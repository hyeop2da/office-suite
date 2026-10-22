// HWPX 무손실 편집 코어.
// 원칙: 바꾼 글자가 들어 있는 XML 조각만 고치고, 나머지는 바이트 그대로 둔다.
// 근거: HWPX구조분석.md 3.4 (문단 id 재번호, 메타태그 손실, 빈 hp:t 변동 등은 전체를 다시 쓸 때 생긴다).
//
// 탭·줄바꿈·문단 나누기 (2026-10-22 추가)
// 근거: fixtures/hwpx-generated/assembly-briefing.hwpx — rhwp(MIT)로 실제 국회 보도자료 HWP를 변환해
// 얻은 실물 샘플에서 직접 관찰함(디컴파일 아님. 우리가 소지한 도구로 우리 시료를 변환한 것).
// 관찰한 사실: <hp:tab width="4000" leader="0" type="1"/> 형태로 **hp:t 안에 중첩**되어 나타난다(문자와
// 같은 레벨). <hp:lineBreak/>는 속성 없이 자기닫힘. width는 탭 위치마다 달라 보여(1496~51024) 실제 배치
// 계산값으로 추정되지만, 그 계산 방법은 코드 분석 없이는 알 수 없다 — 그래서 새로 넣는 탭은 샘플에서 가장
// 흔했던 width="4000"을 기본값으로 쓴다. 한/글이 이 폭을 그대로 쓰는지 다시 계산하는지는 미검증.
import { ZipArchive } from './zip.ts';
import { HWPX_CONFIG, FormatError, findElementSpan, onlyTags, lastTagOpenBefore, encodeText, decodeEntities, type ScanResult, type TextNode } from './xml-scan.ts';
import { TextPart, EditError, type EditOptions, type ParagraphView } from './text-part.ts';

export type { EditOptions, ParagraphView };
export { EditError, FormatError };

const TAB_CLOSE = /<hp:tab\b[^>]*\/>/g;
const BREAK_CLOSE = /<hp:lineBreak\s*\/>/g;
/** 실측(assembly-briefing.hwpx)에서 가장 많이 나온 값. 실제 배치 폭 계산은 알 수 없어 고정값을 쓴다. */
const DEFAULT_TAB_XML = '<hp:tab width="4000" leader="0" type="1"/>';
const BREAK_XML = '<hp:lineBreak/>';

/** raw(hp:t 내용) 안에 hp:tab·hp:lineBreak가 아닌 다른 태그가 있으면 true(그러면 글자로 다루지 않는다). */
function hasOtherTags(raw: string): boolean {
  return /</.test(raw.replace(TAB_CLOSE, '').replace(BREAK_CLOSE, ''));
}

/** hp:t 내용을 글자로 푼다. hp:tab→\t, hp:lineBreak→\n. hasOtherTags가 false일 때만 믿을 수 있다. */
function richDecode(raw: string): string {
  let out = '';
  let i = 0;
  while (i < raw.length) {
    const lt = raw.indexOf('<', i);
    if (lt < 0) { out += decodeEntities(raw.slice(i)); break; }
    out += decodeEntities(raw.slice(i, lt));
    const tab = /^<hp:tab\b[^>]*\/>/.exec(raw.slice(lt));
    if (tab) { out += '\t'; i = lt + tab[0].length; continue; }
    const br = /^<hp:lineBreak\s*\/>/.exec(raw.slice(lt));
    if (br) { out += '\n'; i = lt + br[0].length; continue; }
    const gt = raw.indexOf('>', lt); // 여기 올 일은 없어야 하지만(hasOtherTags로 먼저 거름) 방어적으로 건너뜀
    i = gt < 0 ? raw.length : gt + 1;
  }
  return out;
}

/** 글자를 hp:t 내용으로 쓴다. \t→hp:tab(기본 폭), \n→hp:lineBreak, 나머지는 보통 글자 이스케이프. */
function richEncode(s: string): string {
  let out = '';
  for (const ch of s) out += ch === '\t' ? DEFAULT_TAB_XML : ch === '\n' ? BREAK_XML : encodeText(ch);
  return out;
}

/** 탭·줄바꿈만 허용하는 글자 검증(문단 나누기는 별도 메서드가 처리하므로 \n은 "줄 안 줄바꿈"으로만 쓰인다). */
function validateRichText(text: string): void {
  if (typeof text !== 'string') throw new EditError('글자는 문자열이어야 합니다');
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c === 0x09 || c === 0x0a) continue;
    if (c === 0x0d) throw new EditError('캐리지 리턴(U+000D)은 쓸 수 없습니다. 줄바꿈은 \\n으로 쓰세요');
    if (c < 0x20 || c === 0xfffe || c === 0xffff) throw new EditError(`XML에 쓸 수 없는 제어 문자입니다(U+${c.toString(16).padStart(4, '0')})`);
    if (c >= 0xd800 && c <= 0xdbff) {
      const d = text.charCodeAt(i + 1);
      if (!(d >= 0xdc00 && d <= 0xdfff)) throw new EditError('짝이 맞지 않는 서로게이트 글자입니다');
      i++;
    } else if (c >= 0xdc00 && c <= 0xdfff) throw new EditError('짝이 맞지 않는 서로게이트 글자입니다');
  }
}

export class HwpxDocument {
  readonly zip: ZipArchive;
  private parts = new Map<string, TextPart>();

  private constructor(zip: ZipArchive) { this.zip = zip; }

  static open(bytes: Uint8Array): HwpxDocument {
    const zip = new ZipArchive(bytes);
    if (!zip.has('mimetype') || zip.readText('mimetype').trim() !== 'application/hwp+zip') {
      throw new FormatError('HWPX가 아닙니다(mimetype 불일치)');
    }
    const doc = new HwpxDocument(zip);
    if (doc.sectionNames().length === 0) throw new FormatError('HWPX에 본문 구역(Contents/section*.xml)이 없습니다');
    return doc;
  }

  sectionNames(): string[] {
    return this.zip.entries.map(e => e.name)
      .filter(n => /^Contents\/section\d+\.xml$/.test(n))
      .sort((a, b) => num(a) - num(b));
  }

  part(section: string): TextPart {
    let p = this.parts.get(section);
    if (!p) {
      if (!this.sectionNames().includes(section)) throw new EditError(`구역이 없습니다: ${section}`);
      this.parts.set(section, p = new TextPart(this.zip, section, HWPX_CONFIG));
    }
    return p;
  }

  scan(section: string): ScanResult { return this.part(section).scan(); }

  /** 구역의 모든 hp:t 글자를 이어 붙인다. */
  plainText(section: string): string { return this.part(section).plainText(); }

  /** 코드포인트 기준 글자 수. */
  charCount(section: string): number { return [...this.plainText(section)].length; }

  /** hp:t 하나의 글자를 바꾼다. 하위 요소(탭 등)가 있는 노드는 거부한다. */
  setText(section: string, nodeIndex: number, text: string, opt: EditOptions = {}): void {
    this.part(section).setText(nodeIndex, text, opt);
  }

  /** 글자가 있는 모든 문단(표 셀, 머리말·꼬리말, 각주 포함). */
  paragraphs(): (ParagraphView & { section: string })[] {
    return this.sectionNames().flatMap(s => this.part(s).paragraphs().map(p => ({ ...p, section: s })));
  }

  setParagraphText(section: string, paraIndex: number, text: string, opt: EditOptions = {}): void {
    this.part(section).setParagraphText(paraIndex, text, opt);
  }

  /** 한 hp:t 안에서 find를 replace로 바꾼다. 바꾼 횟수를 돌려준다. 여러 조각에 걸친 글자는 찾지 못한다. */
  replaceText(find: string, replace: string, opt: EditOptions & { all?: boolean } = {}): number {
    if (!find) throw new EditError('찾을 글자가 비어 있습니다');
    let count = 0;
    for (const section of this.sectionNames()) {
      const part = this.part(section);
      const edits: { node: number; text: string }[] = [];
      for (const n of part.scan().texts) {
        if (n.hasChildren || n.context === 'fallback') continue;
        const hits = n.text.split(find).length - 1;
        if (!hits) continue;
        if (opt.all) {
          edits.push({ node: n.index, text: n.text.split(find).join(replace) });
          count += hits;
        } else {
          edits.push({ node: n.index, text: n.text.replace(find, () => replace) });
          count = 1;
          break;
        }
      }
      if (edits.length) part.applyNodeEdits(edits, opt);
      if (!opt.all && count) return count;
    }
    return count;
  }

  /** 본문 첫 글자 노드 맨 앞에 글자를 넣는다(편집 시험용 마커 삽입 등). */
  prependText(text: string, opt: EditOptions = {}): boolean {
    for (const section of this.sectionNames()) {
      const nodes = this.part(section).scan().texts;
      const k = nodes.findIndex(n => !n.hasChildren && n.text.length > 0 && n.context !== 'fallback');
      if (k >= 0) { this.part(section).setText(k, text + nodes[k].text, opt); return true; }
    }
    return false;
  }

  /** content.hpf의 ModifiedDate를 갱신한다. */
  touchModified(date: Date = new Date()): boolean {
    const name = 'Contents/content.hpf';
    if (!this.zip.has(name)) return false;
    const iso = date.toISOString().replace(/\.\d{3}Z$/, 'Z');
    const src = this.zip.readText(name);
    const out = src.replace(/(<opf:meta name="ModifiedDate"[^>]*>)[^<]*(<\/opf:meta>)/, `$1${iso}$2`);
    if (out === src) return false;
    this.zip.replace(name, new TextEncoder().encode(out));
    return true;
  }

  /**
   * Preview/PrvText.txt에서 바뀐 글자를 찾아 같이 바꾼다(모든 글자를 새로 만들지는 않는다).
   * 한 번만 나오는 경우에만 바꾸고 결과를 돌려준다.
   */
  syncPreviewText(before: string, after: string): boolean {
    const name = 'Preview/PrvText.txt';
    if (!this.zip.has(name) || !before) return false;
    const src = this.zip.readText(name);
    const i = src.indexOf(before);
    if (i < 0 || src.indexOf(before, i + 1) >= 0) return false;
    this.zip.replace(name, new TextEncoder().encode(src.slice(0, i) + after + src.slice(i + before.length)));
    return true;
  }

  /**
   * 문단 글자를 탭(\t)·줄바꿈(\n, Shift+Enter에 해당하는 "줄 안 줄바꿈")을 포함해 읽는다.
   * 문단에 탭·줄바꿈 외의 다른 요소(필드, 중첩 개체 등)가 있으면 그 부분은 빈 글자로 보인다(쓰기는 거부됨).
   */
  richParagraphText(section: string, paraIndex: number): string {
    const part = this.part(section);
    const scan = part.scan();
    if (!Number.isInteger(paraIndex) || paraIndex < 0 || paraIndex >= scan.paras.length) throw new EditError(`문단 번호 범위 밖: ${paraIndex}`);
    const xml = part.xml();
    return scan.texts.filter(t => t.paraIndex === paraIndex && t.context !== 'fallback')
      .map(n => n.selfClosing ? '' : richDecode(xml.slice(n.contentStart, n.contentEnd))).join('');
  }

  /**
   * 문단 전체 글자를 탭·줄바꿈을 포함해 새 값으로 바꾼다(문단을 나누지는 않는다 — 그건 splitParagraph).
   * 문단 안의 모든 hp:t가 글자·탭·줄바꿈만 담고 있어야 한다. 다른 요소가 섞여 있으면 거부한다.
   */
  setParagraphRichText(section: string, paraIndex: number, newText: string, opt: EditOptions = {}): void {
    validateRichText(newText);
    const part = this.part(section);
    const scan = part.scan();
    if (!Number.isInteger(paraIndex) || paraIndex < 0 || paraIndex >= scan.paras.length) throw new EditError(`문단 번호 범위 밖: ${paraIndex}`);
    const nodes = scan.texts.filter(t => t.paraIndex === paraIndex && t.context !== 'fallback');
    const xml = part.xml();
    const raws = nodes.map(n => (n.selfClosing ? '' : xml.slice(n.contentStart, n.contentEnd)));
    if (raws.some(r => hasOtherTags(r))) throw new EditError('이 문단에는 탭·줄바꿈 외의 다른 요소가 있어 리치 텍스트 편집을 지원하지 않습니다');
    const richTexts = raws.map(richDecode);
    const old = richTexts.join('');
    if (old === newText) return;
    const cp = (s: string) => Array.from(s);
    const a = cp(old), b = cp(newText);
    let pre = 0; while (pre < a.length && pre < b.length && a[pre] === b[pre]) pre++;
    let suf = 0; while (suf < a.length - pre && suf < b.length - pre && a[a.length - 1 - suf] === b[b.length - 1 - suf]) suf++;
    const delStart = pre, delEnd = a.length - suf;
    const insert = b.slice(pre, b.length - suf).join('');
    let pos = 0;
    const spans = nodes.map((n, i) => { const chars = cp(richTexts[i]); const s = pos, e = pos + chars.length; pos = e; return { n, chars, s, e }; });
    const touched = spans.filter(sp => sp.s < delEnd && sp.e > delStart);
    const target = touched[0] ?? spans.find(sp => delStart > sp.s && delStart <= sp.e) ?? spans.find(sp => sp.s === delStart);
    if (!target) throw new EditError('글자를 넣을 수 있는 조각이 없습니다');
    const ops: [number, number, string][] = [];
    for (const sp of spans) {
      const del = sp.s < delEnd && sp.e > delStart;
      if (!del && sp !== target) continue;
      const chars = sp.chars.slice();
      const local = del ? Math.max(delStart, sp.s) - sp.s : delStart - sp.s;
      if (del) chars.splice(local, Math.min(delEnd, sp.e) - sp.s - local);
      if (sp === target) chars.splice(local, 0, ...cp(insert));
      const newRich = chars.join('');
      if (newRich === sp.chars.join('')) continue;
      const n = sp.n;
      if (n.selfClosing) {
        if (newRich === '') continue;
        const open = xml.slice(n.tagStart, n.tagEnd).replace(/\s*\/>$/, '>');
        ops.push([n.tagStart, n.tagEnd, `${open}${richEncode(newRich)}</hp:t>`]);
      } else {
        ops.push([n.contentStart, n.contentEnd, richEncode(newRich)]);
      }
    }
    part.applyRawEdits(ops, [paraIndex], opt);
  }

  /**
   * 본문 문단 하나를 charOffset(코드포인트 기준) 위치에서 둘로 나눈다. 두 새 문단은 원래 문단과 같은
   * id를 쓴다(문서 안에 id 중복이 이미 흔해 한/글이 여는 데 지장이 없다고 보고됨 — HWPX구조분석.md 3.4-1).
   * 줄 배치 캐시(hp:linesegarray)는 두 새 문단 모두에서 지운다(레이아웃이 바뀌므로). 표·그림·필드 등
   * run/hp:t/linesegarray/lineseg 이외의 요소가 있거나 본문(body) 문맥이 아니면 거부한다.
   */
  splitParagraph(section: string, paraIndex: number, charOffset: number): void {
    const part = this.part(section);
    const scan = part.scan();
    if (!Number.isInteger(paraIndex) || paraIndex < 0 || paraIndex >= scan.paras.length) throw new EditError(`문단 번호 범위 밖: ${paraIndex}`);
    const nodes = scan.texts.filter(t => t.paraIndex === paraIndex && t.context !== 'fallback');
    if (nodes.some(n => part.isLocked(n))) throw new EditError('탭·필드 등이 있는 문단은 아직 나눌 수 없습니다');
    const context = nodes[0]?.context ?? 'body';
    if (context !== 'body') throw new EditError(`'${context}' 문맥의 문단은 아직 나눌 수 없습니다(본문만 지원)`);
    const chars = [...nodes.map(n => n.text).join('')];
    if (!Number.isInteger(charOffset) || charOffset < 0 || charOffset > chars.length) throw new EditError(`나눌 위치가 범위 밖입니다: ${charOffset}`);

    const xml = part.xml();
    const para = scan.paras[paraIndex];
    const span = findElementSpan(xml, 'hp:p', para.start);
    if (!span) throw new EditError('문단 끝을 찾지 못했습니다');
    const inner = xml.slice(span.openEnd, span.closeStart);
    if (!onlyTags(inner, ['hp:run', 'hp:t', 'hp:linesegarray', 'hp:lineseg'])) {
      throw new EditError('이 문단에는 나누기를 지원하지 않는 요소(표·그림·컨트롤 등)가 있습니다');
    }
    if (nodes.length === 0) throw new EditError('빈 문단은 나눌 수 없습니다');

    let acc = 0, si = -1, lo = 0;
    for (let i = 0; i < nodes.length; i++) {
      const len = [...nodes[i].text].length;
      if (charOffset <= acc + len) { si = i; lo = charOffset - acc; break; }
      acc += len;
    }
    if (si < 0) { si = nodes.length - 1; lo = [...nodes[si].text].length; }
    const straddle = nodes[si];

    const closeTagLen = '</hp:t>'.length;
    const nodeFullEnd = (n: TextNode) => (n.selfClosing ? n.tagEnd : n.contentEnd + closeTagLen);
    const runStart = lastTagOpenBefore(xml, 'hp:run', straddle.tagStart);
    const runOpenEnd = xml.indexOf('>', runStart) + 1;
    const runOpenTag = xml.slice(runStart, runOpenEnd);
    const runCloseEnd = xml.indexOf('</hp:run>', nodeFullEnd(straddle)) + '</hp:run>'.length;

    const chs = [...straddle.text];
    const leftText = chs.slice(0, lo).join(''), rightText = chs.slice(lo).join('');
    // onlyTags가 이미 run/t/linesegarray/lineseg만 허용했으므로(탭·줄바꿈 없음) 보통 글자 이스케이프만 하면 된다.
    const buildT = (t: string) => (t === '' ? '<hp:t/>' : `<hp:t>${encodeText(t)}</hp:t>`);
    const leftRun = `${runOpenTag}${buildT(leftText)}</hp:run>`;
    const rightRun = `${runOpenTag}${buildT(rightText)}</hp:run>`;

    const runsBefore = xml.slice(span.openEnd, runStart);
    const runsAfter = xml.slice(runCloseEnd, span.closeStart).replace(/<hp:linesegarray>[\s\S]*?<\/hp:linesegarray>/, '');
    const openTag = xml.slice(span.start, span.openEnd);

    const paraA = `${openTag}${runsBefore}${leftRun}</hp:p>`;
    const paraB = `${openTag}${rightRun}${runsAfter}</hp:p>`;
    part.applyRawEdits([[span.start, span.end, paraA + paraB]], []);
  }

  save(): Uint8Array { return this.zip.write(); }
}

function num(name: string): number { return Number(/section(\d+)\.xml$/.exec(name)![1]); }
