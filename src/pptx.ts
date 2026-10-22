// PPTX 최소 편집 코어: 슬라이드와 발표자 노트의 글자(a:t)를 무손실로 고친다.
// 필드(슬라이드 번호·날짜 등 a:fld)는 고치지 않는다. 레이아웃·마스터·도형·그림·애니메이션은 건드리지 않는다.
//
// 탭·줄바꿈·문단 나누기 (2026-10-22 추가)
// ECMA-376/ISO 29500(DrawingML, 공개 규격) 확인: a:p가 허용하는 자식은 a:pPr?, (a:r|a:br|a:fld)*,
// a:endParaRPr? 뿐이다. a:br은 a:r(런)의 **형제**로 온다(DOCX의 w:br이 w:r *안의* w:t와 형제인 것과 다르다 —
// 한 a:r은 rPr? 뒤에 a:t 딱 하나만 담을 수 있어 그 안에 끼워 넣을 자리가 없다). 탭은 전용 요소가 없다 — a:t는
// xsd:string이라 탭 문자(U+0009)를 그대로 담을 수 있고, 이것이 유일한 방법이다(디컴파일 없이 스키마로 확인
// 가능한 결론). 줄 안에서 줄을 바꾸려면 그 위치의 런을 둘로 쪼개고 사이에 <a:br/>을 끼운다.
import { ZipArchive } from './zip.ts';
import { PPTX_CONFIG, FormatError, findElementSpan, lastTagOpenBefore, encodeText, decodeEntities, type TextNode } from './xml-scan.ts';
import { TextPart, EditError, type EditOptions, type ParagraphView } from './text-part.ts';

function buildAT(s: string): string { return s === '' ? '' : `<a:t>${encodeText(s)}</a:t>`; }
function nodeFullEnd(n: TextNode): number { return n.selfClosing ? n.tagEnd : n.contentEnd + '</a:t>'.length; }

/** span 안의 모든 a:r이 (a:rPr? 하나) 뒤에 a:t 하나만 담고 있는지(필드·줄바꿈이 섞인 문단은 아직 미지원). */
function onlyPlainRuns(xml: string): boolean {
  const re = /<a:r\b[^>]*>([\s\S]*?)<\/a:r>|<a:r\b[^>]*\/>/g;
  let last = 0, m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    if (m.index !== last) return false;
    last = re.lastIndex;
    const body = (m[1] ?? '').replace(/^<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>|^<a:rPr\b[^>]*\/>/, '');
    if (!/^(?:<a:t\b[^>]*>[\s\S]*<\/a:t>|<a:t\b[^>]*\/>)$/.test(body.trim())) return false;
  }
  return last === xml.length;
}

export interface SlideInfo { number: number; part: string; notesPart?: string }

const OPEN_TAG = (name: string) => new RegExp(`<${name}\\b((?:[^>"']|"[^"]*"|'[^']*')*?)(/?)>`, 'g');
function attr(attrs: string, name: string): string | undefined {
  const m = new RegExp(`(?:^|\\s)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(attrs);
  return m ? decodeEntities(m[1] ?? m[2]) : undefined;
}
function resolve(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/'); parts.pop();
  for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.') parts.push(seg); }
  return parts.join('/');
}

export class PptxDocument {
  readonly zip: ZipArchive;
  private slideList: SlideInfo[] = [];
  private parts = new Map<string, TextPart>();

  private constructor(zip: ZipArchive) {
    this.zip = zip;
    const rels = new Map<string, string>();
    if (zip.has('ppt/_rels/presentation.xml.rels')) {
      for (const m of zip.readText('ppt/_rels/presentation.xml.rels').matchAll(OPEN_TAG('Relationship'))) {
        const id = attr(m[1], 'Id'), target = attr(m[1], 'Target'), type = attr(m[1], 'Type') ?? '';
        if (id && target && /\/slide$/.test(type)) rels.set(id, resolve('ppt/presentation.xml', target));
      }
    }
    for (const m of zip.readText('ppt/presentation.xml').matchAll(OPEN_TAG('p:sldId'))) {
      const rid = /(?:^|\s)[\w]+:id\s*=\s*"([^"]*)"/.exec(m[1])?.[1];
      const part = rid ? rels.get(rid) : undefined;
      if (!part || !zip.has(part)) continue;
      const slide: SlideInfo = { number: this.slideList.length + 1, part };
      const relsPath = part.replace(/([^/]+)$/, '_rels/$1.rels');
      if (zip.has(relsPath)) {
        for (const r of zip.readText(relsPath).matchAll(OPEN_TAG('Relationship'))) {
          if (/\/notesSlide$/.test(attr(r[1], 'Type') ?? '')) {
            const np = resolve(part, attr(r[1], 'Target') ?? '');
            if (zip.has(np)) slide.notesPart = np;
          }
        }
      }
      this.slideList.push(slide);
    }
  }

  static open(bytes: Uint8Array): PptxDocument {
    const zip = new ZipArchive(bytes);
    if (!zip.has('[Content_Types].xml') || !zip.has('ppt/presentation.xml')) throw new FormatError('PPTX가 아닙니다(ppt/presentation.xml 없음)');
    const doc = new PptxDocument(zip);
    if (doc.slideList.length === 0) throw new FormatError('PPTX에 슬라이드가 없습니다');
    return doc;
  }

  slides(): SlideInfo[] { return this.slideList.slice(); }

  /** 편집 가능한 부품 이름(슬라이드 순서, 각 슬라이드 뒤에 노트) */
  partNames(): string[] { return this.slideList.flatMap(s => s.notesPart ? [s.part, s.notesPart] : [s.part]); }

  label(part: string): string {
    for (const s of this.slideList) {
      if (s.part === part) return `슬라이드 ${s.number}`;
      if (s.notesPart === part) return `슬라이드 ${s.number} 노트`;
    }
    return part;
  }

  part(name: string): TextPart {
    let p = this.parts.get(name);
    if (!p) {
      if (!this.partNames().includes(name)) throw new EditError(`부품이 없습니다: ${name}`);
      this.parts.set(name, p = new TextPart(this.zip, name, PPTX_CONFIG));
    }
    return p;
  }

  paragraphs(): (ParagraphView & { section: string; group: string })[] {
    return this.partNames().flatMap(n => this.part(n).paragraphs().map(p => ({ ...p, section: n, group: this.label(n) })));
  }

  setParagraphText(section: string, paraIndex: number, text: string, opt: EditOptions = {}): void {
    this.part(section).setParagraphText(paraIndex, text, opt);
  }

  /**
   * 문단 안 charOffset(코드포인트) 위치에 줄 안 줄바꿈(Shift+Enter에 해당)을 끼워 넣는다. a:br은 a:r의
   * 형제라 그 위치를 담은 런을 둘로 쪼개고 사이에 끼운다(서식은 양쪽에 복제).
   */
  insertLineBreak(section: string, paraIndex: number, charOffset: number): void {
    const part = this.part(section);
    const scan = part.scan();
    if (!Number.isInteger(paraIndex) || paraIndex < 0 || paraIndex >= scan.paras.length) throw new EditError(`문단 번호 범위 밖: ${paraIndex}`);
    // 필드(a:fld)는 nodes에서 뺀 뒤 위치를 계산하므로, 문단 전체에서 먼저 필드 유무를 확인해야 한다
    // (필드가 속한 run은 그대로 두더라도, 같은 문단의 다른 run을 건드리는 것도 아직은 보수적으로 거부).
    if (scan.texts.some(t => t.paraIndex === paraIndex && part.isLocked(t))) {
      throw new EditError('필드 등이 있는 문단에는 아직 줄바꿈을 넣을 수 없습니다');
    }
    const nodes = scan.texts.filter(t => t.paraIndex === paraIndex && t.context !== 'field');
    const lens = nodes.map(n => [...n.text].length);
    const total = lens.reduce((a, b) => a + b, 0);
    if (!Number.isInteger(charOffset) || charOffset < 0 || charOffset > total) throw new EditError(`넣을 위치가 범위 밖입니다: ${charOffset}`);
    let acc = 0, target: TextNode | undefined, local = 0;
    for (let i = 0; i < nodes.length; i++) {
      if (charOffset <= acc + lens[i]) { target = nodes[i]; local = charOffset - acc; break; }
      acc += lens[i];
    }
    if (!target) throw new EditError('넣을 수 있는 글자 조각이 없는 빈 문단입니다');
    if (part.isLocked(target)) throw new EditError('하위 요소(필드 등)가 있는 글자 조각에는 넣을 수 없습니다');
    const xml = part.xml();
    const runStart = lastTagOpenBefore(xml, 'a:r', target.tagStart);
    if (runStart < 0) throw new EditError('글자를 담은 런(a:r)을 찾지 못했습니다');
    const runOpenEnd = xml.indexOf('>', runStart) + 1;
    const runOpenTag = xml.slice(runStart, runOpenEnd);
    const runCloseStart = xml.indexOf('</a:r>', nodeFullEnd(target));
    if (runCloseStart < 0) throw new EditError('런이 닫히지 않았습니다');
    const runCloseEnd = runCloseStart + '</a:r>'.length;
    const rPrM = /^<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>|^<a:rPr\b[^>]*\/>/.exec(xml.slice(runOpenEnd, target.tagStart));
    const rPrXml = rPrM ? rPrM[0] : '';
    const chars = [...target.text];
    const left = chars.slice(0, local).join(''), right = chars.slice(local).join('');
    const leftRun = left === '' ? '' : `${runOpenTag}${rPrXml}${buildAT(left)}</a:r>`;
    const rightRun = right === '' ? '' : `${runOpenTag}${rPrXml}${buildAT(right)}</a:r>`;
    part.applyRawEdits([[runStart, runCloseEnd, `${leftRun}<a:br/>${rightRun}`]], [paraIndex]);
  }

  /**
   * 본문 문단 하나를 charOffset 위치에서 둘로 나눈다(Enter에 해당). a:pPr(문단 서식)은 양쪽에 복제한다.
   * 줄바꿈(a:br)·필드(a:fld)가 섞인 문단은 아직 거부한다. 끝에 a:endParaRPr이 있으면(문단 끝 서식) 뒤쪽
   * 문단에 그대로 남는다(원래도 문단 맨 끝에 있던 것이므로 나뉜 뒤에도 같은 자리).
   */
  splitParagraph(section: string, paraIndex: number, charOffset: number): void {
    const part = this.part(section);
    const scan = part.scan();
    if (!Number.isInteger(paraIndex) || paraIndex < 0 || paraIndex >= scan.paras.length) throw new EditError(`문단 번호 범위 밖: ${paraIndex}`);
    const nodes = scan.texts.filter(t => t.paraIndex === paraIndex && t.context !== 'field');
    if (nodes.length === 0) throw new EditError('빈 문단은 나눌 수 없습니다');
    if (nodes.some(n => part.isLocked(n))) throw new EditError('필드 등이 있는 문단은 아직 나눌 수 없습니다');
    const chars = [...nodes.map(n => n.text).join('')];
    if (!Number.isInteger(charOffset) || charOffset < 0 || charOffset > chars.length) throw new EditError(`나눌 위치가 범위 밖입니다: ${charOffset}`);

    const xml = part.xml();
    const para = scan.paras[paraIndex];
    const span = findElementSpan(xml, 'a:p', para.start);
    if (!span) throw new EditError('문단 끝을 찾지 못했습니다');
    const inner = xml.slice(span.openEnd, span.closeStart);
    const pPrM = /^(?:<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>|<a:pPr\b[^>]*\/>)/.exec(inner);
    const pPrXml = pPrM ? pPrM[0] : '';
    const afterPPr = pPrM ? inner.slice(pPrM[0].length) : inner;
    const endM = /(?:<a:endParaRPr\b[^>]*>[\s\S]*?<\/a:endParaRPr>|<a:endParaRPr\b[^>]*\/>)\s*$/.exec(afterPPr);
    const runsOnly = endM ? afterPPr.slice(0, afterPPr.length - endM[0].length) : afterPPr;
    if (!onlyPlainRuns(runsOnly)) {
      throw new EditError('이 문단에는 나누기를 지원하지 않는 요소(줄바꿈·필드 등)가 있습니다');
    }

    let acc = 0, si = -1, lo = 0;
    for (let i = 0; i < nodes.length; i++) {
      const len = [...nodes[i].text].length;
      if (charOffset <= acc + len) { si = i; lo = charOffset - acc; break; }
      acc += len;
    }
    if (si < 0) { si = nodes.length - 1; lo = [...nodes[si].text].length; }
    const straddle = nodes[si];

    const runStart = lastTagOpenBefore(xml, 'a:r', straddle.tagStart);
    const runOpenEnd = xml.indexOf('>', runStart) + 1;
    const runOpenTag = xml.slice(runStart, runOpenEnd);
    const runCloseEnd = xml.indexOf('</a:r>', nodeFullEnd(straddle)) + '</a:r>'.length;
    const rPrM = /^<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>|^<a:rPr\b[^>]*\/>/.exec(xml.slice(runOpenEnd, straddle.tagStart));
    const rPrXml = rPrM ? rPrM[0] : '';

    const chs = [...straddle.text];
    const leftText = chs.slice(0, lo).join(''), rightText = chs.slice(lo).join('');
    const leftRun = leftText === '' ? '' : `${runOpenTag}${rPrXml}${buildAT(leftText)}</a:r>`;
    const rightRun = rightText === '' ? '' : `${runOpenTag}${rPrXml}${buildAT(rightText)}</a:r>`;

    const runsBefore = xml.slice(span.openEnd + pPrXml.length, runStart);
    const runsAfter = xml.slice(runCloseEnd, span.closeStart);
    const openTag = xml.slice(span.start, span.openEnd);

    const paraA = `${openTag}${pPrXml}${runsBefore}${leftRun}</a:p>`;
    const paraB = `<a:p>${pPrXml}${rightRun}${runsAfter}</a:p>`;
    part.applyRawEdits([[span.start, span.end, paraA + paraB]], []);
  }

  save(): Uint8Array { return this.zip.write(); }
}
