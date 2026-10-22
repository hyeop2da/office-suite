// DOCX 최소 편집 코어: 본문, 머리글·바닥글, 각주·미주의 글자(w:t)를 무손실로 고친다.
// 서식(w:rPr, w:pPr), 표, 그림, 스타일은 건드리지 않는다. 글자 요소 안쪽만 바꾼다.
//
// 탭·줄바꿈·문단 나누기 (2026-10-22 추가)
// OOXML(ECMA-376, 공개 규격)에서 w:tab·w:br는 w:t와 **같은 레벨의 형제 요소**로 w:r(런) 안에 둘 수 있다
// (HWPX의 hp:tab이 hp:t "안에" 중첩되는 것과 다르다 — 각 형식의 실제 구조를 반영한 설계다). 따라서 자유
// 텍스트 안에 \t/\n을 끼워 넣는 방식이 아니라, 커서 위치에 w:tab/w:br 요소를 직접 끼워 넣는 전용 메서드로
// 구현한다. 문단 나누기는 w:p를 그대로 복제해 w:pPr(문단 서식)을 양쪽에 남기고 run을 나눈다.
import { ZipArchive } from './zip.ts';
import { DOCX_CONFIG, FormatError, findElementSpan, lastTagOpenBefore, encodeText, type TextNode } from './xml-scan.ts';
import { TextPart, EditError, type EditOptions, type ParagraphView } from './text-part.ts';

function needsPreserve(s: string): boolean { return /^\s|\s$|\s\s/.test(s); }
function buildWT(s: string): string {
  return s === '' ? '' : `<w:t${needsPreserve(s) ? ' xml:space="preserve"' : ''}>${encodeText(s)}</w:t>`;
}
function nodeFullEnd(n: TextNode): number { return n.selfClosing ? n.tagEnd : n.contentEnd + '</w:t>'.length; }

/**
 * span 안의 모든 w:r이 (w:rPr? 하나) 뒤에 w:t 하나만 담고 있는지 확인한다(하이퍼링크·그림·필드·탭·줄바꿈이
 * 섞인 문단은 아직 지원하지 않는다는 뜻). 통과하면 true.
 */
function onlyPlainRuns(xml: string): boolean {
  const re = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>|<w:r\b[^>]*\/>/g;
  let last = 0, m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    if (m.index !== last) return false;
    last = re.lastIndex;
    const body = (m[1] ?? '').replace(/^<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>|^<w:rPr\b[^>]*\/>/, '');
    if (!/^(?:<w:t\b[^>]*>[\s\S]*<\/w:t>|<w:t\b[^>]*\/>)$/.test(body.trim())) return false;
  }
  return last === xml.length;
}

const PART_PATTERNS = [
  /^word\/document\.xml$/,
  /^word\/header\d*\.xml$/,
  /^word\/footer\d*\.xml$/,
  /^word\/footnotes\.xml$/,
  /^word\/endnotes\.xml$/,
];

export class DocxDocument {
  readonly zip: ZipArchive;
  private parts = new Map<string, TextPart>();

  private constructor(zip: ZipArchive) { this.zip = zip; }

  static open(bytes: Uint8Array): DocxDocument {
    const zip = new ZipArchive(bytes);
    if (!zip.has('[Content_Types].xml') || !zip.has('word/document.xml')) throw new FormatError('DOCX가 아닙니다(word/document.xml 없음)');
    if (!/wordprocessingml\.(document|template)/.test(zip.readText('[Content_Types].xml'))) {
      throw new FormatError('DOCX가 아닙니다(콘텐츠 형식이 워드 문서가 아님)');
    }
    return new DocxDocument(zip);
  }

  partNames(): string[] {
    return this.zip.entries.map(e => e.name).filter(n => PART_PATTERNS.some(p => p.test(n)))
      .sort((a, b) => (a === 'word/document.xml' ? -1 : b === 'word/document.xml' ? 1 : a.localeCompare(b)));
  }

  part(name: string): TextPart {
    let p = this.parts.get(name);
    if (!p) {
      if (!this.partNames().includes(name)) throw new EditError(`부품이 없습니다: ${name}`);
      this.parts.set(name, p = new TextPart(this.zip, name, DOCX_CONFIG));
    }
    return p;
  }

  paragraphs(): (ParagraphView & { section: string })[] {
    return this.partNames().flatMap(n => this.part(n).paragraphs().map(p => ({ ...p, section: n })));
  }

  setParagraphText(section: string, paraIndex: number, text: string, opt: EditOptions = {}): void {
    this.part(section).setParagraphText(paraIndex, text, opt);
  }

  plainText(section: string): string { return this.part(section).plainText(); }

  /** 문단 안 charOffset(코드포인트) 위치에 탭(<w:tab/>)을 끼워 넣는다. */
  insertTab(section: string, paraIndex: number, charOffset: number): void {
    this.insertMark(section, paraIndex, charOffset, '<w:tab/>');
  }

  /** 문단 안 charOffset 위치에 줄 안 줄바꿈(<w:br/>, Shift+Enter에 해당)을 끼워 넣는다. */
  insertLineBreak(section: string, paraIndex: number, charOffset: number): void {
    this.insertMark(section, paraIndex, charOffset, '<w:br/>');
  }

  private insertMark(section: string, paraIndex: number, charOffset: number, markXml: string): void {
    const part = this.part(section);
    const scan = part.scan();
    if (!Number.isInteger(paraIndex) || paraIndex < 0 || paraIndex >= scan.paras.length) throw new EditError(`문단 번호 범위 밖: ${paraIndex}`);
    const nodes = scan.texts.filter(t => t.paraIndex === paraIndex && t.context !== 'fallback');
    const lens = nodes.map(n => [...n.text].length);
    const total = lens.reduce((a, b) => a + b, 0);
    if (!Number.isInteger(charOffset) || charOffset < 0 || charOffset > total) throw new EditError(`넣을 위치가 범위 밖입니다: ${charOffset}`);
    let acc = 0, target: TextNode | undefined, local = 0;
    for (let i = 0; i < nodes.length; i++) {
      if (charOffset <= acc + lens[i]) { target = nodes[i]; local = charOffset - acc; break; }
      acc += lens[i];
    }
    if (!target) throw new EditError('넣을 수 있는 글자 조각이 없는 빈 문단입니다');
    if (part.isLocked(target)) throw new EditError('하위 요소가 있는 글자 조각에는 넣을 수 없습니다');
    const chars = [...target.text];
    const left = chars.slice(0, local).join(''), right = chars.slice(local).join('');
    const replacement = `${buildWT(left)}${markXml}${buildWT(right)}`;
    part.applyRawEdits([[target.tagStart, nodeFullEnd(target), replacement]], [paraIndex]);
  }

  /**
   * 본문 문단 하나를 charOffset(코드포인트) 위치에서 둘로 나눈다. w:pPr(문단 서식)은 양쪽에 복제한다.
   * 첫 번째 새 문단은 원래 <w:p> 여는 태그(속성 포함)를 그대로 쓰고, 두 번째는 속성 없는 `<w:p>`를 쓴다
   * (w14:paraId 같은 식별자를 중복시키지 않기 위함). 하이퍼링크·그림·필드·탭·줄바꿈이 섞인 문단, 표 셀이
   * 아닌 문맥이 아니면 거부한다.
   */
  splitParagraph(section: string, paraIndex: number, charOffset: number): void {
    const part = this.part(section);
    const scan = part.scan();
    if (!Number.isInteger(paraIndex) || paraIndex < 0 || paraIndex >= scan.paras.length) throw new EditError(`문단 번호 범위 밖: ${paraIndex}`);
    const nodes = scan.texts.filter(t => t.paraIndex === paraIndex && t.context !== 'fallback');
    if (nodes.length === 0) throw new EditError('빈 문단은 나눌 수 없습니다');
    if (nodes.some(n => part.isLocked(n))) throw new EditError('하이퍼링크·그림 등이 있는 문단은 아직 나눌 수 없습니다');
    // w:p는 DOCX에서 서로 중첩되지 않으므로(표는 문단의 형제이지 자식이 아님) 문맥 제한은 필요 없다 —
    // 본문·표 셀·머리글·바닥글·글상자 모두 같은 w:p 구조이고 onlyPlainRuns가 내용을 검사한다.
    const chars = [...nodes.map(n => n.text).join('')];
    if (!Number.isInteger(charOffset) || charOffset < 0 || charOffset > chars.length) throw new EditError(`나눌 위치가 범위 밖입니다: ${charOffset}`);

    const xml = part.xml();
    const para = scan.paras[paraIndex];
    const span = findElementSpan(xml, 'w:p', para.start);
    if (!span) throw new EditError('문단 끝을 찾지 못했습니다');
    const inner = xml.slice(span.openEnd, span.closeStart);
    const pPrM = /^(?:<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>|<w:pPr\b[^>]*\/>)/.exec(inner);
    const pPrXml = pPrM ? pPrM[0] : '';
    const afterPPr = pPrM ? inner.slice(pPrM[0].length) : inner;
    // onlyPlainRuns가 이미 "각 w:r은 (w:rPr 하나)? 뒤에 w:t 하나"만 검사한다. w:rPr 안쪽(w:b 등 서식
    // 토글)은 일부러 들여다보지 않는 불투명 블록으로 취급한다 — 서식 종류를 제한할 이유가 없기 때문이다.
    if (!onlyPlainRuns(afterPPr)) {
      throw new EditError('이 문단에는 나누기를 지원하지 않는 요소(하이퍼링크·그림·필드·탭·줄바꿈 등)가 있습니다');
    }

    let acc = 0, si = -1, lo = 0;
    for (let i = 0; i < nodes.length; i++) {
      const len = [...nodes[i].text].length;
      if (charOffset <= acc + len) { si = i; lo = charOffset - acc; break; }
      acc += len;
    }
    if (si < 0) { si = nodes.length - 1; lo = [...nodes[si].text].length; }
    const straddle = nodes[si];

    const runStart = lastTagOpenBefore(xml, 'w:r', straddle.tagStart);
    const runOpenEnd = xml.indexOf('>', runStart) + 1;
    const runOpenTag = xml.slice(runStart, runOpenEnd);
    const runCloseEnd = xml.indexOf('</w:r>', nodeFullEnd(straddle)) + '</w:r>'.length;
    const rPrM = /^<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>|^<w:rPr\b[^>]*\/>/.exec(xml.slice(runOpenEnd, straddle.tagStart));
    const rPrXml = rPrM ? rPrM[0] : '';

    const chs = [...straddle.text];
    const leftText = chs.slice(0, lo).join(''), rightText = chs.slice(lo).join('');
    // 어느 한쪽이 빈 글자가 되면 그 run 자체를 만들지 않는다(서식만 있고 글자가 없는 빈 run을 남기지 않기 위해 —
    // 경계에서 정확히 나누는 흔한 경우, 예: '굵게|보통' 두 런 사이에서 나눌 때).
    const leftRun = leftText === '' ? '' : `${runOpenTag}${rPrXml}${buildWT(leftText)}</w:r>`;
    const rightRun = rightText === '' ? '' : `${runOpenTag}${rPrXml}${buildWT(rightText)}</w:r>`;

    const runsBefore = xml.slice(span.openEnd + pPrXml.length, runStart);
    const runsAfter = xml.slice(runCloseEnd, span.closeStart);
    const openTag = xml.slice(span.start, span.openEnd);

    const paraA = `${openTag}${pPrXml}${runsBefore}${leftRun}</w:p>`;
    const paraB = `<w:p>${pPrXml}${rightRun}${runsAfter}</w:p>`;
    part.applyRawEdits([[span.start, span.end, paraA + paraB]], []);
  }

  save(): Uint8Array { return this.zip.write(); }
}
