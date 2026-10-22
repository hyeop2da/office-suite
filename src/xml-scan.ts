// 본문 XML(HWPX, DOCX)을 DOM으로 다시 쓰지 않고 위치(offset)만 찾는 스캐너.
// 편집은 문자열 일부만 바꾸므로 나머지 바이트(속성 순서, 선언, 알 수 없는 속성 포함)는 그대로 남는다.

export class FormatError extends Error {
  constructor(message: string) { super(message); this.name = 'FormatError'; }
}

/** 형식별 태그 이름 설정. */
export interface PartConfig {
  /** 글자 요소 (HWPX hp:t, DOCX w:t) */
  text: string;
  /** 문단 요소 (hp:p, w:p) */
  para: string;
  /** 줄 배치 캐시 요소(HWPX hp:linesegarray). 없으면 생략 */
  cache?: string;
  /** 조상 요소 이름 → 문맥 이름. 가장 안쪽 조상이 우선하고 'fallback'은 항상 우선한다. */
  contexts: Record<string, string>;
  /** w:t처럼 앞뒤 공백 보존 속성이 필요한 형식 */
  preserveSpaceAttr?: boolean;
  /** 이 문맥의 글자 조각은 고칠 수 없다(필드처럼 프로그램이 값을 채우는 글자) */
  lockedContexts?: string[];
  /**
   * 글자(text)의 형제로 오는 자기닫힘 표시 요소(DOCX w:tab/w:br 등) → 이 요소가 나타내는 고정 글자.
   * 형제 레벨 구조(HWPX처럼 글자 "안에" 중첩되는 구조와 다름)에만 쓴다 — 켜면 이 요소가 보통 글자 조각과
   * 같은 TextNode로 texts에 들어가되 hasChildren=true(= isLocked)로 표시되어, paragraphs()의 .text에는
   * 해당 글자가 보이지만 글자 편집(setParagraphText 등)으로 직접 고치거나 지나쳐 지울 수는 없다.
   */
  markers?: Record<string, string>;
  /**
   * 탭 문자(U+0009)를 글자 요소 내용에 그대로 쓸 수 있는 형식(PPTX). OOXML DrawingML(a:t)에는 HWPX
   * hp:tab·DOCX w:tab 같은 전용 탭 요소가 없다 — a:t는 xsd:string이라 탭 문자를 그대로 담을 수 있고,
   * Paragraph(a:p)가 허용하는 자식도 a:r/a:br/a:fld/a:pPr/a:endParaRPr뿐이라 탭 전용 요소를 넣을 자리가
   * 없다(공개 규격 ECMA-376/ISO 29500 확인, 디컴파일 아님). 꺼져 있으면(기본) validateText가 탭을 막는다.
   */
  allowLiteralTab?: boolean;
}

export const PPTX_CONFIG: PartConfig = {
  text: 'a:t', para: 'a:p',
  contexts: { 'a:tc': 'cell', 'a:fld': 'field' },
  lockedContexts: ['field'],
  markers: { 'a:br': '\n' },
  allowLiteralTab: true,
};

export const HWPX_CONFIG: PartConfig = {
  text: 'hp:t', para: 'hp:p', cache: 'hp:linesegarray',
  contexts: {
    'hp:tc': 'cell', 'hp:header': 'header', 'hp:footer': 'footer',
    'hp:footNote': 'footnote', 'hp:endNote': 'endnote', 'hp:drawText': 'textbox',
  },
};

export const DOCX_CONFIG: PartConfig = {
  text: 'w:t', para: 'w:p',
  contexts: {
    'w:tc': 'cell', 'w:txbxContent': 'textbox', 'mc:Fallback': 'fallback',
    'w:footnote': 'footnote', 'w:endnote': 'endnote', 'w:hdr': 'header', 'w:ftr': 'footer',
  },
  preserveSpaceAttr: true,
  markers: { 'w:tab': '\t', 'w:br': '\n' },
};

export interface TextNode {
  index: number;
  /** 글자 시작 태그의 [시작, 끝) */
  tagStart: number;
  tagEnd: number;
  /** 내용 [시작, 끝). 자기닫힘이면 tagEnd와 같고 selfClosing이 true */
  contentStart: number;
  contentEnd: number;
  selfClosing: boolean;
  /** 내용 안에 하위 요소(hp:tab 등)가 있으면 true */
  hasChildren: boolean;
  /** 엔티티를 푼 글자(하위 요소는 건너뜀) */
  text: string;
  paraIndex: number;
  context: string;
}

export interface ParaInfo {
  index: number;
  start: number;
  /** 줄 배치 캐시 범위(없으면 undefined) */
  lineSegs?: [number, number];
}

export interface ScanResult {
  texts: TextNode[];
  paras: ParaInfo[];
}

const TAG = /<(\/?)([A-Za-z_][\w:.-]*)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/y;

export function scanXml(xml: string, cfg: PartConfig): ScanResult {
  const texts: TextNode[] = [];
  const paras: ParaInfo[] = [];
  const stack: { name: string; textIdx: number; start: number }[] = [];
  const paraStack: number[] = [];
  let i = 0;
  const contextNow = (): string => {
    let ctx = 'body';
    let found = false;
    for (let k = stack.length - 1; k >= 0; k--) {
      const c = cfg.contexts[stack[k].name];
      if (c === 'fallback') return 'fallback';
      if (c && !found) { ctx = c; found = true; }
    }
    return ctx;
  };
  while (i < xml.length) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) break;
    if (xml.startsWith('<?', lt)) { i = skipTo(xml, '?>', lt); continue; }
    if (xml.startsWith('<!--', lt)) { i = skipTo(xml, '-->', lt); continue; }
    if (xml.startsWith('<![CDATA[', lt)) { i = skipTo(xml, ']]>', lt); continue; }
    if (xml.startsWith('<!', lt)) { i = skipTo(xml, '>', lt); continue; }
    TAG.lastIndex = lt;
    const m = TAG.exec(xml);
    if (!m) throw new FormatError(`XML 태그를 해석하지 못했습니다(위치 ${lt})`);
    const [whole, closing, name, , selfClose] = m;
    const end = lt + whole.length;
    if (closing) {
      const top = stack.pop();
      if (!top || top.name !== name) throw new FormatError(`닫는 태그가 맞지 않습니다: </${name}> (위치 ${lt})`);
      if (name === cfg.text && top.textIdx >= 0) {
        const t = texts[top.textIdx];
        t.contentEnd = lt;
        t.text = decodeText(xml, t.contentStart, lt);
      }
      if (name === cfg.para) paraStack.pop();
      if (cfg.cache && name === cfg.cache) {
        const owner = paraStack[paraStack.length - 1];
        if (owner !== undefined) paras[owner].lineSegs = [top.start, end];
      }
    } else {
      const parent = stack[stack.length - 1];
      if (parent && parent.name === cfg.text && parent.textIdx >= 0) texts[parent.textIdx].hasChildren = true;
      if (name === cfg.para) {
        paras.push({ index: paras.length, start: lt });
        if (!selfClose) paraStack.push(paras.length - 1);
      }
      if (name === cfg.text) {
        const node: TextNode = {
          index: texts.length, tagStart: lt, tagEnd: end, contentStart: end, contentEnd: end,
          selfClosing: !!selfClose, hasChildren: false, text: '',
          paraIndex: paraStack[paraStack.length - 1] ?? -1,
          context: contextNow(),
        };
        texts.push(node);
        if (!selfClose) stack.push({ name, textIdx: node.index, start: lt });
      } else if (cfg.markers && selfClose && name in cfg.markers) {
        texts.push({
          index: texts.length, tagStart: lt, tagEnd: end, contentStart: end, contentEnd: end,
          selfClosing: true, hasChildren: true, text: cfg.markers[name],
          paraIndex: paraStack[paraStack.length - 1] ?? -1, context: contextNow(),
        });
      } else if (!selfClose) {
        stack.push({ name, textIdx: -1, start: lt });
      } else if (cfg.cache && name === cfg.cache) {
        const owner = paraStack[paraStack.length - 1];
        if (owner !== undefined) paras[owner].lineSegs = [lt, end];
      }
    }
    i = end;
  }
  if (stack.length) throw new FormatError(`닫히지 않은 태그가 있습니다: ${stack[stack.length - 1].name}`);
  return { texts, paras };
}

/** HWPX 본문 스캔(이전 이름 유지). */
export function scanSection(xml: string): ScanResult { return scanXml(xml, HWPX_CONFIG); }

function skipTo(xml: string, end: string, from: number): number {
  const k = xml.indexOf(end, from);
  if (k < 0) throw new FormatError(`'${end}'로 끝나는 구문이 닫히지 않았습니다(위치 ${from})`);
  return k + end.length;
}

/** 글자 요소 내용에서 글자만 모은다(하위 요소는 건너뜀). */
function decodeText(xml: string, from: number, to: number): string {
  return decodeEntities(xml.slice(from, to).replace(/<[^>]*>/g, ''));
}

export function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_, e: string) => {
    if (e === 'amp') return '&';
    if (e === 'lt') return '<';
    if (e === 'gt') return '>';
    if (e === 'quot') return '"';
    if (e === 'apos') return "'";
    const cp = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
    return cp >= 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : '';
  });
}

export function encodeText(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export interface ElementSpan { start: number; openEnd: number; closeStart: number; end: number }

/**
 * openStart는 `<tag ...>` 여는 태그의 '<' 위치. 같은 이름의 중첩(예: 표 안의 중첩 문단)을 깊이로 세어
 * 짝이 맞는 닫는 태그를 찾는다. 문단 나누기처럼 "문단 전체 범위"가 필요한 연산에 쓴다.
 */
export function findElementSpan(xml: string, tag: string, openStart: number): ElementSpan | undefined {
  const re = new RegExp(`<${tag}\\b((?:[^>"']|"[^"]*"|'[^']*')*?)(/?)>|</${tag}>`, 'g');
  re.lastIndex = openStart;
  const openM = re.exec(xml);
  if (!openM || openM.index !== openStart || openM[0].startsWith('</')) return undefined;
  const openEnd = openStart + openM[0].length;
  if (openM[2] === '/') return { start: openStart, openEnd, closeStart: openEnd, end: openEnd };
  let depth = 1;
  for (let m = re.exec(xml); m; m = re.exec(xml)) {
    if (m[0].startsWith('</')) { depth--; if (depth === 0) return { start: openStart, openEnd, closeStart: m.index, end: m.index + m[0].length }; }
    else if (m[2] !== '/') depth++;
  }
  return undefined;
}

/**
 * before 이전에서 `<tag` 여는 태그의 마지막 시작 위치를 찾는다. 단순 `lastIndexOf('<w:r', ...)`는
 * `<w:rPr>`처럼 긴 이름이 짧은 이름을 접두사로 가지면 잘못된 위치를 찾는다 — 태그 이름 뒤에 공백·'/'·'>'가
 * 오는지 경계를 확인해 그 문제를 피한다.
 */
export function lastTagOpenBefore(xml: string, tag: string, before: number): number {
  const re = new RegExp(`<${tag}(?=[\\s/>])`, 'g');
  let last = -1;
  for (let m = re.exec(xml); m && m.index < before; m = re.exec(xml)) last = m.index;
  return last;
}

/** span 안에 나오는 모든 태그 이름(여는/닫는/자기닫힘 모두)이 allowed 안에 있는지. 없으면 false. */
export function onlyTags(xml: string, allowed: string[]): boolean {
  const re = /<\/?([A-Za-z_][\w:.-]*)/g;
  for (let m = re.exec(xml); m; m = re.exec(xml)) if (!allowed.includes(m[1])) return false;
  return true;
}
