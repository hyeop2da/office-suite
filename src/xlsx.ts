// XLSX 최소 편집 코어: 시트 목록, 셀 읽기, 셀 값 편집(숫자는 숫자로 유지).
// 원칙: 바꾼 셀 요소(<c>)만 다시 쓰고 나머지는 바이트 그대로 둔다. 공유 문자열 표(sharedStrings.xml)와 스타일은 건드리지 않는다.
import { ZipArchive } from './zip.ts';
import { FormatError, decodeEntities, encodeText } from './xml-scan.ts';
import { EditError } from './text-part.ts';

export interface SheetInfo { index: number; name: string; part: string; state: 'visible' | 'hidden' | 'veryHidden' }
export type CellKind = 'empty' | 'number' | 'string' | 'bool' | 'error' | 'formula' | 'other';
export interface CellView {
  ref: string; row: number; col: number; kind: CellKind;
  /** 저장된 값 그대로의 글자(숫자는 원래 표기) */
  text: string;
  /** 날짜 서식 숫자의 ISO 표시 등 */
  display?: string;
  formula?: string;
  editable: boolean;
  /** 서식이 섞인 공유 문자열(편집하면 서식이 사라짐) */
  rich?: boolean;
  /** 병합된 범위에 속하면 그 범위 참조와, 이 칸이 왼쪽 위 중심 칸인지 */
  merged?: { ref: string; anchor: boolean };
}

interface CellPos { ref: string; row: number; col: number; start: number; end: number; attrs: string; inner: string }
interface RowPos { num: number; start: number; end: number; openEnd: number; closeStart: number; attrs: string; cells: CellPos[] }
interface MergeRange { ref: string; r1: number; c1: number; r2: number; c2: number }
interface SheetIndex {
  xml: string; rows: RowPos[]; sdOpen: number; sdClose: number; sdSelfClosing: boolean; byRef: Map<string, CellPos>;
  merges: MergeRange[];
}

export function colToNum(col: string): number { let n = 0; for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
export function numToCol(n: number): string { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
export function parseRef(ref: string): { row: number; col: number } {
  const m = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(ref);
  if (!m) throw new EditError(`셀 주소가 올바르지 않습니다: ${ref}`);
  const col = colToNum(m[1]), row = Number(m[2]);
  if (col > 16384 || row > 1048576) throw new EditError(`셀 주소가 범위 밖입니다: ${ref}`);
  return { row, col };
}

function attr(attrs: string, name: string): string | undefined {
  const m = new RegExp(`(?:^|\\s)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(attrs);
  return m ? decodeEntities(m[1] ?? m[2]) : undefined;
}

const NUM = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;
const OPEN_TAG = (name: string) => new RegExp(`<${name}\\b((?:[^>"']|"[^"]*"|'[^']*')*?)(/?)>`, 'g');

function siText(content: string): { text: string; rich: boolean } {
  const noPh = content.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
  const rich = /<r[\s>]/.test(noPh);
  let text = '';
  for (const m of noPh.matchAll(/<t\b[^>]*?(?:\/>|>([\s\S]*?)<\/t>)/g)) text += decodeEntities(m[1] ?? '');
  return { text, rich };
}

export class XlsxDocument {
  readonly zip: ZipArchive;
  private sheetList: SheetInfo[];
  private shared: { text: string; rich: boolean }[] = [];
  private dateStyles = new Set<number>();
  private date1904 = false;
  private index = new Map<number, SheetIndex>();
  private hasFormulaCache: boolean | undefined;

  private constructor(zip: ZipArchive) {
    this.zip = zip;
    const wb = zip.readText('xl/workbook.xml');
    this.date1904 = /<workbookPr\b[^>]*date1904\s*=\s*"(1|true)"/.test(wb);
    const rels = new Map<string, string>();
    if (zip.has('xl/_rels/workbook.xml.rels')) {
      for (const m of zip.readText('xl/_rels/workbook.xml.rels').matchAll(OPEN_TAG('Relationship'))) {
        const id = attr(m[1], 'Id'), target = attr(m[1], 'Target'), type = attr(m[1], 'Type') ?? '';
        if (id && target && /\/worksheet$/.test(type)) rels.set(id, target.startsWith('/') ? target.slice(1) : `xl/${target}`);
      }
    }
    this.sheetList = [];
    for (const m of wb.matchAll(OPEN_TAG('sheet'))) {
      const name = attr(m[1], 'name');
      const rid = /(?:^|\s)[\w]+:id\s*=\s*"([^"]*)"/.exec(m[1])?.[1];
      const part = rid ? rels.get(rid) : undefined;
      if (name === undefined) continue;
      if (!part) continue; // 차트 시트 등 워크시트가 아닌 것은 목록에서 뺀다
      const state = (attr(m[1], 'state') ?? 'visible') as SheetInfo['state'];
      this.sheetList.push({ index: this.sheetList.length, name, part, state });
    }
    if (this.sheetList.length === 0) throw new FormatError('XLSX에 워크시트가 없습니다');
    if (zip.has('xl/sharedStrings.xml')) {
      for (const m of zip.readText('xl/sharedStrings.xml').matchAll(/<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g)) this.shared.push(siText(m[1] ?? ''));
    }
    if (zip.has('xl/styles.xml')) this.readStyles(zip.readText('xl/styles.xml'));
  }

  static open(bytes: Uint8Array): XlsxDocument {
    const zip = new ZipArchive(bytes);
    if (!zip.has('[Content_Types].xml') || !zip.has('xl/workbook.xml')) throw new FormatError('XLSX가 아닙니다(xl/workbook.xml 없음)');
    return new XlsxDocument(zip);
  }

  private readStyles(xml: string): void {
    const custom = new Map<number, string>();
    for (const m of xml.matchAll(OPEN_TAG('numFmt'))) {
      const id = Number(attr(m[1], 'numFmtId')), code = attr(m[1], 'formatCode');
      if (Number.isFinite(id) && code !== undefined) custom.set(id, code);
    }
    const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml)?.[1] ?? '';
    let i = 0;
    for (const m of xfs.matchAll(/<xf\b((?:[^>"']|"[^"]*"|'[^']*')*?)(\/>|>)/g)) {
      const id = Number(attr(m[1], 'numFmtId') ?? 0);
      const builtinDate = (id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 45 && id <= 47) || (id >= 50 && id <= 58);
      const code = custom.get(id);
      const customDate = code !== undefined && /[dmyhs]/i.test(code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, ''));
      if (builtinDate || customDate) this.dateStyles.add(i);
      i++;
    }
  }

  sheets(): SheetInfo[] { return this.sheetList.slice(); }

  private sheetOf(i: number): SheetInfo {
    const s = this.sheetList[i];
    if (!s) throw new EditError(`시트 번호가 범위 밖입니다: ${i}`);
    return s;
  }

  private idx(i: number): SheetIndex {
    let x = this.index.get(i);
    if (!x) { x = this.buildIndex(this.zip.readText(this.sheetOf(i).part)); this.index.set(i, x); }
    return x;
  }

  private buildIndex(xml: string): SheetIndex {
    const open = /<sheetData\b((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/.exec(xml);
    if (!open) {
      if (/<\w+:sheetData\b/.test(xml)) throw new FormatError('접두사가 붙은 XLSX 요소(예: x:sheetData)는 아직 지원하지 않습니다');
      throw new FormatError('시트에 sheetData가 없습니다');
    }
    const sdSelfClosing = open[2] === '/';
    const sdOpen = open.index + open[0].length;
    const sdClose = sdSelfClosing ? sdOpen : xml.indexOf('</sheetData>', sdOpen);
    if (sdClose < 0) throw new FormatError('sheetData가 닫히지 않았습니다');
    const rows: RowPos[] = [];
    const byRef = new Map<string, CellPos>();
    if (!sdSelfClosing) {
      const body = xml.slice(sdOpen, sdClose);
      const rowRe = /<row\b((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/g;
      let prevRow = 0;
      for (let m = rowRe.exec(body); m; m = rowRe.exec(body)) {
        const start = sdOpen + m.index, openEnd = start + m[0].length;
        let end = openEnd, closeStart = -1;
        if (m[2] !== '/') { closeStart = xml.indexOf('</row>', openEnd); if (closeStart < 0) throw new FormatError('row가 닫히지 않았습니다'); end = closeStart + 6; rowRe.lastIndex = end - sdOpen; }
        const num = Number(attr(m[1], 'r') ?? prevRow + 1);
        prevRow = num;
        const row: RowPos = { num, start, end, openEnd, closeStart, attrs: m[1], cells: [] };
        if (closeStart >= 0) {
          const inner = xml.slice(openEnd, closeStart);
          const cRe = /<c\b((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/g;
          let prevCol = 0;
          for (let c = cRe.exec(inner); c; c = cRe.exec(inner)) {
            const cs = openEnd + c.index;
            let ce = cs + c[0].length, content = '';
            if (c[2] !== '/') { const cl = xml.indexOf('</c>', ce); if (cl < 0) throw new FormatError('c가 닫히지 않았습니다'); content = xml.slice(ce, cl); ce = cl + 4; cRe.lastIndex = ce - openEnd; }
            const r = attr(c[1], 'r');
            const pos = r ? parseRef(r) : { row: num, col: prevCol + 1 };
            prevCol = pos.col;
            const ref = `${numToCol(pos.col)}${pos.row}`;
            const cell: CellPos = { ref, row: pos.row, col: pos.col, start: cs, end: ce, attrs: c[1], inner: content };
            row.cells.push(cell);
            byRef.set(ref, cell);
          }
        }
        rows.push(row);
      }
    }
    const merges: MergeRange[] = [];
    for (const m of xml.matchAll(OPEN_TAG('mergeCell'))) {
      const ref = attr(m[1], 'ref');
      if (!ref) continue;
      const [a, b = a] = ref.split(':');
      try {
        const p1 = parseRef(a), p2 = parseRef(b);
        merges.push({ ref, r1: Math.min(p1.row, p2.row), c1: Math.min(p1.col, p2.col), r2: Math.max(p1.row, p2.row), c2: Math.max(p1.col, p2.col) });
      } catch { /* 잘못된 ref는 건너뜀 */ }
    }
    return { xml, rows, sdOpen, sdClose, sdSelfClosing, byRef, merges };
  }

  /** 이 시트의 병합된 범위 목록(예: ["A1:B2"]). mergeCells 요소는 건드리지 않고 읽기만 한다. */
  mergedRanges(sheet: number): string[] { return this.idx(sheet).merges.map(m => m.ref); }

  private mergeOf(sx: SheetIndex, row: number, col: number): MergeRange | undefined {
    return sx.merges.find(m => row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2);
  }

  /** 사용된 범위(값·수식이 있는 칸 기준)의 행·열 수 */
  usedRange(sheet: number): { rows: number; cols: number } {
    let rows = 0, cols = 0;
    for (const c of this.idx(sheet).byRef.values()) {
      if (this.viewOf(c).kind === 'empty') continue;
      rows = Math.max(rows, c.row); cols = Math.max(cols, c.col);
    }
    return { rows, cols };
  }

  private viewOf(c: CellPos): CellView {
    const t = attr(c.attrs, 't');
    const s = Number(attr(c.attrs, 's') ?? 0);
    const f = /<f\b((?:[^>"']|"[^"]*"|'[^']*')*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(c.inner);
    const v = /<v\b[^>]*?(?:\/>|>([\s\S]*?)<\/v>)/.exec(c.inner);
    const vtext = v ? decodeEntities(v[1] ?? '') : '';
    const base = { ref: c.ref, row: c.row, col: c.col };
    if (f) {
      return { ...base, kind: 'formula', text: vtext, formula: decodeEntities(f[2] ?? '') || `(공유 수식 ${attr(f[1], 'si') ?? ''})`, editable: false };
    }
    if (t === 's') {
      const e = this.shared[Number(vtext)];
      return { ...base, kind: 'string', text: e ? e.text : '', editable: !!e, rich: e?.rich };
    }
    if (t === 'inlineStr') {
      const is = /<is\b[^>]*?>([\s\S]*?)<\/is>/.exec(c.inner);
      const e = siText(is ? is[1] : '');
      return { ...base, kind: 'string', text: e.text, editable: true, rich: e.rich };
    }
    if (t === 'str') return { ...base, kind: 'string', text: vtext, editable: true };
    if (t === 'b') return { ...base, kind: 'bool', text: vtext === '1' || vtext === 'true' ? 'TRUE' : 'FALSE', editable: true };
    if (t === 'e') return { ...base, kind: 'error', text: vtext, editable: false };
    if (t === 'd') return { ...base, kind: 'other', text: vtext, editable: false };
    if (!v || vtext === '') return { ...base, kind: 'empty', text: '', editable: true };
    const out: CellView = { ...base, kind: 'number', text: vtext, editable: true };
    if (this.dateStyles.has(s) && Number.isFinite(Number(vtext))) out.display = this.serialToIso(Number(vtext));
    return out;
  }

  cell(sheet: number, ref: string): CellView {
    const { row, col } = parseRef(ref);
    const sx = this.idx(sheet);
    const c = sx.byRef.get(ref);
    const view = c ? this.viewOf(c) : ({ ref, row, col, kind: 'empty', text: '', editable: true } as CellView);
    const mg = this.mergeOf(sx, row, col);
    if (mg) view.merged = { ref: mg.ref, anchor: row === mg.r1 && col === mg.c1 };
    return view;
  }

  /** startRow부터 rows개 행, 1열부터 cols개 열의 격자 */
  grid(sheet: number, rows: number, cols: number, startRow = 1): CellView[][] {
    const out: CellView[][] = [];
    for (let r = startRow; r < startRow + rows; r++) {
      const line: CellView[] = [];
      for (let c = 1; c <= cols; c++) line.push(this.cell(sheet, `${numToCol(c)}${r}`));
      out.push(line);
    }
    return out;
  }

  private serialToIso(serial: number): string {
    const base = this.date1904 ? Date.UTC(1904, 0, 1) : (serial < 61 ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30));
    const d = new Date(base + Math.round(serial * 86400000));
    const iso = d.toISOString();
    return Number.isInteger(serial) ? iso.slice(0, 10) : iso.slice(0, 19).replace('T', ' ');
  }

  private isoToSerial(iso: string): number | undefined {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
    if (!m) return undefined;
    const t = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    const d = new Date(t);
    if (d.getUTCFullYear() !== Number(m[1]) || d.getUTCMonth() !== Number(m[2]) - 1 || d.getUTCDate() !== Number(m[3])) return undefined;
    return Math.round((t - (this.date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30))) / 86400000);
  }

  /**
   * 입력 글자를 해석한다. 원래 칸의 종류를 지킨다: 숫자 칸은 숫자로, 글자 칸은 글자로 남는다.
   * 잘못된 입력은 EditError(문서는 바뀌지 않음).
   */
  checkInput(sheet: number, ref: string, text: string): { type: 'number' | 'string' | 'bool' | 'empty'; value: string } {
    if (typeof text !== 'string') throw new EditError('글자는 문자열이어야 합니다');
    const cur = this.cell(sheet, ref);
    if (cur.merged && !cur.merged.anchor) throw new EditError(`병합된 범위(${cur.merged.ref})의 왼쪽 위 칸에서만 편집할 수 있습니다`);
    if (!cur.editable) throw new EditError(cur.kind === 'formula' ? '수식 칸은 편집할 수 없습니다(결과 값은 수식이 정합니다)' : '이 칸은 편집할 수 없습니다');
    const norm = text.replace(/\r\n?/g, '\n');
    for (let i = 0; i < norm.length; i++) {
      const c = norm.charCodeAt(i);
      if ((c < 0x20 && c !== 0x09 && c !== 0x0a) || c === 0xfffe || c === 0xffff) throw new EditError(`셀에 쓸 수 없는 제어 문자입니다(U+${c.toString(16).padStart(4, '0')})`);
      if (c >= 0xd800 && c <= 0xdbff) { const d = norm.charCodeAt(i + 1); if (!(d >= 0xdc00 && d <= 0xdfff)) throw new EditError('짝이 맞지 않는 서로게이트 글자입니다'); i++; }
      else if (c >= 0xdc00 && c <= 0xdfff) throw new EditError('짝이 맞지 않는 서로게이트 글자입니다');
    }
    if (norm.length > 32767) throw new EditError('셀 글자는 32,767자를 넘을 수 없습니다');
    const trimmed = norm.trim();
    const number = (): string => {
      if (!NUM.test(trimmed)) throw new EditError('숫자 칸에는 숫자만 넣을 수 있습니다(예: 12, -3.5, 1e3). 글자로 바꾸려면 먼저 칸을 비우세요');
      const n = Number(trimmed);
      if (!Number.isFinite(n)) throw new EditError('숫자가 너무 큽니다');
      return trimmed.replace(/^\+/, '');
    };
    switch (cur.kind) {
      case 'number': {
        if (trimmed === '') return { type: 'empty', value: '' };
        if (cur.display !== undefined) { const s = this.isoToSerial(trimmed); if (s !== undefined) return { type: 'number', value: String(s) }; }
        return { type: 'number', value: number() };
      }
      case 'bool': {
        const u = trimmed.toUpperCase();
        if (trimmed === '') return { type: 'empty', value: '' };
        if (u === 'TRUE' || u === '1') return { type: 'bool', value: '1' };
        if (u === 'FALSE' || u === '0') return { type: 'bool', value: '0' };
        throw new EditError('논리값 칸에는 TRUE 또는 FALSE만 넣을 수 있습니다');
      }
      case 'string': return norm === '' ? { type: 'empty', value: '' } : { type: 'string', value: norm };
      default: // 빈 칸: 숫자처럼 보이면 숫자, 아니면 글자
        if (norm === '') return { type: 'empty', value: '' };
        if (NUM.test(trimmed) && Number.isFinite(Number(trimmed))) return { type: 'number', value: trimmed.replace(/^\+/, '') };
        return { type: 'string', value: norm };
    }
  }

  setCell(sheet: number, ref: string, text: string): void {
    const input = this.checkInput(sheet, ref, text);
    const sx = this.idx(sheet);
    const { row, col } = parseRef(ref);
    const old = sx.byRef.get(ref);
    // 기존 속성에서 t만 빼고 유지한다(r, s, cm 등)
    const keep = (old ? old.attrs : ` r="${ref}"`).replace(/\s+t\s*=\s*(?:"[^"]*"|'[^']*')/, '');
    let newXml: string;
    switch (input.type) {
      case 'empty': newXml = `<c${keep}/>`; break;
      case 'number': newXml = `<c${keep}><v>${input.value}</v></c>`; break;
      case 'bool': newXml = `<c${keep} t="b"><v>${input.value}</v></c>`; break;
      default: newXml = `<c${keep} t="inlineStr"><is><t xml:space="preserve">${encodeText(input.value)}</t></is></c>`;
    }
    let xml = sx.xml;
    if (old) {
      if (input.type === 'empty' && !old.inner.trim()) return;
      xml = xml.slice(0, old.start) + newXml + xml.slice(old.end);
    } else {
      if (input.type === 'empty') return;
      const rowPos = sx.rows.find(r => r.num === row);
      if (rowPos) {
        let at: number;
        let ins = newXml;
        if (rowPos.closeStart < 0) { // <row .../> 를 <row ...>cell</row>로
          const open = xml.slice(rowPos.start, rowPos.openEnd).replace(/\s*\/>$/, '>');
          xml = xml.slice(0, rowPos.start) + open + ins + '</row>' + xml.slice(rowPos.end);
          xml = this.widen(xml, rowPos.start, col);
        } else {
          const after = rowPos.cells.find(c => c.col > col);
          at = after ? after.start : rowPos.closeStart;
          xml = xml.slice(0, at) + ins + xml.slice(at);
          xml = this.widen(xml, rowPos.start, col);
        }
      } else {
        const after = sx.rows.find(r => r.num > row);
        const rowXml = `<row r="${row}">${newXml}</row>`;
        if (sx.sdSelfClosing) {
          const m = /<sheetData\b[^>]*?\/>/.exec(xml)!;
          xml = xml.slice(0, m.index) + m[0].replace(/\s*\/>$/, '>') + rowXml + '</sheetData>' + xml.slice(m.index + m[0].length);
        } else {
          const at = after ? after.start : sx.sdClose;
          xml = xml.slice(0, at) + rowXml + xml.slice(at);
        }
      }
    }
    xml = this.fixDimension(xml, row, col);
    this.commit(sheet, xml);
  }

  /** 행의 spans="a:b" 힌트가 새 열을 포함하도록 넓힌다. */
  private widen(xml: string, rowStart: number, col: number): string {
    const end = xml.indexOf('>', rowStart);
    const tag = xml.slice(rowStart, end + 1);
    const m = /\sspans\s*=\s*"(\d+):(\d+)"/.exec(tag);
    if (!m) return xml;
    const lo = Math.min(Number(m[1]), col), hi = Math.max(Number(m[2]), col);
    return xml.slice(0, rowStart) + tag.replace(m[0], ` spans="${lo}:${hi}"`) + xml.slice(end + 1);
  }

  private fixDimension(xml: string, row: number, col: number): string {
    const m = /<dimension\b[^>]*?\sref\s*=\s*"([^"]*)"[^>]*?\/>/.exec(xml);
    if (!m) return xml;
    const [a, b = a] = m[1].split(':');
    let p1: { row: number; col: number }, p2: { row: number; col: number };
    try { p1 = parseRef(a); p2 = parseRef(b); } catch { return xml; }
    const r1 = Math.min(p1.row, row), c1 = Math.min(p1.col, col), r2 = Math.max(p2.row, row), c2 = Math.max(p2.col, col);
    if (r1 === p1.row && c1 === p1.col && r2 === p2.row && c2 === p2.col) return xml;
    const ref = `${numToCol(c1)}${r1}:${numToCol(c2)}${r2}`;
    return xml.slice(0, m.index) + m[0].replace(m[1], ref) + xml.slice(m.index + m[0].length);
  }

  private hasFormulas(): boolean {
    if (this.hasFormulaCache === undefined) this.hasFormulaCache = this.sheetList.some(s => /<f[\s>/]/.test(this.zip.readText(s.part)));
    return this.hasFormulaCache;
  }

  private commit(sheet: number, xml: string): void {
    const s = this.sheetOf(sheet);
    this.zip.replace(s.part, new TextEncoder().encode(xml));
    this.index.delete(sheet);
    // 수식이 있는 통합문서는 열 때 다시 계산하게 한다(수식 셀의 저장된 결과 값은 편집 전 값이다)
    if (this.hasFormulas()) this.requestRecalc();
  }

  private requestRecalc(): void {
    const name = 'xl/workbook.xml';
    const wb = this.zip.readText(name);
    let out: string;
    const m = /<calcPr\b((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/.exec(wb);
    if (m) {
      if (/\sfullCalcOnLoad\s*=\s*"(1|true)"/.test(m[1])) return;
      const attrs = /\sfullCalcOnLoad\s*=\s*"[^"]*"/.test(m[1]) ? m[1].replace(/\sfullCalcOnLoad\s*=\s*"[^"]*"/, ' fullCalcOnLoad="1"') : `${m[1]} fullCalcOnLoad="1"`;
      out = wb.slice(0, m.index) + `<calcPr${attrs}${m[2]}>` + wb.slice(m.index + m[0].length);
    } else {
      // 스키마 순서상 calcPr은 definedNames 뒤, oleSize 등 앞이다
      const after = /<(oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b/.exec(wb);
      const at = after ? after.index : wb.lastIndexOf('</workbook>');
      if (at < 0) return;
      out = wb.slice(0, at) + '<calcPr fullCalcOnLoad="1"/>' + wb.slice(at);
    }
    this.zip.replace(name, new TextEncoder().encode(out));
  }

  save(): Uint8Array { return this.zip.write(); }
}
