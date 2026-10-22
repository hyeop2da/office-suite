// XLSX 편집 세션: 원본은 바꾸지 않고 칸별 수정 목록만 쌓았다가 저장할 때 한 번에 적용한다(session.ts와 같은 구조).
import { XlsxDocument, numToCol, parseRef } from './xlsx.ts';
import type { SaveResult } from './session.ts';

const VERIFY_LIMIT = 200_000;

export function cellKey(sheet: number, ref: string): string { return `${sheet}!${ref}`; }
export function parseKey(key: string): { sheet: number; ref: string } {
  const k = key.indexOf('!');
  return { sheet: Number(key.slice(0, k)), ref: key.slice(k + 1) };
}

export class SheetSession {
  readonly original: Uint8Array;
  readonly filename: string;
  readonly doc: XlsxDocument;
  /** "시트번호!칸" → 새 글자 */
  readonly edits = new Map<string, string>();

  constructor(bytes: Uint8Array, filename: string) {
    this.original = bytes;
    this.filename = filename;
    this.doc = XlsxDocument.open(bytes);
  }

  get changeCount(): number { return this.edits.size; }

  /** 입력을 검사하고 수정 목록에 넣는다. 잘못된 입력은 EditError(목록은 바뀌지 않음). 원래 값과 같아지면 목록에서 뺀다. */
  setText(key: string, text: string): void {
    const { sheet, ref } = parseKey(key);
    this.doc.checkInput(sheet, ref, text);
    const orig = this.doc.cell(sheet, ref);
    if (text === orig.text) this.edits.delete(key);
    else this.edits.set(key, text);
  }

  revert(key: string): void { this.edits.delete(key); }
  revertAll(): void { this.edits.clear(); }

  /** 화면에 보일 글자: 수정이 있으면 수정 값, 없으면 원본 값 */
  textOf(sheet: number, ref: string): string { return this.edits.get(cellKey(sheet, ref)) ?? this.doc.cell(sheet, ref).text; }

  save(): SaveResult {
    const fresh = XlsxDocument.open(this.original);
    const keys = [...this.edits.keys()].sort((a, b) => {
      const pa = parseKey(a), pb = parseKey(b);
      if (pa.sheet !== pb.sheet) return pa.sheet - pb.sheet;
      const ra = parseRef(pa.ref), rb = parseRef(pb.ref);
      return ra.row - rb.row || ra.col - rb.col;
    });
    for (const k of keys) { const { sheet, ref } = parseKey(k); fresh.setCell(sheet, ref, this.edits.get(k)!); }
    const bytes = fresh.save();
    const problems: string[] = [];
    try {
      const again = XlsxDocument.open(bytes);
      if (again.sheets().length !== this.doc.sheets().length) problems.push('시트 수가 달라졌습니다');
      // 수정한 칸: 기대한 종류와 값인지
      for (const k of keys) {
        const { sheet, ref } = parseKey(k);
        const want = this.doc.checkInput(sheet, ref, this.edits.get(k)!);
        const got = again.cell(sheet, ref);
        const ok = want.type === 'empty' ? got.kind === 'empty'
          : want.type === 'number' ? got.kind === 'number' && Number(got.text) === Number(want.value)
          : want.type === 'bool' ? got.kind === 'bool' && (got.text === 'TRUE') === (want.value === '1')
          : got.kind === 'string' && got.text === want.value;
        if (!ok) problems.push(`${this.doc.sheets()[sheet].name}!${ref}: 저장본의 값이 다릅니다`);
      }
      // 수정하지 않은 칸: 모두 같은지(큰 시트는 건너뜀)
      const edited = new Set(keys);
      this.doc.sheets().forEach((s, i) => {
        const a = this.doc.usedRange(i), b = again.usedRange(i);
        const rows = Math.max(a.rows, b.rows), cols = Math.max(a.cols, b.cols);
        if (rows * cols > VERIFY_LIMIT) return;
        for (let r = 1; r <= rows; r++) for (let c = 1; c <= cols; c++) {
          const ref = `${numToCol(c)}${r}`;
          if (edited.has(cellKey(i, ref))) continue;
          const x = this.doc.cell(i, ref), y = again.cell(i, ref);
          if (x.kind !== y.kind || x.text !== y.text) { problems.push(`${s.name}!${ref}: 수정하지 않은 칸이 달라졌습니다`); return; }
        }
      });
    } catch (e) {
      problems.push(`저장본을 다시 열지 못했습니다: ${(e as Error).message}`);
    }
    return { bytes, verified: problems.length === 0, problems };
  }
}
