// PDF 메모 세션: 원본은 바꾸지 않고, 추가할 텍스트 메모 목록만 쌓았다가 저장할 때 증분 갱신으로 덧붙인다.
import { PdfDocument, PdfError, type TextNote } from './pdf.ts';
import type { SaveResult } from './session.ts';

export interface PendingNote { key: string; page: number; x: number; y: number; contents: string }

export class PdfSession {
  readonly original: Uint8Array;
  readonly filename: string;
  readonly doc: PdfDocument;
  /** "쪽|x|y|순번" → 메모 글자 (추가한 순서 유지) */
  readonly edits = new Map<string, string>();
  private seq = 0;

  constructor(bytes: Uint8Array, filename: string) {
    this.original = bytes;
    this.filename = filename;
    this.doc = PdfDocument.open(bytes);
  }

  get changeCount(): number { return this.edits.size; }
  get existing(): TextNote[] { return this.doc.notes(); }

  private parse(key: string): { page: number; x: number; y: number } {
    const m = /^(\d+)\|(-?\d+(?:\.\d+)?)\|(-?\d+(?:\.\d+)?)\|(\d+)$/.exec(key);
    if (!m) throw new PdfError(`메모 식별자가 올바르지 않습니다: ${key}`);
    const page = Number(m[1]);
    if (page >= this.doc.pages().length) throw new PdfError(`쪽 번호가 범위 밖입니다: ${page + 1}`);
    this.seq = Math.max(this.seq, Number(m[4]) + 1);
    return { page, x: Number(m[2]), y: Number(m[3]) };
  }

  /** 새 메모를 추가하고 식별자를 돌려준다. */
  addNote(page: number, x: number, y: number, text: string): string {
    const key = `${page}|${x.toFixed(1)}|${y.toFixed(1)}|${this.seq++}`;
    this.setText(key, text);
    return key;
  }

  /** 식별자로 메모 글자를 정한다(임시 보존 복구에도 쓴다). 빈 글자는 메모를 지운다. */
  setText(key: string, text: string): void {
    this.parse(key);
    if (typeof text !== 'string') throw new PdfError('메모는 글자여야 합니다');
    if (text.trim() === '') this.edits.delete(key);
    else {
      if (text.length > 5000) throw new PdfError('메모는 5,000자를 넘을 수 없습니다');
      this.edits.set(key, text);
    }
  }

  revert(key: string): void { this.edits.delete(key); }
  revertAll(): void { this.edits.clear(); }

  pending(): PendingNote[] {
    return [...this.edits].map(([key, contents]) => ({ key, contents, ...this.parse(key) }));
  }

  save(): SaveResult {
    const list = this.pending();
    const bytes = this.doc.withNotes(list.map(p => ({ page: p.page, x: p.x, y: p.y, contents: p.contents })));
    const problems: string[] = [];
    try {
      const again = PdfDocument.open(bytes);
      if (again.pages().length !== this.doc.pages().length) problems.push('쪽 수가 달라졌습니다');
      const before = this.doc.notes();
      const after = again.notes();
      if (after.length !== before.length + list.length) problems.push(`메모 수가 ${before.length + list.length}개여야 하는데 ${after.length}개입니다`);
      // 모든 쪽의 메모를 (쪽 순서, 쪽 안 추가 순서)로 맞춰 본다: 기존 메모 뒤에 새 메모가 붙는다
      const expected = new Map<number, string[]>();
      for (const n of before) { const l = expected.get(n.page) ?? []; l.push(n.contents); expected.set(n.page, l); }
      for (const p of list) { const l = expected.get(p.page) ?? []; l.push(p.contents); expected.set(p.page, l); }
      const got = new Map<number, string[]>();
      for (const n of after) { const l = got.get(n.page) ?? []; l.push(n.contents); got.set(n.page, l); }
      for (const [pg, l] of expected) if (JSON.stringify(got.get(pg) ?? []) !== JSON.stringify(l)) problems.push(`${pg + 1}쪽의 메모가 기대와 다릅니다`);
      if (bytes !== this.original && !(bytes.length >= this.original.length && this.original.every((b, i) => bytes[i] === b))) problems.push('원본 바이트가 보존되지 않았습니다');
    } catch (e) {
      problems.push(`저장본을 다시 열지 못했습니다: ${(e as Error).message}`);
    }
    return { bytes, verified: problems.length === 0, problems };
  }
}
