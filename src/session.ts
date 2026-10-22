// 편집 세션: 원본 바이트는 건드리지 않고, 문단별 수정 목록만 쌓았다가 저장할 때 한 번에 적용한다.
// UI와 분리되어 있어 Node 테스트로 검증한다.
import { openDocument, type DocParagraph, type DocumentEngine } from './engine.ts';
import type { EditOptions } from './text-part.ts';

export interface SaveResult {
  bytes: Uint8Array;
  /** 저장본을 다시 열어 모든 수정이 반영되었는지 확인한 결과 */
  verified: boolean;
  problems: string[];
}

export class EditSession {
  readonly original: Uint8Array;
  readonly filename: string;
  readonly paragraphs: DocParagraph[];
  /** 문단 id → 새 글자 */
  readonly edits = new Map<string, string>();
  readonly engine: DocumentEngine;

  constructor(bytes: Uint8Array, filename: string) {
    this.original = bytes;
    this.filename = filename;
    this.engine = openDocument(bytes, filename);
    this.paragraphs = this.engine.paragraphs();
  }

  originalText(id: string): string {
    const p = this.paragraphs.find(x => x.id === id);
    if (!p) throw new Error(`문단 없음: ${id}`);
    return p.text;
  }

  /** 수정 기록. 원래 글자와 같아지면 수정 목록에서 뺀다. */
  setText(id: string, text: string): void {
    if (text === this.originalText(id)) this.edits.delete(id);
    else this.edits.set(id, text);
  }

  revert(id: string): void { this.edits.delete(id); }
  revertAll(): void { this.edits.clear(); }
  get changeCount(): number { return this.edits.size; }

  /** 수정 목록을 원본에 적용해 새 파일을 만들고 다시 열어 확인한다. */
  save(opt: EditOptions = {}): SaveResult {
    const fresh = openDocument(this.original, this.filename);
    for (const [id, text] of this.edits) fresh.setParagraphText(id, text, opt);
    const bytes = fresh.save();
    const problems: string[] = [];
    try {
      const again = openDocument(bytes, this.filename).paragraphs();
      const byId = new Map(again.map(p => [p.id, p.text]));
      for (const p of this.paragraphs) {
        const want = this.edits.get(p.id) ?? p.text;
        if (byId.get(p.id) !== want) problems.push(`${p.id}: 저장본 글자가 다릅니다`);
      }
      if (again.length !== this.paragraphs.length) problems.push(`문단 수가 ${this.paragraphs.length}에서 ${again.length}로 달라졌습니다`);
    } catch (e) {
      problems.push(`저장본을 다시 열지 못했습니다: ${(e as Error).message}`);
    }
    return { bytes, verified: problems.length === 0, problems };
  }
}
