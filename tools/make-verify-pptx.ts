// PowerPoint 직접 열기 검증용 PPTX 복사본을 verify-out/pptx/에 만든다. 사용: node tools/make-verify-pptx.ts
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { PptxDocument } from '../src/pptx.ts';

const out = new URL('../verify-out/pptx/', import.meta.url);
mkdirSync(out, { recursive: true });
const orig = new Uint8Array(readFileSync(new URL('../fixtures/pptx/basic.pptx', import.meta.url)));
const put = (name: string, b: Uint8Array) => { const u = new URL(name, out); if (!existsSync(u) || Buffer.compare(readFileSync(u), Buffer.from(b)) !== 0) writeFileSync(u, b); };
put('P-basic-original.pptx', orig);
const d = PptxDocument.open(orig);
const ps = d.paragraphs();
const title = ps.find(p => p.text.startsWith('제목 슬라이드'))!;
d.setParagraphText(title.section, title.index, 'EDIT & <확인> "따옴표" 😀');
const run = ps.find(p => p.text === '굵게 보통 글자')!;
d.setParagraphText(run.section, run.index, '굵게 보통 추가 글자');
const cell = ps.filter(p => p.context === 'cell')[1];
d.setParagraphText(cell.section, cell.index, 'EDIT 셀');
const tb = ps.find(p => p.text.startsWith('텍스트 상자'))!;
d.setParagraphText(tb.section, tb.index, 'EDIT 상자');
const note = ps.find(p => p.group === '슬라이드 3 노트' && p.text.startsWith('발표자'))!;
d.setParagraphText(note.section, note.index, 'EDIT 노트');
put('P-basic-edited.pptx', d.save());
console.log('ok');
