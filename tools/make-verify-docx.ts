// Word 직접 열기 검증(E6·E7)용 DOCX 복사본을 verify-out/docx/에 만든다. 사용: node tools/make-verify-docx.ts
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { DocxDocument } from '../src/docx.ts';

const out = new URL('../verify-out/docx/', import.meta.url);
mkdirSync(out, { recursive: true });
const load = (n: string) => new Uint8Array(readFileSync(new URL(`../fixtures/docx/${n}.docx`, import.meta.url)));
const put = (name: string, b: Uint8Array) => {
  const u = new URL(name, out);
  if (existsSync(u) && Buffer.compare(readFileSync(u), Buffer.from(b)) === 0) return;
  writeFileSync(u, b);
};

for (const n of ['basic', 'table', 'headerfooter', 'textbox']) put(`E6-${n}-original.docx`, load(n));

{ // basic: 본문 문단 교체(특수문자 포함), 공백 문단, 여러 글자 모양 문단의 삽입
  const d = DocxDocument.open(load('basic'));
  const ps = d.paragraphs();
  const a = ps.find(p => p.text.startsWith('안녕하세요'))!;
  d.setParagraphText(a.section, a.index, 'EDIT & <확인> "따옴표" 😀');
  const b = ps.find(p => p.text === '굵게 보통 글자')!;
  d.setParagraphText(b.section, b.index, '굵게 보통 추가 글자');
  const c = ps.find(p => p.text === '이모지 😀 와 한자 漢字')!;
  d.setParagraphText(c.section, c.index, '  앞뒤  공백  ');
  put('E6-basic-edited.docx', d.save());
}
{ // table: 셀 하나
  const d = DocxDocument.open(load('table'));
  const c = d.paragraphs().filter(p => p.context === 'cell')[1];
  d.setParagraphText(c.section, c.index, 'EDIT 셀');
  put('E6-table-edited.docx', d.save());
}
{ // headerfooter: 머리글·바닥글
  const d = DocxDocument.open(load('headerfooter'));
  const ps = d.paragraphs();
  const h = ps.find(p => p.context === 'header')!, f = ps.find(p => p.context === 'footer')!;
  d.setParagraphText(h.section, h.index, 'EDIT 머리글');
  d.setParagraphText(f.section, f.index, 'EDIT 바닥글');
  put('E6-headerfooter-edited.docx', d.save());
}
{ // E7: 글상자 선택(Choice) 쪽만 편집. 대체(Fallback) 쪽은 그대로 남는다.
  const d = DocxDocument.open(load('textbox'));
  const b = d.paragraphs().find(p => p.context === 'textbox')!;
  d.setParagraphText(b.section, b.index, 'EDIT 상자');
  put('E7-textbox-edited.docx', d.save());
}
console.log('ok');
