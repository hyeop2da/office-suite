import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, writeFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { DocxDocument } from '../src/docx.ts';
import { ZipArchive } from '../src/zip.ts';
import { FormatError } from '../src/xml-scan.ts';

const dir = new URL('../fixtures/docx/', import.meta.url);
const load = (f: string) => new Uint8Array(readFileSync(new URL(f, dir)));

/** 독립 검증: python-docx가 읽은 문단·표·머리글 글자. */
function pythonRead(bytes: Uint8Array): { paragraphs: string[]; cells: string[]; header: string[]; footer: string[] } | null {
  const p = join(mkdtempSync(join(tmpdir(), 'docx-')), 'x.docx');
  writeFileSync(p, bytes);
  const r = spawnSync('python', ['-I', '-X', 'utf8', '-c', `
import sys, json, docx
d = docx.Document(sys.argv[1])
s = d.sections[0]
print(json.dumps({
 'paragraphs': [x.text for x in d.paragraphs],
 'cells': [c.text for t in d.tables for r in t.rows for c in r.cells],
 'header': [x.text for x in s.header.paragraphs],
 'footer': [x.text for x in s.footer.paragraphs]}, ensure_ascii=False))`, p], { encoding: 'utf8' });
  if (r.error) return null;
  assert.equal(r.status, 0, r.stderr);
  return JSON.parse(r.stdout);
}

test('무수정 저장은 파일 전체 바이트가 같다', () => {
  for (const f of ['basic', 'table', 'headerfooter', 'textbox']) {
    const orig = load(`${f}.docx`);
    assert.equal(Buffer.compare(Buffer.from(DocxDocument.open(orig).save()), Buffer.from(orig)), 0, f);
  }
});

test('본문 글자 교체: 바뀌는 항목은 document.xml뿐이고 특수문자가 안전하다', (t) => {
  const orig = load('basic.docx');
  const doc = DocxDocument.open(orig);
  const paras = doc.paragraphs();
  const p0 = paras.find(p => p.text.startsWith('안녕하세요'))!;
  doc.setParagraphText(p0.section, p0.index, 'A & B <c> "d" \'e\' 한글 😀');
  const out = doc.save();
  const a = new ZipArchive(orig), b = new ZipArchive(out);
  assert.deepEqual(b.entries.map(e => e.name), a.entries.map(e => e.name));
  for (const e of a.entries) if (e.name !== 'word/document.xml') assert.deepEqual(b.read(e.name), a.read(e.name), e.name);
  const py = pythonRead(out);
  if (!py) return t.skip('python 없음');
  assert.ok(py.paragraphs.includes('A & B <c> "d" \'e\' 한글 😀'));
  assert.ok(py.paragraphs.includes('  앞뒤 공백이 있는 문단  '), '다른 문단은 그대로');
});

test('앞뒤 공백이 있는 새 글자에는 xml:space="preserve"가 붙는다', (t) => {
  const doc = DocxDocument.open(load('basic.docx'));
  const p = doc.paragraphs().find(x => x.text === '이모지 😀 와 한자 漢字')!;
  doc.setParagraphText(p.section, p.index, '  앞뒤  공백  ');
  const out = doc.save();
  assert.match(new ZipArchive(out).readText('word/document.xml'), /<w:t xml:space="preserve">  앞뒤  공백  <\/w:t>/);
  const py = pythonRead(out);
  if (py) assert.ok(py.paragraphs.includes('  앞뒤  공백  '));
});

test('여러 글자 모양(런)이 있는 문단: 삽입은 앞 조각 모양을 따르고 서식 요소는 그대로다', () => {
  const orig = load('basic.docx');
  const doc = DocxDocument.open(orig);
  const p = doc.paragraphs().find(x => x.text === '굵게 보통 글자')!;
  assert.equal(p.runs, 2);
  doc.setParagraphText(p.section, p.index, '굵게 보통 추가 글자');
  const xml = new ZipArchive(doc.save()).readText('word/document.xml');
  const before = new ZipArchive(orig).readText('word/document.xml');
  assert.equal(xml.match(/<w:rPr>.*?<\/w:rPr>/g)?.join(), before.match(/<w:rPr>.*?<\/w:rPr>/g)?.join(), '서식 요소 동일');
  assert.ok(xml.includes('굵게'));
  assert.ok(DocxDocument.open(doc.save()).paragraphs().some(x => x.text === '굵게 보통 추가 글자'));
});

test('문단 중간 삭제와 전체 삭제', () => {
  const doc = DocxDocument.open(load('basic.docx'));
  const p = doc.paragraphs().find(x => x.text === '굵게 보통 글자')!;
  doc.setParagraphText(p.section, p.index, '굵 글자'); // 두 조각에 걸친 삭제
  assert.ok(DocxDocument.open(doc.save()).paragraphs().some(x => x.text === '굵 글자'));
  const q = doc.paragraphs().find(x => x.text.startsWith('안녕하세요'))!;
  doc.setParagraphText(q.section, q.index, '');
  const out = DocxDocument.open(doc.save());
  assert.ok(!out.paragraphs().some(x => x.text.startsWith('안녕하세요')));
});

test('표 셀 글자를 고치고 표 구조는 그대로 둔다', (t) => {
  const orig = load('table.docx');
  const doc = DocxDocument.open(orig);
  const cells = doc.paragraphs().filter(p => p.context === 'cell');
  assert.equal(cells.length, 4);
  doc.setParagraphText(cells[1].section, cells[1].index, '고친 셀');
  const out = doc.save();
  const py = pythonRead(out);
  if (!py) return t.skip('python 없음');
  assert.deepEqual(py.cells, ['셀 1-1', '고친 셀', '셀 2-1', '셀 2-2']);
  const strip = (x: Uint8Array) => new ZipArchive(x).readText('word/document.xml').replace(/<w:t[^>]*>[^<]*<\/w:t>/g, '<w:t/>');
  assert.equal(strip(out), strip(orig), '글자 요소 밖은 바이트 동일');
});

test('머리글·바닥글 글자를 고친다', (t) => {
  const doc = DocxDocument.open(load('headerfooter.docx'));
  const h = doc.paragraphs().find(p => p.context === 'header')!;
  const f = doc.paragraphs().find(p => p.context === 'footer')!;
  doc.setParagraphText(h.section, h.index, '새 머리글');
  doc.setParagraphText(f.section, f.index, '새 바닥글');
  const py = pythonRead(doc.save());
  if (!py) return t.skip('python 없음');
  assert.deepEqual(py.header, ['새 머리글']);
  assert.deepEqual(py.footer, ['새 바닥글']);
});

test('텍스트 상자: 호환용 대체 내용(mc:Fallback)은 목록에서 빠지고 고칠 수 없다', () => {
  const doc = DocxDocument.open(load('textbox.docx'));
  const boxes = doc.paragraphs().filter(p => p.text === '상자 안 글자');
  assert.equal(boxes.length, 1, '선택(Choice) 쪽 하나만 보인다');
  assert.equal(boxes[0].context, 'textbox');
  const part = doc.part('word/document.xml');
  const fb = part.scan().texts.find(t => t.context === 'fallback')!;
  assert.throws(() => part.setText(fb.index, 'X'), /호환용 대체 내용/);
  // 알려진 한계: 선택 쪽만 고치면 대체 쪽은 옛 글자가 남는다
  doc.setParagraphText(boxes[0].section, boxes[0].index, '새 상자');
  const xml = new ZipArchive(doc.save()).readText('word/document.xml');
  assert.ok(xml.includes('새 상자') && xml.includes('상자 안 글자'));
});

test('DOCX가 아닌 파일은 거부한다', () => {
  assert.throws(() => DocxDocument.open(new Uint8Array([1, 2, 3])), /ZIP/);
  assert.throws(() => DocxDocument.open(new Uint8Array(readFileSync(new URL('../fixtures/hwpx/pps-manual-14-7.hwpx', import.meta.url)))), FormatError);
});
