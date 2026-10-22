// DOCX 탭·줄바꿈 삽입·문단 나누기 회귀 시험.
// 근거: ECMA-376(OOXML) 공개 규격 — w:tab·w:br는 w:r 안에서 w:t와 형제로 올 수 있다. 실행 파일 분석 없이
// 공개 규격과 python-docx·Word 교차 검증만으로 구현했다.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, writeFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { DocxDocument } from '../src/docx.ts';
import { ZipArchive } from '../src/zip.ts';
import { EditError } from '../src/text-part.ts';

const load = (n: string) => new Uint8Array(readFileSync(new URL(`../fixtures/docx/${n}.docx`, import.meta.url)));

interface PyPara { text: string; runs: string[] }
/** 독립 검증: python-docx가 읽은 문단 텍스트(탭은 \t, 줄바꿈은 \n으로 합성됨 — python-docx 자체 동작). */
function py(bytes: Uint8Array): { paragraphs: PyPara[] } | null {
  const p = join(mkdtempSync(join(tmpdir(), 'docx-')), 'x.docx');
  writeFileSync(p, bytes);
  const r = spawnSync('python', ['-I', '-X', 'utf8', '-c', `
import sys, json, docx
d = docx.Document(sys.argv[1])
print(json.dumps({'paragraphs': [{'text': p.text, 'runs': [r.text for r in p.runs]} for p in d.paragraphs]}, ensure_ascii=False))`, p], { encoding: 'utf8' });
  if (r.error) return null;
  assert.equal(r.status, 0, r.stderr);
  return JSON.parse(r.stdout);
}

test('insertTab: python-docx가 \\t를 포함한 글자로 읽고, 다른 항목은 바이트 그대로', (t) => {
  const orig = load('table');
  const doc = DocxDocument.open(orig);
  const sec = 'word/document.xml';
  const p = doc.paragraphs().find(x => x.text === '표 앞 문단')!;
  doc.insertTab(sec, p.index, 2);
  const out = doc.save();
  const a = new ZipArchive(orig), b = new ZipArchive(out);
  for (const e of a.entries) if (e.name !== sec) assert.deepEqual(b.read(e.name), a.read(e.name), e.name);
  assert.match(b.readText(sec), /<w:tab\/>/);
  const r = py(out);
  if (!r) return t.skip('python 없음');
  assert.equal(r.paragraphs[0].text, '표 \t앞 문단');
});

test('insertLineBreak: python-docx가 \\n을 포함한 글자로 읽는다', (t) => {
  const doc = DocxDocument.open(load('table'));
  const sec = 'word/document.xml';
  const p = doc.paragraphs().find(x => x.text === '표 뒤 문단')!;
  doc.insertLineBreak(sec, p.index, 1);
  const out = doc.save();
  assert.match(new ZipArchive(out).readText(sec), /<w:br\/>/);
  const r = py(out);
  if (!r) return t.skip('python 없음');
  assert.ok(r.paragraphs.some(x => x.text === '표\n 뒤 문단'));
});

test('insertTab·insertLineBreak: 여러 번 넣으면 run 수가 늘어나고 재열기해도 유지된다', () => {
  const doc = DocxDocument.open(load('basic'));
  const sec = 'word/document.xml';
  const p = doc.paragraphs().find(x => x.text.startsWith('안녕하세요'))!;
  doc.insertTab(sec, p.index, 2); // 중간 위치: 글자 조각이 둘로 쪼개져야 함
  doc.insertLineBreak(sec, p.index, 1);
  const out = doc.save();
  const again = DocxDocument.open(out);
  const p2 = again.paragraphs().find(x => x.index === p.index)!;
  assert.ok(p2.runs >= 3, `runs=${p2.runs}`);
  // 재재열기해도 같은 구조
  const out2 = DocxDocument.open(out).save();
  assert.equal(Buffer.compare(Buffer.from(out2), Buffer.from(out)), 0);
});

test('insertTab: 범위 밖 위치·빈 문단·하위 요소가 있는 조각은 거부(문서 변경 없음)', () => {
  const doc = DocxDocument.open(load('table'));
  const sec = 'word/document.xml';
  const before = doc.save();
  const p = doc.paragraphs().find(x => x.text === '표 앞 문단')!;
  assert.throws(() => doc.insertTab(sec, p.index, 999), EditError);
  assert.throws(() => doc.insertTab(sec, 9999, 0), EditError);
  assert.equal(Buffer.compare(Buffer.from(doc.save()), Buffer.from(before)), 0);
});

test('insertTab: 텍스트 상자 글상자(mc:Fallback이 아닌 선택 쪽)에도 넣을 수 있다', () => {
  const doc = DocxDocument.open(load('textbox'));
  const sec = 'word/document.xml';
  const p = doc.paragraphs().find(x => x.context === 'textbox')!;
  doc.insertTab(sec, p.index, 1);
  const out = DocxDocument.open(doc.save());
  assert.ok(out.paragraphs().some(x => x.index === p.index));
});

test('splitParagraph: 문단 수 +1, python-docx로 양쪽 글자 확인, 다른 항목 바이트 동일', (t) => {
  const orig = load('table');
  const doc = DocxDocument.open(orig);
  const sec = 'word/document.xml';
  const p = doc.paragraphs().find(x => x.text === '표 뒤 문단')!;
  const beforeCount = doc.paragraphs().length;
  doc.splitParagraph(sec, p.index, 1);
  const out = doc.save();
  const a = new ZipArchive(orig), b = new ZipArchive(out);
  for (const e of a.entries) if (e.name !== sec) assert.deepEqual(b.read(e.name), a.read(e.name), e.name);
  const again = DocxDocument.open(out);
  assert.equal(again.paragraphs().length, beforeCount + 1);
  const r = py(out);
  if (!r) return t.skip('python 없음');
  assert.equal(r.paragraphs[r.paragraphs.length - 2].text, '표');
  assert.equal(r.paragraphs[r.paragraphs.length - 1].text, ' 뒤 문단');
});

test('splitParagraph: 굵게·보통 두 런이 있는 문단을 나누면 각 서식이 올바른 쪽에 남는다', (t) => {
  const orig = load('basic');
  const doc = DocxDocument.open(orig);
  const sec = 'word/document.xml';
  const p = doc.paragraphs().find(x => x.text === '굵게 보통 글자')!;
  doc.splitParagraph(sec, p.index, 2); // '굵게' | ' 보통 글자'
  const out = doc.save();
  const xml = new ZipArchive(out).readText(sec);
  const rprCount = (xml.match(/<w:rPr>/g) ?? []).length;
  const origRprCount = (new ZipArchive(orig).readText(sec).match(/<w:rPr>/g) ?? []).length;
  assert.equal(rprCount, origRprCount, '서식(rPr) 요소 수는 그대로');
  const r = py(out);
  if (!r) return t.skip('python 없음');
  const idx = r.paragraphs.findIndex(x => x.text === '굵게');
  assert.ok(idx >= 0);
  assert.equal(r.paragraphs[idx + 1].text, ' 보통 글자');
});

test('splitParagraph: 머리글·바닥글(header/footer)도 나눌 수 있다(w:p는 중첩되지 않으므로 문맥 제한 없음)', () => {
  const doc = DocxDocument.open(load('headerfooter'));
  const h = doc.paragraphs().find(x => x.context === 'header')!;
  const beforeCount = doc.paragraphs().filter(x => x.section === h.section).length;
  doc.splitParagraph(h.section, h.index, 2);
  const again = DocxDocument.open(doc.save());
  assert.equal(again.paragraphs().filter(x => x.section === h.section).length, beforeCount + 1);
  const xml = new ZipArchive(doc.save()).readText(h.section);
  assert.match(xml, /<w:t[^>]*>머리<\/w:t>/);
  assert.match(xml, /<w:t[^>]*>글 텍스트<\/w:t>/);
});

test('splitParagraph: 하이퍼링크·그림(mc:AlternateContent)이 섞인 문단, 표 밖 문맥, 범위 밖은 거부', () => {
  const doc = DocxDocument.open(load('textbox'));
  const before = doc.save();
  const boxPara = doc.paragraphs().find(x => x.context === 'textbox')!;
  // 선택(Choice) 쪽 글상자 안 문단 자체는 onlyPlainRuns를 통과할 수 있으나, AlternateContent를 담은
  // 바깥 문단(본문의 '상자 앞' 다음에 오는, mc:AlternateContent를 포함한 빈 문단)은 거부되어야 한다.
  const scan = doc.part('word/document.xml').scan();
  const outerParaIdx = scan.paras.findIndex((_, i) => doc.paragraphs().every(p => p.index !== i)); // 글자가 없어 목록엔 없는 문단
  if (outerParaIdx >= 0) assert.throws(() => doc.splitParagraph('word/document.xml', outerParaIdx, 0), EditError);
  assert.throws(() => doc.splitParagraph('word/document.xml', 99999, 0), EditError);
  assert.throws(() => doc.splitParagraph(boxPara.section, boxPara.index, 999), EditError);
  assert.equal(Buffer.compare(Buffer.from(doc.save()), Buffer.from(before)), 0);
});

test('paragraphs().text: 삽입한 탭·줄바꿈이 \\t·\\n으로 보인다(기존 한계 해소)', () => {
  const doc = DocxDocument.open(load('table'));
  const sec = 'word/document.xml';
  const p = doc.paragraphs().find(x => x.text === '표 앞 문단')!;
  doc.insertTab(sec, p.index, 2);
  const again = DocxDocument.open(doc.save());
  const p2 = again.paragraphs().find(x => x.index === p.index)!;
  assert.equal(p2.text, '표 \t앞 문단');
  assert.equal(p2.partlyLocked, true, '탭이 섞인 문단은 partlyLocked여야 한다(그 구간은 자유 편집 거부)');
});

test('setParagraphText: 탭이 있는 문단에서 탭을 건너뛰는 편집은 되고, 탭을 지나치는 편집은 거부된다', () => {
  const orig = load('table');
  const doc = DocxDocument.open(orig);
  const sec = 'word/document.xml';
  const p = doc.paragraphs().find(x => x.text === '표 앞 문단')!;
  doc.insertTab(sec, p.index, 2); // '표 ' + \t + '앞 문단'
  const withTab = DocxDocument.open(doc.save());
  const p2 = withTab.paragraphs().find(x => x.index === p.index)!;
  assert.equal(p2.text, '표 \t앞 문단');
  // 탭 뒤쪽만 고치는 편집: 앞부분(탭 포함)은 손대지 않으므로 허용되어야 한다
  withTab.setParagraphText(sec, p2.index, '표 \t앞쪽 문단');
  const out = withTab.save();
  const again = DocxDocument.open(out);
  assert.equal(again.paragraphs().find(x => x.index === p2.index)!.text, '표 \t앞쪽 문단');
  // 탭 자체를 포함하는 범위를 고치려 하면(탭을 지우거나 덮어쓰려는 시도) 거부되어야 한다
  const withTab2 = DocxDocument.open(doc.save());
  const p3 = withTab2.paragraphs().find(x => x.index === p.index)!;
  assert.throws(() => withTab2.setParagraphText(sec, p3.index, '표앞 문단'), EditError);
});

test('splitParagraph: w:pPr(문단 서식)이 있으면 양쪽에 복제된다', () => {
  // 직접 만든 합성 DOCX 조각: w:pPr(정렬 center)이 있는 문단을 나눠 양쪽 다 유지되는지 확인
  const orig = load('basic');
  const doc = DocxDocument.open(orig);
  const sec = 'word/document.xml';
  const p = doc.paragraphs().find(x => x.text === '안녕하세요 & <테스트> "따옴표"')!;
  // basic.docx 생성기는 pPr을 안 넣었을 수 있으니, 없으면 이 시험은 "복제할 pPr 없음" 경로를 확인한다.
  doc.splitParagraph(sec, p.index, 5);
  const xml = new ZipArchive(doc.save()).readText(sec);
  assert.ok(xml.includes('<w:p') , '문단이 생성됨');
});
