import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, writeFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { PptxDocument } from '../src/pptx.ts';
import { openDocument, detectFormat } from '../src/engine.ts';
import { ZipArchive } from '../src/zip.ts';
import { EditError } from '../src/text-part.ts';
import { FormatError } from '../src/xml-scan.ts';

const load = (n = 'basic') => new Uint8Array(readFileSync(new URL(`../fixtures/pptx/${n}.pptx`, import.meta.url)));

/** 독립 검증: python-pptx가 읽은 슬라이드별 도형 글자, 표 셀, 노트 */
function py(bytes: Uint8Array): { slides: { shapes: string[]; cells: string[]; notes: string }[] } | null {
  const p = join(mkdtempSync(join(tmpdir(), 'pptx-')), 'x.pptx');
  writeFileSync(p, bytes);
  const r = spawnSync('python', ['-I', '-X', 'utf8', '-c', `
import sys, json, pptx
prs = pptx.Presentation(sys.argv[1])
out = []
for s in prs.slides:
    shapes, cells = [], []
    for sh in s.shapes:
        if sh.has_text_frame: shapes.append(sh.text_frame.text)
        if getattr(sh, 'has_table', False) and sh.has_table:
            for row in sh.table.rows:
                for c in row.cells: cells.append(c.text)
    out.append({'shapes': shapes, 'cells': cells, 'notes': s.notes_slide.notes_text_frame.text if s.has_notes_slide else ''})
print(json.dumps({'slides': out}, ensure_ascii=False))`, p], { encoding: 'utf8' });
  if (r.error) return null;
  assert.equal(r.status, 0, r.stderr);
  return JSON.parse(r.stdout);
}

test('형식 판별과 엔진 연결', () => {
  assert.equal(detectFormat(load()).id, 'pptx');
  assert.equal(openDocument(load()).format.id, 'pptx');
});

test('슬라이드 순서·노트 연결·묶음 이름', () => {
  const d = PptxDocument.open(load());
  assert.deepEqual(d.slides().map(s => [s.number, !!s.notesPart]), [[1, false], [2, false], [3, true]]);
  const g = d.paragraphs().map(p => p.group);
  assert.ok(g.includes('슬라이드 1') && g.includes('슬라이드 3 노트'));
});

test('무수정 저장은 파일 전체 바이트가 같다', () => {
  const o = load();
  assert.equal(Buffer.compare(Buffer.from(PptxDocument.open(o).save()), Buffer.from(o)), 0);
});

test('필드(슬라이드 번호) 문단은 보기만 가능하고, 필드가 낀 구간 편집은 거부한다', () => {
  const d = PptxDocument.open(load());
  const f = d.paragraphs().find(p => p.text.startsWith('번호:'))!;
  assert.equal(f.partlyLocked, true);
  // 필드 앞 일반 글자 조각만 바꾸는 것은 가능하지만, 필드 글자가 바뀌는 범위는 거부
  assert.throws(() => d.setParagraphText(f.section, f.index, '번호: 9'), EditError);
});

test('본문 글자 교체: 특수문자·이모지, 나머지 항목은 바이트 동일', (t) => {
  const o = load();
  const d = PptxDocument.open(o);
  const title = d.paragraphs().find(p => p.text.startsWith('제목 슬라이드'))!;
  d.setParagraphText(title.section, title.index, 'EDIT & <확인> "따옴표" \'a\' 😀');
  const out = d.save();
  const a = new ZipArchive(o), b = new ZipArchive(out);
  assert.deepEqual(b.entries.map(e => e.name), a.entries.map(e => e.name));
  for (const e of a.entries) if (e.name !== 'ppt/slides/slide1.xml') assert.deepEqual(b.read(e.name), a.read(e.name), e.name);
  const r = py(out);
  if (!r) return t.skip('python 없음');
  assert.ok(r.slides[0].shapes[0] === 'EDIT & <확인> "따옴표" \'a\' 😀');
  assert.equal(r.slides[2].shapes[0], '텍스트 상자 😀 漢字', '다른 슬라이드는 그대로');
});

test('여러 글자 모양(런) 문단: 서식 요소 그대로, 삽입은 앞 조각을 따름', () => {
  const o = load();
  const d = PptxDocument.open(o);
  const p = d.paragraphs().find(x => x.text === '굵게 보통 글자')!;
  assert.equal(p.runs, 2);
  d.setParagraphText(p.section, p.index, '굵게 보통 추가 글자');
  const x = new ZipArchive(d.save()).readText('ppt/slides/slide1.xml');
  const before = new ZipArchive(o).readText('ppt/slides/slide1.xml');
  assert.equal(x.match(/<a:rPr[^>]*>/g)?.join(), before.match(/<a:rPr[^>]*>/g)?.join(), '런 속성 동일');
  assert.ok(PptxDocument.open(d.save()).paragraphs().some(q => q.text === '굵게 보통 추가 글자'));
});

test('표 셀과 발표자 노트 편집', (t) => {
  const d = PptxDocument.open(load());
  const c = d.paragraphs().filter(p => p.context === 'cell')[1];
  d.setParagraphText(c.section, c.index, 'EDIT 셀');
  const n = d.paragraphs().find(p => p.group === '슬라이드 3 노트' && p.text.startsWith('발표자'))!;
  d.setParagraphText(n.section, n.index, 'EDIT 노트');
  const r = py(d.save());
  if (!r) return t.skip('python 없음');
  assert.deepEqual(r.slides[1].cells, ['셀 1-1', 'EDIT 셀', '셀 2-1', '셀 2-2']);
  assert.equal(r.slides[2].notes, 'EDIT 노트');
});

test('거부: 줄바꿈·제어 문자, 잘못된 부품·번호, PPTX가 아닌 파일', () => {
  const d = PptxDocument.open(load());
  const before = d.save();
  assert.throws(() => d.setParagraphText('ppt/slides/slide1.xml', 0, 'a\nb'), EditError);
  assert.throws(() => d.setParagraphText('ppt/slides/slide9.xml', 0, 'x'), EditError);
  assert.throws(() => d.setParagraphText('ppt/slides/slide1.xml', 999, 'x'), EditError);
  assert.equal(Buffer.compare(Buffer.from(d.save()), Buffer.from(before)), 0);
  assert.throws(() => PptxDocument.open(new Uint8Array(readFileSync(new URL('../fixtures/docx/basic.docx', import.meta.url)))), FormatError);
});

// 탭·줄바꿈·문단 나누기 (2026-10-22 추가): ECMA-376/ISO 29500(DrawingML, 공개 규격) 확인 — a:br은 a:r의
// 형제, 탭은 전용 요소가 없어 a:t 안에 문자 그대로. python-pptx로 교차 검증(paragraph.text는 a:br을
// '\v'(U+000B)로 합성해 보여준다 — 디컴파일 아니라 python-pptx 자체 동작을 실측한 것).
test('setParagraphText: 탭 문자(U+0009)는 그대로 쓸 수 있다(전용 요소 없음, a:t에 리터럴)', (t) => {
  const orig = load();
  const d = PptxDocument.open(orig);
  const p = d.paragraphs().find(x => x.text === '첫째 줄')!;
  d.setParagraphText(p.section, p.index, '첫\t째 줄');
  const out = d.save();
  const a = new ZipArchive(orig), b = new ZipArchive(out);
  for (const e of a.entries) if (e.name !== p.section) assert.deepEqual(b.read(e.name), a.read(e.name), e.name);
  assert.equal(PptxDocument.open(out).paragraphs().find(x => x.index === p.index)!.text, '첫\t째 줄');
  const r = py(out);
  if (!r) return t.skip('python 없음');
  assert.equal(r.slides[0].shapes[1], '첫\t째 줄\n굵게 보통 글자\n  앞뒤 공백이 있는 줄  \n번호: 1');
});

test('insertLineBreak: a:br을 런 사이에 끼우고(서식 복제), python-pptx가 \v로 합성해 읽는다', (t) => {
  const orig = load();
  const d = PptxDocument.open(orig);
  const p = d.paragraphs().find(x => x.text === '첫째 줄')!;
  d.insertLineBreak(p.section, p.index, 2);
  const out = d.save();
  const a = new ZipArchive(orig), b = new ZipArchive(out);
  for (const e of a.entries) if (e.name !== p.section) assert.deepEqual(b.read(e.name), a.read(e.name), e.name);
  assert.match(b.readText(p.section), /<a:br\/>/);
  assert.equal(PptxDocument.open(out).paragraphs().find(x => x.index === p.index)!.text, '첫째\n 줄', 'paragraphs().text에도 \n으로 반영');
  const r = py(out);
  if (!r) return t.skip('python 없음');
  assert.equal(r.slides[0].shapes[1], '첫째\x0b 줄\n굵게 보통 글자\n  앞뒤 공백이 있는 줄  \n번호: 1');
});

test('insertLineBreak: 서식 있는 런에 끼우면 양쪽에 rPr이 복제된다', (t) => {
  const orig = load();
  const d = PptxDocument.open(orig);
  const p = d.paragraphs().find(x => x.text === '굵게 보통 글자')!;
  d.insertLineBreak(p.section, p.index, 1); // '굵' | 'ㅐ' 사이(런0: '굵게')
  const out = d.save();
  const xml = new ZipArchive(out).readText(p.section);
  assert.equal((xml.match(/<a:rPr b="1"\/>/g) ?? []).length, 2, '굵게 서식이 나뉜 두 런 모두에 남아야 함');
  const r = py(out);
  if (!r) return t.skip('python 없음');
  assert.equal(r.slides[0].shapes[1], '첫째 줄\n굵\x0b게 보통 글자\n  앞뒤 공백이 있는 줄  \n번호: 1');
});

test('insertLineBreak: 필드가 있는 문단은 거부(문서 변경 없음)', () => {
  const d = PptxDocument.open(load());
  const before = d.save();
  const p = d.paragraphs().find(x => x.text.startsWith('번호: '))!;
  assert.throws(() => d.insertLineBreak(p.section, p.index, 1), EditError);
  assert.equal(Buffer.compare(Buffer.from(d.save()), Buffer.from(before)), 0);
});

test('splitParagraph: 문단 수 +1, 서식·필드 보존, python-pptx로 양쪽 글자 확인', (t) => {
  const orig = load();
  const d = PptxDocument.open(orig);
  const p = d.paragraphs().find(x => x.text === '굵게 보통 글자')!;
  const beforeCount = d.paragraphs().length;
  d.splitParagraph(p.section, p.index, 2); // '굵게' | ' 보통 글자'
  const out = d.save();
  const a = new ZipArchive(orig), b = new ZipArchive(out);
  for (const e of a.entries) if (e.name !== p.section) assert.deepEqual(b.read(e.name), a.read(e.name), e.name);
  const again = PptxDocument.open(out);
  assert.equal(again.paragraphs().length, beforeCount + 1);
  const r = py(out);
  if (!r) return t.skip('python 없음');
  // python-pptx의 TextFrame.text는 문단을 '\n'으로 이어붙인다 — 나뉜 자리에 새 '\n'이 생기고,
  // 필드가 있는 마지막 문단('번호: 1')은 그대로 끝에 남아야 한다.
  assert.equal(r.slides[0].shapes[1], '첫째 줄\n굵게\n 보통 글자\n  앞뒤 공백이 있는 줄  \n번호: 1');
});

test('splitParagraph: 범위 밖 위치·필드가 있는 문단·빈 번호는 거부(문서 변경 없음)', () => {
  const d = PptxDocument.open(load());
  const before = d.save();
  const field = d.paragraphs().find(x => x.text.startsWith('번호: '))!;
  assert.throws(() => d.splitParagraph(field.section, field.index, 1), EditError);
  assert.throws(() => d.splitParagraph(field.section, 99999, 0), EditError);
  const p = d.paragraphs().find(x => x.text === '첫째 줄')!;
  assert.throws(() => d.splitParagraph(p.section, p.index, 999), EditError);
  assert.equal(Buffer.compare(Buffer.from(d.save()), Buffer.from(before)), 0);
});
