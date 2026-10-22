import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, writeFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { PdfDocument, PdfError } from '../src/pdf.ts';
import { detectFormat } from '../src/engine.ts';

const load = (n: string) => new Uint8Array(readFileSync(new URL(`../fixtures/pdf/${n}.pdf`, import.meta.url)));

interface PyPdf { pages: number; notes: { page: number; contents: string; subtype: string; rect: number[]; author: string | null }[]; text: string[]; links: number }
/** 독립 검증: pypdf가 읽은 쪽 수, 주석, 쪽별 글자 */
function py(bytes: Uint8Array): PyPdf | null {
  const p = join(mkdtempSync(join(tmpdir(), 'pdf-')), 'x.pdf');
  writeFileSync(p, bytes);
  const r = spawnSync('python', ['-I', '-X', 'utf8', '-c', `
import sys, json, pypdf
r = pypdf.PdfReader(sys.argv[1], strict=True)
notes, links = [], 0
for i, pg in enumerate(r.pages):
    for a in (pg.get('/Annots') or []):
        a = a.get_object()
        st = str(a.get('/Subtype'))
        if st == '/Text':
            notes.append({'page': i, 'contents': str(a.get('/Contents')), 'subtype': st, 'rect': [float(x) for x in a['/Rect']], 'author': str(a['/T']) if '/T' in a else None})
        elif st == '/Link': links += 1
print(json.dumps({'pages': len(r.pages), 'notes': notes, 'text': [pg.extract_text() for pg in r.pages], 'links': links}, ensure_ascii=False))`, p], { encoding: 'utf8' });
  if (r.error) return null;
  assert.equal(r.status, 0, r.stderr);
  return JSON.parse(r.stdout);
}

for (const [name, pages] of [['classic', 3], ['links', 2], ['xrefstream', 1]] as const) {
  test(`${name}: 쪽 수·쪽 크기를 읽는다`, () => {
    const d = PdfDocument.open(load(name));
    assert.equal(d.pages().length, pages);
    assert.ok(d.pages()[0].width > 100 && d.pages()[0].height > 100);
  });
}

test('쪽 크기: A4(classic)와 사용자 지정(xrefstream)', () => {
  assert.deepEqual([PdfDocument.open(load('classic')).pages()[0].width.toFixed(1), PdfDocument.open(load('classic')).pages()[0].height.toFixed(1)], ['595.3', '841.9']);
  assert.deepEqual([PdfDocument.open(load('xrefstream')).pages()[0].width, PdfDocument.open(load('xrefstream')).pages()[0].height], [300, 200]);
});

test('메모를 넣지 않으면 원본 바이트 그대로', () => {
  const o = load('classic');
  assert.equal(PdfDocument.open(o).withNotes([]), o);
});

for (const name of ['classic', 'links', 'xrefstream']) {
  test(`${name}: 증분 갱신으로 메모 추가 — 원본 바이트 보존, pypdf 읽기, 재열기`, (t) => {
    const o = load(name);
    const out = PdfDocument.open(o).withNotes([
      { page: 0, x: 100, y: 700, contents: '첫 메모 & <확인> 😀', author: '시험' },
      { page: 0, x: 200, y: 600, contents: '둘째 메모' },
    ]);
    assert.ok(out.length > o.length);
    assert.deepEqual(out.subarray(0, o.length), o, '원본 바이트가 앞부분에 그대로 남는다');
    const again = PdfDocument.open(out);
    const notes = again.notes();
    assert.deepEqual(notes.map(n => n.contents), ['첫 메모 & <확인> 😀', '둘째 메모']);
    assert.equal(notes[0].author, '시험');
    assert.equal(again.pages().length, PdfDocument.open(o).pages().length);
    const r = py(out), r0 = py(o);
    if (!r || !r0) return t.skip('python 없음');
    assert.equal(r.pages, r0.pages);
    assert.deepEqual(r.notes.map(n => n.contents), ['첫 메모 & <확인> 😀', '둘째 메모']);
    assert.equal(r.notes[0].subtype, '/Text');
    assert.deepEqual(r.text, r0.text, '쪽 글자는 그대로');
    assert.equal(r.links, r0.links, '기존 링크 주석은 그대로');
  });
}

test('이미 /Annots가 있는 쪽: 기존 주석을 지우지 않고 뒤에 붙인다', (t) => {
  const out = PdfDocument.open(load('links')).withNotes([{ page: 0, x: 50, y: 500, contents: '링크 쪽 메모' }]);
  const r = py(out);
  if (!r) return t.skip('python 없음');
  assert.equal(r.links, 1);
  assert.equal(r.notes.length, 1);
});

test('두 번 연속 증분 갱신해도 모두 남는다(Prev 연결)', (t) => {
  const once = PdfDocument.open(load('classic')).withNotes([{ page: 1, x: 10, y: 100, contents: '1회차' }]);
  const twice = PdfDocument.open(once).withNotes([{ page: 1, x: 10, y: 200, contents: '2회차' }, { page: 2, x: 10, y: 200, contents: '3쪽 메모' }]);
  assert.deepEqual(twice.subarray(0, once.length), once);
  assert.deepEqual(PdfDocument.open(twice).notes().map(n => [n.page, n.contents]), [[1, '1회차'], [1, '2회차'], [2, '3쪽 메모']]);
  const r = py(twice);
  if (r) assert.equal(r.notes.length, 3);
});

test('잘못된 요청과 입력은 거부한다', () => {
  const d = PdfDocument.open(load('classic'));
  assert.throws(() => d.withNotes([{ page: 9, x: 1, y: 1, contents: 'x' }]), PdfError);
  assert.throws(() => d.withNotes([{ page: -1, x: 1, y: 1, contents: 'x' }]), PdfError);
  assert.throws(() => d.withNotes([{ page: 0, x: NaN, y: 1, contents: 'x' }]), PdfError);
  assert.throws(() => PdfDocument.open(new Uint8Array([1, 2, 3])), PdfError);
  assert.throws(() => PdfDocument.open(new TextEncoder().encode('%PDF-1.4\n' + 'x'.repeat(100))), PdfError, 'startxref 없음');
  const o = load('classic');
  assert.throws(() => PdfDocument.open(o.subarray(0, o.length - 30)), PdfError, '끝이 잘린 파일');
  const enc = new TextEncoder().encode(new TextDecoder('latin1').decode(o).replace('/Root', '/Encrypt 99 0 R /Root'));
  assert.throws(() => PdfDocument.open(enc), PdfError, '암호화 표시');
});

test('형식 판별은 PDF로 하고, 지원 표시는 열람·메모', () => {
  assert.equal(detectFormat(load('classic')).id, 'pdf');
});

// 독립 검증 2: 실제 PDF 뷰어 엔진(pdf.js)이 우리가 덧붙인 메모를 주석으로 인식하는가(쪽 그리기까지)
for (const name of ['classic', 'xrefstream']) {
  test(`${name}: pdf.js가 증분 갱신본의 메모를 주석으로 읽고 쪽을 연다`, async () => {
    const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
    const out = PdfDocument.open(load(name)).withNotes([{ page: 0, x: 100, y: 150, contents: '뷰어 확인 메모' }]);
    const doc = await pdfjs.getDocument({ data: out.slice(), isEvalSupported: false, verbosity: 0 }).promise;
    assert.equal(doc.numPages, PdfDocument.open(load(name)).pages().length);
    const page = await doc.getPage(1);
    const annots = await page.getAnnotations();
    const note = annots.find((a: { subtype: string }) => a.subtype === 'Text') as { contentsObj?: { str: string }; rect: number[] } | undefined;
    assert.ok(note, 'Text 주석이 보임');
    assert.equal(note!.contentsObj?.str, '뷰어 확인 메모');
    assert.deepEqual([note!.rect[0], note!.rect[3]].map(Math.round), [100, 150], '왼쪽 위 기준점은 그대로(pdf.js는 아이콘 크기를 22pt로 맞춘다)');
    await doc.destroy();
  });
}
