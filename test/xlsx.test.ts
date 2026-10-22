import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, writeFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { XlsxDocument } from '../src/xlsx.ts';
import { SheetSession } from '../src/sheet-session.ts';
import { ZipArchive } from '../src/zip.ts';
import { EditError } from '../src/text-part.ts';
import { FormatError } from '../src/xml-scan.ts';

const load = (n: string) => new Uint8Array(readFileSync(new URL(`../fixtures/xlsx/${n}.xlsx`, import.meta.url)));

interface PyCell { v: unknown; type: string; fmt: string }
/** 독립 검증: openpyxl이 읽은 첫 시트 셀의 값·파이썬 형·서식, 시트 목록과 상태, 수식 */
function py(bytes: Uint8Array): { cells: Record<string, PyCell>; formulas: Record<string, unknown>; sheets: [string, string][]; calc: boolean | null; second: Record<string, unknown> } | null {
  const p = join(mkdtempSync(join(tmpdir(), 'xlsx-')), 'x.xlsx');
  writeFileSync(p, bytes);
  const r = spawnSync('python', ['-I', '-X', 'utf8', '-c', `
import sys, json, datetime, openpyxl
f = sys.argv[1]
wb = openpyxl.load_workbook(f)
wv = openpyxl.load_workbook(f, data_only=True)
ws = wb.worksheets[0]
def enc(v):
    if isinstance(v, (datetime.datetime, datetime.date)): return v.isoformat()
    return v
cells = {}
formulas = {}
for row in ws.iter_rows():
    for c in row:
        if c.value is None: continue
        if isinstance(c.value, str) and c.value.startswith('='): formulas[c.coordinate] = c.value
        cells[c.coordinate] = {'v': enc(c.value), 'type': type(c.value).__name__, 'fmt': c.number_format}
second = {c.coordinate: enc(c.value) for row in wb.worksheets[1].iter_rows() for c in row if c.value is not None} if len(wb.worksheets) > 1 else {}
print(json.dumps({'cells': cells, 'formulas': formulas, 'sheets': [[s.title, s.sheet_state] for s in wb.worksheets],
                  'calc': wb.calculation.fullCalcOnLoad, 'second': second}, ensure_ascii=False))`, p], { encoding: 'utf8' });
  if (r.error) return null;
  assert.equal(r.status, 0, r.stderr);
  return JSON.parse(r.stdout);
}

test('시트 목록: 이름(특수문자 포함), 숨김 상태', () => {
  const d = XlsxDocument.open(load('types'));
  assert.deepEqual(d.sheets().map(s => [s.name, s.state]), [['Sheet1', 'visible'], ['두 번째 & 시트', 'visible'], ['숨김', 'hidden']]);
});

test('셀 읽기: 종류와 원래 표기', () => {
  const d = XlsxDocument.open(load('types'));
  const c = (r: string) => d.cell(0, r);
  assert.deepEqual([c('B1').kind, c('B1').text], ['number', '42']);
  assert.equal(c('B2').text, '3.14159');
  assert.equal(c('B4').text, '1.5e-07', '지수 표기는 원문 그대로');
  assert.equal(c('B5').text, '12345678901');
  assert.equal(c('B6').text, '0.25');
  assert.deepEqual([c('B7').kind, c('B7').display], ['number', '2026-10-22'], '날짜 서식 숫자는 ISO 표시를 함께 준다');
  assert.deepEqual([c('B8').kind, c('B8').text], ['bool', 'TRUE']);
  assert.deepEqual([c('B9').kind, c('B9').formula, c('B9').editable], ['formula', 'SUM(B1:B3)', false]);
  assert.equal(c('A10').text, '특수 & <문자> "따옴표"');
  assert.equal(c('B10').text, '  앞뒤 공백  ');
  assert.equal(c('A12').text, '줄\n바꿈');
  assert.equal(c('C3').kind, 'empty');
  assert.deepEqual(d.usedRange(0), { rows: 12, cols: 4 });
});

for (const name of ['types', 'types-inline', 'rich']) {
  test(`${name}: 무수정 저장은 파일 전체 바이트가 같다`, () => {
    const o = load(name);
    assert.equal(Buffer.compare(Buffer.from(XlsxDocument.open(o).save()), Buffer.from(o)), 0);
  });
}

test('숫자 칸은 숫자로 남는다(문자열로 바뀌지 않는다)', (t) => {
  const d = XlsxDocument.open(load('types'));
  d.setCell(0, 'B1', '43');
  d.setCell(0, 'B2', '2.5');
  d.setCell(0, 'B3', '-1e3');
  d.setCell(0, 'B5', ' 7 ');
  const out = d.save();
  const xml = new ZipArchive(out).readText('xl/worksheets/sheet1.xml');
  assert.match(xml, /<c r="B1"><v>43<\/v><\/c>/);
  assert.doesNotMatch(xml, /<c r="B[1235]"[^>]*t="(s|inlineStr|str)"/, '숫자 칸에 글자 형식이 붙지 않음');
  const r = py(out);
  if (!r) return t.skip('python 없음');
  assert.deepEqual([r.cells.B1.v, r.cells.B1.type], [43, 'int']);
  assert.deepEqual([r.cells.B2.v, r.cells.B2.type], [2.5, 'float']);
  assert.equal(r.cells.B3.v, -1000);
  assert.equal(r.cells.B5.v, 7);
  assert.equal(r.cells.B4.v, 1.5e-7, '건드리지 않은 숫자 칸은 그대로');
});

test('숫자 칸에 숫자가 아닌 입력은 거부하고 문서를 바꾸지 않는다', () => {
  const d = XlsxDocument.open(load('types'));
  const before = d.save();
  for (const bad of ['abc', '1,234', '12abc', '1.2.3', 'NaN', 'Infinity', '0x10', '１２']) {
    assert.throws(() => d.setCell(0, 'B1', bad), EditError, bad);
  }
  assert.equal(Buffer.compare(Buffer.from(d.save()), Buffer.from(before)), 0);
});

test('숫자 칸을 비우면 값만 지우고 서식(s)은 남는다', (t) => {
  const d = XlsxDocument.open(load('types'));
  d.setCell(0, 'B6', '');
  const xml = new ZipArchive(d.save()).readText('xl/worksheets/sheet1.xml');
  assert.match(xml, /<c r="B6" s="1"\/>/);
  const r = py(d.save());
  if (r) assert.ok(!('B6' in r.cells));
});

test('공유 문자열: 같은 문자열을 쓰는 다른 칸과 공유 표는 그대로', (t) => {
  const o = load('types');
  const d = XlsxDocument.open(o);
  assert.equal(d.cell(0, 'A11').text, d.cell(0, 'B11').text);
  d.setCell(0, 'A11', '바뀐말');
  const out = d.save();
  const a = new ZipArchive(o), b = new ZipArchive(out);
  assert.deepEqual(b.read('xl/sharedStrings.xml'), a.read('xl/sharedStrings.xml'), 'sharedStrings.xml 바이트 동일');
  for (const e of a.entries) if (e.name !== 'xl/worksheets/sheet1.xml') assert.deepEqual(b.read(e.name), a.read(e.name), e.name);
  const r = py(out);
  if (!r) return t.skip('python 없음');
  assert.equal(r.cells.A11.v, '바뀐말');
  assert.equal(r.cells.B11.v, '같은말');
});

test('글자 칸: 특수문자·앞뒤 공백·줄바꿈·이모지 왕복', (t) => {
  const d = XlsxDocument.open(load('types'));
  const text = ' A&B <c> "q" \'a\' 😀\n두 줄  ';
  d.setCell(0, 'A10', text);
  const out = d.save();
  assert.equal(XlsxDocument.open(out).cell(0, 'A10').text, text);
  assert.match(new ZipArchive(out).readText('xl/worksheets/sheet1.xml'), /<t xml:space="preserve"> A&amp;B &lt;c&gt;/);
  const r = py(out);
  if (r) assert.equal(r.cells.A10.v, text);
});

test('날짜 서식 숫자: ISO 날짜를 넣으면 일련번호 숫자로 저장되고 서식은 유지된다', (t) => {
  const d = XlsxDocument.open(load('types'));
  d.setCell(0, 'B7', '2027-01-01');
  const out = d.save();
  const c = XlsxDocument.open(out).cell(0, 'B7');
  assert.deepEqual([c.kind, c.display], ['number', '2027-01-01']);
  const r = py(out);
  if (!r) return t.skip('python 없음');
  assert.equal(r.cells.B7.fmt, 'yyyy-mm-dd');
  assert.match(String(r.cells.B7.v), /^2027-01-01/);
  assert.throws(() => d.setCell(0, 'B7', '2027-02-30'), EditError, '없는 날짜');
});

test('논리값 칸: TRUE/FALSE/0/1만 허용', (t) => {
  const d = XlsxDocument.open(load('types'));
  d.setCell(0, 'B8', 'false');
  assert.throws(() => d.setCell(0, 'B8', '네'), EditError);
  const r = py(d.save());
  if (r) assert.equal(r.cells.B8.v, false);
  assert.equal(d.cell(0, 'B8').text, 'FALSE');
});

test('수식 칸은 편집 불가, 다른 칸을 고치면 열 때 다시 계산하도록 표시', (t) => {
  const d = XlsxDocument.open(load('types'));
  assert.throws(() => d.setCell(0, 'B9', '5'), EditError);
  d.setCell(0, 'B1', '100');
  const r = py(d.save());
  if (!r) return t.skip('python 없음');
  assert.equal(r.formulas.B9, '=SUM(B1:B3)', '수식 보존');
  assert.equal(r.calc, true, 'fullCalcOnLoad');
});

test('빈 칸에 쓰기: 같은 행의 순서 위치, 새 행, 범위(dimension) 갱신', (t) => {
  const d = XlsxDocument.open(load('types'));
  d.setCell(0, 'C1', '새 글자');      // 기존 행 중간(B1 뒤, 이후 칸 없음)
  d.setCell(0, 'C5', '7');             // D5가 있는 행: D5 앞에 들어가야 한다
  d.setCell(0, 'A20', 'x');            // 새 행
  d.setCell(0, 'F2', '1.5');           // 범위 확장(열)
  const out = d.save();
  const xml = new ZipArchive(out).readText('xl/worksheets/sheet1.xml');
  assert.ok(xml.indexOf('r="C5"') < xml.indexOf('r="D5"'), '열 순서');
  assert.match(xml, /<dimension ref="A1:F20"\/>/);
  const rows = [...xml.matchAll(/<row r="(\d+)"/g)].map(m => Number(m[1]));
  assert.deepEqual(rows, [...rows].sort((a, b) => a - b), '행 순서');
  const again = XlsxDocument.open(out);
  assert.equal(again.cell(0, 'C1').text, '새 글자');
  assert.deepEqual([again.cell(0, 'C5').kind, again.cell(0, 'C5').text], ['number', '7'], '빈 칸에 숫자처럼 보이는 입력은 숫자');
  assert.equal(again.cell(0, 'A20').text, 'x');
  const r = py(out);
  if (r) { assert.equal(r.cells.C5.v, 7); assert.equal(r.cells.D5.v, '떨어진 칸'); assert.equal(r.cells.F2.v, 1.5); }
});

test('다른 시트는 바이트 그대로이고, 시트 이름·상태가 유지된다', (t) => {
  const o = load('types');
  const d = XlsxDocument.open(o);
  d.setCell(1, 'B1', '200');
  const out = d.save();
  const a = new ZipArchive(o), b = new ZipArchive(out);
  for (const e of a.entries) if (e.name !== 'xl/worksheets/sheet2.xml' && e.name !== 'xl/workbook.xml') assert.deepEqual(b.read(e.name), a.read(e.name), e.name);
  const r = py(out);
  if (!r) return t.skip('python 없음');
  assert.deepEqual(r.sheets, [['Sheet1', 'visible'], ['두 번째 & 시트', 'visible'], ['숨김', 'hidden']]);
  assert.equal(r.second.B1, 200);
});

test('서식 있는 공유 문자열: 읽을 때 rich 표시, 편집하면 일반 글자가 된다', () => {
  const d = XlsxDocument.open(load('rich'));
  const c = d.cell(0, 'A1');
  assert.deepEqual([c.text, c.rich], ['굵게 기울임 보통', true]);
  d.setCell(0, 'A1', '일반으로');
  assert.equal(d.cell(0, 'A1').rich, false);
  assert.equal(d.cell(0, 'A2').text, '일반 문자열');
});

test('inlineStr 칸도 읽고 편집한다', () => {
  const d = XlsxDocument.open(load('types-inline'));
  assert.equal(d.cell(0, 'A1').text, '정수');
  d.setCell(0, 'A1', '바꿈');
  assert.equal(XlsxDocument.open(d.save()).cell(0, 'A1').text, '바꿈');
});

test('잘못된 요청 거부: 셀 주소, 시트 번호, 제어 문자, 너무 긴 글자', () => {
  const d = XlsxDocument.open(load('types'));
  const before = d.save();
  assert.throws(() => d.setCell(0, 'a1', 'x'), EditError);
  assert.throws(() => d.setCell(0, 'A0', 'x'), EditError);
  assert.throws(() => d.setCell(0, 'XFE1', 'x'), EditError);
  assert.throws(() => d.setCell(0, 'A1048577', 'x'), EditError);
  assert.throws(() => d.setCell(9, 'A1', 'x'), EditError);
  assert.throws(() => d.setCell(0, 'A1', 'a\u0000b'), EditError);
  assert.throws(() => d.setCell(0, 'A1', 'x'.repeat(32768)), EditError);
  assert.equal(Buffer.compare(Buffer.from(d.save()), Buffer.from(before)), 0);
});

test('XLSX가 아닌 파일은 거부', () => {
  assert.throws(() => XlsxDocument.open(new Uint8Array(readFileSync(new URL('../fixtures/docx/basic.docx', import.meta.url)))), FormatError);
});

test('SheetSession: 수정 목록 → 저장 → 재열기 검증, 되돌리기, 잘못된 입력은 목록에 들어가지 않는다', () => {
  const o = load('types');
  const s = new SheetSession(o, 'x.xlsx');
  s.setText('0!B1', '43');
  s.setText('0!A11', '바뀐말');
  s.setText('0!C3', '9');
  assert.equal(s.changeCount, 3);
  assert.throws(() => s.setText('0!B2', 'abc'), EditError);
  assert.equal(s.changeCount, 3, '거부된 입력은 목록에 없음');
  s.setText('0!B1', '42'); // 원래 값으로 돌아가면 목록에서 빠진다
  assert.equal(s.changeCount, 2);
  const r = s.save();
  assert.deepEqual(r.problems, []);
  assert.ok(r.verified);
  const again = XlsxDocument.open(r.bytes);
  assert.equal(again.cell(0, 'A11').text, '바뀐말');
  assert.equal(again.cell(0, 'B11').text, '같은말');
  assert.deepEqual([again.cell(0, 'C3').kind, again.cell(0, 'C3').text], ['number', '9']);
  s.revertAll();
  assert.equal(Buffer.compare(Buffer.from(s.save().bytes), Buffer.from(o)), 0, '수정이 없으면 원본과 같다');
});

test('calcPr가 없는 통합문서: 수식이 있으면 스키마 순서에 맞게 fullCalcOnLoad를 넣는다', () => {
  const z = new ZipArchive(load('types'));
  const wb = z.readText('xl/workbook.xml').replace(/<calcPr[^>]*\/>/, '');
  assert.ok(!wb.includes('calcPr'));
  z.replace('xl/workbook.xml', new TextEncoder().encode(wb));
  const d = XlsxDocument.open(z.write());
  d.setCell(0, 'B1', '1');
  const out = new ZipArchive(d.save()).readText('xl/workbook.xml');
  assert.match(out, /<definedNames\/><calcPr fullCalcOnLoad="1"\/><\/workbook>/);
  // fullCalcOnLoad="0"인 경우는 1로 바꾼다
  const z2 = new ZipArchive(load('types'));
  z2.replace('xl/workbook.xml', new TextEncoder().encode(z2.readText('xl/workbook.xml').replace('fullCalcOnLoad="1"', 'fullCalcOnLoad="0"')));
  const d2 = XlsxDocument.open(z2.write());
  d2.setCell(0, 'B1', '1');
  assert.match(new ZipArchive(d2.save()).readText('xl/workbook.xml'), /fullCalcOnLoad="1"/);
});

// 병합 셀 (2026-10-22 추가): 실제 시료에 병합 셀이 없어, 공개 규격(OOXML mergeCells/mergeCell)대로
// 합성 주입해 재현한다. 실행 파일 분석은 하지 않았다 — ECMA-376 공개 스키마만으로 충분했다.
function withMerge(base: string, ref: string): Uint8Array {
  const z = new ZipArchive(load(base));
  const xml = z.readText('xl/worksheets/sheet1.xml');
  const injected = xml.replace('</sheetData>', `</sheetData><mergeCells count="1"><mergeCell ref="${ref}"/></mergeCells>`);
  z.replace('xl/worksheets/sheet1.xml', new TextEncoder().encode(injected));
  return z.write();
}

test('mergedRanges: 주입한 mergeCell을 그대로 읽는다', () => {
  const d = XlsxDocument.open(withMerge('types', 'F1:G2'));
  assert.deepEqual(d.mergedRanges(0), ['F1:G2']);
});

test('cell: 병합 범위 안 칸은 merged 정보를 갖고, 왼쪽 위만 anchor=true', () => {
  const d = XlsxDocument.open(withMerge('types', 'F1:G2'));
  assert.deepEqual(d.cell(0, 'F1').merged, { ref: 'F1:G2', anchor: true });
  assert.deepEqual(d.cell(0, 'G1').merged, { ref: 'F1:G2', anchor: false });
  assert.deepEqual(d.cell(0, 'F2').merged, { ref: 'F1:G2', anchor: false });
  assert.deepEqual(d.cell(0, 'G2').merged, { ref: 'F1:G2', anchor: false });
  assert.equal(d.cell(0, 'H1').merged, undefined, '범위 밖 칸은 merged가 없다');
});

test('setCell: 병합 범위의 중심(왼쪽 위) 칸만 편집할 수 있고, 다른 칸은 거부(문서 변경 없음)', (t) => {
  const orig = withMerge('types', 'F1:G2');
  const d = XlsxDocument.open(orig);
  d.setCell(0, 'F1', '합친 제목');
  const out = d.save();
  assert.equal(new ZipArchive(out).readText('xl/worksheets/sheet1.xml').includes('mergeCell ref="F1:G2"'), true, 'mergeCells 요소는 그대로 남는다');
  const again = XlsxDocument.open(out);
  assert.equal(again.cell(0, 'F1').text, '합친 제목');
  const r = py(out);
  if (r) {
    // openpyxl의 merged_cells 범위도 그대로인지 별도로 확인
    const p = join(mkdtempSync(join(tmpdir(), 'xlsx-')), 'x.xlsx');
    writeFileSync(p, out);
    const rr = spawnSync('python', ['-I', '-X', 'utf8', '-c', `
import sys, json, openpyxl
wb = openpyxl.load_workbook(sys.argv[1])
ws = wb.worksheets[0]
print(json.dumps({'merged': [str(x) for x in ws.merged_cells.ranges], 'f1': ws['F1'].value}, ensure_ascii=False))`, p], { encoding: 'utf8' });
    if (rr.status === 0) {
      const info = JSON.parse(rr.stdout);
      assert.deepEqual(info.merged, ['F1:G2']);
      assert.equal(info.f1, '합친 제목');
    }
  } else t.skip('python 없음');

  const d2 = XlsxDocument.open(orig);
  const before = d2.save();
  assert.throws(() => d2.setCell(0, 'G1', 'x'), EditError);
  assert.throws(() => d2.setCell(0, 'F2', 'x'), EditError);
  assert.throws(() => d2.setCell(0, 'G2', 'x'), EditError);
  assert.equal(Buffer.compare(Buffer.from(d2.save()), Buffer.from(before)), 0);
});
