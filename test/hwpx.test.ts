import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync, writeFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { DOMParser } from '@xmldom/xmldom';
import { HwpxDocument } from '../src/hwpx.ts';
import { ZipArchive } from '../src/zip.ts';
import { scanSection } from '../src/xml-scan.ts';

const dir = new URL('../fixtures/hwpx/', import.meta.url);
const fixtures = readdirSync(dir).filter(f => f.endsWith('.hwpx'));
const MARK = '〔편집시험〕';

function load(f: string): Uint8Array { return new Uint8Array(readFileSync(new URL(f, dir))); }
function meta(z: ZipArchive) { return z.entries.map(e => `${e.name}:${e.method}`); }
function wellFormed(xml: string, name: string) {
  const errors: string[] = [];
  new DOMParser({ errorHandler: { warning() {}, error: m => errors.push(String(m)), fatalError: m => errors.push(String(m)) } })
    .parseFromString(xml, 'text/xml');
  assert.deepEqual(errors, [], `${name} XML 오류`);
}

test('시료가 5개 있다', () => { assert.equal(fixtures.length, 5); });

for (const f of fixtures) {
  test(`${f}: 무수정 열기→저장은 모든 항목이 같다`, () => {
    const orig = load(f);
    const out = HwpxDocument.open(orig).save();
    const a = new ZipArchive(orig), b = new ZipArchive(out);
    assert.deepEqual(meta(b), meta(a), '항목 순서·압축 방식');
    for (const e of a.entries) assert.deepEqual(b.read(e.name), a.read(e.name), e.name);
    // 항목 바이트까지 같은지(압축 바이트 그대로 복사)
    assert.equal(Buffer.compare(Buffer.from(out), Buffer.from(orig)), 0, '파일 전체 바이트 동일');
  });

  test(`${f}: 글자 하나 넣어도(keep 정책) 바뀌는 곳은 마커뿐이다`, () => {
    const orig = load(f);
    const doc = HwpxDocument.open(orig);
    const sec = doc.sectionNames()[0];
    const before = new ZipArchive(orig).readText(sec);
    const beforeChars = doc.charCount(sec);
    assert.ok(doc.prependText(MARK, { layout: 'keep' }));
    const out = doc.save();

    const a = new ZipArchive(orig), b = new ZipArchive(out);
    assert.deepEqual(meta(b), meta(a));
    for (const e of a.entries) {
      if (e.name === sec) continue;
      assert.deepEqual(b.read(e.name), a.read(e.name), `${e.name}은 그대로여야 함`);
    }
    const after = b.readText(sec);
    assert.equal(after.replace(MARK, ''), before, '마커를 빼면 원본 XML과 글자 단위로 같다');
    wellFormed(after, sec);
    assert.equal(HwpxDocument.open(out).charCount(sec), beforeChars + [...MARK].length);
    // 문단 id, 누름틀 메타태그 등 손실 목록 항목이 그대로인지
    for (const pat of [/<hp:p [^>]*>/g, /metatag="[^"]*"/g, /<hp:metaTag>[^<]*<\/hp:metaTag>/g, /zorder="[^"]*"/g]) {
      assert.deepEqual(after.match(pat) ?? [], before.match(pat) ?? [], String(pat));
    }
  });

  test(`${f}: 저장본을 파이썬 zipfile이 검증한다`, (t) => {
    const doc = HwpxDocument.open(load(f));
    doc.prependText(MARK);
    const p = join(mkdtempSync(join(tmpdir(), 'hwpx-')), 'out.hwpx');
    writeFileSync(p, doc.save());
    const r = spawnSync('python', ['-I', '-X', 'utf8', '-c',
      'import sys,zipfile;z=zipfile.ZipFile(sys.argv[1]);bad=z.testzip();print(bad);sys.exit(1 if bad else 0)', p],
      { encoding: 'utf8' });
    if (r.error) return t.skip('python 없음');
    assert.equal(r.status, 0, r.stdout + r.stderr);
  });
}

test('줄 배치 캐시: drop은 바뀐 문단만 지운다', () => {
  const doc = HwpxDocument.open(load('pps-manual-14-7.hwpx'));
  const sec = doc.sectionNames()[0];
  const before = doc.scan(sec);
  const segsBefore = before.paras.filter(p => p.lineSegs).length;
  assert.ok(doc.prependText('X', { layout: 'drop' }));
  const after = HwpxDocument.open(doc.save()).scan(sec);
  assert.equal(after.paras.filter(p => p.lineSegs).length, segsBefore - 1);
});

test('replaceText: 엔티티와 빈 hp:t를 다룬다', () => {
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec>' +
    '<hp:p id="7"><hp:run><hp:t>A &amp; B</hp:t></hp:run><hp:run><hp:t/></hp:run>' +
    '<hp:run><hp:t>x<hp:tab/>y</hp:t></hp:run></hp:p></hs:sec>';
  const s = scanSection(xml);
  assert.equal(s.texts.length, 3);
  assert.equal(s.texts[0].text, 'A & B');
  assert.equal(s.texts[1].selfClosing, true);
  assert.equal(s.texts[2].hasChildren, true);
  assert.equal(s.texts[2].text, 'xy');
});

test('중첩 표 안 문단도 소속 문단을 올바르게 찾는다', () => {
  const xml = '<hs:sec><hp:p id="1"><hp:run><hp:tbl><hp:tr><hp:tc><hp:subList>' +
    '<hp:p id="2"><hp:run><hp:t>셀</hp:t></hp:run><hp:linesegarray><hp:lineseg a="1"/></hp:linesegarray></hp:p>' +
    '</hp:subList></hp:tc></hp:tr></hp:tbl></hp:run><hp:linesegarray><hp:lineseg a="2"/></hp:linesegarray></hp:p></hs:sec>';
  const s = scanSection(xml);
  assert.equal(s.paras.length, 2);
  assert.equal(s.texts[0].paraIndex, 1, '셀 글자는 안쪽 문단 소속');
  assert.ok(xml.slice(...s.paras[1].lineSegs!).includes('a="1"'));
  assert.ok(xml.slice(...s.paras[0].lineSegs!).includes('a="2"'));
});
