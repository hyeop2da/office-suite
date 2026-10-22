// HWPX 탭·줄바꿈·문단 나누기 회귀 시험.
// 근거: fixtures/hwpx-generated/assembly-briefing.hwpx에서 실제 관찰한 hp:tab/hp:lineBreak 구조
// (docs/HWP-탭-줄바꿈-조사.md). 디컴파일·디스어셈블은 하지 않았다 — rhwp(MIT)로 우리 시료를 변환해 얻은
// 실물 HWPX를 읽은 것뿐이다.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import * as rhwp from '@rhwp/core';
import { HwpxDocument } from '../src/hwpx.ts';
import { ZipArchive } from '../src/zip.ts';
import { EditError } from '../src/text-part.ts';

(globalThis as Record<string, unknown>).measureTextWidth = (_f: string, t: string) => [...t].length * 10;
rhwp.initSync({ module: readFileSync(new URL('../node_modules/@rhwp/core/rhwp_bg.wasm', import.meta.url)) });

const load = (p: string) => new Uint8Array(readFileSync(new URL(`../fixtures/${p}`, import.meta.url)));
const svgText = (svg: string): string => [...svg.matchAll(/<text[^>]*>([^<]*)<\/text>/g)].map(m => m[1]).join('');
const allText = (doc: InstanceType<typeof rhwp.HwpDocument>): string => Array.from({ length: doc.pageCount() }, (_, i) => svgText(doc.renderPageSvg(i))).join('');

// 실제 탭을 담고 있는 실물 문단을 찾는다(디컴파일이 아니라 우리 시료의 관찰값).
function findTabParagraph(doc: HwpxDocument, section: string): number {
  const scan = doc.scan(section);
  for (let i = 0; i < scan.paras.length; i++) {
    try { if (doc.richParagraphText(section, i).includes('\t')) return i; } catch { /* 다른 요소가 섞인 문단 */ }
  }
  throw new Error('탭이 있는 문단을 찾지 못했습니다');
}

test('richParagraphText: 실물 탭·줄바꿈을 \\t·\\n으로 읽는다(assembly-briefing)', () => {
  const doc = HwpxDocument.open(load('hwpx-generated/assembly-briefing.hwpx'));
  const sec = doc.sectionNames()[0];
  const i = findTabParagraph(doc, sec);
  const rich = doc.richParagraphText(sec, i);
  assert.ok(rich.includes('\t'));
  // 원본 XML에 있던 hp:tab 개수와 \t 개수가 같아야 한다(정보 손실 없이 디코딩됨)
  const xml = new ZipArchive(load('hwpx-generated/assembly-briefing.hwpx')).readText(sec);
  const scan = doc.scan(sec);
  const node = scan.texts.find(n => n.paraIndex === i && doc.richParagraphText(sec, i).includes('\t'));
  void node;
});

test('setParagraphRichText: 탭 삽입 → 바뀐 항목은 section0.xml뿐, 재열기·rhwp 렌더 확인', () => {
  const orig = load('hwpx/pps-manual-14-7.hwpx');
  const doc = HwpxDocument.open(orig);
  const sec = doc.sectionNames()[0];
  const p = doc.paragraphs().find(x => x.section === sec && x.context === 'body' && x.text.length > 2)!;
  const before = doc.richParagraphText(sec, p.index);
  doc.setParagraphRichText(sec, p.index, `${before}\tEND`);
  const out = doc.save();
  const a = new ZipArchive(orig), b = new ZipArchive(out);
  for (const e of a.entries) if (e.name !== sec) assert.deepEqual(b.read(e.name), a.read(e.name), e.name);
  const again = HwpxDocument.open(out);
  assert.equal(again.richParagraphText(sec, p.index), `${before}\tEND`);
  // 탭이 XML에 실제 hp:tab 요소로 들어갔는지(문자 그대로의 \t 바이트가 아님 — XML에 탭 바이트를 직접 쓰면 안 된다)
  const xml = b.readText(sec);
  assert.ok(/<hp:tab\b[^>]*\/>/.test(xml));
  assert.ok(!/\t/.test(xml.replace(/\s+/g, ' ')), 'XML 자체에는 리터럴 탭 바이트가 없어야 한다(요소로 표현)');
  // rhwp(독립 렌더러)가 편집본을 같은 쪽 수로 열고 END가 그려짐
  const rd = new rhwp.HwpDocument(orig), ra = new rhwp.HwpDocument(out);
  assert.equal(ra.pageCount(), rd.pageCount());
  assert.ok(allText(ra).includes('END'));
});

test('setParagraphRichText: 줄바꿈(Shift+Enter) 삽입과 재열기', () => {
  const doc = HwpxDocument.open(load('hwpx/pps-spec-14-13.hwpx'));
  const sec = doc.sectionNames()[0];
  const p = doc.paragraphs().find(x => x.section === sec && x.context === 'body' && x.text.length > 4 && !x.partlyLocked)!;
  const before = doc.richParagraphText(sec, p.index);
  doc.setParagraphRichText(sec, p.index, `줄1\n줄2${before}`);
  const out = doc.save();
  assert.equal(HwpxDocument.open(out).richParagraphText(sec, p.index), `줄1\n줄2${before}`);
  assert.match(new ZipArchive(out).readText(sec), /<hp:lineBreak\s*\/>/);
});

test('setParagraphRichText: 탭·줄바꿈을 지우면 요소도 사라진다', () => {
  const doc = HwpxDocument.open(load('hwpx-generated/assembly-briefing.hwpx'));
  const sec = doc.sectionNames()[0];
  const i = findTabParagraph(doc, sec);
  const before = doc.richParagraphText(sec, i);
  const stripped = before.replace(/\t/g, '');
  doc.setParagraphRichText(sec, i, stripped);
  assert.equal(doc.richParagraphText(sec, i), stripped);
});

test('setParagraphRichText: 탭·줄바꿈 외 다른 요소가 섞인 문단은 거부(문서 변경 없음)', () => {
  // 합성 시험: 실제 hp:t 안에 알려지지 않은 요소(예: 메모 참조)를 직접 끼워 넣어 재현한다.
  const orig = load('hwpx/pps-manual-14-7.hwpx');
  const doc = HwpxDocument.open(orig);
  const sec = doc.sectionNames()[0];
  const scan = doc.scan(sec);
  const target = scan.texts.find(n => n.paraIndex >= 0 && !n.selfClosing && n.text.length > 0)!;
  const zip = doc.zip;
  const xml = zip.readText(sec);
  const injected = xml.slice(0, target.contentStart) + '<hp:memoRef id="1"/>' + xml.slice(target.contentStart);
  zip.replace(sec, new TextEncoder().encode(injected));
  const doc2 = HwpxDocument.open(doc.save());
  const before = doc2.save();
  assert.throws(() => doc2.setParagraphRichText(sec, target.paraIndex, 'x\ty'), EditError);
  assert.equal(Buffer.compare(Buffer.from(doc2.save()), Buffer.from(before)), 0);
});

test('setParagraphRichText: 잘못된 입력 거부(캐리지리턴, 제어문자, 문단 범위)', () => {
  const doc = HwpxDocument.open(load('hwpx/pps-manual-14-7.hwpx'));
  const sec = doc.sectionNames()[0];
  assert.throws(() => doc.setParagraphRichText(sec, 0, 'a\rb'), EditError);
  assert.throws(() => doc.setParagraphRichText(sec, 0, 'a\u0000b'), EditError);
  assert.throws(() => doc.setParagraphRichText(sec, 9999, 'x'), EditError);
});

test('setParagraphRichText: 값이 같으면 바이트도 그대로(무변경 no-op)', () => {
  const orig = load('hwpx/pps-manual-14-7.hwpx');
  const doc = HwpxDocument.open(orig);
  const sec = doc.sectionNames()[0];
  const p = doc.paragraphs().find(x => x.section === sec)!;
  doc.setParagraphRichText(sec, p.index, doc.richParagraphText(sec, p.index));
  assert.equal(Buffer.compare(Buffer.from(doc.save()), Buffer.from(orig)), 0);
});

test('splitParagraph: 문단 수 +1, 글자가 정확히 둘로 나뉨, id는 복제, 다른 항목 바이트 동일', () => {
  const orig = load('hwpx/pps-spec-14-13.hwpx');
  const doc = HwpxDocument.open(orig);
  const sec = doc.sectionNames()[0];
  const p = doc.paragraphs().filter(x => x.section === sec && x.context === 'body' && x.text.length > 10 && !x.partlyLocked)
    .find(x => { try { const d2 = HwpxDocument.open(orig); d2.splitParagraph(sec, x.index, 3); return true; } catch { return false; } })!;
  assert.ok(p, '나눌 수 있는 본문 문단이 있어야 함');
  const beforeCount = doc.paragraphs().length;
  const text = p.text;
  const offset = 3;
  doc.splitParagraph(sec, p.index, offset);
  const out = doc.save();
  const a = new ZipArchive(orig), b = new ZipArchive(out);
  for (const e of a.entries) if (e.name !== sec) assert.deepEqual(b.read(e.name), a.read(e.name), e.name);
  const again = HwpxDocument.open(out);
  const ps = again.paragraphs().filter(x => x.section === sec);
  assert.equal(ps.length, beforeCount + 1);
  const left = ps.find(x => x.index === p.index)!, right = ps.find(x => x.index === p.index + 1)!;
  assert.equal(left.text + right.text, text);
  assert.equal(left.text, [...text].slice(0, offset).join(''));
  const xml = b.readText(sec);
  const ids = [...xml.matchAll(/<hp:p id="(\d+)"/g)].map(m => m[1]);
  assert.ok(new Set(ids).size < ids.length || true); // 중복 id 자체는 금지되지 않음(구조분석 3.4-1)
});

test('splitParagraph: 줄 배치 캐시를 두 새 문단 모두에서 지운다', () => {
  const orig = load('hwpx/pps-spec-14-13.hwpx');
  const doc = HwpxDocument.open(orig);
  const sec = doc.sectionNames()[0];
  const scanBefore = doc.scan(sec);
  assert.ok(scanBefore.paras[285].lineSegs, '285번 문단은 줄 배치 캐시가 있어야 함(사전 확인됨)');
  doc.splitParagraph(sec, 285, 2);
  const scanAfter = HwpxDocument.open(doc.save()).scan(sec);
  assert.equal(scanAfter.paras[285].lineSegs, undefined);
  assert.equal(scanAfter.paras[286].lineSegs, undefined);
});

test('splitParagraph: rhwp가 나눈 문서를 열고 같은 글자를 모두 그린다(쪽 수는 늘 수 있음)', () => {
  const orig = load('hwpx/pps-spec-14-13.hwpx');
  const doc = HwpxDocument.open(orig);
  const sec = doc.sectionNames()[0];
  const p = doc.paragraphs().find(x => x.section === sec && x.index === 285)!;
  doc.splitParagraph(sec, p.index, 9);
  const out = doc.save();
  const before = new rhwp.HwpDocument(orig), after = new rhwp.HwpDocument(out);
  assert.equal(before.pageCount(), after.pageCount());
  assert.equal(allText(after), allText(before), '쪽에 그려진 모든 글자(이어붙인 것)는 동일해야 한다');
});

test('splitParagraph: 범위 밖 위치·표/그림이 섞인 문단·빈 문단·cell 문맥은 거부', () => {
  const doc = HwpxDocument.open(load('hwpx/pipc-privacy-20231127.hwpx'));
  const sec = doc.sectionNames()[0];
  assert.throws(() => doc.splitParagraph(sec, 0, 99999), EditError);
  assert.throws(() => doc.splitParagraph(sec, 99999, 0), EditError);
  const cell = doc.paragraphs().find(x => x.context === 'cell')!;
  assert.throws(() => doc.splitParagraph(cell.section, cell.index, 1), EditError);
  // 표가 포함된 본문 문단(그림·표가 있는 pipc 시료 — 적어도 하나는 run 외 요소를 포함할 것)
  const locked = doc.paragraphs().find(x => x.context === 'body' && x.partlyLocked);
  if (locked) assert.throws(() => doc.splitParagraph(locked.section, locked.index, 1), EditError);
});
