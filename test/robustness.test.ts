// 손상된 입력, 잘못된 편집 요청, 문단·표 셀·머리말·바이트 보존에 대한 회귀 시험.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { HwpxDocument } from '../src/hwpx.ts';
import { ZipArchive, ZipError } from '../src/zip.ts';
import { scanSection, FormatError } from '../src/xml-scan.ts';
import { EditError } from '../src/text-part.ts';
import { detectFormat, openDocument, UnsupportedFormatError } from '../src/engine.ts';

const fx = (p: string) => new Uint8Array(readFileSync(new URL(`../fixtures/${p}`, import.meta.url)));
const MANUAL = 'hwpx/pps-manual-14-7.hwpx';
const SPEC = 'hwpx/pps-spec-14-13.hwpx';

test('손상: 잘린 ZIP, 빈 입력, 쓰레기 바이트는 ZipError', () => {
  const good = fx(MANUAL);
  for (const bad of [good.slice(0, good.length - 30), good.slice(0, 10), new Uint8Array(0), new Uint8Array(100).fill(7)]) {
    assert.throws(() => HwpxDocument.open(bad), (e: unknown) => e instanceof ZipError);
  }
});

test('손상: 압축 데이터가 깨진 항목을 읽으면 ZipError', () => {
  const bytes = fx(MANUAL).slice();
  const zip = new ZipArchive(bytes);
  const e = zip.entries.find(x => x.name === 'Contents/section0.xml')!;
  // 로컬 헤더 뒤 데이터 한가운데를 망가뜨린다
  const off = bytes.length > 0 ? bytes.indexOf(e.raw[40]) : 0; void off;
  const broken = bytes.slice();
  const start = 30 + new DataView(e.raw.buffer, e.raw.byteOffset).getUint16(26, true);
  const pos = findSub(broken, e.raw.subarray(0, 30)) + start + 20;
  for (let i = 0; i < 40; i++) broken[pos + i] ^= 0xff;
  const doc = HwpxDocument.open(broken); // 열기는 되어도
  assert.throws(() => doc.paragraphs(), (err: unknown) => err instanceof ZipError, '읽을 때 거부');
});

function findSub(hay: Uint8Array, needle: Uint8Array): number {
  outer: for (let i = 0; i < hay.length - needle.length; i++) {
    for (let j = 0; j < needle.length; j++) if (hay[i + j] !== needle[j]) continue outer;
    return i;
  }
  return -1;
}

test('손상: XML이 깨진 구역은 FormatError', () => {
  assert.throws(() => scanSection('<hs:sec><hp:p><hp:t>x</hp:p></hs:sec>'), FormatError);
  assert.throws(() => scanSection('<hs:sec><hp:p>'), FormatError);
  assert.throws(() => scanSection('<hs:sec><!-- 닫히지 않음'), FormatError);
});

test('HWPX가 아닌 ZIP, mimetype 없음은 FormatError', () => {
  assert.throws(() => HwpxDocument.open(fx('docx/basic.docx')), FormatError);
});

test('잘못된 편집 요청은 문서를 바꾸지 않고 거부한다', () => {
  const doc = HwpxDocument.open(fx(MANUAL));
  const s = doc.sectionNames()[0];
  const before = doc.save();
  assert.throws(() => doc.setText(s, 99999, 'x'), EditError);
  assert.throws(() => doc.setText(s, -1, 'x'), EditError);
  assert.throws(() => doc.setText(s, 1.5, 'x'), EditError);
  assert.throws(() => doc.setText('Contents/section9.xml', 0, 'x'), EditError);
  assert.throws(() => doc.setParagraphText(s, 99999, 'x'), EditError);
  assert.throws(() => doc.replaceText('', 'x'), EditError);
  for (const bad of ['a\nb', 'a\rb', 'a\tb', 'a\u0000b', 'a\u0008b', '\ud800', 'x\udc00y']) {
    assert.throws(() => doc.setText(s, 0, bad), EditError, JSON.stringify(bad));
  }
  assert.equal(Buffer.compare(Buffer.from(doc.save()), Buffer.from(before)), 0, '거부된 요청은 아무것도 바꾸지 않는다');
});

test('특수문자·한글·공백 왕복: 엔티티 인코딩과 재열기', () => {
  const doc = HwpxDocument.open(fx(MANUAL));
  const s = doc.sectionNames()[0];
  const k = doc.scan(s).texts.findIndex(t => !t.hasChildren && t.text.length > 2);
  const text = ` A&B <tag> "q" 'a' 한글 😀 ]]> 끝  `;
  doc.setText(s, k, text);
  const again = HwpxDocument.open(doc.save());
  assert.equal(again.scan(s).texts[k].text, text);
  const raw = new ZipArchive(doc.save()).readText(s);
  assert.ok(raw.includes('A&amp;B &lt;tag&gt;'));
});

test('replaceText: 같은 노드의 여러 번 나온 글자를 모두 바꾸고 $ 패턴을 글자 그대로 둔다', () => {
  const xml = '<hs:sec><hp:p id="1"><hp:run><hp:t>aXbXc</hp:t></hp:run></hp:p></hs:sec>';
  assert.equal(scanSection(xml).texts[0].text, 'aXbXc');
  const doc = HwpxDocument.open(fx(MANUAL));
  const s = doc.sectionNames()[0];
  const nodes = doc.scan(s).texts.filter(t => !t.hasChildren);
  const target = nodes.find(t => t.text.length >= 3)!;
  const ch = target.text[1];
  const n = doc.replaceText(ch, '$&', { all: true });
  assert.ok(n >= 1);
  assert.ok(doc.plainText(s).includes('$&'));
});

test('문단 편집: 앞뒤 공통부분은 그대로, 중간만 바뀌고 문단 id는 보존', () => {
  const orig = fx(SPEC);
  const doc = HwpxDocument.open(orig);
  const p = doc.paragraphs().find(x => x.context === 'body' && [...x.text].length >= 6)!;
  const s = p.section;
  const before = new ZipArchive(orig).readText(s);
  const mid = [...p.text];
  mid.splice(2, 2, '가', '나', '다');
  doc.setParagraphText(s, p.index, mid.join(''));
  const out = HwpxDocument.open(doc.save());
  assert.equal(out.paragraphs().find(x => x.section === s && x.index === p.index)!.text, mid.join(''));
  const after = new ZipArchive(doc.save()).readText(s);
  const ids = (x: string) => x.match(/<hp:p [^>]*>/g);
  assert.deepEqual(ids(after), ids(before));
});

test('표 셀 글자 편집: rhwp 없이도 셀 문단이 context=cell로 보이고 편집된다', () => {
  const doc = HwpxDocument.open(fx(SPEC));
  const cells = doc.paragraphs().filter(p => p.context === 'cell');
  assert.ok(cells.length > 50, `셀 문단 ${cells.length}개`);
  const c = cells.find(x => x.text.length >= 2)!;
  doc.setParagraphText(c.section, c.index, '셀편집시험');
  const again = HwpxDocument.open(doc.save()).paragraphs().find(x => x.section === c.section && x.index === c.index)!;
  assert.equal(again.text, '셀편집시험');
  assert.equal(again.context, 'cell');
});

test('머리말·꼬리말·각주 문맥(합성 XML): 실제 시료에는 이런 글자가 없어 합성 구문으로만 확인한다', () => {
  const xml = '<hs:sec><hp:p id="1"><hp:run><hp:ctrl><hp:header><hp:subList><hp:p id="2"><hp:run><hp:t>머리말</hp:t></hp:run></hp:p></hp:subList></hp:header></hp:ctrl>' +
    '<hp:ctrl><hp:footer><hp:subList><hp:p id="3"><hp:run><hp:t>꼬리말</hp:t></hp:run></hp:p></hp:subList></hp:footer></hp:ctrl>' +
    '<hp:ctrl><hp:footNote><hp:subList><hp:p id="4"><hp:run><hp:t>각주</hp:t></hp:run></hp:p></hp:subList></hp:footNote></hp:ctrl>' +
    '<hp:t>본문</hp:t></hp:run></hp:p></hs:sec>';
  const s = scanSection(xml);
  assert.deepEqual(s.texts.map(t => [t.text, t.context]),
    [['머리말', 'header'], ['꼬리말', 'footer'], ['각주', 'footnote'], ['본문', 'body']]);
});

test('실제 시료의 문맥 분포(셀, 글상자). 머리말·꼬리말·각주가 든 실제 시료는 아직 없다', () => {
  const seen = new Set<string>();
  for (const f of ['hwpx/pipc-privacy-20231127.hwpx', 'hwpx/pps-spec-14-13.hwpx']) {
    for (const p of HwpxDocument.open(fx(f)).paragraphs()) seen.add(p.context);
  }
  assert.ok(seen.has('cell') && seen.has('textbox') && seen.has('body'));
  assert.ok(!seen.has('header') && !seen.has('footer') && !seen.has('footnote'));
});

test('바이트 보존·XML 구조 보존·글자 보존을 따로 확인한다(글자 편집 한 군데)', () => {
  const orig = fx('hwpx/pipc-privacy-20231127.hwpx');
  const doc = HwpxDocument.open(orig);
  const p = doc.paragraphs().find(x => x.context === 'body' && x.text.length > 5)!;
  doc.setParagraphText(p.section, p.index, p.text + '추가', { layout: 'keep' });
  const out = doc.save();
  const a = new ZipArchive(orig), b = new ZipArchive(out);
  // 1) 바이트 보존: 바뀌지 않은 항목은 압축 바이트까지 같다
  let sameRaw = 0;
  for (const e of b.entries) {
    const o = a.entries.find(x => x.name === e.name)!;
    if (e.name !== p.section) { assert.deepEqual(e.raw, o.raw, e.name); sameRaw++; }
  }
  assert.equal(sameRaw, a.entries.length - 1);
  // 2) XML 구조 보존: 글자 내용을 지운 뒤 같다
  const strip = (x: string) => x.replace(/<hp:t>[^<]*<\/hp:t>/g, '<hp:t/>');
  assert.equal(strip(b.readText(p.section)), strip(a.readText(p.section)));
  // 3) 글자 보존: 편집한 문단 외 모든 문단 글자가 같다
  const A = HwpxDocument.open(orig).paragraphs(), B = HwpxDocument.open(out).paragraphs();
  assert.equal(A.length, B.length);
  for (let i = 0; i < A.length; i++) if (A[i].index !== p.index) assert.equal(B[i].text, A[i].text);
});

test('수정 일시와 미리보기 갱신 정책', () => {
  const doc = HwpxDocument.open(fx(MANUAL));
  const hpfBefore = new ZipArchive(fx(MANUAL)).readText('Contents/content.hpf');
  assert.ok(doc.touchModified(new Date('2026-10-22T03:04:05.678Z')));
  const hpf = doc.zip.readText('Contents/content.hpf');
  assert.match(hpf, /name="ModifiedDate" content="text">2026-10-22T03:04:05Z</);
  assert.equal(hpf.replace(/2026-10-22T03:04:05Z/, '2025-06-02T01:07:00Z'), hpfBefore, 'ModifiedDate 외 content.hpf 동일');
  // 미리보기: 한 번만 나오는 글자만 바꾼다
  const prv = doc.zip.readText('Preview/PrvText.txt');
  const word = prv.match(/[가-힣]{4,}/g)!.find(w => prv.split(w).length === 2)!;
  assert.ok(doc.syncPreviewText(word, '바뀐말'));
  assert.ok(doc.zip.readText('Preview/PrvText.txt').includes('바뀐말'));
  assert.equal(doc.syncPreviewText('절대없는글자열', 'x'), false);
  assert.equal(doc.syncPreviewText('', 'x'), false);
});

test('UTF-8 BOM이 있는 부품도 BOM을 지키며 편집한다', () => {
  // 합성 ZIP 없이 TextPart 수준에서: BOM이 남는지
  const doc = HwpxDocument.open(fx(MANUAL));
  const s = doc.sectionNames()[0];
  const bytes = doc.zip.read(s);
  const bom = new Uint8Array(bytes.length + 3);
  bom.set([0xef, 0xbb, 0xbf]); bom.set(bytes, 3);
  doc.zip.replace(s, bom);
  const fresh = HwpxDocument.open(doc.save());
  fresh.setText(s, 0, fresh.scan(s).texts[0].text + '!');
  const out = new Uint8Array(new ZipArchive(fresh.save()).read(s));
  assert.deepEqual([...out.subarray(0, 3)], [0xef, 0xbb, 0xbf]);
});

test('형식 판별과 지원 여부', () => {
  assert.equal(detectFormat(fx(MANUAL)).id, 'hwpx');
  assert.equal(detectFormat(fx('docx/basic.docx')).id, 'docx');
  assert.equal(detectFormat(new TextEncoder().encode('%PDF-1.7\n')).id, 'pdf');
  const ole = new Uint8Array(16); ole.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  assert.equal(detectFormat(ole, 'a.hwp').id, 'hwp');
  assert.equal(detectFormat(ole, 'a.doc').id, 'doc');
  assert.equal(detectFormat(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8])).id, 'unknown');
  assert.throws(() => openDocument(ole, 'a.hwp'), UnsupportedFormatError);
  assert.equal(openDocument(fx(MANUAL)).format.id, 'hwpx');
  assert.equal(openDocument(fx('docx/table.docx')).format.id, 'docx');
});

test('자기닫힘 문단(<hp:p/>)이 뒤 문단의 소속을 어긋나게 하지 않는다', () => {
  const xml = '<hs:sec><hp:p id="1"/><hp:p id="2"><hp:run><hp:t>둘째</hp:t></hp:run></hp:p><hp:p id="3"><hp:run><hp:t>셋째</hp:t></hp:run></hp:p></hs:sec>';
  const s = scanSection(xml);
  assert.equal(s.paras.length, 3);
  assert.deepEqual(s.texts.map(t => t.paraIndex), [1, 2]);
});

test('DRM으로 암호화된 파일은 형식을 알아보고 열기를 거부한다(우회하지 않는다)', () => {
  // 이 PC의 한/글이 저장한 파일의 머리글을 본뜬 합성 바이트(실제 암호화 본문 없음)
  const head = new TextEncoder().encode(String.fromCharCode(0x9b) + ' DRMONE  This Document is encrypted and protected by Fasoo DRM ');
  const bytes = new Uint8Array(4096);
  bytes.set(head);
  const info = detectFormat(bytes, 'a.hwpx');
  assert.equal(info.id, 'drm');
  assert.equal(info.support, 'none');
  assert.match(info.note, /우회하지 않습니다/);
  assert.throws(() => openDocument(bytes, 'a.hwpx'), UnsupportedFormatError);
});
