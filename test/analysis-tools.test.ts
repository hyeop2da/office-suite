// 분석 도구(inventory, hwp_inspect)가 시료에 대해 재현 가능한 값을 낸다. 시료 적합성(커버리지 공백)도 못 박아 둔다.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { inventory } from '../src/inventory.ts';

const dir = (d: string) => new URL(`../fixtures/${d}/`, import.meta.url);
const inv = (d: string, f: string) => inventory(new Uint8Array(readFileSync(new URL(f, dir(d)))));

test('한/글 작성 시료의 기능 목록(그림·글상자·누름틀·쪽 번호는 pipc에만 있다)', () => {
  const pipc = inv('hwpx', 'pipc-privacy-20231127.hwpx');
  assert.equal(pipc.counts['그림(hp:pic)'], 4);
  assert.equal(pipc.binData, 4);
  assert.equal(pipc.counts['글상자(hp:drawText)'], 1);
  assert.equal(pipc.counts['쪽 번호(hp:pageNum)'], 1);
  assert.deepEqual(pipc.fieldTypes, { CLICK_HERE: 1 });
  assert.equal(inv('hwpx', 'pps-spec-14-13.hwpx').counts['표 셀(hp:tc)'], 223);
});

test('커버리지 공백: 한/글 작성 시료 5개 모두에 아래 기능이 없다(생기면 이 시험이 알려 준다)', () => {
  const absent = ['머리말(hp:header)', '꼬리말(hp:footer)', '각주(hp:footNote)', '미주(hp:endNote)', '책갈피(hp:bookmark)', '수식(hp:equation)', 'OLE(hp:ole)',
    '탭(hp:tab)', '줄바꿈(hp:lineBreak)', '자동 번호(hp:autoNum)', '변경 추적 표시(hp:insertBegin·deleteBegin)', '여러 단(colCount>1)', '묶음 개체(hp:container)'];
  for (const f of readdirSync(dir('hwpx')).filter(x => x.endsWith('.hwpx'))) {
    const i = inv('hwpx', f);
    assert.equal(i.sections, 1, `${f} 구역 수`);
    for (const k of absent) assert.equal(i.counts[k], 0, `${f}: ${k}`);
    assert.ok(!('HYPERLINK' in i.fieldTypes), `${f}: 하이퍼링크는 한/글 시료에 없음`);
  }
});

test('HWP 5.x 시료 구조(olefile): 버전 5.x, 압축, 구역 1개, rhwp 저장본과 레코드 수 동일', (t) => {
  const r = spawnSync('python', ['-I', '-X', 'utf8', 'tools/hwp_inspect.py', 'fixtures/hwp'], { encoding: 'utf8', cwd: new URL('..', import.meta.url) });
  if (r.error) return t.skip('python 없음');
  assert.equal(r.status, 0, r.stderr);
  const rows = JSON.parse(r.stdout) as { file: string; version: string; flags: string[]; sections: number; section0: { bytes: number; records: number; ['HWPTAG_PARA_TEXT(0x43=67)']: number; ['HWPTAG_TABLE(0x4d=77)']: number } }[];
  assert.equal(rows.length, 10);
  for (const x of rows) {
    assert.match(x.version, /^5\./);
    assert.deepEqual(x.flags, ['압축']);
    assert.equal(x.sections, 1);
  }
  for (const orig of rows.filter(x => !x.file.includes('rhwp-saved'))) {
    const saved = rows.find(x => x.file === orig.file.replace('.hwp', '.rhwp-saved.hwp'))!;
    const { bytes: ob, ...oRest } = orig.section0, { bytes: sb, ...sRest } = saved.section0;
    assert.deepEqual(sRest, oRest, `${orig.file}: 첫 구역 레코드·문단·표 수`);
    // 이전 하네스가 넣은 마커 〔트랙A수정〕은 7글자(UTF-16 14바이트)다. 실측: 두 파일은 정확히 14바이트,
    // 세 파일은 50바이트 커졌다(14 + 36). 36바이트의 정체는 아직 분석하지 않았다(미확인).
    const measured: Record<string, number> = { 'assembly-audit.hwp': 14, 'assembly-briefing.hwp': 14, 'assembly-nabo.hwp': 50, 'seoul-judge-sample.hwp': 50, 'seoul-menu-3e.hwp': 50 };
    assert.equal(sb - ob, measured[orig.file], `${orig.file}: 압축 해제한 구역 크기 차이`);
  }
});
