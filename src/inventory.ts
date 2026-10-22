// HWPX 시료가 어떤 기능을 실제로 담고 있는지 세는 도구. 시험 범위(시료 적합성)를 숫자로 보이기 위한 것이다.
import { ZipArchive } from './zip.ts';
import { parseXml, type XNode } from './xmldiff.ts';

export interface HwpxInventory {
  sections: number;
  binData: number;
  counts: Record<string, number>;
  fieldTypes: Record<string, number>;
}

/** 이름 → 세는 규칙. 앞쪽은 요소 이름, 함수는 조건부 계산. */
const TAGS: Record<string, string[]> = {
  '표(hp:tbl)': ['hp:tbl'],
  '표 셀(hp:tc)': ['hp:tc'],
  '그림(hp:pic)': ['hp:pic'],
  '도형(rect·ellipse·line·curve·polygon·arc)': ['hp:rect', 'hp:ellipse', 'hp:line', 'hp:curve', 'hp:polygon', 'hp:arc'],
  '묶음 개체(hp:container)': ['hp:container'],
  '글상자(hp:drawText)': ['hp:drawText'],
  'OLE(hp:ole)': ['hp:ole'],
  '수식(hp:equation)': ['hp:equation'],
  '머리말(hp:header)': ['hp:header'],
  '꼬리말(hp:footer)': ['hp:footer'],
  '각주(hp:footNote)': ['hp:footNote'],
  '미주(hp:endNote)': ['hp:endNote'],
  '책갈피(hp:bookmark)': ['hp:bookmark'],
  '쪽 번호(hp:pageNum)': ['hp:pageNum'],
  '자동 번호(hp:autoNum)': ['hp:autoNum'],
  '탭(hp:tab)': ['hp:tab'],
  '줄바꿈(hp:lineBreak)': ['hp:lineBreak'],
  '변경 추적 표시(hp:insertBegin·deleteBegin)': ['hp:insertBegin', 'hp:deleteBegin'],
  '메모 정의(hh:memoPr)': ['hh:memoPr'],
};

function walk(n: XNode, f: (n: XNode) => void): void { for (const c of n.children) { f(c); walk(c, f); } }

export function inventory(bytes: Uint8Array): HwpxInventory {
  const zip = new ZipArchive(bytes);
  const secs = zip.entries.map(e => e.name).filter(n => /^Contents\/section\d+\.xml$/.test(n));
  const inv: HwpxInventory = { sections: secs.length, binData: zip.entries.filter(e => e.name.startsWith('BinData/')).length, counts: {}, fieldTypes: {} };
  for (const k of Object.keys(TAGS)) inv.counts[k] = 0;
  inv.counts['여러 단(colCount>1)'] = 0;
  const parts = [...secs, 'Contents/header.xml'].filter(n => zip.has(n));
  for (const name of parts) {
    walk(parseXml(zip.readText(name)), n => {
      for (const [k, tags] of Object.entries(TAGS)) if (tags.includes(n.name)) inv.counts[k]++;
      if (n.name === 'hp:fieldBegin') { const t = n.attrs.find(a => a[0] === 'type')?.[1] ?? '?'; inv.fieldTypes[t] = (inv.fieldTypes[t] ?? 0) + 1; }
      if (n.name === 'hp:colPr' && Number(n.attrs.find(a => a[0] === 'colCount')?.[1] ?? 1) > 1) inv.counts['여러 단(colCount>1)']++;
    });
  }
  return inv;
}
