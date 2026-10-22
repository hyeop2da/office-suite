// 두 XML 부품과 두 ZIP 패키지의 차이를 세는 비교 도구(분석·시험용. 편집 코어와 별개).
// 요소는 같은 이름끼리 문서 순서대로 짝짓는다. 같은 이름의 요소 수가 다르면 짝짓지 않고 수만 기록한다.
import { ZipArchive } from './zip.ts';
import { decodeEntities } from './xml-scan.ts';

export interface XNode { name: string; attrs: [string, string][]; children: XNode[]; text: string }

const TAG = /<(\/?)([A-Za-z_][\w:.-]*)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/y;
const ATTR = /([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

export function parseXml(xml: string): XNode {
  const root: XNode = { name: '#root', attrs: [], children: [], text: '' };
  const stack: XNode[] = [root];
  let i = 0;
  while (i < xml.length) {
    const lt = xml.indexOf('<', i);
    const textEnd = lt < 0 ? xml.length : lt;
    if (textEnd > i) stack[stack.length - 1].text += decodeEntities(xml.slice(i, textEnd));
    if (lt < 0) break;
    if (xml.startsWith('<?', lt)) { i = xml.indexOf('?>', lt) + 2; continue; }
    if (xml.startsWith('<!--', lt)) { i = xml.indexOf('-->', lt) + 3; continue; }
    if (xml.startsWith('<![CDATA[', lt)) {
      const e = xml.indexOf(']]>', lt);
      stack[stack.length - 1].text += xml.slice(lt + 9, e);
      i = e + 3; continue;
    }
    if (xml.startsWith('<!', lt)) { i = xml.indexOf('>', lt) + 1; continue; }
    TAG.lastIndex = lt;
    const m = TAG.exec(xml);
    if (!m) throw new Error(`XML 태그 해석 실패(위치 ${lt})`);
    const [whole, closing, name, attrSrc, self] = m;
    if (closing) {
      const top = stack.pop();
      if (!top || top.name !== name) throw new Error(`닫는 태그 불일치: ${name}`);
    } else {
      const attrs: [string, string][] = [];
      ATTR.lastIndex = 0;
      for (let a = ATTR.exec(attrSrc); a; a = ATTR.exec(attrSrc)) attrs.push([a[1], decodeEntities(a[2] ?? a[3] ?? '')]);
      const node: XNode = { name, attrs, children: [], text: '' };
      stack[stack.length - 1].children.push(node);
      if (!self) stack.push(node);
    }
    i = lt + whole.length;
  }
  if (stack.length !== 1) throw new Error('닫히지 않은 태그');
  return root;
}

function flatten(n: XNode, out: Map<string, XNode[]>): void {
  for (const c of n.children) {
    let l = out.get(c.name);
    if (!l) out.set(c.name, l = []);
    l.push(c);
    flatten(c, out);
  }
}

export interface XmlDiff {
  /** 이름별 요소 수가 달라진 것: [원본, 새 파일] */
  tagCount: Record<string, [number, number]>;
  /** 요소 수가 같은 이름에서, 원본에만 있던 속성 "태그@속성" → 개수 */
  attrLost: Record<string, number>;
  attrAdded: Record<string, number>;
  attrChanged: Record<string, number>;
  /** 요소 수가 달라 짝짓지 못한 이름 */
  unaligned: string[];
  /** 글자가 비어 있는 요소 수(원본, 새 파일). 지정한 이름만 */
  emptyCount: Record<string, [number, number]>;
}

export function diffXml(a: string, b: string, emptyTags: string[] = ['hp:t']): XmlDiff {
  const fa = new Map<string, XNode[]>(), fb = new Map<string, XNode[]>();
  flatten(parseXml(a), fa); flatten(parseXml(b), fb);
  const res: XmlDiff = { tagCount: {}, attrLost: {}, attrAdded: {}, attrChanged: {}, unaligned: [], emptyCount: {} };
  const inc = (o: Record<string, number>, k: string) => { o[k] = (o[k] ?? 0) + 1; };
  for (const name of new Set([...fa.keys(), ...fb.keys()])) {
    const la = fa.get(name) ?? [], lb = fb.get(name) ?? [];
    if (la.length !== lb.length) { res.tagCount[name] = [la.length, lb.length]; res.unaligned.push(name); continue; }
    for (let i = 0; i < la.length; i++) {
      const ma = new Map(la[i].attrs), mb = new Map(lb[i].attrs);
      for (const [k, v] of ma) {
        if (!mb.has(k)) inc(res.attrLost, `${name}@${k}`);
        else if (mb.get(k) !== v) inc(res.attrChanged, `${name}@${k}`);
      }
      for (const k of mb.keys()) if (!ma.has(k)) inc(res.attrAdded, `${name}@${k}`);
    }
  }
  for (const t of emptyTags) {
    const cnt = (l: XNode[] = []) => l.filter(n => n.text === '' && n.children.length === 0).length;
    res.emptyCount[t] = [cnt(fa.get(t)), cnt(fb.get(t))];
  }
  return res;
}

export interface ZipDiff {
  names: { a: string[]; b: string[] };
  added: string[];
  removed: string[];
  /** 압축 해제한 내용이 같은 항목 */
  identical: string[];
  changed: string[];
  /** 내용은 같고 압축 방식만 달라진 항목 */
  methodChanged: string[];
  orderChanged: boolean;
  /** 같은 이름 항목의 압축된 바이트까지 같은 것 */
  rawIdentical: string[];
}

export function diffZip(aBytes: Uint8Array, bBytes: Uint8Array): ZipDiff {
  const a = new ZipArchive(aBytes), b = new ZipArchive(bBytes);
  const an = a.entries.map(e => e.name), bn = b.entries.map(e => e.name);
  const res: ZipDiff = { names: { a: an, b: bn }, added: bn.filter(n => !an.includes(n)), removed: an.filter(n => !bn.includes(n)), identical: [], changed: [], methodChanged: [], orderChanged: false, rawIdentical: [] };
  const common = an.filter(n => bn.includes(n));
  res.orderChanged = JSON.stringify(common) !== JSON.stringify(bn.filter(n => an.includes(n)));
  for (const n of common) {
    const ea = a.entries.find(e => e.name === n)!, eb = b.entries.find(e => e.name === n)!;
    const same = Buffer.compare(Buffer.from(a.read(n)), Buffer.from(b.read(n))) === 0;
    if (same) { res.identical.push(n); if (ea.method !== eb.method) res.methodChanged.push(n); }
    else res.changed.push(n);
    if (Buffer.compare(Buffer.from(ea.raw), Buffer.from(eb.raw)) === 0) res.rawIdentical.push(n);
  }
  return res;
}

export interface PackageDiff {
  zip: ZipDiff;
  /** 내용이 달라진 XML 부품별 요소·속성 차이 */
  xml: Record<string, XmlDiff>;
}

/** 두 HWPX/OOXML 패키지를 비교한다. XML이 아닌 항목은 바이트 비교만 한다. */
export function diffPackages(aBytes: Uint8Array, bBytes: Uint8Array): PackageDiff {
  const zip = diffZip(aBytes, bBytes);
  const a = new ZipArchive(aBytes), b = new ZipArchive(bBytes);
  const xml: Record<string, XmlDiff> = {};
  for (const n of zip.changed) {
    if (!/\.(xml|hpf|rels|rdf)$/i.test(n)) continue;
    xml[n] = diffXml(a.readText(n), b.readText(n));
  }
  return { zip, xml };
}
