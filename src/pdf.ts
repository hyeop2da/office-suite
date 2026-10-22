// PDF 최소 코어: 쪽 수·쪽 크기 읽기, 텍스트 메모(/Text 주석) 목록과 추가.
// 저장은 **증분 갱신**(원본 바이트는 그대로 두고 끝에 새 객체와 xref를 덧붙임)이라 원본이 손상되지 않는다.
// 암호화된 PDF와 손상된 xref는 거부한다. 쪽 그리기는 앱 화면에서 pdf.js가 맡고, 이 파일은 구조만 다룬다.
import { unzlibSync } from 'fflate';

export class PdfError extends Error { constructor(m: string) { super(m); this.name = 'PdfError'; } }

// ---------- 값 모델 ----------
export type PdfValue =
  | { t: 'null' } | { t: 'bool'; v: boolean } | { t: 'num'; raw: string }
  | { t: 'name'; v: string } | { t: 'str'; v: Uint8Array }
  | { t: 'ref'; n: number; g: number } | { t: 'arr'; v: PdfValue[] }
  | { t: 'dict'; v: Map<string, PdfValue>; stream?: { start: number; end: number } };

const WS = new Set([0, 9, 10, 12, 13, 32]);
const DELIM = new Set('()<>[]{}/%'.split('').map(c => c.charCodeAt(0)));
const isWs = (c: number) => WS.has(c);
const isReg = (c: number) => !WS.has(c) && !DELIM.has(c);

class Lexer {
  pos: number;
  readonly b: Uint8Array;
  constructor(b: Uint8Array, pos = 0) { this.b = b; this.pos = pos; }
  skipWs(): void {
    for (;;) {
      while (this.pos < this.b.length && isWs(this.b[this.pos])) this.pos++;
      if (this.b[this.pos] === 0x25) { while (this.pos < this.b.length && this.b[this.pos] !== 10 && this.b[this.pos] !== 13) this.pos++; continue; }
      return;
    }
  }
  word(): string {
    this.skipWs();
    const s = this.pos;
    while (this.pos < this.b.length && isReg(this.b[this.pos])) this.pos++;
    return String.fromCharCode(...this.b.subarray(s, this.pos));
  }
  peekWord(): string { const p = this.pos; const w = this.word(); this.pos = p; return w; }

  value(depth = 0): PdfValue {
    if (depth > 64) throw new PdfError('PDF 구문이 너무 깊게 중첩되었습니다');
    this.skipWs();
    const c = this.b[this.pos];
    if (c === undefined) throw new PdfError('PDF가 예기치 않게 끝났습니다');
    if (c === 0x2f) { this.pos++; return { t: 'name', v: this.nameBody() }; }
    if (c === 0x28) return { t: 'str', v: this.literalString() };
    if (c === 0x3c) {
      if (this.b[this.pos + 1] === 0x3c) return this.dict(depth);
      return { t: 'str', v: this.hexString() };
    }
    if (c === 0x5b) {
      this.pos++;
      const v: PdfValue[] = [];
      for (;;) { this.skipWs(); if (this.b[this.pos] === 0x5d) { this.pos++; break; } if (this.pos >= this.b.length) throw new PdfError('배열이 닫히지 않았습니다'); v.push(this.value(depth + 1)); }
      return { t: 'arr', v };
    }
    const w = this.word();
    if (w === '') throw new PdfError(`PDF 구문 오류(위치 ${this.pos})`);
    if (w === 'true' || w === 'false') return { t: 'bool', v: w === 'true' };
    if (w === 'null') return { t: 'null' };
    if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(w)) {
      if (/^\d+$/.test(w)) { // 간접 참조 "n g R"인지 본다
        const save = this.pos;
        const g = this.word();
        if (/^\d+$/.test(g) && this.word() === 'R') return { t: 'ref', n: Number(w), g: Number(g) };
        this.pos = save;
      }
      return { t: 'num', raw: w };
    }
    throw new PdfError(`알 수 없는 PDF 토큰: ${w}`);
  }
  private nameBody(): string {
    let s = '';
    while (this.pos < this.b.length && isReg(this.b[this.pos])) {
      const c = this.b[this.pos];
      if (c === 0x23 && this.pos + 2 < this.b.length) { s += String.fromCharCode(parseInt(String.fromCharCode(this.b[this.pos + 1], this.b[this.pos + 2]), 16)); this.pos += 3; }
      else { s += String.fromCharCode(c); this.pos++; }
    }
    return s;
  }
  private literalString(): Uint8Array {
    this.pos++;
    const out: number[] = [];
    let depth = 1;
    while (this.pos < this.b.length) {
      let c = this.b[this.pos++];
      if (c === 0x5c) {
        c = this.b[this.pos++];
        if (c === 0x6e) out.push(10); else if (c === 0x72) out.push(13); else if (c === 0x74) out.push(9);
        else if (c === 0x62) out.push(8); else if (c === 0x66) out.push(12);
        else if (c >= 0x30 && c <= 0x37) { let v = c - 0x30; for (let k = 0; k < 2; k++) { const d = this.b[this.pos]; if (d >= 0x30 && d <= 0x37) { v = v * 8 + d - 0x30; this.pos++; } else break; } out.push(v & 255); }
        else if (c === 13) { if (this.b[this.pos] === 10) this.pos++; }
        else if (c === 10) { /* 줄 이음 */ }
        else out.push(c);
      } else if (c === 0x28) { depth++; out.push(c); }
      else if (c === 0x29) { depth--; if (depth === 0) return Uint8Array.from(out); out.push(c); }
      else out.push(c);
    }
    throw new PdfError('문자열이 닫히지 않았습니다');
  }
  private hexString(): Uint8Array {
    this.pos++;
    const out: number[] = [];
    let hi = -1;
    for (;;) {
      const c = this.b[this.pos++];
      if (c === undefined) throw new PdfError('16진 문자열이 닫히지 않았습니다');
      if (c === 0x3e) break;
      if (isWs(c)) continue;
      const v = parseInt(String.fromCharCode(c), 16);
      if (Number.isNaN(v)) throw new PdfError('16진 문자열 오류');
      if (hi < 0) hi = v; else { out.push(hi * 16 + v); hi = -1; }
    }
    if (hi >= 0) out.push(hi * 16);
    return Uint8Array.from(out);
  }
  private dict(depth: number): PdfValue {
    this.pos += 2;
    const v = new Map<string, PdfValue>();
    for (;;) {
      this.skipWs();
      if (this.b[this.pos] === 0x3e && this.b[this.pos + 1] === 0x3e) { this.pos += 2; break; }
      if (this.b[this.pos] !== 0x2f) throw new PdfError(`딕셔너리 키가 이름이 아닙니다(위치 ${this.pos})`);
      this.pos++;
      const k = this.nameBody();
      v.set(k, this.value(depth + 1));
    }
    const d: PdfValue = { t: 'dict', v };
    // 스트림이 이어지는지
    const save = this.pos;
    if (this.word() === 'stream') {
      if (this.b[this.pos] === 13) this.pos++;
      if (this.b[this.pos] === 10) this.pos++;
      d.stream = { start: this.pos, end: -1 };
    } else this.pos = save;
    return d;
  }
}

// ---------- 직렬화 ----------
const enc = new TextEncoder();
function serName(n: string): string { return '/' + [...n].map(ch => { const c = ch.charCodeAt(0); return c <= 32 || c >= 127 || '()<>[]{}/%#'.includes(ch) ? '#' + c.toString(16).padStart(2, '0') : ch; }).join(''); }
function hex(b: Uint8Array): string { return '<' + [...b].map(x => x.toString(16).padStart(2, '0')).join('') + '>'; }
export function serialize(v: PdfValue): string {
  switch (v.t) {
    case 'null': return 'null';
    case 'bool': return v.v ? 'true' : 'false';
    case 'num': return v.raw;
    case 'name': return serName(v.v);
    case 'str': return hex(v.v);
    case 'ref': return `${v.n} ${v.g} R`;
    case 'arr': return '[' + v.v.map(serialize).join(' ') + ']';
    case 'dict': return '<<' + [...v.v].map(([k, x]) => `${serName(k)} ${serialize(x)}`).join('') + '>>';
  }
}
const num = (n: number): PdfValue => ({ t: 'num', raw: Number.isInteger(n) ? String(n) : n.toFixed(3).replace(/0+$/, '').replace(/\.$/, '') });
const name = (s: string): PdfValue => ({ t: 'name', v: s });
/** 글자를 PDF 텍스트 문자열(UTF-16BE + BOM)로 */
export function textString(s: string): PdfValue {
  const out = [0xfe, 0xff];
  for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); out.push(c >> 8, c & 255); }
  return { t: 'str', v: Uint8Array.from(out) };
}
export function decodeText(b: Uint8Array): string {
  if (b[0] === 0xfe && b[1] === 0xff) { let s = ''; for (let i = 2; i + 1 < b.length; i += 2) s += String.fromCharCode((b[i] << 8) | b[i + 1]); return s; }
  if (b[0] === 0xff && b[1] === 0xfe) { let s = ''; for (let i = 2; i + 1 < b.length; i += 2) s += String.fromCharCode((b[i + 1] << 8) | b[i]); return s; }
  if (b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) return new TextDecoder().decode(b.subarray(3));
  return String.fromCharCode(...b); // PDFDocEncoding(라틴 범위 근사)
}

// ---------- 문서 ----------
interface XEntry { type: 1 | 2; a: number; b: number } // 1: 오프셋a 세대b, 2: 객체스트림 a의 b번째
export interface TextNote { id: number; page: number; x: number; y: number; contents: string; author?: string; isNew?: boolean }
export interface PageInfo { index: number; objNum: number; width: number; height: number }

export class PdfDocument {
  readonly bytes: Uint8Array;
  private xref = new Map<number, XEntry>();
  private trailer = new Map<string, PdfValue>();
  private cache = new Map<number, PdfValue>();
  private objStreams = new Map<number, Map<number, PdfValue>>();
  private usesXrefStream = false;
  private firstXrefOffset = 0;
  private pageList: PageInfo[] | undefined;

  private constructor(bytes: Uint8Array) { this.bytes = bytes; }

  static open(bytes: Uint8Array): PdfDocument {
    if (bytes.length < 20 || String.fromCharCode(...bytes.subarray(0, 5)) !== '%PDF-') throw new PdfError('PDF가 아닙니다(%PDF- 머리글 없음)');
    const d = new PdfDocument(bytes);
    try { d.readXref(); } catch (e) { if (e instanceof PdfError) throw e; throw new PdfError(`PDF 구조를 읽지 못했습니다: ${(e as Error).message}`); }
    if (d.trailer.has('Encrypt')) throw new PdfError('암호화된 PDF는 지원하지 않습니다');
    if (!d.trailer.has('Root')) throw new PdfError('PDF에 Root가 없습니다');
    return d;
  }

  private readXref(): void {
    const tail = String.fromCharCode(...this.bytes.subarray(Math.max(0, this.bytes.length - 2048)));
    const sx = tail.lastIndexOf('startxref');
    const m = sx < 0 ? null : /^startxref\s+(\d+)/.exec(tail.slice(sx));
    if (!m) throw new PdfError('startxref를 찾지 못했습니다');
    let off = Number(m[1]);
    this.firstXrefOffset = off;
    const seen = new Set<number>();
    let first = true;
    while (off >= 0) {
      if (seen.has(off) || off >= this.bytes.length) throw new PdfError('xref 연결이 올바르지 않습니다');
      seen.add(off);
      const lx = new Lexer(this.bytes, off);
      let trailer: Map<string, PdfValue>;
      if (lx.peekWord() === 'xref') {
        lx.word();
        for (;;) {
          lx.skipWs();
          if (lx.peekWord() === 'trailer') { lx.word(); break; }
          const start = Number(lx.word()), count = Number(lx.word());
          if (!Number.isInteger(start) || !Number.isInteger(count) || count < 0) throw new PdfError('xref 구간이 올바르지 않습니다');
          for (let i = 0; i < count; i++) {
            const o = lx.word(), g = lx.word(), f = lx.word();
            if (!this.xref.has(start + i) && f === 'n') this.xref.set(start + i, { type: 1, a: Number(o), b: Number(g) });
            else if (!this.xref.has(start + i) && f === 'f') this.xref.set(start + i, { type: 1, a: -1, b: 0 });
          }
        }
        const t = lx.value();
        if (t.t !== 'dict') throw new PdfError('trailer가 올바르지 않습니다');
        trailer = t.v;
        if (first) this.usesXrefStream = false;
        const xs = trailer.get('XRefStm');
        if (xs && xs.t === 'num') this.readXrefStream(Number(xs.raw), false);
      } else {
        trailer = this.readXrefStream(off, true);
        if (first) this.usesXrefStream = true;
      }
      if (first) { for (const [k, v] of trailer) this.trailer.set(k, v); first = false; }
      const prev = trailer.get('Prev');
      off = prev && prev.t === 'num' ? Number(prev.raw) : -1;
    }
  }

  private readXrefStream(off: number, wantTrailer: boolean): Map<string, PdfValue> {
    const lx = new Lexer(this.bytes, off);
    lx.word(); lx.word(); if (lx.word() !== 'obj') throw new PdfError('xref 스트림 객체가 아닙니다');
    const d = lx.value();
    if (d.t !== 'dict' || !d.stream) throw new PdfError('xref 스트림이 올바르지 않습니다');
    const data = this.streamData(d);
    const w = (d.v.get('W') as { t: 'arr'; v: PdfValue[] } | undefined)?.v.map(x => Number((x as { raw: string }).raw));
    if (!w || w.length !== 3) throw new PdfError('xref 스트림의 W가 올바르지 않습니다');
    const size = Number((d.v.get('Size') as { raw: string }).raw);
    const index = (d.v.get('Index') as { t: 'arr'; v: PdfValue[] } | undefined)?.v.map(x => Number((x as { raw: string }).raw)) ?? [0, size];
    const rl = w[0] + w[1] + w[2];
    let p = 0;
    const rd = (n: number): number => { let v = 0; for (let i = 0; i < n; i++) v = v * 256 + data[p++]; return v; };
    for (let k = 0; k < index.length; k += 2) {
      for (let i = 0; i < index[k + 1]; i++) {
        if (p + rl > data.length) throw new PdfError('xref 스트림이 짧습니다');
        const t = w[0] === 0 ? 1 : rd(w[0]), a = rd(w[1]), b = rd(w[2]);
        const n = index[k] + i;
        if (this.xref.has(n)) continue;
        if (t === 1) this.xref.set(n, { type: 1, a, b });
        else if (t === 2) this.xref.set(n, { type: 2, a, b });
        else if (t === 0) this.xref.set(n, { type: 1, a: -1, b: 0 });
      }
    }
    return wantTrailer ? d.v : new Map();
  }

  /** 스트림 데이터(FlateDecode만 지원, PNG 예측기 지원) */
  private streamData(d: PdfValue & { t: 'dict' }): Uint8Array {
    const s = d.stream!;
    let len = this.resolveNum(d.v.get('Length'));
    if (len === undefined || s.start + len > this.bytes.length) { // 길이가 없거나 틀리면 endstream을 찾는다
      const hay = new TextDecoder('latin1').decode(this.bytes.subarray(s.start, Math.min(this.bytes.length, s.start + 50_000_000)));
      const e = hay.indexOf('endstream');
      if (e < 0) throw new PdfError('endstream을 찾지 못했습니다');
      len = Math.max(0, e - (hay[e - 1] === '\n' ? (hay[e - 2] === '\r' ? 2 : 1) : hay[e - 1] === '\r' ? 1 : 0));
    }
    let data = this.bytes.subarray(s.start, s.start + len);
    const filter = d.v.get('Filter');
    const filters = filter ? (filter.t === 'arr' ? filter.v : [filter]) : [];
    for (const f of filters) {
      const fname = f.t === 'name' ? f.v : '';
      if (fname !== 'FlateDecode' && fname !== 'Fl') throw new PdfError(`지원하지 않는 스트림 필터입니다: ${fname}`);
      try { data = unzlibSync(data); } catch { throw new PdfError('스트림 압축 해제에 실패했습니다'); }
      const parms = d.v.get('DecodeParms');
      const dp = parms && parms.t === 'dict' ? parms : undefined;
      const pred = dp ? this.resolveNum(dp.v.get('Predictor')) : undefined;
      if (pred && pred >= 10) data = this.pngPredictor(data, this.resolveNum(dp!.v.get('Columns')) ?? 1);
    }
    return data;
  }

  private pngPredictor(data: Uint8Array, columns: number): Uint8Array {
    const rl = columns + 1, rows = Math.floor(data.length / rl);
    const out = new Uint8Array(rows * columns);
    for (let r = 0; r < rows; r++) {
      const f = data[r * rl];
      for (let c = 0; c < columns; c++) {
        const x = data[r * rl + 1 + c];
        const a = c > 0 ? out[r * columns + c - 1] : 0, b = r > 0 ? out[(r - 1) * columns + c] : 0, cc = c > 0 && r > 0 ? out[(r - 1) * columns + c - 1] : 0;
        let v = x;
        if (f === 1) v = x + a; else if (f === 2) v = x + b; else if (f === 3) v = x + ((a + b) >> 1);
        else if (f === 4) { const p = a + b - cc, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - cc); v = x + (pa <= pb && pa <= pc ? a : pb <= pc ? b : cc); }
        out[r * columns + c] = v & 255;
      }
    }
    return out;
  }

  private resolveNum(v: PdfValue | undefined): number | undefined {
    if (!v) return undefined;
    if (v.t === 'ref') { const o = this.obj(v.n); return o.t === 'num' ? Number(o.raw) : undefined; }
    return v.t === 'num' ? Number(v.raw) : undefined;
  }

  /** 객체 번호로 값을 가져온다(없으면 null). */
  obj(n: number): PdfValue {
    const hit = this.cache.get(n);
    if (hit) return hit;
    const e = this.xref.get(n);
    let v: PdfValue = { t: 'null' };
    if (e && e.type === 1 && e.a >= 0) {
      const lx = new Lexer(this.bytes, e.a);
      const num = Number(lx.word()); lx.word();
      if (lx.word() !== 'obj' || num !== n) throw new PdfError(`객체 ${n}의 위치가 올바르지 않습니다`);
      v = lx.value();
    } else if (e && e.type === 2) {
      let table = this.objStreams.get(e.a);
      if (!table) {
        const holder = this.obj(e.a);
        if (holder.t !== 'dict' || !holder.stream) throw new PdfError('객체 스트림이 올바르지 않습니다');
        const data = this.streamData(holder);
        const nObj = this.resolveNum(holder.v.get('N')) ?? 0, first = this.resolveNum(holder.v.get('First')) ?? 0;
        const hl = new Lexer(data, 0);
        const pairs: [number, number][] = [];
        for (let i = 0; i < nObj; i++) pairs.push([Number(hl.word()), Number(hl.word())]);
        table = new Map();
        for (const [on, oo] of pairs) table.set(on, new Lexer(data, first + oo).value());
        this.objStreams.set(e.a, table);
      }
      v = table.get(n) ?? { t: 'null' };
    }
    this.cache.set(n, v);
    return v;
  }

  resolve(v: PdfValue | undefined): PdfValue { return !v ? { t: 'null' } : v.t === 'ref' ? this.obj(v.n) : v; }
  private dictOf(v: PdfValue | undefined): Map<string, PdfValue> { const r = this.resolve(v); return r.t === 'dict' ? r.v : new Map(); }

  get version(): string { return String.fromCharCode(...this.bytes.subarray(5, 8)); }

  pages(): PageInfo[] {
    if (this.pageList) return this.pageList;
    const out: PageInfo[] = [];
    const seen = new Set<number>();
    const walk = (ref: PdfValue, inherited: Map<string, PdfValue>): void => {
      const n = ref.t === 'ref' ? ref.n : -1;
      if (n >= 0) { if (seen.has(n)) return; seen.add(n); }
      const d = this.dictOf(ref);
      const inh = new Map(inherited);
      for (const k of ['MediaBox', 'Rotate', 'Resources', 'CropBox']) { const x = d.get(k); if (x) inh.set(k, x); }
      const kids = this.resolve(d.get('Kids'));
      if (kids.t === 'arr' && (!d.has('Type') || (d.get('Type') as { v: string }).v === 'Pages')) { for (const k of kids.v) walk(k, inh); return; }
      const mb = this.resolve(inh.get('MediaBox'));
      const box = mb.t === 'arr' ? mb.v.map(x => Number(this.resolve(x).t === 'num' ? (this.resolve(x) as { raw: string }).raw : 0)) : [0, 0, 612, 792];
      out.push({ index: out.length, objNum: n, width: Math.abs(box[2] - box[0]), height: Math.abs(box[3] - box[1]) });
    };
    walk(this.trailer.get('Root')!.t === 'ref' ? (this.dictOf(this.trailer.get('Root')).get('Pages') ?? { t: 'null' }) : { t: 'null' }, new Map());
    if (out.length === 0) throw new PdfError('PDF에 쪽이 없습니다');
    this.pageList = out;
    return out;
  }

  /** 모든 쪽의 텍스트 메모(/Text 주석) */
  notes(): TextNote[] {
    const out: TextNote[] = [];
    for (const p of this.pages()) {
      if (p.objNum < 0) continue;
      const annots = this.resolve(this.dictOf({ t: 'ref', n: p.objNum, g: 0 }).get('Annots'));
      if (annots.t !== 'arr') continue;
      for (const a of annots.v) {
        const d = this.dictOf(a);
        if ((d.get('Subtype') as { v?: string } | undefined)?.v !== 'Text') continue;
        const rect = this.resolve(d.get('Rect'));
        const r = rect.t === 'arr' ? rect.v.map(x => Number((this.resolve(x) as { raw?: string }).raw ?? 0)) : [0, 0, 0, 0];
        const c = this.resolve(d.get('Contents')), t = this.resolve(d.get('T'));
        out.push({ id: a.t === 'ref' ? a.n : -1, page: p.index, x: r[0], y: r[3], contents: c.t === 'str' ? decodeText(c.v) : '', author: t.t === 'str' ? decodeText(t.v) : undefined });
      }
    }
    return out;
  }

  /** 증분 갱신으로 텍스트 메모를 덧붙인 새 PDF 바이트. 원본 바이트는 앞부분에 그대로 남는다. */
  withNotes(added: { page: number; x: number; y: number; contents: string; author?: string }[]): Uint8Array {
    if (added.length === 0) return this.bytes;
    const pages = this.pages();
    const size0 = Math.max(Number(this.resolveNum(this.trailer.get('Size')) ?? 0), ...[...this.xref.keys()].map(k => k + 1));
    let next = size0;
    const parts: Uint8Array[] = [];
    let offset = this.bytes.length;
    const newObjs: { n: number; off: number }[] = [];
    const emit = (n: number, body: string): void => {
      const s = enc.encode(`${n} 0 obj\n${body}\nendobj\n`);
      newObjs.push({ n, off: offset });
      parts.push(s); offset += s.length;
    };
    const lead = this.bytes[this.bytes.length - 1] === 10 ? '' : '\n';
    if (lead) { const l = enc.encode(lead); parts.push(l); offset += l.length; }

    // 쪽별로 새 주석 객체를 만들고 쪽(또는 Annots 배열)을 갱신한다
    const byPage = new Map<number, typeof added>();
    for (const a of added) {
      if (!Number.isInteger(a.page) || a.page < 0 || a.page >= pages.length) throw new PdfError(`쪽 번호가 범위 밖입니다: ${a.page}`);
      if (pages[a.page].objNum < 0) throw new PdfError('쪽 객체를 찾을 수 없어 메모를 추가할 수 없습니다');
      if (!(Number.isFinite(a.x) && Number.isFinite(a.y))) throw new PdfError('메모 위치가 올바르지 않습니다');
      let l = byPage.get(a.page); if (!l) byPage.set(a.page, l = []); l.push(a);
    }
    for (const [pi, list] of byPage) {
      const page = pages[pi];
      const pd = this.dictOf({ t: 'ref', n: page.objNum, g: 0 });
      const newRefs: PdfValue[] = [];
      for (const a of list) {
        const n = next++;
        const x = Math.max(0, a.x), y = Math.max(24, a.y);
        const d = new Map<string, PdfValue>([
          ['Type', name('Annot')], ['Subtype', name('Text')],
          ['Rect', { t: 'arr', v: [num(x), num(y - 24), num(x + 24), num(y)] }],
          ['Contents', textString(a.contents)], ['Name', name('Comment')], ['F', num(28)],
          ['C', { t: 'arr', v: [num(1), num(0.9), num(0.2)] }], ['Open', { t: 'bool', v: false }],
          ['P', { t: 'ref', n: page.objNum, g: 0 }],
        ]);
        if (a.author) d.set('T', textString(a.author));
        emit(n, serialize({ t: 'dict', v: d }));
        newRefs.push({ t: 'ref', n, g: 0 });
      }
      const existing = pd.get('Annots');
      if (existing && existing.t === 'ref') { // 간접 배열: 그 배열 객체를 새 판으로 다시 쓴다
        const arr = this.obj(existing.n);
        if (arr.t !== 'arr') throw new PdfError('Annots 배열이 올바르지 않습니다');
        emit(existing.n, serialize({ t: 'arr', v: [...arr.v, ...newRefs] }));
      } else { // 직접 배열이거나 없음: 쪽 사전을 새 판으로 다시 쓴다
        const arr: PdfValue = existing && existing.t === 'arr' ? { t: 'arr', v: [...existing.v, ...newRefs] } : { t: 'arr', v: newRefs };
        const nd = new Map(pd); nd.set('Annots', arr);
        emit(page.objNum, serialize({ t: 'dict', v: nd }));
      }
    }

    // 새 xref 구간 + 트레일러
    const rootRef = this.trailer.get('Root')!;
    const extra = new Map<string, PdfValue>();
    for (const k of ['Info', 'ID']) { const v = this.trailer.get(k); if (v) extra.set(k, v); }
    newObjs.sort((a, b) => a.n - b.n);
    const xrefOff = offset;
    if (!this.usesXrefStream) {
      let t = 'xref\n';
      for (let i = 0; i < newObjs.length;) {
        let j = i; while (j + 1 < newObjs.length && newObjs[j + 1].n === newObjs[j].n + 1) j++;
        t += `${newObjs[i].n} ${j - i + 1}\n`;
        for (let k = i; k <= j; k++) t += `${String(newObjs[k].off).padStart(10, '0')} 00000 n \n`;
        i = j + 1;
      }
      const tr = new Map<string, PdfValue>([['Size', num(next)], ['Root', rootRef], ['Prev', num(this.firstXrefOffset)], ...extra]);
      t += `trailer\n${serialize({ t: 'dict', v: tr })}\nstartxref\n${xrefOff}\n%%EOF\n`;
      parts.push(enc.encode(t));
    } else {
      // xref 스트림(비압축)으로 갱신. 새 xref 스트림 자체도 항목으로 올린다.
      const xn = next++;
      const entries = [...newObjs, { n: xn, off: xrefOff }].sort((a, b) => a.n - b.n);
      const index: PdfValue[] = [];
      const data: number[] = [];
      for (let i = 0; i < entries.length;) {
        let j = i; while (j + 1 < entries.length && entries[j + 1].n === entries[j].n + 1) j++;
        index.push(num(entries[i].n), num(j - i + 1));
        for (let k = i; k <= j; k++) { const o = entries[k].off; data.push(1, (o >>> 24) & 255, (o >>> 16) & 255, (o >>> 8) & 255, o & 255, 0, 0); }
        i = j + 1;
      }
      const d = new Map<string, PdfValue>([
        ['Type', name('XRef')], ['Size', num(next)], ['W', { t: 'arr', v: [num(1), num(4), num(2)] }], ['Index', { t: 'arr', v: index }],
        ['Root', rootRef], ['Prev', num(this.firstXrefOffset)], ['Length', num(data.length)], ...extra,
      ]);
      parts.push(enc.encode(`${xn} 0 obj\n${serialize({ t: 'dict', v: d })}\nstream\n`), Uint8Array.from(data), enc.encode(`\nendstream\nendobj\nstartxref\n${xrefOff}\n%%EOF\n`));
    }
    const total = this.bytes.length + parts.reduce((s, p) => s + p.length, 0);
    const out = new Uint8Array(total);
    out.set(this.bytes, 0);
    let o = this.bytes.length;
    for (const p of parts) { out.set(p, o); o += p.length; }
    return out;
  }
}
