// 무손실 ZIP 읽기·쓰기. 바꾸지 않은 항목은 압축된 바이트를 그대로 복사한다.
// 항목 순서, 압축 방식, 플래그, 시각, 속성도 원본 그대로 둔다.
// 압축 해제·압축은 fflate(MIT)를 쓴다. Node와 브라우저에서 같은 코드가 돈다.
import { inflateSync, deflateSync } from 'fflate';

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_DESC = 0x08074b50;

export class ZipError extends Error {
  constructor(message: string) { super(message); this.name = 'ZipError'; }
}

export interface ZipEntry {
  name: string;
  /** 중앙 디렉터리 레코드 원본 바이트(이름·extra·주석 포함). */
  central: Uint8Array;
  /** 로컬 헤더 + 압축 데이터 + (있으면) 데이터 서술자. 바뀌지 않았을 때만 그대로 쓴다. */
  raw: Uint8Array;
  method: number;
  /** 바뀐 항목의 새 압축 해제 내용. 없으면 raw 재사용. */
  replaced?: Uint8Array;
}

export class ZipArchive {
  entries: ZipEntry[] = [];

  constructor(bytes: Uint8Array) {
    try {
      this.parse(bytes);
    } catch (e) {
      if (e instanceof ZipError) throw e;
      throw new ZipError(`ZIP 구조가 손상되었습니다: ${(e as Error).message}`);
    }
  }

  private parse(bytes: Uint8Array): void {
    if (bytes.length < 22) throw new ZipError('ZIP 파일이 너무 짧습니다');
    const dv = view(bytes);
    let eocd = -1;
    for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
      if (dv.getUint32(i, true) === SIG_EOCD) { eocd = i; break; }
    }
    if (eocd < 0) throw new ZipError('ZIP 끝 레코드(EOCD)를 찾지 못했습니다');
    const count = dv.getUint16(eocd + 10, true);
    let cd = dv.getUint32(eocd + 16, true);
    if (count === 0xffff || cd === 0xffffffff) throw new ZipError('ZIP64는 지원하지 않습니다');
    const seen = new Set<string>();
    for (let n = 0; n < count; n++) {
      if (cd + 46 > bytes.length || dv.getUint32(cd, true) !== SIG_CENTRAL) throw new ZipError('중앙 디렉터리 레코드 서명 오류');
      const flags = dv.getUint16(cd + 8, true);
      const method = dv.getUint16(cd + 10, true);
      const csize = dv.getUint32(cd + 20, true);
      const nlen = dv.getUint16(cd + 28, true);
      const elen = dv.getUint16(cd + 30, true);
      const clen = dv.getUint16(cd + 32, true);
      const lho = dv.getUint32(cd + 42, true);
      if (flags & 0x01) throw new ZipError('암호화된 ZIP 항목은 지원하지 않습니다');
      if (method !== 0 && method !== 8) throw new ZipError(`지원하지 않는 압축 방식입니다(${method})`);
      const name = new TextDecoder().decode(bytes.subarray(cd + 46, cd + 46 + nlen));
      if (seen.has(name)) throw new ZipError(`같은 이름의 항목이 둘 있습니다: ${name}`);
      seen.add(name);
      const recLen = 46 + nlen + elen + clen;
      if (lho + 30 > bytes.length || dv.getUint32(lho, true) !== SIG_LOCAL) throw new ZipError(`로컬 헤더 서명 오류: ${name}`);
      const lnlen = dv.getUint16(lho + 26, true);
      const lelen = dv.getUint16(lho + 28, true);
      let end = lho + 30 + lnlen + lelen + csize;
      if (end > bytes.length) throw new ZipError(`항목 데이터가 파일 끝을 넘습니다: ${name}`);
      if (flags & 0x08) end += end + 4 <= bytes.length && dv.getUint32(end, true) === SIG_DESC ? 16 : 12;
      this.entries.push({
        name, method,
        central: bytes.slice(cd, cd + recLen),
        raw: bytes.subarray(lho, Math.min(end, bytes.length)),
      });
      cd += recLen;
    }
  }

  has(name: string): boolean { return this.entries.some(e => e.name === name); }

  private find(name: string): ZipEntry {
    const e = this.entries.find(x => x.name === name);
    if (!e) throw new ZipError(`ZIP 항목 없음: ${name}`);
    return e;
  }

  read(name: string): Uint8Array {
    const e = this.find(name);
    if (e.replaced) return e.replaced;
    const dv = view(e.raw);
    const csize = view(e.central).getUint32(20, true);
    const usize = view(e.central).getUint32(24, true);
    const start = 30 + dv.getUint16(26, true) + dv.getUint16(28, true);
    const data = e.raw.subarray(start, start + csize);
    if (e.method === 0) return data.slice();
    try {
      const out = inflateSync(data);
      if (out.length !== usize) throw new ZipError(`압축 해제 크기가 다릅니다: ${name}`);
      return out;
    } catch (err) {
      if (err instanceof ZipError) throw err;
      throw new ZipError(`항목 압축 해제 실패(손상): ${name}`);
    }
  }

  /** UTF-8 글자로 읽는다. BOM은 지우지 않고 그대로 남긴다(다시 쓸 때 원형을 지키기 위해). */
  readText(name: string): string { return new TextDecoder('utf-8', { ignoreBOM: true }).decode(this.read(name)); }

  /** 항목 내용을 바꾼다. 위치·압축 방식은 유지한다. */
  replace(name: string, data: Uint8Array): void { this.find(name).replaced = data; }

  isModified(name: string): boolean { return !!this.find(name).replaced; }

  /** ZIP 바이트를 쓴다. */
  write(): Uint8Array {
    const parts: Uint8Array[] = [];
    const centrals: Uint8Array[] = [];
    let offset = 0;
    for (const e of this.entries) {
      let local: Uint8Array;
      let central = e.central.slice();
      if (!e.replaced) {
        local = e.raw;
      } else {
        const data = e.replaced;
        const comp = e.method === 0 ? data : deflateSync(data, { level: 9 });
        const crc = crc32(data);
        const cdv = view(e.central);
        const nameBytes = e.central.subarray(46, 46 + cdv.getUint16(28, true));
        const flags = cdv.getUint16(8, true) & ~0x08; // 서술자를 쓰지 않는다
        const lh = new Uint8Array(30 + nameBytes.length);
        const ldv = view(lh);
        ldv.setUint32(0, SIG_LOCAL, true);
        ldv.setUint16(4, cdv.getUint16(6, true), true);
        ldv.setUint16(6, flags, true);
        ldv.setUint16(8, e.method, true);
        ldv.setUint16(10, cdv.getUint16(12, true), true);
        ldv.setUint16(12, cdv.getUint16(14, true), true);
        ldv.setUint32(14, crc, true);
        ldv.setUint32(18, comp.length, true);
        ldv.setUint32(22, data.length, true);
        ldv.setUint16(26, nameBytes.length, true);
        ldv.setUint16(28, 0, true);
        lh.set(nameBytes, 30);
        local = concat([lh, comp]);
        const c = view(central);
        c.setUint16(8, flags, true);
        c.setUint32(16, crc, true);
        c.setUint32(20, comp.length, true);
        c.setUint32(24, data.length, true);
      }
      view(central).setUint32(42, offset, true);
      parts.push(local);
      centrals.push(central);
      offset += local.length;
    }
    const cdStart = offset;
    const cdBytes = concat(centrals);
    const eocd = new Uint8Array(22);
    const edv = view(eocd);
    edv.setUint32(0, SIG_EOCD, true);
    edv.setUint16(8, this.entries.length, true);
    edv.setUint16(10, this.entries.length, true);
    edv.setUint32(12, cdBytes.length, true);
    edv.setUint32(16, cdStart, true);
    return concat([...parts, cdBytes, eocd]);
  }
}

function view(b: Uint8Array): DataView { return new DataView(b.buffer, b.byteOffset, b.byteLength); }

function concat(list: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(list.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of list) { out.set(p, o); o += p.length; }
  return out;
}

let table: Uint32Array | undefined;
export function crc32(data: Uint8Array): number {
  if (!table) {
    table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c >>> 0;
    }
  }
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i++) crc = table[(crc ^ data[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}
