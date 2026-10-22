// XML 부품 하나(HWPX 구역, DOCX 본문 등)의 글자를 무손실로 고치는 공통 편집기.
import { ZipArchive } from './zip.ts';
import { scanXml, encodeText, FormatError, type PartConfig, type ScanResult, type TextNode } from './xml-scan.ts';

export class EditError extends Error {
  constructor(message: string) { super(message); this.name = 'EditError'; }
}

export interface EditOptions {
  /**
   * 글자가 바뀐 문단의 줄 배치 캐시(hp:linesegarray)를 어떻게 할지.
   * 'drop'(기본)은 바뀐 문단의 캐시만 지워서 한/글이 다시 계산하게 한다.
   * 'keep'은 원본을 그대로 둔다. 한/글 2022 실측(docs/HWPX-검증결과.md, E3): 줄 수가 늘어나는 편집에서
   * 'keep'은 글자가 겹쳐 그려지고 'drop'은 정상 재조판되었다. 'keep'은 시험·비교용이다.
   */
  layout?: 'keep' | 'drop';
}

export interface ParagraphView {
  /** 부품 안 문단 번호 */
  index: number;
  text: string;
  context: string;
  /** 하위 요소가 있는 글자 조각이 섞여 있어 일부만 편집할 수 있으면 true */
  partlyLocked: boolean;
  /** 글자 조각 수 */
  runs: number;
}

/**
 * XML에 쓸 수 없는 글자와 이 편집기가 아직 다루지 못하는 글자를 거른다.
 * allowTab: PPTX처럼 전용 탭 요소가 없어 탭 문자(U+0009)를 글자 내용에 그대로 쓰는 형식(xml-scan.ts의
 * PartConfig.allowLiteralTab 참고). HWPX·DOCX는 전용 요소(hp:tab/w:tab)가 있어 기본값(거부)을 쓴다.
 */
export function validateText(text: string, opt: { allowTab?: boolean } = {}): void {
  if (typeof text !== 'string') throw new EditError('글자는 문자열이어야 합니다');
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c === 0x0a || c === 0x0d) throw new EditError('줄바꿈은 글자 편집으로 넣을 수 없습니다(문단 나누기는 아직 지원하지 않음)');
    if (c === 0x09) {
      if (!opt.allowTab) throw new EditError('탭은 글자 편집으로 넣을 수 없습니다(형식별 탭 요소가 필요함)');
    } else if (c < 0x20 || c === 0xfffe || c === 0xffff) throw new EditError(`XML에 쓸 수 없는 제어 문자입니다(U+${c.toString(16).padStart(4, '0')})`);
    if (c >= 0xd800 && c <= 0xdbff) {
      const d = text.charCodeAt(i + 1);
      if (!(d >= 0xdc00 && d <= 0xdfff)) throw new EditError('짝이 맞지 않는 서로게이트 글자입니다');
      i++;
    } else if (c >= 0xdc00 && c <= 0xdfff) throw new EditError('짝이 맞지 않는 서로게이트 글자입니다');
  }
}

interface NodeEdit { node: number; text: string }

export class TextPart {
  private cache: string | undefined;
  private scanned: ScanResult | undefined;

  readonly zip: ZipArchive;
  readonly name: string;
  readonly cfg: PartConfig;

  constructor(zip: ZipArchive, name: string, cfg: PartConfig) {
    this.zip = zip;
    this.name = name;
    this.cfg = cfg;
  }

  xml(): string {
    if (this.cache === undefined) this.cache = this.zip.readText(this.name);
    return this.cache;
  }

  scan(): ScanResult {
    if (!this.scanned) this.scanned = scanXml(this.xml(), this.cfg);
    return this.scanned;
  }

  private commit(xml: string): void {
    this.cache = xml;
    this.scanned = undefined;
    this.zip.replace(this.name, new TextEncoder().encode(xml));
  }

  /** 고칠 수 없는 글자 조각: 하위 요소(탭 등)가 있거나 필드처럼 프로그램이 채우는 글자 */
  isLocked(n: TextNode): boolean { return n.hasChildren || !!this.cfg.lockedContexts?.includes(n.context); }

  plainText(): string { return this.scan().texts.map(t => t.text).join(''); }

  /** 글자 조각들을 한 번에 바꾼다. 같은 스캔 결과 기준이므로 위치가 어긋나지 않는다. */
  applyNodeEdits(edits: NodeEdit[], opt: EditOptions = {}): void {
    const scan = this.scan();
    let xml = this.xml();
    const ops: [number, number, string][] = [];
    const dropParas = new Set<number>();
    const seen = new Set<number>();
    for (const e of edits) {
      validateText(e.text, { allowTab: this.cfg.allowLiteralTab });
      if (!Number.isInteger(e.node) || e.node < 0 || e.node >= scan.texts.length) throw new EditError(`글자 번호 범위 밖: ${e.node}`);
      if (seen.has(e.node)) throw new EditError(`같은 글자 조각을 두 번 고칠 수 없습니다: ${e.node}`);
      seen.add(e.node);
      const n = scan.texts[e.node];
      if (this.isLocked(n)) throw new EditError('하위 요소(탭 등)가 있는 글자 조각은 고칠 수 없습니다');
      if (n.context === 'fallback') throw new EditError('호환용 대체 내용(mc:Fallback)은 고칠 수 없습니다');
      if (e.text === n.text) continue;
      const needPreserve = this.cfg.preserveSpaceAttr && /^\s|\s$|\s\s/.test(e.text);
      if (n.selfClosing) {
        let open = xml.slice(n.tagStart, n.tagEnd).replace(/\s*\/>$/, '>');
        if (needPreserve && !/xml:space=/.test(open)) open = open.replace(/>$/, ' xml:space="preserve">');
        ops.push([n.tagStart, n.tagEnd, `${open}${encodeText(e.text)}</${this.cfg.text}>`]);
      } else {
        if (needPreserve && !/xml:space=/.test(xml.slice(n.tagStart, n.tagEnd))) {
          ops.push([n.tagStart, n.tagEnd, xml.slice(n.tagStart, n.tagEnd).replace(/>$/, ' xml:space="preserve">')]);
        }
        ops.push([n.contentStart, n.contentEnd, encodeText(e.text)]);
      }
      if (n.paraIndex >= 0) dropParas.add(n.paraIndex);
    }
    if ((opt.layout ?? 'drop') === 'drop') {
      for (const p of dropParas) {
        const seg = scan.paras[p].lineSegs;
        if (seg) ops.push([seg[0], seg[1], '']);
      }
    }
    if (!ops.length) return;
    ops.sort((a, b) => b[0] - a[0]);
    for (const [s, e, r] of ops) xml = xml.slice(0, s) + r + xml.slice(e);
    this.commit(xml);
  }

  setText(nodeIndex: number, text: string, opt: EditOptions = {}): void {
    this.applyNodeEdits([{ node: nodeIndex, text }], opt);
  }

  /** 글자가 있는 문단 목록(호환용 대체 내용은 뺀다). */
  paragraphs(): ParagraphView[] {
    const scan = this.scan();
    const byPara = new Map<number, TextNode[]>();
    for (const t of scan.texts) {
      if (t.paraIndex < 0 || t.context === 'fallback') continue;
      let list = byPara.get(t.paraIndex);
      if (!list) byPara.set(t.paraIndex, list = []);
      list.push(t);
    }
    const out: ParagraphView[] = [];
    for (const [index, nodes] of byPara) {
      const text = nodes.map(n => n.text).join('');
      if (text.length === 0) continue;
      out.push({ index, text, context: nodes[0].context, partlyLocked: nodes.some(n => this.isLocked(n)), runs: nodes.length });
    }
    return out.sort((a, b) => a.index - b.index);
  }

  /**
   * 문단 전체 글자를 새 값으로 바꾼다. 앞뒤 공통 부분은 그대로 두고 달라진 구간만 해당 글자 조각에서 고친다.
   * 글자 모양이 다른 조각이 여러 개면 삽입은 앞 조각의 모양을 따른다.
   */
  setParagraphText(paraIndex: number, newText: string, opt: EditOptions = {}): void {
    if (typeof newText !== 'string') throw new EditError('글자는 문자열이어야 합니다');
    const scan = this.scan();
    if (!Number.isInteger(paraIndex) || paraIndex < 0 || paraIndex >= scan.paras.length) throw new EditError(`문단 번호 범위 밖: ${paraIndex}`);
    const nodes = scan.texts.filter(t => t.paraIndex === paraIndex && t.context !== 'fallback');
    const old = nodes.map(n => n.text).join('');
    if (old === newText) return;
    const cp = (s: string) => Array.from(s);
    const a = cp(old), b = cp(newText);
    let pre = 0;
    while (pre < a.length && pre < b.length && a[pre] === b[pre]) pre++;
    let suf = 0;
    while (suf < a.length - pre && suf < b.length - pre && a[a.length - 1 - suf] === b[b.length - 1 - suf]) suf++;
    const delStart = pre, delEnd = a.length - suf; // [delStart, delEnd) 지움(코드포인트 위치)
    const insert = b.slice(pre, b.length - suf).join('');
    // 전체 newText가 아니라 "진짜 새로 들어오는 부분"만 검사한다 — 탭·줄바꿈이 이미 들어 있던(마커로
    // 표시된) 문단을 다른 자리에서 고칠 때, 그 기존 글자가 newText 전체에도 그대로 남아 있다는 이유로
    // 다시 거부되면 안 되기 때문이다(이미 한 번 유효했던 내용은 재검사하지 않음).
    validateText(insert, { allowTab: this.cfg.allowLiteralTab });
    // 글자 조각별 코드포인트 구간
    const spans: { n: TextNode; s: number; e: number; chars: string[] }[] = [];
    let pos = 0;
    for (const n of nodes) { const chars = cp(n.text); spans.push({ n, s: pos, e: pos + chars.length, chars }); pos += chars.length; }
    const result = new Map<number, string>();
    const touched = spans.filter(sp => sp.s < delEnd && sp.e > delStart);
    for (const sp of touched) {
      if (this.isLocked(sp.n)) throw new EditError('하위 요소(탭 등)가 있는 글자 조각이 바뀌는 범위에 걸렸습니다');
    }
    // 삽입 위치: 지우는 구간이 있으면 그 첫 조각, 없으면 delStart를 포함하는 조각(경계면 앞 조각)
    let target = touched[0];
    if (!target) {
      target = spans.find(sp => !this.isLocked(sp.n) && delStart > sp.s && delStart <= sp.e)
        ?? spans.find(sp => !this.isLocked(sp.n) && sp.s === delStart);
    }
    if (!target) throw new EditError('글자를 넣을 수 있는 조각이 없습니다');
    for (const sp of spans) {
      if (this.isLocked(sp.n)) continue;
      const del = sp.s < delEnd && sp.e > delStart;
      if (!del && sp !== target) continue;
      let chars = sp.chars.slice();
      let local = 0; // 이 조각 안에서 지운 개수
      if (del) {
        const from = Math.max(delStart, sp.s) - sp.s, to = Math.min(delEnd, sp.e) - sp.s;
        chars.splice(from, to - from);
        local = from;
      } else {
        local = delStart - sp.s;
      }
      if (sp === target) chars.splice(local, 0, ...cp(insert));
      result.set(sp.n.index, chars.join(''));
    }
    this.applyNodeEdits([...result].map(([node, text]) => ({ node, text })), opt);
  }

  /**
   * 호출자가 직접 계산한 [시작, 끝, 대체문자열] 구간들을 한 번에 적용하는 저수준 API.
   * 탭·줄바꿈 삽입, 문단 나누기처럼 단일 글자 조각을 넘어서는 구조 편집에 쓴다(HWPX·DOCX 코어 전용).
   * 검증은 호출자 책임이다(이 메서드 자체는 위치만 믿고 그대로 적용한다).
   */
  applyRawEdits(ops: [number, number, string][], touchedParas: Iterable<number> = [], opt: EditOptions = {}): void {
    if (!ops.length) return;
    const scan = this.scan();
    let xml = this.xml();
    const all = ops.slice();
    if ((opt.layout ?? 'drop') === 'drop') {
      for (const p of touchedParas) {
        const seg = scan.paras[p]?.lineSegs;
        if (seg) all.push([seg[0], seg[1], '']);
      }
    }
    all.sort((a, b) => b[0] - a[0]);
    for (const [s, e, r] of all) xml = xml.slice(0, s) + r + xml.slice(e);
    this.commit(xml);
  }
}

export { FormatError };
