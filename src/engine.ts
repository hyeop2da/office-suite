// 형식 판별과 공통 문서 엔진 인터페이스. UI는 이 파일만 알면 된다.
import { HwpxDocument } from './hwpx.ts';
import { DocxDocument } from './docx.ts';
import { PptxDocument } from './pptx.ts';
import { ZipArchive, ZipError } from './zip.ts';
import type { EditOptions } from './text-part.ts';

export type FormatId = 'hwpx' | 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'hwp' | 'doc' | 'xls' | 'ppt' | 'drm' | 'unknown';

export type SupportLevel = 'text-edit' | 'cell-edit' | 'annotate' | 'view' | 'none';

export interface FormatInfo {
  id: FormatId;
  label: string;
  support: SupportLevel;
  /** 사용자에게 보여 줄 지원 범위 설명 */
  note: string;
}

export const FORMATS: Record<FormatId, FormatInfo> = {
  hwpx: { id: 'hwpx', label: 'HWPX', support: 'text-edit', note: '글자 편집·저장(원본 구조 보존). 표 셀·머리말 글자 포함, 탭·줄바꿈(Shift+Enter) 삽입 가능. 서식 편집·문단 나누기는 아직 화면에 없음' },
  docx: { id: 'docx', label: 'DOCX', support: 'text-edit', note: '글자 편집·저장(원본 구조 보존). 서식·표 구조 편집 없음' },
  xlsx: { id: 'xlsx', label: 'XLSX', support: 'cell-edit', note: '칸 값 편집·저장(숫자는 숫자로 유지, 원본 구조 보존). 수식·서식·차트 편집과 수식 재계산은 없습니다' },
  pptx: { id: 'pptx', label: 'PPTX', support: 'text-edit', note: '슬라이드·노트 글자 편집·저장(원본 구조 보존). 필드(슬라이드 번호 등)·도형·서식·애니메이션 편집은 없습니다' },
  pdf: { id: 'pdf', label: 'PDF', support: 'annotate', note: '열람·쪽 이동·확대와 텍스트 메모 추가(원본 바이트를 그대로 두고 끝에 덧붙여 저장). 글자·그림 편집, 암호화된 PDF는 지원하지 않습니다' },
  hwp: { id: 'hwp', label: 'HWP(구형 바이너리)', support: 'view', note: 'HWP 5.x 바이너리는 쪽 미리보기(읽기 전용)만 지원합니다. 편집·저장은 되지 않습니다. HWPX와 다른 형식입니다' },
  doc: { id: 'doc', label: 'DOC(구형 바이너리)', support: 'none', note: '구형 바이너리는 별도 평가 후 지원 예정' },
  xls: { id: 'xls', label: 'XLS(구형 바이너리)', support: 'none', note: '구형 바이너리는 별도 평가 후 지원 예정' },
  ppt: { id: 'ppt', label: 'PPT(구형 바이너리)', support: 'none', note: '구형 바이너리는 별도 평가 후 지원 예정' },
  drm: { id: 'drm', label: '문서보안(DRM) 암호화 파일', support: 'none', note: 'Fasoo DRM 등으로 암호화된 문서는 열 수 없습니다. 이 앱은 문서보안을 우회하지 않습니다. 해제가 필요하면 기관의 승인된 절차를 따르세요' },
  unknown: { id: 'unknown', label: '알 수 없음', support: 'none', note: '지원하는 형식이 아닙니다' },
};

/** 파일 내용으로 형식을 판별한다(확장자는 쓰지 않는다). */
export function detectFormat(bytes: Uint8Array, filename = ''): FormatInfo {
  const head = String.fromCharCode(...bytes.subarray(0, 5));
  // Fasoo DRM이 저장 시 씌우는 머리글(" DRMONE  This Document is encrypted and protected by Fasoo DRM")
  if (/DRMONE|Fasoo DRM/.test(String.fromCharCode(...bytes.subarray(0, 96)))) return FORMATS.drm;
  if (head === '%PDF-') return FORMATS.pdf;
  const isZip = bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && (bytes[2] === 3 || bytes[2] === 5);
  if (isZip) {
    try {
      const zip = new ZipArchive(bytes);
      if (zip.has('mimetype') && zip.readText('mimetype').trim() === 'application/hwp+zip') return FORMATS.hwpx;
      if (zip.has('word/document.xml')) return FORMATS.docx;
      if (zip.has('xl/workbook.xml')) return FORMATS.xlsx;
      if (zip.has('ppt/presentation.xml')) return FORMATS.pptx;
    } catch (e) {
      if (e instanceof ZipError) return FORMATS.unknown;
      throw e;
    }
    return FORMATS.unknown;
  }
  if (bytes.length >= 8 && bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) {
    // OLE2: 확장자로만 구분한다(내부 스트림 해석은 아직 하지 않는다).
    const ext = filename.toLowerCase().split('.').pop();
    if (ext === 'hwp') return FORMATS.hwp;
    if (ext === 'doc') return FORMATS.doc;
    if (ext === 'xls') return FORMATS.xls;
    if (ext === 'ppt') return FORMATS.ppt;
    return FORMATS.unknown;
  }
  if (head.startsWith('HWP D') || String.fromCharCode(...bytes.subarray(0, 17)) === 'HWP Document File') return FORMATS.hwp;
  return FORMATS.unknown;
}

export interface DocParagraph {
  /** 같은 문서 안에서 바뀌지 않는 식별자: "부품|문단번호" */
  id: string;
  /** 화면에 보일 묶음 이름(예: PPTX의 "슬라이드 3") */
  group?: string;
  part: string;
  context: string;
  text: string;
  partlyLocked: boolean;
}

export interface DocumentEngine {
  readonly format: FormatInfo;
  /** 문단 글자 안에 탭·줄바꿈을 그대로 넣고 뺄 수 있으면 true(지원 형식만). 거짓이면 UI가 \t\n을 막는다. */
  readonly supportsTabBreak: boolean;
  paragraphs(): DocParagraph[];
  setParagraphText(id: string, text: string, opt?: EditOptions): void;
  save(): Uint8Array;
}

function parseId(id: string): [string, number] {
  const k = id.lastIndexOf('|');
  return [id.slice(0, k), Number(id.slice(k + 1))];
}

class HwpxEngine implements DocumentEngine {
  readonly format = FORMATS.hwpx;
  readonly supportsTabBreak = true;
  private doc: HwpxDocument;
  constructor(bytes: Uint8Array) { this.doc = HwpxDocument.open(bytes); }
  paragraphs(): DocParagraph[] {
    // richParagraphText는 탭·줄바꿈을 \t·\n으로 보여준다(색인 범위만 확인하고, 다른 요소가 섞여도
    // 던지지 않는다 — hasOtherTags 검사는 setParagraphRichText 쓰기 쪽에만 있다. 읽기는 늘 안전하다).
    return this.doc.paragraphs().map(p => ({
      id: `${p.section}|${p.index}`, part: p.section, context: p.context,
      text: this.doc.richParagraphText(p.section, p.index), partlyLocked: p.partlyLocked,
    }));
  }
  setParagraphText(id: string, text: string, opt?: EditOptions): void {
    const [part, i] = parseId(id);
    this.doc.setParagraphRichText(part, i, text, opt);
  }
  save(): Uint8Array { return this.doc.save(); }
}

class DocxEngine implements DocumentEngine {
  readonly format = FORMATS.docx;
  readonly supportsTabBreak = false;
  private doc: DocxDocument;
  constructor(bytes: Uint8Array) { this.doc = DocxDocument.open(bytes); }
  paragraphs(): DocParagraph[] {
    return this.doc.paragraphs().map(p => ({ id: `${p.section}|${p.index}`, part: p.section, context: p.context, text: p.text, partlyLocked: p.partlyLocked }));
  }
  setParagraphText(id: string, text: string, opt?: EditOptions): void {
    const [part, i] = parseId(id);
    this.doc.setParagraphText(part, i, text, opt);
  }
  save(): Uint8Array { return this.doc.save(); }
}

export class UnsupportedFormatError extends Error {
  readonly info: FormatInfo;
  constructor(info: FormatInfo) { super(`${info.label}: ${info.note}`); this.name = 'UnsupportedFormatError'; this.info = info; }
}

class PptxEngine implements DocumentEngine {
  readonly format = FORMATS.pptx;
  readonly supportsTabBreak = false;
  private doc: PptxDocument;
  constructor(bytes: Uint8Array) { this.doc = PptxDocument.open(bytes); }
  paragraphs(): DocParagraph[] {
    return this.doc.paragraphs().map(p => ({ id: `${p.section}|${p.index}`, group: p.group, part: p.section, context: p.context, text: p.text, partlyLocked: p.partlyLocked }));
  }
  setParagraphText(id: string, text: string, opt?: EditOptions): void {
    const [part, i] = parseId(id);
    this.doc.setParagraphText(part, i, text, opt);
  }
  save(): Uint8Array { return this.doc.save(); }
}

export function openDocument(bytes: Uint8Array, filename = ''): DocumentEngine {
  const info = detectFormat(bytes, filename);
  if (info.id === 'hwpx') return new HwpxEngine(bytes);
  if (info.id === 'docx') return new DocxEngine(bytes);
  if (info.id === 'pptx') return new PptxEngine(bytes);
  throw new UnsupportedFormatError(info);
}
