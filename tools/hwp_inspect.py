# HWP 5.x(OLE2) 시료의 헤더와 스트림 구성을 읽는다(읽기 전용, 공개 규격의 FileHeader 구조 기준).
# 사용: python -I -X utf8 tools/hwp_inspect.py fixtures/hwp [--md]
import os, struct, sys, zlib
import olefile

FLAGS = ['압축', '암호', '배포용', '스크립트 저장', 'DRM', 'XML 서식 문서', '문서 이력 관리', '전자 서명 정보', '공인 인증서 암호', '공인 인증서 DRM', 'CCL 문서']


def inspect(path):
    ole = olefile.OleFileIO(path)
    try:
        streams = ['/'.join(s) for s in ole.listdir()]
        hdr = ole.openstream('FileHeader').read()
        sig = hdr[:32].rstrip(b'\0').decode('ascii', 'replace')
        ver = struct.unpack('<I', hdr[32:36])[0]
        flags = struct.unpack('<I', hdr[36:40])[0]
        v = f'{(ver >> 24) & 0xff}.{(ver >> 16) & 0xff}.{(ver >> 8) & 0xff}.{ver & 0xff}'
        flagnames = [n for i, n in enumerate(FLAGS) if flags & (1 << i)]
        sections = sorted(s for s in streams if s.startswith('BodyText/Section'))
        bindata = [s for s in streams if s.startswith('BinData/')]
        compressed = bool(flags & 1)
        sec0 = None
        if sections:
            raw = ole.openstream(sections[0]).read()
            try:
                data = zlib.decompress(raw, -15) if compressed else raw
                # 레코드 헤더: 태그 10비트, 레벨 10비트, 크기 12비트
                tags = {}
                off = 0
                while off + 4 <= len(data):
                    h = struct.unpack('<I', data[off:off + 4])[0]
                    tag, size = h & 0x3ff, h >> 20
                    off += 4
                    if size == 0xfff:
                        size = struct.unpack('<I', data[off:off + 4])[0]; off += 4
                    tags[tag] = tags.get(tag, 0) + 1
                    off += size
                sec0 = {'bytes': len(data), 'records': sum(tags.values()), 'HWPTAG_PARA_TEXT(0x43=67)': tags.get(67, 0), 'HWPTAG_TABLE(0x4d=77)': tags.get(77, 0)}
            except Exception as e:
                sec0 = {'error': str(e)}
        return {'file': os.path.basename(path), 'bytes': os.path.getsize(path), 'signature': sig, 'version': v,
                'flags': flagnames, 'sections': len(sections), 'binData': len(bindata),
                'hasSummary': '\x05HwpSummaryInformation' in streams, 'hasPrvText': 'PrvText' in streams, 'hasPrvImage': 'PrvImage' in streams,
                'streamCount': len(streams), 'section0': sec0}
    finally:
        ole.close()


if __name__ == '__main__':
    d = sys.argv[1]
    rows = [inspect(os.path.join(d, f)) for f in sorted(os.listdir(d)) if f.endswith('.hwp')]
    if '--md' in sys.argv:
        print('| 파일 | 크기 | 버전 | 플래그 | 구역 | BinData | 요약정보 | 첫 구역 레코드 | 문단 텍스트 | 표 |')
        print('|---|---:|---|---|---:|---:|---|---:|---:|---:|')
        for r in rows:
            s = r['section0'] or {}
            print(f"| {r['file']} | {r['bytes']} | {r['version']} | {', '.join(r['flags']) or '-'} | {r['sections']} | {r['binData']} | {'예' if r['hasSummary'] else '**없음**'} | {s.get('records', '-')} | {s.get('HWPTAG_PARA_TEXT(0x43=67)', '-')} | {s.get('HWPTAG_TABLE(0x4d=77)', '-')} |")
    else:
        import json
        print(json.dumps(rows, ensure_ascii=False, indent=2))
