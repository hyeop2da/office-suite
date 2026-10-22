# 앱 아이콘(PNG)을 의존성 없이 만든다. 사용: python tools/make_icons.py web
import struct, zlib, sys, os

def png(size, path):
    s = size
    px = bytearray()
    bg = (11, 92, 173); fg = (255, 255, 255); line = (11, 92, 173)
    for y in range(s):
        px.append(0)
        for x in range(s):
            u, v = x / s, y / s
            c = bg
            # 문서 모양(흰 종이)
            if 0.26 <= u <= 0.74 and 0.18 <= v <= 0.82:
                c = fg
                # 접힌 모서리
                if u > 0.58 and v < 0.34 and (u - 0.58) + (0.34 - v) > 0.16: c = bg
                # 글줄
                for k in range(4):
                    vy = 0.42 + k * 0.095
                    if vy <= v <= vy + 0.03 and 0.33 <= u <= (0.67 if k < 3 else 0.55): c = line
            px.extend(c)
    raw = bytes(px)
    def chunk(t, d): return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    data = b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', s, s, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw, 9)) + chunk(b'IEND', b'')
    open(path, 'wb').write(data)

out = sys.argv[1]
for n in (180, 192, 512): png(n, os.path.join(out, f'icon-{n}.png'))
print('ok')
