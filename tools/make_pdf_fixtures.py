# 시험용 PDF를 만든다(직접 생성. Acrobat·Word가 만든 파일이 아님).
# classic.pdf: reportlab, xref 표 / links.pdf: 이미 /Annots가 있는 쪽 / xrefstream.pdf: 객체 스트림 + xref 스트림(손으로 조립)
# 사용: python tools/make_pdf_fixtures.py fixtures/pdf
import os, sys, zlib
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

out = sys.argv[1]
os.makedirs(out, exist_ok=True)

c = canvas.Canvas(f'{out}/classic.pdf', pagesize=A4, pageCompression=0)
c.setTitle('classic')
c.drawString(72, 760, 'Page one: Hello PDF')
c.showPage()
c.drawString(72, 760, 'Page two: second page')
c.showPage()
c.drawString(72, 760, 'Page three')
c.save()

c = canvas.Canvas(f'{out}/links.pdf', pagesize=A4, pageCompression=0)
c.drawString(72, 760, 'Page with an existing link annotation')
c.linkURL('https://example.com/', (72, 750, 300, 775), relative=0)
c.showPage()
c.drawString(72, 760, 'Second page')
c.save()

# 손으로 조립한 xref 스트림 PDF
objs = {}
objs[1] = b'<</Type/Catalog/Pages 2 0 R>>'
objs[2] = b'<</Type/Pages/Kids[3 0 R]/Count 1>>'
objs[3] = b'<</Type/Page/Parent 2 0 R/MediaBox[0 0 300 200]/Contents 4 0 R/Resources<</Font<</F1 5 0 R>>>>>>'
objs[5] = b'<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>'
content = b'BT /F1 18 Tf 20 100 Td (XRef stream page) Tj ET'
body = b'%PDF-1.5\n%\xe2\xe3\xcf\xd3\n'
off = {}
off[4] = len(body)
body += b'4 0 obj\n<</Length %d>>\nstream\n' % len(content) + content + b'\nendstream\nendobj\n'
# 객체 스트림(6): 1,2,3,5
order = [1, 2, 3, 5]
hdr, data, pos = b'', b'', 0
for n in order:
    hdr += b'%d %d ' % (n, pos)
    data += objs[n] + b'\n'
    pos = len(data)
raw = hdr + data
comp = zlib.compress(raw)
off[6] = len(body)
body += b'6 0 obj\n<</Type/ObjStm/N 4/First %d/Filter/FlateDecode/Length %d>>\nstream\n' % (len(hdr), len(comp)) + comp + b'\nendstream\nendobj\n'
off[7] = len(body)
rows = b'\x00\x00\x00\x00\x00\x00\x00'                  # 0: free
idx = {1: 0, 2: 1, 3: 2, 5: 3}
for n in range(1, 8):
    if n in idx: rows += b'\x02' + (6).to_bytes(4, 'big') + idx[n].to_bytes(2, 'big')
    else: rows += b'\x01' + off[n].to_bytes(4, 'big') + b'\x00\x00'
xr = b'7 0 obj\n<</Type/XRef/Size 8/W[1 4 2]/Root 1 0 R/Length %d>>\nstream\n' % len(rows) + rows + b'\nendstream\nendobj\n'
body += xr
body += b'startxref\n%d\n%%%%EOF\n' % off[7]
open(f'{out}/xrefstream.pdf', 'wb').write(body)
print('ok', sorted(os.listdir(out)))
