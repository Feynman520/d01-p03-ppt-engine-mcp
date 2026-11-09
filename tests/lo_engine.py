"""
macOS 엔진 검증(LibreOffice) — engine/lo_bridge.py + engine/lo_presentation.py 를 진짜 soffice 로 돌린다.

2장짜리 프레젠테이션을 flat ODP(XML)로 직접 써서 soffice 로 pptx 를 만든 뒤, PDF·슬라이드 이미지를 낸다.
실행: python tests/lo_engine.py      (LibreOffice 가 없으면 SKIP)
"""
import os
import shutil
import struct
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from engine.lo_bridge import LoSession, find_soffice  # noqa: E402
from engine import lo_presentation as pres  # noqa: E402

FODP = """<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.presentation">
 <office:automatic-styles><style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="25.4cm" fo:page-height="14.29cm"/></style:page-layout></office:automatic-styles>
 <office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles>
 <office:body><office:presentation>
  <draw:page draw:name="p1" draw:master-page-name="Default"><draw:frame svg:x="2cm" svg:y="2cm" svg:width="20cm" svg:height="3cm"><draw:text-box><text:p>첫 슬라이드 IRIS</text:p></draw:text-box></draw:frame></draw:page>
  <draw:page draw:name="p2" draw:master-page-name="Default"><draw:frame svg:x="2cm" svg:y="2cm" svg:width="20cm" svg:height="3cm"><draw:text-box><text:p>둘째 슬라이드</text:p></draw:text-box></draw:frame></draw:page>
 </office:presentation></office:body>
</office:document>"""


def png_size(path):
    with open(path, "rb") as f:
        head = f.read(24)
    assert head[:8] == bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]), "not a PNG"
    return struct.unpack(">II", head[16:24])


def main():
    soffice = find_soffice()
    if not soffice:
        print("SKIP: LibreOffice 없음")
        return 0
    tmp = tempfile.mkdtemp(prefix="iris-lo-test-")
    fails = 0

    def check(cond, name):
        nonlocal fails
        print(("PASS" if cond else "FAIL"), name, flush=True)
        if not cond:
            fails += 1

    try:
        fodp = os.path.join(tmp, "deck.fodp")
        with open(fodp, "w", encoding="utf-8") as f:
            f.write(FODP)
        subprocess.run([soffice, "-env:UserInstallation=file://" + os.path.join(tmp, "conv-profile"), "--headless", "--norestore",
                        "--convert-to", "pptx", "--outdir", tmp, fodp], capture_output=True, timeout=180)
        src = os.path.join(tmp, "deck.pptx")
        check(os.path.exists(src), "시험용 pptx 만들기")
        s = LoSession(profile_dir=os.path.join(tmp, "profile"))
        s.start()
        pdf = s.run(lambda a: pres.export_pdf(a, src, os.path.join(tmp, "out", "deck.pdf")))
        with open(pdf, "rb") as f:
            data = f.read()
        check(data[:5] == b"%PDF-", "ppt_export_pdf → PDF")
        # 화면 없는 LibreOffice 가 시스템 한글 글꼴을 보는지(못 보면 한글이 □ 로 나오고 한글 글꼴이 박히지 않는다)
        check(any(k in data for k in (b"AppleSDGothic", b"Nanum", b"AppleMyungjo", b"AppleGothic")),
              "PDF 에 한글 글꼴이 박혔다(시스템 글꼴을 본다 — fonts.conf)")
        r = s.run(lambda a: pres.export_images(a, src, os.path.join(tmp, "imgs")))
        names = [os.path.basename(p) for p in r["files"]]
        check(r["count"] == 2 and names == ["slide_001.png", "slide_002.png"], f"슬라이드마다 PNG 한 장, PowerPoint 와 같은 이름 ({names})")
        w, h = png_size(r["files"][0])
        import re
        box = re.search(rb"/MediaBox\s*\[\s*0\s+0\s+([\d.]+)\s+([\d.]+)", data)
        want = (round(float(box.group(1)) / 72 * 96), round(float(box.group(2)) / 72 * 96)) if box else None
        check(want is not None and abs(w - want[0]) <= 1 and abs(h - want[1]) <= 1,
              f"크기 미지정 → 슬라이드 크기 96dpi(PDF 쪽 크기 기준 {want}) ({w}×{h})")
        r2 = s.run(lambda a: pres.export_images(a, src, os.path.join(tmp, "imgs2"), fmt="png", width=320, height=180))
        check(png_size(r2["files"][1]) == (320, 180), "width·height 둘 다 주면 그 크기")
        r3 = s.run(lambda a: pres.export_images(a, src, os.path.join(tmp, "imgs3"), fmt="JPG"))
        with open(r3["files"][0], "rb") as f:
            check(f.read(3) == b"\xff\xd8\xff" and r3["files"][0].endswith(".jpg"), "JPG 형식")
        try:
            s.run(lambda a: pres.export_images(a, src, tmp, fmt="WEBP"))
            check(False, "모르는 형식은 거부")
        except ValueError:
            check(True, "모르는 형식은 거부")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print(f"---- FAILS: {fails}")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
