"""
스모크 테스트 — 본 머신에서 PowerPoint COM 엔진 자동화가 실제로 동작하는지 검증한다.
MCP 계층 없이 engine(session+presentation)을 직접 호출한다.

실행:  .venv\\Scripts\\python.exe tests\\smoke_com.py

통과 기준: 예외/팝업 없이 끝까지 진행하고 deck.pptx / out.pdf / 슬라이드 PNG 2장이 생성된다.
"""
import os
import sys
import traceback

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from engine import presentation  # noqa: E402
from engine.session import PowerPointSession  # noqa: E402

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_out")
os.makedirs(OUT_DIR, exist_ok=True)


def log(msg):
    print(f"[smoke] {msg}", flush=True)


def _make_deck(app, path):
    """2장짜리 입력 프레젠테이션을 COM으로 생성한다(python-pptx가 만든 입력 대역)."""
    pres = app.Presentations.Add(WithWindow=0)  # msoFalse
    for i, title in enumerate(("엔진 렌더 테스트", "둘째 슬라이드"), start=1):
        layout = 12 if i == 1 else 12  # ppLayoutBlank 계열
        slide = pres.Slides.Add(i, layout)
        box = slide.Shapes.AddTextbox(1, 60, 60, 600, 120)  # msoTextOrientationHorizontal
        box.TextFrame.TextRange.Text = f"{i}. {title}"
    if os.path.exists(path):
        os.remove(path)
    pres.SaveAs(path)  # .pptx
    pres.Close()
    return path


def main():
    deck = os.path.join(OUT_DIR, "deck.pptx")
    pdf = os.path.join(OUT_DIR, "out.pdf")
    imgs = os.path.join(OUT_DIR, "imgs")

    log("PowerPointSession 시작 ...")
    s = PowerPointSession()
    s.start()
    try:
        log(f"PowerPoint Version = {s.run(lambda app: str(app.Version))}")

        log("입력 프레젠테이션 생성 ...")
        s.run(lambda app: _make_deck(app, deck))
        assert os.path.exists(deck), "deck.pptx 미생성"

        log("① export_pdf ...")
        s.run(lambda app: presentation.export_pdf(app, deck, pdf))
        assert os.path.exists(pdf), "PDF 미생성"

        log("② export_images (PNG) ...")
        res = s.run(lambda app: presentation.export_images(app, deck, imgs, "PNG", 1280, 720))
        log(f"   생성된 이미지 = {res['count']}장")
        assert res["count"] == 2, f"이미지 수 불일치: {res['count']}"
        for f in res["files"]:
            assert os.path.exists(f), f"이미지 미생성: {f}"

        log("=== SMOKE PASS: PowerPoint COM PDF + 슬라이드 이미지 렌더링 OK ===")
        return 0
    finally:
        log("PowerPointSession 종료 ...")
        s.stop()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        sys.exit(1)
