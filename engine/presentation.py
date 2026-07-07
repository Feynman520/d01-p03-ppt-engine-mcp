"""
engine/presentation.py — PowerPoint 엔진 전용 작업: ① PDF 렌더링  ② 슬라이드 이미지 렌더링.

모든 함수는 `app`(PowerPoint.Application COM)을 첫 인자로 받고 전용 워커 스레드에서 호출된다
(session.run). 라이브러리(python-pptx)에는 렌더링 엔진이 없어 불가능한, 진짜 PowerPoint
렌더 엔진을 쓰는 작업만 노출한다. 원본 보존: 읽기전용/창없음으로 열고 새 경로로만 출력한다.
"""
from __future__ import annotations

import os
from typing import Optional

# --- 후기/조기바인딩 공용 숫자 상수 ------------------------------------------
PP_SAVE_AS_PDF = 32   # ppSaveAsPDF
MSO_TRUE = -1         # msoTrue
MSO_FALSE = 0         # msoFalse
_IMG_EXT = {"PNG": "png", "JPG": "jpg", "JPEG": "jpg", "GIF": "gif", "BMP": "bmp", "TIF": "tif"}


def _abspath(path: str) -> str:
    return os.path.abspath(os.path.expanduser(path))


def _ensure_parent(path: str) -> None:
    os.makedirs(os.path.dirname(_abspath(path)), exist_ok=True)


def _open(app, path: str):
    """프레젠테이션을 읽기전용·창없음으로 연다(원본 보존, 화면 미표시)."""
    src = _abspath(path)
    if not os.path.exists(src):
        raise FileNotFoundError(f"원본을 찾을 수 없습니다: {src}")
    # Open(FileName, ReadOnly, Untitled, WithWindow)
    return app.Presentations.Open(src, MSO_TRUE, MSO_FALSE, MSO_FALSE)


# ============================================================================
# ① PDF 렌더링 — PowerPoint 렌더 엔진으로 PDF 출력
# ============================================================================
def export_pdf(app, src_path: str, out_path: str) -> str:
    """프레젠테이션을 PowerPoint 렌더 엔진으로 PDF 출력한다(레이아웃·폰트·도형 그대로).

    python-pptx 등 라이브러리에는 렌더 엔진이 없어 PDF로 구울 수 없다.
    """
    if os.path.splitext(out_path)[1].lower() != ".pdf":
        raise ValueError("out_path 확장자는 .pdf 여야 합니다.")
    _ensure_parent(out_path)
    pres = _open(app, src_path)
    try:
        pres.SaveAs(_abspath(out_path), PP_SAVE_AS_PDF)
    finally:
        try:
            pres.Close()
        except Exception:
            pass
    if not os.path.exists(_abspath(out_path)):
        raise RuntimeError(f"PDF 저장 실패: {out_path}")
    return _abspath(out_path)


# ============================================================================
# ② 슬라이드 이미지 렌더링 — 각 슬라이드를 이미지 파일로 내보내기
# ============================================================================
def export_images(
    app,
    src_path: str,
    out_dir: str,
    fmt: str = "PNG",
    width: Optional[int] = None,
    height: Optional[int] = None,
) -> dict:
    """각 슬라이드를 이미지(PNG/JPG 등) 파일로 렌더링한다(썸네일·미리보기·문서 삽입용).

    width/height 미지정 시 슬라이드 기본 픽셀 크기로 출력한다.
    파일명: out_dir/slide_001.<ext>, slide_002.<ext> ...
    """
    fmt = (fmt or "PNG").upper()
    if fmt not in _IMG_EXT:
        raise ValueError(f"지원하지 않는 형식: {fmt} (가능: {', '.join(_IMG_EXT)})")
    ext = _IMG_EXT[fmt]
    out_dir_abs = _abspath(out_dir)
    os.makedirs(out_dir_abs, exist_ok=True)

    pres = _open(app, src_path)
    files = []
    try:
        count = int(pres.Slides.Count)
        for i in range(1, count + 1):
            target = os.path.join(out_dir_abs, f"slide_{i:03d}.{ext}")
            slide = pres.Slides(i)
            if width and height:
                slide.Export(target, fmt, int(width), int(height))
            else:
                slide.Export(target, fmt)
            files.append(target)
    finally:
        try:
            pres.Close()
        except Exception:
            pass
    missing = [f for f in files if not os.path.exists(f)]
    if missing:
        raise RuntimeError(f"이미지 일부 미생성: {missing[:3]} ...")
    return {"out_dir": out_dir_abs, "count": len(files), "files": files}
