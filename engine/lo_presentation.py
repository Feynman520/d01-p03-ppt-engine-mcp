"""
engine/lo_presentation.py — macOS 엔진의 ① PDF ② 슬라이드 이미지(engine/presentation.py 와 같은 함수·같은 반환).

첫 인자 `app` 은 COM 의 PowerPoint.Application 대신 engine/lo_bridge.LoBridge 다. 렌더링은 LibreOffice Impress 가
한다 — PowerPoint 전용 효과(일부 SmartArt·글꼴 대체)는 LibreOffice 가 읽는 만큼 반영된다.
"""
from __future__ import annotations

import os
from typing import Optional

_IMG_EXT = {"PNG": "png", "JPG": "jpg", "JPEG": "jpg", "GIF": "gif", "BMP": "bmp", "TIF": "tif"}


def _abspath(path: str) -> str:
    return os.path.abspath(os.path.expanduser(path))


def _require(path: str) -> str:
    src = _abspath(path)
    if not os.path.exists(src):
        raise FileNotFoundError(f"원본을 찾을 수 없습니다: {src}")
    return src


def export_pdf(app, src_path: str, out_path: str) -> str:
    """프레젠테이션 전체를 PDF 로."""
    if os.path.splitext(out_path)[1].lower() != ".pdf":
        raise ValueError("out_path 확장자는 .pdf 여야 합니다.")
    return app.run("impress_pdf", src=_require(src_path), out=_abspath(out_path))["out_path"]


def export_images(app, src_path: str, out_dir: str, fmt: str = "PNG", width: Optional[int] = None, height: Optional[int] = None) -> dict:
    """슬라이드마다 이미지: out_dir/slide_001.<ext> … width·height 를 둘 다 줘야 그 크기(아니면 슬라이드 크기 96dpi)."""
    f = (fmt or "PNG").upper()
    if f not in _IMG_EXT:
        raise ValueError(f"지원하지 않는 형식: {f} (가능: {', '.join(_IMG_EXT)})")
    both = bool(width and height)
    return app.run("impress_images", src=_require(src_path), out_dir=_abspath(out_dir), fmt=f,
                   width=int(width) if both else None, height=int(height) if both else None)
