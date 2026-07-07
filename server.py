"""
server.py — PowerPoint 엔진 자동화 MCP 서버 (stdio).

PowerPoint.Application 데스크톱을 COM으로 자동화해, 라이브러리(python-pptx 등)에 렌더링
엔진이 없어 불가능한 **렌더링/변환** 작업을 핵심 도구로 노출한다:
  ① ppt_export_pdf      프레젠테이션 → PDF (렌더 엔진)
  ② ppt_export_images   각 슬라이드 → 이미지(PNG/JPG) 파일

모든 PowerPoint 호출은 단일 STA 워커 스레드(engine.session)에서 직렬화되며, 도구는 async로
그 블로킹 작업을 워커 스레드에 위임해 이벤트 루프를 막지 않는다. 세션은 **지연 초기화**:
첫 도구 호출 시 PowerPoint를 띄우고, 서버 종료 시 닫는다.

실행:  python server.py        (Claude Code가 stdio로 기동)
"""
from __future__ import annotations

import threading
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from typing import Any, Optional

import anyio
from mcp.server.fastmcp import FastMCP

from engine import presentation
from engine.session import PowerPointSession

# ---- 단일 세션 관리(지연 초기화) -------------------------------------------
_session: Optional[PowerPointSession] = None
_lock = threading.Lock()


def _get_session() -> PowerPointSession:
    global _session
    with _lock:
        if _session is None or not _session.alive:
            s = PowerPointSession()
            s.start()
            _session = s
        return _session


def _stop_session() -> None:
    global _session
    with _lock:
        if _session is not None:
            _session.stop()
            _session = None


@asynccontextmanager
async def _lifespan(_server: FastMCP) -> AsyncIterator[dict]:
    try:
        yield {}
    finally:
        await anyio.to_thread.run_sync(_stop_session)


mcp = FastMCP(
    "ppt-automation",
    lifespan=_lifespan,
    instructions=(
        "PowerPoint.Application COM 엔진 자동화. 라이브러리(python-pptx)가 못 하는 "
        "렌더링/변환 전용. ppt_export_pdf로 PDF 출력, ppt_export_images로 각 슬라이드를 "
        "이미지로 렌더링한다. 경로는 절대경로 권장. 원본은 보존되며 결과는 out_path/out_dir로 저장된다."
    ),
)


def _run(fn) -> Any:
    return _get_session().run(fn)


# ============================================================================
# 핵심 2개 도구 (+ 진단 2)
# ============================================================================
@mcp.tool()
async def ppt_export_pdf(src_path: str, out_path: str) -> dict:
    """① 프레젠테이션을 PowerPoint 렌더 엔진으로 PDF 출력한다.

    레이아웃·폰트·도형·SmartArt가 화면 그대로 렌더된 PDF. python-pptx에는 렌더 엔진이 없다.

    Args:
        src_path: 입력 .pptx/.ppt 절대경로
        out_path: 출력 .pdf 절대경로
    Returns: {"out_path"}
    """
    def work():
        return {"out_path": _run(lambda app: presentation.export_pdf(app, src_path, out_path))}
    return await anyio.to_thread.run_sync(work)


@mcp.tool()
async def ppt_export_images(
    src_path: str,
    out_dir: str,
    fmt: str = "PNG",
    width: Optional[int] = None,
    height: Optional[int] = None,
) -> dict:
    """② 각 슬라이드를 이미지(PNG/JPG 등) 파일로 렌더링한다(썸네일·미리보기·문서삽입용).

    Args:
        src_path: 입력 .pptx/.ppt 절대경로
        out_dir: 이미지들을 저장할 폴더(없으면 생성). slide_001.png, slide_002.png ...
        fmt: 이미지 형식 PNG/JPG/GIF/BMP/TIF (기본 PNG)
        width: 가로 픽셀(미지정 시 슬라이드 기본 크기) — width/height 둘 다 줘야 적용
        height: 세로 픽셀
    Returns: {"out_dir","count","files"}
    """
    def work():
        return _run(lambda app: presentation.export_images(app, src_path, out_dir, fmt, width, height))
    return await anyio.to_thread.run_sync(work)


# ---- 진단/복구 -------------------------------------------------------------
@mcp.tool()
async def ppt_health() -> dict:
    """엔진 상태를 점검한다(세션 기동/PowerPoint 버전). 첫 호출 시 PowerPoint를 띄운다."""
    def work():
        s = _get_session()
        try:
            ver = s.run(lambda app: str(app.Version))
        except Exception as e:  # noqa: BLE001
            ver = f"버전 조회 실패: {e}"
        return {"alive": s.alive, "powerpoint_version": ver}
    return await anyio.to_thread.run_sync(work)


@mcp.tool()
async def ppt_restart() -> dict:
    """PowerPoint 세션을 종료 후 재기동한다(COM 오류 복구용)."""
    def work():
        _stop_session()
        s = _get_session()
        return {"alive": s.alive}
    return await anyio.to_thread.run_sync(work)


if __name__ == "__main__":
    mcp.run(transport="stdio")
