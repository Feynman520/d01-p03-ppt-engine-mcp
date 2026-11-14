# ppt-engine-mcp — 한국어 안내

[English → README.md](README.md)

PowerPoint.Application 데스크톱을 COM으로 자동화해, **라이브러리(python-pptx)에 렌더링 엔진이
없어 불가능한 렌더링/변환 작업**을 핵심 도구로 노출하는 로컬 MCP 서버.

> 설계 사상: 슬라이드 생성·텍스트 채우기는 python-pptx가 빠르고, **PDF·이미지 렌더링**은
> 엔진만 가능하다. 이 MCP는 후자(엔진 전용)만 담당한다 — 그래서 도구가 4개뿐이다.

## 요구사항
- Windows 10+ / **로그인된 인터랙티브 데스크톱 세션**
- Microsoft Office(PowerPoint) 설치 — 검증 환경: **Office 2016+ (PowerPoint 16.0)**
- Python 3.10+ — 검증: **3.12**
- [Claude Code](https://claude.com/claude-code) 또는 임의의 MCP 클라이언트

## 설치

```powershell
git clone https://github.com/Feynman520/d01-p03-ppt-engine-mcp.git
cd d01-p03-ppt-engine-mcp
py -3.12 -m venv .venv          # 또는: python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### Claude Code 등록

클론한 폴더 안에서 아래 한 줄 실행 (절대경로가 박히므로 이후 어디서든 동작):

```powershell
claude mcp add ppt-automation --scope user -- "$PWD\.venv\Scripts\python.exe" "$PWD\server.py"
```

`--scope user`는 모든 프로젝트에서 사용 가능. 한 프로젝트에서만 쓰려면 `--scope project`.

### 동작 검증

```powershell
$py = ".\.venv\Scripts\python.exe"; $env:PYTHONUTF8 = "1"
& $py tests\smoke_com.py      # COM PDF·슬라이드 이미지 렌더링·좀비정리 검증
& $py tests\server_tools.py   # MCP 4개 도구 등록 검증(PowerPoint 미기동)
```

## 도구 (핵심 2 + 진단 2)

| # | 도구 | 입력 → 출력 | 엔진 전용 이유 |
|---|---|---|---|
| ① | `ppt_export_pdf` | `src_path, out_path` → `{out_path}` | PowerPoint 렌더 엔진 PDF(python-pptx에 렌더 엔진 없음) |
| ② | `ppt_export_images` | `src_path, out_dir, fmt?, width?, height?` → `{out_dir,count,files}` | 각 슬라이드를 이미지로 렌더(썸네일·미리보기·문서삽입) |
| — | `ppt_health` | → `{alive, powerpoint_version}` | 세션 점검(첫 호출 시 PowerPoint 기동) |
| — | `ppt_restart` | → `{alive}` | COM 오류 복구 |

전형적 흐름: python-pptx 등으로 만든 `.pptx` → `ppt_export_pdf`로 배포용 PDF /
`ppt_export_images`로 슬라이드 썸네일. 원본은 절대 수정하지 않고 결과는 항상
`out_path`/`out_dir`로 저장한다.

## 아키텍처 핵심
- **단일 STA 워커 스레드**(`engine/session.py`): 모든 PowerPoint 호출을 전용 스레드 1개에 직렬화.
- **세션 지연 초기화**: 첫 도구 호출 시 PowerPoint 기동, 서버 종료 시 닫음. 1개 인스턴스 재사용.
- **`DispatchEx` + `gencache.EnsureDispatch`(조기바인딩)**: 전용 인스턴스를 새로 띄우고 조기바인딩 호출.
  PDF는 `Presentation.SaveAs(path, 32=ppSaveAsPDF)`, 이미지는 `Slide.Export(path, "PNG", w, h)`.
- **`Visible=False` 금지**: PowerPoint는 앱 창 숨김 설정 시 예외 → 설정하지 않고
  `Presentations.Open(..., WithWindow=msoFalse)`로 **창 없이** 연다(화면 미표시·원본 보존).
- **좀비 방지**: PowerPoint.Application은 Hwnd를 안 주므로, DispatchEx 전후 `POWERPNT.EXE`
  프로세스 목록을 비교(diff)해 전용 인스턴스 PID를 식별하고, 종료 시 남으면 강제 종료한다.
- **RPC 거부 재시도 / stdout 보호**: excel-engine-mcp와 동일.

## 한계
- 무인/서비스 세션 부적합(데스크톱 세션 필요).
- 슬라이드 생성·편집은 python-pptx가 빠름 — 의도된 역할 분담이다.

## 라이선스

[MIT](LICENSE)

## 기여

의견과 PR 안내는 [CONTRIBUTING.md](CONTRIBUTING.md)를, 보안 문제 신고는 [SECURITY.md](SECURITY.md)를 봐 주세요.
