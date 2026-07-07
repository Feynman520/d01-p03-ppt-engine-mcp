# ppt-engine-mcp

[한국어 안내 → README.ko.md](README.ko.md)

A local MCP server that drives the **real PowerPoint desktop application** (COM automation) to do the one thing python-pptx fundamentally cannot: **render** — true-fidelity **PDF export and per-slide image export**.

> Design philosophy: this server **complements** library workflows instead of replacing them.
> Building slides and filling text is faster with python-pptx — but no library has
> PowerPoint's render engine. This MCP handles only the engine-exclusive part, keeping its
> tool surface tiny (4 tools).

## Requirements

- Windows 10+ with a **logged-in interactive desktop session** (PowerPoint has no true headless mode)
- Microsoft Office (PowerPoint) installed and licensed — verified on **Office 2016+ (PowerPoint 16.0)**
- Python 3.10+ — verified on **3.12**
- [Claude Code](https://claude.com/claude-code) or any MCP client

## Install

```powershell
git clone https://github.com/Feynman520/ppt-engine-mcp.git
cd ppt-engine-mcp
py -3.12 -m venv .venv          # or: python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### Register with Claude Code

Run this in the cloned folder (uses absolute paths, so it works from anywhere afterwards):

```powershell
claude mcp add ppt-automation --scope user -- "$PWD\.venv\Scripts\python.exe" "$PWD\server.py"
```

`--scope user` makes it available in every project. Use `--scope project` to limit it to one project.

### Verify

```powershell
$py = ".\.venv\Scripts\python.exe"; $env:PYTHONUTF8 = "1"
& $py tests\smoke_com.py      # COM PDF / slide image rendering / zombie cleanup
& $py tests\server_tools.py   # MCP tool registration (does not launch PowerPoint)
```

## Tools (2 core + 2 diagnostics)

| # | Tool | Input → Output | Why engine-only |
|---|---|---|---|
| ① | `ppt_export_pdf` | `src_path, out_path` → `{out_path}` | PowerPoint render engine PDF (python-pptx has no render engine) |
| ② | `ppt_export_images` | `src_path, out_dir, fmt?, width?, height?` → `{out_dir,count,files}` | Renders each slide to an image (thumbnails, previews, embedding in documents) |
| — | `ppt_health` | → `{alive, powerpoint_version}` | Session check (launches PowerPoint on first call) |
| — | `ppt_restart` | → `{alive}` | Recovery from COM errors |

Typical flow: build a `.pptx` with python-pptx → `ppt_export_pdf` for the distributable PDF /
`ppt_export_images` for slide thumbnails.
Originals are never modified; results are always written to `out_path`/`out_dir`.

## Architecture notes

- **Single STA worker thread** (`engine/session.py`): every PowerPoint call is serialized onto
  one dedicated thread (win32com COM objects are thread-bound; FastMCP may hop threads).
- **Lazy session**: PowerPoint starts on the first tool call, is reused, and closes with the server.
- **`DispatchEx` + early binding (`gencache.EnsureDispatch`)**: a dedicated instance with
  type-library calls. PDF via `Presentation.SaveAs(path, 32 /* ppSaveAsPDF */)`, images via
  `Slide.Export(path, "PNG", w, h)`.
- **No `Visible=False`**: unlike Excel/Word, PowerPoint throws when you hide the app window —
  instead presentations are opened windowless via `Presentations.Open(..., WithWindow=msoFalse)`.
- **Zombie prevention**: PowerPoint gives no usable window handle, so the dedicated instance
  PID is identified by diffing the `POWERPNT.EXE` process list before/after `DispatchEx`, and
  force-killed at shutdown if it survives `Quit()`.
- **RPC-rejection retry / stdout protection**: same as excel-engine-mcp.

## Limitations

- Not suitable for unattended/service sessions (needs an interactive desktop).
- Slide creation/editing is faster with python-pptx — that is by design.

## License

[MIT](LICENSE)
