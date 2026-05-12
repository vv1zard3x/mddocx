"""FastAPI app exposing the mddocx converter over HTTP.

One endpoint does the work — POST /convert. The OpenAPI schema at
``/openapi.json`` (UI at ``/docs``) is detailed enough for an LLM client
to call the service without extra instructions.
"""

from __future__ import annotations

import io
import shutil
import tempfile
import zipfile
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from starlette.background import BackgroundTask

from mddocx import convert

MAX_BYTES = 50 * 1024 * 1024  # 50 MB
STATIC_DIR = Path(__file__).parent / "static"

app = FastAPI(
    title="mddocx",
    version="0.1.0",
    summary="Convert a Word .docx into a zip with clean markdown, images, and inline comments.",
    description=(
        "Single-purpose service. Upload a `.docx`; receive a `.zip` containing\n"
        "`report.md` (body in GitHub-flavored markdown, including tables and lists),\n"
        "`media/` (every embedded image), and `comments/cN.md` (one per Word comment).\n"
        "Comments are also embedded inline in `report.md` as collapsible `<details>` blocks.\n\n"
        "Intended for both interactive use (drag-and-drop UI at `/`) and "
        "programmatic use (`POST /convert` accepts multipart/form-data with one file)."
    ),
)


@app.get("/", include_in_schema=False, response_class=HTMLResponse)
def index() -> HTMLResponse:
    return HTMLResponse((STATIC_DIR / "index.html").read_text(encoding="utf-8"))


@app.get("/health", tags=["meta"], summary="Liveness probe")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post(
    "/convert",
    tags=["convert"],
    summary="Convert a .docx file to a zip with markdown, media, and comments",
    response_class=FileResponse,
    responses={
        200: {
            "content": {"application/zip": {}},
            "description": (
                "ZIP archive named `<basename>.zip` containing `report.md`, "
                "`media/`, and (if any) `comments/cN.md`. "
                "Conversion stats are returned in the `X-Mddocx-Stats` response header "
                "as a compact JSON string (totals for comments and media)."
            ),
        },
        400: {"description": "Invalid input (wrong extension, empty file, too large)"},
        500: {"description": "Conversion failed (pandoc or parsing error)"},
    },
)
def convert_endpoint(
    file: UploadFile = File(
        ...,
        description="A `.docx` file. Maximum size 50 MB.",
    ),
):
    name = (file.filename or "").strip()
    if not name.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="expected a .docx file")

    tmp_root = Path(tempfile.mkdtemp(prefix="mddocx-"))
    cleanup = BackgroundTask(shutil.rmtree, tmp_root, ignore_errors=True)
    try:
        src_path = tmp_root / "input.docx"
        size = 0
        with src_path.open("wb") as dst:
            while True:
                chunk = file.file.read(1 << 20)
                if not chunk:
                    break
                size += len(chunk)
                if size > MAX_BYTES:
                    raise HTTPException(
                        status_code=400,
                        detail=f"file exceeds maximum size ({MAX_BYTES} bytes)",
                    )
                dst.write(chunk)
        if size == 0:
            raise HTTPException(status_code=400, detail="empty upload")

        out_dir = tmp_root / "out"
        try:
            stats = convert(src_path, out_dir)
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e)) from e
        except RuntimeError as e:
            raise HTTPException(status_code=500, detail=str(e)) from e

        stem = Path(name).stem or "report"
        zip_path = tmp_root / f"{stem}.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in sorted(out_dir.rglob("*")):
                if path.is_file():
                    zf.write(path, path.relative_to(out_dir))

        import json
        return FileResponse(
            zip_path,
            filename=f"{stem}.zip",
            media_type="application/zip",
            headers={"X-Mddocx-Stats": json.dumps(stats, ensure_ascii=False)},
            background=cleanup,
        )
    except HTTPException:
        cleanup.func(*cleanup.args, **cleanup.kwargs)
        raise
    except Exception as e:
        cleanup.func(*cleanup.args, **cleanup.kwargs)
        return JSONResponse(status_code=500, content={"detail": f"internal error: {e}"})
