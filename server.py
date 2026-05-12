"""FastAPI app exposing the mddocx converter over HTTP and MCP.

Surfaces:
    GET  /                — drag-and-drop UI (single-page HTML)
    POST /convert         — multipart docx upload, returns zip (REST)
    GET  /docs            — Swagger UI
    GET  /openapi.json    — OpenAPI schema
    *    /mcp/            — MCP streamable HTTP endpoint (one tool: convert_docx)

The OpenAPI schema and the MCP tool description are both detailed enough for
an LLM client to call the service without extra instructions.
"""

from __future__ import annotations

import base64
import ipaddress
import json
import shutil
import socket
import tempfile
import urllib.parse
import urllib.request
import zipfile
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from mcp.server.fastmcp import FastMCP
from starlette.background import BackgroundTask

from mddocx import convert, parse_comments, parse_source_paragraphs

MAX_BYTES = 50 * 1024 * 1024  # 50 MB
STATIC_DIR = Path(__file__).parent / "static"

# ---------------------------------------------------------------------------
# MCP server (set up before FastAPI so it can be mounted with proper lifespan)
# ---------------------------------------------------------------------------

mcp_server = FastMCP(
    "mddocx",
    streamable_http_path="/",  # mount at FastAPI's /mcp prefix
    instructions=(
        "Convert Word .docx files into clean markdown. The single tool "
        "`convert_docx` accepts either an HTTPS URL or base64-encoded docx "
        "bytes and returns the rendered markdown, structured comments, and "
        "conversion stats."
    ),
)


def _is_safe_url(url: str) -> tuple[bool, str]:
    """Reject non-HTTPS, malformed, or private/loopback targets (basic SSRF guard)."""
    try:
        parsed = urllib.parse.urlparse(url)
    except ValueError as e:
        return False, f"unparseable url: {e}"
    if parsed.scheme != "https":
        return False, "only https:// URLs are accepted"
    host = parsed.hostname
    if not host:
        return False, "missing host in URL"
    try:
        ip_str = socket.gethostbyname(host)
        ip = ipaddress.ip_address(ip_str)
    except (socket.gaierror, ValueError) as e:
        return False, f"could not resolve host: {e}"
    if ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved or ip.is_multicast:
        return False, f"host {host} resolves to a non-public address ({ip})"
    return True, ""


def _fetch_docx_bytes(source: str) -> bytes:
    """Decode source into raw docx bytes; enforce the 50 MB cap."""
    if source.startswith(("http://", "https://")):
        ok, reason = _is_safe_url(source)
        if not ok:
            raise ValueError(reason)
        req = urllib.request.Request(source, headers={"User-Agent": "mddocx-mcp/0.1"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read(MAX_BYTES + 1)
            if len(data) > MAX_BYTES:
                raise ValueError(f"downloaded file exceeds {MAX_BYTES} bytes")
            return data
    try:
        data = base64.b64decode(source, validate=True)
    except Exception as e:
        raise ValueError(f"source is neither an https URL nor valid base64: {e}") from e
    if len(data) > MAX_BYTES:
        raise ValueError(f"decoded file exceeds {MAX_BYTES} bytes")
    return data


@mcp_server.tool()
def convert_docx(source: str) -> dict:
    """Convert a Word .docx file into clean markdown with extracted images and inline comments.

    Args:
        source: Either an `https://` URL pointing to a publicly reachable .docx
            file (max 50 MB; private/loopback addresses are refused), or a
            base64-encoded .docx file body (same size limit).

    Returns:
        A dict with:
        - ``report_md`` (str):  full markdown including inline ``<details>``
          comment blocks placed next to the paragraph they refer to.
        - ``comments`` (list):  one entry per Word comment::

              {"id": str, "author": str, "body": str, "range": str}

          where ``range`` is the exact text the comment was attached to.
        - ``stats`` (dict):  ``comments_total``, ``comments_inline``,
          ``comments_trailer``, ``media_source``, ``media_output``.
        - ``media`` (list):  filenames of images extracted from the docx
          (located under ``media/`` inside the zip from the REST endpoint).
    """
    tmp = Path(tempfile.mkdtemp(prefix="mddocx-mcp-"))
    try:
        try:
            data = _fetch_docx_bytes(source)
        except ValueError as e:
            raise ValueError(str(e)) from e

        src = tmp / "input.docx"
        src.write_bytes(data)
        out = tmp / "out"

        stats = convert(src, out)

        report_md = (out / "report.md").read_text(encoding="utf-8")

        raw = parse_comments(src)
        _, range_for = parse_source_paragraphs(src)
        comments_struct = [
            {
                "id": cid,
                "author": c.author,
                "body": c.body,
                "range": range_for.get(cid, ""),
            }
            for cid, c in raw.items()
        ]

        media_dir = out / "media"
        media = sorted(p.name for p in media_dir.iterdir()) if media_dir.exists() else []

        return {
            "report_md": report_md,
            "comments": comments_struct,
            "stats": stats,
            "media": media,
        }
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ---------------------------------------------------------------------------
# FastAPI app
# ---------------------------------------------------------------------------

# Build the MCP HTTP app eagerly so its session_manager is available for lifespan.
_mcp_http_app = mcp_server.streamable_http_app()


@asynccontextmanager
async def _lifespan(_app: FastAPI):
    async with mcp_server.session_manager.run():
        yield


app = FastAPI(
    title="mddocx",
    version="0.1.0",
    summary="Convert a Word .docx into a zip with clean markdown, images, and inline comments.",
    description=(
        "Single-purpose service. Upload a `.docx`; receive a `.zip` containing\n"
        "`report.md` (body in GitHub-flavored markdown, including tables and lists),\n"
        "`media/` (every embedded image), and `comments/cN.md` (one per Word comment).\n"
        "Comments are also embedded inline in `report.md` as collapsible `<details>` blocks.\n\n"
        "Surfaces:\n"
        "- `POST /convert` — multipart upload, returns zip (REST).\n"
        "- `*    /mcp/`    — MCP streamable HTTP transport with one tool `convert_docx`,\n"
        "                    suitable for Cursor / Claude Desktop / any MCP client.\n"
        "- `GET  /`        — drag-and-drop browser UI."
    ),
    lifespan=_lifespan,
)

app.mount("/mcp", _mcp_http_app)


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
