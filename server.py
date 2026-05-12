"""FastAPI app exposing both directions of mddocx over HTTP and MCP.

Surfaces:
    GET  /                — drag-and-drop UI (toggle: docx→md / md→docx)
    GET  /rules           — rendered ``MARKDOWN_RULES.md`` (md formatting contract)
    GET  /template.zip    — sample md + media/ that converts cleanly with docxmd
    GET  /static/...      — anything under ``static/`` (e.g. ``template/template.md``)
    POST /convert         — docx upload → zip with markdown/media/comments (REST)
    POST /convert/md      — single .md or .zip(md+media/) upload → .docx (REST)
    GET  /docs            — Swagger UI
    GET  /openapi.json    — OpenAPI schema
    *    /mcp/            — MCP streamable HTTP endpoint; tools: convert_docx, convert_md

OpenAPI descriptions and MCP tool docstrings are written for an LLM client
to call the service without external instructions.
"""

from __future__ import annotations

import base64
import io
import ipaddress
import json
import os
import shutil
import socket
import tempfile
import urllib.parse
import urllib.request
import zipfile
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from markdown_it import MarkdownIt
from mcp.server.fastmcp import FastMCP
from mcp.server.transport_security import TransportSecuritySettings
from starlette.background import BackgroundTask

from docxmd import ConversionError as MdConversionError
from docxmd import convert as md_to_docx_convert
from mddocx import convert, parse_comments, parse_source_paragraphs

MAX_BYTES = 50 * 1024 * 1024  # 50 MB
BASE_DIR = Path(__file__).parent
STATIC_DIR = BASE_DIR / "static"
RULES_MD_PATH = BASE_DIR / "MARKDOWN_RULES.md"
GOST_RULES_MD_PATH = BASE_DIR / "GOST_RULES.md"


# ---------------------------------------------------------------------------
# MCP server (set up before FastAPI so it can be mounted with proper lifespan)
# ---------------------------------------------------------------------------


def _csv_env(name: str) -> list[str]:
    raw = os.getenv(name, "")
    return [v.strip() for v in raw.split(",") if v.strip()]


_allowed_hosts = _csv_env("MDDOCX_ALLOWED_HOSTS")
_allowed_origins = _csv_env("MDDOCX_ALLOWED_ORIGINS")

# DNS-rebinding protection is opt-in: enable only when the operator supplies an
# explicit host list (typically set in docker-compose env to match the public
# domain). Out of the box, local `uvicorn server:app` works on any Host header.
if _allowed_hosts:
    _security = TransportSecuritySettings(
        enable_dns_rebinding_protection=True,
        allowed_hosts=_allowed_hosts,
        allowed_origins=_allowed_origins,
    )
else:
    _security = TransportSecuritySettings(enable_dns_rebinding_protection=False)


mcp_server = FastMCP(
    "mddocx",
    streamable_http_path="/",  # mount at FastAPI's /mcp prefix
    instructions=(
        "Two-way Word/Markdown bridge.\n"
        "  • convert_docx: .docx → clean markdown + structured comments.\n"
        "  • convert_md:   markdown (with optional images packed in a zip) → "
        ".docx formatted per GOST 7.32 (Times New Roman 12, поля 30/15/20/20, "
        "Heading 1 по центру, остальные по левому краю, полуторный интервал, "
        "автоподписи «Рисунок N. ...» к рисункам).\n"
        "Both tools support https:// URLs or base64-encoded payloads."
    ),
    transport_security=_security,
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
    if (
        ip.is_private
        or ip.is_loopback
        or ip.is_link_local
        or ip.is_reserved
        or ip.is_multicast
    ):
        return False, f"host {host} resolves to a non-public address ({ip})"
    return True, ""


def _fetch_bytes(source: str) -> bytes:
    """Decode *source* (https URL or base64) into raw bytes; enforce ``MAX_BYTES``."""
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


_ZIP_SIGNATURES = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")


def _looks_like_zip(data: bytes) -> bool:
    return any(data.startswith(sig) for sig in _ZIP_SIGNATURES)


def _safe_extract_zip(data: bytes, dest: Path) -> None:
    """Extract a zip into *dest*, rejecting traversal and absolute paths."""
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        for info in zf.infolist():
            name = info.filename
            if not name or name.endswith("/"):
                continue
            # Reject absolute paths and any "../" segments.
            if name.startswith("/") or ".." in Path(name).parts:
                raise ValueError(f"unsafe path in zip: {name!r}")
            target = (dest / name).resolve()
            if not str(target).startswith(str(dest.resolve())):
                raise ValueError(f"zip entry escapes destination: {name!r}")
        zf.extractall(dest)


def _prepare_md_source_dir(data: bytes, tmp_root: Path, *, fallback_name: str) -> Path:
    """Materialize a bytes payload into a folder docxmd can consume.

    Accepts either a raw markdown body or a zip archive containing exactly one
    ``*.md`` and an optional ``media/`` subdirectory. Returns the directory path.
    """
    src_dir = tmp_root / "src"
    src_dir.mkdir(parents=True, exist_ok=True)
    if _looks_like_zip(data):
        _safe_extract_zip(data, src_dir)
        # The .md may sit at the zip root or one level down (common when zipping
        # a folder). Flatten if there's a single top-level directory.
        entries = list(src_dir.iterdir())
        if len(entries) == 1 and entries[0].is_dir():
            inner = entries[0]
            for child in inner.iterdir():
                shutil.move(str(child), src_dir / child.name)
            inner.rmdir()
    else:
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError as e:
            raise ValueError(
                "payload is neither a zip nor a valid UTF-8 markdown document"
            ) from e
        md_name = fallback_name if fallback_name.lower().endswith(".md") else "document.md"
        (src_dir / md_name).write_text(text, encoding="utf-8")
    return src_dir


# ---------------------------------------------------------------------------
# MCP tools
# ---------------------------------------------------------------------------


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
            data = _fetch_bytes(source)
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
        media = (
            sorted(p.name for p in media_dir.iterdir()) if media_dir.exists() else []
        )

        return {
            "report_md": report_md,
            "comments": comments_struct,
            "stats": stats,
            "media": media,
        }
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


@mcp_server.tool()
def convert_md(source: str, filename: str = "document.docx") -> dict:
    """Convert markdown (optionally with bundled images) into a GOST-styled .docx.

    The markdown must follow the contract described at ``/rules`` of this server
    (see also ``MARKDOWN_RULES.md`` in the repository). The output document is
    formatted per GOST 7.32: Times New Roman 12pt, 1.5 line spacing, 30/15/20/20mm
    margins, Heading 1 centered, lower headings left-aligned, image captions
    "Рисунок N. <alt-text>".

    Args:
        source: One of:
            • plain UTF-8 markdown text (when no images are needed);
            • a base64-encoded **zip** archive containing exactly one ``*.md``
              file and an optional ``media/`` subdirectory referenced by image
              links in the markdown;
            • an ``https://`` URL pointing to the same kind of zip (or to a
              plain ``.md`` file). Private/loopback addresses are refused.
            All payloads are capped at 50 MB.
        filename: Desired output filename (must end with ``.docx``; defaults to
            ``document.docx``).

    Returns:
        A dict with:
        - ``filename`` (str):       echoed output filename.
        - ``content_base64`` (str): base64-encoded ``.docx`` bytes.
        - ``size_bytes`` (int):     size of the decoded docx, for sanity checks.
    """
    if not filename.lower().endswith(".docx"):
        filename = filename + ".docx"

    tmp = Path(tempfile.mkdtemp(prefix="docxmd-mcp-"))
    try:
        try:
            data = _fetch_bytes(source)
        except ValueError as e:
            raise ValueError(str(e)) from e

        try:
            src_dir = _prepare_md_source_dir(
                data, tmp, fallback_name="document.md"
            )
        except ValueError as e:
            raise ValueError(str(e)) from e

        out_path = tmp / "output.docx"
        try:
            md_to_docx_convert(src_dir, output=out_path)
        except MdConversionError as e:
            raise ValueError(str(e)) from e

        body = out_path.read_bytes()
        return {
            "filename": filename,
            "content_base64": base64.b64encode(body).decode("ascii"),
            "size_bytes": len(body),
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
    version="0.2.0",
    summary=(
        "Two-way bridge between Word .docx and Markdown. "
        "docx → markdown + media + inline comments; markdown(+media) → "
        "GOST-styled .docx."
    ),
    description=(
        "Surfaces:\n"
        "- `POST /convert`     — docx upload, returns a zip with `report.md`, "
        "`media/`, and `comments/cN.md`. Comments are also embedded inline as "
        "collapsible `<details>` blocks.\n"
        "- `POST /convert/md`  — single `.md` upload or `.zip(md+media/)` "
        "upload, returns a `.docx` formatted per GOST 7.32 (TNR 12, 1.5 line "
        "spacing, fields 30/15/20/20mm, captions «Рисунок N. ...»).\n"
        "- `GET  /rules`       — rendered markdown contract for `/convert/md` "
        "(copy/paste-friendly for LLM system prompts).\n"
        "- `GET  /template.zip`— sample `template.md` + `media/` that "
        "round-trips through `/convert/md` cleanly.\n"
        "- `*    /mcp/`        — MCP streamable HTTP transport with two tools "
        "(`convert_docx`, `convert_md`), suitable for Cursor / Claude Desktop.\n"
        "- `GET  /`            — drag-and-drop browser UI with direction toggle."
    ),
    lifespan=_lifespan,
)

app.mount("/mcp", _mcp_http_app)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/", include_in_schema=False, response_class=HTMLResponse)
def index() -> HTMLResponse:
    return HTMLResponse((STATIC_DIR / "index.html").read_text(encoding="utf-8"))


_RULES_PARSER = (
    MarkdownIt("commonmark", {"breaks": False, "html": False})
    .enable("table")
    .enable("strikethrough")
)


@app.get("/rules", include_in_schema=False, response_class=HTMLResponse)
def rules() -> HTMLResponse:
    gost_html = _RULES_PARSER.render(GOST_RULES_MD_PATH.read_text(encoding="utf-8"))
    md_html = _RULES_PARSER.render(RULES_MD_PATH.read_text(encoding="utf-8"))
    layout = (STATIC_DIR / "rules.html").read_text(encoding="utf-8")
    return HTMLResponse(
        layout.replace("<!--GOST-->", gost_html).replace("<!--MARKDOWN-->", md_html)
    )


@app.get(
    "/template.zip",
    include_in_schema=False,
    response_class=Response,
)
def template_zip() -> Response:
    """Pack the bundled sample md + media/ into a fresh zip on every request."""
    template_dir = STATIC_DIR / "template"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(template_dir.rglob("*")):
            if path.is_file():
                zf.write(path, path.relative_to(template_dir))
    return Response(
        content=buf.getvalue(),
        media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="template.zip"'},
    )


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


@app.post(
    "/convert/md",
    tags=["convert"],
    summary="Convert a Markdown file or a (md + media/) zip into a GOST-styled .docx",
    response_class=FileResponse,
    responses={
        200: {
            "content": {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document": {}
            },
            "description": (
                "DOCX document formatted per GOST 7.32 (Times New Roman 12, "
                "1.5 spacing, fields 30/15/20/20mm, Heading 1 centered, "
                "image captions «Рисунок N. ...»). See `/rules` for the "
                "markdown structure contract."
            ),
        },
        400: {
            "description": (
                "Invalid input (wrong extension, missing/multiple .md in zip, "
                "empty file, traversal in zip, too large)"
            )
        },
        500: {"description": "Conversion failed"},
    },
)
def convert_md_endpoint(
    file: UploadFile = File(
        ...,
        description=(
            "Either a single `.md` file or a `.zip` archive containing exactly "
            "one `.md` file and an optional `media/` subdirectory. Max 50 MB."
        ),
    ),
):
    name = (file.filename or "").strip()
    lower = name.lower()
    if not (lower.endswith(".md") or lower.endswith(".zip") or lower.endswith(".markdown")):
        raise HTTPException(
            status_code=400, detail="expected a .md, .markdown, or .zip file"
        )

    tmp_root = Path(tempfile.mkdtemp(prefix="docxmd-"))
    cleanup = BackgroundTask(shutil.rmtree, tmp_root, ignore_errors=True)
    try:
        buf = io.BytesIO()
        size = 0
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
            buf.write(chunk)
        if size == 0:
            raise HTTPException(status_code=400, detail="empty upload")

        data = buf.getvalue()
        try:
            src_dir = _prepare_md_source_dir(
                data, tmp_root, fallback_name=Path(name).name or "document.md"
            )
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e)) from e
        except zipfile.BadZipFile as e:
            raise HTTPException(
                status_code=400, detail=f"invalid zip archive: {e}"
            ) from e

        stem = Path(name).stem or "document"
        out_path = tmp_root / f"{stem}.docx"
        try:
            md_to_docx_convert(src_dir, output=out_path)
        except MdConversionError as e:
            raise HTTPException(status_code=400, detail=str(e)) from e
        except Exception as e:
            raise HTTPException(
                status_code=500, detail=f"conversion failed: {e}"
            ) from e

        return FileResponse(
            out_path,
            filename=f"{stem}.docx",
            media_type=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
            background=cleanup,
        )
    except HTTPException:
        cleanup.func(*cleanup.args, **cleanup.kwargs)
        raise
    except Exception as e:
        cleanup.func(*cleanup.args, **cleanup.kwargs)
        return JSONResponse(status_code=500, content={"detail": f"internal error: {e}"})
