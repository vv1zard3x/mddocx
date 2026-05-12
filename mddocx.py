"""Convert a .docx into clean markdown + media + per-comment files.

Body conversion goes through pandoc (system binary). Comments are parsed
directly from word/comments.xml and word/document.xml via stdlib, then
embedded inline as collapsible <details> blocks next to the paragraph they
refer to. Comments that cannot be unambiguously located in the rendered
markdown fall back to a `## Comments` trailer at the bottom.

Usage:
    mddocx INPUT.docx [OUTPUT_DIR] [--force]

Produces:
    OUTPUT_DIR/
        report.md
        media/
        comments/cN.md     (only if the source has comments)
"""

from __future__ import annotations

import argparse
import html
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = NS["w"]

_BOOKMARK_EMPTY = re.compile(r"\[\]\{#[A-Za-z0-9_.\-:]+\}")
_TRAILING_ATTR = re.compile(r"[ \t]*\{#[A-Za-z0-9_.\-:]+\}[ \t]*$", re.MULTILINE)

# Markdown stripping regexes used for matching pandoc output against the
# plain text we pulled from document.xml.
_MD_IMAGE = re.compile(r"!\[[^\]]*\]\([^)]*\)")
_MD_LINK = re.compile(r"\[([^\]]+)\]\([^)]*\)")
_MD_CODE_BLOCK = re.compile(r"```.*?```", re.S)
_MD_INLINE_CODE = re.compile(r"`([^`]+)`")
_MD_BOLD = re.compile(r"(\*\*|__)(.+?)\1")
_MD_ITAL = re.compile(r"(?<!\*)\*([^*\n]+)\*(?!\*)|(?<!_)_([^_\n]+)_(?!_)")
_MD_STRIKE = re.compile(r"~~(.+?)~~")
_MD_HTML = re.compile(r"<[^>]+>")
_MD_LIST = re.compile(r"^[ \t]*(?:[-*+]|\d+\.)\s+", re.M)
_MD_HEADING = re.compile(r"^#+\s+", re.M)
_MD_BLOCKQUOTE = re.compile(r"^>+\s*", re.M)
_MD_ESCAPE = re.compile(r"\\([\\`*_{}\[\]()#+\-.!~<>|])")
_WS = re.compile(r"\s+")


def normalize(s: str) -> str:
    """Strip markdown formatting + collapse whitespace for fuzzy matching."""
    s = _MD_CODE_BLOCK.sub("", s)
    s = _MD_IMAGE.sub("", s)
    s = _MD_LINK.sub(r"\1", s)
    s = _MD_INLINE_CODE.sub(r"\1", s)
    s = _MD_BOLD.sub(r"\2", s)
    s = _MD_ITAL.sub(lambda m: m.group(1) or m.group(2) or "", s)
    s = _MD_STRIKE.sub(r"\1", s)
    s = _MD_HTML.sub("", s)
    s = _MD_LIST.sub("", s)
    s = _MD_HEADING.sub("", s)
    s = _MD_BLOCKQUOTE.sub("", s)
    s = _MD_ESCAPE.sub(r"\1", s)
    s = _WS.sub(" ", s).strip()
    return s


@dataclass
class Comment:
    cid: str
    author: str
    body: str

    def snippet(self, limit: int = 80) -> str:
        flat = " ".join(self.body.split())
        return flat[:limit] + ("…" if len(flat) > limit else "")


@dataclass
class SourcePara:
    """A document.xml paragraph that contains the end of at least one comment range."""
    text: str
    comment_ids: list[str] = field(default_factory=list)


def parse_comments(docx: Path) -> dict[str, Comment]:
    with zipfile.ZipFile(docx) as z:
        if "word/comments.xml" not in z.namelist():
            return {}
        xml_bytes = z.read("word/comments.xml")

    root = ET.fromstring(xml_bytes)
    out: dict[str, Comment] = {}
    for c in root.findall("w:comment", NS):
        cid = c.attrib.get(f"{{{W}}}id", "?")
        author = c.attrib.get(f"{{{W}}}author", "Unknown")
        paragraphs: list[str] = []
        for p in c.findall(".//w:p", NS):
            text = "".join(t.text or "" for t in p.findall(".//w:t", NS))
            paragraphs.append(text)
        body = "\n\n".join(p for p in paragraphs if p)
        if not body:
            body = "".join(t.text or "" for t in c.findall(".//w:t", NS))
        out[cid] = Comment(cid=cid, author=author, body=body)
    return out


def parse_source_paragraphs(docx: Path) -> tuple[list[SourcePara], dict[str, str]]:
    """Return:
    - list of paragraphs (in document order) that close at least one comment range,
      with the comment IDs that end in each;
    - mapping comment_id -> commented range text (best-effort).
    """
    with zipfile.ZipFile(docx) as z:
        if "word/document.xml" not in z.namelist():
            return [], {}
        xml_bytes = z.read("word/document.xml")

    root = ET.fromstring(xml_bytes)
    body = root.find("w:body", NS)
    if body is None:
        return [], {}

    paragraphs: list[SourcePara] = []
    range_text: dict[str, list[str]] = {}
    open_ranges: set[str] = set()

    for p in body.iter(f"{{{W}}}p"):
        para_text_parts: list[str] = []
        closing_here: list[str] = []
        for elem in p.iter():
            tag = elem.tag.split("}", 1)[-1]
            if tag == "commentRangeStart":
                cid = elem.attrib.get(f"{{{W}}}id")
                if cid:
                    open_ranges.add(cid)
                    range_text.setdefault(cid, [])
            elif tag == "commentRangeEnd":
                cid = elem.attrib.get(f"{{{W}}}id")
                if cid:
                    open_ranges.discard(cid)
                    closing_here.append(cid)
            elif tag == "t" and elem.text:
                para_text_parts.append(elem.text)
                for cid in open_ranges:
                    range_text.setdefault(cid, []).append(elem.text)

        text = "".join(para_text_parts).strip()
        if closing_here and text:
            paragraphs.append(SourcePara(text=text, comment_ids=closing_here))

    range_text_flat = {cid: "".join(parts).strip() for cid, parts in range_text.items()}
    return paragraphs, range_text_flat


def run_pandoc(src: Path, out_dir: Path) -> None:
    if shutil.which("pandoc") is None:
        raise RuntimeError("pandoc not found in PATH; install via your package manager")
    cmd = [
        "pandoc",
        str(src),
        "-t", "gfm",
        "-o", str(out_dir / "report.md"),
        f"--extract-media={out_dir}",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"pandoc failed: {result.stderr.strip()}")
    if result.stderr.strip():
        print(result.stderr, file=sys.stderr)


def clean_pandoc_artifacts(report_md: Path) -> None:
    text = report_md.read_text(encoding="utf-8")
    text = _BOOKMARK_EMPTY.sub("", text)
    text = _TRAILING_ATTR.sub("", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    report_md.write_text(text, encoding="utf-8")


def write_comment_files(comments: dict[str, Comment], out_dir: Path) -> None:
    comments_dir = out_dir / "comments"
    comments_dir.mkdir(parents=True, exist_ok=True)
    for c in comments.values():
        if c.body.strip():
            quoted = "\n".join(
                f"> {line}" if line else ">" for line in c.body.splitlines()
            )
        else:
            quoted = "> *(empty)*"
        (comments_dir / f"c{c.cid}.md").write_text(
            f"{quoted}\n>\n> — *{c.author}*\n", encoding="utf-8"
        )


def render_details(comment: Comment, range_text: str) -> str:
    """Inline `<details>` block embedding the commented range and the body.

    Trailing ``---`` before ``</details>`` marks the end of the comment so
    the boundary is visible even in plain-text renderers that do not
    recognise the ``<details>`` tag.
    """
    author = html.escape(comment.author)
    range_quote = (range_text or "").strip()
    if range_quote:
        quoted = "\n".join(f"> {line}" for line in range_quote.splitlines())
    else:
        quoted = "> *(range not captured)*"
    body = comment.body.strip() or "*(empty)*"
    return (
        f"<details><summary>💬 c{comment.cid} — <em>{author}</em></summary>\n"
        f"\n"
        f"{quoted}\n"
        f"\n"
        f"{body}\n"
        f"\n"
        f"---\n"
        f"\n"
        f"</details>"
    )


def iter_md_blocks(lines: list[str]):
    """Yield (start, end, normalized_text) for each markdown block.

    A block is a run of consecutive non-blank lines that is not inside a
    fenced code block. `end` is exclusive (first line after the block).
    """
    in_code = False
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        stripped = line.lstrip()
        if stripped.startswith("```") or stripped.startswith("~~~"):
            in_code = not in_code
            i += 1
            continue
        if in_code or not line.strip():
            i += 1
            continue
        start = i
        while i < n:
            inner = lines[i].lstrip()
            if not lines[i].strip():
                break
            if inner.startswith("```") or inner.startswith("~~~"):
                break
            i += 1
        end = i
        text = "\n".join(lines[start:end])
        yield (start, end, normalize(text))


def embed_inline(
    report_md: Path,
    sources: list[SourcePara],
    comments: dict[str, Comment],
    range_text_for: dict[str, str],
) -> set[str]:
    """Insert `<details>` blocks after matched paragraphs. Return set of IDs that were embedded."""
    md = report_md.read_text(encoding="utf-8")
    lines = md.split("\n")
    blocks = list(iter_md_blocks(lines))

    insertions_by_block: dict[int, list[str]] = {}
    embedded: set[str] = set()
    cursor = 0

    for src in sources:
        target = normalize(src.text)
        if not target:
            continue
        found_idx = None
        for j in range(cursor, len(blocks)):
            block_text = blocks[j][2]
            if target in block_text or block_text in target:
                found_idx = j
                break
        if found_idx is None:
            continue
        for cid in src.comment_ids:
            comment = comments.get(cid)
            if comment is None:
                continue
            rtext = range_text_for.get(cid, "")
            insertions_by_block.setdefault(found_idx, []).append(
                render_details(comment, rtext)
            )
            embedded.add(cid)
        cursor = found_idx + 1

    if not insertions_by_block:
        return embedded

    new_lines: list[str] = []
    i = 0
    block_idx = 0
    while i < len(lines):
        if block_idx < len(blocks) and i == blocks[block_idx][0]:
            start, end, _ = blocks[block_idx]
            new_lines.extend(lines[start:end])
            for details in insertions_by_block.get(block_idx, []):
                new_lines.append("")
                new_lines.extend(details.split("\n"))
            i = end
            block_idx += 1
        else:
            new_lines.append(lines[i])
            i += 1

    text = "\n".join(new_lines)
    text = re.sub(r"\n{3,}", "\n\n", text)
    report_md.write_text(text, encoding="utf-8")
    return embedded


def append_unmatched_trailer(
    report_md: Path,
    comments: dict[str, Comment],
    unmatched_ids: list[str],
    range_text_for: dict[str, str],
) -> None:
    if not unmatched_ids:
        return
    lines = ["", "## Comments (unmatched)", ""]
    for cid in unmatched_ids:
        c = comments.get(cid)
        if c is None:
            continue
        rtext = range_text_for.get(cid, "").strip()
        lines.append(f"### c{cid} — *{c.author}*")
        lines.append("")
        if rtext:
            for rline in rtext.splitlines():
                lines.append(f"> {rline}")
            lines.append("")
        body = c.body.strip() or "*(empty)*"
        lines.append(body)
        lines.append("")
    with report_md.open("a", encoding="utf-8") as f:
        f.write("\n".join(lines))


def media_count_in_source(src: Path) -> int:
    with zipfile.ZipFile(src) as z:
        return sum(1 for n in z.namelist() if n.startswith("word/media/") and not n.endswith("/"))


def convert(src: Path, out_dir: Path) -> dict[str, int]:
    """Convert a .docx into ``out_dir``. ``out_dir`` must not exist.

    Returns a stats dict::

        {
            "comments_total":    int,
            "comments_inline":   int,
            "comments_trailer":  int,
            "media_source":      int,
            "media_output":      int,
        }

    Raises ``ValueError`` for invalid input, ``FileExistsError`` if ``out_dir``
    already exists, or ``RuntimeError`` if pandoc fails.
    """
    if not src.is_file() or src.suffix.lower() != ".docx":
        raise ValueError(f"not a .docx file: {src}")
    if out_dir.exists():
        raise FileExistsError(f"{out_dir} already exists")
    out_dir.mkdir(parents=True)

    run_pandoc(src, out_dir)
    clean_pandoc_artifacts(out_dir / "report.md")

    comments = parse_comments(src)
    inline_count = 0
    trailer_count = 0
    if comments:
        sources, range_text_for = parse_source_paragraphs(src)
        write_comment_files(comments, out_dir)
        embedded = embed_inline(out_dir / "report.md", sources, comments, range_text_for)
        unmatched = [cid for cid in comments if cid not in embedded]
        append_unmatched_trailer(out_dir / "report.md", comments, unmatched, range_text_for)
        inline_count = len(embedded)
        trailer_count = len(unmatched)

    media_dir = out_dir / "media"
    return {
        "comments_total": len(comments),
        "comments_inline": inline_count,
        "comments_trailer": trailer_count,
        "media_source": media_count_in_source(src),
        "media_output": sum(1 for _ in media_dir.iterdir()) if media_dir.exists() else 0,
    }


def main() -> int:
    ap = argparse.ArgumentParser(description="Convert .docx to clean markdown.")
    ap.add_argument("src", type=Path, help="input .docx file")
    ap.add_argument("out", type=Path, nargs="?", help="output directory (default: basename of input)")
    ap.add_argument("--force", action="store_true", help="overwrite output dir if it exists")
    args = ap.parse_args()

    src: Path = args.src.expanduser().resolve()
    out_dir: Path = (args.out or Path.cwd() / src.stem).expanduser().resolve()

    if out_dir.exists():
        if not args.force:
            sys.exit(f"error: {out_dir} exists (use --force to overwrite)")
        shutil.rmtree(out_dir)

    try:
        stats = convert(src, out_dir)
    except (ValueError, FileExistsError, RuntimeError) as e:
        sys.exit(f"error: {e}")

    if stats["comments_total"]:
        print(
            f"comments: embedded inline {stats['comments_inline']}/{stats['comments_total']}, "
            f"trailer {stats['comments_trailer']}",
            file=sys.stderr,
        )
    if stats["media_source"] != stats["media_output"]:
        print(
            f"warning: media count mismatch "
            f"(source: {stats['media_source']}, output: {stats['media_output']})",
            file=sys.stderr,
        )

    print(out_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
