# Markdown rules for docxmd conversion

Use this contract when producing markdown that will be converted to a GOST-formatted `.docx` via docxmd.

One `.md` file per document. Images go in `media/`. Image paths in the markdown are relative to the folder containing the `.md`.

## Headings
- `#` (H1) marks a GOST 7.32 structural section: `Введение`, `Заключение`, `Список использованных источников`, `Приложение А`, `Реферат`, `Содержание`. Rendered centered in UPPERCASE, never numbered. Multiple H1s per file are allowed.
- The document title proper is not written in markdown — it lives on the Word title page.
- `##` (H2) marks a numbered top-level section. The converter prefixes it automatically: `1`, `2`, `3`, ... across the document. Structural H1s do not reset this counter.
- `###`..`######` are nested subsections; auto-prefixed as `1.1`, `1.1.1`, `1.1.1.1`, ... Cascade restart happens automatically when a higher level appears.
- Never type section numbers inside the heading text — Word numbers headings itself. Manual prefixes will produce double numbering.
- Do not skip heading levels (`##` followed directly by `####` without `###`).

## Paragraphs
- One blank line between paragraphs.
- To force a line break inside a paragraph, end the line with two spaces.
- No HTML tags. No YAML front matter.

## Lists
- Bulleted: `- item`.
- Ordered: `1. item`. The digits in the source do not matter; numbering is hierarchical and rendered by Word.
- Nest with 2-space indent. Nested ordered items become `1.1.`, `1.1.1.`, etc.
- Reference and bibliography lists must be ordered (per GOST 7.32), never bulleted.

## Images
- Standalone paragraph: `![Caption text](media/file.png)`. Alt text becomes `Рисунок N — Caption text` below the image.
- For inline icons inside a paragraph: leave the alt empty `![](media/icon.png)`. No caption is rendered.

## Tables
- GFM pipe syntax with a header row and a `---` separator. Column alignment via `:---`, `:---:`, `---:`.
- Caption above the table: a single paragraph starting with `Таблица:` placed immediately before the table. Becomes `Таблица N — Title`. Skip the prefix to skip the caption.

## Blockquotes
- `> quoted text`. Nested with `>>`.

## Code
- Inline: `` `code` ``.
- Block: triple backticks. Optional language tag is ignored.

## Links
- `[text](url)`. Do not put `**bold**` or `*italic*` inside the brackets — formatting inside link text is dropped.

## Forbidden
HTML tags, footnotes (`[^1]`), task lists (`- [x]`), inline math, formula blocks, TOC markers, embedded YAML/TOML metadata.
