# Markdown rules for docxmd conversion

Use this contract when producing markdown that will be converted to a GOST-formatted `.docx` via docxmd.

One `.md` file per document. Images go in `media/`. Image paths in the markdown are relative to the folder containing the `.md`.

## Headings
- `#` document title, exactly one per file.
- `##`, `###`, ... for sections.
- Never include section numbers in the heading text. The converter does not number headings.

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
