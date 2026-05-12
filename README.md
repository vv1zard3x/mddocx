# mddocx-web

Двусторонний мост между Word `.docx` и Markdown в одном сервисе.

- `POST /convert` — `.docx` → `.zip` с `report.md`, `media/`, `comments/`.
- `POST /convert/md` — `.md` или `.zip(md+media/)` → `.docx` по ГОСТ 7.32
  (Times New Roman 12, поля 30/15/20/20 мм, полуторный интервал,
  заголовок 1-го уровня по центру, остальные по левому краю,
  автоподписи рисунков «Рисунок N. …»).
- `GET  /rules` — контракт markdown для второго направления (готов к
  копированию в system-промпт LLM).
- `GET  /template.zip` — эталонный `template.md + media/`, который сам
  чисто проходит через `/convert/md`.
- `*    /mcp/` — MCP streamable-HTTP с двумя инструментами:
  `convert_docx`, `convert_md`.
- `GET  /docs`, `GET  /openapi.json` — Swagger / схема.
- `GET  /` — drag-and-drop UI с тоггл-переключателем направления.

## Запуск (локально)

```bash
docker build -t mddocx-web .
docker run --rm -p 8000:8000 mddocx-web
```

Открыть `http://localhost:8000/` (UI), `/rules` (правила md), `/docs`
(Swagger).

## Деплой за уже работающий Traefik

В репозитории лежит `docker-compose.yml`, рассчитанный на хост, где
Traefik уже поднят и подключён к внешней Docker-сети `proxy`. Имя cert
resolver по умолчанию — `le`, поменяй в compose если в твоём Traefik
оно другое.

Перед первым запуском:

1. DNS A-record `example.com` → IP сервера.
2. В `docker-compose.yml` заменить `example.com` на свой домен (одна правка).
3. Убедиться, что внешняя Docker-сеть существует: `docker network ls | grep proxy`.

```bash
docker compose up -d --build
docker compose logs -f
```

## REST API

### DOCX → Markdown

```bash
curl -fSs -o report.zip -F "file=@report.docx" \
  http://host:8000/convert
```

Что внутри `report.zip`:

```
report.md            # тело в GFM: текст, таблицы, списки, картинки, inline <details>-комменты
media/               # все embedded картинки из word/media/
comments/cN.md       # отдельный md на каждый Word-комментарий
```

Заголовок ответа `X-Mddocx-Stats` содержит JSON со счётчиками
(`comments_total`, `comments_inline`, `comments_trailer`,
`media_source`, `media_output`).

### Markdown → DOCX

Одиночный `.md`:

```bash
curl -fSs -o out.docx -F "file=@document.md" \
  http://host:8000/convert/md
```

Папка с медиа (упаковать в zip и отправить как один файл):

```bash
cd my_doc && zip -r ../pack.zip . && cd ..
curl -fSs -o out.docx -F "file=@pack.zip" \
  http://host:8000/convert/md
```

Структура архива: `document.md` в корне (или вложен в одну
поддиректорию, она будет вынута) + опциональная папка `media/`.
Допускается ровно один `.md`. Структура и допустимые конструкции
markdown описаны на `/rules`.

## MCP (Cursor / Claude Desktop)

Сервис экспонирует MCP streamable-HTTP endpoint на `/mcp/` с двумя
инструментами:

- `convert_docx(source)` — `.docx` → `report_md` + структурированные
  комментарии. `source` — HTTPS URL или base64 (≤ 50 МБ).
- `convert_md(source, filename="document.docx")` — markdown (с
  опционально упакованной медиатекой) → `.docx` по ГОСТ. `source` —
  один из:
  - plain UTF-8 текст markdown,
  - base64-zip с `*.md` + `media/`,
  - HTTPS URL на `.md` или `.zip`.

  Ответ: `{"filename": ..., "content_base64": ..., "size_bytes": ...}`.

**Cursor** (`~/.cursor/mcp.json` или `.cursor/mcp.json` в проекте):

```json
{
  "mcpServers": {
    "mddocx": {
      "url": "https://mddocx.vv1zard3x.ru/mcp/"
    }
  }
}
```

**Claude Desktop** (`~/Library/Application Support/Claude/claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "mddocx": {
      "type": "http",
      "url": "https://mddocx.vv1zard3x.ru/mcp/"
    }
  }
}
```

URL обязательно **со слешем** в конце — без него FastAPI отдаёт
307-редирект, который не все MCP-клиенты корректно обрабатывают на POST.

После перезапуска клиента в списке появятся `convert_docx` и
`convert_md`. Конфиг отображается прямо на главной странице сервиса
(подставляется текущий домен).

Быстрая проверка из терминала:

```bash
curl -sS -X POST https://mddocx.vv1zard3x.ru/mcp/ \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"curl","version":"0"}}}'
```

## Локальная разработка без Docker

Системная зависимость: `pandoc` (нужен только для docx → md;
направление md → docx работает на чистом Python).

```bash
sudo apt install -y pandoc
uv sync
uv run uvicorn server:app --reload --port 8000
```

## Лимиты и поведение

- Максимальный размер аплоада: 50 МБ (`MAX_BYTES` в `server.py`).
- `POST /convert`: не-`.docx` → 400, ошибки pandoc → 500.
- `POST /convert/md`: не-`.md`/`.markdown`/`.zip` → 400, zip с
  traversal/абсолютными путями → 400, более одного `.md` в zip → 400.
- Временные файлы чистятся через `BackgroundTask` после отдачи ответа.
- Один worker uvicorn по умолчанию. Под нагрузку: `uvicorn --workers N`.

## Структура

```
mddocx_web/
├── server.py          # FastAPI: REST + MCP
├── mddocx.py          # ядро docx → md (pandoc + XML)
├── docxmd/            # пакет md → docx (вендорная копия ~/projects/tools/docxmd)
│   ├── __init__.py
│   ├── _styles.py     # ГОСТ-стили
│   └── _convert.py    # walker по токенам markdown-it
├── MARKDOWN_RULES.md  # источник контракта md (рендерится на /rules)
├── static/
│   ├── index.html     # UI с тоггл-переключателем
│   ├── rules.html     # layout для /rules
│   └── template/      # template.md + media/ для /template.zip
├── Dockerfile
└── docker-compose.yml
```

`docxmd/` — вендорная копия независимого пакета
`~/projects/tools/docxmd/`. Источник правды для md→docx живёт там;
при изменении логики синхронизировать вручную (`cp -r ~/projects/tools/docxmd/docxmd
~/projects/tools/mddocx_web/docxmd && cp ~/projects/tools/docxmd/MARKDOWN_RULES.md
~/projects/tools/mddocx_web/`).
