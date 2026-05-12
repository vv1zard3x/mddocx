# mddocx-web

HTTP-обёртка над `mddocx`: один endpoint конвертирует `.docx` в zip
(`report.md` + `media/` + `comments/`). Swagger на `/docs`, OpenAPI на
`/openapi.json`. Веб-интерфейс с drag-and-drop на `/`.

## Запуск (локально)

```bash
docker build -t mddocx-web .
docker run --rm -p 8000:8000 mddocx-web
```

Открыть `http://localhost:8000/` (UI) или `http://localhost:8000/docs` (Swagger).

## Деплой за уже работающий Traefik

В репозитории лежит `docker-compose.yml`, рассчитанный на хост, где Traefik
уже поднят и подключён к внешней Docker-сети `proxy`. Имя cert resolver по
умолчанию — `le`, поменяй в compose если в твоём Traefik оно другое.

Перед первым запуском:

1. DNS A-record `example.com` → IP сервера.
2. В `docker-compose.yml` заменить `example.com` на свой домен (одна правка).
3. Убедиться, что внешняя Docker-сеть существует:
   `docker network ls | grep proxy`.

```bash
docker compose up -d --build
docker compose logs -f
```

## Programmatic / LLM use

Открытый API, без аутентификации. Эндпоинт самодостаточен — описание в
OpenAPI достаточно подробное, чтобы LLM могла его дёрнуть по ссылке на
`/openapi.json`.

```bash
curl -fSs -o report.zip -F "file=@report.docx" http://host:8000/convert
```

Что внутри `report.zip`:

```
report.md            # тело в GFM: текст, таблицы, списки, картинки, inline <details>-комменты
media/               # все embedded картинки из word/media/
comments/cN.md       # отдельный md на каждый Word-комментарий
```

Заголовок ответа `X-Mddocx-Stats` содержит JSON со счётчиками
(`comments_total`, `comments_inline`, `comments_trailer`, `media_source`,
`media_output`) — пригодно для проверки качества конвертации без распаковки
архива.

## MCP (Cursor / Claude Desktop)

Сервис экспонирует MCP streamable-HTTP endpoint на `/mcp/`. Один инструмент
`convert_docx(source)` — принимает HTTPS URL на `.docx` (только публичные
адреса, без localhost/private) **или** base64-закодированные байты docx, до
50 МБ. Возвращает структурированный JSON: `report_md`, `comments[]`
(`{id, author, body, range}`), `stats`, `media[]`.

**Cursor** (`~/.cursor/mcp.json` для глобального или `.cursor/mcp.json` в
корне проекта):

```json
{
  "mcpServers": {
    "mddocx": {
      "url": "https://mddocx.vv1zard3x.ru/mcp/"
    }
  }
}
```

**Claude Desktop** (`~/Library/Application Support/Claude/claude_desktop_config.json`
на macOS, аналог на других ОС):

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

После перезапуска клиента инструмент `convert_docx` появится в списке
доступных. URL обязательно **со слешем** в конце — без него FastAPI отдаст
307-редирект, который не все MCP-клиенты корректно обрабатывают на POST.

Быстрая проверка из терминала (raw JSON-RPC):

```bash
curl -sS -X POST https://mddocx.vv1zard3x.ru/mcp/ \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"curl","version":"0"}}}'
```

## Локальная разработка без Docker

Системная зависимость: `pandoc`.

```bash
sudo apt install -y pandoc
uv sync
uv run uvicorn server:app --reload --port 8000
```

## Лимиты и поведение

- Максимальный размер аплоада: 50 МБ (константа `MAX_BYTES` в `server.py`).
- Любой не-`.docx` → 400.
- Ошибки парсинга pandoc → 500 с текстом ошибки в `detail`.
- Временные файлы чистятся через `BackgroundTask` после отдачи ответа.
- Один worker uvicorn по умолчанию. Под нагрузку: `uvicorn --workers N`
  (CPU-bound через pandoc, ставить ~количество ядер).
