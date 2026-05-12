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
уже поднят и подключён к внешней Docker-сети. По умолчанию имя сети —
`traefik`, имя cert resolver — `le`. Поменяй их в compose, если в твоём
Traefik они называются иначе.

Перед первым запуском:

1. DNS A-record `example.com` → IP сервера.
2. В `docker-compose.yml` заменить `example.com` на свой домен (одна правка).
3. Убедиться, что внешняя Docker-сеть существует:
   `docker network ls | grep traefik`.

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
