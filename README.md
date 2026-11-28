# Парсер вакансий Telegram (облачная версия)

Читает публичные Telegram-каналы через веб-версию t.me/s/<канал>, отбирает вакансии
и заказы по темам (ИИ, сайты, боты и автоматизация, внедрение), убирает дубли и шлёт
лучшие (оценка 7+) в Telegram-бот. Работает в GitHub Actions, данные хранит в Supabase.

Это перенос локального парсера «Вайб-сайты\Парсер Telegram-вакансий» в облако:
логика чтения, фильтров и оценок — та же (`it_monitor.py`), поменяно только хранение.

## Как работает

- **Расписание:** запуск обоих workflow происходит через Supabase (pg_cron + pg_net):
  функция `dispatch_watchers()` каждые 5 минут проверяет окно 06:00–20:59 по Москве
  (03:00–17:59 UTC) и вызывает GitHub API `workflow_dispatch` для `kwork.yml`
  (каждые 5 минут) и `vacancies.yml` (в чётные десятки минут). В самих workflow
  блок `schedule:` убран — GitHub-расписание для свежего репозитория не стартовало.
  Токен GitHub хранится в таблице `gh_token` (RLS без политик, читается только
  сервисным ключом). Плановые прогоны видны во вкладке Actions с событием
  `workflow_dispatch`.
- **Две группы:** каналы делятся по чётности позиции в `channels.txt`, за запуск
  обрабатывается половина (чередование по десяткам минут UTC), как у локальной задачи.
- **Хранение:** состояние каналов — таблица `vp_state`, все посты — `vp_posts`
  (Supabase, доступ только по секретному ключу, RLS включён).
- **Уведомления:** напрямую в бот через Bot API; не отправленное из-за сети доотправится
  в следующем прогоне (посты остаются `notified=false`).
- **Дедупликация:** отпечаток текста (хэш набора слов) против всей базы + jaccard
  по шинглам внутри прогона, как в локальной версии.
- **Таблица xlsx в облаке не строится** — посты копятся в Supabase, выгрузку можно
  добавить позже.

## Таблицы Supabase

```sql
create table if not exists vp_posts (
  post_url text primary key,
  channel text not null,
  post_id bigint not null,
  date timestamptz,
  text text,
  links jsonb default '[]'::jsonb,
  buttons jsonb default '[]'::jsonb,
  fingerprint text,
  notified boolean not null default false,
  created_at timestamptz not null default now()
);
create table if not exists vp_state (
  channel text primary key,
  last_seen bigint not null default 0,
  updated_at timestamptz
);
alter table vp_posts enable row level security;
alter table vp_state enable row level security;
```

RLS включён без политик: анонимные запросы не видят данные, работает только
сервисный ключ (он в секретах GitHub, в репозиторий не попадает).

## Секреты репозитория

| Секрет | Что это |
|---|---|
| `SUPABASE_URL` | адрес проекта Supabase (Project URL) |
| `SUPABASE_SERVICE_KEY` | сервисный ключ (secret) — держать только в секретах |
| `TG_BOT_TOKEN` | токен бота для уведомлений |
| `TG_CHAT_ID` | чат, куда приходят уведомления |

## Файлы

- `cloud_monitor.py` — облачный монитор вакансий (запускается Actions'ом).
- `kwork_watch.py` — облачный сторож заказов Kwork (тот же бот, таблица `kw_orders`;
  правила отбора и формат уведомлений — по навыку kwork-notify; отклики не отправляются).
- `it_monitor.py`, `fetch_posts.py`, `label_posts.py`, `export_table.py` — модули
  локального парсера (используются как библиотека: темы, оценки, фильтры, парсер страниц).
- `channels.txt` — список каналов (один на строку; добавить канал = добавить строку).
- `seed_supabase.py` — разовый перенос старой базы `it_posts.json`/`it_state.json`
  в Supabase (запуск на своём компьютере, переменные `SUPABASE_URL`/`SUPABASE_SERVICE_KEY`).

## Проверить вручную

Вкладка Actions → «vacancies» → Run workflow. Лог покажет группы каналов,
HTTP-коды, статистику отсева и отправку в бот.
