# Журнал посещаемости (Telegram Mini App) — без своего сервера

Фронтенд (`index.html`) лежит на GitHub Pages, «сервер» — бесплатный Google Apps Script, база — Google-таблица.
Python, Docker, Railway и Volume больше не нужны.

## Запуск

### 1. Бэкенд (Google)
1. Создайте новую Google-таблицу (sheets.new).
2. Меню **Расширения → Apps Script**. Удалите содержимое файла и вставьте код из `Code.gs`. Сохраните.
3. **Project Settings (шестерёнка) → Script Properties → Add**:
   - `BOT_TOKEN` — токен от @BotFather
   - `OWNER_IDS` — ваш Telegram ID (узнать можно у бота @userinfobot), несколько — через запятую
   - `TZ_NAME` — например `Asia/Novosibirsk` (необязательно)
   - `SCHEDULE_URL_TEMPLATE`, `MARK_BEFORE_MIN`, `MARK_AFTER_MIN` — необязательно, как раньше
4. **Deploy → New deployment → тип Web app**: Execute as — **Me**, Who has access — **Anyone**. Нажмите Deploy, разрешите доступ.
5. Скопируйте адрес вида `https://script.google.com/macros/s/.../exec`.

### 2. Фронтенд (GitHub Pages)
1. В `index.html` найдите `API_URL` и вставьте адрес из шага 1.5.
2. Загрузите `index.html` в репозиторий (в корень ветки, которую раздаёт Pages).

### 3. Telegram
В @BotFather: `/mybots` → ваш бот → **Bot Settings → Menu Button → Configure menu button** → вставьте ссылку на GitHub Pages (`https://…github.io/…/`).
Либо `/newapp` — получите ссылку вида `t.me/бот/приложение`.

Процесс бота больше не нужен: Mini App открывается по кнопке меню, а команды `/start`, `/id` без Python-бота не отвечают.

## Важно
- **Обновление кода:** после правок в Apps Script делайте **Deploy → Manage deployments → Edit → New version**, иначе работает старая версия.
- **Выгрузка:** вместо `/export` — меню **Журнал → Обновить отчёты по группам** в таблице (листы «Отчёт …»; `+` был, `н` не был).
- **Бэкап:** Google-таблица хранит историю версий (Файл → История версий), листы `grps`, `students`, `lessons`, `marks` — это и есть база. Не редактируйте их вручную без необходимости.
- **ICS:** поддерживаются обычные события и повторы DAILY/WEEKLY. Более сложные правила повторов (MONTHLY и т. п.) берутся как одна пара.
- **Лимиты Google:** бесплатно хватает на группы и потоки в сотни студентов; при очень массовых одновременных отметках возможны задержки в пару секунд.
- Подпись Telegram (`initData`) проверяется на стороне Apps Script по токену бота, подделать отметку без Telegram нельзя.
