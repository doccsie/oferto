# Оферто — MVP

Сервис превращает Excel-реестр исполнителей в пакет заполненных договоров-оферт (.docx), упакованных в ZIP. Вся обработка происходит **в браузере клиента** — сервера нет, данные никуда не отправляются. Поэтому проект хостится как статический сайт на GitHub Pages.

## Структура

```
index.html            — главная (о сервисе)
app.html              — личный кабинет (без авторизации)
assets/style.css      — стили
assets/app.js         — парсинг xlsx → генерация docx → zip
templates/registry_example.xlsx   — пример реестра для клиентов
templates/contract_template.docx  — шаблон договора с плейсхолдерами {FIO}, {INN}, …
vendor/               — библиотеки (SheetJS, docxtemplater, PizZip, JSZip, FileSaver)
test/e2e.js           — e2e-тест (playwright): загрузка реестра → ZIP
```

## Как задеплоить на GitHub Pages

1. Создай репозиторий на GitHub (например, `oferto`).
2. В папке проекта:
   ```bash
   git remote add origin https://github.com/<твой-логин>/oferto.git
   git push -u origin main --tags
   ```
3. В репозитории: **Settings → Pages → Source: Deploy from a branch → Branch: `main` / `(root)` → Save**.
4. Через минуту сайт будет доступен на `https://<твой-логин>.github.io/oferto/`.

## Версионирование

Версии подписываются git-тегами (`v0.1.0`, `v0.2.0`, …) и дублируются штампом в подвале сайта.
Новая версия: внеси изменения → обнови номер в подвале `index.html` и `app.html` → закоммить → `git tag v0.x.0` → `git push --tags`.

## Как это работает

1. Клиент скачивает в ЛК пример реестра (14 колонок: ФИО, паспорт, ИНН, СНИЛС, статус, специальность, email, банковские реквизиты).
2. Загружает заполненный файл — заголовки распознаются по названиям (устойчиво к опечаткам и порядку колонок), данные показываются в таблице предпросмотра с отметкой незаполненных полей.
3. По кнопке на каждого исполнителя рендерится `contract_template.docx` (docxtemplater заменяет плейсхолдеры), все договоры отдаются одним ZIP: `Договор_<ФИО>.docx`.

## Локальная разработка

```bash
python3 -m http.server 8901   # или любой статический сервер
# e2e-тест (нужен playwright + chromium):
node test/e2e.js templates/registry_example.xlsx
```

## Обновление шаблона договора

Замени `templates/contract_template.docx`. Плейсхолдеры в фигурных скобках: `{FIO}`, `{Status}`, `{Registration}`, `{INN}`, `{SNILS}`, `{SerialNumbPas}`, `{Issued}`, `{IssuedDate}`, `{Specialty}`, `{EMAIL}`, `{AccNumb}`, `{BANK}`, `{BIC}`, `{CorrAcc}`.
