# Карта проекта (Project Map)

**Актуально для:** W-ANALYTIC-06E (Фикс ведущего нуля в номерах шагов) | **Дата:** 2026-11-19

## VBA-модули (корень репозитория, формат .bas)

| Модуль | Назначение | Ключевые функции/процедуры |
|---|---|---|
| Mod_Util.bas | Журналирование, конфигурация, настройки, пути | WriteLog ("[i] <текст>"), WriteError ("ERROR [<код>]: <текст>"), WriteWarning ("[!] <текст>"), DebugLog, FormatMissingFile, FindConfigFile (.xml > .txt; только .pilotcfx — предупреждение), GetConfigValue, **GetProjectRoot** (читает ProjectRoot из «Настроек», пусто = ThisWorkbook.Path), **GetSettingLong(name, defaultValue)** (предупреждение U007 при пустом/некорректном значении), **GetSettingText(name, defaultValue)** (пусто → default + U007), **LegendRange** (именованный диапазон на лист «Настройки», B8), **MaxTypesPerTable** (поле B7, именованный диапазон, читается через GetSettingLong), **EnsureSettingsSheet** (6 полей B2:B7: ProjectRoot, CheckMode, AutoOpenResults, PromptGenVersion, PromptFixVersion, MaxTypesPerTable + 7-е поле B8 «Диапазон легенды» / LegendRange; Validation на B3/B4/B7; именованные диапазоны всех 6 полей; защита «pilot», разблокированы B2:B8; вызов Mod_Patterns.EnsurePatternsSheet с обёрткой On Error Resume Next + U008) |
| Mod_Folders.bas | Структура папок и проверка исходников | CreateProjectFolders (создаёт в т.ч. "01_Исходные данные\04_Дополнительные данные"), CheckSourceFolders (Схема .xlsx / Описание .docx / Конфигурация через FindConfigFile / сводка по доп. данным; возвращает True/False), GetStepFolder |
| Mod_Structure.bas | Парсинг схемы ЖЦ ОКС в адаптированную структуру | BuildFlatStructure (динамическая легенда: `GetSettingText("LegendRange", DEFAULT_LEGEND_RANGE)` вместо хардкода `legCells(1 To 6)`; при некорректном адресе — S011 + fallback B1:C3; восстановление prevScreenUpdating/prevDisplayAlerts/prevStatusBar, закрытие wbOut/wbSrc во всех ветках Fail/OpenErr/saveErr), **WriteFlatTable** (лист «Адаптированная_структура»; **FLAT_COLS = 7**, колонки A–G: «№шага/подшага», «Отображаемое имя», «Категория» (бывш. «Тип»), «Паттерн», «Входит в», «Состоит из», «Аномалия»; умная таблица `tblFlatStructure` A1:G<lastRow>; подсветка A:G; AutoFit A:G), WriteLog (лист «Аномалии схемы»; заголовки через Worksheet.Range("A1:H1")), SafeAutoFit, WriteTraceability (заголовок «Тип» → «Категория»), CheckNumberDuplicates, HighlightDuplicateRows (колонка «Аномалия» в диапазоне A:G), BackupFlatTableSheet, ArchiveOldSchemaArtifacts |
| Mod_Prompt.bas | Генерация промтов | **BuildPromptForStep** (общая цепочка сборки промта шага по prompt_step_template_v1.md), **ApplyCorrectionsBlock**, **FillAllPlaceholders**, **BuildStepTypeList** (распознаёт новые канонические заголовки «Категория» и «№шага/подшага» наравне со старыми «Тип»/«Шаг» — обратная совместимость; буквенные индексы колонок не используются), **GenerateStepZeroPrompt** (устаревшая ветка: заглушка с MsgBox о замене — W7.1-C, шаблон перемещён в prompts/archive/), **MakeNextStepPrompt**, MakeFixPrompt (не подключён — W-FIX-PROMPT), ReadStepList, ListFolderFiles, …DECOMPOSITION_RULES_V713. **GetPromptProfile** (W7.1-A: чистая функция без MsgBox, маршрутизация по lastStep: "" → ""; нет/00.i → «Папки»; 0/00/01.i/01.i.k → «Паттерны»; 01/NN≥2 → «Документы»; невалидное → ""; ответственность за "" — на вызывающем). **GetFileListArray** (W-ANALYTIC-04-FIX: сбор имён файлов папки в массив одноразовым проходом Dir() до начала обработки — защита от сброса stateful-итератора). **GenerateAnalyticalPrompts_Click** (W-ANALYTIC-04-FIX: обход файлов папок 02_Описание/03_Конфигурация/04_Дополнительные данные через массив `GetFileListArray` + `For`, проверка существования промта — `FileExistsSafe`; за один запуск создаёт промты для всех файлов). **CommandGenerateTables** (W-ANALYTIC-05, кнопка 13: выбор шаблона 00_2/00_3 по настройке PatternsState, номер итерации через GetNextIterationNumber, подстановка плейсхолдеров актуальных шаблонов W-ANALYTIC-01, {LAST_STEP_NUMBER} из настройки LastStepNumber (W-ANALYTIC-09), валидация профиля через GetPromptProfile, сохранение Prompt_00_X_iter<N>.md в 02_Подготовка данных\05_Фрагменты шагов). **GetNextIterationNumber** (Private, W-ANALYTIC-05: сканирование 05_Фрагменты шагов и Архив через GetFileListArray, извлечение max N из имён Prompt_00_X_iter*.md, возврат max+1).   **CommandParseAiResponse** — реализован полный цикл: маркерный парсинг 00_2/00_3, обновление «Адаптированной структуры», архивация обработанных файлов и промтов оператором Name (с обработкой коллизий), обновление настроек LastStepNumber (посегментное сравнение) и StepsState. |
| Mod_Assembler.bas | Сборка итоговых настроек | сборка файлов настроек по результатам шагов |
| Mod_Docs.bas | Документация проекта | генерация отчётов/документов |
| Mod_Registry.bas | Реестры | Реестр_типов.txt, Реестр_покрытия.txt |
| Mod_TestGen.bas, Mod_TestGen_Assembler.bas | Генерация тестов | тестовые сценарии для собранных настроек |
| Mod_UI.bas | Кнопки/обработчики интерфейса | BuildFlatStructure_Click и др.; **RestoreUI создаёт 14 кнопок** на листе «Препроцессор»; кнопка 13 «Сформировать таблицы» → **CommandGenerateTables** (W-ANALYTIC-05: реализация в Mod_Prompt; заглушка S0 из Mod_UI удалена), кнопка 14 «Разобрать ответ ИИ» → **CommandParseAiResponse** (заглушка в Mod_Prompt до W-ANALYTIC-06; заглушка из Mod_UI удалена — иначе «Ambiguous name detected» при неквалифицированном OnAction; вызов только по кнопке 14, не из MakeNextStepPrompt; запись в лист «Паттерны» — W7.3-PATTERNS-UPDATE) |
| Mod_Patterns.bas | Лист «Паттерны» (PRE-04c): каркас для Промта №0.2 | **EnsurePatternsSheet** (идемпотентна: создание листа в конце книги, **без Cells.Clear** на существующем листе — PRE-04c-fix; секция A–D: заголовки в строке 2 (A2 «Паттерн», B2 «Типы», C2 «Количество», D2 «Итерация»), данные с 3-й; секция H–O: заголовки в строке 2, данные с 3-й, **8 предзаполненных правил** — Папки ЦИМ, Год (Маска/Год*), маска bim_*, точные системные типы document/ecmDocument/folder/project/projectfolder; Validation H/I/J/L/O на строках 3–1000; заголовки A1/H1 — названия секций (жирным, объединены A1:D1, H1:O1); защита паролем «pilot», разблокированы A2:D1000 и H3:O1000; MigrateLegacyLayout, ShiftLegacyYearRule, SyncPrefillCell для миграции старых каркасов); вызывается из Mod_Util.EnsureSettingsSheet; коды PT001/PT002/PT003. **GetPatternExceptions** (W7.1-A: секция H–O, фильтр O = «активно», PT003 при отсутствии листа/пустой секции) и **GetKnownPatterns** (W7.1-A: секция A–D, колонка A, заголовок «Известные паттерны:» внутри функции, "" при пустой секции, PT003 при отсутствии листа); лист не изменяется, защита не снимается. H–O предзаполнено 8 правилами, A–D пополняется по итерациям (пустая A–D — защитный сценарий, маркер не вводится). **W7.1a: PATTERNS_TABLE_NAME = "tblPatterns" (Public Const), EnsurePatternsListObject** — идемпотентное создание/расширение умной таблицы A2:O1000, стиль TableStyleMedium6, вызов из EnsurePatternsSheet до установки защиты |



## Папки рабочего проекта (создаются CreateProjectFolders вне репозитория)

```text
<корень .xlsm>/
├─ 01_Исходные данные/
│   ├── 01_Схема/
│   ├── 02_Описание/              ← промты 00_4 (Prompt_*.md)
│   ├── 03_Конфигурация/          ← промт 00_1 (Prompt_*.md)
│   └── 04_Дополнительные данные/ ← промты 00_4 (Prompt_*.md)
│
├─ 02_Подготовка данных/
│   ├── 01_Схема/
│   ├── 02_Описание/              ← выжимки (результаты 00_4)
│   ├── 03_Конфигурация/
│   ├── 04_Дополнительные данные/ ← выжимки (результаты 00_4)
│   ├── 05_Фрагменты шагов/       ← промты 00_2, 00_3 + ответы ИИ (таблицы)
│   └── Архив/                    ← обработанные ответы + промты
│
├─ 03_Результаты генерации/
│   ├── 01_По шагам/
│   │   ├── Шаг_00/
│   │   ├── Шаг_00.1/
│   │   ├── Шаг_01/
│   │   └── ...
│   ├── 02_Сводные/
│   ├── 03_Реестры/
│   └── 04_Промты_готовые/
│
├─ 04_Промты/
│   ├── Шаблоны/
│   │   ├── prompt_step_00_1_template.md
│   │   ├── prompt_step_00_2_template.md
│   │   ├── prompt_step_00_3_template.md
│   │   ├── prompt_step_00_4_template.md
│   │   ├── prompt_step_template_v1.md
│   │   ├── prompt_fix_v1.0.md
│   │   └── prompt_settings_template.md
│   └── project_settings.md
│
└─ 05_Логи/                     (Препроцессор_YYYY-MM-DD.log)


```

## Формат Шаги.txt

Канонический разделитель — pipe: `NN|Имя_папки` (например `01|Шаг_01_Скелет`).
Разделитель "=" поддерживается ReadStepList обратно совместимо, но помечается предупреждением.

## Сообщения (единый формат, W-CORRECTION-PACK-1)

- ERROR: `ERROR [<код>]: <текст>`
- WARN:  `[!] <текст>`
- INFO:  `[i] <текст>`
- Отсутствие обязательного файла: `[!] Отсутствует обязательный файл: <имя>`
- Код F005: исходные данные не готовы — генерация промтов остановлена.

## Известные подводные камни / соглашения VBA

- ListObject.HeaderRange после Resize может выдавать ошибку 438. Для чтения заголовков использовать Worksheet.Range("A1:H1").
- Внутри циклов `For i = 1 To N` в обработчиках ошибок использовать именованные метки (`Resume NextNodeLog`) вместо `Resume Next`, чтобы переходить к следующей итерации цикла, а не к следующей инструкции.
- В VBA строковые литералы с кавычками внутри требуют экранирования (""), иначе Syntax error. Проверять компиляцию (Debug -> Compile) перед коммитом.
- При присвоении результата Split() динамическому массиву использовать Variant (Private arr As Variant), иначе ошибка компиляции.
- Все Private Const должны находиться до первой процедуры Sub/Function в модуле.
- Глобальный итератор `Dir()`: любой вызов `Dir(<путь>)` с аргументом сбрасывает единственный stateful-итератор VBA; внутри `Do While fName <> "" ... fName = Dir()` (в т.ч. во вложенных вызовах `SaveTextFileUtf8` → `StripUtf8Bom` → `Dir(filePath, vbNormal)`) цикл молча завершается после первой итерации. Обход файлов — только по массиву (`Mod_Prompt.GetFileListArray`), проверка существования — `FileExistsSafe` (GetAttr). (W-ANALYTIC-04-FIX)
- Дублирование `Public Sub` с одинаковым именем в разных модулях + неквалифицированный `OnAction` кнопки Form Control → runtime-ошибка «Ambiguous name detected» (W-ANALYTIC-05). Решение: единственная реализация в профильном модуле (`Mod_Prompt`), в UI-модуле (`Mod_UI`) — только точка входа без дублирующего имени.

## Карта промтов (templates/prompts — папка prompts/ в репозитории)

Источник: отчёт QWEN (секция 2) с уточнениями W-DOCS-MAP.

| Шаблон | Назначение | Плейсхолдеры (полный список) | Кто заполняет (процедура VBA) | Где вызывается | Статус |
|---|---|---|---|---|---|
| `prompt_step_00_template.md` | Декомпозиция (старая v7.13) | `{SCHEMA_FILENAME}`, `{DESCRIPTION_FILENAME}`, `{CONFIG_FILENAME}`, `{EXTRA_DATA_LIST}`, `{Decomposition_Rules}` | `GenerateStepZeroPrompt` | `MakeNextStepPrompt` | устаревший; перемещён в `prompts/archive/` (W7.1-C) |
| `prompt_step_00_1_template.md` | Базовые типы из Конфигурации | `{CONFIG_FILENAME}` (вход), `{{СПИСОК_БАЗОВЫХ_ТИПОВ}}` (выход) | отсутствует | не вызывается | не подключён (ручной) |
| `prompt_step_00_2_template.md` | Паттерны и типы-справочники из Схемы | `{SCHEMA_FILENAME}`, `{DESCRIPTION_FILENAME}`, `{EXTRA_DATA_LIST}`, `{{СПИСОК_ИЗВЕСТНЫХ_ПАТТЕРНОВ}}`, `{{ИСКЛЮЧЕНИЯ_ПАТТЕРНОВ}}` | `CommandGenerateTables` (W-ANALYTIC-05) | кнопка 13 «Сформировать таблицы» | работает (W-ANALYTIC-05); плейсхолдеры приведены к актуальному шаблону W-ANALYTIC-01 |
| `prompt_step_00_3_template.md` | Список шагов (3 профиля: Папки / Паттерны / Документы) | `{LAST_STEP_NUMBER}`, `{INPUT_FILENAME}`, `{DESCRIPTION_FILENAME}`, `{EXTRA_DATA_LIST}`, `{ЛИМИТ_ТИПОВ_В_ТАБЛИЦЕ}` | `CommandGenerateTables` (W-ANALYTIC-05) | кнопка 13 «Сформировать таблицы» | работает (W-ANALYTIC-05); плейсхолдеры приведены к актуальному шаблону W-ANALYTIC-01 |
| `prompt_step_00_4_template.md` | Выжимка ГОСТа и Описания (PRE-05) | `{DOC_FILENAME}`, `{DESCRIPTION_FILENAME}` (одинарные) | отсутствует | не вызывается | не подключён (ручной) |
| `prompt_step_template_v1.md` | Инвариантный шаблон шага NN (17 блоков; блок 16 — корректировки) | `{{NN}}`, `{{ИМЯ_ШАГА}}`, `{{ЗАДАЧА_ШАГА_ОСОБЕННОСТИ}}`, `{{СТРУКТУРА_ФАЙЛ}}`, `{{ОПИСАНИЕ_ФАЙЛ}}`, `{{КОНФИГУРАЦИЯ_ФАЙЛ}}`, `{{ФАЙЛЫ_ДОП_ДАННЫХ}}`, `{{ШАГИ_ФАЙЛ}}`, `{{СПИСОК_ВХОДОВ}}`, `{{ТАБЛИЦА_ИЗВЕСТНЫХ_GUID}}`, `{{РЕЕСТР_ТИПОВ}}`, `{{РЕЕСТР_ПОКРЫТИЯ}}`, `{{СПИСОК_БАЗОВЫХ_ТИПОВ}}`, `{{ПРАВИЛО_ОБЪЕКТА}}`, `{{СПИСОК_ТИПОВ_ШАГА}}`, `{{ТЕКУЩИЙ_ШАГ}}`, `{{ДЕКОМПОЗИЦИЯ_ШАГОВ}}`, `{{ФОРМАТ_РЕЕСТРОВ}}`, `{{ПРОТОКОЛ_ШАГА}}`, `{{МАРКЕРЫ}}` | `BuildPromptForStep` / `FillAllPlaceholders` | `MakeNextStepPrompt` | работает |
| `prompt_fix_v1.0.md` | Исправление ошибок шага | `{NN}`, `{ERRORS}`, `{REGISTRY_TYPES}`, `{REGISTRY_COVERAGE}`, `{SCHEMA}` | `MakeFixPrompt` | `MakeFixPrompt_Click` | заготовка; плейсхолдеры `{NN}`, `{ERRORS}`, `{REGISTRY_TYPES}`, `{REGISTRY_COVERAGE}`, `{SCHEMA}` будут добавлены вместе с этапом исправления ошибок генерации (W-FIX-PROMPT, после W7.2); до этого MakeFixPrompt не использовать |
| `prompt_generator_v7.13.md` | Старый генератор | Н/Д | отсутствует | не вызывается | мёртвая ветка |
| `prompt_settings_template.md` | Памятка по настройкам (8 ключей project_settings.md) | Н/Д (текстовый файл) | не читается кодом | не вызывается | шаблон-памятка; в генерации промта не участвует |
| `project_settings.md.example` | Образец рабочего файла настроек `project_settings.md` (8 ключей) | `{{СПИСОК_ВХОДОВ}}`, `{{ТАБЛИЦА_ИЗВЕСТНЫХ_GUID}}`, `{{СПИСОК_БАЗОВЫХ_ТИПОВ}}`, `{{ПРАВИЛО_ОБЪЕКТА}}`, `{{ФОРМАТ_РЕЕСТРОВ}}`, `{{ПРОТОКОЛ_ШАГА}}`, `{{МАРКЕРЫ}}` | `LoadProjectSettings` | `FillAllPlaceholders` | работает |
| `reference/attribute_params_examples.md` | Справочник образцов «Дополнительных параметров» (НЕ инвариант) | Н/Д | не читается кодом | не вызывается | справочник |

Внешняя папка "Шаблоны" (копия шаблонов у Пользователя) в репозиторий НЕ входит
и обновляется Пользователем вручную после принятия PR.

## Поток данных

1. Проверка/создание папок: Mod_Folders.CreateProjectFolders, CheckSourceFolders.
2. Подготовка схемы: Mod_Structure.BuildFlatStructure (7 колонок A–G, динамическая легенда из B8; дубли → «Аномалии схемы», бэкап листа перед перезаписью).
3. Реестры: Mod_Registry.
4. Промты: Mod_Prompt (№0 → Шаги.txt → промты шагов → фиксы); перед генерацией — CheckSourceFolders (F005).
5. Сборка: Mod_Assembler; тесты: Mod_TestGen*.

## Поток данных аналитических промтов

1. Кнопка «Сгенерировать аналитические промты»:
   - Читает файлы из `01_Исходные данные/02_Описание`, `01_Исходные данные/03_Конфигурация`, `01_Исходные данные/04_Дополнительные данные`.
   - Генерирует промты 00_1 и 00_4.
   - Ставит «Выжимки и анализ конфигурации» = «Да».

2. Пользователь отправляет промты ИИ, сохраняет:
   - Выжимки → `02_Подготовка данных/02_Описание` и `02_Подготовка данных/04_Дополнительные данные`.
   - Таблицы → `02_Подготовка данных/05_Фрагменты шагов`.

3. Кнопка «Сформировать таблицы»:
   - Генерирует промт 00_2 или 00_3 в `02_Подготовка данных/05_Фрагменты шагов`.

4. Кнопка «Разобрать ответ ИИ»:
   - Читает таблицы из `02_Подготовка данных/05_Фрагменты шагов`.
   - Заполняет «Адаптированную структуру» (столбцы «Паттерн», «№шага/подшага», «Аномалия»).
   - Перемещает обработанные файлы в `02_Подготовка данных/Архив`.

5. Кнопка «Сделать промт следующего шага»:
   - Читает «Адаптированную структуру».
   - Формирует фрагмент схемы (CSV с заголовками) для следующего шага.
   - Генерирует промт шага из `prompt_step_template_v1.md`.
   - Создаёт папку `Шаг_NN`.

## Модули

### Mod_Prompt
- Генерация аналитических промтов 00_1 и 00_4 (`GenerateAnalyticalPrompts_Click`).
- Генерация промтов 00_2 и 00_3 с итерациями (`CommandGenerateTables`).
- Парсинг ответов ИИ с автоопределением формата (`CommandParseAiResponse`).
- Генерация промтов шагов без `Шаги.txt` (`MakeNextStepPrompt_Click`).
- Чтение и запись текстовых файлов в UTF-8.

### Mod_Folders
- Создание структуры папок.
- Инициализация параметров «Настройки» при первоначальном создании.

### Модуль настроек
- Чтение/запись параметров листа «Настройки».
- Три новых параметра: «Выжимки и анализ конфигурации», «Разбивка на паттерны», «Разбивка на шаги».

## Бэклог

| ID | Область | Статус |
|---|---|---|
| W7.1 | Mod_Patterns, Mod_Prompt, Mod_UI, Mod_Structure (VBA): GetPatternExceptions, GetKnownPatterns, GetPromptProfile, CommandGenerateTables = S0, FLAT_SHEET_NAME → Public, GenerateStepZeroPrompt → комментирование + MsgBox, CommandParseAiResponse — заглушка | Выполнено |
| W7.1a | Mod_Patterns (VBA): умные таблицы на листе «Паттерны» | Выполнено |
| W7.2 | Mod_Prompt, Mod_UI (VBA): BuildPrompt02Patterns, BuildPrompt03Profile, CommandParseAiResponse (реальный разбор) | Частично: генерация 00_2/00_3 с итерациями закрыта W-ANALYTIC-05 (CommandGenerateTables); BuildPrompt02Patterns/BuildPrompt03Profile и реальный разбор (CommandParseAiResponse) — не начато |
| W-FIX-PROMPT | prompts/prompt_fix_v1.0.md, Mod_Prompt (VBA): плейсхолдеры промта исправления + формирование промта по сводной таблице и листу отчёта по проверкам | Не начато (после W7.2; зависимость — сводная таблица + лист отчёта по проверкам) |
| W-DOCS-MAP | docs/Project_Map.md, docs/Spec_Optimization.md, docs/Journal.md, docs/Handoff.md, prompts/archive/README.md: привести документацию к состоянию dc288ac | Выполнено |

*Примечание: ветка Script не используется.*
Объём правки W7.2-GAPS-6: актуализированы строки Mod_Prompt / Mod_UI / Mod_Patterns, статус prompt_step_00_template.md, бэклог W7.1 и W7.1a → «Выполнено». Архитектура (состав модулей, поток данных) не менялась.
