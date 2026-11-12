# Карта проекта (Project Map)

Обновлено: 2026-11-11 (W7.2-GAPS-6: W7.1-A/B/C, W7.1a — функции чтения листа «Паттерны», PT003, GetPromptProfile, CommandGenerateTables = S0, умная таблица tblPatterns)

## VBA-модули (корень репозитория, формат .bas)

| Модуль | Назначение | Ключевые функции/процедуры |
|---|---|---|
| Mod_Util.bas | Журналирование, конфигурация, настройки, пути | WriteLog ("[i] <текст>"), WriteError ("ERROR [<код>]: <текст>"), WriteWarning ("[!] <текст>"), DebugLog, FormatMissingFile, FindConfigFile (.xml > .txt; только .pilotcfx — предупреждение), GetConfigValue, **GetProjectRoot** (читает ProjectRoot из «Настроек», пусто = ThisWorkbook.Path), **GetSettingLong(name, defaultValue)** (предупреждение U007 при пустом/некорректном значении), **GetSettingText(name, defaultValue)** (пусто → default + U007), **LegendRange** (именованный диапазон на лист «Настройки», B8), **MaxTypesPerTable** (поле B7, именованный диапазон, читается через GetSettingLong), **EnsureSettingsSheet** (6 полей B2:B7: ProjectRoot, CheckMode, AutoOpenResults, PromptGenVersion, PromptFixVersion, MaxTypesPerTable + 7-е поле B8 «Диапазон легенды» / LegendRange; Validation на B3/B4/B7; именованные диапазоны всех 6 полей; защита «pilot», разблокированы B2:B8; вызов Mod_Patterns.EnsurePatternsSheet с обёрткой On Error Resume Next + U008) |
| Mod_Folders.bas | Структура папок и проверка исходников | CreateProjectFolders (создаёт в т.ч. "01_Исходные данные\04_Дополнительные данные"), CheckSourceFolders (Схема .xlsx / Описание .docx / Конфигурация через FindConfigFile / сводка по доп. данным; возвращает True/False), GetStepFolder |
| Mod_Structure.bas | Парсинг схемы ЖЦ ОКС в адаптированную структуру | BuildFlatStructure (динамическая легенда: `GetSettingText("LegendRange", DEFAULT_LEGEND_RANGE)` вместо хардкода `legCells(1 To 6)`; при некорректном адресе — S011 + fallback B1:C3; восстановление prevScreenUpdating/prevDisplayAlerts/prevStatusBar, закрытие wbOut/wbSrc во всех ветках Fail/OpenErr/saveErr), **WriteFlatTable** (лист «Адаптированная_структура»; **FLAT_COLS = 7**, колонки A–G: «№шага/подшага», «Отображаемое имя», «Категория» (бывш. «Тип»), «Паттерн», «Входит в», «Состоит из», «Аномалия»; умная таблица `tblFlatStructure` A1:G<lastRow>; подсветка A:G; AutoFit A:G), WriteLog (лист «Аномалии схемы»; заголовки через Worksheet.Range("A1:H1")), SafeAutoFit, WriteTraceability (заголовок «Тип» → «Категория»), CheckNumberDuplicates, HighlightDuplicateRows (колонка «Аномалия» в диапазоне A:G), BackupFlatTableSheet, ArchiveOldSchemaArtifacts |
| Mod_Prompt.bas | Генерация промтов | **BuildPromptForStep** (общая цепочка сборки промта шага по prompt_step_template_v1.md), **ApplyCorrectionsBlock**, **FillAllPlaceholders**, **BuildStepTypeList** (распознаёт новые канонические заголовки «Категория» и «№шага/подшага» наравне со старыми «Тип»/«Шаг» — обратная совместимость; буквенные индексы колонок не используются), **GenerateStepZeroPrompt** (устаревшая ветка: генерирует только Промт №0 по prompt_step_00_template.md — заменяется в W7.1; подставляет {ЛИМИТ_ТИПОВ_В_ТАБЛИЦЕ} через GetSettingLong("MaxTypesPerTable", 50) — T3.1), **MakeNextStepPrompt**, MakeFixPrompt (не подключён — W-FIX-PROMPT), ReadStepList, ListFolderFiles, …DECOMPOSITION_RULES_V713. **GetPromptProfile** (W7.1-A: чистая функция без MsgBox, маршрутизация по lastStep: "" → ""; нет/00.i → «Папки»; 0/00/01.i/01.i.k → «Паттерны»; 01/NN≥2 → «Документы»; невалидное → ""; ответственность за "" — на вызывающем). **GenerateStepZeroPrompt — заглушка с MsgBox о замене (W7.1-C)**, шаблон перемещён в prompts/archive/. BuildPrompt02Patterns / BuildPrompt03Profile — W7.2 |
| Mod_Assembler.bas | Сборка итоговых настроек | сборка файлов настроек по результатам шагов |
| Mod_Docs.bas | Документация проекта | генерация отчётов/документов |
| Mod_Registry.bas | Реестры | Реестр_типов.txt, Реестр_покрытия.txt |
| Mod_TestGen.bas, Mod_TestGen_Assembler.bas | Генерация тестов | тестовые сценарии для собранных настроек |
| Mod_UI.bas | Кнопки/обработчики интерфейса | BuildFlatStructure_Click и др.; **RestoreUI создаёт 14 кнопок** на листе «Препроцессор»; кнопка 13 «Сформировать таблицы» → **CommandGenerateTables** (W7.1-B: реализует S0, вызывает Mod_Patterns.GetPatternExceptions / GetKnownPatterns и Mod_Prompt.GetPromptProfile; не заглушка), кнопка 14 «Разобрать ответ ИИ» → **CommandParseAiResponse** (заглушка W7.2; вызов только по кнопке 14, не из MakeNextStepPrompt; запись в лист «Паттерны» — W7.3-PATTERNS-UPDATE) |
| Mod_Patterns.bas | Лист «Паттерны» (PRE-04c): каркас для Промта №0.2 | **EnsurePatternsSheet** (идемпотентна: создание листа в конце книги, **без Cells.Clear** на существующем листе — PRE-04c-fix; секция A–D: заголовки в строке 2 (A2 «Паттерн», B2 «Типы», C2 «Количество», D2 «Итерация»), данные с 3-й; секция H–O: заголовки в строке 2, данные с 3-й, **8 предзаполненных правил** — Папки ЦИМ, Год (Маска/Год*), маска bim_*, точные системные типы document/ecmDocument/folder/project/projectfolder; Validation H/I/J/L/O на строках 3–1000; заголовки A1/H1 — названия секций (жирным, объединены A1:D1, H1:O1); защита паролем «pilot», разблокированы A2:D1000 и H3:O1000; MigrateLegacyLayout, ShiftLegacyYearRule, SyncPrefillCell для миграции старых каркасов); вызывается из Mod_Util.EnsureSettingsSheet; коды PT001/PT002/PT003. **GetPatternExceptions** (W7.1-A: секция H–O, фильтр O = «активно», PT003 при отсутствии листа/пустой секции) и **GetKnownPatterns** (W7.1-A: секция A–D, колонка A, заголовок «Известные паттерны:» внутри функции, "" при пустой секции, PT003 при отсутствии листа); лист не изменяется, защита не снимается. H–O предзаполнено 8 правилами, A–D пополняется по итерациям (пустая A–D — защитный сценарий, маркер не вводится). **W7.1a: PATTERNS_TABLE_NAME = "tblPatterns" (Public Const), EnsurePatternsListObject** — идемпотентное создание/расширение умной таблицы A2:O1000, стиль TableStyleMedium6, вызов из EnsurePatternsSheet до установки защиты |

## Папки рабочего проекта (создаются CreateProjectFolders вне репозитория)

```text
<корень .xlsm>/
├─ 01_Исходные данные/
│  ├─ 01_Схема/                 (*.xlsx — обязательно)
│  ├─ 02_Описание/              (*.docx — обязательно)
│  ├─ 03_Конфигурация/          (*.xml или *.txt; .pilotcfx не принимается)
│  └─ 04_Дополнительные данные/ (необязательно; список подставляется в промты шагов)
├─ 02_Подготовка схемы/         (результаты BuildFlatStructure: Адаптированная_структура_<дата>.xlsx;
│                              старые версии — в подпапке Архив/)
├─ 03_Результаты генерации/
│  ├─ 01_По шагам/              (папки Шаг_NN_...)
│  ├─ 03_Реестры/
│  └─ 04_Промты_готовые/
├─ 04_Промты/                   (Шаги.txt, готовые Промт_шаг_00_Декомпозиция.md, копия шаблонов Шаблоны/)
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

## Карта промтов (templates/prompts — папка prompts/ в репозитории)

Источник: отчёт QWEN (секция 2) с уточнениями W-DOCS-MAP.

| Шаблон | Назначение | Плейсхолдеры (полный список) | Кто заполняет (процедура VBA) | Где вызывается | Статус |
|---|---|---|---|---|---|
| `prompt_step_00_template.md` | Декомпозиция (старая v7.13) | `{SCHEMA_FILENAME}`, `{DESCRIPTION_FILENAME}`, `{CONFIG_FILENAME}`, `{EXTRA_DATA_LIST}`, `{Decomposition_Rules}` | `GenerateStepZeroPrompt` | `MakeNextStepPrompt` | устаревший; перемещён в `prompts/archive/` (W7.1-C) |
| `prompt_step_00_1_template.md` | Базовые типы из Конфигурации | `{CONFIG_FILENAME}` (вход), `{{СПИСОК_БАЗОВЫХ_ТИПОВ}}` (выход) | отсутствует | не вызывается | не подключён (ручной) |
| `prompt_step_00_2_template.md` | Паттерны и типы-справочники из Схемы | `{{ИСКЛЮЧЕНИЯ_ПАТТЕРНОВ}}`, `{{СПИСОК_БАЗОВЫХ_ТИПОВ}}`, `{{СПИСОК_ИЗВЕСТНЫХ_ПАТТЕРНОВ}}`, `{{ТАБЛИЦА_ИЗВЕСТНЫХ_GUID}}`, `{{ТИПЫ_СПРАВОЧНИКИ}}` | отсутствует | не вызывается | не подключён (W7.2 — BuildPrompt02Patterns) |
| `prompt_step_00_3_template.md` | Список шагов (3 профиля: Папки / Паттерны / Документы) | `{{СПИСОК_БАЗОВЫХ_ТИПОВ}}`, `{{ТАБЛИЦА_ИЗВЕСТНЫХ_GUID}}` | отсутствует | не вызывается | не подключён (W7.2 — BuildPrompt03Profile) |
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

## Бэклог

| ID | Область | Статус |
|---|---|---|
| W7.1 | Mod_Patterns, Mod_Prompt, Mod_UI, Mod_Structure (VBA): GetPatternExceptions, GetKnownPatterns, GetPromptProfile, CommandGenerateTables = S0, FLAT_SHEET_NAME → Public, GenerateStepZeroPrompt → комментирование + MsgBox, CommandParseAiResponse — заглушка | Выполнено |
| W7.1a | Mod_Patterns (VBA): умные таблицы на листе «Паттерны» | Выполнено |
| W7.2 | Mod_Prompt, Mod_UI (VBA): BuildPrompt02Patterns, BuildPrompt03Profile, CommandParseAiResponse (реальный разбор) | Не начато (после W7.1) |
| W-FIX-PROMPT | prompts/prompt_fix_v1.0.md, Mod_Prompt (VBA): плейсхолдеры промта исправления + формирование промта по сводной таблице и листу отчёта по проверкам | Не начато (после W7.2; зависимость — сводная таблица + лист отчёта по проверкам) |
| W-DOCS-MAP | docs/Project_Map.md, docs/Spec_Optimization.md, docs/Journal.md, docs/Handoff.md, prompts/archive/README.md: привести документацию к состоянию dc288ac | Выполнено |

*Примечание: ветка Script не используется.*
Объём правки W7.2-GAPS-6: актуализированы строки Mod_Prompt / Mod_UI / Mod_Patterns, статус prompt_step_00_template.md, бэклог W7.1 и W7.1a → «Выполнено». Архитектура (состав модулей, поток данных) не менялась.
