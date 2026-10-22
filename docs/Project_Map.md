# Карта проекта (Project Map)

Обновлено: 2026-10-22 (W-CORRECTION-PACK-1).

## VBA-модули (корень репозитория, формат .bas)

| Модуль | Назначение | Ключевые функции/процедуры |
|---|---|---|
| Mod_Util.bas | Журналирование, конфигурация, пути | WriteLog ("[i] <текст>"), WriteError ("ERROR [<код>]: <текст>"), WriteWarning ("[!] <текст>"), FormatMissingFile, FindConfigFile (.xml > .txt; только .pilotcfx — предупреждение), GetConfigValue, GetProjectRoot |
| Mod_Folders.bas | Структура папок и проверка исходников | CreateProjectFolders (создаёт в т.ч. "01_Исходные данные\04_Дополнительные данные"), CheckSourceFolders (Схема .xlsx / Описание .docx / Конфигурация через FindConfigFile / сводка по доп. данным; возвращает True/False), GetStepFolder |
| Mod_Structure.bas | Парсинг схемы ЖЦ ОКС в плоскую таблицу | BuildFlatStructure, WriteFlatTable (лист "Плоская_таблица"), WriteLog (лист "Аномалии схемы"), WriteTraceability, CheckNumberDuplicates (дубли номеров/наименований — только предупреждения), BackupFlatTableSheet (бэкапы "Плоская_таблица_bak_<YYYYMMDD_HHMMSS>", ротация 5) |
| Mod_Prompt.bas | Генерация промтов | GenerateStepZeroPrompt (Промт №0; блокировка F005 при неготовых исходниках), MakeNextStepPrompt (промт шага + блок "ДОПОЛНИТЕЛЬНЫЕ ДАННЫЕ ПРОЕКТА"; блокировка F005), MakeFixPrompt, ReadStepList (канонический разделитель "\|" из Шаги.txt), BuildExtraDataBlock, ListFolderFiles, DECOMPOSITION_RULES_V713 (алгоритм «анализ → прогноз +50% → шаги → формат NN\|Имя_папки») |
| Mod_Assembler.bas | Сборка итоговых настроек | сборка файлов настроек по результатам шагов |
| Mod_Docs.bas | Документация проекта | генерация отчётов/документов |
| Mod_Registry.bas | Реестры | Реестр_типов.txt, Реестр_покрытия.txt |
| Mod_TestGen.bas, Mod_TestGen_Assembler.bas | Генерация тестов | тестовые сценарии для собранных настроек |
| Mod_UI.bas | Кнопки/обработчики интерфейса | BuildFlatStructure_Click и др. |

## Папки рабочего проекта (создаются CreateProjectFolders вне репозитория)

```text
<корень .xlsm>/
├─ 01_Исходные данные/
│  ├─ 01_Схема/                 (*.xlsx — обязательно)
│  ├─ 02_Описание/              (*.docx — обязательно)
│  ├─ 03_Конфигурация/          (*.xml или *.txt; .pilotcfx не принимается)
│  └─ 04_Дополнительные данные/ (необязательно; список подставляется в промты шагов)
├─ 02_Подготовка схемы/         (результаты BuildFlatStructure: Плоская_таблица_<дата>.xlsx)
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

## templates/prompts (папка prompts/ в репозитории)

| Файл | Назначение |
|---|---|
| prompt_step_00_template.md | Шаблон Промта №0 (декомпозиция): алгоритм декомпозиции, проверка файлов (.xml/.txt; .pilotcfx не принимается ИИ), вывод `NN|Имя_папки` |
| prompt_step_template_v1.md | Шаблон промта следующего шага ({NN}, {STEP_NAME}, {REGISTRY_TYPES}, {REGISTRY_COVERAGE}, {SCHEMA}) |
| prompt_fix_v1.0.md | Шаблон промта исправления (MakeFixPrompt) |
| prompt_generator_v7.13.md | Общая версия генератора промтов |
| prompt_settings_template.md | Шаблон настроек |

Внешняя папка "Шаблоны" (копия шаблонов у Пользователя) в репозиторий НЕ входит
и обновляется Пользователем вручную после принятия PR.

## Поток данных

1. Проверка/создание папок: Mod_Folders.CreateProjectFolders, CheckSourceFolders.
2. Подготовка схемы: Mod_Structure.BuildFlatStructure (дубли → "Аномалии схемы", бэкап листа перед перезаписью).
3. Реестры: Mod_Registry.
4. Промты: Mod_Prompt (№0 → Шаги.txt → промты шагов → фиксы); перед генерацией — CheckSourceFolders (F005).
5. Сборка: Mod_Assembler; тесты: Mod_TestGen*.
