# Образцы «Дополнительных параметров» (колонка 12 листа «Атрибуты»)

Назначение: справочник валидных XML-образцов для колонки 12 «Дополнительные
параметры» листа «Атрибуты» (Промт шага, блоки 9–10). Файл передаётся вместе
с промтом шага (пункт 3.5 шаблона `prompt_step_template_v1.md`).

Общие правила:
- Колонка 12 — валидный XML-фрагмент (кроме «Наследуемого» — там текст).
- Корневой тег зависит от ТИПА АТРИБУТА, а не от модуля.
- Для Enum, Object и ElementBook корневой тег — всегда `Configuration`;
  `Attr` и `ElementBook` — вложенные элементы, корневыми быть НЕ могут.
- GUID — только из {{ТАБЛИЦА_ИЗВЕСТНЫХ_GUID}} либо корень БД
  `00000001-0001-0001-0001-000000000001` (fallback). Конфигурация GUID
  НЕ содержит; выдумывать GUID запрещено.
- Обратные кавычки и маски (`Mask`) запрещены.
- Если формата для типа нет в этом файле — пометь «формат не определён»,
  не выдумывай.

## 1. Организационная единица

Корневой тег — `Configuration`.

```xml
<Configuration AllowMultiSelect="False" AllowDepartments="True" AllowPersons="True" AllowGroups="False"/>
```

Необязательные параметры: AutoFill, AutoFillValueSource, AutoFillSourceTypes,
AutoFillSourceAttributeName, AutoFillValueConverter, DialogCaption, DialogHint,
InputHint, ItemLinesCount, ShowInHistory.

Устаревший формат (НЕ использовать): `<Attr Kind="OrgUnit">…</Attr>` со
строковым содержимым.

## 2. Строка: справочник перечислений и объектный справочник

### 2.1. Enum (Kind="Enum")

Корневой тег — `Configuration`, внутри — `Attr Kind="Enum"`, внутри Attr —
теги `Item` (по одному значению на тег).

```xml
<Configuration>
<Attr Kind="Enum" IsEditable="True" DefaultItemIndex="1">
<Item>Значение 1</Item>
<Item>Значение 2</Item>
</Attr>
</Configuration>
```

Параметры Attr: `IsEditable` (True/False), `DefaultItemIndex` — индекс
значения по умолчанию, **1-based** (1 = первый элемент списка). Параметра
`Defaultindex` НЕ существует — использовать только `DefaultItemIndex`.
Без `DefaultItemIndex` значение по умолчанию не задано.

Значения `Item` — дословно, на кириллице; закрывающие теги `</Attr>` и
`</Configuration>` обязательны. Голых `<Item>` вне `Attr` быть не должно.

Однострочный вариант (если элементов мало):

```xml
<Configuration><Attr Kind="Enum" IsEditable="True" DefaultItemIndex="1"><Item>Вариант А</Item><Item>Вариант Б</Item></Attr></Configuration>
```

### 2.2. Object (Kind="Object")

Корневой тег — `Configuration`, внутри — `Attr Kind="Object"`.

```xml
<Configuration>
<Attr Kind="Object" Source="[GUID]" IsEditable="True" StringFormat="{name}" EditorType="Dialog" ElementsTypes="…" CreateLink="True"/>
</Configuration>
```

Параметры:
- `Source` — GUID КОНТЕЙНЕРА (справочника), в котором лежит тип из
  ElementsTypes. Единственный источник — {{ТАБЛИЦА_ИЗВЕСТНЫХ_GUID}};
  fallback — корень БД `00000001-0001-0001-0001-000000000001` (валидное
  значение, пометок не требует). Правило выбора: собственный GUID, если тип
  сам является справочником из таблицы; иначе — GUID ближайшего предка-
  справочника по дереву «Входит в» (наименьшая глубина); иначе — корень БД.
- `ElementsTypes` — КОНТЕКСТНЫЙ ТИП ЭКЗЕМПЛЯРА, на который ссылается
  атрибут (из контекста, Схемы и Конфигурации). НЕ фиксированный список,
  НЕ корень справочника (`reference_*` сюда НЕ ставится). Примеры:
  `ElementsTypes="dogovor"`, `ElementsTypes="prikaz"`,
  `ElementsTypes="contractor,individual_entrepreneur,physical_person"`.
- `StringFormat` — формат отображения, `{attr}`; можно комбинировать:
  `{inn}-{name}`.
- `EditorType` — `ComboBox` (раскрывающийся список, элементы без связей)
  или `Dialog` (диалог, связи, множественный выбор).
- `CreateLink="True"` — только при `EditorType="Dialog"`.
- `AllowMultiSelect`, `AutoFill` — по необходимости.

Маппинг «справочник → типы» НЕ требуется и НЕ используется: ElementsTypes
определяется контекстом атрибута, Схемой и Конфигурацией.

Контрагенты (известный образец):

```xml
<Configuration>
<Attr Kind="Object" Source="bca77be0-726d-4ae6-9699-ff465db5ba01" IsEditable="True" StringFormat="{name}" EditorType="Dialog" ElementsTypes="contractor" CreateLink="True"/>
</Configuration>
```

Если тип экземпляра не найден в Конфигурации — пометь «ElementsTypes не
определён: нет в Конфигурации», не выдумывай.

Устаревший формат (НЕ использовать): `<Attr Kind="Type">…</Attr>`.

## 3. Строка (многострочная)

```xml
<Configuration AllowLineBreaks="True"/>
```

## 4. Дата и время

```xml
<Configuration Format=" { 0 :dd.MM.yyyy HH:mm}" AutoFill="OnOpen"/>
```

Параметры: UseExtendedDateTimeEditor, AllowPastDates, ShowInHistory, Format,
DefaultValue, DefaultValueFormat, DefaultDaysCount, AutoFill
(OnOpen / OnFocus / Disabled), CalendarSourceAttrName.

Только дата:

```xml
<Configuration Format=" { 0 :dd.MM.yyyy }"/>
```

## 5. Целое число

Колонка 12 — ПУСТО.

## 6. Вещественное число

Колонка 12 — ПУСТО.

## 7. Десятичный (денежная единица)

```xml
<Configuration Format=" {0:C} " Culture="ru-RU" />
```

Доллары: `Culture="en-US"`. Произвольный формат:

```xml
<Configuration Format=" {0:$#, ##0.00; ($#, ##0.00); }" />
```

## 8. Нумератор

Поля «Дополнительные параметры» НЕТ — настройка живёт в поле «Описание»
(например `BX={Counter:d5}-{CurrentDate:yyyy}`). Колонка 12 — ПУСТО.

## 9. Состояние

```xml
<Configuration ShowInListView="true" EnablePopupCommands="false" />
```

`StateMachineId` НЕ указывать (даже если он встречался в исходных примерах)
— параметр необязательный и в генерации не используется.

## 10. Организационная единица (расширенный образец)

С автозаполнением и настройкой диалога:

```xml
<Configuration AllowMultiSelect="True" AllowDepartments="True" AllowPersons="True" AllowGroups="True" AutoFill="True" DialogCaption="Выбор исполнителя" InputHint="Укажите подразделение" ItemLinesCount="10" ShowInHistory="True"/>
```

## 11. Справочник элементов

Корневой тег — `Configuration`, внутри — `ElementBook` (сам `ElementBook`
корневым быть НЕ может).

```xml
<Configuration>
<ElementBook Source="[GUID]" StringFormat="{name}" AllowMultiSelect="false" ElementsTypes="…" CreateLink="true" AutoFill="true"/>
</Configuration>
```

Параметры `ElementBook`: Source, StringFormat, AllowMultiSelect,
ElementsTypes, ShowInListView, EnablePopupCommands, CreateLink, AutoFill.

`Source` — по тем же правилам, что и для `Attr Kind="Object"` (раздел 2.2):
только из {{ТАБЛИЦА_ИЗВЕСТНЫХ_GUID}} (ближайший предок-справочник по дереву
«Входит в»), иначе корень БД `00000001-0001-0001-0001-000000000001`
(папка «Начало», все элементы БД).

`AttachmentToTaskAutoFill` — НЕ использовать (задания не создаются).

## 12. Наследуемый

Поля «Дополнительные параметры» НЕТ (настраивается в UI: тип родительского
элемента + наследуемый атрибут). В колонке 12 — ТЕКСТ, не XML:

```text
Наследуется от: <тип>.<атрибут>
```

Например: `Наследуется от: project.code`

## 13. Булевый

```xml
<Configuration DefaultValue="True"/>
```

`DefaultValue="False"` — если по умолчанию «нет».

---

## Приложение. Известные GUID (образец)

| Справочник | GUID | Тип справочника (системное имя) |
|---|---|---|
| Корень БД | `00000001-0001-0001-0001-000000000001` | Root_object_type |
| Контрагенты | `bca77be0-726d-4ae6-9699-ff465db5ba01` | reference_contractors |
| Разделы | `8a523294-f2be-4fc6-8a70-52602a6ca234` | reference_sections |
| Подразделы | `519cae9d-10dc-4843-b151-9145f78bfbf7` | reference_subsection |
| Основные комплекты | `989acd39-e3c7-4a82-9b70-55f5504ad88c` | reference_sets |

Рабочий источник GUID — {{ТАБЛИЦА_ИЗВЕСТНЫХ_GUID}} из `project_settings.md`;
этот список приведён только как образец формата.
