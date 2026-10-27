# Образцы «Дополнительных параметров» для типов атрибутов Pilot-BIM

Справочный файл (НЕ инвариант). Источник — справка Pilot-ICE/Pilot-BIM
(«Атрибуты») + уточнения пользователя.

## 0. Общее правило корневого тега (критично)
- Enum и Object (тип «Строка», Kind="Enum"/"Object") → корневой <Attr …>,
  БЕЗ обёртки <Configuration>.
- ElementBook (тип «Справочник элементов») → вложен в <Configuration>;
  корневой тег — <Configuration>.
- Остальные (Организационная единица, многострочная Строка, Дата и время,
  Десятичный, Состояние, Булевый) → корневой <Configuration …>.
- Наследуемый — текст «Наследуется от: <тип>.<атрибут>», не XML.
- Нумератор, Целое, Вещественное — колонка 12 ПУСТО.

Устаревшие (НЕ использовать): Attr Kind="OrgUnit", Attr Kind="Type".
Если формат не найден — пометить «формат не определён», не выдумывать.

## 1. Enum (Строка, Kind="Enum")
<Attr Kind="Enum" IsEditable="True" DefaultItemIndex="1">
<Item>Значение 1</Item>
<Item>Значение 2</Item>
</Attr>
DefaultItemIndex — 1-based (1 = первый). «Defaultindex» не существует.

## 2. Object (Строка, Kind="Object")
Корневой тег — <Attr …>, без <Configuration>.
<Attr Kind="Object" Source="[GUID]" IsEditable="True" StringFormat="{name}" EditorType="Dialog" ElementsTypes="…" CreateLink="True"/>

2.1. ElementsTypes — КОНТЕКСТНЫЙ ТИП ЭКЗЕМПЛЯРА (контекст + Схема +
Конфигурация). НЕ корень справочника (reference_* не ставить), НЕ
фиксированный список. Маппинг «справочник → типы» не требуется.
Примеры:
- Договор → ElementsTypes="dogovor"
- Приказ внутри контрагента → ElementsTypes="prikaz"
- Контрагент → ElementsTypes="contractor" / "individual_entrepreneur" /
  "physical_person" / перечисление

2.2. Source — GUID контейнера:
1) тип сам — справочник из {{ТАБЛИЦА_ИЗВЕСТНЫХ_GUID}} → его GUID;
2) тип дочерний к справочнику R из таблицы → GUID ближайшего предка R;
3) иначе → "00000001-0001-0001-0001-000000000001" (корень БД); валидно,
   без пометок.
Конфигурация GUID не содержит.

2.3. EditorType: ComboBox — без связей; Dialog — со связями,
множественный выбор. CreateLink="True" — только при Dialog.
AllowMultiSelect, AutoFill — по необходимости.

Пример (контрагент):
<Attr Kind="Object" Source="bca77be0-726d-4ae6-9699-ff465db5ba01" IsEditable="True" StringFormat="{name}" EditorType="Dialog" ElementsTypes="contractor,individual_entrepreneur,physical_person" CreateLink="True"/>

Пример (договор вне справочников из таблицы):
<Attr Kind="Object" Source="00000001-0001-0001-0001-000000000001" IsEditable="True" StringFormat="{name}" EditorType="Dialog" ElementsTypes="dogovor" CreateLink="True"/>

## 3. Строка (многострочная)
<Configuration AllowLineBreaks="True"/>

## 4. Организационная единица
<Configuration AllowMultiSelect="False" AllowDepartments="True" AllowPersons="True" AllowGroups="False"/>
Параметры: AutoFill, AutoFillValueSource, AutoFillSourceTypes,
AutoFillSourceAttributeName, AutoFillValueConverter, DialogCaption,
DialogHint, InputHint, ItemLinesCount, ShowInHistory.
Устаревшее: Attr Kind="OrgUnit" — НЕ использовать.

## 5. Дата и время
<Configuration Format=" { 0 :dd.MM.yyyy HH:mm}" AutoFill="OnOpen"/>
Параметры: UseExtendedDateTimeEditor, AllowPastDates, ShowInHistory,
Format, DefaultValue, DefaultValueFormat, DefaultDaysCount, AutoFill,
CalendarSourceAttrName.

## 6. Десятичный (денежная единица)
<Configuration Format=" {0:C} " Culture="ru-RU" />
Доллары: Culture="en-US".
Произвольный формат: <Configuration Format=" {0:$#, ##0.00; ($#, ##0.00); }" />

## 7. Состояние
<Configuration ShowInListView="true" EnablePopupCommands="false" />
StateMachineId НЕ указывать (даже если был в исходных примерах).

## 8. Справочник элементов (ElementBook)
Корневой тег — <Configuration>; ElementBook вложен.
<Configuration>
<ElementBook Source="[GUID]" StringFormat="{name}" AllowMultiSelect="false" CreateLink="true" AutoFill="true"/>
</Configuration>
Параметры: Source ("00000001-0001-0001-0001-000000000001" = папка Начало),
StringFormat, AllowMultiSelect, ElementsTypes, ShowInListView,
EnablePopupCommands, CreateLink, AutoFill.
Правила Source — как в разделе 2.2 (таблица GUID / корень БД).
AttachmentToTaskAutoFill — НЕ использовать.

## 9. Булевый
<Configuration DefaultValue="True"/>

## 10. Наследуемый
Поля «Дополнительные параметры» НЕТ. В колонке 12 — текст:
Наследуется от: <тип>.<атрибут>
Например: Наследуется от: project.code

## 11. Нумератор
Поля «Дополнительные параметры» НЕТ. Настройка в «Описание»:
BX={Counter:d5}-{CurrentDate:yyyy}. Колонка 12 — ПУСТО.

## 12. Целое число, Вещественное число
Колонка 12 — ПУСТО.

## 13. Устаревшие форматы (НЕ использовать)
- <Attr Kind="OrgUnit" … /> (Строка с оргструктурой).
- <Attr Kind="Type" … />.
- Параметр «Defaultindex» (только DefaultItemIndex).
- StateMachineId для «Состояние».