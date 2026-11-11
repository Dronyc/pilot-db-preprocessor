Attribute VB_Name = "Mod_Patterns"
Option Explicit

' ============================================================================
' Модуль: Mod_Patterns
' Назначение: лист «Паттерны» (PRE-04c) — каркас листа для Промта №0.2
'             («Паттерны и справочники»):
'   Строка 1 — названия секций: A1 «СПИСОК ПАТТЕРНОВ» (объединено A1:D1),
'               H1 «ИСКЛЮЧЕНИЯ ПАТТЕРНОВ» (объединено H1:O1); жирным.
'   Секция A-D  («Список паттернов», append-only): заголовки в строке 2
'               «Паттерн», «Типы», «Количество», «Итерация»; строки 3+
'               пустые — заполняются кодом позже (разбор ответа ИИ, W7).
'   Секция H-O  («ИСКЛЮЧЕНИЯ ПАТТЕРНОВ»): заголовки в строке 2 «Область»,
'               «Поле», «Тип матчинга», «Значение», «Действие», «Примеры»,
'               «Причина», «Статус»; предзаполнены стартовые правила
'               (папки ЦИМ/Год — ВКЛЮЧАТЬ; базовые системные типы —
'               ИСКЛЮЧАТЬ) — только при создании нового листа, со строки 3.
' Выпадающие списки (Validation) навешаны на колонки H, I, J, L, O
' (строки 3..PATTERNS_LAST_INPUT_ROW — PRE-04c-fix2: заголовки строки 2
' списками не перекрываются); защита листа паролем "pilot",
' редактируемыми остаются только диапазоны ввода данных A2:D1000 и H3:O1000.
' PRE-04c-fix: защита от потери данных — Cells.Clear убран; если лист
' «Паттерны» уже существует, данные в A2:D1000 и H2:O1000 НЕ ТРОГАЮТСЯ —
' процедура лишь дописывает недостающие заголовки, пересоздаёт Validation
' и переустанавливает защиту. Полная очистка не выполняется ни в одном
' из сценариев (новый лист создаётся пустым по определению).
' PRE-04c-fix2: удалена дублирующая старая строка правила «Год»
' (Категория/Отображаемое/Точное/Год) из предзаполнения — осталось одно
' правило «Год» (Тип/Отображаемое/Маска/Год*); Validation применён со
' строки 3; область ввода секции H-O — H3:O1000.
' Зависимости: Mod_Util (WriteLog / WriteError / WriteWarning).
' Реестр кодов ошибок/предупреждений модуля (единый реестр проекта):
'   PT001 - ошибка подготовки листа «Паттерны» (EnsurePatternsSheet);
'   PT002 - предупреждение: не удалось применить Validation к ячейкам
'          секции H-O (лист подготовлен, списки нужно настроить вручную);
'   PT003 - предупреждение: лист «Паттерны» отсутствует или секция пуста
'          (чтение W7.1-A: GetPatternExceptions / GetKnownPatterns —
'          функция возвращает "").
' Примечание: процедура EnsurePatternsSheet вызывается из
' Mod_Util.EnsureSettingsSheet (главная точка подготовки служебных листов
' книги), поэтому вызов защищён On Error Resume Next: сбой подготовки
' листа «Паттерны» не должен блокировать работу с листом «Настройки».
' ============================================================================

' ---------------------------------------------------------------------------
' КОНСТАНТЫ
' ---------------------------------------------------------------------------

Private Const SHEET_NAME As String = "Паттерны"
Private Const PATTERNS_PASSWORD As String = "pilot" ' пароль единый для книги
Private Const PATTERNS_LAST_INPUT_ROW As Long = 1000 ' нижняя граница диапазонов ввода
Public Const PATTERNS_TABLE_NAME As String = "tblPatterns"

' Названия секций (строка 1)
Private Const TITLE_LIST As String = "СПИСОК ПАТТЕРНОВ"
Private Const TITLE_EXC As String = "ИСКЛЮЧЕНИЯ ПАТТЕРНОВ"

' Заголовки секции A-D (список паттернов)
Private Const HDR_A As String = "Паттерн"
Private Const HDR_B As String = "Типы"
Private Const HDR_C As String = "Количество"
Private Const HDR_D As String = "Итерация"

' Заголовки секции H-O (ИСКЛЮЧЕНИЯ паттернов)
Private Const HDR_H As String = "Область"
Private Const HDR_I As String = "Поле"
Private Const HDR_J As String = "Тип матчинга"
Private Const HDR_K As String = "Значение"
Private Const HDR_L As String = "Действие"
Private Const HDR_M As String = "Примеры"
Private Const HDR_N As String = "Причина"
Private Const HDR_O As String = "Статус"

' ---------------------------------------------------------------------------
' Подготовка листа «Паттерны»: создание каркаса, предзаполнение исключений,
' выпадающие списки, форматирование и защита.
' Идемпотентна и безопасна для данных (PRE-04c-fix): безусловный Cells.Clear
' убран.
'   * Лист не существует -> создаётся, рисуются названия секций (строка 1) и
'     заголовки (строка 2), навешивается Validation, предзаполняются
'     исключения (со строки 3), лист защищается.
'   * Лист уже существует -> данные НЕ ОЧИЩАЮТСЯ: проверяются заголовки
'     (если отсутствуют или смещены в строку 1 старого каркаса PRE-04c —
'     дописываются/переносятся в строку 2), пересоздаётся Validation,
'     переустанавливается защита. Значения в диапазонах A2:D1000 и H2:O1000
'     (кроме ячеек заголовков строки 2 и названий секций строки 1) не
'     затрагиваются.
' ---------------------------------------------------------------------------
Public Sub EnsurePatternsSheet()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsAny As Worksheet
    Dim created As Boolean
    Dim excData As Variant
    Dim r As Long
    Dim c As Long

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook

    ' 1. Поиск листа «Паттерны»; если нет - создать в конце книги
    For Each wsAny In wb.Worksheets
        If StrComp(wsAny.Name, SHEET_NAME, vbTextCompare) = 0 Then
            Set ws = wsAny
            Exit For
        End If
    Next wsAny

    created = (ws Is Nothing)

    If created Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_NAME
    Else
        ws.Unprotect Password:=PATTERNS_PASSWORD   ' снять защиту перед правкой
        ' КРИТИЧНО (PRE-04c-fix): полная очистка листа (Cells.Clear) НЕ
        ' выполняется — пользовательские данные A2:D1000 и H2:O1000
        ' сохраняются.
    End If

    ' 2. Лист со старым каркасом PRE-04c (заголовки в строке 1, данные с
    '    строки 2): перенос раскладки - строки данных 2..lastRow смещаются
    '    вниз на одну строку (значения сохраняются), заголовки переносятся
    '    в строку 2, объединения строки 1 снимаются. Операция выполняется
    '    один раз: после переноса HasHeaderAt(ws,2) = True.
    If Not created Then
        If HasHeaderAt(ws, 1) And Not HasHeaderAt(ws, 2) Then
            On Error Resume Next
            ws.Range("A1:D1").MergeCells = False
            ws.Range("H1:O1").MergeCells = False
            On Error GoTo ErrHandler
            MigrateLegacyLayout ws
        End If
    End If

    ' 3. Строка 1: названия секций (жирным, объединённые ячейки A1:D1 и
    '    H1:O1). Существующие чужие значения не перезаписываются.
    EnsureSectionTitle ws, "A1", TITLE_LIST, "A1:D1"
    EnsureSectionTitle ws, "H1", TITLE_EXC, "H1:O1"

    ' 4. Секция A-D: список паттернов (заголовки «Паттерн», «Типы»,
    '    «Количество», «Итерация» в строке 2; ниже — пустые строки ввода,
    '    заполняются кодом позже, напр. разбором ответа ИИ - W7).
    '    Если заголовков нет - они дописываются в строку 2; существующие
    '    значения ячеек не трогаются.
    EnsureHeaderCell ws, 2, 1, HDR_A
    EnsureHeaderCell ws, 2, 2, HDR_B
    EnsureHeaderCell ws, 2, 3, HDR_C
    EnsureHeaderCell ws, 2, 4, HDR_D

    ' 5. Секция H-O: ИСКЛЮЧЕНИЯ паттернов (заголовки «Область», «Поле»,
    '    «Тип матчинга», «Значение», «Действие», «Примеры», «Причина»,
    '    «Статус» в строке 2; дописываются, если отсутствуют)
    EnsureHeaderCell ws, 2, 8, HDR_H
    EnsureHeaderCell ws, 2, 9, HDR_I
    EnsureHeaderCell ws, 2, 10, HDR_J
    EnsureHeaderCell ws, 2, 11, HDR_K
    EnsureHeaderCell ws, 2, 12, HDR_L
    EnsureHeaderCell ws, 2, 13, HDR_M
    EnsureHeaderCell ws, 2, 14, HDR_N
    EnsureHeaderCell ws, 2, 15, HDR_O

    ' 6. Предзаполнение стартовых правил исключений (матрица 8 строк x
    '    8 колонок H..O; PRE-04c-fix2: старая строка правила «Год»
    '    удалена) - ТОЛЬКО при создании нового листа, со строки 3
    '    (строка 2 — заголовки). Для существующего листа правила НЕ
    '    перезаписываются: пользовательские данные сохраняются; при этом
    '    выполняется точечная сверка значений предзаполнения с эталоном
    '    фикса (см. шаг 6.1).
    If created Then
        ' PRE-04c-fix2: старая строка правила «Год» (Категория, Отображаемое,
        ' Точное, Год) удалена — осталась только одна строка для правила
        ' «Год»: Тип / Отображаемое / Маска / Год* / ВКЛЮЧАТЬ.
        excData = Array( _
            Array("Категория", "Отображаемое", "Точное", "Папки ЦИМ", "ВКЛЮЧАТЬ", _
                  "КЦИМ; ППЦИМ", "позиционный паттерн", "активно"), _
            Array("Тип", "Отображаемое", "Маска", "Год*", "ВКЛЮЧАТЬ", _
                  "Год 2024; Год 2025", "позиционный паттерн", "активно"), _
            Array("Тип", "Системное", "Маска", "bim_*", "ИСКЛЮЧАТЬ", _
                  "bim_coordinationModel; bim_modelPartFolder", "базовые типы", "активно"), _
            Array("Тип", "Системное", "Точное", "document", "ИСКЛЮЧАТЬ", _
                  ChrW(9472), "базовый тип", "активно"), _
            Array("Тип", "Системное", "Точное", "ecmDocument", "ИСКЛЮЧАТЬ", _
                  ChrW(9472), "базовый тип", "активно"), _
            Array("Тип", "Системное", "Точное", "folder", "ИСКЛЮЧАТЬ", _
                  ChrW(9472), "базовый тип", "активно"), _
            Array("Тип", "Системное", "Точное", "project", "ИСКЛЮЧАТЬ", _
                  ChrW(9472), "базовый тип", "активно"), _
            Array("Тип", "Системное", "Точное", "projectfolder", "ИСКЛЮЧАТЬ", _
                  ChrW(9472), "базовый тип", "активно"))

        For r = LBound(excData) To UBound(excData)
            For c = LBound(excData(r)) To UBound(excData(r))
                ws.Cells(r + 3, c + 8).Value = CStr(excData(r)(c))
            Next c
        Next r
    Else
        ' 6.1. Точечная сверка предзаполнения на существующем листе
        '      (PRE-04c-fix2): старая строка правила «Год» (Категория,
        '      Отображаемое, Точное, Год) удаляется — все правила, лежащие
        '      ниже неё, поднимаются на одну строку вверх; если нижняя
        '      строка содержит копию старого эталона предзаполнения
        '      (не редактировалась пользователем), она очищается. После
        '      сдвига значения приводятся к эталону только в тех ячейках
        '      стартовых правил (строки 3..10 секции H-O после переноса
        '      раскладки), которые ещё содержат старые эталонные значения
        '      (до фикса); ячейки, изменённые пользователем или добавленные
        '      позже, не затрагиваются. Колонки H..O = 8..15.
        ShiftLegacyYearRule ws
        SyncPrefillCell ws, 3, 8, "Категория", "Папки ЦИМ"     ' правило «Папки ЦИМ»
        SyncPrefillCell ws, 4, 10, "маска", "Маска"            ' правило bim_*
        ' Правило «Год» (вторая строка предзаполнения, PRE-04c-fix2):
        ' Область Тип, Поле Отображаемое, Тип матчинга Маска, Значение Год*,
        ' Действие ВКЛЮЧАТЬ, Примеры "Год 2024; Год 2025", Причина
        ' "позиционный паттерн", Статус активно.
        SyncPrefillCell ws, 4, 8, "Тип", "Тип"
        SyncPrefillCell ws, 4, 9, "Отображаемое", "Отображаемое"
        SyncPrefillCell ws, 4, 10, "Маска", "Маска"
        SyncPrefillCell ws, 4, 11, "Год*", "Год*"
        SyncPrefillCell ws, 4, 12, "ВКЛЮЧАТЬ", "ВКЛЮЧАТЬ"
        SyncPrefillCell ws, 4, 13, "Год 2024; Год 2025", "Год 2024; Год 2025"
        SyncPrefillCell ws, 4, 14, "позиционный паттерн", "позиционный паттерн"
        SyncPrefillCell ws, 4, 15, "активно", "активно"
    End If

    ' 7. Выпадающие списки (Validation) для колонок H, I, J, L, O
    '    (навешиваются всегда: если списки сбились - восстанавливаются;
    '     сами значения ячеек не изменяются).
    '    PRE-04c-fix2: списки применяются начиная со строки 3 (строка 2 —
    '    заголовки секции H-O, выпадающие списки на них не навешиваются).
    ApplyListValidation ws.Range("H3:H" & CStr(PATTERNS_LAST_INPUT_ROW)), "Тип,Категория"
    ApplyListValidation ws.Range("I3:I" & CStr(PATTERNS_LAST_INPUT_ROW)), "Отображаемое,Системное"
    ApplyListValidation ws.Range("J3:J" & CStr(PATTERNS_LAST_INPUT_ROW)), "Точное,Маска"
    ApplyListValidation ws.Range("L3:L" & CStr(PATTERNS_LAST_INPUT_ROW)), "ВКЛЮЧАТЬ,ИСКЛЮЧАТЬ"
    ApplyListValidation ws.Range("O3:O" & CStr(PATTERNS_LAST_INPUT_ROW)), "активно,отменено"

    ' 8. Форматирование: названия секций (строка 1) и заголовки (строка 2)
    '    жирным, центрирование названий секций, автоширина колонок
    '    (значения ячеек не меняются)
    With ws.Range("A1:D1").Font
        .Bold = True
    End With
    With ws.Range("H1:O1").Font
        .Bold = True
    End With
    ws.Range(ws.Cells(2, 1), ws.Cells(2, 15)).Font.Bold = True
    ws.Range("A1").HorizontalAlignment = xlCenter
    ws.Range("H1").HorizontalAlignment = xlCenter
    On Error Resume Next
    ws.Rows("1:2").AutoFit
    ws.Columns("A:O").AutoFit
    On Error GoTo ErrHandler

    ' 9. Умная таблица (ListObject): создание или обновление диапазона
    '    (W7.1a: идемпотентно, без дубликатов, стиль TableStyleMedium6)
    Call EnsurePatternsListObject(ws)

    ' 10. Защита листа: разблокированы только области ввода данных
    '    (ячейки вне A2:D1000 и H2:O1000 блокируются заново — после
    '     Unprotect свойство Locked могло быть сброшено пользователем)
    ws.Cells.Locked = True
    ws.Range("A2:D" & CStr(PATTERNS_LAST_INPUT_ROW)).Locked = False
    ' PRE-04c-fix2: заголовки строки 2 секции H-O защищены (область ввода
    ' данных — строки 3..1000, как и диапазоны Validation)
    ws.Range("H3:O" & CStr(PATTERNS_LAST_INPUT_ROW)).Locked = False
    ws.Protect Password:=PATTERNS_PASSWORD

    If created Then
        WriteLog "EnsurePatternsSheet: лист '" & SHEET_NAME & "' создан " & _
                 "(названия секций A1/H1, заголовки строки 2, предзаполнение " & _
                 "исключений со строки 3, Validation, защита)"
    Else
        WriteLog "EnsurePatternsSheet: лист '" & SHEET_NAME & "' обновлён без " & _
                 "очистки (проверены заголовки, восстановлены Validation и " & _
                 "защита; данные A2:D1000 и H2:O1000 сохранены)"
    End If
    Exit Sub

ErrHandler:
    WriteError "PT001", "EnsurePatternsSheet: ошибка подготовки листа '" & _
               SHEET_NAME & "'. Ошибка: " & Err.Description
    MsgBox "Ошибка PT001: не удалось подготовить лист '" & SHEET_NAME & "'." & _
           vbCrLf & Err.Description, vbCritical, SHEET_NAME
    Err.Clear
End Sub

' ---------------------------------------------------------------------------
' Навешивает выпадающий список (xlValidateList) на диапазон ячеек.
' При ошибке применения пишется предупреждение P002 (подготовка листа
' продолжается: списки можно настроить вручную).
' ---------------------------------------------------------------------------
Private Sub ApplyListValidation(ByVal rng As Range, ByVal listItems As String)
    On Error GoTo FailHandler
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=listItems
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    Exit Sub
FailHandler:
    WriteWarning "PT002", "EnsurePatternsSheet: не удалось применить список """ & _
               listItems & """ к диапазону " & rng.Address(False, False) & _
               ". Ошибка: " & Err.Description
    Err.Clear
End Sub

' ---------------------------------------------------------------------------
' Пишет название секции в строку 1 и объединяет ячейки (PRE-04c-fix).
' Если в целевой ячейке уже лежит заголовок старого каркаса PRE-04c или
' другое чужое значение - текст НЕ перезаписывается (переносом заголовков
' занимается MigrateLegacyLayout); объединение выполняется только если
' прочие ячейки диапазона пусты.
' ---------------------------------------------------------------------------
Private Sub EnsureSectionTitle(ByVal ws As Worksheet, ByVal targetCell As String, _
                               ByVal titleText As String, ByVal mergeRange As String)
    Dim cur As String
    On Error GoTo FailHandler
    cur = Trim(CStr(ws.Range(targetCell).Value))
    If cur = "" Then
        ws.Range(targetCell).Value = titleText
    ElseIf StrComp(cur, titleText, vbTextCompare) <> 0 Then
        ' в ячейке названия секции уже есть значение (старый каркас или
        ' правка пользователя) - не трогаем данные
        Exit Sub
    End If
    With ws.Range(mergeRange)
        If Not .MergeCells Then
            If Application.WorksheetFunction.CountA(.Cells) <= 1 Then .Merge
        End If
    End With
    Exit Sub
FailHandler:
    WriteWarning "PT002", "EnsurePatternsSheet: не удалось оформить название " & _
               "секции '" & titleText & "' (" & targetCell & "). Ошибка: " & _
               Err.Description
    Err.Clear
End Sub

' ---------------------------------------------------------------------------
' Проверяет, стоят ли заголовки каркаса в указанной строке: достаточно
' совпадения A<row> = «Паттерн» или H<row> = «Область».
' ---------------------------------------------------------------------------
Private Function HasHeaderAt(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    On Error Resume Next
    If Trim(CStr(ws.Cells(rowNum, 1).Value)) = HDR_A Then
        HasHeaderAt = True
    ElseIf Trim(CStr(ws.Cells(rowNum, 8).Value)) = HDR_H Then
        HasHeaderAt = True
    End If
    On Error GoTo 0
End Function

' ---------------------------------------------------------------------------
' Дописывает заголовок каркаса в ячейку (строка rowNum, колонка colNum),
' если она пуста. Существующие значения (в т.ч. пользовательские) не
' перезаписываются.
' ---------------------------------------------------------------------------
Private Sub EnsureHeaderCell(ByVal ws As Worksheet, ByVal rowNum As Long, _
                             ByVal colNum As Long, ByVal hdrText As String)
    Dim cur As String
    On Error GoTo FailHandler
    cur = Trim(CStr(ws.Cells(rowNum, colNum).Value))
    If cur = "" Then
        ws.Cells(rowNum, colNum).Value = hdrText
    End If
    Exit Sub
FailHandler:
    WriteWarning "PT002", "EnsurePatternsSheet: не удалось дописать заголовок " & _
               "'" & hdrText & "' в " & ws.Cells(rowNum, colNum).Address(False, False) & _
               ". Ошибка: " & Err.Description
    Err.Clear
End Sub

' ---------------------------------------------------------------------------
' Перенос каркаса со старой раскладки PRE-04c (заголовки в строке 1,
' предзаполнение с строки 2) на новую (названия секций в строке 1,
' заголовки в строке 2, данные с строки 3). Строки 2..lastRow смещаются
' вниз на одну строку вместе со значениями — данные сохраняются, ничего
' не удаляется. Объединения строк 1 снимаются вызывающей процедурой до
' переноса; названия секций дописываются в строку 1 после переноса.
' ---------------------------------------------------------------------------
Private Sub MigrateLegacyLayout(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim aLast As Long
    Dim hoLast As Long
    On Error GoTo FailHandler
    aLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    hoLast = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row
    lastRow = aLast
    If hoLast > lastRow Then lastRow = hoLast
    If lastRow < 2 Then lastRow = 2   ' заголовки всё равно переносятся в строку 2
    If lastRow > PATTERNS_LAST_INPUT_ROW Then lastRow = PATTERNS_LAST_INPUT_ROW
    ' смещение заполненных строк каркаса (2..lastRow) на одну строку вниз
    ws.Range(ws.Rows(2), ws.Rows(lastRow)).Cut Destination:=ws.Rows(3)
    ' очистка ячеек строки 1 от старых заголовков (названия секций будут
    ' дописаны вызывающей процедурой)
    ws.Range("A1:O1").ClearContents
    Exit Sub
FailHandler:
    WriteWarning "PT002", "EnsurePatternsSheet: не удалось перенести старый " & _
               "каркас листа '" & SHEET_NAME & "' (заголовки из строки 1 в " & _
               "строку 2). Ошибка: " & Err.Description
    Err.Clear
End Sub

' ---------------------------------------------------------------------------
' Удаление старой строки правила «Год» (Категория / Отображаемое / Точное /
' Год) на существующем листе (PRE-04c-fix2). Процедура срабатывает только
' если в строке 4 секции H-O лежит именно старое эталонное правило «Год»
' (т.е. пользователь его не редактировал): все заполненные строки ниже
' (5..lastRow) поднимаются на одну строку вверх, последняя заполненная
' строка очищается. Если правило изменено пользователем или отсутствует —
' лист не затрагивается. Данные пользователя при этом сохраняются:
' удаляется только копия устаревшего эталона предзаполнения.
' ---------------------------------------------------------------------------
Private Sub ShiftLegacyYearRule(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim hoLast As Long
    Dim aLast As Long
    Dim srcRow As Long
    On Error GoTo FailHandler
    ' проверка старого эталона правила «Год» в строке 4 (колонки H..O = 8..15)
    If Trim(CStr(ws.Cells(4, 8).Value)) <> "Категория" Then Exit Sub
    If Trim(CStr(ws.Cells(4, 9).Value)) <> "Отображаемое" Then Exit Sub
    If StrComp(Trim(CStr(ws.Cells(4, 10).Value)), "Точное", vbBinaryCompare) <> 0 Then Exit Sub
    If Trim(CStr(ws.Cells(4, 11).Value)) <> "Год" Then Exit Sub
    If StrComp(Trim(CStr(ws.Cells(4, 12).Value)), "ВКЛЮЧАТЬ", vbBinaryCompare) <> 0 Then Exit Sub
    ' границы диапазона секции H-O (строки 3..PATTERNS_LAST_INPUT_ROW)
    hoLast = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row
    aLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastRow = hoLast
    If aLast > lastRow Then lastRow = aLast
    If lastRow > PATTERNS_LAST_INPUT_ROW Then lastRow = PATTERNS_LAST_INPUT_ROW
    ' поднять строки 5..lastRow на одну строку вверх (значения сохраняются)
    For srcRow = 5 To lastRow
        ws.Range(ws.Cells(srcRow - 1, 8), ws.Cells(srcRow - 1, 15)).Value = _
            ws.Range(ws.Cells(srcRow, 8), ws.Cells(srcRow, 15)).Value
    Next srcRow
    ' очистить нижнюю строку (копия предыдущей последней строки)
    ws.Range(ws.Cells(lastRow, 8), ws.Cells(lastRow, 15)).ClearContents
    WriteLog "EnsurePatternsSheet: удалена старая строка правила 'Год' " & _
             "(Категория/Точное/Год) из предзаполнения листа '" & SHEET_NAME & "'"
    Exit Sub
FailHandler:
    WriteWarning "PT002", "EnsurePatternsSheet: не удалось удалить старую " & _
               "строку правила 'Год' на листе '" & SHEET_NAME & "'. Ошибка: " & _
               Err.Description
    Err.Clear
End Sub

' ---------------------------------------------------------------------------
' Точечная сверка одного значения предзаполнения на существующем листе
' (PRE-04c-fix, без очистки и без пересоздания данных): если в ячейке
' (rowNum, colNum) секции H-O всё ещё лежит старое эталонное значение
' oldValue (т.е. правило не редактировалось пользователем), оно заменяется
' на эталон фикса newValue. Пустые и изменённые ячейки не затрагиваются.
' ---------------------------------------------------------------------------
Private Sub SyncPrefillCell(ByVal ws As Worksheet, ByVal rowNum As Long, _
                            ByVal colNum As Long, ByVal oldValue As String, _
                            ByVal newValue As String)
    Dim cur As String
    On Error GoTo FailHandler
    If rowNum < 3 Or rowNum > PATTERNS_LAST_INPUT_ROW Then Exit Sub
    cur = Trim(CStr(ws.Cells(rowNum, colNum).Value))
    If cur <> "" And StrComp(cur, oldValue, vbBinaryCompare) = 0 Then
        If StrComp(cur, newValue, vbBinaryCompare) <> 0 Then
            ws.Cells(rowNum, colNum).Value = newValue
        End If
    End If
    Exit Sub
FailHandler:
    WriteWarning "PT002", "EnsurePatternsSheet: не удалось сверить значение " & _
               "предзаполнения в " & ws.Cells(rowNum, colNum).Address(False, False) & _
               ". Ошибка: " & Err.Description
    Err.Clear
End Sub

' ---------------------------------------------------------------------------
' Возвращает Worksheet листа «Паттерны», если лист существует в книге;
' если листа нет - Nothing. Только поиск, без создания и без изменения
' чего-либо (W7.1-A).
' ---------------------------------------------------------------------------
Private Function GetPatternsSheetOrNull() As Worksheet
    Dim wsAny As Worksheet
    For Each wsAny In ThisWorkbook.Worksheets
        If StrComp(wsAny.Name, SHEET_NAME, vbTextCompare) = 0 Then
            Set GetPatternsSheetOrNull = wsAny
            Exit Function
        End If
    Next wsAny
    Set GetPatternsSheetOrNull = Nothing
End Function

' ---------------------------------------------------------------------------
' Чтение одной ячейки секции H-O листа «Паттерны» (строка rowNum,
' колонка colNum) с проверкой границ диапазона ввода
' (строки 3..PATTERNS_LAST_INPUT_ROW).
' Значение ячейки приводится к тексту через CStr + Trim (только чтение).
' ---------------------------------------------------------------------------
Private Function ExcCell(ByVal ws As Worksheet, ByVal rowNum As Long, _
                         ByVal colNum As Long) As String
    If rowNum < 3 Or rowNum > PATTERNS_LAST_INPUT_ROW Then
        ExcCell = ""
    Else
        ExcCell = Trim(CStr(ws.Cells(rowNum, colNum).Value))
    End If
End Function

' ---------------------------------------------------------------------------
' W7.1-A: GetPatternExceptions — чтение секции H-O («ИСКЛЮЧЕНИЯ
' ПАТТЕРНОВ») листа «Паттерны» (только чтение: лист не изменяется,
' защита не снимается).
'   Источник: заголовки в строке 2, данные со строки 3; колонки
'   H=Область, I=Поле, J=Тип матчинга (Точное/Маска), K=Значение,
'   L=Действие (ВКЛЮЧАТЬ/ИСКЛЮЧАТЬ), M=Примеры, N=Причина,
'   O=Статус (активно/отменено).
'   Фильтр: пропускаются пустые строки и строки, где O <> "активно".
'   Формат строк вывода (J = "Маска" — с суффиксом " (маска)" и
'   оператором "~"; иначе — "Точное" и оператор "="):
'     "- <H>: <I> = '<K>' -> <L>"          (I непусто, Точное)
'     "- <H>: <I> ~ '<K>' -> <L> (маска)"  (I непусто, Маска)
'     "- <H> = '<K>' -> <L>"               (I пусто, Точное)
'     "- <H> ~ '<K>' -> <L> (маска)"       (I пусто, Маска)
'   Первая строка результата (при наличии хотя бы одной строки):
'   "Правила обработки паттернов:".
'   Лист отсутствует или секция H-O пуста (нет ни одной активной строки)
'   — вернуть "" и WriteWarning PT003.
' ---------------------------------------------------------------------------
Public Function GetPatternExceptions() As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim resLines() As String
    Dim nRes As Long
    Dim sArea As String
    Dim sField As String
    Dim sMatch As String
    Dim sValue As String
    Dim sAction As String
    Dim sStatus As String
    Dim sLine As String
    Dim sOp As String
    Dim sSuffix As String

    GetPatternExceptions = ""

    On Error GoTo ErrHandler

    ' 1. Лист «Паттерны» (только чтение, без создания/подготовки)
    Set ws = GetPatternsSheetOrNull()
    If ws Is Nothing Then
        WriteWarning "PT003", "GetPatternExceptions: лист '" & SHEET_NAME & _
                   "' отсутствует в книге - исключения паттернов не читаны."
        Exit Function
    End If

    ' 2. Нижняя граница данных секции H-O (колонка H = 8)
    lastRow = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row
    If lastRow > PATTERNS_LAST_INPUT_ROW Then lastRow = PATTERNS_LAST_INPUT_ROW

    ' 3. Обход данных (со строки 3): фильтр пустых строк и O <> "активно"
    ReDim resLines(1 To 1)
    nRes = 0
    For r = 3 To lastRow
        sArea = ExcCell(ws, r, 8)    ' H - Область
        sField = ExcCell(ws, r, 9)   ' I - Поле
        sMatch = ExcCell(ws, r, 10)  ' J - Тип матчинга
        sValue = ExcCell(ws, r, 11)  ' K - Значение
        sAction = ExcCell(ws, r, 12) ' L - Действие
        sStatus = ExcCell(ws, r, 15) ' O - Статус
        ' пустая строка (нет ключевых значений) - пропустить
        If sArea <> "" Or sValue <> "" Then
            ' неактивные правила (отменено / пустой статус) - пропустить
            If StrComp(sStatus, "активно", vbTextCompare) = 0 Then
                If StrComp(sMatch, "Маска", vbTextCompare) = 0 Then
                    sOp = "~"
                    sSuffix = " (маска)"
                Else
                    ' "Точное" и любые прочие значения J — формат «Точное»
                    sOp = "="
                    sSuffix = ""
                End If
                If sField <> "" Then
                    sLine = "- " & sArea & ": " & sField & " " & sOp & _
                            "'" & sValue & "' -> " & sAction & sSuffix
                Else
                    sLine = "- " & sArea & " " & sOp & _
                            "'" & sValue & "' -> " & sAction & sSuffix
                End If
                nRes = nRes + 1
                If nRes > UBound(resLines) Then
                    ReDim Preserve resLines(1 To nRes)
                End If
                resLines(nRes) = sLine
            End If
        End If
    Next r

    ' 4. Пустая секция H-O (ни одной активной строки) - "" + PT003
    If nRes = 0 Then
        WriteWarning "PT003", "GetPatternExceptions: секция H-O листа '" & _
                   SHEET_NAME & "' пуста (нет активных правил) - " & _
                   "исключения паттернов не сформированы."
        Exit Function
    End If

    ' 5. Сборка результата: заголовок + строки правил
    Dim sOut As String
    sOut = "Правила обработки паттернов:"
    Dim i As Long
    For i = 1 To nRes
        sOut = sOut & vbCrLf & resLines(i)
    Next i
    GetPatternExceptions = sOut
    Exit Function

ErrHandler:
    WriteWarning "PT003", "GetPatternExceptions: ошибка чтения листа '" & _
               SHEET_NAME & "'. Ошибка: " & Err.Description
    GetPatternExceptions = ""
    Err.Clear
End Function

' ---------------------------------------------------------------------------
' W7.1-A: GetKnownPatterns — чтение секции A-D («СПИСОК ПАТТЕРНОВ») листа
' «Паттерны» (только чтение: лист не изменяется, защита не снимается).
'   Источник: заголовки в строке 2, данные со строки 3; используется
'   только колонка A («Паттерн»).
'   Формат: заголовок "Известные паттерны:" (первой строкой, только если
'   есть хотя бы одна строка данных), далее каждая непустая строка
'   колонки A с префиксом "- ".
'   Пустая секция A-D — вернуть "" (без заголовка, без предупреждения);
'   лист отсутствует — вернуть "" и WriteWarning PT003.
' ---------------------------------------------------------------------------
Public Function GetKnownPatterns() As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim pat As String
    Dim sOut As String
    Dim hasData As Boolean

    GetKnownPatterns = ""

    On Error GoTo ErrHandler

    ' 1. Лист «Паттерны» (только чтение, без создания/подготовки)
    Set ws = GetPatternsSheetOrNull()
    If ws Is Nothing Then
        WriteWarning "PT003", "GetKnownPatterns: лист '" & SHEET_NAME & _
                   "' отсутствует в книге - список известных паттернов не читан."
        Exit Function
    End If

    ' 2. Нижняя граница данных секции A-D (колонка A = 1)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > PATTERNS_LAST_INPUT_ROW Then lastRow = PATTERNS_LAST_INPUT_ROW

    ' 3. Обход данных (со строки 3): непустые значения колонки A
    hasData = False
    sOut = ""
    For r = 3 To lastRow
        If r >= 3 And r <= PATTERNS_LAST_INPUT_ROW Then
            pat = Trim(CStr(ws.Cells(r, 1).Value))
            If pat <> "" Then
                If Not hasData Then
                    sOut = "Известные паттерны:"
                    hasData = True
                End If
                sOut = sOut & vbCrLf & "- " & pat
            End If
        End If
    Next r

    ' 4. Пустая секция A-D - вернуть "" (без заголовка)
    If Not hasData Then
        GetKnownPatterns = ""
        Exit Function
    End If

    GetKnownPatterns = sOut
    Exit Function

ErrHandler:
    WriteWarning "PT003", "GetKnownPatterns: ошибка чтения листа '" & _
               SHEET_NAME & "'. Ошибка: " & Err.Description
    GetKnownPatterns = ""
    Err.Clear
End Function

' ---------------------------------------------------------------------------
' W7.1a: Обеспечивает наличие умной таблицы (ListObject) на листе «Паттерны».
' Идемпотентна: если таблица с именем PATTERNS_TABLE_NAME существует,
' проверяет и при необходимости расширяет её диапазон до PATTERNS_LAST_INPUT_ROW.
' Если таблицы нет — создаёт её на диапазоне A2:O<PATTERNS_LAST_INPUT_ROW>.
' Стиль таблицы: TableStyleMedium6 (как в tblAnomalies).
' ---------------------------------------------------------------------------
Private Sub EnsurePatternsListObject(ByVal ws As Worksheet)
    Dim lo As ListObject
    Dim targetRange As Range
    Dim tblName As String
    
    tblName = PATTERNS_TABLE_NAME
    Set targetRange = ws.Range("A2:O" & CStr(PATTERNS_LAST_INPUT_ROW))
    
    On Error Resume Next
    Set lo = ws.ListObjects(tblName)
    On Error GoTo 0
    
    If lo Is Nothing Then
        ' Таблица не найдена — создаём новую
        On Error GoTo ErrHandler
        Set lo = ws.ListObjects.Add(xlSrcRange, targetRange, , xlYes)
        lo.Name = tblName
        lo.TableStyle = "TableStyleMedium6"
        On Error GoTo 0
    Else
        ' Таблица существует — проверяем и при необходимости расширяем диапазон
        If lo.Range.Address <> targetRange.Address Then
            On Error GoTo ErrHandler
            lo.Resize targetRange
            On Error GoTo 0
        End If
    End If
    
    Exit Sub
ErrHandler:
    WriteWarning "PT002", "EnsurePatternsListObject: не удалось создать или изменить " & _
               "умную таблицу '" & tblName & "'. Ошибка: " & Err.Description
    Err.Clear
End Sub
