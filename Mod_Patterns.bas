Attribute VB_Name = "Mod_Patterns"
Option Explicit

' ============================================================================
' Модуль: Mod_Patterns
' Назначение: лист «Паттерны» (PRE-04c) — каркас листа для Промта №0.2
'             («Паттерны и справочники»):
'   Секция A-D  («Список паттернов», append-only): заголовки
'               «Паттерн», «Типы», «Количество», «Итерация»; строки 2+
'               пустые — заполняются кодом позже (разбор ответа ИИ, W7).
'   Секция H-O  («ИСКЛЮЧЕНИЯ паттернов»): заголовки «Область», «Поле»,
'               «Тип матчинга», «Значение», «Действие», «Примеры»,
'               «Причина», «Статус»; предзаполнены стартовые правила
'               (папки ЦИМ/Год — ВКЛЮЧАТЬ; базовые системные типы —
'               ИСКЛЮЧАТЬ).
' Выпадающие списки (Validation) навешаны на колонки H, I, J, L, O
' (строки 2..PATTERNS_LAST_INPUT_ROW); защита листа паролем "pilot",
' редактируемыми остаются только диапазоны ввода данных A2:D1000 и H2:O1000.
' Зависимости: Mod_Util (WriteLog / WriteError / WriteWarning).
' Реестр кодов ошибок/предупреждений модуля (единый реестр проекта):
'   PT001 - ошибка подготовки листа «Паттерны» (EnsurePatternsSheet);
'   PT002 - предупреждение: не удалось применить Validation к ячейкам
'          секции H-O (лист подготовлен, списки нужно настроить вручную).
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
' Идемпотентна: повторный вызов пересоздаёт каркас (Cells.Clear), список
' паттернов (A2:D...) заполняется кодом позже и затирается только при
' явном пересоздании листа этой процедурой.
' ---------------------------------------------------------------------------
Public Sub EnsurePatternsSheet()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsAny As Worksheet
    Dim found As Boolean
    Dim excData As Variant
    Dim r As Long
    Dim c As Long
    Dim lastCol As Long

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook

    ' 1. Поиск листа «Паттерны»; если нет - создать в конце книги
    For Each wsAny In wb.Worksheets
        If StrComp(wsAny.Name, SHEET_NAME, vbTextCompare) = 0 Then
            Set ws = wsAny
            found = True
            Exit For
        End If
    Next wsAny

    If Not found Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_NAME
    Else
        ws.Unprotect Password:=PATTERNS_PASSWORD   ' снять защиту перед правкой
    End If

    ' 2. Полная очистка листа (каркас пересоздаётся целиком)
    ws.Cells.Clear

    ' 3. Секция A-D: список паттернов (заголовки; строки 2+ пустые —
    '    заполняются кодом позже, например разбором ответа ИИ)
    ws.Range("A1").Value = HDR_A
    ws.Range("B1").Value = HDR_B
    ws.Range("C1").Value = HDR_C
    ws.Range("D1").Value = HDR_D

    ' 4. Секция H-O: ИСКЛЮЧЕНИЯ паттернов (заголовки + предзаполнение)
    ws.Range("H1").Value = HDR_H
    ws.Range("I1").Value = HDR_I
    ws.Range("J1").Value = HDR_J
    ws.Range("K1").Value = HDR_K
    ws.Range("L1").Value = HDR_L
    ws.Range("M1").Value = HDR_M
    ws.Range("N1").Value = HDR_N
    ws.Range("O1").Value = HDR_O

    ' Матрица предзаполненных правил: 8 строк x 8 колонок (H..O)
    excData = Array( _
        Array("Категория", "Отображаемое", "Точное", "Папка ЦИМ", "ВКЛЮЧАТЬ", _
              "КЦИМ; ППЦИМ", "позиционный паттерн", "активно"), _
        Array("Категория", "Отображаемое", "Точное", "Год", "ВКЛЮЧАТЬ", _
              "Год 2024; Год 2025", "позиционный паттерн", "активно"), _
        Array("Тип", "Системное", "маска", "bim_*", "ИСКЛЮЧАТЬ", _
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
            ws.Cells(r + 2, c + 8).Value = CStr(excData(r)(c))
        Next c
    Next r

    ' 5. Выпадающие списки (Validation) для колонок H, I, J, L, O
    ApplyListValidation ws.Range("H2:H" & CStr(PATTERNS_LAST_INPUT_ROW)), "Тип,Категория"
    ApplyListValidation ws.Range("I2:I" & CStr(PATTERNS_LAST_INPUT_ROW)), "Отображаемое,Системное"
    ApplyListValidation ws.Range("J2:J" & CStr(PATTERNS_LAST_INPUT_ROW)), "Точное,маска"
    ApplyListValidation ws.Range("L2:L" & CStr(PATTERNS_LAST_INPUT_ROW)), "ВКЛЮЧАТЬ,ИСКЛЮЧАТЬ"
    ApplyListValidation ws.Range("O2:O" & CStr(PATTERNS_LAST_INPUT_ROW)), "активно,отменено"

    ' 6. Форматирование: заголовки жирным, автоширина колонок
    lastCol = 15 ' колонка O
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Font
        .Bold = True
    End With
    ws.Rows(1).AutoFit
    On Error Resume Next
    ws.Columns("A:O").AutoFit
    On Error GoTo ErrHandler

    ' 7. Защита листа: разблокированы только области ввода данных
    ws.Cells.Locked = True
    ws.Range("A2:D" & CStr(PATTERNS_LAST_INPUT_ROW)).Locked = False
    ws.Range("H2:O" & CStr(PATTERNS_LAST_INPUT_ROW)).Locked = False
    ws.Protect Password:=PATTERNS_PASSWORD

    WriteLog "EnsurePatternsSheet: лист '" & SHEET_NAME & "' создан/обновлён " & _
             "(секции A-D и H-O, Validation, защита)"
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
