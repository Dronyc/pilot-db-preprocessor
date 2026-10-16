Attribute VB_Name = "Mod_Util"
Option Explicit

' ============================================================================
' Модуль утилит: пути проекта, журналирование, чтение конфигурации.
' Реализация W4 (разделы 2, 5 спецификации Spec_Dispatcher.md).
' Пути строятся как в Mod_Registry / Mod_Prompt: <корень>\NN_Имя_папки,
' корень берётся из именованного диапазона ProjectRoot листа "Настройки".
' Коды ошибок модуля: U003 - ошибка подготовки листа "Настройки",
' U004 - ошибка выбора папки проекта, U002 - ошибка чтения configuration.txt.
' ============================================================================

' ---------------------------------------------------------------------------
' КОНСТАНТЫ И ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (Private)
' ---------------------------------------------------------------------------

' Имена именованных диапазонов листа "Настройки" (раздел 5 спецификации).
Private Const NAME_CHECK_MODE As String = "CheckMode"
Private Const NAME_AUTO_OPEN_RESULTS As String = "AutoOpenResults"
Private Const NAME_PROMPT_GEN_VERSION As String = "PromptGenVersion"
Private Const NAME_PROMPT_FIX_VERSION As String = "PromptFixVersion"

' Пароль защиты служебных листов (единый для книги Препроцессора).
Private Const SHEET_PASSWORD As String = "pilot"

' Возвращает путь без завершающего разделителя "\".
Private Function NormPath(ByVal p As String) As String
    Do While Len(p) > 1 And Right(p, 1) = "\"
        p = Left(p, Len(p) - 1)
    Loop
    NormPath = p
End Function

' ---------------------------------------------------------------------------
' ПУБЛИЧЕСКИЕ ПРОЦЕДУРЫ И ФУНКЦИИ
' ---------------------------------------------------------------------------

' Возвращает корневую папку проекта из именованного диапазона ProjectRoot
' на листе "Настройки". Если диапазон не найден или значение пустое -
' пишет в лог "GetProjectRoot: ProjectRoot не задан" и возвращает "".
Public Function GetProjectRoot() As String
    ' Корень проекта — всегда папка, где лежит файл .xlsm.
    ' Если книга не сохранена (ThisWorkbook.Path = "") — возвращаем "".
    If ThisWorkbook.path <> "" Then
        GetProjectRoot = NormPath(ThisWorkbook.path)
    Else
        GetProjectRoot = ""
    End If
End Function
' Возвращает путь к папке логов: <корень>\05_Логи.
Public Function GetLogsFolder() As String
    Dim root As String
    root = GetProjectRoot()
    If root = "" Then
        GetLogsFolder = ""
    Else
        GetLogsFolder = root & "\05_Логи"
    End If
End Function

' Возвращает путь к папке результатов: <корень>\03_Результаты генерации.
Public Function GetResultsFolder() As String
    Dim root As String
    root = GetProjectRoot()
    If root = "" Then
        GetResultsFolder = ""
    Else
        GetResultsFolder = root & "\03_Результаты генерации"
    End If
End Function

' Возвращает путь к папке промптов: <корень>\04_Промты.
Public Function GetPromptsFolder() As String
    Dim root As String
    root = GetProjectRoot()
    If root = "" Then
        GetPromptsFolder = ""
    Else
        GetPromptsFolder = root & "\04_Промты"
    End If
End Function

' Возвращает путь к папке шагов: <корень>\03_Результаты генерации\01_По шагам.
Public Function GetStepsFolder() As String
    Dim root As String
    root = GetProjectRoot()
    If root = "" Then
        GetStepsFolder = ""
    Else
        GetStepsFolder = root & "\03_Результаты генерации\01_По шагам"
    End If
End Function

' Записывает сообщение в журнал <корень>\05_Логи\Препроцессор_YYYY-MM-DD.log.
' Формат строки: [YYYY-MM-DD HH:NN:SS] <message>.
' ВАЖНО: обработка ошибок здесь сознательно НЕ применяется - WriteLog
' вызывается из обработчиков ошибок других процедур, и любая собственная
' ошибка могла бы зациклить журналирование. Если корень не задан или
' запись невозможна - молча выходит.
Public Sub WriteLog(ByVal message As String)
    Dim root As String
    Dim logsPath As String
    Dim logFile As String
    Dim fnum As Integer

    root = GetProjectRoot()
    If root = "" Then Exit Sub          ' корень не задан - молча выходим

    logsPath = root & "\05_Логи"
    If Dir(logsPath, vbDirectory) = "" Then MkDir logsPath

    logFile = logsPath & "\Препроцессор_" & Format(Date, "yyyy-mm-dd") & ".log"

    fnum = FreeFile
    Open logFile For Append As #fnum
    Print #fnum, "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] " & message
    Close #fnum
End Sub

' Записывает ошибку с кодом в журнал: "[ERROR <код>] <сообщение>".
Public Sub WriteError(ByVal code As String, ByVal message As String)
    WriteLog "[ERROR " & code & "] " & message
End Sub

' Записывает предупреждение с кодом в журнал: "[WARN <код>] <сообщение>".
Public Sub WriteWarning(ByVal code As String, ByVal message As String)
    WriteLog "[WARN " & code & "] " & message
End Sub

' Возвращает значение параметра конфигурации по ключу из файла
' <корень>\01_Исходные данные\03_Конфигурация\configuration.txt.
' Формат строки: "key=value" (без пробелов вокруг "="), регистр ключа значим.
' Если корень/файл/ключ не найдены - возвращает "".
Public Function GetConfigValue(ByVal key As String) As String
    Dim root As String
    Dim cfgPath As String
    Dim fnum As Integer
    Dim s As String
    Dim posEq As Long

    GetConfigValue = ""

    On Error GoTo ErrHandler

    root = GetProjectRoot()
    If root = "" Then Exit Function     ' корень не задан - тихо возвращаем ""

    cfgPath = root & "\01_Исходные данные\03_Конфигурация\configuration.txt"
    If Dir(cfgPath) = "" Then Exit Function   ' файла нет - возвращаем ""

    fnum = FreeFile
    Open cfgPath For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, s
        posEq = InStr(s, "=")
        If posEq > 0 Then
            ' сравнение ключа с учётом регистра ("key=value", без пробелов)
            If StrComp(Left(s, posEq - 1), key, vbBinaryCompare) = 0 Then
                GetConfigValue = Mid(s, posEq + 1)
                Close #fnum
                Exit Function
            End If
        End If
    Loop
    Close #fnum

    ' ключ не найден - GetConfigValue уже равно ""
    Exit Function

ErrHandler:
    WriteError "U002", "GetConfigValue: ошибка чтения configuration.txt. Ошибка: " & Err.Description
    GetConfigValue = ""
    Err.Clear
End Function

' Открывает указанную папку в проводнике Windows.
' Если путь пустой или папки нет - молча выходит.
Public Sub OpenFolderInExplorer(ByVal path As String)
    If path = "" Then Exit Sub
    If Dir(path, vbDirectory) = "" Then Exit Sub
    Shell "explorer.exe """ & path & """", vbNormalFocus
End Sub

' Создаёт лист "Настройки" (если его нет), заполняет поля, выпадающие
' списки и именованные диапазоны по разделу 5 спецификации Spec_Dispatcher.md:
'   B2 ProjectRoot       - Корневая папка проекта;
'   B3 CheckMode         - Режим проверки (Черновик / Строгая);
'   B4 AutoOpenResults   - Автооткрытие папки результатов (да / нет);
'   B5 PromptGenVersion  - Версия промта-генератора (по умолчанию 7.13);
'   B6 PromptFixVersion  - Версия промта-исправления (по умолчанию 1.0).
' Лист защищается паролем "pilot", редактируемыми остаются только B2:B6.
' Существующие значения пользовательских настроек при повторном запуске
' не перезаписываются.
Public Sub EnsureSettingsSheet()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim RowIdx As Long
    Dim labels As Variant
    Dim namesArr As Variant
    Dim defaults As Variant
    Dim i As Long

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook

    ' 1. Поиск листа "Настройки"; если нет - создать в конце книги
    On Error Resume Next
    Set ws = wb.Worksheets("Настройки")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ErrHandler
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "Настройки"
    Else
        On Error GoTo ErrHandler
        ws.Unprotect Password:=SHEET_PASSWORD   ' снимаем защиту перед правкой
    End If

    ' 2. Подписи (колонка A), имена диапазонов и значения по умолчанию -
    '    пять полей раздела 5 спецификации
    labels = Array("Режим проверки", _
               "Автооткрытие папки результатов", _
               "Версия промта-генератора", _
               "Версия промта-исправления")
    namesArr = Array(NAME_CHECK_MODE, NAME_AUTO_OPEN_RESULTS, _
                     NAME_PROMPT_GEN_VERSION, NAME_PROMPT_FIX_VERSION)
    defaults = Array("Черновик", "да", "7.13", "1.0")

    For i = 0 To 3
        RowIdx = i + 2
        ws.Cells(RowIdx, 1).value = CStr(labels(i))
        ' значение по умолчанию ставим только если ячейка пуста
        If Trim(CStr(ws.Cells(RowIdx, 2).value)) = "" Then
            ws.Cells(RowIdx, 2).value = CStr(defaults(i))
        End If
    Next i

    ws.Range("A1").value = "Настройки Препроцессора"
    ws.Columns("A:B").AutoFit

    ' 3. Выпадающие списки (Validation): B3 - режим проверки, B4 - автооткрытие
    With ws.Range("B3").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="Черновик,Строгая"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    With ws.Range("B4").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="да,нет"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' 4. Именованные диапазоны - создавать, если их ещё нет
    '    (существующие имена не трогаем, чтобы не переопределить ссылку)
    For i = 0 To 3
        On Error Resume Next
        Dim nmChk As Name
        Set nmChk = wb.Names(namesArr(i))
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo ErrHandler
            wb.Names.Add Name:=namesArr(i), _
                         RefersTo:="=" & ws.Name & "!$B$" & (i + 2)
        End If
        On Error GoTo ErrHandler
        Set nmChk = Nothing
    Next i

    ' 5. Защита листа от случайной правки: разблокированы только B2:B6
    ws.Cells.Locked = True
    ws.Range("B2:B5").Locked = False
    ws.Protect Password:=SHEET_PASSWORD

    WriteLog "EnsureSettingsSheet: лист 'Настройки' создан/обновлён"
    Exit Sub

ErrHandler:
    WriteError "U003", "EnsureSettingsSheet: ошибка подготовки листа 'Настройки'. Ошибка: " & Err.Description
    MsgBox "Ошибка U003: не удалось подготовить лист 'Настройки'." & vbCrLf & Err.Description, _
           vbCritical, "Настройки"
    Err.Clear
End Sub
