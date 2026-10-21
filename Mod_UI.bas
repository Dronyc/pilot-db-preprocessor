Attribute VB_Name = "Mod_UI"
Option Explicit

' ============================================================
' Модуль: Mod_UI
' Назначение: UI-слой проекта "Препроцессор структуры БД Pilot".
'             Создание кнопок на листе "Препроцессор", форматирование
'             листов, защита служебных листов, а также обработчики
'             нажатий кнопок, делегирующие работу профильным модулям.
' Зависимости: Mod_Util, Mod_Folders, Mod_Structure, Mod_Assembler,
'              Mod_Prompt, Mod_Docs.
' ============================================================

' ------------------------------------------------------------
' Восстановление UI: кнопки, оформление, защита листов.
' ------------------------------------------------------------
Public Sub RestoreUI()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim i As Long
    Dim rowBtn As Long
    Dim captions As Variant
    Dim actions As Variant
    Dim tips As Variant
    Dim wsAny As Worksheet
    Dim rng As Range
    Dim lastCol As Long
    Dim srv As Variant
    Dim wsSrv As Worksheet
    Dim pass As String
    Dim n As Long

    On Error GoTo ErrHandler

    ' Проверка наличия листа "Препроцессор"
    Set ws = Nothing
    For Each wsAny In ThisWorkbook.Worksheets
        If StrComp(wsAny.Name, "Препроцессор", vbTextCompare) = 0 Then
            Set ws = wsAny
            Exit For
        End If
    Next wsAny

    If ws Is Nothing Then
        MsgBox "Лист Препроцессор не найден", vbExclamation
        Exit Sub
    End If

    ' Создание служебных листов, если их нет
    Call EnsureSheet_Readme(ThisWorkbook)
    Call Mod_Util.EnsureSettingsSheet

    ' Удаление существующих Form Controls
    Do While ws.Buttons.Count > 0
        ws.Buttons(1).Delete
    Loop

    ' Удаление существующих кнопок
    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
    Loop
    
    ' Установка ширины столбцов До создания кнопок
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 60
    ws.Columns("C").Font.Italic = True
    ws.Columns("C").Font.Color = RGB(89, 89, 89)

    ' Подготовка списков кнопок
    captions = Array( _
        "Создать папки проекта", _
        "Проверить исходные данные", _
        "Собрать плоскую таблицу из Схемы", _
        "Повторный анализ преобразованного", _
        "Собрать типы и атрибуты (черновик)", _
        "Собрать типы и атрибуты (строгая)", _
        "Сделать промт следующего шага", _
        "Сделать промт исправления ошибок", _
        "Показать, что нашёл Препроцессор", _
        "Открыть папку Логи", _
        "Обновить Readme", _
        "Восстановить кнопки и оформление")

    actions = Array( _
        "CreateProjectFolders_Click", _
        "CheckSourceFiles_Click", _
        "BuildFlatStructure_Click", _
        "RecheckFlatTable_Click", _
        "AssembleDraft_Click", _
        "AssembleStrict_Click", _
        "MakeNextStepPrompt_Click", _
        "MakeFixPrompt_Click", _
        "ShowDiscoveryReport_Click", _
        "OpenLogsFolder_Click", _
        "UpdateReadme_Click", _
        "RestoreUI_Click")

    tips = Array( _
        "Создаёт структуру папок проекта", _
        "Проверяет наличие файлов в 01_Исходные данные", _
        "Преобразует исходную схему в плоскую таблицу", _
        "Анализирует готовую плоскую таблицу без пересборки", _
        "Собирает сводную в черновом режиме", _
        "Собирает сводную в строгом режиме", _
        "Формирует промт следующего шага и создаёт его папку", _
        "Формирует промт исправления для текущего шага", _
        "Показывает карту найденных шагов и файлов", _
        "Открывает папку 05_Логи", _
        "Формирует Readme.md из листа Readme", _
        "Пересоздаёт кнопки, оформление и защиту листов")

    ' Создание 12 кнопок в столбце B (строки 2, 4, ..., 24)
    ' Используем Shapes.AddFormControl + Placement = xlFreeFloating,
    ' чтобы кнопки не растягивались при изменении ширины столбцов и AutoFit.
    rowBtn = 2
    For i = LBound(captions) To UBound(captions)
        Dim shpBtn As Shape
        Set shpBtn = ws.Shapes.AddFormControl( _
            xlButtonControl, _
            ws.Cells(rowBtn, 2).Left, _
            ws.Cells(rowBtn, 2).Top, _
            200, 22)
        shpBtn.Placement = xlFreeFloating
        shpBtn.Name = "btnPilot_" & CStr(i + 1)
        shpBtn.OnAction = CStr(actions(i))
        With shpBtn.TextFrame
            .Characters.text = CStr(captions(i))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        shpBtn.Fill.ForeColor.RGB = RGB(91, 155, 213)      ' синий
        shpBtn.Line.ForeColor.RGB = RGB(0, 0, 0)           ' чёрная рамка
        shpBtn.Line.Weight = 1
        ws.Cells(rowBtn, 3).value = CStr(tips(i))
        rowBtn = rowBtn + 2
    Next i
    
    ' Форматирование всех листов книги
    For Each wsAny In ThisWorkbook.Worksheets
        If StrComp(wsAny.Name, "Препроцессор", vbTextCompare) <> 0 Then
            On Error Resume Next
            wsAny.Unprotect Password:="pilot"
            On Error GoTo ErrHandler
    
            ' Первая строка - жирным
            wsAny.Rows(1).Font.Bold = True
    
            ' Автоширина столбцов
            On Error Resume Next
            wsAny.Cells.EntireColumn.AutoFit
            On Error GoTo ErrHandler
    
            ' Закрепление первой строки
            On Error Resume Next
            wsAny.Activate
            wsAny.Range("A2").Select
            ActiveWindow.FreezePanes = False
            ActiveWindow.FreezePanes = True
            On Error GoTo ErrHandler
        End If
    Next wsAny

    ' Защита служебных листов
    pass = "pilot"
    srv = Array("_Реестр_типов", "_Реестр_покрытия", "_Ошибки", "_Предупреждения")
    For n = LBound(srv) To UBound(srv)
        Set wsSrv = Nothing
        For Each wsAny In ThisWorkbook.Worksheets
            If StrComp(wsAny.Name, CStr(srv(n)), vbTextCompare) = 0 Then
                Set wsSrv = wsAny
                Exit For
            End If
        Next wsAny
        If Not wsSrv Is Nothing Then
            On Error Resume Next
            wsSrv.Protect Password:=pass
            On Error GoTo ErrHandler
        End If
    Next n

    ' Активируем лист "Препроцессор"
    On Error Resume Next
    ws.Activate
    On Error GoTo ErrHandler
    
    ' Упорядочить листы: Препроцессор, Настройки, Readme, остальные
    Call ReorderSheets(ThisWorkbook)

    Mod_Util.WriteLog "UI восстановлен: 12 кнопок, оформление, защита"
    Exit Sub

ErrHandler:
    Mod_Util.WriteLog "Ошибка в RestoreUI: " & Err.Description
    MsgBox "Ошибка: " & Err.Description, vbExclamation
    Err.Clear
End Sub

' ------------------------------------------------------------
' 1. Создать папки проекта
' ------------------------------------------------------------
Public Sub CreateProjectFolders_Click()
    On Error GoTo ErrHandler

    Call Mod_Folders.CreateProjectFolders
    Mod_Util.WriteLog "CreateProjectFolders_Click: папки созданы"
    Exit Sub

ErrHandler:
    Mod_Util.WriteError "F003", "CreateProjectFolders_Click: " & Err.Description
    MsgBox "Ошибка F003: " & Err.Description, vbCritical, "Папки проекта"
    Err.Clear
End Sub

' ------------------------------------------------------------
' 2. Проверить исходные данные
' ------------------------------------------------------------
Public Sub CheckSourceFiles_Click()
    Dim ok As Boolean

    On Error GoTo ErrHandler

    ok = Mod_Folders.CheckSourceFolders
    If ok Then
        MsgBox "Все исходные файлы найдены.", vbInformation, "Проверка исходных данных"
    Else
        MsgBox "Есть проблемы. Смотрите лог 05_Логи.", vbExclamation, "Проверка исходных данных"
    End If
    Exit Sub

ErrHandler:
    Mod_Util.WriteError "F004", "CheckSourceFiles_Click: " & Err.Description
    MsgBox "Ошибка F004: " & Err.Description, vbCritical, "Проверка исходных данных"
    Err.Clear
End Sub

' ------------------------------------------------------------
' 3. Собрать плоскую таблицу из Схемы
' ------------------------------------------------------------
Public Sub BuildFlatStructure_Click()
    On Error GoTo ErrHandler
    Mod_Structure.BuildFlatStructure
    Mod_Util.WriteLog "Плоская таблица собрана"
    MsgBox "Плоская таблица собрана"
    Exit Sub
ErrHandler:
    Mod_Util.WriteLog "Ошибка в BuildFlatStructure_Click: " & Err.Description
    MsgBox "Ошибка: " & Err.Description, vbExclamation
    Err.Clear
End Sub

' ------------------------------------------------------------
' 4. Повторный анализ преобразованного
' ------------------------------------------------------------
Public Sub RecheckFlatTable_Click()
    On Error GoTo ErrHandler
    Mod_Util.WriteLog "Повторный анализ: функция появится в W4."
    MsgBox "Функция появится в W4"
    Exit Sub
ErrHandler:
    Mod_Util.WriteLog "Ошибка в RecheckFlatTable_Click: " & Err.Description
    MsgBox "Ошибка: " & Err.Description, vbExclamation
    Err.Clear
End Sub

' ------------------------------------------------------------
' 5. Собрать типы и атрибуты (черновик)
' ------------------------------------------------------------
Public Sub AssembleDraft_Click()
    On Error GoTo ErrHandler
    Mod_Assembler.ValidateAndAssemble True
    Mod_Util.WriteLog "Сводная собрана (черновик)"
    MsgBox "Сводная собрана (черновик)"
    Exit Sub
ErrHandler:
    Mod_Util.WriteLog "Ошибка в AssembleDraft_Click: " & Err.Description
    MsgBox "Ошибка: " & Err.Description, vbExclamation
    Err.Clear
End Sub

' ------------------------------------------------------------
' 6. Собрать типы и атрибуты (строгая)
' ------------------------------------------------------------
Public Sub AssembleStrict_Click()
    On Error GoTo ErrHandler
    Mod_Assembler.ValidateAndAssemble False
    Mod_Util.WriteLog "Сводная собрана (строгая)"
    MsgBox "Сводная собрана (строгая)"
    Exit Sub
ErrHandler:
    Mod_Util.WriteLog "Ошибка в AssembleStrict_Click: " & Err.Description
    MsgBox "Ошибка: " & Err.Description, vbExclamation
    Err.Clear
End Sub

' ------------------------------------------------------------
' 7. Сделать промт следующего шага
' ------------------------------------------------------------
Public Sub MakeNextStepPrompt_Click()
    On Error GoTo ErrHandler
    ' Работа выполняется в Mod_Prompt (разделы 9 и 13 спецификации)
    Call Mod_Prompt.MakeNextStepPrompt
    Mod_Util.WriteLog "MakeNextStepPrompt_Click: запуск формирования промта следующего шага"
    Exit Sub
ErrHandler:
    Mod_Util.WriteError "P002", "MakeNextStepPrompt_Click: " & Err.Description
    MsgBox "Ошибка: " & Err.Description, vbExclamation
    Err.Clear
End Sub

' ------------------------------------------------------------
' 8. Сделать промт исправления ошибок
' ------------------------------------------------------------
Public Sub MakeFixPrompt_Click()
    On Error GoTo ErrHandler
    ' Работа выполняется в Mod_Prompt (раздел 13 спецификации)
    Call Mod_Prompt.MakeFixPrompt
    Mod_Util.WriteLog "MakeFixPrompt_Click: запуск формирования промта исправления"
    Exit Sub
ErrHandler:
    Mod_Util.WriteError "P003", "MakeFixPrompt_Click: " & Err.Description
    MsgBox "Ошибка: " & Err.Description, vbExclamation
    Err.Clear
End Sub

' ------------------------------------------------------------
' 9. Показать, что нашёл Препроцессор
' ------------------------------------------------------------
' W-UI-DISCOVERY-FIX: вспомогательные функции для безопасного построения
' отчёта об обнаружении. Раньше ошибка "Invalid procedure call or argument"
' (Err.Number = 5) возникала из-за:
'   * обращения к несуществующей папке/файлу (Dir/GetAttr/FileDateTime),
'   * обращения к листу, отсутствующему в книге,
'   * работы с пустым диапазоном (LastRow <= 1).
' Все эти операции теперь выполняются только после явных проверок.

' Возвращает True, если папка существует на диске.
Private Function FolderExists(ByVal folderPath As String) As Boolean
    Dim attr As Long
    If Len(folderPath) = 0 Then Exit Function
    On Error Resume Next
    attr = GetAttr(folderPath)
    On Error GoTo 0
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    FolderExists = ((attr And vbDirectory) = vbDirectory)
End Function

' Ищет самый свежий .xlsx в папке (пропуская временные файлы ~$...).
' Возвращает полное имя файла или "" , если подходящих файлов нет.
Private Function NewestXlsxFile(ByVal folderPath As String) As String
    Dim fName As String
    Dim best As String
    Dim bestDate As Date
    Dim tmpDate As Date

    NewestXlsxFile = ""
    If Not FolderExists(folderPath) Then Exit Function
    fName = Dir(folderPath & "\*.xlsx")
    Do While Len(fName) > 0
        If Left(fName, 2) <> "~$" Then
            On Error Resume Next
            tmpDate = FileDateTime(folderPath & "\" & fName)
            If Err.Number = 0 Then
                If best = "" Then
                    best = fName
                    bestDate = tmpDate
                ElseIf tmpDate > bestDate Then
                    best = fName
                    bestDate = tmpDate
                End If
            Else
                Err.Clear
            End If
            On Error GoTo 0
        End If
        fName = Dir()
    Loop
    NewestXlsxFile = best
End Function

' Проверяет наличие листа с указанным именем в книге.
Private Function SheetExistsInBook(ByVal wb As Workbook, _
                                   ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If Err.Number <> 0 Then
        Err.Clear
        SheetExistsInBook = False
    Else
        SheetExistsInBook = (Not ws Is Nothing)
    End If
End Function

' Возвращает номер последней непустой строки на листе (по столбцу 1).
' Для пустого диапазона возвращает 0 (а не 1 по умолчанию у xlUp).
Private Function LastUsedRowOfSheet(ByVal ws As Worksheet) As Long
    Dim lr As Long
    On Error Resume Next
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    On Error GoTo 0
    If Err.Number <> 0 Then
        Err.Clear
        LastUsedRowOfSheet = 0
    ElseIf lr < 1 Then
        LastUsedRowOfSheet = 0
    Else
        ' Пустой лист: xlUp указывает на строку 1, но ячейка A1 пуста.
        If Len(Trim(CStr(ws.Cells(lr, 1).Value))) = 0 Then
            LastUsedRowOfSheet = 0
        Else
            LastUsedRowOfSheet = lr
        End If
    End If
End Function

Public Sub ShowDiscoveryReport_Click()
    Const MSG_NO_REPORT As String = _
        "Отчет об обнаружении не найден. Сначала выполните анализ схемы."

    Dim root As String
    Dim prepPath As String
    Dim reportFile As String
    Dim reportWb As Workbook
    Dim wsRep As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim report As String
    Dim lineText As String
    Dim wasOpen As Boolean

    On Error GoTo ErrHandler

    root = Mod_Util.GetProjectRoot()
    If root = "" Then
        MsgBox MSG_NO_REPORT, vbExclamation
        Exit Sub
    End If

    ' --- Проверка 1: файл отчёта действительно существует -------------
    ' Отчёт строится по результатам анализа схемы (BuildFlatStructure):
    ' книга "Плоская_таблица_*.xlsx" в папке "02_Подготовка схемы".
    prepPath = root & "\02_Подготовка схемы"
    If Not FolderExists(prepPath) Then
        MsgBox MSG_NO_REPORT, vbExclamation
        Exit Sub
    End If

    reportFile = NewestXlsxFile(prepPath)
    If reportFile = "" Then
        ' Файла результатов нет - анализ схемы ещё не выполнялся.
        MsgBox MSG_NO_REPORT, vbExclamation
        Exit Sub
    End If
    If Dir(prepPath & "\" & reportFile) = "" Then
        MsgBox MSG_NO_REPORT, vbExclamation
        Exit Sub
    End If

    ' --- Открытие книги отчёта (без дублирующего открытия) ------------
    wasOpen = False
    On Error Resume Next
    Set reportWb = Workbooks(reportFile)
    On Error GoTo ErrHandler
    If Not reportWb Is Nothing Then
        wasOpen = True
    Else
        Set reportWb = Workbooks.Open(prepPath & "\" & reportFile, _
                                      ReadOnly:=True, UpdateLinks:=0, _
                                      AddToMru:=False)
    End If
    If reportWb Is Nothing Then
        MsgBox MSG_NO_REPORT, vbExclamation
        Exit Sub
    End If

    ' --- Проверка 2: лист, к которому идёт обращение, существует ------
    If Not SheetExistsInBook(reportWb, "Плоская_таблица") Then
        If Not wasOpen Then reportWb.Close SaveChanges:=False
        MsgBox MSG_NO_REPORT, vbExclamation
        Exit Sub
    End If
    Set wsRep = reportWb.Worksheets("Плоская_таблица")

    ' --- Проверка 3: диапазон не пуст (LastRow > 1) -------------------
    lastRow = LastUsedRowOfSheet(wsRep)
    If lastRow <= 1 Then
        If Not wasOpen Then reportWb.Close SaveChanges:=False
        MsgBox MSG_NO_REPORT, vbExclamation
        Exit Sub
    End If

    ' --- Формирование текста отчёта -----------------------------------
    report = "Отчет об обнаружении (" & reportFile & "):" & vbCrLf & vbCrLf
    For r = 1 To lastRow
        lineText = Trim(CStr(wsRep.Cells(r, 1).Value))
        If Len(lineText) > 0 Then
            report = report & lineText & vbCrLf
        End If
    Next r

    ' Защита от переполнения MsgBox (обрезка очень длинного отчёта).
    If Len(report) > 3000 Then
        report = Left(report, 3000) & vbCrLf & "... (см. полный отчёт в файле)"
    End If

    If Not wasOpen Then reportWb.Close SaveChanges:=False

    MsgBox report, vbInformation
    Mod_Util.WriteLog "Показан отчёт обнаружения: " & reportFile
    Exit Sub

CleanUp:
    ' Закрываем книгу отчёта, если открывали её сами.
    On Error Resume Next
    If Not reportWb Is Nothing Then
        If Not wasOpen Then reportWb.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Exit Sub

ErrHandler:
    Mod_Util.WriteLog "Ошибка в ShowDiscoveryReport_Click: " & Err.Description
    MsgBox MSG_NO_REPORT, vbExclamation
    Err.Clear
    GoTo CleanUp
End Sub

' ------------------------------------------------------------
' 10. Открыть папку Логи
' ------------------------------------------------------------
Public Sub OpenLogsFolder_Click()
    Dim p As String
    On Error GoTo ErrHandler
    p = Mod_Util.GetLogsFolder()
    Mod_Util.OpenFolderInExplorer p
    Mod_Util.WriteLog "Открыта папка логов: " & p
    Exit Sub
ErrHandler:
    Mod_Util.WriteLog "Ошибка в OpenLogsFolder_Click: " & Err.Description
    MsgBox "Ошибка: " & Err.Description, vbExclamation
    Err.Clear
End Sub

' ------------------------------------------------------------
' 11. Обновить Readme
' ------------------------------------------------------------
Public Sub UpdateReadme_Click()
    On Error GoTo ErrHandler
    Call Mod_Docs.UpdateReadme
    Mod_Util.WriteLog "UpdateReadme_Click: документ обновлён"
    Exit Sub
ErrHandler:
    Mod_Util.WriteError "D003", "UpdateReadme_Click: ошибка обновления документа. Ошибка: " & Err.Description
    MsgBox "Ошибка D003: не удалось обновить документ.", vbCritical, "Документация"
    Err.Clear
End Sub

' ------------------------------------------------------------
' 12. Восстановить кнопки и оформление
' ------------------------------------------------------------
Public Sub RestoreUI_Click()
    On Error GoTo ErrHandler
    RestoreUI
    Mod_Util.WriteLog "UI восстановлен вручную"
    Exit Sub
ErrHandler:
    Mod_Util.WriteLog "Ошибка в RestoreUI_Click: " & Err.Description
    MsgBox "Ошибка: " & Err.Description, vbExclamation
    Err.Clear
End Sub

Private Sub EnsureSheet_Readme(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim wsAny As Worksheet

    ' Проверка: есть ли лист "Readme"
    Set ws = Nothing
    For Each wsAny In wb.Worksheets
        If StrComp(wsAny.Name, "Readme", vbTextCompare) = 0 Then
            Set ws = wsAny
            Exit For
        End If
    Next wsAny

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "Readme"
        ws.Range("A1").value = "Препроцессор структуры БД Pilot - инструкция"
        ws.Range("A2").value = "Инструкция формируется процедурой Mod_Docs.UpdateReadme."
        ws.Range("A3").value = "Нажмите кнопку ""Обновить Readme"" для пересборки."
        ws.Range("A1").Font.Bold = True
        ws.Columns("A").ColumnWidth = 100
        ws.Rows("2:3").WrapText = True
        Mod_Util.WriteLog "EnsureSheet_Readme: лист 'Readme' создан"
    Else
        Mod_Util.WriteLog "EnsureSheet_Readme: лист 'Readme' уже существует"
    End If
End Sub

'===============================================================
' Переставляет листы в порядке: Препроцессор, Настройки, Readme
'===============================================================
Private Sub ReorderSheets(ByVal wb As Workbook)
    Dim order As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim anchor As Worksheet

    order = Array("Препроцессор", "Настройки", "Readme")

    Set anchor = Nothing
    For i = LBound(order) To UBound(order)
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(order(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            If anchor Is Nothing Then
                ws.Move Before:=wb.Worksheets(1)
            Else
                ws.Move After:=anchor
            End If
            Set anchor = ws
        End If
    Next i
End Sub

