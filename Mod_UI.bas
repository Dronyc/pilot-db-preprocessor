Attribute VB_Name = "Mod_UI"
Option Explicit

' ============================================================
' Модуль: Mod_UI
' Назначение: UI-слой проекта «Препроцессор структуры БД Pilot».
'             Создание кнопок на листе «Препроцессор», форматирование
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
    Dim btn As Shape
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

    ' Проверка наличия листа «Препроцессор»
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

    ' Удаление существующих Form Controls
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            shp.Delete
        End If
    Next shp

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
        "CreateProjectFolders", _
        "CheckSourceFiles", _
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
    rowBtn = 2
    For i = LBound(captions) To UBound(captions)
        Set btn = ws.Shapes.AddFormControl(xlButtonControl, _
            ws.Cells(rowBtn, 2).Left, ws.Cells(rowBtn, 2).Top, 260, 22)
        With btn
            .Name = "btnPilot_" & CStr(i + 1)
            .OnAction = CStr(actions(i))
            .ScreenTip = CStr(tips(i))
            .TextFrame.Characters.Text = CStr(captions(i))
        End With
        rowBtn = rowBtn + 2
    Next i

    ' Форматирование всех листов книги
    For Each wsAny In ThisWorkbook.Worksheets
        On Error Resume Next
        wsAny.Unprotect Password:="pilot"
        On Error GoTo ErrHandler

        ' Первая строка — жирным
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

    ' Активируем лист «Препроцессор»
    On Error Resume Next
    ws.Activate
    On Error GoTo ErrHandler

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
Public Sub CreateProjectFolders()
    On Error GoTo ErrHandler
    Mod_Folders.CreateProjectFolders
    Mod_Util.WriteLog "Создание папок проекта выполнено"
    MsgBox "Папки проекта созданы"
    Exit Sub
ErrHandler:
    Mod_Util.WriteLog "Ошибка в CreateProjectFolders: " & Err.Description
    MsgBox "Ошибка: " & Err.Description, vbExclamation
    Err.Clear
End Sub

' ------------------------------------------------------------
' 2. Проверить исходные данные
' ------------------------------------------------------------
Public Sub CheckSourceFiles()
    Dim ok As Boolean
    On Error GoTo ErrHandler
    ok = Mod_Folders.CheckSourceFolders
    Mod_Util.WriteLog "Проверка исходных данных: " & IIf(ok, "OK", "есть проблемы")
    MsgBox "Проверка завершена. Смотрите лог."
    Exit Sub
ErrHandler:
    Mod_Util.WriteLog "Ошибка в CheckSourceFiles: " & Err.Description
    MsgBox "Ошибка: " & Err.Description, vbExclamation
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
Public Sub ShowDiscoveryReport_Click()
    Dim root As String
    Dim subName As String
    Dim report As String
    Dim cnt As Long
    Dim total As Long
    Dim folderPath As String
    Dim fileName As String

    On Error GoTo ErrHandler

    root = Mod_Util.GetStepsFolder()
    If root = "" Then
        MsgBox "Папка шагов не найдена", vbExclamation
        Exit Sub
    End If

    report = "Карта найденных шагов:" & vbCrLf & vbCrLf
    total = 0

    subName = Dir(root & "\", vbDirectory)
    Do While Len(subName) > 0
        If subName <> "." And subName <> ".." Then
            If (GetAttr(root & "\" & subName) And vbDirectory) = vbDirectory Then
                folderPath = root & "\" & subName
                cnt = 0
                fileName = Dir(folderPath & "\*.*")
                Do While Len(fileName) > 0
                    cnt = cnt + 1
                    fileName = Dir()
                Loop
                report = report & subName & " — файлов: " & CStr(cnt) & vbCrLf
                total = total + cnt
            End If
        End If
        subName = Dir()
    Loop

    report = report & vbCrLf & "Всего файлов: " & CStr(total)
    MsgBox report, vbInformation
    Mod_Util.WriteLog "Показан отчёт обнаружения"
    Exit Sub
ErrHandler:
    Mod_Util.WriteLog "Ошибка в ShowDiscoveryReport_Click: " & Err.Description
    MsgBox "Ошибка: " & Err.Description, vbExclamation
    Err.Clear
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
