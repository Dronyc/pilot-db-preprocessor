Attribute VB_Name = "Mod_Structure"
Option Explicit

'=====================================================================
' Подготовка цветовой структуры ЖЦ ОКС к обработке ИИ
' (или полная замена этапа извлечения: на выходе готовая плоская таблица)
'=====================================================================

Private Const FIRST_ROW As Long = 8                        ' строка, с которой начинается структура (Объект)
Private Const LAST_COL  As Long = 30                       ' максимальный столбец сканирования

' W-PROMPT-00-SCHEMA-FIX: имя листа и префикс сохраняемого файла результата
' подготовки схемы (переименовано из "Плоская_таблица").
Private Const FLAT_SHEET_NAME As String = "Адаптированная_структура"
' Старое имя листа (обратная совместимость бэкапа листов в книге-инструменте).
Private Const FLAT_SHEET_NAME_OLD As String = "Плоская_таблица"
' Подпапка архива старых версий файла результата в 02_Подготовка схемы.
Private Const PREP_ARCHIVE_SUBFOLDER As String = "Архив"

Private Type NodeRec
    Name       As String
    NodeType   As String
    Level      As Long
    RowIdx     As Long
    ColIdx     As Long
    parentIdx  As Long
    ColorCode  As Long
    Children   As String
End Type

Private Nodes() As NodeRec
Private NodeCount As Long

' Словарь проблемных узлов: имя узла -> "E" (ошибка) / "W" (предупреждение)
' Заполняется в ProcessShortcuts и WriteLog. Используется в WriteFlatTable для подсветки.
Private problemsDict As Object

Public Sub BuildFlatStructure()
    Dim ws As Worksheet
    Dim wbSrc As Workbook          ' книга-источник схемы (из фиксированной папки)
    Dim wbOut As Workbook          ' новая книга для результатов (сохраняется в .xlsx)
    Dim srcWasOpen As Boolean      ' источник уже был открыт до запуска
    Dim path As String             ' папка 01_Схема
    Dim fileName As String         ' единственный xlsx-файл схемы
    Dim f As String                ' текущий файл при поиске через Dir
    Dim resultsPath As String      ' папка 02_Подготовка схемы (результаты)
    On Error GoTo Fail

    '---------- ИЗМЕНЕНИЕ 1: источник - фиксированная папка 01_Схема ----------
    path = Mod_Util.GetProjectRoot & "\01_Исходные данные\01_Схема"

    ' проверка существования папки
    If Dir(path, vbDirectory) = "" Then
        Err.Clear
        GoTo NoFolder
    End If

    ' сбор списка *.xlsx (без учёта временных файлов вида ~$...)
    fileName = ""
    f = Dir(path & "\*.xlsx")
    Do While f <> ""
        If Left(f, 2) <> "~$" Then
            If fileName = "" Then
                fileName = f
            Else
                Err.Clear
                GoTo ManyFiles
            End If
        End If
        f = Dir()
    Loop

    If fileName = "" Then
        Err.Clear
        GoTo NoFiles
    End If

    ' открытие файла схемы (если уже открыт - используем его)
    srcWasOpen = False
    On Error Resume Next
    Set wbSrc = Workbooks(fileName)
    On Error GoTo Fail
    If Not wbSrc Is Nothing Then
        srcWasOpen = True
    Else
        ' W-AV-REFACTOR, п.1: AddToMru:=False — файл схемы не попадает в список
        ' последних файлов Excel (снижение эвристического скора AV).
        Set wbSrc = Workbooks.Open(path & "\" & fileName, ReadOnly:=True, AddToMru:=False)
    End If
    
    ' === ПРАВКА B: скрыть окно исходной книги, если открывали сами ===
    If Not srcWasOpen Then
        On Error Resume Next
        wbSrc.Windows(1).Visible = False
        On Error GoTo Fail
    End If
    ' === КОНЕЦ ПРАВКИ B ===

    ' лист-источник структуры - так же, как раньше работал с SRC_SHEET
    On Error Resume Next
    ' работаем с первым листом книги-источника
    ' (в будущем можно расширить до обхода нескольких листов)
    Set ws = wbSrc.Worksheets(1)
    On Error GoTo Fail
    If ws Is Nothing Then
        Err.Clear
        GoTo OpenErr
    End If

    ' ИЗМЕНЕНИЕ 2: путь сохранения результатов
    resultsPath = Mod_Util.GetProjectRoot & "\02_Подготовка схемы"

    ' новая книга для результатов (не засоряем книгу инструмента листами)
    Set wbOut = Workbooks.Add

    ' ЧАСТЬ A: скрыть окно wbOut на время работы
    Application.ScreenUpdating = False
    On Error Resume Next
    wbOut.Windows(1).Visible = False
    On Error GoTo Fail

    '---------- 1. Легенда: цвет -> тип ----------
    Dim legCells(1 To 6) As String
    Dim legNames(1 To 6) As String
    Dim legColors(1 To 6) As Long
    legCells(1) = "B1": legCells(2) = "C1": legCells(3) = "B2"
    legCells(4) = "C2": legCells(5) = "B3": legCells(6) = "C3"
    Dim i As Long
    For i = 1 To 6
        legNames(i) = Trim(CStr(ws.Range(legCells(i)).value))
        legColors(i) = ws.Range(legCells(i)).Interior.Color
    Next i

    '---------- 2. Сканирование структуры ----------
    ReDim Nodes(1 To 20000)
    NodeCount = 0
    Dim lastAtLevel(1 To LAST_COL) As Long
    For i = 1 To LAST_COL: lastAtLevel(i) = 0: Next i

    Dim lastRow As Long, lr As Long, c As Long, r As Long
    lastRow = 0
    For c = 1 To LAST_COL
        lr = ws.Cells(ws.Rows.Count, c).End(xlUp).row
        If lr > lastRow Then lastRow = lr
    Next c

    For r = FIRST_ROW To lastRow
        For c = 1 To LAST_COL
            Dim v As Variant
            v = ws.Cells(r, c).value
            If Not IsEmpty(v) Then
                If Len(Trim(CStr(v))) > 0 Then
                    Dim colr As Long
                    colr = ws.Cells(r, c).Interior.Color
                    Dim tName As String
                    tName = ColorToType(colr, legColors, legNames)
                    If tName <> "" Then                      ' цвет из легенды => это узел структуры
                        NodeCount = NodeCount + 1
                        With Nodes(NodeCount)
                            .Name = Trim(CStr(v))
                            .NodeType = tName
                            .Level = c
                            .RowIdx = r
                            .ColIdx = c
                            .ColorCode = colr
                            .Children = ""
                            If c = 1 Then
                                .parentIdx = 0
                            Else
                                .parentIdx = lastAtLevel(c - 1)
                            End If
                        End With
                        ' спец-правило: корень
                        If Nodes(NodeCount).Name = "Объект" Then
                            Nodes(NodeCount).NodeType = "Папка"
                            Nodes(NodeCount).parentIdx = 0
                        End If
                        lastAtLevel(c) = NodeCount
                        Dim k As Long
                        For k = c + 1 To LAST_COL            ' сброс более глубоких уровней
                            lastAtLevel(k) = 0
                        Next k
                    End If
                End If
            End If
        Next c
    Next r

    '---------- 3. Агрегация детей ----------
    For i = 1 To NodeCount
        If Nodes(i).parentIdx > 0 Then
            Dim p As Long
            p = Nodes(i).parentIdx
            If Len(Nodes(p).Children) = 0 Then
                Nodes(p).Children = Nodes(i).Name
            Else
                Nodes(p).Children = Nodes(p).Children & "; " & Nodes(i).Name
            End If
        End If
    Next i
    
    ' Инициализация словаря проблем (для подсветки в WriteFlatTable)
    Set problemsDict = CreateObject("Scripting.Dictionary")

    ' ЧАСТЬ B: обработка ярлыков - перенос связей к целевым объектам
    ' и пометка строк ярлыков как удаляемых (ДО вывода плоской таблицы)
    Call ProcessShortcuts

    '---------- 4. Вывод листов в новую книгу wbOut ----------
    ' ВАЖНО: WriteLog вызывается ПЕРЕД WriteFlatTable, чтобы собрать все ошибки
    ' в problemsDict до того, как WriteFlatTable начнет красить строки.
    WriteLog wbOut

    ' W-CORRECTION-PACK-1, п.3.1: проверка дублей номеров/наименований
    ' (добавляет строки "[!] ..." в "Аномалии схемы", сборку НЕ блокирует)
    CheckNumberDuplicates wbOut

    BackupFlatTableSheet          ' W-CORRECTION-PACK-1, п.3.2 (до перезаписи)
    WriteFlatTable wbOut
    WriteTraceability wbOut

    ' ЧАСТЬ B: удалить пустой лист Лист1, созданный Workbooks.Add
    Dim wsEmpty As Worksheet
    On Error Resume Next
    Set wsEmpty = wbOut.Worksheets("Лист1")
    If wsEmpty Is Nothing Then Set wsEmpty = wbOut.Worksheets("Sheet1")
    On Error GoTo Fail
    If Not wsEmpty Is Nothing Then
        Application.DisplayAlerts = False
        wsEmpty.Delete
        Application.DisplayAlerts = True
        Set wsEmpty = Nothing
    End If

    ' ИЗМЕНЕНИЕ 2: результаты сохраняются в отдельный файл .xlsx
    ' в папке 02_Подготовка схемы (листам Препроцессора больше не засоряются).
    ' 1. Убедиться, что папка есть
    If Dir(resultsPath, vbDirectory) = "" Then MkDir resultsPath
    
    Dim saveName As String
    saveName = resultsPath & "\" & FLAT_SHEET_NAME & "_" & Format$(Now, "yyyy-mm-dd_hhnnss") & ".xlsx"

    ' W-PROMPT-00-SCHEMA-FIX: архивация старых версий перед сохранением новой.
    ' Все файлы "Адаптированная_структура_*.xlsx", найденные в корне
    ' 02_Подготовка схемы, перемещаются в подпапку "Архив" (создаётся при
    ' отсутствии), а старые книги "Плоская_таблица*" — туда же (обратная
    ' совместимость). В корне папки остаётся только последний сохранённый файл.
    ArchiveOldSchemaArtifacts resultsPath
    
    ' 2. Показать окно ДО SaveAs (важно — так SaveAs записывает всё корректно)
    On Error Resume Next
    wbOut.Windows(1).Visible = True
    On Error GoTo Fail
    
    ' 3. Дать Excel время на отрисовку
    DoEvents
    
    ' 4. Показать пользователю, что идёт сохранение (в строке состояния)
    Application.StatusBar = "Сохранение результатов: " & saveName
    
    ' 5. Само сохранение
    On Error Resume Next
    Err.Clear
    wbOut.SaveAs fileName:=saveName, FileFormat:=xlOpenXMLWorkbook
    Dim saveErr As Long
    saveErr = Err.Number
    On Error GoTo Fail
    
    ' 6. Дать Excel время на запись файла
    Application.Wait Now + TimeValue("00:00:02")
    DoEvents
    
    ' 7. Сбросить строку состояния
    Application.StatusBar = False
    
    ' 8. Обработка ошибки сохранения
    If saveErr <> 0 Then
        Mod_Util.WriteError "S005", "Ошибка сохранения результатов: " & Error$(saveErr)
        MsgBox "Ошибка сохранения результатов.", vbExclamation
    
        ' Закрыть книгу без сохранения
        On Error Resume Next
        wbOut.Close SaveChanges:=False
        On Error GoTo Fail
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    ' 9. Закрыть книгу (сохранить уже не надо — только что сохранили)
    wbOut.Close SaveChanges:=False
    
    ' 10. Восстановить экран
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    '---------- 5. Финал ----------
    Mod_Util.WriteLog "BuildFlatStructure: узлов " & NodeCount
    Mod_Util.WriteLog "BuildFlatStructure: результаты сохранены в " & saveName
    MsgBox "Готово. Узлов: " & NodeCount, vbInformation

    ' === ПРАВКА B: закрыть источник, если открывали его сами ===
    If Not srcWasOpen And Not wbSrc Is Nothing Then
        On Error Resume Next
        wbSrc.Windows(1).Visible = True    ' вернуть видимость перед закрытием
        On Error GoTo Fail
        wbSrc.Close SaveChanges:=False
    End If
    ' === КОНЕЦ ПРАВКИ B ===
    Exit Sub

'---------- Обработка ошибок (изменения 1 и 2) ----------
NoFolder:                                   ' папка 01_Схема не найдена
    Mod_Util.WriteError "S003", "Папка 01_Схема не найдена: " & path
    MsgBox "Папка 01_Схема не найдена: " & path, vbExclamation
    Exit Sub

NoFiles:                                    ' в папке нет xlsx-файла
    Mod_Util.WriteError "S001", "В папке 01_Схема нет xlsx-файла"
    MsgBox "В папке 01_Схема нет xlsx-файла", vbExclamation
    Exit Sub

ManyFiles:                                  ' в папке больше одного xlsx
    Mod_Util.WriteError "S002", "В папке 01_Схема больше одного xlsx"
    MsgBox "В папке 01_Схема больше одного xlsx. Оставьте один.", vbExclamation
    Exit Sub

OpenErr:                                    ' в книге нет первого листа (или он недоступен)
    Mod_Util.WriteError "S004", "BuildFlatStructure: не удалось получить первый лист из книги """ & _
        wbSrc.Name & """."
    MsgBox "Ошибка S004: не удалось получить лист из книги """ & wbSrc.Name & """.", _
           vbExclamation, "Схема"
    If Not srcWasOpen And Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Exit Sub

Fail:                                       ' прочие ошибки (например, создание папки результатов)
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    Mod_Util.WriteError "S004", "Ошибка открытия файла схемы: " & errDesc
    MsgBox "Ошибка открытия файла схемы.", vbExclamation
End Sub

'---------------- Вспомогательные функции ----------------

Private Function ColorToType(ByVal colr As Long, legColors() As Long, legNames() As String) As String
    Dim i As Long
    For i = 1 To 6
        If colr = legColors(i) Then
            ColorToType = legNames(i)
            Exit Function
        End If
    Next i
    ColorToType = ""          ' цвет не из легенды => служебная ячейка
End Function

Private Function LeadCode(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^(\d+(?:\.\d+)*)"
    If re.Test(s) Then
        LeadCode = re.Execute(s)(0).SubMatches(0)
    Else
        LeadCode = ""
    End If
End Function

Private Function CodeDepth(ByVal code As String) As Long
    If code = "" Then CodeDepth = 0: Exit Function
    CodeDepth = UBound(Split(code, ".")) + 1
End Function

Private Function HexRGB(ByVal c As Long) As String
    Dim r As Long, g As Long, b As Long
    r = c Mod 256: g = (c \ 256) Mod 256: b = (c \ 65536) Mod 256
    HexRGB = UCase(Right("00" & Hex(r), 2) & Right("00" & Hex(g), 2) & Right("00" & Hex(b), 2))
End Function

Private Function ParentName(ByVal idx As Long) As String
    If idx = 0 Then
        ParentName = ""
    Else
        ParentName = Nodes(idx).Name
    End If
End Function

'---------------- Обработка ярлыков ----------------
' В схеме узлы с типом "Ярлык" ссылаются на другие узлы (по совпадению
' поля Name), но в плоской таблице их быть не должно.
' Процедура переносит связи от ярлыков к их целевым объектам и помечает
' строки ярлыков как удаляемые (NodeType = "УДАЛЕН"), чтобы WriteFlatTable
' исключил их из итоговой таблицы.
' Шаги:
'   1. Пройти по массиву Nodes от 1 до NodeCount.
'   2. Найти узлы с NodeType = "Ярлык".
'   3. Для каждого ярлыка найти целевой узел (то же Name, но не "Ярлык");
'      если цель не найдена - пропустить (ошибка схемы, обработается в WriteLog).
'   4. Если у ярлыка есть родитель: дописать родителя ярлыка в Children
'      целевого узла и имя ярлыка в Children родителя ярлыка (без дублей).
'   5. Пометить узел ярлыка как удаляемый (NodeType = "УДАЛЕН").
Private Sub ProcessShortcuts()
    Dim i As Long, j As Long
    On Error GoTo ErrHandler

    For i = 1 To NodeCount
        If Nodes(i).NodeType = "Ярлык" Then
            '--- поиск целевого узла: Name совпадает, но это не Ярлык ---
            Dim targetIdx As Long
            targetIdx = 0
            For j = 1 To NodeCount
                If j <> i Then
                    If Nodes(j).Name = Nodes(i).Name And Nodes(j).NodeType <> "Ярлык" Then
                        targetIdx = j
                        Exit For
                    End If
                End If
            Next j
            
            ' цель НЕ найдена -> фиксируем ошибку
            If targetIdx = 0 Then
                Nodes(i).NodeType = "ОШИБКА_ЯРЛЫКА"
                ' Зафиксировать в словаре проблем
                If Not problemsDict.Exists(Nodes(i).Name) Then
                    problemsDict.Add Nodes(i).Name, "E"   ' E = ошибка
                End If
            Else
                '--- цель найдена: переносим связи ---
                Dim parentIdx As Long
                Dim sParent As String
                parentIdx = Nodes(i).parentIdx
                sParent = ParentName(parentIdx)
                
                If Len(sParent) > 0 Then
                    ' 1. Дописать родителя ярлыка в Children целевого узла
                    If InStr(1, ";" & Nodes(targetIdx).Children & ";", ";" & sParent & ";") = 0 Then
                        If Len(Nodes(targetIdx).Children) = 0 Then
                            Nodes(targetIdx).Children = sParent
                        Else
                            Nodes(targetIdx).Children = Nodes(targetIdx).Children & "; " & sParent
                        End If
                    End If
                    
                    ' 2. Дописать имя ярлыка в Children родителя ярлыка
                    If InStr(1, ";" & Nodes(parentIdx).Children & ";", ";" & Nodes(i).Name & ";") = 0 Then
                        If Len(Nodes(parentIdx).Children) = 0 Then
                            Nodes(parentIdx).Children = Nodes(i).Name
                        Else
                            Nodes(parentIdx).Children = Nodes(parentIdx).Children & "; " & Nodes(i).Name
                        End If
                    End If
                End If
                ' успешно обработанный ярлык - на удаление
                Nodes(i).NodeType = "УДАЛЕН"
            End If
        End If
    Next i
    Exit Sub

ErrHandler:
    Dim errDescPS As String
    errDescPS = Err.Description
    Err.Clear
    Mod_Util.WriteError "S006", "ProcessShortcuts: " & errDescPS
End Sub
' Возвращает лист с именем nm в книге wb: если листа нет - создаёт его
' в конце книги; если есть - очищает через Cells.Clear (без удаления,
' чтобы Excel не показывал предупреждений об удалении листов).
Private Sub GetSheet(ByVal wb As Workbook, ByVal nm As String, ByRef out As Worksheet)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = nm
    Else
        On Error GoTo 0
        ws.Cells.Clear              ' очистка вместо удаления листа
    End If
    Set out = ws
End Sub

'---------------- Вспомогательная процедура: Умная таблица ----------------
' Создаёт на листе ws Умную таблицу (ListObject) поверх диапазона rng.
' При повторном запуске старая таблица с таким же именем сначала удаляется,
' чтобы не было конфликта имён. Автофильтры (Range.AutoFilter) больше не
' используются: фильтры Умной таблицы не уезжают при прокрутке,
' а оформление задаётся через TableStyle.
Private Sub CreateSmartTable(ByVal ws As Worksheet, _
                              ByVal rng As Range, _
                              ByVal tableName As String, _
                              ByVal styleName As String)
    Dim lo As ListObject
    Dim newLo As ListObject

    On Error Resume Next

    ' Удалить старую таблицу с таким же именем (при повторном запуске)
    For Each lo In ws.ListObjects
        If lo.Name = tableName Then
            lo.Delete
            Exit For
        End If
    Next lo

    ' Создать новую таблицу
    Set newLo = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=rng, _
        XlListObjectHasHeaders:=xlYes)

    If Not newLo Is Nothing Then
        newLo.Name = tableName
        newLo.TableStyle = styleName
    End If

    On Error GoTo 0
End Sub

'---------------- Лист 1: адаптированная структура (плоская таблица) --------
Private Sub WriteFlatTable(ByVal wbOut As Workbook)
    Dim ws As Worksheet
    GetSheet wbOut, FLAT_SHEET_NAME, ws
    ws.Range("A1:D1").value = Array("Отображаемое имя", "Тип", "Входит в", "Состоит из")
    Dim i As Long, r As Long
    Dim hasError As Boolean
    Dim isShortcut As Boolean
    
    r = 2
    For i = 1 To NodeCount
        ' Пропускаем успешно обработанные ярлыки
        If Nodes(i).NodeType = "УДАЛЕН" Then GoTo NextNode
        
        hasError = False
        isShortcut = False
        
        ' Если это ярлык без цели - это ошибка
        If Nodes(i).NodeType = "ОШИБКА_ЯРЛЫКА" Then
            hasError = True
            isShortcut = True
        End If
        
        ' Проверяем словарь проблем (добавлено в WriteLog и ProcessShortcuts)
        If problemsDict.Exists(Nodes(i).Name) Then
            hasError = True
        End If
        
        Dim parentCell As String
        If Nodes(i).Name = "Объект" Then
            parentCell = "projectfolder"
        Else
            parentCell = ParentName(Nodes(i).parentIdx)
        End If
        
        ws.Cells(r, 1).value = Nodes(i).Name
        ' Возвращаем пользователю исходное имя типа, а не служебное
        ws.Cells(r, 2).value = IIf(isShortcut, "Ярлык", Nodes(i).NodeType)
        ws.Cells(r, 3).value = parentCell
        ws.Cells(r, 4).value = Nodes(i).Children
        
        ' Подсветка проблемной строки
        If hasError Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Interior.Color = RGB(255, 199, 206)
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Font.Color = RGB(156, 0, 6)
        End If
        
        r = r + 1
NextNode:
    Next i
    
    ' Создаём Умную таблицу
    Dim lastRowData As Long
    lastRowData = r - 1
    
    If lastRowData >= 2 Then
        Call CreateSmartTable(ws, ws.Range("A1:D" & lastRowData), _
                              "tblFlatStructure", "TableStyleMedium2")
    End If
    
    ws.Columns("A:D").AutoFit
End Sub
'---------------- Лист 2: трассировка ----------------
Private Sub WriteTraceability(ByVal wbOut As Workbook)
    Dim ws As Worksheet
    GetSheet wbOut, "Трассировка", ws
    ws.Range("A1:H1").value = Array("№", "Строка", "Столбец", "Уровень", "Отображаемое имя", "Тип", "Цвет RGB", "Родитель")
    Dim i As Long
    For i = 1 To NodeCount
        ws.Cells(i + 1, 1).value = i
        ws.Cells(i + 1, 2).value = Nodes(i).RowIdx
        ws.Cells(i + 1, 3).value = Nodes(i).ColIdx
        ws.Cells(i + 1, 4).value = Nodes(i).Level
        ws.Cells(i + 1, 5).value = Nodes(i).Name
        ws.Cells(i + 1, 6).value = Nodes(i).NodeType
        ws.Cells(i + 1, 7).value = HexRGB(Nodes(i).ColorCode)
        ws.Cells(i + 1, 8).value = ParentName(Nodes(i).parentIdx)
    Next i
    
    ' Создаём Умную таблицу (вместо автофильтра Range.AutoFilter)
    Dim lastRowTrace As Long
    lastRowTrace = NodeCount + 1
    
    If lastRowTrace >= 2 Then
        Call CreateSmartTable(ws, ws.Range("A1:H" & lastRowTrace), _
                              "tblTraceability", "TableStyleMedium4")
    End If
    
    ws.Columns("A:H").AutoFit
End Sub

'---------------- Лист 3: аномалии схемы ----------------
Private Sub WriteLog(ByVal wbOut As Workbook)
    Mod_Util.WriteLog "Mod_Structure: узлов " & NodeCount
    Dim ws As Worksheet
    GetSheet wbOut, "Аномалии схемы", ws
    ws.Range("A1:H1").value = Array("Код проблемы", "Строка", "Столбец", _
                                    "Уровень", "Тип", "Элемент", "Родитель", "Описание")
    Dim r As Long
    r = 2
    Dim i As Long
    For i = 1 To NodeCount
        ' Пропускаем успешно обработанные ярлыки, чтобы не засорять журнал
        If Nodes(i).NodeType = "УДАЛЕН" Then GoTo NextNodeLog
        
        Dim nm As String, pn As String
        nm = Nodes(i).Name
        pn = ParentName(Nodes(i).parentIdx)

        ' Ошибка: ярлык без цели
        If Nodes(i).NodeType = "ОШИБКА_ЯРЛЫКА" Then
            ws.Cells(r, 1).value = "SHORTCUT_NO_TARGET"
            ws.Cells(r, 2).value = Nodes(i).RowIdx
            ws.Cells(r, 3).value = Nodes(i).ColIdx
            ws.Cells(r, 4).value = Nodes(i).Level
            ws.Cells(r, 5).value = "Ярлык"
            ws.Cells(r, 6).value = nm
            ws.Cells(r, 8).value = "Не найден исходный тип (цель) для ярлыка"
            r = r + 1
            If Not problemsDict.Exists(nm) Then problemsDict.Add nm, "E"
            GoTo NextNodeLog
        End If

        ' сирота
        If Nodes(i).parentIdx = 0 And nm <> "Объект" Then
            ws.Cells(r, 1).value = "ORPHAN"
            ws.Cells(r, 2).value = Nodes(i).RowIdx
            ws.Cells(r, 3).value = Nodes(i).ColIdx
            ws.Cells(r, 4).value = Nodes(i).Level
            ws.Cells(r, 5).value = Nodes(i).NodeType
            ws.Cells(r, 6).value = nm
            ws.Cells(r, 8).value = "Не найден родитель уровня N-1"
            r = r + 1
            If Not problemsDict.Exists(nm) Then problemsDict.Add nm, "E"
        End If

        ' проверки кодов
        Dim cc As String, pc As String
        cc = LeadCode(nm)
        pc = LeadCode(pn)
        If cc <> "" And pc <> "" Then
            If Left(cc, Len(pc) + 1) <> pc & "." Then
                ws.Cells(r, 1).value = "CODE_PARENT_MISMATCH"
                ws.Cells(r, 2).value = Nodes(i).RowIdx
                ws.Cells(r, 3).value = Nodes(i).ColIdx
                ws.Cells(r, 4).value = Nodes(i).Level
                ws.Cells(r, 5).value = Nodes(i).NodeType
                ws.Cells(r, 6).value = nm
                ws.Cells(r, 7).value = pn
                ws.Cells(r, 8).value = "Код ребенка не продолжает код родителя"
                r = r + 1
                If Not problemsDict.Exists(nm) Then problemsDict.Add nm, "E"
            End If
            If CodeDepth(cc) <> CodeDepth(pc) + 1 Then
                ws.Cells(r, 1).value = "CODE_DEPTH_JUMP"
                ws.Cells(r, 2).value = Nodes(i).RowIdx
                ws.Cells(r, 3).value = Nodes(i).ColIdx
                ws.Cells(r, 4).value = Nodes(i).Level
                ws.Cells(r, 5).value = Nodes(i).NodeType
                ws.Cells(r, 6).value = nm
                ws.Cells(r, 7).value = pn
                ws.Cells(r, 8).value = "Глубина кода ребенка не равна глубине родителя + 1"
                r = r + 1
                If Not problemsDict.Exists(nm) Then problemsDict.Add nm, "E"
            End If
        End If
NextNodeLog:
    Next i
    
    Dim lastRowLog As Long
    lastRowLog = r - 1
    
    If lastRowLog >= 2 Then
        Call CreateSmartTable(ws, ws.Range("A1:H" & lastRowLog), _
                              "tblAnomalies", "TableStyleMedium6")
    End If
    
    ws.Columns("A:H").AutoFit
End Sub

'---------------- W-CORRECTION-PACK-1, п.3.1: проверка дублей ----------------
' Выполняется ПОСЛЕ сборки плоской таблицы (данные Nodes[] уже готовы) и
' ПЕРЕД записью листа "Адаптированная_структура" (ранее — "Плоская_таблица").
' Дубли выявляются по паре
' (номер, наименование): номер — LeadCode из имени узла, наименование —
' часть имени после номера. Результаты добавляются в лист "Аномалии схемы"
' в виде строк "[!] ..." (столбец "Описание"). Сборку НЕ блокируем —
' только предупреждения в отчёте.
Private Sub CheckNumberDuplicates(ByVal wbOut As Workbook)
    Dim ws As Worksheet
    Dim dictNum As Object           ' номер -> Collection(уникальных наименований)
    Dim dictName As Object          ' наименование -> Collection(уникальных номеров)
    Dim i As Long
    Dim nm As String, code As String, namePart As String
    Dim col As Collection
    Dim v As Variant
    Dim foundIt As Boolean
    Dim dupCount As Long
    Dim r As Long

    On Error GoTo ErrHandler

    Set dictNum = CreateObject("Scripting.Dictionary")
    Set dictName = CreateObject("Scripting.Dictionary")

    For i = 1 To NodeCount
        If Nodes(i).NodeType <> "УДАЛЕН" And Nodes(i).NodeType <> "ОШИБКА_ЯРЛЫКА" Then
            nm = Nodes(i).Name
            code = LeadCode(nm)
            If code <> "" Then
                namePart = Trim$(Mid$(nm, Len(code) + 1))
                If Left$(namePart, 1) = "." Then namePart = Trim$(Mid$(namePart, 2))
                If namePart <> "" Then
                    ' номер -> множество уникальных наименований
                    If Not dictNum.Exists(code) Then
                        Set col = New Collection
                        col.Add namePart
                        Set dictNum(code) = col
                    Else
                        foundIt = False
                        For Each v In dictNum(code)
                            If CStr(v) = namePart Then foundIt = True: Exit For
                        Next v
                        If Not foundIt Then dictNum(code).Add namePart
                    End If
                    ' наименование -> множество уникальных номеров
                    If Not dictName.Exists(namePart) Then
                        Set col = New Collection
                        col.Add code
                        Set dictName(namePart) = col
                    Else
                        foundIt = False
                        For Each v In dictName(namePart)
                            If CStr(v) = code Then foundIt = True: Exit For
                        Next v
                        If Not foundIt Then dictName(namePart).Add code
                    End If
                End If
            End If
        End If
    Next i

    ' Нет дублей — лист "Аномалии схемы" не трогаем
    If dictNum.Count = 0 And dictName.Count = 0 Then Exit Sub

    GetSheet wbOut, "Аномалии схемы", ws
    r = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row
    If r < 2 Then r = 2 Else r = r + 1

    ' 1) Один номер — несколько разных наименований
    Dim k As Variant
    For Each k In dictNum.Keys
        If dictNum(k).Count > 1 Then
            ws.Cells(r, 8).Value = "[!] Дублирование номера: " & k & " — " & _
                                   JoinCollection(dictNum(k), " / ")
            r = r + 1
            dupCount = dupCount + 1
        End If
    Next k

    ' 2) Одно наименование — несколько разных номеров
    For Each k In dictName.Keys
        If dictName(k).Count > 1 Then
            ws.Cells(r, 8).Value = "[!] Одинаковое наименование: " & k & " — " & _
                                   JoinCollection(dictName(k), " / ")
            r = r + 1
            dupCount = dupCount + 1
        End If
    Next k

    If dupCount > 0 Then
        Mod_Util.WriteLog "Проверка дублей: найдено аномалий - " & dupCount
    End If
    Exit Sub

ErrHandler:
    ' проверка дублей не должна ронять сборку — только предупреждение
    Mod_Util.WriteWarning "S006", "CheckNumberDuplicates: ошибка проверки дублей. Ошибка: " & Err.Description
    Err.Clear
End Sub

' Склеивает элементы Collection через разделитель.
Private Function JoinCollection(ByVal col As Collection, ByVal sep As String) As String
    Dim i As Long
    Dim res As String
    res = ""
    For i = 1 To col.Count
        If i > 1 Then res = res & sep
        res = res & CStr(col(i))
    Next i
    JoinCollection = res
End Function

'---------------- W-CORRECTION-PACK-1, п.3.2: бэкап листа структуры ---------
' W-PROMPT-00-SCHEMA-FIX: лист переименован в "Адаптированная_структура";
' перед перезаписью проверяются оба имени — новое и старое ("Плоская_таблица",
' leftover от предыдущих версий): если такой лист существует
' (например, leftover от предыдущего запуска — на новом wbOut его обычно нет),
' копируем его в лист "Адаптированная_структура_bak_<YYYYMMDD_HHMMSS>". Храним последние
' 5 бэкапов (и новые, и старые имена), более старые удаляем. О факте бэкапа пишем INFO в лог.
Private Sub BackupFlatTableSheet()
    Dim wbTool As Workbook
    Dim wsSrc As Worksheet
    Dim ws As Worksheet
    Dim bakName As String
    Dim namesArr() As String
    Dim cnt As Long
    Dim i As Long, j As Long
    Dim tmp As String

    On Error GoTo ErrHandler

    Set wbTool = ThisWorkbook

    ' существует ли старый лист "Адаптированная_структура"?
    On Error Resume Next
    Set wsSrc = wbTool.Worksheets(FLAT_SHEET_NAME)
    If wsSrc Is Nothing Then
        ' обратная совместимость: leftover со старым именем "Плоская_таблица"
        Err.Clear
        Set wsSrc = wbTool.Worksheets(FLAT_SHEET_NAME_OLD)
    End If
    On Error GoTo ErrHandler
    If wsSrc Is Nothing Then Exit Sub      ' старого листа нет — бэкап не нужен

    bakName = FLAT_SHEET_NAME & "_bak_" & Format$(Now, "yyyymmdd_hhnnss")

    Application.DisplayAlerts = False
    wsSrc.Copy After:=wbTool.Sheets(wbTool.Sheets.Count)
    ActiveSheet.Name = bakName
    Application.DisplayAlerts = True

    Mod_Util.WriteLog "Бэкап адаптированной структуры: " & bakName

    ' ротация: храним последние 5 бэкапов (новое и старое имена префиксов),
    ' старые удаляем
    ReDim namesArr(1 To wbTool.Sheets.Count)
    cnt = 0
    For Each ws In wbTool.Worksheets
        If Left$(ws.Name, Len(FLAT_SHEET_NAME & "_bak_")) = FLAT_SHEET_NAME & "_bak_" Or _
           Left$(ws.Name, Len(FLAT_SHEET_NAME_OLD & "_bak_")) = FLAT_SHEET_NAME_OLD & "_bak_" Then
            cnt = cnt + 1
            namesArr(cnt) = ws.Name
        End If
    Next ws

    If cnt > 5 Then
        ' сортировка имён по возрастанию (формат даты позволяет сортировать как строки)
        For i = 1 To cnt - 1
            For j = i + 1 To cnt
                If namesArr(j) < namesArr(i) Then
                    tmp = namesArr(i)
                    namesArr(i) = namesArr(j)
                    namesArr(j) = tmp
                End If
            Next j
        Next i
        Application.DisplayAlerts = False
        For i = 1 To cnt - 5                    ' самые старые — удалить
            wbTool.Worksheets(namesArr(i)).Delete
        Next i
        Application.DisplayAlerts = True
    End If
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Mod_Util.WriteWarning "S007", "BackupFlatTableSheet: ошибка бэкапа листа. Ошибка: " & Err.Description
    Err.Clear
End Sub

'---------------- W-PROMPT-00-SCHEMA-FIX: архивация старых версий -----------
' Перед сохранением нового файла результата в папке 02_Подготовка схемы
' проверяется наличие старых файлов "Адаптированная_структура_*.xlsx" (и
' устаревших "Плоская_таблица*.xlsx"). Если они есть — перемещаются в
' подпапку "02_Подготовка схемы\Архив" (создаётся, если отсутствует).
' В итоге в корне 02_Подготовка схемы остаётся только последний сохранённый
' файл. Ошибки перемещения (например, файл открыт в Excel) логируются и НЕ
' блокируют сохранение новой версии.
Private Sub ArchiveOldSchemaArtifacts(ByVal prepPath As String)
    Dim masks(1 To 2) As String
    Dim archPath As String
    Dim fName As String
    Dim i As Long
    Dim moved As Long

    On Error GoTo ErrHandler

    If Dir(prepPath, vbDirectory) = "" Then Exit Sub

    archPath = prepPath & "\" & PREP_ARCHIVE_SUBFOLDER
    If Dir(archPath, vbDirectory) = "" Then MkDir archPath

    masks(1) = FLAT_SHEET_NAME & "_*.xlsx"   ' Адаптированная_структура_*.xlsx
    masks(2) = FLAT_SHEET_NAME_OLD & "*.xlsx" ' Плоская_таблица*.xlsx (старые версии)

    For i = 1 To 2
        fName = Dir(prepPath & "\" & masks(i))
        Do While fName <> ""
            If Left(fName, 2) <> "~$" Then   ' пропускаем временные файлы Office
                Name prepPath & "\" & fName As archPath & "\" & fName
                moved = moved + 1
            End If
            fName = Dir()
        Loop
    Next i

    If moved > 0 Then
        Mod_Util.WriteLog "ArchiveOldSchemaArtifacts: перемещено старых версий: " & moved & _
                          " -> " & archPath
    End If
    Exit Sub

ErrHandler:
    Mod_Util.WriteWarning "S008", "ArchiveOldSchemaArtifacts: ошибка архивации '" & _
                                  fName & "'. Ошибка: " & Err.Description
    Err.Clear
End Sub



