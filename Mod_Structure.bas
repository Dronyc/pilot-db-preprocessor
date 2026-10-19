Attribute VB_Name = "Mod_Structure"
Option Explicit

'=====================================================================
' Подготовка цветовой структуры ЖЦ ОКС к обработке ИИ
' (или полная замена этапа извлечения: на выходе готовая плоская таблица)
'=====================================================================

Private Const FIRST_ROW As Long = 8                        ' строка, с которой начинается структура (Объект)
Private Const LAST_COL  As Long = 30                       ' максимальный столбец сканирования

Private Type NodeRec
    Name       As String
    NodeType   As String
    Level      As Long
    RowIdx     As Long
    ColIdx     As Long
    ParentIdx  As Long
    ColorCode  As Long
    Children   As String
End Type

Private Nodes() As NodeRec
Private NodeCount As Long

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
        Set wbSrc = Workbooks.Open(path & "\" & fileName, ReadOnly:=True)
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
                                .ParentIdx = 0
                            Else
                                .ParentIdx = lastAtLevel(c - 1)
                            End If
                        End With
                        ' спец-правило: корень
                        If Nodes(NodeCount).Name = "Объект" Then
                            Nodes(NodeCount).NodeType = "Папка"
                            Nodes(NodeCount).ParentIdx = 0
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
        If Nodes(i).ParentIdx > 0 Then
            Dim p As Long
            p = Nodes(i).ParentIdx
            If Len(Nodes(p).Children) = 0 Then
                Nodes(p).Children = Nodes(i).Name
            Else
                Nodes(p).Children = Nodes(p).Children & "; " & Nodes(i).Name
            End If
        End If
    Next i

    '---------- 4. Вывод листов в новую книгу wbOut ----------
    WriteFlatTable wbOut
    WriteTraceability wbOut
    WriteLog wbOut

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
    saveName = resultsPath & "\Плоская_таблица_" & Format$(Now, "yyyy-mm-dd_hhnnss") & ".xlsx"
    
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

'---------------- Лист 1: целевая плоская таблица ----------------
Private Sub WriteFlatTable(ByVal wbOut As Workbook)
    Dim ws As Worksheet
    GetSheet wbOut, "Плоская_таблица", ws
    ws.Range("A1:D1").value = Array("Отображаемое имя", "Тип", "Входит в", "Состоит из")
    Dim i As Long, r As Long
    r = 2
    For i = 1 To NodeCount
        Dim parentCell As String
        If Nodes(i).Name = "Объект" Then
            parentCell = "projectfolder"
        Else
            parentCell = ParentName(Nodes(i).ParentIdx)
        End If
        ws.Cells(r, 1).value = Nodes(i).Name
        ws.Cells(r, 2).value = Nodes(i).NodeType
        ws.Cells(r, 3).value = parentCell
        ws.Cells(r, 4).value = Nodes(i).Children
        r = r + 1
    Next i
    
    ' Создаём Умную таблицу
    Dim loTable As ListObject
    Dim lastRowData As Long
    lastRowData = r - 1    ' r — это следующая свободная строка после цикла
    
    If lastRowData >= 2 Then    ' таблица не пустая
        ' Удалить старую таблицу, если есть (при повторном запуске)
        On Error Resume Next
        If ws.ListObjects.Count > 0 Then
            ws.ListObjects(1).Delete
        End If
        On Error GoTo Fail
    
        ' Создать новую таблицу
        Set loTable = ws.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=ws.Range("A1:D" & lastRowData), _
            XlListObjectHasHeaders:=xlYes)
        loTable.Name = "tblFlatStructure"
        loTable.TableStyle = "TableStyleMedium2"
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
        ws.Cells(i + 1, 8).value = ParentName(Nodes(i).ParentIdx)
    Next i
    
    ' Создаём Умную таблицу
    Dim loTrace As ListObject
    Dim lastRowTrace As Long
    lastRowTrace = NodeCount + 1
    
    If lastRowTrace >= 2 Then
        On Error Resume Next
        If ws.ListObjects.Count > 0 Then ws.ListObjects(1).Delete
        On Error GoTo Fail
    
        Set loTrace = ws.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=ws.Range("A1:H" & lastRowTrace), _
            XlListObjectHasHeaders:=xlYes)
        loTrace.Name = "tblTraceability"
        loTrace.TableStyle = "TableStyleMedium4"
    End If
    
    ws.Columns("A:H").AutoFit
End Sub

'---------------- Лист 3: аномалии схемы ----------------
Private Sub WriteLog(ByVal wbOut As Workbook)
    ' ИЗМЕНЕНИЕ 3: логирование - через Mod_Util (текст сообщения с итогами).
    Mod_Util.WriteLog "Mod_Structure: узлов " & NodeCount
    Dim ws As Worksheet
    GetSheet wbOut, "Аномалии схемы", ws
    ' Колонки: A - Код проблемы, B - Строка, C - Столбец, D - Уровень,
    '          E - Тип, F - Элемент, G - Родитель, H - Описание
    ws.Range("A1:H1").value = Array("Код проблемы", "Строка", "Столбец", _
                                    "Уровень", "Тип", "Элемент", "Родитель", "Описание")
    Dim r As Long
    r = 2
    Dim i As Long
    For i = 1 To NodeCount
        Dim nm As String, pn As String
        nm = Nodes(i).Name
        pn = ParentName(Nodes(i).ParentIdx)

        ' сирота (не корень и без родителя)
        If Nodes(i).ParentIdx = 0 And nm <> "Объект" Then
            ws.Cells(r, 1).value = "ORPHAN"
            ws.Cells(r, 2).value = Nodes(i).RowIdx
            ws.Cells(r, 3).value = Nodes(i).ColIdx
            ws.Cells(r, 4).value = Nodes(i).Level
            ws.Cells(r, 5).value = Nodes(i).NodeType
            ws.Cells(r, 6).value = nm
            ws.Cells(r, 8).value = "Не найден родитель уровня N-1"
            r = r + 1
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
            End If
        End If
    Next i
    
    ' Создаём Умную таблицу
    Dim loLog As ListObject
    Dim lastRowLog As Long
    lastRowLog = r - 1
    
    If lastRowLog >= 2 Then
        On Error Resume Next
        If ws.ListObjects.Count > 0 Then ws.ListObjects(1).Delete
        On Error GoTo Fail
    
        Set loLog = ws.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=ws.Range("A1:H" & lastRowLog), _
            XlListObjectHasHeaders:=xlYes)
        loLog.Name = "tblAnomalies"
        loLog.TableStyle = "TableStyleMedium6"
    End If
    
    ws.Columns("A:H").AutoFit
End Sub



