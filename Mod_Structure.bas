Option Explicit

'=====================================================================
' Подготовка цветовой структуры ЖЦ ОКС к обработке ИИ
' (или полная замена этапа извлечения: на выходе готовая плоская таблица)
'=====================================================================

Private Const SRC_SHEET As String = "Структура ЖЦ ИМ ИСП"  ' имя листа со структурой
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
    Set ws = ThisWorkbook.Worksheets(SRC_SHEET)

    '---------- 1. Легенда: цвет -> тип ----------
    Dim legCells(1 To 6) As String
    Dim legNames(1 To 6) As String
    Dim legColors(1 To 6) As Long
    legCells(1) = "B1": legCells(2) = "C1": legCells(3) = "B2"
    legCells(4) = "C2": legCells(5) = "B3": legCells(6) = "C3"
    Dim i As Long
    For i = 1 To 6
        legNames(i) = Trim(CStr(ws.Range(legCells(i)).Value))
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
        lr = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If lr > lastRow Then lastRow = lr
    Next c

    For r = FIRST_ROW To lastRow
        For c = 1 To LAST_COL
            Dim v As Variant
            v = ws.Cells(r, c).Value
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

    '---------- 4. Вывод листов ----------
    WriteFlatTable
    WriteTraceability
    WriteLog
    MsgBox "Готово. Узлов: " & NodeCount, vbInformation
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

Private Sub GetSheet(ByVal nm As String, ByRef out As Worksheet)
    On Error Resume Next
    ThisWorkbook.Worksheets(nm).Delete
    On Error GoTo 0
    Set out = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    out.Name = nm
End Sub

'---------------- Лист 1: целевая плоская таблица ----------------
Private Sub WriteFlatTable()
    Dim ws As Worksheet
    GetSheet "Плоская_таблица", ws
    ws.Range("A1:D1").Value = Array("Отображаемое имя", "Тип", "Входит в", "Состоит из")
    Dim i As Long, r As Long
    r = 2
    For i = 1 To NodeCount
        Dim parentCell As String
        If Nodes(i).Name = "Объект" Then
            parentCell = "projectfolder"
        Else
            parentCell = ParentName(Nodes(i).ParentIdx)
        End If
        ws.Cells(r, 1).Value = Nodes(i).Name
        ws.Cells(r, 2).Value = Nodes(i).NodeType
        ws.Cells(r, 3).Value = parentCell
        ws.Cells(r, 4).Value = Nodes(i).Children
        r = r + 1
    Next i
    ws.Columns("A:D").AutoFit
End Sub

'---------------- Лист 2: трассировка ----------------
Private Sub WriteTraceability()
    Dim ws As Worksheet
    GetSheet "Трассировка", ws
    ws.Range("A1:H1").Value = Array("№", "Строка", "Столбец", "Уровень", "Отображаемое имя", "Тип", "Цвет RGB", "Родитель")
    Dim i As Long
    For i = 1 To NodeCount
        ws.Cells(i + 1, 1).Value = i
        ws.Cells(i + 1, 2).Value = Nodes(i).RowIdx
        ws.Cells(i + 1, 3).Value = Nodes(i).ColIdx
        ws.Cells(i + 1, 4).Value = Nodes(i).Level
        ws.Cells(i + 1, 5).Value = Nodes(i).Name
        ws.Cells(i + 1, 6).Value = Nodes(i).NodeType
        ws.Cells(i + 1, 7).Value = HexRGB(Nodes(i).ColorCode)
        ws.Cells(i + 1, 8).Value = ParentName(Nodes(i).ParentIdx)
    Next i
    ws.Columns("A:H").AutoFit
End Sub

'---------------- Лист 3: журнал аномалий ----------------
Private Sub WriteLog()
    Dim ws As Worksheet
    GetSheet "Журнал", ws
    ws.Range("A1:D1").Value = Array("Код проблемы", "Элемент", "Родитель", "Описание")
    Dim r As Long
    r = 2
    Dim i As Long
    For i = 1 To NodeCount
        Dim nm As String, pn As String
        nm = Nodes(i).Name
        pn = ParentName(Nodes(i).ParentIdx)

        ' сирота (не корень и без родителя)
        If Nodes(i).ParentIdx = 0 And nm <> "Объект" Then
            ws.Cells(r, 1).Value = "ORPHAN": ws.Cells(r, 2).Value = nm
            ws.Cells(r, 4).Value = "Не найден родитель уровня N-1"
            r = r + 1
        End If

        ' ECM-документ не должен иметь детей
        If Nodes(i).NodeType = "ECM-документ" And Len(Nodes(i).Children) > 0 Then
            ws.Cells(r, 1).Value = "LEAF_HAS_CHILDREN": ws.Cells(r, 2).Value = nm
            ws.Cells(r, 4).Value = "ECM-документ с дочерними элементами: " & Nodes(i).Children
            r = r + 1
        End If

        ' проверки кодов
        Dim cc As String, pc As String
        cc = LeadCode(nm)
        pc = LeadCode(pn)
        If cc <> "" And pc <> "" Then
            If Left(cc, Len(pc) + 1) <> pc & "." Then
                ws.Cells(r, 1).Value = "CODE_PARENT_MISMATCH": ws.Cells(r, 2).Value = nm
                ws.Cells(r, 3).Value = pn
                ws.Cells(r, 4).Value = "Код ребенка не продолжает код родителя"
                r = r + 1
            End If
            If CodeDepth(cc) <> CodeDepth(pc) + 1 Then
                ws.Cells(r, 1).Value = "CODE_DEPTH_JUMP": ws.Cells(r, 2).Value = nm
                ws.Cells(r, 3).Value = pn
                ws.Cells(r, 4).Value = "Глубина кода ребенка не равна глубине родителя + 1"
                r = r + 1
            End If
        End If
    Next i
    ws.Columns("A:D").AutoFit
End Sub
