Attribute VB_Name = "Mod_Structure"
Option Explicit

' ============================================================================
' КОНСТАНТЫ
' ============================================================================
Const COL_NAME As Long = 1          ' Колонка "Отображаемое имя"
Const COL_TYPE As Long = 2          ' Колонка "Тип"
Const COL_PARENT As Long = 3        ' Колонка "Входит в"
Const COL_CHILDREN As Long = 4      ' Колонка "Состоит из"

Const SEPARATOR As String = "; "    ' Единый разделитель
Const SHEET_RESULT As String = "Плоская_таблица"
Const SHEET_TRACE As String = "Трассировка"
Const SHEET_LOG As String = "Журнал"

' Системные родители, которые НЕ должны считаться висячими ссылками
Const SYSTEM_PARENTS As String = "projectfolder;reference_contractors"

' Допустимые типы (для проверки на неизвестный тип)
Const VALID_TYPES As String = "Папка;ECM-документ;Документ;Папка ЦИМ;Ярлык"

' Цвета подсветки
Const COLOR_DANGLING As Long = 13158655     ' RGB(255, 200, 200) Красный — висячая ссылка
Const COLOR_REPEAT As Long = 11862015       ' RGB(255, 255, 180) Жёлтый — легитимное повторение
Const COLOR_EMPTY As Long = 9886975         ' RGB(255, 220, 150) Оранжевый — пустое поле
Const COLOR_UNKNOWN_TYPE As Long = 16763100 ' RGB(220, 200, 255) Фиолетовый — неизвестный тип

' ============================================================================
' ГЛАВНАЯ ПРОЦЕДУРА
' ============================================================================
Sub MergeStructureFiles()
    Dim schemaFolder As String
    Dim schemaFile As String
    Dim fileCount As Long
    Dim fileName As String
    Dim wbResult As Workbook
    Dim wsResult As Worksheet
    Dim wsTrace As Worksheet
    Dim wsLog As Worksheet
    Dim dictFull As Object
    Dim dictNameType As Object
    Dim dictNames As Object
    Dim sysParents As Variant
    Dim sp As Long
    Dim totalRows As Long
    Dim criticalDupes As Long
    Dim legitRepeats As Long
    Dim uniqueRows As Long
    Dim outputRow As Long
    Dim logText As String
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim lastRow As Long, srcRow As Long
    Dim keyFull As String, keyNameType As String
    Dim nameVal As String, typeVal As String, parentVal As String, childVal As String
    Dim startRow As Long
    Dim i As Long
    Dim danglingRefs As Long
    Dim danglingLog As String
    Dim dictRepeatRows As Object
    Dim r As Long
    Dim rowKey As String
    Dim firstRow As Long
    Dim parents As String
    Dim hasDangling As Boolean
    Dim arrParents As Variant
    Dim p As Long
    Dim parentName As String
    Dim curName As String, curType As String
    Dim validArr As Variant
    Dim v As Long, typeFound As Boolean
    Dim keyRepeat As Variant
    Dim storedRow As Long
    Dim absRow As Long
    Dim r2 As Long
    Dim rk2 As String
    Dim logFilePath As String
    Dim logFolder As String
    Dim fso As Object
    Dim ts As Object
    Dim msg As String
    Dim savePath As String
    Dim resultFolder As String
    Dim wbSave As Workbook
    Dim sheetNames As Variant
    Dim sn As Long
    
    On Error GoTo ErrorHandler
    
    ' 1. Получаем путь к папке 01_Схема
    schemaFolder = Mod_Util.GetProjectRoot & "\01_Исходные данные\01_Схема"
    
    ' Проверяем существование папки
    If Dir(schemaFolder, vbDirectory) = "" Then
        Mod_Util.WriteError "S003", "Папка 01_Схема не найдена: " & schemaFolder
        MsgBox "Папка 01_Схема не найдена: " & schemaFolder, vbCritical, "Ошибка"
        Exit Sub
    End If
    
    ' 2. Ищем xlsx-файлы в папке
    fileCount = 0
    schemaFile = ""
    fileName = Dir(schemaFolder & "\*.xlsx")
    Do While fileName <> ""
        fileCount = fileCount + 1
        schemaFile = schemaFolder & "\" & fileName
        fileName = Dir
    Loop
    
    If fileCount = 0 Then
        Mod_Util.WriteError "S001", "В папке 01_Схема нет xlsx-файла"
        MsgBox "В папке 01_Схема нет xlsx-файла", vbCritical, "Ошибка"
        Exit Sub
    ElseIf fileCount > 1 Then
        Mod_Util.WriteError "S002", "В папке 01_Схема больше одного xlsx. Оставьте один файл"
        MsgBox "В папке 01_Схема больше одного xlsx. Оставьте один файл", vbCritical, "Ошибка"
        Exit Sub
    End If
    
    Mod_Util.WriteLog "MergeStructureFiles: найден файл схемы: " & schemaFile
    
    ' 3. Создание книги-результата
    Set wbResult = Workbooks.Add(xlWBATWorksheet)
    wbResult.Worksheets(1).Name = SHEET_RESULT
    wbResult.Worksheets.Add(After:=wbResult.Worksheets(1)).Name = SHEET_TRACE
    wbResult.Worksheets.Add(After:=wbResult.Worksheets(2)).Name = SHEET_LOG
    
    Set wsResult = wbResult.Worksheets(SHEET_RESULT)
    Set wsTrace = wbResult.Worksheets(SHEET_TRACE)
    Set wsLog = wbResult.Worksheets(SHEET_LOG)
    
    ' 4. Заголовки
    wsResult.Range("A1").Value = "Отображаемое имя"
    wsResult.Range("B1").Value = "Тип"
    wsResult.Range("C1").Value = "Входит в"
    wsResult.Range("D1").Value = "Состоит из"
    
    ' 5. Словари для проверки дубликатов
    Set dictFull = CreateObject("Scripting.Dictionary")
    Set dictNameType = CreateObject("Scripting.Dictionary")
    Set dictNames = CreateObject("Scripting.Dictionary")
    
    ' Добавляем системных родителей в словарь имён
    sysParents = Split(SYSTEM_PARENTS, ";")
    For sp = LBound(sysParents) To UBound(sysParents)
        dictNames(Trim(sysParents(sp))) = True
    Next sp
    
    outputRow = 2
    
    ' 6. Обработка файла
    logText = "=== ОТЧЕТ О СБОРКЕ ===" & vbCrLf
    logText = logText & "Дата: " & Now() & vbCrLf
    logText = logText & "Обработан файл: " & schemaFile & vbCrLf & vbCrLf
    
    On Error GoTo FileOpenError
    Set wbSrc = Workbooks.Open(schemaFile, ReadOnly:=True)
    On Error GoTo ErrorHandler
    
    Set wsSrc = wbSrc.Worksheets(1)
    
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, COL_NAME).End(xlUp).Row
    
    startRow = 2
    If InStr(UCase(Trim(wsSrc.Cells(1, COL_NAME).Value)), "ОТОБРАЖАЕМОЕ") > 0 Then
        startRow = 2
    Else
        startRow = 1
    End If
    
    logText = logText & "  Строк данных: " & (lastRow - startRow + 1) & vbCrLf
    
    For srcRow = startRow To lastRow
        nameVal = Trim(CStr(wsSrc.Cells(srcRow, COL_NAME).Value))
        typeVal = Trim(CStr(wsSrc.Cells(srcRow, COL_TYPE).Value))
        parentVal = NormalizeSeparators(Trim(CStr(wsSrc.Cells(srcRow, COL_PARENT).Value)))
        childVal = NormalizeSeparators(Trim(CStr(wsSrc.Cells(srcRow, COL_CHILDREN).Value)))
        
        If nameVal = "" And typeVal = "" Then GoTo NextRow
        
        totalRows = totalRows + 1
        
        keyFull = nameVal & "|||" & typeVal & "|||" & parentVal & "|||" & childVal
        keyNameType = nameVal & "|||" & typeVal
        
        If dictFull.Exists(keyFull) Then
            criticalDupes = criticalDupes + 1
            logText = logText & "  [КРИТ.ДУБЛ] " & nameVal & " (" & typeVal & ")" & vbCrLf
            GoTo NextRow
        End If
        
        If dictNameType.Exists(keyNameType) Then
            legitRepeats = legitRepeats + 1
        Else
            dictNameType.Add keyNameType, 1
        End If
        
        If nameVal <> "" And Not dictNames.Exists(nameVal) Then
            dictNames.Add nameVal, True
        End If
        
        dictFull.Add keyFull, True
        wsResult.Cells(outputRow, COL_NAME).Value = nameVal
        wsResult.Cells(outputRow, COL_TYPE).Value = typeVal
        wsResult.Cells(outputRow, COL_PARENT).Value = parentVal
        wsResult.Cells(outputRow, COL_CHILDREN).Value = childVal
        outputRow = outputRow + 1
        uniqueRows = uniqueRows + 1
        
NextRow:
    Next srcRow
    
    wbSrc.Close SaveChanges:=False
    logText = logText & vbCrLf
    
    ' ========================================================================
    ' ПРОВЕРКА ЦЕЛОСТНОСТИ СВЯЗЕЙ + ПОДСВЕТКА ПРОБЛЕМ
    ' ========================================================================
    danglingLog = ""
    danglingRefs = 0
    
    Set dictRepeatRows = CreateObject("Scripting.Dictionary")
    
    For r = 2 To outputRow - 1
        rowKey = Trim(CStr(wsResult.Cells(r, COL_NAME).Value)) & "|||" & Trim(CStr(wsResult.Cells(r, COL_TYPE).Value))
        If rowKey <> "|||" Then
            If Not dictRepeatRows.Exists(rowKey) Then
                dictRepeatRows.Add rowKey, r
            Else
                firstRow = dictRepeatRows(rowKey)
                dictRepeatRows(rowKey) = -firstRow
            End If
        End If
    Next r
    
    For r = 2 To outputRow - 1
        parents = CStr(wsResult.Cells(r, COL_PARENT).Value)
        hasDangling = False
        
        If parents <> "" Then
            arrParents = Split(parents, SEPARATOR)
            For p = LBound(arrParents) To UBound(arrParents)
                parentName = Trim(arrParents(p))
                If parentName <> "" And Not dictNames.Exists(parentName) Then
                    danglingRefs = danglingRefs + 1
                    hasDangling = True
                    danglingLog = danglingLog & "  Строка " & r & ": """ & parentName & """ (родитель для """ & wsResult.Cells(r, COL_NAME).Value & """)" & vbCrLf
                End If
            Next p
        End If
        
        If hasDangling Then
            wsResult.Range("A" & r & ":D" & r).Interior.Color = COLOR_DANGLING
        End If
        
        curName = Trim(CStr(wsResult.Cells(r, COL_NAME).Value))
        curType = Trim(CStr(wsResult.Cells(r, COL_TYPE).Value))
        If (curName = "" And curType <> "") Or (curName <> "" And curType = "") Then
            wsResult.Range("A" & r & ":D" & r).Interior.Color = COLOR_EMPTY
        End If
        
        If curType <> "" Then
            validArr = Split(VALID_TYPES, ";")
            typeFound = False
            For v = LBound(validArr) To UBound(validArr)
                If Trim(validArr(v)) = curType Then
                    typeFound = True
                    Exit For
                End If
            Next v
            If Not typeFound Then
                wsResult.Range("A" & r & ":D" & r).Interior.Color = COLOR_UNKNOWN_TYPE
            End If
        End If
    Next r
    
    For Each keyRepeat In dictRepeatRows.Keys
        storedRow = dictRepeatRows(keyRepeat)
        If storedRow < 0 Then
            absRow = Abs(storedRow)
            HighlightIfNotDangling wsResult, absRow
            For r2 = 2 To outputRow - 1
                rk2 = Trim(CStr(wsResult.Cells(r2, COL_NAME).Value)) & "|||" & Trim(CStr(wsResult.Cells(r2, COL_TYPE).Value))
                If rk2 = keyRepeat And r2 <> absRow Then
                    HighlightIfNotDangling wsResult, r2
                End If
            Next r2
        End If
    Next keyRepeat
    
    logText = logText & "=== ПРОВЕРКА ЦЕЛОСТНОСТИ СВЯЗЕЙ ===" & vbCrLf
    If danglingRefs = 0 Then
        logText = logText & "Висячих ссылок не обнаружено. Все родители существуют в таблице." & vbCrLf
    Else
        logText = logText & "Найдено висячих ссылок: " & danglingRefs & vbCrLf
        logText = logText & danglingLog & vbCrLf
    End If
    logText = logText & vbCrLf
    
    ' ========================================================================
    ' ИТОГОВЫЙ ОТЧЁТ
    ' ========================================================================
    logText = logText & "=== ИТОГ ===" & vbCrLf
    logText = logText & "Всего строк до проверки: " & totalRows & vbCrLf
    logText = logText & "Критических дубликатов удалено: " & criticalDupes & vbCrLf
    logText = logText & "Легитимных повторений сохранено: " & legitRepeats & vbCrLf
    logText = logText & "Итого уникальных строк: " & uniqueRows & vbCrLf
    logText = logText & "Висячих ссылок: " & danglingRefs & vbCrLf
    logText = logText & vbCrLf
    logText = logText & "=== ЛЕГЕНДА ПОДСВЕТКИ ===" & vbCrLf
    logText = logText & "Красный фон    — висячая ссылка (родитель не найден)" & vbCrLf
    logText = logText & "Жёлтый фон     — объект встречается в нескольких родителях" & vbCrLf
    logText = logText & "Оранжевый фон  — пустое обязательное поле (имя или тип)" & vbCrLf
    logText = logText & "Фиолетовый фон — неизвестный тип (не из списка допустимых)" & vbCrLf
    
    ' Сохранение отчёта в текстовый файл
    resultFolder = Mod_Util.GetProjectRoot & "\02_Подготовка схемы"
    
    ' Проверяем/создаём папку результатов
    If Dir(resultFolder, vbDirectory) = "" Then
        On Error Resume Next
        MkDir resultFolder
        On Error GoTo ErrorHandler
    End If
    
    logFilePath = resultFolder & "\Отчет_о_сборке_" & Format(Now, "yyyymmdd_hhmmss") & ".txt"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(logFilePath, True, True)
    ts.Write logText
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    Mod_Util.WriteLog "MergeStructureFiles: отчёт сохранён в " & logFilePath
    
    ' Записываем краткую сводку на лист Журнал
    With wsLog
        .Range("A1").Value = "ПОЛНЫЙ ОТЧЁТ СОХРАНЁН В ФАЙЛ:"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        
        .Hyperlinks.Add Anchor:=.Range("A2"), _
                        Address:=logFilePath, _
                        TextToDisplay:=logFilePath
        
        .Range("A4").Value = "КРАТКАЯ СВОДКА:"
        .Range("A4").Font.Bold = True
        
        .Range("A5").Value = "Дата и время сборки:"
        .Range("B5").Value = Now
        
        .Range("A6").Value = "Обработано файлов:"
        .Range("B6").Value = 1
        
        .Range("A7").Value = "Всего строк до проверки:"
        .Range("B7").Value = totalRows
        
        .Range("A8").Value = "Критических дубликатов удалено:"
        .Range("B8").Value = criticalDupes
        
        .Range("A9").Value = "Легитимных повторений сохранено:"
        .Range("B9").Value = legitRepeats
        
        .Range("A10").Value = "Итого уникальных строк:"
        .Range("B10").Value = uniqueRows
        
        .Range("A11").Value = "Висячих ссылок:"
        .Range("B11").Value = danglingRefs
        
        .Range("A11").Font.Color = IIf(danglingRefs > 0, RGB(200, 0, 0), RGB(0, 128, 0))
        .Range("B11").Font.Color = IIf(danglingRefs > 0, RGB(200, 0, 0), RGB(0, 128, 0))
        
        .Columns("A").ColumnWidth = 35
        .Columns("B").ColumnWidth = 80
    End With
    
    ' Форматирование листа результата
    FormatResultSheet wsResult, outputRow - 1
    
    wsResult.Columns("A:D").AutoFit
    wsResult.Columns("C:D").ColumnWidth = 60
    
    ' ========================================================================
    ' СОХРАНЕНИЕ КОПИЙ В ПАПКУ 02_Подготовка схемы
    ' ========================================================================
    On Error GoTo SaveError
    
    savePath = resultFolder & "\Результат_подготовки_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    
    ' Копируем листы результата в новую книгу для сохранения
    Set wbSave = Workbooks.Add
    wbSave.Worksheets(1).Name = SHEET_RESULT
    wbSave.Worksheets.Add(After:=wbSave.Worksheets(1)).Name = SHEET_TRACE
    wbSave.Worksheets.Add(After:=wbSave.Worksheets(2)).Name = SHEET_LOG
    
    ' Копируем данные
    wsResult.Cells.Copy wbSave.Worksheets(SHEET_RESULT).Range("A1")
    wsTrace.Cells.Copy wbSave.Worksheets(SHEET_TRACE).Range("A1")
    wsLog.Cells.Copy wbSave.Worksheets(SHEET_LOG).Range("A1")
    
    wbSave.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    wbSave.Close SaveChanges:=False
    Set wbSave = Nothing
    
    Mod_Util.WriteLog "MergeStructureFiles: результат сохранён в " & savePath
    On Error GoTo ErrorHandler
    
    ' Финальное сообщение
    msg = "Сборка завершена!" & vbCrLf & _
          "Уникальных строк: " & uniqueRows & vbCrLf & _
          "Удалено дубликатов: " & criticalDupes & vbCrLf & _
          "Легитимных повторений: " & legitRepeats & vbCrLf & _
          "Висячих ссылок: " & danglingRefs & vbCrLf & vbCrLf & _
          "Результат сохранён в:" & vbCrLf & savePath
    
    If danglingRefs > 0 Then
        msg = msg & vbCrLf & vbCrLf & "ВНИМАНИЕ: найдены висячие ссылки (подсвечены красным)!"
        MsgBox msg, vbExclamation, "Результат (с предупреждениями)"
    Else
        MsgBox msg, vbInformation, "Результат"
    End If
    
    Mod_Util.WriteLog "MergeStructureFiles: узлов " & uniqueRows & ", дубликатов " & criticalDupes & ", висячих " & danglingRefs
    Exit Sub
    
FileOpenError:
    Mod_Util.WriteError "S004", "Ошибка открытия файла схемы: " & schemaFile & " — " & Err.Description
    MsgBox "Ошибка открытия файла схемы: " & schemaFile, vbCritical, "Ошибка"
    Exit Sub
    
SaveError:
    Mod_Util.WriteError "S005", "Ошибка сохранения результатов: " & Err.Description
    MsgBox "Ошибка сохранения результатов: " & Err.Description, vbCritical, "Ошибка"
    Exit Sub
    
ErrorHandler:
    Mod_Util.WriteError "S999", "Необработанная ошибка в MergeStructureFiles: " & Err.Description
    MsgBox "Ошибка: " & Err.Description, vbCritical, "Ошибка"
End Sub

' ============================================================================
' ВСПОМОГАТЕЛЬНАЯ: подсветка строки, если она ещё не красная
' ============================================================================
Private Sub HighlightIfNotDangling(ws As Worksheet, rowNum As Long)
    Dim currentColor As Long
    currentColor = ws.Cells(rowNum, COL_NAME).Interior.Color
    If currentColor <> COLOR_DANGLING Then
        ws.Range("A" & rowNum & ":D" & rowNum).Interior.Color = COLOR_REPEAT
    End If
End Sub

' ============================================================================
' НОРМАЛИЗАЦИЯ РАЗДЕЛИТЕЛЕЙ
' ============================================================================
Private Function NormalizeSeparators(ByVal text As String) As String
    text = Replace(text, ";", SEPARATOR)
    text = Replace(text, ";;", SEPARATOR)
    text = Replace(text, ", ", SEPARATOR)
    text = Replace(text, " ,", SEPARATOR)
    
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    
    text = Trim(text)
    If Left(text, Len(SEPARATOR)) = SEPARATOR Then
        text = Mid(text, Len(SEPARATOR) + 1)
    End If
    If Right(text, Len(SEPARATOR)) = SEPARATOR Then
        text = Left(text, Len(text) - Len(SEPARATOR))
    End If
    
    NormalizeSeparators = text
End Function

' ============================================================================
' ФОРМАТИРОВАНИЕ ЛИСТА РЕЗУЛЬТАТА (без сортировки!)
' ============================================================================
Private Sub FormatResultSheet(ws As Worksheet, lastRow As Long)
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 220, 240)
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("A1:D" & lastRow).Borders.LineStyle = xlContinuous
    ws.Range("A1:D1").AutoFilter
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub
