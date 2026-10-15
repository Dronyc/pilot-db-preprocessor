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
Const SHEET_RESULT As String = "Сводная таблица"
Const SHEET_LOG As String = "Отчет о сборке"

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
    Dim fileDialog As fileDialog
    Dim selectedFiles As Variant
    Dim i As Long, filesCount As Long
    
    ' 1. Выбор файлов пользователем
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Выберите файлы Excel для слияния (можно несколько)"
        .Filters.Clear
        .Filters.Add "Excel файлы", "*.xlsx; *.xls; *.xlsm"
        .AllowMultiSelect = True
        If .Show = False Then Exit Sub
        selectedFiles = Array()
        ReDim selectedFiles(.SelectedItems.Count - 1)
        For i = 1 To .SelectedItems.Count
            selectedFiles(i - 1) = .SelectedItems(i)
        Next i
    End With
    filesCount = UBound(selectedFiles) + 1
    If filesCount = 0 Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' 2. Создание книги-результата
    Dim wbResult As Workbook
    Set wbResult = Workbooks.Add(xlWBATWorksheet)
    wbResult.Worksheets(1).Name = SHEET_RESULT
    wbResult.Worksheets.Add(After:=wbResult.Worksheets(1)).Name = SHEET_LOG
    
    ' 3. Заголовки
    Dim wsResult As Worksheet
    Set wsResult = wbResult.Worksheets(SHEET_RESULT)
    wsResult.Range("A1").Value = "Отображаемое имя"
    wsResult.Range("B1").Value = "Тип"
    wsResult.Range("C1").Value = "Входит в"
    wsResult.Range("D1").Value = "Состоит из"
    
    ' 4. Словари для проверки дубликатов
    Dim dictFull As Object      ' Ключ = полное совпадение всех 4 полей
    Dim dictNameType As Object  ' Ключ = Имя + Тип (для легитимных повторений)
    Dim dictNames As Object     ' Все уникальные имена (для проверки связей)
    Set dictFull = CreateObject("Scripting.Dictionary")
    Set dictNameType = CreateObject("Scripting.Dictionary")
    Set dictNames = CreateObject("Scripting.Dictionary")
    
    ' Добавляем системных родителей в словарь имён
    Dim sysParents As Variant
    sysParents = Split(SYSTEM_PARENTS, ";")
    Dim sp As Long
    For sp = LBound(sysParents) To UBound(sysParents)
        dictNames(Trim(sysParents(sp))) = True
    Next sp
    
    Dim totalRows As Long
    Dim criticalDupes As Long
    Dim legitRepeats As Long
    Dim uniqueRows As Long
    Dim outputRow As Long
    outputRow = 2
    
    ' 5. Обработка каждого файла
    Dim logText As String
    logText = "=== ОТЧЕТ О СБОРКЕ ===" & vbCrLf
    logText = logText & "Дата: " & Now() & vbCrLf
    logText = logText & "Обработано файлов: " & filesCount & vbCrLf & vbCrLf
    
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim lastRow As Long, srcRow As Long
    Dim keyFull As String, keyNameType As String
    Dim nameVal As String, typeVal As String, parentVal As String, childVal As String
    
    For i = 0 To filesCount - 1
        Set wbSrc = Workbooks.Open(selectedFiles(i), ReadOnly:=True)
        Set wsSrc = wbSrc.Worksheets(1) ' Берём первый лист каждого файла
        
        lastRow = wsSrc.Cells(wsSrc.Rows.Count, COL_NAME).End(xlUp).Row
        
        ' Определяем стартовую строку (пропуск заголовка)
        Dim startRow As Long
        startRow = 2
        If InStr(UCase(Trim(wsSrc.Cells(1, COL_NAME).Value)), "ОТОБРАЖАЕМОЕ") > 0 Then
            startRow = 2
        Else
            startRow = 1
        End If
        
        logText = logText & "Файл: " & wbSrc.Name & vbCrLf
        logText = logText & "  Строк данных: " & (lastRow - startRow + 1) & vbCrLf
        
        For srcRow = startRow To lastRow
            nameVal = Trim(CStr(wsSrc.Cells(srcRow, COL_NAME).Value))
            typeVal = Trim(CStr(wsSrc.Cells(srcRow, COL_TYPE).Value))
            parentVal = NormalizeSeparators(Trim(CStr(wsSrc.Cells(srcRow, COL_PARENT).Value)))
            childVal = NormalizeSeparators(Trim(CStr(wsSrc.Cells(srcRow, COL_CHILDREN).Value)))
            
            ' Пропускаем полностью пустые строки
            If nameVal = "" And typeVal = "" Then GoTo NextRow
            
            totalRows = totalRows + 1
            
            ' Ключи для проверки дубликатов
            keyFull = nameVal & "|||" & typeVal & "|||" & parentVal & "|||" & childVal
            keyNameType = nameVal & "|||" & typeVal
            
            ' Проверка на критический дубликат
            If dictFull.Exists(keyFull) Then
                criticalDupes = criticalDupes + 1
                logText = logText & "  [КРИТ.ДУБЛ] " & nameVal & " (" & typeVal & ")" & vbCrLf
                GoTo NextRow
            End If
            
            ' Проверка на легитимное повторение
            If dictNameType.Exists(keyNameType) Then
                legitRepeats = legitRepeats + 1
            Else
                dictNameType.Add keyNameType, 1
            End If
            
            ' Добавляем имя в словарь всех имён (для последующей проверки связей)
            If nameVal <> "" And Not dictNames.Exists(nameVal) Then
                dictNames.Add nameVal, True
            End If
            
            ' Добавляем строку в результат
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
    Next i
    
    ' ========================================================================
    ' ПРОВЕРКА ЦЕЛОСТНОСТИ СВЯЗЕЙ + ПОДСВЕТКА ПРОБЛЕМ
    ' ========================================================================
    Dim danglingRefs As Long
    Dim danglingLog As String
    danglingLog = ""
    danglingRefs = 0
    
    ' Словарь для отслеживания строк с повторениями (для подсветки)
    Dim dictRepeatRows As Object
    Set dictRepeatRows = CreateObject("Scripting.Dictionary")
    
    ' Собираем информацию о повторениях
    Dim r As Long
    For r = 2 To outputRow - 1
        Dim rowKey As String
        rowKey = Trim(CStr(wsResult.Cells(r, COL_NAME).Value)) & "|||" & Trim(CStr(wsResult.Cells(r, COL_TYPE).Value))
        If rowKey <> "|||" Then
            If Not dictRepeatRows.Exists(rowKey) Then
                dictRepeatRows.Add rowKey, r  ' Запоминаем первую строку
            Else
                ' Помечаем обе строки как повторяющиеся
                Dim firstRow As Long
                firstRow = dictRepeatRows(rowKey)
                dictRepeatRows(rowKey) = -firstRow  ' Отрицательное = есть повтор
            End If
        End If
    Next r
    
    ' Проверка висячих ссылок и подсветка
    For r = 2 To outputRow - 1
        Dim parents As String
        parents = CStr(wsResult.Cells(r, COL_PARENT).Value)
        Dim hasDangling As Boolean
        hasDangling = False
        
        If parents <> "" Then
            Dim arrParents As Variant
            arrParents = Split(parents, SEPARATOR)
            Dim p As Long
            For p = LBound(arrParents) To UBound(arrParents)
                Dim parentName As String
                parentName = Trim(arrParents(p))
                If parentName <> "" And Not dictNames.Exists(parentName) Then
                    danglingRefs = danglingRefs + 1
                    hasDangling = True
                    danglingLog = danglingLog & "  Строка " & r & ": """ & parentName & """ (родитель для """ & wsResult.Cells(r, COL_NAME).Value & """)" & vbCrLf
                End If
            Next p
        End If
        
        ' Подсветка висячей ссылки (красный)
        If hasDangling Then
            wsResult.Range("A" & r & ":D" & r).Interior.Color = COLOR_DANGLING
        End If
        
        ' Подсветка пустых полей (оранжевый)
        Dim curName As String, curType As String
        curName = Trim(CStr(wsResult.Cells(r, COL_NAME).Value))
        curType = Trim(CStr(wsResult.Cells(r, COL_TYPE).Value))
        If (curName = "" And curType <> "") Or (curName <> "" And curType = "") Then
            wsResult.Range("A" & r & ":D" & r).Interior.Color = COLOR_EMPTY
        End If
        
        ' Подсветка неизвестного типа (фиолетовый)
        If curType <> "" Then
            Dim validArr As Variant
            validArr = Split(VALID_TYPES, ";")
            Dim v As Long, typeFound As Boolean
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
    
    ' Подсветка легитимных повторений (жёлтый) — только если нет более серьёзной проблемы
    Dim keyRepeat As Variant
    For Each keyRepeat In dictRepeatRows.Keys
        Dim storedRow As Long
        storedRow = dictRepeatRows(keyRepeat)
        If storedRow < 0 Then
            ' Есть повтор — подсвечиваем обе строки (если они ещё не подсвечены красным)
            Dim absRow As Long
            absRow = Abs(storedRow)
            HighlightIfNotDangling wsResult, absRow
            ' Найти вторую строку с тем же ключом
            Dim r2 As Long
            For r2 = 2 To outputRow - 1
                Dim rk2 As String
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
    ' ИТОГОВЫЙ ОТЧЁТ — сохранение в текстовый файл (Unicode)
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
    
    ' --- Сохранение полного отчёта в .txt файл ---
    Dim logFilePath As String
    Dim logFolder As String
    
    ' Определяем папку для сохранения (папка результирующей книги)
    logFolder = wbResult.Path
    If logFolder = "" Then
        ' Если книга ещё не сохранена — используем папку "Мои документы"
        logFolder = Environ("USERPROFILE") & "\Documents"
    End If
    
    ' Формируем имя файла с датой и временем
    logFilePath = logFolder & "\Отчет_о_сборке_" & Format(Now, "yyyymmdd_hhmmss") & ".txt"
    
    ' Записываем файл в кодировке Unicode (третий параметр True)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.CreateTextFile(logFilePath, True, True) ' True, True = перезапись + Unicode
    ts.Write logText
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    ' --- В Excel записываем только краткую сводку со ссылкой на файл ---
    Dim wsLog As Worksheet
    Set wsLog = wbResult.Worksheets(SHEET_LOG)
    
    With wsLog
        .Range("A1").Value = "ПОЛНЫЙ ОТЧЁТ СОХРАНЁН В ФАЙЛ:"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        
        ' Кликабельная гиперссылка на файл отчёта
        .Hyperlinks.Add Anchor:=.Range("A2"), _
                        Address:=logFilePath, _
                        TextToDisplay:=logFilePath
        
        .Range("A4").Value = "КРАТКАЯ СВОДКА:"
        .Range("A4").Font.Bold = True
        
        .Range("A5").Value = "Дата и время сборки:"
        .Range("B5").Value = Now
        
        .Range("A6").Value = "Обработано файлов:"
        .Range("B6").Value = filesCount
        
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
    
    ' ========================================================================
    ' ФОРМАТИРОВАНИЕ (без сортировки!)
    ' ========================================================================
    FormatResultSheet wsResult, outputRow - 1
    
    ' Автоширина колонок
    wsResult.Columns("A:D").AutoFit
    wsResult.Columns("C:D").ColumnWidth = 60  ' Широкие колонки для списков
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    ' Финальное сообщение
    Dim msg As String
    msg = "Сборка завершена!" & vbCrLf & _
          "Уникальных строк: " & uniqueRows & vbCrLf & _
          "Удалено дубликатов: " & criticalDupes & vbCrLf & _
          "Легитимных повторений: " & legitRepeats & vbCrLf & _
          "Висячих ссылок: " & danglingRefs
    
    If danglingRefs > 0 Then
        msg = msg & vbCrLf & vbCrLf & "? ВНИМАНИЕ: найдены висячие ссылки (подсвечены красным)!"
        MsgBox msg, vbExclamation, "Результат (с предупреждениями)"
    Else
        MsgBox msg, vbInformation, "Результат"
    End If
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
    ' Заменяем различные варианты разделителей на единый "; "
    text = Replace(text, ";", SEPARATOR)
    text = Replace(text, ";;", SEPARATOR)
    text = Replace(text, ", ", SEPARATOR)
    text = Replace(text, " ,", SEPARATOR)
    
    ' Удаляем двойные пробелы вокруг разделителя
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    
    ' Удаляем лишние разделители в начале и конце
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
    ' Заголовки
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 220, 240)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Границы
    ws.Range("A1:D" & lastRow).Borders.LineStyle = xlContinuous
    
    ' Автофильтр
    ws.Range("A1:D1").AutoFilter
    
    ' Закрепление заголовка
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

