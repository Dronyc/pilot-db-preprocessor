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

' Фиксированные параметры источника входной схемы и выходных листов
Const SRC_SHEET As String = SHEET_RESULT ' Лист-источник входной схемы (первый лист книги)
Const FIRST_ROW As Long = 2              ' Первая строка данных (после заголовка)
Const LAST_COL As Long = COL_CHILDREN    ' Последняя колонка данных ("Состоит из")

' Цвета подсветки
Const COLOR_DANGLING As Long = 13158655     ' RGB(255, 200, 200) Красный — висячая ссылка
Const COLOR_REPEAT As Long = 11862015       ' RGB(255, 255, 180) Жёлтый — легитимное повторение
Const COLOR_EMPTY As Long = 9886975         ' RGB(255, 220, 150) Оранжевый — пустое поле
Const COLOR_UNKNOWN_TYPE As Long = 16763100 ' RGB(220, 200, 255) Фиолетовый — неизвестный тип




' ============================================================================
' ГЛАВНАЯ ПРОЦЕДУРА (точка входа — вызывается из Mod_UI)
' Источник входного файла — фиксированная папка <корень>\01_Исходные данные\01_Схема\
' ============================================================================
Sub BuildFlatStructure()
    Dim schemaFolder As String   ' Папка с исходной схемой (фиксированная)
    Dim schemaFile As String     ' Найденный xlsx-файл схемы
    Dim fsoChk As Object         ' Проверка существования папки
    Dim filesCount As Long       ' Количество xlsx-файлов в папке схемы
    Dim selectedFiles As Variant ' Список выбранных файлов (теперь — фиксированный, один файл)
    Dim i As Long

    On Error GoTo ErrOpen

    ' 1. Фиксированная папка источника вместо интерактивного выбора файла
    schemaFolder = Mod_Util.GetProjectRoot & "\01_Исходные данные\01_Схема"

    ' Проверяем, что папка существует
    Set fsoChk = CreateObject("Scripting.FileSystemObject")
    If Not fsoChk.FolderExists(schemaFolder) Then
        Set fsoChk = Nothing
        Mod_Util.WriteError "S003", "Папка 01_Схема не найдена: " & schemaFolder
        MsgBox "Папка 01_Схема не найдена: " & schemaFolder, vbExclamation
        Exit Sub
    End If
    Set fsoChk = Nothing

    ' Собираем список *.xlsx в папке схемы (ожидаем ровно один файл)
    filesCount = 0
    schemaFile = ""
    schemaFile = Dir(schemaFolder & "\*.xlsx")
    Do While schemaFile <> ""
        filesCount = filesCount + 1
        schemaFile = Dir()
    Loop

    If filesCount = 0 Then
        Mod_Util.WriteError "S001", "В папке 01_Схема нет xlsx-файла"
        MsgBox "В папке 01_Схема нет xlsx-файла", vbExclamation
        Exit Sub
    End If

    If filesCount > 1 Then
        Mod_Util.WriteError "S002", "В папке 01_Схема больше одного xlsx. Оставьте один файл."
        MsgBox "В папке 01_Схема больше одного xlsx. Оставьте один файл.", vbExclamation
        Exit Sub
    End If

    ' Ровно один файл — запоминаем его имя и формируем фиксированный список источника
    schemaFile = Dir(schemaFolder & "\*.xlsx")
    ReDim selectedFiles(0 To 0)
    selectedFiles(0) = schemaFolder & "\" & schemaFile
    filesCount = 1

ErrOpen:
    If Err.Number <> 0 Then
        Err.Clear
        Mod_Util.WriteError "S004", "Ошибка открытия файла схемы: " & Err.Description
        MsgBox "Ошибка открытия файла схемы.", vbExclamation
        Exit Sub
    End If

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
        
        ' Определяем стартовую строку данных (пропуск заголовка) — фиксированная FIRST_ROW
        Dim startRow As Long
        startRow = FIRST_ROW
        
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
    
    ' ========================================================================
    ' ВЫГРУЗКА РЕЗУЛЬТАТОВ — запись плоской таблицы, трассировки и журнала
    ' ========================================================================
    WriteFlatTable wbResult
    WriteTraceability wbResult, logFilePath, filesCount, totalRows, _
                      criticalDupes, legitRepeats, uniqueRows, danglingRefs
    WriteLog logText

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
    
    ' Финальное сообщение об успехе (сначала — в журнал через Mod_Util)
    Dim NodeCount As Long
    NodeCount = uniqueRows
    Mod_Util.WriteLog "BuildFlatStructure: узлов " & NodeCount
    If danglingRefs > 0 Then
        MsgBox "Готово. Узлов: " & NodeCount & vbCrLf & _
               "? ВНИМАНИЕ: найдены висячие ссылки (подсвечены красным)!", vbExclamation
    Else
        MsgBox "Готово. Узлов: " & NodeCount, vbInformation
    End If
    Exit Sub

End Sub

' ============================================================================
' ВЫГРУЗКА ПЛОСКОЙ ТАБЛИЦЫ — сохранение книги результатов
' в фиксированную папку <корень>\02_Подготовка схемы
' ============================================================================
Private Sub WriteFlatTable(ByVal wbResult As Workbook)
    Dim resultsPath As String    ' Фиксированная папка результатов подготовки схемы
    Dim fsoDir As Object         ' Создание папки результатов при необходимости

    On Error GoTo ErrSave

    ' Путь сохранения результатов — фиксированная папка
    resultsPath = Mod_Util.GetProjectRoot & "\02_Подготовка схемы"

    ' Создаём папку результатов, если её ещё нет (копируем по образцу ранее выполнявшегося сохранения)
    Set fsoDir = CreateObject("Scripting.FileSystemObject")
    If Not fsoDir.FolderExists(resultsPath) Then fsoDir.CreateFolder (resultsPath)
    Set fsoDir = Nothing

    ' Сохраняем книгу результатов (лист Плоская_таблица) в фиксированную папку
    Application.DisplayAlerts = False
    wbResult.SaveAs resultsPath & "\Плоская_таблица.xlsx", FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    Mod_Util.WriteLog "WriteFlatTable: результаты сохранены в " & resultsPath
    Exit Sub

ErrSave:
    Err.Clear
    Application.DisplayAlerts = True
    Mod_Util.WriteError "S005", "Ошибка сохранения результатов: " & Err.Description
    MsgBox "Ошибка сохранения результатов.", vbExclamation
End Sub

' ============================================================================
' ТРАССИРОВКА ПРЕОБРАЗОВАНИЙ — краткая сводка на листе Журнал отчёта
' (структура листа и состав колонок сохранены без изменений)
' ============================================================================
Private Sub WriteTraceability(ByVal wb As Workbook, _
                              ByVal logFilePath As String, _
                              ByVal filesCount As Long, _
                              ByVal totalRows As Long, _
                              ByVal criticalDupes As Long, _
                              ByVal legitRepeats As Long, _
                              ByVal uniqueRows As Long, _
                              ByVal danglingRefs As Long)
    Dim wsLog As Worksheet
    Set wsLog = wb.Worksheets(SHEET_LOG)

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
End Sub

' ============================================================================
' ЖУРНАЛ ВЫПОЛНЕНИЯ — запись текста отчёта через единый журнал Mod_Util
' (процедура оставлена как обёртка; фактическую запись выполняет Mod_Util)
' ============================================================================
Private Sub WriteLog(ByVal logText As String)
    Mod_Util.WriteLog logText
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

