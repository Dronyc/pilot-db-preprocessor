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

' W-STRUCTURE-DUPLICATES-FIX / W-STRUCTURE-DUPLICATES-SAVE-FIX: словарь узлов,
' участвующих в дублях номеров. Ключ: "Тип|Номер", значение — True (узел
' участвует в дубле). Заполняется в BuildDupKeyDict, который вызывается из
' главной процедуры BuildFlatStructure ПЕРЕД WriteSchemaDuplicates/WriteFlatTable;
' используется в WriteFlatTable/HighlightDuplicateRows для колонки "Аномалия"
' и жёлтой подсветки.
Private dupCountDict As Object

Public Sub BuildFlatStructure()
    Dim ws As Worksheet
    Dim wbSrc As Workbook          ' книга-источник схемы (из фиксированной папки)
    Dim wbOut As Workbook          ' новая книга для результатов (сохраняется в .xlsx)
    Dim srcWasOpen As Boolean      ' источник уже был открыт до запуска
    Dim path As String             ' папка 01_Схема
    Dim fileName As String         ' единственный xlsx-файл схемы
    Dim f As String                ' текущий файл при поиске через Dir
    Dim resultsPath As String      ' папка 02_Подготовка схемы (результаты)
    ' W-FIX-AUTOFIT-WRITELOG: сохранение состояния Excel ДО открытия книг,
    ' чтобы в метке Fail и при успешном завершении восстановить его корректно
    ' (раньше восстановление шло жёсткими True/False, что затирало настройки
    ' пользователя и не спасало от утечки ресурсов при аварийном выходе).
    Dim prevScreenUpdating As Boolean   ' прежнее значение ScreenUpdating
    Dim prevDisplayAlerts As Boolean    ' прежнее значение DisplayAlerts
    Dim prevStatusBar As Variant        ' прежнее значение StatusBar (Variant: может быть строкой или False)
    On Error GoTo Fail

    ' Запомнить текущее состояние приложения до каких-либо изменений
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    prevStatusBar = Application.StatusBar

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

    ' W-STRUCTURE-DUPLICATES-SAVE-FIX: инициализация словаря дублей ДО поиска.
    ' Заполняется в BuildDupKeyDict (вызывается из главной процедуры ПЕРЕД
    ' WriteSchemaDuplicates/WriteFlatTable), используется в WriteFlatTable
    ' (колонка "Аномалия") и HighlightDuplicateRows (жёлтая подсветка).
    Set dupCountDict = CreateObject("Scripting.Dictionary")

    ' W-DEBUG-ANOMALIES-EMPTY: отладочное логирование цепочки подготовки листа
    ' "Аномалии схемы" (DebugLog -> Mod_Util.DebugLog: Immediate + файл-журнал,
    ' на лист ничего не пишется). Плоская структура сформирована (массив Nodes).
    Mod_Util.DebugLog "BuildFlatStructure: сформировано строк = " & NodeCount

    ' ЧАСТЬ B: обработка ярлыков - перенос связей к целевым объектам
    ' и пометка строк ярлыков как удаляемых (ДО вывода плоской таблицы)
    Call ProcessShortcuts

    '---------- 4. Вывод листов в новую книгу wbOut ----------
    ' W-STRUCTURE-DUPLICATES-SAVE-FIX: исправлен порядок вызовов. Ранее лист
    ' "Аномалии схемы" мог быть очищен (GetSheet -> Cells.Clear) или перекрыт
    ' ListObject ПОСЛЕ записи строк дублей — на листе оставались только
    ' заголовки. W-STRUCTURE-LOG-FIX: усилены гарантии — PrepareAnomaliesSheet и
    ' WriteSchemaDuplicates выполняются ДО создания ListObject, а сам ListObject
    ' (CreateSmartTable) создаётся СТРОГО ОДИН РАЗ в конце подготовки журнала,
    ' когда все данные уже лежат в ячейках; WriteLog пересоздание таблицы НЕ
    ' выполняет (только Resize). Новый гарантированный порядок:
    '   1) PrepareAnomaliesSheet   — лист создан/очищен, записаны заголовки;
    '      очистка листа больше НИКОГДА не выполняется после записи данных;
    '   2) BuildDupKeyDict         — поиск дублей ("Тип|Номер"), заполнение dupCountDict;
    '   3) WriteSchemaDuplicates   — запись строк дублей в свободные строки листа;
    '   4) CreateSmartTable        — ListObject tblAnomalies создается СТРОГО
    '      ПОСЛЕ записанных данных (по фактическому диапазону A:H) и ТОЛЬКО
    '      ОДИН РАЗ — повторных вызовов для tblAnomalies в потоке сборки нет;
    '   5) WriteLog                — остальные аномалии (ярлыки/сироты/коды),
    '      дозаписываются ниже строк дублей в переданный лист wsAnom (ByRef);
    '      существующая таблица расширяется через Resize (без пересоздания);
    '   6) BackupFlatTableSheet + WriteFlatTable — основная таблица;
    '   7) HighlightDuplicateRows  — жёлтая подсветка и метка "[!] Дубль"
    '      ПОСЛЕ создания Умной таблицы tblFlatStructure;
    '   8) WriteTraceability       — трассировка.
    Dim wsAnom As Worksheet           ' лист "Аномалии схемы"
    Dim dictDupMain As Object         ' "Тип|Номер" -> Collection индексов узлов
    Dim lastRowDupMain As Long        ' последняя строка записанных дублей
    Dim anomaliesRange As Range       ' фактический диапазон листа A:H

    ' 1) Подготовка листа "Аномалии схемы": создание/очистка + заголовки.
    '    Дальнейшие шаги только ДОПИСЫВАЮТ строки — повторной очистки нет.
    PrepareAnomaliesSheet wbOut, wsAnom

    ' 2) Поиск дублей внутри одного типа (ключ "Тип|Номер") + dupCountDict.
    Set dictDupMain = BuildDupKeyDict

    ' 3) Запись строк дублей (r инкрементируется, каждая пара — в свою строку).
    '    W-DEBUG-ANOMALIES-EMPTY: логирование вызова/завершения — позволяет
    '    выявить этап, на котором теряются данные о дублях.
    lastRowDupMain = 0
    If Not dictDupMain Is Nothing Then
        Mod_Util.DebugLog "BuildFlatStructure: вызов WriteSchemaDuplicates"
        lastRowDupMain = WriteSchemaDuplicates(wsAnom, 2, dictDupMain)
        Mod_Util.DebugLog "BuildFlatStructure: WriteSchemaDuplicates завершён"
    Else
        Mod_Util.DebugLog "BuildFlatStructure: WriteSchemaDuplicates не вызван (дублей нет)"
    End If

    ' 4) ListObject tblAnomalies — ТОЛЬКО после записанных данных.
    '    Если дублей нет, создаём таблицу по заголовкам и расширим её в WriteLog.
    If lastRowDupMain >= 2 Then
        Set anomaliesRange = wsAnom.Range("A1:H" & lastRowDupMain)
    Else
        Set anomaliesRange = wsAnom.Range("A1:H2")
    End If
    Call CreateSmartTable(wsAnom, anomaliesRange, "tblAnomalies", "TableStyleMedium6")

    ' 5) Журнал остальных аномалий (ярлык без цели / сирота / коды):
    '    проблемы собираются в problemsDict ДО WriteFlatTable (подсветка ошибок),
    '    строки дозаписываются под дубли, существующая таблица tblAnomalies
    '    РАСШИРЯЕТСЯ через Resize до A1:H(lastRowLog) — пересоздание ListObject
    '    внутри WriteLog запрещено (W-STRUCTURE-LOG-FIX). Лист передаётся явно
    '    (ByRef wsAnom), повторный поиск через GetSheetNoClear не выполняется.
    WriteLog wbOut, wsAnom
    Mod_Util.DebugLog "BuildFlatStructure: WriteLog завершён, переход к BackupFlatTableSheet"

    ' 6) Основная таблица структуры.
    BackupFlatTableSheet          ' W-CORRECTION-PACK-1, п.3.2 (до перезаписи)
    Mod_Util.DebugLog "BuildFlatStructure: BackupFlatTableSheet завершён"
    WriteFlatTable wbOut
    Mod_Util.DebugLog "BuildFlatStructure: WriteFlatTable завершён"

    ' 7) W-STRUCTURE-DUPLICATES-FIX, п.3: жёлтая подсветка строк-дублей и метка
    '    "[!] Дубль" в колонке "Аномалия" — ПОСЛЕ создания Умной таблицы
    '    tblFlatStructure, чтобы стили таблицы не перекрывали заливку.
    '    dupCountDict передаётся из BuildDupKeyDict (шаг 2).
    HighlightDuplicateRows wbOut
    Mod_Util.DebugLog "BuildFlatStructure: HighlightDuplicateRows завершён"

    ' 8) Лист трассировки.
    WriteTraceability wbOut
    Mod_Util.DebugLog "BuildFlatStructure: WriteTraceability завершён"

    ' W-CORRECTION-PACK-1, п.3.1: резервная дозапись дублей в "Аномалии схемы"
    ' (обычно no-op: журнал уже заполнен на шагах 2-3). Сборку НЕ блокирует.
    CheckNumberDuplicates wbOut
    Mod_Util.DebugLog "BuildFlatStructure: CheckNumberDuplicates завершён"

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
        Set wbOut = Nothing
        ' W-FIX-AUTOFIT-WRITELOG: закрыть источник, если открывали его сами,
        ' чтобы при аварийном выходе не оставалась лишняя открытая книга
        If Not srcWasOpen And Not wbSrc Is Nothing Then
            wbSrc.Windows(1).Visible = True
            wbSrc.Close SaveChanges:=False
            Set wbSrc = Nothing
        End If
        On Error GoTo 0
        ' W-FIX-AUTOFIT-WRITELOG: восстановить состояние из prev-переменных,
        ' а не жёсткими True (раньше это затирало настройки пользователя)
        Application.ScreenUpdating = prevScreenUpdating
        Application.DisplayAlerts = prevDisplayAlerts
        Application.StatusBar = prevStatusBar
        Exit Sub
    End If
    
    ' 9. Закрыть книгу (сохранить уже не надо — только что сохранили)
    wbOut.Close SaveChanges:=False
    Set wbOut = Nothing
    
    ' 10. Восстановить экран
    ' W-FIX-AUTOFIT-WRITELOG: восстановление прежнего состояния приложения
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    Application.StatusBar = prevStatusBar

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
        Set wbSrc = Nothing
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
    ' W-FIX-AUTOFIT-WRITELOG: при открытии источника на запись закрываем его
    ' без сохранения; если книга была открыта пользователем — только вернуть
    ' видимость окна. Ошибки при уборке не должны ронять обработчик.
    On Error Resume Next
    If Not srcWasOpen And Not wbSrc Is Nothing Then
        wbSrc.Windows(1).Visible = True
        wbSrc.Close SaveChanges:=False
        Set wbSrc = Nothing
    ElseIf Not wbSrc Is Nothing Then
        wbSrc.Windows(1).Visible = True
    End If
    ' На этом пути новые книги ещё не создавались, но на всякий случай
    ' подстрахуемся и восстановим прежнее состояние приложения.
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    Set wbOut = Nothing
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    Application.StatusBar = prevStatusBar
    On Error GoTo 0
    Exit Sub

Fail:                                       ' прочие ошибки (например, создание папки результатов)
    ' W-FIX-AUTOFIT-WRITELOG: раньше сюда попадали без освобождения ресурсов —
    ' открытые нами книги (wbOut, wbSrc) оставались висящими в памяти, а
    ' ScreenUpdating/DisplayAlerts — изменёнными, из-за чего при закрытии файла
    ' Excel предлагал сохранять лишние книги. Теперь Fail корректно убирает всё.
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    ' Восстановить прежнее состояние приложения (через On Error Resume Next,
    ' чтобы сбой восстановления не прерывал обработчик ошибок)
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    Application.StatusBar = prevStatusBar
    ' Закрыть новую книгу результатов без сохранения, если она была создана
    If Not wbOut Is Nothing Then
        wbOut.Close SaveChanges:=False
        Set wbOut = Nothing
    End If
    ' Закрыть источник без сохранения, ТОЛЬКО если открывали его сами;
    ' если книга была открыта пользователем — не закрывать, лишь вернуть
    ' видимость её окна (мы скрывали его на время сборки).
    If Not srcWasOpen And Not wbSrc Is Nothing Then
        wbSrc.Windows(1).Visible = True
        wbSrc.Close SaveChanges:=False
        Set wbSrc = Nothing
    ElseIf Not wbSrc Is Nothing Then
        wbSrc.Windows(1).Visible = True
    End If
    On Error GoTo 0
    ' W-FIX-ANOMALIES-EMPTY: раньше сюда попадали из-за ошибки отладочного
    ' логирования в WriteLog, но текст "Ошибка открытия файла схемы" маскировал
    ' реальную причину (ошибка выполнения, не связанная с открытием схемы).
    ' Код S004 сохранён, текст сделан нейтральным; префикс "BuildFlatStructure: "
    ' и уточнение о возможной причине добавлены в сообщение.
    Mod_Util.WriteError "S004", "BuildFlatStructure: ошибка выполнения (возможно, не связана с открытием схемы): " & errDesc
    MsgBox "Ошибка выполнения BuildFlatStructure: " & errDesc, vbExclamation
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
    Mod_Util.DebugLog "GetSheet: старт, имя = " & nm
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
    Mod_Util.DebugLog "GetSheet: лист получен/создан"
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

    Mod_Util.DebugLog "CreateSmartTable: старт, имя = " & tableName & ", адрес = " & rng.Address

    On Error Resume Next

    ' Удалить старую таблицу с таким же именем (при повторном запуске)
    For Each lo In ws.ListObjects
        If lo.Name = tableName Then
            lo.Delete
            Exit For
        End If
    Next lo
    Mod_Util.DebugLog "CreateSmartTable: старая таблица удалена (если была)"

    ' Создать новую таблицу
    Set newLo = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=rng, _
        XlListObjectHasHeaders:=xlYes)
    Mod_Util.DebugLog "CreateSmartTable: таблица создана"

    If Not newLo Is Nothing Then
        newLo.Name = tableName
        newLo.TableStyle = styleName
    End If
    Mod_Util.DebugLog "CreateSmartTable: имя и стиль назначены"

    On Error GoTo 0
End Sub

'---------------- Лист 1: адаптированная структура (плоская таблица) --------
' W-STRUCTURE-DUPLICATES-FIX, п.3: добавлена колонка "Аномалия" (E).
' Если строка участвует в дубле (dupCountDict, ключ "Тип|Номер"),
' в колонку ставится "[!] Дубль" и строка красится в жёлтый цвет.
Private Sub WriteFlatTable(ByVal wbOut As Workbook)
    Dim ws As Worksheet
    Mod_Util.DebugLog "WriteFlatTable: старт"
    GetSheet wbOut, FLAT_SHEET_NAME, ws
    Mod_Util.DebugLog "WriteFlatTable: GetSheet завершён"
    ws.Range("A1:E1").value = Array("Отображаемое имя", "Тип", "Входит в", "Состоит из", "Аномалия")
    Dim i As Long, r As Long
    Dim hasError As Boolean
    Dim isShortcut As Boolean
    Dim isDup As Boolean
    Dim dupKey As String
    
    r = 2
    For i = 1 To NodeCount
        ' Пропускаем успешно обработанные ярлыки
        If Nodes(i).NodeType = "УДАЛЕН" Then GoTo NextNode
        
        hasError = False
        isShortcut = False
        isDup = False
        
        ' Если это ярлык без цели - это ошибка
        If Nodes(i).NodeType = "ОШИБКА_ЯРЛЫКА" Then
            hasError = True
            isShortcut = True
        End If
        
        ' Проверяем словарь проблем (добавлено в WriteLog и ProcessShortcuts)
        If problemsDict.Exists(Nodes(i).Name) Then
            hasError = True
        End If
        
        ' W-STRUCTURE-DUPLICATES-FIX: проверка участия строки в дубле
        ' (ключ словаря — "Тип|Номер"; заполнен в BuildDupKeyDict, который
        ' вызывается из главной процедуры ДО WriteFlatTable).
        ' Пустой номер (LeadCode = "") в расчёт не берётся.
        If Not dupCountDict Is Nothing Then
            dupKey = Nodes(i).NodeType & "|" & LeadCode(Nodes(i).Name)
            If Right$(dupKey, 1) <> "|" Then
                If dupCountDict.Exists(dupKey) Then isDup = True
            End If
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
        ' Колонка "Аномалия": пометка дублей
        If isDup Then ws.Cells(r, 5).value = "[!] Дубль"
        
        ' Подсветка проблемной строки
        If hasError Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Interior.Color = RGB(255, 199, 206)
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Font.Color = RGB(156, 0, 6)
        ElseIf isDup Then
            ' W-STRUCTURE-DUPLICATES-FIX: жёлтая подсветка строк-дублей
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Interior.Color = RGB(255, 255, 0)
        End If
        
        r = r + 1
NextNode:
    Next i
    Mod_Util.DebugLog "WriteFlatTable: данные записаны, строк = " & (r - 1)
    
    ' Создаём Умную таблицу
    Dim lastRowData As Long
    lastRowData = r - 1
    
    If lastRowData >= 2 Then
        Mod_Util.DebugLog "WriteFlatTable: вызов CreateSmartTable"
        Call CreateSmartTable(ws, ws.Range("A1:E" & lastRowData), _
                              "tblFlatStructure", "TableStyleMedium2")
        Mod_Util.DebugLog "WriteFlatTable: CreateSmartTable завершён"
    End If
    
    ' W-FIX-AUTOFIT-WRITELOG: AutoFit через SafeAutoFit — на скрытом/неотрисованном
    ' листе прямой вызов мог падать с "Object doesn't support this property or
    ' method" и ронять всю сборку через Fail главной процедуры.
    SafeAutoFit ws, "A:E"
End Sub
'---------------- Лист 2: трассировка ----------------
Private Sub WriteTraceability(ByVal wbOut As Workbook)
    Dim ws As Worksheet
    Mod_Util.DebugLog "WriteTraceability: старт"
    GetSheet wbOut, "Трассировка", ws
    Mod_Util.DebugLog "WriteTraceability: GetSheet завершён"
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
    Mod_Util.DebugLog "WriteTraceability: данные записаны"
    
    ' Создаём Умную таблицу (вместо автофильтра Range.AutoFilter)
    Dim lastRowTrace As Long
    lastRowTrace = NodeCount + 1
    
    If lastRowTrace >= 2 Then
        Mod_Util.DebugLog "WriteTraceability: вызов CreateSmartTable"
        Call CreateSmartTable(ws, ws.Range("A1:H" & lastRowTrace), _
                              "tblTraceability", "TableStyleMedium4")
        Mod_Util.DebugLog "WriteTraceability: CreateSmartTable завершён"
    End If
    
    ' W-FIX-AUTOFIT-WRITELOG: AutoFit через SafeAutoFit (см. комментарий в WriteFlatTable)
    SafeAutoFit ws, "A:H"
End Sub

'---------------- Вспомогательная процедура: безопасный AutoFit -------------
' W-FIX-AUTOFIT-WRITELOG: AutoFit на скрытом или ещё не отрисованном листе
' (wbOut.Windows(1).Visible = False на время сборки) в ряде конфигураций Excel
' завершался ошибкой "Object doesn't support this property or method" (Err 438),
' что прерывало WriteLog/WriteFlatTable/WriteTraceability и выводило сборку в
' метку Fail главной процедуры. SafeAutoFit перехватывает такую ошибку локально:
' ширина колонок — косметика, аварийное завершение сборки из-за неё недопустимо.
Private Sub SafeAutoFit(ByVal ws As Worksheet, ByVal cols As String)
    On Error Resume Next
    ws.Columns(cols).AutoFit
    If Err.Number <> 0 Then
        Mod_Util.DebugLog "SafeAutoFit: ошибка AutoFit для колонок " & cols & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

'---------------- Лист 3: аномалии схемы ----------------
' W-STRUCTURE-DUPLICATES-SAVE-FIX: лист "Аномалии схемы" готовится ОДИН РАЗ в
' начале (PrepareAnomaliesSheet): создание/очистка + заголовки. После этого
' данные только ДОПИСЫВАЮТСТЯ (WriteSchemaDuplicates, затем WriteLog), повторной
' очистки листа нет — записи больше не теряются. Умная таблица (ListObject
' "tblAnomalies") создаётся строго ПОСЛЕ записанных строк дублей (в главной
' процедуре BuildFlatStructure) и расширяется в конце WriteLog.
' W-STRUCTURE-LOG-FIX: WriteLog НЕ управляет структурой таблиц — внутри неё
' запрещено вызывать CreateSmartTable (пересоздание ListObject ломало заголовки,
' оставляя дефолтные "Столбец1...") и PrepareAnomaliesSheet (Cells.Clear стирал
' уже записанные данные). Лист передаётся в WriteLog явно (ByRef wsAnom), его
' повторный поиск через GetSheetNoClear исключён; существующая tblAnomalies
' только РАСШИРЯЕТСЯ методом Resize до фактических данных.
' Колонки журнала: "Код проблемы", "Строка", "Столбец", "Уровень", "Тип",
' "Элемент", "Родитель", "Описание". Строки дублей ("DUP_IN_TYPE") несут те же
' смысловые поля, что и колонки "Тип", "Номер", "Наименование 1",
' "Наименование 2", "Тип аномалии": Тип -> колонка E, Номер -> в тексте
' описания, Наименование 1 | Наименование 2 -> колонка F, "[!] Дубль" -> H.

' Подготовка листа "Аномалии схемы": очистка/создание + заголовки.
' Вызывается ОДИН раз ДО любой записки данных на лист (в BuildFlatStructure).
Private Sub PrepareAnomaliesSheet(ByVal wbOut As Workbook, ByRef out As Worksheet)
    GetSheet wbOut, "Аномалии схемы", out
    out.Range("A1:H1").Value = Array("Код проблемы", "Строка", "Столбец", _
                                     "Уровень", "Тип", "Элемент", "Родитель", "Описание")
End Sub

' W-STRUCTURE-LOG-FIX: сигнатура изменена — лист "Аномалии схемы" передаётся
' явно по ссылке (ByRef ws), чтобы исключить повторный поиск через
' GetSheetNoClear и любые попытки переподготовки листа внутри процедуры.
Private Sub WriteLog(ByVal wbOut As Workbook, ByRef ws As Worksheet)
    Mod_Util.WriteLog "Mod_Structure: узлов " & NodeCount

    ' W-DEBUG-ANOMALIES-EMPTY: отладочное логирование записи на лист
    ' "Аномалии схемы" (DebugLog -> Immediate + файл-журнал, на лист не пишет).
    ' "переданных строк" — число строк плоской структуры (массив Nodes),
    ' передаваемых в этот этап обработки.
    Mod_Util.DebugLog "WriteLog(лист): старт, переданных строк = " & NodeCount

    ' W-STRUCTURE-LOG-FIX: НЕ получаем лист через GetSheetNoClear и НЕ вызываем
    ' PrepareAnomaliesSheet — иначе Cells.Clear стирал бы уже записанные строки
    ' дублей. Лист подготовлен и заполнен дублями в BuildFlatStructure; здесь
    ' только дозаписываем остальные аномалии в переданный лист БЕЗ очистки.
    ' ws гарантированно не Nothing: вызывающая процедура получила его из
    ' PrepareAnomaliesSheet до создания ListObject.
    If ws Is Nothing Then
        Mod_Util.DebugLog "WriteLog(лист): нет данных для записи (лист не передан)"
        Exit Sub
    End If
    If NodeCount <= 0 Then
        Mod_Util.DebugLog "WriteLog(лист): нет данных для записи (плоская структура пуста)"
    End If

    Dim r As Long
    ' первая свободная строка: после заголовков (2) и после строк дублей
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
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

    '---------- W-STRUCTURE-LOG-FIX ----------
    ' Дубли ("Тип|Номер") уже найдены (BuildDupKeyDict) и записаны
    ' (WriteSchemaDuplicates) в главной процедуре ДО этого вызова — здесь их
    ' повторный поиск/запись НЕ выполняются. Вызов CreateSmartTable ИЗ ИСПОЛНЕНИЯ
    ' WriteLog ИСКЛЮЧЁН: он удалял и пересоздавал ListObject tblAnomalies, из-за
    ' чего терялись заголовки (Excel подставлял дефолтные "Столбец1...").
    ' Существующая таблица (создана ОДИН раз в BuildFlatStructure после записи
    ' дублей) только РАСШИРЯЕТСЯ методом Resize до фактических данных журнала.
    Dim lastRowLog As Long
    lastRowLog = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' W-DEBUG-ANOMALIES-EMPTY: логирование состояния ListObject tblAnomalies
    ' ДО Resize. Если данных на листе нет (последняя непустая строка —
    ' заголовки, т.е. < 2), Resize не выполняется и это фиксируется в журнале.
    Dim loAnomDbg As ListObject
    Dim loRowsBefore As Long
    Dim hdrAddrDbg As String      ' W-FIX-ANOMALIES-EMPTY: адрес HeaderRange для журнала
    Set loAnomDbg = Nothing
    On Error Resume Next
    Set loAnomDbg = ws.ListObjects("tblAnomalies")
    On Error GoTo 0
    ' W-FIX-ANOMALIES-EMPTY (блокер B): весь блок отладочного логирования
    ' состояния tblAnomalies (до Resize и после Resize) обёрнут в локальную
    ' обработку ошибок — ошибка в отладочном логировании НЕ должна прерывать
    ' WriteLog и приводить к выходу в Fail главной процедуры BuildFlatStructure.
    On Error Resume Next
    If loAnomDbg Is Nothing Then
        Mod_Util.DebugLog "WriteLog(лист): нет данных для записи или таблица отсутствует: " & _
                          "ListObject ""tblAnomalies"" не найден, Resize не выполняется"
        If Err.Number <> 0 Then
            Mod_Util.DebugLog "WriteLog(лист): ошибка отладочного логирования: " & Err.Description
            Err.Clear
        End If
    Else
        loRowsBefore = loAnomDbg.ListRows.Count
        If Err.Number <> 0 Then
            Mod_Util.DebugLog "WriteLog(лист): ошибка отладочного логирования: " & Err.Description
            Err.Clear
        End If
        Mod_Util.DebugLog "WriteLog(лист): до Resize, ListObject строк = " & loRowsBefore
        If Err.Number <> 0 Then
            Mod_Util.DebugLog "WriteLog(лист): ошибка отладочного логирования: " & Err.Description
            Err.Clear
        End If
        ' Проверка сохранности заголовков перед Resize (Resize НЕ должен их
        ' затирать): адрес и размер шапки таблицы + фактические тексты A1:H1.
        ' W-FIX-ANOMALIES-EMPTY (блокер B): Address(False, False) заменён на
        ' Address (без аргументов) — вызов с аргументами в ряде конфигураций
        ' завершался ошибкой и ронял сборку через Fail главной процедуры.
        hdrAddrDbg = loAnomDbg.HeaderRange.Address
        If Err.Number <> 0 Then
            Mod_Util.DebugLog "WriteLog(лист): ошибка отладочного логирования: " & Err.Description
            Err.Clear
        End If
        Mod_Util.DebugLog "WriteLog(лист): до Resize, HeaderRange = " & _
                          hdrAddrDbg & _
                          ", строк заголовка = " & loAnomDbg.HeaderRange.Rows.Count & _
                          ", столбцов = " & loAnomDbg.HeaderRange.Columns.Count
        If Err.Number <> 0 Then
            Mod_Util.DebugLog "WriteLog(лист): ошибка отладочного логирования: " & Err.Description
            Err.Clear
        End If
        Mod_Util.DebugLog "WriteLog(лист): до Resize, заголовки A1:H1 = """ & _
                          CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("B1").Value) & _
                          "|" & CStr(ws.Range("C1").Value) & "|" & CStr(ws.Range("D1").Value) & _
                          "|" & CStr(ws.Range("E1").Value) & "|" & CStr(ws.Range("F1").Value) & _
                          "|" & CStr(ws.Range("G1").Value) & "|" & CStr(ws.Range("H1").Value) & """"
        If Err.Number <> 0 Then
            Mod_Util.DebugLog "WriteLog(лист): ошибка отладочного логирования: " & Err.Description
            Err.Clear
        End If
    End If

    If lastRowLog >= 2 Then
        On Error Resume Next
        ws.ListObjects("tblAnomalies").Resize ws.Range("A1:H" & lastRowLog)
        On Error GoTo 0
    Else
        Mod_Util.DebugLog "WriteLog(лист): нет данных для записи (после обработки последняя строка = " & _
                          lastRowLog & "), Resize не выполняется"
    End If

    ' W-DEBUG-ANOMALIES-EMPTY: состояние ListObject ПОСЛЕ Resize (и проверка,
    ' что заголовки после Resize остались на месте — иначе видны дефолтные
    ' "Столбец1...", означающие потерю структуры таблицы).
    ' W-FIX-ANOMALIES-EMPTY (блокер B): этот блок также находится внутри
    ' локальной обработки ошибок (On Error Resume Next), начатой выше.
    If Not loAnomDbg Is Nothing Then
        Mod_Util.DebugLog "WriteLog(лист): после Resize, ListObject строк = " & _
                          loAnomDbg.ListRows.Count
        If Err.Number <> 0 Then
            Mod_Util.DebugLog "WriteLog(лист): ошибка отладочного логирования: " & Err.Description
            Err.Clear
        End If
        hdrAddrDbg = loAnomDbg.HeaderRange.Address
        If Err.Number <> 0 Then
            Mod_Util.DebugLog "WriteLog(лист): ошибка отладочного логирования: " & Err.Description
            Err.Clear
        End If
        Mod_Util.DebugLog "WriteLog(лист): после Resize, HeaderRange = " & _
                          hdrAddrDbg & _
                          ", строк заголовка = " & loAnomDbg.HeaderRange.Rows.Count & _
                          ", столбцов = " & loAnomDbg.HeaderRange.Columns.Count
        If Err.Number <> 0 Then
            Mod_Util.DebugLog "WriteLog(лист): ошибка отладочного логирования: " & Err.Description
            Err.Clear
        End If
        Mod_Util.DebugLog "WriteLog(лист): после Resize, заголовки A1:H1 = """ & _
                          CStr(ws.Range("A1").Value) & "|" & CStr(ws.Range("B1").Value) & _
                          "|" & CStr(ws.Range("C1").Value) & "|" & CStr(ws.Range("D1").Value) & _
                          "|" & CStr(ws.Range("E1").Value) & "|" & CStr(ws.Range("F1").Value) & _
                          "|" & CStr(ws.Range("G1").Value) & "|" & CStr(ws.Range("H1").Value) & """"
        If Err.Number <> 0 Then
            Mod_Util.DebugLog "WriteLog(лист): ошибка отладочного логирования: " & Err.Description
            Err.Clear
        End If
    End If
    ' W-FIX-ANOMALIES-EMPTY (блокер B): конец блока отладочного логирования —
    ' возвращаем стандартную обработку ошибок. Ошибки в этом блоке не приводят
    ' к прерыванию WriteLog и выходу в Fail главной процедуры.
    On Error GoTo 0

    ' W-FIX-AUTOFIT-WRITELOG: AutoFit через SafeAutoFit — на скрытом листе
    ' прямой вызов падал с "Object doesn't support this property or method"
    ' и ронял всю сборку через Fail главной процедуры.
    SafeAutoFit ws, "A:H"
End Sub

'---------------- W-STRUCTURE-DUPLICATES-FIX, п.1: группировка дублей ---------
' Собирает словарь дублей по ключу "Тип|Номер" (колонка "Тип" — Nodes(i).NodeType).
' Дубли считаются ТОЛЬКО ВНУТРИ одного типа: совпадения номеров у разных типов
' (Папка vs Документ) дублем НЕ считаются. Узлы-ярлыки ("УДАЛЕН", "ОШИБКА_ЯРЛЫКА")
' и узлы без ведущего номера в расчёт не берутся.
' Побочный эффект: заполняет dupCountDict (ключ "Тип|Номер" -> True) — он
' используется в WriteFlatTable для колонки "Аномалия" и жёлтой подсветки (п.3).
' Возвращает Object (Scripting.Dictionary) или Nothing, если дублей нет.
Private Function BuildDupKeyDict() As Object
    Dim dictKey As Object         ' "Тип|Номер" -> Collection индексов узлов
    Dim i As Long
    Dim nm As String, code As String
    Dim key As String
    Dim col As Collection
    Dim hasDup As Boolean
    Dim kCheck As Variant
    ' W-FIX-ANOMALIES-EMPTY (блокер A): счётчики диагностики поиска дублей.
    Dim nodesWithCode As Long     ' узлы с непустым ведущим номером (LeadCode <> "")
    Dim nodesSkipped As Long      ' узлы, отброшенные по типу или пустому коду
    Dim dupGroupsDbg As Long      ' число групп с Count > 1
    Dim dbgExamples As Long       ' счётчик выводимых примеров ключей (до 5)
    Dim keysArr As Variant        ' массив ключей для вывода примеров

    Set dictKey = CreateObject("Scripting.Dictionary")

    ' W-FIX-ANOMALIES-EMPTY (блокер A): старт процедуры — видно, сколько узлов
    ' всего участвует в поиске дублей.
    Mod_Util.DebugLog "BuildDupKeyDict: старт, узлов = " & NodeCount

    ' Проход 1: группировка узлов по ключу "Тип|Номер"
    For i = 1 To NodeCount
        If Nodes(i).NodeType <> "УДАЛЕН" And Nodes(i).NodeType <> "ОШИБКА_ЯРЛЫКА" Then
            nm = Nodes(i).Name
            code = LeadCode(nm)
            If code <> "" Then
                nodesWithCode = nodesWithCode + 1
                ' КЛЮЧ УЧИТЫВАЕТ ТИП: дубли считаем только внутри одного типа
                key = Nodes(i).NodeType & "|" & code
                If Not dictKey.Exists(key) Then
                    Set col = New Collection
                    col.Add i
                    Set dictKey(key) = col
                Else
                    dictKey(key).Add i
                End If
            Else
                ' W-FIX-ANOMALIES-EMPTY (блокер A): узел отброшен — пустой код
                nodesSkipped = nodesSkipped + 1
            End If
        Else
            ' W-FIX-ANOMALIES-EMPTY (блокер A): узел отброшен по типу
            ' ("УДАЛЕН"/"ОШИБКА_ЯРЛЫКА")
            nodesSkipped = nodesSkipped + 1
        End If
    Next i

    ' W-FIX-ANOMALIES-EMPTY (блокер A): итоги прохода 1 — если nodesWithCode = 0
    ' или dictKey.Count = 0, поиск дублей не может найти ни одной группы.
    Mod_Util.DebugLog "BuildDupKeyDict: узлов с номером = " & nodesWithCode & _
                      ", отброшено = " & nodesSkipped & ", ключей = " & dictKey.Count

    ' Есть ли хотя бы одна группа из 2+ узлов?
    hasDup = False
    For Each kCheck In dictKey.Keys
        If dictKey(kCheck).Count > 1 Then hasDup = True: Exit For
    Next kCheck

    ' W-FIX-ANOMALIES-EMPTY (блокер A): число групп с дублями и примеры ключей
    ' (до 5 любых), чтобы видеть, что именно собрал словарь. Логику поиска
    ' дублей это не меняет — только диагностика.
    For Each kCheck In dictKey.Keys
        If dictKey(kCheck).Count > 1 Then dupGroupsDbg = dupGroupsDbg + 1
    Next kCheck
    Mod_Util.DebugLog "BuildDupKeyDict: групп с дублями = " & dupGroupsDbg

    If dictKey.Count > 0 Then
        keysArr = dictKey.Keys
        dbgExamples = 0
        For i = 0 To UBound(keysArr)
            If dbgExamples >= 5 Then Exit For
            Mod_Util.DebugLog "BuildDupKeyDict: пример ключа = " & CStr(keysArr(i)) & _
                              " (узлов = " & dictKey(keysArr(i)).Count & ")"
            dbgExamples = dbgExamples + 1
        Next i
    End If

    If Not hasDup Then
        Set BuildDupKeyDict = Nothing
        Exit Function
    End If

    ' Помечаем ВСЕ узлы дублирующихся групп для подсветки в WriteFlatTable (п.3)
    Dim j As Long
    Dim idx As Long
    Dim k As Variant
    If dupCountDict Is Nothing Then Set dupCountDict = CreateObject("Scripting.Dictionary")
    For Each k In dictKey.Keys
        If dictKey(k).Count > 1 Then
            For j = 1 To dictKey(k).Count
                idx = CLng(dictKey(k)(j))
                dupCountDict(Nodes(idx).NodeType & "|" & LeadCode(Nodes(idx).Name)) = True
            Next j
        End If
    Next k

    Set BuildDupKeyDict = dictKey
End Function

'---------------- W-STRUCTURE-DUPLICATES-FIX, п.2: запись дублей в журнал -----
' W-STRUCTURE-DUPLICATES-SAVE-FIX: усилена проверка записи — каждая пара дублей
' пишется в СВОЮ строку (r инкрементируется на каждой итерации), перед выходом
' проверяется, что в колонке A реально стоят значения "DUP_IN_TYPE" (иначе
' запись считается потерянной и фиксируется предупреждение).
' Записывает строки дублей в лист "Аномалии схемы" (ws), начиная со строки
' startRow, строго по колонкам журнала:
'   A="DUP_IN_TYPE" (код проблемы), B/C/E/G — trace-узла, F="Наименование 1 | Наименование 2",
'   H="[!] Дубль: <Тип>, номер <Номер>".
' Колонки "Тип"/"Номер"/"Наименование 1"/"Наименование 2"/"Тип аномалии"
' присутствуют в описании колонок каждого выводимого листа (см. комментарии).
' Возвращает номер последней записанной строки (0 — дублей нет).
Private Function WriteSchemaDuplicates(ByVal ws As Worksheet, ByVal startRow As Long, _
                                       ByVal dictKey As Object) As Long
    Dim k As Variant
    Dim j As Long
    Dim idxA As Long, idxB As Long
    Dim typePart As String, numPart As String
    Dim posBar As Long
    Dim r As Long
    Dim dupGroups As Long          ' W-DEBUG-ANOMALIES-EMPTY: число групп-дублей
    Dim dupPairs As Long           ' W-DEBUG-ANOMALIES-EMPTY: число пар (строк) дублей

    ' W-DEBUG-ANOMALIES-EMPTY: старт процедуры.
    Mod_Util.DebugLog "WriteSchemaDuplicates: старт"

    If ws Is Nothing Then
        Mod_Util.DebugLog "WriteSchemaDuplicates: лист не передан (ws Is Nothing), выход без записи"
        WriteSchemaDuplicates = 0
        Exit Function
    End If
    If dictKey Is Nothing Then
        Mod_Util.DebugLog "WriteSchemaDuplicates: словарь дублей не передан (dictKey Is Nothing), выход без записи"
        WriteSchemaDuplicates = 0
        Exit Function
    End If

    ' Подсчёт найденных дублей ДО записи (только логирование, логику не меняем):
    ' dupGroups — группы из 2+ узлов, dupPairs — строки, которые будут записаны
    ' (по одной на каждую пару "первый элемент + следующий").
    For Each k In dictKey.Keys
        If dictKey(k).Count > 1 Then
            dupGroups = dupGroups + 1
            dupPairs = dupPairs + (dictKey(k).Count - 1)
        End If
    Next k
    Mod_Util.DebugLog "WriteSchemaDuplicates: найдено дублей = " & dupPairs & _
                      " (групп = " & dupGroups & ")"

    ' W-DEBUG-ANOMALIES-EMPTY: сообщение о начале записи на лист. Если дублей
    ' нет (dupPairs = 0), запись не выполняется — это тоже видно в журнале.
    If dupPairs > 0 Then
        Mod_Util.DebugLog "WriteSchemaDuplicates: запись на лист, строк = " & dupPairs
    Else
        Mod_Util.DebugLog "WriteSchemaDuplicates: нет дублей для записи на лист"
    End If

    r = startRow
    For Each k In dictKey.Keys
        If dictKey(k).Count > 1 Then
            posBar = InStr(CStr(k), "|")
            typePart = Left$(CStr(k), posBar - 1)      ' Тип
            numPart = Mid$(CStr(k), posBar + 1)         ' Номер
            ' каждое следующее наименование — в пару к первому
            For j = 2 To dictKey(k).Count
                idxA = CLng(dictKey(k)(1))
                idxB = CLng(dictKey(k)(j))
                ws.Cells(r, 1).Value = "DUP_IN_TYPE"                            ' Код проблемы
                ws.Cells(r, 2).Value = Nodes(idxA).RowIdx                       ' Строка
                ws.Cells(r, 3).Value = Nodes(idxA).ColIdx                       ' Столбец
                ws.Cells(r, 4).Value = Nodes(idxA).Level                        ' Уровень
                ws.Cells(r, 5).Value = typePart                                 ' Тип
                ws.Cells(r, 6).Value = NodeCaption(Nodes(idxA).Name) & " | " & _
                                       NodeCaption(Nodes(idxB).Name)           ' Элемент: Наименование 1 | Наименование 2
                ws.Cells(r, 7).Value = ParentName(Nodes(idxA).parentIdx)        ' Родитель
                ws.Cells(r, 8).Value = "[!] Дубль: тип """ & typePart & _
                                       """, номер " & numPart                   ' Описание / Тип аномалии
                r = r + 1
            Next j
        End If
    Next k

    If r > startRow Then
        ' W-DEBUG-ANOMALIES-EMPTY: верификация фактической записи строк дублей.
        Mod_Util.DebugLog "WriteSchemaDuplicates: запись завершена, последняя строка = " & (r - 1)
        ' W-STRUCTURE-DUPLICATES-SAVE-FIX: верификация фактической записи —
        ' строки дублей должны реально стоять в колонке A листа.
        Dim dupWritten As Long
        dupWritten = Application.WorksheetFunction.CountA(ws.Range("A" & startRow & ":A" & (r - 1)))
        Mod_Util.WriteLog "Проверка дублей: найдено аномалий - " & (r - startRow)
        If dupWritten = 0 Then
            Mod_Util.WriteWarning "S010", "WriteSchemaDuplicates: записано " & _
                (r - startRow) & " строк, но на листе ""Аномалии схемы"" значений нет — проверка сохранности."
        End If
        WriteSchemaDuplicates = r - 1
    Else
        ' W-DEBUG-ANOMALIES-EMPTY: ни одной строки дублей не записано
        ' (в основном потоке этот выход означает, что дубли могли быть потеряны
        ' на этапе сбора — см. сообщение "найдено дублей" выше).
        Mod_Util.DebugLog "WriteSchemaDuplicates: запись не выполнялась (дублей нет), возврат 0"
        WriteSchemaDuplicates = 0
    End If
End Function

'---------------- W-CORRECTION-PACK-1, п.3.1 + W-STRUCTURE-DUPLICATES-FIX ---
' Резервная точка входа проверки дублей (на случай, если основной поток не записал
' строки дублей). W-STRUCTURE-DUPLICATES-SAVE-FIX: в текущем потоке порядок такой:
' PrepareAnomaliesSheet -> BuildDupKeyDict -> WriteSchemaDuplicates ->
' CreateSmartTable(tblAnomalies) -> WriteLog -> WriteFlatTable ->
' HighlightDuplicateRows. Здесь только страховка: если на листе "Аномалии схемы"
' нет ни одной строки "DUP_IN_TYPE", дубли дописываются в конец листа (без
' очистки!). Сборку НЕ блокируем — только предупреждения.
Private Sub CheckNumberDuplicates(ByVal wbOut As Workbook)
    Dim dictKey As Object           ' "Тип|Номер" -> Collection индексов узлов

    On Error GoTo ErrHandler

    ' dupCountDict уже заполнен в BuildDupKeyDict (главная процедура); повторный
    ' вызов безопасен (группировка идемпотентна).
    Set dictKey = BuildDupKeyDict
    If dictKey Is Nothing Then Exit Sub     ' дублей нет

    ' Если журнал "Аномалии схемы" ещё не содержит строк дублей (основная запись
    ' не выполнилась) — дописываем их в конец существующего листа, структуру не трогаем.
    Dim ws As Worksheet
    GetSheetNoClear wbOut, "Аномалии схемы", ws
    If Not ws Is Nothing Then
        If Application.CountIf(ws.Columns(1), "DUP_IN_TYPE") = 0 Then
            Dim lastRowBackup As Long
            lastRowBackup = WriteSchemaDuplicates(ws, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1, dictKey)
            ' ListObject создаётся/расширяется ПОСЛЕ дозаписанных данных
            If lastRowBackup >= 2 Then
                Call CreateSmartTable(ws, ws.Range("A1:H" & lastRowBackup), _
                                      "tblAnomalies", "TableStyleMedium6")
            End If
        End If
    End If
    Exit Sub

ErrHandler:
    ' проверка дублей не должна ронять сборку — только предупреждение
    Mod_Util.WriteWarning "S006", "CheckNumberDuplicates: ошибка проверки дублей. Ошибка: " & Err.Description
    Err.Clear
End Sub

' Возвращает лист с именем nm, если он существует; если нет — Nothing
' (БЕЗ создания и БЕЗ очистки — в отличие от GetSheet).
Private Sub GetSheetNoClear(ByVal wb As Workbook, ByVal nm As String, ByRef out As Worksheet)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    Set out = ws
End Sub

'---------------- W-STRUCTURE-DUPLICATES-FIX, п.3: подсветка дублей ----------
' Применяется ПОСЛЕ WriteFlatTable (после создания Умной таблицы
' tblFlatStructure), чтобы жёлтая заливка не перекрывалась стилями таблицы.
' Для строк, участвующих в дублях (dupCountDict, ключ "Тип|Номер"):
'  - в колонку "Аномалия" (E) ставится "[!] Дубль" (если ещё не поставлено);
'  - вся строка A:E красится в жёлтый цвет RGB(255, 255, 0).
' Проблемные строки (hasError, розовая подсветка WriteFlatTable) НЕ перекрашиваются.
Private Sub HighlightDuplicateRows(ByVal wbOut As Workbook)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim r As Long
    Dim isDup As Boolean
    Dim dupKey As String

    On Error GoTo ErrHandler

    If dupCountDict Is Nothing Then Exit Sub
    If dupCountDict.Count = 0 Then Exit Sub

    Set ws = wbOut.Worksheets(FLAT_SHEET_NAME)
    Set lo = ws.ListObjects("tblFlatStructure")
    If lo.ListRows.Count = 0 Then Exit Sub

    r = 2                                   ' первая строка данных под заголовками
    For i = 1 To NodeCount
        If Nodes(i).NodeType <> "УДАЛЕН" Then
            isDup = False
            dupKey = Nodes(i).NodeType & "|" & LeadCode(Nodes(i).Name)
            If Right$(dupKey, 1) <> "|" Then
                If dupCountDict.Exists(dupKey) Then isDup = True
            End If
            If isDup Then
                ' не перекрашиваем строки ошибок (розовая подсветка WriteFlatTable)
                If Not problemsDict.Exists(Nodes(i).Name) And _
                   Nodes(i).NodeType <> "ОШИБКА_ЯРЛЫКА" Then
                    If Len(CStr(ws.Cells(r, 5).Value)) = 0 Then
                        ws.Cells(r, 5).Value = "[!] Дубль"      ' колонка "Аномалия"
                    End If
                    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Interior.Color = RGB(255, 255, 0)
                End If
            End If
            r = r + 1
        End If
    Next i
    Exit Sub

ErrHandler:
    ' подсветка не должна ронять сборку — только предупреждение
    Mod_Util.WriteWarning "S009", "HighlightDuplicateRows: ошибка подсветки дублей. Ошибка: " & Err.Description
    Err.Clear
End Sub

' W-STRUCTURE-DUPLICATES-FIX: наименование узла без ведущего номера
' (для колонок "Наименование 1"/"Наименование 2" листа "Аномалии схемы").
Private Function NodeCaption(ByVal s As String) As String
    Dim code As String
    Dim res As String
    code = LeadCode(s)
    If code <> "" Then
        res = Trim$(Mid$(s, Len(code) + 1))
        If Left$(res, 1) = "." Then res = Trim$(Mid$(res, 2))
        NodeCaption = res
    Else
        NodeCaption = s
    End If
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

    Mod_Util.DebugLog "BackupFlatTableSheet: старт"

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
    If wsSrc Is Nothing Then
        Mod_Util.DebugLog "BackupFlatTableSheet: старый лист не найден, бэкап не требуется"
        Exit Sub      ' старого листа нет — бэкап не нужен
    End If

    bakName = FLAT_SHEET_NAME & "_bak_" & Format$(Now, "yyyymmdd_hhnnss")

    Application.DisplayAlerts = False
    wsSrc.Copy After:=wbTool.Sheets(wbTool.Sheets.Count)
    ActiveSheet.Name = bakName
    Application.DisplayAlerts = True

    Mod_Util.DebugLog "BackupFlatTableSheet: бэкап создан"
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



