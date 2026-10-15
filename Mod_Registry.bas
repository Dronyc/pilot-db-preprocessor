Attribute VB_Name = "Mod_Registry"
Option Explicit

' ============================================================================
' Модуль реестров: построение, запись файлов реестров, обновление листов.
' Реализация W6 (OPT-13). Данные берутся из листов "Типы" и "Атрибуты"
' книги сводной таблицы; файлы пишутся в
' <корень>\03_Результаты генерации\03_Реестры\ (см. Spec_Dispatcher.md,
' разделы 2, 7 и 11).
' Коды ошибок модуля: R001 — нет листа "Типы", R002 — нет листа "Атрибуты",
' R003 — ошибка записи файла, R004 — ошибка создания папки реестров.
' ============================================================================

' ---------------------------------------------------------------------------
' КОНСТАНТЫ
' ---------------------------------------------------------------------------
Private Const SHEET_TYPES As String = "Типы"
Private Const SHEET_ATTRS As String = "Атрибуты"
Private Const SHEET_REG_TYPES As String = "_Реестр_типов"
Private Const SHEET_REG_COV As String = "_Реестр_покрытия"
Private Const FILE_REG_TYPES As String = "Реестр_типов.txt"
Private Const FILE_REG_COV As String = "Реестр_покрытия.txt"
Private Const RESULTS_SUBFOLDER As String = "\03_Результаты генерации\03_Реестры"
Private Const SHEET_PASSWORD As String = "pilot"

' ---------------------------------------------------------------------------
' МОДУЛЬНЫЕ ПЕРЕМЕННЫЕ (Private, явная очистка в начале BuildRegistries)
' ---------------------------------------------------------------------------
Private RegTypes As Object      ' словарь: все системные имена типов из листа "Типы"
Private RegCoverage As Object   ' словарь: уникальные имена типов из листа "Атрибуты"
Private m_wbSource As Workbook  ' книга-источник данных для реестров
Private m_wsTypes As Worksheet  ' найденный лист "Типы"
Private m_wsAttrs As Worksheet  ' найденный лист "Атрибуты"

' ---------------------------------------------------------------------------
' ВСПОМОГАТЕЛЬНЫЕ ПРОЦЕДУРЫ (только для внутреннего использования модуля)
' ---------------------------------------------------------------------------

' Инициализирует словари реестров (создаёт, если ещё не созданы).
Private Sub EnsureDicts()
    If RegTypes Is Nothing Then Set RegTypes = CreateObject("Scripting.Dictionary")
    If RegCoverage Is Nothing Then Set RegCoverage = CreateObject("Scripting.Dictionary")
End Sub

' Явная очистка состояния реестров перед новым построением (OPT-09).
Private Sub ClearState()
    Call EnsureDicts
    RegTypes.RemoveAll
    RegCoverage.RemoveAll
    Set m_wbSource = Nothing
    Set m_wsTypes = Nothing
    Set m_wsAttrs = Nothing
End Sub

' Возвращает книгу-источник данных для реестров:
' 1) активная книга, если в ней есть листы "Типы" и "Атрибуты";
' 2) иначе — последняя открытая книга с такими листами (книга сводной);
' 3) иначе — книга Препроцессора (ThisWorkbook).
Private Function ResolveSourceWorkbook() As Workbook
    Dim wb As Workbook
    Dim wbFound As Workbook

    Set wb = Application.ActiveWorkbook
    If Not wb Is Nothing Then
        If SheetExists(wb, SHEET_TYPES) And SheetExists(wb, SHEET_ATTRS) Then
            Set ResolveSourceWorkbook = wb
            Exit Function
        End If
    End If

    For Each wb In Application.Workbooks
        If SheetExists(wb, SHEET_TYPES) And SheetExists(wb, SHEET_ATTRS) Then
            Set wbFound = wb    ' берём последнюю подходящую открытую книгу
        End If
    Next wb
    If Not wbFound Is Nothing Then
        Set ResolveSourceWorkbook = wbFound
        Exit Function
    End If

    Set ResolveSourceWorkbook = ThisWorkbook
End Function

' Возвращает True, если в книге wb есть лист с именем sheetName.
Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Nothing
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' Возвращает лист sheetName из книги wb; если листа нет — Nothing.
Private Function GetSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    If SheetExists(wb, sheetName) Then Set GetSheet = wb.Worksheets(sheetName)
End Function

' Собирает системные имена типов из колонки 1 листа ws, начиная со 2-й строки,
' в словарь dict (ключ — имя, элемент — количество строк). Пустые ячейки пропускаются.
Private Sub CollectNamesFromColumn1(ByVal ws As Worksheet, ByRef dict As Object)
    Dim lastRow As Long
    Dim r As Long
    Dim nm As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        nm = Trim$(CStr(ws.Cells(r, 1).Value))
        If nm <> "" Then
            If dict.Exists(nm) Then
                dict(nm) = dict(nm) + 1
            Else
                dict.Add nm, 1
            End If
        End If
    Next r
End Sub

' Возвращает ключи словаря dict, отсортированные по алфавиту (регистронезависимо),
' в виде одномерного массива строк.
Private Function GetSortedKeys(dict As Object) As Variant
    Dim keys As Variant
    Dim i As Long
    Dim j As Long
    Dim tmp As Variant

    If dict.Count = 0 Then
        GetSortedKeys = Array()
        Exit Function
    End If

    keys = dict.Keys
    For i = LBound(keys) + 1 To UBound(keys)
        tmp = keys(i)
        j = i - 1
        Do While j >= LBound(keys)
            If StrComp(CStr(keys(j)), CStr(tmp), vbTextCompare) <= 0 Then Exit Do
            keys(j + 1) = keys(j)
            j = j - 1
        Loop
        keys(j + 1) = tmp
    Next i
    GetSortedKeys = keys
End Function

' Создаёт папку path, если её ещё нет (проверка прав через Err.Number).
' При неудаче выбрасывает ошибку — вызывающий код фиксирует её кодом R004.
Private Sub EnsureFolder(ByVal path As String)
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
End Sub

' Сохраняет список names (массив строк) в текстовый файл filePath,
' перезаписывая существующий файл. Одна запись на строку.
Private Sub WriteTextLines(ByVal filePath As String, names As Variant)
    Dim fnum As Integer
    Dim i As Long

    fnum = FreeFile
    Open filePath For Output As #fnum
    For i = LBound(names) To UBound(names)
        Print #fnum, CStr(names(i))
    Next i
    Close #fnum
End Sub

' Возвращает текст ошибки в безопасном виде (Err может быть пуст).
Private Function ErrDescSafe() As String
    ErrDescSafe = ""
    On Error Resume Next
    ErrDescSafe = Err.Description
    On Error GoTo 0
End Function

' ============================================================================
' 1. Строит внутренние реестры объектов.
' Собирает данные для двух реестров из сводной таблицы:
' RegTypes    — все системные имена типов из листа "Типы" (колонка 1);
' RegCoverage — уникальные системные имена типов из листа "Атрибуты" (колонка 1).
' ============================================================================
Public Sub BuildRegistries()
    On Error GoTo ErrHandler

    ' Явная очистка состояния в начале построения (OPT-09)
    Call ClearState

    ' Источник данных: активная книга либо книга сводной, если она открыта
    Set m_wbSource = ResolveSourceWorkbook()

    ' Лист "Типы": обязателен (код ошибки R001)
    Set m_wsTypes = GetSheet(m_wbSource, SHEET_TYPES)
    If m_wsTypes Is Nothing Then
        Call ReportMissingSheet("R001", SHEET_TYPES)
        GoTo CleanExit
    End If

    ' Лист "Атрибуты": обязателен (код ошибки R002)
    Set m_wsAttrs = GetSheet(m_wbSource, SHEET_ATTRS)
    If m_wsAttrs Is Nothing Then
        Call ReportMissingSheet("R002", SHEET_ATTRS)
        GoTo CleanExit
    End If

    ' Проход по листу "Типы" (со 2-й строки): собираем все системные имена типов
    Call CollectNamesFromColumn1(m_wsTypes, RegTypes)

    ' Проход по листу "Атрибуты" (со 2-й строки): собираем уникальные имена типов
    Call CollectNamesFromColumn1(m_wsAttrs, RegCoverage)

    Mod_Util.WriteLog "BuildRegistries: типов " & RegTypes.Count & ", покрытий " & RegCoverage.Count
    GoTo CleanExit

ErrHandler:
    Mod_Util.WriteError "R001", "BuildRegistries: ошибка построения реестров. Ошибка: " & ErrDescSafe()

CleanExit:
    Err.Clear
End Sub

' ============================================================================
' 2. Записывает файлы реестров на диск.
' Папка: <корень>\03_Результаты генерации\03_Реестры\
' Файлы: Реестр_типов.txt и Реестр_покрытия.txt — по одному системному имени
' типа на строку, сортировка по алфавиту, перезапись существующих файлов.
' ============================================================================
Public Sub WriteRegistryFiles()
    Dim rootPath As String
    Dim folderPath As String
    Dim typesArr As Variant
    Dim covArr As Variant

    On Error GoTo ErrFolder

    ' Если реестры ещё не построены — строим их сейчас
    If RegTypes Is Nothing Or RegCoverage Is Nothing Then
        Call BuildRegistries
    End If

    rootPath = Mod_Util.GetProjectRoot()
    If rootPath = "" Then
        Mod_Util.WriteError "R004", "WriteRegistryFiles: Mod_Util.GetProjectRoot вернул пустой путь, " & _
                                    "папка реестров не определена."
        GoTo CleanExit
    End If
    If Right$(rootPath, 1) = "\" Then rootPath = Left$(rootPath, Len(rootPath) - 1)

    folderPath = rootPath & RESULTS_SUBFOLDER

    ' Создание папки реестров с проверкой прав (при ошибке — код R004)
    On Error GoTo ErrFolder
    Call EnsureFolder(folderPath)

    ' Сортировка имён по алфавиту
    typesArr = GetSortedKeys(RegTypes)
    covArr = GetSortedKeys(RegCoverage)

    ' Запись файлов с перезаписью существующих (при ошибке — код R003)
    On Error GoTo ErrWrite
    Call WriteTextLines(folderPath & "\" & FILE_REG_TYPES, typesArr)
    Call WriteTextLines(folderPath & "\" & FILE_REG_COV, covArr)

    Mod_Util.WriteLog "WriteRegistryFiles: сохранено " & RegTypes.Count & " типов, " & _
                      RegCoverage.Count & " покрытий"
    GoTo CleanExit

ErrFolder:
    Mod_Util.WriteError "R004", "WriteRegistryFiles: не удалось создать папку реестров '" & _
                                folderPath & "'. Ошибка: " & ErrDescSafe()
    GoTo CleanExit

ErrWrite:
    Mod_Util.WriteError "R003", "WriteRegistryFiles: ошибка записи файла реестра в папку '" & _
                                folderPath & "'. Ошибка: " & ErrDescSafe()

CleanExit:
    Err.Clear
End Sub

' ============================================================================
' 3. Обновляет листы реестров в книге Препроцессора.
' Листы _Реестр_типов и _Реестр_покрытия: очищаются (или создаются),
' заполняются системными именами типов в колонке A и защищаются паролем "pilot".
' ============================================================================
Public Sub UpdateRegistrySheets()
    Dim typesArr As Variant
    Dim covArr As Variant

    On Error GoTo ErrHandler

    ' Если реестры ещё не построены — строим их сейчас
    If RegTypes Is Nothing Or RegCoverage Is Nothing Then
        Call BuildRegistries
    End If

    typesArr = GetSortedKeys(RegTypes)
    covArr = GetSortedKeys(RegCoverage)

    Call WriteRegistrySheet(ThisWorkbook, SHEET_REG_TYPES, typesArr)
    Call WriteRegistrySheet(ThisWorkbook, SHEET_REG_COV, covArr)

    Mod_Util.WriteLog "UpdateRegistrySheets: обновлено " & RegTypes.Count & " типов, " & _
                      RegCoverage.Count & " покрытий"
    GoTo CleanExit

ErrHandler:
    Mod_Util.WriteError "R003", "UpdateRegistrySheets: ошибка обновления листов реестров. Ошибка: " & ErrDescSafe()

CleanExit:
    Err.Clear
End Sub

' Записывает отсортированный список names на служебный лист sheetName книги wb:
' лист создаётся при отсутствии, существующий очищается (UsedRange.ClearContents),
' затем в колонку A записываются имена и лист защищается паролем "pilot".
Private Sub WriteRegistrySheet(ByVal wb As Workbook, ByVal sheetName As String, names As Variant)
    Dim ws As Worksheet
    Dim i As Long

    ' Найти или создать лист
    Set ws = GetSheet(wb, sheetName)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    Else
        ' Существующий лист мог быть защищён ранее — снимаем защиту
        On Error Resume Next
        ws.Unprotect Password:=SHEET_PASSWORD
        On Error GoTo 0
        ' Очистка содержимого использованного диапазона
        ws.UsedRange.ClearContents
    End If

    ' Заголовок и данные в колонке A
    ws.Cells(1, 1).Value = "Системное имя типа"
    ws.Cells(1, 1).Font.Bold = True
    For i = LBound(names) To UBound(names)
        ws.Cells(i + 2, 1).Value = CStr(names(i))
    Next i
    ws.Columns(1).AutoFit

    ' Защита служебного листа паролем "pilot"
    On Error Resume Next
    ws.Protect Password:=SHEET_PASSWORD
    On Error GoTo 0
End Sub
