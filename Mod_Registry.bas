Attribute VB_Name = "Mod_Registry"
Option Explicit

' ============================================================================
' Модуль реестров: построение, запись файлов реестров, обновление листов.
' Реализация W4 (OPT-08): RegTypes — все системные имена из листа «Типы»,
' RegCoverage — уникальные системные имена типов с покрытием из листа
' «Атрибуты». Файлы пишутся в 03_Результаты генерации\03_Реестры,
' листы _Реестр_типов и _Реестр_покрытия защищаются паролем «pilot».
' ============================================================================

' Модульные словари-реестры (доступны только внутри Mod_Registry)
Private RegTypes As Object      ' системные имена типов (множество)
Private RegCoverage As Object   ' системные имена типов с покрытием (множество)

' Возвращает отсортированный по алфавиту массив ключей словаря.
' Пузырьковая сортировка достаточна для имён реестров и не требует
' дополнительных зависимостей.
Private Function SortedKeys(ByVal dict As Object) As String()
    Dim keys As Variant
    Dim arr() As String
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim tmp As String

    If dict.Count = 0 Then
        ReDim arr(0 To -1)
        SortedKeys = arr
        Exit Function
    End If

    keys = dict.keys
    n = UBound(keys) - LBound(keys) + 1
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = CStr(keys(i - 1))
    Next i

    For i = 2 To n
        tmp = arr(i)
        j = i - 1
        Do While j >= 1
            If StrComp(arr(j), tmp, vbTextCompare) <= 0 Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i

    SortedKeys = arr
End Function

' Сохраняет массив строк в текстовый файл (перезаписывает существующий).
Private Sub WriteTextFile(ByVal filePath As String, ByVal arr() As String)
    Dim fnum As Integer
    Dim i As Long

    fnum = FreeFile
    Open filePath For Output As #fnum
    For i = LBound(arr) To UBound(arr)
        Print #fnum, arr(i)
    Next i
    Close #fnum
End Sub

' Строит внутренние реестры объектов.
Public Sub BuildRegistries()
    Dim wsT As Worksheet      ' лист «Типы»
    Dim wsA As Worksheet      ' лист «Атрибуты»
    Dim lastRow As Long
    Dim r As Long
    Dim nm As String

    On Error GoTo ErrHandler

    ' Очистка модульных словарей перед заполнением (гарантия повторного запуска)
    Set RegTypes = CreateObject("Scripting.Dictionary")
    Set RegCoverage = CreateObject("Scripting.Dictionary")

    ' Поиск листа «Типы» в книге сводной (активная книга сборки)
    On Error Resume Next
    Set wsT = ActiveWorkbook.Worksheets("Типы")
    On Error GoTo ErrHandler
    If wsT Is Nothing Then
        Mod_Util.WriteError "R001", "BuildRegistries: в книге не найден лист 'Типы'."
        MsgBox "Ошибка R001: в книге не найден лист 'Типы'.", vbCritical, "Реестры"
        Err.Clear
        Exit Sub
    End If

    ' Поиск листа «Атрибуты»
    On Error Resume Next
    Set wsA = ActiveWorkbook.Worksheets("Атрибуты")
    On Error GoTo ErrHandler
    If wsA Is Nothing Then
        Mod_Util.WriteError "R002", "BuildRegistries: в книге не найден лист 'Атрибуты'."
        MsgBox "Ошибка R002: в книге не найден лист 'Атрибуты'.", vbCritical, "Реестры"
        Err.Clear
        Exit Sub
    End If

    ' Реестр типов: обход листа «Типы» со 2-й строки, колонка 1
    lastRow = wsT.Cells(wsT.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        nm = Trim(CStr(wsT.Cells(r, 1).Value))
        If nm <> "" Then
            If Not RegTypes.Exists(nm) Then RegTypes.Add nm, True
        End If
    Next r

    ' Реестр покрытия: уникальные системные имена типов из листа «Атрибуты»
    lastRow = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        nm = Trim(CStr(wsA.Cells(r, 1).Value))
        If nm <> "" Then
            If Not RegCoverage.Exists(nm) Then RegCoverage.Add nm, True
        End If
    Next r

    Mod_Util.WriteLog "BuildRegistries: типов " & RegTypes.Count & ", покрытий " & RegCoverage.Count
    Exit Sub

ErrHandler:
    Mod_Util.WriteError "R001", "BuildRegistries: ошибка построения реестров. Ошибка: " & Err.Description
    MsgBox "Ошибка при построении реестров: " & Err.Description, vbCritical, "Реестры"
    Err.Clear
End Sub

' Записывает файлы реестров на диск.
Public Sub WriteRegistryFiles()
    Dim path As String
    Dim arrT() As String
    Dim arrC() As String

    On Error GoTo ErrHandler

    If RegTypes Is Nothing Or RegCoverage Is Nothing Then
        Mod_Util.WriteError "R003", "WriteRegistryFiles: реестры не построены, вызовите BuildRegistries."
        MsgBox "Ошибка R003: реестры не построены. Сначала вызовите BuildRegistries.", vbCritical, "Реестры"
        Err.Clear
        Exit Sub
    End If

    ' Папка реестров: корень проекта + подпуть (по разделу 2 спецификации)
    path = Mod_Util.GetProjectRoot & "\03_Результаты генерации\03_Реестры"

    ' Создание папки, если её нет (проверка прав через ошибку MkDir)
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If

    ' Файл Реестр_типов.txt: по одному имени на строку, по алфавиту
    arrT = SortedKeys(RegTypes)
    WriteTextFile path & "\Реестр_типов.txt", arrT

    ' Файл Реестр_покрытия.txt: по одному имени на строку, по алфавиту
    arrC = SortedKeys(RegCoverage)
    WriteTextFile path & "\Реестр_покрытия.txt", arrC

    Mod_Util.WriteLog "WriteRegistryFiles: сохранено " & RegTypes.Count & " типов, " & _
                      RegCoverage.Count & " покрытий"
    Exit Sub

ErrHandler:
    Mod_Util.WriteError "R004", "WriteRegistryFiles: не удалось создать папку или записать файлы реестров (" & _
                                 path & "). Проверьте права доступа. Ошибка: " & Err.Description
    MsgBox "Ошибка R004: не удалось записать файлы реестров. Проверьте права доступа к папке." & vbCrLf & _
           path, vbCritical, "Реестры"
    Err.Clear
End Sub

' Обновляет листы реестров в книге.
Public Sub UpdateRegistrySheets()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim arr() As String
    Dim i As Long

    On Error GoTo ErrHandler

    If RegTypes Is Nothing Or RegCoverage Is Nothing Then
        Mod_Util.WriteError "R003", "UpdateRegistrySheets: реестры не построены, вызовите BuildRegistries."
        MsgBox "Ошибка R003: реестры не построены. Сначала вызовите BuildRegistries.", vbCritical, "Реестры"
        Err.Clear
        Exit Sub
    End If

    Set wb = ActiveWorkbook

    ' ---------- ЛИСТ _Реестр_типов ----------
    On Error Resume Next
    Set ws = wb.Worksheets("_Реестр_типов")
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        ' Листа нет — создаём в конце книги
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "_Реестр_типов"
    Else
        ws.Unprotect Password:="pilot"   ' снимаем защиту перед очисткой
        ws.UsedRange.ClearContents       ' очищаем содержимое UsedRange
    End If

    arr = SortedKeys(RegTypes)
    For i = LBound(arr) To UBound(arr)
        ws.Cells(i - LBound(arr) + 1, 1).Value = arr(i)
    Next i
    ws.Protect Password:="pilot"         ' защита паролем «pilot»
    Set ws = Nothing

    ' ---------- ЛИСТ _Реестр_покрытия ----------
    On Error Resume Next
    Set ws = wb.Worksheets("_Реестр_покрытия")
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "_Реестр_покрытия"
    Else
        ws.Unprotect Password:="pilot"
        ws.UsedRange.ClearContents
    End If

    arr = SortedKeys(RegCoverage)
    For i = LBound(arr) To UBound(arr)
        ws.Cells(i - LBound(arr) + 1, 1).Value = arr(i)
    Next i
    ws.Protect Password:="pilot"
    Set ws = Nothing

    Mod_Util.WriteLog "UpdateRegistrySheets: обновлено " & RegTypes.Count & " типов, " & _
                      RegCoverage.Count & " покрытий"
    Exit Sub

ErrHandler:
    Mod_Util.WriteError "R003", "UpdateRegistrySheets: ошибка обновления листов реестров. Ошибка: " & Err.Description
    MsgBox "Ошибка R003: не удалось обновить листы реестров. " & vbCrLf & Err.Description, _
           vbCritical, "Реестры"
    Err.Clear
End Sub
