Attribute VB_Name = "Mod_Prompt"
Option Explicit

' ============================================================================
' Модуль промптов: формирование промптов для следующего шага и исправлений,
' чтение списка и текущей позиции шагов.
' Реализация W4 (разделы 2, 3, 8, 9, 13 спецификации Spec_Dispatcher.md).
' Коды ошибок: P001 — нет Шаги.txt, P002 — ошибка формирования промта шага,
' P003 — ошибка формирования промта исправления.
' ============================================================================

' ---------------------------------------------------------------------------
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (Private)
' ---------------------------------------------------------------------------

' Возвращает путь без завершающего разделителя "\".
Private Function NormPath(ByVal p As String) As String
    Do While Len(p) > 1 And Right(p, 1) = "\"
        p = Left(p, Len(p) - 1)
    Loop
    NormPath = p
End Function

' Создаёт цепочку папок по полному пути (аналог mkdir -p).
' Препроцессор никогда не удаляет папки и файлы (раздел 3 спецификации).
Private Sub EnsureFolder(ByVal folderPath As String)
    Dim parts() As String
    Dim cur As String
    Dim i As Long

    If Dir(folderPath, vbDirectory) <> "" Then Exit Sub

    parts = Split(folderPath, "\")
    cur = parts(0)                       ' диск или UNC-префикс
    For i = 1 To UBound(parts)
        If parts(i) <> "" Then
            cur = cur & "\" & parts(i)
            If Dir(cur, vbDirectory) = "" Then MkDir cur
        End If
    Next i
End Sub

' Читает текстовый файл целиком в строку (через Line Input).
' Если файла нет — возвращает пустую строку (вызывающий сам решает,
' считать ли это ошибкой).
Private Function ReadTextFile(ByVal filePath As String) As String
    Dim fnum As Integer
    Dim s As String
    Dim res As String

    If Dir(filePath) = "" Then
        ReadTextFile = ""
        Exit Function
    End If

    fnum = FreeFile
    Open filePath For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, s
        res = res & s & vbCrLf
    Loop
    Close #fnum

    ' убрать финальный перевод строки
    If Len(res) >= 2 Then
        If Right(res, 2) = vbCrLf Then res = Left(res, Len(res) - 2)
    End If
    ReadTextFile = res
End Function

' Сохраняет строку в текстовый файл (перезаписывает существующий).
Private Sub SaveTextFile(ByVal filePath As String, ByVal content As String)
    Dim fnum As Integer
    Dim lines() As String
    Dim i As Long

    fnum = FreeFile
    Open filePath For Output As #fnum
    lines = Split(content, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        Print #fnum, lines(i)
    Next i
    Close #fnum
End Sub

' Заменяет плейсхолдер в тексте только если он там есть.
' Плейсхолдеры, отсутствующие в шаблоне, не добавляются.
Private Function ReplaceIfPresent(ByVal text As String, _
                                  ByVal placeholder As String, _
                                  ByVal value As String) As String
    If InStr(text, placeholder) > 0 Then
        ReplaceIfPresent = Replace(text, placeholder, value)
    Else
        ReplaceIfPresent = text
    End If
End Function

' Извлекает номер шага из записи коллекции "NN|Имя_папки".
Private Function StepNumFromEntry(ByVal entry As String) As String
    Dim pos As Long
    pos = InStr(entry, "|")
    If pos > 0 Then
        StepNumFromEntry = Trim(Left(entry, pos - 1))
    Else
        StepNumFromEntry = Trim(entry)
    End If
End Function

' Извлекает имя папки из записи коллекции "NN|Имя_папки".
Private Function StepNameFromEntry(ByVal entry As String) As String
    Dim pos As Long
    pos = InStr(entry, "|")
    If pos > 0 Then
        StepNameFromEntry = Trim(Mid(entry, pos + 1))
    Else
        StepNameFromEntry = ""
    End If
End Function

' Приводит номер шага к числовому виду (для сравнения "больше текущего").
Private Function StepNumToLong(ByVal nn As String) As Long
    On Error Resume Next
    StepNumToLong = CLng(nn)
    On Error GoTo 0
End Function

' Возвращает имя последнего изменённого файла по указанной маске
' в заданной папке. Если файлов нет — возвращает "".
Private Function LastModifiedFile(ByVal folderPath As String, ByVal mask As String) As String
    Dim fName As String
    Dim best As String
    Dim bestDate As Date

    fName = Dir(folderPath & "\" & mask)
    Do While fName <> ""
        If Left(fName, 2) <> "~$" Then           ' пропускаем временные файлы Excel
            If best = "" Then
                best = fName
                bestDate = FileDateTime(folderPath & "\" & fName)
            ElseIf FileDateTime(folderPath & "\" & fName) > bestDate Then
                best = fName
                bestDate = FileDateTime(folderPath & "\" & fName)
            End If
        End If
        fName = Dir()
    Loop
    LastModifiedFile = best
End Function

' Читает лист «Ошибки» книги сводной таблицы в текст
' (строки через vbCrLf, колонки через " | ").
' Возвращает False, если лист не найден.
Private Function ReadErrorsSheetText(ByVal wb As Workbook, ByRef outText As String) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Long
    Dim lineText As String
    Dim res As String

    On Error Resume Next
    Set ws = wb.Worksheets("Ошибки")
    On Error GoTo 0
    If ws Is Nothing Then
        outText = ""
        ReadErrorsSheetText = False
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1

    ' строка 1 — заголовки («Код», «Источник», «Сообщение»), строки 2+ — ошибки
    For r = 2 To lastRow
        lineText = ""
        For c = 1 To lastCol
            If c > 1 Then lineText = lineText & " | "
            lineText = lineText & Trim(CStr(ws.Cells(r, c).Value))
        Next c
        If Trim(lineText) <> "" Then res = res & lineText & vbCrLf
    Next r

    If Len(res) >= 2 Then
        If Right(res, 2) = vbCrLf Then res = Left(res, Len(res) - 2)
    End If
    outText = res
    ReadErrorsSheetText = True
End Function

' Собирает фрагмент схемы из листа «Плоская_таблица» книги, сохранённой
' в папке 02_Подготовка схемы (последняя по времени изменения книга).
' Если файла/листа нет — возвращает "" и пишет предупреждение в лог.
Private Function GetSchemaFragment() As String
    Dim prepPath As String
    Dim fileName As String
    Dim candidates(1 To 3) As String
    Dim newest As String
    Dim newestDate As Date
    Dim tmpDate As Date
    Dim i As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Long
    Dim lineText As String
    Dim res As String

    GetSchemaFragment = ""

    prepPath = Mod_Util.GetProjectRoot & "\02_Подготовка схемы"
    If Dir(prepPath, vbDirectory) = "" Then
        Mod_Util.WriteWarning "P002", "GetSchemaFragment: папка 02_Подготовка схемы не найдена, схема не включена."
        Exit Function
    End If

    ' ищем самый свежий файл результата подготовки схемы
    candidates(1) = "Плоская_таблица.xlsx"
    candidates(2) = "*.xlsx"
    candidates(3) = "*.xlsm"
    newest = ""
    For i = 1 To 3
        fileName = LastModifiedFile(prepPath, candidates(i))
        If fileName <> "" Then
            tmpDate = FileDateTime(prepPath & "\" & fileName)
            If newest = "" Or tmpDate > newestDate Then
                newest = fileName
                newestDate = tmpDate
            End If
        End If
    Next i

    If newest = "" Then
        Mod_Util.WriteWarning "P002", "GetSchemaFragment: в 02_Подготовка схемы нет файла Плоская_таблица.xlsx, схема не включена."
        Exit Function
    End If

    On Error Resume Next
    ' W-AV-REFACTOR, п.1: AddToMru:=False — файлы шагов не попадают в список
    ' последних файлов Excel (массовое добавление в MRU воспринимается AV как
    ' подозрительное сканирование).
    Set wb = Workbooks.Open(prepPath & "\" & newest, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Mod_Util.WriteWarning "P002", "GetSchemaFragment: не удалось открыть файл схемы '" & newest & "', схема не включена."
        Exit Function
    End If
    On Error GoTo 0

    On Error Resume Next
    Set ws = wb.Worksheets("Плоская_таблица")
    On Error GoTo 0
    If ws Is Nothing Then
        Mod_Util.WriteWarning "P002", "GetSchemaFragment: в файле '" & newest & "' нет листа 'Плоская_таблица', схема не включена."
        wb.Close SaveChanges:=False
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1

    For r = 1 To lastRow
        lineText = ""
        For c = 1 To lastCol
            If c > 1 Then lineText = lineText & " | "
            lineText = lineText & Trim(CStr(ws.Cells(r, c).Value))
        Next c
        If Trim(lineText) <> "" Then res = res & lineText & vbCrLf
    Next r

    If Len(res) >= 2 Then
        If Right(res, 2) = vbCrLf Then res = Left(res, Len(res) - 2)
    End If

    wb.Close SaveChanges:=False
    GetSchemaFragment = res
End Function

' ---------------------------------------------------------------------------
' ПУБЛИЧЕСКИЕ ПРОЦЕДУРЫ И ФУНКЦИИ
' ---------------------------------------------------------------------------

' Читает список шагов (<корень>\04_Промты\Шаги.txt) и возвращает его
' коллекцией. Каждая запись — строка вида "NN|Имя_папки".
' При отсутствии файла: ошибка P001, MsgBox, пустая коллекция.
Public Function ReadStepList() As Collection
    Dim col As Collection
    Dim fnum As Integer
    Dim s As String
    Dim posEq As Long
    Dim nn As String
    Dim nm As String

    Set col = New Collection

    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = Mod_Util.GetProjectRoot & "\04_Промты\Шаги.txt"

    ' файла нет ? P001, MsgBox, пустая коллекция
    If Dir(filePath) = "" Then
        Mod_Util.WriteError "P001", "ReadStepList: не найден Шаги.txt"
        MsgBox "Ошибка P001: не найден файл Шаги.txt (" & vbCrLf & filePath & ")", _
               vbCritical, "Промты"
        Set ReadStepList = col
        Exit Function
    End If

    ' формат строки: "NN = Имя_папки"
    fnum = FreeFile
    Open filePath For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, s
        s = Trim(s)
        If s <> "" Then
            posEq = InStr(s, "=")
            If posEq > 0 Then
                nn = Trim(Left(s, posEq - 1))
                nm = Trim(Mid(s, posEq + 1))
                If nn <> "" And nm <> "" Then col.Add nn & "|" & nm
            End If
        End If
    Loop
    Close #fnum

    Set ReadStepList = col
    Exit Function

ErrHandler:
    Mod_Util.WriteError "P001", "ReadStepList: ошибка чтения Шаги.txt. Ошибка: " & Err.Description
    MsgBox "Ошибка P001: не удалось прочитать файл Шаги.txt." & vbCrLf & Err.Description, _
           vbCritical, "Промты"
    Err.Clear
    Set ReadStepList = col
End Function

' Возвращает номер текущего шага — последней существующей папки Шаг_*
' в 03_Результаты генерации\01_По шагам\ согласно порядку из Шаги.txt.
' Если папок нет — возвращает "".
Public Function ReadCurrentStep() As String
    Dim steps As Collection
    Dim stepsPath As String
    Dim i As Long
    Dim nn As String
    Dim nm As String
    Dim found As String

    ReadCurrentStep = ""

    On Error GoTo ErrHandler

    Set steps = ReadStepList()
    If steps Is Nothing Then Exit Function
    If steps.Count = 0 Then Exit Function

    stepsPath = Mod_Util.GetProjectRoot & "\03_Результаты генерации\01_По шагам"
    If Dir(stepsPath, vbDirectory) = "" Then Exit Function

    ' идём по порядку Шаги.txt и запоминаем последнюю существующую папку
    found = ""
    For i = 1 To steps.Count
        nn = StepNumFromEntry(steps(i))
        nm = StepNameFromEntry(steps(i))
        If Dir(stepsPath & "\" & nm, vbDirectory) <> "" Then found = nn
    Next i

    ReadCurrentStep = found
    Exit Function

ErrHandler:
    Mod_Util.WriteError "P002", "ReadCurrentStep: ошибка определения текущего шага. Ошибка: " & Err.Description
    MsgBox "Ошибка P002: не удалось определить текущий шаг." & vbCrLf & Err.Description, _
           vbCritical, "Промты"
    Err.Clear
    ReadCurrentStep = ""
End Function

' Возвращает номер следующего шага (шаг из Шаги.txt, идущий сразу за
' текущим). Если текущего шага нет — первый шаг из Шаги.txt.
' Если следующий не найден — возвращает "".
Public Function ReadNextStep() As String
    Dim steps As Collection
    Dim current As String
    Dim i As Long

    ReadNextStep = ""

    On Error GoTo ErrHandler

    Set steps = ReadStepList()
    If steps Is Nothing Then Exit Function
    If steps.Count = 0 Then Exit Function

    current = ReadCurrentStep()

    ' текущего шага нет — возвращаем первый шаг из Шаги.txt
    If current = "" Then
        ReadNextStep = StepNumFromEntry(steps(1))
        Exit Function
    End If

    ' следующий — запись, идущая сразу за текущей
    For i = 1 To steps.Count - 1
        If StepNumFromEntry(steps(i)) = current Then
            ReadNextStep = StepNumFromEntry(steps(i + 1))
            Exit Function
        End If
    Next i

    ' следующий не найден
    ReadNextStep = ""
    Exit Function

ErrHandler:
    Mod_Util.WriteError "P002", "ReadNextStep: ошибка определения следующего шага. Ошибка: " & Err.Description
    MsgBox "Ошибка P002: не удалось определить следующий шаг." & vbCrLf & Err.Description, _
           vbCritical, "Промты"
    Err.Clear
    ReadNextStep = ""
End Function

' Формирует промпт для выполнения следующего шага (разделы 9 и 13).
Public Sub MakeNextStepPrompt()
    Dim root As String
    Dim steps As Collection
    Dim stepsPath As String
    Dim promptsPath As String
    Dim readyPath As String
    Dim registriesPath As String
    Dim current As String
    Dim nextNN As String
    Dim nextName As String
    Dim template As String
    Dim regTypes As String
    Dim regCoverage As String
    Dim schema As String
    Dim result As String
    Dim stepFolder As String
    Dim promptFile As String
    Dim i As Long
    Dim j As Long

    On Error GoTo ErrHandler

    root = NormPath(Mod_Util.GetProjectRoot)
    stepsPath = root & "\03_Результаты генерации\01_По шагам"
    promptsPath = root & "\04_Промты"
    readyPath = root & "\03_Результаты генерации\04_Промты_готовые"
    registriesPath = root & "\03_Результаты генерации\03_Реестры"

    ' 1. Список шагов (при отсутствии Шаги.txt ReadStepList сама
    '    сообщит ошибку P001 и вернёт пустую коллекцию)
    Set steps = ReadStepList()
    If steps Is Nothing Then Exit Sub
    If steps.Count = 0 Then Exit Sub

    ' 2. Текущий и следующий шаги
    current = ReadCurrentStep()
    nextNN = ReadNextStep()

    ' 3. Следующего шага нет ? финальный аудит (раздел 9)
    If nextNN = "" Then
        ' проверка: нет ли папки с номером больше текущего (неожиданная декомпозиция)
        If current <> "" Then
            Dim subName As String
            Dim folderNN As String
            subName = Dir(stepsPath & "\Шаг_*", vbDirectory)
            Do While subName <> ""
                If (GetAttr(stepsPath & "\" & subName) And vbDirectory) = vbDirectory Then
                    folderNN = Mid(subName, 6)          ' "Шаг_NN..." -> "NN..."
                    i = InStr(folderNN, "_")
                    If i > 0 Then folderNN = Left(folderNN, i - 1)
                    ' сверяем папку со списком Шаги.txt: если её там нет и
                    ' номер больше текущего — unexpected step (раздел 9)
                    Dim inList As Boolean
                    inList = False
                    For j = 1 To steps.Count
                        If StepNameFromEntry(steps(j)) = subName Then inList = True
                    Next j
                    If Not inList And StepNumToLong(folderNN) > StepNumToLong(current) Then
                        Mod_Util.WriteError "P002", "MakeNextStepPrompt: найдена папка '" & _
                            subName & "' с номером больше текущего, но в Шаги.txt нет " & _
                            "соответствующей строки. Обновите Шаги.txt."
                        MsgBox "Ошибка P002: в 01_По шагам есть папка '" & subName & "', " & _
                               "которой нет в Шаги.txt." & vbCrLf & _
                               "Добавьте строку для этого шага в Шаги.txt.", _
                               vbCritical, "Промт следующего шага"
                        Exit Sub
                    End If
                End If
                subName = Dir()
            Loop
        End If

        ' читаем шаблон того же уровня (prompt_step_template_v1.md)
        template = ReadTextFile(promptsPath & "\prompt_step_template_v1.md")
        If template = "" Then
            Mod_Util.WriteError "P002", "MakeNextStepPrompt: не найден шаблон prompt_step_template_v1.md"
            MsgBox "Ошибка P002: не найден шаблон prompt_step_template_v1.md в папке 04_Промты.", _
                   vbCritical, "Промт следующего шага"
            Exit Sub
        End If

        ' реестры и фрагмент схемы (при отсутствии — пустая строка и предупреждение)
        regTypes = ReadTextFile(registriesPath & "\Реестр_типов.txt")
        If regTypes = "" Then
            Mod_Util.WriteWarning "P002", "MakeNextStepPrompt: Реестр_типов.txt не найден или пуст, подставлена пустая строка."
        End If
        regCoverage = ReadTextFile(registriesPath & "\Реестр_покрытия.txt")
        If regCoverage = "" Then
            Mod_Util.WriteWarning "P002", "MakeNextStepPrompt: Реестр_покрытия.txt не найден или пуст, подставлена пустая строка."
        End If
        schema = GetSchemaFragment()

        ' финальный аудит: NN = 99, STEP_NAME = Шаг_99_Финал
        result = template
        result = ReplaceIfPresent(result, "{NN}", "99")
        result = ReplaceIfPresent(result, "{STEP_NAME}", "Шаг_99_Финал")
        result = ReplaceIfPresent(result, "{REGISTRY_TYPES}", regTypes)
        result = ReplaceIfPresent(result, "{REGISTRY_COVERAGE}", regCoverage)
        result = ReplaceIfPresent(result, "{SCHEMA}", schema)

        EnsureFolder stepsPath & "\Шаг_99_Финал"
        EnsureFolder readyPath
        SaveTextFile readyPath & "\Промт_финального_аудита.md", result

        Mod_Util.WriteLog "MakeNextStepPrompt: сформирован финальный аудит"
        MsgBox "Все шаги из Шаги.txt выполнены." & vbCrLf & _
               "Сформирован Промт_финального_аудита.md, создана папка Шаг_99_Финал.", _
               vbInformation, "Финальный аудит"
        Exit Sub
    End If

    ' 4. Имя папки следующего шага из Шаги.txt
    nextName = ""
    For i = 1 To steps.Count
        If StepNumFromEntry(steps(i)) = nextNN Then
            nextName = StepNameFromEntry(steps(i))
            Exit For
        End If
    Next i
    If nextName = "" Then
        Mod_Util.WriteError "P002", "MakeNextStepPrompt: в Шаги.txt не найдено имя папки для шага " & nextNN
        MsgBox "Ошибка P002: в Шаги.txt не найдено имя папки для шага " & nextNN & ".", _
               vbCritical, "Промт следующего шага"
        Exit Sub
    End If

    ' 5. Шаблон промта шага
    template = ReadTextFile(promptsPath & "\prompt_step_template_v1.md")
    If template = "" Then
        Mod_Util.WriteError "P002", "MakeNextStepPrompt: не найден шаблон prompt_step_template_v1.md"
        MsgBox "Ошибка P002: не найден шаблон prompt_step_template_v1.md в папке 04_Промты.", _
               vbCritical, "Промт следующего шага"
        Exit Sub
    End If

    ' 6. Реестры из 03_Реестры (при отсутствии — пустая строка и предупреждение)
    regTypes = ReadTextFile(registriesPath & "\Реестр_типов.txt")
    If regTypes = "" Then
        Mod_Util.WriteWarning "P002", "MakeNextStepPrompt: Реестр_типов.txt не найден или пуст, подставлена пустая строка."
    End If
    regCoverage = ReadTextFile(registriesPath & "\Реестр_покрытия.txt")
    If regCoverage = "" Then
        Mod_Util.WriteWarning "P002", "MakeNextStepPrompt: Реестр_покрытия.txt не найден или пуст, подставлена пустая строка."
    End If

    ' 7. Фрагмент схемы из 02_Подготовка схемы (Плоская_таблица.xlsx, если есть)
    schema = GetSchemaFragment()

    ' 8. Подстановка плейсхолдеров (только тех, что есть в шаблоне)
    result = template
    result = ReplaceIfPresent(result, "{NN}", nextNN)
    result = ReplaceIfPresent(result, "{STEP_NAME}", nextName)
    result = ReplaceIfPresent(result, "{REGISTRY_TYPES}", regTypes)
    result = ReplaceIfPresent(result, "{REGISTRY_COVERAGE}", regCoverage)
    result = ReplaceIfPresent(result, "{SCHEMA}", schema)

    ' 9. Папка шага создаётся только по факту формирования промта (раздел 3)
    stepFolder = stepsPath & "\" & nextName
    EnsureFolder stepFolder

    ' 10. Сохранение готового промта
    EnsureFolder readyPath
    promptFile = readyPath & "\Промт_шаг_" & nextNN & ".md"
    SaveTextFile promptFile, result

    Mod_Util.WriteLog "MakeNextStepPrompt: сформирован промт шага " & nextNN
    MsgBox "Сформирован промт шага " & nextNN & " (" & nextName & "):" & vbCrLf & _
           promptFile & vbCrLf & vbCrLf & _
           "Создана папка шага: " & stepFolder, vbInformation, "Промт следующего шага"
    Exit Sub

ErrHandler:
    Mod_Util.WriteError "P002", "MakeNextStepPrompt: ошибка формирования промта шага. Ошибка: " & Err.Description
    MsgBox "Ошибка P002: не удалось сформировать промт следующего шага." & vbCrLf & _
           Err.Description, vbCritical, "Промт следующего шага"
    Err.Clear
End Sub

' Формирует промпт для исправления ошибок текущего шага (раздел 13).
Public Sub MakeFixPrompt()
    Dim root As String
    Dim current As String
    Dim summaryPath As String
    Dim finalFile As String
    Dim promptsPath As String
    Dim readyPath As String
    Dim registriesPath As String
    Dim template As String
    Dim regTypes As String
    Dim regCoverage As String
    Dim schema As String
    Dim errorsText As String
    Dim result As String
    Dim wb As Workbook
    Dim fileName As String

    On Error GoTo ErrHandler

    root = NormPath(Mod_Util.GetProjectRoot)
    promptsPath = root & "\04_Промты"
    readyPath = root & "\03_Результаты генерации\04_Промты_готовые"
    registriesPath = root & "\03_Результаты генерации\03_Реестры"
    summaryPath = root & "\03_Результаты генерации\02_Сводные"

    ' 1. Текущий шаг
    current = ReadCurrentStep()
    If current = "" Then
        Mod_Util.WriteError "P003", "MakeFixPrompt: текущий шаг не определён (нет папок Шаг_* в 01_По шагам)"
        MsgBox "Ошибка P003: текущий шаг не определён — в 03_Результаты генерации\01_По шагам " & _
               "нет ни одной папки шага.", vbCritical, "Промт исправления"
        Exit Sub
    End If

    ' 2. Последняя сводная из 02_Сводные
    If Dir(summaryPath, vbDirectory) = "" Then
        Mod_Util.WriteError "P003", "MakeFixPrompt: папка 02_Сводные не найдена: " & summaryPath
        MsgBox "Ошибка P003: папка 02_Сводные не найдена." & vbCrLf & summaryPath, _
               vbCritical, "Промт исправления"
        Exit Sub
    End If

    ' фиксированное имя последней сводной; если его нет — последний xlsx по дате
    fileName = Dir(summaryPath & "\zhc_okks_gost_FINAL.xlsx")
    If fileName = "" Then
        fileName = LastModifiedFile(summaryPath, "*.xlsx")
    End If
    If fileName = "" Then
        Mod_Util.WriteError "P003", "MakeFixPrompt: в 02_Сводные нет файла zhc_okks_gost_FINAL.xlsx"
        MsgBox "Ошибка P003: в папке 02_Сводные не найден файл сводной " & _
               "zhc_okks_gost_FINAL.xlsx.", vbCritical, "Промт исправления"
        Exit Sub
    End If
    finalFile = summaryPath & "\" & fileName

    Set wb = Nothing
    On Error Resume Next
    ' W-AV-REFACTOR, п.1: AddToMru:=False — см. комментарий выше (Mod_Assembler).
    Set wb = Workbooks.Open(finalFile, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    On Error GoTo ErrHandler
    If wb Is Nothing Then
        Mod_Util.WriteError "P003", "MakeFixPrompt: не удалось открыть сводную: " & finalFile
        MsgBox "Ошибка P003: не удалось открыть файл сводной:" & vbCrLf & finalFile, _
               vbCritical, "Промт исправления"
        Exit Sub
    End If

    ' 3. Лист «Ошибки»
    errorsText = ""
    If Not ReadErrorsSheetText(wb, errorsText) Then
        Mod_Util.WriteError "P003", "MakeFixPrompt: в сводной не найден лист 'Ошибки': " & finalFile
        wb.Close SaveChanges:=False
        MsgBox "Ошибка P003: в файле сводной не найден лист 'Ошибки'." & vbCrLf & finalFile, _
               vbCritical, "Промт исправления"
        Exit Sub
    End If
    wb.Close SaveChanges:=False

    ' 4. Ошибок нет — ничего не формируем (раздел 13)
    If errorsText = "" Then
        Mod_Util.WriteLog "MakeFixPrompt: ошибок нет"
        MsgBox "Ошибок нет, исправлять нечего.", vbInformation, "Промт исправления"
        Exit Sub
    End If

    ' 5. Шаблон промта исправления
    template = ReadTextFile(promptsPath & "\prompt_fix_v1.0.md")
    If template = "" Then
        Mod_Util.WriteError "P003", "MakeFixPrompt: не найден шаблон prompt_fix_v1.0.md"
        MsgBox "Ошибка P003: не найден шаблон prompt_fix_v1.0.md в папке 04_Промты.", _
               vbCritical, "Промт исправления"
        Exit Sub
    End If

    ' 6. Оба реестра (при отсутствии — пустая строка и предупреждение)
    regTypes = ReadTextFile(registriesPath & "\Реестр_типов.txt")
    If regTypes = "" Then
        Mod_Util.WriteWarning "P003", "MakeFixPrompt: Реестр_типов.txt не найден или пуст, подставлена пустая строка."
    End If
    regCoverage = ReadTextFile(registriesPath & "\Реестр_покрытия.txt")
    If regCoverage = "" Then
        Mod_Util.WriteWarning "P003", "MakeFixPrompt: Реестр_покрытия.txt не найден или пуст, подставлена пустая строка."
    End If

    ' 7. Фрагмент схемы текущего шага
    schema = GetSchemaFragment()

    ' 8. Подстановка плейсхолдеров (только тех, что есть в шаблоне)
    result = template
    result = ReplaceIfPresent(result, "{NN}", current)
    result = ReplaceIfPresent(result, "{ERRORS}", errorsText)
    result = ReplaceIfPresent(result, "{REGISTRY_TYPES}", regTypes)
    result = ReplaceIfPresent(result, "{REGISTRY_COVERAGE}", regCoverage)
    result = ReplaceIfPresent(result, "{SCHEMA}", schema)

    ' 9. Сохранение готового промта
    EnsureFolder readyPath
    SaveTextFile readyPath & "\Промт_исправление_шаг_" & current & ".md", result

    Mod_Util.WriteLog "MakeFixPrompt: сформирован промт исправления шага " & current
    MsgBox "Сформирован промт исправления шага " & current & ":" & vbCrLf & _
           readyPath & "\Промт_исправление_шаг_" & current & ".md", _
           vbInformation, "Промт исправления"
    Exit Sub

ErrHandler:
    Mod_Util.WriteError "P003", "MakeFixPrompt: ошибка формирования промта исправления. Ошибка: " & Err.Description
    MsgBox "Ошибка P003: не удалось сформировать промт исправления." & vbCrLf & _
           Err.Description, vbCritical, "Промт исправления"
    Err.Clear
End Sub
