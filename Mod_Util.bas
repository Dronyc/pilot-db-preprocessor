Attribute VB_Name = "Mod_Util"
Option Explicit

' ============================================================================
' Модуль утилит: пути проекта, журналирование, чтение конфигурации.
' Реализация W4 (разделы 2, 5 спецификации Spec_Dispatcher.md).
' Пути строятся как в Mod_Registry / Mod_Prompt: <корень>\NN_Имя_папки,
' корень берётся из поля ProjectRoot листа "Настройки" (T3.1); если поле
' пустое - используется папка файла .xlsm (ThisWorkbook.Path) как значение
' по умолчанию (обратная совместимость).
' Реестр кодов ошибок/предупреждений модуля (OPT-06, единый реестр):
'   U002 - ошибка чтения configuration.txt;
'   U003 - ошибка подготовки листа "Настройки";
'   U004 - ошибка выбора папки проекта;
'   U005 - ошибка открытия папки в проводнике;
'   U006 - предупреждение: в папке найден только .pilotcfx;
'   U007 - предупреждение (T3.1): числовая настройка пуста/некорректна -
'          использовано значение по умолчанию.
' ============================================================================

' ---------------------------------------------------------------------------
' КОНСТАНТЫ И ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (Private)
' ---------------------------------------------------------------------------

' Имена именованных диапазонов листа "Настройки" (раздел 5 спецификации).
' T3.1: добавлены ProjectRoot (B2) и MaxTypesPerTable (B7).
Private Const NAME_PROJECT_ROOT As String = "ProjectRoot"
Private Const NAME_CHECK_MODE As String = "CheckMode"
Private Const NAME_AUTO_OPEN_RESULTS As String = "AutoOpenResults"
Private Const NAME_PROMPT_GEN_VERSION As String = "PromptGenVersion"
Private Const NAME_PROMPT_FIX_VERSION As String = "PromptFixVersion"
Private Const NAME_MAX_TYPES_PER_TABLE As String = "MaxTypesPerTable"

' Значение по умолчанию для MaxTypesPerTable (T3.1): используется
' Mod_Prompt при подстановке {ЛИМИТ_ТИПОВ_В_ТАБЛИЦЕ} и как default
' для GetSettingLong.
Public Const DEFAULT_MAX_TYPES_PER_TABLE As Long = 50

' Пароль защиты служебных листов (единый для книги Препроцессора).
Private Const SHEET_PASSWORD As String = "pilot"

' Возвращает путь без завершающего разделителя "\".
Private Function NormPath(ByVal p As String) As String
    Do While Len(p) > 1 And Right(p, 1) = "\"
        p = Left(p, Len(p) - 1)
    Loop
    NormPath = p
End Function

' ---------------------------------------------------------------------------
' ПУБЛИЧЕСКИЕ ПРОЦЕДУРЫ И ФУНКЦИИ
' ---------------------------------------------------------------------------

' Возвращает корневую папку проекта (T3.1, H2 исправлен):
' 1. читает значение поля ProjectRoot (B2) листа "Настройки" - через
'    именованный диапазон ProjectRoot, либо напрямую по ячейке B2, если
'    имя ещё не создано;
' 2. если поле пустое (или лист "Настройки" отсутствует) - возвращает
'    папку файла .xlsm (ThisWorkbook.Path) - обратная совместимость;
' 3. если книга не сохранена - возвращает "".
' Значение нормализуется (снимаются завершающие "\").
Public Function GetProjectRoot() As String
    Dim v As String
    Dim ws As Worksheet
    Dim wb As Workbook

    Set wb = ThisWorkbook

    ' 1. ProjectRoot из листа "Настройки" (именованный диапазон / ячейка B2).
    ' T3.1-fix: читаем фактическое значение диапазона (nm.RefersToRange.Value),
    ' а не строку определения имени (nm.Value).
    On Error Resume Next
    Dim nm As Name
    Set nm = wb.Names(NAME_PROJECT_ROOT)
    If Err.Number = 0 Then
        v = CStr(nm.RefersToRange.Value)
        If Err.Number <> 0 Then v = "": Err.Clear
    Else
        Err.Clear
        Set ws = wb.Worksheets("Настройки")
        If Err.Number = 0 Then v = CStr(ws.Range("B2").Value)
        Err.Clear
    End If
    On Error GoTo 0

    v = Trim$(v)
    If v <> "" Then
        GetProjectRoot = NormPath(v)
        Exit Function
    End If

    ' 2. Поле пусто - папка файла книги (обратная совместимость).
    If ThisWorkbook.path <> "" Then
        GetProjectRoot = NormPath(ThisWorkbook.path)
    Else
        GetProjectRoot = ""
    End If
End Function

' Читает числовую настройку листа "Настройки" по имени именованного
' диапазона (T3.1). Поведение:
'   * значение корректное целое (>= 1) -> возвращается оно;
'   * поле пустое / значение не является целым >= 1 -> возвращается
'     defaultValue + WriteWarning U007 (штатная ситуация, не ошибка);
'   * имя/лист не найдены -> возвращается defaultValue без предупреждения
'     (корректное состояние до первого запуска EnsureSettingsSheet).
' Используется для MaxTypesPerTable (Mod_Prompt: {ЛИМИТ_ТИПОВ_В_ТАБЛИЦЕ}).
Public Function GetSettingLong(ByVal name As String, ByVal defaultValue As Long) As Long
    Dim raw As String
    Dim pos As Long
    Dim parsed As Double
    Dim sNum As String

    GetSettingLong = defaultValue

    On Error GoTo ErrHandler

    ' T3.1-fix: читаем фактическое значение диапазона (RefersToRange.Value),
    ' а не строку определения имени (Names(...).Value); при ошибке/недоступности
    ' raw = "" - далее штатный fallback на defaultValue.
    On Error Resume Next
    raw = Trim$(CStr(ThisWorkbook.Names(name).RefersToRange.Value))
    If Err.Number <> 0 Then raw = "": Err.Clear
    On Error GoTo 0
    If raw = "" Then
        WriteWarning "U007", "GetSettingLong: настройка """ & name & """ пуста - " & _
                   "использовано значение по умолчанию " & defaultValue & "."
        Exit Function
    End If

    ' Разделитель дробной части зависит от локали - пробуем оба варианта.
    sNum = raw
    pos = InStr(sNum, ",")
    If pos > 0 Then Mid$(sNum, pos, 1) = "."

    If Not IsNumeric(sNum) Then
        WriteWarning "U007", "GetSettingLong: настройка """ & name & """=""" & raw & _
                   """ не является числом - использовано значение по умолчанию " & _
                   defaultValue & "."
        Exit Function
    End If

    parsed = CDbl(sNum)
    If parsed < 1 Or parsed <> Int(parsed) Then
        WriteWarning "U007", "GetSettingLong: настройка """ & name & """=""" & raw & _
               """ должно быть целым >= 1 - использовано значение по умолчанию " & _
               defaultValue & "."
        Exit Function
    End If

    GetSettingLong = CLng(parsed)
    Exit Function

ErrHandler:
    ' Имя не найдено / лист "Настройки" отсутствует - тихо отдаём default.
    GetSettingLong = defaultValue
    Err.Clear
End Function
' Возвращает путь к папке логов: <корень>\05_Логи.
Public Function GetLogsFolder() As String
    Dim root As String
    root = GetProjectRoot()
    If root = "" Then
        GetLogsFolder = ""
    Else
        GetLogsFolder = root & "\05_Логи"
    End If
End Function

' Возвращает путь к папке результатов: <корень>\03_Результаты генерации.
Public Function GetResultsFolder() As String
    Dim root As String
    root = GetProjectRoot()
    If root = "" Then
        GetResultsFolder = ""
    Else
        GetResultsFolder = root & "\03_Результаты генерации"
    End If
End Function

' Возвращает путь к папке промптов: <корень>\04_Промты.
Public Function GetPromptsFolder() As String
    Dim root As String
    root = GetProjectRoot()
    If root = "" Then
        GetPromptsFolder = ""
    Else
        GetPromptsFolder = root & "\04_Промты"
    End If
End Function

' Возвращает путь к папке шагов: <корень>\03_Результаты генерации\01_По шагам.
Public Function GetStepsFolder() As String
    Dim root As String
    root = GetProjectRoot()
    If root = "" Then
        GetStepsFolder = ""
    Else
        GetStepsFolder = root & "\03_Результаты генерации\01_По шагам"
    End If
End Function

' ---------------------------------------------------------------------------
' W-CORRECTION-PACK-1, п.1.1: УНИФИЦИРОВАННЫЙ ФОРМАТ СООБЩЕНИЙ ЖУРНАЛА
'   ERROR:  "ERROR [<код>]: <текст>"   (задаётся внутри WriteError)
'   WARN:   "[!] <текст>"              (задаётся внутри WriteWarning)
'   INFO:   "[i] <текст>"              (добавляется внутри WriteLog, если
'                                        текст ещё не содержит маркера)
' Вызывающий код передаёт чистый текст; маркер формата подставляется здесь,
' поэтому все записи журнала выглядят одинаково независимо от модуля.
' Допускается явная префиксация текста ("[i] ...", "[!] ...", "ERROR ...") —
' в этом случае дублирование маркера исключается.
' ---------------------------------------------------------------------------

' Возвращает true, если текст уже начинается с одного из унифицированных
' маркеров формата ("[i]", "[!]", "ERROR").
Private Function HasUnifiedPrefix(ByVal message As String) As Boolean
    Dim m As String
    m = Trim$(message)
    HasUnifiedPrefix = (Left$(m, 3) = "[i]" Or Left$(m, 3) = "[!]" Or _
                        Left$(m, 5) = "ERROR")
End Function

' Возвращает стандартное сообщение об отсутствующем обязательном файле
' (W-CORRECTION-PACK-1, п.1.1): "[!] Отсутствует обязательный файл: <FileName>".
Public Function FormatMissingFile(ByVal FileName As String) As String
    FormatMissingFile = "[!] Отсутствует обязательный файл: " & FileName
End Function

' Записывает сообщение в журнал <корень>\05_Логи\Препроцессор_YYYY-MM-DD.log.
' Формат строки: [YYYY-MM-DD HH:NN:SS] <message>, где <message> для
' информационных сообщений имеет вид "[i] <текст>" (W-CORRECTION-PACK-1,
' п.1.1). Если текст уже помечен маркером "[i]"/"[!]"/"ERROR" (сообщения
' WriteError/WriteWarning), повторная пометка НЕ добавляется.
' ВАЖНО (T3.2-fix): локальные обработчики ошибок применяются ТОЛЬКО внутри
' самой WriteLog (On Error Resume Next / On Error GoTo 0) и НИКОГДА не
' вызывают другие процедуры журналирования, поэтому рекурсия/зацикливание
' исключены: при любом сбое (MkDir, Open, Print) WriteLog молча выходит.
' Без такой защиты MkDir (ошибки 75/76 - нет прав или путь недоступен)
' ронял всю вызвавшую команду. Если корень не задан или запись невозможна -
' молча выходит.
Public Sub WriteLog(ByVal message As String)
    Dim root As String
    Dim logsPath As String
    Dim logFile As String
    Dim fnum As Integer

    root = GetProjectRoot()
    If root = "" Then Exit Sub          ' корень не задан - молча выходим

    logsPath = root & "\05_Логи"
    If Dir(logsPath, vbDirectory) = "" Then
        ' T3.2-fix: MkDir может упасть (права/недоступный путь) - защищаем
        ' локально, чтобы не ронять вызвавшую процедуру.
        On Error Resume Next
        MkDir logsPath
        If Err.Number <> 0 Then
            ' Папку создать не удалось - выходим без записи в лог,
            ' но НЕ роняем вызвавшую процедуру.
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
        ' T3.2-fix: убеждаемся, что папка действительно создалась.
        If Dir(logsPath, vbDirectory) = "" Then Exit Sub
    End If

    logFile = logsPath & "\Препроцессор_" & Format(Date, "yyyy-mm-dd") & ".log"

    ' T3.2-fix: блок работы с файлом тоже защищён от ошибок (файл занят,
    ' нет прав и т.п.) - сбой записи не должен ронять основную команду.
    On Error Resume Next
    fnum = FreeFile
    Open logFile For Append As #fnum
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub                        ' файл не открылся - молча выходим
    End If
    ' W-CORRECTION-PACK-1, п.1.1: информационные сообщения помечаются "[i]";
    ' сообщения ERROR/WARN приходят уже с маркером (см. WriteError/WriteWarning).
    If HasUnifiedPrefix(message) Then
        Print #fnum, "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] " & message
    Else
        Print #fnum, "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] [i] " & message
    End If
    Close #fnum
    Err.Clear
    On Error GoTo 0
End Sub

' Записывает ошибку в журнал в унифицированном формате (W-CORRECTION-PACK-1,
' п.1.1): "ERROR [<код>]: <текст>".
Public Sub WriteError(ByVal code As String, ByVal message As String)
    WriteLog "ERROR [" & code & "]: " & message
End Sub

' Записывает предупреждение в журнал в унифицированном формате
' (W-CORRECTION-PACK-1, п.1.1): "[!] <текст>". Код сохраняется в тексте
' сообщения для трассировки источника предупреждения.
Public Sub WriteWarning(ByVal code As String, ByVal message As String)
    WriteLog "[!] (" & code & ") " & message
End Sub

' W-DEBUG-ANOMALIES-EMPTY: отладочное логирование цепочки подготовки листа
' "Аномалии схемы" (BuildFlatStructure -> WriteSchemaDuplicates -> WriteLog).
' Сообщения с префиксом "[DEBUG] ":
'   * пишутся в Immediate (Debug.Print) — всегда;
'   * передаются в существующую файловую WriteLog этого модуля (журнал
'     <корень>\05_Логи\Препроцессор_YYYY-MM-DD.log) — только если корень
'     проекта задан (иначе WriteLog молча выходит и запись теряется);
'   * на лист "Аномалии схемы" НЕ пишутся.
' Рекурсия исключена: DebugLog вызывает WriteLog напрямую, WriteLog вызов
' DebugLog не выполняет. Ошибки журналирования не должны ронять сборку —
' собственная ошибка фиксируется в Immediate и подавляется.
Public Sub DebugLog(ByVal msg As String)
    Dim s As String
    s = "[DEBUG] " & msg

    ' 1. Immediate — доступно даже без настроенного корня проекта.
    Debug.Print s

    ' 2. Файловый журнал через существующий механизм WriteLog (без рекурсии).
    On Error GoTo ErrHandler
    If GetProjectRoot() <> "" Then WriteLog s
    Exit Sub

ErrHandler:
    Debug.Print "[DEBUG] DebugLog: ошибка записи в файл журнала: " & Err.Description
    Err.Clear
End Sub

' Возвращает значение параметра конфигурации по ключу из файла
' <корень>\01_Исходные данные\03_Конфигурация\configuration.txt.
' Формат строки: "key=value" (без пробелов вокруг "="), регистр ключа значим.
' Если корень/файл/ключ не найдены - возвращает "".
Public Function GetConfigValue(ByVal key As String) As String
    Dim root As String
    Dim cfgPath As String
    Dim fnum As Integer
    Dim s As String
    Dim posEq As Long

    GetConfigValue = ""

    On Error GoTo ErrHandler

    root = GetProjectRoot()
    If root = "" Then Exit Function     ' корень не задан - тихо возвращаем ""

    cfgPath = root & "\01_Исходные данные\03_Конфигурация\configuration.txt"
    If Dir(cfgPath) = "" Then Exit Function   ' файла нет - возвращаем ""

    fnum = FreeFile
    Open cfgPath For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, s
        posEq = InStr(s, "=")
        If posEq > 0 Then
            ' сравнение ключа с учётом регистра ("key=value", без пробелов)
            If StrComp(Left(s, posEq - 1), key, vbBinaryCompare) = 0 Then
                GetConfigValue = Mid(s, posEq + 1)
                Close #fnum
                Exit Function
            End If
        End If
    Loop
    Close #fnum

    ' ключ не найден - GetConfigValue уже равно ""
    Exit Function

ErrHandler:
    WriteError "U002", "GetConfigValue: ошибка чтения configuration.txt. Ошибка: " & Err.Description
    GetConfigValue = ""
    Err.Clear
End Function

' ---------------------------------------------------------------------------
' W-CORRECTION-PACK-1, п.1.2: поиск файла конфигурации в папке FolderPath.
' Формат конфигурации: принимается .xml (приоритет) или .txt; .pilotcfx ИИ не
' принимается. Правила:
'   * найден .xml  - вернуть полный путь к нему;
'   * иначе найден .txt - вернуть полный путь к нему;
'   * иначе в папке есть хотя бы один .pilotcfx (и нет .xml/.txt) - вернуть ""
'     и записать предупреждение "[!] Найден только .pilotcfx, требуется .xml или .txt";
'   * иначе (файлов конфигурации нет вовсе) - вернуть "".
' Временные файлы Office (~$) пропускаются. Если папка не существует - "".
' ---------------------------------------------------------------------------
Public Function FindConfigFile(ByVal FolderPath As String) As String
    Dim fName As String
    Dim xmlFound As String
    Dim txtFound As String
    Dim pilotcfxFound As Boolean

    FindConfigFile = ""

    If Dir(FolderPath, vbDirectory) = "" Then Exit Function

    ' проход по всем файлам папки: фиксируем первый .xml, первый .txt
    ' и наличие хотя бы одного .pilotcfx. Dir без атрибутов возвращает
    ' только файлы (не подпапки).
    fName = Dir(FolderPath & "\*")
    Do While fName <> ""
        If Left$(fName, 2) <> "~$" Then           ' пропускаем временные файлы Office
            If Right$(LCase$(fName), 4) = ".xml" Then
                If xmlFound = "" Then xmlFound = FolderPath & "\" & fName
            ElseIf Right$(LCase$(fName), 4) = ".txt" Then
                If txtFound = "" Then txtFound = FolderPath & "\" & fName
            ElseIf Right$(LCase$(fName), 8) = ".pilotcfx" Then
                pilotcfxFound = True
            End If
        End If
        fName = Dir()
    Loop

    ' приоритет: .xml > .txt
    If xmlFound <> "" Then
        FindConfigFile = xmlFound
        Exit Function
    End If
    If txtFound <> "" Then
        FindConfigFile = txtFound
        Exit Function
    End If

    ' только .pilotcfx без .xml/.txt - предупреждение, пустой результат
    If pilotcfxFound Then
        WriteWarning "U006", "Найден только .pilotcfx, требуется .xml или .txt (" & FolderPath & ")"
        FindConfigFile = ""
        Exit Function
    End If

    FindConfigFile = ""
End Function

' Открывает указанную папку в проводнике Windows.
' Если путь пустой или папки нет - молча выходит.
' W-AV-REFACTOR, п.2: использование Application.FollowHyperlink вместо Shell
' для предотвращения ложных срабатываний AV (запуск процессов через Shell —
' типичный эвристический признак "опасного" макроса). FollowHyperlink —
' нативное средство Excel: UI не ломается, фокус корректно передаётся
' окну Проводника.
Public Sub OpenFolderInExplorer(ByVal path As String)
    If path = "" Then Exit Sub
    If Dir(path, vbDirectory) = "" Then Exit Sub
    On Error Resume Next
    Application.FollowHyperlink path
    ' Если переход по гиперссылке невозможен (например, сетевой путь
    ' недоступен) — молча выходим: ошибка не должна прерывать UI.
    If Err.Number <> 0 Then
        WriteError "U005", "OpenFolderInExplorer: не удалось открыть папку " & path & _
                           ". Ошибка: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' W-AV-REFACTOR, п.4: безопасное удаление файла.
' Единая точка удаления файлов проекта: проверяет существование через Dir(),
' выполняет Kill под защитой обработки ошибок и возвращает результат.
' Прямые вызовы Kill из других модулей заменены на эту функцию, чтобы
' снизить эвристический скор антивирусов на "массовое удаление файлов".
' Возвращает True, если файла уже нет или он успешно удалён;
' False — если удалить не удалось (файл занят, нет прав и т.п.).
' ---------------------------------------------------------------------------
Public Function DeleteFileIfExists(ByVal filePath As String) As Boolean
    DeleteFileIfExists = True
    If Len(Dir(filePath)) = 0 Then Exit Function
    On Error Resume Next
    Kill filePath
    If Err.Number <> 0 Then
        DeleteFileIfExists = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

' Создаёт лист "Настройки" (если его нет), заполняет поля, выпадающие
' списки и именованные диапазоны по разделу 5 спецификации
' Spec_Dispatcher.md (T3.1 - шесть полей B2:B7):
'   B2 ProjectRoot       - Корневая папка проекта (текст; пусто = ThisWorkbook.Path);
'   B3 CheckMode         - Режим проверки (список: Черновик / Строгая; по умолчанию Черновик);
'   B4 AutoOpenResults   - Автооткрытие папки результатов (список: да / нет; по умолчанию да);
'   B5 PromptGenVersion  - Версия промта-генератора (текст; по умолчанию 7.13);
'   B6 PromptFixVersion  - Версия промта-исправления (текст; по умолчанию 1.0);
'   B7 MaxTypesPerTable  - Лимит типов в таблице (целое >= 1; по умолчанию 50).
' Validation навешивается СТРОГО на свою ячейку (H1 исправлен):
' CheckMode -> B3, AutoOpenResults -> B4, MaxTypesPerTable -> B7;
' B2/B5/B6 - текстовые поля без списков.
' Именованные диапазоны создаются для ВСЕХ шести полей; существующие имена
' не переопределяются. Лист защищается паролем "pilot", редактируемыми
' остаются только B2:B7. Существующие значения пользовательских настроек
' при повторном запуске не перезаписываются (значения по умолчанию
' проставляются только в пустые ячейки новых/существующих полей).
Public Sub EnsureSettingsSheet()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim RowIdx As Long
    Dim labels As Variant
    Dim namesArr As Variant
    Dim defaults As Variant
    Dim i As Long

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook

    ' 1. Поиск листа "Настройки"; если нет - создать в конце книги
    On Error Resume Next
    Set ws = wb.Worksheets("Настройки")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ErrHandler
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "Настройки"
    Else
        On Error GoTo ErrHandler
        ws.Unprotect Password:=SHEET_PASSWORD   ' снимаем защиту перед правкой
    End If

    ' 2. Подписи (колонка A), имена диапазонов и значения по умолчанию -
    '    шесть полей B2:B7 (T3.1)
    labels = Array("Корневая папка проекта", _
               "Режим проверки", _
               "Автооткрытие папки результатов", _
               "Версия промта-генератора", _
               "Версия промта-исправления", _
               "Максимум типов в таблице")
    namesArr = Array(NAME_PROJECT_ROOT, NAME_CHECK_MODE, NAME_AUTO_OPEN_RESULTS, _
                     NAME_PROMPT_GEN_VERSION, NAME_PROMPT_FIX_VERSION, _
                     NAME_MAX_TYPES_PER_TABLE)
    ' ProjectRoot остаётся пустым: пусто = ThisWorkbook.Path (см. GetProjectRoot).
    defaults = Array("", "Черновик", "да", "7.13", "1.0", CStr(DEFAULT_MAX_TYPES_PER_TABLE))

    For i = 0 To 5
        RowIdx = i + 2
        ws.Cells(RowIdx, 1).value = CStr(labels(i))
        ' значение по умолчанию ставим только если ячейка пуста -
        ' пользовательские значения при повторном запуске НЕ затираются
        If Trim(CStr(ws.Cells(RowIdx, 2).value)) = "" Then
            If CStr(defaults(i)) <> "" Then
                ws.Cells(RowIdx, 2).value = CStr(defaults(i))
            End If
        End If
    Next i

    ws.Range("A1").value = "Настройки Препроцессора"
    ws.Columns("A:B").AutoFit

    ' 3. Validation строго на своих ячейках (H1 исправлен):
    '    B3 CheckMode - список; B4 AutoOpenResults - список;
    '    B7 MaxTypesPerTable - целое >= 1; B2/B5/B6 - без списков
    '    (старое смещённое правило на B3 пересоздаётся ниже корректно,
    '     правило на B2 снимается).
    ws.Range("B2").Validation.Delete
    With ws.Range("B3").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="Черновик,Строгая"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    With ws.Range("B4").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="да,нет"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    With ws.Range("B7").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
             Operator:=xlGreaterEqual, Formula1:="1"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Лимит типов в таблице"
        .ErrorTitle = "Некорректное значение"
        .InputMessage = "Целое число >= 1 (по умолчанию 50)."
        .ErrorMessage = "MaxTypesPerTable должно быть целым числом >= 1."
    End With

    ' 4. Именованные диапазоны для всех шести полей - создавать, если их
    '    ещё нет (существующие имена не трогаем, чтобы не переопределить
    '    ссылку)
    For i = 0 To 5
        On Error Resume Next
        Dim nmChk As Name
        Set nmChk = wb.Names(namesArr(i))
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo ErrHandler
            wb.Names.Add Name:=namesArr(i), _
                         RefersTo:="=" & ws.Name & "!$B$" & (i + 2)
        End If
        On Error GoTo ErrHandler
        Set nmChk = Nothing
    Next i

    ' 5. Защита листа от случайной правки: разблокированы B2:B7
    ws.Cells.Locked = True
    ws.Range("B2:B7").Locked = False
    ws.Protect Password:=SHEET_PASSWORD

    WriteLog "EnsureSettingsSheet: лист 'Настройки' создан/обновлён"

    ' 6. T3.1 (H4): фиксируем в журнале статус полей, которые пока не
    '    используются кодом (это не ошибка - информационная запись).
    WriteLog "EnsureSettingsSheet: поля CheckMode/AutoOpenResults/PromptGenVersion/" & _
             "PromptFixVersion созданы; чтение этих настроек кодом пока не " & _
             "реализовано (зарезервировано на след. задачи). MaxTypesPerTable " & _
             "читается через GetSettingLong (Mod_Prompt: {ЛИМИТ_ТИПОВ_В_ТАБЛИЦЕ})."
    Exit Sub

ErrHandler:
    WriteError "U003", "EnsureSettingsSheet: ошибка подготовки листа 'Настройки'. Ошибка: " & Err.Description
    MsgBox "Ошибка U003: не удалось подготовить лист 'Настройки'." & vbCrLf & Err.Description, _
           vbCritical, "Настройки"
    Err.Clear
End Sub
