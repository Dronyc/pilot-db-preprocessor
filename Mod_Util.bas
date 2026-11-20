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
'          использовано значение по умолчанию; с PRE-04 распространяется
'          и на текстовые настройки (GetSettingText, напр. LegendRange);
'   U008 - предупреждение (PRE-04c): не удалось подготовить лист
'          "Паттерны" из Mod_Util.EnsureSettingsSheet (вызов
'          Mod_Patterns.EnsurePatternsSheet (коды модуля - PT###);
'          подготовка книги продолжается.
' ============================================================================

' ---------------------------------------------------------------------------
' КОНСТАНТЫ И ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (Private)
' ---------------------------------------------------------------------------

' Имена именованных диапазонов листа "Настройки" (раздел 5 спецификации).
' T3.1: добавлены ProjectRoot (B2) и MaxTypesPerTable (B7).
' PRE-04: добавлен LegendRange (B8) - диапазон легенды цветов на листе схемы.
Private Const NAME_PROJECT_ROOT As String = "ProjectRoot"
Private Const NAME_CHECK_MODE As String = "CheckMode"
Private Const NAME_AUTO_OPEN_RESULTS As String = "AutoOpenResults"
Private Const NAME_PROMPT_GEN_VERSION As String = "PromptGenVersion"
Private Const NAME_PROMPT_FIX_VERSION As String = "PromptFixVersion"
Private Const NAME_MAX_TYPES_PER_TABLE As String = "MaxTypesPerTable"
Private Const NAME_LEGEND_RANGE As String = "LegendRange"

' PRE-04: значение по умолчанию для LegendRange (адрес диапазона легенды
' цветов на первом листе книги-схемы): используется как default для
' GetSettingText("LegendRange", DEFAULT_LEGEND_RANGE) в Mod_Structure.
Public Const DEFAULT_LEGEND_RANGE As String = "B1:C3"

' W-FLEX-STRUCTURE: новые настройки для гибкого выбора листов и начала структуры
Private Const NAME_LEGEND_SHEET As String = "LegendSheetName"
Private Const NAME_DATA_SHEET As String = "DataSheetName"
Private Const NAME_START_CELL As String = "StartCell"
Public Const DEFAULT_START_CELL As String = "A8" ' Соответствует старому FIRST_ROW=8, Col=1
Private Const NAME_STRUCTURE_SHEET As String = "StructureSheet"

' W-ANALYTIC-08: имена именованных диапазонов трёх параметров состояния
' (раздел 5 спецификации Spec_Dispatcher.md): "Выжимки и анализ конфигурации",
' "Разбивка на паттерны", "Разбивка на шаги". Значения по умолчанию - "Нет";
' значения ставятся только в пустые ячейки (идемпотентность).
Private Const NAME_ANALYTICS_STATE As String = "AnalyticsState"
Private Const NAME_PATTERNS_STATE As String = "PatternsState"
Private Const NAME_STEPS_STATE As String = "StepsState"

' W-ANALYTIC-09: имя именованного диапазона для номера последнего шага
Private Const NAME_LAST_STEP_NUMBER As String = "LastStepNumber"
Private Const NAME_LAST_GENERATED_STEP_NUMBER As String = "LastGeneratedStepNumber"

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
    Dim nm As name
    Set nm = wb.names(NAME_PROJECT_ROOT)
    If Err.Number = 0 Then
        v = CStr(nm.RefersToRange.value)
        If Err.Number <> 0 Then v = "": Err.Clear
    Else
        Err.Clear
        Set ws = wb.Worksheets("Настройки")
        If Err.Number = 0 Then v = CStr(ws.Range("B2").value)
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
    raw = Trim$(CStr(ThisWorkbook.names(name).RefersToRange.value))
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

' Читает ТЕКСТОВУЮ настройку листа "Настройки" по имени именованного
' диапазона (PRE-04; аналог GetSettingLong для строковых значений).
' Поведение:
'   * значение непустое -> возвращается как есть (только Trim);
'   * поле пустое / недоступно -> возвращается defaultValue + WriteWarning
'     U007 (штатная ситуация, не ошибка);
'   * имя/лист не найдены -> возвращается defaultValue без предупреждения
'     (корректное состояние до первого запуска EnsureSettingsSheet).
' Используется для LegendRange (Mod_Structure.BuildFlatStructure: адрес
' диапазона легенды цветов на листе схемы, напр. "B1:C3").
Public Function GetSettingText(ByVal name As String, ByVal defaultValue As String) As String
    Dim raw As String

    GetSettingText = defaultValue

    On Error GoTo ErrHandler

    ' Читаем фактическое значение диапазона (RefersToRange.Value), а не
    ' строку определения имени (Names(...).Value) - см. T3.1-fix.
    On Error Resume Next
    raw = Trim$(CStr(ThisWorkbook.names(name).RefersToRange.value))
    If Err.Number <> 0 Then raw = "": Err.Clear
    On Error GoTo 0
    If raw = "" Then
        WriteWarning "U007", "GetSettingText: настройка """ & name & """ пуста - " & _
                   "использовано значение по умолчанию """ & defaultValue & """."
        Exit Function
    End If

    GetSettingText = raw
    Exit Function

ErrHandler:
    ' Имя не найдено / лист "Настройки" отсутствует - тихо отдаём default.
    GetSettingText = defaultValue
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
Public Function FormatMissingFile(ByVal fileName As String) As String
    FormatMissingFile = "[!] Отсутствует обязательный файл: " & fileName
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
Public Function FindConfigFile(ByVal folderPath As String) As String
    Dim fName As String
    Dim xmlFound As String
    Dim txtFound As String
    Dim pilotcfxFound As Boolean

    FindConfigFile = ""

    If Dir(folderPath, vbDirectory) = "" Then Exit Function

    ' проход по всем файлам папки: фиксируем первый .xml, первый .txt
    ' и наличие хотя бы одного .pilotcfx. Dir без атрибутов возвращает
    ' только файлы (не подпапки).
    fName = Dir(folderPath & "\*")
    Do While fName <> ""
        If Left$(fName, 2) <> "~$" Then           ' пропускаем временные файлы Office
            If Right$(LCase$(fName), 4) = ".xml" Then
                If xmlFound = "" Then xmlFound = folderPath & "\" & fName
            ElseIf Right$(LCase$(fName), 4) = ".txt" Then
                If txtFound = "" Then txtFound = folderPath & "\" & fName
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
        WriteWarning "U006", "Найден только .pilotcfx, требуется .xml или .txt (" & folderPath & ")"
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

' ============================================================================
' Создаёт лист "Настройки" (если его нет), заполняет все 14 полей, выпадающие
' списки и именованные диапазоны. Использует единый массив конфигурации.
' Логика строго идемпотентна: существующие подписи не дублируются,
' пользовательские значения в столбце B не перезаписываются.
' ============================================================================
Public Sub EnsureSettingsSheet()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim i As Long
    Dim lastRow As Long
    Dim targetRow As Long
    Dim foundCell As Range
    Dim nmChk As name
    Dim patErr As Boolean
    Dim r As Long
    
    ' Единый массив конфигурации: 0=Название, 1=Значение по умолчанию, 2=Подсказка, 3=Нужен список "Да;Нет" (1=Да, 0=Нет)
    Dim config As Variant
    config = Array( _
        Array("Корень проекта", "", "Путь к корневой папке проекта. Если пусто, используется путь к файлу .xlsm", 0), _
        Array("Режим проверки", "Строгий", "Режим проверки исходных данных (Строгий / Черновик)", 0), _
        Array("Автооткрытие папки результатов", "Да", "Автоматически открывать папку результатов после успешной операции", 0), _
        Array("Версия промта-генератора", "7.13", "Версия промта-генератора", 0), _
        Array("Версия промта-исправления", "1.0", "Версия промта исправления ошибок", 0), _
        Array("Макс. кол-во типов в таблице", "50", "Максимальное количество типов для обработки в одной итерации", 0), _
        Array("Имя листа с легендой", "", "Имя листа в исходной схеме, содержащего легенду", 0), _
        Array("Диапазон легенды", "", "Адрес диапазона ячеек с легендой (например, B1:C3)", 0), _
        Array("Имя листа со структурой", "", "Имя листа в исходной схеме, содержащего данные структуры", 0), _
        Array("Ячейка начала таблицы", "", "Адрес ячейки начала таблицы данных (например, A4)", 0), _
        Array("Выжимки и анализ конфигурации", "Нет", "Фиксирует генерацию промтов выдержек (00_4) и анализа конфигурации (00_1). Ставится автоматически при нажатии кнопки «Сгенерировать аналитические промты».", 1), _
        Array("Разбивка на паттерны", "Нет", "Фиксирует завершение итераций по промту 00_2. Ставится вручную пользователем.", 1), _
        Array("Разбивка на шаги", "Нет", "Фиксирует завершение распределения типов по шагам. Ставится автоматически после успешного парсинга промта 00_3 при полном покрытии.", 1), _
        Array("Номер последнего шага", "нет", "Номер последнего обработанного шага (нет, 00, 01, 01.1, 02 и т.д.). Обновляется автоматически при разборе ответа ИИ.", 0), _
        Array("Номер последнего сгенерированного шага", "нет", "Автоматически обновляется при генерации промта. Если равен ""Номеру последнего шага"", все шаги обработаны.", 0) _
    )
    
    ' Массив имён для именованных диапазонов (должен строго соответствовать порядку в config)
    Dim namesArr As Variant
    namesArr = Array( _
        "ProjectRoot", "CheckMode", "AutoOpenResults", _
        "PromptGenVersion", "PromptFixVersion", "MaxTypesPerTable", _
        "LegendSheetName", "LegendRange", "StructureSheet", "StartCell", _
        "AnalyticsState", "PatternsState", "StepsState", "LastStepNumber", "LastGeneratedStepNumber" _
    )

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Set wb = ThisWorkbook

    ' 1. Поиск или создание листа "Настройки"
    On Error Resume Next
    Set ws = wb.Worksheets("Настройки")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ErrHandler
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.name = "Настройки"
    Else
        On Error GoTo ErrHandler
        ws.Unprotect Password:="pilot"
    End If

    ' 2. Определение последней заполненной строки в столбце A для корректного добавления новых
    lastRow = 1
    For r = 2 To 100
        If Trim$(CStr(ws.Cells(r, 1).value)) <> "" Then
            lastRow = r
        End If
    Next r

    ' 3. Обработка всех 14 параметров в едином цикле
    For i = 0 To UBound(config)
        ' Поиск существующей строки по названию в столбце A
        Set foundCell = ws.Columns(1).Find(What:=CStr(config(i)(0)), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        
        If Not foundCell Is Nothing Then
            targetRow = foundCell.row
        Else
            ' Если не найдено, добавляем в первую свободную строку после lastRow
            targetRow = lastRow + 1
            ' Дополнительная защита от наложения на случайно заполненные строки ниже
            Do While Trim$(CStr(ws.Cells(targetRow, 1).value)) <> ""
                targetRow = targetRow + 1
            Loop
            lastRow = targetRow
        End If
        
        ' Запись названия (если строка была создана новой)
        If foundCell Is Nothing Then
            ws.Cells(targetRow, 1).value = CStr(config(i)(0))
        End If
        
        ' Запись значения по умолчанию ТОЛЬКО если ячейка B пуста (идемпотентность)
        If Trim$(CStr(ws.Cells(targetRow, 2).value)) = "" Then
            ws.Cells(targetRow, 2).value = CStr(config(i)(1))
        End If
        
        ' W-ANALYTIC-06E / W-ANALYTIC-09: Принудительный текстовый формат для полей с номерами шагов и версиями (сохранение ведущих нулей)
        Dim settingName As String
        settingName = CStr(config(i)(0))
        If settingName = "Номер последнего шага" Or settingName = "Номер последнего сгенерированного шага" Or settingName = "Версия промта-генератора" Or settingName = "Версия промта-исправления" Then
            ws.Cells(targetRow, 2).NumberFormat = "@"
        End If
        
        ' Запись подсказки в столбец C, если он пуст
        If Trim$(CStr(ws.Cells(targetRow, 3).value)) = "" Then
            ws.Cells(targetRow, 3).value = CStr(config(i)(2))
        End If
        
        ' Применение Data Validation
        If CInt(config(i)(3)) = 1 Then
            With ws.Cells(targetRow, 2).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Да,Нет"
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
        Else
            ' Очистка валидации для текстовых полей (на случай, если она там была ошибочно)
            ws.Cells(targetRow, 2).Validation.Delete
        End If
        
        ' Создание именованного диапазона (если не существует)
        On Error Resume Next
        Set nmChk = wb.names(CStr(namesArr(i)))
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo ErrHandler
            wb.names.Add name:=CStr(namesArr(i)), RefersTo:="=" & ws.name & "!$B$" & targetRow
        End If
        On Error GoTo ErrHandler
        Set nmChk = Nothing
    Next i

    ' 4. Форматирование и защита листа
    ws.Range("A1").value = "Настройки Препроцессора"
    ws.Columns("A:C").AutoFit

    ' Разблокируем все ячейки, затем блокируем только столбец B от 2-й до последней строки настроек
    ws.Cells.Locked = True
    ws.Range("B2:B" & lastRow).Locked = False
    ws.Protect Password:="pilot"

    WriteLog "EnsureSettingsSheet: лист 'Настройки' создан/обновлён (14 полей, идемпотентная инициализация)"

    ' 5. Подготовка листа "Паттерны" (сохраняем существующую логику вызова)
    patErr = False
    On Error Resume Next
    Mod_Patterns.EnsurePatternsSheet
    If Err.Number <> 0 Then patErr = True
    Err.Clear
    On Error GoTo 0
    If patErr Then
        WriteWarning "U008", "EnsureSettingsSheet: не удалось подготовить лист 'Паттерны' (Mod_Patterns.EnsurePatternsSheet)."
    End If
    
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    WriteError "U003", "EnsureSettingsSheet: ошибка подготовки листа 'Настройки'. Ошибка: " & Err.Description
    MsgBox "Ошибка U003: не удалось подготовить лист 'Настройки'." & vbCrLf & Err.Description, vbCritical, "Настройки"
    Err.Clear
End Sub

' ---------------------------------------------------------------------------
' W-ANALYTIC-08: инициализация трёх параметров состояния на листе "Настройки"
' (раздел 5 спецификации Spec_Dispatcher.md):
'   A) "Выжимки и анализ конфигурации" (AnalyticsState);
'   B) "Разбивка на паттерны"         (PatternsState);
'   C) "Разбивка на шаги"             (StepsState).
' Поведение по каждому параметру:
'   * если подпись уже есть в колонке A диапазона A2:A100 - строка
'     переиспользуется (строки не дублируются, B2:B11 не сдвигаются);
'   * иначе находится первая пустая строка колонки A после последнего
'     заполненного поля (фактически следующая за B11) и поле создаётся;
'   * значение по умолчанию "Нет" ставится ТОЛЬКО в пустую ячейку B
'     (пользовательские значения не сбрасываются при повторном вызове);
'   * Data Validation xlValidateList Formula1:="Да;Нет" перенавешивается
'     на свою ячейку B (идемпотентно);
'   * подсказка (текст назначения из раздела 5) записывается в колонку C;
'   * именованный диапазон создаётся, если ещё не существует.
' Процедура вызывается из EnsureSettingsSheet при снятой защите листа;
' новые ячейки ввода (B) разблокируются (.Locked = False) до повторной
' защиты. Строгая обработка ошибок: On Error GoTo, код ошибки U003
' (реестр Mod_Util), логирование через WriteLog / WriteError.
Private Sub EnsureAnalyticStateFields(ByVal wb As Workbook, ByVal ws As Worksheet)
    Dim labels As Variant
    Dim namesArr As Variant
    Dim hints As Variant
    Dim i As Long
    Dim lastRow As Long
    Dim targetRow As Long
    Dim cellLabel As Range
    Dim cellValue As Range
    Dim cellHint As Range
    Dim rw As Long
    Dim found As Boolean
    Dim nmChk As name

    On Error GoTo ErrHandler

    labels = Array("Выжимки и анализ конфигурации", _
                   "Разбивка на паттерны", _
                   "Разбивка на шаги")

    namesArr = Array(NAME_ANALYTICS_STATE, NAME_PATTERNS_STATE, NAME_STEPS_STATE)

    hints = Array( _
        "Фиксирует генерацию промтов выдержек (00_4) и анализа конфигурации (00_1). " & _
        "Ставится автоматически при нажатии кнопки «Сгенерировать аналитические промты».", _
        "Фиксирует завершение итераций по промту 00_2. Ставится вручную пользователем.", _
        "Фиксирует завершение распределения типов по шагам. Ставится автоматически " & _
        "после успешного парсинга промта 00_3 при полном покрытии.")

    ' Последняя заполненная строка подписей (для поиска первой пустой строки
    ' после существующих настроек; ограничимся разумным диапазоном A2:A100).
    lastRow = 1
    For rw = 2 To 100
        If Trim$(CStr(ws.Cells(rw, 1).value)) <> "" Then lastRow = rw
    Next rw

    For i = 0 To 2
        ' 1. Ищем существующую подпись (идемпотентность: не дублируем поле).
        found = False
        targetRow = 0
        Set cellLabel = ws.Range("A2:A100").Find(What:=CStr(labels(i)), _
                        LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not cellLabel Is Nothing Then
            found = True
            targetRow = cellLabel.row
        End If

        ' 2. Не найдено - первая пустая строка колонки A после последнего поля.
        If Not found Then
            targetRow = lastRow + 1
            ws.Cells(targetRow, 1).value = CStr(labels(i))
            lastRow = targetRow
        End If

        Set cellValue = ws.Cells(targetRow, 2)
        Set cellHint = ws.Cells(targetRow, 3)

        ' 3. Значение по умолчанию "Нет" - только в пустую ячейку
        ' (повторный вызов не сбрасывает установленное пользователем).
        If Trim$(CStr(cellValue.value)) = "" Then
            cellValue.value = "Нет"
        End If

        ' 4. Выпадающий список "Да;Нет" строго на своей ячейке B
        ' (перенавешивание идемпотентно: Delete + Add).
        With cellValue.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="Да; Нет"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With

        ' 5. Подсказка в колонке C (обновляется актуальным текстом).
        cellHint.value = CStr(hints(i))

        ' 6. Ячейка ввода разблокирована под защитой "pilot".
        cellValue.Locked = False

        ' 7. Именованный диапазон (существующие имена не переопределяются).
        On Error Resume Next
        Set nmChk = wb.names(CStr(namesArr(i)))
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo ErrHandler
            wb.names.Add name:=CStr(namesArr(i)), _
                         RefersTo:="=" & ws.name & "!$B$" & targetRow
        End If
        On Error GoTo ErrHandler
        Set nmChk = Nothing
    Next i

    WriteLog "EnsureAnalyticStateFields: параметры состояния 'Выжимки и анализ " & _
             "конфигурации', 'Разбивка на паттерны', 'Разбивка на шаги' " & _
             "проверены/созданы (список Да;Нет, подсказки в колонке C)"
    Exit Sub

ErrHandler:
    WriteError "U003", "EnsureAnalyticStateFields: ошибка добавления параметров " & _
               "состояния на лист 'Настройки'. Ошибка: " & Err.Description
    Err.Clear
End Sub
