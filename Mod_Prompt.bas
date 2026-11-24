Attribute VB_Name = "Mod_Prompt"
Option Explicit

' ============================================================================
' Модуль промптов: формирование промптов для следующего шага и исправлений,
' чтение списка и текущей позиции шагов.
' Реализация W4 (разделы 2, 3, 8, 9, 13 спецификации Spec_Dispatcher.md).
' W7.1-A: добавлена чистая функция GetPromptProfile (маршрутизация профиля
' промта №0.3 "Папки"/"Паттерны"/"Документы" по lastStep) - без кодов
' ошибок (невалидное значение -> ""; MsgBox - ответственность вызывающего).
' Коды ошибок: P001 - нет Шаги.txt, P002 - ошибка формирования промта шага,
' P003 - ошибка формирования промта исправления,
' P004 - формирование Промта №0 (декомпозиция) при отсутствии Шаги.txt
'      (W-PROMPT-GEN-STEP-00),
' P005 - ошибка сохранения текстового файла в UTF-8 (W-UTF8-EXPORT),
' P006 - предупреждение: ошибка/пустой результат чтения текстового файла
'      с автоопределением кодировки (W-UTF8-READ).
' ============================================================================

' ---------------------------------------------------------------------------
' КОНСТАНТЫ (W-PROMPT-GEN-STEP-00)
' ---------------------------------------------------------------------------

' Имя шаблона Промта №0 (декомпозиция) в 04_Промты\Шаблоны.
Private Const TEMPLATE_STEP0_NAME As String = "prompt_step_00_template.md"
' W-PROMPT-PATH-FIX: имя шаблона промта шага в 04_Промты\Шаблоны.
Private Const TEMPLATE_STEP_V1_NAME As String = "prompt_step_template_v1.md"
' Имя файла Промта №0, сохраняемого в 04_Промты.
Private Const STEP0_PROMPT_FILE As String = "Промт_шаг_00_Декомпозиция.md"
' Плейсхолдер вместо ненайденного входного файла.
Private Const STEP0_NOT_FOUND As String = "НЕ НАЙДЕН"

Private Const UA_HEADER_PREFIX As String = "ТРЕБУЮТ РЕШЕНИЯ"

' ---------------------------------------------------------------------------
' КОНСТАНТЫ W-PROMPT-INVARIANTS-CODE (подстановка нового шаблона шага NN)
' ---------------------------------------------------------------------------

' Перечень плейсхолдеров нового шаблона prompt_step_template_v1.md
' (раздел 4 "Легенда плейсхолдеров" отчёта аналитика; сверено с файлом).
' Разделитель - ";". Плейсхолдеры, отсутствующие в конкретном шаблоне,
' молча пропускаются при подстановке (FillPh проверяет наличие).
Private Const TEMPLATE_PLACEHOLDERS As String = _
    "{{NN}};{{ИМЯ_ШАГА}};{{ЗАДАЧА_ШАГА_ОСОБЕННОСТИ}};{{СТРУКТУРА_ФАЙЛ}};" & _
    "{{ОПИСАНИЕ_ФАЙЛ}};{{КОНФИГУРАЦИЯ_ФАЙЛ}};{{ФАЙЛЫ_ДОП_ДАННЫХ}};{{ШАГИ_ФАЙЛ}};" & _
    "{{СПИСОК_ВХОДОВ}};{{ТАБЛИЦА_ИЗВЕСТНЫХ_GUID}};{{РЕЕСТР_ТИПОВ}};{{РЕЕСТР_ПОКРЫТИЯ}};" & _
    "{{СПИСОК_БАЗОВЫХ_ТИПОВ}};{{ПРАВИЛО_ОБЪЕКТА}};{{СПИСОК_ТИПОВ_ШАГА}};{{ТЕКУЩИЙ_ШАГ}};" & _
    "{{ДЕКОМПОЗИЦИЯ_ШАГОВ}};{{ФОРМАТ_РЕЕСТРОВ}};{{ПРОТОКОЛ_ШАГА}};{{МАРКЕРЫ}}"

' Массив известных плейсхолдеров для проверки "неизвестный {{...}} в шаблоне"
' (инициализируется один раз из TEMPLATE_PLACEHOLDERS, см. EnsureKnownPh).
Private SPLIT_KNOWN_PH() As Variant
Private SPLIT_KNOWN_PH_READY As Boolean


' Имя файла общих настроек проекта (константы промта: списки, правила,
' форматы). Хранится рядом с книгой .xlsm или во внешней папке "Шаблоны";
' формат строк: {{ИМЯ_ПЛЕЙСХОЛДЕРА}} = <значение>. Значения НЕ хардкодятся
' в коде - только из этого файла (требование задания).
Private Const PROJECT_SETTINGS_FILE As String = "project_settings.md"

' Маркер-страховка задачи (ЧАСТЬ D п.7): подставляется в блок 2 вместо
' постоянного текста задачи, если блок 6 пуст (типов шага нет).
Private Const TASK_NOT_SET_MARKER As String = "ЗАДАЧА НЕ ЗАДАНА - генерация запрещена"

' Заготовка блока 16 в новом шаблоне (плейсхолдер {{NN}} сохраняется до
' общего прогона подстановки - сравнение идёт по этому тексту).
Private Const CORRECTION_STUB_TEXT As String = "Корректировок для шага {{NN}} нет."

' Тексты блоков 4.1/4.2 (ЧАСТЬ B п.3).
Private Const REG_EMPTY_FIRST_STEP As String = "Реестры пусты - первый шаг (P002 штатно)"
Private Const REG_NOT_FOUND As String = "[РЕЕСТР НЕ НАЙДЕН]"

' Обязательные поля строки корректировки (ЧАСТЬ D п.6) и канон колонок
' адаптированной структуры для фильтрации блока 6 (ЧАСТЬ B п.4).
Private Const FIELD_STEP As String = "ШАГ:"
Private Const FIELD_TYPE As String = "ТИП:"
Private Const FIELD_COND As String = "УСЛОВИЕ:"
Private Const FIELD_CLAR As String = "УТОЧНЕНИЕ:"
Private Const COL_STEP As String = "Шаг"
' PRE-04: в плоской таблице (Mod_Structure.WriteFlatTable, 7 колонок A-G)
' заголовок "Тип" переименован в "Категория", а колонка номера шага получила
' канонический заголовок "№шага/подшага". Константы ниже - новые канонические
' имена; старые ("Тип", "Шаг") сохранены в распознавании заголовков для
' обратной совместимости с артефактами, собранными до PRE-04.
Private Const COL_CATEGORY As String = "Категория"       ' новый канон (PRE-04)
Private Const COL_STEP_NUM As String = "№шага/подшага"   ' новый канон (PRE-04)
Private Const COL_SYS_NAME As String = "Тип"             ' старый заголовок категории
Private Const COL_VIS_NAME As String = "Отображаемое имя"

' W-PROMPT-00-SCHEMA-FIX: имя артефакта подготовки схемы (лист книги и
' префикс сохраняемого файла в папке 02_Подготовка схемы).
Private Const SCHEMA_SHEET_NAME As String = "Адаптированная_структура"
' Маска файлов артефакта в 02_Подготовка схемы.
Private Const SCHEMA_ARTIFACT_MASK As String = "Адаптированная_структура_*.xlsx"
' Подпапка архива старых версий артефакта.
Private Const SCHEMA_ARCHIVE_SUBFOLDER As String = "Архив"

Private Const FRAGMENTS_FOLDER_NAME As String = "05_Фрагменты шагов"
Private Const ARCHIVE_FOLDER_NAME As String = "Архив"
Private Const MARKER_PATTERNS As String = "PATTERNS_CSV"
Private Const MARKER_STEPS As String = "STEPS_CSV"
Private Const MARKER_STOP As String = "СТОП"
Private Const MARKER_LEGACY As String = "legacy"
Private Const SKIP_FILE_PREFIX_1 As String = "~$"
Private Const SKIP_FILE_PREFIX_2 As String = "prompt_"
Private Const SKIP_FILE_EXACT As String = "Требуют_решения"

' ---------------------------------------------------------------------------
' W-ANALYTIC-06C: константы маркеров и заголовков колонок для разбора 00_2/00_3
' ---------------------------------------------------------------------------
Private Const COL_PATTERN As String = "Паттерн"
Private Const COL_ANOMALY As String = "Аномалия"

' Маркеры блоков ответа ИИ
Private Const MK_PAT_START As String = "<!-- PATTERNS_CSV_START -->"
Private Const MK_PAT_END   As String = "<!-- PATTERNS_CSV_END -->"
Private Const MK_STP_START As String = "<!-- STEPS_CSV_START -->"
Private Const MK_STP_END   As String = "<!-- STEPS_CSV_END -->"
Private Const MK_REF_START As String = "<!-- REFERENCEBOOKS_START -->"
Private Const MK_REF_END   As String = "<!-- REFERENCEBOOKS_END -->"
Private Const MK_UA_START  As String = "<!-- USER_ACTION_START -->"
Private Const MK_UA_END    As String = "<!-- USER_ACTION_END -->"
Private Const MK_STAT_START As String = "<!-- STATISTICS_START -->"
Private Const MK_STAT_END   As String = "<!-- STATISTICS_END -->"

' Служебные строки
Private Const MARK_STOP    As String = "СТОП"
Private Const MARK_NO      As String = "нет"
Private Const UA_NO_HEADER As String = "ТРЕБУЮТ РЕШЕНИЯ ПОЛЬЗОВАТЕЛЯ: нет"
Private Const UA_HDR_SUBSTR As String = "ТРЕБУЮТ РЕШЕНИЯ"
Private Const PAT_HDR_SUBSTR As String = "Паттерн|Типы|Количество"
Private Const REF_MARKER_001 As String = "{{СПИСОК_БАЗОВЫХ_ТИПОВ}}"
Private Const STAT_GEN_STEPS_HDR As String = "Сгенерированные шаги:"
Private Const STAT_NOT_APPLICABLE As String = "не применимо"
Private Const STAT_EMPTY_MARK As String = "-"

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

' ---------------------------------------------------------------------------
' БЕЗОПАСНОЕ ЧТЕНИЕ ТЕКСТОВЫХ ФАЙЛОВ (W-UTF8-READ)
' ---------------------------------------------------------------------------

' Читает текстовый файл целиком в строку с автоопределением кодировки.
' Встроенный Open ... For Input читает в ANSI (Windows-1251), из-за чего
' файлы шаблонов / Шаги.txt / Реестры, сохранённые в UTF-8, превращаются
' в "кракозябры". Поэтому сначала используется ADODB.Stream с
' Charset = "utf-8"; если чтение через UTF-8 не удалось или дало пустой
' результат - fallback на Charset = "windows-1251" (ANSI), затем - на
' стандартное Open ... For Input. При любой ошибке чтения - запись [WARN]
' в лог (код P006) и возврат пустой строки (без прерывания вызывающего).
' Если файла нет - возвращает пустую строку (вызывающий сам решает,
' считать ли это ошибкой).
Private Function SafeReadText(ByVal filePath As String) As String
    Dim res As String

    SafeReadText = ""

    ' W-PROMPT-PATHS-REFACTOR: vbNormal - строгая проверка существования
    ' именно файла (а не папки с таким же именем).
    If Len(Dir(filePath, vbNormal)) = 0 Then Exit Function

    On Error Resume Next

    ' --- Попытка 1: UTF-8 через ADODB.Stream -------------------------------
    Err.Clear
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2                 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    res = stream.ReadText(-1)       ' adReadAll
    stream.Close
    Set stream = Nothing
    If Err.Number <> 0 Then
        ' поток мог остаться открытым - закрыть его, игнорируя ошибки
        If Not stream Is Nothing Then
            Err.Clear
            stream.Close
            Set stream = Nothing
        End If
        res = ""
    End If
    Err.Clear

    ' W-PROMPT-00-REFINE, п.4: ADODB.Stream при чтении UTF-8 возвращает
    ' символом U+FEFF (BOM) в начале строки - удаляем, чтобы BOM не
    ' переносился в генерируемые промты.
    If Left$(res, 1) = ChrW(&HFEFF) Then res = Mid$(res, 2)

    ' --- Fallback 1: Windows-1251 (ANSI) -----------------------------------
    If res = "" Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2             ' adTypeText
        stream.Charset = "windows-1251"
        stream.Open
        stream.LoadFromFile filePath
        res = stream.ReadText(-1)   ' adReadAll
        stream.Close
        Set stream = Nothing
        If Err.Number <> 0 Then
            If Not stream Is Nothing Then
                Err.Clear
                stream.Close
                Set stream = Nothing
            End If
            res = ""
        End If
        Err.Clear
    End If

    ' --- Fallback 2: стандартное чтение Open ... For Input (ANSI) ----------
    If res = "" Then
        Dim fNum As Integer
        Dim s As String
        fNum = FreeFile
        Open filePath For Input Access Read As #fNum
        Do While Not EOF(fNum)
            Line Input #fNum, s
            res = res & s & vbCrLf
        Loop
        Close #fNum
        If Err.Number <> 0 Then
            If fNum <> 0 Then Close #fNum
            res = ""
        End If
        ' убрать финальный перевод строки
        If Len(res) >= 2 Then
            If Right(res, 2) = vbCrLf Then res = Left(res, Len(res) - 2)
        End If
        Err.Clear
    End If

    On Error GoTo 0

    ' Файл существует, но ни один из способов чтения не вернул текст либо
    ' произошла ошибка - фиксируем предупреждение в журнале ([WARN] P006)
    ' и возвращаем пустую строку.
    If res = "" Then
        Mod_Util.WriteWarning "P006", "SafeReadText: файл пуст или не может " & _
            "быть прочитан: " & filePath
    End If

    SafeReadText = res
End Function

' Собирает список имён файлов в папке folderPath через Dir() без рекурсии
' (W-CORRECTION-PACK-1, п.4.3). Временные файлы Office (~$) пропускаются.
' Возвращает строку "- имя1" & vbCrLf & "- имя2" ... либо "" (папка пуста
' или не существует).
Private Function ListFolderFiles(ByVal folderPath As String) As String
    Dim fName As String
    Dim res As String

    ListFolderFiles = ""
    If Dir(folderPath, vbDirectory) = "" Then Exit Function

    fName = Dir(folderPath & "\*")        ' Dir без атрибутов - только файлы
    Do While fName <> ""
        If Left$(fName, 2) <> "~$" Then   ' пропускаем временные файлы Office
            If res <> "" Then res = res & vbCrLf
            res = res & "- " & fName
        End If
        fName = Dir()
    Loop
    ListFolderFiles = res
End Function

' W-ANALYTIC-04-FIX: собирает имена элементов папки folderPath в массив
' (1 To N) одноразовым итератором Dir() БЕЗ вложенных вызовов Dir() с
' аргументом внутри цикла.
' W-ANALYTIC-07.1-FIX: добавлен атрибут vbDirectory и маска "\*" вместо
' "\*.*" - теперь в массив попадают и ПОДПАПКИ ( Шаг_01, Шаг_01.1 ),
' а не только файлы. Служебные элементы "." и ".." отбрасываются.
' Папка отсутствует или пуста -> возвращается пустой массив Array().
Private Function GetFileListArray(ByVal folderPath As String) As Variant
    Dim col As Collection
    Dim fName As String
    Dim arr() As String
    Dim i As Long

    GetFileListArray = Array()
    If Not FolderExistsSafe(folderPath) Then Exit Function

    Set col = New Collection
    fName = Dir(folderPath & "\*", vbDirectory)
    Do While fName <> ""
        If fName <> "." And fName <> ".." Then col.Add fName
        fName = Dir()
    Loop
    If col.Count = 0 Then Exit Function

    ReDim arr(1 To col.Count)
    For i = 1 To col.Count
        arr(i) = CStr(col(i))
    Next i
    GetFileListArray = arr
End Function

' УСТАРЕЛО (W-PROMPT-INVARIANTS-CODE-2): функция BuildExtraDataBlock удалена -
' старая сборка блока "ДОПОЛНИТЕЛЬНЫЕ ДАННЫЕ ПРОЕКТА" для промта шага заменена на
' BuildExtraDataListBlock (блок 3.4, вызывается из FillAllPlaceholders внутри
' BuildPromptForStep). Для Промта №0 (декомпозиция) сохранён
' BuildStepZeroExtraDataList.

' Формирует содержимое плейсхолдера {EXTRA_DATA_LIST} для Промта №0
' (W-PROMPT-00-REFINE, п.2/п.5; W-PROMPT-00-FORMAT-FIX, п.5): список
' имён и типов файлов из папки "01_Исходные данные\04_Дополнительные
' данные" - он подставляется в раздел "Проверка файлов" шаблона
' prompt_step_00_template.md на отдельной строке плейсхолдера.
' Заголовок "Найденные файлы доп. данных:" формируется ЗДЕСЬ (в шаблоне
' статичная строка с заголовком удалена), поэтому:
' - если файлов нет (папка пуста или отсутствует) - возвращается пустая
'   строка "" без заголовка (в итоговом markdown не остаётся висящих
'   заголовков и лишних пустых строк);
' - если файлы есть - возвращается заголовок и имена файлов с отступом.
Private Function BuildStepZeroExtraDataList(ByVal root As String) As String
    Dim filesList As String

    BuildStepZeroExtraDataList = ""
    filesList = ListFolderFiles(root & "\02_Подготовка данных\04_Дополнительные данные")
    If filesList = "" Then Exit Function

    BuildStepZeroExtraDataList = "- Найденные файлы доп. данных:" & _
        vbCrLf & filesList
End Function

' Приводит текст к единым разделителям строк vbCrLf (W-UTF8-READ):
' ADODB.Stream сохраняет исходные переводы строки файла - CRLF, LF или CR.
Private Function NormalizeLineEndings(ByVal text As String) As String
    Dim res As String
    res = Replace(text, vbCrLf, vbLf)     ' CRLF -> LF
    res = Replace(res, vbCr, vbLf)        ' одиночный CR -> LF
    NormalizeLineEndings = Replace(res, vbLf, vbCrLf)  ' LF -> CRLF
End Function

' Читает текстовый файл целиком в строку (W-UTF8-READ: делегирует
' SafeReadText с автоопределением UTF-8 / Windows-1251).
' Если файла нет или чтение не удалось - возвращает пустую строку
' (вызывающий сам решает, считать ли это ошибкой).
' W-DOCS-README-UPDATE: Public - функция используется Mod_Docs.UpdateReadme
' для чтения шаблонов промтов при включении их текстов в Readme.
Public Function ReadTextFile(ByVal filePath As String) As String
    ReadTextFile = SafeReadText(filePath)
End Function

' W-PROMPT-INVARIANTS-CODE: существование именно ФАЙЛА (не папки);
' Dir(..., vbNormal) даёт ложный False для имён с кириллицей/пробелами,
' поэтому проверка выполняется через GetAttr (ошибка 76 - файла нет).
Private Function FileExistsSafe(ByVal filePath As String) As Boolean
    Dim attrs As Long

    FileExistsSafe = False
    If Len(filePath) = 0 Then Exit Function
    On Error Resume Next
    Err.Clear
    attrs = GetAttr(filePath)
    If Err.Number = 0 Then
        ' обычный файл: без атрибута vbDirectory и без vbVolume
        If (attrs And vbDirectory) = 0 And (attrs And vbVolume) = 0 Then _
            FileExistsSafe = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

' W-PROMPT-INVARIANTS-CODE: существование папки (без ошибок Dir по
' кириллическим путям).
Private Function FolderExistsSafe(ByVal folderPath As String) As Boolean
    Dim attrs As Long

    FolderExistsSafe = False
    If Len(folderPath) = 0 Then Exit Function
    On Error Resume Next
    Err.Clear
    attrs = GetAttr(folderPath)
    If Err.Number = 0 Then
        If (attrs And vbDirectory) = vbDirectory Then FolderExistsSafe = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

' Сохраняет строку в текстовый файл в кодировке UTF-8 БЕЗ BOM
' (W-UTF8-EXPORT, W-PROMPT-00-REFINE п.5). Перезаписывает существующий
' файл. Встроенный оператор Print # пишет в ANSI (Windows-1251), из-за
' чего кириллица в промтах отображается как "кракозябры" в UTF-8
' редакторах и на GitHub. Поэтому используется ADODB.Stream с
' Charset = "utf-8". Стандартный WriteText через Charset="utf-8" может
' добавить заголовок EF BB BF (BOM); чтобы этого избежать, содержимое
' перед записью очищается от ведущего U+FEFF, а после сохранения файл
' перечитывается в байтовый поток и, при наличии BOM, перезаписывается
' без первых трёх байт. При ошибке записи - запись в лог (код P005)
' и возврат без прерывания вызывающей процедуры.
Public Sub SaveTextFileUtf8(ByVal filePath As String, ByVal content As String)
    On Error GoTo ErrHandler

    Dim stream As Object

    ' Страховка: если содержимое начинается с BOM (U+FEFF) - убрать его,
    ' иначе ADODB.Stream сохранит удвоенный заголовок.
    If Left$(content, 1) = ChrW(&HFEFF) Then content = Mid$(content, 2)

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2                 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText content
    stream.SaveToFile filePath, 2   ' adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing

    ' Вырезаем BOM (EF BB BF), если он всё же был добавлен при записи.
    StripUtf8Bom filePath
    Exit Sub

ErrHandler:
    ' Освободить поток, если он остался открытым.
    If Not stream Is Nothing Then
        On Error Resume Next
        stream.Close
        Set stream = Nothing
        On Error GoTo 0
    End If
    Mod_Util.WriteError "P005", "SaveTextFileUtf8: ошибка сохранения файла " & _
        filePath & ". Ошибка: " & Err.Description
End Sub

' Удаляет из начала файла filePath заголовок UTF-8 BOM (байты EF BB BF),
' если он там есть (W-PROMPT-00-REFINE, п.4/п.5). Реализация через
' ADODB.Stream (бинарный режим); при любой ошибке файл остаётся как есть.
Private Sub StripUtf8Bom(ByVal filePath As String)
    Dim bytes() As Byte
    Dim bin As Object

    On Error GoTo ErrHandler

    If Dir(filePath, vbNormal) = "" Then Exit Sub

    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1                    ' adTypeBinary
    bin.Open
    bin.LoadFromFile filePath

    If bin.Size >= 3 Then
        bin.Position = 0
        bytes = bin.Read(3)
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            ' Первые три байта - BOM: перечитываем тело без него.
            Dim body As Variant
            body = bin.Read(bin.Size - 3)
            bin.Close
            Dim out As Object
            Set out = CreateObject("ADODB.Stream")
            out.Type = 1            ' adTypeBinary
            out.Open
            If UBound(body) >= 0 Then out.Write body
            out.SaveToFile filePath, 2   ' adSaveCreateOverWrite
            out.Close
            Set out = Nothing
            Exit Sub
        End If
    End If

    bin.Close
    Set bin = Nothing
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not bin Is Nothing Then
        If bin.State = 1 Then bin.Close   ' adStateOpen
        Set bin = Nothing
    End If
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
' в заданной папке. Если файлов нет - возвращает "".
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

' Читает лист "Ошибки" книги сводной таблицы в текст
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

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1

    ' строка 1 - заголовки ("Код", "Источник", "Сообщение"), строки 2+ - ошибки
    For r = 2 To lastRow
        lineText = ""
        For c = 1 To lastCol
            If c > 1 Then lineText = lineText & " | "
            lineText = lineText & Trim(CStr(ws.Cells(r, c).value))
        Next c
        If Trim(lineText) <> "" Then res = res & lineText & vbCrLf
    Next r

    If Len(res) >= 2 Then
        If Right(res, 2) = vbCrLf Then res = Left(res, Len(res) - 2)
    End If
    outText = res
    ReadErrorsSheetText = True
End Function

' Возвращает имя первого найденного файла по маске в папке (для Промта №0).
' Если файлов нет - возвращает "". Временные файлы Office (~$) пропускаются.
Private Function FirstMatchFile(ByVal folderPath As String, ByVal mask As String) As String
    Dim fName As String

    FirstMatchFile = ""
    If Dir(folderPath, vbDirectory) = "" Then Exit Function

    fName = Dir(folderPath & "\" & mask)
    Do While fName <> ""
        If Left(fName, 2) <> "~$" Then       ' пропускаем временные файлы Office
            FirstMatchFile = fName
            Exit Function
        End If
        fName = Dir()
    Loop
End Function



' W-PROMPT-00-SCHEMA-FIX: извлекает из имени файла подстроку вида
' yyyy-mm-dd_hhnnss (16 символов). Сравнивать такие подстроки как строки
' корректно, поскольку формат лексикографически упорядочен. Если шаблон в
' имени не найден - возвращается "".
Private Function ExtractDateStampFromName(ByVal fileName As String) As String
    Dim i As Long
    Dim ok As Boolean
    Dim ch As String

    ExtractDateStampFromName = ""
    For i = 1 To Len(fileName) - 15
        ok = Mid$(fileName, i + 4, 1) = "-" And Mid$(fileName, i + 7, 1) = "-" And _
             Mid$(fileName, i + 10, 1) = "_"
        If ok Then
            Dim j As Long
            For j = 0 To 15
                Select Case j
                    Case 0 To 3, 5 To 6, 8 To 9, 11 To 15
                        ch = Mid$(fileName, i + j, 1)
                        If ch < "0" Or ch > "9" Then
                            ok = False
                            Exit For
                        End If
                    Case 4, 7
                        ' разделители "-" уже проверены
                    Case 10
                        ' разделитель "_" уже проверен
                End Select
            Next j
        End If
        If ok Then
            ExtractDateStampFromName = Mid$(fileName, i, 16)
            Exit Function
        End If
    Next i
End Function

' Собирает фрагмент схемы из листа "Адаптированная_структура" книги, сохранённой
' в папке 02_Подготовка схемы (W-PROMPT-00-SCHEMA-FIX: берётся самый свежий
' файл "Адаптированная_структура_*.xlsx"; поддержка старого имени
' "Плоская_таблица*" сохранена для книг, созданных до переименования).
' Если файла/листа нет - возвращает "" и пишет предупреждение в лог.
Private Function GetSchemaFragment() As String
    Dim root As String
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

    ' W-PROMPT-PATHS-REFACTOR: нормализация корневого пути (защита от
    ' двойных слешей, если GetProjectRoot вернёт путь с завершающим "\")
    root = NormPath(Mod_Util.GetProjectRoot)
    prepPath = root & "\02_Подготовка данных\01_Схема"
    If Dir(prepPath, vbDirectory) = "" Then
        Mod_Util.WriteWarning "P002", "GetSchemaFragment: папка 02_Подготовка схемы не найдена, схема не включена."
        Exit Function
    End If

    ' W-PROMPT-00-SCHEMA-FIX: ищем самый свежий файл артефакта
    ' "Адаптированная_структура_*.xlsx"; если таких нет - старый
    ' "Плоская_таблица*" (обратная совместимость с книгами до переименования),
    ' затем любой xlsx/xlsm в папке.
    candidates(1) = SCHEMA_ARTIFACT_MASK
    candidates(2) = "Плоская_таблица*"
    candidates(3) = "*.xlsx"
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
        Mod_Util.WriteWarning "P002", "[!] Отсутствует обязательный файл: Адаптированная структура" & _
                                 " (в 02_Подготовка схемы нет файла " & SCHEMA_ARTIFACT_MASK & "), схема не включена."
        Exit Function
    End If

    On Error Resume Next
    ' W-AV-REFACTOR, п.1: AddToMru:=False - файлы шагов не попадают в список
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
    Set ws = wb.Worksheets(SCHEMA_SHEET_NAME)
    If ws Is Nothing Then
        ' обратная совместимость: книги, созданные до переименования
        Err.Clear
        Set ws = wb.Worksheets("Плоская_таблица")
    End If
    On Error GoTo 0
    If ws Is Nothing Then
        Mod_Util.WriteWarning "P002", "GetSchemaFragment: в файле '" & newest & "' нет листа '" & _
                                      SCHEMA_SHEET_NAME & "', схема не включена."
        wb.Close SaveChanges:=False
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1

    For r = 1 To lastRow
        lineText = ""
        For c = 1 To lastCol
            If c > 1 Then lineText = lineText & " | "
            lineText = lineText & Trim(CStr(ws.Cells(r, c).value))
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
' W-PROMPT-INVARIANTS-CODE: СБОРКА ПРОМТА ШАГА ПОД НОВЫЙ ШАБЛОН
' (prompt_step_template_v1.md, 17 блоков; инвариантное ядро 7-11 - жёсткий
' текст шаблона, код его не генерирует)
' ---------------------------------------------------------------------------

' Заменяет плейсхолдер {{...}} в тексте на значение и возвращает True, если
' замена выполнена. Если плейсхолдера нет в тексте - ничего не делает
' (возвращает False). Если значение пустое - подставляет явный текст
' "[НЕ ЗАПОЛНЕНО: <имя>]" и пишет WriteWarning с кодом и именем плейсхолдера
' (требование ЧАСТИ A: незаполненный плейсхолдер не оставляется в промте
' молча). Реализация через Split/Join - надёжнее встроенной Replace для
' длинных текстов (защита от ошибки 6 "Permitted exception" при больших
' строках).
Private Function FillPh(ByVal text As String, ByVal phName As String, _
                        ByVal value As String, ByRef outText As String) As Boolean
    Dim parts() As String

    FillPh = False
    outText = text
    If InStr(text, phName) = 0 Then Exit Function

    If Trim$(value) = "" Then
        value = "[НЕ ЗАПОЛНЕНО: " & phName & "]"
        Mod_Util.WriteWarning "P002", "FillPh: плейсхолдер " & phName & _
            " не удалось заполнить - подставлен явный маркер ""[НЕ ЗАПОЛНЕНО]""."
    End If

    parts = Split(text, phName)
    outText = Join(parts, value)
    FillPh = True
End Function

' Подставляет ВСЕ плейсхолдеры нового шаблона промта шага NN
' (W-PROMPT-INVARIANTS-CODE, ЧАСТЬ A). Список соответствует реальному
' содержимому prompt_step_template_v1.md; плейсхолдеры, отсутствующие в
' шаблоне, молча пропускаются (FillPh проверяет InStr). Значения собираются
' в словарь по имени плейсхолдера; отсутствующее в словаре значение
' трактуется как незаполненное ("[НЕ ЗАПОЛНЕНО: ...]").
' ВАЖНО (W-PROMPT-INVARIANTS-CODE-2): вызывается ПОСЛЕ ApplyCorrectionsBlock
' - цепочка сборки: загрузка шаблона ? блоки 3.4/4.1/4.2/6 ? сохранение
' секции 16 (ApplyCorrectionsBlock, работает с {{NN}} как есть) ?
' FillAllPlaceholders (подстановку {{КОРРЕКТИРОВКИ_ШАГА_NN}} не выполняет:
' такого плейсхолдера в новом шаблоне нет).
' Возвращает True, если блок 6 ({СПИСОК_ТИПОВ_ШАГА}) заполнен хотя бы одним
' типом шага (для маркер-страховки задачи, ЧАСТЬ D п.7).
Private Function FillAllPlaceholders(ByVal templateText As String, _
                                     ByVal nn As String, _
                                     ByVal stepName As String, _
                                     ByVal registriesPath As String, _
                                     ByVal root As String, _
                                     ByVal valuesPath As String, _
                                     ByRef outResult As String, _
                                     ByRef hasStepTypes As Boolean) As Boolean
    Dim values As Object          ' Scripting.Dictionary: имя плейсхолдера -> текст
    Dim pairs() As String
    Dim i As Long
    Dim phName As String
    Dim filledCount As Long
    Dim totalNames As Long
    Dim tmp As String
    Dim ok As Boolean
    Dim dummy As Boolean

    On Error GoTo ErrHandler

    Set values = CreateObject("Scripting.Dictionary")
    values.CompareMode = vbBinaryCompare

    ' --- Variants: значения, зависящие от текущего шага --------------------
    values("{{NN}}") = nn
    values("{{ИМЯ_ШАГА}}") = stepName
    ' Блок 3.1: артефакт адаптированной структуры (имя файла + назначение);
    ' сам лист читается отдельно в блоке 6 (BuildStepTypeList).
    values("{{СТРУКТУРА_ФАЙЛ}}") = BuildInputFileLine(root, "02_Подготовка схемы", _
        SCHEMA_ARTIFACT_MASK, "Адаптированная структура (источник иерархии, полей ""Тип"" и аномалий)")
    values("{{ОПИСАНИЕ_ФАЙЛ}}") = BuildInputFileLine(root, "01_Исходные данные\02_Описание", "*.docx", "Описание структуры ЖЦ ОКС (источник атрибутов и бизнес-логики)")
    values("{{КОНФИГУРАЦИЯ_ФАЙЛ}}") = BuildConfigFileLine(root)
    values("{{ФАЙЛЫ_ДОП_ДАННЫХ}}") = BuildExtraDataListBlock(root, nn)   ' блок 3.4
    values("{{ШАГИ_ФАЙЛ}}") = BuildStepsFileLine(root)
    values("{{РЕЕСТР_ТИПОВ}}") = BuildRegistryBlock(registriesPath, "Реестр_типов.txt", nn)     ' блок 4.1
    values("{{РЕЕСТР_ПОКРЫТИЯ}}") = BuildRegistryBlock(registriesPath, "Реестр_покрытия.txt", nn) ' блок 4.2
    values("{{СПИСОК_ТИПОВ_ШАГА}}") = BuildStepTypeList(root, nn)       ' блок 6
    hasStepTypes = (Trim$(CStr(values("{{СПИСОК_ТИПОВ_ШАГА}}"))) <> "")

    ' --- Константы проекта: из общих настроек (файл project_settings.md),
    '     НЕ хардкодятся в коде (требование ЧАСТИ A / "НЕ хардкодить") -----
    LoadProjectSettings values, valuesPath

    ' Текст задачи шага: из настроек; если не задан и типов шага нет -
    ' маркер-страховка (ЧАСТЬ D п.7): триггер в блоке 12 уже есть в шаблоне,
    ' здесь меняем только текст задачи в блоке 2.
    If values.Exists("{{ЗАДАЧА_ШАГА_ОСОБЕННОСТИ}}") Then
        If Trim$(CStr(values("{{ЗАДАЧА_ШАГА_ОСОБЕННОСТИ}}"))) = "" Then _
            values.RemoveKey "{{ЗАДАЧА_ШАГА_ОСОБЕННОСТИ}}"
    End If
    If Not hasStepTypes Then
        values("{{ЗАДАЧА_ШАГА_ОСОБЕННОСТИ}}") = TASK_NOT_SET_MARKER
    End If

    ' --- Перечень плейсхолдеров нового шаблона -----------------------------
    pairs = Split(TEMPLATE_PLACEHOLDERS, ";")
    totalNames = UBound(pairs) - LBound(pairs) + 1

    outResult = templateText
    filledCount = 0
    For i = LBound(pairs) To UBound(pairs)
        phName = Trim$(pairs(i))
        If phName <> "" Then
            tmp = ""
            If values.Exists(phName) Then tmp = CStr(values(phName))
            ok = FillPh(outResult, phName, tmp, outResult)
            If ok Then filledCount = filledCount + 1
        End If
    Next i

    ' Проверка "Сверься с шаблоном": если в шаблоне остался необработанный
    ' плейсхолдер {{...}} (одинарные { } - служебные метки комментариев) -
    ' предупреждение в лог, генерацию не блокируем.
    dummy = HasUnknownPlaceholder(outResult)

    ' Если в шаблоне вообще не найдено ни одногоKnown-плейсхолдера -
    ' вероятна подстановка старого шаблона: предупреждаем.
    If filledCount = 0 Then
        Mod_Util.WriteWarning "P002", "FillAllPlaceholders: в шаблоне не найдено " & _
            "ни одного плейсхолдера нового формата ({" & totalNames & " имён}). " & _
            "Проверьте, что в 04_Промты\Шаблоны лежит актуальный " & TEMPLATE_STEP_V1_NAME & "."
    End If

    FillAllPlaceholders = hasStepTypes
    Exit Function

ErrHandler:
    Mod_Util.WriteError "P002", "FillAllPlaceholders: ошибка подстановки плейсхолдеров. " & _
        "Ошибка: " & Err.Description
    Err.Clear
    If outResult = "" Then outResult = templateText
    FillAllPlaceholders = hasStepTypes
End Function

' Пишет WriteWarning, если в собранном тексте остались плейсхолдеры {{...}},
' не входящие в известный перечень (новый плейсхолдер в шаблоне - обработать
' его в коде; пока - явный сигнал в журнале).
Private Function HasUnknownPlaceholder(ByVal text As String) As Boolean
    Dim pos As Long
    Dim rest As String
    Dim name As String
    Dim endPos As Long
    Dim known As Boolean
    Dim i As Long

    EnsureKnownPh
    HasUnknownPlaceholder = False
    rest = text
    Do
        pos = InStr(rest, "{{")
        If pos = 0 Then Exit Do
        endPos = InStr(pos + 2, rest, "}}")
        If endPos = 0 Then Exit Do
        name = Mid$(rest, pos, endPos - pos + 2)
        known = False
        For i = 0 To UBound(SPLIT_KNOWN_PH)
            If SPLIT_KNOWN_PH(i) = name Then
                known = True
                Exit For
            End If
        Next i
        If Not known Then
            Mod_Util.WriteWarning "P002", "HasUnknownPlaceholder: в шаблоне найден " & _
                "неизвестный плейсхолдер " & name & " - подстановка не выполнена, " & _
                "требуется доработка кода (Сверка с шаблоном, ЧАСТЬ A)."
            HasUnknownPlaceholder = True
        End If
        rest = Mid$(rest, endPos + 2)
    Loop
End Function

' Читает файл общих настроек проекта (project_settings.md) и кладёт значения
' плейсхолдеров в словарь values. Формат файла: строки вида
'   {{ИМЯ_ПЛЕЙСХОЛДЕРА}} = <значение в одну строку>
' Многострочное значение можно задать несколько раз - строки склеиваются
' через vbCrLf. Строки-комментарии (начинаются с '#') и пустые пропускаются.
' Если файл не найден - WriteWarning (константы проекта должны собираться из
' общих настроек, а не хардкодиться); незаполненные плейсхолдеры позже
' получат "[НЕ ЗАПОЛНЕНО: ...]".
Private Sub LoadProjectSettings(ByVal values As Object, ByVal settingsPath As String)
    Dim text As String
    Dim lines() As String
    Dim i As Long
    Dim s As String
    Dim eqPos As Long
    Dim keyName As String
    Dim valText As String

    On Error GoTo ErrHandler

    text = ReadTextFile(settingsPath)
    If text = "" Then
        If Not FileExistsSafe(settingsPath) Then
            Mod_Util.WriteWarning "P002", "LoadProjectSettings: файл общих настроек " & _
                "проекта не найден (" & settingsPath & "). Константы проекта " & _
                "({{СПИСОК_ВХОДОВ}}, {{ТАБЛИЦА_ИЗВЕСТНЫХ_GUID}}, {{СПИСОК_БАЗОВЫХ_ТИПОВ}}, " & _
                "{{ПРАВИЛО_ОБЪЕКТА}}, {{ДЕКОМПОЗИЦИЯ_ШАГОВ}}, {{ФОРМАТ_РЕЕСТРОВ}}, " & _
                "{{ПРОТОКОЛ_ШАГА}}, {{МАРКЕРЫ}}) будут помечены ""[НЕ ЗАПОЛНЕНО]""."
        Else
            Mod_Util.WriteWarning "P002", "LoadProjectSettings: файл общих настроек " & _
                "проекта пуст или не читается: " & settingsPath
        End If
        Exit Sub
    End If

    lines = Split(NormalizeLineEndings(text), vbCrLf)
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If Left$(s, 1) = "#" Then GoTo NextSettingLine   ' комментарий
        If s <> "" Then
            eqPos = InStr(s, "=")
            If eqPos > 1 Then
                keyName = Trim$(Left$(s, eqPos - 1))
                valText = Trim$(Mid$(s, eqPos + 1))
                If Left$(keyName, 2) = "{{" And Right$(keyName, 2) = "}}" Then
                    If values.Exists(keyName) Then
                        values(keyName) = CStr(values(keyName)) & vbCrLf & valText
                    Else
                        values.Add keyName, valText
                    End If
                End If
            End If
        End If
NextSettingLine:
    Next i

    Mod_Util.WriteLog "LoadProjectSettings: загружено значений настроек: " & values.Count
    Exit Sub

ErrHandler:
    Mod_Util.WriteError "P002", "LoadProjectSettings: ошибка чтения общих настроек. " & _
        "Ошибка: " & Err.Description
    Err.Clear
End Sub

' Блок 3.4 ({{ФАЙЛЫ_ДОП_ДАННЫХ}}): перечень файлов папки
' 01_Исходные данные\04_Дополнительные данные в формате
' "<имя файла> - <назначение в одну строку>". Назначение берётся из общего
' файла назначений "Назначения_доп_данных.txt" (строки вида
' "<имя файла> = <назначение>", регистр имён не значим); если назначение не
' задано - только имя файла. При отсутствии файлов - заготовка
' "Дополнительных данных для шага {{NN}} нет" (плейсхолдер {{NN}} сохраняется
' и будет заменён общим прогоном подстановки).
Private Function BuildExtraDataListBlock(ByVal root As String, ByVal nn As String) As String
    Dim folder As String
    Dim mapPath As String
    Dim fName As String
    Dim lineText As String
    Dim res As String
    Dim purpose As String
    Dim assignedDict As Object

    On Error GoTo ErrHandler

    folder = root & "\01_Исходные данные\04_Дополнительные данные"
    BuildExtraDataListBlock = ""
    If Not FolderExistsSafe(folder) Then Exit Function

    ' Карта назначений (необязательный общий файл)
    Set assignedDict = CreateObject("Scripting.Dictionary")
    assignedDict.CompareMode = vbTextCompare
    mapPath = folder & "\Назначения_доп_данных.txt"
    If FileExistsSafe(mapPath) Then
        Dim mapLines() As String
        Dim mi As Long
        Dim ms As String
        Dim meq As Long
        mapLines = Split(NormalizeLineEndings(ReadTextFile(mapPath)), vbCrLf)
        For mi = LBound(mapLines) To UBound(mapLines)
            ms = Trim$(mapLines(mi))
            meq = InStr(ms, "=")
            If meq > 1 Then
                assignedDict(Trim$(Left$(ms, meq - 1))) = Trim$(Mid$(ms, meq + 1))
            End If
        Next mi
    End If

    fName = Dir(folder & "\*")
    Do While fName <> ""
        If Left$(fName, 2) <> "~$" Then                 ' временные файлы Office
            If StrComp(fName, "Назначения_доп_данных.txt", vbTextCompare) <> 0 Then
                If Not FolderEntryIsDirectory(folder & "\" & fName) Then
                    purpose = ""
                    If assignedDict.Exists(fName) Then purpose = CStr(assignedDict(fName))
                    If res <> "" Then res = res & vbCrLf
                    If purpose <> "" Then
                        res = res & fName & " - " & purpose
                    Else
                        res = res & fName               ' только имя файла
                    End If
                End If
            End If
        End If
        fName = Dir()
    Loop

    If res = "" Then
        BuildExtraDataListBlock = "Дополнительных данных для шага {{NN}} нет"
    Else
        BuildExtraDataListBlock = res
    End If
    Exit Function

ErrHandler:
    Mod_Util.WriteError "P002", "BuildExtraDataListBlock: ошибка перечня доп. данных. " & _
        "Ошибка: " & Err.Description
    Err.Clear
End Function

' Элемент пути - папка? (без ошибок Dir по кириллическим путям)
Private Function FolderEntryIsDirectory(ByVal fullPath As String) As Boolean
    Dim attrs As Long
    FolderEntryIsDirectory = False
    On Error Resume Next
    Err.Clear
    attrs = GetAttr(fullPath)
    If Err.Number = 0 Then
        If (attrs And vbDirectory) = vbDirectory Then FolderEntryIsDirectory = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

' Строка вида "<имя файла> - <назначение>" для входного файла из папки srcRel
' (поиск по маске mask). Если файл не найден - "[НЕ НАЙДЕН: <папка>\<маска>]".
Private Function BuildInputFileLine(ByVal root As String, ByVal srcRel As String, _
                                    ByVal mask As String, ByVal purpose As String) As String
    Dim folder As String
    Dim fName As String

    folder = root & "\" & srcRel
    fName = FirstMatchFile(folder, mask)
    If fName = "" Then
        BuildInputFileLine = "[НЕ НАЙДЕН: " & srcRel & "\" & mask & "]"
        Exit Function
    End If
    BuildInputFileLine = fName & " - " & purpose
End Function

' Строка для {{КОНФИГУРАЦИЯ_ФАЙЛ}}: имя файла конфигурации через
' Mod_Util.FindConfigFile (.xml > .txt; только .pilotcfx - предупреждение).
Private Function BuildConfigFileLine(ByVal root As String) As String
    Dim cfgFolder As String
    Dim cfgFull As String
    Dim cfgName As String

    cfgFolder = root & "\01_Исходные данные\03_Конфигурация"
    cfgFull = Mod_Util.FindConfigFile(cfgFolder)
    If cfgFull = "" Then
        BuildConfigFileLine = "[НЕ НАЙДЕН: 01_Исходные данные\03_Конфигурация\*.xml или *.txt]"
        Exit Function
    End If
    cfgName = Mid$(cfgFull, InStrRev(cfgFull, "\") + 1)
    BuildConfigFileLine = cfgName & " - Конфигурация Pilot-BIM (примеры XML ""Дополнительных параметров"", справочник базовых типов)"
End Function

' Строка для {{ШАГИ_ФАЙЛ}}: путь к Шаги.txt (он же источник сетки шагов).
Private Function BuildStepsFileLine(ByVal root As String) As String
    Dim stepsFile As String

    stepsFile = root & "\04_Промты\Шаги.txt"
    If Not FileExistsSafe(stepsFile) Then
        BuildStepsFileLine = "[НЕ НАЙДЕН: 04_Промты\Шаги.txt]"
        Exit Function
    End If
    BuildStepsFileLine = "04_Промты\Шаги.txt - сетка шагов декомпозиции (формат строк NN|Имя_папки)"
End Function

' Блоки 4.1/4.2 ({{РЕЕСТР_ТИПОВ}}, {{РЕЕСТР_ПОКРЫТИЯ}}):
'  - шаг 00 - заготовка "Реестры пусты - первый шаг (P002 штатно)";
'  - шаг > 00 - содержимое накопленного файла реестра;
'  - файл отсутствует на шаге > 00 - WriteWarning (P002) и явный текст
'    "[РЕЕСТР НЕ НАЙДЕН]" (а не пустая строка).
Private Function BuildRegistryBlock(ByVal registriesPath As String, _
                                    ByVal fileName As String, _
                                    ByVal nn As String) As String
    Dim filePath As String
    Dim content As String

    BuildRegistryBlock = ""
    If StepNumToLong(nn) = 0 Then
        BuildRegistryBlock = REG_EMPTY_FIRST_STEP
        Exit Function
    End If

    filePath = registriesPath & "\" & fileName
    If Not FileExistsSafe(filePath) Then
        Mod_Util.WriteWarning "P002", Mod_Util.FormatMissingFile(fileName) & _
            " (папка 03_Результаты генерации\03_Реестры, шаг " & nn & _
            ") - подставлен текст "" & REG_NOT_FOUND & ""."
        BuildRegistryBlock = REG_NOT_FOUND
        Exit Function
    End If

    content = NormalizeLineEndings(ReadTextFile(filePath))
    If Trim$(content) = "" Then
        Mod_Util.WriteWarning "P002", "BuildRegistryBlock: " & fileName & _
            " пуст на шаге " & nn & " - подставлен текст "" & REG_NOT_FOUND & ""."
        BuildRegistryBlock = REG_NOT_FOUND
        Exit Function
    End If
    BuildRegistryBlock = content
End Function

' Блок 6 ({{СПИСОК_ТИПОВ_ШАГА}}): типы шага NN из артефакта адаптированной
' структуры (02_Подготовка схемы\Адаптированная_структура_*.xlsx, лист
' "Адаптированная_структура"; читается через GetSchemaArtifactWorkbook -
' книга открывается только для чтения и сразу закрывается).
' Колонки берутся по ЗАГОЛОВКАМ (канон: "Шаг", "Системное имя типа",
' "Видимое имя типа"; обратная совместимость со старыми книгами:
' "Тип"/"Отображаемое имя"). PRE-04: канонические заголовки плоской таблицы
' A-G - "№шага/подшага" (колонка шага) и "Категория" (бывш. "Тип"); оба
' новых имени распознаются наравне со старыми, жёстких буквенных индексов
' колонок нет. Фильтрация по колонке шага: совпадение числа
' (NN) или диапазона "NN/MM", "NN-MM". Порядок строк - как в структуре
' (без сортировки). Формат вывода: "<Видимое имя типа> | <Тип>"; системное
' имя типа в блок 6 НЕ выводится (вторая колонка - значение поля "Тип"
' адаптированной структуры, напр. "Папка", "Проект").
' Если колонки "Шаг" в структуре нет - WriteWarning и пустая строка
' (плейсхолдер получит маркер "[НЕ ЗАПОЛНЕНО]", блок 2 - маркер-страховку).
Private Function BuildStepTypeList(ByVal root As String, ByVal nn As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim openedHere As Boolean
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Long
    Dim colStep As Long
    Dim colSys As Long
    Dim colVis As Long
    Dim colType As Long
    Dim headerName As String
    Dim stepVal As String
    Dim visName As String
    Dim sysName As String
    Dim typeVal As String
    Dim res As String

    On Error GoTo ErrHandler

    BuildStepTypeList = ""

    Set wb = GetSchemaArtifactWorkbook(root, openedHere)
    If wb Is Nothing Then Exit Function

    On Error Resume Next
    Set ws = wb.Worksheets(SCHEMA_SHEET_NAME)
    If ws Is Nothing Then
        ' обратная совместимость: книги, созданные до переименования
        Err.Clear
        Set ws = wb.Worksheets("Плоская_таблица")
    End If
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Mod_Util.WriteWarning "P002", "BuildStepTypeList: в файле схемы '" & _
            wb.name & "' нет листа '" & SCHEMA_SHEET_NAME & "' - блок 6 не собран."
        GoTo CleanExit
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1

    ' 1. Индексы колонок по заголовкам (строка 1).
    ' PRE-04: новые канонические заголовки плоской таблицы A-G -
    ' "№шага/подшага" (колонка шага) и "Категория" (бывш. "Тип");
    ' старые заголовки ("Шаг", "Тип") распознаются для обратной
    ' совместимости с артефактами, собранными до PRE-04. Жёстких
    ' буквенных индексов (A-E) здесь нет - только поиск по имени.
    colStep = -1
    colSys = -1
    colVis = -1
    colType = -1
    For c = 1 To lastCol
        headerName = Trim$(CStr(ws.Cells(1, c).value))
        Select Case headerName
            Case COL_STEP, COL_STEP_NUM: colStep = c   ' "Шаг" / новый "№шага/подшага"
            Case "Системное имя типа": colSys = c
            Case "Видимое имя типа": colVis = c
            Case COL_SYS_NAME: colSys = c              ' старый заголовок "Тип"
            Case COL_CATEGORY: colSys = c              ' новый "Категория" (PRE-04)
            Case COL_VIS_NAME: colVis = c              ' "Отображаемое имя"
            Case "Тип": colType = c
        End Select
    Next c

    If colVis < 0 Or (colSys < 0 And colType < 0) Then
        Mod_Util.WriteWarning "P002", "BuildStepTypeList: в структуре не найдены " & _
            "колонки ""Видимое имя типа""/""Отображаемое имя"" и " & _
            """Системное имя типа""/""Тип""/""Категория"" - блок 6 не собран."
        GoTo CleanExit
    End If
    If colStep < 0 Then
        Mod_Util.WriteWarning "P002", "BuildStepTypeList: в структуре нет колонки """ & _
            COL_STEP & """ (фильтрация типов по шагу невозможна) - блок 6 не собран."
        GoTo CleanExit
    End If

    ' 2. Отбор строк шага NN в порядке структуры (без сортировки)
    For r = 2 To lastRow
        stepVal = Trim$(CStr(ws.Cells(r, colStep).value))
        If StepValueMatches(stepVal, nn) Then
            visName = Trim$(CStr(ws.Cells(r, colVis).value))
            If colType >= 0 Then
                typeVal = Trim$(CStr(ws.Cells(r, colType).value))
            Else
                typeVal = ""
            End If
            If typeVal = "" And colSys >= 0 Then
                ' запасной вариант: поле "Тип" пустое - берём системное имя
                typeVal = Trim$(CStr(ws.Cells(r, colSys).value))
            End If
            If visName <> "" And typeVal <> "" Then
                If res <> "" Then res = res & vbCrLf
                res = res & visName & " | " & typeVal
            End If
        End If
    Next r

CleanExit:
    ' Аварийный выход обязан закрыть книгу, которую открыли сами
    If openedHere And Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    BuildStepTypeList = res
    Exit Function

ErrHandler:
    Mod_Util.WriteError "P002", "BuildStepTypeList: ошибка сборки блока 6. " & _
        "Ошибка: " & Err.Description
    Err.Clear
    If openedHere And Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    BuildStepTypeList = res
End Function

' W-PROMPT-INVARIANTS-CODE-2: открывает самый свежий артефакт адаптированной
' структуры (02_Подготовка схемы) для чтения. Возвращает Workbook; если
' книга уже была открыта пользователем/кодом - возвращается она же
' (openedHere=False, закрывать нельзя); иначе openedHere=True и вызывающий
' обязан закрыть книгу (требование к аварийным выходам). При отсутствии
' файла/ошибке открытия - WriteWarning и Nothing.
Private Function GetSchemaArtifactWorkbook(ByVal root As String, _
                                           ByRef openedHere As Boolean) As Workbook
    Dim prepPath As String
    Dim fileName As String
    Dim fullName As String
    Dim wb As Workbook

    Set GetSchemaArtifactWorkbook = Nothing
    openedHere = False

    On Error Resume Next
    Err.Clear
    prepPath = NormPath(root) & "\02_Подготовка схемы"
    fileName = GetNewestSchemaArtifactName(prepPath)
    If fileName = "" Then
        Mod_Util.WriteWarning "P002", "GetSchemaArtifactWorkbook: артефакт " & _
            SCHEMA_ARTIFACT_MASK & " не найден в 02_Подготовка схемы - блок 6 не собран."
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    fullName = prepPath & "\" & fileName

    ' если книга уже открыта - не открываем повторно и не закрываем
    For Each wb In Workbooks
        If StrComp(wb.name, fileName, vbTextCompare) = 0 Then
            Set GetSchemaArtifactWorkbook = wb
            openedHere = False
            On Error GoTo 0
            Exit Function
        End If
    Next wb

    ' W-AV-REFACTOR, п.1: AddToMru:=False - файлы шагов не попадают в список
    ' последних файлов Excel.
    Set wb = Workbooks.Open(fullName, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    If Err.Number <> 0 Then
        Mod_Util.WriteWarning "P002", "GetSchemaArtifactWorkbook: не удалось открыть " & _
            "файл схемы '" & fileName & "'. Ошибка: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Set GetSchemaArtifactWorkbook = Nothing
        Exit Function
    End If
    On Error GoTo 0

    openedHere = True
    Set GetSchemaArtifactWorkbook = wb
End Function

' Совпадает ли значение колонки "Шаг" с номером NN: точное число ("05"),
' либо диапазон с префиксом ("05/06", "5-6"). leading zero не значим.
Private Function StepValueMatches(ByVal stepVal As String, ByVal nn As String) As Boolean
    Dim plainNN As String
    Dim plainVal As String
    Dim sepPos As Long

    StepValueMatches = False
    plainNN = CStr(StepNumToLong(nn))
    If plainNN = "" Then Exit Function
    If stepVal = "" Then Exit Function

    plainVal = CStr(StepNumToLong(stepVal))
    If plainVal = plainNN Then
        StepValueMatches = True
        Exit Function
    End If

    ' диапазоны "NN/MM" и "NN-MM": принадлежность, если NN - граница
    sepPos = 0
    If InStr(stepVal, "/") > 0 Then sepPos = InStr(stepVal, "/")
    If sepPos = 0 And InStr(stepVal, "-") > 0 Then sepPos = InStr(stepVal, "-")
    If sepPos > 1 Then
        If CStr(StepNumToLong(Left$(stepVal, sepPos - 1))) = plainNN Then StepValueMatches = True
        If CStr(StepNumToLong(Mid$(stepVal, sepPos + 1))) = plainNN Then StepValueMatches = True
    End If
End Function

' ---------------------------------------------------------------------------
' БЛОК 16: КОРРЕКТИРОВКИ (W-PROMPT-INVARIANTS-CODE, ЧАСТИ C и D)
' ---------------------------------------------------------------------------

' Удаляет из текста HTML-комментарии <!-- ... --> (многострочные включительно).
Private Function StripHtmlComments(ByVal text As String) As String
    Dim openPos As Long
    Dim closePos As Long
    Dim res As String
    Dim startPos As Long

    StripHtmlComments = text
    Do
        openPos = InStr(startPos + 1, text, "<!--")
        If openPos = 0 Then Exit Do
        closePos = InStr(openPos + 4, text, "-->")
        If closePos = 0 Then Exit Do           ' незакрытый комментарий - оставляем как есть
        res = res & Mid$(text, startPos + 1, openPos - startPos - 1)
        startPos = closePos + 3
    Loop
    res = res & Mid$(text, startPos + 1)
    StripHtmlComments = res
End Function

' Извлекает содержимое секции 16 из текста промта: всё между заголовком
' "## 16. …" и следующим заголовком верхнего уровня "## " (или концом файла).
' Возвращает "" , если секция не найдена.
Private Function ExtractSection16(ByVal promptText As String) As String
    Dim lines() As String
    Dim i As Long
    Dim s As String
    Dim inSec As Boolean
    Dim res As String

    ExtractSection16 = ""
    If promptText = "" Then Exit Function

    lines = Split(NormalizeLineEndings(promptText), vbCrLf)
    inSec = False
    For i = LBound(lines) To UBound(lines)
        s = lines(i)
        If Not inSec Then
            If Left$(Trim$(s), 6) = "## 16." Then inSec = True
        Else
            If Left$(Trim$(s), 3) = "## " Then Exit For
            If res = "" Then res = s Else res = res & vbCrLf & s
        End If
    Next i
    ExtractSection16 = res
End Function

' Валидация одной строки корректировки (ЧАСТЬ D п.6): должна начинаться с
' "- " и содержать обязательные поля ШАГ:, ТИП:, УСЛОВИЕ:, УТОЧНЕНИЕ:.
Private Function IsCorrectionLineValid(ByVal lineText As String) As Boolean
    Dim s As String
    IsCorrectionLineValid = False
    s = Trim$(lineText)
    If Left$(s, 2) <> "- " Then Exit Function
    If InStr(s, FIELD_STEP) = 0 Then Exit Function
    If InStr(s, FIELD_TYPE) = 0 Then Exit Function
    If InStr(s, FIELD_COND) = 0 Then Exit Function
    If InStr(s, FIELD_CLAR) = 0 Then Exit Function
    IsCorrectionLineValid = True
End Function

' Извлекает значение поля "ШАГ: X" из строки корректировки (до "|" или конца
' строки). Возвращает "" при отсутствии/некорректности поля.
Private Function CorrectionStepValue(ByVal lineText As String) As String
    Dim s As String
    Dim p As Long
    Dim tail As String
    Dim pipePos As Long

    CorrectionStepValue = ""
    s = lineText
    p = InStr(s, FIELD_STEP)
    If p = 0 Then Exit Function
    tail = Mid$(s, p + Len(FIELD_STEP))
    pipePos = InStr(tail, "|")
    If pipePos > 0 Then tail = Left$(tail, pipePos - 1)
    CorrectionStepValue = Trim$(tail)
End Function

' Извлекает пользовательские корректировки из произвольного текста (секции 16
' предыдущей версии промта или файла corrections.md): строки, начинающиеся
' с "- ", проходят валидацию скоупа (ЧАСТЬ D п.6): чужие шаги (ШАГ ? nn) и
' некорректные по формату - отбрасываются с логом; если после фильтрации не
' осталось ни одной корректной и были отброшенные - WriteWarning.
' Заготовка "Корректировок для шага NN нет." пользовательской не является.
Private Function CollectValidCorrections(ByVal sectionText As String, _
                                         ByVal nn As String, _
                                         ByRef droppedCount As Long) As String
    Dim cleaned As String
    Dim lines() As String
    Dim i As Long
    Dim s As String
    Dim res As String
    Dim corrStep As String

    droppedCount = 0
    CollectValidCorrections = ""
    If Trim$(sectionText) = "" Then Exit Function

    cleaned = StripHtmlComments(sectionText)
    lines = Split(NormalizeLineEndings(cleaned), vbCrLf)

    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If Left$(s, 2) = "- " Then
            If Not IsCorrectionLineValid(s) Then
                droppedCount = droppedCount + 1
                Mod_Util.WriteWarning "P002", "CollectValidCorrections: отброшена " & _
                    "некорректная по формату корректировка (нужны " - " и поля " & _
                    FIELD_STEP & " " & FIELD_TYPE & " " & FIELD_COND & " " & FIELD_CLAR & _
                    "): """ & s & """"
            Else
                corrStep = CorrectionStepValue(s)
                If StrComp(CStr(StepNumToLong(corrStep)), CStr(StepNumToLong(nn)), vbBinaryCompare) <> 0 Then
                    droppedCount = droppedCount + 1
                    Mod_Util.WriteWarning "P002", "CollectValidCorrections: отброшена " & _
                        "корректировка чужого шага: ШАГ=" & corrStep & ", ожидался " & nn
                Else
                    If res <> "" Then res = res & vbCrLf
                    res = res & s
                End If
            End If
        End If
    Next i

    If res = "" And droppedCount > 0 Then
        Mod_Util.WriteWarning "P002", "CollectValidCorrections: после фильтрации " & _
            "скоупа не осталось ни одной корректной корректировки шага " & nn & _
            " (отброшено записей: " & droppedCount & ")."
    End If

    CollectValidCorrections = res
End Function

' Содержимое секции 16 после удаления заготовки и комментариев (для проверки
' "были ли пользовательские строки").
Private Function Section16WithoutStub(ByVal sectionText As String) As String
    Dim cleaned As String
    Dim lines() As String
    Dim i As Long
    Dim s As String
    Dim res As String

    Section16WithoutStub = ""
    cleaned = StripHtmlComments(sectionText)
    lines = Split(NormalizeLineEndings(cleaned), vbCrLf)
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If s <> "" Then
            If s <> CORRECTION_STUB_TEXT Then
                If res = "" Then res = s Else res = res & vbCrLf & s
            End If
        End If
    Next i
End Function

' Сборка блока 16 (ЧАСТЬ C п.5):
'  1. предыдущая версия промта того же шага (файлы Промт_шаг_NN.md ищутся в
'     04_Промты и 03_Результаты генерации\04_Промты_готовые);
'  2. из её секции 16 извлекаются пользовательские корректировки (с валидацией
'     скоупа); дополнительно перечитывается свежий файл 04_Промты\corrections.md
'     (пользователь мог дополнить корректировки после прошлой генерации);
'  3. если пользовательских строк нет - остаётся заготовка шаблона
'     "Корректировок для шага {{NN}} нет." ({{NN}} затем превратится в NN);
'  4. если есть - заготовка заменяется списком корректировок (без дублирования).
' Вызывается ДО FillAllPlaceholders - плейсхолдер {{NN}} в тексте ещё жив.
Private Sub ApplyCorrectionsBlock(ByRef resultText As String, ByVal nn As String, ByVal root As String)
    Dim prevText As String
    Dim prevSection As String
    Dim userCorr As String
    Dim droppedPrev As Long
    Dim droppedFile As Long
    Dim corrFileText As String
    Dim corrSection As String
    Dim stubPos As Long
    Dim nextHeadPos As Long
    Dim beforeStub As String
    Dim afterStub As String
    Dim candidates(1 To 4) As String
    Dim ci As Long
    Dim foundPath As String

    On Error GoTo ErrHandler

    ' --- 1. Предыдущая версия промта шага NN --------------------------------
    candidates(1) = root & "\04_Промты\Промт_шаг_" & nn & ".md"
    candidates(2) = root & "\04_Промты\Промт_шаг_" & Right$("0" & nn, 2) & ".md"
    candidates(3) = root & "\03_Результаты генерации\04_Промты_готовые\Промт_шаг_" & nn & ".md"
    candidates(4) = root & "\03_Результаты генерации\04_Промты_готовые\Промт_шаг_" & Right$("0" & nn, 2) & ".md"
    prevText = ""
    foundPath = ""
    For ci = 1 To 4
        If FileExistsSafe(candidates(ci)) Then
            prevText = ReadTextFile(candidates(ci))
            If prevText <> "" Then
                foundPath = candidates(ci)
                Exit For
            End If
        End If
    Next ci

    userCorr = ""
    If prevText <> "" Then
        prevSection = ExtractSection16(prevText)
        If Section16WithoutStub(prevSection) <> "" Then
            ' в предыдущей версии есть что-то кроме заготовки/комментариев
            userCorr = CollectValidCorrections(prevSection, nn, droppedPrev)
            If userCorr = "" Then
                Mod_Util.WriteLog "ApplyCorrectionsBlock: в предыдущей версии промта " & _
                    "(" & foundPath & ") пользовательских корректировок шага " & nn & _
                    " после валидации не осталось - используется заготовка шаблона."
            End If
        End If
    End If

    ' --- 2. Внешний файл корректировок (04_Промты\corrections.md) ----------
    corrFileText = ReadTextFile(root & "\04_Промты\corrections.md")
    If corrFileText <> "" Then
        corrSection = ExtractSection16(corrFileText)
        If corrSection = "" Then corrSection = corrFileText   ' файл без заголовка - весь текст
        Dim fromFile As String
        fromFile = CollectValidCorrections(corrSection, nn, droppedFile)
        If fromFile <> "" Then
            If userCorr = "" Then
                userCorr = fromFile
            ElseIf InStr(userCorr, fromFile) = 0 Then
                userCorr = userCorr & vbCrLf & fromFile       ' без дублирования
            End If
        End If
    End If

    If userCorr = "" Then
        Mod_Util.WriteLog "ApplyCorrectionsBlock: корректировок шага " & nn & _
            " нет - оставлена заготовка шаблона (блок 16)."
        Exit Sub
    End If

    ' --- 3. Замена заготовки на скорректированный список --------------------
    ' Ищем строку заготовки (плейсхолдер {{NN}} ещё не подставлен).
    stubPos = InStr(resultText, CORRECTION_STUB_TEXT)
    If stubPos = 0 Then
        ' заготовка изменена вручную в шаблоне - вставляем список перед
        ' следующим заголовком "## " после начала секции 16
        Dim sec16Pos As Long
        sec16Pos = InStr(resultText, "## 16.")
        If sec16Pos = 0 Then
            Mod_Util.WriteWarning "P002", "ApplyCorrectionsBlock: секция 16 не найдена " & _
                "в шаблоне - корректировки НЕ вставлены."
            Exit Sub
        End If
        nextHeadPos = InStr(sec16Pos + 6, resultText, vbCrLf & "## ")
        If nextHeadPos = 0 Then nextHeadPos = Len(resultText) + 1
        resultText = Left$(resultText, nextHeadPos - 1) & vbCrLf & userCorr & Mid$(resultText, nextHeadPos)
        Exit Sub
    End If

    ' конец строки заготовки
    Dim lineEnd As Long
    lineEnd = InStr(stubPos, resultText, vbCrLf)
    If lineEnd = 0 Then lineEnd = Len(resultText) + 1
    beforeStub = Left$(resultText, stubPos - 1)
    afterStub = Mid$(resultText, lineEnd)
    resultText = beforeStub & userCorr & afterStub

    Mod_Util.WriteLog "ApplyCorrectionsBlock: в блок 16 вставлены пользовательские " & _
        "корректировки шага " & nn & " (предыдущая версия: " & _
        IIf(foundPath <> "", foundPath, "нет") & ")."
    Exit Sub

ErrHandler:
    Mod_Util.WriteError "P002", "ApplyCorrectionsBlock: ошибка сохранения секции 16. " & _
        "Ошибка: " & Err.Description
    Err.Clear
End Sub

' Мини-IIf по образцу Mod_Assembler/Mod_Structure (проект использует IIf из
' VBA-библиотеки; локальный дубль гарантирует работу модуля без ссылки).
Private Function IIf(cond As Boolean, byTrue As String, byFalse As String) As String
    If cond Then IIf = byTrue Else IIf = byFalse
End Function

' ---------------------------------------------------------------------------
' W-PROMPT-INVARIANTS-CODE-2: ОБЩАЯ ЦЕПОЧКА СБОРКИ ПРОМТА ШАГА
' ---------------------------------------------------------------------------

' Полный путь к файлу общих настроек проекта (project_settings.md).
' Поиск без хардкода одного места (как механизм Mod_Folders для шаблонов):
' 1) 04_Промты\Шаблоны (внутренняя папка шаблонов, куда копируется внешняя
'    папка "Шаблоны"); 2) рядом с книгой .xlsm; 3) внешняя папка "Шаблоны";
' 4) корень 04_Промты. Если ничего не найдено - возвращается путь варианта 1
'   (LoadProjectSettings сам пишет WriteWarning "файл не найден").
Private Function FindProjectSettingsPath(ByVal root As String) As String
    Dim candidates(1 To 4) As String
    Dim i As Long

    candidates(1) = root & "\04_Промты\Шаблоны\" & PROJECT_SETTINGS_FILE
    candidates(2) = root & "\" & PROJECT_SETTINGS_FILE
    candidates(3) = root & "\Шаблоны\" & PROJECT_SETTINGS_FILE
    candidates(4) = root & "\04_Промты\" & PROJECT_SETTINGS_FILE
    For i = 1 To 4
        If FileExistsSafe(candidates(i)) Then
            FindProjectSettingsPath = candidates(i)
            Exit Function
        End If
    Next i
    FindProjectSettingsPath = candidates(1)
End Function

' Убирает из текста промта незаполненные маркеры "[НЕ ЗАПОЛНЕНО: {{ТЕКУЩИЙ_ШАГ}}]"
' и им подобные (строку с маркером целиком). Используется точечно: когда
' соответствующая секция уже заполнена локальным блоком (например, сеткой
' шагов из Шаги.txt), чтобы в готовом промте не оставался служебный маркер.
Private Function RemoveUnfilledMarkerLines(ByVal text As String, ByVal phName As String) As String
    Dim marker As String
    Dim lines() As String
    Dim i As Long
    Dim res As String

    RemoveUnfilledMarkerLines = text
    marker = "[НЕ ЗАПОЛНЕНО: " & phName & "]"
    If InStr(text, marker) = 0 Then Exit Function

    lines = Split(NormalizeLineEndings(text), vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), marker) = 0 Then
            If res = "" Then res = lines(i) Else res = res & vbCrLf & lines(i)
        End If
    Next i
    RemoveUnfilledMarkerLines = res
End Function

' Блок {{ДЕКОМПОЗИЦИЯ_ШАГОВ}}: сетка шагов из Шаги.txt (формат NN|Имя_папки)
' в виде markdown-таблицы. Для Промта №0 (декомпозиция) Шаги.txt ещё нет -
' возвращается заготовка "Сетка шагов будет определена на этом шаге."
Private Function BuildDecompositionBlock(ByVal root As String) As String
    Dim steps As Collection
    Dim i As Long
    Dim nn As String
    Dim nm As String
    Dim res As String

    On Error Resume Next
    Set steps = ReadStepList()
    On Error GoTo 0
    If steps Is Nothing Then
        BuildDecompositionBlock = ""
        Exit Function
    End If
    If steps.Count = 0 Then
        BuildDecompositionBlock = "Сетка шагов будет определена на этом шаге (декомпозиция)."
        Exit Function
    End If

    res = "| Шаг | Название |" & vbCrLf & "|---|---|"
    For i = 1 To steps.Count
        nn = StepNumFromEntry(steps(i))
        nm = StepNameFromEntry(steps(i))
        res = res & vbCrLf & "| " & nn & " | " & nm & " |"
    Next i
    BuildDecompositionBlock = res
End Function

' Общая цепочка сборки промта шага по новому шаблону prompt_step_template_v1.md
' (W-PROMPT-INVARIANTS-CODE-2, ЧАСТЬ B п.4; дублирование между
' MakeNextStepPrompt и GenerateStepZeroPrompt исключено). Порядок:
'  1. загрузка шаблона из 04_Промты\Шаблоны (путь через существующий
'     механизм Mod_Folders.CopyPromptTemplates - внешняя папка "Шаблоны"
'     копируется туда при создании структуры; путь не хардкодится);
'  2. ApplyCorrectionsBlock - сохранение пользовательских корректировок из
'     предыдущей версии промта шага + corrections.md (валидация скоупа
'     IsCorrectionLineValid/CollectValidCorrections внутри; работает ДО
'     подстановки, пока {{NN}} в тексте жив);
'  3. FillAllPlaceholders - блоки 3.1-3.5 (BuildInputFileLine/
'     BuildExtraDataListBlock/BuildStepsFileLine), 4.1/4.2
'     (BuildRegistryBlock), 6 (BuildStepTypeList), константы проекта
'     (LoadProjectSettings ? project_settings.md) и маркер-страховка
'     "ЗАДАЧА НЕ ЗАДАНА" в блок 2 при пустом блоке 6;
'  4. локальные дополнения {{ТЕКУЩИЙ_ШАГ}}/{{ДЕКОМПОЗИЦИЯ_ШАГОВ}} (если
'     общие настройки их не задали) и очистка служебных маркеров.
' Возвращает собранный текст промта; Empty - при фатальной ошибке (шаблон
' не найден/не читается; вызывающий пишет P002/P004 и останавливает работу).
' hasStepTypes - результат проверки блока 6 (для логов точек входа).
Private Function BuildPromptForStep(ByVal root As String, _
                                    ByVal nn As String, _
                                    ByVal stepName As String, _
                                    ByRef hasStepTypes As Boolean) As String
    Dim templatePath As String
    Dim settingsPath As String
    Dim registriesPath As String
    Dim template As String
    Dim result As String
    Dim tmp As String
    Dim curVal As String
    Dim decVal As String

    On Error GoTo ErrHandler

    hasStepTypes = False

    ' --- 1. Шаблон prompt_step_template_v1.md из 04_Промты\Шаблоны ---------
    templatePath = root & "\04_Промты\Шаблоны\" & TEMPLATE_STEP_V1_NAME
    Mod_Util.WriteLog "[i] Поиск шаблона шага по пути: " & templatePath
    template = ReadTextFile(templatePath)
    If template = "" Then
        ' Diagnostics-before-fix: фиксируем, какие альтернативные пути
        ' проверены (внешняя папка "Шаблоны" рядом с книгой).
        Mod_Util.DebugLog "BuildPromptForStep: шаблон не найден/пуст по пути " & _
            templatePath & "; проверен альтернативный путь: " & _
            root & "\Шаблоны\" & TEMPLATE_STEP_V1_NAME
        If FileExistsSafe(root & "\Шаблоны\" & TEMPLATE_STEP_V1_NAME) Then
            template = ReadTextFile(root & "\Шаблоны\" & TEMPLATE_STEP_V1_NAME)
        End If
    End If
    If template = "" Then
        BuildPromptForStep = ""
        Exit Function
    End If
    template = NormalizeLineEndings(template)

    ' --- 2. Сохранение корректировок (блок 16) ДО подстановки --------------
    result = template
    ApplyCorrectionsBlock result, nn, root

    ' --- 3. Подстановка всех плейсхолдеров ---------------------------------
    registriesPath = root & "\03_Результаты генерации\03_Реестры"
    settingsPath = FindProjectSettingsPath(root)
    Mod_Util.DebugLog "BuildPromptForStep: общие настройки проекта - " & settingsPath
    ' ВАЖНО: ByRef-аргументы FillAllPlaceholders должны быть РАЗНЫМИ
    ' переменными (иначе в VBA один и тот же параметр, переданный по ссылке
    ' дважды, интерпретируется как one-element array - ошибка 13/449).
    Dim filled As Boolean
    Dim typesOk As Boolean
    filled = FillAllPlaceholders(result, nn, stepName, _
                                 registriesPath, root, settingsPath, _
                                 result, typesOk)
    hasStepTypes = typesOk
    If Not filled Then
        ' Блок 6 пуст: маркер-страховка "ЗАДАЧА НЕ ЗАДАНА" уже подставлен в
        ' блок 2 внутри FillAllPlaceholders (ЧАСТЬ D п.7); дополнительно -
        ' явная запись в журнал.
        Mod_Util.WriteWarning "P002", "BuildPromptForStep: блок 6 (типы шага " & _
            nn & ") пуст - в блок 2 подставлен маркер-страховка """ & _
            TASK_NOT_SET_MARKER & """."
    End If

    ' --- 4. Локальные дополнения (если общие настройки не задали) ----------
    ' {{ТЕКУЩИЙ_ШАГ}}: статус сетки шагов относительно NN.
    If InStr(result, "{{ТЕКУЩИЙ_ШАГ}}") > 0 Then
        curVal = ""
        If Trim$(CStr(LoadSettingValue(settingsPath, "{{ТЕКУЩИЙ_ШАГ}}"))) <> "" Then _
            curVal = CStr(LoadSettingValue(settingsPath, "{{ТЕКУЩИЙ_ШАГ}}"))
        If curVal = "" Then
            If StepNumToLong(nn) = 0 Then
                curVal = "Текущий шаг: 00 (декомпозиция). Шаги ещё не распределены;" & _
                         " результатом этого шага станет файл Шаги.txt."
            Else
                curVal = "Текущий шаг: " & nn & " из сетки Шаги.txt. Все шаги с номером " & _
                         "меньше " & nn & " считаются выполненными."
            End If
        End If
        tmp = result
        If FillPh(tmp, "{{ТЕКУЩИЙ_ШАГ}}", curVal, result) Then result = tmp
    End If

    ' {{ДЕКОМПОЗИЦИЯ_ШАГОВ}}: сетка шагов из Шаги.txt (markdown-таблица).
    If InStr(result, "{{ДЕКОМПОЗИЦИЯ_ШАГОВ}}") > 0 Then
        decVal = ""
        If Trim$(CStr(LoadSettingValue(settingsPath, "{{ДЕКОМПОЗИЦИЯ_ШАГОВ}}"))) <> "" Then _
            decVal = CStr(LoadSettingValue(settingsPath, "{{ДЕКОМПОЗИЦИЯ_ШАГОВ}}"))
        If decVal = "" Then decVal = BuildDecompositionBlock(root)
        tmp = result
        If FillPh(tmp, "{{ДЕКОМПОЗИЦИЯ_ШАГОВ}}", decVal, result) Then result = tmp
    End If

    ' Очистка строк со служебными маркерами незаполненности для блоков,
    ' которые мы уже заполнили локально (чтобы в промте не висел [НЕ ЗАПОЛНЕНО]).
    result = RemoveUnfilledMarkerLines(result, "{{ТЕКУЩИЙ_ШАГ}}")
    result = RemoveUnfilledMarkerLines(result, "{{ДЕКОМПОЗИЦИЯ_ШАГОВ}}")

    BuildPromptForStep = result
    Exit Function

ErrHandler:
    Mod_Util.WriteError "P002", "BuildPromptForStep: ошибка сборки промта шага " & _
        nn & ". Ошибка: " & Err.Description
    Err.Clear
    BuildPromptForStep = ""
End Function

' Читает из файла общих настроек проекта ОДНО значение плейсхолдера
' (последнее определение ключа). Используется точечно в BuildPromptForStep
' для проверки: задал ли пользователь значение {{ТЕКУЩИЙ_ШАГ}} /
' {{ДЕКОМПОЗИЦИЯ_ШАГОВ}} вручную. Пустая строка - значение не задано.
Private Function LoadSettingValue(ByVal settingsPath As String, ByVal phName As String) As String
    Dim text As String
    Dim lines() As String
    Dim i As Long
    Dim s As String
    Dim eqPos As Long
    Dim keyName As String

    LoadSettingValue = ""
    On Error GoTo ErrHandler

    text = ReadTextFile(settingsPath)
    If text = "" Then Exit Function

    lines = Split(NormalizeLineEndings(text), vbCrLf)
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If Left$(s, 1) <> "#" And s <> "" Then
            eqPos = InStr(s, "=")
            If eqPos > 1 Then
                keyName = Trim$(Left$(s, eqPos - 1))
                If keyName = phName Then
                    LoadSettingValue = Trim$(Mid$(s, eqPos + 1))
                End If
            End If
        End If
    Next i
    Exit Function

ErrHandler:
    LoadSettingValue = ""
    Err.Clear
End Function

' ---------------------------------------------------------------------------
' ПУБЛИЧЕСКИЕ ПРОЦЕДУРЫ И ФУНКЦИИ
' ---------------------------------------------------------------------------

' Читает список шагов (<корень>\04_Промты\Шаги.txt) и возвращает его
' коллекцией. Каждая запись - строка вида "NN|Имя_папки".
' W-UTF8-READ: файл читается целиком через SafeReadText (автоопределение
' UTF-8 / Windows-1251), затем разбивается на строки Split(text, vbCrLf);
' пустые строки пропускаются.
' W-CORRECTION-PACK-1, п.4.4: основной формат строки - "NN|Имя_папки"
' (разделитель - pipe). Устаревший вариант "NN = Имя_папки" принимается с
' предупреждением в логе; каждая запись нормализуется к виду "NN|Имя_папки".
' При отсутствии файла: ошибка P001, MsgBox, пустая коллекция.
' При ошибке чтения: [WARN] P006 в лог, пустая коллекция без прерывания.
Public Function ReadStepList() As Collection
    Dim col As Collection
    Dim root As String
    Dim text As String
    Dim lines() As String
    Dim i As Long
    Dim s As String
    Dim posPipe As Long
    Dim posEq As Long
    Dim nn As String
    Dim nm As String

    Set col = New Collection

    On Error GoTo ErrHandler

    Dim filePath As String
    ' W-PROMPT-PATHS-REFACTOR: нормализация корневого пути (защита от
    ' двойных слешей, если GetProjectRoot вернёт путь с завершающим "\")
    root = NormPath(Mod_Util.GetProjectRoot)
    filePath = root & "\04_Промты\Шаги.txt"

    ' файла нет ? P002, MsgBox, пустая коллекция
    If Len(Dir(filePath, vbNormal)) = 0 Then
        Mod_Util.WriteWarning "P002", "ReadStepList: файл Шаги.txt упразднён (W-CLEANUP-STEP0). " & _
            "Шаги определяются автоматически по листу Адаптированная_структура."
        Set ReadStepList = col
        Exit Function
    End If

    ' W-UTF8-READ: чтение всего файла с автоопределением кодировки
    text = SafeReadText(filePath)
    If text = "" Then
        ' SafeReadText уже записал [WARN] в лог - здесь только предупреждаем
        ' пользователя и возвращаем пустую коллекцию (без прерывания).
        MsgBox "Предупреждение P006: не удалось прочитать файл Шаги.txt " & _
               "(ошибки кодировки отсутствуют, но файл пуст или недоступен)." & _
               vbCrLf & filePath, vbExclamation, "Промты"
        Set ReadStepList = col
        Exit Function
    End If

    ' Разбить текст на строки; формат строки: "NN|Имя_папки"
    ' (W-UTF8-READ: ADODB.Stream сохраняет переводы строки файла как есть,
    ' поэтому сначала нормализуем CRLF/LF/CR к vbCrLf).
    lines = Split(NormalizeLineEndings(text), vbCrLf)

    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        ' пропуск пустых строк
        If s <> "" Then
            posPipe = InStr(s, "|")
            If posPipe > 0 Then
                ' канонический формат (W-CORRECTION-PACK-1, п.4.4): NN|Имя_папки
                nn = Trim(Left(s, posPipe - 1))
                nm = Trim(Mid(s, posPipe + 1))
                If nn <> "" And nm <> "" Then col.Add nn & "|" & nm
            Else
                ' устаревший формат "NN = Имя_папки" - принимаем с
                ' предупреждением, чтобы старые Шаги.txt не ломали работу
                posEq = InStr(s, "=")
                If posEq > 0 Then
                    nn = Trim(Left(s, posEq - 1))
                    nm = Trim(Mid(s, posEq + 1))
                    If nn <> "" And nm <> "" Then
                        Mod_Util.WriteWarning "P001", "ReadStepList: строка " & _
                            (i + 1) & " Шаги.txt использует разделитель ""="" вместо ""|"": """ & _
                            s & """. Обновите файл до формата ""NN|Имя_папки""."
                        col.Add nn & "|" & nm
                    End If
                End If
            End If
        End If
    Next i

    Set ReadStepList = col
    Exit Function

ErrHandler:
    ' Любая ошибка чтения - [WARN] в лог, пустая коллекция без прерывания.
    Mod_Util.WriteWarning "P002", "ReadStepList: ошибка чтения устаревшего файла Шаги.txt. Ошибка: " & Err.Description
    MsgBox "Предупреждение P002: не удалось прочитать устаревший файл Шаги.txt." & vbCrLf & Err.Description, _
        vbExclamation, "Промты"
    Err.Clear
    Set ReadStepList = col
End Function

' Возвращает номер текущего шага - последней существующей папки Шаг_*
' в 03_Результаты генерации\01_По шагам\ согласно порядку из Шаги.txt.
' Если папок нет - возвращает "".
Public Function ReadCurrentStep() As String
    Dim steps As Collection
    Dim root As String
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

    ' W-PROMPT-PATHS-REFACTOR: нормализация корневого пути (защита от
    ' двойных слешей, если GetProjectRoot вернёт путь с завершающим "\")
    root = NormPath(Mod_Util.GetProjectRoot)
    stepsPath = root & "\03_Результаты генерации\01_По шагам"
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
' текущим). Если текущего шага нет - первый шаг из Шаги.txt.
' Если следующий не найден - возвращает "".
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

    ' текущего шага нет - возвращаем первый шаг из Шаги.txt
    If current = "" Then
        ReadNextStep = StepNumFromEntry(steps(1))
        Exit Function
    End If

    ' следующий - запись, идущая сразу за текущей
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

' W-ANALYTIC-05-FIX-2: добавлены параметры excludeAssigned (исключить строки с
' уже заполненным «№шага/подшага» - итеративная разбивка 00_3 подаёт ИИ только
' нераспределённые типы) и outRowCount (число выгруженных строк данных - защита
' от генерации промта с пустой таблицей). Все новые параметры Optional - старые
' вызовы (MakeNextStepPrompt_FromSheet, 00_2) не затронуты.
Private Function BuildSchemaCsvFragment(ByVal ws As Worksheet, _
                                        Optional ByVal filterStep As String = "", _
                                        Optional ByVal maxTypes As Long = 0, _
                                        Optional ByVal excludeAssigned As Boolean = False, _
                                        Optional ByRef outRowCount As Long) As String
    Dim lastRow As Long
    Dim r As Long
    Dim colStep As Long, colVis As Long, colCat As Long, colPat As Long, colIn As Long, colOut As Long, colAnom As Long
    Dim res As String
    Dim rowCount As Long
    Dim include As Boolean
    Dim stepVal As String
    Dim visVal As String, catVal As String, patVal As String, inVal As String, outVal As String, anomVal As String

    outRowCount = 0

    ' Поиск канонических колонок плоской таблицы
    colStep = FindColumn(ws, "№шага/подшага")
    colVis = FindColumn(ws, "Отображаемое имя")
    colCat = FindColumn(ws, "Категория")
    colPat = FindColumn(ws, "Паттерн")
    colIn = FindColumn(ws, "Входит в")
    colOut = FindColumn(ws, "Состоит из")
    colAnom = FindColumn(ws, "Аномалия")

    If colStep = 0 Then
        BuildSchemaCsvFragment = ""
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, colStep).End(xlUp).row

    ' Формирование заголовка CSV (разделитель точка с запятой, все поля в кавычках для надёжности)
    res = """№шага/подшага"";""Отображаемое имя"";""Категория"";""Паттерн"";""Входит в"";""Состоит из"";""Аномалия""" & vbCrLf

    rowCount = 0
    For r = 2 To lastRow
        stepVal = Trim$(CStr(ws.Cells(r, colStep).value))

        include = True
        If filterStep <> "" Then
            ' Точечный выбор строк конкретного шага (промт шага v1)
            If stepVal <> filterStep Then include = False
        ElseIf excludeAssigned Then
            ' Разбивка на шаги (00_3): исключаем типы, уже распределённые по шагам
            If stepVal <> "" Then include = False
        End If

        If include Then
            visVal = Trim$(CStr(ws.Cells(r, colVis).value))
            catVal = Trim$(CStr(ws.Cells(r, colCat).value))
            patVal = Trim$(CStr(ws.Cells(r, colPat).value))
            inVal = Trim$(CStr(ws.Cells(r, colIn).value))
            outVal = Trim$(CStr(ws.Cells(r, colOut).value))
            anomVal = Trim$(CStr(ws.Cells(r, colAnom).value))

            ' Экранирование внутренних кавычек (двойные кавычки) и оборачивание полей
            res = res & """" & Replace(stepVal, """", """""") & """;" & _
                  """" & Replace(visVal, """", """""") & """;" & _
                  """" & Replace(catVal, """", """""") & """;" & _
                  """" & Replace(patVal, """", """""") & """;" & _
                  """" & Replace(inVal, """", """""") & """;" & _
                  """" & Replace(outVal, """", """""") & """;" & _
                  """" & Replace(anomVal, """", """""") & """" & vbCrLf

            rowCount = rowCount + 1
            If maxTypes > 0 And rowCount >= maxTypes Then
                Exit For
            End If
        End If
    Next r

    outRowCount = rowCount
    BuildSchemaCsvFragment = res
End Function

' Формирует промпт для выполнения следующего шага (разделы 9 и 13).
' W-PROMPT-INVARIANTS-CODE-2, ЧАСТЬ A: старая логика сборки заменена на
' общую цепочку BuildPromptForStep (ApplyCorrectionsBlock ?
' FillAllPlaceholders) по новому шаблону prompt_step_template_v1.md.
' Заменённые старые вызовы (удалены из тела процедуры):
'  - ReadTextFile(registriesPath\Реестр_типов.txt / _покрытия.txt) +
'    WriteWarning "подставлена пустая строка"  ? BuildRegistryBlock
'    (внутри FillAllPlaceholders; шаг 00 - "Реестры пусты…", отсутствие
'    файла на шаге >00 - WriteWarning P002 + "[РЕЕСТР НЕ НАЙДЕН]");
'  - GetSchemaFragment() + ReplaceIfPresent "{SCHEMA}" ? блок 6 собирает
'    BuildStepTypeList (фильтрация по шагу NN прямо из артефакта структуры),
'    {{СТРУКТУРА_ФАЙЛ}} - BuildInputFileLine;
'  - BuildExtraDataBlock(root) & template + ReplaceIfPresent "{NN}",
'    "{STEP_NAME}", "{REGISTRY_TYPES}", "{REGISTRY_COVERAGE}", "{SCHEMA}" ?
'    FillAllPlaceholders (плейсхолдеры {{...}} нового шаблона);
'  - сохранение в 04_Промты_готовые ? сохранение в 04_Промты\Промт_шаг_NN.md
'    (как требует задание; копия в 04_Промты_готовые оставлена для
'    обратной совместимости MakeFixPrompt/пользовательских ссылок).
Public Sub MakeNextStepPrompt()
    On Error GoTo ErrHandler
    Call MakeNextStepPrompt_FromSheet
    Exit Sub
ErrHandler:
    Mod_Util.WriteError "P002", "MakeNextStepPrompt: " & Err.Description
    MsgBox "Oshibka P002: " & Err.Description, vbCritical, "Promt sleduyushego shaga"
    Err.Clear
End Sub

Private Sub MakeNextStepPrompt_FromSheet()
    On Error GoTo ErrHandler

    Mod_Util.WriteLog "MakeNextStepPrompt_FromSheet: === НАЧАЛО ==="

    Dim root As String
    root = NormPath(Mod_Util.GetProjectRoot())
    If root = "" Then
        Mod_Util.WriteError "P002", "MakeNextStepPrompt_FromSheet: корень не задан."
        MsgBox "Ошибка P002: корень проекта не задан.", vbCritical, "Следующий шаг"
        Exit Sub
    End If

    Dim pathCfg As String: pathCfg = root & "\02_Подготовка данных\01_Схема"
    Dim pathDesc As String: pathDesc = root & "\02_Подготовка данных\02_Описание"
    Dim pathConfig As String: pathConfig = root & "\02_Подготовка данных\03_Конфигурация"
    Dim pathAdd As String: pathAdd = root & "\02_Подготовка данных\04_Дополнительные данные"
    Dim pathSteps As String: pathSteps = Mod_Util.GetStepsFolder()

    ' --- 1. Проверка состояния ---
    Dim stepsState As String: stepsState = Mod_Util.GetSettingText("StepsState", "Нет")
    If StrComp(Trim$(stepsState), "Да", vbTextCompare) <> 0 Then
        Mod_Util.WriteError "P002", "MakeNextStepPrompt_FromSheet: Разбивка на шаги <> Да."
        MsgBox "Ошибка P002: настройка 'Разбивка на шаги' не включена.", vbCritical, "Следующий шаг"
        Exit Sub
    End If

    ' --- 2. Проверка папок ---
    Dim cfgFiles As Variant: cfgFiles = GetFileListArray(pathCfg)
    If Not FolderExistsSafe(pathCfg) Or Not IsArray(cfgFiles) Or UBound(cfgFiles) < LBound(cfgFiles) Then
        Mod_Util.WriteError "P002", "MakeNextStepPrompt_FromSheet: папка 01_Схема отсутствует или пуста."
        MsgBox "Ошибка P002: папка 01_Схема отсутствует или пуста.", vbCritical, "Следующий шаг"
        Exit Sub
    End If

    Dim descFiles As Variant: descFiles = GetFileListArray(pathDesc)
    If Not FolderExistsSafe(pathDesc) Or Not IsArray(descFiles) Or UBound(descFiles) < LBound(descFiles) Then
        Mod_Util.WriteError "P002", "MakeNextStepPrompt_FromSheet: папка 02_Описание отсутствует или пуста."
        MsgBox "Ошибка P002: папка 02_Описание отсутствует или пуста.", vbCritical, "Следующий шаг"
        Exit Sub
    End If

    Dim configFiles As Variant: configFiles = GetFileListArray(pathConfig)
    If Not FolderExistsSafe(pathConfig) Or Not IsArray(configFiles) Or UBound(configFiles) < LBound(configFiles) Then
        Mod_Util.WriteError "P002", "MakeNextStepPrompt_FromSheet: папка 03_Конфигурация отсутствует или пуста."
        MsgBox "Ошибка P002: папка 03_Конфигурация отсутствует или пуста.", vbCritical, "Следующий шаг"
        Exit Sub
    End If

    ' --- 2.1. Проверка папки 04_Дополнительные данные (P006) ---
    Dim warnP006 As String: warnP006 = ""
    If Not FolderExistsSafe(pathAdd) Then
        Mod_Util.WriteWarning "P006", "MakeNextStepPrompt_FromSheet: папка 04_Дополнительные данные отсутствует."
        warnP006 = "Папка 04_Дополнительные данные отсутствует."
    Else
        Dim addFiles As Variant: addFiles = GetFileListArray(pathAdd)
        If Not IsArray(addFiles) Or UBound(addFiles) < LBound(addFiles) Then
            Mod_Util.WriteWarning "P006", "MakeNextStepPrompt_FromSheet: папка 04_Дополнительные данные пуста."
            warnP006 = "Папка 04_Дополнительные данные пуста."
        End If
    End If

    ' --- 3. Чтение курсоров ---
    Dim lastGen As String: lastGen = Mod_Util.GetSettingText("LastGeneratedStepNumber", "нет")
    If Trim$(lastGen) = "" Then lastGen = "нет"
    Dim lastStep As String: lastStep = Mod_Util.GetSettingText("LastStepNumber", "нет")
    If Trim$(lastStep) = "" Then lastStep = "нет"

    If lastGen = lastStep Then
        Mod_Util.WriteLog "MakeNextStepPrompt_FromSheet: все шаги обработаны (lastGen = lastStep)"
        MsgBox "Все шаги обработаны.", vbInformation, "Следующий шаг"
        Exit Sub
    End If

    ' --- 5. Поиск nextStep ---
    Dim schemaFolder As String: schemaFolder = pathCfg
    Dim schemaFile As String: schemaFile = GetNewestSchemaArtifactName(schemaFolder)
    
    If schemaFile = "" Then
        Mod_Util.WriteError "P002", "MakeNextStepPrompt_FromSheet: файл схемы не найден."
        MsgBox "Ошибка P002: файл схемы не найден.", vbCritical, "Следующий шаг"
        Exit Sub
    End If
    
    Dim schemaWb As Workbook
    Dim wasOpen As Boolean: wasOpen = False
    On Error Resume Next
    Set schemaWb = Workbooks(schemaFile)
    On Error GoTo ErrHandler
    
    If schemaWb Is Nothing Then
        Set schemaWb = Workbooks.Open(schemaFolder & "\" & schemaFile, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    Else
        wasOpen = True
    End If
    
    If schemaWb Is Nothing Then
        Mod_Util.WriteError "P002", "MakeNextStepPrompt_FromSheet: не удалось открыть файл схемы."
        MsgBox "Ошибка P002: не удалось открыть файл схемы.", vbCritical, "Следующий шаг"
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = schemaWb.Worksheets(Mod_Structure.FLAT_SHEET_NAME)
    If ws Is Nothing Then
        If Not wasOpen Then schemaWb.Close SaveChanges:=False
        Mod_Util.WriteError "P002", "MakeNextStepPrompt_FromSheet: лист Адаптированная_структура не найден."
        MsgBox "Ошибка P002: лист Адаптированная_структура не найден.", vbCritical, "Следующий шаг"
        Exit Sub
    End If
    
    Dim colStepNum As Long: colStepNum = FindColumn(ws, COL_STEP_NUM)
    If colStepNum = 0 Then
        If Not wasOpen Then schemaWb.Close SaveChanges:=False
        Mod_Util.WriteError "P002", "MakeNextStepPrompt_FromSheet: колонка №шага/подшага не найдена."
        MsgBox "Ошибка P002: колонка №шага/подшага не найдена.", vbCritical, "Следующий шаг"
        Exit Sub
    End If
    
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, colStepNum).End(xlUp).row
    Dim nextStep As String: nextStep = ""
    Dim r As Long, cellVal As String
    Dim candidateFolder As String
    
    For r = 2 To lastRow
        cellVal = Trim$(CStr(ws.Cells(r, colStepNum).value))
        If Len(cellVal) > 0 Then
            If IsStepGreater(cellVal, lastGen) Then
                candidateFolder = pathSteps & "\Шаг_" & cellVal
                If Not FolderExistsSafe(candidateFolder) Then
                    If nextStep = "" Or IsStepGreater(nextStep, cellVal) Then
                        nextStep = cellVal
                    End If
                End If
            End If
        End If
    Next r
    
    If Len(nextStep) = 0 Then
        If Not wasOpen Then schemaWb.Close SaveChanges:=False
        Mod_Util.WriteError "P002", "MakeNextStepPrompt_FromSheet: следующий шаг не найден."
        MsgBox "Ошибка P002: следующий шаг не найден после " & lastGen, vbCritical, "Следующий шаг"
        Exit Sub
    End If

    ' --- 7. Действие и генерация промта ---
    Dim stepFolder As String: stepFolder = pathSteps & "\Шаг_" & nextStep
    EnsureFolder stepFolder
    Mod_Util.WriteLog "MakeNextStepPrompt_FromSheet: создана папка " & stepFolder

    ' === НОВАЯ ЛОГИКА: Генерация промта шага ===
    Dim csvContent As String
    csvContent = BuildSchemaCsvFragment(ws, nextStep, 0)
    SaveTextFileUtf8 stepFolder & "\Структура_шага.csv", csvContent
    
    Dim templatePath As String
    templatePath = root & "\04_Промты\Шаблоны\prompt_step_template_v1.md"
    Dim template As String
    template = ReadTextFile(templatePath)
    
    Dim result As String
    result = template
    
    ' Поиск {{ИМЯ_ШАГА}} для nextStep
    Dim stepName As String: stepName = ""
    Dim colVisName As Long: colVisName = FindColumn(ws, COL_VIS_NAME)
    If colVisName > 0 And colStepNum > 0 Then
        Dim rName As Long
        For rName = 2 To lastRow
            If Trim$(CStr(ws.Cells(rName, colStepNum).value)) = nextStep Then
                stepName = Trim$(CStr(ws.Cells(rName, colVisName).value))
                Exit For
            End If
        Next rName
    End If
    
    Dim descFile As String
    descFile = FirstMatchFile(pathDesc, "Выжимка_Описание.*")
    If descFile = "" Then descFile = "нет"
    
    Dim configFile As String
    configFile = Mod_Util.FindConfigFile(pathConfig)
    If configFile = "" Then configFile = "нет"
    
    Dim addFilesList As String
    addFilesList = ListFolderFiles(pathAdd)
    If addFilesList = "" Then addFilesList = "нет"
    
    result = ReplaceIfPresent(result, "{{NN}}", nextStep)
    result = ReplaceIfPresent(result, "{{ИМЯ_ШАГА}}", stepName)
    result = ReplaceIfPresent(result, "{{СТРУКТУРА_ФАЙЛ}}", "Структура_шага.csv")
    result = ReplaceIfPresent(result, "{{ОПИСАНИЕ_ФАЙЛ}}", descFile)
    result = ReplaceIfPresent(result, "{{КОНФИГУРАЦИЯ_ФАЙЛ}}", configFile)
    result = ReplaceIfPresent(result, "{{ФАЙЛЫ_ДОП_ДАННЫХ}}", addFilesList)
    
    SaveTextFileUtf8 stepFolder & "\Промт_шага.md", result
    ' ==============================================

    ' 3. Закрыть книгу схемы, если она была открыта этой процедурой
    If Not wasOpen Then
        schemaWb.Close SaveChanges:=False
    End If

    UpdateSetting "Номер последнего сгенерированного шага", nextStep
    Mod_Util.WriteLog "MakeNextStepPrompt_FromSheet: обновлён счётчик LastGeneratedStepNumber = " & nextStep

    Dim successMsg As String
    successMsg = "Определён следующий шаг: " & nextStep & vbCrLf
    successMsg = successMsg & "Создана папка: " & stepFolder & vbCrLf
    successMsg = successMsg & "Сгенерирован промт: Промт_шага.md" & vbCrLf
    successMsg = successMsg & "Обновлён счётчик: «Номер последнего сгенерированного шага» = " & nextStep
    
    If warnP006 <> "" Then
        successMsg = successMsg & vbCrLf & vbCrLf & "Предупреждение: " & warnP006
        MsgBox successMsg, vbExclamation, "Следующий шаг"
    Else
        MsgBox successMsg, vbInformation, "Следующий шаг"
    End If

    Mod_Util.WriteLog "MakeNextStepPrompt_FromSheet: === ЗАВЕРШЕНО УСПЕШНО ==="
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not schemaWb Is Nothing And Not wasOpen Then schemaWb.Close SaveChanges:=False
    On Error GoTo 0
    Mod_Util.WriteError "P002", "MakeNextStepPrompt_FromSheet: " & Err.Description
    MsgBox "Ошибка P002: " & Err.Description, vbCritical, "Следующий шаг"
    Err.Clear
End Sub

' УСТАРЕЛО (W-CLEANUP-STEP0): Шаги.txt упразднён, Промт №0 формируется через MakeNextStepPrompt.
' ВНИМАНИЕ (W7.1): Процедура заменена на заглушку.
' Актуальная логика формирования промта шага 0 теперь должна проходить
' через MakeNextStepPrompt (при пустом списке шагов) или использовать
' архивный шаблон из prompts/archive/prompt_step_00_template.md
Public Sub GenerateStepZeroPrompt()
    MsgBox "Процедура GenerateStepZeroPrompt заменена." & vbCrLf & _
           "Используйте кнопку 'Сделать промт следующего шага' или " & _
           "обратитесь к актуальному шаблону в prompts/archive/.", _
           vbInformation, "GenerateStepZeroPrompt"
    
    ' TODO: При необходимости восстановить логику чтения prompt_step_00_template.md
    ' и подстановки {SCHEMA_FILENAME}, {DESCRIPTION_FILENAME}, {CONFIG_FILENAME}
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
        MsgBox "Ошибка P003: текущий шаг не определён - в 03_Результаты генерации\01_По шагам " & _
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

    ' фиксированное имя последней сводной; если его нет - последний xlsx по дате
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
    ' W-AV-REFACTOR, п.1: AddToMru:=False - см. комментарий выше (Mod_Assembler).
    Set wb = Workbooks.Open(finalFile, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    On Error GoTo ErrHandler
    If wb Is Nothing Then
        Mod_Util.WriteError "P003", "MakeFixPrompt: не удалось открыть сводную: " & finalFile
        MsgBox "Ошибка P003: не удалось открыть файл сводной:" & vbCrLf & finalFile, _
               vbCritical, "Промт исправления"
        Exit Sub
    End If

    ' 3. Лист "Ошибки"
    errorsText = ""
    If Not ReadErrorsSheetText(wb, errorsText) Then
        Mod_Util.WriteError "P003", "MakeFixPrompt: в сводной не найден лист 'Ошибки': " & finalFile
        wb.Close SaveChanges:=False
        MsgBox "Ошибка P003: в файле сводной не найден лист 'Ошибки'." & vbCrLf & finalFile, _
               vbCritical, "Промт исправления"
        Exit Sub
    End If
    wb.Close SaveChanges:=False

    ' 4. Ошибок нет - ничего не формируем (раздел 13)
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

    ' 6. Оба реестра (при отсутствии - пустая строка и предупреждение)
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
    ' 8. Подстановка плейсхолдеров (согласно Spec_Dispatcher.md, Раздел 13)
    result = template
    result = ReplaceIfPresent(result, "{NN}", current)
    result = ReplaceIfPresent(result, "{STEP_NAME}", current) ' Имя папки шага совпадает с номером в текущей логике
    result = ReplaceIfPresent(result, "{ERRORS}", errorsText)
    result = ReplaceIfPresent(result, "{REGISTRY_TYPES}", regTypes)
    result = ReplaceIfPresent(result, "{REGISTRY_COVERAGE}", regCoverage)
    result = ReplaceIfPresent(result, "{SCHEMA}", schema)

    ' 9. Сохранение готового промта
    EnsureFolder readyPath
    SaveTextFileUtf8 readyPath & "\Промт_исправление_шаг_" & current & ".md", result

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

' Разовая инициализация массива известных плейсхолдеров.
Private Sub EnsureKnownPh()
    If SPLIT_KNOWN_PH_READY Then Exit Sub
    SPLIT_KNOWN_PH = Split(TEMPLATE_PLACEHOLDERS, ";")
    SPLIT_KNOWN_PH_READY = True
End Sub


' ---------------------------------------------------------------------------
' W7.1-A: GetPromptProfile - маршрутизация профиля промта №0.3 по номеру
' последнего шага (lastStep). Чистая функция: без MsgBox, без записи в
' журнал, без обращения к книге/файлам. Возвращает строго один из
' "профилей": "Папки" / "Паттерны" / "Документы" либо "" (пустое или
' невалидное значение lastStep).
'   Маршрутизация (Spec_Optimization.md, W7.2-NOTE-2; форма допустимых
'   значений - prompt_step_00_3_template.md, раздел "Валидация
'   {LAST_STEP_NUMBER}"):
'     ""                    -> ""        (пусто - невалидно)
'     "нет"                 -> "Папки"   (Шаг 0)
'     "00.i"                -> "Папки"   (продолжение Шага 0, i >= 1)
'     "0", "00"             -> "Паттерны" (Шаг 1 с 01.1)
'     "01.i", "01.i.k"      -> "Паттерны" (продолжение Шага 1, i,k >= 1)
'     "01"                  -> "Документы" (Шаг 1 завершён -> Шаг 2)
'     "NN", NN >= 2         -> "Документы" (продолжение; NN - две и более
'                                           цифры, канон "02".."99")
'     невалидное значение   -> ""
'   Невалидным считаются: литерал "{LAST_STEP_NUMBER}", "01.15.2.1"
'   (4+ сегмента), буквы и прочие формы, не перечисленные выше.
'   Ответственность вызывающего: profile = "" -> MsgBox + Exit Sub
'   (в GetPromptProfile не реализуется, п. 3.6 задачи W7.1-A).
' ---------------------------------------------------------------------------
Public Function GetPromptProfile(ByVal lastStep As String) As String
    Dim s As String
    Dim parts As Variant
    Dim nParts As Long
    Dim head As String

    GetPromptProfile = ""
    s = Trim$(lastStep)
    If s = "" Then Exit Function                       ' пусто -> невалидно (W7.2-NOTE-2)

    If StrComp(s, "нет", vbTextCompare) = 0 Then
        GetPromptProfile = "Папки"
        Exit Function
    End If

    If StrComp(s, "{LAST_STEP_NUMBER}", vbTextCompare) = 0 Then Exit Function

    parts = Split(s, ".")
    nParts = UBound(parts) + 1
    If nParts < 1 Or nParts > 3 Then Exit Function     ' 4+ сегмента - невалидно

    head = parts(0)
    Select Case head
        Case "0", "00"
            If nParts = 1 Then
                GetPromptProfile = "Паттерны"          ' 0 / 00
            ElseIf head = "00" And IsPositiveIntPart(parts(1)) Then
                GetPromptProfile = "Папки"             ' 00.i
            End If
        Case "01"
            If nParts = 1 Then
                GetPromptProfile = "Документы"         ' 01 завершён
            ElseIf IsPositiveIntPart(parts(1)) And _
                   (nParts = 2 Or IsPositiveIntPart(parts(2))) Then
                GetPromptProfile = "Паттерны"          ' 01.i / 01.i.k
            End If
        Case Else
            If nParts = 1 Then
                If IsPlainStepNumber(head) Then
                    GetPromptProfile = "Документы"     ' NN >= 02 (канон)
                ElseIf IsPositiveIntPart(head) And CLng(head) >= 2 Then
                    ' W-ANALYTIC-06F: нормализация Excel-формы "3".."9" (утерян ведущий ноль)
                    GetPromptProfile = "Документы"
                End If
            End If
    End Select
End Function

' Вспомогательная для GetPromptProfile (W7.1-A): строка - положительное
' целое без знака/пробелов/ведущих нулей ("0" невалиден, "1", "15" -
' корректные сегменты подшага i / куска k).
Private Function IsPositiveIntPart(ByVal seg As String) As Boolean
    Dim i As Long
    Dim ch As String
    IsPositiveIntPart = False
    If Len(seg) = 0 Then Exit Function
    For i = 1 To Len(seg)
        ch = Mid$(seg, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    If seg = "0" Then Exit Function           ' подшаг 0 недопустим
    If Len(seg) > 1 And Left$(seg, 1) = "0" Then Exit Function ' 01, 007
    IsPositiveIntPart = True
End Function

' Вспомогательная для GetPromptProfile (W7.1-A): номер шага NN >= 2 -
' только цифры, не менее двух разрядов в канонической форме "NN"
' (ведущий ноль допустим: "02", "03" - см. Шаги.txt/папки Шаг_NN;
' односимвольные "2"/"3" невалидны; "00"/"01" обрабатываются отдельно).
Private Function IsPlainStepNumber(ByVal nn As String) As Boolean
    Dim i As Long
    Dim ch As String
    IsPlainStepNumber = False
    If Len(nn) < 2 Then Exit Function
    For i = 1 To Len(nn)
        ch = Mid$(nn, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    ' числовое значение должно быть >= 2 ("00"/"01" сюда не попадают -
    ' они обработаны в Select Case отдельно, но защита от прочих форм)
    If CLng(nn) < 2 Then Exit Function
    IsPlainStepNumber = True
End Function

' ============================================================================
' W7.2: Формирование Промта №0.2 (Паттерны)
' ============================================================================
Public Sub BuildPrompt02Patterns()
    Dim root As String
    Dim promptsPath As String
    Dim templatePath As String
    Dim template As String
    Dim result As String
    Dim regPath As String
    Dim regTypes As String
    Dim regCoverage As String

    On Error GoTo ErrHandler

    root = NormPath(Mod_Util.GetProjectRoot)
    promptsPath = root & "\04_Промты"
    templatePath = promptsPath & "\Шаблоны\prompt_step_00_2_template.md"
    regPath = root & "\03_Результаты генерации\03_Реестры"

    template = ReadTextFile(templatePath)
    If template = "" Then
        Mod_Util.WriteError "P002", "BuildPrompt02Patterns: не найден шаблон prompt_step_00_2_template.md"
        MsgBox "Ошибка P002: не найден шаблон prompt_step_00_2_template.md в папке 04_Промты\Шаблоны.", vbCritical, "Промт №0.2"
        Exit Sub
    End If

    result = template
    ' Подстановка плейсхолдеров из листа "Паттерны"
    result = ReplaceIfPresent(result, "{{СПИСОК_ИЗВЕСТНЫХ_ПАТТЕРНОВ}}", Mod_Patterns.GetKnownPatterns())
    result = ReplaceIfPresent(result, "{{ИСКЛЮЧЕНИЯ_ПАТТЕРНОВ}}", Mod_Patterns.GetPatternExceptions())
    
    ' Подстановка плейсхолдеров из реестров (W7.2-NOTE-3)
    regTypes = ReadTextFile(regPath & "\Реестр_типов.txt")
    regCoverage = ReadTextFile(regPath & "\Реестр_покрытия.txt")
    result = ReplaceIfPresent(result, "{{СПИСОК_БАЗОВЫХ_ТИПОВ}}", regTypes)
    result = ReplaceIfPresent(result, "{{ТАБЛИЦА_ИЗВЕСТНЫХ_GUID}}", regCoverage)

    EnsureFolder promptsPath
    SaveTextFileUtf8 promptsPath & "\Промт_шаг_00_2_Паттерны.md", result

    Mod_Util.WriteLog "BuildPrompt02Patterns: сформирован Промт_шаг_00_2_Паттерны.md"
    MsgBox "Сформирован промт:" & vbCrLf & promptsPath & "\Промт_шаг_00_2_Паттерны.md", vbInformation, "Промт №0.2"
    Exit Sub

ErrHandler:
    Mod_Util.WriteError "P002", "BuildPrompt02Patterns: ошибка. " & Err.Description
    MsgBox "Ошибка P002: не удалось сформировать промт №0.2." & vbCrLf & Err.Description, vbCritical, "Промт №0.2"
    Err.Clear
End Sub

' ============================================================================
' W7.2: Формирование Промта №0.3 (Профиль)
' ============================================================================
Public Sub BuildPrompt03Profile()
    Dim root As String
    Dim promptsPath As String
    Dim templatePath As String
    Dim template As String
    Dim result As String
    Dim lastStep As String
    Dim profile As String
    Dim schemaFolder As String, schemaFile As String
    Dim descFolder As String, descFile As String
    Dim extraFolder As String, extraList As String
    Dim limitTypes As Long
    Dim fName As String, bestFile As String, bestDate As Date

    On Error GoTo ErrHandler

    root = NormPath(Mod_Util.GetProjectRoot)
    promptsPath = root & "\04_Промты"
    templatePath = promptsPath & "\Шаблоны\prompt_step_00_3_template.md"

    lastStep = ReadCurrentStep()
    profile = GetPromptProfile(lastStep)
    
    If profile = "" Then
        Mod_Util.WriteError "P002", "BuildPrompt03Profile: не удалось определить профиль для шага '" & lastStep & "'"
        MsgBox "Ошибка P002: не удалось определить профиль обработки для текущего шага '" & lastStep & "'.", vbCritical, "Промт №0.3"
        Exit Sub
    End If

    template = ReadTextFile(templatePath)
    If template = "" Then
        Mod_Util.WriteError "P002", "BuildPrompt03Profile: не найден шаблон prompt_step_00_3_template.md"
        MsgBox "Ошибка P002: не найден шаблон prompt_step_00_3_template.md в папке 04_Промты\Шаблоны.", vbCritical, "Промт №0.3"
        Exit Sub
    End If

    result = template

    ' 1. {LAST_STEP_NUMBER}
    result = ReplaceIfPresent(result, "{LAST_STEP_NUMBER}", lastStep)

    ' 2. {INPUT_FILENAME} / {SCHEMA_FILENAME} (Схема: самая новая из 02_Подготовка данных\01_Схема)
    schemaFolder = root & "\02_Подготовка данных\01_Схема"
    fName = Dir(schemaFolder & "\Адаптированная_структура_*.xlsx")
    Do While fName <> ""
        If Left(fName, 2) <> "~$" Then
            If bestFile = "" Or FileDateTime(schemaFolder & "\" & fName) > bestDate Then
                bestFile = fName
                bestDate = FileDateTime(schemaFolder & "\" & fName)
            End If
        End If
        fName = Dir()
    Loop
    schemaFile = bestFile
    
    If schemaFile = "" Then
        Mod_Util.WriteError "P002", "BuildPrompt03Profile: не найдена адаптированная схема в " & schemaFolder
        MsgBox "Ошибка P002: не найдена адаптированная схема (*.xlsx) в папке:" & vbCrLf & schemaFolder, vbCritical, "Промт №0.3"
        Exit Sub
    End If
    result = ReplaceIfPresent(result, "{INPUT_FILENAME}", schemaFile)
    result = ReplaceIfPresent(result, "{SCHEMA_FILENAME}", schemaFile)

    ' 3. {DESCRIPTION_FILENAME} (Описание: Выжимка_Описание.md или fallback *.docx)
    descFolder = root & "\02_Подготовка данных\02_Описание"
    descFile = FirstMatchFile(descFolder, "Выжимка_Описание.md")
    If descFile = "" Then descFile = FirstMatchFile(descFolder, "*.docx")
    
    If descFile = "" Then
        Mod_Util.WriteError "P002", "BuildPrompt03Profile: не найдено описание в " & descFolder
        MsgBox "Ошибка P002: не найдено описание (Выжимка_Описание.md или *.docx) в папке:" & vbCrLf & descFolder, vbCritical, "Промт №0.3"
        Exit Sub
    End If
    result = ReplaceIfPresent(result, "{DESCRIPTION_FILENAME}", descFile)

    ' 4. {EXTRA_DATA_LIST}
    extraFolder = root & "\02_Подготовка данных\04_Дополнительные данные"
    extraList = ListFolderFiles(extraFolder)
    If extraList = "" Then extraList = "НЕ НАЙДЕНЫ"
    result = ReplaceIfPresent(result, "{EXTRA_DATA_LIST}", extraList)

    ' 5. {ЛИМИТ_ТИПОВ_В_ТАБЛИЦЕ}
    limitTypes = Mod_Util.GetSettingLong("MaxTypesPerTable", 50)
    result = ReplaceIfPresent(result, "{ЛИМИТ_ТИПОВ_В_ТАБЛИЦЕ}", CStr(limitTypes))

    ' 6. {{ИСКЛЮЧЕНИЯ_ПАТТЕРНОВ}} и {{СПИСОК_ИЗВЕСТНЫХ_ПАТТЕРНОВ}}
    result = ReplaceIfPresent(result, "{{ИСКЛЮЧЕНИЯ_ПАТТЕРНОВ}}", Mod_Patterns.GetPatternExceptions())
    result = ReplaceIfPresent(result, "{{СПИСОК_ИЗВЕСТНЫХ_ПАТТЕРНОВ}}", Mod_Patterns.GetKnownPatterns())

    ' Сохранение (W7.2-NOTE-5: строго в 04_Промты, НЕ в папку шага)
    EnsureFolder promptsPath
    SaveTextFileUtf8 promptsPath & "\Промт_шаг_00_3_" & profile & ".md", result

    Mod_Util.WriteLog "BuildPrompt03Profile: сформирован Промт_шаг_00_3_" & profile & ".md"
    MsgBox "Сформирован промт:" & vbCrLf & promptsPath & "\Промт_шаг_00_3_" & profile & ".md", vbInformation, "Промт №0.3"
    Exit Sub

ErrHandler:
    Mod_Util.WriteError "P002", "BuildPrompt03Profile: ошибка. " & Err.Description
    MsgBox "Ошибка P002: не удалось сформировать промт №0.3." & vbCrLf & Err.Description, vbCritical, "Промт №0.3"
    Err.Clear
End Sub

' ============================================================================
' W7.2 / W-ANALYTIC-06C: Команда разбора ответа ИИ (форматы 00_2 и 00_3)
' ============================================================================
Public Sub CommandParseAiResponse()
    Dim root As String, fragmentsPath As String, archivePath As String
    Dim fileArr As Variant, i As Long, f As String, fPath As String, fText As String
    Dim coll002 As New Collection, coll003 As New Collection
    Dim arr002() As String, arr003() As String
    Dim n002 As Long, n003 As Long, j As Long, temp As String
    Dim processFiles() As String, totalFiles As Long
    Dim successCount As Long, errorCount As Long, warnCount As Long
    Dim prevScreenUpdating As Boolean
    Dim maxStep As String ' F4: объявление переменной для отслеживания макс. шага
    
    On Error GoTo ErrHandler
    
    root = NormPath(Mod_Util.GetProjectRoot())
    If root = "" Then
        MsgBox "Корневая папка проекта не задана.", vbExclamation, "Ошибка"
        Exit Sub
    End If
    
    fragmentsPath = root & "\02_Подготовка данных\05_Фрагменты шагов"
    archivePath = root & "\02_Подготовка данных\Архив"
    
    If Not FolderExistsSafe(fragmentsPath) Then
        MsgBox "Папка фрагментов не найдена: " & fragmentsPath, vbExclamation, "Ошибка"
        Exit Sub
    End If
    
    ' F4: Сбор и классификация всех файлов ответов
    ' W-ANALYTIC-08-PARSE-DIAG-A: фильтр по префиксу Otvet_ или Ansver_ (вместо расширения .md)
    fileArr = GetFileListArray(fragmentsPath)
    Dim totalFilesInFolder As Long: totalFilesInFolder = 0
    Dim matchingFilesCount As Long: matchingFilesCount = 0
    
    If IsArray(fileArr) Then
        totalFilesInFolder = UBound(fileArr) - LBound(fileArr) + 1
        For i = LBound(fileArr) To UBound(fileArr)
            f = CStr(fileArr(i))
            
            ' Проверка префикса: Otvet_ (6 символов) или Ansver_ (7 символов)
            Dim isTargetFile As Boolean
            isTargetFile = (StrComp(Left$(f, 6), "Otvet_", vbTextCompare) = 0) Or _
                           (StrComp(Left$(f, 7), "Ansver_", vbTextCompare) = 0)
            
            If isTargetFile Then
                ' Пропуск временных файлов Office и накопительного файла
                If Left$(f, 2) <> "~$" And StrComp(f, "Требуют_решения.md", vbTextCompare) <> 0 Then
                    matchingFilesCount = matchingFilesCount + 1
                    fPath = fragmentsPath & "\" & f
                    fText = ReadTextFile(fPath)
                    If InStr(fText, MK_STP_START) > 0 Then
                        coll003.Add f
                    Else
                        coll002.Add f
                    End If
                End If
            End If
        Next i
    End If
    
    ' W-ANALYTIC-08-PARSE-DIAG-A: явная диагностика при отсутствии подходящих файлов
    If matchingFilesCount = 0 Then
        Dim diagMsg As String
        diagMsg = "В папке '" & fragmentsPath & "' не найдено ни одного файла ответа ИИ" & vbCrLf & _
                  "с префиксом 'Otvet_' или 'Ansver_' (найдено файлов всего: " & totalFilesInFolder & ")." & vbCrLf & vbCrLf & _
                  "Убедитесь, что:" & vbCrLf & _
                  "1. Файл ответа ИИ сохранён в папке '05_Фрагменты шагов'" & vbCrLf & _
                  "2. Имя файла начинается с 'Otvet_' или 'Ansver_' (например: Otvet_шаг03.md, Ansver_шаг04.txt)" & vbCrLf & _
                  "3. Расширение файла может быть любым (.md, .txt, .md.txt и т.д.)"
        Mod_Util.WriteWarning "P010", "CommandParseAiResponse: " & diagMsg
        MsgBox diagMsg, vbExclamation, "Разбор ответа ИИ — файлы не найдены"
        Application.ScreenUpdating = prevScreenUpdating
        Exit Sub
    End If
    
    ' Сортировка по имени (пузырьком)
    n002 = coll002.Count
    If n002 > 0 Then
        ReDim arr002(1 To n002)
        For i = 1 To n002: arr002(i) = coll002(i): Next i
        For i = 1 To n002 - 1
            For j = i + 1 To n002
                If StrComp(arr002(i), arr002(j), vbTextCompare) > 0 Then
                    temp = arr002(i): arr002(i) = arr002(j): arr002(j) = temp
                End If
            Next j
        Next i
    End If
    
    n003 = coll003.Count
    If n003 > 0 Then
        ReDim arr003(1 To n003)
        For i = 1 To n003: arr003(i) = coll003(i): Next i
        For i = 1 To n003 - 1
            For j = i + 1 To n003
                If StrComp(arr003(i), arr003(j), vbTextCompare) > 0 Then
                    temp = arr003(i): arr003(i) = arr003(j): arr003(j) = temp
                End If
            Next j
        Next i
    End If
    
    totalFiles = n002 + n003
    If totalFiles = 0 Then
        MsgBox "Файлы ответов ИИ не найдены в папке: " & fragmentsPath, vbInformation, "Разбор ответов ИИ"
        Exit Sub
    End If
    
    ReDim processFiles(1 To totalFiles)
    For i = 1 To n002: processFiles(i) = arr002(i): Next i
    For i = 1 To n003: processFiles(n002 + i) = arr003(i): Next i
    
    prevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    successCount = 0: errorCount = 0: warnCount = 0
    maxStep = ""
    
    ' === ЦИКЛ ОБРАБОТКИ ВСЕХ ФАЙЛОВ (F4) ===
    For i = 1 To totalFiles
        Dim aiResponseFile As String
        aiResponseFile = processFiles(i)
        Dim aiResponseText As String
        aiResponseText = ReadTextFile(fragmentsPath & "\" & aiResponseFile)
        
        If aiResponseText = "" Then
            Mod_Util.WriteWarning "P006", "CommandParseAiResponse: файл пуст или не прочитан: " & aiResponseFile
            errorCount = errorCount + 1
            GoTo NextFile
        End If
        
        Dim isFormat002 As Boolean, isFormat003 As Boolean
        isFormat002 = (InStr(aiResponseText, MK_PAT_START) > 0)
        isFormat003 = (InStr(aiResponseText, MK_STP_START) > 0)
        
        If Not isFormat002 And Not isFormat003 Then
            Mod_Util.WriteWarning "P008", "CommandParseAiResponse: нераспознанный формат файла " & aiResponseFile
            errorCount = errorCount + 1
            GoTo NextFile
        End If

        ' --- ЛОКАЛЬНАЯ ОБРАБОТКА ОДНОГО ФАЙЛА С ИЗОЛЯЦИЕЙ ОШИБОК ---
        On Error GoTo FileError
        
        Dim patternsCsv As String, stepsCsv As String
        Dim referenceBooks As String, userAction As String, statistics As String
        ExtractBlock aiResponseText, "<!-- PATTERNS_CSV_START -->", "<!-- PATTERNS_CSV_END -->", patternsCsv
        ExtractBlock aiResponseText, "<!-- STEPS_CSV_START -->", "<!-- STEPS_CSV_END -->", stepsCsv
        ExtractBlock aiResponseText, "<!-- REFERENCEBOOKS_START -->", "<!-- REFERENCEBOOKS_END -->", referenceBooks
        ExtractBlock aiResponseText, "<!-- USER_ACTION_START -->", "<!-- USER_ACTION_END -->", userAction
        ExtractBlock aiResponseText, "<!-- STATISTICS_START -->", "<!-- STATISTICS_END -->", statistics
        
        Dim schemaPath As String, schemaFile As String
        schemaPath = root & "\02_Подготовка данных\01_Схема"
        schemaFile = GetNewestSchemaArtifactName(schemaPath)
        If schemaFile = "" Then
            Mod_Util.WriteError "P010", "CommandParseAiResponse: не найдена Адаптированная_структура_*.xlsx"
            errorCount = errorCount + 1
            GoTo NextFile
        End If
        
        Dim wb As Workbook, ws As Worksheet
        Set wb = Workbooks.Open(schemaPath & "\" & schemaFile, ReadOnly:=False)
        Set ws = wb.Worksheets(Mod_Structure.FLAT_SHEET_NAME)
        
        If ws Is Nothing Then
            Mod_Util.WriteError "P010", "CommandParseAiResponse: лист '" & Mod_Structure.FLAT_SHEET_NAME & "' не найден"
            wb.Close SaveChanges:=False
            Set wb = Nothing
            errorCount = errorCount + 1
            GoTo NextFile
        End If
        
        If ws.ProtectContents Then
            On Error Resume Next
            ws.Unprotect Password:="pilot"
            On Error GoTo FileError
        End If
        
        Dim colName As Long, colCat As Long, colPat As Long, colAnom As Long, colStep As Long
        colName = FindColumn(ws, "Отображаемое имя")
        colCat = FindColumn(ws, "Категория")
        colPat = FindColumn(ws, "Паттерн")
        colAnom = FindColumn(ws, "Аномалия")
        colStep = FindColumn(ws, "№шага/подшага")
        If colStep = 0 Then colStep = FindColumn(ws, "Шаг")
        
        If colName = 0 Or colAnom = 0 Then
            Mod_Util.WriteError "P010", "CommandParseAiResponse: не найдены обязательные колонки"
            wb.Close SaveChanges:=False
            Set wb = Nothing
            errorCount = errorCount + 1
            GoTo NextFile
        End If
        
        Dim lastRow As Long, decisionContent As String
        lastRow = ws.Cells(ws.Rows.Count, colName).End(xlUp).row
        decisionContent = ""

                ' === W7.3-PATTERNS-UPDATE: запись паттернов и исключений в ListObject ===
        ' Все переменные с суффиксом W (Write) — гарантия отсутствия конфликтов
        ' с уже объявленными iPat, jPat, rPat, linesPat, partsPat и т.д.
        Dim wsPatternsW As Worksheet
        Dim tblListW As ListObject
        Dim tblExcW As ListObject
        Dim vCsvLinesW As Variant
        Dim vLinePartsW As Variant
        Dim iPatW As Long
        Dim iExcW As Long
        Dim newRowW As ListRow
        Dim patternsAddedW As Long
        Dim excAddedW As Long

        On Error Resume Next
        Set wsPatternsW = ThisWorkbook.Worksheets("Паттерны")
        If Not wsPatternsW Is Nothing Then
            Set tblListW = wsPatternsW.ListObjects(Mod_Patterns.TBL_PATTERNS_LIST)
            Set tblExcW = wsPatternsW.ListObjects(Mod_Patterns.TBL_PATTERNS_EXC)
        End If
        On Error GoTo FileError

        If patternsCsv <> "" Then
            ' Разбиваем CSV на строки (разделитель vbCrLf)
            vCsvLinesW = Split(patternsCsv, vbCrLf)

            ' Список паттернов: строки из 3-4 полей (Паттерн|Типы|Количество[|Итерация])
            If Not tblListW Is Nothing Then
                For iPatW = LBound(vCsvLinesW) To UBound(vCsvLinesW)
                    If Trim(CStr(vCsvLinesW(iPatW))) <> "" Then
                        vLinePartsW = Split(CStr(vCsvLinesW(iPatW)), "|")
                        If IsArray(vLinePartsW) And UBound(vLinePartsW) >= 2 And UBound(vLinePartsW) <= 3 Then
                            Set newRowW = tblListW.ListRows.Add
                            newRowW.Range.Cells(1, 1).Value = CStr(vLinePartsW(0))  ' Паттерн
                            newRowW.Range.Cells(1, 2).Value = CStr(vLinePartsW(1))  ' Типы
                            newRowW.Range.Cells(1, 3).Value = CStr(vLinePartsW(2))  ' Количество
                            If UBound(vLinePartsW) = 3 Then newRowW.Range.Cells(1, 4).Value = CStr(vLinePartsW(3))  ' Итерация
                            patternsAddedW = patternsAddedW + 1
                        End If
                    End If
                Next iPatW
            End If

            ' Исключения: строки из 8+ полей (Область|Поле|Тип матчинга|Значение|Действие|Примеры|Причина|Статус)
            If Not tblExcW Is Nothing Then
                For iExcW = LBound(vCsvLinesW) To UBound(vCsvLinesW)
                    If Trim(CStr(vCsvLinesW(iExcW))) <> "" Then
                        vLinePartsW = Split(CStr(vCsvLinesW(iExcW)), "|")
                        If IsArray(vLinePartsW) And UBound(vLinePartsW) >= 7 Then
                            Set newRowW = tblExcW.ListRows.Add
                            newRowW.Range.Cells(1, 1).Value = CStr(vLinePartsW(0))  ' Область
                            newRowW.Range.Cells(1, 2).Value = CStr(vLinePartsW(1))  ' Поле
                            newRowW.Range.Cells(1, 3).Value = CStr(vLinePartsW(2))  ' Тип матчинга
                            newRowW.Range.Cells(1, 4).Value = CStr(vLinePartsW(3))  ' Значение
                            newRowW.Range.Cells(1, 5).Value = CStr(vLinePartsW(4))  ' Действие
                            newRowW.Range.Cells(1, 6).Value = CStr(vLinePartsW(5))  ' Примеры
                            newRowW.Range.Cells(1, 7).Value = CStr(vLinePartsW(6))  ' Причина
                            newRowW.Range.Cells(1, 8).Value = CStr(vLinePartsW(7))  ' Статус
                            excAddedW = excAddedW + 1
                        End If
                    End If
                Next iExcW
            End If

            Mod_Util.WriteLog "W7.3-PATTERNS-UPDATE: паттернов добавлено: " & CStr(patternsAddedW) & _
                ", исключений добавлено: " & CStr(excAddedW)
        End If
        ' === END W7.3-PATTERNS-UPDATE ===
        
        ' =========================================================================
        ' ЛОГИКА 00_2 (PATTERNS_CSV)
        ' =========================================================================
        If isFormat002 Then
            Dim linesPat As Variant, iPat As Long, jPat As Long, rPat As Long
            Dim linePat As String, partsPat As Variant
            Dim patternName As String, typesStrPat As String, typesArrPat As Variant, typeItemPat As String
            linesPat = Split(patternsCsv, vbCrLf)
            For iPat = LBound(linesPat) To UBound(linesPat)
                linePat = Trim(linesPat(iPat))
                If linePat = "" Or InStr(linePat, "Паттерн|Типы|Количество") > 0 Then GoTo NextLinePat
                partsPat = Split(linePat, "|")
                If UBound(partsPat) < 2 Then GoTo NextLinePat
                patternName = Trim(partsPat(0))
                typesStrPat = Trim(partsPat(1))
                typesArrPat = Split(typesStrPat, ";")
                
                Dim categories As Object
                Set categories = CreateObject("Scripting.Dictionary")
                For jPat = LBound(typesArrPat) To UBound(typesArrPat)
                    typeItemPat = Trim(typesArrPat(jPat))
                    If typeItemPat = "" Then GoTo NextTypePat
                    For rPat = 2 To lastRow
                        On Error Resume Next
                        If Trim(CStr(ws.Cells(rPat, colName).value)) = typeItemPat Then
                            Dim cat As String: cat = Trim(CStr(ws.Cells(rPat, colCat).value))
                            If cat <> "" Then categories(cat) = 1
                        End If
                        On Error GoTo FileError
                    Next rPat
NextTypePat:
                Next jPat
                
                If categories.Count > 1 Then
                    Mod_Util.WriteWarning "P010", "Гомогенность нарушена для паттерна '" & patternName & "': " & Join(categories.keys, ", ")
                    decisionContent = decisionContent & patternName & " | " & typesStrPat & " | Смешение категорий" & vbCrLf
                    warnCount = warnCount + 1
                    GoTo NextLinePat
                End If
                
                For jPat = LBound(typesArrPat) To UBound(typesArrPat)
                    typeItemPat = Trim(typesArrPat(jPat))
                    If typeItemPat = "" Then GoTo NextTypeUpdatePat
                    For rPat = 2 To lastRow
                        On Error Resume Next
                        If Trim(CStr(ws.Cells(rPat, colName).value)) = typeItemPat Then
                            Dim oldPat As String: oldPat = Trim(CStr(ws.Cells(rPat, colPat).value))
                            If oldPat <> patternName And patternName <> "" Then
                                ws.Cells(rPat, colPat).value = patternName
                            End If
                        End If
                        On Error GoTo FileError
                    Next rPat
NextTypeUpdatePat:
                Next jPat
NextLinePat:
            Next iPat
        End If
        
        ' =========================================================================
        ' ЛОГИКА 00_3 (STEPS_CSV)
        ' =========================================================================
        If isFormat003 Then
            If colStep > 0 Then
                Dim linesStep As Variant, iStep As Long, jStep As Long, rStep As Long
                Dim lineStep As String, partsStep As Variant
                Dim nn As String, typesStrStep As String, kStr As String, patternStep As String, warning As String
                Dim typesArrStep As Variant, typeItemStep As String
                Dim parsedNNs As Object
                Set parsedNNs = CreateObject("Scripting.Dictionary")
                
                linesStep = Split(stepsCsv, vbCrLf)
                For iStep = LBound(linesStep) To UBound(linesStep)
                    lineStep = Trim(linesStep(iStep))
                    If lineStep = "" Then GoTo NextLineStep
                    partsStep = Split(lineStep, "|")
                    If UBound(partsStep) < 3 Then GoTo NextLineStep
                    
                    nn = Trim(partsStep(0))
                    typesStrStep = Trim(partsStep(1))
                    kStr = Trim(partsStep(2))
                    patternStep = Trim(partsStep(3))
                    warning = ""
                    If UBound(partsStep) >= 4 Then warning = Trim(partsStep(4))
                    
                    If nn = "" Or nn = "нет" Or GetPromptProfile(nn) = "" Then
                        Mod_Util.WriteWarning "P010", "Неверный формат NN в STEPS_CSV: '" & nn & "'"
                        warnCount = warnCount + 1
                        GoTo NextLineStep
                    End If
                    
                    parsedNNs(nn) = 1
                    If IsStepGreater(nn, maxStep) Then maxStep = nn
                    
                    typesArrStep = Split(typesStrStep, ";")
                    For jStep = LBound(typesArrStep) To UBound(typesArrStep)
                        typeItemStep = Trim(typesArrStep(jStep))
                        If typeItemStep = "" Then GoTo NextTypeStep
                        For rStep = 2 To lastRow
                            On Error Resume Next
                            If Trim(CStr(ws.Cells(rStep, colName).value)) = typeItemStep Then
                                Dim oldStep As String: oldStep = Trim(CStr(ws.Cells(rStep, colStep).value))
                                If oldStep <> nn And nn <> "" Then
                                    ws.Cells(rStep, colStep).NumberFormat = "@"
                                    ws.Cells(rStep, colStep).value = nn
                                End If
                                If warning <> "" Then
                                    Dim curAnomStep As String: curAnomStep = Trim(CStr(ws.Cells(rStep, colAnom).value))
                                    If InStr(curAnomStep, warning) = 0 Then
                                        If curAnomStep = "" Then
                                            ws.Cells(rStep, colAnom).value = warning
                                        Else
                                            ws.Cells(rStep, colAnom).value = curAnomStep & "; " & warning
                                        End If
                                    End If
                                End If
                            End If
                            On Error GoTo FileError
                        Next rStep
NextTypeStep:
                    Next jStep
NextLineStep:
                Next iStep
                
                ' F5: Несоответствие шагов -> WriteWarning P010 (не WriteError)
                Dim statLines As Variant, genStepsStr As String, statIdx As Long
                statLines = Split(statistics, vbCrLf)
                genStepsStr = ""
                For statIdx = LBound(statLines) To UBound(statLines)
                    lineStep = Trim(statLines(statIdx))
                    If InStr(1, lineStep, "Сгенерированные шаги:", vbTextCompare) > 0 Then
                        genStepsStr = Trim(Mid(lineStep, InStr(lineStep, ":") + 1))
                        Exit For
                    End If
                Next statIdx
                
                If genStepsStr <> "" And InStr(1, genStepsStr, "СТОП", vbTextCompare) = 0 And InStr(1, genStepsStr, "не применимо", vbTextCompare) = 0 Then
                    Dim statNNs As Object, statParts As Variant, sNN As String, key As Variant, mismatch As Boolean
                    Set statNNs = CreateObject("Scripting.Dictionary")
                    genStepsStr = Replace(genStepsStr, ";", ",")
                    statParts = Split(genStepsStr, ",")
                    For statIdx = LBound(statParts) To UBound(statParts)
                        sNN = Trim(statParts(statIdx))
                        If sNN <> "" And sNN <> "-" Then statNNs(sNN) = 1
                    Next statIdx
                    
                    mismatch = False
                    If parsedNNs.Count <> statNNs.Count Then
                        mismatch = True
                    Else
                        For Each key In parsedNNs.keys
                            If Not statNNs.Exists(CStr(key)) Then
                                mismatch = True
                                Exit For
                            End If
                        Next key
                    End If
                    
                    If mismatch Then
                        Mod_Util.WriteWarning "P010", "Несоответствие сгенерированных шагов. Разобранные: " & Join(parsedNNs.keys, ", ") & ". В статистике: " & genStepsStr
                        warnCount = warnCount + 1
                    End If
                End If
            End If
        End If
        
        ' =========================================================================
        ' ОБЩАЯ ЛОГИКА: USER_ACTION (F1)
        ' =========================================================================
        Dim userActionTrimmed As String, uaClean As String, uaLines As Variant, uaIdx As Long, uaLine As String
        userActionTrimmed = Trim(userAction)
        If userActionTrimmed <> "" And InStr(userActionTrimmed, UA_NO_HEADER) = 0 And Left(LCase(userActionTrimmed), 3) <> "нет" Then
            uaClean = ""
            uaLines = Split(userActionTrimmed, vbCrLf)
            For uaIdx = LBound(uaLines) To UBound(uaLines)
                uaLine = Trim(uaLines(uaIdx))
                If uaLine <> "" And Left(uaLine, Len(UA_HEADER_PREFIX)) <> UA_HEADER_PREFIX Then
                    If uaClean <> "" Then uaClean = uaClean & vbCrLf
                    uaClean = uaClean & uaLine
                End If
            Next uaIdx
            
            If uaClean <> "" Then
                decisionContent = decisionContent & uaClean & vbCrLf
                Dim uaType As String, reason As String, curAnomUA As String, anomMarkUA As String, partsStepUA As Variant
                For uaIdx = LBound(uaLines) To UBound(uaLines)
                    uaLine = Trim(uaLines(uaIdx))
                    If uaLine = "" Or Left(uaLine, Len(UA_HEADER_PREFIX)) = UA_HEADER_PREFIX Then GoTo NextUaLine2
                    partsStepUA = Split(uaLine, "|")
                    
                    ' F1: Ветка UBound=1 для формата 00_3
                    If UBound(partsStepUA) = 1 Then
                        uaType = Trim(partsStepUA(1))
                        For rStep = 2 To lastRow
                            On Error Resume Next
                            If Trim(CStr(ws.Cells(rStep, colName).value)) = uaType Then
                                curAnomUA = Trim(CStr(ws.Cells(rStep, colAnom).value))
                                anomMarkUA = "требует решения (см. Требуют_решения.md)"
                                If InStr(curAnomUA, anomMarkUA) = 0 Then
                                    If curAnomUA = "" Then ws.Cells(rStep, colAnom).value = anomMarkUA _
                                    Else: ws.Cells(rStep, colAnom).value = curAnomUA & "; " & anomMarkUA
                                End If
                            End If
                            On Error GoTo FileError
                        Next rStep
                        
                    ' F1: Ветка UBound>=2 для формата 00_2
                    ElseIf UBound(partsStepUA) >= 2 Then
                        uaType = Trim(partsStepUA(1))
                        reason = Trim(partsStepUA(2))
                        For rStep = 2 To lastRow
                            On Error Resume Next
                            If Trim(CStr(ws.Cells(rStep, colName).value)) = uaType Then
                                curAnomUA = Trim(CStr(ws.Cells(rStep, colAnom).value))
                                anomMarkUA = "Требует решения: " & reason
                                If InStr(curAnomUA, anomMarkUA) = 0 Then
                                    If curAnomUA = "" Then ws.Cells(rStep, colAnom).value = anomMarkUA _
                                    Else: ws.Cells(rStep, colAnom).value = curAnomUA & "; " & anomMarkUA
                                End If
                            End If
                            On Error GoTo FileError
                        Next rStep
                    End If
NextUaLine2:
                Next uaIdx
            End If
        End If
        
        If decisionContent <> "" Then AppendToDecisionFile fragmentsPath, decisionContent
        
        ' Обработка REFERENCEBOOKS
        Dim refTrimmed As String
        refTrimmed = Trim(referenceBooks)
        If refTrimmed <> "" Then
            If InStr(refTrimmed, "{{СПИСОК_БАЗОВЫХ_ТИПОВ}}") > 0 Then
                Mod_Util.WriteWarning "P008", "CommandParseAiResponse: обнаружен формат 00_1 в REFERENCEBOOKS."
            Else
                Mod_Util.WriteLog "REFERENCEBOOKS извлечены: " & Replace(refTrimmed, vbCrLf, " | ")
            End If
        End If
        
        ' F2: Обработка STATISTICS
        Dim statTrimmed As String, statLine As String, statKey As String, statVal As String
        statTrimmed = Trim(statistics)
        If statTrimmed <> "" Then
            statLines = Split(statTrimmed, vbCrLf)
            For statIdx = LBound(statLines) To UBound(statLines)
                statLine = Trim(statLines(statIdx))
                If statLine = "" Then GoTo NextStatLine
                
                If InStr(statLine, "~") > 0 Or InStr(LCase(statLine), "около") > 0 Or InStr(LCase(statLine), "примерно") > 0 Then
                    Mod_Util.WriteWarning "P010", "STATISTICS: неопределённость: " & statLine
                    warnCount = warnCount + 1
                End If
                
                statParts = Split(statLine, ":")
                If UBound(statParts) >= 1 Then
                    statKey = Trim(statParts(0))
                    statVal = Trim(statParts(1))
                    If statKey = "PATTERN_NOT_RETURNED" Or statKey = "MOVED_TYPE" Or statKey = "LOST_TYPE" Or statKey = "DUPLICATE_TYPE" Then
                        If LCase(statVal) <> "нет" And statVal <> "" Then
                            Mod_Util.WriteWarning "P010", "STATISTICS: " & statKey & " = " & statVal
                            warnCount = warnCount + 1
                        End If
                    End If
                End If
NextStatLine:
            Next statIdx
        End If
        
        ' F5: Ошибка сохранения книги -> P009 (не P010)
        If Not wb.Saved Then
            On Error Resume Next
            wb.Save
            If Err.Number <> 0 Then
                Mod_Util.WriteError "P009", "CommandParseAiResponse: не удалось сохранить " & schemaFile & ". Ошибка: " & Err.Description
                Err.Clear
            End If
            On Error GoTo FileError
        End If
        
        UpdateStepsState wb, Mod_Structure.FLAT_SHEET_NAME
        wb.Close SaveChanges:=False
        Set wb = Nothing
        
        ' ===== АРХИВАЦИЯ И ОБНОВЛЕНИЕ =====
        EnsureFolder archivePath
        ArchiveFile fragmentsPath & "\" & aiResponseFile, archivePath, aiResponseFile
        
        ' F3: Архивация промтов ТОЛЬКО распознанного типа
        Dim targetPromptPrefix As String, archArr As Variant, k As Long, pf As String
        If isFormat002 Then
            targetPromptPrefix = "Prompt_00_2_"
        ElseIf isFormat003 Then
            targetPromptPrefix = "Prompt_00_3_"
        Else
            targetPromptPrefix = ""
        End If
        
        If targetPromptPrefix <> "" Then
            archArr = GetFileListArray(fragmentsPath)
            If IsArray(archArr) Then
                For k = LBound(archArr) To UBound(archArr)
                    pf = CStr(archArr(k))
                    ' Архивируем промты с нужным префиксом независимо от расширения
                    If Left$(pf, Len(targetPromptPrefix)) = targetPromptPrefix Then
                        ArchiveFile fragmentsPath & "\" & pf, archivePath, pf
                    End If
                Next k
            End If
        End If
        
        If maxStep <> "" Then UpdateLastStepNumber maxStep
        successCount = successCount + 1
        On Error GoTo ErrHandler ' Возврат глобального обработчика перед следующей итерацией
        
NextFile:
        ' Очистка объектов текущего файла перед следующей итерацией (изоляция)
        On Error Resume Next
        If Not wb Is Nothing Then
            wb.Close SaveChanges:=False
            Set wb = Nothing
        End If
        On Error GoTo ErrHandler
    Next i
    
    Application.ScreenUpdating = prevScreenUpdating
    Mod_Util.WriteLog "CommandParseAiResponse: разбор завершён. Успешно: " & successCount & ", Ошибок: " & errorCount & ", Предупреждений P010: " & warnCount
    MsgBox "Разбор ответов ИИ завершён." & vbCrLf & _
           "Обработано файлов: " & totalFiles & vbCrLf & _
           "Успешно: " & successCount & vbCrLf & _
           "С ошибками: " & errorCount & vbCrLf & _
           "Предупреждений P010: " & warnCount, vbInformation, "Разбор ответов ИИ"
    Exit Sub

FileError:
    ' F4: Изоляция ошибок по файлам
    Mod_Util.WriteError "P010", "CommandParseAiResponse: ошибка в файле " & aiResponseFile & ". " & Err.Description
    Err.Clear
    errorCount = errorCount + 1
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
    End If
    On Error GoTo ErrHandler
    Resume NextFile

ErrHandler:
    Application.ScreenUpdating = prevScreenUpdating
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Mod_Util.WriteError "P010", "CommandParseAiResponse: " & Err.Description
    MsgBox "Ошибка P010: " & Err.Description, vbCritical, "Разбор ответа ИИ"
    Err.Clear
End Sub

' ---------------------------------------------------------------------------
' Вспомогательные функции для CommandParseAiResponse
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' W-ANALYTIC-08-PARSE-DIAG-A: Поиск самого свежего файла ответа ИИ по префиксу.
' Ищет файлы с префиксами Otvet_ или Ansver_ (регистронезависимо) независимо
' от расширения (.md, .txt, .md.txt и т.д.). Возвращает имя самого свежего
' файла (по FileDateTime) или "" если подходящих файлов нет.
' ---------------------------------------------------------------------------
Private Function GetLatestAiResponseFile(ByVal folderPath As String) As String
    Dim fileArr As Variant
    Dim i As Long
    Dim f As String
    Dim bestFile As String
    Dim bestDate As Date
    Dim curDate As Date
    Dim isTargetFile As Boolean

    GetLatestAiResponseFile = ""
    If Not FolderExistsSafe(folderPath) Then Exit Function

    fileArr = GetFileListArray(folderPath)
    If Not IsArray(fileArr) Then Exit Function

    For i = LBound(fileArr) To UBound(fileArr)
        f = CStr(fileArr(i))
        
        ' W-ANALYTIC-08-A: проверка префикса вместо расширения .md
        isTargetFile = (StrComp(Left$(f, 6), "Otvet_", vbTextCompare) = 0) Or _
                       (StrComp(Left$(f, 7), "Ansver_", vbTextCompare) = 0)
        
        If isTargetFile Then
            ' пропуск: ~$*, Требуют_решения.md (накопительный, не архивируется)
            If Left$(f, 2) <> "~$" And StrComp(f, "Требуют_решения.md", vbTextCompare) <> 0 Then
                On Error Resume Next
                curDate = FileDateTime(folderPath & "\" & f)
                If Err.Number <> 0 Then
                    Err.Clear
                    curDate = 0
                End If
                On Error GoTo 0
                
                If bestFile = "" Or curDate > bestDate Then
                    bestFile = f
                    bestDate = curDate
                End If
            End If
        End If
    Next i
    GetLatestAiResponseFile = bestFile
End Function


Private Sub ExtractBlock(ByVal text As String, ByVal startMarker As String, ByVal endMarker As String, ByRef outBlock As String)
    Dim splitStart As Variant
    Dim splitEnd As Variant
    
    outBlock = ""
    splitStart = Split(text, startMarker)
    If UBound(splitStart) >= 1 Then
        splitEnd = Split(splitStart(1), endMarker)
        If UBound(splitEnd) >= 0 Then
            outBlock = splitEnd(0)
        End If
    End If
End Sub

Private Function FindColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim rng As Range
    Set rng = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not rng Is Nothing Then
        FindColumn = rng.Column
    Else
        FindColumn = 0
    End If
End Function

Private Sub AppendToDecisionFile(ByVal folderPath As String, ByVal content As String)
    Dim filePath As String
    filePath = folderPath & "\Требуют_решения.md"
    Dim existing As String
    existing = ReadTextFile(filePath)                   ' БЕЗ Mod_Util.
    Dim newContent As String

    content = Trim(content)
    If content = "" Then Exit Sub

    If existing = "" Or InStr(existing, "# Требуют решения") = 0 Then
        newContent = "# Требуют решения" & vbCrLf & vbCrLf & content & vbCrLf
    Else
        Dim lines As Variant                            ' Split строго в Variant
        lines = Split(existing, vbCrLf)
        Dim i As Long
        Dim line As String
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")

        For i = LBound(lines) To UBound(lines)
            line = Trim(lines(i))
            If line <> "" And line <> "# Требуют решения" Then
                dict(line) = 1
            End If
        Next i

        Dim newLines As Variant                         ' Split строго в Variant
        newLines = Split(content, vbCrLf)
        For i = LBound(newLines) To UBound(newLines)
            line = Trim(newLines(i))
            If line <> "" Then
                dict(line) = 1
            End If
        Next i

        newContent = "# Требуют решения" & vbCrLf & vbCrLf
        Dim key As Variant
        For Each key In dict.keys
            newContent = newContent & CStr(key) & vbCrLf
        Next key
    End If

    SaveTextFileUtf8 filePath, newContent               ' БЕЗ Mod_Util.
End Sub

' W-ANALYTIC-06B: Dir-цикл заменён на GetFileListArray + For; фильтр по
' маске SCHEMA_ARTIFACT_MASK сохранён; видимость Private (используется
' только внутри Mod_Prompt).
Private Function GetNewestSchemaArtifactName(ByVal prepPath As String) As String
    Dim fileArr As Variant
    Dim i As Long
    Dim fName As String
    Dim best As String
    Dim bestStamp As String
    Dim curStamp As String
    Dim bestDate As Date
    Dim curDate As Date

    GetNewestSchemaArtifactName = ""
    If Not FolderExistsSafe(prepPath) Then Exit Function

    fileArr = GetFileListArray(prepPath)
    If Not IsArray(fileArr) Then Exit Function          ' пустая папка -> Array()

    For i = LBound(fileArr) To UBound(fileArr)
        fName = CStr(fileArr(i))
        If Left$(fName, 2) <> "~$" Then                 ' временные файлы Office
            ' ИСПРАВЛЕНИЕ: убран LCase$() - оператор Like регистрозависимый,
            ' маска SCHEMA_ARTIFACT_MASK уже содержит правильный регистр.
            If fName Like SCHEMA_ARTIFACT_MASK Then
                On Error Resume Next
                curDate = FileDateTime(prepPath & "\" & fName)
                If Err.Number <> 0 Then
                    Err.Clear
                    curDate = 0
                End If
                On Error GoTo 0
                curStamp = ExtractDateStampFromName(fName)
                If best = "" Then
                    best = fName
                    bestStamp = curStamp
                    bestDate = curDate
                Else
                    If curStamp > bestStamp Or _
                       (curStamp = bestStamp And curDate > bestDate) Then
                        best = fName
                        bestStamp = curStamp
                        bestDate = curDate
                    End If
                End If
            End If
        End If
    Next i
    
    ' Отладочное логирование: что нашли (или не нашли)
    If best = "" Then
        Mod_Util.WriteWarning "P010", "GetNewestSchemaArtifactName: в папке " & _
            prepPath & " не найдено файлов по маске '" & SCHEMA_ARTIFACT_MASK & "'"
    Else
        Mod_Util.WriteLog "GetNewestSchemaArtifactName: выбран файл " & best
    End If
    
    GetNewestSchemaArtifactName = best
End Function

' ============================================================================
' W-ANALYTIC-05: Итеративная генерация промтов 00_2 / 00_3 (кнопка 13).
' Выбор шаблона: "Разбивка на паттерны" (PatternsState) = "Нет" -> 00_2,
' "Да" -> 00_3. Номер итерации N - максимум из Prompt_00_X_iter*.md в
' папках "05_Фрагменты шагов" и "Архив" + 1 (GetNextIterationNumber,
' безопасный обход через GetFileListArray). Подстановка плейсхолдеров -
' по актуальным шаблонам W-ANALYTIC-01. {LAST_STEP_NUMBER} читается из
' параметра LastStepNumber (W-ANALYTIC-09).
' ЕДИНСТВЕННАЯ реализация имени CommandGenerateTables в проекте (заглушка
' S0 из Mod_UI удалена - W-ANALYTIC-05-FIX, иначе "Ambiguous name detected").
' ============================================================================
Public Sub CommandGenerateTables()
    Dim root As String
    Dim promptsPath As String
    Dim fragmentsPath As String
    Dim templatePath As String
    Dim template As String
    Dim result As String
    Dim patternsState As String
    Dim promptType As String
    Dim nextIter As Long
    Dim fileName As String
    Dim filePath As String
    Dim schemaFolder As String
    Dim descFolder As String
    Dim extraFolder As String
    Dim schemaFile As String
    Dim descFile As String
    Dim extraList As String
    Dim lastStep As String
    Dim profile As String
    Dim limitTypes As Long
    Dim archivePath As String
    Dim schemaWb As Workbook
    Dim wsSchema As Worksheet
    Dim csvContent As String
    Dim csvFileName As String
    Dim csvRows As Long

    On Error GoTo ErrHandler

    root = NormPath(Mod_Util.GetProjectRoot())
    If root = "" Then
        MsgBox "Корневая папка проекта не задана.", vbExclamation, "Ошибка"
        Exit Sub
    End If

    promptsPath = root & "\04_Промты\Шаблоны"
    fragmentsPath = root & "\02_Подготовка данных\05_Фрагменты шагов"
    archivePath = root & "\02_Подготовка данных\Архив"

    ' 1. Выбор шаблона по настройке "Разбивка на паттерны"
    patternsState = Mod_Util.GetSettingText("PatternsState", "Нет")

    If StrComp(patternsState, "Да", vbTextCompare) = 0 Then
        promptType = "3"
        templatePath = promptsPath & "\prompt_step_00_3_template.md"
    Else
        promptType = "2"
        templatePath = promptsPath & "\prompt_step_00_2_template.md"
    End If

    ' 2. Чтение шаблона (безопасное UTF-8/CP1251)
    template = ReadTextFile(templatePath)
    If template = "" Then
        Mod_Util.WriteError "P002", "CommandGenerateTables: не найден шаблон " & templatePath
        MsgBox "Ошибка P002: не найден шаблон промта 00_" & promptType & ":" & vbCrLf & templatePath, _
               vbCritical, "Сформировать таблицы"
        Exit Sub
    End If

    ' 3. Номер следующей итерации (безопасно: GetFileListArray, без вложенных Dir())
    nextIter = GetNextIterationNumber(fragmentsPath, archivePath, promptType)

    ' 4. Общие входы для обоих шаблонов
    schemaFolder = root & "\02_Подготовка данных\01_Схема"
    descFolder = root & "\02_Подготовка данных\02_Описание"
    extraFolder = root & "\02_Подготовка данных\04_Дополнительные данные"

    schemaFile = GetNewestSchemaArtifactName(schemaFolder)
    If schemaFile = "" Then schemaFile = FirstMatchFile(schemaFolder, "*.xlsx")

    ' --- Открытие книги схемы и проверка листа ---
    If schemaFile <> "" Then
        Set schemaWb = Workbooks.Open(schemaFolder & "\" & schemaFile, ReadOnly:=True, AddToMru:=False)
        On Error Resume Next
        Set wsSchema = schemaWb.Worksheets("Адаптированная_структура")
        On Error GoTo ErrHandler
        
        If wsSchema Is Nothing Then
            schemaWb.Close SaveChanges:=False
            Set schemaWb = Nothing
            Mod_Util.WriteError "P002", "CommandGenerateTables: не найден лист 'Адаптированная_структура' в " & schemaFile
            MsgBox "Ошибка P002: не найден лист 'Адаптированная_структура' в файле:" & vbCrLf & schemaFile, _
                   vbCritical, "Сформировать таблицы"
            Exit Sub
        End If
    End If

    ' Поиск файла описания: универсальная маска (любой формат)
    descFile = FirstMatchFile(descFolder, "Выжимка_Описание.*")
    If descFile = "" Then descFile = "нет"

    extraList = ListFolderFiles(extraFolder)
    If extraList = "" Then extraList = "нет"

    ' 5. Подстановка плейсхолдеров по типу промта
    If promptType = "2" Then
        ' Промт 00_2: Схема обязательна (шаблон: "Пусто - недопустимо")
        If schemaFile = "" Then
            Mod_Util.WriteError "P002", "CommandGenerateTables: не найдена адаптированная схема в " & schemaFolder
            MsgBox "Ошибка P002: не найдена адаптированная схема (*.xlsx) в папке:" & vbCrLf & schemaFolder, _
                   vbCritical, "Сформировать таблицы"
            Exit Sub
        End If
        
        ' Генерация CSV фрагмента для 00_2
        csvContent = BuildSchemaCsvFragment(wsSchema, "", 0)
        csvFileName = "Schema_00_2_iter" & CStr(nextIter) & ".csv"
        SaveTextFileUtf8 fragmentsPath & "\" & csvFileName, csvContent
        
        result = template
        result = ReplaceIfPresent(result, "{SCHEMA_FILENAME}", csvFileName)
        result = ReplaceIfPresent(result, "{DESCRIPTION_FILENAME}", descFile)
        result = ReplaceIfPresent(result, "{EXTRA_DATA_LIST}", extraList)
        result = ReplaceIfPresent(result, "{{СПИСОК_ИЗВЕСТНЫХ_ПАТТЕРНОВ}}", Mod_Patterns.GetKnownPatterns())
        result = ReplaceIfPresent(result, "{{ИСКЛЮЧЕНИЯ_ПАТТЕРНОВ}}", Mod_Patterns.GetPatternExceptions())
        
    Else
        ' Промт 00_3: {LAST_STEP_NUMBER} из настройки LastStepNumber (W-ANALYTIC-09)
        lastStep = Mod_Util.GetSettingText("LastStepNumber", "нет")
        If Trim$(lastStep) = "" Then lastStep = "нет"
        profile = GetPromptProfile(lastStep)
        If profile = "" Then
            Mod_Util.WriteError "P002", "CommandGenerateTables: невалидный номер последнего шага '" & lastStep & "'"
            MsgBox "Ошибка P002: невалидное значение ""Номер последнего шага"": '" & lastStep & "'." & vbCrLf & _
                   "Допустимые формы: нет, 0, 00, 00.i, 01, 01.i, 01.i.k, NN (>= 02)." & vbCrLf & _
                   "Обновите значение на листе ""Настройки"".", vbCritical, "Сформировать таблицы"
            Exit Sub
        End If
        If schemaFile = "" Then
            Mod_Util.WriteError "P002", "CommandGenerateTables: не найдена адаптированная схема (вход 00_3) в " & schemaFolder
            MsgBox "Ошибка P002: не найден входной файл схемы для промта 00_3 в папке:" & vbCrLf & schemaFolder, _
                   vbCritical, "Сформировать таблицы"
            Exit Sub
        End If
        
        ' Генерация CSV фрагмента для 00_3 с лимитом
        limitTypes = Mod_Util.GetSettingLong("MaxTypesPerTable", 50)

        ' Генерация CSV фрагмента для 00_3 с лимитом; уже распределённые
        ' по шагам типы в вход не попадают (W-ANALYTIC-05-FIX-2)
        csvContent = BuildSchemaCsvFragment(wsSchema, "", limitTypes, True, csvRows)
        If csvRows = 0 Then
            ' Все типы уже распределены - промт с пустой таблицей не создаём
            If Not schemaWb Is Nothing Then
                schemaWb.Close SaveChanges:=False
                Set schemaWb = Nothing
            End If
            Mod_Util.WriteLog "CommandGenerateTables: все типы уже распределены по шагам - промт 00_3 не сформирован."
            MsgBox "Все типы уже распределены по шагам (колонка «№шага/подшага» заполнена)." & vbCrLf & _
                   "Новый промт 00_3 не требуется: дождитесь разбора ответа ИИ (кнопка 14) " & _
                   "или переходите к промтам шагов («Сделать промт следующего шага»).", _
                   vbInformation, "Сформировать таблицы"
            Exit Sub
        End If
        csvFileName = "Schema_00_3_iter" & CStr(nextIter) & ".csv"
        SaveTextFileUtf8 fragmentsPath & "\" & csvFileName, csvContent
        
        result = template
        result = ReplaceIfPresent(result, "{LAST_STEP_NUMBER}", lastStep)
        result = ReplaceIfPresent(result, "{INPUT_FILENAME}", csvFileName)
        result = ReplaceIfPresent(result, "{DESCRIPTION_FILENAME}", descFile)
        result = ReplaceIfPresent(result, "{EXTRA_DATA_LIST}", extraList)
        result = ReplaceIfPresent(result, "{ЛИМИТ_ТИПОВ_В_ТАБЛИЦЕ}", CStr(limitTypes))
    End If

    ' 6. Сохранение (папка создаётся при необходимости)
    EnsureFolder fragmentsPath
    fileName = "Prompt_00_" & promptType & "_iter" & CStr(nextIter) & ".md"
    filePath = fragmentsPath & "\" & fileName
    SaveTextFileUtf8 filePath, result

    ' 7. Уведомление и лог
    Dim msg As String
    msg = "Сформирован промт 00_" & promptType & " (итерация " & CStr(nextIter) & "):" & vbCrLf & filePath
    Mod_Util.WriteLog "CommandGenerateTables: " & msg
    MsgBox msg, vbInformation, "Сформировать таблицы"

    ' Очистка памяти: штатное закрытие книги схемы
    If Not schemaWb Is Nothing Then
        schemaWb.Close SaveChanges:=False
        Set schemaWb = Nothing
    End If

    Exit Sub

ErrHandler:
    ' Очистка памяти при ошибке (гарантированное закрытие)
    On Error Resume Next
    If Not schemaWb Is Nothing Then
        schemaWb.Close SaveChanges:=False
        Set schemaWb = Nothing
    End If
    On Error GoTo 0
    
    Mod_Util.WriteError "P002", "CommandGenerateTables: " & Err.Description
    MsgBox "Ошибка P002: не удалось сформировать промт." & vbCrLf & Err.Description, _
           vbCritical, "Сформировать таблицы"
    Err.Clear
End Sub

' ---------------------------------------------------------------------------
' Вспомогательная функция: определяет следующий номер итерации для промта 00_2 или 00_3.
' Сканирует папки "05_Фрагменты шагов" и "Архив" с использованием безопасного
' обхода GetFileListArray (защита от сброса stateful-итератора Dir()).
' ---------------------------------------------------------------------------
Private Function GetNextIterationNumber(ByVal folderFragments As String, ByVal folderArchive As String, ByVal promptType As String) As Long
    Dim maxN As Long
    Dim fileArr As Variant
    Dim i As Long
    Dim fName As String
    Dim prefix As String
    Dim numStr As String
    
    maxN = 0
    prefix = "Prompt_00_" & promptType & "_iter"
    
    ' Сканирование папки фрагментов
    If FolderExistsSafe(folderFragments) Then
        fileArr = GetFileListArray(folderFragments)
        For i = LBound(fileArr) To UBound(fileArr)
            fName = CStr(fileArr(i))
            If Left$(fName, Len(prefix)) = prefix And Right$(fName, 3) = ".md" Then
                numStr = Mid$(fName, Len(prefix) + 1, Len(fName) - Len(prefix) - 3)
                If IsNumeric(numStr) Then
                    If CLng(numStr) > maxN Then maxN = CLng(numStr)
                End If
            End If
        Next i
    End If
    
    ' Сканирование папки архива
    If FolderExistsSafe(folderArchive) Then
        fileArr = GetFileListArray(folderArchive)
        For i = LBound(fileArr) To UBound(fileArr)
            fName = CStr(fileArr(i))
            If Left$(fName, Len(prefix)) = prefix And Right$(fName, 3) = ".md" Then
                numStr = Mid$(fName, Len(prefix) + 1, Len(fName) - Len(prefix) - 3)
                If IsNumeric(numStr) Then
                    If CLng(numStr) > maxN Then maxN = CLng(numStr)
                End If
            End If
        Next i
    End If
    
    ' Если файлов не найдено, maxN = 0, функция вернёт 1
    GetNextIterationNumber = maxN + 1
End Function

' ---------------------------------------------------------------------------
' W-ANALYTIC-04 / W-ANALYTIC-04-FIX: Генерация аналитических промтов 00_1 и 00_4
' ИСПРАВЛЕНИЕ: список файлов каждой папки собирается в массив ДО начала
' обработки (GetFileListArray), обход - через For. Это защищает от сброса
' stateful-итератора Dir() вложенными вызовами (SaveTextFileUtf8 ->
' StripUtf8Bom -> Dir(...)) и проверки существования файла промта
' (FileExistsSafe вместо Dir(promptPath)). За один запуск теперь создаются
' промты для ВСЕХ файлов папок 02_Описание, 03_Конфигурация,
' 04_Дополнительные данные.
' ---------------------------------------------------------------------------
Public Sub GenerateAnalyticalPrompts_Click()
    Dim root As String
    Dim configPath As String, descPath As String, extraPath As String
    Dim template001Path As String, template004Path As String
    Dim fName As String, promptName As String, promptPath As String
    Dim created001 As Long, skipped001 As Long
    Dim created004 As Long, skipped004 As Long
    Dim ws As Worksheet, r As Long, found As Boolean
    Dim tpl001 As String, tpl004 As String
    Dim paths As Variant, p As Variant
    Dim fileArr As Variant
    Dim i As Long

    On Error GoTo ErrHandler

    root = NormPath(Mod_Util.GetProjectRoot())
    If root = "" Then
        MsgBox "Корневая папка проекта не задана.", vbExclamation, "Ошибка"
        Exit Sub
    End If

    configPath = root & "\01_Исходные данные\03_Конфигурация"
    descPath = root & "\01_Исходные данные\02_Описание"
    extraPath = root & "\01_Исходные данные\04_Дополнительные данные"

    template001Path = root & "\04_Промты\Шаблоны\prompt_step_00_1_template.md"
    template004Path = root & "\04_Промты\Шаблоны\prompt_step_00_4_template.md"

    ' Чтение шаблонов (безопасное чтение UTF-8/CP1251)
    tpl001 = ReadTextFile(template001Path)
    tpl004 = ReadTextFile(template004Path)

    If tpl001 = "" Then Mod_Util.WriteWarning "P001", "Шаблон prompt_step_00_1_template.md не найден или пуст."
    If tpl004 = "" Then Mod_Util.WriteWarning "P001", "Шаблон prompt_step_00_4_template.md не найден или пуст."

    ' 1. Генерация 00_1 (Конфигурация) - обход по массиву файлов
    If FolderExistsSafe(configPath) Then
        fileArr = GetFileListArray(configPath)
        For i = LBound(fileArr) To UBound(fileArr)
            fName = CStr(fileArr(i))
            ' Пропуск: временные файлы Office (~$), уже созданные промты
            ' (Prompt_* - защита от рекурсии), файлы без расширения
            If Left$(fName, 2) <> "~$" And LCase$(Left$(fName, 7)) <> "prompt_" And InStr(fName, ".") > 0 Then
                promptName = "Prompt_" & fName & ".md"
                promptPath = configPath & "\" & promptName
                ' FileExistsSafe (GetAttr) - НЕ сбрасывает итератор Dir()
                If Not FileExistsSafe(promptPath) Then
                    If tpl001 <> "" Then
                        SaveTextFileUtf8 promptPath, Replace(tpl001, "{CONFIG_FILENAME}", fName)
                        created001 = created001 + 1
                    End If
                Else
                    skipped001 = skipped001 + 1
                End If
            End If
        Next i
    End If

    ' 2. Генерация 00_4 (Описание и Доп. данные) - обход по массиву файлов
    paths = Array(descPath, extraPath)
    For Each p In paths
        If FolderExistsSafe(CStr(p)) Then
            fileArr = GetFileListArray(CStr(p))
            For i = LBound(fileArr) To UBound(fileArr)
                fName = CStr(fileArr(i))
                If Left$(fName, 2) <> "~$" And LCase$(Left$(fName, 7)) <> "prompt_" And InStr(fName, ".") > 0 Then
                    promptName = "Prompt_" & fName & ".md"
                    promptPath = CStr(p) & "\" & promptName
                    If Not FileExistsSafe(promptPath) Then
                        If tpl004 <> "" Then
                            SaveTextFileUtf8 promptPath, Replace(tpl004, "{DOC_FILENAME}", fName)
                            created004 = created004 + 1
                        End If
                    Else
                        skipped004 = skipped004 + 1
                    End If
                End If
            Next i
        End If
    Next p

    ' 3. Обновление настроек (лист "Настройки")
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Настройки")
    On Error GoTo ErrHandler
    If Not ws Is Nothing Then
        found = False
        For r = 2 To 100
            If StrComp(Trim(CStr(ws.Cells(r, 1).value)), "Выжимки и анализ конфигурации", vbTextCompare) = 0 Then
                ws.Unprotect Password:="pilot"
                ws.Cells(r, 2).value = "Да"
                ws.Protect Password:="pilot"
                found = True
                Exit For
            End If
        Next r
        If Not found Then
            Mod_Util.WriteWarning "U009", "Параметр 'Выжимки и анализ конфигурации' не найден на листе 'Настройки'."
        End If
    End If

    ' 4. Логирование и уведомление (количество создано/пропущено)
    Dim msg As String
    msg = "Генерация аналитических промтов завершена." & vbCrLf & _
          "Создано промтов 00_1: " & created001 & " (пропущено: " & skipped001 & ")" & vbCrLf & _
          "Создано промтов 00_4: " & created004 & " (пропущено: " & skipped004 & ")"

    Mod_Util.WriteLog msg
    MsgBox msg, vbInformation, "Генерация аналитических промтов"

    Exit Sub

ErrHandler:
    Mod_Util.WriteError "P007", "GenerateAnalyticalPrompts_Click: " & Err.Description
    MsgBox "Ошибка: " & Err.Description, vbCritical, "Генерация аналитических промтов"
    Err.Clear
End Sub

' ============================================================================
' W-ANALYTIC-06D: Архивация и обновление настроек
' ============================================================================

' Перемещает файл из srcPath в destFolder с именем fileName.
' При коллизии добавляет суффикс _yyyy-mm-dd_hhnnss и пишет предупреждение P011.
Private Sub ArchiveFile(ByVal srcPath As String, ByVal destFolder As String, ByVal fileName As String)
    Dim destPath As String
    Dim ext As String
    Dim nameWithoutExt As String
    Dim timestamp As String
    Dim dotPos As Long

    ' W-ANALYTIC-06D-FIX: идемпотентность. Если исходник уже перемещён
    ' (повторный вызов / артефакт предыдущего прогона) - молча пропускаем.
    If Not FileExistsSafe(srcPath) Then
        Mod_Util.WriteLog "ArchiveFile: исходный файл уже перемещён, пропуск: " & srcPath
        Exit Sub
    End If

    destPath = destFolder & "\" & fileName
    If FileExistsSafe(destPath) Then
        dotPos = InStrRev(fileName, ".")
        If dotPos > 0 Then
            ext = Mid$(fileName, dotPos)
            nameWithoutExt = Left$(fileName, dotPos - 1)
        Else
            ext = "": nameWithoutExt = fileName
        End If
        timestamp = "_" & Format(Now, "yyyy-mm-dd_hhnnss")
        destPath = destFolder & "\" & nameWithoutExt & timestamp & ext
        Mod_Util.WriteWarning "P011", "Коллизия при архивации: файл " & fileName & _
            " переименован в " & nameWithoutExt & timestamp & ext
    End If

    On Error Resume Next
    Name srcPath As destPath
    If Err.Number <> 0 Then
        Mod_Util.WriteError "P012", "Ошибка перемещения файла " & srcPath & _
            " в " & destPath & ". " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' Посегментное сравнение номеров шагов (возвращает True, если step1 > step2).
' Корректно обрабатывает форматы вида "01.10" > "01.2".
Private Function IsStepGreater(ByVal step1 As String, ByVal step2 As String) As Boolean
    Dim parts1 As Variant, parts2 As Variant
    Dim i As Long, maxParts As Long
    Dim val1 As Long, val2 As Long

    ' Если step1 пуст - точно не больше
    If Trim(step1) = "" Then
        IsStepGreater = False
        Exit Function
    End If
    ' Если step2 пуст, а step1 нет - step1 считается больше (инициализация максимума)
    If Trim(step2) = "" Then
        IsStepGreater = True
        Exit Function
    End If
    
    If StrComp(Trim$(step2), "нет", vbTextCompare) = 0 Then
        IsStepGreater = True
        Exit Function
    End If

    parts1 = Split(step1, ".")
    parts2 = Split(step2, ".")
    maxParts = IIf(UBound(parts1) > UBound(parts2), UBound(parts1), UBound(parts2))
    For i = 0 To maxParts
        val1 = 0: val2 = 0
        If i <= UBound(parts1) Then
            If IsNumeric(Trim(CStr(parts1(i)))) Then val1 = CLng(Trim(CStr(parts1(i))))
        End If
        If i <= UBound(parts2) Then
            If IsNumeric(Trim(CStr(parts2(i)))) Then val2 = CLng(Trim(CStr(parts2(i))))
        End If
        If val1 > val2 Then
            IsStepGreater = True
            Exit Function
        ElseIf val1 < val2 Then
            IsStepGreater = False
            Exit Function
        End If
    Next i
    IsStepGreater = False
End Function

' Обновляет настройку LastStepNumber, только если newStep строго больше текущего.
Private Sub UpdateLastStepNumber(ByVal newStep As String)
    Dim currentStep As String
    currentStep = Mod_Util.GetSettingText("LastStepNumber", "нет")
    
    If IsStepGreater(newStep, currentStep) Then
        UpdateSetting "Номер последнего шага", newStep
        Mod_Util.WriteLog "UpdateLastStepNumber: обновлён на " & newStep
    End If
End Sub

' Проверяет заполненность "№шага/подшага" для всех строк с "Отображаемое имя".
' При полном покрытии устанавливает "Разбивка на шаги" = "Да" (автоматически не снимается).
Private Sub UpdateStepsState(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim colVisName As Long, colStepNum As Long
    Dim allFilled As Boolean
    Dim hasData As Boolean
    
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    colVisName = FindColumn(ws, COL_VIS_NAME)
    colStepNum = FindColumn(ws, COL_STEP_NUM)
    
    If colVisName = 0 Or colStepNum = 0 Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, colVisName).End(xlUp).row
    allFilled = True
    hasData = False
    
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, colVisName).value)) <> "" Then
            hasData = True
            If Trim(CStr(ws.Cells(r, colStepNum).value)) = "" Then
                allFilled = False
                Exit For
            End If
        End If
    Next r
    
    If hasData And allFilled Then
        UpdateSetting "Разбивка на шаги", "Да"
        Mod_Util.WriteLog "UpdateStepsState: установлено значение ""Да"""
    End If
End Sub

' Обновляет значение настройки на листе "Настройки" (снимает/возвращает защиту).
Private Sub UpdateSetting(ByVal settingName As String, ByVal newValue As String)
    Dim ws As Worksheet
    Dim r As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Настройки")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    ws.Unprotect Password:="pilot"
    For r = 2 To 100
        If StrComp(Trim(CStr(ws.Cells(r, 1).value)), settingName, vbTextCompare) = 0 Then
            ws.Cells(r, 2).value = newValue
            Exit For
        End If
    Next r
    ws.Protect Password:="pilot"
End Sub

