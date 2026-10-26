Attribute VB_Name = "Mod_Prompt"
Option Explicit

' ============================================================================
' Модуль промптов: формирование промптов для следующего шага и исправлений,
' чтение списка и текущей позиции шагов.
' Реализация W4 (разделы 2, 3, 8, 9, 13 спецификации Spec_Dispatcher.md).
' Коды ошибок: P001 — нет Шаги.txt, P002 — ошибка формирования промта шага,
' P003 — ошибка формирования промта исправления,
' P004 — формирование Промта №0 (декомпозиция) при отсутствии Шаги.txt
'      (W-PROMPT-GEN-STEP-00),
' P005 — ошибка сохранения текстового файла в UTF-8 (W-UTF8-EXPORT),
' P006 — предупреждение: ошибка/пустой результат чтения текстового файла
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

' ---------------------------------------------------------------------------
' КОНСТАНТЫ W-PROMPT-INVARIANTS-CODE (подстановка нового шаблона шага NN)
' ---------------------------------------------------------------------------

' Перечень плейсхолдеров нового шаблона prompt_step_template_v1.md
' (раздел 4 «Легенда плейсхолдеров» отчёта аналитика; сверено с файлом).
' Разделитель — ";". Плейсхолдеры, отсутствующие в конкретном шаблоне,
' молча пропускаются при подстановке (FillPh проверяет наличие).
Private Const TEMPLATE_PLACEHOLDERS As String = _
    "{{NN}};{{ИМЯ_ШАГА}};{{ЗАДАЧА_ШАГА_ОСОБЕННОСТИ}};{{СТРУКТУРА_ФАЙЛ}};" & _
    "{{ОПИСАНИЕ_ФАЙЛ}};{{КОНФИГУРАЦИЯ_ФАЙЛ}};{{ФАЙЛЫ_ДОП_ДАННЫХ}};{{ШАГИ_ФАЙЛ}};" & _
    "{{СПИСОК_ВХОДОВ}};{{ТАБЛИЦА_ИЗВЕСТНЫХ_GUID}};{{РЕЕСТР_ТИПОВ}};{{РЕЕСТР_ПОКРЫТИЯ}};" & _
    "{{СПИСОК_БАЗОВЫХ_ТИПОВ}};{{ПРАВИЛО_ОБЪЕКТА}};{{СПИСОК_ТИПОВ_ШАГА}};{{ТЕКУЩИЙ_ШАГ}};" & _
    "{{ДЕКОМПОЗИЦИЯ_ШАГОВ}};{{ФОРМАТ_РЕЕСТРОВ}};{{ПРОТОКОЛ_ШАГА}};{{МАРКЕРЫ}}"

' Массив известных плейсхолдеров для проверки «неизвестный {{...}} в шаблоне»
' (инициализируется один раз из TEMPLATE_PLACEHOLDERS, см. EnsureKnownPh).
Private SPLIT_KNOWN_PH() As String
Private SPLIT_KNOWN_PH_READY As Boolean

' Разовая инициализация массива известных плейсхолдеров.
Private Sub EnsureKnownPh()
    If SPLIT_KNOWN_PH_READY Then Exit Sub
    SPLIT_KNOWN_PH = Split(TEMPLATE_PLACEHOLDERS, ";")
    SPLIT_KNOWN_PH_READY = True
End Sub

' Имя файла общих настроек проекта (константы промта: списки, правила,
' форматы). Хранится рядом с книгой .xlsm или во внешней папке «Шаблоны»;
' формат строк: {{ИМЯ_ПЛЕЙСХОЛДЕРА}} = <значение>. Значения НЕ хардкодятся
' в коде — только из этого файла (требование задания).
Private Const PROJECT_SETTINGS_FILE As String = "project_settings.md"

' Маркер-страховка задачи (ЧАСТЬ D п.7): подставляется в блок 2 вместо
' постоянного текста задачи, если блок 6 пуст (типов шага нет).
Private Const TASK_NOT_SET_MARKER As String = "ЗАДАЧА НЕ ЗАДАНА — генерация запрещена"

' Заготовка блока 16 в новом шаблоне (плейсхолдер {{NN}} сохраняется до
' общего прогона подстановки — сравнение идёт по этому тексту).
Private Const CORRECTION_STUB_TEXT As String = "Корректировок для шага {{NN}} нет."

' Тексты блоков 4.1/4.2 (ЧАСТЬ B п.3).
Private Const REG_EMPTY_FIRST_STEP As String = "Реестры пусты — первый шаг (P002 штатно)"
Private Const REG_NOT_FOUND As String = "[РЕЕСТР НЕ НАЙДЕН]"

' Обязательные поля строки корректировки (ЧАСТЬ D п.6) и канон колонок
' адаптированной структуры для фильтрации блока 6 (ЧАСТЬ B п.4).
Private Const FIELD_STEP As String = "ШАГ:"
Private Const FIELD_TYPE As String = "ТИП:"
Private Const FIELD_COND As String = "УСЛОВИЕ:"
Private Const FIELD_CLAR As String = "УТОЧНЕНИЕ:"
Private Const COL_STEP As String = "Шаг"
Private Const COL_SYS_NAME As String = "Тип"
Private Const COL_VIS_NAME As String = "Отображаемое имя"

' W-PROMPT-00-SCHEMA-FIX: имя артефакта подготовки схемы (лист книги и
' префикс сохраняемого файла в папке 02_Подготовка схемы).
Private Const SCHEMA_SHEET_NAME As String = "Адаптированная_структура"
' Маска файлов артефакта в 02_Подготовка схемы.
Private Const SCHEMA_ARTIFACT_MASK As String = "Адаптированная_структура_*.xlsx"
' Подпапка архива старых версий артефакта.
Private Const SCHEMA_ARCHIVE_SUBFOLDER As String = "Архив"

' Алгоритм декомпозиции (W-CORRECTION-PACK-1, п.4.1): заменяет старые
' «Правила декомпозиции» v7.13 (пункты 10.1/10.4 и фиксированная сетка из
' 19 шагов удалены из шаблона Промта №0). Сетка шагов больше НЕ задаётся
' заранее — ИИ строит её сам по алгоритму «анализ → прогноз → шаги».
' Строковая константа для подстановки в {Decomposition_Rules} шаблона
' prompt_step_00_template.md.
Private Const DECOMPOSITION_RULES_V713 As String = _
    "Алгоритм декомпозиции (выполнять строго по порядку):" & vbCrLf & _
    "Шаг 1 - Анализ вложений: подсчитай уникальные типы, атрибуты, распределение по фазам/стадиям/этапам. Выведи сводку ""Типов: N, Атрибутов: M, Фаз: F, Стадий: S, Этапов: E"". Учти также файлы из папки 04_Дополнительные данные (ГОСТы, регламенты и т.п.), если они переданы вместе со Схемой и Описанием." & vbCrLf & _
    "Важно: Список типов в представленной Схеме (Плоской таблице) является ИСЧЕРПЫВАЮЩИМ. Тебе не нужно придумывать новые типы. Однако, список атрибутов в Схеме/Описании может быть неполным: часть атрибутов потребуется выводить из контекста (ГОСТ, логика Pilot-BIM) при генерации шагов." & vbCrLf & _
    "Шаг 2 - Прогноз объёма: определи лимит типов/атрибутов на один шаг генерации. При прогнозе количества атрибутов учитывай не только явные атрибуты из таблицы, но и закладывай резерв на: (а) выводимые атрибуты (по ГОСТу и логике), (б) унификацию повторяющихся паттернов (Договоры, Заявки, Акты), (в) параметризуемые типы. Итоговое количество шагов рассчитай с запасом +50%." & vbCrLf & _
    "Шаг 3 - Генерация списка шагов:" & vbCrLf & _
    "- ШАГ 0 - СКЕЛЕТ: все папки (фазы, стадии, этапы) + их атрибуты." & vbCrLf & _
    "- ШАГ 1 - УНИВЕРСАЛЬНЫЕ И ПАРАМЕТРИЗУЕМЫЕ типы + их атрибуты." & vbCrLf & _
    "- ШАГИ 2...N - расписание фаз/стадий/этапов: полное покрытие схемы без пропусков." & vbCrLf & _
    "- ШАГ N+1 - ФИНАЛЬНЫЙ АУДИТ." & vbCrLf & _
    "Шаг 4 - Формат вывода Шаги.txt: ""NN|Имя_папки"", разделитель - pipe (|), НЕ ""=""! Без пояснений. (NN - номер шага, Имя_папки - имя папки вида Шаг_NN_Название в 03_Результаты генерации\01_По шагам)."

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
' в «кракозябры». Поэтому сначала используется ADODB.Stream с
' Charset = "utf-8"; если чтение через UTF-8 не удалось или дало пустой
' результат — fallback на Charset = "windows-1251" (ANSI), затем — на
' стандартное Open ... For Input. При любой ошибке чтения — запись [WARN]
' в лог (код P006) и возврат пустой строки (без прерывания вызывающего).
' Если файла нет — возвращает пустую строку (вызывающий сам решает,
' считать ли это ошибкой).
Private Function SafeReadText(ByVal filePath As String) As String
    Dim res As String

    SafeReadText = ""

    ' W-PROMPT-PATHS-REFACTOR: vbNormal — строгая проверка существования
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
        ' поток мог остаться открытым — закрыть его, игнорируя ошибки
        If Not stream Is Nothing Then
            Err.Clear
            stream.Close
            Set stream = Nothing
        End If
        res = ""
    End If
    Err.Clear

    ' W-PROMPT-00-REFINE, п.4: ADODB.Stream при чтении UTF-8 возвращает
    ' символом U+FEFF (BOM) в начале строки — удаляем, чтобы BOM не
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
        Dim fnum As Integer
        Dim s As String
        fnum = FreeFile
        Open filePath For Input Access Read As #fnum
        Do While Not EOF(fnum)
            Line Input #fnum, s
            res = res & s & vbCrLf
        Loop
        Close #fnum
        If Err.Number <> 0 Then
            If fnum <> 0 Then Close #fnum
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
    ' произошла ошибка — фиксируем предупреждение в журнале ([WARN] P006)
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

' Формирует блок "ДОПОЛНИТЕЛЬНЫЕ ДАННЫЕ ПРОЕКТА" для подстановки в начало
' промта (W-CORRECTION-PACK-1, п.4.3): если папка
' "01_Исходные данные\04_Дополнительные данные" не пуста - перечисляет её
' файлы; иначе возвращает "" (блок не добавляется).
Private Function BuildExtraDataBlock(ByVal root As String) As String
    Dim filesList As String

    BuildExtraDataBlock = ""
    filesList = ListFolderFiles(root & "\01_Исходные данные\04_Дополнительные данные")
    If filesList = "" Then Exit Function

    BuildExtraDataBlock = "### ДОПОЛНИТЕЛЬНЫЕ ДАННЫЕ ПРОЕКТА" & vbCrLf & _
        "В папке 04_Дополнительные данные найдены файлы:" & vbCrLf & _
        filesList & vbCrLf & _
        "Учитывай их при генерации, если они относятся к текущему шагу." & _
        vbCrLf & vbCrLf & vbCrLf
End Function

' Формирует содержимое плейсхолдера {EXTRA_DATA_LIST} для Промта №0
' (W-PROMPT-00-REFINE, п.2/п.5; W-PROMPT-00-FORMAT-FIX, п.5): список
' имён и типов файлов из папки "01_Исходные данные\04_Дополнительные
' данные" — он подставляется в раздел "Проверка файлов" шаблона
' prompt_step_00_template.md на отдельной строке плейсхолдера.
' Заголовок "Найденные файлы доп. данных:" формируется ЗДЕСЬ (в шаблоне
' статичная строка с заголовком удалена), поэтому:
' - если файлов нет (папка пуста или отсутствует) — возвращается пустая
'   строка "" без заголовка (в итоговом markdown не остаётся висящих
'   заголовков и лишних пустых строк);
' - если файлы есть — возвращается заголовок и имена файлов с отступом.
Private Function BuildStepZeroExtraDataList(ByVal root As String) As String
    Dim filesList As String

    BuildStepZeroExtraDataList = ""
    filesList = ListFolderFiles(root & "\01_Исходные данные\04_Дополнительные данные")
    If filesList = "" Then Exit Function

    BuildStepZeroExtraDataList = "- Найденные файлы доп. данных:" & _
        vbCrLf & filesList
End Function

' Приводит текст к единым разделителям строк vbCrLf (W-UTF8-READ):
' ADODB.Stream сохраняет исходные переводы строки файла — CRLF, LF или CR.
Private Function NormalizeLineEndings(ByVal text As String) As String
    Dim res As String
    res = Replace(text, vbCrLf, vbLf)     ' CRLF -> LF
    res = Replace(res, vbCr, vbLf)        ' одиночный CR -> LF
    NormalizeLineEndings = Replace(res, vbLf, vbCrLf)  ' LF -> CRLF
End Function

' Читает текстовый файл целиком в строку (W-UTF8-READ: делегирует
' SafeReadText с автоопределением UTF-8 / Windows-1251).
' Если файла нет или чтение не удалось — возвращает пустую строку
' (вызывающий сам решает, считать ли это ошибкой).
' W-DOCS-README-UPDATE: Public — функция используется Mod_Docs.UpdateReadme
' для чтения шаблонов промтов при включении их текстов в Readme.
Public Function ReadTextFile(ByVal filePath As String) As String
    ReadTextFile = SafeReadText(filePath)
End Function

' W-PROMPT-INVARIANTS-CODE: существование именно ФАЙЛА (не папки);
' Dir(..., vbNormal) даёт ложный False для имён с кириллицей/пробелами,
' поэтому проверка выполняется через GetAttr (ошибка 76 — файла нет).
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
' чего кириллица в промтах отображается как «кракозябры» в UTF-8
' редакторах и на GitHub. Поэтому используется ADODB.Stream с
' Charset = "utf-8". Стандартный WriteText через Charset="utf-8" может
' добавить заголовок EF BB BF (BOM); чтобы этого избежать, содержимое
' перед записью очищается от ведущего U+FEFF, а после сохранения файл
' перечитывается в байтовый поток и, при наличии BOM, перезаписывается
' без первых трёх байт. При ошибке записи — запись в лог (код P005)
' и возврат без прерывания вызывающей процедуры.
Private Sub SaveTextFileUtf8(ByVal filePath As String, ByVal content As String)
    On Error GoTo ErrHandler

    Dim stream As Object

    ' Страховка: если содержимое начинается с BOM (U+FEFF) — убрать его,
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
            ' Первые три байта — BOM: перечитываем тело без него.
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

' Возвращает имя первого найденного файла по маске в папке (для Промта №0).
' Если файлов нет — возвращает "". Временные файлы Office (~$) пропускаются.
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

' W-PROMPT-00-SCHEMA-FIX: возвращает имя самого свежего файла артефакта
' "Адаптированная структура" (маска Адаптированная_структура_*.xlsx) в папке
' 02_Подготовка схемы. Файл выбирается по максимальной дате, встроенной в
' имя (формат yyyy-mm-dd_hhnnss); если даты из имён извлечь не удалось —
' берётся файл с последней датой изменения на диске. Если подходящих
' файлов несколько — возвращается последний (самый свежий), если ни одного
' — пустая строка. Временные файлы Office (~$...) пропускаются.
Private Function GetNewestSchemaArtifactName(ByVal prepPath As String) As String
    Dim fName As String
    Dim best As String
    Dim bestStamp As String
    Dim curStamp As String
    Dim bestDate As Date
    Dim curDate As Date

    GetNewestSchemaArtifactName = ""
    If Dir(prepPath, vbDirectory) = "" Then Exit Function

    fName = Dir(prepPath & "\" & SCHEMA_ARTIFACT_MASK)
    Do While fName <> ""
        If Left(fName, 2) <> "~$" Then          ' пропускаем временные файлы Office
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
                ' приоритет — дата в имени файла; при её отсутствии/равенстве —
                ' дата изменения на диске
                If curStamp > bestStamp Or _
                   (curStamp = bestStamp And curDate > bestDate) Then
                    best = fName
                    bestStamp = curStamp
                    bestDate = curDate
                End If
            End If
        End If
        fName = Dir()
    Loop
    GetNewestSchemaArtifactName = best
End Function

' W-PROMPT-00-SCHEMA-FIX: извлекает из имени файла подстроку вида
' yyyy-mm-dd_hhnnss (16 символов). Сравнивать такие подстроки как строки
' корректно, поскольку формат лексикографически упорядочен. Если шаблон в
' имени не найден — возвращается "".
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

' Собирает фрагмент схемы из листа «Адаптированная_структура» книги, сохранённой
' в папке 02_Подготовка схемы (W-PROMPT-00-SCHEMA-FIX: берётся самый свежий
' файл "Адаптированная_структура_*.xlsx"; поддержка старого имени
' "Плоская_таблица*" сохранена для книг, созданных до переименования).
' Если файла/листа нет — возвращает "" и пишет предупреждение в лог.
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
    prepPath = root & "\02_Подготовка схемы"
    If Dir(prepPath, vbDirectory) = "" Then
        Mod_Util.WriteWarning "P002", "GetSchemaFragment: папка 02_Подготовка схемы не найдена, схема не включена."
        Exit Function
    End If

    ' W-PROMPT-00-SCHEMA-FIX: ищем самый свежий файл артефакта
    ' "Адаптированная_структура_*.xlsx"; если таких нет — старый
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
' W-PROMPT-INVARIANTS-CODE: СБОРКА ПРОМТА ШАГА ПОД НОВЫЙ ШАБЛОН
' (prompt_step_template_v1.md, 17 блоков; инвариантное ядро 7–11 — жёсткий
' текст шаблона, код его не генерирует)
' ---------------------------------------------------------------------------

' Заменяет плейсхолдер {{...}} в тексте на значение и возвращает True, если
' замена выполнена. Если плейсхолдера нет в тексте — ничего не делает
' (возвращает False). Если значение пустое — подставляет явный текст
' "[НЕ ЗАПОЛНЕНО: <имя>]" и пишет WriteWarning с кодом и именем плейсхолдера
' (требование ЧАСТИ A: незаполненный плейсхолдер не оставляется в промте
' молча). Реализация через Split/Join — надёжнее встроенной Replace для
' длинных текстов (защита от ошибки 6 «Permitted exception» при больших
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
            " не удалось заполнить — подставлен явный маркер «[НЕ ЗАПОЛНЕНО]»."
    End If

    parts = Split(text, phName)
    outText = Join(parts, value)
    FillPh = True
End Function

' Подставляет ВСЕ плейсхолдеры нового шаблона промта шага NN
' (W-PROMPT-INVARIANTS-CODE, ЧАСТЬ A). Список соответствует реальному
' содержимому prompt_step_template_v1.md; плейсхолдеры, отсутствующие в
' шаблоне, молча пропускаются (FillPh проверяет InStr). Значения собираются
' в colValues по имени плейсхолдера; отсутствующее в коллекции значение
' трактуется как незаполненное ("[НЕ ЗАПОЛНЕНО: ...]").
' ВАЖНО: вызывается ДО ApplyCorrectionsBlock — плейсхолдер {{NN}} внутри
' блока 16 должен сохраниться до извлечения корректировок.
' Возвращает True, если блок 6 ({СПИСОК_ТИПОВ_ШАГА}) заполнен хотя бы одним
' типом шага (для маркер-страховки задачи, ЧАСТЬ D п.7).
Private Function FillAllPlaceholders(ByVal templateText As String, _
                                     ByVal nn As String, _
                                     ByVal stepName As String, _
                                     ByVal schemaText As String, _
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
    values("{{СТРУКТУРА_ФАЙЛ}}") = schemaText
    values("{{ОПИСАНИЕ_ФАЙЛ}}") = BuildInputFileLine(root, "01_Исходные данные\02_Описание", "*.docx", "Описание структуры ЖЦ ОКС (источник атрибутов и бизнес-логики)")
    values("{{КОНФИГУРАЦИЯ_ФАЙЛ}}") = BuildConfigFileLine(root)
    values("{{ФАЙЛЫ_ДОП_ДАННЫХ}}") = BuildExtraDataListBlock(root, nn)   ' блок 3.4
    values("{{ШАГИ_ФАЙЛ}}") = BuildStepsFileLine(root)
    values("{{РЕЕСТР_ТИПОВ}}") = BuildRegistryBlock(registriesPath, "Реестр_типов.txt", nn)     ' блок 4.1
    values("{{РЕЕСТР_ПОКРЫТИЯ}}") = BuildRegistryBlock(registriesPath, "Реестр_покрытия.txt", nn) ' блок 4.2
    values("{{СПИСОК_ТИПОВ_ШАГА}}") = BuildStepTypeList(schemaText, nn)  ' блок 6
    hasStepTypes = (Trim$(CStr(values("{{СПИСОК_ТИПОВ_ШАГА}}"))) <> "")

    ' --- Константы проекта: из общих настроек (файл project_settings.md),
    '     НЕ хардкодятся в коде (требование ЧАСТИ A / «НЕ хардкодить») -----
    LoadProjectSettings values, valuesPath

    ' Текст задачи шага: из настроек; если не задан и типов шага нет —
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

    ' Проверка «Сверься с шаблоном»: если в шаблоне остался необработанный
    ' плейсхолдер {{...}} (одинарные { } — служебные метки комментариев) —
    ' предупреждение в лог, генерацию не блокируем.
    dummy = HasUnknownPlaceholder(outResult)

    ' Если в шаблоне вообще не найдено ни одногоKnown-плейсхолдера —
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
' не входящие в известный перечень (новый плейсхолдер в шаблоне — обработать
' его в коде; пока — явный сигнал в журнале).
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
                "неизвестный плейсхолдер " & name & " — подстановка не выполнена, " & _
                "требуется доработка кода (Сверка с шаблоном, ЧАСТЬ A)."
            HasUnknownPlaceholder = True
        End If
        rest = Mid$(rest, endPos + 2)
    Loop
End Function

' Читает файл общих настроек проекта (project_settings.md) и кладёт значения
' плейсхолдеров в словарь values. Формат файла: строки вида
'   {{ИМЯ_ПЛЕЙСХОЛДЕРА}} = <значение в одну строку>
' Многострочное значение можно задать несколько раз — строки склеиваются
' через vbCrLf. Строки-комментарии (начинаются с '#') и пустые пропускаются.
' Если файл не найден — WriteWarning (константы проекта должны собираться из
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
                "{{ПРОТОКОЛ_ШАГА}}, {{МАРКЕРЫ}}) будут помечены «[НЕ ЗАПОЛНЕНО]»."
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
' "<имя файла> — <назначение в одну строку>". Назначение берётся из общего
' файла назначений "Назначения_доп_данных.txt" (строки вида
' "<имя файла> = <назначение>", регистр имён не значим); если назначение не
' задано — только имя файла. При отсутствии файлов — заготовка
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
                        res = res & fName & " — " & purpose
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

' Элемент пути — папка? (без ошибок Dir по кириллическим путям)
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

' Строка вида "<имя файла> — <назначение>" для входного файла из папки srcRel
' (поиск по маске mask). Если файл не найден — "[НЕ НАЙДЕН: <папка>\<маска>]".
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
    BuildInputFileLine = fName & " — " & purpose
End Function

' Строка для {{КОНФИГУРАЦИЯ_ФАЙЛ}}: имя файла конфигурации через
' Mod_Util.FindConfigFile (.xml > .txt; только .pilotcfx — предупреждение).
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
    BuildConfigFileLine = cfgName & " — Конфигурация Pilot-BIM (примеры XML «Дополнительных параметров», справочник базовых типов)"
End Function

' Строка для {{ШАГИ_ФАЙЛ}}: путь к Шаги.txt (он же источник сетки шагов).
Private Function BuildStepsFileLine(ByVal root As String) As String
    Dim stepsFile As String

    stepsFile = root & "\04_Промты\Шаги.txt"
    If Not FileExistsSafe(stepsFile) Then
        BuildStepsFileLine = "[НЕ НАЙДЕН: 04_Промты\Шаги.txt]"
        Exit Function
    End If
    BuildStepsFileLine = "04_Промты\Шаги.txt — сетка шагов декомпозиции (формат строк NN|Имя_папки)"
End Function

' Блоки 4.1/4.2 ({{РЕЕСТР_ТИПОВ}}, {{РЕЕСТР_ПОКРЫТИЯ}}):
'  - шаг 00 — заготовка "Реестры пусты — первый шаг (P002 штатно)";
'  - шаг > 00 — содержимое накопленного файла реестра;
'  - файл отсутствует на шаге > 00 — WriteWarning (P002) и явный текст
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
            ") — подставлен текст «" & REG_NOT_FOUND & "»."
        BuildRegistryBlock = REG_NOT_FOUND
        Exit Function
    End If

    content = NormalizeLineEndings(ReadTextFile(filePath))
    If Trim$(content) = "" Then
        Mod_Util.WriteWarning "P002", "BuildRegistryBlock: " & fileName & _
            " пуст на шаге " & nn & " — подставлен текст «" & REG_NOT_FOUND & "»."
        BuildRegistryBlock = REG_NOT_FOUND
        Exit Function
    End If
    BuildRegistryBlock = content
End Function

' Блок 6 ({{СПИСОК_ТИПОВ_ШАГА}}): типы шага NN из текста адаптированной
' структуры (GetSchemaFragment, колонки через " | ", строка 1 — заголовки).
' Фильтрация по колонке "Шаг": совпадение числа (NN) или префикса "NN/"
' (диапазон шагов). Порядок строк — как в структуре (без сортировки).
' Формат вывода: "Видимое имя типа | Тип" (системное имя в отдельной колонке
' НЕ выводится; "Тип" — системное имя, канон колонок листа «Типы»).
' Если колонка "Шаг" в структуре отсутствует — WriteWarning и пустая строка
' (плейсхолдер получит маркер "[НЕ ЗАПОЛНЕНО]", блок 2 — маркер-страховку).
Private Function BuildStepTypeList(ByVal schemaText As String, ByVal nn As String) As String
    Dim lines() As String
    Dim headers() As String
    Dim i As Long
    Dim cols() As String
    Dim colStep As Long
    Dim colSys As Long
    Dim colVis As Long
    Dim c As Long
    Dim headerFound As Boolean
    Dim res As String
    Dim stepVal As String

    BuildStepTypeList = ""
    If Trim$(schemaText) = "" Then Exit Function

    lines = Split(NormalizeLineEndings(schemaText), vbCrLf)

    ' 1. Строка заголовков — первая непустая
    headerFound = False
    For i = LBound(lines) To UBound(lines)
        If Trim$(lines(i)) <> "" Then
            headers = Split(lines(i), " | ")
            headerFound = True
            Exit For
        End If
    Next i
    If Not headerFound Then Exit Function

    ' 2. Индексы колонок по канону
    colStep = -1
    colSys = -1
    colVis = -1
    For c = LBound(headers) To UBound(headers)
        Select Case Trim$(headers(c))
            Case COL_STEP: colStep = c
            Case COL_SYS_NAME: colSys = c
            Case COL_VIS_NAME: colVis = c
        End Select
    Next c

    If colSys < 0 Or colVis < 0 Then
        Mod_Util.WriteWarning "P002", "BuildStepTypeList: в структуре не найдены " & _
            "колонки «" & COL_SYS_NAME & "»/«" & COL_VIS_NAME & "» — блок 6 не собран."
        Exit Function
    End If
    If colStep < 0 Then
        Mod_Util.WriteWarning "P002", "BuildStepTypeList: в структуре нет колонки «" & _
            COL_STEP & "» (фильтрация типов по шагу невозможна) — блок 6 не собран."
        Exit Function
    End If

    ' 3. Отбор строк шага NN в порядке структуры (без сортировки)
    For i = LBound(lines) + 1 To UBound(lines)
        If Trim$(lines(i)) <> "" Then
            cols = Split(lines(i), " | ")
            If UBound(cols) >= colStep And UBound(cols) >= colSys And UBound(cols) >= colVis Then
                stepVal = Trim$(cols(colStep))
                If StepValueMatches(stepVal, nn) Then
                    If res <> "" Then res = res & vbCrLf
                    res = res & Trim$(cols(colVis)) & " | " & Trim$(cols(colSys))
                End If
            End If
        End If
    Next i

    BuildStepTypeList = res
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

    ' диапазоны "NN/MM" и "NN-MM": принадлежность, если NN — граница
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
        If closePos = 0 Then Exit Do           ' незакрытый комментарий — оставляем как есть
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
' с "- ", проходят валидацию скоупа (ЧАСТЬ D п.6): чужие шаги (ШАГ ≠ nn) и
' некорректные по формату — отбрасываются с логом; если после фильтрации не
' осталось ни одной корректной и были отброшенные — WriteWarning.
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
                    "некорректная по формату корректировка (нужны «- » и поля " & _
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
' «были ли пользовательские строки»).
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
'  3. если пользовательских строк нет — остаётся заготовка шаблона
'     "Корректировок для шага {{NN}} нет." ({{NN}} затем превратится в NN);
'  4. если есть — заготовка заменяется списком корректировок (без дублирования).
' Вызывается ДО FillAllPlaceholders — плейсхолдер {{NN}} в тексте ещё жив.
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
                    " после валидации не осталось — используется заготовка шаблона."
            End If
        End If
    End If

    ' --- 2. Внешний файл корректировок (04_Промты\corrections.md) ----------
    corrFileText = ReadTextFile(root & "\04_Промты\corrections.md")
    If corrFileText <> "" Then
        corrSection = ExtractSection16(corrFileText)
        If corrSection = "" Then corrSection = corrFileText   ' файл без заголовка — весь текст
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
            " нет — оставлена заготовка шаблона (блок 16)."
        Exit Sub
    End If

    ' --- 3. Замена заготовки на скорректированный список --------------------
    ' Ищем строку заготовки (плейсхолдер {{NN}} ещё не подставлен).
    stubPos = InStr(resultText, CORRECTION_STUB_TEXT)
    If stubPos = 0 Then
        ' заготовка изменена вручную в шаблоне — вставляем список перед
        ' следующим заголовком "## " после начала секции 16
        Dim sec16Pos As Long
        sec16Pos = InStr(resultText, "## 16.")
        If sec16Pos = 0 Then
            Mod_Util.WriteWarning "P002", "ApplyCorrectionsBlock: секция 16 не найдена " & _
                "в шаблоне — корректировки НЕ вставлены."
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
' ПУБЛИЧЕСКИЕ ПРОЦЕДУРЫ И ФУНКЦИИ
' ---------------------------------------------------------------------------

' Читает список шагов (<корень>\04_Промты\Шаги.txt) и возвращает его
' коллекцией. Каждая запись — строка вида "NN|Имя_папки".
' W-UTF8-READ: файл читается целиком через SafeReadText (автоопределение
' UTF-8 / Windows-1251), затем разбивается на строки Split(text, vbCrLf);
' пустые строки пропускаются.
' W-CORRECTION-PACK-1, п.4.4: основной формат строки — "NN|Имя_папки"
' (разделитель — pipe). Устаревший вариант "NN = Имя_папки" принимается с
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

    ' файла нет ? P001, MsgBox, пустая коллекция
    If Len(Dir(filePath, vbNormal)) = 0 Then
        Mod_Util.WriteError "P001", "ReadStepList: не найден Шаги.txt"
        MsgBox "Ошибка P001: не найден файл Шаги.txt (" & vbCrLf & filePath & ")", _
               vbCritical, "Промты"
        Set ReadStepList = col
        Exit Function
    End If

    ' W-UTF8-READ: чтение всего файла с автоопределением кодировки
    text = SafeReadText(filePath)
    If text = "" Then
        ' SafeReadText уже записал [WARN] в лог — здесь только предупреждаем
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
                ' устаревший формат "NN = Имя_папки" — принимаем с
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
    ' Любая ошибка чтения — [WARN] в лог, пустая коллекция без прерывания.
    Mod_Util.WriteWarning "P006", "ReadStepList: ошибка чтения Шаги.txt. Ошибка: " & Err.Description
    MsgBox "Предупреждение P006: не удалось прочитать файл Шаги.txt." & vbCrLf & Err.Description, _
           vbExclamation, "Промты"
    Err.Clear
    Set ReadStepList = col
End Function

' Возвращает номер текущего шага — последней существующей папки Шаг_*
' в 03_Результаты генерации\01_По шагам\ согласно порядку из Шаги.txt.
' Если папок нет — возвращает "".
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
    Dim templatePath As String   ' W-PROMPT-PATH-FIX: полный путь к шаблону промта шага
    Dim i As Long
    Dim j As Long

    On Error GoTo ErrHandler

    ' W-CORRECTION-PACK-1, п.2.3: блокировка генерации при неготовых исходниках
    If Not Mod_Folders.CheckSourceFolders() Then
        Mod_Util.WriteError "F005", "MakeNextStepPrompt: Исходные данные не готовы. Генерация остановлена."
        MsgBox "ERROR [F005]: Исходные данные не готовы. Генерация остановлена.", _
               vbCritical, "Промт следующего шага"
        Exit Sub
    End If

    root = NormPath(Mod_Util.GetProjectRoot)
    stepsPath = root & "\03_Результаты генерации\01_По шагам"
    promptsPath = root & "\04_Промты"
    readyPath = root & "\03_Результаты генерации\04_Промты_готовые"
    registriesPath = root & "\03_Результаты генерации\03_Реестры"

    ' 0. W-PROMPT-GEN-STEP-00: если Шаги.txt отсутствует — вместо промта
    '    следующего шага формируем Промт №0 (декомпозиция) и выходим.
    If Len(Dir(promptsPath & "\Шаги.txt", vbNormal)) = 0 Then
        GenerateStepZeroPrompt
        Exit Sub
    End If

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

        ' читаем шаблон того же уровня (04_Промты\Шаблоны\prompt_step_template_v1.md)
        ' W-PROMPT-PATH-FIX: шаблон лежит в подпапке "Шаблоны", а не в корне "04_Промты".
        templatePath = promptsPath & "\Шаблоны\" & TEMPLATE_STEP_V1_NAME
        Mod_Util.WriteLog "[i] Поиск шаблона шага по пути: " & templatePath
        If Len(Dir(templatePath, vbNormal)) = 0 Then
            template = ""
        Else
            template = ReadTextFile(templatePath)
        End If
        If template = "" Then
            Mod_Util.WriteError "P002", "MakeNextStepPrompt: не найден шаблон " & TEMPLATE_STEP_V1_NAME & _
                " (проверен путь: " & templatePath & ")"
            MsgBox "Ошибка P002: не найден шаблон " & TEMPLATE_STEP_V1_NAME & _
                   " в папке 04_Промты\Шаблоны.", _
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

        ' W-CORRECTION-PACK-1, п.4.3: блок доп. данных перед текстом промта
        result = BuildExtraDataBlock(root) & template
        result = ReplaceIfPresent(result, "{NN}", "99")
        result = ReplaceIfPresent(result, "{STEP_NAME}", "Шаг_99_Финал")
        result = ReplaceIfPresent(result, "{REGISTRY_TYPES}", regTypes)
        result = ReplaceIfPresent(result, "{REGISTRY_COVERAGE}", regCoverage)
        result = ReplaceIfPresent(result, "{SCHEMA}", schema)

        EnsureFolder stepsPath & "\Шаг_99_Финал"
        EnsureFolder readyPath
        SaveTextFileUtf8 readyPath & "\Промт_финального_аудита.md", result

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

    ' 5. Шаблон промта шага (04_Промты\Шаблоны\prompt_step_template_v1.md)
    ' W-PROMPT-PATH-FIX: шаблон лежит в подпапке "Шаблоны", а не в корне "04_Промты".
    templatePath = promptsPath & "\Шаблоны\" & TEMPLATE_STEP_V1_NAME
    Mod_Util.WriteLog "[i] Поиск шаблона шага по пути: " & templatePath
    If Len(Dir(templatePath, vbNormal)) = 0 Then
        template = ""
    Else
        template = ReadTextFile(templatePath)
    End If
    If template = "" Then
        Mod_Util.WriteError "P002", "MakeNextStepPrompt: не найден шаблон " & TEMPLATE_STEP_V1_NAME & _
            " (проверен путь: " & templatePath & ")"
        MsgBox "Ошибка P002: не найден шаблон " & TEMPLATE_STEP_V1_NAME & _
               " в папке 04_Промты\Шаблоны.", _
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

    ' 7. Фрагмент схемы из 02_Подготовка схемы (Адаптированная_структура_*.xlsx, если есть)
    schema = GetSchemaFragment()

    ' 8. Подстановка плейсхолдеров (только тех, что есть в шаблоне).
    '    W-CORRECTION-PACK-1, п.4.3: если папка 04_Дополнительные данные не
    '    пуста — в начало промта вставляется блок "ДОПОЛНИТЕЛЬНЫЕ ДАННЫЕ ПРОЕКТА".
    result = BuildExtraDataBlock(root) & template
    result = ReplaceIfPresent(result, "{NN}", nextNN)
    result = ReplaceIfPresent(result, "{STEP_NAME}", nextName)
    result = ReplaceIfPresent(result, "{REGISTRY_TYPES}", regTypes)
    result = ReplaceIfPresent(result, "{REGISTRY_COVERAGE}", regCoverage)
    result = ReplaceIfPresent(result, "{SCHEMA}", schema)

    ' 9. Папка шага создаётся только по факту формирования промта (раздел 3)
    stepFolder = stepsPath & "\" & nextName
    ' Создаём папку шага как маркер для ReadCurrentStep,
    ' сам промт сохраняется в readyPath.
    EnsureFolder stepFolder

    ' 10. Сохранение готового промта в 04_Промты_готовые
    ' (централизованное хранение готовых промтов)
    EnsureFolder readyPath
    promptFile = readyPath & "\Промт_шаг_" & nextNN & ".md"
    SaveTextFileUtf8 promptFile, result

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

' Формирует Промт №0 (декомпозиция) при отсутствии Шаги.txt
' (W-PROMPT-GEN-STEP-00):
' 1. Собирает пути к папкам 02_Подготовка схемы / 01_Исходные данные\02_Описание /
'    03_Конфигурация и находит реальные имена файлов через Dir по маскам
'    (Адаптированная_структура_*.xlsx; *.docx; *.xml или *.pilotcfx).
' 2. Если файла нет — MsgBox "Отсутствует файл в папке X", [WARN] в лог,
'    в промт подставляется плейсхолдер "НЕ НАЙДЕН", работа продолжается.
'    ИСКЛЮЧЕНИЕ (W-PROMPT-00-SCHEMA-FIX): Схема (адаптированная структура)
'    является обязательной — если в 02_Подготовка схемы нет ни одного файла
'    "Адаптированная_структура_*.xlsx", генерация Промта №0 останавливается.
' 3. Читает шаблон prompt_step_00_template.md из 04_Промты\Шаблоны,
'    подставляет {SCHEMA_FILENAME}, {DESCRIPTION_FILENAME},
'    {CONFIG_FILENAME} и {Decomposition_Rules} (правила декомпозиции v7.13).
' 4. Сохраняет результат в 04_Промты\Промт_шаг_00_Декомпозиция.md.
Private Sub GenerateStepZeroPrompt()
    Dim root As String
    Dim promptsPath As String
    Dim prepFolder As String            ' 02_Подготовка схемы (артефакт-схема)
    Dim descFolder As String            ' 01_Исходные данные\02_Описание
    Dim configFolder As String          ' 01_Исходные данные\03_Конфигурация
    Dim templatePath As String          ' 04_Промты\Шаблоны\prompt_step_00_template.md
    Dim outFile As String               ' 04_Промты\Промт_шаг_00_Декомпозиция.md
    Dim schemaName As String            ' найденное имя файла схемы
    Dim descName As String              ' найденное имя файла описания
    Dim configName As String            ' найденное имя файла конфигурации
    Dim missingList As String           ' перечень отсутствующих входных файлов
    Dim template As String
    Dim result As String

    On Error GoTo ErrHandler

    ' W-CORRECTION-PACK-1, п.2.3: блокировка генерации при неготовых исходниках
    If Not Mod_Folders.CheckSourceFolders() Then
        Mod_Util.WriteError "F005", "GenerateStepZeroPrompt: Исходные данные не готовы. Генерация остановлена."
        MsgBox "ERROR [F005]: Исходные данные не готовы. Генерация остановлена.", _
               vbCritical, "Промт №0 (декомпозиция)"
        Exit Sub
    End If

    root = NormPath(Mod_Util.GetProjectRoot)
    promptsPath = root & "\04_Промты"
    prepFolder = root & "\02_Подготовка схемы"
    descFolder = root & "\01_Исходные данные\02_Описание"
    configFolder = root & "\01_Исходные данные\03_Конфигурация"
    templatePath = promptsPath & "\Шаблоны\" & TEMPLATE_STEP0_NAME
    outFile = promptsPath & "\" & STEP0_PROMPT_FILE

    Mod_Util.WriteLog "GenerateStepZeroPrompt: Шаги.txt не найден, формирую Промт №0 (декомпозиция)"

    ' 1. Реальные имена входных файлов:
    '    W-PROMPT-00-SCHEMA-FIX: схема (адаптированная структура) берётся из
    '    папки 02_Подготовка схемы — самый свежий файл по дате в имени
    '    (при равенстве/отсутствии даты в имени — по дате изменения); если
    '    файлов несколько, берётся последний; если файлов нет — ошибка и
    '    остановка генерации.
    '    описание — *.docx; конфигурация — через FindConfigFile
    '    (валидны .xml и .txt; если только .pilotcfx — предупреждение от
    '    FindConfigFile, генерация НЕ блокируется).
    schemaName = GetNewestSchemaArtifactName(prepFolder)

    ' Обязательность Схемы (W-PROMPT-00-SCHEMA-FIX): без адаптированной
    ' структуры Промт №0 не формируется.
    If schemaName = "" Then
        Mod_Util.WriteError "P004", "[!] Отсутствует обязательный файл: Адаптированная структура" & _
                                     " (папка 02_Подготовка схемы, маска " & SCHEMA_ARTIFACT_MASK & ")."
        MsgBox "[!] Отсутствует обязательный файл: Адаптированная структура" & vbCrLf & _
               "Папка: " & prepFolder & vbCrLf & _
               "Маска: " & SCHEMA_ARTIFACT_MASK & vbCrLf & vbCrLf & _
               "Выполните подготовку схемы (Собрать плоскую структуру) и повторите попытку." & vbCrLf & _
               "Генерация Промта №0 остановлена.", _
               vbCritical, "Промт №0 (декомпозиция)"
        Exit Sub
    End If

    descName = FirstMatchFile(descFolder, "*.docx")

    configName = Mod_Util.FindConfigFile(configFolder)
    If configName <> "" Then configName = Mid$(configName, InStrRev(configName, "\") + 1)

    ' 2. Отсутствующие файлы — MsgBox + [!] (FormatMissingFile) + плейсхолдер
    '    "НЕ НАЙДЕН"; формирование Промта №0 при этом НЕ прерывается
    If descName = "" Then
        Mod_Util.WriteWarning "P004", Mod_Util.FormatMissingFile("01_Исходные данные\02_Описание\*.docx")
        If missingList <> "" Then missingList = missingList & ", "
        missingList = missingList & "02_Описание"
        descName = STEP0_NOT_FOUND
    End If
    If configName = "" Then
        Mod_Util.WriteWarning "P004", Mod_Util.FormatMissingFile("01_Исходные данные\03_Конфигурация\*.xml или *.txt")
        If missingList <> "" Then missingList = missingList & ", "
        missingList = missingList & "03_Конфигурация"
        configName = STEP0_NOT_FOUND
    End If
    If missingList <> "" Then
        MsgBox "[!] Отсутствует обязательный файл: " & missingList & "." & vbCrLf & _
               "В Промте №0 вместо имени подставлено """ & STEP0_NOT_FOUND & """.", _
               vbExclamation, "Промт №0 (декомпозиция)"
    End If

    ' 3. Шаблон Промта №0 из 04_Промты\Шаблоны
    template = ReadTextFile(templatePath)
    If template = "" Then
        Mod_Util.WriteError "P004", "GenerateStepZeroPrompt: не найден шаблон " & templatePath
        MsgBox "Ошибка P004: не найден шаблон " & TEMPLATE_STEP0_NAME & _
               " в папке 04_Промты\Шаблоны.", vbCritical, "Промт №0 (декомпозиция)"
        Exit Sub
    End If

    ' 4. Подстановка плейсхолдеров реальными именами файлов и правилами v7.13
    result = template
    result = ReplaceIfPresent(result, "{SCHEMA_FILENAME}", schemaName)
    result = ReplaceIfPresent(result, "{DESCRIPTION_FILENAME}", descName)
    result = ReplaceIfPresent(result, "{CONFIG_FILENAME}", configName)
    ' W-PROMPT-00-REFINE, п.2/п.5: список файлов доп. данных (папка
    ' 01_Исходные данные\04_Дополнительные данные) подставляется в блок
    ' "Проверка файлов" шаблона; если папка пуста — плейсхолдер заменяется
    ' на пустую строку.
    result = ReplaceIfPresent(result, "{EXTRA_DATA_LIST}", BuildStepZeroExtraDataList(root))
    result = ReplaceIfPresent(result, "{Decomposition_Rules}", DECOMPOSITION_RULES_V713)

    ' 5. Сохранение результата в 04_Промты
    EnsureFolder promptsPath
    SaveTextFileUtf8 outFile, result

    Mod_Util.WriteLog "GenerateStepZeroPrompt: сохранён " & outFile

    ' 6. Инструкция пользователю
    MsgBox "Файл Шаги.txt не найден. Сформирован Промт №0:" & vbCrLf & _
           outFile & vbCrLf & vbCrLf & _
           "Загрузите его в ИИ, получите Шаги.txt и сохраните в 04_Промты.", _
           vbInformation, "Промт №0 (декомпозиция)"
    Exit Sub

ErrHandler:
    Mod_Util.WriteError "P004", "GenerateStepZeroPrompt: ошибка формирования Промта №0. Ошибка: " & Err.Description
    MsgBox "Ошибка P004: не удалось сформировать Промт №0 (декомпозицию)." & vbCrLf & _
           Err.Description, vbCritical, "Промт №0 (декомпозиция)"
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
