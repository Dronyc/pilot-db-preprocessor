Attribute VB_Name = "Mod_Folders"
Option Explicit

' ============================================================================
' Модуль работы с папками проекта: создание, проверка, папки шагов.
' Реализация W4 (разделы 2, 3, 4 спецификации Spec_Dispatcher.md).
' Коды ошибок модуля:
'   F001 - папка проекта не определена (книга .xlsm не сохранена);
'   F002 - ошибка создания папки (MkDir);
'   F003 - ошибка копирования файла шаблона (FileCopy);
'   F004 - внешняя папка "Шаблоны" не найдена (предупреждение);
'   F005 - несоответствие исходных файлов в CheckSourceFolders.
' Пути строятся от корня проекта — папки, где лежит .xlsm
' (ThisWorkbook.Path). Никаких настроек и диалогов выбора.
' ============================================================================

' ---------------------------------------------------------------------------
' КОНСТАНТЫ И ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (Private)
' ---------------------------------------------------------------------------

' Относительные пути папок структуры проекта (раздел 2 спецификации).
Private Const FOLDER_LOGS As String = "05_Логи"                          ' создаётся ПЕРВОЙ
Private Const FOLDER_SOURCE As String = "01_Исходные данные"
Private Const FOLDER_SCHEMA As String = "01_Исходные данные\01_Схема"
Private Const FOLDER_DESC As String = "01_Исходные данные\02_Описание"
Private Const FOLDER_CONFIG As String = "01_Исходные данные\03_Конфигурация"
Private Const FOLDER_PREP As String = "02_Подготовка схемы"
Private Const FOLDER_RESULTS As String = "03_Результаты генерации"
Private Const FOLDER_STEPS As String = "03_Результаты генерации\01_По шагам"
Private Const FOLDER_SUMMARY As String = "03_Результаты генерации\02_Сводные"
Private Const FOLDER_REGISTRY As String = "03_Результаты генерации\03_Реестры"
Private Const FOLDER_READY_PROMPTS As String = "03_Результаты генерации\04_Промты_готовые"
Private Const FOLDER_PROMPTS As String = "04_Промты"
Private Const FOLDER_PROMPT_TEMPLATES As String = "04_Промты\Шаблоны"   ' папка шаблонов промтов
Private Const FOLDER_EXTERNAL_TEMPLATES As String = "Шаблоны"           ' внешняя папка рядом с .xlsm

' Папка существует?
Private Function FolderExists(ByVal fullPath As String) As Boolean
    FolderExists = (Dir(fullPath, vbDirectory) <> "")
End Function

' Создаёт цепочку папок по полному пути (аналог mkdir -p).
' Существующие папки не трогаются. Препроцессор никогда ничего не удаляет.
Private Sub EnsureFolder(ByVal folderPath As String)
    Dim parts() As String
    Dim cur As String
    Dim i As Long

    If Dir(folderPath, vbDirectory) <> "" Then Exit Sub

    parts = Split(folderPath, "\")
    cur = parts(0)                        ' диск или UNC-префикс
    For i = 1 To UBound(parts)
        If parts(i) <> "" Then
            cur = cur & "\" & parts(i)
            If Dir(cur, vbDirectory) = "" Then MkDir cur
        End If
    Next i
End Sub

' Считает количество файлов с указанным расширением в папке.
' Расширение передаётся в виде ".xlsx" и т.п. Учитываются только файлы.
Private Function CountFilesByExt(ByVal folderPath As String, ByVal ext As String) As Long
    Dim cnt As Long
    Dim Name As String

    CountFilesByExt = 0
    If Not FolderExists(folderPath) Then Exit Function

    Name = Dir(folderPath & "\*" & ext)
    Do While Name <> ""
        ' Dir без атрибутов возвращает только файлы (не папки)
        cnt = cnt + 1
        Name = Dir()
    Loop
    CountFilesByExt = cnt
End Function

' ---------------------------------------------------------------------------
' ПУБЛИЧЕСКИЕ ПРОЦЕДУРЫ И ФУНКЦИИ
' ---------------------------------------------------------------------------

' Создаёт структуру папок проекта (раздел 2 спецификации).
' 05_Логи создаётся ПЕРВОЙ - она нужна для записи журнала.
' Существующие папки не трогаются. Препроцессор никогда ничего не удаляет.
Public Sub CreateProjectFolders()
    Dim root As String
    Dim relList As Variant          ' относительные пути папок
    Dim i As Long
    Dim full As String

    On Error GoTo ErrHandler

    root = Mod_Util.GetProjectRoot()
    If root = "" Then
        Mod_Util.WriteError "F001", "CreateProjectFolders: не удалось определить папку проекта (книга не сохранена?)"
        MsgBox "Ошибка F001: не удалось определить папку проекта. Сохраните файл .xlsm и повторите.", _
               vbCritical, "Папки проекта"
        Exit Sub
    End If


    ' Порядок важен: 05_Логи - первая (нужна для лога), затем остальные.
    relList = Array(FOLDER_LOGS, _
                    FOLDER_SOURCE, _
                    FOLDER_SCHEMA, _
                    FOLDER_DESC, _
                    FOLDER_CONFIG, _
                    FOLDER_PREP, _
                    FOLDER_RESULTS, _
                    FOLDER_STEPS, _
                    FOLDER_SUMMARY, _
                    FOLDER_REGISTRY, _
                    FOLDER_READY_PROMPTS, _
                    FOLDER_PROMPTS)

    For i = LBound(relList) To UBound(relList)
        full = root & "\" & CStr(relList(i))
        If Not FolderExists(full) Then
            MkDir full              ' при ошибке уйдём в ErrHandler (F002)
        End If
    Next i

    ' Подпапка шаблонов промтов: "04_Промты\Шаблоны" (через EnsureFolder).
    EnsureFolder root & "\" & FOLDER_PROMPT_TEMPLATES

    Mod_Util.WriteLog "CreateProjectFolders: структура создана/проверена"

    ' Копирование шаблонов из внешней папки "Шаблоны" (рядом с .xlsm).
    ' Отсутствие папки или ошибки записи не прерывают создание структуры.
    CopyPromptTemplates root

    MsgBox "Структура папок проекта создана/проверена.", vbInformation, "Папки проекта"
    Exit Sub

ErrHandler:
    ' пишем ошибку с указанием конкретной папки, в которой произошёл сбой
    Mod_Util.WriteError "F002", "CreateProjectFolders: ошибка создания папки """ & _
                                root & "\" & CStr(relList(i)) & """. Ошибка: " & Err.Description
    MsgBox "Ошибка F002: не удалось создать папку." & vbCrLf & Err.Description, _
           vbCritical, "Папки проекта"
    Err.Clear
End Sub

' Копирует все файлы из внешней папки "Шаблоны" (в корне проекта, рядом с
' .xlsm) во внутреннюю папку "04_Промты\Шаблоны".
' * Если внешняя папка "Шаблоны" отсутствует - выводится MsgBox и в лог
'   пишется предупреждение; создание остальной структуры НЕ прерывается.
' * Ошибки отдельных FileCopy (например, нет прав на запись) фиксируются
'   в логе, но не приводят к аварийному завершению (On Error Resume Next).
Private Sub CopyPromptTemplates(ByVal root As String)
    Dim srcPath As String             ' внешняя папка "Шаблоны"
    Dim dstPath As String             ' "04_Промты\Шаблоны"
    Dim fileName As String
    Dim copiedCount As Long
    Dim failedCount As Long

    srcPath = root & "\" & FOLDER_EXTERNAL_TEMPLATES
    dstPath = root & "\" & FOLDER_PROMPT_TEMPLATES

    If Not FolderExists(srcPath) Then
        Mod_Util.WriteWarning "F004", "Mod_Folders: папка ""Шаблоны"" не найдена, " & _
                                      "копирование шаблонов пропущено."
        MsgBox "Папка 'Шаблоны' не найдена. Шаблоны промтов не скопированы. " & _
               "Вы можете восстановить их из файла Readme.", _
               vbExclamation, "Шаблоны промтов"
        Exit Sub
    End If

    ' Папка назначения должна существовать (создана в CreateProjectFolders).
    EnsureFolder dstPath

    ' Копирование с пофайловой фиксацией ошибок: отсутствие прав на запись
    ' не должно ломать создание структуры.
    On Error Resume Next
    fileName = Dir(srcPath & "\*")
    Do While fileName <> ""
        Err.Clear
        FileCopy srcPath & "\" & fileName, dstPath & "\" & fileName
        If Err.Number <> 0 Then
            failedCount = failedCount + 1
            Mod_Util.WriteError "F003", "CopyPromptTemplates: не удалось скопировать файл """ & _
                                        fileName & """. Ошибка: " & Err.Description
            Err.Clear
        Else
            copiedCount = copiedCount + 1
        End If
        fileName = Dir()
    Loop
    On Error GoTo 0

    Mod_Util.WriteLog "CopyPromptTemplates: скопировано файлов: " & copiedCount & _
                      ", с ошибками: " & failedCount
End Sub

' Проверяет наличие исходных папок и файлов в них (кнопка 2, раздел 4):
'   * 01_Схема        - ровно один .xlsx;
'   * 02_Описание     - ровно один .docx;
'   * 03_Конфигурация - ровно один .pilotcfx.
' Все несоответствия пишутся в лог предупреждениями (F005).
' Возвращает True, если все условия выполнены, иначе False.
Public Function CheckSourceFolders() As Boolean
    Dim root As String
    Dim folders As Variant
    Dim exts As Variant
    Dim i As Long
    Dim cnt As Long
    Dim folderFull As String
    Dim result As Boolean

    On Error GoTo ErrHandler

    result = True

    root = Mod_Util.GetProjectRoot()
    If root = "" Then
        Mod_Util.WriteError "F001", "CheckSourceFolders: папка проекта не определена (книга .xlsm не сохранена?)"
        MsgBox "Ошибка F001: не удалось определить папку проекта." & vbCrLf & _
           "Сохраните файл .xlsm и повторите.", _
           vbCritical, "Проверка исходных данных"
        CheckSourceFolders = False
        Exit Function
    End If

    ' три проверяемые папки: имя и требуемое расширение
    folders = Array(FOLDER_SCHEMA, FOLDER_DESC, FOLDER_CONFIG)
    exts = Array(".xlsx", ".docx", ".pilotcfx")

    For i = LBound(folders) To UBound(folders)
        folderFull = root & "\" & CStr(folders(i))
        cnt = CountFilesByExt(folderFull, CStr(exts(i)))

        ' каждая папка - ровно один файл нужного расширения
        If cnt = 0 Then
            Mod_Util.WriteWarning "F005", "CheckSourceFolders: " & CStr(folders(i)) & _
                                       " - файл " & CStr(exts(i)) & " не найден"
            result = False
        ElseIf cnt > 1 Then
            Mod_Util.WriteWarning "F005", "CheckSourceFolders: " & CStr(folders(i)) & _
                                       " - найдено " & cnt & " файлов " & CStr(exts(i)) & _
                                       ", ожидается ровно 1"
            result = False
        End If
    Next i

    Mod_Util.WriteLog "CheckSourceFolders: результат " & IIf(result, "True", "False")
    CheckSourceFolders = result
    Exit Function

ErrHandler:
    Mod_Util.WriteError "F005", "CheckSourceFolders: ошибка проверки. Ошибка: " & Err.Description
    MsgBox "Ошибка F005: не удалось выполнить проверку исходных папок." & vbCrLf & _
           Err.Description, vbCritical, "Проверка исходных данных"
    Err.Clear
    CheckSourceFolders = False
End Function

' Возвращает путь к папке указанного шага: ищет в
' <корень>\03_Результаты генерации\01_По шагам папку с префиксом
' "Шаг_<stepNumber>_". Если корень или папка не найдены - возвращает "".
Public Function GetStepFolder(ByVal stepNumber As String) As String
    Dim root As String
    Dim basePath As String
    Dim Name As String
    Dim prefix As String
    Dim attrs As Long

    On Error GoTo ErrHandler

    GetStepFolder = ""

    root = Mod_Util.GetProjectRoot()
    If root = "" Then Exit Function             ' корень не задан - пустая строка

    basePath = root & "\" & FOLDER_STEPS
    If Not FolderExists(basePath) Then Exit Function

    prefix = "Шаг_" & stepNumber & "_"
    Name = Dir(basePath & "\" & prefix & "*", vbDirectory)
    Do While Name <> ""
        ' Dir с vbDirectory может вернуть и файлы - пропускаем только папки
        attrs = 0
        On Error Resume Next
        attrs = GetAttr(basePath & "\" & Name)
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo ErrHandler
        If (attrs And vbDirectory) = vbDirectory Then
            GetStepFolder = basePath & "\" & Name
            Exit Function
        End If
        Name = Dir()
    Loop
    Exit Function

ErrHandler:
    Mod_Util.WriteError "F002", "GetStepFolder: ошибка поиска папки шага. Ошибка: " & Err.Description
    Err.Clear
    GetStepFolder = ""
End Function

' Создаёт папку для шага "Шаг_<stepNumber>_<stepName>" в
' <корень>\03_Результаты генерации\01_По шагам и возвращает полный путь.
' Если папка уже существует - возвращает её путь без пересоздания.
Public Function CreateStepFolder(ByVal stepNumber As String, ByVal stepName As String) As String
    Dim root As String
    Dim basePath As String
    Dim fullPath As String

    On Error GoTo ErrHandler

    CreateStepFolder = ""

    root = Mod_Util.GetProjectRoot()
    If root = "" Then
        Mod_Util.WriteError "F001", "CreateStepFolder: папка проекта не определена (книга .xlsm не сохранена?)"
        MsgBox "Ошибка F001: не удалось определить папку проекта." & vbCrLf & _
               "Сохраните файл .xlsm и повторите.", _
               vbCritical, "Папки проекта"
        Exit Function
    End If

    basePath = root & "\" & FOLDER_STEPS
    fullPath = basePath & "\Шаг_" & stepNumber & "_" & stepName

    If Not FolderExists(fullPath) Then
        MkDir fullPath              ' при ошибке уйдём в ErrHandler (F002)
    End If

    Mod_Util.WriteLog "CreateStepFolder: " & fullPath
    CreateStepFolder = fullPath
    Exit Function

ErrHandler:
    Mod_Util.WriteError "F002", "CreateStepFolder: ошибка создания папки """ & fullPath & _
                                """. Ошибка: " & Err.Description
    MsgBox "Ошибка F002: не удалось создать папку шага." & vbCrLf & Err.Description, _
           vbCritical, "Папки проекта"
    Err.Clear
    CreateStepFolder = ""
End Function


