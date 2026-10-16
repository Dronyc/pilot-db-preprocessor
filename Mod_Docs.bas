Attribute VB_Name = "Mod_Docs"
Option Explicit

' ============================================================================
' Модуль документации: обновление README и генерация руководства пользователя.
' Реализация W4 (разделы 6 и 10 спецификации Spec_Dispatcher.md).
' Коды ошибок: D001 — нет листа 'Readme', D002 — ошибка создания папки docs
' или записи файла Readme.md.
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

' Создаёт цепочку папок по полному пути (аналог mkdir -p), если их ещё нет.
' При отсутствии прав на создание папки MkDir выдаёт ошибку, которую
' перехватывает вызывающий код (On Error GoTo ErrHandler).
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

' Сохраняет текст в файл встроенными средствами VBA
' (Open For Output, Print, Close). Существующий файл перезаписывается.
Private Sub SaveTextToFile(ByVal filePath As String, ByVal content As String)
    Dim fnum As Integer
    Dim lines() As String
    Dim i As Long

    content = Replace(content, vbCrLf, vbLf)
    lines = Split(content, vbLf)

    fnum = FreeFile
    Open filePath For Output As #fnum
    For i = LBound(lines) To UBound(lines)
        Print #fnum, lines(i)
    Next i
    Close #fnum
End Sub

' ---------------------------------------------------------------------------
' ПУБЛИЧЕСКИЕ ПРОЦЕДУРЫ И ФУНКЦИИ
' ---------------------------------------------------------------------------

' Обновляет файл Readme.md (<корень>\docs\Readme.md): собирает текст из
' колонки A листа «Readme» книги Препроцессора; если лист пуст — использует
' BuildUserGuide (раздел 10 спецификации).
Public Sub UpdateReadme()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim root As String
    Dim docsPath As String
    Dim filePath As String
    Dim lastRow As Long
    Dim r As Long
    Dim textFromSheet As String
    Dim lineVal As String

    On Error GoTo ErrHandler

    ' 1. Книга Препроцессора: ThisWorkbook либо ActiveWorkbook
    Set wb = ThisWorkbook
    If wb.Worksheets.Count = 0 Then Set wb = ActiveWorkbook

    ' 2. Лист «Readme». Если нет — D001 (MsgBox, Exit Sub).
    On Error Resume Next
    Set ws = wb.Worksheets("Readme")
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Mod_Util.WriteError "D001", "UpdateReadme: лист 'Readme' не найден."
        MsgBox "Ошибка D001: лист 'Readme' не найден.", vbCritical, "Документация"
        Exit Sub
    End If

    ' 3. Сбор текста из колонки A построчно с удалением пустых строк в конце.
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    textFromSheet = ""
    If lastRow >= 1 Then
        For r = 1 To lastRow
            lineVal = Trim(CStr(ws.Cells(r, 1).Value))
            textFromSheet = textFromSheet & lineVal & vbLf
        Next r
    End If
    ' пропуск пустых строк в конце
    Do While Len(textFromSheet) >= 1 And (Right(textFromSheet, 1) = vbLf _
                                       Or Right(textFromSheet, 1) = vbCr)
        textFromSheet = Left(textFromSheet, Len(textFromSheet) - 1)
    Loop

    ' 4. Если текст с листа пуст — используем руководство пользователя.
    If Trim(textFromSheet) = "" Then
        textFromSheet = BuildUserGuide()
        Mod_Util.WriteLog "UpdateReadme: лист 'Readme' пуст, использован BuildUserGuide."
    End If

    ' 5. Путь сохранения: <корень>\docs\Readme.md; при необходимости
    '    создаём папку docs (ошибка прав/записи ? D002).
    root = NormPath(Mod_Util.GetProjectRoot())
    If root = "" Then root = NormPath(ThisWorkbook.Path)
    docsPath = root & "\docs"
    filePath = docsPath & "\Readme.md"

    If Dir(docsPath, vbDirectory) = "" Then EnsureFolder docsPath

    ' 6. Сохранение файла встроенными средствами VBA.
    SaveTextToFile filePath, textFromSheet

    Mod_Util.WriteLog "UpdateReadme: файл сохранён в " & filePath
    MsgBox "Readme.md обновлён. Путь: " & filePath, vbInformation, "Документация"
    Exit Sub

ErrHandler:
    Mod_Util.WriteError "D002", "UpdateReadme: ошибка создания папки docs или записи файла. Ошибка: " & Err.Description
    MsgBox "Ошибка D002: не удалось сохранить Readme.md." & vbCrLf & Err.Description, _
           vbCritical, "Документация"
    Err.Clear
End Sub

' Формирует текст руководства пользователя по разделам из раздела 10
' спецификации Spec_Dispatcher.md. Используется, если лист «Readme» пуст.
Public Function BuildUserGuide() As String
    Dim s As String

    ' 1. Назначение Препроцессора
    s = s & "# Руководство пользователя Препроцессора" & vbLf
    s = s & vbLf
    s = s & "## Назначение Препроцессора" & vbLf
    s = s & "Препроцессор подготавливает структуру базы данных пилота к генерации:" & vbLf
    s = s & "преобразует исходную схему в плоскую таблицу, собирает сводную таблицу," & vbLf
    s = s & "реестры типов и покрытия, формирует промты для следующего шага." & vbLf
    s = s & vbLf

    ' 2. Требования к среде
    s = s & "## Требования к среде" & vbLf
    s = s & "Excel 2016 или выше (.xlsm), разрешённые макросы." & vbLf
    s = s & "Доступ на запись к корневой папке проекта и папке docs." & vbLf
    s = s & vbLf

    ' 3. Структура папок проекта
    s = s & "## Структура папок проекта" & vbLf
    s = s & "01_Исходные данные — файлы схемы и справочников;" & vbLf
    s = s & "02_Подготовка схемы — плоская таблица и результаты анализа;" & vbLf
    s = s & "03_Результаты генерации — сводная таблица, реестры, папки шагов;" & vbLf
    s = s & "04_Промты и docs — шаблоны промтов, Шаги.txt, документация; 05_Логи — журнал." & vbLf
    s = s & vbLf

    ' 4. Требования к исходным файлам
    s = s & "## Требования к исходным файлам" & vbLf
    s = s & "Файлы схемы размещаются в 01_Исходные данные без изменения имён." & vbLf
    s = s & "Кнопка 2 проверяет наличие обязательных файлов и показывает отсутствующие." & vbLf
    s = s & vbLf

    ' 5. Пошаговый сценарий работы (12 кнопок)
    s = s & "## Пошаговый сценарий работы (12 кнопок)" & vbLf
    s = s & "1. Создать папки проекта. 2. Проверить исходные данные." & vbLf
    s = s & "3. Собрать плоскую таблицу из «Схемы». 4. Повторный анализ преобразованного." & vbLf
    s = s & "5. Собрать типы и атрибуты (черновик). 6. Собрать типы и атрибуты (строгая)." & vbLf
    s = s & "7. Сделать промт следующего шага. 8. Сделать промт исправления ошибок." & vbLf
    s = s & "9. Показать, что нашёл Препроцессор. 10. Открыть папку Логи." & vbLf
    s = s & "11. Обновить Readme. 12. Восстановить кнопки и оформление." & vbLf
    s = s & vbLf

    ' 6. Расшифровка листов результата
    s = s & "## Расшифровка листов результата" & vbLf
    s = s & "«Типы» — системные имена типов; «Атрибуты» — атрибуты типов." & vbLf
    s = s & "«Ошибки» — ошибки последней сборки; _Реестр_типов и _Реестр_покрытия —" & vbLf
    s = s & "копии соответствующих файлов реестров для просмотра внутри Excel." & vbLf
    s = s & vbLf

    ' 7. Типовые ошибки и действия
    s = s & "## Типовые ошибки и действия" & vbLf
    s = s & "T### — проблемы с типами, A### — с атрибутами, R### — с реестрами," & vbLf
    s = s & "F### — файловые операции, P### — промты, D### — документация." & vbLf
    s = s & "При ошибке откройте папку логов (кнопка 10) и устраните причину по коду." & vbLf
    s = s & vbLf

    ' 8. Правило: правки только в файлы шагов
    s = s & "## Правило: правки только в файлы шагов" & vbLf
    s = s & "Правки результатов вносятся только в файлы папок шагов." & vbLf
    s = s & "Сводную таблицу и служебные листы вручную править нельзя." & vbLf
    s = s & vbLf

    ' 9. Ограничения
    s = s & "## Ограничения" & vbLf
    s = s & "Препроцессор не вызывает внешний искусственный интеллект." & vbLf
    s = s & "Не правит исходные файлы и не удаляет папки и файлы без явной команды." & vbLf
    s = s & vbLf

    ' 10. Порядок обновления промтов
    s = s & "## Порядок обновления промтов" & vbLf
    s = s & "Локально хранится только актуальная версия промта." & vbLf
    s = s & "История версий — в GitHub, папка /prompts/, каждая версия — новый файл." & vbLf

    BuildUserGuide = s
End Function
