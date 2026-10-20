Attribute VB_Name = "Mod_Assembler"
Option Explicit

' ============================================================
' ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ
' ============================================================
Private colErrors As Collection
Private colWarnings As Collection
Private colInfos As Collection
Private knownTypes As Object
Private attrDict As Object
Private attrSourceRow As Object
Private groupDict As Object
Private attrTypesDict As Object
Private uniqDict As Object
Private boolValuesDict As Object
Private baseTypesDict As Object
Private typesProblemsDict As Object
Private attrsProblemsDict As Object
Private sourceFolderPath As String
Private wsLog As Worksheet
Private isSoftMode As Boolean
Private uncoveredTypesCount As Long     ' НОВОЕ: типов без атрибутов
Private orphanCoverageCount As Long     ' НОВОЕ: покрытие без типа в Реестре

' ============================================================
' СОСТОЯНИЕ СЕССИИ (разделено по этапам, OPT-10)
' ============================================================
Private wbOut As Workbook           ' итоговая книга
Private wsTypes As Worksheet        ' лист "Типы"
Private wsAttrs As Worksheet        ' лист "Атрибуты"
Private wsErrors As Worksheet       ' лист "Ошибки"
Private wsWarnings As Worksheet     ' лист "Предупреждения"
Private wsReg As Worksheet          ' лист "Реестры"
Private wsReport As Worksheet       ' лист "Сводка"
Private rTypes As Long              ' следующая свободная строка на листе "Типы"
Private rAttrs As Long              ' следующая свободная строка на листе "Атрибуты"
Private rLog As Long                ' следующая свободная строка на листе "Лог дубликатов"
Private outputFile As String        ' путь итогового файла

' ============================================================
' ГЛАВНАЯ ПРОЦЕДУРА — тонкая обёртка над шестью этапами (OPT-10)
' softMode:=True  — промежуточный прогон (forward-ссылки и пустое "Состоит из" = предупреждения)
' softMode:=False — финальный прогон по склеенной таблице (полный контроль)
' ============================================================
Sub ValidateAndAssemble(Optional softMode As Boolean = False)
    On Error GoTo ErrHandler

    ' Логирование старта сборки: видно в логе, когда именно началась сборка
    ' и в каком режиме она шла (черновик / строгая проверка).
    Mod_Util.WriteLog "Старт операции: Сборка типов и атрибутов (режим: " & IIf(softMode, "черновик", "строгая") & ")"

    Call InitSession(softMode)
    Call LoadAllTypes
    Call LoadAllAttributes
    Call ValidateHierarchy
    Call BuildReports
    Call SaveResult

    Exit Sub
ErrHandler:
    ' Исключение этапа сборки: фиксируем и показываем пользователю
    Mod_Util.WriteError "F005", "ValidateAndAssemble: ошибка выполнения этапа: " & Err.Description
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Ошибка при сборке: " & Err.Description, vbCritical, "Валидация прервана"
    Err.Clear
End Sub

' ============================================================
' ЭТАП 1: ИНИЦИАЛИЗАЦИЯ СЕССИИ (OPT-10)
' Очистка состояния, словари, фиксированная входная папка, создание листов
' ============================================================
Private Sub InitSession(ByVal softMode As Boolean)
    On Error GoTo ErrHandler

    isSoftMode = softMode
    uncoveredTypesCount = 0
    orphanCoverageCount = 0

    ' Явная очистка словарей и коллекций перед заполнением (OPT-09)
    Call InitDictionaries

    ' Диалог выбора папки заменён фиксированной папкой проекта (OPT-10)
    sourceFolderPath = Mod_Util.GetProjectRoot & "\03_Результаты генерации\01_По шагам"
    If Right(sourceFolderPath, 1) <> "\" Then sourceFolderPath = sourceFolderPath & "\"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbOut = Workbooks.Add
    Set wsTypes = wbOut.Sheets(1)
    wsTypes.Name = "Типы"
    Set wsAttrs = wbOut.Sheets.Add(After:=wsTypes)
    wsAttrs.Name = "Атрибуты"
    Set wsErrors = wbOut.Sheets.Add(After:=wsAttrs)
    wsErrors.Name = "Ошибки"
    Set wsWarnings = wbOut.Sheets.Add(After:=wsErrors)
    wsWarnings.Name = "Предупреждения"
    Set wsLog = wbOut.Sheets.Add(After:=wsWarnings)
    wsLog.Name = "Лог дубликатов"
    Set wsReg = wbOut.Sheets.Add(After:=wsLog)      ' НОВОЕ
    wsReg.Name = "Реестры"                            ' НОВОЕ
    Set wsReport = wbOut.Sheets.Add(After:=wsReg)
    wsReport.Name = "Сводка"

    Call WriteHeaders(wsTypes, wsAttrs, wsErrors, wsWarnings, wsLog)

    rTypes = 2
    rAttrs = 2
    rLog = 2

    Exit Sub
ErrHandler:
    ' Ошибка инициализации сессии не должна прерывать книгу без разбора
    AddError "F005", "InitSession: не удалось инициализировать сессию сборки. Ошибка: " & Err.Description
    Err.Clear
End Sub

' ============================================================
' ЭТАП 2: ЗАГРУЗКА ВСЕХ ФАЙЛОВ *_types.xlsx (OPT-10)
' ============================================================
Private Sub LoadAllTypes()
    Dim fName As String
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim rSrc As Long
    Dim lastRow As Long
    Dim j As Long
    Dim sysName As String

    On Error GoTo ErrHandler

    ' ============================================================
    ' ФАЗА 1: ЗАГРУЗКА ВСЕХ ТИПОВ
    ' ============================================================
    colInfos.Add "=== ФАЗА 1: Загрузка типов ==="
    fName = Dir(sourceFolderPath & "*_types.xlsx")
    Do While fName <> ""
        colInfos.Add "Файл типов: " & fName

        On Error Resume Next
        Set wbSrc = Workbooks.Open(sourceFolderPath & fName, ReadOnly:=True, UpdateLinks:=0)
        If Err.Number <> 0 Then
            AddError "T005", "Не удалось открыть файл: " & fName & ". Ошибка: " & Err.Description
            On Error GoTo 0
            fName = Dir()
            GoTo NextTypeFile
        End If
        On Error GoTo 0

        On Error Resume Next
        Set wsSrc = wbSrc.Sheets(1)
        On Error GoTo 0

        If Not wsSrc Is Nothing Then
            lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).row
            For rSrc = 2 To lastRow
                sysName = Trim(CStr(wsSrc.Cells(rSrc, 1).value))
                If sysName = "" Then GoTo NextTypeRow

                For j = 1 To 12
                    wsTypes.Cells(rTypes, j).value = wsSrc.Cells(rSrc, j).value
                Next j

                Call ValidateTypeBasic(wsSrc, rSrc, fName, sysName)

                rTypes = rTypes + 1
NextTypeRow:
            Next rSrc
            Set wsSrc = Nothing
        Else
            AddWarning "T005", "В файле " & fName & " нет доступных листов"
        End If

        wbSrc.Close SaveChanges:=False
NextTypeFile:
        fName = Dir()
    Loop

    colInfos.Add "Всего типов загружено: " & (rTypes - 2)
    Exit Sub
ErrHandler:
    AddError "T005", "LoadAllTypes: ошибка загрузки файлов типов. Ошибка: " & Err.Description
    Err.Clear
End Sub

' ============================================================
' ЭТАП 3: ЗАГРУЗКА ВСЕХ ФАЙЛОВ *_attributes.xlsx (OPT-10)
' ============================================================
Private Sub LoadAllAttributes()
    Dim fName As String
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim rSrc As Long
    Dim lastRow As Long
    Dim sysName As String
    Dim attrName As String

    On Error GoTo ErrHandler

    ' ============================================================
    ' ФАЗА 2: ЗАГРУЗКА ВСЕХ АТРИБУТОВ
    ' ============================================================
    colInfos.Add "=== ФАЗА 2: Загрузка атрибутов ==="
    fName = Dir(sourceFolderPath & "*_attributes.xlsx")
    Do While fName <> ""
        colInfos.Add "Файл атрибутов: " & fName

        On Error Resume Next
        Set wbSrc = Workbooks.Open(sourceFolderPath & fName, ReadOnly:=True, UpdateLinks:=0)
        If Err.Number <> 0 Then
            AddError "A017", "Не удалось открыть файл: " & fName & ". Ошибка: " & Err.Description
            On Error GoTo 0
            fName = Dir()
            GoTo NextAttrFile
        End If
        On Error GoTo 0

        On Error Resume Next
        Set wsSrc = wbSrc.Sheets(1)
        On Error GoTo 0

        If Not wsSrc Is Nothing Then
            lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).row
            For rSrc = 2 To lastRow
                sysName = Trim(CStr(wsSrc.Cells(rSrc, 1).value))
                attrName = Trim(CStr(wsSrc.Cells(rSrc, 2).value))
                If sysName = "" And attrName = "" Then GoTo NextAttrRow

                Call CopyAttributeRow(wsSrc, rSrc, fName, sysName, attrName, wsAttrs, rAttrs, rLog)

NextAttrRow:
            Next rSrc
            Set wsSrc = Nothing
        Else
            AddWarning "A017", "В файле " & fName & " нет доступных листов"
        End If

        wbSrc.Close SaveChanges:=False
NextAttrFile:
        fName = Dir()
    Loop

    colInfos.Add "Всего атрибутов загружено (после дедупликации): " & (rAttrs - 2)
    Exit Sub
ErrHandler:
    AddError "A017", "LoadAllAttributes: ошибка загрузки файлов атрибутов. Ошибка: " & Err.Description
    Err.Clear
End Sub

' ============================================================
' ЭТАП 4: ИЕРАРХИЯ И ПРОВЕРКА ССЫЛОК (OPT-10)
' ============================================================
Private Sub ValidateHierarchy()
    On Error GoTo ErrHandler

    ' ============================================================
    ' ФАЗА 3: АВТОИСПРАВЛЕНИЕ ИЕРАРХИИ И ПРОВЕРКА ССЫЛОК
    ' ============================================================
    colInfos.Add "=== ФАЗА 3: Автоисправление иерархии ==="
    Call AutoFixHierarchy(wsTypes, rTypes)

    If Not isSoftMode Then
        Call SyncChildrenFromParents(wsTypes, rTypes)
    End If

    colInfos.Add "=== ФАЗА 4: Проверка ссылок ==="
    Call ValidateAllTypeReferences(wsTypes, rTypes)
    Call ValidateAllAttributeReferences(wsAttrs, rAttrs)
    Call FinalValidation(wsTypes, wsAttrs, rTypes)
    Exit Sub
ErrHandler:
    AddError "T018", "ValidateHierarchy: ошибка этапа проверки иерархии и ссылок. Ошибка: " & Err.Description
    Err.Clear
End Sub

' ============================================================
' ЭТАП 5: ОТЧЁТЫ (OPT-10, OPT-08)
' Реестры (BuildRegistrySheet) из сборки убраны — их формирует Mod_Registry (W6)
' ============================================================
Private Sub BuildReports()
    On Error GoTo ErrHandler

    ' ============================================================
    ' ФАЗА 5: ФОРМИРОВАНИЕ ОТЧЁТА
    ' ============================================================
    colInfos.Add "=== ФАЗА 5: Формирование отчёта ==="
    Call BuildErrorsReport(wsErrors)
    Call BuildWarningsReport(wsWarnings)
    Call HighlightProblemsInFinalFile(wsTypes, wsAttrs, rTypes, rAttrs)
    ' OPT-08: вызов BuildRegistrySheet из сборки убран — реестры формирует Mod_Registry (W6)
    Call FillSummarySheet(wsReport, rTypes, rAttrs)

    ' Реестры (W4/OPT-08): строим словари, пишем файлы, обновляем служебные листы
    Call Mod_Registry.BuildRegistries
    Call Mod_Registry.WriteRegistryFiles
    Call Mod_Registry.UpdateRegistrySheets

    wsTypes.UsedRange.Columns.AutoFit
    wsAttrs.UsedRange.Columns.AutoFit
    wsErrors.UsedRange.Columns.AutoFit
    wsWarnings.UsedRange.Columns.AutoFit
    wsLog.UsedRange.Columns.AutoFit
    wsReg.UsedRange.Columns.AutoFit     ' НОВОЕ
    wsReport.UsedRange.Columns.AutoFit

    wsTypes.Rows(1).Font.Bold = True
    wsAttrs.Rows(1).Font.Bold = True
    wsReport.Rows(1).Font.Bold = True
    wsLog.Rows(1).Font.Bold = True
    Exit Sub
ErrHandler:
    AddError "R002", "BuildReports: ошибка формирования отчётов. Ошибка: " & Err.Description
    Err.Clear
End Sub

' ============================================================
' ЭТАП 6: СОХРАНЕНИЕ ИТОГОВОГО ФАЙЛА И ИТОГОВОЕ СООБЩЕНИЕ (OPT-10)
' ============================================================
Private Sub SaveResult()
    On Error GoTo ErrHandler

    outputFile = Mod_Util.GetProjectRoot & "\03_Результаты генерации\02_Сводные\zhc_okks_gost_FINAL.xlsx"
    On Error Resume Next
    Kill outputFile
    On Error GoTo 0

    On Error Resume Next
    wbOut.SaveAs fileName:=outputFile, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        AddError "F004", "Не удалось сохранить файл: " & outputFile & ". Ошибка: " & Err.Description
        Err.Clear
    Else
        colInfos.Add "Файл сохранён: " & outputFile
    End If
    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim msg As String
    msg = "=== РЕЗУЛЬТАТЫ ВАЛИДАЦИИ ===" & vbCrLf
    msg = msg & "Режим: " & IIf(isSoftMode, "ПРОМЕЖУТОЧНЫЙ (soft)", "ФИНАЛЬНЫЙ") & vbCrLf & vbCrLf
    msg = msg & "Типов: " & (rTypes - 2) & vbCrLf
    msg = msg & "Атрибутов (после удаления дубликатов): " & (rAttrs - 2) & vbCrLf
    msg = msg & "Уникальных пар: " & attrDict.Count & vbCrLf
    msg = msg & "Типов без покрытия атрибутами: " & uncoveredTypesCount & vbCrLf   ' НОВОЕ
    msg = msg & "ОШИБОК: " & colErrors.Count & vbCrLf
    msg = msg & "ПРЕДУПРЕЖДЕНИЙ: " & colWarnings.Count & vbCrLf & vbCrLf
    msg = msg & "См. листы:" & vbCrLf
    msg = msg & "  - 'Ошибки' — критические проблемы (красный, с гиперссылками)" & vbCrLf
    msg = msg & "  - 'Предупреждения' — рекомендации (жёлтый, с гиперссылками)" & vbCrLf
    msg = msg & "  - 'Лог дубликатов' — история удаления дубликатов" & vbCrLf
    msg = msg & "  - 'Реестры' — Реестр типов и Реестр покрытия атрибутами" & vbCrLf   ' НОВОЕ
    msg = msg & "  - 'Сводка' — общая статистика" & vbCrLf & vbCrLf
    msg = msg & "Итоговый файл сохранён: " & outputFile & vbCrLf
    msg = msg & "В листах 'Типы' и 'Атрибуты' проблемные строки раскрашены."

    If colErrors.Count > 0 Then
        msg = msg & vbCrLf & vbCrLf & "[FAIL] Исправьте " & colErrors.Count & " ошибок перед импортом в БД."
        MsgBox msg, vbExclamation, "Валидация завершена"
    Else
        msg = msg & vbCrLf & vbCrLf & "[PASS] Критических ошибок не обнаружено. Файл готов к импорту."
        MsgBox msg, vbInformation, "Валидация завершена"
    End If
    Exit Sub
ErrHandler:
    AddError "F004", "SaveResult: ошибка сохранения итогового файла. Ошибка: " & Err.Description
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Err.Clear
End Sub

' ============================================================
' ИНИЦИАЛИЗАЦИЯ СЛОВАРЕЙ
' OPT-09: явная очистка всех словарей и коллекций перед
' повторным заполнением — гарантирован корректный повторный запуск
' без перезагрузки книги.
' ============================================================
Private Sub InitDictionaries()
    ' Очистка коллекций (пересоздание)
    Set colErrors = New Collection
    Set colWarnings = New Collection
    Set colInfos = New Collection

    ' Очистка словарей: RemoveAll, если словарь уже создан,
    ' иначе — создание нового экземпляра
    If Not knownTypes Is Nothing Then
        knownTypes.RemoveAll
    Else
        Set knownTypes = CreateObject("Scripting.Dictionary")
    End If
    If Not attrDict Is Nothing Then
        attrDict.RemoveAll
    Else
        Set attrDict = CreateObject("Scripting.Dictionary")
    End If
    If Not attrSourceRow Is Nothing Then
        attrSourceRow.RemoveAll
    Else
        Set attrSourceRow = CreateObject("Scripting.Dictionary")
    End If
    If Not groupDict Is Nothing Then
        groupDict.RemoveAll
    Else
        Set groupDict = CreateObject("Scripting.Dictionary")
    End If
    If Not attrTypesDict Is Nothing Then
        attrTypesDict.RemoveAll
    Else
        Set attrTypesDict = CreateObject("Scripting.Dictionary")
    End If
    If Not uniqDict Is Nothing Then
        uniqDict.RemoveAll
    Else
        Set uniqDict = CreateObject("Scripting.Dictionary")
    End If
    If Not boolValuesDict Is Nothing Then
        boolValuesDict.RemoveAll
    Else
        Set boolValuesDict = CreateObject("Scripting.Dictionary")
    End If
    If Not baseTypesDict Is Nothing Then
        baseTypesDict.RemoveAll
    Else
        Set baseTypesDict = CreateObject("Scripting.Dictionary")
    End If
    If Not typesProblemsDict Is Nothing Then
        typesProblemsDict.RemoveAll
    Else
        Set typesProblemsDict = CreateObject("Scripting.Dictionary")
    End If
    If Not attrsProblemsDict Is Nothing Then
        attrsProblemsDict.RemoveAll
    Else
        Set attrsProblemsDict = CreateObject("Scripting.Dictionary")
    End If

    AddBaseType "Root_object_type"
    AddBaseType "Smart_folder_type"
    AddBaseType "Project_folder"
    AddBaseType "File"
    AddBaseType "projectfolder"
    AddBaseType "project"
    AddBaseType "building"
    AddBaseType "section"
    AddBaseType "subsection"
    AddBaseType "main_set"
    AddBaseType "document"
    AddBaseType "ecm_doc"
    AddBaseType "referencebook_folder"
    AddBaseType "reference_contractors"
    AddBaseType "contractor"
    AddBaseType "reference_sections"
    AddBaseType "item_reference_sections"
    AddBaseType "reference_subsection"
    AddBaseType "item_reference_subsection"
    AddBaseType "reference_sets"
    AddBaseType "item_reference_sets"
    AddBaseType "Task_Folder"
    AddBaseType "Chat_type"
    AddBaseType "Chat_message_type"
    AddBaseType "Extension"
    AddBaseType "Extension_folder"
    AddBaseType "Report_6088AF81_061E_456E_9225_CF65B7B25368"
    AddBaseType "Report_Folder_F2CC6F1D_70E1_4E9B_B32F_BEB3E991318F"
    AddBaseType "Shortcut_E67517F1_93F5_4756_B651_133B816D43C8"
    AddBaseType "Task_template_folder_A0A09765_E6FB_4272_87EE_37793283DBC5"
    AddBaseType "Task_template_05339782_CCBA_4786_80C3_0F6C7E0EF3C5"
    AddBaseType "Document_template_folder_793D0CE8_65E6_484E_AAF9_7E095AF9DBD2"
    AddBaseType "Document_template_89B9E233_A6F9_4B9C_B970_55B3B3A77CED"
    AddBaseType "bim_coordinationModel"
    AddBaseType "root_tasks"
    AddBaseType "bim_modelPart"
    AddBaseType "workflow_common"
    AddBaseType "stage_common"
    AddBaseType "task_approval"
    AddBaseType "photo360Album"
    AddBaseType "bim_viewPointsFolder"
    AddBaseType "photo360_remark"
    AddBaseType "bim_viewPoint"
    AddBaseType "photo360_issue"
    AddBaseType "bim_relationProxyFolder"
    AddBaseType "bim_relationProxy"
    AddBaseType "bim_settings"
    AddBaseType "bim_modelCheckerJournalsFolder"
    AddBaseType "bim_clashDetectionJournal"
    AddBaseType "bim_modelPartSolids"
    AddBaseType "WebExtension"
    AddBaseType "$PilotServerExtension"
    AddBaseType "doc_remarksFolder"
    AddBaseType "doc_remark"
    AddBaseType "doc_issue"
    AddBaseType "externalFile_"
    AddBaseType "bim_modelRemark"
    AddBaseType "bim_searchSetsFolder"
    AddBaseType "bim_searchSet"
    AddBaseType "bim_informationModelCheckFolder"
    AddBaseType "bim_informationModelCheck"
    AddBaseType "bim_analyticSchemaFolder"
    AddBaseType "bim_analyticSchema"
    AddBaseType "bim_modelPartFolder"
    AddBaseType "recycle_bin_folder"
    AddBaseType "calendar_work_schedule_9E24"
    AddBaseType "calendar_exception_holidays_F018"
    AddBaseType "calendar_exception_sick_days_CB7F"
    AddBaseType "calendar_exception_vacation_FF6B"
    AddBaseType "calendar_exception_work_days_9AD0"
    AddBaseType "calendar_event_folder_2501"
    AddBaseType "task_simple"
    AddBaseType "task_toReviewDocument"
    AddBaseType "task_toReadDocument"
    AddBaseType "workflow_approval"
    AddBaseType "stage_approval"
    AddBaseType "bim_modelPart_other3Dmodel"
    AddBaseType "bim_modelPart_pointCloud"
    AddBaseType "object"
    AddBaseType "procurement_folder"
    AddBaseType "applications_folder"
    AddBaseType "qualification_folder"
    
    AddAttrType "Строка"
    AddAttrType "Дата и время"
    AddAttrType "Целое число"
    AddAttrType "Вещественное число"
    AddAttrType "Десятичный (денежная единица)"
    AddAttrType "Нумератор"
    AddAttrType "Состояние"
    AddAttrType "Организационная единица"
    AddAttrType "Справочник элементов"
    AddAttrType "Наследуемый"
    AddAttrType "Булевый"
    
    AddUniqueness "Не проверять"
    AddUniqueness "Показывать предупреждение"
    AddUniqueness "Запрещать создание дубликатов"
    
    boolValuesDict.Add "да", True
    boolValuesDict.Add "нет", True
End Sub

Private Sub AddBaseType(typeName As String)
    baseTypesDict.Add typeName, True
    knownTypes.Add typeName, True
End Sub

Private Sub AddAttrType(typeName As String)
    attrTypesDict.Add typeName, True
End Sub

Private Sub AddUniqueness(value As String)
    uniqDict.Add value, True
End Sub

' ============================================================
' ЗАПИСЬ ЗАГОЛОВКОВ
' ============================================================
Private Sub WriteHeaders(wsTypes As Worksheet, wsAttrs As Worksheet, _
                         wsErrors As Worksheet, wsWarnings As Worksheet, _
                         wsLog As Worksheet)
    Dim typeHeaders As Variant
    typeHeaders = Array("Системное имя типа", "Видимое имя типа", "Может быть корневым", _
        "Может монтироваться на диск", "Отображается на диске", "Является проектом", _
        "Является документом", "Может содержать умные папки", "Служебный", _
        "Дополнительные параметры", "Состоит из", "Входит в")
    Dim i As Long
    For i = LBound(typeHeaders) To UBound(typeHeaders)
        wsTypes.Cells(1, i + 1).value = typeHeaders(i)
    Next i
    
    Dim attrHeaders As Variant
    attrHeaders = Array("Системное имя типа", "name", "Наименование", "Группа атрибутов", _
        "Тип атрибута", "Отображать атрибут в названиях", "Не индексировать", _
        "Обязательный для заполнения", "Объединять с предыдущим", "Служебный", _
        "Уникальность", "Дополнительные параметры")
    For i = LBound(attrHeaders) To UBound(attrHeaders)
        wsAttrs.Cells(1, i + 1).value = attrHeaders(i)
    Next i
    
    wsErrors.Cells(1, 1).value = "Код"
    wsErrors.Cells(1, 2).value = "Источник"
    wsErrors.Cells(1, 3).value = "Сообщение"
    wsErrors.Rows(1).Font.Bold = True
    wsErrors.Rows(1).Interior.Color = RGB(255, 199, 206)
    
    wsWarnings.Cells(1, 1).value = "Код"
    wsWarnings.Cells(1, 2).value = "Источник"
    wsWarnings.Cells(1, 3).value = "Сообщение"
    wsWarnings.Rows(1).Font.Bold = True
    wsWarnings.Rows(1).Interior.Color = RGB(255, 235, 156)
    
    wsLog.Cells(1, 1).value = "Тип"
    wsLog.Cells(1, 2).value = "Атрибут"
    wsLog.Cells(1, 3).value = "Действие"
    wsLog.Cells(1, 4).value = "Причина"
    wsLog.Cells(1, 5).value = "Источник (файл)"
    wsLog.Rows(1).Font.Bold = True
End Sub

' ============================================================
' ДОБАВЛЕНИЕ ОШИБКИ / ПРЕДУПРЕЖДЕНИЯ
' ============================================================
Private Sub AddError(code As String, message As String, _
                     Optional fileName As String = "", _
                     Optional sheetName As String = "", _
                     Optional rowNumber As Long = 0, _
                     Optional objKey As String = "", _
                     Optional objKind As String = "")
    Dim entry As String
    entry = code & "||" & message & "||" & fileName & "||" & sheetName & "||" & rowNumber
    colErrors.Add entry
    
    If objKey <> "" Then
        If objKind = "TYPE" Then
            If Not typesProblemsDict.Exists(objKey) Then typesProblemsDict.Add objKey, New Collection
            typesProblemsDict(objKey).Add code & "|E"
        ElseIf objKind = "ATTR" Then
            If Not attrsProblemsDict.Exists(objKey) Then attrsProblemsDict.Add objKey, New Collection
            attrsProblemsDict(objKey).Add code & "|E"
        End If
    End If
End Sub

Private Sub AddWarning(code As String, message As String, _
                       Optional fileName As String = "", _
                       Optional sheetName As String = "", _
                       Optional rowNumber As Long = 0, _
                       Optional objKey As String = "", _
                       Optional objKind As String = "")
    Dim entry As String
    entry = code & "||" & message & "||" & fileName & "||" & sheetName & "||" & rowNumber
    colWarnings.Add entry
    
    If objKey <> "" Then
        If objKind = "TYPE" Then
            If Not typesProblemsDict.Exists(objKey) Then typesProblemsDict.Add objKey, New Collection
            typesProblemsDict(objKey).Add code & "|W"
        ElseIf objKind = "ATTR" Then
            If Not attrsProblemsDict.Exists(objKey) Then attrsProblemsDict.Add objKey, New Collection
            attrsProblemsDict(objKey).Add code & "|W"
        End If
    End If
End Sub

' ============================================================
' БАЗОВАЯ ВАЛИДАЦИЯ ТИПА
' ============================================================
Private Sub ValidateTypeBasic(ws As Worksheet, row As Long, fileName As String, sysName As String)
    Dim visibleName As String
    Dim isDocument As String
    Dim smartFolders As String
    Dim displayOnDisk As String
    Dim val As String
    Dim bc As Variant
    Dim lowerName As String
    Dim children As String
    Dim parents As String
    Dim colName As String
    Dim colIdx As Long
    
    visibleName = Trim(CStr(ws.Cells(row, 2).value))
    displayOnDisk = LCase(Trim(CStr(ws.Cells(row, 5).value)))
    isDocument = LCase(Trim(CStr(ws.Cells(row, 7).value)))
    smartFolders = LCase(Trim(CStr(ws.Cells(row, 8).value)))
    children = Trim(CStr(ws.Cells(row, 11).value))
    parents = Trim(CStr(ws.Cells(row, 12).value))
    
    ' T016: Видимое имя
    If visibleName = "" Then
        AddError "T016", "Тип '" & sysName & "': поле 'Видимое имя типа' пустое. Заполните наименование на русском языке.", _
                 fileName, "Типы", row, sysName, "TYPE"
    End If
    
    ' T016/T006: Булевы поля
    Dim boolCols As Variant
    boolCols = Array(Array(3, "Может быть корневым"), _
                     Array(4, "Может монтироваться на диск"), _
                     Array(5, "Отображается на диске"), _
                     Array(6, "Является проектом"), _
                     Array(7, "Является документом"), _
                     Array(8, "Может содержать умные папки"), _
                     Array(9, "Служебный"))
    
    For Each bc In boolCols
        colIdx = bc(0)
        colName = bc(1)
        val = LCase(Trim(CStr(ws.Cells(row, colIdx).value)))
        
        If val = "" Then
            AddError "T016", "Тип '" & sysName & "': поле '" & colName & "' пустое. Допустимые значения: 'да' или 'нет'.", _
                     fileName, "Типы", row, sysName, "TYPE"
        ElseIf Not boolValuesDict.Exists(val) Then
            AddError "T006", "Тип '" & sysName & "': колонка '" & colName & "' содержит значение '" & val & "'. Допустимо только 'да' или 'нет'.", _
                     fileName, "Типы", row, sysName, "TYPE"
        End If
    Next bc
    
    ' T016: Состоит из (в soft-режиме — предупреждение)
    If children = "" Then
        If isSoftMode Then
            AddWarning "T016", "Тип '" & sysName & "': поле 'Состоит из' пустое. Дети-документы могут быть добавлены на следующих шагах.", _
                     fileName, "Типы", row, sysName, "TYPE"
        Else
            AddError "T016", "Тип '" & sysName & "': поле 'Состоит из' пустое. Укажите дочерние типы или оставьте пустым только для типов без дочерних элементов.", _
                     fileName, "Типы", row, sysName, "TYPE"
        End If
    End If
    
    ' T016: Входит в
    If parents = "" Then
        AddError "T016", "Тип '" & sysName & "': поле 'Входит в' пустое. Укажите родительский тип.", _
                 fileName, "Типы", row, sysName, "TYPE"
    End If
    
    ' T002: имя типа — camelCase
    If Not IsCamelTypeName(sysName) Then
        AddError "T002", "Системное имя типа '" & sysName & "' нарушает правило camelCase. Допустимы только латинские буквы без подчёркиваний; первое слово со строчной; исключения: суффикс _doc, фазы phase0–phase3, коды форм ksN.", _
                 fileName, "Типы", row, sysName, "TYPE"
    End If
    
    ' T017: цифры только в phase0–phase3 и ksN
    If ContainsForbiddenDigit(sysName) Then
        AddError "T017", "Тип '" & sysName & "': цифры в системном имени допустимы только у фаз phase0–phase3 и кодов форм ksN.", _
                 fileName, "Типы", row, sysName, "TYPE"
    End If
    
    ' T003: Запрещённые префиксы
    lowerName = LCase(sysName)
    If Left(lowerName, 3) = "lc_" Or Left(lowerName, 8) = "lc_gost_" Then
        AddError "T003", "Системное имя '" & sysName & "' использует запрещённый префикс lc_/lc_gost_. Удалите префикс.", _
                 fileName, "Типы", row, sysName, "TYPE"
    End If
    If InStr(lowerName, "gost") > 0 And Left(lowerName, 5) <> "gost_" Then
        AddWarning "T003", "Имя '" & sysName & "' содержит 'gost'. Проверьте соответствие Схеме.", _
                   fileName, "Типы", row, sysName, "TYPE"
    End If
    
    ' T004: Дубликат
    If knownTypes.Exists(sysName) And Not baseTypesDict.Exists(sysName) Then
        AddError "T004", "Дубликат системного имени типа '" & sysName & "'. Переименуйте тип или удалите дублирующую строку.", _
                 fileName, "Типы", row, sysName, "TYPE"
    Else
        If Not knownTypes.Exists(sysName) Then knownTypes.Add sysName, True
    End If
    
    ' T007: Отображается на диске (OPT-07: проверки нет в спецификации — управляется конфигурацией)
    If Mod_Util.GetConfigValue("EnableDiskDisplayCheck") = "да" Then
        If displayOnDisk = "да" Then
            If sysName <> "bim_coordinationModel" And sysName <> "bim_modelPartFolder" Then
                AddError "T007", "Тип '" & sysName & "': 'Отображается на диске'=да недопустимо. Разрешено только для bim_coordinationModel и bim_modelPartFolder.", _
                         fileName, "Типы", row, sysName, "TYPE"
            End If
        End If
    End If

    ' T008: Умные папки (OPT-07: проверки нет в спецификации — управляется конфигурацией)
    If Mod_Util.GetConfigValue("EnableSmartFolderCheck") = "да" Then
        If smartFolders = "да" Then
            Dim allowedSmartFolderTypes As Variant
            allowedSmartFolderTypes = Array("project", "object", "contractor", _
                "procurement_folder", "applications_folder", "qualification_folder", _
                "reference_contractors", "referencebook_folder")
            Dim isAllowed As Boolean
            isAllowed = False
            Dim sft As Variant
            For Each sft In allowedSmartFolderTypes
                If LCase(sysName) = LCase(CStr(sft)) Then
                    isAllowed = True
                    Exit For
                End If
            Next sft
            If Not isAllowed Then
                AddWarning "T008", "Тип '" & sysName & "': 'Может содержать умные папки'=да. Разрешено только для ключевых типов (project, object, contractor, procurement_folder, applications_folder, qualification_folder).", _
                           fileName, "Типы", row, sysName, "TYPE"
            End If
        End If
    End If
    ' T009: Документ без doc_remarksFolder
    If isDocument = "да" Then
        If InStr(children, "doc_remarksFolder") = 0 Then
            AddError "T009", "Тип '" & sysName & "' является документом, но в 'Состоит из' отсутствует 'doc_remarksFolder'. Добавьте его.", _
                     fileName, "Типы", row, sysName, "TYPE"
        End If
    End If
    
    ' T015: object -> projectfolder
    If LCase(sysName) = "object" Then
        If InStr(LCase(parents), "projectfolder") = 0 Then
            AddError "T015", "Тип 'object' должен иметь родительский тип 'projectfolder' в колонке 'Входит в'.", _
                     fileName, "Типы", row, sysName, "TYPE"
        End If
    End If
End Sub

' ============================================================
' КОПИРОВАНИЕ СТРОКИ АТРИБУТА С ДЕДУПЛИКАЦИЕЙ
' ============================================================
Private Sub CopyAttributeRow(ws As Worksheet, row As Long, fileName As String, _
                             sysType As String, attrName As String, _
                             wsAttrs As Worksheet, ByRef rAttrs As Long, ByRef rLog As Long)
    Dim j As Long
    Dim attrKey As String
    attrKey = sysType & "|" & attrName
    
    If attrDict.Exists(attrKey) Then
        Dim existingTargetRow As Long
        Dim existingSourceRow As Long
        Dim scoreExisting As Long
        Dim scoreNew As Long
        
        existingTargetRow = CLng(attrDict(attrKey))
        existingSourceRow = CLng(attrSourceRow(attrKey))
        scoreExisting = CalcRowScore(ws, existingSourceRow)
        scoreNew = CalcRowScore(ws, row)
        
        If scoreNew > scoreExisting Then
            For j = 1 To 12
                wsAttrs.Cells(existingTargetRow + 1, j).value = ws.Cells(row, j).value
            Next j
            wsLog.Cells(rLog, 1).value = sysType
            wsLog.Cells(rLog, 2).value = attrName
            wsLog.Cells(rLog, 3).value = "ЗАМЕНА"
            wsLog.Cells(rLog, 4).value = "Новая строка полнее (оценка: " & scoreNew & " vs " & scoreExisting & ")"
            wsLog.Cells(rLog, 5).value = fileName
            wsLog.Cells(rLog, 3).Interior.Color = RGB(198, 239, 206)
            attrSourceRow(attrKey) = row
            rLog = rLog + 1
        ElseIf scoreNew = scoreExisting Then
            wsLog.Cells(rLog, 1).value = sysType
            wsLog.Cells(rLog, 2).value = attrName
            wsLog.Cells(rLog, 3).value = "ПРОПУСК"
            wsLog.Cells(rLog, 4).value = "Идентичные данные, оставлена первая"
            wsLog.Cells(rLog, 5).value = fileName
            wsLog.Cells(rLog, 3).Interior.Color = RGB(255, 235, 156)
            rLog = rLog + 1
        Else
            wsLog.Cells(rLog, 1).value = sysType
            wsLog.Cells(rLog, 2).value = attrName
            wsLog.Cells(rLog, 3).value = "ПРОПУСК"
            wsLog.Cells(rLog, 4).value = "Существующая строка полнее (оценка: " & scoreExisting & " vs " & scoreNew & ")"
            wsLog.Cells(rLog, 5).value = fileName
            wsLog.Cells(rLog, 3).Interior.Color = RGB(255, 235, 156)
            rLog = rLog + 1
        End If
    Else
        attrDict.Add attrKey, rAttrs - 1
        attrSourceRow.Add attrKey, row
        For j = 1 To 12
            wsAttrs.Cells(rAttrs, j).value = ws.Cells(row, j).value
        Next j
        rAttrs = rAttrs + 1
    End If
End Sub

' ============================================================
' АВТОИСПРАВЛЕНИЕ ИЕРАРХИИ (родитель -> ребёнок)
' ============================================================
Private Sub AutoFixHierarchy(wsTypes As Worksheet, rTypes As Long)
    ' OPT-11: управляемость автоисправления (GetConfigValue("EnableAutoFix"))
    If Mod_Util.GetConfigValue("EnableAutoFix") <> "да" Then
        colInfos.Add "AutoFixHierarchy: автоисправление отключено конфигурацией EnableAutoFix"
        Mod_Util.WriteWarning "T014", "AutoFixHierarchy: автоисправление иерархии пропущено (EnableAutoFix <> да), связи 'Входит в' не дополнялись."
        Exit Sub
    End If

    Dim r As Long
    Dim typeName As String
    Dim children As String
    Dim refs() As String
    Dim ref As Variant
    Dim refName As String
    Dim parents As String
    Dim fixedCount As Long
    fixedCount = 0
    
    Dim knownTypesLocal As Object
    Set knownTypesLocal = CreateObject("Scripting.Dictionary")
    For r = 2 To rTypes - 1
        Dim tName As String
        tName = Trim(CStr(wsTypes.Cells(r, 1).value))
        If tName <> "" Then
            If Not knownTypesLocal.Exists(tName) Then
                knownTypesLocal.Add tName, r
            End If
        End If
    Next r
    
    For r = 2 To rTypes - 1
        typeName = Trim(CStr(wsTypes.Cells(r, 1).value))
        If typeName = "" Then GoTo NextRow
        
        children = Trim(CStr(wsTypes.Cells(r, 11).value))
        If children = "" Then GoTo NextRow
        
        refs = Split(Replace(children, ";", ","), ",")
        For Each ref In refs
            refName = Trim(CStr(ref))
            If refName = "" Then GoTo NextRef
            
            If Not knownTypesLocal.Exists(refName) Then GoTo NextRef
            
            Dim childRow As Long
            childRow = CLng(knownTypesLocal(refName))
            
            parents = Trim(CStr(wsTypes.Cells(childRow, 12).value))
            If InStr(LCase(parents), LCase(typeName)) = 0 Then
                If parents = "" Then
                    wsTypes.Cells(childRow, 12).value = typeName
                    ' OPT-11: протокол каждого исправления — что было, что стало
                    Mod_Util.WriteLog "AutoFixHierarchy: тип " & refName & ", добавлен родитель " & typeName & _
                                      " (было: ''; стало: '" & typeName & "')"
                Else
                    wsTypes.Cells(childRow, 12).value = parents & "; " & typeName
                    ' OPT-11: протокол каждого исправления — что было, что стало
                    Mod_Util.WriteLog "AutoFixHierarchy: тип " & refName & ", добавлен родитель " & typeName & _
                                      " (было: '" & parents & "'; стало: '" & parents & "; " & typeName & "')"
                End If
                fixedCount = fixedCount + 1
            End If
            
NextRef:
        Next ref
NextRow:
    Next r
    
    If fixedCount > 0 Then
        colInfos.Add "Автоматически исправлено связей иерархии: " & fixedCount
    End If
End Sub

' ============================================================
' ТОЧНОЕ СРАВНЕНИЕ ИМЕНИ В СПИСКЕ
' ============================================================
Private Function HasName(ByVal listStr As String, ByVal nm As String) As Boolean
    Dim a() As String
    Dim v As Variant
    a = Split(listStr, ";")
    For Each v In a
        If Trim(CStr(v)) = nm Then
            HasName = True
            Exit Function
        End If
    Next v
    HasName = False
End Function

' ============================================================
' ОБРАТНАЯ СИНХРОНИЗАЦИЯ (ребёнок -> родитель), финальный режим
' ============================================================
Private Sub SyncChildrenFromParents(wsTypes As Worksheet, rTypes As Long)
    ' OPT-11: управляемость автоисправления (GetConfigValue("EnableAutoFix"))
    If Mod_Util.GetConfigValue("EnableAutoFix") <> "да" Then
        colInfos.Add "SyncChildrenFromParents: автоисправление отключено конфигурацией EnableAutoFix"
        Mod_Util.WriteWarning "T014", "SyncChildrenFromParents: обратная синхронизация 'Состоит из' пропущена (EnableAutoFix <> да)."
        Exit Sub
    End If

    Dim r As Long
    Dim pr As Long
    Dim childName As String
    Dim parentsStr As String
    Dim ch As String
    Dim rowByName As Object
    Set rowByName = CreateObject("Scripting.Dictionary")
    
    For r = 2 To rTypes - 1
        childName = Trim(CStr(wsTypes.Cells(r, 1).value))
        If childName <> "" And Not rowByName.Exists(childName) Then
            rowByName.Add childName, r
        End If
    Next r
    
    Dim syncCount As Long
    syncCount = 0
    
    For r = 2 To rTypes - 1
        childName = Trim(CStr(wsTypes.Cells(r, 1).value))
        parentsStr = Trim(CStr(wsTypes.Cells(r, 12).value))
        If parentsStr = "" Or childName = "" Then GoTo NextSync
        
        Dim refs() As String
        refs = Split(Replace(parentsStr, ";", ","), ",")
        Dim ref As Variant
        For Each ref In refs
            Dim pn As String
            pn = Trim(CStr(ref))
            If pn <> "" And rowByName.Exists(pn) Then
                pr = rowByName(pn)
                ch = Trim(CStr(wsTypes.Cells(pr, 11).value))
                If ch = "" Then
                    wsTypes.Cells(pr, 11).value = childName
                    ' OPT-11: протокол каждого исправления — что было, что стало
                    Mod_Util.WriteLog "SyncChildrenFromParents: тип " & pn & ", добавлен ребёнок " & childName & _
                                      " (было: ''; стало: '" & childName & "')"
                    syncCount = syncCount + 1
                ElseIf Not HasName(ch, childName) Then
                    wsTypes.Cells(pr, 11).value = ch & "; " & childName
                    ' OPT-11: протокол каждого исправления — что было, что стало
                    Mod_Util.WriteLog "SyncChildrenFromParents: тип " & pn & ", добавлен ребёнок " & childName & _
                                      " (было: '" & ch & "'; стало: '" & ch & "; " & childName & "')"
                    syncCount = syncCount + 1
                End If
            End If
        Next ref
NextSync:
    Next r
    
    If syncCount > 0 Then
        colInfos.Add "Обратная синхронизация 'Состоит из': " & syncCount
    End If
End Sub

' ============================================================
' ПРОВЕРКА ВСЕХ ССЫЛОК В ТИПАХ
' ============================================================
Private Sub ValidateAllTypeReferences(wsTypes As Worksheet, rTypes As Long)
    Dim r As Long
    Dim typeName As String
    Dim children As String
    Dim parents As String
    Dim refs() As String
    Dim ref As Variant
    Dim refName As String
    
    For r = 2 To rTypes - 1
        typeName = Trim(CStr(wsTypes.Cells(r, 1).value))
        If typeName = "" Then GoTo NextTypeRef
        
        children = Trim(CStr(wsTypes.Cells(r, 11).value))
        parents = Trim(CStr(wsTypes.Cells(r, 12).value))
        
        If children <> "" Then
            refs = Split(Replace(children, ";", ","), ",")
            For Each ref In refs
                refName = Trim(CStr(ref))
                If refName <> "" Then
                    If Not knownTypes.Exists(refName) Then
                        If isSoftMode Then
                            AddWarning "T010", "Тип '" & typeName & "': ссылка 'Состоит из' на '" & refName & "' ещё не создана (будет добавлена на следующем шаге).", _
                                     "", "Типы", r, typeName, "TYPE"
                        Else
                            AddError "T010", "Тип '" & typeName & "': в колонке 'Состоит из' указана ссылка на несуществующий тип '" & refName & "'. Создайте этот тип или исправьте имя.", _
                                     "", "Типы", r, typeName, "TYPE"
                        End If
                    End If
                End If
            Next ref
        End If
        
        If parents <> "" Then
            refs = Split(Replace(parents, ";", ","), ",")
            For Each ref In refs
                refName = Trim(CStr(ref))
                If refName <> "" Then
                    If Not knownTypes.Exists(refName) Then
                        If isSoftMode Then
                            AddWarning "T011", "Тип '" & typeName & "': ссылка 'Входит в' на '" & refName & "' ещё не создана (будет добавлена на следующем шаге).", _
                                     "", "Типы", r, typeName, "TYPE"
                        Else
                            AddError "T011", "Тип '" & typeName & "': в колонке 'Входит в' указана ссылка на несуществующий тип '" & refName & "'. Создайте этот тип или исправьте имя.", _
                                     "", "Типы", r, typeName, "TYPE"
                        End If
                    End If
                End If
            Next ref
        End If
        
        ' T012: самоссылка
        If children <> "" Then
            refs = Split(Replace(children, ";", ","), ",")
            For Each ref In refs
                If Trim(CStr(ref)) = typeName Then
                    AddError "T012", "Тип '" & typeName & "' ссылается сам на себя в 'Состоит из'. Удалите самоссылку.", _
                             "", "Типы", r, typeName, "TYPE"
                    Exit For
                End If
            Next ref
        End If
        If parents <> "" Then
            refs = Split(Replace(parents, ";", ","), ",")
            For Each ref In refs
                If Trim(CStr(ref)) = typeName Then
                    AddError "T012", "Тип '" & typeName & "' ссылается сам на себя в 'Входит в'. Удалите самоссылку.", _
                             "", "Типы", r, typeName, "TYPE"
                    Exit For
                End If
            Next ref
        End If
        
NextTypeRef:
    Next r
End Sub

' ============================================================
' ПРОВЕРКА ВСЕХ ССЫЛОК В АТРИБУТАХ
' ============================================================
Private Sub ValidateAllAttributeReferences(wsAttrs As Worksheet, rAttrs As Long)
    Dim r As Long
    Dim sysType As String
    Dim attrName As String
    Dim attrType As String
    Dim uniqueness As String
    Dim additionalParams As String
    Dim visibleName As String
    Dim groupName As String
    Dim val As String
    Dim bc As Variant
    Dim attrKey As String
    Dim colName As String
    Dim colIdx As Long
    
    For r = 2 To rAttrs - 1
        sysType = Trim(CStr(wsAttrs.Cells(r, 1).value))
        attrName = Trim(CStr(wsAttrs.Cells(r, 2).value))
        visibleName = Trim(CStr(wsAttrs.Cells(r, 3).value))
        groupName = Trim(CStr(wsAttrs.Cells(r, 4).value))
        attrType = Trim(CStr(wsAttrs.Cells(r, 5).value))
        uniqueness = Trim(CStr(wsAttrs.Cells(r, 11).value))
        additionalParams = Trim(CStr(wsAttrs.Cells(r, 12).value))
        
        If sysType = "" And attrName = "" Then GoTo NextAttrRef
        
        attrKey = sysType & "|" & attrName
        
        If sysType = "" Then
            AddError "A015", "Строка " & r & ": поле 'Системное имя типа' пустое. Укажите тип, к которому относится атрибут.", _
                     "", "Атрибуты", r, "|", "ATTR"
            GoTo NextAttrRef
        End If
        
        If attrName = "" Then
            AddError "A015", "Тип '" & sysType & "': поле 'name' пустое. Укажите системное имя атрибута в формате camelCase.", _
                     "", "Атрибуты", r, sysType & "|", "ATTR"
            GoTo NextAttrRef
        End If
        
        If visibleName = "" Then
            AddError "A015", "Атрибут '" & sysType & "." & attrName & "': поле 'Наименование' пустое. Заполните видимое имя атрибута на русском языке.", _
                     "", "Атрибуты", r, attrKey, "ATTR"
        End If
        
        If attrType = "" Then
            AddError "A015", "Атрибут '" & sysType & "." & attrName & "': поле 'Тип атрибута' пустое. Укажите тип из списка myAdmin.", _
                     "", "Атрибуты", r, attrKey, "ATTR"
        End If
        
        Dim boolAttrCols As Variant
        boolAttrCols = Array(Array(6, "Отображать атрибут в названиях"), _
                             Array(7, "Не индексировать"), _
                             Array(8, "Обязательный для заполнения"), _
                             Array(9, "Объединять с предыдущим"), _
                             Array(10, "Служебный"))
        
        For Each bc In boolAttrCols
            colIdx = bc(0)
            colName = bc(1)
            val = LCase(Trim(CStr(wsAttrs.Cells(r, colIdx).value)))
            
            If val = "" Then
                AddError "A015", "Атрибут '" & sysType & "." & attrName & "': поле '" & colName & "' пустое. Допустимые значения: 'да' или 'нет'.", _
                         "", "Атрибуты", r, attrKey, "ATTR"
            ElseIf Not boolValuesDict.Exists(val) Then
                AddError "A008", "Атрибут '" & sysType & "." & attrName & "': колонка '" & colName & "' содержит значение '" & val & "'. Допустимо только 'да' или 'нет'.", _
                         "", "Атрибуты", r, attrKey, "ATTR"
            End If
        Next bc
        
        If uniqueness = "" Then
            AddError "A015", "Атрибут '" & sysType & "." & attrName & "': поле 'Уникальность' пустое. Укажите значение из списка myAdmin.", _
                     "", "Атрибуты", r, attrKey, "ATTR"
        End If
        
        ' A001: неизвестный тип
        If sysType <> "" And Not knownTypes.Exists(sysType) Then
            AddError "A001", "Атрибут '" & attrName & "' ссылается на несуществующий тип '" & sysType & "'. Создайте этот тип в листе 'Типы' или исправьте имя.", _
                     "", "Атрибуты", r, attrKey, "ATTR"
        End If
        
        ' A003: имя атрибута — camelCase
        If Not IsCamelAttrName(attrName) Then
            AddError "A003", "Атрибут '" & sysType & "." & attrName & "': имя атрибута должно быть camelCase: только латинские буквы без подчёркиваний и пробелов; первое слово со строчной.", _
                     "", "Атрибуты", r, attrKey, "ATTR"
        End If
        
        ' A016: цифры в имени атрибута запрещены
        If ContainsDigit(attrName) Then
            AddError "A016", "Атрибут '" & sysType & "." & attrName & "' содержит цифры в имени. Уберите номера (v1_0, 1_1 и т.п.).", _
                     "", "Атрибуты", r, attrKey, "ATTR"
        End If
        
        ' A006: тип атрибута
        If attrType <> "" And Not attrTypesDict.Exists(attrType) Then
            AddError "A006", "Атрибут '" & sysType & "." & attrName & "': недопустимый тип атрибута '" & attrType & "'. Используйте значения из myAdmin.", _
                     "", "Атрибуты", r, attrKey, "ATTR"
        End If
        
        ' A007: уникальность
        If uniqueness <> "" And Not uniqDict.Exists(uniqueness) Then
            AddError "A007", "Атрибут '" & sysType & "." & attrName & "': недопустимое значение уникальности '" & uniqueness & "'.", _
                     "", "Атрибуты", r, attrKey, "ATTR"
        End If
        
        ' A011: наследуемый без источника
        If attrType = "Наследуемый" Then
            If InStr(additionalParams, "Наследуется от:") = 0 Then
                AddError "A011", "Атрибут '" & sysType & "." & attrName & "' имеет тип 'Наследуемый', но не указан источник. Формат: 'Наследуется от: parent_type.attribute_name'.", _
                         "", "Атрибуты", r, attrKey, "ATTR"
            End If
        End If
        
        ' A013: состояние без UserStateRootNode
        If attrType = "Состояние" Then
            If InStr(additionalParams, "UserStateRootNode") = 0 And _
               InStr(additionalParams, "Выбрать машину состояний") = 0 Then
                AddWarning "A013", "Атрибут '" & sysType & "." & attrName & "' имеет тип 'Состояние', но XML не содержит 'UserStateRootNode'.", _
                           "", "Атрибуты", r, attrKey, "ATTR"
            End If
        End If
        
        ' A014: справочник без Source
        If attrType = "Справочник элементов" Then
            If InStr(additionalParams, "ElementBook") > 0 And InStr(additionalParams, "Source=") = 0 Then
                AddError "A014", "Атрибут '" & sysType & "." & attrName & "' имеет тип 'Справочник элементов', но в XML не указан 'Source'.", _
                         "", "Атрибуты", r, attrKey, "ATTR"
            End If
        End If
        
        If groupName <> "" Then
            If groupDict.Exists(groupName) Then
                groupDict(groupName) = groupDict(groupName) + 1
            Else
                groupDict.Add groupName, 1
            End If
        End If
        
NextAttrRef:
    Next r
End Sub

' ============================================================
' ФИНАЛЬНЫЕ ПРОВЕРКИ
' ============================================================
Private Sub FinalValidation(wsTypes As Worksheet, wsAttrs As Worksheet, rTypes As Long)
    Dim k As Variant
    For Each k In groupDict.Keys
        If groupDict(k) < 2 Then
            AddWarning "A009", "Группа атрибутов '" & CStr(k) & "' содержит только " & groupDict(k) & " атрибут(ов). Группа должна содержать минимум 2 атрибута."
        End If
    Next k
    
    Dim r As Long
    Dim r2 As Long
    Dim typeName As String
    Dim children As String
    Dim refs() As String
    Dim ref As Variant
    Dim refName As String
    Dim found As Boolean
    Dim refParents As String
    
    For r = 2 To rTypes - 1
        typeName = Trim(CStr(wsTypes.Cells(r, 1).value))
        children = Trim(CStr(wsTypes.Cells(r, 11).value))
        
        If children <> "" Then
            refs = Split(Replace(children, ";", ","), ",")
            For Each ref In refs
                refName = Trim(CStr(ref))
                If refName <> "" Then
                    found = False
                    For r2 = 2 To rTypes - 1
                        If Trim(CStr(wsTypes.Cells(r2, 1).value)) = refName Then
                            refParents = Trim(CStr(wsTypes.Cells(r2, 12).value))
                            If InStr(refParents, typeName) = 0 Then
                                AddWarning "T013", "Тип '" & refName & "' указан в 'Состоит из' типа '" & typeName & "', но не указывает '" & typeName & "' в своём 'Входит в'."
                            End If
                            found = True
                            Exit For
                        End If
                    Next r2
                End If
            Next ref
        End If
    Next r
End Sub

' ============================================================
' ОТЧЁТ ОБ ОШИБКАХ
' ============================================================
Private Sub BuildErrorsReport(wsErrors As Worksheet)
    Dim rErr As Long
    Dim entry As Variant
    Dim parts() As String
    rErr = 2
    
    For Each entry In colErrors
        parts = Split(CStr(entry), "||")
        
        wsErrors.Cells(rErr, 1).value = parts(0)
        
        If UBound(parts) >= 4 And parts(2) <> "" And parts(4) <> "" Then
            Dim linkText As String
            linkText = parts(2) & " [" & parts(3) & "!A" & parts(4) & "]"
            wsErrors.Cells(rErr, 2).value = linkText
            
            Dim fullPath As String
            fullPath = sourceFolderPath & parts(2)
            On Error Resume Next
            wsErrors.Hyperlinks.Add _
                Anchor:=wsErrors.Cells(rErr, 2), _
                Address:=fullPath, _
                SubAddress:="'" & parts(3) & "'!A" & parts(4), _
                TextToDisplay:=linkText, _
                ScreenTip:="Перейти к проблемной строке"
            On Error GoTo 0
        Else
            wsErrors.Cells(rErr, 2).value = ""
        End If
        
        wsErrors.Cells(rErr, 3).value = parts(1)
        
        Dim rngErr As Range
        Set rngErr = wsErrors.Range(wsErrors.Cells(rErr, 1), wsErrors.Cells(rErr, 3))
        rngErr.Interior.Color = RGB(255, 199, 206)
        rngErr.Font.Color = RGB(156, 0, 6)
        
        rErr = rErr + 1
    Next entry
End Sub

' ============================================================
' ОТЧЁТ О ПРЕДУПРЕЖДЕНИЯХ
' ============================================================
Private Sub BuildWarningsReport(wsWarnings As Worksheet)
    Dim rWrn As Long
    Dim entry As Variant
    Dim parts() As String
    rWrn = 2
    
    For Each entry In colWarnings
        parts = Split(CStr(entry), "||")
        
        wsWarnings.Cells(rWrn, 1).value = parts(0)
        
        If UBound(parts) >= 4 And parts(2) <> "" And parts(4) <> "" Then
            Dim linkTextW As String
            linkTextW = parts(2) & " [" & parts(3) & "!A" & parts(4) & "]"
            wsWarnings.Cells(rWrn, 2).value = linkTextW
            
            Dim fullPathW As String
            fullPathW = sourceFolderPath & parts(2)
            On Error Resume Next
            wsWarnings.Hyperlinks.Add _
                Anchor:=wsWarnings.Cells(rWrn, 2), _
                Address:=fullPathW, _
                SubAddress:="'" & parts(3) & "'!A" & parts(4), _
                TextToDisplay:=linkTextW, _
                ScreenTip:="Перейти к проблемной строке"
            On Error GoTo 0
        Else
            wsWarnings.Cells(rWrn, 2).value = ""
        End If
        
        wsWarnings.Cells(rWrn, 3).value = parts(1)
        
        Dim rngWrn As Range
        Set rngWrn = wsWarnings.Range(wsWarnings.Cells(rWrn, 1), wsWarnings.Cells(rWrn, 3))
        rngWrn.Interior.Color = RGB(255, 235, 156)
        rngWrn.Font.Color = RGB(156, 101, 0)
        
        rWrn = rWrn + 1
    Next entry
End Sub

' ============================================================
' РАСКРАСКА ПРОБЛЕМНЫХ СТРОК
' ============================================================
Private Sub HighlightProblemsInFinalFile(wsTypes As Worksheet, wsAttrs As Worksheet, _
                                         rTypes As Long, rAttrs As Long)
    Dim r As Long
    Dim key As String
    Dim codes As Collection
    Dim code As Variant
    Dim hasError As Boolean
    Dim hasWarning As Boolean
    Dim rng As Range
    Dim sev() As String
    
    For r = 2 To rTypes - 1
        key = Trim(CStr(wsTypes.Cells(r, 1).value))
        If typesProblemsDict.Exists(key) Then
            Set codes = typesProblemsDict(key)
            hasError = False
            hasWarning = False
            For Each code In codes
                sev = Split(CStr(code), "|")
                If UBound(sev) >= 1 Then
                    If sev(1) = "E" Then hasError = True Else hasWarning = True
                End If
            Next code
            
            Set rng = wsTypes.Range(wsTypes.Cells(r, 1), wsTypes.Cells(r, 12))
            If hasError Then
                rng.Interior.Color = RGB(255, 199, 206)
                rng.Font.Color = RGB(156, 0, 6)
            ElseIf hasWarning Then
                rng.Interior.Color = RGB(255, 235, 156)
                rng.Font.Color = RGB(156, 101, 0)
            End If
        End If
    Next r
    
    For r = 2 To rAttrs - 1
        key = Trim(CStr(wsAttrs.Cells(r, 1).value)) & "|" & Trim(CStr(wsAttrs.Cells(r, 2).value))
        If attrsProblemsDict.Exists(key) Then
            Set codes = attrsProblemsDict(key)
            hasError = False
            hasWarning = False
            For Each code In codes
                sev = Split(CStr(code), "|")
                If UBound(sev) >= 1 Then
                    If sev(1) = "E" Then hasError = True Else hasWarning = True
                End If
            Next code
            
            Set rng = wsAttrs.Range(wsAttrs.Cells(r, 1), wsAttrs.Cells(r, 12))
            If hasError Then
                rng.Interior.Color = RGB(255, 199, 206)
                rng.Font.Color = RGB(156, 0, 6)
            ElseIf hasWarning Then
                rng.Interior.Color = RGB(255, 235, 156)
                rng.Font.Color = RGB(156, 101, 0)
            End If
        End If
    Next r
End Sub

' ============================================================
' НОВОЕ: ЛИСТ "РЕЕСТРЫ" — Реестр типов + Реестр покрытия атрибутами
' ============================================================
Private Sub BuildRegistrySheet(wsReg As Worksheet, wsTypes As Worksheet, wsAttrs As Worksheet, _
                               rTypes As Long, rAttrs As Long)
    Dim attrCount As Object
    Dim typeNames As Object
    Set attrCount = CreateObject("Scripting.Dictionary")
    Set typeNames = CreateObject("Scripting.Dictionary")
    
    Dim r As Long
    Dim nm As String
    
    ' Подсчет атрибутов по типам
    For r = 2 To rAttrs - 1
        nm = Trim(CStr(wsAttrs.Cells(r, 1).value))
        If nm <> "" Then
            If attrCount.Exists(nm) Then attrCount(nm) = attrCount(nm) + 1 Else attrCount.Add nm, 1
        End If
    Next r
    
    ' Множество имен типов из листа "Типы"
    For r = 2 To rTypes - 1
        nm = Trim(CStr(wsTypes.Cells(r, 1).value))
        If nm <> "" And Not typeNames.Exists(nm) Then typeNames.Add nm, r
    Next r
    
    uncoveredTypesCount = 0
    orphanCoverageCount = 0
    
    ' ---------- БЛОК 1: РЕЕСТР ТИПОВ ----------
    wsReg.Cells(1, 1).value = "РЕЕСТР ТИПОВ (созданные в итоговой сборке)"
    wsReg.Cells(1, 1).Font.Bold = True
    wsReg.Cells(2, 1).value = "№"
    wsReg.Cells(2, 2).value = "Системное имя типа"
    wsReg.Cells(2, 3).value = "Видимое имя типа"
    wsReg.Cells(2, 4).value = "Атрибутов"
    wsReg.Cells(2, 5).value = "Покрытие атрибутами"
    wsReg.Range(wsReg.Cells(2, 1), wsReg.Cells(2, 5)).Font.Bold = True
    wsReg.Range(wsReg.Cells(2, 1), wsReg.Cells(2, 5)).Interior.Color = RGB(221, 235, 247)
    
    Dim i As Long, rowOut As Long, cnt As Long
    i = 0
    For r = 2 To rTypes - 1
        nm = Trim(CStr(wsTypes.Cells(r, 1).value))
        If nm <> "" Then
            i = i + 1
            rowOut = i + 2
            cnt = 0
            If attrCount.Exists(nm) Then cnt = attrCount(nm)
            
            wsReg.Cells(rowOut, 1).value = i
            wsReg.Cells(rowOut, 2).value = nm
            wsReg.Cells(rowOut, 3).value = Trim(CStr(wsTypes.Cells(r, 2).value))
            wsReg.Cells(rowOut, 4).value = cnt
            If cnt > 0 Then
                wsReg.Cells(rowOut, 5).value = "ДА"
            Else
                wsReg.Cells(rowOut, 5).value = "НЕТ"
                wsReg.Range(wsReg.Cells(rowOut, 1), wsReg.Cells(rowOut, 5)).Interior.Color = RGB(255, 199, 206)
                wsReg.Range(wsReg.Cells(rowOut, 1), wsReg.Cells(rowOut, 5)).Font.Color = RGB(156, 0, 6)
                uncoveredTypesCount = uncoveredTypesCount + 1
                AddWarning "R001", "Тип '" & nm & "' отсутствует в Реестре покрытия атрибутами: в листе 'Атрибуты' нет ни одной строки для этого типа.", _
                           "", "Реестры", rowOut, nm, "TYPE"
            End If
        End If
    Next r
    
    ' ---------- БЛОК 2: РЕЕСТР ПОКРЫТИЯ АТРИБУТАМИ ----------
    Dim start2 As Long
    start2 = i + 2 + 3
    wsReg.Cells(start2, 1).value = "РЕЕСТР ПОКРЫТИЯ ТИПОВ АТРИБУТАМИ"
    wsReg.Cells(start2, 1).Font.Bold = True
    wsReg.Cells(start2 + 1, 1).value = "№"
    wsReg.Cells(start2 + 1, 2).value = "Системное имя типа"
    wsReg.Cells(start2 + 1, 3).value = "Количество атрибутов"
    wsReg.Cells(start2 + 1, 4).value = "Наличие в Реестре типов"
    wsReg.Range(wsReg.Cells(start2 + 1, 1), wsReg.Cells(start2 + 1, 4)).Font.Bold = True
    wsReg.Range(wsReg.Cells(start2 + 1, 1), wsReg.Cells(start2 + 1, 4)).Interior.Color = RGB(221, 235, 247)
    
    Dim k As Variant, jj As Long, row2 As Long
    jj = 0
    For Each k In attrCount.Keys
        jj = jj + 1
        row2 = start2 + 1 + jj
        wsReg.Cells(row2, 1).value = jj
        wsReg.Cells(row2, 2).value = CStr(k)
        wsReg.Cells(row2, 3).value = attrCount(k)
        If typeNames.Exists(CStr(k)) Then
            wsReg.Cells(row2, 4).value = "ДА"
        Else
            wsReg.Cells(row2, 4).value = "НЕТ"
            wsReg.Range(wsReg.Cells(row2, 1), wsReg.Cells(row2, 4)).Interior.Color = RGB(255, 199, 206)
            wsReg.Range(wsReg.Cells(row2, 1), wsReg.Cells(row2, 4)).Font.Color = RGB(156, 0, 6)
            orphanCoverageCount = orphanCoverageCount + 1
        End If
    Next k
    
    ' ---------- БЛОК 3: КОНТРОЛЬНЫЕ ИТОГИ ----------
    Dim start3 As Long
    start3 = start2 + 1 + jj + 3
    wsReg.Cells(start3, 1).value = "КОНТРОЛЬНЫЕ ИТОГИ"
    wsReg.Cells(start3, 1).Font.Bold = True
    wsReg.Cells(start3 + 1, 1).value = "Типов в Реестре типов"
    wsReg.Cells(start3 + 1, 2).value = i
    wsReg.Cells(start3 + 2, 1).value = "Типов в Реестре покрытия"
    wsReg.Cells(start3 + 2, 2).value = jj
    wsReg.Cells(start3 + 3, 1).value = "Типов без покрытия атрибутами (R001)"
    wsReg.Cells(start3 + 3, 2).value = uncoveredTypesCount
    If uncoveredTypesCount > 0 Then wsReg.Cells(start3 + 3, 2).Font.Color = RGB(156, 0, 6)
    wsReg.Cells(start3 + 4, 1).value = "Записей покрытия без типа в Реестре"
    wsReg.Cells(start3 + 4, 2).value = orphanCoverageCount
    If orphanCoverageCount > 0 Then wsReg.Cells(start3 + 4, 2).Font.Color = RGB(156, 0, 6)
    
    wsReg.Columns("A:E").AutoFit
End Sub

' ============================================================
' СВОДКА
' ============================================================
Private Sub FillSummarySheet(wsReport As Worksheet, rTypes As Long, rAttrs As Long)
    wsReport.Cells(1, 1).value = "Метрика"
    wsReport.Cells(1, 2).value = "Значение"
    wsReport.Rows(1).Font.Bold = True
    
    wsReport.Cells(2, 1).value = "Всего типов"
    wsReport.Cells(2, 2).value = rTypes - 2
    wsReport.Cells(3, 1).value = "Всего атрибутов (после удаления дубликатов)"
    wsReport.Cells(3, 2).value = rAttrs - 2
    wsReport.Cells(4, 1).value = "Уникальных пар (тип|name)"
    wsReport.Cells(4, 2).value = attrDict.Count
    wsReport.Cells(5, 1).value = "Ошибок (требуют исправления)"
    wsReport.Cells(5, 2).value = colErrors.Count
    wsReport.Cells(5, 1).Font.Bold = True
    If colErrors.Count > 0 Then wsReport.Cells(5, 2).Font.Color = RGB(156, 0, 6)
    wsReport.Cells(6, 1).value = "Предупреждений (рекомендации)"
    wsReport.Cells(6, 2).value = colWarnings.Count
    wsReport.Cells(6, 1).Font.Bold = True
    If colWarnings.Count > 0 Then wsReport.Cells(6, 2).Font.Color = RGB(156, 101, 0)
    
    Dim logCount As Long
    If wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).row > 1 Then
        logCount = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).row - 1
    Else
        logCount = 0
    End If
    wsReport.Cells(7, 1).value = "Обработано дубликатов"
    wsReport.Cells(7, 2).value = logCount
    wsReport.Cells(8, 1).value = "Режим"
    wsReport.Cells(8, 2).value = IIf(isSoftMode, "ПРОМЕЖУТОЧНЫЙ (soft)", "ФИНАЛЬНЫЙ")
    wsReport.Cells(9, 1).value = "Дата генерации"
    wsReport.Cells(9, 2).value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ' НОВОЕ: показатели покрытия
    wsReport.Cells(10, 1).value = "Типов без покрытия атрибутами (R001)"
    wsReport.Cells(10, 2).value = uncoveredTypesCount
    wsReport.Cells(10, 1).Font.Bold = True
    If uncoveredTypesCount > 0 Then wsReport.Cells(10, 2).Font.Color = RGB(156, 0, 6)
    wsReport.Cells(11, 1).value = "Записей покрытия без типа в Реестре"
    wsReport.Cells(11, 2).value = orphanCoverageCount
    wsReport.Cells(11, 1).Font.Bold = True
    If orphanCoverageCount > 0 Then wsReport.Cells(11, 2).Font.Color = RGB(156, 0, 6)
End Sub

' ============================================================
' ОЦЕНКА КАЧЕСТВА СТРОКИ АТРИБУТА
' ============================================================
Private Function CalcRowScore(ws As Worksheet, row As Long) As Long
    Dim score As Long
    Dim c As Long
    Dim cellVal As String
    
    score = 0
    For c = 1 To 12
        cellVal = Trim(CStr(ws.Cells(row, c).value))
        If cellVal <> "" Then
            score = score + 1
            Select Case c
                Case 5: score = score + 10
                Case 8: If LCase(cellVal) = "да" Then score = score + 5
                Case 11:
                    If cellVal = "Запрещать создание дубликатов" Then score = score + 3
                    If cellVal = "Показывать предупреждение" Then score = score + 2
                Case 12: score = score + 8
                Case 4: score = score + 2
            End Select
        End If
    Next c
    
    CalcRowScore = score
End Function

' ============================================================
' ИМЯ ТИПА — camelCase (T002)
' OPT-12: сверено со спецификацией: имена типов — camelCase,
' исключения — фазы phase0–phase3, суффикс _doc и коды форм ksN.
' Правила в коде совпадают со спецификацией — логика оставлена.
' ============================================================
Private Function IsCamelTypeName(ByVal s As String) As Boolean
    Dim re As Object, t As String
    t = s
    If t = "phase0" Or t = "phase1" Or t = "phase2" Or t = "phase3" Then IsCamelTypeName = True: Exit Function
    If Right(t, 4) = "_doc" Then t = Left(t, Len(t) - 4)
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "ks[0-9]+": t = re.Replace(t, "")
    re.Pattern = "Ks[0-9]+": t = re.Replace(t, "")
    re.Pattern = "^[a-z][a-zA-Z]*$"
    IsCamelTypeName = re.Test(t)
End Function

' ============================================================
' ИМЯ АТРИБУТА — camelCase (A003)
' OPT-12: по спецификации имена атрибутов — camelCase БЕЗ ЦИФР,
' поэтому цифры в шаблоне запрещены ([0-9] исключён).
' Проверка наличия цифр выполняется отдельно ContainsDigit (A016).
' ============================================================
Private Function IsCamelAttrName(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[a-z][a-zA-Z]*$"
    IsCamelAttrName = re.Test(s)
End Function

' ============================================================
' ЦИФРЫ В ИМЕНИ АТРИБУТА (A016)
' OPT-12: по спецификации в именах атрибутов цифры запрещены полностью —
' правила в коде совпадают со спецификацией, логика оставлена.
' ============================================================
Private Function ContainsDigit(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "[0-9]"
    ContainsDigit = re.Test(s)
End Function

' ============================================================
' ЦИФРЫ В ИМЕНИ ТИПА (T017): только phase0–phase3 и ksN
' OPT-12: по спецификации цифры в именах типов допустимы только у фаз
' phase0–phase3 и кодов форм ksN — правила в коде совпадают
' со спецификацией, логика оставлена.
' ============================================================
Private Function ContainsForbiddenDigit(ByVal s As String) As Boolean
    Dim re As Object
    Dim t As String
    t = s
    If t = "phase0" Or t = "phase1" Or t = "phase2" Or t = "phase3" Then
        ContainsForbiddenDigit = False
        Exit Function
    End If
    If Right(t, 4) = "_doc" Then t = Left(t, Len(t) - 4)
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "ks[0-9]+"
    t = re.Replace(t, "")
    re.Pattern = "[0-9]"
    ContainsForbiddenDigit = re.Test(t)
End Function

' ============================================================
' ТОЧКИ ВХОДА
' ============================================================
Sub RunSoftMode()
    ValidateAndAssemble True
End Sub

Sub RunFinalMode()
    ValidateAndAssemble False
End Sub

