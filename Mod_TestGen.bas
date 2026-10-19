Attribute VB_Name = "Mod_TestGen"
Option Explicit

' ============================================================
' ГЕНЕРАТОР ТЕСТОВОГО ФАЙЛА (ТЕСТОВЫЙ ПОЛИГОН)
' Создаёт Схему с 4 типами ошибок для проверки валидатора.
' ============================================================
Public Sub GenerateTestSchema()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim savePath As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
    ws.Name = "Структура"
    
    ' --- 1. Легенда (A1:B6) ---
    ' ВАЖНО: В Mod_Structure названия берутся из колонки B (B1, B2, B3),
    ' а цвета — из колонки C (C1, C2, C3).
    ' Но в реальной схеме часто бывает до 6 типов.
    ' Мы заполним B1:B6 названиями, а C1:C6 - цветами.
    
    ' Задаём названия (B) и цвета (C)
    ws.Range("B1").value = "Проект":       ws.Range("C1").Interior.Color = RGB(248, 203, 173) ' розовый
    ws.Range("B2").value = "Папка":        ws.Range("C2").Interior.Color = RGB(255, 230, 153) ' жёлтый
    ws.Range("B3").value = "ECM-документ": ws.Range("C3").Interior.Color = RGB(198, 224, 180) ' зелёный
    ws.Range("B4").value = "Документ":     ws.Range("C4").Interior.Color = RGB(180, 198, 231) ' синий
    ws.Range("B5").value = "Папка ЦИМ":    ws.Range("C5").Interior.Color = RGB(220, 197, 237) ' фиолетовый
    ws.Range("B6").value = "Ярлык":        ws.Range("C6").Interior.Color = RGB(219, 219, 219) ' серый
    
    ' --- 2. Данные (с строки 8) ---
    Dim r As Long
    r = 8
    
    ' Строка 8: Корень (без ошибок). Уровень 1 (Столбец A)
    ws.Cells(r, 1).value = "Объект"
    ws.Cells(r, 1).Interior.Color = ws.Range("C1").Interior.Color ' Проект
    r = r + 1
    
    ' Строка 9: Нормальная папка. Уровень 2 (Столбец B)
    ws.Cells(r, 2).value = "1 Нормальная папка"
    ws.Cells(r, 2).Interior.Color = ws.Range("C2").Interior.Color ' Папка
    r = r + 1
    
    ' Строка 10: Нормальный ECM-документ. Уровень 3 (Столбец C)
    ws.Cells(r, 3).value = "1.1 Корректный документ"
    ws.Cells(r, 3).Interior.Color = ws.Range("C3").Interior.Color ' ECM
    r = r + 1
    
    ' --- ОШИБКА 1: CODE_PARENT_MISMATCH ---
    ' Код ребенка "3.1" не продолжает код родителя "2"
    ws.Cells(r, 2).value = "2 Папка с ошибкой кода"
    ws.Cells(r, 2).Interior.Color = ws.Range("C2").Interior.Color
    r = r + 1
    ws.Cells(r, 3).value = "3.1 Неверный код"
    ws.Cells(r, 3).Interior.Color = ws.Range("C3").Interior.Color
    r = r + 1
    
    ' --- ОШИБКА 2: CODE_DEPTH_JUMP ---
    ' Родитель на 2 уровне, ребенок сразу на 4-м.
    ws.Cells(r, 2).value = "3 Папка для скачка"
    ws.Cells(r, 2).Interior.Color = ws.Range("C2").Interior.Color
    r = r + 1
    ' Пропускаем уровень 3 (Столбец C) и пишем на 4 (Столбец D)
    ws.Cells(r, 4).value = "3.1.1.1 Скачок"
    ws.Cells(r, 4).Interior.Color = ws.Range("C3").Interior.Color
    r = r + 1
    
    ' --- ОШИБКА 3: ORPHAN (Сирота) ---
    ' Важно: чтобы алгоритм не нашел родителя, столбец B (уровень 2)
    ' в этой строке должен быть ПУСТ. Также нужно очистить "память"
    ' предыдущего уровня. Для этого вставим пустую строку или
    ' просто оставим B пустым.
    ws.Cells(r, 2).value = "" ' Пусто в уровне 2
    ws.Cells(r, 3).value = "4.1 Сирота"
    ws.Cells(r, 3).Interior.Color = ws.Range("C3").Interior.Color
    r = r + 1
    
    ' --- ОШИБКА 4: SHORTCUT_NO_TARGET (Ярлык без цели) ---
    ws.Cells(r, 2).value = "4 Папка для ярлыка"
    ws.Cells(r, 2).Interior.Color = ws.Range("C2").Interior.Color
    r = r + 1
    ws.Cells(r, 3).value = "Несуществующий документ"
    ws.Cells(r, 3).Interior.Color = ws.Range("C6").Interior.Color ' Ярлык
    
    ' --- 3. Сохранение ---
    savePath = Mod_Util.GetProjectRoot & "\01_Исходные данные\01_Схема\Тест_Ошибок.xlsx"
    
    On Error Resume Next
    Kill savePath
    On Error GoTo 0
    
    wb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Тестовый файл создан:" & vbCrLf & savePath & vbCrLf & vbCrLf & _
           "В нём 4 ошибки: CODE_PARENT_MISMATCH, CODE_DEPTH_JUMP, ORPHAN, SHORTCUT_NO_TARGET.", vbInformation
End Sub

