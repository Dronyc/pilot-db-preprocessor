Attribute VB_Name = "Mod_TestGen"
Option Explicit

' ============================================================
' ÃÅÍÅÐÀÒÎÐ ÒÅÑÒÎÂÎÃÎ ÔÀÉËÀ (ÒÅÑÒÎÂÛÉ ÏÎËÈÃÎÍ)
' Ñîçäà¸ò Ñõåìó ñ 4 òèïàìè îøèáîê äëÿ ïðîâåðêè âàëèäàòîðà.
' ============================================================
Public Sub GenerateTestSchema()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim savePath As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
    ws.Name = "Ñòðóêòóðà"
    
    ' --- 1. Ëåãåíäà (A1:C3) ---
    ' Çàäà¸ì áàçîâûå öâåòà, êàê â ðåàëüíîé ñõåìå (èëè ëþáûå äðóãèå)
    ' Ïðîåêò (æ¸ëòûé)
    ws.Range("B1").value = "Ïðîåêò"
    ws.Range("C1").Interior.Color = RGB(255, 230, 153)
    ' ECM-äîêóìåíò (çåë¸íûé)
    ws.Range("B2").value = "ECM-äîêóìåíò"
    ws.Range("C2").Interior.Color = RGB(198, 224, 180)
    ' Ïàïêà (æ¸ëòûé, êàê ïðîåêò, èëè äðóãîé)
    ws.Range("B3").value = "Ïàïêà"
    ws.Range("C3").Interior.Color = RGB(255, 230, 153)
    ' ßðëûê (ñåðûé)
    ' Äîáàâèì â B4:C4, òàê êàê â Mod_Structure LAST_COL = 30 è èùåò òîëüêî 6 ÿ÷ååê,
    ' íî äëÿ òåñòà íàì õâàòèò è ýòèõ. Â ðåàëüíîé ñõåìå ÿðëûê ìîæåò áûòü â ëåãåíäå.
    ' ×òîáû Mod_Structure åãî óâèäåë, íóæíî ïðîïèñàòü â B1:C3.
    ' Ñäåëàåì ßðëûê âìåñòî Ïàïêè â C3? Íåò, ëó÷øå ðàñøèðèì ëåãåíäó.
    ' Â êîäå Mod_Structure ÷èòàþòñÿ ÿ÷åéêè: B1,C1, B2,C2, B3,C3.
    ' Ïî óìîë÷àíèþ òàì 6 òèïîâ. Ñäåëàåì òàê:
    ' B1: Ïðîåêò, C1: Ïàïêà
    ' B2: ECM-äîêóìåíò, C2: ßðëûê
    ' B3: Äîêóìåíò, C3: Ïàïêà ÖÈÌ
    
    ' Ïåðåïèøåì ëåãåíäó ïîä 6 òèïîâ:
    ' 1: Ïðîåêò (ðîçîâûé)
    ws.Range("B1").value = "Ïðîåêò"
    ws.Range("C1").Interior.Color = RGB(248, 203, 173)
    ' 2: Ïàïêà (æ¸ëòûé)
    ws.Range("B2").value = "Ïàïêà"
    ws.Range("C2").Interior.Color = RGB(255, 230, 153)
    ' 3: ECM-äîêóìåíò (çåë¸íûé)
    ws.Range("B3").value = "ECM-äîêóìåíò"
    ws.Range("C3").Interior.Color = RGB(198, 224, 180)
    ' 4: Äîêóìåíò (ñèíèé)
    ws.Range("B4").value = "Äîêóìåíò"
    ws.Range("C4").Interior.Color = RGB(180, 198, 231)
    ' 5: Ïàïêà ÖÈÌ (ôèîëåòîâûé)
    ws.Range("B5").value = "Ïàïêà ÖÈÌ"
    ws.Range("C5").Interior.Color = RGB(220, 197, 237)
    ' 6: ßðëûê (ñåðûé)
    ws.Range("B6").value = "ßðëûê"
    ws.Range("C6").Interior.Color = RGB(219, 219, 219)
    
    ' --- 2. Äàííûå (ñ ñòðîêè 8) ---
    Dim r As Long
    r = 8
    
    ' 1. Êîðåíü (áåç îøèáîê)
    ws.Cells(r, 1).value = "Îáúåêò"
    ws.Cells(r, 1).Interior.Color = ws.Range("C1").Interior.Color ' Ïðîåêò
    r = r + 1
    
    ' 2. Íîðìàëüíàÿ ïàïêà
    ws.Cells(r, 2).value = "1 Íîðìàëüíàÿ ïàïêà"
    ws.Cells(r, 2).Interior.Color = ws.Range("C2").Interior.Color ' Ïàïêà
    r = r + 1
    
    ' 3. Íîðìàëüíûé ECM-äîêóìåíò
    ws.Cells(r, 3).value = "1.1 Êîððåêòíûé äîêóìåíò"
    ws.Cells(r, 3).Interior.Color = ws.Range("C3").Interior.Color ' ECM
    r = r + 1
    
    ' 4. ÎØÈÁÊÀ CODE_PARENT_MISMATCH (êîä ðåáåíêà íå ïðîäîëæàåò ðîäèòåëÿ)
    ws.Cells(r, 2).value = "2 Ïàïêà ñ îøèáêîé êîäà"
    ws.Cells(r, 2).Interior.Color = ws.Range("C2").Interior.Color
    r = r + 1
    ws.Cells(r, 3).value = "3.1 Íåâåðíûé êîä"
    ws.Cells(r, 3).Interior.Color = ws.Range("C3").Interior.Color
    r = r + 1
    
    ' 5. ÎØÈÁÊÀ CODE_DEPTH_JUMP (ñêà÷îê ãëóáèíû)
    ws.Cells(r, 2).value = "3 Ïàïêà äëÿ ñêà÷êà"
    ws.Cells(r, 2).Interior.Color = ws.Range("C2").Interior.Color
    r = r + 1
    ws.Cells(r, 3).value = "3.1 Ïðîìåæóòî÷íàÿ"
    ws.Cells(r, 3).Interior.Color = ws.Range("C2").Interior.Color
    r = r + 1
    ' Ïðîïóñêàåì óðîâåíü C (4) è ïèøåì ñðàçó íà D (5)
    ws.Cells(r, 4).value = "3.1.1.1 Ñêà÷îê"
    ws.Cells(r, 4).Interior.Color = ws.Range("C3").Interior.Color
    r = r + 1
    
    ' 6. ÎØÈÁÊÀ ORPHAN (ñèðîòà, íåò ðîäèòåëÿ)
    ' Ïèøåì ñðàçó íà 3 óðîâåíü, ïðîïóñòèâ 2
    ws.Cells(r, 3).value = "4.1 Ñèðîòà"
    ws.Cells(r, 3).Interior.Color = ws.Range("C3").Interior.Color
    r = r + 1
    
    ' 7. ÎØÈÁÊÀ SHORTCUT_NO_TARGET (ÿðëûê áåç öåëè)
    ws.Cells(r, 2).value = "4 Ïàïêà äëÿ ÿðëûêà"
    ws.Cells(r, 2).Interior.Color = ws.Range("C2").Interior.Color
    r = r + 1
    ws.Cells(r, 3).value = "Íåñóùåñòâóþùèé äîêóìåíò"
    ws.Cells(r, 3).Interior.Color = ws.Range("C6").Interior.Color ' ßðëûê
    
    ' --- 3. Ñîõðàíåíèå ---
    savePath = Mod_Util.GetProjectRoot & "\01_Èñõîäíûå äàííûå\01_Ñõåìà\Òåñò_Îøèáîê.xlsx"
    
    ' Óäàëèòü ñòàðûé ôàéë, åñëè åñòü
    On Error Resume Next
    Kill savePath
    On Error GoTo 0
    
    ' Ñîõðàíÿåì
    wb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Òåñòîâûé ôàéë ñîçäàí:" & vbCrLf & savePath & vbCrLf & vbCrLf & _
           "Â í¸ì 4 îøèáêè: CODE_PARENT_MISMATCH, CODE_DEPTH_JUMP, ORPHAN, SHORTCUT_NO_TARGET.", vbInformation
End Sub

