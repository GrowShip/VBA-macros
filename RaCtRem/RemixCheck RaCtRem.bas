Attribute VB_Name = "RemixCheck"
Option Explicit

Public Sub CompareRemixGui()
    'Dim ask As Long: ask = frmDidYouUpload.Loading("Ты загрузил для сравнения [Media] в [RemixUpload] и [MediaGuiLangAttr] в [GuiREMIXupload] для данного цикла?")
    
    'If ask = 2 Then
    '    Exit Sub
    'ElseIf ask = 6 Then
    '    CompareRemixAndGuiForMistakes
    'ElseIf ask = 7 Then
    '    MsgBox "Дозагрузи необходимые вкладки"
    '    Exit Sub
    'End If
    CompareRemixAndGuiForMistakes
End Sub

Private Sub CompareRemixAndGuiForMistakes()
    frmLoad.Loading
    ЭтаКнига.Opened
    
    Remix.CreateRemixAndGuiInLock
     
    CompareRemix
    CompareGui
     
    ЭтаКнига.Closed
    frmLoad.Unloading
End Sub

Private Sub CompareRemix()
    Dim remixSh As Worksheet: Set remixSh = ThisWorkbook.Sheets("REMIXlock")
    Dim remixUplSh As Worksheet: Set remixUplSh = ThisWorkbook.Sheets("REMIXupload")
    
    Dim rowsSh As Long: rowsSh = MyFunct.countRowBest(remixSh.name)
    Dim rowsUplSh As Long: rowsUplSh = MyFunct.countRowBest(remixUplSh.name)
    Dim i As Long
    
    Dim searchebleRg As Range: Set searchebleRg = remixSh.Range("B2:B" & rowsSh)
    
    For i = 2 To rowsUplSh
        Dim rgFinded As Range
        
        Set rgFinded = searchebleRg.Find(Split(remixUplSh.Cells(i, 2).value, "_")(2), LookIn:=xlValues, LookAt:=xlPart)
        If Not rgFinded Is Nothing Then
            remixSh.Cells(rgFinded.EntireRow.row, 58) = i
            remixUplSh.Range("A" & i & ":BD" & i).Interior.Color = xlNone
            Call LetsCompareRemix(remixSh, remixUplSh, rgFinded.EntireRow.row, i)
        Else
            remixUplSh.Range("A" & i & ":BD" & i).Interior.Color = RGB(255, 102, 102)
        End If
    Next i
    
    If rowsSh <> rowsUplSh Then MsgBox "Кол-во элементов [REMIXlock/upload]=[Media] ОТЛИЧАЕТСЯ!"
End Sub

Private Sub LetsCompareRemix(ByRef lockedSh As Worksheet, ByRef uplSh As Worksheet, lockRow As Long, uplRow As Long)
    With lockedSh
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("B" & lockRow & ":B" & lockRow), uplSh.Range("B" & uplRow & ":B" & uplRow))  'parent
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("C" & lockRow & ":C" & lockRow), uplSh.Range("C" & uplRow & ":C" & uplRow))  'exp st
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("D" & lockRow & ":D" & lockRow), uplSh.Range("D" & uplRow & ":D" & uplRow))  'exp end
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("E" & lockRow & ":E" & lockRow), uplSh.Range("E" & uplRow & ":E" & uplRow))  'mediaT
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("F" & lockRow & ":F" & lockRow), uplSh.Range("F" & uplRow & ":F" & uplRow))  'mediaCat
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("H" & lockRow & ":H" & lockRow), uplSh.Range("H" & uplRow & ":H" & uplRow))  'Ratin
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("I" & lockRow & ":I" & lockRow), uplSh.Range("I" & uplRow & ":I" & uplRow))  'ParLock
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("L" & lockRow & ":L" & lockRow), uplSh.Range("L" & uplRow & ":L" & uplRow))  'Coll
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("O" & lockRow & ":O" & lockRow), uplSh.Range("O" & uplRow & ":O" & uplRow))  'Media
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("U" & lockRow & ":U" & lockRow), uplSh.Range("U" & uplRow & ":U" & uplRow))  'platform
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("Y" & lockRow & ":Y" & lockRow), uplSh.Range("Y" & uplRow & ":Y" & uplRow))  'mf wide
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("Z" & lockRow & ":Z" & lockRow), uplSh.Range("Z" & uplRow & ":Z" & uplRow))  'mf mpeg
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("AA" & lockRow & ":AA" & lockRow), uplSh.Range("AA" & uplRow & ":AA" & uplRow))  'mf
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("AB" & lockRow & ":AB" & lockRow), uplSh.Range("AB" & uplRow & ":AB" & uplRow))  'mf
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("AC" & lockRow & ":AC" & lockRow), uplSh.Range("AC" & uplRow & ":AC" & uplRow))  'runtime
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("AD" & lockRow & ":AD" & lockRow), uplSh.Range("AD" & uplRow & ":AD" & uplRow))  'encrypted
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("AE" & lockRow & ":AE" & lockRow), uplSh.Range("AE" & uplRow & ":AE" & uplRow))  'provider
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("AF" & lockRow & ":AF" & lockRow), uplSh.Range("AF" & uplRow & ":AF" & uplRow))  'encoding
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("AJ" & lockRow & ":AJ" & lockRow), uplSh.Range("AJ" & uplRow & ":AJ" & uplRow))  'genre
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("AP" & lockRow & ":AP" & lockRow), uplSh.Range("AP" & uplRow & ":AP" & uplRow))  'featored
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("AR" & lockRow & ":AR" & lockRow), uplSh.Range("AR" & uplRow & ":AR" & uplRow))  'image
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("AU" & lockRow & ":AU" & lockRow), uplSh.Range("AU" & uplRow & ":AU" & uplRow))  'YEAR
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("BD" & lockRow & ":BD" & lockRow), uplSh.Range("BD" & uplRow & ":BD" & uplRow))  'APPS
    End With
End Sub

Private Sub CompareGui()
    Dim remixSh As Worksheet: Set remixSh = ThisWorkbook.Sheets("GuiREMIXlock")
    Dim remixUplSh As Worksheet: Set remixUplSh = ThisWorkbook.Sheets("GuiREMIXupload")
    
    Dim rowsSh As Long: rowsSh = MyFunct.countRowBest(remixSh.name)
    Dim rowsUplSh As Long: rowsUplSh = MyFunct.countRowBest(remixUplSh.name)
    Dim i As Long
    
    Dim searchebleRg As Range: Set searchebleRg = remixSh.Range("R2:R" & rowsSh)
    
    For i = 2 To rowsUplSh
        Dim rgFinded As Range
        
        Set rgFinded = searchebleRg.Find(What:=Trim(remixUplSh.Cells(i, 1).value) & " " & Trim(LCase(remixUplSh.Cells(i, 4).value)), LookIn:=xlValues, LookAt:=xlWhole)
        If Not rgFinded Is Nothing Then
            remixSh.Cells(rgFinded.EntireRow.row, 17) = i
            remixUplSh.Range("A" & i & ":O" & i).Interior.Color = xlNone
            Call LetsCompareGui(remixSh, remixUplSh, rgFinded.EntireRow.row, i)
        Else
            remixUplSh.Range("A" & i & ":O" & i).Interior.Color = RGB(255, 102, 102)
        End If
    Next i
    
    If rowsSh <> rowsUplSh Then MsgBox "Кол-во элементов [GuiREMIXlock/upload]=[MediaGuiLangAttr] ОТЛИЧАЕТСЯ!"
End Sub

Private Sub LetsCompareGui(ByRef lockedSh As Worksheet, ByRef uplSh As Worksheet, lockRow As Long, uplRow As Long)
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("A" & lockRow & ":A" & lockRow), uplSh.Range("A" & uplRow & ":A" & uplRow))  'parent
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("B" & lockRow & ":B" & lockRow), uplSh.Range("B" & uplRow & ":B" & uplRow))  'title
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("C" & lockRow & ":C" & lockRow), uplSh.Range("C" & uplRow & ":C" & uplRow))  'episode
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("D" & lockRow & ":D" & lockRow), uplSh.Range("D" & uplRow & ":D" & uplRow))  'lang
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("E" & lockRow & ":E" & lockRow), uplSh.Range("E" & uplRow & ":E" & uplRow))  'syn
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("F" & lockRow & ":F" & lockRow), uplSh.Range("F" & uplRow & ":F" & uplRow))  'sub
    'Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("H" & lockRow & ":H" & lockRow), uplSh.Range("H" & uplRow & ":H" & uplRow))  'Ratin
    'Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("I" & lockRow & ":I" & lockRow), uplSh.Range("I" & uplRow & ":I" & uplRow))  'ParLock
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("J" & lockRow & ":J" & lockRow), uplSh.Range("J" & uplRow & ":J" & uplRow))  'audio
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("L" & lockRow & ":L" & lockRow), uplSh.Range("L" & uplRow & ":L" & uplRow))  'artiscs
    Call CTRfirst.Compare2ItemsCTR(lockedSh.Range("N" & lockRow & ":N" & lockRow), uplSh.Range("N" & uplRow & ":N" & uplRow))  'dir
End Sub
