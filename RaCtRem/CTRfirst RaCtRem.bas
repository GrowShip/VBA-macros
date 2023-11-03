Attribute VB_Name = "CTRfirst"
Option Explicit
'Public ddd As Variant
'Public mmm As Variant
'Public yyy As Variant


'создание на основе сметы и перенос в upload CTR и remove
Public Sub CreateCTR()
    Dim ask As Long: ask = frmDidYouUpload.Loading("“ы загрузил старый RAVE CTR вкладка [VIDEO] в [CTRupload] и вкладку [Video and Audio Deletions] в [RemoveUpload]?")
    
    If ask = 2 Then
        Exit Sub
    ElseIf ask = 6 Then
        CTRupload.CreateFirstCtr
    ElseIf ask = 7 Then
        CreateFirstBasedOnKC
    End If
End Sub

Private Sub CreateFirstBasedOnKC()
    Ёта нига.Opened
    
    clearing.ClearCTRlock (False)
    clearing.ClearRemoveLock (False)
    
    CreationCTR
    
    Ёта нига.Opened
    frmLoad.Loading
    
    Dim shCTRlock As Worksheet, shCTRupl As Worksheet
    Set shCTRlock = ThisWorkbook.Sheets("CTRlock")
    Set shCTRupl = ThisWorkbook.Sheets("CTRupload")

    
    Dim rowCtr As Long, i As Long, status As String, counterDel As Long
    rowCtr = MyFunct.countRowBest(shCTRlock.name)
    counterDel = MyFunct.countRowBest("removeLock")
    
    clearing.ClearRemoveUpload (False)
    
    For i = 2 To rowCtr
        status = shCTRlock.Cells(i, 8).value
            If InStr(1, status, "remove", vbTextCompare) > 0 Then
                With ThisWorkbook.Sheets("RemoveLock")
                    .Cells(counterDel, 1).value = shCTRlock.Cells(i, 1).value
                    .Cells(counterDel, 2).value = shCTRlock.Cells(i, 2).value
                    .Cells(counterDel, 3).value = shCTRlock.Cells(i, 6).value
                    .Cells(counterDel, 4).value = shCTRlock.Cells(i, 30).value
                End With
                counterDel = counterDel + 1
            End If
    Next i
    
    counterDel = CTRupload.RemoveStatusLine(shCTRlock, "I", "delete", rowCtr)
    counterDel = CTRupload.RemoveStatusLine(shCTRlock, "I", "remove", rowCtr)
    
    clearing.ClearCTRupload (False)
    clearing.ClearRemoveUpload (False)
    
    Inserting.CopyCTRlockTOupload (False)
    Inserting.CopyRemoveLockToUpload (False)
    
    clearing.ClearCTRlock (False)
    clearing.ClearRemoveLock (False)
    
    OpeningBook.SetBookConfig
    
    Ёта нига.Closed
    frmLoad.Unloading
    
    MsgBox "Ѕыло удалено со статусом Remove " & counterDel & " строк!"
End Sub

'—оздаетс€ с нул€ под CTRlock
'¬ыдел€етс€ красным внутри lock и upload что отличаетс€ между друг другом
Public Sub CompareCTR()
    Ёта нига.Opened
    
    Dim shCTRlock As Worksheet, shCTRupl As Worksheet, sh As Worksheet, shFF As Worksheet, shRemlock As Worksheet, shRemupl As Worksheet
    Set sh = ThisWorkbook.Sheets("Initial")
    Set shFF = ThisWorkbook.Sheets("Filenames")
    Set shCTRlock = ThisWorkbook.Sheets("CTRlock")
    Set shCTRupl = ThisWorkbook.Sheets("CTRupload")
    Set shRemlock = ThisWorkbook.Sheets("RemoveLock")
    Set shRemupl = ThisWorkbook.Sheets("RemoveUpload")
    
    Dim rowLock As Long
    Dim rowUpl As Long: rowUpl = MyFunct.countRowBest(shCTRupl.name)
    Dim rowRemoveUpl As Long: rowRemoveUpl = MyFunct.countRowBest(shRemupl.name)
    
    If rowUpl < 2 Then
        Ёта нига.Closed
        MsgBox "¬ы не загрузили CTR upload"
        Exit Sub
    End If
    
    If rowRemoveUpl < 2 Then
        Ёта нига.Closed
        MsgBox "¬ы не загрузили Video and Audio Deletions"
        Exit Sub
    End If

    Call CreationCTR
    rowLock = MyFunct.countRowBest(shCTRlock.name)
    rowUpl = MyFunct.countRowBest(shCTRupl.name)
    
    Dim rg As Range: Set rg = shCTRlock.Range("AN2:AN" & rowLock)
    Dim i As Long
    
    For i = 2 To rowUpl
        
        Dim title As String
        Dim rgFinded As Range
        
        With shCTRupl
            If Len(.Cells(i, 2).value) > 1 Then
                title = Trim(.Cells(i, 1).value) & " " & Trim(.Cells(i, 2).value)
            Else: title = Trim(.Cells(i, 1).value)
            End If
        End With
        
        Set rgFinded = rg.Find(title, LookIn:=xlValues, LookAt:=xlWhole)
        If Not rgFinded Is Nothing Then
            shCTRupl.Range("A" & i & ":AL" & i).Interior.Color = xlNone
            shCTRlock.Cells(rgFinded.EntireRow.row, 39).value = i
            Call LetsCompare(shCTRlock, shCTRupl, rgFinded.EntireRow.row, i)
        Else:
            shCTRupl.Range("A" & i & ":AL" & i).Interior.Color = RGB(255, 102, 102)
        End If
    Next i
    
    Dim rowRemoveLock As Long: rowRemoveLock = MyFunct.countRowBest(shRemlock.name)
    Set rg = shRemlock.Range("A2:D" & rowRemoveLock)
    Call frmLoad.Loading
    For i = 2 To rowRemoveUpl
        
        Set rgFinded = rg.Find(shRemupl.Cells(i, 4), LookIn:=xlValues, LookAt:=xlWhole)
        If Not rgFinded Is Nothing Then
            Call Compare2ItemsCTR(shRemlock.Range("A" & rgFinded.EntireRow.row & ":A" & rgFinded.EntireRow.row), _
                                                  shRemupl.Range("A" & i & ":A" & i))  'title
            Call Compare2ItemsCTR(shRemlock.Range("B" & rgFinded.EntireRow.row & ":B" & rgFinded.EntireRow.row), _
                                                  shRemupl.Range("B" & i & ":B" & i))  'episode
            Call Compare2ItemsCTR(shRemlock.Range("C" & rgFinded.EntireRow.row & ":C" & rgFinded.EntireRow.row), _
                                                  shRemupl.Range("C" & i & ":C" & i))  'season numb
            Call Compare2ItemsCTR(shRemlock.Range("D" & rgFinded.EntireRow.row & ":D" & rgFinded.EntireRow.row), _
                                                  shRemupl.Range("D" & i & ":D" & i))  'episod numb
        Else:
            shRemupl.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 102, 102)
        End If
    Next i
    Ёта нига.Closed
    Call frmLoad.Unloading
    
End Sub

Private Sub LetsCompare(ByRef CTRlock As Worksheet, ByRef CTRupl As Worksheet, lockRow As Long, uplRow As Long)
    With CTRlock
    Call Compare2ItemsCTR(CTRlock.Range("A" & lockRow & ":A" & lockRow), CTRupl.Range("A" & uplRow & ":A" & uplRow))  'title
    Call Compare2ItemsCTR(CTRlock.Range("B" & lockRow & ":B" & lockRow), CTRupl.Range("B" & uplRow & ":B" & uplRow))  'episode
    Call Compare2ItemsCTR(CTRlock.Range("C" & lockRow & ":C" & lockRow), CTRupl.Range("C" & uplRow & ":C" & uplRow))  'season numb
    Call Compare2ItemsCTR(CTRlock.Range("D" & lockRow & ":D" & lockRow), CTRupl.Range("D" & uplRow & ":D" & uplRow))  'episod numb
    Call Compare2ItemsCTR(CTRlock.Range("E" & lockRow & ":E" & lockRow), CTRupl.Range("E" & uplRow & ":E" & uplRow))  'priority
    Call Compare2ItemsCTR(CTRlock.Range("F" & lockRow & ":F" & lockRow), CTRupl.Range("F" & uplRow & ":F" & uplRow))  'media
    Call Compare2ItemsCTR(CTRlock.Range("G" & lockRow & ":G" & lockRow), CTRupl.Range("G" & uplRow & ":G" & uplRow))  'runtime
    Call Compare2ItemsCTR(CTRlock.Range("H" & lockRow & ":H" & lockRow), CTRupl.Range("H" & uplRow & ":H" & uplRow))  'version
    Call Compare2ItemsCTR(CTRlock.Range("I" & lockRow & ":I" & lockRow), CTRupl.Range("I" & uplRow & ":I" & uplRow))  'status
    Call Compare2ItemsCTR(CTRlock.Range("J" & lockRow & ":J" & lockRow), CTRupl.Range("J" & uplRow & ":J" & uplRow))  'date start
    Call Compare2ItemsCTR(CTRlock.Range("K" & lockRow & ":K" & lockRow), CTRupl.Range("K" & uplRow & ":K" & uplRow))  'date end
    Call Compare2ItemsCTR(CTRlock.Range("L" & lockRow & ":L" & lockRow), CTRupl.Range("L" & uplRow & ":L" & uplRow))  'd 1
    Call Compare2ItemsCTR(CTRlock.Range("M" & lockRow & ":M" & lockRow), CTRupl.Range("M" & uplRow & ":M" & uplRow))  'd 2
    Call Compare2ItemsCTR(CTRlock.Range("N" & lockRow & ":N" & lockRow), CTRupl.Range("N" & uplRow & ":N" & uplRow))  'd 3
    Call Compare2ItemsCTR(CTRlock.Range("O" & lockRow & ":O" & lockRow), CTRupl.Range("O" & uplRow & ":O" & uplRow))  'd 4
    Call Compare2ItemsCTR(CTRlock.Range("P" & lockRow & ":P" & lockRow), CTRupl.Range("P" & uplRow & ":P" & uplRow))  'd 5
    Call Compare2ItemsCTR(CTRlock.Range("Q" & lockRow & ":Q" & lockRow), CTRupl.Range("Q" & uplRow & ":Q" & uplRow))  'd 6
    Call Compare2ItemsCTR(CTRlock.Range("R" & lockRow & ":R" & lockRow), CTRupl.Range("R" & uplRow & ":R" & uplRow))  'd 7
    Call Compare2ItemsCTR(CTRlock.Range("S" & lockRow & ":S" & lockRow), CTRupl.Range("S" & uplRow & ":S" & uplRow))  'd 8
    Call Compare2ItemsCTR(CTRlock.Range("T" & lockRow & ":T" & lockRow), CTRupl.Range("T" & uplRow & ":T" & uplRow))  'b 1
    Call Compare2ItemsCTR(CTRlock.Range("U" & lockRow & ":U" & lockRow), CTRupl.Range("U" & uplRow & ":U" & uplRow))  'b 2
    Call Compare2ItemsCTR(CTRlock.Range("V" & lockRow & ":V" & lockRow), CTRupl.Range("V" & uplRow & ":V" & uplRow))  's 1
    Call Compare2ItemsCTR(CTRlock.Range("W" & lockRow & ":W" & lockRow), CTRupl.Range("W" & uplRow & ":W" & uplRow))  's 2
    Call Compare2ItemsCTR(CTRlock.Range("X" & lockRow & ":X" & lockRow), CTRupl.Range("X" & uplRow & ":X" & uplRow))  's 3
    Call Compare2ItemsCTR(CTRlock.Range("Y" & lockRow & ":Y" & lockRow), CTRupl.Range("Y" & uplRow & ":Y" & uplRow))  's 4
    Call Compare2ItemsCTR(CTRlock.Range("Z" & lockRow & ":Z" & lockRow), CTRupl.Range("Z" & uplRow & ":Z" & uplRow))  's 5
    Call Compare2ItemsCTR(CTRlock.Range("AA" & lockRow & ":AA" & lockRow), CTRupl.Range("AA" & uplRow & ":AA" & uplRow))  's 6
    Call Compare2ItemsCTR(CTRlock.Range("AB" & lockRow & ":AB" & lockRow), CTRupl.Range("AB" & uplRow & ":AB" & uplRow))  's 7
    Call Compare2ItemsCTR(CTRlock.Range("AC" & lockRow & ":AC" & lockRow), CTRupl.Range("AC" & uplRow & ":AC" & uplRow))  's 8
    Call Compare2ItemsCTR(CTRlock.Range("AD" & lockRow & ":AD" & lockRow), CTRupl.Range("AD" & uplRow & ":AD" & uplRow))  'parent
    Call Compare2ItemsCTR(CTRlock.Range("AE" & lockRow & ":AE" & lockRow), CTRupl.Range("AE" & uplRow & ":AE" & uplRow))  'aspect
    Call Compare2ItemsCTR(CTRlock.Range("AF" & lockRow & ":AF" & lockRow), CTRupl.Range("AF" & uplRow & ":AF" & uplRow))  'resol
    Call Compare2ItemsCTR(CTRlock.Range("AJ" & lockRow & ":AJ" & lockRow), CTRupl.Range("AJ" & uplRow & ":AJ" & uplRow))  'distr
    Call Compare2ItemsCTR(CTRlock.Range("AK" & lockRow & ":AK" & lockRow), CTRupl.Range("AK" & uplRow & ":AK" & uplRow))  'lab
    End With
End Sub

Public Sub Compare2ItemsCTR(CTRlockRangeOne As Range, CTRuplRangeOne As Range)
    Dim compared: compared = StrComp(CTRlockRangeOne, CTRuplRangeOne, vbTextCompare)
    
    If compared = 0 Then
        CTRuplRangeOne.Interior.Color = xlNone
    ElseIf Len(CTRlockRangeOne.text) = 0 And Len(CTRuplRangeOne) = 0 Then
        CTRuplRangeOne.Interior.Color = xlNone
    ElseIf compared = 1 Then
        CTRuplRangeOne.Interior.Color = RGB(204, 255, 204)
    ElseIf compared = -1 Then
        CTRuplRangeOne.Interior.Color = RGB(255, 102, 102)
    End If
End Sub

'создание с нул€ по всем статусам
Public Sub CreationCTR()
    VBAProject.Ёта нига.Opened
    'Dim flagform As Boolean: flagform = False
    
    'flagform = dateForSmtFrom.Loading("ƒата дл€ filenames")
    'If Not flagform Then
        'VBAProject.Ёта нига.Closed
        'Exit Sub
    'End If
    
    Call frmLoad.Loading
    
    Call clearing.ClearCTRlock(False)
    Call clearing.ClearRemoveLock(False)
    
    'ThisWorkbook.Sheets("Notes").Cells(6, 1).value = ddd & "|" & mmm & "|" & yyy
    
    Dim dictStatus As Dictionary: Set dictStatus = GetRaveStatus
    
    '¬ словаре 1 - new, 2 - old, 3 - remove, 4 - delete
    Call FillingCTRlock("delete", dictStatus)
    Call FillingCTRlock("remove", dictStatus)
    Call FillingCTRlock("old", dictStatus)
    Call FillingCTRlock("new", dictStatus)
    
    OpeningBook.SetBookConfig
    
    frmLoad.Unloading
    
    VBAProject.Ёта нига.Closed
End Sub

Public Sub FindRaveStatus(ByRef sh As Worksheet, ByRef noteSh As Worksheet, i As Long)
  
    Dim stNew As Range: Set stNew = noteSh.Range("G1")
    Dim stOld As Range: Set stOld = noteSh.Range("G2")
    Dim stRemove As Range: Set stRemove = noteSh.Range("G3")
    Dim stDelete As Range: Set stDelete = noteSh.Range("G4")
   
    If InStr(1, sh.Cells(i, 16).value, "new", vbTextCompare) > 0 Then
        stNew = stNew.value & CStr(i) & " "
    ElseIf InStr(1, sh.Cells(i, 16).value, "old", vbTextCompare) > 0 Then
        stOld = stOld.value & CStr(i) & " "
    ElseIf InStr(1, sh.Cells(i, 16).value, "remove", vbTextCompare) > 0 Then
        stRemove = stRemove.value & CStr(i) & " "
    ElseIf InStr(1, sh.Cells(i, 16).value, "delete", vbTextCompare) > 0 Then
        stDelete = stDelete.value & CStr(i) & " "
    Else
        End If

End Sub

Private Function GetRaveStatus() As Dictionary
    Dim stNew As String, stOld As String, stDelete As String, stRemove As String
    Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("NOTES")
    
    stNew = ThisWorkbook.Sheets("NOTES").Cells(1, 7).value
    stOld = ThisWorkbook.Sheets("NOTES").Cells(2, 7).value
    stRemove = ThisWorkbook.Sheets("NOTES").Cells(3, 7).value
    stDelete = ThisWorkbook.Sheets("NOTES").Cells(4, 7).value
    
    Dim dictStatus As Dictionary
    Set dictStatus = New Dictionary
    
    If Len(stNew) > 2 Then dictStatus.Item("new") = GetStatusArray(stNew)
    If Len(stOld) > 2 Then dictStatus.Item("old") = GetStatusArray(stOld)
    If Len(stRemove) > 2 Then dictStatus.Item("remove") = GetStatusArray(stRemove)
    If Len(stDelete) > 2 Then dictStatus.Item("delete") = GetStatusArray(stDelete)
    
    Set GetRaveStatus = dictStatus
End Function

Private Function GetStatusArray(statusStr As String) As Variant
        GetStatusArray = Split(statusStr, " ")
End Function

Public Sub FillingCTRlock(status As String, Optional dictStatus As Dictionary)
    Dim i As Long, j As Long, rowKC As Long
    Dim ctrLockSh As Worksheet: Set ctrLockSh = ThisWorkbook.Sheets("CTRlock")
    Dim initSh As Worksheet: Set initSh = ThisWorkbook.Sheets("Initial")
    Dim filenamesSh As Worksheet: Set filenamesSh = ThisWorkbook.Sheets("Filenames")
    Dim removeLock As Worksheet: Set removeLock = ThisWorkbook.Sheets("RemoveLock")
    Dim oneStatList As Variant
    
    If Not (dictStatus Is Nothing) Then
        oneStatList = dictStatus(status)
    ElseIf InStr(1, status, "new", vbTextCompare) > 0 Then
        oneStatList = GetStatusArray(ThisWorkbook.Sheets("NOTES").Cells(1, 7))
    ElseIf InStr(1, status, "old", vbTextCompare) > 0 Then
        oneStatList = GetStatusArray(ThisWorkbook.Sheets("NOTES").Cells(2, 7))
    ElseIf InStr(1, status, "remove", vbTextCompare) > 0 Then
        oneStatList = GetStatusArray(ThisWorkbook.Sheets("NOTES").Cells(3, 7))
    ElseIf InStr(1, status, "delete", vbTextCompare) > 0 Then
        oneStatList = GetStatusArray(ThisWorkbook.Sheets("NOTES").Cells(4, 7))
    End If
    
    j = MyFunct.countRowBest(ctrLockSh.name) + 1
    For i = 0 To UBound(oneStatList) - 1
        If ctrLockSh.Range("AM1:AM" & j).Find(oneStatList(i), LookIn:=xlValues, LookAt:=xlWhole) Is Nothing Then
            rowKC = CLng(oneStatList(i))
        
            If (StrComp(status, "old", vbTextCompare) = 0) Then
                ctrLockSh.Cells(j, 9).value = "Holdover"
            ElseIf (StrComp(status, "new", vbTextCompare) = 0) Then
                ctrLockSh.Cells(j, 9).value = "New"
            ElseIf (StrComp(status, "remove", vbTextCompare) = 0) Then
                ctrLockSh.Cells(j, 9).value = "Remove"
            Else: ctrLockSh.Cells(j, 9).value = status
            End If
        
            Call CtrMainInfo(ctrLockSh, initSh, filenamesSh, j, rowKC, removeLock)
            j = j + 1
        End If
    Next i
End Sub

Private Sub CtrMainInfo(ByRef ctrLockSh As Worksheet, ByRef initSh As Worksheet, ByRef filenamesSh As Worksheet, rowCtr As Long, rowKC As Long, ByRef removeLock As Worksheet)
    With ctrLockSh
        .Cells(rowCtr, 1).value = filenamesSh.Cells(rowKC, 1).value 'title
        .Cells(rowCtr, 2).value = filenamesSh.Cells(rowKC, 2).value 'episode
        .Cells(rowCtr, 3).value = filenamesSh.Cells(rowKC, 3).value  'season numb
        .Cells(rowCtr, 4).value = filenamesSh.Cells(rowKC, 4).value  'episode numb
        .Cells(rowCtr, 5).value = "No" 'Priority
        .Cells(rowCtr, 6).value = filenamesSh.Cells(rowKC, 7).value  'MorTV
        .Cells(rowCtr, 7).value = initSh.Cells(rowKC, 34).value  'Runtime
        .Cells(rowCtr, 8).value = filenamesSh.Cells(rowKC, 17).value  'Version
        .Cells(rowCtr, 10).value = CDate(filenamesSh.Cells(rowKC, 13) & "-" & filenamesSh.Cells(rowKC, 12) & "-01") 'date start
        .Cells(rowCtr, 11).value = CDate(filenamesSh.Cells(rowKC, 13) & "-12-31") 'date end
        '.Cells(rowCtr,20).value =   'burned subs to 21
        '.Cells(rowCtr,22).value =   'dyn sub to 29
        .Cells(rowCtr, 30).value = ParentTitle(ctrLockSh, filenamesSh, rowCtr, rowKC) 'parent title
        .Cells(rowCtr, 31).value = filenamesSh.Cells(rowKC, 16).value ' Aspect size
        .Cells(rowCtr, 32).value = "480p" ' Resolution
        .Cells(rowCtr, 36).value = initSh.Cells(rowKC, 19).value 'Distributor
        .Cells(rowCtr, 37).value = Replace(initSh.Cells(rowKC, 20).value, "MG Lab", "Digital IFE Services Ltd", , , vbTextCompare) 'Lab
        
        If Len(.Cells(rowCtr, 2).value) > 1 Then
            .Cells(rowCtr, 40).value = Trim(.Cells(rowCtr, 1).value) & " " & Trim(.Cells(rowCtr, 2).value)
        Else: .Cells(rowCtr, 40).value = Trim(.Cells(rowCtr, 1).value)
        End If
        
        If InStr(1, .Cells(rowCtr, 9), "remove", vbTextCompare) > 0 Then
            Dim last As Long: last = MyFunct.countRowBest(removeLock.name) + 1
            removeLock.Cells(last, 1) = .Cells(rowCtr, 1).value
            removeLock.Cells(last, 2) = .Cells(rowCtr, 2).value
            removeLock.Cells(last, 3) = .Cells(rowCtr, 6).value
            removeLock.Cells(last, 4) = .Cells(rowCtr, 30).value
        End If
    End With
    
    Call CtrLockDubInserting(ctrLockSh, filenamesSh, rowCtr, rowKC) 'dubs to 19
    Call CtrLockSubInserting(ctrLockSh, filenamesSh, rowCtr, rowKC) 'sub to 29
    
    
End Sub

Private Function ParentTitle(ctrLockSh As Worksheet, filenamesSh As Worksheet, rowCtr As Long, rowKC As Long) As String
    If (Len(filenamesSh.Cells(rowKC, 9).value) > 1 And Len(filenamesSh.Cells(rowKC, 18).value) > 1) Then
        Dim burnetDub As New StringBuilderMy
        If Len(filenamesSh.Cells(rowKC, 37)) > 0 Then
            Dim i As Integer
            For i = 0 To 5
                If Not IsEmpty(filenamesSh.Cells(rowKC, 37 + i)) Then
                    burnetDub.Append (filenamesSh.Cells(rowKC, 37 + i))
                End If
            Next i
        End If
        ParentTitle = Replace(Replace(Replace(Replace(filenamesSh.Cells(rowKC, 9).value, "#", ""), "_SSS", "", , , vbTextCompare), "DDD", Replace(filenamesSh.Cells(rowKC, 18).value, "AD", ""), , , vbTextCompare), "|", "") & burnetDub.ToString & ".mp4"
    Else: ParentTitle = ""
    End If
End Function

Private Function DateForTitle(ByRef filenamesSh As Worksheet, rowKC As Long) As String
    If InStr(1, filenamesSh.Cells(rowKC, 8).value, "new", vbTextCompare) > 0 Then
        DateForTitle = CStr(yyy) & "-" & CStr(mmm) & "-01"
    Else:
        DateForTitle = CStr(filenamesSh.Cells(rowKC, 13).value) & "-" & CStr(filenamesSh.Cells(rowKC, 12).value) & "-01"
    End If
End Function

Private Sub CtrLockDubInserting(ByRef ctrLockSh As Worksheet, ByRef filenamesSh As Worksheet, rowCtr As Long, rowKC As Long)
    Dim i As Long
    Dim dvsNotIn As Boolean: dvsNotIn = True
    
    For i = 12 To 19
        If (Len(filenamesSh.Cells(rowKC, i + 7).value) > 1) Then
            ctrLockSh.Cells(rowCtr, i).value = filenamesSh.Cells(rowKC, i + 7).value
        ElseIf (Len(filenamesSh.Cells(rowKC, 35).value) > 1 And dvsNotIn) Then
            ctrLockSh.Cells(rowCtr, i).value = Left(Replace(filenamesSh.Cells(rowKC, 35).value, "EngAD", "Dvs"), 3)
            dvsNotIn = False
        Else:
            Exit For
        End If
    Next i
End Sub

Private Sub CtrLockSubInserting(ByRef ctrLockSh As Worksheet, ByRef filenamesSh As Worksheet, rowCtr As Long, rowKC As Long)
    Dim i As Long
    Dim ccNotIn As Boolean: ccNotIn = True
    Dim burnedNotIn As Boolean: burnedNotIn = True
    
    For i = 22 To 29
        If (Len(filenamesSh.Cells(rowKC, i + 7).value) > 1) And (i + 7 < 34) Then
            ctrLockSh.Cells(rowCtr, i).value = Replace(Replace(Replace(filenamesSh.Cells(rowKC, 9).value, "#", ""), "SSS", Left(filenamesSh.Cells(rowKC, i + 7).value, 3)), "_DDD", "") & ".srt"
        ElseIf (Len(filenamesSh.Cells(rowKC, 36).value) > 1 And ccNotIn) Then
            ctrLockSh.Cells(rowCtr, i).value = Replace(Replace(Replace(filenamesSh.Cells(rowKC, 9).value, "#", ""), "SSS", filenamesSh.Cells(rowKC, 36).value), "_DDD", "") & ".srt"
            ccNotIn = False
        ElseIf (Len(filenamesSh.Cells(rowKC, 37).value) > 1) And burnedNotIn Then
            ctrLockSh.Cells(rowCtr, 20).value = Left(filenamesSh.Cells(rowKC, 37).value, 3) 'Replace(Replace(Replace(filenamesSh.Cells(rowKC, 9).value, "#", ""), "SSS", Left(filenamesSh.Cells(rowKC, 37).value, 3)), "_DDD", "") & ".srt"
            burnedNotIn = False
        Else:
            Exit For
        End If
    Next i
End Sub
