Attribute VB_Name = "CTRupload"
Option Explicit
'на основе сметы
Public Sub CreateFirstCtr()
    Ёта нига.Opened
    
    'Dim flagform As Boolean: flagform = False
    'flagform = dateForSmtFrom.Loading("ƒата дл€ CTR filenames NEW")
    'If Not flagform Then
        'VBAProject.Ёта нига.Closed
        'Exit Sub
    'End If
    
    Dim shCTRlock As Worksheet, shCTRupl As Worksheet, sh As Worksheet, shFF As Worksheet
    Set sh = ThisWorkbook.Sheets("Initial")
    Set shFF = ThisWorkbook.Sheets("Filenames")
    Set shCTRlock = ThisWorkbook.Sheets("CTRlock")
    Set shCTRupl = ThisWorkbook.Sheets("CTRupload")
    Dim status As String
    Dim rowsCtr As Long, i As Long
    Dim counterDel As Long: counterDel = 2
    
    'проверка lock
    rowsCtr = MyFunct.countRowBest(shCTRlock.name)
    If rowsCtr > 1 Then clearing.ClearCTRlock (False)
    clearing.ClearRemoveLock (False)
    
    Inserting.CopyCTRuploadTOlock (False)
    clearing.ClearCTRupload (False)
    rowsCtr = MyFunct.countRowBest(shCTRlock.name)
    
    frmLoad.Loading
    For i = 2 To rowsCtr
        Call searchOne(shFF, shCTRlock, i)
        If Len(shCTRlock.Cells(i, 39).value) > 0 Then
            'change status
            status = shFF.Cells(shCTRlock.Cells(i, 39).value, 8).value
            shCTRlock.Cells(i, 9).value = Replace(status, "old", "Holdover", , , vbTextCompare)
            If InStr(1, status, "remove", vbTextCompare) > 0 Or InStr(1, status, "delete", vbTextCompare) > 0 Then
                With ThisWorkbook.Sheets("RemoveLock")
                    .Cells(counterDel, 1).value = shCTRlock.Cells(i, 1).value
                    .Cells(counterDel, 2).value = shCTRlock.Cells(i, 2).value
                    .Cells(counterDel, 3).value = shCTRlock.Cells(i, 6).value
                    .Cells(counterDel, 4).value = shCTRlock.Cells(i, 30).value
                End With
                counterDel = counterDel + 1
            End If
        End If
    Next i
    
    Call CTRfirst.FillingCTRlock("New")
    Dim counter As Long
    counter = RemoveStatusLine(shCTRlock, "I", "Remove", rowsCtr)
    
    Inserting.CopyCTRlockTOupload (False)
    Inserting.CopyRemoveLockToUpload (False)
    
    clearing.ClearCTRlock (False)
    clearing.ClearRemoveLock (False)
    
    OpeningBook.SetBookConfig
    
    Ёта нига.Closed
    frmLoad.Unloading
    
    MsgBox "Ѕыло удалено со статусом Remove " & counter & " строк!"
End Sub

Public Function RemoveStatusLine(sh As Worksheet, columnLetter As String, status As String, Optional rowsCtr As Long) As Long
    Dim rg As Range: Set rg = sh.Range(columnLetter & "1:" & columnLetter & rowsCtr)
    Dim counter As Long: counter = 0
    With rg
        Dim rgRemove As Range: Set rgRemove = .Find(status, LookIn:=xlValues, LookAt:=xlWhole)
        If Not rgRemove Is Nothing Then
            Do
                sh.rows(rgRemove.EntireRow.row).EntireRow.Delete
                counter = counter + 1
                rowsCtr = MyFunct.countRowBest(sh.name)
                Set rg = sh.Range(columnLetter & "1:" & columnLetter & rowsCtr)
                Set rgRemove = .Find(status, LookIn:=xlValues, LookAt:=xlWhole)
            Loop While Not rgRemove Is Nothing
        End If
    End With
    RemoveStatusLine = counter
End Function

Private Sub searchOne(ByRef shFn As Worksheet, ByRef shCTRlock As Worksheet, rowCtr As Long)
    Dim name As String, rowNum As Range
    Dim searchableRange As Range: Set searchableRange = shFn.Range("J1:J" & MyFunct.countRowBest(shFn.name))
           
    If Len(shCTRlock.Cells(rowCtr, 2).value) > 2 Then
        name = Trim(shCTRlock.Cells(rowCtr, 1).value) + " " + Trim(shCTRlock.Cells(rowCtr, 2).value)
    Else: name = Trim(shCTRlock.Cells(rowCtr, 1).value)
    End If
        
    Set rowNum = searchableRange.Find(name, LookIn:=xlValues, LookAt:=xlWhole)
    If rowNum Is Nothing Then
        shCTRlock.Range("A" & rowCtr & ":AK" & rowCtr).Interior.Color = RGB(255, 102, 102)
    Else:
        shCTRlock.Range("A" & rowCtr & ":AK" & rowCtr).Interior.Color = xlNone
        shCTRlock.Cells(rowCtr, 39) = rowNum.EntireRow.row
    End If
End Sub

Private Function searchAll() As Dictionary
    Dim counter As Long, lastRow As Long, i As Long, name As String, rowNum As Range, notFined As String
    
    
    Dim shFn As Worksheet: Set shFn = ThisWorkbook.Sheets("Filenames")
    lastRow = MyFunct.countRowBest(shFn.name)
    Dim searchableRange As Range: Set searchableRange = shFn.Range("J1:J" & lastRow)
    
    Dim shCTRupl As Worksheet: Set shCTRupl = ThisWorkbook.Sheets("CTRupload")
    lastRow = MyFunct.countRowBest(shCTRupl.name)
    
    Dim dictStatus As Dictionary
    Set dictStatus = New Dictionary
    
    For i = 2 To lastRow
        If Len(shCTRupl.Cells(i, 2).value) > 2 Then
            name = Trim(shCTRupl.Cells(i, 1).value) + " " + Trim(shCTRupl.Cells(i, 2).value)
        Else: name = Trim(shCTRupl.Cells(i, 1).value)
        End If
        
        Set rowNum = searchableRange.Find(name, LookIn:=xlValues)
        If rowNum Is Nothing Then
            notFined = notFined + "—трока: " + CStr(i) + " Ќазвание: " + name + Chr(10)
        Else: dictStatus.Item(i) = rowNum.EntireRow.row
        End If
    Next i
    
    MsgBox notFined
    Set searchAll = dictStatus
End Function

