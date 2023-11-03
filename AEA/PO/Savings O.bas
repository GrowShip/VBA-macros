Attribute VB_Name = "Savings"
Option Explicit

Sub SavePOs()
    Dim sh As Worksheet
    Dim typeSave As String
    Dim datet As String
    Dim arr, studia As String
    
    Set sh = ThisWorkbook.Sheets("Information")
    
    If Len(sh.Cells(22, 12).value) = 0 Then
        MsgBox "Не выбран тип сохранения"
        Exit Sub
    Else
        sh.Cells(22, 13).value = sh.Cells(22, 12).value
        typeSave = sh.Cells(22, 13).value
        sh.Cells(22, 12).ClearContents
    End If
    'date
    arr = Split(ThisWorkbook.Sheets("Notes").Cells(6, 1).value, "|")
    datet = arr(1) & Right(arr(2), 2)
    'studia
    studia = ThisWorkbook.Sheets("Information").Cells(22, 9).value
    
    Dim pathToPOFolder As String:
        pathToPOFolder = GetPathPOFolder("POpdf")
    Select Case (typeSave)
        Case "Single"
            Call SingleSaving(datet, studia, pathToPOFolder)
        Case "Separately"
            Call SeparetelySaving(datet, studia, pathToPOFolder)
    End Select
    MsgBox "Saved"
End Sub

Private Sub SingleSaving(datet As String, studia As String, pathFolder)
    Dim rowsCount As Long: rowsCount = MyFunct.countRows("Prep list") - 4
    Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Prep list")
    'Dim rr As Range: Set rr = ThisWorkbook.Sheets("Result").Range("A" & count * 33 + 1 & ":K" & count * 33 + 33)
    Dim count As Long: count = CLng(ThisWorkbook.Sheets("Notes").Cells(8, 1).value)
    Dim rr As Range: Set rr = ThisWorkbook.Sheets("Result").Range("A1:K" & count)
    
    Dim ff: ff = "UX" & datet & "_" & studia & "_PO.pdf"
    Dim pdfPath: pdfPath = pathFolder & ff

    rr.ExportAsFixedFormat Type:=xlTypePDF, _
                               fileName:=pdfPath, _
                               Quality:=xlQualityStandard, _
                               IncludeDocProperties:=True, _
                               IgnorePrintAreas:=False
End Sub

Private Sub SeparetelySaving(datet As String, studia As String, pathFolder As String)
    Dim rowsCount As Long: rowsCount = CLng(ThisWorkbook.Sheets("Notes").Cells(8, 1).value) / 33 - 1
    Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Prep list")
    Dim count As Long ': count = CLng(ThisWorkbook.Sheets("Notes").Cells(8, 1).value) / 33
    Dim rr As Range
    Dim title As String, PO As String, ff As String, mypath As String
    Dim pdfPath As String
    Dim i As Long
    
    If Not IsFolderExist(pathFolder, studia) Then
        mypath = CreateFolder(pathFolder, studia)
    Else
        mypath = pathFolder & studia
    End If
    
     If Right(mypath, 1) <> "\" Then
        mypath = mypath & "\"
    End If
    
    For i = 0 To rowsCount
        count = i
        
        PO = ThisWorkbook.Sheets("Result").Cells(i * 33 + 1, 2)
        title = Trim(ThisWorkbook.Sheets("Result").Cells(i * 33 + 9, 2))
        title = MyFunct.RemoveSpecSymbols(title)
        
        ff = "UX" & datet & "_PO_" & PO & "_" & title & ".pdf"
        pdfPath = mypath & ff
        
        Set rr = ThisWorkbook.Sheets("Result") _
                             .Range("A" & count * 33 + 1 & ":K" & count * 33 + 33)
                             
        rr.ExportAsFixedFormat Type:=xlTypePDF, _
                               fileName:=pdfPath, _
                               Quality:=xlQualityStandard, _
                               IncludeDocProperties:=True, _
                               IgnorePrintAreas:=False
    Next i
End Sub

Private Function GetPathPOFolder(nameFolder As String) As String
    Dim pathToFile As String: pathToFile = ThisWorkbook.Path
    Dim pathToPOFolder As String
    
    pathToPOFolder = pathToFile & "\" & nameFolder 'POpdf"
    
    If Not IsFolderExist(pathToFile, "POpdf") Then GetPathPOFolder = CreateFolder(pathToFile, "POpdf")
    'Else: MasgBox "Existed"
    GetPathPOFolder = pathToFile & "\" & "POpdf" & "\"
End Function

Private Function IsFolderExist(pathWhere As String, nameFolder As String) As Boolean
    Dim pathToPOFolder: pathToPOFolder = pathWhere & "\" & nameFolder
    
    If Right(pathToPOFolder, 1) <> "\" Then
        pathToPOFolder = pathToPOFolder & "\"
    End If
    
    If Dir(pathToPOFolder, vbDirectory) <> vbNullString Then
        IsFolderExist = True
        Exit Function
    Else
        IsFolderExist = False
        Exit Function
    End If
End Function

Private Function CreateFolder(pathWhere As String, nameFolder As String) As String
    
    If Right(pathWhere, 1) <> "\" Then
        pathWhere = pathWhere & "\"
    End If
    
    Dim pathToPOFolder: pathToPOFolder = pathWhere & nameFolder
    
    Dim fdObj As Object
    Set fdObj = CreateObject("Scripting.FileSystemObject")
    
    fdObj.CreateFolder (pathToPOFolder)
    CreateFolder = pathToPOFolder
End Function
