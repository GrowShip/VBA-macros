Option Explicit

Sub Opened()
    Call Unlocking
End Sub

Sub Closed()
    Call LockSomeFields
End Sub

Private Sub Workbook_Open()
    Call CreateStudiaList
    Call CreateSaveList
    Call LockSomeFields
End Sub

Private Sub CreateStudiaList()
    ThisWorkbook.Sheets("Information").ComboStudia.value = Null
    ThisWorkbook.Sheets("Information").Range("H22:I22").ClearContents
    With ThisWorkbook.Sheets("Information").ComboStudia
        .AddItem "HBO"
        .AddItem "NBC Universal"
        .AddItem "Paramount"
        .AddItem "Sony Pictures"
        .AddItem "Walt Disney"
        .AddItem "Warner Bros"
        .AddItem "Other"
        .AddItem "All"
    End With
    
End Sub

Private Sub CreateSaveList()
    ThisWorkbook.Sheets("Information").ComboSave.value = Null
    ThisWorkbook.Sheets("Information").Range("L22:M22").ClearContents
    With ThisWorkbook.Sheets("Information").ComboSave
        .AddItem "Single"
        .AddItem "Separately"
    End With
End Sub

Private Sub LockSomeFields()
    'Dim sh As Worksheet
    'Set sh = ThisWorkbook.Sheets("Information")
    ThisWorkbook.Sheets("Information").Protect Password:="12345", UserInterfaceOnly:=True '12345
    ThisWorkbook.Sheets("Initial").Protect Password:="12345", UserInterfaceOnly:=True
    ThisWorkbook.Sheets("Filenames").Protect Password:="12345", UserInterfaceOnly:=True
    ThisWorkbook.Sheets("Notes").Protect Password:="12345", UserInterfaceOnly:=True
    ThisWorkbook.Sheets("Result").Protect Password:="12345", UserInterfaceOnly:=True
End Sub

Private Sub Unlocking()
    With ThisWorkbook
        .Sheets("Information").Unprotect "12345"
        .Sheets("Initial").Unprotect "12345"
        .Sheets("Filenames").Unprotect "12345"
        .Sheets("Notes").Unprotect "12345"
        .Sheets("Result").Unprotect "12345"
    End With
End Sub

Private Sub HideMenu()
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"
    Application.DisplayFormulaBar = False
End Sub

Private Sub ShowMenu()
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",True)"
    Application.DisplayFormulaBar = True
End Sub
