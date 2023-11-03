Attribute VB_Name = "clearing"
Option Explicit

Private Sub ClearSheet(sheetname As String, rowAfter As Long, columnUntil As Long, Optional ask As Boolean = True)
    If ask Then If InitialFilling.AskText("Точно очистить лист " & sheetname & " ?") Then Exit Sub
        
    If ask Then frmLoad.Loading
    If ask Then ЭтаКнига.Opened
    
    Dim row As Long
    row = MyFunct.countRows(sheetname)
    If (row <= rowAfter) Then GoTo Continue
    'ThisWorkbook.Sheets(sheetName).Range(Cells(rowAfter, 1), Cells(row, columnUntil)).clear
    ThisWorkbook.Sheets(sheetname).rows(rowAfter + 1 & ":" & row).Clear
    'MsgBox "Cleared"
Continue:
    If ask Then ЭтаКнига.Closed
    If ask Then frmLoad.Unloading
    
    ThisWorkbook.Sheets("Information").Activate
End Sub

Public Sub ClearAllTubs()
    If InitialFilling.AskText("Точно очистить ВСЕ листы?") Then Exit Sub
    Call ClearSheet("CTRlock", 1, 40, False)
    Call ClearSheet("CTRupload", 1, 37, False)
    Call ClearSheet("RemoveLock", 1, 4, False)
    Call ClearSheet("RemoveUpload", 1, 4, False)
    Call ClearSheet("REMIXlock", 1, 50, False)
    Call ClearSheet("REMIXupload", 1, 50, False)
    Call ClearSheet("GuiREMIXlock", 1, 14, False)
    Call ClearSheet("GuiREMIXupload", 1, 14, False)
End Sub
    
Public Sub ClearCTRlock(Optional ask As Boolean = True)
    Call ClearSheet("CTRlock", 1, 40, ask)
End Sub

Public Sub ClearCTRupload(Optional ask As Boolean = True)
    Call ClearSheet("CTRupload", 1, 37, ask)
End Sub

Public Sub ClearRemoveLock(Optional ask As Boolean = True)
    Call ClearSheet("RemoveLock", 1, 4, ask)
End Sub

Public Sub ClearRemoveUpload(Optional ask As Boolean = True)
    Call ClearSheet("RemoveUpload", 1, 4, ask)
End Sub

Public Sub ClearREMIXlock(Optional ask As Boolean = True)
    Call ClearSheet("REMIXlock", 1, 50, ask)
End Sub

Public Sub ClearREMIXupload(Optional ask As Boolean = True)
    Call ClearSheet("REMIXupload", 1, 50, ask)
End Sub

Public Sub ClearGuiREMIXlock(Optional ask As Boolean = True)
    Call ClearSheet("GuiREMIXlock", 1, 14, ask)
End Sub

Public Sub ClearGuiREMIXupload(Optional ask As Boolean = True)
    Call ClearSheet("GuiREMIXupload", 1, 14, ask)
End Sub

Public Sub ClearOrderGoogle(Optional ask As Boolean = True)
    Call ClearSheet("OrderGoogle", 1, 64, ask)
End Sub
