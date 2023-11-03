Attribute VB_Name = "TheHub"
Option Explicit
Public Const sh As String = "The Hub"

Sub ClearTheHubList()
    Dim rows As Long
    rows = MyFunct.countRows(sh)
    If rows > 2 Then
        ThisWorkbook.Sheets(sh).Range("A3:V" & rows).Clear
    End If
End Sub
Sub TheHubFilling(MmYy As String)
    Call Main.FillingInLab(MmYy, sh)
End Sub
