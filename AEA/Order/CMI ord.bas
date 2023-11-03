Attribute VB_Name = "CMI"
Option Explicit
Public Const sh As String = "CMI"

Sub ClearCMIList()
    Dim rows As Long
    rows = MyFunct.countRows(sh)
    If rows > 2 Then
        ThisWorkbook.Sheets(sh).Range("A3:X" & rows).Clear
    End If
End Sub

Sub CMIFilling(MmYy As String)
    Call Main.FillingInLab(MmYy, sh)
End Sub

