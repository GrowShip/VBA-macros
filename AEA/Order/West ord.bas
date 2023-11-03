Attribute VB_Name = "West"
Option Explicit
Public Const sh = "West"
Sub ClearWestList()
    Dim rows As Long
    rows = MyFunct.countRows(sh)
    If rows > 2 Then
        ThisWorkbook.Sheets(sh).Range("A3:EY" & rows).Clear
    End If
End Sub

Sub WestFilling(MmYy As String)
    Call Main.FillingInLab(MmYy, sh)
End Sub
