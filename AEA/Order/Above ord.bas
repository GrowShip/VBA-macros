Attribute VB_Name = "Above"
Option Explicit
Public Const sh As String = "Above"
Public loopi As Long
Public actRow As Long

Sub ClearAbovelist()
    Dim rows As Long
    rows = MyFunct.countRows(sh)
    If rows > 2 Then
        ThisWorkbook.Sheets(sh).Range("A3:AI" & rows).Clear
    End If
End Sub

Sub AboveFilling(MmYy As String)
    Call Main.FillingInLab(MmYy, sh)
End Sub

