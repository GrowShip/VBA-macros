Attribute VB_Name = "LabAero"
Option Explicit
Public Const sh As String = "Lab.Aero"

Sub ClearLabAeroList()
    Dim rows As Long
    rows = MyFunct.countRows(sh)
    If rows > 2 Then
        ThisWorkbook.Sheets(sh).Range("A3:U" & rows).Clear
    End If
End Sub

Sub LabAeroFilling(MmYy As String)
    Call Main.FillingInLab(MmYy, sh)
End Sub
