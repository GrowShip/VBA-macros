Attribute VB_Name = "MyFunct"
Option Explicit

Function countRows(sheetname As String) As Long
    countRows = ThisWorkbook.Sheets(sheetname).UsedRange.rows(ThisWorkbook.Sheets(sheetname).UsedRange.rows.count).row
End Function

Function countRowBest(sheetname As String) As Long
    countRowBest = ThisWorkbook.Sheets(sheetname).Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
End Function

Function RemoveSpecSymbols(title As String) As String
    Dim charr As Variant
'RemoveSpecialSymbols
        Const SpecialCharacters As String = "<|>|:|""|/|\|?|*"
        
        For Each charr In Split(SpecialCharacters, "|")
            title = Replace(title, charr, "")
        Next
    RemoveSpecSymbols = title
End Function

Function GetAspect(fromSheeti As String, fromCol As Long, fromRow As Long, asWord As Boolean) As String
    If asWord Then
        If InStr(1, LCase(ThisWorkbook.Sheets(fromSheeti).Cells(fromRow, fromCol)), "theatrical") Then
            GetAspect = "Th"
        ElseIf InStr(1, LCase(ThisWorkbook.Sheets(fromSheeti).Cells(fromRow, fromCol)), "edited") Then
            GetAspect = "Ed"
        End If
    Else
        If InStr(1, LCase(ThisWorkbook.Sheets(fromSheeti).Cells(fromRow, fromCol)), "16") Then
            GetAspect = "16x9"
        ElseIf InStr(1, LCase(ThisWorkbook.Sheets(fromSheeti).Cells(fromRow, fromCol)), "4") Then
            GetAspect = "4x3"
        End If
    End If
End Function

Function GiveMeAspect(value As String, typeR As Long)
    Select Case typeR
        Case 0 'Нужно получить стринг формат
            If InStr(1, value, "thea") > 0 Then
                GiveMeAspect = "Th"
            ElseIf InStr(1, value, "edit") > 0 Then
                GiveMeAspect = "Ed"
            End If
        Case 1 'Нужно получить цифровое значение
            If InStr(1, value, "thea") > 0 Then
                GiveMeAspect = "16х9"
            ElseIf InStr(1, value, "edit") > 0 Then
                GiveMeAspect = "4х3"
            End If
    End Select
End Function

Function RemoveSymbols(yourString As String, Optional changeSpace As String) As String
Dim charr As Variant
Dim arr As Variant
    If IsEmpty(changeSpace) Then changeSpace = ""
    arr = Split(yourString, "|")
    If Len(arr(1)) = 0 Then
        yourString = arr(0)
    Else
        yourString = Replace(yourString, "|", " ")
    End If
    Const SpecialCharacters As String = "’| |!|?|:|,|'|.|-|–|…|@|#|$|%|^|&|*|(|)|{|[|]|/|\}"
        For Each charr In Split(SpecialCharacters, "|")
            yourString = Trim(Replace(yourString, charr, ""))
        Next
        RemoveSymbols = Replace(yourString, " ", changeSpace)
End Function

Function CheckStatus(row As Long, lab As String, Optional status As String) As String()
    'Проверка статусов для тайтла
    Dim i As Long
    Dim result(3) As String
    Dim sheet As String
    
    If Len(status) <= 0 Then status = "new"
    
    sheet = "Initial"
    
    For i = 0 To 2
        If InStr(1, ThisWorkbook.Sheets("Initial").Cells(row, 20).value, lab, vbTextCompare) > 0 Then
            If InStr(1, ThisWorkbook.Sheets(sheet).Cells(row, 16 + i).value, "new", vbTextCompare) > 0 Then
                result(i) = "new"
            End If
        End If
    Next i
    
    CheckStatus = result
End Function

Function GiveMeTitle(name As String) As String
    Dim number As Long
    number = InStr(1, name, "Season", vbTextCompare)
    
    If number > 0 Then
        GiveMeTitle = Left(name, number - 3)
    Else
        GiveMeTitle = name
    End If
End Function

Function GiveMeSeason(name As String) As String
    Dim number As Long
    number = InStr(1, name, "Season", vbTextCompare)
    
    If number > 0 Then
        GiveMeSeason = Right(name, 1)
    Else
        GiveMeSeason = ""
    End If
End Function

Function GiveMeDubs(row As Long) As String
    Dim cell As Variant
    Dim result As String
    For Each cell In ThisWorkbook.Sheets("Filenames").Range("S" & row & ":AJ" & row)
        If Len(cell) > 0 Then
            If Len(result) = 0 Then
                result = cell
            ElseIf InStr(1, cell, "ad", vbTextCompare) > 0 Then
                result = result & "/" & "Dvs"
            Else
                result = result & "/" & cell
            End If
        'Else: Exit For
        End If
    Next cell
    GiveMeDubs = result
End Function

Sub MakeFormatRow(sh As String, actRow As Long)
    With ThisWorkbook.Sheets(sh).Range("A" & actRow & ":AI" & actRow)
        .Font.Size = "9"
        .EntireRow.RowHeight = 25
    End With
End Sub

Function GetDubForWow(fileName As String) As String

    Const SpecialCharactersChi As String = "yu|cm"
    Dim charr As Variant
    'Китайский
    For Each charr In Split(SpecialCharactersChi, "|")
        fileName = Replace(fileName, charr, "Zh", , , vbTextCompare)
    Next
    
    'Казахский
        fileName = Replace(fileName, "Ka", "Kk", , , vbTextCompare)
        
    'Испанский
        fileName = Replace(fileName, "Sp", "Es", , , vbTextCompare)
        
    'Турецкий
        fileName = Replace(fileName, "Tu", "Tr", , , vbTextCompare)
        
    GetDubForWow = fileName
End Function

Function GetDVSinFilename(row As Long) As String
    If InStr(1, ThisWorkbook.Sheets("Filenames").Cells(row, 35), "AD") > 0 Then
        GetDVSinFilename = "Dvs"
    Else: GetDVSinFilename = ""
    End If
End Function

