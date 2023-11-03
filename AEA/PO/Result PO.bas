Attribute VB_Name = "Result"
Option Explicit

Sub ClearResultList()
    Dim response
    response = MsgBox("Точно очистить лист Result?", vbOKCancel)
    If response = vbOK Then
        Dim row As Long
        row = MyFunct.countRows("Result")
        ThisWorkbook.Sheets("Result").Pictures.Delete
        ThisWorkbook.Sheets("Result").Cells.Clear
        MsgBox "Cleared"
    End If
End Sub

Sub MakingResult()
    Call ЭтаКнига.Opened
    
    Call ShowForm
    If flagform = False Then Exit Sub
    ThisWorkbook.Sheets("Notes").Cells(6, 1).value = ddd & "|" & mmm & "|" & yyy
    
    Call TransmitInfo
    MsgBox "POs Ready"
    
    Call ЭтаКнига.Closed
End Sub
 
 Private Sub ShowForm()
    POform.Caption = "Дата в название файла"
    POform.textInfo.Caption = "Введи " & Chr(10) & _
                              "день (2 симв)," & Chr(10) & _
                              "месяц (2 симв)," & Chr(10) & _
                              "год(4 симв)" & Chr(10) & _
                              "для имени файла PO"
    POform.Show
    'ddd = POform.Dd
    'mmm = POform.Mm
    'yyy = POform.YYYY
    
End Sub

Private Sub CopyPasteSample(numberRow As Long)
    ThisWorkbook.Sheets("Sample").Range("A1:K33").Copy _
        Destination:=ThisWorkbook.Sheets("Result").Range("A" & numberRow)
End Sub

Private Sub TransmitInfo()
    Dim sh As Worksheet
    Dim i As Long, row As Long, resultRow As Long, j As Long, serialRow As Long, serialCol As Long
    Dim positionWB As Long: positionWB = 0
    Dim delta As Long: delta = 0
    Dim lastTitle As String: lastTitle = ""
    Dim arr, datet, dates, dateDel
    
    arr = Split(ThisWorkbook.Sheets("Notes").Cells(5, 1).value, "|")
    
    Set sh = ThisWorkbook.Sheets("Prep list")
    row = MyFunct.countRows(sh.name)
    dates = arr(0) & "." & arr(1) & "." & arr(2)
    datet = DateValue(dates)
    dateDel = DateAdd("d", 20, datet)
    
    For i = 4 To row
        Dim colum As Long: colum = 0
        Dim flagWB As Boolean
        
        'TV сериал и студия warner
        If (InStr(1, sh.Cells(i, 3).value, "Warner", vbTextCompare) > 0 Or _
            InStr(1, sh.Cells(i, 3).value, "HBO", vbTextCompare) > 0) And _
            InStr(1, sh.Cells(i, 26).value, "TV", vbTextCompare) > 0 Then
            If StrComp(Split(sh.Cells(i, 2).value, "|")(0), lastTitle, vbTextCompare) = 0 Then
                positionWB = (i - 4) * 33
                delta = positionWB - resultRow
                flagWB = True
            Else
                If flagWB Then
                    serialRow = 0
                    serialCol = 0
                    flagWB = False
                End If
                
                resultRow = (i - 4) * 33 - delta
                positionWB = resultRow
                lastTitle = Split(sh.Cells(i, 2).value, "|")(0)
            End If
            'Series
            If serialRow = 15 Then
                serialRow = 0
                serialCol = 3
            End If
            If serialRow = 1 And serialCol = 0 Then ThisWorkbook.Sheets("Result").Cells(resultRow + 1, 7 + serialCol).value = Split(sh.Cells(i - 1, 2).value, "|")(1)
            If serialRow < 15 And serialCol <= 3 Then ThisWorkbook.Sheets("Result").Cells(resultRow + 1 + serialRow, 7 + serialCol).value = Split(sh.Cells(i, 2).value, "|")(1)
            serialRow = serialRow + 1
        Else
            If flagWB Then
                positionWB = (i - 5) * 33
                delta = positionWB - resultRow
                
                serialRow = 0
                serialCol = 0
            Else
                
            End If
            
            flagWB = False
            resultRow = (i - 4) * 33 - delta
        End If
        
        
        If flagWB Then
            ThisWorkbook.Sheets("Result").Cells(resultRow + 9, 2).value = lastTitle
        Else
            Call CopyPasteSample(resultRow + 1)
            With ThisWorkbook.Sheets("Result")
                .Cells(resultRow + 1, 2).value = sh.Cells(i, 1).value 'po
                .Cells(resultRow + 2, 2).value = datet 'date
                .Cells(resultRow + 4, 2).value = sh.Cells(i, 3).value 'Distr
                .Cells(resultRow + 5, 2).value = sh.Cells(i, 4).value 'Lab
                .Cells(resultRow + 9, 2).value = Replace(sh.Cells(i, 2).value, "|", " ") 'Title
                .Cells(resultRow + 10, 2).value = sh.Cells(i, 7).value 'Year
                .Cells(resultRow + 12, 2).value = sh.Cells(i, 5).value 'play period
            End With
            
            For j = 0 To 2
                If Len(sh.Cells(i, 19 + j).value) > 1 Then
                    If colum = 0 Then
                        ThisWorkbook.Sheets("Result").Range("A" & resultRow + 22 & ":C" & resultRow + 28).value = ThisWorkbook.Sheets("Notes").Cells(3, 3 + j).value
                    ElseIf colum = 1 Then
                        ThisWorkbook.Sheets("Result").Range("D" & resultRow + 22 & ":G" & resultRow + 28).value = ThisWorkbook.Sheets("Notes").Cells(3, 3 + j).value
                    ElseIf colum = 2 Then
                        ThisWorkbook.Sheets("Result").Range("H" & resultRow + 22 & ":K" & resultRow + 28).value = ThisWorkbook.Sheets("Notes").Cells(3, 3 + j).value
                    End If
                    With ThisWorkbook.Sheets("Result")
                        .Range("A" & resultRow + 17 + colum & ":K" & resultRow + 17 + colum).Interior.Color = RGB(247, 249, 241)
                        .Range("A" & resultRow + 17 + colum & ":K" & resultRow + 17 + colum).Borders.Color = RGB(221, 217, 196)
                        .Range("A" & resultRow + 17 + colum).Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
                        .Range("K" & resultRow + 17 + colum).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
                        .Cells(resultRow + 17 + colum, 1).value = sh.Cells(2, 19 + j).value ' system
                        .Cells(resultRow + 17 + colum, 2).value = sh.Cells(3, 19 + j).value ' fformat
                        .Cells(resultRow + 17 + colum, 3).value = sh.Cells(i, 9).value ' dub1
                        .Cells(resultRow + 17 + colum, 4).value = sh.Cells(i, 10).value ' dub2
                        .Cells(resultRow + 17 + colum, 5).value = sh.Cells(i, 11).value ' dub3
                        .Cells(resultRow + 17 + colum, 6).value = sh.Cells(i, 12).value ' dub4
                        .Cells(resultRow + 17 + colum, 7).value = sh.Cells(i, 13).value ' dub5
                        .Cells(resultRow + 17 + colum, 8).value = sh.Cells(i, 14).value ' dub6
                        .Cells(resultRow + 17 + colum, 9).value = Left(sh.Cells(i, 15).value, 7) ' sub1
                        .Cells(resultRow + 17 + colum, 10).value = Left(sh.Cells(i, 16).value, 7) ' sub2
                        .Cells(resultRow + 17 + colum, 11).value = sh.Cells(i, 8).value ' aspect
                    End With
                    
                    If colum = 0 Then
                        With ThisWorkbook.Sheets("Result")
                            .Cells(resultRow + 13, 2 + colum).value = sh.Cells(2, 19 + j).value ' system
                            .Cells(resultRow + 14, 2 + colum).value = dateDel
                            .Cells(resultRow + 15, 2 + colum).value = sh.Cells(1, 19 + j).value ' amount of flight
                        End With
                        colum = colum + 1
                    ElseIf colum = 1 Then
                        With ThisWorkbook.Sheets("Result")
                            .Range("C" & resultRow + 13 & ":D" & resultRow + 13).value = sh.Cells(2, 19 + j).value ' system
                            .Range("C" & resultRow + 14 & ":D" & resultRow + 14).value = dateDel
                            .Range("C" & resultRow + 15 & ":D" & resultRow + 15).value = sh.Cells(1, 19 + j).value ' amount of flight
                        End With
                        colum = colum + 1
                    ElseIf colum = 2 Then
                        With ThisWorkbook.Sheets("Result")
                            .Range("E" & resultRow + 13 & ":F" & resultRow + 13).value = sh.Cells(2, 19 + j).value ' system
                            .Range("E" & resultRow + 14 & ":F" & resultRow + 14).value = dateDel
                            .Range("E" & resultRow + 15 & ":F" & resultRow + 15).value = sh.Cells(1, 19 + j).value ' amount of flight
                        End With
                        colum = colum + 1
                    End If
                End If
            Next j
        End If
               
    Next i
    ThisWorkbook.Sheets("Notes").Cells(8, 1).value = resultRow + 33
    ThisWorkbook.Sheets("Result").PageSetup.PrintArea = "A1:K" & resultRow + 33
End Sub
