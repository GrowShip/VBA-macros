Attribute VB_Name = "FillingCTRrave"
Option Explicit
Dim lastTittle As String
Dim ordinalEpison As Integer
Public wb As Workbook
Public initSh As Worksheet
Public filenameSh As Worksheet
Public CTRsheet As Worksheet

Sub FillInCTRRave()
    Set wb = ThisWorkbook
    Set initSh = wb.Sheets("Initial")
    Set filenameSh = wb.Sheets("Filenames")
    Set CTRsheet = wb.Sheets("CTRlock")
     
    Dim countRow As Long, countRowCtr As Long
    Dim nameCtr As String, nameInit As String
    Dim i As Long, j As Long


    countRow = MyFunc.FindLastRow(wb.name, filenameSh.name, "A")
    countRowCtr = MyFunc.FindLastRow(wb.name, CTRsheet.name, "A")

    For j = 2 To countRowCtr + 1
        nameCtr = NameInCtr(CTRsheet, j, "A", "B", "C")
        For i = 3 To countRow
            nameInit = NameInCtr(initSh, i, "L", "O")
        
            If InStr(1, initSh.Range("P" & i), "new", vbTextCompare) > 0 And j = countRowCtr + 1 Then
                Call FillTheGaps(wb.name, CTRsheet.name, i, j, initSh.name, "draft", True)
            ElseIf StrComp(Split(nameInit, "|")(0), Split(nameCtr, "|")(0), vbTextCompare) = 0 And _
                    StrComp(Split(nameInit, "|")(1), Split(nameCtr, "|")(1), vbTextCompare) = 0 And _
                    j <> countRowCtr + 1 Then
                Call FillTheGaps(wb.name, CTRsheet.name, i, j, initSh.name, "draft", False)
                Exit For
            ElseIf i = countRow - 1 And j < countRowCtr + 1 Then
                filenameSh.Range("I" & j).Interior.Color = vbBlue
            End If
        Next i
    Next j

    MsgBox "Done"

End Sub

Function NameInCtr(sh As Worksheet, row As Long, tit As String, Optional epis As String, Optional seas As String) As String
    Dim title As String
    Dim episode As String
    
    title = sh.Range(tit & row).value ' Get the title
    
    episode = sh.Range(epis & row).value ' Get the episode
    
    ' If the title contains "Season" then remove it
    If InStr(1, title, "Season") Then
        title = Split(title, ".")(0)
    End If
    
    ' Append the episode number to the title
    title = title & "| " & episode
    
    NameInCtr = title ' Set the function return value
End Function

Sub FillTheGaps(wb As String, ctrSh As String, row As Long, ByVal ctrRow As Long, initSheet As String, draftSh As String, Optional flagS As Boolean)

        '—начала проверка по статусу чтобы убрать из списка все DELETE
        
        'HoldOver
        Dim initCol As String
        Dim raveCol As String
        
        ' ака€ сейчас колонка дл€ рейва в смете
        initCol = "P" '16 Cells(row, 16)
        
        ' ака€ сейчас колонка дл€ рейвав  стр
        raveCol = "I" '9
        
        If Len(initSh.Cells(row, 16).value) < 1 Or InStr(1, initSh.Cells(row, 16).value, "delete", vbTextCompare) Then
            Workbooks(wb).Sheets(ctrSh).Range(raveCol & ctrRow).value = "Empty/Delete"
            Workbooks(wb).Sheets(ctrSh).Range(raveCol & ctrRow).Font.Color = vbRed
            GoTo itWasEmptyOrDelete
        ElseIf InStr(1, initSh.Cells(row, 16).value, "remove", vbTextCompare) Then
            Workbooks(wb).Sheets(ctrSh).Range(raveCol & ctrRow).value = "Delete"
            'Workbooks(wb).Sheets(CTRsh).Range(raveCol & ctrRow).Font.Color = vbRed
            Workbooks(wb).Sheets(ctrSh).Range(raveCol & ctrRow).Font.Bold = True
            GoTo itWasDelete
        ElseIf InStr(1, initSh.Cells(row, 16).value, "new", vbTextCompare) Then
            If flagS Then
                ctrRow = MyFunc.FindLastRow(wb, ctrSh, "A") + 1
            End If
            Workbooks(wb).Sheets(ctrSh).Range(raveCol & ctrRow).value = "New"
            Workbooks(wb).Sheets(ctrSh).Range(raveCol & ctrRow).Font.Color = vbBlack
           
        ElseIf InStr(1, initSh.Cells(row, 16).value, "old", vbTextCompare) Then
            Workbooks(wb).Sheets(ctrSh).Range(raveCol & ctrRow).value = "Holdover"
            Workbooks(wb).Sheets(ctrSh).Range(raveCol & ctrRow).Font.Color = vbBlack
            GoTo itWasEmptyOrDelete
        End If
        'Workbooks(wb).Sheets(CTRsh).Range("I" & CTRRow) = initSh.Range("O" & initRow)
        
        Dim datet As String
        Dim title As String
        Dim titEngCol As String
        Dim MorTVcol As String
        
        '—егодн€шн€€ дата мас€ц√од
        datet = ThisWorkbook.Sheets(ctrSh).Cells(2, 4).value
        titEngCol = "L" '12 Cells(row, 12).Value
        MorTVcol = "A"  '1 Cells(ctrRow, 1).Value
        
        If LCase(initSh.Cells(row, 1).value) Like "tv" And _
            Len(initSh.Cells(row, 12).value) > 0 And _
            InStr(1, LCase(initSh.Cells(row, 2).value), "document") = 0 Then
            
            'Tittle(eng)
            If UBound(Split(initSh.Cells(row, 12).value, ". ")) > 0 Then
                title = initSh.Cells(row, 12).value
                Workbooks(wb).Sheets(ctrSh).Range("A" & ctrRow).value = Split(initSh.Cells(row, 12).value, ". ")(0)
                
            'Season
                Workbooks(wb).Sheets(ctrSh).Range("C" & ctrRow).value = Split(initSh.Cells(row, 12).value, ". ")(1)
            Else
                title = initSh.Cells(row, 12).value
                Workbooks(wb).Sheets(ctrSh).Range("A" & ctrRow).value = initSh.Cells(row, 12).value
            End If
            
            'EpisodeTittle
                Workbooks(wb).Sheets(ctrSh).Range("B" & ctrRow).value = initSh.Range("O" & row).value
                
            'Episode
                Call EpisodeNumbering(title, wb, initSh.name, ctrSh, titEngCol, row, ctrRow)
                
        ElseIf Len(initSh.Cells(row, 12).value) > 1 Then
        
        'Tittle(eng)
            lastTittle = ""
            title = initSh.Cells(row, 12).value
            Workbooks(wb).Sheets(ctrSh).Range(MorTVcol & ctrRow).value = initSh.Cells(row, 12).value
        Else
            Workbooks(wb).Sheets(ctrSh).Range(MorTVcol & ctrRow).value = "Empty"
            Workbooks(wb).Sheets(ctrSh).Range(MorTVcol & ctrRow).Font.Color = vbRed
            GoTo itWasEmptyOrDelete
        End If
        
    'RemoveSpecialSymbols
        Const SpecialCharacters As String = "Т|!|?|:|,|'| |.|-|Е|^|&|*|(|)"
        Dim charr As Variant
    
        For Each charr In Split(SpecialCharacters, "|")
            title = Replace(title, charr, "")
        Next
        
        
        With Workbooks(wb).Sheets(ctrSh)
            .Range("E" & ctrRow).value = "No" ' Priority
            .Range("F" & ctrRow).value = initSh.Range("A" & row).value ' Category
            .Range("G" & ctrRow).value = initSh.Range("AH" & row).value  ' RunTime
        End With
        
    'VersionEdTh
        Dim versionEdTh As String
        Dim initCell As Range
        Set initCell = initSh.Range("AA" & row)
        If InStr(1, LCase(initCell.value), "theatrical") Then
            versionEdTh = "Th"
        ElseIf InStr(1, LCase(initCell.value), "edited") Then
            versionEdTh = "Ed"
        End If
        Workbooks(wb).Sheets(ctrSh).Range("H" & ctrRow).value = versionEdTh
        
    'DateSt
        Dim stDate As Date
        stDate = initSh.Range("AC" & row).value
        With Workbooks(wb).Sheets(ctrSh).Range("J" & ctrRow)
            .value = stDate
            .NumberFormat = "mm.dd.yy"
        End With
        
    'DateEnd
        Dim lastDay As Date
        lastDay = DateSerial(Year(Date), 12, 31)
        With Workbooks(wb).Sheets(ctrSh).Range("K" & ctrRow)
            .value = lastDay
            .NumberFormat = "mm.dd.yy"
        End With
        
    'DubLang
        Dim k As Integer
        Dim flag As Boolean
        
        'метка что dvs уже добавлен
            flag = True
        
        'Copy language information from draft sheet to CTR sheet
        For k = 0 To 10
            If Len(Workbooks(wb).Sheets(draftSh).Range(Chr(67 + k) & row).value) > 2 Then
                 Workbooks(wb).Sheets(ctrSh).Range(Chr(76 + k) & ctrRow) = Workbooks(wb).Sheets(draftSh).Range(Chr(67 + k) & row)
            ElseIf Len(Workbooks(wb).Sheets(draftSh).Range("S" & row).value) > 2 And flag Then
                Workbooks(wb).Sheets(ctrSh).Range(Chr(76 + k) & ctrRow) = "Dvs"
                flag = False
                Exit For
            End If
        Next k
        
        'Copy audio dynamics information from draft sheet to CTR sheet
        For k = 0 To 6
            If Len(Workbooks(wb).Sheets(draftSh).Cells(row, 13 + k).value) > 1 Then
                Call MakeNameSubDin(wb, ctrSh, row, ctrRow, datet, title, Workbooks(wb).Sheets(draftSh).Cells(row, 13 + k).value, 22 + k)
                'Workbooks(wb).Sheets(CTRsh).Cells(CTRRow, 22 + k) = Workbooks(wb).Sheets(draftSH).Cells(initRow, 23 + k)
            ElseIf Len(Workbooks(wb).Sheets(draftSh).Cells(row, 20).value) > 1 Then
                Call MakeNameSubDin(wb, ctrSh, row, ctrRow, datet, title, Workbooks(wb).Sheets(draftSh).Cells(row, 20).value, 22 + k)
                Exit For
            End If
        Next k
        
        'Copy burned audio information from draft sheet to CTR sheet
        For k = 0 To 2
            If Len(Workbooks(wb).Sheets(draftSh).Cells(row, 21 + k).value) > 1 Then
                Workbooks(wb).Sheets(ctrSh).Cells(ctrRow, 20 + k) = Left(Workbooks(wb).Sheets(draftSh).Cells(row, 21 + k), Len(Workbooks(wb).Sheets(draftSh).Cells(row, 21 + k)) - 1)
            End If
        Next k
    
    ' Set Parent Title
    Dim ThEd As String
    ThEd = ""
    If Not IsEmpty(Workbooks(wb).Sheets(ctrSh).Range("H" & ctrRow)) Then
        ThEd = Workbooks(wb).Sheets(ctrSh).Range("H" & ctrRow) & "_"
    End If
    If Len(Workbooks(wb).Sheets(ctrSh).Range("T" & ctrRow).value) > 0 Then
        Dim burned As String
        burned = Workbooks(wb).Sheets(ctrSh).Range("T" & ctrRow).value & "S"
            If Len(Workbooks(wb).Sheets(ctrSh).Range("U" & ctrRow).value) > 0 Then
                burned = burned & Workbooks(wb).Sheets(ctrSh).Range("U" & ctrRow).value & "S"
            End If
    End If
    title = filenameSh.Cells(row, 5).value
    Workbooks(wb).Sheets(ctrSh).Range("AD" & ctrRow).value = "KZR_" & datet & "_" & title & "_" & ThEd & _
    Workbooks(wb).Sheets(ctrSh).Range("L" & ctrRow).value & _
    Workbooks(wb).Sheets(ctrSh).Range("M" & ctrRow).value & _
    Workbooks(wb).Sheets(ctrSh).Range("N" & ctrRow).value & _
    Workbooks(wb).Sheets(ctrSh).Range("O" & ctrRow).value & _
    Workbooks(wb).Sheets(ctrSh).Range("P" & ctrRow).value & _
    Workbooks(wb).Sheets(ctrSh).Range("Q" & ctrRow).value & _
    Workbooks(wb).Sheets(ctrSh).Range("R" & ctrRow).value & _
    Workbooks(wb).Sheets(ctrSh).Range("S" & ctrRow).value & burned & ".mp4"
    
    ' Set Aspect
    Dim aspect As String
    If InStr(1, LCase(initSh.Range("AA" & row).value), "16") Then
        aspect = "16x9"
    ElseIf InStr(1, LCase(initSh.Range("AA" & row).value), "4") Then
        aspect = "4x3"
    End If
    
    With Workbooks(wb).Sheets(ctrSh)
        .Range("AE" & ctrRow).value = aspect ' Set Aspect
        .Range("AF" & ctrRow).value = "480p" ' Set Resolution
        .Range("AJ" & ctrRow).value = initSh.Range("S" & row).value ' Set Studio
        .Range("AK" & ctrRow).value = initSh.Range("T" & row).value ' Set Lab
    End With
    
itWasDelete:
    If InStr(1, LCase(initSh.Cells(row, 16)), "remove") Then
        Dim lastRowDelete As Long
        Dim CTRdeleteSH As String
        
        CTRdeleteSH = "RAVE_CTR_Remove"
        lastRowDelete = MyFunc.FindLastRow(wb, CTRdeleteSH, "A") + 1
        
        ' Transfer data to RAVE_CTR_Remove sheet
        With Workbooks(wb).Sheets(CTRdeleteSH)
            .Range("A" & lastRowDelete).value = Workbooks(wb).Sheets(ctrSh).Range("A" & ctrRow).value ' Title
            .Range("B" & lastRowDelete).value = Workbooks(wb).Sheets(ctrSh).Range("B" & ctrRow).value ' EpisodeTitle
            .Range("C" & lastRowDelete).value = initSh.Range("A" & row).value ' Category
            .Range("D" & lastRowDelete).value = Workbooks(wb).Sheets(ctrSh).Range("AD" & ctrRow).value ' ParentTitle
        End With
    End If

itWasEmptyOrDelete:
End Sub

Sub MakeNameSubDin(wb As String, ctrSh As String, row As Long, ctrRow As Long, datet As String, title As String, lang As String, CTRcol As Integer)
    Dim fileName As String
    fileName = "KZR_" & datet & "_" & title & "_" & Replace(lang, " -DYN Sub", "") & "_" & Workbooks(wb).Sheets(ctrSh).Range("H" & ctrRow) & ".srt"
    Workbooks(wb).Sheets(ctrSh).Cells(ctrRow, CTRcol).value = fileName
End Sub

Sub EpisodeNumbering(title As String, wb As String, initSheet As String, CTRsheet As String, titEngCol As String, row As Long, ctrRow As Long)
    'lastTitle
    'ordinalEpison
    
    If lastTittle = title Then
        ordinalEpison = ordinalEpison + 1
    Else
        ordinalEpison = 1
        lastTittle = title
    End If
    
    Workbooks(wb).Sheets(CTRsheet).Range("D" & ctrRow).value = ordinalEpison
    
    Dim ctrValue As String
    ctrValue = LCase(Trim(Workbooks(wb).Sheets(CTRsheet).Range("C" & ctrRow).value))
    If InStr(ctrValue, "season") = 1 Then
        Workbooks(wb).Sheets(CTRsheet).Range("C" & ctrRow).value = Trim(Right(ctrValue, Len(ctrValue) - Len("season")))
    End If
End Sub

