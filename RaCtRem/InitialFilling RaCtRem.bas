Attribute VB_Name = "InitialFilling"
Option Explicit
Public ddd As Variant
Public mmm As Variant
Public yyy As Variant

Public Function AskText(text As String) As Boolean
    Dim response
    response = MsgBox(text, vbOKCancel)
    If response = vbOK Then
        AskText = False
    Else:
        AskText = True
    End If
End Function

Sub ClearKClist()

    If AskText("“очно очистить лист KC?") Then Exit Sub
        
    Call frmLoad.Loading
    Dim row As Long
    row = MyFunct.countRows("KC")
    ThisWorkbook.Sheets("KC").Cells.Clear
    
    Call frmLoad.Unloading
End Sub

Sub InitialListFilling()
        
    If AskText(" опируем смету из KC?") Then Exit Sub
    
    frmLoad.Loading
    Ёта нига.Opened
    
    'clear status list
    ThisWorkbook.Sheets("NOTES").Range("G1:G4").ClearContents
    
    Dim rows As Long, initRows As Long
    Dim init As String, kc As String
    kc = "KC"
    init = "Initial"
    
    rows = MyFunct.countRows(kc)
    initRows = MyFunct.countRows(init)
    
    If initRows > 2 Then ThisWorkbook.Sheets(init).Range("A3:BA" & initRows).Clear
    
    ThisWorkbook.Sheets(kc).Range("A4:BA" & rows).Copy _
        Destination:=ThisWorkbook.Sheets(init).Range("A3")
        
    initRows = MyFunct.countRows(init)
    
    Ёта нига.Closed
    frmLoad.Unloading
End Sub

Sub FilenamesFilling()
    
    If AskText("ѕодготавливаем скопированную смету?") Then Exit Sub
    
    VBAProject.Ёта нига.Opened
    
    Dim flagform As Boolean: flagform = False
    
    flagform = dateForSmtFrom.Loading("ƒата дл€ filenames")
    If Not flagform Then
        VBAProject.Ёта нига.Closed
        Exit Sub
    End If
    
    ThisWorkbook.Sheets("Notes").Cells(6, 1).value = ddd & "|" & mmm & "|" & yyy
    
    Call frmLoad.Loading
    
    With ThisWorkbook.Sheets("NOTES").Range("G1:G4")
        .ClearContents
        .NumberFormat = "@"
    End With
    
    Dim ArrLanInitial() As String
    Dim sh As String
    Dim shFn As String
    Dim SHdr As String
    Dim countRows As Long
    Dim result As String
    Dim title As String
    Dim dubs As String
    Dim i As Long
    Dim j As Long
    
    'Dubs all
    title = "A"
    
    dubs = "K"
    
    sh = "Initial"
    shFn = "Filenames"
    SHdr = "Draft"
    
    'Ќайдем сколько всего €чеек
    countRows = MyFunct.countRowBest(shFn)
    If countRows > 2 Then ThisWorkbook.Sheets(shFn).Range("A3:BA" & countRows).Clear
    countRows = MyFunct.countRowBest(sh)
    
    'countRows = ThisWorkbook.Sheets("Initial").UsedRange.rows(ThisWorkbook.Sheets("Initial").UsedRange.rows.Count).row
    'i = 3
    'j = 3
    
    For i = 3 To countRows
        '—оздаем массив из €зыков в €чейка
        ArrLanInitial = Split(Replace(ThisWorkbook.Sheets(sh).Range("V" & i), " ", ""), "/")
        
        '”казать колонку свойств субтитров
        result = MakeLangStr(ArrLanInitial, sh, "X", i)
        
        With ThisWorkbook
            'title
            .Sheets(shFn).Range(title & i) = Replace(ThisWorkbook.Sheets(sh).Range("L" & i), Chr(10), "")
            'episode
            .Sheets(shFn).Range("B" & i) = Replace(ThisWorkbook.Sheets(sh).Range("O" & i), Chr(10), "")
            'aspect
            .Sheets(shFn).Range("Q" & i) = MyFunct.GetAspect(sh, 27, i, True)
            .Sheets(shFn).Range("P" & i) = MyFunct.GetAspect(sh, 27, i, False)
            'M/TV
            .Sheets(shFn).Range("G" & i) = ThisWorkbook.Sheets(sh).Range("A" & i)
            'status
            .Sheets(shFn).Range("H" & i) = ThisWorkbook.Sheets(sh).Range("P" & i)
            'rating
            .Sheets(shFn).Range("N" & i) = ThisWorkbook.Sheets(sh).Range("I" & i)
            'year
            .Sheets(shFn).Range("O" & i) = ThisWorkbook.Sheets(sh).Range("H" & i)
        End With
        
        If InStr(1, result, "!") > 0 Then
            ThisWorkbook.Sheets(shFn).Range(dubs & i) = result
            ThisWorkbook.Sheets(shFn).Range(dubs & i).Interior.Color = vbRed
        Else:
            ThisWorkbook.Sheets(shFn).Range(dubs & i) = result
            ThisWorkbook.Sheets(shFn).Range(dubs & i).Font.Color = vbBlack
        End If
        
        'date of start
        If Len(ThisWorkbook.Sheets(sh).Range("AC" & i)) > 2 Then
            With ThisWorkbook.Sheets(shFn)
                .Range("L" & i).NumberFormat = "@"
                If InStr(1, .Cells(i, 8).value, "new", vbTextCompare) > 0 Then
                    .Range("L" & i) = Format(Split(ThisWorkbook.Sheets("Notes").Cells(6, 1).value, "|")(1), "00")
                    .Range("M" & i) = Split(ThisWorkbook.Sheets("Notes").Cells(6, 1).value, "|")(2)
                Else
                    .Range("L" & i) = Format(Month(ThisWorkbook.Sheets(sh).Range("AC" & i)), "00")
                    .Range("M" & i) = Year(ThisWorkbook.Sheets(sh).Range("AC" & i))
                End If
            End With
        End If
        
        Call MovieOrTV(ThisWorkbook.Sheets(shFn), i)
        Call DividingLanguages(i)
        Call CTRfirst.FindRaveStatus(ThisWorkbook.Sheets(sh), ThisWorkbook.Sheets("NOTES"), i)
    Next i
    
    'MsgBox "√отово"
    
    Ёта нига.Closed
    frmLoad.Unloading
End Sub

Private Sub MovieOrTV(ByRef shFn As Worksheet, i As Long)
        'Dim shFn As Worksheet: Set shFn = ThisWorkbook.Sheets("Filenames")
        Dim typeTitle As String, title As String, lastTitle As String, episode As String, numSeason As String, numEpisode As String
        'Dim i As Long
        
        'For i = 3 To lastRow
            typeTitle = shFn.Cells(i, 7).value
            
            lastTitle = shFn.Cells(i - 1, 1).value
            
            If (StrComp(typeTitle, "Movie", vbTextCompare) = 0) Then
                'title withoutSymbols
                title = RemoveSpecSymbols(shFn.Cells(i, 1).value)
                'search title
                shFn.Cells(i, 10).value = Trim(shFn.Cells(i, 1).value)
                
                Call InsertTitleEpisodeFilenames(i, title)
                If (Len(shFn.Cells(i, 17).value) > 0 And Len(shFn.Cells(i, 16).value) > 0) Then
                    shFn.Cells(i, 9) = "KZR_" & shFn.Cells(i, 12) & Mid(shFn.Cells(i, 13), 3) & "_" & Replace(shFn.Cells(i, 5), " ", "#") & "_SSS_" & shFn.Cells(i, 17) & "_DDD"
                End If
            ElseIf (StrComp(typeTitle, "TV", vbTextCompare) = 0) Then
                title = shFn.Cells(i, 1).value
                episode = shFn.Cells(i, 2).value
                
                'as default it will be single
                If (Len(shFn.Cells(i, 17).value) > 0 And Len(shFn.Cells(i, 16).value) > 0) Then
                    shFn.Cells(i, 44).value = "SingleEpisode"
                End If
                
                'search title
                If Len(episode) > 1 Then
                    shFn.Cells(i, 10).value = Trim(title) & " " & Trim(episode)
                Else: shFn.Cells(i, 10).value = Trim(title)
                End If
                
                'episode number
                If StrComp(lastTitle, title, vbTextCompare) = 0 Then
                    numEpisode = CLng(shFn.Cells(i - 1, 4).value) + 1
                    If InStr(1, shFn.Cells(i - 1, 44).value, "single", vbTextCompare) > 0 Or IsEmpty(shFn.Cells(i - 1, 44).value) Then shFn.Cells(i - 1, 44).value = "BoxedEpisode"
                    shFn.Cells(i, 44).value = "BoxedEpisode"
                Else:
                    lastTitle = title
                    numEpisode = 1
                End If
                
                'season number
                If InStr(1, title, "season", vbTextCompare) > 0 Then
                    numSeason = Replace(Mid(title, InStr(1, title, "season", vbTextCompare) + 7), " ", "")
                    title = Left(title, InStr(1, title, "season", vbTextCompare) - 2)
                Else:
                    numSeason = 1
                End If
                
                'remove symbols if smth inside
                If Len(title) > 1 Then title = RemoveSpecSymbols(title)
                If Len(episode) > 1 Then episode = RemoveSpecSymbols(episode)
                
                'check do we need insert number season and episodes
                If (Len(shFn.Cells(i, 17).value) > 0 And Len(shFn.Cells(i, 16).value) > 0) Then
                    If InStr(1, shFn.Cells(i, 8).value, "box", vbTextCompare) > 0 Or InStr(1, ThisWorkbook.Sheets("Initial").Cells(i, 6).value, "best series", vbTextCompare) > 0 Then
                        Call InsertTitleEpisodeFilenames(i, title, episode, numSeason, numEpisode)
                        shFn.Cells(i, 9) = "KZR_" & shFn.Cells(i, 12) & Mid(shFn.Cells(i, 13), 3) & "_" & Replace(shFn.Cells(i, 5) & " Season " & shFn.Cells(i, 3) & " Ep " & shFn.Cells(i, 4), " ", "#") & "_SSS_" & shFn.Cells(i, 17) & "_DDD"
                    Else:
                        Call InsertTitleEpisodeFilenames(i, title, episode, numSeason, numEpisode)
                        Dim epis As String:
                        If Len(shFn.Cells(i, 6)) = 0 Then
                            epis = ""
                        Else: epis = " " & shFn.Cells(i, 6)
                        End If
                        shFn.Cells(i, 9) = "KZR_" & shFn.Cells(i, 12) & Mid(shFn.Cells(i, 13), 3) & "_" & Replace(shFn.Cells(i, 5) & epis, " ", "#") & "_SSS_" & shFn.Cells(i, 17) & "_DDD"
                    End If
                End If
            End If
        'Next i
        
End Sub

Private Sub InsertTitleEpisodeFilenames(row As Long, title As String, Optional episode As String = "", _
                                        Optional numSeason As String = "", Optional numEpisode As String = "")
    With ThisWorkbook.Sheets("Filenames")
        'title
        .Cells(row, 5).value = title
        'episde
        .Cells(row, 6).value = episode
        'numSeas
        .Cells(row, 3).value = numSeason
        'numEpis
        .Cells(row, 4).value = numEpisode
    End With
End Sub

Private Sub DividingLanguages(i As Long)
    Dim arr
    Dim ArColDub() As Variant, ArColSub() As Variant, ArColSubS() As Variant
    Dim colDub As Integer, colSub As Integer, colSubS As Integer
    Dim j As Long
    Dim countRow As Long
    Dim forFN As String
    Dim charr As Variant
    Dim draftSh As String, fnSheet As String
    
    draftSh = "draft"
    fnSheet = "Filenames"
    
    'i = 3
    ArColDub = Array("S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB")
    ArColSub = Array("AC", "AD", "AE", "AF", "AG", "AH")
    ArColSubS = Array("AK", "AL", "AM", "AN", "AO", "AP")
    
    'countRow = ThisWorkbook.Sheets("Filenames").UsedRange.rows(ThisWorkbook.Sheets("Filenames").UsedRange.rows.count).row
    
    'ThisWorkbook.Worksheets(fnSheet).Range("R3:AP" & countRow).Clear
    
    'Do While i < countRow + 1
        With ThisWorkbook.Worksheets(fnSheet).Range("S" & i & ":AP" & i)
        .Interior.Color = RGB(255, 230, 153)
        .Borders(7).Color = RGB(0, 0, 0)
        .Borders(8).Color = RGB(0, 0, 0)
        .Borders(9).Color = RGB(0, 0, 0)
        .Borders(10).Color = RGB(0, 0, 0)
        .Borders(11).Color = RGB(0, 0, 0)
        End With
             
        colDub = 0
        colSub = 0
        colSubS = 0
        
        'Array of Dubs and Subs
        arr = Split(ThisWorkbook.Worksheets(fnSheet).Range("K" & i), "/")
        Dim el As Variant
        For Each el In arr
            If InStr(3, el, "S") > 0 And Not InStr(1, el, "Sub") > 0 Then
                ThisWorkbook.Sheets(fnSheet).Range(ArColSubS(colSubS) & i).value = el
                colSubS = colSubS + 1
            ElseIf InStr(1, el, "DYN") > 0 Or InStr(1, el, "Sub") > 0 Then
                ThisWorkbook.Sheets(fnSheet).Range(ArColSub(colSub) & i).value = el
                colSub = colSub + 1
            ElseIf InStr(1, el, "CC") > 0 Or InStr(el, "——") > 0 Then
                ThisWorkbook.Sheets(fnSheet).Range("AJ" & i).value = el
            ElseIf InStr(1, el, "AD") > 0 Then
                ThisWorkbook.Sheets(fnSheet).Range("AI" & i).value = el
            Else
                ThisWorkbook.Sheets(fnSheet).Range(ArColDub(colDub) & i).value = el
                colDub = colDub + 1
            End If
        Next el
            Dim dvs As String
            If InStr(1, ThisWorkbook.Sheets(fnSheet).Cells(i, 35).value, "EngAD", vbTextCompare) Then
                dvs = "Dvs"
            ElseIf Len(ThisWorkbook.Sheets(fnSheet).Cells(i, 35).value) > 1 Then
                dvs = ThisWorkbook.Sheets(fnSheet).Cells(i, 35).value
            Else: dvs = ""
            End If
            
        Dim cellsCount As Long: cellsCount = 10
        Dim result As StringBuilderMy:
        Set result = New StringBuilderMy
        
        For j = 0 To cellsCount
            If Len(ThisWorkbook.Sheets(fnSheet).Cells(i, 19 + j).value) > 0 Then
                result.Append (ThisWorkbook.Sheets(fnSheet).Cells(i, 19 + j).value & "|")
            ElseIf Len(dvs) > 0 Then
                result.Append (dvs & "|")
                Exit For
            Else
                Exit For
            End If
        Next j
        If (Len(result.ToString)) > 0 Then ThisWorkbook.Sheets(fnSheet).Cells(i, 18).value = Left(result.ToString, Len(result.ToString) - 1)
        'ThisWorkbook.Sheets(fnSheet).Cells(i, 18).value = ThisWorkbook.Sheets(fnSheet).Cells(i, 19).value & _
                                                         ThisWorkbook.Sheets(fnSheet).Cells(i, 20).value & _
                                                         ThisWorkbook.Sheets(fnSheet).Cells(i, 21).value & _
                                                         ThisWorkbook.Sheets(fnSheet).Cells(i, 22).value & _
                                                         ThisWorkbook.Sheets(fnSheet).Cells(i, 23).value & _
                                                         ThisWorkbook.Sheets(fnSheet).Cells(i, 24).value & _
                                                         ThisWorkbook.Sheets(fnSheet).Cells(i, 25).value & _
                                                         dvs
        Set result = New StringBuilderMy
        cellsCount = 6
        For j = 0 To cellsCount
            If Len(ThisWorkbook.Sheets(fnSheet).Cells(i, 29 + j).value) > 0 Then
                result.Append (ThisWorkbook.Sheets(fnSheet).Cells(i, 29 + j).value & "|")
            ElseIf Len(ThisWorkbook.Sheets(fnSheet).Cells(i, 36).value) > 0 Then
                result.Append (ThisWorkbook.Sheets(fnSheet).Cells(i, 36).value & "|")
                Exit For
            Else
                Exit For
            End If
        Next j
        If (Len(result.ToString)) > 0 Then ThisWorkbook.Sheets(fnSheet).Cells(i, 43).value = Left(result.ToString, Len(result.ToString) - 1)
        
        'i = i + 1
        Erase arr
    'Loop
End Sub

Function RemoveSpecSymbols(title As String, Optional episode As String) As String
    Dim charr As Variant
    
        title = Trim(title) & Trim(episode)
'RemoveSpecialSymbols
        Const SpecialCharacters As String = "Т|!|?|:|,|Ц|'|.|-|Е|^|&|*|(|)|\|/"
        
        For Each charr In Split(SpecialCharacters, "|")
            title = Replace(title, charr, "")
        Next
            title = Replace(title, Chr(10), "")
    RemoveSpecSymbols = title
End Function

Private Function MakeLangStr(ArrLanInitial() As String, nameSH, nameCol, numbRow) As String
   
    Dim result As String
    Dim resultRu As String
    Dim kel As Variant
    Dim i As Integer
    
    i = 0
    For Each kel In ArrLanInitial

        If i > 0 Then
            result = result + "/"
        End If
        
            If InStr(1, LCase(kel), "афр", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Afr")
                'result = result + "Afr"
            ElseIf InStr(1, LCase(kel), "амх", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Amh")
                'result = result + "Amh"
            ElseIf InStr(1, LCase(kel), "ара", vbTextCompare) > 0 Or InStr(1, LCase(kel), "A–јЅ", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Ara")
                'result = result + "Ara"
            ElseIf InStr(1, LCase(kel), "азер", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Aze")
                'result = result + "Aze"
            ElseIf InStr(1, LCase(kel), "бенг", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Ben")
                'result = result + "Ben"
            ElseIf InStr(1, LCase(kel), "болг", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Bul")
                'result = result + "Bul"
            ElseIf InStr(1, LCase(kel), "кат", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Cat")
                'result = result + "Cat"
            ElseIf InStr(1, LCase(kel), "кит", vbTextCompare) > 0 Or _
                   InStr(1, LCase(kel), "кан", vbTextCompare) > 0 Or _
                   InStr(1, LCase(kel), "ман", vbTextCompare) > 0 Or _
                   InStr(1, kel, "MAЌ", vbTextCompare) > 0 Then
                If InStr(1, kel, "кан", vbTextCompare) > 0 Or LCase(kel) Like "c" Or LCase(kel) Like "с" Then
                    result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Yue")
                    'result = result + "Zhc"
                ElseIf InStr(1, kel, "(m", vbTextCompare) > 0 Or InStr(1, kel, "(м", vbTextCompare) > 0 Or InStr(1, kel, "ман", vbTextCompare) > 0 Or InStr(1, kel, "MAЌ", vbTextCompare) > 0 Then
                    result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Cmn")
                    'result = result + "Cmn"
                ElseIf InStr(1, LCase(kel), "(tai", vbTextCompare) > 0 Or InStr(1, LCase(kel), "(тай", vbTextCompare) > 0 Then
                        result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Zht")
                        'result = result + "Zht"
                ElseIf InStr(1, LCase(kel), "cbt", vbTextCompare) > 0 Or InStr(1, LCase(kel), "сбт", vbTextCompare) > 0 Then
                    If InStr(1, LCase(kel), "si", vbTextCompare) > 0 Or InStr(1, LCase(kel), "сим", vbTextCompare) > 0 Then
                        result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Chi")
                        'result = result + "Chi"
                    Else
                        result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Zho")
                        'result = result + "Chi"
                    End If
                Else
                    result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Zhc")
                    'result = result + "Zhc"
                End If
            ElseIf InStr(1, LCase(kel), "хорв", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Hrv")
                'result = result + "Hrv"
            ElseIf InStr(1, LCase(kel), "чеж", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Ces")
                'result = result + "Ces"
            ElseIf InStr(1, LCase(kel), "дац", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Dan")
                'result = result + "Dan"
            ElseIf InStr(1, LCase(kel), "гол", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Nld")
                'result = result + "Nld"
            'ElseIf InStr(1, LCase(kel), "ad", vbTextCompare) > 0 Then
                'Result = Result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Afr")
                'result = result + "Dvs"
            ElseIf InStr(1, LCase(kel), "анг", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Eng")
                'result = result + "Eng"
            ElseIf InStr(1, LCase(kel), "эст", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Est")
                'result = result + "Est"
            ElseIf InStr(1, LCase(kel), "фарс", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Fas")
                'result = result + "Fas"
            ElseIf InStr(1, LCase(kel), "‘инс", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Fin")
                'result = result + "Fin"
            ElseIf InStr(1, LCase(kel), "фр", vbTextCompare) > 0 Then
                If InStr(1, kel, "фр(can") Then
                    result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Cfr")
                    'result = result + "Cfr"
                ElseIf InStr(1, kel, "фр") > 0 Then
                    result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Fra")
                    'result = result + "Fra"
                Else
                    result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Fra")
                    'result = result + kel
                End If
            ElseIf InStr(1, LCase(kel), "гэль", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Gle")
                'result = result + "Gle"
            ElseIf InStr(1, LCase(kel), "нем", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Deu")
                'result = result + "Deu"
            ElseIf InStr(1, LCase(kel), "греч", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Ell")
                'result = result + "Ell"
            ElseIf InStr(1, LCase(kel), "ивр", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Heb")
                'result = result + "Heb"
            ElseIf InStr(1, LCase(kel), "хин", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Hin")
                'result = result + "Hin"
            ElseIf InStr(1, LCase(kel), "венг", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Hun")
                'result = result + "Hun"
            ElseIf InStr(1, LCase(kel), "исл", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Isl")
                'result = result + "Isl"
            ElseIf InStr(1, LCase(kel), "игбо", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Ibo")
                'result = result + "Ibo"
            ElseIf InStr(1, LCase(kel), "индо", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Ind")
                'result = result + "Ind"
            ElseIf InStr(1, LCase(kel), "ит", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Ita")
                'result = result + "Ita"
            ElseIf InStr(1, LCase(kel), "€п", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Jpn")
                'result = result + "Jpn"
            ElseIf InStr(1, LCase(kel), "каз", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Kaz")
                'result = result + "Kaz"
            ElseIf InStr(1, LCase(kel), "кинь", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Kin")
                'result = result + "Kin"
            ElseIf InStr(1, LCase(kel), "киту", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Ktu")
                'result = result + "Ktu"
            ElseIf InStr(1, LCase(kel), "кор", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Kor")
                'result = result + "Kor"
            ElseIf InStr(1, LCase(kel), "лао", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Lao")
                'result = result + "Lao"
            ElseIf InStr(1, LCase(kel), "лат", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Lav")
                'result = result + "Lav"
            ElseIf InStr(1, LCase(kel), "линг", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Lin")
                'result = result + "Lin"
            ElseIf InStr(1, LCase(kel), "литв", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Lit")
                'result = result + "Lit"
            ElseIf InStr(1, LCase(kel), "мак", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Mkd")
                'result = result + "Mkd"
            ElseIf InStr(1, LCase(kel), "мал", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Msa")
                'result = result + "Msa"
            ElseIf InStr(1, LCase(kel), "норв", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Nor")
                'result = result + "Nor"
            ElseIf InStr(1, LCase(kel), "про", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Pol")
                'result = result + "Pol"
            ElseIf InStr(1, LCase(kel), "порт", vbTextCompare) > 0 Then
                If InStr(1, kel, "порт*(bra)*") Then
                    result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Pbr")
                    'result = result + "Pbr"
                ElseIf InStr(1, kel, "порт(eu)*") Then
                    result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Por")
                    'result = result + "Por"
                Else
                    result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Por")
                    'result = result + kel
                End If
            ElseIf InStr(1, LCase(kel), "панд", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Pan")
                'result = result + "Pan"
            ElseIf InStr(1, LCase(kel), "*рум", vbTextCompare) > 0 Or _
                   InStr(1, LCase(kel), "молд", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Ron")
                'result = result + "Ron"
            ElseIf InStr(1, LCase(kel), "рус", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Rus")
                'result = result + "Rus"
            ElseIf InStr(1, LCase(kel), "серб", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Hbs")
                'result = result + "Hbs"
            ElseIf InStr(1, LCase(kel), "синг", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Sin")
                'result = result + "Sin"
            ElseIf InStr(1, LCase(kel), "слов", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Slk")
                'result = result + "Slk"
            ElseIf InStr(1, LCase(kel), "исп", vbTextCompare) > 0 Then
                'If InStr(1, kel, "исп(c") Then
                    result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Spa")
                    'result = result + "Spa"
                'ElseIf InStr(1, kel, "исп(l") Then
                    'result = result + LookingForDubsOrSubs(kel,  nameSH, nameCol, numbRow, "Spl")
                    'result = result + "Spl"
                'Else
                    'result = result + LookingForDubsOrSubs(kel,  nameSH, nameCol, numbRow, "Spl")
                    'result = result + kel
                'End If
            ElseIf InStr(1, LCase(kel), "суа", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Swa")
                'result = result + "Swa"
            ElseIf InStr(1, LCase(kel), "швед", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Swe")
                'result = result + "Swe"
            ElseIf InStr(1, LCase(kel), "шв-нм", vbTextCompare) > 0 Or _
                   InStr(1, LCase(kel), "алем", vbTextCompare) > 0 Or _
                   InStr(1, LCase(kel), "эльз", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Gsw")
                'result = result + "Gsw"
            ElseIf InStr(1, LCase(kel), "таг", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Tgl")
                'result = result + "Tgl"
            ElseIf InStr(1, LCase(kel), "тамил", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Tam")
                'result = result + "Tam"
            ElseIf InStr(1, LCase(kel), "тай", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Tha")
                'result = result + "Tha"
            ElseIf InStr(1, LCase(kel), "тур", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Tur")
                'result = result + "Tur"
            ElseIf InStr(1, LCase(kel), "укр", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Ukr")
                'result = result + "Ukr"
            ElseIf InStr(1, LCase(kel), "урд", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Urd")
                'result = result + "Urd"
            ElseIf InStr(1, LCase(kel), "вьет", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Vie")
                'result = result + "Vie"
            ElseIf InStr(1, LCase(kel), "вол", vbTextCompare) > 0 Then
                result = result + LookingForDubsOrSubs(kel, nameSH, nameCol, numbRow, "Wol")
                'result = result + "Wol"
            Else
                result = result & "!" & kel
        End If
        i = i + 1
        Next kel
        'MsgBox result
        MakeLangStr = result
End Function

Private Function LookingForDubsOrSubs(kel As Variant, nameSH, nameCol, numbRow, langMark As String) As String
    Dim val As String
    If numbRow = 412 Then
        'MsgBox "a"
    End If
    val = LCase(ThisWorkbook.Sheets(nameSH).Range(nameCol & numbRow).value)
    
    If InStr(1, LCase(kel), "cbt", vbTextCompare) Or InStr(1, LCase(kel), "сбт", vbTextCompare) Then
        If InStr(1, val, "дин", vbTextCompare) = 0 And InStr(1, val, "прош", vbTextCompare) = 0 Then val = "динамические"
                    If InStr(1, val, ",", vbTextCompare) Then
                        Dim twoVal() As String
                        Dim i As Long
                        twoVal = Split(val, ",")
                        For i = LBound(twoVal) To UBound(twoVal)
                            If langMark = "Eng" And InStr(1, twoVal(i), "англ", vbTextCompare) Then
                                val = twoVal(i)
                            ElseIf langMark = "Rus" And InStr(1, twoVal(i), "рус", vbTextCompare) Then
                                val = twoVal(i)
                            End If
                        Next i
                    End If
                    If InStr(1, val, "динамические", vbTextCompare) > 0 Or _
                        LCase(ThisWorkbook.Sheets(nameSH).Range(nameCol & numbRow)) = "" Then
                        LookingForDubsOrSubs = langMark + " -DYN Sub"
                    ElseIf InStr(1, val, "прошитые", vbTextCompare) > 0 Then
                        LookingForDubsOrSubs = langMark & "S"
                    Else: LookingForDubsOrSubs = langMark & "!Sub"
                    End If
    ElseIf InStr(1, LCase(kel), "cc", vbTextCompare) Or InStr(1, LCase(kel), "сс", vbTextCompare) Then
                    LookingForDubsOrSubs = langMark & "CC"
    ElseIf InStr(1, LCase(kel), "ad", vbTextCompare) Or InStr(1, LCase(kel), "тифло", vbTextCompare) Then
                    LookingForDubsOrSubs = langMark & "AD"
    Else: LookingForDubsOrSubs = langMark
    End If
End Function



