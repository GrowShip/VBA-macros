Attribute VB_Name = "PrepList"
Option Explicit
Public ddd As Variant
Public mmm As Variant
Public yyy As Variant
Public flagform As Boolean

Sub FillPrepList()

    Call ShowForm
    
    If flagform = False Then Exit Sub
    ThisWorkbook.Sheets("Notes").Cells(5, 1).value = ddd & "|" & mmm & "|" & yyy
    
    Dim typeStudio As String
    Dim dict As New Dictionary
    
    If Len(ThisWorkbook.Sheets("Information").Cells(22, 8).value) = 0 Then
        MsgBox "Вы не выбрали студию"
        Exit Sub
    Else
        ThisWorkbook.Sheets("Information").Cells(22, 9).value = ThisWorkbook.Sheets("Information").Cells(22, 8).value
        typeStudio = ThisWorkbook.Sheets("Information").Cells(22, 9).value
        ThisWorkbook.Sheets("Information").Cells(22, 8).ClearContents
    End If
    
    Set dict = FindingAndFillingTitles(typeStudio)
    
    If dict.count > 0 Then
        'MsgBox "Не пусто"
        Call repParseDict(dict)
        Call MainInfo
    Else
        MsgBox "Для студии " + typeStudio + " нет новинок("
    End If
End Sub

Private Function FindingAndFillingTitles(nameStudio) As Dictionary
    Dim dict As Dictionary
    Set dict = New Dictionary
    
    Dim isSmallSt As Boolean
    Dim row As Long, i As Long
    Dim sh As Worksheet
    Dim asMinAsOne As Boolean
    
    If StrComp("other", nameStudio, vbTextCompare) = 0 Then
        Set FindingAndFillingTitles = FindingSmallStudios()
        Exit Function
    ElseIf StrComp("all", nameStudio, vbTextCompare) = 0 Then
        Set FindingAndFillingTitles = FindingAllStudios()
        Exit Function
    End If
    
    Set sh = ThisWorkbook.Sheets("Initial")
    row = MyFunct.countRows(sh.name)
    
    For i = 3 To row
        If InStr(1, sh.Cells(i, 14).value, nameStudio, vbTextCompare) > 0 And _
           InStr(1, sh.Cells(i, 15).value, "Aerogroup", vbTextCompare) = 0 Then
            Dim arr As Variant
            Dim ex3, exw, ipad
            ex3 = ""
            exw = ""
            ipad = ""
            'Rave
            If InStr(1, sh.Cells(i, 11).value, "new", vbTextCompare) > 0 Then ex3 = "New"
            'Ipad
            If InStr(1, sh.Cells(i, 12).value, "new", vbTextCompare) > 0 Then exw = "New"
            'Wow
            If InStr(1, sh.Cells(i, 13).value, "new", vbTextCompare) > 0 Then ipad = "New"
            
            If ex3 = "New" Or exw = "New" Or ipad = "New" Then
                arr = Array(ex3, exw, ipad)
                dict.Item(i) = arr
                asMinAsOne = True
            End If
        End If
    Next i
    If asMinAsOne Then
        Set FindingAndFillingTitles = dict
    Else
        Set FindingAndFillingTitles = Nothing
        Exit Function
    End If
End Function

Private Sub repParseDict(dict As Dictionary)
    Dim totalRow As Long, i As Long
    Dim sh As Worksheet
    Dim row
    Set sh = ThisWorkbook.Sheets("Prep list")
    
    totalRow = MyFunct.countRows(sh.name)
        
    If ThisWorkbook.Sheets("Information").Cells(25, 8).value = 1 Then
        i = totalRow
    ElseIf ThisWorkbook.Sheets("Information").Cells(25, 8).value = 0 Then
        If totalRow > 4 Then sh.Range("A4:Z" & totalRow).ClearContents
        i = 4
    Else
        MsgBox "Вы не выбрали вконец заполнять или сначала"
        Exit Sub
    End If
    
    ThisWorkbook.Sheets("Information").Cells(25, 8).ClearContents
    For Each row In dict
        With sh
            .Cells(i, 25).value = row
            .Cells(i, 19).value = dict(row)(0)
            .Cells(i, 20).value = dict(row)(1)
            .Cells(i, 21).value = dict(row)(2)
        End With
        i = i + 1
    Next row
    'MsgBox "Prep list Done"
End Sub
Private Function FindingAllStudios() As Dictionary
    Dim dict As Dictionary
    Set dict = New Dictionary
    
    Dim sh As Worksheet
    Dim row As Long
    Dim asOneAsMin As Boolean
    Dim i As Long
    Set sh = ThisWorkbook.Sheets("Initial")
    row = MyFunct.countRows(sh.name)
    
    For i = 3 To row
        If InStr(1, sh.Cells(i, 15).value, "Aerogroup", vbTextCompare) = 0 Then
            Dim arr As Variant
            Dim ex3, exw, ipad
            ex3 = ""
            exw = ""
            ipad = ""
            'Rave
            If InStr(1, sh.Cells(i, 11).value, "new", vbTextCompare) > 0 Then ex3 = "New"
            'Ipad
            If InStr(1, sh.Cells(i, 12).value, "new", vbTextCompare) > 0 Then exw = "New"
            'Wow
            If InStr(1, sh.Cells(i, 13).value, "new", vbTextCompare) > 0 Then ipad = "New"
            
            If ex3 = "New" Or exw = "New" Or ipad = "New" Then
                arr = Array(ex3, exw, ipad)
                dict.Item(i) = arr
                asOneAsMin = True
            End If
        End If
    Next i
    
    If asOneAsMin Then
        Set FindingAllStudios = dict
    Else
        Set FindingAllStudios = Nothing
    End If
End Function

Private Function FindingSmallStudios() As Dictionary
    Dim dict As Dictionary
    Set dict = New Dictionary
    
    Dim sh As Worksheet
    Dim row As Long
    Dim asOneAsMin As Boolean
    Dim i As Long
    Set sh = ThisWorkbook.Sheets("Initial")
    row = MyFunct.countRows(sh.name)
    
    For i = 3 To row
        If InStr(1, sh.Cells(i, 15).value, "Aerogroup", vbTextCompare) = 0 Then
            If InStr(1, sh.Cells(i, 14).value, "Disney", vbTextCompare) = 0 And _
               InStr(1, sh.Cells(i, 14).value, "Warner", vbTextCompare) = 0 And _
               InStr(1, sh.Cells(i, 14).value, "NBC", vbTextCompare) = 0 And _
               InStr(1, sh.Cells(i, 14).value, "Sony", vbTextCompare) = 0 And _
               InStr(1, sh.Cells(i, 14).value, "Paramount", vbTextCompare) = 0 And _
               InStr(1, sh.Cells(i, 14).value, "HBO", vbTextCompare) = 0 Then
                Dim arr As Variant
                Dim ex3, exw, ipad
                ex3 = ""
                exw = ""
                ipad = ""
                'Rave
                If InStr(1, sh.Cells(i, 11).value, "new", vbTextCompare) > 0 Then ex3 = "New"
                'Ipad
                If InStr(1, sh.Cells(i, 12).value, "new", vbTextCompare) > 0 Then exw = "New"
                'Wow
                If InStr(1, sh.Cells(i, 13).value, "new", vbTextCompare) > 0 Then ipad = "New"
                
                If ex3 = "New" Or exw = "New" Or ipad = "New" Then
                    arr = Array(ex3, exw, ipad)
                    dict.Item(i) = arr
                    asOneAsMin = True
                End If
            End If
        End If
    Next i
    If asOneAsMin Then
        Set FindingSmallStudios = dict
    Else
        Set FindingSmallStudios = Nothing
    End If
End Function

Private Sub MainInfo()
    Dim sh As Worksheet
    Dim row As Long, i As Long
    Set sh = ThisWorkbook.Sheets("Prep list")
    row = MyFunct.countRows(sh.name)
    
    For i = 4 To row
        Call MainInfoFill(i, sh.Cells(i, 25).value)
    Next i
    MsgBox " Prep List Ready "
End Sub

Private Sub MainInfoFill(rowpr As Long, rowfn As Long)

    Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Prep list")
    Dim shIn As Worksheet: Set shIn = ThisWorkbook.Sheets("Initial")
    Dim shFn As Worksheet: Set shFn = ThisWorkbook.Sheets("Filenames")
    With sh
        .Cells(rowpr, 1).value = shIn.Cells(rowfn, 48).value 'PO
        .Cells(rowpr, 2).value = shIn.Cells(rowfn, 8).value & "|" & shIn.Cells(rowfn, 9).value 'Tittle|Episode
        .Cells(rowpr, 3).value = shIn.Cells(rowfn, 14).value 'Studia
        .Cells(rowpr, 4).value = shIn.Cells(rowfn, 15).value 'Lab
        .Cells(rowpr, 5).value = GetAvailablePeriod(shFn.Cells(rowfn, 14).value) 'Period ?
        .Cells(rowpr, 6).value = "Under annual deal" 'Deal
        .Cells(rowpr, 7).value = shIn.Cells(rowfn, 4).value 'Year
        .Cells(rowpr, 8).value = shIn.Cells(rowfn, 22).value 'Aspect
        .Cells(rowpr, 9).value = shFn.Cells(rowfn, 19).value 'Dub1
        .Cells(rowpr, 10).value = shFn.Cells(rowfn, 20).value 'Dub2
        .Cells(rowpr, 11).value = shFn.Cells(rowfn, 21).value  'Dub3
        .Cells(rowpr, 12).value = shFn.Cells(rowfn, 22).value  'Dub4
        .Cells(rowpr, 13).value = shFn.Cells(rowfn, 23).value  'Dub5
        .Cells(rowpr, 14).value = shFn.Cells(rowfn, 24).value  'Dub6
        .Cells(rowpr, 26).value = shIn.Cells(rowfn, 1).value  'M/TV
    End With
        Call TakeAsubs(rowpr, rowfn)
End Sub

Private Sub TakeAsubs(rowpr As Long, rowfn As Long)
    Dim shFn As Worksheet, shpr As Worksheet
    Dim i As Long
    
    Set shFn = ThisWorkbook.Sheets("Filenames")
    Set shpr = ThisWorkbook.Sheets("Prep list")
    
    For i = 0 To 6
        If Len(shFn.Cells(rowfn, 29 + i).value) > 1 Then
            shpr.Cells(rowpr, 15 + i).value = shFn.Cells(rowfn, 29 + i).value
        ElseIf Len(shFn.Cells(rowfn, 36).value) > 1 Then
            shpr.Cells(rowpr, 15 + i).value = shFn.Cells(rowfn, 36).value
            Exit For
        Else: Exit For
        End If
    Next i
    
End Sub

Private Sub ShowForm()
    POform.Show
End Sub

Private Function GetAvailablePeriod(period As String) As String
    If InStr(1, period, "-", vbTextCompare) = 0 Then
        GetAvailablePeriod = ""
        Exit Function
    End If
    
    Dim arr As Variant: arr = Split(period, "-")
    Dim i As Integer
    
    For i = 0 To UBound(arr)
        arr(i) = Trim(arr(i))
    Next i
    
    Dim fromDate As String: fromDate = SplitMonthAndYear(CStr(arr(0)))
    Dim toDate As String: toDate = SplitMonthAndYear(CStr(arr(1)))
    
    GetAvailablePeriod = fromDate + "- " + Left(toDate, Len(toDate) - 1)
End Function

Private Function SplitMonthAndYear(fullDate As String) As String
    Dim monthM As String, yearY As String
    Dim datet As String: datet = fullDate
    
    If InStr(1, fullDate, " ") > 0 Then
        monthM = TranslateMonth(CStr(Split(datet, " ")(0))) + " "
        yearY = Split(datet, " ")(1) + " "
    Else
        monthM = TranslateMonth(datet)
        yearY = " "
    End If
    
    SplitMonthAndYear = monthM + yearY
End Function

Private Function TranslateMonth(month As String) As String
    If InStr(1, month, "янва", vbTextCompare) > 0 Then
        TranslateMonth = "January"
    ElseIf InStr(1, month, "февр", vbTextCompare) > 0 Then
        TranslateMonth = "February"
    ElseIf InStr(1, month, "март", vbTextCompare) > 0 Then
        TranslateMonth = "March"
    ElseIf InStr(1, month, "апр", vbTextCompare) > 0 Then
        TranslateMonth = "April"
    ElseIf InStr(1, month, "май", vbTextCompare) > 0 Then
        TranslateMonth = "May"
    ElseIf InStr(1, month, "июнь", vbTextCompare) > 0 Then
        TranslateMonth = "June"
    ElseIf InStr(1, month, "июль", vbTextCompare) > 0 Then
        TranslateMonth = "July"
    ElseIf InStr(1, month, "авгу", vbTextCompare) > 0 Then
        TranslateMonth = "August"
    ElseIf InStr(1, month, "сент", vbTextCompare) > 0 Then
        TranslateMonth = "September"
    ElseIf InStr(1, month, "окт", vbTextCompare) > 0 Then
        TranslateMonth = "October"
    ElseIf InStr(1, month, "нояб", vbTextCompare) > 0 Then
        TranslateMonth = "November"
    ElseIf InStr(1, month, "дек", vbTextCompare) > 0 Then
        TranslateMonth = "December"
    Else: TranslateMonth = month
    End If
End Function
