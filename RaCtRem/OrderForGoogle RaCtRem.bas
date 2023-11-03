Attribute VB_Name = "OrderForGoogle"
Option Explicit
Public Sub CreateGooOrder()
    Dim ask As Long: ask = frmDidYouUpload.Loading("Создаем заказ для Гудини лабы?")
    
    If ask = 2 Then
        Exit Sub
    ElseIf ask = 6 Then
        FillInOrdForGoogle
    ElseIf ask = 7 Then
        Exit Sub
    End If
End Sub


Sub FillInOrdForGoogle()
    
    frmLoad.Loading
    ЭтаКнига.Opened
    
    Dim countRow As Long
    Dim sh As Worksheet, shFn As Worksheet
    Dim orderSh As Worksheet
    Dim i As Long
    Dim j As Long
    Dim newRaveIpadWow() As Variant
    Dim fileName As String
    Dim datet As String
    Dim aspect As String
    Dim dateArr: dateArr = Split(ThisWorkbook.Sheets("NOTES").Cells(6, 1), "|")
    
    Set sh = ThisWorkbook.Sheets("Initial")
    Set orderSh = ThisWorkbook.Sheets("OrderGoogle")
    Set shFn = ThisWorkbook.Sheets("Filenames")
    
    countRow = MyFunct.countRowBest(orderSh.name)
    If countRow > 1 Then orderSh.rows("2:" & countRow).Clear
    
    
    'ДАТА Мм Yy
    datet = dateArr(1) & Right(dateArr(2), 2)
    'actual Row
    j = 2
    'Подсчет строк для обработки
    countRow = MyFunct.countRowBest(sh.name)
    
    For i = 3 To countRow + 1
        Dim whatIs As String
        If i = 270 Then
            'MsgBox "s"
        End If
        newRaveIpadWow = CheckFlag(sh, i)

        'MAIN info
        If InStr(1, newRaveIpadWow(3), "yes") > 0 Then
            Call FillInMainInfo(sh, shFn, i, j, orderSh)
            fileName = Split(shFn.Cells(i, 9), "_")(2)
            aspect = shFn.Cells(i, 17)
            whatIs = "mg"
        ElseIf InStr(1, newRaveIpadWow(3), "tr") > 0 Then
            Call FillInMainInfo(sh, shFn, i, j, orderSh)
            whatIs = "tr"
            fileName = Split(shFn.Cells(i, 9), "_")(2)
            aspect = shFn.Cells(i, 17)
            ' значит заебись
        Else: GoTo positionNext
        End If
        
        'RAVE
        If InStr(1, newRaveIpadWow(0), "new") > 0 Then
            If whatIs = "tr" Then
                'lang
                orderSh.Cells(j, 19) = "АНГЛ"
                'filename
                orderSh.Cells(j, 28) = Replace(Replace(Replace(shFn.Cells(i, 9).value, "#", ""), "_SSS", "TR"), "_DDD", "")
                'sub
                orderSh.Cells(j, 31) = ""
            Else
                'lang
                orderSh.Cells(j, 19) = sh.Cells(i, 22)
                Call FillInFileNames(datet, fileName, orderSh, shFn, i, j, "rave")
            End If
        End If
        
        'IPAD
        If InStr(1, newRaveIpadWow(1), "new") > 0 Then
            If whatIs = "tr" Then
                'lang
                orderSh.Cells(j, 34) = "АНГЛ"
                orderSh.Cells(j, 35) = "KC_" & Right(datet, 2) & Left(datet, 2) & "_" & LCase(Replace(fileName, "#", "")) & "_trl"
            Else
                'lang
                orderSh.Cells(j, 34) = sh.Cells(i, 22)
                Call FillInFileNames(datet, fileName, orderSh, shFn, i, j, "ipad")
            End If
        End If
        
        'WOW
        If InStr(1, newRaveIpadWow(2), "new") > 0 Then
            If whatIs <> "tr" Then
                'lang
                orderSh.Cells(j, 36) = sh.Cells(i, 22)
                Call FillInFileNames(datet, fileName, orderSh, shFn, i, j, "wow")
            End If
        End If
        
        
        j = j + 1
positionNext:
    Next i
    
    ЭтаКнига.Closed
    frmLoad.Unloading
End Sub

Function CheckFlag(ByRef sh As Worksheet, actRow As Long) As Variant

    Dim arrRaveIpadWow(4) As Variant
    Dim flagMgLab As Boolean
    Dim whatIs As String
    flagMgLab = False
    
    'Флажок что ни один не был New
    arrRaveIpadWow(3) = "no"
    
    'Проверка лабы MG Lab
    If InStr(1, LCase(sh.Cells(actRow, 20)), "mg lab") > 0 Then
       flagMgLab = True
       whatIs = "yes"
    ElseIf Not sh.Range("P" & actRow & ":Q" & actRow).Find(What:="new", LookAt:=xlPart, LookIn:=xlValues) Is Nothing And InStr(1, sh.Cells(actRow, 6).value, "New Releases", vbTextCompare) > 0 Then
        flagMgLab = True
        whatIs = "tr"
    Else
        arrRaveIpadWow(0) = 0
        arrRaveIpadWow(1) = 0
        arrRaveIpadWow(2) = 0
        CheckFlag = arrRaveIpadWow
        Exit Function
    End If
    
    'new и 0 - заполнения массива по столбцам для дальнейшего заполнения таблицы по пунктам где new
    'Для Rave проверка
    If InStr(1, LCase(sh.Cells(actRow, 16)), "new") > 0 And flagMgLab Then
        arrRaveIpadWow(0) = "new"
        arrRaveIpadWow(3) = whatIs
    Else: arrRaveIpadWow(0) = 0
    End If
    'Для ipad проверка
    If InStr(1, LCase(sh.Cells(actRow, 17)), "new") > 0 And flagMgLab Then
        arrRaveIpadWow(1) = "new"
        arrRaveIpadWow(3) = whatIs
    Else: arrRaveIpadWow(1) = 0
    End If
     'Для Wow проверка
    If InStr(1, LCase(sh.Cells(actRow, 18)), "new") > 0 And flagMgLab Then
        arrRaveIpadWow(2) = "new"
        arrRaveIpadWow(3) = whatIs
    Else: arrRaveIpadWow(2) = 0
    End If
    
    CheckFlag = arrRaveIpadWow
End Function

Sub FillInMainInfo(ByRef sh As Worksheet, ByRef shFn As Worksheet, actRow As Long, ordRow As Long, orderSh As Worksheet)
    'M/TV
        orderSh.Range("A" & ordRow) = sh.Range("A" & actRow)
    'Year
        orderSh.Range("B" & ordRow) = sh.Range("H" & actRow)
    'Title in Rus
        orderSh.Range("C" & ordRow) = sh.Range("K" & actRow)
    'Title in Eng
        orderSh.Range("D" & ordRow) = shFn.Range("A" & actRow)
    'Series
        orderSh.Range("E" & ordRow) = shFn.Cells(actRow, 2)
    'Aspect
        orderSh.Range("K" & ordRow) = shFn.Cells(actRow, 16)
    'Studio
        orderSh.Range("L" & ordRow) = sh.Range("S" & actRow)
    'Chrono
        orderSh.Range("M" & ordRow) = sh.Range("AH" & actRow)
    'Type work
        orderSh.Range("L" & ordRow) = sh.Range("Y" & actRow)
    'Type work руфилмс
        orderSh.Range("N" & ordRow) = sh.Range("Z" & actRow)
    'Type sub
        orderSh.Range("Q" & ordRow) = sh.Range("X" & actRow)
End Sub

Function RemoveSymbolAndCollectSentens(MorTV As String, title As String, Optional season As String, Optional episod As String) As String
    'RemoveSpecialSymbols
        Const SpecialCharacters As String = "|’|!|?|:|,|'|.|-|…|@|#|$|%|^|&|*|(|)|{|[|]|/|\|}"
        Dim charr As Variant
    If InStr(1, LCase(MorTV), "tv") And UBound(Split(title, ".")) > 0 Then
        season = Replace(Trim(Split(title, ".")(1)), " ", "_") & "_"
        title = Replace(Trim(Split(title, ".")(0)), " ", "_") & "_"
    End If
    For Each charr In Split(SpecialCharacters, "|")
        title = Replace(Trim(Replace(title, charr, "")), " ", "_")
        season = Replace(season, charr, "")
        If Len(episod) > 0 Then
            episod = Replace(Trim(Replace(episod, charr, "")), " ", "_")
        ElseIf Len(season) > 0 Then
            season = Left(season, Len(season) - 1)
        End If
    Next
    If Len(episod) > 0 Then
            episod = "_" & episod
        End If
    
    'title = Replace(title, " ", "_")
    'season = Replace(season, " ", "")
    'episod = Replace(episod, " ", "_")
    
    RemoveSymbolAndCollectSentens = title & season & episod
        
End Function

Sub FillInFileNames(datet As String, fileName As String, ordSh As Worksheet, _
                    draftSh As Worksheet, initRow As Long, actRow As Long, typeS As String)
    
    If typeS = "rave" Then
        Dim langString As New StringBuilderMy
        Dim oneLang As String
        Dim colNumb As Long
        Dim k As Long
        Dim subArr As Variant
        
        If InStr(1, draftSh.Cells(initRow, 43), "|", vbTextCompare) > 0 Then
            subArr = Split(draftSh.Cells(initRow, 43), "|")
        ElseIf Len(draftSh.Cells(initRow, 43)) > 0 Then
            Set subArr = draftSh.Cells(initRow, 43)
        Else: GoTo dubcont
        End If
       
        'Собираем языки Sub для rave
        Dim el As Variant
        For Each el In subArr
            langString.Append (Replace(Replace(Replace(draftSh.Cells(initRow, 9).value, "#", ""), "_DDD", ""), "SSS", Replace(el, " -DYN Sub", "")) & Chr(10))
        Next
            ordSh.Cells(actRow, 31) = Left(langString.ToString, Len(langString.ToString) - 1)
        langString.Clear
        
dubcont:
         'Собираем языки Dub для rave
        langString.Append (Replace(draftSh.Cells(initRow, 18).value, "|", ""))
        For k = 0 To 5
            If IsEmpty(draftSh.Cells(initRow, 37 + k).value) Then
                Exit For
            Else: langString.Append (draftSh.Cells(initRow, 37 + k).value)
            End If
        Next k
        
        'Вставляем в заказ
        ordSh.Cells(actRow, 28) = Replace(Replace(Replace(draftSh.Cells(initRow, 9).value, "#", ""), "_SSS", ""), "DDD", langString.ToString)
        
        langString.Clear
        
    ElseIf typeS = "ipad" Then
        ordSh.Cells(actRow, 35) = "KC_" & Right(datet, 2) & Left(datet, 2) & "_" & _
                                                           LCase(Replace(fileName, "#", ""))
    ElseIf typeS = "wow" Then
        colNumb = 19
        Dim colWow As Long
        colWow = 37
        'Собираем языки Dub для Wow и вставляем
        For k = 0 To 9
            If IsEmpty(draftSh.Cells(initRow, colNumb + k)) And colNumb = 19 Then
                Exit For
            ElseIf k = 0 Then
                oneLang = ChangeLangForWow(draftSh.Cells(initRow, colNumb + k))
                ordSh.Cells(actRow, colWow) = "KC_" & datet & "-W-N-" & Replace(fileName, "#", "_") & "-" & Left(oneLang, 2)
                colWow = colWow + 3
            Else
                oneLang = ChangeLangForWow(draftSh.Cells(initRow, colNumb + k))
                ordSh.Cells(actRow, colWow) = "KC_" & datet & "-W-N-" & Replace(fileName, "#", "_") & "-DUB_" & Left(oneLang, 2)
                colWow = colWow + 3
            End If
        Next k
        
        
        langString.Clear
        
        'Собираем языки Sub для Wow
        colNumb = 29
        
        colWow = 55
        For k = 0 To 5
            If IsEmpty(draftSh.Cells(initRow, colNumb + k)) Then
                Exit For
            Else
                oneLang = ChangeLangForWow(Left(Replace(draftSh.Cells(initRow, colNumb + k), " -DYN Sub", ""), 2))
                ordSh.Cells(actRow, colWow) = "KC_" & datet & "-W-N-" & Replace(fileName, "#", "_") & "-" & "SUB_" & oneLang
                colWow = colWow + 3
            End If
        Next k
        
        'Вставляем в заказ
        'ordSh.Range(colDub & actRow) = "KZR_" & datet & "_" & Replace(filename, "_", "") & "_" & langString & "_" & aspect
    End If
End Sub

Function ChangeLangForWow(fileName As String) As String

    Const SpecialCharactersChi As String = "Yu|Cm|Ch"
    Dim charr As Variant
    'Китайский
    For Each charr In Split(SpecialCharactersChi, "|")
        fileName = Replace(fileName, charr, "Zh", , , vbTextCompare)
    Next
    
    'Казахский
        fileName = Replace(fileName, "Ka", "Kk", , , vbTextCompare)
        
    'Испанский
        fileName = Replace(fileName, "Sp", "Es", , , vbTextCompare)
        
    ChangeLangForWow = fileName
End Function

Function MakeLangRaveWow(fileName As String, ByRef draftSh As Workbook, ByRef ordSh As Workbook, _
                         initRow As Long, actRow As Long, typeS As String, datet As String) As String
    Dim langString As String
    Dim colNumb As Long
    Dim k As Long
    Dim colWow As Long
    
    colNumb = 3
    colWow = 29
    
    For k = 0 To 15
        If IsEmpty(draftSh.Cells(initRow, colNumb + k)) And colNumb = 3 Then
            colNumb = 11
            k = 9
        ElseIf IsEmpty(draftSh.Cells(initRow, colNumb + k)) And colNumb = 11 Then
            k = 16
        Else:
            If typeS = "wow" Then
                If k = 0 Then
                    ordSh.Cells(actRow, colWow) = "KC_" & datet & "-W-N-" & fileName & "-" & langString
                ElseIf colNumb = 3 Then
                    langString = ChangeLangForWow(draftSh.Cells(initRow, colNumb + k))
                    ordSh.Cells(actRow, colWow) = "KC_" & datet & "-W-N-" & fileName & "_" & "DUB_" & langString
                Else: k = 16
                End If
                colWow = colWow + 3
            Else:
                langString = langString + draftSh.Cells(initRow, colNumb + k)
            End If
        End If
    Next k
    MakeLangRaveWow = langString
End Function
