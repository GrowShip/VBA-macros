Attribute VB_Name = "Main"
Option Explicit

Sub FillingInLab(MmYy As String, laba As String)
    Dim sh As String
    Dim i As Long, j As Long, k As Long, totalRow As Long, rowList As Long
    Dim shIn As Worksheet
    Dim system, typi, datet
    
    sh = laba
    
    Set shIn = ThisWorkbook.Sheets("Initial")
    datet = Split(MmYy, "|")
    system = Array("ex3", "exW", "Jetpack IFE")
    typi = Array("m", "s")
    totalRow = MyFunct.countRows(shIn.name)
    rowList = 3
    For i = 1 To 2 'Цикл 1-фильмы, 2 - не фильмы
        For j = 0 To 2 'Цикл 1-ex3, 2-exW, 3-ipad
            For k = 3 To totalRow
                If InStr(1, shIn.Cells(k, 11 + j).value, "new", vbTextCompare) > 0 And _
                   InStr(1, shIn.Cells(k, 15).value, laba, vbTextCompare) > 0 And _
                   ((i = 1 And StrComp(shIn.Cells(k, 1).value, "movie", vbTextCompare) = 0) Or _
                   (i = 2 And StrComp(shIn.Cells(k, 1).value, "movie", vbTextCompare) <> 0)) Then
                    Call FillMainInfo(sh, rowList, k, system(j))
                    Call FillFileNames(sh, rowList, k, system(j), typi(i - 1), datet(0), datet(1), ThisWorkbook.Sheets("Filenames").Cells(k, 3).value)
                    ThisWorkbook.Sheets(sh).rows(rowList).RowHeight = "30"
                    rowList = rowList + 1
                End If
            Next k
        Next j
    Next i
    MsgBox "Done"
End Sub

Private Sub FillMainInfo(laba As String, row As Long, rowfn As Long, ByVal systemS As String)
    Dim airl As Long, system As Long, po As Long, mov As Long, seas As Long, epis As Long, ver As Long, rt As Long, typi As Long, _
        dist As Long, dubAll As Long, dub1 As Long, dub2 As Long, dub3 As Long, dub4 As Long, dub5 As Long, dub6 As Long, dub7 As Long, _
        sub1 As Long, sub2 As Long, sub3 As Long, embsub As Long, year As Long, fformat As Long, aspect As Long, bitRate As Long, fname As Long, _
        delivMeth As Long, ShipTo As Long, clcell As Long, countCell As Long
        
    Dim sh As Worksheet, shFn As Worksheet, shIn As Worksheet
    Dim airline As String
    
    Set shFn = ThisWorkbook.Sheets("Filenames")
    Set shIn = ThisWorkbook.Sheets("Initial")
    clcell = 54
    Select Case laba
        Case "Above"
            Set sh = ThisWorkbook.Sheets("Above")
            airl = 5
            sh.Cells(row, airl) = "UX (Air Europa)"
            system = 6
            po = 7
            mov = 9
            seas = 11
            epis = 10
            ver = 13
            rt = 15
            dist = 16
            dubAll = 54
            dub1 = 17
            dub2 = 18
            dub3 = 19
            dub4 = 20
            dub5 = 21
            dub6 = 22
            dub7 = 23
            sub1 = 24
            sub2 = 25
            sub3 = 26
            embsub = 27
            year = 54
            fformat = 28
            aspect = 29
            bitRate = 30
            fname = 32
            delivMeth = 34
            ShipTo = 35
            typi = 54
            countCell = 4
        Case "CMI"
            Set sh = ThisWorkbook.Sheets("CMI")
            With sh
                .Cells(row, 6).Interior.Color = RGB(252, 228, 214)
                .Cells(row, 10).Interior.Color = RGB(217, 225, 242)
                .Cells(row, 21).Interior.Color = RGB(226, 239, 218)
                .Cells(row, 22).Interior.Color = RGB(255, 242, 204)
            End With
            airl = 7
            sh.Cells(row, airl) = "UX"
            sh.Cells(row, 13).value = "CMI" 'laba
            sh.Cells(row, 24).value = "Digital IFE Services Limited" 'Bill to
            system = 8
            po = 4
            mov = 10
            seas = 54
            epis = 11
            ver = 16
            rt = 15
            dist = 14
            dubAll = 17
            dub1 = 54
            dub2 = 54
            dub3 = 54
            dub4 = 54
            dub5 = 54
            dub6 = 54
            dub7 = 54
            sub1 = 54
            sub2 = 54
            sub3 = 54
            embsub = 54
            year = 12
            fformat = 19
            aspect = 18
            bitRate = 20
            fname = 21
            delivMeth = 54
            ShipTo = 23
            typi = 9
            countCell = 5
        Case "Lab.Aero"
            Set sh = ThisWorkbook.Sheets("Lab.Aero")
            With sh
                .Cells(row, 5).Interior.Color = RGB(252, 228, 214)
                .Cells(row, 7).Interior.Color = RGB(217, 225, 242)
                .Cells(row, 19).Interior.Color = RGB(226, 239, 218)
            End With
            airl = 4
            sh.Cells(row, airl) = "UX"
            sh.Cells(row, 10).value = "Lab.Aero" 'laba
            sh.Cells(row, 21).value = "Digital IFE Services Limited" 'Bill to
            system = 5
            po = 1
            mov = 7
            seas = 54
            epis = 8
            ver = 13
            rt = 12
            dist = 11
            dubAll = 14
            dub1 = 54
            dub2 = 54
            dub3 = 54
            dub4 = 54
            dub5 = 54
            dub6 = 54
            dub7 = 54
            sub1 = 54
            sub2 = 54
            sub3 = 54
            embsub = 54
            year = 9
            fformat = 16
            aspect = 15
            bitRate = 17
            fname = 18
            delivMeth = 54
            ShipTo = 20
            typi = 6
            countCell = 2
        Case "The Hub"
            Set sh = ThisWorkbook.Sheets("The Hub")
            With sh
                .Cells(row, 6).Interior.Color = RGB(252, 228, 214)
                .Cells(row, 8).Interior.Color = RGB(217, 225, 242)
                .Cells(row, 19).Interior.Color = RGB(226, 239, 218)
                .Cells(row, 20).Interior.Color = RGB(255, 242, 204)
            End With
            airl = 5
            sh.Cells(row, airl) = "UX"
            sh.Cells(row, 11).value = "The Hub" 'laba
            sh.Cells(row, 22).value = "Digital IFE Services Limited" 'Bill to
            system = 6
            po = 2
            mov = 8
            seas = 54
            epis = 9
            ver = 14
            rt = 13
            dist = 12
            dubAll = 15
            dub1 = 54
            dub2 = 54
            dub3 = 54
            dub4 = 54
            dub5 = 54
            dub6 = 54
            dub7 = 54
            sub1 = 54
            sub2 = 54
            sub3 = 54
            embsub = 54
            year = 10
            fformat = 17
            aspect = 16
            bitRate = 18
            fname = 19
            delivMeth = 54
            ShipTo = 21
            typi = 7
            countCell = 3
        Case "West"
            Set sh = ThisWorkbook.Sheets("West")
            With sh
                .Cells(row, 7).Interior.Color = RGB(252, 228, 214)
                .Cells(row, 70).Interior.Color = RGB(226, 239, 218)
            End With
            airl = 1
            sh.Cells(row, airl) = "Air Europe"
            sh.Cells(row, 34).value = "West Entertainment" 'laba
            sh.Cells(row, 154).value = "Digital IFE Services Limited" 'Bill to
            system = 7
            po = 77
            mov = 28
            seas = 160
            epis = 160
            ver = 35
            rt = 36
            dist = 33
            dubAll = 160
            dub1 = 41
            dub2 = 42
            dub3 = 43
            dub4 = 44
            dub5 = 45
            dub6 = 46
            dub7 = 47
            sub1 = 55
            sub2 = 56
            sub3 = 57
            embsub = 53
            year = 39
            fformat = 61
            aspect = 59
            bitRate = 63
            fname = 69
            delivMeth = 160
            ShipTo = 76
            typi = 160
            countCell = 160
            clcell = 160
    End Select
    
    With sh
        .Cells(row, system).value = systemS
        .Cells(row, po).value = shIn.Cells(rowfn, 48).value
        .Cells(row, mov).value = MyFunct.GiveMeTitle(shIn.Cells(rowfn, 8).value)
        .Cells(row, seas).value = MyFunct.GiveMeSeason(shIn.Cells(rowfn, 8).value)
        .Cells(row, epis).value = shIn.Cells(rowfn, 9).value
        .Cells(row, ver).value = shFn.Cells(rowfn, 17).value
        .Cells(row, rt).value = shIn.Cells(rowfn, 18).value
        .Cells(row, dist).value = shIn.Cells(rowfn, 14).value
        .Cells(row, year).value = shIn.Cells(rowfn, 4).value
        .Cells(row, dubAll).value = shFn.Cells(rowfn, 11).value
        .Cells(row, dub1).value = shFn.Cells(rowfn, 19).value
        .Cells(row, dub2).value = shFn.Cells(rowfn, 20).value
        .Cells(row, dub3).value = shFn.Cells(rowfn, 21).value
        .Cells(row, dub4).value = shFn.Cells(rowfn, 22).value
        .Cells(row, dub5).value = shFn.Cells(rowfn, 23).value
        .Cells(row, dub6).value = shFn.Cells(rowfn, 24).value
        .Cells(row, dub7).value = shFn.Cells(rowfn, 25).value
        .Cells(row, sub1).value = shFn.Cells(rowfn, 29).value
        .Cells(row, sub2).value = shFn.Cells(rowfn, 30).value
        .Cells(row, sub3).value = shFn.Cells(rowfn, 31).value
        .Cells(row, embsub).value = shFn.Cells(rowfn, 37).value
        .Cells(row, aspect).value = shFn.Cells(rowfn, 16).value
        .Cells(row, typi).value = shIn.Cells(rowfn, 1).value
        .Cells(row, countCell).value = row - 2
        .Cells(row, clcell).Clear
    End With
    Call AboveOrWestCC(row, rowfn, laba)
End Sub

Private Sub FillFileNames(laba As String, row As Long, rowfn As Long, ByVal system As String, _
                          ByVal typi As String, ByVal mm As String, ByVal yy As String, count As String)
    Dim fformat, bitRate, fname, delivMethod, ShipTo As String
    Dim i As Long
    'ff, bR, fN, dM, sT
    Dim columnCase
    i = 54
    Select Case (laba)
        Case "Above"
                              'ff, bR, fN, dM, sT
            columnCase = Array(28, 30, 32, 34, 35)
        Case "CMI"
                               'ff, bR, fN, dM, sT
            columnCase = Array(19, 20, 21, i, 23)
        Case "Lab.Aero"
                               'ff, bR, fN, dM, sT
            columnCase = Array(16, 17, 18, i, 20)
        Case "The Hub"
                               'ff, bR, fN, dM, sT
            columnCase = Array(17, 18, 19, i, 21)
        Case "West"
            i = 160
                               'ff, bR, fN, dM, sT
            columnCase = Array(61, 63, 69, 6, 76)
    End Select
    Select Case (system)
        Case "ex3"
            fformat = "Mpeg 4"
            bitRate = "1.5"
            fname = "m4"
            delivMethod = "Panasonic"
            ShipTo = "Panasonic"
        Case "exW"
            fformat = "Mpeg 4"
            bitRate = "800"
            fname = "z4"
            delivMethod = "Panasonic"
            ShipTo = "Panasonic"
        Case "Jetpack IFE"
            fformat = "h.265 codec in an m4v container"
            bitRate = "VBR, Aiming for no more then 2000."
            delivMethod = "Aspera"
            ShipTo = "Jetpack IFE"
    End Select
    If StrComp(system, "Jetpack IFE", vbTextCompare) = 0 Then
        fname = "UX_" & MyFunct.RemoveSymbols(GiveMeTitle(ThisWorkbook.Sheets("Initial").Cells(rowfn, 8).value) & "|", "") & "_Ep" & _
                ThisWorkbook.Sheets("Filenames").Cells(rowfn, 6).value & "_" & mm & yy & "_" & _
                ThisWorkbook.Sheets("Filenames").Cells(rowfn, 18).value & ".m4v"
        GoTo IPADDS
    End If
    fname = "ux" & typi & mm & yy & count & fname & ".mpg"
IPADDS:
    With ThisWorkbook.Sheets(laba)
        .Cells(row, columnCase(0)).value = fformat
        .Cells(row, columnCase(1)).NumberFormat = "@"
        .Cells(row, columnCase(1)).value = bitRate
        .Cells(row, columnCase(2)).value = fname & FillFileNamesSub(rowfn, system, typi, mm, yy, count)
        .Cells(row, columnCase(3)).value = delivMethod
        .Cells(row, columnCase(4)).value = ShipTo
        .Cells(row, i).Clear
    End With
    If StrComp(laba, "Above", vbTextCompare) = 0 Then ThisWorkbook.Sheets(laba).Cells(row, columnCase(3)).value = "SmartJog"
End Sub

Private Function FillFileNamesSub(rowfn As Long, system As String, typi As String, mm As String, yy As String, count As String) As String
    Dim leftStr As String, rightStr As String, result As String
    Dim i As Long
    result = ""
    Select Case (system)
        Case "ex3"
            leftStr = "ux" & typi & mm & yy & count & "m4"
            rightStr = ".zip"
        Case "exW"
            leftStr = "ux" & typi & mm & yy & count & "z4"
            rightStr = ".zip"
        Case "Jetpack IFE"
            FillFileNamesSub = result
            Exit Function
    End Select
    
    For i = 0 To 4
        If Len(ThisWorkbook.Sheets("Filenames").Cells(rowfn, 29 + i).value) > 1 Then
            result = result & Chr(10) & leftStr & "_" & LCase(Left(ThisWorkbook.Sheets("Filenames").Cells(rowfn, 29 + i).value, 3)) & "_sub" & rightStr
        ElseIf Len(ThisWorkbook.Sheets("Filenames").Cells(rowfn, 36).value) > 1 Then
            FillFileNamesSub = result & Chr(10) & leftStr & "_" & LCase(Left(ThisWorkbook.Sheets("Filenames").Cells(rowfn, 36).value, 3)) & "_cap" & rightStr
            Exit Function
        Else
            FillFileNamesSub = result
            Exit Function
        End If
    Next i
End Function

Private Sub AboveOrWestCC(row As Long, rowfn As Long, laba As String)
    Dim k, i, colu
    If StrComp(laba, "above", vbTextCompare) = 0 Then
        k = 2
        colu = 24
    ElseIf StrComp(laba, "west", vbTextCompare) = 0 Then
        k = 3
        colu = 55
    Else: Exit Sub
    End If
    
    For i = 0 To k
        If Len(ThisWorkbook.Sheets(laba).Cells(row, colu + i).value) = 0 Then
            ThisWorkbook.Sheets(laba).Cells(row, colu + i).value = Left(ThisWorkbook.Sheets("Filenames").Cells(rowfn, 36).value, 5)
            Exit Sub
        End If
    Next i
End Sub

