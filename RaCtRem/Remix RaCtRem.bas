Attribute VB_Name = "Remix"
Option Explicit

Public Sub CreateCTR()
    Dim ask As Long: ask = frmDidYouUpload.Loading("—оздаем [RemixUpload] = [Media] и [GuiREMIXupload] = [MediaGuiLangAttr]?")
    
    If ask = 2 Then
        Exit Sub
    ElseIf ask = 6 Then
        CreateRemixAndGui
    ElseIf ask = 7 Then
        Exit Sub
    End If
End Sub

Private Sub CreateRemixAndGui()
    frmLoad.Loading
    Ёта нига.Opened
    
    CreateRemixAndGuiInLock
    
    Dim rowsCtr As Long
    rowsCtr = MyFunct.countRowBest("REMIXupload")
    If rowsCtr > 1 Then clearing.ClearREMIXupload (False)
    rowsCtr = MyFunct.countRowBest("GuiREMIXupload")
    If rowsCtr > 1 Then clearing.ClearGuiREMIXupload (False)
    
    Inserting.CopyRemixLockToUpload (False)
    Inserting.CopyGuiLockToUpload (False)
    
    OpeningBook.SetBookConfig
    
    clearing.ClearREMIXlock (False)
    clearing.ClearGuiREMIXlock (False)
    
    Ёта нига.Closed
    frmLoad.Unloading
End Sub


Public Sub CreateRemixAndGuiInLock()
    Dim shCTRlock As Worksheet: Set shCTRlock = ThisWorkbook.Sheets("CTRlock")
    Dim shFilenames As Worksheet: Set shFilenames = ThisWorkbook.Sheets("Filenames")
    Dim shRemix As Worksheet: Set shRemix = ThisWorkbook.Sheets("REMIXlock")
    Dim guiSh As Worksheet: Set guiSh = ThisWorkbook.Sheets("GuiREMIXlock")
    Dim rowsCtr As Long
    Dim i As Long
    Dim rowKC As Long
    
    rowsCtr = MyFunct.countRowBest(guiSh.name)
    If rowsCtr > 1 Then clearing.ClearGuiREMIXlock (False)
    rowsCtr = MyFunct.countRowBest(shRemix.name)
    If rowsCtr > 1 Then clearing.ClearREMIXlock (False)
    rowsCtr = MyFunct.countRowBest(shCTRlock.name)
    If rowsCtr > 1 Then clearing.ClearCTRlock (False)
    
    
    Call CTRfirst.FillingCTRlock("New")
    
    rowsCtr = MyFunct.countRowBest(shCTRlock.name)
    For i = 2 To rowsCtr
        rowKC = shFilenames.Range("J:J").Find(What:=shCTRlock.Cells(i, 40).value, LookIn:=xlValues, LookAt:=xlWhole).row
        shCTRlock.Cells(i, 39).value = rowKC
        Call RemixInserting(shCTRlock, shRemix, rowKC, i)
        Call GuiInserting(shCTRlock, CInt(rowKC), i)
    Next i
    
End Sub

Private Sub RemixInserting(ByRef sheetFrom As Worksheet, ByRef sheetTo As Worksheet, rowKC As Long, rowTo As Long)
    With sheetTo
        .Cells(rowTo, 2) = Replace(sheetFrom.Cells(rowTo, 30).value, ".mp4", "") 'parent
        .Cells(rowTo, 3) = sheetFrom.Cells(rowTo, 10).value 'startDate
        .Cells(rowTo, 4) = sheetFrom.Cells(rowTo, 11).value 'endDate
        .Cells(rowTo, 5) = "VIDEO" 'mediaType
        .Cells(rowTo, 6) = sheetFrom.Cells(rowTo, 6).value 'MediaCat
        .Cells(rowTo, 8) = ThisWorkbook.Sheets("Filenames").Cells(rowKC, 14) 'rating
        .Cells(rowTo, 9) = GetPatentlock(ThisWorkbook.Sheets("Filenames").Cells(rowKC, 14)) 'parentlock
        .Cells(rowTo, 12) = "No" 'collection
        .Cells(rowTo, 15) = sheetFrom.Cells(rowTo, 30).value 'mediafile
        .Cells(rowTo, 21) = "Remix" 'platform
        .Cells(rowTo, 25) = ThisWorkbook.Sheets("Filenames").Cells(rowKC, 16) 'MF widescreen
        .Cells(rowTo, 26) = "4" 'MF MPEG format
        .Cells(rowTo, 27) = "" '?
        .Cells(rowTo, 28) = "" '?
        .Cells(rowTo, 29) = sheetFrom.Cells(rowTo, 7).value 'runtime
        .Cells(rowTo, 30) = "Yes" 'encrypted
        .Cells(rowTo, 31) = sheetFrom.Cells(rowTo, 36).value 'studia
        .Cells(rowTo, 32) = sheetFrom.Cells(rowTo, 37).value 'lab
        .Cells(rowTo, 36) = GetCorrectGenre(ThisWorkbook.Sheets("Initial").Cells(rowKC, 6), rowKC) 'genre
        .Cells(rowTo, 42) = "No" 'Featured
        .Cells(rowTo, 44) = Replace(sheetFrom.Cells(rowTo, 30).value, "mp4", "png") 'image
        .Cells(rowTo, 47) = ThisWorkbook.Sheets("Filenames").Cells(rowKC, 15) 'year
        .Cells(rowTo, 51) = sheetFrom.Cells(rowTo, 3).value 'season
        .Cells(rowTo, 52) = sheetFrom.Cells(rowTo, 4).value 'episode
        .Cells(rowTo, 53) = ThisWorkbook.Sheets("Filenames").Cells(rowKC, 44) 'boxType
        .Cells(rowTo, 56) = GetRightApps(ThisWorkbook.Sheets("Initial").Cells(rowKC, 6), rowKC) 'Apps
    End With
End Sub

Private Function GetPatentlock(reating As String) As String
    If StrComp(Trim(reating), "R", vbTextCompare) = 0 Then
        GetPatentlock = "Locked"
    Else: GetPatentlock = "Unlocked"
    End If
End Function

Private Function GetCorrectGenre(genres As String, rowKC As Long) As String
    If InStr(1, ThisWorkbook.Sheets("Initial").Cells(rowKC, 3).value, "movies", vbTextCompare) > 0 Then
        genres = Replace(genres, "Movies", "Kids Movies", , , vbTextCompare)
    End If
    GetCorrectGenre = Replace(genres, ", ", " | ", , , vbTextCompare)
End Function

Private Function GetRightApps(apps As String, rowKC As Long) As String
    With ThisWorkbook.Sheets("Initial")
    If InStr(1, .Cells(rowKC, 3).value, "kids", vbTextCompare) > 0 And _
       InStr(1, .Cells(rowKC, 1).value, "TV", vbTextCompare) > 0 Then
        GetRightApps = "Kids_TV"
    ElseIf InStr(1, .Cells(rowKC, 3).value, "Discover Kazakhstan", vbTextCompare) > 0 Then
        GetRightApps = "Discover Kazakhstan"
    ElseIf InStr(1, .Cells(rowKC, 3).value, "kids", vbTextCompare) > 0 And _
           InStr(1, .Cells(rowKC, 1).value, "Movie", vbTextCompare) > 0 Then
        GetRightApps = "MOVIE | Kids_Movies"
    ElseIf StrComp(.Cells(rowKC, 3).value, "Movies", vbTextCompare) = 0 Then
        GetRightApps = "MOVIE"
    ElseIf InStr(1, .Cells(rowKC, 3).value, "series", vbTextCompare) > 0 And _
       InStr(1, .Cells(rowKC, 1).value, "TV", vbTextCompare) > 0 Then
        GetRightApps = "Series and TV"
    Else: GetRightApps = "ATTENTION!"
    End If
    End With
End Function

'ByRef sheetFrom As Worksheet, rowFrom As Long

Private Sub GuiInserting(CTRsheet As Worksheet, rowff As Integer, rowIn As Long)
    Dim sortedLang As New Dictionary, additionalLang As New Dictionary
    Dim langArr As Variant
    Dim initSh As Worksheet, guiSh As Worksheet
    Set initSh = ThisWorkbook.Sheets("Initial")
    Set guiSh = ThisWorkbook.Sheets("GuiREMIXlock")
    Dim indexDub As Long, indexSub As Long, looper As Long
    Dim dubRange As Range, subRange As Range, findedDub As Range, findedSub As Range
    Dim title As String: title = Replace(CTRsheet.Cells(rowIn, 30), ".mp4", "")
    Dim ffSheet As Worksheet: Set ffSheet = ThisWorkbook.Sheets("Filenames")
    Dim additionalAdded As Boolean
    
    With sortedLang
        .Add Key:="Eng", Item:=0
        .Add Key:="Rus", Item:=1
        .Add Key:="Kaz", Item:=2
        .Add Key:="Deu", Item:=3
        .Add Key:="Fra", Item:=4
        .Add Key:="Ita", Item:=5
        .Add Key:="Spa", Item:=6
        .Add Key:="Dau", Item:=7
        .Add Key:="Por", Item:=8
        .Add Key:="Tha", Item:=9
        .Add Key:="Hin", Item:=10
        .Add Key:="Ara", Item:=11
        .Add Key:="Tur", Item:=12
        .Add Key:="Kor", Item:=13
        .Add Key:="Jpn", Item:=14
        .Add Key:="Zho", Item:=15
        .Add Key:="Chi", Item:=16
    End With
    
    langArr = sortedLang.Keys
    
    Dim countLang As Long: countLang = WorksheetFunction.CountA(ffSheet.Range("S" & rowff & ":AH" & rowff))
    'ffSheet.Range("S" & rowff & ":AH" & rowff).Cells.SpecialCells(xlCellTypeConstants).count
    Dim countLast As Long
    Dim langua As String
    
    Set dubRange = ffSheet.Range("S" & rowff & ":AB" & rowff)
    Set subRange = ffSheet.Range("AC" & rowff & ":AH" & rowff)
    
    
    Do While (indexDub + indexSub < countLang Or looper < 3)
        additionalAdded = False
        
        If Len(ffSheet.Cells(rowff, 19 + indexDub)) > 0 And Not sortedLang.Exists(Left(ffSheet.Cells(rowff, 19 + indexDub), 3)) Then
            Call AddAdditionalLang(sortedLang, indexDub, ffSheet, rowff, "dub")
            additionalAdded = True
        End If
        If Len(ffSheet.Cells(rowff, 29 + indexSub)) > 0 And Not sortedLang.Exists(Left(ffSheet.Cells(rowff, 29 + indexSub), 3)) Then
            Call AddAdditionalLang(sortedLang, indexSub, ffSheet, rowff, "sub")
            additionalAdded = True
        End If
        If looper = sortedLang.count - 1 And (ffSheet.Cells(rowff, 35)) > 0 And Not sortedLang.Exists("dvs") Then
            sortedLang.Add Key:="dvs", Item:="-1"
            additionalAdded = True
            countLang = countLang + 1
        End If
        
        If additionalAdded Then langArr = sortedLang.Keys
        
        countLast = MyFunct.countRowBest("GuiREMIXlock") + 1
        langua = langArr(looper)
              
        Set findedDub = dubRange.Find(What:=langArr(looper), LookIn:=xlValues, LookAt:=xlPart)
        Set findedSub = subRange.Find(What:=langArr(looper), LookIn:=xlValues, LookAt:=xlPart)
        
        If langArr(looper) = "Eng" Or langArr(looper) = "Rus" Or langArr(looper) = "Kaz" Then
            Call InsertThreeMainLangInGui(guiSh, countLast, CStr(langArr(looper)), ffSheet, initSh, rowff, title)
        End If
        
         If looper = 1 And InStr(1, ffSheet.Cells(rowff, 35), "RusAD", vbTextCompare) > 0 Then
            countLang = countLang + 1
            indexDub = indexDub + 1
            guiSh.Cells(countLast, 10).value = indexDub
        End If
        
        If Not findedDub Is Nothing Then
            If findedDub.Column = 19 + indexDub Then
                indexDub = indexDub + 1
                guiSh.Cells(countLast, 10).value = indexDub
            End If
        End If
        If Not findedSub Is Nothing Then
            If findedSub.Column = 29 + indexSub Then
                guiSh.Cells(countLast, 6).value = Replace(Replace(Replace(ffSheet.Cells(rowff, 9), "#", ""), "_DDD", ".srt"), "SSS", Replace(findedSub.text, " -DYN Sub", ""))
                indexSub = indexSub + 1
            End If
        End If
        
        If Not findedDub Is Nothing Or Not findedSub Is Nothing Or langArr(looper) = "dvs" Then
            If langArr(looper) = "dvs" Then
                indexDub = indexDub + 1
                guiSh.Cells(countLast, 10).value = indexDub
            End If
            With guiSh
                .Cells(countLast, 1) = title 'parent
                .Cells(countLast, 4) = LCase(langArr(looper)) 'lang
                .Cells(countLast, 18) = title & " " & LCase(langArr(looper))
            End With
        End If
        
        looper = looper + 1
    Loop
    With guiSh.Range("A" & countLast & ":O" & countLast).Borders(xlEdgeBottom)
        .LineStyle = XlLineStyle.xlContinuous
        .Color = vbBlack
    End With
End Sub

Private Sub AddAdditionalLang(ByRef additionalLang As Dictionary, ByRef indexer As Long, ffSheet As Worksheet, rowff As Integer, whatType As String)
    Dim startFr As Long
    If whatType = "dub" Then
        startFr = 19
    ElseIf whatType = "sub" Then
        startFr = 29
    Else: Exit Sub
    End If
    
    If additionalLang.Exists(Left(ffSheet.Cells(rowff, startFr + indexer), 3)) Then
        additionalLang.Item(Left(ffSheet.Cells(rowff, startFr + indexer), 3)) = "dub sub"
        'indexer = indexer + 1
    Else
        additionalLang.Item(Left(ffSheet.Cells(rowff, startFr + indexer), 3)) = whatType
        'indexer = indexer + 1
    End If
End Sub

Private Sub InsertThreeMainLangInGui(guiSh As Worksheet, rowIn As Long, lang As String, ffSheet As Worksheet, initSh As Worksheet, rowFrom As Integer, titleFf As String)
    Dim title As String, episode As String
    Dim x As Long
    If lang = "Eng" Then
        x = 44
        title = initSh.Cells(rowFrom, 12)
        episode = initSh.Cells(rowFrom, 15)
    ElseIf lang = "Rus" Then
        x = 39
        title = initSh.Cells(rowFrom, 11)
        episode = initSh.Cells(rowFrom, 14)
    ElseIf lang = "Kaz" Then
        x = 50
        title = initSh.Cells(rowFrom, 48)
        episode = initSh.Cells(rowFrom, 13)
    Else: Exit Sub
    End If
    
    If InStr(1, ffSheet.Cells(rowFrom, 36), lang, vbTextCompare) > 0 Then
        guiSh.Cells(rowIn, 6) = Replace(Replace(Replace(ffSheet.Cells(rowFrom, 9), "#", ""), "_DDD", ".srt"), "SSS", ffSheet.Cells(rowFrom, 36).value)
    End If
    
    With guiSh
        .Cells(rowIn, 1) = titleFf 'parent
        .Cells(rowIn, 2) = title 'title
        .Cells(rowIn, 3) = episode 'episode
        .Cells(rowIn, 4) = LCase(lang) 'lang
        .Cells(rowIn, 5) = initSh.Cells(rowFrom, x + 2) 'synop
        .Cells(rowIn, 12) = initSh.Cells(rowFrom, x + 1) 'stars
        .Cells(rowIn, 14) = initSh.Cells(rowFrom, x) 'dir
        .Cells(rowIn, 18) = titleFf & " " & LCase(lang)
        If InStr(1, initSh.Cells(rowFrom, 2).value, "Document", vbTextCompare) Then .Cells(rowIn, 3) = title
    End With
End Sub

