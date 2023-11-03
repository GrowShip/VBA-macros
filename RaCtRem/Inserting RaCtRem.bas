Attribute VB_Name = "Inserting"
Option Explicit

Private Sub InsertFromTo(frowRange As Range, toRange As Range, Optional ask As Boolean = True)
    If ask Then If InitialFilling.AskText("Точно скопировать из " & frowRange.Worksheet.name & " в " & toRange.Worksheet.name & " ?") Then Exit Sub
    frowRange.Copy Destination:=toRange
End Sub

Public Sub CopyCTRlockTOupload(Optional ask As Boolean = True)
    Dim countFrom As Long: countFrom = MyFunct.countRowBest("CTRlock")
    Dim countTo As Long: countTo = MyFunct.countRowBest("CTRupload") + 1
    
    Call InsertFromTo(ThisWorkbook.Sheets("CTRlock").Range("A2:AK" & countFrom), _
                      ThisWorkbook.Sheets("CTRupload").Range("A" & countTo), ask)
End Sub

Public Sub CopyCTRuploadTOlock(Optional ask As Boolean = True)
    Dim countFrom As Long: countFrom = MyFunct.countRowBest("CTRupload")
    Dim countTo As Long: countTo = MyFunct.countRowBest("CTRlock") + 1
    
    Call InsertFromTo(ThisWorkbook.Sheets("CTRupload").Range("A2:AK" & countFrom), _
                      ThisWorkbook.Sheets("CTRlock").Range("A" & countTo), ask)
End Sub

Public Sub CopyRemoveLockToUpload(Optional ask As Boolean = True)
    Dim countFrom As Long: countFrom = MyFunct.countRowBest("RemoveLock")
    Dim countTo As Long: countTo = MyFunct.countRowBest("RemoveUpload") + 1
    
    Call InsertFromTo(ThisWorkbook.Sheets("RemoveLock").Range("A2:D" & countFrom), _
                      ThisWorkbook.Sheets("RemoveUpload").Range("A" & countTo), ask)
End Sub


Public Sub CopyRemoveUploadToLock(Optional ask As Boolean = True)
    Dim countFrom As Long: countFrom = MyFunct.countRowBest("RemoveUpload")
    Dim countTo As Long: countTo = MyFunct.countRowBest("RemoveLock") + 1
    
    Call InsertFromTo(ThisWorkbook.Sheets("RemoveUpload").Range("A2:D" & countFrom), _
                      ThisWorkbook.Sheets("RemoveLock").Range("A" & countTo), ask)
End Sub

Public Sub CopyRemixLockToUpload(Optional ask As Boolean = True)
    Dim countFrom As Long: countFrom = MyFunct.countRowBest("REMIXlock")
    Dim countTo As Long: countTo = MyFunct.countRowBest("REMIXupload") + 1
    
    Call InsertFromTo(ThisWorkbook.Sheets("REMIXlock").Range("A2:BD" & countFrom), _
                      ThisWorkbook.Sheets("REMIXupload").Range("A" & countTo), ask)
End Sub

Public Sub CopyGuiLockToUpload(Optional ask As Boolean = True)
    Dim countFrom As Long: countFrom = MyFunct.countRowBest("GuiREMIXlock")
    Dim countTo As Long: countTo = MyFunct.countRowBest("GuiREMIXupload") + 1
    
    Call InsertFromTo(ThisWorkbook.Sheets("GuiREMIXlock").Range("A2:O" & countFrom), _
                      ThisWorkbook.Sheets("GuiREMIXupload").Range("A" & countTo), ask)
End Sub

