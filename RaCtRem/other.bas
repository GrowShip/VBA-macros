Private Sub cmdDoIt_Click()
    Dim selectedEvent As String
    selectedEvent = Me.listCTRraveEvents
    If InStr(1, selectedEvent, "1.", vbTextCompare) > 0 Then
        Call CTRfirst.CreateCTR
    ElseIf InStr(1, selectedEvent, "2.", vbTextCompare) > 0 Then
        If InitialFilling.AskText("Íà÷èíàåì ïîèñê îøèáîê â RAVE CTR?") Then Exit Sub
        Call CTRfirst.CompareCTR
    ElseIf InStr(1, selectedEvent, "3.", vbTextCompare) > 0 Then
        Call Remix.CreateCTR
    ElseIf InStr(1, selectedEvent, "4.", vbTextCompare) > 0 Then
        If InitialFilling.AskText("Íà÷èíàåì ïîèñê îøèáîê â RAVE REMIX?") Then Exit Sub
        Call RemixCheck.CompareRemixGui
    ElseIf InStr(1, selectedEvent, "5.", vbTextCompare) > 0 Then
        Call OrderForGoogle.CreateGooOrder
    End If
End Sub
