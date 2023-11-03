Attribute VB_Name = "GenerateDate"
Option Explicit
Public userInput1 As Variant
Public userInput2 As Variant
Public flag As Boolean

Sub GenerateDateAbove()
    Call ShowForm
    If Len(userInput1) = 2 And Len(userInput2) = 2 And flag Then
        Call Above.AboveFilling(userInput1 & "|" & userInput2)
    End If
End Sub

Sub GenerateDateCMI()
    Call ShowForm
    If Len(userInput1) = 2 And Len(userInput2) = 2 And flag Then
        Call CMIFilling(userInput1 & "|" & userInput2)
    End If
End Sub

Sub GenerateDateLAbAero()
    Call ShowForm
    If Len(userInput1) = 2 And Len(userInput2) = 2 And flag Then
        Call LabAeroFilling(userInput1 & "|" & userInput2)
    End If
End Sub

Sub GenerateDateTheHub()
    Call ShowForm
    If Len(userInput1) = 2 And Len(userInput2) = 2 And flag Then
        Call TheHubFilling(userInput1 & "|" & userInput2)
    End If
End Sub

Sub GenerateDateWest()
    Call ShowForm
    If Len(userInput1) = 2 And Len(userInput2) = 2 And flag Then
        Call WestFilling(userInput1 & "|" & userInput2)
    End If
End Sub
Sub ShowForm()
    CycleForm.Show ' Display the user form
End Sub


