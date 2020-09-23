Attribute VB_Name = "Message"
Option Explicit
'   Enum for AltisMsgBoxResult
    Public Enum MessageBoxResult
        MsgBoxButton1 = 0
        MsgBoxButton2 = 1
        MsgBoxButton3 = 2
        MsgBoxButton4 = 3
    End Enum
    ' Set Enum for ButtonStyle
    Public Enum ButtonStyle
        MSGAbortRetryIgnore = 1
        MSGDefaultButton1 = 10
        MSGDefaultButton2 = 20
        MSGDefaultButton3 = 30
        MSGDefaultButton4 = 40
        MSGMSGBoxHelpButton = 2
        MSGOKCancel = 3
        MSGOKOnly = 4
        MSGRetryCancel = 5
        MSGYesNo = 6
        MSGYesNoCancel = 7
    End Enum


Public Function MessageBox(Prompt As String, Button As ButtonStyle) As MessageBoxResult
    Dim TheForm As Form
    Dim I As Integer
    ' Disable All Forms
    For Each TheForm In Forms
        DoEvents
        TheForm.Enabled = False
    Next TheForm
    ' Load Message Box
    Load frmMSGBOX
    ' Initial Value
    If Button = 0 Then Button = 4
    ' Set Caption
    frmMSGBOX.lblMessage.Caption = Prompt
    ' Choose ButtonStyle
    With frmMSGBOX
        Select Case Button
            Case 10 To 16
                .Command1(0).Default = True
                Button = Button - 10
            Case 20 To 26
                .Command1(1).Default = True
                Button = Button - 20
            Case 30 To 36
                .Command1(2).Default = True
                Button = Button - 30
            Case 40 To 46
                .Command1(3).Default = True
                Button = Button - 40
        End Select
        Select Case Button
            Case Is = 1
                For I = 0 To 2
                    .Command1(I).Visible = True
                Next I
                .Command1(0).Caption = "Abort"
                .Command1(1).Caption = "Retry"
                .Command1(2).Caption = "Ignore"
            Case Is = 2
                .Command1(3).Visible = True
                .Command1(3).Caption = "Help"
            Case Is = 3
                For I = 0 To 1
                    .Command1(I).Visible = True
                Next I
                .Command1(0).Caption = "OK"
                .Command1(1).Caption = "Cancel"
                .Command1(I).Cancel = True
            Case Is = 4
                .Command1(0).Visible = True
                .Command1(0).Caption = "OK"
                .Command1(0).Default = True
            Case Is = 5
                For I = 0 To 1
                    .Command1(I).Visible = True
                Next I
                .Command1(0).Caption = "Retry"
                .Command1(1).Caption = "Cancel"
                .Command1(1).Cancel = True
            Case Is = 6
                For I = 0 To 1
                    .Command1(I).Visible = True
                Next I
                .Command1(0).Caption = "Yes"
                .Command1(1).Caption = "No"
            Case Is = 7
                For I = 0 To 2
                    .Command1(I).Visible = True
                Next I
                .Command1(0).Caption = "Yes"
                .Command1(1).Caption = "No"
                .Command1(2).Caption = "Cancel"
                .Command1(2).Cancel = True
        End Select
    End With
    frmMSGBOX.Show
    ' Disable Window until user responses
    Do
        DoEvents
    Loop Until frmMSGBOX.Clicked = True
    ' Initial value
    MessageBox = frmMSGBOX.ButtonNumber
    ' Enable All Forms
    For Each TheForm In Forms
        DoEvents
        TheForm.Enabled = True
    Next TheForm
    ' Initial Value
    frmMSGBOX.Clicked = False
End Function

