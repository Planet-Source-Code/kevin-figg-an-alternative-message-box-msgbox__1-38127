Attribute VB_Name = "Message"
' Originally written by............
'
' Modifed and imporved by Kevin Figg
'
' Added by Kevin Figg
'
' 1. Different Icons
' 2. The Excellent Chamelon Button by Gonchuki
' 3. Windows 2000 Style (or not as the case may be) and XP Button style (optional)
' 4. The designer
' 5. Positionable message box
' 6. OKOnly button is centred
' 7. VBModal replication option
' 8. Timed message (Disappears after specified amount of time)
' 9. Scrollbars to message (optional)
'
'
'

Option Explicit

Public ButtonNumber As Integer

'   Enum for AltisMsgBoxResult
    Public Enum MessageBoxResult
        MsgBoxButton1 = 0
        MsgBoxButton2 = 1
        MsgBoxButton3 = 2
        MsgBoxButton4 = 3
    End Enum
    
'Enum for Icon Image
'The reason for not using a image list is because I didn't want to assume the person using this code would have the MS Common Controls
' OCX in their project, and I certainly didn't want them to include another OCX BECAUSE of my app.
    Public Enum Icon
        vbWin32 = 0
        vbInformation = 1
        vbExclamation = 2
        vbCritical = 3
        vbQuestion = 4
        vbComputer = 5
        vbHardware = 6
        vbInternet = 7
        vbPower = 8
        vbSearch = 9
        vbSecurity = 10
        vbSound = 11
        vbUsers = 12
    End Enum
    
    ' Set Enum for ButtonStyle
    Public Enum ButtonStyle
       vbAbortRetryIgnore = 1
       vbDefaultButton1 = 10
       vbDefaultButton2 = 20
       vbDefaultButton3 = 30
       vbDefaultButton4 = 40
       vbMsgBoxHelpButton = 2
       vbOKCancel = 3
       vbOKOnly = 4
       vbRetryCancel = 5
       vbYesNo = 6
       vbYesNoCancel = 7
       vbNoButtons = 8
    End Enum
   
   Public Enum msgResult
        vbtimedout = -1
        vbCancel = 0
        vbOK = 1
        vbRetry = 2
        vbYes = 3
        vbNo = 4
        vbHelp = 5
        vbAbort = 6
        vbIgnore = 7
    End Enum
    Public Enum Modal
        vbNoModal = 0
        vbModal = 1
        'vbSystemModal = 2 : ' Not working Yet
    End Enum
    Public Enum Scrollbars
        vbNoScrollBars = 0
        vbScrollBars = 1
    End Enum
    Public Enum ButtonGStyle
        vbWindows32 = 0
        vbWindowsXP = 1
    End Enum

Public Function Messagebox(Prompt As String, Button As ButtonStyle, Optional Icontype As Icon = 0, Optional title As String = "", Optional appmodal As Modal = 0, Optional useScrollbars As Scrollbars = 0, Optional DialogHeader As String = "", Optional TextHeader As String = "", Optional timed As Long = 0, Optional xpButton As ButtonGStyle = 0, Optional Top As Long = -1, Optional Left As Long = -1) As msgResult

'pic top 135
'pic left 105

    Dim TheForm As Form
    Dim X, i As Integer
    ' Disable All Forms
If appmodal = vbModal Then
    For Each TheForm In Forms
        DoEvents
        TheForm.Enabled = False
    Next TheForm
End If
    ' Load Message Box
    Load frmMSGBOX
    

    
    
    If xpButton = 1 Then
        For X = 0 To 3
            frmMSGBOX.command1(X).ButtonType = [Windows XP]
        Next X
    Else
        For X = 0 To 3
            frmMSGBOX.command1(X).ButtonType = [Windows 32-bit]
        Next X
    End If
    
    ' Initial Value
    If Button = 0 Then Button = 4
    ' Set Caption
    frmMSGBOX.lblMessage.BackColor = frmMSGBOX.BackColor
    frmMSGBOX.lblscrollbars.BackColor = frmMSGBOX.BackColor
    
    frmMSGBOX.Label1.Caption = ""
    If DialogHeader <> "" Then frmMSGBOX.Label1.Caption = DialogHeader
    
    frmMSGBOX.Label2.Caption = ""
    If TextHeader <> "" Then frmMSGBOX.Label2.Caption = TextHeader
      
      frmMSGBOX.lblscrollbars.Top = frmMSGBOX.lblMessage.Top
      frmMSGBOX.lblscrollbars.Left = frmMSGBOX.lblMessage.Left
    
    If useScrollbars = 0 Then
        frmMSGBOX.lblMessage.Visible = True
        frmMSGBOX.lblscrollbars.Visible = False
        frmMSGBOX.lblMessage.Text = ""
        frmMSGBOX.lblMessage.Text = Prompt
    Else
        frmMSGBOX.lblMessage.Visible = False
        frmMSGBOX.lblscrollbars.Visible = True
        frmMSGBOX.lblscrollbars.Text = ""
        frmMSGBOX.lblscrollbars.Text = Prompt
    End If
    
   ' frmMSGBOX.lblMessage.Text = ""
    'frmMSGBOX.lblMessage.Text = Prompt
    If title = "" Then title = App.title
    frmMSGBOX.Caption = title
    
    'set icon
    If Icontype = 0 Then
        frmMSGBOX.Picture1.Picture = frmMSGBOX.Picture2.Picture
    Else
        frmMSGBOX.Picture1.Picture = frmMSGBOX.Image1(Icontype - 1).Picture
    End If
    
    ' Choose ButtonStyle
    With frmMSGBOX
        .command1(0).Left = 2895
        Select Case Button
            Case 10 To 18
                .command1(0).Default = True
                Button = Button - 10
            Case 20 To 28
                .command1(1).Default = True
                Button = Button - 20
            Case 30 To 38
                .command1(2).Default = True
                Button = Button - 30
            Case 40 To 48
                .command1(3).Default = True
                Button = Button - 40
        End Select
        Select Case Button
            Case Is = 1
                For i = 0 To 3
                    .command1(i).Visible = True
                Next i
                .command1(2).Visible = False
                .command1(0).Caption = "Abort"
                .command1(1).Caption = "Retry"
                .command1(3).Caption = "Ignore"
            Case Is = 2
                .command1(2).Visible = True
                .command1(2).Caption = "Help"
            Case Is = 3
                For i = 0 To 1
                    .command1(i).Visible = True
                Next i
                .command1(0).Caption = "OK"
                .command1(1).Caption = "Cancel"
                .command1(i).Cancel = True
            Case Is = 4
                .command1(0).Visible = True
                .command1(0).Caption = "OK"
                .command1(0).Default = True
                .command1(0).Left = (frmMSGBOX.Width - .command1(1).Width) / 2
            Case Is = 5
                For i = 0 To 1
                    .command1(i).Visible = True
                Next i
                .command1(0).Caption = "Retry"
                .command1(1).Caption = "Cancel"
                .command1(1).Cancel = True
            Case Is = 6
                For i = 0 To 1
                    .command1(i).Visible = True
                Next i
                .command1(0).Caption = "Yes"
                .command1(1).Caption = "No"
            Case Is = 7
                For i = 0 To 3
                    .command1(i).Visible = True
                Next i
                .command1(2).Visible = False
                .command1(3).Caption = "Yes"
                .command1(0).Caption = "No"
                .command1(1).Caption = "Cancel"
                .command1(3).Cancel = True
              Case Is = 8
              
                For i = 0 To 3
                    .command1(i).Visible = False
                Next i
        End Select
    End With
    frmMSGBOX.Show
    
    If Top <> -1 Then
    frmMSGBOX.Top = Top
End If
If Left <> -1 Then
    frmMSGBOX.Left = Left
End If
    
    If timed <> 0 Then
        If timed <> -1 Then
            frmMSGBOX.etimer.Interval = timed
            frmMSGBOX.etimer.Enabled = True
        End If
    End If
    
    ' Disable Window until user responses
    Do
        DoEvents
    Loop Until frmMSGBOX.Clicked = True
    ' Initial value
    
    
    'MsgBox (frmMSGBOX.Text1.Text)
    
        Messagebox = ButtonNumber
    
 '   MsgBox (DoMessagebox)
    ' Enable All Forms
    If appmodal = vbModal Then
    For Each TheForm In Forms
        DoEvents
        TheForm.Enabled = True
    Next TheForm
    End If
    ' Initial Value
    frmMSGBOX.Clicked = False
    Unload frmMSGBOX
End Function


