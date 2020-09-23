VERSION 5.00
Begin VB.Form frm_designer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message Box Designer"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   Icon            =   "Frm_designer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chk_inc_res 
      Caption         =   "Include Optional If-End If for result"
      Height          =   360
      Left            =   5535
      TabIndex        =   34
      Top             =   5970
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   8865
      TabIndex        =   33
      Top             =   6990
      Width           =   900
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   7830
      TabIndex        =   31
      Top             =   5340
      Width           =   2865
   End
   Begin VB.CommandButton cmd_load 
      Caption         =   "Load"
      Height          =   375
      Left            =   7815
      TabIndex        =   30
      Top             =   6990
      Width           =   990
   End
   Begin VB.CheckBox chk_clipboard 
      Caption         =   "Copy to clipboard"
      Height          =   315
      Left            =   5520
      TabIndex        =   29
      Top             =   5670
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "Save As"
      Height          =   375
      Left            =   9765
      TabIndex        =   28
      Top             =   6990
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Readme"
      Height          =   360
      Left            =   1080
      TabIndex        =   27
      Top             =   6915
      Width           =   1635
   End
   Begin VB.ComboBox cmb_Icon 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Frm_designer.frx":5D52
      Left            =   1305
      List            =   "Frm_designer.frx":5D7D
      TabIndex        =   6
      Text            =   "vbWin32"
      Top             =   1155
      Width           =   1515
   End
   Begin VB.TextBox txt_DialogHeader 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   3465
      TabIndex        =   5
      Text            =   "[Dialog Header]"
      Top             =   1110
      Width           =   3885
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   930
      Left            =   1200
      TabIndex        =   25
      Top             =   870
      Width           =   6270
   End
   Begin VB.CheckBox chk_test 
      Caption         =   "Test Message"
      Height          =   315
      Left            =   5520
      TabIndex        =   23
      Top             =   5385
      Value           =   1  'Checked
      Width           =   1410
   End
   Begin VB.TextBox txt_left 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   450
      TabIndex        =   21
      Text            =   "-1"
      Top             =   540
      Width           =   690
   End
   Begin VB.TextBox txt_Top 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1575
      TabIndex        =   19
      Text            =   "-1"
      Top             =   195
      Width           =   690
   End
   Begin VB.CheckBox chk_XP_Style 
      Caption         =   "XP Style Buttons"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7785
      TabIndex        =   17
      Top             =   3435
      Width           =   1725
   End
   Begin VB.ComboBox cmb_Modal 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Frm_designer.frx":5E15
      Left            =   7725
      List            =   "Frm_designer.frx":5E1F
      TabIndex        =   15
      Text            =   "vbNoModal"
      Top             =   1815
      Width           =   2460
   End
   Begin VB.TextBox txt_Timer 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9585
      TabIndex        =   13
      Text            =   "-1"
      Top             =   2760
      Width           =   690
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   -405
      TabIndex        =   11
      Top             =   4950
      Width           =   12840
   End
   Begin VB.TextBox txt_source 
      Height          =   1380
      Left            =   1110
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Text            =   "Frm_designer.frx":5E37
      Top             =   5370
      Width           =   4335
   End
   Begin VB.CheckBox chk_Scrollbar 
      Caption         =   "Add Scrollbar (does not affect this scrollbar)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7755
      TabIndex        =   8
      Top             =   2190
      Width           =   3195
   End
   Begin VB.TextBox txt_Dialog_Caption 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Text            =   "[Dialog Caption]"
      Top             =   540
      Width           =   6285
   End
   Begin VB.TextBox txt_text_Header 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1305
      TabIndex        =   4
      Text            =   "[Text Header]"
      Top             =   1980
      Width           =   6030
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   1185
      TabIndex        =   3
      Top             =   1785
      Width           =   6285
   End
   Begin VB.TextBox txt_Caption 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   1305
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Frm_designer.frx":5E56
      Top             =   2355
      Width           =   5955
   End
   Begin VB.ComboBox cmd_ButtonStyle 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Frm_designer.frx":5E6C
      Left            =   3300
      List            =   "Frm_designer.frx":5E85
      TabIndex        =   1
      Text            =   "VbOKOnly"
      Top             =   3735
      Width           =   2460
   End
   Begin VB.CommandButton cmd_test 
      Caption         =   "Get Source"
      Height          =   375
      Left            =   5535
      TabIndex        =   0
      Top             =   6390
      Width           =   1155
   End
   Begin VB.CommandButton cmd_frm 
      Enabled         =   0   'False
      Height          =   4125
      Left            =   1170
      TabIndex        =   24
      Top             =   510
      Width           =   6345
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Saved Message Boxes"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   7785
      TabIndex        =   32
      Top             =   5100
      Width           =   2430
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Options"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   210
      Left            =   8430
      TabIndex        =   26
      Top             =   1125
      Width           =   1425
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "(-1 = Centred in Screen)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   240
      Left            =   2325
      TabIndex        =   22
      Top             =   195
      Width           =   2325
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   240
      Left            =   75
      TabIndex        =   20
      Top             =   585
      Width           =   675
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Top"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   240
      Left            =   1200
      TabIndex        =   18
      Top             =   210
      Width           =   675
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Modal"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   210
      Left            =   7725
      TabIndex        =   16
      Top             =   1545
      Width           =   1425
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "in m/s.  min=0  max = 60000 (1 minute)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   7830
      TabIndex        =   14
      Top             =   3060
      Width           =   3240
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Timeout (-1 no timeout)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   240
      Left            =   7785
      TabIndex        =   12
      Top             =   2805
      Width           =   2220
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Source Code"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   210
      Left            =   1140
      TabIndex        =   10
      Top             =   5130
      Width           =   1425
   End
End
Attribute VB_Name = "frm_designer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_load_Click()
    Dim nme As String
    Dim x As Long
    
    
    For x = 0 To File1.ListCount - 1
        If File1.Selected(x) = True Then nme = File1.List(x): Exit For
    Next x
    load_it nme
End Sub

Private Sub cmd_save_Click()
Dim saveas As String
saveas = InputBox("Enter a name. (No extension required)", "Save Msgbox file")

If saveas = "" Then MsgBox ("Cancelled"): Exit Sub

SaveFile saveas

End Sub
Private Sub SaveFile(nme As String)

Dim F As Integer
Dim app_path As String
F = FreeFile
app_path = App.Path
If Right$(app_path, 1) <> "\" Then app_path = app_path + "\"

If Right$(UCase$(nme), 4) = ".MSF" Then nme = Mid(nme, 1, Len(nme) - 4)

Open app_path + nme + ".msf" For Output As #F

    Print #F, txt_Top
    Print #F, txt_left
    Print #F, txt_Dialog_Caption
    Print #F, cmb_Icon.Text
    Print #F, txt_DialogHeader
    Print #F, txt_text_Header
    Print #F, rClean(txt_Caption)
    Print #F, cmd_ButtonStyle.Text
    Print #F, cmb_Modal.Text
    Print #F, chk_Scrollbar.Value
    Print #F, txt_Timer
    Print #F, chk_XP_Style.Value
    Print #F, chk_test.Value
    Print #F, chk_clipboard.Value
    Print #F, chk_inc_res.Value
Close F

If isfile(app_path + "lastused.msf") = True Then
    Kill app_path + "lastused.msf"
    FileCopy app_path + nme + ".msf", app_path + "LastUsed.msf1"
End If


frm_designer.Caption = "Message Box Designer - (" + nme + ".msf)"


File1.Refresh
End Sub
Private Sub load_it(nme As String)
On Error GoTo err_cap
Dim F As Integer
Dim app_path, tmp As String
F = FreeFile
app_path = App.Path
If Right$(app_path, 1) <> "\" Then app_path = app_path + "\"
If isfile(app_path + nme) = False Then Exit Sub
Open app_path + nme For Input As #F

    Line Input #F, tmp
     txt_Top = tmp
    
    Line Input #F, tmp
    txt_left = tmp
    
    Line Input #F, tmp
     txt_Dialog_Caption = tmp
    
    Line Input #F, tmp
    cmb_Icon.Text = tmp
    
    Line Input #F, tmp
     txt_DialogHeader = tmp
    
    Line Input #F, tmp
    txt_text_Header = tmp
    
    Line Input #F, tmp
    txt_Caption = rUnClean(tmp)
    
    Line Input #F, tmp
    cmd_ButtonStyle.Text = tmp
    
    Line Input #F, tmp
    cmb_Modal.Text = tmp
    
    Line Input #F, tmp
    chk_Scrollbar.Value = Val(tmp)
    
    Line Input #F, tmp
    txt_Timer = tmp
    
    Line Input #F, tmp
    chk_XP_Style.Value = Val(tmp)
    
    Line Input #F, tmp
    chk_test.Value = Val(tmp)
    
    Line Input #F, tmp
    chk_clipboard.Value = Val(tmp)
    
    Line Input #F, tmp
    chk_inc_res.Value = Val(tmp)
    
    
Close F

If UCase$(nme) <> "LASTUSED.MSF1" Then
    frm_designer.Caption = "Message Box Designer - (" + nme + ")"
End If

Exit Sub
err_cap:
MsgBox ("An error occured loading. It may be that the msf file does not inlcude an expected value because the designer is a newer version" + vbCrLf + "You should be able to continue, but you will need to re-save all the files in order to update them")
Resume Next
End Sub

Private Sub cmd_test_Click()
    
    Dim ANS, Q, C, SC As String
    
    Dim Tim, L, T As Long
    
    Dim Scr As Scrollbars
    Dim XP As ButtonGStyle
    Dim Bt As ButtonStyle
    Dim IC As Icon
    Dim Moda As Modal
    
    
    
    Q = Chr$(34)
    C = Chr$(44)
    
    SC = "Dim Res" + vbCrLf + "Res = Messagebox("
    
    SC = SC + Q + Clean(txt_Caption.Text) + Q + C
    
    SC = SC + cmd_ButtonStyle.Text + C
    
            Select Case cmd_ButtonStyle.Text
                Case "vbOkOnly"
                    Bt = vbOKOnly
                    ANS = "'If Res = vbOKOnly then " + vbCrLf + "'" + vbCrLf + "'End if"
                Case "vbRetryCancel"
                    Bt = vbRetryCancel
                    ANS = "'If Res = vbRetry then " + vbCrLf + "'" + vbCrLf + "'End if" + vbCrLf
                    ANS = ANS + "'If Res = vbCancel then " + vbCrLf + "'" + vbCrLf + "'End if" + vbCrLf
                Case "vbYesNo"
                    Bt = vbYesNo
                    ANS = "'If Res = vbYes then " + vbCrLf + "'" + vbCrLf + "'End if" + vbCrLf
                    ANS = ANS + "'If Res = vbNo then " + vbCrLf + "'" + vbCrLf + "'End if" + vbCrLf
                Case "vbYesNoCancel"
                    Bt = vbYesNoCancel
                    ANS = "'If Res = vbYes then " + vbCrLf + "'" + vbCrLf + "'End if" + vbCrLf
                    ANS = ANS + "'If Res = vbNo then " + vbCrLf + "'" + vbCrLf + "'End if" + vbCrLf
                    ANS = ANS + "'If Res = vbCancel then " + vbCrLf + "'" + vbCrLf + "'End if" + vbCrLf
                Case "vbOKCancel"
                    Bt = vbOKCancel
                    ANS = "'If Res = vbOK then " + vbCrLf + "'" + vbCrLf + "'End if" + vbCrLf
                    ANS = ANS + "'If Res = vbCancel then " + vbCrLf + "'" + vbCrLf + "'End if" + vbCrLf
                Case "vbAbortRetryIgnore"
                    Bt = vbAbortRetryIgnore
                    ANS = "'If Res = vbAbort then " + vbCrLf + "'" + vbCrLf + "'End if" + vbCrLf
                    ANS = ANS + "'If Res = vbRetry then " + vbCrLf + "'" + vbCrLf + "'End if" + vbCrLf
                    ANS = ANS + "'If Res = vbIgnore then " + vbCrLf + "'" + vbCrLf + "'End if" + vbCrLf
                Case "VBNoButtons"
                    Bt = vbNoButtons
            End Select
    
    
    SC = SC + cmb_Icon.Text + C
            
        Select Case cmb_Icon.Text
                
            Case "vbWin32"
                IC = vbWin32
            Case "vbInformation"
                IC = vbInformation
            Case "vbExclamation"
                IC = vbExclamation
            Case "vbCritical"
                IC = vbCritical
            Case "vbQuestion"
                IC = vbQuestion
            Case "vbComputer"
                IC = vbComputer
            Case "vbHardware"
                IC = vbHardware
            Case "vbInternet"
                IC = vbInternet
            Case "vbPower"
                IC = vbPower
            Case "vbSearch"
                IC = vbSearch
            Case "vbSecurity"
                IC = vbSecurity
            Case "vbSound"
                IC = vbSound
            Case "vbUsers"
                IC = vbUsers
        End Select
    
    SC = SC + Q + txt_Dialog_Caption + Q + C
    
    SC = SC + cmb_Modal.Text + C
    
        Select Case cmb_Modal.Text
                
            Case "vbModal"
                Moda = vbModal
            Case "vbNoModal"
                Moda = vbNoModal
        End Select
    
    
    If chk_Scrollbar.Value = vbChecked Then
        SC = SC + "vbScrollBars" + C
        Scr = vbScrollBars
    Else
        SC = SC + "vbNoScrollBars" + C
        Scr = vbNoScrollBars
    End If
    
    SC = SC + Q + txt_DialogHeader.Text + Q + C
    SC = SC + Q + txt_text_Header.Text + Q + C
    SC = SC + txt_Timer + C
    
    If chk_XP_Style.Value = vbChecked Then
        SC = SC + "vbWindowsXP" + C
        XP = vbWindowsXP
    Else
        SC = SC + "vbWindows32" + C
        XP = vbWindows32
    End If
    
    SC = SC + txt_left.Text + C
    SC = SC + txt_Top.Text + ")" + vbCrLf
    
    txt_source = SC
    
    If chk_inc_res.Value = vbChecked Then
        txt_source = txt_source + vbCrLf + ANS
    End If
    
    If chk_test.Value = vbChecked Then
        Dim Res
        Res = Messagebox(txt_Caption.Text, Bt, IC, txt_Dialog_Caption, Moda, Scr, txt_DialogHeader, txt_text_Header, Val(txt_Timer), XP, Val(txt_left.Text), Val(txt_Top.Text))
        
    
    
    If Res = vbYes Then MsgBox ("You pressed Yes!")
    If Res = vbtimedout Then MsgBox ("Timed Out")
    If Res = vbAbort Then MsgBox ("Abort pressed")
    If Res = vbRetry Then MsgBox ("Retry pressed")
    If Res = vbIgnore Then MsgBox ("Ignore pressed")
    If Res = vbNo Then MsgBox ("You pressed No")
    If Res = vbCancel Then MsgBox ("You pressed Cancel")
     If Res = vbOK Then MsgBox ("You pressed OK")
    
    End If
    
    If chk_clipboard.Value = vbChecked Then
    
        
        Clipboard.SetText txt_source.Text
    
    End If
    

End Sub

Public Function Clean(txt_Clean As String) As String

   ' this function simply replaces all bad file-save characters
   ' like quote, comma and return and turns them into good characters like
   ' {rtn} = return, {cma} = comma, {qte} = quote, {crg} = carrage return
        Dim CH As String
        'CH = "{VBCRLF}"
        CH = "Chr$(34)+vbcrlf+Chr$(34)"
        '"+vbcrlf+"
       'txt_Clean = Replace$(txt_Clean, Chr$(34), "+chr$(34)+")
       txt_Clean = Replace$(txt_Clean, Chr$(13), CH)
       txt_Clean = Replace$(txt_Clean, "Chr$(34)", Chr$(34))
       txt_Clean = Replace$(txt_Clean, Chr$(10), "")
       Clean = txt_Clean

End Function

Private Sub Command1_Click()
  
  Dim nme As String
    Dim x As Long
    Dim chec As Boolean
    chec = False
    For x = 0 To File1.ListCount - 1
        If File1.Selected(x) = True Then nme = File1.List(x): chec = True: Exit For
    Next x
    If chec = False Then Exit Sub: ' coz nothing is selected
    SaveFile nme
End Sub

Private Sub Command2_Click()
    Shell "notepad.exe readme.txt", vbNormalFocus
End Sub




Private Sub File1_DblClick()
    cmd_load_Click
End Sub

Private Sub Form_Load()
    load_it "Lastused.msf1"
    File1.Path = App.Path
    File1.Pattern = "*.msf"
End Sub
Private Function rClean(txt_Clean As String) As String

   ' this function simply replaces all bad file-save characters
   ' like quote, comma and return and turns them into good characters like
   ' {rtn} = return, {cma} = comma, {qte} = quote, {crg} = carrage return

       txt_Clean = Replace$(txt_Clean, Chr$(34), "{qte}")
       txt_Clean = Replace$(txt_Clean, Chr$(13), "{rtn}")
       txt_Clean = Replace$(txt_Clean, Chr$(10), "{crg}")
       txt_Clean = Replace$(txt_Clean, ",", "{cma}")
    
       rClean = txt_Clean

End Function
 Private Function rUnClean(txt_UnClean As String) As String

   ' this function simply undo's the Clean_String function so it can be viewed
   ' how it should.
    
       txt_UnClean = Replace$(txt_UnClean, "{qte}", Chr$(34))
       txt_UnClean = Replace$(txt_UnClean, "{rtn}", Chr$(13))
       txt_UnClean = Replace$(txt_UnClean, "{crg}", Chr$(10))
       txt_UnClean = Replace$(txt_UnClean, "{cma}", ",")
    
       rUnClean = txt_UnClean

End Function
Private Function isfile(filename As String) As Boolean

   'Funct to return boolean if a file exists
   '-------------------
   
   Dim F As Byte

       On Error GoTo err
       F = FreeFile
       Open filename For Input As #F
       Close F
       isfile = True

Exit Function

err:
       isfile = False
       Close F

End Function
