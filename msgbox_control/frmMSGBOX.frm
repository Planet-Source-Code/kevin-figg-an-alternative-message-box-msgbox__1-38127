VERSION 5.00
Begin VB.Form frmMSGBOX 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5580
   ControlBox      =   0   'False
   Icon            =   "frmMSGBOX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   165
      TabIndex        =   7
      Top             =   1050
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer etimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   90
      Top             =   2850
   End
   Begin VB.TextBox lblscrollbars 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   4995
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmMSGBOX.frx":000C
      Top             =   1620
      Visible         =   0   'False
      Width           =   5355
   End
   Begin VB.TextBox lblMessage 
      BorderStyle     =   0  'None
      Height          =   1140
      HideSelection   =   0   'False
      Left            =   90
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Text            =   "frmMSGBOX.frx":0012
      Top             =   1515
      Width           =   5355
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   7860
      Picture         =   "frmMSGBOX.frx":0018
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   855
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   225
      Picture         =   "frmMSGBOX.frx":1B3A
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   180
      Width           =   795
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -330
      TabIndex        =   0
      Top             =   885
      Width           =   6225
   End
   Begin CMSG.chameleonButton command1 
      Height          =   360
      Index           =   0
      Left            =   3165
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   635
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMSGBOX.frx":365C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CMSG.chameleonButton command1 
      Height          =   360
      Index           =   1
      Left            =   4350
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   635
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMSGBOX.frx":3678
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CMSG.chameleonButton command1 
      Height          =   360
      Index           =   2
      Left            =   540
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   635
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMSGBOX.frx":3694
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CMSG.chameleonButton command1 
      Height          =   360
      Index           =   3
      Left            =   1770
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   635
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMSGBOX.frx":36B0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   180
      TabIndex        =   5
      Top             =   1020
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   3780
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   11
      Left            =   6870
      Picture         =   "frmMSGBOX.frx":36CC
      Top             =   900
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   10
      Left            =   6345
      Picture         =   "frmMSGBOX.frx":450E
      Top             =   900
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   9
      Left            =   11010
      Picture         =   "frmMSGBOX.frx":5350
      Top             =   270
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   8
      Left            =   10485
      Picture         =   "frmMSGBOX.frx":6192
      Top             =   270
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   9960
      Picture         =   "frmMSGBOX.frx":6FD4
      Top             =   255
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   9435
      Picture         =   "frmMSGBOX.frx":7E16
      Top             =   255
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   8910
      Picture         =   "frmMSGBOX.frx":8C58
      Top             =   255
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   8385
      Picture         =   "frmMSGBOX.frx":9A9A
      Top             =   255
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   7845
      Picture         =   "frmMSGBOX.frx":A8DC
      Top             =   255
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   7320
      Picture         =   "frmMSGBOX.frx":B71E
      Top             =   255
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   6795
      Picture         =   "frmMSGBOX.frx":C560
      Top             =   255
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   6270
      Picture         =   "frmMSGBOX.frx":D3A2
      Top             =   255
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   -75
      Top             =   -90
      Width           =   5655
   End
End
Attribute VB_Name = "frmMSGBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Public Clicked As Boolean


Private Sub Command1_Click(Index As Integer)
    ' Initial Value
    Clicked = True
    ' Initial value
    
        ' mCancel = 0
       ' mOk = 1
       
       ' mRetry = 3
       ' mYes = 4
       ' mNo = 5
       ' mHelp = 6
       ' mAbort = 7
        'mIgnore = 8
    
    If command1(Index).Caption = "Cancel" Then ButtonNumber = 0
    If command1(Index).Caption = "OK" Then ButtonNumber = 1
    If command1(Index).Caption = "Retry" Then ButtonNumber = 2
    If command1(Index).Caption = "Yes" Then ButtonNumber = 3
    If command1(Index).Caption = "No" Then ButtonNumber = 4
    If command1(Index).Caption = "Help" Then ButtonNumber = 5
    If command1(Index).Caption = "Abort" Then ButtonNumber = 6
    If command1(Index).Caption = "Ignore" Then ButtonNumber = 7
    'MsgBox (buttonnumber)
   
    ' Unload Form
    
End Sub

Private Sub etimer_Timer()
    etimer.Enabled = False
    Clicked = True
    ButtonNumber = -1
End Sub

Private Sub Form_Activate()
    ' Beep
    Beep
End Sub

