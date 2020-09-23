VERSION 5.00
Begin VB.Form frmMSGBOX 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5070
   ControlBox      =   0   'False
   Icon            =   "frmMSGBOX.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   3
      Left            =   3840
      TabIndex        =   5
      Top             =   2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   2
      Left            =   2580
      TabIndex        =   4
      Top             =   2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Top             =   2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picMSGBOX 
      Height          =   1680
      Left            =   720
      ScaleHeight     =   1620
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   585
      Width           =   4395
      Begin VB.Label lblMessage 
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   240
         TabIndex        =   1
         Top             =   180
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmMSGBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Clicked As Boolean
Public ButtonNumber As Integer

Private Sub Command1_Click(Index As Integer)
    ' Initial Value
    Clicked = True
    ' Initial value
    ButtonNumber = Index
    ' Unload Form
    Unload Me
End Sub

Private Sub Form_Activate()
    ' Beep
    Beep
End Sub
