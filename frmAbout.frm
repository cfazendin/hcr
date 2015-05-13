VERSION 4.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About HCR32..."
   ClientHeight    =   2145
   ClientLeft      =   2970
   ClientTop       =   2715
   ClientWidth     =   3735
   ForeColor       =   &H00000000&
   Height          =   2550
   Icon            =   "FRMABOUT.frx":0000
   Left            =   2910
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   3735
   Top             =   2370
   Width           =   3855
   Begin VB.CommandButton cmdCloseAbout 
      BackColor       =   &H00FFFF80&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox picIcon 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      Picture         =   "FRMABOUT.frx":0442
      ScaleHeight     =   615
      ScaleWidth      =   555
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   555
   End
   Begin VB.Label lblHttpAddress 
      AutoSize        =   -1  'True
      Caption         =   "http://www.winternet.com/~faz/HCR/"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   840
      TabIndex        =   8
      Top             =   1920
      Width           =   2745
   End
   Begin VB.Label lblEmailAddress 
      AutoSize        =   -1  'True
      Caption         =   "HCR@Kingpin.Com"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   300
      Left            =   960
      TabIndex        =   7
      Top             =   1560
      Width           =   2340
   End
   Begin VB.Label lblText5 
      AutoSize        =   -1  'True
      Caption         =   "E-mail all comments, bugs, $$ to"
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Top             =   1320
      Width           =   2265
   End
   Begin VB.Label lblText4 
      AutoSize        =   -1  'True
      Caption         =   "This is a freely distributable program."
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   1080
      Width           =   2895
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      X1              =   840
      X2              =   4560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblVersionInformation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HCR 2.04 for Windows95"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 1995 Christopher Fazendin"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label lblMainHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HTML Color Reference"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   840
      TabIndex        =   2
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   2760
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload frmAbout
End Sub


Private Sub cmdCloseAbout_Click()
    Unload frmAbout
End Sub


