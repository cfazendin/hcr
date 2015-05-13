VERSION 4.00
Begin VB.Form frmHCR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTML Color Reference v2.04"
   ClientHeight    =   3765
   ClientLeft      =   1035
   ClientTop       =   1665
   ClientWidth     =   8130
   ClipControls    =   0   'False
   Height          =   4455
   Icon            =   "FRMHCR.frx":0000
   Left            =   975
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   8130
   Top             =   1035
   Width           =   8250
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   345
      Left            =   5760
      TabIndex        =   14
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox AlinkTag 
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "alink=""#551A8B"""
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox VlinkTag 
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "vlink=""#551A8B"""
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox linkTag 
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "link=""#5500EE"""
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtTag 
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "text=""#000000"""
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox bgTag 
      Height          =   300
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "bgcolor=""#C0C0C0"""
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtMain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   $"FRMHCR.frx":0442
      Top             =   3360
      Width           =   7935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Frame fraColControl 
      Caption         =   "Color Control"
      ClipControls    =   0   'False
      Height          =   2070
      Left            =   4200
      TabIndex        =   21
      Top             =   1200
      Width           =   1455
      Begin VB.OptionButton optALink 
         Caption         =   "Active Link"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton optVLink 
         Caption         =   "Visited Link"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton optLink 
         Caption         =   "Link Text"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optNormal 
         Caption         =   "Normal Text"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optBackground 
         Caption         =   "Background"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      DataField       =   "BLAH"
      FontTransparent =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   3915
      TabIndex        =   16
      Top             =   1200
      Width           =   3975
      Begin VB.Label lblALink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This is an ACTIVE LINK"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008B1A55&
         Height          =   240
         Left            =   840
         TabIndex        =   20
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   2430
      End
      Begin VB.Label lblVLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This is a VISITED LINK"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008B1A55&
         Height          =   240
         Left            =   840
         TabIndex        =   19
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   2385
      End
      Begin VB.Label lblLink 
         AutoSize        =   -1  'True
         BackColor       =   &H008B1A55&
         BackStyle       =   0  'Transparent
         Caption         =   "This is a LINK"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EE0000&
         Height          =   240
         Left            =   1320
         TabIndex        =   18
         Top             =   660
         UseMnemonic     =   0   'False
         Width           =   1440
      End
      Begin VB.Label lblNormal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This is NORMAL text"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   960
         TabIndex        =   17
         Top             =   330
         UseMnemonic     =   0   'False
         Width           =   2130
      End
   End
   Begin VB.TextBox txtGreen 
      Height          =   285
      Left            =   5160
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "192"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtBlue 
      Height          =   285
      Left            =   5160
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "192"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtRed 
      Height          =   285
      HideSelection   =   0   'False
      Left            =   5160
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "192"
      Top             =   120
      Width           =   495
   End
   Begin VB.HScrollBar hsbRed 
      Height          =   255
      LargeChange     =   10
      Left            =   840
      Max             =   255
      TabIndex        =   1
      Top             =   120
      Value           =   192
      Width           =   4215
   End
   Begin VB.HScrollBar hsbGreen 
      Height          =   255
      LargeChange     =   10
      Left            =   840
      Max             =   255
      TabIndex        =   3
      Top             =   480
      Value           =   192
      Width           =   4215
   End
   Begin VB.HScrollBar hsbBlue 
      Height          =   270
      LargeChange     =   10
      Left            =   840
      Max             =   255
      TabIndex        =   5
      Top             =   825
      Value           =   192
      Width           =   4215
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00EE0055&
      BorderWidth     =   2
      X1              =   0
      X2              =   8880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblBlue 
      AutoSize        =   -1  'True
      Caption         =   "&Blue"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   390
   End
   Begin VB.Label lblGreen 
      AutoSize        =   -1  'True
      Caption         =   "&Green"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   525
   End
   Begin VB.Label lblRed 
      AutoSize        =   -1  'True
      Caption         =   "&Red"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmHCR"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private SelOption As Integer
Private BackGrnd(1 To 3)
Private NormalTxt(1 To 3)
Private LinkTxt(1 To 3)
Private VLinkTxt(1 To 3)
Private ALinkTxt(1 To 3)
Private xCounter As Integer
Private bgCol(1 To 3)
Private txtCol(1 To 3)
Private linkCol(1 To 3)
Private vlinkCol(1 To 3)
Private alinkCol(1 To 3)
Private x As Integer

Private Sub cmdAbout_Click()
    frmAbout.Show 1
End Sub



Private Sub cmdExit_Click()
    Unload frmHCR
End Sub


Private Sub hsbBlue_Change()
    hsbBlue_Scroll
End Sub

Private Sub hsbBlue_Scroll()
    FindColor
End Sub


Private Sub hsbGreen_Change()
    hsbGreen_Scroll
End Sub


Private Sub hsbGreen_Scroll()
    FindColor
End Sub


Private Sub hsbRed_Change()
    hsbRed_Scroll
End Sub
Private Sub hsbRed_Scroll()
    FindColor
End Sub


Private Sub lblALink_Click()
    optALink.Value = True
    FindOption
End Sub

Private Sub lblLink_Click()
    optLink.Value = True
    FindOption
End Sub

Private Sub lblNormal_Click()
    optNormal.Value = True
    FindOption
End Sub

Private Sub lblVLink_Click()
    optVLink.Value = True
    FindOption
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText Screen.ActiveControl.SelText
End Sub

Private Sub mnuExit_Click()
    Unload frmHCR
End Sub



Private Sub FindColor()
    If (xCounter = 0) Then
        SaveColor
        xCounter = xCounter + 1
    End If
    FindOption
    Select Case SelOption
    Case 1
        picColor.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
    Case 2
        lblNormal.ForeColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
    Case 3
        lblLink.ForeColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
    Case 4
        lblVLink.ForeColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
    Case 5
        lblALink.ForeColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
    End Select
    ShowHex
End Sub

Private Static Sub FindOption()
    If (optBackground.Value = True) Then
        If SelOption <> 1 Then
            hsbRed.Value = BackGrnd(1)
            hsbGreen.Value = BackGrnd(2)
            hsbBlue.Value = BackGrnd(3)
        End If
        txtRed.Text = hsbRed.Value
        txtGreen.Text = hsbGreen.Value
        txtBlue.Text = hsbBlue.Value
        BackGrnd(1) = hsbRed.Value
        BackGrnd(2) = hsbGreen.Value
        BackGrnd(3) = hsbBlue.Value
        SelOption = 1
    ElseIf (optNormal.Value = True) Then
        If SelOption <> 2 Then
            hsbRed.Value = NormalTxt(1)
            hsbGreen.Value = NormalTxt(2)
            hsbBlue.Value = NormalTxt(3)
        End If
        txtRed.Text = hsbRed.Value
        txtGreen.Text = hsbGreen.Value
        txtBlue.Text = hsbBlue.Value
        NormalTxt(1) = hsbRed.Value
        NormalTxt(2) = hsbGreen.Value
        NormalTxt(3) = hsbBlue.Value
        SelOption = 2
    ElseIf (optLink.Value = True) Then
        If SelOption <> 3 Then
            hsbRed.Value = LinkTxt(1)
            hsbGreen.Value = LinkTxt(2)
            hsbBlue.Value = LinkTxt(3)
        End If
        txtRed.Text = hsbRed.Value
        txtGreen.Text = hsbGreen.Value
        txtBlue.Text = hsbBlue.Value
        LinkTxt(1) = hsbRed.Value
        LinkTxt(2) = hsbGreen.Value
        LinkTxt(3) = hsbBlue.Value
        SelOption = 3
    ElseIf (optVLink.Value = True) Then
        If SelOption <> 4 Then
            hsbRed.Value = VLinkTxt(1)
            hsbGreen.Value = VLinkTxt(2)
            hsbBlue.Value = VLinkTxt(3)
        End If
        txtRed.Text = hsbRed.Value
        txtGreen.Text = hsbGreen.Value
        txtBlue.Text = hsbBlue.Value
        VLinkTxt(1) = hsbRed.Value
        VLinkTxt(2) = hsbGreen.Value
        VLinkTxt(3) = hsbBlue.Value
        SelOption = 4
    Else
        If SelOption <> 5 Then
            hsbRed.Value = ALinkTxt(1)
            hsbGreen.Value = ALinkTxt(2)
            hsbBlue.Value = ALinkTxt(3)
        End If
        txtRed.Text = hsbRed.Value
        txtGreen.Text = hsbGreen.Value
        txtBlue.Text = hsbBlue.Value
        ALinkTxt(1) = hsbRed.Value
        ALinkTxt(2) = hsbGreen.Value
        ALinkTxt(3) = hsbBlue.Value
        SelOption = 5
    End If
End Sub

Private Sub SaveColor()
    SelOption = 1
    BackGrnd(1) = 192
    BackGrnd(2) = 192
    BackGrnd(3) = 192
    NormalTxt(1) = 0
    NormalTxt(2) = 0
    NormalTxt(3) = 0
    LinkTxt(1) = 85
    LinkTxt(2) = 0
    LinkTxt(3) = 238
    VLinkTxt(1) = 85
    VLinkTxt(2) = 26
    VLinkTxt(3) = 139
    ALinkTxt(1) = 85
    ALinkTxt(2) = 26
    ALinkTxt(3) = 139
End Sub

Private Sub optALink_Click()
    FindOption
End Sub

Private Sub optBackground_Click()
    FindOption
End Sub

Private Sub optLink_Click()
    FindOption
End Sub

Private Sub optNormal_Click()
    FindOption
End Sub


Private Sub optVLink_Click()
    FindOption
End Sub

Private Sub picColor_Click()
    optBackground.Value = True
    FindOption
End Sub

Private Sub txtBlue_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        hsbBlue.Value = txtBlue.Text
    End If
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
        Beep
    End If
End Sub


Private Sub txtGreen_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        hsbGreen.Value = txtGreen.Text
    End If
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtRed_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        hsbRed.Value = txtRed.Text
    End If
    If (KeyAscii < Asc("0")) Or (KeyAscii > Asc("9")) Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub ShowHex()
    For x = 1 To 3 Step 1
        y = 1
        If BackGrnd(x) > 15 Then
            y = 0
        End If
        bgCol(x) = Hex(BackGrnd(x))
        If y = 1 Then
           bgCol(x) = "0" & bgCol(x)
        End If
    Next x
bgTag.Text = "bgcolor=""#" & bgCol(1) & bgCol(2) & bgCol(3) & """"
    For x = 1 To 3 Step 1
        y = 1
        If NormalTxt(x) > 15 Then
            y = 0
        End If
        txtCol(x) = Hex(NormalTxt(x))
        If y = 1 Then
           txtCol(x) = "0" & txtCol(x)
        End If
    Next x
txtTag.Text = "text=""#" & txtCol(1) & txtCol(2) & txtCol(3) & """"
    For x = 1 To 3 Step 1
        y = 1
        If LinkTxt(x) > 15 Then
            y = 0
        End If
        linkCol(x) = Hex(LinkTxt(x))
        If y = 1 Then
           linkCol(x) = "0" & linkCol(x)
        End If
    Next x
linkTag.Text = "link=""#" & linkCol(1) & linkCol(2) & linkCol(3) & """"
    For x = 1 To 3 Step 1
        y = 1
        If VLinkTxt(x) > 15 Then
            y = 0
        End If
        vlinkCol(x) = Hex(VLinkTxt(x))
        If y = 1 Then
           vlinkCol(x) = "0" & vlinkCol(x)
        End If
    Next x
VlinkTag.Text = "vlink=""#" & vlinkCol(1) & vlinkCol(2) & vlinkCol(3) & """"
    For x = 1 To 3 Step 1
        y = 1
        If ALinkTxt(x) > 15 Then
            y = 0
        End If
        alinkCol(x) = Hex(ALinkTxt(x))
        If y = 1 Then
           alinkCol(x) = "0" & alinkCol(x)
        End If
    Next x
AlinkTag.Text = "alink=""#" & alinkCol(1) & alinkCol(2) & alinkCol(3) & """"
    txtMain.Text = "<body " & bgTag.Text & " " & txtTag.Text & " " & linkTag.Text & " " & VlinkTag.Text & " " & AlinkTag.Text & ">"
End Sub
