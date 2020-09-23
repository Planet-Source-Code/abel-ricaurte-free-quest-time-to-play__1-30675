VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFinal 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmFinal.frx":0000
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlSph 
      Left            =   4980
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   30
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinal.frx":3A78
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinal.frx":4594
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinal.frx":50B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picEnd 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3300
      Left            =   1650
      ScaleHeight     =   220
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   260
      TabIndex        =   9
      Top             =   150
      Visible         =   0   'False
      Width           =   3900
      Begin VB.Timer tmrScroll 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   0
         Top             =   0
      End
      Begin VB.TextBox txtEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3000
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "frmFinal.frx":5BCC
         Top             =   3000
         Width           =   3600
      End
   End
   Begin VB.Timer tmrSign 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2625
      Top             =   1890
   End
   Begin VB.PictureBox picSign 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   2625
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   6
      Top             =   1890
      Width           =   450
   End
   Begin VB.PictureBox picCha 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1350
      Index           =   1
      Left            =   4200
      Picture         =   "frmFinal.frx":5BD5
      ScaleHeight     =   1320
      ScaleWidth      =   1320
      TabIndex        =   1
      Top             =   1275
      Width           =   1350
   End
   Begin VB.PictureBox picCha 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1350
      Index           =   0
      Left            =   150
      ScaleHeight     =   1320
      ScaleWidth      =   1320
      TabIndex        =   0
      Top             =   1275
      Width           =   1350
   End
   Begin VB.Image Image2 
      Height          =   150
      Left            =   1080
      Picture         =   "frmFinal.frx":6238
      Top             =   2880
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   150
      Left            =   5055
      Picture         =   "frmFinal.frx":62C9
      Top             =   2880
      Width           =   225
   End
   Begin VB.Label lblRoll 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Roll"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   2115
      TabIndex        =   12
      Top             =   3570
      Width           =   1500
   End
   Begin VB.Label lblVal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   5325
      TabIndex        =   10
      Top             =   2850
      Width           =   90
   End
   Begin VB.Label lblVal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   1350
      TabIndex        =   8
      Top             =   2850
      Width           =   90
   End
   Begin VB.Label lblStory 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Final Match!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   150
      LinkTimeout     =   100
      TabIndex        =   7
      Top             =   180
      Width           =   5400
   End
   Begin VB.Image imgHP 
      Appearance      =   0  'Flat
      Height          =   120
      Index           =   0
      Left            =   150
      Picture         =   "frmFinal.frx":635A
      Top             =   2700
      Width           =   1350
   End
   Begin VB.Image imgHP 
      Height          =   120
      Index           =   1
      Left            =   4200
      Picture         =   "frmFinal.frx":666F
      Top             =   2700
      Width           =   1350
   End
   Begin VB.Label lblCha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Opponent"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   5
      Top             =   1020
      Width           =   1350
   End
   Begin VB.Label lblCha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblChar(0)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   1020
      Width           =   1350
   End
   Begin VB.Label lblSta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblStat(1)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   3
      Top             =   2850
      Width           =   345
   End
   Begin VB.Label lblSta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblStat(0)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   2850
      Width           =   675
   End
End
Attribute VB_Name = "frmFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iSph As Integer

Private Sub Form_Activate()
  varTmp = Split(ReadINI("Opponent", "Data", strFree), "")
  If UBound(varTmp) = 1 Then lblCha(1) = varTmp(0): lblSta(1) = varTmp(1)
  imgHP(0).Width = frmMain.picHP.Width
  lblVal(0) = frmMain.lblVal(6)
  lblStory = frmMain.strList
  txtEnd.Top = picEnd.Height
  frmMain.Enabled = False
  lblVal(1) = lblVal(0)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblRoll.FontBold = False
End Sub

Private Sub lblRoll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblRoll.FontBold = True
End Sub

Private Sub tmrSign_Timer()
  If iSph = 3 Then iSph = 1 Else iSph = iSph + 1
  picSign.Picture = imlSph.ListImages(iSph).Picture
End Sub

Private Sub tmrScroll_Timer()
  txtEnd.Top = txtEnd.Top - 1
End Sub

Private Sub lblRoll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Select Case lblRoll
    Case "Roll"
      lblStory = ""
      lblRoll = "Stop"
      tmrSign.Enabled = True
      
    Case "Stop"
      tmrSign.Enabled = False
      EndingQuest
      If Me.Tag = "" Then lblRoll = "Roll" Else lblRoll = "Continue"
    
    Case "Continue"
      lblRoll = "The End"
      varTmp = ReadINI("Ending", Me.Tag, strFree)
      txtEnd = Replace(varTmp, "|", vbCrLf)
      tmrScroll.Enabled = True
      picEnd.Visible = True
      
    Case "The End"
      If Me.Tag = "Win" Then txtEnd = txtEnd & vbCrLf & "( +30 Exp. Points )": frmMain.lblExp = frmMain.lblExp + 30
      If frmHall.Rating = True Then frmHall.Show
      frmMain.txtMain = vbCrLf & txtEnd
      frmMain.Enabled = True
      Unload Me
    End Select
End Sub

Private Sub EndingQuest()
  Dim iQ As Integer
  Select Case iSph
    Case 1
      iQ = -6: lblStory = "Mutual Damage!"
      
    Case 2
      iQ = (lblVal(1) * 5): lblStory = "Opponent's " & lblSta(1)
    
    Case 3
      iQ = (lblVal(0) * -5): lblStory = lblSta(0) & "!"
    End Select
    Me.Tag = ""
    If iQ < 0 Then If -iQ < imgHP(1).Width Then imgHP(1).Width = imgHP(1).Width + iQ Else imgHP(0).Width = 90: lblStory = "You Win!": Me.Tag = "Win"
    If iQ = -6 Then iQ = 6
    If iQ > 0 Then If iQ < imgHP(0).Width Then imgHP(0).Width = imgHP(0).Width - iQ Else imgHP(0).Width = 1: lblStory = "You Lose!": Me.Tag = "Lost"
    frmMain.picHP.Width = imgHP(0).Width
End Sub
