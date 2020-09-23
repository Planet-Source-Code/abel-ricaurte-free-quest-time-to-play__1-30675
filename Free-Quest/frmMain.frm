VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: Free-Quest ::"
   ClientHeight    =   6735
   ClientLeft      =   150
   ClientTop       =   555
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAni 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   7260
      Top             =   6300
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000007&
      Height          =   1500
      Left            =   5415
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   36
      Top             =   5220
      Width           =   1815
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00A0A0A0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   8
         Left            =   1200
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   45
         Top             =   30
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00A0A0A0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   7
         Left            =   615
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   44
         Top             =   30
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00A0A0A0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   6
         Left            =   30
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   43
         Top             =   30
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00A0A0A0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   5
         Left            =   1200
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   42
         Top             =   510
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00A0A0A0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   4
         Left            =   615
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   41
         Top             =   510
         Visible         =   0   'False
         Width           =   585
         Begin VB.Image imgGuy 
            Height          =   390
            Left            =   135
            Picture         =   "frmMain.frx":1F5C
            Top             =   45
            Width           =   315
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00A0A0A0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   30
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   40
         Top             =   510
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00A0A0A0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   1200
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   39
         Top             =   990
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00A0A0A0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   615
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   38
         Top             =   990
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00A0A0A0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   30
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   37
         Top             =   990
         Visible         =   0   'False
         Width           =   585
      End
   End
   Begin VB.PictureBox picHP 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   7305
      MouseIcon       =   "frmMain.frx":207A
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":2A44
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   30
      ToolTipText     =   "Restore"
      Top             =   2070
      Width           =   1350
   End
   Begin VB.Frame fraMain 
      Height          =   4995
      Left            =   7260
      TabIndex        =   9
      Top             =   -75
      Width           =   1440
      Begin VB.PictureBox picCha 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1350
         Left            =   45
         ScaleHeight     =   88
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   88
         TabIndex        =   4
         Top             =   735
         Width           =   1350
      End
      Begin VB.Image imgSkill 
         Height          =   150
         Left            =   990
         Picture         =   "frmMain.frx":2D59
         Top             =   3990
         Width           =   225
      End
      Begin VB.Label lblCha 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Character"
         Height          =   225
         Left            =   60
         TabIndex        =   25
         Top             =   195
         Width           =   1350
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000014&
         X1              =   60
         X2              =   1360
         Y1              =   2565
         Y2              =   2565
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000010&
         X1              =   60
         X2              =   1360
         Y1              =   2550
         Y2              =   2550
      End
      Begin VB.Label lblAlias 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "- Alias -"
         Height          =   225
         Left            =   60
         TabIndex        =   31
         Top             =   465
         Width           =   1350
      End
      Begin VB.Label lblEP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Experience Points"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   90
         TabIndex        =   29
         Top             =   4305
         Width           =   1290
      End
      Begin VB.Label lblExp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   600
         TabIndex        =   28
         Top             =   4575
         Width           =   180
      End
      Begin VB.Label lblLevel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   1260
         TabIndex        =   27
         Top             =   2295
         Width           =   90
      End
      Begin VB.Label Level 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   60
         TabIndex        =   26
         Top             =   2295
         Width           =   390
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   6
         Left            =   60
         TabIndex        =   24
         Top             =   3960
         Width           =   315
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dexterity"
         Height          =   225
         Index           =   5
         Left            =   60
         TabIndex        =   23
         Top             =   3735
         Width           =   690
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agility"
         Height          =   225
         Index           =   4
         Left            =   60
         TabIndex        =   22
         Top             =   3510
         Width           =   510
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Constitution"
         Height          =   225
         Index           =   3
         Left            =   60
         TabIndex        =   21
         Top             =   3285
         Width           =   915
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Strengh"
         Height          =   225
         Index           =   2
         Left            =   60
         TabIndex        =   20
         Top             =   3060
         Width           =   540
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perception"
         Height          =   225
         Index           =   1
         Left            =   60
         TabIndex        =   19
         Top             =   2835
         Width           =   780
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Knowledge"
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   18
         Top             =   2610
         Width           =   810
      End
      Begin VB.Label lblVal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   6
         Left            =   1260
         TabIndex        =   17
         Top             =   3960
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   5
         Left            =   1260
         TabIndex        =   16
         Top             =   3735
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   1260
         TabIndex        =   15
         Top             =   3510
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   1260
         TabIndex        =   14
         Top             =   3285
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   1260
         TabIndex        =   13
         Top             =   3060
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   1260
         TabIndex        =   12
         Top             =   2835
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   1260
         TabIndex        =   11
         Top             =   2610
         Width           =   90
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000010&
         X1              =   60
         X2              =   1360
         Y1              =   4245
         Y2              =   4245
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000014&
         X1              =   60
         X2              =   1360
         Y1              =   4260
         Y2              =   4260
      End
   End
   Begin VB.ListBox lstSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Columns         =   4
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      ItemData        =   "frmMain.frx":2DE2
      Left            =   30
      List            =   "frmMain.frx":2DE4
      TabIndex        =   10
      Top             =   5220
      Width           =   5325
   End
   Begin VB.Frame fraBar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   30
      TabIndex        =   5
      Top             =   5820
      Width           =   5340
      Begin VB.Label lblOpt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Talk"
         Height          =   240
         Index           =   1
         Left            =   3810
         TabIndex        =   35
         Top             =   615
         Width           =   720
      End
      Begin VB.Label lblCap 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Agi"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   3030
         TabIndex        =   34
         Tag             =   "4"
         Top             =   615
         Width           =   720
      End
      Begin VB.Label lblCap 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Str"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2295
         TabIndex        =   33
         Tag             =   "2"
         Top             =   615
         Width           =   720
      End
      Begin VB.Label lblCap 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kno"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1560
         TabIndex        =   32
         Tag             =   "0"
         Top             =   615
         Width           =   720
      End
      Begin VB.Image imgBar 
         Height          =   720
         Index           =   6
         Left            =   4545
         Picture         =   "frmMain.frx":2DE6
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblOpt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wear"
         Height          =   240
         Index           =   0
         Left            =   795
         TabIndex        =   8
         Top             =   615
         Width           =   720
      End
      Begin VB.Image imgBar 
         Height          =   720
         Index           =   5
         Left            =   3810
         Picture         =   "frmMain.frx":304E
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image imgBar 
         Height          =   720
         Index           =   1
         Left            =   780
         Picture         =   "frmMain.frx":319C
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image imgBar 
         Height          =   720
         Index           =   4
         Left            =   3030
         Picture         =   "frmMain.frx":33AD
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image imgBar 
         Height          =   720
         Index           =   3
         Left            =   2295
         Picture         =   "frmMain.frx":3531
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image imgBar 
         Height          =   720
         Index           =   2
         Left            =   1560
         Picture         =   "frmMain.frx":36B6
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image imgBar 
         Height          =   720
         Index           =   0
         Left            =   45
         Picture         =   "frmMain.frx":3825
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   3795
         X2              =   3795
         Y1              =   135
         Y2              =   835
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   3780
         X2              =   3780
         Y1              =   135
         Y2              =   835
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   1560
         X2              =   1560
         Y1              =   140
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   1545
         X2              =   1545
         Y1              =   140
         Y2              =   840
      End
      Begin VB.Image imgBack 
         Height          =   720
         Left            =   45
         Picture         =   "frmMain.frx":3A5C
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.ListBox lstTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   2
      ItemData        =   "frmMain.frx":4250
      Left            =   4830
      List            =   "frmMain.frx":4257
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   4230
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.ListBox lstTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      ItemData        =   "frmMain.frx":4263
      Left            =   2430
      List            =   "frmMain.frx":426A
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   4230
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.ListBox lstTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":4276
      Left            =   30
      List            =   "frmMain.frx":427D
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   4230
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.TextBox txtMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4875
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmMain.frx":4289
      Top             =   30
      Width           =   7200
   End
   Begin MSComctlLib.ImageList imlTiles 
      Left            =   8160
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   39
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4651
            Key             =   "E"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":49A1
            Key             =   "ES"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4CFA
            Key             =   "N"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5052
            Key             =   "NE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53A5
            Key             =   "NES"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":56FF
            Key             =   "NS"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A5E
            Key             =   "S"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5DB6
            Key             =   "W"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6105
            Key             =   "WE"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":644F
            Key             =   "WES"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":67A2
            Key             =   "WN"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AF9
            Key             =   "WNE"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6E4B
            Key             =   "WNES"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71A2
            Key             =   "WNS"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":74FE
            Key             =   "WS"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgAni 
      Height          =   390
      Index           =   3
      Left            =   8430
      Picture         =   "frmMain.frx":7857
      Top             =   4920
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgAni 
      Height          =   390
      Index           =   2
      Left            =   8040
      Picture         =   "frmMain.frx":7945
      Top             =   4920
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgAni 
      Height          =   390
      Index           =   1
      Left            =   7650
      Picture         =   "frmMain.frx":7A7E
      Top             =   4920
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgAni 
      Height          =   390
      Index           =   0
      Left            =   7260
      Picture         =   "frmMain.frx":7BB8
      Top             =   4920
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgLook 
      Height          =   450
      Left            =   7770
      Picture         =   "frmMain.frx":7CD6
      Top             =   5730
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgUp 
      Height          =   450
      Left            =   7770
      Picture         =   "frmMain.frx":80CE
      Top             =   5730
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgDown 
      Height          =   450
      Left            =   7770
      Picture         =   "frmMain.frx":84C6
      Top             =   5730
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgArrow 
      Height          =   450
      Index           =   5
      Left            =   8220
      Picture         =   "frmMain.frx":8662
      Tag             =   "You move East."
      Top             =   5730
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   5430
      TabIndex        =   7
      Top             =   4950
      Width           =   1800
   End
   Begin VB.Image imgArrow 
      Height          =   450
      Index           =   1
      Left            =   7770
      Picture         =   "frmMain.frx":871A
      Tag             =   "You move South."
      Top             =   6180
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgArrow 
      Height          =   450
      Index           =   7
      Left            =   7770
      Picture         =   "frmMain.frx":87DB
      Tag             =   "You move North."
      Top             =   5280
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgArrow 
      Height          =   450
      Index           =   3
      Left            =   7320
      Picture         =   "frmMain.frx":889A
      Tag             =   "You move West."
      Top             =   5730
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label lblSel 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   75
      TabIndex        =   6
      Top             =   4950
      Width           =   5250
   End
   Begin VB.Image imgRing 
      Height          =   1350
      Left            =   7320
      Picture         =   "frmMain.frx":8952
      Top             =   5280
      Width           =   1350
   End
   Begin VB.Menu mnuFil 
      Caption         =   "&File"
      Begin VB.Menu mnuFilNew 
         Caption         =   "&New Quest..."
      End
      Begin VB.Menu mnuFilBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilExt 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuVie 
      Caption         =   "&View"
      Begin VB.Menu mnuVieHal 
         Caption         =   "&Hall of Fame..."
      End
   End
   Begin VB.Menu mnuToo 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTooMak 
         Caption         =   "Quest-Maker..."
      End
   End
   Begin VB.Menu mnuHel 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelAbo 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuInv 
      Caption         =   "Inventory"
      Visible         =   0   'False
      Begin VB.Menu mnuInvRem 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuInvBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInvObj 
         Caption         =   "<Empty>"
         Index           =   0
      End
      Begin VB.Menu mnuInvBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInvPic 
         Caption         =   "Pick Up"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Action$, strList As String
Dim r%, c%, Rst$, mnu As Menu
Dim strBuffer As String, strR As String, strFinal As String

Private Sub mnuFilNew_Click()
  frmNew.Show
End Sub

Private Sub mnuVieHal_Click()
  frmHall.Show
End Sub

Private Sub mnuTooMak_Click()
  frmMaker.Show
End Sub

Private Sub mnuHelAbo_Click()
  frmAbout.Show
End Sub

Private Sub mnuInvRem_Click()
  For Each mnu In mnuInvObj
    If mnu.Checked = True Then lstSel.AddItem mnu.Caption
  Next mnu
  
  If lstSel.ListCount = 0 Then lblSel = "You are wearing nothing" Else lblSel = "Remove"
  Action$ = "Remove"
End Sub

Private Sub mnuInvObj_Click(Index As Integer)
  If Index = 0 Then lblSel = "You are carrying nothing": Exit Sub
  If mnuInvObj(Index).Checked = False Then lblSel = "You Drop the " & mnuInvObj(Index).Caption Else lblSel = "You are wearing the " & mnuInvObj(Index).Caption: Exit Sub
  If mnuInvObj.Count = 2 Then mnuInvObj(0).Visible = True
  lstTmp(0).AddItem mnuInvObj(Index).Caption
  Unload mnuInvObj(Index)
End Sub

Private Sub mnuInvPic_Click()
  For i% = 0 To lstTmp(0).ListCount - 1
    If ReadINI("Element", lstTmp(0).List(i%), strFree) = "" Then lstSel.AddItem lstTmp(0).List(i%)
  Next i%
  
  If lstSel.ListCount = 0 Then lblSel = "There is nothing of interest here" Else lblSel = "Pick up"
  Action$ = "Pick"
End Sub

Private Sub mnuFilExt_Click()
  End
End Sub

Private Sub tmrAni_Timer()
  imgGuy.Picture = imgAni(0).Picture
  tmrAni.Enabled = False
End Sub

Private Sub txtMain_Change()
  txtMain.SelStart = Len(txtMain)
End Sub

Private Sub txtMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MouseOver 7
End Sub

Private Sub imgBar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  MouseOver Index
End Sub

Private Sub MouseOver(Index As Integer)
  For i% = 0 To 6
    If i% = Index Then imgBar(i%).BorderStyle = 1 Else imgBar(i%).BorderStyle = 0
  Next i%
End Sub

Private Sub picHP_Click()
  Dim iDiff As Integer
  If picHP.Width < 90 And lblExp > 0 Then iDiff = 90 - picHP.Width Else Exit Sub
  If iDiff > lblExp Then picHP.Width = picHP.Width + lblExp: lblExp = 0 Else picHP.Width = 90: lblExp = lblExp - iDiff
  txtMain = txtMain & vbCrLf & vbCrLf & "Health." & vbCrLf & "( " & Val(picHP.Width / 0.9) & "% Health )"
End Sub

Private Sub imgLook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lstSel.Clear: lblSel = ""
  imgLook.Picture = imgDown.Picture
  If Action$ = "Look" Then Exit Sub Else Action$ = "Look"
  
  varTwo = Split(ReadINI("Room", "Dir", strFree), "|")
  varTmp = Split(ReadINI("Room", "Title", strFree), "|")
  For i% = 0 To UBound(varTwo)
    If varTwo(i%) = strR Then lblTitle = varTmp(i%)
  Next i%
  
  varTmp = ReadINI("Description", strR, strFree)
  txtMain = txtMain & Replace("||Look.|" & varTmp, "|", vbCrLf)
  
  If lstTmp(0).ListCount + lstTmp(1).ListCount + lstTmp(2).ListCount = 0 Then txtMain = txtMain & " This room has been thoroughly searched"
  
  For i% = 0 To lstTmp(0).ListCount - 1
    If i% = 0 Then varTmp = " Also here is a " Else varTmp = " and a "
    txtMain = txtMain & varTmp & lstTmp(0).List(i%)
  Next i%
  
  For i% = 0 To lstTmp(1).ListCount - 1
    txtMain = txtMain & " " & lstTmp(1).List(i%) & " is here. "
  Next i%
End Sub

Private Sub imgLook_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgLook.Picture = imgUp.Picture
End Sub

Private Sub lblOpt_Click(Index As Integer)
  If Index = 0 Then If lblOpt(0) = "Wear" Then lblOpt(0) = "Use" Else lblOpt(0) = "Wear"
  If Index = 1 Then If lblOpt(1) = "Talk" Then lblOpt(1) = "Give" Else lblOpt(1) = "Talk"
End Sub

Private Sub LoadQuest(strIntro As String)
  For i% = 0 To 2
    lstTmp(i%).Clear
  Next i%

  If strIntro <> "" Then txtMain = Replace(ReadINI("Starting", "Info", strFree), "|", vbCrLf): strBuffer = ReadINI("Starting", "Buffer", strFree)
  
  strFinal = ReadINI("Ending", "Encounter", strFree)
  Rst$ = ReadINI("Restriction", "S|" & strR, strFree)
  varTmp = ReadINI("Entrance", strR, strFree)
  If varTmp <> "" Then txtMain = Replace(txtMain & "|" & varTmp, "|", vbCrLf)
  varTmp = ReadINI("Movement", strR, strFree)
  If varTmp <> "" Then SplitAndAdd lstTmp(1), varTmp
  varTmp = Split(ReadINI("Element", strR, strFree), "|")
  If UBound(varTmp) > -1 Then SplitAndAdd lstTmp(0), varTmp(0): SplitAndAdd lstTmp(1), varTmp(1): SplitAndAdd lstTmp(2), varTmp(2)
End Sub

Private Sub SaveQuest()
  Me.Tag = ""
  For i% = 0 To 2
    For e% = 0 To lstTmp(i%).ListCount - 1
      lstTmp(i%).Tag = lstTmp(i%).Tag & "" & lstTmp(i%).List(e%)
    Next e%
    Me.Tag = Me.Tag & lstTmp(i%).Tag & "|"
    lstTmp(i%).Tag = ""
  Next i%
  
  WriteINI "Entrance", strR, "", strFree
  WriteINI "Movement", strR, "", strFree
  WriteINI "Element", strR, Me.Tag, strFree
  WriteINI "Restriction", "S|" & strR, Rst$, strFree
End Sub

Private Sub ShowMap(bCase As Boolean)
  Dim iNum As Integer, sT As String
  
  For i% = c% - 1 To c% + 1
    For e% = r% - 1 To r% + 1
      If ReadINI("Room", e% - 1 & "," & i%, strFree) <> "" Then sT = "W" Else sT = ""
      If ReadINI("Room", e% & "," & i% + 1, strFree) <> "" Then sT = sT & "N"
      If ReadINI("Room", e% + 1 & "," & i%, strFree) <> "" Then sT = sT & "E"
      If ReadINI("Room", e% & "," & i% - 1, strFree) <> "" Then sT = sT & "S"
      For a% = 1 To imlTiles.ListImages.Count
        If imlTiles.ListImages(a%).Key = sT Then pic(iNum).Picture = imlTiles.ListImages(a%).Picture
      Next a%
      varTmp = ReadINI("Room", e% & "," & i%, strFree)
      If varTmp <> "" Or iNum = 4 Then pic(iNum).Visible = bCase Else pic(iNum).Visible = False
      If i% = c% Xor e% = r% Then If varTmp <> "" Then imgArrow(iNum).Visible = True Else imgArrow(iNum).Visible = False
      iNum = iNum + 1
    Next e%
  Next i%
End Sub

Private Sub imgArrow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  lstSel.Clear: lblSel = ""
  tmrAni.Enabled = False
  varTmp = Split(Rst$, "|")
  imgArrow(Index).Visible = False
  If varTmp(Index \ 2) <> "" Then lblSel = varTmp(Index \ 2): Exit Sub
  lblTitle = "": Action$ = ""
  SaveQuest
  If Index = 7 Then c% = c% + 1
  If Index = 5 Then r% = r% + 1
  If Index = 3 Then r% = r% - 1
  If Index = 1 Then c% = c% - 1
  strR = r% & "," & c%
  imgGuy.Picture = imgAni(Index \ 2).Picture
  txtMain = txtMain & vbCrLf & vbCrLf & imgArrow(Index).Tag
  LoadQuest ""
End Sub

Private Sub imgArrow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  tmrAni.Enabled = True
  ShowMap True
End Sub

Private Sub imgBar_Click(Index As Integer)
  lstSel.Clear: lblSel = ""
  
  Select Case Index
    Case 0
      If mnuInvObj.Count = 1 Then mnuInvRem.Enabled = False Else mnuInvRem.Enabled = True
      If mnuInvObj.Count = 6 Then mnuInvPic.Enabled = False Else mnuInvPic.Enabled = True
      PopupMenu mnuInv, vbPopupMenuLeftAlign, fraBar.Left, fraBar.Top + fraBar.Height
      
    Case 1
      If mnuInvObj.Count = 1 Then lblSel = "You are carrying nothing": Exit Sub Else lblSel = lblOpt(0): Action$ = lblOpt(0)
      If lblOpt(0) = "Use" Then
        For Each mnu In mnuInvObj
          If mnu.Index > 0 And mnu.Checked = False Then lstSel.AddItem mnu.Caption
        Next mnu

        If lstTmp(0).ListCount = 0 Then lblSel = "There is nothing to Use": lstSel.Clear
      Else
        For Each mnu In mnuInvObj
          If mnu.Index > 0 Then If mnu.Checked = False Then If ReadINI("Attributes", mnu.Caption, strFree) <> "" Then lstSel.AddItem mnu.Caption
        Next mnu
      End If
      If lstSel.ListCount = 0 Then lblSel = "There is nothing to " & lblOpt(0)
      
    Case 2, 3, 4
      ListToList lstSel, lstTmp(2)
      For i% = 0 To lstTmp(1).ListCount - 1
        varTmp = ReadINI("Resolve", lstTmp(1).List(i%), strFree)
        If varTmp <> "" Then lstSel.AddItem varTmp
      Next i%
      If lstSel.ListCount = 0 Then lblSel = "No Tasks required" Else lblSel = "[ " & lblCap(Index) & " ]" & " Task"
      frmRoll.iSta = lblCap(Index).Tag
      Action$ = "Resolve"
    
    Case 5
      If lstTmp(1).ListCount = 0 Then lblSel = "There is nobody here": Exit Sub Else lblSel = lblOpt(1): Action$ = lblOpt(1)
      If lblOpt(1) = "Give" Then
        If mnuInvObj.Count = 1 Then lblSel = "You are carrying nothing"
      
        For Each mnu In mnuInvObj
          If mnu.Index > 0 And mnu.Checked = False Then lstSel.AddItem mnu.Caption
        Next mnu
      Else
        ListToList lstSel, lstTmp(1)
      End If
    
    Case 6
      ListToList lstSel, lstTmp(0): ListToList lstSel, lstTmp(1)
      If lstSel.ListCount = 0 Then lblSel = "There is nothing of interest here" Else lblSel = "Examine"
      Action$ = "Examine"
  End Select
End Sub

Private Sub lstSel_Click()
  Static iMenu As Integer
  strList = lstSel.Text
  lstSel.Clear: lblSel = ""
  If strList = strFinal Then FormInit False: frmFinal.Show
  
  Select Case Action$
    Case "Pick"
      iMenu = iMenu + 1
      Load mnuInvObj(iMenu)
      mnuInvObj(0).Visible = False
      mnuInvObj(iMenu).Visible = True
      mnuInvObj(iMenu).Caption = strList
      
      lblSel = "You Pick up the " & strList
      FindAndRemove lstTmp(0), strList
    
    Case "Remove"
      varTmp = Split(ReadINI("Attributes", strList, strFree), "")
      txtMain = txtMain & vbCrLf & vbCrLf & "You Remove the " & strList & vbCrLf & "( " & lblSta(varTmp(1)) & " -" & varTmp(0) & " )"
      lblVal(varTmp(1)) = lblVal(varTmp(1)) - Val(varTmp(0))
      
      For Each mnu In mnuInvObj
        If mnu.Caption = strList Then mnu.Checked = False
      Next mnu

    Case "Wear"
      varTmp = Split(ReadINI("Attributes", strList, strFree), "")
      If UBound(varTmp) = -1 Then lblSel = "You can't wear that": Exit Sub
      lblVal(varTmp(1)) = lblVal(varTmp(1)) + Val(varTmp(0))
      Me.Tag = ReadINI("Attributes", "S|" & strList, strFree)
      If Me.Tag <> "" Then Me.Tag = vbCrLf & Me.Tag Else Me.Tag = ""
      txtMain = txtMain & vbCrLf & vbCrLf & "You Put on the " & strList & Me.Tag & vbCrLf & "( " & lblSta(varTmp(1)) & " +" & varTmp(0) & " )"
    
      For Each mnu In mnuInvObj
        If mnu.Caption = strList Then mnu.Checked = True
      Next mnu
      
    Case "Talk"
      varTmp = ReadINI("Dialog", strList, strFree)
      If varTmp <> "" Then txtMain = txtMain & vbCrLf & vbCrLf & "<" & lblCha & "> " & varTmp
      varTmp = ReadINI("Dialog", "S|" & strList, strFree)
      If varTmp <> "" Then txtMain = txtMain & vbCrLf & vbCrLf & "<" & strList & ">" & varTmp
      If varTmp = "" Then lblSel = strList & " does not respond"
      If ReadINI("Dialog", "|" & strList, strFree) = "" Then WriteINI "Dialog", strList, "", strFree: WriteINI "Dialog", "S|" & strList, "", strFree
      
    Case "Give"
      lblSel.Tag = strList
      lblSel = "Give the " & strList & " to"
      ListToList lstSel, lstTmp(1)
      Action$ = "Trade"

    Case "Use"
      lblSel.Tag = strList
      lblSel = "Use the " & strList & " with"
      ListToList lstSel, lstTmp(0)
      Action$ = "Trade"
    
    Case "Trade"
      If ReadINI("Trade", lblSel.Tag, strFree) <> strList Then Else lblSel = "The object is rejected": Exit Sub
      txtMain = txtMain & vbCrLf & vbCrLf & strList & vbCrLf & ReadINI("Trade", "S|" & lblSel.Tag, strFree)
      varTmp = ReadINI("Trade", strList, strFree)
      If varTmp <> "" Then lstTmp(0).AddItem varTmp: txtMain = txtMain & " Here is a " & varTmp
      
      If mnuInvObj.Count = 2 Then mnuInvObj(0).Visible = True

      For Each mnu In mnuInvObj
        If mnu.Caption = lblSel.Tag Then Unload mnu
      Next mnu
    
    Case "Resolve"
      varTmp = ReadINI("Element", strList, strFree)
      If varTmp <> "" Then lstTmp(2).AddItem varTmp
      If InStr(strBuffer, "" & strList & "") = 0 Then frmRoll.Show Else Score 0
        
    Case "Examine"
      varTmp = ReadINI("Description", strList, strFree)
      If varTmp = "" Then lblSel = "Nothing special" Else txtMain = txtMain & vbCrLf & vbCrLf & strList & "." & vbCrLf & varTmp
  End Select
End Sub

Private Sub ChangeStat(iNum As Integer)
  If iNum < 6 Then If iNum = 5 Then lblCap(4).Tag = 4: lblCap(4) = "Agi" Else lblCap(4).Tag = 5: lblCap(4) = "Dex"
  If iNum < 4 Then If iNum = 3 Then lblCap(3).Tag = 2: lblCap(3) = "Str" Else lblCap(3).Tag = 3: lblCap(3) = "Con"
  If iNum < 2 Then If iNum = 1 Then lblCap(2).Tag = 0: lblCap(2) = "Kno" Else lblCap(2).Tag = 1: lblCap(2) = "Per"
End Sub

Public Sub Score(iSum As Integer)
  If iSum > 1 Then ChangeStat frmRoll.iSta
  Select Case iSum
    Case Is >= 0
      FindAndRemove lstTmp(2), strList
      strBuffer = strBuffer & "" & strList & ""
      txtMain = txtMain & vbCrLf & vbCrLf & strList & "."
      AfterTask strList
      If iSum > 0 Then lblExp = lblExp + iSum: txtMain = txtMain & vbCrLf & "( +" & iSum & " Exp. Points )"
      If lblLevel = (lblExp \ 30) Then lblLevel = lblLevel + 1: lblVal(6) = lblVal(6) + 1

    Case Is < 0
      FindAndRemove lstTmp(2), ReadINI("Element", strList, strFree)
      txtMain = txtMain & vbCrLf & vbCrLf & "Failed Intent." & vbCrLf & "( " & iSum & " Hit Points )"
      If iSum < picHP.Width Then picHP.Width = picHP.Width + iSum: Exit Sub Else picHP.Width = 0: FormInit False
    
      varTmp = ReadINI("Ending", "Lost", strFree)
      txtMain = Replace("|" & varTmp, "|", vbCrLf)
  End Select
End Sub

Private Sub AfterTask(strTsk As String)
  Dim varElm As Variant
  varTmp = ReadINI("Resolve", "S|" & strTsk, strFree)
  If varTmp <> "" Then txtMain = txtMain & vbCrLf & varTmp
  varTmp = Split(ReadINI("Restriction", strR, strFree), "|")
  varTwo = Split(ReadINI("Restriction", "S|" & strR, strFree), "|")
  For i% = 0 To UBound(varTmp)
    If strTsk = varTmp(i%) Then varTwo(i%) = ""
  Next i%
  Rst$ = Join(varTwo, "|")
  
  For i% = 0 To lstTmp(1).ListCount - 1
    If strTsk = ReadINI("Resolve", lstTmp(1).List(i%), strFree) Then lstTmp(1).RemoveItem i%
  Next i%
  
  varTmp = ReadINI("ShowUp", strTsk, strFree)
  If varTmp = "" Then varTmp = Split("||", "|") Else varTmp = Split(ReadINI("ShowUp", strTsk, strFree), "|"): WriteINI "ShowUp", strTsk, "", strFree
  For i% = 2 To 0 Step -1
    varElm = Split(varTmp(i%), "")
    If UBound(varElm) > -1 Then
      For e% = 1 To UBound(varElm)
        Me.Tag = ReadINI("ShowUp", "S|" & varElm(e%), strFree)
        If Me.Tag <> "" Then txtMain = txtMain & " " & Me.Tag
        lstTmp(i%).AddItem varElm(e%)
      Next e%
    End If
  Next i%

  varTmp = Split(ReadINI("Movement", strTsk, strFree), ""): WriteINI "Movement", strTsk, "", strFree
  If UBound(varTmp) = 1 Then
    FindAndRemove lstTmp(1), CStr(varTmp(0))
    Me.Tag = ReadINI("Movement", varTmp(1), strFree) & "" & varTmp(0)
    WriteINI "Movement", CStr(varTmp(1)), Me.Tag, strFree
    Me.Tag = ReadINI("Movement", "S|" & strTsk, strFree)
    If Me.Tag <> "" Then txtMain = txtMain & vbCrLf & Me.Tag
  End If
End Sub

Private Sub FormInit(bCase As Boolean)
  imgBack.Visible = Not bCase
  lstSel.Clear: lblSel = ""
  imgLook.Visible = bCase
  ShowMap bCase
  lblTitle = ""
  Action$ = ""
  
  For i% = 0 To 6
    imgBar(i%).Visible = bCase
  Next i%
  
  For Each mnu In mnuInvObj
    If mnu.Index = 0 Then mnu.Visible = True Else Unload mnu
  Next mnu
End Sub

Public Sub Starting(strPath As String, strName As String)
  picHP.Width = 90
  lblLevel = 1
  lblExp = 0
  
On Error Resume Next
  strName = App.Path & "\Data\" & strName
  strFree = App.Path & "\Free.tmp"
  FileCopy strPath, strFree
  
  varTmp = Split(ReadINI("Starting", "Room", strFree), ",")
  r% = varTmp(0): c% = varTmp(1): strR = r% & "," & c%
  
  varTmp = Split(ReadINI("Character", "Data", strName), "")
  lblCha = varTmp(0): frmFinal.lblCha(0) = lblCha: lblAlias = varTmp(1)
  lblSta(6) = varTmp(2): frmFinal.lblSta(0) = lblSta(6): picCha.Tag = varTmp(3)
  picCha.Picture = LoadPicture(picCha.Tag): frmFinal.picCha(0) = picCha

  varTmp = Split(ReadINI("Character", "Stats", strName), "|")
  For i% = 0 To 6
    lblVal(i%) = varTmp(i%)
  Next i%
  
  LoadQuest "Intro"
  FormInit True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  mnuFilExt_Click
End Sub
