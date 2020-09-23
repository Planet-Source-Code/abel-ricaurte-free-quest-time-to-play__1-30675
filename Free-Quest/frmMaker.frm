VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMaker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: Quest-Maker ::"
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
   Icon            =   "frmMaker.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   580
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   2430
      TabIndex        =   38
      Top             =   5220
      Width           =   2925
   End
   Begin VB.TextBox txtRst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2430
      TabIndex        =   37
      Top             =   6390
      Width           =   2925
   End
   Begin VB.TextBox txtRst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2430
      TabIndex        =   36
      Top             =   6000
      Width           =   2925
   End
   Begin MSComDlg.CommonDialog cdlShow 
      Left            =   5520
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.ComboBox cmbRst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      ItemData        =   "frmMaker.frx":030A
      Left            =   750
      List            =   "frmMaker.frx":0311
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   5220
      Width           =   1650
   End
   Begin VB.ComboBox cmbRst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      ItemData        =   "frmMaker.frx":031D
      Left            =   750
      List            =   "frmMaker.frx":0324
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   6390
      Width           =   1650
   End
   Begin VB.ComboBox cmbRst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      ItemData        =   "frmMaker.frx":0330
      Left            =   750
      List            =   "frmMaker.frx":0337
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   6000
      Width           =   1650
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000007&
      Height          =   1500
      Left            =   5415
      Picture         =   "frmMaker.frx":0343
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   18
      Top             =   5220
      Width           =   1815
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         Index           =   2
         Left            =   1200
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   25
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
         Index           =   3
         Left            =   30
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   510
         Visible         =   0   'False
         Width           =   585
         Begin VB.Image imgGuy 
            Height          =   390
            Left            =   135
            Picture         =   "frmMaker.frx":19D5
            Top             =   45
            Width           =   270
         End
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
         TabIndex        =   22
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
         Index           =   6
         Left            =   30
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   21
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
         TabIndex        =   20
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
         Index           =   8
         Left            =   1200
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   19
         Top             =   30
         Visible         =   0   'False
         Width           =   585
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
      Index           =   3
      ItemData        =   "frmMaker.frx":1AED
      Left            =   60
      List            =   "frmMaker.frx":1AF4
      TabIndex        =   17
      Top             =   4230
      Visible         =   0   'False
      Width           =   2280
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
      Index           =   4
      ItemData        =   "frmMaker.frx":1B00
      Left            =   2340
      List            =   "frmMaker.frx":1B07
      TabIndex        =   16
      Top             =   4230
      Visible         =   0   'False
      Width           =   2280
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
      Index           =   5
      ItemData        =   "frmMaker.frx":1B13
      Left            =   4620
      List            =   "frmMaker.frx":1B1A
      TabIndex        =   15
      Top             =   4230
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.TextBox txtRst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   2430
      TabIndex        =   11
      Top             =   5610
      Width           =   2925
   End
   Begin VB.ComboBox cmbRst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      ItemData        =   "frmMaker.frx":1B26
      Left            =   750
      List            =   "frmMaker.frx":1B2D
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   5610
      Width           =   1650
   End
   Begin VB.TextBox txtRoom 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5700
      TabIndex        =   4
      Top             =   4920
      Width           =   1500
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
      ItemData        =   "frmMaker.frx":1B39
      Left            =   4620
      List            =   "frmMaker.frx":1B40
      TabIndex        =   3
      Top             =   3570
      Visible         =   0   'False
      Width           =   2280
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
      ItemData        =   "frmMaker.frx":1B4C
      Left            =   2340
      List            =   "frmMaker.frx":1B53
      TabIndex        =   2
      Top             =   3570
      Visible         =   0   'False
      Width           =   2280
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
      ItemData        =   "frmMaker.frx":1B5F
      Left            =   60
      List            =   "frmMaker.frx":1B66
      TabIndex        =   1
      Top             =   3570
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.TextBox txtMain 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   2430
      Index           =   1
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   7
      Tag             =   "Entrance"
      Top             =   30
      Width           =   7200
   End
   Begin VB.Frame Frame1 
      Height          =   4995
      Left            =   7260
      TabIndex        =   6
      Top             =   -75
      Width           =   1440
      Begin VB.CommandButton cmdElm 
         Caption         =   "Task"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   15
         TabIndex        =   10
         Top             =   4680
         Width           =   1405
      End
      Begin VB.CommandButton cmdElm 
         Caption         =   "Character"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   15
         TabIndex        =   9
         Top             =   4395
         Width           =   1405
      End
      Begin VB.CommandButton cmdElm 
         Caption         =   "Object"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   15
         TabIndex        =   8
         Top             =   120
         Width           =   1405
      End
      Begin VB.PictureBox picBack 
         BorderStyle     =   0  'None
         Height          =   3930
         Left            =   30
         ScaleHeight     =   262
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   92
         TabIndex        =   12
         Top             =   450
         Width           =   1380
         Begin VB.ListBox lstSel 
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
            Height          =   3345
            Left            =   0
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   480
            Width           =   1380
         End
         Begin VB.Image imgTab 
            Height          =   315
            Index           =   2
            Left            =   975
            Picture         =   "frmMaker.frx":1B72
            ToolTipText     =   "Delete"
            Top             =   60
            Width           =   315
         End
         Begin VB.Image imgTab 
            Height          =   315
            Index           =   1
            Left            =   510
            Picture         =   "frmMaker.frx":1C23
            ToolTipText     =   "Edit"
            Top             =   60
            Width           =   315
         End
         Begin VB.Image imgTab 
            Height          =   315
            Index           =   0
            Left            =   30
            Picture         =   "frmMaker.frx":1CFA
            ToolTipText     =   "Add"
            Top             =   60
            Width           =   315
         End
      End
   End
   Begin VB.TextBox txtMain 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   2430
      Index           =   0
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   0
      Tag             =   "Description"
      Top             =   2475
      Width           =   7200
   End
   Begin VB.Image imgOn 
      Height          =   300
      Left            =   5400
      Picture         =   "frmMaker.frx":1DA9
      ToolTipText     =   "Start"
      Top             =   4920
      Width           =   285
   End
   Begin VB.Label lblMove 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add/Move"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7320
      TabIndex        =   39
      Top             =   4950
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   7320
      Picture         =   "frmMaker.frx":1E92
      Top             =   4950
      Width           =   1350
   End
   Begin VB.Image imgRem 
      Height          =   450
      Left            =   7770
      Picture         =   "frmMaker.frx":1F30
      Top             =   5730
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image img 
      Height          =   450
      Index           =   1
      Left            =   7770
      Picture         =   "frmMaker.frx":251C
      Tag             =   "South Room"
      Top             =   6180
      Width           =   450
   End
   Begin VB.Image img 
      Height          =   450
      Index           =   3
      Left            =   7320
      Picture         =   "frmMaker.frx":25DD
      Tag             =   "West Room"
      Top             =   5730
      Width           =   450
   End
   Begin VB.Image img 
      Height          =   450
      Index           =   5
      Left            =   8220
      Picture         =   "frmMaker.frx":2695
      Tag             =   "East Room"
      Top             =   5730
      Width           =   450
   End
   Begin VB.Image img 
      Height          =   450
      Index           =   7
      Left            =   7770
      Picture         =   "frmMaker.frx":274D
      Tag             =   "North Room"
      Top             =   5280
      Width           =   450
   End
   Begin VB.Image imgRing 
      Height          =   1350
      Left            =   7320
      Picture         =   "frmMaker.frx":280C
      Top             =   5280
      Width           =   1350
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7260
      TabIndex        =   35
      Top             =   4920
      Width           =   1380
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "East"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      TabIndex        =   34
      Top             =   5610
      Width           =   675
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "North"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      TabIndex        =   33
      Top             =   5220
      Width           =   675
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "West"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      TabIndex        =   32
      Top             =   6390
      Width           =   675
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "South"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      TabIndex        =   28
      Top             =   6000
      Width           =   675
   End
   Begin VB.Label lblAct 
      BackStyle       =   0  'Transparent
      Caption         =   "Restriction                     Show"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   750
      TabIndex        =   5
      Top             =   4950
      Width           =   2250
   End
   Begin VB.Menu mnuFil 
      Caption         =   "&File"
      Begin VB.Menu mnuFilNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFilOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFilBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilSav 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFilSas 
         Caption         =   "Sa&ve As..."
      End
      Begin VB.Menu mnuFilBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilExt 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuQue 
      Caption         =   "&Quest"
      Begin VB.Menu mnuQueOpt 
         Caption         =   "S&tarting..."
         Index           =   3
      End
      Begin VB.Menu mnuQueOpt 
         Caption         =   "O&pponent..."
         Index           =   4
      End
      Begin VB.Menu mnuQueOpt 
         Caption         =   "E&nding..."
         Index           =   5
      End
   End
   Begin VB.Menu mnuRet 
      Caption         =   "&Return"
   End
End
Attribute VB_Name = "frmMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iElm As Integer, strR As String
Dim r%, c%

Private Sub Form_Load()
  strFile = App.Path & "\Maker.tmp"
  mnuFilNew_Click
  frmMain.Hide
End Sub

Private Sub mnuFilNew_Click()
On Error Resume Next
  Kill strFile
  cdlShow.filename = ""
  LoadQuest "Intro"
End Sub

Private Sub mnuFilOpen_Click()
On Error Resume Next
  cdlShow.Filter = "Free-Quest Data (*.dat)|*.dat"
  cdlShow.InitDir = App.Path
  cdlShow.ShowOpen
  If Err = 0 Then FileCopy cdlShow.filename, strFile: LoadQuest "Intro"
End Sub

Private Sub mnuFilSav_Click()
  If cdlShow.filename = "" Then mnuFilSas_Click Else SaveQuest: FileCopy strFile, cdlShow.filename
End Sub

Private Sub mnuFilSas_Click()
On Error Resume Next
  cdlShow.Filter = "Free-Quest Data (*.dat)|*.dat"
  cdlShow.Flags = cdlOFNOverwritePrompt
  cdlShow.InitDir = App.Path
  cdlShow.ShowSave
  If Err = 0 Then SaveQuest: FileCopy strFile, cdlShow.filename: Me.Caption = ":: Quest-Maker :: " & cdlShow.FileTitle
End Sub

Private Sub mnuFilExt_Click()
  End
End Sub

Private Sub mnuQueOpt_Click(Index As Integer)
  frmAdd.LoadTab Index, ""
End Sub

Private Sub mnuRet_Click()
  Unload Me
End Sub

Private Sub txtMain_LostFocus(Index As Integer)
  If Trim(Replace(txtMain(Index), vbCrLf, "")) = "" Then txtMain(Index) = txtMain(Index).Tag
End Sub

Private Sub LoadQuest(strIntro As String)
  If cdlShow.filename = "" Then Me.Caption = ":: Quest-Maker :: Untitled" Else Me.Caption = ":: Quest-Maker :: " & cdlShow.FileTitle
  txtRoom = "Unknown"
  For i% = 0 To 5
    lstTmp(i%).Clear
  Next i%
  
  varTmp = Split(ReadINI("Starting", "Room", strFile), ",")
  If strIntro <> "" Then If UBound(varTmp) = 1 Then r% = varTmp(0): c% = varTmp(1) Else r% = 0: c% = 0
  strR = r% & "," & c%
  If ReadINI("Starting", "Room", strFile) = "" Then WriteINI "Starting", "Room", strR, strFile
  varTmp = ReadINI("Description", strR, strFile)
  If varTmp = "" Then txtMain(0) = txtMain(0).Tag Else txtMain(0) = Replace(varTmp, "|", vbCrLf)
  varTmp = ReadINI("Entrance", strR, strFile)
  If varTmp = "" Then txtMain(1) = txtMain(1).Tag Else txtMain(1) = Replace(varTmp, "|", vbCrLf)
  varTmp = Split(ReadINI("Element", strR, strFile), "|")
  If UBound(varTmp) = 6 Then For e% = 0 To 5: SplitAndAdd lstTmp(e%), varTmp(e%): Next e%
  varTmp = Split(ReadINI("Room", "Title", strFile), "|"): varTwo = Split(ReadINI("Room", "Dir", strFile), "|")
  For i% = 0 To UBound(varTwo)
    If varTwo(i%) = strR Then txtRoom = varTmp(i%)
  Next i%
  
  ReloadRest strR
  ShowMap True
  ReloadSel
End Sub

Private Sub SaveQuest()
  Me.Tag = ""
  For i% = 0 To 5
    For e% = 0 To lstTmp(i%).ListCount - 1
      Me.Tag = Me.Tag & "" & lstTmp(i%).List(e%)
    Next e%
    Me.Tag = Me.Tag & "|"
  Next i%
  WriteINI "Element", strR, Me.Tag, strFile

  If InStr(ReadINI("Room", "Dir", strFile), strR & "|") = 0 Then varTwo = Split(ReadINI("Room", "Dir", strFile) & strR & "|", "|"): varTmp = Split(ReadINI("Room", "Title", strFile) & txtRoom & "|", "|") Else varTwo = Split(ReadINI("Room", "Dir", strFile), "|"): varTmp = Split(ReadINI("Room", "Title", strFile), "|")
  For i% = 0 To UBound(varTwo)
    If varTwo(i%) = strR Then varTmp(i%) = txtRoom
  Next i%
  WriteINI "Room", "Dir", Join(varTwo, "|"), strFile
  WriteINI "Room", "Title", Join(varTmp, "|"), strFile
  
  If txtMain(0) = txtMain(0).Tag Then Me.Tag = "" Else Me.Tag = Replace(txtMain(0), vbCrLf, "|")
  WriteINI "Description", strR, Me.Tag, strFile
  If txtMain(1) = txtMain(1).Tag Then Me.Tag = "" Else Me.Tag = Replace(txtMain(1), vbCrLf, "|")
  WriteINI "Entrance", strR, Me.Tag, strFile
  WriteINI "Restriction", strR, cmbRst(0).Tag & "|" & cmbRst(1).Tag & "|" & cmbRst(2).Tag & "|" & cmbRst(3).Tag, strFile
  WriteINI "Restriction", "S|" & strR, txtRst(0).Tag & "|" & txtRst(1).Tag & "|" & txtRst(2).Tag & "|" & txtRst(3).Tag, strFile
  
  ReloadRest strR
  ReloadSel
End Sub

Private Sub ShowMap(bCase As Boolean)
  Dim iNum As Integer, sT As String
  WriteINI "Room", strR, "On", strFile
  
  For i% = c% - 1 To c% + 1
    For e% = r% - 1 To r% + 1
      If ReadINI("Room", e% - 1 & "," & i%, strFile) <> "" Then sT = "W" Else sT = ""
      If ReadINI("Room", e% & "," & i% + 1, strFile) <> "" Then sT = sT & "N"
      If ReadINI("Room", e% + 1 & "," & i%, strFile) <> "" Then sT = sT & "E"
      If ReadINI("Room", e% & "," & i% - 1, strFile) <> "" Then sT = sT & "S"
      For a% = 1 To frmMain.imlTiles.ListImages.Count
        If frmMain.imlTiles.ListImages(a%).Key = sT Then pic(iNum).Picture = frmMain.imlTiles.ListImages(a%).Picture
      Next a%
      varTmp = ReadINI("Room", e% & "," & i%, strFile)
      If varTmp = "" Then pic(iNum).Visible = False Else pic(iNum).Visible = bCase
      If e% = r% Xor i% = c% Then If varTmp = "" Then img(iNum).Visible = False Else img(iNum).Visible = True
      If lblMove = "Add/Move" Then If e% = r% Xor i% = c% Then img(iNum).Visible = True
      iNum = iNum + 1
    Next e%
  Next i%
End Sub

Private Sub img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If lblMove = "Remove" Then RemoveRoom Index: Exit Sub
  SaveQuest
  If Index = 7 Then c% = c% + 1
  If Index = 5 Then r% = r% + 1
  If Index = 3 Then r% = r% - 1
  If Index = 1 Then c% = c% - 1
  img(Index).Visible = False
  LoadQuest ""
End Sub

Private Sub RemoveRoom(Index As Integer)
  Dim strMsg As String, strTmp As String
  strMsg = MsgBox("Delete " & img(Index).Tag & "?", vbOKCancel + vbQuestion)
  If strMsg = vbOK Then
    If Index = 7 Then strTmp = r% & "," & c% + 1
    If Index = 5 Then strTmp = r% + 1 & "," & c%
    If Index = 3 Then strTmp = r% - 1 & "," & c%
    If Index = 1 Then strTmp = r% & "," & c% - 1
    If ReadINI("Starting", "Room", strFile) = strTmp Then WriteINI "Starting", "Room", strR, strFile
    WriteINI "Room", strTmp, "", strFile
  End If
    imgRem.Visible = False
    lblMove = "Add/Move"
    ShowMap True
End Sub

Private Sub img_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  img(Index).Visible = True
End Sub

Private Sub imgTab_Click(Index As Integer)
  Dim iMsg As Integer
  lstSel.Tag = lstSel.List(lstSel.ListIndex)
  If Index > 0 Then If lstSel.ListIndex = -1 Then Exit Sub
  If Left(lstSel.Tag, 3) = "..." Then Me.Tag = Mid(lstSel.Tag, 4) Else Me.Tag = lstSel.Tag
  
  Select Case Index
    Case 0
      frmAdd.LoadTab iElm, ""
      
    Case 1
      frmAdd.LoadTab iElm, lstSel.Tag
      FindAndRemove Me.Tag
      
    Case 2
      iMsg = MsgBox("Delete '" & Me.Tag & "'?", vbOKCancel + vbQuestion)
      If iMsg = vbOK Then FindAndRemove Me.Tag
  End Select
End Sub

Private Sub lstSel_DblClick()
  imgTab_Click (1)
End Sub

Private Sub lblMove_Click()
  If lblMove = "Add/Move" Then lblMove = "Remove" Else lblMove = "Add/Move"
  imgRem.Visible = Not imgRem.Visible: ShowMap True
End Sub

Private Sub cmdElm_Click(Index As Integer)
  If Index = 0 Then cmdElm(1).Top = 4395: cmdElm(2).Top = 4680
  If Index = 1 Then If cmdElm(1).Top = 405 Then cmdElm(2).Top = 4680 Else cmdElm(1).Top = 405
  If Index = 2 Then If cmdElm(2).Top = 4680 Then cmdElm(1).Top = 405: cmdElm(2).Top = 690
 
  picBack.Top = 330 + cmdElm(Index).Top
  picBack.SetFocus
  iElm = Index
  ReloadSel
End Sub

Private Sub lstSel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MouseOver 3
End Sub

Private Sub txtMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  MouseOver 3
End Sub

Private Sub imgTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  MouseOver Index
End Sub

Private Sub MouseOver(Index As Integer)
  For i% = 0 To 2
    If i% = Index Then imgTab(i%).BorderStyle = 1 Else imgTab(i%).BorderStyle = 0
  Next i%
End Sub

Private Sub cmbRst_Click(Index As Integer)
  txtRst(Index).Tag = IIf(txtRst(Index) = "", "That way is blocked", txtRst(Index))
  If cmbRst(Index).ListIndex = 0 Then cmbRst(Index).Tag = "": txtRst(Index).Tag = "": txtRst(Index).Enabled = False Else cmbRst(Index).Tag = cmbRst(Index).Text: txtRst(Index) = txtRst(Index).Tag: txtRst(Index).Enabled = True
End Sub

Private Sub ReloadRest(iR As String)
  For i% = 0 To 3
    For e% = 1 To cmbRst(i%).ListCount - 1
      cmbRst(i%).RemoveItem 1
    Next e%
    txtRst(i%) = "That way is blocked"
    cmbRst(i%).ListIndex = 0
  Next i%

  varTmp = ReadINI("Restriction", iR, strFile): varTwo = ReadINI("Restriction", "S|" & iR, strFile)
  If varTmp = "" Then Exit Sub Else varTmp = Split(varTmp, "|"): varTwo = Split(varTwo, "|")
  For i% = 0 To 3
    cmbRst(i%).Tag = varTmp(i%): txtRst(i%).Tag = varTwo(i%)
  Next i%
  
  For i% = 0 To 3
    For e% = 0 To lstTmp(2).ListCount - 1
      cmbRst(i%).AddItem lstTmp(2).List(e%)
    Next e%
    For e% = 0 To lstTmp(5).ListCount - 1
      cmbRst(i%).AddItem lstTmp(5).List(e%)
    Next e%
  Next i%
  
  For i% = 0 To 3
    For e% = 0 To cmbRst(i%).ListCount - 1
      If cmbRst(i%).Tag = cmbRst(i%).List(e%) Then cmbRst(i%).ListIndex = e%
      If cmbRst(i%).ListIndex > 0 Then txtRst(i%) = txtRst(i%).Tag
    Next e%
  Next i%
End Sub

Private Sub ReloadSel()
  lstSel.Clear
  For i% = 0 To 3 Step 3
    If i% = 0 Then Me.Tag = "" Else Me.Tag = "..."
    For e% = 0 To lstTmp(iElm + i%).ListCount - 1
      lstSel.AddItem Me.Tag & lstTmp(iElm + i%).List(e%)
    Next e%
  Next i%
End Sub

Private Sub FindAndRemove(strFind As String)
  For i% = 0 To 5
    For e% = 0 To lstTmp(i%).ListCount - 1
      If lstTmp(i%).List(e%) = strFind Then lstTmp(i%).RemoveItem e%
    Next e%
  Next i%
  ReloadSel
End Sub

Public Sub AddElement(strAdd As String)
  If Left(strAdd, 3) = "..." Then lstTmp(iElm + 3).AddItem Mid(strAdd, 4) Else lstTmp(iElm).AddItem strAdd
  SaveQuest
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.Show
  Cancel = 1
  Me.Hide
End Sub
