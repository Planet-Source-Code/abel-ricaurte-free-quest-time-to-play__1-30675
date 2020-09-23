VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdd 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   270
      Left            =   3150
      TabIndex        =   65
      Top             =   7200
      Width           =   1425
   End
   Begin VB.CommandButton cmdCnl 
      Caption         =   "Cancel"
      Height          =   270
      Left            =   4575
      TabIndex        =   20
      Top             =   7200
      Width           =   1425
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   7140
      Left            =   150
      TabIndex        =   7
      Top             =   15
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   12594
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   4
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Object"
      TabPicture(0)   =   "frmAdd.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtEle(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkSta(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Character"
      TabPicture(1)   =   "frmAdd.frx":03ED
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkEne"
      Tab(1).Control(1)=   "txtEle(1)"
      Tab(1).Control(2)=   "Frame8"
      Tab(1).Control(3)=   "Frame7"
      Tab(1).Control(4)=   "Frame6"
      Tab(1).Control(5)=   "Frame5"
      Tab(1).Control(6)=   "Label8"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Task"
      TabPicture(2)   =   "frmAdd.frx":0498
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label12"
      Tab(2).Control(1)=   "Frame9"
      Tab(2).Control(2)=   "Frame12"
      Tab(2).Control(3)=   "txtEle(2)"
      Tab(2).Control(4)=   "Frame10"
      Tab(2).Control(5)=   "Frame11"
      Tab(2).Control(6)=   "chkType(0)"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Starting"
      TabPicture(3)   =   "frmAdd.frx":0588
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label23"
      Tab(3).Control(1)=   "Label24"
      Tab(3).Control(2)=   "Label40"
      Tab(3).Control(3)=   "Frame17"
      Tab(3).Control(4)=   "txtTitle"
      Tab(3).Control(5)=   "txtAuthor"
      Tab(3).Control(6)=   "cmbStart"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Opponent"
      TabPicture(4)   =   "frmAdd.frx":0635
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label14"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label41"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label44"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label45"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Frame13"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "txtOpp"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "picOpp"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "txtSkill"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Frame19"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "cmbRoomP"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "txtEnc"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).ControlCount=   11
      TabCaption(5)   =   "Ending"
      TabPicture(5)   =   "frmAdd.frx":06CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame18"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Return"
      TabPicture(6)   =   "frmAdd.frx":07A6
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "txtRet"
      Tab(6).Control(1)=   "Frame16"
      Tab(6).Control(2)=   "Frame15"
      Tab(6).Control(3)=   "Label18"
      Tab(6).ControlCount=   4
      Begin VB.TextBox txtEnc 
         Height          =   315
         Left            =   4050
         TabIndex        =   119
         Text            =   "Final Match"
         Top             =   4080
         Width           =   1800
      End
      Begin VB.ComboBox cmbRoomP 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   117
         Top             =   4080
         Width           =   1800
      End
      Begin VB.CheckBox chkEne 
         Alignment       =   1  'Right Justify
         Caption         =   "Enemy"
         Height          =   315
         Left            =   -71880
         TabIndex        =   116
         Top             =   480
         Width           =   1125
      End
      Begin VB.ComboBox cmbStart 
         Height          =   315
         Left            =   -70950
         Style           =   2  'Dropdown List
         TabIndex        =   114
         Top             =   6480
         Width           =   1800
      End
      Begin VB.CheckBox chkType 
         Alignment       =   1  'Right Justify
         Caption         =   "Routinary"
         Height          =   315
         Index           =   0
         Left            =   -71880
         TabIndex        =   111
         Top             =   480
         Width           =   1125
      End
      Begin VB.Frame Frame19 
         Caption         =   "Dialog"
         Height          =   1470
         Left            =   150
         TabIndex        =   105
         Top             =   2370
         Width           =   6000
         Begin VB.TextBox txtDlgP 
            Height          =   675
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   108
            Top             =   645
            Width           =   4650
         End
         Begin VB.TextBox txtAskP 
            Height          =   315
            Left            =   1050
            TabIndex        =   107
            Top             =   240
            Width           =   1800
         End
         Begin VB.CheckBox chkDlg 
            Alignment       =   1  'Right Justify
            Caption         =   "Repeatedly"
            Height          =   315
            Index           =   1
            Left            =   3000
            TabIndex        =   106
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label43 
            Caption         =   "Reply"
            Height          =   210
            Left            =   210
            TabIndex        =   110
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label42 
            Caption         =   "Player"
            Height          =   210
            Left            =   210
            TabIndex        =   109
            Top             =   300
            Width           =   750
         End
      End
      Begin VB.TextBox txtSkill 
         Height          =   315
         Left            =   4050
         TabIndex        =   101
         Top             =   480
         Width           =   1800
      End
      Begin VB.CheckBox chkSta 
         Alignment       =   1  'Right Justify
         Caption         =   "Stationary"
         Height          =   315
         Index           =   0
         Left            =   -71880
         TabIndex        =   98
         Top             =   480
         Width           =   1125
      End
      Begin VB.Frame Frame11 
         Caption         =   "Movement"
         Height          =   1470
         Left            =   -74850
         TabIndex        =   89
         Top             =   3930
         Width           =   6000
         Begin VB.ComboBox cmbRoom 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   240
            Width           =   1800
         End
         Begin VB.ComboBox cmbMov 
            Height          =   315
            ItemData        =   "frmAdd.frx":08AA
            Left            =   1050
            List            =   "frmAdd.frx":08B1
            Style           =   2  'Dropdown List
            TabIndex        =   91
            Top             =   240
            Width           =   1800
         End
         Begin VB.TextBox txtMov 
            Height          =   675
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   90
            Top             =   645
            Width           =   4650
         End
         Begin VB.Label Label36 
            Caption         =   "Send"
            Height          =   210
            Left            =   210
            TabIndex        =   94
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label35 
            Caption         =   "To"
            Height          =   210
            Left            =   3000
            TabIndex        =   93
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label34 
            Caption         =   "Show"
            Height          =   210
            Left            =   210
            TabIndex        =   92
            Top             =   720
            Width           =   750
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Deal"
         Height          =   1470
         Left            =   -74850
         TabIndex        =   75
         Top             =   3930
         Width           =   6000
         Begin VB.CommandButton cmdRtn 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   270
            Index           =   0
            Left            =   5400
            TabIndex        =   99
            Top             =   270
            Width           =   270
         End
         Begin VB.TextBox txtRtn 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   0
            Left            =   3900
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   240
            Width           =   1800
         End
         Begin VB.ComboBox cmbTra 
            Height          =   315
            Index           =   0
            ItemData        =   "frmAdd.frx":08BD
            Left            =   1050
            List            =   "frmAdd.frx":08C4
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   240
            Width           =   1800
         End
         Begin VB.TextBox txtDeal 
            Height          =   675
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   76
            Top             =   645
            Width           =   4650
         End
         Begin VB.Label Label28 
            Caption         =   "Return"
            Height          =   210
            Left            =   3000
            TabIndex        =   83
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label25 
            Caption         =   "Show"
            Height          =   210
            Left            =   210
            TabIndex        =   80
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label22 
            Caption         =   "Use"
            Height          =   210
            Left            =   210
            TabIndex        =   79
            Top             =   300
            Width           =   750
         End
      End
      Begin VB.Frame Frame18 
         Height          =   6570
         Left            =   -74850
         TabIndex        =   73
         Top             =   390
         Width           =   6000
         Begin VB.TextBox txtLost 
            Height          =   1200
            Left            =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   74
            Top             =   5250
            Width           =   5400
         End
         Begin VB.TextBox txtWin 
            Height          =   4500
            Left            =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   420
            Width           =   5400
         End
         Begin VB.Label Label39 
            Caption         =   "Winning"
            Height          =   215
            Left            =   300
            TabIndex        =   113
            Top             =   180
            Width           =   900
         End
         Begin VB.Label Label19 
            Caption         =   "Default Lost"
            Height          =   215
            Left            =   300
            TabIndex        =   112
            Top             =   4995
            Width           =   900
         End
      End
      Begin VB.TextBox txtAuthor 
         Height          =   315
         Left            =   -73800
         TabIndex        =   71
         Text            =   "Anonymous"
         Top             =   900
         Width           =   4650
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   -73800
         TabIndex        =   3
         Text            =   "Untitled"
         Top             =   480
         Width           =   4650
      End
      Begin VB.Frame Frame17 
         Caption         =   "Introduction"
         Height          =   5000
         Left            =   -74850
         TabIndex        =   68
         Top             =   1320
         Width           =   6000
         Begin VB.TextBox txtIntro 
            Height          =   4500
            Left            =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   69
            Top             =   240
            Width           =   5400
         End
      End
      Begin VB.PictureBox picOpp 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1350
         Left            =   4800
         Picture         =   "frmAdd.frx":08D0
         ScaleHeight     =   88
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   88
         TabIndex        =   67
         Top             =   900
         Width           =   1350
      End
      Begin VB.Frame Frame10 
         Caption         =   "Resolve"
         Height          =   1470
         Left            =   -74850
         TabIndex        =   59
         Top             =   2370
         Width           =   6000
         Begin VB.CheckBox chkType 
            Alignment       =   1  'Right Justify
            Caption         =   "Restart"
            Height          =   315
            Index           =   1
            Left            =   3000
            TabIndex        =   103
            Top             =   240
            Width           =   1125
         End
         Begin VB.TextBox txtRea 
            Height          =   315
            Left            =   1050
            TabIndex        =   61
            Top             =   240
            Width           =   1800
         End
         Begin VB.TextBox txtReaS 
            Height          =   675
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   60
            Top             =   645
            Width           =   4650
         End
         Begin VB.Label Label11 
            Caption         =   "Reaction"
            Height          =   210
            Left            =   210
            TabIndex        =   63
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Show"
            Height          =   210
            Left            =   210
            TabIndex        =   62
            Top             =   720
            Width           =   900
         End
      End
      Begin VB.TextBox txtRet 
         Height          =   315
         Left            =   -73800
         TabIndex        =   6
         Top             =   480
         Width           =   1800
      End
      Begin VB.Frame Frame16 
         Caption         =   "Description"
         Height          =   1470
         Left            =   -74850
         TabIndex        =   51
         Top             =   810
         Width           =   6000
         Begin VB.TextBox txtDesR 
            Height          =   1080
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   52
            Top             =   240
            Width           =   4650
         End
         Begin VB.Label Label17 
            Caption         =   "Show"
            Height          =   210
            Left            =   210
            TabIndex        =   53
            Top             =   360
            Width           =   750
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Attributes"
         Height          =   1470
         Left            =   -74850
         TabIndex        =   45
         Top             =   2370
         Width           =   6000
         Begin VB.ComboBox cmbSpe 
            Height          =   315
            Index           =   1
            ItemData        =   "frmAdd.frx":0F1C
            Left            =   1050
            List            =   "frmAdd.frx":0F32
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   240
            Width           =   1800
         End
         Begin VB.ComboBox cmbSpeR 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            ItemData        =   "frmAdd.frx":0F64
            Left            =   3900
            List            =   "frmAdd.frx":0F7A
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   240
            Width           =   1800
         End
         Begin VB.TextBox txtAttR 
            Height          =   675
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Top             =   645
            Width           =   4650
         End
         Begin VB.Label Label20 
            Caption         =   "Quality"
            Height          =   210
            Left            =   210
            TabIndex        =   64
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label16 
            Caption         =   "Show"
            Height          =   210
            Left            =   210
            TabIndex        =   50
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label15 
            Caption         =   "Improve"
            Height          =   210
            Left            =   3000
            TabIndex        =   49
            Top             =   300
            Width           =   900
         End
      End
      Begin VB.TextBox txtOpp 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   480
         Width           =   1800
      End
      Begin VB.Frame Frame13 
         Caption         =   "Description"
         Height          =   1470
         Left            =   150
         TabIndex        =   41
         Top             =   810
         Width           =   4500
         Begin VB.TextBox txtDesP 
            Height          =   1080
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Top             =   240
            Width           =   3105
         End
         Begin VB.Label Label13 
            Caption         =   "Show"
            Height          =   215
            Left            =   210
            TabIndex        =   43
            Top             =   300
            Width           =   855
         End
      End
      Begin VB.TextBox txtEle 
         Height          =   315
         Index           =   2
         Left            =   -73800
         TabIndex        =   2
         Top             =   480
         Width           =   1800
      End
      Begin VB.Frame Frame12 
         Caption         =   "Show Up"
         Height          =   1470
         Left            =   -74850
         TabIndex        =   37
         Top             =   5490
         Width           =   6000
         Begin VB.ComboBox cmbShw 
            Height          =   315
            Index           =   2
            ItemData        =   "frmAdd.frx":0FC1
            Left            =   1050
            List            =   "frmAdd.frx":0FC8
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   240
            Width           =   1800
         End
         Begin VB.TextBox txtShwT 
            Height          =   675
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Top             =   645
            Width           =   4650
         End
         Begin VB.Label Label38 
            Caption         =   "Show"
            Height          =   210
            Left            =   210
            TabIndex        =   96
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label37 
            Caption         =   "After"
            Height          =   210
            Left            =   210
            TabIndex        =   95
            Top             =   300
            Width           =   750
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Resolve"
         Height          =   1470
         Left            =   -74850
         TabIndex        =   36
         Top             =   810
         Width           =   6000
         Begin VB.TextBox txtActS 
            Height          =   1080
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   57
            Top             =   240
            Width           =   4650
         End
         Begin VB.Label Label9 
            Caption         =   "Show"
            Height          =   210
            Left            =   210
            TabIndex        =   58
            Top             =   300
            Width           =   735
         End
      End
      Begin VB.TextBox txtEle 
         Height          =   315
         Index           =   1
         Left            =   -73800
         TabIndex        =   1
         Top             =   480
         Width           =   1800
      End
      Begin VB.Frame Frame8 
         Caption         =   "Description"
         Height          =   1470
         Left            =   -74850
         TabIndex        =   32
         Top             =   810
         Width           =   6000
         Begin VB.TextBox txtDesC 
            Height          =   1080
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   240
            Width           =   4650
         End
         Begin VB.Label Label7 
            Caption         =   "Show"
            Height          =   215
            Left            =   210
            TabIndex        =   34
            Top             =   300
            Width           =   750
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Dialog"
         Height          =   1470
         Left            =   -74850
         TabIndex        =   29
         Top             =   2370
         Width           =   6000
         Begin VB.CheckBox chkDlg 
            Alignment       =   1  'Right Justify
            Caption         =   "Repeatedly"
            Height          =   315
            Index           =   0
            Left            =   3000
            TabIndex        =   104
            Top             =   240
            Width           =   1125
         End
         Begin VB.TextBox txtAsk 
            Height          =   315
            Left            =   1050
            TabIndex        =   55
            Top             =   240
            Width           =   1800
         End
         Begin VB.TextBox txtDlg 
            Height          =   675
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   645
            Width           =   4650
         End
         Begin VB.Label Label5 
            Caption         =   "Player"
            Height          =   210
            Left            =   210
            TabIndex        =   56
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label6 
            Caption         =   "Reply"
            Height          =   210
            Left            =   210
            TabIndex        =   31
            Top             =   720
            Width           =   750
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Treat"
         Height          =   1470
         Left            =   -74850
         TabIndex        =   25
         Top             =   3930
         Width           =   6000
         Begin VB.CommandButton cmdRtn 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   270
            Index           =   1
            Left            =   5400
            TabIndex        =   100
            Top             =   270
            Width           =   270
         End
         Begin VB.TextBox txtTre 
            Height          =   675
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   645
            Width           =   4650
         End
         Begin VB.ComboBox cmbTra 
            Height          =   315
            Index           =   1
            ItemData        =   "frmAdd.frx":0FD4
            Left            =   1050
            List            =   "frmAdd.frx":0FDB
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   240
            Width           =   1800
         End
         Begin VB.TextBox txtRtn 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   1
            Left            =   3900
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label Label31 
            Caption         =   "Show"
            Height          =   210
            Left            =   210
            TabIndex        =   86
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label30 
            Caption         =   "Return"
            Height          =   210
            Left            =   3000
            TabIndex        =   85
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label29 
            Caption         =   "Give"
            Height          =   210
            Left            =   210
            TabIndex        =   84
            Top             =   300
            Width           =   750
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Show Up"
         Height          =   1470
         Left            =   -74850
         TabIndex        =   22
         Top             =   5490
         Width           =   6000
         Begin VB.TextBox txtShwC 
            Height          =   675
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Top             =   645
            Width           =   4650
         End
         Begin VB.ComboBox cmbShw 
            Height          =   315
            Index           =   1
            ItemData        =   "frmAdd.frx":0FE7
            Left            =   1050
            List            =   "frmAdd.frx":0FEE
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label Label33 
            Caption         =   "Show"
            Height          =   210
            Left            =   210
            TabIndex        =   88
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label32 
            Caption         =   "After"
            Height          =   210
            Left            =   210
            TabIndex        =   87
            Top             =   300
            Width           =   750
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Show Up"
         Height          =   1470
         Left            =   -74850
         TabIndex        =   17
         Top             =   5490
         Width           =   6000
         Begin VB.ComboBox cmbShw 
            Height          =   315
            Index           =   0
            ItemData        =   "frmAdd.frx":0FFA
            Left            =   1050
            List            =   "frmAdd.frx":1001
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   240
            Width           =   1800
         End
         Begin VB.TextBox txtShw 
            Height          =   675
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   645
            Width           =   4650
         End
         Begin VB.Label Label27 
            Caption         =   "Show"
            Height          =   210
            Left            =   210
            TabIndex        =   82
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label26 
            Caption         =   "After"
            Height          =   210
            Left            =   210
            TabIndex        =   81
            Top             =   300
            Width           =   750
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Attributes"
         Height          =   1470
         Left            =   -74850
         TabIndex        =   12
         Top             =   2370
         Width           =   6000
         Begin VB.TextBox txtAtt 
            Height          =   675
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   645
            Width           =   4650
         End
         Begin VB.ComboBox cmbSpeR 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            ItemData        =   "frmAdd.frx":100D
            Left            =   3900
            List            =   "frmAdd.frx":1023
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   240
            Width           =   1800
         End
         Begin VB.ComboBox cmbSpe 
            Height          =   315
            Index           =   0
            ItemData        =   "frmAdd.frx":106A
            Left            =   1050
            List            =   "frmAdd.frx":1080
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label Label21 
            Caption         =   "Quality"
            Height          =   255
            Left            =   210
            TabIndex        =   66
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label4 
            Caption         =   "Improve"
            Height          =   210
            Left            =   3000
            TabIndex        =   19
            Top             =   300
            Width           =   900
         End
         Begin VB.Label Label3 
            Caption         =   "Show"
            Height          =   210
            Left            =   210
            TabIndex        =   13
            Top             =   720
            Width           =   750
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Description"
         Height          =   1470
         Left            =   -74850
         TabIndex        =   9
         Top             =   810
         Width           =   6000
         Begin VB.TextBox txtDes 
            Height          =   1080
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   240
            Width           =   4650
         End
         Begin VB.Label Label2 
            Caption         =   "Show"
            Height          =   215
            Left            =   210
            TabIndex        =   11
            Top             =   300
            Width           =   750
         End
      End
      Begin VB.TextBox txtEle 
         Height          =   315
         Index           =   0
         Left            =   -73800
         TabIndex        =   0
         Top             =   480
         Width           =   1800
      End
      Begin VB.Label Label45 
         Caption         =   "Encounter"
         Height          =   210
         Left            =   3150
         TabIndex        =   120
         Top             =   4140
         Width           =   750
      End
      Begin VB.Label Label44 
         Caption         =   "Position"
         Height          =   210
         Left            =   360
         TabIndex        =   118
         Top             =   4140
         Width           =   750
      End
      Begin VB.Label Label40 
         Caption         =   "Start in Room"
         Height          =   210
         Left            =   -72000
         TabIndex        =   115
         Top             =   6540
         Width           =   1200
      End
      Begin VB.Label Label41 
         Caption         =   "Skill"
         Height          =   210
         Left            =   3150
         TabIndex        =   102
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label24 
         Caption         =   "Author"
         Height          =   210
         Left            =   -74640
         TabIndex        =   72
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label23 
         Caption         =   "Title"
         Height          =   210
         Left            =   -74640
         TabIndex        =   70
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label18 
         Caption         =   "Name"
         Height          =   210
         Left            =   -74640
         TabIndex        =   54
         Top             =   540
         Width           =   900
      End
      Begin VB.Label Label14 
         Caption         =   "Name"
         Height          =   210
         Left            =   360
         TabIndex        =   44
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label12 
         Caption         =   "Name"
         Height          =   210
         Left            =   -74640
         TabIndex        =   40
         Top             =   540
         Width           =   900
      End
      Begin VB.Label Label8 
         Caption         =   "Name"
         Height          =   210
         Left            =   -74640
         TabIndex        =   35
         Top             =   540
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   210
         Left            =   -74640
         TabIndex        =   8
         Top             =   540
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iTab As Integer
Private strElm As String

Private Sub Form_Load()
  frmMaker.Enabled = False
End Sub

Public Sub LoadTab(iNum As Integer, strName As String)
  Dim varObj As Variant
  If iNum < 6 Then iTab = iNum: strElm = strName
  If Left(strName, 3) = "..." Then strName = Mid(strName, 4)
  
  For i% = 0 To 6
    If i% = iNum Then SSTab.TabVisible(i%) = True: SSTab.Tab = i% Else SSTab.TabVisible(i%) = False
  Next i%
  
  varTwo = Split(ReadINI("Room", "Dir", strFile), "|")
  If UBound(varTwo) > -1 Then
    For i% = 0 To UBound(varTwo) - 1
      varTmp = Split(ReadINI("Element", varTwo(i%), strFile), "|")
      If UBound(varTmp) = -1 Then varTmp = Split("|||", "|")
      For e% = 0 To 3 Step 3
        varObj = Split(varTmp(e%), "")
        For a% = 1 To UBound(varObj)
          If iNum < 2 Then cmbTra(iNum).AddItem varObj(a%)
        Next a%
      Next e%
    Next i%
  End If
  If iNum < 2 Then cmbTra(iNum).ListIndex = 0
  
  For i% = 0 To 3 Step 3
    For e% = 0 To frmMaker.lstTmp(2 + i%).ListCount - 1
      If iNum < 3 Then cmbShw(iNum).AddItem frmMaker.lstTmp(2 + i%).List(e%)
    Next e%
  Next i%
  If iNum < 3 Then cmbShw(iNum).ListIndex = 0
  
  For i% = 0 To 3 Step 3
    For e% = 0 To frmMaker.lstTmp(1 + i%).ListCount - 1
      cmbMov.AddItem frmMaker.lstTmp(1 + i%).List(e%)
    Next e%
  Next i%
  cmbMov.ListIndex = 0
  
  varTmp = ReadINI("Room", "Title", strFile)
  If varTmp <> "" Then
    varTmp = Split(varTmp, "|")
    For i% = 0 To UBound(varTmp)
      If varTmp(i%) <> "" Then cmbRoom.AddItem varTmp(i%): cmbStart.AddItem varTmp(i%): cmbRoomP.AddItem varTmp(i%)
    Next i%
  End If
  
  cmbSpe(0).ListIndex = 0: cmbSpeR(0).ListIndex = 0
  cmbSpe(1).ListIndex = 0: cmbSpeR(1).ListIndex = 0
  
  Select Case iNum
    Case 0
      txtEle(0) = strName
      FindAndRemove cmbTra(0), strName
      txtDes = ReadINI("Description", strName, strFile)
      If ReadINI("Element", strName, strFile) <> "" Then chkSta(0).Value = 1
      varTmp = Split(ReadINI("Attributes", strName, strFile), "")
      If UBound(varTmp) = 1 Then cmbSpe(0).ListIndex = varTmp(0): cmbSpeR(0).ListIndex = varTmp(1)
      txtAtt = ReadINI("Attributes", "S|" & strName, strFile)
      For i% = 0 To cmbTra(0).ListCount - 1
        If ReadINI("Trade", strName, strFile) = cmbTra(0).List(i%) Then cmbTra(0).ListIndex = i%: txtDeal = ReadINI("Trade", "S|" & strName, strFile): txtRtn(0) = ReadINI("Trade", cmbTra(0).List(i%), strFile)
      Next i%
      For i% = 0 To cmbShw(0).ListCount - 1
        If strName = "" Then varTmp = 0 Else varTmp = InStr(ReadINI("ShowUp", cmbShw(0).List(i%), strFile), "" & strName)
        If varTmp > 0 Then cmbShw(0).ListIndex = i%: txtShw = ReadINI("ShowUp", "S|" & strName, strFile)
      Next i%
      
    Case 1
      txtEle(1) = strName
      txtDesC = ReadINI("Description", strName, strFile)
      If ReadINI("Resolve", strName, strFile) <> "" Then chkEne.Value = 1
      If ReadINI("Dialog", "|" & strName, strFile) <> "" Then chkDlg(0).Value = 1
      txtAsk = ReadINI("Dialog", strName, strFile): txtDlg = ReadINI("Dialog", "S|" & strName, strFile)
      For i% = 0 To cmbTra(1).ListCount - 1
        If ReadINI("Trade", strName, strFile) = cmbTra(1).List(i%) Then cmbTra(1).ListIndex = i%: txtTre = ReadINI("Trade", "S|" & strName, strFile): txtRtn(1) = ReadINI("Trade", cmbTra(1).List(i%), strFile)
      Next i%
      For i% = 0 To cmbShw(1).ListCount - 1
        If strName = "" Then varTmp = 0 Else varTmp = InStr(ReadINI("ShowUp", cmbShw(1).List(i%), strFile), "" & strName)
        If varTmp > 0 Then cmbShw(1).ListIndex = i%: txtShwC = ReadINI("ShowUp", "S|" & strName, strFile)
      Next i%
    
    Case 2
      txtEle(2) = strName
      FindAndRemove cmbShw(2), strName
      txtRea = ReadINI("Resolve", strName, strFile)
      varTmp = ReadINI("Starting", "Buffer", strFile)
      If InStr(varTmp, "" & strName & "") > 0 Then chkType(0).Value = 1
      If ReadINI("Element", strName, strFile) <> "" Then chkType(1).Value = 1
      If txtRea <> "" Then If ReadINI("Element", txtRea, strFile) <> "" Then chkType(1).Value = 1 Else chkType(1).Value = 0
      txtActS = ReadINI("Resolve", "S|" & strName, strFile)
      txtReaS = ReadINI("Resolve", "S|" & txtRea, strFile)
      For i% = 0 To cmbMov.ListCount - 1
        varTmp = InStr(ReadINI("Movement", strName, strFile), cmbMov.List(i%) & "")
        If varTmp > 0 Then cmbMov.ListIndex = i%: txtMov = ReadINI("Movement", "S|" & strName, strFile)
      Next i%
      varTmp = Split(ReadINI("Movement", strName, strFile), "")
      varTwo = Split(ReadINI("Room", "Dir", strFile), "|")
      If UBound(varTmp) > -1 Then
        For i% = 0 To UBound(varTwo)
          If varTmp(1) = varTwo(i%) Then cmbRoom.ListIndex = i%
        Next i%
      End If
      For i% = 0 To cmbShw(2).ListCount - 1
        If strName = "" Then varTmp = 0 Else varTmp = InStr(ReadINI("ShowUp", cmbShw(2).List(i%), strFile), "" & strName)
        If varTmp > 0 Then cmbShw(2).ListIndex = i%: txtShwT = ReadINI("ShowUp", "S|" & strName, strFile)
      Next i%
      
    Case 3
      varTmp = Split(ReadINI("Starting", "Info", strFile), "|")
      If UBound(varTmp) > -1 Then
        txtTitle = varTmp(1): txtAuthor = varTmp(2): txtIntro = varTmp(4)
        For i% = 5 To UBound(varTmp)
          txtIntro = txtIntro & vbCrLf & varTmp(i%)
        Next i%
      End If
      varTmp = Split(ReadINI("Room", "Dir", strFile), "|")
      If UBound(varTmp) > -1 Then
        For i% = 0 To UBound(varTmp)
          If ReadINI("Starting", "Room", strFile) = varTmp(i%) Then cmbStart.ListIndex = i%
        Next i%
      End If

    Case 4
      varTmp = Split(ReadINI("Opponent", "Data", strFile), "")
      If UBound(varTmp) = 1 Then txtOpp = varTmp(0): txtSkill = varTmp(1)
      txtDesP = ReadINI("Description", txtOpp, strFile)
      If ReadINI("Dialog", "|" & txtOpp, strFile) <> "" Then chkDlg(1).Value = 1
      txtAskP = ReadINI("Dialog", txtOpp, strFile): txtDlgP = ReadINI("Dialog", "S|" & txtOpp, strFile)
      txtEnc = ReadINI("Ending", "Encounter", strFile)
      varTwo = Split(ReadINI("Room", "Dir", strFile), "|")
      If UBound(varTmp) > -1 Then
        For i% = 0 To UBound(varTwo)
          If ReadINI("Movement", CStr(varTwo(i%)), strFile) = "" & txtOpp Then cmbRoomP.ListIndex = i%
        Next i%
      End If

    Case 5
      txtWin.Tag = ReadINI("Ending", "Win", strFile): txtWin = Replace(txtWin.Tag, "|", vbCrLf)
      txtLost.Tag = ReadINI("Ending", "Lost", strFile): txtLost = Replace(txtLost.Tag, "|", vbCrLf)
  
    Case 6
      txtRet = strName
      cmdAdd.Caption = "Update"
      txtDesR = ReadINI("Description", txtRet, strFile)
      varTmp = Split(ReadINI("Attributes", txtRet, strFile), "")
      If UBound(varTmp) = 1 Then cmbSpe(1).ListIndex = varTmp(0): cmbSpeR(1).ListIndex = varTmp(1)
      txtAttR = ReadINI("Attributes", "S|" & txtRet, strFile)
  End Select
  frmAdd.Show
End Sub

Private Sub SaveTab(iNum As Integer, strName As String)
  Select Case iNum
    Case 0
      If ValName(strName) = True Then Exit Sub
      WriteINI "Description", strName, txtDes, strFile
      If chkSta(0).Value = 0 Then WriteINI "Element", strName, "", strFile Else WriteINI "Element", strName, strName, strFile
      If cmbSpe(0).ListIndex = 0 Then Me.Tag = "" Else Me.Tag = cmbSpe(0).ListIndex & "" & cmbSpeR(0).ListIndex
      WriteINI "Attributes", strName, Me.Tag, strFile: WriteINI "Attributes", "S|" & strName, txtAtt, strFile
      For i% = 1 To cmbTra(0).ListCount - 1
        If ReadINI("Trade", strName, strFile) <> "" Then WriteINI "Trade", strName & "|" & cmbTra(0).List(i%), "", strFile
      Next i%
      If cmbTra(0).ListIndex > 0 Then WriteINI "Trade", strName, cmbTra(0).List(cmbTra(0).ListIndex), strFile: WriteINI "Trade", "S|" & strName, txtDeal, strFile: WriteINI "Trade", cmbTra(0).List(cmbTra(0).ListIndex), txtRtn(0), strFile
      varTmp = Split(ReadINI("ShowUp", cmbShw(0).Text, strFile), "|")
      If UBound(varTmp) = -1 Then varTmp = Split("||", "|")
      If InStr(varTmp(0), "" & strName) = 0 Then varTmp(0) = varTmp(0) & "" & strName
      If cmbShw(0).ListIndex = 0 Then varTmp(0) = Replace(varTmp(0), "" & strName, "")
      WriteINI "ShowUp", cmbShw(0).List(cmbShw(0).ListIndex), CStr(varTmp(0) & "|" & varTmp(1) & "|" & varTmp(2)), strFile
      WriteINI "ShowUp", "S|" & strName, txtShw, strFile
    
    Case 1
      If ValName(strName) = True Then Exit Sub
      WriteINI "Description", strName, txtDesC, strFile
      If chkEne.Value = 0 Then WriteINI "Resolve", strName, "", strFile: WriteINI "Resolve", "S|" & strName, "", strFile Else WriteINI "Resolve", strName, "Attack to " & strName, strFile: WriteINI "Resolve", "S|" & strName, "You defeated your opponent.", strFile
      If chkDlg(0).Value = 0 Then WriteINI "Dialog", "|" & strName, "", strFile Else WriteINI "Dialog", "|" & strName, strName, strFile
      WriteINI "Dialog", strName, txtAsk, strFile: WriteINI "Dialog", "S|" & strName, txtDlg, strFile
      For i% = 1 To cmbTra(1).ListCount - 1
        If ReadINI("Trade", strName, strFile) <> "" Then WriteINI "Trade", strName & "|" & cmbTra(1).List(i%), "", strFile
      Next i%
      If cmbTra(1).ListIndex > 0 Then WriteINI "Trade", strName, cmbTra(1).List(cmbTra(1).ListIndex), strFile: WriteINI "Trade", "S|" & strName, txtTre, strFile: WriteINI "Trade", cmbTra(1).List(cmbTra(1).ListIndex), txtRtn(1), strFile
      varTmp = Split(ReadINI("ShowUp", cmbShw(1).Text, strFile), "|")
      If UBound(varTmp) = -1 Then varTmp = Split("||", "|")
      If InStr(varTmp(1), "" & strName) = 0 Then varTmp(1) = varTmp(1) & "" & strName
      If cmbShw(1).ListIndex = 0 Then varTmp(1) = Replace(varTmp(1), "" & strName, "")
      WriteINI "ShowUp", cmbShw(1).List(cmbShw(1).ListIndex), CStr(varTmp(0) & "|" & varTmp(1) & "|" & varTmp(2)), strFile
      WriteINI "ShowUp", "S|" & strName, txtShwC, strFile
  
    Case 2
      If ValName(strName) = True Then Exit Sub
      varTmp = ReadINI("Starting", "Buffer", strFile)
      If InStr(varTmp, "" & strName & "") > 0 Then varTmp = Replace(varTmp, "" & strName & "", "")
      If InStr(varTmp, "" & txtRea & "") > 0 Then varTmp = Replace(varTmp, "" & txtRea & "", "")
      If chkType(0).Value = 1 Then WriteINI "Starting", "Buffer", CStr(varTmp) & "" & strName & "" & txtRea & "", strFile Else WriteINI "Starting", "Buffer", "", strFile
      WriteINI "Element", strName, txtRea, strFile
      If chkType(1).Value = 0 Then WriteINI "Element", txtRea, "", strFile
      If chkType(1).Value = 1 Then If txtRea = "" Then WriteINI "Element", strName, strName, strFile Else WriteINI "Element", txtRea, strName, strFile
      WriteINI "Resolve", strName, txtRea, strFile
      WriteINI "Resolve", "S|" & strName, txtActS, strFile: WriteINI "Resolve", "S|" & txtRea, txtReaS, strFile
      varTwo = Split(ReadINI("Room", "Dir", strFile), "|")
      If UBound(varTwo) > -1 Then If cmbRoom.ListIndex > -1 Then Me.Tag = varTwo(cmbRoom.ListIndex)
      If cmbMov.ListIndex > 0 Then WriteINI "Movement", strName, cmbMov.Text & "" & Me.Tag, strFile
      WriteINI "Movement", "S|" & strName, txtMov, strFile
      varTmp = Split(ReadINI("ShowUp", cmbShw(2).Text, strFile), "|")
      If UBound(varTmp) = -1 Then varTmp = Split("||", "|")
      If InStr(varTmp(2), "" & strName) = 0 Then varTmp(2) = varTmp(2) & "" & strName
      If cmbShw(2).ListIndex = 0 Then varTmp(2) = Replace(varTmp(2), "" & strName, "")
      WriteINI "ShowUp", cmbShw(2).List(cmbShw(2).ListIndex), CStr(varTmp(0) & "|" & varTmp(1) & "|" & varTmp(2)), strFile
      WriteINI "ShowUp", "S|" & strName, txtShwT, strFile

    Case 3
      If txtTitle = "" Then txtTitle = "Untitled"
      If txtAuthor = "" Then txtAuthor = "Anonymous"
      txtIntro.Tag = Replace(txtIntro, vbCrLf, "|")
      WriteINI "Starting", "Info", "|" & txtTitle & "|" & txtAuthor & "||" & txtIntro.Tag, strFile
      varTwo = Split(ReadINI("Room", "Dir", strFile), "|")
      If cmbStart.ListIndex > -1 Then WriteINI "Starting", "Room", CStr(varTwo(cmbStart.ListIndex)), strFile

    Case 4
      If ValName(txtOpp) = True Then Exit Sub
      WriteINI "Description", txtOpp, txtDesP, strFile: WriteINI "Opponent", "Data", txtOpp & "" & txtSkill, strFile
      If chkDlg(1).Value = 0 Then WriteINI "Dialog", "|" & txtOpp, "", strFile Else WriteINI "Dialog", "|" & txtOpp, txtOpp, strFile
      WriteINI "Dialog", txtOpp, txtAskP, strFile: WriteINI "Dialog", "S|" & txtOpp, txtDlgP, strFile
      varTmp = Split(ReadINI("Room", "Dir", strFile), "|")
      If cmbRoomP.ListIndex > -1 Then WriteINI "Movement", varTmp(cmbRoomP.ListIndex), "" & txtOpp, strFile
      WriteINI "Ending", "Encounter", txtEnc, strFile
      WriteINI "Resolve", txtOpp, txtEnc, strFile
      
    Case 5
      Me.Tag = Replace(txtWin, vbCrLf, "|"): WriteINI "Ending", "Win", Me.Tag, strFile
      Me.Tag = Replace(txtLost, vbCrLf, "|"): WriteINI "Ending", "Lost", Me.Tag, strFile
    
    Case 6
      cmdAdd.Caption = "Add"
      If ValName(txtRet) = True Then Exit Sub
      If txtRet <> "" Then txtRtn(iTab) = txtRet
      WriteINI "Description", txtRet, txtDesR, strFile
      SSTab.TabVisible(6) = False: SSTab.TabVisible(iTab) = True: SSTab.Tab = iTab
      If cmbSpe(1).ListIndex = 0 Then Me.Tag = "" Else Me.Tag = cmbSpe(1).ListIndex & "" & cmbSpeR(1).ListIndex
      WriteINI "Attributes", txtRet, Me.Tag, strFile: WriteINI "Attributes", "S|" & txtRet, txtAttR, strFile
  End Select
  If iNum < 3 Then If cmbShw(iNum).ListIndex = 0 Then strElm = txtEle(iNum) Else strElm = "..." & txtEle(iNum)
  If iNum < 6 Then Unload Me
End Sub

Private Function ValName(strName As String) As Boolean
  If strName = "" Then ValName = True
  
  For i% = 0 To 15
    varTmp = ReadINI("Element", CStr(i%), strFile)
    If InStr(varTmp, "" & strName) > 0 Then ValName = True
  Next i%
  
  If strElm <> "" Then ValName = False
  If ValName = True Then MsgBox "Invalid Name", vbInformation
End Function

Private Sub FindAndRemove(cmbBox As ComboBox, strFind As String)
  For i% = 0 To cmbBox.ListCount - 1
    If cmbBox.List(i%) = strFind Then cmbBox.RemoveItem i%
  Next i%
End Sub

Private Sub cmdAdd_Click()
  If iTab < 3 Then Me.Tag = txtEle(iTab)
  If cmdAdd.Caption = "Add" Then SaveTab iTab, Me.Tag Else SaveTab 6, ""
End Sub

Private Sub cmbSpe_Click(Index As Integer)
  If cmbSpe(Index).ListIndex = 0 Then cmbSpeR(Index).Enabled = False Else cmbSpeR(Index).Enabled = True
End Sub

Private Sub cmbTra_Click(Index As Integer)
  If cmbTra(Index).ListIndex = 0 Then cmdRtn(Index).Enabled = False Else cmdRtn(Index).Enabled = True
End Sub

Private Sub cmbMov_Click()
  If cmbMov.ListIndex = 0 Then cmbRoom.Enabled = False Else cmbRoom.Enabled = True
End Sub

Private Sub cmdRtn_Click(Index As Integer)
  LoadTab 6, txtRtn(Index)
End Sub

Private Sub cmdCnl_Click()
  If cmdAdd.Caption = "Add" Then Unload Me Else cmdAdd.Caption = "Add": SSTab.TabVisible(6) = False: SSTab.TabVisible(iTab) = True: SSTab.Tab = iTab
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If strElm <> "" Then frmMaker.AddElement strElm
  frmMaker.Enabled = True
End Sub
