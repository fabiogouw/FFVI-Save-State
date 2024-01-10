VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FFVI - Editor de SaveState (ZSNES)"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "frmPrincipal.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraComando"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraGeral"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraEstatisticas"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraEquip"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraStatus"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboHeroi"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraHpMp"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cdlArquivo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Magias"
      TabPicture(1)   =   "frmPrincipal.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraEspers"
      Tab(1).Control(1)=   "fraMagias"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Especiais"
      TabPicture(2)   =   "frmPrincipal.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraLore"
      Tab(2).Control(1)=   "fraRage"
      Tab(2).Control(2)=   "fraSwdtech"
      Tab(2).Control(3)=   "fraBlitz"
      Tab(2).Control(4)=   "fraDance"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "GP/Itens"
      TabPicture(3)   =   "frmPrincipal.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraItem"
      Tab(3).Control(1)=   "fraGPSteps"
      Tab(3).ControlCount=   2
      Begin MSComDlg.CommonDialog cdlArquivo 
         Left            =   6600
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Frame fraHpMp 
         Caption         =   "HP && MP"
         Height          =   1575
         Left            =   2760
         TabIndex        =   126
         Top             =   1080
         Width           =   2055
         Begin VB.TextBox txtCampo 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   3
            Left            =   240
            MaxLength       =   4
            TabIndex        =   147
            Top             =   480
            Width           =   465
         End
         Begin VB.TextBox txtCampo 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   4
            Left            =   1080
            MaxLength       =   4
            TabIndex        =   129
            Top             =   480
            Width           =   465
         End
         Begin VB.TextBox txtCampo 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   5
            Left            =   240
            MaxLength       =   3
            TabIndex        =   128
            Top             =   1080
            Width           =   465
         End
         Begin VB.TextBox txtCampo 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   6
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   127
            Top             =   1080
            Width           =   465
         End
         Begin MSComCtl2.UpDown updCampo 
            Height          =   315
            Index           =   2
            Left            =   720
            TabIndex        =   139
            Top             =   480
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            BuddyControl    =   "txtCampo(3)"
            BuddyDispid     =   196610
            BuddyIndex      =   3
            OrigLeft        =   960
            OrigTop         =   480
            OrigRight       =   1155
            OrigBottom      =   795
            Max             =   9999
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown updCampo 
            Height          =   315
            Index           =   3
            Left            =   1560
            TabIndex        =   140
            Top             =   480
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtCampo(4)"
            BuddyDispid     =   196610
            BuddyIndex      =   4
            OrigLeft        =   1560
            OrigTop         =   480
            OrigRight       =   1755
            OrigBottom      =   795
            Max             =   9999
            Min             =   1
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown updCampo 
            Height          =   315
            Index           =   4
            Left            =   720
            TabIndex        =   141
            Top             =   1080
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            BuddyControl    =   "txtCampo(5)"
            BuddyDispid     =   196610
            BuddyIndex      =   5
            OrigLeft        =   720
            OrigTop         =   1080
            OrigRight       =   915
            OrigBottom      =   1395
            Max             =   999
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown updCampo 
            Height          =   315
            Index           =   5
            Left            =   1560
            TabIndex        =   142
            Top             =   1080
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtCampo(6)"
            BuddyDispid     =   196610
            BuddyIndex      =   6
            OrigLeft        =   1560
            OrigTop         =   1080
            OrigRight       =   1755
            OrigBottom      =   1395
            Max             =   999
            Min             =   1
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "HP"
            Height          =   255
            Left            =   240
            TabIndex        =   133
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "HP máximo"
            Height          =   255
            Left            =   1080
            TabIndex        =   132
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "MP"
            Height          =   255
            Left            =   240
            TabIndex        =   131
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label5 
            Caption         =   "MP máximo"
            Height          =   255
            Left            =   1080
            TabIndex        =   130
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame fraGPSteps 
         Caption         =   "GP && Steps"
         Height          =   975
         Left            =   -74280
         TabIndex        =   115
         Top             =   600
         Width           =   5655
         Begin VB.TextBox txtCampo 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   12
            Left            =   2880
            TabIndex        =   135
            Text            =   "AAAAAAAA"
            Top             =   480
            Width           =   780
         End
         Begin VB.TextBox txtCampo 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   11
            Left            =   840
            TabIndex        =   134
            Text            =   "AAAAAAAA"
            Top             =   480
            Width           =   780
         End
         Begin MSComCtl2.UpDown updCampo 
            Height          =   315
            Index           =   10
            Left            =   1621
            TabIndex        =   148
            Top             =   480
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            BuddyControl    =   "txtCampo(11)"
            BuddyDispid     =   196610
            BuddyIndex      =   11
            OrigLeft        =   1800
            OrigTop         =   480
            OrigRight       =   1995
            OrigBottom      =   795
            Max             =   255
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown updCampo 
            Height          =   315
            Index           =   11
            Left            =   3665
            TabIndex        =   149
            Top             =   480
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCampo(12)"
            BuddyDispid     =   196610
            BuddyIndex      =   12
            OrigLeft        =   3600
            OrigTop         =   480
            OrigRight       =   3795
            OrigBottom      =   795
            Max             =   162
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label31 
            Caption         =   "Steps"
            Height          =   255
            Left            =   3840
            TabIndex        =   117
            Top             =   225
            Width           =   495
         End
         Begin VB.Label Label29 
            Caption         =   "GP"
            Height          =   255
            Left            =   1440
            TabIndex        =   116
            Top             =   225
            Width           =   375
         End
      End
      Begin VB.Frame fraItem 
         Caption         =   "Item"
         Height          =   3375
         Left            =   -74280
         TabIndex        =   107
         Top             =   1680
         Width           =   5655
         Begin MSComctlLib.ProgressBar prbItem 
            Height          =   255
            Left            =   1800
            TabIndex        =   160
            Top             =   2520
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Min             =   1e-4
            Scrolling       =   1
         End
         Begin MSFlexGridLib.MSFlexGrid grdItem 
            Height          =   2175
            Left            =   1800
            TabIndex        =   159
            Top             =   360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   3836
            _Version        =   393216
            FixedCols       =   0
            AllowBigSelection=   0   'False
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.TextBox txtCampo 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   14
            Left            =   4320
            MaxLength       =   3
            TabIndex        =   136
            Text            =   "AAAAAAAA"
            Top             =   600
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.ComboBox cboItem 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   111
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmdOrderItem 
            Caption         =   "Reorganizar"
            Height          =   375
            Left            =   240
            TabIndex        =   110
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CommandButton cmdDelItem 
            Caption         =   "Remover"
            Height          =   375
            Left            =   240
            TabIndex        =   109
            Top             =   1440
            Width           =   1095
         End
         Begin VB.CommandButton cmdAddItem 
            Caption         =   "Adicionar"
            Height          =   375
            Left            =   240
            TabIndex        =   108
            Top             =   960
            Width           =   1095
         End
         Begin MSComCtl2.UpDown updCampo 
            Height          =   315
            Index           =   13
            Left            =   5160
            TabIndex        =   158
            Top             =   600
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            BuddyControl    =   "txtCampo(14)"
            BuddyDispid     =   196610
            BuddyIndex      =   14
            OrigLeft        =   1080
            OrigTop         =   1080
            OrigRight       =   1275
            OrigBottom      =   1395
            Max             =   99
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label lblRegistro 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2040
            TabIndex        =   114
            Top             =   3000
            Width           =   1815
         End
         Begin VB.Label Label28 
            Caption         =   "Quantidade"
            Height          =   255
            Left            =   4320
            TabIndex        =   113
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label27 
            Caption         =   "Item"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame fraMagias 
         Caption         =   "Magias"
         Height          =   1935
         Left            =   -74280
         TabIndex        =   104
         Top             =   600
         Width           =   5775
         Begin VB.ComboBox cboMagias 
            Height          =   315
            ItemData        =   "frmPrincipal.frx":0070
            Left            =   240
            List            =   "frmPrincipal.frx":0072
            Style           =   2  'Dropdown List
            TabIndex        =   151
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Frame fraMagiaOpt 
            Caption         =   "Aprendizado"
            Height          =   1455
            Left            =   1920
            TabIndex        =   150
            Top             =   240
            Width           =   3615
            Begin VB.OptionButton optAprendizado 
               Caption         =   "Magia em aprendizagem (%)"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   156
               Top             =   960
               Width           =   2415
            End
            Begin VB.OptionButton optAprendizado 
               Caption         =   "Magia não aprendida (0%)"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   155
               Top             =   240
               Width           =   2415
            End
            Begin VB.OptionButton optAprendizado 
               Caption         =   "Magia aprendida (100%)"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   154
               Top             =   600
               Width           =   2055
            End
            Begin VB.TextBox txtCampo 
               Enabled         =   0   'False
               Height          =   315
               Index           =   13
               Left            =   2760
               TabIndex        =   152
               Top             =   960
               Width           =   500
            End
            Begin MSComCtl2.UpDown updCampo 
               Height          =   315
               Index           =   12
               Left            =   3240
               TabIndex        =   153
               Top             =   960
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   556
               _Version        =   393216
               Value           =   1
               BuddyControl    =   "txtCampo(13)"
               BuddyDispid     =   196610
               BuddyIndex      =   13
               OrigLeft        =   3240
               OrigTop         =   720
               OrigRight       =   3435
               OrigBottom      =   1035
               Max             =   99
               Min             =   1
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   0
               Enabled         =   0   'False
            End
         End
         Begin VB.ComboBox cboHeroiMagia 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   105
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label25 
            Caption         =   "Magia"
            Height          =   255
            Left            =   240
            TabIndex        =   157
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label24 
            Caption         =   "Personagem"
            Height          =   255
            Left            =   240
            TabIndex        =   106
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame fraEspers 
         Caption         =   "Espers"
         Height          =   2535
         Left            =   -74265
         TabIndex        =   76
         Top             =   2640
         Width           =   5775
         Begin VB.CheckBox chkEspers 
            Caption         =   "Phoenix"
            Height          =   255
            Index           =   26
            Left            =   4155
            TabIndex        =   103
            Top             =   2160
            Width           =   975
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Starlet"
            Height          =   255
            Index           =   25
            Left            =   4155
            TabIndex        =   102
            Top             =   1920
            Width           =   855
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Fenrir"
            Height          =   255
            Index           =   24
            Left            =   4155
            TabIndex        =   101
            Top             =   1680
            Width           =   735
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Unicorn"
            Height          =   255
            Index           =   23
            Left            =   4155
            TabIndex        =   100
            Top             =   1440
            Width           =   975
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Golem"
            Height          =   255
            Index           =   22
            Left            =   4155
            TabIndex        =   99
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Sraphim"
            Height          =   255
            Index           =   21
            Left            =   4155
            TabIndex        =   98
            Top             =   960
            Width           =   975
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Phantom"
            Height          =   255
            Index           =   20
            Left            =   4155
            TabIndex        =   97
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Carbunkl"
            Height          =   255
            Index           =   19
            Left            =   4155
            TabIndex        =   96
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "ZoneSeek"
            Height          =   255
            Index           =   18
            Left            =   4155
            TabIndex        =   95
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Kirin"
            Height          =   255
            Index           =   17
            Left            =   2355
            TabIndex        =   94
            Top             =   2160
            Width           =   615
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Ragnarok"
            Height          =   255
            Index           =   16
            Left            =   2355
            TabIndex        =   93
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Crusader"
            Height          =   255
            Index           =   15
            Left            =   2355
            TabIndex        =   92
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Alexandr"
            Height          =   255
            Index           =   14
            Left            =   2355
            TabIndex        =   91
            Top             =   1440
            Width           =   975
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Bahamut"
            Height          =   255
            Index           =   13
            Left            =   2355
            TabIndex        =   90
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Raiden"
            Height          =   255
            Index           =   12
            Left            =   2355
            TabIndex        =   89
            Top             =   960
            Width           =   855
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Odin"
            Height          =   255
            Index           =   11
            Left            =   2355
            TabIndex        =   88
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Tritoch"
            Height          =   255
            Index           =   10
            Left            =   2355
            TabIndex        =   87
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Palidor"
            Height          =   255
            Index           =   9
            Left            =   2355
            TabIndex        =   86
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Stray"
            Height          =   255
            Index           =   8
            Left            =   555
            TabIndex        =   85
            Top             =   2160
            Width           =   735
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Bismark"
            Height          =   255
            Index           =   7
            Left            =   555
            TabIndex        =   84
            Top             =   1920
            Width           =   855
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Shoat"
            Height          =   255
            Index           =   6
            Left            =   555
            TabIndex        =   83
            Top             =   1680
            Width           =   735
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Maduin"
            Height          =   255
            Index           =   5
            Left            =   555
            TabIndex        =   82
            Top             =   1440
            Width           =   855
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Terrato"
            Height          =   255
            Index           =   4
            Left            =   555
            TabIndex        =   81
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Siren"
            Height          =   255
            Index           =   3
            Left            =   555
            TabIndex        =   80
            Top             =   960
            Width           =   735
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Shiva"
            Height          =   255
            Index           =   2
            Left            =   555
            TabIndex        =   79
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Ifrit"
            Height          =   255
            Index           =   1
            Left            =   555
            TabIndex        =   78
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox chkEspers 
            Caption         =   "Ramuh"
            Height          =   255
            Index           =   0
            Left            =   555
            TabIndex        =   77
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame fraLore 
         Caption         =   "Lore"
         Height          =   2055
         Left            =   -71280
         TabIndex        =   74
         Top             =   600
         Width           =   2775
         Begin VB.ListBox lstLore 
            Height          =   1635
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   75
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fraRage 
         Caption         =   "Rage"
         Height          =   2055
         Left            =   -74400
         TabIndex        =   72
         Top             =   600
         Width           =   2775
         Begin VB.ListBox lstRage 
            Height          =   1635
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   73
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fraSwdtech 
         Caption         =   "SwdTech"
         Height          =   2535
         Left            =   -70200
         TabIndex        =   63
         Top             =   2760
         Width           =   1695
         Begin VB.CheckBox chkSwdtech 
            Caption         =   "Cleave"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   71
            Top             =   2040
            Width           =   1335
         End
         Begin VB.CheckBox chkSwdtech 
            Caption         =   "QuadraSlice"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   70
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CheckBox chkSwdtech 
            Caption         =   "Stunner"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   69
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CheckBox chkSwdtech 
            Caption         =   "Empowerer"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   68
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CheckBox chkSwdtech 
            Caption         =   "Quadra Slam"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   67
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox chkSwdtech 
            Caption         =   "Slash"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   66
            Top             =   840
            Width           =   1335
         End
         Begin VB.CheckBox chkSwdtech 
            Caption         =   "Retort"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   65
            Top             =   600
            Width           =   1335
         End
         Begin VB.CheckBox chkSwdtech 
            Caption         =   "Dispatch"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   64
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraBlitz 
         Caption         =   "Blitz"
         Height          =   2535
         Left            =   -72240
         TabIndex        =   54
         Top             =   2760
         Width           =   1695
         Begin VB.CheckBox chkBlitz 
            Caption         =   "Bum Rush"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   62
            Top             =   2040
            Width           =   1215
         End
         Begin VB.CheckBox chkBlitz 
            Caption         =   "Spiraler"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   61
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CheckBox chkBlitz 
            Caption         =   "Air Blade"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   60
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CheckBox chkBlitz 
            Caption         =   "Mantra"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   59
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox chkBlitz 
            Caption         =   "Fire Dance"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   58
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox chkBlitz 
            Caption         =   "Suplex"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   57
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkBlitz 
            Caption         =   "Aura Bolt"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   56
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chkBlitz 
            Caption         =   "Pummel"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   55
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame fraDance 
         Caption         =   "Dance"
         Height          =   2535
         Left            =   -74400
         TabIndex        =   45
         Top             =   2760
         Width           =   1815
         Begin VB.CheckBox chkDance 
            Caption         =   "Snowman Jazz"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   53
            Top             =   2040
            Width           =   1455
         End
         Begin VB.CheckBox chkDance 
            Caption         =   "Dusk Requium"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   52
            Top             =   1800
            Width           =   1455
         End
         Begin VB.CheckBox chkDance 
            Caption         =   "Water Rondo"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   51
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CheckBox chkDance 
            Caption         =   "Earth Blues"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   50
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CheckBox chkDance 
            Caption         =   "Love Sonata"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   49
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CheckBox chkDance 
            Caption         =   "Desert Aria"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   48
            Top             =   840
            Width           =   1455
         End
         Begin VB.CheckBox chkDance 
            Caption         =   "Forest Suite"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   47
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox chkDance 
            Caption         =   "Wind Song"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   46
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.ComboBox cboHeroi 
         Height          =   315
         ItemData        =   "frmPrincipal.frx":0074
         Left            =   480
         List            =   "frmPrincipal.frx":0076
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   720
         Width           =   1095
      End
      Begin VB.Frame fraStatus 
         Caption         =   "Status"
         Height          =   2775
         Left            =   3840
         TabIndex        =   33
         Top             =   2760
         Width           =   1455
         Begin VB.CheckBox chkStatus 
            Caption         =   "Darkness"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   42
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Zombie"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   41
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Poisoned"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   40
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Invisible"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   39
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Imp"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   38
            Top             =   1560
            Width           =   615
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Stone"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   37
            Top             =   1800
            Width           =   855
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Magitek"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   36
            Top             =   2040
            Width           =   975
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Wounded"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   35
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Float"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   34
            Top             =   1320
            Width           =   735
         End
      End
      Begin VB.Frame fraEquip 
         Caption         =   "Equipamento"
         Height          =   2775
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   3375
         Begin VB.ComboBox cboEquip 
            Height          =   315
            Index           =   0
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox cboEquip 
            Height          =   315
            Index           =   1
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox cboEquip 
            Height          =   315
            Index           =   4
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1680
            Width           =   1335
         End
         Begin VB.ComboBox cboEquip 
            Height          =   315
            Index           =   2
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1080
            Width           =   1335
         End
         Begin VB.ComboBox cboEquip 
            Height          =   315
            Index           =   5
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1680
            Width           =   1335
         End
         Begin VB.ComboBox cboEquip 
            Height          =   315
            Index           =   3
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1080
            Width           =   1335
         End
         Begin VB.ComboBox cboEquip 
            Height          =   315
            Index           =   6
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Mão Esquerda"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "Mão Direita"
            Height          =   255
            Left            =   1800
            TabIndex        =   31
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label22 
            Caption         =   "1º Relic"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Cabeça"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label23 
            Caption         =   "2º Relic"
            Height          =   255
            Left            =   1800
            TabIndex        =   28
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label21 
            Caption         =   "Corpo"
            Height          =   255
            Left            =   1800
            TabIndex        =   27
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label17 
            Caption         =   "Esper equipado"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   2040
            Width           =   1215
         End
      End
      Begin VB.Frame fraEstatisticas 
         Caption         =   "Estatísticas"
         Height          =   1575
         Left            =   5040
         TabIndex        =   14
         Top             =   1080
         Width           =   1935
         Begin VB.TextBox txtCampo 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   10
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   125
            Top             =   1080
            Width           =   465
         End
         Begin VB.TextBox txtCampo 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   9
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   124
            Top             =   480
            Width           =   465
         End
         Begin VB.TextBox txtCampo 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   8
            Left            =   240
            MaxLength       =   3
            TabIndex        =   123
            Top             =   1080
            Width           =   465
         End
         Begin VB.TextBox txtCampo 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   7
            Left            =   240
            MaxLength       =   3
            TabIndex        =   122
            Top             =   480
            Width           =   465
         End
         Begin MSComCtl2.UpDown updCampo 
            Height          =   315
            Index           =   6
            Left            =   720
            TabIndex        =   143
            Top             =   480
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            BuddyControl    =   "txtCampo(7)"
            BuddyDispid     =   196610
            BuddyIndex      =   7
            OrigLeft        =   720
            OrigTop         =   480
            OrigRight       =   915
            OrigBottom      =   795
            Max             =   255
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown updCampo 
            Height          =   315
            Index           =   7
            Left            =   720
            TabIndex        =   144
            Top             =   1080
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            BuddyControl    =   "txtCampo(8)"
            BuddyDispid     =   196610
            BuddyIndex      =   8
            OrigLeft        =   720
            OrigTop         =   1080
            OrigRight       =   915
            OrigBottom      =   1395
            Max             =   255
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown updCampo 
            Height          =   315
            Index           =   8
            Left            =   1560
            TabIndex        =   145
            Top             =   480
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            BuddyControl    =   "txtCampo(9)"
            BuddyDispid     =   196610
            BuddyIndex      =   9
            OrigLeft        =   1560
            OrigTop         =   480
            OrigRight       =   1755
            OrigBottom      =   795
            Max             =   255
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown updCampo 
            Height          =   315
            Index           =   9
            Left            =   1560
            TabIndex        =   146
            Top             =   1080
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            BuddyControl    =   "txtCampo(10)"
            BuddyDispid     =   196610
            BuddyIndex      =   10
            OrigLeft        =   1560
            OrigTop         =   1080
            OrigRight       =   1755
            OrigBottom      =   1395
            Max             =   255
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label13 
            Caption         =   "Speed"
            Height          =   255
            Left            =   1080
            TabIndex        =   118
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label15 
            Caption         =   "Magic"
            Height          =   255
            Left            =   1080
            TabIndex        =   17
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "Stamina"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   830
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "Vigor"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   210
            Width           =   495
         End
      End
      Begin VB.Frame fraGeral 
         Caption         =   "Geral"
         Height          =   1575
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   2295
         Begin MSComCtl2.UpDown updCampo 
            Height          =   315
            Index           =   0
            Left            =   1861
            TabIndex        =   137
            Top             =   480
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtCampo(1)"
            BuddyDispid     =   196610
            BuddyIndex      =   1
            OrigLeft        =   2040
            OrigTop         =   480
            OrigRight       =   2235
            OrigBottom      =   795
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtCampo 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   2
            Left            =   240
            MaxLength       =   8
            TabIndex        =   121
            Top             =   1080
            Width           =   1620
         End
         Begin VB.TextBox txtCampo 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   1
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   120
            Top             =   480
            Width           =   540
         End
         Begin VB.TextBox txtCampo 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   0
            Left            =   240
            MaxLength       =   6
            TabIndex        =   119
            Top             =   480
            Width           =   855
         End
         Begin MSComCtl2.UpDown updCampo 
            Height          =   315
            Index           =   1
            Left            =   1861
            TabIndex        =   138
            Top             =   1080
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtCampo(2)"
            BuddyDispid     =   196610
            BuddyIndex      =   2
            OrigLeft        =   2040
            OrigTop         =   1080
            OrigRight       =   2235
            OrigBottom      =   1395
            Max             =   1677215
            Min             =   1
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label16 
            Caption         =   "Nome"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Experiência"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Nível"
            Height          =   255
            Left            =   1320
            TabIndex        =   11
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame fraComando 
         Caption         =   "Comandos"
         Height          =   2775
         Left            =   5520
         TabIndex        =   1
         Top             =   2760
         Width           =   1455
         Begin VB.ComboBox cboComando 
            Height          =   315
            Index           =   3
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2280
            Width           =   975
         End
         Begin VB.ComboBox cboComando 
            Height          =   315
            Index           =   2
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1680
            Width           =   975
         End
         Begin VB.ComboBox cboComando 
            Height          =   315
            Index           =   1
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1080
            Width           =   975
         End
         Begin VB.ComboBox cboComando 
            Height          =   315
            Index           =   0
            ItemData        =   "frmPrincipal.frx":0078
            Left            =   240
            List            =   "frmPrincipal.frx":007A
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Comando 4"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Comando 3"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Comando 2"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Comando 1"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Personagem"
         Height          =   255
         Left            =   480
         TabIndex        =   44
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuArquivoAbrir 
         Caption         =   "A&brir..."
      End
      Begin VB.Menu mnuArquivoReabrir 
         Caption         =   "&ReAbrir..."
         Begin VB.Menu mnuArquivoReabrirArq 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuArquivoReabrirArq 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuArquivoReabrirArq 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuArquivoReabrirArq 
            Caption         =   ""
            Index           =   3
         End
      End
      Begin VB.Menu mnuArquivoSalvar 
         Caption         =   "&Salvar"
      End
      Begin VB.Menu mnuArquivoFechar 
         Caption         =   "&Fechar"
      End
      Begin VB.Menu mnuArquivoTab1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArquivoSair 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "&Configurações"
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "A&juda"
      Begin VB.Menu mnuAjudaConteudo 
         Caption         =   "Co&nteúdo..."
      End
      Begin VB.Menu mnuAjudaTab1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAjudaSobre 
         Caption         =   "So&bre..."
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub preencheCampos()
    Dim I As Integer
    
    ' preenche o ComboBox que contém todos os heróis
    cboHeroi.AddItem "Terra", 0
    cboHeroi.AddItem "Locke", 1
    cboHeroi.AddItem "Cyan", 2
    cboHeroi.AddItem "Shadow", 3
    cboHeroi.AddItem "Edgar", 4
    cboHeroi.AddItem "Sabin", 5
    cboHeroi.AddItem "Celes", 6
    cboHeroi.AddItem "Strago", 7
    cboHeroi.AddItem "Relm", 8
    cboHeroi.AddItem "Setzer", 9
    cboHeroi.AddItem "Mog", 10
    cboHeroi.AddItem "Gau", 11
    cboHeroi.AddItem "Gogo", 12
    cboHeroi.AddItem "Umaro", 13
    
    ' preenche o ComboBox que contém
    ' todos os heróis que usam magias
    cboHeroiMagia.AddItem "Terra", 0
    cboHeroiMagia.AddItem "Locke", 1
    cboHeroiMagia.AddItem "Cyan", 2
    cboHeroiMagia.AddItem "Shadow", 3
    cboHeroiMagia.AddItem "Edgar", 4
    cboHeroiMagia.AddItem "Sabin", 5
    cboHeroiMagia.AddItem "Celes", 6
    cboHeroiMagia.AddItem "Strago", 7
    cboHeroiMagia.AddItem "Relm", 8
    cboHeroiMagia.AddItem "Setzer", 9
    cboHeroiMagia.AddItem "Mog", 10
    cboHeroiMagia.AddItem "Gau", 11
    
    ' preenche todas as combos de Comandos
    For I = 0 To 3
        cboComando(I).AddItem "Fight", 0
        cboComando(I).AddItem "Item", 1
        cboComando(I).AddItem "Magic", 2
        cboComando(I).AddItem "Morph", 3
        cboComando(I).AddItem "Revert", 4
        cboComando(I).AddItem "Steal", 5
        cboComando(I).AddItem "Capture", 6
        cboComando(I).AddItem "Swdtech", 7
        cboComando(I).AddItem "Throw", 8
        cboComando(I).AddItem "Tools", 9
        cboComando(I).AddItem "Blitz", 10
        cboComando(I).AddItem "Runic", 11
        cboComando(I).AddItem "Lore", 12
        cboComando(I).AddItem "Sketch", 13
        cboComando(I).AddItem "Control", 14
        cboComando(I).AddItem "Slot", 15
        cboComando(I).AddItem "Rage", 16
        cboComando(I).AddItem "Leap", 17
        cboComando(I).AddItem "Mimic", 18
        cboComando(I).AddItem "Dance", 19
        cboComando(I).AddItem "Row", 20
        cboComando(I).AddItem "Def.", 21
        cboComando(I).AddItem "Jump", 22
        cboComando(I).AddItem "X-Magic", 23
        cboComando(I).AddItem "GP Rain", 24
        cboComando(I).AddItem "Summon", 25
        cboComando(I).AddItem "Health", 26
        cboComando(I).AddItem "Shock", 27
        cboComando(I).AddItem "Possess", 28
        cboComando(I).AddItem "MagiTek", 29
        cboComando(I).AddItem " ", 30
    Next
    
    ' preenche a combo do equipamento da mão esquerda
    cboEquip(0).AddItem "Dirk", 0
    cboEquip(0).AddItem "MithrilKnife", 1
    cboEquip(0).AddItem "Guardian", 2
    cboEquip(0).AddItem "Air Lancet", 3
    cboEquip(0).AddItem "ThiefKnife", 4
    cboEquip(0).AddItem "Assassin", 5
    cboEquip(0).AddItem "Man Eater", 6
    cboEquip(0).AddItem "SwordBreaker", 7
    cboEquip(0).AddItem "Graedus", 8
    cboEquip(0).AddItem "ValiantKnife", 9
    cboEquip(0).AddItem "MithrilBlade", 10
    cboEquip(0).AddItem "RegalCutlass", 11
    cboEquip(0).AddItem "Rune Edge", 12
    cboEquip(0).AddItem "Flame Sabre", 13
    cboEquip(0).AddItem "Blizzard", 14
    cboEquip(0).AddItem "ThunderBlade", 15
    cboEquip(0).AddItem "Epee", 16
    cboEquip(0).AddItem "Break Blade", 17
    cboEquip(0).AddItem "Drainer", 18
    cboEquip(0).AddItem "Enhancer", 19
    cboEquip(0).AddItem "Crystal", 20
    cboEquip(0).AddItem "Falchion", 21
    cboEquip(0).AddItem "Soul Sabre", 22
    cboEquip(0).AddItem "Ogre Nix", 23
    cboEquip(0).AddItem "Excalibur", 24
    cboEquip(0).AddItem "Scimiter", 25
    cboEquip(0).AddItem "Illumina", 26
    cboEquip(0).AddItem "Ragnarok", 27
    cboEquip(0).AddItem "Atma Weapon", 28
    cboEquip(0).AddItem "Mithril Pike", 29
    cboEquip(0).AddItem "Trident", 30
    cboEquip(0).AddItem "Stout Spear", 31
    cboEquip(0).AddItem "Partisan", 32
    cboEquip(0).AddItem "Pearl Lance", 33
    cboEquip(0).AddItem "Gold Lance", 34
    cboEquip(0).AddItem "Aura Lance", 35
    cboEquip(0).AddItem "Imp Halberd", 36
    cboEquip(0).AddItem "Imperial", 37
    cboEquip(0).AddItem "Kodachi", 38
    cboEquip(0).AddItem "Blossom", 39
    cboEquip(0).AddItem "Hardened", 40
    cboEquip(0).AddItem "Striker", 41
    cboEquip(0).AddItem "Stunner", 42
    cboEquip(0).AddItem "Ashura", 43
    cboEquip(0).AddItem "Kotetsu", 44
    cboEquip(0).AddItem "Forged", 45
    cboEquip(0).AddItem "Tempest", 46
    cboEquip(0).AddItem "Murasame", 47
    cboEquip(0).AddItem "Aura", 48
    cboEquip(0).AddItem "Strato", 49
    cboEquip(0).AddItem "Sky Render", 50
    cboEquip(0).AddItem "Heal Rod", 51
    cboEquip(0).AddItem "Mithril Rod", 52
    cboEquip(0).AddItem "Fire Rod", 53
    cboEquip(0).AddItem "Ice Rod", 54
    cboEquip(0).AddItem "Thunder Rod", 55
    cboEquip(0).AddItem "Poison Rod", 56
    cboEquip(0).AddItem "Pearl Rod", 57
    cboEquip(0).AddItem "Gravity Rod", 58
    cboEquip(0).AddItem "Punisher", 59
    cboEquip(0).AddItem "Magus Rod", 60
    cboEquip(0).AddItem "Chocobo Brsh", 61
    cboEquip(0).AddItem "DaVinci Brsh", 62
    cboEquip(0).AddItem "Magical Brsh", 63
    cboEquip(0).AddItem "Rainbow Brsh", 64
    cboEquip(0).AddItem "Shuriken", 65
    cboEquip(0).AddItem "Ninja Star", 66
    cboEquip(0).AddItem "Tack Star", 67
    cboEquip(0).AddItem "Flail", 68
    cboEquip(0).AddItem "Full Moon", 69
    cboEquip(0).AddItem "Morning Star", 70
    cboEquip(0).AddItem "Boomerang", 71
    cboEquip(0).AddItem "Rising Sun", 72
    cboEquip(0).AddItem "Hawk Eye", 73
    cboEquip(0).AddItem "Bone Club", 74
    cboEquip(0).AddItem "Sniper", 75
    cboEquip(0).AddItem "Wing Edge", 76
    cboEquip(0).AddItem "Cards", 77
    cboEquip(0).AddItem "Darts", 78
    cboEquip(0).AddItem "Doom Darts", 79
    cboEquip(0).AddItem "Trump", 80
    cboEquip(0).AddItem "Dice", 81
    cboEquip(0).AddItem "Fixed Dice", 82
    cboEquip(0).AddItem "MetalKnuckle", 83
    cboEquip(0).AddItem "Mithril Claw", 84
    cboEquip(0).AddItem "Kaiser", 85
    cboEquip(0).AddItem "Poison Claw", 86
    cboEquip(0).AddItem "Fire Knuckle", 87
    cboEquip(0).AddItem "Dragon Claw", 88
    cboEquip(0).AddItem "Tiger Fangs", 89
    cboEquip(0).AddItem "Buckler", 90
    cboEquip(0).AddItem "Heavy Shld", 91
    cboEquip(0).AddItem "Mithril Shld", 92
    cboEquip(0).AddItem "Gold Shld", 93
    cboEquip(0).AddItem "Aegis Shld", 94
    cboEquip(0).AddItem "Diamond Shld", 95
    cboEquip(0).AddItem "Flame Shld", 96
    cboEquip(0).AddItem "Ice Shld", 97
    cboEquip(0).AddItem "Thunder Shld", 98
    cboEquip(0).AddItem "Crystal Shld", 99
    cboEquip(0).AddItem "Genji Shld", 100
    cboEquip(0).AddItem "TortoiseShld", 101
    cboEquip(0).AddItem "Cursed Shld", 102
    cboEquip(0).AddItem "Paladin Shld", 103
    cboEquip(0).AddItem "Force Shld", 104
    cboEquip(0).AddItem " ", 105
    
    ' preenche a combo do equipamento da mão direita
    cboEquip(1).AddItem "Dirk", 0
    cboEquip(1).AddItem "MithrilKnife", 1
    cboEquip(1).AddItem "Guardian", 2
    cboEquip(1).AddItem "Air Lancet", 3
    cboEquip(1).AddItem "ThiefKnife", 4
    cboEquip(1).AddItem "Assassin", 5
    cboEquip(1).AddItem "Man Eater", 6
    cboEquip(1).AddItem "SwordBreaker", 7
    cboEquip(1).AddItem "Graedus", 8
    cboEquip(1).AddItem "ValiantKnife", 9
    cboEquip(1).AddItem "MithrilBlade", 10
    cboEquip(1).AddItem "RegalCutlass", 11
    cboEquip(1).AddItem "Rune Edge", 12
    cboEquip(1).AddItem "Flame Sabre", 13
    cboEquip(1).AddItem "Blizzard", 14
    cboEquip(1).AddItem "ThunderBlade", 15
    cboEquip(1).AddItem "Epee", 16
    cboEquip(1).AddItem "Break Blade", 17
    cboEquip(1).AddItem "Drainer", 18
    cboEquip(1).AddItem "Enhancer", 19
    cboEquip(1).AddItem "Crystal", 20
    cboEquip(1).AddItem "Falchion", 21
    cboEquip(1).AddItem "Soul Sabre", 22
    cboEquip(1).AddItem "Ogre Nix", 23
    cboEquip(1).AddItem "Excalibur", 24
    cboEquip(1).AddItem "Scimiter", 25
    cboEquip(1).AddItem "Illumina", 26
    cboEquip(1).AddItem "Ragnarok", 27
    cboEquip(1).AddItem "Atma Weapon", 28
    cboEquip(1).AddItem "Mithril Pike", 29
    cboEquip(1).AddItem "Trident", 30
    cboEquip(1).AddItem "Stout Spear", 31
    cboEquip(1).AddItem "Partisan", 32
    cboEquip(1).AddItem "Pearl Lance", 33
    cboEquip(1).AddItem "Gold Lance", 34
    cboEquip(1).AddItem "Aura Lance", 35
    cboEquip(1).AddItem "Imp Halberd", 36
    cboEquip(1).AddItem "Imperial", 37
    cboEquip(1).AddItem "Kodachi", 38
    cboEquip(1).AddItem "Blossom", 39
    cboEquip(1).AddItem "Hardened", 40
    cboEquip(1).AddItem "Striker", 41
    cboEquip(1).AddItem "Stunner", 42
    cboEquip(1).AddItem "Ashura", 43
    cboEquip(1).AddItem "Kotetsu", 44
    cboEquip(1).AddItem "Forged", 45
    cboEquip(1).AddItem "Tempest", 46
    cboEquip(1).AddItem "Murasame", 47
    cboEquip(1).AddItem "Aura", 48
    cboEquip(1).AddItem "Strato", 49
    cboEquip(1).AddItem "Sky Render", 50
    cboEquip(1).AddItem "Heal Rod", 51
    cboEquip(1).AddItem "Mithril Rod", 52
    cboEquip(1).AddItem "Fire Rod", 53
    cboEquip(1).AddItem "Ice Rod", 54
    cboEquip(1).AddItem "Thunder Rod", 55
    cboEquip(1).AddItem "Poison Rod", 56
    cboEquip(1).AddItem "Pearl Rod", 57
    cboEquip(1).AddItem "Gravity Rod", 58
    cboEquip(1).AddItem "Punisher", 59
    cboEquip(1).AddItem "Magus Rod", 60
    cboEquip(1).AddItem "Chocobo Brsh", 61
    cboEquip(1).AddItem "DaVinci Brsh", 62
    cboEquip(1).AddItem "Magical Brsh", 63
    cboEquip(1).AddItem "Rainbow Brsh", 64
    cboEquip(1).AddItem "Shuriken", 65
    cboEquip(1).AddItem "Ninja Star", 66
    cboEquip(1).AddItem "Tack Star", 67
    cboEquip(1).AddItem "Flail", 68
    cboEquip(1).AddItem "Full Moon", 69
    cboEquip(1).AddItem "Morning Star", 70
    cboEquip(1).AddItem "Boomerang", 71
    cboEquip(1).AddItem "Rising Sun", 72
    cboEquip(1).AddItem "Hawk Eye", 73
    cboEquip(1).AddItem "Bone Club", 74
    cboEquip(1).AddItem "Sniper", 75
    cboEquip(1).AddItem "Wing Edge", 76
    cboEquip(1).AddItem "Cards", 77
    cboEquip(1).AddItem "Darts", 78
    cboEquip(1).AddItem "Doom Darts", 79
    cboEquip(1).AddItem "Trump", 80
    cboEquip(1).AddItem "Dice", 81
    cboEquip(1).AddItem "Fixed Dice", 82
    cboEquip(1).AddItem "MetalKnuckle", 83
    cboEquip(1).AddItem "Mithril Claw", 84
    cboEquip(1).AddItem "Kaiser", 85
    cboEquip(1).AddItem "Poison Claw", 86
    cboEquip(1).AddItem "Fire Knuckle", 87
    cboEquip(1).AddItem "Dragon Claw", 88
    cboEquip(1).AddItem "Tiger Fangs", 89
    cboEquip(1).AddItem "Buckler", 90
    cboEquip(1).AddItem "Heavy Shld", 91
    cboEquip(1).AddItem "Mithril Shld", 92
    cboEquip(1).AddItem "Gold Shld", 93
    cboEquip(1).AddItem "Aegis Shld", 94
    cboEquip(1).AddItem "Diamond Shld", 95
    cboEquip(1).AddItem "Flame Shld", 96
    cboEquip(1).AddItem "Ice Shld", 97
    cboEquip(1).AddItem "Thunder Shld", 98
    cboEquip(1).AddItem "Crystal Shld", 99
    cboEquip(1).AddItem "Genji Shld", 100
    cboEquip(1).AddItem "TortoiseShld", 101
    cboEquip(1).AddItem "Cursed Shld", 102
    cboEquip(1).AddItem "Paladin Shld", 103
    cboEquip(1).AddItem "Force Shld", 104
    cboEquip(1).AddItem " ", 105
    
    ' preenche a combo dos elmos
    cboEquip(2).AddItem "Leather Hat", 0
    cboEquip(2).AddItem "Hair Band", 1
    cboEquip(2).AddItem "Plumed Hat", 2
    cboEquip(2).AddItem "Beret", 3
    cboEquip(2).AddItem "Magus Hat", 4
    cboEquip(2).AddItem "Bandana", 5
    cboEquip(2).AddItem "Iron Helmet", 6
    cboEquip(2).AddItem "Coronet", 7
    cboEquip(2).AddItem "Bard's Hat", 8
    cboEquip(2).AddItem "Green Beret", 9
    cboEquip(2).AddItem "Head Band", 10
    cboEquip(2).AddItem "Mithril Helm", 11
    cboEquip(2).AddItem "Tiara", 12
    cboEquip(2).AddItem "Gold Helmet", 13
    cboEquip(2).AddItem "Tiger Mask", 14
    cboEquip(2).AddItem "Red Hat", 15
    cboEquip(2).AddItem "Mystery Veil", 16
    cboEquip(2).AddItem "Circlet", 17
    cboEquip(2).AddItem "Regal Crown", 18
    cboEquip(2).AddItem "Diamond Helm", 19
    cboEquip(2).AddItem "Dark Hood", 20
    cboEquip(2).AddItem "Crystal Helm", 21
    cboEquip(2).AddItem "Oath Veil", 22
    cboEquip(2).AddItem "Cat Hood", 23
    cboEquip(2).AddItem "Genji Helmet", 24
    cboEquip(2).AddItem "Thornlet", 25
    cboEquip(2).AddItem "Titanium", 26
    cboEquip(2).AddItem " ", 27
    
    ' preenche combo das armaduras
    cboEquip(3).AddItem "LeatherArmor", 0
    cboEquip(3).AddItem "Cotton Robe", 1
    cboEquip(3).AddItem "Kung Fu Suit", 2
    cboEquip(3).AddItem "Iron Armor", 3
    cboEquip(3).AddItem "Silk Robe", 4
    cboEquip(3).AddItem "Mithril Vest", 5
    cboEquip(3).AddItem "Ninja Gear", 6
    cboEquip(3).AddItem "White Dress", 7
    cboEquip(3).AddItem "Mithril Mail", 8
    cboEquip(3).AddItem "Gaia Gear", 9
    cboEquip(3).AddItem "Mirage Dress", 10
    cboEquip(3).AddItem "Gold Armor", 11
    cboEquip(3).AddItem "Power Sash", 12
    cboEquip(3).AddItem "Light Robe", 13
    cboEquip(3).AddItem "Diamond Vest", 14
    cboEquip(3).AddItem "Red Jacket", 15
    cboEquip(3).AddItem "Force Armor", 16
    cboEquip(3).AddItem "DiamondArmor", 17
    cboEquip(3).AddItem "Dark Gear", 18
    cboEquip(3).AddItem "Tao Robe", 19
    cboEquip(3).AddItem "Crystal Mail", 20
    cboEquip(3).AddItem "Czarina Gown", 21
    cboEquip(3).AddItem "Genji Armor", 22
    cboEquip(3).AddItem "Imp's Armor", 23
    cboEquip(3).AddItem "Minerva", 24
    cboEquip(3).AddItem "Tabby Suit", 25
    cboEquip(3).AddItem "Chocobo Suit", 26
    cboEquip(3).AddItem "Moogle Suit", 27
    cboEquip(3).AddItem "Nutkin Suit", 28
    cboEquip(3).AddItem "BehemethSuit", 29
    cboEquip(3).AddItem "Snow Muffler", 30
    cboEquip(3).AddItem " ", 31
    
    'preenche a combo do 1º relic
    cboEquip(4).AddItem "Goggles", 0
    cboEquip(4).AddItem "Star Pendant", 1
    cboEquip(4).AddItem "Peace Ring", 2
    cboEquip(4).AddItem "Amulet", 3
    cboEquip(4).AddItem "White Cape", 4
    cboEquip(4).AddItem "Jewel Ring", 5
    cboEquip(4).AddItem "Fair Ring", 6
    cboEquip(4).AddItem "Barrier Ring", 7
    cboEquip(4).AddItem "MithrilGlove", 8
    cboEquip(4).AddItem "Guard Ring", 9
    cboEquip(4).AddItem "RunningShoes", 10
    cboEquip(4).AddItem "Wall Ring", 11
    cboEquip(4).AddItem "Cherub Down", 12
    cboEquip(4).AddItem "Cure Ring", 13
    cboEquip(4).AddItem "True Knight", 14
    cboEquip(4).AddItem "DragoonBoots", 15
    cboEquip(4).AddItem "Zephyr Cape", 16
    cboEquip(4).AddItem "Czarina Ring", 17
    cboEquip(4).AddItem "Cursed Cing", 18
    cboEquip(4).AddItem "Earrings", 19
    cboEquip(4).AddItem "Atlas Armlet", 20
    cboEquip(4).AddItem "BlizzardRing", 21
    cboEquip(4).AddItem "Rage Ring", 22
    cboEquip(4).AddItem "Sneak Ring", 23
    cboEquip(4).AddItem "Pod Bracelet", 24
    cboEquip(4).AddItem "Hero Ring", 25
    cboEquip(4).AddItem "Ribbon", 26
    cboEquip(4).AddItem "Muscle Belt", 27
    cboEquip(4).AddItem "Crystal Orb", 28
    cboEquip(4).AddItem "Gold Hairpin", 29
    cboEquip(4).AddItem "Economizer", 30
    cboEquip(4).AddItem "Thief Glove", 31
    cboEquip(4).AddItem "Gauntlet", 32
    cboEquip(4).AddItem "Genji Glove", 33
    cboEquip(4).AddItem "Hyper Wrist", 34
    cboEquip(4).AddItem "Offering", 35
    cboEquip(4).AddItem "Beads", 36
    cboEquip(4).AddItem "Black Belt", 37
    cboEquip(4).AddItem "Coin Toss", 38
    cboEquip(4).AddItem "FakeMustache", 39
    cboEquip(4).AddItem "Gem Box", 40
    cboEquip(4).AddItem "Dragon Horn", 41
    cboEquip(4).AddItem "Merit Award", 42
    cboEquip(4).AddItem "Momento Ring", 43
    cboEquip(4).AddItem "Safety Bit", 44
    cboEquip(4).AddItem "Relic Ring", 45
    cboEquip(4).AddItem "Moogle Charm", 46
    cboEquip(4).AddItem "Charm Bangle", 47
    cboEquip(4).AddItem "Marvel Shoes", 48
    cboEquip(4).AddItem "Back Gaurd", 49
    cboEquip(4).AddItem "Gale Hairpin", 50
    cboEquip(4).AddItem "Sniper Sight", 51
    cboEquip(4).AddItem "Exp.Egg", 52
    cboEquip(4).AddItem "Tintinabar", 53
    cboEquip(4).AddItem "Sprint Shoes", 54
    cboEquip(4).AddItem " ", 55
    
    'preenche a combo do 2º relic
    cboEquip(5).AddItem "Goggles", 0
    cboEquip(5).AddItem "Star Pendant", 1
    cboEquip(5).AddItem "Peace Ring", 2
    cboEquip(5).AddItem "Amulet", 3
    cboEquip(5).AddItem "White Cape", 4
    cboEquip(5).AddItem "Jewel Ring", 5
    cboEquip(5).AddItem "Fair Ring", 6
    cboEquip(5).AddItem "Barrier Ring", 7
    cboEquip(5).AddItem "MithrilGlove", 8
    cboEquip(5).AddItem "Guard Ring", 9
    cboEquip(5).AddItem "RunningShoes", 10
    cboEquip(5).AddItem "Wall Ring", 11
    cboEquip(5).AddItem "Cherub Down", 12
    cboEquip(5).AddItem "Cure Ring", 13
    cboEquip(5).AddItem "True Knight", 14
    cboEquip(5).AddItem "DragoonBoots", 15
    cboEquip(5).AddItem "Zephyr Cape", 16
    cboEquip(5).AddItem "Czarina Ring", 17
    cboEquip(5).AddItem "Cursed Cing", 18
    cboEquip(5).AddItem "Earrings", 19
    cboEquip(5).AddItem "Atlas Armlet", 20
    cboEquip(5).AddItem "BlizzardRing", 21
    cboEquip(5).AddItem "Rage Ring", 22
    cboEquip(5).AddItem "Sneak Ring", 23
    cboEquip(5).AddItem "Pod Bracelet", 24
    cboEquip(5).AddItem "Hero Ring", 25
    cboEquip(5).AddItem "Ribbon", 26
    cboEquip(5).AddItem "Muscle Belt", 27
    cboEquip(5).AddItem "Crystal Orb", 28
    cboEquip(5).AddItem "Gold Hairpin", 29
    cboEquip(5).AddItem "Economizer", 30
    cboEquip(5).AddItem "Thief Glove", 31
    cboEquip(5).AddItem "Gauntlet", 32
    cboEquip(5).AddItem "Genji Glove", 33
    cboEquip(5).AddItem "Hyper Wrist", 34
    cboEquip(5).AddItem "Offering", 35
    cboEquip(5).AddItem "Beads", 36
    cboEquip(5).AddItem "Black Belt", 37
    cboEquip(5).AddItem "Coin Toss", 38
    cboEquip(5).AddItem "FakeMustache", 39
    cboEquip(5).AddItem "Gem Box", 40
    cboEquip(5).AddItem "Dragon Horn", 41
    cboEquip(5).AddItem "Merit Award", 42
    cboEquip(5).AddItem "Momento Ring", 43
    cboEquip(5).AddItem "Safety Bit", 44
    cboEquip(5).AddItem "Relic Ring", 45
    cboEquip(5).AddItem "Moogle Charm", 46
    cboEquip(5).AddItem "Charm Bangle", 47
    cboEquip(5).AddItem "Marvel Shoes", 48
    cboEquip(5).AddItem "Back Gaurd", 49
    cboEquip(5).AddItem "Gale Hairpin", 50
    cboEquip(5).AddItem "Sniper Sight", 51
    cboEquip(5).AddItem "Exp.Egg", 52
    cboEquip(5).AddItem "Tintinabar", 53
    cboEquip(5).AddItem "Sprint Shoes", 54
    cboEquip(5).AddItem " ", 55
    
    ' preenche o ComboBox que contém todos os Espers
    cboEquip(6).AddItem "Ramuh", 0
    cboEquip(6).AddItem "Ifrit", 1
    cboEquip(6).AddItem "Shiva", 2
    cboEquip(6).AddItem "Siren", 3
    cboEquip(6).AddItem "Terrato", 4
    cboEquip(6).AddItem "Maduin", 5
    cboEquip(6).AddItem "Shoat", 6
    cboEquip(6).AddItem "Bismark", 7
    cboEquip(6).AddItem "Stray", 8
    cboEquip(6).AddItem "Palidor", 9
    cboEquip(6).AddItem "Tritoch", 10
    cboEquip(6).AddItem "Odin", 11
    cboEquip(6).AddItem "Raiden", 12
    cboEquip(6).AddItem "Bahamut", 13
    cboEquip(6).AddItem "Alexandr", 14
    cboEquip(6).AddItem "Crusader", 15
    cboEquip(6).AddItem "Ragnarok", 16
    cboEquip(6).AddItem "Kirin", 17
    cboEquip(6).AddItem "Zoneseek", 18
    cboEquip(6).AddItem "Carbunkl", 19
    cboEquip(6).AddItem "Phantom", 20
    cboEquip(6).AddItem "Sraphim", 21
    cboEquip(6).AddItem "Golem", 22
    cboEquip(6).AddItem "Unicorn", 23
    cboEquip(6).AddItem "Fenrir", 24
    cboEquip(6).AddItem "Startlet", 25
    cboEquip(6).AddItem "Phoenix", 26
    cboEquip(6).AddItem " ", 27
    
    ' preenche a ListBox que contém os Lore's de Strago
    lstLore.AddItem "Condemned", 0
    lstLore.AddItem "Roulette", 1
    lstLore.AddItem "Clean Sweep", 2
    lstLore.AddItem "Aqua Rake", 3
    lstLore.AddItem "Aero", 4
    lstLore.AddItem "Blow Fish", 5
    lstLore.AddItem "Big Guard", 6
    lstLore.AddItem "Revenge", 7
    lstLore.AddItem "Pearl Wind", 8
    lstLore.AddItem "L.5 Doom", 9
    lstLore.AddItem "L.4 Flare", 10
    lstLore.AddItem "L.3 Muddle", 11
    lstLore.AddItem "Reflect???", 12
    lstLore.AddItem "L? Pearl", 13
    lstLore.AddItem "Step Mine", 14
    lstLore.AddItem "Force Field", 15
    lstLore.AddItem "Dischord", 16
    lstLore.AddItem "Sour Mouth", 17
    lstLore.AddItem "Pep Up", 18
    lstLore.AddItem "Rippler", 19
    lstLore.AddItem "Stone", 20
    lstLore.AddItem "Quasar", 21
    lstLore.AddItem "Grand Train", 22
    lstLore.AddItem "Exploder", 23
    
    ' preenche a ListBox que contém os Rage's de Gau
    lstRage.AddItem "Guard", 0
    lstRage.AddItem "Soldier", 1
    lstRage.AddItem "Templar", 2
    lstRage.AddItem "Ninja", 3
    lstRage.AddItem "Samurai", 4
    lstRage.AddItem "Orog", 5
    lstRage.AddItem "Mag Roader", 6
    lstRage.AddItem "Retainer", 7
    lstRage.AddItem "Hazer", 8
    lstRage.AddItem "Dahling", 9
    lstRage.AddItem "Rain Man", 10
    lstRage.AddItem "Brawler", 11
    lstRage.AddItem "Apokryphos", 12
    lstRage.AddItem "Dark Force", 13
    lstRage.AddItem "Whisper", 14
    lstRage.AddItem "Over-Mind", 15
    lstRage.AddItem "Osteosaur", 16
    lstRage.AddItem "Commander", 17
    lstRage.AddItem "Rhodox", 18
    lstRage.AddItem "Were-Rat", 19
    lstRage.AddItem "Ursus", 20
    lstRage.AddItem "Rhinotaur", 21
    lstRage.AddItem "Steroidite", 22
    lstRage.AddItem "Leafer", 23
    lstRage.AddItem "Stray Cat", 24
    lstRage.AddItem "Lobo", 25
    lstRage.AddItem "Doberman", 26
    lstRage.AddItem "Vomammoth", 27
    lstRage.AddItem "Fidor", 28
    lstRage.AddItem "Baskervor", 29
    lstRage.AddItem "Suriander", 30
    lstRage.AddItem "Chimera", 31
    lstRage.AddItem "Behemoth", 32
    lstRage.AddItem "Mesosaur", 33
    lstRage.AddItem "Pterodon", 34
    lstRage.AddItem "FossilFang", 35
    lstRage.AddItem "White Drgn", 36
    lstRage.AddItem "Doom Drgn", 37
    lstRage.AddItem "Brachosaur", 38
    lstRage.AddItem "Tyranosaur", 39
    lstRage.AddItem "Dark Wind", 40
    lstRage.AddItem "Beakor", 41
    lstRage.AddItem "Vulture", 42
    lstRage.AddItem "Harpy", 43
    lstRage.AddItem "Hermit Crab", 44
    lstRage.AddItem "Trapper", 45
    lstRage.AddItem "Hornet", 46
    lstRage.AddItem "Crasshoppr", 47
    lstRage.AddItem "Delta Bug", 48
    lstRage.AddItem "Gilomantis", 49
    lstRage.AddItem "Trilium", 50
    lstRage.AddItem "Nightshade", 51
    lstRage.AddItem "Tumbleweed", 52
    lstRage.AddItem "Bloompire", 53
    lstRage.AddItem "Trilobiter", 54
    lstRage.AddItem "Siegfried", 55
    lstRage.AddItem "Nautiloid", 56
    lstRage.AddItem "Exocite", 57
    lstRage.AddItem "Anguiform", 58
    lstRage.AddItem "Reach Frog", 59
    lstRage.AddItem "Lizard", 60
    lstRage.AddItem "Chickenlip", 61
    lstRage.AddItem "Hoover", 62
    lstRage.AddItem "Rider", 63
    lstRage.AddItem "Chupon", 64
    lstRage.AddItem "Pipsqueak", 65
    lstRage.AddItem "M-TekArmor", 66
    lstRage.AddItem "Sky Armor", 67
    lstRage.AddItem "Telstar", 68
    lstRage.AddItem "Lethal Wpn", 69
    lstRage.AddItem "Vaporite", 70
    lstRage.AddItem "Flan", 71
    lstRage.AddItem "Ing", 72
    lstRage.AddItem "Humpty", 73
    lstRage.AddItem "Brainpan", 74
    lstRage.AddItem "Cruller", 75
    lstRage.AddItem "Cactrot", 76
    lstRage.AddItem "RepoMan", 77
    lstRage.AddItem "Harvester", 78
    lstRage.AddItem "Bomb", 79
    lstRage.AddItem "Still Life", 80
    lstRage.AddItem "Boxed Set", 81
    lstRage.AddItem "Slam Dancer", 82
    lstRage.AddItem "Hades Gigas", 83
    lstRage.AddItem "Pug", 84
    lstRage.AddItem "Magic Urn", 85
    lstRage.AddItem "Mover", 86
    lstRage.AddItem "Figaliz", 87
    lstRage.AddItem "Buffalax", 88
    lstRage.AddItem "Aspik", 89
    lstRage.AddItem "Ghost", 90
    lstRage.AddItem "Crawler", 91
    lstRage.AddItem "Sand Ray", 92
    lstRage.AddItem "Areneid", 93
    lstRage.AddItem "Actaneon", 94
    lstRage.AddItem "Sand Horse", 95
    lstRage.AddItem "Dark Side", 96
    lstRage.AddItem "Mad Oscar", 97
    lstRage.AddItem "Crawly", 98
    lstRage.AddItem "Bleary", 99
    lstRage.AddItem "Marshal", 100
    lstRage.AddItem "Trooper", 101
    lstRage.AddItem "General", 102
    lstRage.AddItem "Covert", 103
    lstRage.AddItem "Ogor", 104
    lstRage.AddItem "Warlock", 105
    lstRage.AddItem "Madam", 106
    lstRage.AddItem "Joker", 107
    lstRage.AddItem "Iron Fist", 108
    lstRage.AddItem "Goblin", 109
    lstRage.AddItem "Apparite", 110
    lstRage.AddItem "PowerDemon", 111
    lstRage.AddItem "Displayer", 112
    lstRage.AddItem "Vector Pup", 113
    lstRage.AddItem "Peepers", 114
    lstRage.AddItem "Sewer Rat", 115
    lstRage.AddItem "Slatter", 116
    lstRage.AddItem "Rhinox", 117
    lstRage.AddItem "Rhobite", 118
    lstRage.AddItem "Wild Cat", 119
    lstRage.AddItem "Red Fang", 120
    lstRage.AddItem "Bounty Man", 121
    lstRage.AddItem "Tusker", 122
    lstRage.AddItem "Ralph", 123
    lstRage.AddItem "Chitonid", 124
    lstRage.AddItem "Wart Puck", 125
    lstRage.AddItem "Rhyos", 126
    lstRage.AddItem "SrBehemoth", 127
    lstRage.AddItem "Vectaur", 128
    lstRage.AddItem "Wyvern", 129
    lstRage.AddItem "Zombone", 130
    lstRage.AddItem "Dragon", 131
    lstRage.AddItem "Brontaur", 132
    lstRage.AddItem "Allosaurus", 133
    lstRage.AddItem "Cirpius", 134
    lstRage.AddItem "Sprinter", 135
    lstRage.AddItem "Gobbler", 136
    lstRage.AddItem "Harpai", 137
    lstRage.AddItem "Gloomshell", 138
    lstRage.AddItem "Drop", 139
    lstRage.AddItem "Mind Candy", 140
    lstRage.AddItem "WeedFeeder", 141
    lstRage.AddItem "Luridan", 142
    lstRage.AddItem "Toe Cutter", 143
    lstRage.AddItem "Over Grunk", 144
    lstRage.AddItem "Exoray", 145
    lstRage.AddItem "Crusher", 146
    lstRage.AddItem "Uroburos", 147
    lstRage.AddItem "Primordite", 148
    lstRage.AddItem "Sky Cap", 149
    lstRage.AddItem "Cephaler", 150
    lstRage.AddItem "Maliga", 151
    lstRage.AddItem "Gigan Toad", 152
    lstRage.AddItem "Geckorex", 153
    lstRage.AddItem "Cluck", 154
    lstRage.AddItem "LandWorm:", 155
    lstRage.AddItem "Test Rider", 156
    lstRage.AddItem "PlutoArmor", 157
    lstRage.AddItem "Tomb Thumb", 158
    lstRage.AddItem "HeavyArmor", 159
    lstRage.AddItem "Chaser", 160
    lstRage.AddItem "Scullion", 161
    lstRage.AddItem "Poplium", 162
    lstRage.AddItem "Intangir", 163
    lstRage.AddItem "Misfit", 164
    lstRage.AddItem "Eland", 165
    lstRage.AddItem "Enuo", 166
    lstRage.AddItem "Deep Eye", 167
    lstRage.AddItem "GreaseMonk", 168
    lstRage.AddItem "NeckHunter", 169
    lstRage.AddItem "Grenadeb", 170
    lstRage.AddItem "Critic", 171
    lstRage.AddItem "Pan Dora", 172
    lstRage.AddItem "SoulDancer", 173
    lstRage.AddItem "Gigantos", 174
    lstRage.AddItem "Mag Roader", 175
    lstRage.AddItem "Spek Tor", 176
    lstRage.AddItem "Parasite", 177
    lstRage.AddItem "EarthGuard", 178
    lstRage.AddItem "Coelecite", 179
    lstRage.AddItem "Anemone", 180
    lstRage.AddItem "Hipocampus", 181
    lstRage.AddItem "Spectre", 182
    lstRage.AddItem "Evil Oscar", 183
    lstRage.AddItem "Slurm", 184
    lstRage.AddItem "Latimeria", 185
    lstRage.AddItem "StillGoing", 186
    lstRage.AddItem "Allo Ver", 187
    lstRage.AddItem "Phase", 188
    lstRage.AddItem "Outsider", 189
    lstRage.AddItem "Barb-e", 190
    lstRage.AddItem "Parasoul", 191
    lstRage.AddItem "Pm Stalker", 192
    lstRage.AddItem "Hemophyte", 193
    lstRage.AddItem "Sp Forces", 194
    lstRage.AddItem "Nohrabbit", 195
    lstRage.AddItem "Wizard", 196
    lstRage.AddItem "Scrapper", 197
    lstRage.AddItem "Ceritops", 198
    lstRage.AddItem "Commando", 199
    lstRage.AddItem "Opinicus", 200
    lstRage.AddItem "Poppers", 201
    lstRage.AddItem "Lunaris", 202
    lstRage.AddItem "Garm", 203
    lstRage.AddItem "Vindr", 204
    lstRage.AddItem "Kiwak", 205
    lstRage.AddItem "Nastidon", 206
    lstRage.AddItem "Rinn", 207
    lstRage.AddItem "Insecare", 208
    lstRage.AddItem "Vermin", 209
    lstRage.AddItem "Mantodea", 210
    lstRage.AddItem "Bogy", 211
    lstRage.AddItem "Prussian", 212
    lstRage.AddItem "Black Drgn", 213
    lstRage.AddItem "Adamanchyt", 214
    lstRage.AddItem "Dante", 215
    lstRage.AddItem "Wirey Drgn", 216
    lstRage.AddItem "Dueller", 217
    lstRage.AddItem "Psycot", 218
    lstRage.AddItem "Muus", 219
    lstRage.AddItem "Karkass", 220
    lstRage.AddItem "Punisher", 221
    lstRage.AddItem "Balloon", 222
    lstRage.AddItem "Gabbldegak", 223
    lstRage.AddItem "GtBehemoth", 224
    lstRage.AddItem "Scorpion", 225
    lstRage.AddItem "Chaos Drgn", 226
    lstRage.AddItem "Spit Fire", 227
    lstRage.AddItem "Vectagoyle", 228
    lstRage.AddItem "Lick", 229
    lstRage.AddItem "Osprey", 230
    lstRage.AddItem "Mag Roader", 231
    lstRage.AddItem "Bug", 232
    lstRage.AddItem "Sea Flower", 233
    lstRage.AddItem "Fortis", 234
    lstRage.AddItem "Abolisher", 235
    lstRage.AddItem "Aquila", 236
    lstRage.AddItem "Junk", 237
    lstRage.AddItem "Mandrake", 238
    lstRage.AddItem "1st Class", 239
    lstRage.AddItem "Tap Dancer", 240
    lstRage.AddItem "Necromancer", 241
    lstRage.AddItem "Borras", 242
    lstRage.AddItem "Mag Roader", 243
    lstRage.AddItem "Wild Rat", 244
    lstRage.AddItem "Gold Bear", 245
    lstRage.AddItem "Innoc", 246
    lstRage.AddItem "Trixter", 247
    lstRage.AddItem "Red Wolf", 248
    lstRage.AddItem "Didalos", 249
    lstRage.AddItem "Woolly", 250
    lstRage.AddItem "Veteran", 251
    lstRage.AddItem "Sky Base", 252
    lstRage.AddItem "IronHitman", 253
    lstRage.AddItem "Io", 254
    
    ' preenche a ListBox que contém oa nomes das magias
    cboMagias.AddItem "Fire", 0
    cboMagias.AddItem "Ice", 1
    cboMagias.AddItem "Bolt", 2
    cboMagias.AddItem "Poison", 3
    cboMagias.AddItem "Drain", 4
    cboMagias.AddItem "Fire2", 5
    cboMagias.AddItem "Ice2", 6
    cboMagias.AddItem "Bolt2", 7
    cboMagias.AddItem "Bio", 8
    cboMagias.AddItem "Fire3", 9
    cboMagias.AddItem "Ice3", 10
    cboMagias.AddItem "Bolt3", 11
    cboMagias.AddItem "Break", 12
    cboMagias.AddItem "Doom", 13
    cboMagias.AddItem "Pearl", 14
    cboMagias.AddItem "Flare", 15
    cboMagias.AddItem "Demi", 16
    cboMagias.AddItem "Quartr", 17
    cboMagias.AddItem "X -Zone", 18
    cboMagias.AddItem "Meteor", 19
    cboMagias.AddItem "Ultima", 20
    cboMagias.AddItem "Quake", 21
    cboMagias.AddItem "W Wind", 22
    cboMagias.AddItem "Merton", 23
    cboMagias.AddItem "Scan", 24
    cboMagias.AddItem "Slow", 25
    cboMagias.AddItem "Rasp", 26
    cboMagias.AddItem "Mute", 27
    cboMagias.AddItem "Safe", 28
    cboMagias.AddItem "Sleep", 29
    cboMagias.AddItem "Muddle", 30
    cboMagias.AddItem "Haste", 31
    cboMagias.AddItem "Stop", 32
    cboMagias.AddItem "Bserk", 33
    cboMagias.AddItem "Float", 34
    cboMagias.AddItem "Imp", 35
    cboMagias.AddItem "Rflect", 36
    cboMagias.AddItem "Shell", 37
    cboMagias.AddItem "Vanish", 38
    cboMagias.AddItem "Haste2", 39
    cboMagias.AddItem "Slow2", 40
    cboMagias.AddItem "Osmose", 41
    cboMagias.AddItem "Warp", 42
    cboMagias.AddItem "Quick", 43
    cboMagias.AddItem "Dispel", 44
    cboMagias.AddItem "Cure", 45
    cboMagias.AddItem "Cure2", 46
    cboMagias.AddItem "Cure3", 47
    cboMagias.AddItem "Life", 48
    cboMagias.AddItem "Life2", 49
    cboMagias.AddItem "Antidot", 50
    cboMagias.AddItem "Remedy", 51
    cboMagias.AddItem "Regen", 52
    cboMagias.AddItem "Life3", 53
    
    ' adiciona os itens na combobox de itens
    cboItem.AddItem "Dirk", 0
    cboItem.ItemData(0) = 2000
    cboItem.AddItem "MithrilKnife", 1
    cboItem.ItemData(1) = 2001
    cboItem.AddItem "Guardian", 2
    cboItem.ItemData(2) = 2002
    cboItem.AddItem "Air Lancet", 3
    cboItem.ItemData(3) = 2003
    cboItem.AddItem "ThiefKnife", 4
    cboItem.ItemData(4) = 2004
    cboItem.AddItem "Assassin", 5
    cboItem.ItemData(5) = 2005
    cboItem.AddItem "Man Eater", 6
    cboItem.ItemData(6) = 2006
    cboItem.AddItem "SwordBreaker", 7
    cboItem.ItemData(7) = 2007
    cboItem.AddItem "Graedus", 8
    cboItem.ItemData(8) = 2008
    cboItem.AddItem "ValiantKnife", 9
    cboItem.ItemData(9) = 2009
    cboItem.AddItem "MithrilBlade", 10
    cboItem.ItemData(10) = 3010
    cboItem.AddItem "RegalCutlass", 11
    cboItem.ItemData(11) = 3011
    cboItem.AddItem "Rune Edge", 12
    cboItem.ItemData(12) = 3012
    cboItem.AddItem "Flame Sabre", 13
    cboItem.ItemData(13) = 3013
    cboItem.AddItem "Blizzard", 14
    cboItem.ItemData(14) = 3014
    cboItem.AddItem "ThunderBlade", 15
    cboItem.ItemData(15) = 3015
    cboItem.AddItem "Epee", 16
    cboItem.ItemData(16) = 3016
    cboItem.AddItem "Break Blade", 17
    cboItem.ItemData(17) = 3017
    cboItem.AddItem "Drainer", 18
    cboItem.ItemData(18) = 3018
    cboItem.AddItem "Enhancer", 19
    cboItem.ItemData(19) = 3019
    cboItem.AddItem "Crystal", 20
    cboItem.ItemData(20) = 3020
    cboItem.AddItem "Falchion", 21
    cboItem.ItemData(21) = 3021
    cboItem.AddItem "Soul Sabre", 22
    cboItem.ItemData(22) = 3022
    cboItem.AddItem "Ogre Nix", 23
    cboItem.ItemData(23) = 3023
    cboItem.AddItem "Excalibur", 24
    cboItem.ItemData(24) = 3024
    cboItem.AddItem "Scimiter", 25
    cboItem.ItemData(25) = 3025
    cboItem.AddItem "Illumina", 26
    cboItem.ItemData(26) = 3026
    cboItem.AddItem "Ragnarok", 27
    cboItem.ItemData(27) = 3027
    cboItem.AddItem "Atma Weapon", 28
    cboItem.ItemData(28) = 3028
    cboItem.AddItem "Mithril Pike", 29
    cboItem.ItemData(29) = 4029
    cboItem.AddItem "Trident", 30
    cboItem.ItemData(30) = 4030
    cboItem.AddItem "Stout Spear", 31
    cboItem.ItemData(31) = 4031
    cboItem.AddItem "Partisan", 32
    cboItem.ItemData(32) = 4032
    cboItem.AddItem "Pearl Lance", 33
    cboItem.ItemData(33) = 4033
    cboItem.AddItem "Gold Lance", 34
    cboItem.ItemData(34) = 4034
    cboItem.AddItem "Aura Lance", 35
    cboItem.ItemData(35) = 4035
    cboItem.AddItem "Imp Halberd", 36
    cboItem.ItemData(36) = 4036
    cboItem.AddItem "Imperial", 37
    cboItem.ItemData(37) = 5037
    cboItem.AddItem "Kodachi", 38
    cboItem.ItemData(38) = 5038
    cboItem.AddItem "Blossom", 39
    cboItem.ItemData(39) = 5039
    cboItem.AddItem "Hardened", 40
    cboItem.ItemData(40) = 5040
    cboItem.AddItem "Striker", 41
    cboItem.ItemData(41) = 5041
    cboItem.AddItem "Stunner", 42
    cboItem.ItemData(42) = 5042
    cboItem.AddItem "Ashura", 43
    cboItem.ItemData(43) = 5043
    cboItem.AddItem "Kotetsu", 44
    cboItem.ItemData(44) = 5044
    cboItem.AddItem "Forged", 45
    cboItem.ItemData(45) = 5045
    cboItem.AddItem "Tempest", 46
    cboItem.ItemData(46) = 5046
    cboItem.AddItem "Murasame", 47
    cboItem.ItemData(47) = 5047
    cboItem.AddItem "Aura", 48
    cboItem.ItemData(48) = 5048
    cboItem.AddItem "Strato", 49
    cboItem.ItemData(49) = 5049
    cboItem.AddItem "Sky Render", 50
    cboItem.ItemData(50) = 5050
    cboItem.AddItem "Heal Rod", 51
    cboItem.ItemData(51) = 6051
    cboItem.AddItem "Mithril Rod", 52
    cboItem.ItemData(52) = 6052
    cboItem.AddItem "Fire Rod", 53
    cboItem.ItemData(53) = 6053
    cboItem.AddItem "Ice Rod", 54
    cboItem.ItemData(54) = 6054
    cboItem.AddItem "Thunder Rod", 55
    cboItem.ItemData(55) = 6055
    cboItem.AddItem "Poison Rod", 56
    cboItem.ItemData(56) = 6056
    cboItem.AddItem "Pearl Rod", 57
    cboItem.ItemData(57) = 6057
    cboItem.AddItem "Gravity Rod", 58
    cboItem.ItemData(58) = 6058
    cboItem.AddItem "Punisher", 59
    cboItem.ItemData(59) = 6059
    cboItem.AddItem "Magus Rod", 60
    cboItem.ItemData(60) = 6060
    cboItem.AddItem "Chocobo Brsh", 61
    cboItem.ItemData(61) = 7061
    cboItem.AddItem "DaVinci Brsh", 62
    cboItem.ItemData(62) = 7062
    cboItem.AddItem "Magical Brsh", 63
    cboItem.ItemData(63) = 7063
    cboItem.AddItem "Rainbow Brsh", 64
    cboItem.ItemData(64) = 7064
    cboItem.AddItem "Shuriken", 65
    cboItem.ItemData(65) = 8065
    cboItem.AddItem "Ninja Star", 66
    cboItem.ItemData(66) = 8066
    cboItem.AddItem "Tack Star", 67
    cboItem.ItemData(67) = 8067
    cboItem.AddItem "Flail", 68
    cboItem.ItemData(68) = 9068
    cboItem.AddItem "Full Moon", 69
    cboItem.ItemData(69) = 9069
    cboItem.AddItem "Morning Star", 70
    cboItem.ItemData(70) = 9070
    cboItem.AddItem "Boomerang", 71
    cboItem.ItemData(71) = 9071
    cboItem.AddItem "Rising Sun", 72
    cboItem.ItemData(72) = 9072
    cboItem.AddItem "Hawk Eye", 73
    cboItem.ItemData(73) = 9073
    cboItem.AddItem "Bone Club", 74
    cboItem.ItemData(74) = 9074
    cboItem.AddItem "Sniper", 75
    cboItem.ItemData(75) = 9075
    cboItem.AddItem "Wing Edge", 76
    cboItem.ItemData(76) = 9076
    cboItem.AddItem "Cards", 77
    cboItem.ItemData(77) = 10077
    cboItem.AddItem "Darts", 78
    cboItem.ItemData(78) = 10078
    cboItem.AddItem "Doom Darts", 79
    cboItem.ItemData(79) = 10079
    cboItem.AddItem "Trump", 80
    cboItem.ItemData(80) = 10080
    cboItem.AddItem "Dice", 81
    cboItem.ItemData(81) = 10081
    cboItem.AddItem "Fixed Dice", 82
    cboItem.ItemData(82) = 10082
    cboItem.AddItem "MetalKnuckle", 83
    cboItem.ItemData(83) = 11083
    cboItem.AddItem "Mithril Claw", 84
    cboItem.ItemData(84) = 11084
    cboItem.AddItem "Kaiser", 85
    cboItem.ItemData(85) = 11085
    cboItem.AddItem "Poison Claw", 86
    cboItem.ItemData(86) = 11086
    cboItem.AddItem "Fire Knuckle", 87
    cboItem.ItemData(87) = 11087
    cboItem.AddItem "Dragon Claw", 88
    cboItem.ItemData(88) = 11088
    cboItem.AddItem "Tiger Fangs", 89
    cboItem.ItemData(89) = 11089
    cboItem.AddItem "Buckler", 90
    cboItem.ItemData(90) = 12090
    cboItem.AddItem "Heavy Shld", 91
    cboItem.ItemData(91) = 12091
    cboItem.AddItem "Mithril Shld", 92
    cboItem.ItemData(92) = 12092
    cboItem.AddItem "Gold Shld", 93
    cboItem.ItemData(93) = 12093
    cboItem.AddItem "Aegis Shld", 94
    cboItem.ItemData(94) = 12094
    cboItem.AddItem "Diamond Shld", 95
    cboItem.ItemData(95) = 12095
    cboItem.AddItem "Flame Shld", 96
    cboItem.ItemData(96) = 12096
    cboItem.AddItem "Ice Shld", 97
    cboItem.ItemData(97) = 12097
    cboItem.AddItem "Thunder Shld", 98
    cboItem.ItemData(98) = 12098
    cboItem.AddItem "Crystal Shld", 99
    cboItem.ItemData(99) = 12099
    cboItem.AddItem "Genji Shld", 100
    cboItem.ItemData(100) = 12100
    cboItem.AddItem "TortoiseShld", 101
    cboItem.ItemData(101) = 12101
    cboItem.AddItem "Cursed Shld", 102
    cboItem.ItemData(102) = 12102
    cboItem.AddItem "Paladin Shld", 103
    cboItem.ItemData(103) = 12103
    cboItem.AddItem "Force Shld", 104
    cboItem.ItemData(104) = 12104
    cboItem.AddItem "Leather Hat", 105
    cboItem.ItemData(105) = 13105
    cboItem.AddItem "Hair Band", 106
    cboItem.ItemData(106) = 13106
    cboItem.AddItem "Plumed Hat", 107
    cboItem.ItemData(107) = 13107
    cboItem.AddItem "Beret", 108
    cboItem.ItemData(108) = 13108
    cboItem.AddItem "Magus Hat", 109
    cboItem.ItemData(109) = 13109
    cboItem.AddItem "Bandana", 110
    cboItem.ItemData(110) = 13110
    cboItem.AddItem "Iron Helmet", 111
    cboItem.ItemData(111) = 13111
    cboItem.AddItem "Coronet", 112
    cboItem.ItemData(112) = 13112
    cboItem.AddItem "Bard's Hat", 113
    cboItem.ItemData(113) = 13113
    cboItem.AddItem "Green Beret", 114
    cboItem.ItemData(114) = 13114
    cboItem.AddItem "Head Band", 115
    cboItem.ItemData(115) = 13115
    cboItem.AddItem "Mithril Helm", 116
    cboItem.ItemData(116) = 13116
    cboItem.AddItem "Tiara", 117
    cboItem.ItemData(117) = 13117
    cboItem.AddItem "Gold Helmet", 118
    cboItem.ItemData(118) = 13118
    cboItem.AddItem "Tiger Mask", 119
    cboItem.ItemData(119) = 13119
    cboItem.AddItem "Red Hat", 120
    cboItem.ItemData(120) = 13120
    cboItem.AddItem "Mystery Veil", 121
    cboItem.ItemData(121) = 13121
    cboItem.AddItem "Circlet", 122
    cboItem.ItemData(122) = 13122
    cboItem.AddItem "Regal Crown", 123
    cboItem.ItemData(123) = 13123
    cboItem.AddItem "Diamond Helm", 124
    cboItem.ItemData(124) = 13124
    cboItem.AddItem "Dark Hood", 125
    cboItem.ItemData(125) = 13125
    cboItem.AddItem "Crystal Helm", 126
    cboItem.ItemData(126) = 13126
    cboItem.AddItem "Oath Veil", 127
    cboItem.ItemData(127) = 13127
    cboItem.AddItem "Cat Hood", 128
    cboItem.ItemData(128) = 13128
    cboItem.AddItem "Genji Helmet", 129
    cboItem.ItemData(129) = 13129
    cboItem.AddItem "Thornlet", 130
    cboItem.ItemData(130) = 13130
    cboItem.AddItem "Titanium", 131
    cboItem.ItemData(131) = 13131
    cboItem.AddItem "LeatherArmor", 132
    cboItem.ItemData(132) = 14132
    cboItem.AddItem "Cotton Robe", 133
    cboItem.ItemData(133) = 14133
    cboItem.AddItem "Kung Fu Suit", 134
    cboItem.ItemData(134) = 14134
    cboItem.AddItem "Iron Armor", 135
    cboItem.ItemData(135) = 14135
    cboItem.AddItem "Silk Robe", 136
    cboItem.ItemData(136) = 14136
    cboItem.AddItem "Mithril Vest", 137
    cboItem.ItemData(137) = 14137
    cboItem.AddItem "Ninja Gear", 138
    cboItem.ItemData(138) = 14138
    cboItem.AddItem "White Dress", 139
    cboItem.ItemData(139) = 14139
    cboItem.AddItem "Mithril Mail", 140
    cboItem.ItemData(140) = 14140
    cboItem.AddItem "Gaia Gear", 141
    cboItem.ItemData(141) = 14141
    cboItem.AddItem "Mirage Dress", 142
    cboItem.ItemData(142) = 14142
    cboItem.AddItem "Gold Armor", 143
    cboItem.ItemData(143) = 14143
    cboItem.AddItem "Power Sash", 144
    cboItem.ItemData(144) = 14144
    cboItem.AddItem "Light Robe", 145
    cboItem.ItemData(145) = 14145
    cboItem.AddItem "Diamond Vest", 146
    cboItem.ItemData(146) = 14146
    cboItem.AddItem "Red Jacket", 147
    cboItem.ItemData(147) = 14147
    cboItem.AddItem "Force Armor", 148
    cboItem.ItemData(148) = 14148
    cboItem.AddItem "DiamondArmor", 149
    cboItem.ItemData(149) = 14149
    cboItem.AddItem "Dark Gear", 150
    cboItem.ItemData(150) = 14150
    cboItem.AddItem "Tao Robe", 151
    cboItem.ItemData(151) = 14151
    cboItem.AddItem "Crystal Mail", 152
    cboItem.ItemData(152) = 14152
    cboItem.AddItem "Czarina Gown", 153
    cboItem.ItemData(153) = 14153
    cboItem.AddItem "Genji Armor", 154
    cboItem.ItemData(154) = 14154
    cboItem.AddItem "Imp's Armor", 155
    cboItem.ItemData(155) = 14155
    cboItem.AddItem "Minerva", 156
    cboItem.ItemData(156) = 14156
    cboItem.AddItem "Tabby Suit", 157
    cboItem.ItemData(157) = 14157
    cboItem.AddItem "Chocobo Suit", 158
    cboItem.ItemData(158) = 14158
    cboItem.AddItem "Moogle Suit", 159
    cboItem.ItemData(159) = 14159
    cboItem.AddItem "Nutkin Suit", 160
    cboItem.ItemData(160) = 14160
    cboItem.AddItem "BehemethSuit", 161
    cboItem.ItemData(161) = 14161
    cboItem.AddItem "Snow Muffler", 162
    cboItem.ItemData(162) = 14163
    cboItem.AddItem "NoiseBlaster", 163
    cboItem.ItemData(163) = 15163
    cboItem.AddItem "Bio Blaster", 164
    cboItem.ItemData(164) = 15164
    cboItem.AddItem "Flash", 165
    cboItem.ItemData(165) = 15165
    cboItem.AddItem "Chain Saw", 166
    cboItem.ItemData(166) = 15166
    cboItem.AddItem "Debilitator", 167
    cboItem.ItemData(167) = 15167
    cboItem.AddItem "Drill", 168
    cboItem.ItemData(168) = 15168
    cboItem.AddItem "Air Anchor", 169
    cboItem.ItemData(169) = 15169
    cboItem.AddItem "AutoCrossbow", 170
    cboItem.ItemData(170) = 15170
    cboItem.AddItem "Fire Skean", 171
    cboItem.ItemData(171) = 16171
    cboItem.AddItem "Water Edge", 172
    cboItem.ItemData(172) = 16172
    cboItem.AddItem "Bolt Edge", 173
    cboItem.ItemData(173) = 16173
    cboItem.AddItem "Inviz Edge", 174
    cboItem.ItemData(174) = 16174
    cboItem.AddItem "Shadow Edge", 175
    cboItem.ItemData(175) = 16175
    cboItem.AddItem "Goggles", 176
    cboItem.ItemData(176) = 17176
    cboItem.AddItem "Star Pendant", 177
    cboItem.ItemData(177) = 17177
    cboItem.AddItem "Peace Ring", 178
    cboItem.ItemData(178) = 17178
    cboItem.AddItem "Amulet", 179
    cboItem.ItemData(179) = 17179
    cboItem.AddItem "White Cape", 180
    cboItem.ItemData(180) = 17180
    cboItem.AddItem "Jewel Ring", 181
    cboItem.ItemData(181) = 17181
    cboItem.AddItem "Fair Ring", 182
    cboItem.ItemData(182) = 17182
    cboItem.AddItem "Barrier Ring", 183
    cboItem.ItemData(183) = 17183
    cboItem.AddItem "MithrilGlove", 184
    cboItem.ItemData(184) = 17184
    cboItem.AddItem "Guard Ring", 185
    cboItem.ItemData(185) = 17185
    cboItem.AddItem "RunningShoes", 186
    cboItem.ItemData(186) = 17186
    cboItem.AddItem "Wall Ring", 187
    cboItem.ItemData(187) = 17187
    cboItem.AddItem "Cherub Down", 188
    cboItem.ItemData(188) = 17188
    cboItem.AddItem "Cure Ring", 189
    cboItem.ItemData(189) = 17189
    cboItem.AddItem "True Knight", 190
    cboItem.ItemData(190) = 17190
    cboItem.AddItem "DragoonBoots", 191
    cboItem.ItemData(191) = 17191
    cboItem.AddItem "Zephyr Cape", 192
    cboItem.ItemData(192) = 17192
    cboItem.AddItem "Czarina Ring", 193
    cboItem.ItemData(193) = 17193
    cboItem.AddItem "Cursed Ring", 194
    cboItem.ItemData(194) = 17194
    cboItem.AddItem "Earrings", 195
    cboItem.ItemData(195) = 17195
    cboItem.AddItem "Atlas Armlet", 196
    cboItem.ItemData(196) = 17196
    cboItem.AddItem "BlizzardRing", 197
    cboItem.ItemData(197) = 17197
    cboItem.AddItem "Rage Ring", 198
    cboItem.ItemData(198) = 17198
    cboItem.AddItem "Sneak Ring", 199
    cboItem.ItemData(199) = 17199
    cboItem.AddItem "Pod Bracelet", 200
    cboItem.ItemData(200) = 17200
    cboItem.AddItem "Hero Ring", 201
    cboItem.ItemData(201) = 17201
    cboItem.AddItem "Ribbon", 202
    cboItem.ItemData(202) = 17202
    cboItem.AddItem "Muscle Belt", 203
    cboItem.ItemData(203) = 17203
    cboItem.AddItem "Crystal Orb", 204
    cboItem.ItemData(204) = 17204
    cboItem.AddItem "Gold Hairpin", 205
    cboItem.ItemData(205) = 17205
    cboItem.AddItem "Economizer", 206
    cboItem.ItemData(206) = 17206
    cboItem.AddItem "Thief Glove", 207
    cboItem.ItemData(207) = 17207
    cboItem.AddItem "Gauntlet", 208
    cboItem.ItemData(208) = 17208
    cboItem.AddItem "Genji Glove", 209
    cboItem.ItemData(209) = 17209
    cboItem.AddItem "Hyper Wrist", 210
    cboItem.ItemData(210) = 17210
    cboItem.AddItem "Offering", 211
    cboItem.ItemData(211) = 17211
    cboItem.AddItem "Beads", 212
    cboItem.ItemData(212) = 17212
    cboItem.AddItem "Black Belt", 213
    cboItem.ItemData(213) = 17213
    cboItem.AddItem "Coin Toss", 214
    cboItem.ItemData(214) = 17214
    cboItem.AddItem "FakeMustache", 215
    cboItem.ItemData(215) = 17215
    cboItem.AddItem "Gem Box", 216
    cboItem.ItemData(216) = 17216
    cboItem.AddItem "Dragon Horn", 217
    cboItem.ItemData(217) = 17217
    cboItem.AddItem "Merit Award", 218
    cboItem.ItemData(218) = 17218
    cboItem.AddItem "Momento Ring", 219
    cboItem.ItemData(219) = 17219
    cboItem.AddItem "Safety Bit", 220
    cboItem.ItemData(220) = 17220
    cboItem.AddItem "Relic Ring", 221
    cboItem.ItemData(221) = 17221
    cboItem.AddItem "Moogle Charm", 222
    cboItem.ItemData(222) = 17222
    cboItem.AddItem "Charm Bangle", 223
    cboItem.ItemData(223) = 17223
    cboItem.AddItem "Marvel Shoes", 224
    cboItem.ItemData(224) = 17224
    cboItem.AddItem "Back Guard", 225
    cboItem.ItemData(225) = 17225
    cboItem.AddItem "Gale Hairpin", 226
    cboItem.ItemData(226) = 17226
    cboItem.AddItem "Sniper Sight", 227
    cboItem.ItemData(227) = 17227
    cboItem.AddItem "Exp.Egg", 228
    cboItem.ItemData(228) = 17228
    cboItem.AddItem "Tintinabar", 229
    cboItem.ItemData(229) = 17229
    cboItem.AddItem "Sprint Shoes", 230
    cboItem.ItemData(230) = 17230
    cboItem.AddItem "Rename Card", 231
    cboItem.ItemData(231) = 1231
    cboItem.AddItem "Tonic", 232
    cboItem.ItemData(232) = 1232
    cboItem.AddItem "Potion", 233
    cboItem.ItemData(233) = 1233
    cboItem.AddItem "X-Potion", 234
    cboItem.ItemData(234) = 1234
    cboItem.AddItem "Tincture", 235
    cboItem.ItemData(235) = 1235
    cboItem.AddItem "Ether", 236
    cboItem.ItemData(236) = 1236
    cboItem.AddItem "X-Ether", 237
    cboItem.ItemData(237) = 1237
    cboItem.AddItem "Elixir", 238
    cboItem.ItemData(238) = 1238
    cboItem.AddItem "Megalixir", 239
    cboItem.ItemData(239) = 1239
    cboItem.AddItem "Phoenix Down", 240
    cboItem.ItemData(240) = 1240
    cboItem.AddItem "Revivify", 241
    cboItem.ItemData(241) = 1241
    cboItem.AddItem "Antidote", 242
    cboItem.ItemData(242) = 1242
    cboItem.AddItem "Eyedrop", 243
    cboItem.ItemData(243) = 1243
    cboItem.AddItem "Soft", 244
    cboItem.ItemData(244) = 1244
    cboItem.AddItem "Remedy", 245
    cboItem.ItemData(245) = 1245
    cboItem.AddItem "Sleeping Bag", 246
    cboItem.ItemData(246) = 1246
    cboItem.AddItem "Tent", 247
    cboItem.ItemData(247) = 1247
    cboItem.AddItem "Green Cherry", 248
    cboItem.ItemData(248) = 1248
    cboItem.AddItem "Magicite", 249
    cboItem.ItemData(249) = 1249
    cboItem.AddItem "Super Ball", 250
    cboItem.ItemData(250) = 1250
    cboItem.AddItem "Echo Screen", 251
    cboItem.ItemData(251) = 1251
    cboItem.AddItem "Smoke Bomb", 252
    cboItem.ItemData(252) = 1252
    cboItem.AddItem "Warp Stone", 253
    cboItem.ItemData(253) = 1253
    cboItem.AddItem "Dried Meat", 254
    cboItem.ItemData(254) = 1254
    cboItem.AddItem " ", 255
    cboItem.ItemData(255) = -1
    
    ' Ajustando as propriedades Max e Min dos controles
    ' UpDown. Alguns variam e por isso não estão aqui.
    updCampo(0).Max = 99        ' level
    updCampo(0).Min = 1
    updCampo(1).Max = 1677215   ' esperiência
    updCampo(1).Min = 1
    updCampo(3).Max = 9999      ' HP máximo
    updCampo(3).Min = 1
    updCampo(5).Max = 999       ' MP máximo
    updCampo(5).Min = 1
    updCampo(6).Max = 255       ' vigor
    updCampo(6).Min = 0
    updCampo(7).Max = 255       ' speed
    updCampo(7).Min = 0
    updCampo(8).Max = 255       ' stamina
    updCampo(8).Min = 0
    updCampo(9).Max = 255       ' magic
    updCampo(9).Min = 0
    updCampo(10).Min = 0        ' GP
    updCampo(10).Max = 1677215
    updCampo(11).Min = 0        ' steps
    updCampo(11).Max = 1677215
    updCampo(12).Max = 255      ' itens
    updCampo(12).Min = 0
    updCampo(12).Max = 99       ' aprendizado das magias
    updCampo(12).Min = 1
    
    grdItem.Col = 0             ' ajustando tamanho e
    grdItem.Row = 0             ' algumas propriedades
    grdItem.Text = "Nome"       ' do grid de itens
    grdItem.Col = 1
    grdItem.Text = "Qtd"
    grdItem.RowData(1) = -1
    grdItem.ColWidth(0) = 1515
    grdItem.ColWidth(1) = 493
End Sub

Public Sub travarControles(liberado As Boolean)
    Dim I As Integer
    If (liberado) Then
        mnuArquivoFechar.Enabled = True
        mnuArquivoSalvar.Enabled = True
    Else
        mnuArquivoFechar.Enabled = False
        mnuArquivoSalvar.Enabled = False
    End If
    
    For I = 0 To 14
        txtCampo(I).Enabled = liberado
    Next
End Sub

Public Sub ajustaMenuReabrir()
    Dim I As Integer
    For I = 0 To 3
        If (Trim(config.Arq(I)) <> "") Then
            mnuArquivoReabrirArq(I).Caption = config.Arq(I)
            mnuArquivoReabrirArq(I).Tag = config.Arq(I)
        Else
            mnuArquivoReabrirArq(I).Caption = "< vazio >"
            mnuArquivoReabrirArq(I).Tag = ""
        End If
    Next
End Sub

Public Sub gravaArquivosRecentes(arquivo As String)
    Dim I, J As Integer
    If (Trim(arquivo) <> "") Then
        I = 0
        Do While (config.Arq(I) <> arquivo And I < 4)
            I = I + 1
        Loop
        If (I > 0) Then
            For J = I To 1 Step -1
                config.Arq(J) = config.Arq(J - 1)
            Next
            config.Arq(0) = arquivo
            WriteINIFile "ArquivosRecentes", "Arq1", config.Arq(0), App.Path & "\" & ARQUIVO_INI
            WriteINIFile "ArquivosRecentes", "Arq2", config.Arq(1), App.Path & "\" & ARQUIVO_INI
            WriteINIFile "ArquivosRecentes", "Arq3", config.Arq(2), App.Path & "\" & ARQUIVO_INI
            WriteINIFile "ArquivosRecentes", "Arq4", config.Arq(3), App.Path & "\" & ARQUIVO_INI
        End If
    End If
End Sub

Private Sub cboHeroi_Click()
    On Error Resume Next
    Dim I, bitTemp As Integer
    If (arquivoAberto) Then
        ' preenche os txtCampo's de acordo com
        ' o personagem escolhido na combo box.
        txtCampo(0).Text = ""
        For I = 0 To 5
            txtCampo(0).Text = txtCampo(0).Text & converteDeLetra(mapaMemoria.personagens(cboHeroi.ListIndex).nome.Caracter(I).valor - CARACTER_NOME)
        Next
    
        txtCampo(1).Text = mapaMemoria.personagens(cboHeroi.ListIndex).level.valor
    
        txtCampo(2).Text = mapaMemoria.personagens(cboHeroi.ListIndex).experiencia(0).valor + (CLng(mapaMemoria.personagens(cboHeroi.ListIndex).experiencia(1).valor) * 256) + (CLng(mapaMemoria.personagens(cboHeroi.ListIndex).experiencia(2).valor) * 65536)
    
        txtCampo(4).Text = mapaMemoria.personagens(cboHeroi.ListIndex).HPmax(0).valor + CLng(mapaMemoria.personagens(cboHeroi.ListIndex).HPmax(1).valor * 256)
    
        updCampo(2).Max = CInt(txtCampo(4).Text)
        txtCampo(3).Text = mapaMemoria.personagens(cboHeroi.ListIndex).HP(0).valor + CLng(mapaMemoria.personagens(cboHeroi.ListIndex).HP(1).valor * 256)
        
        txtCampo(6).Text = mapaMemoria.personagens(cboHeroi.ListIndex).MPmax(0).valor + CLng(mapaMemoria.personagens(cboHeroi.ListIndex).MPmax(1).valor * 256)
    
        updCampo(4).Max = CInt(txtCampo(6).Text)
        txtCampo(5).Text = mapaMemoria.personagens(cboHeroi.ListIndex).MP(0).valor + CLng(mapaMemoria.personagens(cboHeroi.ListIndex).MP(1).valor * 256)
    
        txtCampo(7).Text = mapaMemoria.personagens(cboHeroi.ListIndex).vigor.valor
    
        txtCampo(8).Text = mapaMemoria.personagens(cboHeroi.ListIndex).speed.valor
    
        txtCampo(9).Text = mapaMemoria.personagens(cboHeroi.ListIndex).stamina.valor
    
        txtCampo(10).Text = mapaMemoria.personagens(cboHeroi.ListIndex).magic.valor
    
        'comandos, equipamentos e relics
        For I = 0 To 3
            cboComando(I).Text = cboComando(I).List(mapaMemoria.personagens(cboHeroi.ListIndex).comando(I).valor)
        Next
    
        If (mapaMemoria.personagens(cboHeroi.ListIndex).maoEsq.valor <= 104) Then
            cboEquip(0).Text = cboEquip(0).List(mapaMemoria.personagens(cboHeroi.ListIndex).maoEsq.valor)
        Else
            cboEquip(0).Text = cboEquip(0).List(105)
        End If
        If (mapaMemoria.personagens(cboHeroi.ListIndex).maoDir.valor <= 104) Then
            cboEquip(1).Text = cboEquip(1).List(mapaMemoria.personagens(cboHeroi.ListIndex).maoDir.valor)
        Else
            cboEquip(1).Text = cboEquip(1).List(105)
        End If
        If (mapaMemoria.personagens(cboHeroi.ListIndex).cabeca.valor - OFFSET_ELMO <= 26) Then
            cboEquip(2).Text = cboEquip(2).List(mapaMemoria.personagens(cboHeroi.ListIndex).cabeca.valor - OFFSET_ELMO)
        Else
            cboEquip(2).Text = cboEquip(2).List(27)
        End If
        If (mapaMemoria.personagens(cboHeroi.ListIndex).corpo.valor - OFFSET_ARMADURA <= 30) Then
            cboEquip(3).Text = cboEquip(3).List(mapaMemoria.personagens(cboHeroi.ListIndex).corpo.valor - OFFSET_ARMADURA)
        Else
            cboEquip(3).Text = cboEquip(3).List(31)
        End If
        If (mapaMemoria.personagens(cboHeroi.ListIndex).relic1.valor - OFFSET_RELIC <= 54) Then
            cboEquip(4).Text = cboEquip(4).List(mapaMemoria.personagens(cboHeroi.ListIndex).relic1.valor - OFFSET_RELIC)
        Else
            cboEquip(4).Text = cboEquip(4).List(55)
        End If
        If (mapaMemoria.personagens(cboHeroi.ListIndex).relic2.valor - OFFSET_RELIC <= 54) Then
            cboEquip(5).Text = cboEquip(5).List(mapaMemoria.personagens(cboHeroi.ListIndex).relic2.valor - OFFSET_RELIC)
        Else
            cboEquip(5).Text = cboEquip(5).List(55)
        End If
        If (mapaMemoria.personagens(cboHeroi.ListIndex).esper.valor <= 26) Then
            cboEquip(6).Text = cboEquip(6).List(mapaMemoria.personagens(cboHeroi.ListIndex).esper.valor)
        Else
            cboEquip(6).Text = cboEquip(6).List(27)
        End If
        
        bitTemp = mapaMemoria.personagens(cboHeroi.ListIndex).status.valor
        For I = 1 To 8
            If (bitTemp / (256 / I) = 1) Then
                chkStatus(8 - I).Value = 1
            Else
                chkStatus(8 - I).Value = 0
            End If
            bitTemp = bitTemp Mod 2
        Next
        If (mapaMemoria.personagens(cboHeroi.ListIndex).float.valor > (&H7F)) Then
            chkStatus(8).Value = 1
        Else
            chkStatus(8).Value = 0
        End If
    End If
End Sub

Private Sub cboMagias_Click()
    Dim nivelAprendizado As Byte
    If (arquivoAberto) Then
        If (cboHeroiMagia.Text <> "") Then
            nivelAprendizado = mapaMemoria.personagens(cboHeroiMagia.ListIndex).magias(cboMagias.ListIndex + 1).valor
            If (nivelAprendizado > 99) Then
                ' se a magia está aprendida (100%)
                optAprendizado(0).Value = False
                optAprendizado(1).Value = True
                optAprendizado(2).Value = False
                txtCampo(13).Text = ""
                txtCampo(13).Enabled = False
                updCampo(12).Enabled = False
            Else
                If (nivelAprendizado = 0) Then
                    ' se a magia nem começou a ser aprendida (0%)
                    optAprendizado(0).Value = True
                    optAprendizado(1).Value = False
                    optAprendizado(2).Value = False
                    txtCampo(13).Text = ""
                    txtCampo(13).Enabled = False
                    updCampo(12).Enabled = False
                Else
                    ' magia sendo aprendida
                    optAprendizado(0).Value = False
                    optAprendizado(1).Value = False
                    optAprendizado(2).Value = True
                    txtCampo(13).Text = nivelAprendizado
                    txtCampo(13).Enabled = True
                    updCampo(12).Enabled = True
                End If
            End If
        End If
    End If
End Sub
Private Sub deletaItens()
    If (grdItem.Rows = 2) Then
        ' caso especial onde o MSFlexGrid não permite
        ' que não exista pelo menos uma linha.
        mapaMemoria.itens.tipoItem(0).valor = 255
        mapaMemoria.itens.ctrlItem(0) = -1
        grdItem.Row = 1
        grdItem.Col = 0
        grdItem.Text = " "
        grdItem.Col = 1
        grdItem.Text = " "
        grdItem.RowData(1) = -1
    Else
        mapaMemoria.itens.tipoItem(grdItem.Row - 1).valor = 255
        mapaMemoria.itens.ctrlItem(grdItem.Row - 1) = -1
        grdItem.RemoveItem (grdItem.Row)
        cboItem.Text = cboItem.List(255)
        txtCampo(14).Text = ""
    End If
End Sub

Private Sub cmdDelItem_Click()
    deletaItens
End Sub

Private Sub preencheItens()
    Dim I As Integer
    Dim nomeItem As String, qtdItem As String
    For I = 0 To 254
        ' 255 (&HFF) representa a não existência de itens
        If (mapaMemoria.itens.tipoItem(I).valor <> 255) Then
            nomeItem = cboItem.List(mapaMemoria.itens.tipoItem(I).valor)
            qtdItem = CInt(mapaMemoria.itens.qtdItem(I).valor)
            grdItem.AddItem nomeItem & Chr(9) & qtdItem
            ' guarda um indicador do item para que ele
            ' possa ser achado mais rapidamente durante
            ' a edição da quantidade.
            grdItem.RowData(grdItem.Rows - 1) = cboItem.ItemData(mapaMemoria.itens.tipoItem(I).valor)
        End If
    Next
    ' remove aquela linha que vem
    ' inicialmente com o grid
    If (grdItem.Rows > 2) Then
        grdItem.RemoveItem 1
    End If
End Sub

Private Sub preencheEspeciais()
    Dim bitTemp, I As Integer, erro As Boolean
    checkboxByte mapaMemoria.especiais.blitz, chkBlitz, 0, 8
    checkboxByte mapaMemoria.especiais.swdTech, chkSwdtech, 0, 8
    checkboxByte mapaMemoria.especiais.dance, chkDance, 0, 8
    For I = 0 To 31
        listboxByte mapaMemoria.especiais.rage(I), lstRage, 8 * I
    Next
    For I = 0 To 2
        listboxByte mapaMemoria.especiais.lore(I), lstLore, 8 * I
    Next
End Sub

Private Sub grdItem_Click()
    'só para debug
    lblRegistro.Caption = grdItem.ColWidth(0) & " " & grdItem.ColWidth(1) & " " & grdItem.RowData(grdItem.Row) & " " & grdItem.Rows & " " & grdItem.Row
    
    grdItem.Col = 1
    If (grdItem.RowData(grdItem.Row) <> -1) Then
        txtCampo(14).Text = grdItem.Text
        cboItem.Text = cboItem.List(grdItem.RowData(grdItem.Row) Mod 1000)
    End If
End Sub

Private Sub preencheStepGil()
    txtCampo(11).Text = mapaMemoria.gil(2).valor + (CLng(mapaMemoria.gil(1).valor) * 256) + (CLng(mapaMemoria.gil(0).valor) * 65536)
    txtCampo(12).Text = mapaMemoria.steps(2).valor + (CLng(mapaMemoria.steps(1).valor) * 256) + (CLng(mapaMemoria.steps(0).valor) * 65536)
End Sub

Private Sub preencheEsper()
    Dim I As Integer
    For I = 0 To 3
        checkboxByte mapaMemoria.espers(I), chkEspers, 8 * I, 26
    Next
End Sub

Private Sub mnuArquivoAbrir_Click()
    On Error Resume Next
    cdlArquivo.DialogTitle = NOME_APP
    cdlArquivo.Filter = "ZSNES's save stats (*.zst;*.zs1;*.zs2;*.zs3;*.zs4;*.zs5;*.zs6;*.zs7;*.zs8;*.zs9)|*.zst;*.zs1;*.zs2;*.zs3;*.zs4;*.zs5;*.zs6;*.zs7;*.zs8;*.zs9"
    cdlArquivo.Flags = &H1000& Or &H8& Or &H800& Or &H4&
    cdlArquivo.InitDir = config.Path
    cdlArquivo.FileName = ""
    cdlArquivo.ShowOpen
    'cuida do erro gerado qdo é dado
    'Cancel em uma Common Dialog Box
    If (Err.Number <> 32755) Then
        If (cdlArquivo.FileName <> "") Then
            SSTab1.Tab = 0
            preencheMapaMemoria (cdlArquivo.FileName)
            preencheEspeciais
            preencheItens
            preencheStepGil
            preencheEsper
            gravaArquivosRecentes (cdlArquivo.FileName)
            ajustaMenuReabrir
            arquivoAberto = True
            arquivoAlterado = False
            travarControles (arquivoAberto)
        End If
    End If
End Sub

Private Sub mnuArquivoFechar_Click()
    arquivoAberto = False
    grdItem.Clear
    travarControles (arquivoAberto)
End Sub

Private Sub mnuArquivoReabrirArq_Click(Index As Integer)
    If (mnuArquivoReabrirArq(Index).Tag <> "") Then
        SSTab1.Tab = 0
        preencheMapaMemoria (mnuArquivoReabrirArq(Index).Tag)
        preencheEspeciais
        preencheItens
        preencheStepGil
        preencheEsper
        gravaArquivosRecentes (mnuArquivoReabrirArq(Index).Tag)
        ajustaMenuReabrir
        arquivoAberto = True
        arquivoAlterado = False
        travarControles (arquivoAberto)
    End If
End Sub

Private Sub mnuArquivoSair_Click()
    End
End Sub

Private Sub alteraCampo(Index As Integer)
    Select Case (Index)
        Case 14
            If (txtCampo(14).Text <> "" And grdItem.RowData(1) <> -1) Then
                grdItem.Text = txtCampo(14).Text
            End If
    End Select
End Sub

Private Sub mnuConfig_Click()
    frmConfig.Show 1
End Sub

Private Sub txtCampo_Change(Index As Integer)
    alteraCampo (Index)
End Sub

Private Sub txtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
    ' somente a TextBox de índice 0 pode conter letras
    If (Index <> 0) Then
        If KeyAscii < 48 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub reorganizaItens()
    Dim I, J, somaQtd As Integer
    Dim temp1, temp2 As Byte
    Dim temp3 As Integer
    ' é executado um "bubblesort" para
    ' ordenar os itens como o jogo faz
    For I = 1 To grdItem.Rows - 1
        prbItem.Value = (I / grdItem.Rows) * 100
        For J = I + 1 To grdItem.Rows - 1
            If (grdItem.RowData(I) > grdItem.RowData(J)) Then
                temp1 = grdItem.TextMatrix(I, 0)
                temp2 = grdItem.TextMatrix(I, 1)
                temp3 = grdItem.RowData(I)
                grdItem.TextMatrix(I, 0) = grdItem.TextMatrix(J, 0)
                grdItem.TextMatrix(I, 1) = grdItem.TextMatrix(J, 1)
                grdItem.RowData(I) = grdItem.RowData(J)
                grdItem.TextMatrix(J, 0) = temp1
                grdItem.TextMatrix(J, 1) = temp2
                grdItem.RowData(J) = temp3
            End If
        Next
    Next
    prbItem.Value = prbItem.Max
    MsgBox "Reorganização concluída.", vbOKOnly, NOME_APP
    prbItem.Value = prbItem.Min
End Sub

Private Sub adicionaItens()
    Dim nomeItem As String
    Dim I As Integer
    Dim achou As Boolean
    If (grdItem.Rows > 254) Then
        MsgBox "Não é possível adicionat mais itens.", vbOKOnly, NOME_APP
    Else
        
        If (Trim(cboItem.Text) = "") Then
            MsgBox "Selecione um item.", vbOKOnly, NOME_APP
            cboItem.SetFocus
        Else
            ' procura pelo item que se
            ' quer adicionar pelo grid
            I = 1
            achou = False
            Do While (I < grdItem.Rows And achou = False)
                If (grdItem.RowData(I) <> cboItem.ItemData(cboItem.ListIndex)) Then
                    I = I + 1
                Else
                    achou = True
                End If
            Loop
            ' se I for igual ao número de linhas, quer
            ' dizer que o item não foi achado no grid
            If (I = grdItem.Rows) Then
                nomeItem = cboItem.List(cboItem.ListIndex)
                grdItem.AddItem nomeItem & Chr(9) & "1"
            
                ' guarda um indicador do item para que ele
                ' possa ser achado mais rapidamente durante
                ' a edição da quantidade.
                grdItem.RowData(grdItem.Rows - 1) = cboItem.ItemData(cboItem.ListIndex)
                mapaMemoria.itens.ctrlItem(grdItem.Rows - 2) = cboItem.ItemData(cboItem.ListIndex)
                mapaMemoria.itens.tipoItem(grdItem.Rows - 2).valor = cboItem.ListIndex
                mapaMemoria.itens.qtdItem(grdItem.Rows - 2).valor = 1
                grdItem.Row = grdItem.Rows - 1
                ' caso especial onde o MSFlexGrid não permite
                ' que não exista pelo menos uma linha.
                If (grdItem.RowData(1) = -1) Then
                    grdItem.RemoveItem (1)
                End If
                grdItem.TopRow = grdItem.Rows - 1
            Else
                MsgBox "O item " & cboItem.List(cboItem.ListIndex) & " já consta na lista.", vbOKOnly, NOME_APP
            End If
        End If
    End If
    
End Sub

Private Sub cmdAddItem_Click()
    adicionaItens
End Sub

Private Sub cmdOrderItem_Click()
    reorganizaItens
End Sub
