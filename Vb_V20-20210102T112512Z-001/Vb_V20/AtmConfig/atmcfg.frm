VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form atmcfg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ATM Config 2.0"
   ClientHeight    =   6495
   ClientLeft      =   1275
   ClientTop       =   2010
   ClientWidth     =   9630
   Icon            =   "atmcfg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9630
   Begin VB.CommandButton C_Modify 
      Caption         =   "修 改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton C_Cancel 
      Caption         =   "取 消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7800
      TabIndex        =   1
      Top             =   5760
      Width           =   1245
   End
   Begin VB.CommandButton C_OK 
      Caption         =   "确 定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4200
      MaskColor       =   &H00808080&
      TabIndex        =   0
      Top             =   5760
      Width           =   1230
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5175
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "应用信息设置"
      TabPicture(0)   =   "atmcfg.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(8)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(7)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "EVR"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "AudioCtrl_EndTime"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "AudioCtrl_StartTime"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "AudioCtrl"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "T_Telephone"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "T_BranchCode"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "T_BankCode"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "T_ATMCode"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "通讯＆ATM设置"
      TabPicture(1)   =   "atmcfg.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label32"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label31"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label30"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label29"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label28"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label7"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label19"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(12)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(4)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(3)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(2)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label1(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "T_HostPath"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "T_password"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "T_UserName"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "T_HostAddr"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "ReceiptType"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "OperatorInterface"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "T_MaxBills"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "T_ServerIP"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "T_ManagementPort"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "T_RecvTimeOut"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "T_ServerPort"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "其他信息设置"
      TabPicture(2)   =   "atmcfg.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "版本信息查询"
      TabPicture(3)   =   "atmcfg.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ICBCProVer"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label27"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label2(0)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label5"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label6"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label8"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label9"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label10"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label12"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label13"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label14"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label15"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Label16"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Line1"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "OptevaImgVer"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "MonVer"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "StarterVer"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "MenuVer"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "IdleVer"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "EndvistVer"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "PininputVer"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "CwdVer"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "PinchangeVer"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "InqVer"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "TfrVer"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "OperatorVer"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "Label17"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "OutsevVer"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "Label18"
      Tab(3).Control(28).Enabled=   0   'False
      Tab(3).ControlCount=   29
      Begin VB.Frame Frame6 
         Caption         =   "备份设置"
         Height          =   1065
         Left            =   -72600
         TabIndex        =   80
         Top             =   1740
         Width           =   4005
         Begin VB.DriveListBox BakDrive 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1710
            TabIndex        =   82
            Top             =   210
            Width           =   2175
         End
         Begin VB.TextBox lbl_BakNum 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   315
            Left            =   1710
            TabIndex        =   81
            Text            =   "3"
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "当前备份路径设置"
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   84
            Top             =   300
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "备份周期个数"
            Height          =   225
            Index           =   1
            Left            =   240
            TabIndex        =   83
            Top             =   690
            Width           =   1215
         End
      End
      Begin VB.TextBox T_ATMCode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   31
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox T_BankCode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         MaxLength       =   4
         TabIndex        =   30
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox T_BranchCode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   29
         Top             =   1200
         Width           =   1605
      End
      Begin VB.TextBox T_Telephone 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         MaxLength       =   20
         TabIndex        =   28
         Top             =   1200
         Width           =   1845
      End
      Begin VB.ComboBox AudioCtrl 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "atmcfg.frx":037A
         Left            =   1200
         List            =   "atmcfg.frx":037C
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox AudioCtrl_StartTime 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   26
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox AudioCtrl_EndTime 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   25
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox T_ServerPort 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -72600
         TabIndex        =   24
         Top             =   1320
         Width           =   1755
      End
      Begin VB.TextBox T_RecvTimeOut 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -72600
         TabIndex        =   23
         Top             =   1920
         Width           =   660
      End
      Begin VB.TextBox T_ManagementPort 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -72600
         TabIndex        =   22
         Top             =   2520
         Width           =   1725
      End
      Begin VB.TextBox T_ServerIP 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72600
         TabIndex        =   21
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox T_MaxBills 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -70920
         MaxLength       =   2
         TabIndex        =   20
         Top             =   3120
         Width           =   615
      End
      Begin VB.ComboBox OperatorInterface 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -70920
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3720
         Width           =   3495
      End
      Begin VB.ComboBox ReceiptType 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -70920
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   4320
         Width           =   2535
      End
      Begin VB.TextBox T_HostAddr 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68880
         TabIndex        =   17
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox T_UserName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68880
         TabIndex        =   16
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox T_password 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68880
         TabIndex        =   15
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox T_HostPath 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68880
         TabIndex        =   14
         Top             =   2520
         Width           =   3015
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   2160
         TabIndex        =   11
         Top             =   3240
         Width           =   5055
         Begin VB.CheckBox CHNPrj 
            Caption         =   "打印中文流水"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox ENGPrj 
            Caption         =   "打印英文流水"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   12
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   2160
         TabIndex        =   8
         Top             =   3840
         Width           =   5055
         Begin VB.CheckBox SWPIN 
            Caption         =   "软件加密"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox HWPIN 
            Caption         =   "硬件加密"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   2160
         TabIndex        =   5
         Top             =   4440
         Width           =   5055
         Begin VB.CheckBox SingleDES 
            Caption         =   "16位密钥"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox TriDES 
            Caption         =   "32位密钥"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   6
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CheckBox EVR 
         Caption         =   "连接网络数字硬盘录像机"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   2760
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ATM号(ATMCode):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   720
         TabIndex        =   79
         Top             =   720
         Width           =   1860
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "银行号(BankCode):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   4800
         TabIndex        =   78
         Top             =   720
         Width           =   2190
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "支行号(BranchCode):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   77
         Top             =   1320
         Width           =   2430
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "联系电话:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   5520
         TabIndex        =   76
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label11 
         Caption         =   "语音播放:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "播放起始时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   74
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "播放终止时间:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   73
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "服务器IP地址:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74400
         TabIndex        =   72
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   " 服务器端口:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   71
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "主机回应超时:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74640
         TabIndex        =   70
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "秒"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -71760
         TabIndex        =   69
         Top             =   2040
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "管理命令端口:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -74520
         TabIndex        =   68
         Top             =   2640
         Width           =   1530
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "最大出钞张数："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   -72840
         TabIndex        =   67
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label19 
         Caption         =   "操作员界面："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72480
         TabIndex        =   66
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "凭条格式："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72240
         TabIndex        =   65
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "取款(Cwd)模块："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69480
         TabIndex        =   64
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label OutsevVer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71400
         TabIndex        =   63
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "退出服务(OutOfService)模块："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   62
         Top             =   4560
         Width           =   3255
      End
      Begin VB.Label OperatorVer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71400
         TabIndex        =   61
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label TfrVer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67560
         TabIndex        =   60
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label InqVer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67560
         TabIndex        =   59
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label PinchangeVer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67560
         TabIndex        =   58
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label CwdVer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67560
         TabIndex        =   57
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label PininputVer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71400
         TabIndex        =   56
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label EndvistVer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71400
         TabIndex        =   55
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label IdleVer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67560
         TabIndex        =   54
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label MenuVer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67560
         TabIndex        =   53
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label StarterVer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71400
         TabIndex        =   52
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label MonVer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71400
         TabIndex        =   51
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label OptevaImgVer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70920
         TabIndex        =   50
         Top             =   720
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         BorderWidth     =   2
         X1              =   -74400
         X2              =   -66480
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "操作员(Operator)模块："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74040
         TabIndex        =   49
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "输密码(PinInput)模块："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74040
         TabIndex        =   48
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "主菜单(Menu)模块："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69720
         TabIndex        =   47
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "空闲等待(Idle)模块："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69960
         TabIndex        =   46
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "结束交易(EndVisit)模块："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   45
         Top             =   4080
         Width           =   2775
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "改密(PinChange)模块："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70080
         TabIndex        =   44
         Top             =   4560
         Width           =   2535
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "转帐(Transfer)模块："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69960
         TabIndex        =   43
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "查询(Inquiry)模块："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69840
         TabIndex        =   42
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "系统启动(Starter)模块："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74040
         TabIndex        =   41
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "系统监控(Monitor)模块："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74040
         TabIndex        =   40
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "OPTEVA镜像包版本号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -73680
         TabIndex        =   39
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "中行应用程序版本号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73560
         TabIndex        =   38
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label ICBCProVer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70920
         TabIndex        =   37
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Label Label28 
         Caption         =   "(有效值：1~40)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69960
         TabIndex        =   36
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "FTP主机地址:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70440
         TabIndex        =   35
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "FTP用户名:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   34
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Caption         =   "FTP用户密码:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70440
         TabIndex        =   33
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "FTP主机文件目录:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   32
         Top             =   2640
         Width           =   1935
      End
   End
End
Attribute VB_Name = "atmcfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==========================================================================================
'版权说明:  迪堡公司中国区技术部
'版本号：
'生成日期：2005.8
'作者：  汪林(初始版）
'模块功能： 配置ATM相关信息
'主要函数及其功能
' 全局变量
'修改日志
'-----------------------------------------------------------------------
'<时间>：[]
'<修改者>：
'<当前版本>：
'<详细记录>：
'   <内容>
'==========================================================================================
Private Const GlobalINIPath         As String = "C:\atmwosa\ini\"
Private Const AdvBrowserINIPath     As String = "C:\Windows\"
Private Const AppPath               As String = "C:\atmwosa\"
Private Const RulesINIPath          As String = "C:\AtmWosa\AtmConfig\Rules\"
Private Const sAgilisSDO            As String = "C:\WINNT\AgilisSDO.ini"
Private Const OPTEVAPlateformVer    As String = "SOFTWARE\Diebold"
Private Const MonitorFile           As String = "C:\Program Files\Diebold\Agilis Power\S3EFrame\S3EMonitor.exe"
Private Const StarterFile           As String = "C:\Program Files\Diebold\Agilis Power\S3EFrame\S3EStarter.exe"
Private Const CwdFile               As String = "C:\Atmwosa\Cwd.exe"
Private Const OptFile               As String = "C:\Atmwosa\OperatorDisp.exe"
Private Const EndvistFile           As String = "C:\Atmwosa\EndVisit.exe"
Private Const TfrFile               As String = "C:\Atmwosa\TsfrOut.exe"
Private Const IdleFile              As String = "C:\Atmwosa\Idle.exe"
Private Const MenuFile              As String = "C:\Atmwosa\Menu.exe"
Private Const PininputFile          As String = "C:\Atmwosa\PinInput.exe"
Private Const PinchangeFile         As String = "C:\Atmwosa\PinChange.exe"
Private Const OutserFile            As String = "C:\Atmwosa\OutOfService.exe"
Private Const InqFile               As String = "C:\Atmwosa\Inquiry.exe"

Dim ATMSysPath                      As String
Dim AudioCtrlType                   As String
Dim gIsModify                       As Integer
Dim OpInterFace                     As String
Dim FriendHintType                  As String

Private Sub AudioCtrl_Click()
    If AudioCtrl.ListIndex = 0 And gIsModify = 1 Then
        AudioCtrl_StartTime.Enabled = True
        AudioCtrl_EndTime.Enabled = True
    Else
        AudioCtrl_StartTime.Enabled = False
        AudioCtrl_EndTime.Enabled = False
    End If
End Sub

Private Sub C_Cancel_Click()
    Unload Me
End Sub

Private Sub C_Modify_Click()
    gIsModify = 1
    
    If SSTab.Tab = 0 Then
        AudioCtrl.Enabled = True
        AudioCtrl_EndTime.Enabled = True
        AudioCtrl_StartTime.Enabled = True
      
        T_ATMCode.Enabled = True
        T_BankCode.Enabled = True
        T_BranchCode.Enabled = True
        T_Telephone.Enabled = True
        ENGPrj.Enabled = True
        CHNPrj.Enabled = True
        SingleDES.Enabled = True
        TriDES.Enabled = True
        SWPIN.Enabled = True
        HWPIN.Enabled = True
        
        If AudioCtrlType = "T" Then
            AudioCtrl.ListIndex = 0
            AudioCtrl_StartTime.Enabled = True
            AudioCtrl_EndTime.Enabled = True
        ElseIf AudioCtrlType = "Y" Then
            AudioCtrl.ListIndex = 1
            AudioCtrl_StartTime.Enabled = False
            AudioCtrl_EndTime.Enabled = False
        Else
            AudioCtrl.ListIndex = 2
            AudioCtrl_StartTime.Enabled = False
            AudioCtrl_EndTime.Enabled = False
        End If
    ElseIf SSTab.Tab = 1 Then
        T_MaxBills.Enabled = True
        T_ManagementPort.Enabled = True
        T_RecvTimeOut.Enabled = True
        T_ServerIP.Enabled = True
        T_ServerPort.Enabled = True
        OperatorInterface.Enabled = True
        ReceiptType.Enabled = True
        T_HostAddr.Enabled = True
        T_UserName.Enabled = True
        T_password.Enabled = True
        T_HostPath.Enabled = True
    ElseIf SSTab.Tab = 2 Then
       BakDrive.Enabled = True
       lbl_BakNum.Enabled = True
    End If
    
End Sub

Private Sub C_OK_Click()
    Dim RetVal                  As Variant
    Dim hKey                    As Long
    Dim ret                     As Long
    Dim pStr                    As String
    Dim hKey1                   As Long
    Dim bytary()                As Byte
    Dim tmpStr                  As String
    Dim StartTime               As Variant
    Dim EndTime                 As Variant
    Dim IsIPModify              As Boolean
    Dim IsNameModify            As Boolean
    
    IsIPModify = False
    IsNameModify = False
    
    If gIsModify = 1 Then
        
        If AudioCtrl.ListIndex = 0 Then
            RetVal = SetIniS(GlobalINIPath + "global.ini", "AudioControl", "AudioConfig", "T")
            If IsDate(AudioCtrl_StartTime.Text) Then
                RetVal = SetIniS(GlobalINIPath + "global.ini", "AudioControl", "StartTime", _
                        AudioCtrl_StartTime.Text)
            End If
            If IsDate(AudioCtrl_EndTime.Text) Then
                RetVal = SetIniS(GlobalINIPath + "global.ini", "AudioControl", "EndTime", _
                        AudioCtrl_EndTime.Text)
            End If
        ElseIf AudioCtrl.ListIndex = 1 Then
          RetVal = SetIniS(GlobalINIPath + "global.ini", "AudioControl", "AudioConfig", "Y")
        Else
          RetVal = SetIniS(GlobalINIPath + "global.ini", "AudioControl", "AudioConfig", "N")
        End If
          
        If OperatorInterface.ListIndex = 0 Then
            RetVal = SetIniS(AdvBrowserINIPath + "AgilisAdvBrowser.ini", "Monitor", _
                    "FrontMonitorNbr ", "2")
            RetVal = SetIniS(AdvBrowserINIPath + "AgilisAdvBrowser.ini", "AdvBrowserMaint", _
                    "EmbedXfsKeys ", "0")

        Else
            RetVal = SetIniS(AdvBrowserINIPath + "AgilisAdvBrowser.ini", "Monitor", _
                    "FrontMonitorNbr ", "1")
            RetVal = SetIniS(AdvBrowserINIPath + "AgilisAdvBrowser.ini", "AdvBrowserMaint", _
                    "EmbedXfsKeys ", "1")
        End If
        
        '''''''''''''''''''''''''''
        'Save the SST_COM.INI
        RetVal = SetIniS("sst_com.ini", "SSTSRV", "PrimaryServer", T_ServerIP.Text)
        RetVal = SetIniS("sst_com.ini", "SSTSRV", "PrimaryPort", T_ServerPort.Text)
        RetVal = SetIniS("sst_com.ini", "SSTSRV", "RecvTimeOut", T_RecvTimeOut.Text)
        RetVal = SetIniS("sst_com.ini", "SST_SRV", "PrimaryPort", T_ManagementPort.Text)
        
        'Save the Global.INI
        RetVal = SetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "ATMCode", T_ATMCode.Text)
        
              
        RetVal = SetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "BankCode", T_BankCode.Text)
        RetVal = SetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "BranchCode", T_BranchCode.Text)
        RetVal = SetIniS(GlobalINIPath + "global.ini", "CustomerInfo", "Telephone", T_Telephone.Text)
        RetVal = SetIniS(GlobalINIPath + "FTP.ini", "FtpInfo", "HostAddr", T_HostAddr.Text)
        RetVal = SetIniS(GlobalINIPath + "FTP.ini", "FtpInfo", "UserName", T_UserName.Text)
        RetVal = SetIniS(GlobalINIPath + "FTP.ini", "FtpInfo", "Password", T_password.Text)
        RetVal = SetIniS(GlobalINIPath + "FTP.ini", "FtpInfo", "HostPath", T_HostPath.Text)
        
        If ReceiptType.ListIndex = 0 Then
            RetVal = SetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "Receipt_paper_type", "E")
        Else
            RetVal = SetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "Receipt_paper_type", "B")
        End If
        
        If Not IsNumeric(T_MaxBills.Text) Then
            T_MaxBills.Text = "30"
        ElseIf CInt(T_MaxBills.Text) < 1 Or CInt(T_MaxBills.Text) > 40 Then
            T_MaxBills.Text = "30"
        End If
        
        RetVal = SetIniS(GlobalINIPath + "global.ini", "Withdrawal", "MaxBills", T_MaxBills.Text)

        If ENGPrj.Value Then
            RetVal = SetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "PrjLanguage", "E")
        Else
            RetVal = SetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "PrjLanguage", "C")
        End If
    
        If SWPIN.Value Then
            RetVal = SetIniS(GlobalINIPath + "key.ini", "KeyList", "DESTYPE", "S")
        Else
            RetVal = SetIniS(GlobalINIPath + "key.ini", "KeyList", "DESTYPE", "H")
        End If
    
        If SingleDES.Value Then
            RetVal = SetIniS(GlobalINIPath + "key.ini", "KeyList", "DESMETHOD", "S")
        Else
            RetVal = SetIniS(GlobalINIPath + "key.ini", "KeyList", "DESMETHOD", "T")
        End If
        
        If EVR.Value Then
            RetVal = SetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "EVR", "Y")
        Else
            RetVal = SetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "EVR", "N")
        End If
        Call SaveRegInfo
    End If
    
    Unload Me
End Sub

Private Sub ENGPrj_Click()
    If ENGPrj.Value = 0 Then
        CHNPrj.Value = 1
    Else
        CHNPrj.Value = 0
    End If
End Sub

Private Sub CHNPrj_Click()
    If CHNPrj.Value = 0 Then
        ENGPrj.Value = 1
    Else
        ENGPrj.Value = 0
    End If
End Sub
Private Sub Form_Load()
   
    SSTab.Tab = 0
    gIsModify = 0
    
    AudioCtrl.AddItem "时间控制"
    AudioCtrl.AddItem "连续播放"
    AudioCtrl.AddItem "停止播放"

    OperatorInterface.AddItem "用后载维护屏幕（穿墙式）"
    OperatorInterface.AddItem "用客户交易屏幕（大堂式）"
    
    ReceiptType.AddItem "套打格式"
    ReceiptType.AddItem "空白凭条"
   
    Call ReadSstcomINI
    Call ReadGlobalINI
    Call ReadSysInfo
    Call ReadModuleVer
    Call ReadAdvBrowserINI
    Call ReadVersionINI
    Call GetLogRegInfo
End Sub

Private Sub ReadGlobalINI()
    Dim sStr                As String
    Dim AC_StartTime        As String
    Dim AC_EndTime          As String
    Dim RetVal              As Variant
    
    sStr = GetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "ATMCode", "")
    T_ATMCode.Text = sStr
  
   
    sStr = GetIniS(GlobalINIPath + "global.ini", "Withdrawal", "MaxBills", "")
    T_MaxBills.Text = sStr

    sStr = GetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "BankCode", "")
    T_BankCode.Text = sStr
    sStr = GetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "BranchCode", "")
    T_BranchCode.Text = sStr
    sStr = GetIniS(GlobalINIPath + "global.ini", "CustomerInfo", "Telephone", "")
    T_Telephone.Text = sStr
    sStr = GetIniS(GlobalINIPath + "FTP.ini", "FtpInfo", "HostAddr", "")
    T_HostAddr.Text = sStr
    sStr = GetIniS(GlobalINIPath + "FTP.ini", "FtpInfo", "UserName", "")
    T_UserName.Text = sStr
    sStr = GetIniS(GlobalINIPath + "FTP.ini", "FtpInfo", "Password", "")
    T_password.Text = sStr
    sStr = GetIniS(GlobalINIPath + "FTP.ini", "FtpInfo", "HostPath", "")
    T_HostPath.Text = sStr
    
    sStr = GetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "Receipt_paper_type", "")
    If sStr = "E" Then
        ReceiptType.ListIndex = 0
    Else
        ReceiptType.ListIndex = 1
    End If
    
    AudioCtrlType = GetIniS(GlobalINIPath + "global.ini", "AudioControl", "AudioConfig", "")
    AC_StartTime = GetIniS(GlobalINIPath + "global.ini", "AudioControl", "StartTime", "")
    AC_EndTime = GetIniS(GlobalINIPath + "global.ini", "AudioControl", "EndTime", "")
    
    If AudioCtrlType = "T" Then
        AudioCtrl.ListIndex = 0
    ElseIf AudioCtrlType = "Y" Then
        AudioCtrl.ListIndex = 1
    Else
        AudioCtrl.ListIndex = 2
    End If
    
    AudioCtrl_StartTime.Text = AC_StartTime
    AudioCtrl_EndTime.Text = AC_EndTime
    
    sStr = GetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "PrjLanguage", "IsNull")
    If sStr = "IsNull" Then
        RetVal = SetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "PrjLanguage", "E")
        ENGPrj.Value = 1
        CHNPrj.Value = 0
    ElseIf sStr = "E" Then
        ENGPrj.Value = 1
        CHNPrj.Value = 0
    Else
        CHNPrj.Value = 1
        ENGPrj.Value = 0
    End If

    sStr = GetIniS(GlobalINIPath + "key.ini", "KeyList", "DESTYPE", "IsNull")
    If sStr = "IsNull" Then
        RetVal = SetIniS(GlobalINIPath + "key.ini", "KeyList", "DESTYPE", "S")
        SWPIN.Value = 1
        HWPIN.Value = 0
    ElseIf sStr = "S" Then
        SWPIN.Value = 1
        HWPIN.Value = 0
    Else
        SWPIN.Value = 0
        HWPIN.Value = 1
    End If

    sStr = GetIniS(GlobalINIPath + "key.ini", "KeyList", "DESMETHOD", "IsNull")
    If sStr = "IsNull" Then
        RetVal = SetIniS(GlobalINIPath + "key.ini", "KeyList", "DESMETHOD", "S")
        SingleDES.Value = 1
        TriDES.Value = 0
    ElseIf sStr = "S" Then
        SingleDES.Value = 1
        TriDES.Value = 0
    Else
        SingleDES.Value = 0
        TriDES.Value = 1
    End If

End Sub

Private Sub ReadVersionINI()
    
     ICBCProVer.Caption = GetIniS(GlobalINIPath + "version.ini", "Information", "Project", "")

End Sub

Private Sub ReadAdvBrowserINI()
    
    OpInterFace = GetIniS(AdvBrowserINIPath + "AgilisAdvBrowser.ini", "Monitor", "FrontMonitorNbr ", "")
    
    If OpInterFace = "2" Then
        OperatorInterface.ListIndex = 0
    Else
        OperatorInterface.ListIndex = 1
    End If
End Sub


Private Sub ReadSstcomINI()
    Dim sStr                    As String
    
    sStr = GetIniS("sst_com.ini", "SSTSRV", "PrimaryServer", "")
    T_ServerIP.Text = sStr
    sStr = GetIniS("sst_com.ini", "SSTSRV", "PrimaryPort", "")
    T_ServerPort.Text = sStr
    sStr = GetIniS("sst_com.ini", "SSTSRV", "RecvTimeOut", "")
    T_RecvTimeOut.Text = sStr
    sStr = GetIniS("sst_com.ini", "SST_SRV", "PrimaryPort", "")
    T_ManagementPort.Text = sStr
    
End Sub

Private Sub ReadSysInfo()
    Dim sStr                    As String
    Dim LocalServiceName        As String
    Dim IsDHCP                  As Long
    
    OptevaImgVer.Caption = GetRegKeyS(HKEY_LOCAL_MACHINE, OPTEVAPlateformVer, "OptevaImage", 20, "")
End Sub

Private Sub ReadModuleVer()
    Dim fso                     As New FileSystemObject
    
    If fso.FileExists(MonitorFile) Then
        MonVer.Caption = fso.GetFileVersion(MonitorFile)
    End If
    
     If fso.FileExists(StarterFile) Then
        StarterVer.Caption = fso.GetFileVersion(StarterFile)
    End If
    
    If fso.FileExists(CwdFile) Then
        CwdVer.Caption = fso.GetFileVersion(CwdFile)
    End If
    
    If fso.FileExists(OptFile) Then
        OperatorVer.Caption = fso.GetFileVersion(OptFile)
    End If
   
    If fso.FileExists(EndvistFile) Then
        EndvistVer.Caption = fso.GetFileVersion(EndvistFile)
    End If

    If fso.FileExists(TfrFile) Then
        TfrVer.Caption = fso.GetFileVersion(TfrFile)
    End If
    
    If fso.FileExists(IdleFile) Then
        IdleVer.Caption = fso.GetFileVersion(IdleFile)
    End If
    
    If fso.FileExists(MenuFile) Then
        MenuVer.Caption = fso.GetFileVersion(MenuFile)
    End If
        
    If fso.FileExists(PininputFile) Then
        PininputVer.Caption = fso.GetFileVersion(PininputFile)
    End If
        
    If fso.FileExists(PinchangeFile) Then
        PinchangeVer.Caption = fso.GetFileVersion(PinchangeFile)
    End If
        
    If fso.FileExists(OutserFile) Then
        OutsevVer.Caption = fso.GetFileVersion(OutserFile)
    End If
        
    If fso.FileExists(InqFile) Then
        InqVer.Caption = fso.GetFileVersion(InqFile)
    End If
End Sub

Private Sub SingleDES_Click()
    If SingleDES.Value = 0 Then
        TriDES.Value = 1
    Else
        TriDES.Value = 0
    End If
End Sub

Private Sub TriDES_Click()
    If TriDES.Value = 0 Then
        SingleDES.Value = 1
    Else
        SingleDES.Value = 0
    End If
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
    If SSTab.Tab = 3 Then
        C_Modify.Enabled = False
    Else
        C_Modify.Enabled = True
    End If

End Sub

Private Sub SWPIN_Click()
    If SWPIN.Value = 0 Then
        HWPIN.Value = 1
    Else
        HWPIN.Value = 0
    End If
End Sub

Private Sub HWPIN_Click()
    If HWPIN.Value = 0 Then
        SWPIN.Value = 1
    Else
        SWPIN.Value = 0
    End If
End Sub
'======================= added by ly 2004-11-30 =========================================
'====== Please calling this function when the user clicked ok button======
Private Sub SaveRegInfo()
  Dim iRes            As Integer
  Dim hKey            As Long
  Dim iKeyIndex       As Integer
  Dim sSubKeyName     As String * 255
  Dim hSubKey         As Long
  Dim ayByte()        As Byte
  Dim sTmp            As String
  Dim sCurrDrive      As String
  Dim sBakPath        As String
  Dim lMaxArchives(1) As Long
  Dim objFileSys      As Object
  Dim objFolderDef    As Object
  Dim objSubFolders   As Object
  Dim objSubFolder    As Object
    
    
  On Error Resume Next
  sCurrDrive = Mid(BakDrive.Drive, 1, 2)
  Set objFileSys = CreateObject("Scripting.FileSystemObject")

  iRes = RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\SelfService\S3EArchive", hKey)
  If iRes = 0 Then
     iKeyIndex = 0
     Do
       iRes = RegEnumKey(hKey, iKeyIndex, sSubKeyName, 255)
       If iRes <> 0 Then
         Exit Do
       End If
       sTmp = TranslateStr(sSubKeyName, 255)
       iRes = RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\SelfService\S3EArchive\" + sTmp, hSubKey)
       If iRes = 0 Then
         iRes = GetValue(hSubKey, "ArchivePath", ayByte, REG_SZ)
         Call ByteArrayToString(ayByte, sTmp)
         sBakPath = Mid(sTmp, 3, Len(sTmp) - 2)
         sBakPath = sCurrDrive + sBakPath
         Call CreateFolders(sBakPath, objFileSys)
         iRes = SetValue(hSubKey, "ArchivePath", REG_SZ, sBakPath)
         iRes = RegSetValueEx(hSubKey, "MaxArchives", 0&, REG_DWORD, CLng(lbl_BakNum.Text), 4)
         Call RegCloseKey(hSubKey)
       End If
       iKeyIndex = iKeyIndex + 1
     Loop
     Call RegCloseKey(hKey)

  Else
    MsgBox "设置注册表信息出错！请联系我们"
  End If
End Sub
Private Sub CreateFolders(sPath As String, objFileSystem As Object)
  Dim iCurPos  As Integer
  Dim sCurPath As String
  
  On Error Resume Next
  iCurPos = 0
  Do
    iCurPos = InStr(iCurPos + 1, sPath, "\", vbTextCompare)
    If iCurPos = 0 Or iCurPos = Null Then
      Exit Do
    Else
      sCurPath = Mid(sPath, 1, iCurPos - 1)
      If Not objFileSystem.FolderExists(sCurPath) Then
        objFileSystem.CreateFolder (sCurPath)
      End If
    End If
  Loop
  If Not objFileSystem.FolderExists(sPath) Then
    objFileSystem.CreateFolder (sPath)
  End If
End Sub


Private Function TranslateStr(sInputData As String, iLen As Integer) As String
  Dim i     As Integer
  Dim sRes  As String
  
  sRes = ""
  For i = 1 To iLen
    If Asc(Mid$(sInputData, i, 1)) = 0 Then
       Exit For
    Else
       sRes = sRes & Mid$(sInputData, i, 1)
    End If
  Next
  TranslateStr = sRes
End Function

'=== Please call this function in the form load funtion
Private Sub GetLogRegInfo()
  Dim hKey      As Long
  Dim iRes      As Long
  Dim ayByte()  As Byte
  Dim sTmp      As String
  Dim iKeyIndex As Integer
  Dim sValue(1)  As Long
  On Error GoTo err_Handle
  iRes = RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\SelfService\S3EArchive\S3E", hKey)
  If iRes <> 0 Then
     MsgBox "读取S3EAdvBrowserLog注册表错误！"
  Else
    iRes = GetValue(hKey, "ArchivePath", ayByte, REG_SZ)
    Call ByteArrayToString(ayByte, sTmp)
    BakDrive.Drive = Mid(sTmp, 1, 2)
    If GetValueLong(hKey, "MaxArchives", sValue, REG_SZ) Then
      lbl_BakNum.Text = sValue(0)
    Else
      lbl_BakNum.Text = "3"
    End If
  End If
  Call RegCloseKey(hKey)
  Exit Sub
err_Handle:
  MsgBox "读取S3EAdvBrowserLog注册表错误,请联系我们"
End Sub

