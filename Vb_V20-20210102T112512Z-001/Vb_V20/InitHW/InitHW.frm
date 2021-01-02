VERSION 5.00
Object = "{80C55DB1-86F3-11D3-8B2F-00C04FF20A5D}#1.0#0"; "S3ELineInTcp.ocx"
Object = "{9C37E835-6A58-11D1-80C0-0020AF7093F9}#1.2#0"; "Dl.ocx"
Object = "{6C4DD4AB-27D5-11D3-96C4-000000000000}#1.0#0"; "S3ELineOutTcp.ocx"
Object = "{D659C2E4-44CC-11D3-ACF9-00105A5F6CAB}#1.0#0"; "boclmk1.ocx"
Object = "{EACE4ECF-6930-11D0-AC6C-0020AFBDD1D4}#1.0#0"; "SDOEdm.ocx"
Begin VB.Form FrmInitHW 
   Caption         =   "InitHW"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   Icon            =   "InitHW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdERI 
      Caption         =   "ERI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6800
      TabIndex        =   44
      Top             =   5325
      Width           =   1400
   End
   Begin VB.CommandButton CmdPAN 
      Caption         =   "PAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6800
      TabIndex        =   43
      Top             =   4680
      Width           =   1400
   End
   Begin VB.CommandButton CmdTEX 
      Caption         =   "TEX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3600
      TabIndex        =   42
      Top             =   6600
      Width           =   1400
   End
   Begin VB.CommandButton CmdAEX 
      Caption         =   "AEX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   41
      Top             =   5955
      Width           =   1400
   End
   Begin VB.CommandButton CmdRTT 
      Caption         =   "RTT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5200
      TabIndex        =   40
      Top             =   5955
      Width           =   1400
   End
   Begin VB.CommandButton CmdRDT 
      Caption         =   "RDT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5200
      TabIndex        =   39
      Top             =   5325
      Width           =   1400
   End
   Begin VB.CommandButton CmdCDP 
      Caption         =   "CDP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6840
      TabIndex        =   38
      Top             =   6000
      Width           =   1400
   End
   Begin VB.CommandButton CmdSetDL 
      Caption         =   "DLSet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   8400
      TabIndex        =   31
      Top             =   6600
      Width           =   1400
   End
   Begin VB.CommandButton CmdRWT 
      Caption         =   "RWT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5200
      TabIndex        =   30
      Top             =   4680
      Width           =   1400
   End
   Begin VB.CommandButton CmdTTI 
      Caption         =   "TTI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5200
      TabIndex        =   29
      Top             =   6600
      Width           =   1400
   End
   Begin VB.Frame Frame3 
      Caption         =   "Center"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2070
      Left            =   7350
      TabIndex        =   20
      Top             =   2325
      Width           =   2610
      Begin S3ELINEINLibCtl.S3ELineIn S3ELineIn 
         Height          =   375
         Left            =   360
         OleObjectBlob   =   "InitHW.frx":1272
         TabIndex        =   45
         Top             =   1320
         Width           =   735
      End
      Begin SDOEdmLibCtl.SDOEdm SDOEdm 
         Height          =   735
         Left            =   1320
         OleObjectBlob   =   "InitHW.frx":129C
         TabIndex        =   37
         Top             =   720
         Width           =   1095
      End
      Begin BOCLMK.BOCGDLMK BOCGDLMK 
         Height          =   615
         Left            =   1440
         TabIndex        =   36
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
      End
      Begin S3ELINEOUTLib.S3ELineOut S3ELineOut 
         Height          =   800
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1100
         _Version        =   65536
         _ExtentX        =   1940
         _ExtentY        =   1411
         _StockProps     =   1
      End
      Begin DLLib.DL Pcb3dl 
         Left            =   1320
         Top             =   1200
         _Version        =   65538
         _ExtentX        =   2143
         _ExtentY        =   1296
         _StockProps     =   0
      End
   End
   Begin VB.CommandButton CmdCWC 
      Caption         =   "CWC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2000
      TabIndex        =   14
      Top             =   5325
      Width           =   1400
   End
   Begin VB.Frame Frame2 
      Caption         =   "OutPutResult"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2070
      Left            =   240
      TabIndex        =   13
      Top             =   2325
      Width           =   6900
      Begin VB.TextBox TxtHostReturn 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   240
         Width           =   6650
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "InputFields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2070
      Left            =   240
      TabIndex        =   8
      Top             =   105
      Width           =   9730
      Begin VB.ComboBox CombTFRType 
         Height          =   315
         ItemData        =   "InitHW.frx":12CC
         Left            =   7350
         List            =   "InitHW.frx":12D6
         TabIndex        =   33
         Text            =   "Card"
         Top             =   1560
         Width           =   700
      End
      Begin VB.TextBox TxtTrack2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1100
         MaxLength       =   37
         TabIndex        =   28
         Text            =   "4563510800015012833=0606520100008630"
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox TxtTrack3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1100
         MaxLength       =   104
         TabIndex        =   26
         Top             =   1080
         Width           =   6835
      End
      Begin VB.TextBox TxtNewPin 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   3200
         MaxLength       =   6
         TabIndex        =   24
         Text            =   "654321"
         Top             =   180
         Width           =   1000
      End
      Begin VB.TextBox TxtTfr2Acc 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   5000
         MaxLength       =   19
         TabIndex        =   21
         Text            =   "4563510800015012817"
         Top             =   600
         Width           =   2430
      End
      Begin VB.TextBox TxtCWCNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   5200
         MaxLength       =   6
         TabIndex        =   17
         Top             =   1540
         Width           =   900
      End
      Begin VB.TextBox TxtAmount 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   5800
         MaxLength       =   19
         TabIndex        =   15
         Text            =   "100.00"
         Top             =   180
         Width           =   1275
      End
      Begin VB.TextBox TxtFitAcc 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1100
         MaxLength       =   19
         TabIndex        =   12
         Text            =   "4563510800015012833"
         Top             =   600
         Width           =   2430
      End
      Begin VB.TextBox TxtPin 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1100
         MaxLength       =   12
         TabIndex        =   10
         Text            =   "111111"
         Top             =   180
         Width           =   1000
      End
      Begin VB.Label LabTFR 
         Caption         =   "TFRType:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   32
         Top             =   1600
         Width           =   1095
      End
      Begin VB.Label LabTrack2 
         Caption         =   "Track2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   27
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label LabTrack3 
         Caption         =   "Track3:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   25
         Top             =   1120
         Width           =   855
      End
      Begin VB.Label LblNewPin 
         Caption         =   "NewPin:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2200
         TabIndex        =   23
         Top             =   270
         Width           =   855
      End
      Begin VB.Label LblTfr2Acc 
         Caption         =   "Tfr2AccNo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3700
         TabIndex        =   22
         Top             =   660
         Width           =   1125
      End
      Begin VB.Label LblCwcSerial 
         Caption         =   "ReverseNo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3900
         TabIndex        =   18
         Top             =   1600
         Width           =   1300
      End
      Begin VB.Label Label5 
         Caption         =   "GBLAmount:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4400
         TabIndex        =   16
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "AccNo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   195
         TabIndex        =   11
         Top             =   660
         Width           =   750
      End
      Begin VB.Label Label2 
         Caption         =   "Pin:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   210
         TabIndex        =   9
         Top             =   270
         Width           =   600
      End
   End
   Begin VB.CommandButton CmdCWD 
      Caption         =   "CWD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   380
      TabIndex        =   7
      Top             =   5325
      Width           =   1400
   End
   Begin VB.CommandButton CmdOEX 
      Caption         =   "OEX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3600
      TabIndex        =   6
      Top             =   5325
      Width           =   1400
   End
   Begin VB.CommandButton CmdINQ 
      Caption         =   "INQ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   380
      TabIndex        =   5
      Top             =   6600
      Width           =   1400
   End
   Begin VB.CommandButton CmdPIN 
      Caption         =   "PIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2000
      TabIndex        =   4
      Top             =   6600
      Width           =   1400
   End
   Begin VB.CommandButton CmdRQK 
      Caption         =   "RQK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3600
      TabIndex        =   3
      Top             =   4680
      Width           =   1400
   End
   Begin VB.CommandButton CmdTFR 
      Caption         =   "TFR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   380
      TabIndex        =   2
      Top             =   5970
      Width           =   1400
   End
   Begin VB.CommandButton CmdOpenLine 
      Caption         =   "LineOpen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2000
      TabIndex        =   1
      Top             =   4680
      Width           =   1400
   End
   Begin VB.CommandButton CmdInit 
      Caption         =   "DL_Init"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   380
      TabIndex        =   0
      Top             =   4680
      Width           =   1400
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000013&
      Height          =   2750
      Left            =   200
      TabIndex        =   19
      Top             =   4515
      Width           =   9760
   End
End
Attribute VB_Name = "FrmInitHW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

      Private Declare Sub keybd_event Lib "user32" ( _
         ByVal bVk As Byte, _
         ByVal bScan As Byte, _
         ByVal dwFlags As Long, _
         ByVal dwExtraInfo As Long)

Private Const sGlobalIni = "C:\ATMWosa\Ini\global.ini"
Const IniPath As String = "c:\atmwosa\ini\"
Const EnqRecordFile As String = "C:\ENQRecords.TXT"

Const DEVICE_NOT_INSTALL = "0"
Const DEVICE_ALL_OK = "1"
Const DEVICE_SIMPLE_FAULT = "2"
Const DEVICE_SEVERE_FAULT = "3"
Const DEVICE_CANNOT_DETECT_EXIST = "4"
Const DEVICE_CANNOT_DETECT_STATUS = "5"

Private Const sKeyIni = "C:\ATMWosa\Ini\Key.ini"
Const KEYEVENTF_KEYUP = &H2
Const VK_LWIN = &H5B

Private Type CardTypeRec
    TrackToMatch As Integer
    offset As Integer
    Length As Integer
    MatchChars As String
    MaxWithdrawAmount As Long
'    PinLength As Integer
'    PinMaxAttempts As Integer
    CardType As String
    AccNum_Track As Integer
    AccNum_Len As Integer
    AccNum_Offset As Integer
End Type

Private CardIdx() As CardTypeRec
Private AccnoMismatchB As String
Private AccnoMismatchL As String
Private g_sTrackInfo As String

Dim sLocalMK As String, sJulianDay As String, TermKey As String, NewPinKey As String, NewMACKey As String
Dim g_bDesMethod As Boolean
Private Rc As Integer
Private ReplyCode As Integer
Dim g_strAccNo19 As String
Dim nrc As Integer
Dim g_vHostRejectCode As Variant
Dim G_sDeviceStatus As String
Dim bStartLineIn As Boolean

Private g_ReadTrack(1 To 3) As String
'Dim g_eDevStatusOffset As DevStatus

Private Sub Form_Load()
    bStartLineIn = False
    CombTFRType.ListIndex = 0
    Call InitCardType  'initialize card type index
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bStartLineIn = True Then
        S3ELineIn.DoStopTest
    End If
End Sub

Sub InitCardType()
    Dim i As Integer, j As Integer
    Dim comma As Integer
    Dim StrTmp As String
    
    j = GetIniN(IniPath + "fit.ini", "General", "CurrentRecord", 0)
    If j < 1 Then
        MsgBox "Initcialize fit.ini Error!"
    Else
        ReDim CardIdx(j - 1)
        For i = 0 To j - 1
            CardIdx(i).TrackToMatch = GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(Str(i)), "TrackToMatch", 0)
            CardIdx(i).offset = 1 + GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(Str(i)), "Offset", 0)
            CardIdx(i).Length = GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(Str(i)), "Length", 0)
            CardIdx(i).MatchChars = GetIniS(IniPath + "fit.ini", "CardIndex" + LTrim(Str(i)), "MatchChars", "")
            CardIdx(i).MaxWithdrawAmount = 100 * GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(Str(i)), "MaxWithdrawAmount", 0)
'            CardIdx(i).PinLength = GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(Str(i)), "PinLength", 0)
'            CardIdx(i).PinMaxAttempts = GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(Str(i)), "PinMaxAttempts", 0)
            CardIdx(i).CardType = GetIniS(IniPath + "fit.ini", "CardIndex" + LTrim(Str(i)), "CardType", 0)
            
            CardIdx(i).AccNum_Len = GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(Str(i)), "AccNum_Len", 0)
            CardIdx(i).AccNum_Offset = GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(Str(i)), "AccNum_Offset", 0)
            CardIdx(i).AccNum_Track = GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(Str(i)), "AccNum_Track", 0)
        Next
        StrTmp = GetIniS(IniPath + "fit.ini", "CardMismatch", "CardAccountNumber", "")
        comma = InStr(1, StrTmp, ",")
        AccnoMismatchB = 1 + Val(Left(StrTmp, comma - 1))
        AccnoMismatchL = Val(Right(StrTmp, Len(StrTmp) - comma))
    End If
End Sub

Private Sub CmdSetDL_Click()

    TxtHostReturn.Text = ""
    Call SetInputFields
    
    TxtHostReturn.Text = "DataLink Set OK."
End Sub

Private Sub CmdInit_Click()
    Dim GBLAtmStatus As String
    Dim num As Integer
    Dim sAtmCode As String, sBranchCode As String
    
    TxtHostReturn.Text = ""
    
    sBranchCode = GetIniS(sGlobalIni, "Bank_Environment", "BranchCode", "6955")
    
    sAtmCode = GetIniS(sGlobalIni, "Bank_Environment", "ATMCode", "2088")
    
    G_sDeviceStatus = String(13, "0")
    Pcb3dl.DlSetCharRaw "GBLAtmStatus", "O"
    
    Pcb3dl.DlSetCharRaw "GBLBranchCode", sBranchCode
    Pcb3dl.DlSetCharRaw "GBLAtmCode", sAtmCode
    
    'for Fujitsu
'    Pcb3dl.DlSetCharRaw "GBLAtmCode", "2088"
    'for NCR
'    Pcb3dl.DlSetCharRaw "GBLAtmCode", "1001"

    GBLAtmStatus = Pcb3dl.DlGetCharRaw("GBLAtmStatus")
        
    nrc = Pcb3dl.DlSetCharRaw("CashBoxSts1", "0")
    nrc = Pcb3dl.DlSetCharRaw("CashBoxSts1", "0")
    nrc = Pcb3dl.DlSetCharRaw("CashBoxSts1", "0")
    nrc = Pcb3dl.DlSetCharRaw("CashBoxSts1", "0")
    
    nrc = Pcb3dl.DlSetCharRaw("DevicePRJState", "0")
    nrc = Pcb3dl.DlSetCharRaw("DevicePRRState", "0")
    nrc = Pcb3dl.DlSetCharRaw("DeviceCDMState", "0")
    nrc = Pcb3dl.DlSetCharRaw("GBLUseTriDES", "N")
    nrc = Pcb3dl.DlSetCharRaw("GBLAtmStatus", "C")
    nrc = Pcb3dl.DlSetCharRaw("FitCardType", "00")
'add for boc ShenZhen Test!"

    sAtmCode = Pcb3dl.DlGetCharRaw("GBLAtmCode")
    sBranchCode = Pcb3dl.DlGetCharRaw("GBLBranchCode")

    BOCGDLMK.ATMID = sAtmCode
    BOCGDLMK.BranCode = sBranchCode
    Call BOCGDLMK.MakeLMK
    sLocalMK = BOCGDLMK.LocalMK

'add end
    
'    nRc = Pcb3dl.DlReset("DenomOfCas1")
'    nRc = Pcb3dl.DlSetCharRaw("CasNotesPresent1", "0010")
'
'    nRc = Pcb3dl.DlReset("DenomOfCas2")
'    nRc = Pcb3dl.DlSetCharRaw("CasNotesPresent2", "1500")
'
'    nRc = Pcb3dl.DlReset("DenomOfCas3")
'    nRc = Pcb3dl.DlSetCharRaw("CasNotesPresent3", "0500")
'
'    nRc = Pcb3dl.DlSetCharRaw("DenomOfCas4", String(8, "0"))
'    nRc = Pcb3dl.DlSetCharRaw("CasNotesPresent4", "0000")
'
'    nRc = Pcb3dl.DlSetCharRaw("DenomOfCas5", String(8, "0"))
'    nRc = Pcb3dl.DlSetCharRaw("CasNotesPresent5", "0000")
'
'
'    nRc = Pcb3dl.DlSetCharRaw("ExceptionCode", "99")
'    Pcb3dl.DlSetCharRaw "GBLBranchCode", "2355"
'    Pcb3dl.DlSetCharRaw "GBLAtmCode", "A004"
'    nRc = Pcb3dl.DlReset("TotAccountNum")
'    num = 1
'    num = Pcb3dl.DlGetInt("TotAccountNum")
'    nRc = Pcb3dl.DlSetLong("TotAccountNum", num)
'    num = Pcb3dl.DlGetInt("TotAccountNum")
    'Set Pin
'    nrc = Pcb3dl.DlSetCharRaw("FitAccNo", "750001367035")
    nrc = Pcb3dl.DlSetCharRaw("FitAccNo", TxtFitAcc.Text)
    nrc = Pcb3dl.DlSetCharRaw("PinInputBlock", TxtPin.Text)
    
    Call SetInputFields
    
End Sub

Private Function SetInputFields() As String
'    Dim MasterKey As String
    Dim nAccLen As Integer
    Dim WthAmount As Variant
    Dim StrWthAmount As String
    Dim sValue As String


    nrc = Pcb3dl.DlSetCharRaw("GBLAtmStatus", "O")
    
    'Set FitAcc & FitTracks
    g_ReadTrack(2) = CStr(TxtTrack2.Text)
    g_ReadTrack(3) = CStr(TxtTrack3.Text)
    nAccLen = Len(g_ReadTrack(2))
    nAccLen = Len(g_ReadTrack(3))
    
    nrc = CheckCard()
    If nrc = 0 Then
       MsgBox "Track's contents is invalid!!"
       Exit Function
    ElseIf nrc = 2 Then
       MsgBox "The Card is Operator Card!!"
       Exit Function
    End If
    
End Function

'Return 0 -> Invalid card
'Return 1 -> Validated card
'Return 2 -> Operator card

Private Function CheckCard() As Integer
    Dim i As Integer
    Dim find As Boolean
    Dim bmatch As Boolean
    Dim ProcessTrack As Integer
    Dim AvailAccLenMatch As Boolean
    Dim StrTrack As String
    Dim StrmatchChars As String
'    Dim sCardType As String
    Dim EqualMarkPos As Integer
    Dim AvailTrackLen As Integer
'    Dim nReturnResult As Integer
    
    StrTrack = ""

    find = False
    ProcessTrack = 2
    If ValidateTrack(g_ReadTrack(2)) Then
        StrTrack = g_ReadTrack(2)
    ElseIf ValidateTrack(g_ReadTrack(3)) Then
        StrTrack = g_ReadTrack(3)
        ProcessTrack = 3
    Else
        CheckCard = 0
        Exit Function
    End If
    
'    If ProcessTrack = 2 Then
        EqualMarkPos = GetUnNumberMarkPos(1, StrTrack, 1, 30)
'    Else
'        EqualMarkPos = GetUnNumberMarkPos(3, StrTrack, 1, 30)
'    End If
    AvailTrackLen = EqualMarkPos - 1
    'Skip the check digit
'    AvailTrackLen = AvailTrackLen - 1
    If AvailTrackLen > 19 Then
        AvailTrackLen = 19
    ElseIf AvailTrackLen < 1 Then
        CheckCard = 0
        Exit Function
    End If
    
'            AvailAccLenMatch = (AvailTrackLen = CardIdx(i).AccNum_Len)
    
    If AvailTrackLen Then
        Call InitCardDL(StrTrack, i, ProcessTrack, AvailTrackLen)
'        If CardIdx(i).CardType = "OP" Then
'            CheckCard = 2
'        Else
            CheckCard = 1
'        End If
'        find = True
'        Exit For
    End If
    
'    For i = 0 To UBound(CardIdx)
''        StrTrack = g_ReadTrack(CardIdx(i).TrackToMatch)
''        If StrTrack = "" Then
''            StrTrack = g_ReadTrack(3)
''        End If
'        StrmatchChars = Mid(StrTrack, CardIdx(i).offset, CardIdx(i).Length)
'        bmatch = MatchChars_Compare(StrmatchChars, CardIdx(i).MatchChars, CardIdx(i).Length)
'        If bmatch = True Then
''            StrTrack = g_ReadTrack(CardIdx(i).AccNum_Track)
''            EqualMarkPos = InStr(CardIdx(i).AccNum_Offset + 1, StrTrack, "=", 1)
'            If ProcessTrack = 2 Then
'                EqualMarkPos = GetUnNumberMarkPos(1, StrTrack, 1, 30)
'            Else
'                EqualMarkPos = GetUnNumberMarkPos(3, StrTrack, 1, 30)
'            End If
'            AvailTrackLen = EqualMarkPos - 1
'            If AvailTrackLen > 19 Then
'                AvailTrackLen = 19
'            ElseIf AvailTrackLen < 1 Then
'                CheckCard = 0
'                Exit Function
'            End If
'
''            AvailAccLenMatch = (AvailTrackLen = CardIdx(i).AccNum_Len)
'
'            If AvailTrackLen Then
'                Call InitCardDL(StrTrack, i, ProcessTrack, AvailTrackLen)
'                If CardIdx(i).CardType = "OP" Then
'                    CheckCard = 2
'                Else
'                    CheckCard = 1
'                End If
'                find = True
'                Exit For
'            End If
'        End If
'    Next i
'
'    If find = False Then
'        i = UBound(CardIdx)
'        If CardIdx(i).MatchChars = "9999" Then
'            StrTrack = ""
'
'            If ValidateTrack(g_ReadTrack(2)) Then
'                StrTrack = g_ReadTrack(2)
'            ElseIf ValidateTrack(g_ReadTrack(3)) Then
'                StrTrack = g_ReadTrack(3)
'            End If
'
'            If StrTrack = "" Then
'               g_sTrackInfo = g_ReadTrack(3)
'               CheckCard = 0
'               find = False
'            Else
'                Call InitCardDL(StrTrack, i, ProcessTrack, AvailAccLenMatch)
'                CheckCard = 1
'                find = True
'            End If
'        End If
'    End If
'
'    If find = False Then
'        g_sTrackInfo = g_ReadTrack(3)
'        CheckCard = 0
'    End If
    
End Function

Private Function MatchChars_Compare(pChar1 As String, pChar2 As String, lenght As Integer) As Boolean
    Dim strOne1 As String
    Dim strOne2 As String
    Dim i As Integer
    MatchChars_Compare = True
    
    For i = 1 To lenght
      strOne1 = Mid(pChar1, i, 1)
      strOne2 = Mid(pChar2, i, 1)
      If strOne1 <> strOne2 And strOne2 <> "*" Then
        MatchChars_Compare = False
        Exit For
      End If
    Next
  
End Function
   
Private Function ValidateTrack(FeedTrack As String) As Boolean
    Dim EqualPoz As Integer
    Dim FindSeprator As Boolean
    Dim StrTrack1 As String
    Dim TrackLength As Integer
    Dim i As Integer
    Dim j As Integer
    
    If Trim(FeedTrack) = "" Then
        ValidateTrack = False
        Exit Function
    End If
    
'    EqualPoz = InStr(1, FeedTrack, "=", 1)
    TrackLength = Len(FeedTrack)
    EqualPoz = GetUnNumberMarkPos(1, FeedTrack, 1, TrackLength)
    FindSeprator = CBool(EqualPoz)
    If (Not FindSeprator) Then
        ValidateTrack = FindSeprator
        Exit Function
    End If
    
'    StrTrack1 = Mid(FeedTrack, 1, EqualPoz - 1)
'    FindSeprator = IsNumeric(StrTrack1)
'    If (Not FindSeprator) Then
'        ValidateTrack = FindSeprator
'        Exit Function
'    End If

    ValidateTrack = False
    For i = 1 To (EqualPoz - 1)
        StrTrack1 = Mid(FeedTrack, i, 1)
        If StrTrack1 <> "0" Then
            ValidateTrack = True
        End If
    Next i
    
End Function

Private Sub InitCardDL(ByVal sTrackHandle As String, ByVal row As Integer, ByVal ProcessTrack As Integer, ByVal AccNoLen As Integer)
    
    Dim i      As Integer
    Dim sAccNo As String
    Dim sMatch As String
    Dim sAccType As String
    Dim sAccTmpLen As Integer
    Dim ExpireDate As String
    Dim EqualMarkPos As Integer
    Dim sCardType As String
    Dim StaffCardFlag As String
    Dim CompanyCardFlag As String
    Dim sLocalBankCard As String
    
    Pcb3dl.DlSetCharRaw "FitTrack3Message", g_ReadTrack(3)
    Pcb3dl.DlSetCharRaw "FitTrack2Message", g_ReadTrack(2)

'    If CardIdx(row).CardType = "D" Then
'        Pcb3dl.DlSetCharRaw "FitCardType", "C"
'        Pcb3dl.DlSetCharRaw "GBLAccType", "C"
'    Else
'        Pcb3dl.DlSetCharRaw "FitCardType", CardIdx(row).CardType
'        Pcb3dl.DlSetCharRaw "GBLAccType", CardIdx(row).CardType
'    End If
    
'   For exist card type in FIT
    If ProcessTrack = 2 Then
        sAccNo = Mid(sTrackHandle, 1, AccNoLen)
        'get expire date
        EqualMarkPos = GetUnNumberMarkPos(1, sTrackHandle, 1, 25)
        ExpireDate = Mid(sTrackHandle, EqualMarkPos + 1, 4)
    Else
        sAccNo = Mid(sTrackHandle, 3, AccNoLen)
        i = Len(sTrackHandle)
        EqualMarkPos = GetUnNumberMarkPos(1, sTrackHandle, 2, i)
        EqualMarkPos = EqualMarkPos - 5
        If EqualMarkPos > 0 Then
            ExpireDate = Mid(sTrackHandle, EqualMarkPos, 4)
        Else
            ExpireDate = Format(Now(), "YYMM")
        End If
    End If

    Pcb3dl.DlSetCharRaw "FitExpDate", ExpireDate
    TxtFitAcc.Text = sAccNo
    
    'Pcb3dl.DlSetCharRaw "FitAccNo", sAccNo
    'Pcb3dl.DlSetCharRaw "FitPrrAccNo", Left(sAccNo, AccNoLen - 5) + "****" + Right(sAccNo, 1)

    'Get more information from the account no
'    StaffCardFlag = "0"
'    CompanyCardFlag = "0"
'    sCardType = "O"
'    sLocalBankCard = "N"
'    If CardIdx(row).CardType <> "OP" And CardIdx(row).CardType <> "O" Then
'        sMatch = Mid(sTrackHandle, 1, 2)
'        If sMatch = "63" Then
'            sAccType = Mid(sTrackHandle, 8, 1)
'            Select Case sAccType
'                Case "0"
'                    StaffCardFlag = "1"    '职工卡
'                    sCardType = "S"     '银卡
'                Case "4"
'                    CompanyCardFlag = "1"    '单位卡
'                    sCardType = "G"     '金卡
'                Case "5"
'                Case "6"
'                    sCardType = "G"
'                Case "7"
'                    CompanyCardFlag = "1"    '单位卡
'                    sCardType = "S"     '银卡
'                Case "8"
'                Case "9"
'                    sCardType = "S"     '银卡
'                Case Else
'                    sCardType = "S"     '银卡
'            End Select
'            sAccType = Mid(sTrackHandle, 9, 3)
'            If sAccType = "145" Then    'ISS_BANK_CODE == 145 , 顺德
'                sLocalBankCard = "Y"
'            End If
'        End If
'
'        sMatch = Mid(sTrackHandle, 1, 3)
'        If sMatch = "103" Then
'
'            sCardType = "J"     '借记卡
'            sAccType = Mid(sTrackHandle, 9, 2)
'            If sAccType = "10" Then
'                StaffCardFlag = "1"     '职工卡
'            End If
'
'            sAccType = Mid(sTrackHandle, 10, 1)
'            If sAccType = "3" Then
'                CompanyCardFlag = "1"    '单位卡
'            End If
'
'            sAccType = Mid(sTrackHandle, 4, 4)
'            If sAccType = "2355" Then    'BRANCH_CODE == 2355, 顺德
'                sLocalBankCard = "Y"
'            End If
'        End If
'
'        sMatch = Mid(sTrackHandle, 1, 4)
'        If sMatch = "9559" Then
'            sCardType = "J"     '借记卡
'            sAccType = Mid(sTrackHandle, 6, 1)
'            Select Case sAccType
'                Case "0"
'                    StaffCardFlag = "1"    '职工卡
'                Case "7"
'                    CompanyCardFlag = "1"    '单位卡
'            End Select
'            sAccType = Mid(sTrackHandle, 7, 3)
'            If sAccType = "145" Then    'ISS_BANK_CODE == 145 , 顺德
'                sLocalBankCard = "Y"
'            End If
'        End If
'
'        sMatch = Mid(sTrackHandle, 1, 5)
'        If sMatch = "53591" Or sMatch = "49102" Then
'            sAccType = Mid(sTrackHandle, 6, 1)
'            Select Case sAccType
'                Case "0"
'                    StaffCardFlag = "1"    '职工卡
'                    sCardType = "S"     '银卡
'                Case "4"
'                    CompanyCardFlag = "1"    '单位卡
'                    sCardType = "G"     '金卡
'                Case "5"
'                Case "6"
'                    sCardType = "G"
'                Case "7"
'                    CompanyCardFlag = "1"    '单位卡
'                    sCardType = "S"     '银卡
'                Case "8"
'                Case "9"
'                    sCardType = "S"
'                Case Else
'                    sCardType = "S"
'            End Select
'            sAccType = Mid(sTrackHandle, 7, 3)
'            If sAccType = "145" Then    'ISS_BANK_CODE == 145 , 顺德
'                sLocalBankCard = "Y"
'            End If
'        End If
'
'    End If
'
'    If CardIdx(row).CardType <> "OP" Then
'        Pcb3dl.DlSetCharRaw "StaffCardFlag", StaffCardFlag
'        Pcb3dl.DlSetCharRaw "CompanyCardFlag", CompanyCardFlag
'        Pcb3dl.DlSetCharRaw "MakeTotCardType", sCardType
'        Pcb3dl.DlSetCharRaw "LocalBankFlag", sLocalBankCard
'    End If
    
    
End Sub

Private Sub CmdOpenLine_Click()
    Dim sCurrentDate As String
    Dim sCurrentTime As String
    
    TxtHostReturn.Text = ""
    
'    S3ELineOut.Register (0)
    
    
    TxtHostReturn.Text = ""
    nrc = S3ELineOut.PuOpen
    If nrc <> 0 Then
        TxtHostReturn.Text = "PuOpen Error(" + CStr(nrc) + ")" + vbCrLf
    Else
        TxtHostReturn.Text = "PuOpen OK" + vbCrLf
    End If
    
    If bStartLineIn = False Then
        nrc = S3ELineIn.DoStartTest
        TxtHostReturn.Text = TxtHostReturn.Text + "S3ELineIn Open(" + CStr(nrc) + ")" + vbCrLf
        If nrc = 0 Then
            bStartLineIn = True
            S3ELineIn.BackColor = &HFF00&
        End If
    End If
    
    
End Sub


Private Sub S3ELineIn_HostMessageReceived(ByVal nType As Long)
    Select Case nType
        Case 300:     ' Host Open ATM
            TxtHostReturn.Text = "Host Message Open"
        Case 301:     ' Host Close
            TxtHostReturn.Text = "Host Message Close"
        Case 501:     ' Host Message Check Total
            TxtHostReturn.Text = "Host Message Check Total"
        Case 502:     ' Host Message Check BOX Status
             TxtHostReturn.Text = "Host Message Check BOX Status"
        Case 800:     ' Line Status is Down
            Pcb3dl.DlSetCharRaw "GBLLineStatus", "C"
            S3ELineOut.BackColor = &HFF&
            TxtHostReturn.Text = "Line Status is Down"
        Case 801:     ' Line Status is Active
            Pcb3dl.DlSetCharRaw "GBLLineStatus", "O"
            S3ELineOut.BackColor = &HFF00&
            TxtHostReturn.Text = "Line Status is Active"
    End Select
End Sub

Private Sub S3ELineOut_AtLineOpened(ByVal rcOpen As Integer)
    Dim RcvHostTime As String, RcvHostDate As String
    Dim sHostTranCode As String
    Dim sHostAtmCode As String
    Dim sTermMK As String, DecryptResult As String, DesResult As String, sCheckValue As String
    Dim vRecvData As Variant
    Dim sTermMKEnd As String
    Dim strNum As String
    Dim i As Integer
    
    If rcOpen = 0 Then
        S3ELineOut.BackColor = &HFF00&
        TxtHostReturn.Text = "PuOpen OK" + vbCrLf
        sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
        TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
        If sHostTranCode <> "DNP" And sHostTranCode <> "ANP" Then    'accept
            TxtHostReturn.Text = TxtHostReturn.Text + "INT IS NOT ANP"
            MsgBox ("Host Reject INT!!")
            Exit Sub
        End If
        
        RcvHostDate = Pcb3dl.DlGetCharRaw("HostCurrentDate")
        RcvHostTime = Pcb3dl.DlGetCharRaw("TransJulianDays")
        'Get New Terminal Key from host
        nrc = S3ELineOut.GetData("NewPinKey", vRecvData)
        sTermMK = vRecvData
                       
        sTermMKEnd = ConvertData("o", sTermMK)
        
        nrc = S3ELineOut.GetData("NewPinKeyCheck", vRecvData)
        sCheckValue = vRecvData
        sCheckValue = ConvertData("o", sCheckValue)
        TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + RcvHostDate + "];" + vbCrLf
        TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + RcvHostTime + "];" + vbCrLf
        TxtHostReturn.Text = TxtHostReturn.Text + "NewTermalKey:[" + sTermMKEnd + "];" + vbCrLf
        TxtHostReturn.Text = TxtHostReturn.Text + "NewTermalKeyCheck:[" + sCheckValue + "]" + vbCrLf
                
        DecryptResult = DesDeCrypt(sTermMKEnd, sLocalMK)
        TermKey = DecryptResult
        
        DesResult = DesEncrypt("0000000000000000", DecryptResult)
        
        If Mid(DesResult, 1, 4) = sCheckValue Then
            TxtHostReturn.Text = TxtHostReturn.Text + "MasterKey Check OK!" + vbCrLf
            MsgBox ("INT Check OK!!")
        Else
            TxtHostReturn.Text = TxtHostReturn.Text + "MasterKey Check Failed!!" + vbCrLf
            MsgBox ("MasterKey Check Failed!")
        End If
        '检查完成
    Else
        TxtHostReturn.Text = "PuOpen Error(" + CStr(rcOpen) + ")" + vbCrLf
    End If
End Sub

Private Sub S3ELineOut_DevStateChanged()
    If S3ELineOut.Available Then
        Pcb3dl.DlSetCharRaw "GBLLineStatus", "O"
        S3ELineOut.BackColor = &HFF00&
    Else
        Pcb3dl.DlSetCharRaw "GBLLineStatus", "C"
        S3ELineOut.BackColor = &HFF&
    End If

End Sub

Private Sub CmdRQK_Click()
    Dim sNewPinKey As String, DecryptResult As String, DesResult As String, sCheckValue As String
    Dim sNewMACKey As String, sMacCheckValue As String
    Dim vRecvData As Variant
    Dim RcvHostTime As String, RcvHostDate As String
    Dim sHostTranCode As String
    
    TxtHostReturn.Text = ""
    
    nrc = S3ELineOut.DoSend("RQK", 0)
    If nrc <> 0 Then
        TxtHostReturn.Text = "RQK's DoSend Failed! RC=" + CStr(nrc)
        MsgBox "DoSend RC=" + CStr(nrc)
        Exit Sub
    End If
    
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If

    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "RQK's DoReceive Failed! RC=" + CStr(nrc)
        MsgBox "Receive RC=" + CStr(nrc)
        Exit Sub
    End If

    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    If sHostTranCode <> "DBP" And sHostTranCode <> "ABP" Then     'accept
        TxtHostReturn.Text = TxtHostReturn.Text + "RQK IS NOT ABP"
        Exit Sub
    End If

    RcvHostDate = Pcb3dl.DlGetCharRaw("HostCurrentDate")
    RcvHostTime = Pcb3dl.DlGetCharRaw("TransJulianDays")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + RcvHostDate + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + RcvHostTime + "];" + vbCrLf

    nrc = S3ELineOut.GetData("NewPinKey", vRecvData)
    sNewPinKey = vRecvData
    sNewPinKey = ConvertData("o", sNewPinKey)
     
    nrc = S3ELineOut.GetData("NewPinKeyCheck", vRecvData)
    sCheckValue = vRecvData
    sCheckValue = ConvertData("o", sCheckValue)
     
    TxtHostReturn.Text = TxtHostReturn.Text + "NewPinKey:[" + sNewPinKey + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "NewPinKeyCheck:[" + sCheckValue + "]" + vbCrLf
    
    nrc = S3ELineOut.GetData("NewMacKey", vRecvData)
    sNewMACKey = vRecvData
    sNewMACKey = ConvertData("o", sNewMACKey)
     
    nrc = S3ELineOut.GetData("NewMacKeyCheck", vRecvData)
    sMacCheckValue = vRecvData
    sMacCheckValue = ConvertData("o", sMacCheckValue)
    
    TxtHostReturn.Text = TxtHostReturn.Text + "NewMacKey:[" + sNewMACKey + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "NewMacKeyCheck:[" + sMacCheckValue + "]" + vbCrLf
    
    
    
    DecryptResult = DesDeCrypt(sNewPinKey, TermKey) 'RQK 回来的NewPinKey
    NewPinKey = DecryptResult
    nrc = Pcb3dl.DlSetCharRaw("GBLPrePinKey", DecryptResult)
    '检查checkvalue
    
    DesResult = DesEncrypt("0000000000000000", DecryptResult)
    If Mid(DesResult, 1, 4) = sCheckValue Then
        TxtHostReturn.Text = TxtHostReturn.Text + "New PIN Key Check OK!" + vbCrLf
    Else
        TxtHostReturn.Text = TxtHostReturn.Text + "New PIN Key Check Failed! RC=" + DesResult + vbCrLf
        MsgBox ("New PIN Key Check Failed!")
    End If
    'PinKey检查完成
    
    DecryptResult = DesDeCrypt(sNewMACKey, TermKey) 'RQK 回来的NewMacKey
    NewMACKey = DecryptResult
    '检查checkvalue
    
    DesResult = DesEncrypt("0000000000000000", DecryptResult)
    
    If Mid(DesResult, 1, 4) = sMacCheckValue Then
        Dim aa As String
        aa = "NEW MAC OK: " + DecryptResult
        MsgBox (aa)
        aa = S3ELineOut.MacSwSK
        TxtHostReturn.Text = TxtHostReturn.Text + "New MAC Key Check OK!" + vbCrLf
    Else
        TxtHostReturn.Text = TxtHostReturn.Text + "New MAC Key Check Failed! RC=" + DesResult + vbCrLf
        MsgBox ("New MAC Key Check Failed!")
    End If
    S3ELineOut.MacSwSK = DecryptResult
    'MacKey检查完成
End Sub

Private Sub CmdPIN_Click()
    Dim sCurrentDate As String
    Dim sCurrentTime As String
    Dim PinChangBlock As String
    Dim RcvHostTime As String, RcvHostDate As String
    Dim vRecvData As Variant
    Dim sHostRejectCode As String
    Dim sHostTranCode As String
    Dim sHostAtmCode As String, sPinBlock As String
    Dim sOldPIN As String, sNewPIN As String

    TxtHostReturn.Text = ""
        
'add for boc shenzhen to cal pin diff
    sOldPIN = TxtPin.Text
    sNewPIN = TxtNewPin.Text
    Call genPinBlockSW(sOldPIN, sPinBlock)
    nrc = Pcb3dl.DlSetCharRaw("PinInputBlock", sPinBlock)
    
    Call PIN_Difference(sOldPIN, sNewPIN)
' new pin use ANSI98 len=16  not mac in Guangdong
'   sPinBlock = ""
'    Call genPinBlockSW(sNewPIN, sPinBlock)
'    nrc = Pcb3dl.DlSetCharRaw("PinChangeBlock", sPinBlock)
    
    nrc = S3ELineOut.DoSend("PIN", 0)
    If nrc <> 0 Then
        TxtHostReturn.Text = "PIN's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If

    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "PIN's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostLineNum:[" + Pcb3dl.DlGetCharRaw("HostLineNum") + "];" + vbCrLf
    If sHostTranCode <> "DSP" And sHostTranCode <> "ASP" Then    'accept
        TxtHostReturn.Text = TxtHostReturn.Text + "PIN IS NOT ASP"
        If sHostTranCode = "AGP" Or sHostTranCode = "ATP" Then
            nrc = S3ELineOut.GetData("constHostRejectCode", vRecvData)
            sHostRejectCode = vRecvData
            TxtHostReturn.Text = TxtHostReturn.Text + "HostRejectCode:[" + sHostRejectCode + "];" + vbCrLf
        End If
        Exit Sub
    End If
    
    TxtHostReturn.Text = TxtHostReturn.Text + "HostAccNo:[" + Pcb3dl.DlGetCharRaw("HostAccNo") + "];" + vbCrLf
    nrc = S3ELineOut.GetData("HostTrack3", vRecvData)
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTrack3:[" + CStr(vRecvData) + "];" + vbCrLf
        
End Sub

Private Sub CmdINQ_Click()
    Dim sCurrentDate As String
    Dim sCurrentTime As String
    Dim sHostRejectCode As String
    Dim sHostTranCode As String
    Dim sHostAtmCode As String
    Dim RcvHostTime As String, RcvHostDate As String, sPinBlock As String
    Dim i As Integer
    Dim sPIN As String
    Dim vRecvData As Variant
    
    TxtHostReturn.Text = ""
    
    sPIN = TxtPin.Text
    Call genPinBlockSW(sPIN, sPinBlock)
    nrc = Pcb3dl.DlSetCharRaw("PinInputBlock", sPinBlock)
    
    nrc = S3ELineOut.DoSend("INQ", 0)

    If nrc <> 0 Then
        TxtHostReturn.Text = "INQ's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If

    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "INQ's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostLineNum:[" + Pcb3dl.DlGetCharRaw("HostLineNum") + "];" + vbCrLf
    If sHostTranCode <> "DRP" And sHostTranCode <> "ARP" Then     'accept
        TxtHostReturn.Text = TxtHostReturn.Text + "INQ IS NOT ARP"
        If sHostTranCode = "AGP" Or sHostTranCode = "ATP" Then
            nrc = S3ELineOut.GetData("constHostRejectCode", vRecvData)
            sHostRejectCode = vRecvData
            TxtHostReturn.Text = TxtHostReturn.Text + "HostRejectCode:[" + sHostRejectCode + "];" + vbCrLf
        End If
        Exit Sub
    End If
    
    TxtHostReturn.Text = TxtHostReturn.Text + "HostAccNo:[" + Pcb3dl.DlGetCharRaw("HostAccNo") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostFundAvail:[" + Pcb3dl.DlGetCharRaw("HostFundAvail") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostCurBal:[" + Pcb3dl.DlGetCharRaw("HostCurBal") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostSignCurBal:[" + Pcb3dl.DlGetCharRaw("HostSignCurBal") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostShAvailBal:[" + Pcb3dl.DlGetCharRaw("HostShAvailBal") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostSignAvailBal:[" + Pcb3dl.DlGetCharRaw("HostSignAvailBal") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "KeepHostBasicBalance:[" + Pcb3dl.DlGetCharRaw("HostShTfrAvail") + "];" + vbCrLf
    nrc = S3ELineOut.GetData("SignOfKeepHostBal", vRecvData)
    TxtHostReturn.Text = TxtHostReturn.Text + "SignOfKeepHostBal:[" + CStr(vRecvData) + "];" + vbCrLf
    nrc = S3ELineOut.GetData("HostTrack3", vRecvData)
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTrack3:[" + CStr(vRecvData) + "];" + vbCrLf
                
End Sub

Private Sub CmdCWD_Click()
    Dim SendSerialNo As String
    Dim WthAmount As Variant
    Dim StrWthAmount As String
    Dim sHostRejectCode As String
    Dim sHostTranCode As String
    Dim sHostAtmCode As String
    Dim RcvHostTime As String, RcvHostDate As String
    Dim sPIN As String, sPinBlock As String
    Dim vRecvData As Variant

    TxtHostReturn.Text = ""

    'Set Withdraw Amount
    WthAmount = TxtAmount.Text * 100
    StrWthAmount = Format(WthAmount, "00000000")
    nrc = Pcb3dl.DlSetCharRaw("GBLAmount", StrWthAmount)
            
    'Set Pin
    sPIN = TxtPin.Text
    Call genPinBlockSW(sPIN, sPinBlock)
    nrc = Pcb3dl.DlSetCharRaw("PinInputBlock", sPinBlock)

    nrc = S3ELineOut.DoSend("CWD", 0)
    nrc = S3ELineOut.GetData("GBLLineNum", vRecvData)
    SendSerialNo = vRecvData
    TxtCWCNo.Text = SendSerialNo
    If nrc <> 0 Then
        TxtHostReturn.Text = "CWD's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If

    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "CWD's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
            
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostLineNum:[" + Pcb3dl.DlGetCharRaw("HostLineNum") + "];" + vbCrLf
    If sHostTranCode <> "AQP" Then     'accept
        TxtHostReturn.Text = TxtHostReturn.Text + "CWD IS NOT DQP"
        If sHostTranCode = "AGP" Or sHostTranCode = "ATP" Or sHostTranCode = "AUP" Then
            nrc = S3ELineOut.GetData("constHostRejectCode", vRecvData)
            sHostRejectCode = vRecvData
            TxtHostReturn.Text = TxtHostReturn.Text + "HostRejectCode:[" + sHostRejectCode + "];" + vbCrLf
            If sHostTranCode = "AUP" Then
                TxtHostReturn.Text = TxtHostReturn.Text + "HostAvailBalance:[" + Pcb3dl.DlGetCharRaw("HostAvailBalance") + "];" + vbCrLf
            End If
        End If
        Exit Sub
    End If
            
    TxtHostReturn.Text = TxtHostReturn.Text + "HostAccNo:[" + Pcb3dl.DlGetCharRaw("HostAccNo") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTransAmount:[" + Pcb3dl.DlGetCharRaw("HostTransAmount") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "AuthCode:[" + Pcb3dl.DlGetCharRaw("IcbcHostSeq") + "];" + vbCrLf
    nrc = S3ELineOut.GetData("HostTrack3", vRecvData)
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTrack3:[" + CStr(vRecvData) + "];" + vbCrLf
            
End Sub

Private Sub CmdCWC_Click()
    Dim sCurrentDate As String
    Dim sCurrentTime As String
    Dim SendSerialNo As String
    Dim WthAmount As Variant
    Dim StrWthAmount As String
    Dim sHostRejectCode As String
    Dim sHostTranCode As String
    Dim sHostAtmCode As String
    Dim RcvHostTime As String, RcvHostDate As String
    Dim sPIN As String, sPinBlock As String
    Dim vRecvData As Variant
    
    TxtHostReturn.Text = ""
    
    'Set Pin
    sPIN = TxtPin.Text
    Call genPinBlockSW(sPIN, sPinBlock)
    nrc = Pcb3dl.DlSetCharRaw("PinInputBlock", sPinBlock)
    
    'Set CWC Serial
    nrc = Pcb3dl.DlSetCharRaw("GBLLineSendNum", TxtCWCNo.Text)
    
    'Set Withdraw Amount
    WthAmount = WthAmount * 100
    StrWthAmount = Format(WthAmount, "00000000")
    nrc = Pcb3dl.DlSetCharRaw("GBLAmount", StrWthAmount)
            

    nrc = S3ELineOut.DoSend("CWC", 0)
    If nrc <> 0 Then
        TxtHostReturn.Text = "CWC's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If

    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "CWC's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
            
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostLineNum:[" + Pcb3dl.DlGetCharRaw("HostLineNum") + "];" + vbCrLf
    If sHostTranCode <> "DWP" And sHostTranCode <> "AWP" Then    'accept
        TxtHostReturn.Text = TxtHostReturn.Text + "CWC IS NOT AWP"
        If sHostTranCode = "AGP" Or sHostTranCode = "ATP" Then
            nrc = S3ELineOut.GetData("constHostRejectCode", vRecvData)
            sHostRejectCode = vRecvData
            TxtHostReturn.Text = TxtHostReturn.Text + "HostRejectCode:[" + sHostRejectCode + "];" + vbCrLf
        End If
        Exit Sub
    End If
            
    TxtHostReturn.Text = TxtHostReturn.Text + "HostAccNo:[" + Pcb3dl.DlGetCharRaw("HostAccNo") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTransAmount:[" + Pcb3dl.DlGetCharRaw("HostTransAmount") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "AuthCode:[" + Pcb3dl.DlGetCharRaw("HostRefNumber") + "];" + vbCrLf
    nrc = S3ELineOut.GetData("HostTrack3", vRecvData)
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTrack3:[" + CStr(vRecvData) + "];" + vbCrLf
    
End Sub

Private Sub CmdTFR_Click()
    Dim sCurrentDate As String
    Dim sCurrentTime As String
    Dim SendSerialNo As String
'    Dim TransCode As String
    Dim WthAmount As Variant
    Dim StrWthAmount As String
    Dim sHostRejectCode As String
    Dim sHostTranCode As String
    Dim sHostAtmCode As String
    Dim sTFRType As String
    Dim RcvHostTime As String, RcvHostDate As String
    Dim sPIN As String, sPinBlock As String
    Dim vRecvData As Variant

    TxtHostReturn.Text = ""
    
    'Set Pin
    sPIN = TxtPin.Text
    Call genPinBlockSW(sPIN, sPinBlock)
    nrc = Pcb3dl.DlSetCharRaw("PinInputBlock", sPinBlock)
    
    'Set TFR Amount
    WthAmount = TxtAmount.Text * 100
    StrWthAmount = Format(WthAmount, "00000000")
    nrc = Pcb3dl.DlSetCharRaw("GBLAmount", StrWthAmount)
            
    'Set TFR's account
    nrc = Pcb3dl.DlSetCharRaw("Tfr2ndAccNo", TxtTfr2Acc.Text)
    
    nrc = S3ELineOut.DoSend("TFR", 0)
    If nrc <> 0 Then
        TxtHostReturn.Text = "TFR's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If

    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "TFR's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
            
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostLineNum:[" + Pcb3dl.DlGetCharRaw("HostLineNum") + "];" + vbCrLf
    If sHostTranCode <> "DQP" And sHostTranCode <> "AQP" Then    'accept
        TxtHostReturn.Text = TxtHostReturn.Text + "TFR IS NOT AQP"
        If sHostTranCode = "AGP" Or sHostTranCode = "ATP" Or sHostTranCode = "AUP" Then
            nrc = S3ELineOut.GetData("constHostRejectCode", vRecvData)
            sHostRejectCode = vRecvData
            TxtHostReturn.Text = TxtHostReturn.Text + "HostRejectCode:[" + sHostRejectCode + "];" + vbCrLf
            If sHostTranCode = "AUP" Then
                TxtHostReturn.Text = TxtHostReturn.Text + "HostAvailBalance:[" + Pcb3dl.DlGetCharRaw("HostAvailBalance") + "];" + vbCrLf
            End If
        End If
        Exit Sub
    End If
            
    TxtHostReturn.Text = TxtHostReturn.Text + "HostAccNo:[" + Pcb3dl.DlGetCharRaw("HostAccNo") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTransAmount:[" + Pcb3dl.DlGetCharRaw("HostTransAmount") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "AuthCode:[" + Pcb3dl.DlGetCharRaw("IcbcHostSeq") + "];" + vbCrLf
    nrc = S3ELineOut.GetData("HostTrack3", vRecvData)
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTrack3:[" + CStr(vRecvData) + "];" + vbCrLf
End Sub

Private Sub CmdTFC_Click()
    Dim sCurrentDate As String
    Dim sCurrentTime As String
    Dim SendSerialNo As String
    Dim WthAmount As Variant
    Dim StrWthAmount As String
    Dim sHostRejectCode As String
    Dim sHostTranCode As String
    Dim sHostAtmCode As String
    Dim RcvHostTime As String, RcvHostDate As String
    Dim sPIN As String, sPinBlock As String
    Dim vRecvData As Variant
    
    MsgBox ("TFC Transaction Not Support!")
    Exit Sub
    
    TxtHostReturn.Text = ""
    
    'Set Pin
    sPIN = TxtPin.Text
    Call genPinBlockSW(sPIN, sPinBlock)
    nrc = Pcb3dl.DlSetCharRaw("PinInputBlock", sPinBlock)
    
    'Set TFC Amount
    WthAmount = WthAmount * 100
    StrWthAmount = Format(WthAmount, "00000000")
    nrc = Pcb3dl.DlSetCharRaw("GBLAmount", StrWthAmount)
            
    'Set TFC's account
    nrc = Pcb3dl.DlSetCharRaw("Tfr2ndAccNo", TxtTfr2Acc.Text)
    
    nrc = S3ELineOut.DoSend("TFC", 0)
    nrc = S3ELineOut.GetData("GBLLineNum", vRecvData)
    SendSerialNo = vRecvData
    TxtCWCNo.Text = SendSerialNo
    If nrc <> 0 Then
        TxtHostReturn.Text = "TFC's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If

    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "TFC's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
            
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostLineNum:[" + Pcb3dl.DlGetCharRaw("HostLineNum") + "];" + vbCrLf
    If sHostTranCode <> "DQP" Then     'accept
        TxtHostReturn.Text = TxtHostReturn.Text + "TFC IS NOT AQP"
        If sHostTranCode = "DGP" Or sHostTranCode = "DTP" Then
            nrc = S3ELineOut.GetData("constHostRejectCode", vRecvData)
            sHostRejectCode = vRecvData
            TxtHostReturn.Text = TxtHostReturn.Text + "HostRejectCode:[" + sHostRejectCode + "];" + vbCrLf
        End If
        Exit Sub
    End If
            
    TxtHostReturn.Text = TxtHostReturn.Text + "HostAccNo:[" + Pcb3dl.DlGetCharRaw("HostAccNo") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTransAmount:[" + Pcb3dl.DlGetCharRaw("HostTransAmount") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "AuthCode:[" + Pcb3dl.DlGetCharRaw("IcbcHostSeq") + "];" + vbCrLf
    nrc = S3ELineOut.GetData("HostTrack3", vRecvData)
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTrack3:[" + CStr(vRecvData) + "];" + vbCrLf
End Sub

Private Sub CmdCDP_Click()
    Dim SendSerialNo As String
    Dim WthAmount As Variant
    Dim StrWthAmount As String
    Dim sHostRejectCode As String
    Dim sHostTranCode As String
    Dim sHostAtmCode As String
    Dim RcvHostTime As String, RcvHostDate As String
    Dim sPIN As String, sPinBlock As String
    Dim vRecvData As Variant

    TxtHostReturn.Text = ""

    'Set CDP Amount
    WthAmount = TxtAmount.Text * 100
    StrWthAmount = Format(WthAmount, "00000000")
    nrc = Pcb3dl.DlSetCharRaw("GBLAmount", StrWthAmount)
            
    nrc = S3ELineOut.DoSend("CDP", 0)
    If nrc <> 0 Then
        TxtHostReturn.Text = "CDP's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If

    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "CDP's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
            
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostLineNum:[" + Pcb3dl.DlGetCharRaw("HostLineNum") + "];" + vbCrLf
    If sHostTranCode <> "AQP" Then     'accept
        TxtHostReturn.Text = TxtHostReturn.Text + "CDP IS NOT AQP"
        If sHostTranCode = "AGP" Or sHostTranCode = "ATP" Then
            nrc = S3ELineOut.GetData("constHostRejectCode", vRecvData)
            sHostRejectCode = vRecvData
            TxtHostReturn.Text = TxtHostReturn.Text + "HostRejectCode:[" + sHostRejectCode + "];" + vbCrLf
        End If
        Exit Sub
    End If
            
    TxtHostReturn.Text = TxtHostReturn.Text + "HostAccNo:[" + Pcb3dl.DlGetCharRaw("HostAccNo") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTransAmount:[" + Pcb3dl.DlGetCharRaw("HostTransAmount") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "AuthCode:[" + Pcb3dl.DlGetCharRaw("IcbcHostSeq") + "];" + vbCrLf
    nrc = S3ELineOut.GetData("HostTrack3", vRecvData)
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTrack3:[" + CStr(vRecvData) + "];" + vbCrLf
            
End Sub

Private Sub CmdERI_Click()
    Dim sHostTranCode As String
    Dim vRecvData As Variant

    TxtHostReturn.Text = ""
    nrc = S3ELineOut.DoSend("ERI", 0)
    If nrc <> 0 Then
        TxtHostReturn.Text = "ERI's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If

    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "ERI's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
    If sHostTranCode <> "DIP" And sHostTranCode <> "AIP" Then    'accept
        TxtHostReturn.Text = TxtHostReturn.Text + "ERI IS NOT DIP"
        Exit Sub
    End If
    
    TxtHostReturn.Text = TxtHostReturn.Text + "HostLineNum:[" + Pcb3dl.DlGetCharRaw("HostLineNum") + "];" + vbCrLf
    nrc = S3ELineOut.GetData("Accepted", vRecvData)
    TxtHostReturn.Text = TxtHostReturn.Text + "Accepted:[" + CStr(vRecvData) + "];" + vbCrLf
    nrc = S3ELineOut.GetData("ExchangeRate", vRecvData)
    TxtHostReturn.Text = TxtHostReturn.Text + "ExchangeRate:[" + CStr(vRecvData) + "];" + vbCrLf
End Sub

Private Sub CmdPAN_Click()
    Dim sHostRejectCode As String
    Dim sHostTranCode As String
    Dim sHostAtmCode As String
    Dim RcvHostTime As String, RcvHostDate As String, sPinBlock As String
    Dim i As Integer
    Dim sPIN As String
    Dim vRecvData As Variant
    
    TxtHostReturn.Text = ""
    
    sPIN = TxtPin.Text
    Call genPinBlockSW(sPIN, sPinBlock)
    nrc = Pcb3dl.DlSetCharRaw("PinInputBlock", sPinBlock)
    
    nrc = S3ELineOut.DoSend("PAN", 0)

    If nrc <> 0 Then
        TxtHostReturn.Text = "PAN's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If

    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "PAN's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
    If sHostTranCode <> "DPP" And sHostTranCode <> "APP" Then    'accept
        TxtHostReturn.Text = TxtHostReturn.Text + "PAN IS NOT APP"
'        If sHostTranCode = "AGP" Or sHostTranCode = "ATP" Then
'            nrc = S3ELineOut.GetData("constHostRejectCode", vRecvData)
'            sHostRejectCode = vRecvData
'            TxtHostReturn.Text = TxtHostReturn.Text + "HostRejectCode:[" + sHostRejectCode + "];" + vbCrLf
'        End If
        Exit Sub
    End If
    
    TxtHostReturn.Text = TxtHostReturn.Text + "HostAccNo:[" + Pcb3dl.DlGetCharRaw("HostAccNo") + "];" + vbCrLf
    nrc = S3ELineOut.GetData("HostCardType", vRecvData)
    TxtHostReturn.Text = TxtHostReturn.Text + "HostCardType:[" + CStr(vRecvData) + "];" + vbCrLf
    nrc = S3ELineOut.GetData("HostTrackType", vRecvData)
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTrackType:[" + CStr(vRecvData) + "];" + vbCrLf
End Sub


Private Sub CmdTTI_Click()
    Dim sHostTranCode As String
    Dim vRecvData As Variant

    TxtHostReturn.Text = ""
    nrc = S3ELineOut.DoSend("TTI", 0)
    If nrc <> 0 Then
        TxtHostReturn.Text = "TTI's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If

    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "TTI's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
    If sHostTranCode = "AAP" Then     'accept
        TxtHostReturn.Text = TxtHostReturn.Text + "TTI IS AAP: ATMC Totals matches ATMP Totals"
        Exit Sub
    End If
    
    If sHostTranCode = "ACP" Then     'Reject
        nrc = S3ELineOut.GetData("CasDemo1", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasDemo1:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("CasDemo2", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasDemo2:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("CasDemo3", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasDemo3:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("CasDemo4", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasDemo4:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("CasInitLoad1", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasInitLoad1:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("CasInitLoad2", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasInitLoad2:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("CasInitLoad3", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasInitLoad3:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("CasInitLoad4", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasInitLoad4:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("DenoRef1", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "DenoRef1:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("DenoRef2", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "DenoRef2:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("DenoRef3", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "DenoRef3:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("DenoRef4", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "DenoRef4:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("NoOfHKDWth", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "NoOfHKDWth:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("NoOfRMBWth", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "NoOfRMBWth:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("NoOfDep", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "NoOfDep:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("AmtOfDep", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "AmtOfDep:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("NoOfHKDTfr", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "NoOfHKDTfr:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("AmtOfHKDTfr", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "AmtOfHKDTfr:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("NoOfRMBTfr", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "NoOfRMBTfr:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("AmtOfRMBTfr", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "AmtOfRMBTfr:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("TotCapCardNum", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "TotCapCardNum:[" + CStr(vRecvData) + "];" + vbCrLf
    End If
End Sub

Private Sub CmdRWT_Click()
    Dim sHostTranCode As String
    Dim vRecvData As Variant

    TxtHostReturn.Text = ""
    nrc = S3ELineOut.DoSend("RWT", 0)
    If nrc <> 0 Then
        TxtHostReturn.Text = "RWT's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If
    
    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "RWT's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
    If sHostTranCode = "AAP" Then     'accept
        TxtHostReturn.Text = TxtHostReturn.Text + "RWT IS AAP: ATMC Totals matches ATMP Totals"
        Exit Sub
    End If
    
    If sHostTranCode = "ADP" Then     'Reject
        nrc = S3ELineOut.GetData("CasDemo1", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasDemo1:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("CasDemo2", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasDemo2:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("CasDemo3", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasDemo3:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("CasDemo4", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasDemo4:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("CasInitLoad1", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasInitLoad1:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("CasInitLoad2", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasInitLoad2:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("CasInitLoad3", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasInitLoad3:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("CasInitLoad4", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "CasInitLoad4:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("DenoRef1", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "DenoRef1:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("DenoRef2", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "DenoRef2:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("DenoRef3", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "DenoRef3:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("DenoRef4", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "DenoRef4:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("NoOfHKDWth", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "NoOfHKDWth:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("NoOfRMBWth", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "NoOfRMBWth:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("TotCapCardNum", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "TotCapCardNum:[" + CStr(vRecvData) + "];" + vbCrLf
    End If
End Sub

Private Sub CmdRDT_Click()
    Dim sHostTranCode As String
    Dim vRecvData As Variant

    TxtHostReturn.Text = ""
    nrc = S3ELineOut.DoSend("RDT", 0)
    If nrc <> 0 Then
        TxtHostReturn.Text = "RDT's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If

    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "RDT's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
    If sHostTranCode = "AAP" Then     'accept
        TxtHostReturn.Text = TxtHostReturn.Text + "RDT IS AAP: ATMC Totals matches ATMP Totals"
        Exit Sub
    End If
    
    If sHostTranCode = "AEP" Then     'Reject
        nrc = S3ELineOut.GetData("NoOfDep", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "NoOfDep:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("AmtOfDep", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "AmtOfDep:[" + CStr(vRecvData) + "];" + vbCrLf
    End If
End Sub

Private Sub CmdRTT_Click()
    Dim sHostTranCode As String
    Dim vRecvData As Variant

    TxtHostReturn.Text = ""
    nrc = S3ELineOut.DoSend("RTT", 0)
    If nrc <> 0 Then
        TxtHostReturn.Text = "RTT's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If

    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "RTT's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
    If sHostTranCode = "AAP" Then     'accept
        TxtHostReturn.Text = TxtHostReturn.Text + "RTT IS AAP: ATMC Totals matches ATMP Totals"
        Exit Sub
    End If
    
    If sHostTranCode = "AFP" Then     'Reject
        nrc = S3ELineOut.GetData("NoOfHKDTfr", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "NoOfHKDTfr:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("AmtOfHKDTfr", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "AmtOfHKDTfr:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("NoOfRMBTfr", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "NoOfRMBTfr:[" + CStr(vRecvData) + "];" + vbCrLf
        nrc = S3ELineOut.GetData("AmtOfRMBTfr", vRecvData)
        TxtHostReturn.Text = TxtHostReturn.Text + "AmtOfRMBTfr:[" + CStr(vRecvData) + "];" + vbCrLf
    End If
End Sub

Private Sub CmdAEX_Click()
    Dim sHostRejectCode As String
    Dim sHostTranCode As String
    Dim sHostAtmCode As String
    Dim RcvHostTime As String, RcvHostDate As String, sPinBlock As String
    Dim i As Integer
    Dim sPIN As String
    Dim vRecvData As Variant

    TxtHostReturn.Text = ""
    
    nrc = S3ELineOut.SetData("ExceptionCode", "2036")
    
    nrc = S3ELineOut.DoSend("AEX", 0)
    If nrc <> 0 Then
        TxtHostReturn.Text = "AEX's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If

    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "AEX's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
'    If sHostTranCode <> "AXP" Then     'accept
'        TxtHostReturn.Text = TxtHostReturn.Text + "AEX IS NOT AXP"
''        If sHostTranCode = "AGP" Or sHostTranCode = "ATP" Then
''            nrc = S3ELineOut.GetData("constHostRejectCode", vRecvData)
''            sHostRejectCode = vRecvData
''            TxtHostReturn.Text = TxtHostReturn.Text + "HostRejectCode:[" + sHostRejectCode + "];" + vbCrLf
''        End If
'        Exit Sub
'    End If
'
'    nrc = S3ELineOut.GetData("HostTrack3", vRecvData)
'    TxtHostReturn.Text = TxtHostReturn.Text + "HostTrack3:[" + CStr(vRecvData) + "];" + vbCrLf
End Sub

Private Sub CmdTEX_Click()
    Dim sHostTranCode As String

    TxtHostReturn.Text = ""
    
    nrc = S3ELineOut.SetData("ExceptionCode", "2036")
    
    nrc = S3ELineOut.DoSend("TEX", 0)
    If nrc <> 0 Then
        TxtHostReturn.Text = "TEX's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If
    
    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "TEX's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
End Sub

Private Sub CmdOEX_Click()
    Dim sHostRejectCode As String
    Dim sHostTranCode As String
    Dim sHostAtmCode As String
    Dim RcvHostTime As String, RcvHostDate As String, sPinBlock As String
    Dim i As Integer
    Dim sPIN As String
    Dim vRecvData As Variant

    TxtHostReturn.Text = ""
    
    nrc = S3ELineOut.SetData("ExceptionCode", "2036")
    
    nrc = S3ELineOut.DoSend("OEX", 0)
    If nrc <> 0 Then
        TxtHostReturn.Text = "OEX's DoSend Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    If S3ELineOut.BackColor <> &HFF00& Then
        S3ELineOut.BackColor = &HFF00&
    End If

    nrc = S3ELineOut.DoReceive
    If nrc <> 0 Then
        TxtHostReturn.Text = "OEX's DoReceive Failed! RC=" + CStr(nrc)
        Exit Sub
    End If
    
    sHostTranCode = Pcb3dl.DlGetCharRaw("HostTransCode")
    TxtHostReturn.Text = TxtHostReturn.Text + "HostTranCode:[" + sHostTranCode + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "HostDate:[" + Pcb3dl.DlGetCharRaw("HostCurrentDate") + "];" + vbCrLf
    TxtHostReturn.Text = TxtHostReturn.Text + "JulianDays:[" + Pcb3dl.DlGetCharRaw("TransJulianDays") + "];" + vbCrLf
End Sub

Private Sub SetNewDeviceStatus(ByVal DevOffset As Integer, ByVal bValue As String)

     G_sDeviceStatus = Left(G_sDeviceStatus, DevOffset - 1) + bValue + Right(G_sDeviceStatus, 13 - DevOffset)
'    ElseIf bValue = False And sDevStatusValue = "1" Then
'        G_sNewDeviceStatus = Left(G_sNewDeviceStatus, DevOffset - 1) + "0" _
'                + Right(G_sNewDeviceStatus, 64 - DevOffset)
'    End If
End Sub

Private Function DesEncrypt(DesData As String, DesKey As String) As String
    Dim nrc As Integer
    SDOEdm.CryptMode = True

        SDOEdm.CryptType = 1
        nrc = SDOEdm.DoCryptDataSw(DesData, DesKey)
        If nrc = 0 Then
            DesEncrypt = SDOEdm.CryptResult
            Exit Function
        Else
            DesEncrypt = ""
        End If
End Function

Private Function DesDeCrypt(DesData As String, DesKey As String) As String
    Dim nrc As Integer
    SDOEdm.CryptMode = False
    
    SDOEdm.CryptType = 1
    nrc = SDOEdm.DoCryptDataSw(DesData, DesKey)
        If nrc = 0 Then
            DesDeCrypt = SDOEdm.CryptResult
            Exit Function
        Else
            DesDeCrypt = ""
        End If
End Function
Private Function ConvertData(ConvertFlag As String, ConData As String) As String
    
    Dim strNum As String
    Dim i As Integer
    Select Case ConvertFlag
     Case "o"
     
          For i = 1 To Len(ConData)
            strNum = Mid(ConData, i, 1)
            Select Case strNum
                Case ":"
                    strNum = "A"
                Case ";"
                    strNum = "B"
                Case "<"
                    strNum = "C"
                Case "="
                    strNum = "D"
                Case ">"
                    strNum = "E"
                Case "?"
                    strNum = "F"
            End Select
            ConvertData = ConvertData + strNum
        Next
        
    Case "p"
     
          For i = 1 To Len(ConData)
            strNum = Mid(ConData, i, 1)
            Select Case strNum
                Case "A"
                    strNum = ":"
                Case "B"
                    strNum = ";"
                Case "C"
                    strNum = "<"
                Case "D"
                    strNum = "="
                Case "E"
                    strNum = ">"
                Case "F"
                    strNum = "?"
            End Select
            ConvertData = ConvertData + strNum
        Next
    End Select
    
End Function

Public Sub StrToBin(ByVal inString As String, ByRef bOutArray() As Byte)
    Dim strTwo As String
    Dim i As Integer, j As Integer

    j = 0
    For i = 1 To 16 Step 2
        strTwo = Mid(inString, i, 2)
        bOutArray(j) = Val("&H" + strTwo)
        j = j + 1
    Next

End Sub
Public Sub BinToStr(ByRef InPar() As Byte, ByRef OutPar As String)
    Dim i As Integer
    Dim strNum As String
    

    For i = 0 To 7
        strNum = Hex(InPar(i))
        If Len(strNum) < 2 Then
            strNum = "0" + strNum
        End If
        OutPar = OutPar + strNum
    Next i
End Sub

Private Sub genPinBlockSW(pPin As String, ByRef pPinEnd As String)
    Dim PinLen As String
    Dim DesResult As String
    Dim strPinKey1 As String
    Dim strPinKey2 As String
    Dim strPinKey3 As String

    Dim AccountNo As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim PinDataArray(7) As Byte
    Dim PrePinData As String
    Dim PrePinDataArray(7) As Byte
    Dim CardData As String
    Dim PreCardData As String
    
    Dim PreCardDataArray(7) As Byte
    
    Dim PinData As String
    Dim UseTriDES As String
    
    
    AccountNo = Pcb3dl.DlGetCharRaw("FitAccNo")
    
        strPinKey1 = Pcb3dl.DlGetCharRaw("GBLPrePinKey")
    
    j = Len(pPin)
    If j < 10 Then
        PinLen = Format(CStr(Len(pPin)), "00")
    ElseIf j = 10 Then
        PinLen = "0A"
    ElseIf j = 11 Then
        PinLen = "0B"
    ElseIf j = 12 Then
        PinLen = "0C"
    End If
    
    PrePinData = PinLen + pPin + String(14 - j, "F")
    Call StrToBin(PrePinData, PrePinDataArray)
    
    k = Len(AccountNo)
    If k > 12 Then
        PreCardData = Mid(AccountNo, k - 12, 12)
    Else
        PreCardData = AccountNo
    End If
    
'    CardData = String(4, "0") + Mid(AccountNo, Len(AccountNo) - 12, 12)
    CardData = String(4, "0") + PreCardData
    
    Call StrToBin(CardData, PreCardDataArray)
    
    For i = 0 To 7
        PinDataArray(i) = PrePinDataArray(i) Xor PreCardDataArray(i)
    Next
    
    Call BinToStr(PinDataArray, PinData)
    
    SDOEdm.CryptType = 1
    SDOEdm.CryptMode = True

    nrc = SDOEdm.DoCryptDataSw(PinData, strPinKey1)

    DesResult = SDOEdm.CryptResult
    Dim strNum As String
    For i = 1 To Len(DesResult)
        strNum = Mid(DesResult, i, 1)
        Select Case strNum
            Case "A"
                strNum = ":"
            Case "B"
                strNum = ";"
            Case "C"
                strNum = "<"
            Case "D"
                strNum = "="
            Case "E"
                strNum = ">"
            Case "F"
                strNum = "?"
        End Select
        pPinEnd = pPinEnd + strNum
    Next
    'Call StrToBin(DesResult, pPinEnd)

End Sub
Private Sub PIN_Difference(OldPin As String, NewPin As String)
    Dim lnew As Long, lold As Long, lDifference As Long
    Dim iParity As Integer, iParity1 As Integer
    Dim result As String, result1 As String

    iParity = (CInt(Mid(NewPin, 1, 1)) * 6 + CInt(Mid(NewPin, 2, 1)) * 5 + _
            CInt(Mid(NewPin, 3, 1)) * 4 + CInt(Mid(NewPin, 4, 1)) * 3 + _
            CInt(Mid(NewPin, 5, 1)) * 2 + CInt(Mid(NewPin, 6, 1))) Mod 10
    iParity1 = (10 - iParity) Mod 10
    
    If Len(NewPin) > 5 Then
        lnew = CLng(Mid(NewPin, 1, 5))
    Else
        lnew = CLng(NewPin)
    End If
    
    If Len(OldPin) > 5 Then
        lold = CLng(Mid(OldPin, 1, 5))
    Else
        lold = CLng(OldPin)
    End If
    
    lDifference = lnew - lold
    
    If lDifference < 0 Then
        lDifference = lDifference + 100000
    End If
    
    result = CStr(iParity1)
    result1 = Format((CStr(lDifference)), "0000000")
    result = result + result1
    nrc = Pcb3dl.DlSetCharRaw("PinChangeBlock", result)
End Sub
