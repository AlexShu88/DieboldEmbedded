VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label ReleaseTime2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label ReleaseTime1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "5.6.2发布时间 "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "5.6.1发布时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
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
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Enum RootKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
End Enum

Private Const OPTEVAPlateformVer    As String = "SOFTWARE\Diebold"

Private Sub Form_Load()
  OptevaImgVer.Caption = GetRegKeyS(HKEY_LOCAL_MACHINE, OPTEVAPlateformVer, "OptevaImage", 20, "")
 
 If OptevaImgVer.Caption = "5.6.1" Then
    Text1.Text = "请安装patch，升级到5.6.2"
 ElseIf OptevaImgVer.Caption = "5.6.2" Then
    Text1.Text = "系统版本正确， 可以安装应用"
 Else
    Text1.Text = "系统版本不正确，请安装5.6.2版"
 End If
 
 ReleaseTime1.Caption = "2005-11-02"
 ReleaseTime2.Caption = "2005-12-19"
End Sub

Function GetRegKeyS(ByVal RootKeyName As Long, ByVal SubRegKeyPath As String, ByVal KeyValue As String, ByVal DefSize As Long, ByVal DefValue As String) As String
    Dim nReply As Long
    Dim hKey As Long, nSize As Long
    Dim sDispArray() As Byte
    
    nReply = RegOpenKey(RootKeyName, SubRegKeyPath, hKey)
    If ERROR_SUCCESS = nReply Then
        nSize = DefSize
        ReDim sDispArray(0 To nSize)
        RegQueryValueEx hKey, KeyValue, 0, REG_BINARY, sDispArray(0), nSize
        ReDim Preserve sDispArray(0 To nSize - 2)
        GetRegKeyS = StrConv(sDispArray, vbUnicode)
    Else
        GetRegKeyS = DefValue
    End If
    RegCloseKey hKey
End Function

