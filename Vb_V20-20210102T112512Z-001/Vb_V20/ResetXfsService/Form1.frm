VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ResetPrr 1.0"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6795
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "重启XFS服务"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "秒"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "倒计时："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "正在重启XS服务，请稍候！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const WaitTime As Integer = 5
Dim Timeused As Integer
Dim DisPlayTime As Integer

Private Sub Command1_Click()
    Dim sTmp As String
    Dim nRc As Variant
    
    Command1.Enabled = False
    Label1.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    DisPlayTime = 14
    Label4.Caption = CStr(DisPlayTime)
    sTmp = "net stop " & Chr(34) & "diebold XFS" & Chr(34)
    nRc = Shell(sTmp, 0)
    Timeused = 0
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Dim sTmp As String
    Dim nRc As Variant

    Timer1.Enabled = False
    If Timeused > WaitTime Then
        sTmp = "net start " & Chr(34) & "diebold XFS" & Chr(34)
        nRc = Shell(sTmp, 0)
        Timeused = 0
        DisPlayTime = DisPlayTime - 1
        Label4.Caption = CStr(DisPlayTime)
        Timer2.Enabled = True
    Else
        DisPlayTime = DisPlayTime - 1
        Label4.Caption = CStr(DisPlayTime)
        Timeused = Timeused + 1
        Timer1.Enabled = True
    End If
       
End Sub
Private Sub Timer2_Timer()
    Dim sTmp As String
    Dim nRc As Variant

    Timer2.Enabled = False
    If Timeused > WaitTime Then
        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        sTmp = "请运行桌面上的“ReceiptTest”工具检测收条打印格式！"
        nRc = MsgBox(sTmp, vbOKOnly, "提示")
        Unload Me
    Else
        DisPlayTime = DisPlayTime - 1
        Label4.Caption = CStr(DisPlayTime)
        Timeused = Timeused + 1
        Timer2.Enabled = True
    End If
       
End Sub

