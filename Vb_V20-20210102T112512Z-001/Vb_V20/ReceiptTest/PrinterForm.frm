VERSION 5.00
Object = "{3751B5D1-D348-11D0-AD02-0060970C3D2F}#1.0#0"; "SDOPrr.ocx"
Object = "{9C37E835-6A58-11D1-80C0-0020AF7093F9}#1.2#0"; "Dl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ReceiptTest 1.0"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   Icon            =   "PrinterForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "确认修改"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox T_RowHeight 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "修改行距"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin DLLib.DL PCB3DL 
      Left            =   120
      Top             =   120
      _Version        =   65538
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin SDOPrrLibCtl.SDOPrr S3EPrr 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "PrinterForm.frx":27A2
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打印收条"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "打印收条行距："
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
      Left            =   1440
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "收条打印机正在初始化，请稍候"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const RowHeightPath = "SOFTWARE\XFS\PHYSICAL_SERVICES\DBD_ReceiptPtr\FH"

Const PrintFormName As String = "ATMPrr"
Dim nRc As String
Dim nReply As Integer

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    Label1.Caption = "请输入新的收条行距值后点击“确认修改”！"
    Command2.Enabled = False
    CmdExit.Enabled = False
    Command3.Enabled = True
    T_RowHeight.Enabled = True
End Sub

Private Sub Command2_Click()
    S3EPrr.DoPrintForm (PrintFormName)
End Sub

Private Sub Command3_Click()
    Dim nRev As Variant

    Command2.Enabled = False
    CmdExit.Enabled = False
    Command3.Enabled = False
    T_RowHeight.Enabled = True
    nRev = SetRegKeyValue(HKEY_LOCAL_MACHINE, RowHeightPath, "RowHeight", REG_SZ, T_RowHeight.Text)
    nRev = MsgBox("请运行桌面上的“ResetXfsService”工具！", vbOKOnly, "提示")
    Unload Me
End Sub

Private Sub Form_Load()

    nReply = PCB3DL.DlSetCharRaw("PrrWthMark", "***")
    nReply = PCB3DL.DlSetCharRaw("PrrTransferMark", "***")
    nReply = PCB3DL.DlSetCharRaw("PrrOthersMark", "***")
    nReply = PCB3DL.DlSetCharRaw("PrrCardRetainMark", "***")
    nReply = PCB3DL.DlSetCharRaw("PrrContactBankMark", "***")
    nReply = PCB3DL.DlSetCharRaw("PrrAcceptMark", "***")
    nReply = PCB3DL.DlSetCharRaw("PrrAcceptCode", "0000")
    nReply = PCB3DL.DlSetCharRaw("PrrRejectedCode", "00")
    nReply = PCB3DL.DlSetCharRaw("PrrCardRetainMark", "***")
    nReply = PCB3DL.DlSetCharRaw("PrrContactBankMark", "***")
    
'    nReply = PCB3DL.DlSetCharRaw("PrrTransType", "123450")
    nReply = PCB3DL.DlSetCharRaw("PrrFeeCharge", "手续费:  1.00")
    nReply = PCB3DL.DlSetCharRaw("PrrHostEnqNo", "H-ENQ#:20040921172058490000002")

    PCB3DL.DlSetCharRaw "FitPrrAccNo", "02001000351****4"
    PCB3DL.DlSetCharRaw "GBLBankCode", "1402"
    PCB3DL.DlSetCharRaw "GBLDateYYYYMMDD", "2004/09/21"
    PCB3DL.DlSetCharRaw "GBLTimeHHMM", "17:21"
    PCB3DL.DlSetCharRaw "GBLAtmCode", "14020183"
'    PCB3DL.DlSetCharRaw "PtrReference", "2012"
    PCB3DL.DlSetCharRaw "PrrTransAmount", "RMB 100.00"
    PCB3DL.DlSetCharRaw "PrrTfr2ndAccNo", "609120181320000145"
    PCB3DL.DlSetCharRaw "GBLLineSendNum", "000257"
'    PCB3DL.DlSetCharRaw "PtrCommission", "1234567"
    
'    TxtCmdRc.Text = ""
    T_RowHeight.Text = GetRegKeyS(HKEY_LOCAL_MACHINE, RowHeightPath, "RowHeight", 10, "")

    S3EPrr.Register 0
    nRc = S3EPrr.PuOpen
    If nRc = 0 Then
        Label1.Caption = "收条打印机正常，请测试打印收条！"
        Command2.Enabled = True
        Command1.Enabled = True
    Else
        Label1.Caption = "收条打印机正在初始化，请稍候！"
        Timer1.Enabled = True
    End If
    
'    Label1.Caption = "1064 Receipt Printer PU Open RC: " & nRc
    S3EPrr.Present = True

End Sub

Private Sub S3EPrr_AtPresented()
    Label1.Caption = "请取走收条！"
    S3EPrr.UserReply = 0
End Sub

Private Sub S3EPrr_AtPrintFormEnd(ByVal rc As Integer)
    If rc = 0 Then
        Label1.Caption = "打印收条成功！"
    Else
        Label1.Caption = "打印收条失败，请退出程序并检查收条格式文件是否正确！"
    End If
End Sub

Private Sub S3EPrr_AtPrintFormStart()
    Label1.Caption = "正在打印收条，请稍候！"
    S3EPrr.UserReply = 0
End Sub

Private Sub S3EPrr_BeforePresent()
    S3EPrr.UserReply = 0
End Sub

Private Sub S3EPrr_DevStateChanged()
    If S3EPrr.Available Then
        Label1.Caption = "收条打印机正常，请测试打印收条！"
        Command2.Enabled = True
        Timer1.Enabled = False
'    Else
'        Label1.Caption = "请确认收条打印机是否正常！"
    End If
End Sub

Private Sub Timer1_Timer()
        Timer1.Enabled = False
        Label1.Caption = "请确认收条打印机是否正常！"
End Sub
