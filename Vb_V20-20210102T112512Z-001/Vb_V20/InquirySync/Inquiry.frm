VERSION 5.00
Object = "{B2110643-3E81-11D3-8ACC-00C04FF20A5D}#1.2#0"; "TransProv.dll"
Object = "{192DFCF0-F664-11D3-8BD4-00C04FF20A5D}#1.1#0"; "AdvBrowser.ocx"
Object = "{9C37E835-6A58-11D1-80C0-0020AF7093F9}#1.2#0"; "Dl.ocx"
Object = "{6C4DD4AB-27D5-11D3-96C4-000000000000}#1.0#0"; "S3ELineOutTcp.ocx"
Object = "{E64F71A6-E705-4151-9895-5138B7D67F3A}#1.0#0"; "CHPrj.ocx"
Begin VB.Form ProInquiry 
   Caption         =   "Inquiry"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   Icon            =   "Inquiry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin CHPRJLib.CHPrj SDOPrj 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   1296
      _StockProps     =   1
   End
   Begin TRANSPROVLibCtl.TransactionProvider S3ETrans 
      Height          =   720
      Left            =   1560
      OleObjectBlob   =   "Inquiry.frx":1272
      TabIndex        =   3
      Top             =   240
      Width           =   1545
   End
   Begin S3EADVBROWSERLibCtl.AdvBrowser Browser 
      Height          =   585
      Left            =   1575
      OleObjectBlob   =   "Inquiry.frx":12AC
      TabIndex        =   2
      Top             =   1065
      Width           =   1545
   End
   Begin VB.Timer TimerAction 
      Left            =   4080
      Top             =   1200
   End
   Begin S3ELINEOUTLib.S3ELineOut S3ELineOut 
      Height          =   690
      Left            =   135
      TabIndex        =   1
      Top             =   1005
      Width           =   1380
      _Version        =   65536
      _ExtentX        =   2434
      _ExtentY        =   1217
      _StockProps     =   1
   End
   Begin VB.CommandButton Start 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3210
      TabIndex        =   0
      Top             =   210
      Width           =   1335
   End
   Begin DLLib.DL Pcb3dl 
      Left            =   1560
      Top             =   195
      _Version        =   65538
      _ExtentX        =   2672
      _ExtentY        =   1296
      _StockProps     =   0
   End
End
Attribute VB_Name = "ProInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'All variable need to be declared early
Option Explicit

'版权说明:迪堡公司中国区技术部
'版本号：Agilis power 1.6
'生成日期：2005.8
'作者：汪林（初始版）
'模块功能： 查询模块
'主要函数及其功能
' 全局变量
'================================================================
'修改日志
'<时间>：2005.8.23
'<修改者>：孙世方
'<详细记录>：
'         增加注释；删除无用变量；调整格式
'<时间>：2005.9.07
'<修改者>：孙世方
'<详细记录>：
'         中行客户化
'<时间>：2005.12.20
'<修改者>：孙世方
'<版本号>:1.1.16
'<详细记录>：当前余额前增加符号显示,以前版本显示符号有错位
'<时间>：2005.12.22
'<修改者>：孙世方
'<版本号>:1.2.16
'<详细记录>：增加主机拒绝准备收条打印内容 ，调用DrawInquiryPrr函数
'================================================================
Const NORMAL_MESSAGE        As Integer = 0

Const ReturnOk              As Integer = 212
Const ReturnToMenu          As Integer = 90
Const ReturnHostReject      As Integer = 30
Const ReturnPinNotMatch     As Integer = 31
Const ReturnCommErr         As Integer = 60
Const ReturnTimeout         As Integer = 80
Const ReturnScreenError     As Integer = 81
Const ReturnPressStop       As Integer = 20
Const sGlobalIni            As String = "C:\ATMWosa\Ini\global.ini"

Public Enum ActionType
    DoNothing = 0
    DoInquiryOk = 1
    DoSelectACCMenu = 2
    DoInquiryCommErr = 8
    DoInquiryReject = 9
    DoInquiryPlsWait = 11
    DoScreenError = 97
    DoQuit = 98
    DoError = 99
End Enum
Public NextAction           As ActionType

Private GLnLineNum          As Variant

Dim CardType                As String
Dim nRc                     As Integer
Dim g_sRejectCode           As Variant
Dim g_sPrjLanguage          As String

Private Sub Form_Load()
    Dim sValue As String
    
    sValue = "The version number of " & App.EXEName & ".exe is " & App.Major & "." _
            & App.Minor & ".0." & App.Revision
    
    LogInfo (sValue)
    
    ' Reset the PcB3HtmlBrowser variables
    nRc = Pcb3dl.DlSetCharRaw("HtmlFkeyList", "")
    nRc = Pcb3dl.DlSetCharRaw("HtmlFkeyMap", "3855")
    
    If GetIniS(sGlobalIni, "Bank_Environment", "PrjLanguage", "E") = "E" Then
        g_sPrjLanguage = "E"
    Else
        g_sPrjLanguage = "C"
    End If
    
    S3ETrans.Available = True
    
End Sub

Private Sub S3ETrans_QuitTransaction()
    NextAction = DoQuit
    TimerAction.Interval = 1000
    TimerAction.Enabled = True
End Sub

Private Sub Start_Click()
    
    nRc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
    nRc = Pcb3dl.DlSetCharRaw("HtmlInput2", "")
    
    CardType = Pcb3dl.DlGetCharRaw("FitCardType")
    NextAction = DoSelectACCMenu
    
    TimerAction.Interval = 100
    TimerAction.Enabled = True
End Sub
Private Sub S3ETrans_StartTransaction(ByVal Action As Long)
    
    Start.Enabled = False
    
    nRc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
    nRc = Pcb3dl.DlSetCharRaw("HtmlInput2", "")
    
    CardType = Pcb3dl.DlGetCharRaw("FitCardType")
'    NextAction = DoSelectACCMenu
    NextAction = DoInquiryPlsWait
    TimerAction.Interval = 100
    TimerAction.Enabled = True

End Sub

Private Sub InquiryTotal()
    Dim nTotInquiryNum    As Long

    nTotInquiryNum = Pcb3dl.DlGetInt("TotInquiryNum")
    
    nTotInquiryNum = nTotInquiryNum + 1
    nRc = Pcb3dl.DlSetLong("TotInquiryNum", nTotInquiryNum)

End Sub
Private Sub TimerAction_Timer()
    Dim sSubStData          As String
    Dim bIsTimerAgain       As Boolean
    Dim PrjString           As String
    Dim PrjCHNString        As String
    Dim HostBalance         As String
    Dim HostSignBal         As String
    Dim sHostRejectCard     As String
    
    TimerAction.Enabled = False
    bIsTimerAgain = True
    Select Case NextAction

        Case DoInquiryPlsWait
            nRc = ShowScreenSync(Browser, "Common", "ComPlsWait", sSubStData)
           
            nRc = S3ELineOut.DoSend("INQ", NORMAL_MESSAGE)
            S3ELineOut.GetData "GBLLineNum", GLnLineNum
            Pcb3dl.DlSetCharRaw "GBLLineSendNum", Format(GLnLineNum, "00000")
            PrjString = " " + vbCrLf + _
                        "   **" + "INQ " + Format(Now(), " HH:MM:SS") + " [" + _
                             Format(GLnLineNum, "000000") + "]" + vbCrLf + _
                        "   **ATM CODE:" + Format(Pcb3dl.DlGetCharRaw("GBLAtmCode")) + vbCrLf
            PrjCHNString = " " + vbCrLf + _
                        "   **" + "查询 " + Format(Now(), " HH:MM:SS") + " 流水号：[" + _
                             Format(GLnLineNum, "000000") + "]" + vbCrLf + _
                        "   **ATM号:" + Format(Pcb3dl.DlGetCharRaw("GBLAtmCode")) + vbCrLf
            
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            
            If nRc <> 0 Then
                NextAction = DoInquiryCommErr
            Else
                nRc = S3ELineOut.DoReceive
                
                If nRc = 0 Then
                    
                   g_sRejectCode = Pcb3dl.DlGetCharRaw("HostTransCode")
                    
                    Select Case g_sRejectCode
                    Case "ARP"
                        PrjString = "   **HOST ACCEPT " + vbCrLf + "   **TRANSACTION OK**"
                        PrjCHNString = "   **主机接受" + vbCrLf + "   **交易成功完成**"
                        PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                        
                        '可取款金额
                        HostBalance = Pcb3dl.DlGetCharRaw("HostFundAvail")
                        nRc = Pcb3dl.DlSetCharRaw("HtmlAvailWTHBalance", HostBalance)
                        
                      '当前余额
                        HostBalance = Pcb3dl.DlGetCharRaw("HostCurBal")
                        HostSignBal = Pcb3dl.DlGetCharRaw("HostSignCurBal")
                        If HostSignBal = "D" Or HostSignBal = "-" Then
                             nRc = Pcb3dl.DlSetCharRaw("HtmlBasicBalance", "-" + Right(HostBalance, 11))
                        Else
                             nRc = Pcb3dl.DlSetCharRaw("HtmlBasicBalance", Right(HostBalance, 11))
                        End If
                                                
                        '可用余额
                        HostBalance = Pcb3dl.DlGetCharRaw("HostShAvailBal")
                        HostSignBal = Pcb3dl.DlGetCharRaw("HostSignAvailBal")
                        
                        If HostSignBal = "D" Or HostSignBal = "-" Then
                             nRc = Pcb3dl.DlSetCharRaw("HtmlAvailibleBalance", "-" + Right(HostBalance, 11))
                        Else
                             nRc = Pcb3dl.DlSetCharRaw("HtmlAvailibleBalance", Right(HostBalance, 11))
                        End If
                            
                        '可转帐金额
                        HostBalance = Pcb3dl.DlGetCharRaw("HostShTfrAvail")
                        HostSignBal = Pcb3dl.DlGetCharRaw("SignOfKeepHostBal")
                        
                        If HostSignBal = "+" Then
                             nRc = Pcb3dl.DlSetCharRaw("HtmlAvailTFRBalance", Right(HostBalance, 11))
                        Else
                             nRc = Pcb3dl.DlSetCharRaw("HtmlAvailTFRBalance", "0")
                        End If
                        
                        NextAction = DoInquiryOk
                            
                        Call InquiryTotal
                    Case "ATP"
                        nRc = S3ELineOut.GetData("constHostRejectCode", g_sRejectCode)
                        nRc = Pcb3dl.DlSetCharRaw("ATMPRejectCode", g_sRejectCode)
                        NextAction = DoInquiryReject
                    Case Else
                        NextAction = DoInquiryCommErr
                    End Select
                ElseIf nRc = 97 Then
                    'Host return MAC error,
                    'Set the trickle to download CommKey again in S3EStarter.exe
                    LogError "DoReceive return 97,host return MAC error"
                    nRc = Pcb3dl.DlSetCharRaw("ResetTransKey", "R")
                    NextAction = DoInquiryCommErr
                Else    'doreceive <> 0
                    NextAction = DoInquiryCommErr
                End If  'endif doreceive
            End If     'endif dosend rc = 0
        
        Case DoInquiryCommErr
            nRc = ShowScreenSync(Browser, "Common", "ComCommErr", sSubStData)
            PrjString = "   **NO RESPONSE FROM ATMP. " + vbCrLf + "   **TRANSACTION FAILED**"
            PrjCHNString = "   **主机无响应" + vbCrLf + "   **交易失败**"
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            
            Call SendExceptionMessage(S3ELineOut, Pcb3dl, "64")
            RetToMaster ReturnCommErr
            bIsTimerAgain = False
            
        Case DoInquiryReject
            nRc = ShowScreenSync(Browser, "Common", "ComReject", sSubStData)

            PrjString = "   **HOST REJECT [" + g_sRejectCode + "]" + vbCrLf + _
                                   "   **" + Pcb3dl.DlGetCharRaw("HostRejectEnglish")
            
            PrjCHNString = "   **主机拒绝 [" + g_sRejectCode + "]" + vbCrLf + _
                               "   **" + Pcb3dl.DlGetCharRaw("HostRejectChinese")
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                               
            sHostRejectCard = Pcb3dl.DlGetCharRaw("HostRejectCard")
                
            If sHostRejectCard = "R" Then
                RetToMaster ReturnPinNotMatch
                Exit Sub
            Else
                Call DrawInquiryPrr(Pcb3dl)
                RetToMaster ReturnHostReject
                Exit Sub
            End If
            bIsTimerAgain = False
            
        Case DoInquiryOk
                nRc = ShowScreenSync(Browser, "Inquiry", "InquiryMenu", sSubStData)
            
            If nRc = 0 Then
                 Select Case Browser.SubStData
                    Case "@Continue"
                        RetToMaster ReturnToMenu
                        Exit Sub
                    Case "@stop"
                        RetToMaster ReturnOk
                        Exit Sub
                    Case Else
                        LogError ScreenInfo.Name + " select a impossible function:" + Browser.SubStData
                End Select
            Else
                LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nRc)
                NextAction = DoScreenError
            End If
            
        '???? need modify according to BOC
        Case DoSelectACCMenu
            nRc = ShowScreenSync(Browser, "Inquiry", "SelectCurrType", sSubStData)
            If nRc = 0 Then
                Select Case Browser.SubStData
                  Case "@CNY"
                       nRc = S3ELineOut.SetData("CurrencyCode", "001")
                       NextAction = DoInquiryPlsWait
                    Case "@Continue"
                       nRc = S3ELineOut.SetData("CurrencyCode", "014")
                       NextAction = DoInquiryPlsWait
                    Case Else
                        LogError ScreenInfo.Name + " select a impossible function:" + Browser.SubStData
                        NextAction = DoScreenError
                End Select
            ElseIf nRc = 91 Then
                RetToMaster ReturnTimeout
                Exit Sub
            Else
                LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nRc)
                NextAction = DoScreenError
            End If
        
        Case DoQuit
            Unload ProInquiry
            Exit Sub
                    
        Case DoScreenError
            RetToMaster ReturnScreenError
            bIsTimerAgain = False
            
        Case Else
            LogError "TimerAction next action case error. The next action is:" + _
                CStr(NextAction)
    End Select
    
    If bIsTimerAgain = True Then
        TimerAction.Enabled = True
    End If
End Sub
Private Sub RetToMaster(ByVal S3eRetValue As Integer)
    S3ETrans.Result = S3eRetValue
End Sub

