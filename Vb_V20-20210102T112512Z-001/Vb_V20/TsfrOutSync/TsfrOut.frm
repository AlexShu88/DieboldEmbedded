VERSION 5.00
Object = "{B2110643-3E81-11D3-8ACC-00C04FF20A5D}#1.2#0"; "TransProv.dll"
Object = "{192DFCF0-F664-11D3-8BD4-00C04FF20A5D}#1.1#0"; "AdvBrowser.ocx"
Object = "{9C37E835-6A58-11D1-80C0-0020AF7093F9}#1.2#0"; "Dl.ocx"
Object = "{6C4DD4AB-27D5-11D3-96C4-000000000000}#1.0#0"; "S3ELineOutTcp.ocx"
Object = "{E64F71A6-E705-4151-9895-5138B7D67F3A}#1.0#0"; "CHPrj.ocx"
Begin VB.Form TsfrOut 
   Caption         =   "TsfrOut"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   Icon            =   "TsfrOut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin CHPRJLib.CHPrj SDOPrj 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   1296
      _StockProps     =   1
   End
   Begin VB.TextBox TxtTransDate 
      DataSource      =   "DataTot"
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "0101"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Data DataTot 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Width           =   1380
   End
   Begin TRANSPROVLibCtl.TransactionProvider S3ETrans 
      Height          =   690
      Left            =   1590
      OleObjectBlob   =   "TsfrOut.frx":0E42
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin S3EADVBROWSERLibCtl.AdvBrowser Browser 
      Height          =   375
      Left            =   1545
      OleObjectBlob   =   "TsfrOut.frx":0E7C
      TabIndex        =   2
      Top             =   1155
      Width           =   1710
   End
   Begin VB.Timer TimerAction 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3480
      Top             =   1080
   End
   Begin S3ELINEOUTLib.S3ELineOut S3ELineOut 
      Height          =   705
      Left            =   165
      TabIndex        =   1
      Top             =   885
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   1244
      _StockProps     =   1
   End
   Begin VB.CommandButton Start 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin DLLib.DL Pcb3dl 
      Left            =   195
      Top             =   945
      _Version        =   65538
      _ExtentX        =   2143
      _ExtentY        =   1085
      _StockProps     =   0
   End
End
Attribute VB_Name = "TsfrOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All variable need to be declared before used
Option Explicit

'==========================================================================================
'版权说明:迪堡公司中国区技术部
'版本号：Agilis power 1.6
'生成日期：2005.8
'作者：  汪林(初始版）
'模块功能： 转帐模块
'主要函数及其功能
'全局变量
'       MaxTransferAmount       最大转帐金额
'       str_CurrencyCode        帐户币种        RMB- 2,HKD- 1
'===============================================================
'修改日志
'<时间>：2005.8.23
'<修改者>：孙世方
'<详细记录>：
'         增加注释；删除无用变量；调整格式
'        数据库和文件操作增加on error处理
'================================================================
'<时间>：2005.10.20
'<修改者>：周亮
'<详细记录>：
'÷转帐数据包中帐号域:
'÷帐号长度n<16   16-n +输入帐号 + 000
'           帐号长度n<19   18-n + 输入帐号 + 0
'           帐号长度n=19   输入帐号
'================================================================
'<时间>：2005.10.21
'<修改者>：孙世方
'<详细记录>：中银通卡转帐需要选择帐户类型，其他直接输入帐号
'================================================================
'<时间>：2005.12。9
'<修改者>：孙世方
'<详细记录>：
'  1  修改TfrOutTotal()错误 原来统计值不会累计
'  2  流水打印主机拒绝码
'  3  判断密码错误时变量未赋值
'  4  在主机返回的账号上加上trim函数
'================================================================
'<时间>：[2005.12.12]
'<修改者>：孙世方
'<版本号>:1.2.16
'<详细记录>： 修改TfrOutTotal, 增加记录CutOff.ini，用于主机cutoff时打印流水
'             修改函数DrawTransferPrr，主机拒绝增加调用
'================================================================

Const NORMAL_MESSAGE                        As Integer = 0
Const sGlobalIni                            As String = "C:\ATMWosa\Ini\global.ini"
Const CutOffIni                             As String = "c:\ATMWosa\Ini\CutOff.ini"

Private Type CardTypeRec
    TrackToMatch   As Integer
    offset         As Integer
    Length         As Integer
    MatchChars     As String
    PinLength      As Integer
    PinMaxAttempts As Integer
    cardtype       As String
    AccNum_Track   As Integer
    AccNum_Len     As Integer
    AccNum_Offset  As Integer
  End Type
Private CardIdx() As CardTypeRec

Public Enum pageType
    pageNothing = 0
    pageTsfrOutInput1 = 1
    pageTsfrOutInput2 = 2
    pageTsfrOutInputDiff = 4
    pageTsfrOutInputAmt = 5
    pageTsfrOutAmtInputErr = 6
    pageTsfrOutCommErr = 7
    pageTsfrOutReject = 8
    pageTsfrOutOk = 10
    pageTsfrOutPlsWait = 11
    pageTsfrOutPressStop = 12
    pageTsfrOutConfirm = 14
'    pageSelectCurrType = 19
    pageSelectAccType = 19
    pageTsfrOutConfAcc = 27
    pageTsfrOutInputAccError1 = 21
    pageTsfrOutInputAccError2 = 22
    pageScreenError = 97
    pageError = 99
    pageQuit = 98
End Enum
Private currentPage As pageType

Const ReturnOk              As Integer = 202
Const ReturnToMenu          As Integer = 232
Const ReturnPressStop       As Integer = 20
Const ReturnPinNotMatch     As Integer = 31
Const ReturnHostReject      As Integer = 30
Const ReturnCommErr         As Integer = 60
Const ReturnTimeout         As Integer = 80

Dim cardtype                As String

Dim GLsInput1               As String
Dim GlsInput2               As String
Dim nrc                     As Integer
Dim g_sHostRespCode         As Variant
Dim TsfInCardType           As String
Const iniPath               As String = "C:\atmwosa\ini\"

Dim MaxTransferAmount       As Long
Dim HostSeq                 As String
Dim str_CurrencyCode        As String
Dim g_sPrjLanguage          As String
'==========================================================================================
'函 数 的 功 能 ：VB窗口装载,初始化功能键，打开数据库
'输 入 参 数   ：无
'输 出 参 数   ：无
'返 回 值      ：无
'作 者　　　　 :
'创 建 时 间   :
'==========================================================================================
Private Sub Form_Load()
    Dim sValue As String
    
    sValue = "The version number of " & App.EXEName & ".exe is " & App.Major & "." _
            & App.Minor & ".0." & App.Revision
    
    LogInfo (sValue)

    nrc = Pcb3dl.DlSetCharRaw("HtmlFkeyList", "")
    nrc = Pcb3dl.DlSetCharRaw("HtmlFkeyMap", "3855")
    
    S3ETrans.Available = True
    
    Call InitCardType
    
    If GetIniS(sGlobalIni, "Bank_Environment", "PrjLanguage", "E") = "E" Then
        g_sPrjLanguage = "E"
    Else
        g_sPrjLanguage = "C"
    End If
    
End Sub
'==========================================================================================
'函 数 的 功 能 ：转帐程序退出
'输 入 参 数   ：无
'输 出 参 数   ：无
'返 回 值      ：无
'作 者
'创 建 时 间   :
'==========================================================================================
Private Sub S3ETrans_QuitTransaction()
    currentPage = pageQuit
    TimerAction.Interval = 1000
    TimerAction.Enabled = True
End Sub

'==========================================================================================
'函 数 的 功 能 ：转帐程序入口
'输 入 参 数   ：在ruler调用此模块时Action值
'输 出 参 数   ：无
'返 回 值      ：无
'作 者
'创 建 时 间   :
'==========================================================================================
Private Sub S3ETrans_StartTransaction(ByVal Action As Long)

    Start.Enabled = False
    If Action = 1 Then
                
        MaxTransferAmount = Pcb3dl.DlGetCharRaw("IcbcMaxTfrAmount")
        
        nrc = Pcb3dl.DlSetCharRaw("Tfr2ndAccNo", " ")
        nrc = Pcb3dl.DlSetCharRaw("GBLAmount", "")
        nrc = Pcb3dl.DlSetCharRaw("GBLPrtAmount", "*********.**")
        nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
        nrc = Pcb3dl.DlSetCharRaw("HtmlInput2", "")
        
        cardtype = Pcb3dl.DlGetCharRaw("FitCardType")
        
        If cardtype = "03" Or cardtype = "04" Then
           str_CurrencyCode = "1"
           currentPage = pageSelectAccType
        Else
           str_CurrencyCode = "2"
           currentPage = pageTsfrOutInput1
        End If
    Else
        currentPage = pageTsfrOutPlsWait
    End If
    
    TimerAction.Enabled = True
End Sub

Private Sub Start_Click()
    Dim cardtype As String
    
    str_CurrencyCode = "001"
        MaxTransferAmount = "02000000000"
        nrc = Pcb3dl.DlSetCharRaw("Tfr2ndAccNo", " ")
        nrc = Pcb3dl.DlSetCharRaw("GBLAmount", "")
        nrc = Pcb3dl.DlSetCharRaw("GBLPrtAmount", "*********.**")
        nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
        nrc = Pcb3dl.DlSetCharRaw("HtmlInput2", "")
        
     cardtype = "01"
     If cardtype = "03" Or cardtype = "04" Then
           str_CurrencyCode = "002"
           currentPage = pageSelectAccType
        Else
           str_CurrencyCode = "001"
           currentPage = pageTsfrOutInput1
        End If
    TimerAction.Enabled = True

End Sub

Private Sub TimerAction_Timer()
    Dim sSubStData           As String
    Dim sCurrentDate         As String
    Dim bIsTimerAgain        As Boolean
    Dim PrjString            As String
    Dim PrjCHNString         As String
    Dim sTfrAmount           As String
    Dim dblTfrAmount         As Double
    Dim acc_no               As String
    Dim HostAccNo            As String
    Dim vGlnLineNum          As Variant
    Dim Fee                  As String
    Dim PrrFee               As String
    Dim ServiceNbr           As String
    Dim TsfrAmount           As String
    Dim sDetailRec           As String
    Dim cardflag             As String
    Dim CardInternalAccNo    As String
    Dim subacc               As Variant
    Dim TsfOutCardType       As Variant
    Dim sHostRejectCard      As String
    Dim TsfOutSubAcc         As Variant
    Dim spadding             As String
    Dim HostRejCode          As Variant
    
    TimerAction.Enabled = False
    bIsTimerAgain = True

    Select Case currentPage
        '增加对输入帐号卡位数的判断
        Case pageTsfrOutInput1
            nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
            
            nrc = ShowScreenSync(Browser, "TsfrOut", "TsfrOutInput1", sSubStData)
            If nrc = 0 Then
                Select Case Browser.SubStData
                    Case "@ok"
                        GLsInput1 = Pcb3dl.DlGetCharRaw("HtmlInput1")
                       If Len(GLsInput1) < 11 Or Len(GLsInput1) > 19 Then
                            If Pcb3dl.DlGetCharRaw("GBLSelectLan") = "ENG" Then
                             nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", _
                                            "Wrong Account Number")
                           
                            Else
                            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", _
                                            "输入卡号的位数不正确")
                            End If
                            nrc = ShowScreenSync(Browser, "TsfrOut", "TsfrOutInputError", _
                                                    sSubStData)
                            currentPage = pageTsfrOutInput1
                        Else
                            currentPage = pageTsfrOutInput2
                        End If
                       
                    Case "@Update"
                         currentPage = pageTsfrOutInput1
                    Case "@stop"
                         RetToMaster ReturnToMenu
                         Exit Sub
                    Case Else
                         LogError ScreenInfo.Name + " select a impossible function:" + Browser.SubStData
                End Select
            ElseIf nrc = 91 Then
                RetToMaster ReturnTimeout
                Exit Sub
            Else
                LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nrc)
                currentPage = pageScreenError
            End If
            
    
        Case pageTsfrOutInput2
            nrc = Pcb3dl.DlSetCharRaw("HtmlInput2", "")
        
            nrc = ShowScreenSync(Browser, "TsfrOut", "TsfrOutInput2", sSubStData)
            If nrc = 0 Then
                Select Case Browser.SubStData
                    Case "@ok"
                       GlsInput2 = Pcb3dl.DlGetCharRaw("HtmlInput2")
                       If GLsInput1 <> GlsInput2 Then
                            If Pcb3dl.DlGetCharRaw("GBLSelectLan") = "ENG" Then
                             nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", _
                                            "Wrong Account Number")
                           
                            Else
                            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", _
                                                    "两次输入卡号不一致")
                            End If
                            nrc = ShowScreenSync(Browser, "TsfrOut", "TsfrOutInputError", sSubStData)

                            currentPage = pageTsfrOutInput1
                       Else
                            spadding = "00000"
                            
'                            If Len(GLsInput1) < 16 Then
'                                GLsInput1 = Right(spadding, 16 - Len(GLsInput1)) + GLsInput1 + "000"
'                            ElseIf Len(GLsInput1) < 19 Then
'                                GLsInput1 = Right(spadding, 18 - Len(GLsInput1)) + GLsInput1 + "0"
'                            End If
                            If Len(GLsInput1) < 19 Then
                               GLsInput1 = Right(spadding, 18 - Len(GLsInput1)) + GLsInput1 + Space(1)
                                ' GLsInput1 = GLsInput1 + Space(19 - Len(GLsInput1))
                            End If
                            nrc = Pcb3dl.DlSetCharRaw("Tfr2ndAccNo", GLsInput1)

                            currentPage = pageTsfrOutInputAmt
                      End If
                    Case "@stop"
                        RetToMaster ReturnToMenu
                        Exit Sub
                    Case "@Update"
                         currentPage = pageTsfrOutInput2
                    Case Else
                         LogError ScreenInfo.Name + " select a impossible function:" + Browser.SubStData
                End Select
            ElseIf nrc = 91 Then
                RetToMaster ReturnTimeout
                Exit Sub
            Else
                LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nrc)
                currentPage = pageScreenError
            End If
               
        Case pageTsfrOutPressStop
            nrc = ShowScreenSync(Browser, "Common", "ComPressStop", sSubStData)
            RetToMaster ReturnPressStop
            Exit Sub
            
        Case pageTsfrOutInputDiff
            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", _
                            "对不起，您两次输入的转入帐号不一致")
            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt2", _
                            "Sorry,the two accounts are unmatched.")
            nrc = ShowScreenSync(Browser, "TsfrOut", "TsfrOutInputError", sSubStData)
            currentPage = pageTsfrOutInput1
        
        Case pageTsfrOutInputAmt
            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", MaxTransferAmount)
            nrc = Pcb3dl.DlSetCharRaw("GBLAmount", "")
            nrc = ShowScreenSync(Browser, "TsfrOut", "TsfrOutAmtInput", sSubStData)
        
            If nrc = 0 Then
                 Select Case Browser.SubStData
                    Case "@ok"
                        sTfrAmount = Pcb3dl.DlGetCharRaw("GBLAmount")
                        If Len(sTfrAmount) = 0 Or sTfrAmount = "." Then
                            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", _
                                            "输入金额错误")
                            currentPage = pageTsfrOutAmtInputErr
                        ElseIf CDbl(sTfrAmount) = 0 Then
                            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", _
                                            "输入转帐金额不能为零")
                            currentPage = pageTsfrOutAmtInputErr

                        ElseIf CDbl(sTfrAmount) > CDbl(MaxTransferAmount) Then
                            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", _
                                            "输入金额超过转帐限额")
                            currentPage = pageTsfrOutAmtInputErr
                            
                        Else
                            dblTfrAmount = CDbl(sTfrAmount)
                        
                            nrc = Pcb3dl.DlSetCharRaw("GBLPrtAmount", _
                                Format(dblTfrAmount, "Standard"))
                        
                            dblTfrAmount = dblTfrAmount * 100
                            sTfrAmount = Format(dblTfrAmount, "00000000")
                            nrc = Pcb3dl.DlSetCharRaw("GBLAmount", sTfrAmount)
                            currentPage = pageTsfrOutConfirm
                        End If
                    
                    Case "@Update"
                         currentPage = pageTsfrOutInputAmt
                    Case "@stop"
                        RetToMaster ReturnToMenu
                        Exit Sub
                     
                     Case Else
                         LogError ScreenInfo.Name + " select a impossible function:" + Browser.SubStData
                End Select
            ElseIf nrc = 91 Then
                RetToMaster ReturnTimeout
                Exit Sub
            Else
                LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nrc)
                currentPage = pageScreenError
            End If
                
        Case pageTsfrOutConfirm
                nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", Pcb3dl.DlGetCharRaw("FitAccNo"))
                nrc = ShowScreenSync(Browser, "TsfrOut", "TsfConfirm", sSubStData)
            
            If nrc = 0 Then
                Select Case Browser.SubStData
                    Case "@ok"
                        currentPage = pageTsfrOutPlsWait
                     Case "@stop"
                        RetToMaster ReturnToMenu
                        Exit Sub
                    Case "@modify"
                        nrc = Pcb3dl.DlSetCharRaw("GBLAmount", "")
                        currentPage = pageTsfrOutInput1
                                        
                    Case Else
                         LogError ScreenInfo.Name + " select a impossible function:" + Browser.SubStData
                End Select
            ElseIf nrc = 91 Then
                RetToMaster ReturnTimeout
                Exit Sub
            Else
                LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nrc)
                currentPage = pageScreenError
            End If
                
        
        Case pageTsfrOutAmtInputErr
            nrc = ShowScreenSync(Browser, "TsfrOut", "TsfrOutInputError", sSubStData)

            currentPage = pageTsfrOutInputAmt
            
        Case pageTsfrOutPlsWait
            
            nrc = ShowScreenSync(Browser, "Common", "ComPlsWait", sSubStData)
                         
            sCurrentDate = Format(Now(), "MMDDHHMMSS")
            nrc = S3ELineOut.SetData("TransCurr", str_CurrencyCode)
            nrc = S3ELineOut.DoSend("TFR", NORMAL_MESSAGE)
            
            nrc = S3ELineOut.GetData("GBLLineNum", vGlnLineNum)
            nrc = Pcb3dl.DlSetCharRaw("GBLLineSendNum", _
                                Format(vGlnLineNum, "000000"))
            
            ServiceNbr = Pcb3dl.DlGetCharRaw("Tfr2ndAccNo")
            TsfrAmount = Pcb3dl.DlGetCharRaw("GBLPrtAmount")
            
            PrjString = " " + vbCrLf + _
                        "   " + "TFR " + Format(Now(), " HH:MM:SS") + " [" + Format(vGlnLineNum, "0000") + "]" + vbCrLf + _
                        "    ATM CODE: " + Format(Pcb3dl.DlGetCharRaw("GBLAtmCode")) + vbCrLf + _
                        "    TSF PAN:" + ServiceNbr + vbCrLf + _
                        "    Amount:" + TsfrAmount + vbCrLf
            
            PrjCHNString = " " + vbCrLf + _
                        "    转帐 " + Format(Now(), " HH:MM:SS") + " 流水号：[" + Format(vGlnLineNum, "000000") + "]" + vbCrLf + _
                        "    ATM号： " + Format(Pcb3dl.DlGetCharRaw("GBLAtmCode")) + vbCrLf + _
                        "    转入帐号: " + ServiceNbr + vbCrLf + _
                        "    金额： " + TsfrAmount + vbCrLf
                         ' modi by nicktan
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    
            If nrc <> 0 Then
                currentPage = pageTsfrOutCommErr

            Else
                nrc = S3ELineOut.DoReceive
                If nrc = 0 Then
                    g_sHostRespCode = Pcb3dl.DlGetCharRaw("HostTransCode")
                    If g_sHostRespCode = "AQP" Then
                        
                        HostSeq = Pcb3dl.DlGetCharRaw("IcbcHostSeq")
                        HostAccNo = Trim(Pcb3dl.DlGetCharRaw("HostAccNo"))
                        
                        Pcb3dl.DlSetCharRaw "FitPrrAccNo", _
                        Left(HostAccNo, Len(HostAccNo) - 5) + "****" + Right(HostAccNo, 1)
                               
                        
                        PrjString = "   **HOST ACCEPT " + vbCrLf + _
                           "     Host AccNo: " + HostAccNo + _
                           "     host CardMark: " + Pcb3dl.DlGetCharRaw("FitCardMark") + vbCrLf + _
                           "     Host Date: " + Pcb3dl.DlGetCharRaw("IcbcHostTime") + vbCrLf + _
                           "     Host Fee :" + PrrFee + vbCrLf + _
                           "     Host Seq :" + HostSeq + vbCrLf + _
                           "     TRANSACTION OK"
                        
                        PrjCHNString = "   主机接受 " + vbCrLf + _
                                     "     主机返回帐号：" + HostAccNo + _
                                     "     卡片标识： " + Pcb3dl.DlGetCharRaw("HostCurrentDate") + vbCrLf + _
                                     "     主机时间： " + Pcb3dl.DlGetCharRaw("HostCurrentDate") + vbCrLf + _
                                     "     手续费： " + PrrFee + vbCrLf + _
                                     "     主机检索号：" + HostSeq + vbCrLf + _
                                     "     交易成功完成"
                           ' delete the "*" by nicktan   增加打印卡标识
                        PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                                     
                            CardInternalAccNo = Pcb3dl.DlGetCharRaw("HtmlInput2")
                            nrc = Pcb3dl.DlSetCharRaw("Tfr2ndAccNo", CardInternalAccNo)

                        currentPage = pageTsfrOutOk

                    Else
                        currentPage = pageTsfrOutReject

                    End If
                ElseIf nrc = 97 Then
                    'Host return MAC error,
                    'Set the trickle to download CommKey again in S3EStarter.exe
                    LogError "DoReceive return 97,host return MAC error"
                    nrc = Pcb3dl.DlSetCharRaw("ResetTransKey", "R")
                    currentPage = pageTsfrOutCommErr
                Else
                    LogWarning "S3ELineOut.Receive Return " + CStr(nrc)
                    currentPage = pageTsfrOutCommErr
                End If 'endif do receive
            End If 'endif dosend
        Case pageTsfrOutCommErr
            nrc = ShowScreenSync(Browser, "Common", "ComCommErr", sSubStData)
            Call SendExceptionMessage(S3ELineOut, Pcb3dl, "64")
            PrjString = "   **" + "NO RESPONSE FROM ATMP " + vbCrLf + _
                                   "   TRANSACTION FAILED"
            PrjCHNString = "   **主机无响应" + vbCrLf + "   交易失败"
            ' modi by nicktan
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            RetToMaster ReturnCommErr
            Exit Sub
            
        Case pageTsfrOutOk
            nrc = ShowScreenSync(Browser, "TsfrOut", "TsfrOutOk", sSubStData)
            
            Pcb3dl.DlSetCharRaw "GBLATMLocRejCode", "   "
            
            Call ResetATMPrr(TsfrOut.Pcb3dl)
            Call DrawTransferPrr(Pcb3dl, PrrOK)
            Call TfrOutTotal

            RetToMaster ReturnOk
            Exit Sub
        
        Case pageTsfrOutReject
            nrc = ShowScreenSync(Browser, "Common", "ComReject", sSubStData)
            
            nrc = S3ELineOut.GetData("constHostRejectCode", HostRejCode)
            
            '流水打印主机拒绝码 2005.12.9
            PrjString = "   **HOST REJECT [" + HostRejCode + "]" + vbCrLf + _
                                   "   " + Pcb3dl.DlGetCharRaw("HostRejectEnglish")
            PrjCHNString = "   **主机拒绝 [" + HostRejCode + "]" + vbCrLf + _
                               "   " + Pcb3dl.DlGetCharRaw("HostRejectChinese")
             
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            
            '原来此处判断有误 2005.12.9
            sHostRejectCard = Pcb3dl.DlGetCharRaw("HostRejectCard")
            If sHostRejectCard = "R" Then
                RetToMaster ReturnPinNotMatch
                Exit Sub
            Else
                Call DrawTransferPrr(Pcb3dl, PrrReject)   '2005。12。22 打印拒绝收条
                RetToMaster ReturnHostReject
                Exit Sub
            End If
     
     'add by nicktan 增加对中银卡的转入帐号匹配
        Case pageSelectAccType
            nrc = ShowScreenSync(Browser, "TsfrOut", "SelectAccType", sSubStData)
            If nrc = 0 Then
                Select Case Browser.SubStData
                    Case "@saving"
                         nrc = Pcb3dl.DlSetCharRaw("Tfr2ndAccNo", "0000000000000010   ")

                    Case "@check"
                         nrc = Pcb3dl.DlSetCharRaw("Tfr2ndAccNo", "0000000000000020   ")
                        
                    Case "@credit"
                         nrc = Pcb3dl.DlSetCharRaw("Tfr2ndAccNo", "0000000000000030   ")
                    
                    Case "@default"
                         nrc = Pcb3dl.DlSetCharRaw("Tfr2ndAccNo", "0000000000000000   ")
                         
                    Case "@other"
                         currentPage = pageTsfrOutInput1
                    Case Else
                         LogError ScreenInfo.Name + " select a impossible function:" + Browser.SubStData
                End Select
            ElseIf nrc = 91 Then
                RetToMaster ReturnTimeout
                Exit Sub
            Else
                LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nrc)
                currentPage = pageScreenError
            End If
        
        Case pageScreenError
            bIsTimerAgain = False
            
        Case pageQuit
            Unload TsfrOut
            Exit Sub
            
        Case Else
            LogError "TimerAction next action case error. The next action is:" + _
                CStr(currentPage)
    End Select
    
    If bIsTimerAgain = True Then
        TimerAction.Enabled = True
    End If

End Sub

Private Sub RetToMaster(ByVal S3eRetValue As Integer)
    S3ETrans.Result = S3eRetValue
End Sub
'==========================================================================================
'函 数 的 功 能 :记录转帐成功统计值
'输 入 参 数   ：无
'输 出 参 数   ：无
'返 回 值      ：无
'调 用 函 数   ：无
'被 调 用 情 况：
'作 者         ：汪林
'创 建 时 间   :
'<时间>：2005.12。9
'<修改者>：孙世方
'<详细记录>：修改错误 原来统计值不会累计
'==========================================================================================
'<时间>：[2005.12.12]
'<修改者>：孙世方
'<详细记录>：
'    增加记录CutOff.ini，用于主机cutoff时打印流水

Private Sub TfrOutTotal()
    Dim nTotTfrOutNum         As Long
    Dim dblTotTfrOutAmt       As Double
    Dim sGBLAmount            As String
    Dim num                   As Integer
    Dim szMsg                 As String
    Dim CutOffTfrAmount       As String
    Dim CutOffTfrNumber       As String
    Dim TempNumber            As Long
    Dim TempAmount            As Double
    
On Error GoTo ErrHandler

    nTotTfrOutNum = Pcb3dl.DlGetInt("TotTfrOutNum")
    dblTotTfrOutAmt = Pcb3dl.DlGetDouble("TotTfrOutAmount")
    sGBLAmount = Pcb3dl.DlGetCharRaw("GBLAmount")
    
    nTotTfrOutNum = nTotTfrOutNum + 1
    nrc = Pcb3dl.DlSetLong("TotTfrOutNum", nTotTfrOutNum)

    '原来此处写错 ！ If Len(sGBLAmount) = 0 Then 统计值将不会累计，应使用/符号，不用\
    If Len(sGBLAmount) <> 0 Then
        dblTotTfrOutAmt = dblTotTfrOutAmt + CDbl(sGBLAmount) / 100
        nrc = Pcb3dl.DlSetDouble("TotTfrOutAmount", _
                dblTotTfrOutAmt)
    End If
    
    CutOffTfrNumber = GetIniS(CutOffIni, "HostCutOff", "TfrNumber", "0")
    TempNumber = CLng(CutOffTfrNumber) + 1
    CutOffTfrAmount = GetIniS(CutOffIni, "HostCutOff", "TfrAmount", "0")
    If TempNumber > 2000 Then
        LogError ("TransferNumber Too High ,now Clear")
        TempNumber = 0
        CutOffTfrAmount = "0"
        sGBLAmount = "0"
    End If
    If Len(sGBLAmount) <> 0 Then
        TempAmount = CDbl(CutOffTfrAmount) + CDbl(sGBLAmount) / 100
        nrc = SetIniS(CutOffIni, "HostCutOff", "TfrAmount", CStr(TempAmount))
    End If
    nrc = SetIniS(CutOffIni, "HostCutOff", "TfrNumber", CStr(TempNumber))
    
    Exit Sub
ErrHandler:
    szMsg = CStr(Err.Number) + ": " + Err.Description + " in TfrOutTotal"
    LogError szMsg
    Err.Clear
    
End Sub
'===================================================================================
'函数功能 :检查卡是否可接收（是否在卡表中定义）
'输入参数 ：无
'输出参数：无
'返回值：无
'调用函数：
'被调用情况：
'作者：
'创建时间 : 2004
'====================================================================================
Private Sub CheckCard()
    Dim i                 As Integer
    Dim times             As Integer
    Dim find              As Boolean
    Dim bmatch            As Boolean
    Dim StrmatchChars     As String
    Dim sCardType         As String
    Dim StrTrack          As String
    
    find = False
    times = 1
    
    For i = 0 To UBound(CardIdx)
        If times = 1 Then
            If Len(GLsInput1) < 19 Then
                StrTrack = "0162" + GLsInput1
            Else
                StrTrack = GLsInput1
            End If
        ElseIf times = 2 Then
            If Len(GLsInput1) < 19 Then
                StrTrack = "62" + GLsInput1
            Else
                StrTrack = "01" + GLsInput1
            End If
        Else
            StrTrack = GLsInput1
        End If
         
        StrmatchChars = Mid(StrTrack, CardIdx(i).offset, CardIdx(i).Length)
        bmatch = MatchChars_Compare(StrmatchChars, CardIdx(i).MatchChars, CardIdx(i).Length)
         If bmatch = True Then
            sCardType = CardIdx(i).cardtype
    
            TsfInCardType = Right(sCardType, 2)
            
            find = True
            Exit For
        Else
            If times = 1 Then
                times = 2
                i = i - 1
            ElseIf times = 2 Then
                times = 3
                i = i - 1
            Else
               times = 1
            End If
        End If
    Next
    
    If find = False Then
        TsfInCardType = "000"
    End If
End Sub

Private Function MatchChars_Compare(pChar1 As String, pChar2 As String, lenght As Integer) As Boolean
    Dim strOne1     As String
    Dim strOne2     As String
    Dim i           As Integer
    
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

Sub InitCardType()
    Dim i As Integer
    Dim j As Integer
    
    j = GetIniN(iniPath + "fit.ini", "General", "CurrentRecord", 0)
    
    If j < 1 Then
        LogError "Initcialize fit.ini Error!"
    Else
        ReDim CardIdx(j - 1)
        For i = 0 To j - 1
            CardIdx(i).Length = GetIniN(iniPath + "fit.ini", "CardIndex" + LTrim(CStr(i)), "Length", 0)
            CardIdx(i).MatchChars = GetIniS(iniPath + "fit.ini", "CardIndex" + LTrim(CStr(i)), "MatchChars", "")
            CardIdx(i).cardtype = GetIniS(iniPath + "fit.ini", "CardIndex" + LTrim(CStr(i)), "CardType", 0)
            CardIdx(i).offset = 1 + GetIniN(iniPath + "fit.ini", "CardIndex" + LTrim(CStr(i)), "Offset", 0)
        Next
       
    End If
End Sub

