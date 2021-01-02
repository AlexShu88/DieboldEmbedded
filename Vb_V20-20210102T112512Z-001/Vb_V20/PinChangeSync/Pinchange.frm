VERSION 5.00
Object = "{B2110643-3E81-11D3-8ACC-00C04FF20A5D}#1.2#0"; "TransProv.dll"
Object = "{EACE4ECF-6930-11D0-AC6C-0020AFBDD1D4}#1.0#0"; "SDOEdm.ocx"
Object = "{192DFCF0-F664-11D3-8BD4-00C04FF20A5D}#1.1#0"; "AdvBrowser.ocx"
Object = "{7CCB2EF0-B3E8-11CF-BF8E-0020AF7093F9}#1.0#0"; "SDOPin.ocx"
Object = "{9C37E835-6A58-11D1-80C0-0020AF7093F9}#1.2#0"; "Dl.ocx"
Object = "{6C4DD4AB-27D5-11D3-96C4-000000000000}#1.0#0"; "S3ELineOutTcp.ocx"
Object = "{E64F71A6-E705-4151-9895-5138B7D67F3A}#1.0#0"; "CHPrj.ocx"
Begin VB.Form PinChange 
   Caption         =   "PinChange"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4740
   Icon            =   "Pinchange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin S3EADVBROWSERLibCtl.AdvBrowser Browser 
      Height          =   495
      Left            =   240
      OleObjectBlob   =   "Pinchange.frx":1272
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin CHPRJLib.CHPrj S3EPrj 
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   1
   End
   Begin SDOPinLibCtl.SDOPin S3EPin 
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
      _cx             =   2355
      _cy             =   450
      RelativeSpaceSize=   20
      ForeColor2      =   0
      TimeOutSecondsFirst=   30
      TimeOutSecondsSecond=   20
      TimeOutSecondsLast=   0
      ActiveKeyPair1  =   0
      ActiveKeyPair2  =   0
      ActiveKeyPair3  =   0
      ActiveKeyPair4  =   0
      ActiveKeyPair5  =   0
      ActiveKeyPair6  =   0
      ActiveKeyPair7  =   0
      ActiveKeyPair8  =   0
      ActiveKeyPair9  =   0
      ActiveKeyPair10 =   0
      ActiveKeyPair11 =   0
      ActiveKeyPair12 =   0
      ActiveKeyPair13 =   0
      ActiveKeyPair14 =   0
      ActiveKeyPair15 =   0
      FireScreenClass =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Transparent     =   -1  'True
      UseAmbientFont  =   -1  'True
   End
   Begin TRANSPROVLibCtl.TransactionProvider S3ETrans 
      Height          =   720
      Left            =   1755
      OleObjectBlob   =   "Pinchange.frx":1298
      TabIndex        =   8
      Top             =   120
      Width           =   1245
   End
   Begin S3ELINEOUTLib.S3ELineOut S3ELineOut 
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   885
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   1296
      _StockProps     =   1
   End
   Begin DLLib.DL Pcb3dl 
      Left            =   1725
      Top             =   915
      _Version        =   65538
      _ExtentX        =   2302
      _ExtentY        =   1164
      _StockProps     =   0
   End
   Begin SDOEdmLibCtl.SDOEdm S3EEdm 
      Height          =   570
      Left            =   255
      OleObjectBlob   =   "Pinchange.frx":12D6
      TabIndex        =   6
      Top             =   120
      Width           =   1245
   End
   Begin VB.TextBox DesKey 
      Height          =   375
      Left            =   2970
      TabIndex        =   4
      Text            =   "EFEFEFEFEFEFEFEF"
      Top             =   1740
      Width           =   1575
   End
   Begin VB.OptionButton Software 
      Caption         =   "Software"
      Height          =   330
      Left            =   3255
      TabIndex        =   3
      Top             =   975
      Width           =   1095
   End
   Begin VB.OptionButton Hardware 
      Caption         =   "Hardware"
      Height          =   240
      Left            =   3255
      TabIndex        =   2
      Top             =   1320
      Value           =   -1  'True
      Width           =   1080
   End
   Begin VB.CommandButton Start 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3210
      TabIndex        =   0
      Top             =   135
      Width           =   1290
   End
   Begin VB.Timer TimerAction 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4005
      Top             =   990
   End
   Begin VB.Label Label1 
      Caption         =   "S.W. Key"
      Height          =   240
      Left            =   2160
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
End
Attribute VB_Name = "PinChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All variable need to be declared early
Option Explicit

'==========================================================================================
'版权说明:迪堡公司中国区技术部
'版本号：Agilis power 1.6
'生成日期：2005.8
'作者：  汪林(初始版）
'模块功能： 改密模块
'主要函数及其功能
' 全局变量
'修改日志
'-----------------------------------------------------------------------
'<时间>：2005.12.23
'<修改者>：孙世方
'版本号:1.2.6
'<详细记录>：
'   1 流水打印主机拒绝码
'   2 修正一个bug，两次输入密码时直接按屏幕上的确认键，在PIN_Difference会报错，导致程序异常！
'==========================================================================================

Const NORMAL_MESSAGE        As Integer = 0

Const ReturnOk              As Integer = 211
Const ReturnPressStop       As Integer = 20
Const ReturnHostReject      As Integer = 30
Const ReturnPinNotMatch     As Integer = 31
Const ReturnCommErr         As Integer = 60
Const ReturnTimeout         As Integer = 80
Const ReturnInputOverLimit  As Integer = 110

Const sGlobalIni = "C:\ATMWosa\Ini\global.ini"

Public Enum pageType
    pageNothing = 0
    pagePinChangeInput1 = 1
    pagePinChangeInput2 = 2
    pagePinChangeDiff = 3
    pagePinChangeError = 4
    pagepinchangeOk = 5
    pagePinChangeTimeOut = 7
    pagePinChangeCommErr = 8
    pagePinChangeReject = 9
    pagePinChangePressStop = 10
    pagePinChangePlsWait = 11
    pagePinChangeEppInput1 = 12
    pagePinChangeEppInput2 = 13
    pagePinChangeNotSix1 = 14
    pagePinChangeNotSix2 = 15
    pageScreenError = 97
    pageError = 99
    pageQuit = 98
End Enum
Public currentPage          As pageType
Private GLnLineNum          As Variant

Dim nrc                     As Integer
Dim g_sHostRespCode         As Variant

Dim g_strAccNo20            As String
Dim g_vUnShuffleRjCode      As Variant
Dim G_bIsHardware           As Boolean
Dim g_sPrjLanguage          As String
Dim g_sRejectCode           As Variant

Private Sub Form_Load()
    Dim sValue As String
    
    sValue = "The version number of " & App.EXEName & ".exe is " & App.Major & "." _
            & App.Minor & ".0." & App.Revision
    
    LogInfo (sValue)
    
    ' Reset the PcB3HtmlBrowser variables
    nrc = Pcb3dl.DlSetCharRaw("HtmlFkeyList", "")
    nrc = Pcb3dl.DlSetCharRaw("HtmlFkeyMap", "3855")
    
    If GetIniS(sGlobalIni, "Bank_Environment", "PrjLanguage", "E") = "E" Then
        g_sPrjLanguage = "E"
    Else
        g_sPrjLanguage = "C"
    End If
    
    S3ETrans.Available = True

End Sub

Private Sub S3ETrans_QuitTransaction()
    currentPage = pageQuit
    TimerAction.Interval = 1000
    TimerAction.Enabled = True
End Sub
Private Sub Start_Click()
    g_strAccNo20 = Pcb3dl.DlGetCharRaw("FitAccNo")
    nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
    nrc = Pcb3dl.DlSetCharRaw("HtmlInput2", "")
    nrc = Pcb3dl.DlSetCharRaw("PinChangeRetry", "3")
    
    currentPage = pagePinChangeInput1
    TimerAction.Enabled = True
End Sub
Private Sub S3ETrans_StartTransaction(ByVal Action As Long)

    Start.Enabled = False
    
    g_strAccNo20 = Pcb3dl.DlGetCharRaw("FitAccNo")
    
    nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
    nrc = Pcb3dl.DlSetCharRaw("HtmlInput2", "")
    nrc = Pcb3dl.DlSetCharRaw("PinChangeRetry", "3")
    
    If Pcb3dl.DlGetCharRaw("GBLEncrypType") = "H" Then
        G_bIsHardware = True
        currentPage = pagePinChangeEppInput1
    Else
        G_bIsHardware = False
        currentPage = pagePinChangeInput1
    End If
    TimerAction.Enabled = True
End Sub
'==========================================================================================
'函 数 的 功 能 :记录改密成功总笔数统计值
'输 入 参 数   ：无
'输 出 参 数   ：无
'返 回 值      ：无
'调 用 函 数   ：无
'被 调 用 情 况：
'作 者         ：汪林
'创 建 时 间   :
'==========================================================================================
Private Sub PinChangeTotal()
    Dim nTotPinChangeNum     As Long

    nTotPinChangeNum = Pcb3dl.DlGetInt("TotPinChangeNum")
    
    nTotPinChangeNum = nTotPinChangeNum + 1
    nrc = Pcb3dl.DlSetLong("TotPinChangeNum", nTotPinChangeNum)

End Sub
Private Sub RetToMaster(ByVal S3eRetValue As Integer)
    S3ETrans.result = S3eRetValue
End Sub
Private Sub TimerAction_Timer()
    Dim sSubStData      As String
    Dim sCurrentDate    As String
    Dim bIsTimerAgain   As Boolean
    Dim PrjString       As String
    Dim PrjCHNString    As String
    Dim inputStr1       As String
    Dim inputStr2       As String
    Dim tmp_newPinBlock As String
    Dim tmp_oldPinBlock As String
    Dim LineSendNum     As Variant
    Dim bIsPinBlockOK   As Boolean
    Dim tmp_Retry       As Integer
    Dim sHostRejectCard As String
    Dim TransCode       As String
    Dim HostAccNo       As String
    Dim bTrack3Update   As Boolean
    Dim soldPin         As String
    Dim sPinDifference  As String
    Dim HostRejCode          As Variant
    
    TimerAction.Enabled = False
    bIsTimerAgain = True

    Select Case currentPage
        Case pagePinChangeInput1
            nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
            nrc = Pcb3dl.DlSetCharRaw("HtmlInput2", "")
            nrc = ShowScreenSync(Browser, "PinChange", "PinChangeInput1", sSubStData)
            
            If nrc = 0 Then
                If sSubStData = "@ok" Then
                    inputStr1 = Pcb3dl.DlGetCharRaw("HtmlInput1")
                    If Len(inputStr1) < 6 Then     '2005.12.23
                        currentPage = pagePinChangeInput1
                    Else
                        currentPage = pagePinChangeInput2
                    End If
                ElseIf sSubStData = "@Update" Then
                     nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
                     currentPage = pagePinChangeInput1
                Else
                    currentPage = pagePinChangePressStop
                End If
            ElseIf nrc = 91 Then
                currentPage = pagePinChangeTimeOut
            Else
                LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nrc)
                currentPage = pageScreenError
            End If

        Case pagePinChangeInput2
            nrc = ShowScreenSync(Browser, "PinChange", "PinChangeInput2", sSubStData)
            
            If nrc = 0 Then
                If sSubStData = "@ok" Then
                    inputStr1 = Pcb3dl.DlGetCharRaw("HtmlInput1")
                    inputStr2 = Pcb3dl.DlGetCharRaw("HtmlInput2")
                    tmp_Retry = Pcb3dl.DlGetInt("PinChangeRetry")
                    tmp_Retry = tmp_Retry - 1
                    If inputStr1 <> inputStr2 Then
                        If tmp_Retry > 0 Then
                            nrc = Pcb3dl.DlSetCharRaw("PinChangeRetry", tmp_Retry)
                            currentPage = pagePinChangeDiff
                        Else
                            currentPage = pagePinChangeError
                        End If
                    Else
                        soldPin = Pcb3dl.DlGetCharRaw("ShufflePinBlock")
                        sPinDifference = PIN_Difference(soldPin, inputStr1)
                        nrc = Pcb3dl.DlSetCharRaw("PinChangeBlock", sPinDifference)
                        currentPage = pagePinChangePlsWait
                    End If
                ElseIf sSubStData = "@Update" Then
                     nrc = Pcb3dl.DlSetCharRaw("HtmlInput2", "")
                     currentPage = pagePinChangeInput2
                Else
                    currentPage = pagePinChangePressStop
                End If
            ElseIf nrc = 91 Then
                currentPage = pagePinChangeTimeOut
            Else
                LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nrc)
                currentPage = pageScreenError
            End If

        Case pagePinChangeEppInput1
            nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
            nrc = Pcb3dl.DlSetCharRaw("HtmlInput2", "")
            
            nrc = ShowScreenSync(Browser, "PinChange", "PinChangeEppInput1", sSubStData)
            
            If (nrc = 0 And sSubStData = "@ok") And (S3EPin.DigitsEntered <> 6) Then
                currentPage = pagePinChangeNotSix1
            Else
                bIsPinBlockOK = True
                If nrc = 0 And sSubStData = "@ok" Then
                    bIsPinBlockOK = genPinBlockHW(tmp_newPinBlock)
                End If
                
                If bIsPinBlockOK = False Then
                    nrc = ShowScreenSync(Browser, "Common", "ComPressStop", sSubStData)
                    Call SendExceptionMessage(S3ELineOut, Pcb3dl, "29")
                    PrjString = "  **EPP failed in PinChange"
                    PrjCHNString = "    **修改密码时加密键盘故障"
                    PrintJournal S3EPrj, PrjString, PrjCHNString, g_sPrjLanguage
                    RetToMaster ReturnPressStop
                    Exit Sub
                End If
                
                If nrc = 0 Then
                    Select Case sSubStData
                        Case "@ok"
                            tmp_oldPinBlock = Pcb3dl.DlGetCharRaw("PinInputBlock")
                            nrc = Pcb3dl.DlSetCharRaw("PinChangeBlock", tmp_newPinBlock)
                            currentPage = pagePinChangeEppInput2
                        Case "@stop"
                            currentPage = pagePinChangePressStop
                        Case "@Change"
                            currentPage = pagePinChangeEppInput1
                        Case "@dev_failed"      'only for EPP enable
                            nrc = ShowScreenSync(Browser, "Common", "ComPressStop", sSubStData)
                            Call SendExceptionMessage(S3ELineOut, Pcb3dl, "29")
                            PrjString = "  **EPP failed in PinChange"
                            PrjCHNString = "    **修改密码时加密键盘故障"
                            PrintJournal S3EPrj, PrjString, PrjCHNString, g_sPrjLanguage
                            RetToMaster ReturnPressStop
                            Exit Sub
                        Case "@dev_timeout"     'only for EPP enable
                            PrjString = "  **EPP TimeOut in PinChange"
                            PrjCHNString = "  **修改密码时加密键盘操作超时"
                            PrintJournal S3EPrj, PrjString, PrjCHNString, g_sPrjLanguage
                            RetToMaster ReturnTimeout
                            Exit Sub
                        Case Else
                            nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
                    End Select
                Else
                    LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nrc)
                    currentPage = pageScreenError
                End If
            End If

        Case pagePinChangeEppInput2
            nrc = ShowScreenSync(Browser, "PinChange", "PinChangeEppInput2", sSubStData)
            
            If (nrc = 0 And sSubStData = "@ok") And (S3EPin.DigitsEntered <> 6) Then
                currentPage = pagePinChangeNotSix2
            Else
                bIsPinBlockOK = True
                If nrc = 0 And sSubStData = "@ok" Then
                    bIsPinBlockOK = genPinBlockHW(inputStr1)
                End If
                
                If bIsPinBlockOK = False Then
                    nrc = ShowScreenSync(Browser, "Common", "ComPressStop", sSubStData)
                    Call SendExceptionMessage(S3ELineOut, Pcb3dl, "29")
                    nrc = S3EPrj.DoPrint("EPP failed in PinChange")
                    RetToMaster ReturnPressStop
                    Exit Sub
                End If
                
                If (nrc = 0) Then
                    Select Case sSubStData
                        Case "@ok"
                            inputStr2 = Pcb3dl.DlGetCharRaw("PinChangeBlock")
                            If inputStr2 <> inputStr1 Then
                                currentPage = pagePinChangeDiff
                            Else
                                currentPage = pagePinChangePlsWait
                            End If
                        Case "@stop"
                            currentPage = pagePinChangePressStop
                        Case "@Change"
                            currentPage = pagePinChangeEppInput2
                        Case "@dev_failed"      'only for EPP enable
                            nrc = ShowScreenSync(Browser, "Common", "ComPressStop", sSubStData)
                            Call SendExceptionMessage(S3ELineOut, Pcb3dl, "29")
                            PrjString = "  **EPP failed in PinChange"
                            PrjCHNString = "    **修改密码时加密键盘故障"
                            PrintJournal S3EPrj, PrjString, PrjCHNString, g_sPrjLanguage
                            RetToMaster ReturnPressStop
                            Exit Sub
                        Case "@dev_timeout"     'only for EPP enable
                            PrjString = "  **EPP TimeOut in PinChange"
                            PrjCHNString = "  **修改密码时加密键盘操作超时"
                            PrintJournal S3EPrj, PrjString, PrjCHNString, g_sPrjLanguage
                            RetToMaster ReturnTimeout
                            Exit Sub
                    End Select
                Else
                    LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nrc)
                    currentPage = pageScreenError
                End If
            End If
        
        Case pagePinChangeNotSix1
            nrc = ShowScreenSync(Browser, "PinChange", "PinChangeNotSix", sSubStData)
            If G_bIsHardware = True Then
                currentPage = pagePinChangeEppInput1
            Else
                currentPage = pagePinChangeInput1
            End If
            
        Case pagePinChangeNotSix2
            nrc = Pcb3dl.DlSetCharRaw("HtmlInput2", "")
            nrc = ShowScreenSync(Browser, "PinChange", "PinChangeNotSix", sSubStData)
            If G_bIsHardware = True Then
                currentPage = pagePinChangeEppInput2
            Else
                currentPage = pagePinChangeInput2
            End If
            
        Case pagePinChangeTimeOut
            RetToMaster ReturnTimeout
            Exit Sub

        Case pagePinChangeDiff
            nrc = ShowScreenSync(Browser, "PinChange", "PinChangeDiff", sSubStData)
            If G_bIsHardware = True Then
                currentPage = pagePinChangeEppInput1
            Else
                currentPage = pagePinChangeInput1
            End If
            
        Case pagePinChangeError
            nrc = ShowScreenSync(Browser, "PinChange", "PinChangeError", sSubStData)
            Call SendExceptionMessage(S3ELineOut, Pcb3dl, "48")
            PrjString = "   **Input retries exhausted." + vbCrLf + _
                                   "   **TRANSACTION FAILED**"
            PrjCHNString = "   **修改密码错误次数超限." + vbCrLf + "   **交易失败**"
            PrintJournal S3EPrj, PrjString, PrjCHNString, g_sPrjLanguage
            RetToMaster ReturnInputOverLimit
            Exit Sub
        
        Case pagepinchangeOk
            nrc = ShowScreenSync(Browser, "PinChange", "PinChangeOk", sSubStData)
            Call ResetATMPrr(PinChange.Pcb3dl)
            Call DrawPinChangePrr(Pcb3dl, PrrOK)
            Call PinChangeTotal
            RetToMaster ReturnOk
            Exit Sub
               
        Case pagePinChangeCommErr
            Call SendExceptionMessage(S3ELineOut, Pcb3dl, "64")
            nrc = ShowScreenSync(Browser, "Common", "ComCommErr", sSubStData)
            PrjString = "   **NO RESPONSE FROM ATMP." + vbCrLf + _
                                   "   **TRANSACTION FAILED**"
            PrjCHNString = "   **主机无响应" + vbCrLf + "   **交易失败**"
            PrintJournal S3EPrj, PrjString, PrjCHNString, g_sPrjLanguage
            RetToMaster ReturnCommErr
            Exit Sub
     
        Case pagePinChangeReject
            sHostRejectCard = Pcb3dl.DlGetCharRaw("HostRejectCard")
            If sHostRejectCard = "R" Then
                nrc = Pcb3dl.DlSetCharRaw("HostRejectChinese", "原密码输入错误，改密交易失败")
                nrc = Pcb3dl.DlSetCharRaw("HostRejectEnglish", "Old PIN input error, PIN did not be changed")
            End If
            nrc = ShowScreenSync(Browser, "Common", "ComReject", sSubStData)
            
             nrc = S3ELineOut.GetData("constHostRejectCode", HostRejCode)
            
            '流水打印主机拒绝码 2005.12.9
            PrjString = "   **HOST REJECT [" + HostRejCode + "]" + vbCrLf + _
                                   "   " + Pcb3dl.DlGetCharRaw("HostRejectEnglish")
            PrjCHNString = "   **主机拒绝 [" + HostRejCode + "]" + vbCrLf + _
                               "   " + Pcb3dl.DlGetCharRaw("HostRejectChinese")
            PrintJournal S3EPrj, PrjString, PrjCHNString, g_sPrjLanguage
            
            If sHostRejectCard = "R" Then
                RetToMaster ReturnPinNotMatch
            Else
                Call DrawPinChangePrr(Pcb3dl, PrrReject)
                RetToMaster ReturnHostReject
            End If
            Exit Sub

        Case pagePinChangePressStop
            RetToMaster ReturnPressStop
            Exit Sub
        
        Case pagePinChangePlsWait
            nrc = ShowScreenSync(Browser, "Common", "ComPlsWait", sSubStData)
            
            nrc = S3ELineOut.DoSend("PIN", NORMAL_MESSAGE)
            S3ELineOut.GetData "GBLLineNum", LineSendNum
            Pcb3dl.DlSetCharRaw "GBLLineSendNum", Format(LineSendNum, "0000")
            PrjString = " " + vbCrLf + _
                        "   **" + "PIN " + Format(Now(), " HH:MM:SS") + " [" + _
                             Format(LineSendNum, "000000") + "]" + vbCrLf + _
                        "   **ATM CODE:" + Format(Pcb3dl.DlGetCharRaw("GBLAtmCode")) + vbCrLf
            
            PrjCHNString = " " + vbCrLf + _
                        "   **" + "改密 " + Format(Now(), " HH:MM:SS") + " 流水号：[" + _
                             Format(LineSendNum, "000000") + "]" + vbCrLf + _
                        "   **ATM号:" + Format(Pcb3dl.DlGetCharRaw("GBLAtmCode")) + vbCrLf
            PrintJournal S3EPrj, PrjString, PrjCHNString, g_sPrjLanguage
            If nrc <> 0 Then
                currentPage = pagePinChangeCommErr
            Else
                nrc = S3ELineOut.GetData("GBLLineNum", GLnLineNum)
                nrc = Pcb3dl.DlSetCharRaw("GBLLineSendNum", _
                                Format(GLnLineNum, "00000"))
                
                nrc = S3ELineOut.DoReceive
                
                If nrc = 0 Then
                    g_sHostRespCode = Pcb3dl.DlGetCharRaw("HostTransCode")
                    If g_sHostRespCode = "ASP" Then
                       
                        HostAccNo = Pcb3dl.DlGetCharRaw("HostAccNo")
                        
                        If Trim(HostAccNo) <> Trim(g_strAccNo20) Then
                            currentPage = pagePinChangeCommErr
                        Else
                            'bTrack3Update = CheckHostResponseTrack3Data()
                            
                            Pcb3dl.DlSetCharRaw "FitPrrAccNo", _
                            Left(HostAccNo, Len(HostAccNo) - 5) + "****" + Right(HostAccNo, 1)
                            
                            PrjString = "   **  HOST ACCEPT " + vbCrLf + _
                              "    ** Host AccNo: " + HostAccNo + vbCrLf + _
                              "    ** Host Date: " + Pcb3dl.DlGetCharRaw("HostCurrentDate") + vbCrLf + _
                              "    ** Host Seq :" + Pcb3dl.DlGetCharRaw("IcbcHostSeq") + vbCrLf + _
                              "    ** TRANSACTION OK**"
                                
                             PrjCHNString = "   ** 主机接受 " + vbCrLf + _
                                          "    ** 主机返回帐号：" + HostAccNo + vbCrLf + _
                                          "    ** 主机时间： " + Pcb3dl.DlGetCharRaw("HostCurrentDate") + vbCrLf + _
                                          "    ** 主机检索号：" + Pcb3dl.DlGetCharRaw("IcbcHostSeq") + vbCrLf + _
                                          "    ** 交易成功完成**"
                            PrintJournal S3EPrj, PrjString, PrjCHNString, g_sPrjLanguage
                            nrc = Pcb3dl.DlSetCharRaw("GBLIsPinOk", "Y")
                            nrc = Pcb3dl.DlSetCharRaw("PinInputBlock", Pcb3dl.DlGetCharRaw("PinChangeBlock"))
                            currentPage = pagepinchangeOk
                        
                            Call PinChangeTotal
                        End If
                    Else
                        nrc = S3ELineOut.GetData("constHostRejectCode", g_sRejectCode)
                        nrc = Pcb3dl.DlSetCharRaw("ATMPRejectCode", g_sRejectCode)
                        currentPage = pagePinChangeReject
                    End If
                ElseIf nrc = 97 Then
                    'Host return MAC error,
                    'Set the trickle to download CommKey again in S3EStarter.exe
                    LogError "DoReceive return 97,host return MAC error"
                    nrc = Pcb3dl.DlSetCharRaw("ResetTransKey", "R")
                    currentPage = pagePinChangeCommErr
                Else
                    LogError "Received unknown TransCode, " & TransCode
                    currentPage = pagePinChangeCommErr
                End If    'endif doreceive
            End If

        Case pageScreenError
            bIsTimerAgain = False
            
        Case pageQuit
            Unload PinChange
            Exit Sub
            
        Case Else
            LogError "TimerAction next action case error. The next action is:" + _
                CStr(currentPage)
    End Select
    
    If bIsTimerAgain = True Then
        TimerAction.Enabled = True
    End If
End Sub
Private Sub Str2Bin(ByVal InPar As String, ByRef OutPar() As Byte)
    Dim i As Integer
    
    For i = 1 To 16 Step 2
        OutPar((i + 1) / 2) = Val("&H" + Mid(InPar, i, 2))
    Next
End Sub
Private Sub Bin2Str(ByRef InPar() As Byte, ByRef OutPar As String)
    Dim i        As Integer
    Dim strNum   As String

    For i = 1 To 8
        strNum = Hex(InPar(i))
        If Len(strNum) < 2 Then
            strNum = "0" + strNum
        End If
        OutPar = OutPar + strNum
    Next i

End Sub
Function AKeyXorBKey(ByVal ABuffer As String, ByVal BBuffer As String, ByVal InSize As Integer) As String
    Dim i                  As Integer
    Dim AArray(1 To 8)     As Byte
    Dim BArray(1 To 8)     As Byte
    Dim ABArray(1 To 8)    As Byte
    Dim XorResult          As String

    Call Str2Bin(ABuffer, AArray)
    Call Str2Bin(BBuffer, BArray)

    For i = 1 To InSize \ 2
        ABArray(i) = AArray(i) Xor BArray(i)
    Next
    XorResult = ""
    Call Bin2Str(ABArray, XorResult)
    AKeyXorBKey = XorResult

End Function
Private Function genPinBlockHW(ByRef pPinEnd As String) As Boolean
    Dim AccountNo    As String
    Dim CardData     As String
    Dim k            As Integer
    
    AccountNo = Pcb3dl.DlGetCharRaw("FitAccNo")
    k = Len(AccountNo)
    If k > 12 Then
        CardData = Mid(AccountNo, k - 12, 12)
    Else
        CardData = AccountNo
    End If
    S3EPin.CustData = CardData
    S3EPin.KeyName = "PIN"
    
    nrc = S3EPin.DoPreparePin()
    If nrc <> 0 Then
        genPinBlockHW = False
    Else
        pPinEnd = S3EPin.EPin
        genPinBlockHW = True
    End If
    
End Function
Private Function PIN_Difference(OldPin As String, NewPin As String)
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
'    nRc = Pcb3dl.DlSetCharRaw("PinChangeBlock", result)
    PIN_Difference = result
End Function

Function CheckHostResponseTrack3Data() As Boolean

    'Check the Track3 update
    Dim sTrack3Update As String
    Dim vTrack3Update As Variant
    Dim ii As Integer
    Dim sByte As String
    
    nrc = Pcb3dl.DlSetCharRaw("IcbcTrackUpdate", "")
    CheckHostResponseTrack3Data = False
    
    nrc = S3ELineOut.GetData("HostTrack3", vTrack3Update)
    sTrack3Update = vTrack3Update
    If sTrack3Update <> "" And Len(sTrack3Update) > 37 Then
        sTrack3Update = Replace(sTrack3Update, "D", "=")
        nrc = Pcb3dl.DlSetCharRaw("IcbcTrackUpdate", sTrack3Update)
        CheckHostResponseTrack3Data = True
    End If
End Function



