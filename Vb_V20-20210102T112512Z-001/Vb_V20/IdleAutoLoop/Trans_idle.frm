VERSION 5.00
Object = "{B2110643-3E81-11D3-8ACC-00C04FF20A5D}#1.2#0"; "TransProv.dll"
Object = "{9C37E835-6A58-11D1-80C0-0020AF7093F9}#1.2#0"; "Dl.ocx"
Object = "{6C4DD4AB-27D5-11D3-96C4-000000000000}#1.0#0"; "s3elineouttcp.ocx"
Object = "{DA559591-71AC-11D3-8B0E-00C04FF20A5D}#1.0#0"; "DlWait.ocx"
Object = "{BD8177C0-832C-11CF-BF42-0020AF7093F9}#1.0#0"; "SDOIdc.ocx"
Object = "{192DFCF0-F664-11D3-8BD4-00C04FF20A5D}#1.1#0"; "AdvBrowser.ocx"
Object = "{EACE4ED6-6930-11D0-AC6C-0020AFBDD1D4}#1.0#0"; "SDOFep.ocx"
Object = "{E64F71A6-E705-4151-9895-5138B7D67F3A}#1.0#0"; "CHPrj.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Idle 
   Caption         =   "Trans_Idle"
   ClientHeight    =   2115
   ClientLeft      =   195
   ClientTop       =   450
   ClientWidth     =   5385
   Icon            =   "Trans_idle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   5385
   WindowState     =   1  'Minimized
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4440
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin S3EADVBROWSERLibCtl.AdvBrowser Browser 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "Trans_idle.frx":08CA
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin CHPRJLib.CHPrj SDOPrj 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   1085
      _StockProps     =   1
   End
   Begin DLWaitLibCtl.DLWait S3EDLWaitFreshFitTable 
      Height          =   375
      Left            =   2400
      OleObjectBlob   =   "Trans_idle.frx":08F0
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin TRANSPROVLibCtl.TransactionProvider S3ETrans 
      Height          =   615
      Left            =   1440
      OleObjectBlob   =   "Trans_idle.frx":0940
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin SDOFepLibCtl.SDOFep SDOFep 
      Height          =   690
      Left            =   2895
      OleObjectBlob   =   "Trans_idle.frx":0974
      TabIndex        =   3
      Top             =   900
      Width           =   1260
   End
   Begin SDOIdcLibCtl.SDOIdc SDOIdc 
      Height          =   690
      Left            =   135
      OleObjectBlob   =   "Trans_idle.frx":099E
      TabIndex        =   1
      Top             =   90
      Width           =   1215
   End
   Begin VB.Timer TimerIdle 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3480
      Top             =   240
   End
   Begin VB.Timer TimerAction 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2760
      Top             =   240
   End
   Begin S3ELINEOUTLib.S3ELineOut S3ELineOut 
      Height          =   750
      Left            =   1410
      TabIndex        =   0
      Top             =   855
      Width           =   1425
      _Version        =   65536
      _ExtentX        =   2514
      _ExtentY        =   1323
      _StockProps     =   1
   End
   Begin DLLib.DL Pcb3dl 
      Left            =   1440
      Top             =   885
      _Version        =   65538
      _ExtentX        =   2355
      _ExtentY        =   1191
      _StockProps     =   0
   End
End
Attribute VB_Name = "Idle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'版权说明:迪堡公司中国区技术部
'版本号：Agilis power 1.6
'生成日期：2005.8
'作者：汪林（初始版）
'模块功能： 播放欢迎信息，等待客户插卡
'主要函数及其功能
' 全局变量
'修改日志
'================================================================
'修改日志
'<时间>：2005.8.23
'<修改者>：孙世方
'<详细记录>：
'         增加注释；删除无用变量；调整格式
'<时间>：2005.9.19
'<修改者>：孙世方
'<详细记录>：
'        1 参考江苏中行曾发生的插卡死机现象的修改方法，SDOIdc_AtStartTestEnd事件中rc=99增加相关处理
'        2 修改函数TimerAction_Timer()中pageInsertedWrongCard，增加针对g_sTrackInfo为空时的处理
'================================================================
'修改日志
'<时间>：2005.10.20
'<修改者>：周亮
'<详细记录>：
'         中行总行版上海分行需求：插卡后首先检查卡表，在卡表中的卡属于本行本地卡；否则发送PAN交易来判定卡类型
'================================================================
'修改日志
'<时间>：2005.11
'<修改者>：vincent
'<详细记录>：
'不要帐号中的空格，并给收条赋值
'================================================================
'修改日志
'<时间>：2005.11.21-22
'版本号： 1.1.16 (2005.11.21)
'<修改者>：孙世方
'<详细记录>：
'1 打印不在卡表中的卡号
'2 与声讯公司调试嵌入式网络数字硬盘录象机 EVR(Embedded Net DVR)
' 增加函数 SendMessageToCommPortInsertCard,
'  在插卡进入密码输入前发送到EVR用于在监控屏幕显示卡号
'================================================================

Const IniPath                     As String = "c:\atmwosa\ini\"
Const sGlobalIni                  As String = "C:\ATMWosa\Ini\global.ini"
Const rcDEVIDC_IDC_CCINSERTED     As Integer = 95

Const ReturnOk                    As Integer = 10  'Idle ok - go to PinInput
Const ReturnGotoOperator          As Integer = 101 'Go to Operator
Const ReturnGotoEndVisit          As Integer = 102 'Eject a wrong card - go to EndVisit
Const ReturnGotoOutOfService      As Integer = 103 'Find an error in CheckState() - go to OutOfService
Const ReturnIdcError              As Integer = 105 'IDC error during transaction - go to Idle to let CheckState() check
Const ReturnIdcInitError          As Integer = 106 'IDC error at the beginning of transaction - go to Idle to let CheckState() check
Const ReturnIdcInUse              As Integer = 107 'It should go to OutOfService to display an error message and go to Operator to reboot system
Const ReturnOperator              As Integer = 110  'Idle received Operator card - go to Endvisit

Public Enum pageType
    pageNothing = 0
    pageIdleHead = 2
    pageCardInserted = 3
    pageCardWrongInserted = 4
    pageCardCaptured = 5
    pageInsertedWrongCard = 6
    pageError = 99
    pageQuit = 98
End Enum
Private currentPage As pageType

'For ICBC_HQ Begin Add by lijun
Private Type CardTypeRec
    TrackToMatch     As Integer
    offset           As Integer
    Length           As Integer
    MatchChars       As String
    PinLength        As Integer
    PinMaxAttempts   As Integer
    CardType         As String
    AccNum_Track     As Integer
    AccNum_Len       As Integer
    AccNum_Offset    As Integer
End Type
Private CardIdx() As CardTypeRec

Private AccnoMismatchB         As String
Private AccnoMismatchL         As String
Private g_sTrackInfo           As String
Private g_nPhoneNum            As String
Private g_sTrack2              As String
Private g_sTrack3              As String
Private g_bIsNeedToSendExp     As Boolean
Private nrc                    As Integer
Private g_nRetToMaster         As Integer
Private g_sPrjLanguage         As String

Private Sub Form_Load()
    Dim sValue As String
    
    sValue = "The version number of " & App.EXEName & ".exe is " & App.Major & "." _
            & App.Minor & ".0." & App.Revision
    
    LogInfo (sValue)

'Modified for Opteva, Add FluxInActive. Changed back to 6 in AgilisPower 1.5
    SDOIdc.TracksToRead = 6
'Modified end

    nrc = SDOFep.SetIndicator(ind_fascialight, fascialight_on)
    
    'Reset the PcB3HtmlBrowser variables
    nrc = Pcb3dl.DlSetCharRaw("HtmlFkeyList", "")
    nrc = Pcb3dl.DlSetCharRaw("HtmlFkeyMap", "3855")
    
    g_nPhoneNum = GetIniS(IniPath + "Global.ini", "CustomerInfo", "Telephone", "")
    
    If GetIniS(sGlobalIni, "Bank_Environment", "PrjLanguage", "E") = "E" Then
        g_sPrjLanguage = "E"
    Else
        g_sPrjLanguage = "C"
    End If

    S3ETrans.Available = True
    S3EDLWaitFreshFitTable.Enabled = True
    Call InitCardType  'initialize card type index
    g_bIsNeedToSendExp = True
End Sub

Private Function ShowIdleScreen(ByVal Section As String, ByVal ScreenName As String, ByRef nTimer As Integer)
    Dim sStr          As String
    Dim Path          As String

    sStr = GetIniS("Screens.ini", Section, ScreenName, "")
    ScreenInfo = GetScreenInfo(sStr)

    Path = GetIniS("Screens.ini", Section, "path", "")
    nrc = Browser.DoShowScreenSync(Trim(Path) + "\" + ScreenInfo.Name, 0)
    If (nrc = 0) Then
        nTimer = ScreenInfo.Interval
    Else
        LogError "ShowScreen '" + ScreenInfo.Name + "' Error, Rc = " & CStr(nrc)
        nTimer = 0
    End If
End Function

Private Sub S3EDLWaitFreshFitTable_VariableChanged()
    Call InitCardType  'initialize card type index
End Sub
Private Sub S3ETrans_QuitTransaction()
    nrc = SDOIdc.DoStopTest
    
    currentPage = pageQuit
    TimerAction.Interval = 1000
    TimerAction.Enabled = True
End Sub

Private Sub S3ETrans_StartTransaction(ByVal Action As Long)

    g_nRetToMaster = 0
    
    If Action = 1 Then
        currentPage = pageNothing
        TimerIdle.Enabled = False
        
        g_sTrackInfo = ""
        
        nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", g_nPhoneNum)

        Call CleanData
        
        nrc = CheckState()
        If nrc <> 0 Then
            Exit Sub
        End If
        
        ' Start testing the card insertion
        nrc = SDOIdc.DoStartTest
        If nrc = 0 Then
            g_bIsNeedToSendExp = True
            TimerIdle.Enabled = True
            currentPage = pageIdleHead
            TimerAction.Interval = 100
            TimerAction.Enabled = True
        ElseIf nrc = 90 Then 'IDC is InUse
            Call SendExceptionMessage(S3ELineOut, Pcb3dl, "24")
            S3ETrans.Result = ReturnIdcInUse
        Else
            If g_bIsNeedToSendExp Then
                Call SendExceptionMessage(S3ELineOut, Pcb3dl, "24")
                g_bIsNeedToSendExp = False
            End If
            TimerIdle.Enabled = True
            g_nRetToMaster = ReturnIdcInitError
        End If
    Else
        LogError "Idle has no action:" + CStr(Action)
    End If
End Sub

Private Sub SDOIdc_CardWrongInserted()
    TimerAction.Enabled = False
    currentPage = pageCardWrongInserted
    TimerAction.Interval = 100
    TimerAction.Enabled = True
End Sub

Private Sub SDOIdc_AtStartTestEnd(ByVal Rc As Integer)
    TimerIdle.Enabled = False
    TimerAction.Enabled = False
    If Rc = rcDEVIDC_IDC_CCINSERTED Then 'CC inserted and data available
        If SDOIdc.TracksRead >= 2 Then
            g_sTrack2 = SDOIdc.IsoTrack2
            g_sTrack3 = SDOIdc.IsoTrack3
            currentPage = pageCardInserted
            TimerAction.Interval = 100
            TimerAction.Enabled = True
        Else
            currentPage = pageInsertedWrongCard
            g_sTrackInfo = "UNKNOWN"
            TimerAction.Interval = 100
            TimerAction.Enabled = True
        End If
    ElseIf Rc = 99 Then
         '参考楼翔在江苏中行的修改2005。9。19
        If SDOIdc.CardPosition = devidc_cardpresent Then
            currentPage = pageInsertedWrongCard
            g_sTrackInfo = "UNKNOWN"
            TimerAction.Interval = 100
            TimerAction.Enabled = True
        Else
            Sleep (3000)
            S3ETrans.Result = ReturnIdcError
        End If
    ElseIf Rc = 96 Or Rc = 98 Then
        LogError "Idc_AtStartTestEnd RC=" & CStr(Rc)
        currentPage = pageInsertedWrongCard
        g_sTrackInfo = "UNKNOWN"
        TimerAction.Interval = 100
        TimerAction.Enabled = True
    ElseIf Rc = 92 Then
        LogError "Idc_AtStartTestEnd RC=" & CStr(Rc)
    Else
        LogError "Idc_AtStartTestEnd RC=" & CStr(Rc)
        currentPage = pageCardWrongInserted
        g_sTrackInfo = "UNKNOWN"
        TimerAction.Interval = 100
        TimerAction.Enabled = True
    End If
End Sub

Private Sub TimerAction_Timer()
    Dim sSubStData     As String
    Dim bIsTimerAgain  As Boolean
    Dim nTimeSeconds   As Integer
    Dim PrjString      As String
    Dim PrjCHNString   As String
    
    TimerAction.Enabled = False
    bIsTimerAgain = True
    
    Select Case currentPage
        Case pageCardWrongInserted
            nrc = ShowScreenSync(Browser, "Idle", "CardWrongInserted", sSubStData)
            currentPage = pageIdleHead
            TimerAction.Interval = 100
            
        Case pageIdleHead
            nrc = ShowIdleScreen("Idle", "IdleHead", nTimeSeconds)
            Exit Sub
            
        Case pageCardInserted
            nrc = ShowScreenSync(Browser, "Idle", "CardInserted", sSubStData)
            nrc = CheckCard()
            If nrc = 1 Then
                '判断是否使用嵌入式网络数字硬盘录象机 EVR(Embedded Net DVR)
                If Pcb3dl.DlGetCharRaw("GBLEVRUse") = "Y" Then
                    Call SendMessageToCommPortInsertCard
                End If
                S3ETrans.Result = ReturnOk 'will goning to Pininput model
                bIsTimerAgain = False
            ElseIf nrc = 2 Then
                S3ETrans.Result = ReturnOperator
                bIsTimerAgain = False
            Else
                currentPage = pageInsertedWrongCard
            End If
            
        Case pageInsertedWrongCard
            nrc = ShowScreenSync(Browser, "Idle", "InsertedWrongCard", sSubStData)

            Call SendExceptionMessage(S3ELineOut, Pcb3dl, "26") 'CARD NOT VALID OR UNREADABLE
            If g_sTrackInfo = "UNKNOWN" Then
                Pcb3dl.DlSetCharRaw "FitAccNo", g_sTrackInfo
            ElseIf Len(g_sTrackInfo) <> 0 Then
                Pcb3dl.DlSetCharRaw "FitAccNo", Mid(g_sTrackInfo, AccnoMismatchB, AccnoMismatchL)
            Else            ' 增加针对g_sTrackInfo为空时的处理 2005。9。19
                 Pcb3dl.DlSetCharRaw "FitAccNo", "  "
            End If
            PrjString = "<<<" + vbCrLf + Format(Now(), "MM/DD-HH:MM ") + "  **The Card didn't matched in FIT."
            
            PrjCHNString = "客户插卡" + vbCrLf + _
                                Format(Now(), "MM/DD-HH:MM ") + "  **该卡无法匹配."
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage

            S3ETrans.Result = ReturnGotoEndVisit
            bIsTimerAgain = False
            
        Case pageQuit
            Unload Idle
            Exit Sub
            
        Case Else
            LogError "TimerAction next action case error. The next action is:" + _
                CStr(currentPage)
    End Select
    
    If bIsTimerAgain = True Then
        TimerAction.Enabled = True
    End If

End Sub

Private Sub TimerIdle_Timer()
    TimerIdle.Enabled = False

    nrc = CheckState()
    If nrc <> 0 Then
        TimerAction.Enabled = False
        Exit Sub
    End If
    
    If g_nRetToMaster <> 0 Then
        S3ETrans.Result = g_nRetToMaster
    Else
        TimerIdle.Enabled = True
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

'==========================================================================================
'函数功能 ：更新卡表类型
'输入参数 ：无
'输出参数 ：无
'返回值   :无
'调用函数 ：无
'被调用情况：
'       Form被加载时
'       Wait型DL变量GBLFreshFitTable被更新，表示主机更新了卡表
'作者 ：
'创建时间 :
'==========================================================================================
Sub InitCardType()
    Dim i                 As Integer
    Dim j                  As Integer
    Dim comma             As Integer
    Dim StrTmp            As String
    Dim MaxWithdrawAmount As Long
    Dim MaxTransferAmount As Long
    
    j = GetIniN(IniPath + "fit.ini", "General", "CurrentRecord", 0)
    MaxWithdrawAmount = GetIniN(IniPath + "fit.ini", "General", "MaxWithdrawAmount", 0)
    MaxTransferAmount = GetIniN(IniPath + "fit.ini", "General", "MaxTransferAmount", 0)
    nrc = Pcb3dl.DlSetCharRaw("FitMaxWthAmount", MaxWithdrawAmount)
    nrc = Pcb3dl.DlSetCharRaw("IcbcMaxTfrAmount", MaxTransferAmount)
    
    If j < 1 Then
        LogError "Initcialize fit.ini Error!"
    Else
        ReDim CardIdx(j)
        For i = 0 To j
            CardIdx(i).TrackToMatch = GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(CStr(i)), "TrackToMatch", 0)
            CardIdx(i).offset = 1 + GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(CStr(i)), "Offset", 0)
            CardIdx(i).Length = GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(CStr(i)), "Length", 0)
            CardIdx(i).MatchChars = GetIniS(IniPath + "fit.ini", "CardIndex" + LTrim(CStr(i)), "MatchChars", "")
            
            CardIdx(i).PinLength = GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(CStr(i)), "PinLength", 0)
            CardIdx(i).PinMaxAttempts = GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(CStr(i)), "PinMaxAttempts", 0)
            CardIdx(i).CardType = GetIniS(IniPath + "fit.ini", "CardIndex" + LTrim(CStr(i)), "CardType", 0)
            
            CardIdx(i).AccNum_Len = GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(CStr(i)), "AccNum_Len", 0)
            CardIdx(i).AccNum_Offset = GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(CStr(i)), "AccNum_Offset", 0)
            CardIdx(i).AccNum_Track = GetIniN(IniPath + "fit.ini", "CardIndex" + LTrim(CStr(i)), "AccNum_Track", 0)
        Next
       
        StrTmp = GetIniS(IniPath + "fit.ini", "CardMismatch", "CardAccountNumber", "")
        comma = InStr(1, StrTmp, ",")
        AccnoMismatchB = 1 + Val(Left(StrTmp, comma - 1))
        AccnoMismatchL = Val(Right(StrTmp, Len(StrTmp) - comma))
    End If
End Sub
'==========================================================================================
'函数功能 ：检查状态
'输入参数 ：无
'输出参数 ：无
'返回值   :（整型）0　-　正常
'调用函数 ：无
'被调用情况：
'       Form被加载时
'       Wait型DL变量GBLFreshFitTable被更新，表示主机更新了卡表
'作者 ：汪林
'创建时间 :
'==========================================================================================
Private Function CheckState() As Integer
    Dim OperStatus          As String
    Dim AtmStatus           As String
    Dim sAudioOffAgain      As String
    
    '增加此处处理，以避免声音无法关掉
    '此变量值在EndVisit的RetToMaster中被设置
    sAudioOffAgain = Pcb3dl.DlGetCharRaw("GBLAudioOffAgain")
    If sAudioOffAgain = "Y" Then
        nrc = SDOFep.SetIndicator(ind_audio, audio_off)
        If nrc <> 0 Then
            LogError ("CheckState: Audio is Not OFF! RC=" + CStr(nrc))
        Else
            LogWarning ("CheckState: Audio is OFF!")
            nrc = Pcb3dl.DlSetCharRaw("GBLAudioOffAgain", "N")
        End If
    End If
    
    OperStatus = Pcb3dl.DlGetCharRaw("GBLOperStatus")
    AtmStatus = Pcb3dl.DlGetCharRaw("GBLAtmStatus")

    If Len(OperStatus) = 0 Or Len(AtmStatus) = 0 Then
        LogError "DataLink Error: GBLOperStatus=" + OperStatus + "; GBLAtmStatus=" + AtmStatus
        CheckState = -1
        Exit Function
    End If
    
    If AtmStatus = "O" And OperStatus = "2" Then
        CheckState = 0
    ElseIf OperStatus <> "2" Then
        LogWarning "GBLOperStatus=" + OperStatus + "; GBLAtmStatus=" + AtmStatus
        SDOIdc.DoStopTest
        S3ETrans.Result = ReturnGotoOperator
        CheckState = -1
    Else
        LogWarning "GBLOperStatus=" + OperStatus + "; GBLAtmStatus=" + AtmStatus
        SDOIdc.DoStopTest
        S3ETrans.Result = ReturnGotoOutOfService
        CheckState = -1
    End If
End Function

Private Sub CleanData()
    
    nrc = Pcb3dl.DlSetCharRaw("FitTrack3Message", "")
    nrc = Pcb3dl.DlSetCharRaw("FitTrack2Message", "")
    
    nrc = Pcb3dl.DlSetCharRaw("FitCardPinLength", "")
    nrc = Pcb3dl.DlSetCharRaw("FitPinMaxAttempt", "")
    
    nrc = Pcb3dl.DlSetCharRaw("FitAccNo", "")
   
End Sub

Function InitCardDL(ByVal sTrackHandle As String, ByVal row As Integer) As String
    Dim i              As Integer
    Dim blank          As String
    Dim sAccNo         As String
    Dim sCardType      As String
    Dim sFitCardType   As String
    Dim tempaccno      As Integer
    
    If Len(g_sTrack3) <> 0 Then
        Pcb3dl.DlSetCharRaw "FitTrack3Message", g_sTrack3
    Else
        For i = 1 To 104
          blank = blank + " "
       Next
        Pcb3dl.DlSetCharRaw "FitTrack3Message", blank
    End If
    If Len(g_sTrack2) <> 0 Then
       Pcb3dl.DlSetCharRaw "FitTrack2Message", g_sTrack2
    Else
       For i = 1 To 37
          blank = blank + " "
       Next
       Pcb3dl.DlSetCharRaw "FitTrack2Message", blank
    End If
'Added by Wanglin for BocomNew
    nrc = Pcb3dl.DlSetCharRaw("HostTrack2", "")
    nrc = Pcb3dl.DlSetCharRaw("HostTrack3", "")
'Add end

    Pcb3dl.DlSetCharRaw "FitCardPinLength", CardIdx(row).PinLength
    Pcb3dl.DlSetCharRaw "FitPinMaxAttempt", CardIdx(row).PinMaxAttempts
    sCardType = CardIdx(row).CardType
    
    
'    If Left(sCardType, 1) = "3" Then
'        sFitCardType = "9"
'    ElseIf Left(sCardType, 1) = "2" Then
'        sFitCardType = "8"
'    Else
'        sFitCardType = Mid(sCardType, 2, 1)
'    End If
'
'    sFitCardType = Right(sCardType, 2) + sFitCardType
'
'    Pcb3dl.DlSetCharRaw "FitCardType", sCardType
    '将工行的取卡类型的内容删除
    '中行卡类型就六种01，02，03，04，05，以及99（本行本地卡）
    Pcb3dl.DlSetCharRaw "FitCardType", sCardType
    
    If CardIdx(row).AccNum_Len = 0 Or CardIdx(row).AccNum_Len = 99 Then
        tempaccno = InStr(sTrackHandle, "=") - 1
        If tempaccno <= 0 Or tempaccno > 19 Then
            tempaccno = 19
        End If
    Else
        tempaccno = CardIdx(row).AccNum_Len
    End If
    
    If Len(sTrackHandle) < (tempaccno + CardIdx(row).AccNum_Offset + 1) Then
        InitCardDL = "UNKNOWN"
    Else
        sAccNo = Mid(sTrackHandle, CardIdx(row).AccNum_Offset + 1, tempaccno)
        
        tempaccno = InStr(sAccNo, "=")
        If tempaccno <> 0 Then
            sAccNo = Mid(sAccNo, 1, tempaccno - 1)
        End If
        
        Pcb3dl.DlSetCharRaw "FitAccNo", sAccNo
        If Len(sAccNo) < 5 Then
            InitCardDL = "UNKNOWN"
        Else
            On Error Resume Next
            Pcb3dl.DlSetCharRaw "FitPrrAccNo", Left(sAccNo, Len(sAccNo) - 5) + _
                "****" + Right(sAccNo, 1)
            InitCardDL = CardIdx(row).CardType
        End If
    End If
End Function

Private Function CheckCard() As Integer
    Dim i             As Integer
    Dim find          As Boolean
    Dim bmatch        As Boolean
    Dim StrTrack      As String
    Dim StrmatchChars As String
    Dim sCardType     As String
    Dim AccNo         As String
    Dim CardType      As Variant
    Dim PrjString     As String
    Dim PrjCHNString  As String
    Dim sRejectCode   As String
    
    StrTrack = ""
    find = False
        
    For i = 0 To UBound(CardIdx)
        If CardIdx(i).TrackToMatch = 3 Then
            StrTrack = g_sTrack3
           Else
            StrTrack = g_sTrack2
        End If
        
        If Len(StrTrack) <> 0 And Len(StrTrack) > 20 Then
            StrmatchChars = Mid(StrTrack, CardIdx(i).offset, CardIdx(i).Length)
            bmatch = MatchChars_Compare(StrmatchChars, CardIdx(i).MatchChars, CardIdx(i).Length)
            If bmatch Then
            'add 2001.10.31
                If CardIdx(i).AccNum_Track = 3 Then
                    StrTrack = g_sTrack3
                Else
                    StrTrack = g_sTrack2
                End If
            'end add
                sCardType = InitCardDL(StrTrack, i)
                If sCardType = "UNKNOWN" Then
                    CheckCard = 0
                    Exit Function
                ElseIf sCardType = "OP" Then
                    CheckCard = 2
                Else
                     
                     AccNo = Pcb3dl.DlGetCharRaw("FitAccNo")
                     PrjString = "<<<" + vbCrLf + _
                                Format(Now(), "MM/DD-HH:MM ") + " " + AccNo
                     PrjCHNString = "客户插卡" + vbCrLf + _
                                Format(Now(), "MM/DD-HH:MM ") + " " + AccNo
                     PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                                
                     CheckCard = 1
                End If
                find = True
                Exit For
            End If
        End If
    Next
'周亮修改 2005/10/20''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (Not find) Then
        Pcb3dl.DlSetCharRaw "FitTrack2Message", g_sTrack2
        Pcb3dl.DlSetCharRaw "FitTrack3Message", g_sTrack3

        nrc = S3ELineOut.DoSend("PAN", 0)
        If nrc <> 0 Then
            CheckCard = 0
            Call SendExceptionMessage(S3ELineOut, Pcb3dl, "64")
        Else
            nrc = S3ELineOut.DoReceive
            If nrc <> 0 Then
                CheckCard = 0
                Call SendExceptionMessage(S3ELineOut, Pcb3dl, "64")
            Else
                sRejectCode = Pcb3dl.DlGetCharRaw("HostTransCode")
                If sRejectCode = "APP" Then
                    S3ELineOut.GetData "HostCardType", CardType
                    AccNo = Trim(Pcb3dl.DlGetCharRaw("HostAccNo"))
                    Pcb3dl.DlSetCharRaw "FitCardType", CardType
                    
                    '不要帐号中的空格，并给收条赋值  2005.11 vincent
                    Pcb3dl.DlSetCharRaw "FitAccNo", AccNo
                    
                    Pcb3dl.DlSetCharRaw "FitPrrAccNo", Left(AccNo, Len(AccNo) - 5) + _
                    "****" + Right(Trim(AccNo), 1)
                    
                         PrjString = "<<<" + vbCrLf + _
                                    Format(Now(), "MM/DD-HH:MM ") + " " + AccNo
                         PrjCHNString = "客户插卡" + vbCrLf + _
                                    Format(Now(), "MM/DD-HH:MM ") + " " + AccNo
                         PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    
    
                    CheckCard = 1
                Else
                    CheckCard = 0
                End If
            End If
        End If
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function

Private Sub SendMessageToCommPortInsertCard()
   Dim StrTemp              As String
   Dim LngTemp              As Integer
   Dim i                    As Integer
   
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    StrTemp = Chr(2) + "ATM1"
    For i = 1 To 19
      StrTemp = StrTemp + Chr(0)
    Next
    
    StrTemp = StrTemp + "1000" + Pcb3dl.DlGetCharRaw("FitAccNo") + Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(0) + "0216"
    
    For i = 1 To 40
      StrTemp = StrTemp + Chr(0)
    Next
    
    LngTemp = 0
    For i = 1 To 96
         LngTemp = Asc(Mid(StrTemp, i, 1)) + LngTemp
    Next
    
    StrTemp = StrTemp + Right(Hex(LngTemp), 2) + Chr(0) + Chr(0)
    
    MSComm1.OutBufferCount = 0
    MSComm1.Output = StrTemp
    MSComm1.PortOpen = False

End Sub
