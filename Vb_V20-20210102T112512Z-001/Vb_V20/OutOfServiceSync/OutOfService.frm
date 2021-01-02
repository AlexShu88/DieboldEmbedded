VERSION 5.00
Object = "{B2110643-3E81-11D3-8ACC-00C04FF20A5D}#1.2#0"; "TransProv.dll"
Object = "{6C4DD4AB-27D5-11D3-96C4-000000000000}#1.0#0"; "S3ELineOutTcp.ocx"
Object = "{9C37E835-6A58-11D1-80C0-0020AF7093F9}#1.2#0"; "Dl.ocx"
Object = "{EACE4ED6-6930-11D0-AC6C-0020AFBDD1D4}#1.0#0"; "SDOFep.ocx"
Object = "{EACE4ECF-6930-11D0-AC6C-0020AFBDD1D4}#1.0#0"; "SDOEdm.ocx"
Object = "{192DFCF0-F664-11D3-8BD4-00C04FF20A5D}#1.1#0"; "AdvBrowser.ocx"
Object = "{BD8177C0-832C-11CF-BF42-0020AF7093F9}#1.0#0"; "SDOIdc.ocx"
Object = "{E64F71A6-E705-4151-9895-5138B7D67F3A}#1.0#0"; "CHPrj.ocx"
Begin VB.Form OutOfService 
   Caption         =   "OutOfService"
   ClientHeight    =   2550
   ClientLeft      =   1725
   ClientTop       =   2385
   ClientWidth     =   4230
   Icon            =   "OutOfService.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   4230
   WindowState     =   1  'Minimized
   Begin CHPRJLib.CHPrj SDOPrj 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   1
   End
   Begin SDOIdcLibCtl.SDOIdc S3EIdc 
      Height          =   735
      Left            =   2280
      OleObjectBlob   =   "OutOfService.frx":08CA
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin TRANSPROVLibCtl.TransactionProvider S3ETrans 
      Height          =   735
      Left            =   1440
      OleObjectBlob   =   "OutOfService.frx":08FC
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin S3EADVBROWSERLibCtl.AdvBrowser Browser 
      Height          =   660
      Left            =   165
      OleObjectBlob   =   "OutOfService.frx":0940
      TabIndex        =   4
      Top             =   1680
      Width           =   1905
   End
   Begin SDOEdmLibCtl.SDOEdm S3EEdm 
      Height          =   690
      Left            =   1500
      OleObjectBlob   =   "OutOfService.frx":0966
      TabIndex        =   3
      Top             =   885
      Width           =   1215
   End
   Begin SDOFepLibCtl.SDOFep SDOFep 
      Height          =   690
      Left            =   165
      OleObjectBlob   =   "OutOfService.frx":0996
      TabIndex        =   2
      Top             =   885
      Width           =   1215
   End
   Begin VB.Timer TimerCheckState 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3600
      Top             =   1635
   End
   Begin VB.Timer TimerAction 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   2040
   End
   Begin VB.CommandButton start 
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
      Height          =   630
      Left            =   2865
      TabIndex        =   1
      Top             =   105
      Width           =   1260
   End
   Begin S3ELINEOUTLib.S3ELineOut S3ELineOut 
      Height          =   750
      Left            =   2820
      TabIndex        =   0
      Top             =   840
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   1323
      _StockProps     =   1
   End
   Begin DLLib.DL Pcb3dl 
      Left            =   2835
      Top             =   870
      _Version        =   65538
      _ExtentX        =   2196
      _ExtentY        =   1217
      _StockProps     =   0
   End
End
Attribute VB_Name = "OutOfService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==========================================================================================
'版权说明:迪堡公司中国区技术部
'版本号：Agilis power 1.6
'生成日期：2005.8
'作者：  汪林(初始版）
'模块功能：暂停服务模块
'主要函数及其功能
' 全局变量
'修改日志
'-----------------------------------------------------------------------
'<时间>：2005.11.14
'<修改者>：孙世方
'<当前版本>：中行1.0.16
'<详细记录>：增加关闭磁卡读写器声音的处理
'==========================================================================================
Const INIfile As String = "c:\atmwosa\ini\global.ini"

Private currentPage     As Integer

Const pageOutOfService1         As Integer = 1
Const pageOutOfService2         As Integer = 2
Const pageOutOfService3         As Integer = 3
Const pageCheckShowInterval     As Integer = 4
Const pageQuit                  As Integer = 98
Const pageError                 As Integer = 99

Const ReturnOk                  As Integer = 10  'OutOfVisit ok - go to Idle
Const ReturnGotoOperator        As Integer = 101 'Go to Operator

Dim g_bIdcInUse                 As Boolean
Dim g_lShowInterval             As Long
Dim nRc                         As Integer

Dim g_bLineoutPuOpen            As Boolean
Dim g_bSysShutDown              As Boolean

Private Sub Form_Load()
    Dim sValue          As String
    
    sValue = "The version number of " & App.EXEName & ".exe is " & App.Major & "." _
            & App.Minor & ".0." & App.Revision
    
    LogInfo (sValue)
    
    S3ETrans.Available = True
End Sub
Private Sub S3ETrans_QuitTransaction()
    currentPage = pageQuit
    TimerAction.Interval = 1000
    TimerAction.Enabled = True
End Sub

Private Sub S3ETrans_StartTransaction(ByVal Action As Long)
    Dim Rc              As Integer
    Dim HWStatus        As String
   
    Rc = SDOFep.DoServiceClose
    g_bIdcInUse = False
    
    g_bLineoutPuOpen = False
    
'如果是由于硬件故障导致暂停服务
'DL变量"GBLDoRecovery"将被置位，使S3EMonitor检测到后起动硬件复位
    HWStatus = Pcb3dl.DlGetCharRaw("GBLHWStatus")
    If HWStatus = "C" Then
        Pcb3dl.DlSetCharRaw "GBLDoRecovery", "1"
    End If
    
    nRc = CheckState()
    If S3EIdc.CardPosition = devidc_cardjammed Or S3EIdc.CardPosition = devidc_cardentering Then
        currentPage = pageOutOfService3
    Else
        currentPage = pageOutOfService1
    End If
    If Action = 107 Then
        g_bIdcInUse = True
        Pcb3dl.DlSetCharRaw "HtmlPrompt2", "磁卡读写器故障"
        Pcb3dl.DlSetCharRaw "HtmlWork53", "Out of service due to the trouble of CardReader."
    Else
        TimerCheckState.Enabled = True
    End If
        
    If Pcb3dl.DlGetCharRaw("GBLIsDoRecoverying") = "Y" Then
        Pcb3dl.DlSetCharRaw "HtmlPrompt2", "系统正在自检，请稍后..."
        nRc = ShowIdleScreen("OutOfService", "OutOfService2", g_lShowInterval)
    Else
        TimerAction.Interval = 100
        TimerAction.Enabled = True
    End If
End Sub
Private Sub start_Click()
    Dim PhoneNum As String
    PhoneNum = GetIniS(INIfile, "CustomerInfo", "Telephone", "")
    nRc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", PhoneNum)
    Pcb3dl.DlSetCharRaw "HtmlFkeyList", ""
    Pcb3dl.DlSetCharRaw "HtmlFkeyMap", "3855"
    currentPage = pageOutOfService1
    TimerAction.Interval = 100
    TimerCheckState.Enabled = True
    TimerAction.Enabled = True
    nRc = Pcb3dl.DlSetCharRaw("GBLAtmStatus", "C")
    nRc = Pcb3dl.DlSetCharRaw("GBLLineStatus", "C")
    nRc = Pcb3dl.DlSetCharRaw("GBLOperStatus", "2")
    
End Sub

Private Sub TimerAction_Timer()
    Dim sSubStData                  As String
    Dim nLen                        As Integer
    Dim PrjString                   As String
    Dim PrjCHNString                As String
    Dim g_sPrjLanguage              As String
    
    TimerAction.Enabled = False
    Select Case currentPage
        Case pageOutOfService1
            nRc = ShowIdleScreen("OutOfService", "OutOfService1", g_lShowInterval)
            currentPage = pageCheckShowInterval
            TimerAction.Interval = 100
        
        Case pageCheckShowInterval
            If g_lShowInterval - 30000 > 0 Then
                g_lShowInterval = g_lShowInterval - 30000
                TimerAction.Interval = 30000
            Else
                TimerAction.Interval = g_lShowInterval
                currentPage = pageOutOfService2
            End If
        
        Case pageOutOfService2
            sSubStData = Pcb3dl.DlGetCharRaw("HtmlPrompt2")
            nLen = Len(sSubStData)
            If nLen = 0 Then
                Pcb3dl.DlSetCharRaw "HtmlPrompt2", "本机暂停服务!"
            End If
            nRc = ShowIdleScreen("OutOfService", "OutOfService2", g_lShowInterval)
            currentPage = pageOutOfService1
            TimerAction.Interval = g_lShowInterval
            
         Case pageOutOfService3
            Call SendExceptionMessage(S3ELineOut, Pcb3dl, "C6")
            nRc = ShowIdleScreen("OutOfService", "OutOfService3", g_lShowInterval)
            
            If GetIniS(INIfile, "Bank_Environment", "PrjLanguage", "E") = "E" Then
                g_sPrjLanguage = "E"
            Else
                g_sPrjLanguage = "C"
            End If
            
            PrjString = vbCrLf + "   " + Format(Now(), "YY/MM/DD HH:MM:SS") + vbCrLf + _
                        "   ** IDC Error (Crime Possible) **" + vbCrLf + vbCrLf
            PrjCHNString = vbCrLf + "   " + Format(Now(), "YY/MM/DD HH:MM:SS") + vbCrLf + _
                        "   ** 读卡器故障（可能有人做案！） **" + vbCrLf
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            
            currentPage = pageOutOfService1
            TimerAction.Interval = g_lShowInterval
             
        Case pageQuit
            Unload OutOfService
            Exit Sub
        
        Case Else
            LogError "Timer Current pages case error. The page number is:" + _
                CStr(currentPage)
    End Select
    TimerAction.Enabled = True
End Sub

Private Function ShowIdleScreen(ByVal Section As String, ByVal ScreenName As String, ByRef nTimer As Long)
    Dim Rc   As Integer
    Dim sStr As String
    Dim Path As String

    sStr = GetIniS("Screens.ini", Section, ScreenName, "")
    ScreenInfo = GetScreenInfo(sStr)

    Path = GetIniS("Screens.ini", Section, "path", "")
    Rc = Browser.DoShowScreenSync(Trim(Path) + "\" + ScreenInfo.Name, 0)
    If (Rc = 0) Then
        nTimer = ScreenInfo.Interval
    Else
        LogError "ShowScreen '" + ScreenInfo.Name + "' Error, Rc = " & CStr(Rc)
        nTimer = 0
    End If
    ShowIdleScreen = Rc
End Function

Private Sub TimerCheckState_Timer()
    TimerCheckState.Enabled = False
    nRc = CheckState()
    If nRc = 0 Then
        nRc = SDOFep.DoServiceOpen
        TimerAction.Enabled = False
        S3ETrans.Result = ReturnOk
    ElseIf nRc = 1 Then
        TimerAction.Enabled = False
        S3ETrans.Result = ReturnGotoOperator
    Else
        TimerCheckState.Enabled = True
    End If
End Sub

Private Function CheckState() As Integer
    Dim LineStatus As String
    Dim AtmStatus As String
    Dim OperStatus As String
    Dim HWStatus As String
    Dim PeriodStatus As String
    Dim HostCmdStatus As String
    Dim sStarterStatus As String
    Dim nRc As Integer
    Dim sSysShutDown As String
    Dim StarterStatus As String
    Dim RecoveryStatus As String
    Dim sAudioOffAgain      As String
    
    '增加此处处理，以避免声音无法关掉
    '此变量值在EndVisit的RetToMaster中被设置
    sAudioOffAgain = Pcb3dl.DlGetCharRaw("GBLAudioOffAgain")
    If sAudioOffAgain = "Y" Then
        nRc = SDOFep.SetIndicator(ind_audio, audio_off)
        If nRc <> 0 Then
            LogError ("CheckState: Audio is Not OFF! RC=" + CStr(nRc))
        Else
            LogWarning ("CheckState: Audio is OFF!")
            nRc = Pcb3dl.DlSetCharRaw("GBLAudioOffAgain", "N")
        End If
    End If
     
        
    OperStatus = Pcb3dl.DlGetCharRaw("GBLOperStatus")
    AtmStatus = Pcb3dl.DlGetCharRaw("GBLAtmStatus")
    CheckState = -1

    If OperStatus = "" Or AtmStatus = "" Then
        LogError "DataLink: OperStatus or AtmStatus is EMPTY."
        Exit Function
    End If

'Add for ShutDown System
    sSysShutDown = Pcb3dl.DlGetCharRaw("GBLSysShutDown")
    If AtmStatus = "C" And sSysShutDown = "P" Then
        Pcb3dl.DlSetCharRaw "HtmlPrompt2", "重新启动机器，请稍侯..."
        If Not g_bSysShutDown Then
            g_bSysShutDown = True
            'Info to Monitor shutdowm system
            Pcb3dl.DlSetCharRaw "GBLSysShutDown", "S"
        End If
        Exit Function
    End If
    If AtmStatus = "C" And sSysShutDown = "S" Then
        Pcb3dl.DlSetCharRaw "HtmlPrompt2", "重新启动机器，请稍侯..."
        Exit Function
    End If
'end of Add

    If OperStatus <> "2" Then
        CheckState = 1
    ElseIf g_bIdcInUse = True Then
        Pcb3dl.DlSetCharRaw "HtmlPrompt2", "硬件故障, 需要重新启动机器"
    ElseIf AtmStatus = "O" Then
        If g_bLineoutPuOpen = True Then
            nRc = S3ELineOut.DoSend("0800", 1)
        End If
        CheckState = 0
    Else
        HWStatus = Pcb3dl.DlGetCharRaw("GBLHWStatus")
        PeriodStatus = Pcb3dl.DlGetCharRaw("GBLPeriodStatus")
        HostCmdStatus = Pcb3dl.DlGetCharRaw("GBLHostCmdStatus")
        LineStatus = Pcb3dl.DlGetCharRaw("GBLLineStatus")
      
        StarterStatus = Pcb3dl.DlGetCharRaw("ResetTransKey")
        RecoveryStatus = Pcb3dl.DlGetCharRaw("GBLIsDoRecoverying")
            
        If StarterStatus = "R" Or StarterStatus = "I" Then
            Pcb3dl.DlSetCharRaw "HtmlPrompt2", "系统正在初始化，请稍候"
        ElseIf LineStatus = "C" Then
            Pcb3dl.DlSetCharRaw "HtmlPrompt2", "本机线路中断"
            sStarterStatus = Pcb3dl.DlGetCharRaw("GBLStarterStatus")
            If g_bLineoutPuOpen = False And sStarterStatus = "C" Then
                LogInfo "Line status not OK, Set g_bLineoutPuOpen "
                g_bLineoutPuOpen = True
            End If
        
        ElseIf PeriodStatus = "C" Then
            Pcb3dl.DlSetCharRaw "HtmlPrompt2", "本机会计周期关闭"
            
        ElseIf HostCmdStatus = "C" Then
            Pcb3dl.DlSetCharRaw "HtmlPrompt2", "本机逻辑关机"
            
        ElseIf HWStatus = "C" Then
            Pcb3dl.DlSetCharRaw "HtmlPrompt2", "本机硬件故障"
            
        ElseIf RecoveryStatus = "Y" Then
            Pcb3dl.DlSetCharRaw "HtmlPrompt2", "系统正在自检，请稍后..."
        Else
            Pcb3dl.DlSetCharRaw "HtmlPrompt2", "本机系统故障"
            LogError "Impossible Error take OutOfService live:" + "GBLAtmStatus=" + AtmStatus + _
                     ";GBLPeriodStatus=" + PeriodStatus + ";GBLLineStatus=" + LineStatus + _
                     ";GBLHostCmdStatus=" + HostCmdStatus + ";GBLHWStatus=" + HWStatus
        End If
    End If
End Function
