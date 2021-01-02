VERSION 5.00
Object = "{DA559591-71AC-11D3-8B0E-00C04FF20A5D}#1.0#0"; "DlWait.ocx"
Object = "{5C094E41-67D2-11D0-AC6B-0020AFBDD1D4}#1.0#0"; "SDOCdm.ocx"
Object = "{80C55DB1-86F3-11D3-8B2F-00C04FF20A5D}#1.0#0"; "S3ELineInTcp.ocx"
Object = "{EACE4EC1-6930-11D0-AC6C-0020AFBDD1D4}#1.0#0"; "SDOTtu.ocx"
Object = "{EACE4EDD-6930-11D0-AC6C-0020AFBDD1D4}#1.0#0"; "SDOOps.ocx"
Object = "{EACE4ED6-6930-11D0-AC6C-0020AFBDD1D4}#1.0#0"; "SDOFep.ocx"
Object = "{248BAFE0-D895-11CF-BFA3-0020AF7093F9}#1.0#0"; "SDODoor.ocx"
Object = "{3751B5D1-D348-11D0-AD02-0060970C3D2F}#1.0#0"; "SDOPrr.ocx"
Object = "{EACE4ECF-6930-11D0-AC6C-0020AFBDD1D4}#1.0#0"; "SDOEdm.ocx"
Object = "{BD8177C0-832C-11CF-BF42-0020AF7093F9}#1.0#0"; "SDOIdc.ocx"
Object = "{192DFCF0-F664-11D3-8BD4-00C04FF20A5D}#1.1#0"; "AdvBrowser.ocx"
Object = "{292DFCF0-F664-11D3-8BD4-00C04FF20A5D}#1.0#0"; "AdvBrowserMaint.ocx"
Object = "{9C37E835-6A58-11D1-80C0-0020AF7093F9}#1.2#0"; "Dl.ocx"
Object = "{6C4DD4AB-27D5-11D3-96C4-000000000000}#1.0#0"; "S3ELineOutTcp.ocx"
Object = "{7CCB2EF0-B3E8-11CF-BF8E-0020AF7093F9}#1.0#0"; "SDOPin.ocx"
Object = "{E64F71A6-E705-4151-9895-5138B7D67F3A}#1.0#0"; "CHPrj.ocx"
Begin VB.Form Monitor 
   Caption         =   "S3E Monitor"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   Icon            =   "Monitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin DLWaitLibCtl.DLWait S3eDLWaitResetTransKey 
      Height          =   375
      Left            =   4800
      OleObjectBlob   =   "Monitor.frx":0E42
      TabIndex        =   36
      Top             =   3120
      Width           =   1215
   End
   Begin DLWaitLibCtl.DLWait DLWait3 
      Height          =   30
      Left            =   3960
      OleObjectBlob   =   "Monitor.frx":0E8C
      TabIndex        =   35
      Top             =   3240
      Width           =   135
   End
   Begin DLWaitLibCtl.DLWait S3eDLWaitHost 
      Height          =   495
      Left            =   6120
      OleObjectBlob   =   "Monitor.frx":0EBA
      TabIndex        =   33
      Top             =   2760
      Width           =   2055
   End
   Begin DLLib.DL PCB3DL 
      Left            =   3120
      Top             =   3720
      _Version        =   65538
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   0
   End
   Begin S3EADVBROWSERLibCtl.AdvBrowser Browser 
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "Monitor.frx":0F06
      TabIndex        =   29
      Top             =   4440
      Width           =   1815
   End
   Begin DLWaitLibCtl.DLWait S3EDLWaitSysShutDown 
      Height          =   375
      Left            =   6000
      OleObjectBlob   =   "Monitor.frx":0F2C
      TabIndex        =   27
      Top             =   2280
      Width           =   2535
   End
   Begin DLWaitLibCtl.DLWait S3EDLWaitInitCasStates 
      Height          =   375
      Left            =   6000
      OleObjectBlob   =   "Monitor.frx":0F78
      TabIndex        =   28
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox TxtTransDate 
      DataSource      =   "DataTot"
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   4800
      TabIndex        =   26
      Text            =   "0101"
      Top             =   3960
      Width           =   1035
   End
   Begin VB.Data DataTot 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Width           =   1695
   End
   Begin ADVBROWSERMAINTATLLibCtl.AdvBrowserMaint BrowserMaint 
      Height          =   495
      Left            =   2040
      OleObjectBlob   =   "Monitor.frx":0FC8
      TabIndex        =   25
      Top             =   4440
      Width           =   2055
   End
   Begin DLWaitLibCtl.DLWait S3EDLWaitRecovery 
      Height          =   375
      Left            =   6000
      OleObjectBlob   =   "Monitor.frx":0FEE
      TabIndex        =   24
      Top             =   120
      Width           =   2535
   End
   Begin DLWaitLibCtl.DLWait S3EDLWaitAnomalies 
      Height          =   375
      Left            =   6000
      OleObjectBlob   =   "Monitor.frx":1038
      TabIndex        =   23
      Top             =   480
      Width           =   2535
   End
   Begin DLWaitLibCtl.DLWait S3EDLWaitHostCmd 
      Height          =   375
      Left            =   6000
      OleObjectBlob   =   "Monitor.frx":1086
      TabIndex        =   22
      Top             =   840
      Width           =   2520
   End
   Begin DLWaitLibCtl.DLWait S3EDLWaitPeriod 
      Height          =   375
      Left            =   6000
      OleObjectBlob   =   "Monitor.frx":10D6
      TabIndex        =   21
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Frame FrmModules 
      Height          =   2850
      Left            =   105
      TabIndex        =   10
      Top             =   45
      Width           =   5730
      Begin CHPRJLib.CHPrj SDOPrj 
         Height          =   735
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   1296
         _StockProps     =   1
      End
      Begin SDOPinLibCtl.SDOPin SDOPin 
         Height          =   615
         Left            =   4320
         TabIndex        =   30
         Top             =   2040
         Width           =   1215
         _cx             =   2143
         _cy             =   1085
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
      Begin SDOIdcLibCtl.SDOIdc SDOIdc 
         Height          =   690
         Left            =   165
         OleObjectBlob   =   "Monitor.frx":1124
         TabIndex        =   20
         Top             =   1965
         Width           =   1230
      End
      Begin SDOTtuLibCtl.SDOTtu SDOTtu 
         Height          =   690
         Left            =   4320
         OleObjectBlob   =   "Monitor.frx":1156
         TabIndex        =   19
         Top             =   1125
         Width           =   1230
      End
      Begin SDOOpsLibCtl.SDOOps SDOOps 
         Height          =   690
         Left            =   2940
         OleObjectBlob   =   "Monitor.frx":1180
         TabIndex        =   18
         Top             =   1125
         Width           =   1230
      End
      Begin SDOPrrLibCtl.SDOPrr SDOPrr 
         Height          =   690
         Left            =   1545
         OleObjectBlob   =   "Monitor.frx":11AA
         TabIndex        =   17
         Top             =   1125
         Width           =   1230
      End
      Begin SDODoorLibCtl.SDODoor SDODoor 
         Height          =   690
         Left            =   4320
         OleObjectBlob   =   "Monitor.frx":11DA
         TabIndex        =   16
         Top             =   285
         Width           =   1230
      End
      Begin SDOFepLibCtl.SDOFep SDOFep 
         Height          =   690
         Left            =   2940
         OleObjectBlob   =   "Monitor.frx":120A
         TabIndex        =   15
         Top             =   285
         Width           =   1230
      End
      Begin SDOEdmLibCtl.SDOEdm SDOEdm 
         Height          =   690
         Left            =   1545
         OleObjectBlob   =   "Monitor.frx":1234
         TabIndex        =   14
         Top             =   285
         Width           =   1230
      End
      Begin SDOCdmLibCtl.SDOCdm SDOCdm 
         Height          =   690
         Left            =   165
         OleObjectBlob   =   "Monitor.frx":1264
         TabIndex        =   13
         Top             =   285
         Width           =   1230
      End
      Begin S3ELINEOUTLib.S3ELineOut S3ELineOut 
         Height          =   690
         Left            =   1545
         TabIndex        =   11
         Top             =   1965
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   1217
         _StockProps     =   1
         BackColor       =   12582912
      End
      Begin S3ELINEINLibCtl.S3ELineIn S3ELineIn1 
         Height          =   690
         Left            =   2940
         OleObjectBlob   =   "Monitor.frx":129A
         TabIndex        =   12
         Top             =   1965
         Width           =   1230
      End
   End
   Begin VB.Frame Frame3 
      Height          =   540
      Left            =   105
      TabIndex        =   8
      Top             =   2955
      Width           =   1575
      Begin VB.Label Label1 
         Caption         =   "Last Error:"
         Height          =   270
         Left            =   75
         TabIndex        =   9
         Top             =   195
         Width           =   1395
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   8130
      Top             =   4395
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   8130
      Top             =   3795
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1725
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3045
      Width           =   2790
   End
   Begin VB.Frame Frame2 
      Caption         =   "Availability"
      Height          =   705
      Left            =   105
      TabIndex        =   4
      Top             =   3585
      Width           =   2745
      Begin VB.OptionButton Option5 
         Caption         =   "Out of Service"
         Height          =   255
         Left            =   1290
         TabIndex        =   6
         Top             =   285
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "In Service"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   285
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mode"
      Height          =   1335
      Left            =   6075
      TabIndex        =   0
      Top             =   3600
      Width           =   1575
      Begin VB.OptionButton Option3 
         Caption         =   "Supervisor"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1050
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Operator"
         Height          =   255
         Left            =   255
         TabIndex        =   2
         Top             =   585
         Width           =   930
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Client"
         Height          =   210
         Left            =   255
         TabIndex        =   1
         Top             =   285
         Width           =   750
      End
   End
   Begin DLWaitLibCtl.DLWait DLWait1 
      Height          =   375
      Left            =   0
      OleObjectBlob   =   "Monitor.frx":12C4
      TabIndex        =   32
      Top             =   0
      Width           =   2535
   End
   Begin DLWaitLibCtl.DLWait DLWait2 
      Height          =   495
      Left            =   0
      OleObjectBlob   =   "Monitor.frx":1310
      TabIndex        =   34
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Monitor"
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
'模块功能：打开、监控各模块状态，处理主机管理命令
'主要函数及其功能
' 全局变量
'        g_bIsFirstOpsKeyChanged    忽略第一次系统启动时SDOOps_AtKeyPosChanged事件
'修改日志
'-----------------------------------------------------------------------
'<时间>：2005.11
'<修改者>：孙世方
'<详细记录>：把CalculateAvailability函数中原来对s3elineout.available的判断改为对Datalink变量GBLLineStatus的判断
'----------------------------------------------------------------------
'<时间>：2005/12/07
'<修改者>：唐跃华
'       原来的钞箱状态显示有问题，现将钞箱的初始值显示设为未安装
'<时间>：2005.12.9
'<修改者>：孙世方
'<详细记录>：
'     1 设备故障 Prr ,IDC 时打印流水
'     2 上送打印机状态错，修改GetState函数，添加case 0 的处理。
'<时间>：2005.12.12
'<修改者>：孙世方
'<详细记录>：增加变量HostCutOffFlag，函数ClearCutOffData，用于处理主机CutOff时打印流水并清空统计值
'<时间>：2005.12.20
'<修改者>：孙世方
'<详细记录>:增加S3eDLWaitResetTransKey_VariableChanged,换密钥完全结束后在进入idle之前通过将Linestatus设置为N是不对的
'<时间>：2005.12.27
'<修改者>：孙世方
'版本号：1.3.6
'<详细记录>:
' 1 删除所有Devstate_change中有关S3EDLWaitAnomalies_VariableChanged的调用
' 2 修改CheckSPInfo函数（以前名字是CheckSPInfoNeedRecovery），并将其放到global中。
'3 修改MMDCode.ini 格式，增加取款ReversalState=0时，有可能出钞代码的判断
'==========================================================================================
Private Declare Function OpenEvent Lib "kernel32" Alias "OpenEventA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
                                                                      ByVal lpName As String) As Long
Private Declare Function SetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const keySelfService     As String = "Software\SelfService"
Private Const sGlobalIni         As String = "C:\ATMWosa\Ini\global.ini"

Private Const DB_TotPath         As String = "C:\S3e\Logs\LogTo\cwdlog.mdb"

Private Const EVENT_MODIFY_STATE = 2

Private Const DEVICE_CDM = &H1&
Private Const DEVICE_DEP = &H2&
Private Const DEVICE_DOOR = &H4&
Private Const DEVICE_EDM = &H8&
Private Const DEVICE_FEP = &H10&
Private Const DEVICE_ICC = &H20&
Private Const DEVICE_IDC = &H40&
Private Const DEVICE_OPS = &H80&
Private Const DEVICE_PIN = &H100&
Private Const DEVICE_PRJ = &H200&
Private Const DEVICE_PRR = &H400&
Private Const DEVICE_PRS = &H800&
Private Const DEVICE_SCM = &H1000&
Private Const DEVICE_TTU = &H2000&
Private Const DEVICE_CIM = &H8000&
Private Const DEVICE_LINE_IN = &H10000000
Private Const DEVICE_LINE_OUT = &H20000000

Enum HowExitConst
    EWX_LogOff = 0
    EWX_REBOOT = 2
    EWX_SHUTDOWN = 1
    EWX_FORCE = 4
    EWX_POWEROFF = 8
End Enum

Private Enum DevStatus
'    * 第一组32位状态
    STA01_HSTAT_OUT_OF_SERVICE = 1        '0 x80000000:       暂停客户服务
    STA02_HSTAT_HOST_IN_COMM = 2          '0 x40000000:       与主机通讯正常
    STA03_HSTAT_OPERATOR_SERVICE = 3      '0 x20000000:       进入维护
    STA04_HSTAT_ALARMS_ACTIVATED = 4      '0 x10000000:       保留位
    STA05_HSTAT_SUPPLY_KEYSWITCH = 5      '0 x08000000:       保留位
    STA06_HSTAT_SUPERVISORY_KEYSWITCH = 6 '0 x04000000:       保留位
    STA07_HSTAT_CARD_FAILURE = 7          '0 x02000000:       读卡器故障
    STA08_HSTAT_RECPRT_FAILURE = 8        '0 x01000000:       凭条打印机故障
    STA09_HSTAT_RECPRT_PAPER_OUT = 9      '0 x00800000:       凭条打印机缺纸
    STA10_HSTAT_JOUPRT_FAILURE = 10       '0 x00400000:       日志打印机故障
    STA11_HSTAT_JOUPRT_PAPER_OUT = 11     '0 x00200000:       日志打印机缺纸
    STA12_HSTAT_DEPOSITORY_FAILURE = 12   '0 x00100000:       保留位
    STA17_HSTAT_CDM_RETRACT_FAILURE = 17  '0 x00008000:       吐钞/存钞模块回收故障
    STA18_HSTAT_TERM_DOING_AUDIT = 18     '0 x00002000:       保留位
    STA20_HSTAT_BOOT_PERFORMED = 20       '0 x00001000:       保留位
    STA21_HSTAT_SAFEDOOR_O_C = 21         '0 x00000800：      保险门打开/关闭
    STA23_HSTAT_BILL_TRAP_DETECTED = 23   '0 x00000200:       钱钞陷落
    STA24_LSTAT_CAS5_LOW = 24             '0 x00000100:       钱箱5钞票过少
    STA25_LSTAT_CAS5_ERROR = 25           '0 x00000080:       钱箱5故障
    STA26_HSTAT_WRTTRK3_ERROR_EXCEED = 26 '0 x00000040:       保留位
    STA27_HSTAT_DIVERT_COUNTS_LOST = 27   '0 x00000020:       保留位
    STA30_HSTAT_WDDOOR_OPEN = 30          '0 x00000004：      保留位 //提款门打开
    '
    '    * 第二组32位状态
    STA33_LSTAT_CAS1_LOW = 33             '0 x80000000:       钱箱1钞票过少
    STA34_LSTAT_CAS2_LOW = 34             '0 x40000000:       钱箱2钞票过少
    STA35_LSTAT_CAS3_LOW = 35             '0 x20000000:       钱箱3钞票过少
    STA36_LSTAT_CAS4_LOW = 36             '0 x10000000:       钱箱4钞票过少
    STA37_LSTAT_CAS1_ERROR = 37           '0 x08000000:       钱箱1故障
    STA38_LSTAT_CAS2_ERROR = 38           '0 x04000000:       钱箱2故障
    STA39_LSTAT_CAS3_ERROR = 39           '0 x02000000:       钱箱3故障
    STA40_LSTAT_CAS4_ERROR = 40           '0 x01000000:       钱箱4故障
    STA41_LSTAT_DIVERT_FAULT = 41         '0 x00800000:       废钞箱故障
    STA42_LSTAT_CDM_HDW_ERROR = 42        '0 x00400000：      吐钞/存钞/循环模块硬件故障
    STA43_LSTAT_JOUPRT_PAPER_LOW = 43     '0 x00200000:       日志打印纸过少
    STA44_LSTAT_RECPRT_PAPER_LOW = 44     '0 x00100000:       凭条打印纸过少
    STA47_DEP_RETRACT_FULL = 47           '0 x00020000:       回收箱满
    STA48_LSTAT_DSP_DOOR_ALARM = 48       '0 x00010000:       保留位
    STA49_LSTAT_CAS1_HIGH = 49            '0 x00008000:       钱箱1钞票过多
    STA50_LSTAT_CAS2_HIGH = 50            '0 x00004000:       钱箱2钞票过多
    STA51_LSTAT_CAS3_HIGH = 51            '0 x00002000:       钱箱3钞票过多
    STA52_LSTAT_CAS4_HIGH = 52            '0 x00001000:       钱箱4钞票过多
    STA53_LSTAT_CAS5_HIGH = 53            '0 x00000800:       钱箱5钞票过多
    STA63_LSTAT_CUSTOMER_SERV_CLOSED = 63 '0 x00000002:       保留位
    STA64_LSTAT_TRANSACTION_STARTED = 64  '0 x00000001：      保留位
End Enum
Dim g_eDevStatusOffset              As DevStatus

Dim G_nDevicesCountToUse            As Byte

Dim G_nDevicesToUse                 As Long
Dim G_nOpenedDevicesToUse           As Long
Dim G_nNeededDevices                As Long
Dim G_nCKLInterval                  As Long
Dim G_nCKLCurrentCount              As Long
Dim G_nWaitForReboot                As Long

Dim nrc                             As Integer
Dim nTimerSequence                  As Integer
Dim G_nHoursOfStartPeriod           As Integer

Dim G_bATMAvailable                 As Boolean
'Dim G_bLineAvailable                As Boolean
Dim G_bPeriodAvailable              As Boolean
Dim G_bHostCmdAvailable             As Boolean
Dim G_bHWAvailable                  As Boolean
Dim g_bIsFirstOpsKeyChanged         As Boolean
Dim G_bAutoReboot                   As Boolean
Dim G_bGetAnomaliesBusy             As Boolean
Dim bIsMultiCurrency                As Boolean
Dim G_bIsStartPeriod                As Boolean
Dim G_bIsAudioTimeControl           As Boolean
Dim G_bIsAudioEnabled               As Boolean
Dim g_bLineStsChanged               As Boolean
Dim G_KeyAvailable                  As Boolean

Dim G_sDeviceStatus                 As String
Dim G_sNewDeviceStatus              As String
Dim g_sGBLOperStatus                As String
Dim g_sPrjLanguage                  As String

Dim G_AutoStartPeriod               As Date
Dim G_AutoRebootDateTime            As Date
Dim G_AudioStartTime                As Date
Dim G_AudioEndTime                  As Date

'==========================================================================================
'函数的功能 ：从注册表读取registry变量，register每个使用设备,调用linein模块
'输入参数   ：无
'输出参数   ：无
'返回值     ：无
'作者       ：
'创建时间   :
'==========================================================================================
Private Sub Form_Load()
    Dim nReply           As Integer
    Dim sValue           As String
    Dim nMessNumber      As Long
    Dim GBLHWStatus      As String
    Dim GBLAtmStatus     As String
    Dim GBLPeriodStatus  As String
    Dim GBLLineStatus    As String
    Dim GBLHostCmdStatus As String
    
    sValue = "The version number of " & App.EXEName & ".exe is " & App.Major & "." _
            & App.Minor & ".0." & App.Revision
    
    LogInfo (sValue)
    
    LogInfo ("S3EMonitor Form_Load")

    G_bGetAnomaliesBusy = False
    DataTot.DatabaseName = DB_TotPath
    nReply = PCB3DL.DlSetLong("GBLCdmRecoveryTimes", 3)
    
    nMessNumber = PCB3DL.DlGetInt("MessNumber")
    If nMessNumber < 0 Then
        PCB3DL.DlReset ("MessNumber")
    End If

    'G_sDeviceStatus = "10" + String(36, "0") + "22" + String(24, "0")
    '原来的钞箱状态显示有问题，现将钞箱的初始值显示设为未安装  2005/12/07 唐跃华
    G_sDeviceStatus = "10" + String(22, "0") + "2" + String(11, "0") + "2222" + String(24, "0")
    
    G_sNewDeviceStatus = G_sDeviceStatus
    
    If InitializeCounters() Then
        Timer1.Enabled = True
    End If

    'Set OperStatus to Client Mode
    nReply = PCB3DL.DlSetCharRaw("GBLOperStatus", "2")
    If nReply <> 0 Then
        LogError "    DlSetCharRaw 'GBLOperStatus' returned " & CStr(nReply)
        Text1.Text = "DlSetCharRaw 'GBLOperStatus' returned " & CStr(nReply)
    End If
    
    '硬件状态
    GBLHWStatus = PCB3DL.DlGetCharRaw("GBLHWStatus")
    If GBLHWStatus = "O" Then
        G_bHWAvailable = True
    Else
        GBLHWStatus = "C"
        G_bHWAvailable = False
    End If
    LogInfo "    GBLHWStatus is '" & GBLHWStatus & "'"
    nReply = PCB3DL.DlSetCharRaw("GBLHWStatus", GBLHWStatus)
    If nReply <> 0 Then
        LogError "    DlSetCharRaw 'GBLHWStatus' returned " & CStr(nReply)
        Text1.Text = "DlSetCharRaw 'GBLHWStatus' returned " & CStr(nReply)
    End If

    ' Try to figure out the status of the ATM (maybe the app was already started)
    GBLAtmStatus = PCB3DL.DlGetCharRaw("GBLAtmStatus")
    If GBLAtmStatus = "O" Then
        G_bATMAvailable = True
        Option4.Value = True        ' Check in-service
    Else
        GBLAtmStatus = "C"
        G_bATMAvailable = False
        Option5.Value = True        ' Check out-of-service
    End If
    
    LogInfo "    GBLAtmStatus is '" & GBLAtmStatus & "'"
    nReply = PCB3DL.DlSetCharRaw("GBLAtmStatus", GBLAtmStatus)
    If nReply <> 0 Then
        LogError "    DlSetCharRaw 'GBLAtmStatus' returned " & CStr(nReply)
        Text1.Text = "DlSetCharRaw 'GBLAtmStatus' returned " & CStr(nReply)
    End If
    
    'For CMB Shenzhen
    g_eDevStatusOffset = STA01_HSTAT_OUT_OF_SERVICE
    Call SetNewDeviceStatus(g_eDevStatusOffset, Not G_bATMAvailable)

    '周期开关状态
    GBLPeriodStatus = PCB3DL.DlGetCharRaw("GBLPeriodStatus")
    LogInfo "    GBLPeriodStatus is '" & GBLPeriodStatus & "'"
    If GBLPeriodStatus = "O" Then
        G_bPeriodAvailable = True
    Else
        GBLPeriodStatus = "C"
        G_bPeriodAvailable = False
    End If
    
    '线路状态
    If S3ELineOut.Available Then
        GBLLineStatus = "O"
'        G_bLineAvailable = True
    Else
        GBLLineStatus = "C"
'        G_bLineAvailable = False
    End If
    PCB3DL.DlSetCharRaw "GBLLineStatus", GBLLineStatus
    LogInfo "    GBLLineStatus is '" & GBLLineStatus & "'"

    '主机管理命令状态
    GBLHostCmdStatus = PCB3DL.DlGetCharRaw("GBLHostCmdStatus")
    LogInfo "    GBLHostCmdStatus is '" & GBLHostCmdStatus & "'"
    
    If GBLHostCmdStatus = "O" Or GBLHostCmdStatus = "P" Then
        G_bHostCmdAvailable = True
    Else
        PCB3DL.DlSetCharRaw "GBLHostCmdStatus", "C"
        G_bHostCmdAvailable = False
    End If
    
    '多币种选择
    sValue = GetIniS(sGlobalIni, "Withdrawal", "MultiCurrency", "N")
    If sValue = "Y" Then
        bIsMultiCurrency = True
    Else
        bIsMultiCurrency = False
    End If
    
    'enable所有waitable变量
    LogInfo "    Starting threads for DL change notifications"
    LogInfo "     Starting S3EDLWaitPeriod"
    S3EDLWaitPeriod.Enabled = True
    
    LogInfo "     Starting S3EDLWaitHostCmd"
    S3EDLWaitHostCmd.Enabled = True
    
    LogInfo "     Starting S3EDLWaitRecovery"
    S3EDLWaitRecovery.Enabled = True
    
    LogInfo "     Starting S3EDLWaitAnomalies"
    S3EDLWaitAnomalies.Enabled = True
    
    LogInfo "     Starting S3EDLWaitInitCasStates"
    S3EDLWaitInitCasStates.Enabled = True
    
    LogInfo "     Starting S3EDLWaitSysShutDown"
    S3EDLWaitSysShutDown.Enabled = True
    
    S3eDLWaitHost.Enabled = True
    S3eDLWaitResetTransKey.Enabled = True
    
    '开始LineIn进程
    nReply = S3ELineIn1.DoStartTest()
    LogInfo "    S3ELineIn.DoStartTest returned " & CStr(nReply)

    '从注册表得到使用设备和必须设备配置，并记录Log
    G_nDevicesToUse = GetRegKeyN(HKEY_LOCAL_MACHINE, keySelfService, "DevicesToUse", 4, 0)
    G_nNeededDevices = GetRegKeyN(HKEY_LOCAL_MACHINE, keySelfService, "NeededDevices", 4, 0)
    LogInfo "DevicesToUse = " & Str(G_nDevicesToUse)
    LogInfo "NeededDevices = " & Str(G_nNeededDevices)

    If G_nNeededDevices And DEVICE_CDM Then LogInfo "    CDM is needed"
    If G_nNeededDevices And DEVICE_DEP Then LogInfo "     DEP is needed"
    If G_nNeededDevices And DEVICE_DOOR Then LogInfo "     DOOR is needed"
    If G_nNeededDevices And DEVICE_EDM Then LogInfo "     EDM is needed"
    If G_nNeededDevices And DEVICE_FEP Then LogInfo "     FEP is needed"
    If G_nNeededDevices And DEVICE_ICC Then LogInfo "     ICC is needed"
    If G_nNeededDevices And DEVICE_IDC Then LogInfo "     IDC is needed"
    If G_nNeededDevices And DEVICE_OPS Then LogInfo "     OPS is needed"
    If G_nNeededDevices And DEVICE_PIN Then LogInfo "     PIN is needed"
    If G_nNeededDevices And DEVICE_PRJ Then LogInfo "     PRJ is needed"
    If G_nNeededDevices And DEVICE_PRR Then LogInfo "     PRR is needed"
    If G_nNeededDevices And DEVICE_SCM Then LogInfo "     SCM is needed"
    If G_nNeededDevices And DEVICE_TTU Then LogInfo "     TTU is needed"
    If G_nNeededDevices And DEVICE_LINE_IN Then LogInfo "     LINE_IN is needed"
    If G_nNeededDevices And DEVICE_LINE_OUT Then LogInfo "     LINE_OUT is needed"

    '注册EDM模块
    If G_nDevicesToUse And DEVICE_EDM Then
        LogInfo "    EDM: register"
        SDOEdm.Register 0
        If SDOEdm.Available Then
            LogInfo "    EDM is already available"
            SDOEdm.BackColor = &HFF00&
        Else
            SDOEdm.BackColor = &HFF&
        End If
    Else
        SDOEdm.BackColor = 0
    End If

    '注册PIN模块
    If G_nDevicesToUse And DEVICE_PIN Then
        LogInfo "    PIN: register"
        SDOPin.Register 0
        If SDOPin.Available Then
            LogInfo "    PIN is already available"
            SDOPin.BackColor = &HFF00&
        Else
            SDOPin.BackColor = &HFF&
        End If
    Else
        SDOPin.BackColor = 0
    End If

    '注册PRJ模块
    If G_nDevicesToUse And DEVICE_PRJ Then
        LogInfo "    PRJ: register"
        SDOPrj.Register 0
        If SDOPrj.Available Then
            LogInfo "    PRJ is already available"
            SDOPrj.BackColor = &HFF00&
        Else
            SDOPrj.BackColor = &HFF&
        End If
    Else
        SDOPrj.BackColor = 0
    End If
    
    '注册IDC模块
    If G_nDevicesToUse And DEVICE_IDC Then
        LogInfo "    IDC: register"
        SDOIdc.Register 0
        If SDOIdc.Available Then
            LogInfo "    IDC is already available"
            SDOIdc.BackColor = &HFF00&
        Else
            SDOIdc.BackColor = &HFF&
        End If
    Else
        SDOIdc.BackColor = 0
    End If
    
     '注册FEP模块
    If G_nDevicesToUse And DEVICE_FEP Then
        LogInfo "    FEP: register"
        SDOFep.Register 0
        If SDOFep.Available Then
            LogInfo "    FEP is already available"
            SDOFep.BackColor = &HFF00&
        Else
            SDOFep.BackColor = &HFF&
        End If
        'Add for Tri-color
        SDOFep.GuidLightColor = color_red
    Else
        SDOFep.BackColor = 0
    End If

    '注册DOOR模块
    If G_nDevicesToUse And DEVICE_DOOR Then
        LogInfo "    DOOR: register"
        SDODoor.Register 0
        If SDODoor.Available Then
            LogInfo "    DOOR is already available"
            SDODoor.BackColor = &HFF00&
        Else
            SDODoor.BackColor = &HFF&
        End If
    Else
        SDODoor.BackColor = 0
    End If
    
    '注册CDM模块
    If G_nDevicesToUse And DEVICE_CDM Then
        LogInfo "    CDM: register"
        SDOCdm.Register 0
        LogInfo "    CDM: register Completed"
        If SDOCdm.Available Then
            LogInfo "    CDM is already available"
            SDOCdm.BackColor = &HFF00&
        Else
            SDOCdm.BackColor = &HFF&
        End If
    Else
        SDOCdm.BackColor = 0
    End If

     '注册OPS模块
    If G_nDevicesToUse And DEVICE_OPS Then
        LogInfo "    OPS: register"
        SDOOps.Register 0
        LogInfo "    OPS: register Completed"
        If SDOOps.Available Then
            LogInfo "    OPS is already available"
            SDOOps.BackColor = &HFF00&
        Else
            SDOOps.BackColor = &HFF&
        End If
    Else
        SDOCdm.BackColor = 0
    End If

     '注册TTU模块
    If G_nDevicesToUse And DEVICE_TTU Then
        LogInfo "    TTU: register"
        SDOTtu.Register 0
        If SDOTtu.Available Then
            LogInfo "     TTU is already available"
            SDOTtu.BackColor = &HFF00&
        Else
            SDOTtu.BackColor = &HFF&
        End If
    Else
        SDOTtu.BackColor = 0
    End If
    
     '注册PRR模块
    If G_nDevicesToUse And DEVICE_PRR Then
        LogInfo "    PRR: register"
        SDOPrr.Register 0
        If SDOPrr.Available Then
            LogInfo "    PRR is already available"
            SDOPrr.BackColor = &HFF00&
        Else
            SDOPrr.BackColor = &HFF&
        End If
    Else
        SDOPrr.BackColor = 0
    End If
    
    '注册LINE_OUT
    If G_nDevicesToUse And DEVICE_LINE_OUT Then
        LogInfo "    LINE_OUT: register"
        S3ELineOut.Register 0
        If S3ELineOut.Available Then
            LogInfo "    LINE_OUT is already available"
            S3ELineOut.BackColor = &HFF00&
        Else
            S3ELineOut.BackColor = &HFF&
        End If
    Else
        S3ELineOut.BackColor = 0
    End If
      
    '计算当前状态
    CalculateAvailability

    G_nWaitForReboot = 0
    
    g_bIsFirstOpsKeyChanged = True
    
    G_nDevicesCountToUse = CaculateDeviceNumber(G_nDevicesToUse) - 2
    G_nOpenedDevicesToUse = 0
    
    g_bLineStsChanged = False
    
    nTimerSequence = 1
    
    Call CheckMaxBills    '当每次取款最大出钞数未设置时，设置为默认值30
    
    If Browser.HasSecondMonitor = 0 Then
        BrowserMaint.WindowStyle = WINDOWED
    End If
    
    If GetIniS(sGlobalIni, "Bank_Environment", "PrjLanguage", "E") = "E" Then
        g_sPrjLanguage = "E"
    Else
        g_sPrjLanguage = "C"
    End If

    '按照说明3.5.1.5 对取款模块进行有条件复位
    nrc = PCB3DL.DlReset("GBLCdmRecoveryNeeded")
    
    LogInfo "S3EMonitor Form_Load end"
    
    Timer2.Enabled = True
End Sub
'==========================================================================================
'函数的功能 ：按照当前硬件状态计算设备状态
'输入参数   ：无
'输出参数   ：无
'返回值     ：无
'作者       ：汪林
'创建时间   :2004.8
'==========================================================================================
Private Sub CalculateAvailability()
    Dim ATMAvailable    As Boolean
    Dim HWAvailable     As Boolean
    Dim nReply          As Integer
    Dim bLineAvailable  As Boolean
    
    LogInfo "CalculateAvailability"
    HWAvailable = True
    
    If (G_nNeededDevices And DEVICE_CDM) Then
        If (Not SDOCdm.Available) Then
            HWAvailable = False
        End If
    End If
    If (G_nNeededDevices And DEVICE_PIN) Then
        If (Not SDOPin.Available) Then
            HWAvailable = False
        End If
    End If
    If (G_nNeededDevices And DEVICE_FEP) Then
        If (Not SDOFep.Available) Then
            HWAvailable = False
        End If
    End If
    If (G_nNeededDevices And DEVICE_IDC) Then
        If (Not SDOIdc.Available) Then
            HWAvailable = False
        End If
    End If
    If (G_nNeededDevices And DEVICE_OPS) Then
        If (Not SDOOps.Available) Then
            HWAvailable = False
        End If
    End If
    If (G_nNeededDevices And DEVICE_PRJ) Then
        If (Not SDOPrj.Available) Then
            HWAvailable = False
        End If
    End If
    If (G_nNeededDevices And DEVICE_TTU) Then
        If (Not SDOTtu.Available) Then
            HWAvailable = False
        End If
    End If
    If (G_nNeededDevices And DEVICE_LINE_OUT) Then
        If (Not S3ELineOut.Available) Then
            HWAvailable = False
        End If
    End If
    LogInfo "    HWAvailable is " & CStr(HWAvailable)
    
    'Added for LoadKey status OK ===> G_KeyAvailable 2005.12.20
    LogInfo "   KeyStatus is " & CStr(G_KeyAvailable)
    
    If HWAvailable <> G_bHWAvailable Then
        G_bHWAvailable = HWAvailable
        If G_bHWAvailable Then
            nReply = PCB3DL.DlSetCharRaw("GBLHWStatus", "O")
        Else
            nReply = PCB3DL.DlSetCharRaw("GBLHWStatus", "C")
        End If
        If nReply <> 0 Then
            Text1.Text = "DlSetCharRaw 'GBLHWStatus' returned " & CStr(nReply)
        End If
    End If
    
    LogInfo "   PeriodAvailable is " & CStr(G_bPeriodAvailable)
    LogInfo "   HostCmdAvailable is " & CStr(G_bHostCmdAvailable)
    
    If PCB3DL.DlGetCharRaw("GBLLineStatus") = "O" Then
        bLineAvailable = True
    Else
        bLineAvailable = False
    End If
    
    ATMAvailable = HWAvailable And G_bPeriodAvailable And bLineAvailable _
            And G_bHostCmdAvailable And G_KeyAvailable
    LogInfo "    ATMAvailable is " & CStr(ATMAvailable)
    
    ' Only set the ATMStatus in DataLink when it has changed
    If ATMAvailable <> G_bATMAvailable Then
        G_bATMAvailable = ATMAvailable
        If G_bATMAvailable Then
            nReply = PCB3DL.DlSetCharRaw("GBLAtmStatus", "O")
            Option4.Value = True      ' Check In Service
        Else
            nReply = PCB3DL.DlSetCharRaw("GBLAtmStatus", "C")
            Option5.Value = True        ' Check Out of Service
        End If
        If nReply <> 0 Then
            Text1.Text = "DlSetCharRaw 'GBLAtmStatus' returned " & CStr(nReply)
        End If
    End If
    
    g_eDevStatusOffset = STA01_HSTAT_OUT_OF_SERVICE
    Call SetNewDeviceStatus(g_eDevStatusOffset, Not G_bATMAvailable)
    
    If G_sDeviceStatus <> G_sNewDeviceStatus Then
        Call UpdateStatusMessage
        G_sDeviceStatus = G_sNewDeviceStatus
        g_sGBLOperStatus = PCB3DL.DlGetCharRaw("GBLOperStatus")
        If g_sGBLOperStatus <> "1" Then
            BrowserMaint.DoRefresh
        End If
    End If
    
End Sub

'==========================================================================================
'函数的功能 ：处理所有例外信息
'输入参数   ：无
'输出参数   ：无
'返回值     ：无
'作者       ：汪林
'创建时间   :2004.8
'修改日志：
'==========================================================================================
Private Sub S3EDLWaitAnomalies_VariableChanged()
    Dim bAnomaliesLeft   As Boolean
    Dim stTime           As Date
    Dim nDevId           As Integer
    Dim nTOId            As Integer
    Dim nDOId            As Integer
    Dim nWosaReply       As Long
    Dim sSKBSReply       As String
    Dim sDescr           As String
    Dim sLogicalName     As String
    Dim sOldDescr        As String
    Dim fso              As New FileSystemObject
    Dim AnomalyStream    As TextStream
    Dim TextToPrint      As String
    Dim DlVarName        As String
    Dim sDevice          As String

    If (Not G_bGetAnomaliesBusy) Then
        G_bGetAnomaliesBusy = True
    
        LogInfo "Start GetAnomalies"
        If (Not fso.FileExists("c:\S3E\Logs\LogTO\anomaly.txt")) Then
            LogInfo "Creating new anomaly.txt file"
            Set AnomalyStream = fso.CreateTextFile("c:\S3E\Logs\LogTO\anomaly.txt")
        Else
            LogInfo "Opening existing anomaly.txt file"
            Set AnomalyStream = fso.GetFile("c:\S3E\Logs\LogTO\anomaly.txt").OpenAsTextStream(ForAppending)
        End If
        
        LogInfo "Retrieving anomalies"
        bAnomaliesLeft = SDOOps.GetAnomalyRaw(stTime, nDevId, nTOId, nDOId, nWosaReply, sSKBSReply, sDescr, sLogicalName)
        While bAnomaliesLeft
            Select Case nDevId
                Case 0:  DlVarName = "SiabFEPCode"
                         sDevice = "FEP"
                Case 2:  DlVarName = "SiabBGRCode"
                         sDevice = "IDC"
                Case 3:  DlVarName = "SiabPRJCode"
                         sDevice = "PRJ"
                Case 5:  DlVarName = "SiabDEPCode"
                         sDevice = "DEP"
                Case 8:  DlVarName = "SiabOPDCode"
                         sDevice = "TTU"
                Case 9:  DlVarName = "SiabOPKCode"
                         sDevice = "OPS"
                Case 11: DlVarName = "SiabALMCode"
                         sDevice = "DOOR"
                Case 14: DlVarName = "SiabPRSCode"
                         sDevice = "PRS"
                Case 19: DlVarName = "SiabDAMCode"
                         sDevice = "DAM"
                Case 23: DlVarName = "SiabSCMCode"
                         sDevice = "SCM"
                Case 28: DlVarName = "SiabICCCode"
                         sDevice = "ICC"
                Case 81: DlVarName = "SiabPRRCode"
                         sDevice = "PRR"
                Case 82: DlVarName = "SiabCIMCode"
                         sDevice = "CIM"
                Case 13: DlVarName = "SiabCDMCode"
                         sDevice = "CDM"
                        'Added by lijun for CDM do recovery by criteria in 2005-07-14
                         If (Not CheckSPInfo(sDescr, "NotRecoveryCode")) Then
                            'Restore the information to journal printer & LOG
                             TextToPrint = "*** " & Date$ & " " & Time$ & " ***" & vbCrLf & _
                                           "*** A severe CDM hardware fault happened!!! ***"
                             SDOPrj.DoPrint TextToPrint
                             LogError TextToPrint
                             nrc = PCB3DL.DlSetCharRaw("GBLCdmRecoveryNeeded", "N")
                         End If
                        'End of Add
            End Select
            
            TextToPrint = Date$ & " " & Format(Time$, "HH:MM:SS") & _
                        " (" & sLogicalName & Space(12 - Len(sLogicalName)) & _
                        ") DEV " & CStr(nDevId) & Space(3 - Len(CStr(nDevId))) & _
                        " SDO " & CStr(nTOId) & Space(4 - Len(CStr(nTOId))) & _
                        " TEC " & CStr(nDOId) & Space(4 - Len(CStr(nDOId))) & _
                        " XFS " & CStr(nWosaReply) & Space(5 - Len(CStr(nWosaReply))) & _
                        " SKBS " & sSKBSReply & Space(5 - Len(sSKBSReply)) & sDescr
            
            AnomalyStream.WriteLine TextToPrint
            
            If nWosaReply = 0 And sOldDescr <> sDescr Then
                TextToPrint = "ANOM " & Date$ & " " & Time$ & Space(24 - Len(sDevice)) & sDevice & Chr(13) & Chr(10) & _
                              "    SP Info: " & sDescr
                sOldDescr = sDescr
                SDOPrj.DoPrint TextToPrint
                LogInfo TextToPrint
            End If
            
            bAnomaliesLeft = SDOOps.GetAnomalyRaw(stTime, nDevId, nTOId, nDOId, nWosaReply, sSKBSReply, sDescr, sLogicalName)
        Wend
        
        LogInfo "No more anomalies"
        AnomalyStream.Close
        LogInfo "End GetAnomalies"
        
        G_bGetAnomaliesBusy = False
    Else
        LogInfo "GetAnomalies busy"
    End If

End Sub

Private Sub S3EDLWaitAnomalies_VariableInvalid()
    '如果进入到这里，程序会停滞
    LogError "GBLGetAnomalies is not waitable"
    Text1.Text = "S3EDLWaitAnomalies_VariableInvalid"
End Sub

Private Sub S3eDLWaitHost_VariableChanged()
  Dim HostcutOffFlag As String
  
  HostcutOffFlag = PCB3DL.DlGetCharRaw("HostCutOffFlag")
  
  If HostcutOffFlag = "Y" Then
      Call ClearCutOffData
  
  End If

End Sub
'==========================================================================================
'函数的功能 :用于处理主机CutOff时打印流水并清空统计值
'输入参数 :无
'输出参数 : 无
'返回值   :无
'调用函数 :无
'被调用情况  ：S3eDLWaitHost_VariableChanged
'作者       ：孙世方
'创建时间   :2005.12。12
'==========================================================================================
Sub ClearCutOffData()
    Dim CutOffIni                   As String
    Dim CutOffWithNum               As String
    Dim CutOffWithAmount            As String
    Dim CutOffTfrNum                As String
    Dim CutOffTfrAmount             As String
    Dim PrjString                   As String
    Dim PrjCHNString                As String
    Dim DateTime                 As String
    
    CutOffIni = "c:\ATMWosa\Ini\CutOff.ini"
    CutOffWithNum = GetIniS(CutOffIni, "HostCutOff", "WithdrawNumber", "0")
    CutOffWithAmount = GetIniS(CutOffIni, "HostCutOff", "WithdrawAmount", "0")
    CutOffTfrNum = GetIniS(CutOffIni, "HostCutOff", "TfrNumber", "0")
    CutOffTfrAmount = GetIniS(CutOffIni, "HostCutOff", "TfrAmount", "0")
    DateTime = GetIniS(CutOffIni, "HostCutOff", "DateTime", "0")
    
    nrc = SetIniS(CutOffIni, "HostCutOff", "DateTime", Format(Now, "YYYYMMDDHHMM"))
    If DateTime <> "0" And Format(Now - 0.25, "YYYYMMDDHHMM") > DateTime Then
        Exit Sub
    ElseIf CutOffWithNum = "0" And CutOffTfrNum = "0" Then
        Exit Sub
    Else
        PrjString = "  ** HOST CUT OFF" + vbCrLf + _
                "  Last Working Day Totals" + vbCrLf + _
                "  Type         Count    Amount " + vbCrLf + _
                "  Withdrawals  " + CutOffWithNum + "  " + CutOffWithAmount + vbCrLf + _
                "  Transfer     " + CutOffTfrNum + "  " + CutOffTfrAmount
        PrjCHNString = "  ** 主机正在日结" + vbCrLf + _
                    "  上一周期统计值" + vbCrLf + _
                    "  类型         数量         金额 " + vbCrLf + _
                    "  取款         " + CutOffWithNum + "          " + CutOffWithAmount + vbCrLf + _
                    "  转账         " + CutOffTfrNum + "          " + CutOffTfrAmount
        PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    
        nrc = SetIniS(CutOffIni, "HostCutOff", "WithdrawAmount", "0")
        nrc = SetIniS(CutOffIni, "HostCutOff", "WithdrawNumber", "0")
        nrc = SetIniS(CutOffIni, "HostCutOff", "TfrAmount", "0")
        nrc = SetIniS(CutOffIni, "HostCutOff", "TfrNumber", "0")
        
        nrc = SetIniS(CutOffIni, "Backup", "WithdrawAmount", CutOffWithAmount)
        nrc = SetIniS(CutOffIni, "Backup", "WithdrawNumber", CutOffWithNum)
        nrc = SetIniS(CutOffIni, "Backup", "TfrAmount", CutOffTfrAmount)
        nrc = SetIniS(CutOffIni, "Backup", "TfrNumber", CutOffTfrNum)
    End If
End Sub


'==========================================================================================
'函数的功能 ：收到主机管理
'输入参数   ：无
'输出参数   ：无
'返回值     ：无
'作者       ：汪林
'创建时间   :2004.8
'修改日志：
'==========================================================================================
Private Sub S3EDLWaitHostCmd_VariableChanged()
    Dim GBLHostCmdStatus  As String
    
    GBLHostCmdStatus = PCB3DL.DlGetCharRaw("GBLHostCmdStatus")
    LogInfo "GBLHostCmdStatus is now '" & GBLHostCmdStatus & "'"
    If GBLHostCmdStatus = "O" Then
        G_bHostCmdAvailable = True
    Else
        G_bHostCmdAvailable = False
    End If
    
    CalculateAvailability
End Sub
Private Sub S3EDLWaitHostCmd_VariableInvalid()
    LogError "GBLHostCmdStatus is not waitable"
    Text1.Text = "S3EDLWaitHostCmd_VariableInvalid"
End Sub
'==========================================================================================
'函数的功能 ：钞箱状态改变后，刷新所有钞箱状态
'输入参数   ：无
'输出参数   ：无
'返回值     ：无
'作者       ：汪林
'创建时间   :2004.8
'修改日志：
'==========================================================================================
Private Sub S3EDLWaitInitCasStates_VariableChanged()
    Dim liv_Loop  As Integer
    Dim nPosition As Integer
    
    SDOCdm.DataCriteria = 1
    For liv_Loop = 1 To SDOCdm.NbrOfBoxesUsed
        SDOCdm.CasNbrLogical = liv_Loop
        
        If Len(SDOCdm.CasPosition) > 0 Then
            If IsNumeric(Right(SDOCdm.CasPosition, 1)) Then
                nPosition = CInt(Right(SDOCdm.CasPosition, 1))
                If nPosition < 5 Then
                    Select Case SDOCdm.CasState
                        Case 0
                            g_eDevStatusOffset = STA33_LSTAT_CAS1_LOW + nPosition - 1
                            Call SetNewDeviceStatus(g_eDevStatusOffset, False)
                            g_eDevStatusOffset = STA37_LSTAT_CAS1_ERROR + nPosition - 1
                            Call SetNewDeviceStatus(g_eDevStatusOffset, False)
                        Case 3
                            g_eDevStatusOffset = STA33_LSTAT_CAS1_LOW + nPosition - 1
                            Call SetNewDeviceStatus(g_eDevStatusOffset, True)
                            g_eDevStatusOffset = STA37_LSTAT_CAS1_ERROR + nPosition - 1
                            Call SetNewDeviceStatus(g_eDevStatusOffset, False)
                        Case Else
                            g_eDevStatusOffset = STA37_LSTAT_CAS1_ERROR + nPosition - 1
                            Call SetNewDeviceStatus(g_eDevStatusOffset, True)
                     End Select
                Else
                    Select Case SDOCdm.CasState
                    Case casstate_cdm_ok
                        g_eDevStatusOffset = STA24_LSTAT_CAS5_LOW
                        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
                        g_eDevStatusOffset = STA25_LSTAT_CAS5_ERROR
                        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
                    Case casstate_cdm_low
                        g_eDevStatusOffset = STA24_LSTAT_CAS5_LOW
                        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
                        g_eDevStatusOffset = STA25_LSTAT_CAS5_ERROR
                        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
                    Case Else
                        g_eDevStatusOffset = STA25_LSTAT_CAS5_ERROR
                        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
                    End Select
                End If
            End If
        End If
    Next

    SDOCdm.CasNbrLogical = 0
    If SDOCdm.CasState <> 0 And SDOCdm.CasState <> 2 Then
        g_eDevStatusOffset = STA41_LSTAT_DIVERT_FAULT
        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
    Else
        g_eDevStatusOffset = STA41_LSTAT_DIVERT_FAULT
        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
    End If
    
    If G_sDeviceStatus <> G_sNewDeviceStatus Then
        Call UpdateStatusMessage
        G_sDeviceStatus = G_sNewDeviceStatus
        g_sGBLOperStatus = PCB3DL.DlGetCharRaw("GBLOperStatus")
        If g_sGBLOperStatus <> "1" Then
            BrowserMaint.DoRefresh
        End If
    End If

End Sub
'==========================================================================================
'函数的功能 ：会计周期状态
'输入参数   ：无
'输出参数   ：无
'返回值     ：无
'作者       ：汪林
'创建时间   :2004.8
'修改日志：
'==========================================================================================
Private Sub S3EDLWaitPeriod_VariableChanged()
    Dim GBLPeriodStatus As String
    
    GBLPeriodStatus = PCB3DL.DlGetCharRaw("GBLPeriodStatus")
    LogInfo "GBLPeriodStatus is now '" & GBLPeriodStatus & "'"
    If GBLPeriodStatus = "O" Then
        G_bPeriodAvailable = True
    Else
        G_bPeriodAvailable = False
    End If
    
    CalculateAvailability
End Sub
Private Sub S3EDLWaitPeriod_VariableInvalid()
    LogError "GBLPeriodStatus is not waitable"
    Text1.Text = "S3EDLWaitPeriod_VariableInvalid"
End Sub
'==========================================================================================
'函数的功能 :检查所有模块状态，并对有故障模块进行复位
'输入参数 :无
'输出参数 : 无
'返回值   :无
'调用函数 :无
'被调用情况  ：每次退卡之后
'作者       ：汪林
'创建时间   :2004.8
'==========================================================================================
Private Sub S3EDLWaitRecovery_VariableChanged()
    Dim i              As Integer
    
    Dim sAtmStatus     As String
    Dim sRecoveryValue As String
    
'Add for handling withdrawal crime and the overlimited status of Reject/Retract bin
    If PCB3DL.DlGetCharRaw("GBLDoRecovery") = "C" Then
        Call SetNewDeviceStatus(STA42_LSTAT_CDM_HDW_ERROR, True)
        Call SetGuideLightDispenser
        Exit Sub
    ElseIf PCB3DL.DlGetCharRaw("GBLDoRecovery") = "O" Then
        Call SetNewDeviceStatus(STA42_LSTAT_CDM_HDW_ERROR, True)
        Call SetNewDeviceStatus(STA47_DEP_RETRACT_FULL, True)
        Call SetGuideLightDispenser
        Exit Sub
    ElseIf PCB3DL.DlGetCharRaw("GBLDoRecovery") = "N" Then
        Call SetNewDeviceStatus(STA47_DEP_RETRACT_FULL, False)
        If SDOCdm.Available Then
            Call SetNewDeviceStatus(STA42_LSTAT_CDM_HDW_ERROR, False)
        End If
        Call SetGuideLightDispenser
        Exit Sub
    End If
'End if
    
    'Retrieve anomalies
    S3EDLWaitAnomalies_VariableChanged
    
    sRecoveryValue = PCB3DL.DlGetCharRaw("GBLCdmRecoveryNeeded")
    
    LogInfo "Recovery started"
    If (G_nDevicesToUse And DEVICE_CDM) Then
        If (Not SDOCdm.Available) Then
            If SDOCdm.OperatorType = optype_cdm_allempty Or _
                    SDOCdm.OperatorType = optype_cdm_rejectcasfull Or _
                    SDOCdm.OperatorType = optype_cdm_rejectcasnotinstalled Or _
                    SDOCdm.OperatorType = optype_cdm_rejectcasnotconfigured Or sRecoveryValue = "N" Then
                LogInfo "    CDM not available for all cassette empty or Reject bin problem, Recovering not needed"
            Else
                LogInfo "    CDM not available --> Recovering"
                sAtmStatus = PCB3DL.DlGetCharRaw("GBLAtmStatus")
                If sAtmStatus = "O" Then
                    i = PCB3DL.DlSetCharRaw("GBLAtmStatus", "C")
                    i = PCB3DL.DlSetCharRaw("GBLIsDoRecoverying", "Y")
                   
                    Call CDMRecovery
                   
                    i = PCB3DL.DlSetCharRaw("GBLAtmStatus", "O")
                    i = PCB3DL.DlSetCharRaw("GBLIsDoRecoverying", "N")
                Else
                    Call CDMRecovery
                End If
            End If
        End If
    End If
    
    If (G_nDevicesToUse And DEVICE_EDM) Then
        If (Not SDOEdm.Available) Then
            LogInfo "    EDM not available --> Recovering"
            SDOEdm.DoRecovery
        End If
    End If
    If (G_nDevicesToUse And DEVICE_FEP) Then
        If (Not SDOFep.Available) Then
            LogInfo "    FEP not available --> Recovering"
            SDOFep.DoRecovery
        End If
    End If
    If (G_nDevicesToUse And DEVICE_IDC) Then
        If (Not SDOIdc.Available) Then
            LogInfo "    IDC not available --> Recovering"
            SDOIdc.DoRecovery
            Call SetGuideLightCardReader
        End If
    End If
    If (G_nDevicesToUse And DEVICE_PRJ) Then
        If (Not SDOPrj.Available) Then
            LogInfo "    PRJ not available --> Recovering"
            SDOPrj.DoRecovery
        End If
    End If
    If (G_nDevicesToUse And DEVICE_PRR) Then
        If (Not SDOPrr.Available) Then
            LogInfo "    PRR not available --> Recovering"
            SDOPrr.DoRecovery
            Call SetGuideLightReceipt
        End If
    End If
    If (G_nDevicesToUse And DEVICE_TTU) Then
        If (Not SDOTtu.Available) Then
            LogInfo "    TTU not available --> Recovering"
            SDOTtu.DoRecovery
        End If
    End If
    If (G_nDevicesToUse And DEVICE_PIN) Then
        If (Not SDOPin.Available) Then
            LogInfo "    TTU not available --> Recovering"
            SDOPin.DoRecovery
        End If
    End If
    

    g_eDevStatusOffset = STA03_HSTAT_OPERATOR_SERVICE
    If PCB3DL.DlGetCharRaw("GBLOperStatus") <> "2" Then
        Call SetNewDeviceStatus(STA01_HSTAT_OUT_OF_SERVICE, True)
        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
    Else
        If PCB3DL.DlGetCharRaw("GBLAtmStatus") = "O" Then
             Call SetNewDeviceStatus(STA01_HSTAT_OUT_OF_SERVICE, False)
        End If
        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
    End If
          
    If G_sDeviceStatus <> G_sNewDeviceStatus Then
        Call UpdateStatusMessage
        G_sDeviceStatus = G_sNewDeviceStatus
        g_sGBLOperStatus = PCB3DL.DlGetCharRaw("GBLOperStatus")
        If g_sGBLOperStatus <> "1" Then
            BrowserMaint.DoRefresh
        End If
    End If
    
    If bIsMultiCurrency Then
        Call QueryAllCurrencyAvailable
    End If
    
    LogInfo "Recovery ended"
  
End Sub

Private Sub S3EDLWaitRecovery_VariableInvalid()
    LogError "GBLDoRecovery is not waitable"
    Text1.Text = "S3EDLWaitRecovery_VariableInvalid"
End Sub

Private Sub S3eDLWaitResetTransKey_VariableChanged()
   Dim GBLKeyStatus As String
   
    GBLKeyStatus = PCB3DL.DlGetCharRaw("ResetTransKey")
    LogInfo "TransKeyStatus is now '" & GBLKeyStatus & "'"
    If GBLKeyStatus = "N" Then
        G_KeyAvailable = True
    Else
        G_KeyAvailable = False
    End If
    
    CalculateAvailability

End Sub

'==========================================================================================
'函数的功能 ：重新启动NT系统
'输入参数   ：无
'输出参数   ：无
'返回值     ：无
'调用函数   ：无
'被调用情况  ：半夜重新启动时
'作者       ：汪林
'创建时间   :2004.8
'修改日志： 2005 7 增加GBLSysShutDown变量，以前系统半夜重新启动时并不判断是否有交易正在进行
'          在3030上会出现半夜时正在存款，但系统重新启动的现象。GBLSysShutDown变量只在系统进入
'          outofservice(即没有人做交易时)时会变为"S"，系统才会重新启动。
'==========================================================================================
Private Sub S3EDLWaitSysShutDown_VariableChanged()
    Dim sSysShutDownFlag As String
    Dim hS3EStartStopEvent As Long
    
    sSysShutDownFlag = PCB3DL.DlGetCharRaw("GBLSysShutDown")
    Select Case sSysShutDownFlag
    Case "I"
        LogWarning "GBLSysShutDown Init!"
    Case "P"
        LogWarning "System Will Be Shutdown While The Application Go To Out Of Service..."
        Text1.Text = "System Will Be Shutdown ......"
    Case "S"
        LogError "System Start To Reboot ......"
        hS3EStartStopEvent = OpenEvent(EVENT_MODIFY_STATE, False, "S3EStartStopEvent")
        If hS3EStartStopEvent <> 0 Then
            SetEvent hS3EStartStopEvent
            CloseHandle hS3EStartStopEvent
        Else
            LogError "Failed to open S3EStartStopEvent!!"
            
            nrc = NTSystemShutDown(EWX_FORCE + EWX_REBOOT)
            If nrc <> 0 Then
                LogError "Call System function <ExitWindowsEx->EWX_REBOOT> Failed"
            Else
                LogError "Call System function <ExitWindowsEx->EWX_REBOOT> OK"
            End If
'            'ShutDown Failed. I have to reset the DataLink variable
'            PCB3DL.DlSetCharRaw "GBLSysShutDown", "I"
'            Timer1.Enabled = True
        End If
    Case Else
        LogWarning "GBLSysShutDown Value Is Unknown: " + sSysShutDownFlag
    End Select

End Sub

Private Sub S3EDLWaitSysShutDown_VariableInvalid()
    LogError "GBLSysShutDown is not waitable"
    Text1.Text = "S3EDLWaitSysShutDown_VariableInvalid"
End Sub

Private Sub SDOCdm_CasStateChanged(ByVal CasNbrLogical As Integer, ByVal OldState As SDOCdmLibCtl.tCdmCasState, ByVal NewState As SDOCdmLibCtl.tCdmCasState)
On Error Resume Next
    
    Dim nPosition    As Integer
    Dim nCasPosition As Integer
    Dim sCasPosition As String
    
    If CasNbrLogical = 0 Or CasNbrLogical = 100 Then
        If NewState <> casstate_cdm_ok And NewState <> casstate_cdm_high Then
            g_eDevStatusOffset = STA41_LSTAT_DIVERT_FAULT
            Call SetNewDeviceStatus(g_eDevStatusOffset, True)
        Else
            g_eDevStatusOffset = STA41_LSTAT_DIVERT_FAULT
            Call SetNewDeviceStatus(g_eDevStatusOffset, False)
        End If
    Else
        SDOCdm.DataCriteria = 1
        SDOCdm.CasNbrLogical = CasNbrLogical
        'Try to get a Numeric byte
        sCasPosition = SDOCdm.CasPosition
        nCasPosition = GetPhysicalCasNbr(sCasPosition)
        If nCasPosition > 0 And nCasPosition < 5 Then
            nPosition = nCasPosition
                Select Case NewState
                    Case casstate_cdm_ok
                        g_eDevStatusOffset = STA33_LSTAT_CAS1_LOW + nPosition - 1
                        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
                        g_eDevStatusOffset = STA37_LSTAT_CAS1_ERROR + nPosition - 1
                        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
                         'for BOC
                        PCB3DL.DlSetCharRaw "CashBoxSts" & CStr(nPosition), "0"
                    Case casstate_cdm_low
                        g_eDevStatusOffset = STA33_LSTAT_CAS1_LOW + nPosition - 1
                        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
                        g_eDevStatusOffset = STA37_LSTAT_CAS1_ERROR + nPosition - 1
                        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
                         'for BOC
                        PCB3DL.DlSetCharRaw "CashBoxSts" & CStr(nPosition), "1"
                    Case Else
                        g_eDevStatusOffset = STA37_LSTAT_CAS1_ERROR + nPosition - 1
                        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
                        
                        'for BOC
                        If NewState = casstate_cdm_empty Then
                            PCB3DL.DlSetCharRaw "CashBoxSts" & CStr(nPosition), "2"
                        ElseIf NewState = casstate_cdm_inoperative Then
                            PCB3DL.DlSetCharRaw "CashBoxSts" & CStr(nPosition), "3"
                        Else
                            PCB3DL.DlSetCharRaw "CashBoxSts" & CStr(nPosition), "4"
                        End If
                End Select
           Else
                Select Case NewState
                    Case casstate_cdm_ok
                        g_eDevStatusOffset = STA24_LSTAT_CAS5_LOW
                        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
                        g_eDevStatusOffset = STA25_LSTAT_CAS5_ERROR
                        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
                    Case casstate_cdm_low
                        g_eDevStatusOffset = STA24_LSTAT_CAS5_LOW
                        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
                        g_eDevStatusOffset = STA25_LSTAT_CAS5_ERROR
                        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
                    Case Else
                        g_eDevStatusOffset = STA25_LSTAT_CAS5_ERROR
                        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
                End Select
            End If
        End If
    
    If G_sDeviceStatus <> G_sNewDeviceStatus Then
        Call UpdateStatusMessage
        G_sDeviceStatus = G_sNewDeviceStatus
        g_sGBLOperStatus = PCB3DL.DlGetCharRaw("GBLOperStatus")
        If g_sGBLOperStatus <> "1" Then
            If Browser.HasSecondMonitor <> 0 Then
                BrowserMaint.DoRefresh
            End If
        End If
    End If
End Sub

Private Sub SDOCdm_DevStateChanged()
    Dim Msg          As String
    Dim liv_Loop     As Integer
    Dim nCasPosition As Integer
    Dim sCasPosition As String
    
    g_eDevStatusOffset = STA42_LSTAT_CDM_HDW_ERROR
    
    Msg = "SDOCdm_DevStateChanged (CDM is "
    If SDOCdm.Available Then
        SDOCdm.BackColor = &HFF00&
        Msg = Msg & "available)"
        If PCB3DL.DlGetCharRaw("CWDCrimePossible") = "N" Then
            Call SetNewDeviceStatus(g_eDevStatusOffset, False)
        End If
    Else
        SDOCdm.BackColor = &HFF&
        Msg = Msg & "NOT available)"
        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
    End If
    Call SetGuideLightDispenser      '让取款模块指示灯变色
    
    SDOCdm.DataCriteria = 1
    For liv_Loop = 1 To 4
        SDOCdm.CasNbrLogical = liv_Loop
        'Try to get a Numeric byte
        sCasPosition = SDOCdm.CasPosition
        nCasPosition = GetPhysicalCasNbr(sCasPosition)
        If nCasPosition > 0 And nCasPosition < 5 Then
            Select Case SDOCdm.CasState
                Case casstate_cdm_ok, casstate_cdm_full, casstate_cdm_high
                    PCB3DL.DlSetCharRaw "CashBoxSts" & nCasPosition, "0"
                Case casstate_cdm_low
                    PCB3DL.DlSetCharRaw "CashBoxSts" & nCasPosition, "1"
                Case casstate_cdm_empty
                    PCB3DL.DlSetCharRaw "CashBoxSts" & nCasPosition, "2"
                Case casstate_cdm_inoperative
                    PCB3DL.DlSetCharRaw "CashBoxSts" & nCasPosition, "3"
                Case Else
                    PCB3DL.DlSetCharRaw "CashBoxSts" & nCasPosition, "4"
            End Select
        End If
    Next liv_Loop
    
    LogInfo Msg
    Call CalculateAvailability
    'Call S3EDLWaitAnomalies_VariableChanged   '2005.12.27 为了取款中得到anomaly
End Sub

Private Sub SDODoor_AtDoorPosChanged(ByVal DoorOpen As Boolean)
    Dim PrjString    As String
    Dim PrjCHNString As String

    If DoorOpen Then
        PrjString = DeviceTransExp(" Top-enclosure door was opened.")
        PrjCHNString = DeviceTransExp(" 上箱体门被打开.")
        PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    Else
        PrjString = DeviceTransExp(" Top-enclosure door was closed.")
        PrjCHNString = DeviceTransExp(" 上箱体门被关闭.")
        PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    End If
End Sub

Private Sub SDODoor_AtSafePosChanged(ByVal SafeOpen As Boolean)
    Dim PrjString    As String
    Dim PrjCHNString As String

    g_eDevStatusOffset = STA21_HSTAT_SAFEDOOR_O_C
    If SafeOpen Then
        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
        PrjString = DeviceTransExp(" Safe Door was opened.")
        PrjCHNString = DeviceTransExp(" 下箱体门被打开.")
        PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    Else
        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
        PrjString = DeviceTransExp(" Safe Door was closed.")
        PrjCHNString = DeviceTransExp(" 下箱体门被关闭.")
        PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    End If
    
    Call UpdateStatusMessage
    G_sDeviceStatus = G_sNewDeviceStatus
    
End Sub

Private Sub SDODoor_DevStateChanged()
    Dim Msg As String
    
    Msg = "SDODoor_DevStateChanged (DOOR is "
    If SDODoor.Available Then
        SDODoor.BackColor = &HFF00&
        Msg = Msg & "available)"
        PCB3DL.DlSetCharRaw "SiabALMCode", "0000"
        SDODoor.DoStartTest
    Else
        SDODoor.BackColor = &HFF&
        Msg = Msg & "NOT available)"
    End If
    
    LogInfo Msg
'    S3EDLWaitAnomalies_VariableChanged
End Sub

Private Sub SDOEdm_DevStateChanged()
    Dim Msg As String
    
    Msg = "SDOEdm_DevStateChanged (EDM is "
    If SDOEdm.Available Then
        SDOEdm.BackColor = &HFF00&
        Msg = Msg & "available)"
    Else
        SDOEdm.BackColor = &HFF&
        Msg = Msg & "NOT available)"
        Call SendExceptionMessage("OEX", "2017")
    End If
    LogInfo Msg
    CalculateAvailability
'    S3EDLWaitAnomalies_VariableChanged
End Sub

'added by tyh 2005.7.10 for epp4
Private Sub SDOPin_DevStateChanged()
    Dim Msg As String
    
    Msg = "SDOPin_DevStateChanged (PIN is "
    If SDOPin.Available Then
        SDOPin.BackColor = &HFF00&
        Msg = Msg & "available)"
    Else
        SDOPin.BackColor = &HFF&
        Msg = Msg & "NOT available)"
    End If
    LogInfo Msg
    CalculateAvailability
'    S3EDLWaitAnomalies_VariableChanged
End Sub

Private Sub SDOFep_DevStateChanged()
    Dim Msg As String
    
    Msg = "SDOFep_DevStateChanged (FEP is "
    If SDOFep.Available Then
        SDOFep.BackColor = &HFF00&
        Msg = Msg & "available)"
        PCB3DL.DlSetCharRaw "SiabFEPCode", "0000"
    Else
        SDOFep.BackColor = &HFF&
        Msg = Msg & "NOT available)"
    End If
    LogInfo Msg
    CalculateAvailability
'    S3EDLWaitAnomalies_VariableChanged
End Sub
'<时间>：2005.12.9
'<修改者>：孙世方
'<详细记录>：
'   磁卡读写器故障时打印流水
Private Sub SDOIdc_DevStateChanged()
    Dim Msg     As String
    Dim ExpCode As String
    Dim PrjString    As String
    Dim PrjCHNString As String
    
    g_eDevStatusOffset = STA07_HSTAT_CARD_FAILURE
    
    Msg = "SDOIdc_DevStateChanged (IDC is "
    If SDOIdc.Available Then
        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
        SDOIdc.BackColor = &HFF00&
        Msg = Msg & "available)"
    Else
        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
        SDOIdc.BackColor = &HFF&
        Msg = Msg & "NOT available)"
    
        PrjString = DeviceTransExp(" CardReader Failed")
        PrjCHNString = DeviceTransExp(" 读卡机故障.")
        PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
        Call SendExceptionMessage("OEX", "2010")
    End If
    
    LogInfo Msg
'Add for Tri-color
    Call SetGuideLightCardReader
'Add end
    CalculateAvailability
'    S3EDLWaitAnomalies_VariableChanged
End Sub

'==========================================================================================
'函数的功能 ：处理主机管理命令
'输入参数   ：命令类型
'输出参数   ：无
'返回值     ：无
'调用函数   ：无
'被调用情况  ：接收到主机管理命令后
'作者       ：汪林
'创建时间   :2004.8
'==========================================================================================
Private Sub S3ELineIn1_HostMessageReceived(ByVal nType As Long)
    Dim sText           As String
    Dim sTextCHN        As String
    Dim dTotAmount      As Double
    Dim hS3ELineInEvent As Long
    Dim LineStatus      As String
    
    sText = "---------------------------------------" & Chr(10) & Chr(13)
    sTextCHN = "---------------------------------------" & Chr(10) & Chr(13)
    Select Case nType
        Case 300:     ' Host Open ATM
            sText = sText & "           Host Message Open" & Chr(10) & Chr(13)
            sTextCHN = sTextCHN & "     收到主机开始服务命令！" & Chr(10) & Chr(13)
        Case 301:     ' Host Close ATM
            sText = sText & "           Host Message Close" & Chr(10) & Chr(13)
            sTextCHN = sTextCHN & "     收到主机暂停服务命令！" & Chr(10) & Chr(13)
        Case 501:     ' Host Message Inquiry
            sText = sText & "     Host Message Inquiry" & Chr(10) & Chr(13)
            sTextCHN = sTextCHN & "     收到主机查询交易命令！" & Chr(10) & Chr(13)

            
            hS3ELineInEvent = OpenEvent(EVENT_MODIFY_STATE, False, "S3ELineInEvent")
            If hS3ELineInEvent <> 0 Then
                SetEvent hS3ELineInEvent
                CloseHandle hS3ELineInEvent
            Else
                LogError "Failed to open hS3ELineInEvent"
            End If
        
        Case 502:     ' Host Message Check Status
            sText = sText & " Host Message Check Status" & Chr(10) & Chr(13)
            sTextCHN = sTextCHN & "     收到主机查询状态命令！" & Chr(10) & Chr(13)
           
            Call GetState
            
            hS3ELineInEvent = OpenEvent(EVENT_MODIFY_STATE, False, "S3ELineInEvent")
            If hS3ELineInEvent <> 0 Then
                SetEvent hS3ELineInEvent
                CloseHandle hS3ELineInEvent
            Else
                LogError "Failed to open hS3ELineInEvent"
            End If
        
        Case 800:     ' Line Status is Down
'            G_bLineAvailable = False
            Call SetNewDeviceStatus(g_eDevStatusOffset, False)
            LineStatus = PCB3DL.DlGetCharRaw("GBLLineStatus")
            PCB3DL.DlSetCharRaw "GBLLineStatus", "C"
            S3ELineOut.BackColor = &HFF&
            CalculateAvailability
            sText = sText & "           Line Status is Down" & Chr(10) & Chr(13)
            If LineStatus = "O" Then
                Call SendExceptionMessage("TEX", "2000")
            End If
        Case 801:     ' Line Status is Active
'            G_bLineAvailable = True
            Call SetNewDeviceStatus(g_eDevStatusOffset, True)
            PCB3DL.DlSetCharRaw "GBLLineStatus", "O"
            S3ELineOut.BackColor = &HFF00&
            CalculateAvailability
            sText = sText & "           Line Status is Active" & Chr(10) & Chr(13)
        
        Case Else   ' Unknown host message
                sText = sText & "          Host Message Unknown" & Chr(10) & Chr(13)
    End Select
    sText = sText & "---------------------------------------" & Chr(10) & Chr(13)
    sText = sText & "Mod: Monitor  Time: " & Format(Date$, "YYYY/MM/DD") & " " & Format(Time$(), "HH:MM:SS") & Chr(10) & Chr(13)
    sText = sText & "Bank Code: " & PCB3DL.DlGetCharRaw("GBLBankCode") & "         ATM Code: " & PCB3DL.DlGetCharRaw("GBLAtmCode") & Chr(10) & Chr(13)
    sText = sText & "---------------------------------------"
    
    PrintJournal SDOPrj, sText, sTextCHN, g_sPrjLanguage
    LogInfo sText
End Sub

'==========================================================================================
'针对各模块的statechanged 事件进行处理，包括lineout,ops,pin,prj,prr,ttu,door,edm,fep,idc
'==========================================================================================
Private Sub S3ELineOut_DevStateChanged()
    Dim Msg As String
    Msg = "S3ELineOut_DevStateChanged (LineOut is "
    
    g_eDevStatusOffset = STA02_HSTAT_HOST_IN_COMM
    If S3ELineOut.Available Then
        g_bLineStsChanged = True
'        G_bLineAvailable = True
        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
        PCB3DL.DlSetCharRaw "GBLLineStatus", "O"
        S3ELineOut.BackColor = &HFF00&
        Msg = Msg & "available)"
    Else
'        G_bLineAvailable = False
        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
        PCB3DL.DlSetCharRaw "GBLLineStatus", "C"
        S3ELineOut.BackColor = &HFF&
        Msg = Msg & "NOT available)"
    End If
    LogInfo Msg
    
    CalculateAvailability
End Sub

Private Sub SDOOps_AtKeyPosChanged(ByVal KeyPos As Integer)
    If g_bIsFirstOpsKeyChanged Then
        LogInfo "(First Entry)KeyPos changed to Position " + CStr(KeyPos)
        g_bIsFirstOpsKeyChanged = False
    Else
        nrc = PCB3DL.DlSetCharRaw("GBLOperStatus", "1")
        LogInfo "KeyPos changed to Position " + CStr(KeyPos)
    End If
End Sub

Private Sub SDOOps_DevStateChanged()
    Dim Msg As String
    Msg = "SDOOps_DevStateChanged (OPS is "
    If SDOOps.Available Then
        SDOOps.BackColor = &HFF00&
        Msg = Msg & "available)"
        PCB3DL.DlSetCharRaw "SiabOPKCode", "0000"
        SDOOps.DoStartTest
    Else
        SDOOps.BackColor = &HFF&
        Msg = Msg & "NOT available)"
    End If
    LogInfo Msg
    CalculateAvailability
'    S3EDLWaitAnomalies_VariableChanged
End Sub

Private Sub SDOPrj_DevStateChanged()
    Dim Msg         As String
    Dim ExpCode     As String
    
    g_eDevStatusOffset = STA10_HSTAT_JOUPRT_FAILURE
    
    Msg = "SDOPrj_DevStateChanged (PRJ is "
    If SDOPrj.Available Then
        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
        SDOPrj.BackColor = &HFF00&
        Msg = Msg & "available)"
        PCB3DL.DlSetCharRaw "DevicePRJState", "0"
    Else
        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
        SDOPrj.BackColor = &HFF&
        Msg = Msg & "NOT available)"
       
          Select Case SDOPrj.OperatorType
              Case optype_prj_paper_low, optype_prj_ink_low, optype_prj_retract_high
                  ExpCode = "1"
              Case optype_prj_ink_empty, optype_prj_paper_empty
                  ExpCode = "2"
              Case optype_prj_off_line
                  ExpCode = "3"
              Case optype_prj_retract_full
                  ExpCode = "3"
              Case optype_prj_paper_jammed
                  ExpCode = "3"
              Case Else
                  ExpCode = "3"
          End Select
        PCB3DL.DlSetCharRaw "DevicePRJState", ExpCode
        Call SendExceptionMessage("OEX", "2023")
    End If
    
    If SDOPrj.OperatorType = optype_prr_paper_empty Then
        g_eDevStatusOffset = STA11_HSTAT_JOUPRT_PAPER_OUT
        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
    ElseIf SDOPrj.OperatorType = optype_prr_paper_low Then
        g_eDevStatusOffset = STA43_LSTAT_JOUPRT_PAPER_LOW
        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
    Else
        g_eDevStatusOffset = STA11_HSTAT_JOUPRT_PAPER_OUT
        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
        g_eDevStatusOffset = STA43_LSTAT_JOUPRT_PAPER_LOW
        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
    End If
    
    LogInfo Msg
    CalculateAvailability
'    S3EDLWaitAnomalies_VariableChanged
End Sub
'<时间>：2005.12.9
'<修改者>：孙世方
'<详细记录>：
'收条故障时打印流水
Private Sub SDOPrr_DevStateChanged()
    Dim Msg     As String
    Dim ExpCode As String
    Dim PrjString    As String
    Dim PrjCHNString As String
    
    Msg = "SDOPrr_DevStateChanged (PRR is "
    g_eDevStatusOffset = STA08_HSTAT_RECPRT_FAILURE
    
    If SDOPrr.Available Then
        SDOPrr.BackColor = &HFF00&
        Msg = Msg & "available)"
        PCB3DL.DlSetCharRaw "DevicePRRState", "0"
        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
    Else
        PrjString = DeviceTransExp(" Receipt Printer Failed.")
        PrjCHNString = DeviceTransExp(" 收条打印机故障.")
        PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
        
        SDOPrr.BackColor = &HFF&
        Msg = Msg & "NOT available)"
        Call SetNewDeviceStatus(g_eDevStatusOffset, True)

        Select Case SDOPrr.OperatorType
             Case optype_prr_paper_low, optype_prr_ink_low, optype_prr_retract_high
                ExpCode = "1"
             Case optype_prr_paper_empty
                ExpCode = "2"
             Case optype_prr_ink_empty
                ExpCode = "2"
             Case optype_prr_off_line
                ExpCode = "3"
             Case optype_prr_retract_full
                ExpCode = "3"
             Case optype_prr_paper_jammed
                ExpCode = "3"
             Case Else
                ExpCode = "3"
        End Select
        
        PCB3DL.DlSetCharRaw "DevicePRRState", ExpCode
        Call SendExceptionMessage("OEX", "2022")

    End If
    
    If SDOPrr.OperatorType = optype_prr_paper_empty Then
        g_eDevStatusOffset = STA09_HSTAT_RECPRT_PAPER_OUT
        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
    ElseIf SDOPrr.OperatorType = optype_prr_paper_low Then
        g_eDevStatusOffset = STA44_LSTAT_RECPRT_PAPER_LOW
        Call SetNewDeviceStatus(g_eDevStatusOffset, True)
    Else
        g_eDevStatusOffset = STA09_HSTAT_RECPRT_PAPER_OUT
        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
        g_eDevStatusOffset = STA44_LSTAT_RECPRT_PAPER_LOW
        Call SetNewDeviceStatus(g_eDevStatusOffset, False)
    End If
    
    LogInfo Msg
    Call SetGuideLightReceipt
    CalculateAvailability
'    S3EDLWaitAnomalies_VariableChanged
End Sub
Private Sub SDOTtu_DevStateChanged()
    Dim Msg As String
    Msg = "SDOTtu_DevStateChanged (TTU is "
    If SDOTtu.Available Then
        SDOTtu.BackColor = &HFF00&
        Msg = Msg & "available)"
        PCB3DL.DlSetCharRaw "SiabOPDCode", "0000"
    Else
        SDOTtu.BackColor = &HFF&
        Msg = Msg & "NOT available)"
    End If
    LogInfo Msg
    CalculateAvailability
'    S3EDLWaitAnomalies_VariableChanged
End Sub

'==========================================================================================
'函数的功能 ：定时上送状态报文，重新启动系统
'输入参数   ：无
'输出参数   ：无
'返回值     ：无
'调用函数   ：无
'被调用情况  ：form_load
'作者       ：汪林
'创建时间   :2004.8
'==========================================================================================
Private Sub Timer1_Timer()
Dim sRecoveryValue As String
Dim sCurDetailName As String
Dim curTime As Date
Dim sAtmStatus As String

    If G_nCKLInterval > 0 Then
        G_nCKLCurrentCount = G_nCKLCurrentCount + 1
        If G_nCKLCurrentCount = G_nCKLInterval Then
            G_nCKLCurrentCount = 0
            ' Prepare CKL data
            If S3ELineOut.DevState = 16 Then       'OPERATOR_NEEDED
                LogInfo "CKL: not sending message because S3ELineOut.DevState == OPERATOR_NEEDED"
            Else
                nrc = S3ELineOut.SetData("CapCardNum", Format(PCB3DL.DlGetInt("TotCapCardNum"), "0000"))
                                
                sAtmStatus = PCB3DL.DlGetCharRaw("GBLAtmStatus")
                
                If sAtmStatus = "O" Then
                    nrc = S3ELineOut.SetData("ServiceStatus", "1")
                Else
                    nrc = S3ELineOut.SetData("ServiceStatus", "0")
                End If
                
                Call GetState
                
                Call SendExceptionMessage("TEX", "0000")
                 LogInfo "CKL: Message enqueued"
            End If
        End If
    End If
    
    sRecoveryValue = PCB3DL.DlGetCharRaw("GBLCdmRecoveryNeeded")
    
    If G_bAutoReboot Then
        If (Now > G_AutoRebootDateTime And sRecoveryValue = "Y") Then
            PCB3DL.DlSetCharRaw "GBLAtmStatus", "C"
            G_bAutoReboot = False
            
            PCB3DL.DlSetCharRaw "GBLSysShutDown", "P"
            Timer1.Enabled = False
        End If
    End If
    
    If G_bIsStartPeriod = True And Now > G_AutoStartPeriod And PCB3DL.DlGetCharRaw("GBLLineStatus") = "O" Then
                
        LogInfo ("Time to start period checking again")
        Randomize
        G_AutoStartPeriod = DateSerial(Year(Now), Month(Now), Day(Now) + G_nHoursOfStartPeriod \ 24) + _
                TimeSerial(2, Int(Rnd * 60), 0)
        LogInfo "next Start period on " & Format(G_AutoStartPeriod, "DD/MM/YYYY") & _
                    " at " & Format(G_AutoStartPeriod, "HH:MM:SS")
        
    End If
   
    '用于语音提示时间控制
    If G_bIsAudioTimeControl Then
        curTime = Time()
        If curTime > G_AudioStartTime And curTime < G_AudioEndTime Then
            If Not G_bIsAudioEnabled Then
                nrc = PCB3DL.DlSetCharRaw("GBLAudioControl", "Y")
                G_bIsAudioEnabled = True
            End If
        ElseIf G_bIsAudioEnabled Then
            nrc = PCB3DL.DlSetCharRaw("GBLAudioControl", "N")
            G_bIsAudioEnabled = False
        End If
    End If
End Sub

'==========================================================================================
'函数的功能 ：设置DevStatus表
'输入参数   ：DevStatus表中与设备对应的offset值，布尔变量 true - 正常，false -异常
'输出参数   ：无
'返回值     ：无
'作者       ：汪林
'创建时间   :2004.8
'==========================================================================================
Private Sub SetNewDeviceStatus(ByVal DevOffset As Integer, ByVal bValue As Boolean)
    If (bValue) Then
        G_sNewDeviceStatus = Left(G_sNewDeviceStatus, DevOffset - 1) + "1" _
                + Right(G_sNewDeviceStatus, 64 - DevOffset)
    Else
        G_sNewDeviceStatus = Left(G_sNewDeviceStatus, DevOffset - 1) + "0" _
                + Right(G_sNewDeviceStatus, 64 - DevOffset)
    End If
End Sub
'Send status message to host
Private Sub UpdateStatusMessage()
   nrc = PCB3DL.DlSetCharRaw("GBLDevice_State", G_sNewDeviceStatus)
End Sub

'==========================================================================================
'函数的功能 ：open所有使用设备并用绿色表明设备状态正常，红色表明设备状态不正常
'输入参数   ：无
'输出参数   ：无
'返回值     ：无
'调用函数   ：无
'被调用情况  ：form_load
'作者       ：汪林
'创建时间   :2004.8
'==========================================================================================
Private Sub Timer2_Timer()
    On Error GoTo MyErrHandler

    Dim nReply As Integer

    Timer2.Enabled = False
    Timer2.Interval = 2000
    
    If (G_nDevicesToUse And DEVICE_DOOR) And (Not (G_nOpenedDevicesToUse And DEVICE_DOOR)) Then
        Label1.Caption = "Opening DOOR..."
        nReply = SDODoor.PuOpen()
        LogInfo "DOOR.PuOpen returned " & CStr(nReply)
        If nReply = 0 Then
            SDODoor.BackColor = &HFF00&
        Else
            SDODoor.BackColor = &HFF
        End If
        G_nOpenedDevicesToUse = G_nOpenedDevicesToUse + DEVICE_DOOR
    ElseIf (G_nDevicesToUse And DEVICE_EDM) And (Not (G_nOpenedDevicesToUse And DEVICE_EDM)) Then
        Label1.Caption = "Opening EDM..."
        nReply = SDOEdm.PuOpen()
        LogInfo "EDM.PuOpen returned " & CStr(nReply)
        If nReply = 0 Then
            SDOEdm.BackColor = &HFF00&
        Else
            SDOEdm.BackColor = &HFF
        End If
        G_nOpenedDevicesToUse = G_nOpenedDevicesToUse + DEVICE_EDM
    ElseIf (G_nDevicesToUse And DEVICE_FEP) And (Not (G_nOpenedDevicesToUse And DEVICE_FEP)) Then
        Label1.Caption = "Opening FEP..."
        nReply = SDOFep.PuOpen()
        LogInfo "FEP.PuOpen returned " & CStr(nReply)
        If nReply = 0 Then
            SDOFep.BackColor = &HFF00&
        Else
            SDOFep.BackColor = &HFF
        End If
        G_nOpenedDevicesToUse = G_nOpenedDevicesToUse + DEVICE_FEP
    'added by tyh 2005.7.10 for epp4
    ElseIf (G_nDevicesToUse And DEVICE_PIN) And (Not (G_nOpenedDevicesToUse And DEVICE_PIN)) Then
        Label1.Caption = "Opening PIN..."
        nReply = SDOPin.PuOpen()
        LogInfo "PIN.PuOpen returned " & CStr(nReply)
        If nReply = 0 Then
            SDOPin.BackColor = &HFF00&
        Else
            SDOPin.BackColor = &HFF
        End If
        G_nOpenedDevicesToUse = G_nOpenedDevicesToUse + DEVICE_PIN
    'added end
    ElseIf (G_nDevicesToUse And DEVICE_IDC) And (Not (G_nOpenedDevicesToUse And DEVICE_IDC)) Then
        Label1.Caption = "Opening IDC..."
        nReply = SDOIdc.PuOpen()
        LogInfo "IDC.PuOpen returned " & CStr(nReply)
        If nReply = 0 Then
            SDOIdc.BackColor = &HFF00&
        Else
            SDOIdc.BackColor = &HFF
        End If
        G_nOpenedDevicesToUse = G_nOpenedDevicesToUse + DEVICE_IDC
    ElseIf (G_nDevicesToUse And DEVICE_OPS) And (Not (G_nOpenedDevicesToUse And DEVICE_OPS)) Then
        Label1.Caption = "Opening OPS..."
        nReply = SDOOps.PuOpen()
        LogInfo "OPS.PuOpen returned " & CStr(nReply)
        If nReply = 0 Then
            SDOOps.BackColor = &HFF00&
        Else
            SDOOps.BackColor = &HFF
        End If
        G_nOpenedDevicesToUse = G_nOpenedDevicesToUse + DEVICE_OPS
    ElseIf (G_nDevicesToUse And DEVICE_PRJ) And (Not (G_nOpenedDevicesToUse And DEVICE_PRJ)) Then
        Label1.Caption = "Opening PRJ..."
        nReply = SDOPrj.PuOpen()
        LogInfo "PRJ.PuOpen returned " & CStr(nReply)
        If nReply = 0 Then
            SDOPrj.BackColor = &HFF00&
        Else
            SDOPrj.BackColor = &HFF
        End If
        G_nOpenedDevicesToUse = G_nOpenedDevicesToUse + DEVICE_PRJ
    ElseIf (G_nDevicesToUse And DEVICE_PRR) And (Not (G_nOpenedDevicesToUse And DEVICE_PRR)) Then
        Label1.Caption = "Opening PRR..."
        nReply = SDOPrr.PuOpen()
        LogInfo "PRR.PuOpen returned " & CStr(nReply)
        If nReply = 0 Then
            SDOPrr.BackColor = &HFF00&
        Else
            SDOPrr.BackColor = &HFF
        End If
        G_nOpenedDevicesToUse = G_nOpenedDevicesToUse + DEVICE_PRR
    ElseIf (G_nDevicesToUse And DEVICE_TTU) And (Not (G_nOpenedDevicesToUse And DEVICE_TTU)) Then
        Label1.Caption = "Opening TTU..."
        nReply = SDOTtu.PuOpen()
        LogInfo "TTU.PuOpen returned " & CStr(nReply)
        If nReply = 0 Then
            SDOTtu.BackColor = &HFF00&
            SDOTtu.DoForm "StartLine10", True
        Else
            SDOTtu.BackColor = &HFF
        End If
        G_nOpenedDevicesToUse = G_nOpenedDevicesToUse + DEVICE_TTU
    ElseIf (G_nDevicesToUse And DEVICE_CDM) And (Not (G_nOpenedDevicesToUse And DEVICE_CDM)) Then
        Label1.Caption = "Opening CDM..."
        nReply = SDOCdm.PuOpen()
        LogInfo "CDM.PuOpen returned " & CStr(nReply)
        If nReply = 0 Then
            SDOCdm.BackColor = &HFF00&
        Else
            SDOCdm.BackColor = &HFF
        End If
        G_nOpenedDevicesToUse = G_nOpenedDevicesToUse + DEVICE_CDM
    End If
            
    nTimerSequence = nTimerSequence + 1
    If nTimerSequence < G_nDevicesCountToUse + 1 Then
        Timer2.Enabled = True
    End If
    Exit Sub
    
MyErrHandler:
    LogError "Error: " & Err.Number & " " & Err.Description & " " & Err.Source
    ' Continue running the program. You may want to implement other techniques here.
    If (G_nDevicesToUse And DEVICE_CDM) And (Not (G_nOpenedDevicesToUse And DEVICE_CDM)) Then
        G_nOpenedDevicesToUse = G_nOpenedDevicesToUse + DEVICE_CDM
    End If
    
    nTimerSequence = nTimerSequence + 1
    If nTimerSequence < G_nDevicesCountToUse + 1 Then
        Timer2.Enabled = True
    End If

End Sub
'==========================================================================================
'函数的功能 ：从Global.ini获取有关定时上送报文、自动重启的设置，赋值给相关变量
'输入参数   ：无
'输出参数   ：无
'返回值     ：布尔变量
'调用函数   ：无
'被调用情况  ：form_load
'作者       ：汪林
'创建时间   :2004.8
'==========================================================================================
Private Function InitializeCounters() As Boolean
    Dim dStartTime     As Date
    Dim dEndTime       As Date
    Dim nNumberOfDays  As Long
    Dim StartTime      As String
    Dim EndTime        As String
    Dim nResult        As Long
    Dim sValue         As String
    Dim curTime        As Date
    
    LogInfo "InitializeCounters"

    ' Retrieve configuration information for the CKL message
    G_nCKLCurrentCount = 0
    G_nCKLInterval = GetPrivateProfileInt("Interval", "CKL", 0, sGlobalIni)
    LogInfo "    G_nCKLInterval is " & CStr(G_nCKLInterval)

    G_nHoursOfStartPeriod = 0
    G_nHoursOfStartPeriod = GetPrivateProfileInt("StartPeriod", "IntervalOfHours", 0, sGlobalIni)
    If G_nHoursOfStartPeriod <> 0 Then
        Randomize
        G_AutoStartPeriod = DateSerial(Year(Now), Month(Now), Day(Now) + G_nHoursOfStartPeriod \ 24) + _
                TimeSerial(2, Int(Rnd * 59), 0)
        LogInfo "Start period on " & Format(G_AutoStartPeriod, "DD/MM/YYYY") & _
                    " at " & Format(G_AutoStartPeriod, "HH:MM:SS")
        G_bIsStartPeriod = True
    Else
        G_bIsStartPeriod = False
    End If
    
    nNumberOfDays = GetPrivateProfileInt("AutoReboot", "IntervalOfDays", 0, sGlobalIni)
    If nNumberOfDays = 0 Then
        LogInfo "[AutoReboot] IntervalOfDays is 0 or missing --> No AutoReboot"
        G_bAutoReboot = False
    Else
        StartTime = String(64, " ")
        nResult = GetPrivateProfileString("AutoReboot", "StartTime", "", _
                                          StartTime, Len(StartTime), sGlobalIni)
        StartTime = Left(StartTime, nResult)
        If Len(StartTime) = 0 Then
            LogError "[AutoReboot] StartTime is empty or missing --> No AutoReboot"
            G_bAutoReboot = False
        Else
            EndTime = String(64, " ")
            nResult = GetPrivateProfileString("AutoReboot", "EndTime", "", _
                                              EndTime, Len(EndTime), sGlobalIni)
            EndTime = Left(EndTime, nResult)
            If Len(EndTime) = 0 Then
                LogError "[AutoReboot] EndTime is empty or missing --> No AutoReboot"
                G_bAutoReboot = False
            Else
                G_AutoRebootDateTime = DateSerial(Year(Now), Month(Now), Day(Now) + nNumberOfDays)
              
                dStartTime = TimeValue(StartTime)
                dEndTime = TimeValue(EndTime)
                If dStartTime > dEndTime Then
                    LogError "StartTime > EndTime --> No AutoReboot"
                    G_bAutoReboot = False
                Else
                    Dim nIntervalSize As Integer
                    nIntervalSize = (Hour(dEndTime) - Hour(dStartTime)) * 60 + _
                                    Minute(dEndTime) - Minute(dStartTime)
                    Randomize
                    G_AutoRebootDateTime = G_AutoRebootDateTime + _
                                           TimeSerial(Hour(dStartTime), Minute(dStartTime) + Rnd * nIntervalSize, 0)
                    G_bAutoReboot = True
                    LogInfo "Auto Reboot on " & Format(G_AutoRebootDateTime, "DD/MM/YYYY") & _
                                " at " & Format(G_AutoRebootDateTime, "HH:MM:SS")
                End If
            End If
        End If
    End If
    
'Add for Shutdown system
    If G_bAutoReboot Then
        PCB3DL.DlSetCharRaw "GBLSysShutDown", "I"
    End If

'用于语音提示时间控制
    sValue = GetIniS(sGlobalIni, "AudioControl", "AudioConfig", "N")
    LogInfo "Global's AudioConfig = " & sValue
    G_bIsAudioTimeControl = False
    If sValue = "T" Then
        G_bIsAudioTimeControl = True
        
        nrc = PCB3DL.DlSetCharRaw("GBLAudioControl", "Y")
        G_bIsAudioEnabled = True
        
        sValue = GetIniS(sGlobalIni, "AudioControl", "StartTime", "6:00")
        G_AudioStartTime = TimeValue(sValue)
        sValue = GetIniS(sGlobalIni, "AudioControl", "EndTime", "22:00")
        G_AudioEndTime = TimeValue(sValue)
        
        curTime = Time()
        If curTime > G_AudioStartTime And curTime < G_AudioEndTime Then
            nrc = PCB3DL.DlSetCharRaw("GBLAudioControl", "Y")
            G_bIsAudioEnabled = True
        ElseIf G_bIsAudioEnabled Then
            nrc = PCB3DL.DlSetCharRaw("GBLAudioControl", "N")
            G_bIsAudioEnabled = False
        End If
        
    ElseIf sValue = "Y" Then
        'Enable audio help
        nrc = PCB3DL.DlSetCharRaw("GBLAudioControl", "Y")
    Else
        'Disable audio help
        nrc = PCB3DL.DlSetCharRaw("GBLAudioControl", "N")
    End If
    
    If G_bAutoReboot Or G_nCKLInterval <> 0 Or G_bIsAudioTimeControl Then
        InitializeCounters = True
    Else
        InitializeCounters = False
    End If

    LogInfo "InitializeCounters " & InitializeCounters
End Function

'==========================================================================================
'函数的功能 ：从注册表得到正在使用设备的个数
'输入参数   ：从注册表读取的使用设备表
'输出参数   ：无
'返回值     ：正在使用设备的个数
'调用函数   ：无
'被调用情况  ：form_load
'作者       ：汪林
'创建时间   :2004.8
'==========================================================================================
Function CaculateDeviceNumber(ByVal DeviceRegNum As Long) As Byte
   Dim HexString   As String
   Dim HexLen      As Byte
   Dim HexBit      As Byte
   Dim DivResult   As Byte
   Dim ModResult   As Byte
   Dim SumOneCount As Byte
   Dim i           As Byte
   
   SumOneCount = 0
   HexString = Hex(DeviceRegNum)
   HexLen = Len(HexString)
   For i = 1 To HexLen
       HexBit = Val("&H" + Mid(HexString, i, 1))
       If HexBit <> 0 Then
          DivResult = HexBit
          Do While DivResult <> 0
             ModResult = DivResult Mod 2
             If ModResult = 1 Then
                SumOneCount = SumOneCount + 1
             End If
             DivResult = DivResult \ 2
          Loop
       End If
   Next i
   CaculateDeviceNumber = SumOneCount

End Function
'==========================================================================================
'函数的功能 :组织例外打印流水
'输入参数 :无
'输出参数 :无
'返回值   :无
'调用函数 :无
'被调用情况  ：
'作者       ：赵文明
'创建时间   :2004
'==========================================================================================
Function DeviceTransExp(ByVal ExplainWords As String) As String
    Dim TheTime As String
    
    TheTime = Format(Now(), "YY/MM/DD HH:MM:SS")
    DeviceTransExp = "***  " + TheTime + ExplainWords + vbCrLf

End Function
'==========================================================================================
'函数的功能 :得到钞箱面值（用于多币种）
'输入参数 :无
'输出参数 :无
'返回值   :无
'调用函数 :无
'被调用情况  ：
'作者       ：赵文明
'创建时间   :2004
'==========================================================================================
Private Sub QueryAllCurrencyAvailable()
    Dim AvailOfCNY      As String
    Dim AvailOfHKD      As String
    Dim nNumOfBoxesUsed As Byte
    Dim j               As Byte

    AvailOfCNY = "N"
    AvailOfHKD = "N"
    
    If SDOCdm.Available = False Then
        nrc = PCB3DL.DlSetCharRaw("GBLCashAvailCNY", AvailOfCNY)
        nrc = PCB3DL.DlSetCharRaw("GBLCashAvailHKD", AvailOfHKD)
        Exit Sub
    End If
    
    SDOCdm.DataCriteria = 1
    
    nNumOfBoxesUsed = SDOCdm.NbrOfBoxesUsed
    
    For j = 1 To nNumOfBoxesUsed
        SDOCdm.CasNbrLogical = j
         If (CStr(SDOCdm.CasState) Like "[!4-9]") Then
            Select Case SDOCdm.CasCurrency
                Case "CNY"
                    AvailOfCNY = "Y"
                Case "HKD"
                    AvailOfHKD = "Y"
            End Select
        End If
    Next j

    nrc = PCB3DL.DlSetCharRaw("GBLCashAvailCNY", AvailOfCNY)
    nrc = PCB3DL.DlSetCharRaw("GBLCashAvailHKD", AvailOfHKD)

End Sub
'==========================================================================================
'函数的功能 :用于上送状态报文
'输入参数 :无
'输出参数 : 无
'返回值   :无
'调用函数 :无
'被调用情况  ：
'作者       ：孙世方
'创建时间   :2004

'上送打印机状态错
'修改monitor.frm 函数中的GetState函数，添加case 0 的处理。

'==========================================================================================
Private Sub GetState()
   Dim ExpCode       As String
   Dim liv_Loop      As Integer
   Dim nCasPosition  As Integer
   Dim sCasPosition  As String
    
   Select Case SDOPrj.OperatorType
        Case 0
            ExpCode = "0"
        Case optype_prj_paper_low, optype_prj_ink_low, optype_prj_retract_high
            ExpCode = "1"
        Case optype_prj_ink_empty, optype_prj_paper_empty
            ExpCode = "2"
        Case optype_prj_off_line
            ExpCode = "3"
        Case optype_prj_retract_full
            ExpCode = "3"
        Case optype_prj_paper_jammed
            ExpCode = "3"
        Case Else
            ExpCode = "3"
    End Select
    PCB3DL.DlSetCharRaw "DevicePRJState", ExpCode
                
    Select Case SDOPrr.OperatorType
        Case 0
           ExpCode = "0"
        Case optype_prr_paper_low, optype_prr_ink_low, optype_prr_retract_high
           ExpCode = "1"
        Case optype_prr_paper_empty
           ExpCode = "2"
        Case optype_prr_ink_empty
           ExpCode = "2"
        Case optype_prr_off_line
           ExpCode = "3"
        Case optype_prr_retract_full
           ExpCode = "3"
        Case optype_prr_paper_jammed
           ExpCode = "3"
        Case Else
           ExpCode = "3"
    End Select
        
    PCB3DL.DlSetCharRaw "DevicePRRState", ExpCode
        
    SDOCdm.DataCriteria = 1
    For liv_Loop = 1 To 4
        SDOCdm.CasNbrLogical = liv_Loop
        'Try to get a Numeric byte
        sCasPosition = SDOCdm.CasPosition
        nCasPosition = GetPhysicalCasNbr(sCasPosition)
        If nCasPosition > 0 And nCasPosition < 5 Then
            Select Case SDOCdm.CasState
                Case casstate_cdm_ok, casstate_cdm_full, casstate_cdm_high
                    PCB3DL.DlSetCharRaw "CashBoxSts" & nCasPosition, "0"
                Case casstate_cdm_low
                    PCB3DL.DlSetCharRaw "CashBoxSts" & nCasPosition, "1"
                Case casstate_cdm_empty
                    PCB3DL.DlSetCharRaw "CashBoxSts" & nCasPosition, "2"
                Case casstate_cdm_inoperative
                    PCB3DL.DlSetCharRaw "CashBoxSts" & nCasPosition, "3"
                Case Else
                    PCB3DL.DlSetCharRaw "CashBoxSts" & nCasPosition, "4"
            End Select
        End If
    Next liv_Loop
                
End Sub
'==========================================================================================
'函数的功能 :设置取款灯状态
'         废钞箱满  黄灯长亮   CWDCrimePossible = O
'         犯罪      红灯快闪   CWDCrimePossible = Y
'         模块正常  关闭灯
'         模块不正常 红灯长亮
'输入参数 :无
'输出参数 : 无
'返回值   :无
'调用函数 :无
'被调用情况  ：
'作者       ：汪林
'创建时间   :2004.8
'==========================================================================================
Private Sub SetGuideLightDispenser()
    If PCB3DL.DlGetCharRaw("CWDCrimePossible") = "O" Then
        SDOFep.GuidLightColor = color_amber
        nrc = SDOFep.SetGuidLight(gl_notesdispenser, gl_continuous)
        SDOFep.GuidLightColor = color_red
    ElseIf PCB3DL.DlGetCharRaw("CWDCrimePossible") = "Y" Then
        nrc = SDOFep.SetGuidLight(gl_notesdispenser, gl_quickflash)
    ElseIf SDOCdm.Available Then
        nrc = SDOFep.SetGuidLight(gl_notesdispenser, gl_off)
    Else
        nrc = SDOFep.SetGuidLight(gl_notesdispenser, gl_continuous)
    End If
End Sub
'==========================================================================================
'函数的功能 :设置读卡器灯状态
'         模块正常  关闭灯
'         模块不正常 红灯长亮
'输入参数 :无
'输出参数 : 无
'返回值   :无
'调用函数 :无
'被调用情况  ：
'作者       ：汪林
'创建时间   :2004.8
'==========================================================================================
Private Sub SetGuideLightCardReader()
    If SDOIdc.Available Then
        nrc = SDOFep.SetGuidLight(gl_cardunit, gl_off)
    Else
        nrc = SDOFep.SetGuidLight(gl_cardunit, gl_continuous)
    End If
End Sub
'==========================================================================================
'函数的功能 :设置收条灯状态
'         模块正常  关闭灯
'         模块不正常 红灯长亮
'输入参数 :无
'输出参数 : 无
'返回值   :无
'调用函数 :无
'被调用情况  ：
'作者       ：汪林
'创建时间   :2004.8
'==========================================================================================
Private Sub SetGuideLightReceipt()
    If SDOPrr.Available Then
        nrc = SDOFep.SetGuidLight(gl_receiptprinter, gl_off)
    Else
        nrc = SDOFep.SetGuidLight(gl_receiptprinter, gl_continuous)
    End If
End Sub
'==========================================================================================
'函数的功能 :得到钞箱物理位置
'输入参数 :无
'输出参数 : 无
'返回值   :无
'调用函数 :无
'被调用情况  ：
'作者       ：汪林
'创建时间   :2004.8
'==========================================================================================
Function GetPhysicalCasNbr(ByVal sCasPosition As String) As Integer
    Dim nCasPosition    As Integer
    Dim nLoop           As Integer
    Dim nLenCasPosition As Integer
    Dim sEachByte       As String

    nLenCasPosition = Len(sCasPosition)
    nCasPosition = -1
    If nLenCasPosition Then
        For nLoop = 1 To nLenCasPosition
             sEachByte = Mid(sCasPosition, nLoop, 1)
             If IsNumeric(sEachByte) Then
                  nCasPosition = CInt(sEachByte)
                  Exit For
             End If
        Next nLoop
    End If
    GetPhysicalCasNbr = nCasPosition
End Function
'==========================================================================================
'函数的功能 :当每次取款最大出钞数未设置时，设置为默认值30
'输入参数 :无
'输出参数 : 无
'返回值   :无
'调用函数 :无
'被调用情况  ：Form_load
'作者       ：李军
'创建时间   :2004.8
'==========================================================================================
Private Sub CheckMaxBills()
    Dim sValue As String

    sValue = GetIniS(sGlobalIni, "Withdrawal", "MaxBills", "0")
    If CInt(sValue) < 1 Or CInt(sValue) > 40 Then
        nrc = SetIniS(sGlobalIni, "Withdrawal", "MaxBills", "30")
    End If

End Sub
'==========================================================================================
'函数的功能 :取款模块复位
'输入参数 :无
'输出参数 : 无
'返回值   :无
'调用函数 :无
'被调用情况  ：S3EDLWaitRecovery_VariableChanged
'作者       ：孙世方
'创建时间   :2005.8 26
'==========================================================================================
Private Sub CDMRecovery()
    Dim RecoveryTimes  As Long
    
    RecoveryTimes = PCB3DL.DlGetInt("GBLCdmRecoveryTimes")
    
    If RecoveryTimes <> 0 Then
        nrc = SDOCdm.DoRecovery
        If nrc <> 0 Then
            RecoveryTimes = RecoveryTimes - 1
        Else
            RecoveryTimes = 3
        End If
        Call SetGuideLightDispenser       '让取款模块指示灯变色
    End If
    
    PCB3DL.DlSetLong "GBLCdmRecoveryTimes", RecoveryTimes
    
End Sub
'Add for BOC
Private Sub SendExceptionMessage(ByVal TranCode As String, ByVal ExpCode As String)
    On Error GoTo ErrorTrap
    
    Dim filedata As WIN32_FIND_DATA
    Dim lFileSize As Long
    Dim sTraceMsg As String
        
    If PCB3DL.DlGetCharRaw("GBLLineStatus") = "O" Then
        nrc = S3ELineOut.SetData("ExceptionCode", ExpCode)
        nrc = S3ELineOut.DoSend(TranCode, 1)
    Else
        If g_bLineStsChanged Then
        '如果g_bLineStsChanged = true,表明线路已经发送过OpenMessage成功
        '如果g_bLineStsChanged = false,表明线路没有发送过OpenMessage成功
            lFileSize = 0
            filedata = Findfile("C:\S3ELOut.rcv")        ' Get information
            If filedata.nFileSizeHigh <= 32 Then
                lFileSize = filedata.nFileSizeLow
            Else
                lFileSize = filedata.nFileSizeHigh
            End If
        
            If lFileSize = 0 Then
                'It means that MSG queue is empty
                nrc = S3ELineOut.SetData("ExceptionCode", ExpCode)
                nrc = S3ELineOut.DoSend(TranCode, 1)
            Else
                LogWarning "Line not available and MSG Queue not empty, " + TranCode + " Message not enqueued"
            End If
        End If
    End If
    Exit Sub

ErrorTrap:
    'Log unanticipated error message.
    sTraceMsg = "SendExceptionMessage ==> Error " + CStr(Err.Number) + ": " + Err.Description
    LogError sTraceMsg
    Exit Sub
End Sub

