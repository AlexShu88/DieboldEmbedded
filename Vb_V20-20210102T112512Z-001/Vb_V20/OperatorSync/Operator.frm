VERSION 5.00
Object = "{B2110643-3E81-11D3-8ACC-00C04FF20A5D}#1.2#0"; "TransProv.dll"
Object = "{5C094E41-67D2-11D0-AC6B-0020AFBDD1D4}#1.0#0"; "SDOCdm.ocx"
Object = "{EACE4ECF-6930-11D0-AC6C-0020AFBDD1D4}#1.0#0"; "SDOEdm.ocx"
Object = "{192DFCF0-F664-11D3-8BD4-00C04FF20A5D}#1.1#0"; "AdvBrowser.ocx"
Object = "{3751B5D1-D348-11D0-AD02-0060970C3D2F}#1.0#0"; "SDOPrr.ocx"
Object = "{292DFCF0-F664-11D3-8BD4-00C04FF20A5D}#1.0#0"; "AdvBrowserMaint.ocx"
Object = "{DA559591-71AC-11D3-8B0E-00C04FF20A5D}#1.0#0"; "DlWait.ocx"
Object = "{EACE4ED6-6930-11D0-AC6C-0020AFBDD1D4}#1.0#0"; "SDOFep.ocx"
Object = "{BD8177C0-832C-11CF-BF42-0020AF7093F9}#1.0#0"; "SDOIdc.ocx"
Object = "{9C37E835-6A58-11D1-80C0-0020AF7093F9}#1.2#0"; "Dl.ocx"
Object = "{6C4DD4AB-27D5-11D3-96C4-000000000000}#1.0#0"; "s3elineouttcp.ocx"
Object = "{6580F760-7819-11CF-B86C-444553540000}#1.0#0"; "EZFTP.OCX"
Object = "{E64F71A6-E705-4151-9895-5138B7D67F3A}#1.0#0"; "CHPrj.ocx"
Object = "{F3822055-62E4-4A41-A450-086A3C9B1F79}#1.0#0"; "S3EZip.ocx"
Begin VB.Form Operator 
   Caption         =   "Operator"
   ClientHeight    =   3165
   ClientLeft      =   2910
   ClientTop       =   930
   ClientWidth     =   6885
   Icon            =   "Operator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6885
   WindowState     =   1  'Minimized
   Begin ADVBROWSERMAINTATLLibCtl.AdvBrowserMaint BrowserMaint 
      Height          =   615
      Left            =   2760
      OleObjectBlob   =   "Operator.frx":1272
      TabIndex        =   13
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox TxtTransDate 
      DataSource      =   "DataWTH"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Text            =   "0101"
      Top             =   1560
      Width           =   975
   End
   Begin VB.Data DataWTH 
      Caption         =   "DataWTH"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   465
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Width           =   1920
   End
   Begin S3EZIPLib.S3EZip S3EZip 
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   2640
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin CHPRJLib.CHPrj SDOPrj 
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   873
      _StockProps     =   1
   End
   Begin EZFTPLib.EZFTP EZFTP 
      Left            =   1440
      Top             =   1920
      _Version        =   65536
      _ExtentX        =   800
      _ExtentY        =   800
      _StockProps     =   0
      LocalFile       =   ""
      RemoteFile      =   ""
      RemoteAddres    =   ""
      UserName        =   ""
      Password        =   ""
      Binary          =   0   'False
   End
   Begin SDOIdcLibCtl.SDOIdc SDOIdc 
      Height          =   495
      Left            =   1440
      OleObjectBlob   =   "Operator.frx":1298
      TabIndex        =   10
      Top             =   720
      Width           =   1095
   End
   Begin S3EADVBROWSERLibCtl.AdvBrowser Browser 
      Height          =   735
      Left            =   3840
      OleObjectBlob   =   "Operator.frx":12CA
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin SDOFepLibCtl.SDOFep SDOFep 
      Height          =   495
      Left            =   240
      OleObjectBlob   =   "Operator.frx":12F0
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin SDOEdmLibCtl.SDOEdm SDOEdm 
      Height          =   495
      Left            =   240
      OleObjectBlob   =   "Operator.frx":131A
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin DLWaitLibCtl.DLWait DLWaitMonType 
      Height          =   375
      Left            =   2760
      OleObjectBlob   =   "Operator.frx":134A
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin SDOPrrLibCtl.SDOPrr SDOPrr 
      Height          =   495
      Left            =   240
      OleObjectBlob   =   "Operator.frx":1394
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin SDOCdmLibCtl.SDOCdm SDOCdm 
      Height          =   495
      Left            =   240
      OleObjectBlob   =   "Operator.frx":13C4
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin TRANSPROVLibCtl.TransactionProvider S3ETrans 
      Height          =   495
      Left            =   1440
      OleObjectBlob   =   "Operator.frx":13FA
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin S3ELINEOUTLib.S3ELineOut S3ELineOut 
      Height          =   435
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   1245
      _Version        =   65536
      _ExtentX        =   2196
      _ExtentY        =   767
      _StockProps     =   1
   End
   Begin VB.Timer TimerAction 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2040
      Top             =   1920
   End
   Begin VB.CommandButton start 
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
      Height          =   690
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   1530
   End
   Begin DLLib.DL Pcb3dl 
      Left            =   3120
      Top             =   960
      _Version        =   65538
      _ExtentX        =   2487
      _ExtentY        =   1191
      _StockProps     =   0
   End
End
Attribute VB_Name = "Operator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==========================================================================================
'��Ȩ˵��:�ϱ���˾�й���������
'�汾�ţ�Agilis power 1.6
'�������ڣ�2005.8
'���ߣ�  ����(��ʼ�棩
'ģ�鹦�ܣ� ����Ա����
'��Ҫ�������书��
' ȫ�ֱ���
'�޸���־
'-----------------------------------------------------------------------
'<ʱ��>��[2005.08.29]
'<�޸���>��������
'<��ǰ�汾>������1.0.16
'<��ϸ��¼>��
'   <����>
'     ����1������Ա���24 - ��ʾ�쳣ȡ���ӡ�쳣ȡ�����(�����ݿ�cwdlog��������)
'<ʱ��>��[2005.09.19]
'<�޸���>��������
'<��ǰ�汾>������1.0.16
'<��ϸ��¼>��
'     ���Ӳ���Ա����23 - superoperator��������֮��ص�����
'     SuperAdminBegin
'==========================================================================================
'<ʱ��>��[2005.11.29]
'<�޸���>��������
'<��ǰ�汾>������1.0.16
'<��ϸ��¼>��
'   1 ������峮����������,�����峮����ǰ�ĸ�����һ����ֵ��ͬ��������������װ��ֵ�ۼ�����
'      ����ǰ�ĸ�������ͬ�����ڵ��ĸ����������͵���������װ��ֵ���Ƚϵ��ĸ�������ǰ����������ֵ���ҵ���ֵ��ͬ���ۼ�һ������
'   2  20�����Ϊ�����Ĺر�ϵͳ����ǰ����������������
'   3  ����Ա������Կ���ӶԽ�����жϣ���ǰ���۳ɹ�������ʾ�ɹ���
'==========================================================================================
'<ʱ��>��[2005.12.9]
'<�޸���>��������
'  �޸�CommunicationSubFunction����
'  1   ��������е�ת�˱����������޸����Ͷ��ʽ����е�NoOfHKTfr��ΪNoOfRMBTfr
'  2   �ӳ�������DenoRefill��CasPresent��NoOfRMBWth��NoOfRMBTfr��������,������ǰһ���ڵ��ۼ��������LastCashFilled��LastCashPresent,LastWithDrawNumber,LastTfrNumber����������¼��һ���ڼӳ���
'  3   TTI����������CasPresent���ֵ��;rej��Ӧ���ۼӷϳ�������
'<ʱ��>��[2005.12.15]
'<�޸���>��������
'    1 �޸�08�������й�translog�ļ��Ƿ���ڵ��ж�
'    2 ��֮ǰ����journal.txt��Ϊ����journal001.txt,��Ϊ�ͻ���Ҫ�������������journal.txt����û�е�
'<ʱ��>��[2005.12.16]
'<�޸���>��������
'<��ǰ�汾>������1.2.16
'   1 �޸ĵ㳮���֮ǰ������ͬ��ֵ����ʱ����������
'   2 ɾ��δ�õı�������ˮ��ӡ�м���---�Ĵ�ӡ�Խ�ʡ��ˮֽ
'ʱ��   :2005.12.14
'����   :����
'      ִ��03����ʱ����ӡ�ս�����
'<ʱ��>��[2005.12.22]
'<�޸���>��������
'<��ǰ�汾>�����պӱ�����Ҫ���豸״̬��ѯ������ͨѶ״̬��ʾ
'<ʱ��>��[2005.12.26]
'<�޸���>��������
'         DSM  ���ӳ���������ʾ(R)������ԭ�������ST,�����е��룩
'==========================================================================================
Private Declare Function OpenEvent Lib "kernel32" Alias "OpenEventA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function SetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetUsbDisk Lib "GetUsbDisk" (ByRef sDrive As String) As Boolean
Private Declare Function Ping Lib "S3EPing" (ByVal iPingTimes As Integer, ByVal sHostName As String) As Integer
Private Declare Function CloseS3EWindow Lib "S3EPing" (ByVal WindowName As String, ByVal ClassName As String) As Integer

Const MonitorWinName           As String = "S3E Monitor"
Const MonitorClassName         As String = "ThunderRT6FormDC"
Const MasterWinName            As String = "Agilis Power(tm) Master"
Const MasterClassName          As String = "#32770"

Enum HowExitConst
    EWX_LogOff = 0
    EWX_REBOOT = 2
    EWX_SHUTDOWN = 1
    EWX_FORCE = 4
    EWX_POWEROFF = 8
End Enum

Private Const JourLineSeprator         As String = "----------------------------------" + vbCrLf
Private Const TransLogFile             As String = "C:\TransLog.TXT"
Private Const CHNJOURNALFile           As String = "C:\S3E\archives\logto\Journal001.TXT"
Private Const CHNJOURNALBAKFile        As String = "Journal001.TXT"
Private Const CardRetainBAKFile        As String = "CardRetain.txt"
'Private Const TransLogBAKFile          As String = "TransLog.TXT"
Private Const DB_LogPath               As String = "C:\S3e\Logs\LogTo\CWDLog.mdb"
Private Const LogBackupPath            As String = "C:\Agilis\Logs\"
Private Const IniPath                  As String = "c:\AtmWosa\ini\"
Private Const CardRetainFile           As String = "C:\s3e\logs\logapp\CardRetain.txt"
Private Const sVersionIni              As String = "c:\AtmWosa\ini\Version.ini"
Private Const sGlobalIni               As String = "C:\ATMWosa\Ini\global.ini"
Private Const sKeyIni                  As String = "C:\ATMWosa\Ini\Key.ini"

Private Const MAXRETRACTNOTESTOTAL     As Integer = 60
Private Const MAXREJECTNOTESTOTAL      As Integer = 120
Private Const ReturnOk                 As Integer = 10
Private Const NUMBEROFKEYS             As Integer = 2
Private Const EVENT_MODIFY_STATE       As Integer = 2
Private Const nNumberOfCassettes       As Integer = 5

Private Const DEVICE_EDM = &H8&
Private Const keySelfService     As String = "Software\SelfService"

Private Enum pageType
    pageNothing
    pageClosePeriod10
    pageClosePeriod20
    pageClosePeriod30
    pageClosePeriod40
    pageOprLogCopy10
    pageOprLogCopy20
    pageOprLogCopy30
    pageOprLogCopy40
    pageOprLogCopy50
    pageFirstPage
    pagePrrPrintTOT10
    '������TOTƾ����ֻһ��ʱ��ʾ
    pagePrrPrintTOT15
    pagePrrPrintTOT20
    pagePrrPrintTOT30
    pagePrrPrintTOT35
    pagePrrPrintTOT40
    pagePrrPrintTOT50
    pageFunChoice
    pageLoadBoxWarning
    pageloadbox10
    pageLoadBox11
    pageLoadBox20
    pageLoadBox30
    pageLoadBox40
    pageLoadBox50
    pageLoadBox55
    pageLoadBox57
    pageLoadBox60
    pageLoadBox70
    pageLoadBox61
    pageNoFunAvail
    pageOpenPeriod10
    pageOpenPeriod20
    pageOpenPeriod30
    pageOperReturn10
    pageOperReturn20
    pageShowBoxStat10
    pageShowBoxStat20
    pageShowBoxStat30
    pageShowDev10
    pageShowDev20
    pagePrintTotal10
    pagePrintTotal20
    pagePrintTotal30
    pageRetainCard10
    pageRetainCard20
    pageRetainCard30
    pageRetainCard40
    pageShutdownSys10
    pageShutdownSys20
    pageShutdownSys30
    pageWarnPNC
    pageWarnPNO
    
    pageUpdateMasterKey10
    pageUpdateMasterKey15
    pageUpdateMasterKey20
    pageUpdateMasterKey30
    pageUpdateMasterKey40
    
    pageOpKeyInput10
    pageOpKeyInput20
    pageOpKeyInput22
    pageOpKeyInput24
    pageOpKeyInput26
    pageOpKeyInput30
    pageOpKeyInput35
    pageOpKeyInput40
    pageOpKeyInput45
    pageOpKeyInput50
    
    pageOpChgPwd10
    pageOpChgPwd20
    pageOpChgPwd30
    pageOpChgPwd40
    pageOpChgPwd50
    pageOpPinInput
    pageOpPassWrong
    
    PageChkVersion10
    pageChkVersion20
    
    pageResumeBox10
    pageResumeBox20
    pageResumeBox30
    pageResumeBox40
    pageResumeBox50
    
    pageSelectCopyDisk10
    pageSelectCopyDisk20
    pageSelectCopyDisk21
    
    pageLogBackup10
    pageLogBackup20
    pageLogBackup30
    pageLogBackup31
    pageLogBackup33
    pageLogBackup35
    pageLogBackup60
        
'    pagePingHost10
'    pagePingHost20
'    pagePingHost30
'    pagePingHost40
    
    pageExitApp10
    pageExitApp20
    pageExitApp30
    pageExitApp31
    pageExitApp40
    pageExitApp41
    
    pageResetATM10
    pageResetATM20
    pageResetATM30
    pageResetATM40
    pageResetATM50
    
    pageEnterVendorMode10
    pageEnterVendorMode20
    
    pageIsUpdatePage
    
    pageShowTransItem10
    pageShowTransItem20
    
    pageTestDispenseNoteForEachCas10
    pageTestDispenseNoteForEachCas20
    pageTestDispenseNoteForEachCas30
    pageTestDispenseNoteForEachCas40
    
    pageInputSupAdminPassword
    pageInputSupAdminPasswordOk
    pageInputSupAdminPasswordFailed
    pageSuperFunctionChoice
    pageSuperSetTerminalLuno
    pageSuperSetBankCode
    pageSuperChangeSuperPassword
    
    pageOpResetTransKey
    pageOpResetTransKey1
    pageOpResetTransKey2
    pageOpResetTransKey3
    
    pageSendRTT
    
    pageDispCWD10
    pageDispCDP20
    pageDispCDP30
    pageReturnOk
    pageCmdList10
    pageScreenError
    pageError
    pageQuit
End Enum
Private currentPage As pageType

Private Type AssortLogType
    AssortTransType                 As String
    AssortDate                      As String * 8
    AssortCardType                  As String
    AssortSerial                    As String
    AssortAmount                    As Long
    AssortAccNo                     As String * 20
    AssortKeepAccFlag               As String
    AssortCashinResult              As String
    AssosrtHostReject               As String
End Type
Dim AssortLog()                     As AssortLogType

Private Type BoxLoadCashType
    BoxCurry        As String
    BoxDenom        As Long
    BoxDisp         As Long
    BoxLeftNum      As Long
    PurgedNotes     As Long
    CasLogicalID    As Byte
    BoxState        As String
    BoxStateCHN     As String
    BoxInitial      As Long
End Type
' The (nNumberOfCassettes + 1)th is used for storing and displaying reject box information
Dim WthCassette(1 To nNumberOfCassettes + 1) As BoxLoadCashType

Dim G_nDevicesToUse               As Long

'Dim IsTransLogExist               As Boolean
Dim IsCardRetainExist             As Boolean
Dim GLbIsNewPeriod                As Boolean
Dim G_bTrides                     As Boolean
Dim g_bIsHardware                 As Boolean
Dim g_bIsFindMore                 As Boolean
Dim g_bIsPrrResetTest             As Boolean
Dim g_bIsResumeBox                As Boolean
Dim g_bIsTranslogMore             As Boolean
Dim SuperAdminBegin               As Boolean

Dim g_nLogLastPos                 As Integer
Dim g_nCurKeyTime                 As Integer
Dim nrc                           As Integer
Dim g_nLogCurPos                  As Integer
Dim g_nFindStartLine              As Integer
Dim g_nRejectCount                As Integer
Dim g_nRetractCount               As Integer
Dim g_iTotalNumberOfDisplay       As Integer

Dim sLogTargetDisk                As String
Dim GLsPeriodStatus               As String
Dim sGLtheTime                    As String
Dim GLarrMasKeys(2)               As String
Dim ThermalLineHead               As String
Dim g_sBackupLogFileName          As String
Dim g_sBackupLogFileStartTime     As String
Dim g_sBackupLogFileEndTime       As String
Dim gSelectOprCommand             As String
Dim g_AtmPRRType                  As String
Dim g_sPrrRawData                 As String
Dim g_sResettingDev               As String
Dim g_sPrjLanguage                As String
Dim TOTPrjString                  As String
Dim AtmCode                       As String
Dim g_vInputDate                  As Variant

'���ƾ�����Ĵ�ӡ����
Dim PrrTOTPrintPageNumber         As Integer
Dim PrrLeftPrintPageNumber        As Integer
Dim PrrPrintPosition              As Integer
Const PrrLineNumber = 13

'2005/1/27
Dim IsPrintAmonalyTrans           As Boolean

'2005/12/9
Dim LastCashFilled(5)             As Long
Dim LastCashPresent(5)            As Long
Dim LastWithDrawNumber            As Long
Dim LastTfrNumber                 As Long
'==========================================================================================
'�����Ĺ��ܣ�������Ա�����������ʱ�Ĵ���
'�����������
'�����������
'���أ���
'���ߣ�����
'����ʱ��   :
'==========================================================================================
Private Sub DLWaitMonType_VariableChanged()
    Dim nMonType         As Long
    Dim i                As Integer
    Dim sDisplayStr      As String
    Dim sCasStatus       As String
    Dim sFindStr         As String
    Dim nFindCurLine     As Integer
    
    nMonType = Pcb3dl.DlGetInt("OptevaMonType")
    
    Select Case nMonType
        Case 1
            Call FlushBoxesStatusRetIsPresent
            
            For i = 1 To nNumberOfCassettes
                sCasStatus = sCasStatus + WthCassette(i).BoxStateCHN + "|" + _
                        WthCassette(i).BoxCurry + "|" + _
                        CStr(WthCassette(i).BoxDenom) + "|" + _
                        CStr(WthCassette(i).BoxInitial) + "|" + _
                        CStr(WthCassette(i).BoxLeftNum) + "|" + _
                        CStr(WthCassette(i).BoxDenom * WthCassette(i).BoxLeftNum) + "|"
            Next
            
            sCasStatus = sCasStatus + CStr(Pcb3dl.DlGetInt("TotWithdrawNum") + _
                    Pcb3dl.DlGetInt("IcbcTotExtraWthNum")) + "|" + _
                    CStr(Pcb3dl.DlGetDouble("TotWithdrawAmount") + _
                    Pcb3dl.DlGetDouble("IcbcTotExtraWthAmount")) + "|" + _
                    CStr(g_nRejectCount) + "|" + _
                    CStr(g_nRetractCount) + "|" + _
                    CStr(Pcb3dl.DlGetInt("TotCapCardNum"))
            
            nrc = Pcb3dl.DlSetCharRaw("OptevaCasStatus", sCasStatus)
            
            nrc = ShowOperScreenMaint("Operator", "OpInCasStatus")
    
        Case 2, 21, 22
            If nMonType = 21 And g_nLogCurPos = 0 Then  'Has not before page, do nothing
                Exit Sub
            ElseIf nMonType = 22 And g_bIsTranslogMore = False Then     'Has not next page, do nothing
                Exit Sub
            ElseIf nMonType = 2 Then    'Show first page
                g_nLogCurPos = 0
            ElseIf nMonType = 21 Then   'Show last page
                g_nLogCurPos = g_nLogCurPos - 10
            ElseIf nMonType = 22 Then   'Show next page
                g_nLogCurPos = g_nLogCurPos + 10
            Else
                Exit Sub
            End If
            
            g_nFindStartLine = 0
            g_bIsFindMore = True
            g_bIsTranslogMore = GetLogRecordsAndRetIsMore(g_nLogCurPos, sDisplayStr)
                    
            nrc = Pcb3dl.DlSetCharRaw("OptevaCasStatus", sDisplayStr)
            
            nrc = ShowOperScreenMaint("Operator", "OpInDetail")
            
        Case 23, 24

            sFindStr = Pcb3dl.DlGetCharRaw("HtmlInput1")
            If nMonType = 24 And g_bIsFindMore = False Then
                Exit Sub
            End If
            g_bIsFindMore = GetLogFindRecordsAndRetIsMore(g_nFindStartLine, sFindStr, sDisplayStr, nFindCurLine)
            g_nFindStartLine = nFindCurLine
            nrc = Pcb3dl.DlSetCharRaw("OptevaCasStatus", sDisplayStr)
            nrc = ShowOperScreenMaint("Operator", "OpFindInDetail")
            
        Case 4
            Call FlushBoxesStatusRetIsPresent
            If g_nRejectCount >= MAXREJECTNOTESTOTAL Or _
                    g_nRetractCount >= MAXRETRACTNOTESTOTAL Then
                nrc = Pcb3dl.DlSetCharRaw("CWDCrimePossible", "O")
                LogWarning ("Reject notes over limited")
                Pcb3dl.DlSetCharRaw "GBLDoRecovery", "O"
            End If
        
        Case Else
            nrc = ShowOperScreenMaint("Operator", "OpInService")
    End Select
End Sub
'==========================================================================================
'�����Ĺ��ܣ�VB����װ��,��ʼ�����ܼ��������ݿ⣬�õ�ת�������Ϣ
'�����������
'�����������
'���أ���
'����
'����ʱ��   :
'==========================================================================================
Private Sub Form_Load()
    Dim sValue As String
    
    sValue = "The version number of " & App.EXEName & ".exe is " & App.Major & "." _
            & App.Minor & ".0." & App.Revision
    
    LogInfo (sValue)
    
    ThermalLineHead = Chr(27) + Chr(22) + Chr(50)
    
    
    DataWTH.DatabaseName = DB_LogPath
    
    g_AtmPRRType = Pcb3dl.DlGetCharRaw("GBLRECPrinterType")
    If g_AtmPRRType <> "T" And g_AtmPRRType <> "I" Then
        g_AtmPRRType = "T"
    End If
    
    Pcb3dl.DlSetCharRaw "MaintHtmlFkeyList", ""
    Pcb3dl.DlSetCharRaw "MaintHtmlFkeyMap", "4095"
  
    nrc = ShowOperScreenMaint("Operator", "OpInService")
    
    Pcb3dl.DlSetCharRaw "GBLInitCasStates", "Y"
    Pcb3dl.DlSetCharRaw "TTU01", "��·����"
          
    S3ETrans.Available = True
    
    'Add for CtrlTest
    DLWaitMonType.Enabled = True
    'Add end
    
    If GetIniS(sKeyIni, "keylist", "DESTYPE", "S") = "H" Then
        g_bIsHardware = True
        Pcb3dl.DlSetCharRaw "GBLEncrypType", "H"
    Else
        g_bIsHardware = False
        Pcb3dl.DlSetCharRaw "GBLEncrypType", "S"
    End If
    
    If GetIniS(sKeyIni, "keylist", "DESMETHOD", "S") = "T" Then
        G_bTrides = True
        Pcb3dl.DlSetCharRaw "GBLEncrypMode", "T"
    Else
        G_bTrides = False
        Pcb3dl.DlSetCharRaw "GBLEncrypMode", "S"
    End If
    
    If GetIniS(sGlobalIni, "Bank_Environment", "PrjLanguage", "E") = "E" Then
        g_sPrjLanguage = "E"
    Else
        g_sPrjLanguage = "C"
    End If
    G_nDevicesToUse = GetRegKeyN(HKEY_LOCAL_MACHINE, keySelfService, "DevicesToUse", 4, 0)

End Sub
Private Sub SDOCdm_OpAtLoadCasStart()
    Dim sIsNewPeriod As String

    SDOCdm.TimeOutSecondsFirst = -1
    
    sIsNewPeriod = Pcb3dl.DlGetCharRaw("GBLIsNewPeriod")
    If sIsNewPeriod = "Y" Then
        GLbIsNewPeriod = True
    Else
        GLbIsNewPeriod = False
    End If
    
    SDOCdm.UserReply = 0
End Sub

Private Sub SDOCdm_OpOperatorInfo(ByVal InfoId As Integer)
    SDOCdm.UserReply = 0
End Sub

Private Sub SDOCdm_OpAfterLiftDown()
    If g_bIsResumeBox = False Then
        currentPage = pageLoadBox30
        TimerAction.Enabled = True
    Else
        SDOCdm.UserReply = 0
    End If
End Sub

Private Sub SDOCdm_OpResetRejectedNotes()
    If g_bIsResumeBox = False Then
        If GLbIsNewPeriod Then
            SDOCdm.ResetRejectCas = True
        Else
            SDOCdm.ResetRejectCas = False
        End If
        currentPage = pageLoadBox20
        TimerAction.Enabled = True
    Else
        SDOCdm.ResetRejectCas = False
        SDOCdm.UserReply = 0
    End If
End Sub

Private Sub SDOCdm_OpSetLoadedNotes(ByVal CasNbrLogical As Integer)
    If g_bIsResumeBox = False Then
        If CasNbrLogical = 1 Or CasNbrLogical = 0 Then
            currentPage = pageLoadBox40
            TimerAction.Enabled = True
        Else
            SDOCdm.UserReply = 0
        End If
    Else
        SDOCdm.UserReply = 0
    End If
End Sub

Private Sub SDOCdm_OpAtLoadCasEnd(ByVal LoadCasnRc As Integer)
    If (LoadCasnRc <> 0) And g_bIsResumeBox = True Then
        currentPage = pageResumeBox50
    ElseIf (LoadCasnRc <> 0) Or (FlushBoxesStatusRetIsPresent = False) Then
        Call PrjLoadFeederFailed
        currentPage = pageLoadBox60
    Else
        If g_bIsResumeBox = True Then
            currentPage = pageResumeBox40
        Else
            currentPage = pageLoadBox50
        End If
    End If
    TimerAction.Enabled = True

End Sub

Private Sub SDOCdm_AtWithdrStart()
    LogInfo "OP Test Dispense Notes Begin!"
    SDOCdm.UserReply = 0
End Sub

Private Sub SDOCdm_BefAuthorisation()
    LogInfo "SDOCdm_BefAuthorisation = 0"
    SDOCdm.UserReply = 0
End Sub

Private Sub SDOCdm_BefDeliver()
    LogInfo "SDOCdm_BefAuthorisation = 166"
    SDOCdm.UserReply = 166
End Sub

Private Sub SDOCdm_GetAuthorisation(ByVal WithdrawalAmount As Long)
    LogInfo "SDOCdm_GetAuthorisation = 0  Withdrawal Amount = " + CStr(WithdrawalAmount)
    SDOCdm.UserReply = 0
End Sub

Private Sub SDOCdm_GetWithdrawalAmount()
    Dim iWithdrawAmount     As Integer
    Dim i                   As Integer
    Dim TotNbrOfBoxused     As Integer
    
    SDOCdm.Currency = "CNY"
    iWithdrawAmount = 0
    TotNbrOfBoxused = 0
    
    SDOCdm.DataCriteria = 1
    For i = 1 To SDOCdm.NbrOfBoxesUsed
        SDOCdm.CasNbrLogical = i
        If Len(SDOCdm.CasPosition) > 0 Then
            If SDOCdm.CasState <= casstate_cdm_low And SDOCdm.CasState >= casstate_cdm_ok And _
                    IsNumeric(Right(SDOCdm.CasPosition, 1)) Then
                    iWithdrawAmount = iWithdrawAmount + SDOCdm.CasDenomination
                    TotNbrOfBoxused = TotNbrOfBoxused + 1
            End If
        End If
    Next i
    
    LogInfo "Test Withdrawal Amount = " + CStr(iWithdrawAmount) + _
                "  TotNbrOfBoxused = " + CStr(TotNbrOfBoxused)
    SDOCdm.WithdrawalAmount = iWithdrawAmount
    
  '   �޸ĵ㳮���֮ǰ������ͬ��ֵ����ʱ����������,�����ȸ����ٸ�ÿ�������������
    For i = 1 To SDOCdm.NbrOfBoxesUsed
        SDOCdm.CasNbrLogical = i
        If Len(SDOCdm.CasPosition) > 0 Then
            If SDOCdm.CasState <= casstate_cdm_low And SDOCdm.CasState >= casstate_cdm_ok And _
                    IsNumeric(Right(SDOCdm.CasPosition, 1)) Then
                    SDOCdm.NotesToDispense = 1
            End If
        End If
    Next i
    SDOCdm.UserReply = 0
End Sub

Private Sub SDOCdm_InformDenomNotPresent(ByVal AbsentDenom As Long)
    LogInfo "SDOCdm_InformDenomNotPresent = 0"
    SDOCdm.UserReply = 0
End Sub
Private Sub SDOCdm_AtWithdrEnd(ByVal WithdrRc As Integer)
    LogInfo "SDOCdm_AtWithdrEnd WithdrRc =" + CStr(WithdrRc)
    
    Select Case WithdrRc
        Case 166
            LogInfo "OP Test Dispense Notes OK!"
            currentPage = pageTestDispenseNoteForEachCas30
        Case Else
            LogInfo "OP Test Dispense Notes Error!"
            currentPage = pageTestDispenseNoteForEachCas40
    End Select
    
    TimerAction.Enabled = True
    
End Sub
Private Sub SDOEdm_AtLoadKeyStart()
    SDOEdm.UserReply = 0
End Sub

Private Sub SDOEdm_GetKey1()
    If SDOEdm.KeyType = 0 Then
        SDOEdm.UserReply = 200
    Else
        SDOEdm.UserReply = 0
    End If
End Sub

Private Sub SDOEdm_GetKey2()
    SDOEdm.UserReply = 100
End Sub

Private Sub SDOEdm_AtLoadKeyEnd(ByVal LoadKeyRc As Integer)
    Dim sMasterKeyName     As String
    
    If LoadKeyRc = 100 Then
        currentPage = pageOpKeyInput40
        LogInfo "DoLoadKey OK in Operator. MasterKeyName: " + sMasterKeyName
    Else
        If LoadKeyRc = 200 Then
            nrc = SDOEdm.PuOpen
            currentPage = pageOpKeyInput35
            LogInfo "DoLoadKey OK in Operator. MasterKeyName: " + sMasterKeyName
        Else
            currentPage = pageOpKeyInput50
            LogError "DoLoadKey in Operator return failed, " + CStr(LoadKeyRc)
        End If
    End If
    
    TimerAction.Enabled = True
End Sub
'==========================================================================================
'�������� :��ӡ�ӳ���Ϣ
'������� ����
'������� ����
'����ֵ   ����
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Sub PrjLoadFeederOk()
    Dim theTime         As String
    Dim sDispNumBuf     As String
    Dim i               As Byte
    Dim PrjString       As String
    Dim PrjCHNString    As String

    theTime = Format(Now(), "YY/MM/DD HH:MM")
   
    PrjString = JourLineSeprator + "             LOAD FEEDERS" + vbCrLf + _
                "    " + theTime + "      ATM:" + AtmCode + vbCrLf

    PrjCHNString = JourLineSeprator + "             �ӡ���" + vbCrLf + _
                    "    " + theTime + "     ATM �ţ�" + AtmCode + vbCrLf
                
    PrjString = PrjString + "B"
    
    PrjCHNString = PrjCHNString + "����� "
    
    For i = 1 To nNumberOfCassettes
        PrjString = PrjString + Format(CStr(i) + "#", "@@@@@@@")
        PrjCHNString = PrjCHNString + Format(CStr(i) + "#", "@@@@@@")
    Next i
    PrjString = PrjString + vbCrLf
    PrjCHNString = PrjCHNString + vbCrLf
                
    PrjString = PrjString + "C"
    PrjCHNString = PrjCHNString + "�� �� "
    For i = 1 To nNumberOfCassettes
        PrjString = PrjString + Format(WthCassette(i).BoxCurry, "@@@@@@@")
        PrjCHNString = PrjCHNString + Format(WthCassette(i).BoxCurry, "@@@@@@")
    Next i
    PrjString = PrjString + vbCrLf
    PrjCHNString = PrjCHNString + vbCrLf
    
    PrjString = PrjString + "D"
    PrjCHNString = PrjCHNString + "�� �� "
    For i = 1 To nNumberOfCassettes
        sDispNumBuf = Format(WthCassette(i).BoxDenom, "000")
        PrjString = PrjString + Format(sDispNumBuf, "@@@@@@@")
        PrjCHNString = PrjCHNString + Format(sDispNumBuf, "@@@@@@")
    Next i
    PrjString = PrjString + vbCrLf
    PrjCHNString = PrjCHNString + vbCrLf
    
    PrjString = PrjString + "L"
    PrjCHNString = PrjCHNString + "�� ��"
    For i = 1 To nNumberOfCassettes
        sDispNumBuf = Format(WthCassette(i).BoxLeftNum, "0000")
        PrjString = PrjString + Format(sDispNumBuf, "@@@@@@@")
        PrjCHNString = PrjCHNString + Format(sDispNumBuf, "@@@@@@")
    Next i
    PrjString = PrjString + vbCrLf
    PrjCHNString = PrjCHNString + vbCrLf
    
    PrjString = PrjString + "A"
    PrjCHNString = PrjCHNString + "�� ��"
    For i = 1 To nNumberOfCassettes
        sDispNumBuf = Format(WthCassette(i).BoxDenom * WthCassette(i).BoxLeftNum, "000000")
        PrjString = PrjString + Format(sDispNumBuf, "@@@@@@@")
        PrjCHNString = PrjCHNString + Format(sDispNumBuf, "@@@@@@")
    Next i
    PrjString = PrjString + vbCrLf
    PrjCHNString = PrjCHNString + vbCrLf
    
    LogInfo (PrjString)
    
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    
    If g_sPrjLanguage = "E" Then
        g_sPrrRawData = g_sPrrRawData + vbCrLf + PrjString
    Else
        g_sPrrRawData = g_sPrrRawData + vbCrLf + PrjCHNString
    End If
End Sub
'==========================================================================================
'�������� :��ӡ�ӳ�ʧ����Ϣ
'������� ����
'������� ����
'����ֵ   ����
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Sub PrjLoadFeederFailed()
    Dim theTime       As String
    Dim PrjString     As String
    Dim PrjCHNString  As String
 
    theTime = Format(Now(), "YY/MM/DD HH:MM")
                
    PrjString = JourLineSeprator + "         LOAD FEEDERS ERROR            " + vbCrLf + _
                 " " + theTime + " ATM:" + AtmCode + vbCrLf

    PrjCHNString = JourLineSeprator + "              �ӡ�����������" + vbCrLf + _
                 " " + theTime + " ATM�ţ�" + AtmCode + vbCrLf
    
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
End Sub
'==========================================================================================
'�������� :��ӡ�������Ա��Ϣ
'������� ����
'������� ����
'����ֵ   ����
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Sub PrjStartOperator()
    Dim theTime      As String
    Dim PrjString    As String
    Dim PrjCHNString As String
    
    theTime = Format(Now(), "YY/MM/DD HH:MM")
                
    PrjString = JourLineSeprator + "     START OPERATOR INTERVENTION       " + vbCrLf + _
                "    " + theTime + " ATM:" + AtmCode + vbCrLf

    PrjCHNString = JourLineSeprator + "     ���������롡�١�����Ա��״��̬" + vbCrLf + _
                "    " + theTime + " ATM�ţ�" + AtmCode + vbCrLf
    
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
End Sub
'==========================================================================================
'�������� :��ӡ���볬������Ա��Ϣ
'������� ����
'������� ����
'����ֵ   ����
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Sub PrjStartSuperOperator()
    Dim theTime      As String
    Dim PrjString    As String
    Dim PrjCHNString As String
    
    theTime = Format(Now(), "YY/MM/DD HH:MM")
                
    PrjString = JourLineSeprator + "     START SUPER OPERATOR INTERVENTION       " + vbCrLf + _
                "    " + theTime + " ATM:" + AtmCode + vbCrLf

    PrjCHNString = JourLineSeprator + "     ���������롡�� �� �١�����Ա��״��̬" + vbCrLf + _
                "    " + theTime + " ATM�ţ�" + AtmCode + vbCrLf
    
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
End Sub
'==========================================================================================
'�������� :��ӡ�˳�����Ա��Ϣ
'������� ����
'������� ����
'����ֵ   ����
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Sub PrjEndOperator()
    Dim theTime      As String
    Dim PrjString    As String
    Dim PrjCHNString As String
    
    theTime = Format(Now(), "YY/MM/DD HH:MM")
                
    PrjString = JourLineSeprator + "      END OPERATOR INTERVENTION        " + vbCrLf + _
                "    " + theTime + " ATM:" + AtmCode + vbCrLf

    PrjCHNString = JourLineSeprator + "     �����ˡ������١�����Ա��״��̬" + vbCrLf + _
                "    " + theTime + " ATM�ţ�" + AtmCode + vbCrLf
                
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
End Sub
'==========================================================================================
'�������� :��ӡ��������Ϣ
'������� ����
'������� ����
'����ֵ   ����
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'�޸Ĵ�ӡ��������Ϣ�������н��������뿪���ںϲ������
'==========================================================================================
Private Sub PrjOpenPeriod()
    Dim theTime      As String
    Dim PrjString    As String
    Dim PrjCHNString As String
    
    theTime = Format(Now(), "YY/MM/DD HH:MM")
                
    PrjString = "      CLEAR TOTAL SUCCEED        " + vbCrLf
                
    PrjCHNString = "         �� �� �� �� �� ��" + vbCrLf
                
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
End Sub

'==========================================================================================
'�������� :��ӡͳ����Ϣ
'������� ����ʼ��Ϣ
'������� ����
'����ֵ   ����
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Sub PrjTotal(pTitle As String)
    Dim theTime                 As String
    Dim PrjString               As String
    Dim PrjCHNString            As String
    Dim i                       As Byte
    Dim NorBinTotLeftAmount     As Long
    Dim sDispNumBuf             As String
    Dim pTitleCHN               As String
    
    theTime = Format(Now(), "YY/MM/DD HH:MM")
    
    If pTitle = "TOTAL PRINTING" Then
        pTitleCHN = "��ӡ��ͳ���ơ�ֵ"
    Else
        pTitle = "CHECK TOTAL"
        pTitleCHN = "��  ӡ  ��  ��  ��"
    End If
        
    PrjString = JourLineSeprator + _
                "           " + pTitle + vbCrLf + _
                 "    " + theTime + "      ATM:" + AtmCode + vbCrLf + _
                "RETAINED CARDS NUMBER:" + CStr(Pcb3dl.DlGetInt("TotCapCardNum")) + vbCrLf + _
                "WTH    Num:" + Format(Pcb3dl.DlGetInt("TotWithdrawNum") + _
                                Pcb3dl.DlGetInt("IcbcTotExtraWthNum"), "0000") + _
                    "  AMOUNT:" + Format(Pcb3dl.DlGetDouble("TotWithdrawAmount") + _
                                        Pcb3dl.DlGetDouble("IcbcTotExtraWthAmount"), "Standard") + vbCrLf + _
                "WTHREV Num:" + Format(Pcb3dl.DlGetInt("TotWthReversalNum") + _
                                Pcb3dl.DlGetInt("IcbcTotExtraWthRevNum"), "0000") + _
                    "  AMOUNT:" + Format(Pcb3dl.DlGetDouble("TotWthReversalAmount") + _
                                        Pcb3dl.DlGetDouble("IcbcTotExtraWthRevAmount"), "Standard") + vbCrLf + _
                "TSFOUT Num:" + Format(Pcb3dl.DlGetInt("TotTfrOutNum"), "0000") + _
                    "  AMOUNT:" + Format(Pcb3dl.DlGetDouble("TotTfrOutAmount"), "Standard") + vbCrLf + _
                "INQ    Num:" + Format(Pcb3dl.DlGetInt("TotInquiryNum"), "0000") + vbCrLf + _
                "PINCHG Num:" + Format(Pcb3dl.DlGetInt("TotPinChangeNum"), "0000") + vbCrLf
    PrjCHNString = JourLineSeprator + _
                "           " + pTitleCHN + vbCrLf + _
                "    " + theTime + "     ATM �ţ�" + AtmCode + vbCrLf + _
                "�� �� �� ��  ��" + CStr(Pcb3dl.DlGetInt("TotCapCardNum")) + vbCrLf + _
                "ȡ �� �� �� ��:" + Format(Pcb3dl.DlGetInt("TotWithdrawNum") + _
                                Pcb3dl.DlGetInt("IcbcTotExtraWthNum"), "0000") + vbCrLf + _
                "ȡ �� �� �� ��:" + Format(Pcb3dl.DlGetDouble("TotWithdrawAmount") + _
                                        Pcb3dl.DlGetDouble("IcbcTotExtraWthAmount"), "Standard") + vbCrLf + _
                "�� �� �� �� ��:" + Format(Pcb3dl.DlGetInt("TotWthReversalNum") + _
                                Pcb3dl.DlGetInt("IcbcTotExtraWthRevNum"), "0000") + vbCrLf + _
                "�� �� �� �� ��:" + Format(Pcb3dl.DlGetDouble("TotWthReversalAmount") + _
                                        Pcb3dl.DlGetDouble("IcbcTotExtraWthRevAmount"), "Standard") + vbCrLf + _
                "ת �� �� �� ��:" + Format(Pcb3dl.DlGetInt("TotTfrOutNum"), "0000") + vbCrLf + _
                "ת �� �� �� ��:" + Format(Pcb3dl.DlGetDouble("TotTfrOutAmount"), "Standard") + vbCrLf + _
                "�� ѯ �� �� ��:" + Format(Pcb3dl.DlGetInt("TotInquiryNum"), "0000") + vbCrLf + _
                "�� �� �� �� ��:" + Format(Pcb3dl.DlGetInt("TotPinChangeNum"), "0000") + vbCrLf
    
    PrjString = PrjString + vbCrLf + vbCrLf + "             Cassettes total summary" + vbCrLf
    PrjCHNString = PrjCHNString + vbCrLf + vbCrLf + "           �����䡡ͳ����" + vbCrLf
                
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    If g_sPrjLanguage = "E" Then
        g_sPrrRawData = PrjString
    Else
        g_sPrrRawData = PrjCHNString
    End If
    
    PrjString = "B"
    PrjCHNString = "����� "
    For i = 1 To nNumberOfCassettes
        PrjString = PrjString + Format(CStr(i) + "#", "@@@@@@@")
        PrjCHNString = PrjCHNString + Format(CStr(i) + "#", "@@@@@@")
    Next i
    
    PrjString = PrjString + vbCrLf
    PrjCHNString = PrjCHNString + vbCrLf

    NorBinTotLeftAmount = 0
    
    Call FlushBoxesStatusRetIsPresent
    
    For i = 1 To nNumberOfCassettes
        If (WthCassette(i).CasLogicalID <> 0 And WthCassette(i).BoxCurry <> "XXX") Then
            NorBinTotLeftAmount = NorBinTotLeftAmount + _
                WthCassette(i).BoxDenom * WthCassette(i).BoxLeftNum
        End If
    Next i
    
    PrjString = PrjString + "C"
    PrjCHNString = PrjCHNString + "�� �� "
    
    For i = 1 To nNumberOfCassettes
        PrjString = PrjString + Format(WthCassette(i).BoxCurry, "@@@@@@@")
        PrjCHNString = PrjCHNString + Format(WthCassette(i).BoxCurry, "@@@@@@")
    Next i
    PrjString = PrjString + vbCrLf
    PrjCHNString = PrjCHNString + vbCrLf
    
    PrjString = PrjString + "D"
    PrjCHNString = PrjCHNString + "�� �� "
    For i = 1 To nNumberOfCassettes
        sDispNumBuf = Format(WthCassette(i).BoxDenom, "@@@")
        PrjString = PrjString + Format(sDispNumBuf, "@@@@@@@")
        PrjCHNString = PrjCHNString + Format(sDispNumBuf, "@@@@@@")
    Next i
    PrjString = PrjString + vbCrLf
    PrjCHNString = PrjCHNString + vbCrLf
    
    PrjString = PrjString + "O"
    PrjCHNString = PrjCHNString + "������"
    For i = 1 To nNumberOfCassettes
        sDispNumBuf = Format(WthCassette(i).BoxDisp, "@@@@")
        PrjString = PrjString + Format(sDispNumBuf, "@@@@@@")
        PrjCHNString = PrjCHNString + Format(sDispNumBuf, "@@@@@@")
    Next i
    PrjString = PrjString + vbCrLf
    PrjCHNString = PrjCHNString + vbCrLf
    
    PrjString = PrjString + "R"
    PrjCHNString = PrjCHNString + "�ϳ���"
    For i = 1 To nNumberOfCassettes
        sDispNumBuf = Format(WthCassette(i).PurgedNotes, "@@@@")
        PrjString = PrjString + Format(sDispNumBuf, "@@@@@@")
        PrjCHNString = PrjCHNString + Format(sDispNumBuf, "@@@@@@")
    Next i
    PrjString = PrjString + vbCrLf
    PrjCHNString = PrjCHNString + vbCrLf
    
    PrjString = PrjString + "L"
    PrjCHNString = PrjCHNString + "ʣ����"
    For i = 1 To nNumberOfCassettes
        sDispNumBuf = Format(WthCassette(i).BoxLeftNum, "@@@@")
        PrjString = PrjString + Format(sDispNumBuf, "@@@@@@")
        PrjCHNString = PrjCHNString + Format(sDispNumBuf, "@@@@@@")
    Next i
    PrjString = PrjString + vbCrLf
    PrjCHNString = PrjCHNString + vbCrLf
    
    PrjString = PrjString + "A"
    PrjCHNString = PrjCHNString + "ʣ�����"
    For i = 1 To nNumberOfCassettes
        sDispNumBuf = Format(WthCassette(i).BoxDenom * WthCassette(i).BoxLeftNum, "@@@@@@")
        PrjString = PrjString + Format(sDispNumBuf, "@@@@@@@")
        PrjCHNString = PrjCHNString + Format(sDispNumBuf, "@@@@@@")
    Next i
    PrjString = PrjString + vbCrLf
    PrjCHNString = PrjCHNString + vbCrLf
    
    PrjString = PrjString + "I"
    PrjCHNString = PrjCHNString + "װ������"
    For i = 1 To nNumberOfCassettes
        sDispNumBuf = Format(WthCassette(i).BoxInitial, "@@@@")
        PrjString = PrjString + Format(sDispNumBuf, "@@@@@@@")
        PrjCHNString = PrjCHNString + Format(sDispNumBuf, "@@@@@@")
    Next i
    PrjString = PrjString + vbCrLf
    PrjCHNString = PrjCHNString + vbCrLf
    
    PrjString = PrjString + "T"
    PrjCHNString = PrjCHNString + "װ�����"
    For i = 1 To nNumberOfCassettes
        sDispNumBuf = Format(WthCassette(i).BoxDenom * WthCassette(i).BoxInitial, "@@@@@@")
        PrjString = PrjString + Format(sDispNumBuf, "@@@@@@@")
        PrjCHNString = PrjCHNString + Format(sDispNumBuf, "@@@@@@")
    Next i
    
    PrjString = PrjString + vbCrLf
    PrjCHNString = PrjCHNString + vbCrLf
    
    PrjString = PrjString + "Total Left Amount(A)= " + Format(NorBinTotLeftAmount, "000000") + vbCrLf
    PrjCHNString = PrjCHNString + "ʣ�����= " + Format(NorBinTotLeftAmount, "000000") + vbCrLf
    
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    If g_sPrjLanguage = "E" Then
        g_sPrrRawData = g_sPrrRawData + vbCrLf + PrjString
    Else
        g_sPrrRawData = g_sPrrRawData + vbCrLf + PrjCHNString
    End If
End Sub

'==========================================================================================
'�������� :��ӡ������Ϣ
'������� ����
'������� ����
'����ֵ   ����
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Sub PrjDeviceStatus()
    Dim theTime      As String
    Dim PrjString    As String
    Dim PrjCHNString As String
    Dim i            As Integer
    Dim nNumOfBoxes  As Integer
    
    theTime = Format(Now(), "YY/MM/DD HH:MM")
    
    PrjString = "---------------------------------------" + vbCrLf + _
                "         PERIPHERAL STATUS             " + vbCrLf + _
                "Time :" + theTime + vbCrLf + "      ATM:" + AtmCode + vbCrLf + _
                "    PRJ: " + TranslateDeviceState("PRJ", False) + vbCrLf + _
                "    PRR: " + TranslateDeviceState("PRR", False) + vbCrLf + _
                "    IDC: " + TranslateDeviceState("IDC", False) + vbCrLf + _
                "    EDM: " + TranslateDeviceState("EDM", False) + vbCrLf + _
                "    CDM: " + TranslateDeviceState("CDM", False) + vbCrLf
    PrjCHNString = "---------------------------------------" + vbCrLf + _
                "         �����⡡�衡״��̬            " + vbCrLf + _
                "ʱ�䣺" + theTime + vbCrLf + "      ATM�ţ�" + AtmCode + vbCrLf + _
                "    ��ˮ��ӡ���� " + TranslateDeviceState("PRJ", True) + vbCrLf + _
                "    ƾ����ӡ���� " + TranslateDeviceState("PRR", True) + vbCrLf + _
                "    �ſ���д����" + TranslateDeviceState("IDC", True) + vbCrLf + _
                "    ���ܡ�ģ�飺" + TranslateDeviceState("EDM", True) + vbCrLf + _
                "    ������ģ�飺" + TranslateDeviceState("CDM", True) + vbCrLf

    Call FlushBoxesStatusRetIsPresent
    nNumOfBoxes = SDOCdm.NbrOfBoxesUsed
    
    For i = 1 To nNumberOfCassettes
        PrjString = PrjString + "Cas" + CStr(i) + " State: " + _
                WthCassette(i).BoxState + vbCrLf
        PrjCHNString = PrjCHNString + "����" + CStr(i) + " ״̬��" + _
                WthCassette(i).BoxStateCHN + vbCrLf
    Next
   
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
End Sub

'==========================================================================================
'�������� :��ӡ������Ϣ
'������� ����
'������� ����
'����ֵ   ����
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Sub PrjBoxStatus()
    Dim theTime        As String
    Dim PrjString      As String
    Dim PrjCHNString   As String
     
    theTime = Format(Now(), "YY/MM/DD HH:MM")
    
    PrjString = JourLineSeprator + "         SHOWING FEEDERS STATUS" + vbCrLf + _
                 "    " + theTime + "      ATM:" + AtmCode + vbCrLf
                
    PrjCHNString = JourLineSeprator + "         �����䡡״��̬" + vbCrLf + _
                 "    " + theTime + "      ATM�ţ�" + AtmCode + vbCrLf
                
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
End Sub
Private Sub ClearTotal()
    
    nrc = Pcb3dl.DlSetCharRaw("GBLIsNewPeriod", "Y")
    nrc = Pcb3dl.DlSetLong("TotWithdrawNum", 0)
    nrc = Pcb3dl.DlSetLong("TotTfrOutNum", 0)
    nrc = Pcb3dl.DlSetLong("TotInquiryNum", 0)
    nrc = Pcb3dl.DlSetLong("TotPinChangeNum", 0)
    nrc = Pcb3dl.DlSetLong("TotCapCardNum", 0)
    nrc = Pcb3dl.DlSetLong("TotJournalNum", 0)
    
    nrc = Pcb3dl.DlSetDouble("TotWithdrawAmount", 0)
    nrc = Pcb3dl.DlSetDouble("TotTfrOutAmount", 0)
    
'Add for ICBC
    nrc = Pcb3dl.DlSetLong("IcbcTotExtraWthNum", 0)
    nrc = Pcb3dl.DlSetDouble("IcbcTotExtraWthAmount", 0)
    nrc = Pcb3dl.DlSetLong("IcbcTotExtraWthRevNum", 0)
    nrc = Pcb3dl.DlSetDouble("IcbcTotExtraWthRevAmount", 0)
    nrc = Pcb3dl.DlSetLong("TotWthReversalNum", 0)
    nrc = Pcb3dl.DlSetDouble("TotWthReversalAmount", 0)
'Add end
    DataWTH.RecordSource = "Select * From CWDLOG "
    DataWTH.Refresh
    DataWTH.Database.Execute "Delete * from CWDLOG"
    DataWTH.Recordset.Requery
End Sub
Private Sub SDOPrr_AtAddPage()
    SDOPrr.PageText = g_sPrrRawData
    SDOPrr.UserReply = 0
End Sub

Private Sub SDOPrr_AtPresented()
    SDOPrr.UserReply = 0
End Sub

Private Sub SDOPrr_AtPresentTimeout()
    currentPage = pagePrrPrintTOT35
    TimerAction.Enabled = True
End Sub

Private Sub SDOPrr_AtPrintFormEnd(ByVal Rc As Integer)
    Dim PrjString    As String
    Dim PrjCHNString As String

        If Rc <> 0 And Rc <> 91 Then
        
            PrjString = "PRR out of service."
            PrjCHNString = "ƾ����ӡ������"
            
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            currentPage = pagePrrPrintTOT50
        Else
                PrrLeftPrintPageNumber = PrrLeftPrintPageNumber - 1
                If PrrLeftPrintPageNumber > 0 Then
                currentPage = pagePrrPrintTOT15
                Else
                     If IsPrintAmonalyTrans And gSelectOprCommand = "04" Then   '�ӳ�ǰ��ӡ�쳣ȡ������
                         IsPrintAmonalyTrans = False
                         currentPage = pageLoadBoxWarning
                     Else               '��ӡ��������,ͳ������������ִ���쳣ȡ��
                         currentPage = pageFunChoice
                     End If
                End If
        End If
    TimerAction.Enabled = True
End Sub

Private Sub SDOPrr_AtPrintFormStart()
    currentPage = pagePrrPrintTOT30
    TimerAction.Enabled = True
    
'    SDOPrr.UserReply = 0
End Sub

Private Sub SDOPrr_AtPrintRawEnd(ByVal Rc As Integer)
    Dim PrjString As String
    Dim PrjCHNString As String, sDEVStatus As String
    If g_bIsPrrResetTest Then
        If Rc <> 0 And Rc <> 91 Then
            LogError "PRR out of service"
            PrintJournal SDOPrj, "PRR out of service", "ƾ����ӡ������", g_sPrjLanguage

        Else
            PrintJournal SDOPrj, "PRR OK", "ƾ����ӡ����ӡ���Գɹ�", g_sPrjLanguage
        End If
        sDEVStatus = TranslateDeviceState("PRR", True)
        nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt2", sDEVStatus)
        currentPage = pageResetATM40
    Else

        If Rc <> 0 And Rc <> 91 Then
        
            PrjString = "PRR out of service."
            PrjCHNString = "ƾ����ӡ������"
            
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            currentPage = pagePrrPrintTOT50
        Else
            currentPage = pagePrrPrintTOT40
        End If
    End If
    TimerAction.Enabled = True
End Sub

Private Sub SDOPrr_AtPrintRawStart()
    currentPage = pagePrrPrintTOT20
    TimerAction.Enabled = True
End Sub

Private Sub SDOPrr_BeforePresent()
     currentPage = pagePrrPrintTOT30
     TimerAction.Enabled = True
End Sub

Private Sub start_Click()
    Call CheckupFileExist

    currentPage = pageOpPinInput
    TimerAction.Enabled = True
End Sub
Private Sub S3ETrans_StartTransaction(ByVal Action As Long)
    Dim sSubStData As String
    
    Call CheckupFileExist
    Call PrjStartOperator
    
    Call SendExceptionMessage(S3ELineOut, Pcb3dl, "23")
    g_bIsPrrResetTest = False
    nrc = ShowScreenSync(Browser, "Operator", "OpInMaintain", sSubStData)
    AtmCode = Pcb3dl.DlGetCharRaw("GBLAtmCode")
    nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "")
'Add for Opteva
    If Browser.HasSecondMonitor = 0 Then
        Browser.WindowStyle = WINDOWED
        BrowserMaint.WindowStyle = TOP_FULL_SCREEN
    End If
'Add end
    SDOFep.DoServiceClose
    
    currentPage = pageOpPinInput
    TimerAction.Enabled = True
End Sub
Private Sub S3ETrans_QuitTransaction()
    currentPage = pageQuit
    TimerAction.Interval = 1000
    TimerAction.Enabled = True
End Sub
'==========================================================================================
'�������� :��ӡ�̿���Ϣ
'������� ����
'������� ����
'����ֵ   ����
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Function PrintRetainCard() As Boolean
    Dim theTime      As String
    Dim PrjString    As String
    Dim PrjCHNString As String
    Dim obj          As New FileSystemObject
    Dim LogStream    As TextStream
    Dim sLogRec      As String
        
    theTime = Format(Now(), "YY/MM/DD HH:MM")
                    
    PrjString = JourLineSeprator + "     Print retained card file           " + _
                 " " + theTime + "  ATM:" + AtmCode + vbCrLf
                 
    PrjCHNString = JourLineSeprator + "     ��ӡ�̿���־�ļ�" + _
                 " " + theTime + "  ATM�ţ�" + AtmCode + vbCrLf
    
    If obj.FileExists(CardRetainFile) Then
       
        PrintRetainCard = True
        PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
        
        Set LogStream = obj.OpenTextFile(CardRetainFile, ForReading)
        Do While Not LogStream.AtEndOfStream
            sLogRec = LogStream.ReadLine
            SDOPrj.DoPrint sLogRec + vbCrLf
            SaveCNJournal sLogRec + vbCrLf
        Loop
        LogStream.Close
        PrjString = "     Total retained card Number:[" + CStr(Pcb3dl.DlGetInt("TotCapCardNum")) + "]" + vbCrLf + _
            "     Retain card file printed end.      " + vbCrLf
        PrjCHNString = "     �̡������ܡ�����[" + CStr(Pcb3dl.DlGetInt("TotCapCardNum")) + "]" + vbCrLf + _
            "     �̿���־�ļ���ӡ���" + vbCrLf
        PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    Else
        PrintRetainCard = False
        PrjString = "No Card was captured" + vbCrLf
        PrjCHNString = "������û���̿�" + vbCrLf
        PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage

    End If
End Function
'==========================================================================================
'�������� :��ӡ�ػ���Ϣ
'������� ����
'������� ����
'����ֵ   ����
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Sub PrjShutdownSystem()
    Dim theTime      As String
    Dim PrjString    As String
    Dim PrjCHNString As String
   
    theTime = Format(Now(), "YY/MM/DD HH:MM")
                
    PrjString = JourLineSeprator + "           SHUT DOWN SYSTEM            " + vbCrLf + _
                " " + theTime + " ATM:" + AtmCode + vbCrLf

    PrjCHNString = JourLineSeprator + "           ��ϵ��ͳ���ء���" + vbCrLf + _
                " " + theTime + " ATM�ţ�" + AtmCode + vbCrLf
                
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                
End Sub

Private Sub TimerAction_Timer()
    Dim PrjString           As String
    Dim PrjCHNString        As String
'    Dim hEvent              As Long
    Dim sTTUChoice          As String
    Dim i                   As Integer
    Dim sOperPin            As String
    Dim sHtmlInput          As String
    Dim NorBinTotLeftAmount As Long
    Dim strResult           As String
    Dim vDateNow            As Variant
    Dim fso                 As New FileSystemObject
    Dim sTemp               As String
    Dim sKeyMap             As String
    Dim sHtmlInput1         As String
    Dim sHtmlInput2         As String
    Dim PreMasterKeyA       As String
    Dim PreMasterKeyB       As String
    Dim sDisplayStr         As String
    Dim nInputCount         As Integer
    Dim nDenomBox           As Integer
    Dim sLoadNum            As String
    Dim sCurTime            As String
    Dim sFirstNewPin        As String
    Dim sSecondNewPin       As String
    Dim ReturnValue         As String
'    Dim hS3EStartStopEvent  As Long
    Dim sValue              As String
    Dim bDisableTimer       As Boolean
    Dim iDisplayNum         As Integer
    Dim nRc1                As Integer
    
    TimerAction.Enabled = False
    
    Select Case currentPage
        Case pageFirstPage
            nrc = ShowOperScreenMaint("Operator", "OpFirstPage")
            currentPage = pageFunChoice
            
        Case pageOpPinInput
            sOperPin = GetIniS(IniPath + "Global.ini", "Operator", "OperPassWord", "000000")
            nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
            nrc = ShowOperScreenMaint("Operator", "OpPinInput")
            If nrc = 0 Then
                If BrowserMaint.SubStData = "@ok" Then
                    sHtmlInput = Pcb3dl.DlGetCharRaw("HtmlInput1")
                    If sHtmlInput = sOperPin Then
                        Call PrjStartOperator
                        Call CheckupFileExist
                        currentPage = pageFirstPage
                    Else
                        currentPage = pageOpPassWrong
                    End If
                Else
                    currentPage = pageReturnOk
                End If
            Else
                currentPage = pageReturnOk
            End If
                        
        Case pageOpPassWrong
            nrc = ShowOperScreenMaint("Operator", "OpPassWrong")
            currentPage = pageReturnOk
        
        Case pageFunChoice
            SuperAdminBegin = False
            IsPrintAmonalyTrans = False
            nrc = Pcb3dl.DlSetCharRaw("TTUChoice", "")
            
            nrc = ShowOperScreenMaint("Operator", "OpFunChoice")
            
            sTTUChoice = Pcb3dl.DlGetCharRaw("TTUChoice")
            If Len(sTTUChoice) = 1 Then
                sTTUChoice = "0" + sTTUChoice
            End If

            currentPage = GetOprFunctionPage(sTTUChoice)
            
        Case pageCmdList10
            nrc = ShowOperScreenMaint("Operator", "OpCmdList10")
            currentPage = pageFunChoice
            
        Case pagePrintTotal10
            nrc = ShowOperScreenMaint("Operator", "OpPrintTotal10")
            If BrowserMaint.SubStData = "@ok" Then
                currentPage = pagePrintTotal20
            Else
                currentPage = pageFunChoice
            End If
           
        Case pagePrintTotal20
            nrc = ShowOperScreenMaint("Operator", "OpPrintTotal20")
            Call PrjTotal("TOTAL PRINTING")
            currentPage = pagePrintTotal30
        
        Case pagePrintTotal30   'ѡ���Ƿ��ӡͳ��ƾ��
            nrc = Pcb3dl.DlSetCharRaw("TTUChoice", "")
            nrc = ShowOperScreenMaint("Operator", "OpPrintTotal30")
            
            If BrowserMaint.SubStData = "@ok" Then
                sTemp = Trim(Pcb3dl.DlGetCharRaw("TTUChoice"))
                
                If sTemp = "1" Then
                    If g_sPrjLanguage = "E" Then

                        currentPage = pagePrrPrintTOT10
                    Else
                        If SDOPrr.Available = True Then
                        
    '                   ��������ƾ������ͨ��DoPrintRaw���ķ���ӡ�����ͨ��DoPrintForm�ķ�����ʵ��
                            Call CalPageNum
                            Call PrrTotal
                        Else
                            currentPage = pagePrrPrintTOT50
                        End If
                    End If
                Else
                    currentPage = pageFunChoice
                End If
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageOpenPeriod10
            nrc = ShowOperScreenMaint("Operator", "OpOpenPeriod10")
            If BrowserMaint.SubStData = "@ok" Then
                currentPage = pageOpenPeriod20
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageOpenPeriod20
            nrc = ShowOperScreenMaint("Operator", "OpOpenPeriod20")
            Call ClearTotal
            Call PrjOpenPeriod
            nrc = Pcb3dl.DlSetCharRaw("GBLPeriodStatus", "O")
            sGLtheTime = Format(Now(), "YYYY/MM/DD HH:MM:SS")
            nrc = Pcb3dl.DlSetCharRaw("TotPeriodOpenTime", sGLtheTime)
            currentPage = pageOpenPeriod30
        
        Case pageOpenPeriod30
            nrc = ShowOperScreenMaint("Operator", "OpOpenPeriod30")
            currentPage = pageFunChoice
        
        Case pageWarnPNC
            nrc = ShowOperScreenMaint("Operator", "OpWarnPNC")
            currentPage = pageFunChoice
        '�������������ʱֻ����TTI�����������رջ�����ڣ��ȵ��ӳ�����ʱһ����
        Case pageClosePeriod10
            nrc = ShowOperScreenMaint("Operator", "OpClosePeriod10")
            
'            If Pcb3dl.DlGetCharRaw("CWDCrimePossible") = "O" Then
'                nrc = Pcb3dl.DlSetCharRaw("CWDCrimePossible", "N")
'                Pcb3dl.DlSetCharRaw "GBLDoRecovery", "N"
'            End If

            If BrowserMaint.SubStData = "@ok" Then
                currentPage = pageClosePeriod20
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageClosePeriod20
            nrc = ShowOperScreenMaint("Operator", "OpClosePeriod20")
'            Call PrjTotal("CLOSE PERIOD")
'
'            nrc = Pcb3dl.DlSetCharRaw("GBLPeriodStatus", "C")
'            sGLtheTime = Format(Now(), "YYYY/MM/DD HH:MM:SS")
'            nrc = Pcb3dl.DlSetCharRaw("TotPeriodCloseTime", sGLtheTime)
            Call CommunicationSubFunction("TTI", "AAP")
            
            'AnRchive log files
           
'            hEvent = OpenEvent(EVENT_MODIFY_STATE, 0, "S3EDoArchive")
'            If hEvent <> 0 Then
'                SetEvent hEvent
'                CloseHandle hEvent
'            End If
            currentPage = pageClosePeriod40
        '��������ҳ�汻�������ӳ�������ѯ�ʿͻ��Ƿ��ӡͳ��ƾ��
        Case pageClosePeriod30

            nrc = Pcb3dl.DlSetCharRaw("TTUChoice", "")

            nrc = ShowOperScreenMaint("Operator", "OpClosePeriod30")
            
            If BrowserMaint.SubStData = "@ok" Then
                sTemp = Trim(Pcb3dl.DlGetCharRaw("TTUChoice"))
                
                If sTemp = "1" Then
                    If g_sPrjLanguage = "E" Then
                        currentPage = pagePrrPrintTOT10
                    Else
                        If SDOPrr.Available = True Then
    '                   ��������ƾ������ͨ��DoPrintRaw���ķ���ӡ�����ͨ��DoPrintForm�ķ�����ʵ��
                        Call CalPageNum
                        Call PrrTotal
                        Else
                            currentPage = pagePrrPrintTOT50
                        End If
                    End If
                Else
                    currentPage = pagePrrPrintTOT40
                End If
            Else
                currentPage = pagePrrPrintTOT40
            End If
        '������׽��
        Case pageClosePeriod40
            nrc = ShowOperScreenMaint("Operator", "OpClosePeriod40")
            currentPage = pageFunChoice
            
         Case pagePrrPrintTOT10
            nrc = ShowOperScreenMaint("Operator", "OpPrrPrintTOT10")
            If BrowserMaint.SubStData = "@ok" Then
                If SDOPrr.Available = True Then
                    nrc = SDOPrr.DoPrintRaw()
                    If nrc <> 0 Then
                        currentPage = pagePrrPrintTOT40
                    Else
                        Exit Sub
                    End If
                Else
                    currentPage = pagePrrPrintTOT50
                End If
            Else
                currentPage = pageFunChoice
            End If
        
        Case pagePrrPrintTOT15
             If SDOPrr.Available = True Then
                 If PrrLeftPrintPageNumber = 0 Then
                     If IsPrintAmonalyTrans And gSelectOprCommand = "04" Then   '�ӳ�ǰ��ӡ�쳣ȡ������
                         IsPrintAmonalyTrans = False
                         currentPage = pageLoadBoxWarning
                     Else               '��ӡ��������,ͳ������������ִ���쳣ȡ��
                         currentPage = pageFunChoice
                     End If
                 Else
                     Call PrrTotal
                 End If
             Else
                currentPage = pagePrrPrintTOT50
             End If
                
        Case pagePrrPrintTOT20    '���ڴ�ӡ
            nrc = ShowOperScreenMaint("Operator", "OpPrrPrintTOT20")
            SDOPrr.UserReply = 0
            Exit Sub
            
        Case pagePrrPrintTOT30    '��ȡ����
            nrc = ShowOperScreenMaint("Operator", "OpPrrPrintTOT30")
            SDOPrr.UserReply = 0
            Exit Sub
            
        Case pagePrrPrintTOT35   '�����ȡ����
            nrc = ShowOperScreenMaint("Operator", "OpPrrPrintTOT35")
            SDOPrr.UserReply = 0
            Exit Sub
        
         Case pagePrrPrintTOT40
            If gSelectOprCommand = "04" And IsPrintAmonalyTrans Then
                 IsPrintAmonalyTrans = False
                currentPage = pageLoadBoxWarning
            Else
                currentPage = pageFunChoice
            End If
        
        Case pagePrrPrintTOT50   '��ӡ������
            nrc = ShowOperScreenMaint("Operator", "OpPrrPrintTOT50")
            If IsPrintAmonalyTrans And gSelectOprCommand = "04" Then   '�ӳ�ǰ��ӡ�쳣ȡ������
                         IsPrintAmonalyTrans = False
            Else
                currentPage = pageFunChoice
            End If
                
        Case pageWarnPNO
            nrc = ShowOperScreenMaint("Operator", "OpWarnPNO")
            currentPage = pageFunChoice
            
        Case pageShowDev10
            nrc = ShowOperScreenMaint("Operator", "OpPrintDev10")
            If BrowserMaint.SubStData = "@ok" Then
                Call PrjDeviceStatus
                Call GetDeviceStatus
                currentPage = pageShowDev20
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageShowDev20
            nrc = ShowOperScreenMaint("Operator", "OpPrintDev20")
            currentPage = pageFunChoice
            
        Case pageShowBoxStat10
            nrc = ShowOperScreenMaint("Operator", "OpShowBox10")
            
            Call PrjBoxStatus
            Call GetBoxStatus
            
            If BrowserMaint.SubStData = "@ok" Then
                If SDOCdm.Available = False Then
                    currentPage = pageShowBoxStat20
                Else
                    currentPage = pageShowBoxStat30
                End If
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageShowBoxStat20
            nrc = ShowOperScreenMaint("Operator", "OpShowBox20")
            currentPage = pageShowBoxStat30
            
        Case pageShowBoxStat30
            nrc = ShowOperScreenMaint("Operator", "OpShowBox30")
            currentPage = pageFunChoice
            
        Case pageLoadBoxWarning
            nrc = ShowOperScreenMaint("Operator", "OpLoadBoxWarning")
            If BrowserMaint.SubStData = "@ok" Then
                Call CloseAndOpenPeriod
                currentPage = pageloadbox10
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageloadbox10
            nrc = ShowOperScreenMaint("Operator", "OpLoadBox10")
            
            If BrowserMaint.SubStData = "@ok" Then
                SDOCdm.TimeOutSecondsFirst = -1
                
                g_bIsResumeBox = False
                nrc = SDOCdm.DoLoadCassette
                If nrc <> 0 Then
                    Call PrjLoadFeederFailed
                    currentPage = pageLoadBox60
                Else
                    Exit Sub
                End If
            Else
                currentPage = pageFunChoice
            End If
            
        Case pageLoadBox11
            nrc = ShowOperScreenMaint("Operator", "OpLoadBox11")
            SDOCdm.UserReply = 0
            Exit Sub
            
        Case pageLoadBox20
            nrc = ShowOperScreenMaint("Operator", "OpLoadBox20")
            SDOCdm.UserReply = 0
            Exit Sub
            
        Case pageLoadBox30
            nrc = ShowOperScreenMaint("Operator", "OpLoadBox30")
            SDOCdm.UserReply = 0
            Exit Sub
        
        Case pageLoadBox40
            nrc = ShowOperScreenMaint("Operator", "OpLoadBox40")
            currentPage = pageLoadBox11
        
        Case pageLoadBox50
            For i = 1 To nNumberOfCassettes
                WthCassette(i).BoxLeftNum = 0
                
                If WthCassette(i).CasLogicalID <> 0 And WthCassette(i).BoxCurry <> "XXX" And _
                        WthCassette(i).BoxState <> "MISS" Then
                    nDenomBox = WthCassette(i).BoxDenom
                
                    nrc = Pcb3dl.DlSetCharRaw("HtmlWork51", CStr(i))
                    nrc = Pcb3dl.DlSetCharRaw("TTU01", CStr(nDenomBox))
                    Select Case WthCassette(i).BoxCurry
                        Case "CNY"
                            nrc = Pcb3dl.DlSetCharRaw("TTU02", "�����")
                        Case "HKD"
                            nrc = Pcb3dl.DlSetCharRaw("TTU02", "��  ��")
                        Case "USD"
                            nrc = Pcb3dl.DlSetCharRaw("TTU02", "��  Ԫ")
                        Case Else
                            nrc = Pcb3dl.DlSetCharRaw("TTU02", "MIZZI")
                    End Select
                    
                    nrc = Pcb3dl.DlSetCharRaw("TTUChoice", "")
        
                    nrc = ShowOperScreenMaint("Operator", "OpLoadBox50")
                
                    If BrowserMaint.SubStData = "@ok" Then
                        sLoadNum = Pcb3dl.DlGetCharRaw("TTUChoice")
                        If IsNumeric(sLoadNum) Then
                            nInputCount = Int(sLoadNum)
                        Else
                            nInputCount = 0
                        End If
                    Else
                        nInputCount = 0
                    End If
                    WthCassette(i).BoxLeftNum = nInputCount
                End If
            Next i
            currentPage = pageLoadBox55
            
        Case pageLoadBox55
            nrc = Pcb3dl.DlReset("TTUChoice")
            Call ClearLoadBoxTable

            NorBinTotLeftAmount = 0
            
            For i = 1 To nNumberOfCassettes
                If WthCassette(i).CasLogicalID <> 0 And WthCassette(i).BoxCurry <> "XXX" And _
                        WthCassette(i).BoxState <> "MISS" Then
                    nrc = Pcb3dl.DlSetCharRaw("HtmlWork1" & CStr(i + 1), WthCassette(i).BoxCurry)
                
                    nrc = Pcb3dl.DlSetCharRaw("HtmlWork2" & CStr(i + 1), _
                          Format(WthCassette(i).BoxDenom, "000"))
                          
                    nrc = Pcb3dl.DlSetCharRaw("HtmlWork3" & CStr(i + 1), _
                          Format(WthCassette(i).BoxLeftNum, "0000"))
                          
                    nrc = Pcb3dl.DlSetCharRaw("HtmlWork4" & CStr(i + 1), _
                          Format(WthCassette(i).BoxDenom * WthCassette(i).BoxLeftNum, "000000"))
                      
                    NorBinTotLeftAmount = NorBinTotLeftAmount + WthCassette(i).BoxLeftNum * WthCassette(i).BoxDenom
                End If
            Next i
            
            nrc = Pcb3dl.DlSetCharRaw("HtmlWork52", Format(NorBinTotLeftAmount, "0000000"))

            nrc = ShowOperScreenMaint("Operator", "OpLoadBox55")
            
            If BrowserMaint.SubStData = "@ok" Then
                sTTUChoice = Pcb3dl.DlGetCharRaw("TTUChoice")
                If sTTUChoice = "0" Then
                    currentPage = pageLoadBox57
                Else
                    currentPage = pageLoadBox50
                End If
            Else
                currentPage = pageLoadBox50
            End If
            
        Case pageLoadBox57
            nrc = ShowOperScreenMaint("Operator", "OpLoadBox57")
                        
            SDOCdm.DataCriteria = 1
            'Add for Opteva to reset reject bin count
            If GLbIsNewPeriod Then
                SDOCdm.CasNbrLogical = 0
                SDOCdm.TotNbrPresent = 0
            End If
            'Add end
            For i = 1 To nNumberOfCassettes
                If WthCassette(i).CasLogicalID <> 0 And WthCassette(i).BoxCurry <> "XXX" And _
                        WthCassette(i).BoxState <> "MISS" Then
                    SDOCdm.CasNbrLogical = WthCassette(i).CasLogicalID
                    If GLbIsNewPeriod Then
                        SDOCdm.TotNbrPresent = WthCassette(i).BoxLeftNum
                        SDOCdm.TotNbrDelivered = 0
                        SDOCdm.TotNbrDeliveredNotTaken = 0
                        SDOCdm.TotNbrDispensedNotDelivered = 0
                    ElseIf SDOCdm.TotNbrPresent = -1 Then
                        SDOCdm.TotNbrPresent = WthCassette(i).BoxLeftNum
                    ElseIf WthCassette(i).BoxLeftNum <> 0 Then
                        SDOCdm.TotNbrPresent = SDOCdm.TotNbrPresent + WthCassette(i).BoxLeftNum
                    End If
                ElseIf WthCassette(i).CasLogicalID <> 0 Then
                    SDOCdm.CasNbrLogical = WthCassette(i).CasLogicalID
                    SDOCdm.TotNbrPresent = 0
                    SDOCdm.TotNbrDelivered = 0
                    SDOCdm.TotNbrDeliveredNotTaken = 0
                    SDOCdm.TotNbrDispensedNotDelivered = 0
                End If
            Next i
            currentPage = pageLoadBox70
        
        Case pageLoadBox60
            nrc = ShowOperScreenMaint("Operator", "OpLoadBox60")
            currentPage = pageFunChoice
        
        Case pageLoadBox61
            nrc = ShowOperScreenMaint("Operator", "OpLoadBox61")
            currentPage = pageFunChoice
        
        Case pageLoadBox70
            sGLtheTime = Format(Now(), "YYYY/MM/DD HH:MM:SS")
            nrc = Pcb3dl.DlSetCharRaw("TotLoadNoteTime", sGLtheTime)
            Call PrjLoadFeederOk
     
            If GLbIsNewPeriod = True Then
                nrc = Pcb3dl.DlSetCharRaw("GBLIsNewPeriod", "N")
            End If
        'Ϊ�Ϻ�������Ӽӳ�����
        Call CommunicationSubFunction("RWT", "AAP")
        '����

            nrc = ShowOperScreenMaint("Operator", "OpLoadBox70")
            
            Pcb3dl.DlSetCharRaw "GBLInitCasStates", "Y"
'           �ӳ�������ص�ѯ�ʿͻ��Ƿ�Ҫ��ӡͳ��ֵ
'            currentPage = pageFunChoice
            currentPage = pageClosePeriod30
        
        Case pageOperReturn10
            Call PrjEndOperator
            nrc = ShowOperScreenMaint("Operator", "OpOperReturn10")
            If BrowserMaint.SubStData = "@ok" Then
                Pcb3dl.DlSetCharRaw "TTU01", "��·����"
                currentPage = pageReturnOk
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageRetainCard10
            nrc = ShowOperScreenMaint("Operator", "OpRetainCard10")
            If BrowserMaint.SubStData = "@ok" Then
                If IsCardRetainExist Then
                    currentPage = pageRetainCard20
                Else
                    currentPage = pageRetainCard30
                End If
            Else
                currentPage = pageFunChoice
            End If

        Case pageRetainCard20
            nrc = ShowOperScreenMaint("Operator", "OpRetainCard20")
            If PrintRetainCard Then
                currentPage = pageRetainCard40
            Else
                currentPage = pageRetainCard30
            End If
            
        Case pageRetainCard30
            nrc = ShowOperScreenMaint("Operator", "OpRetainCard30")
            currentPage = pageFunChoice
            
         Case pageRetainCard40
            nrc = ShowOperScreenMaint("Operator", "OpRetainCard40")
            currentPage = pageFunChoice
           
        Case pageOpChgPwd10
            nrc = ShowOperScreenMaint("Operator", "OpChgPwd10")
            If BrowserMaint.SubStData = "@ok" Then
                currentPage = pageOpChgPwd20
            Else
                If SuperAdminBegin Then
                    currentPage = pageSuperFunctionChoice
                Else
                    currentPage = pageFunChoice
                End If

            End If
            
        Case pageOpChgPwd20
            nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
            nrc = ShowOperScreenMaint("Operator", "OpChgPwd20")
            If BrowserMaint.SubStData = "@ok" Then
                currentPage = pageOpChgPwd30
            Else
                If SuperAdminBegin Then
                    currentPage = pageSuperFunctionChoice
                Else
                    currentPage = pageFunChoice
                End If

            End If
            
        Case pageOpChgPwd30
            nrc = Pcb3dl.DlReset("HtmlInput2")
            nrc = ShowOperScreenMaint("Operator", "OpChgPwd30")
            
            If BrowserMaint.SubStData = "@ok" Then
                sCurTime = Format(Now(), "YYYY/MM/DD HH:MM:SS")
                sFirstNewPin = Pcb3dl.DlGetCharRaw("HtmlInput1")
                sSecondNewPin = Pcb3dl.DlGetCharRaw("HtmlInput2")
                If sFirstNewPin = sSecondNewPin Then
                   If SuperAdminBegin Then
                       ReturnValue = SetIniS(IniPath + "Global.ini", "Operator", "SuperOperPassWord", sFirstNewPin)
                       PrjString = JourLineSeprator + "SuperOPR ChgPWD Time: " + sCurTime + "   " + vbCrLf + _
                                   "       Operator Change Password OK     " + vbCrLf
                       PrjCHNString = JourLineSeprator + "��������Ա����ʱ�䣺" + sCurTime + "   " + vbCrLf + _
                               "       ��������Ա���ܳɹ���" + vbCrLf
                    Else
                       ReturnValue = SetIniS(IniPath + "Global.ini", "Operator", "OperPassWord", sFirstNewPin)
                       PrjString = JourLineSeprator + "OPR ChgPWD Time: " + sCurTime + "   " + vbCrLf + _
                                   "       Operator Change Password OK     " + vbCrLf
                        PrjCHNString = JourLineSeprator + "����Ա����ʱ�䣺" + sCurTime + "   " + vbCrLf + _
                               "       ����Ա���ܳɹ���" + vbCrLf
                    End If
                  
                    
                   PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                   currentPage = pageOpChgPwd40
                Else

                   PrjString = JourLineSeprator + "OPR ChgPWD Time: " + sCurTime + "   " + vbCrLf + _
                               "    Operator Change Password Failed    " + vbCrLf
                   PrjCHNString = JourLineSeprator + "����Ա����ʱ�䣺" + sCurTime + "   " + vbCrLf + _
                               "       ����Ա����ʧ�ܣ�" + vbCrLf
                   PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                   currentPage = pageOpChgPwd50
                End If
                nrc = Pcb3dl.DlReset("HtmlInput1")
                nrc = Pcb3dl.DlReset("HtmlInput2")
            Else
                If SuperAdminBegin Then
                    currentPage = pageSuperFunctionChoice
                Else
                    currentPage = pageFunChoice
                End If

            End If
            
        Case pageOpChgPwd40
            nrc = ShowOperScreenMaint("Operator", "OpChgPwd40")
            If SuperAdminBegin Then
                currentPage = pageSuperFunctionChoice
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageOpChgPwd50
            nrc = ShowOperScreenMaint("Operator", "OpChgPwd50")
            If SuperAdminBegin Then
                currentPage = pageSuperFunctionChoice
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageTestDispenseNoteForEachCas10
            nrc = ShowOperScreenMaint("Operator", "OpTestDisNote10")
            If BrowserMaint.SubStData = "@ok" Then
                currentPage = pageTestDispenseNoteForEachCas20
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageTestDispenseNoteForEachCas20
            nrc = ShowOperScreenMaint("Operator", "OpTestDisNote20")
            
            nrc = SDOCdm.DoWithdrawal
            If nrc <> 0 Then
                currentPage = pageTestDispenseNoteForEachCas40
            Else
                Exit Sub
            End If
        
        Case pageTestDispenseNoteForEachCas30
            PrjString = JourLineSeprator + "    TEST DISPENSE NOTES OK" + vbCrLf + _
                        Format(Now(), "YY/MM/DD HH:MM") + " ATM:" + AtmCode + vbCrLf
            PrjCHNString = JourLineSeprator + "    ���������⡡�ԡ��ɡ�����" + vbCrLf + _
                        Format(Now(), "YY/MM/DD HH:MM") + " ATM�ţ�" + AtmCode + vbCrLf
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            nrc = ShowOperScreenMaint("Operator", "OpTestDisNote30")
            currentPage = pageFunChoice
           
        Case pageTestDispenseNoteForEachCas40
            PrjString = JourLineSeprator + "    TEST DISPENSE NOTES ERROR" + vbCrLf + _
                        Format(Now(), "YY/MM/DD HH:MM") + " ATM:" + AtmCode + vbCrLf
            PrjCHNString = JourLineSeprator + "    ���������⡡�ԡ�ʧ���ܣ�" + vbCrLf + _
                        Format(Now(), "YY/MM/DD HH:MM") + " ATM�ţ�" + AtmCode + vbCrLf
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            nrc = ShowOperScreenMaint("Operator", "OpTestDisNote40")
            currentPage = pageFunChoice
        
        Case pageShutdownSys10
            nrc = Pcb3dl.DlReset("TTUChoice")
            nrc = ShowOperScreenMaint("Operator", "OpShutdownSys10")
            sTTUChoice = Pcb3dl.DlGetCharRaw("TTUChoice")
           If sTTUChoice = "1" Or sTTUChoice = "2" Then
                currentPage = pageShutdownSys20
            Else
                currentPage = pageFunChoice
            End If
            
        Case pageShutdownSys20
            nrc = ShowOperScreenMaint("Operator", "OpShutdownSys20")
          If BrowserMaint.SubStData = "@ok" Then
                Call PrjShutdownSystem
                currentPage = pageShutdownSys30
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageShutdownSys30
            nrc = ShowOperScreenMaint("Operator", "OpShutdownSys30")
            sTTUChoice = Pcb3dl.DlGetCharRaw("TTUChoice")
            If sTTUChoice = "2" Then
                nrc = NTSystemShutDown(EWX_FORCE + EWX_REBOOT)
                If nrc <> 0 Then
                    LogError "Call System function <ExitWindowsEx->EWX_REBOOT> Failed"
                Else
                    LogError "Call System function <ExitWindowsEx->EWX_REBOOT> OK"
                End If
            ElseIf sTTUChoice = "1" Then
                nrc = NTSystemShutDown(EWX_FORCE + EWX_SHUTDOWN)
                If nrc <> 0 Then
                    LogError "Call System function <ExitWindowsEx->EWX_REBOOT> Failed"
                Else
                    LogError "Call System function <ExitWindowsEx->EWX_REBOOT> OK"
                End If
            End If
            Exit Sub
            
        Case pageOprLogCopy10
            nrc = ShowOperScreenMaint("Operator", "OpCopyLog10")
            If BrowserMaint.SubStData = "@ok" Then
         
   'ɾ���й�translog�ļ��Ƿ���ڵ��ж� 2005.12.15
'Modified for adding flush disk
                    currentPage = pageSelectCopyDisk10
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageOprLogCopy20
            nrc = ShowOperScreenMaint("Operator", "OpCopyLog20")
            If BrowserMaint.SubStData = "@ok" Then
                currentPage = pageOprLogCopy40
            Else
                If gSelectOprCommand = "12" Then
                    If fso.FileExists("D:\" + g_sBackupLogFileName) Then
                        fso.DeleteFile "D:\" + g_sBackupLogFileName, True
                    End If
                End If
                currentPage = pageFunChoice
            End If
        
        Case pageOprLogCopy30
            nrc = ShowOperScreenMaint("Operator", "OpCopyLog30")
            If gSelectOprCommand = "12" Then
                If (fso.FileExists("D:\" + g_sBackupLogFileName)) Then
                    fso.DeleteFile "D:\" + g_sBackupLogFileName, True
                End If
            End If
            currentPage = pageFunChoice
            
        Case pageOprLogCopy40
            nrc = ShowOperScreenMaint("Operator", "OpCopyLog40")
            On Error GoTo FileCopyFailed
            
             If gSelectOprCommand = "08" Then
                FileCopy CHNJOURNALFile, sLogTargetDisk + CHNJOURNALBAKFile
                
                'ɾ���й�translog�ļ��Ƿ���ڵ��ж� 2005.12.15
                If IsCardRetainExist Then
                    FileCopy CardRetainFile, sLogTargetDisk + CardRetainBAKFile
                End If
                currentPage = pageOprLogCopy50
            Else
                FileCopy "D:\" + g_sBackupLogFileName, sLogTargetDisk + g_sBackupLogFileName
                fso.DeleteFile "D:\" + g_sBackupLogFileName, True
                currentPage = pageLogBackup60
            End If
        
        Case pageOprLogCopy50
            nrc = ShowOperScreenMaint("Operator", "OpCopyLog50")
            currentPage = pageFunChoice
            
        Case pageLogBackup10
            nrc = ShowOperScreenMaint("Operator", "OpLogBackup10")
            If BrowserMaint.SubStData = "@ok" Then
                g_sBackupLogFileName = ""
                currentPage = pageLogBackup20
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageLogBackup20
            nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
            nrc = Pcb3dl.DlSetCharRaw("HtmlInput2", "")
            nrc = ShowOperScreenMaint("Operator", "OpLogBackup20")
            
            If BrowserMaint.SubStData = "@ok" Then
                currentPage = pageLogBackup30
            Else
                currentPage = pageFunChoice
            End If
            
        Case pageLogBackup30
            nrc = Pcb3dl.DlSetCharRaw("TTUChoice", "")
            nrc = ShowOperScreenMaint("Operator", "OpLogBackup30")
            If BrowserMaint.SubStData = "@ok" Then
                sTTUChoice = Pcb3dl.DlGetCharRaw("TTUChoice")
                If sTTUChoice = "0" Then
                    g_vInputDate = (Format(Pcb3dl.DlGetCharRaw("HtmlInput1"), "0000/00/00") + " " + _
                            Format(Pcb3dl.DlGetCharRaw("HtmlInput2"), "00:00"))
                    vDateNow = Now()
                    If IsDate(g_vInputDate) Then
                        g_vInputDate = CDate(g_vInputDate)
                        If g_vInputDate < vDateNow Then
                            g_sBackupLogFileName = GetLogFileName(g_vInputDate)
                            If Len(g_sBackupLogFileName) = 0 Then
                                nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt2", "��־�ļ�������")
                                currentPage = pageOprLogCopy30
                            Else
                                currentPage = pageLogBackup33
                            End If
                        Else
                            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt2", "�������ڷǷ�")
                            currentPage = pageOprLogCopy30
                        End If
                    Else
                        nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt2", "�������ڷǷ�")
                        currentPage = pageOprLogCopy30
                    End If
                Else
                    currentPage = pageLogBackup20
                End If
            Else
                currentPage = pageLogBackup20
            End If
        
        Case pageLogBackup33
            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", g_sBackupLogFileName)
            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt2", g_sBackupLogFileStartTime)
            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt3", g_sBackupLogFileEndTime)
            
            nrc = ShowOperScreenMaint("Operator", "OpLogBackup33")
            If BrowserMaint.SubStData = "@ok" Then
                currentPage = pageLogBackup35
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageLogBackup35
            nrc = ShowOperScreenMaint("Operator", "OpLogBackup35")
            If PrepBackupLogFile(g_sBackupLogFileName) Then
                    currentPage = pageSelectCopyDisk10
            Else
                nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt2", "��־�ļ�׼��ʧ��")
                currentPage = pageOprLogCopy30
            End If
        
        Case pageLogBackup60
            Pcb3dl.DlSetCharRaw "HtmlPrompt1", sLogTargetDisk + g_sBackupLogFileName
            nrc = ShowOperScreenMaint("Operator", "OpLogBackup60")
            currentPage = pageFunChoice
         
        Case PageChkVersion10
           
            nrc = ShowOperScreenMaint("Operator", "OpChkVersion10")
            If BrowserMaint.SubStData = "@ok" Then
                sValue = GetIniS(sVersionIni, "Information", "Project", "No information")
                nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", sValue)
                currentPage = pageChkVersion20
            Else
                currentPage = pageFunChoice
            End If
            
        Case pageChkVersion20
            nrc = ShowOperScreenMaint("Operator", "OpChkVersion20")
            currentPage = pageFunChoice
       
        Case pageUpdateMasterKey10
            nrc = ShowOperScreenMaint("Operator", "OpUpdateMasterKey10")
            If BrowserMaint.SubStData = "@ok" Then
                currentPage = pageUpdateMasterKey15
            Else
                currentPage = pageFunChoice
            End If

        Case pageUpdateMasterKey15
            nrc = Pcb3dl.DlSetCharRaw("TTUChoice", "")
            nrc = ShowOperScreenMaint("Operator", "OpUpdateMasterKey15")
            
            If BrowserMaint.SubStData = "@ok" Then
                sTemp = Trim(Pcb3dl.DlGetCharRaw("TTUChoice"))
                
                If sTemp = "1" Then
                    currentPage = pageOpKeyInput10
                ElseIf sTemp = "2" Then
                    currentPage = pageUpdateMasterKey20
                Else
                    currentPage = pageUpdateMasterKey15
                End If
            Else
                currentPage = pageFunChoice
            End If

        Case pageUpdateMasterKey20
            nrc = ShowOperScreenMaint("Operator", "OpUpdateMasterKey20")
            nrc = UpdateKeyFile(sLogTargetDisk)
            If nrc = 0 Then
                currentPage = pageUpdateMasterKey30
            Else
                currentPage = pageUpdateMasterKey40
            End If

        Case pageUpdateMasterKey30
            nrc = ShowOperScreenMaint("Operator", "OpUpdateMasterKey30")
            currentPage = pageFunChoice

        Case pageUpdateMasterKey40
            nrc = ShowOperScreenMaint("Operator", "OpUpdateMasterKey40")
            currentPage = pageFunChoice
       
        Case pageOpKeyInput10
            g_nCurKeyTime = 0
            If G_bTrides Then
                ' 3DES
                currentPage = pageOpKeyInput22
            Else
                ' DES
                currentPage = pageOpKeyInput20
            End If
            nrc = Pcb3dl.DlSetCharRaw("HtmlWork13", "")

        Case pageOpKeyInput20
            If g_nCurKeyTime >= NUMBEROFKEYS Then
                currentPage = pageOpKeyInput30
            Else
                nrc = Pcb3dl.DlSetCharRaw("TTU01", CStr(g_nCurKeyTime + 1))
                nrc = ShowOperScreenMaint("Operator", "OpKeyInput20")
                If nrc = 0 Then
                    sHtmlInput = Pcb3dl.DlGetCharRaw("HtmlWork13")
                    If BrowserMaint.SubStData = "@stop" Then
                        LogError "SubStData = @stop in pageOpKeyInput20"
                        currentPage = pageOpKeyInput50
                    ElseIf BrowserMaint.SubStData = "@ok" Or Len(sHtmlInput) = 16 Then
                        GLarrMasKeys(g_nCurKeyTime) = sHtmlInput
                        If IsValidKey(GLarrMasKeys(g_nCurKeyTime)) Then
                            nrc = Pcb3dl.DlSetCharRaw("HtmlWork13", "")
                            currentPage = pageOpKeyInput26
                        Else
                            currentPage = pageOpKeyInput24
                        End If
                    Else
                        nrc = Pcb3dl.DlSetCharRaw("HtmlWork13", sHtmlInput + BrowserMaint.SubStData)
                        currentPage = pageOpKeyInput20
                    End If
                Else
                    currentPage = pageFunChoice
                End If
                    
            End If
       
        Case pageOpKeyInput22
            If g_nCurKeyTime >= NUMBEROFKEYS Then
                currentPage = pageOpKeyInput30
            Else
                nrc = Pcb3dl.DlSetCharRaw("TTU01", CStr(g_nCurKeyTime + 1))
                nrc = ShowOperScreenMaint("Operator", "OpKeyInput22")
                If nrc = 0 Then
                    sHtmlInput = Pcb3dl.DlGetCharRaw("HtmlWork13")
                    If BrowserMaint.SubStData = "@stop" Then
                        LogError "SubStData = @stop in pageOpKeyInput22"
                        currentPage = pageOpKeyInput24
                    ElseIf BrowserMaint.SubStData = "@ok" Or Len(sHtmlInput) = 32 Then
                        GLarrMasKeys(g_nCurKeyTime) = sHtmlInput
                        If IsValidKey(GLarrMasKeys(g_nCurKeyTime)) Then
                            nrc = Pcb3dl.DlSetCharRaw("HtmlWork13", "")
                            currentPage = pageOpKeyInput26
                        Else
                            currentPage = pageOpKeyInput24
                        End If
                    Else
                        nrc = Pcb3dl.DlSetCharRaw("HtmlWork13", sHtmlInput + BrowserMaint.SubStData)
                        currentPage = pageOpKeyInput22
                    End If
                Else
                    currentPage = pageFunChoice
                End If
            End If
            
        Case pageOpKeyInput24
            nrc = ShowOperScreenMaint("Operator", "OpKeyInput24")
            Call PrjKeyInput(False)
            currentPage = pageFunChoice
        
        Case pageOpKeyInput26
            sHtmlInput = GetCheckValue(GLarrMasKeys(g_nCurKeyTime), G_bTrides)
            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", sHtmlInput)
            nrc = ShowOperScreenMaint("Operator", "OpKeyInput26")
            If nrc = 0 Then
                If BrowserMaint.SubStData <> "@stop" Then
                    g_nCurKeyTime = g_nCurKeyTime + 1
                End If
                If G_bTrides Then
                    currentPage = pageOpKeyInput22
                Else
                    currentPage = pageOpKeyInput20
                End If
            Else
                LogError "Show(OpKeyInput26) return failed " + CStr(nrc)
            End If
            
        Case pageOpKeyInput30
            If g_bIsHardware Then
                ' load init key
                nrc = ShowOperScreenMaint("Operator", "OpKeyInput30")
                SDOEdm.KeyType = 0
                nrc = SDOEdm.DoLoadKey()
                If 0 <> nrc Then
                    'nRc = Pcb3dl.DlSetCharRaw("GBLLoadKeyStatus", "E")
                    LogError "DoLoadKey in pageOpKeyInput30 return failed, " + CStr(nrc)
                    currentPage = pageOpKeyInput45
                Else
                    Exit Sub    'Going into the DoLoadKey() process
                End If
            Else
                ' software calc
                If G_bTrides Then
                    ' 3DES
                
        
                    PreMasterKeyB = sHtmlInput1 + sHtmlInput2
                Else
                    ' DES
                  
                End If
                
                SetIniS IniPath + "key.ini", "KeyList", "AK", PreMasterKeyA
                SetIniS IniPath + "key.ini", "KeyList", "BK", PreMasterKeyB
                
                currentPage = pageOpKeyInput35
            End If
            
        Case pageOpKeyInput35
            nrc = ShowOperScreenMaint("Operator", "OpKeyInput35")
            
            sHtmlInput = DoXorKeys()
            
            'only for host simulator
            If g_bIsHardware = True Then
                sKeyMap = ""
                SDOEdm.CryptMode = True
                SDOEdm.CryptType = 1
                
                If G_bTrides Then
                    sHtmlInput1 = Left(sHtmlInput, 16)
                    sHtmlInput2 = Right(sHtmlInput, 16)
                    
                    
                    nrc = SDOEdm.DoCryptDataSw(sHtmlInput1, "EFEFEFEFEFEFEFEF")
                    If nrc = 0 Then
                        sKeyMap = SDOEdm.CryptResult
                    End If
                    nRc1 = SDOEdm.DoCryptDataSw(sHtmlInput2, "EFEFEFEFEFEFEFEF")
                    If nRc1 = 0 Then
                        sKeyMap = sKeyMap + SDOEdm.CryptResult
                    End If
                    
                Else
                    nrc = SDOEdm.DoCryptDataSw(sHtmlInput, "EFEFEFEFEFEFEFEF")
                    If nrc = 0 Then
                        sKeyMap = SDOEdm.CryptResult
                    End If
                    nRc1 = 0
                End If
                
                If nrc = 0 And nRc1 = 0 Then
                    SDOEdm.HostKey = sKeyMap
                    SDOEdm.KeyType = 1
                    SDOEdm.KeyEncName = "OFFL0"
                    SDOEdm.KeyName = "MASKEY"
                    
                    'Loading key for Encrypt, Function(EPP) and Key LoadKey
                    SDOEdm.KeyUse = &H23&
                    
                    nrc = Pcb3dl.DlSetCharRaw("GBLMasterKey", sHtmlInput)
                    nrc = SDOEdm.DoLoadKey()
                    If 0 <> nrc Then
                        'nRc = Pcb3dl.DlSetCharRaw("GBLLoadKeyStatus", "E")
                        LogError "DoLoadKey in pageOpKeyInput35 return failed, " + CStr(nrc)
                        currentPage = pageOpKeyInput50
                    Else
                        Exit Sub    'Going into the DoLoadKey() process
                    End If
                Else
                    LogError "DoCryptData in pageOpKeyInput30 return failed, " + CStr(nrc)
                    'nRc = Pcb3dl.DlSetCharRaw("GBLLoadKeyStatus", "E")
                    currentPage = pageOpKeyInput50
                End If
            Else
                SDOEdm.CryptMode = True
                SDOEdm.CryptType = 1
                sKeyMap = ""
                If G_bTrides Then
                    sHtmlInput1 = Left(sHtmlInput, 16)
                    sHtmlInput2 = Right(sHtmlInput, 16)
                    nrc = SDOEdm.DoCryptDataSw(sHtmlInput1, "EFEFEFEFEFEFEFEF")
                    If nrc = 0 Then
                        sKeyMap = SDOEdm.CryptResult
                    End If
                    nrc = SDOEdm.DoCryptDataSw(sHtmlInput2, "EFEFEFEFEFEFEFEF")
                    If nrc = 0 Then
                        sKeyMap = sKeyMap + SDOEdm.CryptResult
                    End If
                Else
                    nrc = SDOEdm.DoCryptDataSw(sHtmlInput, "EFEFEFEFEFEFEFEF")
                    If nrc = 0 Then
                        sKeyMap = SDOEdm.CryptResult
                    End If
                End If
                nrc = Pcb3dl.DlSetCharRaw("GBLMasterKey", sHtmlInput)
                currentPage = pageOpKeyInput40
            End If
            
        Case pageOpKeyInput40
            nrc = ShowOperScreenMaint("Operator", "OpKeyInput40")
            Call PrjKeyInput(True)
            currentPage = pageFunChoice
            
        Case pageOpKeyInput45
            nrc = ShowOperScreenMaint("Operator", "OpKeyInput45")
            Call PrjKeyInput(False)
            currentPage = pageFunChoice
        
        Case pageOpKeyInput50
            nrc = ShowOperScreenMaint("Operator", "OpKeyInput50")
            Call PrjKeyInput(False)
            currentPage = pageFunChoice

        Case pageResumeBox10
            nrc = ShowOperScreenMaint("Operator", "OpResumeBox10")
            If BrowserMaint.SubStData = "@ok" Then
                currentPage = pageResumeBox20
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageResumeBox20
            nrc = ShowOperScreenMaint("Operator", "OpResumeBox20")
            If BrowserMaint.SubStData = "@ok" Then
                currentPage = pageResumeBox30
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageResumeBox30
            nrc = ShowOperScreenMaint("Operator", "OpResumeBox30")
            
            g_bIsResumeBox = True
                        
            nrc = SDOCdm.DoLoadCassette
            If nrc <> 0 Then
                currentPage = pageResumeBox50
            Else
                Exit Sub
            End If
            
         Case pageResumeBox40
            PrjString = JourLineSeprator + "        RESUME BOX OK            " + vbCrLf + _
                        Format(Now(), "YY/MM/DD HH:MM") + " ATM:" + AtmCode + vbCrLf
            PrjCHNString = JourLineSeprator + "     �֡����������䡡�ɡ�����" + vbCrLf + _
                        Format(Now(), "YY/MM/DD HH:MM") + " ATM�ţ�" + AtmCode + vbCrLf
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
           
            Pcb3dl.DlSetCharRaw "GBLInitCasStates", "Y"
            nrc = ShowOperScreenMaint("Operator", "OpResumeBox40")
            currentPage = pageFunChoice
           
        Case pageResumeBox50
            PrjString = JourLineSeprator + "        RESUME FEEDERS ERROR            " + vbCrLf + _
                        Format(Now(), "YY/MM/DD HH:MM") + " ATM:" + AtmCode + vbCrLf
            PrjCHNString = JourLineSeprator + "     �֡����������䡡ʧ���ܣ�" + vbCrLf + _
                        Format(Now(), "YY/MM/DD HH:MM") + " ATM�ţ�" + AtmCode + vbCrLf
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            nrc = ShowOperScreenMaint("Operator", "OpResumeBox50")
            currentPage = pageFunChoice
           
        Case pageSelectCopyDisk10
            TimerAction.Interval = 100

            nrc = Pcb3dl.DlSetCharRaw("TTUChoice", "")
            nrc = ShowOperScreenMaint("Operator", "OpSelectCopyDisk10")
            
            If BrowserMaint.SubStData = "@ok" Then
                sTemp = Trim(Pcb3dl.DlGetCharRaw("TTUChoice"))
                
                If sTemp = "1" Then
                    sLogTargetDisk = "A:\"
                    If gSelectOprCommand = "18" Then
                        currentPage = pageUpdateMasterKey20
                    Else
                        currentPage = pageOprLogCopy20
                    End If
                ElseIf sTemp = "2" Then
                    currentPage = pageSelectCopyDisk20
                Else
                    currentPage = pageSelectCopyDisk21
                End If
            Else
                currentPage = pageFunChoice
            End If
       
        Case pageSelectCopyDisk20
            nrc = ShowOperScreenMaint("Operator", "OpSelectCopyDisk20")
            If BrowserMaint.SubStData = "@ok" Then
               
                If GetUsbDisk(strResult) Then
                    sLogTargetDisk = Left(strResult, 3)
                    If gSelectOprCommand = "18" Then
                        currentPage = pageUpdateMasterKey20
                    Else
                        currentPage = pageOprLogCopy40
                    End If
                Else
                    currentPage = pageSelectCopyDisk21
                End If
            Else
                currentPage = pageFunChoice
            End If
                    
        Case pageSelectCopyDisk21
            nrc = ShowOperScreenMaint("Operator", "OpSelectCopyDisk21")
            currentPage = pageFunChoice
      
'        Case pagePingHost10
'            nrc = ShowOperScreenMaint("Operator", "OpPingHost10")
'            If BrowserMaint.SubStData = "@ok" Then
'                currentPage = pagePingHost20
'            Else
'                currentPage = pageFunChoice
'            End If
'
'        Case pagePingHost20
'            sTemp = GetIniS("C:\Windows\SST_COM.ini", "SSTSRV", "PrimaryServer", "0.0.0.0")
'            nrc = Pcb3dl.DlSetCharRaw("TTU01", sTemp)
'            If sTemp <> "0.0.0.0" Then
'                nrc = ShowOperScreenMaint("Operator", "OpPingHost20")
'                nrc = Ping(2, sTemp)
'                If nrc <> 0 Then
'                    currentPage = pagePingHost40
'                Else
'                    currentPage = pagePingHost30
'                End If
'            Else
'                currentPage = pagePingHost40
'            End If
'
'        Case pagePingHost30
'            nrc = ShowOperScreenMaint("Operator", "OpPingHost30")
'            PrjString = JourLineSeprator + _
'                    Format(Now(), "YY/MM/DD HH:MM") + "  " + "Ping Host Ok!!" + vbCrLf
'            PrjCHNString = JourLineSeprator + _
'                    Format(Now(), "YY/MM/DD HH:MM") + "  " + "Ping�����������ɡ���!!" + vbCrLf
'            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
'            currentPage = pageFunChoice
'
'        Case pagePingHost40
'            nrc = ShowOperScreenMaint("Operator", "OpPingHost40")
'            PrjString = JourLineSeprator + _
'                    Format(Now(), "YY/MM/DD HH:MM") + "  " + "Ping Host Failed" + vbCrLf
'            PrjCHNString = JourLineSeprator + _
'                    Format(Now(), "YY/MM/DD HH:MM") + "  " + "Ping����������ʧ����!!" + vbCrLf
'            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
'
'            currentPage = pageFunChoice
                
        Case pageExitApp10
            nrc = Pcb3dl.DlReset("TTUChoice")
            nrc = ShowOperScreenMaint("Operator", "OpExitApp10")
            sTTUChoice = Pcb3dl.DlGetCharRaw("TTUChoice")
            If BrowserMaint.SubStData = "@ok" And sTTUChoice = "0" Then
                currentPage = pageExitApp20
            Else
                currentPage = pageFunChoice
            End If
            
        Case pageExitApp20
            nrc = Pcb3dl.DlSetCharRaw("TTUChoice", "")
            nrc = ShowOperScreenMaint("Operator", "OpExitApp20")
            sTTUChoice = Pcb3dl.DlGetCharRaw("TTUChoice")
            If BrowserMaint.SubStData = "@ok" And sTTUChoice = "0" Then
                PrjString = JourLineSeprator + "           EXIT APPLICATION      " + vbCrLf + _
                            " " + Format(Now(), "YY/MM/DD HH:MM") + " ATM:" + AtmCode + vbCrLf
                PrjCHNString = JourLineSeprator + "      �ء��ա�Ӧ���á��̡���" + vbCrLf + _
                            " " + Format(Now(), "YY/MM/DD HH:MM") + " ATM�ţ�" + AtmCode + vbCrLf
                PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                currentPage = pageExitApp30
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageExitApp30
            nrc = ShowOperScreenMaint("Operator", "OpExitApp30")
            nrc = CloseS3EWindow(MonitorWinName, MonitorClassName)
            If nrc = 0 Then
                PrjString = "    Exit S3EMonitor Ok!"
                PrjCHNString = "    �˳�S3EMonitor����ɹ���"
                PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                currentPage = pageExitApp40
                TimerAction.Interval = 5000
            Else
                PrjString = "    Exit S3EMonitor Failed!"
                PrjCHNString = "    �˳�S3EMonitor����ʧ�ܣ�"
                PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                currentPage = pageExitApp31
            End If
       
        Case pageExitApp31
            nrc = ShowOperScreenMaint("Operator", "OpExitApp31")
            currentPage = pageFunChoice
       
        Case pageExitApp40
            nrc = ShowOperScreenMaint("Operator", "OpExitApp40")
            nrc = CloseS3EWindow(MasterWinName, MasterClassName)
            If nrc = 0 Then
                PrjString = "    Exit PowerMaster Ok!"
                PrjCHNString = "    �˳�PowerMaster����ɹ���"
                PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                Exit Sub
            Else
                PrjString = "    Exit PowerMaster Failed!"
                PrjCHNString = "    �˳�PowerMaster����ʧ�ܣ�"
                PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                currentPage = pageExitApp41
                TimerAction.Interval = 100
            End If
       
        Case pageExitApp41
            nrc = ShowOperScreenMaint("Operator", "OpExitApp41")
            currentPage = pageFunChoice
       
        Case pageNoFunAvail
            nrc = ShowOperScreenMaint("Operator", "OpNoFunAvail")
            currentPage = pageFunChoice
        
        Case pageReturnOk
            SDOFep.DoServiceOpen
            
            If Pcb3dl.DlGetCharRaw("CWDCrimePossible") = "Y" Then
                Pcb3dl.DlSetCharRaw "CWDCrimePossible", "N"
                Pcb3dl.DlSetCharRaw "GBLDoRecovery", "N"
            End If
            
            If Browser.HasSecondMonitor = 0 Then
                BrowserMaint.WindowStyle = WINDOWED
                Browser.WindowStyle = TOP_FULL_SCREEN
            End If
            
            nrc = ShowOperScreenMaint("Operator", "OpInService")
            
            nrc = Pcb3dl.DlSetCharRaw("GBLOperStatus", "2")
            nrc = Pcb3dl.DlSetCharRaw("GBLDoRecovery", "1")
'            Pcb3dl.DlSetCharRaw "GBLInitCasStates", "Y"
            S3ETrans.Result = ReturnOk
            Exit Sub

'Add for reset ATM
        Case pageResetATM10
            nrc = Pcb3dl.DlReset("HtmlInput1")
            nrc = ShowOperScreenMaint("Operator", "OpResetATM10")
            If BrowserMaint.SubStData = "@ok" Then
                nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt2", "")
                nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt3", "")
                currentPage = pageResetATM20
            Else
                currentPage = pageFunChoice
            End If
            
        Case pageResetATM20
            If Browser.HasSecondMonitor = 0 Then
                nrc = ShowOperScreenMaint("Operator", "OpResetATM20")
            Else
                nrc = Pcb3dl.DlSetCharRaw("TTUChoice", "")
                nrc = ShowOperScreenMaint("Operator", "OpResetATMTtu")
            End If
            If (BrowserMaint.SubStData <> "@stop") Then
                If Browser.HasSecondMonitor = 0 Then
                    g_sResettingDev = Mid(BrowserMaint.SubStData, 2, 3)
                Else
                    sTemp = Trim(Pcb3dl.DlGetCharRaw("TTUChoice"))
                    Select Case sTemp
                    Case "1"
                        g_sResettingDev = "PRJ"
                    Case "2"
                        g_sResettingDev = "PRR"
                    Case "3"
                        g_sResettingDev = "IDC"
                    Case "4"
                        g_sResettingDev = "EDM"
                    Case "5"
                        g_sResettingDev = "CDM"
                    Case "7"
                        g_sResettingDev = "DEV"
                    End Select
                End If
                If g_sResettingDev <> "DEV" Then
                    currentPage = pageResetATM30
                    Select Case g_sResettingDev
                        Case "PRJ"
                            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt3", "��ˮ��ӡ��")
                        Case "PRR"
                            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt3", "ƾ����ӡ��")
                        Case "IDC"
                            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt3", "�ſ���д��")
                        Case "EDM"
                            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt3", "����ģ��")
                        Case "CDM"
                            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt3", "ȡ��ģ��")
                    End Select
                Else
                    Call PrjDeviceStatus
                    Call GetDeviceStatus
                    currentPage = pageResetATM50
                End If
            Else
                currentPage = pageFunChoice
            End If
            
        Case pageResetATM30
            nrc = ShowOperScreenMaint("Operator", "OpResetATM30")
            
            
            bDisableTimer = ResetReleatedDevice(g_sResettingDev)
            If bDisableTimer = True Then
                Exit Sub
            End If
                        
        Case pageResetATM40
            nrc = Pcb3dl.DlReset("HtmlInput1")
            nrc = ShowOperScreenMaint("Operator", "OpResetATM40")
            currentPage = pageResetATM20
            
        Case pageResetATM50
            nrc = Pcb3dl.DlReset("HtmlInput1")
            nrc = ShowOperScreenMaint("Operator", "OpPrintDev20")
            currentPage = pageResetATM20
           
'Add for entering Vendor Mode
        Case pageEnterVendorMode10
            nrc = ShowOperScreenMaint("Operator", "OpEnterVendorMode10")
            If BrowserMaint.SubStData = "@ok" Then
                currentPage = pageEnterVendorMode20
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageEnterVendorMode20
            nrc = SDOFep.DoEnterVDM
            If nrc <> 0 Then
                nrc = ShowOperScreenMaint("Operator", "OpEnterVendorMode20")
            End If
            currentPage = pageFunChoice
'end of add forentering Vendor Mode
        
        Case pageShowTransItem10
            nrc = ShowOperScreenMaint("Operator", "OpShowTransItem10")
            If BrowserMaint.SubStData = "@ok" Then
                g_nLogCurPos = 0
                g_bIsTranslogMore = True
                currentPage = pageShowTransItem20
            Else
                currentPage = pageFunChoice
            End If
        
        Case pageShowTransItem20
            g_bIsTranslogMore = GetLogRecordsAndRetIsMore(g_nLogCurPos, sDisplayStr)
                        
            nrc = Pcb3dl.DlSetCharRaw("OptevaCasStatus", sDisplayStr)
            nrc = ShowOperScreenMaint("Operator", "OpShowTransItem20")
            If BrowserMaint.SubStData = "@PGUP" Then
            'Show last page
                g_nLogCurPos = g_nLogCurPos - 10
                If g_nLogCurPos < 0 Then
                    g_nLogCurPos = 0
                End If
                currentPage = pageShowTransItem20
            ElseIf BrowserMaint.SubStData = "@PGDN" Then
            'Show next page
                If g_bIsTranslogMore Then
                    g_nLogCurPos = g_nLogCurPos + 10
                End If
                currentPage = pageShowTransItem20
            Else
                g_nLogCurPos = 0
                g_bIsTranslogMore = True
                currentPage = pageFunChoice
            End If

        Case pageIsUpdatePage
            nrc = ShowOperScreenMaint("Operator", "OpIsUpdatePage")
            If BrowserMaint.SubStData = "@ok" Then
                Call PrjShutdownSystem
                currentPage = pageShutdownSys30
            Else
                currentPage = pageFunChoice
            End If
            
 '��ʾ��ǰ������ȡ��ײ���ӡ��ˮ��ѡ���Ƿ��ӡ����
        Case pageDispCWD10
            iDisplayNum = PrepareDisplayRecords("CWD")
            If iDisplayNum = 0 Then
                If gSelectOprCommand = "24" Then
                    currentPage = pageDispCDP20
                Else
                    currentPage = pageLoadBoxWarning
                End If
            Else
                Call DisplayRecords("CWD", "CWD")
            End If
            
        Case pageDispCDP20
        nrc = ShowOperScreenMaint("Operator", "OpDispCDP20")
        currentPage = pageFunChoice
        
        Case pageDispCDP30
            g_nLogCurPos = 0
            g_nLogLastPos = 0
            PrintJournal SDOPrj, TOTPrjString, TOTPrjString, g_sPrjLanguage
            nrc = ShowOperScreenMaint("Operator", "OpPrrPrintTOT10")
    
            If BrowserMaint.SubStData = "@ok" Then
                If SDOPrr.Available = True Then
                      g_sPrrRawData = TOTPrjString
                       IsPrintAmonalyTrans = True
    '                   ��������ƾ������ͨ��DoPrintRaw���ķ���ӡ�����ͨ��DoPrintForm�ķ�����ʵ��
                            Call CalPageNum
                            Call PrrTotal
                Else
                    currentPage = pagePrrPrintTOT50
                End If
            Else
                If gSelectOprCommand = "24" Then
                    currentPage = pageFunChoice
                Else
                    currentPage = pageLoadBoxWarning
                End If
            End If
            
'���ӳ�������Ա����
        Case pageInputSupAdminPassword
            sOperPin = GetIniS(IniPath + "Global.ini", "Operator", "SuperOperPassWord", "888888")
            nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
            nrc = ShowOperScreenMaint("Operator", "OpPinInput")
            If nrc = 0 Then
               If BrowserMaint.SubStData = "@ok" Then
                    sHtmlInput = Pcb3dl.DlGetCharRaw("HtmlInput1")
                    If sHtmlInput = sOperPin Then
                        Call PrjStartSuperOperator
                        currentPage = pageSuperFunctionChoice
                        SuperAdminBegin = True
                    Else
                        nrc = Pcb3dl.DlReset("HtmlInput1")
                        nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "��������Ա�����������")
                        nrc = ShowOperScreenMaint("Operator", "OpShowDialogInputEnter")
                        currentPage = pageInputSupAdminPassword
                    End If
                Else
                    currentPage = pageFunChoice
                End If
            Else
                currentPage = pageFunChoice
            End If
            
        Case pageSuperFunctionChoice
            nrc = Pcb3dl.DlSetCharRaw("TTUChoice", "")
            nrc = ShowOperScreenMaint("Operator", "OpSuperFunChoice")
            currentPage = GetSuperOprFunctionPage(Pcb3dl.DlGetCharRaw("TTUChoice"))
            
        Case pageSuperSetTerminalLuno
            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "�������ն˺�")
            Pcb3dl.DlReset ("TTUChoice")
            nrc = ShowOperScreenMaint("Operator", "OpShowDialogInputLuno")
            If nrc = 0 Then
                If BrowserMaint.SubStData = "@ok" Then
                    sTTUChoice = Pcb3dl.DlGetCharRaw("TTUChoice")
                    nrc = InStr(1, sTTUChoice, ".")
                    If nrc = 0 And IsNumeric(sTTUChoice) Then
                        ReturnValue = SetIniS(IniPath + "Global.ini", "Bank_Environment", "ATMCode", sTTUChoice)
                        nrc = Pcb3dl.DlSetCharRaw("GBLAtmCode", sTTUChoice)
                        nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "�����ն˺ųɹ�")
                        Pcb3dl.DlReset ("HtmlInput1")
                        nrc = ShowOperScreenMaint("Operator", "OpShowDialogInputEnter")
                        currentPage = pageSuperFunctionChoice
                    Else
                        Pcb3dl.DlReset ("HtmlInput1")
                        nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "������������������")
                        nrc = ShowOperScreenMaint("Operator", "OpShowDialogInputEnter")
                        currentPage = pageSuperSetTerminalLuno
                    End If
                Else
                    currentPage = pageSuperFunctionChoice
                End If
            Else
                currentPage = pageSuperFunctionChoice
            End If
            
        Case pageSuperSetBankCode
            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "�޸����д���")
            nrc = Pcb3dl.DlReset("TTUChoice")
            nrc = ShowOperScreenMaint("Operator", "OpChangeBankCode")
            If nrc = 0 Then
                If BrowserMaint.SubStData = "@ok" Then
                    sTTUChoice = Pcb3dl.DlGetCharRaw("TTUChoice")
                    If IsNumeric(sTTUChoice) Then
                        ReturnValue = SetIniS(IniPath + "Global.ini", "Bank_Environment", "BranchCode", sTTUChoice)
                        nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "�������д���ɹ�")
                        Pcb3dl.DlReset ("HtmlInput1")
                        nrc = ShowOperScreenMaint("Operator", "OpShowDialogInputEnter")
                        currentPage = pageSuperFunctionChoice
                    Else
                        Pcb3dl.DlReset ("HtmlInput1")
                        nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "������������������")
                        nrc = ShowOperScreenMaint("Operator", "OpShowDialogInputEnter")
                        currentPage = pageSuperSetBankCode
                    End If
                Else
                    currentPage = pageSuperFunctionChoice
                End If
            Else
                currentPage = pageSuperFunctionChoice
            End If
        '���ӳ�ʼ��key�Ľ��洦��
        Case pageOpResetTransKey
            nrc = ShowOperScreenMaint("Operator", "OpResetTransKey")
            If BrowserMaint.SubStData = "@ok" Then
                nrc = Pcb3dl.DlSetCharRaw("ResetTransKey", "O")
                nrc = ShowOperScreenMaint("Operator", "OpResetTransKey1")
                currentPage = pageOpResetTransKey1
            Else
                currentPage = pageSuperFunctionChoice
            End If
            
        '���ڳ�ʼ��
        Case pageOpResetTransKey1
            If Pcb3dl.DlGetCharRaw("ResetTransKey") = "N" Then
                currentPage = pageOpResetTransKey2
            ElseIf Pcb3dl.DlGetCharRaw("ResetTransKey") = "W" Then
                currentPage = pageOpResetTransKey3
                nrc = Pcb3dl.DlSetCharRaw("ResetTransKey", "R")
            End If
            
        '��ʼ���ɹ�
        Case pageOpResetTransKey2
            nrc = ShowOperScreenMaint("Operator", "OpResetTransKey2")
            currentPage = pageSuperFunctionChoice
            
        '��ʼ��ʧ��
        Case pageOpResetTransKey3
            nrc = ShowOperScreenMaint("Operator", "OpResetTransKey3")
            currentPage = pageSuperFunctionChoice
        '������ת�˽��׻���
        Case pageSendRTT
            Call CommunicationSubFunction("RTT", "AAP")
            nrc = ShowOperScreenMaint("Operator", "OpSendRTT")
                currentPage = pageFunChoice
        
        Case pageScreenError
            Exit Sub
        
        Case pageQuit
            Unload Operator
            Exit Sub
            
        Case Else
            LogError "TimerAction next action case error. The next action is:" + _
                CStr(currentPage)
    End Select
    
    TimerAction.Enabled = True
    Exit Sub
    
FileCopyFailed:
    Select Case Err.Number
        Case 61
            Pcb3dl.DlSetCharRaw "HtmlPrompt2", "���̻�USB���������ļ�̫��!"
        Case 70
            Pcb3dl.DlSetCharRaw "HtmlPrompt2", "���̱�д����!"
        Case 7
            Pcb3dl.DlSetCharRaw "HtmlPrompt2", "���̻�USB��δ�ź�!"
        Case 76
            Pcb3dl.DlSetCharRaw "HtmlPrompt2", "δ�������̻�USB��!"
        Case Else
            Pcb3dl.DlSetCharRaw "HtmlPrompt2", "���̻�USB���ѻ�!"
    End Select

    If gSelectOprCommand = "18" Then
        currentPage = pageUpdateMasterKey40
    Else
        currentPage = pageOprLogCopy30
    End If
    TimerAction.Enabled = True
    Exit Sub
    
End Sub
Private Function GetOprFunctionPage(OprCommand As String) As pageType
    Dim theTime          As String
    Dim PrjString        As String
    Dim PrjCHNString     As String
    
    gSelectOprCommand = ""
    
    Select Case OprCommand
'        Case "01" 'Show operator command list
'            GetOprFunctionPage = pageCmdList10
    
        Case "02"  'Print Totals
            GetOprFunctionPage = pagePrintTotal10
        
        Case "07" '�豸�Լ�
            GetOprFunctionPage = pageResetATM10
        
        Case "11" 'Return to idle
            nrc = Pcb3dl.DlSetLong("GBLCdmRecoveryTimes", 3)
            GetOprFunctionPage = pageOperReturn10
        
        Case "15" 'Change Operator password
            GetOprFunctionPage = pageOpChgPwd10
              
'        Case "06" 'Open period
'            GLsPeriodStatus = Pcb3dl.DlGetCharRaw("GBLPeriodStatus")
'            If GLsPeriodStatus = "C" Then
'                GetOprFunctionPage = pageOpenPeriod10
'            Else
'                GetOprFunctionPage = pageWarnPNC
'            End If
            
        Case "03"     '�������,�������򿪹رջ�����ڣ����������ͳ��ֵ
            Call PrintCutOffData    'by Chenlei for Boc_Fujian, ��ӡ�ս�����
            GetOprFunctionPage = pageClosePeriod10
            
        Case "08" 'Operator copy trans log
            GetOprFunctionPage = pageOprLogCopy10
            gSelectOprCommand = OprCommand
        
        Case "09" 'Display device status
            GetOprFunctionPage = pageShowDev10
        
        Case "10" '��ת�˽���
            GetOprFunctionPage = pageSendRTT
            
        Case "04" 'Load cassettes
            Call PrintCutOffData    'by Chenlei for Boc_Fujian, ��ӡ�ս�����
            GetOprFunctionPage = pageDispCWD10   'ǿ�ƴ�ӡ�쳣ȡ����ˮ
            gSelectOprCommand = OprCommand
            
        Case "12" 'DBLOGXX trace file copy
            GetOprFunctionPage = pageLogBackup10
            gSelectOprCommand = OprCommand
                        
        Case "13" 'Print retain card table
            GetOprFunctionPage = pageRetainCard10
                    
        
        Case "17" 'Resume box, not load bills
            GetOprFunctionPage = pageResumeBox10
            
'        Case "18"
'            GetOprFunctionPage = pageUpdateMasterKey10
'            gSelectOprCommand = OprCommand
            
        Case "20" 'Shutdown System
            GetOprFunctionPage = pageShutdownSys10
                               
        Case "30" 'Exit Application
            GetOprFunctionPage = pageExitApp10
                               
        Case "52" 'Display boxes status
            GetOprFunctionPage = pageShowBoxStat10
        
        Case "06" 'Check system version
            GetOprFunctionPage = PageChkVersion10
        
        Case "23"
            GetOprFunctionPage = pageInputSupAdminPassword
        
        Case "24" '�鿴���ɽ�����ϸ
            GetOprFunctionPage = pageDispCWD10
            gSelectOprCommand = OprCommand
            
'        Case "25"
'            GetOprFunctionPage = pageShowTransItem10
            
        Case "28" '�㳮����
            GetOprFunctionPage = pageTestDispenseNoteForEachCas10
             
'        Case "90" 'Enter Vendor Mode
'            GetOprFunctionPage = pageEnterVendorMode10
        
        Case Else
            nrc = Pcb3dl.DlSetCharRaw("TTU01", OprCommand)
            GetOprFunctionPage = pageNoFunAvail
    End Select
    
    theTime = Format(Now(), "HH:MM")
    If GetOprFunctionPage = pageNoFunAvail Then
        PrjString = theTime + " FUNCTION" + OprCommand + " INVALIDATED" + vbCrLf
        PrjCHNString = theTime + " ����Ա���� [" + OprCommand + "] ��Ч" + vbCrLf
    Else
        PrjString = theTime + " FUNCTION " + OprCommand + " SELECTED" + vbCrLf
        PrjCHNString = theTime + " ����Աѡ������ [" + OprCommand + "]" + vbCrLf
    End If
    
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage

End Function
Private Function ShowOperScreenMaint(ByVal Section As String, ByVal ScreenName As String) As Integer
    Dim sStr As String
    Dim Path As String
    Dim ReturnCode As Integer
    
    sStr = GetIniS("Screens.ini", "OperatorMaint", ScreenName, "")
    ScreenInfo = GetScreenInfo(sStr)

    Path = GetIniS("Screens.ini", "OperatorMaint", "path", "")
    If Len(ScreenInfo.Name) <> 0 Then
        ReturnCode = BrowserMaint.DoShowScreenSync(Trim(Path) + "\" + ScreenInfo.Name, 0)
        If ReturnCode <> 0 Then
            LogError "ShowOperScreenMaint '" + ScreenInfo.Name + "' Error, Rc = " & CStr(ReturnCode)
        End If
    Else
        ReturnCode = -1
    End If
    
    ShowOperScreenMaint = ReturnCode
End Function
Private Sub CheckupFileExist()
'    Dim TempA As String
    Dim TempB As String
 
'    TempA = Dir(TransLogFile)
'    If Len(TempA) <> 0 Then
'       IsTransLogExist = True
'    Else
'       IsTransLogExist = False
'    End If
    TempB = Dir(CardRetainFile)
    If Len(TempB) <> 0 Then
       IsCardRetainExist = True
    Else
       IsCardRetainExist = False
    End If

    Call FlushBoxesStatusRetIsPresent
End Sub

Private Sub GetBoxStatus()
    Dim i                   As Byte
    Dim NorBinTotLeftAmount As Long
    Dim RejBinTotNotes      As Long
    Dim RejBinTotAmount     As Long
    Dim TmpValue            As String
    Dim j                   As Byte
    
    Call FlushBoxesStatusRetIsPresent
    
    For j = 1 To nNumberOfCassettes
        Select Case j
            Case 1:
                TmpValue = "XXX"
            Case 2:
                TmpValue = "000"
            Case 3:
                TmpValue = "0000"
            Case 4:
                TmpValue = "000000"
            Case 5:
                TmpValue = "MISS"
        End Select
        For i = 1 To nNumberOfCassettes + 1
            Pcb3dl.DlSetCharRaw "HtmlWork" & CStr(j) & CStr(i + 1), TmpValue
        Next i
    Next j
    
    Pcb3dl.DlSetCharRaw "HtmlWork62", "0000000"
    Pcb3dl.DlSetCharRaw "HtmlWork64", "000000"
    
    RejBinTotAmount = 0
    RejBinTotNotes = 0
    NorBinTotLeftAmount = 0

    For i = 1 To nNumberOfCassettes
        RejBinTotNotes = RejBinTotNotes + WthCassette(i).PurgedNotes
        RejBinTotAmount = RejBinTotAmount + WthCassette(i).BoxDenom * WthCassette(i).PurgedNotes
        NorBinTotLeftAmount = NorBinTotLeftAmount + WthCassette(i).BoxDenom * WthCassette(i).BoxLeftNum
    Next
    SDOCdm.CasNbrLogical = 0
    
    WthCassette(nNumberOfCassettes + 1).BoxLeftNum = RejBinTotAmount
    WthCassette(nNumberOfCassettes + 1).BoxState = TranslateBoxState(SDOCdm.CasState, False)
    WthCassette(nNumberOfCassettes + 1).BoxStateCHN = TranslateBoxState(SDOCdm.CasState, True)
    
    For i = 1 To nNumberOfCassettes + 1
        If i = nNumberOfCassettes + 1 Then
            Pcb3dl.DlSetCharRaw "HtmlWork3" & CStr(i + 1), Format(RejBinTotNotes, "0000")
            Pcb3dl.DlSetCharRaw "HtmlWork4" & CStr(i + 1), Format(RejBinTotAmount, "000000")
        Else
            Pcb3dl.DlSetCharRaw "HtmlWork1" & CStr(i + 1), WthCassette(i).BoxCurry
            Pcb3dl.DlSetCharRaw "HtmlWork2" & CStr(i + 1), Format(WthCassette(i).BoxDenom, "000")
            Pcb3dl.DlSetCharRaw "HtmlWork3" & CStr(i + 1), Format(WthCassette(i).BoxLeftNum, "0000")
            Pcb3dl.DlSetCharRaw "HtmlWork4" & CStr(i + 1), Format(WthCassette(i).BoxDenom * WthCassette(i).BoxLeftNum, "000000")
        End If

        Pcb3dl.DlSetCharRaw "HtmlWork5" & CStr(i + 1), WthCassette(i).BoxStateCHN
    Next i
        
    Pcb3dl.DlSetCharRaw "HtmlWork62", Format(NorBinTotLeftAmount, "0000000")
    Pcb3dl.DlSetCharRaw "HtmlWork64", Format(RejBinTotAmount, "000000")
    
End Sub
Private Sub GetDeviceStatus()
    Dim i         As Byte
    Dim PrjStatus As String
    Dim PrrStatus As String
    Dim BGRStatus As String
    Dim CDMStatus As String
    Dim LineStatus As String
    
    For i = 1 To 4
        Pcb3dl.DlSetCharRaw "HtmlWork1" & CStr(i + 1), "����"
    Next i
     
    PrjStatus = TranslateDeviceState("PRJ", True)
    Pcb3dl.DlSetCharRaw "HtmlWork12", PrjStatus
     
    PrrStatus = TranslateDeviceState("PRR", True)
    Pcb3dl.DlSetCharRaw "HtmlWork13", PrrStatus
     
    BGRStatus = TranslateDeviceState("IDC", True)
    Pcb3dl.DlSetCharRaw "HtmlWork14", BGRStatus
     
    CDMStatus = TranslateDeviceState("CDM", True)
    Pcb3dl.DlSetCharRaw "HtmlWork15", CDMStatus
    
    LineStatus = Pcb3dl.DlGetCharRaw("GBLLineStatus")
    If LineStatus = "O" Then
        Pcb3dl.DlSetCharRaw "HtmlWork16", "����"
    Else
        Pcb3dl.DlSetCharRaw "HtmlWork16", "����"
        
    End If
End Sub

'==========================================================================================
'�������� : �õ�����״̬
'������� ������״ֵ̬,�Ƿ���������ʾ״̬��Ϣ
'������� ����
'����ֵ   ��״̬��Ϣ���ַ�����
'���ú��� ��
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Function TranslateBoxState(ByVal StateValue As Byte, ByVal bChinese As Boolean) As String
    Dim CDMBoxState As String
    Dim CDMBoxStateCHN As String

    Select Case StateValue
        Case casstate_cdm_ok:
            CDMBoxState = "OK"
            CDMBoxStateCHN = "������"
        Case casstate_cdm_full, casstate_cdm_high:
            CDMBoxState = "FULL"
            CDMBoxStateCHN = "������"
        Case casstate_cdm_low:
            CDMBoxState = "LOW"
            CDMBoxStateCHN = "���١�"
        Case casstate_cdm_empty:
            CDMBoxState = "EMPT"
            CDMBoxStateCHN = "�����"
        Case casstate_cdm_inoperative:
            CDMBoxState = "BAD"
            CDMBoxStateCHN = "���ϡ�"
        Case casstate_cdm_missing:
            CDMBoxState = "MISS"
            CDMBoxStateCHN = "δ��װ"
        Case Else
            CDMBoxState = "UNKN"
            CDMBoxStateCHN = "δ֪��"
    End Select
    If bChinese Then
        TranslateBoxState = CDMBoxStateCHN
    Else
        TranslateBoxState = CDMBoxState
    End If
End Function

'==========================================================================================
'�������� :������ʾ����״̬��DL����
'������� ����
'������� ����
'����ֵ   ����
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Sub ClearLoadBoxTable()
    Dim i As Byte
    
    For i = 1 To nNumberOfCassettes
        Pcb3dl.DlSetCharRaw "HtmlWork1" & CStr(i + 1), "XXX"
        Pcb3dl.DlSetCharRaw "HtmlWork2" & CStr(i + 1), "000"
        Pcb3dl.DlSetCharRaw "HtmlWork3" & CStr(i + 1), "0000"
        Pcb3dl.DlSetCharRaw "HtmlWork4" & CStr(i + 1), "000000"
    Next i
        
    Pcb3dl.DlSetCharRaw "HtmlWork52", "0000000"

End Sub

'==========================================================================================
'�������� :���Ұ���ĳһ�ض�ʱ���ϵͳ�ļ���udbdxx.log��
'������� ����Ҫ���ҵ�ʱ��㣨��ʽ��YYYYMMDDHHMM��
'������� ����
'����ֵ   ����Ҫ���ҵ��ļ�����(�ַ���)
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Function GetLogFileName(vInputDate As Variant) As String
    Dim fso                   As New FileSystemObject
    Dim ReturnCode            As String
    Dim sFileName             As String
    Dim vDate                 As Variant
    Dim obFile                As File
    Dim nLastFileIndex        As Integer
    Dim nCurIndex             As Integer
    Dim vFileModifiedDate     As Variant
    Dim vLastFileModifiedDate As Variant
    Dim bIsForward            As Boolean

    Set obFile = fso.GetFile(LogBackupPath + "udbd000.log")
    
    vFileModifiedDate = obFile.DateLastModified
    
    If vInputDate < vFileModifiedDate Then
        bIsForward = False
        nCurIndex = 99
    Else
        bIsForward = True
        nCurIndex = 1
    End If
        
    nLastFileIndex = 0
    vLastFileModifiedDate = vFileModifiedDate
    
    Do
        sFileName = "udbd" + Format(nCurIndex, "000") + ".log"

        If Not (fso.FileExists(LogBackupPath + sFileName)) Then
            ReturnCode = "udbd" + Format(CStr(nLastFileIndex), "000") + ".log"
            If bIsForward Then
                g_sBackupLogFileEndTime = Format(Now(), "YY/MM/DD HH:MM:SS")
            Else
                g_sBackupLogFileEndTime = Format(vFileModifiedDate, "YY/MM/DD HH:MM:SS")
            End If
            Exit Do
        Else
            Set obFile = fso.GetFile(LogBackupPath + sFileName)
            vFileModifiedDate = obFile.DateLastModified
            If bIsForward Then
                If vFileModifiedDate < vLastFileModifiedDate Then
                    'Reach the end one of files loop, the lastest file is target file
                    ReturnCode = "udbd" + Format(nLastFileIndex, "000") + ".log"
                    g_sBackupLogFileEndTime = Format(Now(), "YY/MM/DD HH:MM:SS")
                    Exit Do
                ElseIf vInputDate < vFileModifiedDate Then
                    'Found the file
                    g_sBackupLogFileEndTime = Format(vFileModifiedDate, "YY/MM/DD HH:MM:SS")
                    ReturnCode = sFileName
                    Exit Do
                Else
                    nLastFileIndex = nCurIndex
                    nCurIndex = nCurIndex + 1
                    vLastFileModifiedDate = vFileModifiedDate
                End If
            Else
                If vFileModifiedDate > vLastFileModifiedDate Or _
                            vInputDate > vFileModifiedDate Then
                    'Reach the begin one of files loop, the earliest file is target file
                    ReturnCode = "udbd" + Format(nLastFileIndex, "000") + ".log"
                    g_sBackupLogFileEndTime = Format(vLastFileModifiedDate, "YY/MM/DD HH:MM:SS")
                    Exit Do
                Else
                    nLastFileIndex = nCurIndex
                    nCurIndex = nCurIndex - 1
                    vLastFileModifiedDate = vFileModifiedDate
                End If
            End If
        End If
    Loop Until False
    
    'Find log file start time
    vDate = LogFileStartTime(ReturnCode)
    If vInputDate < vDate Then
        ReturnCode = ""
        Exit Function
    End If
    g_sBackupLogFileStartTime = Format(vDate, "YY/MM/DD HH:MM:SS")
    
    GetLogFileName = ReturnCode
End Function
'==========================================================================================
'�������� : ѹ��ϵͳ�ļ���D��
'������� ����Ҫѹ����ϵͳ�ļ�����
'������� ����
'����ֵ   ���Ƿ�ѹ���ɹ���־(����ֵ)
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Function PrepBackupLogFile(sFileName As String) As Boolean
        
    nrc = S3EZip.DoZipTo(LogBackupPath + sFileName, "D:\" + Left(sFileName, 7) + ".zip", 0)
    g_sBackupLogFileName = Left(sFileName, 7) + ".zip"
    
    If nrc = 0 Then
        PrepBackupLogFile = True
    Else
        g_sBackupLogFileName = ""
        PrepBackupLogFile = False
    End If

End Function
'==========================================================================================
'�������� : ͨ���ļ������������Կ
'������� ��Դ����������
'������� ����
'����ֵ   ����������Կ�Ƿ�ɹ���־(����)
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Function UpdateKeyFile(ObjectDisk As String) As Integer
    Dim PreMasterKeyA            As String
    Dim PreMasterKeyB            As String
    Dim UpMaskeyFileLocation     As String
    Dim sValue                   As String
    Dim sValue1                  As String
    Dim sValue2                  As String
    Dim PrjString                As String
    Dim PrjCHNString             As String
    
    On Error Resume Next

    PrjString = " " + vbCrLf + _
                   "   **INSTALL MASTER KEY " + Format(Now(), "HH:MM:SS")
    PrjCHNString = " " + vbCrLf + _
                   "   **��ʼ������Կ " + Format(Now(), "HH:MM:SS")
    
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                   
    UpdateKeyFile = -1
    UpMaskeyFileLocation = ObjectDisk & "ATM_MK.TXT"
    
    Open UpMaskeyFileLocation For Input As #1   ' Open file for input.
    
    If EOF(1) Then
        Close #1
        Exit Function
    End If
    
    Input #1, PreMasterKeyA
    If G_bTrides Then
        ' 3DES
        If Len(PreMasterKeyA) <> 32 Then
            PrjString = "     INSTALL 32 MASTER KEY FAILED!!!"
            PrjCHNString = "     ��ʼ��32λ����Կʧ��!!!"
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            Close #1
            Exit Function
        End If
    Else
        ' DES
        If Len(PreMasterKeyA) <> 16 Then
            PrjString = "     INSTALL 16 MASTER KEY FAILED!!!"
            PrjCHNString = "     ��ʼ��16λ����Կʧ��!!!"
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            Close #1
            Exit Function
        End If
    End If
    
    If EOF(1) Then
        PrjString = "     INSTALL MASTER KEY FAILED!!!"
        PrjCHNString = "     ��ʼ������Կʧ��!!!"
        PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
        Close #1
        Exit Function
    End If
    
    Input #1, PreMasterKeyB
    If G_bTrides Then
        ' 3DES
        If Len(PreMasterKeyB) <> 32 Then
            PrjString = "     INSTALL 32 MASTER KEY FAILED!!!"
            PrjCHNString = "     ��ʼ��32λ����Կʧ��!!!"
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            Close #1
            Exit Function
        End If
    Else
        ' DES
        If Len(PreMasterKeyB) <> 16 Then
            PrjString = "     INSTALL 16 MASTER KEY FAILED!!!"
            PrjCHNString = "     ��ʼ��16λ����Կʧ��!!!"
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            Close #1
            Exit Function
        End If
    End If

    If Not g_bIsHardware Then
        ' software calc
        SetIniS IniPath + "key.ini", "KeyList", "AK", PreMasterKeyA
        
        SetIniS IniPath + "key.ini", "KeyList", "BK", PreMasterKeyB
    End If
    Close #1
    UpdateKeyFile = 0
    
    PrjString = "     INSTALL MASTER KEY OK"
    PrjCHNString = "     ��ʼ������Կ�ɹ�!!!"
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    
    If G_bTrides Then
        ' 3DES
        sValue1 = Left(PreMasterKeyA, 16)
        SDOEdm.CryptType = 1
        SDOEdm.CryptMode = False
        nrc = SDOEdm.DoCryptDataSw(sValue1, "1968707298242918")
        sValue1 = SDOEdm.CryptResult
        
        sValue2 = Right(PreMasterKeyA, 16)
        SDOEdm.CryptType = 1
        SDOEdm.CryptMode = False
        nrc = SDOEdm.DoCryptDataSw(sValue2, "1968707298242918")
        sValue2 = SDOEdm.CryptResult
        
        GLarrMasKeys(0) = sValue1 + sValue2
        
        sValue1 = Left(PreMasterKeyB, 16)
        SDOEdm.CryptMode = False
        nrc = SDOEdm.DoCryptDataSw(sValue1, "8192428927078691")
        sValue1 = SDOEdm.CryptResult
        
        sValue2 = Right(PreMasterKeyB, 16)
        SDOEdm.CryptMode = False
        nrc = SDOEdm.DoCryptDataSw(sValue2, "8192428927078691")
        sValue2 = SDOEdm.CryptResult
        
        GLarrMasKeys(1) = sValue1 + sValue2
        
        sValue = DoXorKeys()
    Else
        ' DES
        SDOEdm.CryptType = 1
        SDOEdm.CryptMode = False
        nrc = SDOEdm.DoCryptDataSw(PreMasterKeyA, "1968707298242918")
        GLarrMasKeys(0) = SDOEdm.CryptResult
        
        
        SDOEdm.CryptMode = False
        nrc = SDOEdm.DoCryptDataSw(PreMasterKeyB, "8192428927078691")
        GLarrMasKeys(1) = SDOEdm.CryptResult
        
        
        sValue = DoXorKeys()
    End If
    
    nrc = Pcb3dl.DlSetCharRaw("GBLMasterKey", sValue)
    
End Function
'==========================================================================================
'�������� :�õ�ϵͳ�ļ���udbdxx.log����¼����ʼʱ��
'������� ��ϵͳ�ļ�����
'������� ����
'����ֵ   ��ʱ��(Variant��)
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Function LogFileStartTime(ByVal sLogFile As String) As Variant
    Dim fso       As New FileSystemObject
    Dim tmpStr    As String
    Dim Cont      As Integer
    Dim LogStream As TextStream
    Dim sLogRec   As String
    Dim sDate     As String
    Dim sText     As String
    Dim i         As Integer
    
    sLogRec = ""
    sDate = ""
    sLogFile = LogBackupPath + sLogFile
    Set LogStream = fso.OpenTextFile(sLogFile, ForReading)
    sLogRec = LogStream.Read(200)
    LogStream.Close
    Cont = 0
    i = 4
    Do
    tmpStr = Mid(sLogRec, i, 2)
    
    sText = WToA(tmpStr, -1, 0)
    
    If Cont = 7 Then sDate = sDate + sText
    
    If sText = Chr(9) Then Cont = Cont + 1
    i = i + 2
    Loop Until ((Cont = 8) Or (i > 210))
    If Len(sDate) > 5 Then sDate = Left$(sDate, Len(sDate) - 5)
    If IsDate(sDate) Then LogFileStartTime = CDate(sDate)
End Function
'==========================================================================================
'�������� :���ݿͻ���������ݴ�ȡ���¼�ļ��в�����Ӧ�ļ�¼
'������� ����ʼ�кţ���������
'������� ����Ļ��ʾ����ǰλ��
'����ֵ   ���ļ���¼�Ƿ��У�����ֵ��
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Function GetLogFindRecordsAndRetIsMore(StartLine As Integer, FindInfo As String, ByRef LogRecDisp As String, ByRef CurPos As Integer) As Boolean
    Dim obj          As New FileSystemObject
    Dim nRecNum      As Integer
    Dim sLogRec      As String
    Dim i            As Integer
    Dim LogRecInfo() As String
    
    nRecNum = 1
    If obj.FileExists(TransLogFile) Then
        Dim LogStream As TextStream
        LogRecDisp = ""
        Set LogStream = obj.OpenTextFile(TransLogFile, ForReading)
        For i = 1 To StartLine
            LogStream.SkipLine
        Next
        CurPos = StartLine
        Do While (Not LogStream.AtEndOfStream) And (nRecNum <= 10)
            sLogRec = LogStream.ReadLine
            LogRecInfo = Split(sLogRec, "|", -1, 1)

            If InStr(LogRecInfo(3), FindInfo) > 0 Then
                LogRecDisp = LogRecDisp + sLogRec + "|"
                nRecNum = nRecNum + 1
            End If
            CurPos = CurPos + 1
        Loop
        GetLogFindRecordsAndRetIsMore = Not LogStream.AtEndOfStream
        LogStream.Close
    Else
        LogError TransLogFile + " is not exist"
    End If
    For i = nRecNum To 10
        LogRecDisp = LogRecDisp + "&nbsp|&nbsp|&nbsp|&nbsp|&nbsp|&nbsp|&nbsp|"
    Next
End Function
'==========================================================================================
'�������� :��ȡ���¼�ļ��еõ�����ʾ��Ӧ�ļ�¼
'������� ����ʼ�к�
'������� ����Ļ��ʾ
'����ֵ   ���ļ���¼�Ƿ��У�����ֵ��
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Function GetLogRecordsAndRetIsMore(StartLine As Integer, ByRef LogRecDisp As String) As Boolean
    Dim obj           As New FileSystemObject
    Dim nRecNum       As Integer
    Dim sLogRec       As String
    Dim i             As Integer
    Dim LogStream     As TextStream
    
    nRecNum = 1
    If obj.FileExists(TransLogFile) Then
        LogRecDisp = ""
        
        Set LogStream = obj.OpenTextFile(TransLogFile, ForReading)
        
        For i = 1 To StartLine
            LogStream.SkipLine
        Next
        
        Do While (Not LogStream.AtEndOfStream) And (nRecNum <= 10)
            sLogRec = LogStream.ReadLine
            LogRecDisp = LogRecDisp + sLogRec + "|"
            nRecNum = nRecNum + 1
        Loop
        
        GetLogRecordsAndRetIsMore = Not LogStream.AtEndOfStream
        
        LogStream.Close
        
    Else
        LogError TransLogFile + " is not exist"
    End If

    For i = nRecNum To 10
        LogRecDisp = LogRecDisp + "&nbsp|&nbsp|&nbsp|&nbsp|&nbsp|&nbsp|&nbsp|"
    Next
        
End Function
'==========================================================================================
'�������� :ˢ�µ�ǰ����״̬
'������� ����
'������� ����
'����ֵ   ���Ƿ�ɹ�������ֵ��
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Function FlushBoxesStatusRetIsPresent() As Boolean
    Dim bCheckResult       As Boolean
    Dim nNumOfBoxesUsed    As Byte
    Dim nNumOfAvailBoxes   As Byte
    Dim i                  As Byte
    Dim j                  As Byte
    Dim CasPosition        As Byte

    bCheckResult = True
    
    nNumOfBoxesUsed = SDOCdm.NbrOfBoxesUsed
    
    SDOCdm.DataCriteria = 1
        
    For i = 1 To nNumberOfCassettes + 1
        WthCassette(i).BoxCurry = "XXX"
        WthCassette(i).BoxDenom = 0
        WthCassette(i).BoxDisp = 0
        WthCassette(i).PurgedNotes = 0
        WthCassette(i).BoxState = "MISS"
        WthCassette(i).BoxStateCHN = "��"
        WthCassette(i).BoxLeftNum = 0
        WthCassette(i).CasLogicalID = 0
        WthCassette(i).BoxInitial = 0
    Next i
    
    nNumOfAvailBoxes = 0
    
    j = 1
    
    g_nRetractCount = 0
      For i = 1 To nNumOfBoxesUsed
        SDOCdm.CasNbrLogical = i
        If Len(SDOCdm.CasPosition) > 0 Then
            If SDOCdm.CasState <= casstate_cdm_empty And SDOCdm.CasState >= casstate_cdm_ok And _
                IsNumeric(Right(SDOCdm.CasPosition, 1)) Then
                CasPosition = CByte(Right(SDOCdm.CasPosition, 1))
                WthCassette(CasPosition).CasLogicalID = i
                WthCassette(CasPosition).BoxCurry = SDOCdm.CasCurrency
                WthCassette(CasPosition).BoxDenom = SDOCdm.CasDenomination
                WthCassette(CasPosition).BoxDisp = SDOCdm.TotNbrDelivered
                WthCassette(CasPosition).BoxLeftNum = SDOCdm.TotNbrPresent
                WthCassette(CasPosition).PurgedNotes = SDOCdm.TotNbrDispensedNotDelivered + SDOCdm.TotNbrDeliveredNotTaken
                WthCassette(CasPosition).BoxState = TranslateBoxState(SDOCdm.CasState, False)
                WthCassette(CasPosition).BoxStateCHN = TranslateBoxState(SDOCdm.CasState, True)
                WthCassette(CasPosition).BoxInitial = SDOCdm.InitialCount
                'Add for new OpInCasStatus.htm
                g_nRetractCount = g_nRetractCount + SDOCdm.TotNbrDeliveredNotTaken
                nNumOfAvailBoxes = nNumOfAvailBoxes + 1
            End If
        End If
    Next i
    
    SDOCdm.CasNbrLogical = 0
    g_nRejectCount = SDOCdm.TotNbrPresent
    
    If SDOCdm.CasState = casstate_cdm_missing Or nNumOfAvailBoxes = 0 Then
        FlushBoxesStatusRetIsPresent = False
    Else
        FlushBoxesStatusRetIsPresent = True
    End If

End Function
'==========================================================================================
'�������� : �õ��豸״̬
'������� ���豸����,�Ƿ���������ʾ״̬��Ϣ
'������� ����
'����ֵ   ��״̬��Ϣ���ַ�����
'���ú��� ��
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Function TranslateDeviceState(ByVal sDeviceName As String, ByVal bChinese As Boolean) As String
On Error Resume Next

    Dim sDeviceSts As String
    Dim sEngDeviceSts As String
    
    sDeviceSts = "δ֪״̬"
    sEngDeviceSts = "Unknown"
    Select Case sDeviceName
    Case "PRJ", "��ˮ��ӡ��"
        Select Case SDOPrj.OperatorType
        Case optype_ok
            If SDOPrj.Available = False Then
                sDeviceSts = "�豸���ɲ���"
                sEngDeviceSts = "Inoperative"
            Else
                sDeviceSts = "����"
                sEngDeviceSts = "OK"
            End If
        Case optype_prr_ink_empty
            sDeviceSts = "ī��"
            sEngDeviceSts = "Ink Empty"
        Case optype_prr_ink_low
            sDeviceSts = "ī��"
            sEngDeviceSts = "Ink Low"
        Case optype_prr_off_line
            sDeviceSts = "�豸�ѻ�"
            sEngDeviceSts = "Offline"
        Case optype_prr_paper_empty
            sDeviceSts = "ֽ��"
            sEngDeviceSts = "Paper Empty"
        Case optype_prr_paper_jammed
            sDeviceSts = "��ֽ"
            sEngDeviceSts = "Paper Jam"
        Case optype_prr_paper_low
            sDeviceSts = "ֽ��"
            sEngDeviceSts = "Paper Low"
        Case optype_prr_retract_full
            sDeviceSts = "����ֽϻ��"
            sEngDeviceSts = "Retract Full"
        Case optype_prr_retract_high
            sDeviceSts = "����ֽϻ��"
            sEngDeviceSts = "Retract High"
        Case Else
            sDeviceSts = "δ֪״̬"
            sEngDeviceSts = "Unknown"
        End Select
    Case "PRR", "ƾ����ӡ��"
        Select Case SDOPrr.OperatorType
        Case optype_ok
            If SDOPrr.Available = False Then
                sDeviceSts = "�豸���ɲ���"
                sEngDeviceSts = "Inoperative"
            Else
                sDeviceSts = "����"
                sEngDeviceSts = "OK"
            End If
        Case optype_prr_ink_empty
            sDeviceSts = "ī��"
            sEngDeviceSts = "Ink Empty"
        Case optype_prr_ink_low
            sDeviceSts = "ī��"
            sEngDeviceSts = "Ink Low"
        Case optype_prr_off_line
            sDeviceSts = "�豸�ѻ�"
            sEngDeviceSts = "Offline"
        Case optype_prr_paper_empty
            sDeviceSts = "ֽ��"
            sEngDeviceSts = "Paper Empty"
        Case optype_prr_paper_jammed
            sDeviceSts = "��ֽ"
            sEngDeviceSts = "Paper Jam"
        Case optype_prr_paper_low
            sDeviceSts = "ֽ��"
            sEngDeviceSts = "Paper Low"
        Case optype_prr_retract_full
            sDeviceSts = "����ֽϻ��"
            sEngDeviceSts = "Retract Full"
        Case optype_prr_retract_high
            sDeviceSts = "����ֽϻ��"
            sEngDeviceSts = "Retract High"
        Case Else
            sDeviceSts = "δ֪״̬"
            sEngDeviceSts = "Unknown"
        End Select
    Case "IDC", "�ſ���д��"
        Select Case SDOIdc.OperatorType
        Case optype_idc_card_jammed
            sDeviceSts = "�ſ�����"
            sEngDeviceSts = "Card Jam"
        Case optype_idc_retract_full
            sDeviceSts = "���տ�Ƭϻ��"
            sEngDeviceSts = "Retract Full"
        Case optype_ok
            sDeviceSts = "����"
            sEngDeviceSts = "OK"
        Case Else
            sDeviceSts = "δ֪״̬"
            sEngDeviceSts = "Unknown"
        End Select
    Case "EDM", "����ģ��"
        Select Case SDOEdm.OperatorType
        Case 1
            sDeviceSts = "δ��ʼ��"
            sEngDeviceSts = "Not Initialized"
        Case 0
            sDeviceSts = "����"
            sEngDeviceSts = "OK"
        Case Else
            sDeviceSts = "δ֪״̬"
            sEngDeviceSts = "Unknown"
        End Select
    
    Case "CDM", "ȡ��ģ��"
        Select Case SDOCdm.OperatorType
        Case optype_cdm_allempty
            sDeviceSts = "���г����"
            sEngDeviceSts = "All Cassettes Empty"
        Case optype_cdm_casinop
            sDeviceSts = "ĳ���仵"
            sEngDeviceSts = "Cassette Inoperative"
        Case optype_cdm_casinvalid
            sDeviceSts = "���䲻����"
            sEngDeviceSts = "Cassette Invalid"
        Case optype_cdm_casnotconfigured
            sDeviceSts = "����δ����"
            sEngDeviceSts = "Cassette UnConfig"
        Case optype_cdm_casnotinstalled
            sDeviceSts = "�г���δװ"             '2005.12.26
            sEngDeviceSts = "Cassette UnInstall"
        Case optype_cdm_device_inop
            sDeviceSts = "�豸���ɲ���"
            sEngDeviceSts = "CDM Inoperative"
        Case optype_cdm_device_offline
            sDeviceSts = "�豸�ѻ�"
            sEngDeviceSts = "Offline"
        Case optype_cdm_dispense_status_unknown
            sDeviceSts = "����״̬δ֪"
            sEngDeviceSts = "Dispense Unknown"
        Case optype_cdm_notesproblem
            sDeviceSts = "����"
            sEngDeviceSts = "Notes Jam"
        Case optype_cdm_rejectcasfull
            sDeviceSts = "�ϳ�����"
            sEngDeviceSts = "RejectBox Full"
        Case optype_cdm_rejectcasnotconfigured
            sDeviceSts = "�ϳ���δ����"
            sEngDeviceSts = "RejectBox UnConfig"
        Case optype_cdm_rejectcasnotinstalled
            sDeviceSts = "�ϳ���δ��װ"
            sEngDeviceSts = "RejectBox UnInstall"
        Case optype_cdm_retractlimitexceeded
            sDeviceSts = "���ճ�����"
            sEngDeviceSts = "Retract Exceeded"
        Case optype_cdm_shutterproblem
            sDeviceSts = "����������"
            sEngDeviceSts = "Shutter Problem"
        Case optype_cdm_somecasslow
            sDeviceSts = "����Ǯ��"
            sEngDeviceSts = "Cassette Low"
        Case optype_ok
            sDeviceSts = "����"
            sEngDeviceSts = "OK"
        Case Else
            sDeviceSts = "δ֪״̬"
            sEngDeviceSts = "Unknown"
        End Select
    End Select
    
    If bChinese Then
        TranslateDeviceState = sDeviceSts
    Else
        TranslateDeviceState = sEngDeviceSts
    End If
End Function
'==========================================================================================
'�������� :�豸��λ
'������� ���豸����
'������� ����
'����ֵ   ���Ƿ�λ�ɹ�������ֵ��
'���ú��� ��PrjOperatorResetATMDev
'�����������������Աѡ��Ҫ�Լ���豸ʱ
'���ߣ�    ����
'����ʱ�� :
'==========================================================================================
Function ResetReleatedDevice(ByVal sDeviceName As String) As Boolean
    Dim sDEVStatus As String

    ResetReleatedDevice = False
    Call PrjOperatorResetATMDev(sDeviceName)
    Select Case sDeviceName
        Case "PRJ", "��ˮ��ӡ��"
            LogError sDeviceName + " FALSE --> Recovering"
            nrc = SDOPrj.DoRecovery
            If nrc = 0 Then
                LogWarning (sDeviceName + " DoRecovery OK")
            Else
                LogError (sDeviceName + " DoRecovery Failed = " + CStr(nrc))
            End If
            If SDOPrj.Available Then
                
                PrintJournal SDOPrj, "SERVICE CLIENT JOURNAL PRINTER TEST", "��ˮ��ӡ���Լ��ӡ����", g_sPrjLanguage

                'nrc = SDOPrj.DoPrintTest("SERVICE CLIENT JOURNAL PRINTER TEST")
            End If
            
        Case "PRR", "ƾ����ӡ��"
            LogError sDeviceName + " FALSE --> Recovering"
            nrc = SDOPrr.DoRecovery
            If nrc = 0 Then
                LogWarning (sDeviceName + " DoRecovery OK")
            Else
                LogError (sDeviceName + " DoRecovery Failed = " + CStr(nrc))
            End If
            If SDOPrr.Available Then
                SDOPrr.Present = True
                g_sPrrRawData = "SERVICE CLIENT RECEIPT PRINTER TEST"
                g_bIsPrrResetTest = True
                nrc = SDOPrr.DoPrintRaw()
                If nrc = 0 Then
                    ResetReleatedDevice = True
                    Exit Function
                End If
            End If
            
        Case "IDC", "�ſ���д��"
            LogError sDeviceName + " FALSE --> Recovering"
            nrc = SDOIdc.DoRecovery
            If nrc = 0 Then
                LogWarning (sDeviceName + " DoRecovery OK")
            Else
                LogError (sDeviceName + " DoRecovery Failed = " + CStr(nrc))
            End If
            
        Case "EDM", "����ģ��"
            If G_nDevicesToUse And DEVICE_EDM Then
                LogError sDeviceName + " FALSE --> Recovering"
                nrc = SDOEdm.DoRecovery
                If nrc = 0 Then
                    LogWarning (sDeviceName + " DoRecovery OK")
                Else
                    LogError (sDeviceName + " DoRecovery Failed = " + CStr(nrc))
                End If
            End If
            
        Case "CDM", "ȡ��ģ��"
            LogError sDeviceName + " FALSE --> Recovering"
            nrc = SDOCdm.DoRecovery
            If nrc = 0 Then
                LogWarning (sDeviceName + " DoRecovery OK")
                nrc = Pcb3dl.DlSetCharRaw("GBLCdmRecoveryNeeded", "Y")   '2005.12.26
            Else
                LogError (sDeviceName + " DoRecovery Failed = " + CStr(nrc))
            End If
            
    End Select
    
    sDEVStatus = TranslateDeviceState(sDeviceName, True)
    nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt2", sDEVStatus)
    currentPage = pageResetATM40
            
End Function

'==========================================================================================
'�������� :��ӡ�豸�Լ���Ϣ
'������� ���豸����
'������� ����
'����ֵ   ����
'���ú��� ����
'�����������ResetReleatedDevice����
'���ߣ�    ����
'����ʱ�� :
'==========================================================================================
Private Sub PrjOperatorResetATMDev(pResetDevName As String)
    Dim theTime      As String
    Dim PrjString    As String
    Dim PrjCHNString As String
    
    theTime = Format(Now(), "YY/MM/DD HH:MM")
    
    PrjString = JourLineSeprator + "       OPERATOR RESET ATM DEVICES" + vbCrLf + _
                "    " + theTime + "      ATM:" + AtmCode + vbCrLf + _
                "       DEVICE[" + pResetDevName + "] was recoverying." + vbCrLf
    
    Select Case pResetDevName
        Case "PRJ"
            pResetDevName = "��ˮ��ӡ��"
        Case "PRR"
            pResetDevName = "ƾ����ӡ��"
        Case "IDC"
            pResetDevName = "�ſ���д��"
        Case "EDM"
            pResetDevName = "����ģ��"
        Case "CDM"
            pResetDevName = "ȡ��ģ��"
    End Select
    
    PrjCHNString = JourLineSeprator + "   �� �� Ա �� �� �� �� �� ��" + vbCrLf + _
                " " + theTime + " ATM�ţ�" + AtmCode + vbCrLf + _
                "  �豸[" + pResetDevName + "] �����Լ�..." + vbCrLf

    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
End Sub

'==========================================================================================
'�������� : У�������Key�ַ����Ƿ���Ч
'������� ����ҪУ���Key�ַ���
'������� ����
'����ֵ   ���Ƿ�Ϊ��ЧKey(����ֵ)
'���ú��� ����
'�����������������Կʱ
'���ߣ�
'����ʱ�� :
'==========================================================================================
Function IsValidKey(sNeedToValid As String) As Boolean
    Dim i As Integer, LenOfStr As Integer, bResult As Boolean
    bResult = True
    LenOfStr = Len(sNeedToValid)
    For i = 1 To LenOfStr
        If Mid(sNeedToValid, i, 1) Like "[0-9,A-F]" Then
        Else
            bResult = False
            Exit For
        End If
    Next
    IsValidKey = bResult
End Function

'==========================================================================================
'�������� : ��ӡ������Կ�Ƿ�ɹ�
'������� ��������Կ�Ƿ�ɹ���־(����ֵ)
'������� ����
'����ֵ   ����
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Private Sub PrjKeyInput(bIsSucceed As Boolean)
    Dim theTime          As String
    Dim PrjString        As String
    Dim PrjCHNString     As String
    Dim sTitle           As String
    Dim sCHNTitle        As String

    theTime = Format(Now(), "YY/MM/DD HH:MM")
    If bIsSucceed Then
        sTitle = "        INSERT MASTER KEY OK!"
        sCHNTitle = "      ��������Կ�ɹ�!"
    Else
        sTitle = "        INSERT MASTER KEY FAILED!"
        sCHNTitle = "      ��������Կʧ�ܣ�"
    End If
    
    PrjString = JourLineSeprator + sTitle + vbCrLf + _
                "    " + theTime + " ATM:" + AtmCode + vbCrLf
                
    PrjCHNString = JourLineSeprator + sCHNTitle + vbCrLf + _
                   "    " + theTime + "   ATM�ţ�" + AtmCode + vbCrLf

    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
End Sub
'==========================================================================================
'�������� : ��ӡ������Կ�Ƿ�ɹ�
'������� ��������Կ�Ƿ�ɹ���־(����ֵ)
'������� ����
'����ֵ   ����
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Function GetCheckValue(DesKey As String, bIsTriDes As Boolean) As String
    Dim sTempResult    As String
    
    SDOEdm.CryptMode = True
    SDOEdm.CryptType = 1
    If bIsTriDes Then
        nrc = SDOEdm.DoCryptDataSw("0000000000000000", Left(DesKey, 16))
        SDOEdm.CryptMode = False
        sTempResult = SDOEdm.CryptResult
        nrc = SDOEdm.DoCryptDataSw(sTempResult, Right(DesKey, 16))
        sTempResult = SDOEdm.CryptResult
        SDOEdm.CryptMode = True
        nrc = SDOEdm.DoCryptDataSw(sTempResult, Left(DesKey, 16))
        GetCheckValue = SDOEdm.CryptResult
    Else
        nrc = SDOEdm.DoCryptDataSw("0000000000000000", DesKey)
        GetCheckValue = SDOEdm.CryptResult
    End If
    
End Function
'==========================================================================================
'�������� :����Key���
'������� ����
'������� ����
'����ֵ   ������Key���Ľ�����ַ�����
'���ú��� ����
'�����������
'���ߣ�
'����ʱ�� :
'==========================================================================================
Function DoXorKeys() As String
    Dim i                   As Integer
    Dim Byte8KeyArr1(7)     As Byte
    Dim byte8KeyArr2(7)     As Byte
    Dim Byte16KeyArr1(15)   As Byte
    Dim byte16KeyArr2(15)   As Byte
    Dim sXorResult          As String
    
    If G_bTrides Then
        Call StrToBin(GLarrMasKeys(0), Byte16KeyArr1, 32)
        Call StrToBin(GLarrMasKeys(1), byte16KeyArr2, 32)
        For i = 0 To 15
            Byte16KeyArr1(i) = Byte16KeyArr1(i) Xor byte16KeyArr2(i)
        Next
        Call StrToBin(GLarrMasKeys(2), byte16KeyArr2, 32)
        For i = 0 To 15
            Byte16KeyArr1(i) = Byte16KeyArr1(i) Xor byte16KeyArr2(i)
        Next
        Call BinToStr(Byte16KeyArr1, sXorResult, 16)
    Else
        Call StrToBin(GLarrMasKeys(0), Byte8KeyArr1, 16)
        Call StrToBin(GLarrMasKeys(1), byte8KeyArr2, 16)
        For i = 0 To 7
            Byte8KeyArr1(i) = Byte8KeyArr1(i) Xor byte8KeyArr2(i)
        Next
        Call StrToBin(GLarrMasKeys(2), byte8KeyArr2, 16)
        For i = 0 To 7
            Byte8KeyArr1(i) = Byte8KeyArr1(i) Xor byte8KeyArr2(i)
        Next
        Call BinToStr(Byte8KeyArr1, sXorResult, 8)
    End If
    
    DoXorKeys = sXorResult
End Function

Public Sub StrToBin(ByVal inString As String, ByRef bOutArray() As Byte, LenOfStr As Integer)
    Dim strTwo     As String
    Dim i          As Integer
    Dim j          As Integer

    j = 0
    For i = 1 To LenOfStr Step 2
        strTwo = Mid(inString, i, 2)
        bOutArray(j) = Val("&H" + strTwo)
        j = j + 1
    Next

End Sub
Public Sub BinToStr(ByRef InPar() As Byte, ByRef OutPar As String, LenOfBin As Integer)
    Dim i          As Integer
    Dim strNum     As String
    
    For i = 0 To LenOfBin - 1
        strNum = Hex(InPar(i))
        If Len(strNum) < 2 Then
            strNum = "0" + strNum
        End If
        OutPar = OutPar + strNum
    Next i
End Sub

'==========================================================================================
'�������� :�������ݿ�õ��쳣����ȡ����������ӡ��ˮ��׼��������ӡ����ʾ����
'������� ����
'������� ����
'����ֵ  ��0 ���ݿ�û�м�¼��1 ���ݿ��м�¼
'���ú��� ����
'�����������
'����   :������
'����ʱ�� : 2005.6
'�޸ļ�¼��
'<ʱ��>��[2005.12.20]
'<�޸���>��������
'       ȡ���쳣���׼���GB withend 99 else ȡ��������
'==========================================================================================
Private Function PrepareDisplayRecords(ByRef TransType As String) As Integer
    Dim i                           As Integer
    Dim j                           As Integer
    Dim NumberOfNotKeep             As Integer
    Dim NumberOfNeedcheck           As Integer
    Dim NumberOfCashjam             As Integer
    Dim iNumberOfReception1         As Integer
    Dim iNumberOfReception2         As Integer
    Dim NumberOfReversalOK          As Integer
    Dim NumberOfTakeTimeout         As Integer
    
    PrepareDisplayRecords = 0
    TOTPrjString = " "
    
    DataWTH.RecordSource = "Select * From CWDLOG Where TransType = '" & TransType & "' and KeepAccountFlag <>'Y' order by TransDate"
    DataWTH.Refresh
    g_iTotalNumberOfDisplay = DataWTH.Recordset.RecordCount
    
    '���ȡ�ʱδ��(PT),ȡ�����(GB) 2005��12��20
    If TransType = "CWD" Then
        DataWTH.RecordSource = "Select * From CWDLOG Where (TransType = 'CWD' and AccountErrorReason ='PT') or (TransType = 'CWD'  and AccountErrorReason ='GB')"
        DataWTH.Refresh
        g_iTotalNumberOfDisplay = g_iTotalNumberOfDisplay + DataWTH.Recordset.RecordCount
    End If
    
    If g_iTotalNumberOfDisplay = 0 Then
        Exit Function
    End If
    
    ReDim AssortLog(1 To g_iTotalNumberOfDisplay) As AssortLogType
    NumberOfNotKeep = 0
    NumberOfNeedcheck = 0
    NumberOfCashjam = 0
    NumberOfReversalOK = 0
    NumberOfTakeTimeout = 0
    
    TOTPrjString = "�� �� ȡ �� ͳ �� ��" + vbCrLf
    For i = 1 To 5
        'δ����ͳ��
        If i = 1 Then
            DataWTH.RecordSource = "Select * From CWDLOG Where TransType = '" & TransType & "'and KeepAccountFlag ='N'"
            DataWTH.Refresh
            NumberOfNotKeep = DataWTH.Recordset.RecordCount
            If NumberOfNotKeep <> 0 Then
                TOTPrjString = "δ �� �� " + CStr(NumberOfNotKeep) + " ��" + vbCrLf + "����     �ʺ�              ���" + "��ˮ��" + vbCrLf
            End If
        '�����������ͳ��
        ElseIf i = 2 Then
            DataWTH.RecordSource = "Select * From CWDLOG Where TransType = '" & TransType & "' and KeepAccountFlag ='U'"
            DataWTH.Refresh
            NumberOfNeedcheck = DataWTH.Recordset.RecordCount
            If NumberOfNeedcheck <> 0 Then
                TOTPrjString = TOTPrjString + vbCrLf + "�� Ҫ �� �� �� �� " + CStr(NumberOfNeedcheck) + " ��" + vbCrLf + "����     �ʺ�              ���" + vbCrLf + "��ˮ��  ԭ ��" + vbCrLf
            End If
        ElseIf i = 3 Then
            '�ѳ������ɹ�
            DataWTH.RecordSource = "Select * From CWDLOG Where TransType = '" & TransType & "' and KeepAccountFlag ='R'"
            DataWTH.Refresh
            NumberOfReversalOK = DataWTH.Recordset.RecordCount
            If NumberOfReversalOK <> 0 Then
                TOTPrjString = TOTPrjString + vbCrLf + "�� �� �� �� " + CStr(NumberOfReversalOK) + " ��" + vbCrLf + "����     �ʺ�              ���" + vbCrLf + "��ˮ��  ����ԭ��" + vbCrLf
            End If
        ElseIf i = 4 Then
           'ȡ�ʱδ��(PT)
                DataWTH.RecordSource = "Select * From CWDLOG Where (TransType = 'CWD' and AccountErrorReason ='PT') "
                DataWTH.Refresh
                NumberOfTakeTimeout = DataWTH.Recordset.RecordCount
                If NumberOfTakeTimeout <> 0 Then
                    TOTPrjString = TOTPrjString + vbCrLf + "ȡ �� �� ʱ δ ȡ " + CStr(NumberOfTakeTimeout) + " ��" + vbCrLf + "����        �ʺ�                  ���" + vbCrLf + "��ˮ��" + vbCrLf
                End If
        Else
           ' ȡ�����(GB) 2005��12.20
                DataWTH.RecordSource = "Select * From CWDLOG Where  (TransType = 'CWD'  and AccountErrorReason ='GB') or (TransType = 'CWD'  and AccountErrorReason ='PF')"
                DataWTH.Refresh
                NumberOfCashjam = DataWTH.Recordset.RecordCount
                If NumberOfCashjam <> 0 Then
                    TOTPrjString = TOTPrjString + vbCrLf + "ȡ �� �� �� " + CStr(NumberOfCashjam) + " ��" + vbCrLf + "����     �ʺ�              ���" + vbCrLf + "��ˮ�� ԭ��" + vbCrLf
                End If
                
        End If
        
        If DataWTH.Recordset.RecordCount <> 0 Then
            PrepareDisplayRecords = 1
            
            If i = 1 Then
                iNumberOfReception1 = 1
                iNumberOfReception2 = NumberOfNotKeep
            ElseIf i = 2 Then
                iNumberOfReception1 = NumberOfNotKeep + 1
                iNumberOfReception2 = NumberOfNotKeep + NumberOfNeedcheck
            ElseIf i = 3 Then
                iNumberOfReception1 = NumberOfNotKeep + NumberOfNeedcheck + 1
                iNumberOfReception2 = NumberOfNotKeep + NumberOfNeedcheck + NumberOfReversalOK
            ElseIf i = 4 Then
                iNumberOfReception1 = NumberOfNotKeep + NumberOfNeedcheck + NumberOfReversalOK + 1
                iNumberOfReception2 = NumberOfNotKeep + NumberOfNeedcheck + NumberOfReversalOK + NumberOfTakeTimeout
            Else
                iNumberOfReception1 = NumberOfNotKeep + NumberOfNeedcheck + NumberOfReversalOK + NumberOfTakeTimeout + 1
                iNumberOfReception2 = NumberOfNotKeep + NumberOfNeedcheck + NumberOfCashjam + NumberOfReversalOK + NumberOfTakeTimeout
            End If
            
            For j = iNumberOfReception1 To iNumberOfReception2
                TxtTransDate.DataField = "TransType"
                AssortLog(j).AssortTransType = TxtTransDate.Text
                
                TxtTransDate.DataChanged = True
                TxtTransDate.DataField = "TransDate"
                AssortLog(j).AssortDate = TxtTransDate.Text
                
                TxtTransDate.DataChanged = True
                TxtTransDate.DataField = "TransCardType"
                AssortLog(j).AssortCardType = TxtTransDate.Text
                
                TxtTransDate.DataChanged = True
                TxtTransDate.DataField = "TransSerial"
                AssortLog(j).AssortSerial = TxtTransDate.Text
                
                TxtTransDate.DataChanged = True
                TxtTransDate.DataField = "TransAmount"
                AssortLog(j).AssortAmount = TxtTransDate.Text
                
                TxtTransDate.DataChanged = True
                TxtTransDate.DataField = "TransAccNo"
                AssortLog(j).AssortAccNo = TxtTransDate.Text
                
                TxtTransDate.DataChanged = True
                TxtTransDate.DataField = "KeepAccountFlag"
                AssortLog(j).AssortKeepAccFlag = TxtTransDate.Text
                
                TxtTransDate.DataChanged = True
                TxtTransDate.DataField = "AccountErrorReason"
                AssortLog(j).AssortCashinResult = TxtTransDate.Text
                
                TxtTransDate.DataChanged = True
                TxtTransDate.DataField = "HostRejectCode"
                AssortLog(j).AssosrtHostReject = TxtTransDate.Text
                
                DataWTH.Recordset.MoveNext
                
                '��ˮ��ӡ
                TOTPrjString = TOTPrjString + AssortLog(j).AssortDate + " " + Format(AssortLog(j).AssortAccNo, "@@@@@@@@@@@@@@@@@@@@") + _
                          " " + CStr(AssortLog(j).AssortAmount) + vbCrLf + AssortLog(j).AssortSerial
            
                If i = 2 Or i = 3 Then
                    TOTPrjString = TOTPrjString + TranslateCWDReason(AssortLog(j).AssortCashinResult) + vbCrLf
                End If
                If i = 5 Then
                    TOTPrjString = TOTPrjString + AssortLog(j).AssortCashinResult + vbCrLf
                End If
            
                If AssortLog(j).AssortTransType = "CWD" Then
                    AssortLog(j).AssortCashinResult = TranslateCWDReason(AssortLog(j).AssortCashinResult)
                End If
                
                Select Case AssortLog(j).AssortCardType
                Case "99"
                    AssortLog(j).AssortCardType = "����"
                Case "98", "97", "96"                   '2005.12.21
                    AssortLog(j).AssortCardType = "����"
                Case "01"
                    AssortLog(j).AssortCardType = "����"
                Case Else
                    AssortLog(j).AssortCardType = "���"
                End Select
                
                Select Case AssortLog(j).AssortKeepAccFlag
                Case "N"
                    AssortLog(j).AssortKeepAccFlag = "δ����"
                Case "U"
                    AssortLog(j).AssortKeepAccFlag = "����  "
                Case "Y"
                    AssortLog(j).AssortKeepAccFlag = "�Ѽ���"
                Case "R"
                    AssortLog(j).AssortKeepAccFlag = "�����ɹ�"
                End Select
            Next j
        End If                  'endif DataWTH.Recordset.RecordCount<>0
    Next i
    
End Function


'==========================================================================================
'�������� ������Ļ����ʾ��ϸ��¼
'������� ����������
'������� ����
'����ֵ  ����
'���ú��� ����
'����������� TimerAction_Timer ����ʾ�쳣�����ϸ���쳣ȡ����ϸ�����н�����ϸ
'����   :������
'����ʱ�� : 2005.7
'==========================================================================================
Private Function DisplayRecords(ByVal TransType As String, ByVal TransState As String)
    Dim sDisplayStr                 As String
    Dim sFKList                     As String
    Dim i                           As Integer
    Dim j                           As Integer
    Dim Display                     As Boolean
    Dim iTotPages                   As Integer
    Dim iDivider                    As Integer
    Dim iReminder                   As Byte
    
    g_nLogLastPos = 1
    g_nLogCurPos = 1
    Display = True
    sFKList = ""
    
    iDivider = g_iTotalNumberOfDisplay \ 8
    iReminder = g_iTotalNumberOfDisplay Mod 8
    If iReminder <> 0 Then
        iTotPages = iDivider + 1
    Else
        iTotPages = iDivider
    End If
    
    nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", CStr(iTotPages))
    Do While (Display)
        
        For i = 1 + (g_nLogCurPos - 1) * 8 To 8 * g_nLogCurPos
            If i <= g_iTotalNumberOfDisplay Then
                If AssortLog(i).AssortTransType = "CDP" Then
                    AssortLog(i).AssortTransType = "���"
                ElseIf AssortLog(i).AssortTransType = "CWD" Then
                    AssortLog(i).AssortTransType = "ȡ��"
                End If
                If i = 1 + (g_nLogCurPos - 1) * 8 Then
                    sDisplayStr = AssortLog(i).AssortTransType + "|" + Left(AssortLog(i).AssortDate, 2) + "��" + Mid(AssortLog(i).AssortDate, 3, 2) + "��" + Mid(AssortLog(i).AssortDate, 5, 2) + ":" + Right(AssortLog(i).AssortDate, 2) + "|" + _
                    AssortLog(i).AssortAccNo + "|" + AssortLog(i).AssortCardType + "|" + _
                    CStr(AssortLog(i).AssortAmount) + "|" + AssortLog(i).AssortKeepAccFlag + "|" + _
                    AssortLog(i).AssortCashinResult + "|" + AssortLog(i).AssortSerial + "|" + AssortLog(i).AssosrtHostReject + "|"
                Else
                    sDisplayStr = sDisplayStr + AssortLog(i).AssortTransType + "|" + Left(AssortLog(i).AssortDate, 2) + "��" + Mid(AssortLog(i).AssortDate, 3, 2) + "��" + Mid(AssortLog(i).AssortDate, 5, 2) + ":" + Right(AssortLog(i).AssortDate, 2) + "|" + _
                    AssortLog(i).AssortAccNo + "|" + AssortLog(i).AssortCardType + "|" + _
                    CStr(AssortLog(i).AssortAmount) + "|" + AssortLog(i).AssortKeepAccFlag + "|" + _
                    AssortLog(i).AssortCashinResult + "|" + AssortLog(i).AssortSerial + "|" + AssortLog(i).AssosrtHostReject + "|"
                End If
                If i = g_iTotalNumberOfDisplay Then
                     sFKList = "@PGDN,"
                End If
            Else
                sFKList = "@PGDN,"
                For j = g_iTotalNumberOfDisplay + 1 To 8 * g_nLogCurPos
                    sDisplayStr = sDisplayStr + "&nbsp|&nbsp|&nbsp|&nbsp|&nbsp|&nbsp|&nbsp|&nbsp|&nbsp|"
                Next
                Exit For
            End If
        Next
        
        If g_nLogCurPos = 1 Then
            sFKList = sFKList + "@PGUP,"
        End If
        nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt2", CStr(g_nLogCurPos))
        nrc = Pcb3dl.DlSetCharRaw("MaintHtmlFkeyList", sFKList)
        nrc = Pcb3dl.DlSetCharRaw("OptevaCasStatus", sDisplayStr)
        nrc = ShowOperScreenMaint("Operator", "OpDispCDP10")
         If BrowserMaint.SubStData = "@PGUP" Then
        '��ʾ��һҳ
            g_nLogCurPos = g_nLogLastPos
            If g_nLogLastPos > 0 Then
            g_nLogLastPos = g_nLogLastPos - 1
            End If
            sFKList = ""
            nrc = Pcb3dl.DlSetCharRaw("MaintHtmlFkeyList", "")
        ElseIf BrowserMaint.SubStData = "@PGDN" Then
        '��ʾ��һҳ
            g_nLogLastPos = g_nLogCurPos
            g_nLogCurPos = g_nLogCurPos + 1
            sFKList = ""
            nrc = Pcb3dl.DlSetCharRaw("MaintHtmlFkeyList", "")
        Else
            currentPage = pageDispCDP30
            Display = False
        End If
    Loop

End Function

'==========================================================================================
'�������� :��ȡ���쳣ԭ����Ӣ�ķ�������ģ�������ʾ��������ӡ
'������� ����
'������� ����
'����ֵ  �����Ľ���
'���ú��� ����
'�����������PrepareReceptionCashinRecords����
'����   :������
'����ʱ�� : 2005.6
'==========================================================================================
Private Function TranslateCWDReason(ByRef sEnglishReason As String) As String

    Select Case sEnglishReason
    Case "DF"
        TranslateCWDReason = "�䳮ʧ�ܳ���"
    Case "GB"
        TranslateCWDReason = "���齻��"
    Case "ST"
        TranslateCWDReason = "����ԭ�����"
    Case "PF"
        TranslateCWDReason = "���齻��"
    Case "PT"
        TranslateCWDReason = "��ʱδ��"
    Case "CS"
        TranslateCWDReason = "����ʧ��"
    Case "CR"
        TranslateCWDReason = "ͨѶʧ��"
    Case "CU"
        TranslateCWDReason = "��������"
    Case "OK"
        TranslateCWDReason = "��������"
    Case "CM"
        TranslateCWDReason = "�����з���"
    Case Else
        TranslateCWDReason = "δ֪ԭ��"
    End Select
End Function
Private Function GetSuperOprFunctionPage(OprCommand As String) As pageType
    gSelectOprCommand = ""
    Select Case OprCommand
        Case "01"
            GetSuperOprFunctionPage = pageSuperSetTerminalLuno
        Case "03"
            GetSuperOprFunctionPage = pageSuperSetBankCode
'        Case "05"
'             GetSuperOprFunctionPage = pageOpKeyInput10
'             Pcb3dl.DlSetCharRaw "HtmlWork13", ""
        Case "09"
             GetSuperOprFunctionPage = pageOpChgPwd10
        '��������ͨѶ��Կ�Ĺ���
        Case "07"
             GetSuperOprFunctionPage = pageOpResetTransKey
        Case "00"
             GetSuperOprFunctionPage = pageFunChoice
        Case Else
             nrc = Pcb3dl.DlSetCharRaw("TTU01", OprCommand)
             nrc = ShowOperScreenMaint("Operator", "OpNoFunAvail")
             GetSuperOprFunctionPage = pageSuperFunctionChoice
    End Select
End Function

'==========================================================================================
'��������: ����������ͨѶ
'������������ױ��ı�־������ƽʱ������������
'�����������
'����ֵ ����
'���ú�������
'�������� ��������07 ���رջ������ʱ���ã�����11��װ���ɹ������
'˵��          ������TTI, ����ƽʱ������������ ACP
'                װ��  RWT , ����ƽʱ������������ADP
'                ��մ���RDT , ����ƽʱ������������AEP
'               ���� AAP - ����һ�£�AVP - ��Կ��ͬ��
'�� ��         ��������
'�� �� ʱ ��   : 2005-6-28
'==========================================================================================
Sub CommunicationSubFunction(ByRef TransFlag, ByRef HostReturn)
    Dim iCassNumber                 As Integer
    Dim nCount                      As Integer
    Dim TransCode                   As String
    Dim i                           As Integer
    Dim iRc                         As Integer
    Dim CasDenomArray(1 To 5)       As Integer
    Dim CasInitCountArray(1 To 5)   As Integer
    Dim FindSameDenom               As Boolean
    Dim PrjString               As String
    Dim PrjCHNString            As String
    
    Select Case TransFlag
    Case "RWT" '���ͼӳ�����
        For i = 1 To 4
            If (WthCassette(i).CasLogicalID <> 0) Then
                SDOCdm.CasNbrLogical = WthCassette(i).CasLogicalID
                If SDOCdm.CasState <= 4 Then
                    CasDenomArray(i) = SDOCdm.CasDenomination
                    CasInitCountArray(i) = SDOCdm.InitialCount
                    nrc = S3ELineOut.SetData("CasDemo" & CStr(i), _
                            Format(CasDenomArray(i), "0000"))
                            
                    nrc = S3ELineOut.SetData("DenoRefill" & CStr(i), _
                            Format(LastCashFilled(i), "0000"))
                            
                    nrc = S3ELineOut.SetData("CasPresent" & CStr(i), _
                            Format(LastCashPresent(i), "0000"))
                        
                    nrc = S3ELineOut.SetData("RepliDeno" & CStr(i), _
                            Format(SDOCdm.InitialCount, "0000"))
                            
                Else
                    nrc = S3ELineOut.SetData("CasDemo" & CStr(i), "0000")
                    nrc = S3ELineOut.SetData("DenoRefill" & CStr(i), "0000")
                End If
            Else
                nrc = S3ELineOut.SetData("CasDemo" & CStr(i), "0000")
                nrc = S3ELineOut.SetData("DenoRefill" & CStr(i), "0000")
            End If
        Next i

        '������峮���������� 2005.11.29
        If (WthCassette(5).CasLogicalID <> 0) Then
            SDOCdm.CasNbrLogical = WthCassette(5).CasLogicalID
            If SDOCdm.CasState <= 4 Then
                CasDenomArray(5) = SDOCdm.CasDenomination
                
                FindSameDenom = False
                For i = 1 To 4
                    If CasDenomArray(5) = CasDenomArray(i) Then
                        nrc = S3ELineOut.SetData("DenoRefill" & CStr(i), _
                            Format(CasInitCountArray(i) + SDOCdm.InitialCount, "0000"))
                        FindSameDenom = True
                    End If
                Next
                
                If (Not FindSameDenom) Then
                    nrc = S3ELineOut.SetData("DenoRefill4", Format(SDOCdm.InitialCount, "0000"))
                    For i = 1 To 3
                        If CasDenomArray(4) = CasDenomArray(i) Then
                            nrc = S3ELineOut.SetData("DenoRefill" & CStr(i), _
                                Format(CasInitCountArray(i) + CasInitCountArray(4), "0000"))
                        End If
                    Next
                End If
            End If
        End If
        

    Case "RTT" '������ת�ʽ���
        
        iRc = S3ELineOut.SetData("NoOfRMBTfr", Format(Pcb3dl.DlGetInt("TotTfrOutNum"), "0000"))
        iRc = S3ELineOut.SetData("AmtOfRMBTfr", Format(CLng(Pcb3dl.DlGetDouble("TotTfrOutAmount")) * 100, "000000000"))
        
    Case Else '���Ͷ��ʽ���
        SDOCdm.DataCriteria = 1        ' Query by logical number
        For nCount = 1 To 4
            LastCashFilled(nCount) = 0
            LastCashPresent(nCount) = 0
            If WthCassette(nCount).CasLogicalID <> 0 Then
                SDOCdm.CasNbrLogical = WthCassette(nCount).CasLogicalID
                
                iRc = S3ELineOut.SetData("CassDeno" & CStr(nCount), Format(SDOCdm.CasDenomination, "0000"))
                
                iRc = S3ELineOut.SetData("DenoRefill" & CStr(nCount), _
                    Format(SDOCdm.InitialCount, "0000"))
                 LastCashFilled(nCount) = SDOCdm.InitialCount
                 
                iCassNumber = SDOCdm.TotNbrDelivered
                If iCassNumber < 0 Then
                    iCassNumber = 0
                End If
                iRc = S3ELineOut.SetData("CasPresent" & CStr(nCount), _
                    Format(iCassNumber, "0000"))
                 LastCashPresent(nCount) = iCassNumber
                    
                iCassNumber = SDOCdm.TotNbrDelivered + SDOCdm.TotNbrDispensedNotDelivered
                If iCassNumber < 0 Then
                    iCassNumber = 0
                End If
                iRc = S3ELineOut.SetData("CasPurge" & CStr(nCount), _
                    Format(iCassNumber, "0000"))
                    
                iCassNumber = SDOCdm.TotNbrDispensedNotDelivered
                iRc = S3ELineOut.SetData("CasRej" & CStr(nCount), _
                    Format(iCassNumber, "0000"))
            End If
        Next
        
        LastWithDrawNumber = Pcb3dl.DlGetInt("TotWithdrawNum")
        LastTfrNumber = Pcb3dl.DlGetInt("TotTfrOutNum")
        iRc = S3ELineOut.SetData("NoOfRMBWth", Format(Pcb3dl.DlGetInt("TotWithdrawNum"), "0000"))
        iRc = S3ELineOut.SetData("NoOfRMBTfr", Format(Pcb3dl.DlGetInt("TotTfrOutNum"), "0000"))
        If CLng(Pcb3dl.DlGetDouble("TotTfrOutAmount")) > 0 Then
            iRc = S3ELineOut.SetData("AmtOfRMBTfr", Format(CLng(Pcb3dl.DlGetDouble("TotTfrOutAmount")) * 100, "000000000"))
        Else
            iRc = S3ELineOut.SetData("AmtOfRMBTfr", "000000000")
        End If
        
                
        iRc = S3ELineOut.SetData("TotCapCardNum", _
                Format(Pcb3dl.DlGetInt("TotCapCardNum"), "0000"))
    End Select            'end of select TransFlag

    iRc = S3ELineOut.DoSend(TransFlag, 0)
    If iRc <> 0 Then
        LogError "Send " & TransFlag & " failed, " & CStr(iRc)
        Call SendExceptionMessage(S3ELineOut, Pcb3dl, "64")
        Select Case TransFlag
        Case "RWT"
            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "���ͼӳ�����ʱͨѶʧ��")
            PrjString = "RWT send error"
            PrjCHNString = "���ͼӳ�����ʱͨѶʧ��"
        Case "RTT"
            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "������ת�˽���ʱͨѶʧ��")
            PrjCHNString = "������ת�˽���ʱͨѶʧ��"
            PrjString = "RTT send error"
        Case Else
            nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "���Ͷ��˽���ʱͨѶʧ��")
            PrjCHNString = "���Ͷ��˽���ʱͨѶʧ��"
            PrjString = "TTI send error"
        End Select

    
    Else
        iRc = S3ELineOut.DoReceive
        If iRc = 0 Then
            TransCode = Pcb3dl.DlGetCharRaw("HostTransCode")
            If TransCode = HostReturn Then
                Select Case TransFlag
                Case "RWT"
                    nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "�������ܼӳ�����")
                    PrjString = "Host Accept RWT"
                    PrjCHNString = "�������ܼӳ�����"
                Case "RTT"
                    nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "����������ת�˽���")
                    PrjString = "Host Accept RTT"
                    PrjCHNString = "����������ת�˽���"
                Case Else
                    PrjString = "Host Accept TTI"
                    PrjCHNString = "�������ܶ��˽���"
                    nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "�������ܶ��˽���")
                End Select
            Else
                LogError "Received unknown TransCode, " & TransCode
                Select Case TransFlag
                Case "RWT"
                    If TransCode = "ADP" Then
                        nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "ͳ��ֵ��������ƥ��")
                        PrjString = "RWT Host not match "
                        PrjCHNString = "ͳ��ֵ��������ƥ��"
                    Else
                        nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "�����ܾ��ӳ�����")
                        PrjString = "RWT Host reject"
                        PrjCHNString = "�����ܾ��ӳ�����"
                    End If
                Case "RTT"
                    nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "�����ܾ���ת�˽���")
                    PrjString = "RTT Host reject"
                    PrjCHNString = "�����ܾ���ת�˽���"
                Case Else
                    If TransCode = "ACP" Then
                        PrjString = "TTI Host not match"
                        PrjCHNString = "ͳ��ֵ��������ƥ��"
                        nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "ͳ��ֵ��������ƥ��")
                    Else
                        PrjString = "TTI Host reject"
                        PrjCHNString = "�����ܾ����˽���"
                        nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "�����ܾ����˽���")
                    End If
                    
                End Select
                
            End If
        Else
                Select Case TransFlag
                Case "RWT"
                    PrjString = "RWT receive error"
                    PrjCHNString = "���ռӳ�����ʧ��"
                    nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "���ռӳ�����ʧ��")
                Case "RTT"
                    PrjString = "RTT receive error"
                    PrjCHNString = "������ת�˽���ʧ��"
                    nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "������ת�˽���ʧ��")
                Case Else
                    PrjString = "TTI receive error"
                    PrjCHNString = "���ն��˽���ʧ��"
                    nrc = Pcb3dl.DlSetCharRaw("HtmlPrompt1", "���ն��˽���ʧ��")
                End Select
            
            Call SendExceptionMessage(S3ELineOut, Pcb3dl, "64")
            LogError "DoReceive" & TransFlag & "failed, " & CStr(iRc)
        End If
    End If
    
   PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    
End Sub

'==========================================================================================
'��������: ������Ա�еĿ��ػ�����ڽ��кϲ�
'�����������
'�����������
'����ֵ ����
'���ú�������
'�������� �����ڼӳ�ǰ������п��ػ������
'˵��          ��
'�� ��         ����ΰʤ
'�� �� ʱ ��   : 2005-10-21
'==========================================================================================
Private Sub CloseAndOpenPeriod()
            Dim hEvent              As Long
            '�رջ������
            If Pcb3dl.DlGetCharRaw("CWDCrimePossible") = "O" Then
                nrc = Pcb3dl.DlSetCharRaw("CWDCrimePossible", "N")
                Pcb3dl.DlSetCharRaw "GBLDoRecovery", "N"
            End If
            Call PrjTotal("CLOSE PERIOD")
            
            'nrc = Pcb3dl.DlSetCharRaw("GBLPeriodStatus", "C")
            
            nrc = Pcb3dl.DlSetCharRaw("TotPeriodCloseTime", Format(Now(), "YYYY/MM/DD HH:MM:SS"))
            Call CommunicationSubFunction("TTI", "AAP")
            
            'AnRchive log files
           
            hEvent = OpenEvent(EVENT_MODIFY_STATE, 0, "S3EDoArchive")
            If hEvent <> 0 Then
                SetEvent hEvent
                CloseHandle hEvent
            End If
            '�򿪻������
            Call ClearTotal
            nrc = Pcb3dl.DlSetCharRaw("GBLPeriodStatus", "O")
             Call PrjOpenPeriod
           sGLtheTime = Format(Now(), "YYYY/MM/DD HH:MM:SS")
            nrc = Pcb3dl.DlSetCharRaw("TotPeriodOpenTime", sGLtheTime)


End Sub
Private Sub CalPageNum()
   Dim Position1 As Integer, Position2 As Integer
   Dim i As Integer
   
    Position1 = 1
    Position2 = 1
    i = 0
    While (Position2 <> 0)
     Position2 = InStr(Position1, g_sPrrRawData, vbCrLf)
     Position1 = Position2 + Len(vbCrLf)
     i = i + 1
     Wend
     If i Mod PrrLineNumber = 0 Then
        PrrTOTPrintPageNumber = i \ PrrLineNumber
     Else
        PrrTOTPrintPageNumber = i \ PrrLineNumber + 1
     End If
     PrrLeftPrintPageNumber = PrrTOTPrintPageNumber
End Sub
Private Sub PrrTotal()
   Dim Position1 As Integer, Position2 As Integer
   Dim temp_str As String
   Dim i As Integer
   
    For i = 1 To 20
       nrc = Pcb3dl.DlSetCharRaw("PrrRow" & CStr(i), " ")
    Next
    
    If PrrLeftPrintPageNumber <> PrrTOTPrintPageNumber Then
        Position1 = PrrPrintPosition
    Else
        Position1 = 1
    End If
    Position2 = 1
    i = 1
    While (Position2 <> 0 And i <= PrrLineNumber)
        Position2 = InStr(Position1, g_sPrrRawData, vbCrLf)
        If Position2 <> 0 Then
            temp_str = Mid(g_sPrrRawData, Position1, Position2 - Position1)
            Position1 = Position2 + Len(vbCrLf)
        Else
            temp_str = Right(g_sPrrRawData, Len(g_sPrrRawData) + 1 - Position1)
        End If
            nrc = Pcb3dl.DlSetCharRaw("PrrRow" & CStr(i), temp_str)
            i = i + 1
     Wend
   
     If PrrLeftPrintPageNumber > 1 Then
        PrrPrintPosition = Position1
     End If
     nrc = SDOPrr.DoPrintForm("TOTPrr")
     'If nRc = 0 Then
     ' PrrLeftPrintPageNumber = PrrLeftPrintPageNumber - 1
     ' nRc = ShowScreen("Operator", "OpPrrPrintTOT30", pagePrrPrintTOT30)
     'Else
     ' nRc = ShowScreen("Operator", "OpPrrPrintTOT40", pagePrrPrintTOT40)
     'End If
End Sub
'==========================================================================================
'�����Ĺ��� :��ӡCutOffͳ��ֵ
'������� :��
'������� : ��
'����ֵ   :��
'���ú��� :��
'���������  ��
'����       ������
'����ʱ��   :2005.12.14
'==========================================================================================
Sub PrintCutOffData()
    Dim CutOffIni                   As String
    Dim CutOffWithNum               As String
    Dim CutOffWithAmount            As String
    Dim CutOffTfrNum                As String
    Dim CutOffTfrAmount             As String
    Dim PrjString                   As String
    Dim PrjCHNString                As String
    
    CutOffIni = "c:\ATMWosa\Ini\CutOff.ini"
    CutOffWithNum = GetIniS(CutOffIni, "Backup", "WithdrawNumber", "0")
    CutOffWithAmount = GetIniS(CutOffIni, "Backup", "WithdrawAmount", "0")
    CutOffTfrNum = GetIniS(CutOffIni, "Backup", "TfrNumber", "0")
    CutOffTfrAmount = GetIniS(CutOffIni, "Backup", "TfrAmount", "0")
    
    PrjString = "==============================" + vbCrLf + _
            "  Last Working Day ATMP Totals" + vbCrLf + _
            "  Type         Count    Amount " + vbCrLf + _
            "  Withdrawals  " + CutOffWithNum + "  " + CutOffWithAmount + vbCrLf + _
            "  Transfer     " + CutOffTfrNum + "  " + CutOffTfrAmount
    PrjCHNString = "==============================" + vbCrLf + _
                "  ��һ��P��ͳ��ֵ" + vbCrLf + _
                "  ����         ����         ��� " + vbCrLf + _
                "  ȡ��         " + CutOffWithNum + "          " + CutOffWithAmount + vbCrLf + _
                "  ת��         " + CutOffTfrNum + "          " + CutOffTfrAmount
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    
End Sub



