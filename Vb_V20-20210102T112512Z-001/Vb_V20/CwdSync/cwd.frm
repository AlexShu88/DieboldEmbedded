VERSION 5.00
Object = "{B2110643-3E81-11D3-8ACC-00C04FF20A5D}#1.2#0"; "TransProv.dll"
Object = "{5C094E41-67D2-11D0-AC6B-0020AFBDD1D4}#1.0#0"; "SDOCdm.ocx"
Object = "{192DFCF0-F664-11D3-8BD4-00C04FF20A5D}#1.1#0"; "AdvBrowser.ocx"
Object = "{BD8177C0-832C-11CF-BF42-0020AF7093F9}#1.0#0"; "SDOIdc.ocx"
Object = "{EACE4ED6-6930-11D0-AC6C-0020AFBDD1D4}#1.0#0"; "SDOFep.ocx"
Object = "{9C37E835-6A58-11D1-80C0-0020AF7093F9}#1.2#0"; "Dl.ocx"
Object = "{6C4DD4AB-27D5-11D3-96C4-000000000000}#1.0#0"; "S3ELineOutTcp.ocx"
Object = "{E64F71A6-E705-4151-9895-5138B7D67F3A}#1.0#0"; "CHPrj.ocx"
Begin VB.Form Cwd 
   Caption         =   "Cwd"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5715
   Icon            =   "cwd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin CHPRJLib.CHPrj SDOPrj 
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   1
   End
   Begin SDOFepLibCtl.SDOFep SDOFep 
      Height          =   495
      Left            =   240
      OleObjectBlob   =   "cwd.frx":0E42
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin SDOIdcLibCtl.SDOIdc SDOIdc 
      Height          =   735
      Left            =   1680
      OleObjectBlob   =   "cwd.frx":0E6C
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox TxtTransDate 
      DataSource      =   "DataTot"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Text            =   "0101"
      Top             =   1200
      Width           =   975
   End
   Begin SDOCdmLibCtl.SDOCdm SDOCdm 
      Height          =   735
      Left            =   240
      OleObjectBlob   =   "cwd.frx":0E9E
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Data DataWTH 
      Caption         =   "DataWTH"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   465
      Left            =   3225
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1935
      Width           =   2040
   End
   Begin TRANSPROVLibCtl.TransactionProvider SDOTrans 
      Height          =   735
      Left            =   1560
      OleObjectBlob   =   "cwd.frx":0ED4
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin S3EADVBROWSERLibCtl.AdvBrowser Browser 
      Height          =   465
      Left            =   1605
      OleObjectBlob   =   "cwd.frx":0F14
      TabIndex        =   2
      Top             =   1935
      Width           =   1590
   End
   Begin S3ELINEOUTLib.S3ELineOut SDOLineOut 
      Height          =   825
      Left            =   3120
      TabIndex        =   3
      Top             =   885
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   1455
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
      Height          =   615
      Left            =   3105
      TabIndex        =   0
      Top             =   150
      Width           =   1335
   End
   Begin DLLib.DL Pcb3Dl 
      Left            =   3120
      Top             =   930
      _Version        =   65538
      _ExtentX        =   2328
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin VB.Line Line1 
      X1              =   15
      X2              =   5175
      Y1              =   1740
      Y2              =   1740
   End
End
Attribute VB_Name = "Cwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All variable need to be declared early
'==========================================================================================
'版权说明:  迪堡公司中国区技术部
'版本来源：原工行版Opteva取款程序（有Timer方式）版本号1.3.16
'版本号： 1.1.0.16 (2005.11.10)
'生成日期：2005.8
'作者：  孙世方
'模块功能： 取款流程
'主要函数及其功能
' 全局变量
'    HtmlPrompt1,HtmlPrompt2,HtmlPrompt3,HtmlPrompt4 : 10,20,50,100 面值总金额，用于屏幕显示
'             GBLCWDResult                   取款结论
'            硬件原因：  DF    取款点钞失败       冲正待查
'                       ST    出钞失败           冲正待查
'                       GB    取款极特殊的情况   已记帐,无法判断是否出钞，待查
'                       PF    RevFloating       已记帐,无法判断是否出钞，待查
'            操作原因：  PT    超时未拿          已记帐
'                       PO    超时未拿但回收钞票张数为零  已记帐
'                       CR      做案              已记帐
'                       CC     卡被吞            冲正待查
'            通讯原因：  CS    dosend<>0         需手工对帐
'                       CE    doreceive<>0      需手工对帐
'                       CU    非定义返回码       需手工对帐
'            GBLKeepAccountFlag      取款记帐标志 初值是 N
'                     收到主机确认，已记帐   Y
'                      冲正成功             R
'                     通讯原因或冲正失败     U     需手工对帐
'       流程修改：
'         1 先输入金额，然后调用dowithdrawal
'         2 在取款流程中不再调用timer
'==========================================================================================
'<时间>：[2005.08.22]
'<修改者>：孙世方
'<当前版本>：1.0.16(中行版）
'<详细记录>：
'   从3030取款移植到Opteva取款，移植工行统计值处理  cwdoktotal sendexception CwdReversalTotal
'   删除cwdlog：交行需求,增加WriteTranslog：工行
'   删除3030特有函数：RecordTakenNotesTimeOut，CassetteNotesChangeVerify，
'                    GetCashUnitsTotal , GetCashUnitsType
'   删除变量g_sCassettesNotesDetail，不在记录billbox
'
'   跨钞箱(不同面值)取款
'     通过将XFS中“DispByPosition”和“Count Control dispensing”的参数设置为YES，可以实现以下功能：
'    1）跨钞箱(不同面值)取款
'    2）指定每个钞箱（即使有同面额的钞箱）各出一张钱
'    3）每个钞箱NoteToDispense的数是准确的
'
' 为了满足中行先退卡后出钞的流程，引用IDC，增加关于退卡的处理；在SDOCdm_BefDeliver判断配置参数。
' 在Atmconfig配置退卡方式，先退卡"A" - After,"B" - Before  GBLTakeCardSequence
'==========================================================================================
'<时间>：[2005.10.22]
'<修改者>：孙世方
'<详细记录>：
'      中银卡取款前增加发送ERI，显示兑换率
'<时间>：[2005.11.1]
'<修改者>：孙世方
'<详细记录>：
'    1 在CwdCrimeFunction函数中增加记录cwdoktotal 和打印钞箱情况PrintCassLeftNum
'    2 RevStateFloating时增加记录cwdoktotal 和打印钞箱情况PrintCassLeftNum
'==========================================================================================
'<时间>：[2005.11.9]
'<修改者>：赵文明
'<详细记录>：修改流水记录内容
'==========================================================================================
'<时间>：[2005.11.10]
'<修改者>：孙世方
'<详细记录>：
' 1 主机返回流水号与发送不同时不发送冲正
' 2 继续修改流水内容
' 3 增加发送AEX处理
'       2009  客户操作超时
'       2012  客户未取卡
'       2013  主机返回流水号或金额有误，或与上送不一致
'       2015  客户超时未取钞票
'       2036  客户取消交易
'4 主机会返回AUP,现与ATP一样处理；广东行返回DQP,DTP,DUP
'==========================================================================================
'<时间>：[2005.11.22]
'<修改者>:vincent
'<详细记录>：
'   1 只有点钞失败时发送冲正，通讯失败不发送
'   2 修改冲正方式，直接发,不在队列
'==========================================================================================
'<时间>：[2005.11.25]
'<修改者>:孙世方
'<详细记录>：
'注册表中cashunit数目在某种情况下会随着动钞箱而增长，SDOCdm.NbrOfBoxesUsed也会随之增长，
'之前的程序会出现CasLefNum越界的情况。修改PrintCassLeftNum函数解决这个问题。
'==========================================================================================
'<时间>：[2005.12.12]
'<修改者>：孙世方
'版本号： 1.2.16 (2005.12.12)
'<详细记录>： 修改CwdOkTotal, 增加记录CutOff.ini，用于主机cutoff时打印流水
'<时间>：[2005.12.20]
'版本号： 1.3.16
'<详细记录>: 修改SDOCdm_AtWithdrEnd中99 else处理，增加记录translog
'   删除ResetATMPrr函数的调用，增加调用DrawWthPrr打印拒绝收条
'<时间>：[2005.12.27]
'版本号： 1.4.16
'<详细记录>:
'   1 修改SendCwdReversal函数，增加对冲正结果的判断
'   2 增加ST交易，冲正
'   3 增加CheckReversalPossible,在ReversalState=0时，根据例外代码和MMDCode.ini中内容的对比，决定是否停机，是否冲正
'   4 修改CwdReversalTotal函数,冲正金额\100
Option Explicit
'Return values in At_WithdrawEnd
Const rcDEVCDM_DOWTH_T_O_ON_TAKE_NOTES      As Integer = 95

'define the line error status 2004/12/25
Const rcDEV_CONNECT_ERROR                   As Integer = 200
Const rcDEV_WRITE_ERROR                     As Integer = 201
Const rcDEV_NO_WRITE_INVALID_MSG            As Integer = 202
Const rcDEV_RECEIVE_ERROR                   As Integer = 203
Const rcDEV_RECEIVE_TIMEOUT                 As Integer = 204
Const rcDEV_RECEIVE_INVALID_MSG             As Integer = 205
Const BeepEffect_WFS_SIU_ON                 As Integer = 2

'the possible value of "RevState" property in SDOCdm SDO
Const RevStateNeeded                        As Integer = 1
Const RevStateFloating                      As Integer = 2

Const NormalAction                          As Integer = 1
Const PinErrorRetryAction                   As Integer = 2

Const NORMAL_MESSAGE                        As Integer = 0

Const ReturnOk                              As Integer = 209
Const ReturnPossibleCrime                   As Integer = 309
Const ReturnNoteNotTaken                    As Integer = 409
Const ReturnCwdReverse                      As Integer = 509
Const ReturnPressStop                       As Integer = 20
Const ReturnHostReject                      As Integer = 30
Const ReturnPinNotMatch                     As Integer = 31
Const ReturnCommErr                         As Integer = 60
Const ReturnTimeout                         As Integer = 80
Const ReturnScreenError                     As Integer = 81
Const ReturnFailed                          As Integer = 120
Const ReturnRevFloat                        As Integer = 130
Const ReturnCardCapture                     As Integer = 800

Const UserReplyCommErr                      As Integer = 160
Const UserReplyHostReject                   As Integer = 130
Const UserReplyAmountNoEqu                  As Integer = 200
Const UserReplyCardCapture                  As Integer = 210

Const DB_WthLogPath                         As String = "C:\S3e\Logs\LogTo\CWDLog.mdb"
Const TransLogFile                          As String = "C:\TransLog.TXT"
Const CardRetainFile                        As String = "C:\s3e\logs\logapp\CardRetain.txt"
Const sLogAppPath                           As String = "C:\S3E\Logs\LogApp\"
Const sGlobalIni                            As String = "C:\ATMWosa\Ini\global.ini"
Const CutOffIni                             As String = "c:\ATMWosa\Ini\CutOff.ini"

'the mininum denomination of available cassette(s)
Dim GLdblMaxWthAmount                       As Long
Private GLnMinAvailDenom                    As Integer
Private GLnMaxAvailDenom                    As Integer

Dim IsNotesPresented                        As Boolean
Dim g_bIsWthReversal                        As Boolean
Dim g_HostCurrentDate                       As String
Dim g_HostCurrentTime                       As String
Dim g_iWithdrawAmount                       As Integer
Dim GLnAction                               As Integer
Dim iRc                                     As Integer
Dim g_nNumOfRetractCount                    As Integer
Dim TakeCardSenqueue                        As String

Dim g_sDetailLineNum                        As String
Dim g_bIsLocalBankCard                      As Boolean
Dim g_sCurrentDate                          As String
Dim HostSeq                                 As String
Dim g_sRejectCode                           As Variant
Dim g_sPrjLanguage                          As String

Dim CasDenomArray(1 To 5)                   As Integer
Dim TotDenomNum                             As Byte
'==========================================================================================
'函 数 的 功 能 ：VB窗口装载,初始化功能键，打开数据库
'输 入 参 数   ：无
'输 出 参 数   ：无
'返 回 值      ：无
'作 者
'创 建 时 间   :
'==========================================================================================
Private Sub Form_Load()

    ' Reset the PcB3HtmlBrowser variables
    iRc = Pcb3Dl.DlSetCharRaw("HtmlFkeyList", "")
    iRc = Pcb3Dl.DlSetCharRaw("HtmlFkeyMap", "3855")
    
    '交易明细数据库，用于结帐时查询、打印
    DataWTH.DatabaseName = DB_WthLogPath
    DataWTH.RecordSource = "Select * from CWDLOG"
    DataWTH.Refresh
    
    If GetIniS(sGlobalIni, "Bank_Environment", "PrjLanguage", "E") = "E" Then
        g_sPrjLanguage = "E"
    Else
        g_sPrjLanguage = "C"
    End If
     
    '在这里设置退卡顺序  A 后退卡，B 先退卡
    TakeCardSenqueue = GetIniS(sGlobalIni, "Bank_Environment", _
                        "EjectCardMode", "A")
    
    SDOTrans.Available = True
    Cwd.WindowState = 1
End Sub
'==========================================================================================
'函 数 的 功 能 ：取款程序退出
'输 入 参 数   ：无
'输 出 参 数   ：无
'返 回 值      ：无
'作 者
'创 建 时 间   :
'==========================================================================================
Private Sub SDOTrans_QuitTransaction()
    Unload Cwd
End Sub
'==========================================================================================
'函数功能 ：初始化数据，得到与取款相关的信息
'输入参数：无
'输出参数：无
'返回值 ：布尔变量 true - 可以取款  false-不能取款
'调用函数：
'  GetWTHCassettesTotal :  得到当前钞箱钞票总张数(限于取款钞箱)
'  GetAllCasDenominations: 得到当前可用钞箱最大、最小面值；当前允许最大取款限额
'被 调 用 情 况：SDOTrans_StartTransaction
'作者：郭健
'创建时间: 2005-08-05
'==========================================================================================
Private Function InitWithdrawal() As Boolean
    Dim i                                           As Integer
    Dim sCasState                                   As String
    Dim szMsg                                       As String
    
    InitWithdrawal = True
    iRc = Pcb3Dl.DlSetCharRaw("GBLAmount", "")
    iRc = Pcb3Dl.DlSetCharRaw("GBLPrtAmount", "00000000")
    g_iWithdrawAmount = 0
    
    SDOCdm.DataCriteria = 1
    For i = 1 To SDOCdm.NbrOfBoxesUsed
        SDOCdm.CasNbrLogical = i
        Select Case SDOCdm.CasState
        Case 0, 1, 2
            sCasState = "0"
        Case 3
            sCasState = "1"
        Case 4
            sCasState = "2"
        Case Else
            sCasState = "4"
        End Select
        If Len(SDOCdm.CasPosition) > 0 Then
            If IsNumeric(Right(SDOCdm.CasPosition, 1)) Then
                iRc = Pcb3Dl.DlSetCharRaw("DeviceCasState" & Right(SDOCdm.CasPosition, 1), sCasState)
            End If
        End If
    Next i
    
    g_nNumOfRetractCount = SDOCdm.RetractCount
    
    iRc = GetAllCasDenominations()
    If iRc = 1 Then
        InitWithdrawal = False
    End If
    
End Function

'==========================================================================================
'函 数 的 功 能 ：取款程序入口
'输 入 参 数   ： 在ruler调用此模块时Action值
'输 出 参 数   ：无
'返 回 值      ：无
'作 者
'创 建 时 间   :
'==========================================================================================
Private Sub SDOTrans_StartTransaction(ByVal action As Long)
    Dim bRet                                        As Boolean
    Dim sSubStData                                  As String
    Dim PrjString                                   As String
    Dim PrjCHNString                                As String
    Dim InputAmount                                 As String
    Dim sFitCardType                                As String
    Dim g_sRejectCode                               As Variant
    Dim ExchangeRate                                As Variant
    Dim sMsg                                        As String
    
    Start.Enabled = False
    
    sFitCardType = Pcb3Dl.DlGetCharRaw("FitCardType")
    
    '中银卡：旧CardType=03  新 CardType =04 需要先查询兑换率
    If sFitCardType = "03" Or sFitCardType = "04" Then
       iRc = SDOLineOut.DoSend("ERI", NORMAL_MESSAGE)
       If iRc <> 0 Then
           iRc = ShowScreenSync(Browser, "Common", "ComCommErr", sSubStData)
           RetToMaster ReturnCommErr
           Exit Sub
       End If
       iRc = SDOLineOut.DoReceive()
       If iRc <> 0 Then
           iRc = ShowScreenSync(Browser, "Common", "ComCommErr", sSubStData)
           RetToMaster ReturnCommErr
           Exit Sub
       End If
       
       g_sRejectCode = Pcb3Dl.DlGetCharRaw("HostTransCode")
       If g_sRejectCode = "AIP" Or g_sRejectCode = "DIP" Then
           iRc = SDOLineOut.GetData("ExchangeRate", ExchangeRate)
           iRc = Pcb3Dl.DlSetCharRaw("HtmlPrompt1", ExchangeRate)
           iRc = ShowScreenSync(Browser, "Cwd", "CwdRateDsp", sSubStData)
           
           Select Case (iRc)
           Case 0:
               Select Case sSubStData
                   Case "@Continue":
                   
                   Case Else:
                       iRc = ShowScreenSync(Browser, "Common", "ComPressStop", sSubStData)
                       RetToMaster ReturnPressStop
                       Exit Sub
               End Select
            Case 91:
                RetToMaster ReturnTimeout
                Exit Sub
            Case Else:
                RetToMaster ReturnPressStop
                Exit Sub
            End Select
        ElseIf g_sRejectCode = "ATP" Or g_sRejectCode = "DTP" Then
            iRc = SDOLineOut.GetData("constHostRejectCode", g_sRejectCode)
            iRc = Pcb3Dl.DlSetCharRaw("ATMPRejectCode", g_sRejectCode)
            iRc = ShowScreenSync(Browser, "Common", "ComReject", sSubStData)
            RetToMaster ReturnHostReject
            Exit Sub
        Else
           iRc = ShowScreenSync(Browser, "Common", "ComCommErr", sSubStData)
           RetToMaster ReturnCommErr
           Exit Sub
        End If
    End If
    
    '银联卡（非本行卡）CardType =01
    If sFitCardType = "01" Then
        g_bIsLocalBankCard = False
    Else
        g_bIsLocalBankCard = True
    End If
    
    GLnAction = action
    
    g_bIsWthReversal = False
    IsNotesPresented = False
    
    iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "N")
    iRc = Pcb3Dl.DlSetCharRaw("ATMPRejectCode", "0000")
    
    If action = NormalAction Then
        bRet = InitWithdrawal()
        If (Not bRet) Then
            PrjString = "  **  No cassettes could be used "
            PrjCHNString = "  无可用取款钞箱 " + vbCrLf
            sMsg = "TotDenomNum is zero, so no legal cassette is available."
            Call PrintJournalMedia(PrjString, PrjCHNString)
            Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "41")
            Call LogWarning(sMsg)
        
            iRc = ShowScreenSync(Browser, "Cwd", "CwdFailed", sSubStData)
            RetToMaster ReturnFailed
            Exit Sub
        End If
        
        iRc = FastWithdrawalSelect()
        If iRc = 4 Then
            iRc = InputWithdrawalAmount()
        End If
        
        Select Case iRc
        Case 0:    'InputWithdrawalAmount OK
            InputAmount = Pcb3Dl.DlGetCharRaw("GBLAmount")
            PrjString = "    Input Amount is :" + InputAmount
            PrjCHNString = " 输入金额为: " + InputAmount + vbCrLf
            Call PrintJournalMedia(PrjString, PrjCHNString)
        Case 2:
            PrjString = "    Keyboard input timeout in Withdrawal"
            PrjCHNString = " 客户键盘输入超时 " + vbCrLf
            Call PrintJournalMedia(PrjString, PrjCHNString)
            
            Call SendAEXMessage("2009")
            RetToMaster ReturnTimeout
            Exit Sub
        Case Else:                 ' 取消 1  画面错误 3
            sMsg = "    Customer Exit in Cwd"
            Call LogInfo(sMsg)
            Call SendAEXMessage("2036")
            iRc = ShowScreenSync(Browser, "Common", "ComPressStop", sSubStData)
            RetToMaster ReturnPressStop
            Exit Sub
        End Select
    End If
           
    iRc = SDOCdm.DoWithdrawal
    If iRc <> 0 Then
        PrjString = " **   DoWithdrawal error RC=" & CStr(iRc)
        PrjCHNString = " 客户取款返回错误 RC=" & CStr(iRc) + vbCrLf
        sMsg = PrjString
        
        Call PrintJournalMedia(PrjString, PrjCHNString)
        Call LogError(sMsg)
        Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "41")
        
        iRc = ShowScreenSync(Browser, "Cwd", "CwdFailed", sSubStData)
        RetToMaster ReturnFailed
    Else
        SDOCdm.TimeOutSecondsFirst = -1
    End If
        
End Sub

Private Sub Start_Click()
    Dim action                                      As Integer
    Dim sSubStData                                  As String
    Dim PrjString                                   As String
    Dim PrjCHNString                                As String
    Dim ExchangeRate                                As String
    
    Call CheckReversalPossible
    
    action = 1
    
    ExchangeRate = "0010788"
    iRc = Pcb3Dl.DlSetCharRaw("HtmlPrompt1", ExchangeRate)
    iRc = ShowScreenSync(Browser, "Cwd", "CwdRateDsp", sSubStData)
           
           Select Case (iRc)
           Case 0:
               Select Case sSubStData
                   Case "@Continue":
                   
                   Case Else:
                       iRc = ShowScreenSync(Browser, "Common", "ComPressStop", sSubStData)
                       RetToMaster ReturnPressStop
                       Exit Sub
               End Select
            Case 91:
                RetToMaster ReturnTimeout
                Exit Sub
            Case Else:
                RetToMaster ReturnPressStop
                Exit Sub
            End Select
    
    GLnAction = action
    
    If action = NormalAction Then
        iRc = Pcb3Dl.DlSetCharRaw("GBLAmount", "")
        
        iRc = Pcb3Dl.DlSetCharRaw("GBLPrtAmount", "00000000")
                
        iRc = Pcb3Dl.DlSetCharRaw("CwdAmtRetry", "3")
        
        iRc = Pcb3Dl.DlReset("GBLATMLocRejCode")
        GLdblMaxWthAmount = 300000
        iRc = Pcb3Dl.DlSetCharRaw("HtmlPrompt1", CStr(GLdblMaxWthAmount))
        
 
    End If
    
    'iRc = Pcb3Dl.DlSetCharRaw("PrrReplyCode", "0000")
    
    iRc = InputWithdrawalAmount
    Select Case iRc
    Case 0:
        'InputWithdrawalAmount OK
    Case 2:
        PrjString = " **   Keyboard input timeout in Withdrawal"
        PrjCHNString = "    **客户取款交易时输入超时"
        'add by nicktan add "*"
        Call PrintJournalMedia(PrjString, PrjCHNString)
        
        Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "46")
        RetToMaster ReturnTimeout
        Exit Sub
    Case Else:
        PrjString = "  **  Customer Exit in Cwd"
        PrjCHNString = "    **客户在取款交易时选择退出"
         'add by nicktan add "*"
        Call PrintJournalMedia(PrjString, PrjCHNString)
        
        Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "45")
        iRc = ShowScreenSync(Browser, "Common", "ComPressStop", sSubStData)
        RetToMaster ReturnPressStop
        Exit Sub
    End Select
        
End Sub
'==========================================================================================
'版本号：Agilis 1.6
'参见sdohelp文件DoWithdrawal方法，针对其中九个事件进行处理
'==========================================================================================
Private Sub SDOCdm_AtWithdrStart()
    Call LogInfo("SDOCdm_AtWithdrStart=0")
    SDOCdm.UserReply = 0
End Sub

Private Sub SDOCdm_InformDenomNotPresent(ByVal AbsentDenom As Long)
    Call LogInfo("SDOCdm_InformDenomNotPresent=0")
    SDOCdm.UserReply = 0
End Sub

Private Sub SDOCdm_GetWithdrawalAmount()
    Dim lUserReply              As Long
    Dim sAmount                 As String
    
    SDOCdm.Currency = "CNY"
    
    Select Case GLnAction
    Case NormalAction
        sAmount = Pcb3Dl.DlGetCharRaw("GBLAmount")
        
        SDOCdm.WithdrawalAmount = CInt(sAmount)
        lUserReply = 0
    Case PinErrorRetryAction
'密码输入错误，主机拒绝后，进入再次输入密码，然后直接进入取款，不用再次输入金额
        iRc = Pcb3Dl.DlSetCharRaw("GBLAmount", CStr(SDOCdm.WithdrawalAmount))
        lUserReply = 0
    Case Else
        LogError "Select GLnAction error, GLnAction = " + CStr(GLnAction)
        lUserReply = 100
    End Select
    
    Call LogInfo("SDOCdm_GetWithdrawalAmount=" & CStr(lUserReply))
    SDOCdm.UserReply = lUserReply
End Sub
Private Sub SDOCdm_BefAuthorisation()
    Call LogInfo("SDOCdm_BefAuthorisation=0")
    SDOCdm.UserReply = 0
End Sub
'==========================================================================================
' 注释：在Opteva上，跨钞箱时平台提供的每个钞箱点钞等数据是准确的
'==========================================================================================
Private Sub SDOCdm_GetAuthorisation(ByVal WithdrawalAmount As Long)
    Dim WthAmount          As Variant
    Dim CommReturn         As Integer
    Dim sSubStData         As String
    Dim sMsg               As String
    
    WthAmount = Pcb3Dl.DlGetCharRaw("GBLAmount")
    If WithdrawalAmount <> WthAmount Then
        LogError "<GetAuthorisation>: GetWithdrawalAmount=" + _
                CStr(WithdrawalAmount) + "InputWthAmount=" + _
                CStr(WthAmount)
        
        SDOCdm.UserReply = UserReplyAmountNoEqu
    Else
        CommReturn = CommunicationSubFunction()
        sMsg = "SDOCdm_GetAuthorisation UserReply =" & CStr(CommReturn)
        Call LogInfo(sMsg)
        SDOCdm.UserReply = CommReturn
        If CommReturn = 0 Then
            iRc = ShowScreenSync(Browser, "Cwd", "CwdProcIdle", sSubStData)
        End If
    End If
End Sub
Private Sub SDOCdm_BefDeliver()
    Dim ssTrack3Update     As String
    Dim TakeCard           As String
    Dim sSubStData         As String
    Dim PrjString          As String
    Dim PrjCHNString       As String
    
    '后退卡
    If TakeCardSenqueue = "A" Then
        Call LogInfo("SDOCdm_BefDeliver=0")
        SDOCdm.UserReply = 0
    Else
        '先退卡
        iRc = ShowScreenSync(Browser, "EndVisit", "TakeCard", sSubStData)
            
        ssTrack3Update = Pcb3Dl.DlGetCharRaw("IcbcTrackUpdate")
        If Len(ssTrack3Update) <> 0 Then
            SDOIdc.IsoTrack3 = ssTrack3Update
        End If
        
        iRc = SDOIdc.DoEjectCard
        If iRc <> 0 Then
            LogError "DoEjectCard method Error. RC=" & iRc
            Call CaptureCard
            Call LogInfo("SDOCdm_BefDeliver=" & CStr(UserReplyCardCapture))
            SDOCdm.UserReply = UserReplyCardCapture
        End If
    End If
End Sub
Private Sub SDOCdm_NotesPresented()
    Dim sSubStData As String
    
    IsNotesPresented = True
    Call LogInfo("SDOCdm_NotesPresented")
    iRc = ShowScreenSync(Browser, "Cwd", "CwdTakeNote", sSubStData)
    If iRc <> 0 Then
        LogError ScreenInfo.Name + "Return error, iRc = " + CStr(iRc)
    End If
End Sub
Private Sub SDOCdm_PleaseTakeNotes()
    Call LogInfo("SDOCdm_PleaseTakeNotes=0")
    SDOCdm.UserReply = 0
End Sub
Private Sub SDOCdm_AtWithdrEnd(ByVal WithdrRc As Integer)
    Dim sSubStData                              As String
    Dim sHostRejectCard                         As String
    Dim nCwdReturnValue                         As Integer
    Dim PrjString                               As String
    Dim PrjCHNString                            As String
    Dim sCorrCode                               As String
    Dim ReversalFlag                            As Boolean
    
    nCwdReturnValue = ReturnFailed
    PrjString = "SDOCdm_AtWithdrEnd WithdrRc =" + CStr(WithdrRc)
    Call LogInfo(PrjString)
    
     'Inform Operator of recounting the reject and retract notes
    iRc = Pcb3Dl.DlSetLong("OptevaMonType", 4)
    
    Select Case WithdrRc
    Case 0:
        PrjString = "     TRANSACTION OK"
        PrjCHNString = "　　交易成功完成" + vbCrLf
        Call PrintJournalMedia(PrjString, PrjCHNString)
        '记录数据库、交易日志
        iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "OK")
        Call RecordDB_CWDLog
        Call WriteTranslog("取款成功")
        
        Call CwdOkFunction
        nCwdReturnValue = ReturnOk
    
    Case UserReplyHostReject:
        iRc = ShowScreenSync(Browser, "Common", "ComReject", sSubStData)
        PrjString = "   **HOST REJECT [" + g_sRejectCode + "]" + vbCrLf + _
                                   "   **" + Pcb3Dl.DlGetCharRaw("HostRejectEnglish")
                
        PrjCHNString = "   **主机拒绝 [" + g_sRejectCode + "]" + vbCrLf + _
                               "   **" + Pcb3Dl.DlGetCharRaw("HostRejectChinese") + vbCrLf
        Call PrintJournalMedia(PrjString, PrjCHNString)
        
        sHostRejectCard = Pcb3Dl.DlGetCharRaw("HostRejectCard")
                
        If sHostRejectCard = "R" Then
            nCwdReturnValue = ReturnPinNotMatch
        Else
            Call DrawWthPrr(Pcb3Dl, WthPrrReject)   '2005.12.21打印拒绝收条
            nCwdReturnValue = ReturnHostReject
        End If
        
    Case UserReplyCommErr:
        Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "64")
        
        '增加记录文件 2005.11.14
        Call WriteTranslog("通信失败")
        
        iRc = ShowScreenSync(Browser, "Common", "ComCommErr", sSubStData)
        If g_bIsWthReversal = True Then
            Call DrawWthPrr(Pcb3Dl, WthPrrCWC)
            Call RecordDB_CWDLog
            nCwdReturnValue = ReturnCwdReverse
        Else
            nCwdReturnValue = ReturnCommErr
        End If

    Case UserReplyAmountNoEqu:
        LogError "SDOCdm_AtWithdrEnd's WithdrRc = " + CStr(WithdrRc)
        iRc = ShowScreenSync(Browser, "Cwd", "CwdFailed", sSubStData)
        nCwdReturnValue = ReturnFailed
             
    Case UserReplyCardCapture
        'Send withdrawal reversal to host
        iRc = SendCwdReversal("4002")
        Call CwdReversalTotal
        Call SendAEXMessage("2012")
        
        PrjString = "  ** Withdraw Reversal 4002" + vbCrLf + "  Customer does not take card in CWD" + vbCrLf
        PrjCHNString = "  取款冲正" + vbCrLf + "  冲正原因：客户未取卡." + vbCrLf
        Call PrintJournalMedia(PrjString, PrjCHNString)
        
        '记录数据库，交易日志
        iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "CC")
        Call RecordDB_CWDLog
        Call WriteTranslog("卡片被吞")
        
        nCwdReturnValue = ReturnCardCapture
             
    Case rcDEVCDM_DOWTH_T_O_ON_TAKE_NOTES:
        If SDOCdm.Available = False And SDOCdm.OperatorType = optype_cdm_shutterproblem Then
            
            Call CwdCrimeFunction
            
            nCwdReturnValue = ReturnPossibleCrime
        Else
            If g_nNumOfRetractCount = SDOCdm.RetractCount Then
               
                PrjString = "  ** Withdrawl Timeout, But Retract Failed!" + vbCrLf
                PrjCHNString = "  取款超时，但回收钞票张数为零" + vbCrLf
                Call PrintJournalMedia(PrjString, PrjCHNString)
                
                '记录数据库、交易日志
                iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "PO")
                Call RecordDB_CWDLog
                Call WriteTranslog("取款成功")
                
                Call CwdOkFunction
                nCwdReturnValue = ReturnOk
            Else
                '打印流水
                PrjString = "** Take notes timeout"
                PrjCHNString = "  客户未取钞" + vbCrLf
                Call PrintJournalMedia(PrjString, PrjCHNString)
                 
                '记录取款统计值
                Call CwdOkTotal
                 
                '记录数据库，交易日志
                iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "PT")
                Call RecordDB_CWDLog
                Call WriteTranslog("取钞超时")
                
                '发送例外信息
                Call SendAEXMessage("2015")
                
                iRc = ShowScreenSync(Browser, "Cwd", "CwdTakeNoteTimeout", sSubStData)
                
                 '准备收条打印内容 超时未取钞
                Call DrawWthPrr(Pcb3Dl, WthPrrTimeout)
                nCwdReturnValue = ReturnNoteNotTaken
            End If
        End If
           
    Case Else:                  '98,99 和其他
        LogError "SDOCdm_AtWithdrEnd's WithdrRc = " + CStr(WithdrRc)
        LogError "sdocdm.reversalstate = " + CStr(SDOCdm.ReversalState)
        
       Select Case (SDOCdm.ReversalState)
        Case RevStateNeeded:
            '打印流水
            PrjString = "** Cash dispenser error RC=" & CStr(WithdrRc) + " Need Reverse"
            PrjCHNString = "  点钞时取款模块故障 RC=" & CStr(WithdrRc) + " 需要冲正" + vbCrLf
            Call PrintJournalMedia(PrjString, PrjCHNString)
            
            '记录数据库，交易日志
            iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "DF")
          
            Call WriteTranslog("出钞失败")
            
            Select Case SDOCdm.OperatorType
            Case optype_cdm_somecasslow, optype_cdm_casnotconfigured, optype_cdm_notesproblem, optype_cdm_casinvalid
                sCorrCode = "4007"
            Case optype_cdm_allempty
                sCorrCode = "4012"
            Case Else
                sCorrCode = "4009"
            End Select
                        
            '打印流水
            PrjString = "    Send Reversal , code =" & sCorrCode
            PrjCHNString = "    发送冲正,代码=" & sCorrCode + vbCrLf
            Call PrintJournalMedia(PrjString, PrjCHNString)
            
             '发送冲正和例外信息
            iRc = SendCwdReversal(sCorrCode)
            Call CwdReversalTotal
            Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "41")
            Call RecordDB_CWDLog
            
            iRc = ShowScreenSync(Browser, "Cwd", "CwdFailed", sSubStData)
            
            '准备收条打印内容 冲正
            Call DrawWthPrr(Pcb3Dl, WthPrrCWC)
            nCwdReturnValue = ReturnCwdReverse
            
        Case RevStateFloating:
            '打印流水
            PrjString = "** Cash presenter error RC=" & CStr(WithdrRc) + " No Reverse"
            PrjCHNString = "  送钞时取款模块故障 RC=" & CStr(WithdrRc) + " 未冲正" + vbCrLf
            'modi by nicktan change the "#" to "*"
            Call PrintJournalMedia(PrjString, PrjCHNString)
            
             '记录数据库，交易日志
            iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "PF")
            Call RecordDB_CWDLog
            Call WriteTranslog("出钞异常")
            Call CwdOkTotal
            Call PrintCassLeftNum
            
            '发送例外信息
            Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "41")
            iRc = ShowScreenSync(Browser, "Cwd", "CwdRevFloat", sSubStData)

            '准备收条打印内容 取款待查
            Call DrawWthPrr(Pcb3Dl, WthPrrFloat)
            nCwdReturnValue = ReturnRevFloat
        
        Case Else:
            If IsNotesPresented = True Then
                If (SDOCdm.Available = False And SDOCdm.OperatorType = optype_cdm_shutterproblem) Then
                    Call CwdCrimeFunction
                    RetToMaster ReturnPossibleCrime
                Else
                    Call PrintJournalMedia("   **  TRANSACTION OK**", "  交易成功完成")
                     '记录数据库、交易日志
                    iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "OK")
                    Call RecordDB_CWDLog
                    Call WriteTranslog("出后故障")
                    Call CwdOkFunction
                End If
            Else
                
                If (Not CheckReversalPossible) Then
                      '打印流水
                    PrjString = " Other reason Reversal"
                    PrjCHNString = " 其他原因冲正," + vbCrLf
                    Call PrintJournalMedia(PrjString, PrjCHNString)
                    
                    iRc = SendCwdReversal("4009")
                    Call CwdReversalTotal
                     '记录数据库
                    iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "ST")
                    Call RecordDB_CWDLog
                    Call WriteTranslog("其他冲正")          ' 增加记录2005。12。20
                Else
                     '打印流水
                    PrjString = "Need Check"
                    PrjCHNString = "待查交易" + vbCrLf
                    Call PrintJournalMedia(PrjString, PrjCHNString)
                    
                     '记录数据库
                    iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "GB")
                    Call RecordDB_CWDLog
                    Call WriteTranslog("取款待查")          ' 增加记录2005。12。20
                End If
                
                '发送例外信息
                Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "41")
       
                iRc = ShowScreenSync(Browser, "Cwd", "CwdFailed", sSubStData)
                nCwdReturnValue = ReturnFailed
            End If
        End Select              'ReversalState
    End Select
    
    RetToMaster nCwdReturnValue
    
End Sub
'==========================================================================================
'函数功能 ：输入取款金额
'输入参数：无
'输出参数：无
'返回值 ：
'          0: OK
'          1: Customer Cancel
'          2: Input Timeout
'          3: Error
'调用函数：
'被 调 用 情 况：SDOTrans_StartTransaction
'作者：郭健
'创建时间: 2005-08-05
'==========================================================================================
Private Function InputWithdrawalAmount() As Integer
    Dim bLoop               As Boolean
    Dim sSubStData          As String
    Dim StrWthAmount        As String
    Dim lWthAmount          As Long
 
    Dim sPrompt             As String
    Dim i                   As Integer
    
    sPrompt = ""
    For i = 1 To TotDenomNum
        If i = TotDenomNum Then
            sPrompt = sPrompt + CStr(CasDenomArray(i))
        Else
            sPrompt = sPrompt + CStr(CasDenomArray(i)) + "和"
        End If
    Next
    Pcb3Dl.DlSetCharRaw "HtmlPrompt2", sPrompt
    
    '选择其他金额进入输入画面
    bLoop = True
    While (bLoop)
        iRc = Pcb3Dl.DlReset("GBLAmount")

        iRc = ShowScreenSync(Browser, "Cwd", "CwdAmtInput", sSubStData)
        Select Case (iRc)
        Case 0:
            Select Case sSubStData
            Case "@ok":
                StrWthAmount = Pcb3Dl.DlGetCharRaw("GBLAmount")
                If (Not IsNumeric(StrWthAmount)) Then
                    StrWthAmount = "0"
                End If
                lWthAmount = CLng(StrWthAmount)
                If lWthAmount = 0 Or (lWthAmount Mod GLnMinAvailDenom <> 0) Then
                    iRc = ShowScreenSync(Browser, "Cwd", "CwdAmtError", sSubStData)
                ElseIf (lWthAmount > GLdblMaxWthAmount) Then
                    iRc = ShowScreenSync(Browser, "Cwd", "CwdAmtOverLimit", sSubStData)
                Else
                    InputWithdrawalAmount = 0   '确认0
                    iRc = ShowScreenSync(Browser, "Common", "ComPlsWait", sSubStData)
                    Exit Function
                End If
            Case "@Change":
               
            Case Else
                InputWithdrawalAmount = 1   '取消 1
                Exit Function
            End Select
        Case 91:
            InputWithdrawalAmount = 2    '超时2
            Exit Function
        Case Else:
            LogError ScreenInfo.Name + "Return error, irc = " + CStr(iRc)
            InputWithdrawalAmount = 3    '画面错误 3
            Exit Function
        End Select
    
    Wend

End Function
'==========================================================================================
'函数功能 ：选择快速取款金额并确认
'输入参数：无
'输出参数：无
'返回值 ：
'          0: OK
'          1: Customer Cancel
'          2: Input Timeout
'          3: Error
'          4：输入金额
'调用函数：
'被 调 用 情 况：SDOTrans_StartTransaction
'作者：孙世方
'创建时间: 2005-09-07
'==========================================================================================
Private Function FastWithdrawalSelect() As Integer
    Dim FastWthAmount       As String
    Dim bLoop               As Boolean
    Dim sSubStData          As String
    
'快速取款选择
    bLoop = True
    
    '在页面根据当前允许最大取款金额来屏蔽快速取款项
    iRc = Pcb3Dl.DlSetCharRaw("HtmlPrompt1", CStr(GLdblMaxWthAmount))
    
    iRc = ShowScreenSync(Browser, "Cwd", "CwdMenu", sSubStData)
    Select Case (iRc)
    Case 0:
        If sSubStData = "@others" Then              '输入金额
            FastWithdrawalSelect = 4
            bLoop = False
        Else
            If (Not IsNumeric(sSubStData)) Then
                FastWthAmount = "0"
            End If
            If CInt(sSubStData) > 0 Then
                iRc = Pcb3Dl.DlSetCharRaw("HtmlFastcashAmount", sSubStData)
                iRc = Pcb3Dl.DlSetCharRaw("GBLAmount", sSubStData)
                FastWithdrawalSelect = 0
            Else
                LogWarning "Amount Select cancel, SubstData = " + sSubStData
                FastWithdrawalSelect = 1
                Exit Function
            End If
        End If
    Case 91:
        FastWithdrawalSelect = 2
        Exit Function
    Case Else:
        LogError ScreenInfo.Name + "Return error, irc = " + CStr(iRc)
        FastWithdrawalSelect = 3
        Exit Function
    End Select
    
    '已选择金额，需要确认
    If (FastWithdrawalSelect = 0) And bLoop Then       'InputCwdMenu = 0
        iRc = ShowScreenSync(Browser, "Cwd", "CwdConfirmMenu", sSubStData)
        Select Case (iRc)
        Case 0:
            If sSubStData = "@ok" Then
                FastWithdrawalSelect = 0   '确认0
                iRc = ShowScreenSync(Browser, "Common", "ComPlsWait", sSubStData)
            Else
                FastWithdrawalSelect = 1   '取消 1
            End If
        Case 91:
             '继续选择
            FastWithdrawalSelect = 2
        Case Else:
            LogError ScreenInfo.Name + "Return error, irc = " + CStr(iRc)
            FastWithdrawalSelect = 3     '画面错误 3
        End Select
        Exit Function
    End If

End Function
'==========================================================================================
'函 数 的 功 能 ：返回主控模块master
'输 入 参 数   ：返回主控模块值
'输 出 参 数   ：无
'返 回 值      ：无
'调 用 函 数   ：无
'被 调 用 情 况：
'作 者         ：汪林
'创 建 时 间   :
'==========================================================================================
Private Sub RetToMaster(ByVal S3eRetValue As Integer)
    SDOTrans.Result = S3eRetValue
End Sub
'==========================================================================================
'函 数 的 功 能 ：发送取款冲正报文
'输 入 参 数   ：冲正交易例外码
'输 出 参 数   ：无
'返 回 值      ：发送冲正交易的结果
'调 用 函 数   ：
'被 调 用 情 况：
'作 者         ：汪林
'创 建 时 间   :
'==========================================================================================
  Private Function SendCwdReversal(pExceCode As String) As Integer
    Dim vLineNum As Variant
    Dim PrjString  As String
    Dim PrjCHNString As String
    
    iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "U")
    g_bIsWthReversal = True
    
    g_sCurrentDate = Format(Now(), "MMDDHHMMSS")
    
    SDOLineOut.SetData "CurrentDate", g_sCurrentDate
    SDOLineOut.SetData "LocalTransTime", Right(g_sCurrentDate, 6)
    SDOLineOut.SetData "LocalTransDate", Left(g_sCurrentDate, 4)
    SDOLineOut.SetData "CurrencyCode", "001"
    SDOLineOut.SetData "CorrectionCode", pExceCode
    
    iRc = SDOLineOut.DoSend("CWC", 0)
    If iRc <> 0 Then
         LogError "Send Resersal Message(CWC) Failed!"
         PrjString = "Send Resersal Message Failed，please check this transaction" + vbCrLf
         PrjCHNString = "** 发送冲正报文失败,请手工对账" + vbCrLf
    Else
        iRc = SDOLineOut.DoReceive
        If iRc <> 0 Then
            LogError "Receive Resersal Message(CWC) Failed!"
            PrjString = "Receive Resersal Message Failed，please check this transaction" + vbCrLf
            PrjCHNString = "** 接收冲正报文失败,请手工对账" + vbCrLf
        Else
            g_sRejectCode = Pcb3Dl.DlGetCharRaw("HostTransCode")
            If g_sRejectCode = "AWP" Or g_sRejectCode = "DWP" Then
                PrjString = "Reversal OK" + vbCrLf
                PrjCHNString = "冲正成功" + vbCrLf
                iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "R")
            Else
                PrjString = "Reversal Host Reject" + vbCrLf
                PrjCHNString = "冲正被拒绝，,请手工对账" + vbCrLf
            End If
        End If
    End If
    
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    SDOLineOut.GetData "GBLLineNum", vLineNum
    g_sDetailLineNum = Format(vLineNum, "000000")
        
    SendCwdReversal = iRc
End Function
'==========================================================================================
'函 数 的 功 能 :记录取款成功统计值
'输 入 参 数   ：无
'输 出 参 数   ：无
'返 回 值      ：无
'调 用 函 数   ：无
'被 调 用 情 况：CwdOkFunction 和take notes timeout
'作 者         ：汪林
'创 建 时 间   :
'<时间>：[2005.12.12]
'<修改者>：孙世方
'<详细记录>：
'    增加记录CutOff.ini，用于主机cutoff时打印流水
'==========================================================================================
Private Sub CwdOkTotal()
    Dim WthAmount      As Long
    Dim StrWthAmount   As String
    Dim nTotWthNum     As Long
    Dim dblTotWthAmt   As Long
    Dim CutOffWthAmount As String
    Dim CutOffWthNumber As String
    Dim TempNumber As Long
    Dim TempAmount As Long
 
    StrWthAmount = Pcb3Dl.DlGetCharRaw("GBLAmount") \ 100
    WthAmount = CLng(StrWthAmount)
    
    If g_bIsLocalBankCard = True Then
        nTotWthNum = Pcb3Dl.DlGetInt("TotWithdrawNum")
        dblTotWthAmt = Pcb3Dl.DlGetDouble("TotWithdrawAmount")
        nTotWthNum = nTotWthNum + 1
        dblTotWthAmt = dblTotWthAmt + WthAmount
        iRc = Pcb3Dl.DlSetLong("TotWithdrawNum", nTotWthNum)
        iRc = Pcb3Dl.DlSetDouble("TotWithdrawAmount", dblTotWthAmt)
    Else
        nTotWthNum = Pcb3Dl.DlGetInt("IcbcTotExtraWthNum")
        dblTotWthAmt = Pcb3Dl.DlGetDouble("IcbcTotExtraWthAmount")
        nTotWthNum = nTotWthNum + 1
        dblTotWthAmt = dblTotWthAmt + WthAmount
        iRc = Pcb3Dl.DlSetLong("IcbcTotExtraWthNum", nTotWthNum)
        iRc = Pcb3Dl.DlSetDouble("IcbcTotExtraWthAmount", dblTotWthAmt)
    End If
    
    '增加记录CutOff.ini，用于主机cutoff时打印流水
    CutOffWthNumber = GetIniS(CutOffIni, "HostCutOff", "WithdrawNumber", "0")
    TempNumber = CLng(CutOffWthNumber) + 1
    CutOffWthAmount = GetIniS(CutOffIni, "HostCutOff", "WithdrawAmount", "0")
    TempAmount = CLng(CutOffWthAmount) + WthAmount
    If (TempNumber > 5000) Then
        LogError ("CutOffWithAmountTooHigh,Now Clear")
        TempNumber = 0
        TempAmount = 0
    End If
    iRc = SetIniS(CutOffIni, "HostCutOff", "WithdrawNumber", CStr(TempNumber))
    iRc = SetIniS(CutOffIni, "HostCutOff", "WithdrawAmount", CStr(TempAmount))
    
End Sub
'==========================================================================================
'函数功能 ：得到当前可用钞箱最大、最小面值；当前允许最大取款限额
'输入参数 ：无
'输出参数 ：无
'返 回 值 ：0 ---- 表示已经检索到可以使用的钞箱了
'          1---- 表示未检索到合格的钞箱
'作者 ：赵文明
'创建时间 :
'修改记录：2005.8.10
'修改者： 孙世方
'==========================================================================================
Function GetAllCasDenominations() As Integer
    Dim CurCasDenom                 As Long
    Dim i                           As Integer
    Dim j                           As Integer
    Dim MaxBillsCanPresent          As Long
    Dim lFitMaxWthAmount            As Long
    Dim lTotalPresentAmount         As Long
    Dim lCasPresentAmount           As Long
    Dim lMaxAmountCanPresnet        As Long
    Dim IsInArray                   As Boolean
    Dim CurNum                      As Integer
    Dim TmpNum                      As Integer
    
    GLnMaxAvailDenom = 0
    GLnMinAvailDenom = 999
    TotDenomNum = 0
    
    SDOCdm.DataCriteria = 1
    lTotalPresentAmount = 0
    For i = 1 To SDOCdm.NbrOfBoxesUsed
        SDOCdm.CasNbrLogical = i
        If Len(SDOCdm.CasPosition) > 0 Then
            If SDOCdm.CasState <= casstate_cdm_low And SDOCdm.CasState >= casstate_cdm_ok And _
                IsNumeric(Right(SDOCdm.CasPosition, 1)) Then
                CurCasDenom = SDOCdm.CasDenomination
                
                lCasPresentAmount = CurCasDenom * SDOCdm.TotNbrPresent
                lTotalPresentAmount = lTotalPresentAmount + lCasPresentAmount
                
                If CurCasDenom > GLnMaxAvailDenom Then
                    GLnMaxAvailDenom = CurCasDenom
                End If
                If CurCasDenom < GLnMinAvailDenom Then
                    GLnMinAvailDenom = CurCasDenom
                End If
                
                IsInArray = False
                For j = 1 To TotDenomNum
                    If (CurCasDenom = CasDenomArray(j)) Then
                        IsInArray = True
                        Exit For
                    End If
                Next j
                If Not IsInArray Then
                    TotDenomNum = TotDenomNum + 1
                    CasDenomArray(TotDenomNum) = CurCasDenom
                End If
    
            End If
        End If
    Next i
    
    For i = 1 To TotDenomNum - 1
        CurNum = i
        TmpNum = CasDenomArray(i)
        
        For j = i + 1 To TotDenomNum
            If (CasDenomArray(j) > TmpNum) Then
                CurNum = j
                TmpNum = CasDenomArray(j)
            End If
        Next j
        
        CasDenomArray(CurNum) = CasDenomArray(i)
        CasDenomArray(i) = TmpNum
    
    Next i
    
    MaxBillsCanPresent = CLng(Pcb3Dl.DlGetCharRaw("GBLMaxBills"))
               
    lFitMaxWthAmount = CLng(Pcb3Dl.DlGetCharRaw("FitMaxWthAmount"))
    If (GLnMaxAvailDenom <> 0) Then
        lMaxAmountCanPresnet = MaxBillsCanPresent * GLnMaxAvailDenom
        If lFitMaxWthAmount <= lMaxAmountCanPresnet Then
             GLdblMaxWthAmount = lFitMaxWthAmount
        Else
             GLdblMaxWthAmount = lMaxAmountCanPresnet
        End If
        
        If GLdblMaxWthAmount >= lTotalPresentAmount Then
             GLdblMaxWthAmount = lTotalPresentAmount
        End If
        
        iRc = Pcb3Dl.DlSetCharRaw("HtmlPrompt1", _
                      CStr(GLdblMaxWthAmount))
                
         iRc = Pcb3Dl.DlSetCharRaw("CwdAvailDenom", CStr(GLnMinAvailDenom))
         iRc = Pcb3Dl.DlSetCharRaw("CwdMaxDenom", CStr(GLnMaxAvailDenom))
        GetAllCasDenominations = 0
        
    Else
        LogError ("TotDenomNum is zero, so no legal cassette is available.")
        GetAllCasDenominations = 1
    End If

End Function
'==========================================================================================
'函 数 的 功 能 ：与主机进行通讯
'输 入 参 数   ：发送报文类型
'输 出 参 数   ：无
'返 回 值      ：通讯结果，决定出钞还是冲正还是显示主机拒绝 初值 200
'   主机返回码 00  返回0 ，出钞
'   五种情况返回160，不出钞:
'        1 dosend <> 0 2 doreceive <> 0 3 主机返回码 00 4 主机返回流水号、金额与上送不符 5 MAC 校验错
'   以下情况冲正：
'           1 dosend <>0   2 doreceive <> 0
'调 用 函 数   ：
' 与报文加密相关的函数
' 与通讯相关的函数
'           SDOLineOut1.DoSend("0201", NORMAL_MESSAGE)
'           SDOLineOut1.DoReceive()
'被 调 用 情 况：
'     SDOCdm_GetAuthorisation 函数
'作 者         ：
'创 建 时 间   :2005.8
'==========================================================================================
Function CommunicationSubFunction() As Integer
    Dim WthAmount                   As Variant
    Dim HostAccNo                   As String
    Dim StrWthAmount                As String
    Dim PrjString                   As String
    Dim PrjCHNString                As String
    Dim HostResLineNum              As String
    Dim HostResAmount               As Variant
    Dim vGlnLineNum                 As Variant
    Dim Fee                         As String
    Dim PrrFee                      As String
    Dim bReverseCWC                 As Boolean
    Dim iLoop                       As Integer
    Dim nCassNumber                 As Integer
    
    CommunicationSubFunction = UserReplyAmountNoEqu
            
    WthAmount = Pcb3Dl.DlGetCharRaw("GBLAmount")
    
    StrWthAmount = Format(WthAmount, "Standard")
    iRc = Pcb3Dl.DlSetCharRaw("GBLPrtAmount", _
            StrWthAmount)
    
    WthAmount = WthAmount * 100
    StrWthAmount = Format(WthAmount, "00000000")
    iRc = Pcb3Dl.DlSetCharRaw("GBLAmount", StrWthAmount)

    For iLoop = 1 To 4
        iRc = SDOLineOut.SetData("CasNotesPresent" & CStr(iLoop), "00")
    Next iLoop
    
    SDOCdm.DataCriteria = 1
    nCassNumber = SDOCdm.NbrOfBoxesUsed
    For iLoop = 1 To nCassNumber
        SDOCdm.CasNbrLogical = iLoop
            
        If (SDOCdm.CasState < 5) Then
            iRc = SDOLineOut.SetData("CasNotesPresent" & CInt(Right(SDOCdm.CasPosition, 1)), _
                        Format(CStr(SDOCdm.NotesToDispense), "00"))
        End If
    Next iLoop
    
    PrjString = vbCrLf + "     " + "CWD Send" + Format(Now(), " HH:MM:SS")
    PrjCHNString = " 取款报文发送主机" + Format(Now(), " HH:MM:SS") + vbCrLf
    Call PrintJournalMedia(PrjString, PrjCHNString)
    Call LogInfo(PrjString)
    
    iRc = SDOLineOut.DoSend("CWD", NORMAL_MESSAGE)
    
    g_HostCurrentDate = Format(Now(), "MMDD")
    g_HostCurrentTime = Format(Now(), "HHMM")
    SDOLineOut.GetData "GBLLineNum", vGlnLineNum
    g_sDetailLineNum = Format(vGlnLineNum, "000000")
    Pcb3Dl.DlSetCharRaw "GBLLineSendNum", g_sDetailLineNum
            
    PrjString = " " + vbCrLf + _
                "     " + "CWD " + Format(Now(), " HH:MM:SS") + " [" + Format(vGlnLineNum, "000000") + "]" + vbCrLf + _
                "     ATM CODE: " + Format(Pcb3Dl.DlGetCharRaw("GBLAtmCode")) + vbCrLf + _
                "     Amount:" + Pcb3Dl.DlGetCharRaw("GBLPrtAmount") + vbCrLf
    
 
    PrjCHNString = " " + vbCrLf + _
                "     取款 " + Format(Now(), " HH:MM:SS") + " 流水号：[" + Format(vGlnLineNum, "000000") + "]" + vbCrLf + _
                "     ATM号： " + Format(Pcb3Dl.DlGetCharRaw("GBLAtmCode")) + vbCrLf + _
                "     金额： " + Pcb3Dl.DlGetCharRaw("GBLPrtAmount") + vbCrLf

    Call PrintJournalMedia(PrjString, PrjCHNString)

    bReverseCWC = True
    If iRc <> 0 Then
        Select Case (iRc)
            Case 96:  'not send
                bReverseCWC = False
                PrjString = "  **" + CStr(iRc) + ": Not send to ATMP" + vbCrLf
                PrjCHNString = "  " + CStr(iRc) + ":未发送到主机" + vbCrLf
            Case 90:  ' in use
                bReverseCWC = False
                PrjString = "  **" + CStr(iRc) + ": Line In Use" + vbCrLf
                PrjCHNString = " " + CStr(iRc) + ": 线路正在使用中" + vbCrLf
            Case Else
                PrjString = "  **" + CStr(iRc) + ": Line Status Unknown." + vbCrLf
                PrjCHNString = " " + CStr(iRc) + ":线路状态未知" + vbCrLf
        End Select
        
        If bReverseCWC Then
            PrjString = PrjString + "  NO RESPONSE FROM ATMP. " + vbCrLf + "  ** Need check this transaction" + vbCrLf
            PrjCHNString = "  ** 未收到主机响应." + vbCrLf + "  **需手工对账" + vbCrLf
            Call PrintJournalMedia(PrjString, PrjCHNString)
            
            Pcb3Dl.DlSetCharRaw "GBLCWDResult", "CS"
            iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "U")
            g_bIsWthReversal = True
        End If
        Call SendAEXMessage("2013")
        CommunicationSubFunction = UserReplyCommErr
        Exit Function
    End If
        
    iRc = SDOLineOut.DoReceive
    g_HostCurrentDate = Format(Now(), "MMDD")
    g_HostCurrentTime = Format(Now(), "HHMM")
    
    Select Case iRc
        Case 0:
            HostResLineNum = Pcb3Dl.DlGetCharRaw("HostLineNum")
            
            If Not IsNumeric(HostResLineNum) Then
                PrjString = "  **  Received HostLine Error "
                PrjCHNString = "  主机返回流水号有误" + vbCrLf
                Call PrintJournalMedia(PrjString, PrjCHNString)
               
                Pcb3Dl.DlSetCharRaw "GBLCWDResult", "CU"
                Call SendAEXMessage("2013")
                iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "U")
                g_bIsWthReversal = True
                CommunicationSubFunction = UserReplyCommErr
                Exit Function
            End If
            
            If CDbl(HostResLineNum) = CDbl(g_sDetailLineNum) Then
                g_sRejectCode = Pcb3Dl.DlGetCharRaw("HostTransCode")
                Select Case g_sRejectCode
                    Case "AQP", "DQP":
                        HostResAmount = Pcb3Dl.DlGetCharRaw("HostTransAmount")
                    
                        If IsNumeric(HostResAmount) Then
                            If CDbl(StrWthAmount) = CDbl(HostResAmount) Then
                        
                                HostAccNo = Pcb3Dl.DlGetCharRaw("HostAccNo")
                                HostSeq = Pcb3Dl.DlGetCharRaw("IcbcHostSeq")
                               
                                Pcb3Dl.DlSetCharRaw "FitPrrAccNo", _
                                Left(HostAccNo, Len(HostAccNo) - 5) + "****" + Right(HostAccNo, 1)
                             
                                PrjString = "     HOST ACCEPT " + vbCrLf + _
                                         "     Host AccNo: " + HostAccNo + _
                                         "     host CardMark: " + Pcb3Dl.DlGetCharRaw("FitCardMark") + vbCrLf + _
                                         "     Host Date: " + Pcb3Dl.DlGetCharRaw("HostCurrentDate") + vbCrLf + _
                                         "     Host Seq :" + HostSeq + vbCrLf
                            
                                PrjCHNString = "    主　机　接　受 " + vbCrLf + _
                                                     "     主机返回帐号：" + HostAccNo + _
                                                     "     主机时间：" + Pcb3Dl.DlGetCharRaw("HostCurrentDate") + vbCrLf + _
                                                     "     主机检索号：" + HostSeq + vbCrLf
                                Call PrintJournalMedia(PrjString, PrjCHNString)
                                                     
                                iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "Y")
                                CommunicationSubFunction = 0
                            Else        'StrWthAmount <>HostResAmount
                                PrjString = "  **  Received Amount <> Send Amount"
                                PrjCHNString = "  返回金额与上送不一致" + vbCrLf
                                Call PrintJournalMedia(PrjString, PrjCHNString)
                                Pcb3Dl.DlSetCharRaw "GBLCWDResult", "CU"
                                Call SendAEXMessage("2013")
                                iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "U")
                                g_bIsWthReversal = True
                                CommunicationSubFunction = UserReplyCommErr
                            End If
                        Else        'HostResAmount not numeric
                            PrjString = " **   Received Amount Erlaotror "
                            PrjCHNString = "  主机返回金额有误" + vbCrLf
                            Call SendAEXMessage("2013")
                            Call PrintJournalMedia(PrjString, PrjCHNString)
                            Pcb3Dl.DlSetCharRaw "GBLCWDResult", "CU"
                            iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "U")
                            g_bIsWthReversal = True
                            CommunicationSubFunction = UserReplyCommErr
                        End If
                    Case "ATP", "DTP", "AUP", "DUP":
                        iRc = SDOLineOut.GetData("constHostRejectCode", g_sRejectCode)
                        iRc = Pcb3Dl.DlSetCharRaw("ATMPRejectCode", g_sRejectCode)
                        CommunicationSubFunction = UserReplyHostReject
                    'need add AUP 2005/12/26
                        
                    Case Else:
                        PrjString = "  **  Received Unknown Code " & g_sRejectCode
                        PrjCHNString = "  收到其他拒绝码" & g_sRejectCode + vbCrLf
                        Pcb3Dl.DlSetCharRaw "GBLCWDResult", "CU"
                        Call PrintJournalMedia(PrjString, PrjCHNString)
                        iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "U")
                        g_bIsWthReversal = True
                        CommunicationSubFunction = UserReplyCommErr
                End Select
            Else            'HostResLineNum <>g_sDetailLineNum
                PrjString = "  **  Received LineNo <> Send LineNo"
                PrjCHNString = "  返回流水号与上送不一致" + vbCrLf
                Call PrintJournalMedia(PrjString, PrjCHNString)
                Pcb3Dl.DlSetCharRaw "GBLCWDResult", "CU"
                Call SendAEXMessage("2013")
                iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "U")
                g_bIsWthReversal = True
                CommunicationSubFunction = UserReplyCommErr
            End If
        
        Case 97:
            LogError "DoReceive return 97,host return MAC error"
            PrjString = "  ** Receiving host return MAC error" + vbCrLf
            PrjCHNString = "  MAC校验失败" + vbCrLf
            Call PrintJournalMedia(PrjString, PrjCHNString)
            iRc = Pcb3Dl.DlSetCharRaw("ResetTransKey", "R")
            CommunicationSubFunction = UserReplyCommErr
        Case Else:
'            HostSeq = "00000000000000000000000"
'            iRc = SendCwdReversal("4001")
'            Call CwdReversalTotal
            PrjString = "  Receiving host response error" + vbCrLf + "  ** Need check this transaction " + vbCrLf
            PrjCHNString = "  通讯故障." + vbCrLf + "  **需手工对账" + vbCrLf
            Call PrintJournalMedia(PrjString, PrjCHNString)
            iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "U")
            g_bIsWthReversal = True
            Pcb3Dl.DlSetCharRaw "GBLCWDResult", "CE"
            CommunicationSubFunction = UserReplyCommErr
    End Select
End Function
'==========================================================================================
'函 数 的 功 能 :记录数据库
'输 入 参 数   ：无
'输 出 参 数   ：无
'返 回 值      ：无
'调 用 函 数   ：无
'被 调 用 情 况：
'作 者         ：孙世方
'创 建 时 间   : 2005.6.23
'--------数据库内容------------
'记帐标志： Y 已记帐  N 未记帐 U 待查
'错帐原因：见解释
'主机拒绝码：仅用于特殊主机拒绝码需要收钞
'交易类型：存款 取款
'卡类型标志
'主机时间
'交易帐号
'交易金额
'交易流水号： 5 位
'==========================================================================================
Private Sub RecordDB_CWDLog()
    Dim sHostTransDate     As String
    Dim sCardType          As String
    Dim sHostRejectCode    As String
    Dim sLocalRejCode      As String
    Dim sKeepAccountFlag   As String
    Dim lTransAmount       As Long
    Dim strLineNum         As String
    Dim AccNo              As String
    Dim szMsg              As String
    
On Error GoTo ErrHandler
    sHostTransDate = Right(g_HostCurrentDate, 4) + g_HostCurrentTime
    If Len(sHostTransDate) = 0 Then
      sHostTransDate = Format(Now(), "MMDDHHMM")
    End If
    
    AccNo = Format(Pcb3Dl.DlGetCharRaw("FitAccNo"), "00000000000000000000")
    sLocalRejCode = Pcb3Dl.DlGetCharRaw("GBLCWDResult")
    sKeepAccountFlag = Pcb3Dl.DlGetCharRaw("GBLKeepAccountFlag")
    sCardType = Pcb3Dl.DlGetCharRaw("FitCardType")
    lTransAmount = CLng(Pcb3Dl.DlGetCharRaw("GBLPrtAmount"))
    sHostRejectCode = Pcb3Dl.DlGetCharRaw("ATMPRejectCode")
    strLineNum = Format(Pcb3Dl.DlGetCharRaw("GBLLineSendNum"), "000000")
    
    If DataWTH.Recordset.RecordCount <> 0 Then
       DataWTH.Recordset.MoveLast
    End If
    
    With DataWTH.Recordset
        .AddNew
        !TransType = "CWD"
        !TransDate = sHostTransDate
        !TransCardType = sCardType
        !TransAmount = lTransAmount
        !TransAccNo = AccNo
        !TransSerial = strLineNum
        !KeepAccountFlag = sKeepAccountFlag
        !AccountErrorReason = sLocalRejCode
        !HostRejectCode = sHostRejectCode
        .Update
    End With
    Exit Sub
    
ErrHandler:
   iRc = ErrorHandlerFunction("RecordDB_CWDLog:", 99)
End Sub
'================================================================================
'函数功能 :打印流水
'输入参数 ：流水打印buffer
'输出参数：无
'返回值：无
'调用函数：无
'被调用情况：打印流水时调用
'作者：孙世方
'创建时间 : 2005.6.22
'================================================================================
Sub PrintJournalMedia(ByRef JournalBuf As String, ByRef CHNJournalBuf As String)
    If (Len(JournalBuf) <> 0) And (Len(CHNJournalBuf) <> 0) Then
        PrintJournal SDOPrj, JournalBuf, CHNJournalBuf, g_sPrjLanguage
    End If
End Sub
'===================================================================================
'函数功能 :记录冲正统计值
'输入参数 ：无
'输出参数：无
'返回值：无
'调用函数：无
'被调用情况：
'作者：
'创建时间 : 2004
'====================================================================================
Private Sub CwdReversalTotal()
    Dim WthAmount         As Long
    Dim StrWthAmount      As String
    Dim nTotWthNum        As Long
    Dim dblTotWthAmt      As Long
        
    StrWthAmount = Pcb3Dl.DlGetCharRaw("GBLAmount") \ 100
    WthAmount = CLng(StrWthAmount)
    
    If g_bIsLocalBankCard = True Then
        nTotWthNum = Pcb3Dl.DlGetInt("TotWthReversalNum")
        dblTotWthAmt = Pcb3Dl.DlGetDouble("TotWthReversalAmount")
        nTotWthNum = nTotWthNum + 1
        dblTotWthAmt = dblTotWthAmt + WthAmount
        iRc = Pcb3Dl.DlSetLong("TotWthReversalNum", nTotWthNum)
        iRc = Pcb3Dl.DlSetDouble("TotWthReversalAmount", dblTotWthAmt)
    Else
        nTotWthNum = Pcb3Dl.DlGetInt("IcbcTotExtraWthRevNum")
        dblTotWthAmt = Pcb3Dl.DlGetDouble("IcbcTotExtraWthRevAmount")
        nTotWthNum = nTotWthNum + 1
        dblTotWthAmt = dblTotWthAmt + WthAmount
        iRc = Pcb3Dl.DlSetLong("IcbcTotExtraWthRevNum", nTotWthNum)
        iRc = Pcb3Dl.DlSetDouble("IcbcTotExtraWthRevAmount", dblTotWthAmt)
    End If
End Sub
'===================================================================================
'函数功能 :取款成功记录统计值、显示画面、准备收条内容
'输入参数 ：无
'输出参数：无
'返回值：无
'调用函数：
'被调用情况：AtWithdrEnd = 0 ；95:取款超时，但回收钞票张数为零时 ；99:IsNotesPresented = True
'作者：孙世方
'创建时间 : 2005.8.23
'====================================================================================
Private Sub CwdOkFunction()
    Dim sSubStData     As String
    
    Call CwdOkTotal
    Call PrintCassLeftNum
    
    iRc = ShowScreenSync(Browser, "Cwd", "CwdTransOk", sSubStData)

    '准备收条打印内容 取款成功
    Call DrawWthPrr(Pcb3Dl, WthPrrOK)
End Sub
'===================================================================================
'函数功能 :取款可能出现犯罪情况时记录数据库、交易日志，显示画面，准备收条内容
'输入参数 ：无
'输出参数：无
'返回值：无
'调用函数：
'被调用情况：95 optype_cdm_shutterproblem;99 IsNotesPresented = True and optype_cdm_shutterproblem
'作者：孙世方
'创建时间 : 2005.8.23
'--------------------------
'<时间>：[2005.11.1]
'<修改者>：孙世方
'<详细记录>：
'增加记录cwdoktotal 和打印钞箱情况PrintCassLeftNum
'====================================================================================
Private Sub CwdCrimeFunction()
    Dim sSubStData     As String
    Dim PrjString      As String
    Dim PrjCHNString   As String
    
    iRc = Pcb3Dl.DlSetCharRaw("CWDCrimePossible", "Y")
    Pcb3Dl.DlSetCharRaw "GBLDoRecovery", "C"
    PrjString = " TRANSACTION OK (Crime Possible)" + vbCrLf
    PrjCHNString = "**取款成功（但可能有人做案！）" + vbCrLf
    Call PrintJournalMedia(PrjString, PrjCHNString)

    '记录数据库、交易日志
    iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "CR")
    Call RecordDB_CWDLog
    Call WriteTranslog("可疑出钞")
    Call CwdOkTotal
    Call PrintCassLeftNum
    
    iRc = ShowScreenSync(Browser, "Cwd", "CwdCrime", sSubStData)

    '准备收条打印内容 取款成功
    Call DrawWthPrr(Pcb3Dl, WthPrrOK)
 End Sub
'===================================================================================
'函数功能 :记录取款交易日志（用于后屏显示）
'输入参数 ：无
'输出参数：无
'返回值：无
'调用函数：
'被调用情况:
'作者：汪林
'创建时间 : 2004
'====================================================================================
Private Sub WriteTranslog(TransState As String)
    Dim fso                As New FileSystemObject
    Dim TransLogStream     As TextStream
    Dim sTransLogRec       As String
    Dim szMsg              As String
    
 On Error GoTo ErrHandler
    If Not fso.FileExists(TransLogFile) Then
        Set TransLogStream = fso.CreateTextFile(TransLogFile)
    Else
        Set TransLogStream = fso.GetFile(TransLogFile).OpenAsTextStream(ForAppending)
    End If
    
    sTransLogRec = "取款|" + Format(Now(), "MM/DD|HH:MM|") + _
            Pcb3Dl.DlGetCharRaw("FitAccNo") + "|" + Pcb3Dl.DlGetCharRaw("GBLLineSendNum") + _
            "|" + Format(CLng(Pcb3Dl.DlGetCharRaw("GBLAmount") \ 100), "standard") + "|" + _
            TransState
            
    TransLogStream.WriteLine sTransLogRec
    TransLogStream.Close
    Exit Sub
    
ErrHandler:
   iRc = ErrorHandlerFunction("WriteTranslog:", 99)
End Sub
'==========================================================================================
'版本号：Agilis 1.6
'参见sdohelp文件DoEjectCard方法，针对其中4个事件进行处理
'==========================================================================================
Private Sub SDOIdc_AtEjectStart()
    Call LogInfo("SDOIdc_AtEjectStart")
    SDOIdc.UserReply = 0
    iRc = SDOFep.SetIndicator(ind_audio, audio_continuous + BeepEffect_WFS_SIU_ON)
End Sub
Private Sub SDOIdc_EjectCardTimeOut()
    Dim sSubStData     As String
    
    Call LogInfo("SDOIdc_EjectCardTimeOut")
    iRc = ShowScreenSync(Browser, "EndVisit", "TakeCardwarning", sSubStData)
    SDOIdc.UserReply = 0
    
End Sub
Private Sub SDOIdc_CardWillBeCaptured()
    Call LogInfo("SDOIdc_CardWillBeCaptured")
    SDOIdc.UserReply = 0
End Sub
Private Sub SDOIdc_AtEjectEnd(ByVal rcEjectCard As Integer)
    Dim PrjString           As String
    Dim PrjCHNString        As String
    
    iRc = SDOFep.SetIndicator(ind_audio, audio_off)
    If rcEjectCard = rcdevidc_doejectcc_cccaptured Then
        
        PrjString = "   **TimeOut:card not taken by client"
        PrjCHNString = "   **超时：客户未取卡" + vbCrLf
        Call PrintJournalMedia(PrjString, PrjCHNString)
        
        '记录吞卡文件
        Call RecordCpdCardLog("1035")
        
        '打印吞卡收条
        Call DrawCpdCardPrr(Pcb3Dl)

        LogWarning "SDOCdm_BefDeliver UserReply = UserReplyCardCapture at SDOIdc_AtEjectEnd"
        SDOCdm.UserReply = UserReplyCardCapture
        
    ElseIf rcEjectCard = 0 Then
        LogInfo "SDOCdm_BefDeliver UserReply = 0 at SDOIdc_AtEjectEnd"
        SDOCdm.UserReply = 0
    Else
        LogError "AtEjectEnd in CWD RC=" & rcEjectCard
        Call CaptureCard
'        Call PrintJournalMedia("", "", "SDOCdm_BefDeliver UserReply = UserReplyCardCapture", "")
        SDOCdm.UserReply = UserReplyCardCapture
    End If
End Sub
'===================================================================================
'函数功能 :吞卡处理：显示画面、打印流水、执行吞卡命令、准备吞卡收条
'输入参数 ：无
'输出参数：无
'返回值：无
'调用函数：DoTakeCard
'被调用情况: DoEjectCard <>0; SDOIdc_AtEjectEnd
'作者：孙世方
'创建时间 : 2005.8.23
'====================================================================================
Private Sub CaptureCard()
    Dim sSubStData     As String
    
    iRc = ShowScreenSync(Browser, "EndVisit", "EjectCardError", sSubStData)
    iRc = SDOFep.SetIndicator(ind_audio, audio_off)
    Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "24")
    
    PrintJournalMedia "   **Eject Card Err in CWD", "   **取款时退卡失败"
               
    iRc = SDOIdc.DoTakeCard 'capture the card
    If iRc <> 0 Then
        PrintJournalMedia "   **Capture Card Err in CWD.", "   **取款时吞卡失败"
    End If
    Call RecordCpdCardLog("1035")
    
    Call DrawCpdCardPrr(Pcb3Dl)
    
End Sub
'===================================================================================
'函数功能 :记录吞卡文件
'输入参数 ：导致吞卡的例外代码
'输出参数：无
'返回值：无
'调用函数：
'被调用情况： SDOIdc_AtEjectEnd   超时未拿卡;CaptureCard
'作者：
'创建时间 : 2004
'====================================================================================
Private Sub RecordCpdCardLog(ExceptCode As String)
    Dim sTime         As String
    Dim FullCardAccNo As String

    Pcb3Dl.DlSetLong "TotCapCardNum", Pcb3Dl.DlGetInt("TotCapCardNum") + 1
    
    sTime = Format(Now(), "YYYYMMDDHHMM")
    FullCardAccNo = Format(Pcb3Dl.DlGetCharRaw("FitAccNo"), "@@@@@@@@@@@@@@@@@@@@!")
        
    Open CardRetainFile For Append As #1
    Print #1, sTime + " " + FullCardAccNo + " " + ExceptCode
    Close #1
        
End Sub
'==========================================================================================
'函数功能 ：取款成功后，打印各取款钞箱剩钞张数
'输入参数：无
'输出参数：无
'返回值 ：无
'调用函数：
'被 调 用 情 况：
'作者：李军
'创建时间: 2005-08-29
'==========================================================================================
'<时间>：[2005.11.25]
'<修改者>:孙世方
'<详细记录>：
'注册表中cashunit数目在某种情况下会随着动钞箱而增长，SDOCdm.NbrOfBoxesUsed也会随之增长，
'之前的程序会出现CasLefNum越界的情况。
'==========================================================================================
Private Sub PrintCassLeftNum()
On Error GoTo ErrHandler
    Dim i                   As Integer
    Dim sCasState           As String
    Dim szMsg               As String
    Dim PrjString           As String
    Dim PrjCHNString        As String
    Dim CasPosition         As Integer
    Dim CasLefNum(1 To 6)   As String
    Dim j                   As Integer
    
    PrjString = ""
    PrjCHNString = ""
    
    For i = 1 To 6
        CasLefNum(i) = ""
    Next i
    
    j = 1
    SDOCdm.DataCriteria = 1
    For i = 1 To SDOCdm.NbrOfBoxesUsed
        SDOCdm.CasNbrLogical = i
        If Len(SDOCdm.CasPosition) <> 0 Then
            If IsNumeric(Right(SDOCdm.CasPosition, 1)) Then
                CasPosition = CInt(Right(SDOCdm.CasPosition, 1))
                If SDOCdm.CasState >= casstate_cdm_ok And SDOCdm.CasState <= casstate_cdm_empty Then
                    CasLefNum(j) = Format(CStr(SDOCdm.TotNbrPresent), "0000")
                    j = j + 1
                End If
            End If
        End If
    Next i
    
    For i = 1 To 5
        If Len(CasLefNum(i)) > 0 Then
            PrjString = PrjString + "BIN" + CStr(i) + ": " + CasLefNum(i) + " "
            PrjCHNString = PrjCHNString + "钞箱" + CStr(i) + ": " + CasLefNum(i) + " "
        End If
    Next i
    
    PrjCHNString = PrjCHNString + vbCrLf
    PrintJournalMedia PrjString, PrjCHNString
    
    Exit Sub
    
ErrHandler:
    szMsg = CStr(Err.Number) + ": " + Err.Description + " in PrintCassLeftNum"
    LogError szMsg
    Err.Clear
    Exit Sub
End Sub

'===================================================================================
'函数功能 :发送AEX例外通讯报文
'输入参数 ：例外代码
'输出参数：无
'返回值：无
'调用函数：无
'被调用情况：
'作者：孙世方
'创建时间 : 2005.11.10
'====================================================================================
Sub SendAEXMessage(ByVal ExpCode As String)
    Dim sCurrentDate       As String
    Dim nrc                As Integer
    
    sCurrentDate = Format(Now(), "MMDDHHMMSS")
    SDOLineOut.SetData "CurrentDate", sCurrentDate
    
    nrc = SDOLineOut.SetData("ExceptionCode", ExpCode)
    
    nrc = SDOLineOut.DoSend("AEX", 0)
    
End Sub
'===================================================================================
'函数功能 :检查例外代码是否与文件中的相同
'输入参数 ：无
'输出参数：无
'返回值：
'调用函数：无
'被调用情况：
'作者：孙世方
'创建时间 : 2005.11.10
'====================================================================================
Private Function CheckReversalPossible() As Boolean
    Dim bAnomaliesLeft As Boolean
    Dim stTime As Date
    Dim nDevId As Integer, nTOId As Integer, nDOId As Integer
    Dim nWosaReply As Long
    Dim sSKBSReply As String, sDescr As String, sLogicalName As String, sOldDescr As String
    Dim fso As New FileSystemObject
    Dim AnomalyStream As TextStream
    Dim TextToPrint As String
    Dim DlVarName As String
    Dim sDevice As String
    Dim CheckResult As Boolean
    Dim PrjString                                   As String
    Dim PrjCHNString                                As String
    
    CheckResult = False

    LogInfo "Start GetAnomalies"
    If Not fso.FileExists("c:\S3E\Logs\LogTO\anomaly.txt") Then
        ' Create a new anomaly file
        LogInfo "Creating new anomaly.txt file"
        Set AnomalyStream = fso.CreateTextFile("c:\S3E\Logs\LogTO\anomaly.txt")
    Else
        ' Open the existing anomaly file for appending
        LogInfo "Opening existing anomaly.txt file"
        Set AnomalyStream = fso.GetFile("c:\S3E\Logs\LogTO\anomaly.txt").OpenAsTextStream(ForAppending)
    End If
    
    ' Retrieve anomalies
    LogInfo "Retrieving anomalies"
    bAnomaliesLeft = SDOCdm.GetAnomalyRaw(stTime, nDevId, nTOId, nDOId, nWosaReply, sSKBSReply, sDescr, sLogicalName)
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
            Case 9:  DlVarName = "SiabOPKCode"
                     sDevice = "OPS"
            Case 11: DlVarName = "SiabALMCode"
                     sDevice = "DOOR"
            Case 23: DlVarName = "SiabSCMCode"
                     sDevice = "SCM"
            Case 81: DlVarName = "SiabPRRCode"
                     sDevice = "PRR"
            Case 82: DlVarName = "SiabCIMCode"
                     sDevice = "CIM"
            Case 13: DlVarName = "SiabCDMCode"
                     sDevice = "CDM"
        End Select
        
        TextToPrint = Date$ & " " & Format(Time$, "HH:MM:SS") & _
                      " (" & sLogicalName & Space(12 - Len(sLogicalName)) & _
                      ") DEV " & Str(nDevId) & Space(3 - Len(CStr(nDevId))) & _
                      " TO " & Str(nTOId) & Space(4 - Len(CStr(nTOId))) & _
                      " DO " & Str(nDOId) & Space(4 - Len(CStr(nDOId))) & _
                      " WOSA " & Str(nWosaReply) & Space(5 - Len(CStr(nWosaReply))) & _
                      " SKBS " & sSKBSReply & Space(5 - Len(sSKBSReply)) & sDescr
        AnomalyStream.WriteLine TextToPrint
       
        
        If nWosaReply = 0 And sOldDescr <> sDescr Then
            PrjString = "ANOM " & Date$ & " " & Time$ & Space(24 - Len(sDevice)) & sDevice & Chr(13) & Chr(10) & _
                          "    SP Info: " & sDescr
            PrjCHNString = PrjString
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            sOldDescr = sDescr
            LogInfo PrjString
        End If
        
        If (Not CheckSPInfo(sDescr, "NotRecoveryCode")) Then
             'Restore the information to journal printer & LOG
            PrjString = "*** " & Date$ & " " & Time$ & " ***" & vbCrLf & _
                          "*** A severe CDM hardware fault happened!!! ***"
            PrjCHNString = " 吐钞机故障，请检查传输通道是否有卡钞！！" + vbCrLf
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            LogError PrjString
            iRc = Pcb3Dl.DlSetCharRaw("GBLCdmRecoveryNeeded", "N")
        End If

        If sDevice = "CDM" And (Not CheckSPInfo(sDescr, "FloatingCode")) Then     '例外码与文件中的相同
            CheckResult = True
            bAnomaliesLeft = False
        Else
            bAnomaliesLeft = SDOCdm.GetAnomalyRaw(stTime, nDevId, nTOId, nDOId, nWosaReply, sSKBSReply, sDescr, sLogicalName)
        End If
    Wend
    
    LogInfo "No more anomalies"
    AnomalyStream.Close
    LogInfo "End GetAnomalies"
    
    CheckReversalPossible = CheckResult
    
End Function


