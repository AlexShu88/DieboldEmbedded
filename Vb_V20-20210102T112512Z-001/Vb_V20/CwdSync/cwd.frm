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
'��Ȩ˵��:  �ϱ���˾�й���������
'�汾��Դ��ԭ���а�Optevaȡ�������Timer��ʽ���汾��1.3.16
'�汾�ţ� 1.1.0.16 (2005.11.10)
'�������ڣ�2005.8
'���ߣ�  ������
'ģ�鹦�ܣ� ȡ������
'��Ҫ�������书��
' ȫ�ֱ���
'    HtmlPrompt1,HtmlPrompt2,HtmlPrompt3,HtmlPrompt4 : 10,20,50,100 ��ֵ�ܽ�������Ļ��ʾ
'             GBLCWDResult                   ȡ�����
'            Ӳ��ԭ��  DF    ȡ��㳮ʧ��       ��������
'                       ST    ����ʧ��           ��������
'                       GB    ȡ���������   �Ѽ���,�޷��ж��Ƿ����������
'                       PF    RevFloating       �Ѽ���,�޷��ж��Ƿ����������
'            ����ԭ��  PT    ��ʱδ��          �Ѽ���
'                       PO    ��ʱδ�õ����ճ�Ʊ����Ϊ��  �Ѽ���
'                       CR      ����              �Ѽ���
'                       CC     ������            ��������
'            ͨѶԭ��  CS    dosend<>0         ���ֹ�����
'                       CE    doreceive<>0      ���ֹ�����
'                       CU    �Ƕ��巵����       ���ֹ�����
'            GBLKeepAccountFlag      ȡ����ʱ�־ ��ֵ�� N
'                     �յ�����ȷ�ϣ��Ѽ���   Y
'                      �����ɹ�             R
'                     ͨѶԭ������ʧ��     U     ���ֹ�����
'       �����޸ģ�
'         1 �������Ȼ�����dowithdrawal
'         2 ��ȡ�������в��ٵ���timer
'==========================================================================================
'<ʱ��>��[2005.08.22]
'<�޸���>��������
'<��ǰ�汾>��1.0.16(���а棩
'<��ϸ��¼>��
'   ��3030ȡ����ֲ��Optevaȡ���ֲ����ͳ��ֵ����  cwdoktotal sendexception CwdReversalTotal
'   ɾ��cwdlog����������,����WriteTranslog������
'   ɾ��3030���к�����RecordTakenNotesTimeOut��CassetteNotesChangeVerify��
'                    GetCashUnitsTotal , GetCashUnitsType
'   ɾ������g_sCassettesNotesDetail�����ڼ�¼billbox
'
'   �糮��(��ͬ��ֵ)ȡ��
'     ͨ����XFS�С�DispByPosition���͡�Count Control dispensing���Ĳ�������ΪYES������ʵ�����¹��ܣ�
'    1���糮��(��ͬ��ֵ)ȡ��
'    2��ָ��ÿ�����䣨��ʹ��ͬ���ĳ��䣩����һ��Ǯ
'    3��ÿ������NoteToDispense������׼ȷ��
'
' Ϊ�������������˿�����������̣�����IDC�����ӹ����˿��Ĵ�����SDOCdm_BefDeliver�ж����ò�����
' ��Atmconfig�����˿���ʽ�����˿�"A" - After,"B" - Before  GBLTakeCardSequence
'==========================================================================================
'<ʱ��>��[2005.10.22]
'<�޸���>��������
'<��ϸ��¼>��
'      ������ȡ��ǰ���ӷ���ERI����ʾ�һ���
'<ʱ��>��[2005.11.1]
'<�޸���>��������
'<��ϸ��¼>��
'    1 ��CwdCrimeFunction���������Ӽ�¼cwdoktotal �ʹ�ӡ�������PrintCassLeftNum
'    2 RevStateFloatingʱ���Ӽ�¼cwdoktotal �ʹ�ӡ�������PrintCassLeftNum
'==========================================================================================
'<ʱ��>��[2005.11.9]
'<�޸���>��������
'<��ϸ��¼>���޸���ˮ��¼����
'==========================================================================================
'<ʱ��>��[2005.11.10]
'<�޸���>��������
'<��ϸ��¼>��
' 1 ����������ˮ���뷢�Ͳ�ͬʱ�����ͳ���
' 2 �����޸���ˮ����
' 3 ���ӷ���AEX����
'       2009  �ͻ�������ʱ
'       2012  �ͻ�δȡ��
'       2013  ����������ˮ�Ż������󣬻������Ͳ�һ��
'       2015  �ͻ���ʱδȡ��Ʊ
'       2036  �ͻ�ȡ������
'4 �����᷵��AUP,����ATPһ�������㶫�з���DQP,DTP,DUP
'==========================================================================================
'<ʱ��>��[2005.11.22]
'<�޸���>:vincent
'<��ϸ��¼>��
'   1 ֻ�е㳮ʧ��ʱ���ͳ�����ͨѶʧ�ܲ�����
'   2 �޸ĳ�����ʽ��ֱ�ӷ�,���ڶ���
'==========================================================================================
'<ʱ��>��[2005.11.25]
'<�޸���>:������
'<��ϸ��¼>��
'ע�����cashunit��Ŀ��ĳ������»����Ŷ������������SDOCdm.NbrOfBoxesUsedҲ����֮������
'֮ǰ�ĳ�������CasLefNumԽ���������޸�PrintCassLeftNum�������������⡣
'==========================================================================================
'<ʱ��>��[2005.12.12]
'<�޸���>��������
'�汾�ţ� 1.2.16 (2005.12.12)
'<��ϸ��¼>�� �޸�CwdOkTotal, ���Ӽ�¼CutOff.ini����������cutoffʱ��ӡ��ˮ
'<ʱ��>��[2005.12.20]
'�汾�ţ� 1.3.16
'<��ϸ��¼>: �޸�SDOCdm_AtWithdrEnd��99 else�������Ӽ�¼translog
'   ɾ��ResetATMPrr�����ĵ��ã����ӵ���DrawWthPrr��ӡ�ܾ�����
'<ʱ��>��[2005.12.27]
'�汾�ţ� 1.4.16
'<��ϸ��¼>:
'   1 �޸�SendCwdReversal���������ӶԳ���������ж�
'   2 ����ST���ף�����
'   3 ����CheckReversalPossible,��ReversalState=0ʱ��������������MMDCode.ini�����ݵĶԱȣ������Ƿ�ͣ�����Ƿ����
'   4 �޸�CwdReversalTotal����,�������\100
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
'�� �� �� �� �� ��VB����װ��,��ʼ�����ܼ��������ݿ�
'�� �� �� ��   ����
'�� �� �� ��   ����
'�� �� ֵ      ����
'�� ��
'�� �� ʱ ��   :
'==========================================================================================
Private Sub Form_Load()

    ' Reset the PcB3HtmlBrowser variables
    iRc = Pcb3Dl.DlSetCharRaw("HtmlFkeyList", "")
    iRc = Pcb3Dl.DlSetCharRaw("HtmlFkeyMap", "3855")
    
    '������ϸ���ݿ⣬���ڽ���ʱ��ѯ����ӡ
    DataWTH.DatabaseName = DB_WthLogPath
    DataWTH.RecordSource = "Select * from CWDLOG"
    DataWTH.Refresh
    
    If GetIniS(sGlobalIni, "Bank_Environment", "PrjLanguage", "E") = "E" Then
        g_sPrjLanguage = "E"
    Else
        g_sPrjLanguage = "C"
    End If
     
    '�����������˿�˳��  A ���˿���B ���˿�
    TakeCardSenqueue = GetIniS(sGlobalIni, "Bank_Environment", _
                        "EjectCardMode", "A")
    
    SDOTrans.Available = True
    Cwd.WindowState = 1
End Sub
'==========================================================================================
'�� �� �� �� �� ��ȡ������˳�
'�� �� �� ��   ����
'�� �� �� ��   ����
'�� �� ֵ      ����
'�� ��
'�� �� ʱ ��   :
'==========================================================================================
Private Sub SDOTrans_QuitTransaction()
    Unload Cwd
End Sub
'==========================================================================================
'�������� ����ʼ�����ݣ��õ���ȡ����ص���Ϣ
'�����������
'�����������
'����ֵ ���������� true - ����ȡ��  false-����ȡ��
'���ú�����
'  GetWTHCassettesTotal :  �õ���ǰ���䳮Ʊ������(����ȡ���)
'  GetAllCasDenominations: �õ���ǰ���ó��������С��ֵ����ǰ�������ȡ���޶�
'�� �� �� �� ����SDOTrans_StartTransaction
'���ߣ�����
'����ʱ��: 2005-08-05
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
'�� �� �� �� �� ��ȡ��������
'�� �� �� ��   �� ��ruler���ô�ģ��ʱActionֵ
'�� �� �� ��   ����
'�� �� ֵ      ����
'�� ��
'�� �� ʱ ��   :
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
    
    '����������CardType=03  �� CardType =04 ��Ҫ�Ȳ�ѯ�һ���
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
    
    '���������Ǳ��п���CardType =01
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
            PrjCHNString = "  �޿���ȡ��� " + vbCrLf
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
            PrjCHNString = " ������Ϊ: " + InputAmount + vbCrLf
            Call PrintJournalMedia(PrjString, PrjCHNString)
        Case 2:
            PrjString = "    Keyboard input timeout in Withdrawal"
            PrjCHNString = " �ͻ��������볬ʱ " + vbCrLf
            Call PrintJournalMedia(PrjString, PrjCHNString)
            
            Call SendAEXMessage("2009")
            RetToMaster ReturnTimeout
            Exit Sub
        Case Else:                 ' ȡ�� 1  ������� 3
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
        PrjCHNString = " �ͻ�ȡ��ش��� RC=" & CStr(iRc) + vbCrLf
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
        PrjCHNString = "    **�ͻ�ȡ���ʱ���볬ʱ"
        'add by nicktan add "*"
        Call PrintJournalMedia(PrjString, PrjCHNString)
        
        Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "46")
        RetToMaster ReturnTimeout
        Exit Sub
    Case Else:
        PrjString = "  **  Customer Exit in Cwd"
        PrjCHNString = "    **�ͻ���ȡ���ʱѡ���˳�"
         'add by nicktan add "*"
        Call PrintJournalMedia(PrjString, PrjCHNString)
        
        Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "45")
        iRc = ShowScreenSync(Browser, "Common", "ComPressStop", sSubStData)
        RetToMaster ReturnPressStop
        Exit Sub
    End Select
        
End Sub
'==========================================================================================
'�汾�ţ�Agilis 1.6
'�μ�sdohelp�ļ�DoWithdrawal������������оŸ��¼����д���
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
'����������������ܾ��󣬽����ٴ��������룬Ȼ��ֱ�ӽ���ȡ������ٴ�������
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
' ע�ͣ���Opteva�ϣ��糮��ʱƽ̨�ṩ��ÿ������㳮��������׼ȷ��
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
    
    '���˿�
    If TakeCardSenqueue = "A" Then
        Call LogInfo("SDOCdm_BefDeliver=0")
        SDOCdm.UserReply = 0
    Else
        '���˿�
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
        PrjCHNString = "�������׳ɹ����" + vbCrLf
        Call PrintJournalMedia(PrjString, PrjCHNString)
        '��¼���ݿ⡢������־
        iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "OK")
        Call RecordDB_CWDLog
        Call WriteTranslog("ȡ��ɹ�")
        
        Call CwdOkFunction
        nCwdReturnValue = ReturnOk
    
    Case UserReplyHostReject:
        iRc = ShowScreenSync(Browser, "Common", "ComReject", sSubStData)
        PrjString = "   **HOST REJECT [" + g_sRejectCode + "]" + vbCrLf + _
                                   "   **" + Pcb3Dl.DlGetCharRaw("HostRejectEnglish")
                
        PrjCHNString = "   **�����ܾ� [" + g_sRejectCode + "]" + vbCrLf + _
                               "   **" + Pcb3Dl.DlGetCharRaw("HostRejectChinese") + vbCrLf
        Call PrintJournalMedia(PrjString, PrjCHNString)
        
        sHostRejectCard = Pcb3Dl.DlGetCharRaw("HostRejectCard")
                
        If sHostRejectCard = "R" Then
            nCwdReturnValue = ReturnPinNotMatch
        Else
            Call DrawWthPrr(Pcb3Dl, WthPrrReject)   '2005.12.21��ӡ�ܾ�����
            nCwdReturnValue = ReturnHostReject
        End If
        
    Case UserReplyCommErr:
        Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "64")
        
        '���Ӽ�¼�ļ� 2005.11.14
        Call WriteTranslog("ͨ��ʧ��")
        
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
        PrjCHNString = "  ȡ�����" + vbCrLf + "  ����ԭ�򣺿ͻ�δȡ��." + vbCrLf
        Call PrintJournalMedia(PrjString, PrjCHNString)
        
        '��¼���ݿ⣬������־
        iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "CC")
        Call RecordDB_CWDLog
        Call WriteTranslog("��Ƭ����")
        
        nCwdReturnValue = ReturnCardCapture
             
    Case rcDEVCDM_DOWTH_T_O_ON_TAKE_NOTES:
        If SDOCdm.Available = False And SDOCdm.OperatorType = optype_cdm_shutterproblem Then
            
            Call CwdCrimeFunction
            
            nCwdReturnValue = ReturnPossibleCrime
        Else
            If g_nNumOfRetractCount = SDOCdm.RetractCount Then
               
                PrjString = "  ** Withdrawl Timeout, But Retract Failed!" + vbCrLf
                PrjCHNString = "  ȡ�ʱ�������ճ�Ʊ����Ϊ��" + vbCrLf
                Call PrintJournalMedia(PrjString, PrjCHNString)
                
                '��¼���ݿ⡢������־
                iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "PO")
                Call RecordDB_CWDLog
                Call WriteTranslog("ȡ��ɹ�")
                
                Call CwdOkFunction
                nCwdReturnValue = ReturnOk
            Else
                '��ӡ��ˮ
                PrjString = "** Take notes timeout"
                PrjCHNString = "  �ͻ�δȡ��" + vbCrLf
                Call PrintJournalMedia(PrjString, PrjCHNString)
                 
                '��¼ȡ��ͳ��ֵ
                Call CwdOkTotal
                 
                '��¼���ݿ⣬������־
                iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "PT")
                Call RecordDB_CWDLog
                Call WriteTranslog("ȡ����ʱ")
                
                '����������Ϣ
                Call SendAEXMessage("2015")
                
                iRc = ShowScreenSync(Browser, "Cwd", "CwdTakeNoteTimeout", sSubStData)
                
                 '׼��������ӡ���� ��ʱδȡ��
                Call DrawWthPrr(Pcb3Dl, WthPrrTimeout)
                nCwdReturnValue = ReturnNoteNotTaken
            End If
        End If
           
    Case Else:                  '98,99 ������
        LogError "SDOCdm_AtWithdrEnd's WithdrRc = " + CStr(WithdrRc)
        LogError "sdocdm.reversalstate = " + CStr(SDOCdm.ReversalState)
        
       Select Case (SDOCdm.ReversalState)
        Case RevStateNeeded:
            '��ӡ��ˮ
            PrjString = "** Cash dispenser error RC=" & CStr(WithdrRc) + " Need Reverse"
            PrjCHNString = "  �㳮ʱȡ��ģ����� RC=" & CStr(WithdrRc) + " ��Ҫ����" + vbCrLf
            Call PrintJournalMedia(PrjString, PrjCHNString)
            
            '��¼���ݿ⣬������־
            iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "DF")
          
            Call WriteTranslog("����ʧ��")
            
            Select Case SDOCdm.OperatorType
            Case optype_cdm_somecasslow, optype_cdm_casnotconfigured, optype_cdm_notesproblem, optype_cdm_casinvalid
                sCorrCode = "4007"
            Case optype_cdm_allempty
                sCorrCode = "4012"
            Case Else
                sCorrCode = "4009"
            End Select
                        
            '��ӡ��ˮ
            PrjString = "    Send Reversal , code =" & sCorrCode
            PrjCHNString = "    ���ͳ���,����=" & sCorrCode + vbCrLf
            Call PrintJournalMedia(PrjString, PrjCHNString)
            
             '���ͳ�����������Ϣ
            iRc = SendCwdReversal(sCorrCode)
            Call CwdReversalTotal
            Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "41")
            Call RecordDB_CWDLog
            
            iRc = ShowScreenSync(Browser, "Cwd", "CwdFailed", sSubStData)
            
            '׼��������ӡ���� ����
            Call DrawWthPrr(Pcb3Dl, WthPrrCWC)
            nCwdReturnValue = ReturnCwdReverse
            
        Case RevStateFloating:
            '��ӡ��ˮ
            PrjString = "** Cash presenter error RC=" & CStr(WithdrRc) + " No Reverse"
            PrjCHNString = "  �ͳ�ʱȡ��ģ����� RC=" & CStr(WithdrRc) + " δ����" + vbCrLf
            'modi by nicktan change the "#" to "*"
            Call PrintJournalMedia(PrjString, PrjCHNString)
            
             '��¼���ݿ⣬������־
            iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "PF")
            Call RecordDB_CWDLog
            Call WriteTranslog("�����쳣")
            Call CwdOkTotal
            Call PrintCassLeftNum
            
            '����������Ϣ
            Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "41")
            iRc = ShowScreenSync(Browser, "Cwd", "CwdRevFloat", sSubStData)

            '׼��������ӡ���� ȡ�����
            Call DrawWthPrr(Pcb3Dl, WthPrrFloat)
            nCwdReturnValue = ReturnRevFloat
        
        Case Else:
            If IsNotesPresented = True Then
                If (SDOCdm.Available = False And SDOCdm.OperatorType = optype_cdm_shutterproblem) Then
                    Call CwdCrimeFunction
                    RetToMaster ReturnPossibleCrime
                Else
                    Call PrintJournalMedia("   **  TRANSACTION OK**", "  ���׳ɹ����")
                     '��¼���ݿ⡢������־
                    iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "OK")
                    Call RecordDB_CWDLog
                    Call WriteTranslog("�������")
                    Call CwdOkFunction
                End If
            Else
                
                If (Not CheckReversalPossible) Then
                      '��ӡ��ˮ
                    PrjString = " Other reason Reversal"
                    PrjCHNString = " ����ԭ�����," + vbCrLf
                    Call PrintJournalMedia(PrjString, PrjCHNString)
                    
                    iRc = SendCwdReversal("4009")
                    Call CwdReversalTotal
                     '��¼���ݿ�
                    iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "ST")
                    Call RecordDB_CWDLog
                    Call WriteTranslog("��������")          ' ���Ӽ�¼2005��12��20
                Else
                     '��ӡ��ˮ
                    PrjString = "Need Check"
                    PrjCHNString = "���齻��" + vbCrLf
                    Call PrintJournalMedia(PrjString, PrjCHNString)
                    
                     '��¼���ݿ�
                    iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "GB")
                    Call RecordDB_CWDLog
                    Call WriteTranslog("ȡ�����")          ' ���Ӽ�¼2005��12��20
                End If
                
                '����������Ϣ
                Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "41")
       
                iRc = ShowScreenSync(Browser, "Cwd", "CwdFailed", sSubStData)
                nCwdReturnValue = ReturnFailed
            End If
        End Select              'ReversalState
    End Select
    
    RetToMaster nCwdReturnValue
    
End Sub
'==========================================================================================
'�������� ������ȡ����
'�����������
'�����������
'����ֵ ��
'          0: OK
'          1: Customer Cancel
'          2: Input Timeout
'          3: Error
'���ú�����
'�� �� �� �� ����SDOTrans_StartTransaction
'���ߣ�����
'����ʱ��: 2005-08-05
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
            sPrompt = sPrompt + CStr(CasDenomArray(i)) + "��"
        End If
    Next
    Pcb3Dl.DlSetCharRaw "HtmlPrompt2", sPrompt
    
    'ѡ���������������뻭��
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
                    InputWithdrawalAmount = 0   'ȷ��0
                    iRc = ShowScreenSync(Browser, "Common", "ComPlsWait", sSubStData)
                    Exit Function
                End If
            Case "@Change":
               
            Case Else
                InputWithdrawalAmount = 1   'ȡ�� 1
                Exit Function
            End Select
        Case 91:
            InputWithdrawalAmount = 2    '��ʱ2
            Exit Function
        Case Else:
            LogError ScreenInfo.Name + "Return error, irc = " + CStr(iRc)
            InputWithdrawalAmount = 3    '������� 3
            Exit Function
        End Select
    
    Wend

End Function
'==========================================================================================
'�������� ��ѡ�����ȡ���ȷ��
'�����������
'�����������
'����ֵ ��
'          0: OK
'          1: Customer Cancel
'          2: Input Timeout
'          3: Error
'          4��������
'���ú�����
'�� �� �� �� ����SDOTrans_StartTransaction
'���ߣ�������
'����ʱ��: 2005-09-07
'==========================================================================================
Private Function FastWithdrawalSelect() As Integer
    Dim FastWthAmount       As String
    Dim bLoop               As Boolean
    Dim sSubStData          As String
    
'����ȡ��ѡ��
    bLoop = True
    
    '��ҳ����ݵ�ǰ�������ȡ���������ο���ȡ����
    iRc = Pcb3Dl.DlSetCharRaw("HtmlPrompt1", CStr(GLdblMaxWthAmount))
    
    iRc = ShowScreenSync(Browser, "Cwd", "CwdMenu", sSubStData)
    Select Case (iRc)
    Case 0:
        If sSubStData = "@others" Then              '������
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
    
    '��ѡ�����Ҫȷ��
    If (FastWithdrawalSelect = 0) And bLoop Then       'InputCwdMenu = 0
        iRc = ShowScreenSync(Browser, "Cwd", "CwdConfirmMenu", sSubStData)
        Select Case (iRc)
        Case 0:
            If sSubStData = "@ok" Then
                FastWithdrawalSelect = 0   'ȷ��0
                iRc = ShowScreenSync(Browser, "Common", "ComPlsWait", sSubStData)
            Else
                FastWithdrawalSelect = 1   'ȡ�� 1
            End If
        Case 91:
             '����ѡ��
            FastWithdrawalSelect = 2
        Case Else:
            LogError ScreenInfo.Name + "Return error, irc = " + CStr(iRc)
            FastWithdrawalSelect = 3     '������� 3
        End Select
        Exit Function
    End If

End Function
'==========================================================================================
'�� �� �� �� �� ����������ģ��master
'�� �� �� ��   ����������ģ��ֵ
'�� �� �� ��   ����
'�� �� ֵ      ����
'�� �� �� ��   ����
'�� �� �� �� ����
'�� ��         ������
'�� �� ʱ ��   :
'==========================================================================================
Private Sub RetToMaster(ByVal S3eRetValue As Integer)
    SDOTrans.Result = S3eRetValue
End Sub
'==========================================================================================
'�� �� �� �� �� ������ȡ���������
'�� �� �� ��   ����������������
'�� �� �� ��   ����
'�� �� ֵ      �����ͳ������׵Ľ��
'�� �� �� ��   ��
'�� �� �� �� ����
'�� ��         ������
'�� �� ʱ ��   :
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
         PrjString = "Send Resersal Message Failed��please check this transaction" + vbCrLf
         PrjCHNString = "** ���ͳ�������ʧ��,���ֹ�����" + vbCrLf
    Else
        iRc = SDOLineOut.DoReceive
        If iRc <> 0 Then
            LogError "Receive Resersal Message(CWC) Failed!"
            PrjString = "Receive Resersal Message Failed��please check this transaction" + vbCrLf
            PrjCHNString = "** ���ճ�������ʧ��,���ֹ�����" + vbCrLf
        Else
            g_sRejectCode = Pcb3Dl.DlGetCharRaw("HostTransCode")
            If g_sRejectCode = "AWP" Or g_sRejectCode = "DWP" Then
                PrjString = "Reversal OK" + vbCrLf
                PrjCHNString = "�����ɹ�" + vbCrLf
                iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "R")
            Else
                PrjString = "Reversal Host Reject" + vbCrLf
                PrjCHNString = "�������ܾ���,���ֹ�����" + vbCrLf
            End If
        End If
    End If
    
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    SDOLineOut.GetData "GBLLineNum", vLineNum
    g_sDetailLineNum = Format(vLineNum, "000000")
        
    SendCwdReversal = iRc
End Function
'==========================================================================================
'�� �� �� �� �� :��¼ȡ��ɹ�ͳ��ֵ
'�� �� �� ��   ����
'�� �� �� ��   ����
'�� �� ֵ      ����
'�� �� �� ��   ����
'�� �� �� �� ����CwdOkFunction ��take notes timeout
'�� ��         ������
'�� �� ʱ ��   :
'<ʱ��>��[2005.12.12]
'<�޸���>��������
'<��ϸ��¼>��
'    ���Ӽ�¼CutOff.ini����������cutoffʱ��ӡ��ˮ
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
    
    '���Ӽ�¼CutOff.ini����������cutoffʱ��ӡ��ˮ
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
'�������� ���õ���ǰ���ó��������С��ֵ����ǰ�������ȡ���޶�
'������� ����
'������� ����
'�� �� ֵ ��0 ---- ��ʾ�Ѿ�����������ʹ�õĳ�����
'          1---- ��ʾδ�������ϸ�ĳ���
'���� ��������
'����ʱ�� :
'�޸ļ�¼��2005.8.10
'�޸��ߣ� ������
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
'�� �� �� �� �� ������������ͨѶ
'�� �� �� ��   �����ͱ�������
'�� �� �� ��   ����
'�� �� ֵ      ��ͨѶ����������������ǳ���������ʾ�����ܾ� ��ֵ 200
'   ���������� 00  ����0 ������
'   �����������160��������:
'        1 dosend <> 0 2 doreceive <> 0 3 ���������� 00 4 ����������ˮ�š���������Ͳ��� 5 MAC У���
'   �������������
'           1 dosend <>0   2 doreceive <> 0
'�� �� �� ��   ��
' �뱨�ļ�����صĺ���
' ��ͨѶ��صĺ���
'           SDOLineOut1.DoSend("0201", NORMAL_MESSAGE)
'           SDOLineOut1.DoReceive()
'�� �� �� �� ����
'     SDOCdm_GetAuthorisation ����
'�� ��         ��
'�� �� ʱ ��   :2005.8
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
    PrjCHNString = " ȡ��ķ�������" + Format(Now(), " HH:MM:SS") + vbCrLf
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
                "     ȡ�� " + Format(Now(), " HH:MM:SS") + " ��ˮ�ţ�[" + Format(vGlnLineNum, "000000") + "]" + vbCrLf + _
                "     ATM�ţ� " + Format(Pcb3Dl.DlGetCharRaw("GBLAtmCode")) + vbCrLf + _
                "     �� " + Pcb3Dl.DlGetCharRaw("GBLPrtAmount") + vbCrLf

    Call PrintJournalMedia(PrjString, PrjCHNString)

    bReverseCWC = True
    If iRc <> 0 Then
        Select Case (iRc)
            Case 96:  'not send
                bReverseCWC = False
                PrjString = "  **" + CStr(iRc) + ": Not send to ATMP" + vbCrLf
                PrjCHNString = "  " + CStr(iRc) + ":δ���͵�����" + vbCrLf
            Case 90:  ' in use
                bReverseCWC = False
                PrjString = "  **" + CStr(iRc) + ": Line In Use" + vbCrLf
                PrjCHNString = " " + CStr(iRc) + ": ��·����ʹ����" + vbCrLf
            Case Else
                PrjString = "  **" + CStr(iRc) + ": Line Status Unknown." + vbCrLf
                PrjCHNString = " " + CStr(iRc) + ":��·״̬δ֪" + vbCrLf
        End Select
        
        If bReverseCWC Then
            PrjString = PrjString + "  NO RESPONSE FROM ATMP. " + vbCrLf + "  ** Need check this transaction" + vbCrLf
            PrjCHNString = "  ** δ�յ�������Ӧ." + vbCrLf + "  **���ֹ�����" + vbCrLf
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
                PrjCHNString = "  ����������ˮ������" + vbCrLf
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
                            
                                PrjCHNString = "    ���������ӡ��� " + vbCrLf + _
                                                     "     ���������ʺţ�" + HostAccNo + _
                                                     "     ����ʱ�䣺" + Pcb3Dl.DlGetCharRaw("HostCurrentDate") + vbCrLf + _
                                                     "     ���������ţ�" + HostSeq + vbCrLf
                                Call PrintJournalMedia(PrjString, PrjCHNString)
                                                     
                                iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "Y")
                                CommunicationSubFunction = 0
                            Else        'StrWthAmount <>HostResAmount
                                PrjString = "  **  Received Amount <> Send Amount"
                                PrjCHNString = "  ���ؽ�������Ͳ�һ��" + vbCrLf
                                Call PrintJournalMedia(PrjString, PrjCHNString)
                                Pcb3Dl.DlSetCharRaw "GBLCWDResult", "CU"
                                Call SendAEXMessage("2013")
                                iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "U")
                                g_bIsWthReversal = True
                                CommunicationSubFunction = UserReplyCommErr
                            End If
                        Else        'HostResAmount not numeric
                            PrjString = " **   Received Amount Erlaotror "
                            PrjCHNString = "  �������ؽ������" + vbCrLf
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
                        PrjCHNString = "  �յ������ܾ���" & g_sRejectCode + vbCrLf
                        Pcb3Dl.DlSetCharRaw "GBLCWDResult", "CU"
                        Call PrintJournalMedia(PrjString, PrjCHNString)
                        iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "U")
                        g_bIsWthReversal = True
                        CommunicationSubFunction = UserReplyCommErr
                End Select
            Else            'HostResLineNum <>g_sDetailLineNum
                PrjString = "  **  Received LineNo <> Send LineNo"
                PrjCHNString = "  ������ˮ�������Ͳ�һ��" + vbCrLf
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
            PrjCHNString = "  MACУ��ʧ��" + vbCrLf
            Call PrintJournalMedia(PrjString, PrjCHNString)
            iRc = Pcb3Dl.DlSetCharRaw("ResetTransKey", "R")
            CommunicationSubFunction = UserReplyCommErr
        Case Else:
'            HostSeq = "00000000000000000000000"
'            iRc = SendCwdReversal("4001")
'            Call CwdReversalTotal
            PrjString = "  Receiving host response error" + vbCrLf + "  ** Need check this transaction " + vbCrLf
            PrjCHNString = "  ͨѶ����." + vbCrLf + "  **���ֹ�����" + vbCrLf
            Call PrintJournalMedia(PrjString, PrjCHNString)
            iRc = Pcb3Dl.DlSetCharRaw("GBLKeepAccountFlag", "U")
            g_bIsWthReversal = True
            Pcb3Dl.DlSetCharRaw "GBLCWDResult", "CE"
            CommunicationSubFunction = UserReplyCommErr
    End Select
End Function
'==========================================================================================
'�� �� �� �� �� :��¼���ݿ�
'�� �� �� ��   ����
'�� �� �� ��   ����
'�� �� ֵ      ����
'�� �� �� ��   ����
'�� �� �� �� ����
'�� ��         ��������
'�� �� ʱ ��   : 2005.6.23
'--------���ݿ�����------------
'���ʱ�־�� Y �Ѽ���  N δ���� U ����
'����ԭ�򣺼�����
'�����ܾ��룺���������������ܾ�����Ҫ�ճ�
'�������ͣ���� ȡ��
'�����ͱ�־
'����ʱ��
'�����ʺ�
'���׽��
'������ˮ�ţ� 5 λ
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
'�������� :��ӡ��ˮ
'������� ����ˮ��ӡbuffer
'�����������
'����ֵ����
'���ú�������
'�������������ӡ��ˮʱ����
'���ߣ�������
'����ʱ�� : 2005.6.22
'================================================================================
Sub PrintJournalMedia(ByRef JournalBuf As String, ByRef CHNJournalBuf As String)
    If (Len(JournalBuf) <> 0) And (Len(CHNJournalBuf) <> 0) Then
        PrintJournal SDOPrj, JournalBuf, CHNJournalBuf, g_sPrjLanguage
    End If
End Sub
'===================================================================================
'�������� :��¼����ͳ��ֵ
'������� ����
'�����������
'����ֵ����
'���ú�������
'�����������
'���ߣ�
'����ʱ�� : 2004
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
'�������� :ȡ��ɹ���¼ͳ��ֵ����ʾ���桢׼����������
'������� ����
'�����������
'����ֵ����
'���ú�����
'�����������AtWithdrEnd = 0 ��95:ȡ�ʱ�������ճ�Ʊ����Ϊ��ʱ ��99:IsNotesPresented = True
'���ߣ�������
'����ʱ�� : 2005.8.23
'====================================================================================
Private Sub CwdOkFunction()
    Dim sSubStData     As String
    
    Call CwdOkTotal
    Call PrintCassLeftNum
    
    iRc = ShowScreenSync(Browser, "Cwd", "CwdTransOk", sSubStData)

    '׼��������ӡ���� ȡ��ɹ�
    Call DrawWthPrr(Pcb3Dl, WthPrrOK)
End Sub
'===================================================================================
'�������� :ȡ����ܳ��ַ������ʱ��¼���ݿ⡢������־����ʾ���棬׼����������
'������� ����
'�����������
'����ֵ����
'���ú�����
'�����������95 optype_cdm_shutterproblem;99 IsNotesPresented = True and optype_cdm_shutterproblem
'���ߣ�������
'����ʱ�� : 2005.8.23
'--------------------------
'<ʱ��>��[2005.11.1]
'<�޸���>��������
'<��ϸ��¼>��
'���Ӽ�¼cwdoktotal �ʹ�ӡ�������PrintCassLeftNum
'====================================================================================
Private Sub CwdCrimeFunction()
    Dim sSubStData     As String
    Dim PrjString      As String
    Dim PrjCHNString   As String
    
    iRc = Pcb3Dl.DlSetCharRaw("CWDCrimePossible", "Y")
    Pcb3Dl.DlSetCharRaw "GBLDoRecovery", "C"
    PrjString = " TRANSACTION OK (Crime Possible)" + vbCrLf
    PrjCHNString = "**ȡ��ɹ���������������������" + vbCrLf
    Call PrintJournalMedia(PrjString, PrjCHNString)

    '��¼���ݿ⡢������־
    iRc = Pcb3Dl.DlSetCharRaw("GBLCWDResult", "CR")
    Call RecordDB_CWDLog
    Call WriteTranslog("���ɳ���")
    Call CwdOkTotal
    Call PrintCassLeftNum
    
    iRc = ShowScreenSync(Browser, "Cwd", "CwdCrime", sSubStData)

    '׼��������ӡ���� ȡ��ɹ�
    Call DrawWthPrr(Pcb3Dl, WthPrrOK)
 End Sub
'===================================================================================
'�������� :��¼ȡ�����־�����ں�����ʾ��
'������� ����
'�����������
'����ֵ����
'���ú�����
'���������:
'���ߣ�����
'����ʱ�� : 2004
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
    
    sTransLogRec = "ȡ��|" + Format(Now(), "MM/DD|HH:MM|") + _
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
'�汾�ţ�Agilis 1.6
'�μ�sdohelp�ļ�DoEjectCard�������������4���¼����д���
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
        PrjCHNString = "   **��ʱ���ͻ�δȡ��" + vbCrLf
        Call PrintJournalMedia(PrjString, PrjCHNString)
        
        '��¼�̿��ļ�
        Call RecordCpdCardLog("1035")
        
        '��ӡ�̿�����
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
'�������� :�̿�������ʾ���桢��ӡ��ˮ��ִ���̿����׼���̿�����
'������� ����
'�����������
'����ֵ����
'���ú�����DoTakeCard
'���������: DoEjectCard <>0; SDOIdc_AtEjectEnd
'���ߣ�������
'����ʱ�� : 2005.8.23
'====================================================================================
Private Sub CaptureCard()
    Dim sSubStData     As String
    
    iRc = ShowScreenSync(Browser, "EndVisit", "EjectCardError", sSubStData)
    iRc = SDOFep.SetIndicator(ind_audio, audio_off)
    Call SendExceptionMessage(SDOLineOut, Pcb3Dl, "24")
    
    PrintJournalMedia "   **Eject Card Err in CWD", "   **ȡ��ʱ�˿�ʧ��"
               
    iRc = SDOIdc.DoTakeCard 'capture the card
    If iRc <> 0 Then
        PrintJournalMedia "   **Capture Card Err in CWD.", "   **ȡ��ʱ�̿�ʧ��"
    End If
    Call RecordCpdCardLog("1035")
    
    Call DrawCpdCardPrr(Pcb3Dl)
    
End Sub
'===================================================================================
'�������� :��¼�̿��ļ�
'������� �������̿����������
'�����������
'����ֵ����
'���ú�����
'����������� SDOIdc_AtEjectEnd   ��ʱδ�ÿ�;CaptureCard
'���ߣ�
'����ʱ�� : 2004
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
'�������� ��ȡ��ɹ��󣬴�ӡ��ȡ���ʣ������
'�����������
'�����������
'����ֵ ����
'���ú�����
'�� �� �� �� ����
'���ߣ����
'����ʱ��: 2005-08-29
'==========================================================================================
'<ʱ��>��[2005.11.25]
'<�޸���>:������
'<��ϸ��¼>��
'ע�����cashunit��Ŀ��ĳ������»����Ŷ������������SDOCdm.NbrOfBoxesUsedҲ����֮������
'֮ǰ�ĳ�������CasLefNumԽ��������
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
            PrjCHNString = PrjCHNString + "����" + CStr(i) + ": " + CasLefNum(i) + " "
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
'�������� :����AEX����ͨѶ����
'������� ���������
'�����������
'����ֵ����
'���ú�������
'�����������
'���ߣ�������
'����ʱ�� : 2005.11.10
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
'�������� :�����������Ƿ����ļ��е���ͬ
'������� ����
'�����������
'����ֵ��
'���ú�������
'�����������
'���ߣ�������
'����ʱ�� : 2005.11.10
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
            PrjCHNString = " �³������ϣ����鴫��ͨ���Ƿ��п�������" + vbCrLf
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            LogError PrjString
            iRc = Pcb3Dl.DlSetCharRaw("GBLCdmRecoveryNeeded", "N")
        End If

        If sDevice = "CDM" And (Not CheckSPInfo(sDescr, "FloatingCode")) Then     '���������ļ��е���ͬ
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


