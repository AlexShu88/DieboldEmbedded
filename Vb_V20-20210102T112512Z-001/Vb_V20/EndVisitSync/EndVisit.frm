VERSION 5.00
Object = "{B2110643-3E81-11D3-8ACC-00C04FF20A5D}#1.2#0"; "TransProv.dll"
Object = "{9C37E835-6A58-11D1-80C0-0020AF7093F9}#1.2#0"; "Dl.ocx"
Object = "{6C4DD4AB-27D5-11D3-96C4-000000000000}#1.0#0"; "S3ELineOutTcp.ocx"
Object = "{3751B5D1-D348-11D0-AD02-0060970C3D2F}#1.0#0"; "SDOPrr.ocx"
Object = "{BD8177C0-832C-11CF-BF42-0020AF7093F9}#1.0#0"; "SDOIdc.ocx"
Object = "{192DFCF0-F664-11D3-8BD4-00C04FF20A5D}#1.1#0"; "AdvBrowser.ocx"
Object = "{EACE4ED6-6930-11D0-AC6C-0020AFBDD1D4}#1.0#0"; "SDOFep.ocx"
Object = "{E64F71A6-E705-4151-9895-5138B7D67F3A}#1.0#0"; "CHPrj.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form EndVisit 
   Caption         =   "EndVisit"
   ClientHeight    =   2430
   ClientLeft      =   2220
   ClientTop       =   345
   ClientWidth     =   4305
   Icon            =   "EndVisit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4305
   WindowState     =   1  'Minimized
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3720
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin CHPRJLib.CHPrj SDOPrj 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   1
   End
   Begin S3EADVBROWSERLibCtl.AdvBrowser Browser 
      Height          =   615
      Left            =   1440
      OleObjectBlob   =   "EndVisit.frx":1272
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin TRANSPROVLibCtl.TransactionProvider S3ETransactionProvider 
      Height          =   690
      Left            =   2760
      OleObjectBlob   =   "EndVisit.frx":1298
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin SDOPrrLibCtl.SDOPrr SDOPrr 
      Height          =   675
      Left            =   1440
      OleObjectBlob   =   "EndVisit.frx":12D4
      TabIndex        =   3
      Top             =   870
      Width           =   1215
   End
   Begin SDOIdcLibCtl.SDOIdc SDOIdc 
      Height          =   675
      Left            =   105
      OleObjectBlob   =   "EndVisit.frx":1304
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin S3ELINEOUTLib.S3ELineOut S3ELineOut 
      Height          =   705
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   1244
      _StockProps     =   1
   End
   Begin DLLib.DL Pcb3dl 
      Left            =   105
      Top             =   1650
      _Version        =   65538
      _ExtentX        =   2196
      _ExtentY        =   1217
      _StockProps     =   0
   End
   Begin SDOFepLibCtl.SDOFep SDOFep 
      Height          =   675
      Left            =   1440
      OleObjectBlob   =   "EndVisit.frx":1336
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "EndVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================================================
'��Ȩ˵��:  �ϱ���˾�й���������
'�汾�ţ�ԭ���а�Optevaȡ�������Timer��ʽ���汾��1.4.15
'�������ڣ�2005.8
'���ߣ�  ����
'ģ�鹦�ܣ� ������������
'��Ҫ�������书��
'   EjectCard           : �˿�
'   EjectCardError      : �˿�����
'   TakeTicketChoice    : ������ӡѡ��
'   TakeTicketError     : ������ӡ����
'   MultiTransaction    : �ཻ��ѡ��
'   RetToMaster         : ��������ģ��
'   SendExceptionMessage: ����������Ϣ����
'   RecordCpdCardLog    : ��¼�̿��ļ�
'ȫ�ֱ���
'   g_sFormToPrint      : ���ڴ�ӡ��Form����
'   g_bIsPrrCapture     : �Ƿ��ӡ�̿�����
'   bForceToEjectCard   : �Ƿ�ǿ���̿�
'   g_bIsAudioOn        : �Ƿ�ر��������˿�������
'�����޸ģ�
'         1 �ڽ������������в��ٵ���timer
'==========================================================================================
'<ʱ��>��[2005.08.30]
'<�޸���>������
'<��ǰ�汾>��1.4.16
'================================================================
'<ʱ��>��2005.11.22
'<�޸���>��������
'<��ϸ��¼>��
' ����Ѷ��˾����Ƕ��ʽ��������Ӳ��¼��� EVR(Embedded Net DVR)
' ���Ӻ��� SendMessageToCommPortEjectCard,��������µ��ã�
'  �˿��ɹ����̿��ɹ������͵�EVR��IDLE�忨ʱ��ʾ�Ŀ��ŴӼ����Ļ��ɾ��
'================================================================
'<ʱ��>��2005.11.29
'<�޸���>��������
'<��ϸ��¼>�����Ӵ�ӡ����Ҫ���̿�ʱ����ˮ
'<ʱ��>��2005.12.22
'<�޸���>��������
'�汾�ţ� 1.1.16
'   ���ճ����ڸ������е��޸�
'   1 �ܾ���ӡ����  2��ReturnMasterʱ����ResetATMPrr�� 3  ȡ����ʱ�̿�ʱ����ӡƾ����
'================================================================

Option Explicit

Const CardRetainFile                    As String = "C:\s3e\logs\logapp\CardRetain.txt"

Const rcDEVIDC_DOEJECTCC_CCCAPTURED     As Integer = 98

Const ReturnOk                          As Integer = 10  'EndVisit ok - go to Idle
Const ReturnCapturedCard                As Integer = 104 'Capture a timeout card ok - go to Idle
Const ReturnIdcError                    As Integer = 105 'IDC error during transaction - go to Idle to let CheckState() check
Const ReturnPrrError                    As Integer = 40
Const ReturnHostReject                  As Integer = 50
Const ReturnMultiTransaction            As Integer = 110
Const ReturnToOperator                  As Integer = 302

Const BeepEffect_WFS_SIU_ON             As Integer = 2
Const BeepEffect_WFS_SIU_SLOW_FLASH     As Integer = 4
Const BeepEffect_WFS_SIU_MEDIUM_FLASH   As Integer = 8
Const BeepEffect_WFS_SIU_QUICK_FLASH    As Integer = 16
Const GlobalINIPath                     As String = "C:\atmwosa\ini\"

Const OperatorAction                    As Integer = 2

Private nrc                             As Integer
Private g_sFormToPrint                  As String
Private g_bIsPrrCapture                 As Boolean
Private bForceToEjectCard               As Boolean
Private GLnAction                       As Long
Private g_AtmPrrType                    As String
'Added for Icbc3030
Private g_bIsAudioOn                    As Boolean
Dim g_sPrjLanguage                      As String
Dim TakeCard                            As String
'==========================================================================================
'�����Ĺ��� ��VB����װ��,��ʼ����������
'�������   ����
'�������   ����
'����ֵ     ����
'����
'����ʱ��   :
'==========================================================================================
Private Sub Form_Load()
    Dim sValue                      As String
    Dim ReceiptPaperType            As String
    
    sValue = "The version number of " & App.EXEName & ".exe is " & App.Major & "." _
            & App.Minor & ".0." & App.Revision
    
    LogInfo (sValue)

    ReceiptPaperType = GetIniS(GlobalINIPath + "global.ini", "Bank_Environment", _
                    "Receipt_paper_type", "")
    
    TakeCard = GetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "EjectCardMode", "A")

    
    If ReceiptPaperType = "B" Then
        g_sFormToPrint = "ATMPrr1"
    Else
        g_sFormToPrint = "ATMPrr"
    End If
    
    If GetIniS(GlobalINIPath + "global.ini", "Bank_Environment", "PrjLanguage", "E") = "E" Then
        g_sPrjLanguage = "E"
    Else
        g_sPrjLanguage = "C"
    End If

    S3ETransactionProvider.Available = True
End Sub
'==========================================================================================
'�����Ĺ��� ���������׳����˳�
'�������   ����
'�������   ����
'����ֵ     ����
'����
'����ʱ��   :
'==========================================================================================
Private Sub S3ETransactionProvider_QuitTransaction()
    Unload EndVisit
End Sub
'==========================================================================================
'�����Ĺ��� ���������׳������
'�������   �� ��ruler���ô�ģ��ʱActionֵ
'�������   ����
'����ֵ     ����
'����
'����ʱ��   :
'==========================================================================================
Private Sub S3ETransactionProvider_StartTransaction(ByVal Action As Long)
    Dim HostRejectCard              As String
    Dim sPrrOthersMark              As String
    Dim bForcePrintReciept          As Boolean
    Dim sSubStData                  As String
    Dim PrjString                   As String
    Dim PrjCHNString                As String
    
    LogInfo "Start Transaction action=" + CStr(Action)
    
    GLnAction = Action
    
    bForcePrintReciept = False
    bForceToEjectCard = False
    g_bIsPrrCapture = False
    g_bIsAudioOn = False
    
    sPrrOthersMark = Pcb3dl.DlGetCharRaw("PrrOthersMark")
    HostRejectCard = Pcb3dl.DlGetCharRaw("HostRejectCard")
    nrc = Pcb3dl.DlSetCharRaw("GBLAudioOffAgain", "N")
        
    Select Case Action
    Case 30:      '�����ܾ�
        LogInfo "HostRejectCard = " + HostRejectCard
        Select Case HostRejectCard
        Case "E":
            'by Chenlei for Boc_Fujian, have a choice to print receipt while a transaciton rejected by host.
            'nrc = EjectCard()
            bForceToEjectCard = True
            nrc = TakeTicketChoice()
            Exit Sub
        Case "Y":
            Pcb3dl.DlSetCharRaw "HostRejectCard", "N"
            
            Call SendExceptionMessage(S3ELineOut, Pcb3dl, "95")
            nrc = SDOIdc.DoTakeCard 'capture the card
            If nrc <> 0 Then
                Call SendExceptionMessage(S3ELineOut, Pcb3dl, "24")   'BGR OUT OF SERVICE DURING IDLE
                PrjString = FormTransExp(Pcb3dl.DlGetCharRaw("FitAccNo"), "   **Capture Card Err in EndVisit.")
                PrjCHNString = (FormTransExpCHN(Pcb3dl.DlGetCharRaw("FitAccNo"), "   **�̿�����."))
                PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            End If
            
            PrjString = FormTransExp(Pcb3dl.DlGetCharRaw("FitAccNo"), "   Host Capture Card.")
            PrjCHNString = (FormTransExpCHN(Pcb3dl.DlGetCharRaw("FitAccNo"), "   ����Ҫ���̿�."))
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            
            Call RecordCpdCardLog("1095")
'            Call ResetATMPrr(EndVisit.Pcb3dl)
            Call DrawCpdCardPrr(Pcb3dl)
            bForcePrintReciept = True
            g_bIsPrrCapture = True
        Case Else     'For N or other situation
            nrc = MultiTransaction()
            Exit Sub
        End Select

    Case 209 'Withdrawal OK
        If TakeCard = "B" Then     'ȡ�����˿������
            bForceToEjectCard = True
        End If
    
    Case 309 'Withdraw possible crime
        bForcePrintReciept = True
        bForceToEjectCard = True
    Case 211 'Pinchange
        bForceToEjectCard = True
    Case 202
        bForcePrintReciept = True
                                                
    Case Else
        nrc = EjectCard()
        Exit Sub
    End Select
    
    If (bForcePrintReciept) Then
        If (SDOPrr.Available) Then
            nrc = SDOPrr.DoPrintForm(g_sFormToPrint)
            If nrc <> 0 Then
                nrc = TakeTicketError(g_bIsPrrCapture)
            Else
                '��ʱ���̽���SDOPRR���ơ�
                If (g_bIsPrrCapture) Then
                    nrc = ShowScreenSync(Browser, "EndVisit", "HostCapturedCard", sSubStData)
                Else
                    nrc = ShowScreenSync(Browser, "Common", "ComTakeTicket", sSubStData)
                End If
            End If
        Else
            If (g_bIsPrrCapture) Then
                nrc = TakeTicketError(True)
            ElseIf (bForceToEjectCard) Then
                '��ʱ���̽���SDOIDC���ơ�
                nrc = EjectCard()
            Else
                nrc = MultiTransaction()
            End If
        End If
    Else
        If (Not SDOPrr.Available) Then
            If (Not bForceToEjectCard) Then
                nrc = MultiTransaction()
            Else
                '��ʱ���̽���SDOIDC���ơ�
                nrc = EjectCard()
            End If
        Else
            nrc = TakeTicketChoice()
        End If
    End If
End Sub
'==========================================================================================
'�汾�ţ�Agilis 1.6
'�μ�sdohelp�ļ�DoEjectCard��������������ĸ��¼����д���
'==========================================================================================
Private Sub SDOIdc_AtEjectStart()
    LogInfo "SDOIdc_AtEjectStart: UserReply = 0"
    SDOIdc.UserReply = 0
End Sub
Private Sub SDOIdc_EjectCardTimeOut()
    Dim sSubStData                  As String
    
    nrc = ShowScreenSync(Browser, "EndVisit", "TakeCardwarning", sSubStData)
    LogInfo "SDOIdc_EjectCardTimeOut: UserReply = 0"
    SDOIdc.UserReply = 0
End Sub
Private Sub SDOIdc_CardWillBeCaptured()
    LogInfo "SDOIdc_CardWillBeCaptured: UserReply = 0"
    SDOIdc.UserReply = 0
End Sub
Private Sub SDOIdc_AtEjectEnd(ByVal rcEjectCard As Integer)
On Error Resume Next
    Dim PrjString                   As String
    Dim PrjCHNString                As String
    Dim sSubStData                  As String
    Dim nRet                        As Integer
    
    LogInfo "SDOIdc_AtEjectEnd = " + CStr(rcEjectCard)
    
    nrc = SDOFep.SetIndicator(ind_audio, audio_off)
    If nrc <> 0 Then
        PrjString = "SDOIdc_AtEjectEnd: Audio is Not OFF! RC=" + CStr(nrc) + _
                    ". XFSCode=" + CStr(SDOFep.LastReturn)
        LogError (PrjString)
        g_bIsAudioOn = True
    End If
    Select Case (rcEjectCard)
    Case 0:
        If (GLnAction = OperatorAction) Then
            nrc = Pcb3dl.DlSetCharRaw("GBLOperStatus", "1")
            nRet = ReturnToOperator
        Else
            nRet = ReturnOk
        End If
        '�ж��Ƿ�ʹ��Ƕ��ʽ��������Ӳ��¼��� EVR(Embedded Net DVR)
        If Pcb3dl.DlGetCharRaw("GBLEVRUse") = "Y" Then
            Call SendMessageToCommPortEjectCard
        End If
        nrc = RetToMaster(nRet)
    Case rcDEVIDC_DOEJECTCC_CCCAPTURED:
        g_bIsPrrCapture = True
        
        Pcb3dl.DlSetCharRaw "GBLATMLocRejCode", "%CT"
        PrjString = FormTransExp(Pcb3dl.DlGetCharRaw("FitAccNo"), "   **TimeOut:card not taken by client.")
        PrjCHNString = (FormTransExpCHN(Pcb3dl.DlGetCharRaw("FitAccNo"), "   **�ͻ���ʱδȡ��."))
        PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
        
        Call RecordCpdCardLog("1035")
          '�ж��Ƿ�ʹ��Ƕ��ʽ��������Ӳ��¼��� EVR(Embedded Net DVR)
        If Pcb3dl.DlGetCharRaw("GBLEVRUse") = "Y" Then
            Call SendMessageToCommPortEjectCard
        End If
        
        'by Chenlei for Boc_Fujian, ��ʱδȡ�����̿�������ӡƾ����
        nrc = ShowScreenSync(Browser, "EndVisit", "CardCaptured", sSubStData)
        nrc = RetToMaster(ReturnOk)
        
'        If (SDOPrr.Available) Then
''            Call ResetATMPrr(EndVisit.Pcb3dl)
'            Call DrawCpdCardPrr(Pcb3dl)
'
'            nrc = SDOPrr.DoPrintForm(g_sFormToPrint)
'            If nrc <> 0 Then
'                nrc = TakeTicketError(True)
'            Else
'                '��ʱ���̽���SDOPRR���ơ�
'                LogInfo "DoPrintForm 'CpdCardPrr' in SDOIdc_AtEjectEnd return 0"
'                nrc = ShowScreenSync(Browser, "EndVisit", "CardCaptured", sSubStData)
'                Exit Sub
'            End If
'        Else
'            nrc = TakeTicketError(True)
'        End If
       
    Case Else:
        nrc = EjectCardError()
    End Select
End Sub
'==========================================================================================
'�汾�ţ�Agilis 1.6
'�μ�sdohelp�ļ�DoPrintForm��������������ĸ��¼����д���
'==========================================================================================
Private Sub SDOPrr_AtPrintFormStart()
    
    nrc = SDOFep.SetIndicator(ind_audio, audio_exclamation + audio_continuous)
    If nrc <> 0 Then
        LogError ("AtPrintFormStart: Audio is Not ON! RC=" + CStr(nrc) + ". XFSCode=" + CStr(SDOFep.LastReturn))
    End If
    
    LogInfo "SDOPrr_AtPrintFormStart: UserReply = 0"
    SDOPrr.UserReply = 0
End Sub
Private Sub SDOPrr_AtPresented()
    LogInfo "SDOPrr_AtPresented: UserReply = 0"
    SDOPrr.UserReply = 0
End Sub
Private Sub SDOPrr_AtPresentTimeout()
    LogInfo "SDOPrr_AtPresentTimeout: UserReply = 0"
    SDOPrr.UserReply = 0
End Sub
Private Sub SDOPrr_AtPrintFormEnd(ByVal Rc As Integer)
    LogInfo "SDOPrr_AtPrintFormEnd = " + CStr(Rc)
    
    nrc = SDOFep.SetIndicator(ind_audio, audio_off)
    If nrc <> 0 Then
        LogError ("AtPrintFormEnd: Audio is Not OFF! RC=" + CStr(nrc) + ". XFSCode=" + CStr(SDOFep.LastReturn))
        g_bIsAudioOn = True
    End If
    If (Not g_bIsPrrCapture) Then
        Select Case (Rc)
        Case 0, 91, 98
            If (Not bForceToEjectCard) Then
                nrc = MultiTransaction()
            Else
                '��ʱ���̽���SDOIDC���ơ�
                nrc = EjectCard()
            End If
        Case Else
            nrc = TakeTicketError(False)
        End Select
    Else
        g_bIsPrrCapture = False
        Select Case (Rc)
        Case 0, 91, 98
            nrc = RetToMaster(ReturnOk)
        Case Else
            nrc = TakeTicketError(True)
        End Select
    End If
End Sub
'==========================================================================================
'�����Ĺ��� ���˿� EjectCard
'�������   ����
'�������   ����
'����ֵ     ����
'���ú���   ��RetToMaster, EjectCardError
'��������� ��
'����       ��
'����ʱ��   :
'==========================================================================================
Private Function EjectCard() As Integer
On Error GoTo ErrHandler
    Dim sSubStData                  As String
    
    If (SDOIdc.CardPosition = devidc_cardpresent) Then
        nrc = ShowScreenSync(Browser, "EndVisit", "TakeCard", sSubStData)
        nrc = SDOIdc.DoEjectCard
        If nrc <> 0 Then
            nrc = EjectCardError()
        Else
            '��ʱ���̽���SDOIDC���ơ�
            nrc = SDOFep.SetIndicator(ind_audio, audio_keypress + audio_continuous)
            If nrc <> 0 Then
                sSubStData = "pageTakeCard: Audio is Not ON! RC=" + CStr(nrc) + _
                            ". XFSCode=" + CStr(SDOFep.LastReturn)
                LogError (sSubStData)
            End If
        End If
    Else
        nrc = RetToMaster(ReturnOk)
    End If
    Exit Function
ErrHandler:
    nrc = ErrorHandlerFunction("EjectCard:", 99)
    nrc = SDOIdc.DoEjectCard
    If nrc <> 0 Then
        nrc = EjectCardError()
    End If
End Function
'==========================================================================================
'�����Ĺ��� ���˿����� EjectCardError
'�������   ����
'�������   ����
'����ֵ     ����
'���ú���   ��RetToMaster, SendExceptionMessage
'��������� ��
'����       ��
'����ʱ��   :
'==========================================================================================
Private Function EjectCardError() As Integer
On Error GoTo ErrHandler
    Dim sSubStData                  As String
    
    nrc = ShowScreenSync(Browser, "EndVisit", "EjectCardError", sSubStData)
    LogError "DoEjectCard method Error. RC=" & nrc
    Call SendExceptionMessage(S3ELineOut, Pcb3dl, "24")
    
    nrc = SDOPrj.DoPrint(FormTransExp(Pcb3dl.DlGetCharRaw("FitAccNo"), "   Eject Card Err in EndVisit."))
    SaveCNJournal (FormTransExpCHN(Pcb3dl.DlGetCharRaw("FitAccNo"), "   **�˿�������δ�˳�."))
    
    Call RecordCpdCardLog("1024")
    nrc = RetToMaster(ReturnIdcError)
    Exit Function
ErrHandler:
    nrc = ErrorHandlerFunction("EjectCardError:", 99)
    nrc = RetToMaster(ReturnIdcError)
End Function
'==========================================================================================
'�����Ĺ��� ��������ӡѡ�� TakeTicketChoice
'�������   ����
'�������   ����
'����ֵ     ����
'���ú���   ��TakeTicketError, EjectCard, MultiTransaction
'��������� ��
'����       ��
'����ʱ��   :
'==========================================================================================
Private Function TakeTicketChoice() As Integer
On Error GoTo ErrHandler
    Dim sSubStData                  As String
    
    nrc = ShowScreenSync(Browser, "Common", "ComTakeTicketChoice", sSubStData)
    Select Case (nrc)
    Case 0:
        Select Case sSubStData
        Case "@ok"
            nrc = SDOPrr.DoPrintForm(g_sFormToPrint)
            If nrc <> 0 Then
                nrc = TakeTicketError(False)
            Else
                nrc = ShowScreenSync(Browser, "Common", "ComTakeTicket", sSubStData)
            End If
        Case "@stop"
            If (Not bForceToEjectCard) Then
                nrc = MultiTransaction()
            Else
                nrc = EjectCard()
            End If
        Case Else
            LogError ScreenInfo.Name + " select a impossible function:" + sSubStData
            nrc = EjectCard()
        End Select
    Case 91:
        nrc = EjectCard()
    Case Else
        LogError ScreenInfo.Name + "Return error, nRc = " + Str(nrc)
        nrc = EjectCard()
    End Select
    Exit Function
ErrHandler:
    nrc = ErrorHandlerFunction("TakeTicketChoice:", 99)
    nrc = EjectCard()
End Function
'==========================================================================================
'�����Ĺ��� ��������ӡ���� TakeTicketError
'�������   ���Ƿ�Ϊ�̿���ӡ
'�������   ����
'����ֵ     ����
'���ú���   ��RetToMaster, EjectCard, MultiTransaction
'��������� ��
'����       ��
'����ʱ��   :
'==========================================================================================
Private Function TakeTicketError(ByVal bCardCapture As Boolean) As Integer
On Error GoTo ErrHandler
    Dim sSubStData                  As String
    Dim PrjString                   As String
    Dim PrjCHNString                As String
    
    Call SendExceptionMessage(S3ELineOut, Pcb3dl, "28")
    PrjString = FormTransExp(Pcb3dl.DlGetCharRaw("FitAccNo"), "   **PRR out of service.")
    PrjCHNString = (FormTransExpCHN(Pcb3dl.DlGetCharRaw("FitAccNo"), "   **ƾ����ӡ������."))
    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
    If (bCardCapture) Then
        nrc = ShowScreenSync(Browser, "EndVisit", "CardCapturePrrErr", sSubStData)
        nrc = RetToMaster(ReturnPrrError)
    Else
        nrc = ShowScreenSync(Browser, "Common", "ComTakeTicketError", sSubStData)
        If (Not bForceToEjectCard) Then
            nrc = MultiTransaction()
        Else
            nrc = EjectCard()
        End If
    End If
    Exit Function
ErrHandler:
    nrc = ErrorHandlerFunction("TakeTicketError:", 99)
    nrc = EjectCard()
End Function
'==========================================================================================
'�����Ĺ��� ���ཻ��ѡ�� MultiTransaction
'�������   ����
'�������   ����
'����ֵ     ����
'���ú���   ��RetToMaster, EjectCard
'��������� ��
'����       ��
'����ʱ��   :
'==========================================================================================
Private Function MultiTransaction() As Integer
On Error GoTo ErrHandler
    Dim sSubStData                  As String
    
    nrc = ShowScreenSync(Browser, "EndVisit", "MultiTransaction", sSubStData)
    Select Case (nrc)
    Case 0:
        Select Case sSubStData
        Case "@Continue"
            nrc = RetToMaster(ReturnMultiTransaction)
        Case "@stop"
            nrc = EjectCard()
        Case Else
            LogError ScreenInfo.Name + " select a impossible function:" + sSubStData
            nrc = EjectCard()
        End Select
    Case 91:
        nrc = EjectCard()
    Case Else
        LogError ScreenInfo.Name + "Return error, nRc = " + Str(nrc)
        nrc = EjectCard()
    End Select
    Exit Function
ErrHandler:
    nrc = ErrorHandlerFunction("MultiTransaction:", 99)
    nrc = EjectCard()
End Function
'==========================================================================================
'�����Ĺ��� ����������ģ�� RetToMaster
'�������   ����������ģ��ֵ
'�������   ����
'����ֵ     ����
'���ú���   ����
'��������� ��
'����       ��
'����ʱ��   :
'==========================================================================================
Private Function RetToMaster(ByVal ReturnValue As Integer) As Integer
On Error GoTo ErrHandler
    Dim PrjString                   As String
    
    Call ResetATMPrr(EndVisit.Pcb3dl)
    
    Pcb3dl.DlSetCharRaw "HostRejectCard", ""
    If (g_bIsAudioOn) Then
        nrc = SDOFep.SetIndicator(ind_audio, audio_off)
        If nrc <> 0 Then
            PrjString = "RetToMaster: Audio is Not OFF! RC=" + CStr(nrc) + _
                        ". XFSCode=" + CStr(SDOFep.LastReturn)
            LogError (PrjString)
        Else
            g_bIsAudioOn = False
        End If
    End If
    If (g_bIsAudioOn) Then
        nrc = Pcb3dl.DlSetCharRaw("GBLAudioOffAgain", "Y")
    End If

    If (ReturnValue = ReturnMultiTransaction) Then
        S3ETransactionProvider.Result = ReturnValue
    Else
        'modify the value to let Monitor to check whether it should do recovery
        Pcb3dl.DlSetCharRaw "GBLDoRecovery", "1"
        Sleep (500)
        S3ETransactionProvider.Result = ReturnValue
    End If
    Exit Function
ErrHandler:
    nrc = ErrorHandlerFunction("RetToMaster:", 99)
    S3ETransactionProvider.Result = ReturnIdcError
End Function
'===================================================================================
'��������  :��¼�̿��ļ� RecordCpdCardLog
'�������  �������̿����������
'�������  ����
'����ֵ    ����
'���ú���  ��
'����������� SDOIdc_AtEjectEnd   ��ʱδ�ÿ�;EjectCardError
'���ߣ�
'����ʱ��  :
'====================================================================================
Private Sub RecordCpdCardLog(ExceptCode As String)
On Error GoTo ErrHandler
    Dim sTime                       As String
    Dim FullCardAccNo               As String

    Pcb3dl.DlSetLong "TotCapCardNum", Pcb3dl.DlGetInt("TotCapCardNum") + 1
    
    sTime = Format(Now(), "YYYYMMDDHHMM")
    FullCardAccNo = Format(Pcb3dl.DlGetCharRaw("FitAccNo"), "@@@@@@@@@@@@@@@@@@@@!")
        
    Open CardRetainFile For Append As #1
    Print #1, sTime + " " + FullCardAccNo + " " + ExceptCode
    Close #1
    Exit Sub
ErrHandler:
    nrc = ErrorHandlerFunction("RecordCpdCardLog:", 99)
End Sub
Private Sub SendMessageToCommPortEjectCard()
    Dim StrTemp              As String
    Dim LngTemp              As Integer
    Dim i                    As Integer
    
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    StrTemp = Chr(2) + "ATM2"
    
    For i = 1 To 91
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
