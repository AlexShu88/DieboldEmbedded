VERSION 5.00
Object = "{B2110643-3E81-11D3-8ACC-00C04FF20A5D}#1.2#0"; "TransProv.dll"
Object = "{6C4DD4AB-27D5-11D3-96C4-000000000000}#1.0#0"; "S3ELineOutTcp.ocx"
Object = "{9C37E835-6A58-11D1-80C0-0020AF7093F9}#1.2#0"; "Dl.ocx"
Object = "{3751B5D1-D348-11D0-AD02-0060970C3D2F}#1.0#0"; "SDOPrr.ocx"
Object = "{192DFCF0-F664-11D3-8BD4-00C04FF20A5D}#1.1#0"; "AdvBrowser.ocx"
Object = "{5C094E41-67D2-11D0-AC6B-0020AFBDD1D4}#1.0#0"; "SDOCdm.ocx"
Object = "{E64F71A6-E705-4151-9895-5138B7D67F3A}#1.0#0"; "CHPrj.ocx"
Begin VB.Form Menu 
   Caption         =   "MenuCashIn"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin CHPRJLib.CHPrj SDOPrj 
      Height          =   615
      Left            =   1440
      TabIndex        =   6
      Top             =   840
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   1085
      _StockProps     =   1
   End
   Begin TRANSPROVLibCtl.TransactionProvider S3ETrans 
      Height          =   660
      Left            =   120
      OleObjectBlob   =   "Menu.frx":0E42
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin S3EADVBROWSERLibCtl.AdvBrowser Browser 
      Height          =   420
      Left            =   165
      OleObjectBlob   =   "Menu.frx":0E76
      TabIndex        =   4
      Top             =   1620
      Width           =   1995
   End
   Begin SDOPrrLibCtl.SDOPrr S3EPrr 
      Height          =   645
      Left            =   1455
      OleObjectBlob   =   "Menu.frx":0E9C
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin SDOCdmLibCtl.SDOCdm S3ECdm 
      Height          =   645
      Left            =   120
      OleObjectBlob   =   "Menu.frx":0ECC
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer TimerAction 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1665
      Top             =   1680
   End
   Begin S3ELINEOUTLib.S3ELineOut S3ELineOut 
      Height          =   765
      Left            =   2670
      TabIndex        =   1
      Top             =   795
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   1349
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
      Height          =   660
      Left            =   2730
      TabIndex        =   0
      Top             =   120
      Width           =   1305
   End
   Begin DLLib.DL Pcb3Dl 
      Left            =   2685
      Top             =   840
      _Version        =   65538
      _ExtentX        =   2275
      _ExtentY        =   1164
      _StockProps     =   0
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All variable need to be declared early
Option Explicit

'��Ȩ˵��:�ϱ���˾�й���������
'�汾�ţ�Agilis power 1.6
'�������ڣ�2004.8
'���ߣ����֣���ʼ�棩
'ģ�鹦�ܣ� ���׻���ѡ��
'��Ҫ�������书��
' ȫ�ֱ���
'�޸���־
'================================================================
'�޸���־
'<ʱ��>��2005.8.23
'<�޸���>��������
'<��ϸ��¼>��
'         ����ע�ͣ�ɾ�����ñ�����������ʽ
'================================================================
'�޸���־
'<ʱ��>��2005.11.22
'<�޸���>��vincent jiang
'<��ϸ��¼>��
'    ��������ͨ��cardtype = 03��04�����Ӵŵ��ж��Ƿ���Ҫ��ʾ�����ʻ���
'    ��һ�����з��еĿ�cardtype = 04�����ŵ��ǿյģ��ܹ����ܣ����ܷ�����ANSI 98
'================================================================
'�޸���־
'<ʱ��>��2005.11.28
'<�޸���>��������
'�汾�ţ� 1.1.16 (2005.11.28)
'<��ϸ��¼>��
' ���ӱ��ؿ�����98������ת�ʵĿ���97������ת�ʲ��ܸ��ܵĿ�,ֻ��ȡ���ѯ
' 99��98��97ȫ�ڿ��������ã��ڱ�����PAN����ֻ����01:�����������ͱ�����ؿ���02����
'�Ϸ������ڿ���ֻ����99����PAN���׻᷵��01����������02��������ؿ���03������ͨ�ɿ���04������ͨ�¿���05�����н�ǿ�
'<ʱ��>��2005.12.20
'<�޸���>��������
'�汾�ţ� 1.2.16 (2005.12.20)
'���ӿ�����96������ת�ʡ�ȡ�ֻ�ܲ�ѯ�͸���
'================================================================
Private Enum pageType
    pageNothing = 0
    pageFirstMenu = 1
    pageSelectAccType = 3
    pageSelectAccType1 = 5
    pageMenuStop = 2
    pageSelectPrr = 4
    pageScreenError = 97
    pageError = 99
    pageQuit = 98
End Enum
Private currentPage As pageType

Private Enum ReturnType
    ResultTfrOut = 202
    ResultStop = 204
    ResultCwd = 209
    ResultPin = 211
    ResultInq = 212
    ReturnTimeout = 80
End Enum
Private S3eReturn       As ReturnType

Private nrc             As Integer

Dim IsSAN1 As Boolean, IsSAN2 As Boolean

Private Sub Form_Load()
    Dim sValue As String
    
    sValue = "The version number of " & App.EXEName & ".exe is " & App.Major & "." _
            & App.Minor & ".0." & App.Revision
    
    LogInfo (sValue)
        
    ' Reset the PcB3HtmlBrowser variables
    nrc = Pcb3Dl.DlSetCharRaw("HtmlFkeyList", "")
    nrc = Pcb3Dl.DlSetCharRaw("HtmlFkeyMap", "3855")
    
    S3ETrans.Available = True
End Sub
Private Sub S3ETrans_QuitTransaction()
    currentPage = pageQuit
    TimerAction.Interval = 1000
    TimerAction.Enabled = True
End Sub

Private Sub S3ETrans_StartTransaction(ByVal Action As Long)
    Dim strCardType As String
    
    nrc = Pcb3Dl.DlSetCharRaw("GBLAccType", "00")
    strCardType = Pcb3Dl.DlGetCharRaw("FitCardType")
    
    Start.Enabled = False
    IsSAN1 = False
    IsSAN1 = False
    
    If S3EPrr.Available = False And Action = 1 Then
        currentPage = pageSelectPrr
    Else
        'strCardType = "99"��ʾ���б��ؿ���ֱ�ӽ��뽻��
        '����cardtype���Ͷ��Ǿ���PAN���׺�õ���������Ҫ�Ƚ����ʻ�����ѡ�����
        If strCardType = "99" Then
            currentPage = pageFirstMenu
        Else
            'ֻ��04�࿨�������Ų�Ϊ�գ���Ҫȥ�жϸ����ʺţ���������Ϊ�յ�04���������и��ܹ��ܣ�����Ҫ����һ��·
            If (strCardType = "03" Or strCardType = "04") And Len(Pcb3Dl.DlGetCharRaw("FitTrack3Message")) > 2 Then
                Select Case GetHongKongCardPAN()
                Case "1"
                    currentPage = pageSelectAccType1
                Case "2"
                    nrc = Pcb3Dl.DlSetCharRaw("GBLAccType", "90")
                    currentPage = pageFirstMenu
                Case Else
                    currentPage = pageSelectAccType
                End Select
            Else
                If strCardType = "04" And Len(Pcb3Dl.DlGetCharRaw("FitTrack3Message")) < 2 Then
                    currentPage = pageFirstMenu
                Else
                    currentPage = pageSelectAccType
                End If
            End If
        End If
    End If

    TimerAction.Enabled = True
End Sub

Private Sub Start_Click()
Dim FkeyList        As String
    
    FkeyList = ""
    
    nrc = Pcb3Dl.DlSetCharRaw("HtmlFkeyList", FkeyList)
               
    If S3EPrr.Available = False Then
        currentPage = pageSelectPrr
    Else
        currentPage = pageFirstMenu
    End If
    TimerAction.Enabled = True
End Sub

Private Sub RetToMaster(ByVal S3eRetValue As Integer)
    S3ETrans.Result = S3eRetValue
End Sub

Private Sub TimerAction_Timer()
    Dim sSubStData    As String
    Dim bIsTimerAgain As Boolean
    Dim FkeyList      As String
    Dim strCardType   As String
    
    TimerAction.Enabled = False
    bIsTimerAgain = True

    Select Case currentPage
        Case pageSelectAccType
            nrc = ShowScreenSync(Browser, "Menu", "SelectAccType", sSubStData)
            If nrc = 0 Then
                Select Case sSubStData
                Case "@SAVING"
                    nrc = Pcb3Dl.DlSetCharRaw("GBLAccType", "01")
                Case "@CHECK"
                    nrc = Pcb3Dl.DlSetCharRaw("GBLAccType", "02")
                Case "@CREDIT"
                    nrc = Pcb3Dl.DlSetCharRaw("GBLAccType", "03")
                Case "@DEFAULT"
                    nrc = Pcb3Dl.DlSetCharRaw("GBLAccType", "00")
                Case Else
                    nrc = Pcb3Dl.DlSetCharRaw("GBLAccType", "00")
                End Select
                currentPage = pageFirstMenu
            ElseIf nrc = 91 Then
                RetToMaster ReturnTimeout
                bIsTimerAgain = False
            Else
                LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nrc)
                currentPage = pageScreenError
            End If
        'added by vincent for HongKong card 2005/11/07
        Case pageSelectAccType1
            If Not IsSAN1 Then
                FkeyList = FkeyList + "@Add1,"
            End If
            If Not IsSAN2 Then
                FkeyList = FkeyList + "@Add2,"
            End If
            nrc = Pcb3Dl.DlSetCharRaw("HtmlFkeyList", FkeyList)

            nrc = ShowScreenSync(Browser, "Menu", "SelectAccType1", sSubStData)
            If nrc = 0 Then
                Select Case sSubStData
                    Case "@Main" '---���ʻ�
                        nrc = Pcb3Dl.DlSetCharRaw("GBLAccType", "90")
                        currentPage = pageFirstMenu
                    Case "@Add1"   '---��һ�����ʻ�
                        nrc = Pcb3Dl.DlSetCharRaw("GBLAccType", "91")
                        currentPage = pageFirstMenu
                    Case "@Add2"  '---�ڶ������ʻ�
                        nrc = Pcb3Dl.DlSetCharRaw("GBLAccType", "92")
                        currentPage = pageFirstMenu
                    Case "@stop"
                        currentPage = pageMenuStop
                End Select
            ElseIf nrc = 91 Then
                RetToMaster ReturnTimeout
                bIsTimerAgain = False
            Else
                LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nrc)
                currentPage = pageScreenError
            End If
            
        Case pageFirstMenu
            FkeyList = ""
            Select Case Pcb3Dl.DlGetCharRaw("FitCardType")
            Case "98"
                FkeyList = FkeyList + "@TransferOut,"
            Case "97"
                FkeyList = FkeyList + "@PinChange," + "@TransferOut,"
            Case "96"
                FkeyList = FkeyList + "@TransferOut," + "@Withdrawal,"
            Case "01"
                    FkeyList = FkeyList + "@PinChange," + "@TransferOut,"
            Case "02"
                    FkeyList = FkeyList + "@PinChange," + "@TransferOut,"
            Case "03"
                    FkeyList = FkeyList + "@PinChange," + "@TransferOut,"
            Case "04"
                    '���и����˺�ѡ��Ŀ�����ʱ������ܹ��ܾ�Ҫ���
                    If Left(Trim(Pcb3Dl.DlGetCharRaw("GBLAccType")), 1) = "9" Then
                        FkeyList = FkeyList + "@PinChange," + "@TransferOut,"
                    Else
                        FkeyList = FkeyList + "@TransferOut,"
                    End If
            Case "05"
                    FkeyList = FkeyList + "@TransferOut,"
            End Select
            
            'if Cash Dispenser Module is not available,
            'the Withdrawal transaction should be disable.
            If S3ECdm.Available = False Or Pcb3Dl.DlGetCharRaw("CWDCrimePossible") <> "N" Then
                FkeyList = FkeyList + "@Withdrawal,"
            End If
            nrc = Pcb3Dl.DlSetCharRaw("HtmlFkeyList", FkeyList)
            
            nrc = ShowScreenSync(Browser, "Menu", "FirstMenu", sSubStData)
            If nrc = 0 Then
                Select Case sSubStData
                    Case "@Withdrawal"
                        S3eReturn = ResultCwd
                        RetToMaster S3eReturn
                        Exit Sub
                    
                    Case "@PinChange"
                        S3eReturn = ResultPin
                        RetToMaster S3eReturn
                        Exit Sub
                    
                    Case "@Inquiry"
                        S3eReturn = ResultInq
                        RetToMaster S3eReturn
                        Exit Sub
                    
                    Case "@TransferOut"
                        S3eReturn = ResultTfrOut
                        RetToMaster S3eReturn
                        Exit Sub
                    
                    Case "@stop"
                        currentPage = pageMenuStop
                    
                    Case Else
                        LogError "Case SubstData Error in pageFirstMenu, substData: " + _
                                Browser.SubStData
                        currentPage = pageScreenError
                End Select
            ElseIf nrc = 91 Then
                RetToMaster ReturnTimeout
                bIsTimerAgain = False
            Else
                LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nrc)
                currentPage = pageScreenError
            End If
            
            
        Case pageMenuStop
            nrc = ShowScreenSync(Browser, "Common", "ComPressStop", sSubStData)
            Call SendExceptionMessage(S3ELineOut, Pcb3Dl, "45")
            RetToMaster ResultStop
            Exit Sub
        
        Case pageSelectPrr
            nrc = ShowScreenSync(Browser, "Menu", "SelectPrr", sSubStData)
            If nrc = 0 Then
                If sSubStData = "@Continue" Then
                    'strCardType = "99"��ʾ���б��ؿ���ֱ�ӽ��뽻��
                    '����cardtype���Ͷ��Ǿ���PAN���׺�õ���������Ҫ�Ƚ����ʻ�����ѡ�����
                    If Pcb3Dl.DlGetCharRaw("FitCardType") = "99" Then
                        currentPage = pageFirstMenu
                    Else
                        
                        strCardType = Pcb3Dl.DlGetCharRaw("FitCardType")
                        
                        'ֻ��04�࿨�������Ų�Ϊ�գ���Ҫȥ�жϸ����ʺţ���������Ϊ�յ�04���������и��ܹ��ܣ�����Ҫ����һ��·
                        If (strCardType = "03" Or strCardType = "04") And Len(Pcb3Dl.DlGetCharRaw("FitTrack3Message")) > 2 Then
                            Select Case GetHongKongCardPAN()
                            Case "1"
                                currentPage = pageSelectAccType1
                            Case "2"
                                nrc = Pcb3Dl.DlSetCharRaw("GBLAccType", "90")
                                currentPage = pageFirstMenu
                            Case Else
                                currentPage = pageSelectAccType
                            End Select
                        Else
                            If strCardType = "04" And Len(Pcb3Dl.DlGetCharRaw("FitTrack3Message")) < 2 Then
                                currentPage = pageFirstMenu
                            Else
                                currentPage = pageSelectAccType
                            End If
                        End If
                    End If
                Else
                    currentPage = pageMenuStop
                End If
            ElseIf nrc = 91 Then
                RetToMaster ReturnTimeout
                bIsTimerAgain = False
            Else
                LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nrc)
                currentPage = pageScreenError
            End If
            
        Case pageScreenError
            bIsTimerAgain = False
            
        Case pageQuit
            Unload Menu
            Exit Sub
            
        Case Else
            LogError "TimerAction next action case error. The next action is:" + _
                CStr(currentPage)
    End Select
    
    If bIsTimerAgain = True Then
        TimerAction.Enabled = True
    End If

End Sub
'===================================================================================
'�������� :�ж��Ƿ���Ҫ��ʾ�����ʻ�
'������� �������ͣ�������Ϣ
'�����������
'����ֵ����
'���ú�����
'�����������
'���ߣ���ΰʤ
'����ʱ�� : 2005/11
'====================================================================================
Private Function GetHongKongCardPAN() As String
    Dim TypeOfPAN       As String, strPAN As String
    Dim TypeOfSAN1      As String, strSAN1 As String
    Dim TypeOfSAN2      As String, strSAN2 As String
    Dim IssuingIndustry As String, PrimaryID As String
    Dim Track3          As String
        
    Track3 = Pcb3Dl.DlGetCharRaw("FitTrack3Message")
    TypeOfPAN = Mid(Track3, 49, 2)
    TypeOfSAN1 = Mid(Track3, 51, 2)
    TypeOfSAN2 = Mid(Track3, 53, 2)
    IssuingIndustry = Mid(Track3, 3, 2)
    Select Case IssuingIndustry
    Case "49", "53"
        strPAN = Mid(Track3, 9, 12)
        strPAN = Mid(Track3, 3, 6) + Left(strPAN, 10)
    Case "23"
        strPAN = Mid(Track3, 9, 12)
        strPAN = Mid(Track3, 3, 6) + Right(strPAN, 11)
    
    Case "54"
        PrimaryID = Mid(Track3, 5, 4)
        If IsNumeric(PrimaryID) Then
            strPAN = Mid(Track3, 9, 12)
            If CInt(PrimaryID) >= 1150 Then '�˿�ΪM/C
                strPAN = Mid(Track3, 3, 6) + Left(strPAN, 10)
            Else '�˿�Ϊ������
                strPAN = Mid(Track3, 3, 6) + Right(strPAN, 11)
            End If
        Else
            strPAN = Mid(Track3, 3, 18)
        End If
    Case Else
        strPAN = Mid(Track3, 3, 18)
    End Select
        
    If TypeOfSAN1 <> "00" Then
        strSAN1 = Mid(Track3, 61, 12)
        strSAN1 = Mid(Track3, 35, 4) + Right(strSAN1, 11)
    End If
    If TypeOfSAN2 <> "00" Then
        strSAN2 = Mid(Track3, 73, 12)
        strSAN2 = Mid(Track3, 35, 4) + Right(strSAN2, 11)
    End If
        
    Select Case TypeOfPAN
    Case "10"
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt1", "�����˻�")
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt2", strPAN)
    Case "20"
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt1", "֧Ʊ�˻�")
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt2", strPAN)
    Case "30"
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt1", "���ÿ��˻�")
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt2", strPAN)
    Case Else
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt1", "�����˻�")
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt2", strPAN)
    End Select
        
    Select Case TypeOfSAN1
    Case "00"
        IsSAN1 = False
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt3", "")
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt4", "")
        
    Case "10"
        IsSAN1 = True
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt3", "�����˻�")
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt4", strSAN1)
        
    Case "20"
        IsSAN1 = True
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt3", "֧Ʊ�˻�")
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt4", strSAN1)
    Case "30"
        IsSAN1 = True
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt3", "���ÿ��˻�")
        nrc = Pcb3Dl.DlSetCharRaw("HtmlPrompt4", strSAN1)
    Case Else
        IsSAN1 = False
    End Select
        
    Select Case TypeOfSAN2
    Case "00"
        IsSAN2 = False
        nrc = Pcb3Dl.DlSetCharRaw("HtmlWork13", "")
        nrc = Pcb3Dl.DlSetCharRaw("HtmlWork23", "")
    Case "10"
        IsSAN2 = True
        nrc = Pcb3Dl.DlSetCharRaw("HtmlWork13", "�����˻�")
        nrc = Pcb3Dl.DlSetCharRaw("HtmlWork23", strSAN2)
    Case "20"
        IsSAN2 = True
        nrc = Pcb3Dl.DlSetCharRaw("HtmlWork13", "֧Ʊ�˻�")
        nrc = Pcb3Dl.DlSetCharRaw("HtmlWork23", strSAN2)
    Case "30"
        IsSAN2 = True
        nrc = Pcb3Dl.DlSetCharRaw("HtmlWork13", "���ÿ��˻�")
        nrc = Pcb3Dl.DlSetCharRaw("HtmlWork23", strSAN2)
    Case Else
        IsSAN1 = False
    End Select
    
    If TypeOfSAN1 = "00" And TypeOfSAN2 = "00" Then
     '��ʾ�ÿ�û�и����ʻ�����ֱ�ӽ���firstmenu
     GetHongKongCardPAN = "2"
    Else
     '��ʾ�ÿ�������һ�������ʻ�������selectAcctype1
     GetHongKongCardPAN = "1"
    End If
 
End Function
