VERSION 5.00
Object = "{B2110643-3E81-11D3-8ACC-00C04FF20A5D}#1.2#0"; "TransProv.dll"
Object = "{EACE4ECF-6930-11D0-AC6C-0020AFBDD1D4}#1.0#0"; "SDOEdm.ocx"
Object = "{192DFCF0-F664-11D3-8BD4-00C04FF20A5D}#1.1#0"; "AdvBrowser.ocx"
Object = "{7CCB2EF0-B3E8-11CF-BF8E-0020AF7093F9}#1.0#0"; "SDOPin.ocx"
Object = "{9C37E835-6A58-11D1-80C0-0020AF7093F9}#1.2#0"; "Dl.ocx"
Object = "{6C4DD4AB-27D5-11D3-96C4-000000000000}#1.0#0"; "S3ELineOutTcp.ocx"
Object = "{E64F71A6-E705-4151-9895-5138B7D67F3A}#1.0#0"; "CHPrj.ocx"
Begin VB.Form PinInput 
   Caption         =   "PinInput"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   Icon            =   "PinInput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin CHPRJLib.CHPrj SDOPrj 
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   1296
      _StockProps     =   1
   End
   Begin S3EADVBROWSERLibCtl.AdvBrowser Browser 
      Height          =   495
      Left            =   2520
      OleObjectBlob   =   "PinInput.frx":1272
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin SDOPinLibCtl.SDOPin SDOPin 
      Height          =   735
      Left            =   1680
      TabIndex        =   4
      Top             =   840
      Width           =   1215
      _cx             =   2143
      _cy             =   1296
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
   Begin TRANSPROVLibCtl.TransactionProvider SDOTrans 
      Height          =   690
      Left            =   1590
      OleObjectBlob   =   "PinInput.frx":1298
      TabIndex        =   3
      Top             =   60
      Width           =   1245
   End
   Begin S3ELINEOUTLib.S3ELineOut S3ELineOut 
      Height          =   705
      Left            =   150
      TabIndex        =   1
      Top             =   1590
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   1244
      _StockProps     =   1
   End
   Begin DLLib.DL Pcb3dl 
      Left            =   240
      Top             =   1635
      _Version        =   65538
      _ExtentX        =   2196
      _ExtentY        =   1111
      _StockProps     =   0
   End
   Begin SDOEdmLibCtl.SDOEdm SDOEdm 
      Height          =   690
      Left            =   210
      OleObjectBlob   =   "PinInput.frx":12D4
      TabIndex        =   2
      Top             =   45
      Width           =   1245
   End
   Begin VB.Timer TimerAction 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1605
      Top             =   1815
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
      Height          =   675
      Left            =   2955
      TabIndex        =   0
      Top             =   75
      Width           =   1335
   End
End
Attribute VB_Name = "PinInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All variable need to be declared early
Option Explicit

'==========================================================================================
'版权说明:迪堡公司中国区技术部
'版本号：Agilis power 1.6
'生成日期：2004.8
'作者：汪林(初始版）
'模块功能：输入密码模块
'主要函数及其功能
' 全局变量
'修改日志
'-----------------------------------------------------------------------------------------
'<时间>：2005。11。18
'<修改者>：孙世方
'<详细记录>：
'增加对中银通卡的处理，包括校验密码(函数VerifyPin)和按照NCR方式生成针对中银通卡的Pinblock(函数NCRPinBlock)
'<时间>：2005。11。22
'<修改者>: vincent
'<当前版本>：中行1.1.16
'<详细记录>：修改函数NCRPinBlock
'<时间>：2005。12。22
'<修改者>：孙世方
'<详细记录>：1 输入密码必须大于4位，2 修改genPinBlockSW，帐号小于12位的处理
'==========================================================================================
Const ReturnOk              As Integer = 10
Const ReturnPressStop       As Integer = 20
Const ReturnTimeout         As Integer = 80
Const GlobalINIPath         As String = "C:\atmwosa\ini\"

Const MULTI_INT             As String = "947124"
Const SHUFFLE_KEY           As String = "489624461835"

Public Enum pageType
    pageNothing = 0
    pagePinInputPin = 1
    pagePinInputPressStop = 2
    pagePinInputWelcome = 3
    pagePinInputPinEpp = 4
    pagePinInputEppNotSix = 5
    pagePinSelectLang = 6
    pageScreenError = 97
    pageError = 99
    pageQuit = 98
End Enum
Public currentPage          As pageType

Dim GLnAction               As Integer
Dim nrc                     As Integer
Dim bIsAntiFunctionKeyCrime As Boolean
Dim FriendHint(1 To 4)      As String
Dim g_sPrjLanguage          As String

Private Sub Form_Load()
    Dim sValue             As String
    Dim sIsEnableAntiCrime As String
    Dim i                  As Integer
    
    sValue = "The version number of " & App.EXEName & ".exe is " & App.Major & "." _
            & App.Minor & ".0." & App.Revision
    
    LogInfo (sValue)
    
    nrc = Pcb3dl.DlSetCharRaw("HtmlFkeyList", "")
    nrc = Pcb3dl.DlSetCharRaw("HtmlFkeyMap", "3855")
    
    sIsEnableAntiCrime = GetIniS(GlobalINIPath + "FrndHint.ini", "FriendHint", _
                    "FriendlyPromption", "")
    
    If sIsEnableAntiCrime = "Y" Then
        bIsAntiFunctionKeyCrime = True
        For i = 1 To 4
            FriendHint(i) = GetIniS(GlobalINIPath + "FrndHint.ini", "FriendHint", _
                        "FriendlyPromption" & CStr(i), "")
        Next i
    Else
        bIsAntiFunctionKeyCrime = False
    End If
    
    If GetIniS(GlobalINIPath + "Global.ini", "Bank_Environment", "PrjLanguage", "E") = "E" Then
        g_sPrjLanguage = "E"
    Else
        g_sPrjLanguage = "C"
    End If

    SDOTrans.Available = True
End Sub

Private Sub SDOTrans_QuitTransaction()
    currentPage = pageQuit
    TimerAction.Interval = 1000
    TimerAction.Enabled = True
End Sub

Private Sub Start_Click()
    Dim sAccNo     As String
    
    Call NCRPinBlock("121212", sAccNo)
    
    GLnAction = 1
    
    nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
    
    sAccNo = "00000000"
    
    nrc = Pcb3dl.DlSetCharRaw("SiabBGRCode", "0000")
    
    currentPage = pagePinInputWelcome
    
    TimerAction.Enabled = True
End Sub
Private Sub SDOTrans_StartTransaction(ByVal Action As Long)
    
    Start.Enabled = False
    GLnAction = Action
    nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
    
    If Action = 1 And bIsAntiFunctionKeyCrime = True Then
        'currentPage = pagePinSelectLang
        currentPage = pagePinInputWelcome
    Else
        If Pcb3dl.DlGetCharRaw("GBLEncrypType") = "H" Then
            currentPage = pagePinInputPinEpp
        Else
            currentPage = pagePinSelectLang
            'currentPage = pagePinInputPin
        End If
    End If

    TimerAction.Enabled = True
End Sub
Private Sub RetToMaster(ByVal SDORetValue As Integer)
    SDOTrans.Result = SDORetValue
End Sub
Private Sub TimerAction_Timer()
    Dim sSubStData       As String
    Dim bIsTimerAgain    As Boolean
    Dim PrjString        As String
    Dim PrjCHNString     As String
    Dim InputPin         As String
    Dim strPinBlock      As String
    Dim i                As Integer
    Dim bIsPinBlockOK    As Boolean
    Dim sScreenName      As String
    Dim PWDisWrong       As Boolean
    
    TimerAction.Enabled = False
    bIsTimerAgain = True

    Select Case currentPage
        '周亮修改
        Case pagePinSelectLang
            nrc = ShowScreenSync(Browser, "PinInput", "PinSelectLan", sSubStData)
            If nrc = 0 Then
                If sSubStData = "@CHN" Then
                    nrc = Pcb3dl.DlSetCharRaw("GBLSelectLan", "CHN")
                Else
                    nrc = Pcb3dl.DlSetCharRaw("GBLSelectLan", "ENG")
                End If
                currentPage = pagePinInputWelcome
            Else
                RetToMaster ReturnTimeout
                Exit Sub
            End If
    
        Case pagePinInputWelcome
            For i = 1 To 4
                If FriendHint(i) <> "" Then
                    Pcb3dl.DlSetCharRaw "HtmlWork" & CStr(i) & "3", CStr(i) & "." & FriendHint(i)
                End If
            Next i
            
            nrc = ShowScreenSync(Browser, "PinInput", "PinInputWelcome", sSubStData)
            If nrc = 0 Then
                If sSubStData <> "@stop" Then
                    If Pcb3dl.DlGetCharRaw("GBLEncrypType") = "H" Then
                        currentPage = pagePinInputPinEpp
                    Else
                        currentPage = pagePinInputPin
                    End If
                Else
                    currentPage = pagePinInputPressStop
                End If
            Else
                RetToMaster ReturnTimeout
                Exit Sub
            End If
        
        Case pagePinInputPin
            nrc = ShowScreenSync(Browser, "PinInput", "PinInputPin", sSubStData)
            
            If nrc = 0 Then
                Select Case sSubStData
                    Case "@ok":
                        InputPin = Pcb3dl.DlGetCharRaw("HtmlInput1")
                                                                    
                        If Len(InputPin) < 4 Then
                            nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
                            nrc = Pcb3dl.DlSetCharRaw("GBLPwdIsWrong", "Y")
                            currentPage = pagePinInputPin
                            TimerAction.Enabled = True
                            Exit Sub
                        Else
                            Pcb3dl.DlSetCharRaw "PinInputPin", InputPin
                            PWDisWrong = False
                            If Pcb3dl.DlGetCharRaw("FitCardType") = "04" Or Pcb3dl.DlGetCharRaw("FitCardType") = "03" Then
                                If Left(Pcb3dl.DlGetCharRaw("FitTrack3Message"), 2) = "92" Or Left(Pcb3dl.DlGetCharRaw("FitTrack3Message"), 2) = "93" Then
                                '对于中银通卡要预判密码
                                    If VerifyPin() <> 0 Then '密码预判不正确
                                        PWDisWrong = True
                                    End If
                                End If
                            End If
                        End If
                        If PWDisWrong Then
                            nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
                            nrc = Pcb3dl.DlSetCharRaw("GBLPwdIsWrong", "Y")
                            currentPage = pagePinInputPin
                        Else
                            nrc = Pcb3dl.DlSetCharRaw("GBLPwdIsWrong", "N")
                            '对于3磁信息小于48位的卡，加密时会报13错，另外存在这样的卡，卡类型是04，但确没有三磁，加密方法用ansi98，该卡叫总行双币卡
                            If (Pcb3dl.DlGetCharRaw("FitCardType") = "04" Or Pcb3dl.DlGetCharRaw("FitCardType") = "03") And Len(Pcb3dl.DlGetCharRaw("FitTrack3Message")) > 48 Then
                                Call NCRPinBlock(InputPin, strPinBlock)
                            Else
                                Call genPinBlockSW(InputPin, strPinBlock)
                            End If
                        
                            'Save the input PIN for PINChange
                            Pcb3dl.DlSetCharRaw "PinInputBlock", strPinBlock
                        
                            nrc = Pcb3dl.DlSetCharRaw("ShufflePinBlock", InputPin)
                            
                            If GLnAction <> 1 Then
                                RetToMaster GLnAction
                            Else
                                RetToMaster ReturnOk
                            End If
                            Exit Sub
                        End If
                    Case "@Update":
                        nrc = Pcb3dl.DlSetCharRaw("HtmlInput1", "")
                        currentPage = pagePinInputPin
                    Case "@stop"
                        currentPage = pagePinInputPressStop
                    Case Else
                        currentPage = pagePinInputPressStop
                  End Select
            ElseIf nrc = 91 Then
                RetToMaster ReturnTimeout
                Exit Sub
            Else
                LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nrc)
                currentPage = pageScreenError
            End If
        
        Case pagePinInputPinEpp, pagePinInputEppNotSix
        
            If currentPage = pagePinInputPinEpp Then
                sScreenName = "PinInputPinEpp"
            Else
                sScreenName = "PinInputEppNotSix"
            End If
            
            nrc = ShowScreenSync(Browser, "PinInput", sScreenName, sSubStData)
            
            If (nrc = 0 And sSubStData = "@ok") And (SDOPin.DigitsEntered < 4 Or SDOPin.DigitsEntered = 5 Or SDOPin.DigitsEntered > 6) Then
                currentPage = pagePinInputEppNotSix
            Else
                bIsPinBlockOK = True
                If nrc = 0 And sSubStData = "@ok" Then
                    strPinBlock = ""
                    bIsPinBlockOK = genPinBlockHW(strPinBlock)
                End If
                
                If bIsPinBlockOK = False Then
                    nrc = ShowScreenSync(Browser, "Common", "ComTimeOut", sSubStData)
                    Call SendExceptionMessage(S3ELineOut, Pcb3dl, "29")
                    PrjString = "  **EPP failed in PinInput"
                    PrjCHNString = "    **输入密码时加密键盘故障"
                    PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                    RetToMaster ReturnTimeout
                    Exit Sub
                End If
                
                If nrc = 0 Then
                    Select Case sSubStData
                        Case "@ok"
                            Pcb3dl.DlSetCharRaw "PinInputBlock", strPinBlock
                            
                            If GLnAction <> 1 Then
                                RetToMaster GLnAction
                            Else
                                RetToMaster ReturnOk
                            End If
                            Exit Sub
                        Case "@Change"
                            currentPage = pagePinInputPinEpp
                        Case "@stop"
                            currentPage = pagePinInputPressStop
                        Case "@dev_failed"      'only for EPP enable
                            nrc = ShowScreenSync(Browser, "Common", "ComPressStop", sSubStData)
                            Call SendExceptionMessage(S3ELineOut, Pcb3dl, "29")
                            PrjString = "  **EPP failed in PinInput"
                            PrjCHNString = "    **输入密码时加密键盘故障"
                            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                            RetToMaster ReturnPressStop
                            Exit Sub
                        Case "@dev_timeout"     'only for EPP enable
                            PrjString = "  **EPP TimeOut in PinInput"
                            PrjCHNString = "  **输入密码时加密键盘操作超时"
                            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
                            RetToMaster ReturnTimeout
                            Exit Sub
                    End Select
                Else
                    LogError ScreenInfo.Name + "Return error, nRc = " + CStr(nrc)
                    currentPage = pageScreenError
                End If
            End If
        
        Case pagePinInputPressStop
            nrc = ShowScreenSync(Browser, "Common", "ComPressStop", sSubStData)
            Call SendExceptionMessage(S3ELineOut, Pcb3dl, "45")
            PrjString = "  ##Customer Exit in PinInput."
            PrjCHNString = "  ##客户输入密码时退出."
            PrintJournal SDOPrj, PrjString, PrjCHNString, g_sPrjLanguage
            RetToMaster ReturnPressStop
            Exit Sub
        
        Case pageScreenError
            bIsTimerAgain = False
            
        Case pageQuit
            Unload PinInput
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
    Dim i   As Integer
    
    For i = 1 To 16 Step 2
        OutPar((i + 1) / 2) = Val("&H" + Mid(InPar, i, 2))
    Next
End Sub

Private Sub Bin2Str(ByRef InPar() As Byte, ByRef OutPar As String)
    Dim i          As Integer
    Dim strNum     As String

    For i = 1 To 8
        strNum = Hex(InPar(i))
        If Len(strNum) < 2 Then
            strNum = "0" + strNum
        End If
        OutPar = OutPar + strNum
    Next i
End Sub
Public Sub StrToBin(ByVal inString As String, ByRef bOutArray() As Byte, LenOfStr As Integer)
    Dim strTwo     As String
    Dim i          As Integer
    Dim j          As Integer

    j = 0
    For i = 1 To LenOfStr
        strTwo = Mid(inString, i, 1)
        bOutArray(j) = Hex(strTwo)
        j = j + 1
    Next

End Sub
'==========================================================================================
'函数功能 ：ANSI98 PinBlock软件算法
'输入参数 ：密码
'输出参数 ：PinBlock
'返回值   ：无
'调用函数 ：无
'被调用情况：
'作者  ：郭健
'创建时间 :2005
'==========================================================================================
Private Sub genPinBlockSW(pPin As String, ByRef pPinEnd As String)
    Dim iLoop                    As Integer
    Dim iLen                     As Integer
    
    Dim PinDataArray(1 To 8)     As Byte
    Dim PrePinDataArray(1 To 8)  As Byte
    Dim PreCardDataArray(1 To 8) As Byte
    
    Dim PrePinData               As String
    Dim CardData                 As String
    Dim PreCardData              As String
    Dim PinData                  As String
    Dim UseTriDES                As String
    Dim PinLen                   As String
    Dim DesResult                As String
    Dim strPinKey1               As String
    Dim strPinKey2               As String
    Dim strPinKey3               As String
    Dim strNum                   As String
    Dim AccountNo                As String
    
    AccountNo = Pcb3dl.DlGetCharRaw("FitAccNo")
    
    strPinKey1 = Pcb3dl.DlGetCharRaw("GBLPrePinKey")
    
    iLen = Len(pPin)
    If iLen < 10 Then
        PinLen = Format(CStr(Len(pPin)), "00")
    ElseIf iLen = 10 Then
        PinLen = "0A"
    ElseIf iLen = 11 Then
        PinLen = "0B"
    ElseIf iLen = 12 Then
        PinLen = "0C"
    End If
    
    PrePinData = PinLen + pPin + String(14 - iLen, "F")
    Call Str2Bin(PrePinData, PrePinDataArray)
    
    iLen = Len(AccountNo)
    If iLen > 12 Then
        PreCardData = Mid(AccountNo, iLen - 12, 12)
        CardData = String(4, "0") + PreCardData
    Else    '2005。12。22修改帐号小于12位的处理
        CardData = String(17 - Len(AccountNo), "0") + Left(AccountNo, Len(AccountNo) - 1)
    End If
    
    Call Str2Bin(CardData, PreCardDataArray)
    
    For iLoop = 1 To 8
        PinDataArray(iLoop) = PrePinDataArray(iLoop) Xor PreCardDataArray(iLoop)
    Next
    
    Call Bin2Str(PinDataArray, PinData)
    
    SDOEdm.CryptType = 1
    SDOEdm.CryptMode = True

    nrc = SDOEdm.DoCryptDataSw(PinData, strPinKey1)

    DesResult = SDOEdm.CryptResult
   
    For iLoop = 1 To Len(DesResult)
        strNum = Mid(DesResult, iLoop, 1)
        Select Case strNum
            Case "A"
                strNum = ":"
            Case "B"
                strNum = ";"
            Case "C"
                strNum = "<"
            Case "D"
                strNum = "="
            Case "E"
                strNum = ">"
            Case "F"
                strNum = "?"
        End Select
        pPinEnd = pPinEnd + strNum
    Next

End Sub
'==========================================================================================
'函数功能 ：ANSI98 PinBlock硬件算法
'输入参数 ：无
'输出参数 ：PinBlock
'返回值   ：无
'调用函数 ：无
'被调用情况：
'作者  ：郭健
'创建时间 :2005
'==========================================================================================
Private Function genPinBlockHW(ByRef pPinEnd As String) As Boolean
    Dim AccountNo     As String
    Dim CardData      As String
    Dim sTempStr      As String
    Dim nAccLen       As Integer
    
    AccountNo = Pcb3dl.DlGetCharRaw("FitAccNo")
    nAccLen = Len(AccountNo)
    If nAccLen > 12 Then
        CardData = Mid(AccountNo, nAccLen - 12, 12)
    Else
        sTempStr = "000000000000" + Left(AccountNo, nAccLen - 1)
        CardData = Right(sTempStr, 12)
    End If
    SDOPin.CustData = CardData
    SDOPin.KeyName = "PIN"
    
    nrc = SDOPin.DoPreparePin()
    If nrc <> 0 Then
        genPinBlockHW = False
    Else
        pPinEnd = SDOPin.EPin
        genPinBlockHW = True
    End If
    
End Function
'==========================================================================================
'函数功能 ：中银通卡校验密码
'输入参数 ：无
'输出参数 ：无
'返回值   ：0- 密码正确 1 - 密码错误
'调用函数 ：无
'被调用情况：输完密码后
'作者  :江伟胜
'创建时间 :2005.11
'==========================================================================================
Private Function VerifyPin() As Integer
    Dim sCurrentDate As String
    Dim sPinInputRaw As String
    Dim bPinBlock As String
    Dim PinCheckDigit As String
    Dim PinVerifyResult As Integer
    Dim i As Integer
    
    sPinInputRaw = Format(Pcb3dl.DlGetCharRaw("PinInputPin"), "000000")
    PinCheckDigit = Mid(Pcb3dl.DlGetCharRaw("FitTrack3Message"), 42, 1)
    
    PinVerifyResult = 0
    For i = 1 To 6
        PinVerifyResult = PinVerifyResult + ((Mid(sPinInputRaw, i, 1) * (7 - i)) Mod 10)
    Next
    
    PinVerifyResult = 10 - PinVerifyResult Mod 10
    If 10 = PinVerifyResult Then
        PinVerifyResult = 0
    End If
    
    If PinCheckDigit = PinVerifyResult Then
        VerifyPin = 0
    Else
        VerifyPin = 1
    End If
End Function
'==========================================================================================
'函数功能 ：NCR PinBlock软件算法
'输入参数 ：密码
'输出参数 ：PinBlock
'返回值   ：无
'调用函数 ：无
'被调用情况：输完密码后
'作者  ：孙世方
'创建时间 :2005.11.18
'==========================================================================================
Private Sub NCRPinBlock(pPin As String, ByRef pPinEnd As String)
    Dim Track3Data   As String
    Dim pinoffset    As String
    Dim lDifference  As String
    Dim Hi_Num       As String
    Dim Low_Num      As String
    Dim Shuffer_data As String
    Dim PinBlockByte As String
    Dim Pin_block    As String
    Dim Atm_code     As String
    Dim i            As Integer, iLoop As Integer
    Dim strPinKey1   As String
    Dim DesResult    As String
    Dim strNum       As String
    
    Atm_code = Pcb3dl.DlGetCharRaw("GBLAtmCode")
    Track3Data = Pcb3dl.DlGetCharRaw("FitTrack3Message")
    pinoffset = Mid(Track3Data, 43, 5)
   
    'for test
    'pinoffset = "30788"
    'Atm_code = "1001"
    
    lDifference = Left(pPin, 5) - pinoffset
    
    If lDifference < 0 Then
        lDifference = lDifference + 100000
    End If

    Hi_Num = Left(CStr(lDifference), 2) * MULTI_INT
    
    Hi_Num = Right(Hi_Num, 8) + "000"
    
    Low_Num = Right(lDifference, 3) * MULTI_INT
    
    Shuffer_data = CStr(CDbl(Hi_Num) + CDbl(Low_Num))
    Shuffer_data = Format(Shuffer_data, "000000000000")
    Pin_block = ""
    For i = 1 To 12
        PinBlockByte = Hex(Mid(Shuffer_data, i, 1) Xor Mid(SHUFFLE_KEY, i, 1))
        Pin_block = Pin_block + PinBlockByte
    Next
    
'    pPinEnd = Atm_code + Pin_block
    
    strPinKey1 = Pcb3dl.DlGetCharRaw("GBLPrePinKey")
    SDOEdm.CryptType = 1
    SDOEdm.CryptMode = True

    nrc = SDOEdm.DoCryptDataSw(CStr(Atm_code + Pin_block), strPinKey1)

    DesResult = SDOEdm.CryptResult
   
    For iLoop = 1 To Len(DesResult)
        strNum = Mid(DesResult, iLoop, 1)
        Select Case strNum
            Case "A"
                strNum = ":"
            Case "B"
                strNum = ";"
            Case "C"
                strNum = "<"
            Case "D"
                strNum = "="
            Case "E"
                strNum = ">"
            Case "F"
                strNum = "?"
        End Select
        pPinEnd = pPinEnd + strNum
    Next
    
End Sub
