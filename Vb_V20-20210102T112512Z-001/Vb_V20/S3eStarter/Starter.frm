VERSION 5.00
Object = "{B2110643-3E81-11D3-8ACC-00C04FF20A5D}#1.2#0"; "TransProv.dll"
Object = "{80C55DB1-86F3-11D3-8B2F-00C04FF20A5D}#1.0#0"; "S3ELineInTcp.ocx"
Object = "{DA559591-71AC-11D3-8B0E-00C04FF20A5D}#1.0#0"; "DlWait.ocx"
Object = "{EACE4ECF-6930-11D0-AC6C-0020AFBDD1D4}#1.0#0"; "SDOEdm.ocx"
Object = "{9C37E835-6A58-11D1-80C0-0020AF7093F9}#1.2#0"; "Dl.ocx"
Object = "{6C4DD4AB-27D5-11D3-96C4-000000000000}#1.0#0"; "S3ELineOutTcp.ocx"
Object = "{192DFCF0-F664-11D3-8BD4-00C04FF20A5D}#1.1#0"; "AdvBrowser.ocx"
Object = "{E64F71A6-E705-4151-9895-5138B7D67F3A}#1.0#0"; "CHPrj.ocx"
Object = "{D659C2E4-44CC-11D3-ACF9-00105A5F6CAB}#1.0#0"; "boclmk1.ocx"
Begin VB.Form Starter 
   Caption         =   "S3E Starter"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   Icon            =   "Starter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin DLWaitLibCtl.DLWait DLWaitResetTransKey 
      Height          =   495
      Left            =   2520
      OleObjectBlob   =   "Starter.frx":0ECA
      TabIndex        =   10
      Top             =   2280
      Width           =   1935
   End
   Begin S3ELINEINLibCtl.S3ELineIn S3ELineIn1 
      Height          =   855
      Left            =   1560
      OleObjectBlob   =   "Starter.frx":0F14
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin BOCLMK.BOCGDLMK BOCGDLMK 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2520
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin CHPRJLib.CHPrj S3EPrj 
      Height          =   615
      Left            =   1560
      TabIndex        =   8
      Top             =   120
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   1085
      _StockProps     =   1
   End
   Begin VB.CheckBox CheckKey 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Value           =   1  'Checked
      Width           =   255
   End
   Begin TRANSPROVLibCtl.TransactionProvider S3ETrans 
      Height          =   720
      Left            =   3000
      OleObjectBlob   =   "Starter.frx":0F3E
      TabIndex        =   5
      Top             =   135
      Width           =   1320
   End
   Begin SDOEdmLibCtl.SDOEdm SDOEdm 
      Height          =   690
      Left            =   225
      OleObjectBlob   =   "Starter.frx":0F78
      TabIndex        =   4
      Top             =   135
      Width           =   1230
   End
   Begin S3EADVBROWSERLibCtl.AdvBrowser Browser 
      Height          =   375
      Left            =   3030
      OleObjectBlob   =   "Starter.frx":0FA8
      TabIndex        =   3
      Top             =   1680
      Width           =   1410
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   960
   End
   Begin S3ELINEOUTLib.S3ELineOut S3ELineOut1 
      Height          =   630
      Left            =   225
      TabIndex        =   1
      Top             =   960
      Width           =   1170
      _Version        =   65536
      _ExtentX        =   2064
      _ExtentY        =   1111
      _StockProps     =   1
      BackColor       =   12582912
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   960
   End
   Begin DLLib.DL Pcb3DL1 
      Left            =   255
      Top             =   960
      _Version        =   65538
      _ExtentX        =   1984
      _ExtentY        =   1111
      _StockProps     =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Check Pin&&Mac Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1770
      Width           =   2940
   End
End
Attribute VB_Name = "Starter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==========================================================================================
'版权说明:  迪堡公司中国区技术部
'版本号：Agilis 1.6 for ICBC HQ OPTEVA
'生成日期：2005.8
'作者： 汪林(初始版）
'       李军(改进版)
'模块功能： 读文件中的配置参数，发送开机请求，交换密钥，打印版本号
'主要函数及其功能
' 全局变量
'修改日志
'-----------------------------------------------------------------------
'<时间>：[2005.6.6]
'<修改者>：李军
'<详细记录>：修改下载卡表不正常时的处理
'-----------------------------------------------------------------------
'<时间>：[2005.9.7]
'<修改者>：孙世方
'<详细记录>：中行客户化
'        ResetTransKey      Datalink变量，I-系统刚启动 N-交换完密钥  R-密钥不同步，需要重新换密钥
'==========================================================================================
'<时间>：[2005.11.29]
'<修改者>：孙世方
'<当前版本>：中行1.1.16
'<详细记录>：配合操作员申请密钥命令的修改，增加两种状态 O- 操作员申请密钥， W- 操作员申请密钥失败
'           操作员申请密钥中间过程有错误时不重复
'<时间>：[2005.12.20]
'<详细记录>：读取数字录象使用与否的配置,给GBLEVRUse赋值
'==========================================================================================
Private Const keySelfService = "Software\SelfService"

Private Const DEVICE_PRJ = &H200&
Private Const DEVICE_LINE_IN = &H10000000
Private Const DEVICE_LINE_OUT = &H20000000

Private Const sGlobalIni      As String = "C:\ATMWosa\Ini\global.ini"
Private Const sKeyIni         As String = "C:\ATMWosa\Ini\Key.ini"
Private Const sFitIni         As String = "C:\AtmWosa\Ini\Fit.ini"
Private Const sVersionIni     As String = "C:\ATMWosa\Ini\Version.ini"

Dim G_nCounter                As Long
Dim G_nDevicesToUse           As Long
Dim nrc                       As Integer
Dim G_sStartPeriodStatus      As String
Dim g_sProtocolType           As String
Dim G_bTrides                 As Boolean
Dim G_bIsHardware             As Boolean
Dim g_DisableXFSKey           As Boolean
Dim g_sPrjLanguage            As String
Dim g_sLocalMasterKey         As String
Dim g_nRQKTimes               As Integer
Dim g_sTerminalKey            As String
Dim g_sNewPinKey              As String
Dim g_sNewMACKey              As String

Private Sub DLWaitResetTransKey_VariableChanged()
    Dim sResetTransKey As String
    
    sResetTransKey = Pcb3DL1.DlGetCharRaw("ResetTransKey")
    LogInfo "WaitStartPeriod_VariableChanged, Dl(ResetTransKey) is " + sResetTransKey
            
    If sResetTransKey = "R" Or sResetTransKey = "O" Then
        G_nCounter = 8
        G_sStartPeriodStatus = "Y"
        Timer1.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Dim sSubStData   As String
    Dim sValue       As String
    Dim sAtmCode     As String
    Dim sBranchCode  As String
    
    sValue = "The version number of " & App.EXEName & ".exe is " & App.Major & "." _
            & App.Minor & ".0." & App.Revision
    
    LogInfo (sValue)
    LogInfo "Started"

    nrc = ShowScreenSync(Browser, "S3EStarter", "StartSystem", sSubStData)
    
    If GetIniS(sGlobalIni, "Bank_Environment", "PrjLanguage", "E") = "E" Then
        g_sPrjLanguage = "E"
    Else
        g_sPrjLanguage = "C"
    End If

    G_nCounter = 8
    G_nDevicesToUse = 0

    Call PrepareDataLink

'Add for checking CWDCrimePossible datalink value
    Call CheckCWDCrimePossible
'Add end
    DLWaitResetTransKey.Enabled = True
    
    nrc = Pcb3DL1.DlSetCharRaw("GBLGetAnomalies", "1")
    nrc = Pcb3DL1.DlSetCharRaw("ResetTransKey", "I")
    
    nrc = Pcb3DL1.DlSetCharRaw("GBLLineStatus", "C")
    nrc = Pcb3DL1.DlSetCharRaw("GBLLineSendNum", "00000")
    PrintStartupMessage
    
    G_sStartPeriodStatus = "Y"
    S3ETrans.Available = True
    
    sAtmCode = Pcb3DL1.DlGetCharRaw("GBLAtmCode")
    sBranchCode = Pcb3DL1.DlGetCharRaw("GBLBranchCode")
    g_nRQKTimes = 0
    
    BOCGDLMK.ATMID = sAtmCode
    BOCGDLMK.BranCode = sBranchCode
    Call BOCGDLMK.MakeLMK
    g_sLocalMasterKey = BOCGDLMK.LocalMK
    
    Timer1.Enabled = True
End Sub

'==========================================================================================
'函数功能:初始化DataLink变量
'输入参数:无
'输出参数:无
'返回值:无
'作者:
'创建时间:
'==========================================================================================
Private Sub PrepareDataLink()
    Dim sValue As String, sValue1 As String, sValue2 As String
    Dim sValueAK As String, sValueBK As String
    
    sValue = GetIniS(sGlobalIni, "Bank_Environment", "SystemEnvironment", "1")
    If Pcb3DL1.DlSetCharRaw("GBLSysEnvir", sValue) <> 0 Then
        LogError "DlSetCharRaw(GBLSysEnvir) failed"
    End If

    sValue = GetIniS(sGlobalIni, "Bank_Environment", "BankCode", "00000000")
    If Pcb3DL1.DlSetCharRaw("GBLBankCode", sValue) <> 0 Then
        LogError "DlSetCharRaw(GBLBankCode) failed"
    End If
    
    sValue = GetIniS(sGlobalIni, "Bank_Environment", "BranchCode", "0000")
    If Pcb3DL1.DlSetCharRaw("GBLBranchCode", sValue) <> 0 Then
        LogError "DlSetCharRaw(GBLBranchCode) failed"
    End If
    
    sValue = GetIniS(sGlobalIni, "Bank_Environment", "TrackPriority", "3")
    If Pcb3DL1.DlSetCharRaw("GBLTrackPriority", sValue) <> 0 Then
        LogError "DlSetCharRaw(GBLTrackPriority) failed"
    End If

    sValue = GetIniS(sGlobalIni, "Withdrawal", "MaxBills", "40")
    If Pcb3DL1.DlSetCharRaw("GBLMaxBills", sValue) <> 0 Then
        LogError "DlSetCharRaw(GBLMaxBills) failed"
    End If

    sValue = GetIniS(sGlobalIni, "Bank_Environment", "ATMCode", "00000000")
    If Pcb3DL1.DlSetCharRaw("GBLAtmCode", sValue) <> 0 Then
        LogError "DlSetCharRaw(GBLAtmCode) failed"
    End If
    
    sValue = GetIniS(sGlobalIni, "Bank_Environment", "AreaCode", "0000")
    If Pcb3DL1.DlSetCharRaw("IcbcAreaCode", sValue) <> 0 Then
        LogError "DlSetCharRaw(IcbcAreaCode) failed"
    End If
    
    '读取数字录象使用与否的配置
    sValue = GetIniS(sGlobalIni, "Bank_Environment", "EVR", "N")
    If Pcb3DL1.DlSetCharRaw("GBLEVRUse", sValue) <> 0 Then
        LogError "DlSetCharRaw(GBLEVRUse) failed"
    End If
    
    sValue = GetIniS(sGlobalIni, "Bank_Environment", "RegionalCode", "0000")
    If Pcb3DL1.DlSetCharRaw("IcbcRegionalCode", sValue) <> 0 Then
        LogError "DlSetCharRaw(IcbcRegionalCode) failed"
    End If
    If Pcb3DL1.DlSetCharRaw("LineLuNumber", Right(sValue, 3)) <> 0 Then
        LogError "DlSetCharRaw(LineLuNumber) failed"
    End If
       
    If GetIniS(sKeyIni, "keylist", "DESTYPE", "S") = "H" Then
        G_bIsHardware = True
        Pcb3DL1.DlSetCharRaw "GBLEncrypType", "H"
    Else
        G_bIsHardware = False
        Pcb3DL1.DlSetCharRaw "GBLEncrypType", "S"
    End If
    
    If GetIniS(sKeyIni, "keylist", "DESMETHOD", "S") = "T" Then
        G_bTrides = True
        Pcb3DL1.DlSetCharRaw "GBLEncrypMode", "T"
    Else
        G_bTrides = False
        Pcb3DL1.DlSetCharRaw "GBLEncrypMode", "S"
    End If
    
    If (G_bIsHardware = False) Then
    
        If (G_bTrides) Then
            
            sValue = GetIniS(sKeyIni, "KeyList", "AK", String(32, "0"))
            SDOEdm.CryptType = 1
            SDOEdm.CryptMode = False
            sValue1 = Left(sValue, 16)
            nrc = SDOEdm.DoCryptDataSw(sValue1, "1968707298242918")
            sValue1 = SDOEdm.CryptResult
            
            sValue2 = Right(sValue, 16)
            nrc = SDOEdm.DoCryptDataSw(sValue2, "1968707298242918")
            sValue2 = SDOEdm.CryptResult
            
            sValueAK = sValue1 + sValue2
            
            sValue = GetIniS(sKeyIni, "KeyList", "BK", String(32, "0"))
            sValue1 = Left(sValue, 16)
            nrc = SDOEdm.DoCryptDataSw(sValue1, "8192428927078691")
            sValue1 = SDOEdm.CryptResult
            
            sValue2 = Right(sValue, 16)
            nrc = SDOEdm.DoCryptDataSw(sValue2, "8192428927078691")
            sValue2 = SDOEdm.CryptResult
            
            sValueBK = sValue1 + sValue2
            
            sValue1 = Left(sValueAK, 16)
            sValue2 = Left(sValueBK, 16)
            sValue = AKeyXorBKey(sValue1, sValue2, 16)
            
            sValue1 = Right(sValueAK, 16)
            sValue2 = Right(sValueBK, 16)
            sValue = sValue + AKeyXorBKey(sValue1, sValue2, 16)
        Else
        
            sValue1 = GetIniS(sKeyIni, "KeyList", "AK", String(16, "0"))
            SDOEdm.CryptType = 1
            SDOEdm.CryptMode = False
            nrc = SDOEdm.DoCryptDataSw(sValue1, "1968707298242918")
            sValue1 = SDOEdm.CryptResult
            
            sValue2 = GetIniS(sKeyIni, "KeyList", "BK", String(16, "0"))
            nrc = SDOEdm.DoCryptDataSw(sValue2, "8192428927078691")
            sValue2 = SDOEdm.CryptResult
            
            sValue = AKeyXorBKey(sValue1, sValue2, 16)
        End If
        
        If Pcb3DL1.DlSetCharRaw("GBLMasterKey", sValue) <> 0 Then
            LogError "DlSetCharRaw(GBLMasterKey) failed"
        End If
        
    End If
    
    sValue = GetIniS(sGlobalIni, "CustomerInfo", "Telephone", "95588")
    If Pcb3DL1.DlSetCharRaw("GBLPhoneNumber", sValue) <> 0 Then
        LogError "DlSetCharRaw(GBLPhoneNumber) failed"
    End If
    
    G_nDevicesToUse = GetRegKeyN(HKEY_LOCAL_MACHINE, keySelfService, "DevicesToUse", 4, 0)
    
    'for BOC
    'g_sProtocolType = GetIniS(sSNACOMIni, "HostProtocol", "ProtocolType", "TCP")
    g_sProtocolType = "SNA"     'need modify!!!
End Sub
'==========================================================================================
'函数功能:通讯启动完成后的处理
'输入参数:启动通讯后的返回值
'输出参数:无
'返回值:无
'作者:
'创建时间:
'==========================================================================================
Private Sub S3ELineOut1_AtLineOpened(ByVal rcOpen As Integer)
    Dim PrjString         As String
    Dim PrjCHNString      As String
    Dim sTermMK           As String
    Dim DecryptResult     As String
    Dim DesResult         As String
    Dim sCheckValue       As String
    Dim sHostTranCode     As String
    Dim vRecvData         As Variant
 
    LogInfo "AtLineOpened " & CStr(rcOpen)
    If rcOpen = 0 Then
        
'for BOC
        sHostTranCode = Pcb3DL1.DlGetCharRaw("HostTransCode")
        If sHostTranCode <> "DNP" And sHostTranCode <> "ANP" Then     'accept
            PrjString = "  " & Format(Now(), "YYYY/MM/DD HH:MM:SS") & " Not ANP!" & vbCrLf
            PrjCHNString = "  " & Format(Now(), "YYYY/MM/DD HH:MM:SS") & " 报文头错!" & vbCrLf
            PrintJournal S3EPrj, PrjString, PrjCHNString, g_sPrjLanguage
            G_nCounter = 8
            If Pcb3DL1.DlGetCharRaw("ResetTransKey") = "O" Then
                Timer1.Enabled = False
                nrc = Pcb3DL1.DlSetCharRaw("ResetTransKey", "W")
            Else
                Timer1.Enabled = True
            End If
            Exit Sub
        End If
    
        'Get New Terminal Key from host
        nrc = S3ELineOut1.GetData("NewPinKey", vRecvData)
        sTermMK = vRecvData
        sTermMK = ConvertData("D", sTermMK)
        
        DecryptResult = DesDeCrypt(sTermMK, g_sLocalMasterKey)
        g_sTerminalKey = DecryptResult
        g_sTerminalKey = ConvertData("D", g_sTerminalKey)
        
        DesResult = DesEncrypt("0000000000000000", DecryptResult)
        'Get New Terminal Key check value from host
        nrc = S3ELineOut1.GetData("NewPinKeyCheck", vRecvData)
        sCheckValue = vRecvData
        sCheckValue = ConvertData("D", sCheckValue)
         
        If UCase(g_sProtocolType) = "SNA" Then
            If Mid(DesResult, 1, 4) <> sCheckValue Then
                PrjString = "  " & Format(Now(), "YYYY/MM/DD HH:MM:SS") & " PinCheck error!" & vbCrLf
                PrjCHNString = "  " & Format(Now(), "YYYY/MM/DD HH:MM:SS") & " 校验错!" & vbCrLf
                PrintJournal S3EPrj, PrjString, PrjCHNString, g_sPrjLanguage
                G_nCounter = 8
                If Pcb3DL1.DlGetCharRaw("ResetTransKey") = "O" Then
                    Timer1.Enabled = False
                    nrc = Pcb3DL1.DlSetCharRaw("ResetTransKey", "W")
                Else
                    Timer1.Enabled = True
                End If
                Exit Sub
            End If
        End If
        g_nRQKTimes = 0
'end of BOC
        
        PrjString = "  " & Format(Now(), "YYYY/MM/DD HH:MM:SS") & " AtLineOpen OK!" & vbCrLf
        PrjCHNString = "  " & Format(Now(), "YYYY/MM/DD HH:MM:SS") & " 与主机建立通讯连接!" & vbCrLf
        PrintJournal S3EPrj, PrjString, PrjCHNString, g_sPrjLanguage
        
        Timer2.Enabled = True
    Else
        Label1.Caption = "S3ELineOut AtLineOpened " & CStr(rcOpen)
        G_nCounter = 8
        If Pcb3DL1.DlGetCharRaw("ResetTransKey") = "O" Then
            Timer1.Enabled = False
            nrc = Pcb3DL1.DlSetCharRaw("ResetTransKey", "W")
        Else
            Timer1.Enabled = True
        End If
    End If
    
End Sub

Private Sub S3ETrans_QuitTransaction()
    Unload Starter
End Sub

Private Sub SDOEdm_AtLoadKeyStart()
    SDOEdm.UserReply = 0
End Sub

Private Sub SDOEdm_GetKey1()
    SDOEdm.UserReply = 0
End Sub

Private Sub SDOEdm_GetKey2()
    If SDOEdm.KeyName = "PIN" Then
        ' Load PINKey
        SDOEdm.UserReply = 200
    Else
        ' Load CDKey
        SDOEdm.UserReply = 100
    End If
End Sub

Private Sub SDOEdm_AtLoadKeyEnd(ByVal LoadKeyRc As Integer)
    Dim FinalResult As String
    Dim DataBuffer As String
    Dim sKeyName As String
    
    'ADD BY GUO JIAN 2004-06-17
    If g_DisableXFSKey = True Then
        g_DisableXFSKey = False
        Browser.XfsKeysEnabled = True
        LogInfo "SET Browser.XfsKeysEnabled: TRUE on AtLoadKeyEnd"
    End If
    
    If LoadKeyRc = 100 Then
    
    ElseIf LoadKeyRc = 200 Then
    
    Else
        nrc = SDOEdm.PuOpen
        nrc = Pcb3DL1.DlSetCharRaw("GBLDAMStatus", "C")
        nrc = Pcb3DL1.DlSetCharRaw("GBLLoadKeyStatus", "F")
        LogError "DoLoadKey in ResetATMWorkingKey return failed, " + CStr(LoadKeyRc)
    End If

End Sub

Private Sub Timer1_Timer()
    If G_nCounter > 0 Then
        Label1.Caption = "Seconds before open: " & CStr(G_nCounter)
        LogInfo "Seconds before open: " & CStr(G_nCounter)
        G_nCounter = G_nCounter - 1
    Else
        Timer1.Enabled = False
        OpenDevicesToUse
        S3ETrans.Result = 0
        G_nCounter = 8
 
    End If
End Sub

Private Sub OpenDevicesToUse()
On Error GoTo myerrhandler
    Dim nReply       As Integer
    Dim sCurrentDate As String
    
    If G_nDevicesToUse And DEVICE_LINE_IN Then
        Label1.Caption = "Opening LINE IN..."
        LogInfo "LINE_IN opened"
        S3ELineIn1.BackColor = &HFF00&
    End If
    If G_nDevicesToUse And DEVICE_LINE_OUT Then
        Label1.Caption = "Opening LINE OUT..."
     
        sCurrentDate = Format(Now(), "MMDDHHMMSS")
        S3ELineOut1.SetData "CurrentDate", sCurrentDate
        If G_bTrides Then
            S3ELineOut1.SetData "DesMode", "010"
        Else
            S3ELineOut1.SetData "DesMode", "011"
        End If
        S3ELineOut1.SetData "NetMIC", "001"
        nReply = S3ELineOut1.PuOpen()
        LogInfo "LINE_OUT.PuOpen returned " & CStr(nReply)
        If nReply = 0 Then
            S3ELineOut1.BackColor = &HFF00&
        Else
            S3ELineOut1.BackColor = &HFF
        End If
    End If
    Exit Sub
myerrhandler:
    LogError "Error:" & Err.Number & " " & Err.Description & " " & Err.Source
    Resume Next
End Sub
'==========================================================================================
'函数功能:打印开机信息
'输入参数:无
'输出参数:无
'返回值:无
'作者:
'创建时间:
'==========================================================================================
Private Sub PrintStartupMessage()
    Dim sMsg           As String
    Dim sProjectName   As String
    Dim sSysInfo       As String
    Dim PrjCHNString   As String
    
    'Print the information of Version and Patch
    sProjectName = GetIniS(sVersionIni, "Information", "Project", "")
    sSysInfo = "========================================"
    sSysInfo = sSysInfo + "***    Project: " + sProjectName + vbCrLf
    sSysInfo = sSysInfo + "========================================" + vbCrLf
    
    PrjCHNString = "========================================" + _
                    "***    项目名称: " + sProjectName + vbCrLf + _
                    "========================================" + vbCrLf
                    
    PrintJournal S3EPrj, sSysInfo, PrjCHNString, g_sPrjLanguage
    
    sMsg = "---------------------------------------" & vbCrLf
    sMsg = sMsg & "              System Start" & vbCrLf
    sMsg = sMsg & "---------------------------------------" & vbCrLf
    sMsg = sMsg & "Mod: Starter  Time: " & Format(Date$, "YYYY/MM/DD") & " " & Format(Time$(), "HH:MM:SS") & vbCrLf
    sMsg = sMsg & "Bank Code: " & Left(Pcb3DL1.DlGetCharRaw("GBLBankCode"), 5) & " ATM Code: " & Pcb3DL1.DlGetCharRaw("GBLAtmCode") & vbCrLf
    sMsg = sMsg & "---------------------------------------"

    PrjCHNString = "---------------------------------------" & vbCrLf + _
                    "      系　统　重　新　启　动" & vbCrLf + _
                    "---------------------------------------" & vbCrLf + _
                    "模块名: Starter  时间: " & Format(Date$, "YYYY/MM/DD") & " " & Format(Time$(), "HH:MM:SS") & vbCrLf + _
                    "银行号: " & Left(Pcb3DL1.DlGetCharRaw("GBLBankCode"), 5) & " ATM号: " & Pcb3DL1.DlGetCharRaw("GBLAtmCode") & vbCrLf
    
    If G_nDevicesToUse And DEVICE_PRJ Then
        PrintJournal S3EPrj, sMsg, PrjCHNString, g_sPrjLanguage
    Else
        LogInfo sMsg
    End If

End Sub
Private Sub Timer2_Timer()
    If G_nCounter > 0 Then
        G_nCounter = G_nCounter - 1
    Else
        Timer2.Enabled = False
        If G_sStartPeriodStatus = "N" Then
            Pcb3DL1.DlSetCharRaw "GBLDoRecovery", "1" 'modify the value to let Monitor to check whether it should do recovery
            LogInfo "Suspending"
            Me.WindowState = 1
        Else
            If BOCRQKOperation() Then
                G_sStartPeriodStatus = "N"
                G_nCounter = 5
            Else
                If Pcb3DL1.DlGetCharRaw("ResetTransKey") = "O" Then
                    nrc = Pcb3DL1.DlSetCharRaw("ResetTransKey", "W")
                    Exit Sub
                Else
                    G_nCounter = 5
                End If
                
            End If
            
            Timer2.Enabled = True
        End If
    End If
    Exit Sub
End Sub
'函数功能:检查之前是否有ATM犯罪
'输入参数:无
'输出参数:无
'返回值:无
'作者:
'创建时间:
'==========================================================================================
Sub CheckCWDCrimePossible()
    If Pcb3DL1.DlGetCharRaw("CWDCrimePossible") = "O" Then
        Pcb3DL1.DlSetCharRaw "GBLDoRecovery", "O"
    ElseIf Pcb3DL1.DlGetCharRaw("CWDCrimePossible") = "Y" Then
        Pcb3DL1.DlSetCharRaw "GBLDoRecovery", "C"
    End If
End Sub
'函数功能:更换pinkey,mackey
'输入参数:无
'输出参数:无
'返回值:布尔变量
'作者:郭健
'创建时间:
'==========================================================================================
Private Function BOCRQKOperation() As Boolean
    Dim sNewPinKey    As String
    Dim DecryptResult As String
    Dim DesResult     As String
    Dim sCheckValue   As String
    Dim sNewMACKey    As String
    Dim sTransCode    As String
    Dim vRecvData     As Variant

On Error GoTo myerrhandler

    BOCRQKOperation = False
    
    Timer1.Enabled = False
    g_nRQKTimes = g_nRQKTimes + 1
    If g_nRQKTimes > 3 Then
        g_nRQKTimes = 0
        Timer2.Enabled = False
        G_nCounter = 8
        Timer1.Enabled = True
        Exit Function
    End If
    
    Label1.Caption = "Send RQK... "
    S3ELineOut1.SetData "RQKMode", "C"
    nrc = S3ELineOut1.DoSend("RQK", 0)
    If nrc <> 0 Then
        Label1.Caption = "Send RQK Failed: " + CStr(nrc)
        Exit Function
    End If
    
    nrc = S3ELineOut1.DoReceive
    If nrc <> 0 Then
        Label1.Caption = "Receive RQK Failed: " + CStr(nrc)
        Exit Function
    End If
    
    sTransCode = Pcb3DL1.DlGetCharRaw("HostTransCode")
    If sTransCode <> "ABP" And sTransCode <> "DBP" Then    'accept
        Label1.Caption = "RQK IS NOT ABP: " + sTransCode
        Exit Function
    End If
    
    nrc = S3ELineOut1.GetData("NewPinKey", vRecvData)
    sNewPinKey = vRecvData
    sNewPinKey = ConvertData("D", sNewPinKey)
    
    DecryptResult = DesDeCrypt(sNewPinKey, g_sTerminalKey) 'RQK 回来的NewPinKey
    
    '检查checkvalue
    DesResult = DesEncrypt("0000000000000000", DecryptResult)
    nrc = S3ELineOut1.GetData("NewPinKeyCheck", vRecvData)
    sCheckValue = vRecvData
    sCheckValue = ConvertData("D", sCheckValue)
    
    If UCase(g_sProtocolType) = "SNA" Then
        If Mid(DesResult, 1, 4) <> sCheckValue Then
            Label1.Caption = "PinKey IS NOT Match!!"
            Exit Function
        End If
    End If
    nrc = Pcb3DL1.DlSetCharRaw("GBLPrePinKey", DecryptResult)
    'NewPinKey检查完成
    
    nrc = S3ELineOut1.GetData("NewMacKey", vRecvData)
    sNewMACKey = vRecvData
    sNewMACKey = ConvertData("D", sNewMACKey)
    
    DecryptResult = DesDeCrypt(sNewMACKey, g_sTerminalKey) 'RQK 回来的NewMacKey
    g_sNewMACKey = DecryptResult
    
    '检查checkvalue
    DesResult = DesEncrypt("0000000000000000", DecryptResult)
    nrc = S3ELineOut1.GetData("NewMacKeyCheck", vRecvData)
    sCheckValue = vRecvData
    sCheckValue = ConvertData("D", sCheckValue)
    
    If UCase(g_sProtocolType) = "SNA" Then
        If Mid(DesResult, 1, 4) <> sCheckValue Then
            Label1.Caption = "MACKey IS NOT Match!!"
            Exit Function
        End If
    End If
    'NewMacKey检查完成
    
    S3ELineOut1.MacSwSK = g_sNewMACKey
    Label1.Caption = "Finish RQK!!!"
    LogInfo "PinKey = " + g_sNewPinKey + "  MACKey = " + g_sNewMACKey
    
    g_nRQKTimes = 0
    nrc = Pcb3DL1.DlSetCharRaw("ResetTransKey", "N")
        
    BOCRQKOperation = True
    Exit Function

myerrhandler:
    MsgBox "Error:" & Err.Number & " " & Err.Description & " " & Err.Source
    Resume Next
End Function
Function AKeyXorBKey(ByVal ABuffer As String, ByVal BBuffer As String, ByVal InSize As Integer) As String
    Dim i                 As Integer
    Dim AArray(1 To 8)    As Byte
    Dim BArray(1 To 8)    As Byte
    Dim ABArray(1 To 8)   As Byte
    Dim XorResult         As String

    Call Str2Bin(ABuffer, AArray)
    Call Str2Bin(BBuffer, BArray)

    For i = 1 To InSize / 2
        ABArray(i) = AArray(i) Xor BArray(i)
    Next
    XorResult = ""
    Call Bin2Str(ABArray, XorResult)
    AKeyXorBKey = XorResult

End Function

Private Sub Str2Bin(ByVal InPar As String, ByRef OutPar() As Byte)
    Dim i As Integer
    
    For i = 1 To 16 Step 2
        OutPar((i + 1) / 2) = Val("&H" + Mid(InPar, i, 2))
    Next
End Sub

Private Sub Bin2Str(ByRef InPar() As Byte, ByRef OutPar As String)
    Dim i         As Integer
    Dim strNum    As String

    For i = 1 To 8
        strNum = Hex(InPar(i))
        If Len(strNum) < 2 Then
            strNum = "0" + strNum
        End If
        OutPar = OutPar + strNum
    Next

End Sub

Private Function DesDeCrypt(DesData As String, DesKey As String) As String
    Dim nrc As Integer
    SDOEdm.CryptMode = False
    
    SDOEdm.CryptType = 1
    nrc = SDOEdm.DoCryptDataSw(DesData, DesKey)
    If nrc = 0 Then
        DesDeCrypt = SDOEdm.CryptResult
        Exit Function
    Else
        DesDeCrypt = ""
    End If
End Function

Private Function DesEncrypt(DesData As String, DesKey As String) As String
    Dim nrc As Integer
    SDOEdm.CryptMode = True

    SDOEdm.CryptType = 1
    nrc = SDOEdm.DoCryptDataSw(DesData, DesKey)
    If nrc = 0 Then
        DesEncrypt = SDOEdm.CryptResult
        Exit Function
    Else
        DesEncrypt = ""
    End If
End Function
'函数功能:转换字符
'输入参数:转换标志，转换数据
'输出参数:无
'返回值:转换后的字符串
'作者:谭立科
'创建时间:2005.10.11
'==========================================================================================
Private Function ConvertData(ConvertFlag As String, ConData As String) As String
    Dim strNum         As String
    Dim i              As Integer
    
    Select Case ConvertFlag
     Case "D"
          For i = 1 To Len(ConData)
            strNum = Mid(ConData, i, 1)
            Select Case strNum
                Case ":"
                    strNum = "A"
                Case ";"
                    strNum = "B"
                Case "<"
                    strNum = "C"
                Case "="
                    strNum = "D"
                Case ">"
                    strNum = "E"
                Case "?"
                    strNum = "F"
            End Select
            ConvertData = ConvertData + strNum
        Next
        
    Case "E"
     
          For i = 1 To Len(ConData)
            strNum = Mid(ConData, i, 1)
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
            ConvertData = ConvertData + strNum
        Next
    End Select
    
End Function

