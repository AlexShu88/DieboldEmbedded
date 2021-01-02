Attribute VB_Name = "Glboal"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function LogWriteEntry Lib "Logger.dll" (ByVal Source As String, ByVal Message As String, ByVal Level As Long, ByVal msgID As Long) As Integer
Public Declare Function IcbcAccCheck Lib "AtmUtil.dll" (ByVal AccNo As Variant, ByVal sLen As Integer) As Integer
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'For registry operation
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Option Explicit
Enum RootKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Public Const LOG_ERROR = &H2
Public Const LOG_WARNING = &H4
Public Const LOG_INFORMATION = &H8

''''''For Show Screen and timer switch''''''''''''
Public Type ScreenTotalInfo
   Name As String
   Interval As Long
   Sound As String
End Type
Public ScreenInfo As ScreenTotalInfo
''''''For Show Screen and timer switch END''''''''

Enum ValueType
    REG_NONE = 0
    REG_SZ = 1         'StringData
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_BIG_ENDIAN = 5
    REG_MULTI_SZ = 7   'MultiString
End Enum

Enum ErrorCode
    ERROR_SUCCESS = 0&
    ERROR_MORE_DATA = 234&
End Enum
Private Const sMMDIni            As String = "C:\ATMWosa\Ini\MMDCode.ini"

Public Sub LogInfo(Msg As String)
    Dim nReply As Integer
    Dim CurRecordSource As String
    Dim CurRecordMsgId As Long
    
    CurRecordSource = App.Title
    If Len(CurRecordSource) = 0 Then
        CurRecordSource = "AgilisDefault"
    End If
    
    If IsNumeric(App.FileDescription) Then
        CurRecordMsgId = CLng(App.FileDescription)
    Else
        CurRecordMsgId = 6000
    End If
    
    nReply = LogWriteEntry(CurRecordSource, Msg, LOG_INFORMATION, CurRecordMsgId)

End Sub

Public Sub LogWarning(Msg As String)
    Dim nReply As Integer
    Dim CurRecordSource As String
    Dim CurRecordMsgId As Long
    
    CurRecordSource = App.Title
    If Len(CurRecordSource) = 0 Then
        CurRecordSource = "AgilisDefault"
    End If
    
    If IsNumeric(App.FileDescription) Then
        CurRecordMsgId = CLng(App.FileDescription)
    Else
        CurRecordMsgId = 6000
    End If
    
    nReply = LogWriteEntry(CurRecordSource, Msg, LOG_WARNING, CurRecordMsgId)

End Sub

Public Sub LogError(Msg As String)
    Dim nReply As Integer
    Dim CurRecordSource As String
    Dim CurRecordMsgId As Long
    
    CurRecordSource = App.Title
    If Len(CurRecordSource) = 0 Then
        CurRecordSource = "AgilisDefault"
    End If
    
    If IsNumeric(App.FileDescription) Then
        CurRecordMsgId = CLng(App.FileDescription)
    Else
        CurRecordMsgId = 6000
    End If
    
    nReply = LogWriteEntry(CurRecordSource, Msg, LOG_ERROR, CurRecordMsgId)

End Sub

Public Function GetScreenInfo(ByVal sStr As String) As ScreenTotalInfo
    Dim pPos As Integer
    Dim tmpStr As String
    
    pPos = InStr(1, sStr, ";")
    If pPos > 1 Then
        GetScreenInfo.Name = Trim(Left(sStr, pPos - 1))
        sStr = Right(sStr, Len(sStr) - pPos)
        pPos = InStr(1, sStr, ";")
        If pPos > 1 Then
            tmpStr = Trim(Left(sStr, pPos - 1))
            If IsNumeric(tmpStr) Then
                GetScreenInfo.Interval = CDbl(tmpStr) * 1000
            Else
                GetScreenInfo.Interval = 0
            End If
        Else
            If IsNumeric(sStr) Then
                GetScreenInfo.Interval = CDbl(sStr) * 1000
            Else
                GetScreenInfo.Interval = 0
            End If
            GetScreenInfo.Sound = ""
        End If
        
    Else
        GetScreenInfo.Name = Trim(sStr)
        GetScreenInfo.Interval = 0
        GetScreenInfo.Sound = ""
    End If
    
End Function

Function GetIniS(ByVal AppProfileName As String, ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String) As String

    Dim ResultString As String * 144, Temp As Integer
    Dim s As String, i As Integer
    Temp = GetPrivateProfileString(SectionName, KeyWord, "", ResultString, 144, AppProfileName)

    If Temp% > 0 Then ' not null
        s = ""
        For i = 1 To 144
            If Asc(Mid$(ResultString, i, 1)) = 0 Then
                Exit For
            Else
                s = s & Mid$(ResultString, i, 1)
            End If
        Next
    Else
        s = DefString
    End If
    GetIniS = s
End Function
Function SetIniS(ByVal AppProfileName As String, ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String) As Boolean
    Dim res As Integer
    res = WritePrivateProfileString(SectionName, KeyWord, ValStr, AppProfileName)
    If res <> 0 Then
        SetIniS = True
    Else
        SetIniS = False
    End If
End Function

Function GetIniN(ByVal AppProfileName As String, ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Integer) As Long
    GetIniN = GetPrivateProfileInt(SectionName, KeyWord, DefValue, AppProfileName)
End Function
Function GetIniD(ByVal AppProfileName As String, ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Double) As Double
    Dim tValue As String
    tValue = GetIniS(AppProfileName, SectionName, KeyWord, "")
    If IsNumeric(tValue) Then
        GetIniD = CDbl(tValue)
    Else
        GetIniD = DefValue
    End If
End Function
Function GetRegKeyS(ByVal RootKeyName As Long, ByVal SubRegKeyPath As String, ByVal KeyValue As String, ByVal DefSize As Long, ByVal DefValue As String) As String
    Dim nReply As Long
    Dim hKey As Long, nSize As Long
    Dim sDispArray() As Byte
    
    nReply = RegOpenKey(RootKeyName, SubRegKeyPath, hKey)
    If ERROR_SUCCESS = nReply Then
        nSize = DefSize
        ReDim sDispArray(0 To nSize)
        RegQueryValueEx hKey, KeyValue, 0, REG_BINARY, sDispArray(0), nSize
        ReDim Preserve sDispArray(0 To nSize - 2)
        GetRegKeyS = StrConv(sDispArray, vbUnicode)
    Else
        GetRegKeyS = DefValue
    End If
    RegCloseKey hKey
End Function

Function GetRegKeyN(ByVal RootKeyName As Long, ByVal SubRegKeyPath As String, ByVal KeyValue As String, ByVal DefSize As Long, ByVal DefValue As Long) As Long
    Dim nReply As Long
    Dim hKey As Long, nSize As Long
    Dim lResult As Long
    
    nReply = RegOpenKey(RootKeyName, SubRegKeyPath, hKey)
    If ERROR_SUCCESS = nReply Then
        nSize = DefSize
        RegQueryValueEx hKey, KeyValue, 0, REG_NONE, lResult, nSize
        GetRegKeyN = lResult
    Else
        GetRegKeyN = DefValue
    End If
    RegCloseKey hKey
End Function

Function SetRegKeyValue(ByVal RootKeyName As Long, ByVal SubRegKeyPath As String, ByVal KeyValue As String, ByVal vType As Long, Value As Variant, Optional ByVal lenValue As Long) As Boolean
    Dim bRet As Boolean, bArr() As Byte
    Dim hKey As Long
    Dim nReply As Integer
    Dim i As Integer
    
    bRet = False
    
    On Error GoTo ErrorExit
    
    nReply = RegOpenKey(RootKeyName, SubRegKeyPath, hKey)
    If ERROR_SUCCESS = nReply Then
        Select Case vType
            Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                If Len(Value) = 0 Then
                    nReply = RegSetValueEx(hKey, KeyValue, 0&, vType, 0, 1)
                Else
                    nReply = RegSetValueEx(hKey, KeyValue, 0&, vType, ByVal CStr(Value), LenB(StrConv(Value, vbFromUnicode)) + 1)
                End If
            Case REG_DWORD, REG_DWORD_BIG_ENDIAN
                nReply = RegSetValueEx(hKey, KeyValue, 0&, vType, CLng(Value), 4)
            Case REG_BINARY
                
                ReDim bArr(0 To lenValue - 1)
                For i = 0 To lenValue - 1
                    bArr(i) = Value(i)
                Next
                nReply = RegSetValueEx(hKey, KeyValue, 0&, vType, bArr(0), lenValue)
        End Select
        bRet = True
    End If
    
    RegCloseKey hKey
    
    SetRegKeyValue = bRet
    
ErrorExit:

End Function

Function GetAnomalyCode(ByVal ExpCode As String) As String
    GetAnomalyCode = GetIniS("c:\AtmWosa\ini\Anomaly.ini", "AnomalyTable", ExpCode, "")
End Function

Function ShowScreenSync(Browser As AdvBrowser, ByVal Section As String, ByVal ScreenName As String, ByRef SubStData As String) As Integer
On Error GoTo ErrHandler
    Dim Rc As Integer
    Dim sStr As String, Path As String
    
    sStr = GetIniS("Screens.ini", Section, ScreenName, "")
    ScreenInfo = GetScreenInfo(sStr)

    Path = GetIniS("Screens.ini", Section, "path", "")
    Rc = Browser.DoShowScreenSync(Trim(Path) + "\" + ScreenInfo.Name, ScreenInfo.Interval / 1000)
    If Rc = 0 Then
        SubStData = Browser.SubStData
    ElseIf Rc = 99 Then
        Rc = Browser.DoShowScreenSync(Trim(Path) + "\" + ScreenInfo.Name, ScreenInfo.Interval / 1000)
        If Rc = 0 Then
            SubStData = Browser.SubStData
        End If
    End If
    ShowScreenSync = Rc
    Exit Function

ErrHandler:
    ShowScreenSync = ErrorHandlerFunction("ShowScreenSync:" + ScreenInfo.Name, 91)
    Exit Function
End Function

Public Sub ResetATMPrr(ByVal InFormPcB3Dl As DL)
    Dim nReply As Byte

    nReply = InFormPcB3Dl.DlReset("PrrTransAmount")
    nReply = InFormPcB3Dl.DlReset("PrrTfr2ndAccNo")
    nReply = InFormPcB3Dl.DlReset("PrrCashInMark")
    nReply = InFormPcB3Dl.DlReset("PrrWthMark")
    nReply = InFormPcB3Dl.DlReset("PrrTransferMark")
    nReply = InFormPcB3Dl.DlReset("PrrOthersMark")
    nReply = InFormPcB3Dl.DlReset("PrrCardRetainMark")
    nReply = InFormPcB3Dl.DlReset("PrrContactBankMark")
    nReply = InFormPcB3Dl.DlReset("PrrAcceptMark")
    nReply = InFormPcB3Dl.DlReset("PrrAcceptCode")
    nReply = InFormPcB3Dl.DlReset("PrrRejectedCode")
    
    nReply = InFormPcB3Dl.DlReset("PrrTransType")
    nReply = InFormPcB3Dl.DlReset("PrrFeeCharge")
    nReply = InFormPcB3Dl.DlReset("PrrHostEnqNo")


End Sub


'Sub TranslogAdd(ByVal Pcb3Dl As DL, ByVal TransCode As String, _
'                    ByVal ExpCode As String)
'    Dim strTranslogRecord As String
'    strTranslogRecord = Format(Now(), "MMDD|HH:MM") + "|" + _
'                        Format(Pcb3Dl.DlGetCharRaw("GBLAtmCode")) + "|" + _
'                        TransCode + "|" + _
'                        Format(Pcb3Dl.DlGetCharRaw("GBLLineSendNum"), "@@@@") + "|" + _
'                        Format(" ", "@@@@") + "|" + _
'                        Format(Pcb3Dl.DlGetCharRaw("FitAccNo"), "@@@@@@@@@@@@@@@@@@@!") + "|" + _
'                        Format(" ", "@@@@@@@@@@@@@@@@@@@@!") + "|" + _
'                        "00000000" + "|" + _
'                        Format(ExpCode, "@@@@")
'
'
'    Open "C:\TransLog.txt" For Append As #1
'
'    Print #1, strTranslogRecord
'    Close #1
'End Sub

Sub SaveCNJournal(ByVal PrjString As String)
    Open "C:\S3E\Logs\LogTo\CNJournal.txt" For Append As #1
    Print #1, PrjString
    Close #1
End Sub

'===================================================================================
'函数功能 :发送例外通讯报文
'输入参数 ：例外代码
'输出参数：无
'返回值：无
'调用函数：GetAnomalyCode
'被调用情况：
'作者：
'创建时间 : 2004
'====================================================================================
Sub SendExceptionMessage(ByVal S3ELineOut As S3ELineOut, ByVal Pcb3Dl As DL, ByVal ExpCode As String)
    Dim sAnomalyCode       As String
    Dim sCurrentDate       As String
    Dim nrc                As Integer
    
    sAnomalyCode = GetAnomalyCode(ExpCode)
    If Len(sAnomalyCode) = 0 Or sAnomalyCode = "9999" Then
        LogInfo "Do not send exception, ExpCode = " + ExpCode
        Exit Sub
    End If
    
    sCurrentDate = Format(Now(), "MMDDHHMMSS")
    S3ELineOut.SetData "CurrentDate", sCurrentDate
    
    nrc = S3ELineOut.SetData("ExceptionCode", sAnomalyCode)
    
    nrc = S3ELineOut.DoSend("OEX", 0)
    
End Sub

Sub PrintJournal(ByVal CHPrj As CHPrj, ByVal PrjString As String, _
                ByVal PrjCHNString As String, ByVal sPrjLanguage As String)
    If sPrjLanguage = "E" Then
        CHPrj.DoPrint PrjString
        SaveCNJournal PrjCHNString
    Else
        CHPrj.DoPrintCH PrjCHNString
    End If

End Sub
Sub SaveApplErrorInfo(ByVal szInfo As String)
    Dim szMsg                       As String
    
    If Len(szInfo) > 0 Then
        LogError szInfo
        
        szMsg = Format(Now(), "YYYY-MM-DD HH:MM:SS ") + szInfo
        Open "C:\S3E\Logs\LogTo\ApplErrorInfo.txt" For Append As #1
        Print #1, szMsg
        Close #1
    End If
End Sub

Function ErrorHandlerFunction(ByVal szInfo As String, ByVal nRet As Integer) As Integer
    Dim szMsg                       As String
    
    szMsg = szInfo + " On " + Err.Source + " Error=" + CStr(Err.Number) + "-> " + Err.Description
    SaveApplErrorInfo "  ##程序异常[" + szMsg + "]."
    LogError szMsg
    Err.Clear
    
    ErrorHandlerFunction = nRet
End Function

'==========================================================================================
'函数的功能 :检查是否需要有条件复位
'输入参数 : 硬件代码
'输出参数 : 无
'返回值   : 是否需要复位
'调用函数 :无
'被调用情况  ：
'作者       ：赵文明、李军
'创建时间   :2005.7
'-----------------------------------------------
'<时间>：[2005.12.27]
'<修改者>：孙世方
'    增加一个参数，用于取款程序判断情况需要冲正的代码，在文件中找到代码返回False
'==========================================================================================
Function CheckSPInfo(SPInfo As String, sSName As String) As Boolean
    Dim sINIContent       As String
    Dim bResult           As Boolean
    Dim lPosition         As Long
    Dim fso               As New FileSystemObject
    Dim sSectionName      As String
    Dim nLoop             As Integer
    Dim sKeyName          As String
    
    bResult = True
    nLoop = 1
        Do
            sKeyName = "Code" & CStr(nLoop)
            sINIContent = GetIniS(sMMDIni, sSName, sKeyName, "")
            If Len(sINIContent) = 0 Then
                Exit Do
            End If
                
                lPosition = InStrRev(SPInfo, sINIContent)
                If lPosition > 0 Then
                    bResult = False
                    Exit Do
                End If
                nLoop = nLoop + 1
        Loop
    CheckSPInfo = bResult

End Function

