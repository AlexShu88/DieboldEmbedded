Attribute VB_Name = "Module1"
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
                (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) _
                As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
                (ByVal hKey As Long, ByVal lpValueName As String, _
                 ByVal lpReserved As Long, lpType As Long, lpData As Any, _
                 lpcbData As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Function RetrieveString(lpAppName As String, lpValueName As String, nSize As Long, lpFileName As String) As String
    Dim sValue As String
    sValue = String(nSize, " ")
    Dim nResult As Long
    nResult = GetPrivateProfileString(lpAppName, lpValueName, "", sValue, Len(sValue), lpFileName)
    RetrieveString = Left(sValue, nResult)
'    If RetrieveString = "" Then
'        LogError "[" & lpAppName & "] " & lpValueName & " not found"
'    End If
End Function

Public Sub StrToBin(ByVal inString As String, ByRef bOutArray() As Byte)
    Dim strTwo As String
    Dim i As Integer, j As Integer

    j = 0
    For i = 1 To 16 Step 2
        strTwo = Mid(inString, i, 2)
        bOutArray(j) = Val("&H" + strTwo)
        j = j + 1
    Next

End Sub

Public Sub BinToStr(ByRef InPar() As Byte, ByRef OutPar As String)
    Dim i As Integer
    Dim strNum As String
    
'    For i = 1 To 8
'        strBuf1(i) = MidB(StrConv(InPar, vbFromUnicode), i, 1)
'        bBuffer(i) = AscB(strBuf1(i))
'    Next

    For i = 0 To 7
'        strNum = Hex(bBuffer(i))
        strNum = Hex(InPar(i))
        If Len(strNum) < 2 Then
            strNum = "0" + strNum
        End If
        OutPar = OutPar + strNum
    Next i
End Sub


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
'        nSize = 20
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
'        RegCloseKey hKey
    Else
        GetRegKeyN = DefValue
    End If
    RegCloseKey hKey
End Function

Function SetRegKeyValue(ByVal RootKeyName As Long, ByVal SubRegKeyPath As String, ByVal KeyValue As String, ByVal vType As Long, Value As Variant, Optional ByVal lenValue As Long) As Boolean
    Dim bRet As Boolean, bArr() As Byte
    Dim nReply As Integer
    
    bRet = False
    
    On Error GoTo ErrorExit
    
    nReply = RegOpenKey(RootKeyName, SubRegKeyPath, hKey)
    If ERROR_SUCCESS = nReply Then
        Select Case vType
            Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                If Value = "" Then
                    nReply = RegSetValueEx(hKey, KeyValue, 0&, vType, 0, 1)
                Else
                    nReply = RegSetValueEx(hKey, KeyValue, 0&, vType, ByVal CStr(Value), LenB(StrConv(Value, vbFromUnicode)) + 1)
                End If
            Case REG_DWORD, REG_DWORD_BIG_ENDIAN
                nReply = RegSetValueEx(hKey, KeyValue, 0&, vType, CLng(Value), 4)
            Case REG_BINARY
                Dim i As Integer
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

Function GetUnNumberMarkPos(ByVal Start As Integer, ByVal SearchString As String, ByVal Times As Integer, ByVal MaxLength As Integer) As Integer
    Dim Counter As Integer
    Dim sEachByte As String
    Dim FindSeprator As Boolean
    
    FindSeprator = False
    Counter = Start
    Do While Counter < MaxLength
        sEachByte = Mid(SearchString, Counter, 1)
        If IsNumeric(sEachByte) = False Then
            Times = Times - 1
            FindSeprator = True
            If Times = 0 Then
                Exit Do
            End If
        End If
      Counter = Counter + 1
    Loop
    If FindSeprator = True Then
        GetUnNumberMarkPos = Counter
    Else
        GetUnNumberMarkPos = 0
    End If
    
End Function
