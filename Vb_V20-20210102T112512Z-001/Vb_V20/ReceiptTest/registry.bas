Attribute VB_Name = "registry"
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


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

Enum ErrorCode
    ERROR_SUCCESS = 0&
    ERROR_MORE_DATA = 234&
End Enum

Enum ValueType
    REG_NONE = 0
    REG_SZ = 1         'StringData
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_BIG_ENDIAN = 5
    REG_MULTI_SZ = 7   'MultiString
End Enum
'Sub SetIniS(ByVal AppProfileName As String, ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String)
'    Dim res%
'    res% = WritePrivateProfileString(SectionName, KeyWord, ValStr, AppProfileName)
'End Sub
Function SetIniS(ByVal AppProfileName As String, ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String) As Boolean
    Dim res As Integer
    res = WritePrivateProfileString(SectionName, KeyWord, ValStr, AppProfileName)
    If res <> 0 Then
        SetIniS = True
    Else
        SetIniS = False
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
Function GetIniD(ByVal AppProfileName As String, ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Double) As Double
    Dim tValue As String
    tValue = GetIniS(AppProfileName, SectionName, KeyWord, "")
    If IsNumeric(tValue) Then
        GetIniD = CDbl(tValue)
    Else
        GetIniD = DefValue
    End If
End Function

Function GetIniN(ByVal AppProfileName As String, ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Long) As Long
    GetIniN = GetPrivateProfileInt(SectionName, KeyWord, DefValue, AppProfileName)
End Function
Function GetValueLong(ByVal hKey As Long, ByVal ValueName As String, Value() As Long, vType As ValueType) As Boolean
    Dim ret As Long, length As Long, i As Integer
    GetValueLong = False
    length = 4
    ret = RegQueryValueEx(hKey, ValueName, 0&, REG_DWORD, Value(0), length)
    If ret = 0 Then GetValueLong = True
End Function


Function GetValue(ByVal hKey As Long, ByVal ValueName As String, Value() As Byte, vType As ValueType) As Boolean
    Dim ret As Long, length As Long, i As Integer

    ret = RegQueryValueEx(hKey, ValueName, 0&, REG_BINARY, 0&, length)
    If ret = 0 Or ret = ERROR_MORE_DATA Then
        ReDim Value(0 To length - 1)
        vType = REG_BINARY
        ret = RegQueryValueEx(hKey, ValueName, 0&, vType, Value(0), length)
        If ret = 0 Then GetValue = True
        If vType = REG_SZ Or vType = REG_EXPAND_SZ Or vType = REG_MULTI_SZ Then
            If UBound(Value) > 0 Then
                ReDim Preserve Value(0 To length - 2)
            End If
        End If
    End If
End Function
Function SetValue(ByVal hKey As Long, ByVal ValueName As String, ByVal vType As Long, Value As Variant, Optional ByVal lenValue As Integer) As Boolean
    Dim ret As Long, bArr() As Byte

    On Error GoTo ErrorExit
    Select Case vType
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
            If Value = "" Then
                ret = RegSetValueEx(hKey, ValueName, 0&, vType, 0, 1)
            Else
                ret = RegSetValueEx(hKey, ValueName, 0&, vType, ByVal CStr(Value), LenB(StrConv(Value, vbFromUnicode)) + 1)
            End If
        Case REG_DWORD, REG_DWORD_BIG_ENDIAN
            ret = RegSetValueEx(hKey, ValueName, 0&, vType, CLng(Value), 4)
        Case REG_BINARY
            Dim i As Integer
            ReDim bArr(0 To lenValue - 1)
            For i = 0 To lenValue - 1
                bArr(i) = Value(i)
            Next
            ret = RegSetValueEx(hKey, ValueName, 0&, vType, bArr(0), lenValue)
    End Select
    SetValue = (ret = 0)
ErrorExit:
End Function

Sub ByteArrayToString(bArray() As Byte, s As String)
    s = StrConv(bArray, vbUnicode)
End Sub

Function DeleteSubkeyTree(ByVal hKey As Long, ByVal Subkey As String) As Boolean
    Dim ret As Long, Index As Long, Name As String
    Dim hSubKey As Long

    ret = RegOpenKey(hKey, Subkey, hSubKey)
    If ret <> 0 Then
        DeleteSubkeyTree = False
        Exit Function
    End If
    ret = RegDeleteKey(hSubKey, "")
    If ret <> 0 Then
        While GetSubkeyByIndex(hSubKey, 0, Name) And _
              DeleteSubkeyTree(hSubKey, Name)
        Wend
        ret = RegDeleteKey(hSubKey, "")
    End If
    DeleteSubkeyTree = (ret = 0)
End Function

Function GetSubkeyByIndex(ByVal hKey As Long, ByVal Index As Long, KeyName As String) As Boolean
    Dim ret As Long, Name As String, length As Long

    Name = String(256, Chr(0))
    ret = RegEnumKey(hKey, Index, Name, 256)
    If ret = 0 Then
        KeyName = Left(Name, InStr(Name, Chr(0)) - 1)
        GetSubkeyByIndex = True
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
    Dim hKey As Long
    
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
