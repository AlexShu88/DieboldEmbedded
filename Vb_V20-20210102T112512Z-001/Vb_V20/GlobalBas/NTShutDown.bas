Attribute VB_Name = "NTShutDown"
Option Explicit
'Private Const EWX_LogOff As Long = 0
'Private Const EWX_SHUTDOWN As Long = 1
'Private Const EWX_REBOOT As Long = 2
'Private Const EWX_FORCE As Long = 4
'Private Const EWX_POWEROFF As Long = 8

'ExitWindowsEx函数可以退出登录、关机或者重新启动系统
Private Declare Function ExitWindowsEx Lib "user32" _
(ByVal dwOptions As Long, _
ByVal dwReserved As Long) As Long

'GetLastError函数返回本线程的最后一次错误代码。错误代码是按照线程
'储存的，多线程也不会覆盖其他线程的错误代码。
Private Declare Function GetLastError Lib "kernel32" () As Long

'GetCurrentProcess函数返回当前进程的一个句柄。
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Type LUID
UsedPart As Long
IgnoredForNowHigh32BitPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
TheLuid As LUID
Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
PrivilegeCount As Long
TheLuid As LUID
Attributes As Long
End Type
    
'OpenProcessToken函数打开一个进程的访问代号。
Private Declare Function OpenProcessToken Lib "advapi32" _
(ByVal ProcessHandle As Long, _
ByVal DesiredAccess As Long, _
TokenHandle As Long) As Long

'LookupPrivilegeValue函数获得本地唯一的标示符(LUID)，用于在特定的系统中
'表示特定的优先权。
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
Alias "LookupPrivilegeValueA" _
(ByVal lpSystemName As String, _
ByVal lpName As String, _
lpLuid As LUID) As Long

'AdjustTokenPrivileges函数使能或者禁用指定访问记号的优先权。
'使能或者禁用优先权需要TOKEN_ADJUST_PRIVILEGES访问权限。
Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
(ByVal TokenHandle As Long, _
ByVal DisableAllPrivileges As Long, _
NewState As TOKEN_PRIVILEGES, _
ByVal BufferLength As Long, _
PreviousState As TOKEN_PRIVILEGES, _
ReturnLength As Long) As Long

Private Declare Sub SetLastError Lib "kernel32" _
(ByVal dwErrCode As Long)

'********************************************************************
'* 这个过程设置正确的优先权，以允许在Windows NT下关机或者重新启动。
'********************************************************************
Private Function AdjustToken() As Long
    Const TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8
    Const SE_PRIVILEGE_ENABLED = &H2
    
    Dim hdlProcessHandle As Long
    Dim hdlTokenHandle As Long
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    
    '使用SetLastError函数设置错误代码为0。
    '这样做，GetLastError函数如果没有错误会返回0
    SetLastError 0
    
    ' GetCurrentProcess函数设置 hdlProcessHandle变量
    hdlProcessHandle = GetCurrentProcess()
    
    If GetLastError <> 0 Then
        AdjustToken = GetLastError
        Exit Function
    End If
    
    OpenProcessToken hdlProcessHandle, _
    (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle
    
    If GetLastError <> 0 Then
        AdjustToken = GetLastError
        Exit Function
    End If
    
    ' 获得关机优先权的LUID
    LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
    
    If GetLastError <> 0 Then
        AdjustToken = GetLastError
        Exit Function
    End If
    
    tkp.PrivilegeCount = 1 ' 设置一个优先权
    tkp.TheLuid = tmpLuid
    tkp.Attributes = SE_PRIVILEGE_ENABLED
    
    ' 对当前进程使能关机优先权
    AdjustTokenPrivileges hdlTokenHandle, _
    False, _
    tkp, _
    Len(tkpNewButIgnored), _
    tkpNewButIgnored, _
    lBufferNeeded
    
    If GetLastError <> 0 Then
        AdjustToken = GetLastError
    Else
        AdjustToken = 0
    End If
End Function

Function NTSystemShutDown(ByVal lShutDownType As Long) As Long
Dim nRet As Long
    nRet = AdjustToken
    If nRet <> 0 Then
        NTSystemShutDown = nRet
        Exit Function
    End If
    
    nRet = ExitWindowsEx(lShutDownType, &HFFFF)
    If nRet = 0 Then
        NTSystemShutDown = GetLastError
    Else
        NTSystemShutDown = 0
    End If
End Function
