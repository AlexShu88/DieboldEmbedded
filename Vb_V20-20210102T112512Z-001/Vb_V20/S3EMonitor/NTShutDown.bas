Attribute VB_Name = "NTShutDown"
Option Explicit
'Private Const EWX_LogOff As Long = 0
'Private Const EWX_SHUTDOWN As Long = 1
'Private Const EWX_REBOOT As Long = 2
'Private Const EWX_FORCE As Long = 4
'Private Const EWX_POWEROFF As Long = 8

'ExitWindowsEx���������˳���¼���ػ�������������ϵͳ
Private Declare Function ExitWindowsEx Lib "user32" _
(ByVal dwOptions As Long, _
ByVal dwReserved As Long) As Long

'GetLastError�������ر��̵߳����һ�δ�����롣��������ǰ����߳�
'����ģ����߳�Ҳ���Ḳ�������̵߳Ĵ�����롣
Private Declare Function GetLastError Lib "kernel32" () As Long

'GetCurrentProcess�������ص�ǰ���̵�һ�������
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
    
'OpenProcessToken������һ�����̵ķ��ʴ��š�
Private Declare Function OpenProcessToken Lib "advapi32" _
(ByVal ProcessHandle As Long, _
ByVal DesiredAccess As Long, _
TokenHandle As Long) As Long

'LookupPrivilegeValue������ñ���Ψһ�ı�ʾ��(LUID)���������ض���ϵͳ��
'��ʾ�ض�������Ȩ��
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
Alias "LookupPrivilegeValueA" _
(ByVal lpSystemName As String, _
ByVal lpName As String, _
lpLuid As LUID) As Long

'AdjustTokenPrivileges����ʹ�ܻ��߽���ָ�����ʼǺŵ�����Ȩ��
'ʹ�ܻ��߽�������Ȩ��ҪTOKEN_ADJUST_PRIVILEGES����Ȩ�ޡ�
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
'* �������������ȷ������Ȩ����������Windows NT�¹ػ���������������
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
    
    'ʹ��SetLastError�������ô������Ϊ0��
    '��������GetLastError�������û�д���᷵��0
    SetLastError 0
    
    ' GetCurrentProcess�������� hdlProcessHandle����
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
    
    ' ��ùػ�����Ȩ��LUID
    LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
    
    If GetLastError <> 0 Then
        AdjustToken = GetLastError
        Exit Function
    End If
    
    tkp.PrivilegeCount = 1 ' ����һ������Ȩ
    tkp.TheLuid = tmpLuid
    tkp.Attributes = SE_PRIVILEGE_ENABLED
    
    ' �Ե�ǰ����ʹ�ܹػ�����Ȩ
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
