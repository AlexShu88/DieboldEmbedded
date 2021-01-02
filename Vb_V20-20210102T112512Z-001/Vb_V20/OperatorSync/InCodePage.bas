Attribute VB_Name = "basInCodePage"
Option Compare Binary
Option Explicit
Private Declare Function GetACP Lib "kernel32" () As Long

' The OS functions, if you prefer to use them
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long
'--------------------------------
'   UNICODE to ANSI conversion, via a given codepage
'--------------------------------
Public Function WToA(ByVal st As String, Optional ByVal cpg As Long = -1, Optional lFlags As Long = 0) As String
    Dim stBuffer As String
    Dim cwch As Long
    Dim pwz As Long
    Dim pwzBuffer As Long
    
    If cpg = -1 Then cpg = GetACP()
    pwz = StrPtr(st)
    cwch = WideCharToMultiByte(cpg, lFlags, pwz, -1, 0&, 0&, ByVal 0&, ByVal 0&)
    stBuffer = String$(cwch + 1, vbNullChar)
    pwzBuffer = StrPtr(stBuffer)
    cwch = WideCharToMultiByte(cpg, lFlags, pwz, -1, pwzBuffer, Len(stBuffer), ByVal 0&, ByVal 0&)
    WToA = Left$(stBuffer, cwch - 1)
End Function


