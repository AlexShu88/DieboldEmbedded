Attribute VB_Name = "Module1"
Option Explicit
Function FormTransExp(ByVal AccNo As String, ByVal explain As String) As String
    Dim TheTime                     As String
   
    TheTime = Format(Now(), "MM/DD HH:MM")
    FormTransExp = ">>>" + vbCrLf + TheTime + "AccNo:" + AccNo + vbCrLf + _
                   "Explain:" + explain + vbCrLf
End Function
Function FormTransExpCHN(ByVal AccNo As String, ByVal explain As String) As String
    Dim TheTime                     As String
   
    TheTime = Format(Now(), "MM/DD HH:MM")
    FormTransExpCHN = ">>>" + vbCrLf + TheTime + "¿¨ºÅ£º" + AccNo + vbCrLf + _
                   "ËµÃ÷£º" + explain + vbCrLf
End Function

Function DeviceTransExp(ByVal ExplainWords As String) As String
    Dim TheTime                     As String
    
    TheTime = Format(Now(), "YY/MM/DD HH:MM:SS")
    DeviceTransExp = "***  " + TheTime + ExplainWords + vbCrLf
End Function

Sub DrawCpdCardPrr(ByVal Pcb3dl As DL)
    Dim nReply                      As Byte
    Dim FullAccNo                   As String
    Dim LocRejCode                  As String
    
    FullAccNo = Pcb3dl.DlGetCharRaw("FitAccNo")
    nReply = Pcb3dl.DlSetCharRaw("FitPrrAccNo", FullAccNo)
    nReply = Pcb3dl.DlSetCharRaw("PrrCardRetainMark", "***")
    nReply = Pcb3dl.DlSetCharRaw("PrrContactBankMark", "***")
End Sub
