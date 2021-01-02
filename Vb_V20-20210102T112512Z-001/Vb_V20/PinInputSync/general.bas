Attribute VB_Name = "Module1"
Option Explicit

'Function FormTransExp(ByVal Pcb3dl As DL) As String
'
'    Dim AtmCode, TheTime, AccNo As String
'    Dim LineNum As String
'    Dim LocalRejCode As String
'    Dim JournalHead As String
'    Dim nLineNum As Integer
'    Dim nReply As Byte
'
'    TheTime = Format(Now(), "YY/MM/DD HH:MM:SS")
'    AtmCode = Pcb3dl.DlGetCharRaw("GBLAtmCode")
'
'    nLineNum = Pcb3dl.DlGetInt("MessNumber")
'    nReply = Pcb3dl.DlSetCharRaw("GBLLineSendNum", Format(nLineNum, "0000"))
'    LineNum = Pcb3dl.DlGetCharRaw("GBLLineSendNum")
'    AccNo = Pcb3dl.DlGetCharRaw("FitAccNo")
'    LocalRejCode = Pcb3dl.DlGetCharRaw("GBLATMLocRejCode")
'    JournalHead = "----------- Trans Exception -----------"
'
'    nLineNum = nLineNum + 1
'    If (nLineNum > 9999) Then
'        nLineNum = 0
'    End If
'    nReply = Pcb3dl.DlSetLong("MessNumber", nLineNum)
'
'    FormTransExp = JournalHead + vbCrLf + _
'                   TheTime + "  " + AtmCode + LineNum + vbCrLf + _
'                   AccNo + " PIN" + LocalRejCode + vbCrLf + _
'                   "---------------------------------------" + vbCrLf
'End Function



