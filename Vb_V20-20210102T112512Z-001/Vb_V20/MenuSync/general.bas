Attribute VB_Name = "Module1"
Option Explicit

'Function FormTransExp(ByVal Pcb3Dl As DL) As String
'    Dim AtmCode, TheTime, AccNo As String
'    Dim LineNum As String
'    Dim LocalRejCode As String
'    Dim JournalHead As String
'    Dim nLineNum As Integer
'    Dim nReply As Byte
'
'
'    TheTime = Format(Now(), "YY/MM/DD HH:MM:SS")
'    AtmCode = Pcb3Dl.DlGetCharRaw("GBLAtmCode")
'    nLineNum = Pcb3Dl.DlGetInt("MessNumber")
'    nReply = Pcb3Dl.DlSetCharRaw("GBLLineSendNum", Format(nLineNum, "0000"))
'
'    LineNum = Pcb3Dl.DlGetCharRaw("GBLLineSendNum")
'    AccNo = Pcb3Dl.DlGetCharRaw("FitAccNo")
'    LocalRejCode = Pcb3Dl.DlGetCharRaw("GBLATMLocRejCode")
'    JournalHead = "----------- Trans Exception -----------"
'
'    nLineNum = nLineNum + 1
'    If (nLineNum > 9999) Then
'        nLineNum = 0
'    End If
'    nReply = Pcb3Dl.DlSetLong("MessNumber", nLineNum)
'
'    FormTransExp = JournalHead + vbCrLf + _
'                   TheTime + "  " + AtmCode + LineNum + vbCrLf + _
'                   AccNo + " MEN" + LocalRejCode + vbCrLf + _
'                   "---------------------------------------" + vbCrLf
'End Function
'
'
