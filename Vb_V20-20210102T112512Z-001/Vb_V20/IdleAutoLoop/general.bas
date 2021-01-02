Attribute VB_Name = "Module1"

'Function FormTransExp(ByRef PcB3Dl As PcB3Dl, _
'                      ByVal Captured As String, _
'                      ByVal ExpCode As String, _
'                      ByVal PUName As String, _
'                      ByVal explain As String) As String
'
'    Dim AccNo As String, AtmCode As String, Bank As String, TheTime As String
'    Dim sJournalNum As String, nJournalNum As Long
    
'    nJournalNum = PcB3Dl.DlGetInt("TotJournalNum")
'    sJournalNum = Format(nJournalNum, "00000")
'    nJournalNum = nJournalNum + 1
'    PcB3Dl.DlSetLong "TotJournalNum", nJournalNum
   
'    AccNo = PcB3Dl.DlGetCharRaw("FitAccNo")
'    AtmCode = PcB3Dl.DlGetCharRaw("GBLAtmCode")
'    Bank = PcB3Dl.DlGetCharRaw("GBLBankCode")
'    LineNum = PcB3Dl.DlGetCharRaw("GBLLineNum")
'    TheTime = Format(Now(), "YY/MM/DD HH:MM")
'
'    FormTransExp = "----------- Trans Exception ------------" + vbCrLf + _
'                   "Mod:" + Model + " PRJNo:" + sJournalNum + " " + TheTime + vbCrLf + _
'                   "Bank:" + Bank + " ATM:" + AtmCode + vbCrLf + _
'                   "AccNo:" + AccNo + " Sn:" + LineNum + vbCrLf + _
'                   "ExpCode:" + ExpCode + " Cpd:" + Captured + " P.U:" + PUName + vbCrLf + _
'                   "Explain:" + explain + vbCrLf + _
'                   "----------------------------------------" + vbCrLf
'End Function

