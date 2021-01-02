Attribute VB_Name = "Module1"
Option Explicit
Public Const PrrOK As Byte = 1
Public Const PrrReject As Byte = 2
Sub DrawTransferPrr(ByVal Pcb3dl As DL, ByVal PrrType As Byte)
    Dim nReply              As Byte
    Dim TransPrrAmount      As String
    Dim TransPrrTfr2ndAccNo As String
    Dim FeeCharge           As String
    Dim HostSeqNo           As String
    
    TransPrrTfr2ndAccNo = Pcb3dl.DlGetCharRaw("Tfr2ndAccNo")
    nReply = Pcb3dl.DlSetCharRaw("PrrTfr2ndAccNo", TransPrrTfr2ndAccNo)
    
    nReply = Pcb3dl.DlSetCharRaw("PrrTransferMark", "***")
    
    TransPrrAmount = Pcb3dl.DlGetCharRaw("GBLPrtAmount")
    nReply = Pcb3dl.DlSetCharRaw("PrrTransAmount", "RMB " + TransPrrAmount)
    nReply = Pcb3dl.DlSetCharRaw("PrrAcceptMark", "***")
    nReply = Pcb3dl.DlSetCharRaw("PrrAcceptCode", "(0000)")
    nReply = Pcb3dl.DlSetCharRaw("PrrRejectedCode", Pcb3dl.DlGetCharRaw("ATMPRejectCode"))
    
    nReply = Pcb3dl.DlSetCharRaw("PrrTransType", "400000")
    FeeCharge = Pcb3dl.DlGetCharRaw("Icbccommicharge")
    If Not IsNumeric(FeeCharge) Then
        nReply = Pcb3dl.DlSetCharRaw("PrrFeeCharge", "")
    ElseIf CDbl(FeeCharge) = 0 Then
        nReply = Pcb3dl.DlSetCharRaw("PrrFeeCharge", "")
    Else
        nReply = Pcb3dl.DlSetCharRaw("PrrFeeCharge", "ÊÖÐø·Ñ:  " + FeeCharge)
    End If
    HostSeqNo = Pcb3dl.DlGetCharRaw("IcbcHostSeq")
    nReply = Pcb3dl.DlSetCharRaw("PrrHostEnqNo", "H-ENQ#:" + HostSeqNo)
End Sub
