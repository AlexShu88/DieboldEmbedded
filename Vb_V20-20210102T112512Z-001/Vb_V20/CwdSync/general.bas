Attribute VB_Name = "Module1"
Option Explicit
'Define the different withdraw receipt type for CMB Shenzhen
Public Const WthPrrOK As Byte = 1
Public Const WthPrrReject As Byte = 2
Public Const WthPrrCWC As Byte = 3
Public Const WthPrrTimeout As Byte = 4
Public Const WthPrrFloat As Byte = 5
Sub DrawWthPrr(ByVal Pcb3Dl As DL, ByVal WthPrrType As Byte)
    Dim nReply As Byte
    Dim TransCurrency As String
    Dim TransPrrAmount As String
    Dim LocRejCode As String
    Dim FeeCharge As String
    Dim HostSeqNo As String
    
    LocRejCode = Pcb3Dl.DlGetCharRaw("GBLATMLocRejCode")
    TransCurrency = Pcb3Dl.DlGetCharRaw("GBLCurrency_code")
    TransPrrAmount = Pcb3Dl.DlGetCharRaw("GBLPrtAmount")
    
    nReply = Pcb3Dl.DlSetCharRaw("PrrTransAmount", "RMB " + TransPrrAmount)
    nReply = Pcb3Dl.DlSetCharRaw("PrrWthMark", "***")
    nReply = Pcb3Dl.DlSetCharRaw("PrrAcceptCode", "(0000)")
    nReply = Pcb3Dl.DlSetCharRaw("PrrRejectedCode", "00")
    
    nReply = Pcb3Dl.DlSetCharRaw("PrrTransType", "010000")
    FeeCharge = Pcb3Dl.DlGetCharRaw("Icbccommicharge")
    If Not IsNumeric(FeeCharge) Then
        nReply = Pcb3Dl.DlSetCharRaw("PrrFeeCharge", "")
    ElseIf CDbl(FeeCharge) = 0 Then
        nReply = Pcb3Dl.DlSetCharRaw("PrrFeeCharge", "")
    Else
        nReply = Pcb3Dl.DlSetCharRaw("PrrFeeCharge", "手续费:  " + FeeCharge)
    End If
    HostSeqNo = Pcb3Dl.DlGetCharRaw("IcbcHostSeq")
    nReply = Pcb3Dl.DlSetCharRaw("PrrHostEnqNo", "H-ENQ#:" + HostSeqNo)
    
    Select Case WthPrrType
    
        Case WthPrrOK
            nReply = Pcb3Dl.DlSetCharRaw("PrrAcceptMark", "***")
         
        Case WthPrrReject    '2005.12.21增加拒绝收条打印
            nReply = Pcb3Dl.DlSetCharRaw("PrrRejectedCode", Pcb3Dl.DlGetCharRaw("ATMPRejectCode"))
            nReply = Pcb3Dl.DlSetCharRaw("PrrRejectMark", "***")
         
        Case WthPrrCWC
            nReply = Pcb3Dl.DlSetCharRaw("PrrContactBankMark", "***")
            nReply = Pcb3Dl.DlSetCharRaw("PrrOthersMark", "***")
            
        Case WthPrrTimeout, WthPrrFloat
            nReply = Pcb3Dl.DlSetCharRaw("PrrAcceptMark", "***")
            nReply = Pcb3Dl.DlSetCharRaw("PrrContactBankMark", "***")

    End Select

End Sub
Sub DrawCpdCardPrr(ByVal Pcb3Dl As DL)
    Dim nReply         As Byte
    Dim FullAccNo      As String
    
    FullAccNo = Pcb3Dl.DlGetCharRaw("FitAccNo")
    nReply = Pcb3Dl.DlSetCharRaw("FitPrrAccNo", FullAccNo)
    nReply = Pcb3Dl.DlSetCharRaw("PrrCardRetainMark", "***")
    nReply = Pcb3Dl.DlSetCharRaw("PrrContactBankMark", "***")
End Sub
