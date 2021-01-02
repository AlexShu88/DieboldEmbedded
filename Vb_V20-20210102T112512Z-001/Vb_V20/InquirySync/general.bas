Attribute VB_Name = "Module1"
Option Explicit

Sub DrawInquiryPrr(ByVal Pcb3dl As DL)
    Dim nReply As Byte
    
    nReply = Pcb3dl.DlSetCharRaw("PrrOthersMark", "INQ")
    nReply = Pcb3dl.DlSetCharRaw("PrrRejectMark", "***")
    nReply = Pcb3dl.DlSetCharRaw("PrrRejectedCode", Pcb3dl.DlGetCharRaw("ATMPRejectCode"))
    
End Sub
