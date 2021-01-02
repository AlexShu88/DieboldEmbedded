Attribute VB_Name = "Module1"
Option Explicit

Public Const PrrOK As Byte = 1
Public Const PrrReject As Byte = 2

Sub DrawPinChangePrr(ByVal Pcb3dl As DL, ByVal PrrType As Byte)
    Dim nReply As Integer

    nReply = Pcb3dl.DlSetCharRaw("PrrOthersMark", "PIN")
    If PrrType = PrrOK Then
        nReply = Pcb3dl.DlSetCharRaw("PrrAcceptMark", "***")
    Else
        nReply = Pcb3dl.DlSetCharRaw("PrrRejectMark", "***")
        nReply = Pcb3dl.DlSetCharRaw("PrrRejectedCode", Pcb3dl.DlGetCharRaw("ATMPRejectCode"))
    End If
    
End Sub

