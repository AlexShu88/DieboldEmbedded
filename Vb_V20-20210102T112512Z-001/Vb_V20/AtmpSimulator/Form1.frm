VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   Caption         =   "ATMP"
   ClientHeight    =   7425
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRPTStatus 
      Height          =   375
      Left            =   7920
      TabIndex        =   20
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox txtJPTStatus 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   19
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CheckBox AutoIncRespCode 
      Caption         =   "Auto Inc Resp"
      Height          =   375
      Left            =   6420
      TabIndex        =   16
      Top             =   1440
      Width           =   1875
   End
   Begin VB.CheckBox AutoReply 
      Caption         =   "Auto Reply"
      Height          =   375
      Left            =   6840
      TabIndex        =   15
      Top             =   840
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox AtmpSnd 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   6120
      Width           =   9855
   End
   Begin VB.TextBox AtmpRcv 
      Height          =   3375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   2040
      Width           =   9855
   End
   Begin VB.CommandButton ATMPDisConnectButton 
      Caption         =   "DisConnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8400
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ComboBox ComboRespCode 
      Height          =   330
      ItemData        =   "Form1.frx":0000
      Left            =   4980
      List            =   "Form1.frx":0002
      TabIndex        =   11
      Text            =   "0000"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton SendResp 
      Caption         =   "Send Response"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8400
      TabIndex        =   9
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox TextPort 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Text            =   "12007"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox TextKey 
      Height          =   375
      Left            =   4980
      TabIndex        =   2
      Text            =   "12345678"
      Top             =   840
      Width           =   1785
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   9600
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6666
   End
   Begin VB.Line Line8 
      X1              =   7560
      X2              =   7920
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line7 
      X1              =   7560
      X2              =   7920
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line2 
      X1              =   4440
      X2              =   4800
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line1 
      X1              =   4440
      X2              =   4800
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Receipt"
      Height          =   375
      Left            =   6600
      TabIndex        =   18
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Journal"
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Top             =   5640
      Width           =   975
   End
   Begin VB.Line Line14 
      X1              =   4260
      X2              =   4980
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line13 
      X1              =   4260
      X2              =   4980
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line16 
      X1              =   4740
      X2              =   4980
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line15 
      X1              =   4740
      X2              =   4980
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line6 
      X1              =   1200
      X2              =   1800
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line5 
      X1              =   1200
      X2              =   1800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Response Code"
      Height          =   375
      Left            =   3300
      TabIndex        =   10
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Line Line4 
      X1              =   4260
      X2              =   4980
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line3 
      X1              =   4260
      X2              =   4980
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CLOSED"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4980
      TabIndex        =   8
      Top             =   240
      Width           =   1785
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "STATUS"
      Height          =   375
      Left            =   3300
      TabIndex        =   7
      Top             =   240
      Width           =   945
   End
   Begin VB.Image ncrlog 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   7200
      Picture         =   "Form1.frx":0004
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2205
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "RECEIVE ATMC TRANSACTION"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "ATMP Port"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Trm Key"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3300
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Current Sended Package"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Current Received Package"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RcvData As String
Dim RcvTime As String
Dim SndTime As String
Dim IsBeConnectionOpen As Boolean
Dim sTK As String
Dim sTKOld As String
Dim iRespCodeIndex As Integer
Const iTKUpdate As Long = 35827629

Private Sub ATMPDisConnectButton_Click()
If Winsock1.State <> sckClosed Then Winsock1.Close

IsBeConnectionOpen = False
Label15.Caption = "Closed"
Label15.BackColor = 255

'Disable
ATMPDisConnectButton.Enabled = False
SendResp.Enabled = False

txtRPTStatus.Text = ""
txtJPTStatus.Text = ""

txtRPTStatus.BackColor = vbWhite
txtJPTStatus.BackColor = vbWhite

Winsock1.Listen

End Sub

Private Sub AtmpRcv_Change()
If Len(AtmpRcv.Text) > 5000 Then
    AtmpRcv.Text = ""
End If
End Sub

Private Sub Form_Load()
Dim s As String
Dim sArray
Dim i As Integer

If App.PrevInstance = True Then
    End
End If

sTK = TextKey.Text
sTKOld = sTK

AtmpRcv.Text = ""
AtmpSnd.Text = ""


i = 0
Open "REJCODE.CFG" For Input As #2
Do While Not EOF(2)
  Line Input #2, s
  If Len(s) >= 3 Then
    If Mid(s, 1, 3) = "END" Or Mid(s, 1, 3) = "end" Then
      Exit Do
    ElseIf Mid(s, 1, 1) <> "#" And Mid(s, 1, 1) <> " " Then
      ComboRespCode.List(i) = s
      i = i + 1
    End If
  End If
Loop
Close #2

iRespCodeIndex = 0
ComboRespCode.Text = ComboRespCode.List(0)

On Error GoTo PARAMERR
Open "PARAM.CFG" For Input As #3
Do While Not EOF(3)
  Line Input #3, s
  If Len(s) > 3 Then
    sArray = Split(s, "=", -1, vbTextCompare)
    Select Case sArray(0)
      Case "TXKEY"
        TextKey.Text = sArray(1)
      Case "ATMPPORT"
        TextPort.Text = sArray(1)
      Case "AUTOREPLY"
        If sArray(1) = "1" Then
          AutoReply.Value = 1
        Else
          AutoReply.Value = 0
        End If
      Case "INCRESP"
        If sArray(1) = "1" Then
          AutoIncRespCode.Value = 1
        Else
          AutoIncRespCode.Value = 0
        End If
    End Select
  End If
Loop
Close #3

With Winsock1
    .Protocol = sckTCPProtocol
    .LocalPort = CLng(TextPort.Text)
    .Listen
End With
IsBeConnectionOpen = False


If False Then
PARAMERR: MsgBox "PARAM.CFG ERROR"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim FileNumber As Integer

FileNumber = FreeFile
Open "PARAM.CFG" For Output As #FileNumber
Print #FileNumber, "TXKEY=" + TextKey.Text
Print #FileNumber, "ATMPPORT=" + TextPort.Text
If AutoReply.Value = 1 Then
  Print #FileNumber, "AUTOREPLY=1"
Else
  Print #FileNumber, "AUTOREPLY=0"
End If
If AutoIncRespCode.Value = 1 Then
  Print #FileNumber, "INCRESP=1"
Else
  Print #FileNumber, "INCRESP=0"
End If
Close #FileNumber
End Sub

Private Sub SendResp_Click()
Dim FileNumber As Integer

If AtmpSnd.Text <> "" Then
    AtmpSnd.Text = "C" + AtmpSnd.Text
    Winsock1.SendData (AtmpSnd.Text)

    SndTime = Format(Time, "hh:mm:ss")
    FileNumber = FreeFile
    Open "atmp.log" For Append As #FileNumber
    Print #FileNumber, (Mid(SndTime, 1, 8) + " Snd: " + AtmpSnd.Text)
    Print #FileNumber, ""
    Close #FileNumber

Else
    ATMPDisConnectButton_Click
End If

If AutoIncRespCode.Value <> 0 Then
  iRespCodeIndex = iRespCodeIndex + 1
  If iRespCodeIndex >= ComboRespCode.ListCount Then iRespCodeIndex = 0
  ComboRespCode.Text = ComboRespCode.List(iRespCodeIndex)
End If

End Sub

Private Sub Winsock1_Close()
If Winsock1.State <> sckClosed Then Winsock1.Close

IsBeConnectionOpen = False
Label15.Caption = "Closed"
Label15.BackColor = 255

'Disable
ATMPDisConnectButton.Enabled = False
SendResp.Enabled = False

'clear jptr and rptr status
txtRPTStatus.Text = ""
txtJPTStatus.Text = ""

txtRPTStatus.BackColor = vbWhite
txtJPTStatus.BackColor = vbWhite

Winsock1.Listen

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.Close

Winsock1.Accept requestID

IsBeConnectionOpen = True
Label15.Caption = "Connected"
Label15.BackColor = &HFF00&

ATMPDisConnectButton.Enabled = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim FileNumber As Integer

Winsock1.GetData RcvData, vbString, 1500
Dim sJPT As String
Dim sRPT As String
Dim iCase As Integer

sJPT = Mid(RcvData, 20, 1)
sRPT = Mid(RcvData, 21, 1)

iCase = Val(sJPT)
Select Case iCase
    Case 0:
    txtJPTStatus.Text = "0-Good"
    txtJPTStatus.BackColor = &HFF00&
    Case 1:
    txtJPTStatus.Text = "1-Low"
    txtJPTStatus.BackColor = &HFF00FF
    Case 2:
    txtJPTStatus.Text = "2-Out"
    txtJPTStatus.BackColor = &HFF&
    Case 3:
    txtJPTStatus.Text = "3-inoportive"
    txtJPTStatus.BackColor = &HFFFF&

End Select

iCase = Val(sRPT)
Select Case iCase
    Case 0:
    txtRPTStatus.Text = "0-Good"
    txtRPTStatus.BackColor = &HFF00&
    Case 1:
    txtRPTStatus.Text = "1-Low"
    txtRPTStatus.BackColor = &HFF00FF
    Case 2:
    txtRPTStatus.Text = "2-Out"
    txtRPTStatus.BackColor = &HFF&
    Case 3:
    txtRPTStatus.Text = "3-inoportive"
    txtRPTStatus.BackColor = &HFFFF&

End Select
txtRPTStatus.Refresh
txtJPTStatus.Refresh

RcvTime = Format(Time, "hh:mm:ss")
AtmpRcv.Text = RcvTime + RcvData + vbCrLf + AtmpRcv.Text
'
'If AtmpRcv.Text <> "" Then
'AtmpRcv.Text = AtmpRcv.Text + (vbCrLf & RcvTime)
'Else
'AtmpRcv.Text = AtmpRcv.Text + RcvTime
'End If
'AtmpRcv.Text = AtmpRcv.Text + " " + RcvData
'AtmpSnd.Text = ""

ATMPDisConnectButton.Enabled = True
SendResp.Enabled = True

FileNumber = FreeFile
Open "atmp.log" For Append As #FileNumber
Print #FileNumber, (Mid(RcvTime, 1, 8) + " Rcv: " + RcvData)
Close #FileNumber

Call RspPkg
If AutoReply.Value = 1 Then
    Call SendResp_Click
End If


End Sub


Private Sub RspPkg()

Dim s As String
Dim StrArray, TmpArray
Dim i As Integer
Dim j As Integer
Dim FileNumber As Integer
Dim sTPC As String, sRejCode As String

AtmpSnd.Text = ""

sTPC = ""
sRejCode = "0000"
If IsNumeric(ComboRespCode.Text) Then
  If Val(ComboRespCode.Text) <> 0 Then
    sTPC = "ATP"
    sRejCode = ComboRespCode.Text
  End If
Else
  StrArray = Split(ComboRespCode.Text, ":", -1, vbTextCompare)
  If UBound(StrArray) - LBound(StrArray) > 0 Then
    sTPC = StrArray(1)
    sRejCode = StrArray(0)
  Else
    sTPC = StrArray(0)
  End If
End If

FileNumber = FreeFile
Open "atmp.cfg" For Input As #FileNumber
Do While Not EOF(FileNumber)
  Line Input #FileNumber, s
  If Len(s) > 0 And Mid(s, 1, 1) <> "#" Then
    StrArray = Split(s, "|", -1, vbTextCompare)
    If sTPC = StrArray(0) Or (Len(sTPC) = 0 And InStr(RcvData, StrArray(0)) > 0) Then
      i = 1
      Do While Mid(StrArray(i), 1, 1) <> "#"
        AtmpSnd.Text = AtmpSnd.Text + ConvertFormat(StrArray(i), RcvData, sRejCode)
        i = i + 1
      Loop
      Exit Do
    End If
  End If
Loop
Close #FileNumber
Me.Refresh
End Sub

Public Function ConvertFormat(ByVal sMsg As String, Optional ByVal sRcv As String, Optional ByVal sRej As String)
Dim i As Integer
Dim j As Integer
Dim MyDate As String
Dim MyTime As String
Dim TmpStr As String
Dim TmpArray

    If Mid(sMsg, 1, 3) = "YYY" Then
        If Mid(sMsg, 1, 8) = "YYYYMMDD" Then
            MyDate = Format(Date, "yyyymmdd")
        ElseIf Mid(sMsg, 1, 8) = "YYYYDDMM" Then
            MyDate = Format(Date, "yyyyddmm")
        End If
        ConvertFormat = MyDate

    ElseIf Mid(sMsg, 1, 3) = "YYM" Then
        MyDate = Format(Date, "yymmdd")
        ConvertFormat = MyDate
 
    ElseIf Mid(sMsg, 1, 3) = "YYD" Then
        MyDate = Format(Date, "yyddmm")
        If Mid(sMsg, 1, 6) = "YYDDMM" Then
            ConvertFormat = MyDate
        ElseIf Mid(sMsg, 1, 5) = "YYDDD" Then
            ConvertFormat = Mid(MyDate, 1, 2) + Right("000" + Format(Date, "y"), 3)
        End If

    ElseIf Mid(sMsg, 1, 3) = "HHM" Then
        MyTime = Format(Time, "hh:mm:ss")
        If Mid(sMsg, 1, 6) = "HHMMSS" Then
            ConvertFormat = Mid(MyTime, 1, 2) + Mid(MyTime, 4, 2) + Mid(MyTime, 7, 2)
        Else
            ConvertFormat = Mid(MyTime, 1, 2) + Mid(MyTime, 4, 2)
        End If

    ElseIf Mid(sMsg, 1, 3) = "FIL" Then
        TmpArray = Split(sMsg, ":", -1, vbTextCompare)
        TmpStr = ""
        For j = 1 To Val(TmpArray(2))
            TmpStr = TmpStr + TmpArray(1)
        Next
        ConvertFormat = TmpStr

    ElseIf Mid(sMsg, 1, 3) = "ORG" Then
        TmpArray = Split(sMsg, ":", -1, vbTextCompare)
        ConvertFormat = Mid(sRcv, Val(TmpArray(1)), Val(TmpArray(2)))
    
    ElseIf Mid(sMsg, 1, 3) = "REJ" Then
        ConvertFormat = sRej
        
    Else
        ConvertFormat = sMsg

    End If

End Function

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "WinSock Error: " + Description
Winsock1.Close
Winsock1.Listen

End Sub
