VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Connect 
      Caption         =   "Connect"
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton EjectCard 
      Caption         =   "EjectCard"
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Disconnect 
      Caption         =   "Disconnect"
      Height          =   735
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton InsertCard 
      Caption         =   "InsertCard"
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   720
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'typedef struct{
'    BYTE        startCode[4];         0x02 ATM
'    UINT32      cmdCode;            0x31 30 30 30
'BYTE        time[16];            0000000000000000
'    UINT32      chanIdx;                    31303030
'BYTE        cardno[24];              ASII后补零
'UINT32      cardnoOSD;                30323136
'BYTE        tradeType[4];               0000
'UINT32      tradeTypeOSD;             0000
'BYTE        amount[8];                  00000000
'UINT32      amountOSD;                 0000
'BYTE        serialNo[16];                000000000000000
'UINT32      serailNoOSD;                   0000
'BYTE        checksum[4];            前96个字节相加，累加和是2位，进位丢弃
'}CMD_DATA;


Private Sub Connect_Click()
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
End Sub

Private Sub Disconnect_Click()
  If MSComm1.PortOpen = True Then
   MSComm1.PortOpen = False
End If

End Sub

Private Sub EjectCard_Click()
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    StrTemp = Chr(2) + "ATM2"
    
    For i = 1 To 91
      StrTemp = StrTemp + Chr(0)
    Next
    
    LngTemp = 0
    For i = 1 To 96
         LngTemp = Asc(Mid(StrTemp, i, 1)) + LngTemp
    Next
    
    StrTemp = StrTemp + Right(Hex(LngTemp), 2) + Chr(0) + Chr(0)
    
    MSComm1.OutBufferCount = 0
    MSComm1.Output = StrTemp
    MSComm1.PortOpen = False

End Sub

Private Sub MSComm1_OnComm()
 
   Select Case MSComm1.CommEvent
   ' Handle each event or error by placing
   ' code below each case statement

   ' Errors
      Case comEventBreak   ' A Break was received.
      Case comEventFrame   ' Framing Error
      Case comEventOverrun   ' Data Lost.
      Case comEventRxOver   ' Receive buffer overflow.
      Case comEventRxParity   ' Parity Error.
      Case comEventTxFull   ' Transmit buffer full.
      Case comEventDCB   ' Unexpected error retrieving DCB]

   ' Events
      Case comEvCD   ' Change in the CD line.
      Case comEvCTS   ' Change in the CTS line.
      Case comEvDSR   ' Change in the DSR line.
      Case comEvRing   ' Change in the Ring Indicator.
      Case comEvReceive   ' Received RThreshold # of
                        ' chars.
      Case comEvSend   ' There are SThreshold number of
                     ' characters in the transmit
                     ' buffer.
      Case comEvEOF   ' An EOF charater was found in
                     ' the input stream
   End Select
End Sub
Private Sub InsertCard_Click()
   Dim StrTemp              As String
   Dim ByteArr(100)         As Byte
   Dim LngTemp              As Integer
   
   
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    StrTemp = Chr(2) + "ATM1"
    For i = 1 To 19
      StrTemp = StrTemp + Chr(0)
    Next
    
    StrTemp = StrTemp + "1000" + "1234567890123456789" + Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(0) + "0216"
    
    For i = 1 To 40
      StrTemp = StrTemp + Chr(0)
    Next
    
    LngTemp = 0
    For i = 1 To 96
         LngTemp = Asc(Mid(StrTemp, i, 1)) + LngTemp
    Next
    
    StrTemp = StrTemp + Right(Hex(LngTemp), 2) + Chr(0) + Chr(0)
    
    MSComm1.OutBufferCount = 0
    MSComm1.Output = StrTemp
    MSComm1.PortOpen = False

End Sub
