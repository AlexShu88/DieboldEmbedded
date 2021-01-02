VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public Event LoadKeyOK()
Public Event CryptOK()
Public Event Failure(Description As String)

Public Enum KeyLoadMode
    EncryptFirst = 1
    DecryptFirst = 2
    NoAction = 3
End Enum

Private Type KeyLoadInfo
    KeyName As String
    KeyValue() As Byte
    sEncKeyname As String
    sCheckValue As String
    KeyType As XFSPINKeyType
    eKeyMode As KeyLoadMode
End Type
    
Private Type KeyInfo
    KeyName As String
    KeyType As XFSPINKeyType
End Type

Private Const OP_LoadKey As String = "LOADKEY"
Private Const OP_Encrypt As String = "ENCRYPT"
Private Const OP_Decrypt As String = "DECRYPT"
Private Const OP_MAC As String = "MAC"

Private Const KEY_USE_CRYPT As Byte = &H1
Private Const KEY_USE_FUNCTION As Byte = &H2
Private Const KEY_USE_MACING As Byte = &H4
Private Const KEY_USE_KEYENCKEY As Byte = &H20
Private Const KEY_USE_NO_DUPLICATE As Byte = &H40
Private Const KEY_USE_SVENCKEY As Byte = &H80

Private WithEvents moEncrypt As XFSPINPad
Attribute moEncrypt.VB_VarHelpID = -1

Private mKeyInfo() As KeyInfo
Private mlKeyCount As Long

Private msCurrentAction As String
Private mlRequestID As Long
Private mbLoadKey As Boolean
Private mbGenerateCheckValue As Boolean

Private mKeyLoadInfo() As KeyLoadInfo
Private mlKeyLoadCount As Long
Private mlKeyIndex As Long

'***********************End Of Variable Definition *******************************************************

Private Sub Class_Initialize()
On Error GoTo ErrorHandler
    Dim piAuditor As New VBAuditor
    piAuditor.Initialise "clsSecurity", "", "Initialise", AUD_EVENTARRIVED
    
    msCurrentAction = ""
    Erase mKeyInfo
    Erase mKeyLoadInfo
    mlKeyLoadCount = 0
    
    Dim vNames As Variant
    Set moEncrypt = New XFSPINPad
    vNames = moEncrypt.GetDeviceNames
    moEncrypt.SessionModel = XFS_PROCESS_BASED
    moEncrypt.SetDevice vNames(0)
    moEncrypt.TimeOut = 2000
    moEncrypt.ExceptionEventRequired = True
    
    mbLoadKey = False
    mbGenerateCheckValue = False
    Exit Sub
ErrorHandler:
    piAuditor.AuditHRESULT Err
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrorHandler
    Dim piAuditor As New VBAuditor
    piAuditor.Initialise "clsSecurity", "", "Terminate", AUD_EVENTARRIVED
    
    Set moEncrypt = Nothing
    EmptyKeyInfo
    EmptyKeyBuffer
    
    Exit Sub
ErrorHandler:
    piAuditor.AuditHRESULT Err
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub ConfigKeys()
On Error GoTo ErrorHandler
    Dim piAuditor As New VBAuditor
    piAuditor.Initialise "clsSecurity", "", "ConfigKeys", AUD_METHOD
    
    Dim lUBound As Long
    Dim lLBound As Long
    Dim vKeys As Variant
    Dim oKeys As XFSDictionary
    
    Set oKeys = moEncrypt.EncryptionKeys
    vKeys = oKeys.Keys
    If oKeys.Count = 0 Then
        Exit Sub
    End If
    
    lLBound = LBound(vKeys)
    lUBound = UBound(vKeys)
    Dim lLoop As Long
    
    mlKeyCount = oKeys.Count
    ReDim mKeyInfo(mlKeyCount - 1) As KeyInfo
    For lLoop = lLBound To lUBound
        mKeyInfo(lLoop - lLBound).KeyName = vKeys(lLoop)
        Set mKeyInfo(lLoop - lLBound).KeyType = GetKeyType(CByte(oKeys.Item(vKeys(lLoop))))
    Next lLoop
        
    Set oKeys = Nothing
    
    Exit Sub
ErrorHandler:
    piAuditor.AuditHRESULT Err
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Function AddKey(ByRef sKeyName As String, ByRef arrKeyValue() As Byte, ByRef sEncKey As String, ByRef sChkValue As String, eLoadMode As KeyLoadMode) As Boolean
On Error GoTo ErrorHandler
    Dim piAuditor As New VBAuditor
    piAuditor.Initialise "clsSecurity", "", "AddKey", AUD_METHOD, sKeyName, sEncKey

    AddKey = False
    
    Dim lIdx As Long
    If mlKeyLoadCount > 0 Then              ' Check if has already been added
        For lIdx = 0 To mlKeyLoadCount - 1
            With mKeyLoadInfo(lIdx)
                If .KeyName = sKeyName Then
                    Erase .KeyValue
                    .KeyValue = arrKeyValue
                    .sEncKeyname = sEncKey
                    .sCheckValue = sChkValue
                    If .sEncKeyname = "" Then
                        .eKeyMode = NoAction
                    Else
                        If eLoadMode = NoAction Then
                            .sEncKeyname = ""
                        End If
                        .eKeyMode = eLoadMode
                    End If
                    AddKey = True
                    Exit Function
                End If
            End With
        Next lIdx
    End If
    
    If mlKeyCount > 0 Then
        'Check if exists
        For lIdx = 0 To mlKeyCount - 1
            If mKeyInfo(lIdx).KeyName = sKeyName Then
                ReDim Preserve mKeyLoadInfo(mlKeyLoadCount) As KeyLoadInfo
                With mKeyLoadInfo(mlKeyLoadCount)
                    .KeyName = sKeyName
                    .KeyValue = arrKeyValue
                    .sCheckValue = sChkValue
                    Set .KeyType = mKeyInfo(lIdx).KeyType
                    .sEncKeyname = sEncKey
                    If .sEncKeyname = "" Then
                        .eKeyMode = NoAction
                    Else
                        If eLoadMode = NoAction Then
                            .sEncKeyname = ""
                        End If
                        .eKeyMode = eLoadMode
                    End If
                End With
                
                mlKeyLoadCount = mlKeyLoadCount + 1
                
                AddKey = True
                Exit Function
            End If
        Next lIdx
        
        piAuditor.Comment "Key (" & sKeyName & ") not configed"
        
    End If
    
    Exit Function
ErrorHandler:
    piAuditor.AuditHRESULT Err
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Sub EmptyKeyBuffer()
On Error GoTo ErrorHandler
    Dim piAuditor As New VBAuditor
    piAuditor.Initialise "clsSecurity", "", "EmptyKeyBuffer", AUD_METHOD

    If mlKeyLoadCount > 0 Then
        Dim lIdx As Long
        For lIdx = 0 To mlKeyLoadCount - 1
            Set mKeyLoadInfo(lIdx).KeyType = Nothing
        Next lIdx
    End If
    
    Erase mKeyLoadInfo
    mlKeyLoadCount = 0
    
    Exit Sub
ErrorHandler:
    piAuditor.AuditHRESULT Err
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub LoadKeys()
On Error GoTo ErrorHandler
    Dim piAuditor As New VBAuditor
    piAuditor.Initialise "clsSecurity", "", "LoadKeys", AUD_METHOD
    
    If mlKeyLoadCount = 0 Then
        piAuditor.Comment "No key need to load"
        RaiseEvent LoadKeyOK
    Else
        If moEncrypt.IsInitialised = True Then
            msCurrentAction = OP_LoadKey
            
            mbLoadKey = True
            mlKeyIndex = 0
            LoadKey mlKeyIndex
        Else
            RaiseEvent Failure("Encryptor not initialised")
        End If
    End If
    
    
    Exit Sub
ErrorHandler:
    piAuditor.AuditHRESULT Err
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub Encrypt(ByRef sKeyName As String, CryptData() As Byte, Algor As PINEncryptionAlgorithm)
On Error GoTo ErrorHandler
    Dim piAuditor As New VBAuditor
    piAuditor.Initialise "clsSecurity", "", "Encrypt", AUD_METHOD, sKeyName
    
    Dim EmptyKey() As Byte
    
    
    Dim lUBound As Long
    On Error Resume Next
    lUBound = UBound(mKeyInfo)
    If Err.Number = 9 Then
        RaiseEvent Failure("Key (" & sKeyName & ") not configed")
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim lIdx As Long
    Dim bFound As Boolean
    'Check if exists
    bFound = False
    For lIdx = 0 To lUBound
        If mKeyInfo(lIdx).KeyName = sKeyName Then
            bFound = True
            Exit For
        End If
    Next lIdx
    
    If bFound = False Then
        RaiseEvent Failure("Key (" & sKeyName & ") not configed")
        Exit Sub
    End If
    
    If (moEncrypt.EncryptionStatus = PIN_ENC_INITIALIZED) _
            Or (moEncrypt.EncryptionStatus = PIN_ENC_READY) Then
        msCurrentAction = OP_Encrypt
        mlRequestID = moEncrypt.CryptData(PIN_ENCRYPT, sKeyName, Algor, CryptData, 0, EmptyKey, EmptyKey, "", 0)
    Else
        RaiseEvent Failure("Encryptor not ready")
    End If


    Exit Sub
ErrorHandler:
    piAuditor.AuditHRESULT Err
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub MAC(ByRef sKeyName As String, MACData() As Byte)
On Error GoTo ErrorHandler
    Dim piAuditor As New VBAuditor
    piAuditor.Initialise "clsSecurity", "", "MAC", AUD_METHOD, sKeyName
    
    Dim EmptyKey() As Byte
    
    
    Dim lUBound As Long
    On Error Resume Next
    lUBound = UBound(mKeyInfo)
    If Err.Number = 9 Then
        RaiseEvent Failure("Key (" & sKeyName & ") not configed")
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim lIdx As Long
    Dim bFound As Boolean
    'Check if exists
    bFound = False
    For lIdx = 0 To lUBound
        If mKeyInfo(lIdx).KeyName = sKeyName Then
            bFound = True
            Exit For
        End If
    Next lIdx
    
    If bFound = False Then
        RaiseEvent Failure("Key (" & sKeyName & ") not configed")
        Exit Sub
    End If
    
    If (moEncrypt.EncryptionStatus = PIN_ENC_INITIALIZED) _
            Or (moEncrypt.EncryptionStatus = PIN_ENC_READY) Then
        msCurrentAction = OP_MAC
        mlRequestID = moEncrypt.CryptData(PIN_MAC, sKeyName, PIN_DES_MAC, MACData, 0, EmptyKey, EmptyKey, "", 0)
    Else
        RaiseEvent Failure("Encryptor not ready")
    End If


    Exit Sub
ErrorHandler:
    piAuditor.AuditHRESULT Err
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub


Public Sub Decrypt(ByRef sKeyName As String, CryptData() As Byte, Algor As PINEncryptionAlgorithm)
On Error GoTo ErrorHandler
    Dim piAuditor As New VBAuditor
    piAuditor.Initialise "clsSecurity", "", "Decrypt", AUD_METHOD, sKeyName
    
    
    Dim EmptyKey() As Byte
    
    
    Dim lUBound As Long
    On Error Resume Next
    lUBound = UBound(mKeyInfo)
    If Err.Number = 9 Then
        RaiseEvent Failure("Key (" & sKeyName & ") not configed")
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim lIdx As Long
    Dim bFound As Boolean
    'Check if exists
    bFound = False
    For lIdx = 0 To lUBound
        If mKeyInfo(lIdx).KeyName = sKeyName Then
            bFound = True
            Exit For
        End If
    Next lIdx
    
    If bFound = False Then
        RaiseEvent Failure("Key (" & sKeyName & ") not configed")
        Exit Sub
    End If
    
    If (moEncrypt.EncryptionStatus = PIN_ENC_INITIALIZED) _
            Or (moEncrypt.EncryptionStatus = PIN_ENC_READY) Then
        msCurrentAction = OP_Decrypt
        mlRequestID = moEncrypt.CryptData(PIN_DECRYPT, sKeyName, Algor, CryptData, 0, EmptyKey, EmptyKey, "", 0)
    Else
        RaiseEvent Failure("Encryptor not ready")
    End If


    Exit Sub
ErrorHandler:
    piAuditor.AuditHRESULT Err
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Function GetData() As Byte()
On Error GoTo ErrorHandler
    Dim piAuditor As New VBAuditor
    piAuditor.Initialise "clsSecurity", "", "GetData", AUD_METHOD
    
    GetData = moEncrypt.EncryptorData
    Exit Function
ErrorHandler:
    piAuditor.AuditHRESULT Err
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Sub moEncrypt_AvailabilityChanged(ByVal status As ACTIVEXFSCONTROLSLib.XFSStatus, ByVal workstationID As String, ByVal timeStamp As Date)
    If (moEncrypt.DeviceAvailability = XFS_AVAILABLE) Then
        ConfigKeys
    End If
    
End Sub

Private Sub moEncrypt_CryptFailed(ByVal requestID As Long, ByVal reason As ACTIVEXFSCONTROLSLib.PINCryptFailure, ByVal timeStamp As Date)
    Dim piAuditor As New VBAuditor
    piAuditor.Initialise "clsSecurity", "", "moEncrypt_CryptFailed", AUD_EVENTARRIVED, requestID
    
    Dim str As String
    Dim sReason As String

    Select Case reason
        Case PIN_CRYPT_ACCESS_DENIED
            sReason = "PIN_CRYPT_ACCESS_DENIED"
        Case PIN_CRYPT_CANCELED
            sReason = "PIN_CRYPT_CANCELED"
        Case PIN_CRYPT_COMMAND_ERROR
            sReason = "PIN_CRYPT_COMMAND_ERROR"
        Case PIN_CRYPT_DEVICE_LOCKED
            sReason = "PIN_CRYPT_DEVICE_LOCKED"
        Case PIN_CRYPT_DEVICE_NOT_READY
            sReason = "PIN_CRYPT_DEVICE_NOT_READY"
        Case PIN_CRYPT_HARDWARE_ERROR
            sReason = "PIN_CRYPT_HARDWARE_ERROR"
        Case PIN_CRYPT_INVALID_KEY_LENGTH
            sReason = "PIN_CRYPT_INVALID_KEY_LENGTH"
        Case PIN_CRYPT_KEY_NO_VALUE
            sReason = "PIN_CRYPT_KEY_NO_VALUE"
        Case PIN_CRYPT_KEY_NOT_FOUND
            sReason = "PIN_CRYPT_KEY_NOT_FOUND"
        Case PIN_CRYPT_MODE_NOT_SUPPORTED
            sReason = "PIN_CRYPT_MODE_NOT_SUPPORTED"
        Case PIN_CRYPT_TIME_OUT
            sReason = "PIN_CRYPT_TIME_OUT"
        Case PIN_CRYPT_UNSUPPORTED_COMMAND
            sReason = "PIN_CRYPT_UNSUPPORTED_COMMAND"
        Case PIN_CRYPT_USE_VIOLATION
            sReason = "PIN_CRYPT_USE_VIOLATION"
        Case Else
            sReason = CStr(reason)
    End Select
    str = Format(timeStamp, "YY/MM/DD hh:mm:ss") & " CryptFailed Raised." _
        & vbCrLf & vbTab & "RequestID=" & requestID & "(" & mlRequestID & ")" _
        & vbCrLf & vbTab & "Reason=" & sReason _
        & vbCrLf & vbTab & "Current Action=" & msCurrentAction
    
    piAuditor.Comment str
    
    RaiseEvent Failure(msCurrentAction & ":" & sReason)
End Sub

Private Sub moEncrypt_CryptOK(ByVal requestID As Long, ByVal Data As ACTIVEXFSCONTROLSLib.IXFSDictionary, ByVal timeStamp As Date)
    Dim piAuditor As New VBAuditor
    piAuditor.Initialise "clsSecurity", "", "moEncrypt_CryptOK", AUD_EVENTARRIVED, requestID
    
    piAuditor.Comment "Encryption result:" & UnPack(moEncrypt.EncryptorData)
    
    If mbLoadKey = True Then
        If mKeyLoadInfo(mlKeyIndex).eKeyMode = DecryptFirst Then
            ' Decrypt key first
            mKeyLoadInfo(mlKeyIndex).KeyValue = moEncrypt.EncryptorData
            mKeyLoadInfo(mlKeyIndex).sEncKeyname = ""
            mKeyLoadInfo(mlKeyIndex).eKeyMode = NoAction
            LoadKey mlKeyIndex
        Else
            If Mid(UnPack(moEncrypt.EncryptorData), 1, 4) = Mid(mKeyLoadInfo(mlKeyIndex).sCheckValue, 1, 4) Then
                mlKeyIndex = mlKeyIndex + 1
                If mlKeyIndex < mlKeyLoadCount Then
                    ' Continue load key
                    LoadKey mlKeyIndex
                Else
                    ' Key load finished
                    mbLoadKey = False
                    RaiseEvent LoadKeyOK
                End If
            Else
                ' Check value not equal
                piAuditor.Comment "Check value not Equal."
                RaiseEvent Failure("Check value not equal")
            End If
        End If
    Else
        RaiseEvent CryptOK
    End If
End Sub

Private Sub moEncrypt_ImportFailed(ByVal requestID As Long, ByVal reason As ACTIVEXFSCONTROLSLib.PINImportFailure, ByVal timeStamp As Date)
    Dim piAuditor As New VBAuditor
    piAuditor.Initialise "clsSecurity", "", "moEncrypt_ImportFailed", AUD_EVENTARRIVED, requestID
    
    Dim str As String
    Dim sReason As String
    
    mbLoadKey = False

    Select Case reason
        Case PIN_IMPORT_ACCESS_DENIED
            sReason = "PIN_IMPORT_ACCESS_DENIED"
        Case PIN_IMPORT_CANCELED
            sReason = "PIN_IMPORT_CANCELED"
        Case PIN_IMPORT_COMMAND_ERROR
            sReason = "PIN_IMPORT_COMMAND_ERROR"
        Case PIN_IMPORT_DEVICE_LOCKED
            sReason = "PIN_IMPORT_DEVICE_LOCKED"
        Case PIN_IMPORT_DEVICE_NOT_READY
            sReason = "PIN_IMPORT_DEVICE_NOT_READY"
        Case PIN_IMPORT_DUPLICATE_KEY
            sReason = "PIN_IMPORT_DUPLICATE_KEY"
        Case PIN_IMPORT_HARDWARE_ERROR
            sReason = "PIN_IMPORT_HARDWARE_ERROR"
        Case PIN_IMPORT_INVALID_ID
            sReason = "PIN_IMPORT_INVALID_ID"
        Case PIN_IMPORT_INVALID_KEY_LENGTH
            sReason = "PIN_IMPORT_INVALID_KEY_LENGTH"
        Case PIN_IMPORT_KEY_NO_VALUE
            sReason = "PIN_IMPORT_KEY_NO_VALUE"
        Case PIN_IMPORT_KEY_NOT_FOUND
            sReason = "PIN_IMPORT_KEY_NOT_FOUND"
        Case PIN_IMPORT_TIME_OUT
            sReason = "PIN_IMPORT_TIME_OUT"
        Case PIN_IMPORT_UNSUPPORTED_COMMAND
            sReason = "PIN_IMPORT_UNSUPPORTED_COMMAND"
        Case PIN_IMPORT_USE_VIOLATION
            sReason = "PIN_IMPORT_USE_VIOLATION"
        Case Else
            sReason = CStr(reason)
    End Select
    str = Format(timeStamp, "YY/MM/DD hh:mm:ss") & " ImportFailed Raised." _
        & vbCrLf & vbTab & "RequestID=" & requestID _
        & vbCrLf & vbTab & "Reason=" & sReason

    piAuditor.Comment str
    
    RaiseEvent Failure(msCurrentAction & ":" & sReason)
End Sub

Private Sub moEncrypt_ImportOK(ByVal requestID As Long, ByVal verificationData As ACTIVEXFSCONTROLSLib.IXFSDictionary, ByVal timeStamp As Date)
    Dim arrBytes(0 To 7) As Byte
    
    Erase arrBytes
    
    If mbGenerateCheckValue = True Then
        Encrypt mKeyLoadInfo(mlKeyIndex).KeyName, arrBytes, PIN_DES_CBC
    Else
        mlKeyIndex = mlKeyIndex + 1
        If mlKeyIndex < mlKeyLoadCount Then
            LoadKey mlKeyIndex
        Else
            mbLoadKey = False
            RaiseEvent LoadKeyOK
        End If
    End If
End Sub

Private Function GetKeyType(ByVal Data As Byte) As XFSPINKeyType
    Dim lKeyType As XFSPINKeyType
    
    Set lKeyType = New XFSPINKeyType
    
    
    If (Data And KEY_USE_CRYPT) <> 0 Then
        lKeyType.USE_CRYPT = True
    Else
        lKeyType.USE_CRYPT = False
    End If
    
    If (Data And KEY_USE_FUNCTION) <> 0 Then
        lKeyType.USE_FUNCTION = True
    Else
        lKeyType.USE_FUNCTION = False
    End If
    
    If (Data And KEY_USE_KEYENCKEY) <> 0 Then
        lKeyType.USE_KEYENCKEY = True
    Else
        lKeyType.USE_KEYENCKEY = False
    End If
    
    If (Data And KEY_USE_MACING) <> 0 Then
        lKeyType.USE_MACING = True
    Else
        lKeyType.USE_MACING = False
    End If
    
    If (Data And KEY_USE_NO_DUPLICATE) <> 0 Then
        lKeyType.USE_NO_DUPLICATE = True
    Else
        lKeyType.USE_NO_DUPLICATE = False
    End If
    
    If (Data And KEY_USE_SVENCKEY) <> 0 Then
        lKeyType.USE_SVENCKEY = True
    Else
        lKeyType.USE_SVENCKEY = False
    End If
    
    Set GetKeyType = lKeyType
End Function


Private Sub EmptyKeyInfo()
On Error GoTo ErrorHandler
    Dim piAuditor As New VBAuditor
    piAuditor.Initialise "clsSecurity", "", "EmptyKeyInfo", AUD_METHOD

    If mlKeyCount > 0 Then
        Dim lIdx As Long
        For lIdx = 0 To mlKeyCount - 1
            Set mKeyInfo(lIdx).KeyType = Nothing
        Next lIdx
    End If
    
    Erase mKeyInfo
    mlKeyCount = 0
    
    Exit Sub
ErrorHandler:
    piAuditor.AuditHRESULT Err
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub


Private Sub LoadKey(ByVal lIndex As Long)
Dim dummy() As Byte

If mKeyLoadInfo(lIndex).eKeyMode = DecryptFirst Then
    If (moEncrypt.EncryptionStatus = PIN_ENC_INITIALIZED) _
            Or (moEncrypt.EncryptionStatus = PIN_ENC_READY) Then
        mlRequestID = moEncrypt.CryptData(PIN_DECRYPT, _
                                        mKeyLoadInfo(lIndex).sEncKeyname, _
                                        PIN_DES_CBC, _
                                        mKeyLoadInfo(lIndex).KeyValue, _
                                        0, dummy, dummy, "", 0)
    Else
        RaiseEvent Failure("Encryptor not ready")
    End If
Else
    If mKeyLoadInfo(lIndex).sCheckValue = "" Then
        mbGenerateCheckValue = False
    Else
        mbGenerateCheckValue = True
    End If
    
    
    mlRequestID = moEncrypt.ImportKey(mKeyLoadInfo(lIndex).KeyName, _
                                    mKeyLoadInfo(lIndex).KeyValue, _
                                    mKeyLoadInfo(lIndex).KeyType, _
                                    mKeyLoadInfo(lIndex).sEncKeyname, dummy)
End If

End Sub
