Attribute VB_Name = "Module1"
Public Function UnPack(ByRef arrBytes() As Byte) As String
    Dim lIndex As Long
    Dim sHexValue As String
    Dim lUBound As Long
    Dim sRet As String
    
    lUBound = UBound(arrBytes)
    sRet = ""
    'Unpack
    For lIndex = LBound(arrBytes) To lUBound
        sHexValue = UCase(Hex(arrBytes(lIndex)))
        If (Len(sHexValue) = 1) Then
            sHexValue = "0" & sHexValue
        End If
        sRet = sRet & sHexValue
    Next lIndex
    
    UnPack = sRet
End Function


