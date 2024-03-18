Public Function HEX_HMACSHA256(ByVal sTextToHash As String, ByVal sSharedSecretKey As Variant)
    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA256")
 
    TextToHash = asc.Getbytes_4(sTextToHash)
    SharedSecretKey = asc.Getbytes_4(sSharedSecretKey)
    enc.Key = SharedSecretKey


    Dim Bytes() As Byte
    Bytes = enc.ComputeHash_2((TextToHash))
    HEX_HMACSHA256 = ConvToHexString(Bytes)
    Set asc = Nothing
    Set enc = Nothing
 
End Function
 
Private Function ConvToHexString(vIn As Variant) As Variant
    
    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing


End Function

Sub hahah()
 Debug.Print HEX_HMACSHA256(Range("a157"), Range("a158").Value)
End Sub

