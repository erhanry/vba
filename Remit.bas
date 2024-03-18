Public Sub SHA256()

Dim Encoder As Object
Dim Encoder_SHA256 As Object
Dim TextToHash() As Byte
Dim bytes() As Byte
Dim oD As Object
Dim kod As String

Set kod = "erhan"
Set Encoder = CreateObject("System.Text.UTF8Encoding")
Set Encoder_SHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
TextToHash = Encoder.GetBytes_4(kod)
bytes = Encoder_SHA256.ComputeHash_2((TextToHash))
Set oD = CreateObject("MSXML2.DOMDocument")
      
With oD
  .LoadXML "<root />"
  .DocumentElement.DataType = "bin.Hex"
  .DocumentElement.nodeTypedValue = bytes
End With

ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
Debug.Print ConvToHexString
End Sub