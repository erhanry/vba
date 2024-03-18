Attribute VB_Name = "BALANS"
Function REMAINDER(A1 As Integer, T1 As Integer, A2 As Integer, T2 As Integer)

Dim X As Integer
Dim Z As Integer

X = A1 - T1
Z = A2 - T2

If A1 = T1 Or A2 = T2 Or (X > 0 And Z > 0) Or (X < 0 And Z < 0) Then
    REMAINDER = X
ElseIf (X < 0 And X + Z > 0) Or (X > 0 And X + Z < 0) Then
    REMAINDER = 0
Else
    REMAINDER = X + Z
End If

End Function


Function TRANSFER(A1 As Integer, T1 As Integer, A2 As Integer, T2 As Integer)

Dim X As Integer
Dim Z As Integer

X = A1 - T1
Z = A2 - T2

If X >= 0 Or (X < 0 And Z < 0) Then
    TRANSFER = 0
ElseIf (X < 0 And X + Z > 0) Or (X > 0 And X + Z < 0) Then
    TRANSFER = -X
Else
    TRANSFER = Z
End If

End Function

Function BALANS(A As Integer, T As Integer, Optional CRITERIA As String = "A")

If CRITERIA = "A" Then
    BALANS = A - T
ElseIf CRITERIA = "P" And A - T > 0 Then
    BALANS = A - T
ElseIf CRITERIA = "O" And A - T < 0 Then
    BALANS = A - T
Else
    BALANS = 0
End If
End Function

















