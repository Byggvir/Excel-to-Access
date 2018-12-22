Attribute VB_Name = "TAStringFunctions"
Public Function StrNS(Z As Variant) As String

  s = Str(Z)
  If Left(s, 1) = " " Then
    StrNS = Right(s, Len(s) - 1)
  Else
    StrNS = s
  End If

End Function


Public Function PadZero(Zahl As String, Optional Width As Integer = 2) As String

  Dim strTemp As String
  Dim i As Integer
  
  For i = 1 To Width
     strTemp = strTemp & "0"
  Next
  
  PadZero = Right(strTemp & Zahl, Width)
  
End Function
