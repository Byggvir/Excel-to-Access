Attribute VB_Name = "TARegExp"
' Makros, die ausf regulären Ausdrücke basieren und diese erweitern
' Macros based on an expanding regulare expressions


' (cc) 2018 by Thomas Arend
'

' Funktion ersetze um reguläre Ausdrücke in Formeln einzusetzen.

Public Function ersetze(ByVal Source As String, ByVal Muster As String, ByVal Ersatz As String) As String

  Dim regex As New RegExp
 
  With regex
    .Global = True
    .Pattern = Muster
  End With

  Set Fundstellen = regex.Execute(Source)

  ersetze = regex.Replace(Source, Ersatz)

End Function
