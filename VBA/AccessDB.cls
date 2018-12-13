VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public strFarbe As String
Private strDBProvider As String
Private strDBFile As String
Private Sub Class_Initialize()

  strDBProvider = "Microsoft.ACE.OLEDB.12.0"
  strDBFile = ""
  
End Sub

Public Property Let Database(DBFile As String)
  
  strDBFile = DBFile

End Property

Public Property Get Database() As String
    Database = strDatabase
End Property

Public Function GetLong(sSQL As String) As Long

' sSQL. Abfrage die ausgeführt werden soll

    ' Definieren der Objekte
    Dim cn As Object
    Dim rs As Recordset
    Dim sConnection As String
    
    ' Verbinden zur Datenbank
    Set cn = CreateObject("ADODB.Connection")
    sConnection = "Provider=" & strDBProvider & "; Data Source=" & strDBFile & ";"
    
    ' Abfrage ausführen
    On Error GoTo ErrorSQL
      cn.Open sConnection
      Set rs = cn.Execute(sSQL)
    On Error GoTo 0
    
    ' Ergebnis zurückgeben
    If Not rs.EOF Then
      On Error GoTo ErrorNoNumber
        sql_Long = rs.Fields(0)
      On Error GoTo 0
    Else
      sql_Long = 0
    End If
        
    ' Bereinigen der Objekte
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
  
    Exit Function
    
ErrorNoNumber:
  On Error GoTo 0
  GetLong = 0
  
  ' Bereinigen der Objekte
  rs.Close
  Set rs = Nothing
  cn.Close
  Set cn = Nothing
  Exit Function

ErrorSQL:
  On Error GoTo 0
  GetLong = -1
  
  ' Bereinigen der Objekte
  Set rs = Nothing
  cn.Close
  Set cn = Nothing
  Exit Function

End Function

Public Function GetString(sSQL As String) As String

' sSQL. Abfrage die ausgeführt werden soll

    ' Definieren der Objekte
  Dim cn As Object
  Dim rs As Object
  Dim sConnection As String
    
  ' Verbinden zur Datenbank
  Set cn = CreateObject("ADODB.Connection")
  sConnection = "Provider=" & strDBProvider & "; Data Source=" & strDBFile & ";"
  cn.Open sConnection
    
  ' Abfrage ist im Übergabeparameter definiert
       
  ' Abfrage ausführen

  On Error GoTo ErrorSQL
    Set rs = cn.Execute(sSQL)
  On Error GoTo 0
 
    ' Ergebnis zurückgeben
  If Not rs.EOF Then
    On Error GoTo ErrorNoString
      GetString = rs.Fields(0)
    On Error GoTo 0
  Else
    GetString = "--EOF--"
  End If
        
  ' Bereinigen der Objekte
  rs.Close
  Set rs = Nothing
  cn.Close
  Set cn = Nothing
  Exit Function
       
       
ErrorSQL:
  On Error GoTo 0
  GetString = "--Err: SQL--"
  rs.Close
  Set rs = Nothing
  cn.Close
  Set cn = Nothing
  Exit Function
       
ErrorNoString:
  On Error GoTo 0
  GetString = "--Err: No String--"
  rs.Close
  Set rs = Nothing
  cn.Close
  Set cn = Nothing
  Exit Function
    
End Function
Public Function GetList(sSQL As String) As String

' sSQL. Abfrage die ausgeführt werden soll

  ' Definieren der Objekte
  Dim cn As Object
  Dim rs As Object
  Dim sConnection As String
    
  Dim tmpList As String
  Dim Trenner As String
    
  ' Verbinden zur Datenbank
  Set cn = CreateObject("ADODB.Connection")
  sConnection = "Provider=" & strDBProvider & "; Data Source=" & strDBFile & ";"
  cn.Open sConnection
    
  ' Abfrage ist im Übergabeparameter definiert

  ' Abfrage ausführen

  Set rs = cn.Execute(sSQL)
   
  ' Ergebnis zurückgeben
    
  tmpList = ""
  Trenner = ""

  While Not rs.EOF
      
    On Error GoTo ErrorNoString
      tmpList = tmpList & Trenner & rs.Fields(0)
    On Error GoTo 0
      
    Trenner = vbLf
    rs.MoveNext
      
  Wend
    
  GetList = tmpList
    
  ' Bereinigen der Objekte
  rs.Close
  Set rs = Nothing
  cn.Close
  Set cn = Nothing
  Exit Function
       
ErrorNoString:
  On Error GoTo 0
  GetList = "--Err: No String--"
  rs.Close
  Set rs = Nothing
  cn.Close
  Set cn = Nothing
  Exit Function
    
End Function

Public Sub InsertQueryResult(sSQL As String, Optional AtCell As String = "?")


' sSQL: Abfrage die ausgeführt werden soll.
' AtCell: Position (Cell) ab der das Abfrageergebnis kopiert werden soll.
'         Default: Aktive Zelle

' Definieren der Objekte

  Dim cn As Object
  Dim rs As Object
  Dim sConnection As String
  Dim Target As Range
  
  If AtCell = "?" Then
    Set Target = Selection
  Else
    Set Target = Range(AtCell)
  End If
    
  ' Verbinden zur Datenbank
  Set cn = CreateObject("ADODB.Connection")
  sConnection = "Provider=" & strDBProvider & "; Data Source=" & strDBFile & ";"
     
' Abfrage ausführen
  cn.Open sConnection
  Set rs = cn.Execute(sSQL)
' Copy the recordset to Excel
  Target.CopyFromRecordset rs

   
End Sub
