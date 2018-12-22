Attribute VB_Name = "Testing"
Sub test()

 Dim DB As TAAccessDB
 
 Set DB = New TAAccessDB
 
 DB.Database = "C:\Users\arend\Documents\Datenbanken\Spielwiese.accdb"

 DB.InsertQueryResult ("select * from Personen")
 
 Set DB = Nothing
 
End Sub

Public Function ListOf(Source As String) As String
 
 Dim DB As TAAccessDB
 Set DB = New TAAccessDB
 DB.Database = "C:\Users\arend\Documents\Datenbanken\Spielwiese.accdb"
 
 ListOf = DB.GetList("select Nachname from Personen where Vorname= """ & Source & """;")
  
End Function


Public Function NachnameOf(Source As Long) As String
 
 Dim DB As TAAccessDB
 Set DB = New TAAccessDB
 DB.Database = "C:\Users\arend\Documents\Datenbanken\Spielwiese.accdb"
 
 NachnameOf = DB.GetList("select Nachname from Personen where pnr = " & Str(Source) & ";")
  
End Function


Public Sub GehaltsPlus()
 
 Dim DB As TAAccessDB
 Set DB = New TAAccessDB
 DB.Database = "C:\Users\arend\Documents\Datenbanken\Spielwiese.accdb"
 
 DB.Update ("update Personen set [grund_gehalt] = [grund_gehalt] + $3 where nachname = ""$1"" and vorname = ""$2"";")
  
End Sub

Sub testmsg()

  MsgBox ListOf("Thomas")
  MsgBox NachnameOf(10212624)
  
End Sub
