VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InsertQueryForm 
   Caption         =   "Insert SQL query"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7410
   OleObjectBlob   =   "InsertQueryForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "InsertQueryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SQLTextBox_Change()

End Sub

Private Sub UserForm_Initialize()

  Me.AtCell.Text = Selection.Address
 
End Sub
Private Sub CancelButton1_Click()
  
  Me.Hide
    
End Sub

Private Sub Label2_Click()

End Sub

Private Sub OkButton1_Click()

 Dim DB As AccessDB
 
 Set DB = New AccessDB
 
 DB.Database = "C:\Users\arend\Documents\Datenbanken\Spielwiese.accdb"

 DB.InsertQueryResult Me.SQLTextBox.Text, Me.AtCell.Text
  
 Set DB = Nothing
 
 Me.Hide
 
 
End Sub
