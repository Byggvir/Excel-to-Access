VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InsertQueryForm 
   Caption         =   "Insert SQL query"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9690.001
   OleObjectBlob   =   "InsertQueryForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "InsertQueryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label4_Click()

End Sub

Private Sub RefreshSQLBtn_Click()

  Me.SQLTextBox.Value = ersetze(SQL_Insert, "\$1", ActiveSheet.Name)

End Sub

Private Sub MyDatabase_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  
  Dim lngCount As Long
 
  ' Open the file dialog
  
  With Application.FileDialog(msoFileDialogOpen)
    .AllowMultiSelect = False
    .Show
    If .SelectedItems.Count = 1 Then
      MyDatabase.Value = .SelectedItems(1)
    End If
  End With
  
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Initialize()

  Me.AtCell.Text = Selection.Address
  Me.SQLTextBox.Value = ersetze(SQL_Insert, "\$1", ActiveSheet.Name)
 
End Sub
Private Sub CancelButton1_Click()
  
  Me.Hide
    
End Sub


Private Sub OkButton1_Click()

 Dim DB As TAAccessDB
 
 Set DB = New TAAccessDB
 
 DB.Database = MyDatabase.Value
 
 DB.InsertQueryResult Me.SQLTextBox.Text, Me.AtCell.Text, Me.WithHeader.Value
  
 Set DB = Nothing
 
 Me.Hide
 
 
End Sub

Private Sub WithHeader_Click()

End Sub
