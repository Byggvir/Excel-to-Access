VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UpdateQueryForm 
   Caption         =   "Update Access DB from Excel"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7500
   OleObjectBlob   =   "UpdateQueryForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UpdateQueryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton1_Click()
  
  Me.Hide
  
End Sub



Private Sub FromCell_Change()

  Me.SQLTextBox.Value = MkUpdateSQLFromCell(Me.FromCell.Text)
  
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

Private Sub OkButton1_Click()
 
 Dim DB As TAAccessDB
 
 Set DB = New TAAccessDB
 DB.Database = Me.MyDatabase.Value
 DB.Update (Me.SQLTextBox)
 
 Me.Hide

End Sub

Private Sub UserForm_Activate()
  
  Me.FromCell.Text = Selection.Address
  
  Me.SQLTextBox.Value = MkUpdateSQLFromCell()
   
End Sub

