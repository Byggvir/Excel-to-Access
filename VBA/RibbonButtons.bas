Attribute VB_Name = "RibbonButtons"
Public Sub InsertQuery(ByVal control As IRibbonControl)
  
  InsertQueryForm.Show
  
End Sub

Public Sub UpdateQuery(ByVal control As IRibbonControl)
  
  UpdateQueryForm.Show
  
End Sub
Public Sub TAE2ACopyrightBtnClick(ByVal control As IRibbonControl)

  MsgBox "(cc) 2018 by Thomas Arend; E-Mail: thomas@arend.xyz", vbOKOnly, "Copyright"
  
End Sub
