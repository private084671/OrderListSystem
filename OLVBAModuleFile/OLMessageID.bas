Attribute VB_Name = "OLMessageID"
Option Explicit

Public Function GetID(ByRef TargetItem As MailItem) As String

    Const SchemaName As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"

    GetID = TargetItem.PropertyAccessor.GetProperty(SchemaName)
    
End Function

Private Sub Test_PrintID()

    Dim Items() As MailItem
    
    Items = SelectItems.GetSelectMailItems

    Dim i  As Long
    
    For i = 0 To UBound(Items)
    
        Debug.Print GetID(Items(i))
    
    Next i
    
End Sub
