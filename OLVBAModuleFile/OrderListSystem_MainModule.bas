Attribute VB_Name = "OrderListSystem_MainModule"
Option Explicit

Public Sub SaveMails(ByVal EntryIDCollection As String)

    Dim clsEntryID As OLEntryID
    
    Set clsEntryID = New OLEntryID
    
    Dim Items() As MailItem
    
    Set Items = clsEntryID.GetMailItems(EntryIDCollection)
    
    'Items����̏ꍇ�̏������K�v
    
End Sub
