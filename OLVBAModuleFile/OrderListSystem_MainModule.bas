Attribute VB_Name = "OrderListSystem_MainModule"
Option Explicit

Public Sub SaveMails(ByVal EntryIDCollection As String)

    Dim clsEntryID As OLEntryID
    
    Set clsEntryID = New OLEntryID
    
    Dim Items() As MailItem
    
    Items = clsEntryID.GetMailItems(EntryIDCollection)
    
    'Itemsが空の場合の処理が必要
    
End Sub
