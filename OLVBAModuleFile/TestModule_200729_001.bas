Attribute VB_Name = "TestModule_200729_001"
Option Explicit

Private Sub GetMailItemEntryID_001()

    Dim clsEntryID As OLEntryID
    
    Set clsEntryID = New OLEntryID
    
    Dim Items() As MailItem
    
    Items = clsEntryID.GetMailItems("")
    
    Stop

End Sub

Private Sub GetMailItemEntryID_002()

    Dim clsEntryID As OLEntryID
    
    Set clsEntryID = New OLEntryID
    
    Dim Items() As MailItem
    
    Items = clsEntryID.GetMailItems("aaa")
    
    Dim i As Long
    
    Dim ErrMessages() As String
    
    ErrMessages = clsEntryID.ErrMessages
    
    For i = 0 To UBound(ErrMessages)
    
        Debug.Print ErrMessages(i)
    
    Next i
    
    Stop

End Sub

Private Sub ErrObject_001()

    Dim TestErr As ErrObject
    
    Set TestErr = New ErrObject
    
    Stop

End Sub
