Attribute VB_Name = "TestModule_200807_001"
Option Explicit

Public Sub SaveMSGFileTest()

    Dim clsSelection As OLSelection
    
    Set clsSelection = New OLSelection
    
    Dim clsItemInFo As OrderListSystem_ItemInfo
    
    Set clsItemInFo = New OrderListSystem_ItemInfo
    
    Call clsSelection.CreateMailItems
    
    Dim Items() As MailItem
    
    Items = clsSelection.GetItems
    
    Set clsItemInFo.Item = Items(0)
    
    Call clsItemInFo.SaveMSGFile(CreateObject("WScript.Shell").SpecialFolders("Desktop"))
    
    Call MsgBox("•Û‘¶‚ªŠ®—¹‚µ‚Ü‚µ‚½")
    
End Sub
