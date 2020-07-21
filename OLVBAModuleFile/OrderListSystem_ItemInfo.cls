VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderListSystem_ItemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private mItem As MailItem
    Private mGUID As String

Public Property Set Item(ByRef TargetItem As MailItem)

    Set mItem = TargetItem

End Property

Public Property Get GetGUID() As String

    GetGUID = mGUID
 
End Property
Public Function GetRecord() As String

    Dim TimeStamp As String
    
    TimeStamp = Format(mItem.ReceivedTime, "YYYY/MM/DD HH:NN")

    Dim SenderName As String
    
    SenderName = mItem.SenderName

    Dim SenderEmailAddress As String
    
    SenderEmailAddress = mItem.SenderEmailAddress

    Dim Subject As String
    
    Subject = GetSubject
    
    Dim BodyString As String
    
    BodyString = GetBodyString
    
    Dim MessageID As String
    
    MessageID = GetMessageID

    GetRecord = Join(Array(TimeStamp, SenderName, SenderEmailAddress, Subject, BodyString, MessageID, mGUID), vbTab)
    
End Function

Private Function GetSubject() As String

    Dim Target As String
    
    Target = mItem.Subject
    
    If Len(Target) > 50 Then Target = Left(Target, 50) & "Åc"

    GetSubject = Target
    
End Function

Private Function GetBodyString() As String

    Dim Target As String
    
    Target = mItem.Body
    
    Target = Replace(Target, Chr(10), "")
    
    Target = Replace(Target, Chr(13), "")

    Target = Replace(Target, vbTab, "")
    
    If Len(Target) > 50 Then Target = Left(Target, 50) & "Åc"
    
    GetBodyString = Target
    
End Function

Private Function GetMessageID() As String

    Dim Target As String
    
    Target = OLMessageID.GetID(mItem)
    
    Target = Replace(Target, "<", "")
    
    Target = Replace(Target, ">", "")
    
    GetMessageID = Target
    
End Function

Private Sub Class_Initialize()

    mGUID = GUID.Create

End Sub
