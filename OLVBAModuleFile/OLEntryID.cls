VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OLEntryID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private TargetSpace As NameSpace
    
    Private mErrMessages() As String
    Private MailItems() As MailItem

Public Property Get ErrMessages() As String()

    ErrMessages = mErrMessages

End Property

Public Function GetMailItems(ByVal EntryIDCollection As String) As MailItem()

    Dim IDList() As String
    
    IDList = Split(EntryIDCollection, ",")
    
    ReDim MailItem(0)
    
    Dim i As Long
    
    Dim TargetItem As Object
    
    For i = 0 To UBound(IDList)
    
        Set TargetItem = GetItem(IDList(i))
    
        If Not TargetItem Is Nothing Then Call AddItem(TargetItem)
        
    Next i
    
    GetMailItems = MailItems
    
End Function

Private Function GetItem(ByVal ID As String) As Object

    On Error GoTo TargetErr

    Set GetItem = TargetSpace.GetItemFromID(ID)

TargetErr:

    Dim ErrMessage As String

    ErrMessage = _
        Format(Now, "YYYY/MM/DD HH:NN") & vbTab & _
        Err.Number & vbTab & _
        Err.Description & vbTab & _
        "IDとして指定された文字列：" & ID & vbTab & _
        "エラーが発生したクラス：OLEntryID"
    
    Call AddErrMessage(ErrMessage)
    
End Function

Private Sub AddItem(ByRef TargetItem As Object)

    If TargetItem.Class <> 43 Then Exit Sub

    If MailItems(0) Is Nothing Then
    
        Set MailItems(0) = TargetItem
    
    Else
    
        ReDim Preserve MailItems(UBound(MailItems) + 1)
        
        Set MailItems(UBound(MailItems)) = TargetItem
    
    End If

End Sub

Private Sub AddErrMessage(ByVal TargetMessage As String)

    If mErrMessages(0) = "" Then
    
        mErrMessages(0) = TargetMessage
    
    Else
    
        ReDim Preserve mErrMessages(UBound(mErrMessages))
    
        mErrMessages(UBound(mErrMessages)) = TargetMessage
    
    End If
End Sub


Private Sub Class_Initialize()

    Set TargetSpace = Application.GetNamespace("MAPI")

    ReDim mErrMessages(0)
    
    ReDim MailItems(0)
    
End Sub
