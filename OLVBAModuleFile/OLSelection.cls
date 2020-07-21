VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OLSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private MailItems() As MailItem
    
Public Property Get ItemSelected() As Boolean

    'Application～が長いので関数化した
    
    If Application.ActiveExplorer.Selection.Count > 0 Then ItemSelected = True

End Property

Public Property Get ArrayIsBrank() As Boolean

    'Initializeの段階でRedimしているので Trueなら
    'CreateItemsを実行してないか Selection内にMailItem型アイテムが存在しない
    
    If MailItems(0) Is Nothing Then ArrayIsBrank = True

End Property

Public Property Get GetItems() As MailItem()

    GetItems = MailItems

End Property

Public Sub CreateMailItems()

    If Not ItemSelected Then Exit Sub

    Dim SelectionItems As Selection
    
    Set SelectionItems = Application.ActiveExplorer.Selection

    ReDim MailItems(0)
    
    Dim i As Long
    
    For i = 1 To SelectionItems.Count
    
        If SelectionItems(i).Class = 43 Then Call AddMailItems(SelectionItems(i))
        
    Next i
        
End Sub

Private Sub AddMailItems(ByRef TargetItem As MailItem)

    If ArrayIsBrank Then
    
        Set MailItems(0) = TargetItem

    Else
    
        ReDim Preserve MailItems(UBound(MailItems) + 1)
        
        Set MailItems(UBound(MailItems)) = TargetItem
        
    End If
    
End Sub

Private Sub Class_Initialize()

    ReDim MailItems(0)
    
End Sub
