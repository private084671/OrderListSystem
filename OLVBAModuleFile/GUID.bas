Attribute VB_Name = "GUID"
Option Explicit

Public Function Create() As String
    
    'GUID = Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)
    
    Dim GUID As String
    
    GUID = Right$("0000" & Hex(Rnd() * 65536), 4) & Right$("0000" & Hex(Rnd() * 65536), 4)
    
    GUID = GUID & "-" & Right$("0000" & Hex(Rnd() * 65536), 4)
    
    GUID = GUID & "-4" & Right$("0000" & Hex(Rnd() * 65536), 3)
    
    GUID = GUID & "-" & Right$("0000" & Hex(32768 + Rnd() * 16384), 4)
    
    Create = GUID & "-" & Right$("0000" & Hex(Rnd() * 65536), 4) & Right$("0000" & Hex(Rnd() * 65536), 4) & Right$("0000" & Hex(Rnd() * 65536), 4)

End Function

Private Sub Test_CreateDirectory()

    Dim TargetPath As String
    
    TargetPath = CreateObject("Wscript.Shell").SpecialFolders("Desktop") & "\" & Create

    Call MkDir(TargetPath)

End Sub
