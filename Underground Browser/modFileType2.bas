Attribute VB_Name = "modFileType"
Public Function MakeFileType(ByVal Extension As String, ByVal NameOfType As String, ByVal DefaultIcon As String, ByVal NameOfAction As String, ByVal ActionPath As String, Optional ByVal ShellNew As Boolean, Optional ByVal QuickView As Boolean, Optional logs As Boolean) As Boolean
    'On Error GoTo Oops
    Dim dotExtension As String, Extensionfile As String
    Dim correctNameOfAction As String
    Dim writes As String
    dotExtension = "." & Extension
    Extensionfile = Extension & "file"
    correctNameOfAction = ReplaceChars(NameOfAction, " ", "_")
    
    If logs = True Then
      writes = GetString1(HKEY_CLASSES_ROOT, Extensionfile & "\Shell\" & correctNameOfAction, "command", "")
      writeini "Reg", Extension & "exe", writes, App.path & "\reg.ini"
      writes = GetString1(HKEY_CLASSES_ROOT, Extensionfile, "DefaultIcon", "")
      writeini "Reg", Extension & "ico", writes, App.path & "\reg.ini"
      writes = GetString1(HKEY_CLASSES_ROOT, Extensionfile, "shell", "")
      writeini "Reg", Extension & "act", writes, App.path & "\reg.ini"
    End If
    
    CreateKey HKEY_CLASSES_ROOT, dotExtension
    CreateKey HKEY_CLASSES_ROOT, Extensionfile
    CreateKey HKEY_CLASSES_ROOT, Extensionfile, "DefaultIcon"
    CreateKey HKEY_CLASSES_ROOT, Extensionfile, "Shell"
    CreateKey HKEY_CLASSES_ROOT, Extensionfile & "\Shell", correctNameOfAction
    CreateKey HKEY_CLASSES_ROOT, Extensionfile & "\Shell\" & correctNameOfAction, "command"
        
        
    SaveString1 HKEY_CLASSES_ROOT, dotExtension, "", "", Extensionfile
    SaveString1 HKEY_CLASSES_ROOT, Extensionfile, "", "", NameOfType
    SaveString1 HKEY_CLASSES_ROOT, Extensionfile, "DefaultIcon", "", DefaultIcon
    SaveString1 HKEY_CLASSES_ROOT, Extensionfile, "Shell", "", correctNameOfAction
    SaveString1 HKEY_CLASSES_ROOT, Extensionfile & "\Shell", correctNameOfAction, "", "&" & NameOfAction
    SaveString1 HKEY_CLASSES_ROOT, Extensionfile & "\Shell\" & correctNameOfAction, "command", "", ActionPath
    
    If Not IsMissing(ShellNew) Then
        EnableShellNew Extension, ShellNew
    End If
    
    If Not IsMissing(QuickView) Then
        EnableQuickView Extension, QuickView
    End If
    MakeFileType = True
    Exit Function
Oops:
    MakeFileType = False
    Exit Function
    Resume Next
End Function
   
Public Function EditDefaultIcon(ByVal Extension As String, ByVal NewIconPath As String) As Boolean
    On Error GoTo IconOops
    Dim Extensionfile As String
    Extensionfile = Extension & "file"
    
    SaveString1 HKEY_CLASSES_ROOT, Extensionfile, "DefaultIcon", "", NewIconPath
    EditDefaultIcon = True
    Exit Function

IconOops:
    EditDefaultIcon = False
    Exit Function
    Resume Next
End Function

Public Function EnableQuickView(ByVal Extension As String, ByVal QuickView As Boolean) As Boolean
    On Error GoTo QuickViewOops
    Dim Extensionfile As String
    Extensionfile = Extension & "file"
    
    If QuickView = True Then
        'enable QuickView
        CreateKey HKEY_CLASSES_ROOT, Extensionfile, "QuickView"
        SaveString1 HKEY_CLASSES_ROOT, Extensionfile, "QuickView", "", "*"
      Else
        'disable QuickView
        DeleteKey HKEY_CLASSES_ROOT, Extensionfile & "\QuickView"
    End If
    
    EnableQuickView = True
    Exit Function
    
QuickViewOops:
    EnableQuickView = False
    Exit Function
    Resume Next
    
End Function
Public Function AlwaysShowExt(ByVal Extension As String, ByVal ShowExt As Boolean) As Boolean
    'On Error GoTo ExtOops
    Dim Extensionfile As String
    Extensionfile = Extension & "file"
    
    If ShowExt = True Then
        'always show extension
        SaveString1 HKEY_CLASSES_ROOT, Extensionfile, "", "AlwaysShowExt", ""
      Else
        'don't show extension
        DeleteValue1 HKEY_CLASSES_ROOT, Extensionfile, "", "AlwaysShowExt"
    End If
    AlwaysShowExt = True
    Exit Function
    
QuickViewOops:
    AlwaysShowExt = False
    Exit Function
    Resume Next
End Function
Public Function SetAsDefaultAction(ByVal Extension As String, ByVal NameOfAction As String) As Boolean
    On Error GoTo DefOops
    Dim Extensionfile As String, correctNameOfAction As String
    Extensionfile = Extension & "file"
    correctNameOfAction = ReplaceChars(NameOfAction, " ", "_")
    
    SaveString1 HKEY_CLASSES_ROOT, Extensionfile, "Shell", "", correctNameOfAction
    
    SetAsDefaultAction = True
    Exit Function
    
DefOops:
    SetAsDefaultAction = False
    Exit Function
    Resume Next
End Function
Public Function ExistType(ByVal Extension As String) As Boolean
    On Error GoTo OopsExist
    Dim Extensionfile As String, dotExtension As String
    
    Extensionfile = GetString1(HKEY_CLASSES_ROOT, Extension & "file", "", "")
    If Extensionfile <> "" Then
        ExistType = True
      Else
        ExistType = False
    End If
    Exit Function
    
OopsExist:
    ExistType = False
    Exit Function
    Resume Next
End Function
Public Function EnableShellNew(ByVal Extension As String, ByVal ShellNew As Boolean) As Boolean
    On Error GoTo OopsShellN
    Dim dotExtension As String
    dotExtension = "." & Extension
    
    If ShellNew = True Then
        'enable
        CreateKey HKEY_CLASSES_ROOT, dotExtension, "ShellNew"
        SaveString1 HKEY_CLASSES_ROOT, dotExtension, "ShellNew", "NullFile", ""
      Else
        'disable
        DeleteKey HKEY_CLASSES_ROOT, dotExtension & "\ShellNew"
    End If
    EnableShellNew = True
    Exit Function
    
OopsShellN:
    EnableShellNew = False
    Exit Function
    Resume Next
End Function
Public Function ReplaceChars(ByVal Text As String, ByVal Char As String, ReplaceChar As String) As String
    Dim counter As Integer
    
    counter = 1
    Do
        counter = InStr(counter, Text, Char)
        If counter <> 0 Then
            Mid(Text, counter, Len(ReplaceChar)) = ReplaceChar
          Else
            ReplaceChars = Text
            Exit Do
        End If
    Loop

    ReplaceChars = Text
End Function

