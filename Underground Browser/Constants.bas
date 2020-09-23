Attribute VB_Name = "Constants"
Option Explicit
Global PopupKill As Boolean
Public Const PROGRAM_NAME = "Underground Search by Corrupted Inc."
Public Const NEW_TAB_HOME = 1
Public Const NEW_TAB_BLANK = 2
Public Const NEW_TAB_CUR_URL = 3
Public Const NEW_TAB = 4
Public Const NEW_TAB_SEARCH = 5
Public Const BLANK_URL = "about:<html><head><title>Blank</Title></head><body><center><h1>Underground Browser Beta 5 By Spaz<BR></h1><h>By <a href='HTTP://www.astemaninc.net'>Corrupted Inc.</a></h></body></html>"
Public Sub LoadlistBox(path As String, ListName As ListBox)
    
    
    Dim MyString As String, String1 As String
    On Error Resume Next
    Open path$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        
        DoEvents
        ListName.AddItem MyString$
        
    Wend
    Close #1
End Sub

Function ListIsIn(lst As ListBox, zString As String) As Boolean
    
    On Error Resume Next
Dim i As Long

    For i = 0 To lst.ListCount
        If lst.List(i) = zString Then ListIsIn = True: GoTo grr
    Next i
    ListIsIn = False
grr:
End Function

