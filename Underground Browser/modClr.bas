Attribute VB_Name = "modClr"
Public Type colors
  bgClr As Long
  frClr As Long
  fntProp As Long
End Type

Public ClrData(19) As colors

Public Function one_argument(arg1 As String, arg2 As String) As String
    
    Dim X As Integer

    If Mid(arg1, 1, 1) = "'" Then
        X = InStr(2, arg1, "'")
        arg2 = Mid(arg1, 2, X - 2)
        arg1 = Mid(arg1, X + 1, Len(arg1) - X)
        one_argument = arg1
        Exit Function
    End If
    X = InStr(1, arg1, " ")

    If X = 0 Then
        arg2 = arg1
        one_argument = arg1
        Exit Function
    End If
    arg2 = Mid(arg1, 1, X)
    arg1 = Mid(arg1, X + 1, Len(arg1) - X)
    one_argument = arg1
End Function

