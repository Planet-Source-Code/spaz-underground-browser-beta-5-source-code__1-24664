Attribute VB_Name = "Registry"
' -----------------
' ADVAPI32
' -----------------
' function prototypes, constants, and type definitions
' for Windows 32-bit Registry API

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&

' Registry API prototypes

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number

Public Sub savekey(hKey As Long, strPath As String)
    
    Dim keyhand&
    r = RegCreateKey(hKey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)

Exit Sub



    ' Then UDBErrorHandler "(Module) Registry::Sub savekey"
    Resume Next
End Sub

Public Function GetString(hKey As Long, strPath As String, strValue As String)
    
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)

End Sub

Function GetDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    
    r = RegOpenKey(hKey, strPath, keyhand)
    
     ' Get length/data type
    lDataBufSize = 4
        
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDword = lBuf
        End If
    'Else
    '    Call errlog("GetDWORD-" & strPath, False)
    End If
    
    r = RegCloseKey(keyhand)
End Function

Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then Call errlog("SetDWORD", False)
    r = RegCloseKey(keyhand)
End Function

Public Function DeleteKey(ByVal hKey As Long, ByVal strKey As String)
    
    Dim r As Long
    r = RegDeleteKey(hKey, strKey)

End Function

Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function
Public Sub CreateKey(ByVal hKey As Long, ByVal Key As String, Optional SubKey As Variant)

    Dim hHnd As Long
    
    If Not IsMissing(SubKey) Then
        Temp = RegCreateKey(hKey, Key & "\" & SubKey, hHnd)
        Temp = RegCloseKey(hHnd)
    Else
        Temp = RegCreateKey(hKey, Key, hHnd)
        Temp = RegCloseKey(hHnd)
    End If

End Sub
Public Sub SaveString1(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueTitle As String, ByVal ValueData As String)

    Dim hHnd As Long
    Dim Temp As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegCreateKey(hKey, KeyPath, hHnd)
    Temp = RegSetValueEx(hHnd, ValueTitle, 0, REG_SZ, ByVal ValueData, Len(ValueData))
    Temp = RegCloseKey(hHnd)

End Sub
Public Sub DeleteValue1(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal Value As String)

    Dim hHnd As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegOpenKey(hKey, KeyPath, hHnd)
    Temp = RegDeleteValue(hHnd, Value)
    Temp = RegCloseKey(hHnd)

End Sub

Public Function GetString1(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String) As String

    Dim hHnd As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lValueType As Long
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim Temp As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegOpenKey(hKey, KeyPath, hHnd)
    lResult = RegQueryValueEx(hHnd, ValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)

    If lValueType = REG_SZ Then
        strBuf = String(lDataBufferSize, " ")
        lResult = RegQueryValueEx(hHnd, ValueName, 0&, 0&, ByVal strBuf, lDataBufferSize)

        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            
            If intZeroPos > 0 Then
                GetString1 = Left$(strBuf, intZeroPos - 1)
            Else
                GetString1 = strBuf
            End If
        
        End If
    End If
End Function

'Function GetBinaryValue(SubKey As String, Entry As String)
'
'   Call ParseKey(SubKey, MainKeyHandle)
'
'   If MainKeyHandle Then
'      rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey)   'open the key
'      If rtn = ERROR_SUCCESS Then   'if the key could be opened
'         lBufferSize = 1
'         rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, 0, lBufferSize)   'get the value from the registry
'         sBuffer = Space(lBufferSize)
'         rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, sBuffer, lBufferSize)   'get the value from the registry
'         If rtn = ERROR_SUCCESS Then   'if the value could be retreived then
'            rtn = RegCloseKey(hKey)   'close the key
'            GetBinaryValue = sBuffer   'return the value to the user
'         Else   'otherwise, if the value couldnt be retreived
'            GetBinaryValue = "Error"   'return Error to the user
'            If DisplayErrorMsg = True Then   'if the user wants to errors displayed
'               MsgBox ErrorMsg(rtn)   'display the error to the user
'            End If
'         End If
'      Else   'otherwise, if the key couldnt be opened
'         GetBinaryValue = "Error"   'return Error to the user
'         If DisplayErrorMsg = True Then   'if the user wants to errors displayed
'            MsgBox ErrorMsg(rtn)   'display the error to the user
'         End If
'      End If
'   End If
'
'End Function
'
'Function SetBinaryValue(SubKey As String, Entry As String, Value As String)
'   Dim i As Integer
'
'   Call ParseKey(SubKey, MainKeyHandle)
'
'   If MainKeyHandle Then
'      rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey)   'open the key
'      If rtn = ERROR_SUCCESS Then   'if the key was open successfully then
'         lDataSize = Len(Value)
'         ReDim ByteArray(lDataSize)
'         For i = 1 To lDataSize
'            ByteArray(i) = Asc(Mid$(Value, i, 1))
'         Next
'         rtn = RegSetValueExB(hKey, Entry, 0, REG_BINARY, ByteArray(1), lDataSize)   'write the value
'         If Not rtn = ERROR_SUCCESS Then   'if the was an error writting the value
'            If DisplayErrorMsg = True Then   'if the user want errors displayed
'               MsgBox ErrorMsg(rtn)   'display the error
'            End If
'         End If
'         rtn = RegCloseKey(hKey)   'close the key
'      Else   'if there was an error opening the key
'         If DisplayErrorMsg = True Then   'if the user wants errors displayed
'            MsgBox ErrorMsg(rtn)   'display the error
'         End If
'      End If
'   End If
'
'End Function
