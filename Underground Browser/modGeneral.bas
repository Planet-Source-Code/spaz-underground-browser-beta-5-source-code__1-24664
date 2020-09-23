Attribute VB_Name = "modGeneral"
Public Const EM_GETLINECOUNT = &HBA
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hWnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long   ' <---

Public Const WM_GETTEXTLENGTH = &HE
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public PopupList As ListItem
Public PopupHeader As ColumnHeader
Public TransList As ListItem
Public TransHeader As ColumnHeader
Public LineNumbering As Integer
Public HighLight As Integer
Public LeftMargin As Integer
Public WhiteSpace As Integer
Public SelBounds As Integer

Public Function FindInListview(lstListName As ListView, _
                                strStringToFind As String, _
                                Optional bolWholeWordOnly As Boolean, _
                                Optional bolCaseSensitive As Boolean) _
                                As Boolean
    ' setup variables
    Dim lngIndex As Long        ' used for the current index of the parent items
    Dim lngIndexSub As Long     ' used for the current index of the subitems
    Dim strCurrItem As String   ' used to store the text of the currently selected item for compare
    
    ' if we want to be sensitive about the case then make the 'search' all upper case
    If bolCaseSensitive = True Then strStringToFind = UCase(strStringToFind)
    
    ' set the return to the default zero
    EnhListView_Find = 0
    
    ' if there is nothing to search then exit
    If lstListName.ListItems.Count < 1 Then Exit Function
    
    ' if no item is currently selected then select the first item
    If lstListName.SelectedItem.Index = -1 Then lstListName.SelectedItem.Index = 1
    
    ' move through the rows
    For lngIndex = lstListName.SelectedItem.Index - -1 To lstListName.ListItems.Count
        
        ' if we want to be sensitive about the case then...
        If bolCaseSensitive = True Then
            ' fill our variable with the uppercase version of the current text
            strCurrItem = UCase(lstListName.ListItems.Item(lngIndex).Text)
        Else
            ' otherwise, fill our variable with the current text
            strCurrItem = lstListName.ListItems.Item(lngIndex).Text
        End If
        
        If bolWholeWordOnly = True Then
            ' if the current item and the 'search' is an exact match then finalize
            If strCurrItem = strStringToFind Then FindInListview = True    'GoTo Finalize
             
        Else
            ' if the current item contains the 'search' then finalize
            If InStr(strCurrItem, strStringToFind) > 0 Then FindInListview = True ' GoTo Finalize
            
        End If
        
        ' if we have subitems...
        If lstListName.ColumnHeaders.Count > 1 Then
            
            ' move through the subitems of the current row
            For lngIndexSub = 1 To lstListName.ColumnHeaders.Count - 1
                ' if we want to be sensitive about the case then...
                If bolCaseSensitive = True Then
                    ' fill our variable with the uppercase version of the current text
                    strCurrItem = UCase(lstListName.ListItems.Item(lngIndex).SubItems(lngIndexSub))
                Else
                    ' otherwise, fill our variable with the current text
                    strCurrItem = lstListName.ListItems.Item(lngIndex).SubItems(lngIndexSub)
                End If
                
                If bolWholeWordOnly = True Then
                    ' if the current item and the 'search' is an exact match then finalize
                    If strCurrItem = strStringToFind Then FindInListview = True 'GoTo Finalize
                     
                Else
                    ' if the current item contains the 'search' then finalize
                    If InStr(strCurrItem, strStringToFind) > 0 Then FindInListview = True 'GoTo Finalize
                    
                End If
            ' move to next subitem
            Next lngIndexSub
        
        End If
        
    ' move to next row
    Next lngIndex
    
    Exit Function
    
Finalize:
    FindInListview = True
   
End Function


Public Sub LoadPopups()
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    frmSearchOptions.lstPopups.ColumnHeaders.Clear
    frmSearchOptions.lstPopups.ListItems.Clear
    Set PopupHeader = Nothing
    Set PopupHeader = frmSearchOptions.lstPopups.ColumnHeaders.Add(, , "Title", 2000, lvwColumnLeft)
    Set PopupHeader = frmSearchOptions.lstPopups.ColumnHeaders.Add(, , "Url", 7000, lvwColumnLeft)
        
    Set db = OpenDatabase(App.path & "\SearchEngines.udb")
    Set rs = db.OpenRecordset("Engines")
    Do Until rs.EOF = True
        Set PopupList = frmSearchOptions.lstPopups.ListItems.Add(, , rs.Fields("Title"))
        PopupList.SubItems(1) = rs.Fields("Url")
        rs.MoveNext
    Loop
    rs.Close
    db.Close

Exit Sub



    ' Then UDBErrorHandler "(Module) modGeneral::Sub LoadPopups"
    Resume Next
End Sub

Public Sub LoadTransalations()
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    frmSearchOptions.lstTrans.ColumnHeaders.Clear
    frmSearchOptions.lstTrans.ListItems.Clear
    Set TransHeader = Nothing
    Set TransHeader = frmSearchOptions.lstTrans.ColumnHeaders.Add(, , "Title", 2000, lvwColumnLeft)
    Set TransHeader = frmSearchOptions.lstTrans.ColumnHeaders.Add(, , "Url", 7000, lvwColumnLeft)
        
    Set db = OpenDatabase(App.path & "\SearchEngines.udb")
    Set rs = db.OpenRecordset("Translate")
    Do Until rs.EOF = True
        Set PopupList = frmSearchOptions.lstTrans.ListItems.Add(, , rs.Fields("Title"))
        PopupList.SubItems(1) = rs.Fields("Url")
        rs.MoveNext
    Loop
    rs.Close
    db.Close

End Sub

Public Sub SetComboHeight(oComboBox As ComboBox, _
lNewHeight As Long)
    

Dim oldscalemode As Integer

'This procedure does not work with frames: you
'cannot set the ScaleMode to vbPixels, because
'the frame does not have a ScaleMode Property.
'To get round this, you could set the parent control
'to be the form while you run this procedure.

If TypeOf oComboBox.Parent Is Frame Then Exit Sub

'Change the ScaleMode on the parent to Pixels.
oldscalemode = oComboBox.Parent.ScaleMode
oComboBox.Parent.ScaleMode = vbPixels

'Resize the combo box window.
MoveWindow oComboBox.hWnd, oComboBox.Left, _
oComboBox.Top, oComboBox.Width, lNewHeight, 1

'Replace the old ScaleMode
oComboBox.Parent.ScaleMode = oldscalemode

Exit Sub



    ' Then UDBErrorHandler "(Module) modGeneral::Sub SetComboHeight"
    Resume Next
End Sub

Public Function OpenBrowser(strURL As String, lngHwnd As Long)
    
    OpenBrowser = ShellExecute(lngHwnd, vbNullString, strURL, vbNullString, _
    "c:\", SW_SHOWDEFAULT)

Exit Function



    ' Then UDBErrorHandler "(Module) modGeneral::Function OpenBrowser"
    Resume Next
End Function

'=======================================================================
' Description: Resizes all Columns in a ListView to fit the text in
'              the rows
'=======================================================================
Public Function EnhListView_ResizeColumns( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '_______________________________________________________________________
    ' set function return to true
    EnhListView_ResizeColumns = True
    
    '_______________________________________________________________________
    ' if the user has not set LengthPerCharacter use 80
    If LengthPerCharacter = 0 Then LengthPerCharacter = "80"
    
    '_______________________________________________________________________
    ' if there are columns to go through...
    If lstListViewName.ListItems.Count > 0 Then
        ' setup variables
        Dim lngIndexCounter As Long
        Dim lngColumnCounter As Long
        ' move through each column
        For lngColumnCounter = 1 To lstListViewName.ColumnHeaders.Count
            ' move though each entry
            For lngIndexCounter = 1 To lstListViewName.ListItems.Count
                ' if it is not the first column
                If lngColumnCounter > 1 Then
                    ' size the column 85 twips per letter
                    If Len(lstListViewName.ListItems.Item(lngIndexCounter).SubItems(lngColumnCounter - 1)) * LengthPerCharacter > _
                    lstListViewName.ColumnHeaders.Item(lngColumnCounter).Width Then
                        lstListViewName.ColumnHeaders.Item(lngColumnCounter).Width = _
                        Len(lstListViewName.ListItems.Item(lngIndexCounter).SubItems(lngColumnCounter - 1)) * LengthPerCharacter
                    End If
                ' if it is the first column
                Else
                    ' size the column 85 twips per letter
                    If Len(lstListViewName.ListItems.Item(lngIndexCounter).Text) * LengthPerCharacter > _
                    lstListViewName.ColumnHeaders.Item(lngColumnCounter).Width Then
                        lstListViewName.ColumnHeaders.Item(lngColumnCounter).Width = _
                        Len(lstListViewName.ListItems.Item(lngIndexCounter).Text) * LengthPerCharacter
                    End If
                End If
            Next lngIndexCounter
        Next lngColumnCounter
    End If

Exit Function



    ' Then UDBErrorHandler "(Module) modGeneral::Function EnhListView_ResizeColumns"
    Resume Next
End Function

