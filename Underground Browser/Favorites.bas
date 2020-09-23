Attribute VB_Name = "Favorites"
Option Explicit

'//////////////////////////////////////////////////////
'// Favorites.bas
'// Code for retieving IE Favorites
'// Requires modCommon.bas
'Globals
Global GotoFav As String
Global iP(0 To 6) As Long
'Locals
Global a As Long
Dim b As Long
Dim c As Long
Dim FolderInd As Integer
Dim Buffer$
Dim MenuItemName As String
Dim Index As Long

Global FP As FILE_PARAMS
Global Itm As Node
Global favpath As String
Global sLastFolder As String
Global sRoot As String
Global bSubItem As Boolean
Global nCount As Long
Global bCancel As Boolean
Global CurParentDir As String
Global FoldersOnly As Boolean
Global gFavoritesURL As String      'v1.8.99
Public Sub GetFavorites()
    
    '///////////////////////////////////////////////////////////////
    'FAVORITES LOAD
    sLastFolder = ""
    'Load The Tree
    frmBrowser.TreeView1.Clear
    frmBrowser.TreeView1.Refresh
    
    'retrieve the special folder path
    'to the internet favorites
    favpath = GetFolderPath(CSIDL_FAVORITES)
    
    'Initializes the Root Item in the TreeView
    Call LoadTreeView("Internet Favorites", True, True)
    
    If Len(favpath) > 0 Then
        
        'set up the search UDT
        With FP
            .sFileRoot = favpath
            .sFileNameExt = "*.url"
            .bRecurse = True
        End With
        
        'Get Favorite folders only
        'FoldersOnly = True
        Call SearchForFilesArray(FP)
         frmBrowser.TreeView1.ItemExpanded(a) = True
    Else
        
        MsgBox " Could not locate favorites folder! " & _
        "This program requires Microsoft's Internet " & _
        "Explorer to be installed. Program will shutdown now!", _
        vbCritical + vbOKOnly, "FavMenu Error"
        End
        
    End If
    '///////////////////////////////////////////////////////////////
      
End Sub

Public Function TrimNull(startstr As String) As String
       
    'returns the string up to the first
    'null, if present, or the passed string
    Dim Pos As Integer
    
    Pos = InStr(startstr, Chr$(0))
    
    If Pos Then
        TrimNull = Left$(startstr, Pos - 1)
        Exit Function
    End If
    
    TrimNull = startstr
End Function
Private Function GetFileInformation(FP As FILE_PARAMS) As Long
On Error Resume Next
  'local working variables
Dim WFD As WIN32_FIND_DATA
Dim hFile As Long
Dim Pos As Long
Dim sPath As String
Dim sRoot As String
Dim sTmp As String
Dim sURL As String
Dim sShortcut As String
   
         
  'FP.sFileRoot (assigned to sRoot) contains
  'the path to search.
  '
  'FP.sFileNameExt (assigned to sPath) contains
  'the full path and filespec.
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & FP.sFileNameExt
   
  'obtain handle to the first filespec match
   hFile = FindFirstFile(sPath, WFD)
   
  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then

      Do
      
        'remove trailing nulls
         sTmp = TrimNull(WFD.cFileName)
         
        'Even though this routine uses filespecs,
        '*.* is still valid and will cause the search
        'to return folders as well as files, so a
        'check against folders is still required.
         If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) _
            = FILE_ATTRIBUTE_DIRECTORY Then
           
           'determine the link name by removing
           'the .url extension
            Pos = InStr(sTmp, ".url")
            
            If Pos > 0 Then
            
                sShortcut = Left$(sTmp, Pos - 1)
           
                'extract the URL
                sURL = ProfileGetItem("InternetShortcut", "URL", "", sRoot & sTmp)
                If sLastFolder = "" Then
                    'In The Root
                    Call LoadTreeView(sShortcut, False, False, "", sURL)
                Else
                    Call LoadTreeView(sShortcut, False, False, sLastFolder, sURL)
                End If
'              'add to the listview
'               Set itmX = ListView1.ListItems.Add(, , sShortcut)
'               itmX.SubItems(1) = sURL
         
            End If
         
         End If
         
      Loop While FindNextFile(hFile, WFD)
      
     'close the handle
      hFile = FindClose(hFile)
   
   End If
   
  'clean up

   
End Function
Private Function QualifyPath(sPath As String) As String
    
    
    'assures that a passed path ends in a slash
    If Right$(sPath, 1) <> "\" Then
        QualifyPath = sPath & "\"
    Else: QualifyPath = sPath
    End If
    
End Function

Public Function GetFolderPath(CSIDL As Long) As String
    
    Dim sPath As String
    Dim sTmp As String
    
    'fill pidl with the specified folder item
    sPath = Space$(MAX_LENGTH)
    
    If SHGetFolderPath(frmBrowser.hWnd, CSIDL, 0&, SHGFP_TYPE_CURRENT, sPath) = S_OK Then
        sTmp = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
    
    GetFolderPath = sTmp
    
End Function

Public Function ProfileGetItem(lpSectionName As String, _
    lpKeyName As String, _
    defaultValue As String, _
    inifile As String) As String
    
    
    'Retrieves a value from an ini file corresponding
    'to the section and key name passed.
    
    Dim success As Long
    Dim nSize As Long
    Dim ret As String
    
    'call the API with the parameters passed.
    'The return value is the length of the string
    'in ret, including the terminating null. If a
    'default value was passed, and the section or
    'key name are not in the file, that value is
    'returned. If no default value was passed (""),
    'then success will = 0 if not found.
    
    'Pad a string large enough to hold the data.
    ret = Space$(2048)
    nSize = Len(ret)
    success = GetPrivateProfileString(lpSectionName, lpKeyName, _
    defaultValue, ret, nSize, inifile)
    
    If success Then
        ProfileGetItem = Left$(ret, success)
    End If
End Function

Private Sub GetAllFilesSpecified(FP As FILE_PARAMS)
    
    Dim drvCount As Long
    Dim sBuffer As String
    Dim currDrive As String
    
    If FP.sFileRoot = "all fixed disks/partitions" Then
        
        'all drives
        
        'retrieve the available drives on the system
        sBuffer = Space$(64)
        drvCount = GetLogicalDriveStrings(Len(sBuffer), sBuffer)
        
        'drvCount returns the size of the drive string
        If drvCount Then
            
            'strip off trailing nulls
            sBuffer = Left$(sBuffer, drvCount)
            
            'search each drive for the file
            Do Until sBuffer = ""
                
                'strip off one drive item from sBuffer
                FP.sFileRoot = StripItem(sBuffer)
                
                'just search the local file system
                If GetDriveType(FP.sFileRoot) = DRIVE_FIXED Then
                    
                    'this may take a while, so update the
                    'display when the search path changes
                    'Text2.Text = "Working ... searching drive " & FP.sFileRoot
                    
                    DoEvents
                    If bCancel Then Exit Do
                    
                    Call SearchForFilesArray(FP)
                    
                    'Update the display count
                    'Text3.Text = Format$(nCount, sFileSoFar)
                    'Text3.Refresh
                    
                End If
                
            Loop
            
        End If
        
    Else
        
        Call SearchForFilesArray(FP)
        
    End If
   
End Sub

Private Sub SearchForFilesArray(FP As FILE_PARAMS)
On Error Resume Next
  'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
      
  'this routine is primarily interested in the
  'directories, so the file type must be *.*
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & "*.*"
   
  'obtain handle to the first match
   hFile = FindFirstFile(sPath, WFD)
   
  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then
   
     'GetFileInformation function returns the number,
     'of files matching the filespec (FP.sFileNameExt)
     'in the passed folder.
      Call GetFileInformation(FP)

      Do
      
        'if the returned item is a folder...
         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            
           'remove trailing nulls
            sTmp = TrimNull(WFD.cFileName)
            
           'and if the folder is not the default
           'self and parent folders...
            If sTmp <> "." And sTmp <> ".." Then
            
              'get the file
               FP.sFileRoot = sRoot & sTmp
               
               
              'This next If..Then just prevents adding extra
              'lines and unneeded paths to the array when a
              'file search is performed for a specific file type
               If FP.sFileNameExt = "*.*" Then
               
                 'Depending on the purpose, you may want to
                 'exclude the next 4 optional lines.
                 'The first two lines adds a blank entry
                 'to the array as a separator between new
                 'directories in the output file. The last
                 'two add the directory name alone, before
                 'listing the files underneath. These four
                 'lines can be optionally commented out).
                 'Obviously, these extra entries will skew
                 'the actual file counts.
                                 
               End If
               
              'call again
               Call SearchForFilesArray(FP)
            
            End If
               
            
         End If
         
     'continue looping until FindNextFile returns
     '0 (no more matches)
      Loop While FindNextFile(hFile, WFD)
      
     'close the find handle
      hFile = FindClose(hFile)
   
   End If
   
End Sub
Function StripItem(startStrg As String) As String
    
    
    'Take a string separated by Chr(0)'s, and split off 1 item, and
    'shorten the string so that the next item is ready for removal.
    Dim Pos As Integer
    
    Pos = InStr(startStrg, Chr$(0))
    
    If Pos Then
        StripItem = Mid(startStrg, 1, Pos - 1)
        startStrg = Mid(startStrg, Pos + 1, Len(startStrg))
    End If
   
End Function

Private Sub GetSystemDrives(ctl As ComboBox)
    
    
    Dim drvCount As Long
    Dim sBuffer As String
    Dim currDrive As String
    
    'Retrieve the available drives on the system.
    'The string is padded with enough room to hold
    'the drives, nulls and final terminating null.
    sBuffer = Space$(105)
    drvCount = GetLogicalDriveStrings(Len(sBuffer), sBuffer)
    
    'drvCount returns the size of the drive string
    If drvCount Then
        
        'strip off trailing nulls
        sBuffer = Left$(sBuffer, drvCount)
        
        'search each drive for the file
        Do Until sBuffer = ""
            
            'strip off one drive item from sBuffer
            currDrive = StripItem(sBuffer)
            
            'just search the local file system
            If GetDriveType(currDrive) = DRIVE_FIXED Then
                
                ctl.AddItem Left$(currDrive, 2)
                
            End If
            
        Loop
        
    End If
   
End Sub

Private Function GetFolderName(ByVal sPath As String) As String
    
    Dim length As Long
    Dim xPos As Long
    Dim sTemp As String
    
    GetFolderName = ""
    
    length = Len(sPath)
    xPos = length
    
    If Left(sPath, length) = "\" Then
        sPath = Left(sPath, (length - 1))
    End If
    
    Do Until xPos = 0
        xPos = xPos - 1
        
        If Mid$(sPath, xPos, 1) = "\" Then
            GetFolderName = Mid(sPath, (xPos - 1))
            Exit Do
        End If
        
    Loop
 
End Function

Public Sub LoadTreeView(ItemName As String, bFolder As Boolean, bRoot As Boolean, _
    Optional SubItem As String, Optional sURL As String)
    On Error Resume Next
    frmBrowser.TreeView1.hImageList = frmBrowser.m_cILMenu.hIml
    With frmBrowser.m_cMenu
    Dim AIndex As Long
        If bRoot Then
            a = frmBrowser.TreeView1.Add(0&, FirstChild, "", "Internet Favorites", 5)
            Exit Sub
        End If
        'SubItem - Tells which folder the items are under
        'bFolder - Reports True or False if it is a folder or not
        
        'Parse the file name if its to long.  52 Characters
        MenuItemName$ = ItemName$
        If Len(MenuItemName$) > 53 Then
            MenuItemName$ = Trim(Left$(MenuItemName$, 52)) & "..."
        End If
        '--------------------------------------------------
        
        Index = Index + 1
        If SubItem = "" Then
            If bFolder = True Then
                '               favorites_favoritesmnu
                iP(1) = .AddItem(ItemName, , , iP(0), 5, , , Index & "Folder")
                b = frmBrowser.TreeView1.Add(a, AlphabeticalChild, Index & "Folder", ItemName, 5, 5)
                FolderInd = 2
            Else
                iP(1) = .AddItem(MenuItemName, , , iP(0), 5, , , "URL" & " " & sURL)
                frmBrowser.TreeView1.Add a, LastChild, sURL & " " & Index, ItemName, 5
            End If
        Else
            
            If bFolder = True Then
                
                If Buffer = SubItem Then
                    
                    iP(FolderInd) = .AddItem(ItemName, , , iP(FolderInd - 1), 5, , , Index & "Folder")
                    
                    c = frmBrowser.TreeView1.Add(b, LastChild, Index & "Folder", ItemName, 5, 5)
                    FolderInd = FolderInd + 1
                Else
                    
                    FolderInd = FolderInd - 3
                    iP(FolderInd) = .AddItem(ItemName, , , iP(FolderInd), 5, , , Index & "Folder")
                    
                    c = frmBrowser.TreeView1.Add(b, LastChild, Index & "Folder", ItemName, 5, 5)
                    
                End If
                b = c
            Else
                
                iP(FolderInd) = .AddItem(MenuItemName, , , iP(FolderInd - 1), 5, , , "URL" & " " & sURL)
                frmBrowser.TreeView1.Add b, LastChild, sURL & " " & Index, ItemName, 5
                
            End If
        End If
    End With
    'Check to see if the subfolder has a subfolder of the subfoler..LOL
    Buffer = SubItem
End Sub

Public Function InRoot(ByVal sPath As String) As Boolean
    
    Dim sTmp As String
    
    InRoot = False
    
    sTmp = favpath & "\" & sPath
    
    If Dir(sTmp, vbDirectory) <> "" Then
        InRoot = True
    End If
End Function
