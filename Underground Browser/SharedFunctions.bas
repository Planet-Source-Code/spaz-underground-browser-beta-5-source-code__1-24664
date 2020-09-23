Attribute VB_Name = "SharedFunctions"
Option Explicit

Global G_LoadFinished As Boolean
Global CurTab_Index As Integer
Global MaxTab_Index As Integer
Global AddURL, ClickHistory As Boolean
Global BrowserArrayEmpty As Boolean
Global OpenOk As Boolean
Global OpenURL As String
Global OpenNewTab As Boolean
Global SetFocusOnly As Boolean
Global InTabSetFocus As Boolean
Global Deleting As Boolean
Global found As Boolean
Global URLcount As Integer
Global strLocationName As String
Global strLocationURL As String
Global tempString As String
Global gProgPath As String
Global lLeft, lTop, lWidth, lHeight As Long
Global ViewingFavorites As Boolean
Global ViewingHistory As Boolean
Global UnLoading As Boolean
Global HistoryFileChanged As Boolean
Global bBackSpace As Boolean
Global bDeleteKey As Boolean
Global gbFullScreen As Boolean
Global gbMaximized As Boolean
Global glCurTopPos As Long
Global glCurLeftPos As Long
Global glCurWidth As Long
Global glCurHeight As Long
Global bRightMouse As Boolean
Global gUndergroundTop As Long
Global gUndergroundLeft As Long
Global gUndergroundWidth As Long
Global gUndergroundHeight As Long
Global gUndergroundState As Integer

'Options
Global OptionsSaved As Boolean
'General
Global gMinToSysTray As Long
Global gDefaultBrowser As Long
Global gtxtHistoryDays As String
'Browser Tabs
Global gBrowserTitleLength As Long
Global gtxtBrowserTitleLength As String
Global gchkRefreshBrowser As Long
Global gtxtRefreshBrowser As String
'New Tabs
Global gNewTabHome As Long
Global gNewTabSearch As Long
Global gNewTabAddressTyped As Long
Global gNewTabFavorites As Long
Global gNewTabHistory As Long
Global gDefaultNewButton As Long
'Startup
Global gStartPage As Long

Function GetFormInfo()
    
    'GET THE POSITION AND SIZE OF BROWSER FORM
    If GetString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Position", "BrowserTop") <> "" Then
        frmBrowser.Top = GetString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Position", "BrowserTop")
        frmBrowser.Left = GetString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Position", "BrowserLeft")
        frmBrowser.Width = GetString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Position", "BrowserWidth")
        frmBrowser.Height = GetString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Position", "BrowserHeight")
        Dim tmpBool
        tmpBool = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Split", "ViewFavorites")
        If tmpBool Then
            Call frmBrowser.mnu_ViewFavorites_Click
        End If
        tmpBool = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Split", "ViewHistory")
        If tmpBool Then
            Call frmBrowser.ViewHistory
        End If
        
        G_LoadFinished = True
    End If
End Function

Function SaveFormInfo()
    
    'SAVE THE POSITION AND SIZE OF BROWSER FORM
    Call SaveString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Position", "BrowserTop", str(gUndergroundTop))
    Call SaveString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Position", "BrowserLeft", str(gUndergroundLeft))
    Call SaveString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Position", "BrowserWidth", str(gUndergroundWidth))
    Call SaveString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Position", "BrowserHeight", str(gUndergroundHeight))
    Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Split", "ViewFavorites", ViewingFavorites)
    Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Split", "ViewHistory", ViewingHistory)
    
End Function

Function GetSavedTabURLs()
    
    On Error GoTo GetSavedTabURLs_Error:
    Dim X As Integer
    X = 0
    Dim URL As String
    
    URLcount = Int(GetString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\SavedURLs", "Count"))
    
    If URLcount > 0 Then
        'Set the current tabs browser to saved URL
        frmBrowser.TabStrip1.Tabs(CurTab_Index).Caption = GetString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\SavedURLs", str(X + 1) + "Location")
        frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).Visible = True
        URL = GetString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\SavedURLs", str(X + 1) + "URL")
        frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).Navigate URL
        frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).ZOrder (1)
        X = 1
        
        While X <= URLcount - 1
            'Increment WebBrowser indexes
            MaxTab_Index = MaxTab_Index + 1
            CurTab_Index = CurTab_Index + 1
            frmBrowser.TabStrip1.Tabs.Add
            frmBrowser.TabStrip1.Tabs.Item(CurTab_Index).Caption = GetString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\SavedURLs", str(X + 1) + "Location")
            frmBrowser.TabStrip1.Tabs.Item(CurTab_Index).Selected = True
            frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag = CurTab_Index - 1
            'Load new browser, and enable it
            Load frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag)
            frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).Visible = True
            frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).ZOrder (1)
            URL = GetString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\SavedURLs", str(X + 1) + "URL")
            frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).Navigate URL
            X = X + 1
        Wend
        
        MaxTab_Index = URLcount
        Call DeleteKey(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\SavedURLs")
        
    End If
    Exit Function
    
GetSavedTabURLs_Error:
    ShowErrorMessageBox ("GetSavedTabURLs")
   
End Function

Function SaveCurrentTabURLs()
    
    On Error GoTo SaveCurrentTabURLs_Error:
    Dim X As Integer
    
    For X = 1 To frmBrowser.TabStrip1.Tabs.Count
        Call SaveString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\SavedURLs", str(X) + "URL", frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(X).Tag).LocationURL)
        Call SaveString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\SavedURLs", str(X) + "Location", frmBrowser.TabStrip1.Tabs(X).Caption)
    Next
    
    Call SaveString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\SavedURLs", "Count", str(X - 1))
    Exit Function
    
SaveCurrentTabURLs_Error:
    ShowErrorMessageBox ("SaveCurrentTabURLs")
  
End Function

Public Sub PrintBrowser()
    
    frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).SetFocus
    SendKeys "^p"
End Sub

Sub AddToFavorites()
    
    Dim shellHelper As New ShellUIHelper
    Dim strLocationName, strLocationURL As String
    strLocationName = frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).LocationName
    strLocationURL = frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).LocationURL
    shellHelper.AddFavorite strLocationURL, strLocationName
    
End Sub

Sub SaveOptions()
    
    On Error GoTo SaveOptions_Error:
    
    '/////////////////////////////////////////////////////////////////
    'General tab
    Call SaveString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "Background", frmOptions.txtlocation.Text)
    
    Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkMinToSysTray", frmOptions.chkMinToSysTray.Value)
    gMinToSysTray = frmOptions.chkMinToSysTray.Value
    
    '    Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkDefaultBrowser", frmOptions.chkDefaultBrowser.Value)
    '    gDefaultBrowser = frmOptions.chkDefaultBrowser.Value
    
    'Set File Association
    If frmOptions.chkDefaultBrowser.Value = vbChecked Then
         
'         Associate "Undergound Browser", ".htm", "HTML Document", App.path & "\Agent.ico"
'         Associate "Undergound Browser", ".html", "HTML Document", App.path & "\Agent.ico"
        MakeFileType "htm", "HTML Document", App.path & "\Agent.ico", "Open", App.path & "\" & App.EXEName & ".exe %1", False, True, False
        MakeFileType "html", "HTML Document", App.path & "\Agent.ico", "Open", App.path & "\" & App.EXEName & ".exe %1", False, True, False
    End If
    
    gtxtHistoryDays = frmOptions.txtHistoryDays.Text
    Call SaveString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "txtHistoryDays", gtxtHistoryDays)
    
    'Save The Popup Killa Options
    Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "PopupKilla", frmOptions.chkpopupkilla.Value)
    
    'Save the Timer Interval
    Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "QuickDisable", frmOptions.chkquick.Value)
    
    '/////////////////////////////////////////////////////////////////
    'Browser Tabs tab
    
    Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkBrowserTitleLength", frmOptions.chkBrowserTitleLength.Value)
    gBrowserTitleLength = frmOptions.chkBrowserTitleLength.Value
    
    Call SaveString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "txtBrowserTitleLength", frmOptions.txtBrowserTitleLength.Text)
    gtxtBrowserTitleLength = frmOptions.txtBrowserTitleLength.Text
    gchkRefreshBrowser = frmOptions.chkRefreshBrowser.Value
    gtxtRefreshBrowser = frmOptions.txtRefreshBrowser.Text
    
    Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkRefreshBrowser", gchkRefreshBrowser)
    Call SaveString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "txtRefreshBrowser", gtxtRefreshBrowser)
    
    '/////////////////////////////////////////////////////////////////
    'New Tabs tab
    Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkNewTabHome", frmOptions.chkNewTabHome.Value)
    gNewTabHome = frmOptions.chkNewTabHome.Value
    
    Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkNewTabSearch", frmOptions.chkNewTabSearch.Value)
    gNewTabSearch = frmOptions.chkNewTabSearch.Value
    
    Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkNewTabAddressTyped", frmOptions.chkNewTabAddressTyped.Value)
    gNewTabAddressTyped = frmOptions.chkNewTabAddressTyped.Value
    
    Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkNewTabFavorites", frmOptions.chkNewTabFavorites.Value)
    gNewTabFavorites = frmOptions.chkNewTabFavorites.Value
    
    Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkNewTabHistory", frmOptions.chkNewTabHistory.Value)
    gNewTabHistory = frmOptions.chkNewTabHistory.Value
    
    If frmOptions.optDefaultNewButton(0).Value = True Then
        Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "optDefaultNewButton", 0)
        gDefaultNewButton = 0
    Else
        If frmOptions.optDefaultNewButton(1).Value = True Then
            Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "optDefaultNewButton", 1)
            gDefaultNewButton = 1
        Else
            Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "optDefaultNewButton", 2)
            gDefaultNewButton = 2
        End If
    End If
    
    '/////////////////////////////////////////////////////////////////
    'Startup tab
    If frmOptions.optStartPage(0).Value = True Then
        Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "optStartPage", 0)
        gStartPage = 0
    Else
        If frmOptions.optStartPage(1).Value = True Then
            Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "optStartPage", 1)
            gStartPage = 1
        Else
            Call SaveDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "optStartPage", 2)
            gStartPage = 2
        End If
    End If
    Exit Sub
    
SaveOptions_Error:
    ShowErrorMessageBox ("SaveOptions")
    
    Exit Sub
End Sub

Sub GetOptions()
    
    On Error GoTo GetOptions_Error:
    
    '/////////////////////////////////////////////////////////////////
    'General tab
    gMinToSysTray = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkMinToSysTray")
    gDefaultBrowser = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkDefaultBrowser")
    gtxtHistoryDays = GetString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "txtHistoryDays")
    If gtxtHistoryDays = "" Then gtxtHistoryDays = "20"
    
    '/////////////////////////////////////////////////////////////////
    'Browser Tabs tab
    gBrowserTitleLength = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkBrowserTitleLength")
    gtxtBrowserTitleLength = GetString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "txtBrowserTitleLength")
    If gtxtBrowserTitleLength = "" Then gtxtBrowserTitleLength = "0"
    gchkRefreshBrowser = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkRefreshBrowser")
    gtxtRefreshBrowser = GetString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "txtRefreshBrowser")
    
    '/////////////////////////////////////////////////////////////////
    'New Tabs tab
    gNewTabHome = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkNewTabHome")
    gNewTabSearch = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkNewTabSearch")
    gNewTabAddressTyped = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkNewTabAddressTyped")
    gNewTabFavorites = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkNewTabFavorites")
    gNewTabHistory = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "chkNewTabHistory")
    gDefaultNewButton = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "optDefaultNewButton")
    
    '/////////////////////////////////////////////////////////////////
    'Startup tab
    gStartPage = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "optStartPage")
    Exit Sub
    
GetOptions_Error:
    ShowErrorMessageBox ("GetOptions")
End Sub
Sub SetProgPath()
    
    On Error GoTo SetProgPath_Error:
    ' If dragged file is in the root, append filename.
    If Mid(App.path, Len(App.path)) = "\" Then
        gProgPath = App.path
        ' If dragged file is not in root, append "\" and filename.
    Else
        gProgPath = App.path & "\"
    End If
    Exit Sub
    
SetProgPath_Error:
    ShowErrorMessageBox ("SetProgPath")
End Sub
Sub RepositionProgressBar()
    
    'RESIZE and POSITION THE PROGRESS BAR
    lLeft = frmBrowser.StatusBar.Panels(2).Left + 10
    lTop = frmBrowser.StatusBar.Top + 30
    lWidth = frmBrowser.StatusBar.Panels(2).Width - 20
    lHeight = frmBrowser.StatusBar.Height - 40
    frmBrowser.ProgressBar.Move lLeft, lTop, lWidth, lHeight
   
End Sub

Sub GetTypedURLs()
    
    On Error GoTo GetTypedURLs_Error:
    'Get the TypedURLs from the registry
    'And populate the address list
    Dim done As Boolean
    Dim URLnum As Integer
    
    done = False
    URLnum = -1
    
    While Not done
        
        URLnum = URLnum + 1
        tempString = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs", "url" & URLnum)
        
        If URLnum = 0 And tempString = "" Then
            'Do nothing
        Else
            
            If tempString <> "" Then
                frmBrowser.cboAddress.AddItem tempString
            Else
                done = True
            End If
            
        End If
        
    Wend
    Exit Sub
    
GetTypedURLs_Error:
    ShowErrorMessageBox ("GetTypedURLs")
    
End Sub

Sub SaveTypedURLs()
    
    On Error GoTo SaveTypedURLs_Error:
    'Save the TypedURLs in the address list
    'To the registry in the TypedURLs key used by IE
    Dim done As Boolean
    Dim URLnum As Integer
    
    For URLnum = 0 To frmBrowser.cboAddress.ListCount - 1
        frmBrowser.cboAddress.ListIndex = URLnum
        tempString = frmBrowser.cboAddress.Text
        Call SaveString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs", "url" & URLnum, tempString)
    Next
    Exit Sub
    
SaveTypedURLs_Error:
    ShowErrorMessageBox ("SaveTypedURLs")
End Sub

Sub SelectBrowserTab(Index As Integer)
    
    frmBrowser.TabStrip1.Tabs(Index).Selected = True
    frmBrowser.TabStrip1.Refresh
    frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(Index).Tag).SetFocus
      
End Sub

Sub ShowErrorMessageBox(Where As String)
    
    'MsgBox ("Error in " & Where & ", " & vbCrLf & _
            "Please report the bug to Underground web site http://BlacksWeb.com/Underground" & vbCrLf & _
            "Or contact Jim@BlacksWeb.com.  Thank you.")

Exit Sub

End Sub

Public Sub FullScreen()
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    '*******************************************************
    '** Called when Full Screen button is clicked or F11 Key
    '*******************************************************
    With frmBrowser
        If gbFullScreen = True Then
            gbFullScreen = False
            a = frmBrowser.rbrMain.BandIndexForData("TOOLBAR")
            frmBrowser.rbrMain.BandVisible(a) = False
            b = frmBrowser.rbrMain.BandIndexForData("Address")
            frmBrowser.rbrMain.BandVisible(b) = False
            c = frmBrowser.rbrMain.BandIndexForData("Translate")
            frmBrowser.rbrMain.BandVisible(c) = False
            d = frmBrowser.rbrMain.BandIndexForData("Search")
            frmBrowser.rbrMain.BandVisible(d) = False
            .StatusBar.Visible = True
            .TabStrip1.Visible = True
            frmBrowser.Form_Resize
        ElseIf gbFullScreen = False Then
            gbFullScreen = True
            a = frmBrowser.rbrMain.BandIndexForData("TOOLBAR")
            frmBrowser.rbrMain.BandVisible(a) = True
            b = frmBrowser.rbrMain.BandIndexForData("Address")
            frmBrowser.rbrMain.BandVisible(b) = True
            c = frmBrowser.rbrMain.BandIndexForData("Translate")
            frmBrowser.rbrMain.BandVisible(c) = True
            d = frmBrowser.rbrMain.BandIndexForData("Search")
            frmBrowser.rbrMain.BandVisible(d) = True
             .StatusBar.Visible = True
            frmBrowser.Form_Resize
        End If
    End With
End Sub

Sub MoveBrowserOffFormTab(Index As Integer)
    '
    'Don't move the browser unless you have to
    If frmBrowser.brwWebBrowser(Index).Left <> frmBrowser.ScaleLeft + 10000000 Then
        frmBrowser.brwWebBrowser(Index).Left = frmBrowser.ScaleLeft + 10000000
    End If
    
    Exit Sub
    '
    '
    '
    '    ' Then UDBErrorHandler "(Module) SharedFunctions::Sub MoveBrowserOffFormTab"
    '    Resume Next
End Sub

Sub MoveBrowsers()
    
    '/////////////////////////////////////////////////////////////////
    'RESIZE THE BROWSER WINDOW
    Dim lLeft, lTop, lWidth, lHeight As Long
    lLeft = frmBrowser.TabStrip1.Left + 60
    lTop = frmBrowser.TabStrip1.Top + 340
    lWidth = frmBrowser.TabStrip1.Width - 215
    lHeight = frmBrowser.TabStrip1.Height - 400
    
    'Hide all other browser except current browser
    Dim j As Integer
    On Error Resume Next
    
    For j = 0 To frmBrowser.brwWebBrowser.UBound
        If j = Int(frmBrowser.TabStrip1.Tabs.Item(CurTab_Index).Tag) Then
            Dim something
            something = "something"
            'We only need to resize the current browser, becuase all
            'other browsers have been moved off the form and are hidden.
            If lWidth > 0 And lHeight > 0 Then
                frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs.Item(CurTab_Index).Tag).Move lLeft, lTop, lWidth, lHeight
                frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs.Item(CurTab_Index).Tag).ZOrder 0
                frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs.Item(CurTab_Index).Tag).Visible = True
            Else
                Dim i
                For i = 0 To frmBrowser.brwWebBrowser.UBound
                    'brwWebBrowser(i).Visible = False
                    Call MoveBrowserOffFormTab(frmBrowser.TabStrip1.Tabs(i + 1).Tag)
                Next i
            End If
        Else
            Call MoveBrowserOffFormTab(j)
        End If
    Next j
End Sub

