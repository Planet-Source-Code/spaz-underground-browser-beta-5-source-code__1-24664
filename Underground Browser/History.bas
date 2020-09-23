Attribute VB_Name = "History"
Option Explicit

Global DomainArray() As String
Global DomainCount As Integer
Global strDomain As String
Global strWeekOf As String
Global dtWeekOf As Date
Global dtTodayWeekOf As Date
Global dtLastWeek As Date
Global aDate As Date
Global dtToday As Date
Global tmpString As String
Global strDay As String
Global strLastDay As String
Global strURL, strTitle As String
Global strKey As String
Global Pos As Integer
Global gHistoryURL As String        'v1.8.99

Public Sub GetHistory()
   '
'///////////////////////////////////////////////////////////////
'THE DATABAS IS NO LONGER BEING USED.
On Error GoTo GetHistory_Error:
    Dim Conn As New ADODB.Connection
    Dim Cmd As New ADODB.Command
    Dim HistoryRs As New ADODB.Recordset
    
    InitializeHistoryLoad
    
    'Get History from History table in UndergroundHistory db
    Conn.ConnectionTimeout = 15
    Conn.CommandTimeout = 30
    Conn.open ("DBQ=./Underground.mdb;Driver={Microsoft Access Driver (*.mdb)};DriverId=25;MaxBufferSize=8192;Threads=20;")
    Cmd.ActiveConnection = Conn
    'Get a recordset of history data
    Cmd.CommandType = adCmdTable
    Cmd.CommandText = "History order by aDate asc;"
    HistoryRs.open Cmd, , 0, 1
    
    'Go through HistoryRs and populate the HistoryTree
    On Error Resume Next
    Do While Not HistoryRs.EOF
        aDate = HistoryRs.Fields("aDate")
        strURL = HistoryRs.Fields("URL")
        strTitle = HistoryRs.Fields("Title")
        ProcessHistoryURL
        HistoryRs.MoveNext
    Loop
        
    FinishHistoryLoad

    'Clean up
    Set Conn = Nothing
    Set Cmd = Nothing
    Set HistoryRs = Nothing
    
    Exit Sub

GetHistory_Error:
    Dim strMessage
    strMessage = Err.Description
    strMessage = Err.Source
    'Clean up
    Set Conn = Nothing
    Set Cmd = Nothing
    Set HistoryRs = Nothing

End Sub

Public Sub GetHistory_txt()
'
'///////////////////////////////////////////////////////////////
'Histoyr LOAD
On Error GoTo GetHistory_txt_Error_old:

    Dim tmpString As String
    
    InitializeHistoryLoad
    
    Open gProgPath & "UndergroundHistory.dat" For Input As #1
        
    Do Until EOF(1)
        Line Input #1, tempString
        If Left(tempString, 1) <> "*" Then
            aDate = Trim(Left(tempString, 10))
            tempString = Trim(Mid(tempString, 11))
            Pos = InStr(1, tempString, ";", vbTextCompare)
            strTitle = Trim(Mid(tempString, Pos + 1))
            strURL = Trim(Left(tempString, Len(tempString) - Len(strTitle) - 1))
            
            ProcessHistoryURL
            AddToHistoryCombo
            
        End If
    Loop
    
    Close #1
    
    FinishHistoryLoad
    Exit Sub

GetHistory_txt_Error_old:
    Close #1
    Dim strMessage
    strMessage = Err.Description
    strMessage = Err.Source
    If Err.Description = "File not found" Then
        'File didn't exist, create it with header
        Close #1
        Open gProgPath & "UndergroundHistory.dat" For Append As #1
        Print #1, "***************************************************"
        Print #1, "**  Underground History File"
        Print #1, "**  DO NOT MODIFY THIS FILE"
        Print #1, "**************************************************"
        Close #1
    End If
    FinishHistoryLoad
 
End Sub

Public Sub AddHistory(instrURL As String, instrTitle As String)
 '
On Error GoTo AddHistory_Error:
    Dim Conn As New ADODB.Connection
    Dim Cmd As New ADODB.Command
    Dim aDate As Date
    Dim dtWeekOf As Date
    Dim strWeekOf As String
    Dim strDay As String
    Dim Itm As Node
    
    aDate = Now
    strDay = WeekdayName(Weekday(aDate))
    strURL = instrURL
    strTitle = instrTitle
    
    SetTodayInfo
    
    ProcessHistoryURL
    
    frmBrowser.treeHistory.Refresh
    Exit Sub
    
AddHistory_Error:
    Dim strMessage
    strMessage = Err.Description
    strMessage = Err.Source

End Sub

Public Sub GetWeekOf()
    '
    Dim intDay As Integer
    
    intDay = Weekday(aDate, vbMonday)
    Select Case intDay
        Case 1: strWeekOf = "Week of " & Format(aDate, "mm/dd/yyyy", vbMonday)
                dtWeekOf = aDate
        Case 2: strWeekOf = "Week of " & Format(aDate - 1, "mm/dd/yyyy", vbMonday)
                dtWeekOf = aDate - 1
        Case 3: strWeekOf = "Week of " & Format(aDate - 2, "mm/dd/yyyy", vbMonday)
                dtWeekOf = aDate - 2
        Case 4: strWeekOf = "Week of " & Format(aDate - 3, "mm/dd/yyyy", vbMonday)
                dtWeekOf = aDate - 3
        Case 5: strWeekOf = "Week of " & Format(aDate - 4, "mm/dd/yyyy", vbMonday)
                dtWeekOf = aDate - 4
        Case 6: strWeekOf = "Week of " & Format(aDate - 5, "mm/dd/yyyy", vbMonday)
                dtWeekOf = aDate - 5
        Case 7: strWeekOf = "Week of " & Format(aDate - 6, "mm/dd/yyyy", vbMonday)
                dtWeekOf = aDate - 6
    End Select
    Exit Sub
End Sub

Public Sub MaintainHistoryDB()
    '
    Dim Conn As New ADODB.Connection
    Dim Cmd As New ADODB.Command
    Dim dtToday As Date
    
    'Get today's date
    dtToday = Format(Now, "mm/dd/yy")

    'Keep History table to the Underground Options number of days for History
    Conn.ConnectionTimeout = 15
    Conn.CommandTimeout = 30
    Conn.open ("DBQ=./Underground.mdb;Driver={Microsoft Access Driver (*.mdb)};DriverId=25;MaxBufferSize=8192;Threads=20;")
    Cmd.ActiveConnection = Conn
    Cmd.CommandType = adCmdText
    Cmd.CommandText = "delete FROM History WHERE (((History.aDate)<#" & dtToday - Int(gtxtHistoryDays) & "#));"
    Cmd.Execute
    Conn.Close
    Set Cmd = Nothing
    Set Conn = Nothing
End Sub

Public Sub MaintainHistoryTxt()
    '
    Dim tempString As String
    Dim CurentNode As Node
    Dim Pos As Integer
    Dim strTitle As String
    Dim strURL As String
    Dim strDate As String
    Dim aDate As Date

    Open gProgPath & "UndergroundHistory.dat" For Output As #1
    
    'Print File Header
    Print #1, "***************************************************"
    Print #1, "**  Underground History File"
    Print #1, "**  DO NOT MODIFY THIS FILE"
    Print #1, "**************************************************"
    
    For Each CurentNode In frmBrowser.treeHistory.Nodes
        If Right(CurentNode.Key, 3) = "URL" Then
            Pos = InStr(1, CurentNode.Text, "[", vbTextCompare)
            strTitle = Trim(Left(CurentNode.Text, Pos - 1))
            strURL = CurentNode.Tag
            strDate = Mid(CurentNode.Key, 9, 10)
            aDate = strDate
            If aDate > dtToday - Int(gtxtHistoryDays) Then Print #1, strDate & strURL & ";" & strTitle
        End If
    Next CurentNode
    
    Close #1

MaintainHistoryTxt_Error:
    If Err.Description = "File not found" Then
        'File didn't exist, create it with header
        Close #1
        Open gProgPath & "UndergroundHistory.dat" For Append As #1
        Print #1, "***************************************************"
        Print #1, "**  Underground History File"
        Print #1, "**  DO NOT MODIFY THIS FILE"
        Print #1, "**************************************************"
        Close #1
    End If
End Sub

Public Sub InitializeHistoryLoad()
    
    SetTodayInfo
    
    'Clear and Refresh the tree
    frmBrowser.treeHistory.Nodes.Clear
    frmBrowser.treeHistory.Refresh
    
    DomainCount = 0
    ReDim Preserve DomainArray(DomainCount)
    
Exit Sub



    ' Then UDBErrorHandler "(Module) History::Sub InitializeHistoryLoad"
    Resume Next
End Sub

Public Function GetDayString()
    
    Dim X As Integer
    strDay = ""
    For X = 0 To 6
        If dtToday - X = aDate Then strDay = WeekdayName(Weekday(aDate))
    Next
    If WeekdayName(Weekday(dtToday)) = strDay Then strDay = "Today"

End Function

Public Sub MakeWeekOfEntry()
    
    'Make a "Week Of" entry"
    strKey = strWeekOf
    Set Itm = frmBrowser.treeHistory.Nodes.Add(, , strKey, strWeekOf, 7)
    Itm.Tag = strWeekOf
    dtLastWeek = dtWeekOf
    DomainCount = 0
    ReDim Preserve DomainArray(DomainCount)
End Sub

Public Sub MakeDayOfEntry()
    
    'Make a Day Of entry
    strKey = "Day Of  " & Format(aDate, "mm/dd/yyyy")
    Set Itm = frmBrowser.treeHistory.Nodes.Add(, , strKey, strDay, 7)
    Itm.Tag = aDate
    strLastDay = strDay
    DomainCount = 0
    ReDim Preserve DomainArray(DomainCount)

End Sub

Public Sub GetURLDomain()
    
    'Get Domain of URL
    Pos = InStr(1, strURL, "://", vbTextCompare)
    strDomain = Mid(strURL, Pos + 3)
    Pos = InStr(1, strDomain, "/", vbTextCompare)
    If Pos > 0 Then
        strDomain = Left(strDomain, Pos - 1)
    Else
        strDomain = ""
    End If

End Sub

Public Sub ProcessHistoryURL()
    '
    GetWeekOf
    GetDayString
    
    If Format(dtWeekOf, "mm/dd/yyyy") <> "12/25/1899" Then
        If dtWeekOf <> dtTodayWeekOf Then
            If dtWeekOf <> dtLastWeek Then
                MakeWeekOfEntry
            End If
        Else
            If strDay <> strLastDay Then
                MakeDayOfEntry
            End If
        End If
    End If
     
    GetURLDomain
    
    'Check to see if this Domain has been put into the current week of the tree
    If LBound(DomainArray) = UBound(DomainArray) Then
        'array is empty Insert Domain into array and tree
        DomainCount = DomainCount + 1
        ReDim Preserve DomainArray(DomainCount)
        DomainArray(DomainCount) = strDomain
        Set Itm = frmBrowser.treeHistory.Nodes.Add(strKey, tvwChild, strKey & "_" & strDomain & "_Domain", strDomain, 6)
        Itm.Tag = strDomain
    Else
        Dim i As Integer
        Dim found As Boolean
        i = 1
        found = False
        While Not found And i <= UBound(DomainArray)
            If DomainArray(i) = strDomain Then
                found = True
            Else
                i = i + 1
            End If
        Wend
        If Not found Then
            DomainCount = DomainCount + 1
            ReDim Preserve DomainArray(DomainCount)
            DomainArray(DomainCount) = strDomain
            Set Itm = frmBrowser.treeHistory.Nodes.Add(strKey, tvwChild, strKey & "_" & strDomain & "_Domain", strDomain, 6)
            Itm.Tag = strDomain
        End If
    End If
    
    'Make a URL History entry
    If Left(strURL, 5) <> "about" And _
        Left(strURL, 5) <> "res:/" Then
            Set Itm = frmBrowser.treeHistory.Nodes.Add(strKey & "_" & strDomain & "_Domain", tvwChild, strKey & "_" & strDomain & "_" & strURL & "_URL", strTitle & " [" & strURL & "]", 4)
            Itm.Tag = strURL
    End If

End Sub

Public Sub FinishHistoryLoad()
    '
    'Make entry for TODAY into HistoryTree
    strKey = "Day Of  " & Format(Now, "mm/dd/yyyy")
    Set Itm = frmBrowser.treeHistory.Nodes.Add(, , strKey, "Today", 7)
    Itm.Tag = "Today"
    
    'Reset Domain array and count
    DomainCount = 0
    ReDim Preserve DomainArray(DomainCount)

End Sub

Public Sub SetTodayInfo()
    
    Dim intDay As Integer
    'Get today's date
    dtToday = Format(Now, "mm/dd/yy")
    intDay = Weekday(dtToday, vbMonday)
    Select Case intDay
        Case 1: dtTodayWeekOf = dtToday
        Case 2: dtTodayWeekOf = dtToday - 1
        Case 3: dtTodayWeekOf = dtToday - 2
        Case 4: dtTodayWeekOf = dtToday - 3
        Case 5: dtTodayWeekOf = dtToday - 4
        Case 6: dtTodayWeekOf = dtToday - 5
        Case 7: dtTodayWeekOf = dtToday - 6
    End Select
End Sub

Public Sub AddToHistoryCombo()
    
    frmBrowser.cboHistory.AddItem (strURL)
End Sub
