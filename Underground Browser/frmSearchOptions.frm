VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced Options"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmSearchOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opttrans 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Translation Urls"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1455
   End
   Begin VB.OptionButton optsearch 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search Engines"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.Frame FraSettings 
      Caption         =   "Search Options"
      Height          =   5415
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      Begin MSComctlLib.ListView lstPopups 
         Height          =   3015
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdwizard 
         Caption         =   "Wizard"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdremove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox TxtURL 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox TxtTitle 
         Height          =   285
         Left            =   840
         MaxLength       =   25
         TabIndex        =   3
         Top             =   480
         Width           =   3375
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox CboDefault 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "CboDefault"
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Url"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Title"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Default Search Engines"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   5040
         Width           =   1815
      End
   End
   Begin VB.Frame fraTranslation 
      Caption         =   "Translation Options"
      Height          =   5415
      Left            =   1800
      TabIndex        =   13
      Top             =   240
      Width           =   4455
      Begin VB.CommandButton cmdtransremove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   2520
         TabIndex        =   18
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdtransadd 
         Caption         =   "Add"
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Top             =   1440
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstTrans 
         Height          =   3135
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtTransUrl 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtTransTitle 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Translation Url"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   5775
      Left            =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frmSearchOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageLongByRef Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Integer, _
lParam As Long) _
As Long

Private Declare Function SendMessageByNum Lib "user32" _
Alias "SendMessageA" (ByVal hwnd As Long, ByVal _
wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const vbMsgBoxTopMost As Long = &H40000
Private Const LB_SETHORIZONTALEXTENT = &H194
Private Const LB_SETTABSTOPS = &H192

Private datMain As clsLoadSearchEngines

Private Sub CboDefault_Change()
    
    datMain.ChangeDefault CboDefault.Text

Exit Sub



    ' Then UDBErrorHandler "(Form) frmSearchOptions::Sub CboDefault_Change"
    Resume Next
End Sub

Private Sub CmdAdd_Click()
    
    If TxtTitle.Text = "" Then
        MsgBox "PLease Enter a Title for the Engine", vbInformation, "Error"
        TxtTitle.SetFocus
        Exit Sub
    End If
    If TxtURL.Text = "" Then
        MsgBox "PLease Enter a Url for the Engine", vbInformation, "Error"
        TxtURL.SetFocus
        Exit Sub
    End If
    With datMain
        .AddEngine TxtTitle, TxtURL
    End With
    Call LoadPopups
    TxtTitle.Text = ""
    TxtURL.Text = ""
    Call frmBrowser.LoadSrcEngines
Exit Sub



    ' Then UDBErrorHandler "(Form) frmSearchOptions::Sub CmdAdd_Click"
    Resume Next
End Sub
Private Sub DeleteSelected()
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = OpenDatabase(App.path & "\SearchEngines.udb")
    Set rs = db.OpenRecordset("Engines")
    Do While Not rs.EOF
        If rs.Fields("Title") = lstpopups.SelectedItem Then
            
            Debug.Print "Ready to Delete Redord"
            rs.delete
            Call LoadPopups
            db.Close
            Exit Do
        Else
            rs.MoveNext
        End If
    Loop

Exit Sub



    ' Then UDBErrorHandler "(Form) frmSearchOptions::Sub DeleteSelected"
    Resume Next
End Sub

Private Sub cmdremove_Click()
    
    Call DeleteSelected

Exit Sub



    ' Then UDBErrorHandler "(Form) frmSearchOptions::Sub cmdremove_Click"
    Resume Next
End Sub

Private Sub cmdtransadd_Click()
    
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = OpenDatabase(App.path & "\SearchEngines.udb")
    Set rs = db.OpenRecordset("Translate")
    If txtTransTitle.Text = "" Then
        MsgBox "PLease Enter a Title for the Engine", vbInformation, "Error"
        txtTransTitle.SetFocus
        Exit Sub
    End If
    If txtTransUrl.Text = "" Then
        MsgBox "PLease Enter a Url for the Engine", vbInformation, "Error"
        txtTransUrl.SetFocus
        Exit Sub
    End If
    rs.AddNew
    rs.Fields("Title") = txtTransTitle.Text
    rs.Fields("Url") = txtTransUrl.Text
    rs.Update
    rs.Close
    db.Close
    Call LoadTransalations
   
Exit Sub



    ' Then UDBErrorHandler "(Form) frmSearchOptions::Sub cmdtransadd_Click"
    Resume Next
End Sub

Private Sub cmdtransremove_Click()
    
Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = OpenDatabase(App.path & "\SearchEngines.udb")
    Set rs = db.OpenRecordset("Translate")
    Do While Not rs.EOF
        If rs.Fields("Title") = lstTrans.SelectedItem Then
            
            Debug.Print "Ready to Delete Redord"
            rs.delete
            Call LoadTransalations
            db.Close
            Exit Do
        Else
            rs.MoveNext
        End If
    Loop

Exit Sub



    ' Then UDBErrorHandler "(Form) frmSearchOptions::Sub cmdtransremove_Click"
    Resume Next
End Sub

Private Sub cmdwizard_Click()
    
    frmURLWizard.Show

Exit Sub



    ' Then UDBErrorHandler "(Form) frmSearchOptions::Sub cmdwizard_Click"
    Resume Next
End Sub

Private Sub Form_Load()
    
    Set datMain = New clsLoadSearchEngines
    Call LoadTransalations
    Call LoadPopups
    Call EnhListView_ResizeColumns(lstpopups, False)
    Call EnhListView_ResizeColumns(lstTrans, False)
    With datMain
        '.RefreshLBEngines LstEngines
        .RefreshCBEngines CboDefault
        CboDefault.Text = .GetDefault
    End With
    FraSettings(0).Visible = True
    fraTranslation.Visible = False

Exit Sub



    ' Then UDBErrorHandler "(Form) frmSearchOptions::Sub Form_Load"
    Resume Next
End Sub
Private Sub SetTabsTops(ByVal LBHwnd As Long, Stop1 As Integer)
    
    Dim tabsets&(1)
    tabsets(0) = Stop1
    Call SendMessageLongByRef(LBHwnd, LB_SETTABSTOPS, 2, tabsets(0))

Exit Sub



    ' Then UDBErrorHandler "(Form) frmSearchOptions::Sub SetTabsTops"
    Resume Next
End Sub

Private Sub SetScrollBar(lstHwnd As Long, jLen As Long)
    
    Dim strTemp As String
    Dim intX As Long
    intX = TextWidth(String$(jLen, "X"))
    If ScaleMode = vbTwips Then
        intX = intX / Screen.TwipsPerPixelX  ' if twips change to pixels
        SendMessageByNum lstHwnd, LB_SETHORIZONTALEXTENT, intX, 0
    End If

Exit Sub



    ' Then UDBErrorHandler "(Form) frmSearchOptions::Sub SetScrollBar"
    Resume Next
End Sub

'Private Sub lstPopups_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'With lstPopups
'        If .SortKey <> ColumnHeader.Index - 1 Then
'            .SortKey = ColumnHeader.Index - 1
'            .SortOrder = lvwAscending
'        Else
'            If .SortOrder = lvwAscending Then
'                .SortOrder = lvwDescending
'            Else
'                .SortOrder = lvwAscending
'            End If
'        End If
'        .Sorted = True
'    End With
'End Sub
Private Sub optsearch_Click()
    
    If optsearch.Value = True Then
        FraSettings(0).Visible = True
        fraTranslation.Visible = False
    Else
        FraSettings(0).Visible = False
        fraTranslation.Visible = True
    End If

Exit Sub



    ' Then UDBErrorHandler "(Form) frmSearchOptions::Sub optsearch_Click"
    Resume Next
End Sub

Private Sub opttrans_Click()
    
    If opttrans.Value = True Then
        FraSettings(0).Visible = False
        fraTranslation.Visible = True
    Else
        FraSettings(0).Visible = True
        fraTranslation.Visible = False
    End If

Exit Sub



    ' Then UDBErrorHandler "(Form) frmSearchOptions::Sub opttrans_Click"
    Resume Next
End Sub
