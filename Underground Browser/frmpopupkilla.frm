VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmpopupkilla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Di Popup Killa [Slew Dem]"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "frmpopupkilla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2423
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtcaption 
      Height          =   285
      Left            =   1523
      TabIndex        =   4
      Top             =   3720
      Width           =   4815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Popup Captions / Locations"
      Height          =   2535
      Left            =   143
      TabIndex        =   2
      Top             =   120
      Width           =   6255
      Begin MSComctlLib.ListView lstpopups 
         Height          =   2175
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3836
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
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4463
      TabIndex        =   1
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   503
      TabIndex        =   0
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Popup Caption"
      Height          =   255
      Left            =   203
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "frmpopupkilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAdd_Click()
    
    If cmdAdd.Caption = "Add" Then
        Me.Height = 4740
        txtcaption.Text = ""
        
        cmdAdd.Caption = "Save"
    ElseIf cmdAdd.Caption = "Save" Then
        Dim PopupList As String
        Dim PopCaption As String
        Dim PopupUrl As String
        
        '
        PopCaption = txtcaption.Text
        
        PopupList = App.path & "/popups.dat"
        If txtcaption <> "" Then
            Trim (PopCaption)
            Open PopupList For Append As #1
            Print #1, PopCaption
            Close #1
        End If
        Call LoadPopupList(lstpopups, App.path & "/popups.dat")
        Me.Height = 3930
        cmdAdd.Caption = "Add"
    End If
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Height = 3885
    Call LoadPopupList(lstpopups, App.path & "/popups.dat")
    txtcaption.Text = ""
 Me.Caption = Me.Caption + " | Total Blocked " & Count_Lines_In_File(App.path & "/popups.dat")
End Sub
Public Sub LoadPopupList(MyList As ListView, MyFile As String)
    
    MyList.ListItems.Clear
    MyList.View = lvwReport
    Open MyFile For Input As #1
    Input #1, one$
    X = MyList.ColumnHeaders.Add(, , "The Tabs with the Following Captions will Be Killed", 6000, lvwColumnLeft)
    
    Do Until EOF(1)
        Input #1, one$
        X = MyList.ListItems.Add(, , one$)
    Loop
    Close #1
End Sub
Private Sub lstpopups_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstpopups
        If .SortKey <> ColumnHeader.Index - 1 Then
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        Else
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        End If
        .Sorted = True
    End With
End Sub

Function Count_Lines_In_File(ByVal strFilePath As String) As Integer
    'delcare variables
    Dim fileFile As Integer
    Dim intLinesReadCount As Integer
    intLinesReadCount = 0
    'open file
    fileFile = FreeFile
    
        Open strFilePath For Input As fileFile

    'loop through file
    Dim strBuffer As String
    
    
    Do While Not EOF(fileFile)
        'read line
        Input #fileFile, strBuffer
        'update count
        intLinesReadCount = intLinesReadCount + 1
    Loop
    'close file
    Close fileFile
    'return value
    Count_Lines_In_File = intLinesReadCount
End Function

