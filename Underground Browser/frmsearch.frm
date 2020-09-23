VERSION 5.00
Begin VB.Form frmsearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search The Underground"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optcurrent 
      Caption         =   "Use Current Tab"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   1920
      Width           =   2295
   End
   Begin VB.OptionButton optNewBrowserTab 
      Caption         =   "Search in new Tab"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ComboBox cboEngine 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Text            =   "Select a Search Engine"
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox searchtxt 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Engine to Use"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Search For"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrLocation As String
Dim StrSearch As String

Private Sub cmdclose_Click()
    
    Unload Me

Exit Sub



    ' Then UDBErrorHandler "(Form) frmsearch::Sub cmdclose_Click"
    Resume Next
End Sub

Private Sub cmdsearch_Click()
    
    If cboengine.Text = "Select a Search Engine" Then
        MsgBox "Please select a search engine to search", vbInformation, "Select Engine"
    End If
    If searchtxt.Text = "" Then
        MsgBox "Please enter at least 1 word to search for", vbInformation, "Enter Word"
        'GoTo Skip
    End If
    Select Case cboengine.Text
    Case "Astalavista"
        StrLocation = "http://astalavista.box.sk/cgi-bin/astalavista/robot?srch=" & StrSearch & " &submit=Search"
    End Select
    
    If optNewBrowserTab.Value = True Then
        Call frmBrowser.NewTab(Me, StrSearch, -1)
    Else
        If optcurrent.Value = True Then
            frmBrowser.brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate StrSearch
        End If
    End If

Exit Sub



    ' Then UDBErrorHandler "(Form) frmsearch::Sub cmdsearch_Click"
    Resume Next
End Sub

Private Sub Form_Load()
    
   cboengine.AddItem "Astalavista"

Exit Sub



    ' Then UDBErrorHandler "(Form) frmsearch::Sub Form_Load"
    Resume Next
End Sub

