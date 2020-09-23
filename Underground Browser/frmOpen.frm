VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOpen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open"
   ClientHeight    =   2310
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5745
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   4560
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      Filter          =   "Web Documents (*.htm, *.html)|*.htm, *.html|Any (*.*)|*.*"
   End
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   4935
   End
   Begin VB.CheckBox optNewBrowserTab 
      Caption         =   "Open on a new browser tab"
      Height          =   252
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Value           =   1  'Checked
      Width           =   2532
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   372
      Left            =   4320
      TabIndex        =   5
      Top             =   1800
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   1212
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   372
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "Open:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Type the Internet address of a document or folder, and Underground Browser will open it for you."
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboAddress_Change()
    
    If cboAddress.Text <> "" Then cmdOk.Enabled = True

Exit Sub



    ' Then UDBErrorHandler "(Form) frmOpen::Sub cboAddress_Change"
    Resume Next
End Sub

Private Sub cboAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If cboAddress.Text <> "" Then cmdOk.Enabled = True

Exit Sub



    ' Then UDBErrorHandler "(Form) frmOpen::Sub cboAddress_KeyDown"
    Resume Next
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    
    If cboAddress.Text <> "" Then cmdOk.Enabled = True

Exit Sub



    ' Then UDBErrorHandler "(Form) frmOpen::Sub cboAddress_KeyPress"
    Resume Next
End Sub

Private Sub cmdbrowse_Click()
    
    'Open a common dialog for browsing to a document
    'Then load the path to that document into the address text box
    frmOpen.CommonDialog.Filter = "Web Documents (*.htm, *.html)|*.htm;*.html|Any (*.*)|*.*"
    frmOpen.CommonDialog.Flags = &H1000 Or &H2000000 Or &H800
    frmOpen.CommonDialog.Action = 1
    
    If frmOpen.CommonDialog.FileName <> "" Then
        frmOpen.cboAddress.Text = frmOpen.CommonDialog.FileName
    End If
    
    frmOpen.cmdOk.Enabled = True
    frmOpen.cmdOk.SetFocus

Exit Sub



    ' Then UDBErrorHandler "(Form) frmOpen::Sub cmdBrowse_Click"
    Resume Next
End Sub

Private Sub cmdCancel_Click()
    
    Unload frmOpen

Exit Sub



    ' Then UDBErrorHandler "(Form) frmOpen::Sub cmdCancel_Click"
    Resume Next
End Sub

Private Sub cmdOK_Click()
    
    If optNewBrowserTab.Value = 1 Then
        OpenNewTab = True
    End If
    
    OpenURL = cboAddress
    OpenOk = True
    Unload frmOpen

Exit Sub



    ' Then UDBErrorHandler "(Form) frmOpen::Sub cmdOK_Click"
    Resume Next
End Sub

Private Sub Form_Load()
    
    'Initialize form data
    cmdOk.Enabled = False
    optNewBrowserTab.Value = 1
    cboAddress.Text = "http://"

Exit Sub



    ' Then UDBErrorHandler "(Form) frmOpen::Sub Form_Load"
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Unload Me

Exit Sub



    ' Then UDBErrorHandler "(Form) frmOpen::Sub Form_Unload"
    Resume Next
End Sub
