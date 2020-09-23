VERSION 5.00
Begin VB.Form frmURLWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add new Search Engine"
   ClientHeight    =   3465
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6000
   Icon            =   "frmURLWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<-- Previous"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next -->"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   4815
      Begin VB.TextBox txtURL 
         Height          =   645
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label lblInstruction 
         Caption         =   "Open your browser an go to the search engine you wish to add. Click ""Next"" to continue."
         Height          =   1215
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Label lblStep 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   6000
      X2              =   0
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Search Engine Wizard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmURLWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const conHwndTopmost = -1
Private Const conSwpNoActivate = &H10
Private Const conSwpShowWindow = &H40
Private Const vbMsgBoxTopMost As Long = &H40000

Private TestSearch As String
Private URLStem As String

Private Sub CancelButton_Click()
    
    Unload Me

Exit Sub



    ' Then UDBErrorHandler "(Form) frmURLWizard::Sub CancelButton_Click"
    Resume Next
End Sub

Private Sub cmdNext_Click()
    
    Dim datMain As clsLoadSearchEngines
    Set datMain = New clsLoadSearchEngines
    
    Select Case lblStep.Caption
    Case "1"
        GoStep "2"
    Case "2"
        GoStep "3"
    Case "3"
    If Len(TxtURL) = 0 Then
            MsgBox "The Url cannot be blank", vbMsgBoxTopMost
            TxtURL.SetFocus
            Exit Sub
            Else
        GoStep "4"
        End If
    Case "4"
        GoStep "5"
    Case "5"
        If Len(TxtURL) = 0 Then
            MsgBox "The title cannot be blank", vbMsgBoxTopMost
            TxtURL.SetFocus
            Exit Sub
        Else
            With datMain
                .AddEngine TxtURL, URLStem
                '.RefreshCBEngines frmBrowser.cboDefault
                .RefreshCBEngines frmBrowser.cboengine
                Call LoadPopups
                Call frmBrowser.LoadSrcEngines
            End With
            Unload Me
        End If
    Case "!"
        Unload Me
    End Select

Exit Sub



    ' Then UDBErrorHandler "(Form) frmURLWizard::Sub cmdNext_Click"
    Resume Next
End Sub

Private Sub GoStep(StepNo As String)
    
    
    Select Case StepNo
    Case "1"
        lblStep.Caption = "1"
        lblInstruction.Caption = "Please go to the search engine you wish to add. Click " & Chr(34) & "Next" & Chr(34) & " to continue."
        cmdPrevious.Enabled = False
    Case "2"
        lblStep.Caption = "2"
        lblInstruction.Caption = "In the search box of your search engine, Search for " & Chr(34) & _
        "Winzip" & Chr(34) & " and submit the search. Click " & Chr(34) & _
        "Next" & Chr(34) & " to continue."
        cmdPrevious.Enabled = True
        TxtURL.Visible = False
    Case "3"
        lblStep.Caption = "3"
        lblInstruction.Caption = "When the search is complete, copy and paste the internet address " & _
        "from the result page EXACTLY as it appears into the the box below. Click " & Chr(34) & _
        "Next" & Chr(34) & " to continue."
        TxtURL.Visible = True
    Case "4"
        TxtURL.Visible = True
        lblStep.Caption = "4"
        lblInstruction.Caption = "This is the Important part you need to change " & Chr(34) & _
        "The Search term Winzip with SearchString EXACTLY (capital [S]) or it will not work" & Chr(34) & " Click " & Chr(34) & _
        "Next" & Chr(34) & " to continue."
    Case "5"
        lblStep.Caption = "5"
        TxtURL.Height = 285
        
        lblInstruction.Caption = "Please ensure that It is correct, Now Enter a title for the Search Engine and click finish. Otherwise press back or cancel."
        URLStem = TxtURL.Text
        TxtURL.Text = ""
        TxtURL.Visible = True
        cmdNext.Caption = "Finish"
    End Select

Exit Sub



    ' Then UDBErrorHandler "(Form) frmURLWizard::Sub GoStep"
    Resume Next
End Sub

Private Sub cmdPrevious_Click()
    
    Select Case lblStep.Caption
    Case "1"
    Case "2"
        GoStep "1"
    Case "3"
        GoStep "2"
    Case "4"
        GoStep "3"
    Case "5"
        GoStep "4"
    Case "!"
        GoStep "3"
    End Select

Exit Sub



    ' Then UDBErrorHandler "(Form) frmURLWizard::Sub cmdPrevious_Click"
    Resume Next
End Sub

Private Sub Form_Load()
    
    ' SetWindowPos hwnd, conHwndTopmost, ConvertTwipsToPixels(Me.Left, 0), ConvertTwipsToPixels(Me.Top, 1), ConvertTwipsToPixels(Me.Width, 0), ConvertTwipsToPixels(Me.Height, 1), conSwpNoActivate Or conSwpShowWindow

Exit Sub



    ' Then UDBErrorHandler "(Form) frmURLWizard::Sub Form_Load"
    Resume Next
End Sub
