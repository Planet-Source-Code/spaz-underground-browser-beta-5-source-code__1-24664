VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A7C75093-2765-11D3-A0E4-FAFD20CEB591}#5.0#0"; "CBUTTON.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Undergroud Options"
   ClientHeight    =   5370
   ClientLeft      =   4170
   ClientTop       =   3780
   ClientWidth     =   6495
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6495
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optsearch 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search "
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   3000
      Width           =   1215
   End
   Begin VB.OptionButton optstartup 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Startup"
      Height          =   255
      Left            =   240
      TabIndex        =   39
      Top             =   2400
      Width           =   1335
   End
   Begin VB.OptionButton optnew 
      BackColor       =   &H00E0E0E0&
      Caption         =   "New Tabs"
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   1800
      Width           =   1335
   End
   Begin VB.OptionButton optbrowser 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Browser Tabs"
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   1200
      Width           =   1335
   End
   Begin VB.OptionButton optgeneral 
      BackColor       =   &H00E0E0E0&
      Caption         =   "General"
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4335
      Left            =   1920
      TabIndex        =   15
      Top             =   240
      Width           =   4455
      Begin VB.CheckBox chkBrowserTitleLength 
         Caption         =   "Limit Browser Title Length"
         Height          =   252
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   2292
      End
      Begin VB.TextBox txtBrowserTitleLength 
         Height          =   288
         Left            =   2640
         TabIndex        =   21
         Text            =   "35"
         Top             =   360
         Width           =   372
      End
      Begin VB.CheckBox chkRefreshBrowser 
         Caption         =   "Refresh All Browser Tabs Every"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   720
         Value           =   2  'Grayed
         Width           =   2535
      End
      Begin VB.TextBox txtRefreshBrowser 
         Height          =   285
         Left            =   3000
         TabIndex        =   19
         Text            =   "15"
         Top             =   720
         Width           =   375
      End
      Begin VB.Frame Frame6 
         Caption         =   "Popup Killa"
         Height          =   2415
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   4095
         Begin VB.CommandButton cmdpopups 
            Caption         =   "Popups"
            Height          =   375
            Left            =   1320
            TabIndex        =   47
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CheckBox chkpopupkilla 
            Caption         =   "Enable Popup Killa"
            Height          =   255
            Left            =   480
            TabIndex        =   18
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox chkquick 
            Caption         =   "Quick Disable Popups"
            Height          =   255
            Left            =   480
            TabIndex        =   17
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label Label7 
            Caption         =   "Popup options"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   1560
            Width           =   1095
         End
      End
      Begin VB.Label Label4 
         Caption         =   "minutes."
         Height          =   255
         Left            =   3480
         TabIndex        =   23
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   4335
      Left            =   1920
      TabIndex        =   41
      Top             =   240
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "Search Options"
         Height          =   495
         Left            =   1080
         TabIndex        =   42
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Sorry Could Deal with the Move so Click below to view Search Options"
         Height          =   975
         Left            =   360
         TabIndex        =   43
         Top             =   600
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   4455
      Begin MSComDlg.CommonDialog Dialog1 
         Left            =   2640
         Top             =   3840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin CButton.Button cmdbrowse 
         Height          =   255
         Left            =   3480
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BackColor       =   -2147483633
         SelectedBackColor=   -2147483633
         HoverBackColor  =   -2147483633
         HoverColor      =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AmbientColor    =   0   'False
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   "...."
         Alignment       =   0
         GroupNumber     =   0
      End
      Begin VB.TextBox txtlocation 
         Height          =   285
         Left            =   360
         TabIndex        =   45
         Top             =   3480
         Width           =   3015
      End
      Begin VB.CheckBox chkBack 
         Caption         =   "Use Backround Bitmap"
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   3240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Frame Frame5 
         Caption         =   "Underground History"
         Height          =   1455
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   3495
         Begin VB.TextBox txtHistoryDays 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   8
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton cmdClearHistory 
            Caption         =   "Clear History"
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "How many days of History do you want to keep?"
            Height          =   495
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "Days"
            Height          =   255
            Left            =   2760
            TabIndex        =   9
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.CheckBox chkDefaultBrowser 
         Caption         =   "Make Underground Browser the default browser"
         Height          =   252
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkMinToSysTray 
         Caption         =   "Minimize to system tray"
         Height          =   252
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   3612
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   4335
      Left            =   1920
      TabIndex        =   24
      Top             =   240
      Width           =   4455
      Begin VB.CheckBox chkNewTabHome 
         Caption         =   "Home"
         Height          =   252
         Left            =   240
         TabIndex        =   34
         Top             =   960
         Width           =   2052
      End
      Begin VB.CheckBox chkNewTabHistory 
         Caption         =   "History"
         Height          =   252
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Value           =   2  'Grayed
         Width           =   3252
      End
      Begin VB.CheckBox chkNewTabSearch 
         Caption         =   "Search"
         Height          =   252
         Left            =   240
         TabIndex        =   32
         Top             =   1200
         Value           =   1  'Checked
         Width           =   3372
      End
      Begin VB.Frame boxNewTabOption 
         Height          =   1332
         Left            =   240
         TabIndex        =   27
         Top             =   1800
         Width           =   3612
         Begin VB.OptionButton optDefaultNewButton 
            Caption         =   "New blank browser tab"
            Height          =   252
            Index           =   0
            Left            =   360
            TabIndex        =   30
            Tag             =   "2"
            Top             =   480
            Width           =   2172
         End
         Begin VB.OptionButton optDefaultNewButton 
            Caption         =   "New browser tab at current location"
            Height          =   252
            Index           =   1
            Left            =   360
            TabIndex        =   29
            Tag             =   "3"
            Top             =   720
            Width           =   3012
         End
         Begin VB.OptionButton optDefaultNewButton 
            Caption         =   "New browser tab at home page"
            Height          =   252
            Index           =   2
            Left            =   360
            TabIndex        =   28
            Tag             =   "1"
            Top             =   960
            Width           =   2652
         End
         Begin VB.Label Label2 
            Caption         =   "Default action for New button"
            Height          =   252
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   2292
         End
      End
      Begin VB.CheckBox chkNewTabFavorites 
         Caption         =   "Favorites"
         Height          =   252
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Value           =   2  'Grayed
         Width           =   1812
      End
      Begin VB.CheckBox chkNewTabAddressTyped 
         Caption         =   "Address typed in address field"
         Height          =   252
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   2772
      End
      Begin VB.Label Label3 
         Caption         =   "Open a New Tab when going to..."
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   4335
      Left            =   1920
      TabIndex        =   11
      Top             =   240
      Width           =   4455
      Begin VB.OptionButton optStartPage 
         Caption         =   "Start Underground Search at Home Page"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   3855
      End
      Begin VB.OptionButton optStartPage 
         Caption         =   "Start Underground Search with Last Open Sites"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   3855
      End
      Begin VB.OptionButton optStartPage 
         Caption         =   "Start Underground Search on a Blank page"
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   3855
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   5055
      Left            =   120
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkBack_Click()
    If chkBack.Value = vbChecked Then
        txtlocation.Enabled = True
        txtlocation.BackColor = &H80000005
        cmdbrowse.Enabled = True
    Else
        txtlocation.Enabled = False
        txtlocation.BackColor = &HE0E0E0
        cmdbrowse.Enabled = False
    End If
End Sub

Private Sub cmdApply_Click()
    'Save all the options and set focus to OK button
    Call SaveOptions
    OptionsSaved = True
    cmdOk.SetFocus
End Sub

Private Sub cmdbrowse_Click()
    Dialog1.DialogTitle = "Open a document..."
    Dialog1.Filter = "All Supported Files | *bmp;*.bmp;*.jpg;*Gif;*.gif;| All Files |*.*|"
    Dialog1.ShowOpen
    If Dialog1.FileName = "" Then
        Exit Sub
    Else
        txtlocation.Text = Dialog1.FileName
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not OptionsSaved Then Call SaveOptions
    Unload Me
End Sub

Private Sub cmdpopups_Click()
    frmpopupkilla.Show vbModal, frmBrowser
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    frmSearchOptions.Show vbModal, frmBrowser
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    'If the ESC key was entered
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    'Initialize form data
    OptionsSaved = False
    optgeneral.Value = True
    Frame1.Visible = True
    Frame1.Caption = ""
    Frame2.Visible = False
    Frame2.Caption = ""
    Frame3.Visible = False
    Frame3.Caption = ""
    Frame4.Visible = False
    Frame4.Caption = ""
    Frame7.Visible = False
    Frame7.Caption = ""
    '///////////////////////////////////////////////////////
    'GENERAL tab
    
    txtlocation.Text = GetString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "Background")
    chkMinToSysTray.Value = gMinToSysTray
    'chkDefaultBrowser.Enabled = False
    txtHistoryDays = gtxtHistoryDays
    chkpopupkilla = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "PopupKilla")
    '///////////////////////////////////////////////////////
    'BROWSER TABS tab
    chkBrowserTitleLength.Value = gBrowserTitleLength
    txtBrowserTitleLength.Text = gtxtBrowserTitleLength
    chkRefreshBrowser.Value = gchkRefreshBrowser
    txtRefreshBrowser.Text = gtxtRefreshBrowser
    
    '///////////////////////////////////////////////////////
    'NEW BROWSESR TAB tab
    chkNewTabHome.Value = gNewTabHome
    chkNewTabSearch.Value = gNewTabSearch
    chkNewTabAddressTyped.Value = gNewTabAddressTyped
    chkNewTabFavorites.Value = gNewTabFavorites
    chkNewTabHistory.Value = gNewTabHistory
    optDefaultNewButton(gDefaultNewButton).Value = True
    
    '///////////////////////////////////////////////////////
    'START UP tab
    optStartPage(gStartPage).Value = True
    If chkBack.Value = vbChecked Then
        txtlocation.Enabled = True
        txtlocation.BackColor = &H80000005
        cmdbrowse.Enabled = True
    Else
        txtlocation.Enabled = False
        txtlocation.BackColor = &HE0E0E0
        cmdbrowse.Enabled = False
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Unload frmSearchOptions
End Sub
Private Sub tbsOptions_KeyPress(KeyAscii As Integer)
    'If the ESC key was entered
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub optbrowser_Click()
    Frame1.Visible = False
    Frame2.Visible = True
    Frame3.Visible = False
    Frame4.Visible = False
    Frame7.Visible = False
End Sub

Private Sub optgeneral_Click()
    Frame1.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    Frame7.Visible = False
End Sub

Private Sub optnew_Click()
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = True
    Frame4.Visible = False
    Frame7.Visible = False
End Sub

Private Sub optsearch_Click()
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    Frame7.Visible = True
End Sub

Private Sub optstartup_Click()
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = True
    Frame7.Visible = False
End Sub
