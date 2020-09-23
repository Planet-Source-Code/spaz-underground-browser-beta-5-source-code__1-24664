VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A7C75093-2765-11D3-A0E4-FAFD20CEB591}#5.0#0"; "CBUTTON.OCX"
Object = "{54F463F3-0135-11D2-8D52-00C04FA4EE99}#7.2#0"; "VBALTBAR.OCX"
Object = "{1FE9A10D-50A4-431B-89AE-610EC623D3F1}#1.0#0"; "VBALIML.OCX"
Object = "{F1909D6D-FB9D-11D3-B06C-00500427A693}#1.0#0"; "XUITREEVIEW6.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "Underground Search"
   ClientHeight    =   7485
   ClientLeft      =   2340
   ClientTop       =   1230
   ClientWidth     =   10200
   Icon            =   "frmBrowser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboAddress 
      Height          =   315
      ItemData        =   "frmBrowser.frx":08CA
      Left            =   360
      List            =   "frmBrowser.frx":08CC
      TabIndex        =   23
      Text            =   "http://"
      Top             =   720
      Width           =   6930
   End
   Begin MSComctlLib.ListView lstpopups 
      Height          =   1575
      Left            =   0
      TabIndex        =   22
      Top             =   5400
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame FavoritesCoolBar 
      Height          =   400
      Left            =   0
      TabIndex        =   19
      Top             =   1560
      Width           =   2655
      Begin CButton.Button cmdCloseFavorites 
         Height          =   240
         Left            =   2160
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   423
         BackColor       =   -2147483633
         SelectedBackColor=   -2147483633
         HoverBackColor  =   -2147483633
         HoverColor      =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         Caption         =   "X"
         Alignment       =   1
         GroupNumber     =   0
      End
      Begin VB.Label lblFavoritesCoolBar 
         Caption         =   "Favorites"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   125
         Width           =   1455
      End
   End
   Begin vbalIml6.vbalImageList m_cILMenu 
      Left            =   8760
      Top             =   5880
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   16
      Size            =   6580
      Images          =   "frmBrowser.frx":08CE
      KeyCount        =   7
      Keys            =   "ÿÿÿÿÿÿ"
   End
   Begin xuiTreeView6.TreeView TreeView1 
      Height          =   855
      Left            =   0
      TabIndex        =   18
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1508
      Lines           =   0   'False
      LabelEditing    =   0   'False
      PlusMinus       =   0   'False
      RootLines       =   0   'False
      ToolTips        =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxScrollTime   =   0
   End
   Begin VB.PictureBox picAnim 
      BackColor       =   &H00000000&
      Height          =   340
      Left            =   9240
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   15
      Top             =   1680
      Width           =   325
      Begin VB.Image imgIcon 
         Height          =   240
         Left            =   20
         Picture         =   "frmBrowser.frx":22A2
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame fraSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   9975
      Begin VB.TextBox txtsearch 
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   15
         Width           =   2295
      End
      Begin VB.CheckBox chknewtab 
         Caption         =   "Search in New Tab"
         Height          =   255
         Left            =   7560
         TabIndex        =   12
         Top             =   15
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.ComboBox cboengine 
         Height          =   315
         Left            =   2640
         TabIndex        =   11
         Text            =   "Select a Search Engine"
         Top             =   15
         Width           =   3015
      End
      Begin CButton.Button cmdsearch 
         Height          =   255
         Left            =   5760
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   30
         Width           =   1215
         _ExtentX        =   2143
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
         Caption         =   "Search"
         Alignment       =   0
         GroupNumber     =   0
      End
      Begin CButton.Button cmdclose 
         Height          =   255
         Left            =   9360
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Close Current Tab"
         Top             =   15
         Width           =   255
         _ExtentX        =   450
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         Caption         =   "X"
         Alignment       =   3
         GroupNumber     =   0
      End
      Begin CButton.Button cmdcloseall 
         Height          =   255
         Left            =   9600
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Close All Tabs"
         Top             =   15
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Picture         =   "frmBrowser.frx":282C
         BackColor       =   -2147483633
         SelectedBackColor=   -2147483633
         HoverBackColor  =   -2147483633
         HoverColor      =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         Caption         =   ""
         Alignment       =   0
         GroupNumber     =   0
      End
   End
   Begin VB.ComboBox CboTranslate 
      Height          =   315
      Left            =   7320
      TabIndex        =   9
      Top             =   480
      Width           =   2055
   End
   Begin vbalIml6.vbalImageList vbalImageList1 
      Left            =   8760
      Top             =   5280
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   32
      IconSizeY       =   32
      ColourDepth     =   16
      Size            =   48776
      Images          =   "frmBrowser.frx":2E9C
      KeyCount        =   14
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin vbalTBar.cToolbar tbrTools 
      Left            =   240
      Top             =   120
      _ExtentX        =   4260
      _ExtentY        =   661
   End
   Begin vbalTBar.cToolbar tbrMenu 
      Left            =   3120
      Top             =   120
      _ExtentX        =   5741
      _ExtentY        =   661
   End
   Begin vbalTBar.cReBar rbrMain 
      Left            =   120
      Top             =   0
      _ExtentX        =   12091
      _ExtentY        =   1296
   End
   Begin MSComctlLib.ImageList Navigation 
      Left            =   8640
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":ED44
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":F620
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":FEFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":107D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":110B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":11990
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1226C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":12B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":13424
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":13D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":145DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":14EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":15794
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":16070
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1694C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstFreedBrowserList 
      Height          =   255
      Left            =   8640
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picSplitter 
      Height          =   4935
      Left            =   2520
      ScaleHeight     =   4935
      ScaleWidth      =   15
      TabIndex        =   7
      Top             =   2160
      Width           =   15
   End
   Begin VB.ListBox lstHistory 
      Height          =   840
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComctlLib.ImageList imgTreeFavorites 
      Left            =   2040
      Top             =   5760
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":17228
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":177C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":18618
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1946C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":19A08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1A85C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1ADF8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   7200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList imlTrayIcons 
      Left            =   9075
      Top             =   6480
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1BC4C
            Key             =   "earth"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1C09E
            Key             =   "monit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1C4F0
            Key             =   "Underground"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   290
      Left            =   0
      TabIndex        =   2
      Top             =   7200
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Picture         =   "frmBrowser.frx":1CDCC
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   4800
      Index           =   0
      Left            =   2880
      TabIndex        =   1
      Top             =   2040
      Width           =   5370
      ExtentX         =   9472
      ExtentY         =   8467
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5415
      Left            =   2760
      TabIndex        =   0
      Top             =   1680
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9551
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Blank"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   8640
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   30
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1CF28
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1D622
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1DD1C
            Key             =   "Delete All"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1E3A0
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1EA9A
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1F194
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1F88E
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1FF88
            Key             =   "Home"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":20682
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":20D7C
            Key             =   "Favorites"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":21476
            Key             =   "History"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":21B70
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2226A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":22964
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":230DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":23858
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":23FD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2474C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":24EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":25640
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":25DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":26534
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":26CAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":27428
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":27BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2831C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":28A96
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":29190
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2990A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":29FD6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treeHistory 
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2143
      _Version        =   393217
      Indentation     =   212
      LabelEdit       =   1
      Style           =   1
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgTreeFavorites"
      Appearance      =   1
      MouseIcon       =   "frmBrowser.frx":2A646
   End
   Begin VB.ComboBox cboHistory 
      Height          =   315
      Left            =   6600
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Now the Underground Browser was built on a Open Source Project, Started
'by the People at Tiger Studio's. I Have put a lot of time in this one
'to make it what it is, you can compare the two to get a real idea of the
'time and work this Browser got from me. You might also notice the Editor
'it is an Early release of cEdit which is one of the best out there, and
'i used it because of the Syntax highlighting which is pretty cool, plus
'the browser a little more than just a Rich Text Box showing the internal HTML.
Option Explicit
Dim IlValue As Long
Dim StrSearch As String
Dim StrPopupCaption As String
Dim StrLocation As String
Dim TransUrl As String
Dim PageUrl As String
Private WithEvents m_cSplit As cSplitter
Attribute m_cSplit.VB_VarHelpID = -1
Public StartingAddress As String
Public QuickDisable As Boolean
Private LoadEngines As clsLoadSearchEngines
Private WithEvents gSysTray As clsSysTray
Attribute gSysTray.VB_VarHelpID = -1
Public WithEvents m_cMenu As cPopupMenu
Attribute m_cMenu.VB_VarHelpID = -1
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Const CB_FINDSTRING = &H14C
Private Const CB_ERR = (-1)
Private Const CB_SETCURSEL = &H14E
'__________________________________
Implements ISubclass
Private m_emr As EMsgResponse
Private Sub brwWebBrowser_DocumentComplete(Index As Integer, ByVal pDisp As Object, URL As Variant)
    
    On Error GoTo DocumentComplete_Error:
    Dim brwLocation As String
    Dim brwURL As String
    Dim X As Integer
    brwLocation = pDisp.Document.nameProp
    brwURL = URL
    Debug.Print "Location is "; brwLocation
    Debug.Print "Url is "; brwURL
    'Add URL to Underground History Tree
    If brwURL <> "" And brwLocation <> "" And brwLocation <> "Blank" Then
        'This is the Popup Killer
        If PopupKill = True Then
            If FindInListview(lstPopups, brwLocation, False, False) = True Then
                Call DeleteTab
                StatusBar.Panels(1).Text = "Suck U Madda " & brwLocation
                Debug.Print "Suck U Madda " & brwLocation
            ElseIf FindInListview(lstPopups, brwURL, False, False) = True Then
                Call DeleteTab
                StatusBar.Panels(1).Text = "Suck U Madda " & brwURL
                Debug.Print "Suck U Madda " & brwURL
                Call AddHistory(brwURL, brwLocation)
            End If
        End If
    End If
    Exit Sub
    
DocumentComplete_Error:
    ShowErrorMessageBox ("DocumentComplete")
    
End Sub

Private Sub brwWebBrowser_DownloadComplete(Index As Integer)
    
    '*** Occures to often, not used ***
    'SetTabCaption (brwWebBrowser(index).LocationName)
    
End Sub

Private Sub brwWebBrowser_NavigateComplete2(Index As Integer, ByVal pDisp As Object, URL As Variant)
    
    Dim CurName As String
    On Error GoTo NavigateComplete2_Error:
    'Add URL to History Pull Down
    'If URL exists in the list, remove it and
    'Insert it at index 0 to keep the most recent at the top
    
    If AddURL = True Or ClickHistory = True Then
        Dim CurAddress As String
        Dim i As Integer
        Dim found As Boolean
        
        CurName = pDisp.Document.nameProp
        CurAddress = URL
        i = InStr(1, CurAddress, "//", vbTextCompare) + 2
        If Mid(CurAddress, 2, 1) <> ":" Then CurAddress = Mid(CurAddress, i)
        
        found = False
        i = 0
        While i <= cboAddress.ListCount And Not found
            If cboAddress.List(i) = CurAddress Then
                found = True
            End If
            i = i + 1
        Wend
        
        If Not found Then
            cboAddress.AddItem CurAddress, 0
        Else
            'Delete the item and add item as index 0
            cboAddress.RemoveItem i - 1
            cboAddress.Text = URL
            cboAddress.AddItem CurAddress, 0
        End If
        
        AddURL = False
        ClickHistory = False
    End If
    Exit Sub
    
NavigateComplete2_Error:
    ShowErrorMessageBox ("NavigateComplete2")
    
End Sub

Private Sub brwWebBrowser_NewWindow2(Index As Integer, ppDisp As Object, Cancel As Boolean)
    
    On Error GoTo NewWindow2_Error:
    'Should do a New Tab here...
    If QuickDisable = False Then
        Dim URL As String
        URL = ""
        Call NewTab(Me, URL, -99)
        Set ppDisp = brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Object
        Call SelectBrowserTab(CurTab_Index)
        TabStrip1.SetFocus
        Exit Sub
    ElseIf QuickDisable = True Then
        Cancel = True
    End If
NewWindow2_Error:
    ShowErrorMessageBox ("NewWindow2")
    
End Sub

Private Sub brwWebBrowser_OnFullScreen(Index As Integer, ByVal FullScreen As Boolean)
    On Error Resume Next
    'This should bring the Browser Back Down
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).FullScreen = False
    FullScreen = False
    'This should move to a Next Tab, which makes it go back to Normal Size
    Call mnu_GoNextTab_Click
    Call SelectBrowserTab(CurTab_Index)
    MoveBrowsers
End Sub

Private Sub brwWebBrowser_OnQuit(Index As Integer)
    On Error Resume Next
    Call DeleteTab
End Sub

Private Sub brwWebBrowser_ProgressChange(Index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
    
    On Error GoTo ProgressChange_Error:
    'Display Progress Indicator in the status bar
    
    If Index = TabStrip1.tabs(CurTab_Index).Tag And brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Busy Then
        
        If Progress = -1 Then
            ProgressBar.Value = 0
            ProgressBar.Visible = False
        End If
        
        If Progress > 0 And ProgressMax > 0 Then
            
            If ((Progress * 100) / ProgressMax) < 100 Then
                ProgressBar.Value = (Progress * 100) / ProgressMax
                ProgressBar.Visible = True
                ProgressBar.ZOrder (0)
                'make sure progress bar is position correctly
                'some times it doesn't reposition it self properly
                'durring resize
                RepositionProgressBar
            Else
                ProgressBar.Visible = False
            End If
            
        End If
        
    End If
    Exit Sub
    
ProgressChange_Error:
    ShowErrorMessageBox ("ProgressChange_Error")
    
End Sub

Private Sub brwWebBrowser_StatusTextChange(Index As Integer, ByVal Text As String)
    
    On Error GoTo StatusTextChange_Error:
    'Display in Statusbar, the status of the browser
    'Downlaod status and hyperlink fly-overs
    If Index = TabStrip1.tabs(CurTab_Index).Tag Then
        StatusBar.Panels(1).Text = Text
    End If
    Exit Sub
    
StatusTextChange_Error:
    ShowErrorMessageBox ("StatusTextChange")
    
End Sub

Private Sub brwWebBrowser_TitleChange(Index As Integer, ByVal Text As String)
    On Error Resume Next
    'If the title of the doc has changed, make adjustments
    Dim sCaption As String
    Dim theTab As Integer
    Dim found As Boolean
    
    theTab = 1
    found = False
    
    If Len(Text) > 0 Then
        If Len(brwWebBrowser(Index).LocationURL) > 0 Then
            If gBrowserTitleLength = 1 And Int(gtxtBrowserTitleLength) <> 0 And Len(Text) > Int(gtxtBrowserTitleLength) Then
                sCaption = Left(Text, Int(gtxtBrowserTitleLength) - 3) & "..."
            Else
                sCaption = Text
            End If
        End If
        
        'Must find the appropriate tab to update
        While Not found
            'Find the correct tab with a tag = index
            If TabStrip1.tabs(theTab).Tag = Index Then
                found = True
                TabStrip1.tabs(theTab).Caption = sCaption
                If TabStrip1.SelectedItem.Index = theTab Then
                    'Also update cboAddress and form caption
                    'cboAddress.Text = brwWebBrowser(index).LocationURL
                    'If Left(cboAddress.Text, 5) = "about" Then cboAddress.Text = ""
                    Me.Caption = sCaption & " - " & PROGRAM_NAME
                End If
            Else
                theTab = theTab + 1
            End If
        Wend
        
    End If
    
    TabStrip1.Refresh
End Sub
Private Sub brwWebBrowser_WindowClosing(Index As Integer, ByVal IsChildWindow As Boolean, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub cboAddress_Change()
    'This will cause an Auto Complte Effect
    On Error GoTo cboAddress_Change_Error:
    Dim i As Long, j As Long, Pos As Integer
    Dim strPartial As String, strTotal As String
    
    If Not bBackSpace Then
        With cboAddress
            strPartial = .Text
            i = SendMessage(.hWnd, CB_FINDSTRING, -1, ByVal strPartial)
            
            If i <> CB_ERR Then
                strTotal = .List(i)
                j = Len(strTotal) - Len(strPartial)
                
                If j <> 0 Then
                    .SelText = Right$(strTotal, j)
                    .SelStart = Len(strPartial)
                    .SelLength = j
                Else
                End If
            Else
            End If
            
        End With
    End If
    Exit Sub
    
cboAddress_Change_Error:
    ShowErrorMessageBox ("cboAddress_Change")
    Dim strMessage
    strMessage = Err.Description
    strMessage = Err.Source
    
End Sub
Private Sub cboAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    'Still cant see why these wont work. Maybe its because of the Menus
    If Shift = vbCtrlMask And KeyCode <> vbKeyControl Then
        With cboAddress
            Select Case KeyCode
            Case vbKeyV  'PAST
                .SelText = Clipboard.GetText(vbCFText)
                KeyCode = 0
            Case vbKeyX  'CUT
                Clipboard.SetText cboAddress.SelText
                cboAddress.SelText = ""
                KeyCode = 0
            Case vbKeyC  'COPY
                Clipboard.SetText cboAddress.SelText
                KeyCode = 0
            Case vbKeyZ  'UNDO
                KeyCode = 0
            Case vbKeyY  'REDO
                KeyCode = 0
            End Select
        End With
    End If
End Sub


Private Sub cboengine_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdsearch_Click
    End If
End Sub

Private Sub CboTranslate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call TranslatePage
    End If
End Sub

Private Sub cmdclose_Click()
    Call DeleteTab
End Sub

Private Sub cmdcloseall_Click()
    Call DeleteAllTabs
End Sub

Private Sub cmdCloseFavorites_Click()
    'Close the Favorites Side Bar
    On Error GoTo cmdCloseFavorites_Click_Error:
    ViewingFavorites = False
    ViewingHistory = False
    TreeView1.Visible = False
    treeHistory.Visible = False
    FavoritesCoolBar.Visible = False
    picSplitter.Visible = False
    TabStrip1.ZOrder (0)
    TabStrip1.Refresh
    Form_Resize
    
    Exit Sub
    
cmdCloseFavorites_Click_Error:
    ShowErrorMessageBox ("DocumentComplete")
    
End Sub

Private Sub cmdsearch_Click()
    
    If cboEngine.Text = "Select a Search Engine" Then
        MsgBox "Please select a search engine to search", vbInformation, "Select Engine"
    End If
    If txtsearch.Text = "" Then
        MsgBox "Please enter at least 1 word to search for", vbInformation, "Enter Word"
        'GoTo Skip
    End If
    Call CrackSearch
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Handle the F11 key to load full screen
    If KeyCode = 122 Then FullScreen
    
    'HANDLE CONTROL KEYS
    If Shift = vbCtrlMask And KeyCode <> vbKeyControl Then
        Select Case KeyCode
        Case vbKeyN     'New Tab
            KeyCode = 0
            Call NewTab(Me, "", -1)
        Case vbKeyT     'New Tab
            KeyCode = 0
            Call NewTab(Me, "", -1)
        Case vbKeyW     'Delete Current Tab
            KeyCode = 0
            Call DeleteTab
        Case vbKeyD     'Delete All Tabs
            KeyCode = 0
            Call DeleteAllTabs
        End Select
    End If
End Sub

Private Sub Form_Load()
    Dim StrmakeChk As String
    Dim lIndex As Long
    
    On Error Resume Next
    '//////////////////////////////////////////////////////////////
    'Setup the System Tray icon
    'NOTE: icons must be 16 colors MAX
    Set gSysTray = New clsSysTray
    Set gSysTray.SourceWindow = Me
    gSysTray.Icon = imlTrayIcons.ListImages("Underground").Picture
    gSysTray.ChangeToolTip (PROGRAM_NAME)
    
    '//////////////////////////////////////////////////////////////
    'Initialize splitter  10.24.2000  JRB
    Set m_cSplit = New cSplitter
    m_cSplit.Initialise picSplitter, Me
    
    '//////////////////////////////////////////////////////////////
    'Initialize program data
    CurTab_Index = 1
    MaxTab_Index = 1
    
    TabStrip1.tabs.Item(CurTab_Index).Tag = 0
    
    StatusBar.Panels(2).Width = 1200
    StatusBar.Panels(3).Width = 1600
    StatusBar.Panels(1).Width = Me.ScaleWidth - (StatusBar.Panels(2).Width + StatusBar.Panels(3).Width)
    
    bBackSpace = False
    bDeleteKey = False
    gbFullScreen = False
    gbMaximized = False
    FavoritesCoolBar.Visible = False
    picSplitter.Visible = False
    
    Call SetProgPath
    Call GetFormInfo
    Call GetOptions
    
    Call Form_Resize
    cmdcloseall.Left = 9900
    cmdclose.Left = 9600
    
    CurTab_Index = 1
    TabStrip1.tabs(CurTab_Index).Selected = True
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Visible = True
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ZOrder (0)
    TabStrip1.SetFocus
    'Load the Various Information
    lstFreedBrowserList.AddItem ("0")
    Call GetTypedURLs
    cboAddress.Text = brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL
    Call GetHistory_txt
    Call LoadTranslation(CboTranslate)
    Call LoadSrcEngines
    Call GetScriptingInfo
    ' create the Application menu:
    Call pBuildMenu
    ' Make the toolbar:
    tbrMenu.CreateFromMenu m_cMenu
    Call setupToolbar
    StatusBar.Panels(4).Text = m_cILMenu.ItemKey(6) & "Underground Zone"
    ViewingFavorites = False
    ViewingHistory = False
    TreeView1.Visible = False
    treeHistory.Visible = False
    FavoritesCoolBar.Visible = False
    picSplitter.Visible = False
    TabStrip1.ZOrder (0)
    gbFullScreen = True
    'Load the Checked Items in the Menu
    'Caused a lot of Stress, Until i Saw the Documentation
    'lIndex = m_cMenu.IndexForKey("MenuKey") < Nice piece of Information
    Call LoadChkMenus
    PopupKill = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "PopupKilla")
    If PopupKill = True Then
        Call frmpopupkilla.LoadPopupList(lstPopups, App.path & "/popups.dat")
        lIndex = m_cMenu.IndexForKey("mnupopup")
        m_cMenu.Checked(lIndex) = True
    Else
        m_cMenu.Checked(lIndex) = False
    End If
    
    QuickDisable = GetDword(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "QuickDisable")
    lIndex = m_cMenu.IndexForKey("mnuquick")
    If QuickDisable = True Then
        m_cMenu.Checked(lIndex) = True
    Else
        m_cMenu.Checked(lIndex) = False
    End If
    'This will allow me to Open Associated Files in the Same
    'Windows under a differnt Tab, as opposed to a new Instance
    'Alltogather.
    AttachMessage Me, Me.hWnd, WM_COPYDATA
    
    ' Tell the startup module to tag this window:
    TagWindow Me.hWnd
    If Command <> "" Then
        ParseCommand Command
    Else
        If StartingAddress = "" Then
            
            Select Case gStartPage
            Case 0 ' Start at Home Page
                brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoHome
            Case 1 ' Start with Saved URLs
                Call GetSavedTabURLs
            Case 2 ' Start on Blank page
                brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate BLANK_URL
            End Select
            
        Else
            brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate StartingAddress
        End If
    End If
    Me.Refresh
    'Still dont want to use a Splash Screen
    'Makes the App look slow.
    frmBrowser.Show
End Sub
Private Sub cboAddress_Click()
    On Error GoTo cboAddress_Click_Error:
    'Handle selection of URL from history
    ClickHistory = True
    If Not UnLoading Then
        If cboAddress.Text <> "" Then
            brwWebBrowser(TabStrip1.tabs.Item(CurTab_Index).Tag).Navigate cboAddress.Text
            brwWebBrowser(TabStrip1.tabs.Item(CurTab_Index).Tag).SetFocus
        Else
            TabStrip1.SetFocus
            brwWebBrowser(TabStrip1.tabs.Item(CurTab_Index).Tag).SetFocus
        End If
    End If
    Exit Sub
cboAddress_Click_Error:
    ShowErrorMessageBox ("cboAddress_Click")
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error GoTo cboAddress_KeyPress_Error:
    'Handle the entry of a URL into the address text box
    
    On Error Resume Next
    
    If KeyAscii = vbKeyReturn Then
        AddURL = True
        
        If gNewTabAddressTyped Then
            Call NewTab(Me, cboAddress.Text, -1)
        Else
            brwWebBrowser(TabStrip1.tabs.Item(CurTab_Index).Tag).Navigate cboAddress.Text
        End If
        
    End If
    
    'Check for backspace key
    If KeyAscii = 8 Then
        bBackSpace = True
    Else
        bBackSpace = False
    End If
    
    'Check for Ctrl-V
    If KeyAscii = 22 Then
        cboAddress.Text = Clipboard.GetText
    End If
    
    Exit Sub
cboAddress_KeyPress_Error:
    ShowErrorMessageBox ("cboAddress_KeyPress_Error")
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'This Annoyed the Hell out of Me
    '--------------------------------
    '    If MsgBox("This will exit the program." & vbNewLine & "Are you sure you " & _
    '    "want to exit?", vbYesNo + vbQuestion, "Exit") = vbYes Then
    ' stop subclassing:
    DetachMessage frmBrowser, Me.hWnd, WM_COPYDATA    ' clear up mutex...
    Unload Me
    
    'Else
    '    Cancel = 1
    'End If
End Sub

Private Sub ParseCommand(ByVal sCmd As String)
    'RestoreAndActivate Me.hwnd
    sCmd = Trim$(sCmd)
    Call NewTab(Me, sCmd, -1)
End Sub
Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer.EMsgResponse)
    ' This shouldn't really be in SSUBTMR.
    ' In fact, SSUBTMR should have a BeforeMessage and AfterMessage
    ' function if it was going to be easier to use.
    ' Sometime....
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer.EMsgResponse
    ' This will tell you which message you are responding to:
    ' WM_COPYDATA, send response after we've done with it:
    ISubclass_MsgResponse = emrPostProcess
    
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tCDS As COPYDATASTRUCT
    Dim b() As Byte
    Dim sCommand As String
    
    Select Case iMsg
    Case WM_COPYDATA
        ' Copy for processing:
        CopyMemory tCDS, ByVal lParam, Len(tCDS)
        If (tCDS.cbData > 0) Then
            ReDim b(0 To tCDS.cbData - 1) As Byte
            CopyMemory b(0), ByVal tCDS.lpData, tCDS.cbData
            sCommand = StrConv(b, vbUnicode)
            
            ' We've got the info, now do it:
            ParseCommand sCommand
        Else
            ' no data.  This is only sent by the main
            ' module if it detects this window is hidden.
            ' since this can't occur in this project,
            ' this won't occur.  However, in a project
            ' where your main window can be hidden, you
            ' would make your window visible and activate
            ' it here.
        End If
        
    End Select
    
End Function

Sub Form_Resize()
    On Error GoTo Form_Resize_Error:
    '****************************
    'One of the hardest resizes I ever did to deal with
    If Me.WindowState <> vbMinimized Then gUndergroundState = Me.WindowState
    
    'Handle resizing MOST controls on the form
    If Me.WindowState = vbMinimized Then
        If gMinToSysTray Then gSysTray.MinToSysTray
    Else
        Dim FormLeft As Long
        Dim xRatio
        xRatio = (Me.ScaleWidth * 100) \ Me.ScaleWidth
        
        '////////////////////////////////////////////////////////////////
        'CALCULATE THE MAIN AREA TOP AND HEIGHT
        Dim lMainTop As Long
        Dim lT As Long
        rbrMain.RebarSize
        lT = frmBrowser.rbrMain.RebarHeight * Screen.TwipsPerPixelY
        
        lMainTop = Me.ScaleTop + lT
        Dim lMainHeight As Long
        If gbFullScreen = True Then
            lMainHeight = Me.ScaleHeight - lT - Me.StatusBar.Height
            Debug.Print lT
        Else
            lMainHeight = Me.ScaleHeight - lT  'rbrMain.RebarHeight
        End If
        
        ' /////////////////////////////////////////////////////////////////
        'Resize the splitter
        With picSplitter
            .Move .Left, lMainTop, .Width, lMainHeight
            .ZOrder (0)
        End With
        '
        '/////////////////////////////////////////////////////////////////
        'Resize the Favorites/History CoolBar
        If FavoritesCoolBar.Visible = True Then
            With FavoritesCoolBar
                .Move .Left, .Top, (picSplitter.Left - 0 * Screen.TwipsPerPixelX) + 0, .Height
            End With
            With cmdCloseFavorites
                .Move FavoritesCoolBar.Left + FavoritesCoolBar.Width - .Width - 55, .Top, .Width, .Height
            End With
        End If
        
        '/////////////////////////////////////////////////////////////////
        'Resize favorites and history no matter what - 10.24.2000 JRB
        '/////////////////////////////////////////////////////////////////
        'RESIZE THE FAVORITES TREE (wont wory about the coolba\toolbar
        FavoritesCoolBar.Top = lMainTop
        lTop = lMainTop + FavoritesCoolBar.Height
        lWidth = FavoritesCoolBar.Width
        lHeight = lMainHeight - FavoritesCoolBar.Height
        
        If lHeight > 0 Then
            TreeView1.Move Me.ScaleLeft, lTop, lWidth, lHeight
            TreeView1.Visible = True
        Else
            TreeView1.Visible = False
        End If
        '/////////////////////////////////////////////////////////////////
        'RESIZE THE HISTORY TREE (wont wory about the coolba\toolbar
        FavoritesCoolBar.Top = lMainTop
        lTop = lMainTop + FavoritesCoolBar.Height
        lWidth = FavoritesCoolBar.Width
        lHeight = lMainHeight - FavoritesCoolBar.Height
        
        If lHeight > 0 Then
            treeHistory.Move Me.ScaleLeft, lTop, lWidth, lHeight
            treeHistory.Visible = True
        Else
            treeHistory.Visible = False
        End If
        
        '/////////////////////////////////////////////////////////////////
        'RESIZE THE BROWSER TABS
        If FavoritesCoolBar.Visible = True Then
            lLeft = FavoritesCoolBar.Left + FavoritesCoolBar.Width + 20 + picSplitter.Width
        Else
            lLeft = FavoritesCoolBar.Left
        End If
        
        lTop = lMainTop
        
        If FavoritesCoolBar.Visible = True Then
            lWidth = Me.Width - FavoritesCoolBar.Width - picSplitter.Width
        Else
            lWidth = Me.Width
        End If
        
        lHeight = lMainHeight
        
        If lWidth > 0 And lHeight > 0 Then
            TabStrip1.Move lLeft, lTop, lWidth, lHeight
            TabStrip1.Visible = True
        Else
            TabStrip1.Move lLeft, lTop, lWidth, lHeight
            TabStrip1.Visible = False
        End If
        
        '/////////////////////////////////////////////////////////////////
        'RESIZE THE BROWSER WINDOWS
        MoveBrowsers
        
        '/////////////////////////////////////////////////////////////////
        'RESIZE THE BROWSER STATUS BAR
        lWidth = Me.ScaleWidth - (StatusBar.Panels(2).Width + StatusBar.Panels(3).Width + StatusBar.Panels(4).Width)
        
        If lWidth > 0 Then
            StatusBar.Panels(1).Width = lWidth
            StatusBar.Panels(1).Visible = True
        Else
            StatusBar.Panels(1).Visible = False
        End If
        
        '/////////////////////////////////////////////////////////////////
        'RESIZE and POSITION THE PROGRESS BAR
        RepositionProgressBar
        
        '////////////////////////////////////////////////////////////////
        'Save Global form properties
        gUndergroundTop = Me.Top
        gUndergroundLeft = Me.Left
        gUndergroundWidth = Me.Width
        gUndergroundHeight = Me.Height
        
    End If
    Exit Sub
Form_Resize_Error:
    ShowErrorMessageBox ("Form_Resize")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Save Form info, Current open URL's and History
    UnLoading = True
    
    gSysTray.RemoveFromSysTray
    Call SaveFormInfo
    Call SaveCurrentTabURLs
    Call SaveTypedURLs
    Dim X As Integer
    On Error Resume Next
    'Unload browsers 2 through MaxTab_Index
    For X = 2 To MaxTab_Index
        Me.brwWebBrowser(Me.TabStrip1.tabs(X).Tag).Stop
        Unload brwWebBrowser(TabStrip1.tabs(X).Tag)
    Next
    MaintainHistoryTxt
    'Unload frmAbout
    LockWindowUpdate Me.hWnd
    rbrMain.RemoveAllRebarBands
    LockWindowUpdate 0
    Unload frmSearchOptions
    Unload frmMain
    Unload frmBrowser
    Unload Me
    EndApp
    End
    
    
End Sub
Private Sub mnu_About_Click()
    
    On Error GoTo mnu_About_Click_Error:
    frmAbout.Show 1
    If cboAddress.Text = "http://www.corrupted1.f2s.com" Then
        cboAddress_KeyPress (vbKeyReturn)
    End If
    Exit Sub
    
mnu_About_Click_Error:
    ShowErrorMessageBox ("mnu_About_Click")
End Sub

Private Sub mnu_AddTab_Click()
    
    On Error GoTo mnu_AddTab_Click_Error:
    Call NewTab(Me, "", -1)
    Exit Sub
    
mnu_AddTab_Click_Error:
    ShowErrorMessageBox ("mnu_AddTab_Click")
    
End Sub

Private Sub mnu_AddToFavorites_Click()
    
    On Error GoTo mnu_AddToFavorites_Click_Error:
    Call AddToFavorites
    Call GetFavorites  'To refresh the favorites tree
    Exit Sub
    
mnu_AddToFavorites_Click_Error:
    ShowErrorMessageBox ("mnu_AddToFavorites_Click")
    
End Sub

Private Sub mnu_DeleteAllTabs_Click()
    
    On Error GoTo mnu_DeleteAllTabs_Click_Error:
    Call DeleteAllTabs
    Exit Sub
    
mnu_DeleteAllTabs_Click_Error:
    ShowErrorMessageBox ("mnu_DeleteAllTabs_Click")
    
End Sub

Private Sub mnu_DeleteTab_Click()
    
    On Error GoTo mnu_DeleteTab_Click_Error:
    Call DeleteTab
    Exit Sub
    
mnu_DeleteTab_Click_Error:
    ShowErrorMessageBox ("mnu_DeleteTab_Click")
    
End Sub

Private Sub mnu_EditCopy_Click()
    
    On Error GoTo mnu_EditCopy_Click_Error:
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
    Exit Sub
    
mnu_EditCopy_Click_Error:
    ShowErrorMessageBox ("mnu_EditCopy_Click")
    
End Sub

Private Sub mnu_EditCut_Click()
    
    On Error GoTo mnu_EditCut_Click_Error:
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT
    Exit Sub
    
mnu_EditCut_Click_Error:
    ShowErrorMessageBox ("mnu_EditCut_Click")
    
End Sub

Private Sub mnu_EditFind_Click()
    
    On Error GoTo mnu_EditFind_Click_Error:
    SetFocusOnly = True
    TabStrip1.SetFocus
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).SetFocus
    SendKeys "^f"
    Exit Sub
    
mnu_EditFind_Click_Error:
    ShowErrorMessageBox ("mnu_EditFind_Click")
    
End Sub

Private Sub mnu_EditPast_Click()
    
    On Error GoTo mnu_EditPast_Click_Error:
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT
    Exit Sub
    
mnu_EditPast_Click_Error:
    ShowErrorMessageBox ("mnu_EditPast_Click")
    
End Sub

Private Sub mnu_EditSelectAll_Click()
    
    On Error GoTo mnu_EditSelectAll_Click_Error:
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
    Exit Sub
    
mnu_EditSelectAll_Click_Error:
    ShowErrorMessageBox ("mnu_EditSelectAll_Click")
    
End Sub

Private Sub mnu_EditViewHistory_Click()
    
    '******************************
    '*** THIS IS NO LONGER USED ***
    '******************************
    'On Error GoTo NotInPath
    '    '//////////////////////////////////////////////
    '    'Save current URL history to file
    '    'Then load it into Notepad for Edit and Viewing
    '    Dim x As Integer
    '
    '    UnLoading = True
    '    Open gProgPath & "History.dat" For Output As #1
    '
    '    For x = 0 To frmBrowser.cboAddress.ListCount - 1
    '        frmBrowser.cboAddress.ListIndex = x
    '        Print #1, frmBrowser.cboAddress.text
    '    Next
    '
    '    Close #1
    '
    '    'Then load it into Notepad for Edit and Viewing
    '    x = Shell("Notepad.exe " & gProgPath & "History.dat", vbMaximizedFocus)
    '
    '    HistoryFileChanged = True
    '    Exit Sub
    '
    'NotInPath:
    '   MsgBox "ERROR, Notepad.exe was not found in your system path."
    
End Sub

Private Sub mnu_Exit_Click()
    
    On Error GoTo mnu_Exit_Click_Error:
    Unload Me
    Exit Sub
    
mnu_Exit_Click_Error:
    ShowErrorMessageBox ("mnu_Exit_Click")
End Sub

Private Sub mnu_GoBack_Click()
    
    On Error GoTo mnu_GoBack_Click_Error:
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoBack
    Exit Sub
    
mnu_GoBack_Click_Error:
    ShowErrorMessageBox ("mnu_GoBack_Click")
    
End Sub

Private Sub mnu_GoForward_Click()
    
    On Error GoTo mnu_GoForward_Click_Error:
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoForward
    Exit Sub
    
mnu_GoForward_Click_Error:
    ShowErrorMessageBox ("mnu_GoForward_Click")
    
End Sub

Private Sub mnu_GoHome_Click()
    
    On Error GoTo mnu_GoHome_Click_Error:
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoHome
    Exit Sub
    
mnu_GoHome_Click_Error:
    ShowErrorMessageBox ("mnu_GoHome_Click")
End Sub

Private Sub mnu_GoNextTab_Click()
    
    On Error GoTo mnu_GoNextTab_Click_Error:
    'Move the current browser off the form\tab
    'Move the next browser onto the form\tab
    Call MoveBrowserOffFormTab(TabStrip1.tabs(CurTab_Index).Tag)
    
    If CurTab_Index = MaxTab_Index Then
        CurTab_Index = 1
    Else
        CurTab_Index = CurTab_Index + 1
    End If
    
    Call SelectBrowserTab(CurTab_Index)
    MoveBrowsers
    Exit Sub
    
mnu_GoNextTab_Click_Error:
    ShowErrorMessageBox ("mnu_GoNextTab_Click")
End Sub

Private Sub mnu_GoPreviousTab_Click()
    
    On Error GoTo mnu_GoPreviousTab_Click_Error:
    'Move the current browser off the form\tab
    'Move the next browser onto the form\tab
    Call MoveBrowserOffFormTab(TabStrip1.tabs(CurTab_Index).Tag)
    
    If CurTab_Index = 1 Then
        CurTab_Index = MaxTab_Index
    Else
        CurTab_Index = CurTab_Index - 1
    End If
    
    Call SelectBrowserTab(CurTab_Index)
    MoveBrowsers
    Exit Sub
    
mnu_GoPreviousTab_Click_Error:
    ShowErrorMessageBox ("mnu_GoPreviousTab_Click")
    
End Sub

Private Sub mnu_GoSearch_Click()
    
    On Error GoTo mnu_GoSearch_Click_Error:
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoSearch
    Exit Sub
    
mnu_GoSearch_Click_Error:
    ShowErrorMessageBox ("mnu_GoSearch_Click")
    
End Sub

Private Sub mnu_InterNetOptions_Click()
    
    Dim RetVal
    RetVal = Shell("rundll32.exe shell32.dll,Control_RunDLL Inetcpl.cpl", vbNormalFocus)
    
End Sub

Private Sub mnu_NewBrowserBlank_Click()
    
    Call NewTab(frmBrowser, "", NEW_TAB_BLANK)
    
End Sub

Private Sub mnu_NewBrowserCurrent_Click()
    
    Call NewTab(frmBrowser, brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL, NEW_TAB_CUR_URL)
    
End Sub

Private Sub mnu_NewBrowserHome_Click()
    
    Call NewTab(frmBrowser, "", NEW_TAB_HOME)
    
End Sub

Private Sub mnu_NextTab_Click()
    
    Call mnu_GoNextTab_Click
    Call SelectBrowserTab(CurTab_Index)
    MoveBrowsers
    
End Sub

Private Sub mnu_Open_Click()
    
    frmOpen.Show 1
    
    If OpenOk = True Then
        If OpenNewTab Then
            Call NewTab(Me, OpenURL, -1)
        Else
            brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate OpenURL
        End If
    End If
    
End Sub

Private Sub mnu_OpenAllHistoryPages_Click()
    
    'SORRY NOT IMPLAMENTED YET
    
End Sub

Private Sub mnu_OpenDomainCurrentTab_Click()
    
    'Navigate current Tab\Browser to the selected URL
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate gHistoryURL
    
End Sub

Private Sub mnu_OpenDomainNewTab_Click()
    
    'Should do a New Tab here...
    Call NewTab(Me, gHistoryURL, -1)
    
End Sub

Private Sub mnu_OpenFavoriteCurrentTab_Click()
    
    'Navigate current Tab\Browser to the selected URL
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate gFavoritesURL
    
End Sub

Private Sub mnu_OpenFavoriteNewTab_Click()
    'Should do a New Tab here...
    Call NewTab(Me, gFavoritesURL, -1)
End Sub

Private Sub mnu_OpenHistoryInCurrentTab_Click()
    
    'Navigate current Tab\Browser to the selected URL
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate gHistoryURL
    
End Sub

Private Sub mnu_OpenHistoryInNewTab_Click()
    
    'Should do a New Tab here...
    Call NewTab(Me, gHistoryURL, -1)
    
End Sub

Private Sub mnu_OrganizeFavorites_Click()
    
    On Error GoTo mnu_OrganizeFavorites_Click_Error:
    Dim lpszRootFolder As String
    Dim success As Long
    Dim CSIDL As Long
    
    'open the organize folder at the path specified by the CSIDL
    CSIDL = CSIDL_FAVORITES
    
    lpszRootFolder = GetFolderPath(CSIDL)
    success = DoOrganizeFavDlg(hWnd, lpszRootFolder)
    
    Call GetFavorites  'To refresh the favorites tree
    Exit Sub
mnu_OrganizeFavorites_Click_Error:
    ShowErrorMessageBox ("mnu_OrganizeFavorites_Click")
    
End Sub

Private Sub mnu_PageSetup_Click()
    
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT
    
End Sub

Private Sub mnu_PrevTab_Click()
    
    Call mnu_GoPreviousTab_Click
    Call SelectBrowserTab(CurTab_Index)
    MoveBrowsers
    
End Sub

Private Sub mnu_Print_Click()
    
    Call PrintBrowser
    
End Sub

Private Sub mnu_Properties_Click()
    
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT
    
End Sub

Private Sub mnu_RefreshHistory_Click()
    
    '******************************
    '*** THIS IS NO LONGER USED ***
    '******************************
    'On Error GoTo mnu_RefreshHistory_Click_Error:
    '    If HistoryFileChanged Then
    '        UnLoading = True
    '
    '        'Delete all current 'Typed' URL's in the registry
    '        Call DeleteKey(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs")
    '
    '        'Load history from history file into Address list
    '        Get_History
    '
    '        cboAddress.text = brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).LocationURL
    '
    '        UnLoading = False
    '
    '        TabStrip1.SetFocus
    '        brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).SetFocus
    '
    '        HistoryFileChanged = False
    '    Else
    '        MsgBox "History file has not been changed." & vbCrLf _
    '               & "Use Edit/View History first, then Refresh History."
    '    End If
    '    Exit Sub
    '
    'mnu_RefreshHistory_Click_Error:
    '    ShowErrorMessageBox ("mnu_RefreshHistory_Click")
    
End Sub

Private Sub mnu_SaveAs_Click()
    
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
    
End Sub

Private Sub mnu_SelectAll_Click()
    
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
    
End Sub

Private Sub mnu_SurfTabOptions_Click()
    frmOptions.Show 1
End Sub

Private Sub mnu_TabPopUp_DeleteAllTabs_Click()
    
    If MsgBox("This will close all Open Sites" & vbNewLine & "Are you sure you " & _
    "want to Continue?", vbYesNo + vbQuestion, "Exit") = vbYes Then
    Call DeleteAllTabs
Else
    Exit Sub
End If
End Sub

Private Sub mnu_TabPopUp_DeleteTab_Click()
    
    Call DeleteTab
    
End Sub

Private Sub mnu_TabPopUp_DuplicateTab_Click()
    
    Call NewTab(Me, brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL, -1)
    
    
End Sub

Private Sub mnu_TabPopUp_NewTab_Click()
    
    Call NewTab(Me, "", -1)
    
End Sub

Private Sub mnu_TrayExit_Click()
    
    On Error GoTo mnu_TrayExit_Click_Error:
    Unload Me
    Exit Sub
    
mnu_TrayExit_Click_Error:
    ShowErrorMessageBox ("mnu_TrayExit_Click")
    
End Sub
Private Sub mnu_Underground_Click()
    
    gSysTray.LButtonDown
    
End Sub

Public Sub mnu_ViewFavorites_Click()
    
    lblFavoritesCoolBar.Caption = "Favorites"
    If ViewingFavorites Then
        If ViewingFavorites And Not ViewingHistory Then
            'Then close the side window where Favorites and History are
            ViewingFavorites = False
            ViewingHistory = False
            TreeView1.Visible = False
            treeHistory.Visible = False
            FavoritesCoolBar.Visible = False
            picSplitter.Visible = False
            TabStrip1.ZOrder (0)
        Else
            If TreeView1.Visible = True Then
                cmdCloseFavorites_Click
            Else
                ViewingHistory = False
                treeHistory.Visible = False
                ViewingFavorites = True
                TreeView1.Visible = True
                FavoritesCoolBar.Visible = True
                TreeView1.ZOrder (0)
                picSplitter.Visible = True
                Call GetFavorites  'To refresh the favorites tree
            End If
        End If
    Else
        ViewingHistory = False
        treeHistory.Visible = False
        ViewingFavorites = True
        TreeView1.Visible = True
        FavoritesCoolBar.Visible = True
        TreeView1.ZOrder (0)
        picSplitter.Visible = True
        Call GetFavorites  'To refresh the favorites tree
    End If
    
    Form_Resize
End Sub

Private Sub mnu_ViewRefresh_Click()
    
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Refresh
    
End Sub

Private Sub mnu_ViewRefreshAllTabs_Click()
    
    Dim X
    For X = 1 To TabStrip1.tabs.Count
        brwWebBrowser(TabStrip1.tabs(X).Tag).Refresh
    Next
    
End Sub

Private Sub mnu_ViewSource_Click()
    
    On Error GoTo mnu_ViewSource_Click_Error:
    If Len(brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Document.documentelement.innerhtml) > 0 Then
        
        frmMain.rt.Text = brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Document.documentelement.innerhtml
        frmMain.Caption = "Underground HTML Editor - [" & brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL & "]"
        frmMain.SetSyntax
        frmMain.Show
    End If
    
    Exit Sub
    
mnu_ViewSource_Click_Error:
    ShowErrorMessageBox ("mnu_ViewSource_Click")
    
End Sub

Private Sub mnu_ViewStop_Click()
    
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Stop
    
End Sub

Private Sub mnu_ViewStopAllTabs_Click()
    
    Dim X
    For X = 1 To TabStrip1.tabs.Count
        brwWebBrowser(TabStrip1.tabs(X).Tag).Stop
    Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_cSplit.MouseMove X
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    m_cSplit.MouseUp X
    
End Sub
Private Sub mnucredits_Click()
    frmCredits.Show
End Sub

Private Sub mnudetoeng_Click()
    
    PageUrl = ""
    PageUrl = brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL
    PageUrl = Right(PageUrl, Len(PageUrl) - 7)
    TransUrl = ""
    TransUrl = "http://babel.altavista.com/translate.dyn?doit=done&BabelFishFrontPage=yes&enc=utf8&bblType=url&url=http%3A%2F%2F" & PageUrl & "%2F&lp=de_en"
    Call NewTab(Me, TransUrl, -1)
    
End Sub

Private Sub mnufrtoeng_Click()
    
    PageUrl = ""
    PageUrl = brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL
    PageUrl = Right(PageUrl, Len(PageUrl) - 7)
    TransUrl = ""
    TransUrl = "http://babel.altavista.com/translate.dyn?doit=done&BabelFishFrontPage=yes&enc=utf8&bblType=url&url=http%3A%2F%2F" & PageUrl & "%2F%3Ff%3Dext&lp=fr_en"
    Debug.Print TransUrl
    Call NewTab(Me, TransUrl, -1)
    
    
End Sub
Private Sub mnupopup_Click()
    
    frmpopupkilla.Show
    
End Sub

Private Sub mnuporttoeng_Click()
    
    PageUrl = ""
    PageUrl = brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL
    PageUrl = Right(PageUrl, Len(PageUrl) - 7)
    TransUrl = ""
    TransUrl = "http://babel.altavista.com/translate.dyn?doit=done&BabelFishFrontPage=yes&enc=utf8&bblType=url&url=http%3A%2F%2F" & PageUrl & "%2F%3Ff%3Dext&lp=pt_en"
    Debug.Print TransUrl
    Call NewTab(Me, TransUrl, -1)
    
    
End Sub

Private Sub mnupreview_Click()
    
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
    
End Sub

Private Sub mnuquick_Click()
    
    
    
End Sub

Private Sub mnurussiantoeng_Click()
    
    PageUrl = ""
    PageUrl = brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL
    PageUrl = Right(PageUrl, Len(PageUrl) - 7)
    TransUrl = ""
    TransUrl = "http://babel.altavista.com/translate.dyn?doit=done&BabelFishFrontPage=yes&enc=utf8&bblType=url&url=http%3A%2F%2F" & PageUrl & "%2F%3Ff%3Dext&lp=ru_en"
    Debug.Print TransUrl
    Call NewTab(Me, TransUrl, -1)
    
End Sub
Private Sub mnusrcoptions_Click()
    frmSearchOptions.Show
    
End Sub

Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    m_cSplit.MouseDown X
End Sub

Private Sub TabStrip1_Click()
    
    TabStrip1.Refresh
End Sub

Private Sub TabStrip1_GotFocus()
    
    '*****************************************
    On Error GoTo TabStrip1_GotFocus_Error:
    'Clear Progress Panel so it gets reset.
    StatusBar.Panels(2).Text = ""
    ProgressBar.Value = 0
    
    'Move current tab off from
    Call MoveBrowserOffFormTab(TabStrip1.tabs(CurTab_Index).Tag)
    
    'Set up the new current tab
    CurTab_Index = TabStrip1.SelectedItem.Index
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Visible = True
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ZOrder 0
    
    Me.Caption = TabStrip1.tabs(CurTab_Index).Caption & " - " & PROGRAM_NAME
    cboAddress.Text = brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL
    If Left(cboAddress.Text, 5) = "about" Then cboAddress.Text = ""
    
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).SetFocus
    Form_Resize
    
TabStrip1_GotFocus_Error:
    ShowErrorMessageBox ("TabStrip1_GotFocus")
    
End Sub

Private Sub TabStrip1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'HANDLE CONTROL KEYS
    If Shift = vbCtrlMask And KeyCode <> vbKeyControl Then
        Select Case KeyCode
        Case vbKeyN     'New Tab
            Call NewTab(Me, "", -1)
            KeyCode = 0
        Case vbKeyW     'Delete Current Tab
            Call DeleteTab
            KeyCode = 0
        End Select
    End If
    
End Sub

Private Sub TabStrip1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'If Button = 2 Then PopupMenu Me.mnu_TabPopUp
    
End Sub
'Private Sub gSysTray_RButtonUP()
'
'    PopupMenu Me.mnu_TrayPopup
'
'End Sub

Public Sub NewTab(b As Object, URL As String, intOption As Integer)
    
    On Error GoTo NewTab_Error:
    
    Dim X
    'Increment Tab indexes
    MaxTab_Index = MaxTab_Index + 1
    CurTab_Index = MaxTab_Index
    
    'Add new tab and set properties
    TabStrip1.tabs.Add
    TabStrip1.tabs.Item(CurTab_Index).Caption = "Blank"
    TabStrip1.tabs.Item(CurTab_Index).Selected = True
    
    'Check FreedBrowserList for an index number
    If lstFreedBrowserList.ListCount = 1 Then
        'List is empty, Increment WebBrowser indexes
        TabStrip1.tabs(CurTab_Index).Tag = MaxTab_Index - 1
    Else
        'List is NOT empty, get first index
        TabStrip1.tabs(CurTab_Index).Tag = lstFreedBrowserList.List(1)
        lstFreedBrowserList.RemoveItem (1)
    End If
    
    'Load new browser, and enable it
    Load brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag)
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Visible = True
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ZOrder 0
    
    '3 Options (Blank, Home or Current)
    If URL = "" Then
        
        If intOption = -99 Then Exit Sub
        
        If intOption = -1 Then
            
            Select Case Int(frmOptions.optDefaultNewButton(gDefaultNewButton).Tag)
                
            Case 1 'NEW_TAB_HOME
                brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoHome
                
            Case 2 'NEW_TAB_BLANK
                brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate BLANK_URL
                
            Case 3 'NEW_TAB_CUR_URL
                brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoHome
                
            End Select
            
        Else
            
            Select Case intOption
                
            Case 1 'NEW_TAB_HOME
                brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoHome
                
            Case 2 'NEW_TAB_BLANK
                brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate BLANK_URL
                
            Case 3 'NEW_TAB_CUR_URL
                brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL
                
            Case 5 'NEW_TAB_SEARCH
                brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoSearch
                
            End Select
            
        End If
        
    Else
        
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate URL
        
    End If
    
    Call SelectBrowserTab(CurTab_Index)
    TabStrip1.SetFocus
    MoveBrowsers
    
    Exit Sub
    
NewTab_Error:
    ShowErrorMessageBox ("NewTab")
    
End Sub

Private Sub DeleteTab()
    On Error GoTo DeleteTab_Error:
    Dim X
    
    'Stop the current tab\browser before deleteing it
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Stop
    
    If TabStrip1.tabs.Count > 1 Then
        
        Deleting = True
        
        'Adjust FreedBrowserList
        If TabStrip1.tabs(CurTab_Index).Tag = 0 Then
            brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate brwWebBrowser(TabStrip1.tabs(CurTab_Index + 1).Tag).LocationURL
            Unload brwWebBrowser(TabStrip1.tabs(CurTab_Index + 1).Tag)
            lstFreedBrowserList.AddItem (TabStrip1.tabs(CurTab_Index + 1).Tag)
            TabStrip1.tabs.Remove (CurTab_Index + 1)
        Else
            Unload brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag)
            lstFreedBrowserList.AddItem (TabStrip1.tabs(CurTab_Index).Tag)
            TabStrip1.tabs.Remove (CurTab_Index)
        End If
        
        'Decrement Tab indexes
        If CurTab_Index = MaxTab_Index Then
            MaxTab_Index = MaxTab_Index - 1
            CurTab_Index = MaxTab_Index
        Else
            MaxTab_Index = MaxTab_Index - 1
        End If
        
        'Set the new current tab
        TabStrip1.tabs.Item(CurTab_Index).Selected = True
        
        Deleting = False
        
    End If
    Call SelectBrowserTab(CurTab_Index)
    MoveBrowsers
    Exit Sub
    
DeleteTab_Error:
    ShowErrorMessageBox ("DeleteTab")
End Sub

Sub DeleteAllTabs()
    On Error GoTo DeleteAllTabs_Error:
    Dim TabCount, X As Integer
    
    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Stop
    
    If TabStrip1.tabs.Count > 1 Then
        
        'Delete tabs and browsers 2 through N
        X = 1
        TabCount = TabStrip1.tabs.Count
        
        While TabCount >= X
            TabStrip1.tabs(TabCount).Selected = True
            Call DeleteTab
            TabCount = TabCount - 1
        Wend
        
        'Adjust FreedBrowserList
        lstFreedBrowserList.Clear
        lstFreedBrowserList.AddItem ("0")
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Refresh
    End If
    TabStrip1.SetFocus
    Exit Sub
    
DeleteAllTabs_Error:
    ShowErrorMessageBox ("DeleteAllTabs")
    
End Sub

Private Sub tbrMenu_ButtonClick(ByVal lButton As Long)
    '    Dim lIndex As Long
    '    lIndex = m_cMenu.IndexForKey("Favorites")
    '    m_cMenu.ShowPopupMenuAtIndex 2500, 500, lIndex:=lIndex
End Sub

Private Sub tbrTools_ButtonClick(ByVal lButton As Long)
    On Error Resume Next
    Dim X As Integer
    Dim Y As Long
    Dim lIndex As Long
    Select Case tbrTools.ButtonKey(lButton)
    Case "Back"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoBack
        
    Case "Forward"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoForward
        
    Case "Refresh"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Refresh
        
    Case "Home"
        If gNewTabHome Then
            Call NewTab(Me, "", NEW_TAB_HOME)
        Else
            brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoHome
        End If
        
    Case "Search"
        Call cmdsearch_Click
        
    Case "Stop"
        StatusBar.Panels(2).Text = ""
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Stop
        Me.Caption = brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationName & " - " & PROGRAM_NAME
        gSysTray.ChangeToolTip (Me.Caption)
        
    Case "NewTab"
        Call NewTab(Me, "", -1)
        
    Case "DeleteTab"
        Call DeleteTab
        
    Case "DeleteAllTabs"
        Call DeleteAllTabs
        
    Case "Print"
        Call PrintBrowser
        
    Case "Favorites"
        Call mnu_ViewFavorites_Click
        
    Case "FullScreen"
        Call FullScreen
        
    Case "Options"
        Call mnu_SurfTabOptions_Click
        
    Case "History"
        ViewHistory
    Case "Media"
        lIndex = m_cMenu.IndexForKey("mnuimages")
        m_cMenu.ShowPopupMenuAtIndex 9000, 700, lIndex:=lIndex
    Case "Security"
        lIndex = m_cMenu.IndexForKey("mnuscripting")
        m_cMenu.ShowPopupMenuAtIndex 8500, 700, lIndex:=lIndex
    Case "Editor"
        frmMain.Show
    Case "Editor"
        frmMain.Show
    Case "Images"
        
    End Select
End Sub


Private Sub treeFavorites_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
        bRightMouse = True
    Else
        bRightMouse = False
    End If
    
End Sub

Private Sub treeFavorites_NodeClick(ByVal Node As MSComctlLib.Node)
    
    On Error GoTo treeFavorites_NodeClick_Error:
    
    'SAVE NODE TO GLOBAL VARIABLE
    Set Itm = Node
    gFavoritesURL = Itm.Tag
    
    'LEFT CLICK ON HISTORY ENTRY
    If bRightMouse And Right(Node.Key, 4) = "_URL" Then
        'PopupMenu Me.mnu_FavoritesPopUp
        bRightMouse = False
    End If
    'RIGHT CLICK ON HISTORY ENTRY
    If Right(Node.Key, 4) = "_URL" Then
        If gNewTabFavorites Then
            Call mnu_OpenFavoriteNewTab_Click
        Else
            Call mnu_OpenFavoriteCurrentTab_Click
        End If
    End If
    
    'LEFT CLICK ON HISTORY DOMAIN FOLDER
    If bRightMouse And Right(Node.Key, 4) <> "_URL" Then
        'PopupMenu Me.mnu_FavoritesFolderPopUp
        bRightMouse = False
    End If
    
    Exit Sub
    
treeFavorites_NodeClick_Error:
    ShowErrorMessageBox ("treeFavorites_NodeClick")
    
End Sub

Public Sub ViewHistory()
    
    With frmBrowser
        .lblFavoritesCoolBar.Caption = "History"
        If ViewingHistory Then
            If ViewingHistory And Not ViewingFavorites Then
                'Then close the side window where Favorites is
                ViewingHistory = False
                ViewingFavorites = False
                .treeHistory.Visible = False
                .treeHistory.Visible = False
                FavoritesCoolBar.Visible = False
                picSplitter.Visible = False
                TabStrip1.ZOrder (0)
            Else
                If treeHistory.Visible = True Then
                    cmdCloseFavorites_Click
                Else
                    .TreeView1.Visible = False
                    ViewingFavorites = False
                    ViewingHistory = True
                    .treeHistory.Visible = True
                    FavoritesCoolBar.Visible = True
                    .treeHistory.ZOrder (0)
                    picSplitter.Visible = True
                End If
            End If
        Else
            .TreeView1.Visible = False
            ViewingFavorites = False
            ViewingHistory = True
            .treeHistory.Visible = True
            FavoritesCoolBar.Visible = True
            .treeHistory.ZOrder (0)
            picSplitter.Visible = True
        End If
        
        Form_Resize
        
    End With
    
End Sub

Private Sub timTimer_Timer()
    
End Sub

Private Sub treeHistory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
        bRightMouse = True
    Else
        bRightMouse = False
    End If
    
End Sub

Private Sub treeHistory_NodeClick(ByVal Node As MSComctlLib.Node)
    
    On Error GoTo treeHistory_NodeClick_Error:
    'SAVE NODE TO GLOBAL VARIABLE
    Set Itm = Node
    gHistoryURL = Itm.Tag
    
    'LEFT CLICK ON HISTORY ENTRY
    If bRightMouse And Right(Node.Key, 4) = "_URL" Then
        'PopupMenu Me.mnu_HistoryPopUp
        bRightMouse = False
    End If
    'RIGHT CLICK ON HISTORY ENTRY
    If Right(Node.Key, 4) = "_URL" Then
        If gNewTabHistory Then
            Call mnu_OpenHistoryInNewTab_Click
        Else
            Call mnu_OpenHistoryInCurrentTab_Click
        End If
    End If
    
    'LEFT CLICK ON HISTORY DOMAIN FOLDER
    If bRightMouse And Right(Node.Key, 7) = "_Domain" Then
        'PopupMenu Me.mnu_HistoryFolderPopUp
        bRightMouse = False
    End If
    
    Exit Sub
    
treeHistory_NodeClick_Error:
    ShowErrorMessageBox ("treeHistory_NodeClick")
    
End Sub

Private Sub m_cSplit_DoSplit(bSplit As Boolean)
    
    ' Can cancell split here
    
End Sub

Private Sub m_cSplit_SplitComplete()
    Form_Resize
    
End Sub

Function ConvertSpaces(strdata As String) As String
    
    'Special Character Conversions Coming Soon.
    Dim i As Integer
    Dim ch As String
    Dim sQ As String
    For i = 1 To Len(strdata)
        ch = Mid(strdata, i, 1)
        If ch = " " Then
            sQ = sQ & "+"
        Else
            sQ = sQ & ch
        End If
    Next
    ConvertSpaces = sQ
End Function
Private Sub GetScriptingInfo()
    
    Dim strScripting As String
    Dim strCookies As String
    Dim strSession As String
    Dim strActiveX As String
    Dim strImages As String
    Dim strSounds As String
    Dim strAnimations As String
    Dim strVideos As String
    
    strScripting = ""
    strScripting = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1400")
    strCookies = ""
    strCookies = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A02")
    strSession = ""
    strSession = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A03")
    strActiveX = ""
    strActiveX = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1200")
    strImages = ""
    strImages = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Display Inline Images")
    strSounds = ""
    strSounds = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Play_Background_Sounds")
    strAnimations = ""
    strAnimations = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Play_Animations")
    strVideos = ""
    strVideos = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Display Inline Videos")
    
    
End Sub

Public Sub LoadSrcEngines()
    
    Set LoadEngines = New clsLoadSearchEngines
    LoadEngines.RefreshCBEngines cboEngine
    cboEngine.Text = "Select Search Engine to Use"
    
End Sub

Private Sub CrackSearch()
    
    Dim strFinal As String
    StrSearch = ""
    StrSearch = ConvertSpaces(txtsearch.Text)
    StrLocation = ""
    StrLocation = LoadEngines.GetURL(cboEngine.Text)
    
    strFinal = replace(StrLocation, "SearchString", StrSearch, 1, -1, vbTextCompare)
    Debug.Print "Final Thing is " & strFinal
    
    If chknewtab.Value = vbChecked Then
        Call NewTab(Me, strFinal, -1)
    Else
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate strFinal
    End If
    
End Sub
Public Sub RegisterasBrowser()
    
    frmBrowser.brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).RegisterasBrowser = True
    
End Sub
Public Sub LoadTranslation(CB As ComboBox)
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = OpenDatabase(App.path & "\SearchEngines.udb")
    Set rs = db.OpenRecordset("Translate")
    
    CB.Clear
    With rs
        '.Requery
        If .BOF And .EOF Then
        Else
            .MoveFirst
            Do Until .EOF
                CB.AddItem !Title
                .MoveNext
            Loop
        End If
    End With
    
End Sub
Public Function GetTransUrl(Title As String) As String
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = OpenDatabase(App.path & "\SearchEngines.udb")
    Set rs = db.OpenRecordset("Translate")
    GetTransUrl = ""
    With rs
        If .BOF And .EOF Then
            Exit Function
        Else
            .MoveFirst
            Do Until .EOF
                If !Title = Title Then
                    GetTransUrl = !URL
                    Exit Function
                End If
                .MoveNext
            Loop
        End If
    End With
    
End Function
Sub TranslatePage()
    
    Dim StrTransLocation As String
    Dim strPageUrl As String
    Dim strTransFinal As String
    strPageUrl = ""
    strPageUrl = brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL
    strPageUrl = Right(strPageUrl, Len(strPageUrl) - 7)
    
    StrTransLocation = ""
    StrTransLocation = GetTransUrl(CboTranslate.Text)
    Debug.Print "TransLocal = "; StrTransLocation
    
    strTransFinal = replace(StrTransLocation, "PageUrl", strPageUrl, 1, -1, vbTextCompare)
    Debug.Print "Final Thing is " & strTransFinal
    
    Call NewTab(Me, strTransFinal, -1)
    
End Sub
Private Sub RefreshTabs()
    
    Dim TabCount, X As Integer
    
    If TabStrip1.tabs.Count > 1 Then
        X = 1
        TabCount = TabStrip1.tabs.Count
        While TabCount >= X
            brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Refresh
            TabCount = TabCount - 1
        Wend
    End If
    
End Sub
'Change out these
Private Sub TreeView1_ItemClick(hItem As Long, RightButton As Boolean)
    Dim UrlFav$
    If Len(TreeView1.ItemKey(hItem)) > 0 Then
        If InStr(TreeView1.ItemKey(hItem), "Folder") = 0 Then
            UrlFav$ = TreeView1.ItemKey(hItem)
            UrlFav$ = Left$(UrlFav$, InStr(UrlFav$, " ") - 1)
            brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate UrlFav$
            cboAddress.Text = TreeView1.ItemKey(hItem)
        End If
    End If
End Sub
Private Sub TreeView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyF5 Then
        '         TreeView1.Nodes.Clear
        TreeView1.Clear
        TreeView1.Refresh
        
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
            
            'get the files
            Call SearchForFilesArray(FP)
            TreeView1.ItemExpanded("") = True
        Else
            
            MsgBox " Could not locate favorites folder! " & _
            "This program requires Microsoft's Internet " & _
            "Explorer to be installed. Program will shutdown now!", _
            vbCritical + vbOKOnly, "FavMenu Error"
            End
        End If
    End If
End Sub
Public Sub pBuildMenu()
    On Error Resume Next
    'Dim iP(0 To 6) As Long
    Set m_cMenu = New cPopupMenu
    
    'The menu settings
    With m_cMenu
        
        .ImageList = m_cILMenu.hIml
        .hwndOwner = Me.hWnd
        .GradientHighlight = True
        ' File menu:
        iP(0) = .AddItem("&File", , , , , , , "mnu_File")
        iP(1) = .AddItem("&New Browser Tab", , , iP(0), , , , "mnu_NewBrowser")
        iP(2) = .AddItem("&Current", , , iP(1), , , , "mnu_NewBrowserCurrent")
        iP(2) = .AddItem("&Home", , , iP(1), , , , "mnu_NewBrowserHome")
        iP(2) = .AddItem("&Blank", , , iP(1), , , , "mnu_NewBrowserBlank")
        iP(1) = .AddItem("&Open..." & vbTab & "Ctrl+O", , , iP(0), , , , "mnu_Open")
        iP(1) = .AddItem("&Save As...", , , iP(0), , , , "mnu_SaveAs")
        iP(1) = .AddItem("-", , , iP(0), , , , "sep1")
        iP(1) = .AddItem("Page Set&up", , , iP(0), , , , "mnu_PageSetup")
        iP(1) = .AddItem("&Print", , , iP(0), , , , "mnu_Print")
        iP(1) = .AddItem("Print Preview", , , iP(0), , , , "mnupreview")
        iP(1) = .AddItem("-", , , iP(0), , , , "line3")
        iP(1) = .AddItem("Properties", , , iP(0), , , , "mnu_Properties")
        iP(1) = .AddItem("&Work Offline", , , iP(0), , , , "mnuworkoffline")
        iP(1) = .AddItem("-", , , iP(0), , , , "sep2")
        iP(1) = .AddItem("E&xit", , , iP(0), , , , "mnu_Exit")
        
        ' Edit menu
        iP(0) = .AddItem("&Edit", , , , , , , "mnu_Edit")
        iP(1) = .AddItem("C&ut" & vbTab & "Ctrl+X", , , iP(0), , , , "mnu_EditCut")
        iP(1) = .AddItem("&Copy" & vbTab & "Ctrl+C", , , iP(0), , , , "mnu_EditCopy")
        iP(1) = .AddItem("&Paste" & vbTab & "Ctrl+V", , , iP(0), , , , "mnu_EditPast")
        iP(1) = .AddItem("-", , , iP(0), , , , "line2")
        iP(1) = .AddItem("Select &All" & vbTab & "Ctrl+A", , , iP(0), , , , "mnu_EditSelectAll")
        iP(1) = .AddItem("-", , , iP(0), , , , "line15")
        iP(1) = .AddItem("&Find on this page" & vbTab & "Ctrl+F", , , iP(0), , , , "mnu_EditFind")
        'iP(1) = .AddItem("", , , iP(0), , , , "")
        
        ' View menu
        iP(0) = .AddItem("&View", , , , , , , "mnu_View")
        iP(1) = .AddItem("Tool Bar", , , iP(0), , True, , "mnutoolbar")
        iP(1) = .AddItem("Address Bar", , , iP(0), , True, , "mnuaddressbar")
        iP(1) = .AddItem("Translation Bar", , , iP(0), , True, , "mnutranslatebar")
        iP(1) = .AddItem("Search Bar", , , iP(0), , True, , "mnusearchbar")
        iP(1) = .AddItem("-", , , iP(0), , , , "sep11")
        iP(1) = .AddItem("&Goto", , , iP(0), , , , "mnu_Go")
        iP(2) = .AddItem("&Back", , , iP(1), , , , "mnu_GoBack")
        iP(2) = .AddItem("&Forward", , , iP(1), , , , "mnu_GoForward")
        iP(2) = .AddItem("-", , , iP(1), , , , "sep4")
        iP(2) = .AddItem("&Home", , , iP(1), , , , "mnu_GoHome")
        iP(2) = .AddItem("&Search", , , iP(1), , , , "mnu_GoSearch")
        iP(2) = .AddItem("-", , , iP(1), , , , "sep5")
        iP(2) = .AddItem("&Previous Tab", , , iP(1), , , , "mnu_GoPreviousTab")
        iP(2) = .AddItem("&Next Tab", , , iP(1), , , , "mnu_GoNextTab")
        iP(1) = .AddItem("&Stop" & vbTab & "Esc", , , iP(0), , , , "mnu_ViewStop")
        iP(1) = .AddItem("Stop All", , , iP(0), , , , "mnu_ViewStopAllTabs")
        iP(1) = .AddItem("&Refresh" & vbTab & "F5", , , iP(0), , , , "mnu_ViewRefresh")
        iP(1) = .AddItem("Refresh All Tabs", , , iP(0), , , , "mnu_ViewRefreshAllTabs")
        iP(1) = .AddItem("-", , , iP(0), , , , "line1")
        iP(1) = .AddItem("Text Size", , , iP(0), , , , "mnutextsize")
        iP(2) = .AddItem("Largest", , , iP(1), , , , "mnuLargest")
        iP(2) = .AddItem("Large", , , iP(1), , , , "mnuLarge")
        iP(2) = .AddItem("Medium", , , iP(1), , , , "mnuMedium")
        iP(2) = .AddItem("Smaller", , , iP(1), , , , "mnuSmaller")
        iP(2) = .AddItem("Smallest", , , iP(1), , , , "mnuSmallest")
        iP(1) = .AddItem("-", , , iP(0), , , , "line12")
        iP(1) = .AddItem("&Source", , , iP(0), , , , "mnu_ViewSource")
        iP(1) = .AddItem("-", , , iP(0), , , , "line7")
        iP(1) = .AddItem("Full Screen" & vbTab & "F11", , , iP(0), , , , "Fullscreen")
        
        'Favorites Menu
        iP(0) = .AddItem("&Favorites", , , , , , , "Favorites")
        iP(1) = .AddItem("&Add To Favorites", , , iP(0), , , , "AddFavorites")
        iP(1) = .AddItem("&Organize Favorites", , , iP(0), , , , "OrganizeFavorites")
        iP(1) = .AddItem("-", , , iP(0), , , , "line12")
        'Setup di Favorites
        favpath = GetFolderPath(CSIDL_FAVORITES)
        frmBrowser.TreeView1.Clear
        frmBrowser.TreeView1.Refresh
        Call LoadTreeView("Internet Favorites", True, True)
        If Len(favpath) > 0 Then
            With FP
                .sFileRoot = favpath
                .sFileNameExt = "*.url"
                .bRecurse = True
            End With
            frmBrowser.TreeView1.ItemExpanded(a) = True
        End If
        'IMPORTANT Shows the (Links)
        Call SearchForFilesArray(FP)
        'Dis is what set it to the current Menu
        Call GetFavorites
        '-------------------------------------
        'Tools / Optoins Menu
        iP(0) = .AddItem("T&ools", , , , , , , "mnu_Options")
        iP(1) = .AddItem("Favorites", , , iP(0), , , , "mnuFav")
        iP(2) = .AddItem("Export Favorites", , , iP(1), , , , "mnuexport")
        iP(2) = .AddItem("Import Bookmarks", , , iP(1), , , , "mnuimport")
        iP(2) = .AddItem("Syncronise Offline Favorites", , , iP(1), , , , "mnusync")
        iP(1) = .AddItem("Mail and News", , , iP(0), , , , "mnu_Readmail")
        iP(2) = .AddItem("New Message", , , iP(1), , , , "mnu_newMessage")
        iP(2) = .AddItem("Send Link By E-mail", , , iP(1), , , , "mnu_Sendlink")
        iP(2) = .AddItem("Send Page By E-Mail", , , iP(1), , , , "mnu_Sendpage")
        iP(2) = .AddItem("-", , , iP(1), , , , "sep11")
        iP(2) = .AddItem("Read News ", , , iP(1), , , , "mnuReadnews")
        iP(2) = .AddItem("-", , , iP(1), , , , "sep12")
        iP(2) = .AddItem("Hotmail", , , iP(1), , , , "mnuhotmail")
        iP(2) = .AddItem("Outlook Express", , , iP(1), , , , "mnu_outlook")
        iP(1) = .AddItem("&Media Loading", , , iP(0), , , , "mnumedia")
        iP(2) = .AddItem("Load Images", , , iP(1), , , , "mnuimages")
        iP(2) = .AddItem("Load Videos", , , iP(1), , , , "mnuvideos")
        iP(2) = .AddItem("Play Sounds", , , iP(1), , , , "mnusounds")
        iP(2) = .AddItem("Play Animations", , , iP(1), , , , "mnuanimation")
        iP(1) = .AddItem("&TranslatePage", , , iP(0), , , , "mnutranslate")
        iP(1) = .AddItem("-", , , iP(0), , , , "sep6")
        iP(1) = .AddItem("Quick Disable", , , iP(0), , , , "mnuquick")
        iP(1) = .AddItem("-", , , iP(0), , , , "sep7")
        iP(1) = .AddItem("PopupKilla", , , iP(0), , , , "mnupopup")
        iP(1) = .AddItem("-", , , iP(0), , , , "sep8")
        iP(1) = .AddItem("Search Options", , , iP(0), , , , "mnusrcoptions")
        iP(1) = .AddItem("-", , , iP(0), , , , "sep9")
        iP(1) = .AddItem("Underground Options", , , iP(0), , , , "mnu_SurfTabOptions")
        iP(1) = .AddItem("Internet Options", , , iP(0), , , , "mnu_InterNetOptions")
        
        'Tabs Menu
        iP(0) = .AddItem("&Tabs", , , , , , , "mnu_Tabs")
        iP(1) = .AddItem("&Next Tab" & vbTab & "Alt+N", , , iP(0), , , , "mnu_NextTab")
        iP(1) = .AddItem("&Previous Tab" & vbTab & "Alt+P", , , iP(0), , , , "mnu_Previous")
        iP(1) = .AddItem("N&ew Tab" & vbTab & "Alt+E", , , iP(0), , , , "mnu_AddTab")
        iP(1) = .AddItem("&Delete Tab" & vbTab & "Alt+W", , , iP(0), , , , "mnu_DeleteTab")
        iP(1) = .AddItem("Delete &All Tabs" & vbTab & "Alt+A", , , iP(0), , , , "mnu_DeleteAllTabs")
        
        'Security
        iP(0) = .AddItem("Se&curity", , , , , , , "mnusecurity")
        iP(1) = .AddItem("Current Internet Zone", , , iP(0), 6, , , "mnu_currentzone")
        iP(1) = .AddItem("Scripting", , , iP(0), , , , "mnuscripting")
        iP(2) = .AddItem("Disable", , , iP(1), , , , "mnujdisable")
        iP(2) = .AddItem("Enable", , , iP(1), , , , "mnujenable")
        iP(2) = .AddItem("Prompt", , , iP(1), , , , "mnujprompt")
        iP(1) = .AddItem("Cookies", , , iP(0), , , , "mnucookies")
        iP(2) = .AddItem("Disable", , , iP(1), , , , "mnuCdisable")
        iP(2) = .AddItem("Enable", , , iP(1), , , , "mnuCenable")
        iP(2) = .AddItem("Prompt", , , iP(1), , , , "mnuCprompt")
        iP(1) = .AddItem("Session", , , iP(0), , , , "mnusession")
        iP(2) = .AddItem("Disable", , , iP(1), , , , "mnuSdisable")
        iP(2) = .AddItem("Enable", , , iP(1), , , , "mnuSenable")
        iP(2) = .AddItem("Prompt", , , iP(1), , , , "mnuSprompt")
        iP(1) = .AddItem("ActiveX", , , iP(0), , , , "mnuactivex")
        iP(2) = .AddItem("Disable", , , iP(1), , , , "mnuAdisable")
        iP(2) = .AddItem("Enable", , , iP(1), , , , "mnuAenable")
        iP(2) = .AddItem("Prompt", , , iP(1), , , , "mnuAprompt")
        iP(1) = .AddItem("-", , , iP(0), , , , "sep13")
        iP(1) = .AddItem("Settings", , , iP(0), 5, , , "mnu_settings")
        
        ' Help menu.
        iP(0) = .AddItem("&Help", , , , , , , "mnu_Help")
        iP(1) = .AddItem("&Contents", , , iP(0), , , , "Contents")
        iP(1) = .AddItem("&Tip of the Day", , , iP(0), , , , "TipOfDay")
        iP(1) = .AddItem("-", , , iP(0), , , , "line3")
        iP(1) = .AddItem("On The Internet", , , iP(0), , , , "Corrupted Inc")
        iP(1) = .AddItem("&Send Feedback", , , iP(0), , , , "Feedback")
        iP(1) = .AddItem("-", , , iP(0), , , , "line2")
        iP(1) = .AddItem("&About...", , , iP(0), , , , "About")
        iP(1) = .AddItem("Credits", , , iP(0), , , , "mnucredits")
        
    End With
    'Other Menus / Popups
    
    '        iP(0) = .AddItem("mnuTrayPopup", , , , ,  , False, "mnu_TrayPopup")
    '        iP(1) = .AddItem("UndergroundSearch", , , iP(0), , , , "mnu_SurfTabs")
    '        iP(1) = .AddItem("Exit", , , iP(0), , , , "mnu_TrayExit")
    '        iP(0) = .AddItem("mnuFavoritesPopUp", , , , , , False, "mnu_FavoritesPopUp")
    '        iP(1) = .AddItem("OpenFavoriteincurrentTab", , , iP(0), , , , "mnu_OpenFavoriteCurrentTab")
    '        iP(1) = .AddItem("OpenFavoriteinNewTab", , , iP(0), , , , "mnu_OpenFavoriteNewTab")
    '        iP(1) = .AddItem("mnuFavoritesFolderPopUp", , , iP(0), , , , "mnu_FavoritesFolderPopUp")
    '        iP(1) = .AddItem("OpenallFavoritesinthisfolder", , , iP(0), , , , "mnu_OpenAllFavoritesInFolder")
    '        iP(0) = .AddItem("mnuHistoryPopUp", , , , , , False, "mnu_HistoryPopUp")
    '        iP(1) = .AddItem("OpenHistoryincurrentTab", , , iP(0), , , , "mnu_OpenHistoryInCurrentTab")
    '        iP(1) = .AddItem("OpenHistoryinNewTab", , , iP(0), , , , "mnu_OpenHistoryInNewTab")
    '        iP(1) = .AddItem("mnuHistoryFolderPopUp", , , iP(0), , , , "mnu_HistoryFolderPopUp")
    '        iP(1) = .AddItem("OpenthisdomaininNewtab", , , iP(0), , , , "mnu_OpomainNewTab")
    '        iP(1) = .AddItem("Openthisdomainincurrenttab", , , iP(0), , , , "mnu_OpomainCurrentTab")
    '        iP(1) = .AddItem("OpenAllHistorypagesinthisdomain", , , iP(0), , , , "mnu_OpenAllHistoryPages")
    '        iP(0) = .AddItem("mnuTabPopUp", , , , , , False, "mnu_TabPopUp")
    '        iP(1) = .AddItem("CloseCurrentTab", , , iP(0), , , , "mnu_TabPopUp_DeleteTab")
    '        iP(1) = .AddItem("OpenNewTab", , , iP(0), , , , "mnu_TabPopUp_NewTab")
    '        iP(1) = .AddItem("DuplicatCurrentTab", , , iP(0), , , , "mnu_TabPopUp_DuplicateTab")
    '        iP(1) = .AddItem("CloseAllTabs", , , iP(0), , , , "mnu_TabPopUp_DeleteAllTabs")
    
End Sub

Sub setupToolbar()
    On Error Resume Next
    Dim i As Long
    
    ' Set up the main toolbar
    With tbrTools
        .ImageSource = CTBExternalImageList
        .SetImageList frmBrowser.vbalImageList1, CTBImageListNormal
        
        .CreateToolbar 32, True, False
        .AddButton "New Tab", 13, , , , CTBNormal Or CTBAutoSize, "NewTab"
        .AddButton , , , , , CTBSeparator
        .AddButton "Go Back", 11, , , , CTBDropDown Or CTBAutoSize, "Back"
        .AddButton "Go Forward", 12, , , , CTBDropDown Or CTBAutoSize, "Forward"
        .AddButton , , , , , CTBSeparator
        .AddButton "Stop", 10, , , , CTBNormal Or CTBAutoSize, "Stop"
        .AddButton "Refresh Page", 9, , , , CTBNormal Or CTBAutoSize, "Refresh"
        .AddButton "Go Home", 7, , , , CTBNormal Or CTBAutoSize, "Home"
        .AddButton , , , , , CTBSeparator
        .AddButton "Search", 8, , , , CTBNormal Or CTBAutoSize, "Search"
        .AddButton "View Favorites", 6, , , , CTBNormal Or CTBAutoSize, "Favorites"
        .AddButton "View History", 5, , , , CTBNormal Or CTBAutoSize, "History"
        .AddButton , , , , , CTBSeparator
        .AddButton "Print Page", 4, , , , CTBNormal Or CTBAutoSize, "Print"
        .AddButton "Browser Options", 3, , , , CTBNormal Or CTBAutoSize, "Options"
        .AddButton "Java Options", 0, , , , CTBNormal Or CTBAutoSize, "Security"
        .AddButton "Image Options", 2, , , , CTBNormal Or CTBAutoSize, "Media"
        .AddButton "Text Editor", 1, , , , CTBNormal Or CTBAutoSize, "Editor"
        
    End With
    
    LockWindowUpdate Me.hWnd
    With frmBrowser.rbrMain
        .Position = erbPositionTop
        '.BackgroundBitmap = GetString(HKEY_CURRENT_USER, "Software\Corrupted Inc\Underground Search\Options", "Background")
        .CreateRebar frmBrowser.hWnd
        .AddBandByHwnd frmBrowser.tbrMenu.hWnd, , , , "MENU"
        .AddBandByHwnd frmBrowser.tbrTools.hWnd, , , , "TOOLBAR"
        .AddBandByHwnd frmBrowser.picAnim.hWnd, , False, True, "ANIM"
        .AddBandByHwnd cboAddress.hWnd, "Address", , , "Address"
        .AddBandByHwnd CboTranslate.hWnd, "Translate Page", False, , "Translate"
        .AddBandByHwnd fraSearch.hWnd, "Underground Search", True, , "Search"
        
        For i = 0 To .BandCount - 1
            If i <> 1 Then
                .BandChildMinWidth(i) = 110
                .BandChildMinHeight(3) = 20
            End If
        Next i
    End With
    LockWindowUpdate 0
End Sub

Private Sub rbrMain_BandChildResize(ByVal wID As Long, ByVal lBandLeft As Long, ByVal lBandTop As Long, ByVal lBandRight As Long, ByVal lBandBottom As Long, lChildLeft As Long, lChildTop As Long, lChildRight As Long, lChildBottom As Long)
    If rbrMain.BandData(rbrMain.BandIndexForId(wID)) = "ANIM" Then
        picAnim.Width = (lChildRight - lChildLeft) * Screen.TwipsPerPixelX
        picAnim.Height = (lChildBottom - lChildTop) * Screen.TwipsPerPixelX
    End If
End Sub

Private Sub m_cMenu_Click(ItemNumber As Long)
    On Error Resume Next
    Dim bS As Boolean
    Dim lIndex As Long
    Dim Val1 As Long
    Dim Val2 As Long
    Dim Cap1 As Long
    Dim Cap2 As Long
    Dim Cap3 As Long
    Dim Cap4 As Long
    Dim RetVal
    Cap1 = m_cMenu.IndexForKey("mnuscripting")
    Cap2 = m_cMenu.IndexForKey("mnucookies")
    Cap3 = m_cMenu.IndexForKey("mnusession")
    Cap4 = m_cMenu.IndexForKey("mnuactivex")
    Select Case m_cMenu.ItemKey(ItemNumber)
        
    Case "mnu_NewBrowserCurrent"
        Call NewTab(frmBrowser, brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL, NEW_TAB_CUR_URL)
        
    Case "mnu_NewBrowserHome"
        Call NewTab(frmBrowser, "", NEW_TAB_HOME)
        
    Case "mnu_NewBrowserBlank"
        Call NewTab(frmBrowser, "", NEW_TAB_BLANK)
        
    Case "mnu_Open"
        frmOpen.Show 1
        
        If OpenOk = True Then
            If OpenNewTab Then
                Call NewTab(Me, OpenURL, -1)
            Else
                brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate OpenURL
            End If
        End If
    Case "mnu_SaveAs"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
        
    Case "mnu_PageSetup"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT
        
    Case "mnu_Print"
        Call PrintBrowser
        
    Case "mnupreview"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
        
    Case "mnu_Properties"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT
        
    Case "mnu_WorkOffline"
        'Place that stuff in a sub and call it
        
    Case "mnu_Exit"
        Unload Me
        
    Case "mnu_EditCut"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT
        'Clipboard.GetText
    Case "mnu_EditCopy"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
        'Clipboard.SetText
    Case "mnu_EditPast"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT
        'Clipboard.GetText
    Case "mnu_EditSelectAll"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
        
    Case "mnu_EditFind"
        SetFocusOnly = True
        TabStrip1.SetFocus
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).SetFocus
        SendKeys "^f"
        
    Case "mnu_ViewStatusBar"
        bS = Not (m_cMenu.Checked(ItemNumber))
        m_cMenu.Checked(ItemNumber) = bS
    Case "mnutoolbar"
        ViewBars ItemNumber
    Case "mnuaddressbar"
        '        bS = Not (m_cMenu.Checked(ItemNumber))
        '        m_cMenu.Checked(ItemNumber) = bS
        ViewBars ItemNumber
    Case "mnutranslatebar"
        '        bS = Not (m_cMenu.Checked(ItemNumber))
        '        m_cMenu.Checked(ItemNumber) = bS
        ViewBars ItemNumber
    Case "mnusearchbar"
        '        bS = Not (m_cMenu.Checked(ItemNumber))
        '        m_cMenu.Checked(ItemNumber) = bS
        ViewBars ItemNumber
        
    Case "mnu_ViewStop"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Stop
        
    Case "mnu_ViewRefresh"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Refresh
        
    Case "mnu_ViewStopAllTabs"
        Dim X
        For X = 1 To TabStrip1.tabs.Count
            brwWebBrowser(TabStrip1.tabs(X).Tag).Stop
        Next
        
        
    Case "mnu_ViewRefreshAllTabs"
        'Dim X
        For X = 1 To TabStrip1.tabs.Count
            brwWebBrowser(TabStrip1.tabs(X).Tag).Refresh
        Next
        
    Case "mnu_ViewSource"
        If Len(brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Document.documentelement.innerhtml) > 0 Then
            
            frmMain.rt.Text = brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Document.documentelement.innerhtml
            frmMain.Caption = "Underground HTML Editor - [" & brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL & "]"
            frmMain.SetSyntax
            frmMain.Show
        End If
    Case "Fullscreen"
        Call FullScreen
    Case "mnu_GoBack"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoBack
        
    Case "mnu_GoForward"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoForward
        
    Case "mnu_GoHome"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoHome
        
    Case "mnu_GoSearch"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).GoSearch
        
        
    Case "mnu_GoPreviousTab"
        'Move the current browser off the form\tab
        'Move the next browser onto the form\tab
        Call MoveBrowserOffFormTab(TabStrip1.tabs(CurTab_Index).Tag)
        
        If CurTab_Index = 1 Then
            CurTab_Index = MaxTab_Index
        Else
            CurTab_Index = CurTab_Index - 1
        End If
        
        Call SelectBrowserTab(CurTab_Index)
        MoveBrowsers
        
    Case "mnu_GoNextTab"
        'Move the current browser off the form\tab
        'Move the next browser onto the form\tab
        Call MoveBrowserOffFormTab(TabStrip1.tabs(CurTab_Index).Tag)
        
        If CurTab_Index = MaxTab_Index Then
            CurTab_Index = 1
        Else
            CurTab_Index = CurTab_Index + 1
        End If
        
        Call SelectBrowserTab(CurTab_Index)
        MoveBrowsers
    Case "Favorites"
        '        lIndex = m_cMenu.IndexForKey("AddFavorites")
        '        m_cMenu.ShowPopupAbsolute 2500, 500, lIndex
        '        '        lIndex = m_cMenu.IndexForKey("AddFavorites")
        '        '        m_cMenu.ShowPopupMenuAtIndex 2500, -100, lIndex:=lIndex
    Case "AddFavorites"
        Call AddToFavorites
        Call GetFavorites  'To refresh the favorites tree
        
    Case "OrganizeFavorites"
        Dim lpszRootFolder As String
        Dim success As Long
        Dim CSIDL As Long
        
        'open the organize folder at the path specified by the CSIDL
        CSIDL = CSIDL_FAVORITES
        
        lpszRootFolder = GetFolderPath(CSIDL)
        success = DoOrganizeFavDlg(hWnd, lpszRootFolder)
        
        Call GetFavorites  'To refresh the favorites tree
        
    Case "mnurussiantoeng"
        PageUrl = ""
        PageUrl = brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL
        PageUrl = Right(PageUrl, Len(PageUrl) - 7)
        TransUrl = ""
        TransUrl = "http://babel.altavista.com/translate.dyn?doit=done&BabelFishFrontPage=yes&enc=utf8&bblType=url&url=http%3A%2F%2F" & PageUrl & "%2F%3Ff%3Dext&lp=ru_en"
        Debug.Print TransUrl
        Call NewTab(Me, TransUrl, -1)
        
    Case "mnudetoeng"
        PageUrl = ""
        PageUrl = brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL
        PageUrl = Right(PageUrl, Len(PageUrl) - 7)
        TransUrl = ""
        TransUrl = "http://babel.altavista.com/translate.dyn?doit=done&BabelFishFrontPage=yes&enc=utf8&bblType=url&url=http%3A%2F%2F" & PageUrl & "%2F&lp=de_en"
        Call NewTab(Me, TransUrl, -1)
        
    Case "mnuporttoeng"
        PageUrl = ""
        PageUrl = brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL
        PageUrl = Right(PageUrl, Len(PageUrl) - 7)
        TransUrl = ""
        TransUrl = "http://babel.altavista.com/translate.dyn?doit=done&BabelFishFrontPage=yes&enc=utf8&bblType=url&url=http%3A%2F%2F" & PageUrl & "%2F%3Ff%3Dext&lp=pt_en"
        Debug.Print TransUrl
        Call NewTab(Me, TransUrl, -1)
        
    Case "mnufrtoeng"
        PageUrl = ""
        PageUrl = brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL
        PageUrl = Right(PageUrl, Len(PageUrl) - 7)
        TransUrl = ""
        TransUrl = "http://babel.altavista.com/translate.dyn?doit=done&BabelFishFrontPage=yes&enc=utf8&bblType=url&url=http%3A%2F%2F" & PageUrl & "%2F%3Ff%3Dext&lp=fr_en"
        Debug.Print TransUrl
        Call NewTab(Me, TransUrl, -1)
        
    Case "mnupopup"
        lIndex = m_cMenu.IndexForKey("mnupopup")
        If m_cMenu.Checked(lIndex) = True Then
            m_cMenu.Checked(lIndex) = False
            PopupKill = False
            Call SaveDword(HKEY_CURRENT_USER, "\Software\Corrupted Inc\Underground Search\Options", "PopupKilla", 0)
        ElseIf m_cMenu.Checked(lIndex) = False Then
            m_cMenu.Checked(lIndex) = True
            PopupKill = True
            Call frmpopupkilla.LoadPopupList(lstPopups, App.path & "/popups.dat")
            Call SaveDword(HKEY_CURRENT_USER, "\Software\Corrupted Inc\Underground Search\Options", "PopupKilla", 1)
        End If
        
    Case "mnulargest"
        'lIndex = m_cMenu.IndexForKey("mnulargest")
        'If m_cMenu.Checked(lIndex) = True Then
        '    m_cMenu.Checked(lIndex) = False
        'Call SetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\International\Scripts\3", "IEFontSize", "4")
        '    End If
        ' HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\International\Scripts\3
    Case "mnusrcoptions"
        frmSearchOptions.Show
        
    Case "mnu_SurfTabOptions"
        frmOptions.Show 1
        
    Case "mnu_settings"
        RetVal = Shell("rundll32.exe shell32.dll,Control_RunDLL Inetcpl.cpl,,1", vbNormalFocus)
    Case "mnu_InterNetOptions"
        
        RetVal = Shell("rundll32.exe shell32.dll,Control_RunDLL Inetcpl.cpl", vbNormalFocus)
        
    Case "mnu_PrevTab"
        Call mnu_GoPreviousTab_Click
        Call SelectBrowserTab(CurTab_Index)
        MoveBrowsers
        
    Case "mnu_NextTab"
        Call mnu_GoNextTab_Click
        Call SelectBrowserTab(CurTab_Index)
        MoveBrowsers
        
    Case "mnu_AddTab"
        Call NewTab(Me, "", -1)
        
    Case "mnu_DeleteTab"
        Call DeleteTab
    Case "mnu_Previous"
        'Move the current browser off the form\tab
        'Move the next browser onto the form\tab
        Call MoveBrowserOffFormTab(TabStrip1.tabs(CurTab_Index).Tag)
        
        If CurTab_Index = 1 Then
            CurTab_Index = MaxTab_Index
        Else
            CurTab_Index = CurTab_Index - 1
        End If
        
        Call SelectBrowserTab(CurTab_Index)
        MoveBrowsers
        
    Case "mnu_DeleteAllTabs"
        Call DeleteAllTabs
        
    Case "mnu_About"
        frmAbout.Show 1
        If cboAddress.Text = "http://www.corrupted1.f2s.com" Then
            cboAddress_KeyPress (vbKeyReturn)
        End If
        
    Case "mnucredits"
        frmCredits.Show
        
    Case "mnu_TrayExit"
        Unload Me
    Case "mnu_newMessage"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate ("mailto:" & "" & "?subject=" & "")
        
    Case "mnu_Sendlink"
        Dim strLink As String
        strLink = ""
        strLink = "mailto:?subject=Link From Underground Browser"
        strLink = strLink + "&body= Please Visit this Link Brought to you By the Underground Browser. "
        strLink = strLink + brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL
        Debug.Print "Link is "; strLink
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate strLink
        
    Case "mnu_Sendpage"
        strLink = ""
        strLink = "mailto:?subject=File From Underground Browser"
        strLink = strLink + "&body= Havent figured out this attachment thing yet. "
        strLink = strLink + brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL
        
        Debug.Print "Link is "; strLink
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate strLink
        
        
    Case "mnu_Readmail"
        Shell "C:\Program Files\Outlook Express\MSIMN.EXE", vbNormalFocus
        
    Case "mnuhotmail"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate "www.hotmail.com"
        
    Case "mnu_outlook"
        Shell "C:\Program Files\Outlook Express\MSIMN.EXE", vbNormalFocus
        
    Case "mnu_OpenFavoriteCurrentTab"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate gFavoritesURL
        
    Case "mnu_OpenFavoriteNewTab"
        Call NewTab(Me, gFavoritesURL, -1)
        
    Case "mnu_OpenHistoryInCurrentTab"
        brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate gHistoryURL
        
    Case "mnu_OpenHistoryInNewTab"
        Call NewTab(Me, gHistoryURL, -1)
        
    Case "mnuquick"
        lIndex = m_cMenu.IndexForKey("mnuquick")
        If m_cMenu.Checked(lIndex) = True Then
            bS = Not (m_cMenu.Checked(ItemNumber))
            m_cMenu.Checked(lIndex) = False
            QuickDisable = False
            Call SaveDword(HKEY_CURRENT_USER, "\Software\Corrupted Inc\Underground Search\Options", "QuickDisable", 0)
        ElseIf m_cMenu.Checked(lIndex) = False Then
            m_cMenu.Checked(lIndex) = True
            QuickDisable = True
            Call SaveDword(HKEY_CURRENT_USER, "\Software\Corrupted Inc\Underground Search\Options", "QuickDisable", 1)
        End If
        
    Case "mnutranslate"
        bS = Not (m_cMenu.Checked(ItemNumber))
        m_cMenu.Checked(ItemNumber) = bS
        
    Case "mnujenable"
        lIndex = m_cMenu.IndexForKey("mnujenable")
        Val1 = m_cMenu.IndexForKey("mnujdisable")
        Val2 = m_cMenu.IndexForKey("mnujprompt")
        bS = Not (m_cMenu.Checked(lIndex))
        Call SaveDword(HKEY_CURRENT_USER, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1400", 0)
        m_cMenu.Checked(lIndex) = bS
        m_cMenu.Checked(Val1) = False
        m_cMenu.Checked(Val2) = False
        m_cMenu.Caption(Cap1) = "Scripting - Enabled"
    Case "mnujdisable"
        lIndex = m_cMenu.IndexForKey("mnujdisable")
        Val1 = m_cMenu.IndexForKey("mnujenable")
        Val2 = m_cMenu.IndexForKey("mnujprompt")
        bS = Not (m_cMenu.Checked(ItemNumber))
        m_cMenu.Checked(lIndex) = bS
        Call SaveDword(HKEY_CURRENT_USER, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1400", 3)
        m_cMenu.Checked(Val1) = False
        m_cMenu.Checked(Val2) = False
        m_cMenu.Caption(Cap1) = "Scripting - Disabled"
    Case "mnujprompt"
        lIndex = m_cMenu.IndexForKey("mnujprompt")
        Val1 = m_cMenu.IndexForKey("mnujenable")
        Val2 = m_cMenu.IndexForKey("mnujdisable")
        bS = Not (m_cMenu.Checked(lIndex))
        m_cMenu.Checked(lIndex) = bS
        Call SaveDword(HKEY_CURRENT_USER, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1400", 1)
        m_cMenu.Checked(Val1) = False
        m_cMenu.Checked(Val2) = False
        m_cMenu.Caption(Cap1) = "Scripting - Prompt"
    Case "mnuCdisable"
        lIndex = m_cMenu.IndexForKey("mnuCdisable")
        Val1 = m_cMenu.IndexForKey("mnuCenable")
        Val2 = m_cMenu.IndexForKey("mnuCprompt")
        bS = Not (m_cMenu.Checked(ItemNumber))
        m_cMenu.Checked(lIndex) = bS
        Call SaveDword(HKEY_CURRENT_USER, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A02", 3)
        m_cMenu.Checked(Val1) = False
        m_cMenu.Checked(Val2) = False
        m_cMenu.Caption(Cap2) = "Cookies - Disabled"
    Case "mnuCenable"
        lIndex = m_cMenu.IndexForKey("mnuCenable")
        Val1 = m_cMenu.IndexForKey("mnuCdisable")
        Val2 = m_cMenu.IndexForKey("mnuCprompt")
        bS = Not (m_cMenu.Checked(lIndex))
        m_cMenu.Checked(lIndex) = bS
        Call SaveDword(HKEY_CURRENT_USER, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A02", 0)
        m_cMenu.Checked(Val1) = False
        m_cMenu.Checked(Val2) = False
        m_cMenu.Caption(Cap2) = "Cookies - Enabled"
    Case "mnuCprompt"
        lIndex = m_cMenu.IndexForKey("mnuCprompt")
        Val1 = m_cMenu.IndexForKey("mnuCenable")
        Val2 = m_cMenu.IndexForKey("mnuCdisable")
        bS = Not (m_cMenu.Checked(lIndex))
        m_cMenu.Checked(lIndex) = bS
        Call SaveDword(HKEY_CURRENT_USER, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A02", 1)
        m_cMenu.Checked(Val1) = False
        m_cMenu.Checked(Val2) = False
        m_cMenu.Caption(Cap2) = "Cookies - Prompt"
    Case "mnuSdisable"
        lIndex = m_cMenu.IndexForKey("mnuSdisable")
        Val1 = m_cMenu.IndexForKey("mnuSenable")
        Val2 = m_cMenu.IndexForKey("mnuSpropmt")
        bS = Not (m_cMenu.Checked(lIndex))
        m_cMenu.Checked(lIndex) = bS
        Call SaveDword(HKEY_CURRENT_USER, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A03", 3)
        m_cMenu.Checked(Val1) = False
        m_cMenu.Checked(Val2) = False
        m_cMenu.Caption(Cap3) = "Session - Disabled"
    Case "mnuSenable"
        lIndex = m_cMenu.IndexForKey("mnuSenable")
        Val1 = m_cMenu.IndexForKey("mnuSdisable")
        Val2 = m_cMenu.IndexForKey("mnuSprompt")
        bS = Not (m_cMenu.Checked(lIndex))
        m_cMenu.Checked(lIndex) = bS
        Call SaveDword(HKEY_CURRENT_USER, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A03", 0)
        m_cMenu.Checked(Val1) = False
        m_cMenu.Checked(Val2) = False
        m_cMenu.Caption(Cap3) = "Session - Enabled"
    Case "mnuSprompt"
        lIndex = m_cMenu.IndexForKey("mnuSprompt")
        Val1 = m_cMenu.IndexForKey("mnuSenable")
        Val2 = m_cMenu.IndexForKey("mnuSdisable")
        bS = Not (m_cMenu.Checked(ItemNumber))
        m_cMenu.Checked(lIndex) = bS
        Call SaveDword(HKEY_CURRENT_USER, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A03", 1)
        m_cMenu.Checked(Val1) = False
        m_cMenu.Checked(Val2) = False
        m_cMenu.Caption(Cap3) = "Session - Prompt"
    Case "mnuAdisable"
        lIndex = m_cMenu.IndexForKey("mnuAdisable")
        Val1 = m_cMenu.IndexForKey("mnuAenable")
        Val2 = m_cMenu.IndexForKey("mnuApropmt")
        bS = Not (m_cMenu.Checked(lIndex))
        m_cMenu.Checked(lIndex) = bS
        Call SaveDword(HKEY_CURRENT_USER, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1200", 3)
        m_cMenu.Checked(Val1) = False
        m_cMenu.Checked(Val2) = False
        m_cMenu.Caption(Cap4) = "ActiveX - Disabled"
    Case "mnuAenable"
        lIndex = m_cMenu.IndexForKey("mnuAenable")
        Val1 = m_cMenu.IndexForKey("mnuAdisable")
        Val2 = m_cMenu.IndexForKey("mnuApropmt")
        bS = Not (m_cMenu.Checked(lIndex))
        m_cMenu.Checked(lIndex) = bS
        Call SaveDword(HKEY_CURRENT_USER, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1200", 0)
        m_cMenu.Checked(Val1) = False
        m_cMenu.Checked(Val2) = False
        m_cMenu.Caption(Cap4) = "ActiveX - Enabled"
    Case "mnuAprompt"
        lIndex = m_cMenu.IndexForKey("mnuApropmt")
        Val1 = m_cMenu.IndexForKey("mnuAenable")
        Val2 = m_cMenu.IndexForKey("mnuAdisable")
        bS = Not (m_cMenu.Checked(lIndex))
        m_cMenu.Checked(lIndex) = bS
        Call SaveDword(HKEY_CURRENT_USER, "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1200", 1)
        m_cMenu.Checked(Val1) = False
        m_cMenu.Checked(Val2) = False
        m_cMenu.Caption(Cap4) = "ActiveX - Prompt"
    Case "mnuimages"
        If m_cMenu.Checked(54) = True Then
            m_cMenu.Checked(54) = False
            Call SaveString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Display Inline Images", "no")
        ElseIf m_cMenu.Checked(54) = False Then
            m_cMenu.Checked(54) = True
            Call SaveString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Display Inline Images", "yes")
        End If
    Case "mnuvideos"
        If m_cMenu.Checked(55) = True Then
            m_cMenu.Checked(55) = False
            Call SaveString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Display Inline Videos", "no")
        ElseIf m_cMenu.Checked(55) = False Then
            m_cMenu.Checked(55) = True
            Call SaveString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Display Inline Videos", "yes")
        End If
    Case "mnusounds"
        If m_cMenu.Checked(56) = True Then
            m_cMenu.Checked(56) = False
            Call SaveString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Play_Background_Sounds", "no")
        ElseIf m_cMenu.Checked(56) = False Then
            m_cMenu.Checked(56) = True
            Call SaveString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Play_Background_Sounds", "yes")
        End If
    Case "mnuanimation"
        If m_cMenu.Checked(57) = True Then
            m_cMenu.Checked(57) = False
            Call SaveString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Play_Animations", "no")
        ElseIf m_cMenu.Checked(57) = False Then
            m_cMenu.Checked(57) = True
            Call SaveString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Play_Animations", "yes")
        End If
        
    Case "mnuworkoffline"
        lIndex = m_cMenu.IndexForKey("mnuworkoffline")
        bS = Not (m_cMenu.Checked(lIndex))
        Dim TabCount
        If m_cMenu.Checked(14) = True Then
            m_cMenu.Checked(14) = False
            If TabStrip1.tabs.Count > 1 Then
                X = 1
                TabCount = TabStrip1.tabs.Count
                While TabCount >= X
                    TabStrip1.tabs(TabCount).Selected = True
                    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Offline = False
                    StatusBar.Panels(1).Text = "Working Online in All Tabs"
                    TabCount = TabCount - 1
                Wend
            Else
                brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Offline = False
            End If
        Else
            m_cMenu.Checked(14) = True
            If TabStrip1.tabs.Count > 1 Then
                X = 1
                TabCount = TabStrip1.tabs.Count
                
                While TabCount >= X
                    TabStrip1.tabs(TabCount).Selected = True
                    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Offline = True
                    brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Refresh
                    StatusBar.Panels(1).Text = "Working Offline in All Tabs"
                    TabCount = TabCount - 1
                Wend
            Else
                brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Offline = True
                
            End If
        End If
        
        
    Case "mnu_TabPopUp_DeleteTab"
        Call DeleteTab
        
    Case "mnu_TabPopUp_NewTab"
        '   Call NewTab
        
    Case "mnu_TabPopUp_DuplicateTab"
        Call NewTab(Me, brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).LocationURL, -1)
        
        
    Case "mnu_TabPopUp_DeleteAllTabs"
        Call DeleteAllTabs
        
    Case "About"
        frmAbout.Show
        
    Case Else
        Dim FavStr$, a%
        FavStr$ = m_cMenu.ItemKey(ItemNumber)
        If InStr(FavStr$, "Folder") = 0 Then
            a% = InStr(FavStr$, "URL")
            If a% > 0 Then
                FavStr$ = Right$(FavStr$, Len(FavStr$) - 3)
                FavStr$ = Trim(FavStr$)
                brwWebBrowser(TabStrip1.tabs(CurTab_Index).Tag).Navigate FavStr$
                cboAddress.Text = FavStr$
            End If
        Else
            MsgBox "This folder is empty!  You may add an item by going to add to favorites and clicking on this folder.", vbExclamation + vbOKOnly, "Empty Favorite Folder"
        End If
    End Select
    
End Sub
Private Function plGetIndexFromKey(ByVal sKey As String) As Long
    Dim iPos As Long
    Dim iNextPos As Long
    iPos = InStr(sKey, "(")
    If iPos > 0 Then
        iNextPos = InStr(iPos, sKey, ")")
        If iNextPos > 0 Then
            plGetIndexFromKey = Mid$(sKey, iPos + 1, iNextPos - iPos - 1)
        End If
    End If
End Function
Sub LoadChkMenus()
    ' Dont Puzzle this section as I stressed for Hours as how to get
    'this section working, using the smallest way possible and all failed
    'Due to the way the menus are built
    'So i had to resort to this method.
    'Declares
    '-----------------------------------------
    Dim strValue1 As String
    Dim strValue2 As String
    Dim strValue3 As String
    Dim strValue4 As String
    Dim strImages As String
    Dim strSounds As String
    Dim strAnimations As String
    Dim strVideos As String
    Dim jVal1 As Long
    Dim jVal2 As Long
    Dim jVal3 As Long
    Dim cVal1 As Long
    Dim cVal2 As Long
    Dim cVal3 As Long
    Dim sVal1 As Long
    Dim sVal2 As Long
    Dim sVal3 As Long
    Dim aVal1 As Long
    Dim aVal2 As Long
    Dim aVal3 As Long
    Dim lIndex As Long
    Dim Cap1 As Long
    Dim Cap2 As Long
    Dim Cap3 As Long
    Dim Cap4 As Long
    'Setting Values
    '-------------------------------------------------
    strValue1 = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1400")
    strValue2 = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A02")
    strValue3 = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A03")
    strValue4 = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1200")
    strImages = ""
    strImages = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Display Inline Images")
    strSounds = ""
    strSounds = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Play_Background_Sounds")
    strAnimations = ""
    strAnimations = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Play_Animations")
    strVideos = ""
    strVideos = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Display Inline Videos")
    '----------------------------------------------
    jVal1 = m_cMenu.IndexForKey("mnujdisable")
    jVal2 = m_cMenu.IndexForKey("mnujenable")
    jVal3 = m_cMenu.IndexForKey("mnujprompt")
    cVal1 = m_cMenu.IndexForKey("mnuCdisable")
    cVal2 = m_cMenu.IndexForKey("mnuCenable")
    cVal3 = m_cMenu.IndexForKey("mnuCprompt")
    sVal1 = m_cMenu.IndexForKey("mnuSdisable")
    sVal2 = m_cMenu.IndexForKey("mnuSenable")
    sVal3 = m_cMenu.IndexForKey("mnuSprompt")
    aVal1 = m_cMenu.IndexForKey("mnuAdisable")
    aVal2 = m_cMenu.IndexForKey("mnuAenable")
    aVal3 = m_cMenu.IndexForKey("mnuAprompt")
    Cap1 = m_cMenu.IndexForKey("mnuscripting")
    Cap2 = m_cMenu.IndexForKey("mnucookies")
    Cap3 = m_cMenu.IndexForKey("mnusession")
    Cap4 = m_cMenu.IndexForKey("mnuactivex")
    '-----------------------------------------------
    
    'Scripting
    Select Case strValue1
    Case "3"
        m_cMenu.Checked(jVal1) = True
        m_cMenu.Caption(Cap1) = m_cMenu.Caption(Cap1) + " - Disabled"
    Case "0"
        m_cMenu.Checked(jVal2) = True
        m_cMenu.Caption(Cap1) = m_cMenu.Caption(Cap1) + " - Enabled"
    Case "1"
        m_cMenu.Checked(jVal3) = True
        m_cMenu.Caption(Cap1) = m_cMenu.Caption(Cap1) + " - Prompt"
    End Select
    'Cookies
    
    Select Case strValue2
    Case "3"
        m_cMenu.Checked(cVal1) = True
        m_cMenu.Caption(Cap2) = m_cMenu.Caption(Cap2) + " -  Disabled"
    Case "0"
        m_cMenu.Checked(cVal2) = True
        m_cMenu.Caption(Cap2) = m_cMenu.Caption(Cap2) + " -  Enabled"
    Case "1"
        m_cMenu.Checked(cVal3) = True
        m_cMenu.Caption(Cap2) = m_cMenu.Caption(Cap2) + " -  Prompt"
    End Select
    'Session
    
    Select Case strValue3
    Case "3"
        m_cMenu.Checked(sVal1) = True
        m_cMenu.Caption(Cap3) = m_cMenu.Caption(Cap3) + " -  Disabled"
    Case "0"
        m_cMenu.Checked(sVal2) = True
        m_cMenu.Caption(Cap3) = m_cMenu.Caption(Cap3) + " -  Enabled"
    Case "1"
        m_cMenu.Checked(sVal3) = True
        m_cMenu.Caption(Cap3) = m_cMenu.Caption(Cap3) + " -  Prompt"
    End Select
    'ActiveX
    
    Select Case strValue4
    Case "3"
        m_cMenu.Checked(aVal1) = True
        m_cMenu.Caption(Cap4) = m_cMenu.Caption(Cap4) + " -  Disabled"
    Case "0"
        m_cMenu.Checked(aVal2) = True
        m_cMenu.Caption(Cap4) = m_cMenu.Caption(Cap4) + " -  Enabled"
    Case "1"
        m_cMenu.Checked(aVal3) = True
        m_cMenu.Caption(Cap4) = m_cMenu.Caption(Cap4) + " -  Prompt"
    End Select
    'Load the Media Stuff
    lIndex = m_cMenu.IndexForKey("mnuimages")
    If strImages = "yes" Then
        m_cMenu.Checked(lIndex) = True
    ElseIf strImages = "no" Then
        m_cMenu.Checked(lIndex) = False
    End If
    lIndex = m_cMenu.IndexForKey("mnuvideos")
    If strVideos = "yes" Then
        m_cMenu.Checked(lIndex) = True
    ElseIf strVideos = "no" Then
        m_cMenu.Checked(lIndex) = False
    End If
    lIndex = m_cMenu.IndexForKey("mnusounds")
    If strSounds = "yes" Then
        m_cMenu.Checked(lIndex) = True
    ElseIf strSounds = "no" Then
        m_cMenu.Checked(lIndex) = False
    End If
    lIndex = m_cMenu.IndexForKey("mnuanimation")
    If strAnimations = "yes" Then
        m_cMenu.Checked(lIndex) = True
    ElseIf strAnimations = "no" Then
        m_cMenu.Checked(lIndex) = False
    End If
    
End Sub
Private Sub ViewBars(ByVal lIndex As Long)
    Dim bS As Boolean
    Dim lItemIndex As Long
    Dim l As Long
    bS = Not (m_cMenu.Checked(lIndex))
    m_cMenu.Checked(lIndex) = bS
    Select Case lIndex
    Case 26
        l = rbrMain.BandIndexForData("TOOLBAR")
        rbrMain.BandVisible(l) = bS
        Call Form_Resize
    Case 27
        l = rbrMain.BandIndexForData("Address")
        rbrMain.BandVisible(l) = bS
        Call Form_Resize
    Case 28
        l = rbrMain.BandIndexForData("Translate")
        rbrMain.BandVisible(l) = bS
        Call Form_Resize
    Case 29
        l = rbrMain.BandIndexForData("Search")
        rbrMain.BandVisible(l) = bS
        Call Form_Resize
    End Select
    
End Sub

Private Sub txtsearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode <> vbKeyControl Then
        With txtsearch
            Select Case KeyCode
            Case vbKeyV  'PAST
                .SelText = Clipboard.GetText(vbCFText)
                KeyCode = 0
            Case vbKeyX  'CUT
                Clipboard.SetText txtsearch.SelText
                txtsearch.SelText = ""
                KeyCode = 0
            Case vbKeyC  'COPY
                Clipboard.SetText txtsearch.SelText
                KeyCode = 0
            Case vbKeyZ  'UNDO
                KeyCode = 0
            Case vbKeyY  'REDO
                KeyCode = 0
            End Select
        End With
    End If
End Sub


Private Sub txtsearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then
        cboAddress.Text = Clipboard.GetText
    End If
End Sub


