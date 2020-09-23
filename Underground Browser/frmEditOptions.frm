VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8A61D56-D8DC-11D2-8064-9D6F06504DA8}#1.1#0"; "AXCOLCTL.OCX"
Begin VB.Form frmEditOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmEditOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picColors 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   360
      ScaleHeight     =   2175
      ScaleWidth      =   4095
      TabIndex        =   4
      Top             =   480
      Width           =   4095
      Begin VB.ComboBox cboFnt 
         Height          =   315
         ItemData        =   "frmEditOptions.frx":030A
         Left            =   2280
         List            =   "frmEditOptions.frx":031D
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ListBox lstOpt 
         Height          =   2010
         ItemData        =   "frmEditOptions.frx":0351
         Left            =   0
         List            =   "frmEditOptions.frx":038B
         TabIndex        =   6
         Top             =   120
         Width           =   2175
      End
      Begin ImgColorPicker.ColorPicker clrFore 
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Appearance      =   2
      End
      Begin ImgColorPicker.ColorPicker clrBack 
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Appearance      =   2
      End
      Begin VB.Label Label3 
         Caption         =   "Font Options:"
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Back Color:"
         Height          =   495
         Left            =   2280
         TabIndex        =   10
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Color:"
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.PictureBox picGeneral 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   360
      ScaleHeight     =   2175
      ScaleWidth      =   4095
      TabIndex        =   3
      Top             =   480
      Width           =   4095
      Begin VB.CheckBox chkSel 
         Caption         =   "Sel Bounds"
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CheckBox chkWhite 
         Caption         =   "Display White Spaces"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox chkLeft 
         Caption         =   "Left Margin"
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox chkLine 
         Caption         =   "Line Numbering"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox chkHighlight 
         Caption         =   "Highlight Line"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tabs 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4683
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General Options"
            Key             =   "goptions"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Color Settings"
            Key             =   "colors"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4560
      Y1              =   2850
      Y2              =   2850
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4560
      Y1              =   2880
      Y2              =   2880
   End
End
Attribute VB_Name = "frmEditOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboFnt_Click()
    
  ClrData(lstOpt.ListIndex).fntProp = cboFnt.ListIndex

Exit Sub



    ' Then UDBErrorHandler "(Form) frmEditOptions::Sub cboFnt_Click"
    Resume Next
End Sub

Private Sub clrBack_ColorChanged()
    
  ClrData(lstOpt.ListIndex).bgClr = clrBack

Exit Sub



    ' Then UDBErrorHandler "(Form) frmEditOptions::Sub clrBack_ColorChanged"
    Resume Next
End Sub

Private Sub clrFore_ColorChanged()
    
  ClrData(lstOpt.ListIndex).frClr = clrFore

Exit Sub



    ' Then UDBErrorHandler "(Form) frmEditOptions::Sub clrFore_ColorChanged"
    Resume Next
End Sub

Private Sub cmdCancel_Click()
    
  Unload Me

Exit Sub



    ' Then UDBErrorHandler "(Form) frmEditOptions::Sub cmdCancel_Click"
    Resume Next
End Sub

Private Sub List1_Click()
    

Exit Sub



    ' Then UDBErrorHandler "(Form) frmEditOptions::Sub List1_Click"
    Resume Next
End Sub

Private Sub cmdOK_Click()
    
  Dim arg1 As String, argument As String, s As String
  For X = 0 To lstOpt.ListCount - 1
    s = ClrData(X).bgClr
    writeini "Colors", "clr" & X & "bg", s, App.path & "\UndergroundBrowse.ini"
    s = ClrData(X).frClr
    writeini "Colors", "clr" & X & "fr", s, App.path & "\UndergroundBrowse.ini"
    s = ClrData(X).fntProp
    writeini "Colors", "clr" & X & "fnt", s, App.path & "\UndergroundBrowse.ini"
  Next
  writeini "Options", "highlight", chkHighlight.Value, App.path & "\UndergroundBrowse.ini"
  writeini "Options", "linenumber", chkLine.Value, App.path & "\UndergroundBrowse.ini"
  writeini "Options", "leftmargin", chkLeft.Value, App.path & "\UndergroundBrowse.ini"
  writeini "Options", "whitespace", chkWhite.Value, App.path & "\UndergroundBrowse.ini"
  writeini "Options", "selbounds", chkSel.Value, App.path & "\UndergroundBrowse.ini"
  
  'Set the colors
  frmMain.rt.SetColor cmClrBookmark, ClrData(0).frClr
  frmMain.rt.SetColor cmClrBookmarkBk, ClrData(0).bgClr
  frmMain.rt.SetColor cmClrComment, ClrData(1).frClr
  frmMain.rt.SetColor cmClrCommentBk, ClrData(1).bgClr
  frmMain.rt.SetColor cmClrHDividerLines, ClrData(2).frClr
  frmMain.rt.SetColor cmClrVDividerLines, ClrData(3).frClr
  frmMain.rt.SetColor cmClrHighlightedLine, ClrData(4).frClr
  frmMain.rt.SetColor cmClrKeyword, ClrData(5).frClr
  frmMain.rt.SetColor cmClrKeywordBk, ClrData(5).bgClr
  frmMain.rt.SetColor cmClrLeftMargin, ClrData(6).frClr
  frmMain.rt.SetColor cmClrLineNumber, ClrData(7).frClr
  frmMain.rt.SetColor cmClrLineNumberBk, ClrData(7).bgClr
  frmMain.rt.SetColor cmClrNumber, ClrData(8).frClr
  frmMain.rt.SetColor cmClrNumberBk, ClrData(8).bgClr
  frmMain.rt.SetColor cmClrOperator, ClrData(9).frClr
  frmMain.rt.SetColor cmClrOperatorBk, ClrData(9).bgClr
  frmMain.rt.SetColor cmClrScopeKeyword, ClrData(10).frClr
  frmMain.rt.SetColor cmClrScopeKeywordBk, ClrData(10).bgClr
  frmMain.rt.SetColor cmClrString, ClrData(11).frClr
  frmMain.rt.SetColor cmClrStringBk, ClrData(11).bgClr
  frmMain.rt.SetColor cmClrTagElementName, ClrData(12).frClr
  frmMain.rt.SetColor cmClrTagElementNameBk, ClrData(12).bgClr
  frmMain.rt.SetColor cmClrTagEntity, ClrData(13).frClr
  frmMain.rt.SetColor cmClrTagEntityBk, ClrData(13).bgClr
  frmMain.rt.SetColor cmClrTagAttributeName, ClrData(14).frClr
  frmMain.rt.SetColor cmClrTagAttributeNameBk, ClrData(14).bgClr
  frmMain.rt.SetColor cmClrTagText, ClrData(15).frClr
  frmMain.rt.SetColor cmClrTagTextBk, ClrData(15).bgClr
  frmMain.rt.SetColor cmClrText, ClrData(16).frClr
  frmMain.rt.SetColor cmClrTextBk, ClrData(16).bgClr
  frmMain.rt.SetColor cmClrWindow, ClrData(17).frClr
  
  'Setup the font styles
  frmMain.rt.SetFontStyle cmStyComment, frmMain.txtProp(ClrData(1).fntProp)
  frmMain.rt.SetFontStyle cmStyLineNumber, frmMain.txtProp(ClrData(7).fntProp)
  frmMain.rt.SetFontStyle cmStyNumber, frmMain.txtProp(ClrData(8).fntProp)
  frmMain.rt.SetFontStyle cmStyOperator, frmMain.txtProp(ClrData(9).fntProp)
  frmMain.rt.SetFontStyle cmStyScopeKeyword, frmMain.txtProp(ClrData(10).fntProp)
  frmMain.rt.SetFontStyle cmStyString, frmMain.txtProp(ClrData(11).fntProp)
  frmMain.rt.SetFontStyle cmStyTagAttributeName, frmMain.txtProp(ClrData(12).fntProp)
  frmMain.rt.SetFontStyle cmStyTagAttributeName, frmMain.txtProp(ClrData(13).fntProp)
  frmMain.rt.SetFontStyle cmStyTagEntity, frmMain.txtProp(ClrData(14).bgClr)
  frmMain.rt.SetFontStyle cmStyKeyword, frmMain.txtProp(ClrData(5).fntProp)
  frmMain.rt.SetFontStyle cmStyTagText, frmMain.txtProp(ClrData(15).fntProp)
  frmMain.rt.SetFontStyle cmStyNumber, frmMain.txtProp(ClrData(8).fntProp)
  frmMain.rt.SetFontStyle cmStyText, frmMain.txtProp(ClrData(16).fntProp)
  
 WhiteSpace = chkWhite.Value
 LineNumbering = chkLine.Value
 HighLight = chkHighlight.Value
 LeftMargin = chkLeft.Value
 SelBounds = chkSel.Value
 
 If HighLight = 0 Then frmMain.rt.HighlightedLine = -1
 If LineNumbering = 0 Then
    frmMain.rt.LineNumbering = False
  Else
    frmMain.rt.LineNumbering = True
  End If
  If LeftMargin = 0 Then
    frmMain.rt.DisplayLeftMargin = False
  Else
    frmMain.rt.DisplayLeftMargin = True
  End If
 If WhiteSpace = 0 Then
   frmMain.rt.DisplayWhitespace = False
 Else
   frmMain.rt.DisplayWhitespace = True
 End If
 If SelBounds = 0 Then
   frmMain.rt.SelBounds = False
 Else
   frmMain.rt.SelBounds = True
 End If
  Unload Me

Exit Sub



    ' Then UDBErrorHandler "(Form) frmEditOptions::Sub cmdOK_Click"
    Resume Next
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    For X = 0 To lstOpt.ListCount - 1
      ClrData(X).bgClr = ReadINI("Colors", "clr" & X & "bg", App.path & "\UndergroundBrowse.ini")
      ClrData(X).frClr = ReadINI("Colors", "clr" & X & "fr", App.path & "\UndergroundBrowse.ini")
      ClrData(X).fntProp = ReadINI("Colors", "clr" & X & "fnt", App.path & "\UndergroundBrowse.ini")
    Next
    HighLight = ReadINI("options", "highlight", App.path & "\UndergroundBrowse.ini")
    WhiteSpace = ReadINI("options", "whitespace", App.path & "\UndergroundBrowse.ini")
    LineNumbering = ReadINI("options", "linenumber", App.path & "\UndergroundBrowse.ini")
    LeftMargin = ReadINI("options", "leftmargin", App.path & "\UndergroundBrowse.ini")
    SelBounds = ReadINI("options", "selbounds", App.path & "\UndergroundBrowse.ini")
    
    chkLeft.Value = LeftMargin
    chkLine.Value = LineNumbering
    chkWhite.Value = WhiteSpace
    chkHighlight.Value = HighLight
    chkSel.Value = SelBounds

Exit Sub



    ' Then UDBErrorHandler "(Form) frmEditOptions::Sub Form_Load"
    Resume Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
  frmMain.rt.SetFocus

Exit Sub



    ' Then UDBErrorHandler "(Form) frmEditOptions::Sub Form_QueryUnload"
    Resume Next
End Sub

Private Sub lstOpt_Click()
    
  SetVis
  clrFore.Color = ClrData(lstOpt.ListIndex).frClr
  clrBack.Color = ClrData(lstOpt.ListIndex).bgClr
  cboFnt.ListIndex = ClrData(lstOpt.ListIndex).fntProp

Exit Sub



    ' Then UDBErrorHandler "(Form) frmEditOptions::Sub lstOpt_Click"
    Resume Next
End Sub

Private Sub tabs_Click()
    
  If tabs.tabs(2).Selected = True Then
    picColors.Visible = True
    picGeneral.Visible = False
    picColors.SetFocus
    lstOpt.ListIndex = 0
  Else
    picColors.Visible = False
    picGeneral.Visible = True
    picGeneral.SetFocus
  End If

Exit Sub



    ' Then UDBErrorHandler "(Form) frmEditOptions::Sub tabs_Click"
    Resume Next
End Sub

Public Function SetVis()
    
  Select Case lstOpt.ListIndex
    Case 0
      ColorsTrue
    Case 1
      AllTrue
    Case 2
      ColorTrue
    Case 3
      ColorTrue
    Case 4
      ColorTrue
    Case 5
      AllTrue
    Case 6
      ColorTrue
    Case 7
      AllTrue
    Case 8
      AllTrue
    Case 9
      AllTrue
    Case 10
      AllTrue
    Case 11
      AllTrue
    Case 12
      AllTrue
    Case 13
      AllTrue
    Case 14
      AllTrue
    Case 15
      AllTrue
    Case 16
      AllTrue
    Case 17
      ColorTrue
  End Select

Exit Function



    ' Then UDBErrorHandler "(Form) frmEditOptions::Function SetVis"
    Resume Next
End Function

Public Function AllTrue()
    
      Label1.Visible = True
      Label2.Visible = True
      Label3.Visible = True
      clrFore.Visible = True
      clrBack.Visible = True
      cboFnt.Visible = True

Exit Function



    ' Then UDBErrorHandler "(Form) frmEditOptions::Function AllTrue"
    Resume Next
End Function

Public Function ColorTrue()
    
      Label1.Visible = True
      Label2.Visible = False
      Label3.Visible = False
      clrFore.Visible = True
      clrBack.Visible = False
      cboFnt.Visible = False

Exit Function



    ' Then UDBErrorHandler "(Form) frmEditOptions::Function ColorTrue"
    Resume Next
End Function

Public Function ColorsTrue()
    
      Label1.Visible = True
      Label2.Visible = True
      Label3.Visible = False
      clrFore.Visible = True
      clrBack.Visible = True
      cboFnt.Visible = False

Exit Function



    ' Then UDBErrorHandler "(Form) frmEditOptions::Function ColorsTrue"
    Resume Next
End Function
