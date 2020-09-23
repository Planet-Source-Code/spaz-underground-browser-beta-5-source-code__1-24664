VERSION 5.00
Object = "{ECEDB943-AC41-11D2-AB20-000000000000}#2.0#0"; "CMAX20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00808080&
   Caption         =   "HTML Editor | Taken from the Best Editor cEdit"
   ClientHeight    =   5640
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8625
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5385
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   970
            MinWidth        =   970
            Text            =   "Line:"
            TextSave        =   "Line:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1676
            MinWidth        =   1676
            Text            =   "Total Lines:"
            TextSave        =   "Total Lines:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cm 
      Left            =   7080
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   7440
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0864
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0976
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   688
      BandCount       =   1
      FixedOrder      =   -1  'True
      EmbossPicture   =   -1  'True
      _CBWidth        =   8625
      _CBHeight       =   390
      _Version        =   "6.7.8988"
      Child1          =   "tBar"
      MinHeight1      =   330
      Width1          =   8625
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tBar 
         Height          =   330
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imgList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   19
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "new"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "open"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "save"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cut"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "copy"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "paste"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "delete"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "undo"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "redo"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "properties"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "help"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "about"
               ImageIndex      =   12
            EndProperty
         EndProperty
         MouseIcon       =   "frmMain.frx":12FC
         Begin VB.ComboBox cboLang 
            Height          =   315
            ItemData        =   "frmMain.frx":145E
            Left            =   4200
            List            =   "frmMain.frx":147D
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   0
            Width           =   1335
         End
      End
   End
   Begin CodeMaxCtl.CodeMax rt 
      Height          =   4575
      Left            =   120
      OleObjectBlob   =   "frmMain.frx":14C9
      TabIndex        =   0
      Top             =   600
      Width           =   8415
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New Document"
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "&Open Document"
         Shortcut        =   ^O
      End
      Begin VB.Menu bar0 
         Caption         =   "-"
      End
      Begin VB.Menu save 
         Caption         =   "&Save Document"
         Shortcut        =   ^S
      End
      Begin VB.Menu saveas 
         Caption         =   "Save Document &As"
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu pagesetup 
         Caption         =   "Page &Setup.."
      End
      Begin VB.Menu print 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu bar21 
         Caption         =   "-"
      End
      Begin VB.Menu clearclip 
         Caption         =   "&Clear Clipboard"
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu undo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu redo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu bar5 
         Caption         =   "-"
      End
      Begin VB.Menu cut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu delete 
         Caption         =   "&Delete"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu bar6 
         Caption         =   "-"
      End
      Begin VB.Menu find 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu findnext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu replace 
         Caption         =   "&Replace"
         Shortcut        =   ^R
      End
      Begin VB.Menu goto 
         Caption         =   "&Go To..."
         Shortcut        =   ^G
      End
      Begin VB.Menu bar7 
         Caption         =   "-"
      End
      Begin VB.Menu selectall 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu bar9 
         Caption         =   "-"
      End
      Begin VB.Menu Timedate 
         Caption         =   "&Time/Date"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu tools 
      Caption         =   "&Tools"
      Begin VB.Menu wordcount 
         Caption         =   "&Word Count"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu options 
      Caption         =   "&Options"
      Begin VB.Menu main 
         Caption         =   "Main Options"
         Shortcut        =   ^H
      End
      Begin VB.Menu bar11 
         Caption         =   "-"
      End
      Begin VB.Menu colors 
         Caption         =   "&Colors"
      End
      Begin VB.Menu bar17 
         Caption         =   "-"
      End
      Begin VB.Menu fileassociations 
         Caption         =   "&Setup File Associations"
         Visible         =   0   'False
      End
      Begin VB.Menu bar20 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu font 
         Caption         =   "&Font"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Visible         =   0   'False
      Begin VB.Menu helpme 
         Caption         =   "&General Help"
         Shortcut        =   {F1}
         Visible         =   0   'False
      End
      Begin VB.Menu readme 
         Caption         =   "&Readme"
         Visible         =   0   'False
      End
      Begin VB.Menu bar15 
         Caption         =   "-"
      End
      Begin VB.Menu website 
         Caption         =   "&Acksoft Website"
         Visible         =   0   'False
      End
      Begin VB.Menu bar16 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu about 
         Caption         =   "&About"
         Shortcut        =   {F6}
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Changed As Boolean
Dim DocFile As Boolean
  Dim r As CodeMaxCtl.Range

Dim DocName As String
Dim DocFileName As String
Private Sub about_Click()
  frmAbout.Show
End Sub

Private Sub cboLang_Click()
    
  On Error Resume Next
  If cboLang.ListIndex > 6 Then
    rt.Language = "java"
    Exit Sub
  End If
  If cboLang.ListIndex <> 0 Then
    rt.Language = cboLang.Text
  Else
    rt.Language = ""
  End If
  rt.SetFocus

End Sub

Private Sub clearclip_Click()
    
  Dim msg As VbMsgBoxResult
  msg = MsgBox("Are you sure you wish to clear the clipboard?", vbYesNo + vbQuestion, "Clipboard")
  If msg = vbYes Then
    Clipboard.SetText ""
  Else
    Exit Sub
  End If
End Sub

Private Sub colors_Click()
    
  frmEditOptions.Show
  frmEditOptions.tabs.tabs(2).Selected = True
End Sub

Private Sub Command1_Click()
    
  rt.ExecuteCmd cmCmdProperties
End Sub

Private Sub copy_Click()
    
  On Error Resume Next
  rt.copy
End Sub

Private Sub cut_Click()
    
  On Error Resume Next
  rt.cut
End Sub

Private Sub delete_Click()
    
  On Error Resume Next
  rt.SelText = ""
End Sub

Private Sub fileassociations_Click()
    
  frmNew.Show
End Sub

Private Sub find_Click()
    
  On Error Resume Next
  Dim cf As CodeMaxCtl.globals
  rt.ExecuteCmd cmCmdFind

End Sub

Private Sub findnext_Click()
    
  On Error Resume Next
  rt.ExecuteCmd cmCmdFindNext
End Sub

Private Sub font_Click()
    
  On Error GoTo er
  cm.CancelError = True
  cm.Flags = cdlCFScreenFonts
  cm.ShowFont
  rt.font.Bold = cm.FontBold
  rt.font.Italic = cm.FontItalic
  rt.font.Name = cm.FontName
  rt.font.Size = cm.FontSize
  rt.font.Strikethrough = cm.FontStrikethru
  rt.font.Underline = cm.FontUnderline
  writeini "Font", "Bold", rt.font.Bold, App.path & "\UndergroundBrowse.ini"
  writeini "Font", "Italic", rt.font.Italic, App.path & "\UndergroundBrowse.ini"
  writeini "Font", "Name", rt.font.Name, App.path & "\UndergroundBrowse.ini"
  writeini "Font", "Size", rt.font.Size, App.path & "\UndergroundBrowse.ini"
  writeini "Font", "Strikethrough", rt.font.Strikethrough, App.path & "\UndergroundBrowse.ini"
  writeini "Font", "Underline", rt.font.Underline, App.path & "\UndergroundBrowse.ini"
er:
  Exit Sub

End Sub

Private Sub Form_Load()
    
  'On Error Resume Next
  Set r = rt.GetSel(False)
  Dim hk As CodeMaxCtl.HotKey, hk_index As Integer
  Dim num_hk As Long, cmGlobals As CodeMaxCtl.globals
  Dim Cmd(7) As CodeMaxCtl.cmCommand, cmd_index As Integer
  
  Set cmGlobals = New CodeMaxCtl.globals
  Set hk = New CodeMaxCtl.HotKey
  
  'Read the ini File data :)
  HighLight = ReadINI("options", "highlight", App.path & "\UndergroundBrowse.ini")
  LineNumbering = ReadINI("options", "linenumber", App.path & "\UndergroundBrowse.ini")
  LeftMargin = ReadINI("options", "leftmargin", App.path & "\UndergroundBrowse.ini")
  WhiteSpace = ReadINI("options", "whitespace", App.path & "\UndergroundBrowse.ini")
  SelBounds = ReadINI("options", "selbounds", App.path & "\UndergroundBrowse.ini")
   
  'Set drag and drop
  rt.EnableDragDrop = True
   
  'Lets load some font data up :)
  rt.font.Bold = ReadINI("font", "bold", App.path & "\UndergroundBrowse.ini")
  rt.font.Italic = ReadINI("font", "italic", App.path & "\UndergroundBrowse.ini")
  rt.font.Name = ReadINI("font", "name", App.path & "\UndergroundBrowse.ini")
  rt.font.Size = ReadINI("font", "size", App.path & "\UndergroundBrowse.ini")
  rt.font.Strikethrough = ReadINI("font", "strikethrough", App.path & "\UndergroundBrowse.ini")
  rt.font.Underline = ReadINI("font", "underline", App.path & "\UndergroundBrowse.ini")
   
  If WhiteSpace = 1 Then rt.DisplayWhitespace = True
  If LeftMargin = 1 Then rt.DisplayLeftMargin = True
  If LineNumbering = 0 Then rt.LineNumbering = False
  If SelBounds = 1 Then rt.SelBounds = True

  'Read the color data from the ini file
  
   For X = 0 To 17
      ClrData(X).bgClr = ReadINI("Colors", "clr" & X & "bg", App.path & "\UndergroundBrowse.ini")
      ClrData(X).frClr = ReadINI("Colors", "clr" & X & "fr", App.path & "\UndergroundBrowse.ini")
      ClrData(X).fntProp = ReadINI("Colors", "clr" & X & "fnt", App.path & "\UndergroundBrowse.ini")
   Next
   
  'Use the color data from the ini file
  rt.SetColor cmClrBookmark, ClrData(0).frClr
  rt.SetColor cmClrBookmarkBk, ClrData(0).bgClr
  rt.SetColor cmClrCommentBk, ClrData(1).bgClr
  rt.SetColor cmClrComment, ClrData(1).frClr
  rt.SetColor cmClrHDividerLines, ClrData(2).frClr
  rt.SetColor cmClrVDividerLines, ClrData(3).frClr
  rt.SetColor cmClrHighlightedLine, ClrData(4).frClr
  rt.SetColor cmClrKeyword, ClrData(5).frClr
  rt.SetColor cmClrKeywordBk, ClrData(5).bgClr
  rt.SetColor cmClrLeftMargin, ClrData(6).frClr
  rt.SetColor cmClrLineNumber, ClrData(7).frClr
  rt.SetColor cmClrLineNumberBk, ClrData(7).bgClr
  rt.SetColor cmClrNumber, ClrData(8).frClr
  rt.SetColor cmClrNumberBk, ClrData(8).bgClr
  rt.SetColor cmClrOperator, ClrData(9).frClr
  rt.SetColor cmClrOperatorBk, ClrData(9).bgClr
  rt.SetColor cmClrScopeKeyword, ClrData(10).frClr
  rt.SetColor cmClrScopeKeywordBk, ClrData(10).bgClr
  rt.SetColor cmClrString, ClrData(11).frClr
  rt.SetColor cmClrStringBk, ClrData(11).bgClr
  rt.SetColor cmClrTagElementName, ClrData(12).frClr
  rt.SetColor cmClrTagElementNameBk, ClrData(12).bgClr
  rt.SetColor cmClrTagEntity, ClrData(13).frClr
  rt.SetColor cmClrTagEntityBk, ClrData(13).bgClr
  rt.SetColor cmClrTagAttributeName, ClrData(14).frClr
  rt.SetColor cmClrTagAttributeNameBk, ClrData(14).bgClr
  rt.SetColor cmClrTagText, ClrData(15).frClr
  rt.SetColor cmClrTagTextBk, ClrData(15).bgClr
  rt.SetColor cmClrText, ClrData(16).frClr
  rt.SetColor cmClrTextBk, ClrData(16).bgClr
  rt.SetColor cmClrWindow, ClrData(17).frClr
  
  'Setup font styles
  rt.SetFontStyle cmStyComment, txtProp(ClrData(1).fntProp)
  rt.SetFontStyle cmStyLineNumber, txtProp(ClrData(7).fntProp)
  rt.SetFontStyle cmStyNumber, txtProp(ClrData(8).fntProp)
  rt.SetFontStyle cmStyOperator, txtProp(ClrData(9).fntProp)
  rt.SetFontStyle cmStyScopeKeyword, txtProp(ClrData(10).fntProp)
  rt.SetFontStyle cmStyString, txtProp(ClrData(11).fntProp)
  rt.SetFontStyle cmStyTagAttributeName, txtProp(ClrData(12).fntProp)
  rt.SetFontStyle cmStyTagAttributeName, txtProp(ClrData(13).fntProp)
  rt.SetFontStyle cmStyTagEntity, txtProp(ClrData(14).bgClr)
  rt.SetFontStyle cmStyKeyword, txtProp(ClrData(5).fntProp)
  rt.SetFontStyle cmStyTagText, txtProp(ClrData(15).fntProp)
  rt.SetFontStyle cmStyNumber, txtProp(ClrData(8).fntProp)
  rt.SetFontStyle cmStyText, txtProp(ClrData(16).fntProp)
  
  'Get the old window data and set it up :)
  Me.Left = ReadINI("Window", "Left", App.path & "\UndergroundBrowse.ini")
  Me.Top = ReadINI("Window", "Top", App.path & "\UndergroundBrowse.ini")
  Me.Width = ReadINI("Window", "Width", App.path & "\UndergroundBrowse.ini")
  Me.Height = ReadINI("Window", "Height", App.path & "\UndergroundBrowse.ini")
  Me.WindowState = ReadINI("Window", "windowstate", App.path & "\UndergroundBrowse.ini")
  
  'Register the perl language
  RegisterPerl rt
  cboLang.ListIndex = 0

  'Unregister a few of the hotkeys in the codemax control
  
  Cmd(1) = cmCmdCut
  Cmd(2) = cmCmdPaste
  Cmd(3) = cmCmdCopy
  Cmd(4) = cmCmdLineCut
  Cmd(7) = cmCmdLineDelete
  Cmd(5) = cmCmdUndo
  Cmd(6) = cmCmdRedo
  
  For cmd_index = 1 To 7
     num_hk = cmGlobals.GetNumHotKeysForCmd(Cmd(cmd_index))
     For hk_index = num_hk - 1 To 0 Step -1
       Set hk = cmGlobals.GetHotKeyForCmd(Cmd(cmd_index), hk_index)
       Call cmGlobals.UnregisterHotKey(hk)
     Next hk_index
  Next cmd_index
  
  'Setup neccisary data for application to run :)
  SetSyntax
  Changed = False
  If Command = "" Then
    doNew
  Else
    cm.FileName = Command
    Changed = False
    DocFile = True
    DocFileName = Command
    DocName = Command
    Me.Caption = "HTML Editor | Taken from the Best Editor cEdit - [" & Command & "]"
    SetSyntax
    rt.OpenFile Command
  End If
    stBar.Panels(4).Text = rt.LineCount
  stBar.Panels(2).Text = r.StartLineNo + 1

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
  Dim msg As VbMsgBoxResult
     If Changed = True Then
    msg = MsgBox("Document: " & DocName & Chr(10) & Chr(10) & "This document has been modified." & Chr(10) & "Do you wish to save?", vbYesNoCancel, "Document Modified.")
    Select Case msg
      Case vbCancel
        Cancel = 1
      Case vbNo
      Case vbYes
        doSave
     End Select
   End If

End Sub

Private Sub Form_Resize()
    
  On Error Resume Next
  rt.Move 0, cBar.Top + cBar.Height, Me.ScaleWidth, Me.ScaleHeight - stBar.Height - cBar.Height

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
  'Write the current width/height data

  writeini "Window", "windowstate", Me.WindowState, App.path & "\UndergroundBrowse.ini"
  Me.WindowState = 0
  writeini "Window", "Left", Me.Left, App.path & "\UndergroundBrowse.ini"
  writeini "Window", "Top", Me.Top, App.path & "\UndergroundBrowse.ini"
  writeini "Window", "Width", Me.Width, App.path & "\UndergroundBrowse.ini"
  writeini "Window", "Height", Me.Height, App.path & "\UndergroundBrowse.ini"
  
  Unload frmAbout
  Unload frmOptions
  Unload Me

End Sub

Private Sub goto_Click()
    
  'On Error Resume Next
  rt.ExecuteCmd cmCmdGotoLine, -1

End Sub

Private Sub main_Click()
    
  frmEditOptions.Show
  frmEditOptions.tabs.tabs(1).Selected = True

End Sub

Private Sub new_Click()
    
  doNew
End Sub

Private Sub open_Click()
    
  Dim msg As VbMsgBoxResult
  'Check the document
  On Error GoTo exitSub
   If Changed = True Then
    msg = MsgBox("Document: " & DocName & Chr(10) & Chr(10) & "This document has been modified." & Chr(10) & "Do you wish to save?", vbYesNoCancel, "Document Modified.")
    Select Case msg
      Case vbCancel
        Exit Sub
      Case vbNo
      Case vbYes
        doSave
     End Select
   End If
  cm.CancelError = True
  cm.DialogTitle = "Open a document..."
  cm.Filter = "All Files|*.*|All Supported Files|*txt;*.htm;*.html;*css;*.js;*.c;*.cpp;*.h;*.pl;*.cgi;*.xml;*.pas;*.bas;*.frm;*.vbp|Text Files|*.txt|Html Files|*.html;*.htm|Java Script Files|*.js|Style Sheets|*.cs|C/C++ Files|*.c;*.cpp;*.h|Perl Files|*.pl|CGI/Perl Files|*.cgi|XML Files|*.xml|Pascal Files|*.pas|Basic Files|*.bas;*.frm;*.vbp"
  cm.ShowOpen
  If cm.FileName = "" Then Exit Sub
  SetSyntax
  rt.OpenFile cm.FileName
  Changed = False
  DocFile = True
  DocFileName = cm.FileName
  DocName = cm.FileName
  Me.Caption = "HTML Editor | Taken from the Best Editor cEdit - [" & cm.FileName & "]"
exitSub:

End Sub

Private Sub pagesetup_Click()
    
  On Error Resume Next
  Call rt.PrintContents(0, cmPrnColor + cmPrnRichFonts)

End Sub

Private Sub paste_Click()
    
  On Error Resume Next
  rt.paste

End Sub

Private Sub print_Click()
    
  Call rt.PrintContents(0, cmPrnColor + cmPrnDefaultPrn + cmPrnRichFonts)

End Sub

Private Sub properties_Click()
    
  frmProp.Show

End Sub

Private Sub readme_Click()
    
  Shell App.path & "\cedit.exe " & App.path & "\readme.txt", vbNormalFocus

End Sub

Private Sub redo_Click()
    
  On Error Resume Next
  rt.redo

End Sub

Private Sub replace_Click()
    
  On Error Resume Next
  rt.ExecuteCmd cmCmdFindReplace

End Sub

Private Sub rt_Change(ByVal Control As CodeMaxCtl.ICodeMax)
    
  Changed = True

End Sub

Private Function rt_MouseDown(ByVal rt As CodeMaxCtl.ICodeMax, ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    

  If Button = 2 Then PopupMenu edit

End Function

Public Function CheckDoc()
 
End Function

Public Function doSave()
    
  If DocFile = True Then
    rt.SaveFile DocFileName, True
    Changed = False
  Else
    DoSaveAs
  End If

End Function

Public Function DoSaveAs()
    
  Dim msg As VbMsgBoxResult
  On Error GoTo exitSub
newDialog:
  cm.DialogTitle = "Save a document..."
  cm.CancelError = True
  cm.Filter = "All Files|*.*|Text Files|*.txt|Html Files|*.html;*.htm|Style Sheets|*.css|Java Scripting|*.js|C Files|*.c|C++ Files|*.cpp|C/C++ H Files|*.h|Perl Files|*.pl|CGI/Perl Files|*.cgi|XML Files|*.xml|Pascal Files|*.pas|Basic Module Files|*.bas|Basic Form Files|*.frm|Basic Project Files|*.vbp"
  cm.ShowSave
  If cm.FileName = "" Or cm.FileName = " " Then Exit Function
  If FileExists(cm.FileName) = True Then
    msg = MsgBox("File: " & cm.FileName & Chr(10) & Chr(10) & "The file you have entered" & Chr(10) & "already exists. Overwrite?", vbYesNoCancel, "Overwrite?")
    Select Case msg
      Case vbNo
        GoTo newDialog
      Case vbCancel
        Exit Function
     End Select
  End If
  rt.SaveFile cm.FileName, True
  SetSyntax
  DocFile = True
  DocFileName = cm.FileName
  Changed = False
  DocName = cm.FileName
  Me.Caption = "HTML Editor | Taken from the Best Editor cEdit - [" & cm.FileName & "]"
exitSub:
  
End Function

Private Function FileExists(FullFileName As String) As Boolean
    
    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #1
        Close #1
        'no error, file exists
        FileExists = True
    Exit Function
MakeF:
        'error, file does Not exist
        FileExists = False
    Exit Function

End Function

Public Function SetSyntax()
    
  Select Case LCase(Right(cm.FileName, 2))
    Case ".c"
      cboLang.ListIndex = 2
      'rt.Language = "c/c++"
    Case ".h"
      cboLang.ListIndex = 2
      rt.Language = "c/c++"
  End Select
  
  Select Case LCase(Right(cm.FileName, 3))
    Case ".pl"
      cboLang.ListIndex = 3
    Case ".js"
      cboLang.ListIndex = 7
    End Select
    
  Select Case LCase(Right(cm.FileName, 4))
    Case ".txt"
      cboLang.ListIndex = 0
    Case ".htm"
      cboLang.ListIndex = 1
    Case ".cpp"
      cboLang.ListIndex = 2
    Case ".cgi"
      cboLang.ListIndex = 3
    Case ".xml"
      cboLang.ListIndex = 4
    Case ".pas"
      cboLang.ListIndex = 6
    Case ".bas"
      cboLang.ListIndex = 5
    Case ".frm"
      cboLang.ListIndex = 5
    Case ".vbp"
      cboLang.ListIndex = 5
    Case ".css"
      cboLang.ListIndex = 8
  End Select
  
  If LCase(Right(cm.FileName, 5)) = ".html" Then
    cboLang.ListIndex = 1
  End If

End Function

Private Sub RegisterPerl(ByVal rt As CodeMaxCtl.ICodeMax)
    
Dim lang As CodeMaxCtl.Language
Set lang = New CodeMaxCtl.Language
lang.CaseSensitive = True
Dim str As String
str = ""
str = "abs" + Chr(10) + "if" + Chr(10) + "accept" + Chr(10) + "accept" + Chr(10) + "alarm" + Chr(10) + "atan2" + Chr(10) + "atan2" + Chr(10) + "bind" + Chr(10) + "bind" + Chr(10) + "binmode" + Chr(10)
str = str + "binmode" + Chr(10) + "bless" + Chr(10) + "bless" + Chr(10) + "caller" + Chr(10) + "chdir" + Chr(10) + "chmod" + Chr(10) + "chmod" + Chr(10) + "chomp" + Chr(10) + "chop" + Chr(10) + "chown" + Chr(10)
str = str + "chr" + Chr(10) + "chroot" + Chr(10) + "close" + Chr(10) + "closedir" + Chr(10) + "connect" + Chr(10) + "connect" + Chr(10) + "cos" + Chr(10) + "crypt" + Chr(10) + "crypt" + Chr(10) + "dbmclose" + Chr(10)
str = str + "dbmopen" + Chr(10) + "dbmopen" + Chr(10) + "dbmopen" + Chr(10) + "dbmopen" + Chr(10) + "dbmopen" + Chr(10) + "dbmopen" + Chr(10) + "dbmopen" + Chr(10) + "defined" + Chr(10) + "delete" + Chr(10) + "die " + Chr(10)
str = str + "do" + Chr(10) + "dump" + Chr(10) + "each" + Chr(10) + "eof" + Chr(10) + "eval" + Chr(10) + "exec" + Chr(10) + "exists" + Chr(10) + "exit" + Chr(10) + "exp" + Chr(10) + "fcntl" + Chr(10)
str = str + "fcntl" + Chr(10) + "fcntl" + Chr(10) + "fileno" + Chr(10) + "flock" + Chr(10) + "flock" + Chr(10) + "fork" + Chr(10) + "format" + Chr(10) + "formline" + Chr(10) + "formline" + Chr(10) + "getc" + Chr(10)
str = str + "getlogin" + Chr(10) + "getpeername" + Chr(10) + "getpgrp" + Chr(10) + "getppid" + Chr(10) + "getpriority" + Chr(10) + "getpriority" + Chr(10) + "getsockname" + Chr(10) + "getsockopt" + Chr(10) + "getsockopt" + Chr(10) + "getsockopt" + Chr(10)
str = str + "glob" + Chr(10) + "gmtime" + Chr(10) + "goto" + Chr(10) + "goto" + Chr(10) + "goto" + Chr(10) + "grep" + Chr(10) + "grep" + Chr(10) + "hex" + Chr(10) + "import" + Chr(10) + "index" + Chr(10)
str = str + "index" + Chr(10) + "index" + Chr(10) + "int" + Chr(10) + "ioctl" + Chr(10) + "ioctl" + Chr(10) + "ioctl" + Chr(10) + "join" + Chr(10) + "join" + Chr(10) + "keys" + Chr(10) + "kill" + Chr(10)
str = str + "last" + Chr(10) + "lc" + Chr(10) + "lcfirst" + Chr(10) + "length" + Chr(10) + "link" + Chr(10) + "link" + Chr(10) + "listen" + Chr(10) + "listen" + Chr(10) + "local" + Chr(10) + "localtime" + Chr(10)
str = str + "log" + Chr(10) + "lstat" + Chr(10) + "map" + Chr(10) + "map" + Chr(10) + "mkdir" + Chr(10) + "mkdir" + Chr(10) + "msgctl" + Chr(10) + "msgctl" + Chr(10) + "msgctl" + Chr(10) + "msgget" + Chr(10)
str = str + "msgget" + Chr(10) + "msgrcv" + Chr(10) + "msgrcv" + Chr(10) + "msgrcv" + Chr(10) + "msgrcv" + Chr(10) + "msgrcv" + Chr(10) + "msgsnd" + Chr(10) + "msgsnd" + Chr(10) + "msgsnd" + Chr(10) + "my" + Chr(10)
str = str + "next" + Chr(10) + "no" + Chr(10) + "oct" + Chr(10) + "open" + Chr(10) + "open" + Chr(10) + "opendir" + Chr(10) + "opendir" + Chr(10) + "ord" + Chr(10) + "pack" + Chr(10) + "pack" + Chr(10)
str = str + "package" + Chr(10) + "pipe" + Chr(10) + "pipe" + Chr(10) + "pop" + Chr(10) + "pos" + Chr(10) + "print" + Chr(10) + "printf" + Chr(10) + "push" + Chr(10) + "push" + Chr(10) + "quotemeta" + Chr(10)
str = str + "rand" + Chr(10) + "read" + Chr(10) + "read" + Chr(10) + "read" + Chr(10) + "read" + Chr(10) + "readdir" + Chr(10) + "readlink" + Chr(10) + "recv" + Chr(10) + "recv" + Chr(10) + "recv" + Chr(10)
str = str + "recv" + Chr(10) + "redo" + Chr(10) + "ref" + Chr(10) + "rename" + Chr(10) + "rename" + Chr(10) + "require" + Chr(10) + "reset" + Chr(10) + "return" + Chr(10) + "return" + Chr(10) + "return" + Chr(10)
str = str + "reverse" + Chr(10) + "rewinddir" + Chr(10) + "remdir" + Chr(10) + "scalar" + Chr(10) + "seek" + Chr(10) + "seek" + Chr(10) + "seek" + Chr(10) + "seekdir" + Chr(10) + "seekdir" + Chr(10) + "select" + Chr(10)
str = str + "semctl" + Chr(10) + "semctl" + Chr(10) + "semctl" + Chr(10) + "semctl" + Chr(10) + "semget" + Chr(10) + "semget" + Chr(10) + "semget" + Chr(10) + "semget" + Chr(10) + "semop" + Chr(10) + "semop" + Chr(10)
str = str + "send" + Chr(10) + "send" + Chr(10) + "send" + Chr(10) + "send" + Chr(10) + "setpgrp" + Chr(10) + "setpgrp" + Chr(10) + "setpriority" + Chr(10) + "setpriority" + Chr(10) + "setpriority" + Chr(10) + "setsockopt" + Chr(10)
str = str + "setsockopt" + Chr(10) + "setsockopt" + Chr(10) + "setsockopt" + Chr(10) + "shift" + Chr(10) + "shmctl" + Chr(10) + "shmctl" + Chr(10) + "shmctl" + Chr(10) + "shmget" + Chr(10) + "shmget" + Chr(10) + "shmget" + Chr(10)
str = str + "shmread" + Chr(10) + "shmread" + Chr(10) + "shmread" + Chr(10) + "shmread" + Chr(10) + "shmwrite" + Chr(10) + "shmwrite" + Chr(10) + "shmwrite" + Chr(10) + "shmwrite" + Chr(10) + "shutdown" + Chr(10) + "shutdown" + Chr(10)
str = str + "sin" + Chr(10) + "sleep" + Chr(10) + "socket" + Chr(10) + "socket" + Chr(10) + "socket" + Chr(10) + "socket" + Chr(10) + "socketpair" + Chr(10) + "socketpair" + Chr(10) + "socketpair" + Chr(10) + "socketpair" + Chr(10)
str = str + "socketpair" + Chr(10) + "sort" + Chr(10) + "splice" + Chr(10) + "splice" + Chr(10) + "splice" + Chr(10) + "splice" + Chr(10) + "split" + Chr(10) + "split" + Chr(10) + "split" + Chr(10) + "sprintf" + Chr(10)
str = str + "sprintf" + Chr(10) + "sqrt" + Chr(10) + "srand" + Chr(10) + "stat" + Chr(10) + "study" + Chr(10) + "substr" + Chr(10) + "substr" + Chr(10) + "substr" + Chr(10) + "symlink" + Chr(10) + "symlink" + Chr(10)
str = str + "syscall" + Chr(10) + "sysopen" + Chr(10) + "sysopen" + Chr(10) + "sysopen" + Chr(10) + "sysread" + Chr(10) + "sysread" + Chr(10) + "sysread" + Chr(10) + "sysread" + Chr(10) + "sysseek" + Chr(10) + "sysseek" + Chr(10)
str = str + "sysseek" + Chr(10) + "system" + Chr(10) + "syswrite" + Chr(10) + "syswrite" + Chr(10) + "syswrite" + Chr(10) + "syswrite" + Chr(10) + "tell" + Chr(10) + "telldir" + Chr(10) + "tie" + Chr(10) + "tie" + Chr(10)
str = str + "tie" + Chr(10) + "tied" + Chr(10) + "time" + Chr(10) + "times" + Chr(10) + "truncate" + Chr(10) + "truncate" + Chr(10) + "uc" + Chr(10) + "ucfirst" + Chr(10) + "umask" + Chr(10) + "undef" + Chr(10)
str = str + "unlink" + Chr(10) + "unpack" + Chr(10) + "unpack" + Chr(10) + "unshift" + Chr(10) + "unshift" + Chr(10) + "utime" + Chr(10) + "values" + Chr(10) + "vec" + Chr(10) + "vec" + Chr(10) + "vec" + Chr(10)
str = str + "wait" + Chr(10) + "waitpid" + Chr(10) + "waitpid" + Chr(10) + "wantarray" + Chr(10) + "warn" + Chr(10) + "write" + Chr(10) + "$" + Chr(10)
lang.Keywords = str
lang.Style = cmLangStyleProcedural
str = "=" + Chr(10) + ">" + Chr(10) + "<" + Chr(10) + "/"
lang.Operators = str
lang.SingleLineComments = "# "
lang.ScopeKeywords1 = "{"
lang.ScopeKeywords2 = "}"
lang.StringDelims = Chr$(34) + Chr$(10) + "'"
lang.EscapeChar = "\"
lang.TerminatorChar = ";"
Dim globals As CodeMaxCtl.globals
Set globals = New CodeMaxCtl.globals
Call globals.RegisterLanguage("Perl", lang)

End Sub

Private Sub rt_SelChange(ByVal Control As CodeMaxCtl.ICodeMax)
    
  On Error Resume Next
  Set r = rt.GetSel(True)
  stBar.Panels(4).Text = rt.LineCount
  stBar.Panels(2).Text = r.StartLineNo + 1
  If HighLight = 1 Then
    rt.HighlightedLine = r.EndLineNo
  End If

End Sub

Private Sub save_Click()
    
  doSave

End Sub

Private Sub saveas_Click()
    
  DoSaveAs

End Sub

Private Sub selectall_Click()
    
  On Error Resume Next
  rt.ExecuteCmd cmCmdSelectAll

End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
  On Error Resume Next
  Select Case Button.Key
    Case "undo"
      rt.undo
    Case "redo"
      rt.redo
    Case "cut"
      rt.cut
    Case "copy"
      rt.copy
    Case "paste"
      rt.paste
    Case "delete"
      rt.SelText = ""
    Case "new"
      new_Click
    Case "open"
      open_Click
    Case "save"
      doSave
'    Case "about"
'      frmAbout.Show
    Case "properties"
      Load frmOptions
      frmEditOptions.Show
      frmEditOptions.tabs.tabs(1).Selected = True
    Case "help"
      Shell App.path & "\cedit.exe " & App.path & "\readme.txt", vbNormalFocus
  End Select

End Sub

Private Sub Timedate_Click()
    
  'Dim r As CodeMaxCtl.Range
  'Set r = New CodeMaxCtl.Range
  rt.SelText = Time & " " & Date
  'r = rt.GetSel(True)
  
  'rt.SetSel r, True

End Sub

Private Sub undo_Click()
    
  On Error Resume Next
  rt.undo

End Sub

Public Function doNew()
    
  Dim msg As VbMsgBoxResult
  'Check the document
    
   If Changed = True Then
    msg = MsgBox("Document: " & DocName & Chr(10) & Chr(10) & "This document has been modified." & Chr(10) & "Do you wish to save?", vbYesNoCancel, "Document Modified.")
    Select Case msg
      Case vbCancel
        Exit Function
      Case vbNo
      Case vbYes
        doSave
     End Select
   End If
   rt.Text = ""
   rt.Language = ""
   Changed = False
   DocFile = False
   DocFileName = ""
   DocName = "New Document"
   cboLang.ListIndex = 0
   Me.Caption = "HTML Editor | Taken from the Best Editor cEdit - [New Document]"

End Function

'Private Sub website_Click()
'  OpenBrowser "http://acksoft.hypermart.net", Me.hwnd
'End Sub

Public Function txtProp(num As Long) As Long
    
  Select Case num
    Case 0
      txtProp = 0
      Exit Function
    Case 1
      txtProp = 1
      Exit Function
    Case 2
      txtProp = 2
      Exit Function
    Case 3
      txtProp = 3
      Exit Function
    Case 4
      txtProp = 4
      Exit Function
 End Select
End Function

Private Sub wordcount_Click()
    
   Dim ua, X As Long
   ua = Split(rt.Text, " ")
   X = UBound(ua)
   If X < 0 Then X = 0
   MsgBox "There are:" & Chr(10) & "Words: " & X & Chr(10) & "Characters: " & Len(rt.Text) & Chr(10) & "in this document", vbOKOnly + vbInformation, "Word Count"

End Sub
