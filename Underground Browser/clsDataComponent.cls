VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoadSearchEngines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim db As DAO.Database
Dim rsDefault As DAO.Recordset
Dim rs As DAO.Recordset

Public Function GetURL(Title As String) As String
    
    GetURL = ""
    With rs
        '.Requery
        If .BOF And .EOF Then
            Exit Function
        Else
            .MoveFirst
            Do Until .EOF
                If !Title = Title Then
                    GetURL = !URL
                    Exit Function
                End If
                .MoveNext
            Loop
        End If
    End With

Exit Function



    ' Then UDBErrorHandler "(Class Module) clsLoadSearchEngines::Function GetURL"
    Resume Next
End Function
Public Function GetDefault() As String
    
    GetDefault = ""
    With rs
        '.Requery
        If .BOF And .EOF Then
            Exit Function
        Else
            .MoveFirst
            'GetDefault = !default
        End If
    End With

Exit Function



    ' Then UDBErrorHandler "(Class Module) clsLoadSearchEngines::Function GetDefault"
    Resume Next
End Function

Public Sub AddEngine(Title As String, URL As String)
    
    Dim msg As Integer
    Dim Duplicate As Boolean
    
    'Validate non-blank title
    If Len(Title) = 0 Then
        MsgBox "The title cannot be blank. Please try again"
        Exit Sub
    End If

    'Validate non-blank URL
    If Len(URL) = 0 Then
        MsgBox "The URL cannot be blank. Please try again"
        Exit Sub
    End If
    
    'add to recordset
    With rs
        If .BOF And .EOF Then
        Else
            .MoveFirst
            Do Until .EOF
                If !Title = Title Then
                    Duplicate = True
                    Exit Do
                End If
                .MoveNext
            Loop
        End If
        If Duplicate = False Then
            .AddNew
            !Title = Title
            !URL = URL
            .Update
        End If
    End With

Exit Sub



    ' Then UDBErrorHandler "(Class Module) clsLoadSearchEngines::Sub AddEngine"
    Resume Next
End Sub
Public Sub DeleteEngine(Title As String)
    
    With rs
        If .BOF And .EOF Then
        Else
            .MoveFirst
            Do Until .EOF
                If !Title = Title Then
                    .delete
                    Exit Do
                End If
                .MoveNext
            Loop
        End If
    End With

Exit Sub



    ' Then UDBErrorHandler "(Class Module) clsLoadSearchEngines::Sub DeleteEngine"
    Resume Next
End Sub
Public Sub RefreshLBEngines(LB As ListBox)
    
    LB.Clear
    With rs
        '.Requery
        If .BOF And .EOF Then
        Else
            .MoveFirst
            Do Until .EOF
                LB.AddItem !Title & vbTab & !URL
                .MoveNext
            Loop
        End If
    End With

Exit Sub



    ' Then UDBErrorHandler "(Class Module) clsLoadSearchEngines::Sub RefreshLBEngines"
    Resume Next
End Sub

Public Sub RefreshCBEngines(CB As ComboBox)
    
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

Exit Sub



    ' Then UDBErrorHandler "(Class Module) clsLoadSearchEngines::Sub RefreshCBEngines"
    Resume Next
End Sub

Public Sub ChangeDefault(Title As String)
    
    Dim Duplicate As Boolean
    With rs
        If .BOF And .EOF Then
        Else
            .MoveFirst
            Do Until .EOF
                If !Title = Title Then
                    Duplicate = True
                    Exit Do
                End If
                .MoveNext
            Loop
        End If
    End With
    With rsDefault
        If Duplicate = True Then
            .MoveFirst
            .edit
            !default = Title
            .Update
        End If
    End With

Exit Sub



    ' Then UDBErrorHandler "(Class Module) clsLoadSearchEngines::Sub ChangeDefault"
    Resume Next
End Sub

Private Sub Class_Initialize()
    
    Set db = OpenDatabase(App.path & "\SearchEngines.udb")
    Set rs = db.OpenRecordset("Engines")
    Set rsDefault = db.OpenRecordset("SELECT Default FROM Default")

Exit Sub



    ' Then UDBErrorHandler "(Class Module) clsLoadSearchEngines::Sub Class_Initialize"
    Resume Next
End Sub

