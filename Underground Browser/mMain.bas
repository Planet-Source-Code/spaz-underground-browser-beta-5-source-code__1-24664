Attribute VB_Name = "mMain"
Option Explicit
' Description:
' A re-usable module for single instance applications which can have a
' a command line.
'
' a) To absolutely prevent two instances, we use a system Mutex via
'    CreateMutex (rather than App.PrevInstance, which may not return True).
'    However this is a pain during development if you press Stop (have to
'    shutdown VB to clear the Mutex) so we just use App.PrevInstance then.
' b) When window is created, it is tagged with a Windows property so any
'    new instances can be accurately identified.
' c) When the user tries to start a second instance (either by double
'    clicking on the EXE or by double clicking an associated file), the
'    window is identified and the command line (if any) is sent to it.
'
' ---------------------------------------------------------------------------
' vbAccelerator - free, advanced source code for VB programmers.
'     http://vbaccelerator.com
' ===========================================================================

Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const ERROR_ALREADY_EXISTS = 183&
Private m_hMutex As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Const SMTO_NORMAL = &H0
Public Const WM_COPYDATA = &H4A
Public Type COPYDATASTRUCT
   dwData As Long
   cbData As Long
   lpData As Long
End Type
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_RESTORE = &HF120
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long


Private m_hWndPrevious As Long
Private m_bInDevelopment As Boolean

' Change this line:
Public Const mcTHISAPPID = "UndergroundBrowser"
Public Sub Main()

   ' Check if this is the first instance:
   If (WeAreAlone(mcTHISAPPID & "_APPLICATION_MUTEX")) Then
      
      ' If it is, then start the app:
      
      ' Change these lines:
     Load frmBrowser
      
   Else
            
      ' There is an existing instance.
      ' First try to find it:
      EnumerateWindows
      
      ' If we get it:
      If (m_hWndPrevious <> 0) Then
      
         ' Do we have a command to send, or is the main window hidden?
         If (Command <> "") Or (IsWindowVisible(m_hWndPrevious) = 0) Then
            
            ' Send.  The app must subclass the WM_COPYDATA message
            ' to get this information:
            Dim tCDS As COPYDATASTRUCT, b() As Byte, lR As Long
            
            If (Command <> "") Then
               b = StrConv(Command, vbFromUnicode)
               tCDS.dwData = 0
               tCDS.cbData = UBound(b) + 1
               tCDS.lpData = VarPtr(b(0))
            Else
               ReDim b(0 To 0) As Byte
               tCDS.dwData = 0
               tCDS.cbData = 1
               tCDS.lpData = VarPtr(b(0))
            End If
            ' Give in if the existing app is not responding:
            lR = SendMessageTimeout(m_hWndPrevious, WM_COPYDATA, 0, tCDS, SMTO_NORMAL, 5000, lR)
      
         Else
            ' Try to activate the existing window:
            RestoreAndActivate m_hWndPrevious
            
         End If
         
      End If
      
   End If
   
End Sub


Public Sub RestoreAndActivate(ByVal hWnd As Long)
   If (IsIconic(hWnd)) Then
      SendMessageByLong hWnd, WM_SYSCOMMAND, SC_RESTORE, 0
   End If
   ActivateWindow hWnd
End Sub

Public Sub TagWindow(ByVal hWnd As Long)
   ' Applies a window property to allow the window to
   ' be clearly identified.
   SetProp hWnd, mcTHISAPPID & "_APPLICATION", 1
End Sub

Public Function IsThisApp(ByVal hWnd As Long) As Boolean
   ' Check if the windows property is set for this
   ' window handle:
   If GetProp(hWnd, mcTHISAPPID & "_APPLICATION") = 1 Then
      IsThisApp = True
   End If
End Function
Public Function EnumWindowsProc( _
        ByVal hWnd As Long, _
        ByVal lParam As Long _
    ) As Long
Dim bStop As Boolean
   ' Customised windows enumeration procedure.  Stops
   ' when it finds another application with the Window
   ' property set, or when all windows are exhausted.
   bStop = False
   If IsThisApp(hWnd) Then
      EnumWindowsProc = 0
      m_hWndPrevious = hWnd
   Else
      EnumWindowsProc = 1
   End If
End Function

Public Function EnumerateWindows() As Boolean
   ' Enumerate top-level windows:
   EnumWindows AddressOf EnumWindowsProc, 0
End Function

Public Sub ActivateWindow(ByVal lHwnd As Long)
    SetForegroundWindow lHwnd
End Sub
Public Function InDevelopment() As Boolean
   ' Debug.Assert code not run in an EXE.  Therefore
   ' m_bInDevelopment variable is never set.
   Debug.Assert InDevelopmentHack() = True
   InDevelopment = m_bInDevelopment
End Function
Public Function InDevelopmentHack() As Boolean
   ' .... '
   m_bInDevelopment = True
   InDevelopmentHack = m_bInDevelopment
End Function
Public Function WeAreAlone(ByVal sMutex As String) As Boolean
   ' Don't call Mutex when in VBIDE because it will apply
   ' for the entire VB IDE session, not just the app's
   ' session.
   If InDevelopment Then
      WeAreAlone = Not (App.PrevInstance)
   Else
      ' Ensures we don't run a second instance even
      ' if the first instance is in the start-up phase
      m_hMutex = CreateMutex(ByVal 0&, 1, sMutex)
      If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
         CloseHandle m_hMutex
      Else
         WeAreAlone = True
      End If
   End If
End Function
Public Function EndApp()
   ' Call this to remove the Mutex.  It will be cleared
   ' anyway by windows, but this ensures it works.
   If (m_hMutex <> 0) Then
      CloseHandle m_hMutex
   End If
   m_hMutex = 0
End Function
