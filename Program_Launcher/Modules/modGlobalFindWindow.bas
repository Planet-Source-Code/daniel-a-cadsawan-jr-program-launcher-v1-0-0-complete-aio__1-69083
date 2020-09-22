Attribute VB_Name = "modGlobalFindWindow"
Option Explicit

' compose of make window top most nad not top most
' enable - disable task manager
' find instance of app
' capture and release
' make form transparent

' to make top most not topmost
Public Const HWND_BOTTOM = 1
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

' API Calls Used To Move A Form With The Mousemove
' borderless form
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_CLOSE = &H10

'API Calls Used To Remove The Title Bar From Window
'(Make A Sizeable Borderless Form)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const WS_DLGFRAME = &H400000
'Requires Windows 2000 or later:
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT = &H20&
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const BM_SETSTATE = &HF3

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Public Declare Function GetActiveWindow Lib "user32" _
    () As Long

Public Declare Function SendMessage Lib "user32.dll" _
    Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

' Declare App Instance Function

Public Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Public Declare Function PostMessage Lib "user32" _
    Alias "PostMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

' capture and release
Public Declare Function _
    ReleaseCapture Lib "user32" () As Long
Public Declare Function _
    SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function _
    GetCapture Lib "user32" () As Long

'DECLARATION
'API Calls Used To Remove The Title Bar From Window
'(Make A Sizeable Borderless Form)
Public Declare Function _
    GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Declare Function _
    SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Public Declare Function SetLayeredWindowAttributes Lib "user32" _
    (ByVal hwnd As Long, ByVal crKey As Long, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Declare Function SendMessageAsLong Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

' code: SetForegroundWindow Me.hwnd
Public Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long

' lock window update: on resize
Public Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwnd As Long) As Long

Public Sub LockWindow(hwnd As Long, blnValue As Boolean)
    If blnValue Then
        LockWindowUpdate hwnd
    Else
        LockWindowUpdate 0
    End If
End Sub

Public Sub MakeTopMost(hwnd As Long)
    ' Make Always on Top
    SetWindowPos hwnd, HWND_TOPMOST, _
        0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub MakeNormal(hwnd As Long)
    ' Make Normal
    SetWindowPos hwnd, HWND_NOTOPMOST, _
        0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub tskbar(show As Boolean)
    Dim primo As Long
    primo = FindWindow("Shell_traywnd", "")
    If show = True Then
        SetWindowPos primo, 0, 0, 0, 0, 0, &H40 'sets the task bar to its original position
    Else
        SetWindowPos primo, 0, 0, 0, 0, 0, &H80  'this hides the taskbar making it invisible
    End If
End Sub

'make windows transparent
Public Sub MakeWindowTransparent(ByVal hwnd As Long, ByVal alphaAmount As Byte)
    Dim lStyle As Long
    
    lStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    lStyle = lStyle Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, lStyle
    SetLayeredWindowAttributes hwnd, 0, alphaAmount, LWA_ALPHA
End Sub

Public Sub SetTrans(hwnd As Long, Trans As Integer)
    Dim Tcall As Long
    Tcall = GetWindowLong(hwnd, GWL_EXSTYLE)
    SetWindowLong hwnd, GWL_EXSTYLE, Tcall Or WS_EX_LAYERED
    SetLayeredWindowAttributes hwnd, RGB(255, 255, 0), Trans, LWA_ALPHA
    Exit Sub
End Sub
