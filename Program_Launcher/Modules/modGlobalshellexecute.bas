Attribute VB_Name = "modGlobalshellexecute"
Option Explicit

'declare shell execute
Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMINIMIZED As Long = 2
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SE_Err_NOASSOC = 31
Public Const sOperation As String = "open"     ' Constants for shell operations
Public Const sRun As String = "RUNDLL32.EXE"
Public Const sParameters As String = "shell32.dll,OpenAs_RunDLL "

Public Declare Function _
    ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function _
    GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function shelldoc(sfile As String)
    Dim sPath As String, RetVal As Long, _
        lRet As Long
    lRet = ShellExecute(GetDesktopWindow(), sOperation, sfile, _
        vbNullString, vbNullString, SW_SHOWNORMAL)
    If lRet = SE_Err_NOASSOC Then ' No association exists
    'Create a buffer
    sPath = Space(255)
    'Get the system directory
    RetVal = GetSystemDirectory(sPath, 255)
    'Remove all unnecessary chr$(0)'s
    'and move on the stack
    sPath = Left$(sPath, RetVal)
    
    lRet = ShellExecute(GetDesktopWindow(), "open", sRun, _
        sParameters + sfile, sPath, SW_SHOWNORMAL)
End If
End Function



