Attribute VB_Name = "modGlobalgetfilename"
Option Explicit

'delete file
Public Declare Function DeleteFile _
    Lib "kernel32" _
    Alias "DeleteFileA" ( _
    ByVal lpFileName As String) _
    As Long

'check if a file exists
Public Function FileExists(FileName As String) As Boolean
    
    Dim intFreeFile As Integer
    On Error GoTo ErrorHandler
    
    intFreeFile = FreeFile
    Open FileName For Input As #intFreeFile
    Close #intFreeFile
    FileExists = True
    Exit Function
    
ErrorHandler:
    FileExists = False
End Function

'remove path to display only filename
Public Function RemovePath(ByVal FileName As String) As String
    Dim pos As Long
    pos = InStrRev(FileName, "\")
    If pos > 0 Then
        RemovePath = Right$(FileName, Len(FileName) - pos)
    Else
        RemovePath = FileName
    End If
End Function
