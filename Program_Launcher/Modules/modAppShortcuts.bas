Attribute VB_Name = "modAppShortcuts"
Option Explicit

'lock workstation
Public Declare Function LockWorkStation Lib "user32.dll" () As Long

Public Function LaunchApp(ByVal stxt As String, frm As Form) As String

    Dim lRet As Long
    Dim sPath As String
    Dim RetVal As Long
    
    lRet = ShellExecute(frm.hwnd, "Open", stxt, _
        vbNullString, "C:\", SW_SHOWNORMAL)
    
    If lRet <= 32 Then
        Dim aButtons(1) As String
        Dim Reply As Integer
        aButtons(0) = "Cancel to discard ..."
        aButtons(1) = "Yes browse for runner ..."
        Reply = MsgBoxEx("Unable to Run Application " & vbCrLf _
            & vbCrLf & stxt & vbCrLf & vbCrLf _
            & "Either the target path is wrong," & vbCrLf & vbCrLf _
            & "or there is no program to open the shortcut." & vbCrLf & vbCrLf _
            & "Right click on window or drag target application." & vbCrLf & vbCrLf _
            & "To open the file with another program click Yes ...", _
            vbCustom, "Error on Running the Shortcut ... >>", , , aButtons, _
            aButtons(1), 600, 400, 400, 230, vbBlue, vbWhite)
        
        Select Case Reply
            Case 0
                'MsgBox "First button clicked"
                If lRet = SE_Err_NOASSOC Then ' No association exists
                sPath = Space(255) 'Create a buffer
                RetVal = GetSystemDirectory(sPath, 255) 'Get the system directory
                'Remove all unnecessary chr$(0)'s
                'and move on the stack
                sPath = Left$(sPath, RetVal)
                lRet = ShellExecute(GetDesktopWindow(), "Open", sRun, _
                    sParameters + stxt, sPath, SW_SHOWNORMAL)
                End If
    
            Case 1
                Exit Function
            Case 2
                Exit Function
        End Select
        
        Exit Function
    End If
            
End Function

Public Function BrowseApp(ByVal stxt As String, _
ByVal slbl As String, ByVal sn As String) As String
On Error GoTo ErrorHandler
    Dim frmCDC As New frmCDC
    'frmCDC.Move (Screen.Width - frmCDC.Width) / 2, (Screen.Height - frmCDC.Height) / 2
    With frmCDC.CommonDialog1
        .CancelError = True
        .DialogTitle = "Choose Application"
        .Filter = "Program Files (*.exe)| *.exe|All Files (*.*)|*.*"
        .FilterIndex = 1
        .flags = cdlOFNNoValidate Or cdlOFNHideReadOnly Or _
            cdlOFNFileMustExist Or cdlOFNPathMustExist Or _
            cdlOFNNoChangeDir
        .ShowOpen
        stxt = .FileName   'put path to text 1
        slbl = .FileTitle ' put file title to label 1
        'save path to registry app path and app name
        'SaveSetting App.EXEName, sn, "Path", stxt
        'SaveSetting App.EXEName, sn, "Label", slbl
        
        Call WriteIniString(sn, "Path", stxt, inifileapp)
        Call WriteIniString(sn, "Label", slbl, inifileapp)

    End With
    Unload frmCDC
    Set frmCDC = Nothing
    Exit Function
ErrorHandler:
    If Err.Number <> cdlCancel Then
        MsgBox Err.Number & " - " & Err.Description, _
            vbOKOnly + vbExclamation
    End If
    Unload frmCDC
    Set frmCDC = Nothing
End Function


Public Function AppDD(ByVal stxt As String, _
ByVal slbl As String, ByVal sn As String, _
ByVal data As DataObject) As String
    On Error Resume Next
    
    Dim strTemp As String
    Dim i As Long
    
    If data.Files.count > 0 Then
        stxt = vbNullString
        slbl = vbNullString
        
        For i = 1 To data.Files.count
            'retrieves path and write to textbox
            stxt = stxt & data.Files(i)
            'retrives the filename from the first file
            strTemp = RemovePath(data.Files(i))
            slbl = strTemp
        Next i
        
    End If
    'save path to registry app path and app name
    'SaveSetting App.EXEName, sn, "Path", stxt
    'SaveSetting App.EXEName, sn, "Label", slbl

    Call WriteIniString(sn, "Path", stxt, inifileapp)
    Call WriteIniString(sn, "Label", slbl, inifileapp)
        
End Function
