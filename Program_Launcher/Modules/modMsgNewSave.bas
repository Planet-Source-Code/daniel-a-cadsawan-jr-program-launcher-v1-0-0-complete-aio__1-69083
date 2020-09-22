Attribute VB_Name = "modMsgNewSave"
Option Explicit

Public Sub CreateNewForm()
On Error GoTo ErrorHandler

    Dim FormName As String
    Dim frm As Form
    Dim X As Long
    Dim Y As Long
    Dim sName As String
    
    iNum = iNum + 1
    If iNum = 100 Then iNum = 1
    'for 2 characters file name
    sName = frmMsg.Name & Format(iNum, "00")
    
    ' Create forms
    Set frm = New frmMsg
    ' when not loading alright this is what i changed!
    FormName = sName '"Form" & sName
    X = X + 100
    Y = Y + 100
    
    With frm
        .RTBNote.Text = " Note:"
        .Left = .Left + X
        .Top = .Top + Y
        .Caption = FormName
        .Tag = FormName
        .show
        
        'saves the text to file
        Open App.Path & "\Data\" & sName & ".TextRTF" For Output As #1
            Print #1, .RTBNote.Text
        Close #1

    End With
Exit Sub
ErrorHandler:
    MsgBox " System Error Number " & Err.Number _
        & " : " & Err.Description, vbInformation
End Sub

Public Sub SaveOpenForms()
On Error GoTo ErrorHandler

    Dim frm As Form
    

    If Not Len(Dir$(inifilename)) = 0 Then
       Open inifilename For Output As #1
        Close #1
    End If
    
    ' Erase whole section to re-write it
    Call WritePrivateProfileSection(ACTIVEFORMS, vbNullString, inifilename)
    
    ' Save Form Information
    iNum = 0
    For Each frm In Forms
        If frm.Tag <> "" Then
            iNum = iNum + 1
            SaveFormInformation frm, iNum
            Unload frm
            Set frm = Nothing
        End If
    Next

Exit Sub
ErrorHandler:
    MsgBox " System Error Number " & Err.Number _
        & " : " & Err.Description
End Sub

