Attribute VB_Name = "modMsgIniFile"
Option Explicit
Public inifilename As String
Public inifileapp As String
Public inifilepw As String

Public iNum As Integer

Public Const ACTIVEFORMS = "ActiveForms"

Public Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nsize As Long, _
    ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias _
    "WritePrivateProfileSectionA" (ByVal lpAppName As String, _
    ByVal lpString As String, ByVal lpFileName As String) As Long

Public Sub GetFormsInformation()
Dim i As Integer
    Dim F As Form
    Dim FormName As String
    
    Do While True
        i = i + 1
        FormName = GetIniString(ACTIVEFORMS, "Form" & i, vbNullString, inifilename)
        If FormName = vbNullString Then Exit Do
        
        Set F = New frmMsg
        
        With F
            '.Tag = FormName
            .Caption = GetIniString(FormName, "Caption", vbNullString, inifilename)
            '.show
            .Top = GetIniString(FormName, "Top", vbNullString, inifilename)
            .Left = GetIniString(FormName, "Left", vbNullString, inifilename)
            .Height = GetIniString(FormName, "Height", vbNullString, inifilename)
            .Width = GetIniString(FormName, "Width", vbNullString, inifilename)
            .Tag = FormName ' this produce another note message
            .show ' this produce another note message
            
            'i added this
            .RTBNote.Font.Name = GetIniString(FormName, "FontofNote", vbNullString, inifilename)
            .RTBNote.Font.Size = GetIniString(FormName, "SizeofNote", vbNullString, inifilename)
            .RTBNote.Font.Bold = GetIniString(FormName, "BoldofNote", vbNullString, inifilename)
            .RTBNote.Font.Italic = GetIniString(FormName, "ItalicofNote", vbNullString, inifilename)
            .RTBNote.Font.Underline = GetIniString(FormName, "UnderlineofNote", vbNullString, inifilename)
            .RTBNote.Font.Strikethrough = GetIniString(FormName, "StrikethruofNote", vbNullString, inifilename)
            .RTBNote.BackColor = GetIniString(FormName, "BackColorofNote", vbNullString, inifilename)
            .picMin.Visible = GetIniString(FormName, "Maximize", vbNullString, inifilename)
            .picFont.Visible = GetIniString(FormName, "picFontVis", vbNullString, inifilename)
            .picColor.Visible = GetIniString(FormName, "picColorVis", vbNullString, inifilename)
            .picMove2.Visible = GetIniString(FormName, "picMove2Vis", vbNullString, inifilename)
            .picDrag.Visible = GetIniString(FormName, "picDragVis", vbNullString, inifilename)
            .picTransOff.Visible = GetIniString(FormName, "picTransOff", vbNullString, inifilename)
            .picSkinOff.Visible = GetIniString(FormName, "SkinOff", vbNullString, inifilename)
            '.imgSkin.Picture = GetIniString(FormName, "SkinFile", vbNullString, inifilename)
            
            Dim skin As String
            skin = GetSetting(App.EXEName, "NoteSkin" & "\" & .Tag, "Value", "")
            'skin = GetIniString(FormName, "SkinFile", vbNullString, inifilename)
            Select Case skin
                Case 1
                    .imgSkin.Picture = .ImgList.ListImages(1).Picture
                Case 2
                    .imgSkin.Picture = .ImgList.ListImages(2).Picture
                Case 3
                    .imgSkin.Picture = .ImgList.ListImages(3).Picture
                Case vbNullString
                 '   .imgSkin.Picture = .ImgList.ListImages(3).Picture
                    MsgBox "Some of the Notes background image were not found." & vbCrLf _
                    & "Some notes will be loaded with the default " & vbCrLf _
                    & "background setting. Loading continues ...." & .Tag, vbInformation, "Note Loading Info"
                    .imgSkin.Picture = .ImgList.ListImages(3).Picture
                Case Else
                    .imgSkin.Picture = LoadPicture(skin)
            End Select
            
            Dim sTemp As String
            Open App.Path & "\Data\" & .Tag & ".TextRTF" For Input As #1
            sTemp = Input(LOF(1), 1)    'Getting the text
            Close #1                        'Closing the file
            .RTBNote.TextRTF = sTemp
            
        End With
    Loop
    
End Sub


Public Sub SaveFormInformation(frm As Form, FormSeq As Integer)
    
    With frm
        If .Tag <> "" Then
            .Tag = .Name & Format(FormSeq, "00")
            .Caption = .Tag
            
            Call WriteIniString(ACTIVEFORMS, "Form" & FormSeq, .Tag, inifilename)
            ' Erase whole section to re-wrtite it
            
            Call WritePrivateProfileSection(.Tag, vbNullString, inifilename)
            Call WriteIniString(.Tag, "Top", .Top, inifilename)
            Call WriteIniString(.Tag, "Left", .Left, inifilename)
            Call WriteIniString(.Tag, "Height", .Height, inifilename)
            Call WriteIniString(.Tag, "Width", .Width, inifilename)
            Call WriteIniString(.Tag, "Caption", .Caption, inifilename)
            'i added this
            Call WriteIniString(.Tag, "FontofNote", .RTBNote.Font.Name, inifilename)
            Call WriteIniString(.Tag, "SizeofNote", .RTBNote.Font.Size, inifilename)
            Call WriteIniString(.Tag, "BoldofNote", .RTBNote.Font.Bold, inifilename)
            Call WriteIniString(.Tag, "ItalicofNote", .RTBNote.Font.Italic, inifilename)
            Call WriteIniString(.Tag, "UnderlineofNote", .RTBNote.Font.Underline, inifilename)
            Call WriteIniString(.Tag, "StrikethruofNote", .RTBNote.Font.Strikethrough, inifilename)
            Call WriteIniString(.Tag, "BackColorofNote", .RTBNote.BackColor, inifilename)
            Call WriteIniString(.Tag, "Maximize", .picMin.Visible, inifilename)
            Call WriteIniString(.Tag, "picFontVis", .picFont.Visible, inifilename)
            Call WriteIniString(.Tag, "picColorVis", .picColor.Visible, inifilename)
            Call WriteIniString(.Tag, "picMove2Vis", .picMove2.Visible, inifilename)
            Call WriteIniString(.Tag, "picDragVis", .picDrag.Visible, inifilename)
            Call WriteIniString(.Tag, "picTransOff", .picTransOff.Visible, inifilename)
            Call WriteIniString(.Tag, "SkinOff", .picSkinOff.Visible, inifilename)
            
            If .imgSkin.Picture = .ImgList.ListImages(1).Picture Then _
                SaveSetting App.EXEName, "NoteSkin" & "\" & .Tag, "Value", "1"
            If .imgSkin.Picture = .ImgList.ListImages(2).Picture Then _
                SaveSetting App.EXEName, "NoteSkin" & "\" & .Tag, "Value", "2"
            If .imgSkin.Picture = .ImgList.ListImages(3).Picture Then _
                SaveSetting App.EXEName, "NoteSkin" & "\" & .Tag, "Value", "3"
            If .picSkinOff.Visible = True Then _
                SaveSetting App.EXEName, "NoteSkin" & "\" & .Tag, "Value", "3"
            
            'If .imgSkin.Picture = .ImgList.ListImages(1).Picture Then _
            'Call WriteIniString(.Tag, "SkinFile", "1", inifilename)
            'If .imgSkin.Picture = .ImgList.ListImages(2).Picture Then _
            'Call WriteIniString(.Tag, "SkinFile", "2", inifilename)
            'If .imgSkin.Picture = .ImgList.ListImages(3).Picture Or .picSkinOff.Visible = True Then _
            'Call WriteIniString(.Tag, "SkinFile", "3", inifilename)
            
            Open App.Path & "\Data\" & .Tag & ".TextRTF" For Output As #1
                Print #1, .RTBNote.TextRTF
            Close #1
        End If
    End With
End Sub


Public Sub WriteIniString(Section As String, Key As String, _
        Value As String, IniFile As String)
    
    Call WritePrivateProfileString(Section, Key, Value, IniFile)
    
End Sub

Public Function GetIniString(Section As String, Key As String, _
        Default As String, IniFile As String) As String
    Dim sValue As String
    Dim lth As Integer
    Dim i As Integer
    
    
    lth = 1000 '255
    sValue = String(lth, Chr(0))
    Call GetPrivateProfileString(Section, Key, Default, sValue, lth, IniFile)
    i = InStr(sValue, Chr(0))
    If i = 0 Then
        GetIniString = vbNullString
    Else
        GetIniString = Mid(sValue, 1, i - 1)
    End If
    
End Function
