Attribute VB_Name = "modAppSubs"
Option Explicit
Public Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" _
(ByVal lpPath As String) As Long

'form load subs

Public Sub AppInstance()
    
    Dim strCap As String
    Dim lngHwnd As Long
    Dim lngRet As Long
    'Store the existing caption in a temp var
    'and change it to something else as
    'we are going to use the Caption to find
    'an earlier instance of this application
    strCap = frmProgramLauncher.Caption
    frmProgramLauncher.Caption = "*" & strCap
    'Find the window of the preveious instance
    'using the original Caption
    lngHwnd = FindWindow(ByVal vbNullString, ByVal strCap)
    'This means App.PrevInstance is True
    If lngHwnd <> 0 Then
        MsgBox App.Title & " is already running!", vbInformation
        'Send a Close message to that window
        PostMessage lngHwnd, WM_CLOSE, 0&, 0&
        
    End If
    'Reset the Caption, so that the next version of the application
    'can find this application
    frmProgramLauncher.Caption = strCap
    
End Sub

' This gets gradient settings from registry
Public Sub Gradient_Settings()
On Error GoTo FirstLoad:
    Dim FirstLoad As Integer
    
    Dim Red1 As Integer
    Dim Green1 As Integer
    Dim Blue1 As Integer
    Dim Red2 As Integer
    Dim Green2 As Integer
    Dim Blue2 As Integer
    
    Red1 = GetSetting(App.EXEName, _
        "Gradient\Red1", "Value", "")
    Green1 = GetSetting(App.EXEName, _
        "Gradient\Green1", "Value", "")
    Blue1 = GetSetting(App.EXEName, _
        "Gradient\Blue1", "Value", "")
    
    Red2 = GetSetting(App.EXEName, _
        "Gradient\Red2", "Value", "")
    Green2 = GetSetting(App.EXEName, _
        "Gradient\Green2", "Value", "")
    Blue2 = GetSetting(App.EXEName, _
        "Gradient\Blue2", "Value", "")
    
    FirstLoad = GetSetting(App.EXEName, _
        "Gradient\Red1", "Value", "")
    
    If FirstLoad = vbNullString Then GoTo FirstLoad:
    PaintGradient frmProgramLauncher, Red1, Green1, Blue1, Red2, Green2, Blue2
    'PaintGradient frmProgramLauncher, 255, 255, 255, 128, 128, 255
    Exit Sub
    
FirstLoad:
    PaintGradient frmProgramLauncher, 255, 255, 255, 128, 128, 255
    SaveSetting App.EXEName, "Gradient\Red1", "Value", "255"
    SaveSetting App.EXEName, "Gradient\Green1", "Value", "255"
    SaveSetting App.EXEName, "Gradient\Blue1", "Value", "255"
    SaveSetting App.EXEName, "Gradient\Red2", "Value", "128"
    SaveSetting App.EXEName, "Gradient\Green2", "Value", "128"
    SaveSetting App.EXEName, "Gradient\Blue2", "Value", "255"
    
End Sub

Public Sub Draw_MenuBmp()
    Dim hMenu As Long, hSubMenu As Long
    'get the handle of the menu
    hMenu = GetMenu(frmProgramLauncher.hwnd)
    'check if there's a menu
    If hMenu = 0 Then
        MsgBox "This form doesn't have a menu!", vbInformation
        Exit Sub
    End If
    'get the first submenu
    hSubMenu = GetSubMenu(hMenu, 0)
    'check if there's a submenu
    If hSubMenu = 0 Then
        MsgBox "This form doesn't have a submenu!", vbInformation
        Exit Sub
    End If
    'set the menu bitmap
    SetMenuItemBitmaps hSubMenu, 0, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(1).Picture, frmProgramLauncher.Imagebmp.ListImages(1).Picture
    SetMenuItemBitmaps hSubMenu, 1, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(2).Picture, frmProgramLauncher.Imagebmp.ListImages(2).Picture
    SetMenuItemBitmaps hSubMenu, 2, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(3).Picture, frmProgramLauncher.Imagebmp.ListImages(3).Picture
    SetMenuItemBitmaps hSubMenu, 3, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(4).Picture, frmProgramLauncher.Imagebmp.ListImages(4).Picture
    SetMenuItemBitmaps hSubMenu, 4, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(5).Picture, frmProgramLauncher.Imagebmp.ListImages(5).Picture
    
    SetMenuItemBitmaps hSubMenu, 6, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(6).Picture, frmProgramLauncher.Imagebmp.ListImages(6).Picture
    
    SetMenuItemBitmaps hSubMenu, 8, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(7).Picture, frmProgramLauncher.Imagebmp.ListImages(7).Picture
    SetMenuItemBitmaps hSubMenu, 9, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(8).Picture, frmProgramLauncher.Imagebmp.ListImages(8).Picture
    SetMenuItemBitmaps hSubMenu, 10, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(9).Picture, frmProgramLauncher.Imagebmp.ListImages(9).Picture
    SetMenuItemBitmaps hSubMenu, 11, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(10).Picture, frmProgramLauncher.Imagebmp.ListImages(10).Picture
    SetMenuItemBitmaps hSubMenu, 12, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(11).Picture, frmProgramLauncher.Imagebmp.ListImages(11).Picture

    SetMenuItemBitmaps hSubMenu, 14, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(12).Picture, frmProgramLauncher.Imagebmp.ListImages(12).Picture
    SetMenuItemBitmaps hSubMenu, 15, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(13).Picture, frmProgramLauncher.Imagebmp.ListImages(13).Picture
    SetMenuItemBitmaps hSubMenu, 16, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(14).Picture, frmProgramLauncher.Imagebmp.ListImages(14).Picture
    SetMenuItemBitmaps hSubMenu, 17, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(15).Picture, frmProgramLauncher.Imagebmp.ListImages(15).Picture

    SetMenuItemBitmaps hSubMenu, 19, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(16).Picture, frmProgramLauncher.Imagebmp.ListImages(16).Picture
    SetMenuItemBitmaps hSubMenu, 20, MF_BYPOSITION, frmProgramLauncher.Imagebmp.ListImages(17).Picture, frmProgramLauncher.Imagebmp.ListImages(17).Picture
End Sub

