Attribute VB_Name = "modMainDrawMenuBar"
Option Explicit
'for bitmap menu
Public Const MF_BYPOSITION = &H400&

Public Const MIM_BACKGROUND As Long = &H2
Public Const MIM_APPLYTOSUBMENUS As Long = &H80000000
Public Type MENUINFO
    cbSize As Long
    fMask As Long
    dwStyle As Long
    cyMax As Long
    hbrBack As Long
    dwContextHelpID As Long
    dwMenuData As Long
End Type
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, mi As MENUINFO) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

'set bitmaps to menu items
Public Declare Function SetMenuItemBitmaps Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, _
ByVal hBitmapChecked As Long) As Long

Public Sub Draw_MenuBar(frm As Form, MColor As ColorConstants, _
        P1Color As ColorConstants, P2Color As ColorConstants)
    Dim mi As MENUINFO
    With mi
        .cbSize = Len(mi)
        .fMask = MIM_BACKGROUND
        .hbrBack = CreateSolidBrush(MColor)
        SetMenuInfo GetMenu(frm.hWnd), mi 'main menu bar
        .fMask = MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS
        .hbrBack = CreateSolidBrush(P1Color)
        SetMenuInfo GetSubMenu(GetMenu(frm.hWnd), 0), mi '(item 0)
        .hbrBack = CreateSolidBrush(P2Color)
        SetMenuInfo GetSubMenu(GetMenu(frm.hWnd), 1), mi '(item 1)
    End With
End Sub
