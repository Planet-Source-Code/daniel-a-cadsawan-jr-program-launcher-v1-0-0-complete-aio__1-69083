Attribute VB_Name = "modMainextractassoc_icon"
Option Explicit

'declare extract icon
Public Declare Function _
    ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" _
    (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long

Public Declare Function DrawIconEx Lib "user32" _
    (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, _
    ByVal hIcon As Long, ByVal cxWidth As Long, _
    ByVal cyWidth As Long, ByVal istepIfAniCur As Long, _
    ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Public Declare Function DestroyIcon Lib "user32" _
    (ByVal hIcon As Long) As Long
