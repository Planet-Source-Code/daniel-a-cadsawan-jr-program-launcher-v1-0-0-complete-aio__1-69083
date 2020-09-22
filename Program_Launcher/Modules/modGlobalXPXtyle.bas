Attribute VB_Name = "modGlobalXPXtyle"
Option Explicit

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
    ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" ( _
    ByVal hLibModule As Long) As Long
Public m_hMod As Long

