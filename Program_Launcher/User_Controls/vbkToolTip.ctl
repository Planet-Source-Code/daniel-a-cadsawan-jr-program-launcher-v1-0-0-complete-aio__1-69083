VERSION 5.00
Begin VB.UserControl vbkToolTip 
   Appearance      =   0  'Flat
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1545
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1185
   ScaleWidth      =   1545
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   480
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Custom ToolTip"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "vbkToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
   
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WS_BORDER = &H800000
Private Const WS_POPUP = &H80000000
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const GWL_STYLE = (-16)

Private Const SW_HIDE = 0
Private Const SW_NORMAL = 1
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOWNA = 8

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Dim P As POINTAPI
Dim PP As POINTAPI
Dim OldHWnd As Long
Dim NewHWnd As Long
Dim NHWnd As Long

Dim HW As Long

Private Sub Timer1_Timer()
    GetCursorPos PP
    NHWnd = WindowFromPoint(PP.X, PP.Y)
    If NHWnd <> OldHWnd Then
        OldHWnd = -1
        Timer1.Enabled = False
        ShowWindow HW, SW_HIDE
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode Then
        HW = frmTip.hwnd
        SetWindowLong HW, GWL_STYLE, WS_POPUP Or WS_BORDER
        ShowWindow HW, SW_HIDE
    End If
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = lblCaption.Width
    UserControl.Height = lblCaption.Height
End Sub

Private Sub UserControl_Terminate()
    DestroyWindow HW
End Sub

Public Sub Fire(obj As Object, Caption As String, Text As String, pxWidth As Long, pxHeight As Long)
    GetCursorPos P
    NewHWnd = WindowFromPoint(P.X, P.Y)
    If NewHWnd <> OldHWnd Then
        Timer1.Enabled = True
        OldHWnd = NewHWnd
        MoveWindow HW, P.X - 150, P.Y, pxWidth, pxHeight, 1
        ShowWindow HW, SW_SHOWNA
    Else
        frmTip.Init Caption, Text, pxWidth, pxHeight
        MoveWindow HW, P.X - 150, P.Y, pxWidth, pxHeight, 1
    End If
End Sub

Public Sub HideTip()
    OldHWnd = -1
    Timer1.Enabled = False
    ShowWindow HW, SW_HIDE
End Sub
