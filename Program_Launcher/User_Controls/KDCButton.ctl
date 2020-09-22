VERSION 5.00
Begin VB.UserControl KDCButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3060
   ClipBehavior    =   0  'None
   ControlContainer=   -1  'True
   DefaultCancel   =   -1  'True
   EditAtDesignTime=   -1  'True
   HitBehavior     =   2  'Use Paint
   MaskColor       =   &H00000000&
   MouseIcon       =   "KDCButton.ctx":0000
   MousePointer    =   99  'Custom
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   66
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   204
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   765
      ScaleHeight     =   390
      ScaleWidth      =   2220
      TabIndex        =   0
      Top             =   90
      Width           =   2220
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00CECFCE&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1065
         TabIndex        =   1
         Top             =   90
         Width           =   75
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   30
         Top             =   30
         Width           =   15
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   90
      Top             =   540
   End
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   45
      ScaleHeight     =   390
      ScaleWidth      =   375
      TabIndex        =   2
      Top             =   0
      Width           =   375
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   30
         Top             =   30
         Width           =   105
      End
   End
End
Attribute VB_Name = "KDCButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*******************************************************************
'*  Copyright (C) Kobi Vazana 2001 - All Rights Reserved           *
'*                                                                 *
'*  FILE:  KDCButton.ctl                                           *
'*                                                                 *
'*  DESCRIPTION:                                                   *
'*      Gradient button with color sets that can be modified       *
'*      At Design time ,Centered icon ,and all min Properties      *
'*  Update ver 1.0.5:                                              *
'*      Added Custom Colors And Borders                            *
'*      Added Events MouseIn, MouseOut, KeyDown, KeyPress,         *
'*                   KeyUp, MouseDown, MouseUp ,MouseMove.         *
'*  \MouseIn, MouseOut is giving solution for hover(as requested)\ *
'*                                                                 *
'*  Update ver 1.0.7 :                                             *
'*        XP Appearance                                            *
'*        Thanks To Thomas Braad Hannibal for creating This Part   *
'*  Update ver 1.0.8 :                                             *
'*        New Vertical And Horizontal Gradient (4 ways)            *
'*        Added GradientFill for faster and non flickering buttons *
'*******************************************************************
''--------\\\\\\\--------Declaration---------\\\\\\\\\-----------------
Private Declare Function GradientFill Lib "msimg32" (ByVal hDC As Long, ByRef pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINT_API) As Long

Private Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
    End Type
    
Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
    End Type

Public Enum AppearanceConst
    Flat = 0
    Autumn = 1
    Spring = 2
    Summer = 3
    Winter = 4
    Purple = 5
    Pink = 6
    Blue = 7
    Yellow = 8
    Brown = 9
    GrayOrang = 10
    NeonBlue = 11
    NeonGreen = 12
    HardGray = 13
    SoftGray = 14
    Custom = 15
End Enum

Public Enum GradientStyleConst
    [Vertical Normal & Xp] = 0
    [Horizontal Normal & Xp] = 1
    [V Normal H Xp] = 2
    [H Normal V Xp] = 3
End Enum

Public Enum StyleConst
        Normal = 0
        Xp = 1
End Enum

Private Type POINT_API
    X As Long
    Y As Long
End Type

Private Type RectAPI
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private CusR1 As String, CusG1 As String, CusB1 As String, CusR2 As String, CusG2 As String, CusB2 As String
Private LCol1 As String, Border1 As String, Border2 As String, Top0 As String, Top1 As String
Private Bottom0 As String, Bottom1 As String, CusBorder1 As String, CusBorder2 As String

Dim vert(2) As TRIVERTEX
Dim gRect As GRADIENT_RECT
Private rctToolTip As RectAPI
Private hWndToolTip As Long
Private MyAppearance As AppearanceConst
Private MyGradientStyle As GradientStyleConst
Private MyStyle As StyleConst
Private MyLastEvent As String
Private MyCaption As String
Private MyFont As Font
Private MyForeColor As OLE_COLOR
Private DefForeColor As OLE_COLOR
Private MyBackColorTop As OLE_COLOR
Private MyBackColorBottom As OLE_COLOR
Private MyBorderColorTop As OLE_COLOR
Private MyBorderColorBottom As OLE_COLOR
Private NewButtonIcon As Picture
Private MyEnabled As Boolean
Private MyHasFocus As Boolean
Private MyLeftFocus As Boolean
Private MyToolTipText As String
Private ToolTipAvailable As Boolean
Private GradientF1 As Integer
Private GradientF2 As Integer
Private Const DefToolTipText = vbNullString
Private Const DefBackColorTop = "&HE5E5E5"
Private Const DefBackColorBottom = "&H808080"
Private Const DefBorderColorTop = "&HF5F5F5"
Private Const DefBorderColorBottom = "&H505050"
Private Const MyDefAppearance = Flat
Private Const MyDefStyle = Normal
Private Const DefCaption = "KDC"
Private Const MyDefEnabled = True
Private Const DefGradientStyle = 0



Public Event Click()
Public Event Resize()
Public Event MouseIn()
Public Event MouseOut()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
''----------Start Click KeyDown KeyPress KeyUp GotFocus LostFocus------------
Private Sub Label1_Change()
''don't remove this event this is re size for changing caption
    Call UserControl_Resize
End Sub
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    Call RaiseEventS("Click")
End Sub
Private Sub Label1_Click()
    Call RaiseEventS("Click")
End Sub
Private Sub image1_Click()
    Call RaiseEventS("Click")
End Sub
Private Sub Pic1_Click()
    Call RaiseEventS("Click")
End Sub
Private Sub Pic1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call RaiseEventS("KeyDown", KeyCode, Shift)
End Sub
Private Sub Pic1_KeyPress(KeyAscii As Integer)
    Call RaiseEventS("KeyPress", KeyAscii)
End Sub
Private Sub Pic1_KeyUp(KeyCode As Integer, Shift As Integer)
    Call RaiseEventS("KeyUp", KeyCode, Shift)
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    Call RaiseEventS("KeyPress", KeyAscii)
End Sub
Private Sub UserControl_GotFocus()
    MyHasFocus = True
End Sub
Private Sub UserControl_LostFocus()
    MyHasFocus = False
End Sub
Private Sub Pic1_GotFocus()
    MyHasFocus = True
End Sub
Private Sub Pic1_LostFocus()
    MyHasFocus = False
End Sub
''----------End Click KeyDown KeyPress KeyUp GotFocus LostFocus-------------


''---------------------Start MouseDown MouseUp Resize-----------------------
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, Pic1.Left + (X \ Screen.TwipsPerPixelX), Pic1.Top + (Y \ Screen.TwipsPerPixelY))
    Call AllMouseDown
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, Pic1.Left + (X \ Screen.TwipsPerPixelX), Pic1.Top + (Y \ Screen.TwipsPerPixelY))
    Call AllMouseUP
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, Pic1.Left + (X \ Screen.TwipsPerPixelX), Pic1.Top + (Y \ Screen.TwipsPerPixelY))
    Call AllMouseDown
End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, Pic1.Left + (X \ Screen.TwipsPerPixelX), Pic1.Top + (Y \ Screen.TwipsPerPixelY))
    Call AllMouseUP
End Sub
Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, Pic1.Left + (X \ Screen.TwipsPerPixelX), Pic1.Top + (Y \ Screen.TwipsPerPixelY))
    Call AllMouseDown
End Sub
Private Sub Pic1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, Pic1.Left + (X \ Screen.TwipsPerPixelX), Pic1.Top + (Y \ Screen.TwipsPerPixelY))
    Call AllMouseUP
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MyLeftFocus = False
    Call RaiseEventS("MouseDown", Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MyLeftFocus = False
    Call RaiseEventS("MouseUp", Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
End Sub
Private Sub AllMouseDown()
    UserControl.ScaleMode = 1
    Pic1.Line (0, 0)-(Pic1.Width - 1, Pic1.Height - 1), Border2, B
    Pic1.Line (0, Pic1.Height - 10)-(Pic1.Width, Pic1.Height - 10), Border1
    Pic1.Line (Pic1.Width - 10, 0)-(Pic1.Width - 10, Pic1.Height - 10), Border1
    Image1.Move (Image1.Left + 11), Image1.Top + 11
    Label1.Move (Label1.Left + 20), Label1.Top + 11
End Sub
Private Sub AllMouseUP()
    UserControl.ScaleMode = 1
    Image1.Move (Image1.Left - 11), Image1.Top - 11
    Label1.Move (Label1.Left - 20), Label1.Top - 11
    Call UserControl_Resize
End Sub
Private Sub UserControl_Resize()
On Error Resume Next
UserControl.ScaleMode = 1
    If UserControl.Width <> 0 Then
            If MyStyle = Normal Then
                Image2.Visible = False
                Image1.Visible = True
                Set Image1.Picture = Image2.Picture
                Pic1.Left = 0
                Pic2.Width = 0
                Pic2.Height = 0
                Pic2.Left = -10
                Pic1.Width = UserControl.Width
                Pic1.Height = UserControl.Height
            Else
                Image1.Visible = False
                Image2.Visible = True
                Set Image2.Picture = Image1.Picture
                Pic2.Width = IIf(Image2.Width < 400, 400, Image2.Width + 180)
                Pic1.Left = Pic2.Width
                Pic2.Height = UserControl.Height
                Pic2.Left = 0
                Pic1.Width = (UserControl.Width - (Pic2.Width + 4))
                Pic1.Height = UserControl.Height
                Image2.Top = (Pic1.Height / 2) - (Image2.Height / 2)
                Image2.Left = (Pic2.Width / 2) - ((Image2.Width) / 2)
            End If
                    If Image1.Width > 15 Then
                        If MyStyle = Normal Then
                        Label1.Left = (Pic1.Width / 2) - ((Label1.Width + Image1.Width) / 2) + Image1.Width
                        Else
                        Label1.Left = (Pic1.Width / 2) - ((Label1.Width) / 2) - 40
                        End If
                        Image1.Left = (Pic1.Width / 2) - ((Label1.Width + Image1.Width) / 2) - 40
                    Else
                        Label1.Left = (Pic1.Width / 2) - ((Label1.Width + Image1.Width) / 2)
                    End If
            Label1.Top = (Pic1.Height / 2) - (Label1.Height / 2)
            Image1.Top = (Pic1.Height / 2) - (Image1.Height / 2)
    End If
    Call SetGradient
End Sub
''-------------------------End MouseDown MouseUp Resize---------------------


''--------------------------Start MouseMove Enter Exit----------------------
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, Pic1.Left + (X \ Screen.TwipsPerPixelX), Pic1.Top + (Y \ Screen.TwipsPerPixelY))
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, Pic2.Left + (X \ Screen.TwipsPerPixelX), Pic2.Top + (Y \ Screen.TwipsPerPixelY))
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, Pic1.Left + (X \ Screen.TwipsPerPixelX), Pic1.Top + (Y \ Screen.TwipsPerPixelY))
End Sub
Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, Pic1.Left + (X \ Screen.TwipsPerPixelX), Pic1.Top + (Y \ Screen.TwipsPerPixelY))
End Sub
Private Sub Pic2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, Pic2.Left + (X \ Screen.TwipsPerPixelX), Pic2.Top + (Y \ Screen.TwipsPerPixelY))
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MyLeftFocus = False
    If UserControl.Ambient.UserMode = True And Not Timer1.Enabled Then
        Timer1.Enabled = True
    End If
    UserControl.ScaleMode = 3
    If X >= 0 And Y >= 0 And _
                X <= UserControl.ScaleWidth And Y <= UserControl.ScaleHeight Then
        Call RaiseEventS("MouseIn")
        Call RaiseEventS("MouseMove", Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
    End If
End Sub
Private Sub Timer1_Timer()
    Dim pnt As POINT_API
    UserControl.ScaleMode = 3
    GetCursorPos pnt
    ScreenToClient UserControl.hwnd, pnt
    If pnt.X < UserControl.ScaleLeft Or _
            pnt.Y < UserControl.ScaleTop Or _
            pnt.X > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
            pnt.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
        Timer1.Enabled = False
    
        Call RaiseEventS("MouseOut")
        MyLeftFocus = True
    Else
        MyLeftFocus = False
    End If
End Sub
''-------------------------End MouseMove Enter Exit--------------------------

'----------------------------Start Sending Events----------------------------
Private Function RaiseEventS(ByVal Name As String, ParamArray Params() As Variant)
  Select Case Name
        Case "Click"
            RaiseEvent Click
        Case "KeyDown"
            RaiseEvent KeyDown(CInt(Params(0)), CInt(Params(1)))
        Case "KeyPress"
            RaiseEvent KeyPress(CInt(Params(0)))
        Case "KeyUp"
            RaiseEvent KeyUp(CInt(Params(0)), CInt(Params(1)))
        Case "MouseDown"
            RaiseEvent MouseDown(CInt(Params(0)), CInt(Params(1)), CSng(Params(2)), CSng(Params(3)))
        Case "MouseUp"
            RaiseEvent MouseUp(CInt(Params(0)), CInt(Params(1)), CSng(Params(2)), CSng(Params(3)))
        Case "MouseMove"
            RaiseEvent MouseMove(CInt(Params(0)), CInt(Params(1)), CSng(Params(2)), CSng(Params(3)))
        Case "MouseOut"
            If MyLastEvent <> "MouseOut" Then
                RaiseEvent MouseOut
            End If
            MyLastEvent = Name
        Case "MouseIn"
        If MyLastEvent <> "MouseIn" Then
                RaiseEvent MouseIn
            End If
            MyLastEvent = Name
        Case "Resize"
            RaiseEvent Resize
    End Select
End Function
'----------------------------End Sending Events-----------------------------


''---------------------------Start User Control-----------------------------
Public Sub Refresh()
    UserControl.Refresh
End Sub
Private Sub UserControl_Initialize()
UserControl.ScaleMode = 1
    Pic1.Left = 0
    Pic1.Top = 0
    UserControl.Height = Pic1.Height
    UserControl.Width = Pic1.Width
Call UserControl_Resize
End Sub
Private Sub UserControl_InitProperties()
    Appearance = Yellow
    Style = Normal
    Caption = DefCaption
    ForeColor = DefForeColor
    Set Font = Ambient.Font
    Enabled = MyDefEnabled
    BackColorTop = DefBackColorTop
    BackColorBottom = DefBackColorBottom
    BorderColorTop = DefBorderColorTop
    BorderColorBottom = DefBorderColorBottom
    GradientStyle = DefGradientStyle
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Appearance", MyAppearance, MyDefAppearance)
    Call PropBag.WriteProperty("Style", MyStyle, MyDefStyle)
    Call PropBag.WriteProperty("Caption", MyCaption, DefCaption)
    Call PropBag.WriteProperty("ForeColor", MyForeColor, DefForeColor)
    Call PropBag.WriteProperty("Font", MyFont, Ambient.Font)
    Call PropBag.WriteProperty("ButtonIcon", Me.ButtonIcon, Nothing)
    Call PropBag.WriteProperty("Enabled", MyEnabled, MyDefEnabled)
    Call PropBag.WriteProperty("BackColorTop", MyBackColorTop, DefBackColorTop)
    Call PropBag.WriteProperty("BackColorBottom", MyBackColorBottom, DefBackColorBottom)
    Call PropBag.WriteProperty("BorderColorTop", MyBorderColorTop, DefBorderColorTop)
    Call PropBag.WriteProperty("BorderColorBottom", MyBorderColorBottom, DefBorderColorBottom)
    Call PropBag.WriteProperty("GradientStyle", MyGradientStyle, DefGradientStyle)
    Call PropBag.WriteProperty("ToolTipText", MyToolTipText, DefToolTipText)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Appearance = PropBag.ReadProperty("Appearance", MyDefAppearance)
    Style = PropBag.ReadProperty("Style", MyDefStyle)
    Caption = PropBag.ReadProperty("Caption", DefCaption)
    ForeColor = PropBag.ReadProperty("ForeColor", DefForeColor)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set ButtonIcon = PropBag.ReadProperty("ButtonIcon", Nothing)
    Enabled = PropBag.ReadProperty("Enabled", MyDefEnabled)
    BackColorTop = PropBag.ReadProperty("BackColorTop", DefBackColorTop)
    BackColorBottom = PropBag.ReadProperty("BackColorBottom", DefBackColorBottom)
    BorderColorTop = PropBag.ReadProperty("BorderColorTop", DefBorderColorTop)
    BorderColorBottom = PropBag.ReadProperty("BorderColorBottom", DefBorderColorBottom)
    GradientStyle = PropBag.ReadProperty("GradientStyle", DefGradientStyle)
    ToolTipText = PropBag.ReadProperty("ToolTipText", DefToolTipText)
End Sub

'-----------------------Start Getting Letting Property's---------------------
Public Property Get hwnd() As Long
Attribute hwnd.VB_MemberFlags = "400"
    hwnd = UserControl.hwnd
End Property
Public Property Get Caption() As String
    Caption = MyCaption
End Property
Public Property Let Caption(ByVal vData As String)
    MyCaption = vData
    Label1.Caption = vData
PropertyChanged "Caption"
End Property
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ToolTipText.VB_MemberFlags = "200"
    If ToolTipAvailable Then
        ToolTipText = Extender.ToolTipText
    Else
        ToolTipText = MyToolTipText
    End If
End Property
Public Property Let ToolTipText(ByVal vData As String)
    If ToolTipAvailable Then
        Extender.ToolTipText = vData
    Else
        MyToolTipText = vData
    End If
    PropertyChanged "ToolTipText"
End Property
Public Property Get GradientStyle() As GradientStyleConst
    GradientStyle = MyGradientStyle
End Property
Public Property Let GradientStyle(ByVal vData As GradientStyleConst)
    MyGradientStyle = vData
    Select Case vData
        Case Is = "0"
            GradientF1 = 1
            GradientF2 = 1
        Case Is = "1"
            GradientF1 = 0
            GradientF2 = 0
        Case Is = "2"
            GradientF1 = 1
            GradientF2 = 0
        Case Is = "3"
            GradientF1 = 0
            GradientF2 = 1
    End Select
    Call SetGradient
PropertyChanged "GradientStyle"
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = MyForeColor
End Property
Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    MyForeColor = vData
    Label1.ForeColor = MyForeColor
PropertyChanged "ForeColor"
End Property
Public Property Get Font() As Font
    Set Font = MyFont
End Property
Public Property Set Font(ByVal vData As Font)
    Set MyFont = vData
    Set UserControl.Font = vData
    Set Label1.Font = MyFont
    Call UserControl_Resize
PropertyChanged "Font"
End Property
Public Property Get Style() As StyleConst
    Style = MyStyle
End Property
Public Property Let Style(ByVal vData As StyleConst)
    MyStyle = vData
    Call UserControl_Resize
PropertyChanged "Style"
End Property
Public Property Get Appearance() As AppearanceConst
    Appearance = MyAppearance
End Property
Public Property Let Appearance(ByVal vData As AppearanceConst)
    MyAppearance = vData
    Call SetGradient
    ForeColor = DefForeColor
PropertyChanged "ForeColor"
PropertyChanged "Appearance"
End Property
Public Property Get ButtonIcon() As Picture
        Set ButtonIcon = Image1.Picture
        Set ButtonIcon = Image2.Picture
End Property
Public Property Set ButtonIcon(ByVal NewButtonIcon As Picture)
        Set Image1.Picture = NewButtonIcon
        Set Image2.Picture = NewButtonIcon
  Call UserControl_Resize
PropertyChanged "ButtonIcon"
End Property
Public Property Get Enabled() As Boolean
    Enabled = MyEnabled
End Property
Public Property Let Enabled(ByVal vData As Boolean)
    MyEnabled = vData
    UserControl.Enabled = MyEnabled
    Call UserControl_Resize
PropertyChanged "Enabled"
End Property
Public Property Get BackColorTop() As OLE_COLOR
    BackColorTop = MyBackColorTop
End Property
Public Property Let BackColorTop(ByVal vData As OLE_COLOR)
    MyBackColorTop = vData
    '-------Splitting HEX(OLE COLOR) TO RGB
    Top0 = FixLen(Hex$(vData), "000000")
    Top1 = Left(Top0, 4)
    CusB1 = "&H" & Left(Top1, 2)
    CusG1 = "&H" & Right(Top1, 2)
    CusR1 = "&H" & Right(Top0, 2)
    Call UserControl_Resize
PropertyChanged "BackColorTop"
End Property
Public Property Get BackColorBottom() As OLE_COLOR
    BackColorBottom = MyBackColorBottom
End Property
Public Property Let BackColorBottom(ByVal vData As OLE_COLOR)
    MyBackColorBottom = vData
    '-------Splitting HEX(OLE COLOR) TO RGB
    Bottom0 = FixLen(Hex$(vData), "000000")
    Bottom1 = Left(Bottom0, 4)
    CusB2 = "&H" & Left(Bottom1, 2)
    CusG2 = "&H" & Right(Bottom1, 2)
    CusR2 = "&H" & Right(Bottom0, 2)
    Call UserControl_Resize
PropertyChanged "BackColorBottom"
End Property
Public Property Get BorderColorTop() As OLE_COLOR
    BorderColorTop = MyBorderColorTop
End Property
Public Property Let BorderColorTop(ByVal vData As OLE_COLOR)
    MyBorderColorTop = vData
    CusBorder1 = vData
    Call UserControl_Resize
PropertyChanged "BorderColorTop"
End Property
Public Property Get BorderColorBottom() As OLE_COLOR
    BorderColorBottom = MyBorderColorBottom
End Property
Public Property Let BorderColorBottom(ByVal vData As OLE_COLOR)
    MyBorderColorBottom = vData
    CusBorder2 = vData
    Call UserControl_Resize
PropertyChanged "BorderColorBottom"
End Property
'-------------------------End Getting Letting Property's---------------------


'---------------------------Start Drawing Button Face------------------------
Private Sub SetGradient()
UserControl.ScaleMode = 1
    Select Case MyAppearance
        Case Is = Flat
            DefForeColor = &HFFFFFF
            Border1 = &HE0E0E0
            Border2 = &H606060
            vert(0).Red = &H8000: vert(1).Red = &H8000
            vert(0).Green = &H8000: vert(1).Green = &H8000
            vert(0).Blue = &H8000: vert(1).Blue = &H8000
        Case Is = Autumn
            DefForeColor = &HC0F0F0
            Border1 = &HB0D0D0
            Border2 = &H608080
            vert(0).Red = &HA000: vert(1).Red = &H6000
            vert(0).Green = &HA000: vert(1).Green = &H6000
            vert(0).Blue = &H8000: vert(1).Blue = &H4000
        Case Is = Spring
            DefForeColor = &H80F0C0
            Border1 = &H90D0B0
            Border2 = &H406060
            vert(0).Red = &H8000: vert(1).Red = &H2000
            vert(0).Green = &HA000: vert(1).Green = &H4000
            vert(0).Blue = &H6000: vert(1).Blue = &H0
        Case Is = Summer
            DefForeColor = &H40A0F0
            Border1 = &H7090D0
            Border2 = &H102040
            vert(0).Red = &HD000: vert(1).Red = &H4000
            vert(0).Green = &H6000: vert(1).Green = &H0
            vert(0).Blue = &H4000: vert(1).Blue = &H0
        Case Is = Winter
            DefForeColor = &H804040
            Border1 = &HF08080
            Border2 = &H802040
            vert(0).Red = &HF000: vert(1).Red = &H6000
            vert(0).Green = &HF000: vert(1).Green = &H6000
            vert(0).Blue = &HF000: vert(1).Blue = &H8000
        Case Is = Purple
            DefForeColor = &HC0C0C0
            Border1 = &H908090
            Border2 = &H402040
            vert(0).Red = &HA000: vert(1).Red = &H4000
            vert(0).Green = &H9000: vert(1).Green = &H0
            vert(0).Blue = &HA000: vert(1).Blue = &H4000
        Case Is = Pink
            DefForeColor = &H202080
            Border1 = &HC0C0F0
            Border2 = &H8080A0
            vert(0).Red = &HD000: vert(1).Red = &H8000
            vert(0).Green = &HA000: vert(1).Green = &H7000
            vert(0).Blue = &HA000: vert(1).Blue = &H7000
        Case Is = Blue
            DefForeColor = &HFFFFFF
            Border1 = &HD05050
            Border2 = &H802020
            vert(0).Red = &H2000: vert(1).Red = &H7000
            vert(0).Green = &H2000: vert(1).Green = &H8000
            vert(0).Blue = &H4000: vert(1).Blue = &HA000
        Case Is = Yellow
            DefForeColor = &H206000
            Border1 = &H80FFFF
            Border2 = &H208080
            vert(0).Red = &HF000: vert(1).Red = &HA000
            vert(0).Green = &HF000: vert(1).Green = &HA000
            vert(0).Blue = &H8000: vert(1).Blue = &H2000
        Case Is = Brown
            DefForeColor = &H20F0A0
            Border1 = &H2080F0
            Border2 = &H104080
            vert(0).Red = &HF000: vert(1).Red = &H8000
            vert(0).Green = &HF000: vert(1).Green = &H3000
            vert(0).Blue = &H6000: vert(1).Blue = &H2000
        Case Is = GrayOrang
            DefForeColor = &H80FF&
            Border1 = &H2080F0
            Border2 = &H104080
            vert(0).Red = &HFF00: vert(1).Red = &HCC00
            vert(0).Green = &HFF00: vert(1).Green = &HCC00
            vert(0).Blue = &HFF00: vert(1).Blue = &HCC00
        Case Is = NeonBlue
            DefForeColor = &HFFFFFF
            Border1 = &HF3CD69
            Border2 = &H6B5007
            vert(0).Red = &H2200: vert(1).Red = &H4400
            vert(0).Green = &HCC00: vert(1).Green = &H4400
            vert(0).Blue = &HFF00: vert(1).Blue = &H4400
        Case Is = NeonGreen
            DefForeColor = &HFFFFFF
            Border1 = &HBFBB0D
            Border2 = &H525805
            vert(0).Red = &H2200: vert(1).Red = &H4400
            vert(0).Green = &HCC00: vert(1).Green = &H4400
            vert(0).Blue = &HCC00: vert(1).Blue = &H4400
        Case Is = HardGray
            DefForeColor = &HFFFFFF
            Border1 = &HC0C0C0
            Border2 = &H404040
            vert(0).Red = &H7700: vert(1).Red = &H1100
            vert(0).Green = &H7700: vert(1).Green = &H1100
            vert(0).Blue = &H7700: vert(1).Blue = &H1100
        Case Is = SoftGray
            DefForeColor = &H0
            Border1 = &HC0C0C0
            Border2 = &H4040403
            vert(0).Red = &HEE00: vert(1).Red = &HAA00
            vert(0).Green = &HEE00: vert(1).Green = &HAA00
            vert(0).Blue = &HEE00: vert(1).Blue = &HAA00
        Case Is = Custom
        On Local Error Resume Next
            DefForeColor = &H0
            Border1 = CusBorder1
            Border2 = CusBorder2
            vert(0).Red = CusR1 + "00": vert(1).Red = CusR2 + "00"
            vert(0).Green = CusG1 + "00": vert(1).Green = CusG2 + "00"
            vert(0).Blue = CusB1 + "00": vert(1).Blue = CusB2 + "00"
    End Select

    Pic1.ScaleMode = vbPixels
    Pic2.ScaleMode = vbPixels
    vert(0).X = 0: vert(1).X = Pic1.ScaleWidth
    vert(0).Y = 0: vert(1).Y = Pic1.ScaleHeight
    gRect.UpperLeft = 1
    gRect.LowerRight = 0
'normal
    GradientFill Pic1.hDC, vert(0), 4, gRect, 1, GradientF1
    Pic1.ScaleMode = 1
Pic1.Line (0, 0)-(Pic1.Width - 1, Pic1.Height - 1), Border1, B
Pic1.Line (0, Pic1.Height - 10)-(Pic1.Width, Pic1.Height - 10), Border2
Pic1.Line (Pic1.Width - 10, 0)-(Pic1.Width - 10, Pic1.Height - 10), Border2

'XP
If MyStyle = Xp Then
    GradientFill Pic2.hDC, vert(0), 4, gRect, 1, GradientF2
    Pic2.ScaleMode = 1
Pic2.Line (0, 0)-(Pic2.Width - 1, Pic2.Height - 1), Border1, B
Pic2.Line (0, Pic2.Height - 10)-(Pic2.Width, Pic2.Height - 10), Border2
Pic2.Line (Pic2.Width - 10, 0)-(Pic2.Width - 10, Pic2.Height - 10), Border2
End If
End Sub
'-------------------------Start Drawing Button Face-------------------------


'-------------------Start Adding "000000" To HEX(OLE COLOR)-----------------
Private Function FixLen(ByVal sIn As String, ByVal sMask As String) As String
    If Len(sIn) < Len(sMask) Then
        FixLen = Left$(sMask, Len(sMask) - Len(sIn)) & sIn
    Else
        FixLen = Right$(sIn, Len(sMask))
    End If
End Function
