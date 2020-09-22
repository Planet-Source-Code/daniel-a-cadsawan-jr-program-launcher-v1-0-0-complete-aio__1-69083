VERSION 5.00
Begin VB.Form frmSlideNote 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   855
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ProgramLauncher.xFrame xFrameSlide 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1508
      Caption         =   "Slide Note : Click here to unload note."
      Enabled         =   -1  'True
      EnableGradient  =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      HeaderGradientBottom=   12611136
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFC0&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         Picture         =   "frmSlideNote.frx":0000
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblSlideNote 
         BackStyle       =   0  'Transparent
         Caption         =   "This is your Hot Spot for Slide In and Out."
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmSlideNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Form_Activate()
    On Error Resume Next
 '   Dim i As Long
 '   Static bActive As Boolean
 '   If bActive Then Exit Sub
 '   bActive = True
 '   For i = 1 To 255 Step 2 '2
 '       MakeWindowTransparent frmSlideNote.hWnd, i
 '       DoEvents    ' need this so form doesn't turn black
 '   Next
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    xFrameSlide.Width = frmSlideNote.Width
    xFrameSlide.Height = frmSlideNote.Height
    
    Dim Ftop As Integer
    Dim Fheight As Integer
    'get main forms window position
    Ftop = val(GetSetting(App.EXEName, "WinPos", "PrLrTop", ""))
    Fheight = val(GetSetting(App.EXEName, "WinPos", "PrLrHeight", ""))
    
    With Me
        .Left = (Screen.Width - 400 - frmSlideNote.Width)
        .Top = Ftop + Fheight - 1000
    End With
    ' start off transparent so form doesn't flicker
    'MakeWindowTransparent frmSlideNote.hWnd, 2 '10
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'CODE : fade form on unload
    Dim X As Long
    For X = 255 To 0 Step -2 '-2
        MakeWindowTransparent frmSlideNote.hWnd, X
    Next
    'Unload our form completely
    Unload frmSlideNote
End Sub

Private Sub xFrameSlide_Click()
    On Error Resume Next
    Unload frmSlideNote
End Sub
