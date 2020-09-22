VERSION 5.00
Begin VB.Form frmGradientForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   14760
   ClientTop       =   6690
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   2880
      Width           =   615
   End
   Begin VB.VScrollBar vsbBlue2 
      Height          =   2415
      Left            =   1560
      Max             =   255
      MouseIcon       =   "frmGradientForm.frx":0000
      TabIndex        =   6
      Top             =   360
      Value           =   128
      Width           =   255
   End
   Begin VB.VScrollBar vsbGreen2 
      Height          =   2415
      Left            =   1320
      Max             =   255
      MouseIcon       =   "frmGradientForm.frx":0CCA
      TabIndex        =   5
      Top             =   360
      Value           =   128
      Width           =   255
   End
   Begin VB.VScrollBar vsbRed2 
      Height          =   2415
      Left            =   1080
      Max             =   255
      MouseIcon       =   "frmGradientForm.frx":1994
      TabIndex        =   4
      Top             =   360
      Value           =   128
      Width           =   255
   End
   Begin VB.VScrollBar vsbRed1 
      DragIcon        =   "frmGradientForm.frx":265E
      Height          =   2415
      Left            =   120
      Max             =   255
      MouseIcon       =   "frmGradientForm.frx":3328
      TabIndex        =   0
      Top             =   360
      Value           =   128
      Width           =   255
   End
   Begin VB.VScrollBar vsbGreen1 
      Height          =   2415
      Left            =   360
      Max             =   255
      MouseIcon       =   "frmGradientForm.frx":3FF2
      TabIndex        =   1
      Top             =   360
      Value           =   128
      Width           =   255
   End
   Begin VB.VScrollBar vsbBlue1 
      Height          =   2415
      Left            =   600
      Max             =   255
      MouseIcon       =   "frmGradientForm.frx":4CBC
      TabIndex        =   3
      Top             =   360
      Value           =   128
      Width           =   255
   End
   Begin ProgramLauncher.xFrame xFrameG 
      Height          =   2895
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   5106
      Button          =   -1  'True
      Caption         =   "Click here to preview"
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
   End
End
Attribute VB_Name = "frmGradientForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Red1 As Integer
Dim Green1 As Integer
Dim Blue1 As Integer
Dim Red2 As Integer
Dim Green2 As Integer
Dim Blue2 As Integer

Private Sub Form_Paint()
    On Error Resume Next
    Red1 = vsbRed1.Value
    Green1 = vsbGreen1.Value
    Blue1 = vsbBlue1.Value
    Red2 = vsbRed2.Value
    Green2 = vsbGreen2.Value
    Blue2 = vsbBlue2.Value
    
    PaintGradient frmGradientForm, Red1, Green1, Blue1, Red2, Green2, Blue2
    ' PaintGradient frmGradientForm, 255, 255, 255, 128, 128, 255
    
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Dim Ftop As Integer
    Dim Fleft As Integer
    'get main forms window position
    Ftop = val(GetSetting(App.EXEName, "WinPos", "PrLrTop", ""))
    Fleft = val(GetSetting(App.EXEName, "WinPos", "PrLrleft", ""))
    
    With Me
        .Left = (Fleft - Me.Width - 50)
        .Top = Ftop
    End With
    
    vsbRed1.Value = GetSetting(App.EXEName, _
        "Gradient\Red1", "Value", Red1)
    vsbGreen1.Value = GetSetting(App.EXEName, _
        "Gradient\Green1", "Value", Green1)
    vsbBlue1.Value = GetSetting(App.EXEName, _
        "Gradient\Blue1", "Value", Blue1)
    vsbRed2.Value = GetSetting(App.EXEName, _
        "Gradient\Red2", "Value", Red2)
    vsbGreen2.Value = GetSetting(App.EXEName, _
        "Gradient\Green2", "Value", Green2)
    vsbBlue2.Value = GetSetting(App.EXEName, _
        "Gradient\Blue2", "Value", Blue2)
    
End Sub

Private Sub vsbRed1_Change()
    On Error Resume Next
    Red1 = vsbRed1.Value
    PaintGradient frmGradientForm, Red1, Green1, Blue1, Red2, Green2, Blue2
    SaveSetting App.EXEName, "Gradient\Red1", "Value", vsbRed1.Value
End Sub
Private Sub vsbGreen1_Change()
    On Error Resume Next
    Green1 = vsbGreen1.Value
    PaintGradient frmGradientForm, Red1, Green1, Blue1, Red2, Green2, Blue2
    SaveSetting App.EXEName, "Gradient\Green1", "Value", vsbGreen1.Value
End Sub
Private Sub vsbBlue1_Change()
    On Error Resume Next
    Blue1 = vsbBlue1.Value
    PaintGradient frmGradientForm, Red1, Green1, Blue1, Red2, Green2, Blue2
    SaveSetting App.EXEName, "Gradient\Blue1", "Value", vsbBlue1.Value
End Sub
Private Sub vsbRed2_Change()
    On Error Resume Next
    Red2 = vsbRed2.Value
    PaintGradient frmGradientForm, Red1, Green1, Blue1, Red2, Green2, Blue2
    SaveSetting App.EXEName, "Gradient\Red2", "Value", vsbRed2.Value
End Sub
Private Sub vsbGreen2_Change()
    On Error Resume Next
    Green2 = vsbGreen2.Value
    PaintGradient frmGradientForm, Red1, Green1, Blue1, Red2, Green2, Blue2
    SaveSetting App.EXEName, "Gradient\Green2", "Value", vsbGreen2.Value
End Sub
Private Sub vsbBlue2_Change()
    On Error Resume Next
    Blue2 = vsbBlue2.Value
    PaintGradient frmGradientForm, Red1, Green1, Blue1, Red2, Green2, Blue2
    SaveSetting App.EXEName, "Gradient\Blue2", "Value", vsbBlue2.Value
End Sub

Private Sub cmdApply_Click()
    On Error Resume Next
    PaintGradient frmProgramLauncher, Red1, Green1, Blue1, Red2, Green2, Blue2
    frmProgramLauncher.Refresh
End Sub

Private Sub cmdReset_Click()
    On Error Resume Next
    vsbRed1.Value = 255
    vsbGreen1.Value = 255
    vsbBlue1.Value = 255
    vsbRed2.Value = 128
    vsbGreen2.Value = 128
    vsbBlue2.Value = 255
    
    SaveSetting App.EXEName, "Gradient\Red1", "Value", vsbRed1.Value
    SaveSetting App.EXEName, "Gradient\Green1", "Value", vsbGreen1.Value
    SaveSetting App.EXEName, "Gradient\Blue1", "Value", vsbBlue1.Value
    SaveSetting App.EXEName, "Gradient\Red2", "Value", vsbRed2.Value
    SaveSetting App.EXEName, "Gradient\Green2", "Value", vsbGreen2.Value
    SaveSetting App.EXEName, "Gradient\Blue2", "Value", vsbBlue2.Value
End Sub

Private Sub cmdExit_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub xFrameG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End Sub


