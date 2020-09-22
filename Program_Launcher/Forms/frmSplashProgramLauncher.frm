VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplashProgramLauncher 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplashProgramLauncher.frx":0000
   ScaleHeight     =   4095
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerSplash 
      Enabled         =   0   'False
      Left            =   2040
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   120
      Picture         =   "frmSplashProgramLauncher.frx":8001
      ScaleHeight     =   1650
      ScaleWidth      =   1410
      TabIndex        =   1
      Top             =   480
      Width           =   1440
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      Picture         =   "frmSplashProgramLauncher.frx":F8B4
      ScaleHeight     =   945
      ScaleWidth      =   2265
      TabIndex        =   0
      Top             =   3000
      Width           =   2295
   End
   Begin MSComctlLib.ProgressBar PBarSplash 
      Height          =   3495
      Left            =   2520
      TabIndex        =   2
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   6165
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
End
Attribute VB_Name = "frmSplashProgramLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Form_Activate()
    On Error Resume Next
    Dim iX As Long
    Static bActive As Boolean
    If bActive Then Exit Sub
    bActive = True
    For iX = 1 To 255 Step 1 '2
        MakeWindowTransparent frmSplashProgramLauncher.hwnd, iX
        DoEvents    ' need this so form doesn't turn black
    Next
    
End Sub


Private Sub Form_Load()
    On Error Resume Next
    
    
    If App.PrevInstance = True Then
        Unload frmSplashProgramLauncher
        Exit Sub
    End If
    
    'checks if the option value is set to yes = no splash
    frmProgramLauncher.chkSplash.Value = GetSetting(App.EXEName, _
        "ShowSplash", "NoShow", "")
    
    If frmProgramLauncher.chkSplash.Value = 0 Then
        
        Label1.Caption = App.Title & _
            " Version " & App.Major & "." & _
            App.Minor & "." & App.Revision
        TimerSplash.Enabled = True
        TimerSplash.Interval = 15
        ' start off transparent so form doesn't flicker
        MakeWindowTransparent frmSplashProgramLauncher.hwnd, 1 '10
    Else
        
        'frmSplashProgramLauncher.Hide
        Unload frmSplashProgramLauncher
        Set frmSplashProgramLauncher = Nothing
        frmProgramLauncher.show
    End If
    
End Sub


Private Sub TimerSplash_Timer()
    On Error Resume Next
    If PBarSplash.Value < 100 Then
        PBarSplash.Value = PBarSplash.Value + 1
    Else
        TimerSplash.Enabled = False
        frmProgramLauncher.show
        Unload frmSplashProgramLauncher
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'CODE : form unload
    TimerSplash.Enabled = False
    'CODE : fade form on unload
    Dim iX As Long
    For iX = 255 To 0 Step -1 '-2
        MakeWindowTransparent frmSplashProgramLauncher.hwnd, iX
    Next
    'Unload our form completely
    Unload frmSplashProgramLauncher
    frmProgramLauncher.show
    
End Sub

