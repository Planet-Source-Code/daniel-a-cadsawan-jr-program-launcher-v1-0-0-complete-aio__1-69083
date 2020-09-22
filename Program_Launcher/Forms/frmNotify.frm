VERSION 5.00
Begin VB.Form frmNotify 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2535
   ClientLeft      =   1365
   ClientTop       =   2010
   ClientWidth     =   5295
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ProgramLauncher.xFrame xFrame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4471
      Caption         =   "PrLr Scheduler Notice !!!!"
      DisplayPicture  =   -1  'True
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
      GradientBottom  =   16744576
      HeaderGradientBottom=   12611136
      Picture         =   "frmNotify.frx":0000
      Begin VB.CheckBox cbxAlwaysPlay 
         BackColor       =   &H00FF8080&
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmNotify.frx":059A
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   2160
         Width           =   255
      End
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "OK"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmNotify.frx":1264
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox lblText 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1575
         Left            =   960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "frmNotify.frx":1F2E
         Top             =   480
         Width           =   4215
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   240
         Picture         =   "frmNotify.frx":1F3A
         Stretch         =   -1  'True
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " Always Allow MP3 To Finish Playing even when Note is discarded."
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   2190
         Width           =   4815
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmNotify.frx":2C04
         Top             =   360
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbxAlwaysPlay_Click()
SaveRegLong HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "AlwaysPlay", cbxAlwaysPlay.Value

End Sub

Private Sub cmdOK_Click()
'user didnt select to allow to finish playing mp3, so we stop it
If cbxAlwaysPlay.Value = 0 Or cbxAlwaysPlay.Enabled = False Then frmSched.MediaPlayer1.Stop
'i added this
frmSched.show
frmSched.dtpTime.Value = Now
frmSched.tmrCaption.Enabled = True
If frmLockScr.WindowState = vbMaximized Then MakeTopMost frmLockScr.hwnd
Unload frmNotify
End Sub

Private Sub Form_Load()
'Dim lspeed As Integer
'lspeed = frmProgramLauncher.hsLoad.Value
    
'check to see if media file is mp3 so we can give option to continue playing,
'since we dont want to continue looping a windows sound event, it will never turn stop.
If Right(frmSched.MediaPlayer1.FileName, 1) = "3" Then cbxAlwaysPlay.Enabled = True

'load always play value
cbxAlwaysPlay.Value = GetRegLong(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "AlwaysPlay")

MakeTopMost frmNotify.hwnd

End Sub

Private Sub xFrame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    End If

End Sub
