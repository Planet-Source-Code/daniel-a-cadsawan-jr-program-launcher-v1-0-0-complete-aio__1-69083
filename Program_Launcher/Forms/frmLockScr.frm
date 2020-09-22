VERSION 5.00
Begin VB.Form frmLockScr 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimLock 
      Enabled         =   0   'False
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3855
      ScaleWidth      =   3915
      TabIndex        =   7
      Top             =   120
      Width           =   3915
      Begin VB.CommandButton cmdInput 
         Caption         =   "Enable"
         Height          =   360
         Left            =   780
         TabIndex        =   5
         Top             =   1530
         Width           =   795
      End
      Begin VB.CommandButton cmdEDUnlock 
         Caption         =   "Unlock w/ Encrypted PW"
         Enabled         =   0   'False
         Height          =   360
         Left            =   1680
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   1
         Top             =   2430
         Width           =   2115
      End
      Begin VB.CommandButton cmdHCUnlock 
         Caption         =   "Unlock w/ Hardcoded PW"
         Height          =   360
         Left            =   1680
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   3
         Top             =   3330
         Width           =   2115
      End
      Begin VB.PictureBox picIcon 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   180
         Picture         =   "frmLockScr.frx":0000
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   12
         Top             =   450
         Width           =   555
      End
      Begin VB.TextBox txtDomain 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1170
         Width           =   2115
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         Locked          =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   810
         Width           =   2115
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   450
         Width           =   2115
      End
      Begin VB.CommandButton cmdDOMUnlock 
         Caption         =   "Unlock PC Screen"
         Enabled         =   0   'False
         Height          =   360
         Left            =   1680
         TabIndex        =   6
         Top             =   1530
         Width           =   2115
      End
      Begin VB.TextBox txtHCPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2970
         Width           =   2115
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   180
         Picture         =   "frmLockScr.frx":0C42
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   9
         Top             =   2970
         Width           =   555
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   180
         Picture         =   "frmLockScr.frx":1884
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   8
         Top             =   2070
         Width           =   555
      End
      Begin VB.TextBox txtEDPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   2070
         Width           =   2115
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOCKED!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   780
         TabIndex        =   22
         Top             =   1620
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOCKED!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   780
         TabIndex        =   20
         Top             =   3420
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOCKED!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   780
         TabIndex        =   19
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unlock Screen Protector"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   930
         TabIndex        =   18
         Top             =   120
         Width           =   2130
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Domain:"
         Height          =   210
         Index           =   4
         Left            =   780
         TabIndex        =   17
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   210
         Index           =   3
         Left            =   780
         TabIndex        =   16
         Top             =   840
         Width           =   780
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         Height          =   210
         Index           =   2
         Left            =   780
         TabIndex        =   15
         Top             =   480
         Width           =   780
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   210
         Index           =   0
         Left            =   780
         TabIndex        =   14
         Top             =   3000
         Width           =   780
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   210
         Index           =   1
         Left            =   780
         TabIndex        =   13
         Top             =   2100
         Width           =   780
      End
   End
   Begin VB.Label lblBanner 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Program Launcher is Protecting this PC!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   60
      TabIndex        =   21
      Top             =   60
      Width           =   8595
   End
End
Attribute VB_Name = "frmLockScr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iTries As Integer

Private Sub Form_Load()
Dim Str As String

Me.WindowState = 2
'Me.Width = Screen.Width
'Me.Height = Screen.Height
picFrame.Left = Screen.Width / 2 - picFrame.Width / 2
picFrame.Top = Screen.Height / 2 - picFrame.Height / 2
lblBanner.Left = Screen.Width / 2 - lblBanner.Width / 2
lblBanner.Top = Screen.Height / 2 - picFrame.Height ' / 2 - lblBanner.Height

If Len(Dir(inifilepw)) <> 0 Then cmdEDUnlock.Enabled = True
txtUserName.Text = CurrentLogonUser
txtDomain.Text = CurrentDomain

' do all locking on windows
'===============================
AntiTaskManagerController False
tskbar False
HookKeyboard
'===============================
MakeTopMost frmLockScr.hwnd

End Sub

Private Sub Unlock_Screen()
AntiTaskManagerController True
tskbar True
UnHookKeyboard
iTries = 0
Unload Me
End Sub

Private Sub cmdEDUnlock_Click()

Dim strPW As String
strPW = ReadSetting("Default", True)

If strPW = vbNullString Then
MsgBox "There is no password currently set" & vbCrLf _
     & "or the password file was altered!" & vbCrLf _
     & "The Encrypted button will now be disable.", vbInformation
txtEDPassword.BackColor = &HC0C0C0
txtEDPassword.Enabled = False
cmdEDUnlock.Enabled = False
Label1.Visible = True
Exit Sub
End If

If txtEDPassword = strPW Then
Call Unlock_Screen
Else
    iTries = iTries + 1
    If iTries >= 3 Then
    MsgBox "You have attempted 3 times to unlock the" & vbCrLf _
         & "Encrypted password but did not succeed!" & vbCrLf _
         & "The Encrypted button will now be disable" & vbCrLf _
         & "as well as all other controls.", vbInformation
    Else
    MsgBox "The Encrypted Password you supply is incorrect!", vbExclamation
    txtEDPassword.SetFocus
    txtEDPassword.SelStart = 0
    txtEDPassword.SelLength = Len(txtEDPassword.Text)
    End If
End If

If iTries = 3 Then
Call Lock_Controls
End If

End Sub

Private Sub Lock_Controls()
txtEDPassword.BackColor = &HC0C0C0
txtEDPassword.Enabled = False
cmdEDUnlock.Enabled = False
txtHCPassword.BackColor = &HC0C0C0
txtHCPassword.Enabled = False
cmdHCUnlock.Enabled = False
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True

cmdInput.Visible = False
txtPassword.Enabled = False
txtPassword.BackColor = &HC0C0C0
cmdDOMUnlock.Enabled = False

TimLock.Interval = 60000 '60000 = 1 minute
TimLock.Enabled = True
End Sub

Private Sub TimLock_Timer()
iTries = 0
txtEDPassword.BackColor = &HFFFFFF
txtEDPassword.Enabled = True
cmdEDUnlock.Enabled = True
txtHCPassword.BackColor = &HFFFFFF
txtHCPassword.Enabled = True
cmdHCUnlock.Enabled = True
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False

cmdInput.Visible = True
cmdInput.Enabled = True
txtPassword.Enabled = False
txtPassword.BackColor = &HC0C0C0
txtPassword.Locked = True
cmdDOMUnlock.Enabled = False

TimLock.Enabled = False
End Sub

Private Sub cmdHCUnlock_Click()

If txtHCPassword = "DaniBoyIncorporated" Then
Call Unlock_Screen
Else
    iTries = iTries + 1
    If iTries >= 3 Then
    MsgBox "You have attempted 3 times to unlock the" & vbCrLf _
         & "HardCoded password but did not succeed!" & vbCrLf _
         & "The HardCoded button will now be disable" & vbCrLf _
         & "as well as all other controls.", vbInformation
    Else
    MsgBox "The HardCoded Password you supply is incorrect!", vbInformation
    txtHCPassword.SetFocus
    txtHCPassword.SelStart = 0
    txtHCPassword.SelLength = Len(txtEDPassword.Text)
    End If
End If

If iTries = 3 Then
Call Lock_Controls
End If

End Sub

Private Sub cmdInput_Click()
txtPassword.Enabled = True
txtPassword.BackColor = &HFFFFFF
txtPassword.Locked = False
cmdDOMUnlock.Enabled = True
cmdInput.Enabled = False
End Sub

Private Sub cmdDOMUnlock_Click()

    Dim LogonDomain As Boolean
    LogonDomain = VerifyLogin(txtUserName.Text, txtDomain.Text, txtPassword.Text)

    If txtUserName.Text = "" Then
        MsgBox "User name is empty!", vbCritical, "Login"
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    
    If txtDomain.Text = "" Then
        MsgBox "Domain name is empty!", vbCritical, "Login"
        txtDomain.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
   
    If txtPassword.Text = "" Then
        MsgBox "Password is empty!", vbCritical, "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
   
    If Not LogonDomain Then
        iTries = iTries + 1
        MsgBox "Incorrect login for username " & txtDomain & "\" & txtUserName, vbCritical, "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
            If iTries = 2 Then
            MsgBox "Only 2 attempts allowed" & vbCrLf _
                 & "for the Domain Unlocking," & vbCrLf _
                 & "All controls will now be disable.", vbInformation
            Call Lock_Controls
            End If
        Exit Sub
    End If
    
    Call Unlock_Screen
    
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Call cmdDOMUnlock_Click
    End If
End Sub

Private Sub txtEDPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Call cmdEDUnlock_Click
    End If
End Sub

Private Sub txtHCPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Call cmdHCUnlock_Click
    End If
End Sub

Private Sub txtUserName_GotFocus()
picIcon.SetFocus
End Sub

Private Sub txtDomain_GotFocus()
picIcon.SetFocus
End Sub
