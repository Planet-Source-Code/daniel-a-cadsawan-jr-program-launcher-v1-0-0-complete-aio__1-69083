VERSION 5.00
Begin VB.Form frmChangePW 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   4215
   ControlBox      =   0   'False
   Icon            =   "frmChangePW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   3975
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtNoPW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   """Blank at the moment"""
         Top             =   480
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2940
         TabIndex        =   6
         Top             =   1560
         Width           =   915
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   1980
         TabIndex        =   5
         Top             =   1560
         Width           =   915
      End
      Begin VB.TextBox txtExistingPassword 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1980
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   480
         Width           =   1875
      End
      Begin VB.TextBox txtNewPassword1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1980
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   840
         Width           =   1875
      End
      Begin VB.TextBox txtNewPassword2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1980
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   1875
      End
      Begin VB.TextBox Text1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1980
         TabIndex        =   1
         Top             =   60
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Screen Protector Password"
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
         Left            =   480
         TabIndex        =   11
         Top             =   120
         Width           =   3030
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter Existing Password"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter New Password"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   900
         Width           =   1635
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Confirm New Password"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1260
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Do not use space / linefeed as password"
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmChangePW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim strPW As String
strPW = ReadSetting("Default", True)
Text1.Text = strPW

If strPW = vbNullString Then
txtNoPW.Visible = True
End If

End Sub

Private Sub cmdOK_Click()

Dim strTemp As String
Dim strPW As String
Dim strNewPW As String
strPW = ReadSetting("Default", True)
'strNewPW = LCase(txtNewPassword2.Text)
strNewPW = txtNewPassword2.Text
    
    'checks to see if you type in the correct password in the existing password field
    'If LCase(strPW) = LCase(txtExistingPassword.Text) Then
    If strPW = txtExistingPassword.Text Then

        'checks the match of the new passwords
        'If LCase(txtNewPassword1.Text) = strNewPW Then
        If txtNewPassword1.Text = strNewPW Then
            Dim PWpath As String
            PWpath = inifilepw
            If Len(Dir(PWpath)) <> 0 Then Kill PWpath

            WriteSetting "Default", strNewPW, True
            MsgBox "Password changed!", 8, "Password Verfication"
        
        Else
            MsgBox "The New Passwords Do Not Match", 8, "Password Error"
            txtNewPassword1.SetFocus
            Exit Sub
        
        End If
        
    Else
        MsgBox "The Existing Password is Incorrect or Altered!", 8, "Password Error"
        txtExistingPassword.SetFocus
        Exit Sub
        
    End If
    
    Unload Me
    DoEvents
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtNoPW_GotFocus()
txtNewPassword1.SetFocus
End Sub
