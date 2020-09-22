VERSION 5.00
Begin VB.Form frmClockHelp 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3900
   ClientLeft      =   6615
   ClientTop       =   5670
   ClientWidth     =   4815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmClockHelp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmClockHelp.frx":0742
   ScaleHeight     =   3900
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   180
      Left            =   1440
      TabIndex        =   1
      Top             =   3240
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   330
      Left            =   1920
      Picture         =   "frmClockHelp.frx":5285
      TabIndex        =   0
      Top             =   3480
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Show Help On Start Up"
      Height          =   240
      Left            =   1680
      TabIndex        =   2
      Top             =   3240
      Width           =   1800
   End
End
Attribute VB_Name = "frmClockHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    frmClock.ShowHelp = Check1.Value
    frmClock.SaveFile
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub
