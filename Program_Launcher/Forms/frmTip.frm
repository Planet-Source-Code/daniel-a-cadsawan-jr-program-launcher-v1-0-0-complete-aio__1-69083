VERSION 5.00
Begin VB.Form frmTip 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LBL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3585
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   1755
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LBL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitle 
      Height          =   225
      Left            =   0
      Picture         =   "frmTip.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1845
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Init(TheCaption$, TheText$, pxWidth As Long, pxHeight As Long)
    lblCaption.Caption = TheCaption
    lblText.Caption = TheText
    Me.ScaleWidth = pxWidth
    Me.ScaleHeight = pxHeight
    imgTitle.Width = pxWidth
    lblCaption.Width = pxWidth
    lblText.Width = pxWidth - 5
End Sub

