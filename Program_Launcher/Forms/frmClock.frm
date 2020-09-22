VERSION 5.00
Begin VB.Form frmClock 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   12120
   ClientTop       =   8190
   ClientWidth     =   5175
   Icon            =   "frmClock.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   15  'Size All
   Picture         =   "frmClock.frx":0742
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   6135
      Top             =   5865
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5370
      Top             =   5865
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   18
      Left            =   4425
      MouseIcon       =   "frmClock.frx":55558
      MousePointer    =   99  'Custom
      ToolTipText     =   "Help"
      Top             =   5730
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   19
      Left            =   4155
      MouseIcon       =   "frmClock.frx":56222
      MousePointer    =   99  'Custom
      ToolTipText     =   "Hide For 10 Seconds"
      Top             =   5910
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   20
      Left            =   4665
      MouseIcon       =   "frmClock.frx":56EEC
      MousePointer    =   99  'Custom
      ToolTipText     =   "Exit"
      Top             =   5910
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   225
      Index           =   16
      Left            =   2715
      MousePointer    =   15  'Size All
      ToolTipText     =   "Move"
      Top             =   2715
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   14
      Left            =   4155
      MouseIcon       =   "frmClock.frx":57BB6
      MousePointer    =   99  'Custom
      Top             =   6180
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   15
      Left            =   4680
      MouseIcon       =   "frmClock.frx":58880
      MousePointer    =   99  'Custom
      Top             =   6180
      Width           =   150
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   2745
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   165
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   2715
      Shape           =   3  'Circle
      Top             =   2715
      Width           =   225
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   151
      X2              =   233
      Y1              =   424
      Y2              =   424
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      BorderWidth     =   4
      Index           =   0
      X1              =   43
      X2              =   125
      Y1              =   424
      Y2              =   424
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      Index           =   1
      X1              =   149
      X2              =   216
      Y1              =   409
      Y2              =   409
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      BorderWidth     =   7
      Index           =   0
      X1              =   41
      X2              =   104
      Y1              =   409
      Y2              =   409
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Index           =   1
      X1              =   151
      X2              =   216
      Y1              =   392
      Y2              =   392
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   9
      Index           =   0
      X1              =   43
      X2              =   104
      Y1              =   392
      Y2              =   392
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   59
      Left            =   6465
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   58
      Left            =   6330
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   57
      Left            =   6165
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   56
      Left            =   6015
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   55
      Left            =   5865
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   54
      Left            =   5715
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   53
      Left            =   5550
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   52
      Left            =   5415
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   51
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   50
      Left            =   5145
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   49
      Left            =   5010
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   48
      Left            =   4860
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   47
      Left            =   4710
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   46
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   45
      Left            =   4410
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   44
      Left            =   4260
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   43
      Left            =   4110
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   42
      Left            =   3945
      Shape           =   3  'Circle
      Top             =   5205
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   41
      Left            =   3780
      Shape           =   3  'Circle
      Top             =   5205
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   40
      Left            =   3630
      Shape           =   3  'Circle
      Top             =   5205
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   39
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   38
      Left            =   3315
      Shape           =   3  'Circle
      Top             =   5205
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   37
      Left            =   3135
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   36
      Left            =   2925
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   35
      Left            =   2745
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   34
      Left            =   2580
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   33
      Left            =   2415
      Shape           =   3  'Circle
      Top             =   5175
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   32
      Left            =   2220
      Shape           =   3  'Circle
      Top             =   5175
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   31
      Left            =   2025
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   30
      Left            =   1845
      Shape           =   3  'Circle
      Top             =   5205
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   29
      Left            =   1665
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   28
      Left            =   1485
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   27
      Left            =   1335
      Shape           =   3  'Circle
      Top             =   5250
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   26
      Left            =   1170
      Shape           =   3  'Circle
      Top             =   5250
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   25
      Left            =   975
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   24
      Left            =   810
      Shape           =   3  'Circle
      Top             =   5265
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   23
      Left            =   660
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   22
      Left            =   6450
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   21
      Left            =   6255
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   20
      Left            =   6060
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   19
      Left            =   5895
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   18
      Left            =   5730
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   17
      Left            =   5340
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   16
      Left            =   5190
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   15
      Left            =   5055
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   14
      Left            =   4875
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   13
      Left            =   4665
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   12
      Left            =   4425
      Shape           =   3  'Circle
      Top             =   5460
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   11
      Left            =   4275
      Shape           =   3  'Circle
      Top             =   5460
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   10
      Left            =   4095
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   9
      Left            =   3930
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   8
      Left            =   3750
      Shape           =   3  'Circle
      Top             =   5460
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   7
      Left            =   3540
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   6
      Left            =   3255
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   5
      Left            =   3015
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   4
      Left            =   2835
      Shape           =   3  'Circle
      Top             =   5505
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   3
      Left            =   2610
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   2
      Left            =   2325
      Shape           =   3  'Circle
      Top             =   5505
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   1
      Left            =   2100
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   0
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   150
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const PI = 3.141592654


Dim Lastx As Integer
Dim Lasty As Integer
Dim Saveleft As Variant
Dim Savetop As Variant
Dim DelayCounter As Long
Dim TopMost As Variant
Dim Startup As Variant
Public ShowHelp As Variant

Private Sub Form_Load()
    
    Dim csize As Integer
    csize = 75
    Dim Ret As Long
    Dim clr As Long
    Me.Hide
    For i = 0 To 59
        If i Mod 5 = 0 Then
            Shape4(i).Left = 188 + Cos(i * 2 * PI / 60 - (0.5 * PI)) * csize - 5
            Shape4(i).Top = 188 + sIn(i * 2 * PI / 60 - (0.5 * PI)) * csize - 5
        Else
            Shape4(i).Left = 188 + Cos(i * 2 * PI / 60 - (0.5 * PI)) * csize - 2.5
            Shape4(i).Top = 188 + sIn(i * 2 * PI / 60 - (0.5 * PI)) * csize - 2.5
        End If
        Shape4(i).BorderColor = vbBlack
    Next i
    Image2(18).Left = 188 + Cos(0 * 2 * PI / 60 - (0.5 * PI)) * csize - 5
    Image2(18).Top = 188 + sIn(0 * 2 * PI / 60 - (0.5 * PI)) * csize - 5
    Image2(19).Left = 188 + Cos(50 * 2 * PI / 60 - (0.5 * PI)) * csize - 5
    Image2(19).Top = 188 + sIn(50 * 2 * PI / 60 - (0.5 * PI)) * csize - 5
    Image2(20).Left = 188 + Cos(10 * 2 * PI / 60 - (0.5 * PI)) * csize - 5
    Image2(20).Top = 188 + sIn(10 * 2 * PI / 60 - (0.5 * PI)) * csize - 5
    Image2(14).Left = 188 + Cos(40 * 2 * PI / 60 - (0.5 * PI)) * csize - 5
    Image2(14).Top = 188 + sIn(40 * 2 * PI / 60 - (0.5 * PI)) * csize - 5
    Image2(15).Left = 188 + Cos(20 * 2 * PI / 60 - (0.5 * PI)) * csize - 5
    Image2(15).Top = 188 + sIn(40 * 2 * PI / 60 - (0.5 * PI)) * csize - 5
    
    clr = RGB(0, 0, 255) 'this color is the color that will be transparent
    'Set the window style to 'Layered'
    Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
    'Set the opacity of the layered window to 128
    SetLayeredWindowAttributes Me.hwnd, clr, 0, LWA_COLORKEY
    OpenFile
    SetTopMost
    'SetStartUp
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Lastx = X
Lasty = Y
   If Button = 1 Then
        ReleaseCapture
        frmClock.Left = frmClock.Left + (X - Lastx)
        frmClock.Top = frmClock.Top + (Y - Lasty)
        SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    End If
SaveFile
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmClock
End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 14
            '     Startup = Not Startup
            '     SetStartUp
        Case 15
            TopMost = Not TopMost
            SetTopMost
            SaveFile
        Case 16
            Lastx = X
            Lasty = Y
        Case 18
            frmClockHelp.show
        Case 19
            frmClock.Hide
            DelayCounter = 0
            Timer2.Interval = 100
        Case 20
            Unload frmClock
    End Select
End Sub
Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 16
            If Button = 1 Then
                frmClock.Left = frmClock.Left + (X - Lastx)
                frmClock.Top = frmClock.Top + (Y - Lasty)
            End If
    End Select
End Sub
Private Sub Image2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SaveFile
End Sub
Private Sub SetStartUp()
    If (Startup) Then
        Image2(14).ToolTipText = "Run Manually"
        SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "Transparent Analog Clock", App.Path & "\" & App.EXEName & ".exe"
        SaveFile
    Else
        Image2(14).ToolTipText = "Run When Windows Starts Up"
        SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "Transparent Analog Clock", "<NonRun>"
        SaveFile
    End If
End Sub
Private Sub SetTopMost()
    If TopMost Then
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        Image2(15).ToolTipText = "Make Not Always On Top"
    Else
        SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        Image2(15).ToolTipText = "Make Always On Top"
    End If
End Sub
Private Sub OpenFile()
    On Error GoTo There_Is_No_File
    Open App.Path & "\Data\" & "Analog Clock.dat" For Input Access Read As #1
        Line Input #1, Saveleft
        Line Input #1, Savetop
        Line Input #1, TopMost
        Line Input #1, Startup
        Line Input #1, ShowHelp
    Close #1
    GoTo The_File_Exist
There_Is_No_File:
    Dim Ftop As Integer
    Dim Fleft As Integer
    'get main forms window position
    Ftop = val(GetSetting(App.EXEName, "WinPos", "PrLrTop", ""))
    Fleft = val(GetSetting(App.EXEName, "WinPos", "PrLrleft", ""))
    Saveleft = (Fleft - Me.Width)
    Savetop = Ftop - 1000
    
    TopMost = True
    Startup = False
    ShowHelp = 1
    SaveFile
The_File_Exist:
    frmClock.Left = Saveleft
    frmClock.Top = Savetop
    frmClockHelp.Check1.Value = ShowHelp
    If ShowHelp Then frmClockHelp.show
End Sub
Public Sub SaveFile()
    Open App.Path & "\Data\" & "Analog Clock.dat" For Output As #1
        Print #1, frmClock.Left
        Print #1, frmClock.Top
        Print #1, TopMost
        Print #1, Startup
        Print #1, ShowHelp
    Close #1
End Sub


Private Sub Timer1_Timer()
    Dim Tim As Long
    Tim = Int(Timer)
    For i = 0 To 1
        'Hour
        Line1(i).X1 = 188
        Line1(i).Y1 = 188
        Line1(i).X2 = 188 + Cos((Tim Mod 43200) * 2 * PI / 43200 - (0.5 * PI)) * 45 '60
        Line1(i).Y2 = 188 + sIn((Tim Mod 43200) * 2 * PI / 43200 - (0.5 * PI)) * 45 '60
        'Minute
        Line2(i).X1 = 188
        Line2(i).Y1 = 188
        Line2(i).X2 = 188 + Cos((Tim \ 60 Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 60 '90
        Line2(i).Y2 = 188 + sIn((Tim \ 60 Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 60 '90
        'Second
        Line3(i).X1 = 188 - Cos((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 15
        Line3(i).Y1 = 188 - sIn((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 15
        Line3(i).X2 = 188 + Cos((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 65 '90
        Line3(i).Y2 = 188 + sIn((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 65 '90
    Next i
End Sub
Private Sub Timer2_Timer()
    DelayCounter = DelayCounter + 1
    If DelayCounter >= 100 Then
        frmClock.show
        Timer2.Interval = 0
    End If
End Sub
