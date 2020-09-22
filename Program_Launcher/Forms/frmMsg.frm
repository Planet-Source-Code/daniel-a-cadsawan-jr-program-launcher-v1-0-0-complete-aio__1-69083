VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMsg 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSkinOff 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1020
      MouseIcon       =   "frmMsg.frx":000C
      MousePointer    =   99  'Custom
      Picture         =   "frmMsg.frx":0CD6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      ToolTipText     =   "Activate skin"
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picSkinOn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1020
      MouseIcon       =   "frmMsg.frx":2F65
      MousePointer    =   99  'Custom
      Picture         =   "frmMsg.frx":3C2F
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      ToolTipText     =   "Deactivate skin"
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picSave 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   765
      MouseIcon       =   "frmMsg.frx":5E20
      MousePointer    =   99  'Custom
      Picture         =   "frmMsg.frx":6AEA
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   11
      ToolTipText     =   "Save note"
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picTransOff 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   510
      MouseIcon       =   "frmMsg.frx":8EA0
      MousePointer    =   99  'Custom
      Picture         =   "frmMsg.frx":9B6A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   10
      ToolTipText     =   "Transparency off"
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picTransOn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   510
      MouseIcon       =   "frmMsg.frx":BD8B
      MousePointer    =   99  'Custom
      Picture         =   "frmMsg.frx":CA55
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   9
      ToolTipText     =   "Transparency on"
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picMin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   255
      MouseIcon       =   "frmMsg.frx":ECCD
      MousePointer    =   99  'Custom
      Picture         =   "frmMsg.frx":F997
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
      ToolTipText     =   "Minimize this note"
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picMax 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   255
      MouseIcon       =   "frmMsg.frx":11BB8
      MousePointer    =   99  'Custom
      Picture         =   "frmMsg.frx":12882
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   8
      ToolTipText     =   "Maximize this note"
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picDump 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      MouseIcon       =   "frmMsg.frx":14AE2
      MousePointer    =   99  'Custom
      Picture         =   "frmMsg.frx":157AC
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      ToolTipText     =   "Discard this note"
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   255
      MouseIcon       =   "frmMsg.frx":17A4C
      MousePointer    =   99  'Custom
      Picture         =   "frmMsg.frx":18716
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   5
      ToolTipText     =   "Change background color"
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox picFont 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      MouseIcon       =   "frmMsg.frx":1AA01
      MousePointer    =   99  'Custom
      Picture         =   "frmMsg.frx":1B6CB
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      ToolTipText     =   "Change font and color"
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3600
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmMsg.frx":1DA65
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      ToolTipText     =   "Resize note"
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox picMove2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   510
      MousePointer    =   5  'Size
      Picture         =   "frmMsg.frx":1FBD8
      ScaleHeight     =   255
      ScaleWidth      =   3090
      TabIndex        =   4
      ToolTipText     =   "LeftClick hold to move"
      Top             =   1200
      Width           =   3090
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picMove 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1275
      MousePointer    =   5  'Size
      Picture         =   "frmMsg.frx":2430A
      ScaleHeight     =   225
      ScaleWidth      =   2550
      TabIndex        =   1
      ToolTipText     =   "LeftClick hold to move"
      Top             =   0
      Width           =   2580
   End
   Begin RichTextLib.RichTextBox RTBNote 
      Height          =   945
      Left            =   15
      TabIndex        =   0
      Top             =   240
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1667
      _Version        =   393217
      BackColor       =   16777152
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMsg.frx":285F4
      MouseIcon       =   "frmMsg.frx":28674
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   2400
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   200
      ImageHeight     =   200
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsg.frx":28690
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsg.frx":2BD9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsg.frx":2F46C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgSkin 
      Appearance      =   0  'Flat
      Height          =   945
      Left            =   20
      Picture         =   "frmMsg.frx":332D1
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private stX As Double, stY As Double

' scrollbar checker
Private Declare Function ShowScrollBar Lib "user32" _
    (ByVal hwnd As Long, ByVal wBar As Long, _
    ByVal bShow As Long) As Long
Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Const SB_HWND = 2
Private Const SB_BOTH = 3
Private Const EM_GETLINECOUNT = &HBA

Private Sub CheckScrollBars()
    Dim lngLineCount As Long
    Static VVisible  As Boolean
    Static HVisible  As Boolean
    Dim lWidth       As Long
    Dim lHeight      As Long
    Dim i            As Integer
    
    lngLineCount = SendMessageAsLong(RTBNote.hwnd, EM_GETLINECOUNT, 0, 0)
    
    For i = 1 To 2 '2 times, because if HScrollbar appeared, VScrollbar code will be different
        
        'If Horizontal ScrollBar is present, Height = RTBNote.Height - [Height of the Hscrollbar]
        If HVisible Then lHeight = RTBNote.Height - 250 Else lHeight = RTBNote.Height - 150
        If Me.TextHeight("A") * lngLineCount > lHeight Then
            ShowScrollBar RTBNote.hwnd, SB_VERT, True  'Higher than Height?
            VVisible = True                          'Make VScrollbar visible
        Else
            ShowScrollBar RTBNote.hwnd, SB_VERT, False 'Hide VScrollbar
            VVisible = False
        End If
        
        'If Vertical ScrollBar is present width = RTBNote.width - [width of the Vscrollbar]
        'If VVisible Then lWidth = RTBNote.Width - 350 Else lWidth = RTBNote.Width - 100
        'If Me.TextWidth(RTBNote.Text) > lWidth Then    'Wider than width?
        'ShowScrollBar RTBNote.hwnd, SB_HORZ, True  'Make HScrollbar visible
        'HVisible = True
        'Else
        'ShowScrollBar RTBNote.hwnd, SB_HORZ, False 'Hide HScrollbar
        'HVisible = False
        'End If
    Next
    
End Sub

Private Sub Form_Activate()
Dim lspeed As Integer
lspeed = frmProgramLauncher.hsLoad.Value
    Dim iX As Long
    'makes the form static
    Static bActive As Boolean
    If bActive Then Exit Sub
    bActive = True
    
    'check for Skin Off or On
    If picSkinOff.Visible = False Then
        RTBNote.Visible = False
        SetWindowLong RTBNote.hwnd, GWL_EXSTYLE, _
        GetWindowLong(RTBNote.hwnd, GWL_EXSTYLE) Or WS_EX_TRANSPARENT
        RTBNote.Visible = True
    End If
    
    'check for transparency window or not then set translucency
    If picTransOff.Visible = True Then
        For iX = lspeed To 150 Step 1
            MakeWindowTransparent Me.hwnd, iX
            DoEvents    ' need this so form doesn't turn black
        Next
    Else
        For iX = lspeed To 255 Step 1
            MakeWindowTransparent Me.hwnd, iX
            DoEvents    ' need this so form doesn't turn black
        Next
    End If
    
End Sub

Private Sub Form_Load()
Dim lspeed As Integer
lspeed = frmProgramLauncher.hsLoad.Value

    Me.Height = 1500
    Me.Width = 3000
    
    MakeWindowTransparent Me.hwnd, lspeed ' need this so form doesn't flicker
    
End Sub

Private Sub Form_Resize()
    
    LockWindow Me.hwnd, True  'To turn it on
    
    If Me.Width < 2000 Then
        Me.Width = 2000
    End If
    If Me.Height < 510 Then
        Me.Height = 510
    End If
    
    RTBNote.Height = Me.Height - 510
    RTBNote.Width = Me.Width - 30
    imgSkin.Height = Me.Height - 510
    imgSkin.Width = Me.Width - 30
    picFont.Top = Me.Height - picFont.Height
    picColor.Top = Me.Height - picColor.Height
    picMove.Width = Me.Width - 255 * 5
    picMove2.Top = Me.Height - picMove2.Height
    picMove2.Width = Me.Width - 255 * 3
    picDrag.Left = Me.Width - picDrag.Width
    picDrag.Top = Me.Height - picDrag.Height
    
    LockWindow Me.hwnd, False 'To Turn it off
    
    Me.Refresh
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim lspeed As Integer
lspeed = frmProgramLauncher.hsLoad.Value
    
    If picTransOff.Visible = True Then
        Dim iX As Long
        For iX = 150 To 0 Step -lspeed '-2
            MakeWindowTransparent Me.hwnd, iX
        Next
    Else
        For iX = 255 To 0 Step -lspeed '-2
            MakeWindowTransparent Me.hwnd, iX
        Next
    End If
    
    Unload Me
End Sub

Private Sub picColor_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        CommonDialog1.flags = cdlCCRGBInit 'Or cdlCCFullOpen
        ' Set initial values for the dialog box.
        CommonDialog1.Color = RTBNote.BackColor
        On Error GoTo ErrorHandler
        
        With CommonDialog1
            .CancelError = True
            .ShowColor
            RTBNote.BackColor = .Color
        End With
        
        'Call SaveOpenForms
        
    End If
Exit Sub
ErrorHandler:
    If Err.Number <> cdlCancel Then
    MsgBox Err.Number & " - " & Err.Description, _
    vbOKOnly + vbExclamation
    End If
End Sub

Private Sub picColor_MouseMove(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    
    If (X Or Y) < 0 Or (X > picColor.Width) Or _
            (Y > picColor.Height) Then
        ReleaseCapture
        If picTransOff.Visible = True Then MakeWindowTransparent Me.hwnd, 150
    ElseIf GetCapture() <> picColor.hwnd Then
        SetCapture picColor.hwnd
        MakeWindowTransparent Me.hwnd, 255
        
    End If
    
End Sub

Private Sub RTBNote_Change()
    'CheckScrollBars
    'RTBNote.Refresh
End Sub

Private Sub RTBNote_MouseMove(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    
    If (X Or Y) < 0 Or (X > RTBNote.Width) Or _
            (Y > RTBNote.Height) Then
        Screen.MousePointer = vbDefault
        ReleaseCapture
        If picTransOff.Visible = True Then MakeWindowTransparent Me.hwnd, 150
    ElseIf GetCapture() <> RTBNote.hwnd Then
        Screen.MousePointer = vbArrow
        SetCapture RTBNote.hwnd
        MakeWindowTransparent Me.hwnd, 255
    End If
    
End Sub
Private Sub picMove_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        ReleaseCapture
        SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
        
        'Call SaveOpenForms
    End If
End Sub

Private Sub picMove_MouseMove(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    
    If (X Or Y) < 0 Or (X > picMove.Width) Or _
            (Y > picMove.Height) Then
        ReleaseCapture
        If picTransOff.Visible = True Then MakeWindowTransparent Me.hwnd, 150
    ElseIf GetCapture() <> picMove.hwnd Then
        SetCapture picMove.hwnd
        MakeWindowTransparent Me.hwnd, 255
        
    End If
    
End Sub

Private Sub picMin_Click()
    Me.Height = 255 * 2
    picFont.Visible = False
    picColor.Visible = False
    picMove2.Visible = False
    picDrag.Visible = False
    picMin.Visible = False
    
    'Call SaveOpenForms
    
End Sub

Private Sub picMin_MouseMove(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    
    If (X Or Y) < 0 Or (X > picMin.Width) Or _
            (Y > picMin.Height) Then
        ReleaseCapture
        If picTransOff.Visible = True Then MakeWindowTransparent Me.hwnd, 150
    ElseIf GetCapture() <> picMin.hwnd Then
        SetCapture picMin.hwnd
        MakeWindowTransparent Me.hwnd, 255
        
    End If
End Sub

Private Sub picMax_Click()
    Me.Height = 2000
    picFont.Visible = True
    picColor.Visible = True
    picMove2.Visible = True
    picDrag.Visible = True
    picMin.Visible = True
    
    'Call SaveOpenForms
    
End Sub

Private Sub picMax_MouseMove(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    
    If (X Or Y) < 0 Or (X > picMax.Width) Or _
            (Y > picMax.Height) Then
        ReleaseCapture
        If picTransOff.Visible = True Then MakeWindowTransparent Me.hwnd, 150
    ElseIf GetCapture() <> picMax.hwnd Then
        SetCapture picMax.hwnd
        MakeWindowTransparent Me.hwnd, 255
        
    End If
End Sub

Private Sub picDump_Click()
    'AYAW GUMANA PA
    'Call DeleteFile(App.Path & "\Data\" & Me.Tag & ".TextRTF")
    
    Unload Me
    Call SaveOpenForms
    Call GetFormsInformation
    
End Sub

Private Sub picDump_MouseMove(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    
    If (X Or Y) < 0 Or (X > picDump.Width) Or _
            (Y > picDump.Height) Then
        ReleaseCapture
        If picTransOff.Visible = True Then MakeWindowTransparent Me.hwnd, 150
    ElseIf GetCapture() <> picDump.hwnd Then
        SetCapture picDump.hwnd
        MakeWindowTransparent Me.hwnd, 255
        
    End If
End Sub

Private Sub picTransOff_Click()
    MakeWindowTransparent Me.hwnd, 255
    picTransOff.Visible = False
    
    'Call SaveOpenForms
End Sub

Private Sub picTransOff_MouseMove(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    
    If (X Or Y) < 0 Or (X > picTransOff.Width) Or _
            (Y > picTransOff.Height) Then
        ReleaseCapture
        If picTransOff.Visible = True Then MakeWindowTransparent Me.hwnd, 150
    ElseIf GetCapture() <> picTransOff.hwnd Then
        SetCapture picTransOff.hwnd
        MakeWindowTransparent Me.hwnd, 255
        
    End If
End Sub

Private Sub picTransOn_Click()
    MakeWindowTransparent Me.hwnd, 150
    picTransOff.Visible = True
    
    'Call SaveOpenForms
End Sub
Private Sub picTransOn_MouseMove(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    
    If (X Or Y) < 0 Or (X > picTransOn.Width) Or _
            (Y > picTransOn.Height) Then
        ReleaseCapture
        If picTransOff.Visible = True Then MakeWindowTransparent Me.hwnd, 150
    ElseIf GetCapture() <> picTransOn.hwnd Then
        SetCapture picTransOn.hwnd
        MakeWindowTransparent Me.hwnd, 255
        
    End If
End Sub

Private Sub picSkinOff_Click()
    SetWindowLong RTBNote.hwnd, GWL_EXSTYLE, _
    GetWindowLong(RTBNote.hwnd, GWL_EXSTYLE) Or WS_EX_TRANSPARENT
    picSkinOff.Visible = False
    imgSkin.Picture = ImgList.ListImages(3).Picture
    SaveSetting App.EXEName, "NoteSkin" & "\" & Me.Tag, "Value", "3"
'    Call SaveOpenForms
End Sub

Private Sub picSkinOff_MouseMove(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    
    If (X Or Y) < 0 Or (X > picSkinOff.Width) Or _
            (Y > picSkinOff.Height) Then
        ReleaseCapture
        If picTransOff.Visible = True Then MakeWindowTransparent Me.hwnd, 150
    ElseIf GetCapture() <> picSkinOff.hwnd Then
        SetCapture picSkinOff.hwnd
        MakeWindowTransparent Me.hwnd, 255
        
    End If
End Sub

Private Sub picSkinOn_Click()
    SetWindowLong RTBNote.hwnd, GWL_EXSTYLE, _
        GetWindowLong(RTBNote.hwnd, GWL_EXSTYLE) And Not WS_EX_TRANSPARENT
    picSkinOff.Visible = True
    'Call SaveOpenForms
End Sub

Private Sub picSkinOn_MouseMove(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    
    If (X Or Y) < 0 Or (X > picSkinOn.Width) Or _
            (Y > picSkinOn.Height) Then
        ReleaseCapture
        If picTransOff.Visible = True Then MakeWindowTransparent Me.hwnd, 150
    ElseIf GetCapture() <> picSkinOn.hwnd Then
        SetCapture picSkinOn.hwnd
        MakeWindowTransparent Me.hwnd, 255
        
    End If
End Sub

Private Sub picSave_Click()
    Call SaveOpenForms
    frmProgramLauncher.mnuShowAllNotes.Enabled = False
    frmProgramLauncher.mnuHideAllNotes.Enabled = True
    frmProgramLauncher.cmdHideAllNotes.Enabled = True
    Call GetFormsInformation
End Sub

Private Sub picSave_MouseMove(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    
    If (X Or Y) < 0 Or (X > picSave.Width) Or _
            (Y > picSave.Height) Then
        ReleaseCapture
        If picTransOff.Visible = True Then MakeWindowTransparent Me.hwnd, 150
    ElseIf GetCapture() <> picSave.hwnd Then
        SetCapture picSave.hwnd
        MakeWindowTransparent Me.hwnd, 255
    End If
    
End Sub
Private Sub picMove2_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    
    'If Button = 1 Then
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    'Call SaveOpenForms
    'ElseIf Button = 2 Then
    'PopupMenu mnuOptions
    'End If
    
End Sub

Private Sub picMove2_MouseMove(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    
    If (X Or Y) < 0 Or (X > picMove2.Width) Or _
            (Y > picMove2.Height) Then
        ReleaseCapture
        If picTransOff.Visible = True Then MakeWindowTransparent Me.hwnd, 150
    ElseIf GetCapture() <> picMove2.hwnd Then
        SetCapture picMove2.hwnd
        MakeWindowTransparent Me.hwnd, 255
        
    End If
    
End Sub
Private Sub picDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Right...first step is to grab the "mousedown" x,y value.
    'This takes the position of the mouse when the button is
    'clicked and held.
    stX = X: stY = Y '"st" = start...this is the start x/y co-ords
    
End Sub

Private Sub picDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If stX = 0 Or stY = 0 Then Exit Sub
    'these two lines resize the form based on x and stX & y and stY
    LockWindow Me.hwnd, True  'To turn it on
    MakeWindowTransparent Me.hwnd, 255
    Me.Width = (stX + X) + picDrag.Left
    Me.Height = (stY + Y) + picDrag.Top
    
End Sub

Private Sub picDrag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    stX = 0: stY = 0 'This line obviously clears the two
    LockWindow Me.hwnd, False  'To turn it off
    If picTransOff.Visible = True Then MakeWindowTransparent Me.hwnd, 150
    'Call SaveOpenForms
    
End Sub

Private Sub picFont_Click()
    On Error GoTo ErrorHandler
    
    ' Set Flags
    CommonDialog1.flags = cdlCFBoth Or cdlCFEffects
    ' Set initial values for the dialog box.
    With CommonDialog1
        .CancelError = True
        .ShowFont
    End With
    ' If canceled, exit.
    If CommonDialog1.FontName = "" Then Exit Sub
    
    With Me.RTBNote
        '   Change the text font according to options selected.
        .Font.Name = CommonDialog1.FontName
        .Font.Size = CommonDialog1.FontSize
        .Font.Bold = CommonDialog1.FontBold
        .Font.Italic = CommonDialog1.FontItalic
        .Font.Underline = CommonDialog1.FontUnderline
        .Font.Strikethrough = CommonDialog1.FontStrikethru
        .SelColor = CommonDialog1.Color
    End With
    
    'Call SaveOpenForms
Exit Sub
ErrorHandler:
    If Err.Number <> cdlCancel Then
    MsgBox Err.Number & " - " & Err.Description, _
    vbOKOnly + vbExclamation
    End If
End Sub

Private Sub picFont_MouseMove(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    
    If (X Or Y) < 0 Or (X > picFont.Width) Or _
            (Y > picFont.Height) Then
        ReleaseCapture
        If picTransOff.Visible = True Then MakeWindowTransparent Me.hwnd, 150
    ElseIf GetCapture() <> picFont.hwnd Then
        SetCapture picFont.hwnd
        MakeWindowTransparent Me.hwnd, 255
        
    End If
    
End Sub

Private Sub RTBNote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandler
    If Button = 2 Then
        
        '--------------------------------------------------------------------------------------------------------------------------
        'CODE AUTOGENERATED WITH:  MC API Menu Code Generator ver 2.0
        '---------------------------------------------------------------------------------------------------------------------------
        Dim hPopupMenu1 As Long ' handle to the popup menu to display
        Dim hPopupMenu2 As Long ' handle to the popup menu to display
        Dim mii1 As MENUITEMINFO   ' describes menu items to add
        Dim mii2 As MENUITEMINFO   ' describes menu items to add
        Dim curpos As POINT_TYPE  ' holds the current mouse coordinates
        Dim menusel As Long       ' ID of what the user selected in the popup menu
        Dim RetVal As Long        ' generic return value
        
        
        'Create the popup menus which are initialy empty.
        hPopupMenu1 = CreatePopupMenu()
        hPopupMenu2 = CreatePopupMenu()
        
        'Create the structure which is the base for all menus:
        With mii1
            .cbSize = Len(mii1) ' The size of this structure.
            .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU ' Which elements of the structure to use.
        End With
        
        'Make all structures equal
        mii2 = mii1
        
        With mii2 '(Change SelText Color & Font)
            .fType = MFT_STRING
            .fState = MFS_ENABLED
            .wID = 1001 ' Assign this item an item identifier.
            .dwTypeData = "Change SelText Color n Font"
            .cch = Len("Change SelText Color n Font")
            .hSubMenu = 0
        End With
        RetVal = InsertMenuItem(hPopupMenu2, 0, 1, mii2)
        
        With mii2 '(Change SelText Color More ...)
            .fType = MFT_STRING
            .fState = MFS_ENABLED
            .wID = 1002 ' Assign this item an item identifier.
            .dwTypeData = "Change SelText Color More ..."
            .cch = Len("Change SelText Color More ...")
            .hSubMenu = 0
        End With
        RetVal = InsertMenuItem(hPopupMenu2, 1, 1, mii2)
        
        With mii2 '(/separator/)
            .fType = MFT_SEPARATOR
            .fState = MFS_ENABLED
            .wID = 1003 ' Assign this item an item identifier.
            .dwTypeData = "/separator/"
            .cch = Len("/separator/")
            .hSubMenu = 0
        End With
        RetVal = InsertMenuItem(hPopupMenu2, 2, 1, mii2)
        
        With mii2 '(Select Skin)
            .fType = MFT_STRING
            .fState = MFS_ENABLED
            .wID = 1004 ' Assign this item an item identifier.
            .dwTypeData = "Select Skin"
            .cch = Len("Select Skin")
            .hSubMenu = hPopupMenu1
        End With
        RetVal = InsertMenuItem(hPopupMenu2, 3, 1, mii2)
        
        With mii1 '(Select Skin/Summer)
            .fType = MFT_STRING
            .fState = MFS_ENABLED
            .wID = 1005 ' Assign this item an item identifier.
            .dwTypeData = "Summer"
            .cch = Len("Summer")
            .hSubMenu = 0
        End With
        RetVal = InsertMenuItem(hPopupMenu1, 0, 1, mii1)
        
        With mii1 '(Select Skin/Spring)
            .fType = MFT_STRING
            .fState = MFS_ENABLED
            .wID = 1006 ' Assign this item an item identifier.
            .dwTypeData = "Spring"
            .cch = Len("Spring")
            .hSubMenu = 0
        End With
        RetVal = InsertMenuItem(hPopupMenu1, 1, 1, mii1)
        
        With mii1 '(Select Skin/Fall)
            .fType = MFT_STRING
            .fState = MFS_ENABLED
            .wID = 1007 ' Assign this item an item identifier.
            .dwTypeData = "Fall"
            .cch = Len("Fall")
            .hSubMenu = 0
        End With
        RetVal = InsertMenuItem(hPopupMenu1, 2, 1, mii1)
        
        With mii1
            .fType = MFT_STRING
            .fState = MFS_ENABLED
            .wID = 1008
            .dwTypeData = "From file ..."
            .cch = Len("From file ...")
            .hSubMenu = 0
        End With
        RetVal = InsertMenuItem(hPopupMenu1, 3, 1, mii1)
        
        With mii2 '(Discard Note)
            .fType = MFT_STRING
            .fState = MFS_ENABLED
            .wID = 1009 ' Assign this item an item identifier.
            .dwTypeData = "Discard Note"
            .cch = Len("Discard Note")
            .hSubMenu = 0
        End With
        RetVal = InsertMenuItem(hPopupMenu2, 4, 1, mii2)
        
        'The following code is for adding pictures into menus, if there are any!
        '------------------------------------------------------------
        '------------------------------------------------------------
        
        '------------------------------------------------------------
        '------------------------------------------------------------
        
        RetVal = GetCursorPos(curpos)
        menusel = TrackPopupMenu(hPopupMenu2, TPM_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_LEFTALIGN Or TPM_LEFTBUTTON, curpos.X, curpos.Y, 0, RTBNote.hwnd, 0)
        RetVal = DestroyMenu(hPopupMenu2)
        '------------------------------------------------------------------------------------------------
        'DOWN BELOW  PUT IN YOUR CODE MANUALY !!!!
        '------------------------------------------------------------------------------------------------
        Select Case menusel
            
            Case 1001 '(Change SelText Color & Font)
                ' Set Flags
                CommonDialog1.flags = cdlCFBoth Or cdlCFEffects
                CommonDialog1.CancelError = True
                CommonDialog1.ShowFont
                
                'If CommonDialog1.FontName = "" Then Exit Sub
                
                With Me.RTBNote
                    '   Change the text font according to options selected.
                    .Font.Name = CommonDialog1.FontName
                    .Font.Size = CommonDialog1.FontSize
                    .Font.Bold = CommonDialog1.FontBold
                    .Font.Italic = CommonDialog1.FontItalic
                    .Font.Underline = CommonDialog1.FontUnderline
                    .Font.Strikethrough = CommonDialog1.FontStrikethru
                    .SelColor = CommonDialog1.Color
                End With
                
            Case 1002 '(Change SelText ... More Color)
                CommonDialog1.flags = cdlCCRGBInit ' Or cdlCCFullOpen
                ' Set initial values for the dialog box.
                ' CommonDialog1.Color = RTBNote.SelColor
                With CommonDialog1
                    .CancelError = True
                    .ShowColor
                    RTBNote.SelLength = Len(RTBNote.SelText)
                    RTBNote.SelColor = .Color
                End With
                '    Call picSave_Click
                
            Case 1005 '(Select Skin/Summer)
                RTBNote.Visible = False
                imgSkin.Picture = ImgList.ListImages(1).Picture
                'Call WriteIniString(.Tag, "SkinFile", "2", inifilename)
                SaveSetting App.EXEName, "NoteSkin" & "\" & Me.Tag, "Value", "1"
                RTBNote.Visible = True
                
            Case 1006 '(Select Skin/Spring)
                RTBNote.Visible = False
                imgSkin.Picture = ImgList.ListImages(2).Picture
                'Call WriteIniString(.Tag, "SkinFile", "2", inifilename)
                SaveSetting App.EXEName, "NoteSkin" & "\" & Me.Tag, "Value", "2"
                RTBNote.Visible = True
                
            Case 1007 '(Select Skin/Fall)
                RTBNote.Visible = False
                imgSkin.Picture = ImgList.ListImages(3).Picture
                SaveSetting App.EXEName, "NoteSkin" & "\" & Me.Tag, "Value", "3"
                'Call WriteIniString(.Tag, "SkinFile", "3", inifilename)
                RTBNote.Visible = True
                                
            Case 1008
            Dim skin As String
            With CommonDialog1
                .CancelError = True
                .DialogTitle = "Choose Skin File"
                .InitDir = App.Path & "\Skins"
                .Filter = "GIF(*.gif)|*.gif|JPG (*.jpg)| *.jpg|JPEGS (*.jpeg)|*.jpeg |BMP(*.bmp)|*.bmp|All Files(*.*)|*.*"
                .FilterIndex = 1
                .ShowOpen
                skin = .FileName
            End With
                RTBNote.Visible = False
                imgSkin.Picture = LoadPicture(skin)
                SaveSetting App.EXEName, "NoteSkin" & "\" & Me.Tag, "Value", skin
                'Call WriteIniString(.Tag, "SkinFile", skin, inifilename)
                RTBNote.Visible = True

            Case 1009 '(Discard Note)
                Call picDump_Click
                
            Case Else
                
        End Select
        
        DestroyMenu hPopupMenu1
        DestroyMenu hPopupMenu2
    End If
Exit Sub
ErrorHandler:
If Err.Number = 481 Then
MsgBox "Invalid Skin File! Choose again.", vbInformation
RTBNote.Visible = True
Exit Sub
ElseIf Err.Number <> cdlCancel Then
MsgBox " System Error Number " & Err.Number _
        & " : " & Err.Description, vbInformation
End If
End Sub
