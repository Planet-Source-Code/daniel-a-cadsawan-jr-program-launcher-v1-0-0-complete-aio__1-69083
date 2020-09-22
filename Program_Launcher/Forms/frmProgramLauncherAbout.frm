VERSION 5.00
Begin VB.Form frmProgramLauncherAbout 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "About MyApp"
   ClientHeight    =   4575
   ClientLeft      =   2295
   ClientTop       =   1500
   ClientWidth     =   6255
   ClipControls    =   0   'False
   Icon            =   "frmProgramLauncherAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin ProgramLauncher.xFrame xFrameAbout 
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8070
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
      Picture         =   "frmProgramLauncherAbout.frx":2D0A
      Begin VB.PictureBox picIcon1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1230
         Left            =   120
         Picture         =   "frmProgramLauncherAbout.frx":3B5C
         ScaleHeight     =   842.8
         ScaleMode       =   0  'User
         ScaleWidth      =   4214
         TabIndex        =   10
         Top             =   480
         Width           =   6030
      End
      Begin VB.PictureBox picIcon2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1680
         Left            =   120
         Picture         =   "frmProgramLauncherAbout.frx":12683
         ScaleHeight     =   1650
         ScaleWidth      =   1410
         TabIndex        =   9
         Top             =   1800
         Width           =   1440
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "frmProgramLauncherAbout.frx":19F36
         Top             =   3720
         Width           =   4695
      End
      Begin VB.CommandButton cmdSysInfo 
         Caption         =   "&System Info..."
         Height          =   345
         Left            =   4920
         TabIndex        =   3
         Top             =   4110
         Width           =   1245
      End
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   345
         Left            =   4920
         TabIndex        =   2
         Top             =   3720
         Width           =   1260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   2
         X1              =   120
         X2              =   6120
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label lblDescription 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Description: Program Launcher is use to launch pre defined external applications."
         ForeColor       =   &H00000000&
         Height          =   585
         Left            =   1680
         TabIndex        =   8
         Top             =   2880
         Width           =   4395
      End
      Begin VB.Label lblAuthor 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Author: Daniel A. Cadsawan Jr."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   2520
         Width           =   4245
      End
      Begin VB.Label lblVersion 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   2160
         Width           =   4365
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "App Title:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1680
         TabIndex        =   5
         Top             =   1800
         Width           =   4365
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   120
         X2              =   6120
         Y1              =   3600
         Y2              =   3600
      End
   End
   Begin VB.TextBox txtWarning 
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmProgramLauncherAbout.frx":1A64E
      Top             =   3360
      Width           =   4695
   End
End
Attribute VB_Name = "frmProgramLauncherAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options... '
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
    KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
    KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' Reg Key ROOT Types... '
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                   ' Unicode nul terminated string '
Const REG_DWORD = 4                ' 32-bit number                 '

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Sub cmdSysInfo_Click()
    On Error Resume Next
    Call StartSysInfo
End Sub
Public Sub MakeTopMost(hwnd As Long)
    On Error Resume Next
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    Unload frmProgramLauncherAbout
End Sub


'CODE : fade form on loading
Public Sub Form_Activate()
    On Error Resume Next
    Dim iX As Long
    Static bActive As Boolean
    If bActive Then Exit Sub
    bActive = True
    For iX = 1 To 255 Step 2 '2
        MakeWindowTransparent frmProgramLauncherAbout.hwnd, iX
        DoEvents    ' need this so form doesn't turn black
    Next
    
End Sub
Private Sub Form_Load()
    On Error Resume Next
    'CODE : position the form
    frmProgramLauncherAbout.Left = _
        (Screen.Width - frmProgramLauncherAbout.Width - frmProgramLauncher.Width) - 200
    frmProgramLauncherAbout.Top = _
        (Screen.Height - frmProgramLauncherAbout.Height) - 1000
    
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    xFrameAbout.Caption = "About " & App.Title
    
    ' start off transparent so form doesn't flicker
    MakeWindowTransparent frmProgramLauncherAbout.hwnd, 2 '10
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'CODE : fade form on unload
    Dim iX As Long
    For iX = 255 To 0 Step -2 '-2
        MakeWindowTransparent frmProgramLauncherAbout.hwnd, iX
    Next
    'Unload our form completely
    Set frmProgramLauncherAbout = Nothing
    Unload frmProgramLauncherAbout
    
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
    
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry... '
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry... '
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version '
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
            ' Error - File Can Not Be Found... '
        Else
            GoTo SysInfoErr
        End If
        ' Error - Registry Entry Can Not Be Found... '
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
    Resume Next
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                  ' Loop Counter                             '
    Dim rc As Long                 ' Return Code                              '
    Dim hKey As Long               ' Handle To An Open Registry Key           '
    Dim hDepth As Long             '                                          '
    Dim KeyValType As Long         ' Data Type Of A Registry Key              '
    Dim tmpVal As String           ' Tempory Storage For A Registry Key Value '
    Dim KeyValSize As Long         ' Size Of Registry Key Variable            '
    ' ------------------------------------------------------------ '
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}            '
    ' ------------------------------------------------------------ '
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)   ' Open Registry Key '
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError   ' Handle Error... '
    
    tmpVal = String$(1024, 0)      ' Allocate Variable Space '
    KeyValSize = 1024              ' Mark Variable Size      '
    
    ' ------------------------------------------------------------ '
    ' Retrieve Registry Key Value...                               '
    ' ------------------------------------------------------------ '
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
        KeyValType, tmpVal, KeyValSize)   ' Get/Create Key Value '
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError   ' Handle Errors '
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then   ' Win95 Adds Null Terminated String...    '
    tmpVal = Left(tmpVal, KeyValSize - 1)       ' Null Found, Extract From String         '
Else                                            ' WinNT Does NOT Null Terminate String... '
    tmpVal = Left(tmpVal, KeyValSize)           ' Null Not Found, Extract String Only     '
End If
    ' ------------------------------------------------------------ '
    ' Determine Key Value Type For Conversion...                   '
    ' ------------------------------------------------------------ '
    Select Case KeyValType                                      ' Search Data Types...               '
        Case REG_SZ                                             ' String Registry Key Data Type      '
            KeyVal = tmpVal                                     ' Copy String Value                  '
        Case REG_DWORD                                          ' Double Word Registry Key Data Type '
            For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit                   '
                KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.         '
            Next
            KeyVal = Format$("&h" + KeyVal)   ' Convert Double Word To String '
    End Select
    
    GetKeyValue = True             ' Return Success     '
    rc = RegCloseKey(hKey)         ' Close Registry Key '
    Exit Function                  ' Exit               '
    
GetKeyError:                                   ' Cleanup After An Error Has Occured... '
    KeyVal = ""                    ' Set Return Val To Empty String        '
    GetKeyValue = False            ' Return Failure                        '
    rc = RegCloseKey(hKey)         ' Close Registry Key                    '
End Function


Private Sub xFrameAbout_Click()
    On Error Resume Next
    Unload frmProgramLauncherAbout
End Sub
