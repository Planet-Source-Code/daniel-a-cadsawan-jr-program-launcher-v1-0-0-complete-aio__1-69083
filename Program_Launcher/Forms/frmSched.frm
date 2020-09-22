VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSched 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   7065
   ClientLeft      =   -45
   ClientTop       =   -435
   ClientWidth     =   6135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHide 
      Cancel          =   -1  'True
      Caption         =   "Hide Scheduler"
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   5160
      Width           =   1455
   End
   Begin ProgramLauncher.xFrame xFrame1 
      Height          =   5235
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8387
      BorderColor     =   7645851
      Button          =   -1  'True
      ButtonColor     =   4487268
      ButtonHighlightColor=   7645851
      ButtonPin       =   -1  'True
      ColorScheme     =   2
      Caption         =   "Set New PrLr Scheduler Notice"
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
      ForeColor       =   4487268
      FontItalic      =   0   'False
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      GradientBottom  =   14938092
      HeaderGradientBottom=   7975330
      HeaderGradientTop=   14938092
      Picture         =   "frmSched.frx":0000
      Begin VB.CommandButton cmdTestStop 
         Caption         =   "SoundTest Stop"
         Height          =   375
         Left            =   4560
         TabIndex        =   38
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdTestPlay 
         Caption         =   "SoundTest Play"
         Height          =   375
         Left            =   3120
         TabIndex        =   37
         Top             =   1680
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpTime 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16384002
         CurrentDate     =   39248
      End
      Begin VB.ComboBox cbAudio 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Text            =   "(None)"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox setDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Text            =   "Date"
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox setTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "Time"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox tbxNotifyText 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "Reminder!"
         Top             =   4800
         Width           =   5895
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Schedule"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4560
         TabIndex        =   3
         Top             =   5160
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Set Schedule"
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   5160
         Width           =   1455
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2310
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   16711680
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MonthColumns    =   2
         MonthBackColor  =   12574687
         ShowWeekNumbers =   -1  'True
         StartOfWeek     =   16384001
         TitleBackColor  =   -2147483645
         TitleForeColor  =   16777215
         TrailingForeColor=   8421504
         CurrentDate     =   37136
         MinDate         =   -52196
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Schedule Message:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   4560
         Width           =   1665
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Schedule Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1380
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Schedule Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Schedule Sound:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   10
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1080
      End
   End
   Begin ProgramLauncher.xFrame xFrame2 
      Height          =   5355
      Left            =   0
      TabIndex        =   15
      Top             =   300
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9446
      Button          =   -1  'True
      ButtonPin       =   -1  'True
      Caption         =   "View PrLr Scheduler Notices"
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
      HeaderGradientBottom=   12611136
      Picture         =   "frmSched.frx":059A
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Schedule"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4560
         TabIndex        =   18
         Top             =   4860
         Width           =   1455
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change Schedule"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   17
         Top             =   4860
         Width           =   1455
      End
      Begin ComctlLib.ListView lvAlerts 
         Height          =   4095
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Alert Date"
            Object.Width           =   4587
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Alert Time"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Alert Message"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Alert Sound"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Schedules:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1380
      End
   End
   Begin ProgramLauncher.xFrame xFrame3 
      Height          =   5085
      Left            =   0
      TabIndex        =   20
      Top             =   600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8969
      BackColor       =   16777215
      BorderColor     =   12298664
      ButtonColor     =   8283750
      ButtonHighlightColor=   12298664
      ColorScheme     =   3
      Caption         =   "View PrLr Scheduler Settings"
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
      ForeColor       =   8283750
      FontItalic      =   0   'False
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      GradientBottom  =   15920108
      HeaderGradientBottom=   14140358
      HeaderGradientTop=   16118000
      Picture         =   "frmSched.frx":0B34
      Begin VB.PictureBox xpFrame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1665
         ScaleWidth      =   5865
         TabIndex        =   27
         Top             =   2760
         Width           =   5895
         Begin VB.CheckBox cbxAlarmMode 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enable Alarm Clock Mode."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   480
            Width           =   2535
         End
         Begin VB.CheckBox cbxStartUp 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Do not Load PrLr Scheduler when Program Launcher Starts."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1200
            Width           =   4815
         End
         Begin VB.CheckBox cbxCleanUp 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Automaticaly Delete Expired Notices when Scheduler Loads."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   840
            Width           =   4815
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Other PrLr Scheduler Settings"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   120
            Width           =   2130
         End
      End
      Begin VB.PictureBox frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   120
         ScaleHeight     =   2025
         ScaleWidth      =   5865
         TabIndex        =   21
         Top             =   480
         Width           =   5895
         Begin VB.OptionButton optSoundDir 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Use Window Default Sounds Directory"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   480
            Value           =   -1  'True
            Width           =   3255
         End
         Begin VB.OptionButton optSoundDir 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Use MP3 Music Directory As Default"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   25
            Top             =   840
            Width           =   3255
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Browse"
            Height          =   375
            Left            =   4320
            TabIndex        =   24
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CheckBox chkMp3Finish 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Always Allow MP3 To Finish Playing After Notifying."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   1680
            Width           =   3975
         End
         Begin VB.TextBox xplMP3Dir 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   240
            TabIndex        =   22
            Top             =   1200
            Width           =   5535
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PrLr Scheduler Sound Settings"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   120
            Width           =   2160
         End
      End
      Begin VB.Label LBLdsd 
         BackStyle       =   0  'Transparent
         Caption         =   "Default Sound Directory:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   2295
      End
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   34
      Top             =   6690
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   5116
            MinWidth        =   5116
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   2118
            MinWidth        =   2118
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   8008
            MinWidth        =   8008
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox pichook 
      Height          =   615
      Left            =   2640
      Picture         =   "frmSched.frx":10CE
      ScaleHeight     =   555
      ScaleWidth      =   315
      TabIndex        =   35
      Top             =   5880
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   2040
      Top             =   5760
   End
   Begin VB.Timer tmrCaption 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   2040
      Top             =   6120
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   0
      TabIndex        =   39
      Top             =   5760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Image IconImage 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   0
      Left            =   3120
      Picture         =   "frmSched.frx":2DC8
      Top             =   6120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image IconImage 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   1
      Left            =   3360
      Picture         =   "frmSched.frx":3352
      Top             =   6120
      Visible         =   0   'False
      Width           =   240
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   735
      Left            =   1440
      TabIndex        =   36
      Top             =   5760
      Width           =   495
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Image IconImage 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   2
      Left            =   3360
      Picture         =   "frmSched.frx":38DC
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image IconImage 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   3
      Left            =   3120
      Picture         =   "frmSched.frx":3E66
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image IconImage 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   4
      Left            =   3600
      Picture         =   "frmSched.frx":43F0
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmSched"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


''''start browse for folder'''''''
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
'Private Const BIF_BROWSEFORCOMPUTER = &H1000

Private Const MAX_PATH = 260
Private Declare Function SHBrowseForFolder Lib "shell32" _
    (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, ByVal _
    lpString2 As String) As Long

Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
    End Type
''''''''end browse for folder'''''

Private f_Change As Long

Dim i As Long
'Dim sIcon' as Variant
Dim sIcon As Long
Dim CountIt As String, FirstAlert As String

Private Sub cbAudio_Click()
'since user clicked hour combobox then we enable clear button
cmdSave.Enabled = True
cmdClear.Enabled = True
'MonthView1.SetFocus
'MsgBox cbAudio.Text
End Sub

Private Sub cmdTestPlay_Click()
On Error GoTo ErrorHandler
Dim Mfile As String
Mfile = File1.Path & "\" & cbAudio.List(cbAudio.ListIndex)
MediaPlayer1.FileName = Mfile
MediaPlayer1.Play
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & vbCrLf _
        & Err.Description & vbCrLf _
        & "Either you don't have a filename or" & vbCrLf _
        & "No MediaPlayer is installed!", vbInformation
End Sub

Private Sub cmdTestStop_Click()
MediaPlayer1.Stop
End Sub
Private Sub cbxAlarmMode_Click()

On Error Resume Next
SaveRegLong HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "AlarmMode", cbxAlarmMode.Value

If cbxAlarmMode.Value = 1 Then
optSoundDir(0).Enabled = False
optSoundDir(1).Enabled = False
xplMP3Dir.Enabled = False
cmdBrowse.Enabled = False
chkMp3Finish.Enabled = False
frmNotify.cbxAlwaysPlay.Enabled = False
frame7.Enabled = False
LBLdsd.Enabled = False
Else

optSoundDir(0).Enabled = True
optSoundDir(1).Enabled = True
xplMP3Dir.Enabled = True
cmdBrowse.Enabled = True
chkMp3Finish.Enabled = True
frmNotify.cbxAlwaysPlay.Enabled = True
frame7.Enabled = True
LBLdsd.Enabled = True
End If

End Sub

Private Sub cbxCleanUp_Click()
SaveRegLong HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "AutoCleanUp", cbxCleanUp.Value

End Sub

Private Sub cbxStartUp_Click()
If cbxStartUp.Value = 1 Then
SaveRegLong HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "AutoStart", cbxStartUp.Value

'SaveRegString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "PrLrScheduler", App.Path & "\" & App.EXEName & ".exe"
Else
'DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "PrLrScheduler"
SaveRegLong HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "AutoStart", cbxStartUp.Value

End If
End Sub

Private Sub chkMp3Finish_Click()
SaveRegLong HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "AlwaysPlay", chkMp3Finish.Value

End Sub

Private Sub cmdBrowse_Click()
Dim FolderPath As String
FolderPath = BrowseFolder
If FolderPath = "" Then Exit Sub
xplMP3Dir.Text = FolderPath
SaveRegString HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "MP3Directory", xplMP3Dir.Text

cbAudio.Clear
If optSoundDir(1).Value = False Then
File1.Pattern = "*.wav"
File1.Path = "C:\WINDOWS\Media"
Else
File1.Pattern = "*.mp3;*.wma;*.wav"
File1.Path = xplMP3Dir.Text
End If

For i = 0 To File1.ListCount - 1
File1.ListIndex = i
cbAudio.AddItem File1.FileName
Next i

cbAudio.Text = "(None)"
End Sub

Private Sub cmdClear_Click()
On Error Resume Next
cbAudio.Text = "(None)"
tbxNotifyText.Text = "Reminder!"
MonthView1.Value = Date
setDate.Text = "Date"
setTime.Text = "Time"
dtpTime.Value = time

'if user selected change but then cleared, we wanna unselect the notice
If f_Change >= 1 Then
lvAlerts.SelectedItem.Selected = False
f_Change = 0
End If
cmdClear.Enabled = False
cmdSave.Enabled = False

'dtpTime.SetFocus
End Sub

Private Sub cmdDelete_Click()

On Error Resume Next

If lvAlerts.SelectedItem.Selected = False Then
MsgBox "Please Select Notice To Delete", vbQuestion
Exit Sub
End If

'Delete alert from registry
DeleteKey HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & lvAlerts.SelectedItem.Text & " - " & lvAlerts.SelectedItem.SubItems(1)

'Delete alert from listview
lvAlerts.ListItems.Remove (lvAlerts.SelectedItem.Index)

'update labels
StatusBar.Panels(1).Text = "You have a total of " & lvAlerts.ListItems.count & " active notices."
    
'if no alerts, then disable buttons
If lvAlerts.ListItems.count = 0 Then
cmdDelete.Enabled = False
cmdChange.Enabled = False
'Switch to Set Alert tab to modify the settings
xFrame1.Expanded = True
End If

End Sub


Private Sub cmdHide_Click()
frmSched.tmrCaption.Enabled = False
frmSched.Hide
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
'if user selected change notice and is now saving it, we wanna delete the
'old notice

If f_Change >= 1 Then

'Delete alert from registry
DeleteKey HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & lvAlerts.SelectedItem.Text & " - " & lvAlerts.SelectedItem.SubItems(1)

'Delete alert from listview
lvAlerts.ListItems.Remove (f_Change)

'update labels
StatusBar.Panels(1).Text = "You have a total of " & lvAlerts.ListItems.count & " active notices."
    
'if no alerts, then disable buttons
If lvAlerts.ListItems.count = 0 Then
cmdDelete.Enabled = False
cmdChange.Enabled = False
End If
If lvAlerts.ListItems.count >= 1 Then lvAlerts.SelectedItem.Selected = False

f_Change = 0
End If
'end update notice


'dont save alert if user didnt set date
If setDate.Text = "Date" Then MsgBox "You Must Set The Alert Date!!", vbInformation: Exit Sub
If setTime.Text = "Time" Then MsgBox "You Must Set The Alert Time!!", vbInformation: Exit Sub

'user selected invalid time
If NoticeExpired(ShortDate(setDate.Text), setTime.Text) = True Then
MsgBox "You Selected A Time That Has Already Past, Please Select New Time", vbInformation
dtpTime.SetFocus
Exit Sub
End If


'if timer was turned off, then we turn it on
If Timer2.Enabled = False Then Timer2.Enabled = True


'load new alert to list view
If cbAudio.List(cbAudio.ListIndex) = "" Then sIcon = 2 Else sIcon = 1

With lvAlerts.ListItems.Add(, , setDate.Text, , sIcon)
    .SubItems(1) = setTime.Text
    .SubItems(2) = tbxNotifyText.Text
    If sIcon = 1 Then
    .SubItems(3) = File1.Path & "\" & cbAudio.List(cbAudio.ListIndex)
    End If

End With

'save new alert to registry
SaveRegString HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & setDate.Text & " - " & setTime.Text, "AlertTime", setTime.Text
SaveRegString HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & setDate.Text & " - " & setTime.Text, "AlertDate", setDate.Text
SaveRegString HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & setDate.Text & " - " & setTime.Text, "AlertMessage", tbxNotifyText.Text
SaveRegString HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & setDate.Text & " - " & setTime.Text, "AlertSound", File1.Path & "\" & cbAudio.List(cbAudio.ListIndex)

'send click to clear set alert input settings and disable save button
cmdClear_Click '

'enable delete button
If cmdDelete.Enabled = False Then cmdDelete.Enabled = True
If cmdChange.Enabled = False Then cmdChange.Enabled = True
'update labels
StatusBar.Panels(1).Text = "You have a total of " & lvAlerts.ListItems.count & " active notices."
tbxNotifyText.Text = "Reminder!"

'set text and date fields according to dtp and monthview
setTime.Text = Format(dtpTime.Value, "h:mm AM/PM")
setDate.Text = Format(MonthView1.Value, "Long Date")

End Sub






Private Sub cmdChange_Click()
On Error Resume Next

If lvAlerts.SelectedItem.Selected = False Then
    MsgBox "Please Select Notice To Change", vbInformation
    Exit Sub
End If

f_Change = lvAlerts.SelectedItem.Index


setTime.Text = lvAlerts.SelectedItem.SubItems(1)

setDate.Text = lvAlerts.SelectedItem.Text

'Load sound event to cbAudio
For i = 1 To cbAudio.ListCount

'If lvAlerts.SelectedItem.SubItems(3) = "" Then
cbAudio.Text = "(None)"
'ElseIf cbAudio.List(i) = lvAlerts.SelectedItem.SubItems(3) Then
'cbAudio.ListIndex = i
'End If

Next i

'Load text message from list view to set alert message
tbxNotifyText.Text = lvAlerts.SelectedItem.SubItems(2)

'load date to monthview1
MonthView1.Value = ShortDate(lvAlerts.SelectedItem.Text)
'you edited the shortdate sub

dtpTime.Value = lvAlerts.SelectedItem.SubItems(1)

'Switch to Set Alert tab to modify the settings
xFrame1.Expanded = True

'cbAudio.Text = GetRegString(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & lvAlerts.SelectedItem.Text & " - " & lvAlerts.SelectedItem.SubItems(1), "AlertSound", "")

End Sub
Private Function BrowseFolder()
On Error Resume Next

'Opens a Treeview control that displays
    '     the directories in a computer
    Dim lpIDList As Long
    Dim sbuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    szTitle = "Choose PrLr Scheduler Sound Folder"


    With tBrowseInfo
        .hwndOwner = Me.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
        
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)


    If (lpIDList) Then
        sbuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sbuffer
        sbuffer = Left(sbuffer, InStr(sbuffer, vbNullChar) - 1)
        
        BrowseFolder = sbuffer
        
    End If
    
End Function
Private Function ShortDate(nvalue As String)
On Error Resume Next

If nvalue = "" Then Exit Function

'this is to convert date (Tuesday, September 28, 2001) to (09/28/2001)
Dim syear, smonth, sday


syear = Right(nvalue, 4)

sday = Left(Right(nvalue, 8), 2)

smonth = Right(nvalue, Len(nvalue) - InStr(nvalue, ",") - 1)
smonth = Left(smonth, Len(smonth) - 9)

Select Case RTrim(LTrim(smonth))

Case "January"
smonth = 1
Case "February"
smonth = 2
Case "March"
smonth = 3
Case "April"
smonth = 4
Case "May"
smonth = 5
Case "June"
smonth = 6
Case "July"
smonth = 7
Case "August"
smonth = 8
Case "September"
smonth = 9
Case "October"
smonth = 10
Case "November"
smonth = 11
Case "December"
smonth = 12
End Select

ShortDate = smonth & "/" & sday & "/" & syear

End Function


Private Function NoticeExpired(comDate As String, comTime As String) As Boolean
On Error Resume Next
Dim nDate As String, tDate As String, nyear As String, tyear As String
Dim ntime As String, ttime As String

nDate = Format(comDate, "MM/DD/YYYY")
tDate = Format(Date, "MM/DD/YYYY")
ntime = Format(comTime, "HH:mm")
ttime = Format(time, "HH:mm")

'years
nyear = Right(nDate, 4)
tyear = Right(tDate, 4)



'yesterday or older we flag
If nDate < tDate And nyear <= tyear Then
NoticeExpired = True
Exit Function
End If

'today but with older time we flag
If nDate = tDate And nyear = tyear And ntime < ttime Then
NoticeExpired = True
Exit Function
End If



'if we got this far, the notice is valid
NoticeExpired = False


End Function

Private Sub dtpTime_Change()
setTime.Text = Format(dtpTime.Value, "h:mm AM/PM")
End Sub

Private Sub dtpTime_Click()
cmdSave.Enabled = True
cmdClear.Enabled = True
End Sub

Private Sub Form_Load()

On Error Resume Next

'here we save our hwnd, i know thers a better way of doing this
'SaveRegLong HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "hWnd", Me.hWnd

Dim rTime As String, rDate As String
'RemoveMenu GetSystemMenu(Me.hWnd, 0), 6, MF_BYPOS
'RemoveMenu GetSystemMenu(Me.hWnd, 0), 5, MF_BYPOS
Me.Height = 120 * 50
Me.Width = 6135
xFrame1.Height = Me.Height - 375
xFrame2.Height = Me.Height - 300 - 375
xFrame3.Height = Me.Height - 600 - 375
'Frame1.Left = 240
'Frame2.Left = 240
'Frame3.Left = 240
'Me.Height = 7995
'Me.Width = 6720

    ' Load pictures into the ImageList.
    For i = 0 To 4
        ImageList1.ListImages.Add , , IconImage(i).Picture
    Next i
    
'load sound directory if any
xplMP3Dir.Text = GetRegString(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "MP3Directory")

'load user selected default
optSoundDir(1).Value = GetRegLong(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "UseMP3Dir")
optSoundDir(0).Value = GetRegLong(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "UseWaveDir")

If optSoundDir(0).Value = False And optSoundDir(1).Value = False Then optSoundDir(0).Value = True

'now we check for user default directory, if not set then we set to windows default
If optSoundDir(1).Value = False Then
File1.Pattern = "*.wav"
File1.Path = "C:\WINDOWS\Media"
xplMP3Dir.Enabled = False
cmdBrowse.Enabled = False
Else
File1.Pattern = "*.mp3;*.wma;*.wav"
File1.Path = xplMP3Dir.Text
xplMP3Dir.Enabled = True
cmdBrowse.Enabled = True
End If

'we load autoclean up
cbxCleanUp.Value = GetRegLong(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "AutoCleanUp")
    
'if user selected to autoclean, then we do it
If cbxCleanUp.Value = 1 Then
retry:
    CountIt = CountRegKeys(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts")
    For i = 0 To CountIt - 1
    FirstAlert = GetRegKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts", i)
    rDate = GetRegString(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & FirstAlert, "AlertDate")
    rTime = GetRegString(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & FirstAlert, "AlertTime")
    'if notice is expired we erase and start check over
    If NoticeExpired(ShortDate(rDate), rTime) = True Then
    DeleteKey HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & FirstAlert
GoTo retry
End If

Next i
End If


'get the amount of notices
CountIt = CountRegKeys(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts")

'disable delete button if not notices
If CountIt = 0 Then cmdDelete.Enabled = False


'load saved alerts to listview
For i = 0 To CountIt - 1
FirstAlert = GetRegKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts", i)



rDate = GetRegString(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & FirstAlert, "AlertDate")
rTime = GetRegString(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & FirstAlert, "AlertTime")



If GetRegString(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & FirstAlert, "AlertSound") = "" Then sIcon = 2 Else sIcon = 1

    With lvAlerts.ListItems.Add(, , rDate, , sIcon)
        .SubItems(1) = rTime
        .SubItems(2) = GetRegString(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & FirstAlert, "AlertMessage")
        .SubItems(3) = GetRegString(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & FirstAlert, "AlertSound")
    End With


Next i
    



'load alarm mode setting
cbxAlarmMode.Value = GetRegLong(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "AlarmMode")

StatusBar.Panels(1).Text = "You have a total of " & CountIt & " active notices."
    
'If no saved alerts then kill timer
If lvAlerts.ListItems.count = 0 Then Timer2.Enabled = False
    
'load always play
chkMp3Finish.Value = GetRegLong(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "AlwaysPlay")


'set both date and time pickers to current time and date
dtpTime.Value = time

MonthView1.Value = Date

'set cbx if PrLrScheduler is loaded at start up
cbxStartUp.Value = GetRegLong(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "AutoStart")


'add audio fils to combobox
cbAudio.Clear
For i = 0 To File1.ListCount - 1
File1.ListIndex = i
cbAudio.AddItem File1.FileName
Next i


'enable or disable delete and  change
If lvAlerts.ListItems.count = 0 Then
cmdDelete.Enabled = False
cmdChange.Enabled = False
Else
cmdDelete.Enabled = True
cmdChange.Enabled = True

End If

'cbHour.SetFocus

'set text and date fields according to dtp and monthview
setTime.Text = Format(dtpTime.Value, "h:mm AM/PM")
setDate.Text = Format(MonthView1.Value, "Long Date")

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

'c_CANCEL = False
Timer2.Enabled = False
tmrCaption.Enabled = False

Unload frmSched

End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
'check for valid date
If MonthView1.Value < Date Then
MsgBox "Sorry, PrLrScheduler Cannot Be Set For A Date That Has Already Past", vbExclamation, "PrLrScheduler Error"
MonthView1.SetFocus
Exit Sub
End If


setDate.Text = Format(MonthView1.Value, "Long Date")

'since user clicked monthview then we enable clear button
cmdSave.Enabled = True
cmdClear.Enabled = True

End Sub

Private Sub optSoundDir_Click(Index As Integer)
On Error Resume Next

Select Case Index

Case 0
optSoundDir(0).Value = True
optSoundDir(1).Value = False
SaveRegLong HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "UseWaveDir", optSoundDir(0).Value
SaveRegLong HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "UseMP3Dir", 0
xplMP3Dir.Enabled = False
cmdBrowse.Enabled = False

Case 1
optSoundDir(0).Value = False
optSoundDir(1).Value = True

SaveRegLong HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "UseMP3Dir", optSoundDir(1).Value
SaveRegLong HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "UseWaveDir", 0
xplMP3Dir.Enabled = True
cmdBrowse.Enabled = True
'If xplMP3Dir.Text = "" Then cmdBrowse_Click

End Select

cbAudio.Clear
If optSoundDir(1).Value = False Then
File1.Pattern = "*.wav"
File1.Path = "C:\WINDOWS\Media"
Else
File1.Pattern = "*.mp3;*.wma;*.wav"
File1.Path = xplMP3Dir.Text
End If

For i = 0 To File1.ListCount - 1
File1.ListIndex = i
cbAudio.AddItem File1.FileName
Next i

cbAudio.Text = "(None)"

End Sub

Private Sub tbxNotifyText_Change()
'since user clicked our combobox then we enable clear button
cmdSave.Enabled = True
cmdClear.Enabled = True
End Sub

Private Sub Timer2_Timer()
On Error Resume Next

For i = 1 To lvAlerts.ListItems.count
If lvAlerts.ListItems(i).SubItems(1) = Format(time, "h:mm AM/PM") And lvAlerts.ListItems(i).Text = Format(Date, "Long Date") Then

If cbxAlarmMode.Value = 0 Then
'''Check to see if we have an audible alert
    If lvAlerts.ListItems(i).SubItems(3) > "" Then
    MediaPlayer1.FileName = lvAlerts.ListItems(i).SubItems(3)
    'MediaPlayer1.Volume = 127
    MediaPlayer1.PlayCount = 1
    MediaPlayer1.Play
    End If

Else
    MediaPlayer1.FileName = "C:\WINDOWS\Media\notify.wav"
    'MediaPlayer1.Volume = 127
    MediaPlayer1.PlayCount = 0
    MediaPlayer1.Play

End If

'''Check to see if we have a text message to display
    If lvAlerts.ListItems(i).SubItems(2) > "" Then
    frmNotify.lblText = lvAlerts.ListItems(i).SubItems(2)
    frmNotify.show
    Else
    frmNotify.lblText = "PrLrScheduler was set to alert you of sumthing,, but since you didnt say what, PrLrScheduler can not tell you. You must have got high."
    frmNotify.show
    End If

'Delete alert from registry
DeleteKey HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler\Alerts\" & lvAlerts.ListItems(i).Text & " - " & lvAlerts.ListItems(i).SubItems(1)

'Delete alert from listview
lvAlerts.ListItems.Remove i

If lvAlerts.ListItems.count = 0 Then cmdDelete.Enabled = False

End If

Next i

'kill timer when no alerts listed
If lvAlerts.ListItems.count = 0 Then Timer2.Enabled = False

End Sub

Private Sub tmrCaption_Timer()
StatusBar.Panels(2).Text = time
StatusBar.Panels(3).Text = Format(Date, "Long Date")
End Sub


Private Sub xFrame1_Click()
    If xFrame1.Expanded = True Then
        xFrame1.Height = Me.Height - 375
        dtpTime.Value = Now
        'set text and date fields according to dtp and monthview
        setTime.Text = Format(dtpTime.Value, "h:mm AM/PM")
        setDate.Text = Format(MonthView1.Value, "Long Date")
    End If
End Sub

Private Sub xFrame2_Click()
    If xFrame2.Expanded = True Then
        xFrame2.Height = Me.Height - 300 - 375
    End If
End Sub

Private Sub xFrame3_Click()
    If xFrame3.Expanded = True Then
        xFrame3.Height = Me.Height - 600 - 375
    End If
End Sub

Private Sub xFrame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    End If

End Sub

Private Sub xFrame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    End If

End Sub

Private Sub xFrame3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    End If

End Sub


Private Sub xplMP3Dir_Change()
SaveRegString HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\PrLrScheduler", "MP3Directory", xplMP3Dir.Text
End Sub

