VERSION 5.00
Begin VB.UserControl Slider 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   ScaleHeight     =   2175
   ScaleWidth      =   3180
   ToolboxBitmap   =   "Slider.ctx":0000
   Begin VB.Image ImgCur 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      Picture         =   "Slider.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   225
   End
   Begin VB.Image SliderBack 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   0
      Picture         =   "Slider.ctx":31E5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Slider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: Horizontal Slider
' Nome dei Files..:
' Data............: 27/10/2004
' Versione........: 1.0
' Sistema.........: Windows
' Scritto da......: Bruno Crepaldi ®
' E-Mail..........: bruno.crepax@libero.it
'=====================================================
'
'                Not For Commercial Use
'=====================================================
'
Option Explicit
'
Private M_Value As Long
Private M_MinValue As Long
Private M_MaxValue As Long
'
Private CursRaporto As Double
Private CursRange As Long
Private CursBlk As Boolean
'                                Dichiarazione Eventi
Public Event Change(Value As Long)
'
'      Inizializza le Variabili ( Solo Progetazione )
'
Private Sub UserControl_InitProperties()
    M_Value = 0                   ' Valore Iniziale
    M_MinValue = 0                ' Valore Iniziale
    M_MaxValue = 10               ' Valore Iniziale
    
    UserControl.Height = 255      ' Altezza
    UserControl.Width = 1830      ' Larghezza
End Sub

'
'                                Property
'
'
Public Property Get Value() As Long
    Value = M_Value
End Property
Public Property Let Value(ByVal NewValue As Long)
    
    If NewValue > M_MaxValue Then NewValue = M_MaxValue
    If NewValue < M_MinValue Then NewValue = M_MinValue
    
    M_Value = NewValue
    PropertyChanged "Value"
    Call Sposta((M_Value - M_MinValue) * CursRaporto)
End Property
'
Public Property Get MinValue() As Long
    MinValue = M_MinValue
End Property
Public Property Let MinValue(ByVal NewValue As Long)
    M_MinValue = NewValue
    PropertyChanged "MinValue"
    CursRaporto = Raporto(M_MinValue, M_MaxValue)
End Property
'
Public Property Get MaxValue() As Long
    MaxValue = M_MaxValue
End Property
Public Property Let MaxValue(ByVal NewValue As Long)
    M_MaxValue = NewValue
    PropertyChanged "MaxValue"
    CursRaporto = Raporto(M_MinValue, M_MaxValue)
End Property
'
Public Property Get Picture() As Picture
    Set Picture = SliderBack.Picture
End Property

Public Property Set Picture(ByVal NewPic As Picture)
    Set SliderBack.Picture = NewPic
    PropertyChanged "Picture"
End Property
'
'                 Read/Write Properties
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    M_Value = PropBag.ReadProperty("Value", 0)
    M_MinValue = PropBag.ReadProperty("MinValue", 0)
    M_MaxValue = PropBag.ReadProperty("MaxValue", 10)
    
    CursRaporto = Raporto(M_MinValue, M_MaxValue)
    
    Set SliderBack.Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", M_Value, 0)
    Call PropBag.WriteProperty("MinValue", M_MinValue, 0)
    Call PropBag.WriteProperty("MaxValue", M_MaxValue, 10)
    Call PropBag.WriteProperty("Picture", SliderBack.Picture, Nothing)
End Sub
'
'                        Eventi
'
Private Sub ChangeEvent(Valore As Long)
    RaiseEvent Change(Valore)
End Sub
'
'                        Resizing
'
Private Sub UserControl_Resize()
    SliderBack.Left = 0
    SliderBack.Top = 0
    SliderBack.Width = ScaleWidth
    SliderBack.Height = ScaleHeight
    ImgCur.Left = 0
    ImgCur.Height = ScaleHeight
End Sub
'
'                        Inizio
'
Public Sub SliderBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CursBlk = True
End Sub

Public Sub SliderBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SliderBack_MouseMove(Button, Shift, X, Y)
    CursBlk = False
End Sub

Public Sub SliderBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim MaxDX As Long
    Dim MinSX As Long
    '
    If CursBlk = False Then Exit Sub
    MaxDX = ScaleWidth - (ImgCur.Width / 2)
    MinSX = (ImgCur.Width / 2)
    Select Case X
        Case Is < MinSX              ' Minimo
            ImgCur.Left = 0
            M_Value = M_MinValue
            GoTo SetValue
        Case Is > MaxDX              ' Massimo
            ImgCur.Left = ScaleWidth - ImgCur.Width
            M_Value = M_MaxValue
            GoTo SetValue
    End Select
    
    Call Sposta(X - MinSX)
    M_Value = (ImgCur.Left / CursRaporto) + M_MinValue
    
SetValue:
    Call ChangeEvent(Value)
End Sub

Private Sub Sposta(Posizione As Long)
    ImgCur.Left = Posizione
End Sub
Private Function Raporto(Min As Long, Max As Long) As Single
    CursRange = Max - Min
    Raporto = (ScaleWidth - ImgCur.Width) / CursRange
End Function

