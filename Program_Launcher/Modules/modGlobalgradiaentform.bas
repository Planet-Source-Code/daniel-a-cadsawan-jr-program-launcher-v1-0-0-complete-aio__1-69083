Attribute VB_Name = "modGlobalgradientform"
' The following code snippet paints a gradient between 2 colors on the entire form.                      '
' It uses Win32 API in order to get the best performances.                                               '
' You can change the colors of the gradient by changing the RGB values passed to PaintGradient function. '
'                                                                                                        '
' Written by Nir Sofer                                                                                   '
' Web site: http://nirsoft.cjb.net                                                                       '


Option Explicit
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Function SafeDiv(X1 As Double, X2 As Double) As Double
    If X2 = 0 Then SafeDiv = 0 Else SafeDiv = X1 / X2
End Function

Public Sub PaintGradient(frm As Form, Red1 As Integer, Green1 As Integer, Blue1 As Integer, Red2 As Integer, Green2 As Integer, Blue2 As Integer)
    Dim WinRect     As RECT
    Dim ColorRect   As RECT
    Dim Y           As Long
    Dim hBrush      As Long
    Dim hPrevBrush  As Long
    Dim DivValue    As Double
    Dim CurrRed     As Integer
    Dim CurrGreen   As Integer
    Dim CurrBlue    As Integer
    
    GetClientRect frm.hWnd, WinRect
    For Y = WinRect.Top To WinRect.Bottom
        DivValue = SafeDiv((WinRect.Bottom - WinRect.Top), (Y - WinRect.Top))
        ' Calculate the Red, Green and Blue values. '
        CurrRed = Red1 + SafeDiv((Red2 - Red1), DivValue)
        CurrGreen = Green1 + SafeDiv((Green2 - Green1), DivValue)
        CurrBlue = Blue1 + SafeDiv((Blue2 - Blue1), DivValue)
        SetRect ColorRect, WinRect.Left, Y, WinRect.Right, Y + 1
        ' Create the brush for the current color '
        hBrush = CreateSolidBrush(RGB(CurrRed, CurrGreen, CurrBlue))
        ' Select the brush into the DC '
        hPrevBrush = SelectObject(frm.hdc, hBrush)
        ' Draw the line of gradient '
        FillRect frm.hdc, ColorRect, hBrush
        SelectObject frm.hdc, hPrevBrush
        ' Release the brush '
        DeleteObject hBrush
    Next
End Sub
