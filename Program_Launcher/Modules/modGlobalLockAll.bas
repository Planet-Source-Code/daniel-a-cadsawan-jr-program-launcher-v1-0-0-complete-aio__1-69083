Attribute VB_Name = "modGlobalLockAll"
' in module
Private Const WH_KEYBOARD_LL = 13 ' LowLevel

Private Const VK_CONTROL = &H11
Private Const VK_DELETE = &H2E
Private Const VK_ESCAPE = &H1B
Private Const VK_LCONTROL = &HA2
Private Const VK_LEFT = &H25
Private Const VK_LSHIFT = &HA0
Private Const VK_LWIN = &H5B
Private Const VK_RCONTROL = &HA3
Private Const VK_RSHIFT = &HA1
Private Const VK_RWIN = &H5C
Private Const VK_SHIFT = &H10 'Left and right SHIFT keys
Private Const VK_TAB = &H9
Private Const VK_MENU = &H12 'ALT key
Private Const VK_CANCEL = &H3
Private Type KBDLLHOOKSTRUCT
vkCode As Long ' virtual key code
scanCode As Long ' scan code
flags As Long ' flags
time As Long ' time stamp for thismessage
dwExtraInfo As Long ' extra info from
' the driver or keybd_event
End Type
Private Const LLKHF_ALTDOWN = 32
Private Const HC_ACTION = 0
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)

Private hHook As Long
Private CTRLDown As Boolean
Private pkbhs As KBDLLHOOKSTRUCT

' wParam - Specifies the identifier of the
' keyboard message.
' This parameter can be one of the
' following messages:
' WM_KEYDOWN,WM_KEYUP,WM_SYSKEYDOWN,
' or WM_SYSKEYUP.
' lParam - pointer to KBDLLHOOKSTRUCT structure

Private Function LowLevelKeyboardProc(ByVal idHook As Integer, ByVal wParam As Long, ByVal lParam As Long) As Long
CTRLDown = False
CopyStructFromPtr pkbhs, lParam, Len(pkbhs)
Select Case idHook
Case HC_ACTION
If (GetAsyncKeyState(VK_CONTROL) And &HF0000000) Then CTRLDown = True
' Disable CTRL + ESC :
If (pkbhs.vkCode = VK_ESCAPE And CTRLDown) Then
LowLevelKeyboardProc = 1
Exit Function
End If
' Disable ATL+TAB
If (pkbhs.vkCode = VK_TAB And (pkbhs.flags And LLKHF_ALTDOWN)) Then
LowLevelKeyboardProc = 1
Exit Function
End If
'Disable ALT+ESC
If (pkbhs.vkCode = VK_ESCAPE And (pkbhs.flags And LLKHF_ALTDOWN)) Then
LowLevelKeyboardProc = 1
Exit Function
End If
' Disable the WINDOWS key
If (pkbhs.vkCode = VK_LWIN Or pkbhs.vkCode = VK_RWIN) Then
LowLevelKeyboardProc = 1
Exit Function
End If
If (pkbhs.vkCode = VK_CANCEL) Then ' disable ctrl + break
LowLevelKeyboardProc = 1
Exit Function
End If
'call the next hook
LowLevelKeyboardProc = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
Case Else
'call the next hook
LowLevelKeyboardProc = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
End Select
End Function

Public Sub HookKeyboard()
hHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0&)
End Sub
Public Sub UnHookKeyboard()
'remove the windows-hook
UnhookWindowsHookEx hHook
End Sub


