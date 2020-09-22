Attribute VB_Name = "modUserDomainPW"

Option Explicit

Private Const VER_PLATFORM_WIN32_NT                 As Long = 2
Private Const HKEY_LOCAL_MACHINE                    As Long = &H80000002
Private Const KEY_READ                              As Long = &H20019
Private Const mcstrAgentKey                         As String * 56 = "System\CurrentControlSet\Services\MSNP32\NetworkProvider"

Private Type OS_VERSION_INFO
    dwOSVersionInfoSize     As Long
    dwMajorVersion          As Long
    dwMinorVersion          As Long
    dwBuildNumber           As Long
    dwPlatformId            As Long
    szCSDVersion            As String * 128
End Type

Private Type WKSTA_USER_INFO_1
    wkui1_username          As Long
    wkui1_logon_domain      As Long
    wkui1_oth_domains       As Long
    wkui1_logon_server      As Long
End Type

'Common APIs
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nsize As Long) As Long
Private Declare Function GetPlatform Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OS_VERSION_INFO) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nsize As Long) As Long

'WinNT/2000/XP APIs
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function NetApiBufferFree Lib "Netapi32.dll" (ByVal lpBuffer As Long) As Long
Private Declare Function NetWkstaUserGetInfo Lib "Netapi32.dll" (ByVal reserved As Any, ByVal level As Long, lpBuffer As Any) As Long

'Win9x/ME APIs
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Public Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long

Const LOGON32_LOGON_INTERACTIVE As Long = 2
Const LOGON32_LOGON_NETWORK As Long = 3
Const LOGON32_PROVIDER_DEFAULT As Long = 0
Const LOGON32_PROVIDER_WINNT50 As Long = 3
Const LOGON32_PROVIDER_WINNT40 As Long = 2
Const LOGON32_PROVIDER_WINNT35 As Long = 1

Public Function CurrentDomain() As String
    Dim lngStructPtr        As Long
    Dim udtUserInfo         As WKSTA_USER_INFO_1
    Select Case NetApiSupport
        Case True:  NetWkstaUserGetInfo 0&, 1&, lngStructPtr
        Case False: Win9xDomainName CurrentDomain
    End Select
    If lngStructPtr = 0 Then Exit Function
    CopyMemory udtUserInfo, ByVal lngStructPtr, Len(udtUserInfo)
    CurrentDomain = StrFromPtr(udtUserInfo.wkui1_logon_domain)
    NetApiBufferFree lngStructPtr
End Function

Private Function StrFromPtr(lngPtr As Long) As String
    Dim bytString()      As Byte
    Dim lngBytes         As Long
    If lngPtr = 0 Then Exit Function
    lngBytes = lstrlenW(lngPtr) * 2
    If lngBytes = 0 Then Exit Function
    ReDim bytString(0 To (lngBytes - 1)) As Byte
    CopyMemory bytString(0), ByVal lngPtr, lngBytes
    StrFromPtr = bytString
End Function

Private Function NetApiSupport() As Boolean
    On Error Resume Next
    Dim udtOS       As OS_VERSION_INFO
    udtOS.dwOSVersionInfoSize = Len(udtOS)
    If Not (GetPlatform(udtOS) = 1) Then Exit Function
    NetApiSupport = (udtOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Private Sub Win9xDomainName(ByRef strDomainName As String)
    On Error Resume Next
    Dim strRegKey           As String
    Dim lngRegKeyHdl        As Long
    Dim lngRegKeyLen        As Long
    Dim lngRegDataType      As Long
    strRegKey = Space(255)
    lngRegKeyLen = Len(strRegKey)
    RegOpenKeyEx HKEY_LOCAL_MACHINE, ByVal mcstrAgentKey, 0&, KEY_READ, lngRegKeyHdl
    RegQueryValueEx lngRegKeyHdl, "AuthenticatingAgent", 0, lngRegDataType, ByVal strRegKey, lngRegKeyLen
    RegCloseKey lngRegKeyHdl
    strDomainName = Left$(strRegKey, lngRegKeyLen - 1)
    strDomainName = Trim$(strDomainName)
End Sub

Public Function CurrentLogonUser() As String
    Dim UserLoginName As String
    UserLoginName = Space(200)
    Call GetUserName(UserLoginName, 200)
    UserLoginName = Trim$(UserLoginName)
    UserLoginName = Mid$(UserLoginName, 1, Len(UserLoginName) - 1)
    CurrentLogonUser = UCase$(UserLoginName)
End Function


Public Function VerifyLogin(sUser As String, sDomain As String, sPassword As String) As Boolean
    Dim token As Long
    VerifyLogin = LogonUser(sUser, sDomain, sPassword, LOGON32_LOGON_NETWORK, LOGON32_PROVIDER_DEFAULT, token)
End Function


Public Function ReadSetting(ByVal Name As String, Optional ByVal ReadEncypted As Boolean = False) As String
    Dim Str As String, Pos As Long, strRndNum As String, RndNum As Double
    
    Str = String(1024, 0)
    GetPrivateProfileString App.EXEName, Name, "", Str, Len(Str), _
    inifilepw
    
    Pos = InStr(1, Str, Chr(0))
    Str = Left(Str, Pos - 1)
    
    If ReadEncypted And Len(Str) > 0 Then
        strRndNum = Space(64)
        GetPrivateProfileString App.EXEName, Name & "_RND", "", strRndNum, Len(strRndNum), _
        inifilepw
        
        RndNum = val(Trim(strRndNum))
        Str = RC4(Str, CStr(RndNum * 0.945))
    End If
    
    ReadSetting = Str
End Function
 
Public Sub WriteSetting(ByVal Name As String, ByVal Str As String, Optional ByVal SaveEncypted As Boolean = False)
    Dim RndNum As Double, K As Long
    
    If SaveEncypted And Len(Str) > 0 Then
        Randomize
        
        For K = 1 To 10
            RndNum = RndNum + ((2 ^ 15) * Rnd)
        Next K
        
        WritePrivateProfileString App.EXEName, Name & "_RND", CStr(RndNum), _
        inifilepw
        
        Str = RC4(Str, CStr(RndNum * 0.945))
    End If
    
    WritePrivateProfileString App.EXEName, Name, Str, _
    inifilepw
End Sub
 
Public Function RC4(ByVal Expression As String, ByVal Password As String) As String
    Dim RB(0 To 255) As Integer
    Dim X As Long, Y As Long, z As Long
    Dim Key() As Byte, ByteArray() As Byte, Temp As Byte
    
    On Error Resume Next
    
    If Len(Password) = 0 Then Exit Function
    If Len(Expression) = 0 Then Exit Function
    
    If Len(Password) > 256 Then
        Key() = StrConv(Left$(Password, 256), vbFromUnicode)
    Else
        Key() = StrConv(Password, vbFromUnicode)
    End If
    
    For X = 0 To 255
        RB(X) = X
    Next X
    
    X = 0
    Y = 0
    z = 0
    
    For X = 0 To 255
        Y = (Y + RB(X) + Key(X Mod Len(Password))) Mod 256
        
        Temp = RB(X)
        RB(X) = RB(Y)
        RB(Y) = Temp
    Next X
    
    X = 0
    Y = 0
    z = 0
    
    ByteArray() = StrConv(Expression, vbFromUnicode)
    
    For X = 0 To Len(Expression)
        Y = (Y + 1) Mod 256
        z = (z + RB(Y)) Mod 256
        
        Temp = RB(Y)
        RB(Y) = RB(z)
        RB(z) = Temp
        
        ByteArray(X) = ByteArray(X) Xor (RB((RB(Y) + RB(z)) Mod 256))
    Next X
    
    RC4 = StrConv(ByteArray, vbUnicode)
    If Err.Number <> 0 Then Err.Clear
End Function

