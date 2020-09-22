Attribute VB_Name = "modTMReg"
Option Explicit

Public Sub CreateRegLong(ByVal EnmHive As RegistryHives, ByVal StrSubKey As String, ByVal strValueName As String, ByVal LngData As Long, Optional ByVal EnmType As RegistryLongTypes = REG_DWORD_LITTLE_ENDIAN)
    Dim hKey As Long
    Call CreateSubKey(EnmHive, StrSubKey)
    hKey = GetSubKeyHandle(EnmHive, StrSubKey, KEY_ALL_ACCESS)
    RegSetValueEx hKey, strValueName, 0, EnmType, LngData, 4
    RegCloseKey hKey
End Sub

Public Sub CreateSubKey(ByVal EnmHive As RegistryHives, ByVal StrSubKey As String)
    Dim hKey As Long
    RegCreateKey EnmHive, StrSubKey & Chr(0), hKey
    RegCloseKey hKey
End Sub

Private Function GetSubKeyHandle(ByVal EnmHive As RegistryHives, ByVal StrSubKey As String, Optional ByVal EnmAccess As RegistryKeyAccess = KEY_READ) As Long
    Dim hKey As Long
    Dim RetVal As Long
    RetVal = RegOpenKeyEx(EnmHive, StrSubKey, 0, EnmAccess, hKey)
    If RetVal <> ERROR_SUCCESS Then
        hKey = 0
    End If
    GetSubKeyHandle = hKey
End Function

Public Function SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
    Dim lRetVal As Long
    Dim hKey As Long
    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)
End Function

Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    Select Case lType
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
    End Select
End Function

Public Function DeleteKey(lPredefinedKey As RegistryHives, sKeyName As String)
    Dim lRetVal As Long         'result of the SetValueEx function
    Dim hKey As Long         'handle of open key
    lRetVal = RegDeleteKey(lPredefinedKey, sKeyName)
End Function

Public Function RemoveRegSubKeys(ByVal eHive As RegistryHives, ByVal sKey As String) As Boolean
    Dim lRegKey As Long, lRegType As Long
    Dim sValue As String, lValueLen As Long
    Dim sData() As Byte, lDataLen As Long
    Dim bFailed As Boolean
    Dim sName As String, lNameLen As Long
    Dim sClass As String, lClassLen As Long
    Dim tFILETIME As FILETIME
    ' Open a Handle to the Recent File List Registry "Key" ("RecentFileList")
    If RegOpenKeyEx(eHive, sKey, 0, KEY_ALL_ACCESS, lRegKey) <> 0 Then Exit Function
    sClass = Space(260)
    sName = Space(260)
    lClassLen = 260
    lNameLen = 260
    ' Enumerate the next SubKey
    Do While RegEnumKeyEx(lRegKey, 0, ByVal sName, lNameLen, 0, ByVal sClass, lClassLen, tFILETIME) = ERROR_SUCCESS
        ' Attempt to Delete it...
        If RegDeleteKey(lRegKey, Left(sName, lNameLen)) <> ERROR_SUCCESS Then
            ' Failed to Delete, Exit and return Failure
            bFailed = True
            Exit Do
        End If
        sClass = Space(260)
        sName = Space(260)
        lClassLen = 260
        lNameLen = 260
    Loop
    ' Close the Key Handle
    Call RegCloseKey(lRegKey)
    ' Return Success/Failure Result
    RemoveRegSubKeys = (Not bFailed)
End Function
