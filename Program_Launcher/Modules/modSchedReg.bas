Attribute VB_Name = "modSchedReg"
Option Explicit

Public Sub CreateKey(hKey As Long, strPath As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)

If lRegResult <> ERROR_SUCCESS Then
  ' there is a problem
End If

lRegResult = RegCloseKey(hCurKey)

End Sub

'Public Sub DeleteKey(ByVal hKey As Long, ByVal strPath As String)
'Dim lRegResult As Long
'lRegResult = RegDeleteKey(hKey, strPath)
'End Sub

Public Sub DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegOpenKey(hKey, strPath, hCurKey)

lRegResult = RegDeleteValue(hCurKey, strValue)

lRegResult = RegCloseKey(hCurKey)

End Sub

Public Function GetRegString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
Dim hCurKey As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long

' Set up default value
If Not IsEmpty(Default) Then
  GetRegString = Default
Else
  GetRegString = ""
End If

' Open the key and get length of string
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

  If lValueType = REG_SZ Then
    ' initialise string buffer and retrieve string
    strBuffer = String(lDataBufferSize, " ")
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
    
    ' format string
    intZeroPos = InStr(strBuffer, Chr$(0))
    If intZeroPos > 0 Then
      GetRegString = Left$(strBuffer, intZeroPos - 1)
    Else
      GetRegString = strBuffer
    End If

  End If

Else
  ' there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub SaveRegString(hKey As Long, strPath As String, strValue As String, strData As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))

If lRegResult <> ERROR_SUCCESS Then
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetRegLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long

Dim lRegResult As Long
Dim lValueType As Long
Dim lBuffer As Long
Dim lDataBufferSize As Long
Dim hCurKey As Long

' Set up default value
If Not IsEmpty(Default) Then
  GetRegLong = Default
Else
  GetRegLong = 0
End If

lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lDataBufferSize = 4       ' 4 bytes = 32 bits = long

lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

  If lValueType = REG_DWORD Then
    GetRegLong = lBuffer
  End If

Else
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)

End Function

Public Sub SaveRegLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal Ldata As Long)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0&, REG_DWORD, Ldata, 4)

If lRegResult <> ERROR_SUCCESS Then
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function CountRegKeys(hKey As Long, strPath As String) As Variant
' Returns: an count of all keys

Dim lRegResult As Long
Dim lCounter As Long
Dim hCurKey As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim strNames() As String
Dim intZeroPos As Integer

lCounter = 0

lRegResult = RegOpenKey(hKey, strPath, hCurKey)

Do

  'initialise buffers (longest possible length=255)
  lDataBufferSize = 255
  strBuffer = String(lDataBufferSize, " ")
  lRegResult = RegEnumKey(hCurKey, lCounter, strBuffer, lDataBufferSize)

  If lRegResult = ERROR_SUCCESS Then
  
    lCounter = lCounter + 1

  Else
    Exit Do
  End If
Loop

CountRegKeys = lCounter
End Function

Public Function GetRegKey(hKey As Long, strPath As String, RegKey) As Variant
' Returns: an array in a variant of strings

Dim lRegResult As Long
Dim lCounter As Long
Dim hCurKey As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim strNames() As String
Dim intZeroPos As Integer

lCounter = 0

lRegResult = RegOpenKey(hKey, strPath, hCurKey)

Do

  'initialise buffers (longest possible length=255)
  lDataBufferSize = 255
  strBuffer = String(lDataBufferSize, " ")
  lRegResult = RegEnumKey(hCurKey, lCounter, strBuffer, lDataBufferSize)

  If lRegResult = ERROR_SUCCESS Then
  
    'tidy up string and save it
    ReDim Preserve strNames(lCounter) As String
    If RegKey = lCounter Then
    GetRegKey = strBuffer
    Exit Do
    Else
    lCounter = lCounter + 1
    End If
  Else
    Exit Do
  End If
Loop

End Function
