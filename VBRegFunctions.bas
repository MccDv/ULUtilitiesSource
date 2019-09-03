Attribute VB_Name = "VBRegFunctions"
Public Const HKCR = 0
Public Const HKCU = 1
Public Const HKLM = 2
Public Const HKU = 3
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
   
'Define severity codes
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_ACCESS_DENIED = 5
Public Const ERROR_MORE_DATA = 234
Public Const ERROR_NO_MORE_ITEMS = 259
Public Const ERROR_ALREADY_EXISTS = 183&
   
'Structures Needed For Registry Prototypes
Public Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type
   
Public Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
   
'Registry Function Prototypes
   Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
     (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
   Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
     (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
     ByVal samDesired As Long, phkResult As Long) As Long
   Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" _
     (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
      ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
   Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
   Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
     (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
      ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
   Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" _
     (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
      ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
      lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, _
      lpdwDisposition As Long) As Long
   Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
     (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
      lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, _
      lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
   Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
     (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
      lpcbName As Long) As Long
   Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
     (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
      lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
      ByVal lpData As String, lpcbData As Long) As Long
   Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
     (ByVal hKey As Long, ByVal lpSubKey As String) As Long
   Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
     (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Function GetKeyValue(ByVal hKey As Long, ByRef RegNode _
   As String, ByRef ValName As String) As Boolean

   Dim phkResult As Long, Result As Long
   Dim lpcbData As Long
   Dim ValFound As Boolean
   
   ValFound = False
   
   Result = RegOpenKey(hKey, RegNode, phkResult)
   If Result = ERROR_SUCCESS Then
      lpcbData = 255
      szData$ = Space$(lpcbData)
      Result = RegQueryValueEx(phkResult, ValName, _
         0, 1, szData, lpcbData)
      If Result = ERROR_SUCCESS Then
         ValName = Left$(szData$, lpcbData& - 1)
         ValFound = True
      End If
      Result = RegCloseKey(phkResult)
   End If
   GetKeyValue = ValFound
   
End Function

Public Function FindSubNodes(ByVal hKey As Long, ByRef NodeToTest _
   As String, ByRef ReturnedKeyNames As Variant) As Long

   Dim phkResult As Long, Result As Long
   Dim lpcbData As Long, Index As Long
   Dim NumKeysFound As Long, Loc As Long
   Dim FT As FILETIME
   Dim SKeys() As String
   Dim szBuffer As String, SubEnum As String
   Dim lBuffSize As Long
   
   Index = 0
   Result = RegOpenKey(hKey, NodeToTest, phkResult)
   If Result = ERROR_SUCCESS Then
      While Result = ERROR_SUCCESS
         szBuffer = Space(255)
         lBuffSize = Len(szBuffer)
         Result = RegEnumKey(phkResult, Index, _
            szBuffer, lBuffSize)
         If Result = ERROR_SUCCESS Then
            Loc& = InStr(1, szBuffer, vbNullChar)
            SubEnum = Left$(szBuffer, Loc& - 1)
            ReDim Preserve SKeys(Index)
            SKeys(Index) = SubEnum
            Index = Index + 1
         End If
      Wend
      ReturnedKeyNames = SKeys()
      NumKeysFound = Index - 1
      Result = RegCloseKey(phkResult)
   End If
   FindSubNodes = NumKeysFound
   
End Function


