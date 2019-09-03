Attribute VB_Name = "SYSINF32"
'This module is compatable with:
'  VB version 4.0 and up
'  No library Dependencies


'information types
Global Const OPSYS = &H100
Global Const DLLINF = &H400
Global Const VXDINF = &H500
Global Const REGINF = &H600
Public Const FILE_READ_ATTRIBUTES = (&H80)              '  all
Public Const GENERIC_READ = (&H80000000)              '  all
Public Const OPEN_EXISTING = 3
Public Const INVALID_HANDLE_VALUE = -1

Global Const FILEDATE = 1
Global Const FILESIZE = 2
Global Const FILEREV = 3

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

Global Const GMEM_DISCARDED = &H4000
Global Const GMEM_ZEROINIT = &H40
Global Const GMEM_MOVEABLE = &H2
Global Const GMEM_FIXED = 0
Public Const GWL_STYLE = (-16)
Public Const WS_CAPTION = &HC00000      '  WS_BORDER Or WS_DLGFRAME

'Define severity codes
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_ACCESS_DENIED = 5
Public Const ERROR_NO_MORE_ITEMS = 259
Public Const ERROR_ALREADY_EXISTS = 183&

Public Const READ_WRITE = 2
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SECTION_EXTEND_SIZE = &H10
Public Const SECTION_MAP_EXECUTE = &H8
Public Const SECTION_MAP_READ = &H4
Public Const SECTION_MAP_WRITE = &H2
Public Const SECTION_QUERY = &H1
Public Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
Public Const FILE_MAP_READ = SECTION_MAP_READ
Public Const FILE_MAP_WRITE = SECTION_MAP_WRITE
Public Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS

Public Const PAGE_NOACCESS = &H1
Public Const PAGE_READONLY = &H2
Public Const PAGE_READWRITE = &H4
Public Const PAGE_WRITECOPY = &H8
Public Const PAGE_EXECUTE = &H10
Public Const PAGE_EXECUTE_READ = &H20
Public Const PAGE_EXECUTE_READWRITE = &H40
Public Const PAGE_EXECUTE_WRITECOPY = &H80
Public Const PAGE_GUARD = &H100
Public Const PAGE_NOCACHE = &H200
Public Const MEM_COMMIT = &H1000
Public Const MEM_RESERVE = &H2000
Public Const MEM_DECOMMIT = &H4000
Public Const MEM_RELEASE = &H8000
Public Const MEM_FREE = &H10000
Public Const MEM_PRIVATE = &H20000
Public Const MEM_MAPPED = &H40000
Public Const MEM_MAPPED_COPIED = &H80000
Public Const MEM_TOP_DOWN = &H100000
Public Const MEM_LARGE_PAGES = &H20000000
Public Const MEM_DOS_LIM = &H40000000

Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Public Const FORMAT_MESSAGE_FROM_STRING = &H400
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

' dwPlatforID Constants
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const INFINITE = &HFFFF

Public Declare Function OpenProcess Lib "kernel32" _
(ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" _
(ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Const DRIVE_CDROM As Long = 5
'2 = "Removable"
'3 = "Drive Fixed"
'4 = "Remote"
'5 = "Cd-Rom"
'6 = "Ram disk"
'Case Else "Unrecognized"

Global gsOS As String
Global gnNumDrvs As Integer
Global gsDvrList() As String
Global gsUInstPath As String
Global gsUInstAltPath As String
Global gnDriverUtilOnCD As Integer
Global gsUserDirec As String
Global gnWow64 As Integer

Dim mfNoForm As Form
Dim msSysDirec As String, msWinDirec As String

Dim msVxDPath As String, msVxDName As String
Dim mnVxDFound As Integer, mnVxDValid As Integer
Dim msKeys() As String

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

Public Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Public Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName(32) As Integer
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(32) As Integer
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

Public Type VERINFO                      'Version FIXEDFILEINFO
    strPad1 As Long               'Pad out struct version
    strPad2 As Long               'Pad out struct signature
    nMSLo As Integer              'Low word of file ver # MS DWord
    nMSHi As Integer              'High word of file ver # MS DWord
    nLSLo As Integer              'Low word of file ver # LS DWord
    nLSHi As Integer              'High word of file ver # LS DWord
    nPVMSLo As Integer           'Low word of product ver # MS DWord
    nPVMSHi As Integer           'Low word of product ver # MS DWord
    nPVLSLo As Integer           'Low word of product ver # MS DWord
    nPVLSHi As Integer           'Low word of product ver # MS DWord
    strPad3(1 To 28) As Byte      'Pad out rest of VERINFO struct (36 bytes)
End Type

Public Type MEMORY_BASIC_INFORMATION
     BaseAddress As Long
     AllocationBase As Long
     AllocationProtect As Long
     RegionSize As Long
     State As Long
     Protect As Long
     lType As Long
End Type

Public Type SYSTEM_INFO
   dwOemID As Long
   dwPageSize As Long
   lpMinimumApplicationAddress As Long
   lpMaximumApplicationAddress As Long
   dwActiveProcessorMask As Long
   dwNumberOrfProcessors As Long
   dwProcessorType As Long
   dwAllocationGranularity As Long
   dwReserved As Long
End Type

  Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
  End Type
   
Public Declare Function GetDriveType Lib "kernel32" _
   Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

#If Win32 Then ' 32-bit VB uses this Declare.
   Public Const WIN32APP = True
   Declare Function GetModuleFileName Lib "kernel32" Alias _
         "GetModuleFileNameA" (ByVal hModule As Long, _
         ByVal lpFileName As String, ByVal nSize As Long) As Long
   Declare Function GetModuleHandle Lib "kernel32" Alias _
         "GetModuleHandleA" (ByVal lpModuleName As String) As Long
   Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
   Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
   Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
   Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
   Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
   Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
   Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
   Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
   Declare Function GetVersion Lib "kernel32" () As Long
   Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
   Declare Function GlobalFlags Lib "kernel32" (ByVal hMem As Long) As Long
   Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
   Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
   Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
   Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
   Declare Function GetFileVersionInfo Lib "VERSION.DLL" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
   Declare Function GetFileVersionInfoSize Lib "VERSION.DLL" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
   Declare Function VerQueryValue Lib "VERSION.DLL" Alias "VerQueryValueA" (lpvVerData As Byte, ByVal lpszSubBlock As String, lplpBuf As Long, lpcb As Long) As Long
'   Declare Sub lmemcpy Lib "STKIT432.DLL" (strDest As Any, ByVal strSrc As Any, ByVal lBytes As Long)
   Declare Sub lmemcpy Lib "VB5STKIT.DLL" (strDest As Any, ByVal strSrc As Any, ByVal lBytes As Long)
'   Declare Sub lmemcpy Lib "VB6STKIT.DLL" (strDest As Any, ByVal strSrc As Any, ByVal lBytes As Long)
   Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long
   Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
   Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
   Declare Function VirtualQuery Lib "kernel32" (lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
   Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
  Private Declare Function GetVersionEx Lib "kernel32" _
      Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Boolean

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

   Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
   Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
      (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
      lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, _
      ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

   Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" _
      (ByVal hFile As Long, lpFileMappingAttributes As SECURITY_ATTRIBUTES, _
      ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, _
      ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
   Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, _
      ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, _
      ByVal dwNumberOfBytesToMap As Long) As Long
   Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" _
      (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
   Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
   Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long

   Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, _
   lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
   ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

   Declare Function GetLastError Lib "kernel32" () As Long
   Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
   (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
   Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, _
   lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, _
   lpLastWriteTime As FILETIME) As Long
   Public Declare Function FileTimeToSystemTime Lib "kernel32" _
   (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
   Public Declare Function SystemTimeToTzSpecificLocalTime _
   Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, _
   lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long

#Else
   Public Const WIN32APP = False
   Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal WinDirec$, ByVal BufferSize%)
   Declare Function GetSystemDirectory% Lib "Kernel" (ByVal SysDirec$, ByVal BufferSize%)
   Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
   Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Integer
   Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
   Declare Function GetVersion Lib "Kernel" () As Long
   Declare Function GlobalSize Lib "Kernel" (ByVal hMem As Integer) As Long
   Declare Function GlobalFlags Lib "Kernel" (ByVal hMem As Integer) As Integer
   Declare Function GlobalAlloc Lib "Kernel" (ByVal wFlags As Integer, ByVal dwBytes As Long) As Integer
   Declare Function GetFileVersionInfoSize Lib "ver.dll" (ByVal lpszFileName As String, lpdwHandle As Long) As Long
   Declare Function GetFileVersionInfo Lib "ver.dll" (ByVal lpszFileName As String, ByVal lpdwHandle As Long, ByVal cbbuf As Long, ByVal lpvdata As String) As Integer
   Declare Function VerQueryValue Lib "ver.dll" (ByVal lpvBlock As String, ByVal lpszSubBlock As String, lplpBuffer As Long, lpcb As Integer) As Integer
   Declare Function lstrcpyn Lib "Kernel" (ByVal lpszString1 As Any, ByVal lpszString2 As Long, ByVal cChars As Integer) As Long
   Declare Function GetWindowLong Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Long
   Declare Function SetWindowLong Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Long) As Long
   Declare Sub lmemcpy Lib "STKIT416.DLL" (strDest As Any, ByVal strSrc As Any, ByVal intBytes As Integer)
#End If

'-----------------------------------------------------------
' FUNCTION: GetFileVersion
'
' Returns the internal file version number for the specified
' file.  This can be different than the 'display' version
' number shown in the File Manager File Properties dialog.
' It is the same number as shown in the VB4 SetupWizard's
' File Details screen.  This is the number used by the
' Windows VerInstallFile API when comparing file versions.
'
' IN: [strFileName] - the file whose version # is desired
'     [fIsRemoteServerSupportFile] - whether or not this file is
'          a remote OLE automation server support file (.VBR)
'          (Enterprise edition only).  If missing, False is assumed.
'
' Returns: The Version number string if found, otherwise
'          gstrNULL
'-----------------------------------------------------------
'
Sub GetFileVersion(ByVal strFileName As String, FileProps() As String, Optional ByVal fIsRemoteServerSupportFile)
   Dim sVerInfo As VERINFO
   Dim strVer As String
   Dim lVerSize As Long
   Dim lVerHandle As Long
   Dim lpBufPtr As Long
   Dim byteVerData() As Byte
   Dim byteCharData() As Byte

   On Error GoTo GFVError

   If IsMissing(fIsRemoteServerSupportFile) Then
       fIsRemoteServerSupportFile = False
   End If
   
   lVerSize = GetFileVersionInfoSize(strFileName, lVerHandle)
   If lVerSize > 0 Then
      ReDim byteVerData(lVerSize)
      If GetFileVersionInfo(strFileName, lVerHandle, lVerSize, byteVerData(0)) <> 0 Then ' (Pass byteVerData array via reference to first element)
         If VerQueryValue(byteVerData(0), "\", lpBufPtr, lVerSize) <> 0 Then
            lmemcpy sVerInfo, lpBufPtr, lVerSize
            FileVerStruct% = True
         End If
      End If
   Else
      FileProps(1) = ""
   End If
   
   If FileVerStruct% Then
      With sVerInfo
         MSHi$ = .nMSHi
         If Val(.nMSHi) < 0 Then MSHi$ = Format(65536 + Val(.nMSHi))
         MSLo$ = .nMSLo
         If Val(.nMSLo) < 0 Then MSLo$ = Format(65536 + Val(.nMSLo))
         LSHi$ = .nLSHi
         If Val(.nLSHi) < 0 Then LSHi$ = Format(65536 + Val(.nLSHi))
         LSLo$ = .nLSLo
         If Val(.nLSLo) < 0 Then LSLo$ = Format(65536 + Val(.nLSLo))
         PVMSHi$ = .nPVMSHi
         If Val(.nPVMSHi) < 0 Then PVMSHi$ = Format(65536 + Val(.nPVMSHi))
         PVMSLo$ = .nPVMSLo
         If Val(.nPVMSLo) < 0 Then PVMSLo$ = Format(65536 + Val(.nPVMSLo))
         PVLSHi$ = .nPVLSHi
         If Val(.nPVLSHi) < 0 Then PVLSHi$ = Format(65536 + Val(.nPVLSHi))
         PVLSLo$ = .nPVLSLo
         If Val(.nPVLSLo) < 0 Then PVLSLo$ = Format(65536 + Val(.nPVLSLo))
      End With
      strVer = MSHi$ & "." & MSLo$ & "."
      strVer = strVer & LSHi$ & "." & LSLo$
      strPVer = PVMSHi$ & "." & PVMSLo$ & "."
      strPVer = strPVer & PVLSHi$ & "." & PVLSLo$
      FileProps(1) = strVer
      FileProps(2) = strPVer
      If VerQueryValue(byteVerData(0), "\VarFileInfo\Translation", lpBufPtr, lVerSize) <> 0 Then
         ReDim byteCharData(lVerSize)
         lmemcpy byteCharData(0), lpBufPtr, lVerSize
         Trans$ = Format$(Hex$(byteCharData(1)), "00") & Format$(Hex$(byteCharData(0)), "00") & Format$(Hex$(byteCharData(3)), "00") & Format$(Hex$(byteCharData(2)), "00")
      Else
         'assume US English and Windows, Multilingual (ANSI)
         Trans$ = "040904E4"
      End If
      For InfType% = 1 To 7
         CurString$ = ""
         TypeString$ = Choose(InfType%, "\CompanyName", "\ProductName", "\FileDescription", "\ProductVersion", "\FileVersion", "\Comments", "\LegalTrademarks")
         If VerQueryValue(byteVerData(0), "\StringFileInfo\" & Trans$ & TypeString$, lpBufPtr, lVerSize) <> 0 Then
            ReDim byteCharData(lVerSize)
            lmemcpy byteCharData(0), lpBufPtr, lVerSize
            For CharCode% = 0 To lVerSize - 1
               If byteCharData(CharCode%) < 32 Then Exit For
               CurString$ = CurString$ + Chr$(byteCharData(CharCode%))
            Next
            FileProps(InfType% + 2) = CurString$
         End If
      Next InfType%
   End If
   
   Exit Sub
    
GFVError:
   If Err = 48 Or Err = 53 Then
      'MsgBox "Can't find DLL 'STKIT432.DLL'.  This file must be in the current directory (" & CurDir$() & "), Windows, Windows\System directory or a directory included in the path.", , Error$(Err)
      MsgBox "Can't find DLL 'VB5STKIT.DLL'.  This file must be in the current directory (" & CurDir$() & "), Windows, Windows\System directory or a directory included in the path.", , Error$(Err)
      'MsgBox "Can't find DLL 'VB6STKIT.DLL'.  This file must be in the current directory (" & CurDir$() & "), Windows, Windows\System directory or a directory included in the path.", , Error$(Err)
      End
   End If
   Err = 0
   
End Sub

Function GetFileVerStruct(ByVal strFileName As String, sVerInfo As VERINFO, Optional ByVal fIsRemoteServerSupportFile) As Boolean


End Function

Public Sub GetDirectory(WinDirec$, SysDirec$)

   'used by GetDLLInfo and GetVxDInfo to determine Windows
   'and System directories for locating DLLs and system.ini

#If Win32 Then
   BufferSize& = 80
   WinDirec$ = Space$(BufferSize&)
   Size& = GetWindowsDirectory&(WinDirec$, BufferSize&)
   WinDirec$ = Left$(WinDirec$, Size&) & "\"
   SysDirec$ = Space$(BufferSize&)
   Size& = GetSystemDirectory&(SysDirec$, BufferSize&)
   SysDirec$ = Left$(SysDirec$, Size&) & "\"
#Else
   BufferSize% = 80
   WinDirec$ = Space$(BufferSize%)
   Size% = GetWindowsDirectory%(WinDirec$, BufferSize%)
   WinDirec$ = Left$(WinDirec$, Size%) & "\"
   SysDirec$ = Space$(BufferSize%)
   Size% = GetSystemDirectory%(SysDirec$, BufferSize%)
   SysDirec$ = Left$(SysDirec$, Size%) & "\"
#End If
   

End Sub

Private Sub GetDLLInfo(TypeOfInfo%, DLLInfo As Variant)

   'returns DLL information on the DLL specified by DLLInfo
   'in the form 'path$, date$, size$, rev$ (derived from time)'
   
   On Error GoTo LookElsewhere
   
   GetDirectory WinDirec$, SysDirec$
   
   'for DLL, search first in current directory, then Windows,
   'then Windows\System, then path

   Path$ = Environ$("PATH")
   DLLDir$ = CurDir$ & "\"
   DLLName$ = DLLInfo
   attempt% = 1
   RawDLLDate$ = FileDateTime(DLLDir$ & DLLName$)
   If NoSuccess% Then
      DLLInfo = "Not Found"
      Exit Sub
   End If
   DDate$ = Format$(RawDLLDate$, "Short Date")
   DTime$ = Format$(RawDLLDate$, "Short Time")
   DSize$ = Format$(FileLen(DLLDir$ & DLLName$), "0")

   For Position% = 1 To Len(DTime$)
      If Mid$(DTime$, Position%, 1) = ":" Then Exit For
   Next Position%
   Rev$ = Left$(DTime$, Position% - 1) & "." & Mid$(DTime$, Position% + 1)

   
   'DLLInfo = DLLDir$ & "," & DDate$ & "," & DSize$ & "," & Rev$

   Select Case TypeOfInfo%
      Case 0
         DLLInfo = DLLDir$
      Case 1
         DLLInfo = DDate$
      Case 2
         DLLInfo = DSize$
      Case 3
         DLLInfo = Rev$
      Case Else
         DLLInfo = Rev$
   End Select

   Exit Sub

   
LookElsewhere:
   If (Err = 53) Or (Err = 76) Then
      If Err = 76 Then MsgBox "Invalid path: " & DLLDir$
      Select Case attempt%
         Case 1
            DLLDir$ = WinDirec$
            attempt% = attempt% + 1
         Case 2
            DLLDir$ = SysDirec$
            attempt% = attempt% + 1
         Case 3
            If i% > Len(Path$) Then
               'MsgBox "Can't find " & DLLName$, , "File Not Found"
               BoardNum% = 0
               NoSuccess% = 1
               Resume Next
            End If
            If Mid$(Path$, Len(Path$), 1) = ";" Then Path$ = Left$(Path$, Len(Path$) - 1)
            Do While Mid$(Path$, i% + 1, 1) <> ";"
               If i% >= Len(Path$) Then Exit Do
               i% = i% + 1
            Loop
            DLLDir$ = Mid$(Path$, d% + 1, i% - d%)
            If Right$(DLLDir$, 1) <> "\" Then DLLDir$ = DLLDir$ & "\"
            i% = i% + 1
            d% = i%
      End Select
   Else
      MsgBox Error$(Err)
      If gnIDERunning Then
         Stop
      Else
         Dim Resp As VbMsgBoxResult
         Resp = MsgBox("This path is a Stop statement " & _
         "in the IDE. Check Local Error Handling options. " _
         & vbCrLf & vbCrLf & "          Click Yes to attempt " & _
         "to continue, No to exit application.", _
         vbYesNo, "Attempt To Continue?")
         If Resp = vbNo Then End
      End If
      Exit Sub
   End If
   Resume 0

   
End Sub

Private Sub GetOpSys(SysVer$)
   
   Dim rOsVersionInfo As OSVERSIONINFO
   Dim sOperatingSystem As String
  
  ' Pass the size of the structure into itself for the API call
  rOsVersionInfo.dwOSVersionInfoSize = Len(rOsVersionInfo)
  
  If GetVersionEx(rOsVersionInfo) Then
  
    Select Case rOsVersionInfo.dwPlatformId
    
      Case VER_PLATFORM_WIN32_NT
        
         If rOsVersionInfo.dwMajorVersion >= 6 Then
            If rOsVersionInfo.dwMinorVersion = 1 Then
               sOperatingSystem = "Windows 7"
               VistaPlatform$ = " 32 Bit"
            ElseIf rOsVersionInfo.dwMinorVersion = 2 Then
               'NameKey$ = "Software\Microsoft\Windows\CurrentVersion\Setup\ImageServicingData\PKeyConfigVersion"
               NameKey$ = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion"
               UserDirec = GetSysInfo(REGINF + HKLM, NameKey$)
               If Left$(UserDirec, 3) = "6.3" Then
                  sOperatingSystem = "Windows 8.1"
               Else
                  sOperatingSystem = "Windows 8"
               End If
               If Left$(UserDirec, 3) = "6.4" Then sOperatingSystem = "Windows 10"
               VistaPlatform$ = " 32 Bit"
            Else
               sOperatingSystem = "Windows Vista"
               VistaPlatform$ = " 32 Bit"
            End If
            If Not Environ("ProgramFiles(x86)") = "" Then
               VistaPlatform$ = " 64 Bit"
               gnWow64 = True
            End If
         Else
            If rOsVersionInfo.dwMajorVersion = 5 Then
               If rOsVersionInfo.dwMinorVersion = 0 Then
                  sOperatingSystem = "Windows 2000"
               Else
                  sOperatingSystem = "Windows XP"
               End If
            Else
               sOperatingSystem = "Windows NT"
            End If
         End If
        
      Case VER_PLATFORM_WIN32_WINDOWS
        If rOsVersionInfo.dwMajorVersion >= 5 Then
           sOperatingSystem = "Windows ME"
        ElseIf rOsVersionInfo.dwMajorVersion = 4 And rOsVersionInfo.dwMinorVersion > 0 Then
           sOperatingSystem = "Windows 98"
        Else
           sOperatingSystem = "Windows 95"
        End If
        
      Case VER_PLATFORM_WIN32s
        sOperatingSystem = "Win32s"
        
    End Select
  End If
  gsOS = sOperatingSystem
  Pos& = InStr(1, rOsVersionInfo.szCSDVersion, Chr(0))
  AddInfo$ = Left$(rOsVersionInfo.szCSDVersion, Pos& - 1)
  SysVer$ = sOperatingSystem & VistaPlatform$ & " " & AddInfo$
Exit Sub
   v& = GetVersion()
   
   dos& = (v& / &H10000) And &HFFFF&
   mvd% = dos& And &HFF
   pvd% = (dos& And &HFF00&) / &H100
   
   w& = (v& And &HFFFF&)
   pvw% = w& And &HFF
   mvw% = (w& And &HFF00&) / &H100

   If WIN32APP Then
      If (dos& And &H8000) Then 'not Windows NT
         If pvw% < 4 Then
            SysVer$ = "Win 32S "
            'remaining bit in high-order word
            'specify Win32S build number
         Else
            SysVer$ = "Win 95 " '(32 bit app)
         End If
      Else
         SysVer$ = "Win NT "
         'remaining bit in high-order word
         'specify NT build number
      End If
   Else
      If mvw% > 11 Then
         If mvd% > 0 Then
            SysVer$ = "Win 95B "
         Else
            SysVer$ = "Win 95 "
         End If
      Else
         SysVer$ = "Win 3.x "
      End If
   End If
   
End Sub

Function GetSysInfo(TypeOfInfo%, ByVal ReturnVal As Variant) As Variant

   Select Case (TypeOfInfo% And &HF00) / &H100
      Case 1
         GetOpSys os$
         GetSysInfo = os$
      Case 2
         GetRegInfo TypeOfInfo% And &HF, ReturnVal
         GetSysInfo = ReturnVal
      Case 4
         GetDLLInfo TypeOfInfo% And &HF, ReturnVal
         GetSysInfo = ReturnVal
      Case 5
         GetVxDInfo TypeOfInfo% And &HF, ReturnVal
         GetSysInfo = ReturnVal
      Case 6
         GetRegInfo TypeOfInfo% And &HF, ReturnVal
         GetSysInfo = ReturnVal
      Case Else
         GetSysInfo = "Invalid information request"
   End Select

End Function

Private Sub GetVxDInfo(TypeOfInfo%, VxDInfo As Variant)

   'returns VxD information on the VxD specified by
   'VxDInfo and information obtained from system.ini
   'in the form 'path$, date$, size$, rev$ (derived from time)'
   
   On Error GoTo VxDProblem
   Driver$ = LCase$(VxDInfo)
   If LCase$(msVxDName) <> Driver$ Then
      mnVxDFound = False
      mnVxDValid = False
      msVxDPath = ""
      msVxDName = Driver$
   End If
   
   If (Len(msWinDirec) = 0) Or (Len(msSysDirec) = 0) Then GetDirectory msWinDirec, msSysDirec
   If Len(msVxDPath) = 0 Then
      FName$ = msWinDirec & "SYSTEM.INI"
      Open FName$ For Input As #1
      Do
         Line Input #1, InString$
         If LCase$(Right$(InString$, Len(Driver$))) = Driver$ Then
            mnVxDFound = True
            mnVxDValid = True
            If (Left$(InString$, 1) = ";") Or (UCase$(Left$(InString$, 3)) = "REM") Then
               msVxDPath = Driver$ & " is commented out of system.ini..."
               mnVxDValid = False
            End If
            If mnVxDValid Then Exit Do
         End If
      Loop Until EOF(1)
      Close #1

      If Not mnVxDFound Then
         msVxDPath = "The device driver " & Driver$ & " was not found in system.ini."
         MsgBox InString$, , "Initialization Error"
      Else
         If mnVxDValid Then
            For Position% = 1 To Len(InString$)
               If Mid$(InString$, Position%, 1) = "=" Then Exit For
            Next Position%
            msVxDPath = Mid$(InString$, Position% + 1, Len(InString$) - Len(Driver$) - Position%)
         End If
      End If
   End If

   If mnVxDValid Then
      DrivePos% = InStr(msVxDPath, ":\")
      If DrivePos% = 0 Then msVxDPath = msWinDirec & msVxDPath
      VxDDate$ = FileDateTime(msVxDPath & Driver$)
      VSize$ = Format$(FileLen(msVxDPath & Driver$), "0")
      
      VDate$ = Format$(VxDDate$, "Short Date")
      VTime$ = Format$(VxDDate$, "Short Time")
      
      For Position% = 1 To Len(VTime$)
         If Mid$(VTime$, Position%, 1) = ":" Then Exit For
      Next Position%
      Rev$ = Left$(VTime$, Position% - 1) & "." & Mid$(VTime$, Position% + 1)
   End If
   
   Select Case TypeOfInfo%
      Case 0
         VxDInfo = msVxDPath
      Case 1
         VxDInfo = VDate$
      Case 2
         VxDInfo = VSize$
      Case 3
         VxDInfo = Rev$
      Case Else
         VxDInfo = Rev$
   End Select
   Exit Sub
   
VxDProblem:
   Select Case Err
      Case 53
         MsgBox "File not found: " & Driver$, , "Error Locating Driver"
      Case 76
         MsgBox "Invalid path: " & VDir$, , "Error Locating Driver"
      Case Else
         MsgBox Error$(Err), , "Error Locating Driver"
   End Select
   VxDInfo = "Not Found"
   Exit Sub

End Sub


Private Sub GetRegInfo(TypeOfInfo%, RegKey As Variant)

   Dim SubKey() As String
'   LF$ = Chr$(13) & Chr$(10)
   t$ = Chr$(9)
   Select Case TypeOfInfo%
      Case 0
         hKey& = HKEY_CLASSES_ROOT
      Case 1
         hKey& = HKEY_CURRENT_USER
      Case 2
         hKey& = HKEY_LOCAL_MACHINE
      Case 3
         hKey& = HKEY_USERS
   End Select
   For i% = Len(RegKey) To 1 Step -1
      If Mid$(RegKey, i%, 1) = "\" Then Exit For
   Next i%
   TopKey$ = Left$(RegKey, i%)
   If (i% + 1) < Len(RegKey) Then msValueName = Mid$(RegKey, i% + 1) & Chr$(0)
   If RegOpenKey(hKey&, TopKey$, phkResult&) = ERROR_SUCCESS Then
      'if no ValueName is specified then search subkeys
      If Len(msValueName) = 0 Then
         For i% = Len(RegKey) - 1 To 1 Step -1
            If Mid$(RegKey, i%, 1) = "\" Then Exit For
         Next i%
         EndKey$ = Mid$(RegKey, i% + 1, Len(RegKey) - 1)
         Select Case LCase$(EndKey$)
            Case "software"
               SubKeyDepth% = 2
               ValueDepth% = 1
            Case "pcmcia"
               SubKeyDepth% = 1
               ValueDepth% = 1
            Case "das component"
               SubKeyDepth% = 0
               ValueDepth% = 0
            Case "pci"
               SubKeyDepth% = 1
               ValueDepth% = 1
         End Select
         If EnumerateSubKeys(phkResult&, SubKeys$) Then
            result& = RegCloseKey(hKey&)
            If UCase$(Left$(SubKeys$, 4)) = "S-1-" Then
               'list of user accounts
               RegKey = SubKeys$
               Exit Sub
            Else
               While ParseSubKeys(phkResult&, SubKeys$, ThisKey$)
                  SaveKey% = False
                  Select Case True
                     Case LCase$(Left$(ThisKey$, 14)) = "computerboards"
                        SaveKey% = True
                     Case LCase$(ThisKey$) = "DAS Component"
                        SaveKey% = True
                     Case LCase$(Left$(ThisKey$, 8)) = "ven_1307"
                        SaveKey% = True
                  End Select
                  If SaveKey% Then
                     ReDim Preserve SubKey(NumSubs%) As String
                     SubKey(NumSubs%) = ThisKey$
                     NumSubs% = NumSubs% + 1
                  End If
               Wend
               For i% = 0 To NumSubs% - 1
                  If RegOpenKey(hKey&, TopKey$ & SubKey(i%), phkResult&) = ERROR_SUCCESS Then
                     If EnumerateSubKeys(phkResult&, SubKeys$) Then
                        While ParseSubKeys(phkResult&, SubKeys$, ThisKey$)
                        Wend
                     End If
                     result& = RegCloseKey(hKey&)
                  End If
               Next i%
            End If
         End If
      End If
      szData$ = Space$(255)
      lpcbData& = Len(szData$)
      If RegQueryValueEx(phkResult&, msValueName, 0, 1, szData$, lpcbData&) = ERROR_SUCCESS Then
         KeyVal$ = Left$(szData$, lpcbData& - 1)
         If Right$(KeyVal$, 10) = "CBUL32.SYS" Then
            NTDriverInfo% = True
            mnVxDValid = True
            mnVxDFound = True
            DrivePos% = InStr(KeyVal$, ":\")
            filespec$ = KeyVal$
            If DrivePos% > 0 Then filespec$ = Mid$(KeyVal$, InStr(KeyVal$, ":\") - 1)
            NameStart& = InStr(UCase(filespec$), "CBUL")
            msVxDName = Mid$(filespec$, NameStart&)
            msVxDPath = Left$(filespec$, NameStart& - 1)
         End If
         If (KeyVal$ = "DAS Component") Or (KeyVal$ = "") Then
            If LCase$(Right$(TopKey$, 10)) = "app paths\" Then
               KeyVal$ = "Application Paths" & vbCrLf
               PathList% = True
            ElseIf LCase$(Right$(TopKey$, 9)) = "enum\pci\" Then
               KeyVal$ = "Enumerated PCI Boards" & vbCrLf
               PCIList% = True
            ElseIf LCase$(Right$(TopKey$, 9)) = "software\" Then
               KeyVal$ = "Registered software" & vbCrLf
               SWList% = True
            Else
               GetDetail% = True
               KeyVal$ = KeyVal$ & vbCrLf
            End If
            If EnumerateSubKeys(phkResult&, SubKeys$) Then
               result& = RegCloseKey(hKey&)
               While ParseSubKeys(phkResult&, SubKeys$, ThisKey$)
                  
                  If PathList% Then
                     GetDetail% = False
                     If LCase$(Left$(ThisKey$, Len(ThisKey$) - 1)) = "daswiz.dll" Then GetDetail% = True
                     If LCase$(Left$(ThisKey$, Len(ThisKey$) - 1)) = "inscal32.exe" Then GetDetail% = True
                  End If
                  If SWList% Then
                     If LCase$(Left$(ThisKey$, 17)) = "universal library" Then GetDetail% = True
                     If LCase$(Left$(ThisKey$, 14)) = "computerboards" Then GetDetail% = True
                  End If
                  If PCIList% Then
                     GetDetail% = False
                     If LCase$(Left$(ThisKey$, 8)) = "ven_1307" Then GetDetail% = True
                     If EnumerateSubKeys(phkResult&, SubKeys2$) Then
                        result& = RegCloseKey(hKey&)
                        LocInString2% = True
                        Do
                           PrevLoc2& = CurrentLoc2& + 1
                           If PrevLoc2& >= Len(SubKeys2$) Then
                              LocInString2% = False
                              Exit Do
                           End If
                           CurrentLoc2& = InStr(PrevLoc2&, SubKeys2$, Chr$(0))
                           ThisKey2$ = Mid$(SubKeys2$, PrevLoc2&, CurrentLoc2& - (PrevLoc2& - 1))
                           If RegOpenKey(hKey&, TopKey$ & ThisKey$ & ThisKey2$, phkResult&) = ERROR_SUCCESS Then
                           End If
                        Loop While LocInString2%
                     End If
                  End If
                  If GetDetail% Then
                     If RegOpenKey(hKey&, TopKey$ & ThisKey$, phkResult&) = ERROR_SUCCESS Then
                        KeyVal$ = KeyVal$ & t$ & Left$(ThisKey$, Len(ThisKey$) - 1) & vbCrLf
                        KeyValue$ = ""
                        If EnumerateKeyVals(phkResult&, KeyValue$) Then
                           ParseKeyVals phkResult&, KeyValue$, KeyVal$
                        End If
                     End If
                  End If
               Wend
            End If
         End If
      Else
         NoInfo% = True
      End If
      result& = RegCloseKey(phkResult&)
   Else
      NoInfo% = True
   End If
   If NoInfo% Then
      RegKey = "No registry entry"
   Else
      If NTDriverInfo% Then
         RegKey = msVxDPath
      Else
         RegKey = KeyVal$
      End If
   End If
   
End Sub

Private Function EnumerateSubKeys(KeyVal As Long, AllSubEnum As String) As Integer

   Dim FT As FILETIME
   Index& = 0
   AllSubEnum = ""
   While lResult = ERROR_SUCCESS
      szBuffer$ = Space(255)
      lBuffSize& = Len(szBuffer$)
      szBuffer2$ = Space(255)
      lBuffSize2& = Len(szBuffer2$)
      lResult = RegEnumKeyEx(KeyVal, Index&, szBuffer$, lBuffSize&, _
                             0, szBuffer2$, lBuffSize2&, FT)
      If lResult = ERROR_SUCCESS Then
         EnumerateSubKeys = True
         SubEnum$ = Left$(szBuffer$, lBuffSize& + 1)
         AllSubEnum = AllSubEnum + SubEnum$
      End If
      Index& = Index& + 1
   Wend

End Function

Private Function EnumerateKeyVals(KeyHandle As Long, KeyVal As String) As Integer

   While lValResult& = ERROR_SUCCESS
      lpmsValueName = Space$(255)
      lpcbValueName& = Len(lpmsValueName)
      lpData$ = Space$(255)
      lpcbData& = Len(lpData$)
      lValResult& = RegEnumValue(KeyHandle, dwIndex&, lpmsValueName, _
      lpcbValueName&, lpReserved&, lpType&, lpData$, lpcbData&)
      If lValResult& = 0 Then
         EnumerateKeyVals = True
         KeyVal = KeyVal & Left$(lpmsValueName, lpcbValueName& + 1)
      End If
      dwIndex& = dwIndex& + 1
   Wend

End Function

Private Sub ParseKeyVals(phkResult As Long, KeyValue$, KeyVal As String)

   LocInValString% = True
   CurrentValLoc& = 0
   'vbCrLf = Chr$(13) & Chr$(10)
   t$ = Chr$(9)
   Do
      PrevValLoc& = CurrentValLoc& + 1
      If PrevValLoc& > Len(KeyValue$) Then
         LocInValString% = False
         Exit Do
      End If
      CurrentValLoc& = InStr(PrevValLoc&, KeyValue$, Chr$(0))
      ThisVal$ = Mid$(KeyValue$, PrevValLoc&, CurrentValLoc& - (PrevValLoc& - 1))
      szData$ = Space$(255)
      lpcbData& = Len(szData$)
      If RegQueryValueEx(phkResult, ThisVal$, 0, 1, szData$, lpcbData&) = ERROR_SUCCESS Then
         ItemVal$ = Left$(szData$, lpcbData& - 1)
         If Len(ThisVal$) = 1 Then ThisVal$ = "(Default) "
         KeyVal = KeyVal & t$ & t$ & Left$(ThisVal$, Len(ThisVal$) - 1) & ":   " & ItemVal$ & vbCrLf
      End If
   Loop While LocInValString%

End Sub

Public Function ParseSubKeys(phkResult As Long, SubKeys As String, ThisKey As String) As Integer

   Static CurrentLoc&
   ParseSubKeys = True
   PrevLoc& = CurrentLoc& + 1
   If PrevLoc& >= Len(SubKeys$) Then
      ParseSubKeys = False
      CurrentLoc& = 0
      Exit Function
   End If
   CurrentLoc& = InStr(PrevLoc&, SubKeys$, Chr$(0))
   ThisKey = Mid$(SubKeys$, PrevLoc&, CurrentLoc& - (PrevLoc& - 1))

End Function

Public Function AllocateMemory(BufferSize As Long) As Long

   Dim lpSystemInfo As SYSTEM_INFO
   Dim lpFileMappingAttributes As SECURITY_ATTRIBUTES
   
   Size& = BufferSize * 2
   
   hFile& = &HFFFF
   flProtect& = PAGE_READWRITE
   dwMaximumSizeHigh& = 0
   dwMaximumSizeLow& = Size&
   lpName$ = ""
   lpFileMappingAttributes.bInheritHandle = False
   lpFileMappingAttributes.lpSecurityDescriptor = 0
   lpFileMappingAttributes.nLength = 12
   
   hFileMappingObject& = CreateFileMapping(hFile&, lpFileMappingAttributes, _
      flProtect&, dwMaximumSizeHigh&, dwMaximumSizeLow&, lpName$)
   
   t& = Err.LastDllError
   If t& > 0 Then
      ULStat = NOWINDOWSMEMORY
      MessageType& = FORMAT_MESSAGE_FROM_SYSTEM
      Message$ = Space(128)
      Length& = FormatMessage(MessageType, 0, t&, 0, Message$, 128, 0)
      If Length& > 0 Then FMessage$ = Left(Message$, Length&)
      MsgBox FMessage$, vbCritical, "Memory Allocation Call Failed"
      AllocateMemory = 0
   Else
      mhMapFileHandle = hFileMappingObject&
   End If
   
   If Not mhMapFileHandle = 0 Then
      dwDesiredAccess& = FILE_MAP_ALL_ACCESS 'FILE_MAP_WRITE Or FILE_MAP_READ
      dwFileOffsetHigh& = 0
      dwFileOffsetLow& = 0
      dwNumberOfBytesToMap& = Size&
      MemRef& = MapViewOfFile(hFileMappingObject&, dwDesiredAccess&, _
      dwFileOffsetHigh&, dwFileOffsetLow&, dwNumberOfBytesToMap&)
      If MemRef& = 0 Then
         t& = Err.LastDllError
         MessageType& = FORMAT_MESSAGE_FROM_SYSTEM
         Message$ = Space(128)
         Length& = FormatMessage(MessageType, 0, t&, 0, Message$, 128, 0)
         If Length& > 0 Then FMessage$ = Left(Message$, Length&)
         MsgBox FMessage$, vbCritical, "Memory Call Failed"
         ULStat = NOWINDOWSMEMORY
         AllocateMemory = MemRef&
         Exit Function
      End If
   End If
   AllocateMemory = MemRef&

End Function

Public Function FreeMemory(MemHandle As Variant) As Integer

   Handle& = CLng(MemHandle)
   MemFreeResult& = UnmapViewOfFile(Handle&)
   If MemFreeResult& = 0 Then
      t& = Err.LastDllError   'GetLastError()
      MessageType& = FORMAT_MESSAGE_FROM_SYSTEM
      Message$ = Space(128)
      Length& = FormatMessage(MessageType, 0, t&, 0, Message$, 128, 0)
      If Length& > 0 Then FMessage$ = Left(Message$, Length&)
      'to do - fix this (never frees memory)
      'MsgBox FMessage$, vbCritical, "Free Memory Call Failed"
      ULStat = NOWINDOWSMEMORY
   End If
   If Not (mhMapFileHandle = 0) Then
      CloseObj& = CloseHandle(mhMapFileHandle)
      If (CloseObj& = 0) Then
         t& = Err.LastDllError
         ULStat = NOWINDOWSMEMORY
         MessageType& = FORMAT_MESSAGE_FROM_SYSTEM
         Message$ = Space(128)
         Length& = FormatMessage(MessageType, 0, t&, 0, Message$, 128, 0)
         If Length& > 0 Then FMessage$ = Left(Message$, Length&)
         MsgBox FMessage$, vbCritical, "Close Handle Call Failed"
      Else
         mhMapFileHandle = 0
      End If
   End If
   'to do - fix this
   FreeMemory = True '(MemFreeResult& <> False)
   
End Function

Public Function GetMapHandle() As Long

   lpName$ = "MyFile"
   dwDesiredAccess& = FILE_MAP_WRITE Or FILE_MAP_READ
   hFileMappingObject& = OpenFileMapping(dwDesiredAccess&, bInheritHandle&, lpName$)
   If hFileMappingObject& = 0 Then ULStat = NOWINDOWSMEMORY
   'If SaveFunc(mfNoForm, WinAPIOpenFileMapping, ULStat, hFileMappingObject&, dwDesiredAccess&, dwFileOffsetHigh&, dwFileOffsetLow&, dwNumberOfBytesToMap&, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
   
   dwFileOffsetHigh& = 0
   dwFileOffsetLow& = 0
   dwNumberOfBytesToMap& = 0  'maps entire buffer
   MemRef& = MapViewOfFile(hFileMappingObject&, dwDesiredAccess&, _
      dwFileOffsetHigh&, dwFileOffsetLow&, dwNumberOfBytesToMap&)
'   CloseObj& = CloseHandle(hFileMappingObject&)
'   ULStat = cbWinBufToArray(MemRef&, DataBuffer%(0), FirstPoint&, CBCount&)
'   CloseView& = UnmapViewOfFile(MemRef&)
'   T& = GetLastError()
   GetMapHandle = MemRef&

End Function

Sub GetDriverPath(ByVal KeyName As String, ByRef PathName As String)

   Dim SubKey() As String
   hKey& = HKEY_LOCAL_MACHINE
   
   For i% = Len(KeyName) To 1 Step -1
      If Mid$(KeyName, i%, 1) = "\" Then Exit For
   Next i%
   TopKey$ = Left$(KeyName, i%)
   If (i% + 1) < Len(RegKey) Then msValueName = Mid$(RegKey, i% + 1) & Chr$(0)
   If RegOpenKey(hKey&, KeyName, phkResult&) = ERROR_SUCCESS Then
      szData$ = Space$(255)
      lpcbData& = Len(szData$)
      msValueName = "ImagePath"
      If RegQueryValueEx(phkResult&, msValueName, 0, 1, szData$, lpcbData&) = ERROR_SUCCESS Then
         KeyVal$ = Left$(szData$, lpcbData& - 1)
         PathName = KeyVal$
      Else
         PathName = " no path exists."
      End If
   Else
      PathName = ""
   End If

End Sub

Function SearchKeyList(KeyPath As String, KeyList As Variant) As Integer

   Dim NewList() As String
   j% = -1
   NumKeys% = UBound(KeyList)
   For i% = 0 To NumKeys%
      hKey& = HKEY_LOCAL_MACHINE
      If RegOpenKey(hKey&, KeyPath & KeyList(i%), phkResult&) = ERROR_SUCCESS Then
         'SubKeys& = GetSubKeys(phkResult&)
         j% = j% + 1
         ReDim Preserve NewList(j%)
         NewList(j%) = KeyList(i%)
         result& = RegCloseKey(phkResult&)
      End If
   Next
   KeyList = NewList()
   SearchKeyList = j%
   
End Function

Public Sub GetKeys(KeyList As Variant)

   KeyList = msKeys()
   
End Sub

Public Function GetFileDir(filespec As String, SearchType As VbFileAttribute, result As Variant) As Long

      'ProfilesDir$ = Mid$(StartDir$, 1, Pos&)
      Dim ResultArray() As String
      d$ = Dir(filespec, SearchType)
      NumFound& = -1
      If Not (d$ = "") Then
         If Not (d$ = "." Or d$ = "..") Then
            NumFound& = NumFound& + 1
            ReDim ResultArray(NumFound&)
            ResultArray(NumFound&) = d$
         End If
         Do
            d$ = Dir()
            If Not (d$ = "" Or d$ = "." Or d$ = "..") Then
               NumFound& = NumFound& + 1
               ReDim Preserve ResultArray(NumFound&)
               ResultArray(NumFound&) = d$
            End If
         Loop While Not (d$ = "")
         result = ResultArray
      End If
      GetFileDir = NumFound&

End Function

Public Function GetSubKeys(hKey As Long, SubKeys As Variant) As Long

   Dim FT As FILETIME
   Index& = 0
   Dim SKeys() As String
   Dim hSKeys() As Long
   
   While lResult = ERROR_SUCCESS
      szBuffer$ = Space(255)
      lBuffSize& = Len(szBuffer$)
      szBuffer2$ = Space(255)
      lBuffSize2& = Len(szBuffer2$)
      lResult = RegEnumKeyEx(hKey, Index&, szBuffer$, lBuffSize&, _
                             0, szBuffer2$, lBuffSize2&, FT)
      If lResult = ERROR_SUCCESS Then
         EnSubKeys% = True
         SubEnum$ = Left$(szBuffer$, lBuffSize& + 1)
         ReDim Preserve SKeys(Index&)
         'ReDim Preserve hSKeys(Index&)
         SKeys(Index&) = SubEnum$
         'hSKeys(Index&) =
         Index& = Index& + 1
      End If
   Wend
   'RegCloseKey (hKey)
   SubKeys = SKeys()
   GetSubKeys = Index& - 1

End Function

Public Function GetRegGroup(hKey As Long, KeyName As String, result As Long) As Integer

   If RegOpenKey(hKey&, KeyName, phkResult&) = ERROR_SUCCESS Then
      GetRegGroup = True
      result = phkResult&
      'RegCloseKey (hKey)
   End If

End Function

Public Function GetKeyValue(phkResult As Long, ValueName As String, KeyVal As String) As Integer

   szData$ = Space$(255)
   lpcbData& = Len(szData$)
   ValueKeyName$ = ValueName & Chr$(0)
   If RegQueryValueEx(phkResult, ValueKeyName$, 0, 1, szData$, lpcbData&) = ERROR_SUCCESS Then
      If lpcbData& > 0 Then KeyVal = Left$(szData$, lpcbData& - 1)
      GetKeyValue = True
   End If

End Function
