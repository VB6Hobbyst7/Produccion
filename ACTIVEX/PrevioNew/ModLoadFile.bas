Attribute VB_Name = "ModLoadFile"
Option Explicit

Public User_name As String
Const REG_DWORD = 4
Const HKEY_DYN_DATA = &H80000006

Private Declare Function GetWindow Lib "user32" _
(ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetParent Lib "user32" _
(ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib _
"user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" _
Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal _
lpString As String, ByVal cch As Long) As Long

Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Type MEMORYSTATUS
    dwLength        As Long ' sizeof(MEMORYSTATUS)
    dwMemoryLoad    As Long ' percent of memory in use
    dwTotalPhys     As Long ' bytes of physical memory
    dwAvailPhys     As Long ' free physical memory bytes
    dwTotalPageFile As Long ' bytes of paging file
    dwAvailPageFile As Long ' free bytes of paging file
    dwTotalVirtual  As Long ' user bytes of address space
    dwAvailVirtual  As Long ' free user bytes
End Type

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const MAX_PATH As Integer = 260

Private Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szExeFile As String * MAX_PATH
End Type

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" _
Alias "CreateToolhelp32Snapshot" _
(ByVal lFlags As Long, ByVal lProcessID As Long) As Long

Private Declare Function ProcessFirst Lib "kernel32" _
Alias "Process32First" _
(ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Private Declare Function ProcessNext Lib "kernel32" _
Alias "Process32Next" _
(ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Private Declare Sub CloseHandle Lib "kernel32" _
(ByVal hPass As Long)
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type
Private Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long
Const shrdNoMRUString = &H2    '2nd bit is set
Private Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
'By Capi 23102008 para controlar la impresora elegida
Public gsImpresoraElegida As String
'


Public Function VerArchivoCargado(ByVal lsArchivo As String) As Boolean
Dim nstr As String
Dim hSnapShot As Long
Dim uProcess As PROCESSENTRY32
Dim r As Long
Dim lsNomEXE As String
Dim lnPos As Integer

lnPos = InStrRev(lsArchivo, "\", , vbTextCompare)

lsNomEXE = UCase(Mid(lsArchivo, lnPos + 1, Len(lsArchivo)))

hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)

If hSnapShot = 0 Then
  Exit Function
End If

VerArchivoCargado = False
uProcess.dwSize = Len(uProcess)

r = ProcessFirst(hSnapShot, uProcess)
Do While r
    nstr = uProcess.szExeFile
    If InStr(1, UCase(Trim(nstr)), lsNomEXE, vbTextCompare) Then
        VerArchivoCargado = True
    End If
    r = ProcessNext(hSnapShot, uProcess)
Loop 'Do all the processs

End Function
