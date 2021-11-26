Attribute VB_Name = "gAcceso"
Option Explicit

Private Const HEAP_ZERO_MEMORY = &H8

Private Const SEC_WINNT_AUTH_IDENTITY_ANSI = &H1

Private Const SECBUFFER_TOKEN = &H2

Private Const SECURITY_NATIVE_DREP = &H10

Private Const SECPKG_CRED_INBOUND = &H1
Private Const SECPKG_CRED_OUTBOUND = &H2

Private Const SEC_I_CONTINUE_NEEDED = &H90312
Private Const SEC_I_COMPLETE_NEEDED = &H90313
Private Const SEC_I_COMPLETE_AND_CONTINUE = &H90314

Private Const VER_PLATFORM_WIN32_NT = &H2

Type SecPkgInfo
fCapabilities As Long
wVersion As Integer
wRPCID As Integer
cbMaxToken As Long
Name As Long
Comment As Long
End Type

Type SecHandle
dwLower As Long
dwUpper As Long
End Type

Type AUTH_SEQ
fInitialized As Boolean
fHaveCredHandle As Boolean
fHaveCtxtHandle As Boolean
hcred As SecHandle
hctxt As SecHandle
End Type

Type SEC_WINNT_AUTH_IDENTITY
User As String
UserLength As Long
Domain As String
DomainLength As Long
Password As String
PasswordLength As Long
Flags As Long
End Type

Type TimeStamp
LowPart As Long
HighPart As Long
End Type

Type SecBuffer
cbBuffer As Long
BufferType As Long
pvBuffer As Long
End Type

Type SecBufferDesc
ulVersion As Long
cBuffers As Long
pBuffers As Long
End Type

Private Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function NT4QuerySecurityPackageInfo Lib "security" _
Alias "QuerySecurityPackageInfoA" (ByVal PackageName As String, _
ByRef pPackageInfo As Long) As Long

Private Declare Function QuerySecurityPackageInfo Lib "secur32" _
Alias "QuerySecurityPackageInfoA" (ByVal PackageName As String, _
ByRef pPackageInfo As Long) As Long

Private Declare Function NT4FreeContextBuffer Lib "security" _
Alias "FreeContextBuffer" (ByVal pvContextBuffer As Long) As Long

Private Declare Function FreeContextBuffer Lib "secur32" _
(ByVal pvContextBuffer As Long) As Long

Private Declare Function NT4InitializeSecurityContext Lib "security" _
Alias "InitializeSecurityContextA" _
(ByRef phCredential As SecHandle, ByRef phContext As SecHandle, _
ByVal pszTargetName As Long, ByVal fContextReq As Long, _
ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
ByRef pInput As SecBufferDesc, ByVal Reserved2 As Long, _
ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function InitializeSecurityContext Lib "secur32" _
Alias "InitializeSecurityContextA" _
(ByRef phCredential As SecHandle, ByRef phContext As SecHandle, _
ByVal pszTargetName As Long, ByVal fContextReq As Long, _
ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
ByRef pInput As SecBufferDesc, ByVal Reserved2 As Long, _
ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function NT4InitializeSecurityContext2 Lib "security" _
Alias "InitializeSecurityContextA" _
(ByRef phCredential As SecHandle, ByVal phContext As Long, _
ByVal pszTargetName As Long, ByVal fContextReq As Long, _
ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
ByVal pInput As Long, ByVal Reserved2 As Long, _
ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function InitializeSecurityContext2 Lib "secur32" _
Alias "InitializeSecurityContextA" _
(ByRef phCredential As SecHandle, ByVal phContext As Long, _
ByVal pszTargetName As Long, ByVal fContextReq As Long, _
ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
ByVal pInput As Long, ByVal Reserved2 As Long, _
ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function NT4AcquireCredentialsHandle Lib "security" _
Alias "AcquireCredentialsHandleA" (ByVal pszPrincipal As Long, _
ByVal pszPackage As String, ByVal fCredentialUse As Long, _
ByVal pvLogonId As Long, _
ByRef pAuthData As SEC_WINNT_AUTH_IDENTITY, _
ByVal pGetKeyFn As Long, ByVal pvGetKeyArgument As Long, _
ByRef phCredential As SecHandle, ByRef ptsExpiry As TimeStamp) _
As Long

Private Declare Function AcquireCredentialsHandle Lib "secur32" _
Alias "AcquireCredentialsHandleA" (ByVal pszPrincipal As Long, _
ByVal pszPackage As String, ByVal fCredentialUse As Long, _
ByVal pvLogonId As Long, _
ByRef pAuthData As SEC_WINNT_AUTH_IDENTITY, _
ByVal pGetKeyFn As Long, ByVal pvGetKeyArgument As Long, _
ByRef phCredential As SecHandle, ByRef ptsExpiry As TimeStamp) _
As Long

Private Declare Function NT4AcquireCredentialsHandle2 Lib "security" _
Alias "AcquireCredentialsHandleA" (ByVal pszPrincipal As Long, _
ByVal pszPackage As String, ByVal fCredentialUse As Long, _
ByVal pvLogonId As Long, ByVal pAuthData As Long, _
ByVal pGetKeyFn As Long, ByVal pvGetKeyArgument As Long, _
ByRef phCredential As SecHandle, ByRef ptsExpiry As TimeStamp) _
As Long

Private Declare Function AcquireCredentialsHandle2 Lib "secur32" _
Alias "AcquireCredentialsHandleA" (ByVal pszPrincipal As Long, _
ByVal pszPackage As String, ByVal fCredentialUse As Long, _
ByVal pvLogonId As Long, ByVal pAuthData As Long, _
ByVal pGetKeyFn As Long, ByVal pvGetKeyArgument As Long, _
ByRef phCredential As SecHandle, ByRef ptsExpiry As TimeStamp) _
As Long

Private Declare Function NT4AcceptSecurityContext Lib "security" _
Alias "AcceptSecurityContext" (ByRef phCredential As SecHandle, _
ByRef phContext As SecHandle, ByRef pInput As SecBufferDesc, _
ByVal fContextReq As Long, ByVal TargetDataRep As Long, _
ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function AcceptSecurityContext Lib "secur32" _
(ByRef phCredential As SecHandle, _
ByRef phContext As SecHandle, ByRef pInput As SecBufferDesc, _
ByVal fContextReq As Long, ByVal TargetDataRep As Long, _
ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function NT4AcceptSecurityContext2 Lib "security" _
Alias "AcceptSecurityContext" (ByRef phCredential As SecHandle, _
ByVal phContext As Long, ByRef pInput As SecBufferDesc, _
ByVal fContextReq As Long, ByVal TargetDataRep As Long, _
ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function AcceptSecurityContext2 Lib "secur32" _
Alias "AcceptSecurityContext" (ByRef phCredential As SecHandle, _
ByVal phContext As Long, ByRef pInput As SecBufferDesc, _
ByVal fContextReq As Long, ByVal TargetDataRep As Long, _
ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function NT4CompleteAuthToken Lib "security" _
Alias "CompleteAuthToken" (ByRef phContext As SecHandle, _
ByRef pToken As SecBufferDesc) As Long

Private Declare Function CompleteAuthToken Lib "secur32" _
(ByRef phContext As SecHandle, _
ByRef pToken As SecBufferDesc) As Long

Private Declare Function NT4DeleteSecurityContext Lib "security" _
Alias "DeleteSecurityContext" (ByRef phContext As SecHandle) _
As Long

Private Declare Function DeleteSecurityContext Lib "secur32" _
(ByRef phContext As SecHandle) _
As Long

Private Declare Function NT4FreeCredentialsHandle Lib "security" _
Alias "FreeCredentialsHandle" (ByRef phContext As SecHandle) _
As Long

Private Declare Function FreeCredentialsHandle Lib "secur32" _
(ByRef phContext As SecHandle) _
As Long

Private Declare Function GetProcessHeap Lib "kernel32" () As Long

Private Declare Function HeapAlloc Lib "kernel32" _
(ByVal hHeap As Long, ByVal dwFlags As Long, _
ByVal dwBytes As Long) As Long

Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, _
ByVal dwFlags As Long, ByVal lpMem As Long) As Long

Private Declare Function GetVersionExA Lib "kernel32" _
(lpVersionInformation As OSVERSIONINFO) As Integer

Dim g_NT4 As Boolean

Public Function GenClientContext(ByRef AuthSeq As AUTH_SEQ, _
ByRef AuthIdentity As SEC_WINNT_AUTH_IDENTITY, _
ByVal pIn As Long, ByVal cbIn As Long, _
ByVal pOut As Long, ByRef cbOut As Long, _
ByRef fDone As Boolean) As Boolean

Dim ss As Long
Dim tsExpiry As TimeStamp
Dim sbdOut As SecBufferDesc
Dim sbOut As SecBuffer
Dim sbdIn As SecBufferDesc
Dim sbIn As SecBuffer
Dim fContextAttr As Long

GenClientContext = False

If Not AuthSeq.fInitialized Then

If g_NT4 Then
ss = NT4AcquireCredentialsHandle(0&, "NTLM", _
SECPKG_CRED_OUTBOUND, 0&, AuthIdentity, 0&, 0&, _
AuthSeq.hcred, tsExpiry)
Else
ss = AcquireCredentialsHandle(0&, "NTLM", _
SECPKG_CRED_OUTBOUND, 0&, AuthIdentity, 0&, 0&, _
AuthSeq.hcred, tsExpiry)
End If

If ss < 0 Then
Exit Function
End If

AuthSeq.fHaveCredHandle = True

End If

' Prepare output buffer
sbdOut.ulVersion = 0
sbdOut.cBuffers = 1
sbdOut.pBuffers = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, _
Len(sbOut))

sbOut.cbBuffer = cbOut
sbOut.BufferType = SECBUFFER_TOKEN
sbOut.pvBuffer = pOut

CopyMemory ByVal sbdOut.pBuffers, sbOut, Len(sbOut)

' Prepare input buffer
If AuthSeq.fInitialized Then

sbdIn.ulVersion = 0
sbdIn.cBuffers = 1
sbdIn.pBuffers = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, _
Len(sbIn))

sbIn.cbBuffer = cbIn
sbIn.BufferType = SECBUFFER_TOKEN
sbIn.pvBuffer = pIn

CopyMemory ByVal sbdIn.pBuffers, sbIn, Len(sbIn)

End If

If AuthSeq.fInitialized Then

If g_NT4 Then
ss = NT4InitializeSecurityContext(AuthSeq.hcred, _
AuthSeq.hctxt, 0&, 0, 0, SECURITY_NATIVE_DREP, sbdIn, _
0, AuthSeq.hctxt, sbdOut, fContextAttr, tsExpiry)
Else
ss = InitializeSecurityContext(AuthSeq.hcred, _
AuthSeq.hctxt, 0&, 0, 0, SECURITY_NATIVE_DREP, sbdIn, _
0, AuthSeq.hctxt, sbdOut, fContextAttr, tsExpiry)
End If

Else

If g_NT4 Then
ss = NT4InitializeSecurityContext2(AuthSeq.hcred, 0&, 0&, _
0, 0, SECURITY_NATIVE_DREP, 0&, 0, AuthSeq.hctxt, _
sbdOut, fContextAttr, tsExpiry)
Else
ss = InitializeSecurityContext2(AuthSeq.hcred, 0&, 0&, _
0, 0, SECURITY_NATIVE_DREP, 0&, 0, AuthSeq.hctxt, _
sbdOut, fContextAttr, tsExpiry)
End If

End If

If ss < 0 Then
GoTo FreeResourcesAndExit
End If

AuthSeq.fHaveCtxtHandle = True

' If necessary, complete token
If ss = SEC_I_COMPLETE_NEEDED _
Or ss = SEC_I_COMPLETE_AND_CONTINUE Then

If g_NT4 Then
ss = NT4CompleteAuthToken(AuthSeq.hctxt, sbdOut)
Else
ss = CompleteAuthToken(AuthSeq.hctxt, sbdOut)
End If

If ss < 0 Then
GoTo FreeResourcesAndExit
End If

End If

CopyMemory sbOut, ByVal sbdOut.pBuffers, Len(sbOut)
cbOut = sbOut.cbBuffer

If Not AuthSeq.fInitialized Then
AuthSeq.fInitialized = True
End If

fDone = Not (ss = SEC_I_CONTINUE_NEEDED _
Or ss = SEC_I_COMPLETE_AND_CONTINUE)

GenClientContext = True

FreeResourcesAndExit:

If sbdOut.pBuffers <> 0 Then
HeapFree GetProcessHeap(), 0, sbdOut.pBuffers
End If

If sbdIn.pBuffers <> 0 Then
HeapFree GetProcessHeap(), 0, sbdIn.pBuffers
End If

End Function


Public Function GenServerContext(ByRef AuthSeq As AUTH_SEQ, _
ByVal pIn As Long, ByVal cbIn As Long, _
ByVal pOut As Long, ByRef cbOut As Long, _
ByRef fDone As Boolean) As Boolean

Dim ss As Long
Dim tsExpiry As TimeStamp
Dim sbdOut As SecBufferDesc
Dim sbOut As SecBuffer
Dim sbdIn As SecBufferDesc
Dim sbIn As SecBuffer
Dim fContextAttr As Long

GenServerContext = False

If Not AuthSeq.fInitialized Then

If g_NT4 Then
ss = NT4AcquireCredentialsHandle2(0&, "NTLM", _
SECPKG_CRED_INBOUND, 0&, 0&, 0&, 0&, AuthSeq.hcred, _
tsExpiry)
Else
ss = AcquireCredentialsHandle2(0&, "NTLM", _
SECPKG_CRED_INBOUND, 0&, 0&, 0&, 0&, AuthSeq.hcred, _
tsExpiry)
End If

If ss < 0 Then
Exit Function
End If

AuthSeq.fHaveCredHandle = True

End If

' Prepare output buffer
sbdOut.ulVersion = 0
sbdOut.cBuffers = 1
sbdOut.pBuffers = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, _
Len(sbOut))

sbOut.cbBuffer = cbOut
sbOut.BufferType = SECBUFFER_TOKEN
sbOut.pvBuffer = pOut

CopyMemory ByVal sbdOut.pBuffers, sbOut, Len(sbOut)

' Prepare input buffer
sbdIn.ulVersion = 0
sbdIn.cBuffers = 1
sbdIn.pBuffers = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, _
Len(sbIn))

sbIn.cbBuffer = cbIn
sbIn.BufferType = SECBUFFER_TOKEN
sbIn.pvBuffer = pIn

CopyMemory ByVal sbdIn.pBuffers, sbIn, Len(sbIn)

If AuthSeq.fInitialized Then

If g_NT4 Then
ss = NT4AcceptSecurityContext(AuthSeq.hcred, AuthSeq.hctxt, _
sbdIn, 0, SECURITY_NATIVE_DREP, AuthSeq.hctxt, sbdOut, _
fContextAttr, tsExpiry)
Else
ss = AcceptSecurityContext(AuthSeq.hcred, AuthSeq.hctxt, _
sbdIn, 0, SECURITY_NATIVE_DREP, AuthSeq.hctxt, sbdOut, _
fContextAttr, tsExpiry)
End If

Else

If g_NT4 Then
ss = NT4AcceptSecurityContext2(AuthSeq.hcred, 0&, sbdIn, 0, _
SECURITY_NATIVE_DREP, AuthSeq.hctxt, sbdOut, _
fContextAttr, tsExpiry)
Else
ss = AcceptSecurityContext2(AuthSeq.hcred, 0&, sbdIn, 0, _
SECURITY_NATIVE_DREP, AuthSeq.hctxt, sbdOut, _
fContextAttr, tsExpiry)
End If

End If

If ss < 0 Then
GoTo FreeResourcesAndExit
End If

AuthSeq.fHaveCtxtHandle = True

' If necessary, complete token
If ss = SEC_I_COMPLETE_NEEDED _
Or ss = SEC_I_COMPLETE_AND_CONTINUE Then

If g_NT4 Then
ss = NT4CompleteAuthToken(AuthSeq.hctxt, sbdOut)
Else
ss = CompleteAuthToken(AuthSeq.hctxt, sbdOut)
End If

If ss < 0 Then
GoTo FreeResourcesAndExit
End If

End If

CopyMemory sbOut, ByVal sbdOut.pBuffers, Len(sbOut)
cbOut = sbOut.cbBuffer

If Not AuthSeq.fInitialized Then
AuthSeq.fInitialized = True
End If

fDone = Not (ss = SEC_I_CONTINUE_NEEDED _
Or ss = SEC_I_COMPLETE_AND_CONTINUE)

GenServerContext = True

FreeResourcesAndExit:

If sbdOut.pBuffers <> 0 Then
HeapFree GetProcessHeap(), 0, sbdOut.pBuffers
End If

If sbdIn.pBuffers <> 0 Then
HeapFree GetProcessHeap(), 0, sbdIn.pBuffers
End If

End Function


Public Function SSPValidateUser(User As String, Domain As String, _
Password As String) As Boolean

Dim pSPI As Long
Dim SPI As SecPkgInfo
Dim cbMaxToken As Long

Dim pClientBuf As Long
Dim pServerBuf As Long

Dim ai As SEC_WINNT_AUTH_IDENTITY

Dim asClient As AUTH_SEQ
Dim asServer As AUTH_SEQ
Dim cbIn As Long
Dim cbOut As Long
Dim fDone As Boolean

Dim osinfo As OSVERSIONINFO

SSPValidateUser = False

' Determine if system is Windows NT (version 4.0 or earlier)
osinfo.dwOSVersionInfoSize = Len(osinfo)
osinfo.szCSDVersion = Space$(128)
GetVersionExA osinfo
g_NT4 = (osinfo.dwPlatformId = VER_PLATFORM_WIN32_NT And _
osinfo.dwMajorVersion <= 4)

' Get max token size
If g_NT4 Then
NT4QuerySecurityPackageInfo "NTLM", pSPI
Else
QuerySecurityPackageInfo "NTLM", pSPI
End If

CopyMemory SPI, ByVal pSPI, Len(SPI)
cbMaxToken = SPI.cbMaxToken

If g_NT4 Then
NT4FreeContextBuffer pSPI
Else
FreeContextBuffer pSPI
End If

' Allocate buffers for client and server messages
pClientBuf = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, _
cbMaxToken)
If pClientBuf = 0 Then
GoTo FreeResourcesAndExit
End If

pServerBuf = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, _
cbMaxToken)
If pServerBuf = 0 Then
GoTo FreeResourcesAndExit
End If

' Initialize auth identity structure
ai.Domain = Domain
ai.DomainLength = Len(Domain)
ai.User = User
ai.UserLength = Len(User)
ai.Password = Password
ai.PasswordLength = Len(Password)
ai.Flags = SEC_WINNT_AUTH_IDENTITY_ANSI

' Prepare client message (negotiate) .
cbOut = cbMaxToken
If Not GenClientContext(asClient, ai, 0, 0, pClientBuf, cbOut, _
fDone) Then
GoTo FreeResourcesAndExit
End If

' Prepare server message (challenge) .
cbIn = cbOut
cbOut = cbMaxToken
If Not GenServerContext(asServer, pClientBuf, cbIn, pServerBuf, _
cbOut, fDone) Then
' Most likely failure: AcceptServerContext fails with
' SEC_E_LOGON_DENIED in the case of bad szUser or szPassword.
' Unexpected Result: Logon will succeed if you pass in a bad
' szUser and the guest account is enabled in the specified domain.
GoTo FreeResourcesAndExit
End If

' Prepare client message (authenticate) .
cbIn = cbOut
cbOut = cbMaxToken
If Not GenClientContext(asClient, ai, pServerBuf, cbIn, pClientBuf, _
cbOut, fDone) Then
GoTo FreeResourcesAndExit
End If

' Prepare server message (authentication) .
cbIn = cbOut
cbOut = cbMaxToken
If Not GenServerContext(asServer, pClientBuf, cbIn, pServerBuf, _
cbOut, fDone) Then
GoTo FreeResourcesAndExit
End If

SSPValidateUser = True

FreeResourcesAndExit:

' Clean up resources
If asClient.fHaveCtxtHandle Then
If g_NT4 Then
NT4DeleteSecurityContext asClient.hctxt
Else
DeleteSecurityContext asClient.hctxt
End If
End If

If asClient.fHaveCredHandle Then
If g_NT4 Then
NT4FreeCredentialsHandle asClient.hcred
Else
FreeCredentialsHandle asClient.hcred
End If
End If

If asServer.fHaveCtxtHandle Then
If g_NT4 Then
NT4DeleteSecurityContext asServer.hctxt
Else
DeleteSecurityContext asServer.hctxt
End If
End If

If asServer.fHaveCredHandle Then
If g_NT4 Then
NT4FreeCredentialsHandle asServer.hcred
Else
FreeCredentialsHandle asServer.hcred
End If
End If

If pClientBuf <> 0 Then
HeapFree GetProcessHeap(), 0, pClientBuf
End If

If pServerBuf <> 0 Then
HeapFree GetProcessHeap(), 0, pServerBuf
End If

End Function

'Obtener Usuario Actual
Public Function GetUserName()

Dim InfoNT
Set InfoNT = CreateObject("WinNTSystemInfo")
GetUserName = InfoNT.UserName

End Function

'Obtener nombre del dominio
Public Function GetDomainName()

Dim InfoNT
Set InfoNT = CreateObject("WinNTSystemInfo")
GetDomainName = LCase(InfoNT.DomainName)

End Function
'Obtener nombre del ordenador
Public Function GetComputerName()

Dim InfoNT
Set InfoNT = CreateObject("WinNTSystemInfo")
GetComputerName = LCase(InfoNT.ComputerName)

End Function




