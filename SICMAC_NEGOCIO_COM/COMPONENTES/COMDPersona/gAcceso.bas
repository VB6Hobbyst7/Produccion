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
flags As Long
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
Public Vusuario As String
'Para Cargar Usuarios
Private Type JoinLong
   X As Long
   Dummy As Integer
End Type

Private Type JoinInt
   Bottom As Integer
   Top As Integer
   Dummy As Integer
End Type

Private Type NETRESOURCE
        dwScope As Long
        dwType As Long
        dwDisplayType As Long
        dwUsage As Long
        lpLocalName As Long
        lpRemoteName As Long
        lpComment As Long
        lpProvider As Long
End Type

Private Espaciado As Integer
Private Declare Function NetGetDCName Lib "netapi32.dll" (ByRef servername As Byte, ByRef DomainName As Byte, ByRef buffer As Long) As Long
Private Declare Function NetUserGetInfo Lib "NETAPI32" (ByRef servername As Byte, ByRef UserName As Byte, ByVal level As Long, ByRef buffer As Long) As Long
Private Declare Function NetQueryDisplayInformation Lib "netapi32.dll" (ByRef servername As Byte, ByVal level As Long, ByVal Index As Long, ByVal EntriesRequested As Long, ByVal PreferredMaximumLength As Long, ByRef ReturnedEntryCount As Long, ByRef SortedBuffer As Long) As Long
Private Declare Function PtrToInt Lib "kernel32" Alias "lstrcpynW" (RetVal As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long
Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
Private Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long
Private Declare Function NetAPIBufferFree Lib "netapi32.dll" Alias "NetApiBufferFree" (ByVal Ptr As Long) As Long
Private Users$(), typRootResourses() As NETRESOURCE, typDomainResourses() As NETRESOURCE

'Para Cambiar el Password
Private Declare Function NetUserChangePassword Lib "netapi32.dll" ( _
     ByVal DomainName As String, ByVal UserName As String, _
     ByVal OldPassword As String, ByVal NewPassword As String) As Long
     
'Para Obtener el Usuario Logeado
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Private sGrupoUsu() As String
Private nNumGrupos As Integer

Private LstUsuarios() As String
Private LstUsuariosNameFull() As String
Private LstUsuariosLoged() As String
Public nNumUsu As Integer

Public MenuItems As String
Public nNumMenus As Integer

Public sCadCon As String
Public sCadMenu As String
Public sCadMenuGrp As String
Public sCadMenuSql As String
Public vError As Boolean
Public sMsgError As String

Private vsTipoPermisoUsu As String
Private vsTipoPermisoGrp As String
Private nPosUsuGrp As Integer

'Para Trabajo con Grupos Globales de NT
Private Type MungeLong
  X As Long
  Dummy As Integer
End Type

Private Type MungeInt
  XLo As Integer
  XHi As Integer
  Dummy As Integer
End Type

Private Declare Function StrLenW Lib "kernel32.dll" Alias "lstrlenW" (ByVal Ptr As Long) As Long
'Private Declare Function NetGroupEnum Lib "netapi32.dll" (servername As Byte, ByVal level As Long, buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long
'Private Declare Function NetUserGetGroups Lib "netapi32.dll" (servername As Byte, UserName As Byte, ByVal level As Long, buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long) As Long

'Para Guardar Los Grupos Globales de NT
Dim Groups() As String
Dim Pass As Long


Private MatMenuPermisos() As String
Private sCadGruposdeUsu As String

Public Function TieneAlgunPermiso() As Boolean
    If IsArray(MatMenuPermisos) Then
        If UBound(MatMenuPermisos) > 0 Then
            TieneAlgunPermiso = True
        Else
            TieneAlgunPermiso = False
        End If
    Else
        TieneAlgunPermiso = False
    End If
End Function

'Private Sub EnumGlobalGroups(ByVal Server As String, Optional ByVal UserName As String)
'  ' Enumerates global groups only - not local groups
'  ' Returns an array of global groups
'  ' If a username is specified, it only returns
'  ' groups that that user is a member of
'  Dim Result As Long
'  Dim bufptr As Long
'  Dim EntriesRead As Long
'  Dim TotalEntries As Long
'  Dim ResumeHandle As Long
'  Dim BufLen As Long
'  Dim SNArray() As Byte
'  Dim GNArray(99) As Byte
'  Dim UNArray() As Byte
'  Dim GName As String
'  Dim I As Integer
'  Dim UNPtr As Long
'  Dim TempPtr As MungeLong
'  Dim TempStr As MungeInt
'
'  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server
'
'  SNArray = Server & vbNullChar      ' Move to byte array
'  UNArray = UserName & vbNullChar    ' Move to Byte array
'  BufLen = 255                       ' Buffer size
'  ResumeHandle = 0                   ' Start with the first entry
'
'  Pass = 0
'  Do
'    If UserName = "" Then
'      Result = NetGroupEnum(SNArray(0), 0, bufptr, BufLen, EntriesRead, TotalEntries, ResumeHandle)
'    Else
'      Result = NetUserGetGroups(SNArray(0), UNArray(0), 0, bufptr, BufLen, EntriesRead, TotalEntries)
'    End If
'    If Result <> 0 And Result <> 234 Then    ' 234 means multiple reads required
'      Err.Raise Result, "EnumGlobalGroups", "Error enumerating global group " & EntriesRead & " of " & TotalEntries
'      Exit Sub
'    End If
'    For I = 1 To EntriesRead
'      ' Get pointer to string from beginning of buffer
'      ' Copy 4 byte block of memory in 2 steps
'      PtrToInt TempStr.XLo, bufptr + (I - 1) * 4, 2
'      PtrToInt TempStr.XHi, bufptr + (I - 1) * 4 + 2, 2
'      LSet TempPtr = TempStr ' munge 2 Integers to a Long
'      ' Copy string to array and convert to a string
'      Result = PtrToStr(GNArray(0), TempPtr.x)
'      GName = Left(GNArray, StrLenW(TempPtr.x))
'      ReDim Preserve Groups(0 To Pass) As String
'      Groups(Pass) = GName
'      Pass = Pass + 1
'    Next I
'  Loop Until EntriesRead = TotalEntries
'  ' The above condition only valid for reading accounts on NT
'  ' but not OK for OS/2 or LanMan
'  NetAPIBufferFree bufptr         ' Don't leak memory
'
'End Sub

Public Function DameMaquinasdeUsuario(ByVal psUser As String, ByVal psDominio As String) As Variant
Dim User As IADsUser
Dim m As Variant
    On Error Resume Next
    Set User = GetObject("WinNT://" & psDominio & "/" & psUser & ",user")
    m = User.LoginWorkstations
    DameMaquinasdeUsuario = m
End Function

Public Function IniciarNuevoDia() As Boolean
Dim sSQL As String
Dim oConec As DCOMConecta
Dim R As ADODB.Recordset
Dim dFecCieSis As Date
Dim dFecIniDia As Date

    Set oConec = New DCOMConecta
    oConec.AbreConexion
    'Obtener fecha de Cierre del Sistema
    sSQL = "Select nConsSisValor from ConstSistema Where nConsSisCod = 13"
    Set R = oConec.CargaRecordSet(sSQL)
    dFecCieSis = CDate(R!nConsSisValor)
    'Obtener fecha de Inicio de Dia
    sSQL = "Select nConsSisValor from ConstSistema Where nConsSisCod = 15"
    Set R = oConec.CargaRecordSet(sSQL)
    dFecIniDia = CDate(R!nConsSisValor)
    
    If dFecCieSis = dFecIniDia Then
        IniciarNuevoDia = True
    Else
        IniciarNuevoDia = False
    End If
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function CuentaBloqueada(ByVal psUser As String, ByVal psDominio As String) As Boolean
Dim User As IADsUser

    Set User = GetObject("WinNT://" & psDominio & "/" & psUser & ",user")
    CuentaBloqueada = User.IsAccountLocked
    
End Function

Public Function DameItemsMenu() As ADODB.Recordset
Dim Conn As DCOMConecta
Dim sSQL As String
Dim R As ADODB.Recordset
Dim i As Integer
Dim Y As Integer
Dim lnFiltroM As Integer

        Set Conn = New DCOMConecta
        If Not Conn.AbreConexion() Then
            vError = True
            sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
            Set Conn = Nothing
            Exit Function
        End If
        
        sSQL = "SELECT nConsSisValor FROM CONSTSISTEMA WHERE nConsSisCod = 101"
        Set R = Conn.CargaRecordSet(sSQL, adLockReadOnly)
                    
        lnFiltroM = R!nConsSisValor
        Set R = Nothing
        
        'FILTRO PARA LIMA (PIGNORATICIO)
        If lnFiltroM = 2 Then       'Muestra Pigno Lima
            sSQL = "Select right(cName,12) as cCodigo,cDescrip,cName from Menu Where cAplicacion = '" & "1" & _
                "' AND cName <> 'M030100000002' AND Substring(cName,1,7) <> 'M030103' Order By nOrden"
        ElseIf lnFiltroM = 1 Then   'Muestra Pigno Trujillo
            sSQL = "Select right(cName,12) as cCodigo,cDescrip,cName from Menu Where cAplicacion = '" & "1" & _
                "' AND cName <> 'M030100000003' AND Substring(cName,1,7) <> 'M030104' Order By nOrden"
        Else                        'Muestra Ambos Lima y Trujillo
            sSQL = "Select right(cName,12) as cCodigo,cDescrip,cName from Menu Where cAplicacion = '" & "1" & "' Order By nOrden"
        End If
        
        Set R = Conn.CargaRecordSet(sSQL, adLockReadOnly)
        nNumMenus = R.RecordCount
        i = 0
        sCadMenu = ""
        sCadMenuSql = "('"
        MenuItems = ""
        
        'EN VES DE CADENA PASAR A MATRIZ PARA PODER EVITAR EL DESBORDE DE MEMORIA
        sCadMenu = sCadMenu & Trim(R!cname)
        Set DameItemsMenu = R
End Function

Public Function DameRSOperaciones() As ADODB.Recordset
Dim rsVar As Recordset
Dim Conn As DCOMConecta
    Dim sSQL As String
    Dim lsFiltroMon As String
    
    Set Conn = New DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set Conn = Nothing
        Exit Function
    End If
    
    sSQL = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O Where O.cOpeVisible = '1'"
    sSQL = sSQL & " Order by O.cOpeCod, O.nOpeNiv"
    Set rsVar = Conn.CargaRecordSet(sSQL, adLockReadOnly)
    Set DameRSOperaciones = rsVar
    Set rsVar = Nothing
    Conn.CierraConexion
    Set Conn = Nothing
    
End Function

Public Function DameUsuario() As String
    If nNumUsu > nPosUsuGrp Then
        DameUsuario = LstUsuarios(nPosUsuGrp)
        nPosUsuGrp = nPosUsuGrp + 1
    Else
        DameUsuario = ""
    End If
End Function

Public Function DameGrupo() As String
    If nPosUsuGrp < UBound(LstUsuarios) Then
        DameGrupo = LstUsuarios(nPosUsuGrp)
    Else
        DameGrupo = ""
    End If
    nPosUsuGrp = nPosUsuGrp + 1
End Function

Public Function DameLogedUsuario() As String
    DameLogedUsuario = LstUsuariosLoged(nPosUsuGrp - 1)
End Function

Public Function DameUsuarioNameFull() As String
    DameUsuarioNameFull = LstUsuariosNameFull(nPosUsuGrp - 1)
End Function

Public Sub DenegarAcceso(ByVal psUsuGrp As String, ByVal psItemMenuName As String, ByVal psTipoUsu As String)
Dim sSQL As String
Dim Conn As DCOMConecta
    vError = False
    sSQL = "DELETE Permiso where cGrupoUsu = '" & psUsuGrp & "' And cName = '" & psItemMenuName & "' And cTipo = '" & psTipoUsu & "' And cMenuOpe = '1'"
    Set Conn = New DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar sSQL
    Conn.CierraConexion
    Set Conn = Nothing
End Sub
Public Sub OtorgarAcceso(ByVal psUsuGrp As String, ByVal psItemMenuName As String, ByVal psTipoUsu As String)
Dim sSQL As String
Dim Conn As DCOMConecta
    vError = False
    sSQL = "INSERT INTO PERMISO(cName,cGrupoUsu,cTipo,cMenuOpe,cAplicacion) VALUES('" & psItemMenuName & "','" & psUsuGrp & "','" & psTipoUsu & "','1'," & Str("1") & ")"
    Set Conn = New DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar sSQL
    Conn.CierraConexion
    Set Conn = Nothing
End Sub

Public Sub DenegarOperacion(ByVal psUsuGrp As String, ByVal psItemMenuName As String, ByVal psTipoUsu As String)
Dim sSQL As String
Dim Conn As DCOMConecta
    vError = False
    sSQL = "DELETE Permiso where cGrupoUsu = '" & psUsuGrp & "' And cName = '" & psItemMenuName & "' And cTipo = '" & psTipoUsu & "' And cMenuOpe = '2'"
    Set Conn = New DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar sSQL
    Conn.CierraConexion
    Set Conn = Nothing
End Sub
Public Sub OtorgarOperacion(ByVal psUsuGrp As String, ByVal psItemMenuName As String, ByVal psTipoUsu As String)
Dim sSQL As String
Dim Conn As DCOMConecta
    vError = False
    sSQL = "INSERT INTO PERMISO(cName,cGrupoUsu,cTipo,cMenuOpe,cAplicacion) VALUES('" & psItemMenuName & "','" & psUsuGrp & "','" & psTipoUsu & "','2','" & "1" & "')"
    Set Conn = New DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar sSQL
    Conn.CierraConexion
    Set Conn = Nothing
End Sub
Public Function GetPDC(pdc As String) As Long
   Dim Result As Long, Server As String, Domain As String
   Dim SNArray() As Byte
   Dim DArray() As Byte
   Dim DCNPtr As Long
   Dim StrArray(100) As Byte
   SNArray = Server & vbNullChar      ' Move to byte array
   DArray = Domain & vbNullChar       ' Move to byte array
   Result = NetGetDCName(SNArray(0), DArray(0), DCNPtr)
   GetPDC = Result
   If Result = 0 Then
      Result = PtrToStr(StrArray(0), DCNPtr)
      pdc = Left(StrArray(), StrLen(DCNPtr))
   Else
      pdc = ""
   End If
   NetAPIBufferFree (DCNPtr)
End Function
Private Sub GetUserInfo(ByVal psPDC As String, User As String, UserName, Logged)
    On Error Resume Next
    Dim Result&, bufptr&, LOn As Long, LOff As Long
    Dim SNArray() As Byte, UNArray() As Byte, StrArray(500) As Byte
    Dim TempPtr As JoinLong, TempStr As JoinInt, X&, pdc$
    Let X = GetPDC(psPDC)
    pdc = psPDC
    SNArray = pdc & vbNullChar
    UNArray = User & vbNullChar
    Result = NetUserGetInfo(SNArray(0), UNArray(0), 3, bufptr)
    'DoEvents
    If Result = 0 Then
        Result = PtrToInt(TempStr.Bottom, bufptr + 36, 2)
        Result = PtrToInt(TempStr.Top, bufptr + 38, 2)
        LSet TempPtr = TempStr
        Result = PtrToStr(StrArray(0), TempPtr.X)
        UserName = Left(StrArray, StrLen(TempPtr.X))
        Result = PtrToInt(TempStr.Bottom, bufptr + 52, 2)
        Result = PtrToInt(TempStr.Top, bufptr + 54, 2)
        LSet TempPtr = TempStr
        LOn = TempPtr.X
        Result = PtrToInt(TempStr.Bottom, bufptr + 56, 2)
        Result = PtrToInt(TempStr.Top, bufptr + 58, 2)
        LSet TempPtr = TempStr
        LOff = TempPtr.X
        If LOn > LOff Then
            Logged = "On"
        Else
            Logged = "Off"
        End If
        Result = NetAPIBufferFree(bufptr)
    End If
End Sub

Private Sub GetUsers(ByVal Domain As String)
    On Error Resume Next
    Dim SNArray() As Byte, level&, Index&, EntriesRequested&, _
    PreferredMaximumLength&, ReturnedEntryCount&, SortedBuffer&, _
    APIResult As Long, StrArray(500) As Byte, i&, TempPtr As JoinLong, _
    TempStr As JoinInt, data$(), Result&, Size%
    Let level = 1
    Let SNArray = Domain & vbNullChar
    Let Index = 0
    Let EntriesRequested = 500
    Let PreferredMaximumLength = 6000
    Do
        'DoEvents
        APIResult = NetQueryDisplayInformation(SNArray(0), level, Index, EntriesRequested, PreferredMaximumLength, ReturnedEntryCount, SortedBuffer)
        If ReturnedEntryCount = 0 Then Exit Do
        For i = 1 To ReturnedEntryCount
            Let Size = Size + 1
            APIResult = PtrToInt(TempStr.Bottom, SortedBuffer + (i - 1) * 24, 2)
            APIResult = PtrToInt(TempStr.Top, SortedBuffer + (i - 1) * 24 + 2, 2)
            LSet TempPtr = TempStr
            APIResult = PtrToStr(StrArray(0), TempPtr.X)
            ReDim Preserve Users$(1 To Size)
            Users(Size) = Left(StrArray, StrLen(TempPtr.X))
            APIResult = PtrToInt(TempStr.Bottom, SortedBuffer + (i - 1) * 24 + 20, 2)
            APIResult = PtrToInt(TempStr.Top, SortedBuffer + (i - 1) * 24 + 22, 2)
            LSet TempPtr = TempStr
            Index = TempPtr.X
            'DoEvents
        Next i
        Result = NetAPIBufferFree(SortedBuffer)
    Loop Until APIResult = 0
End Sub

Public Sub CargaControlUsuarios(ByVal psDominio As String)
Dim i As Integer
Dim container As IADsContainer
Dim containername As String
Dim User As IADsUser

'Recuperar Los Usuarios del Dominio
    Set container = GetObject("WinNT://" & psDominio)
    
    container.Filter = Array("User")
    i = 0
    nNumUsu = 0
    ReDim LstUsuarios(i)
    ReDim LstUsuariosNameFull(i)
    For Each User In container
        If Len(User.Name) = 4 Then
            ReDim Preserve LstUsuarios(i + 1)
            ReDim Preserve LstUsuariosNameFull(i + 1)
            LstUsuarios(i) = User.Name
            'este proceso de obtener el fullname demora se debera considerar
            'LstUsuariosNameFull(i) = User.FullName
            i = i + 1
            nNumUsu = nNumUsu + 1
        End If
    Next
    
'    Call GetUsers(psDominio)
'    For i = 1 To UBound(Users)
'        If Len(Users(i)) <= 5 Then
'            nNumUsu = nNumUsu + 1
'            ReDim Preserve LstUsuarios(nNumUsu)
'            ReDim Preserve LstUsuariosNameFull(nNumUsu)
'            ReDim Preserve LstUsuariosLoged(nNumUsu)
'            LstUsuarios(nNumUsu - 1) = Users(i)
            'Call GetUserInfo(psDominio, LstUsuarios(nNumUsu - 1), LstUsuariosNameFull(nNumUsu - 1), LstUsuariosLoged(nNumUsu - 1))
'        End If
'    Next
End Sub

Public Function MostarNombre(ByVal psDominio As String, ByVal psIniciales As String)
Dim sLoged As String
Dim sNombre As String
    Call GetUserInfo(psDominio, psIniciales, sNombre, sLoged)
    MostarNombre = sNombre
End Function

Public Sub CargaControlGrupos(ByVal psDominio As String)
Dim container As IADsContainer
Dim containername As String
Dim group As IADsGroup

    Set container = GetObject("WinNT://" & psDominio)
    container.Filter = Array("Group")
    nNumGrupos = 0
    ReDim LstUsuarios(nNumGrupos)
    For Each group In container
        If UCase(Left(group.Name, 5)) = "GRUPO" Or UCase(Left(group.Name, 3)) = "GG_" Then
            ReDim Preserve LstUsuarios(nNumGrupos + 1)
            LstUsuarios(nNumGrupos) = group.Name
            nNumGrupos = nNumGrupos + 1
        End If
    Next
    nPosUsuGrp = 0
End Sub

Public Sub RegeneraMenu(ByVal MatMenu As Variant)
Dim Ctl As Control
Dim sSQL As String
Dim Conn As DCOMConecta
Dim R As ADODB.Recordset
Dim sCriterio As String
Dim sCadMenu As String
Dim i As Integer
Dim J As Integer
Dim sTemp As String

    On Error GoTo ErrRegeneraMenu
    vError = False
    sMsgError = ""
    
    Set Conn = New DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Set Conn = Nothing
        Exit Sub
    End If
    
    sTemp = ""
    For i = 0 To UBound(MatMenu) - 1
        sTemp = sTemp & "'" & MatMenu(i, 0) & Right("00" & MatMenu(i, 2), 2) & "',"
    Next i
    sTemp = Mid(sTemp, 1, Len(sTemp) - 1)
    
    'Eliminar Menu
    sSQL = "DELETE Menu Where cAplicacion = '" & Trim(Str("1")) & "' "
    Conn.ConexionActiva.Execute sSQL
    
    'Eliminar los que ya no estan en el Menu
    sSQL = "Delete Permiso where cName in (Select cName from Menu where cName not in (" & sTemp & ")) AND cAplicacion = '" & Trim(Str("1")) & "'"
    Conn.ConexionActiva.Execute sSQL
    sSQL = "Delete Menu where cName in (Select cName from Menu where cName not in (" & sTemp & ")) AND cAplicacion = '" & Trim(Str("1")) & "'"
    Conn.ConexionActiva.Execute sSQL
    
    'Ingresa Nuevos Menus
    sCadMenu = ""
    sSQL = "Select cName,cDescrip from Menu where cAplicacion = '" & Trim("1") & "'"
    Set R = Conn.CargaRecordSet(sSQL, adLockOptimistic)
    Do While Not R.EOF
        sCadMenu = sCadMenu & R!cname
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
        
    For i = 0 To UBound(MatMenu) - 1
        If InStr(sCadMenu, MatMenu(i, 0) & Right("00" & MatMenu(i, 2), 2)) <= 0 Then
            sSQL = "INSERT INTO MENU(CNAME,CDescrip,nOrden,cAplicacion) VALUES('" & MatMenu(i, 0) & Right("00" & MatMenu(i, 2), 2) & "','" & Left(MatMenu(i, 1), 45) & "'," & i + 1 & ",'" & "1" & "')"
            Conn.ConexionActiva.Execute sSQL
        End If
    Next i
    
    Set Conn = Nothing
    Exit Sub
    
ErrRegeneraMenu:
    vError = True
    sMsgError = Err.Description
    Set Conn = Nothing
    Exit Sub
End Sub
Public Function ObtenerNombreComputadora() As String
Dim buffMaq As String
Dim lSizeMaq As Long

    buffMaq = Space(255)
    lSizeMaq = Len(buffMaq)
    Call GetComputerName(buffMaq, lSizeMaq)
    ObtenerNombreComputadora = buffMaq
End Function
Public Function ObtenerUsuario() As String
Dim buffUsu As String
Dim lSizeUsu As Long

    buffUsu = String(255, " ")
    lSizeUsu = Len(buffUsu)
    Call GetUserName(buffUsu, lSizeUsu)
    ObtenerUsuario = Mid(Trim(buffUsu), 1, Len(Trim(buffUsu)) - 1)
       
End Function

Public Function VerificarUsuarioExistaEnRRHH(ByVal psUser As String) As Boolean
Dim sSQL As String
Dim oConn As DCOMConecta
Dim R As ADODB.Recordset
    sSQL = "Select cUser from RRHH Where cUser = '" & psUser & "'"
    Set oConn = New DCOMConecta
    oConn.AbreConexion
    Set R = oConn.CargaRecordSet(sSQL)
    oConn.CierraConexion
    Set oConn = Nothing
    
    If R.RecordCount > 0 Then
        VerificarUsuarioExistaEnRRHH = True
    Else
        VerificarUsuarioExistaEnRRHH = False
    End If
    R.Close
    Set oConn = Nothing
    Set R = Nothing
End Function


Public Sub CargaGruposUsuario(ByVal sUsuario As String, ByVal psDominio As String)
Dim i As Integer
Dim container As IADsContainer
Dim containername As String
Dim group As IADsGroup
Dim User As IADsUser

    'Call EnumGlobalGroups(psPDC, sUsuario)
'    For I = 0 To Pass - 1
'        nNumGrupos = nNumGrupos + 1
'        ReDim Preserve sGrupoUsu(nNumGrupos)
'        sGrupoUsu(nNumGrupos - 1) = Groups(I)
'    Next I
    
    Set container = GetObject("WinNT://" & psDominio)
    Set User = GetObject("WinNT://" & psDominio & "/" & sUsuario & ",user")
    container.Filter = Array("Group")
    i = 0
    sCadGruposdeUsu = "'"
    ReDim sGrupoUsu(i)
    For Each group In User.Groups
        If UCase(Left(group.Name, 5)) = "GRUPO" Or UCase(Left(group.Name, 3)) = "GG_" Then
            ReDim Preserve sGrupoUsu(i + 1)
            sGrupoUsu(i) = group.Name
            sCadGruposdeUsu = sCadGruposdeUsu & sGrupoUsu(i) & "','"
            i = i + 1
        End If
    Next
    
    If sCadGruposdeUsu <> "" And Len(sCadGruposdeUsu) > 1 Then
        sCadGruposdeUsu = Mid(sCadGruposdeUsu, 1, Len(sCadGruposdeUsu) - 2)
    End If
    nNumGrupos = 0
End Sub

Public Sub AgregaGrupoAUsuario(ByVal psDominio As String, ByVal psUsuario As String, ByVal psGrupo As String)
Dim group As IADsGroup
Dim User As IADsUser

Set User = GetObject("WinNT://" & psDominio & "/" & psUsuario & ",user")
Set group = GetObject("WinNT://" & psDominio & "/" & psGrupo & ",group")

Call group.Add(User.ADsPath)
group.SetInfo

End Sub

Public Sub EliminaGrupodeUsuario(ByVal psDominio As String, ByVal psUsuario As String, ByVal psGrupo As String)
Dim group As IADsGroup
Dim User As IADsUser

Set User = GetObject("WinNT://" & psDominio & "/" & psUsuario & ",user")
Set group = GetObject("WinNT://" & psDominio & "/" & psGrupo & ",group")

Call group.Remove(User.ADsPath)

End Sub
Public Function DameGrupoUsuario() As String
        
    If nNumGrupos < UBound(sGrupoUsu) Then
        DameGrupoUsuario = sGrupoUsu(nNumGrupos)
    Else
        DameGrupoUsuario = ""
    End If
    nNumGrupos = nNumGrupos + 1
End Function

Private Function PerteneceAGrupo(ByVal sGrupoPerm As String) As Boolean
Dim i As Integer
    PerteneceAGrupo = False
    If Not IsArray(sGrupoUsu) Then
        Exit Function
    End If
    For i = 0 To UBound(sGrupoUsu) - 1
        If UCase(Trim(sGrupoPerm)) = UCase(sGrupoUsu(i)) Then
            PerteneceAGrupo = True
            Exit For
        End If
    Next i
End Function

Public Function TienePermiso(ByVal psName As String, ByVal psIndex As String, Optional ByVal pbOperacion As Boolean = False) As Boolean
Dim i As Integer
    TienePermiso = False
    For i = 1 To UBound(MatMenuPermisos)
        If Not pbOperacion Then
            If Left(MatMenuPermisos(i), 11) = psName And Right(MatMenuPermisos(i), 2) = psIndex Then
                TienePermiso = True
                Exit For
            End If
        Else
            If Left(MatMenuPermisos(i), 11) = psName Then
                TienePermiso = True
                Exit For
            End If
        End If
    Next i
    
End Function

Public Function CargaMenu(ByVal psDominio As String, Optional ByVal psUsuario As String = "", Optional ByVal sTipoUsu As String = "1")
Dim sSQL As String
Dim oConn As DCOMConecta
Dim RMenu As ADODB.Recordset
Dim RPermisoUsu As ADODB.Recordset
Dim RPermisoGrp As ADODB.Recordset
Dim sUsuario As String
Dim sCriterio As String
Dim i As Integer
Dim Y As Integer

    'Carga Grupos
    If Len(Trim(psUsuario)) = 0 Then
        sUsuario = UCase(ObtenerUsuario())
    Else
        sUsuario = UCase(psUsuario)
    End If
    If sTipoUsu = "1" Then 'Si es un usuario de lo contrario un grupo
        Call CargaGruposUsuario(sUsuario, psDominio)
    End If
    
    Set oConn = New DCOMConecta
    If oConn.AbreConexion() Then
    
        If psUsuario <> "" And sTipoUsu = "1" Then
            sCadGruposdeUsu = sCadGruposdeUsu & ",'" & psUsuario & "' "
            
            sSQL = "Select * from permiso where ((cGrupoUsu in (" & sCadGruposdeUsu & ")) OR ( cGrupoUsu = '" & psUsuario & "')) " & " And cAplicacion = '" & "1" & "' "
            Set RMenu = oConn.CargaRecordSet(sSQL)
            oConn.CierraConexion
            
            sCadMenu = ""
            sCadMenuGrp = ""
            
            Do While Not RMenu.EOF
                sCadMenu = sCadMenu & "*" & Trim(RMenu!cname) & ","
                If UCase(Trim(RMenu!cGrupoUsu)) <> UCase(psUsuario) Then
                    sCadMenuGrp = sCadMenuGrp & "*" & Trim(RMenu!cname) & ","
                End If
                Y = Y + 1
                ReDim Preserve MatMenuPermisos(Y)
                MatMenuPermisos(Y) = RMenu!cname
            
                RMenu.MoveNext
            Loop
            RMenu.Close
            Set RMenu = Nothing
        Else
        
            sSQL = "Select * from Menu where cAplicacion = '" & "1" & "'"
            Set RMenu = oConn.CargaRecordSet(sSQL)
            'permisos por usuario
            sSQL = "Select * from Permiso where cTipo='" & sTipoUsu & "' and upper(rtrim(ltrim(cGrupoUsu))) = '" & UCase(sUsuario) & "' And cAplicacion = '" & "1" & "'"
            Set RPermisoUsu = oConn.CargaRecordSet(sSQL)
            'Permisos por Grupo
            sSQL = "Select * from Permiso where cTipo='" & vsTipoPermisoGrp & "' And cAplicacion = '" & "1" & "'"
            Set RPermisoGrp = oConn.CargaRecordSet(sSQL)
            oConn.CierraConexion
            
            sCadMenu = ""
            sCadMenuGrp = ""
            'For I = 0 To nNumGrupos - 1
            '    If UCase(sGrupoUsu(I)) = "GG_SISTEMAS" Then
            '        sCadMenu = "mnuSeguridad09000000mnuSegurPerm09010000"
            '        sCadMenuGrp = "mnuSeguridad09000000mnuSegurPerm09010000"
            '        Exit For
            '    End If
            'Next I
            Y = 0
            Do While Not RMenu.EOF
                'If RMenu!cname = "M070100000000" Then
                '    sCriterio = sCriterio
                'End If
                'Buscando por usuario
                sCriterio = "cName = '" & Trim(RMenu!cname) & "'"
                If RPermisoUsu.RecordCount > 0 Then
                    RPermisoUsu.MoveFirst
                    RPermisoUsu.Find sCriterio, , , 1
                End If
                If Not RPermisoUsu.EOF Then
                    sCadMenu = sCadMenu & Trim(RPermisoUsu!cname) & ","
                    Y = Y + 1
                    ReDim Preserve MatMenuPermisos(Y)
                    MatMenuPermisos(Y) = RPermisoUsu!cname
                Else
                    'Buscando por Grupo
                    If sTipoUsu = "1" Then
                        sCriterio = "cName = '" & Trim(RMenu!cname) & "'"
                        RPermisoGrp.Filter = sCriterio
                        If RPermisoGrp.RecordCount > 0 Then
                            RPermisoGrp.MoveFirst
                            Do While Not RPermisoGrp.EOF
                                If PerteneceAGrupo(Trim(RPermisoGrp!cGrupoUsu)) Then
                                    sCadMenu = sCadMenu & Trim(RPermisoGrp!cname) & ","
                                    sCadMenuGrp = sCadMenuGrp & Trim(RPermisoGrp!cname) & ","
                                    Y = Y + 1
                                    ReDim Preserve MatMenuPermisos(Y)
                                    MatMenuPermisos(Y) = RPermisoGrp!cname
                                End If
                                RPermisoGrp.MoveNext
                            Loop
                        End If
                        RPermisoGrp.Filter = adFilterNone
                    End If
                End If
                RMenu.MoveNext
            Loop
            RMenu.Close
            
            '******************************************************************************
            '**********************   Para Operaciones  ***********************************
            '******************************************************************************
             oConn.AbreConexion
            sSQL = "Select cOpeCod as cName,cOpeDesc from OpeTpo where cOpeVisible ='1' Order by cOpeCod, cOpeDesc"
            Set RMenu = oConn.CargaRecordSet(sSQL)
            oConn.CierraConexion
            Do While Not RMenu.EOF
                'Buscando por usuario
                sCriterio = "cName = '" & Trim(RMenu!cname) & "'"
                If RPermisoUsu.RecordCount > 0 Then
                    RPermisoUsu.MoveFirst
                    RPermisoUsu.Find sCriterio, , , 1
                End If
                If Not RPermisoUsu.EOF Then
                    sCadMenu = sCadMenu & "*" & Trim(RPermisoUsu!cname) & ","
                    Y = Y + 1
                    ReDim Preserve MatMenuPermisos(Y)
                    MatMenuPermisos(Y) = RPermisoUsu!cname
                Else
                    'Buscando por Grupo
                    If sTipoUsu = "1" Then
                        sCriterio = "cName = '" & Trim(RMenu!cname) & "'"
                        RPermisoGrp.Filter = sCriterio
                        If RPermisoGrp.RecordCount > 0 Then
                            RPermisoGrp.MoveFirst
                            Do While Not RPermisoGrp.EOF
                                If PerteneceAGrupo(Trim(RPermisoGrp!cGrupoUsu)) Then
                                    sCadMenu = sCadMenu & "*" & Trim(RPermisoGrp!cname) & ","
                                    sCadMenuGrp = sCadMenuGrp & "*" & Trim(RPermisoGrp!cname) & ","
                                    Y = Y + 1
                                    ReDim Preserve MatMenuPermisos(Y)
                                    MatMenuPermisos(Y) = RPermisoGrp!cname
                                End If
                                RPermisoGrp.MoveNext
                            Loop
                        End If
                        RPermisoGrp.Filter = adFilterNone
                    End If
                End If
                RMenu.MoveNext
            Loop
            
            RMenu.Close
            Set RMenu = Nothing
            RPermisoGrp.Close
            Set RPermisoGrp = Nothing
            RPermisoUsu.Close
            Set RPermisoUsu = Nothing
        End If
    End If
    Set oConn = Nothing
End Function
Public Function InterconexionCorrecta() As Boolean
Dim oConn As DCOMConecta
    Set oConn = New DCOMConecta
    If oConn.AbreConexion() Then
       InterconexionCorrecta = True
       oConn.CierraConexion
    Else
       InterconexionCorrecta = False
    End If
    Set oConn = Nothing
End Function

Public Sub Desbloquear_Habilitar_Cuenta(ByVal psDominio As String, ByVal psUser As String)
Dim User As IADsUser

    Set User = GetObject("WinNT://" & psDominio & "/" & psUser & ",user")
    User.AccountDisabled = False
    User.SetInfo
    
End Sub

Public Sub AsignarAccesoATodasMaquinas(ByVal psDominio As String, ByVal psUser As String)
Dim User As IADsUser

    Set User = GetObject("WinNT://" & psDominio & "/" & psUser & ",user")
    User.LoginWorkstations = ""
    User.SetInfo
    
End Sub

Public Sub AsignarAccesoAMaquinas(ByVal psDominio As String, ByVal psUser As String, ByVal m As Variant)
Dim User As IADsUser
Dim i As Integer
Dim sCad As String

    Set User = GetObject("WinNT://" & psDominio & "/" & psUser & ",user")
    sCad = ""
    For i = 0 To UBound(m) - 1
        sCad = sCad & m(i) & ","
    Next i
    If Len(sCad) > 0 Then
        sCad = Mid(sCad, 1, Len(sCad) - 1)
    End If
    m(0) = sCad
    User.LoginWorkstations = sCad
    User.SetInfo
    
End Sub

Public Function ChangePassword(ByVal psDominio As String, ByVal psUser As String, ByVal psOldPass As String, ByVal psNewPass As String) As Boolean
Dim User As IADsUser

Set User = GetObject("WinNT://" & psDominio & "/" & psUser & ",user")
    On Error Resume Next
    Call User.ChangePassword(psOldPass, psNewPass)
    User.SetInfo
    
End Function

Public Function ClaveIncorrectaNT(ByVal psUsuario As String, ByVal psClave As String, ByVal psDominio As String) As Boolean
    ClaveIncorrectaNT = SSPValidateUser(psUsuario, psDominio, psClave)
End Function
Private Sub Class_Initialize()
    vsTipoPermisoUsu = "1"
    vsTipoPermisoGrp = "2"
    nPosUsuGrp = 0
    nNumMenus = 0
    Espaciado = 60
    ReDim MatMenuPermisos(0)
End Sub

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
ai.flags = SEC_WINNT_AUTH_IDENTITY_ANSI

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






