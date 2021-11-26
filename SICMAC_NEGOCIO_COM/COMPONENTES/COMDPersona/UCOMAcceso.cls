VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UCOMAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit

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
Private Declare Function NetUserChangePassword Lib "netapi32.dll" (ByVal DomainName As String, ByVal UserName As String, ByVal OldPassword As String, ByVal NewPassword As String) As Long
     
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
Dim sSql As String
Dim oConec As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
Dim dFecCieSis As Date
Dim dFecIniDia As Date

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    'Obtener fecha de Cierre del Sistema
    sSql = "Select nConsSisValor from ConstSistema Where nConsSisCod = 13"
    Set R = oConec.CargaRecordSet(sSql)
    dFecCieSis = CDate(R!nConsSisValor)
    'Obtener fecha de Inicio de Dia
    sSql = "Select nConsSisValor from ConstSistema Where nConsSisCod = 15"
    Set R = oConec.CargaRecordSet(sSql)
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

'**Modificado por DAOR 20071122
'**Se agregó el parametro psGrupoUsu
Public Function DameItemsMenu(Optional ByVal psGrupoUsu As String = "", Optional psAplicacion As String = "1") As ADODB.Recordset
Dim Conn As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset
Dim i As Integer
Dim Y As Integer
Dim lnFiltroM As Integer

        Set Conn = New COMConecta.DCOMConecta
        If Not Conn.AbreConexion() Then
            vError = True
            sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
            Set Conn = Nothing
            Exit Function
        End If
        
        '**Comentado por DAOR 20071122*********************************************
        'sSql = "SELECT nConsSisValor FROM CONSTSISTEMA WHERE nConsSisCod = 101"
        'Set R = Conn.CargaRecordSet(sSql, adLockReadOnly)
        '
        'lnFiltroM = R!nConsSisValor
        'Set R = Nothing
        '
        ''FILTRO PARA LIMA (PIGNORATICIO)
        'If lnFiltroM = 2 Then       'Muestra Pigno Lima
        '    sSql = "Select right(cName,12) as cCodigo,cDescrip,cName from Menu Where cAplicacion = '" & "1" & _
        '        "' AND cName <> 'M030100000002' AND Substring(cName,1,7) <> 'M030103' Order By nOrden"
        'ElseIf lnFiltroM = 1 Then   'Muestra Pigno Trujillo
        '    sSql = "Select right(cName,12) as cCodigo,cDescrip,cName from Menu Where cAplicacion = '" & "1" & _
        '        "' AND cName <> 'M030100000003' AND Substring(cName,1,7) <> 'M030104' Order By nOrden"
        'Else                        'Muestra Ambos Lima y Trujillo
        '    sSql = "Select right(cName,12) as cCodigo,cDescrip,cName from Menu Where cAplicacion = '" & "1" & "' Order By nOrden"
        'End If
        '**************************************************************************
        
        sSql = "exec stp_sel_Menus '" & psGrupoUsu & "','" & psAplicacion & "'"
        Set R = Conn.CargaRecordSet(sSql, adLockReadOnly)
        nNumMenus = R.RecordCount
        i = 0
        sCadMenu = ""
        sCadMenuSql = "('"
        MenuItems = ""
        
        'EN VES DE CADENA PASAR A MATRIZ PARA PODER EVITAR EL DESBORDE DE MEMORIA
        sCadMenu = sCadMenu & Trim(R!cname)
        Set DameItemsMenu = R
End Function

'**Modificado por DAOR 20071122
'**Se agreg{o el parametro psGrupoUsu
Public Function DameRSOperaciones(Optional psGrupoUsu As String = "") As ADODB.Recordset
Dim rsVar As Recordset
Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    Dim lsFiltroMon As String
    
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set Conn = Nothing
        Exit Function
    End If
    
    '**Modificado por DAOR 20071122 *******************************************
    'sSql = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O Where O.cOpeVisible = '1'"
    'sSql = sSql & " Order by O.cOpeCod, O.nOpeNiv"
    sSql = "exec stp_sel_Operaciones '" & psGrupoUsu & "'"
    '**************************************************************************
    
    Set rsVar = Conn.CargaRecordSet(sSql, adLockReadOnly)
    Set DameRSOperaciones = rsVar
    Set rsVar = Nothing
    Conn.CierraConexion
    Set Conn = Nothing
    
End Function

'AMDO 20130726 TI-ERS086-2013 *********
Public Function DameRSOpcEsp() As ADODB.Recordset
Dim rsVar As Recordset
Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    Dim lsFiltroMon As String
    
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set Conn = Nothing
        Exit Function
    End If

    sSql = "exec stp_sel_ERS0862013_OpcEspecial"

    
    Set rsVar = Conn.CargaRecordSet(sSql, adLockReadOnly)
    Set DameRSOpcEsp = rsVar
    Set rsVar = Nothing
    Conn.CierraConexion
    Set Conn = Nothing
End Function

Public Function DameRSCargos() As ADODB.Recordset
Dim rsVar As Recordset
Dim Conn As COMConecta.DCOMConecta
    Dim sSql As String
    Dim lsFiltroMon As String
    
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set Conn = Nothing
        Exit Function
    End If

    sSql = "exec stp_sel_ERS0862013_Cargos"
    
    Set rsVar = Conn.CargaRecordSet(sSql, adLockReadOnly)
    Set DameRSCargos = rsVar
    Set rsVar = Nothing
    Conn.CierraConexion
    Set Conn = Nothing
End Function

Public Function DameSelCargosOpcEsp(ByVal nOpcEsp As Integer) As ADODB.Recordset
    Dim Conn As New COMConecta.DCOMConecta
    Dim sSql As String
    Dim lsFiltroMon As String
    Conn.AbreConexion
    sSql = "stp_sel_ERS0862013_SelCargosOpcEsp " & nOpcEsp
    Set DameSelCargosOpcEsp = Conn.CargaRecordSet(sSql, adLockReadOnly)
    Conn.CierraConexion
    Set Conn = Nothing
End Function

Public Sub InsertaPermisoCargo(ByVal pnOpcEsp As Integer, ByVal psCargoCod As String)
Dim Conn As New COMConecta.DCOMConecta
Dim sSql As String
    vError = False
    sSql = "INSERT INTO PermisoCargo(nOpcEsp,cRHCargoCod) VALUES(" & pnOpcEsp & ",'" & psCargoCod & "')"
    Set Conn = New DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar sSql
    Conn.CierraConexion
    Set Conn = Nothing
End Sub

Public Sub EliminaPermisoCargo(ByVal pnOpcEsp As Integer, ByVal psCargoCod As String)
Dim Conn As New COMConecta.DCOMConecta
Dim sSql As String
    vError = False
    sSql = "DELETE PermisoCargo where nOpcEsp= " & pnOpcEsp & " And cRHCargoCod ='" & psCargoCod & "'"
    Set Conn = New DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar sSql
    Conn.CierraConexion
    Set Conn = Nothing
End Sub
'END AMDO ***************************************************

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

Public Sub DenegarAcceso(ByVal psUsuGrp As String, ByVal psItemMenuName As String, ByVal psTipoUsu As String, Optional psAplicacion As String = "1")
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
    vError = False
    sSql = "DELETE Permiso where cGrupoUsu = '" & psUsuGrp & "' And cName = '" & psItemMenuName & "' And cTipo = '" & psTipoUsu & "' And cMenuOpe = '1' AND cAplicacion='" & psAplicacion & "'"
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar sSql
    Conn.CierraConexion
    Set Conn = Nothing
End Sub
'MADM 20110302 - store
Public Sub OtorgarAcceso(ByVal psUsuGrp As String, ByVal psItemMenuName As String, ByVal psTipoUsu As String, Optional ByVal pdFecVen As Date = "01/01/1900", Optional psAplicacion As String = "1")
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
    vError = False
    sSql = "exec stp_ins_Permiso '" & psItemMenuName & "','" & psUsuGrp & "','" & psTipoUsu & "','1','" & psAplicacion & "','" & Format(pdFecVen, "mm/dd/yyyy") & "'" 'WIOR 20130913 MODIFICÓ pdFecVen A FORMAT(pdFecVen, "mm/dd/yyyy")
    'sSql = "INSERT INTO PERMISO(cName,cGrupoUsu,cTipo,cMenuOpe,cAplicacion) VALUES('" & psItemMenuName & "','" & psUsuGrp & "','" & psTipoUsu & "','1'," & Str("1") & ")"
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar sSql
    Conn.CierraConexion
    Set Conn = Nothing
End Sub

Public Sub DenegarOperacion(ByVal psUsuGrp As String, ByVal psItemMenuName As String, ByVal psTipoUsu As String, Optional psAplicacion As String = "1")
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
    vError = False
    sSql = "DELETE Permiso where cGrupoUsu = '" & psUsuGrp & "' And cName = '" & psItemMenuName & "' And cTipo = '" & psTipoUsu & "' And cMenuOpe = '2' AND cAplicacion='" & psAplicacion & "'"
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar sSql
    Conn.CierraConexion
    Set Conn = Nothing
End Sub
'MADM 20110302 - STORE - PARAMETRO OPCIONAL FECHA
Public Sub OtorgarOperacion(ByVal psUsuGrp As String, ByVal psItemMenuName As String, ByVal psTipoUsu As String, Optional ByVal pdFecha As Date = "1900/01/01", Optional psAplicacion As String = "1")
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
    vError = False
    sSql = "exec stp_ins_Permiso '" & psItemMenuName & "','" & psUsuGrp & "','" & psTipoUsu & "','2','" & "1" & "','" & Format(pdFecha, "mm/dd/yyyy") & "','" & psAplicacion & "'"
''    sSql = "INSERT INTO PERMISO(cName,cGrupoUsu,cTipo,cMenuOpe,cAplicacion) VALUES('" & psItemMenuName & "','" & psUsuGrp & "','" & psTipoUsu & "','2','" & "1" & "')"
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar sSql
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

Public Sub RegeneraMenu(ByVal MatMenu As Variant, Optional psAplicacion As String = "1")
Dim Ctl As Control
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
Dim sCriterio As String
Dim sCadMenu As String
Dim i As Integer
Dim J As Integer
Dim sTemp As String

    On Error GoTo ErrRegeneraMenu
    vError = False
    sMsgError = ""
    
    Set Conn = New COMConecta.DCOMConecta
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
    sSql = "DELETE Menu Where cAplicacion = '" & psAplicacion & "' "
    Conn.ConexionActiva.Execute sSql
    
    'Eliminar los que ya no estan en el Menu
    sSql = "Delete Permiso where cName in (Select cName from Menu where cName not in (" & sTemp & ")) AND cAplicacion = '" & psAplicacion & "'"
    Conn.ConexionActiva.Execute sSql
    sSql = "Delete Menu where cName in (Select cName from Menu where cName not in (" & sTemp & ")) AND cAplicacion = '" & psAplicacion & "'"
    Conn.ConexionActiva.Execute sSql
    
    'Ingresa Nuevos Menus
    sCadMenu = ""
    sSql = "Select cName,cDescrip from Menu where cAplicacion = '" & psAplicacion & "'"
    Set R = Conn.CargaRecordSet(sSql, adLockOptimistic)
    Do While Not R.EOF
        sCadMenu = sCadMenu & R!cname
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
        
    For i = 0 To UBound(MatMenu) - 1
        If InStr(sCadMenu, MatMenu(i, 0) & Right("00" & MatMenu(i, 2), 2)) <= 0 Then
            sSql = "INSERT INTO MENU(CNAME,CDescrip,nOrden,cAplicacion) VALUES('" & MatMenu(i, 0) & Right("00" & MatMenu(i, 2), 2) & "','" & Left(MatMenu(i, 1), 45) & "'," & i + 1 & ",'" & psAplicacion & "')"
            Conn.ConexionActiva.Execute sSql
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
Dim sSql As String
Dim oConn As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
    sSql = "Select cUser from RRHH Where cUser = '" & psUser & "'"
    Set oConn = New COMConecta.DCOMConecta
    oConn.AbreConexion
    Set R = oConn.CargaRecordSet(sSql)
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
    
    sCadGruposdeUsu = IIf(sCadGruposdeUsu = "'", "", sCadGruposdeUsu) 'DAOR 20070206
    
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

Public Function CargaMenu(ByVal psDominio As String, Optional ByVal psUsuario As String = "", Optional ByVal sTipoUsu As String = "1", Optional bExisteAlgunPermisoFecha As Boolean = False, Optional pgdFecsis As Date = "01/01/1900", Optional psAplicacion As String = "1")
Dim sSql As String
Dim oConn As COMConecta.DCOMConecta
Dim RMenu As ADODB.Recordset
Dim RPermisoUsu As ADODB.Recordset
Dim RPermisoGrp As ADODB.Recordset
Dim sUsuario As String
Dim sCriterio As String
Dim i As Integer
Dim Y As Integer
Dim pCadGrupo As String 'MADM 20110303
Dim lsAplicacionGrupoNegocio As String 'ALPA20160219

    pCadGrupo = ""
    'Carga Grupos
    If Len(Trim(psUsuario)) = 0 Then
        sUsuario = UCase(ObtenerUsuario())
    Else
        sUsuario = UCase(psUsuario)
    End If
    
    lsAplicacionGrupoNegocio = psAplicacion
    
            
    If sTipoUsu = "1" Then 'Si es un usuario de lo contrario un grupo
        Call CargaGruposUsuario(sUsuario, psDominio)
    End If
    
    ' *** RIRO 20131108 RFIII
    Dim clsGen As New COMDConstSistema.DCOMGeneral
    Dim rsTemp As New ADODB.Recordset
    Dim sRF3, sTemporal As String
    
    Set rsTemp = clsGen.GetConstante(10027, , "100", "1")
    
    If Not rsTemp Is Nothing Then
        If Not rsTemp.EOF And Not rsTemp.BOF Then
            sRF3 = rsTemp!cDescripcion
        Else
            sRF3 = String(10, "-")
        End If
    Else
        sRF3 = String(10, "-")
    End If
    
    sTemporal = sCadGruposdeUsu
    
    If InStr(1, sCadGruposdeUsu, sRF3) > 0 Then
                
        Dim oPersona As New COMDPersona.DCOMPersonas
        Dim rsRF3 As New ADODB.Recordset
        Set rsRF3 = oPersona.RecuperarGruposRF3(sUsuario)
                
        If Not rsRF3 Is Nothing Then
            If Not rsRF3.BOF And Not rsRF3.EOF Then
                If rsRF3!nEstado = 1 Then
                    sCadGruposdeUsu = "'" & sRF3 & "'"
                Else
                    sCadGruposdeUsu = Replace(sCadGruposdeUsu, sRF3, "")
                End If
            Else
                sCadGruposdeUsu = Replace(sCadGruposdeUsu, sRF3, "")
            End If
        Else
            sCadGruposdeUsu = Replace(sCadGruposdeUsu, sRF3, "")
        End If
    
    End If
    ' *** END RIRO
    
    Set oConn = New COMConecta.DCOMConecta
    If oConn.AbreConexion() Then
    
        If psUsuario <> "" And sTipoUsu = "1" Then
            'sCadGruposdeUsu = sCadGruposdeUsu & ",'" & psUsuario & "' "
           
            '**DAOR 20070206********************************************Capi
            If sCadGruposdeUsu = "" Then
                sCadGruposdeUsu = "'" & psUsuario & "'"
                sTemporal = "'" & psUsuario & "'"
            Else
                sCadGruposdeUsu = sCadGruposdeUsu & ",'" & psUsuario & "' "
                sTemporal = sTemporal & ",'" & psUsuario & "' " ' RIRO20131102 RFIII
                pCadGrupo = Replace(sCadGruposdeUsu, "'", "")
                
            End If
            '***********************************************************
            ''MADM 20110302 SE AGREGO FECHA DE VENC
            'sSql = "Select * from permiso where ((cGrupoUsu in (" & sCadGruposdeUsu & ")) OR ( cGrupoUsu = '" & psUsuario & "')) " & " And cAplicacion = '" & "1" & "'"
            'ALPA20160219

            sSql = " exec stp_sel_DevuelvePermisosGrupoUsuarioFecha '" & pCadGrupo & "','" & psUsuario & "','" & lsAplicacionGrupoNegocio & "','" & Format(pgdFecsis, "yyyymmdd") & "' "
           
            Set RMenu = oConn.CargaRecordSet(sSql)
            oConn.CierraConexion
            
            sCadMenu = ""
            sCadMenuGrp = ""
            bExisteAlgunPermisoFecha = IIf(Not RMenu.EOF, True, False)
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
        
            sSql = "Select * from Menu where cAplicacion in (select valor from dbo.fnc_getTblValoresTexto('" & lsAplicacionGrupoNegocio & "'))"
            Set RMenu = oConn.CargaRecordSet(sSql)
            'permisos por usuario
            sSql = "Select * from Permiso where cTipo='" & sTipoUsu & "' and upper(rtrim(ltrim(cGrupoUsu))) = '" & UCase(sUsuario) & "' And cAplicacion in (select valor from dbo.fnc_getTblValoresTexto('" & lsAplicacionGrupoNegocio & "'))"
            Set RPermisoUsu = oConn.CargaRecordSet(sSql)
            'Permisos por Grupo
            sSql = "Select * from Permiso where cTipo='" & vsTipoPermisoGrp & "' And cAplicacion in (select valor from dbo.fnc_getTblValoresTexto('" & lsAplicacionGrupoNegocio & "'))"
            Set RPermisoGrp = oConn.CargaRecordSet(sSql)
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
            sSql = "Select cOpeCod as cName,cOpeDesc from OpeTpo where cOpeVisible ='1' Order by cOpeCod, cOpeDesc"
            Set RMenu = oConn.CargaRecordSet(sSql)
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
        sCadGruposdeUsu = sTemporal ' *** RIRO 20131108 RF III
    End If
    Set oConn = Nothing
End Function
Public Function InterconexionCorrecta() As Boolean
Dim oConn As COMConecta.DCOMConecta
    Set oConn = New COMConecta.DCOMConecta
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

Public Sub CargaControles(ByVal psDominio As String, _
                        ByRef pMatUser As Variant, _
                        ByRef pMatGrupo As Variant, _
                        ByRef prsMenu As ADODB.Recordset, _
                        ByRef prsOpe As ADODB.Recordset, _
                        Optional ByVal psCoduser As String = "", _
                        Optional ByRef prsOpcEsp As ADODB.Recordset, _
                        Optional ByRef prsCargos As ADODB.Recordset, Optional psAplicacion As String = "1") 'DAOR 20071122
                        'AMDO TI-ERS086-2013 20130726, Se agrego prsOpcEsp,prsCargos
    
Dim sCadTemp As String
Dim i As Integer
            
    Call CargaControlUsuarios(psDominio)
    
    ReDim pMatUser(nNumUsu, 2)
    
    sCadTemp = DameUsuario
    i = 0
    
    While sCadTemp <> ""
        pMatUser(i, 0) = sCadTemp
        pMatUser(i, 1) = DameUsuarioNameFull
        sCadTemp = DameUsuario
        i = i + 1
    Wend
    
    Call CargaControlGrupos(psDominio)
    sCadTemp = DameGrupo
    
    i = 0
    While sCadTemp <> ""
        ReDim Preserve pMatGrupo(i + 1)
        pMatGrupo(i) = sCadTemp
        sCadTemp = DameUsuario
        i = i + 1
    Wend
    
    nPosUsuGrp = 0
    
    '**Modificado por DAOR 20071122******************
    'Set prsMenu = DameItemsMenu
    'Set prsOpe = DameRSOperaciones
    Set prsMenu = DameItemsMenu(psCoduser, psAplicacion)
    Set prsOpe = DameRSOperaciones(psCoduser)
    '*************************************************
    
    '**AMDO 20130726 TI-ERS086-2013 *************
    Set prsOpcEsp = DameRSOpcEsp()
    Set prsCargos = DameRSCargos()
    'END AMDO *********************************
    
End Sub

Public Sub CargarLogin(ByVal psUsuario As String, _
                        ByVal psClave As String, _
                        ByRef pbInterconexCorrecta As Boolean, _
                        ByRef pbClaveIncorrecta As Boolean, _
                        ByRef prsITF As ADODB.Recordset, _
                        ByRef prsVar As ADODB.Recordset, _
                        ByRef psServerName As String, _
                        ByRef psDatabaseName As String, _
                        ByRef psCadenaConexion As String, _
                        ByRef prsMenu As ADODB.Recordset, _
                        ByRef pbTieneAlgunPermiso As Boolean, _
                        ByRef pMatOperac As Variant, _
                        ByRef pnNumOperac As Integer, _
                        ByRef pbTienePermisoRetiroSinFirma As Boolean, _
                        ByRef pbIniciarNuevoDia As Boolean, _
                        ByRef pRsOpeF2 As ADODB.Recordset, _
                        ByRef pRsExtor As ADODB.Recordset, _
                        ByRef pRsOpeRecep As ADODB.Recordset, _
                        ByRef pRsOpeLlam As ADODB.Recordset, _
                        Optional psAplicacion As String = "1")
    
Dim sTitulo As String
Dim i As Integer
Dim Y As Integer
Dim oConec As COMConecta.DCOMConecta
Dim sSql As String
Dim oITF As COMDConstSistema.FCOMITF
Dim R As ADODB.Recordset
Dim psDominio As String
    
'Para las Operaciones F2 y de Extornos
Dim sFiltro() As String
Dim lnFiltraTC As Integer
Dim lnFiltraMP As Integer
Dim oGen As COMDConstSistema.DCOMGeneral
    
'Para el Manejo de la Matriz de Operaciones
Dim MatIndices() As Integer

    pbInterconexCorrecta = InterconexionCorrecta()
    
    If Not pbInterconexCorrecta Then Exit Sub
        
    Set oConec = New COMConecta.DCOMConecta
    
    'Carga Variables del Sistema
    sSql = " Select nConsSisCod,nConsSisValor,nConsSisDesc From ConstSistema " _
            & " Where nConsSisCod in (" & gConstSistFechaSistema & "," & gConstSistNombreAbrevCMAC & "," _
            & gConstSistRutaBackup & "," & gConstSistCodCMAC & "," & gConstSistMargenSupCartas & "," & gConstSistMagenIzqCartas & "," _
            & gConstSistMargenDerCartas & "," & gConstSistNroLineasPagina & "," & gConstSistNroLineasOrdenPago & "," _
            & gConstSistCtaConversionMEDol & "," & gConstSistCtaConversiónMESoles & "," & gConstSistTipoConverión & "," _
            & gConstSistNombreModulo & "," & gConstSistFechaInicioDia & "," & gConstSistDominio & "," & gConstSistPDC _
            & ",40,50," & gConstPersCodCMACT & "," & gConstSistAgenciaEspecial & "," & gConstSistCierreMesNegocio & ", " _
            & gConstSistRutaIcono & ",84,85," & gConstSistVerificaRegistroEfectivo & ") " _
            & "ORDER BY nConsSisCod"
    oConec.AbreConexion
                        
    Set prsVar = oConec.CargaRecordSet(sSql)
    
    psServerName = oConec.servername
    psDatabaseName = oConec.DatabaseName
    psCadenaConexion = oConec.CadenaConexion
    
    'Carga Dominio
    sSql = "SELECT nConsSisValor FROM ConstSistema WHERE nConsSisCod=37"
    Set R = oConec.CargaRecordSet(sSql)
    psDominio = Trim(R!nConsSisValor)
    
    pbClaveIncorrecta = ClaveIncorrectaNT(psUsuario, psClave, psDominio)
    If Not pbClaveIncorrecta Then Exit Sub
    
    'Carga Operaciones
    sSql = "Select cOpeCod,cOpeDesc,cOpeVisible,nOpeNiv,cOpeGruCod,cUltimaActualizacion from OpeTpo WHERE cOpeVisible ='1' Order by cOpeCod"
    Set R = oConec.CargaRecordSet(sSql)
    
    oConec.CierraConexion
    Set oConec = Nothing
        
    'Carga Parametros de ITF
    Set oITF = New COMDConstSistema.FCOMITF
    Set prsITF = oITF.fgITFParametroForm
    Set oITF = Nothing

    
    Set prsMenu = DameItemsMenu(, psAplicacion)
    
    Call CargaMenu(psDominio, psUsuario, , , , psAplicacion)
    
    pbTieneAlgunPermiso = TieneAlgunPermiso
    
    'Habilita Permiso para Operaciones y Reportes
    Y = 0
    i = 0
    Do While Not R.EOF
        If TienePermiso(R!cOpeCod, "", True) Then
            Y = Y + 1
            ReDim Preserve MatIndices(i + 1)
            MatIndices(i) = R.Bookmark
            i = i + 1
        End If
        R.MoveNext
    Loop
    
    pnNumOperac = Y
    ReDim pMatOperac(Y, 5)
    
    R.MoveFirst
    i = 0
    Do While Not R.EOF
        If R.Bookmark = MatIndices(i) Then
            pMatOperac(i, 0) = R!cOpeCod
            pMatOperac(i, 1) = R!cOpeDesc
            pMatOperac(i, 2) = IIf(IsNull(R!cOpeGruCod), "", R!cOpeGruCod)
            pMatOperac(i, 3) = R!cOpeVisible
            pMatOperac(i, 4) = R!nOpeNiv
            i = i + 1
        End If
        R.MoveNext
    Loop
   
    
    'Permiso Retiro sin tener en cuenta las firmas
    pbTienePermisoRetiroSinFirma = TienePermiso(gCapPermEspretSinFirmas, "", True)
        
    'Permiso Para Inicio de Dia
    pbIniciarNuevoDia = False
    If TienePermiso(gCapPermEspInicioDia, "", True) Then
       pbIniciarNuevoDia = IniciarNuevoDia
    End If
    
    'Cargar Operaciones F2 y de Extornos
    Set oGen = New COMDConstSistema.DCOMGeneral
    lnFiltraTC = CInt(oGen.LeeConstSistema(102))
    lnFiltraMP = CInt(oGen.LeeConstSistema(103))
    
    'F2
    ReDim sFiltro(5)
    If lnFiltraMP = 1 Then
        sFiltro(1) = "[1][01234][012]" 'Pigno Trujillo
    ElseIf lnFiltraMP = 2 Then
        sFiltro(1) = "[1][01345][012]"   'Pigno Lima
    ElseIf lnFiltraMP = 0 Then
        sFiltro(1) = "[1][012345][012]"   'Ambos
    End If
           
    sFiltro(2) = "[23][0-2][0123]"    'Captaciones
          
    If lnFiltraTC = 0 Then
        sFiltro(3) = "90002[0-3]"       'Compra Venta
    ElseIf lnFiltraTC = 1 Then
        sFiltro(3) = "90002[0-6]"
    End If
    
    sFiltro(4) = "9010[01][0123456789]"    'Control de Efectivo Boveda y Cajero
    sFiltro(5) = "90003[0-5]"    'Operaciones con Cheques
    
    Set pRsOpeF2 = oGen.GetOperacionesUsuarioMatriz(psUsuario, sFiltro, pMatOperac, pnNumOperac)

    'Extornos
    ReDim sFiltro(9)
    sFiltro(1) = "1[034][79][012]"         'Extornos de Colocaciones
    If lnFiltraMP = 1 Then
        sFiltro(2) = "129"          'Extornos de Prendario Trujillo
    ElseIf lnFiltraMP = 2 Then      'Extornos de Prendario Lima
        sFiltro(2) = "159"
    ElseIf lnFiltraMP = 0 Then      'Extornos de Prendario Lima
        sFiltro(2) = "1[25]9"
    End If
    sFiltro(3) = "2[3457]"      'Extornos de Captaciones
    sFiltro(4) = "3[569]"       'Extornos de Otras Operaciones
    If lnFiltraTC = 0 Then
        sFiltro(5) = "90900[0-3]"
    ElseIf lnFiltraTC = 1 Then
        '  sfiltro(1) = "90900[0-6]"
        sFiltro(5) = "90900[0-6]"
    End If
    sFiltro(6) = "90103[0-9]"   'Extornos de Operaciones de Boveda
    sFiltro(7) = "90102[1-9]"   'Extornos de Operaciones de Cajero
    sFiltro(8) = "90003[6-9]"   'Extornos de Operaciones con Cheque
    sFiltro(9) = "90004[4-6]"   'Extornos de Compra Venta - Tipo de Cambio Especial

    Set pRsExtor = oGen.GetOperacionesUsuarioMatriz(psUsuario, sFiltro, pMatOperac, pnNumOperac)
    
    'Operaciones con CMAC Recepcion
    ReDim sFiltro(4)
    sFiltro(1) = "260[0-3]" 'Operaciones de Captaciones
    sFiltro(2) = "126"      'Operaciones de Prendario
    sFiltro(3) = "106"      'Operaciones de Colocaciones
    sFiltro(4) = "136"      'Operaciones de judiciales

    Set pRsOpeRecep = oGen.GetOperacionesUsuarioMatriz(psUsuario, sFiltro, pMatOperac, pnNumOperac)
    
    'Operaciones con CMAC Llamada
    ReDim sFiltro(3)
    sFiltro(1) = "2605"     'Operaciones de Captaciones
    sFiltro(2) = "127"      'Operaciones de Prendario
    sFiltro(3) = "107"      'Operaciones de Colocaciones
    
    Set pRsOpeLlam = oGen.GetOperacionesUsuarioMatriz(psUsuario, sFiltro, pMatOperac, pnNumOperac)
    
    Set oGen = Nothing
End Sub

Public Sub CargaGrupoUsuarios(ByVal psDominio As String, pMatGrupo As Variant)
                            
Dim sCadTemp As String
Dim i As Integer
    Call CargaControlGrupos(psDominio)
    sCadTemp = DameGrupo
    
    While sCadTemp <> ""
        ReDim Preserve pMatGrupo(i + 1)
        pMatGrupo(i) = sCadTemp
        sCadTemp = DameGrupo 'DameUsuario
        i = i + 1
    Wend
    
End Sub

Public Function CargaUsuarioGrupo(ByVal sUsuario As String, ByVal psDominio As String) As String
Dim i As Integer
Dim container As IADsContainer
Dim containername As String
Dim group As IADsGroup
Dim User As IADsUser
Dim sCadGruposdeUsu As String

    Set container = GetObject("WinNT://" & psDominio)
    Set User = GetObject("WinNT://" & psDominio & "/" & sUsuario & ",user")
    container.Filter = Array("Group")
    sCadGruposdeUsu = ""
    For Each group In User.Groups
        If UCase(Left(group.Name, 5)) = "GRUPO" Or UCase(Left(group.Name, 3)) = "GG_" Then
            If sCadGruposdeUsu = "" Then
                sCadGruposdeUsu = group.Name
            Else
                sCadGruposdeUsu = sCadGruposdeUsu & "','" & group.Name
            End If
        End If
    Next
    CargaUsuarioGrupo = sCadGruposdeUsu
End Function

''MAVM
'Public Function CargaUsuarioGrupoAuditoria(ByVal sUsuario As String, ByVal psDominio As String) As String
'Dim i As Integer
'Dim container As IADsContainer
'Dim containername As String
'Dim group As IADsGroup
'Dim User As IADsUser
'Dim sCadGruposdeUsu As String
'
'    Set container = GetObject("WinNT://" & psDominio)
'    Set User = GetObject("WinNT://" & psDominio & "/" & sUsuario & ",user")
'    container.Filter = Array("Group")
'    sCadGruposdeUsu = ""
'    For Each group In User.Groups
'        If UCase(Left(group.Name, 5)) = "GRUPO" Or UCase(Left(group.Name, 3)) = "GG_" Then
'            If sCadGruposdeUsu = "" Then
'                sCadGruposdeUsu = "'" & group.Name & "'"
'            Else
'                sCadGruposdeUsu = sCadGruposdeUsu & ",'" & group.Name & "'"
'            End If
'        End If
'    Next
'    CargaUsuarioGrupoAuditoria = sCadGruposdeUsu
'End Function

Private Sub Class_Initialize()
    vsTipoPermisoUsu = "1"
    vsTipoPermisoGrp = "2"
    nPosUsuGrp = 0
    nNumMenus = 0
    Espaciado = 60
    ReDim MatMenuPermisos(0)
End Sub

'ARCV 18-07-2006 (Performance)
'**********************************************************************

Public Function TienePermiso_NEW(ByVal psName As String, ByVal psIndex As String, Optional ByVal pbOperacion As Boolean = False) As Boolean
Dim i As Integer
    TienePermiso_NEW = False
        For i = 1 To UBound(MatMenuPermisos)
            If Not pbOperacion Then
                If MatMenuPermisos(i) = psName Then
                    TienePermiso_NEW = True
                    Exit For
                End If
            Else
                If Left(MatMenuPermisos(i), 11) = psName Then
                    TienePermiso_NEW = True
                    Exit For
                End If
            End If
        Next i
End Function




Public Sub CargarLogin_NEW(ByVal psUsuario As String, _
                        ByVal psClave As String, _
                        ByRef pbInterconexCorrecta As Boolean, _
                        ByRef pbClaveIncorrecta As Boolean, _
                        ByRef prsITF As ADODB.Recordset, _
                        ByRef prsVar As ADODB.Recordset, _
                        ByRef psServerName As String, _
                        ByRef psDatabaseName As String, _
                        ByRef psCadenaConexion As String, _
                        ByRef pMatMenuPer As Variant, _
                        ByRef pbTieneAlgunPermiso As Boolean, _
                        ByRef pbTienePermisoRetiroSinFirma As Boolean, _
                        ByRef pbIniciarNuevoDia As Boolean, _
                        ByRef pRsOpeF2 As ADODB.Recordset, _
                        ByRef pRsExtor As ADODB.Recordset, _
                        ByRef pRsOpeRecep As ADODB.Recordset, _
                        ByRef pRsOpeLlam As ADODB.Recordset, _
                        ByRef prsOpeRep As ADODB.Recordset, Optional psAplicacion As String = "1")
    
Dim sTitulo As String
Dim i As Integer
Dim Y As Integer
Dim oConec As COMConecta.DCOMConecta
Dim sSql As String
Dim oITF As COMDConstSistema.FCOMITF
Dim R As ADODB.Recordset
Dim psDominio As String
    
'Para las Operaciones F2 y de Extornos
Dim sFiltro() As String
Dim lnFiltraTC As Integer
Dim lnFiltraMP As Integer
Dim oGen As COMDConstSistema.DCOMGeneral
    
'Para el Manejo de la Matriz de Operaciones
Dim MatIndices() As Integer

'Para carga menu personalizado
Dim MatMenuPer() As Integer
Dim nNumMatMenuPer As Integer

Dim pMatOperac() As String
Dim pnNumOperac As Integer
Dim prsMenu As ADODB.Recordset
Dim pbExisteAlgunPermisoFecha As Boolean
Dim ftemp As Date

    pbInterconexCorrecta = InterconexionCorrecta()
    
    If Not pbInterconexCorrecta Then Exit Sub
        
    Set oConec = New COMConecta.DCOMConecta
    
    'Carga Variables del Sistema
    sSql = " Select nConsSisCod,nConsSisValor,nConsSisDesc From ConstSistema " _
            & " Where nConsSisCod in (" & gConstSistFechaSistema & "," & gConstSistNombreAbrevCMAC & "," _
            & gConstSistRutaBackup & "," & gConstSistCodCMAC & "," & gConstSistMargenSupCartas & "," & gConstSistMagenIzqCartas & "," _
            & gConstSistMargenDerCartas & "," & gConstSistNroLineasPagina & "," & gConstSistNroLineasOrdenPago & "," _
            & gConstSistCtaConversionMEDol & "," & gConstSistCtaConversiónMESoles & "," & gConstSistTipoConverión & "," _
            & gConstSistNombreModulo & "," & gConstSistFechaInicioDia & "," & gConstSistDominio & "," & gConstSistPDC _
            & ",40,50," & gConstPersCodCMACT & "," & gConstSistAgenciaEspecial & "," & gConstSistCierreMesNegocio & ", " _
            & gConstSistRutaIcono & ",84,85," & gConstSistVerificaRegistroEfectivo & ",301,302,303,304,305) " _
            & "ORDER BY nConsSisCod"
    
    'ARCV Codigo 301 para el Documento de Declaracion Jurada
    oConec.AbreConexion
                        
    Set prsVar = oConec.CargaRecordSet(sSql)
    
    'MADM 20110322
    Do While Not prsVar.EOF
        If Trim(prsVar!nConsSiscod) = 16 Then
            ftemp = CDate(Trim(prsVar!nConsSisValor))
            Exit Do
        End If
       prsVar.MoveNext
    Loop
    prsVar.MoveFirst
    'END MADM
    
    psServerName = oConec.servername
    psDatabaseName = oConec.DatabaseName
    psCadenaConexion = oConec.CadenaConexion
    
    'Carga Dominio
    sSql = "SELECT nConsSisValor FROM ConstSistema WHERE nConsSisCod=37"
    Set R = oConec.CargaRecordSet(sSql)
    psDominio = Trim(R!nConsSisValor)
    
    pbClaveIncorrecta = ClaveIncorrectaNT(psUsuario, psClave, psDominio)
    
    If Not pbClaveIncorrecta Then Exit Sub
    
    'Carga Operaciones
    sSql = "Select cOpeCod,cOpeDesc,cOpeVisible,nOpeNiv,cOpeGruCod,cUltimaActualizacion from OpeTpo WHERE cOpeVisible ='1' Order by cOpeCod"
    Set R = oConec.CargaRecordSet(sSql)
    
    oConec.CierraConexion
    Set oConec = Nothing

    'Carga Parametros de ITF
    Set oITF = New COMDConstSistema.FCOMITF
    Set prsITF = oITF.fgITFParametroForm
    Set oITF = Nothing

    
    Set prsMenu = DameItemsMenu(, psAplicacion)
    
    Call CargaMenu(psDominio, psUsuario, , pbExisteAlgunPermisoFecha, ftemp, psAplicacion)
    
    'MADM 20110302
    If pbExisteAlgunPermisoFecha = False Then
        pbTieneAlgunPermiso = False
        Exit Sub
    End If
    
    '*******************************************************************************************************
    'Filtra solo su perfil
    '*******************************************************************************************************
    ReDim MatMenuPer(prsMenu.RecordCount)
    nNumMatMenuPer = 0
    
    Do While Not prsMenu.EOF
        If TienePermiso_NEW(prsMenu!cname, "", False) Then
            MatMenuPer(nNumMatMenuPer) = 1
        Else
            MatMenuPer(nNumMatMenuPer) = 0
        End If
        nNumMatMenuPer = nNumMatMenuPer + 1
        prsMenu.MoveNext
    Loop
    
    pMatMenuPer = MatMenuPer
    
    '*******************************************************************************************************
    
    pbTieneAlgunPermiso = TieneAlgunPermiso
    
    'Habilita Permiso para Operaciones y Reportes
    Y = 0
    i = 0
    Do While Not R.EOF
        If TienePermiso(R!cOpeCod, "", True) Then
            Y = Y + 1
            ReDim Preserve MatIndices(i + 1)
            MatIndices(i) = R.Bookmark
            i = i + 1
        End If
        R.MoveNext
    Loop
    
    pnNumOperac = Y
    ReDim pMatOperac(Y, 5)
    
    R.MoveFirst
    i = 0
    
    'madm 20110302
    If Y = 0 Then
        ReDim Preserve MatIndices(0)
    End If
    'end madm
    
    Do While Not R.EOF
        If R.Bookmark = MatIndices(i) Then
            pMatOperac(i, 0) = R!cOpeCod
            pMatOperac(i, 1) = R!cOpeDesc
            pMatOperac(i, 2) = IIf(IsNull(R!cOpeGruCod), "", R!cOpeGruCod)
            pMatOperac(i, 3) = R!cOpeVisible
            pMatOperac(i, 4) = R!nOpeNiv
            i = i + 1
        End If
        R.MoveNext
    Loop
   
    
    'Permiso Retiro sin tener en cuenta las firmas
    pbTienePermisoRetiroSinFirma = TienePermiso(gCapPermEspretSinFirmas, "", True)
        
    'Permiso Para Inicio de Dia
    pbIniciarNuevoDia = False
    If TienePermiso(gCapPermEspInicioDia, "", True) Then
       pbIniciarNuevoDia = IniciarNuevoDia
    End If
    
    'Cargar Operaciones F2 y de Extornos
    Set oGen = New COMDConstSistema.DCOMGeneral
    lnFiltraTC = CInt(oGen.LeeConstSistema(102))
    lnFiltraMP = CInt(oGen.LeeConstSistema(103))
    
    'F2
'    '***Modificado por ELRO el 20120627, según OYP-RFC024-2012
'     ReDim sFiltro(5)

    'ANDE 20170512 DINERO ELECTRONICO
    'PASI20160613 CCE ************
    'ReDim sFiltro(6)
    ReDim sFiltro(8)
    '*****************************
    '*****************************
    
'    '***Modificado por ELRO***********************************
    If lnFiltraMP = 1 Then
        'sFiltro(1) = "[1][01234][012]" 'Pigno Trujillo
        sFiltro(1) = "[1][01234][0123]" 'Pigno Trujillo 'FRHU 20150504 ERS022-2015
    ElseIf lnFiltraMP = 2 Then
        sFiltro(1) = "[1][01345][012]"   'Pigno Lima
    ElseIf lnFiltraMP = 0 Then
        sFiltro(1) = "[1][012345][012]"   'Ambos
    End If
           
    sFiltro(2) = "[23][0-2][0123]"    'Captaciones
          
    If lnFiltraTC = 0 Then
        sFiltro(3) = "90002[0-3]"       'Compra Venta
    ElseIf lnFiltraTC = 1 Then
        sFiltro(3) = "90002[0-6]"
    End If
    
    sFiltro(4) = "9010[01][0123456789]"    'Control de Efectivo Boveda y Cajero
    sFiltro(5) = "90003[0-5]"    'Operaciones con Cheques
    
'    '***Agregado por ELRO el 20120627, según OYP-RFC024-2012
    sFiltro(6) = "92000[0-2]"    'Operaciones con Vouchers
    'PASI20160613 CCE *********
    sFiltro(7) = "93000"
    '**************************
    'ANDE 20170512 DINERO ELECTRONICO
    sFiltro(8) = "98000[0-2]"
    '**************************
    
'    '***Agregado por ELRO***********************************
    Set pRsOpeF2 = oGen.GetOperacionesUsuarioMatriz(psUsuario, sFiltro, pMatOperac, pnNumOperac)

    'Extornos
    'PASI20160824 CCE************
    'ReDim sFiltro(9)
    ReDim sFiltro(10)
    '****************************
    
    'Modificado por DAOR 20091112*************************************
    'sFiltro(1) = "1[034][79][012]"         'Extornos de Colocaciones
    sFiltro(1) = "1[0345][79][012]"         'Extornos de Colocaciones
    '*****************************************************************
    
    '**Modificado por DAOR  20091112, Las operaciones de extornos de prendario empiezan con 129
    'If lnFiltraMP = 1 Then
    '    sFiltro(2) = "129"          'Extornos de Prendario Trujillo
    'ElseIf lnFiltraMP = 2 Then      'Extornos de Prendario Lima
    '    sFiltro(2) = "159"
    'ElseIf lnFiltraMP = 0 Then      'Extornos de Prendario Lima
    '    sFiltro(2) = "1[25]9"
    'End If
    sFiltro(2) = "129"          'Extornos de Prendario
    '*************************************************************************************
    
    'sFiltro(3) = "2[3457]"      'Extornos de Captaciones
    sFiltro(3) = "2[34579]"      'Extornos de Captaciones 'JUEZ 20140523 Se agregó 9
    'sFiltro(4) = "3[569]"       'Extornos de Otras Operaciones
    sFiltro(4) = "3[3569]"       'Extornos de Otras Operaciones 'FRHU 20150126 ERS048 Se agrego 3: Extorno Notas
    If lnFiltraTC = 0 Then
        sFiltro(5) = "90900[0-3]"
    ElseIf lnFiltraTC = 1 Then
        '  sfiltro(1) = "90900[0-6]"
        sFiltro(5) = "90900[0-6]"
    End If
    sFiltro(6) = "90103[0-9]"   'Extornos de Operaciones de Boveda
    sFiltro(7) = "90102[1-9]"   'Extornos de Operaciones de Cajero
    sFiltro(8) = "90003[6-9]"   'Extornos de Operaciones con Cheque
    sFiltro(9) = "90004[4-6]"   'Extornos de Compra Venta - Tipo de Cambio Especial
    sFiltro(10) = "94000"       'Extornos de Transferencias Interbancarias CCE

    Set pRsExtor = oGen.GetOperacionesUsuarioMatriz(psUsuario, sFiltro, pMatOperac, pnNumOperac)
    
    'Operaciones con CMAC Recepcion
    ReDim sFiltro(4)
    sFiltro(1) = "260[0-3]" 'Operaciones de Captaciones
    sFiltro(2) = "126"      'Operaciones de Prendario
    sFiltro(3) = "106"      'Operaciones de Colocaciones
    sFiltro(4) = "136"      'Operaciones de judiciales

    Set pRsOpeRecep = oGen.GetOperacionesUsuarioMatriz(psUsuario, sFiltro, pMatOperac, pnNumOperac)
    
    'Operaciones con CMAC Llamada
    ReDim sFiltro(3)
    sFiltro(1) = "2605"     'Operaciones de Captaciones
    sFiltro(2) = "127"      'Operaciones de Prendario
    sFiltro(3) = "107"      'Operaciones de Colocaciones
    
    Set pRsOpeLlam = oGen.GetOperacionesUsuarioMatriz(psUsuario, sFiltro, pMatOperac, pnNumOperac)
    
    'ARCV 19-07-2006
    ReDim sFiltro(11)
    
    sFiltro(1) = "1088"
    sFiltro(2) = "1792"
    sFiltro(3) = "178"
    sFiltro(4) = "808"
    sFiltro(5) = "108"
    sFiltro(6) = "1380"
    sFiltro(7) = "1280[012]"
    sFiltro(8) = "1280[3]"
    sFiltro(9) = "1280[4]"
    sFiltro(10) = "14800"
    sFiltro(11) = "1791"
    
    Set prsOpeRep = oGen.GetOperacionesUsuarioMatriz(psUsuario, sFiltro, pMatOperac, pnNumOperac)
    '--------------------
    
    Set oGen = Nothing
End Sub

'**********************************************************************



'ARCV 17-09-2006 (Para el Financiero)

Public Sub CargarMenu_NEW(ByVal psUsuario As String, _
                        ByVal psClave As String, _
                        ByRef pMatMenuPer As Variant, _
                        ByRef prsOpeRep As ADODB.Recordset, _
                        ByVal psAplicacion As String)
    
Dim sTitulo As String
Dim i As Integer
Dim Y As Integer
Dim oConec As COMConecta.DCOMConecta
Dim sSql As String
Dim oITF As COMDConstSistema.FCOMITF
Dim R As ADODB.Recordset
Dim psDominio As String
    
'Para las Operaciones F2 y de Extornos
Dim sFiltro() As String
Dim lnFiltraTC As Integer
Dim lnFiltraMP As Integer
Dim oGen As COMDConstSistema.DCOMGeneral
    
'Para el Manejo de la Matriz de Operaciones
Dim MatIndices() As Integer

'Para carga menu personalizado
Dim MatMenuPer() As Integer
Dim nNumMatMenuPer As Integer

Dim pMatOperac() As String
Dim pnNumOperac As Integer
Dim prsMenu As ADODB.Recordset


    
    'Carga Operaciones
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    'Carga Dominio
    sSql = "SELECT nConsSisValor FROM ConstSistema WHERE nConsSisCod=37"
    Set R = oConec.CargaRecordSet(sSql)
    psDominio = Trim(R!nConsSisValor)
    
    sSql = "Select cOpeCod,cOpeDesc,cOpeVisible,nOpeNiv,cOpeGruCod,cUltimaActualizacion from OpeTpo WHERE cOpeVisible ='1' Order by cOpeCod"
    Set R = oConec.CargaRecordSet(sSql)

    oConec.CierraConexion
    Set oConec = Nothing

    Set prsMenu = DameItemsMenuAplicacion(psAplicacion)
    
    Call CargaMenuAplicacion(psDominio, psUsuario, , psAplicacion)
    
    '*******************************************************************************************************
    'Filtra solo su perfil
    '*******************************************************************************************************
    ReDim MatMenuPer(prsMenu.RecordCount)
    nNumMatMenuPer = 0
    
    Do While Not prsMenu.EOF
        If TienePermiso_NEW(prsMenu!cname, "", False) Then
            MatMenuPer(nNumMatMenuPer) = 1
        Else
            MatMenuPer(nNumMatMenuPer) = 0
        End If
        nNumMatMenuPer = nNumMatMenuPer + 1
        prsMenu.MoveNext
    Loop
    
    pMatMenuPer = MatMenuPer
    
    '*******************************************************************************************************
        
    'Habilita Permiso para Operaciones y Reportes
    Y = 0
    i = 0
    Do While Not R.EOF
        If TienePermiso(R!cOpeCod, "", True) Then
            Y = Y + 1
            ReDim Preserve MatIndices(i + 1)
            MatIndices(i) = R.Bookmark
            i = i + 1
        End If
        R.MoveNext
    Loop
    
    pnNumOperac = Y
    ReDim pMatOperac(Y, 5)
    
    R.MoveFirst
    i = 0
    Do While Not R.EOF
        If R.Bookmark = MatIndices(i) Then
            pMatOperac(i, 0) = R!cOpeCod
            pMatOperac(i, 1) = R!cOpeDesc
            pMatOperac(i, 2) = IIf(IsNull(R!cOpeGruCod), "", R!cOpeGruCod)
            pMatOperac(i, 3) = R!cOpeVisible
            pMatOperac(i, 4) = R!nOpeNiv
            i = i + 1
        End If
        R.MoveNext
    Loop
   
    'Cargar Operaciones F2 y de Extornos
    Set oGen = New COMDConstSistema.DCOMGeneral
        
    ReDim sFiltro(11)
    sFiltro(1) = Left(gContProvisionProveeMN, 2) & "_" & Mid(gContProvisionProveeMN, 4, 1)
    sFiltro(2) = "76"
    sFiltro(3) = "77"
    sFiltro(4) = "78"
    sFiltro(5) = "70[12]1"
    sFiltro(6) = Left("701221", 2) & "_" & Mid("701221", 4, 1)
    sFiltro(7) = "74"
    sFiltro(8) = "4[0-5]"
    sFiltro(9) = "9[0-5]%"
    sFiltro(10) = "4[6-9]"
    sFiltro(11) = "9"

    Set prsOpeRep = oGen.GetOperacionesUsuarioMatriz(psUsuario, sFiltro, pMatOperac, pnNumOperac)
    Set oGen = Nothing

End Sub

Public Function DameItemsMenuAplicacion(Optional ByVal psAplicacion As String = "1") As ADODB.Recordset
Dim Conn As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset
Dim i As Integer
Dim Y As Integer
Dim lnFiltroM As Integer

        Set Conn = New COMConecta.DCOMConecta
        If Not Conn.AbreConexion() Then
            vError = True
            sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
            Set Conn = Nothing
            Exit Function
        End If
        
        sSql = "SELECT nConsSisValor FROM CONSTSISTEMA WHERE nConsSisCod = 101"
        Set R = Conn.CargaRecordSet(sSql, adLockReadOnly)
                    
        lnFiltroM = R!nConsSisValor
        Set R = Nothing
        
        'FILTRO PARA LIMA (PIGNORATICIO)
        If psAplicacion = "1" Then
            If lnFiltroM = 2 Then       'Muestra Pigno Lima
                sSql = "Select right(cName,12) as cCodigo,cDescrip,cName from Menu Where cAplicacion = '" & psAplicacion & _
                    "' AND cName <> 'M030100000002' AND Substring(cName,1,7) <> 'M030103' Order By nOrden"
            ElseIf lnFiltroM = 1 Then   'Muestra Pigno Trujillo
                sSql = "Select right(cName,12) as cCodigo,cDescrip,cName from Menu Where cAplicacion = '" & psAplicacion & _
                    "' AND cName <> 'M030100000003' AND Substring(cName,1,7) <> 'M030104' Order By nOrden"
            Else                        'Muestra Ambos Lima y Trujillo
                sSql = "Select right(cName,12) as cCodigo,cDescrip,cName from Menu Where cAplicacion = '" & psAplicacion & "' Order By nOrden"
            End If
        Else
            sSql = "Select right(cName,12) as cCodigo,cDescrip,cName from Menu Where cAplicacion = '" & psAplicacion & "' Order By nOrden"
        End If
        Set R = Conn.CargaRecordSet(sSql, adLockReadOnly)
        nNumMenus = R.RecordCount
        i = 0
        sCadMenu = ""
        sCadMenuSql = "('"
        MenuItems = ""
        
        'EN VES DE CADENA PASAR A MATRIZ PARA PODER EVITAR EL DESBORDE DE MEMORIA
        sCadMenu = sCadMenu & Trim(R!cname)
        Set DameItemsMenuAplicacion = R
End Function


Public Function CargaMenuAplicacion(ByVal psDominio As String, Optional ByVal psUsuario As String = "", Optional ByVal sTipoUsu As String = "1", _
                                    Optional ByVal psAplicacion As String = "1")
Dim sSql As String
Dim oConn As COMConecta.DCOMConecta
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
    
    Set oConn = New COMConecta.DCOMConecta
    If oConn.AbreConexion() Then
    
        If psUsuario <> "" And sTipoUsu = "1" Then
            sCadGruposdeUsu = sCadGruposdeUsu & ",'" & psUsuario & "' "
            sSql = "Select * from permiso where ((cGrupoUsu in (" & sCadGruposdeUsu & ")) OR ( cGrupoUsu = '" & psUsuario & "')) " & " And cAplicacion = '" & psAplicacion & "' "
            Set RMenu = oConn.CargaRecordSet(sSql)
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
        
            sSql = "Select * from Menu where cAplicacion = '" & psAplicacion & "'"
            Set RMenu = oConn.CargaRecordSet(sSql)
               
            'permisos por usuario
            sSql = "Select * from Permiso where cTipo='" & sTipoUsu & "' and upper(rtrim(ltrim(cGrupoUsu))) = '" & UCase(sUsuario) & "' And cAplicacion = '" & psAplicacion & "'"
            Set RPermisoUsu = oConn.CargaRecordSet(sSql)
            'Permisos por Grupo
            sSql = "Select * from Permiso where cTipo='" & vsTipoPermisoGrp & "' And cAplicacion = '" & psAplicacion & "'"
            Set RPermisoGrp = oConn.CargaRecordSet(sSql)
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
            sSql = "Select cOpeCod as cName,cOpeDesc from OpeTpo where cOpeVisible ='1' Order by cOpeCod, cOpeDesc"
            Set RMenu = oConn.CargaRecordSet(sSql)
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

'**DAOR 20070510, Función que valida el visto bueno
'**del lavado de dinero
Public Function VistoLavDineroEsCorrecto(psUsuario As String, psClave As String, ByRef psVisPersCod As String) As Boolean
Dim oConec As COMConecta.DCOMConecta
Dim sSql As String
Dim oITF As COMDConstSistema.FCOMITF
Dim R As ADODB.Recordset
Dim psDominio As String
Dim bClaveEsCorrecta As Boolean

    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
        'Carga Dominio
        sSql = "SELECT nConsSisValor FROM ConstSistema WHERE nConsSisCod=37"
        Set R = oConec.CargaRecordSet(sSql)
        psDominio = Trim(R!nConsSisValor)
    
        bClaveEsCorrecta = ClaveIncorrectaNT(psUsuario, psClave, psDominio)
        
        VistoLavDineroEsCorrecto = bClaveEsCorrecta
        
        If Not bClaveEsCorrecta Then Exit Function
        
        'Se agrego filtro '007026' RFIII, RIRO20131102
        'By Capi 2001208
        sSql = "select top 1 RC.cPersCod from RRHH R inner join RHCargos RC on R.cPersCod=RC.cPersCod "
        sSql = sSql & " where R.cUser='" & psUsuario & "' and (RC.cRHCargoCod In ('006005','007006', '007026') or Substring(RC.cRHCargoCod,1,3)='002')"
        'By Capi 25022008 se añadio auxiliar administrativo para el visto electronico
                
        Set R = oConec.CargaRecordSet(sSql)
        If R.EOF Then
            VistoLavDineroEsCorrecto = False
        Else
            psVisPersCod = R!cPersCod
            VistoLavDineroEsCorrecto = True
        End If
        
        Set R = Nothing
        oConec.CierraConexion
    Set oConec = Nothing
End Function

'**DAOR 20071113
'**procedimiento que carga los controles utilizados en el mantenimiento
'**de responsabilidades de permisos
Public Sub CargaControlesMantResponsabilidades(ByVal psDominio As String, _
                        ByRef pMatUser As Variant, _
                        ByRef pMatGrupo As Variant, _
                        ByRef prsMenu As ADODB.Recordset, _
                        ByRef prsOpe As ADODB.Recordset, Optional psAplicacion As String = "1")
    
Dim sCadTemp As String
Dim i As Integer
            
    Call CargaControlUsuariosParaMantPermisos(psDominio)
    
    ReDim pMatUser(nNumUsu, 2)
    
    sCadTemp = DameUsuario
    i = 0
    
    While sCadTemp <> ""
        pMatUser(i, 0) = sCadTemp
        pMatUser(i, 1) = DameUsuarioNameFull
        sCadTemp = DameUsuario
        i = i + 1
    Wend
    
    Call CargaControlGruposParaMantPermisos(psDominio)
    sCadTemp = DameGrupo
    
    i = 0
    While sCadTemp <> ""
        ReDim Preserve pMatGrupo(i + 1)
        pMatGrupo(i) = sCadTemp
        sCadTemp = DameGrupo
        i = i + 1
    Wend
    
    Set prsMenu = DameItemsMenu(, psAplicacion)
    Set prsOpe = DameRSOperaciones
    
End Sub

'**DAOR 20071122
'**Carga la lista de usuarios responsables del mantenimiento de permisos
Public Sub CargaControlUsuariosParaMantPermisos(ByVal psDominio As String)
Dim i As Integer
Dim container As IADsContainer
Dim containername As String
Dim User As IADsUser
Dim Conn As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset
Dim sCriterio As String

    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set Conn = Nothing
        Exit Sub
    End If
    sSql = "exec stp_sel_GruposUsuResponsableDePermisos '1' "
    Set R = Conn.CargaRecordSet(sSql, adLockReadOnly)
    Set Conn = Nothing
            
    Set container = GetObject("WinNT://" & psDominio)
    container.Filter = Array("User")
    i = 0
    nNumUsu = 0
    ReDim LstUsuarios(i)
    ReDim LstUsuariosNameFull(i)
    For Each User In container
        If Len(User.Name) = 4 Then
            sCriterio = "cGrupoUsu = '" & User.Name & "'"
            If R.RecordCount > 0 Then
                R.MoveFirst
                R.Find sCriterio, , , 1
            End If
            If Not R.EOF Then
                ReDim Preserve LstUsuarios(i + 1)
                ReDim Preserve LstUsuariosNameFull(i + 1)
                LstUsuarios(i) = User.Name
                LstUsuariosNameFull(i) = User.FullName
                i = i + 1
                nNumUsu = nNumUsu + 1
            End If
        End If
    Next
    Set R = Nothing
End Sub

'**DAOR 20071122
'**Carga la lista de grupos responsables del mantenimiento de permisos
Public Sub CargaControlGruposParaMantPermisos(ByVal psDominio As String)
Dim container As IADsContainer
Dim containername As String
Dim group As IADsGroup
Dim Conn As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset
Dim sCriterio As String

    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set Conn = Nothing
        Exit Sub
    End If
    sSql = "exec stp_sel_GruposUsuResponsableDePermisos '2' "
    Set R = Conn.CargaRecordSet(sSql, adLockReadOnly)
    Set Conn = Nothing
    
    Set container = GetObject("WinNT://" & psDominio)
    container.Filter = Array("Group")
    nNumGrupos = 0
    ReDim LstUsuarios(nNumGrupos)
    For Each group In container
        If UCase(Left(group.Name, 5)) = "GRUPO" Or UCase(Left(group.Name, 3)) = "GG_" Then
            sCriterio = "cGrupoUsu = '" & group.Name & "'"
            If R.RecordCount > 0 Then
                R.MoveFirst
                R.Find sCriterio, , , 1
            End If
            If Not R.EOF Then
                ReDim Preserve LstUsuarios(nNumGrupos + 1)
                LstUsuarios(nNumGrupos) = group.Name
                nNumGrupos = nNumGrupos + 1
            End If
        End If
    Next
    nPosUsuGrp = 0
    Set R = Nothing
End Sub


'**DAOR 20071122
'**Carga las opciones que puede dar Permisos un usuario o Grupo
Public Function CargaOpcionesResponsable(ByVal psDominio As String, Optional ByVal psUsuario As String = "", Optional ByVal sTipoUsu As String = "1")
Dim sSql As String
Dim oConn As COMConecta.DCOMConecta
Dim RMenu As ADODB.Recordset
Dim RPermisoUsu As ADODB.Recordset
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
    
    Set oConn = New COMConecta.DCOMConecta
    If oConn.AbreConexion() Then
        If psUsuario <> "" And sTipoUsu = "1" Then
            If sCadGruposdeUsu = "" Then
                sCadGruposdeUsu = "'" & psUsuario & "'"
            Else
                sCadGruposdeUsu = sCadGruposdeUsu & ",'" & psUsuario & "' "
            End If
            sCadGruposdeUsu = Replace(sCadGruposdeUsu, "'", "''")
            sSql = " exec stp_sel_PermisoResponsabilidadPorUsuario '" & sCadGruposdeUsu & "' "
            Set RMenu = oConn.CargaRecordSet(sSql)
            oConn.CierraConexion
            sCadMenu = ""
            sCadMenuGrp = ""
            
            Do While Not RMenu.EOF
                sCadMenu = sCadMenu & "*" & Trim(RMenu!cOpcion) & ","
                If UCase(Trim(RMenu!cGrupoUsu)) <> UCase(psUsuario) Then
                    sCadMenuGrp = sCadMenuGrp & "*" & Trim(RMenu!cOpcion) & ","
                End If
                Y = Y + 1
                RMenu.MoveNext
            Loop
            RMenu.Close
            Set RMenu = Nothing
        Else
            sSql = "exec stp_sel_MenuOperaciones 0 "
            Set RMenu = oConn.CargaRecordSet(sSql)
            'responsabilidades de permisos por usuario
            sSql = " exec stp_sel_PermisoResponsabilidadPorUsuario '''" & UCase(sUsuario) & "'''"
            Set RPermisoUsu = oConn.CargaRecordSet(sSql)
            oConn.CierraConexion
            
            sCadMenu = ""
            sCadMenuGrp = ""
            Y = 0
            Do While Not RMenu.EOF
                sCriterio = "cName = '" & Trim(RMenu!cname) & "'"
                If RPermisoUsu.RecordCount > 0 Then
                    RPermisoUsu.MoveFirst
                    RPermisoUsu.Find sCriterio, , , 1
                End If
                If Not RPermisoUsu.EOF Then
                    sCadMenu = sCadMenu & Trim(RPermisoUsu!cOpcion) & ","
                    Y = Y + 1
                End If
                RMenu.MoveNext
            Loop
            RMenu.Close
            
            Set RMenu = Nothing
                RPermisoUsu.Close
            Set RPermisoUsu = Nothing
        End If
    End If
    Set oConn = Nothing
End Function


'**DAOR 20071122
'**Otorgar responsabilidad de menu u operacion
Public Sub OtorgarResponsabilidadDePermiso(ByVal psUsuGrp As String, ByVal psItemMenuName As String, ByVal psMenuOpe As String)
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
    vError = False
    sSql = "exec stp_ins_PermisoResponsabilidad '" & psUsuGrp & "','" & psItemMenuName & "','" & psMenuOpe & "'"
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar sSql
    Conn.CierraConexion
    Set Conn = Nothing
End Sub

'**DAOR 20071122
'**Denegar responsabilidad de menu u operacion
Public Sub DenegarResponsabilidadDePermiso(ByVal psUsuGrp As String, ByVal psItemMenuName As String)
Dim sSql As String
Dim Conn As COMConecta.DCOMConecta
    vError = False
    sSql = "exec stp_del_PermisoResponsabilidad '" & psUsuGrp & "','" & psItemMenuName & "'"
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar sSql
    Conn.CierraConexion
    Set Conn = Nothing
End Sub

'*** PEAC 20080812, Función que valida el visto Electrónico
'*** para cualquier proceso

Public Function VistoElectronicoEsCorrecto(psUsuario As String, psClave As String, ByRef psVisPersCod As String, Optional psCasoCod, Optional ByRef psPersNombreValidador As String, Optional ByRef psPersCargoValidador As String, Optional ByRef psMensaje As String, Optional ByRef pcUsuarioVisto As String) As Boolean
Dim oConec As COMConecta.DCOMConecta
Dim sSql As String
Dim oITF As COMDConstSistema.FCOMITF
Dim R As ADODB.Recordset
Dim R2 As ADODB.Recordset
Dim i As Integer
Dim lsCodcargos As String
Dim psDominio As String
Dim bClaveEsCorrecta As Boolean


    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    'Carga Dominio
    sSql = "SELECT nConsSisValor FROM ConstSistema WHERE nConsSisCod=37"
    Set R = oConec.CargaRecordSet(sSql)
    psDominio = Trim(R!nConsSisValor)
    
    bClaveEsCorrecta = ClaveIncorrectaNT(psUsuario, psClave, psDominio)
    VistoElectronicoEsCorrecto = bClaveEsCorrecta
    
    'Add by GITU 22-08-2008
    'sSql = "Select cCargoCod from VistoPorCargoyCaso where cCodCaso = '" & psCasoCod & "'"
    'Set R2 = oConec.CargaRecordSet(sSql)
    'i = 0
    'If Not R2.EOF Then
    '    Do While Not R2.EOF
    '        lsCodcargos = lsCodcargos & R2!cCargoCod
    '        i = i + 1
    '        R2.MoveNext
    '    Loop
    'End If
    'End Gitu
        
    psMensaje = "Usuario y clave incorrecta o usted no tiene permisos para dar el visto bueno"
    
    
    If Not bClaveEsCorrecta Then Exit Function

    Select Case psCasoCod
        Case 0
            
        'Case 1, 2, 3, 5, 6, 7, 9 '** Juez 20120829 Caso 6'** EJVG20131115 Se agregó caso 7 x Visto Elect. en Provisión
        'Case 1, 2, 3, 5, 6, 7, 9, 10 'RECO20140408 ERS160-2013
        'Case 1, 2, 3, 5, 6, 7, 9, 10, 11 'FRHU20140602 ERS063-2014
        'Case 1, 2, 3, 5, 6, 7, 9, 10, 11, 12 'FRHU20141203 ERS048-2014 Se agrego caso 12 x visto elect. notas de cargo y abono
        'Case 1, 2, 3, 5, 6, 7, 9, 10, 11, 12, 13 'RIRO20150605 ERS162-2014, ADD "13" x visto electronico de pago de utilidades a ex trabajadores.
        'Case 1, 2, 3, 5, 6, 7, 9, 10, 11, 12, 13, 14 'FRHU20151130 ERS077-2015  ADD "14" x visto Elect. para Actualizacion y Autorización de Datos
        'Case 1, 2, 3, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15 'PASI20151221 ERS0692015  ADD "15" x visto Elect. para Arqueo de Stock de Tarjetas de Debito
        'Case 1, 2, 3, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 16 'GITU20161004 ADD 16 X visto para cancelacion de CTS
        'Case 1, 2, 3, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17 'RECO20161014 ADD 17 X visto Anulacion Seguro Sepelio
        'Case 1, 2, 3, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18 'GIPO ERS070-2016 ADD 18 X visto Devolucion Sobrantes Adjudicados
        'Case 1, 2, 3, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20 'ANDE ERS021-2017 arqueo de expedientes de ahorro y extorno de arqueo de exp. de ahorro
        'Case 1, 2, 3, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21 'ANDE ERS021-2017 arqueo de expedientes de ahorro y extorno de arqueo de exp. de ahorro, ADD RIRO20170823
        'Case 1, 2, 3, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22 'ANDE ERS021-2017 arqueo de expedientes de ahorro y extorno de arqueo de exp. de ahorro, ADD 22 APRI20180802 ERS028-2017
        Case 1, 2, 3, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23 'JOEP20210910 Campana Prendario retencion 23
        
               sSql = " exec stp_sel_ObtieneDatosPersonaVisto '" & psUsuario & "'," & psCasoCod
               Set R = oConec.CargaRecordSet(sSql)
                'If R.BOF And R.EOF Then
                '    VistoElectronicoEsCorrecto = False
                'Else
                '    psVisPersCod = R!cPersCod
                '    psPersNombreValidador = R!cNombre
                '    psPersCargoValidador = R!cCargo
                '    VistoElectronicoEsCorrecto = True
                'End If
                'Set R = Nothing
                
'        Case 2
'              sSql = " exec stp_sel_ObtieneDatosPersonaVisto '" & psUsuario & "'," & psCasoCod
'              Set R = oConec.CargaRecordSet(sSql)
    End Select

    If R.BOF And R.EOF Then
        VistoElectronicoEsCorrecto = False
    Else
        psVisPersCod = R!cPersCod
        psPersNombreValidador = R!cNombre
        psPersCargoValidador = R!cCargo
        VistoElectronicoEsCorrecto = True
        pcUsuarioVisto = psUsuario
    End If
    Set R = Nothing
                
    oConec.CierraConexion
    Set oConec = Nothing

End Function

'MADM 20100825
Public Function ValidarVistoBuenoxGrupoUser(ByVal sTipoVisto As Integer, ByVal psGrupo As String) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
Dim Conn As COMConecta.DCOMConecta
Dim prsValor As Integer

    Set Conn = New COMConecta.DCOMConecta
    sSql = "exec stp_sel_ValidaVistoxGrupoUser " & sTipoVisto & ",'" & psGrupo & "'"
    Conn.AbreConexion
   
    Set R = Conn.CargaRecordSet(sSql)
    
     If R.BOF And R.EOF Then
        ValidarVistoBuenoxGrupoUser = False
        Exit Function
    Else
        prsValor = R!nValor
        
        If prsValor > 0 Then
            ValidarVistoBuenoxGrupoUser = True
            Exit Function
        End If
    End If
    
    Set R = Nothing
    Conn.CierraConexion
    Set Conn = Nothing
End Function
'modificado x madm 20100909 ---------------------
Public Function VistoElectronicoUser(ByVal psUsuario As String, ByVal psClave As String, psVisPersCod As String, Optional ByVal psCasoCod As Integer, Optional psPersNombreValidador As String, Optional psPersCargoValidador As String, Optional ByVal psMensaje As String) As Boolean
Dim oConec As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    
     If (psCasoCod = 4) Then
     sSql = " exec stp_sel_ObtieneDatosPersonaVistoxOpera '" & psUsuario & "', " & psCasoCod & ""
     Else
     sSql = " exec stp_sel_ObtieneDatosPersonaVisto '" & psUsuario & "', " & psCasoCod & ""
     End If
     Set R = oConec.CargaRecordSet(sSql)
    
    If R.BOF And R.EOF Then
        VistoElectronicoUser = False
    Else
        psVisPersCod = R!cPersCod
        psPersNombreValidador = R!cNombre
        psPersCargoValidador = R!cCargo
        VistoElectronicoUser = True
    End If
    Set R = Nothing
              
    oConec.CierraConexion
    Set oConec = Nothing

End Function
'end madm ------------------------
'***Agregado por ELRO el 20121112, según OYP-RFC115-2012
Public Function cargarUsuarioGrupoAprobacionRechazo(ByVal sUsuario As String, ByVal psDominio As String) As String
Dim i As Integer
Dim container As IADsContainer
Dim containername As String
Dim group As IADsGroup
Dim User As IADsUser
Dim sCadGruposdeUsu As String

    Set container = GetObject("WinNT://" & psDominio)
    Set User = GetObject("WinNT://" & psDominio & "/" & sUsuario & ",user")
    container.Filter = Array("Group")
    sCadGruposdeUsu = ""
    For Each group In User.Groups
        If UCase(Left(group.Name, 5)) = "GRUPO" Or UCase(Left(group.Name, 3)) = "GG_" Then
            If sCadGruposdeUsu = "" Then
                sCadGruposdeUsu = group.Name
            Else
                sCadGruposdeUsu = sCadGruposdeUsu & "," & group.Name
            End If
        End If
    Next
    cargarUsuarioGrupoAprobacionRechazo = sCadGruposdeUsu
End Function
'***Fin Agregado por ELRO el 20121112*******************

' ******* Agregado por RIRO el 20121109 ******************
Public Function listGrupos(ByVal sDominio As String) As String
           
    Dim container As IADsContainer
    Dim group As IADsGroup
    Dim Grupos As String
    
    On Error GoTo error
    
    Set container = GetObject("WinNT://" & sDominio)
    container.Filter = Array("Group")
    For Each group In container
        If UCase(Left(group.Name, 5)) = "GRUPO" Or UCase(Left(group.Name, 3)) = "GG_" Then
            Grupos = Grupos & group.Name & ","
        End If
    Next

    Grupos = Mid(Grupos, 1, Len(Grupos) - 1)
    listGrupos = Grupos
    
    Exit Function
    
error:
    Err.Raise Err.Number, Err.Source, "Error producido al cargar la lista de grupos del dominio, contactarse con el area de TI"
    
End Function
' ******************* Fin RIRO **********************

'WIOR 20130823 ***********************************************************
Public Function ObtenerMensajeSeguridad(Optional ByVal pbEstado As Boolean = False, Optional pbRandom As Boolean = False) As ADODB.Recordset
Dim Conn As New COMConecta.DCOMConecta
Dim sSql As String

Conn.AbreConexion
If pbRandom Then
    sSql = "EXEC stp_sel_SegMensajesRandom"
Else
    sSql = "EXEC stp_sel_SegMensajes " & IIf(pbEstado, 1, 0)
End If

Set ObtenerMensajeSeguridad = Conn.CargaRecordSet(sSql, adLockReadOnly)
Conn.CierraConexion
Set Conn = Nothing
End Function

Public Sub OpeMensajeSeguridad(ByVal pnTipo As Integer, ByVal psMensaje As String, ByVal pbEstado As Boolean, Optional ByVal pnId As Integer = 0)
Dim Conn As New COMConecta.DCOMConecta
Dim sSql As String

Conn.AbreConexion

If pnTipo = 1 Then
    sSql = "EXEC stp_upd_SegMensajes " & pnId & ",'" & psMensaje & "'," & IIf(pbEstado, 1, 0)
ElseIf pnTipo = 2 Then
    sSql = "EXEC stp_del_SegMensajes " & pnId
End If

Conn.Ejecutar sSql
Conn.CierraConexion
Set Conn = Nothing
End Sub

Public Function InsertaMensajeSeguridad(ByVal psMensaje As String, ByVal pbEstado As Boolean) As Long
Dim Conn As New COMConecta.DCOMConecta
Dim rsDatos As ADODB.Recordset
Dim sSql As String


sSql = "EXEC stp_ins_SegMensajes '" & psMensaje & "'," & IIf(pbEstado, 1, 0)
Conn.AbreConexion
Set rsDatos = Conn.CargaRecordSet(sSql)
Conn.CierraConexion

If Not (rsDatos.BOF And rsDatos.EOF) Then
    InsertaMensajeSeguridad = CLng(rsDatos!IdMensaje)
Else
    InsertaMensajeSeguridad = 0
End If

Set Conn = Nothing
Set rsDatos = Nothing
End Function
'WIOR FIN ****************************************************************
'RECO20140326 ERS174-2013 ANEXO 01****************************************
Public Function RecuperaPersCargo(ByVal psRHCargoCod As String) As ADODB.Recordset
Dim Conn As New COMConecta.DCOMConecta
Dim sSql As String

Conn.AbreConexion
sSql = "EXEC stp_sel_RecuperaPersCargo '" & psRHCargoCod & "'"
Set RecuperaPersCargo = Conn.CargaRecordSet(sSql, adLockReadOnly)
Conn.CierraConexion
Set Conn = Nothing
End Function
Public Function RecuperaDatosPersRRHH(ByVal psRHCargoCod As String) As ADODB.Recordset
Dim Conn As New COMConecta.DCOMConecta
Dim sSql As String

Conn.AbreConexion
sSql = "EXEC stp_sel_RecuperaDatosPersRRHH '" & psRHCargoCod & "'"
Set RecuperaDatosPersRRHH = Conn.CargaRecordSet(sSql, adLockReadOnly)
Conn.CierraConexion
Set Conn = Nothing
End Function
'RECO FIN*****************************************************************
'RECO20150326 ERS001-2015********************************************
Public Function ObtieneDatosPersonaRRHH(ByVal psRHUser As String, ByVal psRHPersCod As String) As ADODB.Recordset
Dim Conn As New COMConecta.DCOMConecta
Dim sSql As String

Conn.AbreConexion
sSql = "EXEC stp_sel_ObtieneDatosPersonaRRHH '" & psRHUser & "','" & psRHPersCod & "'"
Set ObtieneDatosPersonaRRHH = Conn.CargaRecordSet(sSql, adLockReadOnly)
Conn.CierraConexion
Set Conn = Nothing
End Function
'RECO FIN ***********************************************************

'GIPO ERS051-2016********************************************
Public Function obtenerPermisoAccesoArqueo(ByVal userArqueado As String, ByVal userArqueador As String) As ADODB.Recordset
Dim Conn As New COMConecta.DCOMConecta
Dim sSql As String

Conn.AbreConexion
sSql = "EXEC sp_sel_ERS0512016_permitirAccesoArqueo '" & userArqueado & "','" & userArqueador & "'"
Set obtenerPermisoAccesoArqueo = Conn.CargaRecordSet(sSql, adLockReadOnly)
Conn.CierraConexion
Set Conn = Nothing
End Function

'ANDE 20170627 Comprobrando si hay actualización
Public Function ComprobarActualizacion() As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    Dim R As ADODB.Recordset
    
    cSql = "stp_sel_ComprobarActualizacion '1'"
    On Error GoTo ErrorComprobarAct
        oCon.AbreConexion
        Set R = oCon.CargaRecordSet(cSql)
        oCon.CierraConexion
        
        Set ComprobarActualizacion = R
    Exit Function
ErrorComprobarAct:
    Err.Raise Err.Number, "Error al comprobar actualización.", Err.Description
End Function
'end ande

'PTI1 2019 03 2019 ACTA N° 034-2019 ******************************
Public Function GruposUsuario(ByVal sUsuario As String, ByVal psDominio As String) As String
Dim i As Integer
Dim container As IADsContainer
Dim containername As String
Dim group As IADsGroup
Dim User As IADsUser


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
    
    sCadGruposdeUsu = IIf(sCadGruposdeUsu = "'", "", sCadGruposdeUsu)
    
    If sCadGruposdeUsu <> "" And Len(sCadGruposdeUsu) > 1 Then
        sCadGruposdeUsu = Mid(sCadGruposdeUsu, 1, Len(sCadGruposdeUsu) - 2)
    End If
    nNumGrupos = 0
    GruposUsuario = sCadGruposdeUsu
End Function
'FIN PTI1 ********************************************
