VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMUAcceso"
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

Public Function DameMaquinasdeUsuario(ByVal psUser As String, ByVal psDominio As String) As Variant
Dim User As IADsUser
Dim m As Variant
    On Error Resume Next
    Set User = GetObject("WinNT://" & psDominio & "/" & psUser & ",user")
    m = User.LoginWorkstations
    DameMaquinasdeUsuario = m
End Function

Public Function IniciarNuevoDia() As Boolean
Dim ssql As String
Dim oConec As COMConecta.DCOMConecta  'COMConecta.DCOMConecta
Dim R As ADODB.Recordset
Dim dFecCieSis As Date
Dim dFecIniDia As Date

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    'Obtener fecha de Cierre del Sistema
    ssql = "Select nConsSisValor from ConstSistema Where nConsSisCod = 13"
    Set R = oConec.CargaRecordSet(ssql)
    dFecCieSis = CDate(R!nConsSisValor)
    'Obtener fecha de Inicio de Dia
    ssql = "Select nConsSisValor from ConstSistema Where nConsSisCod = 15"
    Set R = oConec.CargaRecordSet(ssql)
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
Dim Conn As COMConecta.DCOMConecta
Dim oFun As New COMFunciones.FCOMVarPublicas
Dim ssql As String
Dim R As ADODB.Recordset
Dim i As Integer
Dim Y As Integer
Dim lnFiltroM As Integer
Dim gsMenuApli As Integer
gsMenuApli = oFun.gsMenuAplicac
        Set Conn = New COMConecta.DCOMConecta
        If Not Conn.AbreConexion() Then
            vError = True
            sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
            Set Conn = Nothing
            Exit Function
        End If
        
        ssql = "SELECT nConsSisValor FROM CONSTSISTEMA WHERE nConsSisCod = 101"
        Set R = Conn.CargaRecordSet(ssql, adLockReadOnly)
                    
        lnFiltroM = R!nConsSisValor
        Set R = Nothing
        
        'FILTRO PARA LIMA (PIGNORATICIO)
        If lnFiltroM = 2 Then       'Muestra Pigno Lima
            ssql = "Select right(cName,12) as cCodigo,cDescrip,cName from Menu Where cAplicacion = '" & gsMenuApli & _
                "' AND cName <> 'M030100000002' AND Substring(cName,1,7) <> 'M030103' Order By nOrden"
        ElseIf lnFiltroM = 1 Then   'Muestra Pigno Trujillo
            ssql = "Select right(cName,12) as cCodigo,cDescrip,cName from Menu Where cAplicacion = '" & gsMenuApli & _
                "' AND cName <> 'M030100000003' AND Substring(cName,1,7) <> 'M030104' Order By nOrden"
        Else                        'Muestra Ambos Lima y Trujillo
            ssql = "Select right(cName,12) as cCodigo,cDescrip,cName from Menu Where cAplicacion = '" & gsMenuApli & "' Order By nOrden"
        End If
        
        Set R = Conn.CargaRecordSet(ssql, adLockReadOnly)
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
Dim rsVar As ADODB.Recordset
Dim Conn As COMConecta.DCOMConecta
    Dim ssql As String
    Dim lsFiltroMon As String
    
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set Conn = Nothing
        Exit Function
    End If
    
    ssql = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O Where O.cOpeVisible = '1'"
    ssql = ssql & " Order by O.cOpeCod, O.nOpeNiv"
    Set rsVar = Conn.CargaRecordSet(ssql, adLockReadOnly)
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
Dim ssql As String
Dim Conn As COMConecta.DCOMConecta
    vError = False
    ssql = "DELETE Permiso where cGrupoUsu = '" & psUsuGrp & "' And cName = '" & psItemMenuName & "' And cTipo = '" & psTipoUsu & "' And cMenuOpe = '1'"
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar ssql
    Conn.CierraConexion
    Set Conn = Nothing
End Sub
Public Sub OtorgarAcceso(ByVal psUsuGrp As String, ByVal psItemMenuName As String, ByVal psTipoUsu As String)
Dim ssql As String
Dim Conn As COMConecta.DCOMConecta
Dim oFun As New COMFunciones.FCOMVarPublicas
    vError = False
    ssql = "INSERT INTO PERMISO(cName,cGrupoUsu,cTipo,cMenuOpe,cAplicacion) VALUES('" & psItemMenuName & "','" & psUsuGrp & "','" & psTipoUsu & "','1'," & Str(oFun.gsMenuAplicac) & ")"
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar ssql
    Conn.CierraConexion
    Set Conn = Nothing
End Sub

Public Sub DenegarOperacion(ByVal psUsuGrp As String, ByVal psItemMenuName As String, ByVal psTipoUsu As String)
Dim ssql As String
Dim Conn As COMConecta.DCOMConecta
    vError = False
    ssql = "DELETE Permiso where cGrupoUsu = '" & psUsuGrp & "' And cName = '" & psItemMenuName & "' And cTipo = '" & psTipoUsu & "' And cMenuOpe = '2'"
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar ssql
    Conn.CierraConexion
    Set Conn = Nothing
End Sub
Public Sub OtorgarOperacion(ByVal psUsuGrp As String, ByVal psItemMenuName As String, ByVal psTipoUsu As String)
Dim ssql As String
Dim Conn As COMConecta.DCOMConecta
Dim oFun As New COMFunciones.FCOMVarPublicas
vError = False
    ssql = "INSERT INTO PERMISO(cName,cGrupoUsu,cTipo,cMenuOpe,cAplicacion) VALUES('" & psItemMenuName & "','" & psUsuGrp & "','" & psTipoUsu & "','2','" & oFun.gsMenuAplicac & "')"
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar ssql
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
Dim ssql As String
Dim Conn As COMConecta.DCOMConecta
Dim oFun As New COMFunciones.FCOMVarPublicas
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
    ssql = "DELETE Menu Where cAplicacion = '" & Trim(Str(oFun.gsMenuAplicac)) & "' "
    Conn.ConexionActiva.Execute ssql
    
    'Eliminar los que ya no estan en el Menu
    ssql = "Delete Permiso where cName in (Select cName from Menu where cName not in (" & sTemp & ")) AND cAplicacion = '" & Trim(Str(oFun.gsMenuAplicac)) & "'"
    Conn.ConexionActiva.Execute ssql
    ssql = "Delete Menu where cName in (Select cName from Menu where cName not in (" & sTemp & ")) AND cAplicacion = '" & Trim(Str(oFun.gsMenuAplicac)) & "'"
    Conn.ConexionActiva.Execute ssql
    
    'Ingresa Nuevos Menus
    sCadMenu = ""
    ssql = "Select cName,cDescrip from Menu where cAplicacion = '" & Trim(oFun.gsMenuAplicac) & "'"
    Set R = Conn.CargaRecordSet(ssql, adLockOptimistic)
    Do While Not R.EOF
        sCadMenu = sCadMenu & R!cname
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
        
    For i = 0 To UBound(MatMenu) - 1
        If InStr(sCadMenu, MatMenu(i, 0) & Right("00" & MatMenu(i, 2), 2)) <= 0 Then
            ssql = "INSERT INTO MENU(CNAME,CDescrip,nOrden,cAplicacion) VALUES('" & MatMenu(i, 0) & Right("00" & MatMenu(i, 2), 2) & "','" & Left(MatMenu(i, 1), 45) & "'," & i + 1 & ",'" & oFun.gsMenuAplicac & "')"
            Conn.ConexionActiva.Execute ssql
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
Dim ssql As String
Dim oConn As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
    ssql = "Select cUser from RRHH Where cUser = '" & psUser & "'"
    Set oConn = New COMConecta.DCOMConecta
    oConn.AbreConexion
    Set R = oConn.CargaRecordSet(ssql)
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
Dim ssql As String
Dim oConn As COMConecta.DCOMConecta
Dim oFun As New COMFunciones.FCOMVarPublicas
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
            
            ssql = "Select * from permiso where ((cGrupoUsu in (" & sCadGruposdeUsu & ")) OR ( cGrupoUsu = '" & psUsuario & "')) " & " And cAplicacion = '" & oFun.gsMenuAplicac & "' "
            Set RMenu = oConn.CargaRecordSet(ssql)
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
        
            ssql = "Select * from Menu where cAplicacion = '" & oFun.gsMenuAplicac & "'"
            Set RMenu = oConn.CargaRecordSet(ssql)
            'permisos por usuario
            ssql = "Select * from Permiso where cTipo='" & sTipoUsu & "' and upper(rtrim(ltrim(cGrupoUsu))) = '" & UCase(sUsuario) & "' And cAplicacion = '" & oFun.gsMenuAplicac & "'"
            Set RPermisoUsu = oConn.CargaRecordSet(ssql)
            'Permisos por Grupo
            ssql = "Select * from Permiso where cTipo='" & vsTipoPermisoGrp & "' And cAplicacion = '" & oFun.gsMenuAplicac & "'"
            Set RPermisoGrp = oConn.CargaRecordSet(ssql)
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
            ssql = "Select cOpeCod as cName,cOpeDesc from OpeTpo where cOpeVisible ='1' Order by cOpeCod, cOpeDesc"
            Set RMenu = oConn.CargaRecordSet(ssql)
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
    Dim cAccess As New DCOMAcceso
    ClaveIncorrectaNT = cAccess.SSPValidateUser(psUsuario, psDominio, psClave)
End Function
Private Sub Class_Initialize()
    vsTipoPermisoUsu = "1"
    vsTipoPermisoGrp = "2"
    nPosUsuGrp = 0
    nNumMenus = 0
    Espaciado = 60
    ReDim MatMenuPermisos(0)
End Sub

Public Function Cargousu(ByVal NomUser As String) As ADODB.Recordset
 Dim sql As String
 Dim rs As ADODB.Recordset
 Dim oConecta As COMConecta.DCOMConecta
 sql = "   SELECT top 1 RC.CRHCARGOCOD FROM RRHH RH "
 sql = sql & "  INNER JOIN RHCARGOS RC ON RC.CPERSCOD=RH.CPERSCOD "
 sql = sql & " WHERE RH.CUSER='" & NomUser & "'"
 sql = sql & " order by drhcargoFecha desc "
  
 Set oConecta = New COMConecta.DCOMConecta
 Set rs = New ADODB.Recordset
    oConecta.AbreConexion
    rs.CursorLocation = adUseClient
    Set rs = oConecta.CargaRecordSet(sql)
    Set Cargousu = rs
End Function

Public Function EnviaMensaje(ByVal cOpecod As String, ByVal nMonto As Double, ByVal nmoneda As Integer, ByRef lsmensaje As String) As Boolean
Dim ssql As String, saux As String
Dim rs As ADODB.Recordset, oConecta As COMConecta.DCOMConecta
  Set oConecta = New COMConecta.DCOMConecta
  Set rs = New ADODB.Recordset
    If nmoneda = gMonedaNacional Then
        saux = " and (c.nmontofinsol)>=" & nMonto
    ElseIf nmoneda = gMonedaExtranjera Then
        saux = " and (c.nmontofindol)>=" & nMonto
    End If
    
EnviaMensaje = False

ssql = ssql & " Select rh.cuser,cl.cworkstation,replace(rhc.cpersnombre,'/', ' ') nombre from  rrhh rh "
ssql = ssql & " inner join   (Select Max(cast(convert(char(8),r.drhcargoFecha,112) as int)) as Fecha,r.cperscod,r.crhcargocod,p.cpersnombre from rhcargos r "
ssql = ssql & " inner join persona p on p.cperscod=r.cperscod "
ssql = ssql & " group by r.cperscod,r.crhcargocod,p.cpersnombre) rhc on rhc.cperscod=rh.cperscod "
ssql = ssql & " inner join capautorizacionrango c on c.crhcargocod=rhc.crhcargocod "
ssql = ssql & " inner join caplogeo cl on cl.cuser=rh.cuser "
ssql = ssql & " where  rh.nrhestado='201' and c.copecod='" & cOpecod & "'" & saux

  
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(ssql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    lsmensaje = ""
    If rs.State = 1 Then
        While Not rs.EOF
          ' Shell "Net send " & rs!cworkstation & " URGENTE: Tiene Operacion Pendiente !!!"
            lsmensaje = lsmensaje & rs!cworkstation & "URGENTE " & CStr(rs!NOMBRE) & ": Tiene Operaciones Pendientes !!!" & Chr$(10)
            rs.MoveNext
        Wend
        If rs.RecordCount > 0 Then
            EnviaMensaje = True
        Else
            EnviaMensaje = False
        End If
    End If
Exit Function
MensaError:
    Err.Raise vbObjectError + 100, "ErrorEmiteDocOpe", "Error al Eviar Mensaje"
End Function





