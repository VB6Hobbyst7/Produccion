VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UAcceso"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Para Cargar Usuarios
Private Type JoinLong
   x As Long
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
  x As Long
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
Private sCadGruposdeUsu As String 'EJVG 20111028
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

Public Function CuentaBloqueada(ByVal psUser As String, ByVal psDominio As String) As Boolean
Dim User As IADsUser

    Set User = GetObject("WinNT://" & psDominio & "/" & psUser & ",user")
    CuentaBloqueada = User.IsAccountLocked
    
End Function

Public Function DameItemsMenu() As ADODB.Recordset
Dim Conn As DConecta
Dim sSql As String
Dim R As ADODB.Recordset
Dim i As Integer
Dim y As Integer
        Set Conn = New DConecta
        If Not Conn.AbreConexion() Then
            vError = True
            sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
            Set Conn = Nothing
            Exit Function
        End If
        'sSQL = "Select right(cName,12) as cCodigo,cDescrip,cName from Menu Order By right(cName,12)"
        sSql = "Select right(cName,12) as cCodigo,cDescrip,cName from Menu Where cAplicacion = '" & gsMenuAplicac & "' Order By nOrden"
        
        Set R = Conn.CargaRecordSet(sSql, adLockReadOnly)
        Set DameItemsMenu = R
       If R.EOF Then
            Exit Function
        End If
        nNumMenus = R.RecordCount
        i = 0
        sCadMenu = ""
        sCadMenuSql = "('"
        MenuItems = ""
        
        'EN VES DE CADENA PASAR A MATRIZ PARA PODER EVITAR EL DESBORDE DE MEMORIA
        sCadMenu = sCadMenu & Trim(R!cname)
End Function

Public Function DameRSOperaciones() As ADODB.Recordset
Dim rsVar As Recordset
Dim Conn As DConecta
    Dim sSql As String
    Dim lsFiltroMon As String
    
    Set Conn = New DConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set Conn = Nothing
        Exit Function
    End If
    
    sSql = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O where cOpeVisible ='1'"
    sSql = sSql & " Order by O.cOpeCod, O.nOpeNiv"
    Set rsVar = Conn.CargaRecordSet(sSql, adLockReadOnly)
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
Dim sSql As String
Dim Conn As DConecta
    vError = False
    sSql = "DELETE Permiso where cGrupoUsu = '" & psUsuGrp & "' And cName = '" & psItemMenuName & "' And cTipo = '" & psTipoUsu & "' And cMenuOpe = '1'"
    Set Conn = New DConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar sSql
    Conn.CierraConexion
    Set Conn = Nothing
End Sub
Public Sub OtorgarAcceso(ByVal psUsuGrp As String, ByVal psItemMenuName As String, ByVal psTipoUsu As String)
Dim sSql As String
Dim Conn As DConecta
    vError = False
    sSql = "INSERT INTO PERMISO(cName,cGrupoUsu,cTipo,cMenuOpe,cAplicacion) VALUES('" & psItemMenuName & "','" & psUsuGrp & "','" & psTipoUsu & "','1'," & Str(gsMenuAplicac) & ")"
    Set Conn = New DConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar sSql
    Conn.CierraConexion
    Set Conn = Nothing
End Sub

Public Sub DenegarOperacion(ByVal psUsuGrp As String, ByVal psItemMenuName As String, ByVal psTipoUsu As String)
Dim sSql As String
Dim Conn As DConecta
    vError = False
    sSql = "DELETE Permiso where cGrupoUsu = '" & psUsuGrp & "' And cName = '" & psItemMenuName & "' And cTipo = '" & psTipoUsu & "' And cMenuOpe = '2'"
    Set Conn = New DConecta
    If Not Conn.AbreConexion() Then
        vError = True
        sMsgError = "No se pudo Conectar al Servidor"
        Exit Sub
    End If
    Conn.Ejecutar sSql
    Conn.CierraConexion
    Set Conn = Nothing
End Sub
Public Sub OtorgarOperacion(ByVal psUsuGrp As String, ByVal psItemMenuName As String, ByVal psTipoUsu As String, Optional ByVal pdFecha As Date = "1900/01/01") 'WIOR 20130201 AGREGO pdFecha
Dim sSql As String
Dim Conn As DConecta
    vError = False
    'sSql = "INSERT INTO PERMISO(cName,cGrupoUsu,cTipo,cMenuOpe,cAplicacion) VALUES('" & psItemMenuName & "','" & psUsuGrp & "','" & psTipoUsu & "','2','" & gsMenuAplicac & "')"
    sSql = "exec stp_ins_Permiso '" & psItemMenuName & "','" & psUsuGrp & "','" & psTipoUsu & "','2','" & gsMenuAplicac & "','" & Format(pdFecha, "mm/dd/yyyy") & "'" 'WIOR 20130201 AGREGO pdFecha

    Set Conn = New DConecta
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
    Dim TempPtr As JoinLong, TempStr As JoinInt, x&, pdc$
    Let x = GetPDC(psPDC)
    pdc = psPDC
    SNArray = pdc & vbNullChar
    UNArray = User & vbNullChar
    Result = NetUserGetInfo(SNArray(0), UNArray(0), 3, bufptr)
    'DoEvents
    If Result = 0 Then
        Result = PtrToInt(TempStr.Bottom, bufptr + 36, 2)
        Result = PtrToInt(TempStr.Top, bufptr + 38, 2)
        LSet TempPtr = TempStr
        Result = PtrToStr(StrArray(0), TempPtr.x)
        UserName = Left(StrArray, StrLen(TempPtr.x))
        Result = PtrToInt(TempStr.Bottom, bufptr + 52, 2)
        Result = PtrToInt(TempStr.Top, bufptr + 54, 2)
        LSet TempPtr = TempStr
        LOn = TempPtr.x
        Result = PtrToInt(TempStr.Bottom, bufptr + 56, 2)
        Result = PtrToInt(TempStr.Top, bufptr + 58, 2)
        LSet TempPtr = TempStr
        LOff = TempPtr.x
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
            APIResult = PtrToStr(StrArray(0), TempPtr.x)
            ReDim Preserve Users$(1 To Size)
            Users(Size) = Left(StrArray, StrLen(TempPtr.x))
            APIResult = PtrToInt(TempStr.Bottom, SortedBuffer + (i - 1) * 24 + 20, 2)
            APIResult = PtrToInt(TempStr.Top, SortedBuffer + (i - 1) * 24 + 22, 2)
            LSet TempPtr = TempStr
            Index = TempPtr.x
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
Dim sSql As String
Dim Conn As DConecta
Dim R As ADODB.Recordset
Dim sCriterio As String
Dim sCadMenu As String
Dim i As Integer
Dim j As Integer
Dim sTemp As String

    On Error GoTo ErrRegeneraMenu
    vError = False
    sMsgError = ""
    
    Set Conn = New DConecta
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
            
    'Eliminar los que ya no estan en el Menu
    sSql = "Delete Permiso where cName in (Select cName from Menu where cName not in (" & sTemp & ")) AND cAplicacion = '" & Str(gsMenuAplicac) & "'"
    Conn.ConexionActiva.Execute sSql
    'ALPA 20120209 ***
    'sSql = "Delete Menu where cName in (Select cName from Menu where cName not in (" & sTemp & ")) AND cAplicacion = '" & Str(gsMenuAplicac) & "'"
    sSql = "Delete Menu where cAplicacion = '3'"
    '***
    Conn.ConexionActiva.Execute sSql
    
    'Ingresa Nuevos Menus
    sCadMenu = ""
    sSql = "Select cName,cDescrip from Menu where cAplicacion = '" & LTrim(gsMenuAplicac) & "'"
    Set R = Conn.CargaRecordSet(sSql, adLockOptimistic)
    Do While Not R.EOF
        sCadMenu = sCadMenu & R!cname
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
        
    For i = 0 To UBound(MatMenu) - 1
        If InStr(sCadMenu, MatMenu(i, 0) & Right("00" & MatMenu(i, 2), 2)) <= 0 Then
            sSql = "INSERT INTO MENU(CNAME,CDescrip,nOrden,cAplicacion) VALUES('" & MatMenu(i, 0) & Right("00" & MatMenu(i, 2), 2) & "','" & Left(MatMenu(i, 1), 45) & "'," & i + 1 & ",'" & gsMenuAplicac & "')"
            'Conn.ConexionActiva.Execute sSql
            Conn.Ejecutar sSql
            
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

    buffUsu = Space(255)
    lSizeUsu = Len(buffUsu)
    Call GetUserName(buffUsu, lSizeUsu)
    ObtenerUsuario = Mid(Trim(buffUsu), 1, Len(Trim(buffUsu)) - 1)
       
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
    sCadGruposdeUsu = "'" 'EJVG 20111028
    ReDim sGrupoUsu(i)
    For Each group In User.Groups
        If UCase(Left(group.Name, 5)) = "GRUPO" Or UCase(Left(group.Name, 3)) = "GG_" Then
            ReDim Preserve sGrupoUsu(i + 1)
            sGrupoUsu(i) = group.Name
            sCadGruposdeUsu = sCadGruposdeUsu & sGrupoUsu(i) & "','" 'EJVG 20111028
            i = i + 1
        End If
    Next
    'EJVG 20111028*************************
    sCadGruposdeUsu = IIf(sCadGruposdeUsu = "'", "", sCadGruposdeUsu)
    
    If sCadGruposdeUsu <> "" And Len(sCadGruposdeUsu) > 1 Then
        sCadGruposdeUsu = Mid(sCadGruposdeUsu, 1, Len(sCadGruposdeUsu) - 2)
    End If
    '**************************************
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

'Public Function CargaMenu(ByVal psDominio As String, Optional ByVal psUsuario As String = "", Optional ByVal sTipoUsu As String = "1")
'Dim sSQL As String
'Dim oConn As DConecta
'Dim RMenu As ADODB.Recordset
'Dim RPermisoUsu As ADODB.Recordset
'Dim RPermisoGrp As ADODB.Recordset
'Dim sUsuario As String
'Dim sCriterio As String
'Dim I As Integer
'Dim Y As Integer
'
'    'Carga Grupos
'    If Len(Trim(psUsuario)) = 0 Then
'        sUsuario = UCase(ObtenerUsuario())
'    Else
'        sUsuario = UCase(psUsuario)
'    End If
'    If sTipoUsu = "1" Then 'Si es un usuario de lo contrario un grupo
'        Call CargaGruposUsuario(sUsuario, psDominio)
'    End If
'
'    Set oConn = New DConecta
'    If oConn.AbreConexion() Then
'        sSQL = "Select * from Menu where cAplicacion = '" & gsMenuAplicac & "'"
'        Set RMenu = oConn.CargaRecordSet(sSQL)
'        'permisos por usuario
'        sSQL = "Select * from Permiso where cTipo='" & sTipoUsu & "' and upper(rtrim(ltrim(cGrupoUsu))) = '" & UCase(sUsuario) & "' And cAplicacion = '" & gsMenuAplicac & "'"
'        Set RPermisoUsu = oConn.CargaRecordSet(sSQL)
'        'Permisos por Grupo
'        sSQL = "Select * from Permiso where cTipo='" & vsTipoPermisoGrp & "' And cAplicacion = '" & gsMenuAplicac & "'"
'        Set RPermisoGrp = oConn.CargaRecordSet(sSQL)
'        oConn.CierraConexion
'
'        sCadMenu = ""
'        sCadMenuGrp = ""
'        'For I = 0 To nNumGrupos - 1
'        '    If UCase(sGrupoUsu(I)) = "GG_SISTEMAS" Then
'        '        sCadMenu = "mnuSeguridad09000000mnuSegurPerm09010000"
'        '        sCadMenuGrp = "mnuSeguridad09000000mnuSegurPerm09010000"
'        '        Exit For
'        '    End If
'        'Next I
'        Y = 0
'        Do While Not RMenu.EOF
'            If RMenu!cname = "401402" Then
'                sCriterio = sCriterio
'            End If
'            'Buscando por usuario
'            sCriterio = "cName = '" & Trim(RMenu!cname) & "'"
'            If RPermisoUsu.RecordCount > 0 Then
'                RPermisoUsu.MoveFirst
'                RPermisoUsu.Find sCriterio, , , 1
'            End If
'            If Not RPermisoUsu.EOF Then
'                sCadMenu = sCadMenu & Trim(RPermisoUsu!cname)
'                Y = Y + 1
'                ReDim Preserve MatMenuPermisos(Y)
'                MatMenuPermisos(Y) = RPermisoUsu!cname
'            Else
'                'Buscando por Grupo
'                If sTipoUsu = "1" Then
'                    sCriterio = "cName = '" & Trim(RMenu!cname) & "'"
'                    RPermisoGrp.Filter = sCriterio
'                    If RPermisoGrp.RecordCount > 0 Then
'                        RPermisoGrp.MoveFirst
'                        Do While Not RPermisoGrp.EOF
'                            If PerteneceAGrupo(Trim(RPermisoGrp!cGrupoUsu)) Then
'                                sCadMenu = sCadMenu & Trim(RPermisoGrp!cname)
'                                sCadMenuGrp = sCadMenuGrp & Trim(RPermisoGrp!cname)
'                                Y = Y + 1
'                                ReDim Preserve MatMenuPermisos(Y)
'                                MatMenuPermisos(Y) = RPermisoGrp!cname
'                            End If
'                            RPermisoGrp.MoveNext
'                        Loop
'                    End If
'                    RPermisoGrp.Filter = adFilterNone
'                End If
'            End If
'            RMenu.MoveNext
'        Loop
'        RMenu.Close
'
'        '******************************************************************************
'        '**********************   Para Operaciones  ***********************************
'        '******************************************************************************
'         oConn.AbreConexion
'        sSQL = "Select cOpeCod as cName,cOpeDesc from OpeTpo where cOpeVisible = '1' Order by cOpeCod, cOpeDesc"
'        Set RMenu = oConn.CargaRecordSet(sSQL)
'        oConn.CierraConexion
'        Do While Not RMenu.EOF
'            'Buscando por usuario
'            sCriterio = "cName = '" & Trim(RMenu!cname) & "'"
'            If RPermisoUsu.RecordCount > 0 Then
'                RPermisoUsu.MoveFirst
'                RPermisoUsu.Find sCriterio, , , 1
'            End If
'            If Not RPermisoUsu.EOF Then
'                sCadMenu = sCadMenu & Trim(RPermisoUsu!cname)
'                Y = Y + 1
'                ReDim Preserve MatMenuPermisos(Y)
'                MatMenuPermisos(Y) = RPermisoUsu!cname
'            Else
'                'Buscando por Grupo
'                If sTipoUsu = "1" Then
'                    sCriterio = "cName = '" & Trim(RMenu!cname) & "'"
'                    RPermisoGrp.Filter = sCriterio
'                    If RPermisoGrp.RecordCount > 0 Then
'                        RPermisoGrp.MoveFirst
'                        Do While Not RPermisoGrp.EOF
'                            If PerteneceAGrupo(Trim(RPermisoGrp!cGrupoUsu)) Then
'                                sCadMenu = sCadMenu & Trim(RPermisoGrp!cname)
'                                sCadMenuGrp = sCadMenuGrp & Trim(RPermisoGrp!cname)
'                                Y = Y + 1
'                                ReDim Preserve MatMenuPermisos(Y)
'                                MatMenuPermisos(Y) = RPermisoGrp!cname
'                            End If
'                            RPermisoGrp.MoveNext
'                        Loop
'                    End If
'                    RPermisoGrp.Filter = adFilterNone
'                End If
'            End If
'            RMenu.MoveNext
'        Loop
'
'        RMenu.Close
'        Set RMenu = Nothing
'        RPermisoGrp.Close
'        Set RPermisoGrp = Nothing
'        RPermisoUsu.Close
'        Set RPermisoUsu = Nothing
'
'    End If
'    Set oConn = Nothing
'End Function

Public Function CargaMenu(ByVal psDominio As String, Optional ByVal psUsuario As String = "", Optional ByVal sTipoUsu As String = "1", _
                          Optional ByVal oConex As DConecta = Nothing, Optional ByRef pbTieneAlgunPermiso As Boolean = False, _
                          Optional ByRef psGruposdeUsu As String, Optional pgdFecSis As Date = "01/01/1900") 'Juez 20120715 Se agregó psGruposdeUsu'WIOR 20130201 AGREGO pgdFecsis
Dim sSql As String
Dim oConn As DConecta
Dim RMenu As ADODB.Recordset
Dim RPermisoUsu As ADODB.Recordset
Dim RPermisoGrp As ADODB.Recordset
Dim sUsuario As String
Dim sCriterio As String
Dim i As Integer
Dim y As Integer
Dim pCadGrupo As String 'WIOR 20130201

    'Carga Grupos
    If Len(Trim(psUsuario)) = 0 Then
        sUsuario = UCase(ObtenerUsuario())
    Else
        sUsuario = UCase(psUsuario)
    End If
    If sTipoUsu = "1" Then 'Si es un usuario de lo contrario un grupo
        Call CargaGruposUsuario(sUsuario, psDominio)
    End If
    
    'RIRO 20131108 RFIII **************
    Dim clsGen As New DGeneral
    Dim oPersona As New DPersonas
    Dim rsTemp As ADODB.Recordset
    Dim sRF3 As String, sTemporal  As String
    Set rsTemp = clsGen.GetConstante_New(10027, , "100", "1")
    sRF3 = IIf(Not rsTemp Is Nothing, IIf(Not rsTemp.EOF And Not rsTemp.BOF, rsTemp!cDescripcion, String(10, "-")), String(10, "-"))
    Set rsTemp = Nothing
    If InStr(1, sCadGruposdeUsu, sRF3) > 0 Then
        Set rsTemp = oPersona.RecuperarGruposRF3(sUsuario)
        If Not rsTemp Is Nothing Then
            If Not rsTemp.BOF And Not rsTemp.EOF Then
                If rsTemp!nEstado = 1 Then
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
    'END RIRO *************************
        
    If oConex Is Nothing Then
        Set oConn = New DConecta
        oConn.AbreConexion
    End If
        psGruposdeUsu = Replace(sCadGruposdeUsu, "'", "") 'Juez 20120715
        'EJVG 20111028*************************
        If sCadGruposdeUsu = "" Then
            sCadGruposdeUsu = "'" & psUsuario & "'"
        Else
            sCadGruposdeUsu = sCadGruposdeUsu & ",'" & psUsuario & "' "
            pCadGrupo = Replace(sCadGruposdeUsu, "'", "") 'WIOR 20130201
        End If
        'sSql = "Select * from permiso where cGrupoUsu in (" & sCadGruposdeUsu & ") And cAplicacion = '" & gsMenuAplicac & "'"
        sSql = " exec stp_sel_DevuelvePermisosGrupoUsuarioFecha '" & pCadGrupo & "','" & psUsuario & "','" & gsMenuAplicac & "','" & Format(pgdFecSis, "yyyymmdd") & "' " 'WIOR 20130201

        Set RMenu = oConn.CargaRecordSet(sSql)
        pbTieneAlgunPermiso = IIf(Not RMenu.EOF, True, False)
        
        If pbTieneAlgunPermiso = False Then
            Exit Function
        End If
        '**************************************
    
        
        'sSql = "Select * from Menu where cAplicacion = '" & gsMenuAplicac & "'"
        'If oConex Is Nothing Then
        '    Set RMenu = oConn.CargaRecordSet(sSql)
        'Else
        '    Set RMenu = oConex.CargaRecordSet(sSql)
        'End If
        
        'permisos por usuario
        sSql = "Select * from Permiso where cTipo='" & sTipoUsu & "' and upper(rtrim(ltrim(cGrupoUsu))) = '" & UCase(sUsuario) & "' And cAplicacion = '" & gsMenuAplicac & "'"
        If oConex Is Nothing Then
            Set RPermisoUsu = oConn.CargaRecordSet(sSql)
        Else
            Set RPermisoUsu = oConex.CargaRecordSet(sSql)
        End If
        
        'Permisos por Grupo
        sSql = "Select * from Permiso where cTipo='" & vsTipoPermisoGrp & "' And cAplicacion = '" & gsMenuAplicac & "'"
        If oConex Is Nothing Then
            Set RPermisoGrp = oConn.CargaRecordSet(sSql)
        Else
            Set RPermisoGrp = oConex.CargaRecordSet(sSql)
        End If
        
        
    '    oConn.CierraConexion
        
        sCadMenu = ""
        sCadMenuGrp = ""
        'For I = 0 To nNumGrupos - 1
        '    If UCase(sGrupoUsu(I)) = "GG_SISTEMAS" Then
        '        sCadMenu = "mnuSeguridad09000000mnuSegurPerm09010000"
        '        sCadMenuGrp = "mnuSeguridad09000000mnuSegurPerm09010000"
        '        Exit For
        '    End If
        'Next I
        y = 0
        Do While Not RMenu.EOF
            If RMenu!cname = "401402" Then
                sCriterio = sCriterio
            End If
            'Buscando por usuario
            sCriterio = "cName = '" & Trim(RMenu!cname) & "'"
            If RPermisoUsu.RecordCount > 0 Then
                RPermisoUsu.MoveFirst
                RPermisoUsu.Find sCriterio, , , 1
            End If
            If Not RPermisoUsu.EOF Then
                sCadMenu = sCadMenu & Trim(RPermisoUsu!cname) & ","
                y = y + 1
                ReDim Preserve MatMenuPermisos(y)
                MatMenuPermisos(y) = RPermisoUsu!cname
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
                                y = y + 1
                                ReDim Preserve MatMenuPermisos(y)
                                MatMenuPermisos(y) = RPermisoGrp!cname
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
     
     '    oConn.AbreConexion
        
        sSql = "Select cOpeCod as cName,cOpeDesc from OpeTpo where cOpeVisible = '1' Order by cOpeCod, cOpeDesc"
        
        If oConex Is Nothing Then
            Set RMenu = oConn.CargaRecordSet(sSql)
        Else
            Set RMenu = oConex.CargaRecordSet(sSql)
        End If
     
     '   oConn.CierraConexion
        Do While Not RMenu.EOF
            'Buscando por usuario
            sCriterio = "cName = '" & Trim(RMenu!cname) & "'"
            If RPermisoUsu.RecordCount > 0 Then
                RPermisoUsu.MoveFirst
                RPermisoUsu.Find sCriterio, , , 1
            End If
            If Not RPermisoUsu.EOF Then
                sCadMenu = sCadMenu & "*" & Trim(RPermisoUsu!cname) & ","
                y = y + 1
                ReDim Preserve MatMenuPermisos(y)
                MatMenuPermisos(y) = RPermisoUsu!cname
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
                                y = y + 1
                                ReDim Preserve MatMenuPermisos(y)
                                MatMenuPermisos(y) = RPermisoGrp!cname
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
        
    'End If
    If oConex Is Nothing Then
        Set oConn = Nothing
    End If
    
End Function


Public Function InterconexionCorrecta() As Boolean
Dim oConn As DConecta
    Set oConn = New DConecta
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
End Sub

'ARCV 11-09-2006
Public Sub CargarMenu_NEW(ByVal psUsuario As String, _
                        ByVal psClave As String, _
                        ByRef pMatMenuPer As Variant, _
                        ByRef prsOpeRep As ADODB.Recordset)
    
Dim sTitulo As String
Dim i As Integer
Dim y As Integer
Dim oConec As DConecta
Dim sSql As String
Dim oITF As COMDConstSistema.FCOMITF
Dim R As ADODB.Recordset
Dim psDominio As String
    
'Para las Operaciones F2 y de Extornos
Dim sFiltro() As String
Dim lnFiltraTC As Integer
Dim lnFiltraMP As Integer
Dim oGen As DGeneral
    
'Para el Manejo de la Matriz de Operaciones
Dim MatIndices() As Integer

'Para carga menu personalizado
Dim MatMenuPer() As Integer
Dim nNumMatMenuPer As Integer

Dim pMatOperac() As String
Dim pnNumOperac As Integer
Dim prsMenu As ADODB.Recordset


    
    'Carga Operaciones
    Set oConec = New DConecta
    oConec.AbreConexion
    'Carga Dominio
    sSql = "SELECT nConsSisValor FROM ConstSistema WHERE nConsSisCod=37"
    Set R = oConec.CargaRecordSet(sSql)
    psDominio = Trim(R!nConsSisValor)
    
    sSql = "Select cOpeCod,cOpeDesc,cOpeVisible,nOpeNiv,cOpeGruCod,cUltimaActualizacion from OpeTpo WHERE cOpeVisible ='1' Order by cOpeCod"
    Set R = oConec.CargaRecordSet(sSql)

    Set prsMenu = DameItemsMenu(oConec)
    
    Call CargaMenu(psDominio, psUsuario, , oConec)
    
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
    y = 0
    i = 0
    Do While Not R.EOF
        If TienePermiso(R!cOpeCod, "", True) Then
            y = y + 1
            ReDim Preserve MatIndices(i + 1)
            MatIndices(i) = R.Bookmark
            i = i + 1
        End If
        R.MoveNext
    Loop
    
    pnNumOperac = y
    ReDim pMatOperac(y, 5)
    
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
    Set oGen = New DGeneral
        
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

    Set prsOpeRep = oGen.GetOperacionesUsuarioMatriz(psUsuario, sFiltro, pMatOperac, pnNumOperac, , oConec)
    
    oConec.CierraConexion
    Set oConec = Nothing
    Set oGen = Nothing
End Sub


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
'WIOR 20130823 ***********************************************************
Public Function ObtenerMensajeSeguridad() As ADODB.Recordset
Dim Conn As New DConecta
Dim sSql As String

Conn.AbreConexion
sSql = "EXEC stp_sel_SegMensajesRandom"

Set ObtenerMensajeSeguridad = Conn.CargaRecordSet(sSql, adLockReadOnly)
Conn.CierraConexion
Set Conn = Nothing
End Function
Public Function VerificaExistePermisoCargo(ByVal psCodCargo As String, ByVal pnTipoOpc As Integer) As Boolean
Dim rsVar As New ADODB.Recordset
Dim Conn As New DConecta
Dim sSql As String

sSql = "stp_sel_ERS0862013_VerificaPermisoCargo '" & psCodCargo & "'," & pnTipoOpc

Conn.AbreConexion
Set rsVar = Conn.CargaRecordSet(sSql, adLockReadOnly)
Conn.CierraConexion

If (rsVar.EOF And rsVar.BOF) Then
    VerificaExistePermisoCargo = False
Else
    VerificaExistePermisoCargo = rsVar!bExiste
End If

Set rsVar = Nothing
Set Conn = Nothing
End Function
'WIOR FIN ***************************************************************




