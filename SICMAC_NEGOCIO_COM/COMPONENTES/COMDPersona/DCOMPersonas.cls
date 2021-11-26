VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMPersonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Option Explicit
Dim gConsPersona As String
Dim gConsComunes As String
Dim gConsImagenes As String

Enum TBusqueda
    BusquedaNombre = 1
    BusquedaCodigo = 2
    BusquedaDocumento = 3
    BusquedaEmpleadoNombre = 4
    BusquedaEmpleadoCodigo = 5
    BusquedaEmpleadoDocumento = 6
End Enum
'En caso de error
Public sMsgError As String
Public bError As Boolean

Dim ofun As New COMFunciones.FCOMImpresion
Dim oGen As New COMDConstSistema.DCOMGeneral
Dim ocad As New COMFunciones.FCOMCadenas
Public Function Get_CIIU_Busqueda(ByVal psCod As String) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCargar_CIIU
    
    sSql = "select cCIIUCod , cCIIUdescripcion  from ciiu where cCIIUCod like '%" & psCod & "%'"
    sSql = sSql & "order by cCIIUCod"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set Get_CIIU_Busqueda = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargar_CIIU:
    Set Get_CIIU_Busqueda = Nothing
    Err.Raise Err.Number, "Cargar CIIU ", Err.Description
End Function

Public Function BuscaPersCodConyuge(ByVal psPersCod As String) As String
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    sSql = "Select cPersRelacPersCod from PersRelaciones where nPersRelac = 0 and cPersCod = '" & psPersCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    If R.RecordCount > 0 Then
        BuscaPersCodConyuge = R!cPersRelacPersCod
    Else
        BuscaPersCodConyuge = ""
    End If
    
End Function


Public Function RecuperaPersonasTipo(ByVal psTipo As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaPersonasTipo
    
    sSql = "Select P.cPersCod, P.cPersNombre From Persona P inner join PersTpo PT"
    sSql = sSql & " ON P.cpersCod = PT.cPersCod Where PT.nPersTipo = '" & Right("00" & psTipo, 2) & "' AND PT.nPersEstado = 1 "
    sSql = sSql & " Order By cPersNombre "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPersonasTipo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaPersonasTipo:
    Set RecuperaPersonasTipo = Nothing
    Err.Raise Err.Number, "Recupera Personas de un Tipo", Err.Description
End Function
Public Function RecuperaPersonasTipo_Arbol(ByVal psTipo As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaPersonasTipo
    
    sSql = "Select P.cPersCod, P.cPersNombre, 1 as nItem From Persona P inner join PersTpo PT"
    sSql = sSql & " ON P.cpersCod = PT.cPersCod "
    sSql = sSql & " Where PT.nPersTipo = '" & Right("00" & psTipo, 2) & "' AND PT.nPersEstado = 1 "
    sSql = sSql & " Order By cPersNombre "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPersonasTipo_Arbol = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaPersonasTipo:
    Set RecuperaPersonasTipo_Arbol = Nothing
    Err.Raise Err.Number, "Recupera Personas de un Tipo", Err.Description
End Function

'MADM 20101012
Public Function RecuperaPersonasTipo_Arbol_Agencia(ByVal psTipo As String, pAgecod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaPersonasTipo
    
    sSql = "exec stp_sel_DevuelvePersonaTipoAgencia '" & psTipo & "','" & pAgecod & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPersonasTipo_Arbol_Agencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaPersonasTipo:
    Set RecuperaPersonasTipo_Arbol_Agencia = Nothing
    Err.Raise Err.Number, "Recupera Personas de un Tipo", Err.Description
End Function

Public Function BuscaCliente(ByVal sCadBusca As String, Optional ByVal nTipoBusq As TBusqueda = BusquedaNombre) As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim Conn As COMConecta.DCOMConecta
    bError = False
'    Set cmd = New ADODB.Command
'    Set prm = New ADODB.Parameter
'    cmd.CommandText = "spBuscaClienteApellido"
'    cmd.CommandType = adCmdStoredProc
'    cmd.Name = "spBuscaClienteApellido"
'    Set prm = cmd.CreateParameter("Apellido", adChar, adParamInput, 80)
'    cmd.Parameters.Append prm
    
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set BuscaCliente = Nothing
        Set Conn = Nothing
        Set cmd = Nothing
        Set prm = Nothing
        Exit Function
    End If
    'Set cmd.ActiveConnection = Conn.ConexionActiva
    'cmd.Parameters.Refresh
    'Set BuscaCliNombre = cmd.Execute(, (sCadBusca))
    'cmd.Execute
    sCadBusca = Replace(sCadBusca, "'", "''")
    Conn.ConexionActiva.CommandTimeout = 7200
    Select Case nTipoBusq
        Case 1 'Busqueda por nombre
            Set BuscaCliente = Conn.CargaRecordSet("spBuscaClienteApellido '" & Trim(sCadBusca) & "'")
        Case 2 'Busqueda por Codigo
            Set BuscaCliente = Conn.CargaRecordSet("spBuscaClienteCodigo '" & Trim(sCadBusca) & "'")
        Case 3 'Busqueda por Documento
            Set BuscaCliente = Conn.CargaRecordSet("spBuscaClienteDoc '" & Trim(sCadBusca) & "'")
        Case 4 'Busqueda Empleado por nombre
            Set BuscaCliente = Conn.CargaRecordSet("spBuscaEmpleadoApellido '" & Trim(sCadBusca) & "'")
        Case 5 'Busqueda Empleado por Codigo
            Set BuscaCliente = Conn.CargaRecordSet("spBuscaEmpleadoCodigo '" & Trim(sCadBusca) & "'")
        Case 6 'Busqueda Empleado por Documento
            Set BuscaCliente = Conn.CargaRecordSet("spBuscaEmpleadoDoc '" & Trim(sCadBusca) & "'")
    End Select
    Conn.CierraConexion
    Set Conn = Nothing
End Function
'20090910 MADM ******************
Public Function BuscaCliente_1(ByVal sCadBusca As String, Optional ByVal nTipoBusq As TBusqueda = BusquedaNombre) As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim Conn As COMConecta.DCOMConecta
    bError = False
    
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set BuscaCliente_1 = Nothing
        Set Conn = Nothing
        Set cmd = Nothing
        Set prm = Nothing
        Exit Function
    End If
    sCadBusca = Replace(sCadBusca, "'", "''")
    Conn.ConexionActiva.CommandTimeout = 7200
    Select Case nTipoBusq
        Case 1 'Busqueda por nombre
            Set BuscaCliente_1 = Conn.CargaRecordSet("spBuscaClienteApellido_1 '" & Trim(sCadBusca) & "'")
        Case 2 'Busqueda por Codigo NO TENEMOS AUN
            'Set BuscaCliente_1 = Conn.CargaRecordSet("spBuscaClienteCodigo '" & Trim(sCadBusca) & "'")
        Case 3 'Busqueda por Documento
            Set BuscaCliente_1 = Conn.CargaRecordSet("spBuscaClienteDoc_1 '" & Trim(sCadBusca) & "'")
    End Select
    Conn.CierraConexion
    Set Conn = Nothing
End Function

Public Function BuscaDocumento(ByVal pnTipo As Integer, ByVal psNumero As String, ByVal psPersCod As String) As Boolean
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset
    On Error GoTo ErrorBuscaDocumento
    'COMENTADO POR LARI - PTI1 20210708***************
    'sSql = "Select cPersIDnro from " & gConsPersona & "PersID where cPersIDTpo = '" & Trim(Str(pnTipo)) & "' And cPersIDnro = '" & psNumero & "'"
    sSql = "Select cPersIDnro from " & gConsPersona & "PersID where cPersIDTpo = '" & Trim(Str(pnTipo)) & "' And cPersIDnro = '" & psNumero & "' AND cPersCod <>" & "'" & psPersCod & "'"
    '*************************************************
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        BuscaDocumento = True
    Else
        BuscaDocumento = False
    End If
    R.Close
    Set R = Nothing
    Set oConecta = Nothing
    Exit Function
    
ErrorBuscaDocumento:
    Err.Raise Err.Number, "BuscaDocumento", Err.Description
    BuscaDocumento = False
End Function

Public Function ObtienePorcentajeGarantia(ByVal psCodTipoGar As Integer) As Double
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
Dim sSql As String

    On Error GoTo ErrorObtienePorcentajeGarantia
    sSql = "Select * From Colocparametro where nParamVar = " & "1027" & psCodTipoGar
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        ObtienePorcentajeGarantia = R!nParamValor
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    
    Exit Function
    
ErrorObtienePorcentajeGarantia:
    MsgBox Err.Description, vbInformation, "Aviso"
    
End Function

Private Sub Class_Initialize()
Dim CIni As COMConecta.DCOMClasIni
    Set CIni = New COMConecta.DCOMClasIni
    gConsPersona = CIni.BasePersonas
    gConsComunes = CIni.BaseComunes
    gConsImagenes = CIni.BaseImagenes
    
    Set ofun = New COMFunciones.FCOMImpresion
    Set oGen = New COMDConstSistema.DCOMGeneral
    Set ocad = New COMFunciones.FCOMCadenas

    Set CIni = Nothing
End Sub


' ********************************************************************
' DATOS DE PERSONAS  ' CRSF - 03/10
' *******************************************************************
Public Function dDatosPersonas(ByVal psCodPers As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaPersonasTipo
    
    sSql = "Select P.cPersCod, P.cPersNombre , P.cPersDireccDomicilio, PD.cPersIDnro "
    sSql = sSql & "From Persona P Join PersID PD On P.cPersCod = PD.cPersCod Where P.cPersCod = '" & psCodPers & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set dDatosPersonas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaPersonasTipo:
    Set dDatosPersonas = Nothing
    Err.Raise Err.Number, "Datos Personas ", Err.Description
End Function

Public Function Cargar_CIIU(psCodCmac As String) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCargar_CIIU
    
    If psCodCmac = "102" Then
        sSql = " Select cCIIUcod,cCIIUdescripcion, 1 Nro from CIIU  where cciiucod in ('O9309')  " _
             & "  union all " _
             & "  Select cCIIUcod ,cCIIUdescripcion, 2 nro  from CIIU where cciiucod not in ('O9309')" _
             & " order by nro, cCIIUdescripcion "
    Else
        sSql = "Select cCIIUcod,cCIIUdescripcion from CIIU Order by cCIIUdescripcion"
    End If
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set Cargar_CIIU = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargar_CIIU:
    Set Cargar_CIIU = Nothing
    Err.Raise Err.Number, "Cargar CIIU ", Err.Description
End Function

'*** PEAC 20080412
Public Function Cargar_TIPOCOMP(psCodCmac As String) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCargar_TIPOCOMP
    
    If psCodCmac = "102" Then
        sSql = " Select cCIIUcod,cCIIUdescripcion, 1 Nro from CIIU  where cciiucod in ('O9309')  " _
             & "  union all " _
             & "  Select cCIIUcod ,cCIIUdescripcion, 2 nro  from CIIU where cciiucod not in ('O9309')" _
             & " order by nro, cCIIUdescripcion "
    Else
        'sSql = "Select cCIIUcod,cCIIUdescripcion from CIIU Order by cCIIUdescripcion"
        sSql = "exec stp_sel_ObtieneTipoCompetencia"
    End If
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set Cargar_TIPOCOMP = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargar_TIPOCOMP:
    Set Cargar_TIPOCOMP = Nothing
    Err.Raise Err.Number, "Cargar TIPOCOMP ", Err.Description
End Function
'*** PEAC 20080412
Public Function Cargar_TIPOSISTINFOR(psCodCmac As String) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCargar_TIPOSISTINFOR
    
    If psCodCmac = "102" Then
        sSql = " Select cCIIUcod,cCIIUdescripcion, 1 Nro from CIIU  where cciiucod in ('O9309')  " _
             & "  union all " _
             & "  Select cCIIUcod ,cCIIUdescripcion, 2 nro  from CIIU where cciiucod not in ('O9309')" _
             & " order by nro, cCIIUdescripcion "
    Else
        sSql = "exec stp_sel_ObtieneTipoSistemaInformacion"
    End If
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set Cargar_TIPOSISTINFOR = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargar_TIPOSISTINFOR:
    Set Cargar_TIPOSISTINFOR = Nothing
    Err.Raise Err.Number, "Cargar TIPOSISTINFOR ", Err.Description
End Function

'*** PEAC 20080428
Public Function Cargar_MotivoActPersona(psCodCmac As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCargar_MotivoAct
    
    sSql = "exec stp_sel_ObtieneMotivoActualizaPersona "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set Cargar_MotivoActPersona = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargar_MotivoAct:
    Set Cargar_MotivoActPersona = Nothing
    Err.Raise Err.Number, "Cargar ", Err.Description
End Function
'madm 20100326
Public Function Cargar_AlternativaUsual(psCodCmac As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCargar_AlternativaUsual
    
    sSql = "exec stp_sel_ObtieneAlternativaUsual"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set Cargar_AlternativaUsual = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargar_AlternativaUsual:
    Set Cargar_AlternativaUsual = Nothing
    Err.Raise Err.Number, "Cargar ", Err.Description
End Function
'end madm

'*** PEAC 20080428
Public Function Cargar_AlternativaSiNo(psCodCmac As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCargar_AlternativaSiNo
    
    sSql = "exec stp_sel_ObtieneAlternativaSiNo"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set Cargar_AlternativaSiNo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargar_AlternativaSiNo:
    Set Cargar_AlternativaSiNo = Nothing
    Err.Raise Err.Number, "Cargar ", Err.Description
End Function

'*** PEAC 20080428
Public Function Cargar_TipoPatri(psCodCmac As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCargar_TipoPatri
    
    sSql = "exec stp_sel_ObtieneTipoPatrimonio"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set Cargar_TipoPatri = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargar_TipoPatri:
    Set Cargar_TipoPatri = Nothing
    Err.Raise Err.Number, "Cargar TipoPatri ", Err.Description
End Function

'*** PEAC 20080428
Public Function Cargar_MonePatri(psCodCmac As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCargar_MonePatri
    
    sSql = "exec stp_sel_ObtieneMonePatri"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set Cargar_MonePatri = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargar_MonePatri:
    Set Cargar_MonePatri = Nothing
    Err.Raise Err.Number, "Cargar MonePatri ", Err.Description
End Function



'*** PEAC 20080412
Public Function Cargar_CADENAPROD(psCodCmac As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCargar_CADENAPROD
    
    If psCodCmac = "102" Then
        sSql = " Select cCIIUcod,cCIIUdescripcion, 1 Nro from CIIU  where cciiucod in ('O9309')  " _
             & "  union all " _
             & "  Select cCIIUcod ,cCIIUdescripcion, 2 nro  from CIIU where cciiucod not in ('O9309')" _
             & " order by nro, cCIIUdescripcion "
    Else
        sSql = "exec stp_sel_ObtieneTipoCadenaProductiva"
    End If
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set Cargar_CADENAPROD = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargar_CADENAPROD:
    Set Cargar_CADENAPROD = Nothing
    Err.Raise Err.Number, "Cargar CADENAPROD ", Err.Description
End Function


Public Function CargarTipos_PerJuridicas() As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCargarTipos_PerJuridicas
    
    sSql = "Select cPersJurTpoCod,cPersJurTpoDesc  from persjurtpo Order by cPersJurTpoDesc"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargarTipos_PerJuridicas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargarTipos_PerJuridicas:
    Set CargarTipos_PerJuridicas = Nothing
    Err.Raise Err.Number, "Cargar Tipos de Persona Juridica", Err.Description
End Function
    
Public Function CargarEstadosPersona() As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCargarEstadosPersona
    
    sSql = "Select nConsValor,cConsDescripcion From Constante Where nConsCod = " & Trim(Str(gPersEstado)) & " and nConsValor <> " & Trim(Str(gPersEstado))
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargarEstadosPersona = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargarEstadosPersona:
    Set CargarEstadosPersona = Nothing
    Err.Raise Err.Number, "Cargar Estados de Persona", Err.Description
End Function
'madm 20100226 - 20110223
Public Function CargarOcupaciones(Optional ByVal pcFiltro As String = "") As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCargarOcupaciones
    
    'sSql = "select id as nConsValor,descripcion as cConsDescripcion from Ocupaciones "
    'sSql = "select descripcion as cConsDescripcion,id as nConsValor from Ocupaciones " 'NDX ERS0062020
    sSql = "select descripcion as cConsDescripcion, id as nConsValor from Ocupaciones WHERE id > 1000" 'NDX ERS0062020

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargarOcupaciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorCargarOcupaciones:
    Set CargarOcupaciones = Nothing
    Err.Raise Err.Number, "Cargar Ocupaciones", Err.Description
End Function
'end madm

'** Juez 20120326 ************************************
Public Function CargarCargos(Optional ByVal pcFiltro As String = "") As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCargarCargos
    
    sSql = "select PersCargoDesc as cConsDescripcion,PersCargoId as nConsValor from PersCargo "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargarCargos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorCargarCargos:
    Set CargarCargos = Nothing
    Err.Raise Err.Number, "Cargar Cargos", Err.Description
End Function
'** End Juez ******************************************

Public Function CargarUbicacionesGeograficas(Optional ByVal pbPaises As Boolean = False, _
                                             Optional ByVal pnNivel As Integer = -1, _
                                             Optional ByVal pcFiltro As String = "") As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorCargarUbicacionesGeograficas
    
    If pnNivel = -1 Then    'Filtrar todos los Niveles
        If pbPaises Then
            sSql = "select *, 1 p from UbicacionGeografica where cUbiGeoCod like '0%' "
            sSql = sSql & " Union "
            sSql = sSql & " Select *, 2 p from UbicacionGeografica where cUbiGeoCod like '1%'"
            sSql = sSql & " Union "
            sSql = sSql & " select *, 3 p from UbicacionGeografica where cUbiGeoCod like '2%' "
            sSql = sSql & " Union "
            sSql = sSql & " select *, 4 p from UbicacionGeografica where cUbiGeoCod like '3%' "
            sSql = sSql & " Union "
            sSql = sSql & " select *, 5 p from UbicacionGeografica where cUbiGeoCod like '4%' order by p,cUbiGeoDescripcion "
        Else
            sSql = "Select *, 1 p from UbicacionGeografica where cUbiGeoCod like '1%'"
            sSql = sSql & " Union "
            sSql = sSql & " select *, 2 p from UbicacionGeografica where cUbiGeoCod like '2%' "
            sSql = sSql & " Union "
            sSql = sSql & " select *, 3 p from UbicacionGeografica where cUbiGeoCod like '3%' "
            sSql = sSql & " Union "
            sSql = sSql & " select *, 4 p from UbicacionGeografica where cUbiGeoCod like '4%' order by p,cUbiGeoDescripcion "
        End If
    Else        'Filtro de un Nivel Especifico
        If pbPaises Then
            ' sSql = "SELECT * from UbicacionGeografica where cUbiGeoCod like '" & pnNivel & "%' " & _
            'IIf(pcFiltro <> "", " AND SUBSTRING(cUbiGeoCod,2," & 2 * (pnNivel - 1) & ")=SUBSTRING('" & pcFiltro & "',2," & 2 * (pnNivel - 1) & ")", "") & " ORDER BY cUbiGeoDescripcion"
            
            'WIOR 20120322-INICIO
             sSql = "SELECT" & IIf(pnNivel = 5, " cUbiGeoDescripcion,cUbiGeoCod ", " * ") & "from UbicacionGeografica where (cUbiGeoCod like '" & IIf(pnNivel = 5, (pnNivel - 4), pnNivel) & "%' " & _
             IIf(pcFiltro <> "", " AND SUBSTRING(cUbiGeoCod,2," & 2 * IIf(pnNivel = 5, (pnNivel - 5), (pnNivel - 1)) & ")=SUBSTRING('" & pcFiltro & "',2," & 2 * IIf(pnNivel = 5, (pnNivel - 5), (pnNivel - 1)) & "))", ")") & " ORDER BY cUbiGeoDescripcion"
            ' IIf(pnNivel = 5, " OR (len(cUbiGeoCod)=5 and cUbiGeoCod<>'" & pcFiltro & "') ORDER BY cUbiGeoCod", " ORDER BY cUbiGeoDescripcion")
        Else
            sSql = "SELECT * from UbicacionGeografica where cUbiGeoCod like '" & pnNivel & "%' " & _
                   IIf(pcFiltro <> "", " AND SUBSTRING(cUbiGeoCod,2," & 2 * (pnNivel - 1) & ")=SUBSTRING('" & pcFiltro & "',2," & 2 * (pnNivel - 1) & ")", "") & " ORDER BY cUbiGeoDescripcion"
        End If
    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set CargarUbicacionesGeograficas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorCargarUbicacionesGeograficas:
    Set CargarUbicacionesGeograficas = Nothing
    Err.Raise Err.Number, "Cargar Ubicaciones Geograficas", Err.Description
End Function

Public Function CargarDatosObjetosPersona(ByRef prsRelac As ADODB.Recordset, ByRef prsDocumentos As ADODB.Recordset, _
ByRef prsRefComercial As ADODB.Recordset, ByRef prsPatVehicular As ADODB.Recordset, _
ByRef prsTipoSangre As ADODB.Recordset, ByRef prsDireccCondic As ADODB.Recordset, _
ByRef prsPersoneria As ADODB.Recordset, ByRef prsMagnitudComer As ADODB.Recordset, _
ByRef prsEstCivil As ADODB.Recordset, ByRef prsRelacInst As ADODB.Recordset, _
ByRef prsUbiGeo As ADODB.Recordset, ByRef prsCIIU As ADODB.Recordset, ByVal psCodCmac As String, _
ByRef prsTipoPerJuri As ADODB.Recordset, ByRef prsVisita As ADODB.Recordset, Optional prsTIPOCOMP As ADODB.Recordset = Nothing, Optional prsTIPOSISTINFOR As ADODB.Recordset = Nothing, _
Optional prsCADENAPROD As ADODB.Recordset = Nothing, Optional prsTipoPatri As ADODB.Recordset = Nothing, Optional prsMonePatri As ADODB.Recordset = Nothing, Optional prsAlterSiNo As ADODB.Recordset = Nothing, Optional prsMotivoActu As ADODB.Recordset = Nothing, _
Optional plrsOcupa As ADODB.Recordset = Nothing, Optional plrsCargo As ADODB.Recordset = Nothing) '** Juez 20120326 Se agregó plrsCargos

'*** PEAC 20080801
'SE AGREGO EL PARAMETRO "prsMotivoActu"

'madm 20100305 ------------- Se agrego el parametro plrsOcupa
'madm 20100329 ------------- Se agrego el parametro prsVisita

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim oConstante As COMDConstantes.DCOMConstantes
 
On Error GoTo ErrorCargarDatosObjetosPersona
    

Set oConstante = New COMDConstantes.DCOMConstantes 'DConstante

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

Set prsRelac = oConstante.RecuperaConstantes(gPersRelacion)
Set prsDocumentos = oConstante.RecuperaConstantes(gPersIdTipo)
Set prsRefComercial = oConstante.RecuperaConstantes(3028)
Set prsPatVehicular = oConstante.RecuperaConstantes(3029)
         
Set prsTipoSangre = oConstante.RecuperaConstantes(gPersTpoSangre)
Set prsDireccCondic = oConstante.RecuperaConstantes(gPersCondDomic)
Set prsPersoneria = oConstante.RecuperaConstantes(gPersPersoneria)
'JACA 20110427*****************************************************************
'Set prsMagnitudComer = oConstante.RecuperaConstantes(gPersJurMagnitud)
 Set prsMagnitudComer = oConstante.RecuperaMagnitudPersona(gPersJurMagnitud, 2)
'JACA END*****************************************************************

Set prsEstCivil = oConstante.RecuperaConstantes(gPersEstadoCivil)
Set prsRelacInst = oConstante.RecuperaConstantes(gPersRelacionInst)
          
Set prsUbiGeo = CargarUbicacionesGeograficas(True, 0)
Set plrsOcupa = CargarOcupaciones() 'madm 20100305
Set plrsCargo = CargarCargos() '** Juez 20120326

Set prsCIIU = Cargar_CIIU(psCodCmac)
'*** PEAC 20080412 **********************************************
Set prsTIPOCOMP = Cargar_TIPOCOMP(psCodCmac)
Set prsTIPOSISTINFOR = Cargar_TIPOSISTINFOR(psCodCmac)
Set prsCADENAPROD = Cargar_CADENAPROD(psCodCmac)
Set prsTipoPatri = Cargar_TipoPatri(psCodCmac)
Set prsMonePatri = Cargar_MonePatri(psCodCmac)
Set prsAlterSiNo = Cargar_AlternativaSiNo(psCodCmac)
Set prsMotivoActu = Cargar_MotivoActPersona(psCodCmac)
'*** FIN PEAC  **************************************************

Set prsVisita = Cargar_AlternativaUsual(psCodCmac) 'madm 20100326

Set prsTipoPerJuri = CargarTipos_PerJuridicas
   
oConecta.CierraConexion

Set oConstante = Nothing
Set oConecta = Nothing
Exit Function
    
ErrorCargarDatosObjetosPersona:
    Set CargarDatosObjetosPersona = Nothing
    Err.Raise Err.Number, "Cargar Datos Objetos Persona", Err.Description
End Function

Public Function UltimoCorrelativo(ByVal psCodAge As String, Optional ByVal pbPersDistrib As Boolean = False, _
                                  Optional ByVal pConexDitrib As ADODB.Connection = Nothing) As String

Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta  As COMConecta.DCOMConecta

    On Error GoTo ErrorUltimoCorrelativo
    sSql = "Select MAX(substring(P.cPersCod,6,7)) as Correlativo from " & gConsPersona & "Persona P where substring(P.cPersCod,4,2) = '" & Right(psCodAge, 2) & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        UltimoCorrelativo = IIf(IsNull(R!Correlativo), 0, R!Correlativo)
    Else
        UltimoCorrelativo = 0
    End If
    R.Close
    
    If pbPersDistrib Then
        sSql = "Select MAX(substring(P.cCodPers,3,7)) as Correlativo from Persona P where substring(P.cCodPers,1,2) = '" & Right(psCodAge, 2) & "'"
        R.CursorLocation = adUseClient
        R.Open sSql, pConexDitrib, adOpenStatic, adLockReadOnly, adCmdText
        If Not R.BOF And Not R.EOF Then
            If R!Correlativo > UltimoCorrelativo Then
                UltimoCorrelativo = R!Correlativo
            End If
        End If
        R.Close
        Set R = Nothing
        oConecta.CierraConexion
        Set oConecta = Nothing
    End If
    Exit Function
    
ErrorUltimoCorrelativo:
    Set R = Nothing
    Set oConecta = Nothing
    Err.Raise Err.Number, "Ultimo Correlativo", Err.Description
End Function

'***************************************************
' Funcion que recupera todos los Datos de la Persona
'***************************************************

Public Function RecuperarDatosPersona(ByVal psPersCod As String, _
                                    ByVal psCodAge As String, _
                                    ByRef prsPersona As ADODB.Recordset, _
                                    ByRef prsFirma As ADODB.Recordset, _
                                    ByRef prsDocs As ADODB.Recordset, _
                                    ByRef prsRelac As ADODB.Recordset, _
                                    ByRef prsFtesIng As ADODB.Recordset, _
                                    ByRef prsRefComer As ADODB.Recordset, _
                                    ByRef pnMaxRefComer As Integer, _
                                    ByRef prsRefBanc As ADODB.Recordset, _
                                    ByRef prsPatVehic As ADODB.Recordset, _
                                    ByRef pnMaxPatVehic As Integer, _
                                    Optional ByVal psCtaCod As String = "", _
                                    Optional ByRef prsPatOtros As ADODB.Recordset, _
                                    Optional ByRef pnMaxPatOtros As Integer, _
                                    Optional ByRef prsPatInmuebles As ADODB.Recordset, _
                                    Optional ByRef pnMaxPatInmuebles As Integer, _
                                    Optional ByRef prsVisita As ADODB.Recordset, _
                                    Optional ByRef prsVentas As ADODB.Recordset, _
                                    Optional ByRef prsDatAdc As ADODB.Recordset, _
                                    Optional ByRef prsPersRRHH As ADODB.Recordset, _
                                    Optional ByRef psCoduser As String)

'prsVisita add by MADM 20100327
'prsPersRRHH add by MADM 20101020
Dim sSql As String
Dim oConecta  As COMConecta.DCOMConecta
Dim oConFirma As COMConecta.DCOMConecta
Dim sCadCon As String
Dim R As ADODB.Recordset

    On Error GoTo ErrorRecuperarDatosPersona
    
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
            
        '***** CARGA DATOS DE LA PERSONA ******
'        sSql = "Select  P.nPersRelaInst, P.cPersCodSbs, P.cPersEmail cEmail,P.cPersTelefono2,P.cPersCod,P.cPersNombre,P.dPersNacCreac, "
'        sSql = sSql & " P.cPersDireccUbiGeo,P.cPersDireccDomicilio,P.cPersDireccCondicion, P.nPersValComDomicilio, P.cPersTelefono, "
'        sSql = sSql & " P.nPersPersoneria,P.cPersCIIU,P.cPersEstado,P.cUltimaActualizacion, "
'        sSql = sSql & " PN.cPersNatSexo,PN.nPersNatEstCiv,PN.nPersNatHijos,PN.nPersPeso,PN.nPersTalla,PN.nPersTpoSangre, PN.cNacionalidad, PN.nResidente, "
'        sSql = sSql & " PJ.cPersJurTpo, PJ.cPersJurMagnitud,PJ.cPersJurSigla,PJ.nPersJurEmpleados, CI.cCIIUdescripcion,nPersIngresoProm=ISNULL(nPersIngresoProm,0)," 'CUSCO
'        sSql = sSql & " Dpto = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '1' And Substring(u.cUbigeoCod,2,2) = Substring(P.cPersDireccUbiGeo,2,2) And Substring(u.cUbigeoCod,4,9) = '000000000' ), '' ) , "
'        sSql = sSql & " Prov = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '2' And Substring(u.cUbigeoCod,2,4) = Substring(P.cPersDireccUbiGeo,2,4) And Substring(u.cUbigeoCod,6,7) = '0000000' ), '' ) , "
'        sSql = sSql & " Dist = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '3' And Substring(u.cUbigeoCod,2,6) = Substring(P.cPersDireccUbiGeo,2,6) And Substring(u.cUbigeoCod,8,5) = '00000' ), '' ), "
'        sSql = sSql & " Zona = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where u.cUbigeoCod = P.cPersDireccUbiGeo ),''), "
'        sSql = sSql & " CN7.cConsDescripcion as cPersDireccCondicionDes, CN8.cConsDescripcion as cPersJurMagnitudDes"
'        sSql = sSql & " FROM " & gConsPersona & "Persona P left join " & gConsPersona & "PersonaNat PN ON P.cPersCod=PN.cPersCod "
'        sSql = sSql & " LEFT JOIN " & gConsPersona & "PersonaJur PJ ON P.cPersCod = PJ.cPersCod "
'        sSql = sSql & " LEFT JOIN CIIU CI ON P.cPersCIIU = CI.cCIIUcod "
'        sSql = sSql & " Left Join Constante CN7 ON CN7.nConsValor = P.cPersDireccCondicion AND CN7.nConsCod = 1029 "
'        sSql = sSql & " Left Join Constante CN8 ON CN8.nConsValor = PJ.cPersJurMagnitud AND CN8.nConsCod = 1004 "
'        sSql = sSql & " WHERE P.cPersCod = '" & psPersCod & "'"
    
        sSql = "exec stp_sel_ObtieneDatosPersona_nuevo '" & psPersCod & "'"

    
        Set prsPersona = oConecta.CargaRecordSet(sSql)
        If prsPersona.EOF Then
            oConecta.CierraConexion
            Set prsPersona = Nothing
            Set oConecta = Nothing
            Exit Function
        End If
    
    '***************** Firma ********************
    
        sCadCon = ""
        sSql = "SELECT * FROM Servidor WHERE RIGHT(cCodAge,2) = '" & psCodAge & "' AND cNroSer = '05'"
        Set R = New ADODB.Recordset
        R.Open sSql, oConecta.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
        If R.RecordCount > 0 Then
            sCadCon = "Provider=" & Trim(R!cProvider) & ";User ID=" & Trim(R!cLogin) & ";Initial Catalog=" & Trim(R!cDataBase) & ";Data Source=" & Trim(R!cNomSer) & ";PWD=" & Trim(R!cPass)
        End If
        R.Close
        
        sSql = "Select cPersCod,iPersFirma,cUltimaActualizacion from " & gConsImagenes & "PersImagen PI Where cPersCod = '" & psPersCod & "'"
        Set R = New ADODB.Recordset
        If sCadCon <> "" Then
            Set oConFirma = New COMConecta.DCOMConecta
            Call oConFirma.AbreConexion(sCadCon)
            R.Open sSql, oConFirma.ConexionActiva, adOpenKeyset, adLockOptimistic
        Else
            R.Open sSql, oConecta.ConexionActiva, adOpenKeyset, adLockOptimistic
        End If
        
        Set prsFirma = R
        
        '***************** Documentos ****************
        'sSql = "Select PID.cPersIDnro,PID.cPersIDTpo,C.cConsDescripcion,C.nConsValor From " & gConsPersona & "PersID PID inner join Constante C ON convert(int,PID.cPersIDTpo) = C.nConsValor Where PID.cPersCod = '" & psPersCod & "' And  C.nConsCod = " & Trim(Str(gPersIdTipo)) & " And C.cConsDescripcion <> '" & Trim(Str(gPersIdTipo)) & "'"
        sSql = "exec stp_sel_DocumentosPersona '" & psPersCod & "'" 'JACA 20111004
        Set prsDocs = oConecta.CargaRecordSet(sSql)

        '****************** Relaciones de Personas *********
        sSql = "Select PR.bPersRelacAMP, PR.nPersRelacBenefPorc, PR.bPersRelacBenef, C3.cConsDescripcion cConsDescripcion3, C2.cConsDescripcion cConsDescripcion2,C2.nConsValor nConsValor2, C.cConsDescripcion,PR.cPersRelacPersCod,P.cPersNombre,PR.nPersRelac,PR.cMovNro "
        sSql = sSql & " from " & gConsPersona & "PersRelaciones PR Inner join Persona P ON PR.cPersRelacPersCod = P.cPersCod"
        sSql = sSql & " Inner join Constante C on PR.nPersRelac = C.nConsValor"
        sSql = sSql & " Left join Constante C2 on PR.bPersRelacBenef = C2.nConsValor"
        sSql = sSql & " Left join Constante C3 on PR.bPersRelacAMP = C3.nConsValor"
        sSql = sSql & " Where PR.cPersCod ='" & psPersCod & "' and C.nConsCod = " & gPersRelacion & " And C.nConsValor <> " & gPersRelacion & " and C2.nConsCod = " & gPersRelacBenef & " And C2.nConsValor <> " & gPersRelacBenef & " and C3.nConsCod = " & gPersRelacAmp & " And C3.nConsValor <> " & gPersRelacAmp
        Set prsRelac = oConecta.CargaRecordSet(sSql)

        '******************** Visitas ***************************
        '*** MADM 20100327
        
        sSql = " exec stp_sel_PersRecuperaVisitas '" & psPersCod & "'"
        
        Set prsVisita = oConecta.CargaRecordSet(sSql)
        '*** MADM 20101020
        '***********************************************
        sSql = " exec stp_sel_devuelvePersonalRRHH '" & psCoduser & "'"

        Set prsPersRRHH = oConecta.CargaRecordSet(sSql)
        
        '***********************************************
        
        'MAVM 20100607 BAS II ***
        sSql = " exec stp_sel_PersRecuperaVentas '" & psPersCod & "'"
        Set prsVentas = oConecta.CargaRecordSet(sSql)
        '***
        
        '******************** Recupera Fuentes de Ingreso *******
        
        If psCtaCod = "" Then
    
            sSql = "Select PF.bCostoProd, PF.cNumFuente, PF.cPersCod, PF.cPersFIPersCod, PF.cPersFICargo, "
            sSql = sSql & " PF.dPersFIinicio, PF.cPersFIMoneda, PF.cPersFIComentario, "
            sSql = sSql & " PF.cRazSocUbiGeo, PF.cRazSocDirecc, PF.cRazSocTelef, PF.cRazSocDescrip, "
            sSql = sSql & " PF.nPersTipFte, PF.cUltimaActualizacion, "
        '------------------------------------------------------------------------------------------------------------
            sSql = sSql & "P.cPersDireccUbiGeo,P.cPersDireccDomicilio,P.cPersDireccCondicion, P.cPersTelefono, "
            sSql = sSql & "P.nPersPersoneria,P.cPersCIIU, P.cPersTelefono2, P.cPersCod as cPersFteCod,P.cPersNombre,P.dPersNacCreac, "
            sSql = sSql & "PJ.cPersJurTpo, PJ.cPersJurMagnitud,PJ.cPersJurSigla,PJ.nPersJurEmpleados, CI.cCIIUdescripcion, "
            sSql = sSql & "cNroRUC = (  Select ISNULL(cPersIDnro,'') From PersID WHERE PersID.cPersCod = P.cPersCod and PersID.cPersIDTpo = " & gPersIdRUC & " ), "
            sSql = sSql & "PJT.cPersJurTpoDesc, CN7.cConsDescripcion as cPersDireccCondicionDes, CN8.cConsDescripcion as cPersJurMagnitudDes, "
            sSql = sSql & " Dpto = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '1' And Substring(u.cUbigeoCod,2,2) = Substring(P.cPersDireccUbiGeo,2,2) And Substring(u.cUbigeoCod,4,9) = '000000000' ), '' ) , "
            sSql = sSql & " Prov = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '2' And Substring(u.cUbigeoCod,2,4) = Substring(P.cPersDireccUbiGeo,2,4) And Substring(u.cUbigeoCod,6,7) = '0000000' ), '' ) , "
            sSql = sSql & " Dist = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '3' And Substring(u.cUbigeoCod,2,6) = Substring(P.cPersDireccUbiGeo,2,6) And Substring(u.cUbigeoCod,8,5) = '00000' ), '' ), "
            sSql = sSql & " Zona = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where u.cUbigeoCod = P.cPersDireccUbiGeo ),'') "
            sSql = sSql & " From  PersFteIngreso PF "
            sSql = sSql & " INNER JOIN " & gConsPersona & "Persona P ON PF.cPersFIPersCod = P.cPersCod "
            sSql = sSql & " left join " & gConsPersona & "PersonaNat PN ON P.cPersCod=PN.cPersCod "
            sSql = sSql & " LEFT JOIN " & gConsPersona & "PersonaJur PJ ON P.cPersCod = PJ.cPersCod "
            sSql = sSql & " LEFT JOIN CIIU CI ON P.cPersCIIU = CI.cCIIUcod "
            sSql = sSql & " LEFT JOIN PersJurTpo PJT ON PJ.cPersJurTpo = PJT.cPersJurTpoCod "
            sSql = sSql & " Left Join Constante CN7 ON CN7.nConsValor = P.cPersDireccCondicion AND CN7.nConsCod = 1029 "
            sSql = sSql & " Left Join Constante CN8 ON CN8.nConsValor = PJ.cPersJurMagnitud AND CN8.nConsCod = 1004 "
            sSql = sSql & " Where PF.cPersCod = '" & psPersCod & "'"
            sSql = sSql & " and (exists(Select cnumfuente from persfidependiente where cnumfuente=PF.cNumfuente) or Exists(Select cnumfuente from persfiindependiente where cnumfuente=PF.cnumfuente))"
            sSql = sSql & " Order By dPersFIinicio desc"
        '----------------------------------------------------------------------------------------------------------
        Else
            sSql = "Select PF.bCostoProd, PF.cNumFuente, PF.cPersCod, PF.cPersFIPersCod, PF.cPersFICargo, "
            sSql = sSql & " PF.dPersFIinicio, PF.cPersFIMoneda, PF.cPersFIComentario, "
            sSql = sSql & " PF.cRazSocUbiGeo, PF.cRazSocDirecc, PF.cRazSocTelef, PF.cRazSocDescrip, "
            sSql = sSql & " PF.nPersTipFte, PF.cUltimaActualizacion, "
        '------------------------------------------------------------------------------------------------------------
            sSql = sSql & "P.cPersDireccUbiGeo,P.cPersDireccDomicilio,P.cPersDireccCondicion, P.cPersTelefono, "
            sSql = sSql & "P.nPersPersoneria,P.cPersCIIU, P.cPersTelefono2, P.cPersCod as cPersFteCod,P.cPersNombre,P.dPersNacCreac, "
            sSql = sSql & "PJ.cPersJurTpo, PJ.cPersJurMagnitud,PJ.cPersJurSigla,PJ.nPersJurEmpleados, CI.cCIIUdescripcion, "
            sSql = sSql & "cNroRUC = (  Select ISNULL(cPersIDnro,'') From PersID WHERE PersID.cPersCod = P.cPersCod and PersID.cPersIDTpo = " & gPersIdRUC & " ), "
            sSql = sSql & "PJT.cPersJurTpoDesc, CN7.cConsDescripcion as cPersDireccCondicionDes, CN8.cConsDescripcion as cPersJurMagnitudDes, "
            sSql = sSql & " Dpto = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '1' And Substring(u.cUbigeoCod,2,2) = Substring(P.cPersDireccUbiGeo,2,2) And Substring(u.cUbigeoCod,4,9) = '000000000' ), '' ) , "
            sSql = sSql & " Prov = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '2' And Substring(u.cUbigeoCod,2,4) = Substring(P.cPersDireccUbiGeo,2,4) And Substring(u.cUbigeoCod,6,7) = '0000000' ), '' ) , "
            sSql = sSql & " Dist = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '3' And Substring(u.cUbigeoCod,2,6) = Substring(P.cPersDireccUbiGeo,2,6) And Substring(u.cUbigeoCod,8,5) = '00000' ), '' ), "
            sSql = sSql & " Zona = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where u.cUbigeoCod = P.cPersDireccUbiGeo ),'') "
            sSql = sSql & " From  PersFteIngreso PF "
            sSql = sSql & " INNER JOIN " & gConsPersona & "Persona P ON PF.cPersFIPersCod = P.cPersCod "
            sSql = sSql & " left join " & gConsPersona & "PersonaNat PN ON P.cPersCod=PN.cPersCod "
            sSql = sSql & " LEFT JOIN " & gConsPersona & "PersonaJur PJ ON P.cPersCod = PJ.cPersCod "
            sSql = sSql & " LEFT JOIN CIIU CI ON P.cPersCIIU = CI.cCIIUcod "
            sSql = sSql & " LEFT JOIN PersJurTpo PJT ON PJ.cPersJurTpo = PJT.cPersJurTpoCod "
            sSql = sSql & " Left Join Constante CN7 ON CN7.nConsValor = P.cPersDireccCondicion AND CN7.nConsCod = 1029 "
            sSql = sSql & " Left Join Constante CN8 ON CN8.nConsValor = PJ.cPersJurMagnitud AND CN8.nConsCod = 1004 "
            sSql = sSql & " Inner Join ColocFteIngreso CFI on CFI.cNumFuente=PF.cNumFuente"
            sSql = sSql & " Where PF.cPersCod = '" & psPersCod & "'"
            sSql = sSql & " and (exists(Select cnumfuente from persfidependiente where cnumfuente=PF.cNumfuente) or Exists(Select cnumfuente from persfiindependiente where cnumfuente=PF.cnumfuente))"
            sSql = sSql & " and CFI.cCtaCod='" & psCtaCod & "'"
            sSql = sSql & " Order By dPersFIinicio desc"
        '----------------------------------------------------------------------------------------------------------
        End If
    
        Set prsFtesIng = oConecta.CargaRecordSet(sSql)

        '******************* Recupera Referencias Comerciales *********
'        sSql = " SELECT nNumRefCom, cNomRefCom, nTipoRelaCom, cConsDescripcion, cFonoRefCom "
'        sSql = sSql & " FROM PERSREFCOMERCIAL RC INNER JOIN CONSTANTE C ON RC.nTipoRelaCom = C.nConsValor "
'        sSql = sSql & " WHERE nConsCod = 3028 AND cPersCod = '" & psPersCod & "'"
'        sSql = sSql & " ORDER BY nNumRefCom"

        '*** PEAC 20080412
        sSql = " exec stp_sel_recuperaReferenciaPersComer '" & psPersCod & "'"
        
        Set prsRefComer = oConecta.CargaRecordSet(sSql)

        '******************* Recupera Maxima Referencia Comercial *********
        sSql = "SELECT MAX(nNumRefCom) as nMaxRefCom FROM PERSREFCOMERCIAL WHERE cPersCod = '" & psPersCod & "'"
        Set R = oConecta.CargaRecordSet(sSql)
    
        If Not R.EOF And Not R.BOF Then
            pnMaxRefComer = IIf(IsNull(R!nMaxRefCom), 0, R!nMaxRefCom)
        Else
            pnMaxRefComer = 0
        End If

        '******************* Recupera las Referencias Bancarias *********
        sSql = "SELECT I.cIFTpo + '.' + P.cPersCod cCodBanco, CONVERT(CHAR(30),P.cPersNombre ) AS cBanNombre, "
        sSql = sSql & "cNumCta , cNumTarjeta, nMontoLinCred "
        sSql = sSql & " FROM PERSREFBANCARIA RB INNER JOIN PERSONA P ON RB.cCodBanco = P.cPersCod "
        sSql = sSql & " INNER JOIN INSTITUCIONFINANC I ON P.cPersCod = I.cPersCod "
        sSql = sSql & " WHERE RB.cPersCod = '" & psPersCod & "'"
        Set prsRefBanc = oConecta.CargaRecordSet(sSql)
        
        '******************** Recupera Patrimonio Vehicular *******
'        sSql = " SELECT nNumPatri, nCondPatri, cMarca, cFecFabrica, nValorCom, cConsDescripcion "
'        sSql = sSql & " FROM PERSPATRIMONIO P INNER JOIN CONSTANTE C ON P.nCondPatri = C.nConsValor "
'        sSql = sSql & " WHERE nConsCod = 3029 AND cPersCod = '" & psPersCod & "'"
'        sSql = sSql & " ORDER BY nNumPatri"
        
        sSql = "exec stp_sel_recuperaPatriVehicular '" & psPersCod & "'"
        
        Set prsPatVehic = oConecta.CargaRecordSet(sSql)

        '******************** Maximo Patrimonio Vehicular ************
        sSql = "SELECT MAX(nNumPatri) as nMaxPatVeh FROM PERSPATRIMONIO WHERE cPersCod = '" & psPersCod & "'"
        Set R = oConecta.CargaRecordSet(sSql)
    
        If Not R.EOF And Not R.BOF Then
            pnMaxPatVehic = IIf(IsNull(R!nMaxPatVeh), 0, R!nMaxPatVeh)
        Else
            pnMaxPatVehic = 0
        End If

        '*** PEAC 20080412 ------------------------------------------------------------------
        '******************** Recupera Patrimonio Otros *******
       
        sSql = "exec stp_sel_recuperaPatriOtros '" & psPersCod & "'"
        
        Set prsPatOtros = oConecta.CargaRecordSet(sSql)

        '******************** Maximo Patrimonio Vehicular ************
        sSql = "SELECT MAX(nNumPatri) as nMaxPatOtros FROM PERSPATRIotros WHERE cPersCod = '" & psPersCod & "'"
        Set R = oConecta.CargaRecordSet(sSql)
    
        If Not R.EOF And Not R.BOF Then
            pnMaxPatOtros = IIf(IsNull(R!nMaxPatOtros), 0, R!nMaxPatOtros)
        Else
            pnMaxPatOtros = 0
        End If

        '*** PEAC 20080412 ------------------------------------------------------------------
        '******************** Recupera Patrimonio Inmuebles *******
       
        sSql = "exec stp_sel_recuperaPatriInmuebles '" & psPersCod & "'"
        
        Set prsPatInmuebles = oConecta.CargaRecordSet(sSql)

        '******************** Maximo Patrimonio Vehicular ************
        sSql = "SELECT MAX(nNumPatri) as nMaxPatInmuebles FROM PERSPATRIInmuebles WHERE cPersCod = '" & psPersCod & "'"
        Set R = oConecta.CargaRecordSet(sSql)
    
        If Not R.EOF And Not R.BOF Then
            pnMaxPatInmuebles = IIf(IsNull(R!nMaxPatInmuebles), 0, R!nMaxPatInmuebles)
        Else
            pnMaxPatInmuebles = 0
        End If
        
        '********CTI3 *****DATOS ADICIIONALES
        sSql = "SELECT * FROM PersonaJurDetalles WHERE cPersCod = '" & psPersCod & "'"
        Set prsDatAdc = oConecta.CargaRecordSet(sSql)
        '*************
        
        oConecta.CierraConexion
        Set R = Nothing
        Set oConecta = Nothing
        Exit Function
    
ErrorRecuperarDatosPersona:
    Set R = Nothing
    Set oConecta = Nothing
    Err.Raise Err.Number, "Recuperar Datos Persona", Err.Description
End Function

'**********************************************
' Funcion que almacena los Datos de la Persona
'**********************************************
Public Function GrabarDatosPersona(ByRef psPersCod As String, _
                                    ByVal psCodAge As String, ByVal psCodCmact As String, _
                                    ByVal pdFecha As Date, ByVal psCoduser As String, _
                                    ByVal pnTipoAct As Integer, ByVal prsPersona As ADODB.Recordset, _
                                    ByVal prsFirma As ADODB.Recordset, ByVal prsDocs As ADODB.Recordset, _
                                    ByVal prsRelac As ADODB.Recordset, ByVal prsFteIng As ADODB.Recordset, _
                                    ByVal prsRefCom As ADODB.Recordset, ByVal prsFIDep As ADODB.Recordset, _
                                    ByVal prsRefBan As ADODB.Recordset, ByVal prsPatVeh As ADODB.Recordset, _
                                    ByVal prsFIInd As ADODB.Recordset, ByVal prsFICosto As ADODB.Recordset, _
                                    ByVal prsDatosAdc As ADODB.Recordset, _
                                    ByVal psDoc0 As String, ByVal psDoc1 As String, _
                                    ByVal psDoc2 As String, ByVal psDoc3 As String, _
                                    Optional prsHojEval As ADODB.Recordset, Optional nIntValor As Integer = 0, _
                                    Optional prsPatOtros As ADODB.Recordset, Optional prsPatInmuebles As ADODB.Recordset, _
                                    Optional prsVisitas As ADODB.Recordset, _
                                    Optional prsVentas As ADODB.Recordset, _
                                    Optional prsEstFinan As Recordset, _
                                    Optional pDatAccExiste As Boolean) As Boolean 'MAVM 20100607 BAS II   ****CTI3 07092018
'WIOR 20140321 AGREGO prsEstFinan
'add prsVisitas por MADM 20100327
Dim sSql As String
Dim R As ADODB.Recordset
Dim bPersDistrib As Boolean
Dim oConexDitrib As COMConecta.DCOMConecta   'ADODB.Connection

Dim oConecta  As COMConecta.DCOMConecta

Dim oGen As COMDConstSistema.DCOMGeneral
Dim sUltimaActualizacion As String
'****************************************************************************
Dim sCodPersDistrib As String
Dim lsUltCorr As String
Dim lnCorr As Long
Dim lnCodchequeo As Integer
Dim lsCadFinal As String
Dim sNumFuente As String
Dim sNuevo As String
'Para manejar la Actualizacion de varias Fuentes de Ingreso
Dim sFteTemp As String
Dim sFtesDep() As String
Dim sFtesInd() As String
Dim nFteInd As Integer
Dim nFteDep As Integer

'************************************************************************
'PARA DISTRIBUIDO
'************************************************************************

Dim sPersRelIDistrib As String
Dim TranDistribBegin As Boolean
Dim vDocs(4) As String
Dim sPersRelInst As String
'************************************************************************
Dim dfecha As Date

TranDistribBegin = False
sPersRelIDistrib = ""
    Select Case sPersRelInst
        Case "1"
            sPersRelIDistrib = "D"
        Case "2"
            sPersRelIDistrib = "F"
        Case "3"
            sPersRelIDistrib = "N"
        Case "4"
            sPersRelIDistrib = "O"
        Case "5"
            sPersRelIDistrib = "T"
    End Select

    On Error GoTo ErrorGrabarDatosPersona



    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    '**********************   Para Actualizar el Distribuido ************************
    sSql = "Select nConsSisValor from ConstSistema where nConsSisCod=45"
    Set R = New ADODB.Recordset

    Set R = oConecta.CargaRecordSet(sSql)

    If R.RecordCount > 0 Then
        If Len(Trim(R!nConsSisValor)) > 0 Then
            bPersDistrib = True
            Set oConexDitrib = New COMConecta.DCOMConecta  ' ADODB.Connection
            oConexDitrib.AbreConexion ("Provider=SQLOLEDB.1;User ID=dbaccess;PWD=cmact;Initial Catalog=DBPersona;Data Source=" & R!nConsSisValor & ";")
        Else
            bPersDistrib = False
        End If
    End If
    R.Close

    '************************************************************************************
    'OBTENER CODIGO DE FUENTE DE INGRESO
    sSql = "Select MAX(cNumFuente) as cNumFuente From PersFteIngreso"
    Set R = New ADODB.Recordset
    Set R = oConecta.CargaRecordSet(sSql)

    sNumFuente = Format(CLng(IIf(IsNull(R!cNumFuente), 0, R!cNumFuente)), "00000000")
    'sFteTemp = sNumFuente
    R.Close
    '************************************************************************************

    oConecta.BeginTrans
    '****************************** Actualizacion
    Set oGen = New COMDConstSistema.DCOMGeneral
    sUltimaActualizacion = GeneraMovNroActualiza(pdFecha, psCoduser, psCodCmact, psCodAge)

    '*********************************


    '**********************   Para Actualizar el Distribuido ************************
    If pnTipoAct = 2 Then '2 Then
        'Set oNegPers = New NCOMPersona
        lsUltCorr = UltimoCorrelativo(psCodAge)
        lnCorr = Val(lsUltCorr) + 1
        lsUltCorr = FillNum(Trim(Str(lnCorr)), 7, "0")
        lnCodchequeo = DigitoChequeo(lsUltCorr)
        lsCadFinal = lsUltCorr & Trim(Str(lnCodchequeo))
        sCodPersDistrib = Right(Trim(psCodAge), 2) & Trim(lsCadFinal)
        psPersCod = Trim(psCodCmact) & Trim(psCodAge) & Trim(lsCadFinal) '
        'Set oNegPers = Nothing
    Else
        sCodPersDistrib = Right(psPersCod, 10)
    End If
    '********************************************************************************
    Set R = Nothing
    
    '** Juez 20120326 Para ingregar el codigo Otros (116) y No Declara (36)
    If prsPersona("nPersPersoneria") <> gPersonaNat Then
        prsPersona("sActiGiro1") = ""
        prsPersona("nPersCargo") = 0
    End If
    '** End Juez ***************************************
    If pnTipoAct = 1 Then '1 Then
        'Actualiza Persona
'*** PEAC 20080616
'        sSql = "UPDATE " & gConsPersona & "Persona SET "
'        sSql = sSql & "cPersNombre = '" & prsPersona("cPersNombre") & "',"
'        sSql = sSql & "dPersNacCreac = '" & Format(prsPersona("dPersNacCreac"), "yyyymmdd") & "',"
'        sSql = sSql & "cPersDireccUbiGeo = '" & prsPersona("cPersDireccUbiGeo") & "',"
'        sSql = sSql & "cPersDireccDomicilio = '" & prsPersona("cPersDireccDomicilio") & "',"
'        sSql = sSql & "cPersDireccCondicion = '" & prsPersona("cPersDireccCondicion") & "',"
'        sSql = sSql & "nPersValComDomicilio = " & prsPersona("cnPersValComDomicilio") & ","
'        sSql = sSql & "cPersTelefono = '" & prsPersona("cPersTelefono") & "',"
'        sSql = sSql & "cPersTelefono2 = '" & prsPersona("cPersTelefono2") & "',"
'        sSql = sSql & "cPersEmail = '" & prsPersona("cPersEmail") & "',"
'        sSql = sSql & "nPersPersoneria = '" & Trim(Str(prsPersona("nPersPersoneria"))) & "',"
'        sSql = sSql & "cPersCIIU = '" & prsPersona("cPersCIIU") & "',"
'        sSql = sSql & "cPersEstado = '" & prsPersona("cPersEstado") & "',"
'        sSql = sSql & "cPersCodSbs = '" & prsPersona("cPersCodSbs") & "',"
'        sSql = sSql & "nPersRelaInst = " & prsPersona("nPersRelaInst") & ","
'        '*** PEAC 20080412
'        sSql = sSql & "dPersIngRuc = '" & Format(prsPersona("dPersIngRuc"), "yyyymmdd") & "',"
'        sSql = sSql & "dPersIniActi = '" & Format(prsPersona("dPersIniActi"), "yyyymmdd") & "',"
'        sSql = sSql & "nNumDependi = " & prsPersona("nNumDependi") & ","
'        sSql = sSql & "cActComple = '" & prsPersona("cActComple") & "',"
'        sSql = sSql & "nNumPtosVta = " & prsPersona("nNumPtosVta") & ","
'        sSql = sSql & "cActiGiro = '" & prsPersona("cActiGiro") & "',"
'        '----
'        sSql = sSql & "nPersTipCompe = " & prsPersona("nPersTipoComp") & ","
'        sSql = sSql & "nPersTipSistInform = " & prsPersona("nPersTipoSistInfor") & ","
'        sSql = sSql & "nPersTipCadeProd = " & prsPersona("nPersCadenaProd") & ","
'        sSql = sSql & "nPersMoneyPatri = " & prsPersona("nMonedaPatri") & ","
'        '*** FIN PEAC
'        'CUSCO
'        sSql = sSql & "nPersIngresoProm=" & prsPersona("nPersIngresoProm")
'        sSql = sSql & " Where cPersCod = '" & psPersCod & "'"


    '************************************************************************************
        
        
        '*** PEAC 20080616
        'EJVG20110912 Se agregó cPersCelular,cPersCelular2,cPersCelular3,cPersEmail2
        'EJVG20120813 Se agregó nSujetoObligado y nOfCumplimiento
        'JUEZ 20131007 cPersNegocioUbiGeo,cPersNegocioDireccion,cPersNegocioRef,nRemisionInfoEmail
        sSql = "Exec stp_upd_ActualizaPersona_nuevo '" _
                        & prsPersona("cPersNombre") & "','" & Format(prsPersona("dPersNacCreac"), "yyyymmdd") & "','" _
                        & prsPersona("cPersDireccUbiGeo") & "','" & prsPersona("cPersDireccDomicilio") & "','" _
                        & prsPersona("cPersDireccCondicion") & "'," & prsPersona("cnPersValComDomicilio") & ",'" _
                        & prsPersona("cPersTelefono") & "','" & prsPersona("cPersTelefono2") & "','" _
                        & prsPersona("cPersCelular") & "','" & prsPersona("cPersCelular2") & "','" & prsPersona("cPersCelular3") & "','" _
                        & prsPersona("cPersEmail") & "','" & prsPersona("cPersEmail2") & "','" & Trim(Str(prsPersona("nPersPersoneria"))) & "','" _
                        & prsPersona("cPersCIIU") & "','" & prsPersona("cPersEstado") & "','" _
                        & prsPersona("cPersCodSbs") & "'," & prsPersona("nPersRelaInst") & ",'" _
                        & Format(prsPersona("dPersIngRuc"), "yyyymmdd") & "','" & Format(prsPersona("dPersIniActi"), "yyyymmdd") & "'," _
                        & prsPersona("nNumDependi") & ",'" & prsPersona("cActComple") & "'," _
                        & prsPersona("nNumPtosVta") & ",'" & prsPersona("cActiGiro") & "'," _
                        & prsPersona("nPersTipoComp") & "," & prsPersona("nPersTipoSistInfor") & "," _
                        & prsPersona("nPersCadenaProd") & "," & prsPersona("nMonedaPatri") & "," _
                        & prsPersona("nPersIngresoProm") & ",'" & psPersCod & "','" & sUltimaActualizacion & "','" & prsPersona("cMotivoActu") & "','" & prsPersona("cPersRefDomicilio") & "','" _
                        & prsPersona("sActiGiro1") & "'," _
                        & prsPersona("nPersCargo") & "," & prsPersona("nSujetoObligado") & "," & prsPersona("nOfCumplimiento") & ",'" _
                        & prsPersona("cPersNegocioUbiGeo") & "','" & prsPersona("cPersNegocioDireccion") & "','" & prsPersona("cPersNegocioRef") & "','" & prsPersona("cPersNombreCentroLaboral") & "'," & prsPersona("nRemisionInfoEmail") 'marg 12-05-16 :cPersNombreCentroLaboral
                        
        oConecta.ConexionActiva.Execute sSql

        If prsPersona("nPersPersoneria") = gPersonaNat Then

            sSql = "DELETE " & gConsPersona & "PersonaJur"
            sSql = sSql & " WHERE cPersCod = '" & psPersCod & "'"
            oConecta.ConexionActiva.Execute sSql

            sSql = "DELETE " & gConsPersona & "PersonaNat"
            sSql = sSql & " WHERE cPersCod = '" & psPersCod & "'"
            oConecta.ConexionActiva.Execute sSql

            'Actualiza Persona Natural
            sSql = "INSERT INTO " & gConsPersona
            sSql = sSql & "PersonaNat(cPersCod,cPersnatSexo,nPersNatEstCiv,nPersNatHijos,nPersNatNumEmp,nPersPeso,"
            'MAVM 06042009 Se añadio el campo dPersFallec
            'MADM 20091113 Se añadieron relacion CMAC
            'sSql = sSql & "nPersTalla,nPersTpoSangre, cNacionalidad, nResidente, cUltimaActualizacion, dPersFallec, nCredito, nAhorro, nServicio, nOtro, nPeps)" modificado x jaca para añadir el campo cPersNatMagnitud
            'sSql = sSql & "nPersTalla,nPersTpoSangre, cNacionalidad, nResidente, cUltimaActualizacion, dPersFallec, nCredito, nAhorro, nServicio, nOtro, nPeps,cPersNatMagnitud)" 'JACA 20110427 Comentado x Juez para añadir el campo cPaisReside
            sSql = sSql & "nPersTalla,nPersTpoSangre, cNacionalidad, nResidente, cUltimaActualizacion, dPersFallec, nCredito, nAhorro, nServicio, nOtro, nPeps,cPersNatMagnitud,cPaisReside)" '** Juez 20120327
            sSql = sSql & " VALUES('" & psPersCod & "',"
            sSql = sSql & "'" & prsPersona("cPersnatSexo") & "',"
            sSql = sSql & prsPersona("nPersNatEstCiv") & ","
            sSql = sSql & Trim(Str(prsPersona("nPersNatHijos"))) & ","
            '*** PEAC 20080412
            sSql = sSql & Trim(Str(prsPersona("nPersNatNumEmp"))) & ","
            sSql = sSql & "'" & Format(prsPersona("nPersPeso"), "#0.00") & "',"
            sSql = sSql & "'" & Format(prsPersona("nPersTalla"), "#0.00") & "',"
            sSql = sSql & IIf(prsPersona("nPersTpoSangre") = "", 0, prsPersona("nPersTpoSangre")) & ","
            sSql = sSql & "'" & prsPersona("cNacionalidad") & "',"
            sSql = sSql & prsPersona("nResidente") & ","
            'MAVM 06042009 Se añadio el valor prsPersona("dPersFallec")
            'sSql = sSql & "'" & sUltimaActualizacion & "'" & "," & "'" & Format(prsPersona("dPersFallec"), "yyyymmdd") & "')"
            'MADM 20090913
            sSql = sSql & "'" & sUltimaActualizacion & "'" & "," & "'" & Format(prsPersona("dPersFallec"), "yyyymmdd") & "',"
            sSql = sSql & prsPersona("nCredito") & ","
            sSql = sSql & prsPersona("nAhorro") & ","
            sSql = sSql & prsPersona("nServicio") & ","
            sSql = sSql & prsPersona("nOtro") & ","
            'JACA 20110427*******************************************************************
            'sSql = sSql & prsPersona("nPeps") & ")"
            sSql = sSql & prsPersona("nPeps") & ","
            sSql = sSql & "'" & prsPersona("cPersNatMagnitud") & "',"
            'JACA END************************************************************************
            'END MADM
            sSql = sSql & "'" & prsPersona("cPaisReside") & "')" '** Juez 20120327
            oConecta.ConexionActiva.Execute sSql

        End If

        If prsPersona("nPersPersoneria") <> gPersonaNat Then
            sSql = "DELETE " & gConsPersona & "PersonaNat"
            sSql = sSql & " WHERE cPersCod = '" & psPersCod & "'"
            oConecta.ConexionActiva.Execute sSql

            sSql = "DELETE " & gConsPersona & "PersonaJur"
            sSql = sSql & " WHERE cPersCod = '" & psPersCod & "'"
            oConecta.ConexionActiva.Execute sSql

            'Actualiza Persona Juridica
            'Inserta PersonaJur
            sSql = "INSERT INTO " & gConsPersona
            sSql = sSql & "PersonaJur(cPersCod,cPersJurTpo,cPersJurMagnitud,cPersJurSigla,nPersJurEmpleados,cPersJurObjetoSocial,cUltimaActualizacion)" '** Juez 20120328 Se agregó cPersJurObjetoSocial
            sSql = sSql & "VALUES('" & psPersCod & "',"
            sSql = sSql & "'" & prsPersona("cPersJurTpo") & " ',"
            sSql = sSql & "'" & prsPersona("cPersJurMagnitud") & " ',"
            sSql = sSql & "'" & prsPersona("cPersJurSigla") & " ',"
            sSql = sSql & Trim(Str(prsPersona("nPersJurEmpleados"))) & ","
            sSql = sSql & "'" & prsPersona("cPersJurObjetoSocial") & "'," '** Juez 20120328
            sSql = sSql & "'" & sUltimaActualizacion & "')"
            oConecta.ConexionActiva.Execute sSql
            
        End If
        'Actualiza Firma
        'If Not prsFirma Is Nothing Then
        '    If Not prsFirma.BOF And Not prsFirma.EOF Then
        '        prsFirma!cUltimaActualizacion = sUltimaActualizacion
        '        prsFirma.Update
        '    End If
        'End If

'SE AGREGO EN EL INSERT LOS CAMPOS "nPersTipCompe,nPersTipSistInform,nPersTipCadeProd" PEAC 20080412

    '*** CTI3 ******** 07092018 *********Datos Adicionales Juridicos y Participacion patrimonial ****Modificar Datos
    If pDatAccExiste = True Then
            sSql = "DELETE PersonaJurDetalles"
            sSql = sSql & " WHERE cPersCod = '" & psPersCod & "'"
            oConecta.ConexionActiva.Execute sSql
        If prsDatosAdc.RecordCount > 0 Then
            'prsDatosAdc.MoveFirst
            While Not prsDatosAdc.EOF
            
            sSql = "INSERT INTO PersonaJurDetalles(cPersCod,cNombreJur,vRuc,vCi,cNacionalidad,cCargo,"
            sSql = sSql & "cRazonSocial,nApoSol,nApoPorc,vCodSBS,cTipo)"
            sSql = sSql & "Values('" & psPersCod & "','" & prsDatosAdc("sNombrejur") & "','" & prsDatosAdc("sRUC") & "',"
            sSql = sSql & "'" & prsDatosAdc("sCi") & "','" & prsDatosAdc("sNacionalidad") & "','" & prsDatosAdc("sCargo") & "',"
            sSql = sSql & "'" & prsDatosAdc("sRazonSocial") & "'," & prsDatosAdc("nApoSol") & ","
            sSql = sSql & "" & prsDatosAdc("nApoPorc") & ",'" & prsDatosAdc("sCodSbs") & "'," & prsDatosAdc("sTipo") & ")"
            oConecta.ConexionActiva.Execute sSql
    
            prsDatosAdc.MoveNext
            Wend
        End If
    End If
    '*******************************************************************************************

    Else 'Para Nueva Persona
        If pnTipoAct = 2 Then
            'Insertar en Tabla Persona          ***CUSCO
            sSql = "INSERT INTO " & gConsPersona
            sSql = sSql & "Persona(cPersCod,cPersNombre,dPersNacCreac,cPersDireccUbiGeo,cPersDireccDomicilio"
            'sSql = sSql & ",cPersDireccCondicion, nPersValComDomicilio, cPersTelefono,cPersTelefono2,nPersPersoneria,cPersCIIU,cPersEstado"
            sSql = sSql & ",cPersDireccCondicion, nPersValComDomicilio, cPersTelefono,cPersTelefono2,cPersCelular,cPersCelular2,cPersCelular3,nPersPersoneria,cPersCIIU,cPersEstado" 'EJVG20110912
            'sSql = sSql & ",cPersEmail,dPersIngRuc,dPersIniActi,nNumDependi,cActComple,nNumPtosVta,cActiGiro,nPersTipCompe,nPersTipSistInform,nPersTipCadeProd,nPersMoneyPatri,cTipoActualizacion,cPersRefDomicilio,cUltimaActualizacion, cPersCodSbs, nPersRelaInst,dPersIng,cActiGiro1,nPersIngresoProm)"
            'sSql = sSql & ",cPersEmail,cPersEmail2,dPersIngRuc,dPersIniActi,nNumDependi,cActComple,nNumPtosVta,cActiGiro,nPersTipCompe,nPersTipSistInform,nPersTipCadeProd,nPersMoneyPatri,cTipoActualizacion,cPersRefDomicilio,cUltimaActualizacion, cPersCodSbs, nPersRelaInst,dPersIng,cActiGiro1,nPersIngresoProm)"
            'EJVG20120813 ***
            'sSql = sSql & ",cPersEmail,cPersEmail2,dPersIngRuc,dPersIniActi,nNumDependi,cActComple,nNumPtosVta,cActiGiro,nPersTipCompe,nPersTipSistInform,nPersTipCadeProd,nPersMoneyPatri,cTipoActualizacion,cPersRefDomicilio,cUltimaActualizacion, cPersCodSbs, nPersRelaInst,dPersIng,cActiGiro1,nPersIngresoProm,nPersCargo)" '** Juez 20120326
            sSql = sSql & ",cPersEmail,cPersEmail2,dPersIngRuc,dPersIniActi,nNumDependi,cActComple,nNumPtosVta,cActiGiro,nPersTipCompe,nPersTipSistInform,nPersTipCadeProd,nPersMoneyPatri,cTipoActualizacion,cPersRefDomicilio,cUltimaActualizacion, cPersCodSbs, nPersRelaInst,dPersIng,cActiGiro1,nPersIngresoProm,nPersCargo"
            sSql = sSql & ",nSujetoObligado,nOfCumplimiento"
            sSql = sSql & ",cPersNegocioUbiGeo,cPersNegocioDireccion,cPersNegocioRef,cPersNombreCentroLaboral, nRemisionInfoEmail" 'JUEZ 20131007 marg 12-05-2016
            sSql = sSql & ")"
            'END EJVG ******
            sSql = sSql & "VALUES('" & psPersCod & "',"
            sSql = sSql & "'" & prsPersona("cPersNombre") & "',"
            sSql = sSql & "'" & Format(prsPersona("dPersNacCreac"), "yyyymmdd") & "',"
            sSql = sSql & "'" & prsPersona("cPersDireccUbiGeo") & "',"
            sSql = sSql & "'" & prsPersona("cPersDireccDomicilio") & "',"
            sSql = sSql & "'" & prsPersona("cPersDireccCondicion") & "',"
            sSql = sSql & prsPersona("cnPersValComDomicilio") & ","
            sSql = sSql & "'" & prsPersona("cPersTelefono") & "',"
            sSql = sSql & "'" & prsPersona("cPersTelefono2") & "',"
            'EJVG20111209 ********************************************
            sSql = sSql & "'" & prsPersona("cPersCelular") & "',"
            sSql = sSql & "'" & prsPersona("cPersCelular2") & "',"
            sSql = sSql & "'" & prsPersona("cPersCelular3") & "',"
            '*********************************************************
            sSql = sSql & Trim(Str(prsPersona("nPersPersoneria"))) & ","
            sSql = sSql & "'" & prsPersona("cPersCIIU") & "',"
            sSql = sSql & "'" & prsPersona("cPersEstado") & "',"
            sSql = sSql & "'" & prsPersona("cPersEmail") & "',"
            sSql = sSql & "'" & prsPersona("cPersEmail2") & "'," 'EJVG20111209
            '*** PEAC 20080412
            sSql = sSql & "'" & Format(prsPersona("dPersIngRuc"), "yyyymmdd") & "',"
            sSql = sSql & "'" & Format(prsPersona("dPersIniActi"), "yyyymmdd") & "',"
            sSql = sSql & "" & prsPersona("nNumDependi") & ","
            sSql = sSql & "'" & prsPersona("cActComple") & "',"
            sSql = sSql & "" & prsPersona("nNumPtosVta") & ","
            sSql = sSql & "'" & prsPersona("cActiGiro") & "',"
            '-----
            sSql = sSql & "" & prsPersona("nPersTipoComp") & ","
            sSql = sSql & "" & prsPersona("nPersTipoSistInfor") & ","
            sSql = sSql & "" & prsPersona("nPersCadenaProd") & ","
            sSql = sSql & "" & prsPersona("nMonedaPatri") & ","
            '*** FIN PEAC
            
            '*** PEAC 20080801
            sSql = sSql & "'" & prsPersona("cMotivoActu") & "',"
            sSql = sSql & "'" & prsPersona("cPersRefDomicilio") & "',"
            
            sSql = sSql & "'" & sUltimaActualizacion & "','" & prsPersona("cPersCodSbs") & "'," & prsPersona("nPersRelaInst") & ",'" & Format(pdFecha, "mm/dd/yyyy") & "',"
            sSql = sSql & "'" & prsPersona("sActiGiro1") & "'," & prsPersona("nPersIngresoProm") & ","
            'comentado x madm 20100322 sSql = sSql & prsPersona("nPersIngresoProm") & ")"
            'EJVG20120813 ***
            'sSql = sSql & "'" & prsPersona("nPersCargo") & "')" '** Juez 20120326
            sSql = sSql & "'" & prsPersona("nPersCargo") & "'"
            sSql = sSql & "," & prsPersona("nSujetoObligado") & "," & prsPersona("nOfCumplimiento")
            sSql = sSql & ",'" & prsPersona("cPersNegocioUbiGeo") & "','" & prsPersona("cPersNegocioDireccion") & "','" & prsPersona("cPersNegocioRef") & "','" & prsPersona("cPersNombreCentroLaboral") & "'," & prsPersona("nRemisionInfoEmail") 'JUEZ 20131007
            sSql = sSql & ")"
            'END EJVG *******
            oConecta.ConexionActiva.Execute sSql

            If prsPersona("nPersPersoneria") = gPersonaNat Then
                'Insertar en PersonaNat
                sSql = "INSERT INTO " & gConsPersona
                sSql = sSql & "PersonaNat(cPersCod,cPersnatSexo,nPersNatEstCiv,nPersNatHijos,nPersNatNumEmp,nPersPeso,"
                sSql = sSql & "nPersTalla,nPersTpoSangre, cNacionalidad,"
                'sSql = sSql & "nCredito, nAhorro, nServicio, nOtro,nPeps,nResidente,cUltimaActualizacion)" 'madm 20091110 comentado x jaca cPersNatMagnitud
                'sSql = sSql & "nCredito, nAhorro, nServicio, nOtro,nPeps,nResidente,cUltimaActualizacion,cPersNatMagnitud)" 'jaca 20110427 Comentado x Juez para agregar campo cPaisReside
                sSql = sSql & "nCredito, nAhorro, nServicio, nOtro,nPeps,nResidente,cUltimaActualizacion,cPersNatMagnitud,cPaisReside)" '** Juez 20120327
                sSql = sSql & " VALUES('" & psPersCod & "',"
                sSql = sSql & "'" & prsPersona("cPersnatSexo") & "',"
                sSql = sSql & prsPersona("nPersNatEstCiv") & ","
                sSql = sSql & Trim(Str(prsPersona("nPersNatHijos"))) & ","
                sSql = sSql & Trim(Str(prsPersona("nPersNatNumEmp"))) & ","
                sSql = sSql & "'" & Format(prsPersona("nPersPeso"), "#0.00") & "',"
                sSql = sSql & "'" & Format(prsPersona("nPersTalla"), "#0.00") & "',"
                sSql = sSql & IIf(prsPersona("nPersTpoSangre") = "", 0, prsPersona("nPersTpoSangre")) & ","
                sSql = sSql & "'" & prsPersona("cNacionalidad") & "',"
                  'MADM 20091411
                sSql = sSql & prsPersona("nCredito") & ","
                sSql = sSql & prsPersona("nAhorro") & ","
                sSql = sSql & prsPersona("nServicio") & ","
                sSql = sSql & prsPersona("nOtro") & ","
                sSql = sSql & prsPersona("nPeps") & ","     '20100408
                ' END MADM
                sSql = sSql & prsPersona("nResidente") & ","
                'JACA 20110427********************************************************
                'sSql = sSql & "'" & sUltimaActualizacion & "')"
                sSql = sSql & "'" & sUltimaActualizacion & "',"
                sSql = sSql & "'" & prsPersona("cPersNatMagnitud") & "',"
                'JACA END************************************************************
                sSql = sSql & "'" & prsPersona("cPaisReside") & "')" '** Juez 20120327
                oConecta.ConexionActiva.Execute sSql
            End If

            If prsPersona("nPersPersoneria") <> gPersonaNat Then
                'Inserta PersonaJur
                sSql = "INSERT INTO " & gConsPersona
                sSql = sSql & "PersonaJur(cPersCod,cPersJurTpo,cPersJurMagnitud,cPersJurSigla,nPersJurEmpleados,cPersJurObjetoSocial,cUltimaActualizacion)" '** Juez 20120328 Se agregó cPersJurObjetoSocial
                sSql = sSql & "VALUES('" & psPersCod & "',"
                sSql = sSql & "'" & prsPersona("cPersJurTpo") & " ',"
                sSql = sSql & "'" & prsPersona("cPersJurMagnitud") & " ',"
                sSql = sSql & "'" & prsPersona("cPersJurSigla") & " ',"
                sSql = sSql & Trim(Str(prsPersona("nPersJurEmpleados"))) & ","
                sSql = sSql & "'" & prsPersona("cPersJurObjetoSocial") & "'," '** Juez 20120328
                sSql = sSql & "'" & sUltimaActualizacion & "')"
                oConecta.ConexionActiva.Execute sSql
               
            End If
            
            
    '*** CTI3 ******** 07092018 *********Datos Adicionales Juridicos y Participacion patrimonial ****Modificar Datos
    If pDatAccExiste = True Then
            sSql = "DELETE PersonaJurDetalles"
            sSql = sSql & " WHERE cPersCod = '" & psPersCod & "'"
            oConecta.ConexionActiva.Execute sSql
        If prsDatosAdc.RecordCount > 0 Then

            prsDatosAdc.MoveFirst
            While Not prsDatosAdc.EOF
            
            sSql = "INSERT INTO PersonaJurDetalles(cPersCod,cNombreJur,vRuc,vCi,cNacionalidad,cCargo,"
            sSql = sSql & "cRazonSocial,nApoSol,nApoPorc,vCodSBS,cTipo)"
            sSql = sSql & "Values('" & psPersCod & "','" & prsDatosAdc("sNombrejur") & "','" & prsDatosAdc("sRUC") & "',"
            sSql = sSql & "'" & prsDatosAdc("sCi") & "','" & prsDatosAdc("sNacionalidad") & "','" & prsDatosAdc("sCargo") & "',"
            sSql = sSql & "'" & prsDatosAdc("sRazonSocial") & "'," & prsDatosAdc("nApoSol") & ","
            sSql = sSql & "" & prsDatosAdc("nApoPorc") & ",'" & prsDatosAdc("sCodSbs") & "'," & prsDatosAdc("sTipo") & ")"
            oConecta.ConexionActiva.Execute sSql
            
            prsDatosAdc.MoveNext
            Wend
        End If
    End If
    '*******************************************************************************************
            
            'OJO : 10-12    Chequear Firma
            'Actualiza Firma
            If Not prsFirma Is Nothing Then
                If prsFirma.RecordCount > 0 Then
                    prsFirma!cPersCod = psPersCod
                    prsFirma!cUltimaActualizacion = sUltimaActualizacion
                    prsFirma.Update
                End If
            End If

        End If
    End If

    'Graba Relaciones de la Persona
    While Not prsRelac.EOF
        If prsRelac("nTipoAct") = 1 Then
            sSql = "UPDATE " & gConsPersona & "PersRelaciones SET "
            sSql = sSql & "cPersRelacPersCod = '" & prsRelac("cPersRelacPersCod") & "',"
            sSql = sSql & "nPersRelac = " & prsRelac("nPersRelac") & ","
            sSql = sSql & "bPersRelacBenef = " & prsRelac("bPersRelacBenef") & ","
            sSql = sSql & "nPersRelacBenefPorc = " & prsRelac("nPersRelacBenefPorc") & ","
            sSql = sSql & "bPersRelacAMP = " & prsRelac("bPersRelacAMP")
            'sSql = sSql & "PR.bPersRelacEstado = " & Trim(Right(Relac(i).sPersRelacEstado, 20))
            sSql = sSql & " Where cPersCod = '" & psPersCod & "' And cPersRelacPersCod = '" & prsRelac("cPersCodOrig") & "'"
            oConecta.ConexionActiva.Execute sSql
        Else
            If prsRelac("nTipoAct") = 2 Then
                sSql = "INSERT INTO " & gConsPersona & "PERSRELACIONES(cPersCod,cPersRelacPersCod,"
                sSql = sSql & "nPersRelac,bPersrelacBenef,nPersRelacBenefPorc,bPersRelacAMP,cMovNro)"
                sSql = sSql & "VALUES('" & psPersCod & "',"
                sSql = sSql & "'" & prsRelac("cPersRelacPersCod") & "',"
                sSql = sSql & prsRelac("nPersRelac") & ","
                sSql = sSql & prsRelac("bPersRelacBenef") & ","
                sSql = sSql & prsRelac("nPersRelacBenefPorc") & ","
                sSql = sSql & prsRelac("bPersRelacAMP") & ","
                sSql = sSql & "'" & prsRelac("cMovNro") & "')"
                oConecta.ConexionActiva.Execute sSql
            Else
                If prsRelac("nTipoAct") = 3 Then
                    sSql = "DELETE " & gConsPersona & "PersRelaciones Where "
                    sSql = sSql & " cPersCod = '" & psPersCod & "'"
                    sSql = sSql & " And cPersRelacPersCod = '" & Trim(prsRelac("cPersCodOrig")) & "'"
                    oConecta.ConexionActiva.Execute sSql
                End If
            End If
        End If
        prsRelac.MoveNext
    Wend

    'add visitas personas madm 20100327
    While Not prsVisitas.EOF
        If prsVisitas("nTipoAct") = 1 Then
            sSql = "UPDATE " & gConsPersona & "persVisitas SET "
            sSql = sSql & "cPersRelacPersCod = '" & prsVisitas("cPersRelacPersCod") & "',"
            sSql = sSql & "cDireccion = '" & prsVisitas("cDireccion") & "',"
            sSql = sSql & "dFecVis = '" & Format(prsVisitas("dFecVis"), "yyyymmdd") & "',"
            sSql = sSql & "cVisObserva = '" & prsVisitas("cVisObserva") & "',"
            sSql = sSql & "sUsual = " & prsVisitas("nUsual")
            sSql = sSql & " Where cPersCod = '" & psPersCod & "' And cPersRelacPersCod = '" & prsVisitas("cPersRelacPersCod") & "'"
            oConecta.ConexionActiva.Execute sSql
        Else
            If prsVisitas("nTipoAct") = 2 Then
                sSql = "INSERT INTO " & gConsPersona & "persVisitas(cPersCod,cPersRelacPersCod,cDireccion,"
                sSql = sSql & "dFecVis,cVisObserva,sUsual)"
                sSql = sSql & " VALUES ('" & psPersCod & "',"
                sSql = sSql & "'" & prsVisitas("cPersRelacPersCod") & "',"
                sSql = sSql & "'" & prsVisitas("cDireccion") & "',"
                sSql = sSql & "'" & Format(prsVisitas("dFecVis"), "yyyymmdd") & "',"
                sSql = sSql & "'" & prsVisitas("cVisObserva") & "',"
                sSql = sSql & "" & prsVisitas("nUsual") & ") "
                oConecta.ConexionActiva.Execute sSql
            Else
                If prsVisitas("nTipoAct") = 3 Then
                    sSql = "DELETE " & gConsPersona & "persVisitas Where "
                    sSql = sSql & " cPersCod = '" & psPersCod & "'"
                    sSql = sSql & " And cPersRelacPersCod = '" & Trim(prsVisitas("cPersRelacPersCod")) & "'"
                    oConecta.ConexionActiva.Execute sSql
                End If
            End If
        End If
        prsVisitas.MoveNext
    Wend
'end madm

    'Ventas Anuales
    While Not prsVentas.EOF
        If prsVentas("nTipoActVA") = 1 Then
            sSql = "UPDATE " & gConsPersona & "PersonaJurVentaAnual SET "
            sSql = sSql & "nPeriodo = '" & prsVentas("cPeriodoVA") & "', "
            sSql = sSql & "nMontoVA = '" & prsVentas("nMontoVA") & "', "
            sSql = sSql & "cPersCodAudEEFF = '" & prsVentas("cPersCodVA") & "', "
            sSql = sSql & "dFecAudEEFF = '" & prsVentas("dFecAudVA") & "' "
            sSql = sSql & "Where cPersCod = '" & psPersCod & "' "
            sSql = sSql & "And cPersCodAudEEFF = '" & Trim(prsVentas("cPersCodOrigVA")) & "'"
            oConecta.ConexionActiva.Execute sSql
        Else
            If prsVentas("nTipoActVA") = 2 Then
                sSql = "INSERT INTO " & gConsPersona & "PersonaJurVentaAnual(cPersCod, nPeriodo, nMontoVA, cPersCodAudEEFF, dFecAudEEFF)"
                sSql = sSql & " VALUES ('" & psPersCod & "',"
                sSql = sSql & "'" & prsVentas("cPeriodoVA") & " ',"
                sSql = sSql & "'" & prsVentas("nMontoVA") & " ',"
                sSql = sSql & "'" & prsVentas("cPersCodVA") & " ',"
                sSql = sSql & "'" & prsVentas("dFecAudVA") & "')"
                oConecta.ConexionActiva.Execute sSql
            Else
                If prsVentas("nTipoActVA") = 3 Then
                    sSql = "DELETE " & gConsPersona & "PersonaJurVentaAnual Where "
                    sSql = sSql & " cPersCod = '" & psPersCod & "'"
                    sSql = sSql & " And cPersCodAudEEFF = '" & Trim(prsVentas("cPersCodVA")) & "'"
                    oConecta.ConexionActiva.Execute sSql
                End If
            End If
        End If
        prsVentas.MoveNext
    Wend
    '***
    
    'Graba Referencias Comerciales
    While Not prsRefCom.EOF
        If prsRefCom("nTipoAct") = 1 Then
            sSql = "UPDATE PERSREFCOMERCIAL SET "
            sSql = sSql & "cNomRefCom= '" & prsRefCom("cNomRefCom") & "',"
            sSql = sSql & "nTipoRelaCom = " & prsRefCom("nTipoRelaCom") & ","
            sSql = sSql & "cFonoRefCom = '" & prsRefCom("cFonoRefCom") & "',"
            sSql = sSql & "cComentarioRefCom = '" & prsRefCom("cComentarioRefCom") & "',"
            sSql = sSql & "cDireccionRefCom = '" & prsRefCom("cDireccionRefCom") & "',"
            sSql = sSql & "cUltimaActualizacion = '" & sUltimaActualizacion & "' "
            sSql = sSql & " WHERE cPersCod = '" & psPersCod & "' AND nNumRefCom = " & prsRefCom("nRefComCod")
            oConecta.ConexionActiva.Execute sSql
        Else
            If prsRefCom("nTipoAct") = 2 Then
                sSql = "INSERT INTO PERSREFCOMERCIAL (cPersCod, nNumRefCom, cNomRefCom, nTipoRelaCom, "
                sSql = sSql & "cFonoRefCom,cComentarioRefCom,cDireccionRefCom, cUltimaActualizacion)"
                sSql = sSql & " VALUES ('" & psPersCod & "',"
                sSql = sSql & prsRefCom("nRefComCod") & ","
                sSql = sSql & "'" & prsRefCom("cNomRefCom") & "',"
                sSql = sSql & prsRefCom("nTipoRelaCom") & ","
                sSql = sSql & "'" & prsRefCom("cFonoRefCom") & "',"
                sSql = sSql & "'" & prsRefCom("cComentarioRefCom") & "',"
                sSql = sSql & "'" & prsRefCom("cDireccionRefCom") & "',"
                sSql = sSql & "'" & sUltimaActualizacion & "')"
                oConecta.ConexionActiva.Execute sSql
            Else
                If prsRefCom("nTipoAct") = 3 Then
                    sSql = "DELETE PERSREFCOMERCIAL "
                    sSql = sSql & " WHERE cPersCod = '" & psPersCod & "' AND nNumRefCom = " & prsRefCom("nRefComCod")
                    oConecta.ConexionActiva.Execute sSql
                End If
            End If
        End If
        prsRefCom.MoveNext
    Wend

    'Graba Referencias Bancarias
    While Not prsRefBan.EOF
        If prsRefBan("nTipoAct") = 1 Then
            sSql = "UPDATE PERSREFBANCARIA SET "
            sSql = sSql & "cCodBanco = '" & prsRefBan("cCodBanco") & "',"
            sSql = sSql & "cNumCta = '" & prsRefBan("cNumCta") & "',"
            sSql = sSql & "cNumTarjeta = '" & prsRefBan("cNumTarjeta") & "',"
            sSql = sSql & "nMontoLinCred = " & prsRefBan("nMontoLinCred") & ","
            sSql = sSql & "cUltimaActualizacion = '" & sUltimaActualizacion & "' "
            sSql = sSql & " WHERE cPersCod = '" & psPersCod & "' AND cCodBanco = '" & prsRefBan("sRefBanCodIFOrg") & "'"
            oConecta.ConexionActiva.Execute sSql
        Else
            If prsRefBan("nTipoAct") = 2 Then
                sSql = "INSERT INTO PERSREFBANCARIA (cPersCod, cCodBanco, cNumCta, cNumTarjeta, "
                sSql = sSql & "nMontoLinCred, cUltimaActualizacion)"
                sSql = sSql & " VALUES ('" & psPersCod & "',"
                sSql = sSql & "'" & prsRefBan("cCodBanco") & "',"
                sSql = sSql & "'" & prsRefBan("cNumCta") & "',"
                sSql = sSql & "'" & prsRefBan("cNumTarjeta") & "',"
                sSql = sSql & prsRefBan("nMontoLinCred") & ","
                sSql = sSql & "'" & sUltimaActualizacion & "')"
                oConecta.ConexionActiva.Execute sSql
            Else
                If prsRefBan("nTipoAct") = 3 Then
                    sSql = "DELETE PERSREFBANCARIA "
                    sSql = sSql & " WHERE cPersCod = '" & psPersCod & "' AND cCodBanco = '" & prsRefBan("sRefBanCodIFOrg") & "'"
                    oConecta.ConexionActiva.Execute sSql
                End If
            End If
        End If
        prsRefBan.MoveNext
    Wend

    'Graba Patrimonio Vehicular
    While Not prsPatVeh.EOF
        If prsPatVeh("nTipoAct") = 1 Then
            sSql = "UPDATE PERSPATRIMONIO SET "
            sSql = sSql & "cMarca= '" & prsPatVeh("cMarca") & "',"
            sSql = sSql & "cFecFabrica = '" & prsPatVeh("cFecFabrica") & "',"
            sSql = sSql & "nValorCom = " & prsPatVeh("nValorCom") & ","
            sSql = sSql & "cModelo = '" & prsPatVeh("cModelo") & "',"
            sSql = sSql & "cPlaca  = '" & prsPatVeh("cPlaca") & "',"
            sSql = sSql & "nCondPatri = " & prsPatVeh("nCondPatri") & ","
            sSql = sSql & "cUltimaActualizacion = '" & sUltimaActualizacion & "' "
            sSql = sSql & " WHERE cPersCod = '" & psPersCod & "' AND nNumPatri= " & prsPatVeh("nNumPatri")
            oConecta.ConexionActiva.Execute sSql
        Else
            If prsPatVeh("nTipoAct") = 2 Then
                sSql = "INSERT INTO PERSPATRIMONIO(cPersCod, nNumPatri, nCondPatri, cMarca,cModelo, cPlaca, cFecFabrica, "
                sSql = sSql & "nValorCom, cUltimaActualizacion)"
                sSql = sSql & " VALUES ('" & psPersCod & "',"
                sSql = sSql & prsPatVeh("nNumPatri") & ","
                sSql = sSql & prsPatVeh("nCondPatri") & ","
                sSql = sSql & "'" & prsPatVeh("cMarca") & "',"
                sSql = sSql & "'" & prsPatVeh("cModelo") & "',"
                sSql = sSql & "'" & prsPatVeh("cPlaca") & "',"
                sSql = sSql & "'" & prsPatVeh("cFecFabrica") & "',"
                sSql = sSql & prsPatVeh("nValorCom") & ","
                sSql = sSql & "'" & sUltimaActualizacion & "')"
                oConecta.ConexionActiva.Execute sSql
            Else
                If prsPatVeh("nTipoAct") = 3 Then
                    sSql = "DELETE PERSPATRIMONIO "
                    sSql = sSql & " WHERE cPersCod = '" & psPersCod & "' AND nNumPatri= " & prsPatVeh("nNumPatri")
                    oConecta.ConexionActiva.Execute sSql
                End If
            End If
        End If
        prsPatVeh.MoveNext
    Wend

    '*** PEAC 20080412
    'Graba Patrimonio Otros
    While Not prsPatOtros.EOF
        If prsPatOtros("nTipoAct") = 1 Then
            sSql = "UPDATE PERSPATRIOTROS SET "
            sSql = sSql & "cDescripcion= '" & prsPatOtros("cDescripcion") & "',"
            sSql = sSql & "nValorCom = " & prsPatOtros("nValorCom") & ","
            sSql = sSql & "cUltimaActualizacion = '" & sUltimaActualizacion & "' "
            sSql = sSql & " WHERE cPersCod = '" & psPersCod & "' AND nNumPatri= " & prsPatOtros("nNumPatri")
            oConecta.ConexionActiva.Execute sSql
        Else
            If prsPatOtros("nTipoAct") = 2 Then
                sSql = "INSERT INTO PERSPATRIOTROS(cPersCod, nNumPatri, cDescripcion, "
                sSql = sSql & "nValorCom, cUltimaActualizacion)"
                sSql = sSql & " VALUES ('" & psPersCod & "',"
                sSql = sSql & prsPatOtros("nNumPatri") & ","
                sSql = sSql & "'" & prsPatOtros("cDescripcion") & "',"
                sSql = sSql & prsPatOtros("nValorCom") & ","
                sSql = sSql & "'" & sUltimaActualizacion & "')"
                oConecta.ConexionActiva.Execute sSql
            Else
                If prsPatOtros("nTipoAct") = 3 Then
                    sSql = "DELETE PERSPATRIOTROS "
                    sSql = sSql & " WHERE cPersCod = '" & psPersCod & "' AND nNumPatri= " & prsPatOtros("nNumPatri")
                    oConecta.ConexionActiva.Execute sSql
                End If
            End If
        End If
        prsPatOtros.MoveNext
    Wend

    '*** PEAC 20080412
    'Graba Patrimonio Inmuebles
    While Not prsPatInmuebles.EOF
        If prsPatInmuebles("nTipoAct") = 1 Then
            sSql = "UPDATE PERSPATRIinmuebles SET "
            sSql = sSql & "cUbicacion= '" & prsPatInmuebles("cUbicacion") & "',"
            sSql = sSql & "nAreaTerreno = " & prsPatInmuebles("nAreaTerreno") & ","
            sSql = sSql & "nAreaConstru = " & prsPatInmuebles("nAreaConstru") & ","
            sSql = sSql & "cTipoUso= '" & prsPatInmuebles("cTipoUso") & "',"
            sSql = sSql & "cRRPP= '" & prsPatInmuebles("cRRPP") & "',"
            sSql = sSql & "nValorCom = " & prsPatInmuebles("nValorCom") & ","
            sSql = sSql & "cUltimaActualizacion = '" & sUltimaActualizacion & "' "
            sSql = sSql & " WHERE cPersCod = '" & psPersCod & "' AND nNumPatri= " & prsPatInmuebles("nNumPatri")
            oConecta.ConexionActiva.Execute sSql
        Else
            If prsPatInmuebles("nTipoAct") = 2 Then
                sSql = "INSERT INTO PERSPATRIInmuebles(cPersCod, nNumPatri,cUbicacion ,nAreaTerreno,nAreaConstru,cTipoUso,cRRPP, "
                sSql = sSql & "nValorCom, cUltimaActualizacion)"
                sSql = sSql & " VALUES ('" & psPersCod & "',"
                sSql = sSql & prsPatInmuebles("nNumPatri") & ","
                sSql = sSql & "'" & prsPatInmuebles("cUbicacion") & "',"
                sSql = sSql & prsPatInmuebles("nAreaTerreno") & ","
                sSql = sSql & prsPatInmuebles("nAreaConstru") & ","
                sSql = sSql & "'" & prsPatInmuebles("cTipoUso") & "',"
                sSql = sSql & "'" & prsPatInmuebles("cRRPP") & "',"
                sSql = sSql & prsPatInmuebles("nValorCom") & ","
                sSql = sSql & "'" & sUltimaActualizacion & "')"
                oConecta.ConexionActiva.Execute sSql
            Else
                If prsPatInmuebles("nTipoAct") = 3 Then
                    sSql = "DELETE PERSPATRIInmuebles "
                    sSql = sSql & " WHERE cPersCod = '" & psPersCod & "' AND nNumPatri= " & prsPatInmuebles("nNumPatri")
                    oConecta.ConexionActiva.Execute sSql
                End If
            End If
        End If
        prsPatInmuebles.MoveNext
    Wend

    'Actualizar Los Documentos de las Personas
    While Not prsDocs.EOF
        If prsDocs("nTipoAct") = 1 Then
            sSql = "UPDATE " & gConsPersona & "PERSID SET "
            sSql = sSql & "cPersIdTpo = '" & prsDocs("cPersIdTpo") & "',"
            sSql = sSql & "cPersIDNro = '" & prsDocs("cPersIDNro") & "',"
            sSql = sSql & "cUltimaActualizacion = '" & sUltimaActualizacion & "' "
            sSql = sSql & " WHERE cPersCod = '" & psPersCod & "' And cPersIDTpo = '" & prsDocs("nTipoOrig") & "'"
            oConecta.ConexionActiva.Execute sSql
        Else
            If prsDocs("nTipoAct") = 2 Then
                sSql = "INSERT INTO " & gConsPersona & "PERSID("
                sSql = sSql & "cPersCod,cPersIDTpo,cPersIDNro,cUltimaActualizacion)"
                sSql = sSql & "VALUES('" & psPersCod & "',"
                sSql = sSql & "'" & prsDocs("cPersIdTpo") & "',"
                sSql = sSql & "'" & prsDocs("cPersIDNro") & "',"
                sSql = sSql & "'" & sUltimaActualizacion & "')"
                oConecta.ConexionActiva.Execute sSql
            Else
                If prsDocs("nTipoAct") = 3 Then
                    sSql = "DELETE " & gConsPersona & "PERSID Where "
                    sSql = sSql & " cPersCod = '" & psPersCod & "'"
                    sSql = sSql & " And cPersIDTpo = '" & prsDocs("nTipoOrig") & "'"
                    oConecta.ConexionActiva.Execute sSql
                End If
            End If
        End If
        prsDocs.MoveNext
    Wend

    'Actualizar las Fuentes de Ingreso
    nFteDep = 0
    nFteInd = 0

    While Not prsFteIng.EOF
        If prsFteIng("nTipoAct") = 2 Then
            sNumFuente = Format(CLng(sNumFuente) + 1, "00000000")
            If prsFteIng("nPersTipFte") = 1 Then
                nFteDep = nFteDep + 1
                ReDim Preserve sFtesDep(nFteDep)
                sFtesDep(nFteDep) = sNumFuente
            Else
                nFteInd = nFteInd + 1
                ReDim Preserve sFtesInd(nFteInd)
                sFtesInd(nFteInd) = sNumFuente
            End If
            sSql = "INSERT INTO " & gConsPersona & "PersFteIngreso(cNumFuente, cPersCod, cPersFIPersCod, cPersFICargo,"
            sSql = sSql & "dPersFIinicio, cPersFIMoneda, cPersFIComentario, cRazSocUbiGeo, cRazSocDirecc, "
            sSql = sSql & " cRazSocTelef, cRazSocDescrip, nPersTipFte, cUltimaActualizacion,bCostoProd) "
            sSql = sSql & " VALUES('" & sNumFuente & "','" & psPersCod & "',"
            sSql = sSql & "'" & prsFteIng("cPersFIPersCod") & "',"
            sSql = sSql & "'" & prsFteIng("cPersFICargo") & "',"
            sSql = sSql & "'" & Format(prsFteIng("dPersFIinicio"), "mm/dd/yyyy") & "',"
            sSql = sSql & prsFteIng("cPersFIMoneda") & ","
            sSql = sSql & "'" & prsFteIng("cPersFIComentario") & "',"
            sSql = sSql & "'" & prsFteIng("cRazSocUbiGeo") & "',"
            sSql = sSql & "'" & prsFteIng("cRazSocDirecc") & "',"
            sSql = sSql & "'" & prsFteIng("cRazSocTelef") & "',"
            sSql = sSql & "'" & prsFteIng("cRazSocDescrip") & "',"
            sSql = sSql & prsFteIng("nPersTipFte") & ",'" & sUltimaActualizacion & "'," & prsFteIng("bCostoProd")
            sSql = sSql & ")"
            oConecta.ConexionActiva.Execute sSql
            sNuevo = "S"
        Else
            If prsFteIng("nTipoAct") = 1 Then
                           
                '**** PEAC 200800710 - VERIFICA QUE REALMENTE EXISTA LA FUENTE DE INGRESO ************************
                '**** SI NO EXISTE LO CREA
                
                sSql = "select * from PersFteIngreso where cnumfuente='" & prsFteIng("cNumFuente") & "'"
                Set R = New ADODB.Recordset
                Set R = oConecta.CargaRecordSet(sSql)
            
                If R.RecordCount > 0 Then
                
                    sSql = "UPDATE " & gConsPersona & "PersFteIngreso SET "
                    sSql = sSql & " cPersFIPersCod = '" & prsFteIng("cPersFIPersCod") & "',"
                    sSql = sSql & " cPersFICargo = '" & prsFteIng("cPersFICargo") & "',"
                    sSql = sSql & " dPersFIInicio = '" & Format(prsFteIng("dPersFIinicio"), "mm/dd/yyyy") & "',"
                    sSql = sSql & " cPersFIMoneda = " & prsFteIng("cPersFIMoneda") & ","
                    sSql = sSql & " cPersFIComentario = '" & prsFteIng("cPersFIComentario") & "',"
                    sSql = sSql & " cUltimaActualizacion = '" & sUltimaActualizacion & "',"
                    sSql = sSql & " cRazSocUbiGeo = '" & prsFteIng("cRazSocUbiGeo") & "',"
                    sSql = sSql & " cRazSocDirecc = '" & prsFteIng("cRazSocDirecc") & "',"
                    sSql = sSql & " cRazSocTelef = '" & prsFteIng("cRazSocTelef") & "',"
                    sSql = sSql & " cRazSocDescrip = '" & prsFteIng("cRazSocDescrip") & "', "
                    sSql = sSql & " bCostoProd = " & prsFteIng("bCostoProd")
                    sSql = sSql & " WHERE cNumFuente = '" & prsFteIng("cNumFuente") & "'"
                    oConecta.ConexionActiva.Execute sSql
                
                Else
                
                    sSql = "INSERT INTO " & gConsPersona & "PersFteIngreso(cNumFuente, cPersCod, cPersFIPersCod, cPersFICargo,"
                    sSql = sSql & "dPersFIinicio, cPersFIMoneda, cPersFIComentario, cRazSocUbiGeo, cRazSocDirecc, "
                    sSql = sSql & " cRazSocTelef, cRazSocDescrip, nPersTipFte, cUltimaActualizacion,bCostoProd) "
                    sSql = sSql & " VALUES('" & sNumFuente & "','" & psPersCod & "',"
                    sSql = sSql & "'" & prsFteIng("cPersFIPersCod") & "',"
                    sSql = sSql & "'" & prsFteIng("cPersFICargo") & "',"
                    sSql = sSql & "'" & Format(prsFteIng("dPersFIinicio"), "mm/dd/yyyy") & "',"
                    sSql = sSql & prsFteIng("cPersFIMoneda") & ","
                    sSql = sSql & "'" & prsFteIng("cPersFIComentario") & "',"
                    sSql = sSql & "'" & prsFteIng("cRazSocUbiGeo") & "',"
                    sSql = sSql & "'" & prsFteIng("cRazSocDirecc") & "',"
                    sSql = sSql & "'" & prsFteIng("cRazSocTelef") & "',"
                    sSql = sSql & "'" & prsFteIng("cRazSocDescrip") & "',"
                    sSql = sSql & prsFteIng("nPersTipFte") & ",'" & sUltimaActualizacion & "'," & prsFteIng("bCostoProd")
                    sSql = sSql & ")"
                    oConecta.ConexionActiva.Execute sSql
        
                End If
    
            R.Close
            
            '************************************************************************************
                        
            Else
                If prsFteIng("nTipoAct") = 3 Then
                    sSql = "DELETE PersFIIndependiente Where cNumFuente = '" & prsFteIng("cNumFuente") & "' "
                    oConecta.ConexionActiva.Execute sSql

                    sSql = "DELETE PersFIDependiente Where cNumFuente = '" & prsFteIng("cNumFuente") & "' "
                    oConecta.ConexionActiva.Execute sSql

                    sSql = "DELETE PersFICostoProd Where cNumFuente = '" & prsFteIng("cNumFuente") & "' "
                    oConecta.ConexionActiva.Execute sSql

                    sSql = "DELETE " & gConsPersona & "PersFteIngreso "
                    sSql = sSql & " WHERE cNumFuente = '" & prsFteIng("cNumFuente") & "'"
                    oConecta.ConexionActiva.Execute sSql
                End If
            End If
        End If
        prsFteIng.MoveNext
    Wend

   If prsFIDep.RecordCount > 0 Then

    nFteDep = 0
    'FUENTES DE INGRESO DEPENDIENTES
    While Not prsFIDep.EOF

            If prsFIDep("nTipoAct") = 2 Then
                nFteDep = nFteDep + 1
                'ARCV 12-08-2006
                If Trim(prsFIDep("cNumFuente")) = "" Then
                    sNumFuente = sFtesDep(nFteDep)
                Else
                    sNumFuente = prsFIDep("cNumFuente") 'sFtesDep(nFteDep)
                End If
                '-------------------
                sSql = "INSERT INTO PersFIDependiente(cNumFuente, dPersEval, nPersIngCli, "
                sSql = sSql & " nPersGastoFam, nPersIngCon, nPersOtrIng, dPersFICaduca) "
                sSql = sSql & " VALUES('" & sNumFuente & "',"
                sSql = sSql & "'" & Format(prsFIDep("dPersEval"), "mm/dd/yyyy") & "',"
                sSql = sSql & prsFIDep("nPersIngCli") & ","
                sSql = sSql & prsFIDep("nPersGastoFam") & ","
                sSql = sSql & prsFIDep("nPersIngCon") & ","
                sSql = sSql & prsFIDep("nPersOtrIng") & ","
                sSql = sSql & "'" & Format(prsFIDep("dPersFICaduca"), "mm/dd/yyyy") & "')"
                oConecta.ConexionActiva.Execute sSql
            Else
                If prsFIDep("nTipoAct") = 1 Then
                    sSql = "UPDATE PersFIDependiente SET "
                    sSql = sSql & " nPersIngCli = " & prsFIDep("nPersIngCli") & ","
                    sSql = sSql & " nPersGastoFam = " & prsFIDep("nPersGastoFam") & ","
                    sSql = sSql & " nPersIngCon = " & prsFIDep("nPersIngCon") & ","
                    sSql = sSql & " nPersOtrIng = " & prsFIDep("nPersOtrIng") & ","
                    sSql = sSql & " dPersFICaduca = '" & Format(prsFIDep("dPersFICaduca"), "mm/dd/yyyy") & "' "
                    sSql = sSql & " Where cNumFuente = '" & prsFIDep("cNumFuente") & "' AND dPersEval = '" & Format(prsFIDep("dPersEval"), "mm/dd/yyyy") & "'"
                    oConecta.ConexionActiva.Execute sSql
                Else
                    If prsFIDep("nTipoAct") = 3 Then
                        sSql = "DELETE PersFIDependiente Where cNumFuente = '" & prsFIDep("cNumFuente") & "' AND dPersEval = '" & Format(prsFIDep("dPersEval"), "mm/dd/yyyy") & "'"
                        oConecta.ConexionActiva.Execute sSql
                    End If
                End If
            End If
        prsFIDep.MoveNext
    Wend
        If nIntValor = 1 Then
            Call GrabarFteHojaeval(prsHojEval, nFteDep, sNuevo, sNumFuente, oConecta)
            Call GrabarFteEstFinan(prsEstFinan, sNuevo, sNumFuente, oConecta) 'WIOR 20140321
        End If
    End If

    'FUENTES DE INGRESO INDEPENDIENTES
    nFteInd = 0
    If prsFIInd.RecordCount > 0 Then
    While Not prsFIInd.EOF
            If prsFIInd("nTipoAct") = 2 Then
                nFteInd = nFteInd + 1
                'ARCV 12-08-2006

                If Trim(prsFIInd("cNumFuente")) = "" Then
                    sNumFuente = sFtesInd(nFteInd)
                Else
                    sNumFuente = prsFIInd("cNumFuente") 'sFtesInd(nFteInd)
                End If
                '-----------------
                'peac 20071227 se agrego
                sSql = "INSERT INTO PersFIIndependiente(cNumFuente, dPersEval, "
                sSql = sSql & " dPersFICaduca, nPersFIActivoDisp, nPersFICtasxCobrar, nPersFIInventarios, "
                sSql = sSql & " nPersFIActivosFijos, nPersFIProveedores, nPersFICredCMACT, nPersFICredOtros, "
                sSql = sSql & " nPersFIVentas, nPersFIPatrimonio, nPersFICostoVentas, nPersFIRecupCtasXCobrar, "
                sSql = sSql & " nPersFIEgresosOtros, nPersIngFam, nPersEgrFam, cUltimaActualizacion,dFecEEFF)"
                sSql = sSql & " VALUES('" & sNumFuente & "',"
                sSql = sSql & "'" & Format(prsFIInd("dPersEval"), "mm/dd/yyyy") & "',"
                sSql = sSql & "'" & Format(prsFIInd("dPersFICaduca"), "mm/dd/yyyy") & "',"
                sSql = sSql & prsFIInd("nPersFIActivoDisp") & ","
                sSql = sSql & prsFIInd("nPersFICtasxCobrar") & ","
                sSql = sSql & prsFIInd("nPersFIInventarios") & ","
                sSql = sSql & prsFIInd("nPersFIActivosFijos") & ","
                sSql = sSql & prsFIInd("nPersFIProveedores") & ","
                sSql = sSql & prsFIInd("nPersFICredCMACT") & ","
                sSql = sSql & prsFIInd("nPersFICredOtros") & ","
                sSql = sSql & prsFIInd("nPersFIVentas") & ","
                sSql = sSql & prsFIInd("nPersFIPatrimonio") & ","
                sSql = sSql & prsFIInd("nPersFICostoVentas") & ","
                sSql = sSql & prsFIInd("nPersFIRecupCtasXCobrar") & ","
                sSql = sSql & prsFIInd("nPersFIEgresosOtros") & ","
                sSql = sSql & prsFIInd("nPersIngFam") & ","
                sSql = sSql & prsFIInd("nPersEgrFam") & ","
                sSql = sSql & "'" & sUltimaActualizacion & "',"
                sSql = sSql & "'" & Format(prsFIInd("dFecEEFF"), "mm/dd/yyyy") & "')"
                oConecta.ConexionActiva.Execute sSql
                dfecha = prsFIInd("dPersEval")
            Else
                If prsFIInd("nTipoAct") = 1 Then
                    sSql = " UPDATE PersFIIndependiente SET "
                    sSql = sSql & " nPersFIActivoDisp = " & prsFIInd("nPersFIActivoDisp") & ","
                    sSql = sSql & " nPersFICtasxCobrar = " & prsFIInd("nPersFICtasxCobrar") & ","
                    sSql = sSql & " nPersFIInventarios = " & prsFIInd("nPersFIInventarios") & ","
                    sSql = sSql & " nPersFIActivosFijos = " & prsFIInd("nPersFIActivosFijos") & ","
                    sSql = sSql & " nPersFIProveedores = " & prsFIInd("nPersFIProveedores") & ","
                    sSql = sSql & " nPersFICredCMACT = " & prsFIInd("nPersFICredCMACT") & ","
                    sSql = sSql & " nPersFICredOtros = " & prsFIInd("nPersFICredOtros") & ","
                    sSql = sSql & " nPersFIVentas = " & prsFIInd("nPersFIVentas") & ","
                    sSql = sSql & " nPersFIPatrimonio = " & prsFIInd("nPersFIPatrimonio") & ","
                    sSql = sSql & " nPersFICostoVentas = " & prsFIInd("nPersFICostoVentas") & ","
                    sSql = sSql & " nPersFIRecupCtasXCobrar = " & prsFIInd("nPersFIRecupCtasXCobrar") & ","
                    sSql = sSql & " nPersFIEgresosOtros = " & prsFIInd("nPersFIEgresosOtros") & ","
                    sSql = sSql & " nPersIngFam = " & prsFIInd("nPersIngFam") & ","
                    sSql = sSql & " nPersEgrFam = " & prsFIInd("nPersEgrFam") & ","
                    sSql = sSql & " dFecEEFF = '" & Format(prsFIInd("dFecEEFF"), "mm/dd/yyyy") & "', "
                    sSql = sSql & " dPersFICaduca = '" & Format(prsFIInd("dPersFICaduca"), "mm/dd/yyyy") & "' "
                    sSql = sSql & " Where cNumFuente = '" & prsFIInd("cNumFuente") & "' AND dPersEval = '" & Format(prsFIInd("dPersEval"), "mm/dd/yyyy") & "'"

                    oConecta.ConexionActiva.Execute sSql
                Else
                    If prsFIInd("nTipoAct") = 3 Then
                        sSql = "DELETE PersFIIndependiente Where cNumFuente = '" & prsFIInd("cNumFuente") & "' AND dPersEval = '" & Format(prsFIInd("dPersEval"), "mm/dd/yyyy") & "'"
                        oConecta.ConexionActiva.Execute sSql
                    End If
                End If
            End If
        prsFIInd.MoveNext
    Wend
    
    '*PEAC****20080412************************************************************************
    If nIntValor = 2 Then
        Call GrabarFteHojaeval(prsHojEval, nFteInd, sNuevo, sNumFuente, oConecta)
        Call GrabarFteEstFinan(prsEstFinan, sNuevo, sNumFuente, oConecta) 'WIOR 20140321
    End If
End If



    '*End***************************************************************************************
    'Actualiza Costos de Produccion
    nFteInd = 0
    While Not prsFICosto.EOF
            If prsFICosto("nTipoAct") = 2 Then
                nFteInd = nFteInd + 1
                'ARCV 12-08-2006
                If Trim(prsFICosto("cNumFuente")) = "" Then
                    sNumFuente = sFtesInd(nFteInd)
                Else
                    sNumFuente = prsFICosto("cNumFuente") 'sFtesInd(nFteInd)
                End If
                '------------------
                sSql = "INSERT INTO PersFICostoProd(cNumFuente, dFecEval, nTpoCultivo, nMaquinaria, nJornales, nInsumos, nPesticidas, nOtros, nHectareas, "
                sSql = sSql & " nProduccion, nUniProd, nPreUni,nSiembra,nCosecha,nDesaAgricola,nPlanOtros,nMantenimiento)"
                sSql = sSql & " VALUES('" & sNumFuente & "','" & Format(prsFICosto("dFecEval"), "mm/dd/yyyy") & "'," & prsFICosto("nTpoCultivo") & ","
                sSql = sSql & prsFICosto("nMaquinaria") & "," & prsFICosto("nJornales") & ","
                sSql = sSql & prsFICosto("nInsumos") & "," & prsFICosto("nPesticidas") & ","
                sSql = sSql & prsFICosto("nOtros") & "," & prsFICosto("nHectareas") & ","
                sSql = sSql & prsFICosto("nProduccion") & "," & prsFICosto("nUniProd") & "," & prsFICosto("nPreUni") & ","
                sSql = sSql & prsFICosto("nSiembra") & ","
                sSql = sSql & prsFICosto("nCosecha") & "," & prsFICosto("nDesaAgricola") & "," & prsFICosto("nDesaAgricola") & ","
                sSql = sSql & prsFICosto("nMantenimiento") & ")"

                oConecta.ConexionActiva.Execute sSql

            Else
                If prsFICosto("nTipoAct") = 1 Then
                    sSql = "UPDATE PersFICostoProd SET "
                    sSql = sSql & "nTpoCultivo = " & prsFICosto("nTpoCultivo") & ","
                    sSql = sSql & "nMaquinaria = " & prsFICosto("nMaquinaria") & ","
                    sSql = sSql & "nJornales = " & prsFICosto("nJornales") & ","
                    sSql = sSql & "nInsumos = " & prsFICosto("nInsumos") & ","
                    sSql = sSql & "nPesticidas = " & prsFICosto("nPesticidas") & ","
                    sSql = sSql & "nOtros = " & prsFICosto("nOtros") & ","
                    sSql = sSql & "nHectareas = " & prsFICosto("nHectareas") & ","
                    sSql = sSql & "nProduccion = " & prsFICosto("nProduccion") & ","
                    sSql = sSql & "nUniProd = " & prsFICosto("nUniProd") & ","
                    sSql = sSql & "nPreUni = " & prsFICosto("nPreUni") & ","
                    sSql = sSql & "nSiembra=" & prsFICosto("nSiembra") & ","
                    sSql = sSql & "nCosecha=" & prsFICosto("nCosecha") & ","
                    sSql = sSql & "nDesaAgricola=" & prsFICosto("nDesaAgricola") & ","
                    sSql = sSql & "nPlanOtros=" & prsFICosto("nDesaAgricola") & ","
                    sSql = sSql & "nMantenimiento= " & prsFICosto("nMantenimiento") & " "
                    sSql = sSql & " WHERE cNumFuente = '" & prsFICosto("cNumFuente") & "' AND dFecEval = '" & Format(prsFICosto("dFecEval"), "mm/dd/yyyy") & "'"

                    oConecta.ConexionActiva.Execute sSql
                Else
                    If prsFICosto("nTipoAct") = 3 Then
                        sSql = "DELETE PersFICostoProd Where cNumFuente = '" & prsFICosto("cNumFuente") & "'"
                        oConecta.ConexionActiva.Execute sSql
                    End If
                End If
            End If
        prsFICosto.MoveNext
    Wend

    oConecta.CommitTrans


    '*******************************************************************************
    'Replicar Persona a Distribuido
    '*******************************************************************************
    If bPersDistrib Then
        oConexDitrib.Ejecutar "SET DATEFORMAT MDY"
        oConexDitrib.BeginTrans
        TranDistribBegin = True
        If pnTipoAct <> 2 Then
            '*********************************************************************************
            'Para Actualizar Persona Distribuida
            '*********************************************************************************
            If bPersDistrib Then

            End If

        Else

                '*********************************************************************************
                'Para INSERTAR Persona Distribuida
                '*********************************************************************************
                If bPersDistrib Then
                    sSql = "INSERT INTO Persona(cCodPers,cNomPers,cTidoci,cNudoci,cTidotr,cNudotr,cDirPers, cCodZon, "
                    sSql = sSql & " dFecNac, cTelPers, cEstPers, cTipPers, cCodSBS, dIngPers, dUltAct, cCodUsu, cRelIns, cConDom, cFlagConsol, cEMail, cWeb, cTelefono2) "
                    sSql = sSql & " VALUES('" & sCodPersDistrib & "','" & prsPersona("cPersNombre") & "','" & psDoc0 & "','" & psDoc1 & "','" & psDoc2 & "','" & psDoc3 & "','"
                    sSql = sSql & prsPersona("cPersDireccDomicilio") & "','" & prsPersona("cPersDireccUbiGeo") & "','" & Format(prsPersona("dPersNacCreac"), "mm/dd/yyyy") & "','" & prsPersona("cPersTelefono") & "','1','"
                    sSql = sSql & CInt(prsPersona("nPersPersoneria")) & "','" & prsPersona("cPersCodSbs") & "','" & Mid(sUltimaActualizacion, 1, 4) & "/" & Mid(sUltimaActualizacion, 5, 2) & "/" & Mid(sUltimaActualizacion, 7, 2) & "','" & Mid(sUltimaActualizacion, 1, 4) & "/" & Mid(sUltimaActualizacion, 5, 2) & "/" & Mid(sUltimaActualizacion, 7, 2)
                    sSql = sSql & "','','" & sPersRelIDistrib & "','" & prsPersona("cPersDireccCondicion") & "',NULL,'','','')"

                    oConexDitrib.ConexionActiva.Execute sSql
                End If

                '******************************************************************************
                'ACTUALIZA PERSONA NATURAL EN DISTRIBUIDO
                '******************************************************************************
                If bPersDistrib Then
                    sSql = "INSERT INTO Personanat(cCodPers, cSexPers, cEstCiv, cCodProf, cCodCon, nNumHij)"
                    sSql = sSql & " VALUES('" & sCodPersDistrib & "','" & prsPersona("cPersnatSexo") & "','" & prsPersona("nPersNatEstCiv") & "','',''," & prsPersona("nPersNatHijos") & ")"
                    oConexDitrib.ConexionActiva.Execute sSql
                End If

                '******************************************************************************
                'ACTUALIZA PERSONA JRUDICA EN DISTRIBUIDO
                '******************************************************************************
                If bPersDistrib Then
                    sSql = "Insert into PersonaJur(cCodPers,cTipPerJur,cCodRegPub,cMagEmp,cSigla) "
                    sSql = sSql & "Values('" & sCodPersDistrib & "','0000','','','" & prsPersona("cPersJurSigla") & "')"
                    oConexDitrib.ConexionActiva.Execute sSql
                End If

        End If
        oConexDitrib.CommitTrans
        'oConexDitrib.Close
        Set oConexDitrib = Nothing
        TranDistribBegin = False
    End If

    Exit Function

ErrorGrabarDatosPersona:
    Set R = Nothing
    Set oConecta = Nothing
    'Err.Raise Err.Number, "Grabar Datos Persona", Err.Description
End Function
Public Sub GrabarFteHojaeval(ByVal prsHojEval As ADODB.Recordset, ByVal nFteInd As Integer, ByVal sNuevo As String, ByVal sNumFuente As String, ByVal oConecta As COMConecta.DCOMConecta)
Dim sSql As String
sSql = ""
 If prsHojEval.RecordCount > 0 Then
        Dim rSearch As ADODB.Recordset
        nFteInd = 0
        Dim nPersHojC As Integer
        nPersHojC = 0
        While Not prsHojEval.EOF
            nPersHojC = nPersHojC + 1
            If nPersHojC = 1 Then
                sSql = "update PersFIHojaEvaluacion set nEstado=0 "
                sSql = sSql & "where cNumFuente='" & prsHojEval("cNumFuente") & "' "
                sSql = sSql & "and datediff(day,dPersEval,'" & Format(prsHojEval("dPersEval"), "yyyy/mm/dd") & "')=0 "
                oConecta.ConexionActiva.Execute sSql
            End If
            If sNuevo = "S" Then
                prsHojEval("cNumFuente") = sNumFuente
            End If
            sSql = "select cNumFuente from PersFIHojaEvaluacion "
            sSql = sSql & "where cNumFuente='" & prsHojEval("cNumFuente") & "'"
            sSql = sSql & "and cCodHojEval ='" & Trim(prsHojEval("cCodHojEval")) & "'"
            sSql = sSql & "and datediff(day,dPersEval,'" & Format(prsHojEval("dPersEval"), "yyyy/mm/dd") & "')=0 "
            sSql = sSql & "and nTipoImporte='" & prsHojEval("nTipoAct") & "'"
            Set rSearch = oConecta.CargaRecordSet(sSql)
            If rSearch.RecordCount > 0 Then
                sSql = "update PersFIHojaEvaluacion set nEstado=1,nImporte='" & prsHojEval("nValorEval") & "' "
                sSql = sSql & "where cNumFuente='" & prsHojEval("cNumFuente") & "' "
                sSql = sSql & "and cCodHojEval ='" & Trim(prsHojEval("cCodHojEval")) & "' "
                sSql = sSql & "and datediff(day,dPersEval,'" & Format(prsHojEval("dPersEval"), "yyyy/mm/dd") & "')=0 "
                sSql = sSql & "and nTipoImporte='" & prsHojEval("nTipoAct") & "'"
                oConecta.ConexionActiva.Execute sSql
            Else
                sSql = "insert into PersFIHojaEvaluacion (cNumFuente,cCodHojEval,dPersEval,nTipoImporte,nImporte,nEstado)"
                sSql = sSql & "values ('" & prsHojEval("cNumFuente") & "'," & Trim(prsHojEval("cCodHojEval")) & ",'" & Format(prsHojEval("dPersEval"), "yyyy/mm/dd") & "'," & prsHojEval("nTipoAct") & "," & prsHojEval("nValorEval") & ",1)"
                oConecta.ConexionActiva.Execute sSql
            End If
             prsHojEval.MoveNext
        Wend
    End If
End Sub

'*********************************************************************************
' Funcion que recupera todos los Datos de la Persona para la Solicitud de Credito
'*********************************************************************************

Public Function RecuperarDatosPersona_Solicitud(ByVal psPersCod As String, _
                                    ByVal psCodAge As String, _
                                    ByRef prsPersona As ADODB.Recordset, _
                                    ByRef prsFtesIng As ADODB.Recordset, _
                                    ByRef prsDocs As ADODB.Recordset, Optional gdFecSis As Date = "1900/01/01")

Dim sSql As String
Dim oConecta  As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperarDatosPersona
    
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
            
        '***** CARGA DATOS DE LA PERSONA ******
'        sSql = "Select P.nPersRelaInst, P.cPersCodSbs, P.cPersEmail cEmail,P.cPersTelefono2,P.cPersCod,P.cPersNombre,P.dPersNacCreac, "
'        sSql = sSql & "P.cPersDireccUbiGeo,P.cPersDireccDomicilio,P.cPersDireccCondicion, P.nPersValComDomicilio, P.cPersTelefono, "
'        sSql = sSql & "P.nPersPersoneria,P.cPersCIIU,P.cPersEstado,P.cUltimaActualizacion, "
'        sSql = sSql & "PN.cPersNatSexo,PN.nPersNatEstCiv,PN.nPersNatHijos,PN.nPersPeso,PN.nPersTalla,PN.nPersTpoSangre, PN.cNacionalidad, PN.nResidente, "
'        sSql = sSql & "PJ.cPersJurTpo, PJ.cPersJurMagnitud,PJ.cPersJurSigla,PJ.nPersJurEmpleados, CI.cCIIUdescripcion"
'        sSql = sSql & " FROM " & gConsPersona & "Persona P left join " & gConsPersona & "PersonaNat PN ON P.cPersCod=PN.cPersCod "
'        sSql = sSql & " LEFT JOIN " & gConsPersona & "PersonaJur PJ ON P.cPersCod = PJ.cPersCod "
'        sSql = sSql & " LEFT JOIN CIIU CI ON P.cPersCIIU = CI.cCIIUcod "
'        sSql = sSql & " WHERE P.cPersCod = '" & psPersCod & "'"
        
        '*** PEAC 20080402
        sSql = "exec stp_sel_ObtieneDatosPersonaParaSolicitud '" & psPersCod & "'"
    
        Set prsPersona = oConecta.CargaRecordSet(sSql)
        If prsPersona.EOF Then
            oConecta.CierraConexion
            Set prsPersona = Nothing
            Set oConecta = Nothing
            Exit Function
        End If
    
        '***************** Documentos ****************
        sSql = "Select PID.cPersIDnro,PID.cPersIDTpo,C.cConsDescripcion,C.nConsValor From " & gConsPersona & "PersID PID inner join Constante C ON convert(int,PID.cPersIDTpo) = C.nConsValor Where PID.cPersCod = '" & psPersCod & "' And  C.nConsCod = " & Trim(Str(gPersIdTipo)) & " And C.cConsDescripcion <> '" & Trim(Str(gPersIdTipo)) & "'"
        Set prsDocs = oConecta.CargaRecordSet(sSql)
    
        '******************** Recupera Fuentes de Ingreso *******
        
        sSql = "Select PF.bCostoProd, PF.cNumFuente, PF.cPersCod, PF.cPersFIPersCod, PF.cPersFICargo, "
        sSql = sSql & " PF.dPersFIinicio, PF.cPersFIMoneda, PF.cPersFIComentario, "
        sSql = sSql & " PF.cRazSocUbiGeo, PF.cRazSocDirecc, PF.cRazSocTelef, PF.cRazSocDescrip, "
        sSql = sSql & " PF.nPersTipFte, PF.cUltimaActualizacion, "
        '------------------------------------------------------------------------------------------------------------
        sSql = sSql & "P.cPersDireccUbiGeo,P.cPersDireccDomicilio,P.cPersDireccCondicion, P.cPersTelefono, "
        sSql = sSql & "P.nPersPersoneria,P.cPersCIIU, P.cPersTelefono2, P.cPersCod as cPersFteCod,P.cPersNombre,P.dPersNacCreac, "
        sSql = sSql & "PJ.cPersJurTpo, PJ.cPersJurMagnitud,PJ.cPersJurSigla,PJ.nPersJurEmpleados, CI.cCIIUdescripcion, "
        sSql = sSql & "cNroRUC = (  Select ISNULL(cPersIDnro,'') From PersID WHERE PersID.cPersCod = P.cPersCod and PersID.cPersIDTpo = " & gPersIdRUC & " ), "
        sSql = sSql & "PJT.cPersJurTpoDesc, CN7.cConsDescripcion as cPersDireccCondicionDes, CN8.cConsDescripcion as cPersJurMagnitudDes, "
        sSql = sSql & " Dpto = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '1' And Substring(u.cUbigeoCod,2,2) = Substring(P.cPersDireccUbiGeo,2,2) And Substring(u.cUbigeoCod,4,9) = '000000000' ), '' ) , "
        sSql = sSql & " Prov = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '2' And Substring(u.cUbigeoCod,2,4) = Substring(P.cPersDireccUbiGeo,2,4) And Substring(u.cUbigeoCod,6,7) = '0000000' ), '' ) , "
        sSql = sSql & " Dist = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '3' And Substring(u.cUbigeoCod,2,6) = Substring(P.cPersDireccUbiGeo,2,6) And Substring(u.cUbigeoCod,8,5) = '00000' ), '' ), "
        sSql = sSql & " Zona = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where u.cUbigeoCod = P.cPersDireccUbiGeo ),'') "
        sSql = sSql & " ,isnull(max(x.descr),'1900/01/01') as descr,isnull(max(x.descr),'1900/01/01') as valor " ' PEAC 19/05/2008
        sSql = sSql & " From  PersFteIngreso PF "
        sSql = sSql & " INNER JOIN " & gConsPersona & "Persona P ON PF.cPersFIPersCod = P.cPersCod "
        sSql = sSql & " left join " & gConsPersona & "PersonaNat PN ON P.cPersCod=PN.cPersCod "
        sSql = sSql & " LEFT JOIN " & gConsPersona & "PersonaJur PJ ON P.cPersCod = PJ.cPersCod "
        sSql = sSql & " LEFT JOIN CIIU CI ON P.cPersCIIU = CI.cCIIUcod "
        sSql = sSql & " LEFT JOIN PersJurTpo PJT ON PJ.cPersJurTpo = PJT.cPersJurTpoCod "
        sSql = sSql & " Left Join Constante CN7 ON CN7.nConsValor = P.cPersDireccCondicion AND CN7.nConsCod = 1029 "
        '**PEAC**19/05/2008
        sSql = sSql & " Left Join Constante CN8 ON CN8.nConsValor = PJ.cPersJurMagnitud AND CN8.nConsCod = 1004 "
         sSql = sSql & " Left Join"
         sSql = sSql & " ("
         sSql = sSql & " (select FI.cNumFuente,dPersEval as descr,dPersEval as valor"
         sSql = sSql & " from persFteIngreso FI inner join persFIIndependiente FII"
         sSql = sSql & " on (FI.cNumFuente=FII.cNumFuente)"
         sSql = sSql & " where " ' --FI.cNumFuente=PF.cNumFuente and
         sSql = sSql & "     FII.dPersFICaduca>='" & Format(gdFecSis, "YYYYMMDD") & "'"
         sSql = sSql & " )"
         sSql = sSql & " Union"
         sSql = sSql & " ("
         sSql = sSql & " select FI.cNumFuente,dPersEval as descr,dPersEval as valor"
         sSql = sSql & " from persFteIngreso FI inner join persFIdependiente FID"
         sSql = sSql & " on (FI.cNumFuente=FID.cNumFuente)"
         sSql = sSql & " where " '--FI.cNumFuente=PF.cNumFuente and
         sSql = sSql & "  FID.dPersFICaduca>= '" & Format(gdFecSis, "YYYYMMDD") & "'"
         sSql = sSql & " )"
         sSql = sSql & " )x  on x.cNumFuente=PF.cNumFuente"
         '**End PEAC*******************
        sSql = sSql & " Where PF.cPersCod = '" & psPersCod & "'"
        sSql = sSql & " and (exists(Select cnumfuente from persfidependiente where cnumfuente=PF.cNumfuente) or Exists(Select cnumfuente from persfiindependiente where cnumfuente=PF.cnumfuente))"
        '**PEAC**19/05/2008
        sSql = sSql & " group by PF.bCostoProd,PF.cNumFuente,PF.cPersCod,PF.cPersFIPersCod,PF.cPersFICargo,"
        sSql = sSql & " PF.dPersFIinicio, PF.cPersFIMoneda, PF.cPersFIComentario,  PF.cRazSocUbiGeo, PF.cRazSocDirecc, PF.cRazSocTelef, PF.cRazSocDescrip,  PF.nPersTipFte, PF.cUltimaActualizacion, P.cPersDireccUbiGeo,P.cPersDireccDomicilio,P.cPersDireccCondicion, P.cPersTelefono, P.nPersPersoneria,P.cPersCIIU, P.cPersTelefono2, P.cPersCod,"
        sSql = sSql & " P.cPersNombre,P.dPersNacCreac, PJ.cPersJurTpo, PJ.cPersJurMagnitud,PJ.cPersJurSigla,PJ.nPersJurEmpleados, CI.cCIIUdescripcion,"
        sSql = sSql & " PJT.cPersJurTpoDesc , CN7.cConsDescripcion, CN8.cConsDescripcion"
        '**End PEAC*******************
        sSql = sSql & " Order By dPersFIinicio desc"
        '----------------------------------------------------------------------------------------------------------
    
        Set prsFtesIng = oConecta.CargaRecordSet(sSql)
        
        oConecta.CierraConexion
        Set oConecta = Nothing
        
        Exit Function
    
ErrorRecuperarDatosPersona:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Recuperar Datos Persona", Err.Description
End Function



'************************************
'Para evitar llamar a otro Componente
'************************************

Public Function DigitoChequeo(ByVal psCorrelativo As String) As Integer
Dim liFactor As Variant
Dim liCadena(1 To 7) As Integer
Dim liSum, i As Integer
Dim lnDigito As Integer
    liFactor = Array(2, 7, 6, 5, 4, 3, 2)
    liCadena(1) = Val(Mid(psCorrelativo, 1, 1))
    liCadena(2) = Val(Mid(psCorrelativo, 2, 1))
    liCadena(3) = Val(Mid(psCorrelativo, 3, 1))
    liCadena(4) = Val(Mid(psCorrelativo, 4, 1))
    liCadena(5) = Val(Mid(psCorrelativo, 5, 1))
    liCadena(6) = Val(Mid(psCorrelativo, 6, 1))
    liCadena(7) = Val(Mid(psCorrelativo, 7, 1))
    liSum = 0
    For i = 1 To 7
        liSum = liSum + liCadena(i) * liFactor(i - 1)
    Next i
    
    lnDigito = 11 - (liSum Mod 11)
    If lnDigito = 10 Then
        DigitoChequeo = 0
    ElseIf lnDigito = 11 Then
        DigitoChequeo = 1
    Else
        DigitoChequeo = lnDigito
    End If
End Function

Public Function FillNum(intNumero As String, intLenNum As Integer, ChrFil As String) As String
  FillNum = Left(String(intLenNum, ChrFil), (Len(String(intLenNum, ChrFil)) - Len(Trim(intNumero)))) + Trim(intNumero)
End Function

Public Function GeneraMovNroActualiza(pdFecha As Date, psCoduser As String, psCodCmac As String, psCodAge As String) As String
Dim oConecta As New COMConecta.DCOMConecta
Dim oVar As New COMFunciones.FCOMVarPublicas

oConecta.AbreConexion
GeneraMovNroActualiza = Format(pdFecha & " " & oConecta.GetHoraServer, oVar.gsFormatoMovFechaHora) & psCodCmac & psCodAge & "00" & psCoduser

oConecta.CierraConexion
Set oConecta = Nothing
Set oVar = Nothing
End Function

Public Function CargarDatosObjetosFteIngreso(ByRef prsTipoFte As ADODB.Recordset, _
                                            ByRef prsMoneda As ADODB.Recordset, _
                                            ByRef prsTipoCul As ADODB.Recordset, _
                                            ByRef prsUnidad As ADODB.Recordset, _
                                            ByRef prsFIDep As ADODB.Recordset, _
                                            ByRef prsFIInd As ADODB.Recordset, _
                                            ByRef prsFICos As ADODB.Recordset, _
                                            Optional ByVal pcNumFte As String = "")

Dim oConstante  As COMDConstantes.DCOMConstantes
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    On Error GoTo ErrorCargarDatosObjetosFteIngreso
    
    Set oConstante = New COMDConstantes.DCOMConstantes
    
    Set prsTipoFte = oConstante.RecuperaConstantes(gPersFteIngresoTipo)
    Set prsMoneda = oConstante.RecuperaConstantes(gMoneda)
    Set prsTipoCul = oConstante.RecuperaConstantes(1046)
    Set prsUnidad = oConstante.RecuperaConstantes(1045)
    
    Set oConstante = Nothing
    
    If pcNumFte <> "" Then
    
        Set oConecta = New COMConecta.DCOMConecta
    
        oConecta.AbreConexion
        'sSql = " Select * from ( "
        sSql = " Select PF.cNumFuente, PF.dPersEval, PF.nPersIngCli, PF.nPersGastoFam, PF.nPersIngCon, PF.nPersOtrIng, PF.dPersFICaduca "
        sSql = sSql & " FROM PersFIDependiente PF Where PF.cNumFuente = '" & pcNumFte & "'"
        sSql = sSql & " ORDER BY PF.dPersEval asc "
        'sSql = sSql & " ) X Order by dPersEval asc"
        Set prsFIDep = oConecta.CargaRecordSet(sSql)
    
        sSql = " Select C.dFecEval, CN2.cConsDescripcion as sUniProd, CN1.cConsDescripcion as sTipoCultivo,  C.cNumFuente, C.nTpoCultivo, C.nMaquinaria, C.nJornales, C.nInsumos, C.nPesticidas, C.nOtros, C.nHectareas,C.nMantenimiento, "
        sSql = sSql & " C.nProduccion , C.nUniProd, C.nPreUni,C.nSiembra,C.nCosecha,C.nDesaAgricola,C.nPlanOtros "
        sSql = sSql & "  From PersFICostoProd C Inner Join Constante CN1 ON CN1.nConsCod = 1046 AND C.nTpoCultivo = CN1.nConsValor "
        sSql = sSql & "                 Inner Join Constante CN2 ON CN2.nConsCod = 1045 AND C.nUniProd = CN2.nConsValor "
        sSql = sSql & "  Where cNumFuente = '" & pcNumFte & "' "
        Set prsFICos = oConecta.CargaRecordSet(sSql)
    
        
        sSql = " Select * from ("
        sSql = sSql & " SELECT  top 7 PF.cNumFuente, PF.dPersEval, PF.dPersFICaduca, PF.nPersFIActivoDisp, "
        sSql = sSql & " PF.nPersFICtasxCobrar, PF.nPersFIInventarios, PF.nPersFIActivosFijos, "
        sSql = sSql & " PF.nPersFIProveedores, PF.nPersFICredCMACT, PF.nPersFICredOtros, "
        sSql = sSql & " PF.nPersFIVentas, PF.nPersFIPatrimonio, PF.nPersFICostoVentas, "
        sSql = sSql & " PF.nPersFIRecupCtasXCobrar, PF.nPersFIEgresosOtros, PF.nPersIngFam, "
        sSql = sSql & " PF.nPersEgrFam, PF.cUltimaActualizacion, isnull(pf.dfecEEFF,'')dfecEEFF "
        sSql = sSql & " FROM PersFIIndependiente PF Where PF.cNumFuente = '" & pcNumFte & "'"
        sSql = sSql & " ORDER BY PF.dPersEval desc"
        sSql = sSql & " ) X Order by dPersEval asc"
        
        Set prsFIInd = oConecta.CargaRecordSet(sSql)
        
        oConecta.CierraConexion
        Set oConecta = Nothing
    End If
    Exit Function
    
ErrorCargarDatosObjetosFteIngreso:
    Set oConstante = Nothing
    Set oConecta = Nothing
    Err.Raise Err.Number, "Cargar Datos Objetos Fuente Ingreso", Err.Description
End Function

Public Function ObtenerApoderados(ByVal psApoderados As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
  On Error GoTo derror
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "Select R.cPersCod, P.cPersNombre from RRHH R inner join Persona P ON R.cPersCod = P.cpersCod "
    sSql = sSql & " AND nRHEstado = 201 "
    sSql = sSql & " inner join RHCargos RC ON R.cPersCod = RC.cPersCod "
    sSql = sSql & " where  RC.cRHCargoCod in (" & psApoderados & ") AND RC.dRHCargoFecha = (select MAX(dRHCargoFecha) "
    sSql = sSql & " from RHCargos RHC2 where RHC2.cPersCod = RC.cPersCod)"
    sSql = sSql & " order by P.cPersNombre "
    Set rs = oConecta.CargaRecordSet(sSql)
        Set ObtenerApoderados = rs
    Set rs = Nothing
    Exit Function
derror:
    Err.Raise Err.Number, "Error al cargar Datos", Err.Description
End Function

Public Function ObtenerCodPersPig(ByVal psCodigo As String) As ADODB.Recordset
   Dim sql As String
   Dim rs As New ADODB.Recordset
   Dim oConecta As New COMConecta.DCOMConecta
   
   On Error GoTo xError
        If oConecta.AbreConexion Then
           sql = "Select   replace (replace(substring(cpersnombre,1,case when PATINDEX ( '%,%', cpersnombre )<>0 then PATINDEX ( '%,%', cpersnombre )-1 else PATINDEX ( '%,%', cpersnombre ) end),'/',' '),'\',' ')  as cPersApellido, "
           sql = sql & "  substring(cpersnombre,PATINDEX ( '%,%', cpersnombre )+1,len(cpersnombre)) as cPersNombre "
           sql = sql & " from persona where cperscod='" & psCodigo & "'"
           Set rs = oConecta.CargaRecordSet(sql)
           Set ObtenerCodPersPig = rs
        End If
        oConecta.CierraConexion
   Exit Function
xError:
     Set ObtenerCodPersPig = Nothing
     oConecta.CierraConexion
     Err.Raise Err.Number, "Error al cargar Datos", Err.Description
End Function


'Public Function GeneraImpresionFuenteIngreso(ByVal psPersCod As String, ByVal pnIndice As Integer, ByVal psNomAge As String, _
'                                            ByVal pdFecSis As Date, ByVal pbCostoProd As Boolean, ByVal pnTipo As Integer, _
'                                            ByRef sCadImp As String)
'
'Dim oPersona As COMDPersona.DCOMPersona
'
'    On Error GoTo ErrorGeneraImpresionFuenteIngreso
'
'    Set oPersona = New COMDPersona.DCOMPersona
'
'    Call oPersona.RecuperaPersona(psPersCod)
'
'    If pnTipo = 1 Then
'        Call oPersona.RecuperaFtesIngresoDependiente(pnIndice)
'        sCadImp = oPersona.GenerarImpresionFteIngresoDependiente(pnIndice, psNomAge, pdFecSis, pbCostoProd, "")
'    Else
'        Call oPersona.RecuperaFtesIngresoIndependiente(pnIndice)
'        sCadImp = oPersona.GenerarImpresionFteIngresoIndependiente(pnIndice, psNomAge, pdFecSis, pbCostoProd, "")
'    End If
'
'    Set oPersona = Nothing
'    Exit Function
'
'ErrorGeneraImpresionFuenteIngreso:
'    Err.Raise Err.Number, "Genera Impresion Fuente Ingreso", Err.Description
'
'End Function

'Public Function GenerarImpresionFteIngresoDependiente(ByVal oPersona As DCOMPersona, _
'                                ByVal pnIndex As Integer, _
'                                ByVal psNomAge As String, ByVal pdFecha As Date, ByVal pbCostoProd As Boolean, _
'                                Optional ByVal psNomCmac As String = "") As String
'
'Dim sCadImp As String
'Dim i As Integer
'Dim J As Integer
'Dim lnNegritaON As String
'Dim lnNegritaOFF As String
'
'
'Dim MatDatosFte As Variant
'Dim MatDatosComFte As Variant
'
'
'
'    lnNegritaON = Chr$(27) & Chr$(69)
'    lnNegritaOFF = Chr$(27) + Chr$(70)
'
'    sCadImp = ""
'    sCadImp = sCadImp & ofun.ImpreFormat(psNomCmac, 35, 35) & " - " & psNomAge
'    sCadImp = sCadImp & Space(25) & Format(oGen.FechaHora(pdFecha), "dd/mm/yyyy hh:mm:ss") & Chr$(10) & Chr$(10) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & "Cliente : " & oPersona.PersCodigo & Space(2) & ocad.PstaNombre(oPersona.NombreCompleto) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    'sCadImp = sCadImp & "Fte Ingreso : " & FtesIngreso(pnIndex).sRazSocDescrip & Chr$(10) & Chr$(10) & Chr$(10)
'
'    sCadImp = sCadImp & "Fte Ingreso : " & oPersona.FtesIngreso(pnIndex).sRazSocDescrip & Chr$(10) & Chr$(10)
'
''----------------------------------------------------------------------------------------------------
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("RAZON SOCIAL/NOMBRE COMERCIAL", 35, 0) & ofun.ImpreFormat("RUC", 14, 2) & ofun.ImpreFormat("INICIO ACT.", 12, 0) & ofun.ImpreFormat("PERS. JURIDICA", 24, 2) & ofun.ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    sCadImp = sCadImp & ofun.ImpreFormat(oPersona.FtesIngreso(pnIndex).sRazSocDescrip, 35, 0) & ofun.ImpreFormat(oPersona.FtesIngreso(pnIndex).cFtePersRuc, 14, 2) & ofun.ImpreFormat(Format(oPersona.FtesIngreso(pnIndex).dFtePersFecInicioAct, "dd/mm/yyyy"), 12, 0) & ofun.ImpreFormat(oPersona.FtesIngreso(pnIndex).cFtePersTipoPersJur, 24, 2) & ofun.ImpreFormat(oPersona.FtesIngreso(pnIndex).cFtePersTelefono, 10, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("ACTIVIDAD(CIUU)", 35, 0) & ofun.ImpreFormat("COND. LOCAL", 14, 2) & ofun.ImpreFormat("MAGNITUD EMP.", 35, 0) & ofun.ImpreFormat("Nª EMPLEADOS", 14, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    sCadImp = sCadImp & ofun.ImpreFormat(oPersona.FtesIngreso(pnIndex).cFtePersCIUU, 35, 0) & ofun.ImpreFormat(oPersona.FtesIngreso(pnIndex).cFtePersCondicionDomic, 14, 2) & ofun.ImpreFormat(oPersona.FtesIngreso(pnIndex).cFtePersMagnitudEmp, 35, 0) & ofun.ImpreFormat(Round(oPersona.FtesIngreso(pnIndex).nFtePersNroEmpleados), 14, 0) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("DIRECCION(Avenida, Calle, Jiron)", 35, 0) & ofun.ImpreFormat("DEPARTAMENTO", 12, 2) & ofun.ImpreFormat("PROVINCIA", 12, 2) & ofun.ImpreFormat("DISTRITO", 16, 2) & ofun.ImpreFormat("ZONA", 17, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersDireccion, 35, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersDpto, 12, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersProv, 12, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersDist, 16, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersZona, 17, 2) & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10) & Chr$(10)
''----------------------------------------------------------------------------------------------------
'
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & Space(20) & "EVOLUCION ECONOMICA FINANCIERA DEPENDIENTE" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & Space(23)
'
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
'        sCadImp = sCadImp & Format(FtesIngreso(pnIndex).FtesIngresoDep(i).dFecEval, "dd/mm/yyyy") & Space(3)
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            sCadImp = sCadImp & "Variacion"
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & ofun.ImpreFormat("Ingreso Familiar", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngFam, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngFam = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 5, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngFam / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngFam, "#0.00")), 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    sCadImp = sCadImp & ofun.ImpreFormat("Ingreso Conyugue", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngCony = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 5, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngCony, "#0.00")), 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    sCadImp = sCadImp & ofun.ImpreFormat("Otros Ingresos", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngOtros, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngOtros = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 5, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngOtros / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngOtros, "#0.00")), 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    sCadImp = sCadImp & ofun.ImpreFormat("Gastos Familiares", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nGastosFam, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nGastosFam = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 5, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nGastosFam / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nGastosFam, "#0.00")), 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("Saldo", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i).nIngFam + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(i).nGastosFam, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngCony + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngFam + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nGastosFam = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 5, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(CDbl(Format((FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i).nIngFam + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(i).nGastosFam) / _
'                                (FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngCony + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngFam + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nGastosFam), "#0.00")), 8, 2)
'
'            End If
'        End If
'    Next i
'
'    'IMPRIME COSTO DE PRODUCCION
'    '-------------------------------------
'    If pbCostoProd = True Then
'            sCadImp = sCadImp & Chr$(12)
'            sCadImp = sCadImp & Chr$(10) & Chr$(10)
'            sCadImp = sCadImp & String(100, "-") & Chr$(10)
'            sCadImp = sCadImp & Space(30) & "COSTO DE PRODUCCION" & Chr$(10)
'            sCadImp = sCadImp & String(100, "-") & Chr$(10)
'            sCadImp = sCadImp & ofun.ImpreFormat("Evaluacion", 16)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).dFecEval, 14, 0)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ofun.ImpreFormat("Tipo Cultivo", 20)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).sTpoCultivo, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ofun.ImpreFormat("DESTINO", 20)
'            Dim sDestino As String
'
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                If FtesIngreso(pnIndex).CostoProd(i).nSiembra = 1 Then
'        '            If Len(sDestino) = 0 Then
'        '                sDestino = "SIEMBRA"
'        '            Else
'        '                sDestino = sDestino & ",SIEMBRA"
'                        sCadImp = sCadImp & ofun.ImpreFormat("SIEMBRA", 13)
'                        sCadImp = sCadImp & Chr$(10)
'        '            End If
'                 End If
'
'                If FtesIngreso(pnIndex).CostoProd(i).nMantenimiento = 1 Then
'        '            If Len(sDestino) = 0 Then
'        '                sDestino = "MANTENIMIENTO"
'        '            Else
'        '                sDestino = sDestino & ",MANTENIMIENTO"
'        '                sCadImp = sCadImp & Chr$(10)
'        '            End If
'                        sCadImp = sCadImp & Space(22) & ofun.ImpreFormat("MANTENIMIENTO", 13)
'                        sCadImp = sCadImp & Chr$(10)
'                End If
'                If FtesIngreso(pnIndex).CostoProd(i).nDesAgricola = 1 Then
'        '             If Len(sDestino) = 0 Then
'        '                sDestino = "DESAGRICOLA"
'        '             Else
'        '                sDestino = sDestino & ",DESAGRICOLA"
'        '                sCadImp = sCadImp & Chr$(10)
'        '             End If
'                        sCadImp = sCadImp & Space(22) & ofun.ImpreFormat("DESAGRICOLA", 13)
'                        sCadImp = sCadImp & Chr$(10)
'
'
'                End If
'                If FtesIngreso(pnIndex).CostoProd(i).nCosecha = 1 Then
'        '            If Len(sDestino) = 0 Then
'        '                sDestino = "COSECHA"
'        '            Else
'        '                sDestino = sDestino & ",COSECHA"
'        '                sCadImp = sCadImp & Chr$(10)
'        '            End If
'                        sCadImp = sCadImp & Space(22) & ofun.ImpreFormat("COSECHA", 13)
'                        sCadImp = sCadImp & Chr$(10)
'
'                 End If
'                If FtesIngreso(pnIndex).CostoProd(i).nPlanOtros = 1 Then
'        '            If Len(sDestino) = 0 Then
'        '                sDestino = "OTROS"
'        '            Else
'        '                sDestino = sDestino & ",OTROS"
'        '            End If
'                        sCadImp = sCadImp & Space(22) & ofun.ImpreFormat("OTROS", 13)
'                        'sCadImp = sCadImp & Chr$(10)
'
'                End If
'                'sCadImp = sCadImp & oFun.ImpreFormat(sDestino, 60)
'            Next i
'
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ofun.ImpreFormat("Maquinaria", 12)
'
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nMaquinaria, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ofun.ImpreFormat("Jornales", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nJornales, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ofun.ImpreFormat("Insumos", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nInsumos, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ofun.ImpreFormat("Pesticidas", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ofun.ImpreFormat("Otros", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ofun.ImpreFormat("Total", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10) & Chr$(10)
'
'            sCadImp = sCadImp & ofun.ImpreFormat("Hectareas", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ofun.ImpreFormat("Produccion", 18)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ofun.ImpreFormat(Format(FtesIngreso(pnIndex).CostoProd(i).nProduccion, "#0.00") & " " & FtesIngreso(pnIndex).CostoProd(i).sUniProd, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ofun.ImpreFormat("Precio Uni.", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPreUni, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ofun.ImpreFormat("Ingresos", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ofun.ImpreFormat("Egresos", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros), 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ofun.ImpreFormat("Utilidad", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion) - (FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros)), 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'            sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    End If
'
'
'
'    ' ---------------------------------------------------------------------------------------------------------------
'    ' Muestra los datos de las referencias personales de la fte como persona ----------------------------------------
'    sCadImp = sCadImp & Chr$(10) & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & "REFERENCIAS PERSONALES" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    sCadImp = sCadImp & FtesIngreso(pnIndex).sComentarios & Chr$(10)
'
'    sCadImp = sCadImp & Chr$(10)
'
'    ' Muestra los datos de las referencias bancarias de la fte como persona ----------------------------------------
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("REFERENCIAS BANCARIAS", 42, 0) & ofun.ImpreFormat("NRO.TARJETA", 22, 2) & ofun.ImpreFormat("NRO.CUENTA", 22, 2) & ofun.ImpreFormat("MONTO", 10, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    'MatDatosFte = RecuperaRefBancariaPersFte(FtesIngreso(pnIndex).sFteIngreso)
'    MatDatosFte = RecuperaRefBancariaPersFte(sPersCod)
'
'    For i = 0 To UBound(MatDatosFte) - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(MatDatosFte(i, 0) & "-" & MatDatosFte(i, 1), 40) & ofun.ImpreFormat(MatDatosFte(i, 3), 22, 2) & ofun.ImpreFormat(MatDatosFte(i, 2), 22, 2) & ofun.ImpreFormat(MatDatosFte(i, 4), 10, 2) & Chr$(10)
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'
'   ' Muestra los datos de las referencias comerciales de la fte como persona ----------------------------------------
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("REFERENCIAS COMERCIALES", 66, 0) & ofun.ImpreFormat("TIPO RELAC. COMERCIAL", 22, 2) & ofun.ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    'MatDatosComFte = RecuperaRefComercialPersFte(FtesIngreso(pnIndex).sFteIngreso)
'    MatDatosComFte = RecuperaRefComercialPersFte(sPersCod)
'    For i = 0 To UBound(MatDatosComFte) - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(MatDatosComFte(i, 3) & "-" & MatDatosComFte(i, 0), 64) & ofun.ImpreFormat(MatDatosComFte(i, 2), 22, 2) & ofun.ImpreFormat(MatDatosComFte(i, 1), 10, 2) & Chr$(10)
'    Next i
'
'    ' ---------------------------------------------------------------------------------------------------------------
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    GenerarImpresionFteIngresoDependiente = sCadImp
'End Function
'
'Public Function GenerarImpresionFteIngresoIndependiente(ByVal pnIndex As Integer, _
'    ByVal psNomAge As String, ByVal pdFecha As Date, ByVal pbCostoProd As Boolean, _
'    Optional ByVal psNomCmac As String = "") As String
'
'Dim sCadImp As String
'Dim MatDatos As Variant
'Dim i As Integer
'Dim lnNegritaON As String
'Dim lnNegritaOFF As String
'
'Dim MatDatosFte As Variant
'Dim MatDatosComFte As Variant
'Dim bProd As Boolean
'
'
'
'    lnNegritaON = Chr$(27) & Chr$(69)
'    lnNegritaOFF = Chr$(27) + Chr$(70)
'
'    MatDatos = RecuperaHistoriaCredPersona(sPersCod)
'
'    sCadImp = ""
'
'    ' Comprime todo el reporte ----------------------------------------------------------------------
'    sCadImp = sCadImp & Chr$(27) & Chr$(64)
'    'sCadImp = sCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
'    sCadImp = sCadImp & Chr$(27) & Chr$(65) & Chr$(9)     'Espaciado
'
'    sCadImp = sCadImp & Chr$(27) & Chr$(67) & Chr$(80)  'Longitud de página a 80 líneas'
'    sCadImp = sCadImp & Chr$(27) & Chr$(103)   'Tamaño 10 cpi
'    'sCadImp = sCadImp & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
'    sCadImp = sCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
'    'sCadImp = sCadImp & Chr$(27) + Chr$(18) ' cancela condensada
'    'sCadImp = sCadImp & Chr$(27) + Chr$(72) ' desactiva negrita
'    'lsCadImp = lsCadImp & Chr$(27) & Chr$(108) & Chr$(20)   'Margen  izquierdo  20 caracteres
'    '-------------------------------------------------------------------------------------------------
'
'    'sCadImp = sCadImp & oFun.ImpreFormat(psNomCmac, 35, 35) & psNomAge
'    sCadImp = sCadImp & ofun.ImpreFormat(psNomCmac, 35, 0) & psNomAge
'    sCadImp = sCadImp & Space(25) & Format(pdFecha, "dd/mm/yyy hh:mm:ss") & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & "Cliente : " & sPersCod & Space(2) & ocad.PstaNombre(sNombreCompleto) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    sCadImp = sCadImp & "Fte Ingreso : " & FtesIngreso(pnIndex).sRazSocDescrip & Chr$(10)
'
'
''----------------------------------------------------------------------------------------------------
'    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("RAZON SOCIAL/NOMBRE COMERCIAL", 35, 0) & ofun.ImpreFormat("RUC", 14, 2) & ofun.ImpreFormat("INICIO ACT.", 12, 0) & ofun.ImpreFormat("PERS. JURIDICA", 24, 2) & ofun.ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).sRazSocDescrip, 35, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersRuc, 14, 2) & ofun.ImpreFormat(Format(FtesIngreso(pnIndex).dFtePersFecInicioAct, "dd/mm/yyyy"), 12, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersTipoPersJur, 24, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersTelefono, 10, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("ACTIVIDAD(CIUU)", 35, 0) & ofun.ImpreFormat("COND. LOCAL", 14, 2) & ofun.ImpreFormat("MAGNITUD EMP.", 35, 0) & ofun.ImpreFormat("Nª EMPLEADOS", 14, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersCIUU, 35, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersCondicionDomic, 14, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersMagnitudEmp, 35, 0) & ofun.ImpreFormat(Round(FtesIngreso(pnIndex).nFtePersNroEmpleados), 14, 0) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("DIRECCION(Avenida, Calle, Jiron)", 35, 0) & ofun.ImpreFormat("DEPARTAMENTO", 12, 2) & ofun.ImpreFormat("PROVINCIA", 12, 2) & ofun.ImpreFormat("DISTRITO", 16, 2) & ofun.ImpreFormat("ZONA", 17, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersDireccion, 35, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersDpto, 12, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersProv, 12, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersDist, 16, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersZona, 17, 2) & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
''----------------------------------------------------------------------------------------------------
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & Space(20) & "HISTORIA CREDITICA" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
'     sCadImp = sCadImp & String(100, "-") & Chr$(10)
'
'    sCadImp = sCadImp & "MONEDA             : " & IIf(FtesIngreso(pnIndex).sMoneda = "1", "SOLES", "DOLARES")
'
'    If Not IsArray(MatDatos) Then
'        sCadImp = sCadImp & "NO POSEE CREDITOS" & Chr$(10)
'    Else
'        'sCadImp = sCadImp & "CREDITO No         : "
'        'For I = 0 To UBound(MatDatos) - 1
'        '    sCadImp = sCadImp & MatDatos(I, 0) & Space(2)
'        'Next I
'
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Estado Actual      : "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & MatDatos(i, 1) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Fecha Vigencia     : "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & MatDatos(i, 2) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Fecha Cancelacion  : "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & MatDatos(i, 3) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Monto(Sol/Sug/Apr) : "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & MatDatos(i, 4) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Cuotas/Plazo(Dias) : "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & MatDatos(i, 5) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Cuota (Sol/Sug/Apr): "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & MatDatos(i, 6) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Nota               : "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & MatDatos(i, 7) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Analista           : "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & MatDatos(i, 8) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'    End If
'    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & Space(20) & "EVOLUCION ECONOMICA FINANCIERA" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & Space(25)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & Format(FtesIngreso(pnIndex).FtesIngresoInd(i).dFecEval, "dd/mm/yyyy") & Space(3)
'    Next i
'    sCadImp = sCadImp & "Variacion" & Chr$(10)
'    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("ACTIVO CIRCULANTE", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios) = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios), 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    sCadImp = sCadImp & ofun.ImpreFormat("Disponible", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ofun.ImpreFormat("Cuentas x Cobrar", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ofun.ImpreFormat("Inventario", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("ACTIVO FIJO", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ofun.ImpreFormat("TOTAL ACTIVO", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo), 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ofun.ImpreFormat("PASIVO", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact = 0 Then
'
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact), 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    sCadImp = sCadImp & ofun.ImpreFormat("Proveedores", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ofun.ImpreFormat("Otros Prestamos", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ofun.ImpreFormat("Prestamos CMACT", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("PATRIMONIO", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ofun.ImpreFormat("INGRESOS UE", 20)
'    Dim nIng As Currency
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
''        sCadImp = sCadImp & oFun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob, 8, 2)
''
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob, 8, 2)
'        nIng = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
'
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
''            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob = 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob = 0 Then
'
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
''                sCadImp = sCadImp & oFun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob), 8, 2)
'               sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob), 8, 2)
'
'            End If
'        End If
'    Next i
'
'    'CMACICA_CSTS - 10/11/2003 - SE CAMBIO EL ORDEN (VENTAS, RECUP. X COB., INGRESOS FAMILIARES) -----------------
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    sCadImp = sCadImp & ofun.ImpreFormat("Ventas", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas, 8, 2)
'        If i > 0 And FtesIngreso(pnIndex).nNumFtesIndep - 1 = i Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas, 8, 2)
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ofun.ImpreFormat("Recup. Ctas x Cob", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob, 8, 2)
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
''    sCadImp = sCadImp & oFun.ImpreFormat("Ingresos familiares", 20)
''    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
''        sCadImp = sCadImp & oFun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam, 8, 2)
''        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
''            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam = 0 Then
''                sCadImp = sCadImp & oFun.ImpreFormat(0, 8, 2)
''            Else
''                sCadImp = sCadImp & oFun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam, 8, 2)
''            End If
''        End If
''    Next i
'
''-----------------------------------------------------------------------------------------------------------------------
'    Dim nEg As Currency
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("EGRESOS UE", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
''        sCadImp = sCadImp & oFun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
'         nEg = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
''            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos = 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'
''                sCadImp = sCadImp & oFun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos), 8, 2)
'                sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos), 8, 2)
'            End If
'        End If
'    Next i
'
'    'CMACICA_CSTS - 21/11/2003 - SE CAMBIO EL ORDEN (COMPRA DE MERCADERIAS, OTROS EGRESOS, GASTOS FAMILIARES) ---------
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    'sCadImp = sCadImp & oFun.ImpreFormat("Compra Mercaderia", 20)
'    sCadImp = sCadImp & ofun.ImpreFormat("Costo de Ventas", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad, 8, 2)
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ofun.ImpreFormat("Otros Egresos", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos, 8, 2)
'            End If
'        End If
'    Next i
''
''    sCadImp = sCadImp & Chr$(10)
''    sCadImp = sCadImp & oFun.ImpreFormat("Gastos Familiares", 20)
''    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
''        sCadImp = sCadImp & oFun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam, 8, 2)
''        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
''            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam = 0 Then
''                sCadImp = sCadImp & oFun.ImpreFormat(0, 8, 2)
''            Else
''                sCadImp = sCadImp & oFun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam, 8, 2)
''            End If
''        End If
''    Next i
'
'' ---------------------------------------------------------------------------------------------------------------------
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("UTILIDAD NETA", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
''        sCadImp = sCadImp & oFun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob - _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam - _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad - _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob - _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad - _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
'
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
''            If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - _
''                        (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
''                         FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
''                         FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos) = 0 Then
'            If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - _
'                        (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                         FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos) = 0 Then
'
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'
'            Else
''                sCadImp = sCadImp & oFun.ImpreFormat(((FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
''                            FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
''                            FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) - _
''                            (FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
''                             FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
''                             FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos)) _
''                            / ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
''                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
''                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - _
''                            (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
''                             FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
''                             FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos)), 8, 2)
'                sCadImp = sCadImp & ofun.ImpreFormat(((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) - _
'                            (FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                             FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos)) _
'                            / ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - _
'                            (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                             FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos)), 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    'sCadImp = sCadImp & Chr$(10)
'    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & Space(20) & "INGRESOS-EGRESOS FAMILIARES" & Chr$(10)
'    sCadImp = sCadImp & ofun.ImpreFormat("Ingresos familiares", 20)
'    Dim nIngFam As Currency
'
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam, 8, 2)
'        nIngFam = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ofun.ImpreFormat("Gastos Familiares", 20)
'    Dim nEgrFam As Currency
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam, 8, 2)
'        nEgrFam = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ofun.ImpreFormat("DISPONIBLE", 20)
'    sCadImp = sCadImp & ofun.ImpreFormat(nIngFam + nIng - (nEgrFam + nEg), 8, 2)
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & Space(20) & "RATIOS DE EVALUACION" & Chr$(10)
'    sCadImp = sCadImp & "LIQUIDEZ" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    'Activo Circulante / Pasivo
'    Dim nActCir As Double
'    Dim nPasivo As Double
'    Dim nActCirAnt As Double
'    Dim nPasivoAnt As Double
'    Dim Res1 As Double
'    Dim Res2 As Double
'    sCadImp = sCadImp & ofun.ImpreFormat("Liquidez", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nActCir = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
'        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
'        If nPasivo = 0 Then
'            Res1 = 0
'        Else
'            Res1 = nActCir / nPasivo
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nActCirAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
'            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
'            If nPasivoAnt = 0 Then
'                Res2 = 0
'            Else
'                Res2 = nActCirAnt / nPasivoAnt
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    'Activo Circulante - Pasivo
'    sCadImp = sCadImp & ofun.ImpreFormat("Capital Trabajo", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nActCir = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
'        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
'        Res1 = nActCir - nPasivo
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nActCirAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
'            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
'            Res2 = nActCirAnt - nPasivoAnt
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & "GESTION" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    'Costo de Ventas / Inventario
'    sCadImp = sCadImp & ofun.ImpreFormat("Rotacion Inventario", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios = 0 Then
'            Res1 = 0
'        Else
'            Res1 = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios = 0 Then
'                Res2 = 0
'            Else
'                Res2 = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    'Recup Ctas x Cob / (Ventas / 30)
'    sCadImp = sCadImp & ofun.ImpreFormat("Periodo de Cobro", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
'            Res1 = 0
'        Else
'            Res1 = FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob / (FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas / 30)
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
'                Res2 = 0
'            Else
'                Res2 = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas / 30)
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    Dim nIngresos As Double
'    Dim nEgresos As Double
'    Dim nIngresosAnt As Double
'    Dim nEgresosAnt As Double
'
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & "RENTABILIDAD" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    'Utilidad Operat =  Ingresos - Egresos
'    'Utilidad Operat / Patrimonio
'    sCadImp = sCadImp & ofun.ImpreFormat("Rentab. Patrimonial", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
'        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
'
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
'            Res1 = 0
'        Else
'            Res1 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob
'            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
'                Res2 = 0
'            Else
'                Res2 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio
'            End If
'
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    Dim nTotActivo As Double
'    Dim nTotActivoAnt As Double
'    sCadImp = sCadImp & ofun.ImpreFormat("Rentab. Activo", 20)
'    'Utilidad Operativa / Total Activo
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
'        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
'        nTotActivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
'        If nTotActivo = 0 Then
'            Res1 = 0
'        Else
'            Res1 = (nIngresos - nEgresos) / nTotActivo
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob
'            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
'            nTotActivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo
'            If nTotActivoAnt = 0 Then
'                Res2 = 0
'            Else
'                Res2 = (nIngresosAnt - nEgresosAnt) / nTotActivoAnt
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    sCadImp = sCadImp & ofun.ImpreFormat("Rentab. Ventas", 20)
'    'Utilidad Operativa / Ventas
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
'        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
'            Res1 = 0
'        Else
'            Res1 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob
'            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
'                Res2 = 0
'            Else
'                Res2 = (nIngresosAnt - nEgresosAnt) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    'Ventas Totales = Ventas + Recup Ctas x Cob
'    'Margen Utilidad = Ventas Totales - Costo Ventas / Ventas Totales
'    Dim nVentasTot As Double
'    Dim nVentasTotAnt As Double
'    sCadImp = sCadImp & ofun.ImpreFormat("Margen Utilidad", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nVentasTot = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
'        If nVentasTot = 0 Then
'            Res1 = 0
'        Else
'            Res1 = (nVentasTot - FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad) / nVentasTot
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nVentasTotAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
'            If nVentasTotAnt = 0 Then
'                Res2 = 0
'            Else
'                Res2 = (nVentasTotAnt - FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad) / nVentasTotAnt
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & "ENDEUDAMIENTO" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    sCadImp = sCadImp & ofun.ImpreFormat("Endeudamiento Pat.", 20)
'    'Endeudamiento = Pasivo / Patrimonio
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
'
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
'            Res1 = 0
'        Else
'            Res1 = nPasivo / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
'
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
'                Res2 = 0
'            Else
'                Res2 = nPasivoAnt / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    sCadImp = sCadImp & ofun.ImpreFormat("Endeudamiento", 20)
'    'Endeudamiento = Pasivo / TotalActivo
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nTotActivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
'        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
'
'        If nTotActivo = 0 Then
'            Res1 = 0
'        Else
'            Res1 = nPasivo / nTotActivo
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nTotActivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
'            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
'            If nTotActivoAnt = 0 Then
'                Res2 = 0
'            Else
'                Res2 = nPasivoAnt / nTotActivoAnt
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'
'    If pbCostoProd = True Then
'
'        'IMPRIME COSTO DE PRODUCCION
'        '-------------------------------------
'        sCadImp = sCadImp & Chr$(12)
'        sCadImp = sCadImp & String(100, "-") & Chr$(10)
'        sCadImp = sCadImp & Space(30) & "COSTO DE PRODUCCION" & Chr$(10)
'        sCadImp = sCadImp & String(100, "-") & Chr$(10)
'        sCadImp = sCadImp & ofun.ImpreFormat("Evaluacion", 16)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).dFecEval, 14, 0)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ofun.ImpreFormat("Tipo Cultivo", 20)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).sTpoCultivo, 13)
'        Next i
'
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ofun.ImpreFormat("Destino", 20)
'
'        Dim sDestino As String
'        bProd = False
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            If FtesIngreso(pnIndex).CostoProd(i).nSiembra = 1 Then
'    '            If Len(sDestino) = 0 Then
'    '                sDestino = "SIEMBRA"
'    '            Else
'    '                sDestino = sDestino & ",SIEMBRA"
'                    bProd = True
'                    sCadImp = sCadImp & ofun.ImpreFormat("SIEMBRA", 13)
'                    'sCadImp = sCadImp & Chr$(10)
'    '            End If
'             End If
'
'            If FtesIngreso(pnIndex).CostoProd(i).nMantenimiento = 1 Then
'    '            If Len(sDestino) = 0 Then
'    '                sDestino = "MANTENIMIENTO"
'    '            Else
'    '                sDestino = sDestino & ",MANTENIMIENTO"
'    '                sCadImp = sCadImp & Chr$(10)
'    '            End If
'                If bProd = True Then
'                    sCadImp = sCadImp & Space(7) & ofun.ImpreFormat("MANTENIMIENTO", 13)
'                Else
'                    sCadImp = sCadImp & ofun.ImpreFormat("MANTENIMIENTO", 13)
'                    bProd = True
'                End If
'                   ' sCadImp = sCadImp & Chr$(10)
'            End If
'            If FtesIngreso(pnIndex).CostoProd(i).nDesAgricola = 1 Then
'    '             If Len(sDestino) = 0 Then
'    '                sDestino = "DESAGRICOLA"
'    '             Else
'    '                sDestino = sDestino & ",DESAGRICOLA"
'    '                sCadImp = sCadImp & Chr$(10)
'    '             End If
'                  If bProd = True Then
'                    sCadImp = sCadImp & Space(7) & ofun.ImpreFormat("DESAGRICOLA", 13)
'                  Else
'                    sCadImp = sCadImp & ofun.ImpreFormat("DESAGRICOLA", 13)
'                    bProd = True
'                  End If
'                    'sCadImp = sCadImp & Chr$(10)
'
'
'            End If
'            If FtesIngreso(pnIndex).CostoProd(i).nCosecha = 1 Then
'    '            If Len(sDestino) = 0 Then
'    '                sDestino = "COSECHA"
'    '            Else
'    '                sDestino = sDestino & ",COSECHA"
'    '                sCadImp = sCadImp & Chr$(10)
'    '            End If
'                If bProd = True Then
'                    sCadImp = sCadImp & Space(7) & ofun.ImpreFormat("COSECHA", 13)
'                Else
'                    sCadImp = sCadImp & ofun.ImpreFormat("COSECHA", 13)
'                    bProd = True
'                End If
'                    'sCadImp = sCadImp & Chr$(10)
'
'             End If
'            If FtesIngreso(pnIndex).CostoProd(i).nPlanOtros = 1 Then
'    '            If Len(sDestino) = 0 Then
'    '                sDestino = "OTROS"
'    '            Else
'    '                sDestino = sDestino & ",OTROS"
'    '            End If
'                    If bProd = True Then
'                        sCadImp = sCadImp & Space(7) & ofun.ImpreFormat("OTROS", 13)
'                    Else
'                        sCadImp = sCadImp & ofun.ImpreFormat("OTROS", 13)
'                        bProd = True
'                    End If
'                    'sCadImp = sCadImp & Chr$(10)
'
'            End If
'            'sCadImp = sCadImp & oFun.ImpreFormat(sDestino, 60)
'        Next i
'
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ofun.ImpreFormat("Maquinaria", 12)
'
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nMaquinaria, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ofun.ImpreFormat("Jornales", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nJornales, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ofun.ImpreFormat("Insumos", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nInsumos, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ofun.ImpreFormat("Pesticidas", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ofun.ImpreFormat("Otros", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ofun.ImpreFormat("Total", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10) & Chr$(10)
'
'        sCadImp = sCadImp & ofun.ImpreFormat("Hectareas", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ofun.ImpreFormat("Produccion", 18)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ofun.ImpreFormat(Format(FtesIngreso(pnIndex).CostoProd(i).nProduccion, "#0.00") & " " & FtesIngreso(pnIndex).CostoProd(i).sUniProd, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ofun.ImpreFormat("Precio Uni.", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPreUni, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ofun.ImpreFormat("Ingresos", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ofun.ImpreFormat("Egresos", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros), 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ofun.ImpreFormat("Utilidad", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion) - (FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros)), 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & String(100, "-") & Chr$(10)
' End If
'    ' ---------------------------------------------------------------------------------------------------------------
'    ' Muestra los datos de las referencias personales de la fte como persona ----------------------------------------
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & "REFERENCIAS PERSONALES" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    sCadImp = sCadImp & FtesIngreso(pnIndex).sComentarios & Chr$(10)
'    sCadImp = sCadImp & Chr$(10)
'
'    ' Muestra los datos de las referencias bancarias de la fte como persona ----------------------------------------
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("REFERENCIAS BANCARIAS", 42, 0) & ofun.ImpreFormat("NRO.TARJETA", 22, 2) & ofun.ImpreFormat("NRO.CUENTA", 22, 2) & ofun.ImpreFormat("MONTO", 10, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    MatDatosFte = RecuperaRefBancariaPersFte(FtesIngreso(pnIndex).sFteIngreso)
'
'    For i = 0 To UBound(MatDatosFte) - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(MatDatosFte(i, 0) & "-" & MatDatosFte(i, 1), 40) & ofun.ImpreFormat(MatDatosFte(i, 3), 22, 2) & ofun.ImpreFormat(MatDatosFte(i, 2), 22, 2) & ofun.ImpreFormat(MatDatosFte(i, 4), 10, 2) & Chr$(10)
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'
'   ' Muestra los datos de las referencias comerciales de la fte como persona ----------------------------------------
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ofun.ImpreFormat("REFERENCIAS COMERCIALES", 66, 0) & ofun.ImpreFormat("TIPO RELAC. COMERCIAL", 22, 2) & ofun.ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    MatDatosComFte = RecuperaRefComercialPersFte(FtesIngreso(pnIndex).sFteIngreso)
'
'    For i = 0 To UBound(MatDatosComFte) - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(MatDatosComFte(i, 3) & "-" & MatDatosComFte(i, 0), 64) & ofun.ImpreFormat(MatDatosComFte(i, 2), 22, 2) & ofun.ImpreFormat(MatDatosComFte(i, 1), 10, 2) & Chr$(10)
'    Next i
'
'    ' ---------------------------------------------------------------------------------------------------------------
'    sCadImp = sCadImp & Chr$(10)
'
'    GenerarImpresionFteIngresoIndependiente = sCadImp
'
'End Function

Public Function ObtenerRatiosFinancieros(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = "SELECT LiquidezG= (nPersFIActivoDisp+nPersFICtasxCobrar+nPersFIInventarios)/" & _
            "(nPersFIProveedores+nPersFICredCMACT+nPersFICredOtros), " & _
        "PruebaA= (nPersFIActivoDisp+nPersFICtasxCobrar)/ (nPersFIProveedores+nPersFICredCMACT+nPersFICredOtros),  " & _
        "CapitalT= (nPersFIActivoDisp+nPersFICtasxCobrar+nPersFIInventarios)- " & _
            "(nPersFIProveedores+nPersFICredCMACT+nPersFICredOtros), " & _
        "RotacionC= CASE WHEN nPersFICtasxCobrar=0 THEN 0 ELSE nPersFIVentas/nPersFICtasxCobrar END, " & _
        "PeriodoC= CASE WHEN nPersFIVentas=0 THEN 0 ELSE nPersFICtasxCobrar*365/nPersFIVentas END, " & _
        "RotacionP= CASE WHEN nPersFICtasxCobrar=0 THEN 0 ELSE nPersEgrFam /nPersFICtasxCobrar END, " & _
        "PeriodoP= CASE WHEN nPersEgrFam= 0 THEN 0 ELSE nPersFICtasxCobrar*365/ nPersEgrFam END, " & _
        "RotacionIV= CASE WHEN nPersFIInventarios=0 THEN 0 ELSE nPersFICostoVentas/nPersFIInventarios END, " & _
        "RotacionID= CASE WHEN nPersFICostoVentas=0 THEN 0 ELSE nPersFIInventarios*365/nPersFICostoVentas END, " & _
        "RotacionA= CASE WHEN nPersFIActivosFijos=0 THEN 0 ELSE nPersFIVentas/nPersFIActivosFijos END, " & _
        "MargenB= CASE WHEN nPersFIVentas=0 THEN 0 " & _
        "    ELSE(   (nPersFIVentas+nPersFIRecupCtasXCobrar+nPersIngFam)- " & _
        "    (nPersFICostoVentas+nPersEgrFam+nPersFIEgresosOtros)) /nPersFIVentas END, " & _
        "GradoE=  (nPersFIProveedores+nPersFICredCMACT+nPersFICredOtros)/ " & _
        "    (nPersFIActivoDisp+nPersFICtasxCobrar+nPersFIInventarios+nPersFIActivosFijos)," & _
        "GradoP= nPersFIPatrimonio/(nPersFIActivoDisp+nPersFICtasxCobrar+nPersFIInventarios+nPersFIActivosFijos), " & _
        "EndeudamientoP= CASE WHEN nPersFIPatrimonio=0 THEN 0 " & _
        "    ELSE (nPersFIProveedores+nPersFICredCMACT+nPersFICredOtros)/ nPersFIPatrimonio END, " & _
        "RentabilidadP= 0, " & _
        "RentabilidadV=0 " & _
        "FROM PersFteIngreso F INNER JOIN PersFIindependiente FI ON F.cNumFuente=FI.cNumFuente WHERE cPersCod='" & psPersCod & "'"
        
Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set ObtenerRatiosFinancieros = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function ObtenerDatosReporteFteIngreso(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = "SELECT cConsDescripcion=ISNULL(CON.cConsDescripcion,''),cCIIUDescripcion=ISNULL(CIIU.cCIIUDescripcion,''), cPersIDnro=ISNULL(cPersIDnro,''), " & _
       " nPersJurEmpleados=ISNULL(nPersJurEmpleados,0),Magnitud=ISNULL(MAG.cConsDescripcion,''), " & _
       " cDep= ISNULL((SELECT RTRIM(cUbiGeoDescripcion) FROM UbicacionGeografica WHERE LEFT(cUbiGeoCod,1)='1' AND SUBSTRING(cUbiGeoCod,2,2)= SUBSTRING(cPersDireccUbiGeo,2,2)),'')," & _
       " cProv= ISNULL((SELECT RTRIM(cUbiGeoDescripcion) FROM UbicacionGeografica WHERE LEFT(cUbiGeoCod,1)='2' AND SUBSTRING(cUbiGeoCod,2,4)= SUBSTRING(cPersDireccUbiGeo,2,4)),''), " & _
       " cDist= ISNULL((SELECT RTRIM(cUbiGeoDescripcion) FROM UbicacionGeografica WHERE LEFT(cUbiGeoCod,1)='3' AND SUBSTRING(cUbiGeoCod,2,6)= SUBSTRING(cPersDireccUbiGeo,2,6)),''), " & _
       " cZon= ISNULL((SELECT RTRIM(cUbiGeoDescripcion) FROM UbicacionGeografica WHERE cUbiGeoCod = cPersDireccUbiGeo),'') " & _
       " FROM Persona P LEFT JOIN Constante CON ON CON.nConsValor=P.cPersDireccCondicion AND CON.nConsCod=1029 " & _
       " LEFT JOIN CIIU ON cCIIUCod=cPersCIIU LEFT JOIN PersId PD ON P.cPersCod=PD.cPerscod AND cPersIDTpo=2 " & _
       " LEFT JOIN PersonaJur PJ ON PJ.cPersCod=P.cPersCod LEFT JOIN Constante MAG ON MAG.nConsValor=PJ.cPersJurMagnitud AND MAG.nConscod=1004" & _
       " WHERE P.cPersCod='" & psPersCod & "'"
        
Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set ObtenerDatosReporteFteIngreso = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Private Sub Class_Terminate()
    Set ofun = Nothing
    Set oGen = Nothing
    Set ocad = Nothing
End Sub

'By Capi 13022008
Public Function ObtenerDatosPersonasLavDinero(ByVal sPersCodOrd As String, sPersCodRea As String, sPersCodBen As String, sPersCodTit As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loConecta As COMConecta.DCOMConecta
    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    lsSQL = " Exec stp_sel_ObtenerDatosPersonasLavDinero '" & sPersCodOrd & "','" & sPersCodRea & "','" & sPersCodBen & "','" & sPersCodTit & "'"
    Set ObtenerDatosPersonasLavDinero = loConecta.CargaRecordSet(lsSQL)
    loConecta.CierraConexion
    Set loConecta = Nothing
 End Function
'By Capi 28022008
Public Function ObtenerAcumuladoLavDinero(pnTipoCambio As Double, ByVal pcPeriodo As String, ByVal psPersCodTit As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loConecta As COMConecta.DCOMConecta
    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    lsSQL = " Exec stp_sel_ObtenerPersAcumuladoLavDinero " & pnTipoCambio & ",'" & pcPeriodo & "','" & psPersCodTit & "'"
    Set ObtenerAcumuladoLavDinero = loConecta.CargaRecordSet(lsSQL)
    loConecta.CierraConexion
    Set loConecta = Nothing
 End Function

'By Capi 28022008
Public Function ObtenerAgenCorrelativoLavDinero(ByVal psAgencia As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loConecta As COMConecta.DCOMConecta
    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    lsSQL = " Exec stp_sel_ObtenerAgenCorrelativoLavDinero '" & psAgencia & "'"
    Set ObtenerAgenCorrelativoLavDinero = loConecta.CargaRecordSet(lsSQL)
    loConecta.CierraConexion
    Set loConecta = Nothing
 End Function


'By Capi 28022008
Public Sub ActualizarAgenCorrelativoLavDinero(ByVal psAgencia As String)
    Dim lsSQL As String
    Dim loConecta As COMConecta.DCOMConecta
    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    lsSQL = " Exec stp_upd_ActualizarAgenCorrelativoLavDinero '" & psAgencia & "'"
    loConecta.Ejecutar (lsSQL)
    loConecta.CierraConexion
    Set loConecta = Nothing
 End Sub
 '**ALPA**13042008***********************************************************************
Public Function ObtenerDatosHojEvaluaci(NumFteIngreso As String, dfecha As Date) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = " Exec stp_sel_ObtenersHojaEvaluacion '" & NumFteIngreso & "','" & Format(dfecha, "yyyy/mm/dd") & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerDatosHojEvaluaci = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ObtenerDatosHojEvaluaci = Null
End Function
Public Function ObtenerDatosDocsPers(cCodPers As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = " Exec stp_sel_PersDocsTipo '" & cCodPers & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerDatosDocsPers = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ObtenerDatosDocsPers = Null
End Function
 '****************************************************************************************

'*** MAVM 13/04/2009
Public Function ValidarCodSBS(ByVal sCodSBS As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_sel_ValidarCodSBS '" & sCodSBS & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set ValidarCodSBS = lrDatos
    Set lrDatos = Nothing
End Function

'*** PEAC 20090715
Public Function RecuperaPersListaNegativa() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
   
    On Error GoTo ErrorRecuperaPersListaNegativa
    
    sSql = " exec stp_sel_RecuperaPersListaNegativa "
        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPersListaNegativa = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaPersListaNegativa:
    Set RecuperaPersListaNegativa = Nothing
    Err.Raise Err.Number, "Recupera Personas de un Tipo", Err.Description
End Function

'*** PEAC 20090715
Public Function eliminaperslistanegativa(ByVal pnTipoDoc As Integer, ByVal pcNumDoc As String, ByVal pcMovNro As String) 'JACA 20110528 modificado para incluir el campo pcMovNro
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
sSql = " exec stp_del_EliminaPersNegativa " & pnTipoDoc & ",'" & pcNumDoc & "','" & pcMovNro & "'" 'JACA 20110528 modificado para incluir el campo pcMovNro
    
Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set eliminaperslistanegativa = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing

End Function
'20090910 MADM ********************
Public Function eliminaperslistanegativa_Gral(ByVal pcond As Integer, ByVal pcapepat As String, ByVal pcapemat As String, ByVal pcNombre As String, ByVal pcMovNro As String) 'JACA 20110528 modificado para incluir el campo pcMovNro
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
sSql = " exec stp_del_EliminaPersNegativa_Gral " & pcond & ",'" & pcapepat & "','" & pcapemat & "','" & pcNombre & "','" & pcMovNro & "'" 'JACA 20110528 modificado para incluir el campo pcMovNro
    
Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set eliminaperslistanegativa_Gral = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing

End Function

'*** PEAC 20090715
Public Function VerificaPersListaNegativa(ByVal pnTipoPers As Integer, ByVal pnTipoDoc As Integer, _
        ByVal pcNroDoc As String, ByVal pcMovNro As String) As ADODB.Recordset 'JACA 20110530 SE AGREGO pcMovNro

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_VerificaPersListaNegativa " & pnTipoDoc & ",'" & pcNroDoc & "','" & pcMovNro & "'" 'JACA 20110530 SE AGREGO pcMovNro

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set VerificaPersListaNegativa = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
'*** MADM 100909
Public Function VerificaPersListaNegativa_Gral(ByVal pnTipoPers As Integer, ByVal pcapepat As String, ByVal pcapemat As String, _
        ByVal pcNombre As String, ByVal pcMovNro As String, _
        Optional ByVal psApeCasada As String = "") As ADODB.Recordset 'JACA 20110530 SE AGREGO pcMovNro
        'WIOR 20120710 AGREGO EL PARAMETRO psApeCasada
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

        sSql = "exec stp_sel_VerificaPersListaNegativa_Gral " & pnTipoPers & ",'" & pcapepat & "','" & pcapemat & "','" & pcNombre & "','" & pcMovNro & "','" & psApeCasada & "'" 'JACA 20110530 SE AGREGO pcMovNro

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set VerificaPersListaNegativa_Gral = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
'*****************
'*** PEAC 20090715
'madm 20100525
'madm 20101217 - estado
Public Function RegistraPersNegativo(ByVal pnTipoPers As Integer, ByVal pnTipoDoc As Integer, _
        ByVal pcNroDoc As String, ByVal pcNombre As String, ByVal pcapepat As String, ByVal pcapemat As String, _
        ByVal pcJustificacion As String, ByVal pcFuente As String, ByVal cInstitucion As String, ByVal cCargo As String, ByVal pcMovNro As String, Optional ByVal pnCondicion As Integer = 1, Optional ByVal pnEstado As Integer = 0, Optional ByVal pcApeCasada As String = "")
        'WIOR 20120710 pcApeCasada  se sgrego el parametro
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
sSql = " exec stp_ins_RegistraPersNegativo " & pnTipoPers & "," & pnTipoDoc & ",'" & pcNroDoc & "','" & pcNombre & "','" & pcapepat & "','" & pcapemat & "','" & pcJustificacion & "','" & pcFuente & "','" & pcMovNro & "','" & cInstitucion & "','" & cCargo & "'," & pnCondicion & "," & pnEstado & ",'" & pcApeCasada & "'"
'WIOR 20120710 pcApeCasada  se sgrego el parametro

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set RegistraPersNegativo = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing

'oCon.Ejecutar sSql

End Function
'end madm

'*** 20090928 MADM -- 20100510
'MADM 20101217 - CONDICION
Public Function ModificaPersNegativo(ByVal pnTipoPers As Integer, ByVal pnTipoDoc As Integer, _
        ByVal pcNroDoc As String, ByVal pcNombre As String, ByVal pcapepat As String, ByVal pcapemat As String, _
        ByVal pcJustificacion As String, ByVal pcFuente As String, ByVal cInstitucion As String, ByVal cCargo As String, ByVal pcMovNro As String, Optional ByVal pnCondicion As Integer = 1, Optional ByVal pnEstado As Integer = 0, Optional ByVal pcMovNroA As String = "", _
        Optional ByVal psApeCasada As String = "")
        'WIOR 20120710 AGREGO EL PARAMETRO psApeCasada
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

'If pcNroDoc = "" Then
'sSql = " exec stp_upd_ModificaPersNegativo_Gral " & pnTipoPers & "," & pnTipoDoc & ",'" & pcNroDoc & "','" & pcNombre & "','" & pcapepat & "','" & pcapemat & "','" & pcJustificacion & "','" & pcFuente & "','" & pcMovNro & "','" & cInstitucion & "','" & cCargo & "'," & pnCondicion & ", " & pnEstado & ",'" & pcMovNroA & "'"
'Else
sSql = " exec stp_upd_ModificaPersNegativo " & pnTipoPers & "," & pnTipoDoc & ",'" & pcNroDoc & "','" & pcNombre & "','" & pcapepat & "','" & pcapemat & "','" & pcJustificacion & "','" & pcFuente & "','" & pcMovNro & "','" & cInstitucion & "','" & cCargo & "'," & pnCondicion & ", " & pnEstado & ",'" & pcMovNroA & "','" & psApeCasada & "'"
'End If

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set ModificaPersNegativo = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing

End Function

'JACA 20110217
Public Function ModificaPersNegativoJustifica(ByVal pcMovNro As String, ByVal pcDelito As String, ByVal pcJuzgado As String, ByVal pcOfMultiple As String, ByVal pcOfJuzgado As String, ByVal pcExpediente As String, ByVal pcDepartamento As String, ByVal pnTipo As Integer, ByVal cComentario As String)
        
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
'sSql = " exec stp_ins_DatosPersListaNegativasJustificacion_MovNro '" & pcMovNro & "','" & pcDelito & "','" & pcJuzgado & "','" & pcOfMultiple & "','" & pcOfJuzgado & "','" & pcExpediente & "'"
'WIOR 20120329
sSql = " exec stp_ins_DatosPersListaNegativasJustificacion_MovNro '" & pcMovNro & "','" & pcDelito & "','" & pcJuzgado & "','" & pcOfMultiple & "','" & pcOfJuzgado & "','" & pcExpediente & "','" & pcDepartamento & "'," & pnTipo & ",'" & cComentario & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set ModificaPersNegativoJustifica = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function EliminaPersNegativoJustifica(ByVal pcMovNro As String)
        
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = " exec stp_del_DatosPersListaNegativasJustificacion_MovNro '" & pcMovNro & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set EliminaPersNegativoJustifica = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function
'JACA END

'MADM 20100510 CARGO & INSTITUCION
'MADM 20101217 - ESTADO
Public Function ModificaPersNegativo_Gral(ByVal pnTipoPers As Integer, ByVal pnTipoDoc As Integer, _
        ByVal pcNroDoc As String, ByVal pcNombre As String, ByVal pcapepat As String, ByVal pcapemat As String, _
        ByVal pcJustificacion As String, ByVal pcFuente As String, ByVal cInstitucion As String, ByVal cCargo As String, _
        ByVal pcMovNro As String, Optional ByVal pnCondicion As Integer = 1, Optional ByVal pnEstado As Integer = 0, Optional ByVal pcMovNroA As String = "")
        
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = " exec stp_upd_ModificaPersNegativo_Gral " & pnTipoPers & "," & pnTipoDoc & ",'" & pcNroDoc & "','" & pcNombre & "','" & pcapepat & "','" & pcapemat & "','" & pcJustificacion & "','" & pcFuente & "','" & pcMovNro & "','" & cInstitucion & "','" & cCargo & "'," & pnCondicion & ",'" & pnEstado & "','" & pcMovNroA & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set ModificaPersNegativo_Gral = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing

End Function
'20090929 MADM **************

'MADM 20100518 ---------------------------------------------------------------------------------------------------
Public Function VerificarNivelGerencialPEPS(ByVal pCodUser As String) As Boolean
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset
    On Error GoTo ErrorVerificarNivelGerencialPEPS
    
'    sSql = "select top 1 RH.cPersCod,P.CPersNombre,cRHCod,cUser,RHCT.cRHCargoCod,RHCT.cRHCargoDescripcion from RRHH RH " _
'         & "inner join Persona P ON RH.cPersCod = P.cpersCod AND nRHEstado = 201 " _
'         & "inner join RHCargos RHC on RH.cPersCod = RHC.cPersCod " _
'         & "inner join RHCargostabla RHCT on RHCT.cRHCargoCod = RHC.cRHCargoCod " _
'         & "where cUser = '" & Trim(pCodUser) & "' And RHCT.cRHCargoCod like '00100%' order by 5 "
        
sSql = "exec stp_val_VerificarNivelGerencialAprobar '" & pCodUser & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        VerificarNivelGerencialPEPS = True
    Else
        VerificarNivelGerencialPEPS = False
    End If
    R.Close
    Set R = Nothing
    Set oConecta = Nothing
    Exit Function
ErrorVerificarNivelGerencialPEPS:
    Err.Raise Err.Number, "VerificarNivelGerencial", Err.Description
    VerificarNivelGerencialPEPS = False
    
End Function

Public Function DevuelvePersListaNegativaPorAprobar(ByVal pCodAge As String, Optional ByVal pbPreAut As Boolean = False) As ADODB.Recordset 'WIOR 20121123 AGREGO pbPreAut

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_lis_PersonaNegativaAutorizacionPorAprobar '" & pCodAge & "'," & IIf(pbPreAut, 0, 1) 'WIOR 20121123 AGREGO pnPreAut

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevuelvePersListaNegativaPorAprobar = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'MADM 20100723
Public Function DevuelveRFPorAprobarSession(ByVal pCodAge As String, ByVal dfecha As Date) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_RFPorAprobar '" & pCodAge & "', '" & Format(dfecha, "yyyymmdd") & "' "

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevuelveRFPorAprobarSession = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Sub ModificaRFPorAprobarSession(ByVal pCodigo As Long)
        
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

    On Error GoTo ErrorModificaRFPorAprobarSession

    sSql = " exec stp_upd_SessionRF " & pCodigo & " "

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub
ErrorModificaRFPorAprobarSession:
    Err.Raise Err.Number, "Cargar ", Err.Description
End Sub
'END MADM

Public Function ValidaInsercionAprobacion(ByVal pfecsol As Date, ByVal pNombre As String, ByVal pUser As String, Optional pbolCond As Boolean = False, Optional ByVal pbPreAut As Boolean) As Boolean 'WIOR 20121123 AGREGO pbPreAut
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset
    
'    sSql = "select id,dfechaSol,dfechaApr,cNomPers, cUser,cOcupa,cAgencia nEstado from PersNegativaAutorizacion " _
'         & "where dfechasol = '" & Format(pfecsol, "YYYYMMDD") & "' And cNomPers ='" & pNombre & "' and  cUser ='" & pUser & "' "
    If pbolCond = False Then
        sSql = " exec stp_val_ValidaAutorizacionSensibles '" & Format(pfecsol, "yyyymmdd") & "','" & pNombre & "','" & pUser & "', 0 ," & IIf(pbPreAut, 1, 0) 'WIOR 20121123 AGREGO pbPreAut
    Else
        sSql = " exec stp_val_ValidaAutorizacionSensibles '" & Format(pfecsol, "yyyymmdd") & "','" & pNombre & "','" & pUser & "', 1 ," & IIf(pbPreAut, 1, 0) 'WIOR 20121123 AGREGO pbPreAut
    End If
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        ValidaInsercionAprobacion = True
    Else
        ValidaInsercionAprobacion = False
    End If
    R.Close
    Set R = Nothing
    Set oConecta = Nothing
    Exit Function
End Function
'madm 20110221 - agrego aprueba
Public Sub InsertaPersNegativo_Aprobacion(ByVal pfecsol As Date, ByVal pNombre As String, ByVal pUser As String, ByVal pCodAge As String, ByVal pOcupa As String, Optional pfecapro As Date = "1900/01/01", Optional ByVal pnEstado As Integer = 0, Optional ByVal pcDescripcion As String = "", Optional ByVal pcUserAutoriza As String = "", Optional ByVal pnPreAutorizacion As Integer = 0) 'WIOR 20121123 AGREGO pnPreAutorizacion
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
    
    On Error GoTo ErrorInsertaPersNegativo_Aprobacion
    
    sSql = " exec stp_ins_PersonaNegativaAutorizacion '" & Format(pfecsol, "YYYYMMDD") & "','" & Format(pfecapro, "YYYYMMDD") & "','" & pNombre & "','" & pUser & "','" & pOcupa & "','" & pCodAge & "'," & pnEstado & ", '" & pcDescripcion & "','" & pcUserAutoriza & "'," & pnPreAutorizacion 'WIOR 20121123 AGREGO pnPreAutorizacion

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub
ErrorInsertaPersNegativo_Aprobacion:
    Set oCon = Nothing
    Err.Raise Err.Number, "Cargar ", Err.Description
End Sub
'madm 20110221 - agrego aprueba
Public Sub ModificaPersNegativo_Aprobacion(ByVal pCodigo As Integer, ByVal pfecsol As Date, ByVal pfecapro As Date, ByVal pNombre As String, ByVal pUser As String, ByVal nEstado As Integer, Optional ByVal pcUserAutoriza As String = "", Optional ByVal pbPreAut As Boolean = False) 'WIOR 20121123 AGREGO pnPreAut
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

    On Error GoTo ErrorModificaPersNegativo_Aprobacion

    sSql = " exec stp_upd_PersonaNegativaAutorizacion " & pCodigo & ",'" & Format(pfecsol, "YYYYMMDD") & "','" & Format(pfecapro, "YYYYMMDD") & "','" & pNombre & "','" & pUser & "'," & nEstado & ", '" & pcUserAutoriza & "'," & IIf(pbPreAut, 1, 0) 'WIOR 20121123 AGREGO pbPreAut

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub
ErrorModificaPersNegativo_Aprobacion:
    Err.Raise Err.Number, "Cargar ", Err.Description
End Sub
'END MADM --------------------------------------------------------------------------------------------------------


''*** PEAC 20090716
'Public Function CargaDatosPersListaNegativas(ByVal pcAgencia As String, ByVal pnInmueble As Integer) As ADODB.Recordset
'Dim sSql As String
'Dim oCon As COMConecta.DCOMConecta
'sSql = " exec stp_sel_CargaDatosPersListaNegativas "
'
'Set oCon = New COMConecta.DCOMConecta
'oCon.AbreConexion
'Set CargaDatosPersListaNegativas = oCon.CargaRecordSet(sSql)
'oCon.CierraConexion
'Set oCon = Nothing
'
'End Function

'*** PEAC 20090716
Public Function CargaDatosPersNegativas(ByVal pnTipoDoc As Integer, ByVal pcNumDoc As String, Optional ByVal pcMovNro As String = "") As ADODB.Recordset 'JACA 20110530 SE AGREGO pcMovNro
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = " exec stp_sel_CargaDatosPersListaNegativas " & pnTipoDoc & ",'" & pcNumDoc & "','" & pcMovNro & "'" 'JACA 20110530 SE AGREGO pcMovNro

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set CargaDatosPersNegativas = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing

End Function
'20090928 MADM *********************
Public Function CargaDatosPersNegativas_MovNro(ByVal pmovnro As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = " exec stp_sel_CargaDatosPersListaNegativas_MovNro '" & pmovnro & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set CargaDatosPersNegativas_MovNro = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing

End Function

'MADM 20110216
Public Function CargaDatosPersNegativasJustificacion_MovNro(ByVal pmovnro As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = " stp_sel_CargaDatosPersListaNegativasJustificacion_MovNro '" & pmovnro & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set CargaDatosPersNegativasJustificacion_MovNro = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function
'**MADM ******************


'*** PEAC 20090718
Public Function ObtenerPersListaNegativa() As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = " exec stp_sel_RecuperaPersListaNegativa "

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtenerPersListaNegativa = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

'*** PEAC 20090720
'MARG 20160719*************
Public Function BusquedaEnListaNegativa(ByVal pnTipoDoc As Integer, ByVal psNumDoc As String) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorBusquedaEnListaNegativa
    
    sSql = "exec stp_sel_VerificaPersListaNegativa_nuevo " & pnTipoDoc & ",'" & psNumDoc & "',''" 'JACA 20110601 SE AGRE EL 3ER PARAMETRO=''
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set BusquedaEnListaNegativa = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorBusquedaEnListaNegativa:
    Set BusquedaEnListaNegativa = Nothing
    Err.Raise Err.Number, "Cargar Busqueda en Lista Negra", Err.Description
End Function

'MADM 20100723 -
Public Sub ModificaOcupaPers_REU(ByVal pUser As String, ByVal pOcupa As String)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
       
    sSql = " exec stp_upd_OcupacionPersonaREU '" & pUser & "','" & pOcupa & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub
End Sub
'MADM 20101011 -
Public Function BusquedaNumDocPersona(ByVal pnTipoDoc As Integer, ByVal psNumDoc As String) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorBusquedaEnListaNegativa
    
    sSql = "exec stp_sel_VerificaNumDocPersona " & pnTipoDoc & ",'" & psNumDoc & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set BusquedaNumDocPersona = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorBusquedaEnListaNegativa:
    Set BusquedaNumDocPersona = Nothing
    Err.Raise Err.Number, "Cargar Busqueda en Persona", Err.Description
End Function

'************************************************************************************
'MADM 20101206 - PAGO SERVICIO
Public Function CargaDatosInstitucionesPagoServicios(pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = " exec stp_sel_CargaDatosInstitucionesPagoMes '" & Format(pdFecha, "YYYYMMDD") & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set CargaDatosInstitucionesPagoServicios = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function

'MADM 20100317 - PAGO SERVICIO
Public Function CargaDatosPagoServiciosxCodServicio(ByVal pPersCod As String, ByVal pCodServicio As String, ByVal pFecSis As Date) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = " exec stp_sel_CargaDatosPersPagoServiciosCodServicio '" & pPersCod & "','" & pCodServicio & "','" & Format(pFecSis, "YYYYMMDD") & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set CargaDatosPagoServiciosxCodServicio = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function ValDatosPagoServiciosxCodServicioEsPagado(ByVal pPersCod As String, ByVal pCodServicio As String, ByVal pFecSis As Date) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

sSql = " exec stp_sel_CargaDatosPersPagoServiciosCodServicioPagado '" & pPersCod & "','" & pCodServicio & "','" & Format(pFecSis, "YYYYMMDD") & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set ValDatosPagoServiciosxCodServicioEsPagado = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function

'MADM 20111122
Public Function CargaDatosPagoServiciosxDocLista(ByVal pnNumDeuSist As String, ByVal pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim nCreditos As String
nCreditos = ""

If pnNumDeuSist <> "" Then
        If Len(pnNumDeuSist) > 1 Then
            nCreditos = Replace(Replace(pnNumDeuSist, "'", ""), " ", "")
        Else
            nCreditos = Replace(Replace(pnNumDeuSist, "'", ""), " ", "")
        End If
End If


sSql = " exec stp_sel_CargaDatosPersPagoServiciosCancelaLista '" & nCreditos & "','" & Format(pdFecha, "yyyymmdd") & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set CargaDatosPagoServiciosxDocLista = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function

'MADM 20110721 - 20110606
Public Function CargaDatosPagoServiciosxDoc(ByVal pnNumDeuSist As Long, ByVal pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = " exec stp_sel_CargaDatosPersPagoServiciosCancela " & pnNumDeuSist & ", '" & Format(pdFecha, "yyyymmdd") & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set CargaDatosPagoServiciosxDoc = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function CargaDatosPagoServiciosxInstitucion(ByVal pcPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = " exec stp_sel_CargaDatosPersPagoServiciosInstitucion '" & pcPersCod & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set CargaDatosPagoServiciosxInstitucion = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function
'MADM 20101202
Public Function CargaDatosReportePagoServiciosxDoc(ByVal pPersInst As String, ByVal pfecha As Date, ByVal pind As Integer) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = " exec stp_sel_LogPagoDeuda '" & pPersInst & "','" & Format(pfecha, "YYYYMMDD") & "'," & pind & ""

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set CargaDatosReportePagoServiciosxDoc = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function
'MADM 20111227
Public Function CargaDatosReportePagoServiciosxDocComision(ByVal pPersInst As String, ByVal pfecha As Date, ByVal pind As Integer) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = " exec stp_sel_LogPagoComisionDeuda '" & pPersInst & "','" & Format(pfecha, "YYYYMMDD") & "'," & pind & ""

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set CargaDatosReportePagoServiciosxDocComision = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function
'MADM 20110602
Public Function BuscaClienteServiciosNombre(ByVal sCadBusca As String, Optional ByVal sCodInst As String = "", Optional ByVal nTipoBusq As TBusqueda = BusquedaNombre) As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim Conn As COMConecta.DCOMConecta
    bError = False
    
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set BuscaClienteServiciosNombre = Nothing
        Set Conn = Nothing
        Set cmd = Nothing
        Set prm = Nothing
        Exit Function
    End If
    sCadBusca = Replace(sCadBusca, "'", "''")
    Conn.ConexionActiva.CommandTimeout = 7200
    Select Case nTipoBusq
        Case 1 'Busqueda por nombre
            Set BuscaClienteServiciosNombre = Conn.CargaRecordSet("spBuscaClienteApellidoServicioNom '" & Trim(sCadBusca) & "','" & Trim(sCodInst) & "'")
        Case 3 'Busqueda por Codigo
            Set BuscaClienteServiciosNombre = Conn.CargaRecordSet("spBuscaClienteDocServicioNom '" & Trim(sCadBusca) & "','" & Trim(sCodInst) & "'")
        Case 4 'Busqueda por nombre
            Set BuscaClienteServiciosNombre = Conn.CargaRecordSet("spBuscaClienteCodCliServicio '" & Trim(sCadBusca) & "','" & Trim(sCodInst) & "'")
    End Select
    Conn.CierraConexion
    Set Conn = Nothing
End Function

Public Function BuscaClienteServicios(ByVal sCadBusca As String, Optional ByVal sCodInst As String = "", Optional ByVal nTipoBusq As TBusqueda = BusquedaNombre) As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim Conn As COMConecta.DCOMConecta
    bError = False
    
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set BuscaClienteServicios = Nothing
        Set Conn = Nothing
        Set cmd = Nothing
        Set prm = Nothing
        Exit Function
    End If
    sCadBusca = Replace(sCadBusca, "'", "''")
    Conn.ConexionActiva.CommandTimeout = 7200
    Select Case nTipoBusq
        Case 1 'Busqueda por nombre
            Set BuscaClienteServicios = Conn.CargaRecordSet("spBuscaClienteApellidoServicio '" & Trim(sCadBusca) & "','" & Trim(sCodInst) & "'")
        Case 3 'Busqueda por Codigo
            Set BuscaClienteServicios = Conn.CargaRecordSet("spBuscaClienteDocServicio '" & Trim(sCadBusca) & "','" & Trim(sCodInst) & "'")
        Case 4 'Busqueda por nombre
            Set BuscaClienteServicios = Conn.CargaRecordSet("spBuscaClienteCliDocServicio '" & Trim(sCadBusca) & "','" & Trim(sCodInst) & "'")
    End Select
    Conn.CierraConexion
    Set Conn = Nothing
End Function

'JACA 20110427************************************************************************
Public Function CargarMagnitudPersona(ByRef prsMagnitudPersona As ADODB.Recordset, ByVal pnPersoneria As Integer)

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oConstante As COMDConstantes.DCOMConstantes
     
    On Error GoTo ErrorCargarMagnitudPersona
        
    
    Set oConstante = New COMDConstantes.DCOMConstantes
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set prsMagnitudPersona = oConstante.RecuperaMagnitudPersona(gPersJurMagnitud, pnPersoneria)
    oConecta.CierraConexion
    
    Set oConstante = Nothing
    Set oConecta = Nothing
    Exit Function
        
ErrorCargarMagnitudPersona:
        Set CargarMagnitudPersona = Nothing
        Err.Raise Err.Number, "Cargar Magnitud Persona", Err.Description
End Function
Public Function CargarSectoresCIUU() As Recordset

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    sSql = " exec stp_sel_Sectores_CIIU "
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set CargarSectoresCIUU = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function CargarParamAgeOcupacion(ByVal psAgeCod As String) As Recordset

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    sSql = " exec stp_sel_ParamAgeOcupacion_Lista '" & psAgeCod & "' "
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set CargarParamAgeOcupacion = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Sub insertarParamPersAgeOcupacion(ByVal psAgeCod As String, ByVal pnDesde As Integer, ByVal pnHasta As Integer, ByVal pnMonto As Currency, ByVal psUltimaActualizacion As String, ByVal pnActivo As Integer)

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    sSql = " exec stp_ins_ParamPersAgeOcupacion '" & psAgeCod & "'," & pnDesde & "," & pnHasta & "," & pnMonto & ",'" & psUltimaActualizacion & "'," & pnActivo & " "
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Sub modificarParamPersAgeOcupacion(ByVal psAgeCod As String)

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    sSql = " exec stp_upd_ParamPersAgeOcupacion '" & psAgeCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Function ObtenerPersAcumuladoMontoOpe(ByVal pnTpoCamb As Currency, ByVal pcPeriodo As String, ByVal pcPersCod As String) As Currency

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As Recordset
    Set rs = New Recordset
    sSql = " exec stp_sel_ObtenerPersAcumuladoMontoOpe " & pnTpoCamb & ",'" & pcPeriodo & "','" & pcPersCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
        ObtenerPersAcumuladoMontoOpe = rs!Monto
    Else
        ObtenerPersAcumuladoMontoOpe = 0
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function ObtenerParamPersAgeOcupacionMonto(ByVal psAgeCod As String, ByVal pnOcupacion As Integer) As Currency

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As Recordset
    Set rs = New Recordset
    sSql = " exec stp_sel_ParamPersAgeOcup_Monto '" & psAgeCod & "'," & pnOcupacion & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
        ObtenerParamPersAgeOcupacionMonto = rs!nMonto
    Else
        ObtenerParamPersAgeOcupacionMonto = 99999999
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'MARG 20160719**************************************
Public Function ObtenerDatosPersona(ByVal psPersCod As String) As Recordset

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    sSql = " exec stp_sel_ObtieneDatosPersona_nuevo '" & psPersCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set ObtenerDatosPersona = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

'ANDE 20170515 DINERO ELECTRÓNICO
Public Function ObtenerDatosPersonaXDNI(ByVal psDNI As String)
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_ObtenerDatosPersonaXDNI '" & psDNI & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set ObtenerDatosPersonaXDNI = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'FIN ANDE

Public Sub insertarPersonaAgeOcupacionDatos(ByVal pnMovNro As Long, ByVal pcPersCod As String, ByVal pnMonto As Currency, ByVal pnMontoAcumulado As Currency, ByVal pdFechaOpe As Date, ByVal psUltimaActualizacion As String)

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    sSql = " exec stp_ins_PersOpeAgeOcuDatos " & pnMovNro & ",'" & pcPersCod & "'," & pnMonto & "," & pnMontoAcumulado & ",'" & Format(pdFechaOpe, "yyyy/mm/dd") & "','" & psUltimaActualizacion & "' "
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Function ObtenerPersonaAgeOcupDatos_Verificar(ByVal pcPersCod As String, ByVal pdFechaOpe As Date) As Boolean

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As Recordset
    Set rs = New Recordset
    sSql = " exec stp_sel_PersOpeAgeOcuDatos_Verificar '" & pcPersCod & "','" & Format(pdFechaOpe, "yyyy/mm/dd") & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
        ObtenerPersonaAgeOcupDatos_Verificar = True
    Else
        ObtenerPersonaAgeOcupDatos_Verificar = False
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub ActualizarPersonaOcupIngreProm(ByVal psPersCod As String, ByVal psOcupacion As String, ByVal pnIngreProm As Double)

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    sSql = " exec stp_upd_PersonaOcupIngreProm '" & psPersCod & "','" & psOcupacion & "'," & pnIngreProm & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'JACA END

'MADM 20110914
Public Function DevuelvePersBloqueaRecuperaTodos() As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_BloqueaRecuperaTodos "

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevuelvePersBloqueaRecuperaTodos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function DevuelvePersBloqueaRecupera(ByVal pcPersCod As String, pcCtaCod As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_BloqueaRecupera '" & pcPersCod & "','" & pcCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevuelvePersBloqueaRecupera = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function DevuelvePersBloqueaRecuperaCred(ByVal pcPersCod As String, pcCtaCod As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_BloqueaRecuperaCred '" & pcPersCod & "','" & pcCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DevuelvePersBloqueaRecuperaCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function DevuelvePersBloqueaRecuperaBool(ByVal pcPersCod As String, ByVal pcCtaCod As String) As Boolean
    Dim sSql As String
    Dim rs As Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Set rs = New Recordset
    
    sSql = "exec stp_sel_BloqueaRecupera '" & pcPersCod & "','" & pcCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
        DevuelvePersBloqueaRecuperaBool = True
    Else
        DevuelvePersBloqueaRecuperaBool = False
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Sub ActualizarPersBloqueaRecupera(ByVal psPersCod As String, ByVal pcCtaCod As String, ByVal pdRegistro As Date, ByVal pdVigente As Boolean)
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    sSql = " exec stp_upd_BloqueaRecupera '" & psPersCod & "','" & pcCtaCod & "','" & Format(pdRegistro, "yyyy/mm/dd") & "'," & IIf(pdVigente, 1, 0) & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub InsertaPersBloqueaRecupera(ByVal psPersCod As String, ByVal psPersNombre As String, ByVal pcCtaCod As String, ByVal pdRegistro As Date, ByVal pdVigente As Boolean)
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    sSql = " exec stp_ins_BloqueaRecupera '" & psPersCod & "','" & psPersNombre & "','" & pcCtaCod & "','" & Format(pdRegistro, "yyyy/mm/dd") & "'," & IIf(pdVigente, 1, 0) & ""
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub EliminaPersBloqueaRecupera(ByVal psPersCod As String, ByVal psCtaCod As String)
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    sSql = " exec stp_del_BloqueaRecupera '" & psPersCod & "','" & psCtaCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Function ValidaPersBloqueaCredRecupera(ByVal psPersCod As String, ByVal pcCtaCod As String) As Boolean
    Dim rs As Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set rs = New Recordset
    
    sSql = "exec stp_sel_ValidaPersCreditoTitular '" & psPersCod & "','" & pcCtaCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
     If Not rs.EOF And Not rs.BOF Then
        ValidaPersBloqueaCredRecupera = True
    Else
        ValidaPersBloqueaCredRecupera = False
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'END MADM
'EJVG20110107 ************************
Public Function validaPermisoEditarPersona(ByVal psCargoCod As String, Optional psPersCod As String = "") As Boolean
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_ObtienePermisoEditarPersona '" & psCargoCod & "', '" & psPersCod & "'" 'RIRO20141106 ERS159
    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
        validaPermisoEditarPersona = rs!bPermiso
    Else
        validaPermisoEditarPersona = False
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'EJVG20120815 ***
Public Function TienePermisoEditarSujetoObligadoDJ(ByVal psCodCargo As String) As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim lsListaCargos As String
    Dim MatCargos() As String
    Dim sSql As String
    Dim iDJ As Integer
    
    Set rs = New ADODB.Recordset
    sSql = "Exec stp_sel_ObtenerListaCargosEditarSujetoObligadoDJ"
    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
        lsListaCargos = rs!nConsSisValor
    Else
        lsListaCargos = ""
    End If
    oConecta.CierraConexion
    
    TienePermisoEditarSujetoObligadoDJ = False
    MatCargos = Split(lsListaCargos, ",")
    For iDJ = 0 To UBound(MatCargos)
        If MatCargos(iDJ) = psCodCargo Then
            TienePermisoEditarSujetoObligadoDJ = True
            Exit For
        End If
    Next
    Set oConecta = Nothing
End Function
Public Function RecuperaDatosPersona_Basic(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    sSql = "Exec stp_sel_RecuperaDatosPersona_Basic '" & psPersCod & "'"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set RecuperaDatosPersona_Basic = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'END EJVG *******

'Create by GITU 19-09-2012
Public Function RecuperaDatosPersonaApeLote(ByVal psNroDoc As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    sSql = "Exec stp_sel_RecuperaDatosPersonaApeLote '" & psNroDoc & "'"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set RecuperaDatosPersonaApeLote = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaClientesConvTransf(ByVal psCodPers As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "Exec stp_sel_RecuperaClienteConvTransf '" & psCodPers & "'"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set RecuperaClientesConvTransf = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaOperaFrecuentes(ByVal pnEstado As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "Exec stp_sel_DevulveOpeFrecuentes '" & pnEstado & "'"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set RecuperaOperaFrecuentes = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaCuentasOpeFrec(ByVal pnCodOpeFre As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "Exec stp_sel_DevulveDetalleOpeFrecuentes '" & pnCodOpeFre & "'"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set RecuperaCuentasOpeFrec = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'End GITU
'WIOR 20121123 ***********************************************************************
Public Function VerificarNivelPreAutorizacionSupJA(ByVal pCodUser As String) As Boolean
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset
    On Error GoTo ErrorVerificarNivelPreAutorizacionSupJA
    

        
sSql = "exec stp_val_VerificarNivelPreAprobacion '" & pCodUser & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        VerificarNivelPreAutorizacionSupJA = True
    Else
        VerificarNivelPreAutorizacionSupJA = False
    End If
    R.Close
    Set R = Nothing
    Set oConecta = Nothing
    Exit Function
ErrorVerificarNivelPreAutorizacionSupJA:
    Err.Raise Err.Number, "VerificarNivelPreAutorizacionSupJA", Err.Description
    VerificarNivelPreAutorizacionSupJA = False
    
End Function
'WIOR FIN ****************************************************************************

'MIOL 20130517, SEGUN RQ13251 ********************************************************
Public Function RecuperaPersonasTipo_Arbol_AgenciaTodos(ByVal psTipo As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaPersonasTipo
    
    sSql = "exec stp_sel_DevuelvePersonaTipoAgenciaTodos '" & psTipo & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPersonasTipo_Arbol_AgenciaTodos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaPersonasTipo:
    Set RecuperaPersonasTipo_Arbol_AgenciaTodos = Nothing
    Err.Raise Err.Number, "Recupera Personas de un Tipo", Err.Description
End Function

Public Function RecuperaNroDOI_Inst(ByVal psCodInst As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaPersonasTipo
    
    sSql = "exec stp_sel_NroDOIxCodPerInst '" & psCodInst & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaNroDOI_Inst = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaPersonasTipo:
    Set RecuperaNroDOI_Inst = Nothing
    Err.Raise Err.Number, "Recupera DOI Institución", Err.Description
End Function
'END MIOL ****************************************************************************
'JUEZ 20130717 ***********************************************************************
Public Function VerificarPersonaPREDA(ByVal psCod As String, ByVal pnTipoVerif As Integer) As Boolean
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    VerificarPersonaPREDA = False
    sSql = "exec stp_sel_VerificarPersonaPREDA '" & psCod & "', " & pnTipoVerif
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        VerificarPersonaPREDA = True
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function RecuperaDatosPREDA() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta

    sSql = "exec stp_sel_RecuperaDatosPREDA"
    oCon.AbreConexion
    Set RecuperaDatosPREDA = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Sub InsertaPersonaPREDA(ByVal psPersIDNro As String, ByVal psPersNombre As String)
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    sSql = " exec stp_ins_InsertaPersonaPREDA '" & psPersIDNro & "','" & psPersNombre & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'INICIO ORCR-20140913*********
Public Sub ActualizarPersonaPREDAAutorizado(ByVal psPersIDNro As String, ByVal psAutorizado As String)
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    sSql = " exec stp_upd_ActualizarPersonaPREDAAutorizado '" & psPersIDNro & "'," & IIf(psAutorizado = ".", 1, 0)
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'FIN ORCR-20140913************
'END JUEZ ****************************************************************************
'JUEZ 20130724 ***********************************************************************
Public Function RecuperaInstitucionesCTSPersona(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta

    sSql = "exec stp_sel_RecuperaInstitucionesCTSPersona '" & psPersCod & "'"
    oCon.AbreConexion
    Set RecuperaInstitucionesCTSPersona = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'END JUEZ ****************************************************************************
'WIOR 20130808 *******************************************************************
Public Function ObtenerDatosBasicosPersNegativa(Optional ByVal psNumId As String = "", Optional ByVal psNombre As String = "", _
                                Optional ByVal psApePat As String = "", Optional ByVal psApeMat As String = "", _
                                Optional ByVal psApeCas As String = "", Optional pnTipo As Integer = 0) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    If pnTipo = 0 Then
        sSql = "exec stp_sel_PersListaNegativaDatosBasicos '" & psNumId & "','" & psNombre & "','" & psApePat & "','" & psApeMat & "','" & psApeCas & "'"
    Else
        sSql = "exec stp_sel_PersListaNegativaXNumDoc '" & psNumId & "'"
    End If
    oCon.AbreConexion
    Set ObtenerDatosBasicosPersNegativa = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function PersNegativaExiste(Optional ByVal psNumId As String = "", Optional ByVal psNombre As String = "", _
                                Optional ByVal psApePat As String = "", Optional ByVal psApeMat As String = "", _
                                Optional ByVal psApeCas As String = "", Optional pnTipo As Integer = 0, _
                                Optional pnTipoPers As Integer = 0) As String
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    If pnTipo = 1 Then
        If pnTipoPers = 1 Then
            sSql = "exec stp_sel_PersNegExiteXPersNat '" & psNombre & "','" & psApePat & "','" & psApeMat & "','" & psApeCas & "'"
        ElseIf pnTipoPers = 2 Then
            sSql = "exec stp_sel_PersNegExiteXPersJur '" & psNombre & "'"
        End If
    ElseIf pnTipo = 2 Then
        sSql = "exec stp_sel_PersNegExiteXNumDoc '" & psNumId & "'"
    End If
    
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    
    If Not (rs.EOF And rs.BOF) Then
        PersNegativaExiste = "SI"
    Else
        PersNegativaExiste = "NO"
    End If

    Set oCon = Nothing
    Set rs = Nothing
End Function
'WIOR 20130808 *******************************************************************
'WIOR 20130827 *******************************************************************
Public Sub RegistroDatosCambiosDatosPrinc(ByVal pbDocs As Boolean, ByVal psMovCambio As String, ByVal psPersCod As String, _
                Optional ByVal psNomAnt As String = "", Optional ByVal psNomAct As String = "", _
                Optional ByVal pnTipoDocAnt As Integer = 0, Optional ByVal pnTipoDocAct As Integer = 0, _
                Optional ByVal psNumDocAnt As String = "", Optional ByVal psNumDocAct As String = "")

Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta
    
If Not pbDocs Then
    sSql = "EXEC stp_ins_PersonaCambios '" & psPersCod & "','" & psNomAnt & "','" & psNomAct & "','" & psMovCambio & "'"
Else
    sSql = "EXEC stp_ins_PersonaCambiosDoc '" & psPersCod & "'," & pnTipoDocAnt & "," & pnTipoDocAct & ",'" & psNumDocAnt & "','" & psNumDocAct & "','" & psMovCambio & "'"
End If

oCon.AbreConexion
oCon.Ejecutar sSql
oCon.CierraConexion
Set oCon = Nothing
End Sub
Public Sub RegistroCabeceraDatosCambiosDatosPrinc(ByVal psMovCambio As String, ByVal psPersCod As String, ByVal pnAplicacion As Integer)
Dim sSql As String
Dim oCon As New COMConecta.DCOMConecta

sSql = "EXEC stp_ins_PersonaCambiosCabecera '" & psPersCod & "','" & psMovCambio & "'," & pnAplicacion

oCon.AbreConexion
oCon.Ejecutar sSql
oCon.CierraConexion
Set oCon = Nothing
End Sub
'WIOR FIN ************************************************************************

' *** RIRO SEGUN TI-ERS108-2013 ***

Public Function RecuperarDatosRF3(ByVal pAgencia As String, Optional sCargo As String = "007026") As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta

    sSql = "exec STP_SEL_USUARIOPORCARGO '" & pAgencia & "', '" & sCargo & "'"
    oCon.AbreConexion
    Set RecuperarDatosRF3 = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub RegistrarGruposRF3(ByVal sPersCod As String, ByVal bOperacionesSimultaneas As Boolean, Optional ByRef sGrupo As String = "", _
                              Optional ByVal sCargoRelacion As String = "")
    Dim sSql As String
    Dim i As Integer
    Dim oCon As New COMConecta.DCOMConecta

    oCon.AbreConexion
'    sSql = "exec STP_LimpiarRFIIIGrupos '" & sPersCod & "'"
'    oCon.CargaRecordSet (sSql)
'    For i = 1 To UBound(sGrupos)
'        sSql = "exec STP_INS_InsertarRFIIIGrupos '" & sPersCod & "', " & IIf(bOperacionesSimultaneas, 1, 0) & ", '" & sGrupos(i) & "'"
'        oCon.CargaRecordSet (sSql)
'    Next
    oCon.BeginTrans
    
    sSql = "exec STP_INS_InsertarRFIIIGrupos '" & sPersCod & "', " & IIf(bOperacionesSimultaneas, 1, 0) & ", '" & sGrupo & "'"
    oCon.CargaRecordSet (sSql)
    sSql = "exec STP_INS_RELACIONRFIII_CARGO '" & sPersCod & "', '" & sCargoRelacion & "'"
    oCon.CargaRecordSet (sSql)
    oCon.CommitTrans
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Function RecuperarGruposRF3(ByVal psUser As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta

    sSql = "exec STP_SEL_GRUPOSUSUARIO '" & psUser & "'"
    oCon.AbreConexion
    Set RecuperarGruposRF3 = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
' *** Fin RIRO ***
Public Function RecuperarCargoRF3(ByVal sPersCod As String) As String
    Dim sSql As String, sCargo As String
    Dim oCon As New COMConecta.DCOMConecta
    Dim rsCargo As New Recordset
    'cCargoRel
    sSql = "EXEC STP_SEL_RELACIONRFIII_CARGO '" & sPersCod & "'"
    oCon.AbreConexion
    Set rsCargo = oCon.CargaRecordSet(sSql)
    If Not rsCargo Is Nothing Then
        If Not rsCargo.EOF And Not rsCargo.BOF Then
            sCargo = rsCargo!cCargoRel
        End If
    End If
    oCon.CierraConexion
    RecuperarCargoRF3 = sCargo
End Function
'JUEZ 20131007 ****************************************************************
Public Function PermiteParticiparCampañaDatos(ByVal psPersCod As String) As Boolean
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset
    PermiteParticiparCampañaDatos = False
    
    sSql = "exec stp_sel_PermiteParticiparCampañaDatos '" & psPersCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        PermiteParticiparCampañaDatos = True
    Else
        PermiteParticiparCampañaDatos = False
    End If
    R.Close
    Set R = Nothing
    Set oConecta = Nothing
End Function
Public Function VerificaPersonaConCreditoActivo(ByVal psPersCod As String, Optional pbMalaCalif As Boolean = False) As Boolean
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset
    VerificaPersonaConCreditoActivo = False
    
    If Not pbMalaCalif Then
        sSql = "exec stp_sel_VerificaPersonaConCreditoActivo '" & psPersCod & "'"
    Else
        sSql = "exec stp_sel_VerificaPersonaConCreditoActivoMalaCalificacion '" & psPersCod & "'"
    End If
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        VerificaPersonaConCreditoActivo = True
    Else
        VerificaPersonaConCreditoActivo = False
    End If
    R.Close
    Set R = Nothing
    Set oConecta = Nothing
End Function
Public Sub GrabarDatosPersonaCampDatosBitacora(ByVal psPersCod As String, ByVal prsPersona As ADODB.Recordset, ByVal prsDocs As ADODB.Recordset, _
                                               ByVal psCodUserResp As String, ByVal psActualizacion As String)
Dim sSql As String
Dim i As Integer
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    oCon.BeginTrans
    
    sSql = " exec stp_ins_GrabarDatosPersonaCampDatosBitacora '" & psPersCod & "','" & prsPersona("cPaisReside") & "','" _
             & prsPersona("cPersDireccUbiGeo") & "','" & prsPersona("cPersDireccDomicilio") & "','" _
             & prsPersona("cPersRefDomicilio") & "','" & prsPersona("cPersNegocioUbiGeo") & "','" & prsPersona("cPersNegocioDireccion") & "','" _
             & prsPersona("cPersNegocioRef") & "','" & prsPersona("cActiGiro") & "','" & prsPersona("cPersEmail") & "','" _
             & prsPersona("cPersTelefono") & "','" & prsPersona("cPersTelefono2") & "','" _
             & prsPersona("cPersCelular") & "','" & prsPersona("cPersCelular2") & "','" & prsPersona("cPersCelular3") & "','" _
             & psActualizacion & "','" & psCodUserResp & "'," & prsPersona("nRemisionInfoEmail")
    oCon.ConexionActiva.Execute sSql
    
    prsDocs.MoveFirst
    For i = 1 To prsDocs.RecordCount
        sSql = " exec stp_ins_GrabarDatosPersonaDocsCampDatosBitacora '" & psPersCod & "','" & psActualizacion & "'," & i & "," _
                                                                         & prsDocs("cPersIDTpo") & ",'" & prsDocs("cPersIDNro") & "'"
        oCon.ConexionActiva.Execute sSql
        prsDocs.MoveNext
        If i = 2 Then Exit For
    Next i
    
    oCon.CommitTrans
    Set oCon = Nothing
End Sub
Public Function RecuperaDatosPersonaCampDatosBitacora(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta

    sSql = "exec stp_sel_RecuperaDatosPersonaCampDatosBitacora '" & psPersCod & "'"
    oCon.AbreConexion
    Set RecuperaDatosPersonaCampDatosBitacora = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'Public Function VerificaPersonaParticipaCampDatos(ByVal psCtaCod As String) As Boolean
'    Dim sSql As String
'    Dim rs As ADODB.Recordset
'    Dim oCon As New COMConecta.DCOMConecta
'
'    sSql = "exec stp_sel_VerificaPersonaParticipaCampDatos '" & psCtaCod & "'"
'    oCon.AbreConexion
'    Set rs = oCon.CargaRecordSet(sSql)
'    If rs.RecordCount > 0 Then
'        VerificaPersonaParticipaCampDatos = True
'    Else
'        VerificaPersonaParticipaCampDatos = False
'    End If
'    oCon.CierraConexion
'    Set oCon = Nothing
'End Function
Public Function ReporteArqueoCampActualizacionDatos(ByVal pdFecha As String, ByVal psAgeCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_ReporteArqueoCampActualizacionDatos '" & Format(pdFecha, "yyyymmdd") & "','" & psAgeCod & "'"
    oCon.AbreConexion
    Set ReporteArqueoCampActualizacionDatos = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'END JUEZ *********************************************************************
'RECO20140226 ERS174-2013*********************************************************
Public Function ObtieneDatosFICred(ByVal psNumFuente As String, ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_ObtieneDatosFICred '" & psNumFuente & "','" & psCtaCod & "'"
    oCon.AbreConexion
    Set ObtieneDatosFICred = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'RECO FIN*************************************************************************
'FRHU 20140212 RQ14005
Public Function RecuperarDatosPersonaDesembolsoTercero(ByVal pCtaAhorro As String, ByRef prsPersona As ADODB.Recordset)
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_PersonaDesembolsoTercero '" & pCtaAhorro & "'"
    oCon.AbreConexion
    Set prsPersona = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
End Function
'FRHU 20140212

'WIOR 20140318 *******************************************************************
Public Function RecuperaFuenteIngEstFinan(ByVal psNumFteIngreso As String, ByVal pdFecha As Date) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim sSql As String

    sSql = " Exec stp_sel_PersFIEstadosFinan '" & psNumFteIngreso & "','" & Format(pdFecha, "yyyymmdd") & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaFuenteIngEstFinan = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
End Function
Public Sub GrabarFteEstFinan(ByVal prsEstFinan As ADODB.Recordset, ByVal sNuevo As String, ByVal sNumFuente As String, ByVal oConecta As COMConecta.DCOMConecta)
Dim sSql As String
sSql = ""
 If prsEstFinan.RecordCount > 0 Then

        While Not prsEstFinan.EOF

            If sNuevo = "S" Then
                prsEstFinan("cNumFuente") = sNumFuente
            End If
            sSql = "EXEC stp_ins_upd_PersFIEstadosFinan '" & prsEstFinan("cNumFuente") & "','" & Format(prsEstFinan("dPersEval"), "yyyymmdd") & "'," & prsEstFinan("nTpoEF") & "," & prsEstFinan("nSubTpoEF") & "," & prsEstFinan("nMonto")
            oConecta.ConexionActiva.Execute sSql

             prsEstFinan.MoveNext
        Wend
    End If
End Sub
'WIOR FIN ************************************************************************
'RIRO20140407 ERS017 20143 *******************************************************
Public Function ValidaTitularSueldo(ByVal sCodEmpleador As String, ByVal nMoneda As Integer, ByVal sPersCod As String) As Integer
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rsCtaSueldo As New ADODB.Recordset
    Dim nRespuesta As Integer
    sSql = "Exec stp_sel_validarTitularLote '" & sCodEmpleador & "', " & nMoneda & ", '" & sPersCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rsCtaSueldo = oConecta.CargaRecordSet(sSql)
    nRespuesta = 0
    If Not rsCtaSueldo Is Nothing Then
        If Not rsCtaSueldo.BOF And Not rsCtaSueldo.EOF Then
            If rsCtaSueldo.RecordCount > 0 Then
                nRespuesta = rsCtaSueldo.RecordCount
            End If
        End If
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    ValidaTitularSueldo = nRespuesta
End Function
Public Function ValidaTramaDeposito(ByVal psNroDoc As String, ByVal sCodInstitucion As String, ByVal nMoneda As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    sSql = "Exec STP_SEL_VALIDATRAMADEPOSITO '" & psNroDoc & "', '" & sCodInstitucion & "', " & nMoneda
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ValidaTramaDeposito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function ListarTramaDeposito(ByVal psNroDoc As String, ByVal psMonto As String, ByVal sCodInstitucion As String, ByVal nMoneda As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    sSql = "Exec STP_SEL_LISTARDEPOSITOLOTE '" & psNroDoc & "','" & psMonto & "','" & sCodInstitucion & "', " & nMoneda
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ListarTramaDeposito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'END RIRO ***********************************************************************
'FRHU 20150311 ERS013-2015
Public Function RecuperaEstFinanDetalle(ByVal pnCodEstFinan As Integer) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim sSql As String

    sSql = "Exec stp_sel_PersEstadoFinancieroDet " & pnCodEstFinan

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaEstFinanDetalle = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
End Function
Public Function GrabarEstFinan(ByVal psPersCod As String, ByVal pdEstFinanFecha As Date, ByVal psEstFinanUser As String, ByVal nAuditado As Integer, _
                               ByVal pdFechaPatriIni As String, ByVal pdFechaPatriFin As String) As Long 'LUCV20170915, Agregó: pdFechaPatriIni, pdFechaPatriFin según ERS051-2017
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    Dim oRS As New ADODB.Recordset
    
    sSql = "EXEC stp_ins_EstadosFinan '" & psPersCod & "','" & Format(pdEstFinanFecha, "yyyymmdd") & "','" & psEstFinanUser & "'," & nAuditado & ", '" & pdFechaPatriIni & "', '" & pdFechaPatriFin & "'"
    oCon.AbreConexion
    Set oRS = oCon.CargaRecordSet(sSql)
    If Not oRS.BOF And Not oRS.EOF Then
        If oRS.RecordCount > 0 Then
            GrabarEstFinan = oRS!nCodEstFinan
        Else
            GrabarEstFinan = 0
        End If
    Else
        GrabarEstFinan = 0
    End If
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
End Function
Public Sub GrabarEstFinanDet(ByVal prsEstFinan As ADODB.Recordset, ByVal pnCodEF As Integer)
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    If prsEstFinan.RecordCount > 0 Then
        oCon.AbreConexion
        While Not prsEstFinan.EOF
            sSql = "EXEC stp_ins_upd_EstadosFinanDet " & prsEstFinan("nCodEstFinan") & "," & prsEstFinan("nEstFinanTipo") & "," & prsEstFinan("nEstFinanSubTipo") & "," & prsEstFinan("nMonto")
            oCon.Ejecutar sSql
            prsEstFinan.MoveNext
        Wend
        oCon.CierraConexion
    End If
End Sub
Public Function RecuperarDatosPersonaEstadoFinanciero(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_PersEstadoFinanciero '" & psPersCod & "'"
    oCon.AbreConexion
    Set RecuperarDatosPersonaEstadoFinanciero = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
End Function
Public Sub EliminarEstFinan(ByVal pnCodEstFinan As Long)
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta

    sSql = "EXEC stp_upd_EstadoEstadosFinan " & pnCodEstFinan

    oCon.AbreConexion
    oCon.Ejecutar sSql
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Function GetUltSemestreEstadoFinancieroPersona(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_UltSemestreEstadoFinancieroPersona '" & psPersCod & "'"
    oCon.AbreConexion
    Set GetUltSemestreEstadoFinancieroPersona = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
End Function
Public Function ExistePersonaEstadoFinanciero(ByVal psPersCod As String, ByVal pdEstFinanFecha As Date) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_ExistePersEstadoFinanciero '" & psPersCod & "','" & Format(pdEstFinanFecha, "yyyymmdd") & "'"
    oCon.AbreConexion
    Set ExistePersonaEstadoFinanciero = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
End Function
'FIN FRHU 20150311

'RIRO20150519 ERS162-2014 *********************************
Public Function ValidaTramaUtilidades(ByVal psNroDoc As String, ByVal psUser As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    sSql = "Exec STP_SEL_VALIDA_DOI_PERS '" & psNroDoc & "', '" & psUser & "'"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ValidaTramaUtilidades = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function ObtenerTramaUtilidades(ByVal pnAnio As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    sSql = "Exec STP_SEL_TRAMAS_UTILIDADES " & pnAnio
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ObtenerTramaUtilidades = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function ObtenerTramaUtilidadesDetalle(ByVal pnUtilidadesDetalle As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    sSql = "Exec STP_SEL_TRAMAS_DETALLE " & pnUtilidadesDetalle
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ObtenerTramaUtilidadesDetalle = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function GrabarTramaUtilidades(ByVal psNomApell As String, ByVal psNroDoc As String, _
                                      ByVal psPeriodo As String, ByVal psMoneda As String, _
                                      ByVal psImporte As String, ByVal psUser As String, _
                                      ByVal psCargo As String, ByVal psArea As String, _
                                      ByVal psAgeCod As String, _
                                      ByVal psFechaIngreso As String, ByVal sC09_ParticipAdistribuir As String, _
                                      ByVal sC10_DiasLaborTodosTrabAnio As String, ByVal sC11_RemunPercibTodosTrabAnio As String, _
                                      ByVal sC12_MontDistribXdiasLabor As String, ByVal sC13_MontDistribXremunPercib As String, _
                                      ByVal sC14_TotDiasEfectivLabor As String, ByVal sC15_TotRemuneraciones As String, _
                                      ByVal sC16_ParticipXdiasLabor As String, ByVal sC17_ParticipXremuneraciones As String, _
                                      ByVal sC18_TotParticipUtilidades As String, ByVal sC19_RetencionImpuestoRenta As String, _
                                      ByVal sC20_TotalDescuento As String, ByVal sC21_TotalPagar As String, _
                                      ByVal sRetencion As String, ByVal sOtrosDescuentos As String) As Boolean
   
    On Error GoTo error_handler
    
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "Exec STP_INS_TramaUtilidad '" & psPeriodo & "', '" & _
                                            psMoneda & "', '" & _
                                            psUser & "', '" & _
                                            psNomApell & "', '" & _
                                            psNroDoc & "', '" & _
                                            psImporte & "', '" & _
                                            psCargo & "', '" & _
                                            psArea & "', '" & _
                                            psAgeCod & "', '" & _
                                            psFechaIngreso & "', '" & _
                                            sC09_ParticipAdistribuir & "', '" & _
                                            sC10_DiasLaborTodosTrabAnio & "', '" & _
                                            sC11_RemunPercibTodosTrabAnio & "', '" & _
                                            sC12_MontDistribXdiasLabor & "', '" & _
                                            sC13_MontDistribXremunPercib & "', '" & _
                                            sC14_TotDiasEfectivLabor & "', '" & _
                                            sC15_TotRemuneraciones & "', '" & _
                                            sC16_ParticipXdiasLabor & "', '" & _
                                            sC17_ParticipXremuneraciones & "', '" & _
                                            sC18_TotParticipUtilidades & "', '" & _
                                            sC19_RetencionImpuestoRenta & "', '" & _
                                            sC20_TotalDescuento & "', '" & _
                                            sC21_TotalPagar & "', '" & sRetencion & "', '" & _
                                            sOtrosDescuentos & "' "
                                            
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    oConecta.BeginTrans
    oConecta.CargaRecordSet (sSql)
    oConecta.CommitTrans
    oConecta.CierraConexion
    Set oConecta = Nothing
    GrabarTramaUtilidades = True
    
    Exit Function
error_handler:
    oConecta.RollbackTrans
    GrabarTramaUtilidades = False
End Function
Public Function BajaTramaUtilidades(ByVal pnIdTrama As Long) As Boolean
    
    On Error GoTo error_handler
    
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "Exec STP_BAJA_TRAMAS_UTILIDADES " & pnIdTrama
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    oConecta.BeginTrans
    oConecta.CargaRecordSet (sSql)
    oConecta.CommitTrans
    oConecta.CierraConexion
    Set oConecta = Nothing
    BajaTramaUtilidades = True
    
    Exit Function
error_handler:
    BajaTramaUtilidades = False
End Function
Public Function ObtenerUtilidadesExTrabajador(ByVal psDOI As String, ByVal psAgeCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    sSql = "Exec STP_SEL_TRAMAS_X_TRABAJADORES '" & psDOI & "', '" & psAgeCod & "' "
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ObtenerUtilidadesExTrabajador = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'END RIRO ************************************************
'RECO20151203 ERS073-2015***********************************
Public Function ObtieneTitularCuenta(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta

    sSql = "Exec stp_sel_ObtieneTitularCuenta '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta

    oConecta.AbreConexion
    Set ObtieneTitularCuenta = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'RECO FIN **************************************************
'RECO20151222 ERS040-2015***********************************
Public Sub ActualizaUBIGEOPersona(ByVal psUbiGeo As String, ByVal psDireccion As String, ByVal psPersCod As String)
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta

    sSql = "EXEC stp_upd_ActualizaUBIGEOPersona '" & psUbiGeo & "','" & psDireccion & "','" & psPersCod & "'"

    oCon.AbreConexion
    oCon.Ejecutar sSql
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

'By GIPO 31-10-2016
Public Function seleccionarCodigoSBSRepetido(ByVal codigoSBS As String, ByVal codigoPers As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loConecta As COMConecta.DCOMConecta
    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    lsSQL = " Exec sp_sel_seleccionarCodigoSBSRepetido '" & codigoSBS & "','" & codigoPers & "'"
    Set seleccionarCodigoSBSRepetido = loConecta.CargaRecordSet(lsSQL)
    loConecta.CierraConexion
    Set loConecta = Nothing
 End Function
 
'RECO FIN***************************************************
'FRHU 20160515 ERS020-2016: RIC
'Public Function BuscaClienteConsolidada(ByVal sCadBusca As String, Optional ByVal nTipoBusq As TBusqueda = BusquedaNombre, _
'                                        Optional ByVal pnTipoReporte As Integer = 1, Optional ByVal psFechaCierre As String) As ADODB.Recordset
'Dim cmd As ADODB.Command
'Dim prm As ADODB.Parameter
'Dim Conn As COMConecta.DCOMConecta
'    bError = False
'
'    Set Conn = New COMConecta.DCOMConecta
'    If Not Conn.AbreConexion() Then
'        bError = True
'        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
'        Set BuscaCliente = Nothing
'        Set Conn = Nothing
'        Set cmd = Nothing
'        Set prm = Nothing
'        Exit Function
'    End If
'
'    sCadBusca = Replace(sCadBusca, "'", "''")
'    Conn.ConexionActiva.CommandTimeout = 7200
'    Select Case nTipoBusq
'        Case 1 'Busqueda por nombre
'            Set BuscaCliente = Conn.CargaRecordSet("spBuscaClienteApellido '" & Trim(sCadBusca) & "'")
'        Case 2 'Busqueda por Codigo
'            Set BuscaCliente = Conn.CargaRecordSet("spBuscaClienteCodigoConsolidada '" & Trim(sCadBusca) & "'," & pnTipoReporte & ",'" & psFechaCierre & "'")
'        Case 3 'Busqueda por Documento
'            Set BuscaCliente = Conn.CargaRecordSet("spBuscaClienteDoc '" & Trim(sCadBusca) & "'")
'        Case 4 'Busqueda Empleado por nombre
'            Set BuscaCliente = Conn.CargaRecordSet("spBuscaEmpleadoApellido '" & Trim(sCadBusca) & "'")
'        Case 5 'Busqueda Empleado por Codigo
'            Set BuscaCliente = Conn.CargaRecordSet("spBuscaEmpleadoCodigo '" & Trim(sCadBusca) & "'")
'        Case 6 'Busqueda Empleado por Documento
'            Set BuscaCliente = Conn.CargaRecordSet("spBuscaEmpleadoDoc '" & Trim(sCadBusca) & "'")
'    End Select
'    Conn.CierraConexion
'    Set Conn = Nothing
'End Function
'FIN FRHU

'**********marg ers046-2016*********************************************
Public Function ListarPersRPLAFTVistoContinuidad(ByVal pcAgeCod As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_ERS0462016_ListarPersRPLAFTVistoContinuidad '" & pcAgeCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ListarPersRPLAFTVistoContinuidad = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function VerificarAutorizacionVistoContinuidad(ByVal pcUser As String) As Boolean
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    Dim R As ADODB.Recordset
    On Error GoTo ErrorVerificarAutorizacionVistoContinuidad
           
    sSql = "exec stp_sel_ERS0462016_VerificarAutorizacionVistoContinuidad '" & pcUser & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        VerificarAutorizacionVistoContinuidad = True
    Else
        VerificarAutorizacionVistoContinuidad = False
    End If
    R.Close
    Set R = Nothing
    Set oConecta = Nothing
    Exit Function
ErrorVerificarAutorizacionVistoContinuidad:
    Err.Raise Err.Number, "VerificarAutorizacionVistoContinuidad", Err.Description
    VerificarAutorizacionVistoContinuidad = False
End Function

Public Sub ActualizarPersRPLAFTVistoContinuidad(ByVal pnIdVisto As Integer, ByVal pbAdmision As Boolean, ByVal pcComentario As String, pcUserVisto As String)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

    On Error GoTo ErrorActualizarPersRPLAFTVistoContinuidad

    sSql = " exec stp_upd_ERS0462016_ActualizarPersRPLAFTVistoContinuidad " & pnIdVisto & "," & IIf(pbAdmision = True, 1, 0) & ",'" & pcComentario & "','" & pcUserVisto & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub
ErrorActualizarPersRPLAFTVistoContinuidad:
    Err.Raise Err.Number, "ActualizarPersRPLAFTVistoContinuidad ", Err.Description
End Sub

Public Function InsertarPersRPLAFTVistoContinuidad(pcPersCodCliente As String, pnCondicion As Integer, pcOpeCod As String, pcUserSolicitud As String, pcAgeCod As String, nColocRelacPers As Integer) As Integer
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim oRS As New ADODB.Recordset

    On Error GoTo ErrorInsertarPersRPLAFTVistoContinuidad

    sSql = " exec stp_ins_ERS0462016_InsertarPersRPLAFTVistoContinuidad '" & pcPersCodCliente & "'," & pnCondicion & ",'" & pcOpeCod & "','" & pcUserSolicitud & "','" & pcAgeCod & "'," & nColocRelacPers
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set oRS = oCon.Ejecutar(sSql)
    If Not oRS.BOF And Not oRS.EOF Then
        If oRS.RecordCount > 0 Then
            InsertarPersRPLAFTVistoContinuidad = oRS!nId
        Else
            InsertarPersRPLAFTVistoContinuidad = 0
        End If
    Else
        InsertarPersRPLAFTVistoContinuidad = 0
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrorInsertarPersRPLAFTVistoContinuidad:
    Err.Raise Err.Number, "InsertarPersRPLAFTVistoContinuidad ", Err.Description
    InsertarPersRPLAFTVistoContinuidad = 0
End Function

Public Function SolicitarVistoContinuidad(pcPersCodCliente As String, pnCondicion As Integer, pcOpeCod As String, pcUserSolicitud As String, pcAgeCod As String) As Integer
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim oRS As New ADODB.Recordset

    On Error GoTo ErrorSolicitarVistoContinuidad

    sSql = " exec stp_sel_ERS0462016_SolicitarVistoContinuidad '" & pcPersCodCliente & "'," & pnCondicion & ",'" & pcOpeCod & "','" & pcUserSolicitud & "','" & pcAgeCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set oRS = oCon.Ejecutar(sSql)
    If Not oRS.BOF And Not oRS.EOF Then
        If oRS.RecordCount > 0 Then
            SolicitarVistoContinuidad = oRS!nRes
        Else
            SolicitarVistoContinuidad = 0
        End If
    Else
        SolicitarVistoContinuidad = 0
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrorSolicitarVistoContinuidad:
    Err.Raise Err.Number, "SolicitarVistoContinuidad ", Err.Description
    SolicitarVistoContinuidad = 0
End Function

Public Function ObtenerPersRPLAFTVistoContinuidad(pcPersCodCliente As String, pnCondicion As Integer, pcOpeCod As String, pcUser As String, pcAgeCod As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = " exec stp_sel_ERS0462016_ObtenerPersRPLAFTVistoContinuidad '" & pcPersCodCliente & "'," & pnCondicion & ",'" & pcOpeCod & "','" & pcUser & "','" & pcAgeCod & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerPersRPLAFTVistoContinuidad = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function GetOperacionSolicitudVistoContinuidad(psOpeCod As String) As String
    Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim oRS As New ADODB.Recordset

    On Error GoTo ErrorGetOperacionSolicitudVistoContinuidad

    sSql = " exec stp_sel_ERS0462016_GetOperacionSolicitudVistoContinuidad '" & psOpeCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set oRS = oCon.Ejecutar(sSql)
    If Not oRS.BOF And Not oRS.EOF Then
        If oRS.RecordCount > 0 Then
            GetOperacionSolicitudVistoContinuidad = oRS!cOpeDesc
        Else
            GetOperacionSolicitudVistoContinuidad = 0
        End If
    Else
        GetOperacionSolicitudVistoContinuidad = 0
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrorGetOperacionSolicitudVistoContinuidad:
    Err.Raise Err.Number, "GetOperacionSolicitudVistoContinuidad ", Err.Description
    GetOperacionSolicitudVistoContinuidad = 0
End Function

Public Function GetCorreoDestinoSolicitudVistoContinuidad() As String
    Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim oRS As New ADODB.Recordset

    On Error GoTo ErrorGetCorreoDestinoSolicitudVistoContinuidad

    sSql = " exec stp_sel_ERS0462016_GetCorreoDestinoSolicitudVistoContinuidad "
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set oRS = oCon.Ejecutar(sSql)
    If Not oRS.BOF And Not oRS.EOF Then
        If oRS.RecordCount > 0 Then
            GetCorreoDestinoSolicitudVistoContinuidad = oRS!CorreoDestino
        Else
            GetCorreoDestinoSolicitudVistoContinuidad = ""
        End If
    Else
        GetCorreoDestinoSolicitudVistoContinuidad = ""
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrorGetCorreoDestinoSolicitudVistoContinuidad:
    Err.Raise Err.Number, "GetCorreoDestinoSolicitudVistoContinuidad ", Err.Description
    GetCorreoDestinoSolicitudVistoContinuidad = ""
End Function
'**********end marg*********************************************

'***APRI201601030
Public Function SugerenciaProductoAhorros(ByVal cPersCod As String, ByVal cCtaCod As String) As ADODB.Recordset
Dim sql As String
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim Conn As COMConecta.DCOMConecta
    bError = False
    
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set SugerenciaProductoAhorros = Nothing
        Set Conn = Nothing
        Set cmd = Nothing
        Set prm = Nothing
        Exit Function
    End If
    
    cPersCod = Replace(cPersCod, "'", "''")
    cCtaCod = Replace(cCtaCod, "'", "''")
    Conn.ConexionActiva.CommandTimeout = 7200
    sql = "Exec STP_SEL_SUGERENCIA_PRODUCTO '" & Trim(cPersCod) & "','" & Trim(cCtaCod) & "'"
    
    Set SugerenciaProductoAhorros = Conn.CargaRecordSet(sql)

    Conn.CierraConexion
    Set Conn = Nothing
    
End Function


'APRI 15-03-2017
Public Function SaldoDisponibleDPF(ByVal cPersCod As String) As ADODB.Recordset
   Dim sql As String
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter
    Dim Conn As COMConecta.DCOMConecta
    bError = False
    
    Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set SaldoDisponibleDPF = Nothing
        Set Conn = Nothing
        Set cmd = Nothing
        Set prm = Nothing
        Exit Function
    End If
    
    cPersCod = Replace(cPersCod, "'", "''")
    Conn.ConexionActiva.CommandTimeout = 7200
    sql = "Exec STP_OBTENER_SALDODISPONIBLEDPF '" & cPersCod & "'"
    Set SaldoDisponibleDPF = Conn.CargaRecordSet(sql)
    Conn.CierraConexion
    Set Conn = Nothing
End Function

Public Function RecuperarDatosPersonaGar(ByVal psPersCod As String) As ADODB.Recordset 'PTI1 20170315
'****************** Relaciones de Personas *********
        
        Dim rsPrueba As ADODB.Recordset
        Dim oConecta As COMConecta.DCOMConecta
        Dim sSql As String
        
        On Error GoTo ErrorRecuperarDatosPersonaGar
        sSql = "exec stp_sel_RecuperaRelacPersGar '" & psPersCod & "' "
    
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set RecuperarDatosPersonaGar = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
    Exit Function
ErrorRecuperarDatosPersonaGar:
    MsgBox Err.Description, vbInformation, "Aviso"
End Function
'***JGPA20191210 ACTA N° 106 - 2019
Public Function VerificaPersonaListaNegativa(ByVal psNombre As String, ByVal psApePat As String, ByVal psApeMat As String, ByVal psApeCas As String, ByVal psDoc As String, ByVal pnTipoBusqueda As Integer) As ADODB.Recordset 'JGPA20200117 Added [psDoc,pnTipoBusqueda]
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrVerificaPersonaListaNegativa
    If pnTipoBusqueda = 1 Then 'JGPA20200117
        sSql = "EXEC stp_sel_VerificaCondicionPersListaNegativaDoi '" & psNombre & "','" & psApePat & "','" & psApeMat & "','" & psApeCas & "','" & psDoc & "'" 'JGPA20200117 ANEXO 02  ACTA 106 - 2019
    ElseIf pnTipoBusqueda = 2 Then
        sSql = "Exec stp_sel_VerificaCondicionPersListaNegativa '" & psNombre & "','" & psApePat & "','" & psApeMat & "','" & psApeCas & "'"
    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set VerificaPersonaListaNegativa = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrVerificaPersonaListaNegativa:
    Err.Raise Err.Number, "Verifica Persona Lista Negativa", Err.Description
End Function
'***End JGPA20191210
'INICIO BY JATO 20210215
Public Function GetPuedeGenerar(ByVal gsCodUser As String) As Boolean
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    Dim R As ADODB.Recordset
    sSql = "EXEC stp_sel_puedeVerFirma '" & gsCodUser & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    If R.RecordCount > 0 Then
        If (R!nGenera = True) Then
            GetPuedeGenerar = True
        Else
            GetPuedeGenerar = False
        End If
    End If
    R.Close
    Set R = Nothing
End Function


