VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DPersonas"
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
Public Function RecuperaPersonasTipo(ByVal psTipo As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaPersonasTipo
    
    sSQL = "Select P.cPersCod, P.cPersNombre From Persona P inner join PersTpo PT"
    sSQL = sSQL & " ON P.cpersCod = PT.cPersCod Where PT.cPersTipo = '" & Right("00" & psTipo, 2) & "'"
    sSQL = sSQL & " Order By cPersNombre "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaPersonasTipo = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaPersonasTipo:
    Set RecuperaPersonasTipo = Nothing
    Err.Raise Err.Number, "Recupera Personas de un Tipo", Err.Description
End Function
Public Function BuscaCliente(ByVal sCadBusca As String, Optional ByVal nTipoBusq As TBusqueda = BusquedaNombre) As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim Conn As DConecta
    bError = False
'    Set cmd = New ADODB.Command
'    Set prm = New ADODB.Parameter
'    cmd.CommandText = "spBuscaClienteApellido"
'    cmd.CommandType = adCmdStoredProc
'    cmd.Name = "spBuscaClienteApellido"
'    Set prm = cmd.CreateParameter("Apellido", adChar, adParamInput, 80)
'    cmd.Parameters.Append prm
    
    Set Conn = New DConecta
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

Public Function BuscaDocumento(ByVal pnTipo As Integer, ByVal psNumero As String, ByVal psPersCod As String) As Boolean
Dim oConecta As DConecta
Dim sSQL As String
Dim R As ADODB.Recordset
    On Error GoTo ErrorBuscaDocumento
    sSQL = "Select cPersIDnro from " & gConsPersona & "PersID where cPersIDTpo = '" & Trim(Str(pnTipo)) & "' And cPersIDnro = '" & psNumero & "' And cPersCod <> '" & psPersCod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
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

Private Sub Class_Initialize()
Dim CIni As ClsIni.ClasIni
    Set CIni = New ClsIni.ClasIni
    gConsPersona = CIni.BasePersonas
    gConsComunes = CIni.BaseComunes
    gConsImagenes = CIni.BaseImagenes
    Set CIni = Nothing
End Sub
