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
    BusquedaProveedorAGNombre = 7
    BusquedaProveedorAGCodigo = 8
    BusquedaProveedorAGDocumento = 9
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

Public Function BuscaCliente(ByVal sCadBusca As String, Optional ByVal nTipoBusq As TBusqueda = BusquedaNombre, Optional psProveedorAG As String = "") As ADODB.Recordset
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
        Case 7 'Busqueda Proveedor AG por nombre
            Set BuscaCliente = Conn.CargaRecordSet("spBuscaProveedorAGApellido '" & Trim(sCadBusca) & "','" & psProveedorAG & "' ")
        Case 8 'Busqueda Proveedor AG por Codigo
            Set BuscaCliente = Conn.CargaRecordSet("spBuscaProveedorAGCodigo '" & Trim(sCadBusca) & "','" & psProveedorAG & "' ")
        Case 9 'Busqueda Proveedor AG por Documento
            Set BuscaCliente = Conn.CargaRecordSet("spBuscaProveedorAGDoc '" & Trim(sCadBusca) & "','" & psProveedorAG & "' ")
    End Select
    Conn.CierraConexion
    Set Conn = Nothing
End Function

Public Function BuscaDocumento(ByVal pnTipo As Integer, ByVal psNumero As String, ByVal psPersCod As String) As Boolean

Dim oConecta As DConecta
Dim sSQL As String
Dim R As ADODB.Recordset

    On Error GoTo ErrorBuscaDocumento
    
    sSQL = "Select cPersIDnro from " & gConsPersona & "PersID where cPersIDTpo = '" & Trim(Str(pnTipo)) & "' And cPersIDnro = '" & psNumero & "'"
    Set oConecta = New DConecta

    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    If Not R.BOF And Not R.EOF Then
        BuscaDocumento = True
    Else
        BuscaDocumento = False
    End If
    R.Close

    Call oConecta.CierraConexion
    
    
'    Call oConecta.AbreConexion 'Remota("07")
'
'    sSql = ""
'    sSql = sSql & " Select cNuDoCi From Persona Where cNuDoCi = '" & psNumero & "'"
'    sSql = sSql & " Union All"
'    sSql = sSql & " Select cNuDoTr From Persona Where cNuDoTr = '" & psNumero & "'"
'
'    Set R = oConecta.CargaRecordSet(sSql)

'    If Not R.BOF And Not R.EOF Then
'        BuscaDocumento = True
'    Else
'        BuscaDocumento = False
'    End If

'    R.Close
'    Call oConecta.CierraConexion
    
    Set R = Nothing
    Set oConecta = Nothing
    Exit Function
ErrorBuscaDocumento:
    Err.Raise Err.Number, "BuscaDocumento", Err.Description
    BuscaDocumento = False
End Function


Public Function ObtienePorcentajeGarantia(ByVal psCodTipoGar As Integer) As Double
Dim oConecta As DConecta
Dim R As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrorObtienePorcentajeGarantia
    sSQL = "Select * From Colocparametro where nParamVar = " & "1027" & psCodTipoGar
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
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
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim CIni As ClsIni.ClasIni
    Set CIni = New ClsIni.ClasIni
    gConsPersona = CIni.BasePersonas
    gConsComunes = CIni.BaseComunes
    gConsImagenes = CIni.BaseImagenes
    Set CIni = Nothing
End Sub

Public Function BuscaProveedor(ByVal sCadBusca As String, Optional ByVal nTipoBusq As TBusqueda = BusquedaNombre) As ADODB.Recordset
Dim lsCadena As String

Dim prm As ADODB.Parameter
Dim Conn As DConecta
    bError = False
    Set Conn = New DConecta
    'temporal
   ' lsCadena = "PROVIDER=SQLOLEDB;User ID=sa;Password=cmacica;INITIAL CATALOG=dbcmacicaproveedores;DATA SOURCE=01SRVSICMAC02"
    'If Not Conn.AbreConexion(lsCadena) Then
    If Not Conn.AbreConexion Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set BuscaProveedor = Nothing
        Set Conn = Nothing
        
        Exit Function
    End If
    Conn.ConexionActiva.CommandTimeout = 7200
    Select Case nTipoBusq
        Case 1 'Busqueda por nombre
            Set BuscaProveedor = Conn.CargaRecordSet("spBuscaProveedorApellido '" & Trim(sCadBusca) & "'")
        'Case 2 'Busqueda por Codigo
        '    Set BuscaCliente = Conn.CargaRecordSet("spBuscaClienteCodigo '" & Trim(sCadBusca) & "'")
        Case 3 'Busqueda por Documento
            Set BuscaProveedor = Conn.CargaRecordSet("spBuscaProveedoresDoc '" & Trim(sCadBusca) & "'")
            
    End Select
    Conn.CierraConexion
    Set Conn = Nothing
End Function

Public Function ValidaCodPersona(codigo As String, TipoCod) As Integer


Dim Conn As DConecta
Dim rs As ADODB.Recordset
Dim sSQL As String

Set rs = New ADODB.Recordset
    bError = False
    ValidaCodPersona = 0
    Set Conn = New DConecta
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        ValidaCodPersona = 0
        Set Conn = Nothing
        
        Exit Function
    End If
    Conn.ConexionActiva.CommandTimeout = 7200
    If TipoCod = 1 Then 'cod persona
          sSQL = "select count(*) as cuenta from persona where cPersCod ='" & codigo & "'"
          Set rs = Conn.CargaRecordSet(sSQL)
    ElseIf TipoCod = 2 Then  'Numerodoc
          sSQL = " select count(*) as cuenta from persid where cPersIDnro ='" & codigo & "'"
          Set rs = Conn.CargaRecordSet(sSQL)
    End If
    If rs!Cuenta > 0 Then
        ValidaCodPersona = 1
        ElseIf rs!Cuenta = 0 Then
        ValidaCodPersona = 0
    End If
    
    Conn.CierraConexion
    Set Conn = Nothing
End Function

Public Function InsertaProveedor(CodPersona As String) As Integer

Dim sSQL As String

Dim Conn As DConecta
    
    Set Conn = New DConecta
    
    InsertaProveedor = 0
    If Not Conn.AbreConexion() Then
        InsertaProveedor = 0
        Set Conn = Nothing
        Exit Function
    End If
    Conn.ConexionActiva.CommandTimeout = 7200
    sSQL = "spInsertaPersonaProveedor  '" & Trim(CodPersona) & "' "
    Conn.Ejecutar sSQL
    InsertaProveedor = 1
    Conn.CierraConexion
    Set Conn = Nothing
End Function
'WIOR 20130827 *******************************************************************
Public Sub RegistroDatosCambiosDatosPrinc(ByVal pbDocs As Boolean, ByVal psMovCambio As String, ByVal psPersCod As String, _
                Optional ByVal psNomAnt As String = "", Optional ByVal psNomAct As String = "", _
                Optional ByVal pnTipoDocAnt As Integer = 0, Optional ByVal pnTipoDocAct As Integer = 0, _
                Optional ByVal psNumDocAnt As String = "", Optional ByVal psNumDocAct As String = "")

Dim sSQL As String
Dim oCon As New DConecta
    
If Not pbDocs Then
    sSQL = "EXEC stp_ins_PersonaCambios '" & psPersCod & "','" & psNomAnt & "','" & psNomAct & "','" & psMovCambio & "'"
Else
    sSQL = "EXEC stp_ins_PersonaCambiosDoc '" & psPersCod & "'," & pnTipoDocAnt & "," & pnTipoDocAct & ",'" & psNumDocAnt & "','" & psNumDocAct & "','" & psMovCambio & "'"
End If

oCon.AbreConexion
oCon.Ejecutar sSQL
oCon.CierraConexion
Set oCon = Nothing
End Sub
Public Sub RegistroCabeceraDatosCambiosDatosPrinc(ByVal psMovCambio As String, ByVal psPersCod As String, ByVal pnAplicacion As Integer)
Dim sSQL As String
Dim oCon As New DConecta

sSQL = "EXEC stp_ins_PersonaCambiosCabecera '" & psPersCod & "','" & psMovCambio & "'," & pnAplicacion

oCon.AbreConexion
oCon.Ejecutar sSQL
oCon.CierraConexion
Set oCon = Nothing
End Sub
'WIOR FIN ************************************************************************





