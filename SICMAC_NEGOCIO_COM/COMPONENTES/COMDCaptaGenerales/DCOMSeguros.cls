VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMSeguros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'FRHU 20140610 ERS068-2014
Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String
Dim oError As COMConecta.COMErrorHandling
Public coConex As COMConecta.DCOMConecta
'FIN FRHU

Public bTransaction As Boolean 'RIRO 20170508

'FRHU 20140528 ************************************************************************
'REMESA: CERTIFICADOS POR AGENCIAS
Public Function ObtenerSegTarjetaRemesaAgencia() As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_SegTarjetaRemesaAgencia"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    Set ObtenerSegTarjetaRemesaAgencia = rs
    Set oConec = Nothing
End Function
Public Sub InsertaSegTarjetaRemesaAgencia(ByVal psAgeCod As String, ByVal psRangoDel As String, ByVal psRangoAl As String, ByVal pdFecha As Date, ByVal pnEstado As Integer)
    Dim lsSQL As String
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    lsSQL = "exec stp_ins_SegTarjetaRemesaAgencia '" & psAgeCod & "','" & psRangoDel & "','" & psRangoAl & "','" & Format(pdFecha, "yyyymmdd") & "'," & pnEstado
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
Public Function ValidarSegTarjetaRemesaAgencia(ByVal pnId As Integer, ByVal psRangoDel As String, ByVal psRangoAl As String, ByVal pnOperacion As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_ValidarSegTarjetaRemesaAgencia " & pnId & ",'" & psRangoDel & "','" & psRangoAl & "'," & pnOperacion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    Set ValidarSegTarjetaRemesaAgencia = rs
    Set oConec = Nothing
End Function
Public Sub ActualizarSegTarjetaRemesaAgencia(ByVal pnId As Integer, ByVal psRangoDel As String, ByVal psRangoAl As String, ByVal pdFecha As Date, ByVal pnEstado As Integer, ByVal pnOperacion As Integer)
    Dim lsSQL As String
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    lsSQL = "exec stp_upd_SegTarjetaRemesaAgencia " & pnId & ",'" & psRangoDel & "','" & psRangoAl & "','" & Format(pdFecha, "yyyymmdd") & "'," & pnEstado & "," & pnOperacion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
'FIN REMESA: CERTIFICADOS POR AGENCIAS
'FIN FRHU 20140528 ********************************************************************
'JUEZ 20140515 **********************************************************************
Public Function VerificaSegTarjetaAfiliacion(ByVal psNumTarj As String) As Boolean
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConec As New COMConecta.DCOMConecta

    VerificaSegTarjetaAfiliacion = False
    oConec.AbreConexion
    sSql = "exec stp_sel_VerificaSegTarjetaAfiliacion '" & psNumTarj & "'"
    Set rs = oConec.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        VerificaSegTarjetaAfiliacion = True
    End If
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function RecuperaSegTarjetaParametro(ByVal psParamVar As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_RecuperaSegTarjetaParametro " & psParamVar
    Set RecuperaSegTarjetaParametro = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function RecuperaNroSolicitudAfiliacionMes(ByVal psNumTarj As String, ByVal pnAnio As Integer, ByVal pnMes As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_RecuperaNroSolicitudAfiliacionMes '" & psNumTarj & "'," & pnAnio & "," & pnMes
    Set RecuperaNroSolicitudAfiliacionMes = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Sub InsertaNroSolicitudAfiliacionMes(ByVal psNumTarj As String, ByVal pdFecSis As String)
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta
Dim pbTran As Boolean
Dim lrDatos As ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    lsSQL = "exec stp_ins_InsertaNroSolicitudAfiliacionMes '" & psNumTarj & "','" & Format(pdFecSis & " " & oConec.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "'"
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
Public Function RecuperaPersonaTarjeta(ByVal psNumTarj As String) As ADODB.Recordset
Dim sSql As String
Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_RecuperaPersonaTarjeta '" & psNumTarj & "'"
    Set RecuperaPersonaTarjeta = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Sub InsertaSegTarjetaAfiliacion(ByVal psNumCert As String, ByVal psNumTarj As String, ByVal psCtaCod As String, _
                                       ByVal pdFecSis As String, ByVal psMovNro As String, ByVal pnMovNro As Long, ByVal psPersCod As String)
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta
Dim pbTran As Boolean
Dim lrDatos As ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    lsSQL = "exec stp_ins_SegTarjetaAfiliacion '" & psNumCert & "','" & psNumTarj & "','" & psCtaCod & "','" & _
            Format(pdFecSis & " " & oConec.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & psMovNro & "'," & pnMovNro & ",'" & psPersCod & "'"
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
Public Function RecuperaSegTarjetaAfiliacion(Optional ByVal pnMovNro As Long = 0, Optional ByVal psNumTarj As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_SegTarjetaAfiliacion " & pnMovNro & ",'" & psNumTarj & "'"
    Set RecuperaSegTarjetaAfiliacion = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function ValidaExisteRegistroNroCertificado(ByVal cNumCertificado As String, Optional ByVal nOpeTpo As SegCertificadoTipo = 1) As Boolean
Dim sSql As String
Dim oConec As New COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

    ValidaExisteRegistroNroCertificado = False
    oConec.AbreConexion
    sSql = "exec stp_sel_ValidaExisteRegistroNroCertificado '" & cNumCertificado & "'," & nOpeTpo
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    If Not (rs.EOF And rs.BOF) Then
        ValidaExisteRegistroNroCertificado = True
    End If
    Set oConec = Nothing
End Function
Public Function ValidaNroCertificadoRemesaAgencia(ByVal cNumCertificado As String, ByVal psCodAge As String) As Boolean
Dim sSql As String
Dim oConec As New COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

    ValidaNroCertificadoRemesaAgencia = False
    oConec.AbreConexion
    sSql = "exec stp_sel_ValidaNroCertificadoRemesaAgencia '" & cNumCertificado & "','" & psCodAge & "'"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    If Not (rs.EOF And rs.BOF) Then
        ValidaNroCertificadoRemesaAgencia = True
    End If
    Set oConec = Nothing
End Function
Public Function RecuperaSegAfiliacionTarjetaExtorno(ByVal cNumTarjeta As String, ByVal pdFecSis As String) As ADODB.Recordset
Dim sSql As String
Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_RecuperaSegAfiliacionTarjetaExtorno '" & cNumTarjeta & "','" & Format(pdFecSis, "yyyyMMdd") & "'"
    Set RecuperaSegAfiliacionTarjetaExtorno = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Sub ActualizaEstadoSegTarjetaAfiliacion(ByVal pcNumCertificado As String, ByVal pnMovNro As Long, ByVal pnSegEstado As Integer)
'APRI20171027 ADD pcNumCertificado ERS028-2017
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta
Dim pbTran As Boolean
Dim lrDatos As ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    lsSQL = "exec stp_upd_SegTarjetaAfiliacion '" & pcNumCertificado & "'," & pnMovNro & "," & pnSegEstado 'APRI20171027
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
'END JUEZ ***************************************************************************
'FRHU 20140610 ERS068-2014 **********************************************************
'transact
Public Sub dBeginTrans()
    'If pbConcurrencia Then
    '    coConex.ConexionActiva.Execute "SET TRANSACTION ISOLATION LEVEL SERIALIZABLE"
    'End If
    coConex.BeginTrans
End Sub
Public Sub dRollbackTrans()
    coConex.RollbackTrans
    coConex.Ejecutar "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
End Sub
Public Sub dCommitTrans()
    coConex.CommitTrans
    'coConex.Ejecutar "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
End Sub
Private Sub Class_Initialize()
    Dim loIni As COMConecta.DCOMClasIni
    
    Set loIni = New COMConecta.DCOMClasIni
        csConexion = loIni.CadenaConexion
        csNegocio = loIni.BaseNegocio
        csCentralPer = loIni.BasePersonas
        csCentralCom = loIni.BaseComunes
        csCentralImg = loIni.BaseImagenes
        csAdminist = loIni.BaseAdministracion
    Set loIni = Nothing
    
Set coConex = New COMConecta.DCOMConecta
Set oError = New COMConecta.COMErrorHandling

If coConex.AbreConexion(csConexion) = False Then
    Call oError.RaiseError(oError.MyUnhandledError, "DCOMSeguros:Initialize. Error en Conexion a Base de datos")
End If

End Sub
'RIRO 20170508 Modificacion
Private Sub Class_Terminate()
    If bTransaction = True Then
    Else
        coConex.CierraConexion
        Set coConex = Nothing
    End If
End Sub
Public Function dEjecutaBatch() As Integer
    'Ejecuta procesos Batch
    dEjecutaBatch = coConex.EjecutarBatch
    Exit Function
End Function
'fin transact
Public Function ObtenerSegTarjetaConfigDoc(ByVal pnOperacion As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_SegTarjetaConfigDoc " & pnOperacion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    Set ObtenerSegTarjetaConfigDoc = rs
    Set oConec = Nothing
End Function
Public Sub InsertaActualizaSegTarjetaConfigDoc(ByVal pnId As Integer, ByVal psDescripcion As String, ByVal pnEstado As Integer)
    Dim lsSQL As String
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    lsSQL = "exec stp_ins_upd_SegTarjetaConfigDoc " & pnId & ",'" & psDescripcion & "'," & pnEstado
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
Public Function QuitarSegTarjetaConfigDoc(ByVal pnId As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_upd_SegTarjetaQuitarConfigDoc " & pnId
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    Set QuitarSegTarjetaConfigDoc = rs
    Set oConec = Nothing
End Function
'SOLICITUD: Proceso de Activación de Seguro de Tarjeta
Public Function ValidarSegTarjetaGrabarSolicitud(ByVal psNumSolicitud As String, ByVal psNumActSiniestro As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_ValidarRegistroSolicitudSegTarjeta '" & psNumSolicitud & "','" & psNumActSiniestro & "'"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    Set ValidarSegTarjetaGrabarSolicitud = rs
    Set oConec = Nothing
End Function
Public Sub InsertaSegTarjetaSolicitud(ByVal psNumSolicitud As String, ByVal pdFechaSolicitud As Date, ByVal pnMontoSolicitud As Currency, ByVal pnPrdEstado As Integer, _
                                      ByVal cPersCodAsegurado As String, ByVal cPersCodTercero As String, ByVal pcRelacionTercero As String, _
                                      ByVal psNumActSiniestro As String, ByVal psAgeCodSiniestro As String, ByVal psDescripcionSiniestro As String, ByVal psUbigeoSiniestro As String, _
                                      ByVal psDireccionRefSiniestro As String, ByVal pdFechaSiniestro As Date, ByVal psHoraSiniestro As String, ByVal cUltimaActualizacion As String)
    Dim lsSQL As String
    
    lsSQL = "exec stp_ins_SegTarjetaSolicitud '" & psNumSolicitud & "','" & Format(pdFechaSolicitud, "yyyymmdd") & "'," & pnMontoSolicitud & "," & pnPrdEstado & ","
    lsSQL = lsSQL & "'" & cPersCodAsegurado & "','" & cPersCodTercero & "','" & pcRelacionTercero & "',"
    lsSQL = lsSQL & "'" & psNumActSiniestro & "','" & psAgeCodSiniestro & "','" & psDescripcionSiniestro & "','" & psUbigeoSiniestro & "',"
    lsSQL = lsSQL & "'" & psDireccionRefSiniestro & "','" & Format(pdFechaSiniestro, "yyyymmdd") & "','" & psHoraSiniestro & "','" & cUltimaActualizacion & "'"
    coConex.CargaRecordSet (lsSQL)
End Sub
Public Sub InsertaSegTarjetaSolicitudSiniestro(ByVal psNumSolicitud As String, ByVal pnTipoSiniestro As Integer)
    Dim lsSQL As String

    lsSQL = "exec stp_ins_SegTarjetaSolicitudSiniestro '" & psNumSolicitud & "'," & pnTipoSiniestro
    coConex.CargaRecordSet (lsSQL)
End Sub
Public Sub InsertaSegTarjetaSolicitudDocumento(ByVal psNumSolicitud As String, ByVal pnNumDocumento As Integer, ByVal pnTipoDocumento As Integer)
    Dim lsSQL As String

    lsSQL = "exec stp_ins_SegTarjetaSolicitudDocumento '" & psNumSolicitud & "'," & pnNumDocumento & "," & pnTipoDocumento
    coConex.CargaRecordSet (lsSQL)
End Sub
Public Sub InsertaSegTarjetaEstado(ByVal psNumSolicitud As String, ByVal pdPrdEstado As Date, ByVal pnPrdEstado As Integer, ByVal pcDescripcion As String, ByVal pnMonto As Currency, ByVal psUltimaActualizacion As String)
    Dim lsSQL As String

    lsSQL = "exec stp_ins_SegTarjetaSolicitudEstado '" & psNumSolicitud & "','" & Format(pdPrdEstado, "yyyymmdd") & "'," & pnPrdEstado & ",'" & pcDescripcion & "'," & pnMonto & ",'" & psUltimaActualizacion & "'"
    coConex.CargaRecordSet (lsSQL)
End Sub
Public Function ValidarTarjetasAfiliadasSegTarjetaXPersona(ByVal psCodPersona As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_TarjetasAfiliadasAlSegTarjetaXPersona '" & psCodPersona & "'"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    Set ValidarTarjetasAfiliadasSegTarjetaXPersona = rs
    Set oConec = Nothing
End Function
Public Function ValidarSolicitudesEnProcesoSegTarjeta(ByVal psCodPersona As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_ValidarSolicitudesEnProcesoSegTarjeta '" & psCodPersona & "'"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    Set ValidarSolicitudesEnProcesoSegTarjeta = rs
    Set oConec = Nothing
End Function
'rechazo solicitud
Public Function ObtenerSegTarjetSolictudXNSolicitud(ByVal psNumSolicitud As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_SegTarjetaSolicitudXNSolicitud '" & psNumSolicitud & "'"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    Set ObtenerSegTarjetSolictudXNSolicitud = rs
    Set oConec = Nothing
End Function
Public Sub ActualizarSegTarjetaSolicitudARechazado(ByVal psNumSolicitud As String, ByVal psMovNro As String, ByVal pnEstado As Integer, ByVal pdFechaResp As Date)
    Dim lsSQL As String

    lsSQL = "exec stp_upd_SegTarjetaRechazoSolicitud '" & psNumSolicitud & "','" & psMovNro & "'," & pnEstado & ",'" & Format(pdFechaResp, "yyyymmdd") & "'"
    coConex.CargaRecordSet (lsSQL)

End Sub
'aceptacion de solicitudes
Public Function ObtenerSegTarjetaMovAprobadas() As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_SegTarjetaMovAprobados"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    Set ObtenerSegTarjetaMovAprobadas = rs
    Set oConec = Nothing
End Function
Public Function ObtenerSegTarjetaSolicitudesAprobadasXMov(ByVal psMovNro As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_SegTarjetaSolicitudesAprobadosXMov '" & psMovNro & "'"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    Set ObtenerSegTarjetaSolicitudesAprobadasXMov = rs
    Set oConec = Nothing
End Function
Public Sub ActualizarSegTarjetaSolicitudAceptado(ByVal psNumSolicitud As String, ByVal psUltimaActualizacion As String, ByVal pnPrdEstado As Integer, ByVal pnMonto As Currency)
    Dim lsSQL As String
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    lsSQL = "Exec stp_upd_SegTarjetaSolicitudAceptado '" & psNumSolicitud & "','" & psUltimaActualizacion & "'," & pnPrdEstado & "," & pnMonto
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
Public Sub InsertarSegTarjetaSolicitudEstadoAceptado(ByVal psNumSolicitud As String, ByVal pdPrdEstado As Date, ByVal pnPrdEstado As Integer, ByVal psDescripcion As String, ByVal pnMonto As Currency, ByVal psUltimaActualizacion As String)
    Dim lsSQL As String
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    lsSQL = "Exec stp_ins_SegTarjetaSolicitudEstadoAceptado '" & psNumSolicitud & "','" & Format(pdPrdEstado, "yyyymmdd") & "'," & pnPrdEstado & ",'" & psDescripcion & "'," & pnMonto & ",'" & psUltimaActualizacion & "'"
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
'FIN FRHU 20140610 ********************************************************
'JUEZ 20140616 *******************************************************
Public Function ObtenerSegTarjetaNumCertificado() As String
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_ObtenerSegTarjetaNumCertificado"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not (rs.EOF And rs.BOF) Then
        ObtenerSegTarjetaNumCertificado = rs!cNumCertificado
    End If
    
    Set oConec = Nothing
End Function
'END JUEZ ***********************************************************
'JUEZ 20140620 ***************************************************************
Public Sub ActualizarSegTarjetaDesafiliadoTresMeses(ByVal psMovNro As String)
Dim lsSQL As String
Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    lsSQL = "Exec stp_upd_SegTarjetaDesafiliadoTresMeses '" & psMovNro & "'"
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
Public Function GeneraTramaSegTarjeta(ByVal pnAnio As Integer, ByVal pnMes As Integer, ByVal pnTipoTrama As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_GeneraTramaSegTarjeta " & pnAnio & "," & pnMes & "," & pnTipoTrama
    Set GeneraTramaSegTarjeta = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Sub GrabarGeneracionTramaSegTarjeta(ByVal psFecSis As String, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psTipoTrama As String, _
                                           ByVal pnCantReg As Integer, ByVal pnPrimaTotal As Double, ByVal psMovNro As String)
Dim lsSQL As String
Dim oConec As New COMConecta.DCOMConecta
    oConec.AbreConexion
    lsSQL = "Exec stp_ins_GrabarGeneracionTramaSegTarjeta '" & Format(psFecSis & " " & oConec.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & pnMes & "','" & _
                                                               pnAnio & "','" & psTipoTrama & "'," & pnCantReg & "," & pnPrimaTotal & ",'" & psMovNro & "'"
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
Public Function RecuperaGeneracionTramaSegTarjeta(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psTipoTrama As String) As Boolean
Dim sSql As String
Dim oConec As New COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
    
    RecuperaGeneracionTramaSegTarjeta = False
    oConec.AbreConexion
    sSql = "exec stp_sel_RecuperaGeneracionTramaSegTarjeta '" & pnMes & "','" & pnAnio & "','" & psTipoTrama & "'"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    If Not (rs.EOF And rs.BOF) Then
        RecuperaGeneracionTramaSegTarjeta = True
    End If
    Set oConec = Nothing
End Function
'END JUEZ ********************************************************************
'FRHU 20140903
Public Function ObtenerSegTarjetaMovPendienteXcMovNro(ByVal psMovNro As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_SegTarjetaMovPendienteXcMovNro '" & psMovNro & "'"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    Set ObtenerSegTarjetaMovPendienteXcMovNro = rs
    Set oConec = Nothing
End Function
'FIN FRHU 20140903
'JUEZ 20150510 *************************************************************
Public Sub InsertaSegTarjetaAnulaDevPendiente(ByVal psNumCert As String, ByVal psPersCod As String, ByVal psCtaCodDeb As String, ByVal pdFecAnula As Date, ByVal pnMontoCom As String)
    Dim lsSQL As String

    lsSQL = "exec stp_ins_SegTarjetaAnulaDevPendiente '" & psNumCert & "','" & psPersCod & "','" & psCtaCodDeb & "','" & Format(pdFecAnula, "yyyyMMdd") & "'," & pnMontoCom
    coConex.CargaRecordSet (lsSQL)
End Sub
Public Function RecuperaSegTarjetaAnulaDevPend(ByVal pbRegistrado As Boolean, ByVal pbPagoAseg As Boolean)
Dim sSql As String
Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_SegTarjetaAnulaDevPendiente " & IIf(pbRegistrado, 1, 0) & "," & IIf(pbPagoAseg, 1, 0)
    Set RecuperaSegTarjetaAnulaDevPend = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion

    Set oConec = Nothing
End Function
Public Sub ActualizaSegTarjetaAnulaDevPendiente(ByVal psNumCert As String, ByVal psPersCod As String, ByVal pbRegistrado As Boolean, ByVal pbPagoAseg As Boolean)
    Dim lsSQL As String

    lsSQL = "exec stp_upd_SegTarjetaAnulaDevPendiente '" & psNumCert & "','" & psPersCod & "'," & IIf(pbRegistrado, 1, 0) & "," & IIf(pbPagoAseg, 1, 0) & ",0"
    coConex.CargaRecordSet (lsSQL)
End Sub
'END JUEZ ******************************************************************
'RECO20151102 ERS073-2015***************************************************
Public Function ObtieneParametrosSegSepelio()
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_ObtieneParametrosSegSepelio "
    Set ObtieneParametrosSegSepelio = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion

    Set oConec = Nothing
End Function
Public Function ObtieneCriteriosSegSepelio()
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_ObtieneCriteriosSegSepelio "
    Set ObtieneCriteriosSegSepelio = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion

    Set oConec = Nothing
End Function
Public Sub ActualizaValoresParamSepelio(ByVal pnParamID As Integer, ByVal pnValor As Double)
    Dim lsSQL As String

    lsSQL = "exec stp_upd_ActualizaValoresParamSepelio " & pnParamID & "," & pnValor
    coConex.CargaRecordSet (lsSQL)
End Sub
Public Sub ActualizaCriterioSegSepelio(ByVal pnCriterioID As Integer, ByVal pnValAnio As Integer, ByVal pnValMes As Integer, ByVal pnValDia As Integer)
    Dim lsSQL As String

    lsSQL = "exec stp_upd_ActualizaCriterioSegSepelio " & pnCriterioID & "," & pnValAnio & "," & pnValMes & "," & pnValDia
    coConex.CargaRecordSet (lsSQL)
End Sub
Public Function ObtieneParamNumCertifSepelio()
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_ObtieneParamNumCertifSepelio "
    Set ObtieneParamNumCertifSepelio = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion

    Set oConec = Nothing
End Function
Public Sub ActualizaParamNumCertifSepelio(ByVal psAgeCod As String, ByVal pnDesde As Integer, ByVal pnHasta As Integer, ByVal pnStock As Integer, ByVal psUser As String)
    Dim lsSQL As String

    lsSQL = "exec stp_upd_ActualizaParamNumCertifSepelio '" & psAgeCod & "'," & pnDesde & "," & pnHasta & "," & pnStock & ",'" & psUser & "'"
    coConex.CargaRecordSet (lsSQL)
End Sub

Public Function ObtenerSegSepeleioNumCertificado() As String
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_ObtenerSegSepelioNumCertificado"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not (rs.EOF And rs.BOF) Then
        ObtenerSegSepeleioNumCertificado = rs!cNumCertificado
    End If
    
    Set oConec = Nothing
End Function
Public Function RecuperaSegSepelioParametro(ByVal psParamVar As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_RecuperaSegSepelioParametro " & psParamVar
    Set RecuperaSegSepelioParametro = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Sub RegistraAseguradoSepelio(ByVal psNumCertificado As String, ByVal psAseguradoNombre As String, ByVal pnRelacion As Integer, ByVal pnParticipacion As Integer, ByVal pnEstado As Integer, Optional ByVal pConn As COMConecta.DCOMConecta = Nothing) 'RIRO 20170609 ADD pConn
    Dim lsSQL As String
    'RIRO 20170505 ADD ***
    If Not pConn Is Nothing Then
        Set coConex = pConn
        bTransaction = True
    End If
    'END RIRO ****
    lsSQL = "exec stp_ins_RegistraAseguradoSepelio '" & psNumCertificado & "','" & psAseguradoNombre & "'," & pnRelacion & "," & pnParticipacion & "," & pnEstado
    coConex.CargaRecordSet (lsSQL)
End Sub
Public Function ObtieneDatosPersSegSepelio(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_ObtieneDatosPersSegSepelio '" & psPersCod & "'"
    Set ObtieneDatosPersSegSepelio = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function ActualizaEstadoSeguroSepelio(ByVal psNumCertificado As String, ByVal psFecAfiliacion As String, ByVal pnMovNroReg As Long, ByVal pnEstado As Integer) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As New COMConecta.DCOMConecta
    Dim sFecAfiliacion As String
    
    sFecAfiliacion = IIf(Len(psFecAfiliacion) = 8, psFecAfiliacion, Format(psFecAfiliacion, "yyyyMMdd") & " " & Format(psFecAfiliacion, "hh:mm:ss"))
    oConec.AbreConexion
    
    lsSQL = "exec stp_upd_ActualizaEstadoSegSepelio '" & psNumCertificado & "','" & sFecAfiliacion & "'," & pnMovNroReg & "," & pnEstado
    Set ActualizaEstadoSeguroSepelio = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    
End Function
Public Function ObtieneClienteAfiliacionSepelio(ByVal psFecha As String, ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_ObtieneClienteAfiliacionSepelio '" & psFecha & "','" & psPersCod & "'"
    Set ObtieneClienteAfiliacionSepelio = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function DevulveClienteSeguroActivo(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_DevulveClienteSeguroActivo '" & psPersCod & "'"
    Set DevulveClienteSeguroActivo = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function VerificaAfiliacionDesactivada(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    
    oConec.AbreConexion
    sSql = "exec stp_sel_VerificaAfiliacionDesactivada '" & psPersCod & "'"
    Set VerificaAfiliacionDesactivada = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function SegSepelioDatosOperacion(ByVal pnMovNro As Long, Optional ByVal pConn As COMConecta.DCOMConecta = Nothing) As ADODB.Recordset
    'RIRO20170612 ADD pConn

    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    
    'RIRO 20170505 ADD ***
    If Not pConn Is Nothing Then
        Set oConec = pConn
        bTransaction = True
    Else
        oConec.AbreConexion
    End If
    'END RIRO ****
    
    sSql = "stp_sel_SegSepelioDatosOperacion " & pnMovNro
    Set SegSepelioDatosOperacion = oConec.CargaRecordSet(sSql)
    
    'RIRO 20170612 ***
    If pConn Is Nothing Then
        oConec.CierraConexion
        Set oConec = Nothing
    End If
    'END RIRO 20170612 ***

End Function
Public Function SepelioBeneficiarios(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
        oConec.AbreConexion
    sSql = "stp_sel_SepelioBeneficiarios " & psPersCod
    Set SepelioBeneficiarios = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function SepelioCoberturas() As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
        oConec.AbreConexion
    sSql = "stp_sel_SepelioCoberturas "
    Set SepelioCoberturas = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Sub RegistraSepelioCoberturas(ByVal pnIdSolicitud As Integer, ByVal pnIdCobertura As Integer)
    Dim lsSQL As String
    
    lsSQL = "exec stp_ins_SepelioSolicitudCoberturas " & pnIdSolicitud & "," & pnIdCobertura
    coConex.CargaRecordSet (lsSQL)
End Sub
Public Sub SepelioSolicCtaBenef(ByVal pnAseguradoID As Integer, ByVal psCtaCod As String)
    Dim lsSQL As String
    
    lsSQL = "exec stp_upd_SepelioSolicCtaBenef " & pnAseguradoID & "," & psCtaCod
    coConex.CargaRecordSet (lsSQL)
End Sub
Public Function SepelioDevolverMovimientoExtornar(ByVal psDatoBusqueda As String, ByVal psFecha As String, ByVal psOpeCod As String, ByVal psAgeCod As String, ByVal pnTipoBusqueda As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
        oConec.AbreConexion
    sSql = "stp_sel_SepelioDevolverMovimientoExtornar '" & psDatoBusqueda & "','" & psFecha & "','" & psOpeCod & "','" & psAgeCod & "'," & pnTipoBusqueda
    Set SepelioDevolverMovimientoExtornar = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function SepelioObtieneCriterios() As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
        oConec.AbreConexion
    sSql = "stp_sel_SepelioObtieneCriterios "
    Set SepelioObtieneCriterios = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function SepelioObtieneDatosExtorno(ByVal pnMovNro As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
        oConec.AbreConexion
    sSql = "stp_sel_SepelioObtieneDatosExtorno " & pnMovNro
    Set SepelioObtieneDatosExtorno = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function SepelioObtieneDatos(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
        oConec.AbreConexion
    sSql = "stp_sel_ObtieneDatosSegSepelio '" & psPersCod & "'"
    Set SepelioObtieneDatos = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Sub SepelioActualizaDatos(ByVal psNumCertifNue As String, ByVal psNumCertifAnt As String, ByVal pnTpoPago As Integer, ByVal psCtaCod As String, ByVal pnMovNro As Long, ByVal pnEstado As Integer)
    Dim lsSQL As String
    
    lsSQL = "exec stp_upd_SepelioActualizaDatos '" & psNumCertifNue & "','" & psNumCertifAnt & "'," & pnTpoPago & ",'" & psCtaCod & "'," & pnMovNro & "," & pnEstado
    coConex.CargaRecordSet (lsSQL)
End Sub
Public Sub SepelioLimpiaBeneficiario(ByVal psNumCertif As String)
    Dim lsSQL As String
    
    lsSQL = "exec stp_upd_SepelioLimpiaBeneficiario '" & psNumCertif & "'"
    coConex.CargaRecordSet (lsSQL)
End Sub
Public Function SepelioVerificaPagoMes(ByVal psFecSis As String, ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
        oConec.AbreConexion
    sSql = "stp_sel_SepelioVerificaPagoMes '" & psFecSis & "','" & psPersCod & "'"
    Set SepelioVerificaPagoMes = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

'RECO FIN ******************************************************************
'APRI20171004 ERS028-2017
Public Function SepelioVerificaTpoPrograma(ByVal nTpoSeguro As Integer, ByVal nProd As String, ByVal nTpoPrograma As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "STP_SEL_VALIDATPOPRODSEGURO " & nTpoSeguro & ",'" & nProd & "'," & nTpoPrograma
    Set SepelioVerificaTpoPrograma = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function SepelioVerificaPagoEfectivo(ByVal sPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "STP_VALIDASEPELIOEFECTIVO '" & sPersCod & "'"
    Set SepelioVerificaPagoEfectivo = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function SepelioVerificaProducto(ByVal sPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "Exec stp_sel_CuentasAhorrosSepelio '" & sPersCod & "'"
    Set SepelioVerificaProducto = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function SepelioParametros(ByVal sParamID As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "STP_SEL_PARAMSEGSEPELIO " & sParamID
    Set SepelioParametros = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Sub ActualizaPoliza(ByVal psCtaCod As String)
    Dim lsSQL As String
    lsSQL = "exec STP_UP_POLIZA '" & psCtaCod & "'"
    coConex.CargaRecordSet (lsSQL)
End Sub
Public Function InsertarObservacionesCredSegDesg(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
     Dim oConec As New COMConecta.DCOMConecta
    lsSQL = "exec STP_INS_SEGDESGCREDITOOBS '" & psCtaCod & "'"
    Set InsertarObservacionesCredSegDesg = coConex.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function ObtenerObservacionesCredSegDesg(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As New COMConecta.DCOMConecta
    oConec.AbreConexion
    lsSQL = "exec STP_SEL_CREDITOSOBSSEGDESG '','" & psCtaCod & "'"
    Set ObtenerObservacionesCredSegDesg = coConex.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'END APRI

'CTI2 FERIMORO ERS034-2019  25082020
Public Sub ActuSeguroCronograma(ByVal tpoOpcion As Integer, ByVal psFecAfiliacion As String, ByVal tpoSeguro As Integer, _
                                ByVal pnMovNroReg As Long, ByVal psNumCertificado As String, _
                                ByVal nExisteCrono As Integer)
    Dim lsSQL As String
    
    lsSQL = "exec stp_cns_ERS034_2019_PagosSegurosCronogramas '" & Format(psFecAfiliacion, "YYYYmmdd") & "'," & tpoSeguro & ", " & pnMovNroReg & ",'" & psNumCertificado & "', " & nExisteCrono & ""
    coConex.CargaRecordSet (lsSQL)
    
End Sub
'CTI2 FERIMORO ERS034-2019    29082020
Public Function BuscaFormulariosExisteD(ByVal psNroFormulario As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As New COMConecta.DCOMConecta
    oConec.AbreConexion
    lsSQL = "exec stp_cns_ERS034_2019_VerificaExisteFormulario '" & psNroFormulario & "'"
    Set BuscaFormulariosExisteD = coConex.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function ObtieneDatosSegSoat(ByVal pnMovNro As Long, Optional ByVal pConn As COMConecta.DCOMConecta = Nothing) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As New COMConecta.DCOMConecta
    If Not pConn Is Nothing Then
        Set oConec = pConn
        bTransaction = True
    Else
        oConec.AbreConexion
    End If
    
    sSql = "stp_cns_ERS034_2019_ObtieneDatosSegSoat " & pnMovNro
    Set ObtieneDatosSegSoat = oConec.CargaRecordSet(sSql)

    If pConn Is Nothing Then
        oConec.CierraConexion
        Set oConec = Nothing
    End If
End Function
'CTI2 FERIMORO ERS034-2019    29082020
Public Function BuscaTpoDoc(ByVal psPersCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As New COMConecta.DCOMConecta
    oConec.AbreConexion
    lsSQL = "exec stp_cns_ERS034_2019_mostrarTpoDoc '" & psPersCod & "'"
    Set BuscaTpoDoc = coConex.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'ADD BY PTI1 20210927
Public Function ValidaPerSegSepelioA(ByVal sCuenta As String, ByVal fsPersCod As String) As Boolean 'ADD BY PTI1 20210927
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    On Error GoTo ErrorGetVerificaPerSegSepelio
    
        sSql = "Exec sp_ValidaPerSegSepelioA '" & sCuenta & "','" & fsPersCod & "'"
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set rs = oConecta.CargaRecordSet(sSql)
        If Not rs.EOF And Not rs.BOF Then
            If rs!Resultado = 1 Then
                ValidaPerSegSepelioA = False
            Else
                ValidaPerSegSepelioA = True
            End If
        End If
        oConecta.CierraConexion
        Set oConecta = Nothing
                
        Exit Function
    
ErrorGetVerificaPerSegSepelio:
    Err.Raise Err.Number, "Verifica Cuenta Cliente", "Error al verificar la Cuenta del Cliente, contactarse con el area de TI"

End Function

