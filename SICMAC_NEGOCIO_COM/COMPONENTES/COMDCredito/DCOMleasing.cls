VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMleasing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function RecuperaCalendarioLeasing(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
         
    sSql = "exec stp_sel_obtenerCalensdarioLeasing '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioLeasing = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Sub ActualizarLeasingCotizacion(ByVal psCtaCod As String, ByVal psPersCod As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
         
    sSql = "exec SAF_stp_upd_ActualizaCotizacion '" & psCtaCod & "','" & psPersCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub

Public Function ObtenerAprobacionLeasing(ByVal psCtaCod As String, Optional ByRef pnNroProveedores As Integer = 0) As ADODB.Recordset
Dim sSql As String
Dim oRs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
         
    Set oRs = New ADODB.Recordset
    sSql = "exec SAF_stp_sel_obtenerAprobacion '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    Set ObtenerAprobacionLeasing = oRs
    If oRs.EOF = True Or oRs.BOF = True Then
        pnNroProveedores = 0
        Exit Function
    Else
        pnNroProveedores = oRs.RecordCount
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function ObtenerComisionLeasingDesembolso(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oRs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
         
    Set oRs = New ADODB.Recordset
    sSql = "exec SAF_stp_sel_ComisionLeasingDesembolso '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    Set ObtenerComisionLeasingDesembolso = oRs
    
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaColocCalendarioLeasing(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, ByVal pnAplicado As ColocCalendApl, Optional ByVal pnNroCuota As Integer = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaColocCalendarioLeasing
'    sSql = "Select * from ColocCalendario where cCtacod = '" & psCtaCod & "' AND nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnAplicado
    sSql = "SELECT * "
    sSql = sSql & " From "
    sSql = sSql & " ("
    sSql = sSql & "     select Saf_numero_operacion_core cCtaCod,"
    sSql = sSql & "            2 nNroCalen,1 nColocCalendApl,"
    sSql = sSql & "            convert(int,case when Saf_numero_cuota Not IN ('OC')"
    sSql = sSql & "                           then Saf_numero_cuota"
    sSql = sSql & "                        Else"
    sSql = sSql & "                     (select MAX(Saf_numero_cuota)+1 from SAF..IntSaf_Activacion_Cronograma O where O.Saf_numero_operacion_saf= S.Saf_numero_operacion_saf and Saf_numero_cuota Not IN ('OC'))"
    sSql = sSql & "                  end) nCuota,"
    sSql = sSql & "             Saf_fecha_vencimiento dVenc, null dPago,"
    sSql = sSql & "             0 nColocCalendEstado,"
    sSql = sSql & "             'Calendario de Aprobacion Pagos' cDescripcion,"
    sSql = sSql & "             null cColocCalenFlag,"
    sSql = sSql & "             2 nCalendProc,"
    sSql = sSql & "             null cColocMiVivEval"
    sSql = sSql & "     from SAF..IntSaf_Activacion_Cronograma S"
    sSql = sSql & " ) S"
    sSql = sSql & " where S.cCtaCod= " & psCtaCod
    If pnNroCuota <> -1 Then
        sSql = sSql & " AND nCuota = " & pnNroCuota
    End If
    sSql = sSql & " ORDER BY nCuota"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaColocCalendarioLeasing = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocCalendarioLeasing:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
Public Function RecuperaColocCalendDetLeasing(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, ByVal pnAplicado As ColocCalendApl) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaColocCalendDetLeasing
    sSql = "exec SAF_stp_sel_ColocCalendDetLeasing '" & psCtaCod & "'," & pnNroCalen & "," & pnAplicado
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaColocCalendDetLeasing = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocCalendDetLeasing:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaCtasAhorroxPersonaLeasing(ByVal psPersCod As String, ByVal psMoneda As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaCtasAhorroxPersonaLeasing
    sSql = "exec SAF_stp_sel_CuentasAhorroPorCliente '" & psPersCod & "','" & psMoneda & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCtasAhorroxPersonaLeasing = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCtasAhorroxPersonaLeasing:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaDatosDesembolsoAbonoProveedores(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaDatosDesembolsoAbonoProveedores
    sSql = "exec SAF_stp_sel_DesembolsoProveedorLeasingAbono '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosDesembolsoAbonoProveedores = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosDesembolsoAbonoProveedores:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Sub ActualizarDesembolsoProveedor(ByVal psCtaCod As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
         
    sSql = "exec SAF_stp_upd_DesembolsoProveedorLeasingAbono '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub

Public Sub ActualizarCobranzaFechaPago(ByVal psCtaCod As String, ByVal pdFecPago As Date)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
         
    sSql = "exec SAF_stp_upd_INTSAF_COBRANZA_FECHAPAGO '" & psCtaCod & "','" & Format(pdFecPago, "YYYY/MM/DD") & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub

Public Function RecuperaDatosClienteLeasing(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaDatosClienteLeasing
    sSql = "exec stp_sel_ObtenerTitularCreditoLeasing '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosClienteLeasing = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosClienteLeasing:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function ObtenerComisionLeasingOtrasOperacionesAD(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oRs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
         
    Set oRs = New ADODB.Recordset
    sSql = "exec stp_ObtenerCuotaInicialLeasing '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    Set ObtenerComisionLeasingOtrasOperacionesAD = oRs
    
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function ValidaCreditoPersona(ByVal psCtaCod As String, ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oRs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
         
    Set oRs = New ADODB.Recordset
    sSql = "exec stp_sel_ValidarCreditoPersona '" & psCtaCod & "','" & psPersCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    Set ValidaCreditoPersona = oRs
    
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function ObtenerDatoText(ByVal psSql As String, ByVal psWheDato As String, ByVal psDato As String) As ADODB.Recordset
Dim sSql As String
Dim oRs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
         
    Set oRs = New ADODB.Recordset
    sSql = psSql & "Where " & psWheDato & "='" & psDato & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    Set ObtenerDatoText = oRs
    
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Sub ActualizarUsuarioSAFSICMACM(ByVal psPersCod As String, ByVal psCodSAF As String)
Dim sSql As String
Dim oRs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
         
    Set oRs = New ADODB.Recordset
    sSql = "Exec stp_ins_ActualizarUsuarioSAFSICMACM '" & psPersCod & "','" & psCodSAF & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub
'ALPA 20121228*******************************************************************
Public Function Obtener_MontoFinanciarLeasing(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oRs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
         
    Set oRs = New ADODB.Recordset
    sSql = "exec LE_stp_sel_MontoFinanciarLeasing '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    Set Obtener_MontoFinanciarLeasing = oRs
    
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'********************************************************************************
'ORCR INICIO 20140414 ***
Public Function ObtenerMontoFinanciado(ByVal cCtaCod As String) As ADODB.Recordset
On Error GoTo ErrorObtenerMontoFinanciado
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String

    sSql = "exec stp_sel_SIAFmontoFinanciado '" & cCtaCod & "'"

    oConecta.AbreConexion
    Set ObtenerMontoFinanciado = oConecta.CargaRecordSet(sSql, adLockReadOnly)
    oConecta.CierraConexion

    Set oConecta = Nothing
    
    Exit Function
ErrorObtenerMontoFinanciado:
    Err.Raise Err.Number, "DCOMleasing:ObtenerMontoFinanciado"", Err.Description"
End Function
'ORCR FIN 20140414 ***
