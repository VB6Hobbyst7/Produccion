VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************Automatizacion del Formato de Revision
'***************************Validacion de la calificacion

Dim oCon As COMConecta.DCOMConecta

'Revision Calificacion
Public Function DarValidarCierre(ByVal sFCierre As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_sel_ValidarCierre '" & sFCierre & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set DarValidarCierre = lrDatos
    Set lrDatos = Nothing
End Function


Public Function CargarDatosCalificacionXCliente(ByVal psPersCod As String, ByVal psCodCta As String, ByVal pdFechaFinMes As Date, ByVal pnTipCamb As Double, ByRef psMensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ReporteCarteraCreditosXCliente '" & psPersCod & "', '" & psCodCta & "', '" & Format(pdFechaFinMes, "yyyymmdd") & "'," & pnTipCamb & ""
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos para Calificacion"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos para Calificacion"
        Exit Function
    End If
    Set CargarDatosCalificacionXCliente = lrDatos
    Set lrDatos = Nothing
End Function

'NR MAVM 20090815
Public Function CargarPosicionClienteSBS(ByVal psPersId As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        lsSQL = "exec stp_sel_DarPosicionClienteSBS '" & psPersId & "'"
        Set loReg = New COMConecta.DCOMConecta
            loReg.AbreConexion
        Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
        Set CargarPosicionClienteSBS = lrDatos
        Set lrDatos = Nothing
End Function

'NR MAVM 20090815
Public Function CargarEndeudamientoSBS(ByVal psCodEnd As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        lsSQL = "exec stp_sel_DarEndeudamientoSBS '" & psCodEnd & "'"
        Set loReg = New COMConecta.DCOMConecta
            loReg.AbreConexion
        Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
        Set CargarEndeudamientoSBS = lrDatos
        Set lrDatos = Nothing
End Function

Public Function DarDatosCliente(ByVal psPersCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_DatosCliente '" & psPersCod & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
'    If lrDatos Is Nothing Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
'    If lrDatos.BOF And lrDatos.EOF Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
    Set DarDatosCliente = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarGarantia(ByVal psCodCta As String, ByRef psMensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_GarantiaXCodCta '" & psCodCta & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos para Calificacion"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos para Calificacion"
        Exit Function
    End If
    Set CargarGarantia = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarCreditoXPersona(ByVal psPersCod As String, ByVal sFCierre As String, ByRef psMensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_AuditCreditoXPersona '" & psPersCod & "', '" & sFCierre & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    Set CargarCreditoXPersona = lrDatos
    Set lrDatos = Nothing
End Function

Public Sub RegistrarRevisionCalificacion(ByVal sCodPers As String, ByVal sCodCta As String, ByVal sFCierre As String, ByVal CuTCambio As Currency, ByVal sFRegistro As String, ByVal sCAnalista As String, ByVal sGiro As String, ByVal sFSDCMAC As String, ByVal cMontoCMAC As Currency, ByVal sTMonedaCMAC As String, ByVal sFSDSF As String, ByVal cMontoSF As Currency, ByVal sTMonedaSF As String, ByVal sPNorm As String, ByVal sPCPP As String, ByVal sPDefic As String, ByVal sPDud As String, ByVal sPPerd As String, ByVal sCalificacionCMAC As String, ByVal sCalificacionSF As String, ByVal sCalificacionOCI As String, ByVal sSituacion As String, ByVal sDesarrollo As String, ByVal sGarantia As String, ByVal sInformacion As String, ByVal sEvaluacion As String, ByVal sComentario As String, ByVal sConclusion As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_RevisionCalificacion '" & sCodPers & "', '" & sCodCta & "', '" & sFCierre & "', " & CuTCambio & ", '" & sFRegistro & "', '" & sCAnalista & "', '" & sGiro & "', '" & sFSDCMAC & "', '" & cMontoCMAC & "', '" & sTMonedaCMAC & "', '" & sFSDSF & "', '" & cMontoSF & "', '" & sTMonedaSF & "', '" & sPNorm & "', '" & sPCPP & "', '" & sPDefic & "', '" & sPDud & "', '" & sPPerd & "', '" & sCalificacionCMAC & "', '" & sCalificacionSF & "', '" & sCalificacionOCI & "', '" & sSituacion & "', '" & sDesarrollo & "', '" & sGarantia & "', '" & sInformacion & "', '" & sEvaluacion & "', '" & sComentario & "', '" & sConclusion & "', 1"
    '
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Sub DeleteRevisionCalificacion(ByVal iRevisionId As Integer)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_del_RevisionCalificacion " & iRevisionId & ""
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

'Public Function ListarRevisionCalificacion(ByRef obj As COMANEJADOR.ManejadorError, ByVal psPersCod As String, ByVal pdFechaFinMes As Date, ByRef psMensaje As String) As ADODB.Recordset  ', ByVal pnTipCamb As Double
'On Error GoTo Manejador
'
'        Dim lsSQL As String
'        Dim loReg As COMConecta.DCOMConecta
'        Dim lrDatos As ADODB.Recordset
'
'        lsSQL = "exec stp_sel_ListarRevisionCalificacion '" & psPersCod & "', '" & pdFechaFinMes & "'"  ', " & pnTipCamb & ""
'        Set loReg = New COMConecta.DCOMConecta
'        loReg.AbreConexion
'        Set lrDatos = loReg.CargaControladorRecordSet(obj, lsSQL, adLockReadOnly)
'        If obj.ValidarError Then GoTo Manejador
'        If lrDatos Is Nothing Then
'            psMensaje = "No existen datos"
'            Exit Function
'        End If
'        If lrDatos.BOF And lrDatos.EOF Then
'            psMensaje = "No existen datos"
'            Exit Function
'        End If
'        Set ListarRevisionCalificacion = lrDatos
'        Set lrDatos = Nothing
'        Exit Function
'
'Manejador:
'    obj.ManejarError "ListarRevisionCalificacion", "COMDAuditoria/DCOMRevision", Err
'End Function

Public Function ListarRevisionCalificacion(ByVal psPersCod As String, ByVal pdFechaFinMes As Date, ByRef psMensaje As String) As ADODB.Recordset  ', ByVal pnTipCamb As Double
'On Error GoTo Manejador
        
        Dim lsSQL As String
        Dim loReg As COMConecta.DCOMConecta
        Dim lrDatos As ADODB.Recordset
        
        lsSQL = "exec stp_sel_ListarRevisionCalificacion '" & psPersCod & "', '" & pdFechaFinMes & "'"  ', " & pnTipCamb & ""
        Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
        Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
        'Dim obj As New COMANEJADOR.ManejadorError
        'If obj.ValidarError Then GoTo Manejador
        If lrDatos Is Nothing Then
            psMensaje = "No existen datos"
            Exit Function
        End If
        If lrDatos.BOF And lrDatos.EOF Then
            psMensaje = "No existen datos"
            Exit Function
        End If
        Set ListarRevisionCalificacion = lrDatos
        Set lrDatos = Nothing
        Exit Function

'Manejador:
    'obj.ManejarError "ListarRevisionCalificacion", "COMDAuditoria/DCOMRevision", Err
End Function

Public Function MaximoIdRevision() As Integer
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
Dim lIdRev As Integer
lCod = 0
    
    lsSQL = "exec stp_sel_MaximoIdRevision"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    
    If lrDatos.RecordCount > 0 Then
        lIdRev = lrDatos!valor
    Else
        lIdRev = lIdRev
    End If
    
    MaximoIdRevision = lIdRev
    Set loReg = Nothing
    Exit Function
    
End Function

Public Function DarRevisionXId(ByVal iRevisionId As Integer) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ListarRevisionXId '" & iRevisionId & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
'    If lrDatos Is Nothing Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
'    If lrDatos.BOF And lrDatos.EOF Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
    Set DarRevisionXId = lrDatos
    Set lrDatos = Nothing
End Function

Public Sub ModificarRevisionCalificacion(ByVal iRevisionId As Integer, ByVal sCodPers As String, ByVal sCodCta As String, ByVal sFCierre As String, ByVal CuTCambio As Currency, ByVal sFRegistro As String, ByVal sCAnalista As String, ByVal sGiro As String, ByVal sFSDCMAC As String, ByVal cMontoCMAC As Currency, ByVal sTMonedaCMAC As String, ByVal sFSDSF As String, ByVal cMontoSF As Currency, ByVal sTMonedaSF As String, ByVal sPNorm As String, ByVal sPCPP As String, ByVal sPDefic As String, ByVal sPDud As String, ByVal sPPerd As String, ByVal sCalificacionCMAC As String, ByVal sCalificacionSF As String, ByVal sCalificacionOCI As String, ByVal sSituacion As String, ByVal sDesarrollo As String, ByVal sGarantia As String, ByVal sInformacion As String, ByVal sEvaluacion As String, ByVal sComentario As String, ByVal sConclusion As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_upd_RevisionCalificacion " & iRevisionId & ", '" & sCodPers & "', '" & sCodCta & "', '" & sFCierre & "', " & CuTCambio & ", '" & sFRegistro & "', '" & sCAnalista & "', '" & sGiro & "', '" & sFSDCMAC & "', '" & cMontoCMAC & "', '" & sTMonedaCMAC & "', '" & sFSDSF & "', '" & cMontoSF & "', '" & sTMonedaSF & "', '" & sPNorm & "', '" & sPCPP & "', '" & sPDefic & "', '" & sPDud & "', '" & sPPerd & "', '" & sCalificacionCMAC & "', '" & sCalificacionSF & "', '" & sCalificacionOCI & "', '" & sSituacion & "', '" & sDesarrollo & "', '" & sGarantia & "', '" & sInformacion & "', '" & sEvaluacion & "', '" & sComentario & "', '" & sConclusion & "', 1"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Function ValidarRevision(ByVal psCodCta As String, ByVal pdFechaFinMes As String, ByRef psMensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ValidarRevision '" & psCodCta & "', '" & pdFechaFinMes & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos para Calificacion"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos para Calificacion"
        Exit Function
    End If
    Set ValidarRevision = lrDatos
    Set lrDatos = Nothing
End Function

'Pistas

Public Function CargarDatosPistas(ByVal sFI As String, ByVal sFF As String, ByRef psMensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_Pistas '" & sFI & "', '" & sFF & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    Set CargarDatosPistas = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarDatosPistasReporte(ByVal sFI As String, ByVal sFF As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_Pistas '" & sFI & "', '" & sFF & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarDatosPistasReporte = lrDatos
    Set lrDatos = Nothing
End Function

'*************** MAVM Comentarios de Reporte de Cartera de Creditos

Public Function CargarAnalista(ByVal sAgencia As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_sel_Analista '" & sAgencia & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarAnalista = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarInstituciones() As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_Instituciones "
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarInstituciones = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarDatosCalificacionComentario(ByVal pdFechaFinMes As Date, ByVal pnTipCamb As Double, ByVal cCodInst As String, ByVal sFI As String, ByVal sFF As String, ByVal sImporteI As String, ByVal sImporteF As String, ByVal sAgencia As String, ByVal sAnalista As String, ByVal sMoneda As String, ByVal sCodTipoCred As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ReporteCarteraCreditosComentario '" & Format(pdFechaFinMes, "yyyymmdd") & "'," & pnTipCamb & ", '" & cCodInst & "', '" & sFI & "', '" & sFF & "', '" & sImporteI & "', '" & sImporteF & "', '" & sAgencia & "', '" & sAnalista & "', '" & sMoneda & "', '" & sCodTipoCred & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarDatosCalificacionComentario = lrDatos
    Set lrDatos = Nothing
End Function

' Cartas de Circulacion - Ahorros

Public Function CargarDatosCtaAhorro(ByVal sNroCta As String, ByVal sPersCod As String, ByVal sAgencia As String, ByVal sCodProd As String, ByVal sMoneda As String, ByVal sImporteI As String, ByVal sImporteF As String, ByVal sTasa As String, ByVal sFI As String, ByVal sFF As String, ByRef psMensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ReporteCuentaAhorro '" & sNroCta & "', '" & sPersCod & "', '" & sAgencia & "', '" & sCodProd & "', '" & sMoneda & "', '" & sImporteI & "', '" & sImporteF & "', '" & sTasa & "', '" & sFI & "', '" & sFF & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    Set CargarDatosCtaAhorro = lrDatos
    Set lrDatos = Nothing
End Function

Public Function DarDatosXNroCta(ByVal NroCta As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ListarDatosXNroCta '" & NroCta & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
'    If lrDatos Is Nothing Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
'    If lrDatos.BOF And lrDatos.EOF Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
    Set DarDatosXNroCta = lrDatos
    Set lrDatos = Nothing
End Function

Public Function DarDatosFuncionario() As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_DatosFuncionario "
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
'    If lrDatos Is Nothing Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
'    If lrDatos.BOF And lrDatos.EOF Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
    Set DarDatosFuncionario = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarDatosCtaAhorroXCliente(ByVal psPersCod As String, ByRef psMensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ReporteCuentaAhorroXCliente '" & psPersCod & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    Set CargarDatosCtaAhorroXCliente = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarProductos() As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_CargarProductos "
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarProductos = lrDatos
    Set lrDatos = Nothing
End Function

' Carta de Circulacion Creditos

Public Function CargarUserXCod(ByVal sCodPers As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_DevolverUserXCod '" & sCodPers & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarUserXCod = lrDatos
    Set lrDatos = Nothing
End Function

'NR Se Agrego los parametros moneda y monto de desmbolso
Public Function CargarDatosCredito(ByVal sNroCta As String, ByVal sCodPer As String, ByVal sAgencia As String, ByVal sCodTipCre As String, ByVal sAnalista As String, ByVal sFI As String, ByVal sFF As String, ByVal sMoneda As String, ByVal sImporteI As String, ByVal sImporteF As String, ByRef psMensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ReporteCuentaCredito '" & sNroCta & "', '" & sCodPer & "', '" & sAgencia & "', '" & sCodTipCre & "', '" & sAnalista & "', '" & sFI & "', '" & sFF & "', '" & sMoneda & "', '" & sImporteI & "', '" & sImporteF & "' "
       
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    Set CargarDatosCredito = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarTipoCredito() As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_CargarTipoCredito "
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarTipoCredito = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarDatosCreditoXCliente(ByVal psPersCod As String, ByRef psMensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ReporteCuentaCreditoXCliente '" & psPersCod & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    Set CargarDatosCreditoXCliente = lrDatos
    Set lrDatos = Nothing
End Function

Public Function DarCreditoXNroCta(ByVal NroCta As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_DatosCreditoXNroCta '" & NroCta & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
'    If lrDatos Is Nothing Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
'    If lrDatos.BOF And lrDatos.EOF Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
    Set DarCreditoXNroCta = lrDatos
    Set lrDatos = Nothing
End Function

' Analisis de Cuentas
Public Function GetOperacionesUsuario(ByVal sUsuario As String, ByVal sProducto As String, Optional ByVal psMoneda As String = "") As Recordset
Dim rsVar As Recordset
Dim loReg As COMConecta.DCOMConecta
Dim sSQL As String
Dim lsFiltroMon As String
Dim i As Integer
Dim nCont As Integer
Dim sCodOpe As String
'sSQL = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O " _
'    & "WHERE O.cOpeCod LIKE '" & sProducto & "%' And O.cOpeVisible = '1'  " & lsFiltroMon _
'    & " Order by O.cOpeCod, O.nOpeNiv "
'Set rsVar = dbCmact.CargaRecordSet(sSQL)
'sCodOpe = ""
'nCont = -1
'Do While Not rsVar.EOF
'    For I = nCont + 1 To NroRegOpe - 1
'        If Trim(rsVar!cOpeCod) = Trim(MatOpe(I, 0)) Then
'            sCodOpe = sCodOpe & Trim(rsVar!cOpeCod) & "','"
'            nCont = I
'            Exit For
'        End If
'    Next I
'    rsVar.MoveNext
'Loop
'    If sCodOpe <> "" Then
'        sCodOpe = Mid(sCodOpe, 1, Len(sCodOpe) - 2)
'    End If
Set loReg = New COMConecta.DCOMConecta
loReg.AbreConexion
    
If sProducto = "76" Then
    
sSQL = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O "
'If sCodOpe <> "" Then
    sSQL = sSQL & " Where O.cOpeCod in ('760000','760001','760002','760004','760010','760100','760104','760112','760113','760181','760200','760201','760210','760250','760251','760300','760301','760302','760303','760600','760610','760900','760905','760906','761200','761201','761202','761203','761204','763200','763202','763400','763401','763402','763403','763404','763405','763500','763501','763503','763504','763505','763506','763507','763600','763601','763602','763603','763604','763605','763700','763701','763702','763703')"
Else
    If psMoneda = "2" Then
        sSQL = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O "
        sSQL = sSQL & " Where O.cOpeCod in ('461000','461001','461002','461010','461011','461012','461013','461020','461021','461022','461023','461040','461041','461042','461044','461045','461090','461091','461092','461093','461094','461095','461100','461110','461111','461112','461160','461200','461210','461220','461230','461400','461401','461402','461410','461411','461413','461416','461600','461601')"
    Else
        sSQL = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O "
        sSQL = sSQL & " Where O.cOpeCod in ('462000','462001','462002','462010','462011','462012','462013','462020','462021','462022','462023','462040','462041','462042','462044','462090','462091','462092','462094','462100','462110','462111','462112','462160','462200','462210','462220','462230','462400','462401','462402','462410','462411','462413','462600','462601')"
    End If

End If
'Else
    'sSQL = sSQL & " Where O.cOpeCod in ('')"
'End If
'Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
Set rsVar = loReg.CargaRecordSet(sSQL)
Set GetOperacionesUsuario = rsVar
Set rsVar = Nothing
End Function

'*** Sistemas: Reporte de Usuarios y Perfiles

Public Function CargarAgencias() As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_sel_Agencias"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarAgencias = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarAreas() As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_sel_Areas"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarAreas = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarDatosUsuarioXUser(ByVal psUser As String, ByVal psAgencia As String, ByVal psArea As String, ByRef psMensaje As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_sel_DatosUsuarioXUser '" & psUser & "', '" & psAgencia & "', '" & psArea & "'"
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    Set CargarDatosUsuarioXUser = lrDatos
    Set lrDatos = Nothing
End Function

Public Sub RegistrarUsuarioCMACMAYNASTem(ByVal sNombre As String, ByVal sUsuario As String, ByVal sAgencia As String, ByVal sArea As String, ByVal sGrupo As String, ByVal sOperaciones As String, ByVal sColocaciones As String, ByVal sOtros As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_UsuarioCMACMAYNASTem '" & sNombre & "', '" & sUsuario & "', '" & sAgencia & "', '" & sArea & "', '" & sGrupo & "', '" & sOperaciones & "', '" & sColocaciones & "', '" & sOtros & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Sub EliminarUsuarioCMACMAYNASTem()
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_del_UsuarioCMACMAYNASTem "
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Function CargarUsuarioCMACMAYNASExcel() As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_sel_UsuarioCMACMAYNAS"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarUsuarioCMACMAYNASExcel = lrDatos
    Set lrDatos = Nothing
End Function

''Captaciones
'Public Function CargarOperacionesXGrupo(ByVal psGrupo As String, ByVal svar1 As String, ByVal svar2 As String, ByRef psMensaje As String) As ADODB.Recordset
'    Dim lsSQL As String
'    Dim loReg As COMConecta.DCOMConecta
'    Dim lrDatos As ADODB.Recordset
'    lsSQL = "exec stp_sel_OperacionesXGrupo '" & psGrupo & "', '" & svar1 & "', '" & svar2 & "'"
'    Set loReg = New COMConecta.DCOMConecta
'        loReg.AbreConexion
'    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
'    If lrDatos Is Nothing Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
'    If lrDatos.BOF And lrDatos.EOF Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
'    Set CargarOperacionesXGrupo = lrDatos
'    Set lrDatos = Nothing
'End Function
'
''Colocaciones
'Public Function CargarOperacionesXGrupoColocaciones(ByVal psGrupo As String, ByRef psMensaje As String) As ADODB.Recordset
'    Dim lsSQL As String
'    Dim loReg As COMConecta.DCOMConecta
'    Dim lrDatos As ADODB.Recordset
'    lsSQL = "exec stp_sel_OperacionesXGrupoColocaciones '" & psGrupo & "'"
'    Set loReg = New COMConecta.DCOMConecta
'        loReg.AbreConexion
'    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
'    If lrDatos Is Nothing Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
'    If lrDatos.BOF And lrDatos.EOF Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
'    Set CargarOperacionesXGrupoColocaciones = lrDatos
'    Set lrDatos = Nothing
'End Function
'
''Otros
'Public Function CargarOperacionesXGrupoOtros(ByVal psGrupo As String, ByRef psMensaje As String) As ADODB.Recordset
'    Dim lsSQL As String
'    Dim loReg As COMConecta.DCOMConecta
'    Dim lrDatos As ADODB.Recordset
'    lsSQL = "exec stp_sel_OperacionesXGrupoOtros '" & psGrupo & "'"
'    Set loReg = New COMConecta.DCOMConecta
'        loReg.AbreConexion
'    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
'    If lrDatos Is Nothing Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
'    If lrDatos.BOF And lrDatos.EOF Then
'        psMensaje = "No existen datos"
'        Exit Function
'    End If
'    Set CargarOperacionesXGrupoOtros = lrDatos
'    Set lrDatos = Nothing
'End Function

'CREDITOS DESEMBOLSADOS
Public Function CargarCreditosDesembolsados(ByVal sFI As String, ByVal sFF As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "stp_sel_Cancelados_x_ampliacion '" & sFI & "', '" & sFF & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarCreditosDesembolsados = lrDatos
    Set lrDatos = Nothing
End Function

''Cargar FechaCierre x Año_Mes
'Public Function FechaCierreXAnio_Mes(ByVal sFecha As String) As ADODB.Recordset
'    Dim lsSQL As String
'    Dim loReg As COMConecta.DCOMConecta
'    Dim lrDatos As ADODB.Recordset
'
'        lsSQL = "exec stp_sel_FechaCierreXAnioMes '" & sFecha & "'"
'
'        Set loReg = New COMConecta.DCOMConecta
'            loReg.AbreConexion
'        Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
'        Set FechaCierreXAnio_Mes = lrDatos
'        Set lrDatos = Nothing
'End Function

