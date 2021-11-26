VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub RegistrarSesionDir(ByVal sTipoSesion, ByVal sNroSesion As String, ByVal sFecha As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_SesionDir '" & sTipoSesion & "', '" & sNroSesion & "', '" & sFecha & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Sub ModificarSesionDir(ByVal iSesionDirId As Integer, ByVal sTipoSesion, ByVal sNroSesion As String, ByVal sFecha As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_upd_SesionDir " & iSesionDirId & ", '" & sTipoSesion & "', '" & sNroSesion & "', '" & sFecha & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Sub RegistrarAcuerdoDir(ByVal iSesionDirId As Integer, ByVal sAsunto As String, ByVal sAcuerdo As String, ByVal sDetalle As String, ByVal sSituacion As String, ByVal tSituacion As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_AcuerdoDir " & iSesionDirId & ", '" & sAsunto & "', '" & sAcuerdo & "', '" & sDetalle & "', '" & sSituacion & "', '" & tSituacion & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Sub ModificarAcuerdoDir(ByVal iSesionDirId As Integer, ByVal sNroAcuerdo As String, ByVal sAsunto As String, ByVal sAcuerdo As String, ByVal sDetalle As String, ByVal sSituacion As String, ByVal tSituacion As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_upd_AcuerdoDir " & iSesionDirId & ", '" & sNroAcuerdo & "', '" & sAsunto & "', '" & sAcuerdo & "', '" & sDetalle & "', '" & sSituacion & "', '" & tSituacion & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Function MaximoIdSesionDir() As Integer
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
Dim lIdRev As Integer
    
    lsSQL = "exec stp_sel_MaximoIdSesionDir"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    
    If lrDatos.RecordCount > 0 Then
        lIdRev = lrDatos!valor
    Else
        lIdRev = lIdRev
    End If
    
    MaximoIdSesionDir = lIdRev
    Set loReg = Nothing
    Exit Function
End Function

Public Function ListarAcuerdoDir(ByVal sDesde As String, ByVal sHasta As String, ByRef psMensaje As String) As ADODB.Recordset
        Dim lsSQL As String
        Dim loReg As COMConecta.DCOMConecta
        Dim lrDatos As ADODB.Recordset
        
        lsSQL = "exec stp_sel_ListarAcuerdoDir '" & Format(sDesde, "yyyymmdd") & "', '" & Format(sHasta, "yyyymmdd") & "'"
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
        Set ListarAcuerdoDir = lrDatos
        Set lrDatos = Nothing
        Exit Function
End Function

Public Function DarSesionDirXId(ByVal iSesionDirId As Integer, ByVal sNroAcuerdo As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_ListarSesionDirXId '" & iSesionDirId & "', '" & sNroAcuerdo & "'"
       
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set DarSesionDirXId = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarReporteAcuerdoDir(ByVal sDesde As String, ByVal sHasta As String, ByVal sSituacion As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_ReporteAcuerdoDir '" & Format(sDesde, "yyyymmdd") & "', '" & Format(sHasta, "yyyymmdd") & "', '" & sSituacion & "'"
        
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarReporteAcuerdoDir = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarReporteEstadistico() As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_ReporteEstadistico"
        
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarReporteEstadistico = lrDatos
    Set lrDatos = Nothing
End Function


'Medidas Correctivas

Public Function CargarTipoMedida() As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_CargarTipoMedida "
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarTipoMedida = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarEnte() As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_CargarEnte "
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarEnte = lrDatos
    Set lrDatos = Nothing
End Function

Public Sub RegistrarInformeMed(ByVal sNroInforme As String, ByVal sFecha As String, ByVal sTipo As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_InformeMed '" & sNroInforme & "', '" & sFecha & "', '" & sTipo & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Sub ModificarInformeMed(ByVal iInformeId As Integer, ByVal sTipoSesion, ByVal sNroSesion As String, ByVal sFecha As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_upd_InformeMed " & iInformeId & ", '" & sTipoSesion & "', '" & sNroSesion & "', '" & sFecha & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Function MaximoIdInformeMed() As Integer
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
Dim lIdRev As Integer
    
    lsSQL = "exec stp_sel_MaximoIdInformeMed"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    
    If lrDatos.RecordCount > 0 Then
        lIdRev = lrDatos!valor
    Else
        lIdRev = lIdRev
    End If
    
    MaximoIdInformeMed = lIdRev
    Set loReg = Nothing
    Exit Function
End Function

Public Sub RegistrarMedidasCorrectivas(ByVal iInformeMedId As Integer, ByVal sObservacion As String, ByVal sRecomendacion As String, ByVal sAccion As String, ByVal sSituacion As String, ByVal tSituacion As String, ByVal sEnte As String, ByVal sArea As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_MedidasCorrectivas " & iInformeMedId & ", '" & sObservacion & "', '" & sRecomendacion & "', '" & sAccion & "', '" & sSituacion & "', '" & tSituacion & "', '" & sEnte & "', '" & sArea & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Sub ModificarMedidasCorrectivas(ByVal iInformeMedId As Integer, ByVal iMedidasCorrectivasId As Integer, ByVal sObservacion As String, ByVal sRecomendacion As String, ByVal sAccion As String, ByVal sSituacion As String, ByVal tSituacion As String, ByVal sEnte As String, ByVal sArea As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_upd_MedidasCorrectivas " & iInformeMedId & ", " & iMedidasCorrectivasId & ", '" & sObservacion & "', '" & sRecomendacion & "', '" & sAccion & "', '" & sSituacion & "', '" & tSituacion & "', '" & sEnte & "', '" & sArea & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Function ListarMedidasCorrectivas(ByRef psMensaje As String) As ADODB.Recordset 'ByVal sDesde As String, ByVal sHasta As String,
        Dim lsSQL As String
        Dim loReg As COMConecta.DCOMConecta
        Dim lrDatos As ADODB.Recordset
        
        lsSQL = "exec stp_sel_ListarMedidasCorrectivas " ''" & Format(sDesde, "yyyymmdd") & "', '" & Format(sHasta, "yyyymmdd") & "'"
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
        Set ListarMedidasCorrectivas = lrDatos
        Set lrDatos = Nothing
        Exit Function
End Function

Public Function DarInformeXId(ByVal iNroInformeId As Integer, ByVal iMedidasCorrectivasId As Integer) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_ListarInformeXId '" & iNroInformeId & "', " & iMedidasCorrectivasId & ""
       
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set DarInformeXId = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarReporteMedidasCorrectivas() As ADODB.Recordset 'ByVal sDesde As String, ByVal sHasta As String, ByVal sSituacion As String
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_ReporteMedidasCorrectivas " ''" & Format(sDesde, "yyyymmdd") & "', '" & Format(sHasta, "yyyymmdd") & "', '" & sSituacion & "'"
        
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarReporteMedidasCorrectivas = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarReporteMedidasCorrectivasEstadistico() As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_ReporteMedidasCorrectivasEstadistico"
        
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarReporteMedidasCorrectivasEstadistico = lrDatos
    Set lrDatos = Nothing
End Function










