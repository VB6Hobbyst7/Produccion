VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dPITFunciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function nRegistraConciliacion(pdFecha As Date, pdFechaLogIB As Date, lsUser As String) As Long
Dim loConec As New DConecta
Dim lRs As ADODB.Recordset
Dim lsSQL As String

    lsSQL = " exec PIT_stp_ins_Conciliacion '" & Format(pdFecha, "YYYY-MM-DD") & "','" & Format(pdFechaLogIB, "YYYY-MM-DD") & "','" & lsUser & "'"
    
    loConec.AbreConexion
    
    Set lRs = loConec.Ejecutar(lsSQL)
    
    nRegistraConciliacion = lRs("nConciliacionId")
    
    lsSQL = "exec PIT_stp_ConciliacionOperacionesIntercajas " & nRegistraConciliacion & ",'" & Format(pdFechaLogIB, "YYYY-MM-DD") & "'"
    
    loConec.Ejecutar (lsSQL)
    
    loConec.CierraConexion
    
    Set loConec = Nothing
    
End Function


Public Function recuperaConciliacionPorFechaProc(pdFechaProc As Date) As ADODB.Recordset
Dim loConec As New DConecta
Dim lsSQL As String

    lsSQL = " exec PIT_stp_sel_ConciliacionPorFechaProc '" & Format(pdFechaProc, "YYYY-MM-DD") & "'"
    
    loConec.AbreConexion
    
    Set recuperaConciliacionPorFechaProc = loConec.CargaRecordSet(lsSQL)
        
    loConec.CierraConexion
    
    Set loConec = Nothing

End Function

Public Function nRegistraLogOpeIB(pdFecha As Date, pdFechaLogIB As Date, lsUser As String) As Long
Dim loConec As New DConecta
Dim lRs As ADODB.Recordset
Dim lsSQL As String

    lsSQL = " exec PIT_stp_ins_LogOpeIB '" & Format(pdFecha, "YYYY-MM-DD") & "','" & Format(pdFechaLogIB, "YYYY-MM-DD") & "','" & lsUser & "'"
    
    loConec.AbreConexion
    
    Set lRs = loConec.Ejecutar(lsSQL)
    
    nRegistraLogOpeIB = lRs("nLogOpeId")
    
    loConec.CierraConexion
    
    Set loConec = Nothing
    
End Function

Public Function recuperaLogOpeIBPorFechaLog(pdFechaLogIB As Date) As ADODB.Recordset
Dim loConec As New DConecta
Dim lsSQL As String

    lsSQL = " exec PIT_stp_sel_LogOpeIBPorFechaLog '" & Format(pdFechaLogIB, "YYYY-MM-DD") & "'"
    
    loConec.AbreConexion
    
    Set recuperaLogOpeIBPorFechaLog = loConec.CargaRecordSet(lsSQL)
        
    loConec.CierraConexion
    
    Set loConec = Nothing

End Function

Public Function recuperaBusquedaPersona(pnTipoBusq As Integer, psDatoBusq As String) As ADODB.Recordset
Dim loConec As New DConecta
Dim lsSQL As String
    
    lsSQL = " exec PIT_stp_sel_BuscaCliente " & pnTipoBusq & ",'" & psDatoBusq & "'"
    
    loConec.AbreConexion
    
    Set recuperaBusquedaPersona = loConec.CargaRecordSet(lsSQL)
        
    loConec.CierraConexion
    
    Set loConec = Nothing
   
End Function

Public Function recuperaCMACSParaCombo() As ADODB.Recordset
Dim loConec As New DConecta
Dim lsSQL As String
    
    lsSQL = "exec PIT_stp_sel_CMACS"
    
    loConec.AbreConexion
    
    Set recuperaCMACSParaCombo = loConec.CargaRecordSet(lsSQL)
    
    loConec.CierraConexion
    
    Set loConec = Nothing
       
End Function
