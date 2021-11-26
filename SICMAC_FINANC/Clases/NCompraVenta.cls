VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCompraVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsBaseComunes As String
Dim vsBasePesonas As String
Dim lsSql        As String
Dim oConec  As DConecta

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As New ClasIni
vsBaseComunes = oIni.BaseComunes
vsBasePesonas = oIni.BasePersonas
Set oIni = Nothing
Set oConec = New DConecta
oConec.AbreConexion
End Sub

Private Sub Class_Terminate()
oConec.CierraConexion
Set oConec = Nothing
End Sub

Public Function GetImporteCompraVenta(psOpeCod, pdFecIni As Date, pdFecFin As Date, Optional psAgeCod As String = "") As Recordset
'    lsSql = "SELECT TotalDol = ISNULL(SUM(nMovImporte),0), TotalSol = ISNULL(SUM(ROUND(nMovImporte * mtc.nMovTpoCambio,2)),0), Numero = ISNULL(COUNT(*),0), TCPromedio = ISNULL(AVG(mtc.nMovTpoCambio),0), TCMinimo = ISNULL(MIN(mtc.nMovTpoCambio),0), TCMaximo = ISNULL(MAX(mtc.nMovTpoCambio),0) " _
'         & "FROM mov m join movcompraventa mcv on mcv.nmovnro = m.nmovnro " _
'         & "           join movtpocambio   mtc on mtc.nmovnro = m.nmovnro " _
'         & "WHERE m.nMovEstado = '" & gMovEstContabNoContable & "' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and copecod = '" & psOpeCod & "' and LEFT(cMovNro,8) BETWEEN '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "'"
         
    '*** PEAC 20120928
'    lsSql = "SELECT TotalDol = ISNULL(SUM(nMovImporte),0), TotalSol = ISNULL(SUM(ROUND(nMovImporte * mtc.nMovTpoCambio,2)),0), Numero = ISNULL(COUNT(*),0), TCPromedio = ISNULL(SUM(ROUND(nMovImporte * mtc.nMovTpoCambio,2)),0)/ISNULL(SUM(nMovImporte),1), TCMinimo = ISNULL(MIN(mtc.nMovTpoCambio),0), TCMaximo = ISNULL(MAX(mtc.nMovTpoCambio),0) " _
         & "FROM mov m join movcompraventa mcv on mcv.nmovnro = m.nmovnro " _
         & "           join movtpocambio   mtc on mtc.nmovnro = m.nmovnro " _
         & "WHERE m.nMovEstado = '" & gMovEstContabMovContable & "' and  m.nMovFlag=" & gMovFlagVigente & " and copecod = '" & psOpeCod & "' and LEFT(cMovNro,8) BETWEEN '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "'"
         
'    If psAgeCod <> "" Then
'        lsSql = lsSql & " And substring(cMovNro,18,2) = '" & psAgeCod & "'"
'    End If
    lsSql = "exec stp_sel_GetImporteCompraVenta '" & psOpeCod & "','" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & psAgeCod & "'"
    '*** FIN PEAC

    Set GetImporteCompraVenta = oConec.CargaRecordSet(lsSql)
End Function

'MIOL 20121022, SEGUN RQ12338 *********************************************************************************
Public Function GetImporteCompraVentaFinanzas(psOpeCod, pdFecIni As Date, pdFecFin As Date, Optional psAgeCod As String = "") As Recordset
    lsSql = "exec stp_sel_GetImporteCompraVentaFinanzas '" & psOpeCod & "','" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & psAgeCod & "'"
    Set GetImporteCompraVentaFinanzas = oConec.CargaRecordSet(lsSql)
End Function

Public Function GetImporteCompraVentaMasFinanzas(psOpeCod, pdFecIni As Date, pdFecFin As Date, Optional psAgeCod As String = "", Optional psOpeCodFI As String) As Recordset
    lsSql = "exec stp_sel_GetImporteCompraVentaInFinanzas '" & psOpeCod & "','" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & psAgeCod & "','" & psOpeCodFI & "'"
    Set GetImporteCompraVentaMasFinanzas = oConec.CargaRecordSet(lsSql)
End Function
'END MIOL ******************************************************************************************************
'ALPA 20140226**************************************************************************************************
Public Function ObtenerTipoCambioDiario(ByVal pdFecCamb As Date) As ADODB.Recordset
    Dim sql As String
    
    sql = "exec stp_sel_obtener_TipoCambio '" & Format(pdFecCamb, "YYYY/MM/DD") & "'"
    Set ObtenerTipoCambioDiario = oConec.CargaRecordSet(sql)
End Function
Public Function ObtenerReporteGanaciaPerdidaSpot(ByVal psFechaInicio As String, ByVal psFechaFin As String) As ADODB.Recordset
    Dim sql As String
    
    sql = "exec stp_sel_ReporteGanaciaPerdidaSpot '" & psFechaInicio & "','" & psFechaFin & "'"
    Set ObtenerReporteGanaciaPerdidaSpot = oConec.CargaRecordSet(sql)
End Function
'***************************************************************************************************************
