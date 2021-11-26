VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogPlanAnual"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ActualizaBSGrupos(ByVal pnAnio As Integer)
Dim oConn As New DConecta, sSQL As String

sSQL = "UPDATE d set d.cBSGrupoCod = b.cBSGrupoCod " & _
       "  from LogPlanAnualReqDetalle d inner join LogProSelBienesServicios b on d.cBSCod = b.cProSelBSCod " & _
       " where d.nAnio = " & pnAnio & "   "
       
If oConn.AbreConexion Then
   oConn.Ejecutar sSQL
   oConn.CierraConexion
End If
End Sub

Public Function GrabaCabeceraPlanAnual(ByVal pnAnio As Integer, ByVal psEntidad As String, ByVal psRUC As String, _
                                       ByVal psSiglas As String, ByVal psEjecutora As String, ByVal psAprueba As String, ByVal psPliego As String) As Integer
                                       
Dim sSQL As String, oConn As New DConecta

GrabaCabeceraPlanAnual = 0
If oConn.AbreConexion Then

  'ANULA PLAN ANUAL ANTERIOR -------------------------------------------------------------
  'Anula Cabecera
  
   sSQL = "UPDATE LogPlanAnual SET nPlanAnualEstado=0 WHERE nPlanAnualAnio = " & pnAnio & " "
   oConn.Ejecutar sSQL
   
  'Anula Detalle
   sSQL = "UPDATE LogPlanAnualDetalle SET nPlanAnualEstado=0 WHERE nPlanAnualAnio = " & pnAnio & " "
   oConn.Ejecutar sSQL
   
  
  'GENERA NUEVO PLAN ANUAL -----------------------------------------------
  'CABECERA SOLAMENTE
   sSQL = "insert into LogPlanAnual (nPlanAnualAnio,cPlanAnualEntidad,cPlanAnualRUC," & _
          "                          cPlanAnualSiglas,cPlanAnualEjecutor,cPlanAnualAprueba, " & _
          "                          cPlanAnualPliego) " & _
          "                  Values (" & pnAnio & ",'" & psEntidad & "','" & psRUC & "'," & _
          "                           '" & psSiglas & "','" & psEjecutora & "','" & psAprueba & "', " & _
          "                           '" & psPliego & "')"
   oConn.Ejecutar sSQL
   
   GrabaCabeceraPlanAnual = UltimaSecuenciaIdentidad("LogPlanAnual")
   oConn.CierraConexion
End If
End Function

Public Function RequerimientoPlanAnual(ByVal pnPlanReqNro As Long, Optional ByVal pbSoloActivos As Boolean = False) As ADODB.Recordset
Dim sSQL As String, oConn As New DConecta

Set RequerimientoPlanAnual = Nothing
If oConn.AbreConexion Then
   If pbSoloActivos Then
      sSQL = "select p.cBSCod,cBSDescripcion=substring(b.cBSDescripcion,1,60),t.cUnidad, " & _
           "       p.nMes01,p.nMes02,p.nMes03,p.nMes04,p.nMes05,p.nMes06, " & _
           "       p.nMes07,p.nMes08,p.nMes09,p.nMes10,p.nMes11,p.nMes12, " & _
           "       nTotal = (p.nMes01 + p.nMes02 + p.nMes03 + p.nMes04 + p.nMes05 + p.nMes06 + p.nMes07 + p.nMes08 + p.nMes09 + p.nMes10 + p.nMes11 + p.nMes12) " & _
           "  from LogPlanAnualReqDetalle p inner join LogProSelBienesServicios b on p.cBSCod = b.cProSelBSCod " & _
           "       inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) t on b.nBSUnidad = t.nBSUnidad " & _
           " where p.nPlanReqNro = " & pnPlanReqNro & " and p.nEstado = 1 ORDER BY nItem"
   Else
      sSQL = "select p.cBSCod,cBSDescripcion=substring(b.cBSDescripcion,1,60),t.cUnidad,p.nEstado, " & _
           "       p.nMes01,p.nMes02,p.nMes03,p.nMes04,p.nMes05,p.nMes06, " & _
           "       p.nMes07,p.nMes08,p.nMes09,p.nMes10,p.nMes11,p.nMes12, " & _
           "       nTotal = (p.nMes01 + p.nMes02 + p.nMes03 + p.nMes04 + p.nMes05 + p.nMes06 + p.nMes07 + p.nMes08 + p.nMes09 + p.nMes10 + p.nMes11 + p.nMes12) " & _
           "  from LogPlanAnualReqDetalle p inner join LogProSelBienesServicios b on p.cBSCod = b.cProSelBSCod " & _
           "       inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) t on b.nBSUnidad = t.nBSUnidad " & _
           " where p.nPlanReqNro = " & pnPlanReqNro & "  ORDER BY nItem"
   End If
   Set RequerimientoPlanAnual = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
End If
End Function

Public Function EstadoAprobacionRequerimiento(ByVal pnAnio As Integer, ByVal psPersCod As String, Optional ByVal pnEstado As Integer = 1, Optional ByVal pnNivel As Integer = 1) As ADODB.Recordset
Dim sSQL As String, oConn As New DConecta

sSQL = "select r.nPlanReqNro,a.nEstadoAprobacion " & _
       "  from LogPlanAnualReq r inner join LogPlanAnualAprobacion a on r.nPlanReqNro = a.nPlanReqNro " & _
       " where   r.cPersCod = '" & psPersCod & "' and " & _
       "            r.nAnio =  " & pnAnio & "     and " & _
       "          r.nEstado =  " & pnEstado & "   and " & _
       " a.nNivelAprobacion =  " & pnNivel & " "
       
If oConn.AbreConexion Then
   Set EstadoAprobacionRequerimiento = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
End If
End Function

Public Function AreaCargoAgencia(ByVal psPersCod As String, ByVal psAnioMes As String) As ADODB.Recordset
Dim oConn As New DConecta, sSQL As String, rc As New ADODB.Recordset
Dim rs As New ADODB.Recordset

rc.Fields.Append "cRHCargoCod", adVarChar, 6, adFldMayBeNull
rc.Fields.Append "cRHAreaCod", adVarChar, 3, adFldMayBeNull
rc.Fields.Append "cRHAgeCod", adVarChar, 2, adFldMayBeNull
rc.Fields.Append "cRHCargo", adVarChar, 40, adFldMayBeNull
rc.Fields.Append "cRHArea", adVarChar, 40, adFldMayBeNull
rc.Fields.Append "cRHAgencia", adVarChar, 40, adFldMayBeNull
rc.Fields.Append "cPersona", adVarChar, 60, adFldMayBeNull
rc.Open

sSQL = "select cRHCargoCodOficial as cRHCargoCod,cRHAreaCodOficial as cRHAreaCod," & _
       "       cRHAgenciaCodOficial as cRHAgeCod,t1.cRHCargoDescripcion as cRHCargo, " & _
       "       cPersona=replace(p.cPersNombre,'/',' '), " & _
       "       t2.cAreaDescripcion as cRHArea, t3.cAgeDescripcion as cRHAgencia" & _
       "  from RHCargos c inner join RHCargosTabla t1 on c.cRHCargoCodOficial = t1.cRHCargoCod " & _
       "                  inner join Areas t2 on c.cRHAreaCodOficial = t2.cAreaCod " & _
       "                  inner join Agencias t3 on c.cRHAgenciaCodOficial = t3.cAgeCod " & _
       "                  inner join Persona p on c.cPersCod = p.cPersCod " & _
       " where c.cPersCod = '" & psPersCod & "' and " & _
       "       dRHCargoFecha = (select max(dRHCargoFecha) " & _
       "  from RHCargos where cPersCod  = '" & psPersCod & "' and " & _
       "                 dRHCargoFecha <= '" & psAnioMes & "') "

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      rc.AddNew
      rc.Fields(0) = rs!cRHCargoCod
      rc.Fields(1) = rs!cRHAreaCod
      rc.Fields(2) = rs!cRHAgeCod
      rc.Fields(3) = Mid(rs!cRHCargo, 1, 40)
      rc.Fields(4) = Mid(rs!cRHArea, 1, 40)
      rc.Fields(5) = Mid(rs!cRHAgencia, 1, 40)
      rc.Fields(6) = Mid(rs!cPersona, 1, 60)
      rc.Update
   End If
   Set AreaCargoAgencia = rc
End If
End Function

Public Function TodosBienesRequeridos(ByVal pnAnio As Integer, Optional ByVal pbHallarSinGrupo As Boolean = False) As ADODB.Recordset
Dim oConn As New DConecta, sSQL As String

If pbHallarSinGrupo Then
   sSQL = "select distinct b.cProSelBSCod, b.cBSDescripcion " & _
          "  from LogPlanAnualReq r inner join LogPlanAnualReqDetalle d on r.nPlanReqNro = d.nPlanReqNro " & _
          " inner join LogProSelBienesServicios b on d.cBSCod = b.cProSelBSCod " & _
          " Where d.nAnio = " & pnAnio & " And Len(d.cBSGrupoCod) < 4 And r.nEstado = 1 And d.nEstado = 1 "
Else
   sSQL = "select distinct b.cProSelBSCod, b.cBSDescripcion " & _
          "  from LogPlanAnualReq r inner join LogPlanAnualReqDetalle d on r.nPlanReqNro = d.nPlanReqNro " & _
          " inner join LogProSelBienesServicios b on d.cBSCod = b.cProSelBSCod " & _
          " Where d.nAnio = " & pnAnio & " And Len(d.cBSGrupoCod) = 4 And r.nEstado = 1 And d.nEstado = 1 "
End If

If oConn.AbreConexion Then
   Set TodosBienesRequeridos = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
End If
End Function

Public Function GruposPlanAnualDetalleBS() As ADODB.Recordset
Dim oConn As New DConecta, sSQL As String

sSQL = "select d.cBSGrupoCod,b.cCIIUCod,nObjetoCod = convert(tinyint,substring(b.cProSelBSCod,2,1)),g.cBSGrupoDescripcion,d.nMoneda,sum(d.nPrecioUnitario*d.nCantidad) as nMonto " & _
       "  from LogPlanAnualDetalleBS d inner join BSGrupos g on d.cBSGrupoCod = g.cBSGrupoCod " & _
       " inner join LogProSelBienesServicios b on d.cBSCod = b.cProSelBSCod " & _
       " group by d.cBSGrupoCod,b.cCIIUCod,convert(tinyint,substring(b.cProSelBSCod,2,1)),g.cBSGrupoDescripcion,d.nMoneda "
       
If oConn.AbreConexion Then
   Set GruposPlanAnualDetalleBS = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
End If
End Function
