Attribute VB_Name = "gFunLogistica2"
Option Explicit

'COLORES
Global Const Blanco = "&H80000005"
Global Const Plomo = "&H8000000F"
Global Const Gris = "&H80000010"
Global Const Negro = "&H80000012"
Global Const Granate = "&H00000080"
Global Const Rojo = "&H001905DC"
Global Const Azul = "&H8000000D"
Global Const Azulino = "&H00FF0000"
Global Const Verde = "&H80000001"
Global Const Amarillo = "&H00C0FFFF"

'9020    9020    ESTADOS - CTRL VEHICULAR
'9021    9021    CLASE DE VEHICULOS
'9022    9022    MARCAS DE VEHICULOS
'9023    9023    COLORES - CTRL VEHICULAR
'9024    9024    TIPO DE REGISTRO - CTRL VEHICULAR
'9025    9025    INCIDENCIAS
'9026    9026    TIPO DE VEHICULO
'9027    9027    TIPO DE COMBUSTIBLE
'9028    9028    ESTADOS DE VEHICULOS
'9029    9029    ESTADOS DE CONDUCTORES

Global Const gcAsignacionEstado = 9020
Global Const gcClaseVehiculo = 9021
Global Const gcMarcaVehiculo = 9022
Global Const gcColorVehiculo = 9023
Global Const gcTipoRegistro = 9024
Global Const gcIncidencias = 9025
Global Const gcTipoVehiculo = 9026
Global Const gcTipoCombustible = 9027

Global Const gMonedaNacional = 1

Global Const gcAnulado = 0
Global Const gcDisponible = 1
Global Const gcSolicitud = 2
Global Const gcAprobado = 3
Global Const gcAceptado = 4
Global Const gcVistoBueno = 5
'-------------------------------------------------
Global Const gcActivo = 1
Global Const gcObservado = 2

'9041    9041    ETAPAS DE PROCESOS DE SELECCION
'9042    9042    PROCESOS DE SELECCION PLAN ANUAL A.C.
'9043    9043    AREAS DE APROBACION
'9044    9044    OBJETO DE ADQUISICION
'9045    9045    TIPOS DE PROCESOS DE SELECCION- PLAN ANUAL
'9046    9046    FUENTES DE FINANCIAMIENTO PLAN ANUAL
'Global Const gcEtapasProcesoSel = 9041
Global Const gcProcesoSeleccion = 9042
Global Const gcEstadosRPA = 9043
Global Const gcObjAdquisicion = 9044
Global Const gcTipoProcesoSeleccion = 9045
Global Const gcFuenteFinanciamiento = 9046
Global Const gcResponsableProceso = 9047
Global mMes(1 To 12) As String

Global cLogNro As String
Dim nIndex As Integer

Public Function GetLogMovNro() As String
GetLogMovNro = ""
GetLogMovNro = Format(Date, "YYYYMMDD") + Format(Time, "HHMMSS") + gsCodUser
End Function

'*******************************************************************
'Selecciona un texto dentro de un TEXTBOX
Public Sub SelTexto(vCajaTexto As Control)
vCajaTexto.SelStart = 0
vCajaTexto.SelLength = Len(Trim(vCajaTexto))
End Sub

Public Function GetBSUnidadLog(psBSCod As String) As String
Dim oConn As New DConecta, sSQL As String, rs As New ADODB.Recordset
GetBSUnidadLog = ""
If oConn.AbreConexion Then
   sSQL = "select t.cConsDescripcion as cUnidad from LogProSelBienesServicios b inner join " & _
          " (select nConsValor as nBSUnidad, cConsDescripcion from Constante where nConsCod = 9097) t " & _
          " on b.nBSUnidad = t.nBSUnidad where b.cProSelBSCod = '" & psBSCod & "'"
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      GetBSUnidadLog = rs!cUnidad
   End If
End If
End Function

'Public Function GetCargoCodAprobacion(ByRef vRHCargoReq As String, ByRef vNivelReq As Integer, ByRef vRHAreaCodAprobacion As String, ByRef vTodasAge As Integer) As String
'Dim oConn As New DConecta, sSQL As String, rs as New ADODB.Recordset
'
'GetCargoCodAprobacion = ""
'sSQL = "select n.cRHCargoCodAprobacion,n.nTodasAge, cRHAreaCod = coalesce(x.cAreaCod,'') " & _
'       " from NivelesAprobacion n left outer join AreaCargo x on n.cRHCargoCodAprobacion = x.cRHCargoCod" & _
'       " where n.cRHCargoCod = '" & vRHCargoReq & "' and n.nNivel=" & vNivelReq & " "
'
'If oConn.AbreConexion Then
'   Set rs = oConn.CargaRecordSet(sSQL)
'   If Not rs.EOF Then
'      GetCargoCodAprobacion = rs!cRHCargoCodAprobacion
'      vRHAreaCodAprobacion = rs!cRHAreaCod
'      vTodasAge = rs!nTodasAge
'   End If
'End If
'End Function

Public Function GetNivelesAprobacion(ByVal vRHCargoCod As String) As ADODB.Recordset
Dim oConn As New DConecta, sSQL As String

Set GetNivelesAprobacion = Nothing
If oConn.AbreConexion Then
   sSQL = "select * from LogNivelAprobacion where cRHCargoCod = '" & vRHCargoCod & "' order by nNivelAprobacion "
   Set GetNivelesAprobacion = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
End If
End Function

Public Function GetCargosAprobacion(ByVal vRHCargoCodAprobacion As String) As ADODB.Recordset
Dim oConn As New DConecta, rs As New ADODB.Recordset
Dim sSQL As String

If oConn.AbreConexion Then
   sSQL = "select cRHCargoCod,nNivelAprobacion,nAgencia from LogNivelAprobacion where cRHCargoCodAprobacion = '" & vRHCargoCodAprobacion & "'"
   Set GetCargosAprobacion = oConn.CargaRecordSet(sSQL)
End If
End Function


Public Function GetMinNivelAprobacion(ByVal vRHCargoCodAprobacion As String) As ADODB.Recordset
Dim oConn As New DConecta, rs As New ADODB.Recordset, sSQL As String
Dim rc As New ADODB.Recordset
Dim nNivelApro As Integer, nNivelProc As Integer, nConsolAge As Integer

nNivelApro = 0
nNivelProc = 0
nConsolAge = 0

rc.Fields.Append "cRHCargoCod", adVarChar, 10, adFldMayBeNull
rc.Fields.Append "nNivelProc", adInteger, 0, adFldMayBeNull
rc.Fields.Append "nNivelApro", adInteger, 0, adFldMayBeNull
rc.Fields.Append "nConsolAge", adInteger, 0, adFldMayBeNull
rc.Open

Set GetMinNivelAprobacion = Nothing

If oConn.AbreConexion Then
   sSQL = "select nNivelProceso from LogNivelProceso where cRHCargoCod = '" & vRHCargoCodAprobacion & "'"
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      nNivelProc = rs!nNivelProceso
   Else
      nNivelProc = 1
   End If
   
   sSQL = "select nMinNivel=coalesce(min(nNivelAprobacion),0) from LogNivelAprobacion where cRHCargoCodAprobacion = '" & vRHCargoCodAprobacion & "'"
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      nNivelApro = rs!nMinNivel
   End If
   
   sSQL = "select cRHCargoCod, nConsolAge from LogNivelAprobacion where cRHCargoCodAprobacion = '" & vRHCargoCodAprobacion & "' and nNivelAprobacion=" & nNivelApro & " "
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         rc.AddNew
         rc.Fields(0) = rs!cRHCargoCod
         rc.Fields(1) = nNivelApro
         rc.Fields(2) = nNivelProc
         rc.Fields(3) = rs!nConsolAge
         rc.Update
         rs.MoveNext
      Loop
   End If
End If
Set GetMinNivelAprobacion = rc
End Function

Public Function GetCargoAprobacion(ByVal vRHCargoCod As String, ByRef vRHCargoDescripcion As String, ByVal vNivel As Integer, vSector As Integer) As String
Dim sSQL As String, oConn As New DConecta, rs As New ADODB.Recordset
Dim cConsulta As String

cConsulta = ""
vRHCargoDescripcion = ""
GetCargoAprobacion = ""

If vNivel < 3 Then
   cConsulta = "       AND x.nSector in (0," & vSector & ")"
End If

sSQL = "select x.cRHCargoCodAprobacion, t.cRHCargoDescripcion " & _
       "  from LogPlanAnualCargosOpe x inner join RHCargosTabla t on x.cRHCargoCodAprobacion = t.cRHCargoCod" & _
       " where x.cRHCargoCod = '" & vRHCargoCod & "' and x.nNivelAprobacion = " & vNivel & " " + cConsulta
       
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      GetCargoAprobacion = rs!cRHCargoCodAprobacion
      vRHCargoDescripcion = rs!cRHCargoDescripcion
   End If
End If
End Function

Public Function GetSectorCargo(vRHCargoCod As String) As Integer
Dim sSQL As String, oConn As New DConecta, rs As New ADODB.Recordset

GetSectorCargo = 0
sSQL = "select distinct nSector from LogPlanAnualCargosOpe where cRHCargoCod = '" & vRHCargoCod & "'"
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      GetSectorCargo = rs!nSector
   End If
End If
End Function

Public Function GetUbigeoConsucode(ByVal psUbigeoCod As String) As String
Dim sSQL1 As String, sSQL2 As String, sSQL3 As String
Dim oConn As New DConecta, rs As New ADODB.Recordset
Dim cUbigeo2 As String, cUbigeo4 As String, cUbigeo6 As String

GetUbigeoConsucode = ""

cUbigeo2 = Left(psUbigeoCod, 2)
cUbigeo4 = Left(psUbigeoCod, 4)
cUbigeo6 = psUbigeoCod

sSQL1 = "Select cUbigeoDescripcion from LogProSelUbigeo where cUbigeoCod = '" & cUbigeo2 & "'"
sSQL2 = "Select cUbigeoDescripcion from LogProSelUbigeo where cUbigeoCod = '" & cUbigeo4 & "'"
sSQL3 = "Select cUbigeoDescripcion from LogProSelUbigeo where cUbigeoCod = '" & cUbigeo6 & "'"

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL1)
   If Not rs.EOF Then
      GetUbigeoConsucode = GetUbigeoConsucode + rs!cUbiGeoDescripcion + " / "
   End If
   Set rs = oConn.CargaRecordSet(sSQL2)
   If Not rs.EOF Then
      GetUbigeoConsucode = GetUbigeoConsucode + rs!cUbiGeoDescripcion + " / "
   End If
   Set rs = oConn.CargaRecordSet(sSQL1)
   If Not rs.EOF Then
      GetUbigeoConsucode = GetUbigeoConsucode + rs!cUbiGeoDescripcion
   End If
End If
End Function

Public Function UltimaSecuenciaIdentidad(ByVal psTablaSQL As String) As Integer
Dim oConn As New DConecta, rs As New ADODB.Recordset

UltimaSecuenciaIdentidad = 0
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet("SELECT nUltSec=IDENT_CURRENT('" & psTablaSQL & "')")
   If Not rs.EOF Then
      UltimaSecuenciaIdentidad = rs!nUltSec
   End If
Else
   MsgBox "No se puede establecer conexión..." + Space(10), vbInformation
   UltimaSecuenciaIdentidad = 0
End If
End Function
Public Function GetPrecioUnitario(ByVal pnAnio As Integer, ByVal pnMes As Integer, ByVal psBSCod As String) As Currency
Dim oConn As New DConecta, sSQL As String, rs As New ADODB.Recordset, cAnioMes As String
GetPrecioUnitario = 0

cAnioMes = CStr(pnAnio) + Format(pnMes, "00")

sSQL = "select Fecha=left(m.cMovNro,8), c.cBSCod, nPrecioUnit=sum(nBSValor)/count(c.cBSCod) " & _
       "  from BSControl c inner join Mov m on c.nMovNro = m.nMovNro " & _
       " where m.cMovNro like '" & cAnioMes & "%'  and c.cBSCod = '" & psBSCod & "' " & _
       " group by c.cBSCod,left(m.cMovNro,8) " & _
       " order by left(m.cMovNro,8) desc "
       
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      'Do While rs.EOF
         GetPrecioUnitario = rs!nPrecioUnit
      '   rs.MoveNext
      'Loop
   End If
End If
End Function

Public Function DeterminaProcesoSeleccion(ByVal pnObjetoCod As Integer, ByVal pnMonto As Currency, ByRef pnProSelTpoCod As Integer, ByRef pnProSelSubTpo As Integer) As String
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset
Dim sSQL As String

sSQL = ""
pnProSelTpoCod = 0
pnProSelSubTpo = 0
DeterminaProcesoSeleccion = ""
Select Case pnObjetoCod
    'BIENES
    Case 1
         sSQL = "select r.nProSelTpoCod,r.nProSelSubTpo,t.cProSelTpoDescripcion,r.cProSelSubTpo,r.cAbreviatura " & _
                " from LogProSelTpoRangos r inner join LogProSelTpo t on r.nProSelTpoCod = t.nProSelTpoCod " & _
                "where " & pnMonto & " > nBienesMin  and  " & pnMonto & " < nBienesMax and " & _
                "       nBienesMin>0 and nBienesMax>0"
    'SERVICIOS
    Case 2
         sSQL = "select r.nProSelTpoCod,r.nProSelSubTpo,t.cProSelTpoDescripcion,r.cProSelSubTpo,r.cAbreviatura " & _
                " from LogProSelTpoRangos r inner join LogProSelTpo t on r.nProSelTpoCod = t.nProSelTpoCod " & _
                " where " & pnMonto & " > nServiMin  and  " & pnMonto & " < nServiMax and " & _
                "      nServiMin>0 and nServiMax>0 "
    Case 3
         sSQL = ""
End Select

If Len(sSQL) = 0 Then Exit Function
   
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      'GetProcesoSeleccion = rs!nProSelTpoCod
      'vDescripcion = rs!cProSelTpoDescripcion
      'vAbreviatura = rs!cAbreviatura
      pnProSelTpoCod = rs!nProSelTpoCod
      pnProSelSubTpo = rs!nProSelSubTpo
      DeterminaProcesoSeleccion = rs!cAbreviatura
   End If
End If
End Function

Public Function ObtenerProcesoSeleccion(ByVal pnObjetoCod As Integer, ByVal pnMonto As Currency, ByRef psProcesoSeleccion As String, ByRef psAbreviatura As String) As Boolean
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset
Dim sSQL As String

sSQL = ""
psProcesoSeleccion = ""
psAbreviatura = ""

ObtenerProcesoSeleccion = False
Select Case pnObjetoCod
    'BIENES
    Case 1
         sSQL = "select r.nProSelTpoCod,r.nProSelSubTpo,t.cProSelTpoDescripcion,r.cProSelSubTpo,r.cAbreviatura " & _
                " from LogProSelTpoRangos r inner join LogProSelTpo t on r.nProSelTpoCod = t.nProSelTpoCod " & _
                "where " & pnMonto & " > nBienesMin  and  " & pnMonto & " < nBienesMax and " & _
                "       nBienesMin>0 and nBienesMax>0"
    'SERVICIOS
    Case 2
         sSQL = "select r.nProSelTpoCod,r.nProSelSubTpo,t.cProSelTpoDescripcion,r.cProSelSubTpo,r.cAbreviatura " & _
                " from LogProSelTpoRangos r inner join LogProSelTpo t on r.nProSelTpoCod = t.nProSelTpoCod " & _
                " where " & pnMonto & " > nServiMin  and  " & pnMonto & " < nServiMax and " & _
                "      nServiMin>0 and nServiMax>0 "
    Case 3
         sSQL = ""
End Select

If Len(sSQL) = 0 Then Exit Function
   
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      psProcesoSeleccion = rs!cProSelTpoDescripcion
      psAbreviatura = rs!cAbreviatura
      ObtenerProcesoSeleccion = True
   End If
End If
End Function

Public Function RequerimientosNoAprobados(ByVal pnAnio As Integer) As Integer
Dim oConn As New DConecta, sSQL As String, rs As New ADODB.Recordset

RequerimientosNoAprobados = 0

sSQL = "select nPlanReq=count(*) from " & _
       " (select r.nPlanReqNro,nNro=count(*) from LogPlanAnualAprobacion a " & _
       "  inner join LogPlanAnualReq r on a.nPlanReqNro = r.nPlanReqNro " & _
       "  Where r.nAnio = " & pnAnio & " and r.nEstado=1 group by r.nPlanReqNro) x left join " & _
       " (select nPlanReqNro,nApro=count(*) from LogPlanAnualAprobacion where nEstadoAprobacion=1 group by nPlanReqNro) y on x.nPlanReqNro = y.nPlanReqNro " & _
       "  Where x.nNro <> Y.nApro  "

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      RequerimientosNoAprobados = rs!nPlanReq ' Rs!nProSelReq
   End If
End If
End Function
