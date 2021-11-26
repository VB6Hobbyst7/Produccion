VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RecuperaCalendarioGastos(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
                    ByVal pnNroCuota As Integer, ByVal pnAplicado As ColocCalendApl, _
                    Optional ByVal pbTodos As Boolean = False, Optional pbExcluyeGastos As Boolean = False) As ADODB.Recordset

'CTI2 20181229 ADD pbExcluyeGastos

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta


    On Error GoTo ErrorRecuperaCalendarioGastos
    'MAVM 20120213 ***
    'sSql = "Select * from ColocCalendDet where cCtaCod = '"
    'sSql = "Select cCtaCod, nNroCalen, nColocCalendApl, nCuota, nPrdConceptoCod, nMonto, nMontoPagado, cFlag from ColocCalendDet where cCtaCod = '" & psCtaCod & _
    '        "' AND nNroCalen = " & pnNroCalen
    'JUEZ 20140730 Se modificó script *****************************
    sSql = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED " & _
           "Select cCtaCod, nNroCalen, nColocCalendApl, nCuota, ccd.nPrdConceptoCod, nMonto, nMontoPagado, cFlag " & _
           "From ColocCalendDet ccd " & _
           "Inner Join ProductoConcepto pc on ccd.nPrdConceptoCod=pc.nPrdConceptoCod " & _
           "Where cCtaCod = '" & psCtaCod & _
           "' AND nNroCalen = " & pnNroCalen
    '***
    'END JUEZ *****************************************************
    If Not pbTodos Then
        sSql = sSql & " AND nCuota = " & pnNroCuota
    End If
    'CTI2 20181229 INI ***
    If pbExcluyeGastos Then
        sSql = sSql & " AND isnull(pc.nAplicaConvenio,0) = 0"
    End If
    'CTI2 20181229 FIN ***
    sSql = sSql & " AND nColocCalendApl = " & pnAplicado
    'sSql = sSql & " AND nPrdConceptoCod like '12%' And (nMonto - nMontoPagado) > 0 Order By nCuota "
    sSql = sSql & " AND ccd.nPrdConceptoCod like '12%' And (nMonto - nMontoPagado) > 0 Order By nCuota,pc.cAplicaProceso desc,ccd.nPrdConceptoCod desc " 'JUEZ 20140404
    'MAVM 20120215 ***
    sSql = sSql & " SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
    '***
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioGastos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaCalendarioGastos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaCalendarioGastosPendientes(ByVal psCtaCod As String, _
                    ByVal pnAplicado As ColocCalendApl) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

'ARCV 26-06-2007
Dim oCons As COMDConstSistema.NCOMConstSistema
Dim dFecSisTmp As Date

Set oCons = New COMDConstSistema.NCOMConstSistema
dFecSisTmp = CDate(oCons.LeeConstSistema(gConstSistFechaSistema))
Set oCons = Nothing
'--------
    
    On Error GoTo ErrRecuperaCalendarioGastosPendientes
            
'ColocCred_ObtieneGastoFechaCuota](@psCtaCod char(18),@pnCuota integer,@pdFecha datetime)
'ARCV 26-06-2007
    sSql = " Select CD.* From ColocCalendario C Inner Join ColocCalendDet CD "
    sSql = sSql & " ON C.cCtaCod = CD.cCtaCod AND C.nNroCalen = CD.nNroCalen AND C.nColocCalendApl = CD.nColocCalendApl "
    sSql = sSql & " AND C.nCuota = CD.nCuota "
    sSql = sSql & " Where  C.cCtaCod = '" & psCtaCod & "' AND  C.nColocCalendApl = " & pnAplicado & " AND C.nColocCalendEstado = 0 "
    sSql = sSql & " AND CD.nPrdConceptoCod like '12%' AND C.nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = '" & psCtaCod & "') Order By C.nCuota   "
    
'    sSql = " Select CD.nCuota,CD.nPrdConceptoCod,nMonto=dbo.ColocCred_ObtieneGastoFechaCuota(CD.cCtaCod,CD.nCuota,'" & Format(dFecSisTmp, "mm/dd/yyyy") & "') " & _
'            " From ColocCalendario C Inner Join ColocCalendDet CD "
'    sSql = sSql & " ON C.cCtaCod = CD.cCtaCod AND C.nNroCalen = CD.nNroCalen AND C.nColocCalendApl = CD.nColocCalendApl "
'    sSql = sSql & " AND C.nCuota = CD.nCuota "
'    sSql = sSql & " Where  C.cCtaCod = '" & psCtaCod & "' AND  C.nColocCalendApl = " & pnAplicado & " AND C.nColocCalendEstado = 0 "
'    sSql = sSql & " AND nPrdConceptoCod LIKE '12%' AND C.nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = '" & psCtaCod & "') Order By C.nCuota   "
'-------------
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioGastosPendientes = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrRecuperaCalendarioGastosPendientes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaCalendarioGastosPendientesAFecha(ByVal psCtaCod As String, _
                    ByVal pnAplicado As ColocCalendApl, ByVal pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    
    sSql = " Select PC.cDescripcion, (CD.nMonto-CD.nMontoPagado) as nMonto "
    sSql = sSql & "  From ColocCalendario C Inner Join ColocCalendDet CD "
    sSql = sSql & " ON C.cCtaCod = CD.cCtaCod AND C.nNroCalen = CD.nNroCalen AND C.nColocCalendApl = CD.nColocCalendApl "
    sSql = sSql & " AND C.nCuota = CD.nCuota "
    sSql = sSql & " Inner Join ProductoConcepto PC ON PC.nPrdConceptoCod = CD.nPrdConceptoCod "
    sSql = sSql & " Where  C.cCtaCod = '" & psCtaCod & "' AND  C.nColocCalendApl = " & pnAplicado & " AND C.nColocCalendEstado = 0 "
    sSql = sSql & " AND CD.nPrdConceptoCod like '12%' AND C.nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = '" & psCtaCod & "') "
    sSql = sSql & " AND (C.nCuota = (Select nNroProxCuota From ColocacCred Where cCtaCod = C.cCtaCod) OR C.dVenc <= '" & Format(pdFecha, "mm/dd/yyyy") & "') "
    sSql = sSql & " Order By C.nCuota   "

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioGastosPendientesAFecha = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    
End Function

Public Function RecuperaCalendarioGastosPendientesAgrupado(ByVal psCuentas As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCalendarioGastosPendientesAgrupado
            
    sSql = " Select CD.nPrdConceptoCod, SUM(CD.nMonto - CD.nMontoPagado) as nMonto From ColocCalendario C Inner Join ColocCalendDet CD "
    sSql = sSql & " ON C.cCtaCod = CD.cCtaCod AND C.nNroCalen = CD.nNroCalen AND C.nColocCalendApl = CD.nColocCalendApl "
    sSql = sSql & " AND C.nCuota = CD.nCuota AND C.nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = C.cCtaCod ) "
    sSql = sSql & " Where  C.cCtaCod In  " & psCuentas & " AND  C.nColocCalendEstado = 0 "
    sSql = sSql & " AND CD.nPrdConceptoCod like '12%' AND C.nColocCalendApl = 1 "
    sSql = sSql & " Group By CD.nPrdConceptoCod "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioGastosPendientesAgrupado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaCalendarioGastosPendientesAgrupado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Function RecuperaCalendarioDesemb(ByVal psCtaCod As String, _
Optional ByVal pConn As COMDCredito.DCOMCredActBD = Nothing) As ADODB.Recordset

'pConn RIRO20200301

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    '**Comentado por DAOR 20080318 **************************************************************
    'sSql = "select C.dVenc,C.dPago, C.nCuota,C.nColocCalendEstado, "
    'sSql = sSql & " CASE WHEN C.nColocCalendEstado = 1 THEN 'DESEMBOLSADO' WHEN C.nColocCalendEstado = 0 THEN 'PENDIENTE' END AS cColocCalendEstado,"
    'sSql = sSql & " nCapital=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(1000,1010)),"
    'sSql = sSql & " nIntComp=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(1100,1107)),"
    'sSql = sSql & " nIntGracia=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1102),"
    'sSql = sSql & " nIntMor=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(1101,1108)),"
    'sSql = sSql & " nIntReprog=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1103),"
    'sSql = sSql & " nGasto=(select sum(nMonto) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where nPrdConceptoCod like '12%' )) "
    ''**PEAC 20071204, Capital segun ultimo calendario ************************
    'sSql = sSql & " ,nCapitalUltCal=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod='" & psCtaCod & "') And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(1000,1010))"
    ''By Capi 28022008
    'sSql = sSql & " ,cConvenioBN=IsNull(CCB.cCtaCod,'NO')"
    ''**************************************************************************
    'sSql = sSql & " from ColocCalendario C inner join ColocacCred CC on C.cCtaCod=CC.cCtaCod "
    ''By Capi 28022008
    'sSql = sSql & " Left join ColocConvBcoNac CCB on C.cCtaCod=CCB.cCtaCod "
    ''**DAOR 20071121, Tomar los desembolsos del calendario Original (Calendario 03)*********
    ''sSql = sSql & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl = " & gColocCalendAplDesembolso & " and C.nNroCalen = CC.nNroCalen order by C.nCuota"
    'sSql = sSql & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl = " & gColocCalendAplDesembolso
    'sSql = sSql & "      and C.nNroCalen = (case "
    'sSql = sSql & "          when (select nPrdEstado from Producto where cCtaCod='" & psCtaCod & "') <=2002 then (select nNroCalen from ColocacCred where cCtaCod='" & psCtaCod & "') "
    'sSql = sSql & "          else 3 end ) "
    'sSql = sSql & " order by C.nCuota"
    ''***************************************************************************************
    '****************************************************************************************************
    
    '**DAOR 20080318 ******************************************
    sSql = "exec RecuperaCalendarioDesembolso '" & psCtaCod & "'"
    '**********************************************************
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    'RIRO 20200302 **
    If Not pConn Is Nothing Then
        Set RecuperaCalendarioDesemb = pConn.CargaRecordSet(sSql)
    Else
        Set RecuperaCalendarioDesemb = oConecta.CargaRecordSet(sSql)
    End If
    'END RIRO *******
        
    'Set RecuperaCalendarioDesemb = oConecta.CargaRecordSet(sSql) RIRO 20200301 Comentado
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaCalendarioDesembPagados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    sSql = "select C.dVenc,C.nCuota,C.nColocCalendEstado, "
    sSql = sSql & " nCapital=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(1000,1010)),"
    sSql = sSql & " nIntComp=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(1100,1107)),"
    sSql = sSql & " nIntGracia=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1102),"
    sSql = sSql & " nIntMor=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(1101,1108)),"
    sSql = sSql & " nIntReprog=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1103),"
    sSql = sSql & " nGasto=(select sum(nMonto) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where nPrdConceptoCod like '12%' )) "
    sSql = sSql & " from ColocCalendario C inner join ColocacCred CC on C.cCtaCod=CC.cCtaCod "
    sSql = sSql & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl = " & gColocCalendAplDesembolso & " and C.nNroCalen = CC.nNroCalen AND C.nColocCalendEstado=1 order by C.nCuota"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set RecuperaCalendarioDesembPagados = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function


Public Function RecuperaCalendarioPagosDeuda(ByVal psCtaCod As String, Optional ByVal pnNroCalen As Integer = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    sSql = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
    sSql = sSql & " nCapital=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
    sSql = sSql & " nIntComp=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSql = sSql & " nIntGracia=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSql = sSql & " nIntMor=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSql = sSql & " nIntReprog=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSql = sSql & " nIntSuspenso=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
    sSql = sSql & " nGasto=(select sum(nMonto- nMontoPagado) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' ))"
    sSql = sSql & " from ColocCalendario C "
    If pnNroCalen = -1 Then
        sSql = sSql & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = C.cCtaCod) order by C.nCuota"
    Else
        sSql = sSql & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = " & pnNroCalen & " order by C.nCuota"
    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioPagosDeuda = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaCalendarioPagosRealizados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCalendarioPagosRealizados
    
    sSql = "Select C.dVenc,C.nCuota, C.nColocCalendEstado,"
    sSql = sSql & " nCapital=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
    sSql = sSql & " nIntComp=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSql = sSql & " nIntGracia=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSql = sSql & " nIntMor=(select nMontoPagado  from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSql = sSql & " nIntReprog=(select nMontoPagado  from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSql = sSql & " nIntSuspenso=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
    sSql = sSql & " nGasto=(select sum(nMontoPagado) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' )), "
    sSql = sSql & " dFecCanc = (Select MAX(SubString(M.cMovNro,1,14)) from Mov M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro "
    sSql = sSql & "                    Where MD.cCtaCod = C.cCtaCod AND nNroCalen = C.nNroCalen AND nCuota = C.nCuota) "
    sSql = sSql & " from ColocCalendario C "
    sSql = sSql & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = C.cCtaCod) AND C.nColocCalendEstado = " & gColocCalendEstadoPagado & " order by C.nCuota"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioPagosRealizados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioPagosRealizados:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaNroCuotas(ByVal psCtaCod As String, ByVal pnAplicado As ColocCalendApl, _
    ByVal pnNroCalen As Integer) As Integer
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset

    sSql = "Select Count(*) as nCuotas From  ColocCalendario Where nNroCalen = " & pnNroCalen
    sSql = sSql & " AND nColocCalendApl = " & pnAplicado & " AND cCtaCod = '" & psCtaCod & "' AND nColocCalendEstado = 0 "
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    RecuperaNroCuotas = IIf(IsNull(R!nCuotas), 0, R!nCuotas)
    R.Close
    Set R = Nothing
End Function

Public Function RecuperaCalendarioPagosPendiente(ByVal psCtaCod As String, Optional ByVal pbCalParalelo As Boolean = False, _
Optional ByVal pBrfa As Boolean, Optional ByVal pbExcluyeGastos As Boolean = False, _
Optional ByVal pConn As COMDCredito.DCOMCredActBD = Nothing) As ADODB.Recordset

'CTI2 20181229 ADD pbExcluyeGastos
'pConn 20200302 ADD RIRO

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

'ARCV 27-02-2007
Dim oCons As COMDConstSistema.NCOMConstSistema
Dim dFecSisTmp As Date

Set oCons = New COMDConstSistema.NCOMConstSistema
dFecSisTmp = CDate(oCons.LeeConstSistema(gConstSistFechaSistema))
Set oCons = Nothing
'--------

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    If pBrfa = False Then
            'CONVERTIDO A SP X JUEZ 20150505 *******************************************************
            'sSql = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
            'sSql = sSql & " nMontoPrestamo = (select SUM(nMonto) from ColocCalendDet CD2 Inner Join ColocCalendario C2 ON CD2.cCtaCod=C2.cCtaCod AND CD2.nNroCalen=C2.nNroCalen AND CD2.nColocCalendApl = C2.nColocCalendApl AND CD2.nCuota = C2.nCuota where CD2.cCtaCod = C.cCtaCod And CD2.nNroCalen = C.nNroCalen And CD2.nColocCalendApl=C.nColocCalendApl and CD2.nPrdConceptoCod in(1000,1010) AND C2.nColocCalendEstado = " & gColocCalendEstadoPendiente & "), "
            'sSql = sSql & " nSaldoCap=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(" & gColocConceptoCodCapital & ",1010)),"
            'sSql = sSql & " nCapital=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in( " & gColocConceptoCodCapital & ",1010)),"
            'sSql = sSql & " nIntComp=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in( " & gColocConceptoCodInteresCompensatorio & ",1107)),"
            'sSql = sSql & " nIntCompVenc=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompVencido & "),"
            'sSql = sSql & " nIntGracia=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
            ''ARCV 05-03-2007 (Redondear a dos digitos para evitar las diferencias)
            ''sSql = sSql & " nIntMor=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in ( " & gColocConceptoCodInteresMoratorio & "," & gColocRFAInteresMoratorio & ")),"
            'sSql = sSql & " nIntMor=(select ROUND(nMonto,2) - ROUND(nMontoPagado,2) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in ( " & gColocConceptoCodInteresMoratorio & "," & gColocRFAInteresMoratorio & ")),"
            ''----------
            'sSql = sSql & " nIntReprog=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
            'sSql = sSql & " nIntSuspenso=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
            ''ARCV 27-02-2007 (Cargar solo los Gastos a la Fecha)
            ''sSql = sSql & " nGasto=(select sum(nMonto - nMontoPagado) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' )), "
            ''ALPA20130801*************************************************************************************
            'sSql = sSql & " nGasto= dbo.ColocCred_ObtieneGastoFechaCuotaMiVienda('" & psCtaCod & "',C.nCuota,'" & Format(dFecSisTmp, "mm/dd/yyyy") & "',C.nNroCalen), "
            ''sSql = sSql & " nGasto= dbo.ColocCred_ObtieneGastoFechaCuota('" & psCtaCod & "',C.nCuota,'" & Format(dFecSisTmp, "mm/dd/yyyy") & "'), "
            ''*************************************************************************************************
            ''****************
            'sSql = sSql & " nITF=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(20,21)) "
           
            'sSql = sSql & " from ColocCalendario C "
            'sSql = sSql & " Where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select " & IIf(pbCalParalelo, "nNroCalPar", "nNroCalen") & " from ColocacCred where cCtaCod = C.cCtaCod) "
            
            'sSql = sSql & " AND C.nColocCalendEstado = " & gColocCalendEstadoPendiente
            'sSql = sSql & " order by C.nCuota"
            
            sSql = "stp_sel_RecuperaCalendarioPagosPendiente '" & psCtaCod & "','" & Format(dFecSisTmp, "yyyyMMdd") & "'," & IIf(pbCalParalelo, 1, 0) & ", " & IIf(pbExcluyeGastos, 1, 0)
            'END JUEZ ******************************************************************************
            
            'If pbCalParalelo Then
            '    sSql = sSql & " AND nCuota in (Select nCuota From  ColocCalendario Where cCtaCod = '" & psCtaCod & "' And nColocCalendApl = " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = '" & psCtaCod & "') AND nColocCalendEstado = " & gColocCalendEstadoPendiente & ")"
            'Else
            '    sSql = sSql & " AND C.nColocCalendEstado = " & gColocCalendEstadoPendiente
            'End If
            'sSql = sSql & " order by C.nCuota"
    Else
        
            sSql = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
            sSql = sSql & " nMontoPrestamo = (select SUM(nMonto) from ColocCalendDet CD2 Inner Join ColocCalendario C2 ON CD2.cCtaCod=C2.cCtaCod AND CD2.nNroCalen=C2.nNroCalen AND CD2.nColocCalendApl = C2.nColocCalendApl AND CD2.nCuota = C2.nCuota where CD2.cCtaCod = C.cCtaCod And CD2.nNroCalen = C.nNroCalen And CD2.nColocCalendApl=C.nColocCalendApl and CD2.nPrdConceptoCod in(" & gColocConceptoCodCapital & "," & gColocRFACapital & ") AND C2.nColocCalendEstado = " & gColocCalendEstadoPendiente & "), "
            sSql = sSql & " nSaldoCap=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (" & gColocConceptoCodCapital & "," & gColocRFACapital & ")),"
            sSql = sSql & " nCapital=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (" & gColocConceptoCodCapital & "," & gColocRFACapital & ")),"
            sSql = sSql & " nIntComp=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in ( " & gColocConceptoCodInteresCompensatorio & "," & gColocRFAInteresCompesatorio & ")),"
            sSql = sSql & " nIntCompVenc=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompVencido & "),"
            sSql = sSql & " nIntGracia=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
            'ARCV 05-03-2007 (Redondear a dos digitos para evitar las diferencias)
            'sSql = sSql & " nIntMor=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in ( " & gColocConceptoCodInteresMoratorio & "," & gColocRFAInteresMoratorio & ")),"
            sSql = sSql & " nIntMor=(select ROUND(nMonto,2) - ROUND(nMontoPagado,2) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in ( " & gColocConceptoCodInteresMoratorio & "," & gColocRFAInteresMoratorio & ")),"
            '----------
            sSql = sSql & " nIntReprog=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
            sSql = sSql & " nIntSuspenso=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
            'ARCV 27-02-2007 (Cargar solo los Gastos a la Fecha)
            'sSql = sSql & " nGasto=(select sum(nMonto - nMontoPagado) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' )), "
            'ALPA20130801***********************************
            'sSql = sSql & " nGasto= dbo.ColocCred_ObtieneGastoFechaCuota('" & psCtaCod & "',C.nCuota,'" & Format(dFecSisTmp, "mm/dd/yyyy") & "'), "
            sSql = sSql & " nGasto= dbo.ColocCred_ObtieneGastoFechaCuotaMiVienda('" & psCtaCod & "',C.nCuota,'" & Format(dFecSisTmp, "mm/dd/yyyy") & "',C.nNroCalen), "
            '****************
            sSql = sSql & " nITF=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(20,21)) "
            
            sSql = sSql & " from ColocCalendario C "
            sSql = sSql & " Where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select " & IIf(pbCalParalelo, "nNroCalPar", "nNroCalen") & " from ColocacCred where cCtaCod = C.cCtaCod) "
            
            sSql = sSql & " AND C.nColocCalendEstado = " & gColocCalendEstadoPendiente
            sSql = sSql & " order by C.nCuota"
            
            'If pbCalParalelo Then
            '    sSql = sSql & " AND nCuota in (Select nCuota From  ColocCalendario Where cCtaCod = '" & psCtaCod & "' And nColocCalendApl = " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = '" & psCtaCod & "') AND nColocCalendEstado = " & gColocCalendEstadoPendiente & ")"
            'Else
            '    sSql = sSql & " AND C.nColocCalendEstado = " & gColocCalendEstadoPendiente
            'End If
            'sSql = sSql & " order by C.nCuota"
    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    ' RIRO 20200122 FONDO A CRECER *****
    If Not pConn Is Nothing Then
        Set RecuperaCalendarioPagosPendiente = pConn.CargaRecordSet(sSql)
    Else
        Set RecuperaCalendarioPagosPendiente = oConecta.CargaRecordSet(sSql)
    End If
    ' END ******************************
    
    'Set RecuperaCalendarioPagosPendiente = oConecta.CargaRecordSet(sSql) RIRO 20200122
    
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

'MAVM Corregir Historial del Credito (Gastos)
Public Function RecuperaCalendarioPagosPendienteHistorial(ByVal psCtaCod As String, Optional ByVal pbCalParalelo As Boolean = False, _
Optional ByVal pBrfa As Boolean) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

'ARCV 27-02-2007
Dim oCons As COMDConstSistema.NCOMConstSistema
Dim dFecSisTmp As Date

Set oCons = New COMDConstSistema.NCOMConstSistema
dFecSisTmp = CDate(oCons.LeeConstSistema(gConstSistFechaSistema))
Set oCons = Nothing
'--------

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    If pBrfa = False Then
            sSql = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
            sSql = sSql & " nMontoPrestamo = (select SUM(nMonto) from ColocCalendDet CD2 Inner Join ColocCalendario C2 ON CD2.cCtaCod=C2.cCtaCod AND CD2.nNroCalen=C2.nNroCalen AND CD2.nColocCalendApl = C2.nColocCalendApl AND CD2.nCuota = C2.nCuota where CD2.cCtaCod = C.cCtaCod And CD2.nNroCalen = C.nNroCalen And CD2.nColocCalendApl=C.nColocCalendApl and CD2.nPrdConceptoCod in(1000,1010) AND C2.nColocCalendEstado = " & gColocCalendEstadoPendiente & "), "
            sSql = sSql & " nSaldoCap=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(" & gColocConceptoCodCapital & ",1010)),"
            sSql = sSql & " nCapital=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in( " & gColocConceptoCodCapital & ",1010)),"
            sSql = sSql & " nIntComp=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in( " & gColocConceptoCodInteresCompensatorio & ",1107)),"
            sSql = sSql & " nIntCompVenc=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompVencido & "),"
            sSql = sSql & " nIntGracia=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
            'ARCV 05-03-2007 (Redondear a dos digitos para evitar las diferencias)
            'sSql = sSql & " nIntMor=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in ( " & gColocConceptoCodInteresMoratorio & "," & gColocRFAInteresMoratorio & ")),"
            sSql = sSql & " nIntMor=(select ROUND(nMonto,2) - ROUND(nMontoPagado,2) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in ( " & gColocConceptoCodInteresMoratorio & "," & gColocRFAInteresMoratorio & ")),"
            '----------
            sSql = sSql & " nIntReprog=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
            sSql = sSql & " nIntSuspenso=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
            'ARCV 27-02-2007 (Cargar solo los Gastos a la Fecha)
            sSql = sSql & " nGasto=(select sum(nMonto - nMontoPagado) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' )), "
            'sSql = sSql & " nGasto= dbo.ColocCred_ObtieneGastoFechaCuota('" & psCtaCod & "',C.nCuota,'" & Format(dFecSisTmp, "mm/dd/yyyy") & "'), "
            '****************
            'EJVG20140925 *** nGastoComision= Están otros gastos no detallados
            sSql = sSql & " nGastoComision = ISNULL((select ISNULL(sum(nMonto-nMontoPagado),0.00) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl And nCuota = C.nCuota And nPrdConceptoCod In (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' And nPrdConceptoCod Not In (1217,1231,1243))),0)"
            sSql = sSql & " ,nGastoSegDes = ISNULL((select ISNULL(sum(nMonto-nMontoPagado),0.00) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl And nCuota = C.nCuota And nPrdConceptoCod = 1217),0)"
            sSql = sSql & " ,nGastoPolizaIncendio = ISNULL((select ISNULL(sum(nMonto-nMontoPagado),0.00) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl And nCuota = C.nCuota And nPrdConceptoCod = 1231),0)"
            sSql = sSql & " ,nGastoPolizaVehiculo = ISNULL((select ISNULL(sum(nMonto-nMontoPagado),0.00) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl And nCuota = C.nCuota And nPrdConceptoCod = 1243),0)"
            'END EJVG *******
            sSql = sSql & " ,nITF=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(20,21)) "
            
            sSql = sSql & " from ColocCalendario C "
            sSql = sSql & " Where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select " & IIf(pbCalParalelo, "nNroCalPar", "nNroCalen") & " from ColocacCred where cCtaCod = C.cCtaCod) "
            
            sSql = sSql & " AND C.nColocCalendEstado = " & gColocCalendEstadoPendiente
            sSql = sSql & " order by C.nCuota"
            
            'If pbCalParalelo Then
            '    sSql = sSql & " AND nCuota in (Select nCuota From  ColocCalendario Where cCtaCod = '" & psCtaCod & "' And nColocCalendApl = " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = '" & psCtaCod & "') AND nColocCalendEstado = " & gColocCalendEstadoPendiente & ")"
            'Else
            '    sSql = sSql & " AND C.nColocCalendEstado = " & gColocCalendEstadoPendiente
            'End If
            'sSql = sSql & " order by C.nCuota"
    Else
        
            sSql = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
            sSql = sSql & " nMontoPrestamo = (select SUM(nMonto) from ColocCalendDet CD2 Inner Join ColocCalendario C2 ON CD2.cCtaCod=C2.cCtaCod AND CD2.nNroCalen=C2.nNroCalen AND CD2.nColocCalendApl = C2.nColocCalendApl AND CD2.nCuota = C2.nCuota where CD2.cCtaCod = C.cCtaCod And CD2.nNroCalen = C.nNroCalen And CD2.nColocCalendApl=C.nColocCalendApl and CD2.nPrdConceptoCod in(" & gColocConceptoCodCapital & "," & gColocRFACapital & ") AND C2.nColocCalendEstado = " & gColocCalendEstadoPendiente & "), "
            sSql = sSql & " nSaldoCap=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (" & gColocConceptoCodCapital & "," & gColocRFACapital & ")),"
            sSql = sSql & " nCapital=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (" & gColocConceptoCodCapital & "," & gColocRFACapital & ")),"
            sSql = sSql & " nIntComp=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in ( " & gColocConceptoCodInteresCompensatorio & "," & gColocRFAInteresCompesatorio & ")),"
            sSql = sSql & " nIntCompVenc=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompVencido & "),"
            sSql = sSql & " nIntGracia=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
            'ARCV 05-03-2007 (Redondear a dos digitos para evitar las diferencias)
            'sSql = sSql & " nIntMor=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in ( " & gColocConceptoCodInteresMoratorio & "," & gColocRFAInteresMoratorio & ")),"
            sSql = sSql & " nIntMor=(select ROUND(nMonto,2) - ROUND(nMontoPagado,2) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in ( " & gColocConceptoCodInteresMoratorio & "," & gColocRFAInteresMoratorio & ")),"
            '----------
            sSql = sSql & " nIntReprog=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
            sSql = sSql & " nIntSuspenso=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
            'ARCV 27-02-2007 (Cargar solo los Gastos a la Fecha)
            'sSql = sSql & " nGasto=(select sum(nMonto - nMontoPagado) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' )), "
            sSql = sSql & " nGasto= dbo.ColocCred_ObtieneGastoFechaCuota('" & psCtaCod & "',C.nCuota,'" & Format(dFecSisTmp, "mm/dd/yyyy") & "'), "
            '****************
            'EJVG20140925 *** nGastoComision= Están otros gastos no detallados
            sSql = sSql & " nGastoComision = ISNULL((select ISNULL(sum(nMonto-nMontoPagado),0.00) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl And nCuota = C.nCuota And nPrdConceptoCod In (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' And nPrdConceptoCod Not In (1217,1231,1243))),0)"
            sSql = sSql & " ,nGastoSegDes = ISNULL((select ISNULL(sum(nMonto-nMontoPagado),0.00) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl And nCuota = C.nCuota And nPrdConceptoCod = 1217),0)"
            sSql = sSql & " ,nGastoPolizaIncendio = ISNULL((select ISNULL(sum(nMonto-nMontoPagado),0.00) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl And nCuota = C.nCuota And nPrdConceptoCod = 1231),0)"
            sSql = sSql & " ,nGastoPolizaVehiculo = ISNULL((select ISNULL(sum(nMonto-nMontoPagado),0.00) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl And nCuota = C.nCuota And nPrdConceptoCod = 1243),0)"
            'END EJVG *******
            sSql = sSql & " ,nITF=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(20,21)) "
            
            sSql = sSql & " from ColocCalendario C "
            sSql = sSql & " Where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select " & IIf(pbCalParalelo, "nNroCalPar", "nNroCalen") & " from ColocacCred where cCtaCod = C.cCtaCod) "
            
            sSql = sSql & " AND C.nColocCalendEstado = " & gColocCalendEstadoPendiente
            sSql = sSql & " order by C.nCuota"
            
            'If pbCalParalelo Then
            '    sSql = sSql & " AND nCuota in (Select nCuota From  ColocCalendario Where cCtaCod = '" & psCtaCod & "' And nColocCalendApl = " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = '" & psCtaCod & "') AND nColocCalendEstado = " & gColocCalendEstadoPendiente & ")"
            'Else
            '    sSql = sSql & " AND C.nColocCalendEstado = " & gColocCalendEstadoPendiente
            'End If
            'sSql = sSql & " order by C.nCuota"
    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioPagosPendienteHistorial = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaRefinanciados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaMatrizRefinanciados
    'LUCV20160617 Comento, Según ERS004-2016 **********->
'    sSql = "Select CR.cCtaCod,CR.cCtaCodRef,CR.nMontoRef, "
'    sSql = sSql & "nCapital = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
'    sSql = sSql & "nCapitalPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
'    sSql = sSql & "nIntComp = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
'    sSql = sSql & "nIntCompPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
'    sSql = sSql & "nIntGracia = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
'    sSql = sSql & "nIntGraciaPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
'    sSql = sSql & "nIntMor = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
'    sSql = sSql & "nIntMorPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
'    sSql = sSql & "nIntReprog = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
'    sSql = sSql & "nIntReprogPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
'    sSql = sSql & "nIntSuspenso = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
'    sSql = sSql & "nIntSuspensoPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "), "
'    sSql = sSql & "nGastos = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodGastoVarios & "),"
'    sSql = sSql & "nGastosPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodGastoVarios & ") "
'    sSql = sSql & "From ColocacRefinanc CR Where cCtaCod = '" & psCtaCod & "'"
    'LUCV20160617 Fin <-**********
    
    sSql = " exec stp_sel_RecuperaRefinanciadosAprobados '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaRefinanciados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaMatrizRefinanciados:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaFechaInicioCuota(ByVal psCtaCod As String, ByVal pnCuota As Integer, ByVal pnAplicado As ColocCalendApl, Optional ByVal pbCalendDin As Boolean = False, _
Optional ByVal pConn As COMDCredito.DCOMCredActBD = Nothing) As Date

'pConn RIRO20200302

Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaFechaInicioCuota
    'COMMENT BY MARG20181011----------------
'    If pnCuota > 1 Then
'        sSql = "Select CC.dVenc from ColocCalendario CC "
'        sSql = sSql & " Where CC.nCuota = " & pnCuota - 1 & " And nColocCalendApl = " & pnAplicado & " AND cCtaCod = '" & psCtaCod & "'"
'        sSql = sSql & " And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CC.cCtaCod)"
'    Else
'        If pbCalendDin Then
'            sSql = "Select CC.dVenc from ColocCalendario CC "
'            sSql = sSql & " Where CC.nCuota = " & IIf(pnCuota = 0, 1, pnCuota) & " And nColocCalendApl = 0 AND cCtaCod = '" & psCtaCod & "'"
'            sSql = sSql & " And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CC.cCtaCod)"
'        Else
'            'sSql = "Select dVigencia as dVenc from Colocaciones Where cCtaCod = '" & psCtaCod & "'"
'            'peac 20071219
'            sSql = " select dvenc from coloccalendario where cctacod='" & psCtaCod & "'"
'            sSql = sSql & " and ncoloccalendapl=0"
'            sSql = sSql & " and nnrocalen=(select nnrocalen from colocaccred where cctacod='" & psCtaCod & "')"
'        End If
'    End If
    'END MARG----------------------------
    
    'ADD MARG20181110--------------------
    sSql = " exec stp_sel_RecuperaFechaInicioCuota '" & psCtaCod & "'," & pnCuota & "," & pnAplicado & "," & IIf(pbCalendDin = True, 1, 0)
    'END MARG----------------------------
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    'RIRO 20200123 ********************************
    If Not pConn Is Nothing Then
        Set R = pConn.CargaRecordSet(sSql)
    Else
        Set R = oConecta.CargaRecordSet(sSql)
    End If
    'END RIRO *************************************
    
    'Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    'If DatePart("d", R!dVenc) <= 12 Then
        'If pnCuota = 1 Then
         '   RecuperaFechaInicioCuota = FormateandoFecha(R!dVenc)
     '   Else
            RecuperaFechaInicioCuota = Format(R!dVenc, "dd/mm/yyyy")
      '  End If
    'Else
     '   RecuperaFechaInicioCuota = Format(R!dVenc, "dd/mm/yyyy")
    'End If
    Exit Function

ErrorRecuperaFechaInicioCuota:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function
Function FormateandoFecha(ByVal pFecha As Date) As Date
Dim sDia As String
Dim sMes As String
Dim sAno As String

sDia = Mid(CStr(pFecha), 4, 2)
sMes = Mid(CStr(pFecha), 1, 2)
sAno = Mid(CStr(pFecha), 7, 4)
FormateandoFecha = CDate(sDia & "/" & sMes & "/" & sAno)
End Function

Public Function RecuperaCalendarioPagos(ByVal psCtaCod As String, _
                            Optional ByVal pnNroCalen As Integer = -1, Optional ByVal pbNroCalenParalelo As Boolean = False, _
                            Optional ByVal pbComItemFE As Boolean = False, Optional ByVal pbInicial As Boolean = False, _
                            Optional ByVal pbPagoCuota As Boolean = False, Optional ByVal pbPerdon As Boolean = False, _
                            Optional ByVal pbPagoAnticipado As Boolean = False, Optional psTipoOpeCod As String = "", _
                            Optional ByVal pConn As COMDCredito.DCOMCredActBD = Nothing) As ADODB.Recordset
                            'LUCV20180424, Agregó pbPagoAnticipado
                            'NAGL 20190816 Según ERS036-2018 agregó psTipoOpeCod
                            'pConn RIRO20200302
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim lxpbInicial, lxpbPagoCuota, lxnNroCalenParalelo, lxnNroCalen As Integer
On Error GoTo ErrorRecuperaCalendarioDesemb

    '*** PEAC 20090721 - COMENTADO PARA CREAR UN STP
   If pbComItemFE Then
   
   Else
   
   End If

   If pbInicial Then
       lxpbInicial = 1
   Else
       lxpbInicial = 2
   End If
  
   If pbPagoCuota = True Then
       lxpbPagoCuota = 1
   Else
       lxpbPagoCuota = 2
   End If
        
    If Not pbNroCalenParalelo Then
        lxnNroCalenParalelo = 2
        lxnNroCalen = pnNroCalen
    Else
        lxnNroCalenParalelo = 1
    End If
    
    '*** PEAC 20090721
    If psTipoOpeCod = CStr(gsPerdonMoraCred) Then
         sSql = " exec stp_sel_ERS0362018_RecuperaCalendarioPagosNew '" & psCtaCod & "'," & lxpbInicial & "," & lxpbPagoCuota & "," & lxnNroCalenParalelo & "," & pnNroCalen & "," & IIf(pbPerdon = True, 1, 0) 'Agregado by NAGL 20190816 Según ERS036-2018
    Else
        sSql = " exec stp_sel_RecuperaCalendarioPagos '" & psCtaCod & "'," & lxpbInicial & "," & lxpbPagoCuota & "," & lxnNroCalenParalelo & "," & pnNroCalen & "," & IIf(pbPerdon = True, 1, 0) & "," & IIf(pbPagoAnticipado = True, 1, 0)
        'LUCV20180424, Agregó pbPagoAnticipado
    End If 'NAGL 20190816 Agregó Condicional
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    'RIRO 20200123 *****
    If Not pConn Is Nothing Then
        Set RecuperaCalendarioPagos = pConn.CargaRecordSet(sSql)
    Else
        Set RecuperaCalendarioPagos = oConecta.CargaRecordSet(sSql)
    End If
    'END RIRO 20200123 *****
    
    'Set RecuperaCalendarioPagos = oConecta.CargaRecordSet(sSql) RIRO 20200123
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function


Public Function RecuperaColocCalendario(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, ByVal pnAplicado As ColocCalendApl, Optional ByVal pnNroCuota As Integer = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaColocCalendario
    sSql = "Select * from ColocCalendario where cCtacod = '" & psCtaCod & "' AND nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnAplicado
    If pnNroCuota <> -1 Then
        sSql = sSql & " AND nCuota = " & pnNroCuota
    End If
    sSql = sSql & "  order by nCuota "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaColocCalendario = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocCalendario:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Function RecuperaColocCalendDet(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, ByVal pnAplicado As ColocCalendApl) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaColocCalendDet
    sSql = "Select * from ColocCalendDet where cCtacod = '" & psCtaCod & "' And nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnAplicado
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaColocCalendDet = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocCalendDet:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaGastosCuotaDesemb(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
    ByVal pnAplicado As ColocCalendApl, Optional ByVal pnCuota As Integer = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaGastosCuotaDesemb
    
'    sSql = "Select C.nCuota, C.nMonto, C.nMontoPagado, CN.cDescripcion cGasto, C.nPrdConceptoCod"
'    sSql = sSql & " From ColocCalendDet C Inner Join ProductoConcepto CN ON C.nPrdConceptoCod = CN.nPrdConceptoCod AND CN.nPrdConceptoCod like '" & Mid(Trim(Str(gColocConceptoCodGastoVarios)), 1, 2) & "%'"
'    sSql = sSql & " Where cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnAplicado
'    If pnCuota <> -1 Then
'        sSql = sSql & " AND nCuota = " & pnCuota
'    End If
    
    sSql = "EXEC stp_sel_RecuperaGastosCuotaDesemb '" & psCtaCod & "'," & pnNroCalen & "," & pnAplicado & "," & pnCuota 'APRI20181121 ERS071-2018
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGastosCuotaDesemb = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaGastosCuotaDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
End Function

Public Function CuotasPendientes(ByVal psCtaCod As String) As Integer
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorCuotasPendientes
    sSql = "Select count(*) as nCuotasPend from ColocCalendario CC where cCtaCod = '" & psCtaCod & "' "
    sSql = sSql & " AND nColocCalendApl = " & gColocCalendAplCuota & " AND nColocCalendEstado = " & gColocCalendEstadoPendiente & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = CC.cCtaCod)"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    CuotasPendientes = R!nCuotasPend
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorCuotasPendientes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function FechaCuotaPendiente(ByVal psCtaCod As String) As Date
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    
    sSql = "Select top 1 dVenc from ColocCalendario CC where cCtaCod = '" & psCtaCod & "' AND nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = '" & psCtaCod & "') AND nColocCalendEstado = 0 AND nColocCalendApl = 1 Order By nCuota "
        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.EOF And Not R.BOF Then
        FechaCuotaPendiente = R!dVenc
    Else
        FechaCuotaPendiente = "01/01/1900"
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    

End Function

Public Function RecuperaListaDesembolsosParciales(ByVal pcCodCta As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    
    sSql = "select "
    sSql = sSql & " case when cc.nColocCalendEstado=1 then 'SI' else 'NO' end as cActual, "
    sSql = sSql & " CC.nCuota , dVenc, CCD.nMonto "
    sSql = sSql & " from coloccalendario CC "
    sSql = sSql & " Inner Join ColocacCred CCr ON CC.cCtaCod=CCr.cCtaCod and CC.nNroCalen=CCr.nNroCalen "
    sSql = sSql & " Inner Join ColocCalendDet CCD ON CC.cCtaCod=CCD.cCtaCod and CC.nNroCalen=CCD.nNroCalen "
    sSql = sSql & " and CC.nColocCalendApl=CCD.nColocCalendApl and CC.nCuota=CCD.nCuota "
    sSql = sSql & " and CCD.nPrdConceptoCod=" & gColocConceptoCodCapital & " "
    sSql = sSql & " where CC.cctacod='" & pcCodCta & "' and CC.nColocCalendApl=" & gColocCalendAplDesembolso & " Order By CC.nCuota "
        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaListaDesembolsosParciales = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function VerificaCalendQuincenal(ByVal psCtaCod As String) As Boolean
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "Select nColocCalendCod"
    sSql = sSql & " From ColocacCred"
    sSql = sSql & " Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not rs.EOF And Not rs.BOF Then
        If rs!nColocCalendCod = "81" Then
            VerificaCalendQuincenal = True
        Else
            VerificaCalendQuincenal = False
        End If
    End If
    Set rs = Nothing
End Function


Public Function ObtenerDiasPeriodoParcial(ByVal psCtaCod As String, ByVal pdFecDesem As Date) As Integer
    Dim sSql As String
    Dim sFecha As String
    Dim rs As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    
    sFecha = Format(pdFecDesem, "MM/dd/yyyy")
    sSql = "Select dbo.fn_PeriodoParcia('" & psCtaCod & "','" & sFecha & "') as nDias"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not rs.EOF And Not rs.BOF Then
        ObtenerDiasPeriodoParcial = rs!nDias
    End If
    Set rs = Nothing
    Set oConec = Nothing
End Function

'MAVM 20120808 ***
Public Function RecuperaPagCredDiaAnt(ByVal psCtaCod As String) As ADODB.Recordset
Dim oCons As COMDConstSistema.NCOMConstSistema
Dim dFecSisTmp As Date
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

Set oCons = New COMDConstSistema.NCOMConstSistema
dFecSisTmp = CDate(oCons.LeeConstSistema(gConstSistFechaSistema))
Set oCons = Nothing
    
    On Error GoTo ErrorRecuperaPagCredDiaAnt
    sSql = "Select nMonto "
    sSql = sSql & "From MovColDet MCD Inner Join Mov M on MCD.nMovNro = M.nMovNro Where cCtaCod = '" & psCtaCod & "'"
    sSql = sSql & " And SubString(cMovNro, 1, 8) = '" & Format(dFecSisTmp - 1, "yyyymmdd") & "'"
    sSql = sSql & " And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = MCD.cCtaCod) And nPrdConceptoCod = 1000"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPagCredDiaAnt = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaPagCredDiaAnt:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'***
'JUEZ 20150708 ****************************************************************************************
Public Function VerificaCuotaGastoPagadoCalendAnt(ByVal psCtaCod As String, ByVal pnAnio As Integer, ByVal pnMes As Integer) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorVerificaCuotaGastoPagadoCalendAnt
    
    VerificaCuotaGastoPagadoCalendAnt = False
    
    sSql = "stp_sel_VerificaCuotaGastoPagadoCalendAnt '" & psCtaCod & "'," & pnAnio & "," & pnMes
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If Not R.BOF And Not R.EOF Then VerificaCuotaGastoPagadoCalendAnt = True
    
    Exit Function

ErrorVerificaCuotaGastoPagadoCalendAnt:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'END JUEZ *********************************************************************************************
'APRI20181121 ERS071-2018
Public Function ObtenerCalendarioCredito(ByVal pcCodCta As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    
    sSql = "EXEC stp_sel_CalendarioCredito '" & pcCodCta & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerCalendarioCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
'END APRI

'->***** LUCV20180601, Agregó según ERS022-2018. Interés de gracia en la primera cuota
Public Function TieneCalendarioMetodoInterativo(ByVal psCtaCod As String) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorTieneCalendarioMetodoInterativo

    TieneCalendarioMetodoInterativo = False
    
    sSql = "Exec stp_sel_ERS0222018_AplicarCalendarioMetodoInterativo '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If Not R.EOF And Not R.BOF Then
        TieneCalendarioMetodoInterativo = True
    Else
        TieneCalendarioMetodoInterativo = False
    End If
    
    Exit Function
ErrorTieneCalendarioMetodoInterativo:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'<-*****  Fin LUCV20180601.

'RIRO 20200910 ********************
Public Function ObtieneLiquidacionCab(ByVal psCtaCod As String, psFecha As String) As ADODB.Recordset
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtieneLiquidacionCab
    
    sSql = "Exec stp_sel_LiquidacionCab '" & psCtaCod & "', '" & psFecha & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Set ObtieneLiquidacionCab = R
       
    Exit Function
ErrorObtieneLiquidacionCab:
    Err.Raise Err.Number, "Error En Proceso de obtención de liquidació de crédito", Err.Description
End Function
Public Function VerificaActualizacionLiquidacion(ByVal psCtaCod As String, ByRef psMensaje As String) As Boolean
    
    Dim sSql As String
    Dim R As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim bResultado As Boolean

    On Error GoTo ErrorVerificaActualizacionLiquidacion
    
    bResultado = False
    sSql = "Exec stp_sel_VerificaActualizacionLiquidacion '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If Not R Is Nothing Then
        If R.State = 1 Then
            If Not R.EOF And Not R.BOF Then
                If R.RecordCount > 0 Then
                    bResultado = R!bLiquidacion
                    psMensaje = R!cMensaje
                End If
            End If
        End If
    End If
    
    VerificaActualizacionLiquidacion = bResultado
       
    Exit Function
ErrorVerificaActualizacionLiquidacion:
    Err.Raise Err.Number, "Error En Proceso de obtención de liquidació de crédito", Err.Description
End Function
'END RIRO 20200910 ****************







