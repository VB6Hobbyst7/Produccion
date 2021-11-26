VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dcalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RecuperaCalendarioGastos(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
                    ByVal pnNroCuota As Integer, ByVal pnAplicado As ColocCalendApl, _
                    Optional ByVal pbTodos As Boolean = False) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta


    On Error GoTo ErrorRecuperaCalendarioGastos
    sSQL = "Select * from ColocCalendDet where cCtaCod = '" & psCtaCod & _
            "' AND nNroCalen = " & pnNroCalen
    If Not pbTodos Then
        sSQL = sSQL & " AND nCuota = " & pnNroCuota
    End If
    sSQL = sSQL & " AND nColocCalendApl = " & pnAplicado
    sSQL = sSQL & " AND nPrdConceptoCod like '12%' Order By nCuota "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioGastos = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaCalendarioGastos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaCalendarioGastosPendientes(ByVal psCtaCod As String, _
                    ByVal pnAplicado As ColocCalendApl) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrRecuperaCalendarioGastosPendientes
            
    sSQL = " Select CD.* From ColocCalendario C Inner Join ColocCalendDet CD "
    sSQL = sSQL & " ON C.cCtaCod = CD.cCtaCod AND C.nNroCalen = CD.nNroCalen AND C.nColocCalendApl = CD.nColocCalendApl "
    sSQL = sSQL & " AND C.nCuota = CD.nCuota "
    sSQL = sSQL & " Where  C.cCtaCod = '" & psCtaCod & "' AND  C.nColocCalendApl = " & pnAplicado & " AND C.nColocCalendEstado = 0 "
    sSQL = sSQL & " AND CD.nPrdConceptoCod like '12%' AND C.nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = '" & psCtaCod & "') Order By C.nCuota   "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioGastosPendientes = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrRecuperaCalendarioGastosPendientes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaCalendarioGastosPendientesAFecha(ByVal psCtaCod As String, _
                    ByVal pnAplicado As ColocCalendApl, ByVal pdFecha As Date) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    
    sSQL = " Select PC.cDescripcion, (CD.nMonto-CD.nMontoPagado) as nMonto "
    sSQL = sSQL & "  From ColocCalendario C Inner Join ColocCalendDet CD "
    sSQL = sSQL & " ON C.cCtaCod = CD.cCtaCod AND C.nNroCalen = CD.nNroCalen AND C.nColocCalendApl = CD.nColocCalendApl "
    sSQL = sSQL & " AND C.nCuota = CD.nCuota "
    sSQL = sSQL & " Inner Join ProductoConcepto PC ON PC.nPrdConceptoCod = CD.nPrdConceptoCod "
    sSQL = sSQL & " Where  C.cCtaCod = '" & psCtaCod & "' AND  C.nColocCalendApl = " & pnAplicado & " AND C.nColocCalendEstado = 0 "
    sSQL = sSQL & " AND CD.nPrdConceptoCod like '12%' AND C.nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = '" & psCtaCod & "') "
    sSQL = sSQL & " AND (C.nCuota = (Select nNroProxCuota From ColocacCred Where cCtaCod = C.cCtaCod) OR C.dVenc <= '" & Format(pdFecha, "mm/dd/yyyy") & "') "
    sSQL = sSQL & " Order By C.nCuota   "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioGastosPendientesAFecha = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    
End Function

Public Function RecuperaCalendarioGastosPendientesAgrupado(ByVal psCuentas As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioGastosPendientesAgrupado
            
    sSQL = " Select CD.nPrdConceptoCod, SUM(CD.nMonto - CD.nMontoPagado) as nMonto From ColocCalendario C Inner Join ColocCalendDet CD "
    sSQL = sSQL & " ON C.cCtaCod = CD.cCtaCod AND C.nNroCalen = CD.nNroCalen AND C.nColocCalendApl = CD.nColocCalendApl "
    sSQL = sSQL & " AND C.nCuota = CD.nCuota AND C.nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = C.cCtaCod ) "
    sSQL = sSQL & " Where  C.cCtaCod In  " & psCuentas & " AND  C.nColocCalendEstado = 0 "
    sSQL = sSQL & " AND CD.nPrdConceptoCod like '12%' AND C.nColocCalendApl = 1 "
    sSQL = sSQL & " Group By CD.nPrdConceptoCod "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioGastosPendientesAgrupado = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaCalendarioGastosPendientesAgrupado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Function RecuperaCalendarioDesemb(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    sSQL = "select C.dVenc,C.dPago, C.nCuota,C.nColocCalendEstado, "
    sSQL = sSQL & " CASE WHEN C.nColocCalendEstado = 1 THEN 'DESEMBOLSADO' WHEN C.nColocCalendEstado = 0 THEN 'PENDIENTE' END AS cColocCalendEstado,"
    sSQL = sSQL & " nCapital=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1000),"
    sSQL = sSQL & " nIntComp=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1100),"
    sSQL = sSQL & " nIntGracia=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1102),"
    sSQL = sSQL & " nIntMor=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1101),"
    sSQL = sSQL & " nIntReprog=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1103),"
    sSQL = sSQL & " nGasto=(select sum(nMonto) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where nPrdConceptoCod like '12%' )) "
    sSQL = sSQL & " from ColocCalendario C inner join ColocacCred CC on C.cCtaCod=CC.cCtaCod "
    sSQL = sSQL & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl = " & gColocCalendAplDesembolso & " and C.nNroCalen = CC.nNroCalen order by C.nCuota"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioDesemb = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaCalendarioDesembPagados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    sSQL = "select C.dVenc,C.nCuota,C.nColocCalendEstado, "
    sSQL = sSQL & " nCapital=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1000),"
    sSQL = sSQL & " nIntComp=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1100),"
    sSQL = sSQL & " nIntGracia=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1102),"
    sSQL = sSQL & " nIntMor=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1101),"
    sSQL = sSQL & " nIntReprog=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1103),"
    sSQL = sSQL & " nGasto=(select sum(nMonto) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where nPrdConceptoCod like '12%' )) "
    sSQL = sSQL & " from ColocCalendario C inner join ColocacCred CC on C.cCtaCod=CC.cCtaCod "
    sSQL = sSQL & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl = " & gColocCalendAplDesembolso & " and C.nNroCalen = CC.nNroCalen AND C.nColocCalendEstado=1 order by C.nCuota"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    
    Set RecuperaCalendarioDesembPagados = oConecta.CargaRecordSet(sSQL)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function


Public Function RecuperaCalendarioPagosDeuda(ByVal psCtaCod As String, Optional ByVal pnNroCalen As Integer = -1) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    sSQL = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
    sSQL = sSQL & " nCapital=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
    sSQL = sSQL & " nIntComp=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSQL = sSQL & " nIntGracia=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSQL = sSQL & " nIntMor=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSQL = sSQL & " nIntReprog=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSQL = sSQL & " nIntSuspenso=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
    sSQL = sSQL & " nGasto=(select sum(nMonto- nMontoPagado) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' ))"
    sSQL = sSQL & " from ColocCalendario C "
    If pnNroCalen = -1 Then
        sSQL = sSQL & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = C.cCtaCod) order by C.nCuota"
    Else
        sSQL = sSQL & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = " & pnNroCalen & " order by C.nCuota"
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioPagosDeuda = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaCalendarioPagosRealizados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioPagosRealizados
    
    sSQL = "Select C.dVenc,C.nCuota, C.nColocCalendEstado,"
    sSQL = sSQL & " nCapital=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
    sSQL = sSQL & " nIntComp=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSQL = sSQL & " nIntGracia=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSQL = sSQL & " nIntMor=(select nMontoPagado  from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSQL = sSQL & " nIntReprog=(select nMontoPagado  from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSQL = sSQL & " nIntSuspenso=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
    sSQL = sSQL & " nGasto=(select sum(nMontoPagado) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' )), "
    sSQL = sSQL & " dFecCanc = (Select MAX(SubString(M.cMovNro,1,14)) from Mov M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro "
    sSQL = sSQL & "                    Where MD.cCtaCod = C.cCtaCod AND nNroCalen = C.nNroCalen AND nCuota = C.nCuota) "
    sSQL = sSQL & " from ColocCalendario C "
    sSQL = sSQL & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = C.cCtaCod) AND C.nColocCalendEstado = " & gColocCalendEstadoPagado & " order by C.nCuota"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioPagosRealizados = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioPagosRealizados:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaNroCuotas(ByVal psCtaCod As String, ByVal pnAplicado As ColocCalendApl, _
    ByVal pnNroCalen As Integer) As Integer
Dim oCon As DConecta
Dim sSQL As String
Dim R As ADODB.Recordset

    sSQL = "Select Count(*) as nCuotas From  ColocCalendario Where nNroCalen = " & pnNroCalen
    sSQL = sSQL & " AND nColocCalendApl = " & pnAplicado & " AND cCtaCod = '" & psCtaCod & "' AND nColocCalendEstado = 0 "
    
    Set oCon = New DConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    RecuperaNroCuotas = IIf(IsNull(R!nCuotas), 0, R!nCuotas)
    R.Close
    Set R = Nothing
End Function

Public Function RecuperaCalendarioPagosPendiente(ByVal psCtaCod As String, Optional ByVal pbCalParalelo As Boolean = False, _
Optional ByVal pBrfa As Boolean) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    If pBrfa = False Then
            sSQL = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
            sSQL = sSQL & " nMontoPrestamo = (select SUM(nMonto) from ColocCalendDet CD2 Inner Join ColocCalendario C2 ON CD2.cCtaCod=C2.cCtaCod AND CD2.nNroCalen=C2.nNroCalen AND CD2.nColocCalendApl = C2.nColocCalendApl AND CD2.nCuota = C2.nCuota where CD2.cCtaCod = C.cCtaCod And CD2.nNroCalen = C.nNroCalen And CD2.nColocCalendApl=C.nColocCalendApl and CD2.nPrdConceptoCod = 1000 AND C2.nColocCalendEstado = " & gColocCalendEstadoPendiente & "), "
            sSQL = sSQL & " nSaldoCap=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
            sSQL = sSQL & " nCapital=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
            sSQL = sSQL & " nIntComp=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
            sSQL = sSQL & " nIntCompVenc=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompVencido & "),"
            sSQL = sSQL & " nIntGracia=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
            sSQL = sSQL & " nIntMor=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
            sSQL = sSQL & " nIntReprog=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
            sSQL = sSQL & " nIntSuspenso=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
            sSQL = sSQL & " nGasto=(select sum(nMonto - nMontoPagado) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' )), "
            sSQL = sSQL & " nITF=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(20,21)) "
            
            sSQL = sSQL & " from ColocCalendario C "
            sSQL = sSQL & " Where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select " & IIf(pbCalParalelo, "nNroCalPar", "nNroCalen") & " from ColocacCred where cCtaCod = C.cCtaCod) "
            
            sSQL = sSQL & " AND C.nColocCalendEstado = " & gColocCalendEstadoPendiente
            sSQL = sSQL & " order by C.nCuota"
            
            'If pbCalParalelo Then
            '    sSql = sSql & " AND nCuota in (Select nCuota From  ColocCalendario Where cCtaCod = '" & psCtaCod & "' And nColocCalendApl = " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = '" & psCtaCod & "') AND nColocCalendEstado = " & gColocCalendEstadoPendiente & ")"
            'Else
            '    sSql = sSql & " AND C.nColocCalendEstado = " & gColocCalendEstadoPendiente
            'End If
            'sSql = sSql & " order by C.nCuota"
    Else
        
            sSQL = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
            sSQL = sSQL & " nMontoPrestamo = (select SUM(nMonto) from ColocCalendDet CD2 Inner Join ColocCalendario C2 ON CD2.cCtaCod=C2.cCtaCod AND CD2.nNroCalen=C2.nNroCalen AND CD2.nColocCalendApl = C2.nColocCalendApl AND CD2.nCuota = C2.nCuota where CD2.cCtaCod = C.cCtaCod And CD2.nNroCalen = C.nNroCalen And CD2.nColocCalendApl=C.nColocCalendApl and CD2.nPrdConceptoCod in(" & gColocConceptoCodCapital & "," & gColocRFACapital & ") AND C2.nColocCalendEstado = " & gColocCalendEstadoPendiente & "), "
            sSQL = sSQL & " nSaldoCap=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (" & gColocConceptoCodCapital & "," & gColocRFACapital & ")),"
            sSQL = sSQL & " nCapital=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (" & gColocConceptoCodCapital & "," & gColocRFACapital & ")),"
            sSQL = sSQL & " nIntComp=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in ( " & gColocConceptoCodInteresCompensatorio & "," & gColocRFAInteresCompesatorio & ")),"
            sSQL = sSQL & " nIntCompVenc=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompVencido & "),"
            sSQL = sSQL & " nIntGracia=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
            sSQL = sSQL & " nIntMor=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in ( " & gColocConceptoCodInteresMoratorio & "," & gColocRFAInteresMoratorio & ")),"
            sSQL = sSQL & " nIntReprog=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
            sSQL = sSQL & " nIntSuspenso=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
            sSQL = sSQL & " nGasto=(select sum(nMonto - nMontoPagado) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' )), "
            sSQL = sSQL & " nITF=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(20,21)) "
            
            sSQL = sSQL & " from ColocCalendario C "
            sSQL = sSQL & " Where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select " & IIf(pbCalParalelo, "nNroCalPar", "nNroCalen") & " from ColocacCred where cCtaCod = C.cCtaCod) "
            
            sSQL = sSQL & " AND C.nColocCalendEstado = " & gColocCalendEstadoPendiente
            sSQL = sSQL & " order by C.nCuota"
            
            'If pbCalParalelo Then
            '    sSql = sSql & " AND nCuota in (Select nCuota From  ColocCalendario Where cCtaCod = '" & psCtaCod & "' And nColocCalendApl = " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = '" & psCtaCod & "') AND nColocCalendEstado = " & gColocCalendEstadoPendiente & ")"
            'Else
            '    sSql = sSql & " AND C.nColocCalendEstado = " & gColocCalendEstadoPendiente
            'End If
            'sSql = sSql & " order by C.nCuota"
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioPagosPendiente = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaRefinanciados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaMatrizRefinanciados
    sSQL = "Select CR.cCtaCod,CR.cCtaCodRef,CR.nMontoRef, "
    sSQL = sSQL & "nCapital = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
    sSQL = sSQL & "nCapitalPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
    sSQL = sSQL & "nIntComp = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSQL = sSQL & "nIntCompPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSQL = sSQL & "nIntGracia = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSQL = sSQL & "nIntGraciaPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSQL = sSQL & "nIntMor = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSQL = sSQL & "nIntMorPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSQL = sSQL & "nIntReprog = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSQL = sSQL & "nIntReprogPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSQL = sSQL & "nIntSuspenso = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
    sSQL = sSQL & "nIntSuspensoPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "), "
    sSQL = sSQL & "nGastos = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodGastoVarios & "),"
    sSQL = sSQL & "nGastosPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodGastoVarios & ") "
    sSQL = sSQL & "From ColocacRefinanc CR Where cCtaCod = '" & psCtaCod & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaRefinanciados = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaMatrizRefinanciados:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
    
End Function

Public Function RecuperaFechaInicioCuota(ByVal psCtaCod As String, ByVal pnCuota As Integer, ByVal pnAplicado As ColocCalendApl) As Date
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaFechaInicioCuota
    If pnCuota > 1 Then
        sSQL = "Select CC.dVenc from ColocCalendario CC "
        sSQL = sSQL & " Where CC.nCuota = " & pnCuota - 1 & " And nColocCalendApl = " & pnAplicado & " AND cCtaCod = '" & psCtaCod & "'"
        sSQL = sSQL & " And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CC.cCtaCod)"
    Else
        sSQL = "Select dVigencia as dVenc from Colocaciones Where cCtaCod = '" & psCtaCod & "'"
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    'If DatePart("d", R!dVenc) <= 12 Then
        'If pnCuota = 1 Then
         '   RecuperaFechaInicioCuota = FormateandoFecha(R!dVenc)
     '   Else
            RecuperaFechaInicioCuota = Format(R!DVENC, "dd/mm/yyyy")
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
    Optional ByVal pbComItemFE As Boolean = False, Optional ByVal pbInicial As Boolean = False) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
        If pbComItemFE Then
            sSQL = "select '',C.dVenc,C.nCuota, C.nColocCalendEstado,"
        Else
            sSQL = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
        End If
        sSQL = sSQL & " nCapital=ISNULL((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodCapital & "),0),"
        sSQL = sSQL & " nCapitalPag=ISNULL((select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodCapital & "),0),"
        sSQL = sSQL & " nIntComp=ISNULL((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),0),"
        sSQL = sSQL & " nIntCompPag=ISNULL((select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),0),"
        sSQL = sSQL & " nIntGracia=ISNULL((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),0),"
        sSQL = sSQL & " nIntGraciaPag=ISNULL((select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),0),"
        sSQL = sSQL & " nIntMor=ISNULL((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),0),"
        sSQL = sSQL & " nIntMorPag=ISNULL((select ISNULL(nMontoPagado,0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),0),"
        sSQL = sSQL & " nIntReprog=ISNULL((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),0),"
        sSQL = sSQL & " nIntReprogPag=ISNULL((select ISNULL(nMontoPagado,0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),0),"
        sSQL = sSQL & " nIntSuspenso=ISNULL((select ISNULL(nMonto,0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),0),"
        sSQL = sSQL & " nIntSuspensoPag=ISNULL((select ISNULL(nMontoPagado,0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),0),"
        If pbInicial Then
            sSQL = sSQL & " nGasto=ISNULL((select ISNULL(sum(nMonto),0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' AND cAplicaProceso = 'DE' )),0),"
        Else
            sSQL = sSQL & " nGasto=ISNULL((select ISNULL(sum(nMonto),0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' )),0),"
        End If
        sSQL = sSQL & " nGastoPag=ISNULL((select ISNULL(sum(nMontoPagado),0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' )),0),"
        sSQL = sSQL & " nIntCompVenc=ISNULL((select ISNULL(nMonto,0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompVencido & "),0),"
        sSQL = sSQL & " nIntCompVencPag=ISNULL((select ISNULL(nMontoPagado,0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompVencido & "),0), "
        sSQL = sSQL & " C.dPago"
        sSQL = sSQL & " from ColocCalendario C "
        
            
    If Not pbNroCalenParalelo Then
        If pnNroCalen = -1 Then
            sSQL = sSQL & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = C.cCtaCod) order by C.nCuota"
        Else
            sSQL = sSQL & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = " & pnNroCalen & " order by C.nCuota"
        End If
    Else
        sSQL = sSQL & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalPar from ColocacCred where cCtaCod = C.cCtaCod) order by C.nCuota"
    End If
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioPagos = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function



Public Function RecuperaColocCalendario(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, ByVal pnAplicado As ColocCalendApl, Optional ByVal pnNroCuota As Integer = -1) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaColocCalendario
    sSQL = "Select * from ColocCalendario where cCtacod = '" & psCtaCod & "' AND nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnAplicado
    If pnNroCuota <> -1 Then
        sSQL = sSQL & " AND nCuota = " & pnNroCuota
    End If
    sSQL = sSQL & "  order by nCuota "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocCalendario = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocCalendario:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Function RecuperaColocCalendDet(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, ByVal pnAplicado As ColocCalendApl) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaColocCalendDet
    sSQL = "Select * from ColocCalendDet where cCtacod = '" & psCtaCod & "' And nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnAplicado
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocCalendDet = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocCalendDet:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaGastosCuotaDesemb(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
    ByVal pnAplicado As ColocCalendApl, Optional ByVal pnCuota As Integer = -1) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaGastosCuotaDesemb
    sSQL = "Select C.nCuota, C.nMonto, C.nMontoPagado, CN.cDescripcion cGasto, C.nPrdConceptoCod"
    sSQL = sSQL & " From ColocCalendDet C Inner Join ProductoConcepto CN ON C.nPrdConceptoCod = CN.nPrdConceptoCod AND CN.nPrdConceptoCod like '" & Mid(Trim(Str(gColocConceptoCodGastoVarios)), 1, 2) & "%'"
    sSQL = sSQL & " Where cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnAplicado
    If pnCuota <> -1 Then
        sSQL = sSQL & " AND nCuota = " & pnCuota
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGastosCuotaDesemb = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaGastosCuotaDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
End Function

Public Function CuotasPendientes(ByVal psCtaCod As String) As Integer
Dim sSQL As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorCuotasPendientes
    sSQL = "Select count(*) as nCuotasPend from ColocCalendario CC where cCtaCod = '" & psCtaCod & "' "
    sSQL = sSQL & " AND nColocCalendApl = " & gColocCalendAplCuota & " AND nColocCalendEstado = " & gColocCalendEstadoPendiente & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = CC.cCtaCod)"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
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
Dim sSQL As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    
    sSQL = "Select top 1 dVenc from ColocCalendario CC where cCtaCod = '" & psCtaCod & "' AND nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = '" & psCtaCod & "') AND nColocCalendEstado = 0 AND nColocCalendApl = 1 Order By nCuota "
        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    If Not R.EOF And Not R.BOF Then
        FechaCuotaPendiente = R!DVENC
    Else
        FechaCuotaPendiente = "01/01/1900"
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    

End Function

Public Function RecuperaListaDesembolsosParciales(ByVal pcCodCta As String) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    
    sSQL = "select "
    sSQL = sSQL & " case when cc.nCuota=(select nnroproxdesemb from colocaccred where cctacod='108012021000155918') then 'SI' else 'NO' end as cActual, "
    sSQL = sSQL & " CC.nCuota , dVenc, CCD.nMonto "
    sSQL = sSQL & " from coloccalendario CC "
    sSQL = sSQL & " Inner Join ColocacCred CCr ON CC.cCtaCod=CCr.cCtaCod and CC.nNroCalen=CCr.nNroCalen "
    sSQL = sSQL & " Inner Join ColocCalendDet CCD ON CC.cCtaCod=CCD.cCtaCod and CC.nNroCalen=CCD.nNroCalen "
    sSQL = sSQL & " and CC.nColocCalendApl=CCD.nColocCalendApl and CC.nCuota=CCD.nCuota "
    sSQL = sSQL & " and CCD.nPrdConceptoCod=" & gColocConceptoCodCapital & " "
    sSQL = sSQL & " where CC.cctacod='" & pcCodCta & "' and CC.nColocCalendApl=" & gColocCalendAplDesembolso & " Order By CC.nCuota "
        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaListaDesembolsosParciales = oConecta.CargaRecordSet(sSQL)
    
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function VerificaCalendQuincenal(ByVal psCtaCod As String) As Boolean
    Dim sSQL As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    
    sSQL = "Select nColocCalendCod"
    sSQL = sSQL & " From ColocacCred"
    sSQL = sSQL & " Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
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


