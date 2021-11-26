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
Dim sSql As String
Dim oConecta As DConecta


    On Error GoTo ErrorRecuperaCalendarioGastos
    sSql = "Select * from ColocCalendDet where cCtaCod = '" & psCtaCod & _
            "' AND nNroCalen = " & pnNroCalen
    If Not pbTodos Then
        sSql = sSql & " AND nCuota = " & pnNroCuota
    End If
    sSql = sSql & " AND nColocCalendApl = " & pnAplicado
    sSql = sSql & " AND nPrdConceptoCod like '12%' Order By nCuota "
    Set oConecta = New DConecta
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
Dim oConecta As DConecta

    On Error GoTo ErrRecuperaCalendarioGastosPendientes
            
    sSql = " Select CD.* From ColocCalendario C Inner Join ColocCalendDet CD "
    sSql = sSql & " ON C.cCtaCod = CD.cCtaCod AND C.nNroCalen = CD.nNroCalen AND C.nColocCalendApl = CD.nColocCalendApl "
    sSql = sSql & " AND C.nCuota = CD.nCuota "
    sSql = sSql & " Where  C.cCtaCod = '" & psCtaCod & "' AND  C.nColocCalendApl = " & pnAplicado & " AND C.nColocCalendEstado = 0 "
    sSql = sSql & " AND CD.nPrdConceptoCod like '12%' AND C.nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = '" & psCtaCod & "') Order By C.nCuota   "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioGastosPendientes = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrRecuperaCalendarioGastosPendientes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaCalendarioGastosPendientesAgrupado(ByVal psCuentas As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioGastosPendientesAgrupado
            
    sSql = " Select CD.nPrdConceptoCod, SUM(CD.nMonto - CD.nMontoPagado) as nMonto From ColocCalendario C Inner Join ColocCalendDet CD "
    sSql = sSql & " ON C.cCtaCod = CD.cCtaCod AND C.nNroCalen = CD.nNroCalen AND C.nColocCalendApl = CD.nColocCalendApl "
    sSql = sSql & " AND C.nCuota = CD.nCuota AND C.nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = C.cCtaCod ) "
    sSql = sSql & " Where  C.cCtaCod In  " & psCuentas & " AND  C.nColocCalendEstado = 0 "
    sSql = sSql & " AND CD.nPrdConceptoCod like '12%' AND C.nColocCalendApl = 1 "
    sSql = sSql & " Group By CD.nPrdConceptoCod "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioGastosPendientesAgrupado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaCalendarioGastosPendientesAgrupado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Function RecuperaCalendarioDesemb(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    sSql = "select C.dVenc,C.dPago, C.nCuota,C.nColocCalendEstado, "
    sSql = sSql & " nCapital=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1000),"
    sSql = sSql & " nIntComp=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1100),"
    sSql = sSql & " nIntGracia=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1102),"
    sSql = sSql & " nIntMor=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1101),"
    sSql = sSql & " nIntReprog=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1103),"
    sSql = sSql & " nGasto=(select sum(nMonto) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where nPrdConceptoCod like '12%' )) "
    sSql = sSql & " from ColocCalendario C inner join ColocacCred CC on C.cCtaCod=CC.cCtaCod "
    sSql = sSql & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl = " & gColocCalendAplDesembolso & " and C.nNroCalen = CC.nNroCalen order by C.nCuota"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioDesemb = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaCalendarioDesembPagados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    sSql = "select C.dVenc,C.nCuota,C.nColocCalendEstado, "
    sSql = sSql & " nCapital=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1000),"
    sSql = sSql & " nIntComp=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1100),"
    sSql = sSql & " nIntGracia=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1102),"
    sSql = sSql & " nIntMor=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1101),"
    sSql = sSql & " nIntReprog=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1103),"
    sSql = sSql & " nGasto=(select sum(nMonto) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where nPrdConceptoCod like '12%' )) "
    sSql = sSql & " from ColocCalendario C inner join ColocacCred CC on C.cCtaCod=CC.cCtaCod "
    sSql = sSql & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl = " & gColocCalendAplDesembolso & " and C.nNroCalen = CC.nNroCalen AND C.nColocCalendEstado=1 order by C.nCuota"
    Set oConecta = New DConecta
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
Dim oConecta As DConecta

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
    Set oConecta = New DConecta
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
Dim oConecta As DConecta

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
    
    Set oConecta = New DConecta
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
Dim oCon As DConecta
Dim sSql As String
Dim R As ADODB.Recordset

    sSql = "Select Count(*) as nCuotas From  ColocCalendario Where nNroCalen = " & pnNroCalen
    sSql = sSql & " AND nColocCalendApl = " & pnAplicado & " AND cCtaCod = '" & psCtaCod & "'"
    
    Set oCon = New DConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    RecuperaNroCuotas = IIf(IsNull(R!nCuotas), 0, R!nCuotas)
    R.Close
    Set R = Nothing
End Function

Public Function RecuperaCalendarioPagosPendiente(ByVal psCtaCod As String, Optional ByVal pbCalParalelo As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    sSql = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
    sSql = sSql & " nMontoPrestamo = (select SUM(nMonto) from ColocCalendDet CD2 Inner Join ColocCalendario C2 ON CD2.cCtaCod=C2.cCtaCod AND CD2.nNroCalen=C2.nNroCalen AND CD2.nColocCalendApl = C2.nColocCalendApl AND CD2.nCuota = C2.nCuota where CD2.cCtaCod = C.cCtaCod And CD2.nNroCalen = C.nNroCalen And CD2.nColocCalendApl=C.nColocCalendApl and CD2.nPrdConceptoCod = 1000 AND C2.nColocCalendEstado = " & gColocCalendEstadoPendiente & "), "
    sSql = sSql & " nSaldoCap=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
    sSql = sSql & " nCapital=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
    sSql = sSql & " nIntComp=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSql = sSql & " nIntCompVenc=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompVencido & "),"
    sSql = sSql & " nIntGracia=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSql = sSql & " nIntMor=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSql = sSql & " nIntReprog=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSql = sSql & " nIntSuspenso=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
    sSql = sSql & " nGasto=(select sum(nMonto - nMontoPagado) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' )) "
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
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioPagosPendiente = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaRefinanciados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaMatrizRefinanciados
    sSql = "Select CR.cCtaCod,CR.cCtaCodRef,CR.nMontoRef, "
    sSql = sSql & "nCapital = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
    sSql = sSql & "nCapitalPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
    sSql = sSql & "nIntComp = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSql = sSql & "nIntCompPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSql = sSql & "nIntGracia = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSql = sSql & "nIntGraciaPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSql = sSql & "nIntMor = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSql = sSql & "nIntMorPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSql = sSql & "nIntReprog = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSql = sSql & "nIntReprogPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSql = sSql & "nIntSuspenso = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
    sSql = sSql & "nIntSuspensoPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "), "
    sSql = sSql & "nGastos = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodGastoVarios & "),"
    sSql = sSql & "nGastosPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nPrdConceptoCod = " & gColocConceptoCodGastoVarios & ") "
    sSql = sSql & "From ColocacRefinanc CR Where cCtaCod = '" & psCtaCod & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaRefinanciados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaMatrizRefinanciados:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
    
End Function

Public Function RecuperaFechaInicioCuota(ByVal psCtaCod As String, ByVal pnCuota As Integer, ByVal pnAplicado As ColocCalendApl) As Date
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaFechaInicioCuota
    If pnCuota > 1 Then
        sSql = "Select CC.dVenc from ColocCalendario CC "
        sSql = sSql & " Where CC.nCuota = " & pnCuota - 1 & " And nColocCalendApl = " & pnAplicado & " AND cCtaCod = '" & psCtaCod & "'"
        sSql = sSql & " And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CC.cCtaCod)"
    Else
        sSql = "Select dVigencia as dVenc from Colocaciones Where cCtaCod = '" & psCtaCod & "'"
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    RecuperaFechaInicioCuota = Format(R!dVenc, "dd/mm/yyyy")
    Exit Function

ErrorRecuperaFechaInicioCuota:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaCalendarioPagos(ByVal psCtaCod As String, _
    Optional ByVal pnNroCalen As Integer = -1, Optional ByVal pbNroCalenParalelo As Boolean = False, _
    Optional ByVal pbComItemFE As Boolean = False, Optional ByVal pbInicial As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
        If pbComItemFE Then
            sSql = "select '',C.dVenc,C.nCuota, C.nColocCalendEstado,"
        Else
            sSql = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
        End If
        sSql = sSql & " nCapital=ISNULL((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodCapital & "),0),"
        sSql = sSql & " nCapitalPag=ISNULL((select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodCapital & "),0),"
        sSql = sSql & " nIntComp=ISNULL((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),0),"
        sSql = sSql & " nIntCompPag=ISNULL((select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),0),"
        sSql = sSql & " nIntGracia=ISNULL((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),0),"
        sSql = sSql & " nIntGraciaPag=ISNULL((select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresGracia & "),0),"
        sSql = sSql & " nIntMor=ISNULL((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),0),"
        sSql = sSql & " nIntMorPag=ISNULL((select ISNULL(nMontoPagado,0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & "),0),"
        sSql = sSql & " nIntReprog=ISNULL((select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),0),"
        sSql = sSql & " nIntReprogPag=ISNULL((select ISNULL(nMontoPagado,0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado & "),0),"
        sSql = sSql & " nIntSuspenso=ISNULL((select ISNULL(nMonto,0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),0),"
        sSql = sSql & " nIntSuspensoPag=ISNULL((select ISNULL(nMontoPagado,0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresSuspenso & "),0),"
        If pbInicial Then
            sSql = sSql & " nGasto=ISNULL((select ISNULL(sum(nMonto),0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' AND cAplicaProceso = 'DE' )),0),"
        Else
            sSql = sSql & " nGasto=ISNULL((select ISNULL(sum(nMonto),0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' )),0),"
        End If
        sSql = sSql & " nGastoPag=ISNULL((select ISNULL(sum(nMontoPagado),0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (Select nPrdConceptoCod from ProductoConcepto Where convert(varchar(6),nPrdConceptoCod) like '12%' )),0),"
        sSql = sSql & " nIntCompVenc=ISNULL((select ISNULL(nMonto,0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompVencido & "),0),"
        sSql = sSql & " nIntCompVencPag=ISNULL((select ISNULL(nMontoPagado,0) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = " & gColocConceptoCodInteresCompVencido & "),0) "
        sSql = sSql & " from ColocCalendario C "
        
            
    If Not pbNroCalenParalelo Then
        If pnNroCalen = -1 Then
            sSql = sSql & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = C.cCtaCod) order by C.nCuota"
        Else
            sSql = sSql & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = " & pnNroCalen & " order by C.nCuota"
        End If
    Else
        sSql = sSql & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalPar from ColocacCred where cCtaCod = C.cCtaCod) order by C.nCuota"
    End If
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioPagos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function



Public Function RecuperaColocCalendario(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, ByVal pnAplicado As ColocCalendApl, Optional ByVal pnNroCuota As Integer = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaColocCalendario
    sSql = "Select * from ColocCalendario where cCtacod = '" & psCtaCod & "' AND nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnAplicado
    If pnNroCuota <> -1 Then
        sSql = sSql & " AND nCuota = " & pnNroCuota
    End If
    sSql = sSql & "  order by nCuota "
    Set oConecta = New DConecta
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
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaColocCalendDet
    sSql = "Select * from ColocCalendDet where cCtacod = '" & psCtaCod & "' And nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnAplicado
    Set oConecta = New DConecta
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
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaGastosCuotaDesemb
    sSql = "Select C.nCuota, C.nMonto, C.nMontoPagado, CN.cDescripcion cGasto, C.nPrdConceptoCod"
    sSql = sSql & " From ColocCalendDet C Inner Join ProductoConcepto CN ON C.nPrdConceptoCod = CN.nPrdConceptoCod AND CN.nPrdConceptoCod like '" & Mid(Trim(Str(gColocConceptoCodGastoVarios)), 1, 2) & "%'"
    sSql = sSql & " Where cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnAplicado
    If pnCuota <> -1 Then
        sSql = sSql & " AND nCuota = " & pnCuota
    End If
    Set oConecta = New DConecta
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
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorCuotasPendientes
    sSql = "Select count(*) as nCuotasPend from ColocCalendario CC where cCtaCod = '" & psCtaCod & "' "
    sSql = sSql & " AND nColocCalendApl = " & gColocCalendAplCuota & " AND nColocCalendEstado = " & gColocCalendEstadoPendiente & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = CC.cCtaCod)"
    Set oConecta = New DConecta
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

