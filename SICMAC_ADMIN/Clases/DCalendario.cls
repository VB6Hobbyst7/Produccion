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
                    ByVal pnNroCuota As Integer, ByVal pnAplicado As ColocCalendApl) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioGastos
    sSql = "Select * from ColocCalendDet where cCtaCod = '" & psCtaCod & _
            "' AND nNroCalen = " & pnNroCalen & " AND nCuota = " & pnNroCuota & " AND nColocCalendApl = " & pnAplicado & _
            " AND nColocConceptoCod like '12%'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioGastos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaCalendarioGastos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaCalendarioDesemb(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    sSql = "select C.dVenc,C.nCuota,C.nColocCalendEstado, "
    sSql = sSql & " nCapital=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = 1000),"
    sSql = sSql & " nIntComp=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = 1100),"
    sSql = sSql & " nIntGracia=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = 1102),"
    sSql = sSql & " nIntMor=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = 1101),"
    sSql = sSql & " nIntReprog=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = 1103),"
    sSql = sSql & " nGasto=(select sum(nMonto) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod in (Select nColocConceptoCod from ColocConcepto Where nColocConceptoCod like '12%' )) "
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

Public Function RecuperaCalendarioPagosDeuda(ByVal psCtaCod As String, Optional ByVal pnNroCalen As Integer = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    sSql = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
    sSql = sSql & " nCapital=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodCapital & "),"
    sSql = sSql & " nIntComp=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSql = sSql & " nIntGracia=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSql = sSql & " nIntMor=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSql = sSql & " nIntReprog=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSql = sSql & " nIntSuspenso=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
    sSql = sSql & " nGasto=(select sum(nMonto- nMontoPagado) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod in (Select nColocConceptoCod from ColocConcepto Where convert(varchar(6),nColocConceptoCod) like '12%' ))"
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
    sSql = sSql & " nCapital=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodCapital & "),"
    sSql = sSql & " nIntComp=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSql = sSql & " nIntGracia=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSql = sSql & " nIntMor=(select nMontoPagado  from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSql = sSql & " nIntReprog=(select nMontoPagado  from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSql = sSql & " nIntSuspenso=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
    sSql = sSql & " nGasto=(select sum(nMontoPagado) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod in (Select nColocConceptoCod from ColocConcepto Where convert(varchar(6),nColocConceptoCod) like '12%' )), "
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


Public Function RecuperaCalendarioPagosPendiente(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    sSql = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
    sSql = sSql & " nSaldoCap=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodCapital & "),"
    sSql = sSql & " nCapital=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodCapital & "),"
    sSql = sSql & " nIntComp=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSql = sSql & " nIntGracia=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSql = sSql & " nIntMor=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSql = sSql & " nIntReprog=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSql = sSql & " nIntSuspenso=(select nMonto - nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
    sSql = sSql & " nGasto=(select sum(nMonto - nMontoPagado) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod in (Select nColocConceptoCod from ColocConcepto Where convert(varchar(6),nColocConceptoCod) like '12%' )) "
    sSql = sSql & " from ColocCalendario C "
    sSql = sSql & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = C.cCtaCod) AND C.nColocCalendEstado = " & gColocCalendEstadoPendiente & " order by C.nCuota"
    
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
    sSql = sSql & "nCapital = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nColocConCeptoCod = " & gColocConceptoCodCapital & "),"
    sSql = sSql & "nCapitalPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nColocConCeptoCod = " & gColocConceptoCodCapital & "),"
    sSql = sSql & "nIntComp = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nColocConCeptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSql = sSql & "nIntCompPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nColocConCeptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSql = sSql & "nIntGracia = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nColocConCeptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSql = sSql & "nIntGraciaPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nColocConCeptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSql = sSql & "nIntMor = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nColocConCeptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSql = sSql & "nIntMorPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nColocConCeptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSql = sSql & "nIntReprog = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nColocConCeptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSql = sSql & "nIntReprogPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nColocConCeptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSql = sSql & "nIntSuspenso = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nColocConCeptoCod = " & gColocConceptoCodInteresSuspenso & "),"
    sSql = sSql & "nIntSuspensoPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nColocConCeptoCod = " & gColocConceptoCodInteresSuspenso & "), "
    sSql = sSql & "nGastos = (Select nMonto From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nColocConCeptoCod = " & gColocConceptoCodGastoVarios & "),"
    sSql = sSql & "nGastosPag = (Select nMontoPagado From ColocacRefinancDet Where cCtaCod=cr.cCtaCod AND cCtaCodRef=cr.cCtaCodRef AND nEstado=cr.nEstado AND dEstado=cr.dEstado AND nColocConCeptoCod = " & gColocConceptoCodGastoVarios & ") "
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

Public Function RecuperaCalendarioPagos(ByVal psCtaCod As String, Optional ByVal pnNroCalen As Integer = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    sSql = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
    sSql = sSql & " nCapital=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodCapital & "),"
    sSql = sSql & " nCapitalPag=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodCapital & "),"
    sSql = sSql & " nIntComp=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSql = sSql & " nIntCompPag=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresCompensatorio & "),"
    sSql = sSql & " nIntGracia=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSql = sSql & " nIntGraciaPag=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresGracia & "),"
    sSql = sSql & " nIntMor=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSql = sSql & " nIntMorPag=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresMoratorio & "),"
    sSql = sSql & " nIntReprog=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSql = sSql & " nIntReprogPag=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresReprogramado & "),"
    sSql = sSql & " nIntSuspenso=(select nMonto from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
    sSql = sSql & " nIntSuspensoPag=(select nMontoPagado from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod = " & gColocConceptoCodInteresSuspenso & "),"
    sSql = sSql & " nGasto=(select sum(nMonto) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod in (Select nColocConceptoCod from ColocConcepto Where convert(varchar(6),nColocConceptoCod) like '12%' )),"
    sSql = sSql & " nGastoPag=(select sum(nMontoPagado) from ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nColocConceptoCod in (Select nColocConceptoCod from ColocConcepto Where convert(varchar(6),nColocConceptoCod) like '12%' ))"
    sSql = sSql & " from ColocCalendario C "
    If pnNroCalen = -1 Then
        sSql = sSql & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = (select nNroCalen from ColocacCred where cCtaCod = C.cCtaCod) order by C.nCuota"
    Else
        sSql = sSql & " where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= " & gColocCalendAplCuota & " And nNroCalen = " & pnNroCalen & " order by C.nCuota"
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
    sSql = "Select * from ColocCalendDet where cCtacod = '" & psCtaCod & "' And nNroCalen = " & pnNroCalen & " nColocCalendApl = " & pnAplicado
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocCalendDet = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocCalendDet:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaGastosCuotaDesemb(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, ByVal pnAplicado As ColocCalendApl, Optional ByVal pnCuota As Integer = -1) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaGastosCuotaDesemb
    sSql = "Select C.nCuota, C.nMonto, C.nMontoPagado, CN.cDescripcion cGasto, C.nColocConceptoCod"
    sSql = sSql & " From ColocCalendDet C Inner Join ColocConcepto CN ON C.nColocConceptoCod = CN.nColocConceptoCod AND CN.nColocConceptoCod like '" & Mid(Trim(Str(gColocConceptoCodGastoVarios)), 1, 2) & "%'"
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

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing


End Sub
