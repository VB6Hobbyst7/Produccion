VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function ObtenerCreditosPorCliente(ByVal psCodCliente As String) As ADODB.Recordset
    Dim objC As DConecta
    Dim strSQL As String
    On Error GoTo ErrHandler
        strSQL = ""
        Set objC = New DConecta
            objC.AbreConexion
            
    Exit Function
ErrHandler:
    If Not objcn Is Nothing Then Set objC = Nothing
    Set ObtenerCreditosPorCliente = Nothing
End Function

'Public Function VerificaCreditoRFA(ByVal psCtaCod As String) As Boolean
'    Dim objC As DConecta
'    Dim rs As ADODB.Recordset
'    Dim strSQL As String
'    Dim intCantidad As Integer
'    On Error GoTo ErrHandler
'        strSQL = "Select Count(*) as Cantidad From ColocacCred Where cCtaCod='" & psCtaCod & "' "
'        strSQL = strSQL & " AND cRFA='RFC'"
'
'        Set objC = New DConecta
'        objC.AbreConexion
'        Set rs = objC.CargaRecordSet(strSQL)
'        objC.CierraConexion
'        Set objC = Nothing
'
'        If Not rs.EOF And Not rs.BOF Then
'            intCantidad = rs!cantidad
'        End If
'        Set rs = Nothing
'        If intCantidad = 0 Then
'            VerificaCreditoRFA = False
'        Else
'            VerificaCreditoRFA = True
'        End If
'
'    Exit Function
'ErrHandler:
'    If Not objC Is Nothing Then Set objC = Nothing
'    VerificaCreditoRFA = False
'End Function

Public Function BuscaDNICliente(ByVal psCtaCod As String) As String
    Dim objC As DConecta
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrHandler
    strSQL = "Select cPersIDnro from persid where cperscod='" & psCtaCod & "'"
    Set objC = New DConecta
    objC.AbreConexion
    Set rs = objC.CargaRecordSet(strSQL)
    objC.CierraConexion
    Set objC = Nothing
    If Not rs.EOF And Not rs.BOF Then
        BuscaDNICliente = rs!cPersIdNro
    Else
        BuscaDNICliente = ""
    End If
    Set rs = Nothing
    Exit Function
ErrHandler:
    If Not objC Is Nothing Then Set objC = Nothing
    If Not rs Is Nothing Then Set rs = Nothing
    BuscaDNICliente = ""
End Function

Public Function VerificaCred(ByVal pcCtaCodCliente) As Integer
'1 es verdad
'0 es falso
'-1 error
'Verifica si el cliente posee creditos de rfa
    Dim objC As DConecta
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrHandler
    
        Set objC = New DConecta
        'strSQL = "Select Count(*) as Cantidad From ColocacCred where cRFA in('RFC','RFC','DIF') where cctacod='" & pcCtaCodCliente & "'"
        strSQL = "Select Count(P.cCtaCod) as Cantidad"
        strSQL = strSQL & " from Producto P"
        strSQL = strSQL & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod"
        strSQL = strSQL & " Inner Join colocaccred CC on PP.cCtaCod=CC.cCtaCod"
        strSQL = strSQL & " Where PP.cPersCod='" & pcCtaCodCliente & "' and PP.nPrdPersRelac=20 and"
        strSQL = strSQL & " CC.cRFA IN('RFC','RFA','DIF') and P.nPrdEstado<>2050"
        objC.AbreConexion
        Set rs = objC.CargaRecordSet(strSQL)
        objC.CierraConexion
        Set objC = Nothing
        If Not rs.EOF And Not rs.BOF Then
           If rs!cantidad > 0 Then
            VerificaCred = 1
           Else
            VerificaCred = 0
           End If
        Else
         GoTo ErrHandler
        End If
        Set rs = Nothing
    Exit Function
ErrHandler:
    If Not objC Is Nothing Then Set objC = Nothing
    If Not rs Is Nothing Then Set rs = Nothing
    verificarcred = -1
End Function

Function ListaCreditosPendientes(ByVal pcCodCli As String, ByVal pdFecSis As String) As ADODB.Recordset
      Dim objC As DConecta
      Dim sSQL As String
      On Error GoTo ErrHandler
      sSQL = "set dateformat dmy"
      sSQL = sSQL & " Select CC.cRFA as cCredito,CO.nCuota as nCuota,PP.cCtaCod,"
      sSQL = sSQL & " Capital=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in (" & gColocConceptoCodCapital & "," & gColocRFACapital & ") and CD.nCuota=CO.nCuota),"
      sSQL = sSQL & " IntComp=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in (" & gColocConceptoCodInteresCompensatorio & "," & gColocRFAInteresCompesatorio & ") and CD.nCuota=CO.nCuota),"
      sSQL = sSQL & " IntMor=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in (" & gColocConceptoCodInteresMoratorio & "," & gColocRFAInteresMoratorio & ") and CD.nCuota=CO.nCuota),"
      sSQL = sSQL & " Gastos=(Select isnull(Sum(nMonto-nMontoPagado),0) From coloccalenddet a  where a.cctaCod=PP.cCtaCod and a.nColocCalendApl=1 and a.nCuota=CO.nCuota and a.nPrdConceptoCod Like '12%' and a.nPrdConceptoCod not in('124352','124353','124354','124355')),"
      sSQL = sSQL & " ComCofide=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nCuota=CO.nCuota and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod=124352 and CD.nCuota=CO.nCuota), CO.dVenc, diasAtraso=DateDiff(day,dVenc,'" & pdFecSis & "'),"
      sSQL = sSQL & " Indice=Case cc.cRFA    When 'DIF' Then 1    When 'RFC' Then 2    When 'RFA' Then 3 End "
      sSQL = sSQL & " From ProductoPersona PP"
      sSQL = sSQL & " Inner Join ColocCalendario CO on PP.cCtaCod=CO.cCtaCod"
      sSQL = sSQL & " Inner Join ColocacCred CC on CC.cCtaCod=CO.cCtaCod"
      sSQL = sSQL & " Where PP.cPersCod='" & pcCodCli & "' and  CC.cRFA IN('RFC','DIF','RFA') and"
      sSQL = sSQL & " CO.nColocCalendApl=1 and"
      sSQL = sSQL & " CO.nCuota=(Select Min(nCuota) From ColocCalendario c1"
      sSQL = sSQL & " inner join ColocacCred d on c1.cCtaCod=d.cCtaCod"
      sSQL = sSQL & " inner join ProductoPersona Ps on Ps.cCtaCod=C1.cCtaCod and Ps.nPrdPersRelac=20"
      sSQL = sSQL & " Where Ps.cPersCod='" & pcCodCli & "' and c1.nColocCalendApl=CO.nColocCalendApl and c1.nColocCalendEstado=0"
      sSQL = sSQL & " and d.cRFA IN('RFC','DIF','RFA'))"
      sSQL = sSQL & " Order by nCuota,Indice"
    
    Set objC = New DConecta
    objC.AbreConexion
    Set ListaCreditosPendientes = objC.CargaRecordSet(sSQL)
    objC.CierraConexion
    Set objC = Nothing
      
    Exit Function
ErrHandler:
      If Not objC Is Nothing Then Set objC = Nothing
      Set ListaCreditosPendientes = Null
End Function

Function ListaCreditosPendientesCuota(ByVal pcCodCli As String, ByVal pdFecSis As String, ByVal pnCuota As Integer) As ADODB.Recordset
      Dim objC As DConecta
      Dim sSQL As String
      On Error GoTo ErrHandler
                       
      sSQL = "set dateformat dmy"
      sSQL = sSQL & " Select CC.cRFA as cCredito,CO.nCuota as nCuota,PP.cCtaCod,"
      sSQL = sSQL & " Capital=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in(1000,1010) and CD.nCuota=CO.nCuota),"
      sSQL = sSQL & " IntComp=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in(1100,1107) and CD.nCuota=CO.nCuota),"
      sSQL = sSQL & " IntMor=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in(1101,1108) and CD.nCuota=CO.nCuota),"
      sSQL = sSQL & " Gastos=(Select isnull(Sum(nMonto-nMontoPagado),0) From coloccalenddet a  where a.cctaCod=PP.cCtaCod and a.nColocCalendApl=1 and a.nCuota=CO.nCuota and a.nPrdConceptoCod Like '12%' and a.nPrdConceptoCod not in('124352','124353','124354','124355')),"
      sSQL = sSQL & " ComCofide=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nCuota=CO.nCuota and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod=124352 and CD.nCuota=CO.nCuota), CO.dVenc, diasAtraso=DateDiff(day,dVenc,'" & pdFecSis & "'),"
      sSQL = sSQL & " Indice=Case cc.cRFA    When 'DIF' Then 1    When 'RFC' Then 2    When 'RFA' Then 3 End"
      sSQL = sSQL & " From ProductoPersona PP"
      sSQL = sSQL & " Inner Join ColocCalendario CO on PP.cCtaCod=CO.cCtaCod"
      sSQL = sSQL & " Inner Join ColocacCred CC on CC.cCtaCod=CO.cCtaCod"
      sSQL = sSQL & " Where PP.cPersCod='" & pcCodCli & "' and  CC.cRFA IN('RFC','DIF','RFA') and"
      sSQL = sSQL & " CO.nColocCalendApl=1 and"
      sSQL = sSQL & " CO.nCuota=" & pnCuota
      sSQL = sSQL & " Order by nCuota,Indice"
    
    Set objC = New DConecta
    objC.AbreConexion
    Set ListaCreditosPendientesCuota = objC.CargaRecordSet(sSQL)
    objC.CierraConexion
    Set objC = Nothing
      
    Exit Function
ErrHandler:
      If Not objC Is Nothing Then Set objC = Nothing
      Set ListaCreditosPendientesCuota = Null
End Function


Function ObtenerFactor(ByVal psCtaCod As String) As Double
    Dim oConec As DConecta
    Dim sSQL As String
    Dim rs As ADODB.Recordset
        sSQL = "Select nRfaFactor From FactorRfa "
        sSQL = sSQL & " where cctacod='" & psCtaCod & "'"
        
        Set oConec = New DConecta
        oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sSQL)
        If Not rs.EOF And Not rs.BOF Then
            ObtenerFactor = rs!nRfaFactor
        End If
        Set rs = Nothing
        oConec.CierraConexion
        Set oConec = Nothing
    On Error GoTo ErrHandler
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    If Not rs Is Nothing Then Set rs = Nothing
End Function

Public Function BuscarPersona(ByVal pcPersCod As String) As Recordset
    Dim oConec As DConecta
    Dim sSQL As String
    
    sSQL = "Select cPersCod,cPersNombre"
    sSQL = sSQL & " From Persona"
    sSQL = sSQL & " Where cPersCod='" & pcPersCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set BuscarPersona = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerExtornoRFA(ByVal psPersCod As String, ByVal psUsuario As String, _
       ByVal pdFecha As Date, ByVal psCodAgencia As String) As Recordset
    
    Dim oConec As DConecta
    Dim sSQL As String
    Dim sFecha As String
    
    On Error GoTo ErrHandler
    
    'Convirtiendo la fecha
    sFecha = Mid(CStr(pdFecha), 7, 4) & Mid(CStr(pdFecha), 4, 2) & Mid(CStr(pdFecha), 1, 2)
    sSQL = "Select PP.cCtaCod,"
    sSQL = sSQL & " MC.cOpeCod,"
    sSQL = sSQL & " Substring(M.cMovNro,8,2)+':'+Substring(M.cMovNro,11,2)+ ':' +Substring(M.cMovNro,13,2) as cHora,"
    sSQL = sSQL & " MAX(MC.nMovNro)as nMovNro,"
    sSQL = sSQL & " Sum(MC.nMonto) as nMonto,"
    sSQL = sSQL & " right(M.cMovNro,4) as cUsuario,"
    sSQL = sSQL & " MC.nPrePago,"
    sSQL = sSQL & " CC.cRFA"
    sSQL = sSQL & " From Mov M"
    sSQL = sSQL & " Inner Join  MovCol MC on MC.nMovNro=M.nMovNro"
    sSQL = sSQL & " Inner Join ColocacCred CC on MC.cCtaCod=CC.cCtacod"
    sSQL = sSQL & " Inner Join ProductoPersona PP on PP.cCtaCod=MC.cCtaCod"
    sSQL = sSQL & " Where Left(M.cMovNro,8)='" & sFecha & "' and right(M.cMovNro,4)='" & psUsuario & "' and M.nMovFlag<>2  and MC.cOpeCod not like '109%' and "
    sSQL = sSQL & " CC.cRFA in ('DIF','RFC','RFA') and PP.nPrdPersRelac=20 and PP.cPersCod='" & psPersCod & "'"
    sSQL = sSQL & " and Substring(CC.cCtaCod,4,2)='" & psCodAgencia & "'"
    sSQL = sSQL & " Group By PP.cCtaCod,MC.cOpeCod,Substring(M.cMovNro,8,2)+':'+Substring(M.cMovNro,11,2)+ ':' +Substring(M.cMovNro,13,2),"
    sSQL = sSQL & " right(M.cMovNro,4),MC.nPrePago,CC.cRFA"
    sSQL = sSQL & " Having Sum(MC.nMonto) > 0"
    sSQL = sSQL & " Order By MAX(MC.nMovNro),MC.nPrePago,CC.cRFA"
    
    Set oConec = New DConecta
    
    oConec.AbreConexion
    Set ObtenerExtornoRFA = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ObtenerExtornoRFA = Null
    MsgBox "Se ha producido un error  " & Err.Description, vbInformation, "AVISO"
    
End Function

Public Function ObtenerCuentasExtorno(ByVal psPersCod As String, ByVal psUsuario As String, _
       ByVal pdFecha As Date, ByVal psCodAgencia As String) As Recordset
 Dim sSQL As String
 Dim oConec As DConecta
 Dim sFecha As String
    On Error GoTo ErrHandler
        
        'Convirtiendo la fecha
    sFecha = Mid(CStr(pdFecha), 7, 4) & Mid(CStr(pdFecha), 4, 2) & Mid(CStr(pdFecha), 1, 2)
    
    sSQL = sSQL & "Select PP.cCtaCod,"
    sSQL = sSQL & " Max(MC.nMovNro) As nMovNro,"
    sSQL = sSQL & " Sum(MC.nMonto) as nMonto"
    sSQL = sSQL & " From Mov M"
    sSQL = sSQL & " Inner Join  MovCol MC on MC.nMovNro=M.nMovNro"
    sSQL = sSQL & " Inner Join ColocacCred CC on MC.cCtaCod=CC.cCtacod"
    sSQL = sSQL & " Inner Join ProductoPersona PP on PP.cCtaCod=MC.cCtaCod"
    sSQL = sSQL & " Where Left(M.cMovNro,8)='20040930' and right(M.cMovNro,4)='VDGT' and M.nM vFlag<>2 and"
    sSQL = sSQL & " CC.cRFA in ('DIF','RFC','RFA') and PP.nPrdPersRelac=20 and PP.cPersCod='1080100748200' and"
    sSQL = sSQL & " Substring(CC.cCtaCod,4,2)='01'"
    sSQL = sSQL & " Group By PP.cCtaCod"
    sSQL = sSQL & " Order By Max(MC.nMovNro)"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ObtenerCuentasExtorno = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If oConec Is Nothing Then Set oConec = Nothing
    Set ObtenerCuentasExtorno = Null
End Function

Public Sub ExtornarPago(ByVal pnMovNro As Long, ByVal pdFecSis As Date, ByVal psCodUser As String, ByVal psCodAge As String)
Dim R As ADODB.Recordset
Dim oDCreditos As DCreditos
Dim oBase As DCredActualizaBD
Dim sMovNro As String
Dim nMovNro As Long
Dim oFun As NContFunciones
    Set oFun = New NContFunciones
    sMovNro = oFun.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
    Set oFun = Nothing
    
    Set oDCreditos = New DCreditos
    Set R = RecuperaCreditosdePago(pnMovNro)
    Set oBase = New DCredActualizaBD
    oBase.dBeginTrans
    Call oBase.dUpdateMov(pnMovNro, , , , gMovFlagExtornado, False)
    Call oBase.dInsertMov(sMovNro, gCredExtPagoLote, "Extorno de Pago", gMovEstContabNoContable, gMovFlagDeExtorno, False)
    nMovNro = oBase.dGetnMovNro(sMovNro)
    
    Do While Not R.EOF
        Call ExtornarCredito(R!Cctacod, pdFecSis, psCodUser, psCodAge, pnMovNro, R!copecod, R!nMonto, , True, oBase, nMovNro)
        R.MoveNext
    Loop
    oBase.dCommitTrans
    Set oBase = Nothing
    R.Close
    Set R = Nothing
End Sub

Public Function RecuperaCreditosdePago(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    sSQL = "Select cCtaCod, cOpeCod, nMonto From MovCol Where nMovNro = " & pnMovNro
    sSQL = sSQL & " and nMonto>0 and cOpeCod not like '99%'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosdePago = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Sub ExtornarCredito(ByVal psCtaCod As String, ByVal pdFecSis As Date, ByVal psCodUser As String, _
    ByVal psCodAge As String, ByVal pnMov As Long, psCodOpe As String, ByVal pnMonto As Double, Optional ByVal pnNroDacion As Long = -1, _
    Optional ByVal pbExtornoEnLote As Boolean = False, Optional ByVal poBase As DCredActualizaBD = Nothing, Optional ByVal pnMovNro As Long = -1, _
    Optional ByVal pnPrepago As Integer = -1)
    
Dim oDCred As DCredito
Dim oBase As DCredActualizaBD
Dim oFun As NContFunciones
Dim nMovNro As Long
Dim sMovNro, sMovNroCap As String
Dim MatsMovNro() As String
Dim R, RCap, RCapMov, RCancel As ADODB.Recordset
Dim dCap As Date
Dim sSQL As String
Dim nCuotaMin As Integer
Dim nNroCalen As Integer
Dim nNroCalenPar As Integer
Dim i As Integer
Dim nMontoCond As Double
Dim nMovNroCond As Long

    'Set oDCred = New DCredito
    Set R = RecuperaDatosExtorno(pnMov, psCtaCod)
    If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA _
        Or psCodOpe = gCredPagNorNorCC Or psCodOpe = gCredPagNorMorCC Or psCodOpe = gCredPagNorVenCC Or psCodOpe = gCredPagRefNorCC Or psCodOpe = gCredPagRefMorCC Or psCodOpe = gCredPagRefVenCC Then
        If psCodOpe = gCredPagNorNorCC Or psCodOpe = gCredPagNorMorCC Or psCodOpe = gCredPagNorVenCC Or psCodOpe = gCredPagRefNorCC Or psCodOpe = gCredPagRefMorCC Or psCodOpe = gCredPagRefVenCC Then
            Set RCapMov = oDCred.RecuperaMovimientosAhorros(pnMov, True)
        Else
            Set RCapMov = oDCred.RecuperaMovimientosAhorros(pnMov)
        End If
    End If
    'Set oDCred = Nothing
    
    If pbExtornoEnLote Then
        Set oBase = poBase
    Else
        Set oFun = New NContFunciones
        sMovNro = oFun.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
        Set oFun = Nothing
        Set oBase = New DCredActualizaBD
        ReDim MatsMovNro(0)
        If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Then
            
            Set oFun = New NContFunciones
            Do While Not RCapMov.EOF
                'If RCapMov!cOpeCod = "200301" Then '********* Aca hice la modificacion CAFF
                If RCapMov!copecod = gCredDesembCtaRetiroGastos Or RCapMov!copecod = gCredDesembCtaRetiroGastosDOA Or RCapMov!copecod = gCredDesembCtaRetiroCancelacion Or RCapMov!copecod = gCredDesembCtaRetiroCancelacionDOA Then
                    Sleep 1000
                    ReDim Preserve MatsMovNro(RCapMov.Bookmark)
                    MatsMovNro(RCapMov.Bookmark - 1) = oFun.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
                End If
                RCapMov.MoveNext
            Loop
            RCapMov.MoveFirst
            Set oFun = Nothing
        End If
        Call oBase.dBeginTrans
    End If
    
    Sleep 1000
    Set oFun = New NContFunciones
    If pbExtornoEnLote Then
        sMovNroCap = poBase.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
    Else
        sMovNroCap = oFun.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
    End If
    Set oFun = Nothing
    'Extorna Movimientos
    If Not pbExtornoEnLote Then
        Call oBase.dUpdateMov(pnMov, , , , gMovFlagExtornado, False)
        Call oBase.dInsertMov(sMovNro, gCredExtPago, "Extorno de Pago", gMovEstContabNoContable, gMovFlagDeExtorno, False)
        nMovNro = oBase.dGetnMovNro(sMovNro)
    Else
        nMovNro = pnMovNro
    End If
    
    'Extorna Movimiento de Colocaciones Insertando sus Filas Correspondientes
    If Not pbExtornoEnLote Then
        Call oBase.dInsertMovCol(nMovNro, gCredExtPago, psCtaCod, 0, pnMonto, 0, "", 0, 0, 0, False)
    Else
        Call oBase.dInsertMovCol(nMovNro, gCredExtPagoLote, psCtaCod, 0, pnMonto, 0, "", 0, 0, 0, False)
    End If
    
    
    If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Or psCodOpe = gCredDesembEfec Or psCodOpe = gCredPagNorNorCC Or psCodOpe = gCredPagNorMorCC Or psCodOpe = gCredPagNorVenCC Or psCodOpe = gCredPagRefNorCC Or psCodOpe = gCredPagRefMorCC Or psCodOpe = gCredPagRefVenCC Then
        If psCodOpe = gCredPagNorNorCC Or psCodOpe = gCredPagNorMorCC Or psCodOpe = gCredPagNorVenCC Or psCodOpe = gCredPagRefNorCC Or psCodOpe = gCredPagRefMorCC Or psCodOpe = gCredPagRefVenCC Then
            Call oBase.dExtornaSaldosCalendario(pnMov, gColocCalendAplCuota, psCtaCod, False)
        Else
            'Extorna Saldos Calendario
            Call oBase.dExtornaSaldosCalendario(pnMov, gColocCalendAplDesembolso, psCtaCod, False)
        End If
        If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Then
            Set RCap = oBase.RecuperaMovimientoCapataciones(pnMov)
            'Extorna Retiros por Gastos
            Do While Not RCapMov.EOF
                'If RCapMov!cOpeCod = "200301" Then '********* Aca hice la modificacion CAFF
                If RCapMov!copecod = gCredDesembCtaRetiroGastos Or RCapMov!copecod = gCredDesembCtaRetiroGastosDOA Or RCapMov!copecod = gCredDesembCtaRetiroCancelacion Or RCapMov!copecod = gCredDesembCtaRetiroCancelacionDOA Then
                    'Call oBase.dInsertMov(MatsMovNro(RCapMov.boomark - 1), gCredExtPago, "Extorno de Pago", gMovEstContabNoContable, gMovFlagDeExtorno, False)
                    Call oBase.CapExtornoCargoAho(nMovNro, RCapMov!nMovNro, gAhoExtRetEfec, RCap!Cctacod, MatsMovNro(RCapMov.Bookmark - 1), "Extorno de Desembolso Abono Cuenta", RCapMov!nMonto)
                End If
                RCapMov.MoveNext
            Loop
            'Recupera DatosExtorno
            dCap = CDate(Mid(RCap!cMovnro, 7, 2) & "/" & Mid(RCap!cMovnro, 5, 2) & "/" & Mid(RCap!cMovnro, 1, 4) & " " & Mid(RCap!cMovnro, 9, 2) & ":" & Mid(RCap!cMovnro, 11, 2) & ":" & Mid(RCap!cMovnro, 13, 2))
            'Extorno de Apertura
            'If RCap!cOpeCod = "200101" Then
                        
            If RCap!copecod = "100102" Then
                Call oBase.CapExtornoApertura(sMovNroCap, nMovNro, pnMovNro, gAhoExtApeEfec, RCap!Cctacod, sMovNro, "Extorno de Desembolso Abono Cuenta", RCap!nMonto)
            Else
                Call oBase.CapExtornoAbonoAho(sMovNroCap, nMovNro, pnMovNro, gAhoExtDepEfec, RCap!Cctacod, sMovNro, "Extorno de Desembolso Abono Cuenta", RCap!nMonto)
            End If
            
        ElseIf psCodOpe = gCredPagNorNorCC Or psCodOpe = gCredPagNorMorCC Or psCodOpe = gCredPagNorVenCC Or psCodOpe = gCredPagRefNorCC Or psCodOpe = gCredPagRefMorCC Or psCodOpe = gCredPagRefVenCC Then
            Do While Not RCapMov.EOF
                If RCapMov!copecod = gCredDesembCtaRetiroGastos Or RCapMov!copecod = gCredDesembCtaRetiroGastosDOA Or RCapMov!copecod = gCredDesembCtaRetiroCancelacion Or RCapMov!copecod = gCredDesembCtaRetiroCancelacionDOA Then
                    'Call oBase.CapExtornoCargoAho(nMovNro, RCapMov!nMovNro, gCredDesembCtaRetiroCancelacion, RCapMov!cCtaCod, MatsMovNro(RCapMov.Bookmark - 1), "Extorno de Desembolso Abono Cuenta", RCapMov!nMonto)
                    Call oBase.CapExtornoCargoAho(nMovNro, RCapMov!nMovNro, gCredDesembCtaRetiroCancelacion, RCapMov!Cctacod, sMovNro, "Extorno de Desembolso Abono Cuenta", RCapMov!nMonto)
                End If
                RCapMov.MoveNext
            Loop
            
        End If
    Else
        'Extorna Saldos Calendario
        Call oBase.dExtornaSaldosCalendario(pnMov, gColocCalendAplCuota, psCtaCod, False)
    End If
    
    'Extorna Saldos de Maestros
    'Colocaccred
    Call oBase.dUpdateColocacCred(psCtaCod, R!nDiasMora, , , , , , , R!nMinCuota, , , , , , , R!nMinCuota, , , False)
    
    'Producto
    If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Or psCodOpe = gCredDesembEfec Then
        Call oBase.dUpdateProducto(psCtaCod, , R!nCapital, R!nCredEstado, , , False, 1)
    Else
        Call oBase.dUpdateProducto(psCtaCod, , R!nCapital, R!nCredEstado, , , False, 1, True)
    End If
        
    '**************************************************************
    'En Caso que el pago halla sido con una Dacion en Pago
    '**************************************************************
    If psCodOpe = gCredPagNorNorDacion Or psCodOpe = gCredPagNorMorDacion Or psCodOpe = gCredPagNorVenDacion _
        Or psCodOpe = gCredPagRefNorDacion Or psCodOpe = gCredPagNorNorDacion Or psCodOpe = gCredPagRefMorDacion Or psCodOpe = gCredPagRefVenDacion Then
        Call oBase.dAnularColocGarantRec(R!nFlag, gColocGarantRecEstadoRegistrado, False)
    End If
    
    '**************************************************************
    'Si es Desembolso Con Cancelacion de Creditos
    '**************************************************************
    If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Or psCodOpe = gCredDesembEfec Then
        Set oDCred = New DCredito
        Set RCancel = oDCred.CreditosCanceladoConDesembolso(psCtaCod, pnMov)
        Do While Not RCancel.EOF
            Set R = oDCred.RecuperaDatosExtorno(pnMov, RCancel!Cctacod)
            Call oBase.dInsertMovCol(nMovNro, gCredExtPago, RCancel!Cctacod, 0, RCancel!nMonto, 0, "", 0, 0, 0, False)
            Call oBase.dExtornaSaldosCalendario(pnMov, gColocCalendAplCuota, RCancel!Cctacod, False)
            Call oBase.dUpdateColocacCred(RCancel!Cctacod, R!nDiasMora, , , , , , , R!nMinCuota, , , , , , , R!nMinCuota, , , False)
            Call oBase.dUpdateProducto(RCancel!Cctacod, , R!nCapital, R!nCredEstado, , , False, 0)
            RCancel.MoveNext
        Loop
    End If
    '***************************************************************
    'Si es Desembolso es un Hipotecario ---- CAFF
    '***************************************************************
    If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Or psCodOpe = gCredDesembEfec Then
        If Mid(psCtaCod, 6, 3) = "402" Then
            Call oBase.dDeleteColocCalifMiViv(psCtaCod)
        End If
    End If
    
    
            
    If Not pbExtornoEnLote Then
        Call oBase.dCommitTrans
    End If
    
    
    R.Close
    Set R = Nothing
                    
    
    '***************************************************************
    'Si es Prepago
    '***************************************************************
    Dim dCred As Dcalendario
    
    If Not (psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Or psCodOpe = gCredDesembEfec) Then
        If pnPrepago = 1 Then
    
            Call oBase.dBeginTrans
                        
            sSQL = "Select MIN(nNroCuota) as nCuotaMin from MovColDet where nMovNro = " & pnMov
            Set R = oBase.coConex.CargaRecordSet(sSQL)
            If R.RecordCount > 0 Then
                nCuotaMin = R!nCuotaMin
            End If
            R.Close
            
            Set R = oBase.RecuperaColocacCred(psCtaCod)
            nNroCalen = R!nNroCalen
            R.Close
            
            sSQL = "DELETE ColocCalifMiViv Where nCuota >= " & nCuotaMin & " AND nNroCalen = " & nNroCalen & " AND cCtaCod = '" & psCtaCod & "'"
            oBase.coConex.Ejecutar (sSQL)
            
            Set R = oBase.RecuperaColocacCred(psCtaCod)
            nNroCalen = R!nNroCalen - 2
            
            sSQL = "Select ISNULL(MAX(nCuotaOrig),0) as nCuotaMax from ColocCalifMiViv where cCtaCod = '" & psCtaCod & "'"
            Set R = oBase.coConex.CargaRecordSet(sSQL)
            If R.RecordCount > 0 Then
                i = R!nCuotaMax + 1
            End If
            R.Close
            
            Set dCred = New Dcalendario
            Set R = dCred.RecuperaColocCalendario(psCtaCod, nNroCalen, gColocCalendAplCuota)
            Set dCred = Nothing
            
            Do While Not R.EOF
                If R!nCuota >= nCuotaMin Then
                    Call oBase.dInsertColocCalifMiViv(psCtaCod, i, nNroCalen, R!nCuota, R!DVENC, R!nColocCalendEstado)
                End If
                i = i + 1
                R.MoveNext
            Loop
            
            Call oBase.dUpdateColocacCred(psCtaCod, , , , , , , , , , , , , , nNroCalen, , , , False, , , nNroCalen + 1)
            
            sSQL = "DELETE ColocCalendDet Where nNroCalen = " & nNroCalen + 2 & " AND cCtaCod =  '" & psCtaCod & "'"
            Call oBase.coConex.Ejecutar(sSQL)
            sSQL = "DELETE ColocCalendDet Where nNroCalen = " & nNroCalen + 3 & " AND cCtaCod =  '" & psCtaCod & "'"
            Call oBase.coConex.Ejecutar(sSQL)
            
            sSQL = "DELETE ColocCalendario Where nNroCalen = " & nNroCalen + 2 & " AND cCtaCod =  '" & psCtaCod & "'"
            Call oBase.coConex.Ejecutar(sSQL)
            sSQL = "DELETE ColocCalendario Where nNroCalen = " & nNroCalen + 3 & " AND cCtaCod =  '" & psCtaCod & "'"
            Call oBase.coConex.Ejecutar(sSQL)
            
                nMovNroCond = 0
                sSQL = "Select * From MovRef Where nMovNroRef = " & pnMov
                Set R = oBase.coConex.CargaRecordSet(sSQL)
                If R.RecordCount > 0 Then
                    nMovNroCond = R!nMovNro
                End If
                R.Close
                
                If nMovNroCond <> 0 Then
                    sSQL = "Select * From MovCol Where nMovNro = " & nMovNroCond
                    Set R = oBase.coConex.CargaRecordSet(sSQL)
                    If R.RecordCount > 0 Then
                        nMontoCond = R!nMonto
                    End If
                    R.Close
                End If
            
                If nMontoCond <> 0 Then
                    sSQL = "Update Producto Set nSaldo = nSaldo + " & Format(nMontoCond, "#0.00") & " Where cCtaCod = '" & psCtaCod & "'"
                    Call oBase.coConex.Ejecutar(sSQL)
                    
                    sSQL = "Update Mov Set nMovFlag = 1 Where nMovNro = " & nMovNroCond
                    Call oBase.coConex.Ejecutar(sSQL)
                    
                End If
                
            Call oBase.dCommitTrans
        Else 'Si no es un prepago verificar si ha sido una cuota con evaluacion de Mivivienda
            If Mid(psCtaCod, 6, 3) = "402" Then
                sSQL = "Select MIN(nNroCuota) as nCuotaMin from MovColDet where nMovNro = " & pnMov
                Set R = oBase.coConex.CargaRecordSet(sSQL)
                If R.RecordCount > 0 Then
                    nCuotaMin = R!nCuotaMin
                End If
                R.Close
                                                
                nMovNroCond = 0
                sSQL = "Select * From MovRef Where nMovNroRef = " & pnMov
                Set R = oBase.coConex.CargaRecordSet(sSQL)
                If R.RecordCount > 0 Then
                    nMovNroCond = R!nMovNro
                End If
                R.Close
                
                If nMovNroCond <> 0 Then
                    sSQL = "Select * From MovCol Where nMovNro = " & nMovNroCond
                    Set R = oBase.coConex.CargaRecordSet(sSQL)
                    If R.RecordCount > 0 Then
                        nMontoCond = R!nMonto
                    End If
                    R.Close
                End If
                
                Set R = oBase.RecuperaColocacCred(psCtaCod)
                nNroCalen = R!nNroCalen
                nNroCalenPar = R!nNroCalPar
                R.Close
                
                If nCuotaMin = 5 Or (((nCuotaMin - 5) Mod 6) = 0) Then
                    
                    sSQL = " UPDATE ColocCalifMiViv Set cColocMiVivEval = NULL Where nCuota >= " & Trim(Str(nCuotaMin - 5)) & " AND nCuota <= " & nCuotaMin
                    sSQL = sSQL & " AND cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & nNroCalen
                    Call oBase.coConex.Ejecutar(sSQL)
                    
                    sSQL = " UPDATE ColocCalifMiViv Set nColocCalendEstado = 0 Where nCuota = " & nCuotaMin
                    sSQL = sSQL & " AND cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & nNroCalen
                    Call oBase.coConex.Ejecutar(sSQL)
                
                    sSQL = " UPDATE ColocCalendario Set nColocCalendEstado = 0  Where nCuota >= " & nCuotaMin + 2 & " AND nCuota <= " & nCuotaMin + 7
                    sSQL = sSQL & " AND cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & nNroCalenPar & " AND nColocCalendApl = 1 "
                    Call oBase.coConex.Ejecutar(sSQL)
                    
                    If nMontoCond <> 0 Then
                        sSQL = "Update Producto Set nSaldo = nSaldo + " & Format(nMontoCond, "#0.00") & " Where cCtaCod = '" & psCtaCod & "'"
                        Call oBase.coConex.Ejecutar(sSQL)
                        
                        sSQL = "Update Mov Set nMovFlag = 1 Where nMovNro = " & nMovNroCond
                        Call oBase.coConex.Ejecutar(sSQL)
                        
                    End If
                
                End If
            End If
        End If
    End If
    
    Set oBase = Nothing
    
End Sub

Public Function GetCredRFA(ByVal psPersCod As String)
Dim sql As String
Dim oCon As New DConecta
Set oCon = New DConecta

sql = "SELECT P.CCTACOD " _
    & " FROM    PRODUCTO P " _
    & "         JOIN COLOCACCRED C ON C.CCTACOD = P.CCTACOD " _
    & "         JOIN PRODUCTOPERSONA R ON R.CCTACOD = P.CCTACOD AND nPrdPersRelac=20 " _
    & " WHERE   C.CRFA IN ('RFA','RFC','DIF') AND P.nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2202) " _
    & " AND R.cPersCod='" & psPersCod & "'"


oCon.AbreConexion
Set GetCredRFA = oCon.CargaRecordSet(sql)
oCon.CierraConexion

End Function
Public Function GetCalendariosRFA(ByVal psPersCod As String, ByVal psRFA As String) As ADODB.Recordset
Dim sql As String
Dim oCon As DConecta
Set oCon = New DConecta

sql = "SELECT   C.CCTACOD as cCodCta,  C.nColocCalendEstado AS cEstado, P.nPrdEstado AS nEstCred, " _
    & "         P.NSALDO AS NSALDOK, C1.dVigencia as dFecVig, C1.nMontoCol as nCapDes, " _
    & "         nTasInt = (select nRFATasInt FROM colocaccredrfa AS cRFA WHERE cRFA.cPerscod = R.cPersCod ), " _
    & "         nDiaApr  =  (SELECT CE.nPlazo FROM COLOCACESTADO CE WHERE CE.CCTACOD = CC.CCTACOD AND nPrdEstado=2002), " _
    & "         nSalInt = (   SELECT SUM(CD1.NMONTO - CD1.NMONTOPAGADO) " _
    & "                         FROM COLOCCALENDDET CD1 " _
    & "                         WHERE CD1.CCTACOD = CC.CCTACOD AND CD1.nColocCalendApl=1 AND CD1.nNroCalen =CC.nNroCalen and CD1.nPrdConceptoCod=1100), " _
    & "         nCuotas  = (    SELECT count(*) FROM COLOCCALENDARIO CC1 WHERE CC1.CCTACOD=CC.CCTACOD AND CC1.nColocCalendApl=1  AND CC1.nNroCalen=CC.nNroCalen), " _
    & "         C.DVENC as dFecVen, C.DPAGO,C.NCUOTA as cNrocuo, " _
    & "         nCapita= SUM(CASE WHEN CD.nPrdConceptoCod= 1000 THEN CD.nMonto ELSE 0 END), " _
    & "         nIntere= SUM(CASE WHEN CD.nPrdConceptoCod= 1100 THEN CD.nMonto ELSE 0 END), " _
    & "         nCofide= SUM(CASE WHEN CD.nPrdConceptoCod= 124352 THEN CD.nMonto ELSE 0 END), " _
    & "         nCuota= SUM(CASE WHEN CD.nPrdConceptoCod IN (1000,1100) THEN CD.nMonto ELSE 0 END), " _
    & "         nIntMor= SUM(CASE WHEN CD.nPrdConceptoCod = 1101 THEN CD.nMonto ELSE 0 END), " _
    & "         nCapPag= SUM(CASE WHEN CD.nPrdConceptoCod= 1000 THEN CD.nMontoPagado ELSE 0 END), " _
    & "         nIntPag= SUM(CASE WHEN CD.nPrdConceptoCod= 1100 THEN CD.nMontoPagado ELSE 0 END), " _
    & "         nMorPag= SUM(CASE WHEN CD.nPrdConceptoCod = 1101 THEN CD.nMontoPagado ELSE 0 END) " _
    & " FROM    COLOCCALENDARIO C " _
    & "         JOIN COLOCCALENDDET CD ON CD.CCTACOD = C.CCTACOD AND CD.nNroCalen = C.nNroCalen AND CD.nColocCalendApl= C.nColocCalendApl AND CD.nCuota = C.nCuota " _
    & "         JOIN PRODUCTO P ON P.CCTACOD = C.CCTACOD " _
    & "         JOIN COLOCACCRED CC ON CC.CCTACOD = P.CCTACOD " _
    & "         JOIN COLOCACIONES C1 ON C1.CCTACOD = P.CCTACOD " _
    & "         JOIN PRODUCTOPERSONA R ON R.CCTACOD = P.CCTACOD AND R.nPrdPersRelac=20 " _
    & " WHERE   R.cPersCod='" & psPersCod & " ' AND C.nColocCalendApl <>0 AND CD.nNroCalen = CC.nNroCalen AND CC.CRFA='" & psRFA & "' " _
    & " GROUP BY C.CCTACOD, C.NCUOTA, C.DVENC, C.DPAGO, C.nColocCalendEstado, P.nPrdEstado, "
sql = sql + "       P.nSaldo , C1.dVigencia, P.nTasaInteres, C1.nMontoCol, CC.Cctacod, CC.nNroCalen, C.DVENC " _
    & " ORDER BY C.NCUOTA"

oCon.AbreConexion
Set GetCalendariosRFA = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetCRFA(ByVal psCtaCod As String) As String
    Dim sSQL As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    
    Set oConec = New DConecta
    oConec.AbreConexion
    sSQL = "Select cRFa From ColocacCred Where cCtaCod='" & psCtaCod & "'"
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    If Not rs.EOF And Not rs.BOF Then
        GetCRFA = rs!cRFA
    End If
    Set rs = Nothing
End Function

Public Function RecuperaDatosExtorno(ByVal pnMovNro As Long, ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    sSQL = "Select M.nCredEstado, SUM(MD.nMonto) as nCapital, "
    sSQL = sSQL & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
    sSQL = sSQL & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro AND MD.nPrdConceptoCod in(" & gColocRFACapital & "," & gColocConceptoCodCapital & " ) Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psCtaCod & "' AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psCtaCod & "' and M.nCredEstado<>0 "
    sSQL = sSQL & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSQL)
    If RecuperaDatosExtorno.RecordCount = 0 Then
        sSQL = "Select M.nCredEstado, 0 as nCapital, "
        sSQL = sSQL & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
        sSQL = sSQL & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro AND MD.nPrdConceptoCod = " & gColocRFAInteresCompesatorio & "," & gColocConceptoCodInteresCompensatorio & ") Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psCtaCod & "' AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psCtaCod & "' "
        sSQL = sSQL & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
        Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSQL)
    End If
    Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSQL)
    If RecuperaDatosExtorno.RecordCount = 0 Then
        sSQL = "Select M.nCredEstado, 0 as nCapital, "
        sSQL = sSQL & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
        sSQL = sSQL & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro AND Left(Cast(nPrdConceptoCod as Varchar(4)),2)='12' Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psCtaCod & "' AND Left(Cast(nPrdConceptoCod as Varchar(4)),2)='12' AND M.cCtaCod = '" & psCtaCod & "' "
        sSQL = sSQL & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
        Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

