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
Dim vsCadImpresion As String
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

Public Function VerificaCreditoRFA(ByVal psctacod As String) As Boolean
    Dim objC As DConecta
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim intCantidad As Integer
    On Error GoTo ErrHandler
        strSQL = "Select Count(*) as Cantidad From ColocacCred Where cCtaCod='" & psctacod & "' "
        strSQL = strSQL & " AND cRFA in('RFC','RFA','DIF')"

        Set objC = New DConecta
        objC.AbreConexion
        Set rs = objC.CargaRecordSet(strSQL)
        objC.CierraConexion
        Set objC = Nothing

        If Not rs.EOF And Not rs.BOF Then
            intCantidad = rs!cantidad
        End If
        Set rs = Nothing
        If intCantidad = 0 Then
            VerificaCreditoRFA = False
        Else
            VerificaCreditoRFA = True
        End If

    Exit Function
ErrHandler:
    If Not objC Is Nothing Then Set objC = Nothing
    VerificaCreditoRFA = False
End Function

Public Function BuscaDNICliente(ByVal psctacod As String) As String
    Dim objC As DConecta
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrHandler
    strSQL = "Select cPersIDnro from persid where cperscod='" & psctacod & "'"
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

Function ListaCreditosPendientes(ByVal pcCodCli As String, ByVal pdFecSis As String, Optional ByVal pbRefinac As Boolean = False) As ADODB.Recordset
      Dim objC As DConecta
      Dim sSql As String
      On Error GoTo ErrHandler
      
      If pbRefinac = False Then
        sSql = "set dateformat dmy"
        sSql = sSql & " Select CC.cRFA as cCredito,CO.nCuota as nCuota,PP.cCtaCod,"
        sSql = sSql & " Capital=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in (" & gColocConceptoCodCapital & "," & gColocRFACapital & ") and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen),"
        sSql = sSql & " IntComp=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in (" & gColocConceptoCodInteresCompensatorio & "," & gColocRFAInteresCompesatorio & ") and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen),"
        sSql = sSql & " IntMor=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in (" & gColocConceptoCodInteresMoratorio & "," & gColocRFAInteresMoratorio & ") and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen),"
        sSql = sSql & " Gastos=(Select isnull(Sum(nMonto-nMontoPagado),0) From coloccalenddet a  where a.cctaCod=PP.cCtaCod and a.nColocCalendApl=1 and a.nCuota=CO.nCuota and a.nPrdConceptoCod Like '12%' and a.nPrdConceptoCod not in('124352','124353','124354','124355') and a.nNroCalen=CO.nNroCalen),"
        sSql = sSql & " ComCofide=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nCuota=CO.nCuota and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod=124352 and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen), CO.dVenc, diasAtraso=DateDiff(day,dVenc,'" & pdFecSis & "'),"
        sSql = sSql & " Indice=Case cc.cRFA    When 'DIF' Then 1    When 'RFC' Then 2    When 'RFA' Then 3 End "
        sSql = sSql & " From ProductoPersona PP"
        sSql = sSql & " Inner Join ColocCalendario CO on PP.cCtaCod=CO.cCtaCod"
        sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=CO.cCtaCod and CO.nNroCalen=CC.nNroCalen"
        sSql = sSql & " Where PP.cPersCod='" & pcCodCli & "' and  CC.cRFA IN('RFC','DIF','RFA') and"
        sSql = sSql & " CO.nColocCalendApl=1 and"
        sSql = sSql & " CO.nCuota=(Select Min(nCuota) From ColocCalendario c1"
        sSql = sSql & " inner join ColocacCred d on c1.cCtaCod=d.cCtaCod and C1.nNroCalen=d.nNroCalen"
        sSql = sSql & " inner join ProductoPersona Ps on Ps.cCtaCod=C1.cCtaCod and Ps.nPrdPersRelac=20"
        sSql = sSql & " Where Ps.cPersCod='" & pcCodCli & "' and c1.nColocCalendApl=CO.nColocCalendApl and c1.nColocCalendEstado=0"
        sSql = sSql & " and d.cRFA IN('RFC','DIF','RFA'))"
        sSql = sSql & " Order by nCuota,Indice"
     Else
        sSql = "set dateformat dmy"
        sSql = sSql & " Select CC.cRFA as cCredito,CO.nCuota as nCuota,PP.cCtaCod,"
        sSql = sSql & " Capital=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in (" & gColocConceptoCodCapital & "," & gColocRFACapital & ") and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen),"
        sSql = sSql & " IntComp=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in (" & gColocConceptoCodInteresCompensatorio & "," & gColocRFAInteresCompesatorio & ") and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen),"
        sSql = sSql & " IntMor=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in (" & gColocConceptoCodInteresMoratorio & "," & gColocRFAInteresMoratorio & ") and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen),"
        sSql = sSql & " Gastos=(Select isnull(Sum(nMonto-nMontoPagado),0) From coloccalenddet a  where a.cctaCod=PP.cCtaCod and a.nColocCalendApl=1 and a.nCuota=CO.nCuota and a.nPrdConceptoCod Like '12%' and a.nPrdConceptoCod not in('124352','124353','124354','124355') and a.nNroCalen=CO.nNroCalen),"
        sSql = sSql & " ComCofide=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nCuota=CO.nCuota and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod=124352 and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen), CO.dVenc, diasAtraso=DateDiff(day,dVenc,'" & pdFecSis & "'),"
        sSql = sSql & " Indice=Case cc.cRFA    When 'DIF' Then 1    When 'RFC' Then 2    When 'RFA' Then 3 End "
        sSql = sSql & " From ProductoPersona PP"
        sSql = sSql & " Inner Join ColocCalendario CO on PP.cCtaCod=CO.cCtaCod"
        sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=CO.cCtaCod and CO.nNroCalen=CC.nNroCalen"
        sSql = sSql & " Where PP.cPersCod='" & pcCodCli & "' and  CC.cRFA IN('RFC','DIF','RFA') and"
        sSql = sSql & " CO.nColocCalendApl=1 and"
        sSql = sSql & " CO.nCuota=(Select Min(nCuota) From ColocCalendario c1"
        sSql = sSql & " inner join ColocacCred d on c1.cCtaCod=d.cCtaCod and C1.nNroCalen=d.nNroCalen"
        sSql = sSql & " inner join ProductoPersona Ps on Ps.cCtaCod=C1.cCtaCod and Ps.nPrdPersRelac=20"
        sSql = sSql & " Where Ps.cPersCod='" & pcCodCli & "' and c1.nColocCalendApl=CO.nColocCalendApl and c1.nColocCalendEstado=0"
        sSql = sSql & " and d.cRFA IN('RFC','DIF','RFA'))"
        sSql = sSql & " and CC.bRefinacRFA=1"
        sSql = sSql & " Order by nCuota,Indice"
     End If
    
    Set objC = New DConecta
    objC.AbreConexion
    Set ListaCreditosPendientes = objC.CargaRecordSet(sSql)
    objC.CierraConexion
    Set objC = Nothing
      
    Exit Function
ErrHandler:
      If Not objC Is Nothing Then Set objC = Nothing
      Set ListaCreditosPendientes = Null
End Function

Function ListaCreditosPendientesCuota(ByVal pcCodCli As String, ByVal pdFecSis As String, ByVal pnCuota As Integer, _
Optional ByVal pBRefina As Boolean = False) As ADODB.Recordset
      Dim objC As DConecta
      Dim sSql As String
      On Error GoTo ErrHandler
                       
      If pBRefina = False Then
        sSql = "set dateformat dmy"
        sSql = sSql & " Select CC.cRFA as cCredito,CO.nCuota as nCuota,PP.cCtaCod,"
        sSql = sSql & " Capital=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in(1000,1010) and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen),"
        sSql = sSql & " IntComp=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in(1100,1107) and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen),"
        sSql = sSql & " IntMor=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in(1101,1108) and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen) ,"
        sSql = sSql & " Gastos=(Select isnull(Sum(nMonto-nMontoPagado),0) From coloccalenddet a  where a.cctaCod=PP.cCtaCod and a.nColocCalendApl=1 and a.nCuota=CO.nCuota and a.nPrdConceptoCod Like '12%' and a.nPrdConceptoCod not in('124352','124353','124354','124355') and a.nNroCalen=CO.nNroCalen),"
        sSql = sSql & " ComCofide=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nCuota=CO.nCuota and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod=124352 and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen), CO.dVenc, diasAtraso=DateDiff(day,dVenc,'" & pdFecSis & "'),"
        sSql = sSql & " Indice=Case cc.cRFA    When 'DIF' Then 1    When 'RFC' Then 2    When 'RFA' Then 3 End"
        sSql = sSql & " From ProductoPersona PP"
        sSql = sSql & " Inner Join ColocCalendario CO on PP.cCtaCod=CO.cCtaCod"
        sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=CO.cCtaCod and CO.nNroCalen=CC.nNroCalen"
        sSql = sSql & " Where PP.cPersCod='" & pcCodCli & "' and  CC.cRFA IN('RFC','DIF','RFA') and"
        sSql = sSql & " CO.nColocCalendApl=1 and"
        sSql = sSql & " CO.nCuota=" & pnCuota
        sSql = sSql & " Order by nCuota,Indice"
      Else
        sSql = "set dateformat dmy"
        sSql = sSql & " Select CC.cRFA as cCredito,CO.nCuota as nCuota,PP.cCtaCod,"
        sSql = sSql & " Capital=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in(1000,1010) and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen),"
        sSql = sSql & " IntComp=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in(1100,1107) and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen),"
        sSql = sSql & " IntMor=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod in(1101,1108) and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen),"
        sSql = sSql & " Gastos=(Select isnull(Sum(nMonto-nMontoPagado),0) From coloccalenddet a  where a.cctaCod=PP.cCtaCod and a.nColocCalendApl=1 and a.nCuota=CO.nCuota and a.nPrdConceptoCod Like '12%' and a.nPrdConceptoCod not in('124352','124353','124354','124355') and a.nNroCalen=CO.nNroCalen),"
        sSql = sSql & " ComCofide=(Select isnull(Sum(nMonto-nMontoPagado),0) From ColocCalendDet CD Where CD.cCtaCod=CO.cCtaCod and CD.nCuota=CO.nCuota and CD.nColocCalendApl=CO.nColocCalendApl and CD.nPrdConceptoCod=124352 and CD.nCuota=CO.nCuota and CD.nNroCalen=CO.nNroCalen), CO.dVenc, diasAtraso=DateDiff(day,dVenc,'" & pdFecSis & "'),"
        sSql = sSql & " Indice=Case cc.cRFA    When 'DIF' Then 1    When 'RFC' Then 2    When 'RFA' Then 3 End"
        sSql = sSql & " From ProductoPersona PP"
        sSql = sSql & " Inner Join ColocCalendario CO on PP.cCtaCod=CO.cCtaCod"
        sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=CO.cCtaCod and CO.nNroCalen=CC.nNroCalen"
        sSql = sSql & " Where PP.cPersCod='" & pcCodCli & "' and  CC.cRFA IN('RFC','DIF','RFA') and"
        sSql = sSql & " CO.nColocCalendApl=1 and"
        sSql = sSql & " CO.nCuota=" & pnCuota
        sSql = sSql & " and CC.bRefinacRFA=1"
        sSql = sSql & " Order by nCuota,Indice"
      End If
    
    Set objC = New DConecta
    objC.AbreConexion
    Set ListaCreditosPendientesCuota = objC.CargaRecordSet(sSql)
    objC.CierraConexion
    Set objC = Nothing
      
    Exit Function
ErrHandler:
      If Not objC Is Nothing Then Set objC = Nothing
      Set ListaCreditosPendientesCuota = Null
End Function


Function ObtenerFactor(ByVal psctacod As String) As Double
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
        sSql = "Select nRfaFactor From FactorRfa "
        sSql = sSql & " where cctacod='" & psctacod & "'"
        
        Set oConec = New DConecta
        oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sSql)
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
    Dim sSql As String
    
    sSql = "Select cPersCod,cPersNombre,cPersDireccDomicilio"
    sSql = sSql & " From Persona"
    sSql = sSql & " Where cPersCod='" & pcPersCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set BuscarPersona = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerExtornoRFA(ByVal psPersCod As String, ByVal psUsuario As String, _
       ByVal pdFecha As Date, ByVal psCodAgencia As String) As Recordset
    
    Dim oConec As DConecta
    Dim sSql As String
    Dim sFecha As String
    
    On Error GoTo ErrHandler
    
    'Convirtiendo la fecha
    sFecha = Mid(CStr(pdFecha), 7, 4) & Mid(CStr(pdFecha), 4, 2) & Mid(CStr(pdFecha), 1, 2)
    sSql = "Select PP.cCtaCod,"
    sSql = sSql & " MC.cOpeCod,"
    sSql = sSql & " Substring(M.cMovNro,8,2)+':'+Substring(M.cMovNro,11,2)+ ':' +Substring(M.cMovNro,13,2) as cHora,"
    sSql = sSql & " MAX(MC.nMovNro)as nMovNro,"
    sSql = sSql & " Sum(MC.nMonto) as nMonto,"
    sSql = sSql & " right(M.cMovNro,4) as cUsuario,"
    sSql = sSql & " MC.nPrePago,"
    sSql = sSql & " CC.cRFA"
    sSql = sSql & " From Mov M"
    sSql = sSql & " Inner Join  MovCol MC on MC.nMovNro=M.nMovNro"
    sSql = sSql & " Inner Join ColocacCred CC on MC.cCtaCod=CC.cCtacod"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=MC.cCtaCod"
    sSql = sSql & " Where Left(M.cMovNro,8)='" & sFecha & "' and right(M.cMovNro,4)='" & psUsuario & "' and M.nMovFlag<>2  and MC.cOpeCod not like '109%' and "
    sSql = sSql & " CC.cRFA in ('DIF','RFC','RFA') and PP.nPrdPersRelac=20 and PP.cPersCod='" & psPersCod & "'"
    sSql = sSql & " and Substring(CC.cCtaCod,4,2)='" & psCodAgencia & "'"
    sSql = sSql & " Group By PP.cCtaCod,MC.cOpeCod,Substring(M.cMovNro,8,2)+':'+Substring(M.cMovNro,11,2)+ ':' +Substring(M.cMovNro,13,2),"
    sSql = sSql & " right(M.cMovNro,4),MC.nPrePago,CC.cRFA"
    sSql = sSql & " Having Sum(MC.nMonto) > 0"
    sSql = sSql & " Order By MAX(MC.nMovNro),MC.nPrePago,CC.cRFA"
    
    Set oConec = New DConecta
    
    oConec.AbreConexion
    Set ObtenerExtornoRFA = oConec.CargaRecordSet(sSql)
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
 Dim sSql As String
 Dim oConec As DConecta
 Dim sFecha As String
    On Error GoTo ErrHandler
        
        'Convirtiendo la fecha
    sFecha = Mid(CStr(pdFecha), 7, 4) & Mid(CStr(pdFecha), 4, 2) & Mid(CStr(pdFecha), 1, 2)
    
    sSql = sSql & "Select PP.cCtaCod,"
    sSql = sSql & " Max(MC.nMovNro) As nMovNro,"
    sSql = sSql & " Sum(MC.nMonto) as nMonto"
    sSql = sSql & " From Mov M"
    sSql = sSql & " Inner Join  MovCol MC on MC.nMovNro=M.nMovNro"
    sSql = sSql & " Inner Join ColocacCred CC on MC.cCtaCod=CC.cCtacod"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=MC.cCtaCod"
    sSql = sSql & " Where Left(M.cMovNro,8)='20040930' and right(M.cMovNro,4)='VDGT' and M.nM vFlag<>2 and"
    sSql = sSql & " CC.cRFA in ('DIF','RFC','RFA') and PP.nPrdPersRelac=20 and PP.cPersCod='1080100748200' and"
    sSql = sSql & " Substring(CC.cCtaCod,4,2)='01'"
    sSql = sSql & " Group By PP.cCtaCod"
    sSql = sSql & " Order By Max(MC.nMovNro)"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ObtenerCuentasExtorno = oConec.CargaRecordSet(sSql)
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
        Call ExtornarCredito(R!cCtaCod, pdFecSis, psCodUser, psCodAge, pnMovNro, R!copecod, R!nMonto, , True, oBase, nMovNro)
        R.MoveNext
    Loop
    oBase.dCommitTrans
    Set oBase = Nothing
    R.Close
    Set R = Nothing
End Sub

Public Function RecuperaCreditosdePago(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    sSql = "Select cCtaCod, cOpeCod, nMonto From MovCol Where nMovNro = " & pnMovNro
    sSql = sSql & " and nMonto>0 and cOpeCod not like '99%'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosdePago = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Sub ExtornarCredito(ByVal psctacod As String, ByVal pdFecSis As Date, ByVal psCodUser As String, _
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
Dim sSql As String
Dim nCuotaMin As Integer
Dim nNroCalen As Integer
Dim nNroCalenPar As Integer
Dim i As Integer
Dim nMontoCond As Double
Dim nMovNroCond As Long

    'Set oDCred = New DCredito
    Set R = RecuperaDatosExtorno(pnMov, psctacod)
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
        Call oBase.dInsertMovCol(nMovNro, gCredExtPago, psctacod, 0, pnMonto, 0, "", 0, 0, 0, False)
    Else
        Call oBase.dInsertMovCol(nMovNro, gCredExtPagoLote, psctacod, 0, pnMonto, 0, "", 0, 0, 0, False)
    End If
    
    
    If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Or psCodOpe = gCredDesembEfec Or psCodOpe = gCredPagNorNorCC Or psCodOpe = gCredPagNorMorCC Or psCodOpe = gCredPagNorVenCC Or psCodOpe = gCredPagRefNorCC Or psCodOpe = gCredPagRefMorCC Or psCodOpe = gCredPagRefVenCC Then
        If psCodOpe = gCredPagNorNorCC Or psCodOpe = gCredPagNorMorCC Or psCodOpe = gCredPagNorVenCC Or psCodOpe = gCredPagRefNorCC Or psCodOpe = gCredPagRefMorCC Or psCodOpe = gCredPagRefVenCC Then
            Call oBase.dExtornaSaldosCalendario(pnMov, gColocCalendAplCuota, psctacod, False)
        Else
            'Extorna Saldos Calendario
            Call oBase.dExtornaSaldosCalendario(pnMov, gColocCalendAplDesembolso, psctacod, False)
        End If
        If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Then
            Set RCap = oBase.RecuperaMovimientoCapataciones(pnMov)
            'Extorna Retiros por Gastos
            Do While Not RCapMov.EOF
                'If RCapMov!cOpeCod = "200301" Then '********* Aca hice la modificacion CAFF
                If RCapMov!copecod = gCredDesembCtaRetiroGastos Or RCapMov!copecod = gCredDesembCtaRetiroGastosDOA Or RCapMov!copecod = gCredDesembCtaRetiroCancelacion Or RCapMov!copecod = gCredDesembCtaRetiroCancelacionDOA Then
                    'Call oBase.dInsertMov(MatsMovNro(RCapMov.boomark - 1), gCredExtPago, "Extorno de Pago", gMovEstContabNoContable, gMovFlagDeExtorno, False)
                    Call oBase.CapExtornoCargoAho(nMovNro, RCapMov!nMovNro, gAhoExtRetEfec, RCap!cCtaCod, MatsMovNro(RCapMov.Bookmark - 1), "Extorno de Desembolso Abono Cuenta", RCapMov!nMonto)
                End If
                RCapMov.MoveNext
            Loop
            'Recupera DatosExtorno
            dCap = CDate(Mid(RCap!cMovNro, 7, 2) & "/" & Mid(RCap!cMovNro, 5, 2) & "/" & Mid(RCap!cMovNro, 1, 4) & " " & Mid(RCap!cMovNro, 9, 2) & ":" & Mid(RCap!cMovNro, 11, 2) & ":" & Mid(RCap!cMovNro, 13, 2))
            'Extorno de Apertura
            'If RCap!cOpeCod = "200101" Then
                        
            If RCap!copecod = "100102" Then
                Call oBase.CapExtornoApertura(sMovNroCap, nMovNro, pnMovNro, gAhoExtApeEfec, RCap!cCtaCod, sMovNro, "Extorno de Desembolso Abono Cuenta", RCap!nMonto)
            Else
                Call oBase.CapExtornoAbonoAho(sMovNroCap, nMovNro, pnMovNro, gAhoExtDepEfec, RCap!cCtaCod, sMovNro, "Extorno de Desembolso Abono Cuenta", RCap!nMonto)
            End If
            
        ElseIf psCodOpe = gCredPagNorNorCC Or psCodOpe = gCredPagNorMorCC Or psCodOpe = gCredPagNorVenCC Or psCodOpe = gCredPagRefNorCC Or psCodOpe = gCredPagRefMorCC Or psCodOpe = gCredPagRefVenCC Then
            Do While Not RCapMov.EOF
                If RCapMov!copecod = gCredDesembCtaRetiroGastos Or RCapMov!copecod = gCredDesembCtaRetiroGastosDOA Or RCapMov!copecod = gCredDesembCtaRetiroCancelacion Or RCapMov!copecod = gCredDesembCtaRetiroCancelacionDOA Then
                    'Call oBase.CapExtornoCargoAho(nMovNro, RCapMov!nMovNro, gCredDesembCtaRetiroCancelacion, RCapMov!cCtaCod, MatsMovNro(RCapMov.Bookmark - 1), "Extorno de Desembolso Abono Cuenta", RCapMov!nMonto)
                    Call oBase.CapExtornoCargoAho(nMovNro, RCapMov!nMovNro, gCredDesembCtaRetiroCancelacion, RCapMov!cCtaCod, sMovNro, "Extorno de Desembolso Abono Cuenta", RCapMov!nMonto)
                End If
                RCapMov.MoveNext
            Loop
            
        End If
    Else
        'Extorna Saldos Calendario
        Call oBase.dExtornaSaldosCalendario(pnMov, gColocCalendAplCuota, psctacod, False)
    End If
    
    'Extorna Saldos de Maestros
    'Colocaccred
    Call oBase.dUpdateColocacCred(psctacod, R!nDiasMora, , , , , , , R!nMinCuota, , , , , , , R!nMinCuota, , , False)
    
    'Producto
    If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Or psCodOpe = gCredDesembEfec Then
        Call oBase.dUpdateProducto(psctacod, , R!nCapital, R!nCredEstado, , , False, 1)
    Else
        Call oBase.dUpdateProducto(psctacod, , R!nCapital, R!nCredEstado, , , False, 1, True)
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
        Set RCancel = oDCred.CreditosCanceladoConDesembolso(psctacod, pnMov)
        Do While Not RCancel.EOF
            Set R = oDCred.RecuperaDatosExtorno(pnMov, RCancel!cCtaCod)
            Call oBase.dInsertMovCol(nMovNro, gCredExtPago, RCancel!cCtaCod, 0, RCancel!nMonto, 0, "", 0, 0, 0, False)
            Call oBase.dExtornaSaldosCalendario(pnMov, gColocCalendAplCuota, RCancel!cCtaCod, False)
            Call oBase.dUpdateColocacCred(RCancel!cCtaCod, R!nDiasMora, , , , , , , R!nMinCuota, , , , , , , R!nMinCuota, , , False)
            Call oBase.dUpdateProducto(RCancel!cCtaCod, , R!nCapital, R!nCredEstado, , , False, 0)
            RCancel.MoveNext
        Loop
    End If
    '***************************************************************
    'Si es Desembolso es un Hipotecario ---- CAFF
    '***************************************************************
    If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Or psCodOpe = gCredDesembEfec Then
        If Mid(psctacod, 6, 3) = "402" Then
            Call oBase.dDeleteColocCalifMiViv(psctacod)
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
                        
            sSql = "Select MIN(nNroCuota) as nCuotaMin from MovColDet where nMovNro = " & pnMov
            Set R = oBase.coConex.CargaRecordSet(sSql)
            If R.RecordCount > 0 Then
                nCuotaMin = R!nCuotaMin
            End If
            R.Close
            
            Set R = oBase.RecuperaColocacCred(psctacod)
            nNroCalen = R!nNroCalen
            R.Close
            
            sSql = "DELETE ColocCalifMiViv Where nCuota >= " & nCuotaMin & " AND nNroCalen = " & nNroCalen & " AND cCtaCod = '" & psctacod & "'"
            oBase.coConex.Ejecutar (sSql)
            
            Set R = oBase.RecuperaColocacCred(psctacod)
            nNroCalen = R!nNroCalen - 2
            
            sSql = "Select ISNULL(MAX(nCuotaOrig),0) as nCuotaMax from ColocCalifMiViv where cCtaCod = '" & psctacod & "'"
            Set R = oBase.coConex.CargaRecordSet(sSql)
            If R.RecordCount > 0 Then
                i = R!nCuotaMax + 1
            End If
            R.Close
            
            Set dCred = New Dcalendario
            Set R = dCred.RecuperaColocCalendario(psctacod, nNroCalen, gColocCalendAplCuota)
            Set dCred = Nothing
            
            Do While Not R.EOF
                If R!nCuota >= nCuotaMin Then
                    Call oBase.dInsertColocCalifMiViv(psctacod, i, nNroCalen, R!nCuota, R!dVenc, R!ncoloccalendestado)
                End If
                i = i + 1
                R.MoveNext
            Loop
            
            Call oBase.dUpdateColocacCred(psctacod, , , , , , , , , , , , , , nNroCalen, , , , False, , , nNroCalen + 1)
            
            sSql = "DELETE ColocCalendDet Where nNroCalen = " & nNroCalen + 2 & " AND cCtaCod =  '" & psctacod & "'"
            Call oBase.coConex.Ejecutar(sSql)
            sSql = "DELETE ColocCalendDet Where nNroCalen = " & nNroCalen + 3 & " AND cCtaCod =  '" & psctacod & "'"
            Call oBase.coConex.Ejecutar(sSql)
            
            sSql = "DELETE ColocCalendario Where nNroCalen = " & nNroCalen + 2 & " AND cCtaCod =  '" & psctacod & "'"
            Call oBase.coConex.Ejecutar(sSql)
            sSql = "DELETE ColocCalendario Where nNroCalen = " & nNroCalen + 3 & " AND cCtaCod =  '" & psctacod & "'"
            Call oBase.coConex.Ejecutar(sSql)
            
                nMovNroCond = 0
                sSql = "Select * From MovRef Where nMovNroRef = " & pnMov
                Set R = oBase.coConex.CargaRecordSet(sSql)
                If R.RecordCount > 0 Then
                    nMovNroCond = R!nMovNro
                End If
                R.Close
                
                If nMovNroCond <> 0 Then
                    sSql = "Select * From MovCol Where nMovNro = " & nMovNroCond
                    Set R = oBase.coConex.CargaRecordSet(sSql)
                    If R.RecordCount > 0 Then
                        nMontoCond = R!nMonto
                    End If
                    R.Close
                End If
            
                If nMontoCond <> 0 Then
                    sSql = "Update Producto Set nSaldo = nSaldo + " & Format(nMontoCond, "#0.00") & " Where cCtaCod = '" & psctacod & "'"
                    Call oBase.coConex.Ejecutar(sSql)
                    
                    sSql = "Update Mov Set nMovFlag = 1 Where nMovNro = " & nMovNroCond
                    Call oBase.coConex.Ejecutar(sSql)
                    
                End If
                
            Call oBase.dCommitTrans
        Else 'Si no es un prepago verificar si ha sido una cuota con evaluacion de Mivivienda
            If Mid(psctacod, 6, 3) = "402" Then
                sSql = "Select MIN(nNroCuota) as nCuotaMin from MovColDet where nMovNro = " & pnMov
                Set R = oBase.coConex.CargaRecordSet(sSql)
                If R.RecordCount > 0 Then
                    nCuotaMin = R!nCuotaMin
                End If
                R.Close
                                                
                nMovNroCond = 0
                sSql = "Select * From MovRef Where nMovNroRef = " & pnMov
                Set R = oBase.coConex.CargaRecordSet(sSql)
                If R.RecordCount > 0 Then
                    nMovNroCond = R!nMovNro
                End If
                R.Close
                
                If nMovNroCond <> 0 Then
                    sSql = "Select * From MovCol Where nMovNro = " & nMovNroCond
                    Set R = oBase.coConex.CargaRecordSet(sSql)
                    If R.RecordCount > 0 Then
                        nMontoCond = R!nMonto
                    End If
                    R.Close
                End If
                
                Set R = oBase.RecuperaColocacCred(psctacod)
                nNroCalen = R!nNroCalen
                nNroCalenPar = R!nNroCalPar
                R.Close
                
                If nCuotaMin = 5 Or (((nCuotaMin - 5) Mod 6) = 0) Then
                    
                    sSql = " UPDATE ColocCalifMiViv Set cColocMiVivEval = NULL Where nCuota >= " & Trim(Str(nCuotaMin - 5)) & " AND nCuota <= " & nCuotaMin
                    sSql = sSql & " AND cCtaCod = '" & psctacod & "' AND nNroCalen = " & nNroCalen
                    Call oBase.coConex.Ejecutar(sSql)
                    
                    sSql = " UPDATE ColocCalifMiViv Set nColocCalendEstado = 0 Where nCuota = " & nCuotaMin
                    sSql = sSql & " AND cCtaCod = '" & psctacod & "' AND nNroCalen = " & nNroCalen
                    Call oBase.coConex.Ejecutar(sSql)
                
                    sSql = " UPDATE ColocCalendario Set nColocCalendEstado = 0  Where nCuota >= " & nCuotaMin + 2 & " AND nCuota <= " & nCuotaMin + 7
                    sSql = sSql & " AND cCtaCod = '" & psctacod & "' AND nNroCalen = " & nNroCalenPar & " AND nColocCalendApl = 1 "
                    Call oBase.coConex.Ejecutar(sSql)
                    
                    If nMontoCond <> 0 Then
                        sSql = "Update Producto Set nSaldo = nSaldo + " & Format(nMontoCond, "#0.00") & " Where cCtaCod = '" & psctacod & "'"
                        Call oBase.coConex.Ejecutar(sSql)
                        
                        sSql = "Update Mov Set nMovFlag = 1 Where nMovNro = " & nMovNroCond
                        Call oBase.coConex.Ejecutar(sSql)
                        
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
Public Function GetCalendariosRFA(ByVal psPersCod As String, ByVal psRFA As String, _
Optional ByVal pbReprogramacion As Boolean = False, Optional ByVal pbRefinan As Boolean = False) As ADODB.Recordset
Dim sql As String
Dim oCon As DConecta
Set oCon = New DConecta

If pbReprogramacion = False Then
    sql = "SELECT   C.CCTACOD as cCodCta,  C.nColocCalendEstado AS cEstado, P.nPrdEstado AS nEstCred, " _
        & "         P.NSALDO AS NSALDOK, C1.dVigencia as dFecVig, C1.nMontoCol as nCapDes, " _
        & "         nTasInt = (select nRFATasInt FROM colocaccredrfa AS cRFA WHERE cRFA.cPerscod = R.cPersCod ), " _
        & "         nDiaApr  =  (SELECT CE.nPlazo FROM COLOCACESTADO CE WHERE CE.CCTACOD = CC.CCTACOD AND nPrdEstado=2002), " _
        & "         nSalInt = (   SELECT SUM(CD1.NMONTO - CD1.NMONTOPAGADO) " _
        & "                         FROM COLOCCALENDDET CD1 " _
        & "                         WHERE CD1.CCTACOD = CC.CCTACOD AND CD1.nColocCalendApl=1 AND CD1.nNroCalen =CC.nNroCalen and CD1.nPrdConceptoCod in(1100,1107)), " _
        & "         nCuotas  = (    SELECT count(*) FROM COLOCCALENDARIO CC1 WHERE CC1.CCTACOD=CC.CCTACOD AND CC1.nColocCalendApl=1  AND CC1.nNroCalen=CC.nNroCalen), " _
        & "         C.DVENC as dFecVen, C.DPAGO,C.NCUOTA as cNrocuo, " _
        & "         nCapita= SUM(CASE WHEN CD.nPrdConceptoCod= 1000 OR CD.nPrdConceptoCod= 1010 THEN CD.nMonto ELSE 0 END), " _
        & "         nIntere= SUM(CASE WHEN CD.nPrdConceptoCod= 1100 OR CD.nPrdConceptoCod= 1107 THEN CD.nMonto ELSE 0 END), " _
        & "         nCofide= SUM(CASE WHEN CD.nPrdConceptoCod= 124352 THEN CD.nMonto ELSE 0 END), " _
        & "         nCuota= SUM(CASE WHEN CD.nPrdConceptoCod IN (1000,1100,1010,1107) THEN CD.nMonto ELSE 0 END), " _
        & "         nIntMor= SUM(CASE WHEN CD.nPrdConceptoCod = 1101 OR CD.nPrdConceptoCod = 1108 THEN CD.nMonto ELSE 0 END), " _
        & "         nCapPag= SUM(CASE WHEN CD.nPrdConceptoCod= 1000 OR  CD.nPrdConceptoCod= 1010 THEN CD.nMontoPagado ELSE 0 END), " _
        & "         nIntPag= SUM(CASE WHEN CD.nPrdConceptoCod= 1100  OR CD.nPrdConceptoCod= 1107 THEN CD.nMontoPagado ELSE 0 END), " _
        & "         nMorPag= SUM(CASE WHEN CD.nPrdConceptoCod = 1101 OR CD.nPrdConceptoCod = 1108 THEN CD.nMontoPagado ELSE 0 END) " _
        & " FROM    COLOCCALENDARIO C " _
        & "         JOIN COLOCCALENDDET CD ON CD.CCTACOD = C.CCTACOD AND CD.nNroCalen = C.nNroCalen AND CD.nColocCalendApl= C.nColocCalendApl AND CD.nCuota = C.nCuota " _
        & "         JOIN PRODUCTO P ON P.CCTACOD = C.CCTACOD " _
        & "         JOIN COLOCACCRED CC ON CC.CCTACOD = P.CCTACOD " _
        & "         JOIN COLOCACIONES C1 ON C1.CCTACOD = P.CCTACOD " _
        & "         JOIN PRODUCTOPERSONA R ON R.CCTACOD = P.CCTACOD AND R.nPrdPersRelac=20 " _
        & " WHERE   R.cPersCod='" & psPersCod & " ' AND C.nColocCalendApl <>0 AND CD.nNroCalen = CC.nNroCalen AND CC.CRFA='" & psRFA & "' " _
        & " GROUP BY R.cPersCod,C.CCTACOD, C.NCUOTA, C.DVENC, C.DPAGO, C.nColocCalendEstado, P.nPrdEstado, "
    sql = sql + "       P.nSaldo , C1.dVigencia, P.nTasaInteres, C1.nMontoCol, CC.Cctacod, CC.nNroCalen, C.DVENC " _
        & " ORDER BY C.NCUOTA"
Else
    sql = "SELECT   C.CCTACOD as cCodCta,  C.nColocCalendEstado AS cEstado, P.nPrdEstado AS nEstCred, " _
        & "         P.NSALDO AS NSALDOK, C1.dVigencia as dFecVig, C1.nMontoCol as nCapDes, " _
        & "         nTasInt = (select nRFATasInt FROM colocaccredrfa AS cRFA WHERE cRFA.cPerscod = R.cPersCod ), " _
        & "         nDiaApr  =  (SELECT CE.nPlazo FROM COLOCACESTADO CE WHERE CE.CCTACOD = CC.CCTACOD AND nPrdEstado=2002), " _
        & "         nSalInt = (   SELECT SUM(CD1.NMONTO - CD1.NMONTOPAGADO) " _
        & "                         FROM COLOCCALENDDET CD1 " _
        & "                         WHERE CD1.CCTACOD = CC.CCTACOD AND CD1.nColocCalendApl=1 AND CD1.nNroCalen =CC.nNroCalen and CD1.nPrdConceptoCod in(1100,1107)), " _
        & "         nCuotas  = (    SELECT count(*) FROM COLOCCALENDARIO CC1 WHERE CC1.CCTACOD=CC.CCTACOD AND CC1.nColocCalendApl=1  AND CC1.nNroCalen=CC.nNroCalen), " _
        & "         C.DVENC as dFecVen, C.DPAGO,C.NCUOTA as cNrocuo, " _
        & "         nCapita= SUM(CASE WHEN CD.nPrdConceptoCod= 1000 OR CD.nPrdConceptoCod= 1010 THEN CD.nMonto ELSE 0 END), " _
        & "         nIntere= SUM(CASE WHEN CD.nPrdConceptoCod= 1100 OR CD.nPrdConceptoCod= 1107 THEN CD.nMonto ELSE 0 END), " _
        & "         nCofide= SUM(CASE WHEN CD.nPrdConceptoCod= 124352 THEN CD.nMonto ELSE 0 END), " _
        & "         nCuota= SUM(CASE WHEN CD.nPrdConceptoCod IN (1000,1100,1010,1107) THEN CD.nMonto ELSE 0 END), " _
        & "         nIntMor= SUM(CASE WHEN CD.nPrdConceptoCod = 1101 OR CD.nPrdConceptoCod = 1108 THEN CD.nMonto ELSE 0 END), " _
        & "         nCapPag= SUM(CASE WHEN CD.nPrdConceptoCod= 1000 OR  CD.nPrdConceptoCod= 1010 THEN CD.nMontoPagado ELSE 0 END), " _
        & "         nIntPag= SUM(CASE WHEN CD.nPrdConceptoCod= 1100  OR CD.nPrdConceptoCod= 1107 THEN CD.nMontoPagado ELSE 0 END), " _
        & "         nMorPag= SUM(CASE WHEN CD.nPrdConceptoCod = 1101 OR CD.nPrdConceptoCod = 1108 THEN CD.nMontoPagado ELSE 0 END) " _
        & " FROM    COLOCCALENDARIO C " _
        & "         JOIN COLOCCALENDDET CD ON CD.CCTACOD = C.CCTACOD AND CD.nNroCalen = C.nNroCalen AND CD.nColocCalendApl= C.nColocCalendApl AND CD.nCuota = C.nCuota " _
        & "         JOIN PRODUCTO P ON P.CCTACOD = C.CCTACOD " _
        & "         JOIN COLOCACCRED CC ON CC.CCTACOD = P.CCTACOD " _
        & "         JOIN COLOCACIONES C1 ON C1.CCTACOD = P.CCTACOD " _
        & "         JOIN PRODUCTOPERSONA R ON R.CCTACOD = P.CCTACOD AND R.nPrdPersRelac=20 " _
        & " WHERE   R.cPersCod='" & psPersCod & " ' AND C.nColocCalendApl <>0 AND CD.nNroCalen = CC.nNroCalen AND CC.CRFA='" & psRFA & "'  and CC.bRefinacRFA=" & IIf(pbRefinan = True, 1, 0) _
        & " GROUP BY R.cPersCod,C.CCTACOD, C.NCUOTA, C.DVENC, C.DPAGO, C.nColocCalendEstado, P.nPrdEstado, "
    sql = sql + "       P.nSaldo , C1.dVigencia, P.nTasaInteres, C1.nMontoCol, CC.Cctacod, CC.nNroCalen, C.DVENC " _
        & " ORDER BY C.NCUOTA"

End If
oCon.AbreConexion
Set GetCalendariosRFA = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetCRFA(ByVal psctacod As String) As String
    Dim sSql As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    
    Set oConec = New DConecta
    oConec.AbreConexion
    sSql = "Select cRFa From ColocacCred Where cCtaCod='" & psctacod & "'"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    If Not rs.EOF And Not rs.BOF Then
        GetCRFA = rs!cRFA
    End If
    Set rs = Nothing
End Function

Public Function RecuperaDatosExtorno(ByVal pnMovNro As Long, ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    sSql = "Select M.nCredEstado, SUM(MD.nMonto) as nCapital, "
    sSql = sSql & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
    sSql = sSql & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro AND MD.nPrdConceptoCod in(" & gColocRFACapital & "," & gColocConceptoCodCapital & "," & gColocRFAInteresDiferido & "," & gColocRFAInteresDiferidoRFA & " ) Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psctacod & "' AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psctacod & "' and M.nCredEstado<>0 "
    sSql = sSql & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSql)
    If RecuperaDatosExtorno.RecordCount = 0 Then
        sSql = "Select M.nCredEstado, 0 as nCapital, "
        sSql = sSql & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
        sSql = sSql & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro AND MD.nPrdConceptoCod = " & gColocRFAInteresCompesatorio & "," & gColocConceptoCodInteresCompensatorio & ") Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psctacod & "' AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psctacod & "' "
        sSql = sSql & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
        Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSql)
    
        If RecuperaDatosExtorno.RecordCount = 0 Then
            sSql = "Select M.nCredEstado, 0 as nCapital, "
            sSql = sSql & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
            sSql = sSql & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro AND Left(Cast(nPrdConceptoCod as Varchar(4)),2)='12' Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psctacod & "' AND Left(Cast(nPrdConceptoCod as Varchar(4)),2)='12' AND M.cCtaCod = '" & psctacod & "' "
            sSql = sSql & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
            Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSql)
        End If
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function


Public Function ListaCalendarioByPersona(ByVal psPersCod As String) As Recordset
    Dim oConec As DConecta
    Dim sSql  As String
    
    sSql = "SELECT *"
    sSql = sSql & "FROM ("
    sSql = sSql & " SELECT ORDEN =CASE CC.CRFA"
    sSql = sSql & " WHEN 'DIF'THEN 1"
    sSql = sSql & " WHEN 'RFC' THEN 2"
    sSql = sSql & " END,"
    sSql = sSql & " CA.NCUOTA,"
    sSql = sSql & " CA.CCTACOD,"
    sSql = sSql & " CAPITAL=(SELECT SUM(NMONTO-NMONTOPAGADO)"
    sSql = sSql & " FROM   COLOCCALENDDET CD WHERE  CD.CCTACOD=CA.CCTACOD AND CD.NNROCALEN=CA.NNROCALEN AND"
    sSql = sSql & " CD.NCOLOCCALENDAPL = CA.NCOLOCCALENDAPL And CD.nCuota = CA.nCuota"
    sSql = sSql & " AND CD.NPRDCONCEPTOCOD=1000),"
    sSql = sSql & " INTERES=(SELECT SUM(NMONTO-NMONTOPAGADO)"
    sSql = sSql & " FROM   COLOCCALENDDET CD WHERE CD.CCTACOD=CA.CCTACOD AND CD.NNROCALEN=CA.NNROCALEN AND"
    sSql = sSql & " CD.NCOLOCCALENDAPL = CA.NCOLOCCALENDAPL And CD.nCuota = CA.nCuota"
    sSql = sSql & " AND CD.NPRDCONCEPTOCOD=1100),"
    sSql = sSql & " MORA=(SELECT SUM(NMONTO-NMONTOPAGADO)"
    sSql = sSql & " FROM   COLOCCALENDDET CD  WHERE CD.CCTACOD=CA.CCTACOD AND CD.NNROCALEN=CA.NNROCALEN AND"
    sSql = sSql & " CD.NCOLOCCALENDAPL = CA.NCOLOCCALENDAPL And CD.nCuota = CA.nCuota"
    sSql = sSql & " AND  CD.NPRDCONCEPTOCOD=1101),"
    sSql = sSql & " COFIDE=(SELECT SUM(NMONTO-NMONTOPAGADO)"
    sSql = sSql & " FROM   COLOCCALENDDET CD WHERE CD.CCTACOD=CA.CCTACOD AND CD.NNROCALEN=CA.NNROCALEN AND"
    sSql = sSql & " CD.NCOLOCCALENDAPL = CA.NCOLOCCALENDAPL And CD.nCuota = CA.nCuota"
    sSql = sSql & " AND  CD.NPRDCONCEPTOCOD=124352),"
    sSql = sSql & " GASTOS=(SELECT SUM(NMONTO-NMONTOPAGADO)"
    sSql = sSql & " FROM   COLOCCALENDDET CD WHERE CD.CCTACOD=CA.CCTACOD AND CD.NNROCALEN=CA.NNROCALEN AND"
    sSql = sSql & " CD.NCOLOCCALENDAPL = CA.NCOLOCCALENDAPL And CD.nCuota = CA.nCuota"
    sSql = sSql & " AND CD.NPRDCONCEPTOCOD<>124352 AND CD.NPRDCONCEPTOCOD LIKE '12%'),"
    sSql = sSql & " CA.dVenc,"
    sSql = sSql & " CA.dPago "
    sSql = sSql & " FROM PRODUCTOPERSONA PP"
    sSql = sSql & " INNER JOIN COLOCACCRED CC ON CC.CCTACOD=PP.CCTACOD AND CC.CRFA IN ('RFC','DIF')"
    sSql = sSql & " INNER JOIN COLOCCALENDARIO CA ON CA.CCTACOD=CC.CCTACOD AND CA.NCOLOCCALENDAPL=1 AND CA.NNROCALEN=CC.NNROCALEN"
    sSql = sSql & " WHERE PP.CPERSCOD='" & psPersCod & "')X"
    sSql = sSql & " Union "
    sSql = sSql & " SELECT *"
    sSql = sSql & " FROM ("
    sSql = sSql & " SELECT 3 AS ORDEN,"
    sSql = sSql & " CA.NCUOTA,"
    sSql = sSql & " CA.CCTACOD,"
    sSql = sSql & " CAPITAL=(SELECT SUM(NMONTO-NMONTOPAGADO)"
    sSql = sSql & " FROM   COLOCCALENDDET CD WHERE  CD.CCTACOD=CA.CCTACOD AND CD.NNROCALEN=CA.NNROCALEN AND"
    sSql = sSql & " CD.NCOLOCCALENDAPL = CA.NCOLOCCALENDAPL And CD.nCuota = CA.nCuota"
    sSql = sSql & " AND CD.NPRDCONCEPTOCOD=1010),"
    sSql = sSql & " INTERES=(SELECT SUM(NMONTO-NMONTOPAGADO)"
    sSql = sSql & " FROM   COLOCCALENDDET CD WHERE CD.CCTACOD=CA.CCTACOD AND CD.NNROCALEN=CA.NNROCALEN AND"
    sSql = sSql & " CD.NCOLOCCALENDAPL = CA.NCOLOCCALENDAPL And CD.nCuota = CA.nCuota"
    sSql = sSql & " AND CD.NPRDCONCEPTOCOD=1107),"
    sSql = sSql & " MORA=(SELECT SUM(NMONTO-NMONTOPAGADO)"
    sSql = sSql & " FROM   COLOCCALENDDET CD  WHERE CD.CCTACOD=CA.CCTACOD AND CD.NNROCALEN=CA.NNROCALEN AND"
    sSql = sSql & " CD.NCOLOCCALENDAPL = CA.NCOLOCCALENDAPL And CD.nCuota = CA.nCuota"
    sSql = sSql & " AND  CD.NPRDCONCEPTOCOD=1108),"
    sSql = sSql & " COFIDE=(SELECT SUM(NMONTO-NMONTOPAGADO)"
    sSql = sSql & " FROM   COLOCCALENDDET CD WHERE CD.CCTACOD=CA.CCTACOD AND CD.NNROCALEN=CA.NNROCALEN AND"
    sSql = sSql & " CD.NCOLOCCALENDAPL = CA.NCOLOCCALENDAPL And CD.nCuota = CA.nCuota"
    sSql = sSql & " AND  CD.NPRDCONCEPTOCOD=124352),"
    sSql = sSql & " GASTOS=(SELECT SUM(NMONTO-NMONTOPAGADO)"
    sSql = sSql & " FROM   COLOCCALENDDET CD WHERE CD.CCTACOD=CA.CCTACOD AND CD.NNROCALEN=CA.NNROCALEN AND"
    sSql = sSql & " CD.NCOLOCCALENDAPL = CA.NCOLOCCALENDAPL And CD.nCuota = CA.nCuota"
    sSql = sSql & " AND CD.NPRDCONCEPTOCOD<>124352 AND CD.NPRDCONCEPTOCOD LIKE '12%'),"
    sSql = sSql & " CA.dVenc,"
    sSql = sSql & " CA.dPago "
    sSql = sSql & " FROM PRODUCTOPERSONA PP"
    sSql = sSql & " INNER JOIN COLOCACCRED CC ON CC.CCTACOD=PP.CCTACOD AND CC.CRFA IN ('RFA')"
    sSql = sSql & " INNER JOIN COLOCCALENDARIO CA ON CA.CCTACOD=CC.CCTACOD AND CA.NCOLOCCALENDAPL=1 AND CC.NNROCALEN=CA.NNROCALEN"
    sSql = sSql & " WHERE PP.CPERSCOD='" & psPersCod & "')X"
    sSql = sSql & " ORDER BY  ORDEN,NCUOTA,CCTACOD"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCalendarioByPersona = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Sub ImprimeBoletaExtorno(ByVal psDescrip As String, ByVal psctacod As String, ByVal psPersNombre As String, ByVal psNomAge As String, _
            ByVal psMoneda As String, ByVal psCuotasPagadas As String, ByVal psFecha As String, ByVal psHora As String, _
            ByVal psNroTransac As String, ByVal pnCapPagado As Double, _
            ByVal pnSaldoCap As Double, ByVal psCodUsu As String, ByVal sLpt As String, Optional ByVal psNomCmacAbrev As String = "", Optional ByVal psCodCMAC As String, _
            Optional ByVal psCodUserOri As String = "")
            
Dim nFicSal As Integer
Dim oDCred As DCredito
Dim R As ADODB.Recordset

    On Error GoTo ErrorImprimeBoleta
    vsCadImpresion = ""
    nFicSal = FreeFile

    Open sLpt For Output As nFicSal
    
    Set oDCred = New DCredito
    Set R = oDCred.RecuperaProducto(psctacod)
    Set oDCred = Nothing
    
    pnSaldoCap = R!nSaldo
    R.Close
    Set R = Nothing
    
    Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
    Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
    Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
    'Print #nFicSal, Chr$(27) + Chr$(15) ' condensada
    'Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
    Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
    
    Print #nFicSal, ""
    Call ImprimeCabeceraBoletaPago(psNomAge, psMoneda, psCuotasPagadas, _
            psFecha, psHora, psPersNombre, "", psctacod, "", psNomCmacAbrev, gsCodCMAC)
    
    'Cuerpo de la Boleta
If gsCodCMAC = "102" Then
'    vsCadImpresion = vsCadImpresion & lsNegritaOn
'    vsCadImpresion = vsCadImpresion & ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 40) & lsSaltoLin
'    '45 menos la longitud de su etiqueta
'    vsCadImpresion = vsCadImpresion & lsSaltoLin
'    vsCadImpresion = vsCadImpresion & "Monto  :" & Space(13) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
'    vsCadImpresion = vsCadImpresion & lsSaltoLin
'    vsCadImpresion = vsCadImpresion & String(40, "-") & lsSaltoLin
'    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
'    vsCadImpresion = vsCadImpresion & lsNegritaOff
'    vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(7) & ImpreFormat(pnSaldoCap, 12) & lsSaltoLin
'    'vsCadImpresion = vsCadImpresion & ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & psCodUsu, 53, 0) & ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & psCodUsu, 53, 0) & lsSaltoLin
'    vsCadImpresion = vsCadImpresion & lsSaltoLin
'    Print #nFicSal, vsCadImpresion
'    'Print #nFicSal, lsSaltoLin
'    Print #nFicSal, ""
'    Print #nFicSal, ""
'    Print #nFicSal, ""
'    Close #nFicSal
Else
vsCadImpresion = vsCadImpresion & lsNegritaOn
    vsCadImpresion = vsCadImpresion & ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 50)
    vsCadImpresion = vsCadImpresion & ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 40) & Chr(10)
    '45 menos la longitud de su etiqueta
    vsCadImpresion = vsCadImpresion & Chr(10)
    vsCadImpresion = vsCadImpresion & "Monto  :" & Space(13) & ImpreFormat(pnCapPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Monto   :" & Space(13) & ImpreFormat(pnCapPagado, 12) & Chr(10)
    vsCadImpresion = vsCadImpresion & Chr(10)
    vsCadImpresion = vsCadImpresion & String(40, "-")
    vsCadImpresion = vsCadImpresion & Space(14) & String(40, "-") & Chr(10)
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado, 12) & Chr(10)
    vsCadImpresion = vsCadImpresion & lsNegritaOff
    vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(7) & ImpreFormat(pnSaldoCap, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Saldo Capital :" & Space(7) & ImpreFormat(pnSaldoCap, 12) & Chr(10)
    
    vsCadImpresion = vsCadImpresion & "Usuario Exto. :" & Space(7) & ImpreFormat(psCodUsu, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Usuario Exto. :" & Space(7) & ImpreFormat(psCodUsu, 12) & Chr(10)
    
    vsCadImpresion = vsCadImpresion & "Usuario Gene. :" & Space(7) & ImpreFormat(psCodUserOri, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Usuario Gene. :" & Space(7) & ImpreFormat(psCodUserOri, 12) & Chr(10)
    'vsCadImpresion = vsCadImpresion & ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & psCodUsu, 53, 0) & ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & psCodUsu, 53, 0) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & Chr(10)
    Print #nFicSal, vsCadImpresion
    'Print #nFicSal, lsSaltoLin
    Print #nFicSal, ""
    Print #nFicSal, ""
    Print #nFicSal, ""
    Close #nFicSal
End If
    Exit Sub
ErrorImprimeBoleta:
    Err.Raise Err.Number, "Error En Proceso ImprimeBoleta", Err.Description
    
End Sub


Public Function ObtenNumeroCuotas(ByVal pnMovNro As Double) As Integer
    Dim sSql As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    
    sSql = "Select Count(distinct nNroCuota)-1 as nCantidad"
    sSql = sSql & " From MovColDet"
    sSql = sSql & " Where nMovNro=" & pnMovNro
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        ObtenNumeroCuotas = rs!nCantidad
    End If
    Set rs = Nothing
End Function

Private Sub ImprimeCabeceraBoletaPago(ByVal psNomAge As String, ByVal psMoneda As String, _
        ByVal psCuotasPagadas As String, ByVal psFecha As String, ByVal psHora As String, _
        ByVal psPersNombre As String, ByVal psNroTransac As String, ByVal psctacod As String, _
        ByVal psTipoCredito As String, Optional ByVal psNomCmacAbrev As String = "", Optional ByVal psCodCMAC As String)
       ' psNomAge = "Sede Institucional"
        On Error GoTo ErrorImprimeCabeceraBoletaPago
        'vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(64)
        'vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
        'vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
        'vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
        'vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
        'vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(18)  ' cancela condensada
        'vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)  ' desactiva negrita
        
        vsCadImpresion = vsCadImpresion & lsNegritaOn
        'vsCadImpresion = vsCadImpresion & lsSaltoLin
        'vsCadImpresion = vsCadImpresion & lsSaltoLin
        'vsCadImpresion = vsCadImpresion & lsSaltoLin
        If psCodCMAC = "102" Then
            vsCadImpresion = vsCadImpresion & ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & ImpreFormat(Mid(psTipoCredito, 1, 25), 51, 0), 53) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & lsNegritaOff
            vsCadImpresion = vsCadImpresion & ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & ImpreFormat("Credito : " & psctacod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & lsSaltoLin
        Else
            vsCadImpresion = vsCadImpresion & ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & ImpreFormat(Mid(psTipoCredito, 1, 25), 51, 0), 50, 0) & Trim(ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & ImpreFormat(Mid(psTipoCredito, 1, 25), 51, 0), 53, 0)) & Chr(10)
            vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 51, 0) & Trim(ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0)) & Chr(10)
            vsCadImpresion = vsCadImpresion & lsNegritaOff
            vsCadImpresion = vsCadImpresion & ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 51, 0) & Trim(ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0)) & Chr(10)
            vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psPersNombre, 1, 50), 51, 0) & Trim(ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0)) & Chr(10)
            vsCadImpresion = vsCadImpresion & ImpreFormat("Credito : " & psctacod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 51, 0) & ImpreFormat("Credito : " & psctacod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & Chr(10)
        End If
        
        Exit Sub

ErrorImprimeCabeceraBoletaPago:
        Err.Raise Err.Number, "Error En Proceso ImprimeCabeceraBoletaPago", Err.Description
    
End Sub

 Function ValidaCredito(ByVal psctacod As String) As Boolean
    'Verifica si el credito tiene montos disponibles
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select Count(*) as nCantidad From Producto "
    sSql = sSql & " Where cCtaCod='" & psctacod & "' And nPrdEstado in (2030,2031,2032) and nSaldo>0"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        If rs!nCantidad > 0 Then
            ValidaCredito = True
        Else
            ValidaCredito = False
        End If
    End If
   Set rs = Nothing
 End Function
 
 Public Function MatrizInteresGastosAFecha(ByVal psctacod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date, Optional ByVal pbDesagio As Boolean = False, Optional ByVal pbCalenDin As Boolean = False) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaColocaciones(psctacod)
    Set oCredito = Nothing
    
    dFecIni = R!dVigencia
    nSaldoCal = R!nMontoCol
    R.Close
    Set R = Nothing
        
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaProducto(psctacod)
    Set oCredito = Nothing
    pnTasa = R!nTasaInteres
    R.Close
    Set R = Nothing
    
    'Total Calculado a la fecha
    MatrizInteresGastosAFecha = CDbl(Format(MatrizGastosFecha(psctacod, MatCalend) + _
                MatrizInteresReprogramadoFecha(psctacod, MatCalend) + _
                MatrizInteresSuspensoFecha(psctacod, MatCalend) + _
                MatrizInteresGraciaFecha(psctacod, MatCalend, pdHoy, nSaldoCal) + _
                MatrizInteresMorFecha(psctacod, MatCalend) + _
                MatrizInteresCompAFecha(psctacod, MatCalend, pdHoy, nSaldoCal, pnTasa, pbDesagio, pbCalenDin) + MatrizInteresCompensatorioVencido(MatCalend), "#0.00"))
    
End Function
 
Public Function MatrizGastosFecha(ByVal psctacod As String, ByVal MatCalend As Variant) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 9)) 'Gastos
    Next i
    
    'Total Calculado es
    MatrizGastosFecha = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function

Public Function MatrizInteresReprogramadoFecha(ByVal psctacod As String, ByVal MatCalend As Variant) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 7)) 'Interes Reprogramado
    Next i
    
    'Total Calculado es
    MatrizInteresReprogramadoFecha = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function

Public Function MatrizInteresCompensatorioVencido(ByVal MatCalend As Variant) As Double
Dim nMontoGanado As Double
Dim i As Integer

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 11)) 'Interes Suspenso
    Next i
    
    'Total Calculado es
    MatrizInteresCompensatorioVencido = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function

Public Function MatrizInteresSuspensoFecha(ByVal psctacod As String, ByVal MatCalend As Variant) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 8)) 'Interes Suspenso
    Next i
    
    'Total Calculado es
    MatrizInteresSuspensoFecha = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function


Public Function MatrizInteresGraciaFecha(ByVal psctacod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date, _
    Optional pnMontoCol As Double = -1) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim i As Integer
Dim nSaldoCal As Double
Dim pnTasa As Double
Dim nDiasGracia As Integer
Dim oCalend As Dcalendario

    If pnMontoCol = -1 Then
        Set oCredito = New DCredito
        Set R = oCredito.RecuperaColocaciones(psctacod)
        Set oCredito = Nothing
        nSaldoCal = R!nMontoCol
        R.Close
        Set R = Nothing
    Else
        nSaldoCal = pnMontoCol
    End If
    
    Set oCalend = New Dcalendario
    dFecIni = oCalend.RecuperaFechaInicioCuota(psctacod, CInt(MatCalend(0, 1)), gColocCalendAplCuota)
    Set oCalend = Nothing
    
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaProductoTasaInteres(psctacod, gColocLineaCredTasasIntGracia)
    Set oCredito = Nothing
    If Not R.BOF And Not R.EOF Then
        pnTasa = R!nTasaInteres
    Else
        MatrizInteresGraciaFecha = 0
        R.Close
        Set R = Nothing
        Exit Function
    End If
    R.Close
    Set R = Nothing
    
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaColocacEstado(psctacod, gColocEstAprob)
    Set oCredito = Nothing
    nDiasGracia = R!nPeriodoGracia
    R.Close
    Set R = Nothing
    
    nMontoGanado = 0
    'Si es Cierto ya Vencio la Fecha de Gracia
    If Abs(CDate(Format(pdHoy, "dd/mm/yyyy")) - CDate(Format(dFecIni, "dd/mm/yyyy"))) >= nDiasGracia Then
        'Calculo de Montos Ya Ganados
        For i = 0 To UBound(MatCalend) - 1
            nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 5)) 'Interes Moratorio
        Next i
    Else
        nMontoGanado = TasaIntPerDias(pnTasa, CDate(Format(pdHoy, "dd/mm/yyyy")) - CDate(Format(dFecIni, "dd/mm/yyyy"))) * nSaldoCal
    End If
    'Total Calculado es
    MatrizInteresGraciaFecha = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function

Public Function MatrizInteresMorFecha(ByVal psctacod As String, ByVal MatCalend As Variant) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 6)) 'Interes Moratorio
    Next i
    
    'Total Calculado es
    MatrizInteresMorFecha = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function

Public Function MatrizInteresCompAFecha(ByVal psctacod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date, _
    Optional pnMontoCol As Double = -1, Optional pnTasaInteres As Double = -1, Optional ByVal pbDesagio As Boolean = False, Optional ByVal pbCalenDin As Boolean = False) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim oCalend As Dcalendario
Dim nMontoGanado As Double
Dim nMontoAFecha As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double
Dim nColocCalendCod As Integer
Dim sSql As String
Dim oConec As DConecta
Dim nDiasGracia As Long

    'Calculo de Montos Ya Ganados


    Set oCredito = New DCredito
    Set R = oCredito.RecuperaColocacCred(psctacod)
    Set oCredito = Nothing
    nColocCalendCod = R!nColocCalendCod
    R.Close



    nMontoGanado = 0
    nSaldoCap = 0
    nMontoAFecha = 0

    If nColocCalendCod <> 70 Then
        For i = 0 To UBound(MatCalend) - 1
            If pdHoy >= CDate(MatCalend(i, 0)) Then
                nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 4))
            End If
        Next i
    Else
        nMontoGanado = nMontoGanado + CDbl(MatCalend(0, 4))
    End If

    If nColocCalendCod <> 70 Then
        If pnMontoCol = -1 Then
            Set oCredito = New DCredito
            Set R = oCredito.RecuperaColocaciones(psctacod)
            Set oCredito = Nothing
            nSaldoCal = R!nMontoCol
            R.Close
            Set R = Nothing
        Else
            If pbCalenDin Then
                Dim oDCalenCred As Dcalendario
                Set oDCalenCred = New Dcalendario
                Set R = oDCalenCred.RecuperaCalendarioDesemb(psctacod)
                Set oDCalenCred = Nothing
                nSaldoCal = 0
                Do While Not R.EOF
                    nSaldoCal = nSaldoCal + R!nCapital
                    R.MoveNext
                Loop
                R.Close
                Set R = Nothing
            Else
                nSaldoCal = pnMontoCol
            End If
        End If
        
        Set oCalend = New Dcalendario
        dFecIni = oCalend.RecuperaFechaInicioCuota(psctacod, CInt(MatCalend(0, 1)), gColocCalendAplCuota, pbCalenDin)
        Set oCalend = Nothing

        If pnTasaInteres = -1 Then
            Set oCredito = New DCredito
            Set R = oCredito.RecuperaProducto(psctacod)
            Set oCredito = Nothing
            pnTasa = R!nTasaInteres
            R.Close
            Set R = Nothing
        Else
            pnTasa = pnTasaInteres
        End If

        Set oCredito = New DCredito
        Set R = oCredito.RecuperaColocacEstado(psctacod, gColocEstAprob)
        Set oCredito = Nothing
        nDiasGracia = R!nPeriodoGracia
        R.Close
        Set R = Nothing
        
        'Interes a la fecha
        nMontoAFecha = 0
        For i = 0 To UBound(MatCalend) - 1
            If pdHoy >= CDate(MatCalend(i, 0)) Then
                dFecIni = CDate(MatCalend(i, 0))
            Else
                If (pdHoy - dFecIni) > 0 Then
                    If MatCalend(i, 1) = "1" Then
                        nMontoAFecha = TasaIntPerDias(pnTasa, (pdHoy - dFecIni) - nDiasGracia) * nSaldoCal
                        'nMontoAFecha = TasaIntPerDias(pnTasa, pdHoy -dFecIni) * (CDbl(MatCalend(I, 10)) + CDbl(MatCalend(I, 3)))
                    Else
                        nMontoAFecha = TasaIntPerDias(pnTasa, pdHoy - dFecIni) * CDbl(SaldoK(psctacod, pdHoy))
                        'nMontoAFecha = TasaIntPerDias(pnTasa, pdHoy - dFecIni) * CDbl(MatCalend(I, 10))
                    End If
                End If
                sSql = "Select isnull(nMontoPagado,0)  as nMontoPagado From ColocCalendDet Where cCtaCod = '" & psctacod & "' AND nColocCalendApl = 1 "
                sSql = sSql & " AND nNroCalen =  (Select nNroCalen from ColocacCred Where cctaCod = '" & psctacod & "')"
                sSql = sSql & " AND nPrdConceptoCod = 1107 AND nCuota = " & MatCalend(i, 1)
                sSql = "Select nMontoPagado From ColocCalendDet Where cCtaCod = '" & psctacod & "' AND nColocCalendApl = 1 "
                sSql = sSql & " AND nNroCalen =  (Select nNroCalen from ColocacCred Where cctaCod = '" & psctacod & "')"
                sSql = sSql & " AND nPrdConceptoCod = 1107 AND nCuota = " & MatCalend(i, 1)
                Set oConec = New DConecta
                oConec.AbreConexion
                Set R = oConec.CargaRecordSet(sSql)
                Dim nMonto As Double
                If Not R.EOF And Not R.BOF Then
                    nMonto = IIf(IsNull(R!nMontoPagado), 0, R!nMontoPagado)
                Else
                    nMonto = 0
                End If
                
                nMontoAFecha = nMontoAFecha - nMonto
                If Not pbDesagio And nMontoAFecha < 0 Then
                    nMontoAFecha = 0
                End If
                oConec.CierraConexion
                Set oConec = Nothing
                Exit For
            End If
        Next i
    End If
    'Total Calculado es
    MatrizInteresCompAFecha = CDbl(Format(nMontoGanado + nMontoAFecha, "#0.00"))

End Function

Public Function SaldoK(ByVal pcCtaCod As String, ByVal dPago As Date) As Currency
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    'Calculo del SaldoCapital
        sSql = "Select Top 1 (Select Sum(nMonto) From ColocCalendDet CD Where  C.cCtaCod=CD.cCtaCod and"
        sSql = sSql & " C.nNroCalen = CD.nNroCalen And CD.cCtaCod = C.cCtaCod"
        sSql = sSql & " and CD.nPrdConceptoCod in(1000,1010) and CD.nColocCalendApl=0)-"
        sSql = sSql & " (Select Sum(nMontoPagado) From ColocCalendDet CD Where  C.cCtaCod=CD.cCtaCod and"
        sSql = sSql & " C.nNroCalen = CD.nNroCalen And CD.cCtaCod = C.cCtaCod"
        sSql = sSql & " and CD.nPrdConceptoCod in(1000 ,1010)and CD.nColocCalendApl=1) as SK"
        sSql = sSql & " From colocCalendario C"
        sSql = sSql & " Where C.cCtaCod='" & pcCtaCod & "' and  nNroCalen=(Select nNroCalen From ColocacCred Where cCtaCod='" & pcCtaCod & "')"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    
    If Not rs.EOF And Not rs.BOF Then
        SaldoK = rs!SK
    End If
    'En el Caso pague cuotas atrasada y cancele el credito
    sSql = "Select isnull(Sum(nMonto),0) as Capital"
    sSql = sSql & " From ColocCalendDet a "
    sSql = sSql & " inner join ColocCalendario b on a.cctacod=b.cctacod and a.nnrocalen=b.nnrocalen and "
    sSql = sSql & " a.nColocCalendApl=b.nColocCalendApl and a.ncuota=b.nCuota "
    sSql = sSql & " Where a.nPrdConceptoCod in(1000,1010) and b.dVenc<='" & Format(CStr(dPago), "MM/dd/yyyy") & "' and b.nColocCalendEstado=0  and b.nColocCalendApl=1"
    sSql = sSql & " and a.cctacod='" & pcCtaCod & "' and a.nNroCalen=(Select nNroCalen From ColocacCred Where cCtaCod='" & pcCtaCod & "')"
    
    Set rs = New ADODB.Recordset
    Set rs = oConec.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
        SaldoK = SaldoK - IIf(IsNull(rs!Capital), 0, rs!Capital)
    End If
    Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function TasaIntPerDias(ByVal pnTasaInter As Double, ByVal pnDiasTrans As Integer) As Double
    TasaIntPerDias = ((1 + pnTasaInter / 100) ^ (pnDiasTrans / 30)) - 1
End Function

Public Function MontoIntPerDias(ByVal pnTasaInter As Double, ByVal pnDiasTrans As Integer, ByVal pnMonto As Double) As Double
    MontoIntPerDias = (((1 + pnTasaInter / 100) ^ (pnDiasTrans / 30)) - 1) * pnMonto
    MontoIntPerDias = CDbl(Format(MontoIntPerDias, "#0.00"))
End Function

Public Function MatrizDistribuirCancelacion(ByVal psctacod As String, ByVal MatCalend As Variant, ByVal pnMontoPago As Double, _
                 ByVal psMetLiquidacion As String, ByVal pdHoy As Date, Optional ByVal pbNoCancelar As Boolean = False, Optional ByVal pbCalenDin As Boolean = False) As Variant

Dim MatCalendDistrib As Variant
Dim nMontoGastos As Double
Dim nMontoMora As Double
Dim nMontoInteres As Double
Dim nMontoCapital As Double
Dim nMontoTotalTmp As Double
Dim nMontoPagoTmp As Double
Dim J As Integer

        MatCalendDistrib = CrearMatrizparaAmortizacion(MatCalend)
        nMontoTotalTmp = 0
        nMontoPagoTmp = pnMontoPago
        For J = 1 To 4
                Select Case Mid(psMetLiquidacion, J, 1)
                    Case "G"
                        nMontoGastos = MatrizGastosFecha(psctacod, MatCalend)
                        If nMontoPagoTmp > nMontoGastos Then
                            nMontoPagoTmp = nMontoPagoTmp - nMontoGastos
                        Else
                            nMontoGastos = nMontoPagoTmp
                            nMontoPagoTmp = 0#
                        End If
                        nMontoTotalTmp = nMontoTotalTmp + nMontoGastos
                        Call MatrizDistribuirGastos(MatCalend, MatCalendDistrib, nMontoGastos, True)
                    Case "M"
                        nMontoMora = MatrizInteresMorFecha(psctacod, MatCalend)
                        If nMontoPagoTmp > nMontoMora Then
                            nMontoPagoTmp = nMontoPagoTmp - nMontoMora
                        Else
                            nMontoMora = nMontoPagoTmp
                            nMontoPagoTmp = 0#
                        End If
                        nMontoTotalTmp = nMontoTotalTmp + nMontoMora
                        Call MatrizDistribuirMora(MatCalend, MatCalendDistrib, nMontoMora, True)
                    Case "I", "i", "Y"
                        nMontoInteres = CDbl(Format(MatrizInteresTotalesAFechaSinMora(psctacod, MatCalend, pdHoy, pbCalenDin), "#0.00"))
                        If nMontoPagoTmp > nMontoInteres Then
                            nMontoPagoTmp = nMontoPagoTmp - nMontoInteres
                        Else
                            nMontoInteres = nMontoPagoTmp
                            nMontoPagoTmp = 0#
                        End If
                        nMontoTotalTmp = nMontoTotalTmp + nMontoInteres
                        Call MatrizDistribuirInteresI(MatCalend, MatCalendDistrib, nMontoInteres, True)
                    Case "C"
                        nMontoCapital = MatrizCapitalAFecha(psctacod, MatCalend)
                        If nMontoPagoTmp > nMontoCapital Then
                            nMontoPagoTmp = nMontoPagoTmp - nMontoCapital
                        Else
                            nMontoCapital = nMontoPagoTmp
                            nMontoPagoTmp = 0#
                        End If
                        nMontoTotalTmp = nMontoTotalTmp + nMontoCapital
                        Call MatrizDistribuirCapital(MatCalend, MatCalendDistrib, nMontoCapital, True)
                End Select
        Next J
        If (Mid(psMetLiquidacion, 3, 1) = "i" Or Mid(psMetLiquidacion, 3, 1) = "Y") And pbNoCancelar Then
            If Mid(psMetLiquidacion, 3, 1) = "Y" Then
                Call MatrizActualizarEstadoCalendCancelado(MatCalendDistrib, True, pbNoCancelar, MatCalend)
            Else
                Call MatrizActualizarEstadoCalendCancelado(MatCalendDistrib, False, pbNoCancelar, MatCalend)
            End If
        Else
            Call MatrizActualizarEstadoCalendCancelado(MatCalendDistrib, , False)
        End If
        MatrizDistribuirCancelacion = MatCalendDistrib
End Function
Private Sub MatrizActualizarEstadoCalendCancelado(ByRef MatCalendDistrib As Variant, Optional ByVal pbVerficaSolosCapital As Boolean = False, _
    Optional ByVal pbNoCancelar As Boolean = True, Optional ByVal MatCalend As Variant = "")
Dim i As Integer

    For i = 0 To UBound(MatCalendDistrib) - 1
        If pbNoCancelar Then
            
            If pbVerficaSolosCapital Then
                If CInt(MatCalendDistrib(i, 2)) = gColocCalendEstadoPendiente Then
                    If CDbl(MatCalend(i, 3)) = CDbl(MatCalendDistrib(i, 3)) Then
                        MatCalendDistrib(i, 2) = Trim(Str(gColocCalendEstadoPagado))
                    Else
                        MatCalendDistrib(i, 2) = Trim(Str(gColocCalendEstadoPendiente))
                    End If
                End If
            Else
                If CDbl(MatCalend(i, 8)) = CDbl(MatCalendDistrib(i, 8)) And _
                    CDbl(MatCalend(i, 7)) = CDbl(MatCalendDistrib(i, 7)) And _
                    CDbl(MatCalend(i, 6)) = CDbl(MatCalendDistrib(i, 6)) And _
                    CDbl(MatCalend(i, 5)) = CDbl(MatCalendDistrib(i, 5)) And _
                    CDbl(MatCalend(i, 3)) = CDbl(MatCalendDistrib(i, 3)) Then
                    MatCalendDistrib(i, 2) = Trim(Str(gColocCalendEstadoPagado))
                End If
            End If
        Else
            MatCalendDistrib(i, 2) = Trim(Str(gColocCalendEstadoPagado))
        End If
    Next i
End Sub

Public Function CrearMatrizparaAmortizacion(ByVal MatCalend As Variant) As Variant
Dim MatCalendAmortiz() As String
Dim i As Integer
    ReDim MatCalendAmortiz(UBound(MatCalend), 12)
    
    For i = 0 To UBound(MatCalend) - 1
        MatCalendAmortiz(i, 0) = MatCalend(i, 0)
        MatCalendAmortiz(i, 1) = MatCalend(i, 1)
        MatCalendAmortiz(i, 2) = MatCalend(i, 2)
        MatCalendAmortiz(i, 3) = "0.00"
        MatCalendAmortiz(i, 4) = "0.00"
        MatCalendAmortiz(i, 5) = "0.00"
        MatCalendAmortiz(i, 6) = "0.00"
        MatCalendAmortiz(i, 7) = "0.00"
        MatCalendAmortiz(i, 8) = "0.00"
        MatCalendAmortiz(i, 9) = "0.00"
        MatCalendAmortiz(i, 11) = "0.00"
        MatCalendAmortiz(i, 10) = MatCalend(i, 10)
    Next i
    CrearMatrizparaAmortizacion = MatCalendAmortiz
End Function

Private Function MatrizDistribuirGastos(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, ByRef pnMonto As Double, _
        Optional ByVal DistVert As Boolean = False)
Dim i As Integer

    For i = 0 To UBound(MatCalend) - 1
        If i > 0 And Not DistVert Then
            'Si Aun queda pendiente capital, interes, gastos, mora de la cuota anterior
            If (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 3)) - CDbl(MatCalendDistrib(i - 1, 3))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 8)) - CDbl(MatCalendDistrib(i - 1, 8))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 7)) - CDbl(MatCalendDistrib(i - 1, 7))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 5)) - CDbl(MatCalendDistrib(i - 1, 5))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 4)) - CDbl(MatCalendDistrib(i - 1, 4))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 6)) - CDbl(MatCalendDistrib(i - 1, 6))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 And _
                (CDbl(MatCalend(i - 1, 9)) - CDbl(MatCalendDistrib(i - 1, 9))) > 0) Then
                
                Exit Function
            End If
        End If
    
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 And (CDbl(MatCalend(i, 9)) - CDbl(MatCalendDistrib(i, 9))) > 0 Then
            If pnMonto > (CDbl(MatCalend(i, 9)) - CDbl(MatCalendDistrib(i, 9))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 9)) - CDbl(MatCalendDistrib(i, 9)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 9) = MatCalend(i, 9)
            Else
                MatCalendDistrib(i, 9) = Format(CDbl(MatCalendDistrib(i, 9)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
            If Not DistVert Then
                Exit For
            End If
        End If
     Next
End Function
        
        '*********************************************************
'** Cubrir Verticalmente Todas las Moras de las Cuotas
'*********************************************************
Private Function MatrizDistribuirMora(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, _
ByRef pnMonto As Double, Optional ByVal DistVert As Boolean = False)
Dim i As Integer

    For i = 0 To UBound(MatCalend) - 1
        
        If i > 0 And Not DistVert Then
            'Si Aun queda pendiente capital, interes, gastos, mora de la cuota anterior
            If (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 3)) - CDbl(MatCalendDistrib(i - 1, 3))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 8)) - CDbl(MatCalendDistrib(i - 1, 8))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 7)) - CDbl(MatCalendDistrib(i - 1, 7))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 5)) - CDbl(MatCalendDistrib(i - 1, 5))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 4)) - CDbl(MatCalendDistrib(i - 1, 4))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 6)) - CDbl(MatCalendDistrib(i - 1, 6))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 And _
                (CDbl(MatCalend(i - 1, 9)) - CDbl(MatCalendDistrib(i - 1, 9))) > 0) Then
                
                Exit Function
            End If
        End If
        
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(i, 6)) - CDbl(MatCalendDistrib(i, 6))) > 0 Then
            If pnMonto > (CDbl(MatCalend(i, 6)) - CDbl(MatCalendDistrib(i, 6))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 6)) - CDbl(MatCalendDistrib(i, 6)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 6) = MatCalend(i, 6)
            Else
                MatCalendDistrib(i, 6) = Format(CDbl(MatCalendDistrib(i, 6)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
    Next i
  
End Function

'Public Function MatrizInteresGastosAFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date, Optional ByVal pbDesagio As Boolean = False, Optional ByVal pbCalenDin As Boolean = False) As Double
'Dim R As ADODB.Recordset
'Dim oCredito As DCredito
'Dim nMontoGanado As Double
'Dim dFecini As Date
'Dim nSaldoCal As Double
'Dim pnTasa As Double
'
'    'Calculo de Montos Ya Ganados
'
'    Set oCredito = New DCredito
'    Set R = oCredito.RecuperaColocaciones(psCtaCod)
'    Set oCredito = Nothing
'
'    dFecini = R!dVigencia
'    nSaldoCal = R!nMontoCol
'    R.Close
'    Set R = Nothing
'
'    Set oCredito = New DCredito
'    Set R = oCredito.RecuperaProducto(psCtaCod)
'    Set oCredito = Nothing
'    pnTasa = R!nTasaInteres
'    R.Close
'    Set R = Nothing
'
'    'Total Calculado a la fecha
'    MatrizInteresGastosAFecha = CDbl(Format(MatrizGastosFecha(psCtaCod, MatCalend) + _
'                MatrizInteresReprogramadoFecha(psCtaCod, MatCalend) + _
'                MatrizInteresSuspensoFecha(psCtaCod, MatCalend) + _
'                MatrizInteresGraciaFecha(psCtaCod, MatCalend, pdHoy, nSaldoCal) + _
'                MatrizInteresMorFecha(psCtaCod, MatCalend) + _
'                MatrizInteresCompAFecha(psCtaCod, MatCalend, pdHoy, nSaldoCal, pnTasa, pbDesagio, pbCalenDin) + MatrizInteresCompensatorioVencido(MatCalend), "#0.00"))
'
'End Function

Private Function MatrizDistribuirInteresI(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, _
    ByRef pnMonto As Double, Optional ByVal DistVert As Boolean = False)
Dim i As Integer
Dim bSiCubrio As Boolean
    bSiCubrio = False
    For i = 0 To UBound(MatCalend) - 1
        If i > 0 And Not DistVert Then
            'Si Aun queda pendiente capital, interes, gastos, mora de la cuota anterior
            If (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 3)) - CDbl(MatCalendDistrib(i - 1, 3))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 8)) - CDbl(MatCalendDistrib(i - 1, 8))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 7)) - CDbl(MatCalendDistrib(i - 1, 7))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 5)) - CDbl(MatCalendDistrib(i - 1, 5))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 4)) - CDbl(MatCalendDistrib(i - 1, 4))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 6)) - CDbl(MatCalendDistrib(i - 1, 6))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 And _
                (CDbl(MatCalend(i - 1, 9)) - CDbl(MatCalendDistrib(i - 1, 9))) > 0) Then
                
                Exit Function
            End If
        End If
    
        'Cubre Interes en Suspenso
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(i, 8)) - CDbl(MatCalendDistrib(i, 8))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(i, 8)) - CDbl(MatCalendDistrib(i, 8))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 8)) - CDbl(MatCalendDistrib(i, 8)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 8) = MatCalend(i, 8)
            Else
                MatCalendDistrib(i, 8) = Format(CDbl(MatCalendDistrib(i, 8)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        'Cubre Intereses Reprogramados
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(i, 7)) - CDbl(MatCalendDistrib(i, 7))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(i, 7)) - CDbl(MatCalendDistrib(i, 7))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 7)) - CDbl(MatCalendDistrib(i, 7)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 7) = MatCalend(i, 7)
            Else
                MatCalendDistrib(i, 7) = Format(CDbl(MatCalendDistrib(i, 7)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        'Intereses de Gracia
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(i, 5)) - CDbl(MatCalendDistrib(i, 5))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(i, 5)) - CDbl(MatCalendDistrib(i, 5))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 5)) - CDbl(MatCalendDistrib(i, 5)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 5) = MatCalend(i, 5)
            Else
                MatCalendDistrib(i, 5) = Format(CDbl(MatCalendDistrib(i, 5)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        'Intereses Compensatorio Vencido
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(i, 11)) - CDbl(MatCalendDistrib(i, 11))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(i, 11)) - CDbl(MatCalendDistrib(i, 11))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 11)) - CDbl(MatCalendDistrib(i, 11)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 11) = MatCalend(i, 11)
            Else
                MatCalendDistrib(i, 11) = Format(CDbl(MatCalendDistrib(i, 11)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        'Intereses Compensatorios
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(i, 4)) - CDbl(MatCalendDistrib(i, 4))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(i, 4)) - CDbl(MatCalendDistrib(i, 4))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 4)) - CDbl(MatCalendDistrib(i, 4)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 4) = MatCalend(i, 4)
            Else
                MatCalendDistrib(i, 4) = Format(CDbl(MatCalendDistrib(i, 4)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        If bSiCubrio Then
            If Not DistVert Then
                Exit For
            End If
        End If
    Next i
End Function


Public Function MatrizCapitalAFecha(ByVal psctacod As String, ByVal MatCalend As Variant) As Double
Dim nMontoGanado As Double
Dim i As Integer
    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 3)) 'Capital
    Next i
    
    'Total Calculado es
    MatrizCapitalAFecha = CDbl(Format(nMontoGanado, "#0.00"))

End Function
Private Function MatrizDistribuirCapital(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, _
    ByRef pnMonto As Double, Optional ByVal DistVert As Boolean = False)
Dim i As Integer

    For i = 0 To UBound(MatCalend) - 1
    
        If i > 0 And Not DistVert Then
            'Si Aun queda pendiente capital, interes, gastos, mora de la cuota anterior
            If (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 3)) - CDbl(MatCalendDistrib(i - 1, 3))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 8)) - CDbl(MatCalendDistrib(i - 1, 8))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 7)) - CDbl(MatCalendDistrib(i - 1, 7))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 5)) - CDbl(MatCalendDistrib(i - 1, 5))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 4)) - CDbl(MatCalendDistrib(i - 1, 4))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 6)) - CDbl(MatCalendDistrib(i - 1, 6))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 And _
                (CDbl(MatCalend(i - 1, 9)) - CDbl(MatCalendDistrib(i - 1, 9))) > 0) Then
                
                Exit Function
            End If
        End If
    
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(i, 3)) - CDbl(MatCalendDistrib(i, 3))) > 0 Then
            If pnMonto > (CDbl(MatCalend(i, 3)) - CDbl(MatCalendDistrib(i, 3))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 3)) - CDbl(MatCalendDistrib(i, 3)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 3) = MatCalend(i, 3)
            Else
                MatCalendDistrib(i, 3) = Format(CDbl(MatCalendDistrib(i, 3)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
            If Not DistVert Then
                Exit For
            End If
        End If
    Next i
End Function


Public Function MatrizInteresTotalesAFechaSinMora(ByVal psctacod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date, Optional ByVal pbCalenDin As Boolean = False) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaColocaciones(psctacod)
    Set oCredito = Nothing
    
    dFecIni = R!dVigencia
    nSaldoCal = R!nMontoCol
    R.Close
    Set R = Nothing
        
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaProducto(psctacod)
    Set oCredito = Nothing
    pnTasa = R!nTasaInteres
    R.Close
    Set R = Nothing
    
    'Total Calculado a la fecha
    MatrizInteresTotalesAFechaSinMora = CDbl(Format(MatrizInteresReprogramadoFecha(psctacod, MatCalend) + _
                MatrizInteresSuspensoFecha(psctacod, MatCalend) + _
                MatrizInteresGraciaFecha(psctacod, MatCalend, pdHoy, nSaldoCal) + _
                MatrizInteresCompAFecha(psctacod, MatCalend, pdHoy, nSaldoCal, pnTasa, True, pbCalenDin) + MatrizInteresCompensatorioVencido(MatCalend), "#0.00"))
End Function

Public Function GetNroCalend(ByVal psctacod As String) As Integer
    Dim sSql As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    
    sSql = "Select nNroCalen"
    sSql = sSql & " From Colocaccred"
    sSql = sSql & " where cctacod='" & psctacod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        GetNroCalend = rs!nNroCalen
    End If
    Set rs = Nothing
End Function

Public Sub ActualizaCalendarioReprog(ByVal psctacod As String)
    Dim oConec As DConecta
    Dim sSql As String
    
    
    sSql = "Update ColocCalendario"
    sSql = sSql & " Set dPago = X.dPago"
    sSql = sSql & " From ColocCalendario C"
    sSql = sSql & " Inner Join (Select nCuota,dPago,C.cCtaCod ,C.nNroCalen"
    sSql = sSql & " From ColocCalendario C"
    sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=C.cCtaCod and C.nNroCalen=CC.nNroCalen -1"
    sSql = sSql & " where C.cctaCod='" & psctacod & "'  and C.nColocCalendEstado=1 and C.nColocCalendApl=1) X on X.cCtaCod= C.cCtaCod and X.nNroCalen+1=C.nNroCalen and X.nCuota=C.nCuota"
    sSql = sSql & " Where C.nColocCalendApl=1"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    oConec.ConexionActiva.Execute sSql
    oConec.CierraConexion
    Set oConec = Nothing
    
End Sub

