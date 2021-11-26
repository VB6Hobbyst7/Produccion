VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCredDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RecuperaDatosMantCargoAutomatico(ByVal psCtaCod As String) As ADODB.Recordset
Dim SQL As String
'Dim Rs As New ADODB.Recordset
'Dim RCD As nRcdReportes
'Set RCD = New nRcdReportes
Dim oConecta As COMConecta.DCOMConecta

SQL = " Select Pro.cCtaCod ,P1.cPersCod, P1.cPersNombre Nombre, Pro.nSaldo Saldo from Producto Pro"
SQL = SQL & " Inner Join ProductoPersona PP1 on PP1.cCtaCod = Pro.cCtaCod"
SQL = SQL & " Inner Join Persona P1 on P1.cPersCod = PP1.cPersCod"
SQL = SQL & " Where PP1.nPrdPersRelac = 20"
SQL = SQL & " and Pro.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
SQL = SQL & " and Pro.cCtaCod='" & psCtaCod & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaDatosMantCargoAutomatico = oConecta.CargaRecordSet(SQL)
oConecta.CierraConexion

Set oConecta = Nothing
'Set Rs = RCD.GetQuery(Sql)
'Set RecuperaDatosMantCargoAutomatico = Rs
'Set Rs = Nothing
'Set RCD = Nothing
End Function

Public Function RecuperaDatosColocCargoAutoma(ByVal psCtaCod As String) As ADODB.Recordset
Dim SQL As String
'Dim Rs As New ADODB.Recordset
'Dim RCD As nRcdReportes
'Set RCD = New nRcdReportes
Dim oConecta As COMConecta.DCOMConecta

SQL = " Select cCtaCod,nOrden,cCodCtaAho,cEstado  from ColocCargoAutoma"
SQL = SQL & " where cCtaCod='" & psCtaCod & "' and substring(cCtaCod,9,1)='" & Mid(psCtaCod, 9, 1) & "'"
SQL = SQL & " Order by cEstado, nOrden"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaDatosColocCargoAutoma = oConecta.CargaRecordSet(SQL)
oConecta.CierraConexion

Set oConecta = Nothing

'Set Rs = RCD.GetQuery(Sql)
'Set RecuperaDatosColocCargoAutoma = Rs
'Set Rs = Nothing
'Set RCD = Nothing
End Function

Public Sub ActualizaColocCargoAutoma(ByVal psCtaCod As String, ByVal psCtaCodAho As String, _
ByVal pnOrden As Integer, ByVal pnEstado As Integer)
Dim SQL As String
Dim Co As COMConecta.DCOMConecta
Set Co = New COMConecta.DCOMConecta
SQL = " Update ColocCargoAutoma"
SQL = SQL & " Set nOrden=" & pnOrden & ", cEstado = " & pnEstado
SQL = SQL & " where "
SQL = SQL & " cCtaCod ='" & Trim(psCtaCod) & "'"
SQL = SQL & " and cCodCtaAho ='" & Trim(psCtaCodAho) & "'"
Co.AbreConexion
Call Co.Ejecutar(SQL)
Co.CierraConexion
Set Co = Nothing

End Sub
Public Sub DeleteColocCargoAutoma(ByVal psCtaCod As String, ByVal psCtaCodAho As String)
Dim SQL As String
Dim Co As COMConecta.DCOMConecta
Set Co = New COMConecta.DCOMConecta
SQL = " Delete ColocCargoAutoma"
SQL = SQL & " where cEstado = '1'"
SQL = SQL & " and cCtaCod ='" & Trim(psCtaCod) & "'"
SQL = SQL & " and cCodCtaAho ='" & Trim(psCtaCodAho) & "'"
Co.AbreConexion
Call Co.Ejecutar(SQL)
Co.CierraConexion
Set Co = Nothing
End Sub

Public Function ExisteCtaCargoAutoma(ByVal psCtaCod As String, ByVal psCtaCodAho As String) As Boolean
Dim SQL As String
Dim rs As ADODB.Recordset
'Dim RCD As nRcdReportes
'Set RCD = New nRcdReportes
Dim oConecta As COMConecta.DCOMConecta

SQL = " Select * from ColocCargoAutoma"
SQL = SQL & " where cEstado = '1' "
SQL = SQL & " and cCtaCod ='" & Trim(psCtaCod) & "' "
SQL = SQL & " and cCodCtaAho ='" & Trim(psCtaCodAho) & "' "

Set oConecta = New COMConecta.DCOMConecta

oConecta.AbreConexion
Set rs = oConecta.CargaRecordSet(SQL)
oConecta.CierraConexion
Set oConecta = Nothing

'Set Rs = RCD.GetQuery(Sql)
If rs.EOF And rs.BOF Then
    ExisteCtaCargoAutoma = False
Else
    ExisteCtaCargoAutoma = True
End If
'Set Rs = Nothing
'Set RCD = Nothing
End Function

Public Sub InsertaColocCargoAutoma(ByVal psCtaCod As String, _
ByVal psCtaCodAho As String, ByVal psEstado As String, ByVal pnOrden As Integer)
Dim Co As COMConecta.DCOMConecta
Dim SQL As String
Set Co = New COMConecta.DCOMConecta
SQL = " Insert ColocCargoAutoma "
SQL = SQL & " (cCtaCod,nOrden,cCodCtaAho,cEstado) Values "
SQL = SQL & " ('" & psCtaCod & "'," & pnOrden & ",'" & psCtaCodAho & "','" & psEstado & "')"
Co.AbreConexion
Co.Ejecutar (SQL)
Co.CierraConexion
Set Co = Nothing
End Sub
'By Capi Set 07 Planeamiento
Public Function GetInformacionFepCmac(ByVal psInforme As String, ByVal pnTipoCambio As Currency, ByVal pnMes As Integer, ByVal pnAno As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = " Exec Genera_InformacionFepCmac '" & psInforme & "'," & pnTipoCambio & "," & pnMes & "," & pnAno
    Set GetInformacionFepCmac = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
 End Function

Public Function RecuperaDatosSolicitudCredito(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

'    ssql = "Select Pers.cPersCod, Pers.cPersNombre, PID.cPersIDNro DNI,PID2.cPersIDNro RUC, "
'    ssql = ssql & " Pers.cPersDireccDomicilio, PNat.nPersnatHijos, Pers.dPersNacCreac, PNat.nPersNatEstCiv, "
'    ssql = ssql & " CN.cConsDescripcion cEstadoCivil, PFI.cPersNombre cRazonSocial, PFI.cPersDireccDomicilio cDireccRazon, "
'    ssql = ssql & " PFI.cPersCIIU, CI.cCIIUdescripcion, PFI.dPersNacCreac, PersAna.cPersNombre cAnalista, CN2.cConsDescripcion cDestino, "
'    ssql = ssql & " CN3.cConsDescripcion cCondicion, CE.dPrdEstado dFecAsig, Inst.cPersNombre cNomInst, CV.cCodModular, CE.nMonto nMontoSol,  "
'    ssql = ssql & " CN4.cConsDescripcion cMonedaSol, CE.nCuotas nCuotasSol, CE.nPlazo nPlazoSol, CN5.cConsDescripcion cTipoCred, CN6.cConsDescripcion cPersoneria, PF.cPersCod cFteInd"
'    ssql = ssql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtacod AND PP.nPrdPersRelac = " & gColRelPersTitular
'    ssql = ssql & "                 Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
'    ssql = ssql & "                 Left Join PersID PID ON PID.cPersCod = PP.cPersCod AND PID.cPersIDTpo = " & gPersIdDNI
'    ssql = ssql & "                 Left Join PersID PID2 ON PID2.cPersCod = PP.cPersCod AND PID2.cPersIDTpo = " & gPersIdRUC
'    ssql = ssql & "                 Left Join PersonaNat PNat ON PNat.cPersCod = PP.cPersCod"
'    ssql = ssql & "                 Left Join Constante CN ON CN.nConsValor = PNat.nPersNatEstCiv AND CN.nConsCod = " & gPersEstadoCivil
'    ssql = ssql & "                 Left Join ColocFteIngreso CF ON CF.cCtaCod = P.cCtaCod "
'    ssql = ssql & "                 Left Join PersFteIngreso PF ON PF.cNumFuente = CF.cNumFuente "
'    ssql = ssql & "                 Left Join Persona PFI ON PFI.cPersCod = PF.cPersCod"
'    ssql = ssql & "                 Left Join CIIU CI ON PFI.cPersCIIU = CI.cCIIUcod "
'    ssql = ssql & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAnalista
'    ssql = ssql & "                 Inner Join Persona PersAna ON PersAna.cPersCod = PP2.cPersCod "
'    ssql = ssql & "                 Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
'    ssql = ssql & "                 Left Join Constante CN2 ON CN2.nConsValor = CC.nColocDestino AND CN2.nConsCod = " & gColocDestino
'    ssql = ssql & "                 Left Join Constante CN3 ON CN3.nConsValor = CC.nColocCondicion AND CN3.nConsCod = " & gColocCredCondicion
'    ssql = ssql & "                 Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic
'    ssql = ssql & "                 Left Join ColocacConvenio CV ON CV.cCtaCod = P.cCtaCod "
'    ssql = ssql & "                 Left Join Persona Inst ON Inst.cPersCod = CV.cPersCod "
'    ssql = ssql & "                 Left Join Constante CN4 ON Convert(int, SubString(P.cCtaCod,9,1)) = CN4.nConsValor AND CN4.nConsCod = " & gMoneda
'    ssql = ssql & "                 Left Join Constante CN5 ON Convert(int, SubString(P.cCtaCod,6,3)) = CN5.nConsValor AND CN5.nConsCod = " & gProducto
'    ssql = ssql & "                 Left Join Constante CN6 ON CN6.nConsValor = Pers.nPersPersoneria AND CN6.nConsCod = " & gPersPersoneria
'    ssql = ssql & "                 Left Join PersFIDependiente PFDep ON PFDep.cNumFuente = CF.cNumFuente AND PFDep.dPersEval = CF.dPersEval "
'    ssql = ssql & "                 Left Join PersFIIndependiente PFInd ON PFInd.cNumFuente = CF.cNumFuente AND PFInd.dPersEval = CF.dPersEval "
'    ssql = ssql & " Where P.cCtaCod = '" & psCtaCod & "'"

'***********COMENTADO POR APRI20180215 ERS036-2017

'sSql = "Select Pers.cPersCod, Pers.cPersNombre, PID.cPersIDNro DNI,PID2.cPersIDNro RUC,"
'sSql = sSql & " Pers.cPersDireccDomicilio, PNat.nPersnatHijos, Nacimiento=Pers.dPersNacCreac,"
'sSql = sSql & " PNat.nPersNatEstCiv,  CN.cConsDescripcion cEstadoCivil,"
'sSql = sSql & " PFI.cPersNombre cRazonSocial, PFI.cPersDireccDomicilio cDireccRazon,"
'sSql = sSql & " PFI.cPersCIIU, CI.cCIIUdescripcion, PFI.dPersNacCreac, PersAna.cPersNombre cAnalista,"
'sSql = sSql & " CN2.cConsDescripcion cDestino,  CN3.cConsDescripcion cCondicion, CE.dPrdEstado dFecAsig,"
'sSql = sSql & " Inst.cPersNombre cNomInst, CV.cCodModular, CE.nMonto nMontoSol,"
'sSql = sSql & " CN4.cConsDescripcion cMonedaSol, CE.nCuotas nCuotasSol, CE.nPlazo nPlazoSol,"
'sSql = sSql & " CN5.cConsDescripcion cTipoCred, CN6.cConsDescripcion cPersoneria, PF.cPersCod cFteInd ,"
'sSql = sSql & " PF.nPersTipFte as nTipoFuente"
'sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtacod AND PP.nPrdPersRelac = 20"
'sSql = sSql & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
'sSql = sSql & " Left Join PersID PID ON PID.cPersCod = PP.cPersCod AND PID.cPersIDTpo = 1"
'sSql = sSql & " Left Join PersID PID2 ON PID2.cPersCod = PP.cPersCod AND PID2.cPersIDTpo = 2"
'sSql = sSql & " Left Join PersonaNat PNat ON PNat.cPersCod = PP.cPersCod"
'sSql = sSql & " Left Join Constante CN ON CN.nConsValor = PNat.nPersNatEstCiv AND CN.nConsCod = 1020"
'sSql = sSql & " Left Join ColocFteIngreso CF ON CF.cCtaCod = P.cCtaCod"
'sSql = sSql & " Left Join PersFteIngreso PF ON PF.cNumFuente = CF.cNumFuente"
'sSql = sSql & " Left Join Persona PFI ON PFI.cPersCod = PF.cPersFIPersCod"
'sSql = sSql & " Left Join CIIU CI ON PFI.cPersCIIU = CI.cCIIUcod"
'sSql = sSql & " Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = 28"
'sSql = sSql & " Inner Join Persona PersAna ON PersAna.cPersCod = PP2.cPersCod"
'sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod"
'sSql = sSql & " Left Join Constante CN2 ON CN2.nConsValor = CC.nColocDestino AND CN2.nConsCod = 3016"
'sSql = sSql & " Left Join Constante CN3 ON CN3.nConsValor = CC.nColocCondicion AND CN3.nConsCod = 3015"
'sSql = sSql & " Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado = 2000"
'sSql = sSql & " Left Join ColocacConvenio CV ON CV.cCtaCod = P.cCtaCod"
'sSql = sSql & " Left Join Persona Inst ON Inst.cPersCod = CV.cPersCod"
'sSql = sSql & " Left Join Constante CN4 ON Convert(int, SubString(P.cCtaCod,9,1)) = CN4.nConsValor AND CN4.nConsCod = 1011"
'sSql = sSql & " Left Join Constante CN5 ON Convert(int, SubString(P.cCtaCod,6,3)) = CN5.nConsValor AND CN5.nConsCod = 1001"
'sSql = sSql & " Left Join Constante CN6 ON CN6.nConsValor = Pers.nPersPersoneria AND CN6.nConsCod = 1002"
'sSql = sSql & " Left Join PersFIDependiente PFDep ON PFDep.cNumFuente = CF.cNumFuente AND PFDep.dPersEval = CF.dPersEval"
'sSql = sSql & " Left Join PersFIIndependiente PFInd ON PFInd.cNumFuente = CF.cNumFuente AND PFInd.dPersEval = CF.dPersEval"
'sSql = sSql & " Where P.cCtaCod = '" & psCtaCod & "'"
    
    sSql = "Exec stp_sel_RecuperaDatosSolicitudCredito '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosSolicitudCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaDatosResumenComite(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "Select Pers.cPersCod, Pers.cPersNombre, PID.cPersIDNro DNI,PID2.cPersIDNro RUC, "
    sSql = sSql & " Pers.cPersDireccDomicilio, PNat.nPersnatHijos, Pers.dPersNacCreac, Pers.nperspersoneria, PNat.nPersNatEstCiv, "
    sSql = sSql & " CN.cConsDescripcion cEstadoCivil, CN2.cConsDescripcion cDestino, "
    sSql = sSql & " CN3.cConsDescripcion cCondicion, CN5.cConsDescripcion cTipoCred, L.cDescripcion cLinea, "
    sSql = sSql & " CN6.cConsDescripcion cPersoneria "
    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtacod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                 Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSql = sSql & "                 Left Join PersID PID ON PID.cPersCod = PP.cPersCod AND PID.cPersIDTpo = " & gPersIdDNI
    sSql = sSql & "                 Left Join PersID PID2 ON PID2.cPersCod = PP.cPersCod AND PID2.cPersIDTpo = " & gPersIdRUC
    sSql = sSql & "                 Left Join PersonaNat PNat ON PNat.cPersCod = PP.cPersCod"
    sSql = sSql & "                 Left Join Constante CN ON CN.nConsValor = PNat.nPersNatEstCiv AND CN.nConsCod = " & gPersEstadoCivil
    sSql = sSql & "                 Inner Join Colocaciones C ON C.cCtaCod = P.cCtaCod "
    sSql = sSql & "                 Left Join ColocLineaCredito L ON C.cLineaCred = L.cLineaCred "
    sSql = sSql & "                 Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    sSql = sSql & "                 Left Join Constante CN2 ON CN2.nConsValor = CC.nColocDestino AND CN2.nConsCod = " & gColocDestino
    sSql = sSql & "                 Left Join Constante CN3 ON CN3.nConsValor = CC.nColocCondicion AND CN3.nConsCod = " & gColocCredCondicion
    sSql = sSql & "                 Left Join Constante CN5 ON Convert(int, SubString(P.cCtaCod,6,3)) = CN5.nConsValor AND CN5.nConsCod = " & gProducto
    sSql = sSql & "                 Left Join Constante CN6 ON CN6.nConsValor = Pers.nPersPersoneria AND CN6.nConsCod = " & gPersPersoneria
    sSql = sSql & " Where P.cCtaCod = '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosResumenComite = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

' CMACICA_CSTS - 08/11/2003 VERIFICAR SI ESTA FUNCION DEBERIA IR EN ESTA CLASE O EN ALGUNA CLASE DE PERSONAS -----------

Public Function ObtieneNombrePersonaUser(ByVal psCodUser As String) As String
Dim SQL As String
Dim rs As ADODB.Recordset
Dim Co As COMConecta.DCOMConecta
    Set Co = New COMConecta.DCOMConecta
    SQL = "select cPersNombre from Persona where cPersCod = (select cPersCod from RRHH where cUser = '" & psCodUser & "')"
    Co.AbreConexion
    Set rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
    If rs.BOF Then
        ObtieneNombrePersonaUser = ""
    Else
        ObtieneNombrePersonaUser = rs!cPersNombre
    End If

    Set Co = Nothing
    Set rs = Nothing

End Function
' ----------------------------------------------------------------------------------------------------------------------

Public Function RecuperaDatosDocDesembolso(ByVal psCtaCod As String, Optional ByVal pnRefina As Integer = -1) As ADODB.Recordset
'LUCV20160621, Agregó: pnRefina, Según ERS004-2016
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    'LUCV20160621, Comentó -> Segun ERS004-2016 **********->
    'WIOR 20120815 modificao la columnda de CE.dPrdEstado dFecDesemb por CE2.dPrdEstado
'    sSQL = "Select PersAna.cPersNombre cAnalista, CE.dPrdEstado dFecApr, Pers.cPersCod, "
'    sSQL = sSQL & " Pers.cPersNombre,CE2.dPrdEstado dFecDesemb, PersApod.cPersNombre cApoderado, CN.cConsDescripcion cMoneda,"
'    sSQL = sSQL & " L.cDescripcion cLinea, P.nTasaInteres, CN2.cConsDescripcion cTipoDesemb, CN3.cConsDescripcion cTipoCred "
'    'ssql = ssql & " cPersNombreUser = ( select cPersNombre from Persona where cPersCod = (select cPersCod from RRHH where cUser = '" & gsUser & "'))"
'    sSQL = sSQL & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersAnalista
'    sSQL = sSQL & "                Inner Join Persona PersAna ON PersAna.cPersCod = PP.cPersCod "
'    sSQL = sSQL & "                Inner Join ColocacEstado CE ON P.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstAprob
'    sSQL = sSQL & "                Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersTitular
'    sSQL = sSQL & "                Inner Join Persona Pers ON Pers.cPersCod = PP2.cPersCod "
'    sSQL = sSQL & "                Inner Join ColocacEstado CE2 ON P.cCtaCod = CE2.cCtaCod AND CE2.nPrdEstado = " & gColocEstVigNorm
'    sSQL = sSQL & "                Inner Join ProductoPersona PP3 ON P.cCtaCod = PP3.cCtaCod AND PP3.nPrdPersRelac = " & gColRelPersApoderado
'    sSQL = sSQL & "                Inner Join Persona PersApod ON PersApod.cPersCod = PP3.cPersCod "
'    sSQL = sSQL & "                Inner Join Constante CN ON CN.nConsValor = Convert(Int, Substring(P.cCtaCod,9,1)) AND CN.nConsCod = " & gMoneda
'    sSQL = sSQL & "                Inner Join Colocaciones C ON C.cCtacod = P.cCtaCod "
'    sSQL = sSQL & "                Inner Join ColocLineaCredito L ON L.cLineaCred = C.cLineaCred "
'    sSQL = sSQL & "                Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
'    sSQL = sSQL & "                Inner Join Constante CN2 ON CN2.nConsValor = CC.nTipoDesembolso AND CN2.nConsCod = " & gColocTiposDesembolso
'    sSQL = sSQL & "                Left  Join Constante CN3 ON Convert(int, SubString(P.cCtaCod,6,3)) = CN3.nConsValor AND CN3.nConsCod = " & gProducto
'    sSQL = sSQL & " Where P.cCtaCod = '" & psCtaCod & "'"
    'LUCV20160621 Fin Comentario <-**********
        
    'LUCV20160621, Segun ERS004-2016 **********->
    sSql = "EXEC stp_sel_RecuperaDatosDocDesembolso '" & psCtaCod & "', " & pnRefina & " "
    'LUCV20160621 Fin <-**********
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosDocDesembolso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaDatosDocPlanPagos(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "Select Pers.cPersNombre, PersAna.cPersNombre cAnalista, CN.cConsDescripcion cTipoCuota, "
    sSql = sSql & " nMontoCuota = (Select SUM(CD.nMonto) from ColocCalendDet CD where nNrocalen=(Select nNroCalen From ColocacCred Where cCtaCod = CD.cCtaCod) and cCtaCod = P.cCtaCod AND nColocCalendApl = 1 AND nCuota = 1), "
    sSql = sSql & " P.nTasaInteres, CN2.cConsDescripcion cMoneda, CE.nPlazo, CE.nPeriodoGracia, CE.nMonto, CE2.dVigencia dFecVig, CE. nTipoGracia, CN3.cConsDescripcion cTipoGracia "
    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                INNER Join Colocaciones CE2 ON CE2.cCtaCod = P.cCtaCod "
    sSql = sSql & "                Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSql = sSql & "                Inner Join ProductoPersona PP2 ON PP2.cCtaCod = P.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "                Inner Join Persona PersAna ON PersAna.cPersCod = PP2.cPersCod"
    sSql = sSql & "                Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado = " & gColocEstAprob
    sSql = sSql & "                Inner Join Constante CN ON CN.nConsValor = CE.nColocCalendCod AND CN.nConsCod = " & gColocTipoCalend
    sSql = sSql & "                Inner Join Constante CN2 ON CN2.nConsValor = Convert(int,Substring(P.cCtaCod,9,1)) AND CN2.nConsCod = " & gMoneda
    sSql = sSql & "                Left Join Constante CN3 ON CN3.nConsValor = CE.nTipoGracia AND CN3.nConsCod = 3017"
    sSql = sSql & " Where P.cCtaCod = '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosDocPlanPagos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaDatosBoletaDesembolso(ByVal pnNroMov As Long, ByVal psCtaCod As String, ByVal psNomAge As String, Optional ByVal pnRefina As Integer = -1) As ADODB.Recordset ' LUCV20160621, Agregó: pnRefina, Segun ERS004-2016
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    'sSql = "Select P.cPersNombre,  CASE WHEN SUBSTRING(MC.cOpeCod,1,2) = 12 THEN SUM(MC.nMonto) "
    'sSql = sSql & "                WHEN MC.cOpeCod = '" & gColocConceptoCodCapital & "' THEN SUM(MC.nMonto) "
    'sSql = sSql & "                END as nMonTran, "
    'sSql = sSql & "    FechaTran=dbo.FechaMov(M.cMovNro), HoraTran=dbo.HoraMov(M.cMovNro), CN.cConsDescripcion as sMoneda,"
    'sSql = sSql & "       COUNT(MC.nMovNro) as nNumTran, MC.nNrocuota, RIGHT(M.cMovNro,4) as sUsuario, Cal.dVenc"
    'sSql = sSql & " From ProductoPersona PP Inner Join Persona P ON PP.cPersCod = P.cPersCod And PP.nPrdPersRelac = " & gColRelPersTitular
    'sSql = sSql & "       Inner Join MovColDet MC ON MC.cCtaCod = PP.cCtacod "
    'sSql = sSql & "       Inner Join Mov M ON M.nMovNro = MC.nMovnro "
    'sSql = sSql & "       Inner Join Constante CN ON CN.nConsCod = " & gMoneda & " And CN.nConsValor = " & Mid(psCtaCod, 9, 1)
    'sSql = sSql & "       Inner Join ColocCalendario Cal ON Cal.cCtaCod = PP.cCtaCod And nColocCalendApl = " & gColocCalendAplCuota & " And nCuota = 1"
    'sSql = sSql & " Where PP.cCtaCod = '" & psCtaCod & "' And MC.nMovNro = " & pnNroMov
    'sSql = sSql & " Group By P.cPersNombre,  CN.cConsDescripcion, MC.nNrocuota, RIGHT(M.cMovNro,4), Cal.dVenc, MC.cOpeCod,M.cMovNro"
    
    'sSql = " Select P.cPersNombre,CASE WHEN SUBSTRING(CONVERT(VARCHAR(6),MC.nPrdConceptoCod),1,2) = 12 THEN SUM(MC.nMonto) END as nGasto, CASE WHEN SUBSTRING(CONVERT(VARCHAR(6),MC.nPrdConceptoCod),1,2) = 10 THEN SUM(MC.nMonto) END as nCapital, "
    'sSql = sSql & "     FechaTran=dbo.FechaMov(M.cMovNro), HoraTran=dbo.HoraMov(M.cMovNro), CN.cConsDescripcion as sMoneda,       MC.nNrocuota, RIGHT(M.cMovNro,4) as sUsuario, Cal.dVenc,COUNT(MC.nMovNro) as nNumTran "
    'sSql = sSql & " From ProductoPersona PP Inner Join Persona P ON PP.cPersCod = P.cPersCod And PP.nPrdPersRelac = 20 "
    'sSql = sSql & "       Inner Join MovColDet MC ON MC.cCtaCod = PP.cCtacod "
    'sSql = sSql & "       Inner Join Mov M ON M.nMovNro = MC.nMovnro "
    'sSql = sSql & "       Inner Join Constante CN ON CN.nConsCod = " & gMoneda & " And CN.nConsValor = " & Mid(psCtaCod, 9, 1)
    'sSql = sSql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = PP.cCtaCod And nColocCalendApl = 1 And nCuota = 1 and Cal.nNroCalen = (Select nNroCalen from ColocacCred where cCtaCod = PP.cCtaCod) "
    'sSql = sSql & " Where PP.cCtaCod = '" & psCtaCod & "' And MC.nMovNro = " & pnNroMov & " and MC.cOpeCod in ('100101','100102','100103','100104','100105','100106','100107','100109')"
    'sSql = sSql & " Group By P.cPersNombre,  CN.cConsDescripcion, MC.nNrocuota, RIGHT(M.cMovNro,4), Cal.dVenc, MC.cOpeCod,M.cMovNro,SUBSTRING(CONVERT(VARCHAR(6),MC.nPrdConceptoCod),1,2)"
    
    '18-04-2006
    'LUCV20160621, Segun ERS004-2016 **********->
'    sSQL = " Select P.cPersNombre,CASE WHEN SUBSTRING(CONVERT(VARCHAR(6),MC.nPrdConceptoCod),1,2) = 12 THEN SUM(MC.nMonto) END as nGasto, CASE WHEN SUBSTRING(CONVERT(VARCHAR(6),MC.nPrdConceptoCod),1,2) = 10 THEN SUM(MC.nMonto) END as nCapital, "
'    sSQL = sSQL & "     FechaTran=dbo.FechaMov(M.cMovNro), HoraTran=dbo.HoraMov(M.cMovNro), CN.cConsDescripcion as sMoneda,       MC.nNrocuota, RIGHT(M.cMovNro,4) as sUsuario, Cal.dVenc,COUNT(MC.nMovNro) as nNumTran ,cDNI=ISNULL(ID.cPersIDnro,'')"
'    sSQL = sSQL & " From ProductoPersona PP Inner Join Persona P ON PP.cPersCod = P.cPersCod And PP.nPrdPersRelac = 20 "
'    sSQL = sSQL & " LEFT JOIN PersID ID ON P.cPersCod = ID.cPersCod And ID.cPersIDTpo = 1 " 'ARCV 30-07-2006
'    sSQL = sSQL & "       Inner Join MovColDet MC ON MC.cCtaCod = PP.cCtacod "
'    sSQL = sSQL & "       Inner Join Mov M ON M.nMovNro = MC.nMovnro "
'    sSQL = sSQL & "       Inner Join Constante CN ON CN.nConsCod = " & gMoneda & " And CN.nConsValor = " & Mid(psCtaCod, 9, 1)
'    sSQL = sSQL & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = PP.cCtaCod And nColocCalendApl = 1 And nCuota = 1 and Cal.nNroCalen = (Select nNroCalen from ColocacCred where cCtaCod = PP.cCtaCod) "
'    sSQL = sSQL & " Where PP.cCtaCod = '" & psCtaCod & "' And MC.nMovNro = " & pnNroMov & " and MC.cOpeCod in ('100101','100102','100103','100104','100105','100106','100107','100109','100110','100111')"
'    sSQL = sSQL & " Group By P.cPersNombre,  CN.cConsDescripcion, MC.nNrocuota, RIGHT(M.cMovNro,4), Cal.dVenc, MC.cOpeCod,M.cMovNro,SUBSTRING(CONVERT(VARCHAR(6),MC.nPrdConceptoCod),1,2) ,ID.cPersIDnro"
    'LUCV20160621 Fin <-**********
    
    'LUCV20160621, Segun ERS004-2016 **********->
    sSql = "EXEC stp_sel_RecuperaDatosBoletaDesembolso " & pnNroMov & ",'" & psCtaCod & "', " & pnRefina & ""
    'LUCV20160621 Fin <-**********
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosBoletaDesembolso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function


Public Function RecuperaDatosBoletaGastosDesembolso(ByVal pnNroMov As Long, ByVal psCtaCod As String, ByVal psNomAge As String) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    sSql = "Select P.cPersNombre,  MC.nMonto, G.cDescripcion, "
    sSql = sSql & "       FechaTran=dbo.FechaMov(M.cMovnro), HoraTran=dbo.HoraMov(M.cMovnro), CN.cConsDescripcion as sMoneda, "
    sSql = sSql & "       MC.nNrocuota, RIGHT(M.cMovNro,4) as sUsuario, Cal.dVenc"
    sSql = sSql & " From ProductoPersona PP Inner Join Persona P ON PP.cPersCod = P.cPersCod And PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "       Inner Join MovColDet MC ON MC.cCtaCod = PP.cCtacod "
    sSql = sSql & "       Inner Join Mov M ON M.nMovNro = MC.nMovnro "
    sSql = sSql & "       Inner Join Constante CN ON CN.nConsCod = " & gMoneda & " And CN.nConsValor = " & Mid(psCtaCod, 9, 1)
    sSql = sSql & "       Inner Join ColocCalendario Cal ON Cal.cCtaCod = PP.cCtaCod And nColocCalendApl = " & gColocCalendAplDesembolso & " And nCuota = 1 and Cal.nNroCalen = (Select MAX(nNroCalen) from Coloccalendario where cCtaCod = PP.cCtaCod)"
    sSql = sSql & "       Inner Join ProductoConcepto G ON G.nPrdConceptoCod = MC.nPrdConceptoCod "
    sSql = sSql & " Where PP.cCtaCod = '" & psCtaCod & "' And MC.nMovNro = " & pnNroMov & " And SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2) = '12'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosBoletaGastosDesembolso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaDatosBoletaCancelCredDesembolso(ByVal pnNroMov As Long, ByVal psCtaCod As String, ByVal psNomAge As String) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

Dim sSql As String
    
    'WIOR 20150915 *** COMENTO PARA CAMBIAR A PROCEDIMIENTO ALMACENADO
    'sSql = "Select MCol.cCtacod as cCtacod,   "
    'sSql = sSql & " P.cPersNombre,"
    'sSql = sSql & " nCapital=(Select SUM(MC.nMonto) From MovColdet MC Where MC.cCtaCod = MCol.cCtaCod AND MC.nMovNro = MCol.nMovNro AND MC.nNroCalen = MCol.nNroCalen AND SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2)= '10'),"
    'sSql = sSql & " ninteres=(Select SUM(MC.nMonto) From MovColdet MC Where MC.cCtaCod = MCol.cCtaCod AND MC.nMovNro = MCol.nMovNro AND MC.nNroCalen = MCol.nNroCalen AND SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2)= '11'),"
    'sSql = sSql & " nGastos=(Select SUM(MC.nMonto) From MovColdet MC Where MC.cCtaCod = MCol.cCtaCod AND MC.nMovNro = MCol.nMovNro AND MC.nNroCalen = MCol.nNroCalen AND SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2)= '12'),"
    'sSql = sSql & " nItf =(Select SUM(MC.nMonto) From MovColdet MC Where MC.cCtaCod = MCol.cCtaCod AND MC.nMovNro = MCol.nMovNro AND MC.nPrdConceptoCod = 20 ), FechaTran=dbo.FechaMov(M.cMovNro), "
    'sSql = sSql & " FechaTran=dbo.FechaMov(M.cMovNro), HoraTran=dbo.HoraMov(M.cMovNro),"
    'sSql = sSql & " CN.cConsDescripcion as sMoneda,"
    'sSql = sSql & " COUNT(MCol.nMovNro) as nNumTran,"
    'sSql = sSql & " RIGHT(m.cMovNro, 4) As sUsuario"
    'sSql = sSql & " From"
    'sSql = sSql & " MovCol MCol"
    'sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = MCol.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    'sSql = sSql & " Inner Join Persona P ON P.cPersCod = PP.cPersCod"
    'sSql = sSql & " Inner Join Mov M ON M.nMovNro = MCol.nMovnro"
    'sSql = sSql & " Inner Join Constante CN ON CN.nConsCod = 1011 And CN.nConsValor = SUBSTRING(MCol.cCtaCod,9,1)"
    'sSql = sSql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = PP.cCtaCod And nColocCalendApl = 1  and Cal.nNrocalen = (select MAX(nNrocalen) From ColocCalendario Where cCtaCod = PP.cCtaCod)"
    'sSql = sSql & " Where MCol.nMovnro = " & pnNroMov & " AND MCol.cCtacod <> '" & psCtaCod & "' and MCOL.cOpeCod like '100[234567]%' "
    'sSql = sSql & " Group By MCol.cCtaCod, P.cPersNombre,M.cMovNro,CN.cConsDescripcion, MCol.nMovnro,MCol.nNroCalen "
    
    'WIOR 20150915 ***
    sSql = "EXEC stp_sel_DatosCancelCred " & pnNroMov & ",'" & psCtaCod & "'"
    'WIOR FIN ********
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosBoletaCancelCredDesembolso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function


Public Function ReportePagoDeCreditos(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda, ByVal pnRFA As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String

    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias)
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    'sSql = " SELECT  "
    'sSql = sSql & " CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'"
    'sSql = sSql & "        ELSE ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')"
    'sSql = sSql & "   END AS sProducto,"
    'sSql = sSql & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
    'sSql = sSql & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
    'sSql = sSql & "   END AS sFuenteFinan,"
    'sSql = sSql & "   CASE WHEN (GROUPING(O.cOpeDesc) = 1) THEN 'TOTAL '"
    'sSql = sSql & "        ELSE ISNULL(O.cOpeDesc, 'OPERACION DESCONOCIDA')"
    'sSql = sSql & "   END AS cOpeCod,"
    'sSql = sSql & "   CASE WHEN (GROUPING(MC.cCtaCod) = 1) THEN 'TOTAL'"
    'sSql = sSql & "        ELSE ISNULL(MC.cCtaCod, 'DESCONOCIDO')"
    'sSql = sSql & "   END AS cCtaCod, "
    'sSql = sSql & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = MC.cCtaCod), "
    'sSql = sSql & " sum(isnull(MCD.nMonto,0)) as nCapital,  sum(isnull(MCD2.nMonto,0)) as nInteres,"
    'sSql = sSql & " sum(isnull(MCD4.nMonto,0)) as nInteresMora, sum(isnull(MCD3.nMonto,0)) as nInteresGracia,"
    'sSql = sSql & " sum(isnull(MCD5.nMonto,0)) as nInteresReprog, sum(isnull(MCD6.nMonto,0)) as nInteresSuspenso,"
    'sSql = sSql & " sum(IsNull(MCD7.nMonto, 0)) As nGastos "
    'sSql = sSql & " FROM MovCol MC Inner Join Constante CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = 1001"
    'sSql = sSql & " Inner Join OpeTpo O ON O.cOpeCod = MC.cOpeCod"
    'sSql = sSql & " Inner Join Colocaciones C ON C.cCtaCod = MC.cCtaCod"
    'sSql = sSql & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = C.cLineaCred"
    'sSql = sSql & " Inner Join Persona P ON CL.cPersCod = P.cPersCod"
    'sSql = sSql & " Left Join MovColdet MCD ON MCD.nMovNro = MC.nMovNro And MCD.cOpeCod = MC.cOpeCod And MCD.cCtaCod = MC.cCtaCod And MCD.nNroCalen = MC.nNroCalen And MCD.nPrdConceptoCod = 1000"
    'sSql = sSql & " Left Join MovColdet MCD2 ON MCD2.nMovNro = MC.nMovNro And MCD2.cOpeCod = MC.cOpeCod And MCD2.cCtaCod = MC.cCtaCod And MCD2.nNroCalen = MC.nNroCalen And MCD2.nPrdConceptoCod = 1100"
    'sSql = sSql & " Left Join MovColdet MCD3 ON MCD3.nMovNro = MC.nMovNro And MCD3.cOpeCod = MC.cOpeCod And MCD3.cCtaCod = MC.cCtaCod And MCD3.nNroCalen = MC.nNroCalen And MCD3.nPrdConceptoCod = 1102"
    'sSql = sSql & " Left Join MovColdet MCD4 ON MCD4.nMovNro = MC.nMovNro And MCD4.cOpeCod = MC.cOpeCod And MCD4.cCtaCod = MC.cCtaCod And MCD4.nNroCalen = MC.nNroCalen And MCD4.nPrdConceptoCod = 1101"
    'sSql = sSql & " Left Join MovColdet MCD5 ON MCD5.nMovNro = MC.nMovNro And MCD5.cOpeCod = MC.cOpeCod And MCD5.cCtaCod = MC.cCtaCod And MCD5.nNroCalen = MC.nNroCalen And MCD5.nPrdConceptoCod = 1103"
    'sSql = sSql & " Left Join MovColdet MCD6 ON MCD6.nMovNro = MC.nMovNro And MCD6.cOpeCod = MC.cOpeCod And MCD6.cCtaCod = MC.cCtaCod And MCD6.nNroCalen = MC.nNroCalen And MCD6.nPrdConceptoCod = 1104"
    'sSql = sSql & " Left Join ( Select nMovNro, cOpeCod, cCtaCod, nNroCalen, Convert(int, '1200') as nPrdConceptoCod, SUM(nMonto) as nMonto From MovColDet"
    'sSql = sSql & "       Where SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2) = '12'"
    'sSql = sSql & "       Group By nMovNro, cOpeCod, cCtaCod, nNroCalen, SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2))  MCD7"
    'sSql = sSql & "    ON MCD7.nMovNro = MC.nMovNro And MCD7.cOpeCod = MC.cOpeCod And MCD7.cCtaCod = MC.cCtaCod And MCD7.nNroCalen = MC.nNroCalen And MCD7.nPrdConceptoCod = 1200"
    'sSql = sSql & " Inner Join Mov M ON M.nMovNro = MC.nMovNro "
    'sSql = sSql & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'  and Len(MC.cCtaCod) = 18 And (O.cOpeCod >= '100200' and O.cOpeCod <= '100707') And SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And SUBSTRING(MC.cCtaCod,4,2) IN " & sCadAge
    'sSql = sSql & "     And SUBSTRING(M.cMovNro,1,8) Between  '" & Format(pdFecIni, "yyyymmdd") & "' AND '" & Format(pdFecFin, "yyyymmdd") & "' "
    'sSql = sSql & " GROUP BY CN.cConsDescripcion, P.cPersNombre, O.cOpeDesc, MC.cCtaCod"
    'sSql = sSql & " With ROLLUP "
'**************

    '---------
    sSql = " Select * From ("
    sSql = sSql & " SELECT"
    sSql = sSql & " CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'  ELSE  ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')   END AS sProducto,"
    sSql = sSql & "        CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')   END AS sFuenteFinan,"
    sSql = sSql & "        CASE WHEN (GROUPING(O.cOpeDesc) = 1) THEN 'TOTAL '        ELSE ISNULL(O.cOpeDesc, 'OPERACION DESCONOCIDA')   END AS cOpeCod,"
    sSql = sSql & "        CASE WHEN (GROUPING(MC.cCtaCod) = 1) THEN 'TOTAL'        ELSE ISNULL(MC.cCtaCod, 'DESCONOCIDO')   END AS cCtaCod,"
    sSql = sSql & " M.nMovNro,"
    sSql = sSql & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = MC.cCtaCod),"
    sSql = sSql & " sum(isnull(MCD.nMonto, 0)) as  nCapital,"
    sSql = sSql & " sum(isnull(MCD2.nMonto, 0)) as nInteres,"
    sSql = sSql & " sum(isnull(MCD4.nMonto, 0)) as nInteresMora,"
    sSql = sSql & " sum(isnull(MCD3.nMonto, 0)) as nInteresGracia,"
    sSql = sSql & " sum(isnull(MCD5.nMonto, 0)) as nInteresReprog,"
    sSql = sSql & " sum(isnull(MCD6.nMonto, 0)) as nInteresSuspenso,"
    sSql = sSql & " sum(IsNull(MCD7.nMonto, 0)) As nGastos"
    sSql = sSql & " FROM MovCol MC"
    sSql = sSql & " Inner Join Constante CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = 1001"
    sSql = sSql & " Inner Join OpeTpo O ON O.cOpeCod = MC.cOpeCod"
    sSql = sSql & " Inner Join Colocaciones C ON C.cCtaCod = MC.cCtaCod"
    sSql = sSql & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = C.cLineaCred"
    sSql = sSql & " Inner Join Persona P ON CL.cPersCod = P.cPersCod"
    sSql = sSql & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, nPrdConceptoCod, SUM(nMonto) as nMonto"
    sSql = sSql & "         From MovColDet Where nPrdConceptoCod IN(1000,1109) "
    sSql = sSql & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD ON MCD.nMovNro = MC.nMovNro And MCD.cOpeCod = MC.cOpeCod And MCD.cCtaCod = MC.cCtaCod And MCD.nNroCalen = MC.nNroCalen And MCD.nPrdConceptoCod IN (1000,1109)"
    sSql = sSql & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, 1100 as nPrdConceptoCod, SUM(nMonto) as nMonto"
    sSql = sSql & "         From MovColDet Where nPrdConceptoCod in (1100,1105) "
    sSql = sSql & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen) MCD2 ON MCD2.nMovNro = MC.nMovNro And MCD2.cOpeCod = MC.cOpeCod And MCD2.cCtaCod = MC.cCtaCod And MCD2.nNroCalen = MC.nNroCalen And MCD2.nPrdConceptoCod = 1100"
    sSql = sSql & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, nPrdConceptoCod, SUM(nMonto) as nMonto"
    sSql = sSql & "         From MovColDet Where nPrdConceptoCod = 1102"
    sSql = sSql & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD3 ON MCD3.nMovNro = MC.nMovNro And MCD3.cOpeCod = MC.cOpeCod And MCD3.cCtaCod = MC.cCtaCod And MCD3.nNroCalen = MC.nNroCalen And MCD3.nPrdConceptoCod = 1102"
    sSql = sSql & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, nPrdConceptoCod, SUM(nMonto) as nMonto"
    sSql = sSql & "         From MovColDet Where nPrdConceptoCod = 1101"
    sSql = sSql & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD4 ON MCD4.nMovNro = MC.nMovNro And MCD4.cOpeCod = MC.cOpeCod And MCD4.cCtaCod = MC.cCtaCod And MCD4.nNroCalen = MC.nNroCalen And MCD4.nPrdConceptoCod = 1101"
    sSql = sSql & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, nPrdConceptoCod, SUM(nMonto) as nMonto"
    sSql = sSql & "         From MovColDet Where nPrdConceptoCod = 1103"
    sSql = sSql & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD5 ON MCD5.nMovNro = MC.nMovNro And MCD5.cOpeCod = MC.cOpeCod And MCD5.cCtaCod = MC.cCtaCod And MCD5.nNroCalen = MC.nNroCalen And MCD5.nPrdConceptoCod = 1103"
    sSql = sSql & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, nPrdConceptoCod, SUM(nMonto) as nMonto"
    sSql = sSql & "         From MovColDet Where nPrdConceptoCod = 1104"
    sSql = sSql & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD6 ON MCD6.nMovNro = MC.nMovNro And MCD6.cOpeCod = MC.cOpeCod And MCD6.cCtaCod = MC.cCtaCod And MCD6.nNroCalen = MC.nNroCalen And MCD6.nPrdConceptoCod = 1104"
    sSql = sSql & " Left Join ( Select nMovNro, cOpeCod, cCtaCod, nNroCalen, Convert(int, '1200') as nPrdConceptoCod, SUM(nMonto) as nMonto"
    sSql = sSql & "         From MovColDet Where SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2) = '12'"
    sSql = sSql & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen, SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2))  MCD7  ON MCD7.nMovNro = MC.nMovNro And MCD7.cOpeCod = MC.cOpeCod And MCD7.cCtaCod = MC.cCtaCod And MCD7.nNroCalen = MC.nNroCalen And MCD7.nPrdConceptoCod = 1200"
    sSql = sSql & " Inner Join Mov M ON M.nMovNro = MC.nMovNro And M.nMovFlag = 0  "
    sSql = sSql & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'  and Len(MC.cCtaCod) = 18 And (O.cOpeCod >= '100200' and O.cOpeCod <= '100707')"
    sSql = sSql & " And SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And SUBSTRING(MC.cCtaCod,4,2) IN " & sCadAge
    'ssql = ssql & " And SUBSTRING(M.cMovNro,1,8) =  '" & Format(pdFecFin, "yyyymmdd") & "' "
    sSql = sSql & " And SUBSTRING(M.cMovNro,1,8) Between  '" & Format(pdFecIni, "yyyymmdd") & "' AND '" & Format(pdFecFin, "yyyymmdd") & "' "
    If pnRFA = 1 Then
        sSql = sSql & " AND MC.cCtaCod NOT IN ( "
        sSql = sSql & "     Select cCtaCod From ColocacCred Where cRFA = 'RFA' "
        sSql = sSql & "     ) "
    End If
    sSql = sSql & " GROUP BY CN.cConsDescripcion, P.cPersNombre, O.cOpeDesc, MC.cCtaCod,M.nMovNro With ROLLUP"
    sSql = sSql & " ) as T"
    sSql = sSql & " Where T.nMovNro is not NULL or  T.cCtaCod = 'TOTAL'"
    sSql = sSql & " Order by T.sProducto, T.cCtaCod, T.sTitular "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ReportePagoDeCreditos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosDesembolsados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda, ByVal pMatProducts As Variant) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadProd As String
    

    sCadProd = " And SUBSTRING(MC.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProducts) - 1
        sCadProd = sCadProd & pMatProducts(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProducts) = 0 Then
        sCadProd = ""
    End If
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    sSql = sSql & " SELECT * FROM ("
    sSql = sSql & "SELECT "
    sSql = sSql & "  CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSql = sSql & "   END AS sAgencia,"
    sSql = sSql & "   CASE WHEN (GROUPING(CN.cAbrev) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(CN.cAbrev, 'PRODUCTO DESCONOCIDO ')"
    sSql = sSql & "   END AS sProducto,"
    sSql = sSql & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
    sSql = sSql & "   END AS sFuenteFinan,"
    sSql = sSql & "   CASE WHEN (GROUPING(O.cOpeDesc) = 1) THEN 'TOTAL '"
    sSql = sSql & "        ELSE ISNULL(O.cOpeDesc, 'OPERACION DESCONOCIDA')"
    sSql = sSql & "   END AS cOpeCod,"
    sSql = sSql & "   CASE WHEN (GROUPING(MC.cCtaCod) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(MC.cCtaCod, 'CUENTA DESCONOCIDO ')"
    sSql = sSql & "   END AS nCredito,"
    sSql = sSql & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = " & gColRelPersTitular & " Where PP.cCtaCod = MC.cCtaCod),"
    'sSql = sSql & " SUM(CE.nMonto) as nMontoApr,"
    sSql = sSql & " SUM(isnull(MCD8.nMonto,0)) as nDesembolso,"
    sSql = sSql & " SUM(CASE WHEN MCD.nNroCuota >1 THEN 0 ELSE  CAL.nMonto END) as nIntApr, "
    sSql = sSql & " SUM(CASE WHEN MCD.nNroCuota >1 THEN 0 ELSE  CE.nMonto END) as nMontoApr, "
    sSql = sSql & " SUM(MCD7.nMonto) as Gasto,"
    sSql = sSql & " sAnalista = (Select Top 1 cUser From RRHH Where cPersCod = (select cPersCod From ProductoPersona PP2 Where PP2.cCtaCod = MC.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersAnalista & ")),"
    sSql = sSql & " nTasa = (Select P.nTasaInteres From Producto P Where P.cCtacod = MC.cCtaCod),  "
    sSql = sSql & " cUserDes = (SELECT  RIGHT(CMOVNRO,4)  "
    sSql = sSql & "             FROM    MOV M1  JOIN MOVCOL MC1 ON MC1.NMOVNRO = M1.NMOVNRO "
    sSql = sSql & "             WHERE   (MC1.cOpeCod >= '100101' and MC1.cOpeCod <= '100105') AND MC1.cCtacod=MC.cCtaCod and M1.nMovFlag=0 And MC1.nMovNro=MC.nMovNro) "
    sSql = sSql & " FROM MovCol MC "
    sSql = sSql & " Inner Join ("
    sSql = sSql & "     Select nMovNro, cCtaCod, nNroCuota From MovColDet Where  cOpeCod in ('100101','100102','100103','100104','100105') Group by nMovNro, cCtaCod, nNroCuota"
    sSql = sSql & " ) MCD ON MCD.nMovNro = MC.nMovNro AND  MCD.cCtaCod = MC.cCtaCod"
    'sSql = sSql & " Inner Join Constante CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSql = sSql & " INNER JOIN (Select nConsCod,    nConsValor,  cConsDescripcion, (CONVERT(varchar(10),nConsValor)  + ' ' + cConsDescripcion ) as cAbrev "
    sSql = sSql & "    From Constante ) as CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSql = sSql & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(MC.cCtaCod,4,2)"
    sSql = sSql & " Inner Join OpeTpo O ON O.cOpeCod = MC.cOpeCod"
    sSql = sSql & " Inner Join Colocaciones CC ON CC.cCtaCod = MC.cCtaCod"
    sSql = sSql & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = CC.cLineaCred"
    sSql = sSql & " Inner Join Persona P ON CL.cPersCod = P.cPersCod "
    sSql = sSql & " Left Join ( Select nMovNro, cOpeCod, cCtaCod, nNroCalen, Convert(int, '1200') as nPrdConceptoCod, SUM(nMonto) as nMonto From MovColDet"
    sSql = sSql & "       Where SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2) = '12'"
    sSql = sSql & "       Group By nMovNro, cOpeCod, cCtaCod, nNroCalen, SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2))  MCD7"
    sSql = sSql & "    ON MCD7.nMovNro = MC.nMovNro And MCD7.cOpeCod = MC.cOpeCod And MCD7.cCtaCod = MC.cCtaCod And MCD7.nNroCalen = MC.nNroCalen And MCD7.nPrdConceptoCod = 1200 "
    sSql = sSql & " Left Join ( Select nMovNro, cOpeCod, cCtaCod, nNroCalen, Convert(int, '1000') as nPrdConceptoCod, SUM(nMonto) as nMonto From MovColDet"
    sSql = sSql & "       Where nPrdConceptoCod = 1000 AND (cOpeCod >= '" & gCredDesembEfec & "' and cOpeCod <= '" & gCredDesembCtaNuevaDOA & "') "
    sSql = sSql & "       Group By nMovNro, cOpeCod, cCtaCod, nNroCalen, SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2))  MCD8"
    sSql = sSql & "    ON MCD8.nMovNro = MC.nMovNro And MCD8.cOpeCod = MC.cOpeCod And MCD8.cCtaCod = MC.cCtaCod And MCD8.nNroCalen = MC.nNroCalen And MCD8.nPrdConceptoCod = 1000 "
    sSql = sSql & " Inner Join ColocacEstado CE ON CE.cCtaCod = MC.cCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSql = sSql & " Inner Join ColocacCred C ON C.cCtaCod = MC.cCtaCod"
    sSql = sSql & " Inner Join ( Select CD.cCtaCod, CD.nNrocalen, SUM(CD.nMonto) as nMonto From ColocCalendDet CD"
    sSql = sSql & "        Where CD.nColocCalendApl = 1 And CD.nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio
    sSql = sSql & "        GROUP BY CD.cCtaCod, CD.nNrocalen) CAL ON CAL.cCtaCod = MC.cCtaCod And CAL.nNroCalen = C.nNroCalen"
    sSql = sSql & " Inner Join Mov M ON M.nMovNro = MC.nMovNro "
    sSql = sSql & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'  and Len(MC.cCtaCod) = 18 And (MC.cOpeCod >= '" & gCredDesembEfec & "' and MC.cOpeCod <= '" & gCredDesembCtaNuevaDOA & "') And SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And SUBSTRING(MC.cCtaCod,4,2) IN " & sCadAge
    sSql = sSql & "     And SUBSTRING(M.cMovNro,1,8) Between  '" & Format(pdFecIni, "yyyymmdd") & "' AND '" & Format(pdFecFin, "yyyymmdd") & "' " & sCadProd
    sSql = sSql & " and  nMovFlag <> 2 "
    sSql = sSql & " GROUP BY A.cAgeDescripcion, CN.cAbrev, P.cPersNombre, O.cOpeDesc, MC.cCtaCod,MC.nMovNro "
    sSql = sSql & " With ROLLUP)X "
    sSql = sSql & " Where cUserDes Is Not Null"
    sSql = sSql & " ORDER BY sAgencia, sProducto, sFuenteFinan, cOpeCod, nCredito"
    'sSql = sSql & " ORDER BY  CN.nConsValor "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
     Set RecuperaCreditosDesembolsados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function


Public Function RecuperaSaldoCarteraVigente(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String
   
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    sSql = " SELECT"
    sSql = sSql & " CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSql = sSql & "   END AS sAgencia,"
    sSql = sSql & "   CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')"
    sSql = sSql & "   END AS sProducto,"
    sSql = sSql & "   CASE WHEN (GROUPING(CN2.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(CN2.cConsDescripcion, 'PLAZO DESCONOCIDO ')"
    sSql = sSql & "   END AS sPlazo,"
    sSql = sSql & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
    sSql = sSql & "   END AS sFuenteFinan,"
    sSql = sSql & " COUNT(*) as nCasos,"
    sSql = sSql & " SUM(CE.nMonto) as nMontoApr,"
    'sSql = sSql & " SUM(Cal.nMonto) as nDesembolsado,"
    sSql = sSql & " SUM(CC.nMontoCol) as nDesembolsado,"
    sSql = sSql & " SUM(MC.nSaldo) As nSaldoCapital"
    sSql = sSql & " FROM Producto MC Inner Join Constante CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSql = sSql & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(MC.cCtaCod,4,2)"
    sSql = sSql & " Inner Join Colocaciones CC ON CC.cCtaCod = MC.cCtaCod"
    sSql = sSql & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = CC.cLineaCred"
    sSql = sSql & " Inner Join Persona P ON CL.cPersCod = P.cPersCod"
    sSql = sSql & " Inner join ColocacEstado CE ON CE.cCtaCod = MC.cCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSql = sSql & " Inner join ColocacCred C ON C.cCtaCod = MC.cCtaCod"
    'sSql = sSql & " Inner Join ( "
    'sSql = sSql & "  Select cCtaCod, SUM(nMonto) as nMonto from MovColDet"
    'sSql = sSql & "  where cOpecod in ('100101','100102','100103','100104','100105')"
    'sSql = sSql & "  AND nPrdConceptoCod = 1000"
    'sSql = sSql & "  GROUP BY cCtaCod"
    'sSql = sSql & "  ) CAL ON CAL.cCtaCod = MC.cCtaCod "
    sSql = sSql & " Inner join Constante CN2 ON CN2.nConsValor = Convert(int,SUBSTRING(CC.cLineaCred,6,1)) And CN2.nConsCod = " & gColocLineaCredPlazo
    sSql = sSql & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'  and Len(MC.cCtaCod) = 18  And SUBSTRING(CC.cLineaCred,5,1) = '" & pnMoneda & "' And SUBSTRING(MC.cCtaCod,4,2) IN " & sCadAge
    sSql = sSql & " And MC.nPrdEstado in (2020,2021,2022,2030,2031,2032) "
    sSql = sSql & " GROUP BY A.cAgeDescripcion, CN.cConsDescripcion, CN2.cConsDescripcion, P.cPersNombre"
    sSql = sSql & " With ROLLUP"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSaldoCarteraVigente = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
Public Function RecuperaCreditosCanceladosTodos(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda, _
    ByVal psAnalistas As String) As ADODB.Recordset
Dim sSql As String
Dim sCadAge As String
Dim i As Integer
Dim oConecta As COMConecta.DCOMConecta

    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    sSql = " Select X.sAgencia,X.sPlazo,X.sProducto,X.sFuenteFinan,X.nCredito,"
    sSql = sSql & " Case When RC.cCtaCodAnt is Null Then 'NUEVO CREDITO' "
    sSql = sSql & " Else RC.cCtaCodAnt End  as CuentaAntigua,"
    sSql = sSql & " X.sTitular,X.nDesembolso,X.nCuotasApr,X.nPlazoApr,X.dVigencia,X.sAnalista"
    sSql = sSql & " From ("
    sSql = sSql & " SELECT "
    sSql = sSql & " CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSql = sSql & "   END AS sAgencia,"
    sSql = sSql & "  CASE WHEN (GROUPING(CN2.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(CN2.cConsDescripcion, 'PLAZO DESCONOCIDO ')"
    sSql = sSql & "   END AS sPlazo,"
    sSql = sSql & "   CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')"
    sSql = sSql & "   END AS sProducto,"
    sSql = sSql & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
    sSql = sSql & "   END AS sFuenteFinan,"
    sSql = sSql & "   CASE WHEN (GROUPING(Prd.cCtaCod) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDO ')"
    sSql = sSql & "   END AS nCredito,"
    sSql = sSql & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = " & gColRelPersTitular & " Where PP.cCtaCod = Prd.cCtaCod),"
    sSql = sSql & " SUM(MC.nMonto) as nDesembolso,"
    sSql = sSql & " nCuotasApr = (Select nCuotas From ColocacEstado Where cCtaCod = Prd.cCtaCod And nPrdEstado = " & gColocEstAprob & "),"
    sSql = sSql & " nPlazoApr = (Select nPlazo From ColocacEstado Where cCtaCod = Prd.cCtaCod And nPrdEstado = " & gColocEstAprob & "),"
    sSql = sSql & " dVigencia = (Select dVigencia From Colocaciones Where cCtaCod = Prd.cCtaCod),"
    sSql = sSql & " sAnalista = (Select Top 1 cUser From RRHH Where cPersCod = (select cPersCod From ProductoPersona PP2 Where PP2.cCtaCod = Prd.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersAnalista & "))"
    sSql = sSql & " FROM Producto Prd Inner Join Constante CN ON Convert(Int,SUBSTRING(Prd.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSql = sSql & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & " Inner Join Colocaciones CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = CC.cLineaCred"
    sSql = sSql & " Inner Join Persona P ON CL.cPersCod = P.cPersCod"
    sSql = sSql & " Inner join MovCol MC ON MC.cCtaCod = Prd.cCtaCod And (MC.cOpeCod >= '" & gCredDesembEfec & "' and MC.cOpeCod <= '" & gCredDesembCtaNuevaDOA & "')"
    sSql = sSql & " Inner join Constante CN2 ON CN2.nConsValor = Convert(int,SUBSTRING(CC.cLineaCred,6,1)) And CN2.nConsCod = " & gColocLineaCredPlazo
    sSql = sSql & " Where SUBSTRING(Prd.cCtaCod,6,3)<>'305'  and Len(Prd.cCtaCod) = 18 And Prd.nPrdEstado = " & gColocEstCancelado & " And SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And SUBSTRING(MC.cCtaCod,4,2) IN " & sCadAge
    sSql = sSql & "     And Prd.dPrdEstado Between  '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "' "
    sSql = sSql & " GROUP BY A.cAgeDescripcion, CN2.cConsDescripcion, CN.cConsDescripcion, P.cPersNombre, Prd.cCtaCod"
    sSql = sSql & " With ROLLUP)X"
    sSql = sSql & " Left Join RelCuentas RC on RC.cCtaCod=X.nCredito"
    sSql = sSql & " Where X.nCredito in (Select cCtaCod From ProductoPersona"
    sSql = sSql & " Where cPersCod in (" & psAnalistas & ")" & " And nPrdPersRelac=28)"
    

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosCanceladosTodos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaCreditosCanceladosConsaldo(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sSql = "SELECT "
    sSql = sSql & "  CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSql = sSql & "   END AS sAgencia,"
    sSql = sSql & "  CASE WHEN (GROUPING(CN2.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(CN2.cConsDescripcion, 'PLAZO DESCONOCIDO ')"
    sSql = sSql & "   END AS sPlazo,"
    sSql = sSql & "   CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')"
    sSql = sSql & "   END AS sProducto,"
    sSql = sSql & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
    sSql = sSql & "   END AS sFuenteFinan,"
    sSql = sSql & "   CASE WHEN (GROUPING(Prd.cCtaCod) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDO ')"
    sSql = sSql & "   END AS nCredito,"
    sSql = sSql & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = " & gColRelPersTitular & " Where PP.cCtaCod = Prd.cCtaCod),"
    sSql = sSql & " nSaldo=(Select Top 1 nMonto From MovCol Where cCtaCod = Prd.cCtaCod Order by nMovnro DESC),"
    sSql = sSql & " SUM(ISNULL(MCD.nMonto,0)) as nCapital,"
    sSql = sSql & " SUM(ISNULL(MCD2.nMonto,0)) as nInteres,"
    sSql = sSql & " SUM(ISNULL(MCD4.nMonto,0)) as nMora,"
    sSql = sSql & " SUM(Cal.nMonto) as nIntPerd,"
    sSql = sSql & " cLineaCred = (Select cLineaCred From Colocaciones Where cCtaCod = Prd.cCtaCod),"
    sSql = sSql & " sAnalista = (Select Top 1 cUser From RRHH Where cPersCod = (select cPersCod From ProductoPersona PP2 Where PP2.cCtaCod = Prd.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersAnalista & "))"
    sSql = sSql & " FROM    Producto Prd"
    sSql = sSql & " Inner Join Constante CN ON Convert(Int,SUBSTRING(Prd.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSql = sSql & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & " Inner Join Colocaciones CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = CC.cLineaCred"
    sSql = sSql & " Inner Join Persona P ON CL.cPersCod = P.cPersCod"
    sSql = sSql & " Inner join MovCol MC ON MC.cCtaCod = Prd.cCtaCod And nMovNro = (Select MAX(nMovNro) From MovCol Where cCtaCod = MC.cCtaCod)"
    sSql = sSql & " Left Join ( Select nMovNro, cOpeCod, cCtaCod, nNroCalen, SUM(nMonto) as nMonto"
    sSql = sSql & "             From MovColdet"
    sSql = sSql & "        Where nPrdConceptoCod = " & gColocConceptoCodCapital
    sSql = sSql & "        Group By nMovNro, cOpeCod, cCtaCod, nNroCalen"
    sSql = sSql & "         ) as MCD ON MCD.nMovNro = MC.nMovNro And MCD.cOpeCod = MC.cOpeCod And MCD.cCtaCod = MC.cCtaCod And MCD.nNroCalen = MC.nNroCalen"
    sSql = sSql & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, SUM(nMonto) as nMonto"
    sSql = sSql & "        From MovColdet"
    sSql = sSql & "        Where   nPrdConceptoCod in (" & gColocConceptoCodInteresCompensatorio & "," & gColocConceptoCodInteresGracia & gColocConceptoCodInteresReprogramado & gColocConceptoCodInteresSuspenso & ") "
    sSql = sSql & "        Group By nMovNro, cOpeCod, cCtaCod, nNroCalen"
    sSql = sSql & "         ) as MCD2 ON MCD2.nMovNro = MC.nMovNro And MCD2.cOpeCod = MC.cOpeCod And MCD2.cCtaCod = MC.cCtaCod And MCD2.nNroCalen = MC.nNroCalen"
    sSql = sSql & "    Left Join ( Select nMovNro, cOpeCod, cCtaCod, nNroCalen, SUM(nMonto) as nMonto"
    sSql = sSql & "        From MovColDet"
    sSql = sSql & "        Where nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio
    sSql = sSql & "        Group By nMovNro, cOpeCod, cCtaCod, nNroCalen"
    sSql = sSql & "         ) as MCD4 ON MCD4.nMovNro = MC.nMovNro And MCD4.cOpeCod = MC.cOpeCod And MCD4.cCtaCod = MC.cCtaCod And MCD4.nNroCalen = MC.nNroCalen"
    sSql = sSql & " Inner Join ColocacCred C ON C.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Inner Join ( Select CD.cCtaCod, CD.nNrocalen, SUM(CD.nMonto - ISNULL(CD.nMontoPagado,0)) as nMonto From ColocCalendDet CD"
    sSql = sSql & "        Where CD.nColocCalendApl = 1 And CD.nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio
    sSql = sSql & "        GROUP BY CD.cCtaCod, CD.nNrocalen) CAL ON CAL.cCtaCod = Prd.cCtaCod And CAL.nNroCalen = C.nNroCalen "
    sSql = sSql & " Inner join Constante CN2 ON CN2.nConsValor = Convert(int,SUBSTRING(CC.cLineaCred,6,1)) And CN2.nConsCod = " & gColocLineaCredPlazo
    sSql = sSql & " Where SUBSTRING(Prd.cCtaCod,6,3)<>'305'  and Len(Prd.cCtaCod) = 18 And Prd.nPrdEstado = " & gColocEstCancelado & " And SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And SUBSTRING(MC.cCtaCod,4,2) IN " & sCadAge
    sSql = sSql & "     And Prd.dPrdEstado Between  '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "' "
    sSql = sSql & " GROUP BY A.cAgeDescripcion, CN2.cConsDescripcion, CN.cConsDescripcion, P.cPersNombre, Prd.cCtaCod"
    sSql = sSql & " With ROLLUP "

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosCanceladosConsaldo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaResumenSaldosCarteraPorAnalista(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal P1I As Integer, ByVal P1F As Integer, _
    ByVal P2I As Integer, ByVal P2F As Integer, ByVal P3I As Integer, ByVal P3F As Integer, ByVal P4F As Integer, _
    ByVal pMatCondicion As Variant, ByVal pMatProd As Variant) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sAnalistas As String
Dim oGen As COMDConstSistema.DCOMGeneral

    Set oGen = New COMDConstSistema.DCOMGeneral
    sAnalistas = oGen.LeeConstSistema(gConstSistRHCargoCodAnalistas)
    Set oGen = Nothing

    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For i = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
    Next i
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"
    
    sSql = "Select  "
    sSql = sSql & " CASE WHEN GROUPING(A.cAgeDescripcion)=1 THEN 'TOTAL'"
    sSql = sSql & "   ELSE ISNULL(A.cAgeDescripcion,'AGENCIA DESCONOCIDA')"
    sSql = sSql & " END  as cAgencia,"
    sSql = sSql & " CASE WHEN GROUPING(Q.cUser)=1 THEN 'TOTAL'"
    sSql = sSql & "   ELSE ISNULL(Q.cUser,'AGENCIA DESCONOCIDA')"
    sSql = sSql & " END as cUser,"
    sSql = sSql & " SUM(Q.nCar1Nro) as nCar1Nro, SUM(Q.nCar1Saldo) as nCar1Saldo,"
    sSql = sSql & " SUM(Q.nCar2Nro) as nCar2Nro, SUM(Q.nCar2Saldo) as nCar2Saldo,"
    sSql = sSql & " SUM(Q.nCar3Nro) as nCar3Nro, SUM(Q.nCar3Saldo) as nCar3Saldo,"
    sSql = sSql & " SUM(Q.nCar4Nro) as nCar4Nro, SUM(Q.nCar4Saldo) as nCar4Saldo,"
    sSql = sSql & " SUM(Q.nCar5Nro) as nCar5Nro, SUM(Q.nCar5Saldo) as nCar5Saldo"
    sSql = sSql & " From ("
    sSql = sSql & " select cAgencia, T.cUser + ' ' +  T.cPersNombre as cUser, SUM(T.nCar1Nro) as nCar1Nro, SUM(T.nCar1Saldo) as nCar1Saldo,"
    sSql = sSql & "    SUM(T.nCar2Nro) as nCar2Nro, SUM(T.nCar2Saldo) as nCar2Saldo,"
    sSql = sSql & "    SUM(T.nCar3Nro) as nCar3Nro, SUM(T.nCar3Saldo) as nCar3Saldo,"
    sSql = sSql & "    SUM(T.nCar4Nro) as nCar4Nro, SUM(T.nCar4Saldo) as nCar4Saldo,"
    sSql = sSql & "    SUM(T.nCar5Nro) as nCar5Nro, SUM(T.nCar5Saldo) as nCar5Saldo"
    sSql = sSql & " From ("
    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, COUNT(*) as nCar1Nro, SUM(Prd.nSaldo) as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo "
    sSql = sSql & "    From RRHH RH "
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSql = sSql & "        Inner Join RHCargos RC ON RC.cpersCod = RH.cpersCod "
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "    Where RC.cRHCargoCod in (" & sAnalistas & ")"
    sSql = sSql & "        AND CC.nDiasAtraso >=" & P1I & " AND CC.nDiasAtraso <=" & P1F
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Union"
    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, COUNT(*) as nCar2Nro, SUM(Prd.nSaldo) as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo "
    sSql = sSql & "    From RRHH RH "
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod "
    sSql = sSql & "        Inner Join RHCargos RC ON RC.cpersCod = RH.cpersCod "
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "    Where RC.cRHCargoCod in (" & sAnalistas & ")"
    sSql = sSql & "        AND CC.nDiasAtraso >=" & P2I & " AND CC.nDiasAtraso <=" & P2F
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Union"
    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, COUNT(*) as nCar3Nro, SUM(Prd.nSaldo) as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo "
    sSql = sSql & "    From RRHH RH "
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSql = sSql & "        Inner Join RHCargos RC ON RC.cpersCod = RH.cpersCod "
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "    Where RC.cRHCargoCod in (" & sAnalistas & ")"
    sSql = sSql & "        AND CC.nDiasAtraso >=" & P3I & " AND CC.nDiasAtraso <=" & P3F
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Union"
    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, COUNT(*) as nCar4Nro, SUM(Prd.nSaldo) as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo "
    sSql = sSql & "    From RRHH RH"
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSql = sSql & "        Inner Join RHCargos RC ON RC.cpersCod = RH.cpersCod "
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "    Where RC.cRHCargoCod in (" & sAnalistas & ")"
    sSql = sSql & "        AND CC.nDiasAtraso > " & P4F
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Union"
    sSql = sSql & "    Select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo , COUNT(*) as nCar5Nro, SUM(Prd.nSaldo) as nCar5Saldo "
    sSql = sSql & "    From RRHH RH"
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSql = sSql & "        Inner Join RHCargos RC ON RC.cpersCod = RH.cpersCod "
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod " & sCadProd
    sSql = sSql & "    Where RC.cRHCargoCod in (" & sAnalistas & ")"
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "'"
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & " ) as T"
    sSql = sSql & " Group by  T.cUser + ' ' +  T.cPersNombre, cAgencia"
    sSql = sSql & " ) As Q Inner Join Agencias A ON A.cAgeCod = Q.cAgencia"
    sSql = sSql & " WHERE Q.cAgencia IN " & sCadAge
    sSql = sSql & " GROUP BY A.cAgeDescripcion, Q.cUser"
    sSql = sSql & " With ROLLUP"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaResumenSaldosCarteraPorAnalista = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaMoraInstitucional(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal P1I As Integer, ByVal P1F As Integer, _
    ByVal P2I As Integer, ByVal P2F As Integer, ByVal P3I As Integer, ByVal P3F As Integer, ByVal P4F As Integer, _
    ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, ByVal pbMoraAnt As Boolean) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
    
    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For i = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
    Next i
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"


    sSql = " Select "
    sSql = sSql & "    CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSql = sSql & "    END AS sAgencia,"
    sSql = sSql & "    CASE WHEN GROUPING(T.cRango)=1 THEN 'TOTAL'"
    sSql = sSql & " ELSE ISNULL(T.cRango,'RANGO DESCONOCIDO')"
    sSql = sSql & " END as cRango,"
    sSql = sSql & " CASE WHEN GROUPING(T.cCtaCod)=1 THEN 'TOTAL'"
    sSql = sSql & " ELSE ISNULL(T.cCtaCod,'CUENTA DESCONOCIDA')"
    sSql = sSql & " END as cCtaCod,"
    sSql = sSql & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = T.cCtaCod),"
    sSql = sSql & " SUM(T.nMonto) as nMontoApr,"
    sSql = sSql & " SUM(T.nMontoCol) as nMontoDesemb,"
    sSql = sSql & " SUM(T.nSaldo) as nSaldoCap,"
    'sSql = sSql & " nCuotasApr = (Select nCuotas From ColocacEstado Where cCtaCod = T.cCtaCod And nPrdEstado = 2002),"
    sSql = sSql & " nCuotasApr = (Select COUNT(*) From ColocCalendario CC Where CC.cCtaCod = T.cCtaCod AND CC.nNroCalen = (Select nNroCalen from ColocacCred Where cCtaCod = T.cCtaCod) AND CC.nColocCalendApl = 1),"
    sSql = sSql & " nCuotaPend = (Select nNroProxCuota From ColocacCred Where cCtaCod = T.cCtaCod),"
    sSql = sSql & " SUM(T.nDiasAtraso) as nDiasAtraso,"
    sSql = sSql & " sAnalista = (Select Top 1 cUser From RRHH Where cPersCod = (select cPersCod From ProductoPersona PP2 Where PP2.cCtaCod = T.cCtaCod And PP2.nPrdPersRelac = 28)),"
    sSql = sSql & " nSaldoCapTotal = (Select SUM(nSaldo) From Producto Where nPrdEstado >= 2020 And nPrdEstado <= 2032 And SUBSTRING(cCtaCod,9,1) = '" & pnMoneda & "')"
    sSql = sSql & " From"
    sSql = sSql & " ( Select 'RANGO1' as cRango, Prd.cCtaCod, CE.nMonto, C.nMontoCol, Prd.nSaldo, CC.nDiasAtraso"
    sSql = sSql & " From  Producto Prd"
    sSql = sSql & "     Inner join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & "     Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "     Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        AND CC.nDiasAtraso >=" & P1I & " AND CC.nDiasAtraso <=" & P1F
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSql = sSql & " Union"
    sSql = sSql & " Select 'RANGO2' as cRango, Prd.cCtaCod, CE.nMonto, C.nMontoCol, Prd.nSaldo, CC.nDiasAtraso"
    sSql = sSql & " From  Producto Prd"
    sSql = sSql & "     Inner join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & "     Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "     Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        AND CC.nDiasAtraso >=" & P2I & " AND CC.nDiasAtraso <=" & P2F
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSql = sSql & " Union"
    sSql = sSql & " Select 'RANGO3' as cRango, Prd.cCtaCod, CE.nMonto, C.nMontoCol, Prd.nSaldo, CC.nDiasAtraso"
    sSql = sSql & " From  Producto Prd"
    sSql = sSql & "     Inner join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & "     Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "     Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        AND CC.nDiasAtraso >=" & P3I & " AND CC.nDiasAtraso <=" & P3F
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSql = sSql & " Union"
    sSql = sSql & " Select 'RANGO4' as cRango, Prd.cCtaCod, CE.nMonto, C.nMontoCol, Prd.nSaldo, CC.nDiasAtraso"
    sSql = sSql & " From  Producto Prd"
    sSql = sSql & "     Inner join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & "     Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "     Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        AND CC.nDiasAtraso >=" & P4F
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSql = sSql & " ) as T Inner join Agencias A ON A.cAgeCod = SUBSTRING(T.cCtaCod,4,2)"
    sSql = sSql & " WHERE A.cAgeCod IN " & sCadAge
    sSql = sSql & " Group by A.cAgeDescripcion, T.cRango, T.cCtaCod"
    sSql = sSql & " With ROLLUP "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaMoraInstitucional = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaMoraxAnalista_AtrasoPagoCuotaLibre(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, _
    ByVal pMatAnalistas As Variant, ByVal pnConCuotaLibre As Boolean) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sCadAnalista As String

    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"

    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If

    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For i = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
    Next i
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"

    sCadAnalista = "('"
    For i = 0 To UBound(pMatAnalistas) - 1
        sCadAnalista = sCadAnalista & pMatAnalistas(i) & "','"
    Next i
    sCadAnalista = Mid(sCadAnalista, 1, Len(sCadAnalista) - 2) & ")"

 sSql = " Select CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
 sSql = sSql & "       ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')    END AS cAgencia,"
 sSql = sSql & "   CASE WHEN GROUPING(P2.cPersNombre)=1 THEN 'TOTAL'"
 sSql = sSql & "       ELSE ISNULL(P2.cPersNombre,'PERSONA DESCONOCIDA') END as cAnalista,"
 sSql = sSql & "   CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL'"
 sSql = sSql & "       ELSE ISNULL(Prd.cCtaCod,'CUENTA DESCONOCIDA') END as cCtaCod,"
 sSql = sSql & "   (dbo.ColocCredTitular(Prd.cCtaCod) + '  ' + RTrim(dbo.ColocCredTitularDir(Prd.cCtaCod)) + '  '+ Cast(C.nMontoCol as varchar(20)) + '  ' + Cast(Prd.nSaldo as varchar(20)) +  '  '  + Cast(dbo.CuotaAprobada(Prd.cCtaCod) as varchar(15)) + '      ' + Cast(CCal.nCuota as varchar(10)) + '       ' +  CONVERT(varchar(15),CCal.dvenc,103)) as nDatos,"
 sSql = sSql & "   SUM(DATEDIFF(day,CCal.dVenc,'" & Format(pdFecIni, "mm/dd/yyyy") & "')) as nDiasAtraso,"
 sSql = sSql & "   SUM(ROUND(dbo.ColocMoraCuota(Prd.cCtaCod,CCal.nCuota),2)) as nMora,"
 sSql = sSql & "   Sum (ROUND(dbo.ColocGastoCuota(Prd.cCtaCod, CCal.nCuota),2)) as nGasto, "
 sSql = sSql & "   Sum (ROUND(dbo.ColocSaldoCuota(Prd.cCtaCod, CCal.nCuota),2)) as nSaldoCuota, 'A' AS cTipo "
 sSql = sSql & " From Producto Prd"
 sSql = sSql & "   Inner Join ProductoPersona PP ON PP.cCtaCod = Prd.cCtaCod And PP.nPrdPersRelac=28 And PP.cPersCod in " & sCadAnalista
 sSql = sSql & "   Inner Join Persona P2 ON P2.cPersCod = PP.cPersCod"
 sSql = sSql & "   Inner join Agencias A ON A.cAgeCod = SUBSTRING(PRD.cCtaCod,4,2)"
 sSql = sSql & "   Inner Join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
 sSql = sSql & "   Inner join ColocacCred CC ON CC.cCtaCod =  Prd.cCtaCod"
 sSql = sSql & "   Inner Join ColocCalendario CCal ON CCal.cCtaCod = Prd.cCtaCod And CCal.nNroCalen = CC.nNroCalen"
 sSql = sSql & "       And CCal.nColocCalendApl = 1 And CCal.dVenc < '" & Format(pdFecIni, "mm/dd/yyyy") & "' And CCal.nColocCalendEstado=0"
 sSql = sSql & " WHERE (Prd.nPrdEstado IN (" & gColocEstRefNorm & ", " & gColocEstRefVenc & ", " & gColocEstRefMor & ", " & gColocEstVigNorm & ", " & gColocEstVigVenc & ", " & gColocEstVigMor & ")) "

 sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd & " And A.cAgeCod IN " & sCadAge
 sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion

    If pnConCuotaLibre = True Then
        sSql = sSql & " And CC.nColocCalendCod=" & gColocCalendCodCL & " "
    End If

 sSql = sSql & "  Group by A.cAgeDescripcion,P2.cPersNombre,Prd.cCtaCod,"
 sSql = sSql & " (dbo.ColocCredTitular(Prd.cCtaCod) + '  ' + RTrim(dbo.ColocCredTitularDir(Prd.cCtaCod)) + '  '+ Cast(C.nMontoCol as varchar(20)) + '  ' + Cast(Prd.nSaldo as varchar(20)) +  '  '  + Cast(dbo.CuotaAprobada(Prd.cCtaCod) as varchar(15)) + '      ' + Cast(CCal.nCuota as varchar(10)) + '       ' +  CONVERT(varchar(15),CCal.dvenc,103))"
 sSql = sSql & " With ROLLUP"

    If pnConCuotaLibre = True Then
        'sSql = sSql & " CC.nColocCalendCod, "
    End If


    If pnConCuotaLibre = True Then
        'sSql = sSql & " And CC.nColocCalendCod=" & gColocCalendCodCL & " "
    End If

    If pnConCuotaLibre = True Then
        'sSql = sSql & ", CC.nColocCalendCod "
    End If

' CMACICA_CSTS 17/11/2003 ---------------------------------------------------------------------------------------------

 sSql = sSql & " UNION "
 sSql = sSql & " Select ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ') AS cAgencia,"
 sSql = sSql & " ISNULL(P3.cPersNombre,'PERSONA DESCONOCIDA') AS cAnalista,"
 sSql = sSql & " ISNULL(Prd.cCtaCod,'CUENTA DESCONOCIDA') as cCtaCod,"
 sSql = sSql & " LEFT(SUBSTRING(ISNULL(P2.cPersNombre,'AVAL DESCONOCIDO'),1,32) + SPACE(32), 32) + ' ' +  LEFT(SUBSTRING(ISNULL(P2.cPersDireccDomicilio,'NO EXISTE DIRECCION'),1,32)+ SPACE(32), 32) AS nDatos, "
 sSql = sSql & " 0 as nDiasAtraso, 0 as nMora, 0 as nGasto, 0 as nSaldoCuota, 'B' AS cTipo "
 sSql = sSql & " From Producto Prd"
 sSql = sSql & "   Inner Join ProductoPersona PP ON PP.cCtaCod = Prd.cCtaCod And PP.nPrdPersRelac=25"
 sSql = sSql & "   Inner Join ProductoPersona PP2 ON PP2.cCtaCod = Prd.cCtaCod And PP2.nPrdPersRelac=28 And PP2.cPersCod in " & sCadAnalista
 sSql = sSql & "   Inner Join Persona P2 ON P2.cPersCod = PP.cPersCod"
 sSql = sSql & "   Inner Join Persona P3 ON P3.cPersCod = PP2.cPersCod"
 sSql = sSql & "   Inner join Agencias A ON A.cAgeCod = SUBSTRING(PRD.cCtaCod,4,2)"
 sSql = sSql & "   Inner Join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
 sSql = sSql & "   Inner join ColocacCred CC ON CC.cCtaCod =  Prd.cCtaCod"
 sSql = sSql & "   Inner Join ColocCalendario CCal ON CCal.cCtaCod = Prd.cCtaCod And CCal.nNroCalen = CC.nNroCalen"
 sSql = sSql & "       And CCal.nColocCalendApl = 1 And CCal.dVenc < '" & Format(pdFecIni, "mm/dd/yyyy") & "' And CCal.nColocCalendEstado=0"

 sSql = sSql & " WHERE (Prd.nPrdEstado IN (" & gColocEstRefNorm & ", " & gColocEstRefVenc & ", " & gColocEstRefMor & ", " & gColocEstVigNorm & ", " & gColocEstVigVenc & ", " & gColocEstVigMor & ")) "
 sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd & " And A.cAgeCod IN " & sCadAge
 sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion

    If pnConCuotaLibre = True Then
        sSql = sSql & " And CC.nColocCalendCod=" & gColocCalendCodCL & " "
    End If

 sSql = sSql & "  Group by A.cAgeDescripcion,P3.cPersNombre,Prd.cCtaCod, (LEFT(SUBSTRING(ISNULL(P2.cPersNombre,'AVAL DESCONOCIDO'),1,32) + SPACE(32), 32) + ' ' +  LEFT(SUBSTRING(ISNULL(P2.cPersDireccDomicilio,'NO EXISTE DIRECCION'),1,32)+ SPACE(32), 32))"
 sSql = sSql & "  ORDER BY cAgencia, cAnalista, cCtaCod, cTipo, ndatos"


' ---------------------------------------------------------------------------------------------------------------------
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaMoraxAnalista_AtrasoPagoCuotaLibre = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosProtestados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer

    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sSql = "Select A.cAgeDescripcion, CN.cConsDescripcion, Prd.cCtaCod, Pers.cPersNombre,"
    sSql = sSql & " Prd.nSaldo, SUBSTRING(M.cMovNro,1,8) as sFecha, CC.nNroProxCuota, CD.dVenc,"
    sSql = sSql & " CDet.nMonto as nCuotaApr, CDet.nMonto - CDet.nMontoPagado as nCuotaXAmort, CDet.nMontoPagado as nCuotaAmort,"
    sSql = sSql & " CC.nDiasAtraso"
    sSql = sSql & " From     Producto Prd"
    sSql = sSql & "  Inner Join Constante CN ON Convert(Int,SUBSTRING(Prd.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSql = sSql & "  Inner Join ProductoPersona PP ON PP.cCtaCod = Prd.cCtaCod And PP.nPrdPersrelac = " & gColRelPersTitular
    sSql = sSql & "  Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSql = sSql & "  LEFT Join (Select cCtaCod, MAX(nMovNro) as nMovNro  From MovCol Where SUBSTRING(cCtaCod,6,3) <> '305'  AND cOpeCod like '100[234567]%' Group by cCtaCod) as MC ON MC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "  LEFT Join Mov M ON M.nMovnro = MC.nMovNro"
    sSql = sSql & "  Inner join ColocacCred CC ON CC.cCtacod = Prd.cCtaCod"
    sSql = sSql & "  Inner Join ColocCalendario CD ON CD.cCtaCod = Prd.cCtaCod And CD.nNroCalen = CC.nNroCalen And CD.nCuota=CC.nNroProxCuota And CD.nColocCalendApl = " & gColocCalendAplCuota
    sSql = sSql & "  Inner Join (Select cCtaCod, nNroCalen, nCuota, SUM(nMonto) as nMonto, SUM(nMontoPagado) as nMontoPagado From ColocCalendDet Where nColocCalendApl =" & gColocCalendAplCuota & " Group by cCtaCod, nNroCalen, nCuota) as CDet"
    sSql = sSql & "     On CDet.cCtaCod = Prd.cCtaCod And CDet.nNroCalen = CC.nNroCalen And CDet.nCuota = CC.nNroProxCuota"
    sSql = sSql & "  Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & " WHERE CC.cProtesto=1  and SUBSTRING(Prd.cCtaCod,6,3) in ('101','102','103','104','201','202','203','221','301','302','303','304','320','401','423')"
    sSql = sSql & "  And LEN(Prd.cCtacod)=18"
    sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & " And A.cAgeCod IN " & sCadAge
    sSql = sSql & "        And Prd.nPrdEstado in (2020,2021,2022,2030,2031,2032) "
    sSql = sSql & " Order by A.cAgeDescripcion, CN.cConsDescripcion, Pers.cPersNombre "

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosProtestados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosRetirados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant) As ADODB.Recordset

Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sSql As String

    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For i = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
    Next i
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"

    sSql = " Select A.cAgeDescripcion, Prd.cCtaCod, P.cPersNombre, CE2.nMonto, CN.cConsDescripcion, CE2.dPrdEstado, RH.cUser"
    sSql = sSql & " From Producto Prd"
    sSql = sSql & " Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = Prd.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & " Inner Join Persona P ON P.cPersCod = PP.cPersCod"
    sSql = sSql & " Inner Join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = " & gColocEstRetirado
    sSql = sSql & " Inner Join ColocacEstado CE2 ON CE2.cCtaCod = Prd.cCtaCod And CE2.nPrdEstado = " & gColocEstAprob
    sSql = sSql & " Inner join Constante CN ON CN.nConsValor = CE.nMotivoRechazo And CN.nConsCod = " & gColocMotivRechazo
    sSql = sSql & " Inner Join ProductoPersona PrdPers ON PrdPers.cCtaCod = Prd.cCtaCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & " Inner join RRHH RH ON RH.cPersCod = PrdPers.cPersCod"
    sSql = sSql & " Inner Join Agencias A ON A.cAgeCod = substring(Prd.cCtaCod,4,2)"
    sSql = sSql & " Where Prd.nPrdEstado = " & gColocEstRetirado & " and Prd.dPrdEstado >= '" & Format(pdFecIni, "mm/dd/yyyy") & "' And Prd.dPrdEstado <= '" & Format(pdFecFin, "mm/dd/yyyy") & "' "
    sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd & " And A.cAgeCod IN " & sCadAge
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & " Order by A.cAgeDescripcion, P.cpersNombre "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosRetirados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaCartaCreditosMorosos(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, ByVal pnDiasAtrIni As Integer, _
    ByVal pnDiasAtrFin As Integer, ByVal pMatAnalistas As Variant) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sSql As String
Dim sCadAnalista As String

    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For i = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
    Next i
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"

    sCadAnalista = "('"
    For i = 0 To UBound(pMatAnalistas) - 1
        sCadAnalista = sCadAnalista & pMatAnalistas(i) & "','"
    Next i
    sCadAnalista = Mid(sCadAnalista, 1, Len(sCadAnalista) - 2) & ")"

    sSql = "select A.cAgeDescripcion, Pers.cPersNombre, Pers.cPersDireccDomicilio, U.cUbiGeoDescripcion as Zona, Prd.cCtaCod, Cal.nCuota, Pers2.cPersNombre as cAnalista "
    sSql = sSql & " From Producto Prd"
    sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = Prd.cCtaCod And PP.nPrdPersRelac = 20"
    sSql = sSql & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSql = sSql & " Left Join UbicacionGeografica U ON U.cUbiGeoCod = Pers.cPersDireccUbigeo"
    sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Inner Join (Select cCtaCod, nNroCalen, Max(nCuota) as nCuota From ColocCalendario"
    sSql = sSql & "         Where nColocCalendEstado = 0 And nColocCalendApl = 1 "
    sSql = sSql & "          Group By cCtaCod, nNroCalen"
    sSql = sSql & "        ) as Cal ON Cal.cCtaCod = Prd.cCtaCod And Cal.nNroCalen = CC.nNroCalen"
    sSql = sSql & " Inner Join ProductoPersona PP2 ON PP2.cCtaCod = Prd.cCtaCod And PP2.nPrdPersRelac = 28 And PP2.cPersCod in " & sCadAnalista
    sSql = sSql & " Inner Join Persona Pers2 ON Pers2.cPersCod = PP2.cPersCod"
    sSql = sSql & " Inner Join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & " Where SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd & " And A.cAgeCod IN " & sCadAge
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & "        And CC.nDiasAtraso >= " & pnDiasAtrIni & " And CC.nDiasAtraso <= " & pnDiasAtrFin
    sSql = sSql & " ORDER BY A.cAgeDescripcion, Pers.cPersNombre"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCartaCreditosMorosos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCartaInvitacionCreditosParalelos(ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, ByVal pnNotaIni As Integer, _
    ByVal pnNotaFin As Integer, ByVal pMatAnalistas As Variant, ByVal pnUltCuotas As Integer, ByVal pbPorc As Integer) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sSql As String
Dim sCadAnalista As String

    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For i = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
    Next i
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"

    sCadAnalista = "('"
    For i = 0 To UBound(pMatAnalistas) - 1
        sCadAnalista = sCadAnalista & pMatAnalistas(i) & "','"
    Next i
    sCadAnalista = Mid(sCadAnalista, 1, Len(sCadAnalista) - 2) & ")"

    sSql = "select A.cAgeDescripcion, Pers.cPersNombre, Pers.cPersDireccDomicilio, U.cUbiGeoDescripcion as Zona, Prd.cCtaCod, Cal.nCuotas, Pers2.cPersNombre as cAnalista "
    sSql = sSql & " From Producto Prd"
    sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = Prd.cCtaCod And PP.nPrdPersRelac = 20"
    sSql = sSql & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSql = sSql & " Left Join UbicacionGeografica U ON U.cUbiGeoCod = Pers.cPersDireccUbigeo"
    sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Inner Join (Select cCtaCod, nNroCalen, COUNT(nCuota) as nCuotas From ColocCalendario"
    sSql = sSql & "         Where nColocCalendEstado = 0 And nColocCalendApl = 1 "
    sSql = sSql & "          Group By cCtaCod, nNroCalen"
    sSql = sSql & "        ) as Cal ON Cal.cCtaCod = Prd.cCtaCod And Cal.nNroCalen = CC.nNroCalen"
    sSql = sSql & " Inner Join ProductoPersona PP2 ON PP2.cCtaCod = Prd.cCtaCod And PP2.nPrdPersRelac = 28 And PP2.cPersCod in " & sCadAnalista
    sSql = sSql & " Inner Join Persona Pers2 ON Pers2.cPersCod = PP2.cPersCod"
    sSql = sSql & " Inner Join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & " Inner Join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSql = sSql & " Left Join ( Select cCtaCod, nColocNota From ColocCalificacionAnalista CA Where dColocNotaFecha = (Select MAX(dColocNotaFecha) From ColocCalificacionAnalista Where cCtaCod = CA.cCtaCod)) as CA "
    sSql = sSql & " ON CA.cCtaCod = Prd.cCtaCod "
    sSql = sSql & " Where SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd & " And A.cAgeCod IN " & sCadAge
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & " And CA.nColocNota  >= " & pnNotaIni & " And CA.nColocNota  <= " & pnNotaFin
    If pbPorc <> 1 Then
        sSql = sSql & "    And Cal.nCuotas = " & pnUltCuotas
    Else
        sSql = sSql & "    And Prd.nSaldo <= CE.nMonto * " & Format(pnUltCuotas / 100, "#0.00")
    End If
    sSql = sSql & " ORDER BY A.cAgeDescripcion, Pers.cPersNombre"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCartaInvitacionCreditosParalelos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosPorUbicacionGeo(ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, _
    ByVal psCodUbic As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
    
    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For i = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
    Next i
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"
    
    sSql = "select A.cAgeDescripcion, Pers.cPersNombre, Pers.cPersDireccDomicilio, U.cUbiGeoDescripcion as Zona,"
    sSql = sSql & " Prd.cCtaCod, C.nMontoCol nMonto, Prd.nSaldo, CA.nColocNota, CC.nNroProxCuota, "
    sSql = sSql & " nCuotas = (SELECT Count(*) FROM ColocCalendario WHERE cCtaCod = C.cCtaCod AND nNroCalen = CC.nNroCalen AND nColocCalendApl = 1), "
    sSql = sSql & " RH.cUser From Producto Prd "
    sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = Prd.cCtaCod And PP.nPrdPersRelac=" & gColRelPersTitular & " "
    sSql = sSql & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSql = sSql & " Inner Join UbicacionGeografica U ON U.cUbiGeoCod = Pers.cPersDireccUbigeo And U.cUbiGeoCod = '" & Trim(Right(psCodUbic, 15)) & "'"
    sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod INNER JOIN COLOCACIONES C ON CC.cCtaCod = C.cCtaCod "
    sSql = sSql & " Left Join (Select cCtaCod, nColocNota From ColocCalificacionAnalista CAN Where dColocNotaFecha = (Select MAX(dColocNotaFecha) From ColocCalificacionAnalista Where cCtaCod = CAN.cCtaCod)) as CA"
    sSql = sSql & "     ON CA.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Inner Join ProductoPersona PP2 ON PP2.cCtaCod = Prd.cCtaCod And PP2.nPrdPersRelac=" & gColRelPersAnalista & " "
    sSql = sSql & " Inner Join RRHH RH ON RH.cPersCod = PP2.cPersCod"
    sSql = sSql & " Inner Join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & " And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd & " And A.cAgeCod IN " & sCadAge
    sSql = sSql & " And CC.nColocCondicion in " & sCadCondicion & " ORDER BY PP.cCtaCod "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosPorUbicacionGeo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCredVig_CredVigConCuotLib(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, ByVal pnDiasAtrIni As Integer, _
    ByVal pnDiasAtrFin As Integer, ByVal pnTipoCambio As Double, ByVal pnConCuotaLibre As Boolean) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadProd As String
Dim sCadCondicion As String
   
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    
    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sCadCondicion = "("
    For i = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
    Next i
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"
    
    sSql = " SELECT "
    sSql = sSql & "  CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSql = sSql & "   END AS sAgencia,"
    sSql = sSql & "   CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')"
    sSql = sSql & "   END AS sProducto,"
    sSql = sSql & "   CASE WHEN (GROUPING(CN2.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(CN2.cConsDescripcion, 'PLAZO DESCONOCIDO ')"
    sSql = sSql & "   END AS sPlazo,"
    sSql = sSql & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
    sSql = sSql & "   END AS sFuenteFinan,"
    sSql = sSql & "   CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL'"
    sSql = sSql & "    ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDA')"
    sSql = sSql & "   END AS cCtaCod,"
    sSql = sSql & " COUNT(*) as nCasos,"
    
    'If pnConCuotaLibre = True Then
    '    sSql = sSql & " cc.ncoloccalendcod, "
    'End If
    
    sSql = sSql & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac=" & gColRelPersTitular & " Where PP.cCtaCod = Prd.cCtaCod)," 'Decia 20
    sSql = sSql & " sInstitucion = (Select cPersNombre From Persona Where cPersCod = (Select Top 1 cPersCod From ColocacConvenio Where cCtaCod = Prd.cCtaCod)),"
    sSql = sSql & " dVigencia = (select dVigencia From Colocaciones Where cCtaCod = Prd.cCtaCod),"
    sSql = sSql & " sPlazoCuota = (Select CONVERT(varchar(6),nCuotas) + '/' +  CONVERT(varchar(6), Case nPlazo When 0 Then 30 Else nPlazo End) "
    sSql = sSql & " From ColocacEstado Where cCtaCod = Prd.cCtaCod And nPrdEstado = " & gColocEstAprob & "),"
    sSql = sSql & " SUM(isnull(CE.nMonto,0)) as nMontoApr,"
    sSql = sSql & " SUM(isnull(Prd.nSaldo,0)) as nSaldoCap,"
    sSql = sSql & " nCuotaMora = (select Case WHEN nDiasAtraso <=0 THEN 0 ELSE nNroProxCuota END From ColocacCred Where cCtaCod = Prd.cCtaCod),"
    sSql = sSql & " nDiasAtraso = (select Case WHEN nDiasAtraso <= 0 THEN 0 ELSE nDiasAtraso END From ColocacCred Where cCtaCod = Prd.cCtaCod),"
    sSql = sSql & " nMontoCuota = (Select SUM(round(nMonto,2)) from ColocCalendDet"
    sSql = sSql & "        Where cCtaCod = Prd.cCtaCod And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = Prd.cCtaCod)"
    sSql = sSql & "        And nCuota = (select nNroProxCuota From ColocacCred Where cCtaCod = Prd.cCtaCod)"
    sSql = sSql & "        And nColocCalendApl=" & gColocCalendAplCuota & "), " 'Decia 1
    sSql = sSql & " dFecUtlPago = (Select dbo.FechaMov(M.cMovnro) From Mov M Where nMovNro = "
    sSql = sSql & " (Select MAX(nMovNro) From MovCol where cCtaCod = Prd.cCtaCod AND cOpeCod like '100[234567]%')),"
    sSql = sSql & " SUM(dbo.InteresDevengado(prd.cCtaCod,'" & Format(pdFecIni, "mm/dd/yyyy") & "') ) as IntDeveng,"
    sSql = sSql & " SUM(dbo.DeudaTotal(Prd.cCtaCod)) as nDeudaTotal,"
    sSql = sSql & " SUM(dbo.CapitalVencido(Prd.cCtaCod,'" & Format(pdFecIni, "mm/dd/yyyy") & "')) as nCapVenc,"
    sSql = sSql & " SUM(dbo.CapitalXVencer(Prd.cCtaCod,'" & Format(pdFecIni, "mm/dd/yyyy") & "')) as nCapXVencer,"
    sSql = sSql & " SUM(dbo.GarantiaHipotecaria(Prd.cCtaCod," & Format(pnTipoCambio, "#0.000") & ")) as nMontoGarHip,"
    sSql = sSql & " SUM(dbo.OtrasGarantias(Prd.cCtaCod," & Format(pnTipoCambio, "#0.000") & ")) as nMontoOtrasGarant,"
    sSql = sSql & " cEstadoDesc = (Select CN2.cConsDescripcion From Producto Prd2 Inner Join Constante CN2 ON CN2.nConsValor = Prd2.nPrdEstado And CN2.nConsCod = " & gColocEstado & " Where Prd2.cCtaCod = Prd.cCtaCod)"
    sSql = sSql & " FROM Producto Prd"
    sSql = sSql & " Inner Join Constante CN ON Convert(Int,SUBSTRING(Prd.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSql = sSql & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & " Inner Join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod And CC.nDiasAtraso >= " & pnDiasAtrIni & " And CC.nDiasAtraso <=" & pnDiasAtrFin
    sSql = sSql & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = C.cLineaCred"
    sSql = sSql & " Inner join Constante CN2 ON CN2.nConsValor = Convert(int,SUBSTRING(C.cLineaCred,6,1)) And CN2.nConsCod = " & gColocLineaCredPlazo
    sSql = sSql & " Inner Join Persona P ON CL.cPersCod = P.cPersCod"
    sSql = sSql & " Inner Join ColocacEstado CE ON CE.cCtaCod = Prd.CCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSql = sSql & " WHERE (Prd.nPrdEstado IN (" & gColocEstRefNorm & ", " & gColocEstRefVenc & ", " & gColocEstRefMor & ", " & gColocEstVigNorm & ", " & gColocEstVigVenc & ", " & gColocEstVigMor & ")) "
    sSql = sSql & " And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd & " And A.cAgeCod IN " & sCadAge
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    
    If pnConCuotaLibre = True Then
         sSql = sSql & " And CC.nColocCalendCod=" & gColocCalendCodCL & " "
    End If
    
    sSql = sSql & " GROUP BY A.cAgeDescripcion, CN.cConsDescripcion, CN2.cConsDescripcion, P.cPersNombre, Prd.cCtaCod"
    
    sSql = sSql & " WITH ROLLUP"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCredVig_CredVigConCuotLib = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosXInstitucion(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, _
    ByVal pdFecIni As Date, ByVal pnTipoOrden As Integer, pnIncluyeMora As Integer, ByVal MatInstitucion As Variant) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadMora As String
Dim sCadTipoOrden As String
Dim sCadInst As String

   sCadMora = ""
   If pnIncluyeMora = 0 Then
        sCadMora = "And nPrdConceptoCod <> 1101"
   End If
    
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
        
    sCadInst = "('"
    For i = 0 To UBound(MatInstitucion) - 1
        sCadInst = sCadInst & MatInstitucion(i) & "','"
    Next i
    sCadInst = Mid(sCadInst, 1, Len(sCadInst) - 2) & ")"

    sCadTipoOrden = ""
    Select Case pnTipoOrden
        Case 0 'Por Codigo Modular
            sCadTipoOrden = "cCodModular"
        Case 1 'Orden Alfabetico
            sCadTipoOrden = "sTitular"
        Case 2 'Pagare
            sCadTipoOrden = "cCtaCod"
    End Select
    
    sSql = "        SELECT * "
    sSql = sSql & " From "
    sSql = sSql & " (SELECT  sAgencia, cCtaCod, cInstitucion, nCasos, isnull(sTitular,'ZZZZ') as sTitular, cCodModular, dFecDesemb, nDiasAtraso, nMontoApr, nSaldoCap, "
    sSql = sSql & "         sCuotaPend , nSaldoCuota, nCuota, nITFSalCuota, nItfCuota, dFecPago, nCapPend, nTasMor "
    sSql = sSql & "         From "
    sSql = sSql & "         (SELECT "
    sSql = sSql & "                 CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "                     ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSql = sSql & "                     END AS sAgencia,"
    sSql = sSql & "                 CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL'"
    sSql = sSql & "                     ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDA')"
    sSql = sSql & "                     END AS cCtaCod,"
    sSql = sSql & "                 CASE WHEN GROUPING(P.cPersNombre)=1 THEN 'TOTAL'"
    sSql = sSql & "                     ELSE ISNULL(P.cPersNombre, 'INSTITUCION DESCONOCIDA')"
    sSql = sSql & "                     END AS cInstitucion,"
    sSql = sSql & "                 COUNT(*) as nCasos, "
    sSql = sSql & "                 sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = Prd.cCtaCod),"
    sSql = sSql & "                 cCodModular = (Select cCodModular From ColocacConvenio CV Where CV.cCtaCod = Prd.cCtaCod),"
    sSql = sSql & "                 dFecDesemb = (Select dVigencia From Colocaciones Where cCtaCod = Prd.cCtaCod),"
    sSql = sSql & "                 nDiasAtraso = (Select nDiasAtraso From ColocacCred Where cCtaCod = Prd.cCtaCod), "
    sSql = sSql & "                 SUM(CE.nMonto) as nMontoApr, "
    sSql = sSql & "                 SUM(Prd.nSaldo) as nSaldoCap, "
    sSql = sSql & "                 sCuotaPend = (Select CONVERT(varchar(5),CC.nNroProxCuota) + '/' + CONVERT(varchar(5),CE.nCuotas) From ColocacCred CC Inner Join ColocacEstado CE ON CC.cCtaCod = CE.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & "                               Where CC.cCtaCod = Prd.cCtaCod ),"
    sSql = sSql & "                 SUM(CDet.nMonto) as nSaldoCuota,SUM(CDet.nMonCuota) as nCuota, "
    sSql = sSql & "                 ROUND(dbo.ITFImpuesto(SUM(CDet.nMonto)) ,2) AS nITFSalCuota, ROUND(dbo.ITFImpuesto(SUM(CDet.nMonCuota)) ,2) as nItfCuota, "
    sSql = sSql & "                 dFecPago = (Select dVenc From ColocCalendario CCal Where CCal.cCtaCod = Prd.cCtaCod And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CCal.cCtaCod) And nColocCalendApl = 1 And nCuota = (Select nNroProxCuota From ColocacCred Where cCtaCod = CCal.cCtaCod)), "
    sSql = sSql & "                 nTasMor  = (SELECT PT.nTasaInteres FROM ProductoTasaInteres PT where PT.cctacod = Prd.cCtaCod and PT.nPrdTasaInteres=3), SUM(nCapPend) AS nCapPend "
    sSql = sSql & "            From Producto Prd"
    sSql = sSql & "                 Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & "                 Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "                 Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSql = sSql & "                 Inner Join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & "                 Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "                 Inner Join ( Select cCtaCod, nNroCalen, nCuota,  SUM(nMonto-nMontoPagado) as nMonto,SUM(nMonto) as nMonCuota, nCapPend= SUM(CASE WHEN nPrdConceptoCod=1000 THEN nMonto-nMontoPagado ELSE 0 END) "
    sSql = sSql & "                              From ColocCalendDet CDet "
    sSql = sSql & "                              Where nColocCalendApl = 1 " & sCadMora
    sSql = sSql & "                                 Group By cCtaCod, nNroCalen, nCuota"
    sSql = sSql & "                                 ) as CDet ON CDet.cCtaCod = Prd.cCtaCod And CDet.nNroCalen = CC.nNroCalen And CDet.nCuota = CC.nNroProxCuota"
    sSql = sSql & "             Where Prd.nPrdEstado >= 2020 And Prd.nPrdEstado <= 2032"
    sSql = sSql & "                     And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' And A.cAgeCod IN " & sCadAge & " AND CV.cPersCod in " & sCadInst
    sSql = sSql & "            Group By A.cAgeDescripcion, P.cPersNombre, Prd.cCtaCod"
    sSql = sSql & "             WITH ROLLUP ) as X ) as Y"
    sSql = sSql & " order by  cInstitucion ASC, sagencia," & sCadTipoOrden

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosXInstitucion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosXInstitucionDetAnt(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, _
    ByVal pdFecIni As Date, ByVal pnTipoOrden As Integer, pnIncluyeMora As Integer, ByVal MatInstitucion As Variant, ByVal pdFecSis As Date, _
    Optional ByVal pnDia As Integer = 0) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadMora As String
Dim sCadTipoOrden As String
Dim sCadInst As String

   sCadMora = ""
   If pnIncluyeMora = 0 Then
        sCadMora = "And nPrdConceptoCod <> 1101"
   End If
    
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
        
    sCadInst = "('"
    For i = 0 To UBound(MatInstitucion) - 1
        sCadInst = sCadInst & MatInstitucion(i) & "','"
    Next i
    sCadInst = Mid(sCadInst, 1, Len(sCadInst) - 2) & ")"

    sCadTipoOrden = ""
    Select Case pnTipoOrden
        Case 0 'Por Codigo Modular
            sCadTipoOrden = "cCodModular"
        Case 1 'Orden Alfabetico
            sCadTipoOrden = "sTitular"
        Case 2 'Pagare
            sCadTipoOrden = "cCtaCod"
    End Select
    
    'datediff(day,'" & Format(pdFecSis, "mm/dd/yyyy") & "','" & Format(pdFecIni, "mm/dd/yyyy") & "')
    
    sSql = "        SELECT *  "
    sSql = sSql & " From  ("
    sSql = sSql & "         SELECT  sAgencia, cCtaCod, cInstitucion, cTotCuota, nCasos, isnull(sTitular,'ZZZZ') as sTitular, cCodModular,"
    sSql = sSql & "                 dFecDesemb, nDiasAtraso, nMontoApr, nSaldoCap, nDiaAtraCuota, sCuotaPend , nSaldoCuota, nCuota, nITFSalCuota,"
    sSql = sSql & "                 nItfCuota , dFecPago, nCapPend, nTasMor,"
    sSql = sSql & "                 case when nDiaAtraCuota<=0 then 0 else ROUND(nDiaAtraCuota*nCapPend*(nTasMor/100),2) END AS nMoraCalc,"
    sSql = sSql & "                 case when nDiaAtraCuota<=0 then 0 else ROUND(dbo.ITFImpuesto(datediff(day,'" & Format(pdFecSis, "mm/dd/yyyy") & "','" & Format(pdFecIni, "mm/dd/yyyy") & "')*nCapPend*(nTasMor/100)),2) end AS nItfCalc "
    sSql = sSql & "         From    (SELECT CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "                             ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ') END AS sAgencia,"
    sSql = sSql & "                         CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL'"
    sSql = sSql & "                             ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDA') END AS cCtaCod,"
    sSql = sSql & "                         CASE WHEN GROUPING(P.cPersNombre)=1 THEN 'TOTAL'"
    sSql = sSql & "                             ELSE ISNULL(P.cPersNombre, 'INSTITUCION DESCONOCIDA') END AS cInstitucion,"
    sSql = sSql & "                         CASE WHEN GROUPING(CDET.nCuota)=1 THEN 'TOTAL DEUDA' "
    sSql = sSql & "                               ELSE ISNULL('', '') END AS cTotCuota, "
    sSql = sSql & "                     COUNT(*) as nCasos,"
    sSql = sSql & "                     sTitular = (    select P2.cPersNombre"
    sSql = sSql & "                                     From    Persona P2"
    sSql = sSql & "                                     Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20"
    sSql = sSql & "                                     Where PP.cCtaCod = Prd.cCtaCod),"
    sSql = sSql & "                     cCodModular = ( Select  cCodModular"
    sSql = sSql & "                                     From    ColocacConvenio CV"
    sSql = sSql & "                                     Where CV.cCtaCod = Prd.cCtaCod),"
    sSql = sSql & "                     dFecDesemb = (  Select dVigencia"
    sSql = sSql & "                                     From Colocaciones"
    sSql = sSql & "                                     Where cCtaCod = Prd.cCtaCod),"
    sSql = sSql & "                     nDiasAtraso = ( Select nDiasAtraso"
    sSql = sSql & "                                     From ColocacCred"
    sSql = sSql & "                                     Where cCtaCod = Prd.cCtaCod),"
    sSql = sSql & "                     SUM(CE.nMonto) as nMontoApr,"
    sSql = sSql & "                     SUM(Prd.nSaldo) as nSaldoCap,"
    sSql = sSql & "                     sCuotaPend = (  Select  CONVERT(varchar(5),CDET.nCuota) + '/' + CONVERT(varchar(5),CE.nCuotas)"
    sSql = sSql & "                                     From    ColocacCred CC"
    sSql = sSql & "                                     Inner Join ColocacEstado CE ON CC.cCtaCod = CE.cCtaCod And CE.nPrdEstado = 2002 "
    sSql = sSql & "                                     Where   CC.cCtaCod = Prd.cCtaCod ),"
    sSql = sSql & "                     SUM(CDet.nMonto) as nSaldoCuota,SUM(CDet.nMonCuota) as nCuota,"
    sSql = sSql & "                     ROUND(dbo.ITFImpuesto(SUM(CDet.nMonto)) ,2) AS nITFSalCuota,"
    sSql = sSql & "                     ROUND(dbo.ITFImpuesto(SUM(CDet.nMonCuota)) ,2) as nItfCuota,"
    sSql = sSql & "                     dFecPago = (    Select  dVenc"
    sSql = sSql & "                                     From    ColocCalendario CCal"
    sSql = sSql & "                                     Where   CCAL.Cctacod = Prd.Cctacod"
    sSql = sSql & "                                             And nNroCalen = (   Select nNroCalen"
    sSql = sSql & "                                                         From ColocacCred"
    sSql = sSql & "                                                         Where cCtaCod = CCal.cCtaCod)"
    sSql = sSql & "                                             And nColocCalendApl = 1"
    sSql = sSql & "                                             And nCuota = CDet.nCuota),"
    sSql = sSql & "                      nDiaAtraCuota = (  Select datediff(day,dVenc, '" & Format(pdFecIni, "mm/dd/yyyy") & "') as nAtrCuota"
    sSql = sSql & "                                         From    ColocCalendario CCal"
    sSql = sSql & "                                         Where CCAL.Cctacod = Prd.Cctacod"
    sSql = sSql & "                                         And nNroCalen = (   Select nNroCalen "
    sSql = sSql & "                                                             From ColocacCred "
    sSql = sSql & "                                                             Where cCtaCod = CCal.cCtaCod)"
    sSql = sSql & "                                         And nColocCalendApl = 1"
    sSql = sSql & "                                         And nCuota = CDet.nCuota),"
    sSql = sSql & "                     nTasMor  = (    SELECT  PT.nTasaInteres"
    sSql = sSql & "                                     FROM    ProductoTasaInteres PT"
    sSql = sSql & "                                     where PT.cctacod = Prd.cCtaCod and PT.nPrdTasaInteres=3),"
    sSql = sSql & "                     SUM(nCapPend) As nCapPend"
    sSql = sSql & "             From    Producto Prd"
    sSql = sSql & "                     Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & "                     Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "                     Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSql = sSql & "                     Inner Join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & "                     Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "                     Inner Join (    Select  CDet.cCtaCod, CDet.nNroCalen, CDet.nCuota,  SUM( CASE WHEN CCAL.dVenc <= '" & Format(pdFecIni, "mm/dd/yyyy") & "' THEN CDet.nMonto-CDet.nMontoPagado ELSE 0 END) as nMonto, SUM( CASE WHEN CCAL.dVenc <= '" & Format(pdFecIni, "mm/dd/yyyy") & "' THEN CDet.nMonto ELSE 0 END) as nMonCuota, "
    sSql = sSql & "                                             nCapPend= SUM( CASE WHEN CCAL.dVenc <= '" & Format(pdFecIni, "mm/dd/yyyy") & "' THEN  ( CASE WHEN CDet.nPrdConceptoCod=1000 THEN CDet.nMonto-CDet.nMontoPagado ELSE 0 END ) ELSE 0 END) "
    sSql = sSql & "                                     From    ColocCalendario CCAL"
    sSql = sSql & "                                             JOIN ColocCalendDet CDet ON CDet.cCtaCod = CCAL.cCtaCod AND CDet.nNroCalen = CCAL.nNroCalen "
    sSql = sSql & "                                             AND CDet.nColocCalendApl = CCAL.nColocCalendApl AND CDet.nCuota = CCAL.nCuota "
    sSql = sSql & "                                     Where   CDet.nColocCalendApl = 1 AND CCAL.nColocCalendEstado=0 AND  CCAL.dVenc <= '" & Format(pdFecIni + 30, "mm/dd/yyyy") & "' "
    sSql = sSql & "                                     Group By CDet.cCtaCod, CDet.nNroCalen, CDet.nCuota ) as CDet ON CDet.cCtaCod = Prd.cCtaCod And CDet.nNroCalen = CC.nNroCalen "
    sSql = sSql & "               Where Prd.nPrdEstado >= 2020 And Prd.nPrdEstado <= 2032 "
    sSql = sSql & "                     And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' And A.cAgeCod IN " & sCadAge & " AND CV.cPersCod in " & sCadInst
    sSql = sSql & "            Group By A.cAgeDescripcion, P.cPersNombre, Prd.cCtaCod, CDET.nCuota"
    sSql = sSql & "             WITH ROLLUP ) as X ) as Y"
    If pnDia > 0 Then
    sSql = sSql & "              Where DatePart(day,dFecPago)=" & pnDia
    End If
    sSql = sSql & "     order by  cInstitucion ASC, sagencia," & sCadTipoOrden

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosXInstitucionDetAnt = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosXInstitucionDet(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, _
    ByVal pdFecIni As Date, ByVal pnTipoOrden As Integer, pnIncluyeMora As Integer, ByVal MatInstitucion As Variant, ByVal pdFecSis As Date, _
    Optional ByVal pnDia As Integer = 0) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadMora As String
Dim sCadTipoOrden As String
Dim sCadInst As String

   sCadMora = ""
   If pnIncluyeMora = 0 Then
        sCadMora = "And nPrdConceptoCod <> 1101"
   End If
    
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
        
    sCadInst = "('"
    For i = 0 To UBound(MatInstitucion) - 1
        sCadInst = sCadInst & MatInstitucion(i) & "','"
    Next i
    sCadInst = Mid(sCadInst, 1, Len(sCadInst) - 2) & ")"

    sCadTipoOrden = ""
    Select Case pnTipoOrden
        Case 0 'Por Codigo Modular
            sCadTipoOrden = "cCodModular"
        Case 1 'Orden Alfabetico
            sCadTipoOrden = "sTitular"
        Case 2 'Pagare
            sCadTipoOrden = "cCtaCod"
    End Select
    
    'datediff(day,'" & Format(pdFecSis, "mm/dd/yyyy") & "','" & Format(pdFecIni, "mm/dd/yyyy") & "')
    
    sSql = "        SELECT *  "
    sSql = sSql & " From  ("
    sSql = sSql & "         SELECT  sAgencia, cCtaCod, cInstitucion, nCasos, isnull(sTitular,'ZZZZ') as sTitular, cCodModular,"
    sSql = sSql & "                 dFecDesemb, nDiasAtraso, nMontoApr, nSaldoCap, nDiaAtraCuota, sCuotaPend , nSaldoCuota, nCuota, nITFSalCuota,"
    sSql = sSql & "                 nItfCuota , dFecPago, nCapPend, nTasMor,"
    sSql = sSql & "                 case when nDiaAtraCuota<=0 then 0 else ROUND(nDiaAtraCuota*nCapPend*(nTasMor/100),2) END AS nMoraCalc,"
    sSql = sSql & "                 case when nDiaAtraCuota<=0 then 0 else ROUND(dbo.ITFImpuesto(datediff(day,'" & Format(pdFecSis, "mm/dd/yyyy") & "','" & Format(pdFecIni, "mm/dd/yyyy") & "')*nCapPend*(nTasMor/100)),2) end AS nItfCalc "
    sSql = sSql & "         From    (SELECT A.cAgeDescripcion as sAgencia, Prd.cCtaCod as cCtaCod , P.cPersNombre as cInstitucion, "
    sSql = sSql & "                     COUNT(*) as nCasos,"
    sSql = sSql & "                     sTitular = (    select P2.cPersNombre"
    sSql = sSql & "                                     From    Persona P2"
    sSql = sSql & "                                     Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20"
    sSql = sSql & "                                     Where PP.cCtaCod = Prd.cCtaCod),"
    sSql = sSql & "                     cCodModular = ( Select  cCodModular"
    sSql = sSql & "                                     From    ColocacConvenio CV"
    sSql = sSql & "                                     Where CV.cCtaCod = Prd.cCtaCod),"
    sSql = sSql & "                     dFecDesemb = (  Select dVigencia"
    sSql = sSql & "                                     From Colocaciones"
    sSql = sSql & "                                     Where cCtaCod = Prd.cCtaCod),"
    sSql = sSql & "                     nDiasAtraso = ( Select nDiasAtraso"
    sSql = sSql & "                                     From ColocacCred"
    sSql = sSql & "                                     Where cCtaCod = Prd.cCtaCod),"
    sSql = sSql & "                     SUM(CE.nMonto) as nMontoApr,"
    sSql = sSql & "                     SUM(Prd.nSaldo) as nSaldoCap,"
    sSql = sSql & "                     sCuotaPend = (  Select  CONVERT(varchar(5),CDET.nCuota) + '/' + CONVERT(varchar(5),CE.nCuotas)"
    sSql = sSql & "                                     From    ColocacCred CC"
    sSql = sSql & "                                     Inner Join ColocacEstado CE ON CC.cCtaCod = CE.cCtaCod And CE.nPrdEstado = 2002 "
    sSql = sSql & "                                     Where   CC.cCtaCod = Prd.cCtaCod ),"
    sSql = sSql & "                     SUM(CDet.nMonto) as nSaldoCuota,SUM(CDet.nMonCuota) as nCuota,"
    sSql = sSql & "                     ROUND(dbo.ITFImpuesto(SUM(CDet.nMonto)) ,2) AS nITFSalCuota,"
    sSql = sSql & "                     ROUND(dbo.ITFImpuesto(SUM(CDet.nMonCuota)) ,2) as nItfCuota,"
    sSql = sSql & "                     dFecPago = (    Select  dVenc"
    sSql = sSql & "                                     From    ColocCalendario CCal"
    sSql = sSql & "                                     Where   CCAL.Cctacod = Prd.Cctacod"
    sSql = sSql & "                                             And nNroCalen = (   Select nNroCalen"
    sSql = sSql & "                                                         From ColocacCred"
    sSql = sSql & "                                                         Where cCtaCod = CCal.cCtaCod)"
    sSql = sSql & "                                             And nColocCalendApl = 1"
    sSql = sSql & "                                             And nCuota = CDet.nCuota),"
    sSql = sSql & "                      nDiaAtraCuota = (  Select datediff(day,dVenc, '" & Format(pdFecIni, "mm/dd/yyyy") & "') as nAtrCuota"
    sSql = sSql & "                                         From    ColocCalendario CCal"
    sSql = sSql & "                                         Where CCAL.Cctacod = Prd.Cctacod"
    sSql = sSql & "                                         And nNroCalen = (   Select nNroCalen "
    sSql = sSql & "                                                             From ColocacCred "
    sSql = sSql & "                                                             Where cCtaCod = CCal.cCtaCod)"
    sSql = sSql & "                                         And nColocCalendApl = 1"
    sSql = sSql & "                                         And nCuota = CDet.nCuota),"
    sSql = sSql & "                     nTasMor  = (    SELECT  PT.nTasaInteres"
    sSql = sSql & "                                     FROM    ProductoTasaInteres PT"
    sSql = sSql & "                                     where PT.cctacod = Prd.cCtaCod and PT.nPrdTasaInteres=3),"
    sSql = sSql & "                     SUM(nCapPend) As nCapPend"
    sSql = sSql & "             From    Producto Prd"
    sSql = sSql & "                     Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & "                     Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "                     Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSql = sSql & "                     Inner Join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & "                     Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "                     Inner Join (    Select  CDet.cCtaCod, CDet.nNroCalen, CDet.nCuota,  SUM( CASE WHEN CCAL.dVenc <= '" & Format(pdFecIni, "mm/dd/yyyy") & "' THEN CDet.nMonto-CDet.nMontoPagado ELSE 0 END) as nMonto, SUM( CASE WHEN CCAL.dVenc <= '" & Format(pdFecIni, "mm/dd/yyyy") & "' THEN CDet.nMonto ELSE 0 END) as nMonCuota, "
    sSql = sSql & "                                             nCapPend= SUM( CASE WHEN CCAL.dVenc <= '" & Format(pdFecIni, "mm/dd/yyyy") & "' THEN  ( CASE WHEN CDet.nPrdConceptoCod=1000 THEN CDet.nMonto-CDet.nMontoPagado ELSE 0 END ) ELSE 0 END) "
    sSql = sSql & "                                     From    ColocCalendario CCAL"
    sSql = sSql & "                                             JOIN ColocCalendDet CDet ON CDet.cCtaCod = CCAL.cCtaCod AND CDet.nNroCalen = CCAL.nNroCalen "
    sSql = sSql & "                                             AND CDet.nColocCalendApl = CCAL.nColocCalendApl AND CDet.nCuota = CCAL.nCuota "
    sSql = sSql & "                                     Where   CDet.nColocCalendApl = 1 AND CCAL.nColocCalendEstado=0 AND  CCAL.dVenc <= '" & Format(pdFecIni + 30, "mm/dd/yyyy") & "' "
    sSql = sSql & "                                     Group By CDet.cCtaCod, CDet.nNroCalen, CDet.nCuota ) as CDet ON CDet.cCtaCod = Prd.cCtaCod And CDet.nNroCalen = CC.nNroCalen "
    sSql = sSql & "               Where Prd.nPrdEstado >= 2020 And Prd.nPrdEstado <= 2032 "
    sSql = sSql & "                     And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' And A.cAgeCod IN " & sCadAge & " AND CV.cPersCod in " & sCadInst
    sSql = sSql & "            Group By A.cAgeDescripcion, P.cPersNombre, Prd.cCtaCod, CDET.nCuota"
    sSql = sSql & "            ) as X ) as Y"
    If pnDia > 0 Then
    sSql = sSql & "              Where DatePart(day,dFecPago)=" & pnDia
    End If
    sSql = sSql & "     order by  cInstitucion ASC, sagencia," & sCadTipoOrden

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosXInstitucionDet = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function


'Public Function RecuperaCreditosXInstitucion(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, _
'    ByVal pdFecIni As Date, ByVal pnTipoOrden As Integer, pnIncluyeMora As Integer) As ADODB.Recordset
'
'Dim sSQL As String
'Dim oConecta As COMConecta.DCOMConecta
'Dim i As Integer
'Dim sCadAge As String
'Dim sCadMora As String
'Dim sCadTipoOrden As String
'
'   sCadMora = ""
'   If pnIncluyeMora = 0 Then
'        sCadMora = "And nPrdConceptoCod <> 1101"
'   End If
'
'    sCadAge = "('"
'    For i = 0 To UBound(pMatAgencias) - 1
'        sCadAge = sCadAge & pMatAgencias(i) & "','"
'    Next i
'    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
'
'    sCadTipoOrden = ""
'    Select Case pnTipoOrden
'        Case 0 'Por Codigo Modular
'            sCadTipoOrden = "CV.cCodModular"
'        Case 1 'Orden Alfabetico
'            sCadTipoOrden = "P.cPersNombre"
'        Case 2 'Pagare
'            sCadTipoOrden = "Prd.cCtaCod"
'    End Select
'
'    sSQL = " SELECT "
'    sSQL = sSQL & "  CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
'    sSQL = sSQL & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
'    sSQL = sSQL & "   END AS sAgencia,"
'    sSQL = sSQL & "   CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL'"
'    sSQL = sSQL & "    ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDA')"
'    sSQL = sSQL & "   END AS cCtaCod,"
'    sSQL = sSQL & "   CASE WHEN GROUPING(P.cPersNombre)=1 THEN 'TOTAL'"
'    sSQL = sSQL & "    ELSE ISNULL(P.cPersNombre, 'INSTITUCION DESCONOCIDA')"
'    sSQL = sSQL & "   END AS cInstitucion,"
'    sSQL = sSQL & " COUNT(*) as nCasos, "
'    sSQL = sSQL & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = Prd.cCtaCod),"
'    sSQL = sSQL & " cCodModular = (Select cCodModular From ColocacConvenio CV Where CV.cCtaCod = Prd.cCtaCod),"
'    sSQL = sSQL & " dFecDesemb = (Select dVigencia From Colocaciones Where cCtaCod = Prd.cCtaCod),"
'    sSQL = sSQL & " nDiasAtraso = (Select nDiasAtraso From ColocacCred Where cCtaCod = Prd.cCtaCod), "
'    sSQL = sSQL & " SUM(CE.nMonto) as nMontoApr,"
'    sSQL = sSQL & " SUM(Prd.nSaldo) as nSaldoCap,"
'    sSQL = sSQL & " sCuotaPend = (Select CONVERT(varchar(5),CC.nNroProxCuota) + '/' + CONVERT(varchar(5),CE.nCuotas) From ColocacCred CC Inner Join ColocacEstado CE ON CC.cCtaCod = CE.cCtaCod And CE.nPrdEstado = 2002"
'    sSQL = sSQL & "        Where CC.cCtaCod = Prd.cCtaCod ),"
'    sSQL = sSQL & " SUM(CDet.nMonto) as nCuota,"
'    sSQL = sSQL & " dFecPago = (Select dVenc From ColocCalendario CCal Where CCal.cCtaCod = Prd.cCtaCod And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CCal.cCtaCod) And nColocCalendApl = 1 And nCuota = (Select nNroProxCuota From ColocacCred Where cCtaCod = CCal.cCtaCod))"
'    sSQL = sSQL & " From Producto Prd"
'    sSQL = sSQL & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
'    sSQL = sSQL & " Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
'    sSQL = sSQL & " Inner Join Persona P ON P.cPersCod = CV.cPersCod"
'    sSQL = sSQL & " Inner Join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
'    sSQL = sSQL & " Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
'    sSQL = sSQL & " Inner Join (Select cCtaCod, nNroCalen, nCuota,  SUM(nMonto) as nMonto From ColocCalendDet CDet"
'    sSQL = sSQL & "        Where nColocCalendApl = 1 " & sCadMora
'    sSQL = sSQL & "        Group By cCtaCod, nNroCalen, nCuota"
'    sSQL = sSQL & "        ) as CDet ON CDet.cCtaCod = Prd.cCtaCod And CDet.nNroCalen = CC.nNroCalen And CDet.nCuota = CC.nNroProxCuota"
'    sSQL = sSQL & " Where Prd.nPrdEstado >= 2020 And Prd.nPrdEstado <= 2032"
'    sSQL = sSQL & " And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' And A.cAgeCod IN " & sCadAge
'    sSQL = sSQL & " Group By A.cAgeDescripcion, P.cPersNombre, Prd.cCtaCod"
'    sSQL = sSQL & " WITH ROLLUP"
'    sSQL = sSQL & " Order by sAgencia, cCtaCod, cInstitucion "
'
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set RecuperaCreditosXInstitucion = oConecta.CargaRecordSet(sSQL)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'
'End Function

Public Function RecuperaMoraXInstituciones(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, _
    ByVal pdFecIni As Date, pMatInstitucion As Variant) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadInstitucion As String

    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    sCadInstitucion = "('"
    For i = 0 To UBound(pMatInstitucion) - 1
        sCadInstitucion = sCadInstitucion & pMatInstitucion(i) & "','"
    Next i
    sCadInstitucion = Mid(sCadInstitucion, 1, Len(sCadInstitucion) - 2) & ")"
    
    sSql = " SELECT "
    sSql = sSql & "  CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSql = sSql & "   END AS sAgencia,"
    sSql = sSql & "   CASE WHEN GROUPING(P.cPersNombre)=1 THEN 'TOTAL'"
    sSql = sSql & "    ELSE ISNULL(P.cPersNombre, 'INSTITUCION DESCONOCIDA')"
    sSql = sSql & "   END AS cInstitucion,"
    sSql = sSql & "   CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL'"
    sSql = sSql & "    ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDA')"
    sSql = sSql & "   END AS cCtaCod,"
    sSql = sSql & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = Prd.cCtaCod),"
    'sSql = sSql & " nMontoDesemb = (Select nMontoCol From Colocaciones Where cCtaCod = Prd.cCtaCod),"
    sSql = sSql & " SUM(C.nMontoCol) as nMontoDesemb, "
    sSql = sSql & " SUM(Prd.nSaldo) as nSaldoCap,"
    sSql = sSql & " sCuotaPend = (Select CONVERT(varchar(5),CC.nNroProxCuota) + '/' + CONVERT(varchar(5),CE.nCuotas) From ColocacCred CC Inner Join ColocacEstado CE ON CC.cCtaCod = CE.cCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSql = sSql & "        Where CC.cCtaCod = Prd.cCtaCod ),"
    sSql = sSql & " SUM(CDet.nMonto) as nSaldoCuota,"
    sSql = sSql & " nDiasAtraso = (Select nDiasAtraso From ColocacCred Where cCtaCod = Prd.cCtaCod),"
    sSql = sSql & " SUM(CDet2.nMonto) as nSaldoMora,"
    sSql = sSql & " dFecVencimiento = (Select dVenc From ColocCalendario CCal Where CCal.cCtaCod = Prd.cCtaCod And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CCal.cCtaCod) And nColocCalendApl = 1 And nCuota = (Select nNroProxCuota From ColocacCred Where cCtaCod = CCal.cCtaCod)),"
    sSql = sSql & " COUNT(*) as nCasos"
    sSql = sSql & " From Producto Prd"
    sSql = sSql & " Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod "
    sSql = sSql & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & " Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod And CV.cPersCod in " & sCadInstitucion
    sSql = sSql & " Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Inner Join (Select Cdet.cCtaCod, Cdet.nNroCalen, SUM(Round(nMonto,2)-nMontoPagado) as nMonto "
    sSql = sSql & "     From ColocCalendDet CDet"
    sSql = sSql & "     Inner Join ColocCalendario CC ON Cdet.cCtaCod =   CC.cCtaCod AND Cdet.nNroCalen =   CC.nNroCalen AND Cdet.nCuota =   CC.nCuota AND Cdet.nColocCalendApl =   CC.nColocCalendApl"
    sSql = sSql & "     Where Cdet.nColocCalendApl = 1  AND CC.dVenc <= '" & Format(pdFecIni, "mm/dd/yyyy") & "' "
    sSql = sSql & "     Group By Cdet.cCtaCod, Cdet.nNroCalen"
    sSql = sSql & "        ) as CDet ON CDet.cCtaCod = Prd.cCtaCod And CDet.nNroCalen = CC.nNroCalen "
    sSql = sSql & " Inner Join (Select cCtaCod, nNroCalen, SUM(Round(nMonto,2)-nMontoPagado) as nMonto "
    sSql = sSql & "     From ColocCalendDet CDet"
    sSql = sSql & "     Where nColocCalendApl = 1 And nPrdConceptoCod = 1101"
    sSql = sSql & "     Group By cCtaCod, nNroCalen "
    sSql = sSql & "        ) as CDet2 ON CDet2.cCtaCod = Prd.cCtaCod And CDet2.nNroCalen = CC.nNroCalen "
    sSql = sSql & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & " And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' And A.cAgeCod IN " & sCadAge & " AND nDiasAtraso > 0 "
    sSql = sSql & " Group By A.cAgeDescripcion, P.cPersNombre, Prd.cCtaCod"
    sSql = sSql & " WITH ROLLUP"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaMoraXInstituciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaResumenSaldosCarteraPorInstitucionConsumo(ByVal pMatAgencias As Variant, _
    ByVal pMatProd As Variant, ByVal pdHoy As Date) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
    
    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sSql = "SELECT "
    sSql = sSql & " CASE WHEN (GROUPING(T.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(T.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSql = sSql & "    END AS sAgencia,"
    sSql = sSql & " CASE WHEN (GROUPING(T.cPersNombre) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(T.cPersNombre, 'INSTITUCION DESCONOCIDA')"
    sSql = sSql & "    END AS sInstitucion,"
    sSql = sSql & " SUM(T.nNroCreditosMN) as nNroCreditosMN,"
    sSql = sSql & " SUM(T.nSaldoCapMN) as nSaldoCapMN,"
    sSql = sSql & " SUM(T.nDesembolsadoMN) as nDesembolsadoMN,"
    sSql = sSql & " SUM(T.nNroCreditosME) as nNroCreditosME,"
    sSql = sSql & " SUM(T.nSaldoCapME) as nSaldoCapME,"
    sSql = sSql & " SUM(T.nDesembolsadoME) as nDesembolsadoME,"
    sSql = sSql & " SUM(T.nNroCreditos30MN) as nNroCreditos30MN,"
    sSql = sSql & " SUM(T.nSaldoCap30MN) as nSaldoCap30MN,"
    sSql = sSql & " SUM(T.nNroCreditos30ME) as nNroCreditos30ME,"
    sSql = sSql & " SUM(T.nSaldoCap30ME) As nSaldoCap30ME"
    sSql = sSql & " From ("
    sSql = sSql & " select A.cAgeDescripcion, P.cPersNombre, SUM(Prd.nSaldo) as nSaldoCapMN, SUM(C.nMontoCol) as nDesembolsadoMN, Count(*) as nNroCreditosMN, 0.00 as nSaldoCapME, 0.00 as nDesembolsadoME, 0 as nNroCreditosME,"
    sSql = sSql & "    0 as nSaldoCap30MN, 0 as nNroCreditos30MN,"
    sSql = sSql & "    0 as nSaldoCap30ME, 0 as nNroCreditos30ME"
    sSql = sSql & " From Producto as Prd"
    sSql = sSql & "    Inner Join Colocaciones C ON Prd.cCtaCod = C.cCtaCod"
    sSql = sSql & "    Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "    Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSql = sSql & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor & " and SUBSTRING(Prd.cCtaCod,9,1)='1'"
    sSql = sSql & "    And Prd.cCtaCod <>  '305' And A.cAgecod in " & sCadAge & sCadProd
    sSql = sSql & " Group by P.cPersNombre, A.cAgeDescripcion"
    sSql = sSql & " Union"
    sSql = sSql & " select A.cAgeDescripcion, P.cPersNombre, 0.00 as nSaldoCapMN, 0.00 as nDesembolsadoMN, 0 as nNroCreditosMN, SUM(Prd.nSaldo/dbo.GetTipoCambio(1,'" & Format(pdHoy, "mm/dd/yyyy") & "')) as nSaldoCapME, SUM(C.nMontoCol/dbo.GetTipoCambio(1,'" & Format(pdHoy, "mm/dd/yyyy") & "')) as nDesembolsadoME, Count(*) as nNroCreditosME,"
    sSql = sSql & "    0 as nSaldoCap30MN, 0 as nNroCreditos30MN,"
    sSql = sSql & "    0 as nSaldoCap30ME, 0 as nNroCreditos30ME"
    sSql = sSql & " From Producto as Prd"
    sSql = sSql & "    Inner Join Colocaciones C ON Prd.cCtaCod = C.cCtaCod"
    sSql = sSql & "    Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "    Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSql = sSql & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor & " and SUBSTRING(Prd.cCtaCod,9,1)='2'"
    sSql = sSql & "    And Prd.cCtaCod <>  '305' And A.cAgecod in " & sCadAge & sCadProd
    sSql = sSql & " Group by P.cPersNombre, A.cAgeDescripcion"
    sSql = sSql & " Union"
    sSql = sSql & " Select A.cAgeDescripcion, P.cPersNombre, 0.00 as nSaldoCapMN, 0.00 as nDesembolsadoMN, 0 as nNroCreditosMN, 0 as nSaldoCapME, 0 as nDesembolsadoME, 0 as nNroCreditosME,"
    sSql = sSql & "    SUM(Prd.nSaldo) as nSaldoCap30MN, Count(*) as nNroCreditos30MN,"
    sSql = sSql & "    0 as nSaldoCap30ME,  0 as nNroCreditos30ME"
    sSql = sSql & " From Producto as Prd"
    sSql = sSql & "    Inner Join Colocaciones C ON Prd.cCtaCod = C.cCtaCod"
    sSql = sSql & "    Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "    Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSql = sSql & "    Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSql = sSql & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor & " and SUBSTRING(Prd.cCtaCod,9,1)='1'"
    sSql = sSql & "    And Prd.cCtaCod <>  '305' and CC.nDiasAtraso >= 30 And A.cAgecod in " & sCadAge & sCadProd
    sSql = sSql & " Group by P.cPersNombre, A.cAgeDescripcion"
    sSql = sSql & "    Union"
    sSql = sSql & " Select A.cAgeDescripcion, P.cPersNombre, 0.00 as nSaldoCapMN, 0.00 as nDesembolsadoMN, 0 as nNroCreditosMN, 0 as nSaldoCapME, 0 as nDesembolsadoME, 0 as nNroCreditosME,"
    sSql = sSql & "    0 as nSaldoCap30MN, 0 as nNroCreditos30MN,"
    sSql = sSql & "    SUM(Prd.nSaldo/dbo.GetTipoCambio(1,'" & Format(pdHoy, "mm/dd/yyyy") & "')) as nSaldoCap30ME, Count(*) as nNroCreditos30ME"
    sSql = sSql & " From Producto as Prd"
    sSql = sSql & "    Inner Join Colocaciones C ON Prd.cCtaCod = C.cCtaCod"
    sSql = sSql & "    Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "    Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSql = sSql & "    Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSql = sSql & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor & " and SUBSTRING(Prd.cCtaCod,9,1)='2'"
    sSql = sSql & "    And Prd.cCtaCod <>  '305' and CC.nDiasAtraso >= 30 And A.cAgecod in " & sCadAge & sCadProd
    sSql = sSql & " Group by P.cPersNombre, A.cAgeDescripcion"
    sSql = sSql & "   ) as T"
    sSql = sSql & " Group by T.cAgeDescripcion, T.cPersNombre"
    sSql = sSql & " With ROLLUP"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaResumenSaldosCarteraPorInstitucionConsumo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaResumenSaldosCarteraPorAnalistaConsumo(ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal P1I As Integer, ByVal P1F As Integer, _
    ByVal P2I As Integer, ByVal P2F As Integer, ByVal P4F As Integer, _
    ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, pdFecha As Date) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String

    
    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For i = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
    Next i
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"
    
    
    sSql = "SELECT "
    sSql = sSql & "    CASE WHEN (GROUPING(T.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(T.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSql = sSql & "    END AS sAgencia,"
    sSql = sSql & " CASE WHEN (GROUPING(T.cPersNombre) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(T.cPersNombre, 'INSTITUCION DESCONOCIDA')"
    sSql = sSql & "    END AS sInstitucion,"
    sSql = sSql & " SUM(T.nNroCreditosMN) as nNroCreditosMN,"
    sSql = sSql & " SUM(T.nSaldoCapMN) as nSaldoCapMN,"
    sSql = sSql & " SUM(T.nDesembolsadoMN) as nDesembolsadoMN,"
    sSql = sSql & " SUM(T.nNroCreditosME) as nNroCreditosME,"
    sSql = sSql & " SUM(T.nSaldoCapME) as nSaldoCapMN,"
    sSql = sSql & " SUM(T.nDesembolsadoME) As nDesembolsadoME"
    sSql = sSql & " From ("
    sSql = sSql & " select A.cAgeDescripcion, P.cPersNombre, SUM(Prd.nSaldo) as nSaldoCapMN, SUM(C.nMontoCol) as nDesembolsadoMN, Count(*) as nNroCreditosMN, 0.00 as nSaldoCapME, 0.00 as nDesembolsadoME, 0 as nNroCreditosME"
    sSql = sSql & " From Producto as Prd"
    sSql = sSql & "    Inner Join Colocaciones C ON Prd.cCtaCod = C.cCtaCod"
    sSql = sSql & "    Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "    Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSql = sSql & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor & " and SUBSTRING(Prd.cCtaCod,9,1)='1'"
    sSql = sSql & "    And Prd.cCtaCod <>  '305'"
    sSql = sSql & " Group by P.cPersNombre, A.cAgeDescripcion"
    sSql = sSql & " Union"
    sSql = sSql & " select A.cAgeDescripcion, P.cPersNombre, 0.00 as nSaldoCapMN, 0.00 as nDesembolsadoMN, 0 as nNroCreditosMN, SUM(Prd.nSaldo/dbo.GetTipoCambio(1,'" & Format(pdFecha, "mm/dd/yyyy") & "')) as nSaldoCapME, SUM(C.nMontoCol/dbo.GetTipoCambio(1,'" & Format(pdFecha, "mm/dd/yyyy") & "')) as nDesembolsadoME, Count(*) as nNroCreditosME"
    sSql = sSql & " From Producto as Prd"
    sSql = sSql & "    Inner Join Colocaciones C ON Prd.cCtaCod = C.cCtaCod"
    sSql = sSql & "    Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "    Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSql = sSql & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor & " and SUBSTRING(Prd.cCtaCod,9,1)='2'"
    sSql = sSql & "    And Prd.cCtaCod <>  '305'"
    sSql = sSql & " Group by P.cPersNombre, A.cAgeDescripcion"
    sSql = sSql & "   ) as T"
    sSql = sSql & " Group by T.cAgeDescripcion, T.cPersNombre"
    sSql = sSql & " With ROLLUP"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaResumenSaldosCarteraPorAnalistaConsumo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaPagoConCheque(ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, _
    ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pnTipoBusq As Integer, ByVal psNroDoc As String) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sCadBusq As String

    If pnTipoBusq = 1 Then
        sCadBusq = " AND MD.cDocNro = '" & psNroDoc & "'"
    End If

    sCadProd = " And SUBSTRING(MC.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For i = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
    Next i
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"

    sSql = " select A.cAgeDescripcion as sAgencia, P.cPersNombre as cInstitucion, MC.cCtaCod, sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = MC.cCtaCod),"
    sSql = sSql & " cOperacion = (Select CN.cConsDescripcion From Constante CN  where CN.nConsValor = MC.nCredEstado And CN.nConsCod = 3001),"
    sSql = sSql & "    cNroDoc = (Select cDocNro From MovDoc Where nMovNro = MC.nMovNro),"
    sSql = sSql & "    MC.nMonto as nMonto,"
    sSql = sSql & "    cUser = (Select RIGHT(cMovNro,4) From Mov Where nMovNro = MC.nMovNro),"
    sSql = sSql & "    dMov = (Select SUBSTRING(cMovNro,7,2) + '/' + SUBSTRING(cMovNro,5,2) + '/' + SUBSTRING(cMovNro,1,4) + ' ' + SUBSTRING(cMovNro,9,2) + ':' + SUBSTRING(cMovNro,11,2) + ':' + SUBSTRING(cMovNro,13,2) From Mov Where nMovNro = MC.nMovNro)"
    sSql = sSql & " From MovCol MC"
    sSql = sSql & "    Inner join Agencias A ON A.cAgeCod = SUBSTRING(MC.cCtaCod,4,2)"
    sSql = sSql & "    Inner Join ColocacConvenio CV ON CV.cCtaCod = MC.cCtaCod"
    sSql = sSql & "    Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSql = sSql & "    Inner Join Mov M ON M.nMovNro = MC.nMovNro"
    sSql = sSql & "    Inner Join MovDoc MD ON MD.nMovNro = M.nMovNro" & sCadBusq
    sSql = sSql & "    Inner Join ColocacCred CC ON CC.cCtaCod = MC.cCtaCod"
    sSql = sSql & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'"
    sSql = sSql & "    And SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdFecIni, "yyyymmdd") & "' AND '" & Format(pdFecFin, "yyyymmdd") & "'"
    sSql = sSql & "    And A.cAgecod in " & sCadAge & sCadProd
    sSql = sSql & "    And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & "  And SUBSTRING(MC.cCtaCod, 9, 1) = '" & pnMoneda & "' AND NOT MC.COPECOD LIKE '107%' "
    sSql = sSql & " Order by A.cAgeDescripcion, P.cPersNombre"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPagoConCheque = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaPagoDeOtrasAgencias(ByVal pnMoneda As Moneda, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, psCodAge As String, ByVal psCodCmact As String) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = " Select A.cAgeDescripcion as sAgencia, MC.cCtaCod, "
    sSql = sSql & " cOperacion = (Select O.cOpeDesc From OpeTpo O  Where O.cOpeCod = MC.cOpeCod),"
    sSql = sSql & " cNroDoc = (Select cDocNro From MovDoc Where nMovNro = MC.nMovNro),"
    sSql = sSql & " MC.nMonto as nMonto,"
    sSql = sSql & " cUser = (Select RIGHT(cMovNro,4) From Mov Where nMovNro = MC.nMovNro),"
    sSql = sSql & " dMov = (Select SUBSTRING(cMovNro,7,2) + '/' + SUBSTRING(cMovNro,5,2) + '/' + SUBSTRING(cMovNro,1,4) + ' ' + SUBSTRING(cMovNro,9,2) + ':' + SUBSTRING(cMovNro,11,2) + ':' + SUBSTRING(cMovNro,13,2) From Mov Where nMovNro = MC.nMovNro)"
    sSql = sSql & " From MovCol MC"
    sSql = sSql & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(MC.cCtaCod,4,2)"
    sSql = sSql & " Inner Join Mov M ON M.nMovNro = MC.nMovNro And SUBSTRING(M.cMovNro,15,3) = '" & psCodCmact & "'"
    sSql = sSql & " Left Join MovDoc MD ON MD.nMovNro = M.nMovNro"
    sSql = sSql & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'"
    sSql = sSql & "    And LEN(MC.cCtaCod)=18"
    sSql = sSql & "    And SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdFecIni, "yyyymmdd") & "' AND '" & Format(pdFecFin, "yyyymmdd") & "'"
    sSql = sSql & "    And SUBSTRING(MC.cCtaCod,4,2) <> '" & psCodAge & "' And SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' "
    sSql = sSql & "  And SUBSTRING(MC.cCtaCod, 9, 1) = '" & pnMoneda & "' "
    sSql = sSql & " AND MC.cOpeCod like '100[234567]%' "
    sSql = sSql & " Order by A.cAgeDescripcion, M.cMovNro"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPagoDeOtrasAgencias = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaPagoENOtrasAgencias(ByVal pnMoneda As Moneda, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, psCodAge As String, ByVal psCodCmact As String, _
    ByVal pMatCondicion As Variant, ByVal pMatProd As Variant) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadCondicion As String
Dim sCadProd As String
Dim i As Integer
    sCadProd = " And SUBSTRING(MC.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sCadCondicion = "("
    For i = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
    Next i
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"

    sSql = " Select A.cAgeDescripcion as sAgencia, MC.cCtaCod, P.cPersNombre,"
    sSql = sSql & " cOperacion = (Select O.cOpeDesc From OpeTpo O  Where O.cOpeCod = MC.cOpeCod),"
    sSql = sSql & " cNroDoc = (Select cDocNro From MovDoc Where nMovNro = MC.nMovNro),"
    sSql = sSql & " MC.nMonto as nMonto,"
    sSql = sSql & " cUser = (Select RIGHT(cMovNro,4) From Mov Where nMovNro = MC.nMovNro),"
    sSql = sSql & " dMov = (Select SUBSTRING(cMovNro,7,2) + '/' + SUBSTRING(cMovNro,5,2) + '/' + SUBSTRING(cMovNro,1,4) + ' ' + SUBSTRING(cMovNro,9,2) + ':' + SUBSTRING(cMovNro,11,2) + ':' + SUBSTRING(cMovNro,13,2) From Mov Where nMovNro = MC.nMovNro)"
    sSql = sSql & " From MovCol MC"
    sSql = sSql & " Inner Join Mov M ON M.nMovNro = MC.nMovNro And SUBSTRING(M.cMovNro,15,3) = '" & psCodCmact & "'"
    sSql = sSql & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(M.cMovNro,18,2) "
    sSql = sSql & " Inner join ProductoPersona PP ON PP.cCtaCod = MC.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & " Inner Join Persona P ON P.cPersCod = PP.cPersCod "
    sSql = sSql & " Left Join MovDoc MD ON MD.nMovNro = M.nMovNro"
    sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = MC.cCtaCod "
    sSql = sSql & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'"
    sSql = sSql & "    And LEN(MC.cCtaCod)=18"
    sSql = sSql & "    And SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdFecIni, "yyyymmdd") & "' AND '" & Format(pdFecFin, "yyyymmdd") & "'"
    sSql = sSql & "    And SUBSTRING(MC.cCtaCod,4,2) = '" & psCodAge & "' And SUBSTRING(M.cMovNro,18,2) <> '" & psCodAge & "' "
    sSql = sSql & "  And SUBSTRING(MC.cCtaCod, 9, 1) = '" & pnMoneda & "' " & sCadProd
    sSql = sSql & "    And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & " AND MC.cOpeCod like '100[234567]%' "
    sSql = sSql & " Order by A.cAgeDescripcion, M.cMovNro "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPagoENOtrasAgencias = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function


Public Function RecuperaInteresesEnSuspenso(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatProd As Variant)
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadProd As String

    sCadProd = " And SUBSTRING(MC.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
   
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    sSql = "SELECT "
    sSql = sSql & "   CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSql = sSql & "   END AS sAgencia,"
    sSql = sSql & "   CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')"
    sSql = sSql & "   END AS sProducto,"
    sSql = sSql & "   CASE WHEN (GROUPING(CN2.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(CN2.cConsDescripcion, 'PLAZO DESCONOCIDO ')"
    sSql = sSql & "   END AS sPlazo,"
    sSql = sSql & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
    sSql = sSql & "   END AS sFuenteFinan,"
    sSql = sSql & "   CASE WHEN (GROUPING(MC.cCtaCod) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(MC.cCtaCod, 'CREDITO DESCONOCIDO')"
    sSql = sSql & "   END AS cCtaCod,"
    sSql = sSql & " COUNT(*) as nCasos,"
    sSql = sSql & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = MC.cCtaCod),"
    sSql = sSql & " dRefinanciado = (Select CE.dPrdEstado From Producto CE Where CE.cCtaCod = MC.cCtaCod And CE.nPrdEstado = 2060),"
    sSql = sSql & " SUM(ISNULL(Cal.nMonto,0)) as nInteresComp,"
    sSql = sSql & " SUM(ISNULL(Cal2.nMonto,0)) as nInteresMor,"
    sSql = sSql & " SUM(IsNull(Cal3.nMonto, 0)) As nGastos"
    sSql = sSql & " FROM Producto MC Inner Join Constante CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSql = sSql & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(MC.cCtaCod,4,2)"
    sSql = sSql & " Inner Join Colocaciones CC ON CC.cCtaCod = MC.cCtaCod"
    sSql = sSql & " Inner Join ColocacCRed CCred ON CCred.cCtaCod = MC.cCtaCod"
    sSql = sSql & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = CC.cLineaCred"
    sSql = sSql & " Inner Join Persona P ON CL.cPersCod = P.cPersCod "
    sSql = sSql & " Inner join Constante CN2 ON CN2.nConsValor = Convert(int,SUBSTRING(CC.cLineaCred,6,1)) And CN2.nConsCod = " & gColocLineaCredPlazo
    sSql = sSql & " Inner join ColocacRefinanc CR ON CR.cCtaCodRef = MC.cCtaCod And CR.nEstado = " & gColocEstadoRefinancAprobado
    sSql = sSql & " Left Join ( Select CD.cCtaCodRef, SUM(isnull(CD.nMonto,0) - isnull(CD.nMontoPagado,0)) as nMonto From ColocacRefinancDet CD"
    sSql = sSql & "        Where (CD.nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & " Or CD.nPrdConceptoCod = "
    sSql = sSql & gColocConceptoCodInteresGracia & " Or CD.nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado
    sSql = sSql & " Or CD.nPrdConceptoCod = " & gColocConceptoCodInteresCompVencido & ") And nEstado = 3 "
    sSql = sSql & "        GROUP BY CD.cCtaCodRef) CAL ON CAL.cCtaCodRef = MC.cCtaCod"
    sSql = sSql & " Left Join ( Select CD.cCtaCodRef, SUM(CD.nMonto - isnull(CD.nMontoPagado,0)) as nMonto From ColocacRefinancDet CD "
    sSql = sSql & "        Where CD.nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & " And nEstado = 3 "
    sSql = sSql & "        GROUP BY CD.cCtaCodRef) CAL2 ON CAL2.cCtaCodRef = MC.cCtaCod"
    sSql = sSql & " Left Join ( Select CD.cCtaCodRef, SUM(CD.nMonto - isnull(CD.nMontoPagado,0)) as nMonto From ColocacRefinancDet CD "
    sSql = sSql & "        Where CONVERT(varchar(6),CD.nPrdConceptoCod) like '12%' And nEstado = 3 "
    sSql = sSql & "        GROUP BY CD.cCtaCodRef) CAL3 ON CAL3.cCtaCodRef = MC.cCtaCod "
    sSql = sSql & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'  and Len(MC.cCtaCod) = 18"
    sSql = sSql & " And MC.nPrdEstado = " & gColocEstRefinanc
    sSql = sSql & "    And A.cAgecod in " & sCadAge & sCadProd
    sSql = sSql & "  And SUBSTRING(MC.cCtaCod, 9, 1) = '" & pnMoneda & "' AND CCred.brefCapInt = 0"
    sSql = sSql & " GROUP BY A.cAgeDescripcion, CN.cConsDescripcion, CN2.cConsDescripcion, P.cPersNombre, MC.cCtaCod "
    sSql = sSql & " WITH ROLLUP "

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaInteresesEnSuspenso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
 
Public Function RecuperaCreditosRefinanciados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatProd As Variant)
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadProd As String

    
    sCadProd = " And SUBSTRING(MC.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
   
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    sSql = "SELECT DISTINCT "
    sSql = sSql & "   CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSql = sSql & "   END AS sAgencia,"
    sSql = sSql & "   CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')"
    sSql = sSql & "   END AS sProducto,"
    sSql = sSql & "   CASE WHEN (GROUPING(CN2.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(CN2.cConsDescripcion, 'PLAZO DESCONOCIDO ')"
    sSql = sSql & "   END AS sPlazo,"
    sSql = sSql & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
    sSql = sSql & "   END AS sFuenteFinan,"
    sSql = sSql & "   CASE WHEN (GROUPING(MC.cCtaCod) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(MC.cCtaCod, 'CREDITO DESCONOCIDO')"
    sSql = sSql & "   END AS cCtaCod,"
    sSql = sSql & " COUNT(*) as nCasos,"
    sSql = sSql & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = MC.cCtaCod),"
    sSql = sSql & " dRefinanciado = (Select CE.dVigencia From Colocaciones CE Where CE.cCtaCod = MC.cCtaCod),"
    sSql = sSql & " SUM(ISNULL(Cal.nMonto,0)) as nInteresComp,"
    sSql = sSql & " SUM(ISNULL(Cal2.nMonto,0)) as nInteresMor,"
    sSql = sSql & " SUM(IsNull(Cal3.nMonto, 0)) As nGastos, "
    sSql = sSql & " SUM(IsNull(Cal4.nMonto, 0)) As nCapital "
    sSql = sSql & " FROM Producto MC Inner Join Constante CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSql = sSql & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(MC.cCtaCod,4,2)"
    sSql = sSql & " Inner Join Colocaciones CC ON CC.cCtaCod = MC.cCtaCod"
    sSql = sSql & " Inner Join ColocacCred Cre ON Cre.cCtaCod = MC.cCtaCod"
    sSql = sSql & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = CC.cLineaCred"
    sSql = sSql & " Inner Join Persona P ON CL.cPersCod = P.cPersCod "
    sSql = sSql & " Inner join Constante CN2 ON CN2.nConsValor = Convert(int,SUBSTRING(CC.cLineaCred,6,1)) And CN2.nConsCod = " & gColocLineaCredPlazo
    sSql = sSql & " Inner join ColocacRefinanc CR ON CR.cCtaCod = MC.cCtaCod And CR.nEstado = " & gColocEstadoRefinancAprobado
    sSql = sSql & " Left Join ( Select CD.cCtaCod, SUM(isnull(CD.nMonto,0) - isnull(CD.nMontoPagado,0)) as nMonto From ColocacRefinancDet CD"
    sSql = sSql & "        Where ( CD.nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & " Or CD.nPrdConceptoCod = "
    sSql = sSql & gColocConceptoCodInteresGracia & " Or CD.nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado
    sSql = sSql & " Or CD.nPrdConceptoCod = " & gColocConceptoCodInteresCompVencido & ") AND nEstado=3 "
    sSql = sSql & "        GROUP BY CD.cCtaCod) CAL ON CAL.cCtaCod = MC.cCtaCod"
    sSql = sSql & " Left Join ( Select CD.cCtaCod, SUM(CD.nMonto - isnull(CD.nMontoPagado,0)) as nMonto From ColocacRefinancDet CD "
    sSql = sSql & "        Where (CD.nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & ") AND nEstado=3 "
    sSql = sSql & "        GROUP BY CD.cCtaCod) CAL2 ON CAL2.cCtaCod = MC.cCtaCod"
    sSql = sSql & " Left Join ( Select CD.cCtaCod, SUM(CD.nMonto - isnull(CD.nMontoPagado,0)) as nMonto From ColocacRefinancDet CD "
    sSql = sSql & "        Where (CONVERT(varchar(6),CD.nPrdConceptoCod) like '12%'" & " ) AND nEstado=3 "
    sSql = sSql & "        GROUP BY CD.cCtaCod) CAL3 ON CAL3.cCtaCod = MC.cCtaCod "
    sSql = sSql & " Left Join ( Select CD.cCtaCod, SUM(CD.nMonto - isnull(CD.nMontoPagado,0)) as nMonto From ColocacRefinancDet CD "
    sSql = sSql & "        Where ( CD.nPrdConceptoCod = " & gColocConceptoCodCapital & ") AND nEstado=3 "
    sSql = sSql & "        GROUP BY CD.cCtaCod) CAL4 ON CAL4.cCtaCod = MC.cCtaCod "
    sSql = sSql & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'  and Len(MC.cCtaCod) = 18"
    sSql = sSql & "    And A.cAgecod in " & sCadAge & sCadProd
    sSql = sSql & "  And SUBSTRING(MC.cCtaCod, 9, 1) = '" & pnMoneda & "' "
    sSql = sSql & " GROUP BY A.cAgeDescripcion, CN.cConsDescripcion, CN2.cConsDescripcion, P.cPersNombre, MC.cCtaCod "
    sSql = sSql & " WITH ROLLUP "
    
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosRefinanciados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
 
Public Function RecuperaImpLavDineroM(ByVal psFechaIni As String, ByVal psFechaFin As String, ByVal psAgencias As String, ByVal psMonedas As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    
    Dim sAgencias As String
    Dim sMonedas As String
    
    If Len(Trim(psAgencias)) = 0 Then
        sAgencias = ""
    Else
        sAgencias = " AND SUBSTRING(dbo.Colocaciones.cCtaCod, 1, 2) IN (" & psAgencias & ") "
    End If
    
    If Len(Trim(psMonedas)) = 0 Then
        sMonedas = ""
    Else
        sMonedas = " AND SUBSTRING(dbo.Colocaciones.cCtaCod, 6, 1) IN (" & psMonedas & ") "
    End If
    
    sSql = "SELECT     MovLavDinero.cPersCod, SUBSTRING(Colocaciones.cCtaCod, 6, 1) AS cMoneda, SUBSTRING(Colocaciones.cCtaCod, 1, 2) AS cAgencia, "
    sSql = sSql & " MovCol.cOpeCod, OpeTpo.cOpeDesc, MovCol.nMonto, Colocaciones.cCtaCod, Persona.cPersNombre, "
    sSql = sSql & " SUBSTRING(Colocaciones.cCtaCod, 6, 3) AS cCodProducto, Constante.cConsDescripcion AS cDesProducto, Agencias.cAgeDescripcion, "
    sSql = sSql & " RIGHT(Mov.cMovNro, 4) AS cUser, SUBSTRING(Mov.cMovNro, 7, 2) + '/' + SUBSTRING(Mov.cMovNro, 5, 2) "
    sSql = sSql & " + '/' + SUBSTRING(Mov.cMovNro, 1, 4) + ' ' + substring(Mov.cMovNro,9,2) + ':' + substring(Mov.cMovNro,11,2) + ':' + "
    sSql = sSql & " Substring(Mov.cMovNro,13,2) as cFechaHora, SUBSTRING(Mov.cMovNro, 1, 8) AS cFechaCompleta "
    sSql = sSql & " FROM         MovCol INNER JOIN "
    sSql = sSql & " MovLavDinero ON MovCol.nMovNro = MovLavDinero.nMovNro INNER JOIN "
    sSql = sSql & " OpeTpo ON MovCol.cOpeCod = OpeTpo.cOpeCod INNER JOIN "
    sSql = sSql & " Colocaciones ON MovCol.cCtaCod = Colocaciones.cCtaCod INNER JOIN "
    sSql = sSql & " Persona ON MovLavDinero.cPersCod = Persona.cPersCod INNER JOIN "
    sSql = sSql & " Constante ON SUBSTRING(Colocaciones.cCtaCod, 6, 3) = Constante.nConsValor INNER JOIN "
    sSql = sSql & " Agencias ON SUBSTRING(Colocaciones.cCtaCod, 1, 2) = Agencias.cAgeCod INNER JOIN "
    sSql = sSql & " Mov ON MovCol.nMovNro = Mov.nMovNro "
    sSql = sSql & " Where (Constante.nConsCod = 1001) " & sAgencias & sMonedas
    
    sSql = sSql & " and SUBSTRING(Mov.cMovNro, 1, 8)>='" & Format(psFechaIni, "YYYYMMdd") & "' and SUBSTRING(Mov.cMovNro, 1, 8) <='" & Format(psFechaFin, "YYYYMMdd") & "' "
    sSql = sSql & " ORDER BY Persona.cPersNombre, SUBSTRING(Colocaciones.cCtaCod, 1, 2), Agencias.cAgeDescripcion, Mov.cMovNro"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaImpLavDineroM = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaTCLavDineroM(ByVal pdFechaIni As String) As Double
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As New ADODB.Recordset
    
    sSql = "SELECT TOP 1 nValPond FROM TIPOCAMBIO WHERE datediff(d, dateadd(d," & Mid(pdFechaIni, 1, 2) * -1 & ", '" _
        & Format(pdFechaIni, "mm/dd/yyyy") & "'),  dFecCamb) = 0 ORDER BY dFecCamb "

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    Set rs = oConecta.CargaRecordSet(sSql)
    
    If rs.BOF Then
        RecuperaTCLavDineroM = 0
    Else
        RecuperaTCLavDineroM = rs!nValPond
    End If
    
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaListaCreditosComite(ByVal pdFecSis As Date, ByVal psCodAgencia As String) As Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
        
        sSql = "Select *"
        sSql = sSql & " From ("
        sSql = sSql & "Select P.cCtaCod,"
        sSql = sSql & " Titular=(Select cPersNombre From ProductoPersona PP"
        sSql = sSql & " Inner Join Persona Pers on PP.cPersCod=Pers.cPersCod and PP.nPrdPersRelac=20 and PP.cCtaCod=P.cCtaCod),"
        sSql = sSql & " TipoCredito=(Select cConsDescripcion From Constante Where nConsCod=1001 and nConsValor=Substring(P.cCtaCod,6,3)),"
        sSql = sSql & " Condicion=(Select cConsDescripcion From Constante Where nConsCod=3015 and nConsValor<>3015 and nConsValor=CC.nColocCondicion),"
        sSql = sSql & " CE.nMonto as MontoPropuesto,CE.nPlazo as PlazoPropuesto,"
        sSql = sSql & " CE1.nMonto as MontoAprobado,CE1.nPlazo as PlazoAprobado,"
        sSql = sSql & " CE.nCuotas as CuotasSugeridas,CE1.nCuotas as nCuotasAprobadas,"
        sSql = sSql & " CE.nPeriodoFechaFija as FechaFijaSugerido,CE1.nPeriodoFechaFija as FechaFijaAprobado,"
        sSql = sSql & " Analista=(Select RH.cUser From ProductoPersona PP"
        sSql = sSql & " Inner Join Persona Pers on PP.cPersCod=Pers.cPersCod and PP.nPrdPersRelac=28 and PP.cCtaCod=P.cCtaCod"
        sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=Pers.cPersCod),"
        sSql = sSql & " MontoGarantia=(Select Sum(nGravado) From ColocGarantia Where cCtaCod=P.cCtaCod)"
        sSql = sSql & " From Producto P"
        sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=P.cCtaCod"
        sSql = sSql & " Inner Join ColocacEstado CE on CE.cCtaCod=P.cCtaCod and CE.nPrdEstado=2001"
        sSql = sSql & " Inner Join ColocacEstado CE1 on CE1.cCtaCod=P.cCtaCod and CE1.nPrdEstado=2002"
        sSql = sSql & " Where "
        sSql = sSql & " Cast(datepart(yyyy,CE1.dPrdEstado) as varchar(4))+Right('00',2-len(Cast(datepart(mm,CE1.dPrdEstado) as varchar(2))))+Cast(datepart(mm,CE1.dPrdEstado) as varchar(2))+Right('00',2-len(Cast(datepart(dd,CE1.dPrdEstado) as varchar(2))))+Cast(datepart(dd,CE1.dPrdEstado) as varchar(2))='"
        sSql = sSql & DatePart("yyyy", pdFecSis) & Right("00", 2 - Len(DatePart("m", pdFecSis))) & DatePart("m", pdFecSis) & Right("00", 2 - Len(DatePart("d", pdFecSis))) & DatePart("d", pdFecSis) & "'"
        sSql = sSql & " And Substring(P.cCtaCod,4,2)='" & psCodAgencia & "'"
        sSql = sSql & " and CE.dPrdEstado=(Select Max(dPrdEstado) From ColocacEstado Where cCtaCod=P.cCtaCod and nPrdEstado=2001) "
        sSql = sSql & " and CE1.dPrdEstado=(Select Max(dPrdEstado) From ColocacEstado Where cCtaCod=P.cCtaCod and nPrdEstado=2002)"
        sSql = sSql & " ) X"
        sSql = sSql & " Order By Titular,cCtaCod"

    Set RecuperaListaCreditosComite = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    Set oConec = Nothing
End Function
Public Function RecuperaPersDevCredConv(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, ByVal MatInstitucion As Variant) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadInst As String

sCadAge = "('"
For i = 0 To UBound(pMatAgencias) - 1
    sCadAge = sCadAge & pMatAgencias(i) & "','"
Next i
sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
sCadInst = "('"
For i = 0 To UBound(MatInstitucion) - 1
    sCadInst = sCadInst & MatInstitucion(i) & "','"
Next i
sCadInst = Mid(sCadInst, 1, Len(sCadInst) - 2) & ")"

sSql = "         SELECT C.nMoneda, cMoneda  = (SELECT cConsDescripcion FROM CONSTANTE CC WHERE CC.nConsValor = C.nMoneda and CC.NCONSCOD =1011),"
sSql = sSql + "         C.cCodage, sAgencia = (SELECT cAgeDescripcion FROM AGENCIAS A WHERE A.cAgeCod = C.cCodAge ),"
sSql = sSql + "         C.cPersCodIns, P.cPersNombre AS cInstitucion,"
sSql = sSql + "         P1.cPersNombre as cTitular,"
sSql = sSql + "         C.dRegistro, C.cNrodoc, C.nMonto as nMonto,  "
sSql = sSql + "         cUser = ISNULL((SELECT RIGHT(cMovNro,4) FROM MOV M WHERE M.nMovNro = C.nMovNroReg  ),'') "
sSql = sSql + "  FROM   ColocacConvenioRegDev C"
sSql = sSql + "         JOIN PERSONA P ON C.cPersCodIns = P.cPersCod"
sSql = sSql + "         JOIN PERSONA P1 ON P1.cPersCod = C.cPersCod"
sSql = sSql + "  WHERE  (C.nPendiente IS NULL OR C.nPendiente<>'1') and C.nMoneda=" & pnMoneda & " And C.cCodAge IN " & sCadAge & " AND C.cPersCodIns in " & sCadInst
sSql = sSql + "         AND EXISTS (SELECT M.nMovNro FROM MOV M WHERE M.nMovNro = C.nMovNroReg and M.nMovflag=0 ) "
sSql = sSql + "         ORDER BY C.nMoneda, C.cCodage, C.cPersCodIns, P1.cPersNombre"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaPersDevCredConv = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function
Public Function RecuperaDevCredConv(ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, ByVal MatInstitucion As Variant) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadInst As String

sCadAge = "('"
For i = 0 To UBound(pMatAgencias) - 1
    sCadAge = sCadAge & pMatAgencias(i) & "','"
Next i
sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
sCadInst = "('"
For i = 0 To UBound(MatInstitucion) - 1
    sCadInst = sCadInst & MatInstitucion(i) & "','"
Next i
sCadInst = Mid(sCadInst, 1, Len(sCadInst) - 2) & ")"

sSql = "         Select  C.nMoneda, cMoneda  = (SELECT cConsDescripcion FROM CONSTANTE CC WHERE CC.nConsValor = C.nMoneda and CC.NCONSCOD =1011),"
sSql = sSql + "         C.cCodage, cAgencia = (SELECT cAgeDescripcion FROM AGENCIAS A WHERE A.cAgeCod = C.cCodAge ),"
sSql = sSql + "         C.cPersCodIns, P.cPersNombre AS cInstitucion,"
sSql = sSql + "         P1.cPersNombre as cTitular,"
sSql = sSql + "         SUBSTRING(M.CMOVNRO,7,2) + '/' + SUBSTRING(M.CMOVNRO,5,2) + '/' +SUBSTRING(M.CMOVNRO,1,4) AS dOperacion,"
sSql = sSql + "         CONVERT(CHAR(10),C.dRegistro,103) AS dRegistro, C.cNrodoc, C.nMonto as nMonto, RIGHT(M.cMovNro,4) as cUser"
sSql = sSql + "  from    ColocacConvenioRegDev C"
sSql = sSql + "         JOIN PERSONA P ON C.cPersCodIns = P.cPersCod"
sSql = sSql + "         JOIN PERSONA P1 ON P1.cPersCod = C.cPersCod"
sSql = sSql + "         JOIN MOV M ON M.NMOVNRO = C.NMOVNRO"
sSql = sSql + "  WHERE   C.nPendiente ='1' and M.nMovFlag=0 and C.nMoneda=" & pnMoneda & " And C.cCodAge IN " & sCadAge & " AND C.cPersCodIns in " & sCadInst
sSql = sSql + "         AND LEFT(M.cMovNro,8) BETWEEN '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "'"
sSql = sSql + "         ORDER BY C.nMoneda, C.cCodage, C.cPersCodIns, P1.cPersNombre"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaDevCredConv = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

Public Function ListaProtestosDia(ByVal psCodAge As String, ByVal pdFecha As Date) As Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim sDia As String
    Dim sMes As String
    Dim sAno As String
    Dim sFecha As String
    Dim oConec As COMConecta.DCOMConecta
    
    sDia = Right("00", 2 - Len(CStr(DatePart("d", pdFecha)))) & CStr(DatePart("d", pdFecha))
    sMes = Right("00", 2 - Len(CStr(DatePart("m", pdFecha)))) & CStr(DatePart("m", pdFecha))
    sAno = Right("0000", 4 - Len(CStr(DatePart("yyyy", pdFecha)))) & CStr(DatePart("yyyy", pdFecha))
    
    sFecha = sAno & sMes & sDia
    
    sSql = "Select P.cCtaCod, SUBSTRING(P.cCtaCod,4,2) as cCodAge, dbo.GetCtaAntigua(P.cCtaCod)  as cCtaCodAnt,Pers.cPersNombre,MC.nMonto,"
    sSql = sSql & " Moneda=Case Substring(P.cCtaCod,9,1)"
    sSql = sSql & " When '1' Then 'Soles'"
    sSql = sSql & " When '2' Then 'Dolares'"
    sSql = sSql & " End"
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=P.cCtaCod and PP.nPrdPersRelac=20"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Inner Join MovCol MC on MC.cCtaCod=P.cCtaCod and MC.cOpeCod='100900'"
    sSql = sSql & " Inner Join Mov M on M.nMovNro=MC.nMovNro"
    sSql = sSql & " Where Left(M.cMovNro,8)='" & sFecha & "' and SubString(P.cCtaCod,4,2)  in ('" & psCodAge & "')"
    sSql = sSql & " Order by SUBSTRING(P.cCtaCod,4,2), P.cCtaCod "
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaProtestosDia = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
End Function


Public Function ObtenerTelefonosTitular(ByVal psCtaCod As String) As Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    
    sSql = "Select Pers.cPersTelefono,Pers.cPersTelefono2"
    sSql = sSql & " From ProductoPersona PP"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Where PP.nPrdPersRelac=20 and PP.cCtaCod='" & psCtaCod & "'"
    
     Set oConec = New COMConecta.DCOMConecta
     oConec.AbreConexion
     Set ObtenerTelefonosTitular = oConec.CargaRecordSet(sSql)
     oConec.CierraConexion
End Function

Public Function ObtenerDireccTrabajoTitular(ByVal psCtaCod As String) As String
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sDireccion As String
    
     sSql = "Select PFI.cRazSocDirecc"
     sSql = sSql & " From ColocFteIngreso CFT"
     sSql = sSql & " Inner Join PersFteIngreso PFI on PFI.cNumFuente=CFT.cNumFuente"
     sSql = sSql & " Where CFT.cCtaCod='" & psCtaCod & "'"

     Set oConec = New COMConecta.DCOMConecta
     oConec.AbreConexion
     Set rs = oConec.CargaRecordSet(sSql)
     oConec.CierraConexion
     Set oConec = Nothing
     
     If Not rs.EOF And Not rs.BOF Then
        If Not IsNull(rs!cRazSocDirecc) Then
           sDireccion = rs!cRazSocDirecc
        End If
     End If
     
     Set rs = Nothing
     If sDireccion = "" Then
        sSql = "Select P.cPersDireccDomicilio"
        sSql = sSql & " from ProductoPersona PP"
        sSql = sSql & " Inner Join Persona P on PP.cPersCod=P.cPersCod"
        sSql = sSql & " Where PP.nPrdPersRelac=20 and PP.cCtaCod='" & psCtaCod & "'"
        
        Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        Set oConec = Nothing
        
        If Not rs.EOF And Not rs.BOF Then
            If Not IsNull(rs!cPersDireccDomicilio) Then
                sDireccion = rs!cPersDireccDomicilio
            End If
        End If
        Set rs = Nothing
     End If
     ObtenerDireccTrabajoTitular = sDireccion
End Function

Function ObtenerConvenio(ByVal psCtaCod As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select Pers.cPersNombre"
    sSql = sSql & " From ColocacConvenio CC"
    sSql = sSql & " Inner Join Persona Pers on CC.cPersCod=Pers.cPersCod"
    sSql = sSql & " Where cCtacod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs Is Nothing Then
        If Not rs.EOF And Not rs.BOF Then
           ObtenerConvenio = "Convenio Con :" & rs!cPersNombre
        Else
            ObtenerConvenio = ""
        End If
    Else
        ObtenerConvenio = ""
    End If
    Set rs = Nothing
End Function

Public Function ObtenerCarteraVigentePorUser(ByVal psUser As String, ByVal psCodAgen As String, _
ByVal psProducto As String, ByVal pbFecha As Boolean, Optional pdFecha As Date, _
Optional ByVal psProductoFiltro As String) As Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim nTipoCambio As Double
    'Dim odCredDoc As DCredDoc
    Dim rs As ADODB.Recordset
    
    
    'Set odCredDoc = New DCredDoc
    '    nTipoCambio = odCredDoc.ObtenerTipoCambio(pdFecha)
    'Set odCredDoc = Nothing
       nTipoCambio = ObtenerTipoCambio(pdFecha)
       
    If pbFecha = True Then
        sSql = "Select cUser,Sum(NumCreVig) as NumCreVig,Sum(SaldoDesem) as SaldoDesem,Sum(SaldoKVig) as SaldoKVig,Sum(SaldoPromVig) as SaldoPromVig"
        sSql = sSql & " From ("
        sSql = sSql & " Select RH.cUser,"
        sSql = sSql & " NumCreVig=(Select Count(*)"
        sSql = sSql & " From Producto P"
        sSql = sSql & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=28"
        sSql = sSql & " Inner join ColocacCred C on C.cCtaCod=PP.cCtaCod"
        sSql = sSql & " Where PP.cPersCod=RH.cPersCod and P.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205)"
        sSql = sSql & " and Substring(P.cCtaCod,4,2) in(" & psCodAgen & ") and (C.cRFA <>'RFA' or C.cRFA is null) and Substring(P.cCtaCod,7,2)<>'21'"
        If psProductoFiltro <> "" Then
           sSql = sSql & " and SubString(P.cCtaCod,6,3) in ('" & psProductoFiltro & "')), "
        Else
           sSql = sSql & " ), "
        End If
        sSql = sSql & " SaldoDesem=(Select Sum(Case When Substring(P.cCtaCod,9,1)='1' Then C.nMontoCol"
        sSql = sSql & " When Substring(P.cCtaCod,9,1)='2' Then C.nMontoCol*" & nTipoCambio & " End) as SaldoK"
        sSql = sSql & " From Producto P"
        sSql = sSql & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=28"
        sSql = sSql & " Inner Join Colocaciones C on C.cCtaCod=PP.cCtaCod"
        sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=C.cCtaCod"
        sSql = sSql & " Where P.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205) and Substring(P.cCtaCod,4,2) in(" & psCodAgen & ") and"
        sSql = sSql & " PP.cPersCod=RH.cPersCod and  (CC.cRFA <>'RFA' or CC.cRFA is null) and Substring(P.cCtaCod,7,2)<>'21'"
        If psProductoFiltro <> "" Then
           sSql = sSql & " and SubString(P.cCtaCod,6,3) In ('" & psProductoFiltro & "')),"
        Else
            sSql = sSql & " ), "
        End If
        sSql = sSql & " SaldokVig=(Select Sum(Case When Substring(P.cCtaCod,9,1)='1' Then P.nSaldo"
        sSql = sSql & " When Substring(P.cCtaCod,9,1)='2' Then P.nSaldo*" & nTipoCambio & " end)"
        sSql = sSql & " From Producto P"
        sSql = sSql & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=28"
        sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=PP.cCtaCod"
        sSql = sSql & " Where PP.cPersCod=RH.cPersCod and P.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205) and"
        sSql = sSql & " Substring(P.cCtaCod,4,2) in(" & psCodAgen & ") and (CC.cRFA <>'RFA' or CC.cRFA is null) and Substring(P.cCtaCod,7,2)<>'21'"
        If psProductoFiltro <> "" Then
           sSql = sSql & " and SubString(P.cCtaCod,6,3) in ('" & psProductoFiltro & "')),"
        Else
            sSql = sSql & " ), "
        End If
        sSql = sSql & " SaldoPromVig=  Round((Select Sum(P.nSaldo)"
        sSql = sSql & " From Producto P"
        sSql = sSql & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=28"
        sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=PP.cCtaCod"
        sSql = sSql & " Where PP.cPersCod=RH.cPersCod and P.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205))/"
        sSql = sSql & " (Select Count(*)"
        sSql = sSql & " From Producto P"
        sSql = sSql & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=28"
        sSql = sSql & " Inner join ColocacCred C on C.cCtaCod=PP.cCtaCod"
        sSql = sSql & " Where PP.cPersCod=RH.cPersCod and P.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205)),2)"
        sSql = sSql & " From  RRHH RH Where RH.cUser='" & psUser & "'"
        sSql = sSql & ")X"
        sSql = sSql & " GROUP BY cUser"
   Else
        sSql = "Select cUser,Sum(NumCreVig) as NumCreVig,Sum(SaldoDesem) as SaldoDesem,Sum(SaldoKVig) as SaldoKVig,Sum(SaldoPromVig) as SaldoPromVig"
        sSql = sSql & " From ("
        sSql = sSql & "Select RH.cUser,"
        sSql = sSql & " NumCreVig=(Select Count(*)"
        sSql = sSql & " From ColocacSaldo CS"
        sSql = sSql & " Inner Join ProductoPersona PP on CS.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=28"
        sSql = sSql & " Inner join ColocacCred C on C.cCtaCod=PP.cCtaCod"
        sSql = sSql & " Where PP.cPersCod=RH.cPersCod and CS.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205)"
        sSql = sSql & " and Substring(CS.cCtaCod,4,2) in(" & psCodAgen & ") and (C.cRFA <>'RFA' or C.cRFA is null) and Substring(CS.cCtaCod,7,2)<>'21'"
        sSql = sSql & " and CS.dFecha='" & CStr(Format(pdFecha, "MM/dd/yyyy")) & "'"
        If psProductoFiltro <> "" Then
            sSql = sSql & " and SubString(CS.cCtaCod,6,3) in ('" & psProductoFiltro & "')),"
        Else
            sSql = sSql & " ),"
        End If
        
        sSql = sSql & " SaldoDesem=(Select Sum(Case When Substring(CS.cCtaCod,9,1)='1' Then C.nMontoCol"
        sSql = sSql & " When Substring(CS.cCtaCod,9,1)='2' Then C.nMontoCol*" & nTipoCambio & " End) as SaldoK"
        sSql = sSql & " From ColocacSaldo CS"
        sSql = sSql & " Inner Join ProductoPersona PP on CS.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=28"
        sSql = sSql & " Inner Join Colocaciones C on C.cCtaCod=PP.cCtaCod"
        sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=C.cCtaCod"
        sSql = sSql & " Where CS.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205) and Substring(CS.cCtaCod,4,2) in(" & psCodAgen & ") and"
        sSql = sSql & " PP.cPersCod=RH.cPersCod and  (CC.cRFA <>'RFA' or CC.cRFA is null) and Substring(CS.cCtaCod,7,2)<>'21'"
        sSql = sSql & " and CS.dFecha='" & CStr(Format(pdFecha, "MM/dd/yyyy")) & "'"
        If psProductoFiltro <> "" Then
            sSql = sSql & " and  SubString(CS.cCtacod,6,3) in ('" & psProductoFiltro & "')),"
        Else
            sSql = sSql & " ),"
        End If
        
        sSql = sSql & " SaldokVig=(Select Sum(Case When Substring(CS.cCtaCod,9,1)='1' Then CS.nSaldoCap"
        sSql = sSql & " When Substring(CS.cCtaCod,9,1)='2' Then CS.nSaldoCap*" & nTipoCambio & " end)"
        sSql = sSql & " From ColocacSaldo CS"
        sSql = sSql & " Inner Join ProductoPersona PP on CS.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=28"
        sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=PP.cCtaCod"
        sSql = sSql & " Where PP.cPersCod=RH.cPersCod and CS.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205) and"
        sSql = sSql & " Substring(CS.cCtaCod,4,2) in(" & psCodAgen & ")  and (CC.cRFA <>'RFA' or CC.cRFA is null) and Substring(CS.cCtaCod,7,2)<>'21'"
        sSql = sSql & " and CS.dFecha='" & CStr(Format(pdFecha, "MM/dd/yyyy")) & "'"
        If psProductoFiltro <> "" Then
            sSql = sSql & " and SubString(CS.cCtaCod,6,3) in ('" & psProductoFiltro & "')),"
        
        Else
            sSql = sSql & "),"
        End If
        
        sSql = sSql & " SaldoPromVig=  Round((Select Sum(P.nSaldoCap)"
        sSql = sSql & " From ColocacSaldo P"
        sSql = sSql & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=28"
        sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=PP.cCtaCod"
        sSql = sSql & " Where PP.cPersCod=RH.cPersCod and P.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205))/"
        sSql = sSql & " (Select Count(*)"
        sSql = sSql & " From ColocacSaldo P"
        sSql = sSql & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=28"
        sSql = sSql & " Inner join ColocacCred C on C.cCtaCod=PP.cCtaCod"
        sSql = sSql & " Where PP.cPersCod=RH.cPersCod and P.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205)),2)"
        sSql = sSql & " From  RRHH RH Where RH.cUser='" & psUser & "'"
        sSql = sSql & " )X"
        sSql = sSql & " GROUP BY cUser"
   End If
  
    
    Set oConec = New COMConecta.DCOMConecta
    'Sleep 250
    oConec.AbreConexion
    Set ObtenerCarteraVigentePorUser = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    Exit Function
End Function

Public Function ListaConsolidadoPorAnalista(ByVal pnTipoCambio As Double, ByVal psAgenciaCod As String, ByVal pdFecha As Date, _
ByVal psAnalistas As String, ByVal psProducto As String, ByVal pbFecha As Boolean, Optional ByVal psProductoFiltro As String = "") As Recordset

    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sFecha As String
    Dim sDia As String
    Dim sMes As String
    Dim sAno As String
    Dim sFechaMov As String
    Dim sdInicioMes As String
    Dim nTipoCambio As Double
    'Dim odCredDoc As DCredDoc
    
    sFechaMov = Format(pdFecha, "MM/dd/yyyy")
    sAno = DatePart("yyyy", pdFecha)
    sMes = Right("00", 2 - Len(CStr(DatePart("m", pdFecha)))) & CStr(DatePart("m", pdFecha))
    sDia = Right("00", 2 - Len(CStr(DatePart("d", pdFecha)))) & CStr(DatePart("d", pdFecha))
    
    'Set odCredDoc = New DCredDoc
    nTipoCambio = ObtenerTipoCambio(pdFecha)
    
    sFecha = sAno & sMes & sDia
    sdInicioMes = sAno & sMes & "01"
    
    If pbFecha = True Then ' fecha del sistema
    
            sSql = "SELECT  CCODANA, SUM(nNumCredNew) as nNumCredNew, SUM(nMonCredNew) as nMonCredNew, "
            sSql = sSql & "  SUM(nNumCredRest) as nNumCredRest, SUM(nMonCredRest) as nMonCredRest ,"
            sSql = sSql & "  SUM(nNumCredDes) as nNumCredDes, SUM(nMonCredDes) as nMonCredDes,"
            sSql = sSql & " SUM(NSALDO) as NSALDO,"
            sSql = sSql & " SUM(NNRO1_15) as NNRO1_15, SUM(NSALDO1_15) as NSALDO1_15,"
            sSql = sSql & " SUM(NNRO16_30) as NNRO16_30, SUM(NSALDO16_30) as NSALDO16_30,"
            sSql = sSql & " SUM(NNROMAY31) as NNROMAY31, SUM(NSALDOMAY31) as NSALDOMAY31,"
            sSql = sSql & " SUM(NNROJUD) as NNROJUD, SUM(NSALDOJUD) as NSALDOJUD,"
            sSql = sSql & " SUM(NNROVENCOM+NNROVEN) AS NNROVEN, SUM(NSALDOVENCOM+NSALDOVEN) AS NSALDOVEN"
            sSql = sSql & " FROM CONSTANTE C JOIN"
            sSql = sSql & " (SELECT  CCTACOD, SUBSTRING(CCTACOD,6,3) AS NPRODUCTO, CCODANA, NPRDESTADO,isnull(nNumCredNew,0) as nNumCredNew, isnull(nMonCredNew,0) as nMonCredNew,"
            sSql = sSql & " isnull(nNumCredRest,0) as nNumCredRest, isnull(nMonCredRest,0) as nMonCredRest, isnull(nNumCredDes,0) as nNumCredDes,"
            sSql = sSql & " isnull(nMonCredDes,0) as nMonCredDes, CASE WHEN SUBSTRING(CCTACOD,9,1) ='1' THEN NSALDO ELSE Round((NSALDO*" & pnTipoCambio & "),2) End as NSALDO,"
            sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  1 AND 15 AND nPrdEstado NOT IN (2201,2205) THEN 1 ELSE 0 END  AS NNRO1_15,"
            sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  1 AND 15 AND nPrdEstado NOT IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDO1_15 ,"
            sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  16 AND 30 AND nPrdEstado NOT IN (2201,2205) THEN 1 ELSE 0 END AS NNRO16_30 ,"
            sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  16 AND 30 AND nPrdEstado NOT IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDO16_30 ,"
            sSql = sSql & " CASE WHEN nDiasAtr>=31 AND nPrdEstado NOT IN (2201,2205) THEN 1 ELSE 0 END AS NNROMAY31 ,"
            sSql = sSql & " CASE WHEN nDiasAtr>=31 AND nPrdEstado NOT IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDOMAY31 ,"
            sSql = sSql & " CASE WHEN nPrdEstado IN (2201,2205) THEN 1 ELSE 0 END AS NNROJUD ,"
            sSql = sSql & " CASE WHEN nPrdEstado IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDOJUD ,"
            sSql = sSql & " CASE WHEN nDiasAtr>= 16 and SUBSTRING(CCTACOD,6,1)='1' THEN 1 ELSE 0 END AS NNROVENCOM ,"
            sSql = sSql & " CASE WHEN nDiasAtr>= 16 and SUBSTRING(CCTACOD,6,1)='1' THEN NSALDO ELSE 0 END AS NSALDOVENCOM,"
            sSql = sSql & " CASE WHEN nDiasAtr>= 31 and SUBSTRING(CCTACOD,6,1)<>'1' THEN 1 ELSE 0 END AS NNROVEN ,"
            sSql = sSql & " CASE WHEN nDiasAtr>= 31 and SUBSTRING(CCTACOD,6,1)<>'1' THEN NSALDO ELSE 0 END AS NSALDOVEN"
            sSql = sSql & " From"
            sSql = sSql & " (SELECT P.CCTACOD, P.NPRDESTADO,"
            sSql = sSql & " CASE WHEN CANA.CUSER IS NULL AND SUBSTRING(P.CCTACOD,6,3)='305' THEN 'PREN' ELSE CANA.CUSER END AS CCODANA,"
            sSql = sSql & " nNumCredNew = CASE WHEN C.nColocCondicion = 1 THEN nNumCredNew ELSE 0 END,"
            sSql = sSql & " nMonCredNew = CASE WHEN C.nColocCondicion = 1 THEN nMonCredNew ELSE 0 END,"
            sSql = sSql & " nNumCredRest = CASE WHEN C.nColocCondicion <> 1 OR SUBSTRING(P.CCTACOD,6,3)='305' THEN nNumCredNew ELSE 0 END,"
            sSql = sSql & " nMonCredRest = CASE WHEN C.nColocCondicion <> 1 OR SUBSTRING(P.CCTACOD,6,3)='305' THEN nMonCredNew ELSE 0 END,"
            sSql = sSql & " NDESEM.nNumCredDes , NDESEM.nMonCredDes, NSALDO = CASE WHEN SUBSTRING(P.CCTACOD,9,1) ='1' THEN P.NSALDO ELSE ROUND(P.NSALDO*" & pnTipoCambio & ",2) END,"
            sSql = sSql & " nDiasAtr= CASE WHEN SUBSTRING(P.CCTACOD,6,3)='305' THEN DATEDIFF(DAY, CC.dVenc, '" & sFechaMov & "') ELSE C.nDiasAtraso END"
            sSql = sSql & " FROM    PRODUCTO P"
            sSql = sSql & " JOIN COLOCACIONES CC ON CC.CCTACOD = P.CCTACOD"
            sSql = sSql & " LEFT JOIN COLOCACCRED C ON C.CCTACOD = CC.CCTACOD"
            sSql = sSql & " LEFT JOIN ( SELECT  cCtaCod, COUNT(MC.cCtaCod) nNumCredNew ,"
            sSql = sSql & " SUM(CASE WHEN SUBSTRING(CCTACOD,9,1) ='1' THEN NMONTO ELSE ROUND((NMONTO* " & nTipoCambio & "),2) END) AS nMonCredNew"
            sSql = sSql & " FROM    MOV M"
            sSql = sSql & " JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO"
            sSql = sSql & " WHERE MC.COPECOD IN (120201, 100101,100102, 100103,100104,100105)"
            sSql = sSql & " AND LEFT(M.CMOVNRO,8) BETWEEN '" & sdInicioMes & "' AND '" & sFecha & "'  AND M.NMOVFLAG =0"
            sSql = sSql & " GROUP BY cCtaCod ) AS NCREDNEW ON NCREDNEW.CCTACOD = P.CCTACOD"
            sSql = sSql & " LEFT JOIN ( SELECT  cCtaCod, COUNT(MC.cCtaCod) nNumCredDes ,"
            sSql = sSql & " SUM(CASE WHEN SUBSTRING(CCTACOD,9,1) ='1' THEN NMONTO ELSE ROUND((NMONTO*" & nTipoCambio & "),2) END) AS nMonCredDes"
            sSql = sSql & " FROM    MOV M"
            sSql = sSql & " JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO"
            sSql = sSql & " WHERE   MC.COPECOD IN (120201, 100101,100102, 100103,100104,100105) AND M.NMOVFLAG =0"
            sSql = sSql & " GROUP BY cCtaCod ) AS NDESEM ON NDESEM.CCTACOD = P.CCTACOD"
            sSql = sSql & " LEFT JOIN ( SELECT  R.CCTACOD , RH.CUSER"
            sSql = sSql & " FROM    PRODUCTOPERSONA R"
            sSql = sSql & " JOIN RRHH RH ON RH.CPERSCOD = R.CPERSCOD"
            sSql = sSql & " WHERE   R.nPrdPersRelac = 28 ) AS CANA ON CANA.CCTACOD = P.CCTACOD"
            'sSQL = sSQL & " WHERE   P.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205) AND P.NSALDO>0) AS NCRED ) AS NCRED1"
            sSql = sSql & " WHERE   P.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205)) AS NCRED ) AS NCRED1"
            sSql = sSql & " ON NCRED1.NPRODUCTO = C.NCONSVALOR AND C.NCONSCOD = 1001"
            'sSQL = sSQL & " WHERE   Substring(NCRED1.cCtacod,4,2) in(" & psAgenciaCod & ") AND  CCODANA in ('JMMM','CEMG'," & psAnalistas & ") AND Substring(NCRED1.cCtacod,6,3) in (" & psProducto & ")"
            sSql = sSql & " WHERE   Substring(NCRED1.cCtacod,4,2) in(" & psAgenciaCod & ") AND NOT EXISTS (SELECT CCTACOD FROM COLOCACCRED WHERE CCTACOD=NCRED1.CCTACOD AND CRFA IN ('RFA'))"
            If psProductoFiltro <> "" Then
                sSql = sSql & " AND SubString(NCRED1.cCtaCod,6,3) in('" & psProducto & "')"
            End If
            sSql = sSql & " GROUP BY CCODANA"
            sSql = sSql & " ORDER BY  CCODANA"
    Else
        sSql = "SELECT  CCODANA,"
        sSql = sSql & " SUM(nNumCredNew) as nNumCredNew, SUM(nMonCredNew) as nMonCredNew,"
        sSql = sSql & " SUM(nNumCredRest) as nNumCredRest, SUM(nMonCredRest) as nMonCredRest ,"
        sSql = sSql & " SUM(nNumCredDes) as nNumCredDes, SUM(nMonCredDes) as nMonCredDes, SUM(NSALDO) as NSALDO, SUM(NNRO1_15) as NNRO1_15,"
        sSql = sSql & " SUM(NSALDO1_15) as NSALDO1_15, SUM(NNRO16_30) as NNRO16_30, SUM(NSALDO16_30) as NSALDO16_30, SUM(NNROMAY31) as NNROMAY31,"
        sSql = sSql & " SUM(NSALDOMAY31) as NSALDOMAY31, SUM(NNROJUD) as NNROJUD, SUM(NSALDOJUD) as NSALDOJUD,"
        sSql = sSql & " SUM(NNROVENCOM+NNROVEN) AS NNROVEN, SUM(NSALDOVENCOM+NSALDOVEN) AS NSALDOVEN"
        sSql = sSql & " FROM CONSTANTE C"
        sSql = sSql & " JOIN (SELECT  CCTACOD, SUBSTRING(CCTACOD,6,3) AS NPRODUCTO, CCODANA, NPRDESTADO,isnull(nNumCredNew,0) as nNumCredNew,"
        sSql = sSql & " isnull(nMonCredNew,0) as nMonCredNew, isnull(nNumCredRest,0) as nNumCredRest, isnull(nMonCredRest,0) as nMonCredRest,"
        sSql = sSql & " isnull(nNumCredDes,0) as nNumCredDes, isnull(nMonCredDes,0) as nMonCredDes,"
        sSql = sSql & " CASE WHEN SUBSTRING(CCTACOD,9,1) ='1' THEN NSALDO ELSE Round((NSALDO*" & nTipoCambio & "),2) End as NSALDO,"
        sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  1 AND 15 AND nPrdEstado NOT IN (2201,2205) THEN 1 ELSE 0 END  AS NNRO1_15,"
        sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  1 AND 15 AND nPrdEstado NOT IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDO1_15 ,"
        sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  16 AND 30 AND nPrdEstado NOT IN (2201,2205) THEN 1 ELSE 0 END AS NNRO16_30 ,"
        sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  16 AND 30 AND nPrdEstado NOT IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDO16_30 ,"
        sSql = sSql & " CASE WHEN nDiasAtr>=31 AND nPrdEstado NOT IN (2201,2205) THEN 1 ELSE 0 END AS NNROMAY31 ,"
        sSql = sSql & " CASE WHEN nDiasAtr>=31 AND nPrdEstado NOT IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDOMAY31 ,"
        sSql = sSql & " CASE WHEN nPrdEstado IN (2201,2205) THEN 1 ELSE 0 END AS NNROJUD ,"
        sSql = sSql & " CASE WHEN nPrdEstado IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDOJUD ,"
        sSql = sSql & " CASE WHEN nDiasAtr>= 16 and SUBSTRING(CCTACOD,6,1)='1' THEN 1 ELSE 0 END AS NNROVENCOM ,"
        sSql = sSql & " CASE WHEN nDiasAtr>= 16 and SUBSTRING(CCTACOD,6,1)='1' THEN NSALDO ELSE 0 END AS NSALDOVENCOM,"
        sSql = sSql & " CASE WHEN nDiasAtr>= 31 and SUBSTRING(CCTACOD,6,1)<>'1' THEN 1 ELSE 0 END AS NNROVEN ,"
        sSql = sSql & " CASE WHEN nDiasAtr>= 31 and SUBSTRING(CCTACOD,6,1)<>'1' THEN NSALDO ELSE 0 END AS NSALDOVEN"
        sSql = sSql & " FROM (SELECT CS.CCTACOD, CS.NPRDESTADO,"
        sSql = sSql & " CASE WHEN CANA.CUSER IS NULL AND SUBSTRING(CS.CCTACOD,6,3)='305' THEN 'PREN' ELSE CANA.CUSER END AS CCODANA,"
        sSql = sSql & " nNumCredNew = CASE WHEN C.nColocCondicion = 1 THEN nNumCredNew ELSE 0 END,"
        sSql = sSql & " nMonCredNew = CASE WHEN C.nColocCondicion = 1 THEN nMonCredNew ELSE 0 END,"
        sSql = sSql & " nNumCredRest = CASE WHEN C.nColocCondicion <> 1 OR SUBSTRING(CS.CCTACOD,6,3)='305' THEN nNumCredNew ELSE 0 END,"
        sSql = sSql & " nMonCredRest = CASE WHEN C.nColocCondicion <> 1 OR SUBSTRING(CS.CCTACOD,6,3)='305' THEN nMonCredNew ELSE 0 END,"
        sSql = sSql & " NDESEM.nNumCredDes , NDESEM.nMonCredDes, NSALDO = CASE WHEN SUBSTRING(CS.CCTACOD,9,1) ='1' THEN CS.NSALDOCAP ELSE ROUND(CS.NSALDOCAP*" & nTipoCambio & ",2) END,"
        sSql = sSql & " nDiasAtr = CS.nDiasAtraso"
        sSql = sSql & " FROM    COLOCACSALDO CS"
        sSql = sSql & " JOIN COLOCACIONES CC ON CC.CCTACOD = CS.CCTACOD"
        sSql = sSql & " LEFT JOIN COLOCACCRED C ON C.CCTACOD = CC.CCTACOD"
        sSql = sSql & " LEFT JOIN ( SELECT  cCtaCod, COUNT(MC.cCtaCod) nNumCredNew ,"
        sSql = sSql & " SUM(CASE WHEN SUBSTRING(CCTACOD,9,1) ='1' THEN NMONTO ELSE ROUND((NMONTO*" & nTipoCambio & "),2) END) AS nMonCredNew"
        sSql = sSql & " FROM    MOV M JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO"
        sSql = sSql & " WHERE MC.COPECOD IN (120201, 100101,100102, 100103,100104,100105) AND"
        sSql = sSql & " LEFT(M.CMOVNRO,8) BETWEEN '" & sdInicioMes & "' AND '" & sFecha & "'  AND M.NMOVFLAG =0"
        sSql = sSql & " GROUP BY cCtaCod ) AS NCREDNEW ON NCREDNEW.CCTACOD = CS.CCTACOD"
        sSql = sSql & " LEFT JOIN ( SELECT  cCtaCod, COUNT(MC.cCtaCod) nNumCredDes ,"
        sSql = sSql & " SUM(CASE WHEN SUBSTRING(CCTACOD,9,1) ='1' THEN NMONTO ELSE ROUND((NMONTO*" & nTipoCambio & "),2) END) AS nMonCredDes"
        sSql = sSql & " FROM    MOV M"
        sSql = sSql & " JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO"
        sSql = sSql & " WHERE   MC.COPECOD IN (120201, 100101,100102, 100103,100104,100105) AND M.NMOVFLAG =0 GROUP BY cCtaCod ) AS NDESEM ON NDESEM.CCTACOD = CS.CCTACOD"
        sSql = sSql & " LEFT JOIN ( SELECT  R.CCTACOD , RH.CUSER"
        sSql = sSql & " FROM    PRODUCTOPERSONA R JOIN RRHH RH ON RH.CPERSCOD = R.CPERSCOD WHERE   R.nPrdPersRelac = 28 ) AS CANA ON CANA.CCTACOD = CS.CCTACOD"
        'sSQL = sSQL & " WHERE   CS.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205) AND CS.NSALDOCAP>0 AND CS.DFECHA=' " & CStr(Format(pdFecha, "mm/dd/yyyy")) & " ' ) AS NCRED ) AS NCRED1 ON NCRED1.NPRODUCTO = C.NCONSVALOR AND C.NCONSCOD = 1001"
        sSql = sSql & " WHERE   CS.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205)AND CS.DFECHA=' " & CStr(Format(pdFecha, "mm/dd/yyyy")) & " ' ) AS NCRED ) AS NCRED1 ON NCRED1.NPRODUCTO = C.NCONSVALOR AND C.NCONSCOD = 1001"
        sSql = sSql & " WHERE   Substring(NCRED1.cCtacod,4,2) in(" & psAgenciaCod & ") AND NOT EXISTS (SELECT CCTACOD FROM COLOCACCRED WHERE CCTACOD=NCRED1.CCTACOD AND CRFA IN ('RFA'))"
        If psProductoFiltro <> "" Then
                sSql = sSql & " AND SubString(NCRED1.cCtaCod,6,3) in('" & psProducto & "')"
            End If
        sSql = sSql & " GROUP BY CCODANA ORDER BY  CCODANA"

    End If
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
        Set ListaConsolidadoPorAnalista = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaUser(ByVal psPersCod As String) As String
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    
    Set oConec = New COMConecta.DCOMConecta
    sSql = "Select cUser From RRHH Where cPersCod='" & psPersCod & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        ListaUser = rs!cUser
    End If
    Set rs = Nothing
End Function

Public Sub InsertarTmpConsolidadoAnalistas(ByVal rsAnalista As ADODB.Recordset, _
ByVal psCodUser As String, ByVal pdFecha As Date, ByVal psProducto As String, _
ByVal psCodAgen As String, ByVal pbFecha As Boolean)
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim nNumCredVig As Integer
    Dim nKDesemVig As Double
    Dim nSaldokVig As Double
    Dim nSaldoPromVig As Double
    Dim rs As ADODB.Recordset
    Dim nTipoCambio As Double
    Dim nCodigo As Integer
    'Dim odCredDoc As DCredDoc
    Dim sAgenciaNew As String
    Dim nAgencia As Integer
    Dim i As Integer
''    On Error GoTo ErrHandler
    i = 1
    If Len(psCodAgen) > 4 Then
        nAgencia = Len(psCodAgen)
        Do While i < nAgencia
            If IsNumeric(Mid(psCodAgen, i, 1)) Then
                sAgenciaNew = sAgenciaNew & Mid(psCodAgen, i, 1)
            End If
            i = i + 1
        Loop
        sAgenciaNew = "'" & sAgenciaNew & "'"
    End If
    
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "Select Isnull(Max(IdCodigo),0) as IdCodigo From TmpConsolidadoAnalista"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    If Not rs.EOF And Not rs.BOF Then
        nCodigo = 1 + rs!IdCodigo
    End If
    
    Set rs = Nothing
    oConec.AbreConexion
    If rsAnalista.RecordCount > 0 Then
        'oConec.ConexionActiva.BeginTrans
    End If
    
    If pbFecha = True Then
        'Set odCredDoc = New DCredDoc
        nTipoCambio = ObtenerTipoCambio(pdFecha)
        'Set odCredDoc = Nothing
        If Len(psCodAgen) > 4 Then
            sSql = " InsertarTempSaldoVigConsolAnalista " & sAgenciaNew & "," & nTipoCambio & "," & nCodigo
        Else
            sSql = " InsertarTempSaldoVigConsolAnalista " & psCodAgen & "," & nTipoCambio & "," & nCodigo
        End If
        
        oConec.ConexionActiva.CommandTimeout = 12000
        oConec.Ejecutar sSql
   End If
    
    Do Until rsAnalista.EOF
     'If rsAnalista!CCODANA <> "NULL" Then
            If rsAnalista!cCodAna = "PREN" Then
               Set rs = ListaPrendarioConsolidadAnalista(psCodAgen, pdFecha, pbFecha)
               If Not rs.EOF And Not rs.BOF Then
                    nNumCredVig = IIf(IsNull(rs!cantidad), 0, rs!cantidad)
                    nKDesemVig = IIf(IsNull(rs!Desembolsado), 0, rs!Desembolsado)
                    nSaldokVig = IIf(IsNull(rs!SaldosK), 0, rs!SaldosK)
                    nSaldoPromVig = IIf(IsNull(rs!SALDOSKPROM), 0, rs!SALDOSKPROM)
                End If
                Set rs = Nothing
            End If
            
    If pbFecha = False Then
            If rsAnalista!cCodAna <> "PREN" Then
                Set rs = ObtenerCarteraVigentePorUser(rsAnalista!cCodAna, psCodAgen, psProducto, pbFecha, pdFecha)
                ' Set rs = ObtenerCarteraVigentePorUserStored(psCodAgen, rsAnalista!cCodAna, pdFecha)
                If Not rs.EOF And Not rs.BOF Then
                    nNumCredVig = IIf(IsNull(rs!NumCreVig), 0, rs!NumCreVig)
                    nKDesemVig = IIf(IsNull(rs!SaldoDesem), 0, rs!SaldoDesem)
                    nSaldokVig = IIf(IsNull(rs!SaldoKVig), 0, rs!SaldoKVig)
                    nSaldoPromVig = IIf(IsNull(rs!SaldoPromVig), 0, rs!SaldoPromVig)
                End If
                Set rs = Nothing
             End If
   Else
              If rsAnalista!cCodAna <> "PREN" Then
                sSql = "Select *"
                sSql = sSql & " From TempSaldoVigConsolAnalista"
                sSql = sSql & " Where IdCodigo=" & nCodigo & " and cUser='" & rsAnalista!cCodAna & "'"
                
                Set rs = oConec.CargaRecordSet(sSql)
                   If Not rs.EOF And Not rs.BOF Then
                      nNumCredVig = IIf(IsNull(rs!NumCreVig), 0, rs!NumCreVig)
                      nKDesemVig = IIf(IsNull(rs!SaldoDesem), 0, rs!SaldoDesem)
                      nSaldokVig = IIf(IsNull(rs!SaldoKVig), 0, rs!SaldoKVig)
                      nSaldoPromVig = IIf(IsNull(rs!SaldoPromVig), 0, rs!SaldoPromVig)
                  End If
                  Set rs = Nothing
             End If
   End If
            If nNumCredVig > 0 Then
                sSql = "Insert Into TmpConsolidadoAnalista Values('" & Format(pdFecha, "MM/dd/yyyy") & "','" & psCodUser & "',"
                sSql = sSql & "'" & rsAnalista!cCodAna & "'," & nNumCredVig & "," & Format(nKDesemVig, "#0.00") & "," & Format(nSaldokVig, "#0.00") & ","
                sSql = sSql & Format(nSaldoPromVig, "#0.00") & "," & rsAnalista!nNumCredNew & "," & Format(rsAnalista!nMonCredNew, "#0.00") & "," & rsAnalista!nNumCredRest & ","
                sSql = sSql & Format(rsAnalista!nMonCredRest, "#0.00") & "," & rsAnalista!NNRO1_15 & "," & Format(rsAnalista!NSALDO1_15, "#0.00") & "," & Format((rsAnalista!NNRO1_15 / nNumCredVig) * 100, "#0.00") & "," & rsAnalista!NNRO16_30 & ","
                sSql = sSql & Format(rsAnalista!NSALDO16_30, "#0.00") & "," & Format((rsAnalista!NNRO16_30 / nNumCredVig) * 100, "#0.00") & "," & rsAnalista!NNROMAY31 & "," & Format(rsAnalista!NSALDOMAY31, "#0.00") & "," & Format((rsAnalista!NNROMAY31 / nNumCredVig) * 100, "#0.00") & ","
                sSql = sSql & rsAnalista!NNROJUD & "," & Format(rsAnalista!NSALDOJUD, "#0.00") & "," & Format((rsAnalista!NNROJUD / nNumCredVig) * 100, "#0.00") & "," & rsAnalista!NNROVEN & "," & Format(rsAnalista!NSALDOVEN, "#0.00") & ","
                sSql = sSql & Format((rsAnalista!NNROVEN / nNumCredVig) * 100, "#0.00") & ","
                sSql = sSql & nCodigo & ")"
                oConec.ConexionActiva.Execute sSql
         '   End If
        'Else
         '   sSQL = "Insert Into TmpConsolidadoAnalista Values('" & Format(pdFecha, "MM/dd/yyyy") & "','" & psCodUser & "',"
           ' sSQL = sSQL & "'" & rsAnalista!CCODANA & "'," & nNumCredVig & "," & Format(nKDesemVig, "#0.00") & "," & Format(nSaldoKVig, "#0.00") & ","
          '  sSQL = sSQL & Format(nSaldoPromVig, "#0.00") & "," & rsAnalista!nNumCredNew & "," & Format(rsAnalista!nMonCredNew, "#0.00") & "," & rsAnalista!nNumCredRest & ","
            'sSQL = sSQL & Format(rsAnalista!nMonCredRest, "#0.00") & "," & rsAnalista!NNRO1_15 & "," & Format(rsAnalista!NSALDO1_15, "#0.00") & "," & Format((rsAnalista!NNRO1_15 / nNumCredVig) * 100, "#0.00") & "," & rsAnalista!NNRO16_30 & ","
            'sSQL = sSQL & Format(rsAnalista!NSALDO16_30, "#0.00") & "," & Format((rsAnalista!NNRO16_30 / nNumCredVig) * 100, "#0.00") & "," & rsAnalista!NNROMAY31 & "," & Format(rsAnalista!NSALDOMAY31, "#0.00") & "," & Format((rsAnalista!NNROMAY31 / nNumCredVig) * 100, "#0.00") & ","
           ' sSQL = sSQL & rsAnalista!NNROJUD & "," & Format(rsAnalista!NSALDOJUD, "#0.00") & "," & Format((rsAnalista!NNROJUD / nNumCredVig) * 100, "#0.00") & "," & rsAnalista!NNROVEN & "," & Format(rsAnalista!NSALDOVEN, "#0.00") & ","
            'sSQL = sSQL & Format((rsAnalista!NNROVEN / nNumCredVig) * 100, "#0.00") & ","
            'sSQL = sSQL & nCodigo & ")"
           
        End If
        nNumCredVig = 0
        nKDesemVig = 0
        nSaldokVig = 0
        nSaldoPromVig = 0
        rsAnalista.MoveNext
    Loop
    Set rsAnalista = Nothing
   ' oConec.ConexionActiva.CommitTrans
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Sub
''ErrHandler:
  ''  If Not oConec Is Nothing Then
    ''    If oConec.ConexionActiva.State = adStateOpen Then
      ''      oConec.ConexionActiva.RollbackTrans
        ''End If
    ''End '''' If
End Sub
Public Sub InsertarTmpConsolidadoProducto(ByVal rsProducto As ADODB.Recordset, _
ByVal psCodUser As String, ByVal pdFecha As Date, ByVal psProducto As String, _
ByVal psCodAgen As String, ByVal pbFecha As Boolean, ByVal psProductoFiltro As String, _
ByVal pnCodigo As Integer)
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim nNumCredVig As Integer
    Dim nKDesemVig As Double
    Dim nSaldokVig As Double
    Dim nSaldoPromVig As Double
    Dim rs As ADODB.Recordset
    Dim nCodigo As Integer
    
'    Set oConec = New COMConecta.DCOMConecta
'    oConec.AbreConexion
'    sSQL = "Select IsNull(Max(IdCodigo),0) as IdCodigo From TmpConsolidadoProducto"
'
'    Set rs = oConec.CargaRecordSet(sSQL)
'    oConec.CierraConexion
'
'    If Not rs.EOF And Not rs.BOF Then
'        nCodigo = rs!IdCodigo + 1
'    End If
'    Set rs = Nothing
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Do Until rsProducto.EOF
        If rsProducto!cCodAna = "PREN" Then
            Set rs = ListaPrendarioConsolidadAnalista(psCodAgen, pdFecha, pbFecha)
            If Not rs.EOF And Not rs.BOF Then
                nNumCredVig = IIf(IsNull(rs!cantidad), 0, rs!cantidad)
                nKDesemVig = IIf(IsNull(rs!Desembolsado), 0, rs!Desembolsado)
                nSaldokVig = IIf(IsNull(rs!SaldosK), 0, rs!SaldosK)
            End If
            Set rs = Nothing
        End If
        If rsProducto!cCodAna <> "PREN" Then
            Set rs = ObtenerCarteraVigentePorUser(rsProducto!cCodAna, psCodAgen, "", pbFecha, pdFecha, psProductoFiltro)
            If Not rs.EOF And Not rs.BOF Then
                nNumCredVig = IIf(IsNull(rs!NumCreVig), 0, rs!NumCreVig)
                nKDesemVig = IIf(IsNull(rs!SaldoDesem), 0, rs!SaldoDesem)
                nSaldokVig = IIf(IsNull(rs!SaldoKVig), 0, rs!SaldoKVig)
                nSaldoPromVig = IIf(IsNull(rs!SaldoPromVig), 0, rs!SaldoPromVig)
            End If
            Set rs = Nothing
       End If
       If nNumCredVig > 0 Then
            sSql = "Insert Into TmpConsolidadoProducto Values('" & Format(pdFecha, "MM/dd/yyyy") & "','" & psCodUser & "',"
            sSql = sSql & "'" & rsProducto!cCodAna & "'," & nNumCredVig & "," & Format(nKDesemVig, "#0.00") & "," & Format(nSaldokVig, "#0.00") & ","
            sSql = sSql & Format(nSaldoPromVig, "#0.00") & "," & rsProducto!nNumCredNew & "," & Format(rsProducto!nMonCredNew, "#0.00") & "," & rsProducto!nNumCredRest & ","
            sSql = sSql & Format(rsProducto!nMonCredRest, "#0.00") & "," & rsProducto!NNRO1_15 & "," & Format(rsProducto!NSALDO1_15, "#0.00") & "," & Format((rsProducto!NNRO1_15 / nNumCredVig) * 100, "#0.00") & "," & rsProducto!NNRO16_30 & ","
            sSql = sSql & Format(rsProducto!NSALDO16_30, "#0.00") & "," & Format((rsProducto!NNRO16_30 / nNumCredVig) * 100, "#0.00") & "," & rsProducto!NNROMAY31 & "," & Format(rsProducto!NSALDOMAY31, "#0.00") & "," & Format((rsProducto!NNROMAY31 / nNumCredVig) * 100, "#0.00") & ","
            sSql = sSql & rsProducto!NNROJUD & "," & Format(rsProducto!NSALDOJUD, "#0.00") & "," & Format((rsProducto!NNROJUD / nNumCredVig) * 100, "#0.00") & "," & rsProducto!NNROVEN & "," & Format(rsProducto!NSALDOVEN, "#0.00") & ","
            sSql = sSql & Format((rsProducto!NNROVEN / nNumCredVig) * 100, "#0.00") & ","
            sSql = sSql & "'" & psProductoFiltro & "'," & pnCodigo & ")"
            oConec.ConexionActiva.Execute sSql
       End If
        nNumCredVig = 0
        nKDesemVig = 0
        nSaldokVig = 0
        nSaldoPromVig = 0
        rsProducto.MoveNext
    Loop
    Set rsProducto = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
Public Function ObtenerIdCodigo(ByVal pdFecha As Date, psCodUser As String) As Integer
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "Select Max(IdCodigo) as nCodigo From TmpConsolidadoAnalista Where FechaProceso='" & Format(pdFecha, "MM/dd/yyyy") & "' AND Usuario='" & psCodUser & "'"
    
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not rs.EOF And Not rs.BOF Then
        ObtenerIdCodigo = rs!nCodigo
    End If
    Set rs = Nothing
End Function

Public Function ObtenerIDCodigoProductoAnalista(ByVal pdFecha As Date, psCodUser As String) As Integer
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "Select Max(IdCodigo) nCodigo From TmpConsolidadoProducto Where FechaProceso='" & Format(pdFecha, "MM/dd/yyyy") & "' And Usuario='" & psCodUser & "'"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        ObtenerIDCodigoProductoAnalista = rs!nCodigo
    End If
    Set rs = Nothing
End Function

Public Function ListaConsolidadoRefinac(ByVal psCodUser As String, ByVal psAgencia As String, _
ByVal psProducto As String) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    
    sSql = "Select Count(*) as nRefinanciado,"
    sSql = sSql & " Sum(nSaldo) As SKRefinanciado"
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=P.cCtaCod and PP.nPrdPersRelac=28"
    sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
    sSql = sSql & " Where P.nPrdEstado in (2030,2031,2032) and RH.cUser='RAOC' and Substring(P.cCtaCod,4,2)='01' and"
    sSql = sSql & "Substring(P.cCtaCod,6,3)=''"

   Set ListaConsolidadoRefinac = oConec.CargaRecordSet(sSql)
   oConec.CierraConexion
   Set oConec = Nothing
End Function


Public Function ListaMoraXAnalistaCabecera(ByVal psUser As String, ByVal psProductos As String, _
ByVal psAgencias As String, ByVal psCondicion As String, ByVal psMoneda As String, ByVal pdFecha As String, _
ByVal pnIniDiaMora As Integer, ByVal pnFinDiaMora As Integer, ByVal pnMontoIni As Double, ByVal pnMontoFin As Double, _
Optional ByVal psCodUbiGeo As String) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "Select * From ("
    sSql = sSql & " Select  RH.cUser,P.cCtaCod,"
    sSql = sSql & " Titular=(Select Pers.cPersNombre"
    sSql = sSql & " From ProductoPersona PP1"
    sSql = sSql & " Inner Join Persona Pers on PP1.cPersCod=Pers.cPersCod"
    sSql = sSql & " Where PP1.cCtaCod=P.cCtaCod and PP1.nPrdPersRelac=20),"
    sSql = sSql & " Direccion=(Select Pers.cPersDireccDomicilio"
    sSql = sSql & " From ProductoPersona PP1 "
    sSql = sSql & " Inner Join Persona Pers on PP1.cPersCod=Pers.cPersCod"
    sSql = sSql & " Where PP1.cCtaCod=P.cCtaCod and PP1.nPrdPersRelac=20),"
    sSql = sSql & " KDesembolsado=(Select nMontoCol"
    sSql = sSql & " From Colocaciones"
    sSql = sSql & " Where cCtaCod=P.cCtaCod),"
    sSql = sSql & " P.nSaldo as SaldoK,"
    sSql = sSql & " NroCuotaApr=(Select Count(*)"
    sSql = sSql & " From ColocCalendario CD"
    sSql = sSql & " Where CD.cCtaCod=P.cCtaCod and CD.nColocCalendApl=1 and CD.nNroCalen=(Select Max(nNroCalen) From ColocCalendario"
    sSql = sSql & " Where cCtaCod=P.cCtaCod)),"
    sSql = sSql & " CuotasMora=(Select Count(*)"
    sSql = sSql & " From ColocCalendario CC"
    sSql = sSql & " Where CC.cCtaCod=P.cCtaCod and CC.nColocCalendApl=1 and CC.nNroCalen=(Select Max(nNroCalen) From ColocCalendDet"
    sSql = sSql & " Where cCtaCod=P.cCtaCod) and"
    sSql = sSql & " CC.nColocCalendEstado=0 and dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "')"
    sSql = sSql & " ,MontoCobrar=(Select Sum(CD.nMonto-CD.nMontoPagado)"
    sSql = sSql & " From ColocCalendario CC"
    sSql = sSql & " Inner Join ColocCalendDet CD on CC.cCtaCod=CD.cCtaCod And CC.nNroCalen=CD.nNroCalen and CC.nColocCalendApl=CD.nColocCalendApl and"
    sSql = sSql & " CC.nCuota = CD.nCuota"
    sSql = sSql & " Where CC.cCtaCod=P.cCtaCod and CD.nColocCalendApl=1 and CC.nNroCalen=(Select Max(nNroCalen) From ColocCalendDet"
    sSql = sSql & " Where cCtaCod=P.cCtaCod) and "
    sSql = sSql & " CC.dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "' and CC.nColocCalendEstado=0),"
    sSql = sSql & " cRFA=Case When CC.cRFA Is Null or CC.cRFA='NOR' Then 'NOR' Else CC.cRFA End "
    sSql = sSql & " From RRHH RH"
    sSql = sSql & " Inner Join ProductoPersona PP on RH.cPersCod=PP.cPersCod and PP.nPrdPersRelac=28"
    sSql = sSql & " Inner Join Producto P on P.cCtaCod=PP.cCtaCod"
    sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=P.cCtaCod"
    sSql = sSql & " Inner Join ProductoPersona PP1 on PP1.cCtacod=P.cCtaCod and PP1.nPrdPersRelac=20"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP1.cPersCod"
    
    If Len(psCodUbiGeo) <> 0 Then
        If Mid(Trim(psCodUbiGeo), 1, 1) = "3" Then 'Todos los del distrito
            sSql = sSql & " Inner Join UbicacionGeografica UG on Substring(UG.cUbiGeoCod,2,6)=Substring(Pers.cPersDireccUbiGeo,2,6) and Substring(UG.cUbiGeoCod,1,1)='3' and Substring(UG.cUbiGeoCod,8,5)='00000'"
        Else
            sSql = sSql & " Inner Join UbicacionGeografica UG on UG.cUbiGeoCod=Pers.cPersDireccUbiGeo"
        End If
    End If
    
    sSql = sSql & " Where P.nPrdEstado in (2020,2021,2022,2030,2031,2032) "
    sSql = sSql & " and (Select Count(*)"
    sSql = sSql & " From ColocCalendario CC"
    sSql = sSql & " Where CC.cCtaCod=P.cCtaCod and CC.nColocCalendApl=1 and CC.nNroCalen=(Select Max(nNroCalen) From ColocCalendDet"
    sSql = sSql & " Where cCtaCod=P.cCtaCod) and"
    sSql = sSql & " CC.nColocCalendEstado=0 and dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "')>0"
    sSql = sSql & " and RH.cPersCod in (" & psUser & ") and Substring(P.cCtaCod,6,3) in (" & psProductos & ")"
    'If Len(psCodUbiGeo) <> 0 Then
    '    If Mid(Trim(psCodUbiGeo), 1, 1) <> "3" Then 'Todos los del distrito
          ' por zonas
     If Len(psCodUbiGeo) <> 0 Then
            sSql = sSql & " and UG.cUbiGeoCod='" & Trim(psCodUbiGeo) & "'"
     End If
    '    End If
    'End If
    sSql = sSql & " and Substring(P.cCtaCod,4,2) in (" & psAgencias & ")"
    sSql = sSql & " and CC.nColocCondicion in (" & psCondicion & ")"
    sSql = sSql & " and Substring(P.cCtaCod,9,1) in (" & psMoneda & ")"
    sSql = sSql & " and P.nSaldo>=" & pnMontoIni & " and P.nSaldo<=" & pnMontoFin
    sSql = sSql & " and CC.nDiasAtraso between " & pnIniDiaMora & " and " & pnFinDiaMora & ")X"
    sSql = sSql & " Order By cUser,Titular"
    
    Set ListaMoraXAnalistaCabecera = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaMoraXAnalistaDetalle(ByVal psCtaCod As String, ByVal psFecha As String) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    Set oConec = New COMConecta.DCOMConecta
    sSql = "Select CONVERT(varchar(15),CC.dVenc,103) as dFecha,"
    sSql = sSql & " SaldoGastoMora=isnull((Select Sum(nMonto-nMontoPagado)"
    sSql = sSql & " From ColocCalendDet"
    sSql = sSql & " where cCtaCod=CC.cCtaCod and nNroCalen=CC.nNroCalen and nColocCalendApl=CC.nColocCalendApl and"
    sSql = sSql & " nCuota=CC.nCuota and nPrdConceptoCod in (1101) and nPrdConceptoCod like '12%'),0),"
    sSql = sSql & " SaldoCuota=isnull((Select Sum(nMonto-nMontoPagado)"
    sSql = sSql & " From ColocCalendDet"
    sSql = sSql & " where cCtaCod=CC.cCtaCod and nNroCalen=CC.nNroCalen and nColocCalendApl=CC.nColocCalendApl and"
    sSql = sSql & " nCuota=CC.nCuota),0),"
    sSql = sSql & " DateDiff(day,CC.dVenc,'" & Format(psFecha, "MM/dd/yyyy") & "') as AtrasoCuota,"
    sSql = sSql & " CC.nCuota"
    sSql = sSql & " From ColocCalendario CC"
    sSql = sSql & " Where CC.cCtaCod='" & psCtaCod & "' and CC.nColocCalendApl=1  and CC.nColocCalendEstado=0"
    sSql = sSql & " and CC.dVenc<'" & Format(psFecha, "MM/dd/yyyy") & "'"
    sSql = sSql & " and nNroCalen=(Select Max(nNroCalen) From ColocCalendario Where cCtaCod='" & psCtaCod & "')"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaMoraXAnalistaDetalle = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaDesembolsosXAgenciaXFecha(ByVal psCodAgen As String, ByVal dFecha As Date) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim sDia As String
    Dim sMes As String
    Dim sAno As String
    Dim sFecha As String
    
    sDia = Mid(dFecha, 1, 2)
    sMes = Mid(dFecha, 4, 2)
    sAno = Mid(dFecha, 7, 4)
    
    sFecha = sAno & sMes & sDia
    
    sSql = "Select *"
    sSql = sSql & " From("
    sSql = sSql & " Select P.cCtaCod,Pers.cPersNombre as cTitular,MD.nMonto as MontoDesem,"
    sSql = sSql & " Pers1.cPersNombre as cAnalista, Substring(M.cMovNro,7,2)+'/'+Substring(M.cMovNro,5,2)+'/'+Substring(M.cMovNro,1,4) as dFecha,"
    sSql = sSql & " Right(M.cMovNro,4) as cUsuario "
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join MovColDet MD on MD.cCtaCod=P.cCtaCod and MD.nPrdConceptoCod=1000"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=P.cCtaCod and PP.nPrdPersRelac=20"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Inner Join ProductoPersona PP1 on PP1.cCtaCod=P.cCtaCod and PP1.nPrdPersRelac=28"
    sSql = sSql & " Inner Join Persona Pers1 on Pers1.cPersCod=PP1.cPersCod"
    sSql = sSql & " Inner Join Mov M on M.nMovNro=MD.nMovNro"
    sSql = sSql & " Where MD.cOpeCod in ('100101','100102','100103','100104','100105','100106','100107','100109') and"
    sSql = sSql & " Left(M.cMovNro,8)='" & sFecha & "' and Substring(MD.cCtaCod,4,2) in(" & psCodAgen & ") and M.nMovFlag=0)X"
    sSql = sSql & " Order By cUsuario,cTitular"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaDesembolsosXAgenciaXFecha = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ListaPrendarioConsolidadAnalista(ByVal psCodAgencia As String, ByVal pdFecha As String, _
ByVal pbFecha As Boolean) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim sDia As String
    Dim sMes As String
    Dim sAno As String
    Dim sFecha As String
    
    sDia = Mid(pdFecha, 1, 2)
    sMes = Mid(pdFecha, 4, 2)
    sAno = Mid(pdFecha, 7, 4)
    
    sFecha = sAno & sMes & sDia
    
    
'    sSQL = "SELECT CCODAGE=SUBSTRING(C.CCTACOD,4,2),CANTIDAD=COUNT(*),DESEMBOLSADO=SUM(CO.NMONTOCOL),SALDOSK=SUM(C.NSALDOCAP),"
'    sSQL = sSQL & " SALDOSKPROM=SUM(C.NSALDOCAP)/COUNT(*)"
'    sSQL = sSQL & " FROM COLOCACSALDO C"
'    sSQL = sSQL & " JOIN COLOCACIONES CO ON CO.CCTACOD=C.CCTACOD"
'    sSQL = sSQL & " WHERE C.CCTACOD LIKE '108__305%' AND CONVERT(CHAR(8),C.DFECHA,112)='" & sFecha & "'"
'    sSQL = sSQL & " AND  C.NPRDESTADO IN ('2107','2102','2101','2106','2104') AND SUBSTRING(C.CCTACOD,4,2) in (" & psCodAgencia & ")"
'    sSQL = sSQL & " GROUP BY SUBSTRING(C.CCTACOD,4,2)"

    If pbFecha = True Then
        sSql = " SELECT CCODAGE=SUBSTRING(p.CCTACOD,4,2),CANTIDAD=COUNT(*),"
        sSql = sSql & " DESEMBOLSADO=SUM(CO.NMONTOCOL),SALDOSK=SUM(p.NSALDO),"
        sSql = sSql & " SALDOSKPROM=SUM(p.NSALDO)/COUNT(*)"
        sSql = sSql & " FROM producto p"
        sSql = sSql & " join COLOCACIONES CO on co.cctacod=p.cctacod"
        sSql = sSql & " WHERE Co.CCTACOD LIKE '108__305%' AND p.NPRDESTADO IN ('2107','2102','2101','2106','2104') AND SUBSTRING(p.CCTACOD,4,2) in (" & psCodAgencia & ")"
        sSql = sSql & " GROUP BY SUBSTRING(p.CCTACOD,4,2)"
   Else
            
     sSql = "SELECT CCODAGE=SUBSTRING(C.CCTACOD,4,2),CANTIDAD=COUNT(*),DESEMBOLSADO=SUM(CO.NMONTOCOL),SALDOSK=SUM(C.NSALDOCAP),"
     sSql = sSql & " SALDOSKPROM=SUM(C.NSALDOCAP)/COUNT(*)"
     sSql = sSql & " FROM COLOCACSALDO C"
     sSql = sSql & " JOIN COLOCACIONES CO ON CO.CCTACOD=C.CCTACOD"
     sSql = sSql & " WHERE C.CCTACOD LIKE '108__305%' AND CONVERT(CHAR(8),C.DFECHA,112)='" & sFecha & "'"
     sSql = sSql & " AND  C.NPRDESTADO IN ('2107','2102','2101','2106','2104') AND SUBSTRING(C.CCTACOD,4,2) in (" & psCodAgencia & ")"
     sSql = sSql & " GROUP BY SUBSTRING(C.CCTACOD,4,2)"
   End If
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaPrendarioConsolidadAnalista = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerTipoCambio(ByVal pFechaCambio As Date) As Double
    Dim rs As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim sDia As String
    Dim sMes As String
    Dim sAno As String
    Dim sFecha As String
    Dim nValFijo As Double
    
    sDia = Right("00", 2 - Len(CStr(Day(pFechaCambio)))) & Day(pFechaCambio)
    sMes = Right("00", 2 - Len(CStr(Month(pFechaCambio)))) & Month(pFechaCambio)
    sAno = Year(pFechaCambio)
    
    sFecha = CStr(sAno) + CStr(sMes) + CStr(sDia)
    
    sSql = "select Max(nValFijo) as nValFijo"
    sSql = sSql & " From TipoCambio"
    sSql = sSql & " where cast(datepart(yyyy,dFecCamb) as varchar(4))+"
    sSql = sSql & " right('00',2-len(cast(datepart(mm,dFecCamb) as varchar(2))))+cast(datepart(mm,dFecCamb) as varchar(2))+"
    sSql = sSql & " right('00',2-len(cast(datepart(dd,dFecCamb) as varchar(2))))+cast(datepart(dd,dFecCamb) as varchar(2))='" & sFecha & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not rs.EOF And Not rs.BOF Then
        nValFijo = rs!nValFijo
    End If
    Set rs = Nothing
    
    ObtenerTipoCambio = nValFijo
End Function

Public Function ListaClientesAutomaticos(ByVal dFecha As Date, ByVal psAgencias As String, ByVal psAnalistas As String, _
ByVal psProductos As String) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    Set oConec = New COMConecta.DCOMConecta
    sSql = "SELECT *"
    sSql = sSql & " FROM ("
    sSql = sSql & " SELECT RH.CUSER,PERS.cPersCod,PERS.CPERSNOMBRE,PERS.CPERSDIRECCDOMICILIO,"
    sSql = sSql & " SUM(Cast(CS.DIASATRASO as Numeric(6,2))) AS DIASATRASO,SUM(Cast(CS.NCANTIDAD as numeric(6,2))) AS NCANTIDAD,"
    sSql = sSql & " (SUM(Cast(CS.DIASATRASO as Numeric(6,2)))/SUM(Cast(CS.NCANTIDAD as numeric(6,2)))) as nCalificacion,"
    sSql = sSql & " cCalificacion=Case"
    sSql = sSql & " When (SUM(Cast(CS.DIASATRASO as Numeric(6,2)))/SUM(Cast(CS.NCANTIDAD as numeric(6,2)))) between 0 and 3.5 then 'A'"
    sSql = sSql & " When (SUM(Cast(CS.DIASATRASO as Numeric(6,2)))/SUM(Cast(CS.NCANTIDAD as numeric(6,2)))) between 3.6 and 4.5 then 'B'"
    sSql = sSql & " Else 'NINGUNO' End,"
    sSql = sSql & " P.CCTACOD,"
    sSql = sSql & " P.NSALDO,"
    sSql = sSql & " C.nMontoCol"
    sSql = sSql & " FROM RRHH RH"
    sSql = sSql & " INNER JOIN PRODUCTOPERSONA PP ON PP.CPERSCOD=RH.CPERSCOD AND PP.NPRDPERSRELAC=28"
    sSql = sSql & " INNER JOIN PRODUCTOPERSONA PP1 ON PP1.CCTACOD=PP.CCTACOD AND PP1.NPRDPERSRELAC=20"
    sSql = sSql & " INNER JOIN PERSONA PERS ON PERS.cPersCod=PP1.CPERSCOD"
    sSql = sSql & " INNER JOIN PRODUCTO P ON P.CCTACOD=PP.CCTACOD"
    sSql = sSql & " INNER JOIN COLOCACIONES C ON C.CCTACOD=P.CCTACOD"
    sSql = sSql & " INNER JOIN (SELECT SUM(CASE WHEN DATEDIFF(DAY,CD.DVENC,CD.DPAGO)< 0 THEN 0 ELSE DATEDIFF(DAY,CD.DVENC, CD.DPAGO) END) AS DIASATRASO,"
    sSql = sSql & " (SELECT COUNT(*)"
    sSql = sSql & " FROM COLOCCALENDARIO CO"
    sSql = sSql & " INNER JOIN PRODUCTOPERSONA PP2 ON PP2.CCTACOD=CO.CCTACOD AND PP2.NPRDPERSRELAC=20"
    sSql = sSql & " WHERE  CO.NCOLOCCALENDAPL=1 AND CO.NCOLOCCALENDESTADO=1 AND PP2.CPERSCOD=PPX.CPERSCOD) AS NCANTIDAD,"
    sSql = sSql & " PPX.cPersCod"
    sSql = sSql & " FROM COLOCCALENDARIO CD"
    sSql = sSql & " INNER JOIN PRODUCTOPERSONA PPX ON PPX.CCTACOD=CD.CCTACOD AND PPX.NPRDPERSRELAC=20"
    sSql = sSql & " Where CD.NCOLOCCALENDAPL = 1 And CD.NCOLOCCALENDESTADO = 1"
    sSql = sSql & " GROUP BY PPX.CPERSCOD) CS ON CS.CPERSCOD=PP1.CPERSCOD"
    sSql = sSql & " WHERE RH.CUSER IN (" & psAnalistas & ") AND SUBSTRING(PP1.CCTACOD,4,2) IN (" & psAgencias & ") And "
    sSql = sSql & " (SELECT DATEDIFF(DAY,MIN(CO.DPAGO),'" & CStr(Format(dFecha, "MM/dd/yyyy")) & "')"
    sSql = sSql & " FROM COLOCCALENDARIO CO"
    sSql = sSql & " INNER JOIN PRODUCTOPERSONA PP2 ON PP2.CCTACOD=CO.CCTACOD AND PP2.NPRDPERSRELAC=20"
    sSql = sSql & " WHERE PP2.CPERSCOD=PP1.CPERSCOD AND CO.NCOLOCCALENDAPL=0 AND CO.NCOLOCCALENDESTADO=1)>360 AND"
    sSql = sSql & " (SELECT COUNT(*)"
    sSql = sSql & " FROM PRODUCTOPERSONA PP2"
    sSql = sSql & " INNER JOIN COLOCACCRED CD ON CD.CCTACOD=PP2.CCTACOD AND PP2.NPRDPERSRELAC=20"
    sSql = sSql & " WHERE PP2.CPERSCOD=PP1.CPERSCOD)>=3 AND"
    sSql = sSql & " SUBSTRING(PP1.CCTACOD,6,3)IN (" & psProductos & ")"
    sSql = sSql & " GROUP BY RH.CUSER,PERS.cPersCod,PERS.CPERSNOMBRE,PERS.CPERSDIRECCDOMICILIO,P.CCTACOD,P.NSALDO,C.NMONTOCOL)X"
    sSql = sSql & " WHERE cCalificacion<>'NINGUNO'"
    sSql = sSql & " ORDER BY cCalificacion,CPERSNOMBRE"
    oConec.AbreConexion
    Set ListaClientesAutomaticos = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function Repor108205_ClientesVigentes(ByVal psFecha As Date) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "SELECT PERS.CPERSNOMBRE,C.NMONTOCOL,P.nSaldoCap,C.dVigencia,C.dVenc,"
    sSql = sSql & " AG.CAGEDESCRIPCION AS NOMAGENCIA,"
    sSql = sSql & " MONEDA=CASE WHEN SUBSTRING(PP.CCTACOD,9,1)='1' THEN 'SOLES'ELSE'DOLARES'END,"
    sSql = sSql & " DateDiff(year,PERS.dPersNacCreac,'" & CStr(Format(psFecha, "MM/dd/yyyy")) & "') as EDAD"
    sSql = sSql & " FROM PRODUCTOPERSONA PP"
    sSql = sSql & " INNER JOIN COLOCACSALDO P ON P.CCTACOD=PP.CCTACOD"
    sSql = sSql & " INNER JOIN AGENCIAS AG ON AG.CAGECOD=SUBSTRING(P.CCTACOD,4,2)"
    sSql = sSql & " INNER JOIN COLOCACIONES C ON C.CCTACOD=P.CCTACOD"
    sSql = sSql & " INNER JOIN PERSONA PERS ON PERS.CPERSCOD=PP.CPERSCOD"
    sSql = sSql & " WHERE P.NPRDESTADO IN (2020,2021,2022,2030,2031,2032,2201,2205) AND SUBSTRING(P.CCTACOD,6,3)<>'305' AND"
    sSql = sSql & " SUBSTRING(P.CCTACOD,7,2) NOT LIKE '21%' AND PP.NPRDPERSRELAC=20 AND P.dFecha='" & CStr(Format(psFecha, "MM/dd/yyyy")) & "'"
    sSql = sSql & " AND DATEDIFF(YEAR,Pers.dPersNacCreac,'" & CStr(Format(psFecha, "MM/dd/yyyy")) & "')<=65"
    sSql = sSql & " ORDER BY AG.CAGEDESCRIPCION,PERS.CPERSNOMBRE"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set Repor108205_ClientesVigentes = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ListaCreditosNoDesembolsados(ByVal pdFechaInicio As Date, ByVal psFechaFin As Date, _
ByVal psAgencia As String) As Recordset
  
  Dim oConec As COMConecta.DCOMConecta
  Dim sSql As String
  
  sSql = "Select P.cCtaCod,Pers.cPersNombre,Pers.cPersDireccDomicilio,CE.nMonto,CE.dPrdEstado,"
  sSql = sSql & " Moneda=Case Substring(P.cCtaCod,9,1)"
  sSql = sSql & " When '1' Then 'Soles'"
  sSql = sSql & " When '2' Then 'Dolares'"
  sSql = sSql & " End,"
  sSql = sSql & " RH.cUser"
  sSql = sSql & " From Producto P"
  sSql = sSql & " Inner Join ProductoPersona PP On PP.cCtaCod=P.cCtaCod and PP.nPrdPersRelac=20"
  sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
  sSql = sSql & " Inner Join ColocacEstado CE on CE.cCtaCod=P.cCtaCod  and CE.nPrdEstado=2002"
  sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=P.cCtaCod"
  sSql = sSql & " Inner Join ProductoPersona PP1 on PP1.cCtaCod=P.cCtaCod and PP1.nPrdPersRelac=28"
  sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP1.cPersCod"
  sSql = sSql & " Where P.nPrdEstado=2002 and Substring(P.cCtaCod,4,2) in (" & psAgencia & ") and"
  sSql = sSql & " CE.dPrdEstado Between '" & CStr(Format(pdFechaInicio, "MM/dd/yyyy")) & "' and DateAdd(Day,1,'" & CStr(Format(psFechaFin, "MM/dd/yyyy")) & "')"
  sSql = sSql & " Order By CE.dPrdEstado,RH.cUser,Pers.cPersNombre"
  
  Set oConec = New COMConecta.DCOMConecta
  oConec.AbreConexion
  Set ListaCreditosNoDesembolsados = oConec.CargaRecordSet(sSql)
  oConec.CierraConexion
  Set oConec = Nothing
End Function


Public Function ListaCreditosVencidos(ByVal sAgencias As String, ByVal pdFecha As Date) As Recordset
  Dim oConec As COMConecta.DCOMConecta
  Dim sSql As String
  
  sSql = "SELECT P.cCtaCod,PERS.CPERSNOMBRE AS CNOMTITULAR,RTRIM(LTRIM(PERS.CPERSDIRECCDOMICILIO)) AS DIRECTITULAR,"
  sSql = sSql & " PERS1.CPERSNOMBRE AS CNOMGARANT,PERS1.CPERSDIRECCDOMICILIO AS DIRECGARANTE,"
  sSql = sSql & " RH.CUSER AS CANALISTA,CO.DPAGO AS FECHAPAGO,C.NDIASATRASO,COO.NMONTOCOL,"
  sSql = sSql & " P.NSALDO,"
  sSql = sSql & " CAPITAL=(SELECT SUM(NMONTO-NMONTOPAGADO)"
  sSql = sSql & " FROM COLOCCALENDARIO CX"
  sSql = sSql & " INNER JOIN COLOCCALENDDET CDX ON CDX.CCTACOD=CX.CCTACOD AND CDX.NNROCALEN=CX.NNROCALEN AND"
  sSql = sSql & " CDX.NCOLOCCALENDAPL = CX.NCOLOCCALENDAPL And CDX.nCuota = CX.nCuota"
  sSql = sSql & " WHERE CX.NCOLOCCALENDAPL=1 AND CDX.NPRDCONCEPTOCOD IN (1000,1010) and CX.cCtaCod=C.cCtaCod and"
  sSql = sSql & " CX.dVenc<'" & CStr(Format(pdFecha, "MM/dd/yyyy")) & "'),"
  sSql = sSql & " Interes=(SELECT SUM(NMONTO-NMONTOPAGADO)"
  sSql = sSql & " FROM COLOCCALENDARIO CX"
  sSql = sSql & " INNER JOIN COLOCCALENDDET CDX ON CDX.CCTACOD=CX.CCTACOD AND CDX.NNROCALEN=CX.NNROCALEN AND"
  sSql = sSql & " CDX.NCOLOCCALENDAPL = CX.NCOLOCCALENDAPL And CDX.nCuota = CX.nCuota"
  sSql = sSql & " WHERE CX.NCOLOCCALENDAPL=1 AND CDX.NPRDCONCEPTOCOD IN (1100,1109) and CX.cCtaCod=C.cCtaCod and"
  sSql = sSql & " CX.dVenc<'" & CStr(Format(pdFecha, "MM/dd/yyyy")) & "'),"
  sSql = sSql & " Mora=(SELECT SUM(NMONTO-NMONTOPAGADO)"
  sSql = sSql & " FROM COLOCCALENDARIO CX"
  sSql = sSql & " INNER JOIN COLOCCALENDDET CDX ON CDX.CCTACOD=CX.CCTACOD AND CDX.NNROCALEN=CX.NNROCALEN AND"
  sSql = sSql & " CDX.NCOLOCCALENDAPL = CX.NCOLOCCALENDAPL And CDX.nCuota = CX.nCuota"
  sSql = sSql & " WHERE CX.NCOLOCCALENDAPL=1 AND CDX.NPRDCONCEPTOCOD IN (1101,1108) and CX.cCtaCod=C.cCtaCod and"
  sSql = sSql & " CX.dVenc<'" & CStr(Format(pdFecha, "MM/dd/yyyy")) & "'),"
  sSql = sSql & " DeudaTotal=(SELECT SUM(NMONTO-NMONTOPAGADO)"
  sSql = sSql & " FROM COLOCCALENDARIO CX"
  sSql = sSql & " INNER JOIN COLOCCALENDDET CDX ON CDX.CCTACOD=CX.CCTACOD AND CDX.NNROCALEN=CX.NNROCALEN AND"
  sSql = sSql & " CDX.NCOLOCCALENDAPL = CX.NCOLOCCALENDAPL And CDX.nCuota = CX.nCuota"
  sSql = sSql & " WHERE CX.NCOLOCCALENDAPL=1 and CX.cCtaCod=C.cCtaCod and CX.dVenc<'" & CStr(Format(pdFecha, "MM/dd/yyyy")) & "'),"
  sSql = sSql & " AG.cAgeDescripcion  as sNomAge,"
  sSql = sSql & " Moneda=Case When Substring(P.cCtaCod,9,1)='1' Then 'Soles' Else 'Dolares' End"
  sSql = sSql & " FROM COLOCACCRED C"
  sSql = sSql & " INNER JOIN PRODUCTOPERSONA PP ON C.CCTACOD=PP.CCTACOD AND PP.NPRDPERSRELAC=20"
  sSql = sSql & " INNER JOIN PERSONA PERS ON PERS.CPERSCOD=PP.CPERSCOD"
  sSql = sSql & " LEFT JOIN  PRODUCTOPERSONA PP1 ON PP1.CCTACOD=C.CCTACOD AND PP1.NPRDPERSRELAC=25"
  sSql = sSql & " LEFT JOIN PERSONA PERS1 ON PERS1.CPERSCOD=PP1.CPERSCOD"
  sSql = sSql & " INNER JOIN PRODUCTOPERSONA PP2 ON PP2.CCTACOD=C.CCTACOD  AND PP2.NPRDPERSRELAC=28"
  sSql = sSql & " INNER JOIN RRHH RH ON RH.CPERSCOD=PP2.CPERSCOD"
  sSql = sSql & " INNER JOIN COLOCCALENDARIO CO ON CO.CCTACOD=C.CCTACOD AND CO.NCOLOCCALENDAPL=0 AND CO.NCOLOCCALENDESTADO=1 AND CO.NNROCALEN=C.NNROCALEN"
  sSql = sSql & " INNER JOIN COLOCACIONES COO ON COO.CCTACOD=CO.CCTACOD"
  sSql = sSql & " INNER JOIN PRODUCTO P ON P.CCTACOD=C.CCTACOD"
  sSql = sSql & " Inner Join Agencias AG on AG.cAgeCod=SubString(P.cCtaCod,4,2)"
  sSql = sSql & " Where C.nDiasAtraso>=90 and P.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
  sSql = sSql & " AND Substring(P.cCtaCod,4,2) in (" & sAgencias & ")"
  sSql = sSql & " Order By RH.cUser,PERS.CPERSNOMBRE  "
  
  Set oConec = New COMConecta.DCOMConecta
  oConec.AbreConexion
  Set ListaCreditosVencidos = oConec.CargaRecordSet(sSql)
  oConec.CierraConexion
  Set oConec = Nothing
End Function


Public Function ObtenerGarantes(ByVal psCtaCod As String) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select Pers.cPersNombre,Pers.cPersDireccDomicilio"
    sSql = sSql & " from ProductoPersona PP"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Where  PP.nPrdPersRelac=25 and PP.cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerGarantes = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function Repor_HistorialCliente(ByVal pscPerscod As String) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
'CTI6-20210503-ERS032-2019
'    sSql = "Select  distinct CCred.nDiasAtrasoAcum, C.dVigencia,SUBSTRING(C.cCtaCod,6,3) nTipoCredCod ,"
'    sSql = sSql & "  nNota=ISNULL((Select nColocNota From ColocCalificacionAnalista A "
'    sSql = sSql & "                 where ctipo <> 'S'"
'    sSql = sSql & "                 and dColocNotaFecha = (Select MAX(dColocNotaFecha)"
'    sSql = sSql & "                         From ColocCalificacionAnalista Where cCtaCod = A.cCtaCod )"
'    sSql = sSql & "                     AND cCtaCod = C.cCtaCod),0),"
'    sSql = sSql & " CN3.cConsDescripcion Partic, P.cCtaCod,Cn.cConsDescripcion as EstadoCredito,C.nMontoCol,P.nSaldo,CE.dPrdEstado,C.dVenc,"
'    sSql = sSql & " RH.cUser,CN1.cConsDescripcion as TipoCredito,Pers1.cPersNombre as cNomIntitucion,"
'    sSql = sSql & " Moneda=Case When Substring(P.cCtaCod,9,1)='1' Then 'SOLES' Else 'DOLARES' End,"
'    sSql = sSql & " PID.cPersIDnro,Pers.cPersNombre as cNomCli"
'    'ARCV 03-11-2006
'    sSql = sSql & " ,CCred.nDiasAtraso "
'    '--------------
'    '**DAOR 20071121************
'    sSql = sSql & ", dDesembolso = (Select  top 1 dPrdEstado From ColocacEstado Where cCtaCod=P.cCtaCod and nPrdEstado=2020 order by dPrdEstado) "
'    '***************************
'    sSql = sSql & " From Persona Pers"
'    sSql = sSql & " Inner Join PersId  PID on PID.cPersCod=Pers.cPersCod and Pers.nPersPersoneria = PID.cPersIDTpo"
'    sSql = sSql & " Inner Join ProductoPersona PP on PP.cPersCod=Pers.cPersCod and PP.nPrdPersRelac not in (28,29)"
'    sSql = sSql & " Inner Join Constante CN3 ON CN3.nConsvalor = PP.nPrdpersrelac and CN3.nConsCod = 3002 "
'    sSql = sSql & " Inner Join Producto  P on P.cCtaCod=PP.cCtaCod"
'    sSql = sSql & " Inner Join Constante CN on CN.nConsValor=P.nPrdEstado and CN.nConsCod=3001 and CN.nConsValor<>3001"
'    sSql = sSql & " Inner Join Colocaciones C on C.cCtaCod=P.cCtaCod"
'    'ARCV 03-11-2006
'    sSql = sSql & " Inner Join ColocacCred CCred on CCred.cCtaCod=C.cCtaCod"
'    '--------------
'    sSql = sSql & " Left Join ColocacEstado CE on CE.cCtaCod=P.cCtaCod and CE.nPrdEstado=2002"
'    sSql = sSql & " Left Join ProductoPersona PP1 on PP1.cCtacod=P.cCtaCod and PP1.nPrdPersRelac=28"
'    sSql = sSql & " Left Join RRHH RH on RH.cPersCod=PP1.cPersCod"
'    sSql = sSql & " Inner Join Constante CN1 on CN1.nConsValor=SubString(P.cctaCod,6,3) and CN1.nConsCod=1001 and CN1.nConsValor<>1001"
'    sSql = sSql & " Left Join  ColocacConvenio CC on CC.cCtaCod=P.cCtaCod"
'    sSql = sSql & " Left Join Persona Pers1 on Pers1.cPersCod=CC.cPersCod"
'    sSql = sSql & " Where PP.cPersCod='" & pscPerscod & "'"
'    sSql = sSql & " Order By C.dVenc"
    
    sSql = "exec stp_sel_ReporteHistorialCliente '" & pscPerscod & "'"
    'Fin CTI6-20210503
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set Repor_HistorialCliente = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function Repor_HistorialCreditoCabecera(ByVal psCtaCod As String) As Recordset
   Dim oConec As COMConecta.DCOMConecta
   Dim sSql As String
   
   sSql = "Select Pers.cPersNombre,Pers.cPersCod,P.cCtaCod,CN.cConsDescripcion as cEstado,"
   sSql = sSql & " PTI.nTasaInteres as TasaInteres,C.nMontoCol,C.dVenc,C.dVigencia,RH.cUser"
   sSql = sSql & " From Producto P"
   sSql = sSql & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=20"
   sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
   sSql = sSql & " Inner Join Constante Cn on Cn.nConsValor=P.nPrdEstado and CN.nConsCod=3001 and CN.nConsValor<>3001"
   sSql = sSql & " Inner Join Colocaciones C on C.cCtaCod=P.cCtaCod"
   sSql = sSql & " Inner Join ProductoTasaInteres PTI on PTI.cCtaCod=P.cCtaCod and PTI.nPrdTasaInteres=1"
   sSql = sSql & " Inner Join ProductoPersona PP1 on PP1.cCtaCod=P.cCtaCod and PP1.nPrdPersRelac=28"
   sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP1.cPersCod"
   sSql = sSql & " Where P.cCtaCod='" & psCtaCod & "'"
   
   Set oConec = New COMConecta.DCOMConecta
   oConec.AbreConexion
   Set Repor_HistorialCreditoCabecera = oConec.CargaRecordSet(sSql)
   oConec.CierraConexion
   Set oConec = Nothing
End Function


Public Function Repo_HistorialCreditoDetalle(ByVal psCtaCod As String) As Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta

'Modificado YIHU20141203ERS1692014
'sSql = "Select Y.nNroCuota,  Y.dVenc, Y.FechaPago,"
'sSql = sSql & " Capital=IsNull((Select Sum(nMonto) From MovColDet Where nMovNro=M1.nMovNro and nPrdConceptoCod in (1000) and nNroCuota=Y.nNroCuota AND cCtaCod=Y.cCtaCod and cOpeCod not like '107[1234]%'),0)," 'and cOpeCod in ('100101','100102','100103','100104','100105') --ARCV 04-07-2007
'sSql = sSql & " Interes=IsNull((Select Sum(nMonto) From MovColDet Where nMovNro=M1.nMovNro and nPrdConceptoCod in (1100) and nNroCuota=Y.nNroCuota AND cCtaCod=Y.cCtaCod),0),"
'sSql = sSql & " Mora=IsNull((Select Sum(nMonto) From MovColDet Where nMovNro=M1.nMovNro and nPrdConceptoCod in (1101) and nNroCuota=Y.nNroCuota AND cCtaCod=Y.cCtaCod),0),"
'sSql = sSql & " Gastos=Isnull((Select Sum(nMonto) From MovColDet Where nMovNro=M1.nMovNro and Cast(nPrdConceptoCod  as varchar(10)) like '12%' and nNroCuota=Y.nNroCuota AND cCtaCod=Y.cCtaCod),0),"
''RIRO 20130724 SEGUN ERS098-2013 ***********
'sSql = sSql & " nSaldoCap= ISNULL((SELECT top 1 nSaldoCap FROM MovCol WHERE nMovNro=M1.nMovNro and cCtaCod = y.cCtaCod and (cOpeCod like '100[234567]%' OR cOpeCod = '105001') ORDER by nSaldoCap DESC), 0), "
'sSql = sSql & " IntGra=IsNull((Select Sum(nMonto) From MovColDet Where nMovNro=M1.nMovNro and nPrdConceptoCod in (1102) and nNroCuota=Y.nNroCuota AND cCtaCod=Y.cCtaCod),0),"
''FIN RIRO **********************************
'sSql = sSql & " ITF=Isnull((Select Sum(nMonto) From MovColDet Where nMovNro=M1.nMovNro and nPrdConceptoCod=20),0), Right(M1.cMovNro,4) as cUsuario,"
'sSql = sSql & " nDiasAtraso = DateDiff(Day, CCO.dVenc, FechaPago)"
'sSql = sSql & " From Mov M1"
'sSql = sSql & " Inner Join (Select nMovNro,nNroCuota,dVenc,FechaPago,cCtaCod,nNroCalen,nColocCalendApl"
'sSql = sSql & " From (Select M.nMovNro, MD.nNroCuota, CO.dVenc, Cast((Substring(Left(M.cMovNro,8),5,2)+'/'+ Substring(Left(M.cMovNro,8),7,2)+'/'+Substring(Left(M.cMovNro,8),1,4)) as Datetime) as FechaPago, MD.cCtaCod,CD.nNroCalen,CO.nColocCalendApl"
'sSql = sSql & " From Mov M"
'sSql = sSql & " Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
'sSql = sSql & " Inner Join ColocCalendario CO on CO.cCtaCod=MD.cCtaCod  and Co.nCuota=MD.nNroCuota and CO.nColocCalendApl=(Case When Md.cOpeCod in ('100101','100102','100103','100104','100105') Then 0 else 1 end)"
'sSql = sSql & " Inner Join ColocacCred CD on CD.cCtaCod=MD.cCtaCod and CD.nNroCalen=MD.nNroCalen"
'sSql = sSql & " Where MD.cCtaCod='" & psCtaCod & "' and M.nMovFlag=0 and"
''sSql = sSql & " CO.nNroCalen=(Select nNroCalen From ColocacCred Where cCtaCod='" & psCtaCod & "') and MD.cOpeCod not like '99%') X"
'sSql = sSql & " CO.nNroCalen=(Select nNroCalen From ColocacCred Where cCtaCod='" & psCtaCod & "') and Not (Left(MD.cOpeCod,2) In ('99') Or M.cOpeCod in ('172101'))) X"
'sSql = sSql & " Group By nMovNro,nNroCuota,dVenc,FechaPago,cCtaCod,nNroCalen,nColocCalendApl)Y on M1.nMovNro=Y.nMovNro and M1.nMovFlag=0"
'sSql = sSql & " Inner Join ColocCalendario CCO on CCO.cCtaCod=Y.cCtaCod and CCO.nCuota=Y.nNroCuota and Y.nNroCalen=CCO.nNroCalen and Y.nColocCalendApl=CCO.nColocCalendApl"
'sSql = sSql & " Order By M1.nMovNro"

sSql = "stp_sel_Repo_HistorialCreditoDetalle '" & psCtaCod & "'"

'END YIHU

Set oConec = New COMConecta.DCOMConecta
oConec.AbreConexion
Set Repo_HistorialCreditoDetalle = oConec.CargaRecordSet(sSql)
oConec.CierraConexion
Set oConec = Nothing
End Function

Function Repo_HistorialCredDetalleComplem(ByVal psCtaCod As String, ByVal pdFecha As Date) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
        
    sSql = "Select Co.nCuota,CO.dVenc,NULL as FechaPago,0 as Captital,0 as Interes,"
    sSql = sSql & " 0 as Mora,0 as Gastos,0 as ITF,NULL as cUsuario,DateDiff(Day,CO.dVenc,'" & CStr(Format(pdFecha, "MM/dd/yyyy")) & "') as nDiasAtraso, "
    'RIRO 20130724 SEGUN ERS098-2013 ***********
    sSql = sSql & " 0 as IntGra, "
    sSql = sSql & " nSaldoCap = isnull((SELECT top 1 nSaldoCap FROM MovCol WHERE cCtaCod = CO.cCtaCod and (cOpeCod like '100[234567]%' OR cOpeCod = '105001') and nSaldoCap >0 ORDER by nSaldoCap asc),0)"
    'FIN RIRO **********************************
    sSql = sSql & " From ColocCalendario CO"
    sSql = sSql & " Inner Join ColocacCred CC on CO.cCtaCod=CC.cCtaCod and CO.nNroCalen=CC.nNroCalen"
    sSql = sSql & " Where Co.cCtaCod='" & psCtaCod & "' and CO.nColocCalendApl=1 and Co.nColocCalendEstado=0"
    '**Comentado por Juez 20120315, para solución Incidente Nº INC1203150006*****
    'sSql = sSql & " and (Select Sum(nMontoPagado) From ColocCalendDet Where cCtaCod=CO.cCtaCod and nCuota=CO.nCuota and nColocCalendApl=CO.nColocCalendApl and"
    'sSql = sSql & " nNroCalen=CO.nNroCalen)=0"
    sSql = sSql & " Order By CO.nCuota,nColocCalendApl"
 
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set Repo_HistorialCredDetalleComplem = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function Recup_ClientesByAnalista(ByVal psAnalista As String, ByVal psAgencias As String) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select Distinct Pers.cPersCod,Pers.cPersNombre,Pers.cPersDireccDomicilio,Pers.cPersTelefono,Pers.cPersTelefono2,RH.cUser"
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=20"
    sSql = sSql & " Inner Join ProductoPersona PP1 on P.cCtaCod=PP1.cCtaCod and PP1.nPrdPersRelac=28"
    sSql = sSql & " Inner JOin Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Inner Join RRHH Rh on Rh.cPersCod=PP1.cPersCod"
    sSql = sSql & " Where Rh.cPersCod in (" & psAnalista & ") and SubString(P.cCtaCod,4,2) in ( " & psAgencias & ") and P.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205)"
    sSql = sSql & " Order By RH.cUser,Pers.cPersNombre"

    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set Recup_ClientesByAnalista = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function Recup_CalidadCartera(ByVal pnTipoCambio As Double)
    Dim oConec  As COMConecta.DCOMConecta
    Dim sSql As String
    
    'Recupera la calidad de la cartera
    
    sSql = ""
End Function

Public Function Recup_MontoAtrasado(ByVal pdFecha As Date, ByVal psCtaCod As String) As Double
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select Sum(CD.nMonto-CD.nMontoPagado) as nMonto"
    sSql = sSql & " From ColocacCred C"
    sSql = sSql & " Inner Join  ColocCalendario CO on C.cCtaCod=CO.cCtaCod and C.nNroCalen=CO.nNroCalen"
    sSql = sSql & " Inner Join ColocCalendDet CD on CD.cCtaCod=CO.cCtaCod and  CD.nNroCalen=CO.nNroCalen And CD.nColocCalendApl=CO.nColocCalendApl and CD.nCuota=CO.nCuota"
    sSql = sSql & " Where CD.cCtaCod='" & psCtaCod & "' and dVenc<='" & Format(pdFecha, "MM/dd/yyyy") & "' and nColocCalendEstado=0"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql, adLockReadOnly)
    oConec.CierraConexion
    
    If Not rs.EOF And Not rs.BOF Then
        Recup_MontoAtrasado = Format(rs!nMonto, "#0.00")
    End If
        
    Set rs = Nothing
End Function

Function VerificaConvenioEmp(ByVal psCtaCod As String) As Boolean
    'funcion que verifica si existe convenio a traves del filtro de los gastos
    'se hizo para el gasto de credicom
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "Select Count(*) as nCantidad"
    sSql = sSql & " From ColocacConvenio  CC"
    sSql = sSql & " Inner Join ProductoConceptoFiltro PCF on SubString(CC.cCtaCod,6,3)=PCF.nProdCod and SubString(CC.cCtaCod,4,2)=PCF.cAgeCod and PCF.cIntitucion=CC.cPersCod"
    sSql = sSql & " Where CC.cCtaCod='" & psCtaCod & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
        
    VerificaConvenioEmp = False
        
    If Not rs.EOF And Not rs.BOF Then
        If val(rs!nCantidad) > 0 Then
            VerificaConvenioEmp = True
        Else
            VerificaConvenioEmp = False
        End If
    Else
        VerificaConvenioEmp = False
    End If
    
    Set rs = Nothing
End Function

Public Function Recup_CreditoAntiguo(ByVal psCtaCod As String) As String
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "Select cCtaCodAnt From RelCuentas Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        Recup_CreditoAntiguo = rs!CctacodAnt
    End If
    Set rs = Nothing
End Function

Public Function Recup_FechaCancelacion(ByVal psCtaCod As String) As String
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    'CTI6-20210503-ERS032-2019
'    sSQL = "Select C.dVenc"
'    sSQL = sSQL & " From Producto P"
'    sSQL = sSQL & " Inner Join Colocaciones C on C.cCtaCod=P.cCtaCod"
'    sSQL = sSQL & " Where P.cCtaCod='" & psCtaCod & "' and P.nPrdEstado=2050"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "stp_sel_DatosPosicionRecupFechaCancelacion '" & psCtaCod & "'"
    'Fin CTI6-20210503
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        Recup_FechaCancelacion = IIf(IsNull(rs!dVenc), "NO TIENE CANCELACION", Format(rs!dVenc, "DD/mm/yyyy"))
    End If
    Set rs = Nothing
    If Recup_FechaCancelacion = "" Then
        Recup_FechaCancelacion = "NO TIENE CANCELACION"
    End If
End Function

Public Function GetUsuario(ByVal pnMovNro As Long) As String
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    sSql = "Select Right(cMovNro,4) as cUser From Mov Where nMovNro=" & pnMovNro
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    If Not rs.EOF And Not rs.BOF Then
        GetUsuario = rs!cUser
    End If
    Set rs = Nothing
End Function

Public Function GetCalidadCartera(ByVal pnTipoCambio As Double) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select 'VIGENTE NORMAL (0)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(CCP.cCtaCod,9,1)"
    sSql = sSql & " When '1' Then  CCP.nSaldoCap"
    sSql = sSql & " When '2' Then  CCP.nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(CCP.cCtaCod,9,1)"
    sSql = sSql & " When '1' Then  CCP.nProvision"
    sSql = sSql & " When '2' Then  CCP.nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(CCP.cCtaCod),"
    sSql = sSql & " Count(Distinct CCP.cPersCod) as NumCliente,1 as nIndice"
    sSql = sSql & " From ColocCalifProv  CCP"
    sSql = sSql & " Where CCP.nPrdEstado in(2020,2021,2022,2101,2104,2106,2107) and CCP.nDiasAtraso<=0 And SubString(CCP.cCtaCod,7,2)<>'21'"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'VIGENTE NORMAL (1-8)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & " End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & " End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,2 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2020,2021,2022,2101,2104,2106,2107) and  SubString(cCtaCod,7,2)<>'21' and (nDiasAtraso>=1 and nDiasAtraso<=8)"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'VIGENTE NORMAL (9-30)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,3 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2020,2021,2022,2101,2104,2106,2107) and SubString(cCtaCod,7,2)<>'21' and not (SubString(cCtaCod,6,2)='10' And nDiasAtraso>=16 and nDiasAtraso<=30) and (nDiasAtraso>=9 and nDiasAtraso<=30)"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'VENCIDA (COM) NORMAL (16-30)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,4 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2020,2021,2022,2101,2104,2106,2107) and SubString(cCtaCod,6,2)='10' and  (nDiasAtraso>=16 and nDiasAtraso<=30) and SubString(cCtaCod,7,2)<>'21'"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'VENCIDA NORMAL (31-60)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,5 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2020,2021,2022,2101,2104,2106,2107) and SubString(cCtaCod,7,2)<>'21' and (nDiasAtraso>=31 and nDiasAtraso<=60)"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'VENCIDA NORMAL (61-90)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & " End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,6 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2020,2021,2022) and SubString(cCtaCod,7,2)<>'21' and (nDiasAtraso>=61 and nDiasAtraso<=90)"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'VENCIDA NORMAL (91-120)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & " End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,7 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2020,2021,2022) and SubString(cCtaCod,7,2)<>'21' and (nDiasAtraso>=91 and nDiasAtraso<=120)"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'VENCIDA NORMAL (121-180)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,8 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2020,2021,2022) and SubString(cCtaCod,7,2)<>'21' and (nDiasAtraso>=121 and nDiasAtraso<=180)"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'VENCIDA NORMAL (>=181)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,9 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2020,2021,2022) and SubString(cCtaCod,7,2)<>'21' and nDiasAtraso>=181"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'REFINANCIADA (0)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,10 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2030,2031,2032) and SubString(cCtaCod,7,2)<>'21' and nDiasAtraso<=0"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'REFINANCIADA (1-8)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,11 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2030,2031,2032) and SubString(cCtaCod,7,2)<>'21' and (nDiasAtraso>=1 and nDiasAtraso<=8)"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'REFINANCIADA (9-30)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,12 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2030,2031,2032) and SubString(cCtaCod,7,2)<>'21'  and not (SubString(cCtaCod,6,2)='10' And nDiasAtraso>=16 and nDiasAtraso<=30) and (nDiasAtraso>=9 and nDiasAtraso<=30)"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'REFINANCIADA COM (16-30)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,13 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2030,2031,2032) and SubString(cCtaCod,6,2)='10' and (nDiasAtraso>=16 and nDiasAtraso<=30) and SubString(cCtaCod,7,2)<>'21'"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'REFINANCIADA (31-60)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,14 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2030,2031,2032) and SubString(cCtaCod,7,2)<>'21' and  (nDiasAtraso>=31 and nDiasAtraso<=60)"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'REFINANCIADA (61-90)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,15 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2030,2031,2032) and SubString(cCtaCod,7,2)<>'21' and (nDiasAtraso>=61 and nDiasAtraso<=90)"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'REFINANCIADA (91-120)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,16 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2030,2031,2032) and SubString(cCtaCod,7,2)<>'21' and (nDiasAtraso>=91 and nDiasAtraso<=120)"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'REFINANCIADA (121-180)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,17 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2030,2031,2032) and SubString(cCtaCod,7,2)<>'21' and (nDiasAtraso>=121 and nDiasAtraso<=180)"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'REFINANCIADA (>=181)' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,18 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2030,2031,2032) and SubString(cCtaCod,7,2)<>'21' and nDiasAtraso>=181"
    sSql = sSql & " Union"
    sSql = sSql & " Select 'JUDICIAL' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap*" & pnTipoCambio & "  End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & "  End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,19 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2201,2205) and SubString(cCtaCod,7,2)<>'21'"
    sSql = sSql & " ORDER BY NINDICE"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set GetCalidadCartera = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GetCalidadCarteraCastigados(ByVal pnTipoCambio As Double) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select 'JUDICIAL' as cDescripcion,"
    sSql = sSql & " SaldoK=Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nSaldoCap"
    sSql = sSql & " When '2' Then nSaldoCap* " & pnTipoCambio & " End),"
    sSql = sSql & " SaldoProv =Sum(Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then nProvision"
    sSql = sSql & " When '2' Then nProvision*" & pnTipoCambio & " End),"
    sSql = sSql & " Count(cCtaCod) as NumCreditos,Count(Distinct cPersCod) as NumCliente,19 as nIndice"
    sSql = sSql & " From ColocCalifProv"
    sSql = sSql & " Where nPrdEstado in(2202,2206) and SubString(cCtaCod,7,2)<>'21'"
    sSql = sSql & " ORDER BY NINDICE"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set GetCalidadCarteraCastigados = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
End Function

Public Function GetCalidadCarteraConsumo(ByVal pnTipoCambio As Double, ByVal psEstados As String, _
ByVal pnDiaIni As Integer, ByVal pnDiaFin As Integer) As Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select Sum (Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then (nSaldoCap-nCapVencido)"
    sSql = sSql & " When '2' Then (nSaldoCap-nCapVencido)* " & pnTipoCambio & " End)"
    sSql = sSql & " From [dbcmacicaconsol].[dbo].[CreditoConsolTotal]"
    sSql = sSql & " Where SubString(cCtaCod,6,3) like '[34]%' and nDiasAtraso>=" & pnDiaIni & " and nDiasAtraso<=" & pnDiaFin & " and"
    sSql = sSql & " nPrdEstado in(" & psEstados & ")"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set GetCalidadCarteraConsumo = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaCreditosMoraXTelefono(ByVal psCodAnalistas As String, ByVal pdFecha As Date, ByVal pnDiaInicio As Integer, _
                                           ByVal pnDiaFin As Integer, ByVal psCodAgen As String) As Recordset

    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "Select * "
    sSql = sSql & " From ( Select  RH.cUser,P.cCtaCod,"
    sSql = sSql & " Titular=(Select Pers.cPersNombre"
    sSql = sSql & " From ProductoPersona PP1"
    sSql = sSql & " Inner Join Persona Pers on PP1.cPersCod=Pers.cPersCod"
    sSql = sSql & " Where PP1.cCtaCod=P.cCtaCod and PP1.nPrdPersRelac=20),"
    sSql = sSql & " Direccion=(Select Pers.cPersDireccDomicilio"
    sSql = sSql & " From ProductoPersona PP1"
    sSql = sSql & " Inner Join Persona Pers on PP1.cPersCod=Pers.cPersCod"
    sSql = sSql & " Where PP1.cCtaCod=P.cCtaCod and PP1.nPrdPersRelac=20),"
    sSql = sSql & " KDesembolsado=(Select nMontoCol"
    sSql = sSql & " From Colocaciones"
    sSql = sSql & " Where cCtaCod=P.cCtaCod), P.nSaldo as SaldoK,"
    sSql = sSql & " NroCuotaApr=(Select Count(*)"
    sSql = sSql & " From ColocCalendario CD"
    sSql = sSql & " Where CD.cCtaCod=P.cCtaCod and CD.nColocCalendApl=1 and"
    sSql = sSql & " CD.nNroCalen=(Select Max(nNroCalen)"
    sSql = sSql & " From ColocCalendario Where cCtaCod=P.cCtaCod)),"
    sSql = sSql & " CuotasMora=(Select Count(*)"
    sSql = sSql & " From ColocCalendario CC"
    sSql = sSql & " Where CC.cCtaCod=P.cCtaCod and CC.nColocCalendApl=1 and CC.nNroCalen=(Select Max(nNroCalen)"
    sSql = sSql & " From ColocCalendDet"
    sSql = sSql & " Where cCtaCod=P.cCtaCod) and CC.nColocCalendEstado=0 and dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "') ,"
    sSql = sSql & " MontoCobrar=(Select Sum(CD.nMonto-CD.nMontoPagado)"
    sSql = sSql & " From ColocCalendario CC"
    sSql = sSql & " Inner Join ColocCalendDet CD on CC.cCtaCod=CD.cCtaCod And CC.nNroCalen=CD.nNroCalen and CC.nColocCalendApl=CD.nColocCalendApl and CC.nCuota = CD.nCuota"
    sSql = sSql & " Where CC.cCtaCod=P.cCtaCod and CD.nColocCalendApl=1 and CC.nNroCalen=(Select Max(nNroCalen)"
    sSql = sSql & " From ColocCalendDet"
    sSql = sSql & " Where cCtaCod=P.cCtaCod) and  CC.dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "' and CC.nColocCalendEstado=0),"
    sSql = sSql & " cRFA=Case When CC.cRFA Is Null or CC.cRFA='NOR' Then 'NOR' Else CC.cRFA End"
    sSql = sSql & " From RRHH RH"
    sSql = sSql & " Inner Join ProductoPersona PP on RH.cPersCod=PP.cPersCod and PP.nPrdPersRelac=28"
    sSql = sSql & " Inner Join Producto P on P.cCtaCod=PP.cCtaCod"
    sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=P.cCtaCod"
    sSql = sSql & " Inner Join ProductoPersona PP1 on PP1.cCtacod=P.cCtaCod and PP1.nPrdPersRelac=20"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP1.cPersCod"
    sSql = sSql & " Where P.nPrdEstado in (2020,2021,2022,2030,2031,2032)  and"
    sSql = sSql & " (Select Count(*)"
    sSql = sSql & " From ColocCalendario CC"
    sSql = sSql & " Where CC.cCtaCod=P.cCtaCod and CC.nColocCalendApl=1 and"
    sSql = sSql & " CC.nNroCalen=(Select Max(nNroCalen)"
    sSql = sSql & " From ColocCalendDet"
    sSql = sSql & " Where cCtaCod=P.cCtaCod) and CC.nColocCalendEstado=0 and dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "')>0 and"
    sSql = sSql & " RH.cPersCod in (" & psCodAnalistas & ") and"
    sSql = sSql & " Substring(P.cCtaCod,6,3) in ('101','102','103','201','202','301','302','303','304','320','401','403','423') and"
    sSql = sSql & " Substring(P.cCtaCod,4,2) in (" & psCodAgen & ") and CC.nColocCondicion in ('1','3','2') and"
    sSql = sSql & " Substring(P.cCtaCod,9,1) in (1,2) and CC.nDiasAtraso between " & pnDiaInicio & " and " & pnDiaFin & ")X Order By cUser,Titular"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCreditosMoraXTelefono = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Function ListaAnalistasAgencia(ByVal psCodAgencia As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim sAnalistas As String
    Dim nCont As Integer
    
    sSql = "SELECT  p.cperscod"
    sSql = sSql & " FROM RHCARGOS RC1, RRHH RH, PERSONA P, RHCONTRATO RT1,AREAS A,RHCARGOSTABLA RA, AGENCIAS AG"
    sSql = sSql & " Where RC1.cPersCod = RH.cPersCod"
    sSql = sSql & " AND RT1.cperscod =RC1.cperscod"
    sSql = sSql & " AND P.cperscod =  rc1.cperscod"
    sSql = sSql & " AND A.cAreaCod = RC1.cRHAreaCodOficial"
    sSql = sSql & " AND RA.cRHCargoCod = RC1.cRHCargoCodOficial"
    sSql = sSql & " AND RC1.cRHAgenciaCodOficial = AG.cAgeCod"
    sSql = sSql & " AND dRHCargoFecha = (SELECT MAX (dRHCargoFecha) FROM  RHCARGOS RC2 WHERE RC2.CPERSCOD = RC1.CPERSCOD)"
    sSql = sSql & " AND cRHContratoNro = (SELECT MAX (cRHContratoNro) FROM  RHCONTRATO RT2 WHERE RT2.CPERSCOD = RT1.CPERSCOD)"
    sSql = sSql & " AND  SUBSTRING(cRHCod,1,1) IN ('E','P')"
    sSql = sSql & " AND RH.NRHESTADO=201"
    sSql = sSql & " AND RC1.cRHAgenciaCod='" & psCodAgencia & "'"
    sSql = sSql & " AND NRHESTADO NOT LIKE  '8%' and substring(RA.cRHCargoCod,1,3) in ('004')"
    sSql = sSql & " order by  cAgeDescripcion,RC1.cRHCargoCod,nRHContratoTpo ASC  "
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    nCont = 0
    Do Until rs.EOF
        If nCont = 0 Then
            sAnalistas = "'" & rs!cPersCod & "'"
        Else
            sAnalistas = sAnalistas & ",'" & rs!cPersCod & "'"
        End If
        nCont = nCont + 1
        rs.MoveNext
    Loop
    Set rs = Nothing
    ListaAnalistasAgencia = sAnalistas
End Function

Public Function ListaCreditosConvenioCancaledos(ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, _
                                                 ByVal pnMoneda As String, _
                                                 ByVal psCodAgencia As String, Optional ByVal psCodInstitucion As String) As Recordset

    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim dFechaInicio As String
    Dim dFechaFin As String

    dFechaInicio = Format(pdFechaIni, "MM/dd/yyyy")
    dFechaFin = Format(pdFechaFin, "MM/dd/yyyy")
    
    sSql = "Select P.cCtaCod,CC.cCodModular,Pers.cPersNombre as cTitular,C.nMontoCol,"
    sSql = sSql & " Convert(Varchar(20),P.dPrdEstado,103) as dFecha,PersInstitucion.cPersNombre as cInstitucion"
    sSql = sSql & " From ColocacConvenio CC"
    sSql = sSql & " Inner Join Producto P on P.cCtaCod=CC.cCtaCod"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=P.cCtacod and PP.nPrdPersRelac=20"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Inner Join Colocaciones C on C.cCtaCod=P.cCtaCod"
    sSql = sSql & " Inner Join Persona PersInstitucion on PersInstitucion.cPersCod=CC.cPersCod"
    sSql = sSql & " Where P.nPrdEstado=2050 and (P.dPrdEstado Between '" & dFechaInicio & "' and '" & dFechaFin & "' ) "
    
    If Len(psCodInstitucion) > 0 Then
        sSql = sSql & " and CC.cPersCod='" & psCodInstitucion & "'"
    End If
    
    sSql = sSql & " and SubString(P.cCtaCod,4,2)='" & psCodAgencia & "'"
    sSql = sSql & " Order By PersInstitucion.cPersNombre,P.dPrdEstado "
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCreditosConvenioCancaledos = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GetPersona(ByVal psPersCod As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim sNombre As String
    Dim rs As ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "Select cPersNombre From Persona Where cPersCod='" & psPersCod & "'"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        GetPersona = rs!cPersNombre
    End If
    Set rs = Nothing
End Function

Public Function GetAgencia(ByVal psCodAgen As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim sNombre As String
    Dim rs As ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = " Select cAgeDescripcion From Agencias Where cAgeCod='" & psCodAgen & "'"
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        GetAgencia = rs!cAgeDescripcion
    End If
    Set rs = Nothing
End Function

'Public Function ObtenerCarteraVigentePorUserStored(ByVal psCodAgen As String, ByVal psUser As String, ByVal pdFecha As Date) As Recordset
'
' Dim oConec As COMConecta.DCOMConecta
' Dim nTipoCambio As Double
' Dim oDCredDoc As DCredDoc
'
' Set oDCredDoc = New DCredDoc
' nTipoCambio = oDCredDoc.ObtenerTipoCambio(pdFecha)
' Set oDCredDoc = Nothing
'
' Set oConec = New COMConecta.DCOMConecta
' oConec.AbreConexion
' Set ObtenerCarteraVigentePorUserStored = oConec.Ejecutar("ListaVigentes " & psCodAgen & ",'" & psUser & "', " & nTipoCambio)
' oConec.CierraConexion
' Set oConec = Nothing
'End Function

Public Function ListaCreditosXInstitucionDesembolsados(ByVal psCodAgen As String, ByVal psCodInstitucuion As String, _
                                                       ByVal pdFechaInicial As Date, ByVal dFechaFinal As Date) As Recordset

    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim sFechaInicial As String
    Dim sFechaFinal As String


    sFechaInicial = Format(pdFechaInicial, "YYYYMMdd")
    sFechaFinal = Format(dFechaFinal, "YYYYMMdd")
    
    sSql = "Select P.cCtaCod,C.cCodModular,Pers.cPersNombre,CC.nMontoCol,"
    sSql = sSql & " Fecha=(SubString(Left(M.cMovNro,8),7,2)+'/'+SubString(Left(M.cMovNro,8),5,2)+'/'+SubString(Left(M.cMovNro,8),1,4)),"
    sSql = sSql & " Numerocuota=(Select Count(CO.nCuota)"
    sSql = sSql & "                     From Colocaccred CC1 "
    sSql = sSql & "                           Inner Join ColocCalendario CO on CO.cCtaCod=CC1.cCtaCod and CC1.nNroCalen=CO.nNroCalen"
    sSql = sSql & "                     Where CO.cCtaCod=C.cCtaCod and CO.nColocCalendApl=1),"
    sSql = sSql & "        MontoCuota=(Select Sum(CO.nMonto)"
    sSql = sSql & "                    From Colocaccred CC1 "
    sSql = sSql & "                           Inner Join ColocCalendDet CO on CO.cCtaCod=CC1.cCtaCod and CC1.nNroCalen=CO.nNroCalen"
    sSql = sSql & "                     Where CO.cCtaCod=C.cCtaCod and CO.nColocCalendApl=1 and CO.nCuota=1),  "
    sSql = sSql & "       FechaVenc=(Select CO.dVenc"
    sSql = sSql & "                     From Colocaccred CC1 "
    sSql = sSql & "                           Inner Join ColocCalendario CO on CO.cCtaCod=CC1.cCtaCod and CC1.nNroCalen=CO.nNroCalen"
    sSql = sSql & "                     Where CO.cCtaCod=C.cCtaCod and CO.nColocCalendApl=1 and CO.nCuota=1)   "
    sSql = sSql & " From ColocacConvenio C"
    sSql = sSql & " Inner Join Producto P on C.cCtaCod=P.cCtaCod"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=P.cCtaCod and PP.nPrdPersRelac=20"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Inner Join Colocaciones CC on CC.cCtaCod=P.cCtaCod"
    sSql = sSql & " Inner Join MovColDet MD on MD.cCtaCod=C.cCtaCod"
    sSql = sSql & " Inner Join Mov M on M.nMovNro=MD.nMovNro"
    sSql = sSql & " Where M.nMovFlag=0 and (Left(M.cMovNro,8) between '" & sFechaInicial & "' and '" & sFechaFinal & "')"
    sSql = sSql & " and SubString(P.cCtaCod,4,2)='" & psCodAgen & "' and MD.cOpeCod in ('100101','100102','100103','100104','100105') and"
    sSql = sSql & " MD.cOpeCod not like '107[123456789]%' and MD.nPrdConceptoCod=1000 and "
    sSql = sSql & " C.cPersCod='" & psCodInstitucuion & "'"
    sSql = sSql & " Order By cast(SubString(Left(M.cMovNro,8),5,2)+'/'+SubString(Left(M.cMovNro,8),7,2)+'/'+SubString(Left(M.cMovNro,8),1,4) as Datetime)"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCreditosXInstitucionDesembolsados = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
        
End Function


Public Function ListaConsolidadoPorAnalistaV2(ByVal pnTipoCambio As Double, ByVal psAgenciaCod As String, ByVal pdFecha As Date, _
ByVal psAnalistas As String, ByVal pbFecha As Boolean) As Recordset

    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sFecha As String
    Dim sDia As String
    Dim sMes As String
    Dim sAno As String
    Dim sFechaMov As String
    Dim sdInicioMes As String
    Dim nTipoCambio As Double
    'Dim odCredDoc As DCredDoc
    
    sFechaMov = Format(pdFecha, "MM/dd/yyyy")
    sAno = DatePart("yyyy", pdFecha)
    sMes = Right("00", 2 - Len(CStr(DatePart("m", pdFecha)))) & CStr(DatePart("m", pdFecha))
    sDia = Right("00", 2 - Len(CStr(DatePart("d", pdFecha)))) & CStr(DatePart("d", pdFecha))
    
    'Set odCredDoc = New DCredDoc
    nTipoCambio = ObtenerTipoCambio(pdFecha)
    
    sFecha = sAno & sMes & sDia
    sdInicioMes = sAno & sMes & "01"
    
    If pbFecha = True Then ' fecha del sistema
    
            sSql = "SELECT  CCODANA, SUM(nNumCredNew) as nNumCredNew, SUM(nMonCredNew) as nMonCredNew, "
            sSql = sSql & "  SUM(nNumCredRest) as nNumCredRest, SUM(nMonCredRest) as nMonCredRest ,"
            sSql = sSql & "  SUM(nNumCredDes) as nNumCredDes, SUM(nMonCredDes) as nMonCredDes,"
            sSql = sSql & " SUM(NSALDO) as NSALDO,"
            sSql = sSql & " SUM(NNRO1_15) as NNRO1_15, SUM(NSALDO1_15) as NSALDO1_15,"
            sSql = sSql & " SUM(NNRO16_30) as NNRO16_30, SUM(NSALDO16_30) as NSALDO16_30,"
            sSql = sSql & " SUM(NNROMAY31) as NNROMAY31, SUM(NSALDOMAY31) as NSALDOMAY31,"
            sSql = sSql & " SUM(NNROJUD) as NNROJUD, SUM(NSALDOJUD) as NSALDOJUD,"
            sSql = sSql & " SUM(NNROVENCOM+NNROVEN) AS NNROVEN, SUM(NSALDOVENCOM+NSALDOVEN) AS NSALDOVEN"
            sSql = sSql & " FROM CONSTANTE C JOIN"
            sSql = sSql & " (SELECT  CCTACOD, SUBSTRING(CCTACOD,6,3) AS NPRODUCTO, CCODANA, NPRDESTADO,isnull(nNumCredNew,0) as nNumCredNew, isnull(nMonCredNew,0) as nMonCredNew,"
            sSql = sSql & " isnull(nNumCredRest,0) as nNumCredRest, isnull(nMonCredRest,0) as nMonCredRest, isnull(nNumCredDes,0) as nNumCredDes,"
            sSql = sSql & " isnull(nMonCredDes,0) as nMonCredDes, CASE WHEN SUBSTRING(CCTACOD,9,1) ='1' THEN NSALDO ELSE Round((NSALDO*" & pnTipoCambio & "),2) End as NSALDO,"
            sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  1 AND 15 AND nPrdEstado NOT IN (2201,2205) THEN 1 ELSE 0 END  AS NNRO1_15,"
            sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  1 AND 15 AND nPrdEstado NOT IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDO1_15 ,"
            sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  16 AND 30 AND nPrdEstado NOT IN (2201,2205) THEN 1 ELSE 0 END AS NNRO16_30 ,"
            sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  16 AND 30 AND nPrdEstado NOT IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDO16_30 ,"
            sSql = sSql & " CASE WHEN nDiasAtr>=31 AND nPrdEstado NOT IN (2201,2205) THEN 1 ELSE 0 END AS NNROMAY31 ,"
            sSql = sSql & " CASE WHEN nDiasAtr>=31 AND nPrdEstado NOT IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDOMAY31 ,"
            sSql = sSql & " CASE WHEN nPrdEstado IN (2201,2205) THEN 1 ELSE 0 END AS NNROJUD ,"
            sSql = sSql & " CASE WHEN nPrdEstado IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDOJUD ,"
            sSql = sSql & " CASE WHEN nDiasAtr>= 16 and SUBSTRING(CCTACOD,6,1)='1' THEN 1 ELSE 0 END AS NNROVENCOM ,"
            sSql = sSql & " CASE WHEN nDiasAtr>= 16 and SUBSTRING(CCTACOD,6,1)='1' THEN NSALDO ELSE 0 END AS NSALDOVENCOM,"
            sSql = sSql & " CASE WHEN nDiasAtr>= 31 and SUBSTRING(CCTACOD,6,1)<>'1' THEN 1 ELSE 0 END AS NNROVEN ,"
            sSql = sSql & " CASE WHEN nDiasAtr>= 31 and SUBSTRING(CCTACOD,6,1)<>'1' THEN NSALDO ELSE 0 END AS NSALDOVEN"
            sSql = sSql & " From"
            sSql = sSql & " (SELECT P.CCTACOD, P.NPRDESTADO,"
            sSql = sSql & " CASE WHEN CANA.CUSER IS NULL AND SUBSTRING(P.CCTACOD,6,3)='305' THEN 'PREN' ELSE CANA.CUSER END AS CCODANA,"
            sSql = sSql & " nNumCredNew = CASE WHEN C.nColocCondicion = 1 THEN nNumCredNew ELSE 0 END,"
            sSql = sSql & " nMonCredNew = CASE WHEN C.nColocCondicion = 1 THEN nMonCredNew ELSE 0 END,"
            sSql = sSql & " nNumCredRest = CASE WHEN C.nColocCondicion <> 1 OR SUBSTRING(P.CCTACOD,6,3)='305' THEN nNumCredNew ELSE 0 END,"
            sSql = sSql & " nMonCredRest = CASE WHEN C.nColocCondicion <> 1 OR SUBSTRING(P.CCTACOD,6,3)='305' THEN nMonCredNew ELSE 0 END,"
            sSql = sSql & " NDESEM.nNumCredDes , NDESEM.nMonCredDes, NSALDO = CASE WHEN SUBSTRING(P.CCTACOD,9,1) ='1' THEN P.NSALDO ELSE ROUND(P.NSALDO*" & pnTipoCambio & ",2) END,"
            sSql = sSql & " nDiasAtr= CASE WHEN SUBSTRING(P.CCTACOD,6,3)='305' THEN DATEDIFF(DAY, CC.dVenc, '" & sFechaMov & "') ELSE C.nDiasAtraso END"
            sSql = sSql & " FROM    PRODUCTO P"
            sSql = sSql & " JOIN COLOCACIONES CC ON CC.CCTACOD = P.CCTACOD"
            sSql = sSql & " LEFT JOIN COLOCACCRED C ON C.CCTACOD = CC.CCTACOD"
            sSql = sSql & " LEFT JOIN ( SELECT  cCtaCod, COUNT(MC.cCtaCod) nNumCredNew ,"
            sSql = sSql & " SUM(CASE WHEN SUBSTRING(CCTACOD,9,1) ='1' THEN NMONTO ELSE ROUND((NMONTO* " & nTipoCambio & "),2) END) AS nMonCredNew"
            sSql = sSql & " FROM    MOV M"
            sSql = sSql & " JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO"
            sSql = sSql & " WHERE MC.COPECOD IN (120201, 100101,100102, 100103,100104,100105)"
            sSql = sSql & " AND LEFT(M.CMOVNRO,8) BETWEEN '" & sdInicioMes & "' AND '" & sFecha & "'  AND M.NMOVFLAG =0"
            sSql = sSql & " GROUP BY cCtaCod ) AS NCREDNEW ON NCREDNEW.CCTACOD = P.CCTACOD"
            sSql = sSql & " LEFT JOIN ( SELECT  cCtaCod, COUNT(MC.cCtaCod) nNumCredDes ,"
            sSql = sSql & " SUM(CASE WHEN SUBSTRING(CCTACOD,9,1) ='1' THEN NMONTO ELSE ROUND((NMONTO*" & nTipoCambio & "),2) END) AS nMonCredDes"
            sSql = sSql & " FROM    MOV M"
            sSql = sSql & " JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO"
            sSql = sSql & " WHERE   MC.COPECOD IN (120201, 100101,100102, 100103,100104,100105) AND M.NMOVFLAG =0"
            sSql = sSql & " GROUP BY cCtaCod ) AS NDESEM ON NDESEM.CCTACOD = P.CCTACOD"
            sSql = sSql & " LEFT JOIN ( SELECT  R.CCTACOD , RH.CUSER"
            sSql = sSql & " FROM    PRODUCTOPERSONA R"
            sSql = sSql & " JOIN RRHH RH ON RH.CPERSCOD = R.CPERSCOD"
            sSql = sSql & " WHERE   R.nPrdPersRelac = 28 ) AS CANA ON CANA.CCTACOD = P.CCTACOD"
            'sSQL = sSQL & " WHERE   P.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205) AND P.NSALDO>0) AS NCRED ) AS NCRED1"
            sSql = sSql & " WHERE   P.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205)) AS NCRED ) AS NCRED1"
            sSql = sSql & " ON NCRED1.NPRODUCTO = C.NCONSVALOR AND C.NCONSCOD = 1001"
            'sSQL = sSQL & " WHERE   Substring(NCRED1.cCtacod,4,2) in(" & psAgenciaCod & ") AND  CCODANA in ('JMMM','CEMG'," & psAnalistas & ") AND Substring(NCRED1.cCtacod,6,3) in (" & psProducto & ")"
            sSql = sSql & " WHERE   Substring(NCRED1.cCtacod,4,2) in(" & psAgenciaCod & ") AND NOT EXISTS (SELECT CCTACOD FROM COLOCACCRED WHERE CCTACOD=NCRED1.CCTACOD AND CRFA IN ('RFA'))"
            sSql = sSql & " AND CCODANA='" & psAnalistas & "'"
            sSql = sSql & " GROUP BY CCODANA"
            sSql = sSql & " ORDER BY  CCODANA"
    Else
        sSql = "SELECT  CCODANA,"
        sSql = sSql & " SUM(nNumCredNew) as nNumCredNew, SUM(nMonCredNew) as nMonCredNew,"
        sSql = sSql & " SUM(nNumCredRest) as nNumCredRest, SUM(nMonCredRest) as nMonCredRest ,"
        sSql = sSql & " SUM(nNumCredDes) as nNumCredDes, SUM(nMonCredDes) as nMonCredDes, SUM(NSALDO) as NSALDO, SUM(NNRO1_15) as NNRO1_15,"
        sSql = sSql & " SUM(NSALDO1_15) as NSALDO1_15, SUM(NNRO16_30) as NNRO16_30, SUM(NSALDO16_30) as NSALDO16_30, SUM(NNROMAY31) as NNROMAY31,"
        sSql = sSql & " SUM(NSALDOMAY31) as NSALDOMAY31, SUM(NNROJUD) as NNROJUD, SUM(NSALDOJUD) as NSALDOJUD,"
        sSql = sSql & " SUM(NNROVENCOM+NNROVEN) AS NNROVEN, SUM(NSALDOVENCOM+NSALDOVEN) AS NSALDOVEN"
        sSql = sSql & " FROM CONSTANTE C"
        sSql = sSql & " JOIN (SELECT  CCTACOD, SUBSTRING(CCTACOD,6,3) AS NPRODUCTO, CCODANA, NPRDESTADO,isnull(nNumCredNew,0) as nNumCredNew,"
        sSql = sSql & " isnull(nMonCredNew,0) as nMonCredNew, isnull(nNumCredRest,0) as nNumCredRest, isnull(nMonCredRest,0) as nMonCredRest,"
        sSql = sSql & " isnull(nNumCredDes,0) as nNumCredDes, isnull(nMonCredDes,0) as nMonCredDes,"
        sSql = sSql & " CASE WHEN SUBSTRING(CCTACOD,9,1) ='1' THEN NSALDO ELSE Round((NSALDO*" & nTipoCambio & "),2) End as NSALDO,"
        sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  1 AND 15 AND nPrdEstado NOT IN (2201,2205) THEN 1 ELSE 0 END  AS NNRO1_15,"
        sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  1 AND 15 AND nPrdEstado NOT IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDO1_15 ,"
        sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  16 AND 30 AND nPrdEstado NOT IN (2201,2205) THEN 1 ELSE 0 END AS NNRO16_30 ,"
        sSql = sSql & " CASE WHEN nDiasAtr BETWEEN  16 AND 30 AND nPrdEstado NOT IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDO16_30 ,"
        sSql = sSql & " CASE WHEN nDiasAtr>=31 AND nPrdEstado NOT IN (2201,2205) THEN 1 ELSE 0 END AS NNROMAY31 ,"
        sSql = sSql & " CASE WHEN nDiasAtr>=31 AND nPrdEstado NOT IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDOMAY31 ,"
        sSql = sSql & " CASE WHEN nPrdEstado IN (2201,2205) THEN 1 ELSE 0 END AS NNROJUD ,"
        sSql = sSql & " CASE WHEN nPrdEstado IN (2201,2205) THEN NSALDO ELSE 0 END AS NSALDOJUD ,"
        sSql = sSql & " CASE WHEN nDiasAtr>= 16 and SUBSTRING(CCTACOD,6,1)='1' THEN 1 ELSE 0 END AS NNROVENCOM ,"
        sSql = sSql & " CASE WHEN nDiasAtr>= 16 and SUBSTRING(CCTACOD,6,1)='1' THEN NSALDO ELSE 0 END AS NSALDOVENCOM,"
        sSql = sSql & " CASE WHEN nDiasAtr>= 31 and SUBSTRING(CCTACOD,6,1)<>'1' THEN 1 ELSE 0 END AS NNROVEN ,"
        sSql = sSql & " CASE WHEN nDiasAtr>= 31 and SUBSTRING(CCTACOD,6,1)<>'1' THEN NSALDO ELSE 0 END AS NSALDOVEN"
        sSql = sSql & " FROM (SELECT CS.CCTACOD, CS.NPRDESTADO,"
        sSql = sSql & " CASE WHEN CANA.CUSER IS NULL AND SUBSTRING(CS.CCTACOD,6,3)='305' THEN 'PREN' ELSE CANA.CUSER END AS CCODANA,"
        sSql = sSql & " nNumCredNew = CASE WHEN C.nColocCondicion = 1 THEN nNumCredNew ELSE 0 END,"
        sSql = sSql & " nMonCredNew = CASE WHEN C.nColocCondicion = 1 THEN nMonCredNew ELSE 0 END,"
        sSql = sSql & " nNumCredRest = CASE WHEN C.nColocCondicion <> 1 OR SUBSTRING(CS.CCTACOD,6,3)='305' THEN nNumCredNew ELSE 0 END,"
        sSql = sSql & " nMonCredRest = CASE WHEN C.nColocCondicion <> 1 OR SUBSTRING(CS.CCTACOD,6,3)='305' THEN nMonCredNew ELSE 0 END,"
        sSql = sSql & " NDESEM.nNumCredDes , NDESEM.nMonCredDes, NSALDO = CASE WHEN SUBSTRING(CS.CCTACOD,9,1) ='1' THEN CS.NSALDOCAP ELSE ROUND(CS.NSALDOCAP*" & nTipoCambio & ",2) END,"
        sSql = sSql & " nDiasAtr = CS.nDiasAtraso"
        sSql = sSql & " FROM    COLOCACSALDO CS"
        sSql = sSql & " JOIN COLOCACIONES CC ON CC.CCTACOD = CS.CCTACOD"
        sSql = sSql & " LEFT JOIN COLOCACCRED C ON C.CCTACOD = CC.CCTACOD"
        sSql = sSql & " LEFT JOIN ( SELECT  cCtaCod, COUNT(MC.cCtaCod) nNumCredNew ,"
        sSql = sSql & " SUM(CASE WHEN SUBSTRING(CCTACOD,9,1) ='1' THEN NMONTO ELSE ROUND((NMONTO*" & nTipoCambio & "),2) END) AS nMonCredNew"
        sSql = sSql & " FROM    MOV M JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO"
        sSql = sSql & " WHERE MC.COPECOD IN (120201, 100101,100102, 100103,100104,100105) AND"
        sSql = sSql & " LEFT(M.CMOVNRO,8) BETWEEN '" & sdInicioMes & "' AND '" & sFecha & "'  AND M.NMOVFLAG =0"
        sSql = sSql & " GROUP BY cCtaCod ) AS NCREDNEW ON NCREDNEW.CCTACOD = CS.CCTACOD"
        sSql = sSql & " LEFT JOIN ( SELECT  cCtaCod, COUNT(MC.cCtaCod) nNumCredDes ,"
        sSql = sSql & " SUM(CASE WHEN SUBSTRING(CCTACOD,9,1) ='1' THEN NMONTO ELSE ROUND((NMONTO*" & nTipoCambio & "),2) END) AS nMonCredDes"
        sSql = sSql & " FROM    MOV M"
        sSql = sSql & " JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO"
        sSql = sSql & " WHERE   MC.COPECOD IN (120201, 100101,100102, 100103,100104,100105) AND M.NMOVFLAG =0 GROUP BY cCtaCod ) AS NDESEM ON NDESEM.CCTACOD = CS.CCTACOD"
        sSql = sSql & " LEFT JOIN ( SELECT  R.CCTACOD , RH.CUSER"
        sSql = sSql & " FROM    PRODUCTOPERSONA R JOIN RRHH RH ON RH.CPERSCOD = R.CPERSCOD WHERE   R.nPrdPersRelac = 28 ) AS CANA ON CANA.CCTACOD = CS.CCTACOD"
        'sSQL = sSQL & " WHERE   CS.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205) AND CS.NSALDOCAP>0 AND CS.DFECHA=' " & CStr(Format(pdFecha, "mm/dd/yyyy")) & " ' ) AS NCRED ) AS NCRED1 ON NCRED1.NPRODUCTO = C.NCONSVALOR AND C.NCONSCOD = 1001"
        sSql = sSql & " WHERE   CS.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205)AND CS.DFECHA=' " & CStr(Format(pdFecha, "mm/dd/yyyy")) & " ' ) AS NCRED ) AS NCRED1 ON NCRED1.NPRODUCTO = C.NCONSVALOR AND C.NCONSCOD = 1001"
        sSql = sSql & " WHERE   Substring(NCRED1.cCtacod,4,2) in(" & psAgenciaCod & ") AND NOT EXISTS (SELECT CCTACOD FROM COLOCACCRED WHERE CCTACOD=NCRED1.CCTACOD AND CRFA IN ('RFA'))"
        sSql = sSql & " AND CCODANA='" & psAnalistas & "'"
        
        sSql = sSql & " GROUP BY CCODANA ORDER BY  CCODANA"

    End If
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
        Set ListaConsolidadoPorAnalistaV2 = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerParametroComision() As Double
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select nParamValor"
    sSql = sSql & " From ColocParametro"
    sSql = sSql & " Where nParamVar=9002" '102731"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        ObtenerParametroComision = rs!nParamValor
    End If
    Set rs = Nothing
End Function

Public Function ListaCredNoDesembolsados(ByVal psCodAgen As String, ByVal pdFechaIni As Date, _
ByVal pdFechaFin As Date) As Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sFechaIni As String
    Dim sFechaFin As String
    
    sFechaIni = Format("MM/dd/yyyy", pdFechaIni)
    sFechaFin = Format("MM/dd/yyyy", pdFechaFin)
    
    sSql = "Select P.cCtaCod,Pers.cPersNombre,CE.nMonto,Moneda=Case When SubString(P.cCtaCod,9,1)='1' Then 'Soles' Else 'Dolares' End,"
    sSql = sSql & " Pers.cPersTelefono as cTelefono,Pers.cPersDireccDomicilio as cDireccion,RH.cUser"
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=20"
    sSql = sSql & " Inner Join ProductoPersona PP1 on P.cCtaCod=PP1.cCtaCod and PP1.nPrdPersRelac=28"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Inner Join ColocacEstado CE on CE.cCtaCod=P.cCtaCod and CE.nPrdEstado=2002"
    sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP1.cPersCod"
    sSql = sSql & "Where SubString(P.cCtaCod,4,2) in ('01') and   (CE.dPrdEstado between '" & sFechaIni & " 00:00:00' and '" & sFechaFin & " 23:59:59')   and"
    sSql = sSql & " P.nPrdEstado = 2002"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCredNoDesembolsados = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function Repo_108506(ByVal pdFechaInicial As Date, ByVal pdFechaPagoInicial As Date, ByVal pdFechaPagoFin As Date) As Recordset
    Dim sFechaInicial As String
    Dim sFechaInicialPag As String
    Dim sFechaFinalPag As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    
    sFechaInicial = Format(pdFechaInicial, "MM/dd/yyyy")
    sFechaInicialPag = Format(pdFechaPagoInicial, "YYYYMMDD")
    sFechaFinalPag = Format(pdFechaPagoFin, "YYYYMMDD")
    
    sSql = "Select CN.cConsDescripcion,AG.cAgeDescripcion,CINI.cRFA,ISNULL(CINI.nMontoIni,0) as nMontoIni,ISNULL(CFIN.nMontoFin,0) as nMontoFin,nDIF=ISNULL(CINI.nMontoIni,0)-ISNULL(CFIN.nMontoFin,0) "
    sSql = sSql & " From ("
    sSql = sSql & "     Select SubString(CS.cCtaCod,6,3) as cProducto,SubString(CS.cCtaCod,4,2) as cAgen,C.cRFA,SUM(CS.nSaldoCap) as nMontoIni"
    sSql = sSql & "     From ColocacCred C"
    sSql = sSql & "          Inner Join ColocacSaldo CS on C.cCtaCod=CS.cCtaCod "
    sSql = sSql & "     Where C.cRFA in ('DIF','RFC') and CS.dFecha='" & sFechaInicial & "'"
    sSql = sSql & "     Group By SubString(CS.cCtaCod,6,3),SubString(CS.cCtaCod,4,2),C.cRFA) CINI"
    sSql = sSql & " "
    
    sSql = sSql & " Full Outer Join (Select SubString(MD.cCtaCod,6,3) as cProducto,SubString(MD.cCtaCod,4,2) as cAgen,C.cRFA,SUM(MD.nMonto)as nMontoFin"
    sSql = sSql & "         From Mov M"
    sSql = sSql & "                  Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
    sSql = sSql & "                  Inner Join ColocacCred C on C.cCtaCod=MD.cCtaCod"
    sSql = sSql & "             Where  C.cRFA in ('DIF','RFC') and M.nMovFlag=0 and (Left(M.cMovNro,8) between '" & sFechaInicialPag & "'and '" & sFechaFinalPag & "') and"
    sSql = sSql & "                    MD.nPrdConceptoCod in(1000,1109)  "
    sSql = sSql & "            Group By SubString(MD.cCtaCod,6,3),SubString(MD.cCtaCod,4,2),C.cRFA) CFIN on CINI.cProducto=CFIN.cProducto and CINI.cAgen=CFIN.cAgen and CINI.cRFA=CFIN.cRFA"
    sSql = sSql & " Inner Join Constante CN on CN.nConsValor=CINI.cProducto and CN.nConsCod=1001 and CN.nConsValor<>1001"
    sSql = sSql & " Inner Join Agencias AG on AG.cAgeCod=CINI.cAgen "
    sSql = sSql & " Order By AG.cAgeDescripcion,CN.cConsDescripcion,CINI.cRFA"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set Repo_108506 = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ListaCredAnaPlusMonto(ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, psCodAgen As String, _
pnTipoCambio As Double, ByVal pnMonto As Double) As Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sFechaIni As String
    Dim sFechaFin As String
    
    sFechaIni = Format(pdFechaIni, "yyyymmdd")
    sFechaFin = Format(pdFechaFin, "yyyymmdd")
    
    
    sSql = "Select Count(MD.cCtaCod) as nCantidad,RH.cUser"
    sSql = sSql & " From Mov M"
    sSql = sSql & " Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=MD.cCtaCod and PP.nPrdPersRelac=28"
    sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
    sSql = sSql & " inner Join ColocacCred C on C.cCtaCod=MD.cCtaCod"
    sSql = sSql & " Where SubString(MD.cCtaCod,4,2)='" & psCodAgen & "' and MD.nPrdConceptoCod=1000 and MD.cOpeCod in ('100101','100102','100103','100104','100105') and"
    sSql = sSql & " C.nColocCondicion=1  and (Case When SubString(MD.cCtaCod,9,1)='1' Then MD.nMonto Else MD.nMonto*" & pnTipoCambio & "End)<=" & pnMonto & " and"
    sSql = sSql & " (Left(M.cMovNro,8) between '" & sFechaIni & "' and '" & sFechaFin & "') and M.nMovFlag=0"
    sSql = sSql & " Group By RH.cUser"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCredAnaPlusMonto = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaCredAnaLessMonto(ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, psCodAgen As String, _
pnTipoCambio As Double, ByVal pnMonto As Double) As Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sFechaIni As String
    Dim sFechaFin As String
    
    sFechaIni = Format(pdFechaIni, "yyyymmdd")
    sFechaFin = Format(pdFechaFin, "yyyymmdd")
    
    
    sSql = "Select Count(MD.cCtaCod) as nCantidad,RH.cUser"
    sSql = sSql & " From Mov M"
    sSql = sSql & " Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=MD.cCtaCod and PP.nPrdPersRelac=28"
    sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
    sSql = sSql & " inner Join ColocacCred C on C.cCtaCod=MD.cCtaCod"
    sSql = sSql & " Where SubString(MD.cCtaCod,4,2)='" & psCodAgen & "' and MD.nPrdConceptoCod=1000 and MD.cOpeCod in ('100101','100102','100103','100104','100105') and"
    sSql = sSql & " C.nColocCondicion=1  and (Case When SubString(MD.cCtaCod,9,1)='1' Then MD.nMonto Else MD.nMonto*" & pnTipoCambio & "End)>" & pnMonto & " and"
    sSql = sSql & " (Left(M.cMovNro,8) between '" & sFechaIni & "' and '" & sFechaFin & "') and M.nMovFlag=0"
    sSql = sSql & " Group By RH.cUser"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCredAnaLessMonto = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaCreditosRefinanciados(ByVal pdFechaFin As Date, ByVal psCodAgen As String, _
                 ByVal pnTipoCambio As Double) As Recordset
                 
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    
    
    sSql = "Select Sum(Case When SubString(CS.cCtaCod,9,1)='1' Then CS.nSaldoCap Else CS.nSaldoCap*" & pnTipoCambio & "End),RH.cUser"
    sSql = sSql & " From ColocacSaldo CS"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtacod=CS.cCtaCod and PP.nPrdPersRelac=28"
    sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
    sSql = sSql & " Where SubString(CS.cCtacod,4,2)='" & psCodAgen & "' and CS.nPrdEstado in (2030,2031,2032) and CS.dFecha='" & Format(pdFechaFin, "MM/dd/yyyy") & "'"
    sSql = sSql & " Group By RH.cUser"
    sSql = sSql & " Order By RH.cUser"
        
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCreditosRefinanciados = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtieneCuentasPersona(ByVal psCodPers As String, ByVal psEstados As String, ByVal psMoneda As String) As ADODB.Recordset
Dim SQL As String
Dim oConec As COMConecta.DCOMConecta

On Error GoTo ErrorObtieneCuentasPersona
    'Sql = " Select Pro.cCtaCod, P.cPersNombre from ProductoPersona PP"
    'Sql = Sql & " Inner Join Persona P on P.cPersCod  = PP.cPersCod"
    'Sql = Sql & " Inner Join Producto Pro on Pro.cCtaCod = PP.cCtaCod"
    'Sql = Sql & " Where PP.nPrdPersRelac = 10 And Pro.nPrdEstado in (" & psEstados & ")"
    'Sql = Sql & " and P.cPersCod = '" & psCodPers & "' and substring(Pro.cCtaCod,9,1)='" & psMoneda & "'"
    SQL = " select Pro.cCtaCod, nSaldo from ProductoPersona PP"
    SQL = SQL & " Inner Join Producto Pro on Pro.cCtaCod = PP.cCtaCod"
    SQL = SQL & " where cperscod='" & psCodPers & "'"
    SQL = SQL & " and nPrdPersRelac=10 "
    SQL = SQL & " and nPrdEstado in (" & psEstados & ") "
    SQL = SQL & " and substring(Pro.cCtaCod,9,1)='" & psMoneda & "' Order by Pro.cCtaCod"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtieneCuentasPersona = oConec.CargaRecordSet(SQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
    Exit Function

ErrorObtieneCuentasPersona:
    Err.Raise Err.Number, "Obtiene Cuentas Persona ", Err.Description

End Function

Public Function CargaDatosCuentaCAutomatico(ByVal psCtaCod As String, _
                                            ByRef prsCred As ADODB.Recordset, _
                                            ByRef prsCol As ADODB.Recordset, _
                                            ByRef prsCue As ADODB.Recordset)

Dim oCred As COMDCredito.DCOMCredDoc
Set oCred = New COMDCredito.DCOMCredDoc
Dim sPersCod As String

On Error GoTo ErrorCargaDatosCuentaCAutomatico

        Set prsCred = oCred.RecuperaDatosMantCargoAutomatico(psCtaCod)
        If prsCred.EOF And prsCred.BOF Then
            Set oCred = Nothing
            Exit Function
        End If
        sPersCod = prsCred!cPersCod
        Set prsCol = oCred.RecuperaDatosColocCargoAutoma(psCtaCod)
        
        Set prsCue = oCred.ObtieneCuentasPersona(sPersCod, "1000", IIf(Mid(psCtaCod, 9, 1) = "1", "1", "2"))

    Set oCred = Nothing
    Exit Function
ErrorCargaDatosCuentaCAutomatico:
    Err.Raise Err.Number, "Carga Datos Cuenta Cargo Automatico", Err.Description

End Function

Public Function ActualizaColocCargoAutomaLote(ByVal psCtaCod As String, _
                                            ByVal pMatCargoAutoma As Variant)

Dim i As Integer

On Error GoTo ErrorActualizaColocCargoAutomaLote

    For i = 1 To UBound(pMatCargoAutoma)
        If pMatCargoAutoma(i, 1) <> "" Then
            If ExisteCtaCargoAutoma(psCtaCod, pMatCargoAutoma(i, 2)) = True Then
                Call ActualizaColocCargoAutoma(psCtaCod, pMatCargoAutoma(i, 2), pMatCargoAutoma(i, 3), IIf(pMatCargoAutoma(i, 4) = ".", 1, 0))
            Else
                Call InsertaColocCargoAutoma(psCtaCod, pMatCargoAutoma(i, 2), IIf(pMatCargoAutoma(i, 4) = ".", "1", "0"), pMatCargoAutoma(i, 3))
            End If
        End If
    Next i
    Exit Function
ErrorActualizaColocCargoAutomaLote:
    Err.Raise Err.Number, "Actualiza Cargo Automatico Lote", Err.Description

End Function

'Datos para la Identificacion RCC
Public Function ObtieneDatosIdentificacionRCC(ByVal psCtaCod As String) As ADODB.Recordset
Dim SQL As String
Dim oConecta As COMConecta.DCOMConecta

SQL = "SELECT P.cPersCod,Titular=P.cPersNombre,Documento=TipoDoc.cConsDescripcion,NroDoc=Doc.cPersIDNro, A.cAgeDescripcion, CIIU.cCIIUDescripcion,E.nMonto,E.nPlazo,Analista= An.cPersNombre," & _
        " Moneda=CASE WHEN SUBSTRING(C.cCtaCod,9,1)='2' THEN 'US $' ELSE 'S/.'END, " & _
        " IMN=(nPersFIVentas+nPersFIRecupCtasXCobrar+nPersIngFam),EMN=(nPersFICostoVentas+nPersEgrFam+nPersFIEgresosOtros), " & _
        " Saldo=(nPersFIVentas+nPersFIRecupCtasXCobrar+nPersIngFam)-(nPersFICostoVentas+nPersEgrFam+nPersFIEgresosOtros), " & _
        " MontoCuotaDol=CASE WHEN SUBSTRING(C.cCtaCod,9,1)='2' THEN  dbo.CuotaAprobada(C.cCtaCod)   ELSE dbo.CuotaAprobada(C.cCtaCod)/nValFijoDia END," & _
        " Patrimonio=nPersFIPatrimonio,    DeudaSF=CASE WHEN SUBSTRING(C.cCtaCod,9,1)='2' THEN E.nMonto ELSE E.nMonto/nValFijoDia END, " & _
        " PosicionCambio=CASE WHEN (nPersFIActivoDisp+nPersFICtasxCobrar+nPersFIInventarios+nPersFIActivosFijos)-(nPersFIProveedores+nPersFICredCMACT+nPersFICredOtros)>0 THEN 'Sobre Compra' ElSE 'Sobre Venta'END," & _
        " PosicionCambioValor=ABS(    (nPersFIActivoDisp+nPersFICtasxCobrar+nPersFIInventarios+nPersFIActivosFijos) -(nPersFIProveedores+nPersFICredCMACT+nPersFICredOtros)  ), " & _
        " TotalActivo=(nPersFIActivoDisp+nPersFICtasxCobrar+nPersFIInventarios+nPersFIActivosFijos)," & _
        " MontoCuotaSol=CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN dbo.CuotaAprobada(C.cCtaCod) ELSE dbo.CuotaAprobada(C.cCtaCod)*nValFijoDia END, " & _
        " Cuota_Saldo=(CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN dbo.CuotaAprobada(C.cCtaCod) ELSE dbo.CuotaAprobada(C.cCtaCod)*nValFijoDia END)/ " & _
        " ((nPersFIVentas+nPersFIRecupCtasXCobrar+nPersIngFam)-(nPersFICostoVentas+nPersEgrFam+nPersFIEgresosOtros))," & _
        " DeudaSoles= CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN E.nMonto ELSE E.nMonto*nValFijoDia END, TipoCambio=nValFijoDia," & _
        " Deuda_Patrimonio=(CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN E.nMonto ELSE E.nMonto*nValFijoDia END)/nPersFIPatrimonio, " & _
        " Expuesto= CASE WHEN ((nPersFIVentas+nPersFIRecupCtasXCobrar+nPersIngFam)-(nPersFICostoVentas+nPersEgrFam+nPersFIEgresosOtros)) / " & _
        "    (CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN dbo.CuotaAprobada(C.cCtaCod) ELSE dbo.CuotaAprobada(C.cCtaCod)*nValFijoDia END)>100 THEN 1 ELSE 0 END " & _
        " FROM ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod=PP.cCtaCod AND PP.nPrdPersRelac= " & gColRelPersTitular & _
        " INNER JOIN ColocacEstado E ON C.cCtaCod=E.cCtaCod AND E.dPrdEstado= (SELECT MAX(dPrdEstado)FROM ColocacEstado WHERE cCtaCod=C.cCtaCod) " & _
        " INNER JOIN Persona P ON PP.cPersCod=P.cPersCod INNER JOIN PersID Doc ON P.cPersCod=DOC.cPersCod AND DOC.cPersIDTpo=(SELECT MAX(cPersIDTpo)FROM PersID WHERE cPersCod=P.cPersCod)" & _
        " INNER JOIN Constante TipoDoc ON nConsValor=DOC.cPersIDTpo AND nConsCod=" & gPersIdTipo & " INNER JOIN Agencias A ON SUBSTRING(C.cCtaCod,4,2)=A.cAgeCod " & _
        " INNER JOIN ProductoPersona PPA ON C.cCtaCod=PPA.cCtaCod AND PPA.nPrdPersRelac=" & gColRelPersAnalista & _
        " INNER JOIN Persona An ON  An.cPersCod=PPA.cPersCod INNER JOIN CIIU ON P.cPersCIIU=CIIU.cCIIUCod " & _
        " INNER JOIN ColocFteIngreso CF ON CF.cCtaCod=C.cCtaCod INNER JOIN PersFteIngreso F ON P.cPersCod=F.cPersCod AND F.cNumFuente =CF.cNumFuente" & _
        " INNER JOIN PersFIIndependiente FI ON F.cNumFuente=FI.cNumFuente INNER JOIN TipoCambio TC ON TC.dFecCamb=(SELECT MAX(dFecCamb)FROM TipoCambio)" & _
        " WHERE C.cCtaCod='" & psCtaCod & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosIdentificacionRCC = oConecta.CargaRecordSet(SQL)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

'ARCV 28-10-2006
Public Function RecuperaDatosColocacConvenio(psTipoPlanilla) As ADODB.Recordset
Dim SQL As String
Dim Sql2 As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

Sql2 = "Select * from  CredTxtEducacion"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set R = oConecta.CargaRecordSet(Sql2)


SQL = "SELECT cCodModular=ISNULL(cCodModular,''),cCargo=ISNULL(cCargo,''),cCARBEN=ISNULL(cCARBEN,''),cT_Plani=ISNULL(cT_Plani,'')," & _
    " Per.cPersNombre , nMontoDes = Sum(CD.nMonto - CD.nMontoPagado) " & _
    " FROM ColocacConvenio CV " & _
    " INNER JOIN Producto P ON CV.cCtaCod=P.cCtaCod " & _
    " INNER JOIN ProductoPersona PP ON P.cCtaCod=PP.cCtacod AND PP.nPrdPersRelac=20 " & _
    " INNER JOIN Persona Per ON PP.cPersCod=Per.cPersCod " & _
    " INNER JOIN ColocacCred CC ON P.cCtaCod=CC.cCtaCod " & _
    " INNER JOIN ColocCalendDet CD ON CD.cCtaCod= P.cCtaCod AND CD.nNroCalen=CC.nNroCalen " & _
    " AND nColocCalendApl=1 AND CD.nCuota=CC.nNroProxCuota " & _
    " WHERE nPrdEstado IN (2020,2021,2022,2030,2031,2032) and substring(CV.cctacod,4,2)='" & R!cAgeCod & "'" & _
    " AND cT_Plani='" & psTipoPlanilla & "' AND CV.cPersCod='" & R!cPersCodInst & "'" & _
    " GROUP BY cCodModular,cCargo,cCARBEN,cT_Plani,Per.cPersNombre "


R.Close
'1060100890510 codigo de PErsona de EDUCACION---- OJO

Set RecuperaDatosColocacConvenio = oConecta.CargaRecordSet(SQL)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

Public Function RecuperaActaComiteCredApro( _
ByVal pdCodAge As String, _
ByVal pdFecIni As Date, _
ByVal pdFecFin As Date, _
ByVal pMatProd As Variant) As ADODB.Recordset

'psCodAge, pdFecIni, pdFecFin, pMatProd


' PEAC 20070822
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta 'DConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sAnalistas As String
'Dim oGen As COMDConstSistema.DCOMGeneral 'DGeneral

'    sCadProd = " And SUBSTRING(CE.cCtaCod,6,3) in ('"
'    For i = 0 To UBound(pMatProd) - 1
'        sCadProd = sCadProd & pMatProd(i) & "','"
'    Next i
'    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    sCadProd = " '0"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & "," & pMatProd(i)
    Next i
    sCadProd = sCadProd & "'"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
'sSql = "select  "
'sSql = sSql & " A.cAgeDescripcion,"
'sSql = sSql & " ce.cctacod,"
'sSql = sSql & " Producto=(select cConsDescripcion"
'sSql = sSql & "             From Constante"
'sSql = sSql & "             where nConsCod='1001' and nConsValor=substring(ce.cctacod,6,3)),"
'sSql = sSql & " Cliente = ISNULL((SELECT    cPersNombre"
'sSql = sSql & "                     FROM Persona Per"
'sSql = sSql & "                     INNER JOIN ProductoPersona PP ON Per.cPersCod = PP.cPersCod"
'sSql = sSql & "                     WHERE PP.cCtaCod = ce.cCtaCod AND nPrdPersRelac = 20),''),"
'sSql = sSql & " ce.dprdestado,"
'sSql = sSql & " ce.nmonto,"
'sSql = sSql & " ce.ncuotas,"
'sSql = sSql & " Co.cConsDescripcion,"
'sSql = sSql & " plazo ="
'sSql = sSql & "         (case"
'sSql = sSql & "               when CE.nColocCalendCod between '10' and '39' then ce.nPlazo"
'sSql = sSql & "               when CE.nColocCalendCod between '40' and '69' then ce.nPeriodoFechaFija"
'sSql = sSql & "           end),"
'sSql = sSql & " clc.cDescripcion as Descri_Linea,"
'sSql = sSql & " Analista = Isnull ((SELECT  cUser"
'sSql = sSql & "                     FROM rrhh rh"
'sSql = sSql & "                     INNER JOIN ProductoPersona PP ON rh.cPersCod = PP.cPersCod"
'sSql = sSql & "                     WHERE PP.cCtaCod = ce.cCtaCod AND pp.nPrdPersRelac = 28),'')"
'sSql = sSql & " from colocacEstado ce"
'sSql = sSql & "     inner join colocaciones c on ce.cctacod=c.cctacod"
'sSql = sSql & "     inner join coloclineacredito clc on c.cLineaCred=clc.cLineaCred"
'sSql = sSql & "     inner join Constante co on CE.nColocCalendCod=Co.nConsValor and Co.nConsCod=3005"
'sSql = sSql & "     inner join agencias a on substring(ce.cctacod,4,2)=a.cAgeCod"
'sSql = sSql & " where ce.nPrdEstado='2002' "
'sSql = sSql & " and ce.dprdestado between '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "'"
'sSql = sSql & " and substring(ce.cctacod,4,2) = " & pdCodAge
'sSql = sSql & "" & sCadProd
'sSql = sSql & " order by Producto, ce.dprdestado"
sSql = "exec stp_sel_ReporteActasComitedeCreditos '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "'," & sCadProd & ",'" & pdCodAge & "'," & UBound(pMatProd)
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaActaComiteCredApro = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
'madm 20100210 ----------------------------------
Public Function RecuperaActaComiteCredApro1( _
ByVal pdCodAge As Integer, _
ByVal pdFecIni As Date, _
ByVal pMatProd As String, pid_comite As Integer) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sAnalistas As String
    
sCadProd = "'" & pMatProd & "'"
sSql = "exec stp_sel_ReporteActasComitedeCreditos_1 '" & Format(pdFecIni, "yyyymmdd") & "'," & sCadProd & "," & pdCodAge & "," & pid_comite
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaActaComiteCredApro1 = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
'-------------------------------------------------------
Public Function RecuperaRepXTipoCondProd( _
ByVal pdCodAge As String, ByVal pdFecIni As Date, _
ByVal pdFecFin As Date, _
ByVal pMatProd As Variant, _
ByVal pTipoCambio As Double, _
ByVal psMonedas As String, _
ByVal pMatAnalistas As Variant) As ADODB.Recordset

'psCodAge, pdFecIni, pdFecFin, pMatProd, pnTipCam, psMoneda,pMatAnalistas

' PEAC 20070822
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta 'DConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sCadAnalista As String
Dim sMonedas As String

    sCadAnalista = "('"
    For i = 0 To UBound(pMatAnalistas) - 1
        sCadAnalista = sCadAnalista & pMatAnalistas(i) & "','"
    Next i
    sCadAnalista = Mid(sCadAnalista, 1, Len(sCadAnalista) - 2) & ")"

    sCadProd = " And SUBSTRING(C.cCtaCod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If

    If Len(Trim(psMonedas)) = 0 Then
        sMonedas = ""
    Else
        sMonedas = " AND SUBSTRING(c.cctacod,9,1) IN (" & psMonedas & ") "
    End If


sSql = " select"
sSql = sSql & " c.cCtaCod,"
sSql = sSql & " cCondiProd = co.cconsDescripcion ,"
sSql = sSql & " cTitular=(select cPersNombre from persona where cpersCod=ppTit.cPersCod),"
sSql = sSql & " cMoneda=CASE when substring(c.cCtaCod,9,1)='1' then 'MN' ELSE 'ME' END,"
sSql = sSql & " c.nMontoCol,"
sSql = sSql & " nMontoColSol = case when substring(c.cCtaCod,9,1) = '1' then c.nMontoCol ELSE c.nMontoCol*" & pTipoCambio & " END,"
sSql = sSql & " c.dVigencia,"
sSql = sSql & " cAnalista=(select cUser from rrhh where cpersCod=ppAna.cPersCod)"
sSql = sSql & " from colocaciones c"
sSql = sSql & " inner join producto p on p.cctacod=c.cctacod"
sSql = sSql & " inner JOIN colocacCred cc on p.cctacod=cc.cctacod"
sSql = sSql & " inner join constante co on cc.nColocCondicion = co.nConsVAlor and co.nConsCod=3015"
sSql = sSql & " inner join productopersona ppTit on c.cctacod=ppTit.cctacod and ppTit.nPrdPersRelac=20"
sSql = sSql & " inner join productopersona ppAna on c.cctacod=ppAna.cctacod and ppAna.nPrdPersRelac=28 AND ppAna.cPersCod in " & sCadAnalista
sSql = sSql & " Where p.nPrdEstado = 2020"
sSql = sSql & " and (c.dVigencia between '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "')"
sSql = sSql & " and substring(c.cctacod,4,2)=" & pdCodAge
sSql = sSql & "" & sCadProd
sSql = sSql & "" & sMonedas
sSql = sSql & " order by co.cconsDescripcion,c.dVigencia"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaRepXTipoCondProd = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaRepDetalleCuotasXCobrar( _
ByVal pdCodAge As String, _
ByVal pdFecIni As Date, _
ByVal pdFecFin As Date, _
ByVal pdFecSis As Date, _
ByVal pMatProd As Variant, _
ByVal psMonedas As String, _
ByVal pMatAnalistas As Variant) As ADODB.Recordset

'psCodAge, pdFecIni, pdFecFin, pdFecSis, pMatProd, psMoneda, pMatAnalistas

' PEAC 20070822
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta 'DConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sCadAnalista As String
Dim sMonedas As String

    sCadAnalista = "('"
    For i = 0 To UBound(pMatAnalistas) - 1
        sCadAnalista = sCadAnalista & pMatAnalistas(i) & "','"
    Next i
    sCadAnalista = Mid(sCadAnalista, 1, Len(sCadAnalista) - 2) & ")"

    sCadProd = " And SUBSTRING(a.cCtacod,6,3) in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If

    If Len(Trim(psMonedas)) = 0 Then
        sMonedas = ""
    Else
        sMonedas = " AND SUBSTRING(a.cCtaCod,9,1) IN (" & psMonedas & ") "
    End If

sSql = " select distinct"
sSql = sSql & " moneda=case substring(a.cCtacod,9,1) when 1 then 'NUEVOS SOLES' ELSE 'DOLARES US' END,"
sSql = sSql & " Analista=h.cUser,Credito=a.cCtaCod,Cliente=d.cPersNombre,Desembolso=e.nMontoCol,"
sSql = sSql & " Cuota=b.ncuota,"
sSql = sSql & " SaldoCapital=isnull((select sum(nmonto) from ColocCalendDet"
sSql = sSql & "         Where cCtaCod = a.cCtaCod And nNroCalen = a.nNroCalen"
sSql = sSql & "         and nPrdConceptoCod=1000 and nColocCalendApl=1"
sSql = sSql & "         and nMontoPagado=0),0),"
sSql = sSql & " Atraso=a.nDiasAtraso,"
sSql = sSql & " FechaVenc = B.dVenc"
sSql = sSql & " from colocaccred a"
sSql = sSql & " inner join producto pro on a.cCtaCod = pro.cCtaCod and pro.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205)"
sSql = sSql & " inner join (select max(coloccalendario.cctacod)cctacod, count(coloccalendario.ncuota)ncuota, min(coloccalendario.dvenc)dvenc"
sSql = sSql & "     From coloccalendario"
sSql = sSql & "     inner join colocaccred on coloccalendario.cctacod=colocaccred.cctacod"
sSql = sSql & "     and coloccalendario.nNroCalen=colocaccred.nNroCalen"
sSql = sSql & "     Where coloccalendario.nColocCalendApl = 1"
sSql = sSql & "     and coloccalendario.nColocCalendEstado=0"
sSql = sSql & "     group by coloccalendario.cctacod) b on a.cctacod=b.cctacod"
sSql = sSql & " inner join productopersona c on a.cctacod=c.cctacod and c.nPrdPersRelac=20"
sSql = sSql & " inner join productopersona g on a.cctacod=g.cctacod and g.nPrdPersRelac=28 AND g.cPersCod in " & sCadAnalista
sSql = sSql & " inner join persona d on c.cperscod=d.cperscod"
sSql = sSql & " inner join rrhh h on g.cPersCod=h.cPersCod"
sSql = sSql & " inner join colocaciones e on a.cCtaCod=e.cCtaCod"
sSql = sSql & " where b.dVenc between '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "'"
sSql = sSql & " and substring(a.cctacod,4,2) = " & pdCodAge
sSql = sSql & "" & sCadProd
sSql = sSql & "" & sMonedas
sSql = sSql & " order by Moneda,Analista,a.cctacod"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaRepDetalleCuotasXCobrar = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

' PEAC 20080304

Public Function RecuperaJefeAgencia(ByVal psCodAge As String) As ADODB.Recordset
 
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "exec sql_sel_ObtieneFirmaJefeAgencia '" & psCodAge & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaJefeAgencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function


' PEAC 20070822 -- OPCION 108124
' MADM 20091711 -- 108124 TIPO VIG / CANC
Public Function RecuperaRepCreditosXCampanha( _
ByVal pdCodAge As String, _
ByVal pMatAgencias As Variant, _
ByVal pdFecIni As Date, _
ByVal pdFecFin As Date, _
ByVal psMonedas As String, _
ByVal pMatProd As Variant, _
ByVal pMatAnalistas As Variant, _
ByVal pMatCampanhas As Variant, Optional nTipo As Integer = 1, Optional cMatCondicion As String) As ADODB.Recordset
 
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sCadAnalista As String
Dim sCadCampanha As String
Dim sMonedas As String
Dim sCadTipoCred As String

    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias)
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    'MADM 20100722
    'sCadProd = " And SUBSTRING(C.cCtaCod,6,3) in ('"
    sCadProd = " And c.cTpoCredCod in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & "','"
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sCadAnalista = "('"
    For i = 0 To UBound(pMatAnalistas) - 1
        sCadAnalista = sCadAnalista & pMatAnalistas(i) & "','"
    Next i
    sCadAnalista = Mid(sCadAnalista, 1, Len(sCadAnalista) - 2) & ")"
    
    sCadCampanha = "('"
    For i = 0 To UBound(pMatCampanhas) - 1
        sCadCampanha = sCadCampanha & pMatCampanhas(i) & "','"
    Next i
    sCadCampanha = Mid(sCadCampanha, 1, Len(sCadCampanha) - 2) & ")"
       
    If Len(Trim(psMonedas)) = 0 Then
        sMonedas = ""
    Else
        sMonedas = " AND SUBSTRING(c.cctacod,9,1) IN (" & psMonedas & ") "
    End If
    'MADM 20091711
    sCadTipoCred = ""
    Select Case nTipo
        Case 1 'Por Vigente
            sCadTipoCred = "(2020,2021,2022,2030,2031,2032,2201,2205)"
        Case 2 'Por Cancelados
            sCadTipoCred = "(2050)"
        Case 3 'Todos
            sCadTipoCred = "(2020,2021,2022,2030,2031,2032,2201,2205,2050)"
    End Select
    'MADM 20100713 - AGENCIA - DATRASO
        sSql = " select"
        sSql = sSql & "     Agencia = case convert(int,substring(C.cctacod,4,2)) "
        sSql = sSql & "        when 1 then 'Oficina Principal' "
        sSql = sSql & "        when 2 then 'Agencia Huánuco' "
        sSql = sSql & "        when 3 then 'Agencia Pucallpa' "
        sSql = sSql & "        when 4 then 'Agencia Calle Arequipa' "
        sSql = sSql & "        when 5 then 'Oficina Especial Aeropuerto' "
        sSql = sSql & "        when 6 then 'Agencia Yurimaguas' "
        sSql = sSql & "        when 7 then 'Agencia Tingo María' "
        sSql = sSql & "        when 9 then 'Agencia Belén' "
        sSql = sSql & "        when 10 then 'Agencia Tarapoto' "
        sSql = sSql & "        when 12 then 'Oficina Especial Aguaytía' "
        sSql = sSql & "        when 13 then 'Agencia Requena' "
        sSql = sSql & "        when 24 then 'Agencia Cajamarca' "
        sSql = sSql & "        when 25 then 'Agencia Cerro de Pasco' "
        sSql = sSql & "        when 31 then 'Agencia Punchana' END, "
        sSql = sSql & "     Moneda=case substring(C.cctacod,9,1) when '1' then 'NUEVOS SOLES' ELSE 'DOLARES US' END,"
        sSql = sSql & "     Situacion=case p.nPrdEstado when 2050 then 'Cancelado' else 'Vigente' end,"
        sSql = sSql & "     Campanha=isnull((select cDescripcion"
        sSql = sSql & "             From campanas"
        sSql = sSql & "             where idCampana=cc.idCampana),'SIN CAMPAÑA'),"
        sSql = sSql & "     c.cCtaCod,"
        sSql = sSql & "     Cliente=(select cPersNombre from persona where cpersCod=ppTit.cPersCod),"
        sSql = sSql & "     c.nMontoCol,"
        sSql = sSql & "     CapitalPagado=(c.nMontoCol-p.nSaldo),"
        sSql = sSql & "     p.nSaldo,"
        sSql = sSql & "     c.dVigencia,"
        sSql = sSql & "     Atraso=cc.nDiasAtraso,"
        sSql = sSql & "     Analista=(select cUser from rrhh where cpersCod=ppAna.cPersCod),"
        sSql = sSql & "     Convenio=isnull(( select cNombre + ' - ' + cInstitucion as cNombre"
        sSql = sSql & "                        from CasaComercial where id =cc.id_CasaCom ),'SIN CASA') "
        sSql = sSql & " from colocaciones c"
        sSql = sSql & " INNER join producto p on p.cctacod=c.cctacod and p.nPrdEstado in " & sCadTipoCred  'MADM 20091711
        sSql = sSql & " INNER JOIN colocacCred cc on p.cctacod=cc.cctacod"
        sSql = sSql & " INNER join productopersona ppTit on c.cctacod=ppTit.cctacod and ppTit.nPrdPersRelac=20"
        sSql = sSql & " INNER join productopersona ppAna on c.cctacod=ppAna.cctacod and ppAna.nPrdPersRelac=28 AND ppAna.cPersCod in " & sCadAnalista
        sSql = sSql & " LEFT join Constante CT ON CC.nColocCondicion = CT.nConsValor And CT.nConsCod ='3015' "
        sSql = sSql & " where c.dVigencia between '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "'"
        sSql = sSql & " and substring( p.cctacod,4,2) IN " & sCadAge
        sSql = sSql & " and isnull(cc.idCampana,0) in " & sCadCampanha
        sSql = sSql & " and CC.nColocCondicion IN (" & cMatCondicion & ") "
        sSql = sSql & "" & sCadProd
        sSql = sSql & "" & sMonedas
        'sSql = sSql & " And cc.nDiasAtraso > 0 "
        'sSql = sSql & " ORDER BY Moneda,Situacion,Campanha,c.dVigencia,Agencia"
        sSql = sSql & " ORDER BY Agencia,Moneda,Situacion,Campanha,c.dVigencia"
        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaRepCreditosXCampanha = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

' PEAC 20070921 -- OPCION 108209

Public Function RecuperaRepConvCasillero( _
ByVal pdCodAge As String, _
ByVal pMatAgencias As Variant, _
ByVal pdFecIni As Date, _
ByVal pdFecFin As Date, _
ByVal psMonedas As String, _
ByVal pMatProd As Variant, _
ByVal pMatAnalistas As Variant, _
ByVal pMatCampanhas As Variant, _
ByVal pMatInstitucion As Variant, _
ByVal pMatGastos As Variant) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim sCadAge As String
Dim sCadIns As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sCadAnalista As String
Dim sCadCampanha As String
Dim sCadGastos As String
Dim sMonedas As String

    sCadGastos = "('"
    For i = 0 To UBound(pMatGastos) - 1
        sCadGastos = sCadGastos & pMatGastos(i) & "','"
    Next i
    sCadGastos = Mid(sCadGastos, 1, Len(sCadGastos) - 2) & ")"

    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias)
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadIns = "('"
    For i = 0 To UBound(pMatInstitucion)
        sCadIns = sCadIns & pMatInstitucion(i) & "','"
    Next i
    sCadIns = Mid(sCadIns, 1, Len(sCadIns) - 2) & ")"
'ALPA 20080730****************************************************************************************
'****************************************************************************************************
'sSql = "          select Institucion=lcIns.cpersnombre,lb.cctacod,lw.nmontopagado,lw.ncuota,lcTit.cpersnombre,lf.cuser,lw.dpago,lb.cPersCod,lw.nprdconceptocod,lw.cDescripcion"
'sSql = "          select distinct Institucion=lcIns.cpersnombre,lb.cctacod,AD.nHaber nmontopagado,lw.ncuota,lcTit.cpersnombre,lf.cuser,lw.dpago,lb.cPersCod,lw.nprdconceptocod,lw.cDescripcion"
sSql = "          select Institucion=lcIns.cpersnombre,lb.cctacod,MCD.nMonto nmontopagado,lw.ncuota,lcTit.cpersnombre,lf.cuser,convert(datetime,substring(cMovNro,1,8)) dpago,lb.cPersCod,lw.nprdconceptocod,lw.cDescripcion"
sSql = sSql & "         from colocacconvenio lb"
sSql = sSql & "         inner join (select a.cctacod,a.ncuota,a.nmontopagado,b.dpago,a.nnrocalen,a.nprdconceptocod,d.cDescripcion"
sSql = sSql & "                 from coloccalenddet a"
sSql = sSql & "                 inner join colocaccred c on a.cctacod=c.cctacod"
sSql = sSql & "                 inner join coloccalendario b on a.cctacod=b.cctacod and a.ncuota=b.ncuota"
sSql = sSql & "                     and b.nnrocalen=a.nnrocalen and b.ncoloccalendapl=1"
'sSql = sSql & "                     and b.nnrocalen=a.nnrocalen and b.ncoloccalendestado=1 and b.ncoloccalendapl=1"
sSql = sSql & "                 inner join productoconcepto d on a.nprdconceptocod=d.nprdconceptocod"
sSql = sSql & "                 where a.nprdconceptocod IN " & sCadGastos
sSql = sSql & "                 and a.nmontopagado > 0"
'sSql = sSql & "                 and a.nnrocalen=c.nnrocalen and b.dpago between '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "') lw"
sSql = sSql & "                 and a.nnrocalen=c.nnrocalen and b.dpago is not null) lw"
sSql = sSql & "                 on lb.cctacod=lw.cctacod"
sSql = sSql & "         inner join productopersona ld on lb.cctacod=ld.cctacod and ld.nprdPersRelac=20"
sSql = sSql & "         inner join productopersona le on lb.cctacod=le.cctacod and le.nprdPersRelac=28"
sSql = sSql & "         inner join persona lcIns on lb.cperscod=lcIns.cperscod"
sSql = sSql & "         inner join persona lcTit on ld.cperscod=lcTit.cperscod"
sSql = sSql & "         inner join rrhh lf on le.cperscod=lf.cperscod"
'ALPA 20080730****************************************************************
sSql = sSql & "         inner join movcoldet MCD on (ld.cctacod=MCD.cCtaCod) and (lw.nprdconceptocod=MCD.nprdconceptocod)"
sSql = sSql & "         and (lw.ncuota=MCD.nNroCuota)"
sSql = sSql & "         inner join movcol MC on (MCD.cCtaCod=MC.cCtaCod) and (MCD.cOpeCod=MC.cOpeCod)"
sSql = sSql & "         and (MCD.nMovNro=MC.nMovNro)"
sSql = sSql & "         and (MCD.cCtaCod=MC.cCtaCod)"
sSql = sSql & "         inner join Mov M on (MC.nMovNro=M.nMovNro) /* and  (MC.cOpeCod=M.cOpeCod) */"
'******************************************************************************
sSql = sSql & "         where lb.cPerscod in " & sCadIns
sSql = sSql & "         and substring(lb.cctacod,4,2) IN " & sCadAge
'ALPA 20080730****************************************************************
sSql = sSql & "         and convert(datetime,substring(cMovNro,1,8)) between '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "' and M.nMovFlaG =0"
'sSql = sSql & "         order by lcIns.cpersnombre,lcTit.cpersnombre,lw.ncuota" 'lb.cctacod,
sSql = sSql & "         order by convert(datetime,substring(cMovNro,1,8)),lb.cctacod,lw.ncuota "
'*****************************************************************************
   ' sSql = "exec stp_sel_ReporteCobranzaDiariaConveniosCasillero '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "'," & pMatGastos & "," & sCadAge & "," & sCadCampanha
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaRepConvCasillero = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'peac 20071228
Public Function RecuperaDatosListadoComVigEEFF(ByVal pMatAgencias As Variant) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String

    sCadAge = "''"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "'',''"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 3)

sSql = "stp_sel_ListadoComVigEEFF '" & sCadAge & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

Set RecuperaDatosListadoComVigEEFF = oConecta.CargaRecordSet(sSql)

oConecta.CierraConexion
Set oConecta = Nothing

End Function

'*** PEAC 20080215
Public Function RecuperaDatosDiasTranscDesdeSoli( _
ByVal pdFecIni As Date, _
ByVal pdFecFin As Date, _
ByVal pMatAgencias As Variant, _
ByVal pnDiasIni As Integer, _
ByVal pnDiasFin As Integer) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String

    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

sSql = "exec stp_sel_RepoNumDiasDesdeSolicitud '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & sCadAge & "'," & pnDiasIni & "," & pnDiasFin & ",'2000,2001,2002'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

Set RecuperaDatosDiasTranscDesdeSoli = oConecta.CargaRecordSet(sSql)

oConecta.CierraConexion
Set oConecta = Nothing

End Function

'*** PEAC 20080215
Public Function RecuperaDatosSoliRechazadas( _
ByVal pdFecIni As Date, _
ByVal pdFecFin As Date, _
ByVal pMatAgencias As Variant) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String

    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

sSql = "exec stp_sel_RepoNumDiasDesdeSolicitud '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & sCadAge & "',0,999,'2003'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

Set RecuperaDatosSoliRechazadas = oConecta.CargaRecordSet(sSql)

oConecta.CierraConexion
Set oConecta = Nothing

End Function

'*** PEAC 20080215
Public Function RecuperaDatosSoloEstadoSoli( _
ByVal pdFecIni As Date, _
ByVal pdFecFin As Date, _
ByVal pMatAgencias As Variant) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String

    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

sSql = "exec stp_sel_RepoNumDiasDesdeSolicitud '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & sCadAge & "',0,999,'2000'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

Set RecuperaDatosSoloEstadoSoli = oConecta.CargaRecordSet(sSql)

oConecta.CierraConexion
Set oConecta = Nothing

End Function

'*** PEAC 20080215
Public Function RecuperaDatosSoliProcesadas(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
ByVal pMatAgencias As Variant) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String

    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

sSql = "exec stp_sel_RepoSolicitudesProcesadas '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & sCadAge & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

Set RecuperaDatosSoliProcesadas = oConecta.CargaRecordSet(sSql)

oConecta.CierraConexion
Set oConecta = Nothing

End Function

'*** PEAC 20080221
Public Function RecuperaDatosCliConDistinTiposCred(ByVal pMatAgencias As Variant, _
ByVal pMatProd As Variant) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String, sCadProd As String

'    If UBound(pMatprod) = 0 Then
'
'        sCadProd = ""
'        nEst = 2 '' no selecciona los productos osea muestra todos
'    Else
        sCadProd = ""
        For i = 0 To UBound(pMatProd) - 1
            sCadProd = sCadProd & pMatProd(i) & ","
        Next i
        sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)
'    End If

    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

sSql = "exec stp_sel_RepoCliConDistinTiposCred '" & sCadAge & "','" & sCadProd & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

Set RecuperaDatosCliConDistinTiposCred = oConecta.CargaRecordSet(sSql)

oConecta.CierraConexion
Set oConecta = Nothing

End Function

'***PEAC 20071229
Public Function RecuperaDatosListadoPoliIncendio(ByVal pMatAgencias As Variant, _
                                                 ByVal pMatProd As Variant, ByVal pdFecIni As Date, ByVal pdFecFin As Date) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadProd As String
Dim nEst As Integer

    If UBound(pMatProd) = 0 Then
        sCadProd = ""
        nEst = 2 '' no selecciona los productos osea muestra todos
    Else
        sCadProd = ""
        For i = 0 To UBound(pMatProd) - 1
            sCadProd = sCadProd & pMatProd(i) & ","
        Next i
        sCadProd = sCadProd & "121," & "221"
        sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)
    End If

    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

sSql = "exec stp_sel_ReporteCredConPoliIncendio '" & sCadAge & "','" & sCadProd & "','" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaDatosListadoPoliIncendio = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

'***PEAC 20080303
Public Function RecuperaDatosAECIReporte01(ByVal pdFecIni As String, ByVal pdFecFin As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadProd As String
Dim nEst As Integer

sSql = "exec stp_sel_RepoAECIReporte01 '" & pdFecIni & "','" & pdFecFin & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaDatosAECIReporte01 = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function




'***PEAC 20080219
Public Function RecuperaDatosCredVencPaseCastigo( _
ByVal pnDiaIni As Integer, ByVal pnDiaFin As Integer, ByVal pnUIT As Double, ByVal pnIni As Double, ByVal pnFin As Double, _
ByVal pnTipCam As Double, ByVal pMatAgencias As Variant) As ADODB.Recordset
'MADM 20110515 - UIT x MONTOS

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadProd As String
Dim nEst As Integer

    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

'sSql = "exec stp_sel_RepoCredVigeVencidoParaCastigo " & pnDiaIni & "," & pnDiaFin & "," & pnTipCam & "," & pnUIT & ",'" & sCadAge & "'"
sSql = "exec stp_sel_RepoCredVigeVencidoParaCastigo " & pnDiaIni & "," & pnDiaFin & "," & pnTipCam & "," & pnUIT & "," & pnIni & "," & pnFin & ",'" & sCadAge & "'"


Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaDatosCredVencPaseCastigo = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

' MAVM 19052009
Public Function RecuperaDatosCredVencMay90Dias( _
ByVal pnDiaIni As Integer, ByVal pnDiaFin As Integer, _
ByVal pnTipCam As Double, ByVal pMatAgencias As Variant, ByVal sFI As String, ByVal sFF As String) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim i As Integer
    Dim sCadAge As String
    Dim sCadProd As String
    Dim nEst As Integer
    
        sCadAge = ""
        For i = 0 To UBound(pMatAgencias) - 1
            sCadAge = sCadAge & pMatAgencias(i) & ","
        Next i
        sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
    
    sSql = "exec stp_sel_RepoCredVigeVencidoMayor90Dias " & pnDiaIni & "," & pnDiaFin & "," & pnTipCam & ",'" & sCadAge & "','" & sFI & "','" & sFF & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosCredVencMay90Dias = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

' MAVM 19052009
Public Function RecuperaDatosClientesPreferenciales( _
ByVal pMatAgencias As Variant, ByVal pMatProducts As Variant) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim i As Integer
    Dim sCadAge As String
    Dim sCadProd As String
    
        sCadAge = ""
        For i = 0 To UBound(pMatAgencias) - 1
            sCadAge = sCadAge & pMatAgencias(i) & ","
        Next i
        sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
        
        sCadProd = ""
        For i = 0 To UBound(pMatProducts) - 1
            If i = 0 Then
                sCadProd = "" & pMatProducts(i) & ""
            Else
                sCadProd = sCadProd & "," & pMatProducts(i) & ""
            End If
        Next i
        
        If UBound(pMatProducts) = 0 Then
            sCadProd = ""
        End If
    
    sSql = "exec stp_sel_RepoClientesPreferenciales '" & sCadAge & "', '" & sCadProd & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosClientesPreferenciales = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function


'***PEAC 20080115
Public Function RecuperaCredPigVigVen(ByVal pcCodCli As String, ByVal pdFecSis As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_RecuperaCredPigVigVen '" & pcCodCli & "','" & Format(pdFecSis, "yyyymmdd") & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaCredPigVigVen = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function
'**ALPA***18/04/2008********************************************************
Public Function CanceladosxAmpliacion(ByVal pdFecIni As Date, ByVal pdFecFin As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_Cancelados_x_ampliacion '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set CanceladosxAmpliacion = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function
Public Function PreCancelaciones(ByVal pdFecIni As Date, ByVal pdFecFin As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_reporte_pre_cancelaciones '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set PreCancelaciones = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

Public Function ObtenerFechaProNumFuente(ByVal sNumFte As String, ByVal dFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_ObtenerFechaProNumFuente '" & sNumFte & "','" & Format(dFecha, "YYYY/mm/dd") & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtenerFechaProNumFuente = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'**End***********************************************************************

'*** PEAC 20080412
Public Function ObtenerCodHojEval() As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "exec stp_sel_ObtieneCodHojEval"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtenerCodHojEval = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function



'*** PEAC 20080412
Public Function ObtenerPatrimonio(ByVal psCodCli As String, ByVal psCodBusca As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "exec stp_sel_ObtienePatrimonio '" & psCodCli & "','" & psCodBusca & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtenerPatrimonio = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

'*** PEAC 20080412
Public Function ObtenerDatosDDJJPatri(ByVal psCodCta As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "exec stp_sel_RecuperaDatosDDJJPatrimonial '" & psCodCta & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtenerDatosDDJJPatri = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

'*** PEAC 20080412
Public Function RecuperaRelaBancosPersona(ByVal psCodNat As String, ByVal psCodJur As String, ByVal psCuenta As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
If Len(psCodNat) > 0 Then
    sSql = "exec stp_sel_RecuperaRelaBancosPersona '" & psCodNat & "','1','" & psCuenta & "'"
Else
    sSql = "exec stp_sel_RecuperaRelaBancosPersona '" & psCodJur & "','2'"
End If
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaRelaBancosPersona = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

'*** PEAC 20080510
Public Function RecuperaDatosListadoPoliza(ByVal pdFecSis As Date, pMatAgencias As Variant) As ADODB.Recordset
    
Dim sSql As String
'Dim oConecta As DConecta
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String

'    sCadAge = "('"
'    For i = 0 To UBound(pMatAgencias) - 1
'        sCadAge = sCadAge & pMatAgencias(i) & "','"
'    Next i
'    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

'''---------------------------------------------------
    
    '*** PEAC 20090330
    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

    '''------------------------------------------------------------------
    
    '*** PEAC 20080510
    '"cDireccionInmueble=ISNULL(cDirGarRegPubli,''), " & _*
    
'    sSql = " SELECT  Prd.cCtaCod," & _
'           " cCodPolizaAseg= ISNULL(cCodPolizaAseg,''), " & _
'            " cCiaSeguros= ASEG.cPersNombre, con.cConsDescripcion TipoPoliza," & _
'            "dVigenciaAseg= ISNULL(dVigenciaAseg,0), " & _
'            "dVencAseg= ISNULL(dVencAseg,0), " & _
'            "cCondicionPoliza=CASE WHEN ISNULL(dVencAseg,0)<= '" & Format(pdFecSis, "mm/dd/yyyy") & "' THEN 'VENCIDO'ELSE 'VIGENTE'END, " & _
'            "cDireccionInmueble=ISNULL(g.cDireccion,''), " & _
'            "nSumaAsegurada = IsNull(nSumaAsegurada, 0) " & _
'            "Into #ColocPoliza " & _
'            "FROM    Producto Prd INNER JOIN ColocGarantia CG ON Prd.cCtaCod = CG.cCtaCod " & _
'            "INNER JOIN Garantias G ON CG.cNumGarant = G.cNumGarant " & _
'            "INNER JOIN GarantReal GR ON GR.cNumGarant = G.cNumGarant " & _
'            "INNER JOIN GarantRealTasacion GRT ON GRT.cNumGarant = GR.cNumGarant " & _
'            "AND GRT.dTasacion = (SELECT MAX(dTasacion)FROM GarantRealTasacion WHERE cNumGarant=GRT.cNumGarant) " & _
'            "INNER JOIN GarantPoliza GP ON GP.cNumGarant = GRT.cNumGarant AND GP.dTasacion = GRT.dTasacion " & _
'            "INNER JOIN Poliza Pol ON GP.cNumPoliza = Pol.cNumPoliza " & _
'            "inner join constante con on pol.ntipopoliza=con.nconsvalor and con.nconscod=9066 " & _
'            "INNER JOIN Persona ASEG ON ASEG.cPersCod = Pol.cPersCodAseg " & _
'            "Where Prd.nPrdEstado >= 2020 And Prd.nPrdEstado <= 2032 " & _
'            "AND SUBSTRING(Prd.cCtaCod,4,2) IN " & sCadAge & " AND GP.nEstado= 1"
'
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    oConecta.Ejecutar (sSql)
'
'    sSql = " SELECT CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL' "
'    sSql = sSql & "ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')   END AS sAgencia, "
'sSql = sSql & "CASE WHEN GROUPING(P.cPersNombre)=1 THEN 'TOTAL' "
'sSql = sSql & "ELSE ISNULL(P.cPersNombre, 'CLIENTE DESCONOCIDO')   END AS sCliente, "
'sSql = sSql & "CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL' "
'sSql = sSql & "ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDA')   END AS cCtaCod, "
'sSql = sSql & "SUM(C.nMontoCol) as nMontoDesemb, SUM(Prd.nSaldo) as nSaldoCap, "
'sSql = sSql & "cMoneda=CASE WHEN SUBSTRING(Prd.cCtaCod,9,1)='1'THEN 'MN'ELSE 'ME'END, " & _
'            "dDesembolso=(SELECT dVigencia FROM Colocaciones WHERE cCtaCod =Prd.cCtaCod ), " & _
'            "dTermino = (SELECT MAX(dVenc) FROM ColocCalendario WHERE cCtaCod= Prd.cCtaCod AND nColocCalendApl=1 " & _
'            "    AND nNroCalen = (SELECT nNroCalen FROM ColocacCred WHERE cCtaCod=Prd.cCtaCod)), " & _
'            "TipoPoliza=( SELECT TOP 1 TipoPoliza FROM #ColocPoliza WHERE cCtaCod = Prd.cCtaCod)," & _
'            "cCodPolizaAseg= ( SELECT TOP 1 cCodPolizaAseg FROM #ColocPoliza WHERE cCtaCod = Prd.cCtaCod), " & _
'            "cCiaSeguros= ( SELECT TOP 1 cCiaSeguros FROM #ColocPoliza WHERE cCtaCod = Prd.cCtaCod), " & _
'            "dVigenciaAseg= ( SELECT TOP 1 dVigenciaAseg FROM #ColocPoliza WHERE cCtaCod = Prd.cCtaCod), " & _
'            "dVencAseg= ( SELECT  TOP 1 dVencAseg FROM #ColocPoliza WHERE cCtaCod = Prd.cCtaCod), " & _
'            "cCondicionPoliza= ( SELECT TOP 1 cCondicionPoliza FROM #ColocPoliza WHERE cCtaCod = Prd.cCtaCod), " & _
'            "cDireccionInmueble= ( SELECT TOP 1 cDireccionInmueble FROM #ColocPoliza WHERE cCtaCod = Prd.cCtaCod), " & _
'            "nSumaAsegurada = ( SELECT TOP 1 nSumaAsegurada FROM #ColocPoliza WHERE cCtaCod = Prd.cCtaCod) " & _
'            "From Producto Prd Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod " & _
'            "Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2) " & _
'            "INNER JOIN ProductoPersona PP ON Prd.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = 20 " & _
'            "INNER JOIN Persona P ON P.cPersCod = PP.cPersCod INNER JOIN #ColocPoliza CP ON CP.cCtaCod = Prd.cCtaCod " & _
'            "Where Prd.nPrdEstado >= 2020 And Prd.nPrdEstado <= 2032 " & _
'            " AND A.cAgeCod IN " & sCadAge & " Group By A.cAgeDescripcion, P.cPersNombre, Prd.cCtaCod" & _
'            " WITH ROLLUP ORDER BY sAgencia, sCliente, cCtaCod"
'
'    Set RecuperaDatosListadoPoliza = oConecta.CargaRecordSet(sSql)
'
'    sSql = "DROP TABLE #ColocPoliza"
'
'    oConecta.Ejecutar (sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing

'''--------------------------------------------------------------------
    
    sSql = "exec stp_sel_RecuperaDatosListadoPoliza '" & Format(pdFecSis, "yyyymmdd") & "','" & sCadAge & "'"

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosListadoPoliza = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'*** PEAC 20080620
Public Function ObtenerGarantiasAvales(ByVal psCodCta As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "exec stp_sel_ObtenerGarantAvales '" & psCodCta & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtenerGarantiasAvales = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

'*** PEAC 20080620
Public Function ObtenerGarantiasTitular(ByVal psCodCta As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "exec stp_sel_ObtenerGarantTitular '" & psCodCta & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtenerGarantiasTitular = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

'*** PEAC 20080630
Public Function RecuperaResumenGarTitularAval(ByVal psCuenta As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_ObtenerResumenGarTitularAval '" & psCuenta & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaResumenGarTitularAval = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'ALPA 20080820*******
Public Function ObtenerCalificacionDiasAtraso_Titular(ByVal psPerCod As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "select dbo.ObtenerCalificacionDiasAtraso_Titular('" & psPerCod & "') as cCalificacion "

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtenerCalificacionDiasAtraso_Titular = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'*******************

'***PEAC 20080923
Public Function RecuperaClientesPotencialesSinCredVig(ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pMatAgencias As Variant) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadProd As String
Dim nEst As Integer

    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

sSql = "exec stp_sel_RelacionClientesPotencialesSinCredVig '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & sCadAge & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaClientesPotencialesSinCredVig = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

'*** PEAC 20080923
Public Function ObtieneUltimaFechaRCC()
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "exec stp_sel_ObtieneUltimaFechaRCC "

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneUltimaFechaRCC = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

'***PEAC 20080924
Public Function RecuperaNumYSaldoCredPorProductoConsol(ByVal pnTipoC As Double) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "exec stp_sel_ReporteSaldosCredPorProdConsol " & pnTipoC

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaNumYSaldoCredPorProductoConsol = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

'************************************************************'
'** GITU 20080926 108337 Segun Memo Nº 1705-2008-GM-DI/CMAC *'
Public Function RecuperaCreditosRiesgosCrediticios(ByVal psTipoCambio As String, ByVal pdFechaSistema As Date) As ADODB.Recordset

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim MesAnterior As String
    
    '*** BRGO 16/11/2010
    '*** Modificado el 16/12/2010 según requerimiento RQ10149 - BRGO
    MesAnterior = CStr(Year(DateAdd("m", -1, pdFechaSistema))) & Format(CStr(Month(DateAdd("m", -1, pdFechaSistema))), "00")
        
    sSql = " Select cCodAge, vAgencia, "
    sSql = sSql & "    Sum (Case When nCondCre = 1  And LEFT(CONVERT(VARCHAR(8),dFecVig,112),6) = " & MesAnterior & " Then 1 else 0 end) nNumNuevos, "
    sSql = sSql & "    Sum (Case When nCondCre <> 1 And LEFT(CONVERT(VARCHAR(8),dFecVig,112),6) = " & MesAnterior & " Then 1 else 0 end) nNumRepres, "
    sSql = sSql & "    Sum (Case When nCondCre = 1  And LEFT(CONVERT(VARCHAR(8),dFecVig,112),6) = " & MesAnterior & " Then nSaldoCap else 0 end) nSalNuevos, "
    sSql = sSql & "    Sum (Case When nCondCre <> 1 And LEFT(CONVERT(VARCHAR(8),dFecVig,112),6) = " & MesAnterior & " Then nSaldoCap else 0 end) nSalRepres, "
    sSql = sSql & "    Sum (Case When SubString(nPrdEstado,3,1) = 2 And nDiasAtraso < 30 Then 1 else 0 end) nNumVig, "
    sSql = sSql & "    Sum (Case When SubString(nPrdEstado,3,1) = 3 And nDiasAtraso < 30 Then 1 else 0 end) nNumRef, "
    sSql = sSql & "    Sum (Case When nDiasAtraso >= 30 Then 1 else 0 end) nNumAtr, "
    sSql = sSql & "    Sum (Case When SubString(nPrdEstado,3,1) = 2 And nDiasAtraso < 30 Then nSaldoCap else 0 end) nSalVig, "
    sSql = sSql & "    Sum (Case When SubString(nPrdEstado,3,1) = 3 And nDiasAtraso < 30 Then nSaldoCap else 0 end) nSalRef, "
    sSql = sSql & "    Sum (Case When nDiasAtraso >= 30 Then nSaldoCap else 0 end) nSalAtr, "
    sSql = sSql & "    Sum (Case When LEFT(CONVERT(VARCHAR(8),dFecVig,112),6) = " & MesAnterior & " Then 1 Else 0 End) nNumTotalMes, "
    sSql = sSql & "    Sum (Case When LEFT(CONVERT(VARCHAR(8),dFecVig,112),6) = " & MesAnterior & " Then nSaldoCap Else 0 End) nSalTotalMes, "
    sSql = sSql & "    Sum (1) nNumTotal, "
    sSql = sSql & "    Sum (nSaldoCap) nSalTotal, "
    sSql = sSql & "    Count(distinct (Case When LEFT(CONVERT(VARCHAR(8),dFecVig,112),6) = " & MesAnterior & " Then cCodAnalista End)) Analistas "
    sSql = sSql & " From( "
    sSql = sSql & "    Select AB.cCodAge, AB.vAgencia, CC.cCtaCod, CC.dFecVig,"
    sSql = sSql & "        case substring(CC.cCtaCod,9,1) when '1' then CC.nSaldoCap else (CC.nSaldoCap* " & psTipoCambio & ") end as nSaldoCap, "
    sSql = sSql & "        CC.nDiasAtraso, CC.nCondCre, Cast(CC.nPrdEstado As Varchar) As nPrdEstado, cCodAnalista "
    sSql = sSql & "    From DBConsolidada..CreditoConsol CC "
    sSql = sSql & "        Join ColocConvBcoNac CB On CC.cCtaCod = CB.cCtaCod "
    sSql = sSql & "        Join AgenciasBancoNacion AB On CB.cCodAgeBcoNac = AB.cCodAge "
    sSql = sSql & "    Where CC.nPrdEstado In (2020,2021,2022,2030,2031,2032) "
    sSql = sSql & ") aa "
    sSql = sSql & " Group By  cCodAge, vAgencia "
    sSql = sSql & " Order by vAgencia "
    '*** End BRGO
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosRiesgosCrediticios = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'***PEAC 20080929
Public Function ObtieneResCarteraPorProdSinCF(ByVal pnTipoC As Double) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "exec stp_sel_ReporteSaldosCredPorProdResSinCarFza " & pnTipoC

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneResCarteraPorProdSinCF = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

'*** PEAC 20081003 *****************************************************
Public Function RecuperaDatosConsolRentabilidadCarteraXAnalista(ByVal dFechaIni As Date, ByVal dFechaFin As Date, ByVal nCanDiaMora As Integer, ByVal sAnslis As Variant, ByVal sAgencias As Variant, Optional sProductos As Variant = "") As ADODB.Recordset
'ALPA 20090324******************************
'Se agrego el atributo sProductos a la cabecera de la función
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler
    
    'ALPA 20090324**************************************************************************************************************************************************
    'sSql = "exec stp_sel_ReporteConsolRentabilidadCarteraXAnalista '" & Format(dFechaIni, "yyyymmdd") & "','" & Format(dFechaFin, "yyyymmdd") & "'," & nCanDiaMora & ",'" & sAnslis & "','" & sAgencias & "'"
    sSql = "exec stp_sel_ReporteConsolRentabilidadCarteraXAnalista '" & Format(dFechaIni, "yyyymmdd") & "','" & Format(dFechaFin, "yyyymmdd") & "'," & nCanDiaMora & ",'" & sAnslis & "','" & sAgencias & "', '" & sProductos & "'"
    '***************************************************************************************************************************************************************
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaDatosConsolRentabilidadCarteraXAnalista = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set RecuperaDatosConsolRentabilidadCarteraXAnalista = Null
End Function
'*** BRGO 20111125 *****************************************
Public Function ObtieneDatosInfoGasParaCartas(ByVal sCuenta As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_ObtieneDatosInfoGasParaCartas '" & sCuenta & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set ObtieneDatosInfoGasParaCartas = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'*** END BRGO **********************************************

'***Agregado por ELRO el 20120104, según Acta N° 002-2012/TI-D
Public Function recuperarCreditosCanceladoCastigado(ByVal pcCodCli As String, ByVal pdFecSis As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_CreditosCanceladoCastigado '" & pcCodCli & "','" & Format(pdFecSis, "mm/dd/yyyy") & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set recuperarCreditosCanceladoCastigado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function
'***Fin Agregado por ELRO*************************************
'***Agregado por ELRO el 20120719, según OYP-RFC076-2012
Public Function devolverPignoraticiosAdjudicadosCliente(ByVal pcCodCli As String, ByVal pdFecSis As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "exec stp_sel_DevolverPignoraticiosAdjudicadosCliente '" & pcCodCli & "','" & Format(pdFecSis, "mm/dd/yyyy") & "'"

Set oConecta = New COMConecta.DCOMConecta

oConecta.AbreConexion
Set devolverPignoraticiosAdjudicadosCliente = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion

Set oConecta = Nothing

End Function
'***Fin Agregado por ELRO el 20120719*************************

'** Juez 20120803 ********************************************
Public Function ObtieneDatosReimprDesemb(ByVal pnMovNro As Long) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_ObtieneDatosReimprDesemb " & pnMovNro & ""

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimprDesemb = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

Public Function ObtieneDatosReimprPagoCred(ByVal pnMovNro As Long) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_ObtieneDatosReimprPagoCred " & pnMovNro & ""

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimprPagoCred = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

Public Function ObtieneDatosReimprPagoJud(ByVal pnMovNro As Long) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_ObtieneDatosReimprPagoJud " & pnMovNro & ""

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimprPagoJud = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

Public Function ObtieneDatosReimprCFComision(ByVal pnMovNro As Long) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_ObtieneDatosReimprCFComision " & pnMovNro & ""

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimprCFComision = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

Public Function ObtieneDatosReimprPigno(ByVal pnMovNro As Long) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_ObtieneDatosReimprPigno " & pnMovNro & ""

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimprPigno = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function
'** End Juez *************************************************
'JUEZ 20130906 ***************************************************************
Public Function ObtieneDatosReimprComisiones(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_ObtieneDatosReimprComisiones " & pnMovNro & ""

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimprComisiones = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'END JUEZ ********************************************************************

'MAVM 20140320 ***
Public Function RecuperaGastoConv(ByVal psCtaCod As String) As ADODB.Recordset
On Error GoTo ErrorRecuperaGastoConv
    Dim oConec As New COMConecta.DCOMConecta
    Dim sSql As String
    sSql = "exec stp_sel_RecuperaGastoConv '" & psCtaCod & "'"
    oConec.AbreConexion
    Set RecuperaGastoConv = oConec.CargaRecordSet(sSql, adLockReadOnly)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrorRecuperaGastoConv:
    Err.Raise Err.Number, "DCOMCredDoc:RecuperaGastoConv"", Err.Description"
End Function
'***

'RECO20140305 ERS174-2013*************************************
Public Function ObtieneRatiosFI(ByVal psNumFuente As String, ByVal psCtaCod As String) As ADODB.Recordset
On Error GoTo ErrorObtieneRatiosFI
    Dim oConec As New COMConecta.DCOMConecta
    Dim sSql As String
    sSql = "exec stp_ObtieneRatiosFI '" & psNumFuente & "','" & psCtaCod & "'"
    oConec.AbreConexion
    Set ObtieneRatiosFI = oConec.CargaRecordSet(sSql, adLockReadOnly)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrorObtieneRatiosFI:
    Err.Raise Err.Number, "DCOMCredDoc:RecuperaRatios"", Err.Description"
End Function
Public Function RecuperaRelaBancosPersonaHojaApr(ByVal psCodNat As String, ByVal psCodJur As String, ByVal psCuenta As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    If Len(psCodNat) > 0 Then
        sSql = "exec stp_sel_RecuperaRelaBancosPersonaHojaApr '" & psCodNat & "','1','" & psCuenta & "'"
    Else
        sSql = "exec stp_sel_RecuperaRelaBancosPersonaHojaApr '" & psCodJur & "','2'"
    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaRelaBancosPersonaHojaApr = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'FIN RECO*****************************************************
'PASI20150926 ERS0692015
Public Function ObtieneDatosReimpPagoLeasing(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_ObtieneDatosReimprPagoArrendFinan " & pnMovNro

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpPagoLeasing = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpPagoxHonramiento(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_ObtieneDatosReimprPagoCredxHonramiento " & pnMovNro

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpPagoxHonramiento = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpDesembPigno(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_ObtieneDatosReimprDesembPigno " & pnMovNro

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpDesembPigno = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpDupliContPigno(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_ObtieneDatosReimprDupliContPigno " & pnMovNro

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpDupliContPigno = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpCustoDifPigno(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_ObtieneDatosReimprCustoDifPigno " & pnMovNro

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpCustoDifPigno = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpRecupContPigno(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "stp_sel_ObtieneDatosReimprRecupContPigno " & pnMovNro

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpRecupContPigno = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpApeLoteEfectivoCab(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprApeLoteEfectivoCab " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpApeLoteEfectivoCab = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpApeLoteEfectivoDet(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprApeLoteEfectivoDet " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpApeLoteEfectivoDet = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpDepositoCaptacII(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprDepositoCaptacII " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpDepositoCaptacII = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpRetiroCaptacOrdPagCanje(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprRetiroCaptacOrdPagCanje " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpRetiroCaptacOrdPagCanje = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtienDatosReimpRetiroCaptacOrdPagCanjeComision(ByVal pnMovNro As Long, ByVal psUser As String, ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprRetiroCaptacOrdPagCanjeComision " & pnMovNro & ",'" & psUser & "','" & psCtaCod & "'"
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtienDatosReimpRetiroCaptacOrdPagCanjeComision = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpRetiroCaptacComisTransf(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprRetiroCaptacComisTransf " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpRetiroCaptacComisTransf = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpCancelActivCaptac(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprCancelaActivCaptac " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpCancelActivCaptac = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpRetiroPFIntEfec(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprRetiroPFIntEfect " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpRetiroPFIntEfec = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpRetiroPFIntAboCtaBoletaPF(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprRetiroPFIntAboCtaBoletaPF " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpRetiroPFIntAboCtaBoletaPF = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpRetiroPFIntAboCtaBoletaAho(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprRetiroPFIntAboCtaBoletaAHO " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpRetiroPFIntAboCtaBoletaAho = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpRetiroPFIntCash(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprRetiroPFIntCash " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpRetiroPFIntCash = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpCancelaPFEfectivo(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprCancelaPFEfectivo " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpCancelaPFEfectivo = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpPFAumentCapEfec(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprPFAumentCapEfect " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpPFAumentCapEfec = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpApeCTSLoteEfectivoCab(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprApeCTSLoteEfectivoCab " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpApeCTSLoteEfectivoCab = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpApeCTSLoteEfectivoDet(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprApeCTSLoteEfectivoDet " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpApeCTSLoteEfectivoDet = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpCancelaCTSEfectivo(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprCancelaCTSEfectivo " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpCancelaCTSEfectivo = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpOtrasOperaciones(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimpOtrasOperaciones " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpOtrasOperaciones = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpOtraOpeDepCtaBanco(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimpOtraOpeDepCtaBanco " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpOtraOpeDepCtaBanco = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpDesemOtroGasto(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimpDesembOtroGasto " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpDesemOtroGasto = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpDesemOtroGastoCCHICA(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimpDesembOtroGastoCCHICA " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpDesemOtroGastoCCHICA = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpDevSobOtrasOpeCheque(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimpDevSobOtrasOpeCheque " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpDevSobOtrasOpeCheque = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpComisionOtros(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimpComisionOtros " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpComisionOtros = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpDebitoSP(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimpDebitoSP " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpDebitoSP = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpNotaAbono(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprNotaAbono " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpNotaAbono = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpGiro(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprGiro " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpGiro = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpGiroCambDest(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprGiroCambDest " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpGiroCambDest = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneDatosReimpGiroAnulacion(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "stp_sel_ObtieneDatosReimprGiroAnulacion " & pnMovNro
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosReimpGiroAnulacion = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'end PASI
'***JGPA20190607-----------------------
Public Function ObtieneDatosOperacionREU(ByVal pnMovNro As Long, ByVal pnNroReu As Long, ByVal psOpeCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "Exec stp_sel_ObtieneDatosOperacionREU " & pnMovNro & "," & pnNroReu & ",'" & psOpeCod & "'"
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneDatosOperacionREU = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function ObtieneIntervinientesOperacionREU(ByVal pnMovNro As Long, ByVal pnNroReu As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
sSql = "Exec stp_sel_ObtieneIntervinientesOperacionREU " & pnMovNro & "," & pnNroReu & ""
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set ObtieneIntervinientesOperacionREU = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'***End JGPA20190607-------------------
'by angc20200518
Public Function ObtieneDatosActivatePeru(ByVal pscCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "exec sp_sel_ObtenerDatosreactiva '" & pscCtaCod & "'"
    Set ObtieneDatosActivatePeru = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
 End Function
'**CTI2 FERIMORO : 17/02/2021 : MEJORA EL PROCESO PARA REIMPRIMIR COBRANZA SOAT
Public Function ObtieneDatosSoat(ByVal nMovNro As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "select cMovNro from mov where nMovNro='" & nMovNro & "'"
    Set ObtieneDatosSoat = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
 End Function
'***********************************************************************
'CTI5 20210519***
Public Function ObtenerMatrizGastoCierrexReimprDesemb(ByVal pnMovNro As Long) As Variant
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rsGastoFMV As ADODB.Recordset
    Dim MatApeGastoCierre As Variant
    
    sSql = "stp_sel_ObtieneDatosReimprDesemGastoFMV " & CStr(pnMovNro)
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rsGastoFMV = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
            
    If Not rsGastoFMV.EOF Then
        ReDim MatApeGastoCierre(14)
    
        MatApeGastoCierre(0) = "" 'Cod Titular
        MatApeGastoCierre(1) = "" 'Nombre Titular
        MatApeGastoCierre(2) = 0 'Personería
        MatApeGastoCierre(3) = 0 'Programa Ahorro Corriente
        MatApeGastoCierre(4) = 0 'Tipo de Cuenta
        MatApeGastoCierre(5) = 0 'Tipo de Tasa
        MatApeGastoCierre(6) = 0 'Tasa
        MatApeGastoCierre(7) = CDbl(rsGastoFMV!nMontoCap) 'Monto de apertura
        MatApeGastoCierre(8) = False 'Documento
        MatApeGastoCierre(9) = "" 'Nro de documento
        MatApeGastoCierre(10) = "" 'Codigo IFI
        MatApeGastoCierre(11) = 0 'Plazo Abono
        MatApeGastoCierre(12) = rsGastoFMV!sCtaAho 'Nro Cuenta a generar
        MatApeGastoCierre(13) = CDbl(rsGastoFMV!nMontoITF) 'ITF Abono
        MatApeGastoCierre(14) = "" 'Direccion Titular
    End If
    
    ObtenerMatrizGastoCierrexReimprDesemb = MatApeGastoCierre
End Function
'END CTI5 *******
