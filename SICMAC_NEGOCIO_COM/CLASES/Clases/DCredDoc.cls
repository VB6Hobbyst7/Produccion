VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCredDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function RecuperaDatosMantCargoAutomatico(ByVal psCtaCod As String) As ADODB.Recordset
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim RCD As nRcdReportes
Set RCD = New nRcdReportes

SQL = " Select Pro.cCtaCod ,P1.cPersCod, P1.cPersNombre Nombre, Pro.nSaldo Saldo from Producto Pro"
SQL = SQL & " Inner Join ProductoPersona PP1 on PP1.cCtaCod = Pro.cCtaCod"
SQL = SQL & " Inner Join Persona P1 on P1.cPersCod = PP1.cPersCod"
SQL = SQL & " Where PP1.nPrdPersRelac = 20"
SQL = SQL & " and Pro.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
SQL = SQL & " and Pro.cCtaCod='" & psCtaCod & "'"
Set rs = RCD.GetQuery(SQL)
Set RecuperaDatosMantCargoAutomatico = rs
Set rs = Nothing
Set RCD = Nothing
End Function
Public Function RecuperaDatosColocCargoAutoma(ByVal psCtaCod As String) As ADODB.Recordset
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim RCD As nRcdReportes
Set RCD = New nRcdReportes

SQL = " Select cCtaCod,nOrden,cCodCtaAho,cEstado  from ColocCargoAutoma"
SQL = SQL & " where cCtaCod='" & psCtaCod & "' and substring(cCtaCod,9,1)='" & Mid(psCtaCod, 9, 1) & "'"
SQL = SQL & " Order by cEstado, nOrden"

Set rs = RCD.GetQuery(SQL)
Set RecuperaDatosColocCargoAutoma = rs
Set rs = Nothing
Set RCD = Nothing
End Function
Public Sub ActualizaColocCargoAutoma(ByVal psCtaCod As String, ByVal psCtaCodAho As String, _
ByVal pnOrden As Integer, ByVal pnEstado As Integer)
Dim SQL As String
Dim Co As DConecta
Set Co = New DConecta
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
Dim Co As DConecta
Set Co = New DConecta
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
Dim rs As New ADODB.Recordset
Dim RCD As nRcdReportes
Set RCD = New nRcdReportes

SQL = " Select * from ColocCargoAutoma"
SQL = SQL & " where cEstado = '1' "
SQL = SQL & " and cCtaCod ='" & Trim(psCtaCod) & "' "
SQL = SQL & " and cCodCtaAho ='" & Trim(psCtaCodAho) & "' "
Set rs = RCD.GetQuery(SQL)
If rs.EOF And rs.BOF Then
    ExisteCtaCargoAutoma = False
Else
    ExisteCtaCargoAutoma = True
End If
Set rs = Nothing
Set RCD = Nothing
End Function


Public Sub InsertaColocCargoAutoma(ByVal psCtaCod As String, _
ByVal psCtaCodAho As String, ByVal psEstado As String, ByVal pnOrden As Integer)
Dim Co As DConecta
Dim SQL As String
Set Co = New DConecta
SQL = " Insert ColocCargoAutoma "
SQL = SQL & " (cCtaCod,nOrden,cCodCtaAho,cEstado) Values "
SQL = SQL & " ('" & psCtaCod & "'," & pnOrden & ",'" & psCtaCodAho & "','" & psEstado & "')"
Co.AbreConexion
Co.Ejecutar (SQL)
Co.CierraConexion
Set Co = Nothing
End Sub




Public Function RecuperaDatosSolicitudCredito(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

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

sSql = "Select Pers.cPersCod, Pers.cPersNombre, PID.cPersIDNro DNI,PID2.cPersIDNro RUC,"
sSql = sSql & " Pers.cPersDireccDomicilio, PNat.nPersnatHijos, Pers.dPersNacCreac,"
sSql = sSql & " PNat.nPersNatEstCiv,  CN.cConsDescripcion cEstadoCivil,"
sSql = sSql & " PFI.cPersNombre cRazonSocial, PFI.cPersDireccDomicilio cDireccRazon,"
sSql = sSql & " PFI.cPersCIIU, CI.cCIIUdescripcion, PFI.dPersNacCreac, PersAna.cPersNombre cAnalista,"
sSql = sSql & " CN2.cConsDescripcion cDestino,  CN3.cConsDescripcion cCondicion, CE.dPrdEstado dFecAsig,"
sSql = sSql & " Inst.cPersNombre cNomInst, CV.cCodModular, CE.nMonto nMontoSol,"
sSql = sSql & " CN4.cConsDescripcion cMonedaSol, CE.nCuotas nCuotasSol, CE.nPlazo nPlazoSol,"
sSql = sSql & " CN5.cConsDescripcion cTipoCred, CN6.cConsDescripcion cPersoneria, PF.cPersCod cFteInd ,"
sSql = sSql & " PF.nPersTipFte as nTipoFuente"
sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtacod AND PP.nPrdPersRelac = 20"
sSql = sSql & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
sSql = sSql & " Left Join PersID PID ON PID.cPersCod = PP.cPersCod AND PID.cPersIDTpo = 1"
sSql = sSql & " Left Join PersID PID2 ON PID2.cPersCod = PP.cPersCod AND PID2.cPersIDTpo = 2"
sSql = sSql & " Left Join PersonaNat PNat ON PNat.cPersCod = PP.cPersCod"
sSql = sSql & " Left Join Constante CN ON CN.nConsValor = PNat.nPersNatEstCiv AND CN.nConsCod = 1020"
sSql = sSql & " Left Join ColocFteIngreso CF ON CF.cCtaCod = P.cCtaCod"
sSql = sSql & " Left Join PersFteIngreso PF ON PF.cNumFuente = CF.cNumFuente"
sSql = sSql & " Left Join Persona PFI ON PFI.cPersCod = PF.cPersFIPersCod"
sSql = sSql & " Left Join CIIU CI ON PFI.cPersCIIU = CI.cCIIUcod"
sSql = sSql & " Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = 28"
sSql = sSql & " Inner Join Persona PersAna ON PersAna.cPersCod = PP2.cPersCod"
sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod"
sSql = sSql & " Left Join Constante CN2 ON CN2.nConsValor = CC.nColocDestino AND CN2.nConsCod = 3016"
sSql = sSql & " Left Join Constante CN3 ON CN3.nConsValor = CC.nColocCondicion AND CN3.nConsCod = 3015"
sSql = sSql & " Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado = 2000"
sSql = sSql & " Left Join ColocacConvenio CV ON CV.cCtaCod = P.cCtaCod"
sSql = sSql & " Left Join Persona Inst ON Inst.cPersCod = CV.cPersCod"
sSql = sSql & " Left Join Constante CN4 ON Convert(int, SubString(P.cCtaCod,9,1)) = CN4.nConsValor AND CN4.nConsCod = 1011"
sSql = sSql & " Left Join Constante CN5 ON Convert(int, SubString(P.cCtaCod,6,3)) = CN5.nConsValor AND CN5.nConsCod = 1001"
sSql = sSql & " Left Join Constante CN6 ON CN6.nConsValor = Pers.nPersPersoneria AND CN6.nConsCod = 1002"
sSql = sSql & " Left Join PersFIDependiente PFDep ON PFDep.cNumFuente = CF.cNumFuente AND PFDep.dPersEval = CF.dPersEval"
sSql = sSql & " Left Join PersFIIndependiente PFInd ON PFInd.cNumFuente = CF.cNumFuente AND PFInd.dPersEval = CF.dPersEval"
sSql = sSql & " Where P.cCtaCod = '" & psCtaCod & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosSolicitudCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaDatosResumenComite(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
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
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosResumenComite = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

' CMACICA_CSTS - 08/11/2003 VERIFICAR SI ESTA FUNCION DEBERIA IR EN ESTA CLASE O EN ALGUNA CLASE DE PERSONAS -----------

Public Function ObtieneNombrePersonaUser(ByVal psCodUser As String) As String
Dim SQL As String
Dim rs As ADODB.Recordset
Dim Co As DConecta
    Set Co = New DConecta
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

Public Function RecuperaDatosDocDesembolso(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    sSql = "Select PersAna.cPersNombre cAnalista, CE.dPrdEstado dFecApr, Pers.cPersCod, "
    sSql = sSql & " Pers.cPersNombre, CE.dPrdEstado dFecDesemb, PersApod.cPersNombre cApoderado, CN.cConsDescripcion cMoneda,"
    sSql = sSql & " L.cDescripcion cLinea, P.nTasaInteres, CN2.cConsDescripcion cTipoDesemb, CN3.cConsDescripcion cTipoCred "
    'ssql = ssql & " cPersNombreUser = ( select cPersNombre from Persona where cPersCod = (select cPersCod from RRHH where cUser = '" & gsUser & "'))"
    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "                Inner Join Persona PersAna ON PersAna.cPersCod = PP.cPersCod "
    sSql = sSql & "                Inner Join ColocacEstado CE ON P.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstAprob
    sSql = sSql & "                Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                Inner Join Persona Pers ON Pers.cPersCod = PP2.cPersCod "
    sSql = sSql & "                Inner Join ColocacEstado CE2 ON P.cCtaCod = CE2.cCtaCod AND CE2.nPrdEstado = " & gColocEstVigNorm
    sSql = sSql & "                Inner Join ProductoPersona PP3 ON P.cCtaCod = PP3.cCtaCod AND PP3.nPrdPersRelac = " & gColRelPersApoderado
    sSql = sSql & "                Inner Join Persona PersApod ON PersApod.cPersCod = PP3.cPersCod "
    sSql = sSql & "                Inner Join Constante CN ON CN.nConsValor = Convert(Int, Substring(P.cCtaCod,9,1)) AND CN.nConsCod = " & gMoneda
    sSql = sSql & "                Inner Join Colocaciones C ON C.cCtacod = P.cCtaCod "
    sSql = sSql & "                Inner Join ColocLineaCredito L ON L.cLineaCred = C.cLineaCred "
    sSql = sSql & "                Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    sSql = sSql & "                Inner Join Constante CN2 ON CN2.nConsValor = CC.nTipoDesembolso AND CN2.nConsCod = " & gColocTiposDesembolso
    sSql = sSql & "                Left  Join Constante CN3 ON Convert(int, SubString(P.cCtaCod,6,3)) = CN3.nConsValor AND CN3.nConsCod = " & gProducto
    sSql = sSql & " Where P.cCtaCod = '" & psCtaCod & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosDocDesembolso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaDatosDocPlanPagos(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    sSql = "Select Pers.cPersNombre, PersAna.cPersNombre cAnalista, CN.cConsDescripcion cTipoCuota, "
    sSql = sSql & " nMontoCuota = (Select SUM(CD.nMonto) from ColocCalendDet CD where nNrocalen=(Select nNroCalen From ColocacCred Where cCtaCod = CD.cCtaCod) and cCtaCod = P.cCtaCod AND nColocCalendApl = 1 AND nCuota = 1), "
    sSql = sSql & " P.nTasaInteres, CN2.cConsDescripcion cMoneda, CE.nPlazo, CE.nPeriodoGracia, CE.nMonto, CE2.dVigencia dFecVig "
    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                INNER Join Colocaciones CE2 ON CE2.cCtaCod = P.cCtaCod "
    sSql = sSql & "                Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSql = sSql & "                Inner Join ProductoPersona PP2 ON PP2.cCtaCod = P.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "                Inner Join Persona PersAna ON PersAna.cPersCod = PP2.cPersCod"
    sSql = sSql & "                Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado = " & gColocEstAprob
    sSql = sSql & "                Inner Join Constante CN ON CN.nConsValor = CE.nColocCalendCod AND CN.nConsCod = " & gColocTipoCalend
    sSql = sSql & "                Inner Join Constante CN2 ON CN2.nConsValor = Convert(int,Substring(P.cCtaCod,9,1)) AND CN2.nConsCod = " & gMoneda
    sSql = sSql & " Where P.cCtaCod = '" & psCtaCod & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosDocPlanPagos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaDatosBoletaDesembolso(ByVal pnNroMov As Long, ByVal psCtaCod As String, ByVal psNomAge As String) As ADODB.Recordset
Dim oConecta As DConecta
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
    
    sSql = " Select P.cPersNombre,CASE WHEN SUBSTRING(CONVERT(VARCHAR(6),MC.nPrdConceptoCod),1,2) = 12 THEN SUM(MC.nMonto) END as nGasto, CASE WHEN SUBSTRING(CONVERT(VARCHAR(6),MC.nPrdConceptoCod),1,2) = 10 THEN SUM(MC.nMonto) END as nCapital, "
    sSql = sSql & "     FechaTran=dbo.FechaMov(M.cMovNro), HoraTran=dbo.HoraMov(M.cMovNro), CN.cConsDescripcion as sMoneda,       MC.nNrocuota, RIGHT(M.cMovNro,4) as sUsuario, Cal.dVenc,COUNT(MC.nMovNro) as nNumTran "
    sSql = sSql & " From ProductoPersona PP Inner Join Persona P ON PP.cPersCod = P.cPersCod And PP.nPrdPersRelac = 20 "
    sSql = sSql & "       Inner Join MovColDet MC ON MC.cCtaCod = PP.cCtacod "
    sSql = sSql & "       Inner Join Mov M ON M.nMovNro = MC.nMovnro "
    sSql = sSql & "       Inner Join Constante CN ON CN.nConsCod = " & gMoneda & " And CN.nConsValor = " & Mid(psCtaCod, 9, 1)
    sSql = sSql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = PP.cCtaCod And nColocCalendApl = 1 And nCuota = 1 and Cal.nNroCalen = (Select nNroCalen from ColocacCred where cCtaCod = PP.cCtaCod) "
    sSql = sSql & " Where PP.cCtaCod = '" & psCtaCod & "' And MC.nMovNro = " & pnNroMov & " and MC.cOpeCod='100101'"
    sSql = sSql & " Group By P.cPersNombre,  CN.cConsDescripcion, MC.nNrocuota, RIGHT(M.cMovNro,4), Cal.dVenc, MC.cOpeCod,M.cMovNro,SUBSTRING(CONVERT(VARCHAR(6),MC.nPrdConceptoCod),1,2)"
    
        
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosBoletaDesembolso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function


Public Function RecuperaDatosBoletaGastosDesembolso(ByVal pnNroMov As Long, ByVal psCtaCod As String, ByVal psNomAge As String) As ADODB.Recordset
Dim oConecta As DConecta
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
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosBoletaGastosDesembolso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaDatosBoletaCancelCredDesembolso(ByVal pnNroMov As Long, ByVal psCtaCod As String, ByVal psNomAge As String) As ADODB.Recordset
Dim oConecta As DConecta

Dim sSql As String

'    sSql = " Select P.cPersNombre,  CASE WHEN SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2) = '12' THEN SUM(MC.nMonto) END as nGasto, "
'    sSql = sSql & "           CASE WHEN SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2)= '10' THEN SUM(MC.nMonto)  END as nCapital,"
'    sSql = sSql & "          CASE WHEN SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2) = '11' THEN SUM(MC.nMonto) END as nInteres,"
'    sSql = sSql & "           FechaTran=dbo.FechaMov(M.cMovNro), HoraTran=dbo.HoraMov(M.cMovNro), CN.cConsDescripcion as sMoneda, COUNT(MC.nMovNro) as nNumTran,"
'    sSql = sSql & "           MC.nNrocuota, RIGHT(M.cMovNro,4) as sUsuario, Cal.dVenc, MCol.nSaldocap"
'    sSql = sSql & " From ProductoPersona PP Inner Join Persona P ON PP.cPersCod = P.cPersCod And PP.nPrdPersRelac = 20"
'    sSql = sSql & " Inner Join MovColDet MC ON MC.cCtaCod = PP.cCtacod"
'    sSql = sSql & " Inner Join MovCol MCol ON MCol.nMovnro = MC.nMovNro"
'    sSql = sSql & " Inner Join Mov M ON M.nMovNro = MC.nMovnro"
'    sSql = sSql & " Inner Join Constante CN ON CN.nConsCod = 1011 And CN.nConsValor = 1"
'    sSql = sSql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = PP.cCtaCod And nColocCalendApl = 1 And nCuota = 1 and Cal.nNrocalen = (select MAX(nNrocalen) From ColocCalendario Where cCtaCod = PP.cCtaCod)"
'    sSql = sSql & " Where PP.cCtaCod = '" & psCtaCod & "' And MC.nMovNro = (Select nMovNroRef From MovRef Where nMovNro = " & pnNroMov & ") And SUBSTRING(MC.cCtaCod,6,3)<>'232' "
'    sSql = sSql & " Group By P.cPersNombre, CN.cConsDescripcion, MC.nNrocuota, RIGHT(M.cMovNro,4), Cal.dVenc,SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2), M.cMovNro, Mcol.nSaldoCap"
    
    sSql = "Select MCol.cCtacod as cCtacod,   "
    sSql = sSql & " P.cPersNombre,"
    sSql = sSql & " nCapital=(Select SUM(MC.nMonto) From MovColdet MC Where MC.cCtaCod = MCol.cCtaCod AND MC.nMovNro = MCol.nMovNro AND MC.nNroCalen = MCol.nNroCalen AND SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2)= '10'),"
    sSql = sSql & " ninteres=(Select SUM(MC.nMonto) From MovColdet MC Where MC.cCtaCod = MCol.cCtaCod AND MC.nMovNro = MCol.nMovNro AND MC.nNroCalen = MCol.nNroCalen AND SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2)= '11'),"
    sSql = sSql & " nGastos=(Select SUM(MC.nMonto) From MovColdet MC Where MC.cCtaCod = MCol.cCtaCod AND MC.nMovNro = MCol.nMovNro AND MC.nNroCalen = MCol.nNroCalen AND SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2)= '12'),"
    sSql = sSql & " FechaTran=dbo.FechaMov(M.cMovNro), HoraTran=dbo.HoraMov(M.cMovNro),"
    sSql = sSql & " CN.cConsDescripcion as sMoneda,"
    sSql = sSql & " COUNT(MCol.nMovNro) as nNumTran,"
    sSql = sSql & " RIGHT(m.cMovNro, 4) As sUsuario"
    sSql = sSql & " From"
    sSql = sSql & " MovCol MCol"
    sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = MCol.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & " Inner Join Persona P ON P.cPersCod = PP.cPersCod"
    sSql = sSql & " Inner Join Mov M ON M.nMovNro = MCol.nMovnro"
    sSql = sSql & " Inner Join Constante CN ON CN.nConsCod = 1011 And CN.nConsValor = SUBSTRING(MCol.cCtaCod,9,1)"
    sSql = sSql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = PP.cCtaCod And nColocCalendApl = 1  and Cal.nNrocalen = (select MAX(nNrocalen) From ColocCalendario Where cCtaCod = PP.cCtaCod)"
    sSql = sSql & " Where MCol.nMovnro = " & pnNroMov & " AND MCol.cCtacod <> '" & psCtaCod & "'"
    sSql = sSql & " Group By MCol.cCtaCod, P.cPersNombre,M.cMovNro,CN.cConsDescripcion, MCol.nMovnro,MCol.nNroCalen "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosBoletaCancelCredDesembolso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function


Public Function ReportePagoDeCreditos(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim I As Integer
Dim sCadAge As String

    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias)
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
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
    sSql = sSql & "         From MovColDet Where nPrdConceptoCod = 1000"
    sSql = sSql & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD ON MCD.nMovNro = MC.nMovNro And MCD.cOpeCod = MC.cOpeCod And MCD.cCtaCod = MC.cCtaCod And MCD.nNroCalen = MC.nNroCalen And MCD.nPrdConceptoCod = 1000"
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
    sSql = sSql & " GROUP BY CN.cConsDescripcion, P.cPersNombre, O.cOpeDesc, MC.cCtaCod,M.nMovNro With ROLLUP"
    sSql = sSql & " ) as T"
    sSql = sSql & " Where T.nMovNro is not NULL or  T.cCtaCod = 'TOTAL'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set ReportePagoDeCreditos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosDesembolsados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda, ByVal pMatProducts As Variant) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim I As Integer
Dim sCadAge As String
Dim sCadProd As String
    
    sCadProd = " And SUBSTRING(MC.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProducts) - 1
        sCadProd = sCadProd & pMatProducts(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProducts) = 0 Then
        sCadProd = ""
    End If
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    sSql = "SELECT "
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
    sSql = sSql & "             WHERE   (MC1.cOpeCod >= '100101' and MC1.cOpeCod <= '100105') AND MC1.cCtacod=MC.cCtaCod and M1.nMovFlag=0 ) "
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
    sSql = sSql & " GROUP BY A.cAgeDescripcion, CN.cAbrev, P.cPersNombre, O.cOpeDesc, MC.cCtaCod "
    sSql = sSql & " With ROLLUP "
    sSql = sSql & " ORDER BY sAgencia, sProducto, sFuenteFinan, cOpeCod, nCredito"
    'sSql = sSql & " ORDER BY  CN.nConsValor "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosDesembolsados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function


Public Function RecuperaSaldoCarteraVigente(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String)
Dim sSql As String
Dim oConecta As DConecta
Dim I As Integer
Dim sCadAge As String
   
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
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

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaSaldoCarteraVigente = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
Public Function RecuperaCreditosCanceladosTodos(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda) As ADODB.Recordset
Dim sSql As String
Dim sCadAge As String
Dim I As Integer
Dim oConecta As DConecta

    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sSql = "SELECT "
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
    sSql = sSql & " With ROLLUP"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosCanceladosTodos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaCreditosCanceladosConsaldo(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sCadAge As String
Dim I As Integer
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
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

    Set oConecta = New DConecta
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
Dim oConecta As DConecta
Dim sCadAge As String
Dim I As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sAnalistas As String
Dim oGen As DGeneral

    Set oGen = New DGeneral
    sAnalistas = oGen.LeeConstSistema(gConstSistRHCargoCodAnalistas)
    Set oGen = Nothing

    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For I = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(I) & ","
    Next I
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
    
    Set oConecta = New DConecta
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
Dim oConecta As DConecta
Dim sCadAge As String
Dim I As Integer
Dim sCadCondicion As String
Dim sCadProd As String
    
    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For I = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(I) & ","
    Next I
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
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaMoraInstitucional = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaMoraxAnalista_AtrasoPagoCuotaLibre(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, _
    ByVal pMatAnalistas As Variant, ByVal pnConCuotaLibre As Boolean) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sCadAge As String
Dim I As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sCadAnalista As String

    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"

    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If

    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For I = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(I) & ","
    Next I
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"

    sCadAnalista = "('"
    For I = 0 To UBound(pMatAnalistas) - 1
        sCadAnalista = sCadAnalista & pMatAnalistas(I) & "','"
    Next I
    sCadAnalista = Mid(sCadAnalista, 1, Len(sCadAnalista) - 2) & ")"

 sSql = " Select CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
 sSql = sSql & "       ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')    END AS cAgencia,"
 sSql = sSql & "   CASE WHEN GROUPING(P2.cPersNombre)=1 THEN 'TOTAL'"
 sSql = sSql & "       ELSE ISNULL(P2.cPersNombre,'PERSONA DESCONOCIDA') END as cAnalista,"
 sSql = sSql & "   CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL'"
 sSql = sSql & "       ELSE ISNULL(Prd.cCtaCod,'CUENTA DESCONOCIDA') END as cCtaCod,"
 sSql = sSql & "   dbo.ColocCredTitular(Prd.cCtaCod) + ' ' + dbo.ColocCredTitularDir(Prd.cCtaCod) + ' ' + RIGHT(SPACE(10) + CONVERT(varchar(20),C.nMontoCol),10) +  ' ' + RIGHT(SPACE(10) + CONVERT(varchar(20),Prd.nSaldo),10) +  ' ' + RIGHT(SPACE(10) + CONVERT(varchar(15),dbo.CuotaAprobada(Prd.cCtaCod)),10) + '   ' + CONVERT(varchar(10),CCal.nCuota) + '    ' +  CONVERT(varchar(15),CCal.dvenc,103) as nDatos,"
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
 sSql = sSql & " (dbo.ColocCredTitular(Prd.cCtaCod) + ' ' + dbo.ColocCredTitularDir(Prd.cCtaCod) + ' ' + RIGHT(SPACE(10) + CONVERT(varchar(20),C.nMontoCol),10) +  ' ' + RIGHT(SPACE(10) + CONVERT(varchar(20),Prd.nSaldo),10) +  ' ' + RIGHT(SPACE(10) + CONVERT(varchar(15),dbo.CuotaAprobada(Prd.cCtaCod)),10) + '   ' + CONVERT(varchar(10),CCal.nCuota) + '    ' +  CONVERT(varchar(15),CCal.dvenc,103))"
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
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaMoraxAnalista_AtrasoPagoCuotaLibre = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosProtestados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta
Dim sCadAge As String
Dim I As Integer

    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
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

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosProtestados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosRetirados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant) As ADODB.Recordset

Dim oConecta As DConecta
Dim sCadAge As String
Dim I As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sSql As String

    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For I = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(I) & ","
    Next I
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
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosRetirados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaCartaCreditosMorosos(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, ByVal pnDiasAtrIni As Integer, _
    ByVal pnDiasAtrFin As Integer, ByVal pMatAnalistas As Variant) As ADODB.Recordset
Dim oConecta As DConecta
Dim sCadAge As String
Dim I As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sSql As String
Dim sCadAnalista As String

    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For I = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(I) & ","
    Next I
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"

    sCadAnalista = "('"
    For I = 0 To UBound(pMatAnalistas) - 1
        sCadAnalista = sCadAnalista & pMatAnalistas(I) & "','"
    Next I
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
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCartaCreditosMorosos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCartaInvitacionCreditosParalelos(ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, ByVal pnNotaIni As Integer, _
    ByVal pnNotaFin As Integer, ByVal pMatAnalistas As Variant, ByVal pnUltCuotas As Integer, ByVal pbPorc As Integer) As ADODB.Recordset
Dim oConecta As DConecta
Dim sCadAge As String
Dim I As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sSql As String
Dim sCadAnalista As String

    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For I = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(I) & ","
    Next I
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"

    sCadAnalista = "('"
    For I = 0 To UBound(pMatAnalistas) - 1
        sCadAnalista = sCadAnalista & pMatAnalistas(I) & "','"
    Next I
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
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCartaInvitacionCreditosParalelos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosPorUbicacionGeo(ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, _
    ByVal psCodUbic As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim sCadAge As String
Dim I As Integer
Dim sCadCondicion As String
Dim sCadProd As String
    
    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For I = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(I) & ","
    Next I
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
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosPorUbicacionGeo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCredVig_CredVigConCuotLib(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, ByVal pnDiasAtrIni As Integer, _
    ByVal pnDiasAtrFin As Integer, ByVal pnTipoCambio As Double, ByVal pnConCuotaLibre As Boolean) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim I As Integer
Dim sCadAge As String
Dim sCadProd As String
Dim sCadCondicion As String
   
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    
    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sCadCondicion = "("
    For I = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(I) & ","
    Next I
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

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCredVig_CredVigConCuotLib = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosXInstitucion(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, _
    ByVal pdFecIni As Date, ByVal pnTipoOrden As Integer, pnIncluyeMora As Integer) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim I As Integer
Dim sCadAge As String
Dim sCadMora As String
Dim sCadTipoOrden As String
   
   sCadMora = ""
   If pnIncluyeMora = 0 Then
        sCadMora = "And nPrdConceptoCod <> 1101"
   End If
    
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
        
    sCadTipoOrden = ""
    Select Case pnTipoOrden
        Case 0 'Por Codigo Modular
            sCadTipoOrden = "CV.cCodModular"
        Case 1 'Orden Alfabetico
            sCadTipoOrden = "P.cPersNombre"
        Case 2 'Pagare
            sCadTipoOrden = "Prd.cCtaCod"
    End Select
    
    sSql = " SELECT "
    sSql = sSql & "  CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSql = sSql & "   END AS sAgencia,"
    sSql = sSql & "   CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL'"
    sSql = sSql & "    ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDA')"
    sSql = sSql & "   END AS cCtaCod,"
    sSql = sSql & "   CASE WHEN GROUPING(P.cPersNombre)=1 THEN 'TOTAL'"
    sSql = sSql & "    ELSE ISNULL(P.cPersNombre, 'INSTITUCION DESCONOCIDA')"
    sSql = sSql & "   END AS cInstitucion,"
    sSql = sSql & " COUNT(*) as nCasos, "
    sSql = sSql & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = Prd.cCtaCod),"
    sSql = sSql & " cCodModular = (Select cCodModular From ColocacConvenio CV Where CV.cCtaCod = Prd.cCtaCod),"
    sSql = sSql & " dFecDesemb = (Select dVigencia From Colocaciones Where cCtaCod = Prd.cCtaCod),"
    sSql = sSql & " nDiasAtraso = (Select nDiasAtraso From ColocacCred Where cCtaCod = Prd.cCtaCod), "
    sSql = sSql & " SUM(CE.nMonto) as nMontoApr,"
    sSql = sSql & " SUM(Prd.nSaldo) as nSaldoCap,"
    sSql = sSql & " sCuotaPend = (Select CONVERT(varchar(5),CC.nNroProxCuota) + '/' + CONVERT(varchar(5),CE.nCuotas) From ColocacCred CC Inner Join ColocacEstado CE ON CC.cCtaCod = CE.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & "        Where CC.cCtaCod = Prd.cCtaCod ),"
    sSql = sSql & " SUM(CDet.nMonto) as nCuota,"
    sSql = sSql & " dFecPago = (Select dVenc From ColocCalendario CCal Where CCal.cCtaCod = Prd.cCtaCod And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CCal.cCtaCod) And nColocCalendApl = 1 And nCuota = (Select nNroProxCuota From ColocacCred Where cCtaCod = CCal.cCtaCod))"
    sSql = sSql & " From Producto Prd"
    sSql = sSql & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & " Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSql = sSql & " Inner Join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Inner Join (Select cCtaCod, nNroCalen, nCuota,  SUM(nMonto) as nMonto From ColocCalendDet CDet"
    sSql = sSql & "        Where nColocCalendApl = 1 " & sCadMora
    sSql = sSql & "        Group By cCtaCod, nNroCalen, nCuota"
    sSql = sSql & "        ) as CDet ON CDet.cCtaCod = Prd.cCtaCod And CDet.nNroCalen = CC.nNroCalen And CDet.nCuota = CC.nNroProxCuota"
    sSql = sSql & " Where Prd.nPrdEstado >= 2020 And Prd.nPrdEstado <= 2032"
    sSql = sSql & " And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' And A.cAgeCod IN " & sCadAge
    sSql = sSql & " Group By A.cAgeDescripcion, P.cPersNombre, Prd.cCtaCod"
    sSql = sSql & " WITH ROLLUP"
    sSql = sSql & " Order by sAgencia, cCtaCod, cInstitucion "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosXInstitucion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaMoraXInstituciones(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, _
    ByVal pdFecIni As Date, pMatInstitucion As Variant) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim I As Integer
Dim sCadAge As String
Dim sCadInstitucion As String

    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    sCadInstitucion = "('"
    For I = 0 To UBound(pMatInstitucion) - 1
        sCadInstitucion = sCadInstitucion & pMatInstitucion(I) & "','"
    Next I
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

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaMoraXInstituciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaResumenSaldosCarteraPorInstitucionConsumo(ByVal pMatAgencias As Variant, _
    ByVal pMatProd As Variant, ByVal pdHoy As Date) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim sCadAge As String
Dim I As Integer
Dim sCadCondicion As String
Dim sCadProd As String
    
    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
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
    
    Set oConecta = New DConecta
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
Dim oConecta As DConecta
Dim sCadAge As String
Dim I As Integer
Dim sCadCondicion As String
Dim sCadProd As String

    
    sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For I = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(I) & ","
    Next I
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
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaResumenSaldosCarteraPorAnalistaConsumo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaPagoConCheque(ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, _
    ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pnTipoBusq As Integer, ByVal psNroDoc As String) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim sCadAge As String
Dim I As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sCadBusq As String

    If pnTipoBusq = 1 Then
        sCadBusq = " AND MD.cDocNro = '" & psNroDoc & "'"
    End If

    sCadProd = " And SUBSTRING(MC.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sCadCondicion = "("
    For I = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(I) & ","
    Next I
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
    sSql = sSql & "  And SUBSTRING(MC.cCtaCod, 9, 1) = '" & pnMoneda & "' "
    sSql = sSql & " Order by A.cAgeDescripcion, P.cPersNombre"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaPagoConCheque = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaPagoDeOtrasAgencias(ByVal pnMoneda As Moneda, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, psCodAge As String, ByVal psCodCmact As String) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta

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
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaPagoDeOtrasAgencias = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaPagoENOtrasAgencias(ByVal pnMoneda As Moneda, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, psCodAge As String, ByVal psCodCmact As String, _
    ByVal pMatCondicion As Variant, ByVal pMatProd As Variant) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim sCadCondicion As String
Dim sCadProd As String
Dim I As Integer
    sCadProd = " And SUBSTRING(MC.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
    
    sCadCondicion = "("
    For I = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(I) & ","
    Next I
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
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaPagoENOtrasAgencias = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function


Public Function RecuperaInteresesEnSuspenso(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatProd As Variant)
    
Dim sSql As String
Dim oConecta As DConecta
Dim I As Integer
Dim sCadAge As String
Dim sCadProd As String

    sCadProd = " And SUBSTRING(MC.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
   
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
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

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaInteresesEnSuspenso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
 
Public Function RecuperaCreditosRefinanciados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatProd As Variant)
    
Dim sSql As String
Dim oConecta As DConecta
Dim I As Integer
Dim sCadAge As String
Dim sCadProd As String

    
    sCadProd = " And SUBSTRING(MC.cCtaCod,6,3) in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If
   
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
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
    
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosRefinanciados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
 
Public Function RecuperaImpLavDineroM(ByVal psFechaIni As String, ByVal psFechaFin As String, ByVal psAgencias As String, ByVal psMonedas As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As DConecta
    
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
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaImpLavDineroM = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaTCLavDineroM(ByVal pdFechaIni As String) As Double
    Dim sSql As String
    Dim oConecta As DConecta
    Dim rs As New ADODB.Recordset
    
    sSql = "SELECT TOP 1 nValPond FROM TIPOCAMBIO WHERE datediff(d, dateadd(d," & Mid(pdFechaIni, 1, 2) * -1 & ", '" _
        & Format(pdFechaIni, gsFormatoFecha) & "'),  dFecCamb) = 0 ORDER BY dFecCamb "

    Set oConecta = New DConecta
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
