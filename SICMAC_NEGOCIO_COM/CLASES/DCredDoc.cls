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
Dim sql As String
Dim Rs As New ADODB.Recordset
Dim RCD As nRcdReportes
Set RCD = New nRcdReportes

sql = " Select Pro.cCtaCod ,P1.cPersCod, P1.cPersNombre Nombre, Pro.nSaldo Saldo from Producto Pro"
sql = sql & " Inner Join ProductoPersona PP1 on PP1.cCtaCod = Pro.cCtaCod"
sql = sql & " Inner Join Persona P1 on P1.cPersCod = PP1.cPersCod"
sql = sql & " Where PP1.nPrdPersRelac = 20"
sql = sql & " and Pro.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
sql = sql & " and Pro.cCtaCod='" & psCtaCod & "'"
Set Rs = RCD.GetQuery(sql)
Set RecuperaDatosMantCargoAutomatico = Rs
Set Rs = Nothing
Set RCD = Nothing
End Function
Public Function RecuperaDatosColocCargoAutoma(ByVal psCtaCod As String) As ADODB.Recordset
Dim sql As String
Dim Rs As New ADODB.Recordset
Dim RCD As nRcdReportes
Set RCD = New nRcdReportes

sql = " Select cCtaCod,nOrden,cCodCtaAho,cEstado  from ColocCargoAutoma"
sql = sql & " where cCtaCod='" & psCtaCod & "' and substring(cCtaCod,9,1)='" & Mid(psCtaCod, 9, 1) & "'"
sql = sql & " Order by cEstado, nOrden"

Set Rs = RCD.GetQuery(sql)
Set RecuperaDatosColocCargoAutoma = Rs
Set Rs = Nothing
Set RCD = Nothing
End Function
Public Sub ActualizaColocCargoAutoma(ByVal psCtaCod As String, ByVal psCtaCodAho As String, _
ByVal pnOrden As Integer, ByVal pnEstado As Integer)
Dim sql As String
Dim Co As DConecta
Set Co = New DConecta
sql = " Update ColocCargoAutoma"
sql = sql & " Set nOrden=" & pnOrden & ", cEstado = " & pnEstado
sql = sql & " where "
sql = sql & " cCtaCod ='" & Trim(psCtaCod) & "'"
sql = sql & " and cCodCtaAho ='" & Trim(psCtaCodAho) & "'"
Co.AbreConexion
Call Co.Ejecutar(sql)
Co.CierraConexion
Set Co = Nothing

End Sub
Public Sub DeleteColocCargoAutoma(ByVal psCtaCod As String, ByVal psCtaCodAho As String)
Dim sql As String
Dim Co As DConecta
Set Co = New DConecta
sql = " Delete ColocCargoAutoma"
sql = sql & " where cEstado = '1'"
sql = sql & " and cCtaCod ='" & Trim(psCtaCod) & "'"
sql = sql & " and cCodCtaAho ='" & Trim(psCtaCodAho) & "'"
Co.AbreConexion
Call Co.Ejecutar(sql)
Co.CierraConexion
Set Co = Nothing
End Sub

Public Function ExisteCtaCargoAutoma(ByVal psCtaCod As String, ByVal psCtaCodAho As String) As Boolean
Dim sql As String
Dim Rs As New ADODB.Recordset
Dim RCD As nRcdReportes
Set RCD = New nRcdReportes

sql = " Select * from ColocCargoAutoma"
sql = sql & " where cEstado = '1' "
sql = sql & " and cCtaCod ='" & Trim(psCtaCod) & "' "
sql = sql & " and cCodCtaAho ='" & Trim(psCtaCodAho) & "' "
Set Rs = RCD.GetQuery(sql)
If Rs.EOF And Rs.BOF Then
    ExisteCtaCargoAutoma = False
Else
    ExisteCtaCargoAutoma = True
End If
Set Rs = Nothing
Set RCD = Nothing
End Function


Public Sub InsertaColocCargoAutoma(ByVal psCtaCod As String, _
ByVal psCtaCodAho As String, ByVal psEstado As String, ByVal pnOrden As Integer)
Dim Co As DConecta
Dim sql As String
Set Co = New DConecta
sql = " Insert ColocCargoAutoma "
sql = sql & " (cCtaCod,nOrden,cCodCtaAho,cEstado) Values "
sql = sql & " ('" & psCtaCod & "'," & pnOrden & ",'" & psCtaCodAho & "','" & psEstado & "')"
Co.AbreConexion
Co.Ejecutar (sql)
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
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Co As DConecta
    Set Co = New DConecta
    sql = "select cPersNombre from Persona where cPersCod = (select cPersCod from RRHH where cUser = '" & psCodUser & "')"
    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(sql)
    Co.CierraConexion
    If Rs.BOF Then
        ObtieneNombrePersonaUser = ""
    Else
        ObtieneNombrePersonaUser = Rs!cpersnombre
    End If

    Set Co = Nothing
    Set Rs = Nothing

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
    sSql = sSql & " Where PP.cCtaCod = '" & psCtaCod & "' And MC.nMovNro = " & pnNroMov & " and MC.cOpeCod in ('100101','100102','100103','100104','100105','100106','100107','100109')"
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

    sSql = "Select MCol.cCtacod as cCtacod,   "
    sSql = sSql & " P.cPersNombre,"
    sSql = sSql & " nCapital=(Select SUM(MC.nMonto) From MovColdet MC Where MC.cCtaCod = MCol.cCtaCod AND MC.nMovNro = MCol.nMovNro AND MC.nNroCalen = MCol.nNroCalen AND SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2)= '10'),"
    sSql = sSql & " ninteres=(Select SUM(MC.nMonto) From MovColdet MC Where MC.cCtaCod = MCol.cCtaCod AND MC.nMovNro = MCol.nMovNro AND MC.nNroCalen = MCol.nNroCalen AND SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2)= '11'),"
    sSql = sSql & " nGastos=(Select SUM(MC.nMonto) From MovColdet MC Where MC.cCtaCod = MCol.cCtaCod AND MC.nMovNro = MCol.nMovNro AND MC.nNroCalen = MCol.nNroCalen AND SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2)= '12'),"
    sSql = sSql & " nItf =(Select SUM(MC.nMonto) From MovColdet MC Where MC.cCtaCod = MCol.cCtaCod AND MC.nMovNro = MCol.nMovNro AND MC.nPrdConceptoCod = 20 ), FechaTran=dbo.FechaMov(M.cMovNro), "
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
    sSql = sSql & " Where MCol.nMovnro = " & pnNroMov & " AND MCol.cCtacod <> '" & psCtaCod & "' and MCOL.cOpeCod like '100[234567]%' "
    sSql = sSql & " Group By MCol.cCtaCod, P.cPersNombre,M.cMovNro,CN.cConsDescripcion, MCol.nMovnro,MCol.nNroCalen "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosBoletaCancelCredDesembolso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function


Public Function ReportePagoDeCreditos(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda, ByVal pnRFA As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim i As Integer
Dim sCadAge As String

    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias)
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    sSql = "Select MC.cCtaCod, P.cPersnombre, MC.nMonto, "
    sSql = sSql & " (Select ISNULL(SUM(nMonto),0) from MovColDet Where nMovnro = MC.nMovNro and nPrdConceptoCod=1000 AND cOpeCod like '100[234567]%') as Capital, "
    sSql = sSql & " (Select ISNULL(SUM(nMonto),0) from MovColDet Where nMovnro = MC.nMovNro and nPrdConceptoCod<>1101 and nPrdConceptoCod like '11%' AND cOpeCod like '100[234567]%') as Interes,"
    sSql = sSql & " (Select ISNULL(SUM(nMonto),0) from MovColDet Where nMovnro = MC.nMovNro and nPrdConceptoCod=1101 AND cOpeCod like '100[234567]%') as Mora,"
    sSql = sSql & " (Select ISNULL(SUM(nMonto),0) from MovColDet Where nMovnro = MC.nMovNro and nPrdConceptoCod like '12%' AND cOpeCod like '100[234567]%') as Gastos"
    sSql = sSql & " From Mov M"
    sSql = sSql & " Inner Join MovCol MC ON M.nMoVnro = MC.nMoVnro"
    sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = MC.cCtaCod and PP.nPrdpersrelac=20"
    sSql = sSql & " Inner Join Persona P ON P.cPerscod = PP.cPerscod"
    sSql = sSql & " Where SUBSTRING(M.cMovNro,1,8)  >= '" & Format(pdFecIni, "yyyymmdd") & "' AND  SUBSTRING(M.cMovNro,1,8)  <= '" & Format(pdFecFin, "yyyymmdd") & "'  and MC.cOpeCod like '100[234567]%'"
    sSql = sSql & "     AND SUBSTRING(MC.cCtaCod,9,1)= '" & pnMoneda & "'"
    sSql = sSql & "      AND M.nMovFlag=0 AND SUBSTRING(MC.cCtaCod,4,2) in " & sCadAge
    sSql = sSql & " Order by MC.nMonto"

    
    
    
'**************

'    '---------
'    sSQL = " Select * From ("
'    sSQL = sSQL & " SELECT"
'    sSQL = sSQL & " CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'  ELSE  ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')   END AS sProducto,"
'    sSQL = sSQL & "        CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')   END AS sFuenteFinan,"
'    sSQL = sSQL & "        CASE WHEN (GROUPING(O.cOpeDesc) = 1) THEN 'TOTAL '        ELSE ISNULL(O.cOpeDesc, 'OPERACION DESCONOCIDA')   END AS cOpeCod,"
'    sSQL = sSQL & "        CASE WHEN (GROUPING(MC.cCtaCod) = 1) THEN 'TOTAL'        ELSE ISNULL(MC.cCtaCod, 'DESCONOCIDO')   END AS cCtaCod,"
'    sSQL = sSQL & " M.nMovNro,"
'    sSQL = sSQL & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = MC.cCtaCod),"
'    sSQL = sSQL & " sum(isnull(MCD.nMonto, 0)) as  nCapital,"
'    sSQL = sSQL & " sum(isnull(MCD2.nMonto, 0)) as nInteres,"
'    sSQL = sSQL & " sum(isnull(MCD4.nMonto, 0)) as nInteresMora,"
'    sSQL = sSQL & " sum(isnull(MCD3.nMonto, 0)) as nInteresGracia,"
'    sSQL = sSQL & " sum(isnull(MCD5.nMonto, 0)) as nInteresReprog,"
'    sSQL = sSQL & " sum(isnull(MCD6.nMonto, 0)) as nInteresSuspenso,"
'    sSQL = sSQL & " sum(IsNull(MCD7.nMonto, 0)) As nGastos"
'    sSQL = sSQL & " FROM MovCol MC"
'    sSQL = sSQL & " Inner Join Constante CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = 1001"
'    sSQL = sSQL & " Inner Join OpeTpo O ON O.cOpeCod = MC.cOpeCod"
'    sSQL = sSQL & " Inner Join Colocaciones C ON C.cCtaCod = MC.cCtaCod"
'    sSQL = sSQL & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = C.cLineaCred"
'    sSQL = sSQL & " Inner Join Persona P ON CL.cPersCod = P.cPersCod"
'    sSQL = sSQL & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, nPrdConceptoCod, SUM(nMonto) as nMonto"
'    sSQL = sSQL & "         From MovColDet Where nPrdConceptoCod IN(1000,1109) "
'    sSQL = sSQL & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD ON MCD.nMovNro = MC.nMovNro And MCD.cOpeCod = MC.cOpeCod And MCD.cCtaCod = MC.cCtaCod And MCD.nNroCalen = MC.nNroCalen And MCD.nPrdConceptoCod IN (1000,1109)"
'    sSQL = sSQL & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, 1100 as nPrdConceptoCod, SUM(nMonto) as nMonto"
'    sSQL = sSQL & "         From MovColDet Where nPrdConceptoCod in (1100,1105) "
'    sSQL = sSQL & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen) MCD2 ON MCD2.nMovNro = MC.nMovNro And MCD2.cOpeCod = MC.cOpeCod And MCD2.cCtaCod = MC.cCtaCod And MCD2.nNroCalen = MC.nNroCalen And MCD2.nPrdConceptoCod = 1100"
'    sSQL = sSQL & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, nPrdConceptoCod, SUM(nMonto) as nMonto"
'    sSQL = sSQL & "         From MovColDet Where nPrdConceptoCod = 1102"
'    sSQL = sSQL & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD3 ON MCD3.nMovNro = MC.nMovNro And MCD3.cOpeCod = MC.cOpeCod And MCD3.cCtaCod = MC.cCtaCod And MCD3.nNroCalen = MC.nNroCalen And MCD3.nPrdConceptoCod = 1102"
'    sSQL = sSQL & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, nPrdConceptoCod, SUM(nMonto) as nMonto"
'    sSQL = sSQL & "         From MovColDet Where nPrdConceptoCod = 1101"
'    sSQL = sSQL & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD4 ON MCD4.nMovNro = MC.nMovNro And MCD4.cOpeCod = MC.cOpeCod And MCD4.cCtaCod = MC.cCtaCod And MCD4.nNroCalen = MC.nNroCalen And MCD4.nPrdConceptoCod = 1101"
'    sSQL = sSQL & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, nPrdConceptoCod, SUM(nMonto) as nMonto"
'    sSQL = sSQL & "         From MovColDet Where nPrdConceptoCod = 1103"
'    sSQL = sSQL & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD5 ON MCD5.nMovNro = MC.nMovNro And MCD5.cOpeCod = MC.cOpeCod And MCD5.cCtaCod = MC.cCtaCod And MCD5.nNroCalen = MC.nNroCalen And MCD5.nPrdConceptoCod = 1103"
'    sSQL = sSQL & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, nPrdConceptoCod, SUM(nMonto) as nMonto"
'    sSQL = sSQL & "         From MovColDet Where nPrdConceptoCod = 1104"
'    sSQL = sSQL & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD6 ON MCD6.nMovNro = MC.nMovNro And MCD6.cOpeCod = MC.cOpeCod And MCD6.cCtaCod = MC.cCtaCod And MCD6.nNroCalen = MC.nNroCalen And MCD6.nPrdConceptoCod = 1104"
'    sSQL = sSQL & " Left Join ( Select nMovNro, cOpeCod, cCtaCod, nNroCalen, Convert(int, '1200') as nPrdConceptoCod, SUM(nMonto) as nMonto"
'    sSQL = sSQL & "         From MovColDet Where SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2) = '12'"
'    sSQL = sSQL & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen, SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2))  MCD7  ON MCD7.nMovNro = MC.nMovNro And MCD7.cOpeCod = MC.cOpeCod And MCD7.cCtaCod = MC.cCtaCod And MCD7.nNroCalen = MC.nNroCalen And MCD7.nPrdConceptoCod = 1200"
'    sSQL = sSQL & " Inner Join Mov M ON M.nMovNro = MC.nMovNro And M.nMovFlag = 0  "
'    sSQL = sSQL & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'  and Len(MC.cCtaCod) = 18 And (O.cOpeCod >= '100200' and O.cOpeCod <= '100707')"
'    sSQL = sSQL & " And SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And SUBSTRING(MC.cCtaCod,4,2) IN " & sCadAge
'    'ssql = ssql & " And SUBSTRING(M.cMovNro,1,8) =  '" & Format(pdFecFin, "yyyymmdd") & "' "
'    sSQL = sSQL & " And SUBSTRING(M.cMovNro,1,8) Between  '" & Format(pdFecIni, "yyyymmdd") & "' AND '" & Format(pdFecFin, "yyyymmdd") & "' "
'    If pnRFA = 1 Then
'        sSQL = sSQL & " AND MC.cCtaCod NOT IN ( "
'        sSQL = sSQL & "     Select cCtaCod From ColocacCred Where cRFA = 'RFA' "
'        sSQL = sSQL & "     ) "
'    End If
'    sSQL = sSQL & " GROUP BY CN.cConsDescripcion, P.cPersNombre, O.cOpeDesc, MC.cCtaCod,M.nMovNro With ROLLUP"
'    sSQL = sSQL & " ) as T"
'    sSQL = sSQL & " Where T.nMovNro is not NULL or  T.cCtaCod = 'TOTAL'"
'    sSQL = sSQL & " Order by T.sProducto, T.cCtaCod, T.sTitular "


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
    
'    ssql = ssql & " SELECT * FROM ("
'    ssql = ssql & "SELECT "
'    ssql = ssql & "  CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
'    ssql = ssql & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
'    ssql = ssql & "   END AS sAgencia,"
'    ssql = ssql & "   CASE WHEN (GROUPING(CN.cAbrev) = 1) THEN 'TOTAL'"
'    ssql = ssql & "        ELSE ISNULL(CN.cAbrev, 'PRODUCTO DESCONOCIDO ')"
'    ssql = ssql & "   END AS sProducto,"
'    ssql = ssql & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
'    ssql = ssql & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
'    ssql = ssql & "   END AS sFuenteFinan,"
'    ssql = ssql & "   CASE WHEN (GROUPING(O.cOpeDesc) = 1) THEN 'TOTAL '"
'    ssql = ssql & "        ELSE ISNULL(O.cOpeDesc, 'OPERACION DESCONOCIDA')"
'    ssql = ssql & "   END AS cOpeCod,"
'    ssql = ssql & "   CASE WHEN (GROUPING(MC.cCtaCod) = 1) THEN 'TOTAL'"
'    ssql = ssql & "        ELSE ISNULL(MC.cCtaCod, 'CUENTA DESCONOCIDO ')"
'    ssql = ssql & "   END AS nCredito,"
'    ssql = ssql & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = " & gColRelPersTitular & " Where PP.cCtaCod = MC.cCtaCod),"
'    'sSql = sSql & " SUM(CE.nMonto) as nMontoApr,"
'    ssql = ssql & " SUM(isnull(MCD8.nMonto,0)) as nDesembolso,"
'    ssql = ssql & " SUM(CASE WHEN MCD.nNroCuota >1 THEN 0 ELSE  CAL.nMonto END) as nIntApr, "
'    ssql = ssql & " SUM(CASE WHEN MCD.nNroCuota >1 THEN 0 ELSE  CE.nMonto END) as nMontoApr, "
'    ssql = ssql & " SUM(MCD7.nMonto) as Gasto,"
'    ssql = ssql & " sAnalista = (Select Top 1 cUser From RRHH Where cPersCod = (select cPersCod From ProductoPersona PP2 Where PP2.cCtaCod = MC.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersAnalista & ")),"
'    ssql = ssql & " nTasa = (Select P.nTasaInteres From Producto P Where P.cCtacod = MC.cCtaCod),  "
'    ssql = ssql & " cUserDes = (SELECT  RIGHT(CMOVNRO,4)  "
'    ssql = ssql & "             FROM    MOV M1  JOIN MOVCOL MC1 ON MC1.NMOVNRO = M1.NMOVNRO "
'    ssql = ssql & "             WHERE   (MC1.cOpeCod >= '100101' and MC1.cOpeCod <= '100105') AND MC1.cCtacod=MC.cCtaCod and M1.nMovFlag=0 And MC1.nMovNro=MC.nMovNro) "
'    ssql = ssql & " FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro AND M.nMovFlag<> 2"
'    ssql = ssql & " Inner Join ("
'    ssql = ssql & "     Select nMovNro, cCtaCod, nNroCuota From MovColDet Where  cOpeCod in ('100101','100102','100103','100104','100105') Group by nMovNro, cCtaCod, nNroCuota"
'    ssql = ssql & " ) MCD ON MCD.nMovNro = MC.nMovNro AND  MCD.cCtaCod = MC.cCtaCod"
'    'sSql = sSql & " Inner Join Constante CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
'    ssql = ssql & " INNER JOIN (Select nConsCod,    nConsValor,  cConsDescripcion, (CONVERT(varchar(10),nConsValor)  + ' ' + cConsDescripcion ) as cAbrev "
'    ssql = ssql & "    From Constante ) as CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
'    ssql = ssql & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(MC.cCtaCod,4,2)"
'    ssql = ssql & " Inner Join OpeTpo O ON O.cOpeCod = MC.cOpeCod"
'    ssql = ssql & " Inner Join Colocaciones CC ON CC.cCtaCod = MC.cCtaCod"
'    ssql = ssql & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = CC.cLineaCred"
'    ssql = ssql & " Inner Join Persona P ON CL.cPersCod = P.cPersCod "
'    ssql = ssql & " Left Join ( Select nMovNro, cOpeCod, cCtaCod, nNroCalen, Convert(int, '1200') as nPrdConceptoCod, SUM(nMonto) as nMonto From MovColDet"
'    ssql = ssql & "       Where SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2) = '12'"
'    ssql = ssql & "       Group By nMovNro, cOpeCod, cCtaCod, nNroCalen, SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2))  MCD7"
'    ssql = ssql & "    ON MCD7.nMovNro = MC.nMovNro And MCD7.cOpeCod = MC.cOpeCod And MCD7.cCtaCod = MC.cCtaCod And MCD7.nNroCalen = MC.nNroCalen And MCD7.nPrdConceptoCod = 1200 "
'    ssql = ssql & " Left Join ( Select nMovNro, cOpeCod, cCtaCod, nNroCalen, Convert(int, '1000') as nPrdConceptoCod, SUM(nMonto) as nMonto From MovColDet"
'    ssql = ssql & "       Where nPrdConceptoCod = 1000 AND (cOpeCod >= '" & gCredDesembEfec & "' and cOpeCod <= '" & gCredDesembCtaNuevaDOA & "') "
'    ssql = ssql & "       Group By nMovNro, cOpeCod, cCtaCod, nNroCalen, SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2))  MCD8"
'    ssql = ssql & "    ON MCD8.nMovNro = MC.nMovNro And MCD8.cOpeCod = MC.cOpeCod And MCD8.cCtaCod = MC.cCtaCod And MCD8.nNroCalen = MC.nNroCalen And MCD8.nPrdConceptoCod = 1000 "
'    ssql = ssql & " Inner Join ColocacEstado CE ON CE.cCtaCod = MC.cCtaCod And CE.nPrdEstado = " & gColocEstAprob
'    ssql = ssql & " Inner Join ColocacCred C ON C.cCtaCod = MC.cCtaCod"
'    ssql = ssql & " Inner Join ( Select CD.cCtaCod, CD.nNrocalen, SUM(CD.nMonto) as nMonto From ColocCalendDet CD"
'    ssql = ssql & "        Where CD.nColocCalendApl = 1 And CD.nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio
'    ssql = ssql & "        GROUP BY CD.cCtaCod, CD.nNrocalen) CAL ON CAL.cCtaCod = MC.cCtaCod And CAL.nNroCalen = C.nNroCalen"
'    'ssql = ssql & " Inner Join Mov M ON M.nMovNro = MC.nMovNro "
'    ssql = ssql & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'  and Len(MC.cCtaCod) = 18 And (MC.cOpeCod >= '" & gCredDesembEfec & "' and MC.cOpeCod <= '" & gCredDesembCtaNuevaDOA & "') And SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And SUBSTRING(MC.cCtaCod,4,2) IN " & sCadAge
'    ssql = ssql & "     And SUBSTRING(M.cMovNro,1,8) Between  '" & Format(pdFecIni, "yyyymmdd") & "' AND '" & Format(pdFecFin, "yyyymmdd") & "' " & sCadProd
'    ssql = ssql & "  "
'    ssql = ssql & " GROUP BY A.cAgeDescripcion, CN.cAbrev, P.cPersNombre, O.cOpeDesc, MC.cCtaCod,MC.nMovNro "
'    ssql = ssql & " With ROLLUP)X "
'    ssql = ssql & " Where (cUserDes Is Not Null ) OR nCredito='TOTAL'"    'ARCV 23-06-2007
'    ssql = ssql & " ORDER BY sAgencia, sProducto, sFuenteFinan, cOpeCod, nCredito"
    
    sSql = sSql & " SELECT * FROM ("
    sSql = sSql & "SELECT "
    sSql = sSql & "  CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSql = sSql & "   END AS sAgencia,"
    sSql = sSql & "   CASE WHEN (GROUPING(CN.cAbrev) = 1) THEN 'TOTAL'"
    sSql = sSql & "        ELSE ISNULL(CN.cAbrev, 'PRODUCTO DESCONOCIDO ')"
    sSql = sSql & "   END AS sProducto,"
    sSql = sSql & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'ZTOTAL'"
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
    sSql = sSql & " FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro AND M.nMovFlag<> 2"
    sSql = sSql & " Inner Join ("
    sSql = sSql & "     Select nMovNro, cCtaCod, nNroCuota From MovColDet Where  cOpeCod in ('100101','100102','100103','100104','100105') Group by nMovNro, cCtaCod, nNroCuota"
    sSql = sSql & " ) MCD ON MCD.nMovNro = MC.nMovNro AND  MCD.cCtaCod = MC.cCtaCod"
    'sSql = sSql & " Inner Join Constante CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSql = sSql & " INNER JOIN (Select nConsCod,    nConsValor,  cConsDescripcion, (CONVERT(varchar(10),nConsValor)  + ' ' + cConsDescripcion ) as cAbrev "
    sSql = sSql & "    From Constante ) as CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSql = sSql & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(MC.cCtaCod,4,2)"
    sSql = sSql & " Inner Join OpeTpo O ON O.cOpeCod = MC.cOpeCod"
    sSql = sSql & " Inner Join Colocaciones CC ON CC.cCtaCod = MC.cCtaCod"
    'ssql = ssql & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = CC.cLineaCred"
    sSql = sSql & " Inner Join ProductoPersona CL ON CL.cCtaCod = CC.cCtaCod AND CL.nPrdPersRelac=28"
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
    'ssql = ssql & " Inner Join Mov M ON M.nMovNro = MC.nMovNro "
    sSql = sSql & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'  and Len(MC.cCtaCod) = 18 And (MC.cOpeCod >= '" & gCredDesembEfec & "' and MC.cOpeCod <= '" & gCredDesembCtaNuevaDOA & "') And SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And SUBSTRING(MC.cCtaCod,4,2) IN " & sCadAge
    sSql = sSql & "     And SUBSTRING(M.cMovNro,1,8) Between  '" & Format(pdFecIni, "yyyymmdd") & "' AND '" & Format(pdFecFin, "yyyymmdd") & "' " & sCadProd
    sSql = sSql & "  "
    sSql = sSql & " GROUP BY A.cAgeDescripcion, CN.cAbrev, P.cPersNombre, O.cOpeDesc, MC.cCtaCod,MC.nMovNro "
    sSql = sSql & " With ROLLUP)X "
    sSql = sSql & " Where (cUserDes Is Not Null ) OR nCredito='TOTAL'"    'ARCV 23-06-2007
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

    Set oConecta = New DConecta
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
Dim oConecta As DConecta

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
    sSql = sSql & " Inner join Mov M ON M.nMOvNro = MC.nMovNro" '
    sSql = sSql & " Inner join Constante CN2 ON CN2.nConsValor = Convert(int,SUBSTRING(CC.cLineaCred,6,1)) And CN2.nConsCod = " & gColocLineaCredPlazo
    sSql = sSql & " Where SUBSTRING(Prd.cCtaCod,6,3)<>'305'  and Len(Prd.cCtaCod) = 18 And Prd.nPrdEstado = " & gColocEstCancelado & " And SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And SUBSTRING(MC.cCtaCod,4,2) IN " & sCadAge
    sSql = sSql & "     And Prd.dPrdEstado Between  '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "' AND M.nMovFlag=0"
    sSql = sSql & " GROUP BY A.cAgeDescripcion, CN2.cConsDescripcion, CN.cConsDescripcion, P.cPersNombre, Prd.cCtaCod"
    sSql = sSql & " With ROLLUP)X"
    sSql = sSql & " Left Join RelCuentas RC on RC.cCtaCod=X.nCredito"
    sSql = sSql & " Where X.nCredito in (Select cCtaCod From ProductoPersona"
    sSql = sSql & " Where cPersCod in (" & psAnalistas & ")" & " And nPrdPersRelac=28)"
    

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

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosCanceladosConsaldo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaResumenSaldosCarteraPorAnalista(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal P1I As Integer, ByVal P1F As Integer, _
    ByVal P2I As Integer, ByVal P2F As Integer, ByVal P3I As Integer, ByVal P3F As Integer, ByVal P4F As Integer, _
    ByVal cMatCondicion As String, ByVal pMatProd As Variant, Optional ByVal pnTipCam As Double) As ADODB.Recordset 'DAOR 20070717, se agregó pnTipCam
    
    'ByVal pMatCondicion As Variant
    
Dim sSql As String
Dim oConecta As DConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sAnalistas As String
Dim oGen As DGeneral

    Set oGen = New DGeneral
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

'    sCadCondicion = "("
'    For i = 0 To UBound(pMatCondicion) - 1
'        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
'    Next i
'    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"
    
    
'    If Len(Trim(cMatCondicion)) > 0 Then
'          lsSQL = lsSQL & " AND (dbo.ColocacCred.nColocCondicion IN (" & cCondicion & ")) "
'    End If
    
    'Comentado por DAOR 20070717
'    ssql = "Select  "
'    ssql = ssql & " CASE WHEN GROUPING(A.cAgeDescripcion)=1 THEN 'TOTAL'"
'    ssql = ssql & "   ELSE ISNULL(A.cAgeDescripcion,'AGENCIA DESCONOCIDA')"
'    ssql = ssql & " END  as cAgencia,"
'    ssql = ssql & " CASE WHEN GROUPING(Q.cUser)=1 THEN 'TOTAL'"
'    ssql = ssql & "   ELSE ISNULL(Q.cUser,'AGENCIA DESCONOCIDA')"
'    ssql = ssql & " END as cUser,"
'    ssql = ssql & " SUM(Q.nCar1Nro) as nCar1Nro, SUM(Q.nCar1Saldo) as nCar1Saldo,"
'    ssql = ssql & " SUM(Q.nCar2Nro) as nCar2Nro, SUM(Q.nCar2Saldo) as nCar2Saldo,"
'    ssql = ssql & " SUM(Q.nCar3Nro) as nCar3Nro, SUM(Q.nCar3Saldo) as nCar3Saldo,"
'    ssql = ssql & " SUM(Q.nCar4Nro) as nCar4Nro, SUM(Q.nCar4Saldo) as nCar4Saldo,"
'    ssql = ssql & " SUM(Q.nCar5Nro) as nCar5Nro, SUM(Q.nCar5Saldo) as nCar5Saldo"
'    ssql = ssql & " From ("
'    ssql = ssql & " select cAgencia, T.cUser + ' ' +  T.cPersNombre as cUser, SUM(T.nCar1Nro) as nCar1Nro, SUM(T.nCar1Saldo) as nCar1Saldo,"
'    ssql = ssql & "    SUM(T.nCar2Nro) as nCar2Nro, SUM(T.nCar2Saldo) as nCar2Saldo,"
'    ssql = ssql & "    SUM(T.nCar3Nro) as nCar3Nro, SUM(T.nCar3Saldo) as nCar3Saldo,"
'    ssql = ssql & "    SUM(T.nCar4Nro) as nCar4Nro, SUM(T.nCar4Saldo) as nCar4Saldo,"
'    ssql = ssql & "    SUM(T.nCar5Nro) as nCar5Nro, SUM(T.nCar5Saldo) as nCar5Saldo"
'    ssql = ssql & " From ("
'    ssql = ssql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, COUNT(*) as nCar1Nro, SUM(Prd.nSaldo) as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo "
'    ssql = ssql & "    From RRHH RH "
'    ssql = ssql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
'    ssql = ssql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
'    ssql = ssql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
'    ssql = ssql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
'    ssql = ssql & "    Where  CC.nDiasAtraso >=" & P1I & " AND CC.nDiasAtraso <=" & P1F
'    ssql = ssql & "        And CC.nColocCondicion in " & sCadCondicion
'    ssql = ssql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
'    ssql = ssql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
'    ssql = ssql & "    Union"
'    ssql = ssql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, COUNT(*) as nCar2Nro, SUM(Prd.nSaldo) as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo "
'    ssql = ssql & "    From RRHH RH "
'    ssql = ssql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod "
'    ssql = ssql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
'    ssql = ssql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
'    ssql = ssql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
'    ssql = ssql & "     WHERE  CC.nDiasAtraso >=" & P2I & " AND CC.nDiasAtraso <=" & P2F
'    ssql = ssql & "        And CC.nColocCondicion in " & sCadCondicion
'    ssql = ssql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
'    ssql = ssql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
'    ssql = ssql & "    Union"
'    ssql = ssql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, COUNT(*) as nCar3Nro, SUM(Prd.nSaldo) as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo "
'    ssql = ssql & "    From RRHH RH "
'    ssql = ssql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
'    ssql = ssql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
'    ssql = ssql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
'    ssql = ssql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
'    ssql = ssql & "     WHERE CC.nDiasAtraso >=" & P3I & " AND CC.nDiasAtraso <=" & P3F
'    ssql = ssql & "        And CC.nColocCondicion in " & sCadCondicion
'    ssql = ssql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
'    ssql = ssql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
'    ssql = ssql & "    Union"
'    ssql = ssql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, COUNT(*) as nCar4Nro, SUM(Prd.nSaldo) as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo "
'    ssql = ssql & "    From RRHH RH"
'    ssql = ssql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
'    ssql = ssql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
'    ssql = ssql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
'    ssql = ssql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
'    ssql = ssql & "     WHERE CC.nDiasAtraso > " & P4F
'    ssql = ssql & "        And CC.nColocCondicion in " & sCadCondicion
'    ssql = ssql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
'    ssql = ssql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
'    ssql = ssql & "    Union"
'    ssql = ssql & "    Select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo , COUNT(*) as nCar5Nro, SUM(Prd.nSaldo) as nCar5Saldo "
'    ssql = ssql & "    From RRHH RH"
'    ssql = ssql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
'    ssql = ssql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
'    ssql = ssql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
'    ssql = ssql & "        " & sCadProd
'    ssql = ssql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
'    ssql = ssql & "     WHERE CC.nColocCondicion in " & sCadCondicion
'    ssql = ssql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "'"
'    ssql = ssql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
'    ssql = ssql & " ) as T"
'    ssql = ssql & " Group by  T.cUser + ' ' +  T.cPersNombre, cAgencia"
'    ssql = ssql & " ) As Q Inner Join Agencias A ON A.cAgeCod = Q.cAgencia"
'    ssql = ssql & " WHERE Q.cAgencia IN " & sCadAge
'    ssql = ssql & " GROUP BY A.cAgeDescripcion, Q.cUser"
'    ssql = ssql & " With ROLLUP"
    
    '**DAOR 20070717**********************************************************
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
    sSql = sSql & " SUM(Q.nCar5Nro) as nCar5Nro, SUM(Q.nCar5Saldo) as nCar5Saldo,"
    sSql = sSql & " SUM(Q.nCar6Nro) as nCar6Nro, SUM(Q.nCar6Saldo) as nCar6Saldo"
    sSql = sSql & " From ("
    sSql = sSql & " Select cAgencia, T.cUser + ' ' +  T.cPersNombre as cUser, SUM(T.nCar1Nro) as nCar1Nro, SUM(T.nCar1Saldo) as nCar1Saldo,"
    sSql = sSql & "    SUM(T.nCar2Nro) as nCar2Nro, SUM(T.nCar2Saldo) as nCar2Saldo,"
    sSql = sSql & "    SUM(T.nCar3Nro) as nCar3Nro, SUM(T.nCar3Saldo) as nCar3Saldo,"
    sSql = sSql & "    SUM(T.nCar4Nro) as nCar4Nro, SUM(T.nCar4Saldo) as nCar4Saldo,"
    sSql = sSql & "    SUM(T.nCar5Nro) as nCar5Nro, SUM(T.nCar5Saldo) as nCar5Saldo,"
    sSql = sSql & "    SUM(T.nCar6Nro) as nCar6Nro, SUM(T.nCar6Saldo) as nCar6Saldo"
    sSql = sSql & " From ("
    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, COUNT(*) as nCar1Nro, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldo else Prd.nSaldo*" & Str(pnTipCam) & " end) as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo, 0 as nCar6Nro, 0 as nCar6Saldo "
    sSql = sSql & "    From RRHH RH "
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSql = sSql & "    Where  CC.nDiasAtraso >=" & P1I & " AND CC.nDiasAtraso <=" & P1F
    sSql = sSql & "        And CC.nColocCondicion IN (" & cMatCondicion & ") "
    sSql = sSql & "        " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Union "
    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, COUNT(*) as nCar2Nro, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldo else Prd.nSaldo*" & Str(pnTipCam) & " end) as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo, 0 as nCar6Nro, 0 as nCar6Saldo "
    sSql = sSql & "    From RRHH RH "
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod "
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSql = sSql & "     WHERE  CC.nDiasAtraso >=" & P2I & " AND CC.nDiasAtraso <=" & P2F
    sSql = sSql & "        And CC.nColocCondicion IN (" & cMatCondicion & ") "
    sSql = sSql & "        " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Union "
    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, COUNT(*) as nCar3Nro, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldo else Prd.nSaldo*" & Str(pnTipCam) & " end) as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo, 0 as nCar6Nro, 0 as nCar6Saldo "
    sSql = sSql & "    From RRHH RH "
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSql = sSql & "     WHERE CC.nDiasAtraso >=" & P3I & " AND CC.nDiasAtraso <=" & P3F
    sSql = sSql & "        And CC.nColocCondicion IN (" & cMatCondicion & ") "
    sSql = sSql & "        " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Union "
    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, COUNT(*) as nCar4Nro, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldo else Prd.nSaldo*" & Str(pnTipCam) & " end) as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo, 0 as nCar6Nro, 0 as nCar6Saldo "
    sSql = sSql & "    From RRHH RH"
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSql = sSql & "     WHERE CC.nDiasAtraso > " & P4F
    sSql = sSql & "        And CC.nColocCondicion IN (" & cMatCondicion & ") "
    sSql = sSql & "        " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Union "
    sSql = sSql & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo, count(*) as nCar5Nro, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldo else Prd.nSaldo*" & Str(pnTipCam) & " end) as nCar5Saldo, 0 as nCar6Nro, 0 as nCar6Saldo "
    sSql = sSql & "    From RRHH RH"
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado in (" & gColocEstRecVigJud & ",2205) "
    sSql = sSql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSql = sSql & "     WHERE CC.nColocCondicion IN (" & cMatCondicion & ") "
    sSql = sSql & "        " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Union "
    sSql = sSql & "    Select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo , 0 as nCar5Nro, 0 as nCar5Saldo, count(*) as nCar6Nro, SUM(case substring(Prd.cCtaCod,9,1) when '1' then Prd.nSaldo else Prd.nSaldo*" & Str(pnTipCam) & " end) as nCar6Saldo "
    sSql = sSql & "    From RRHH RH"
    sSql = sSql & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSql = sSql & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And ((Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor & ") Or Prd.nPrdEstado in (" & gColocEstRecVigJud & ",2205))"
    sSql = sSql & "        Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSql = sSql & "     WHERE CC.nColocCondicion IN (" & cMatCondicion & ") "
    sSql = sSql & "        " & sCadProd
    sSql = sSql & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSql = sSql & " ) as T"
    sSql = sSql & " Group by  T.cUser + ' ' +  T.cPersNombre, cAgencia"
    sSql = sSql & " ) As Q Inner Join Agencias A ON A.cAgeCod = Q.cAgencia"
    sSql = sSql & " WHERE Q.cAgencia IN " & sCadAge
    sSql = sSql & " GROUP BY A.cAgeDescripcion, Q.cUser"
    sSql = sSql & " With ROLLUP"
    '***************************************************************************
    
    'ssql = ssql & "     WHERE CC.nColocCondicion in " & sCadCondicion
    
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
    sSql = sSql & " sTitular = ISNULL((select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = T.cCtaCod),''),"
    sSql = sSql & " SUM(T.nMonto) as nMontoApr,"
    sSql = sSql & " SUM(T.nMontoCol) as nMontoDesemb,"
    sSql = sSql & " SUM(T.nSaldo) as nSaldoCap,"
    'sSql = sSql & " nCuotasApr = (Select nCuotas From ColocacEstado Where cCtaCod = T.cCtaCod And nPrdEstado = 2002),"
    sSql = sSql & " nCuotasApr = (Select COUNT(*) From ColocCalendario CC Where CC.cCtaCod = T.cCtaCod AND CC.nNroCalen = (Select nNroCalen from ColocacCred Where cCtaCod = T.cCtaCod) AND CC.nColocCalendApl = 1),"
    sSql = sSql & " nCuotaPend = ISNULL((Select nNroProxCuota From ColocacCred Where cCtaCod = T.cCtaCod),0),"
    sSql = sSql & " SUM(T.nDiasAtraso) as nDiasAtraso,"
    sSql = sSql & " sAnalista = (Select Top 1 ISNULL(cUser,'') From RRHH Where cPersCod = (select cPersCod From ProductoPersona PP2 Where PP2.cCtaCod = T.cCtaCod And PP2.nPrdPersRelac = 28)),"
    sSql = sSql & " nSaldoCapTotal = (Select SUM(nSaldo) From Producto Where nPrdEstado IN (2020,2021,2022,2030,2031,2032,2201,2205) And SUBSTRING(cCtaCod,9,1) = '" & pnMoneda & "')"
    sSql = sSql & " From"
    sSql = sSql & " ( Select 'RANGO1' as cRango, Prd.cCtaCod, CE.nMonto, C.nMontoCol, Prd.nSaldo, CC.nDiasAtraso"
    sSql = sSql & " From  Producto Prd"
    sSql = sSql & "     Inner join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & "     Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "     Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Where Prd.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2201,2205) "
    sSql = sSql & "        AND CC.nDiasAtraso >=" & P1I & " AND CC.nDiasAtraso <=" & P1F
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSql = sSql & " Union"
    sSql = sSql & " Select 'RANGO2' as cRango, Prd.cCtaCod, CE.nMonto, C.nMontoCol, Prd.nSaldo, CC.nDiasAtraso"
    sSql = sSql & " From  Producto Prd"
    sSql = sSql & "     Inner join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & "     Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "     Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Where Prd.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2201,2205) "
    sSql = sSql & "        AND CC.nDiasAtraso >=" & P2I & " AND CC.nDiasAtraso <=" & P2F
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSql = sSql & " Union"
    sSql = sSql & " Select 'RANGO3' as cRango, Prd.cCtaCod, CE.nMonto, C.nMontoCol, Prd.nSaldo, CC.nDiasAtraso"
    sSql = sSql & " From  Producto Prd"
    sSql = sSql & "     Inner join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & "     Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "     Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Where Prd.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2201,2205) "
    sSql = sSql & "        AND CC.nDiasAtraso >=" & P3I & " AND CC.nDiasAtraso <=" & P3F
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSql = sSql & " Union"
    sSql = sSql & " Select 'RANGO4' as cRango, Prd.cCtaCod, CE.nMonto, C.nMontoCol, Prd.nSaldo, CC.nDiasAtraso"
    sSql = sSql & " From  Producto Prd"
    sSql = sSql & "     Inner join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & "     Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "     Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Where Prd.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2201,2205) "
    sSql = sSql & "        AND CC.nDiasAtraso >" & P4F
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    sSql = sSql & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSql = sSql & " ) as T Inner join Agencias A ON A.cAgeCod = SUBSTRING(T.cCtaCod,4,2)"
    sSql = sSql & " WHERE A.cAgeCod IN " & sCadAge
    sSql = sSql & " Group by A.cAgeDescripcion, T.cRango, T.cCtaCod"
    sSql = sSql & " With ROLLUP"
    sSql = sSql & " ORDER BY sAgencia,cRango,cCtaCod " 'ARCV 06-07-2006
    
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
    'ssql = ssql & "        And CC.nColocCondicion in " & sCadCondicion
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
Dim i As Integer
Dim sCadAge As String
Dim sCadTipoCredito As String
Dim sCadCondicion As String
   
   ''''' BRGO 07/06/2010 BASILEA II
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    
    'sCadProd = " And SUBSTRING(PRD.cCtaCod,6,3) in ('"
    sCadTipoCredito = " And C.cTpoCredCod in ('"
    For i = 0 To UBound(pMatProd) - 1
        sCadTipoCredito = sCadTipoCredito & pMatProd(i) & "','"
    Next i
    sCadTipoCredito = Mid(sCadTipoCredito, 1, Len(sCadTipoCredito) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadTipoCredito = ""
    End If
    
    sCadCondicion = "("
    For i = 0 To UBound(pMatCondicion) - 1
        sCadCondicion = sCadCondicion & pMatCondicion(i) & ","
    Next i
    sCadCondicion = Mid(sCadCondicion, 1, Len(sCadCondicion) - 1) & ")"
    sSql = " SELECT * FROM ("
    sSql = sSql & " SELECT "
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
    sSql = sSql & " Inner Join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Inner Join Constante CN ON Convert(Int,C.cTpoCredCod) = CN.nConsValor And CN.nConsCod = 3034" '& gProducto
    
    'sSql = sSql & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & " Inner join Agencias A ON A.cAgeCod = C.cAgeCodAct"
    
   
    sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod And CC.nDiasAtraso >= " & pnDiasAtrIni & " And CC.nDiasAtraso <=" & pnDiasAtrFin
    sSql = sSql & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = C.cLineaCred"
    sSql = sSql & " Inner join Constante CN2 ON CN2.nConsValor = Convert(int,SUBSTRING(C.cLineaCred,6,1)) And CN2.nConsCod = " & gColocLineaCredPlazo
    sSql = sSql & " Inner Join Persona P ON CL.cPersCod = P.cPersCod"
    sSql = sSql & " Inner Join ColocacEstado CE ON CE.cCtaCod = Prd.CCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSql = sSql & " WHERE (Prd.nPrdEstado IN (" & gColocEstRefNorm & ", " & gColocEstRefVenc & ", " & gColocEstRefMor & ", " & gColocEstVigNorm & ", " & gColocEstVigVenc & ", " & gColocEstVigMor & ")) "
    sSql = sSql & " And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadTipoCredito & " And A.cAgeCod IN " & sCadAge
    sSql = sSql & "        And CC.nColocCondicion in " & sCadCondicion
    
    If pnConCuotaLibre = True Then
         sSql = sSql & " And CC.nColocCalendCod=" & gColocCalendCodCL & " "
    End If
    
    sSql = sSql & " GROUP BY A.cAgeDescripcion, CN.cConsDescripcion, CN2.cConsDescripcion, P.cPersNombre, Prd.cCtaCod"
    sSql = sSql & " WITH ROLLUP"
    sSql = sSql & " ) X ORDER BY sAgencia,sProducto,sPlazo,sFuenteFinan,cCtaCod"
    ''''' FIN BRGO
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCredVig_CredVigConCuotLib = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosXInstitucion(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, _
    ByVal pdFecIni As Date, ByVal pnTipoOrden As Integer, pnIncluyeMora As Integer, ByVal MatInstitucion As Variant) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
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
    sSql = sSql & "                 Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod" 'BRGO 07/06/2010
    sSql = sSql & "                 Inner join Agencias A ON A.cAgeCod = C.cAgeCodAct"
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

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosXInstitucion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosXInstitucionDetAnt(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, _
    ByVal pdFecIni As Date, ByVal pnTipoOrden As Integer, pnIncluyeMora As Integer, ByVal MatInstitucion As Variant, ByVal pdFecSis As Date, _
    Optional ByVal pnDia As Integer = 0) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
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

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosXInstitucionDetAnt = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosXInstitucionDet(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, _
    ByVal pdFecIni As Date, ByVal pnTipoOrden As Integer, pnIncluyeMora As Integer, ByVal MatInstitucion As Variant, ByVal pdFecSis As Date, _
    Optional ByVal pnDia As Integer = 0) As ADODB.Recordset
    
Dim sSql As String
Dim sSQL1 As String
Dim oConecta As DConecta
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
    'LSDO
    ''INI
    sSql = sSql & " INTO #CREDINSTITUC "
    ''FIN
    sSql = sSql & " From  ("
    sSql = sSql & "         SELECT  sAgencia, cCtaCod, cInstitucion, nCasos, isnull(sTitular,'ZZZZ') as sTitular, cCodModular,"
    sSql = sSql & "                 dFecDesemb, nDiasAtraso, nMontoApr, nSaldoCap, nDiaAtraCuota, sCuotaPend , nSaldoCuota, nCuota, nCuotaPactada, nITFSalCuota,"
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
    'LSDO
    'INI
    sSql = sSql & "                     SUM(CDet.nCuotaPactada ) AS nCuotaPactada ,"
    'FIN
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
    sSql = sSql & "                     Inner join Colocaciones C ON C.cCtaCod=Prd.cCtaCod"
    sSql = sSql & "                     Inner join Agencias A ON A.cAgeCod = C.cAgeCodAct"
    sSql = sSql & "                     Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "                     Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSql = sSql & "                     Inner Join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSql = sSql & "                     Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & "                     Inner Join (    Select  CDet.cCtaCod, CDet.nNroCalen, CDet.nCuota,  SUM( CASE WHEN CCAL.dVenc <= '" & Format(pdFecIni, "mm/dd/yyyy") & "' THEN CDet.nMonto-CDet.nMontoPagado ELSE 0 END) as nMonto, SUM( CASE WHEN CCAL.dVenc <= '" & Format(pdFecIni, "mm/dd/yyyy") & "' THEN CDet.nMonto ELSE 0 END) as nMonCuota, "
    'LSDO
    'INI
    sSql = sSql & "                                             SUM( CASE CDet.nPrdConceptoCod WHEN 1000 THEN CDet.nMonto WHEN 1100 THEN CDet.nMonto ELSE 0 END) as nCuotaPactada,"
    'FIN
    sSql = sSql & "                                             nCapPend= SUM( CASE WHEN CCAL.dVenc <= '" & Format(pdFecIni, "mm/dd/yyyy") & "' THEN  ( CASE WHEN CDet.nPrdConceptoCod=1000 THEN CDet.nMonto-CDet.nMontoPagado ELSE 0 END ) ELSE 0 END) "
    sSql = sSql & "                                     From    ColocCalendario CCAL"
    sSql = sSql & "                                             JOIN ColocCalendDet CDet ON CDet.cCtaCod = CCAL.cCtaCod AND CDet.nNroCalen = CCAL.nNroCalen "
    sSql = sSql & "                                             AND CDet.nColocCalendApl = CCAL.nColocCalendApl AND CDet.nCuota = CCAL.nCuota "
    'LSDO
    'INI
    sSql = sSql & "                                     Where   CDet.nColocCalendApl = 1 AND CCAL.nColocCalendEstado=0 AND  CCAL.dVenc <= '" & Format(pdFecIni, "mm/dd/yyyy") & "' "
    'FIN
    'ssql = ssql & "                                     Where   CDet.nColocCalendApl = 1 AND CCAL.nColocCalendEstado=0 AND  CCAL.dVenc <= '" & Format(pdFecIni + 30, "mm/dd/yyyy") & "' "
    sSql = sSql & "                                     Group By CDet.cCtaCod, CDet.nNroCalen, CDet.nCuota ) as CDet ON CDet.cCtaCod = Prd.cCtaCod And CDet.nNroCalen = CC.nNroCalen "
    sSql = sSql & "               Where Prd.nPrdEstado >= 2020 And Prd.nPrdEstado <= 2032 "
    sSql = sSql & "                     And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' And A.cAgeCod IN " & sCadAge & " AND CV.cPersCod in " & sCadInst
    sSql = sSql & "            Group By A.cAgeDescripcion, P.cPersNombre, Prd.cCtaCod, CDET.nCuota"
    sSql = sSql & "            ) as X ) as Y"
    If pnDia > 0 Then
    sSql = sSql & "              Where DatePart(day,dFecPago)=" & pnDia
    End If
    sSql = sSql & "     order by  cInstitucion ASC, sagencia," & sCadTipoOrden
    
    'LSDO
    ''INI
    sSQL1 = sSQL1 & "  SELECT sAgencia, cCtaCod, cInstitucion, nCasos, sTitular, cCodModular, dFecDesemb, nDiasAtraso, " & _
                    "  nMontoApr, nSaldoCap, SUM(nDiaAtraCuota) AS nDiaAtraCuota, MIN(sCuotaPend) AS sCuotaPend, " & _
                    "  SUM(nSaldoCuota) AS nSaldoCuota, SUM(nCuota) AS nCuota, SUM(nCuotaPactada) /  COUNT(cCtaCod) AS nCuotaPactada ," & _
                    "  MIN(dFecPago) AS dFecPago , SUM(nCapPend) AS nCapPend, nTasMor, SUM(nMoraCalc) AS nMoraCalc, " & _
                    "  SUM(nITFSalCuota) AS nITFSalCuota , SUM(nItfCuota) AS nItfCuota , SUM(nItfCalc) As nItfCalc " & _
                    "  From #CREDINSTITUC " & _
                    "  GROUP BY sAgencia, cCtaCod, cInstitucion, nCasos, sTitular, cCodModular, dFecDesemb, nDiasAtraso, " & _
                    "  nMontoApr , nSaldoCap,  nTasMor " & _
                    "  order by  cInstitucion ASC, sagencia," & sCadTipoOrden
    ''FIN
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    'LSDO
    ''INI
    oConecta.Ejecutar sSql
    ''FIN
    Set RecuperaCreditosXInstitucionDet = oConecta.CargaRecordSet(sSQL1)
    'oconecta.Ejecutar "DROP TABLE #CREDINSTITUC "
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function


'Public Function RecuperaCreditosXInstitucion(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, _
'    ByVal pdFecIni As Date, ByVal pnTipoOrden As Integer, pnIncluyeMora As Integer) As ADODB.Recordset
'
'Dim sSQL As String
'Dim oConecta As DConecta
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
'    Set oConecta = New DConecta
'    oConecta.AbreConexion
'    Set RecuperaCreditosXInstitucion = oConecta.CargaRecordSet(sSQL)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'
'End Function

Public Function RecuperaMoraXInstituciones(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, _
    ByVal pdFecIni As Date, pMatInstitucion As Variant) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
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
    sSql = "Select * from ( "
    sSql = sSql & " SELECT "
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
    sSql = sSql & " Inner join Agencias A ON A.cAgeCod = C.cAgeCodAct" 'BRGO 17/06/2010 BASILEA II
    sSql = sSql & " Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod And CV.cPersCod in " & sCadInstitucion
    sSql = sSql & " Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSql = sSql & " Inner Join (Select Cdet.cCtaCod, Cdet.nNroCalen, SUM(Round(nMonto,2)-nMontoPagado) as nMonto "
    sSql = sSql & "     From ColocCalendDet CDet"
    sSql = sSql & "     Inner Join ColocCalendario CC ON Cdet.cCtaCod =   CC.cCtaCod AND Cdet.nNroCalen =   CC.nNroCalen AND Cdet.nCuota =   CC.nCuota AND Cdet.nColocCalendApl =   CC.nColocCalendApl"
    sSql = sSql & "     Where Cdet.nColocCalendApl = 1  AND CC.dVenc <= '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND CC.nColocCalendEstado=0 "
    sSql = sSql & "     Group By Cdet.cCtaCod, Cdet.nNroCalen"
    sSql = sSql & "        ) as CDet ON CDet.cCtaCod = Prd.cCtaCod And CDet.nNroCalen = CC.nNroCalen "
    sSql = sSql & " Inner Join (Select CC.cCtaCod, CC.nNroCalen, SUM(Round(nMonto,2)-nMontoPagado) as nMonto "
    sSql = sSql & "     From ColocCalendDet CDet Inner Join ColocCalendario CC on CC.cCtaCod=CDet.cCtaCod And CC.nNroCalen = CDet.nNroCalen And CC.nCuota = CDet.nCuota And CC.nColocCalendApl = CDet.nColocCalendApl "
    sSql = sSql & "     Where CC.nColocCalendApl = 1 And nPrdConceptoCod = 1101 And CC.nColocCalendEstado = 0"
    sSql = sSql & "     Group By CC.cCtaCod, CC.nNroCalen "
    sSql = sSql & "        ) as CDet2 ON CDet2.cCtaCod = Prd.cCtaCod And CDet2.nNroCalen = CC.nNroCalen "
    sSql = sSql & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor
    sSql = sSql & " And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' And A.cAgeCod IN " & sCadAge & " AND nDiasAtraso > 0 "
    sSql = sSql & " Group By A.cAgeDescripcion, P.cPersNombre, Prd.cCtaCod"
    sSql = sSql & " WITH ROLLUP"
    sSql = sSql & " ) A"
    sSql = sSql & " ORDER By A.sAgencia, A.cInstitucion, A.cCtaCod "
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
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
    
    sCadProd = " And C.cTpoCredCod in ('" 'BRGO 07/06/2010 BASILEA II
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
    
    'sSql = sSql & "    Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSql = sSql & "    Inner join Agencias A ON A.cAgeCod = C.cAgeCodAct" 'BRGO 07/06/2010 BASILEA II
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
    sSql = sSql & "    Inner join Agencias A ON A.cAgeCod = C.cAgeCodAct"
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
    sSql = sSql & "    Inner join Agencias A ON A.cAgeCod = C.cAgeCodAct"
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
    sSql = sSql & "    Inner join Agencias A ON A.cAgeCod = C.cAgeCodAct"
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
    Dim Rs As New ADODB.Recordset
    
    sSql = "SELECT TOP 1 nValPond FROM TIPOCAMBIO WHERE datediff(d, dateadd(d," & Mid(pdFechaIni, 1, 2) * -1 & ", '" _
        & Format(pdFechaIni, gsFormatoFecha) & "'),  dFecCamb) = 0 ORDER BY dFecCamb "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    
    Set Rs = oConecta.CargaRecordSet(sSql)
    
    If Rs.BOF Then
        RecuperaTCLavDineroM = 0
    Else
        RecuperaTCLavDineroM = Rs!nValPond
    End If
    
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaListaCreditosComite(ByVal pdFecSis As Date, ByVal psCodAgencia As String) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
    
    
    Set oConec = New DConecta
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
Dim oConecta As DConecta
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

Set oConecta = New DConecta
oConecta.AbreConexion
Set RecuperaPersDevCredConv = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function
Public Function RecuperaDevCredConv(ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, ByVal MatInstitucion As Variant, Optional pCastDev As Integer = 0) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
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
        'MADM 20110930
        If pCastDev <> 0 Then
           sSql = sSql & " AND (bPagoCastigo= 1)"
        Else
           sSql = sSql & " AND (bPagoCastigo is NULL or bPagoCastigo= 0)"
        End If
        'END MADM
sSql = sSql + "         ORDER BY C.nMoneda, C.cCodage, C.cPersCodIns, P1.cPersNombre"

Set oConecta = New DConecta
oConecta.AbreConexion
Set RecuperaDevCredConv = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

Public Function ListaProtestosDia(ByVal psCodAge As String, ByVal pdFecha As Date) As Recordset
    Dim sSql As String
    Dim Rs As ADODB.Recordset
    Dim sDia As String
    Dim sMes As String
    Dim sAno As String
    Dim sFecha As String
    Dim oConec As DConecta
    
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
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaProtestosDia = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
End Function


Public Function ObtenerTelefonosTitular(ByVal psCtaCod As String) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
    
    sSql = "Select Pers.cPersTelefono,Pers.cPersTelefono2 "
    sSql = sSql & " From ProductoPersona PP"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Where PP.nPrdPersRelac=20 and PP.cCtaCod='" & psCtaCod & "'"
    
     Set oConec = New DConecta
     oConec.AbreConexion
     Set ObtenerTelefonosTitular = oConec.CargaRecordSet(sSql)
     oConec.CierraConexion
End Function

Public Function ObtenerDireccTrabajoTitular(ByVal psCtaCod As String) As String
    Dim sSql As String
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Dim sDireccion As String
    
     sSql = "Select PFI.cRazSocDirecc"
     sSql = sSql & " From ColocFteIngreso CFT"
     sSql = sSql & " Inner Join PersFteIngreso PFI on PFI.cNumFuente=CFT.cNumFuente"
     sSql = sSql & " Where CFT.cCtaCod='" & psCtaCod & "'"

     Set oConec = New DConecta
     oConec.AbreConexion
     Set Rs = oConec.CargaRecordSet(sSql)
     oConec.CierraConexion
     Set oConec = Nothing
     
     If Not Rs.EOF And Not Rs.BOF Then
        If Not IsNull(Rs!cRazSocDirecc) Then
           sDireccion = Rs!cRazSocDirecc
        End If
     End If
     
     Set Rs = Nothing
     If sDireccion = "" Then
        sSql = "Select P.cPersDireccDomicilio"
        sSql = sSql & " from ProductoPersona PP"
        sSql = sSql & " Inner Join Persona P on PP.cPersCod=P.cPersCod"
        sSql = sSql & " Where PP.nPrdPersRelac=20 and PP.cCtaCod='" & psCtaCod & "'"
        
        Set oConec = New DConecta
        oConec.AbreConexion
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        Set oConec = Nothing
        
        If Not Rs.EOF And Not Rs.BOF Then
            If Not IsNull(Rs!cPersDireccDomicilio) Then
                sDireccion = Rs!cPersDireccDomicilio
            End If
        End If
        Set Rs = Nothing
     End If
     ObtenerDireccTrabajoTitular = sDireccion
End Function

Function ObtenerConvenio(ByVal psCtaCod As String) As String
    Dim oConec As DConecta
    Dim sSql As String
    Dim Rs As ADODB.Recordset
    
    sSql = "Select Pers.cPersNombre"
    sSql = sSql & " From ColocacConvenio CC"
    sSql = sSql & " Inner Join Persona Pers on CC.cPersCod=Pers.cPersCod"
    sSql = sSql & " Where cCtacod='" & psCtaCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set Rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not Rs Is Nothing Then
        If Not Rs.EOF And Not Rs.BOF Then
           ObtenerConvenio = "Convenio Con :" & Rs!cpersnombre
        Else
            ObtenerConvenio = ""
        End If
    Else
        ObtenerConvenio = ""
    End If
    Set Rs = Nothing
End Function

Public Function ObtenerCarteraVigentePorUser(ByVal psUser As String, ByVal psCodAgen As String, _
ByVal psProducto As String, ByVal pbFecha As Boolean, Optional pdFecha As Date, _
Optional ByVal psProductoFiltro As String) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
    Dim nTipoCambio As Double
    Dim odCredDoc As DCredDoc
    Dim Rs As ADODB.Recordset
    
    
    Set odCredDoc = New DCredDoc
        nTipoCambio = odCredDoc.ObtenerTipoCambio(pdFecha)
    Set odCredDoc = Nothing
    
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
  
    
    Set oConec = New DConecta
    Sleep 250
    oConec.AbreConexion
    Set ObtenerCarteraVigentePorUser = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    Exit Function
End Function

Public Function ListaConsolidadoPorAnalista(ByVal pnTipoCambio As Double, ByVal psAgenciaCod As String, ByVal pdFecha As Date, _
ByVal psAnalistas As String, ByVal psProducto As String, ByVal pbFecha As Boolean, Optional ByVal psProductoFiltro As String = "") As Recordset

    Dim sSql As String
    Dim oConec As DConecta
    Dim sFecha As String
    Dim sDia As String
    Dim sMes As String
    Dim sAno As String
    Dim sFechaMov As String
    Dim sdInicioMes As String
    Dim nTipoCambio As Double
    Dim odCredDoc As DCredDoc
    
    sFechaMov = Format(pdFecha, "MM/dd/yyyy")
    sAno = DatePart("yyyy", pdFecha)
    sMes = Right("00", 2 - Len(CStr(DatePart("m", pdFecha)))) & CStr(DatePart("m", pdFecha))
    sDia = Right("00", 2 - Len(CStr(DatePart("d", pdFecha)))) & CStr(DatePart("d", pdFecha))
    
    Set odCredDoc = New DCredDoc
    nTipoCambio = odCredDoc.ObtenerTipoCambio(pdFecha)
    
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
    Set oConec = New DConecta
    oConec.AbreConexion
        Set ListaConsolidadoPorAnalista = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaUser(ByVal psPersCod As String) As String
    Dim sSql As String
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    
    
    Set oConec = New DConecta
    sSql = "Select cUser From RRHH Where cPersCod='" & psPersCod & "'"
    oConec.AbreConexion
    Set Rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not Rs.EOF And Not Rs.BOF Then
        ListaUser = Rs!cUser
    End If
    Set Rs = Nothing
End Function

Public Sub InsertarTmpConsolidadoAnalistas(ByVal rsAnalista As ADODB.Recordset, _
ByVal psCodUser As String, ByVal pdFecha As Date, ByVal psProducto As String, _
ByVal psCodAgen As String, ByVal pbFecha As Boolean)
    Dim oConec As DConecta
    Dim sSql As String
    Dim nNumCredVig As Integer
    Dim nKDesemVig As Double
    Dim nSaldokVig As Double
    Dim nSaldoPromVig As Double
    Dim Rs As ADODB.Recordset
    Dim nTipoCambio As Double
    Dim nCodigo As Integer
    Dim odCredDoc As DCredDoc
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
    
    
    Set oConec = New DConecta
    oConec.AbreConexion
    sSql = "Select Isnull(Max(IdCodigo),0) as IdCodigo From TmpConsolidadoAnalista"
    Set Rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    If Not Rs.EOF And Not Rs.BOF Then
        nCodigo = 1 + Rs!IdCodigo
    End If
    
    Set Rs = Nothing
    oConec.AbreConexion
    If rsAnalista.RecordCount > 0 Then
        'oConec.ConexionActiva.BeginTrans
    End If
    
    If pbFecha = True Then
        Set odCredDoc = New DCredDoc
        nTipoCambio = odCredDoc.ObtenerTipoCambio(pdFecha)
        Set odCredDoc = Nothing
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
               Set Rs = ListaPrendarioConsolidadAnalista(psCodAgen, pdFecha, pbFecha)
               If Not Rs.EOF And Not Rs.BOF Then
                    nNumCredVig = IIf(IsNull(Rs!cantidad), 0, Rs!cantidad)
                    nKDesemVig = IIf(IsNull(Rs!Desembolsado), 0, Rs!Desembolsado)
                    nSaldokVig = IIf(IsNull(Rs!SaldosK), 0, Rs!SaldosK)
                    nSaldoPromVig = IIf(IsNull(Rs!SALDOSKPROM), 0, Rs!SALDOSKPROM)
                End If
                Set Rs = Nothing
            End If
            
    If pbFecha = False Then
            If rsAnalista!cCodAna <> "PREN" Then
                Set Rs = ObtenerCarteraVigentePorUser(rsAnalista!cCodAna, psCodAgen, psProducto, pbFecha, pdFecha)
                ' Set rs = ObtenerCarteraVigentePorUserStored(psCodAgen, rsAnalista!cCodAna, pdFecha)
                If Not Rs.EOF And Not Rs.BOF Then
                    nNumCredVig = IIf(IsNull(Rs!NumCreVig), 0, Rs!NumCreVig)
                    nKDesemVig = IIf(IsNull(Rs!SaldoDesem), 0, Rs!SaldoDesem)
                    nSaldokVig = IIf(IsNull(Rs!SaldoKVig), 0, Rs!SaldoKVig)
                    nSaldoPromVig = IIf(IsNull(Rs!SaldoPromVig), 0, Rs!SaldoPromVig)
                End If
                Set Rs = Nothing
             End If
   Else
              If rsAnalista!cCodAna <> "PREN" Then
                sSql = "Select *"
                sSql = sSql & " From TempSaldoVigConsolAnalista"
                sSql = sSql & " Where IdCodigo=" & nCodigo & " and cUser='" & rsAnalista!cCodAna & "'"
                
                Set Rs = oConec.CargaRecordSet(sSql)
                   If Not Rs.EOF And Not Rs.BOF Then
                      nNumCredVig = IIf(IsNull(Rs!NumCreVig), 0, Rs!NumCreVig)
                      nKDesemVig = IIf(IsNull(Rs!SaldoDesem), 0, Rs!SaldoDesem)
                      nSaldokVig = IIf(IsNull(Rs!SaldoKVig), 0, Rs!SaldoKVig)
                      nSaldoPromVig = IIf(IsNull(Rs!SaldoPromVig), 0, Rs!SaldoPromVig)
                  End If
                  Set Rs = Nothing
             End If
   End If
            If nNumCredVig > 0 Then
                sSql = "Insert Into TmpConsolidadoAnalista Values('" & Format(pdFecha, "MM/dd/yyyy") & "','" & psCodUser & "',"
                sSql = sSql & "'" & rsAnalista!cCodAna & "'," & nNumCredVig & "," & Format(nKDesemVig, "#0.00") & "," & Format(nSaldokVig, "#0.00") & ","
                sSql = sSql & Format(nSaldoPromVig, "#0.00") & "," & rsAnalista!nNumCredNew & "," & Format(rsAnalista!nMonCredNew, "#0.00") & "," & rsAnalista!nNumCredRest & ","
                sSql = sSql & Format(rsAnalista!nMonCredRest, "#0.00") & "," & rsAnalista!NNRO1_15 & "," & Format(rsAnalista!NSALDO1_15, "#0.00") & "," & Format((rsAnalista!NNRO1_15 / nNumCredVig) * 100, "#0.00") & "," & rsAnalista!NNRO16_30 & ","
                sSql = sSql & Format(rsAnalista!NSALDO16_30, "#0.00") & "," & Format((rsAnalista!NNRO16_30 / nNumCredVig) * 100, "#0.00") & "," & rsAnalista!NNROMAY31 & "," & Format(rsAnalista!NSALDOMAY31, "#0.00") & "," & Format((rsAnalista!NNROMAY31 / nNumCredVig) * 100, "#0.00") & ","
                sSql = sSql & rsAnalista!NNROJUD & "," & Format(rsAnalista!nSaldoJud, "#0.00") & "," & Format((rsAnalista!NNROJUD / nNumCredVig) * 100, "#0.00") & "," & rsAnalista!NNROVEN & "," & Format(rsAnalista!NSALDOVEN, "#0.00") & ","
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
    Dim oConec As DConecta
    Dim sSql As String
    Dim nNumCredVig As Integer
    Dim nKDesemVig As Double
    Dim nSaldokVig As Double
    Dim nSaldoPromVig As Double
    Dim Rs As ADODB.Recordset
    Dim nCodigo As Integer
    
'    Set oConec = New DConecta
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
    Set oConec = New DConecta
    oConec.AbreConexion
    Do Until rsProducto.EOF
        If rsProducto!cCodAna = "PREN" Then
            Set Rs = ListaPrendarioConsolidadAnalista(psCodAgen, pdFecha, pbFecha)
            If Not Rs.EOF And Not Rs.BOF Then
                nNumCredVig = IIf(IsNull(Rs!cantidad), 0, Rs!cantidad)
                nKDesemVig = IIf(IsNull(Rs!Desembolsado), 0, Rs!Desembolsado)
                nSaldokVig = IIf(IsNull(Rs!SaldosK), 0, Rs!SaldosK)
            End If
            Set Rs = Nothing
        End If
        If rsProducto!cCodAna <> "PREN" Then
            Set Rs = ObtenerCarteraVigentePorUser(rsProducto!cCodAna, psCodAgen, "", pbFecha, pdFecha, psProductoFiltro)
            If Not Rs.EOF And Not Rs.BOF Then
                nNumCredVig = IIf(IsNull(Rs!NumCreVig), 0, Rs!NumCreVig)
                nKDesemVig = IIf(IsNull(Rs!SaldoDesem), 0, Rs!SaldoDesem)
                nSaldokVig = IIf(IsNull(Rs!SaldoKVig), 0, Rs!SaldoKVig)
                nSaldoPromVig = IIf(IsNull(Rs!SaldoPromVig), 0, Rs!SaldoPromVig)
            End If
            Set Rs = Nothing
       End If
       If nNumCredVig > 0 Then
            sSql = "Insert Into TmpConsolidadoProducto Values('" & Format(pdFecha, "MM/dd/yyyy") & "','" & psCodUser & "',"
            sSql = sSql & "'" & rsProducto!cCodAna & "'," & nNumCredVig & "," & Format(nKDesemVig, "#0.00") & "," & Format(nSaldokVig, "#0.00") & ","
            sSql = sSql & Format(nSaldoPromVig, "#0.00") & "," & rsProducto!nNumCredNew & "," & Format(rsProducto!nMonCredNew, "#0.00") & "," & rsProducto!nNumCredRest & ","
            sSql = sSql & Format(rsProducto!nMonCredRest, "#0.00") & "," & rsProducto!NNRO1_15 & "," & Format(rsProducto!NSALDO1_15, "#0.00") & "," & Format((rsProducto!NNRO1_15 / nNumCredVig) * 100, "#0.00") & "," & rsProducto!NNRO16_30 & ","
            sSql = sSql & Format(rsProducto!NSALDO16_30, "#0.00") & "," & Format((rsProducto!NNRO16_30 / nNumCredVig) * 100, "#0.00") & "," & rsProducto!NNROMAY31 & "," & Format(rsProducto!NSALDOMAY31, "#0.00") & "," & Format((rsProducto!NNROMAY31 / nNumCredVig) * 100, "#0.00") & ","
            sSql = sSql & rsProducto!NNROJUD & "," & Format(rsProducto!nSaldoJud, "#0.00") & "," & Format((rsProducto!NNROJUD / nNumCredVig) * 100, "#0.00") & "," & rsProducto!NNROVEN & "," & Format(rsProducto!NSALDOVEN, "#0.00") & ","
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
    Dim oConec As DConecta
    Dim sSql As String
    Dim Rs As ADODB.Recordset
    
    Set oConec = New DConecta
    oConec.AbreConexion
    sSql = "Select Max(IdCodigo) as nCodigo From TmpConsolidadoAnalista Where FechaProceso='" & Format(pdFecha, "MM/dd/yyyy") & "' AND Usuario='" & psCodUser & "'"
    
    Set Rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not Rs.EOF And Not Rs.BOF Then
        ObtenerIdCodigo = Rs!nCodigo
    End If
    Set Rs = Nothing
End Function

Public Function ObtenerIDCodigoProductoAnalista(ByVal pdFecha As Date, psCodUser As String) As Integer
    Dim oConec As DConecta
    Dim sSql As String
    Dim Rs As ADODB.Recordset
    
    Set oConec = New DConecta
    oConec.AbreConexion
    sSql = "Select Max(IdCodigo) nCodigo From TmpConsolidadoProducto Where FechaProceso='" & Format(pdFecha, "MM/dd/yyyy") & "' And Usuario='" & psCodUser & "'"
    Set Rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not Rs.EOF And Not Rs.BOF Then
        ObtenerIDCodigoProductoAnalista = Rs!nCodigo
    End If
    Set Rs = Nothing
End Function

Public Function ListaConsolidadoRefinac(ByVal psCodUser As String, ByVal psAgencia As String, _
ByVal psProducto As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    Dim Rs As ADODB.Recordset
    
    Set oConec = New DConecta
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
    Dim oConec As DConecta
    Dim sSql As String
    
    Set oConec = New DConecta
    oConec.AbreConexion
    'ARCV 24-06-2007 mejorar Performance de la Consulta
'******************ALPA
'******************17/03/2008
'    ssql = "Select * From ("
''    ssql = ssql & " Select  RH.cUser,P.cCtaCod,"
''    ssql = ssql & " Titular=(Select Pers.cPersNombre"
''    ssql = ssql & " From ProductoPersona PP1"
''    ssql = ssql & " Inner Join Persona Pers on PP1.cPersCod=Pers.cPersCod"
''    ssql = ssql & " Where PP1.cCtaCod=P.cCtaCod and PP1.nPrdPersRelac=20),"
''    ssql = ssql & " Direccion=(Select Pers.cPersDireccDomicilio"
''    ssql = ssql & " From ProductoPersona PP1 "
''    ssql = ssql & " Inner Join Persona Pers on PP1.cPersCod=Pers.cPersCod"
''    ssql = ssql & " Where PP1.cCtaCod=P.cCtaCod and PP1.nPrdPersRelac=20),"
''    ssql = ssql & " KDesembolsado=(Select nMontoCol"
''    ssql = ssql & " From Colocaciones"
''    ssql = ssql & " Where cCtaCod=P.cCtaCod),"
'
'    ssql = ssql & " Select  "
'    '*********ALPA
'    '**ALPA
'    ssql = ssql & " CC.cCtaCod,RH.cUser,"
'    ssql = ssql & " Titular= Pers.cPersNombre,"
'    ssql = ssql & " isnull((Select (isnull(Pers.cPersTelefono,'') + ' '+ isnull(Pers.cPersTelefono2,' ')) From ProductoPersona PP Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod Where PP.nPrdPersRelac=20 and PP.cCtaCod=P.cCtaCod),'') as TelefonoTitular,"
'    ssql = ssql & " isnull(isnull(Pers2.cPersNombre,''),'') as NombreGarante,"
'    ssql = ssql & " isnull(isnull(Pers2.cPersTelefono,'') + ' '+ isnull(Pers2.cPersTelefono2,' '),'')  telefonoGarante,"
'    ssql = ssql & " KDesembolsado= C.nMontoCol,"
'    ssql = ssql & " NroCuotaApr=(Select Count(*) From ColocCalendario CD Where CD.cCtaCod=P.cCtaCod and CD.nColocCalendApl=1 and CD.nNroCalen=CCA.nNroCalen),"
'    ssql = ssql & " CCA.nCuota,"
'    ssql = ssql & " DateDiff(day,CCA.dVenc,'" & Format(pdFecha, "MM/dd/yyyy") & "') as AtrazoCuota,"
'    ssql = ssql & " CONVERT(VarChar(15), CCA.dVenc, 103) As dFecha,"
'    ssql = ssql & " ((select sum("
'    ssql = ssql & " Case nColocCalendEstado"
'    ssql = ssql & " when 0 then (case"
'    ssql = ssql & " when datediff(day,cal.dvenc,'" & Format(pdFecha, "MM/dd/yyyy") & "') < 0 then 0"
'    ssql = ssql & " else datediff(day,cal.dvenc,'" & Format(pdFecha, "MM/dd/yyyy") & "')"
'    ssql = ssql & " End"
'    ssql = ssql & "     )"
'    ssql = ssql & " else (case"
'    ssql = ssql & " when datediff(day,cal.dvenc,cal.dpago)<0 then 0"
'    ssql = ssql & " else datediff(day,cal.dvenc,cal.dpago)"
'    ssql = ssql & " End"
'    ssql = ssql & " )"
'    ssql = ssql & " End"
'    ssql = ssql & " )"
'    ssql = ssql & " DifFecha"
'    ssql = ssql & " from ColocCalendario cal inner join colocacCred colcad on (cal.cCtaCod=colcad.cCtaCod) and"
'    ssql = ssql & " (cal.nNroCalen=colcad.nNroCalen)"
'    ssql = ssql & " where cal.cCtaCod=P.cCtaCod and cal.nColocCalendApl=1 and dvenc<='" & Format(pdFecha, "MM/dd/yyyy") & "'"
'    ssql = ssql & " )/"
'    'Inicio
'     ssql = ssql & " (select convert(decimal(18,3),count(cal.ncuota))"
'    ssql = ssql & " DifFecha"
'    ssql = ssql & " from ColocCalendario cal inner join colocacCred colcad on (cal.cCtaCod=colcad.cCtaCod) and"
'    ssql = ssql & " (cal.nNroCalen=colcad.nNroCalen)"
'    ssql = ssql & " where cal.cCtaCod=P.cCtaCod and cal.nColocCalendApl=1 and dvenc<='" & Format(pdFecha, "MM/dd/yyyy") & "'"
'    ssql = ssql & " )) as PromDAtra,"
'    'Fin
'     ssql = ssql & " (select max("
'    ssql = ssql & " Case nColocCalendEstado"
'    ssql = ssql & " when 0 then (case"
'    ssql = ssql & " when datediff(day,cal.dvenc,'" & Format(pdFecha, "MM/dd/yyyy") & "') < 0 then 0"
'    ssql = ssql & " else datediff(day,cal.dvenc,'" & Format(pdFecha, "MM/dd/yyyy") & "')"
'    ssql = ssql & " End "
'    ssql = ssql & " )"
'    ssql = ssql & " else (case"
'    ssql = ssql & " when datediff(day,cal.dvenc,cal.dpago)<0 then 0"
'    ssql = ssql & " else datediff(day,cal.dvenc,cal.dpago)"
'    ssql = ssql & " End"
'    ssql = ssql & " )"
'    ssql = ssql & " End"
'    ssql = ssql & " )"
'    ssql = ssql & " DifFecha"
'    ssql = ssql & " from ColocCalendario cal inner join colocacCred colcad on (cal.cCtaCod=colcad.cCtaCod) and"
'    ssql = ssql & " (cal.nNroCalen=colcad.nNroCalen)"
'    ssql = ssql & " where cal.cCtaCod=P.cCtaCod and cal.nColocCalendApl=1 and dvenc<='" & Format(pdFecha, "MM/dd/yyyy") & "'"
'    ssql = ssql & " ) as MaxDiaAtra,"
'     ssql = ssql & " (select min("
'    ssql = ssql & " Case nColocCalendEstado"
'    ssql = ssql & " when 0 then (case"
'    ssql = ssql & " when datediff(day,cal.dvenc,'" & Format(pdFecha, "MM/dd/yyyy") & "') < 0 then 0"
'    ssql = ssql & " else datediff(day,cal.dvenc,'" & Format(pdFecha, "MM/dd/yyyy") & "')"
'    ssql = ssql & " End "
'    ssql = ssql & " )"
'    ssql = ssql & " else (case"
'    ssql = ssql & " when datediff(day,cal.dvenc,cal.dpago)<0 then 0"
'    ssql = ssql & " else datediff(day,cal.dvenc,cal.dpago)"
'    ssql = ssql & " End"
'    ssql = ssql & " )"
'    ssql = ssql & " End"
'    ssql = ssql & " )"
'    ssql = ssql & " DifFecha"
'    ssql = ssql & " from ColocCalendario cal inner join colocacCred colcad on (cal.cCtaCod=colcad.cCtaCod) and"
'    ssql = ssql & " (cal.nNroCalen=colcad.nNroCalen)"
'    ssql = ssql & " where cal.cCtaCod=P.cCtaCod and cal.nColocCalendApl=1 and dvenc<='" & Format(pdFecha, "MM/dd/yyyy") & "'"
'    ssql = ssql & " and datediff(day,cal.dvenc,cal.dpago)>0) as MinDiaAtra,"
'    ssql = ssql & " (Select isnull((Select Sum(nMonto-nMontoPagado) From ColocCalendDet NOLOCK Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen and nColocCalendApl=CC.nColocCalendApl and nCuota=CC.nCuota and nPrdConceptoCod in (1000)),0) From ColocCalendario CC Where CC.cCtaCod=P.cCtaCod and CC.nColocCalendApl=1  and CC.nColocCalendEstado=0 and CC.dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "' and nNroCalen=CCA.nNroCalen and nCuota=CCA.nCuota) as Capital,"
'    ssql = ssql & " (Select isnull((Select Sum(nMonto-nMontoPagado) From ColocCalendDet NOLOCK Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen and nColocCalendApl=CC.nColocCalendApl and nCuota=CC.nCuota   and nPrdConceptoCod in (1100)),0) From ColocCalendario CC Where CC.cCtaCod=P.cCtaCod and CC.nColocCalendApl=1  and CC.nColocCalendEstado=0 and CC.dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "' and nNroCalen=CCA.nNroCalen and nCuota=CCA.nCuota) as Interes,"
'    ssql = ssql & " (Select Mora=isnull((Select Sum(nMonto-nMontoPagado)          From ColocCalendDet NOLOCK Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen And nColocCalendApl = CC.nColocCalendApl and nCuota=CC.nCuota          and nPrdConceptoCod in (1101)),0) From ColocCalendario CC Where CC.cCtaCod=P.cCtaCod and CC.nColocCalendApl=1  and CC.nColocCalendEstado=0 and CC.dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "' and nNroCalen=CCA.nNroCalen and nCuota=CCA.nCuota) as Mora, "
'    ssql = ssql & " (Select Gastos=isnull((Select Sum(nMonto-nMontoPagado)          From ColocCalendDet NOLOCK Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen and nColocCalendApl=CC.nColocCalendApl and nCuota=CC.nCuota          and nPrdConceptoCod like '12%'),0) From ColocCalendario CC Where CC.cCtaCod=P.cCtaCod and CC.nColocCalendApl=1  and CC.nColocCalendEstado=0 and CC.dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "' and nNroCalen=CCA.nNroCalen and nCuota=CCA.nCuota) as Gastos, "
'    ssql = ssql & " MontoCobrar=(Select Sum(CD.nMonto-CD.nMontoPagado) From ColocCalendario CC Inner Join ColocCalendDet CD on CC.cCtaCod=CD.cCtaCod And CC.nNroCalen=CD.nNroCalen and CC.nColocCalendApl=CD.nColocCalendApl and CC.nCuota = CD.nCuota Where CC.cCtaCod=P.cCtaCod and CD.nColocCalendApl=1 and CC.nNroCalen=CCA.nNroCalen and  CC.dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "' and CC.nColocCalendEstado=0),"
'    ssql = ssql & " (Select SaldoCuota=isnull((Select Sum(nMonto-nMontoPagado) "
'    ssql = ssql & " From ColocCalendDet NOLOCK where"
'    ssql = ssql & " cCtaCod=CCC.cCtaCod and nNroCalen=CCC.nNroCalen and nColocCalendApl=CCC.nColocCalendApl and nCuota=CCC.nCuota),0)"
'    ssql = ssql & " From ColocCalendario CCC Where CCC.cCtaCod=P.cCtaCod and CCC.nColocCalendApl=1"
'    ssql = ssql & "     and CCC.nColocCalendEstado=0 and CCC.dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "' and nNroCalen=CCA.nNroCalen and nCuota=CCA.nCuota) as SaldoCuota,"
'    ssql = ssql & " (select nTasaInteres from Producto where cCtaCod=P.cCtaCod) as TasaMor,"
'    ssql = ssql & " (Select SaldoGastoMora=isnull((Select Sum(nMonto-nMontoPagado) From ColocCalendDet NOLOCK Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen And nColocCalendApl = CC.nColocCalendApl and nCuota=CC.nCuota and nPrdConceptoCod in (1101) and nPrdConceptoCod like '12%'),0) From ColocCalendario CC Where CC.cCtaCod=P.cCtaCod and CC.nColocCalendApl=1  and CC.nColocCalendEstado=0 and CC.dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "' and nNroCalen=CCA.nNroCalen and nCuota=CCA.nCuota) as SaldoGastoMora,"
'    ssql = ssql & " (case (Select PFI.cRazSocDirecc From ColocFteIngreso CFT Inner Join PersFteIngreso PFI on PFI.cNumFuente=CFT.cNumFuente Where CFT.cCtaCod=P.cCtaCod) when '' then Pers.cPersDireccDomicilio else (Select PFI.cRazSocDirecc From ColocFteIngreso CFT Inner Join PersFteIngreso PFI on PFI.cNumFuente=CFT.cNumFuente Where CFT.cCtaCod=P.cCtaCod) end) as DireccionTrabajo,"
'    ssql = ssql & " isnull(Pers2.cPersDireccDomicilio,'')  as DireccionGarante,"
'    'ALPA
'    '*********
'    ssql = ssql & " C.dVigencia, "
'    ssql = ssql & " Direccion= Pers.cPersDireccDomicilio,"
'
'    ssql = ssql & " P.nSaldo as SaldoK, "
'
'    ssql = ssql & " CuotasMora=(Select Count(*)"
'    ssql = ssql & " From ColocCalendario CC"
'    ssql = ssql & " Where CC.cCtaCod=P.cCtaCod and CC.nColocCalendApl=1 and CC.nNroCalen=(Select Max(nNroCalen) From ColocCalendDet"
'    ssql = ssql & " Where cCtaCod=P.cCtaCod) and"
'    ssql = ssql & " CC.nColocCalendEstado=0 and dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "')"
''
'    ssql = ssql & " ,cRFA=Case When CC.cRFA Is Null or CC.cRFA='NOR' Then 'NOR' Else CC.cRFA End "
'    ssql = ssql & " From RRHH RH"
'    ssql = ssql & " Inner Join ProductoPersona PP on RH.cPersCod=PP.cPersCod and PP.nPrdPersRelac=28"
'    ssql = ssql & " Inner Join Producto P on P.cCtaCod=PP.cCtaCod"
'    ssql = ssql & " INNER JOIN Colocaciones C ON C.cCtaCod= P.cCtaCod "
'    ssql = ssql & " Inner Join ColocacCred CC on CC.cCtaCod=P.cCtaCod"
'    ssql = ssql & " INNER JOIN colocCalendario CCA on CC.cCtaCod=CCA.cCtaCod"
'    ssql = ssql & " Inner Join ProductoPersona PP1 on PP1.cCtacod=P.cCtaCod and PP1.nPrdPersRelac=20"
'    ssql = ssql & " Inner Join Persona Pers on Pers.cPersCod=PP1.cPersCod"
'    ssql = ssql & " left join ProductoPersona PP2 on P.cCtacod=PP2.cCtacod and PP2.nPrdPersRelac=25"
'    ssql = ssql & " left join Persona Pers2 on Pers2.cPersCod=PP2.cPersCod"
'
'    If Len(psCodUbiGeo) <> 0 Then
'        If Mid(Trim(psCodUbiGeo), 1, 1) = "3" Then 'Todos los del distrito
'            ssql = ssql & " Inner Join UbicacionGeografica UG on Substring(UG.cUbiGeoCod,2,6)=Substring(Pers.cPersDireccUbiGeo,2,6) and Substring(UG.cUbiGeoCod,1,1)='3' and Substring(UG.cUbiGeoCod,8,5)='00000'"
'        Else
'            ssql = ssql & " Inner Join UbicacionGeografica UG on UG.cUbiGeoCod=Pers.cPersDireccUbiGeo"
'        End If
'    End If
'
'    ssql = ssql & " Where CCA.nColocCalendApl=1  and CCA.nColocCalendEstado=0 and "
'    ssql = ssql & " CCA.dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "' and CCA.nNroCalen=(Select Max(nNroCalen)     From ColocCalendario Where cCtaCod=P.cCtaCod) and"
'    ssql = ssql & " P.nPrdEstado in (2020,2021,2022,2030,2031,2032) "
'    ssql = ssql & " and (Select Count(*)"
'    ssql = ssql & " From ColocCalendario CC"
'    ssql = ssql & " Where CC.cCtaCod=P.cCtaCod and CC.nColocCalendApl=1 and CC.nNroCalen=(Select Max(nNroCalen) From ColocCalendDet"
'    ssql = ssql & " Where cCtaCod=P.cCtaCod) and"
'    ssql = ssql & " CC.nColocCalendEstado=0 and dVenc<'" & Format(pdFecha, "MM/dd/yyyy") & "')>0"
'    ssql = ssql & " and RH.cPersCod in (" & psUser & ") and Substring(P.cCtaCod,6,3) in (" & psProductos & ")"
'    'If Len(psCodUbiGeo) <> 0 Then
'    '    If Mid(Trim(psCodUbiGeo), 1, 1) <> "3" Then 'Todos los del distrito
'          ' por zonas
'     If Len(psCodUbiGeo) <> 0 Then
'            ssql = ssql & " and UG.cUbiGeoCod='" & Trim(psCodUbiGeo) & "'"
'     End If
'    '    End If
'    'End If
'    ssql = ssql & " and Substring(P.cCtaCod,4,2) in (" & psAgencias & ")"
'    ssql = ssql & " and CC.nColocCondicion in (" & psCondicion & ")"
'    ssql = ssql & " and Substring(P.cCtaCod,9,1) in (" & psMoneda & ")"
'    ssql = ssql & " and P.nSaldo>=" & pnMontoIni & " and P.nSaldo<=" & pnMontoFin
'    ssql = ssql & " and CC.nDiasAtraso between " & pnIniDiaMora & " and " & pnFinDiaMora & ")X"
'    ssql = ssql & " Order By cUser,Titular"
'**********end ALPA
'**********ALPA 18/03/2008
    sSql = "stp_sel_ReporteMoraDiariaXAnalista '" & Format(pdFecha, "yyyy-mm-dd") & "', '" & psUser & "','" & psProductos & "','" & psCondicion & "' ,'" & psAgencias & "', '" & psMoneda & "', '" & pnIniDiaMora & "','" & pnFinDiaMora & "', '" & pnMontoIni & "','" & pnMontoFin & "', '" & psCodUbiGeo & "'"
'**********End ALPA
    Set ListaMoraXAnalistaCabecera = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaMoraXAnalistaDetalle(ByVal psCtaCod As String, ByVal psFecha As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
'peac 20071025 se agrego al query capital,interes,mora y gastos
    
Set oConec = New DConecta
sSql = "Select CONVERT(varchar(15),CC.dVenc,103) as dFecha,"
sSql = sSql & " SaldoGastoMora=isnull((Select Sum(nMonto-nMontoPagado)"
sSql = sSql & " From ColocCalendDet NOLOCK"
sSql = sSql & " where cCtaCod=CC.cCtaCod and nNroCalen=CC.nNroCalen and nColocCalendApl=CC.nColocCalendApl and"
sSql = sSql & " nCuota=CC.nCuota and nPrdConceptoCod in (1101) and nPrdConceptoCod like '12%'),0),"
sSql = sSql & " capital=isnull((Select Sum(nMonto-nMontoPagado)"
sSql = sSql & "          From ColocCalendDet NOLOCK"
sSql = sSql & "          Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen"
sSql = sSql & "          and nColocCalendApl=CC.nColocCalendApl and nCuota=CC.nCuota"
sSql = sSql & "          and nPrdConceptoCod in (1000)),0),"
sSql = sSql & " Interes=isnull((Select Sum(nMonto-nMontoPagado)"
sSql = sSql & "          From ColocCalendDet NOLOCK"
sSql = sSql & "          Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen"
sSql = sSql & "          and nColocCalendApl=CC.nColocCalendApl and nCuota=CC.nCuota"
sSql = sSql & "          and nPrdConceptoCod in (1100)),0),"
sSql = sSql & " Mora=isnull((Select Sum(nMonto-nMontoPagado)"
sSql = sSql & "          From ColocCalendDet NOLOCK"
sSql = sSql & "          Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen"
sSql = sSql & "          and nColocCalendApl=CC.nColocCalendApl and nCuota=CC.nCuota"
sSql = sSql & "          and nPrdConceptoCod in (1101)),0),"
sSql = sSql & " Gastos=isnull((Select Sum(nMonto-nMontoPagado)"
sSql = sSql & "          From ColocCalendDet NOLOCK"
sSql = sSql & "          Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen"
sSql = sSql & "          and nColocCalendApl=CC.nColocCalendApl and nCuota=CC.nCuota"
sSql = sSql & "          and nPrdConceptoCod like '12%'),0),"
sSql = sSql & " SaldoCuota=isnull((Select Sum(nMonto-nMontoPagado)"
sSql = sSql & " From ColocCalendDet NOLOCK"
sSql = sSql & " where cCtaCod=CC.cCtaCod and nNroCalen=CC.nNroCalen and nColocCalendApl=CC.nColocCalendApl and"
sSql = sSql & " nCuota=CC.nCuota),0),"
sSql = sSql & " DateDiff(day,CC.dVenc,'" & Format(psFecha, "MM/dd/yyyy") & "') as AtrasoCuota,"
sSql = sSql & " CC.nCuota"
sSql = sSql & " From ColocCalendario CC"
sSql = sSql & " Where CC.cCtaCod='" & psCtaCod & "' and CC.nColocCalendApl=1  and CC.nColocCalendEstado=0"
sSql = sSql & " and CC.dVenc<'" & Format(psFecha, "MM/dd/yyyy") & "'"
sSql = sSql & " and nNroCalen=(Select Max(nNroCalen) From ColocCalendario Where cCtaCod='" & psCtaCod & "')"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaMoraXAnalistaDetalle = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaDesembolsosXAgenciaXFecha(ByVal psCodAgen As String, ByVal dFecha As Date) As Recordset
    Dim oConec As DConecta
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
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaDesembolsosXAgenciaXFecha = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ListaPrendarioConsolidadAnalista(ByVal psCodAgencia As String, ByVal pdFecha As String, _
ByVal pbFecha As Boolean) As Recordset
    Dim oConec As DConecta
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
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaPrendarioConsolidadAnalista = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerTipoCambio(ByVal pFechaCambio As Date) As Double
    Dim Rs As ADODB.Recordset
    Dim oConec As DConecta
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

    Set oConec = New DConecta
    oConec.AbreConexion
    Set Rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not Rs.EOF And Not Rs.BOF Then
        nValFijo = Rs!nValFijo
    End If
    Set Rs = Nothing
    
    ObtenerTipoCambio = nValFijo
End Function

Public Function ListaClientesAutomaticos(ByVal dFecha As Date, ByVal psAgencias As String, ByVal psAnalistas As String, _
ByVal psProductos As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    Set oConec = New DConecta
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
    Dim oConec As DConecta
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

    Set oConec = New DConecta
    oConec.AbreConexion
    Set Repor108205_ClientesVigentes = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ListaCreditosNoDesembolsados(ByVal pdFechaInicio As Date, ByVal psFechaFin As Date, _
ByVal psAgencia As String) As Recordset
  
  Dim oConec As DConecta
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
  
  Set oConec = New DConecta
  oConec.AbreConexion
  Set ListaCreditosNoDesembolsados = oConec.CargaRecordSet(sSql)
  oConec.CierraConexion
  Set oConec = Nothing
End Function


Public Function ListaCreditosVencidos(ByVal sAgencias As String, ByVal pdFecha As Date) As Recordset
  Dim oConec As DConecta
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
  
  Set oConec = New DConecta
  oConec.AbreConexion
  Set ListaCreditosVencidos = oConec.CargaRecordSet(sSql)
  oConec.CierraConexion
  Set oConec = Nothing
End Function


Public Function ObtenerGarantes(ByVal psCtaCod As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    sSql = "Select Pers.cPersNombre,Pers.cPersDireccDomicilio"
    sSql = sSql & " from ProductoPersona PP"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Where  PP.nPrdPersRelac=25 and PP.cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ObtenerGarantes = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function Repor_HistorialCliente(ByVal pscPerscod As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    sSql = "Select P.cCtaCod,Cn.cConsDescripcion as EstadoCredito,C.nMontoCol,P.nSaldo,CE.dPrdEstado,C.dVenc,"
    sSql = sSql & " RH.cUser,CN1.cConsDescripcion as TipoCredito,Pers1.cPersNombre as cNomIntitucion,"
    sSql = sSql & " Moneda=Case When Substring(P.cCtaCod,9,1)='1' Then 'SOLES' Else 'DOLARES' End,"
    sSql = sSql & " PID.cPersIDnro,Pers.cPersNombre as cNomCli"
    sSql = sSql & " From Persona Pers"
    sSql = sSql & " Inner Join PersId  PID on PID.cPersCod=Pers.cPersCod"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cPersCod=Pers.cPersCod and PP.nPrdPersRelac=20"
    sSql = sSql & " Inner Join Producto  P on P.cCtaCod=PP.cCtaCod"
    sSql = sSql & " Inner Join Constante CN on CN.nConsValor=P.nPrdEstado and CN.nConsCod=3001 and CN.nConsValor<>3001"
    sSql = sSql & " Inner Join Colocaciones C on C.cCtaCod=P.cCtaCod"
    sSql = sSql & " Left Join ColocacEstado CE on CE.cCtaCod=P.cCtaCod and CE.nPrdEstado=2002"
    sSql = sSql & " Left Join ProductoPersona PP1 on PP1.cCtacod=P.cCtaCod and PP1.nPrdPersRelac=28"
    sSql = sSql & " Left Join RRHH RH on RH.cPersCod=PP1.cPersCod"
    sSql = sSql & " Inner Join Constante CN1 on CN1.nConsValor=SubString(P.cctaCod,6,3) and CN1.nConsCod=1001 and CN1.nConsValor<>1001"
    sSql = sSql & " Left Join  ColocacConvenio CC on CC.cCtaCod=P.cCtaCod"
    sSql = sSql & " Left Join Persona Pers1 on Pers1.cPersCod=CC.cPersCod"
    sSql = sSql & " Where PP.cPersCod='" & pscPerscod & "'"
    sSql = sSql & " Order By C.dVenc"
     
    Set oConec = New DConecta
    oConec.AbreConexion
    Set Repor_HistorialCliente = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function Repor_HistorialCreditoCabecera(ByVal psCtaCod As String) As Recordset
   Dim oConec As DConecta
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
   
   Set oConec = New DConecta
   oConec.AbreConexion
   Set Repor_HistorialCreditoCabecera = oConec.CargaRecordSet(sSql)
   oConec.CierraConexion
   Set oConec = Nothing
End Function


Public Function Repo_HistorialCreditoDetalle(ByVal psCtaCod As String) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
    
sSql = "Select Y.nNroCuota,  Y.dVenc, Y.FechaPago,"
sSql = sSql & " Capital=IsNull((Select Sum(nMonto) From MovColDet Where nMovNro=M1.nMovNro and nPrdConceptoCod in (1000) and nNroCuota=Y.nNroCuota),0),"
sSql = sSql & " Interes=IsNull((Select Sum(nMonto) From MovColDet Where nMovNro=M1.nMovNro and nPrdConceptoCod in (1100) and nNroCuota=Y.nNroCuota),0),"
sSql = sSql & " Mora=IsNull((Select Sum(nMonto) From MovColDet Where nMovNro=M1.nMovNro and nPrdConceptoCod in (1101) and nNroCuota=Y.nNroCuota),0),"
sSql = sSql & " Gastos=Isnull((Select Sum(nMonto) From MovColDet Where nMovNro=M1.nMovNro and Cast(nPrdConceptoCod  as varchar(10)) like '12%' and nNroCuota=Y.nNroCuota),0),"
sSql = sSql & " ITF=Isnull((Select Sum(nMonto) From MovColDet Where nMovNro=M1.nMovNro and nPrdConceptoCod=20),0), Right(M1.cMovNro,4) as cUsuario,"
sSql = sSql & " nDiasAtraso = DateDiff(Day, CCO.dVenc, FechaPago)"
sSql = sSql & " From Mov M1"
sSql = sSql & " Inner Join (Select nMovNro,nNroCuota,dVenc,FechaPago,cCtaCod,nNroCalen,nColocCalendApl"
sSql = sSql & " From (Select M.nMovNro, MD.nNroCuota, CO.dVenc, Cast((Substring(Left(M.cMovNro,8),5,2)+'/'+ Substring(Left(M.cMovNro,8),7,2)+'/'+Substring(Left(M.cMovNro,8),1,4)) as Datetime) as FechaPago, MD.cCtaCod,CD.nNroCalen,CO.nColocCalendApl"
sSql = sSql & " From Mov M"
sSql = sSql & " Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
sSql = sSql & " Inner Join ColocCalendario CO on CO.cCtaCod=MD.cCtaCod  and Co.nCuota=MD.nNroCuota and CO.nColocCalendApl=(Case When Md.cOpeCod in ('100101','100102','100103','100104','100105') Then 0 else 1 end)"
sSql = sSql & " Inner Join ColocacCred CD on CD.cCtaCod=MD.cCtaCod and CD.nNroCalen=MD.nNroCalen"
sSql = sSql & " Where MD.cCtaCod='" & psCtaCod & "' and M.nMovFlag=0 and"
sSql = sSql & " CO.nNroCalen=(Select nNroCalen From ColocacCred Where cCtaCod='" & psCtaCod & "') and MD.cOpeCod not like '99%') X"
sSql = sSql & " Group By nMovNro,nNroCuota,dVenc,FechaPago,cCtaCod,nNroCalen,nColocCalendApl)Y on M1.nMovNro=Y.nMovNro"
sSql = sSql & " Inner Join ColocCalendario CCO on CCO.cCtaCod=Y.cCtaCod and CCO.nCuota=Y.nNroCuota and Y.nNroCalen=CCO.nNroCalen and Y.nColocCalendApl=CCO.nColocCalendApl"
sSql = sSql & " Order By M1.nMovNro"

Set oConec = New DConecta
oConec.AbreConexion
Set Repo_HistorialCreditoDetalle = oConec.CargaRecordSet(sSql)
oConec.CierraConexion
Set oConec = Nothing
End Function

Function Repo_HistorialCredDetalleComplem(ByVal psCtaCod As String, ByVal pdFecha As Date) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
        
    sSql = "Select Co.nCuota,CO.dVenc,NULL as FechaPago,0 as Captital,0 as Interes,"
    sSql = sSql & " 0 as Mora,0 as Gastos,0 as ITF,NULL as cUsuario,DateDiff(Day,CO.dVenc,'" & CStr(Format(pdFecha, "MM/dd/yyyy")) & "') as nDiasAtraso"
    sSql = sSql & " From ColocCalendario CO"
    sSql = sSql & " Inner Join ColocacCred CC on CO.cCtaCod=CC.cCtaCod and CO.nNroCalen=CC.nNroCalen"
    sSql = sSql & " Where Co.cCtaCod='" & psCtaCod & "' and CO.nColocCalendApl=1 and Co.nColocCalendEstado=0"
    sSql = sSql & " and (Select Sum(nMontoPagado) From ColocCalendDet Where cCtaCod=CO.cCtaCod and nCuota=CO.nCuota and nColocCalendApl=CO.nColocCalendApl and"
    sSql = sSql & " nNroCalen=CO.nNroCalen)=0"
    sSql = sSql & " Order By CO.nCuota,nColocCalendApl"
 
    Set oConec = New DConecta
    oConec.AbreConexion
    Set Repo_HistorialCredDetalleComplem = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function Recup_ClientesByAnalista(ByVal psAnalista As String, ByVal psAgencias As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    sSql = "Select Distinct Pers.cPersCod,Pers.cPersNombre,Pers.cPersDireccDomicilio,Pers.cPersTelefono,Pers.cPersTelefono2,RH.cUser"
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=20"
    sSql = sSql & " Inner Join ProductoPersona PP1 on P.cCtaCod=PP1.cCtaCod and PP1.nPrdPersRelac=28"
    sSql = sSql & " Inner JOin Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Inner Join RRHH Rh on Rh.cPersCod=PP1.cPersCod"
    sSql = sSql & " Where Rh.cPersCod in (" & psAnalista & ") and SubString(P.cCtaCod,4,2) in ( " & psAgencias & ") and P.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205)"
    sSql = sSql & " Order By RH.cUser,Pers.cPersNombre"

    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set Recup_ClientesByAnalista = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function Recup_CalidadCartera(ByVal pnTipoCambio As Double)
    Dim oConec  As DConecta
    Dim sSql As String
    
    'Recupera la calidad de la cartera
    
    sSql = ""
End Function

Public Function Recup_MontoAtrasado(ByVal pdFecha As Date, ByVal psCtaCod As String) As Double
    Dim oConec As DConecta
    Dim sSql As String
    Dim Rs As ADODB.Recordset
    
    sSql = "Select Sum(CD.nMonto-CD.nMontoPagado) as nMonto"
    sSql = sSql & " From ColocacCred C"
    sSql = sSql & " Inner Join  ColocCalendario CO on C.cCtaCod=CO.cCtaCod and C.nNroCalen=CO.nNroCalen"
    sSql = sSql & " Inner Join ColocCalendDet CD on CD.cCtaCod=CO.cCtaCod and  CD.nNroCalen=CO.nNroCalen And CD.nColocCalendApl=CO.nColocCalendApl and CD.nCuota=CO.nCuota"
    sSql = sSql & " Where CD.cCtaCod='" & psCtaCod & "' and dVenc<='" & Format(pdFecha, "MM/dd/yyyy") & "' and nColocCalendEstado=0"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set Rs = oConec.CargaRecordSet(sSql, adLockReadOnly)
    oConec.CierraConexion
    
    If Not Rs.EOF And Not Rs.BOF Then
        Recup_MontoAtrasado = Format(Rs!nMonto, "#0.00")
    End If
        
    Set Rs = Nothing
End Function

Function VerificaConvenioEmp(ByVal psCtaCod As String) As Boolean
    'funcion que verifica si existe convenio a traves del filtro de los gastos
    'se hizo para el gasto de credicom
    Dim sSql As String
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    
    sSql = "Select Count(*) as nCantidad"
    sSql = sSql & " From ColocacConvenio  CC"
    sSql = sSql & " Inner Join ProductoConceptoFiltro PCF on SubString(CC.cCtaCod,6,3)=PCF.nProdCod and SubString(CC.cCtaCod,4,2)=PCF.cAgeCod and PCF.cIntitucion=CC.cPersCod"
    sSql = sSql & " Where CC.cCtaCod='" & psCtaCod & "'"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set Rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
        
    VerificaConvenioEmp = False
        
    If Not Rs.EOF And Not Rs.BOF Then
        If Val(Rs!nCantidad) > 0 Then
            VerificaConvenioEmp = True
        Else
            VerificaConvenioEmp = False
        End If
    Else
        VerificaConvenioEmp = False
    End If
    
    Set Rs = Nothing
End Function

Public Function Recup_CreditoAntiguo(ByVal psCtaCod As String) As String
    Dim sSql As String
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    
    sSql = "Select cCtaCodAnt From RelCuentas Where cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set Rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not Rs.EOF And Not Rs.BOF Then
        Recup_CreditoAntiguo = Rs!CctacodAnt
    End If
    Set Rs = Nothing
End Function

Public Function Recup_FechaCancelacion(ByVal psCtaCod As String) As String
    Dim sSql As String
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    
    sSql = "Select C.dVenc"
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join Colocaciones C on C.cCtaCod=P.cCtaCod"
    sSql = sSql & " Where P.cCtaCod='" & psCtaCod & "' and P.nPrdEstado=2050"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set Rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not Rs.EOF And Not Rs.BOF Then
        Recup_FechaCancelacion = IIf(IsNull(Rs!dVenc), "NO TIENE CANCELACION", Format(Rs!dVenc, "DD/mm/yyyy"))
    End If
    Set Rs = Nothing
    If Recup_FechaCancelacion = "" Then
        Recup_FechaCancelacion = "NO TIENE CANCELACION"
    End If
End Function

Public Function GetUsuario(ByVal pnMovNro As Long) As String
    Dim sSql As String
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    
    sSql = "Select Right(cMovNro,4) as cUser From Mov Where nMovNro=" & pnMovNro
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set Rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    If Not Rs.EOF And Not Rs.BOF Then
        GetUsuario = Rs!cUser
    End If
    Set Rs = Nothing
End Function

Public Function GetCalidadCartera(ByVal pnTipoCambio As Double) As Recordset
    Dim oConec As DConecta
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

    Set oConec = New DConecta
    oConec.AbreConexion
    Set GetCalidadCartera = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GetCalidadCarteraCastigados(ByVal pnTipoCambio As Double) As Recordset
    Dim oConec As DConecta
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
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set GetCalidadCarteraCastigados = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
End Function

Public Function GetCalidadCarteraConsumo(ByVal pnTipoCambio As Double, ByVal psEstados As String, _
ByVal pnDiaIni As Integer, ByVal pnDiaFin As Integer) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    
    sSql = "Select Sum (Case SubString(cCtaCod,9,1)"
    sSql = sSql & " When '1' Then (nSaldoCap-nCapVencido)"
    sSql = sSql & " When '2' Then (nSaldoCap-nCapVencido)* " & pnTipoCambio & " End)"
    sSql = sSql & " From [dbcmacicaconsol].[dbo].[CreditoConsolTotal]"
    sSql = sSql & " Where SubString(cCtaCod,6,3) like '[34]%' and nDiasAtraso>=" & pnDiaIni & " and nDiasAtraso<=" & pnDiaFin & " and"
    sSql = sSql & " nPrdEstado in(" & psEstados & ")"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set GetCalidadCarteraConsumo = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaCreditosMoraXTelefono(ByVal psCodAnalistas As String, ByVal pdFecha As Date, ByVal pnDiaInicio As Integer, _
                                           ByVal pnDiaFin As Integer, ByVal psCodAgen As String) As Recordset

    Dim oConec As DConecta
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

    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCreditosMoraXTelefono = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Function ListaAnalistasAgencia(ByVal psCodAgencia As String) As String
    Dim oConec As DConecta
    Dim sSql As String
    Dim Rs As ADODB.Recordset
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
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set Rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    nCont = 0
    Do Until Rs.EOF
        If nCont = 0 Then
            sAnalistas = "'" & Rs!cPersCod & "'"
        Else
            sAnalistas = sAnalistas & ",'" & Rs!cPersCod & "'"
        End If
        nCont = nCont + 1
        Rs.MoveNext
    Loop
    Set Rs = Nothing
    ListaAnalistasAgencia = sAnalistas
End Function

Public Function ListaCreditosConvenioCancaledos(ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, _
                                                 ByVal pnMoneda As String, _
                                                 ByVal psCodAgencia As String, Optional ByVal psCodInstitucion As String) As Recordset

    Dim oConec As DConecta
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
    
    sSql = sSql & " and C.cAgeCodAct='" & psCodAgencia & "'" 'BRGO 07/06/2010 BASILEA II
    sSql = sSql & " Order By PersInstitucion.cPersNombre,P.dPrdEstado "
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCreditosConvenioCancaledos = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GetPersona(ByVal psPersCod As String) As String
    Dim oConec As DConecta
    Dim sSql As String
    Dim sNombre As String
    Dim Rs As ADODB.Recordset
    
    Set oConec = New DConecta
    oConec.AbreConexion
    sSql = "Select cPersNombre From Persona Where cPersCod='" & psPersCod & "'"
    Set Rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not Rs.EOF And Not Rs.BOF Then
        GetPersona = Rs!cpersnombre
    End If
    Set Rs = Nothing
End Function

Public Function GetAgencia(ByVal psCodAgen As String) As String
    Dim oConec As DConecta
    Dim sSql As String
    Dim sNombre As String
    Dim Rs As ADODB.Recordset
    
    Set oConec = New DConecta
    oConec.AbreConexion
    sSql = " Select cAgeDescripcion From Agencias Where cAgeCod='" & psCodAgen & "'"
    Set Rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not Rs.EOF And Not Rs.BOF Then
        GetAgencia = Rs!cAgeDescripcion
    End If
    Set Rs = Nothing
End Function

'Public Function ObtenerCarteraVigentePorUserStored(ByVal psCodAgen As String, ByVal psUser As String, ByVal pdFecha As Date) As Recordset
'
' Dim oConec As DConecta
' Dim nTipoCambio As Double
' Dim oDCredDoc As DCredDoc
'
' Set oDCredDoc = New DCredDoc
' nTipoCambio = oDCredDoc.ObtenerTipoCambio(pdFecha)
' Set oDCredDoc = Nothing
'
' Set oConec = New DConecta
' oConec.AbreConexion
' Set ObtenerCarteraVigentePorUserStored = oConec.Ejecutar("ListaVigentes " & psCodAgen & ",'" & psUser & "', " & nTipoCambio)
' oConec.CierraConexion
' Set oConec = Nothing
'End Function

Public Function ListaCreditosXInstitucionDesembolsados(ByVal psCodAgen As String, ByVal psCodInstitucuion As String, _
                                                       ByVal pdFechaInicial As Date, ByVal dFechaFinal As Date) As Recordset

    Dim oConec As DConecta
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
    sSql = sSql & " and CC.cAgeCodAct='" & psCodAgen & "' and MD.cOpeCod in ('100101','100102','100103','100104','100105') and" ' BRGO 07/06/2010 BASILEA II
    sSql = sSql & " MD.cOpeCod not like '107[123456789]%' and MD.nPrdConceptoCod=1000 and "
    sSql = sSql & " C.cPersCod='" & psCodInstitucuion & "'"
    sSql = sSql & " Order By cast(SubString(Left(M.cMovNro,8),5,2)+'/'+SubString(Left(M.cMovNro,8),7,2)+'/'+SubString(Left(M.cMovNro,8),1,4) as Datetime)"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCreditosXInstitucionDesembolsados = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
        
End Function


Public Function ListaConsolidadoPorAnalistaV2(ByVal pnTipoCambio As Double, ByVal psAgenciaCod As String, ByVal pdFecha As Date, _
ByVal psAnalistas As String, ByVal pbFecha As Boolean) As Recordset

    Dim sSql As String
    Dim oConec As DConecta
    Dim sFecha As String
    Dim sDia As String
    Dim sMes As String
    Dim sAno As String
    Dim sFechaMov As String
    Dim sdInicioMes As String
    Dim nTipoCambio As Double
    Dim odCredDoc As DCredDoc
    
    sFechaMov = Format(pdFecha, "MM/dd/yyyy")
    sAno = DatePart("yyyy", pdFecha)
    sMes = Right("00", 2 - Len(CStr(DatePart("m", pdFecha)))) & CStr(DatePart("m", pdFecha))
    sDia = Right("00", 2 - Len(CStr(DatePart("d", pdFecha)))) & CStr(DatePart("d", pdFecha))
    
    Set odCredDoc = New DCredDoc
    nTipoCambio = odCredDoc.ObtenerTipoCambio(pdFecha)
    
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
    Set oConec = New DConecta
    oConec.AbreConexion
        Set ListaConsolidadoPorAnalistaV2 = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function ColCredRepCredSaldosDiarios(ByVal psAgenciaCod As String, ByVal pdFecha As Date) As Recordset

    Dim sSql As String
    Dim oConec As DConecta
    sSql = " Select cAgeCod,cAgeDescripcion,nConsValor,cConsdescripcion,"
    sSql = sSql & " CSoles , CDolares, JSoles, JDolares, JCSoles, JCDolares, TSoles, TDolares"
    sSql = sSql & " From"
    sSql = sSql & " ("
    sSql = sSql & "     select substring(cctacod,4,2) Age, substring(cctacod,6,3) Prod,"
    sSql = sSql & "     sum(case when substring(cctacod,9,1)= '1' and not nPrdEstado in (2201,2202) then nSaldoCap else 0 end ) CSoles,"
    sSql = sSql & "     sum(case when substring(cctacod,9,1)= '2' and not nPrdEstado in (2201,2202)then nSaldoCap else 0 end ) CDolares,"
    sSql = sSql & "     sum(case when substring(cctacod,9,1)= '1' and nPrdEstado = 2201 then nSaldoCap else 0 end ) JSoles,"
    sSql = sSql & "     sum(case when substring(cctacod,9,1)= '2' and nPrdEstado = 2201 then nSaldoCap else 0 end ) JDolares,"
    sSql = sSql & "     sum(case when substring(cctacod,9,1)= '1' and nPrdEstado = 2202 then nSaldoCap else 0 end ) JCSoles,"
    sSql = sSql & "     sum(case when substring(cctacod,9,1)= '2' and nPrdEstado = 2202 then nSaldoCap else 0 end ) JCDolares,"
    sSql = sSql & "     sum(case when substring(cctacod,9,1)= '1' then nSaldoCap else 0 end ) TSoles,"
    sSql = sSql & "     sum(case when substring(cctacod,9,1)= '2' then nSaldoCap else 0 end ) TDolares"
    sSql = sSql & "     from colocacsaldo where"
    sSql = sSql & "     datediff(day,dfecha,'" & Format(pdFecha, "MM/DD/YYYY") & "')=0 and substring(cctacod,4,2) = '" & Right(psAgenciaCod, 2) & "'"
    sSql = sSql & "     Group by substring(cctacod,4,2),substring(cctacod,6,3)"
    sSql = sSql & " ) S"
    sSql = sSql & " Inner Join Agencias A on cAgeCod = S.Age"
    sSql = sSql & " Inner Join Constante C on nConsValor = convert(int,S.Prod) and nConsCod = 1001"
    sSql = sSql & " Order by cAgeCod, nConsValor"
    
    Set oConec = New DConecta
    oConec.AbreConexion
        Set ColCredRepCredSaldosDiarios = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerParametroComision() As Double
    Dim oConec As DConecta
    Dim sSql As String
    Dim Rs As ADODB.Recordset
    
    sSql = "Select nParamValor"
    sSql = sSql & " From ColocParametro"
    sSql = sSql & " Where nParamVar=102731"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set Rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not Rs.EOF And Not Rs.BOF Then
        ObtenerParametroComision = Rs!nParamValor
    End If
    Set Rs = Nothing
End Function

Public Function ListaCredNoDesembolsados(ByVal psCodAgen As String, ByVal pdFechaIni As Date, _
ByVal pdFechaFin As Date) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
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
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCredNoDesembolsados = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function Repo_108506(ByVal pdFechaInicial As Date, ByVal pdFechaPagoInicial As Date, ByVal pdFechaPagoFin As Date) As Recordset
    Dim sFechaInicial As String
    Dim sFechaInicialPag As String
    Dim sFechaFinalPag As String
    Dim oConec As DConecta
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
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set Repo_108506 = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ListaCredAnaPlusMonto(ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, psCodAgen As String, _
pnTipoCambio As Double, ByVal pnMonto As Double) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
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
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCredAnaPlusMonto = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaCredAnaLessMonto(ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, psCodAgen As String, _
pnTipoCambio As Double, ByVal pnMonto As Double) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
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
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCredAnaLessMonto = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaCreditosRefinanciados(ByVal pdFechaFin As Date, ByVal psCodAgen As String, _
                 ByVal pnTipoCambio As Double) As ADODB.Recordset
                 
    Dim sSql As String
    Dim oConec As DConecta
    
    
    sSql = "Select Sum(Case When SubString(CS.cCtaCod,9,1)='1' Then CS.nSaldoCap Else CS.nSaldoCap*" & pnTipoCambio & "End),RH.cUser"
    sSql = sSql & " From ColocacSaldo CS"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtacod=CS.cCtaCod and PP.nPrdPersRelac=28"
    sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
    sSql = sSql & " Where SubString(CS.cCtacod,4,2)='" & psCodAgen & "' and CS.nPrdEstado in (2030,2031,2032) and CS.dFecha='" & Format(pdFechaFin, "MM/dd/yyyy") & "'"
    sSql = sSql & " Group By RH.cUser"
    sSql = sSql & " Order By RH.cUser"
        
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCreditosRefinanciados = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaCreditosCancelacionProyectados(ByVal pdFechaInicio As Date, ByVal pdFechaFin As Date, _
                    ByVal pMatAnalistas As Variant, ByVal pMatProductos As Variant) As ADODB.Recordset
                 
    Dim sSql As String
    Dim oConec As DConecta
    Dim i As Integer
    Dim lsAnalistas As String
    Dim lsProductos As String
    
    lsAnalistas = "('"
    For i = 0 To UBound(pMatAnalistas) - 1
        lsAnalistas = lsAnalistas & pMatAnalistas(i) & "','"
    Next i
    lsAnalistas = Mid(lsAnalistas, 1, Len(lsAnalistas) - 2) & ")"
    
    lsProductos = "('"
    For i = 0 To UBound(pMatProductos) - 1
        lsProductos = lsProductos & pMatProductos(i) & "','"
    Next i
    lsProductos = Mid(lsProductos, 1, Len(lsProductos) - 2) & ")"
    
    
    'sSQL = "SELECT P.cCtaCod,Titular= T.cPersNombre,T.cPersDireccDomicilio,T.cPersTelefono,C.nMontoCol,P.nSaldo," & _
           " CC.nDiasAtrasoAcum , C.dVenc, CD.nMonto , A.cPersNombre" & _
           " FROM Producto P " & _
           " INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod " & _
           " INNER JOIN ColocacCred CC ON P.cCtaCod = CC.cCtaCod " & _
           " INNER JOIN ProductoPersona PP ON P.cCtaCod =PP.cCtaCod AND PP.nPrdPersRelac= 20 " & _
           " INNER JOIN Persona T ON PP.cPersCod = T.cPersCod " & _
           " INNER JOIN ProductoPersona ANA ON P.cCtaCod =ANA.cCtaCod AND ANA.nPrdPersRelac= 28 " & _
           " INNER JOIN Persona A ON A.cPersCod = ANA.cPersCod " & _
           " INNER JOIN ColocCalendario CAL ON P.cCtaCod = CAL.cCtaCod AND CAL.nNroCalen = CC.nNroCalen " & _
           " AND CAL.nColocCalendApl=0 AND nCuota=1 " & _
           " INNER JOIN ColocCalendDet CD ON CAL.cCtaCod = CD.cCtaCod AND CAL.nNroCalen= CD.nNroCalen " & _
           " AND CAL.nColocCalendApl = CD.nColocCalendApl AND CAL.nCuota = CD.nCuota AND CD.nPrdConceptoCod=1000 " & _
           " WHERE ANA.cPersCod IN " & lsAnalistas & " AND SUBSTRING(P.cCtaCod,6,3) IN " & lsProductos & _
           " AND C.dVenc BETWEEN '" & CStr(Format(pdFechaInicio, "MM/dd/yyyy")) & "' AND '" & CStr(Format(pdFechaFin, "MM/dd/yyyy")) & "'" & _
           " ORDER BY A.cPersNombre"
    sSql = " SELECT DISTINCT CASE WHEN (GROUPING(A.cPersNombre)=1) THEN 'ZZ'ELSE ISNULL(A.cPersNombre,'PERSONA DESCONOCIDA') END as Analista," & _
           " CASE WHEN (GROUPING(CON.cConsDescripcion)=1) THEN 'ZZ'ELSE CON.cConsDescripcion END as Producto," & _
           " Cuenta=ISNULL(C.cCtaCod,'ZZ')," & _
           " Titular= dbo.ColocCredTitular(C.cCtaCod), " & _
           " RTrim(dbo.ColocCredTitularDir(C.cCtaCod))as Direccion," & _
           " Telefono=dbo.ColocCredTitularTelefono(C.cCtaCod)," & _
           " Plazo=(SELECT nPlazo FROM ColocacEstado WHERE cCtaCod=C.cCtaCod AND nPrdEstado=2002), " & _
           " MontoAprob=(SELECT nMontoCol FROM Colocaciones WHERE cCtaCod= C.cCtaCod)," & _
           " Saldo=(SELECT nSaldo FROM Producto WHERE cCtaCod=C.cCtaCod)," & _
           " DiasAtrasoAcum=(SELECT nDiasAtrasoAcum FROM ColocacCred WHERE cCtaCod = C.cCtaCod)," & _
           " Desembolso=ISNULL((SELECT CONVERT(varchar(10),dPrdEstado,103) FROM ColocacEstado WHERE cCtaCod= C.cCtaCod AND nPrdEstado=2020 AND " & _
           " cDescripcion LIKE 'Desembolso%'),''), " & _
           " Vencimiento=(SELECT CONVERT(varchar(10),dVenc,103) FROM Colocaciones WHERE cCtaCod= C.cCtaCod), " & _
           " Moneda=CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1'THEN 'S/.'ELSE 'US $'END, " & _
           " NroCuotas =(select max(nCuota)  from ColocCalendario where cCtaCod=C.cCtaCod and  nNroCalen =3) " & _
           " FROM Colocaciones C " & _
           " INNER JOIN ProductoPersona ANA ON C.cCtaCod =ANA.cCtaCod AND ANA.nPrdPersRelac= 28 " & _
           " INNER JOIN Persona A ON A.cPersCod = ANA.cPersCod " & _
           " INNER JOIN Constante CON ON CON.nConsValor=SUBSTRING(C.cCtaCod,6,3) AND CON.nConsCod=1001 " & _
           " WHERE ANA.cPersCod IN " & lsAnalistas & " " & _
           " AND SUBSTRING(C.cCtaCod,6,3) IN " & lsProductos & " AND C.dVenc >= '" & Format(pdFechaInicio, "yyyymmdd") & "' AND C.dVenc <=  '" & Format(pdFechaFin, "yyyymmdd") & "'" & _
           " GROUP BY A.cPersNombre,CON.cConsDescripcion,C.cCtaCod WITH ROLLUP "
           
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCreditosCancelacionProyectados = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

'ARCV 28-10-2006
Public Function RecuperaCreditosVigentesXInstituciones(ByVal pdFecSis As Date, pMatInstitucion As Variant, pMatAgencias As Variant) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
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
    sSql = sSql & " sCuotaPend = (Select CONVERT(varchar(5),CC.nNroProxCuota) From ColocacCred CC Inner Join ColocacEstado CE ON CC.cCtaCod = CE.cCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSql = sSql & "        Where CC.cCtaCod = Prd.cCtaCod ),"
    sSql = sSql & " SUM(CDet.nMonto) as nSaldoCuota,"
    sSql = sSql & " nCuota = (Select SUM(nMonto) From ColocCalendDet CCal Where CCal.cCtaCod = Prd.cCtaCod And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CCal.cCtaCod) And nColocCalendApl = 1 AND nCuota = (Select nNroProxCuota From ColocacCred Where cCtaCod = CCal.cCtaCod)) ,"
    sSql = sSql & " SUM(Prd.nSaldo) as nSaldoCap,"
    'sSql = sSql & " nDiasAtraso = (Select nDiasAtraso From ColocacCred Where cCtaCod = Prd.cCtaCod),"
    sSql = sSql & " SUM(CDet2.nMonto) as nSaldoMora,"
    sSql = sSql & " dFecVencimiento = (Select dVenc From ColocCalendario CCal Where CCal.cCtaCod = Prd.cCtaCod And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CCal.cCtaCod) And nColocCalendApl = 1 And nCuota = (Select nNroProxCuota From ColocacCred Where cCtaCod = CCal.cCtaCod)), "
    sSql = sSql & " nNroCuotas = convert(varchar(5), (Select MAX(nCuota) From ColocCalendario CCal Where CCal.cCtaCod = Prd.cCtaCod And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CCal.cCtaCod) And nColocCalendApl = 1 ) ),"
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
    sSql = sSql & "     Where Cdet.nColocCalendApl = 1  AND CC.dVenc <= '" & Format(pdFecSis, "mm/dd/yyyy") & "' "
    sSql = sSql & "     Group By Cdet.cCtaCod, Cdet.nNroCalen"
    sSql = sSql & "        ) as CDet ON CDet.cCtaCod = Prd.cCtaCod And CDet.nNroCalen = CC.nNroCalen "
    sSql = sSql & " Inner Join (Select cCtaCod, nNroCalen, SUM(Round(nMonto,2)-nMontoPagado) as nMonto "
    sSql = sSql & "     From ColocCalendDet CDet"
    sSql = sSql & "     Where nColocCalendApl = 1 And nPrdConceptoCod = 1101"
    sSql = sSql & "     Group By cCtaCod, nNroCalen "
    sSql = sSql & "        ) as CDet2 ON CDet2.cCtaCod = Prd.cCtaCod And CDet2.nNroCalen = CC.nNroCalen "
    sSql = sSql & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor
    'sSql = sSql & " And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' And A.cAgeCod IN " & sCadAge & " AND nDiasAtraso > 0 "
    sSql = sSql & " Group By A.cAgeDescripcion, P.cPersNombre, Prd.cCtaCod"
    sSql = sSql & " WITH ROLLUP"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosVigentesXInstituciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'ARCV 03-02-2007
Public Function RecuperaCreditosVigentesXCalificacion(ByVal pdFecSis As Date, pMatAnalistas As Variant, pMatAgencias As Variant, _
                        ByVal pMatCalificacion As Variant) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadAna As String
Dim sCadCalif As String

    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    sCadAna = "('"
    For i = 0 To UBound(pMatAnalistas) - 1
        sCadAna = sCadAna & pMatAnalistas(i) & "','"
    Next i
    sCadAna = Mid(sCadAna, 1, Len(sCadAna) - 2) & ")"
    
    sCadCalif = "('"
    For i = 0 To UBound(pMatCalificacion) - 1
        sCadCalif = sCadCalif & pMatCalificacion(i) & "','"
    Next i
    sCadCalif = Mid(sCadCalif, 1, Len(sCadCalif) - 2) & ")"
    
    sSql = "SELECT CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'        " & _
           " ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')   END AS sAgencia, " & _
           " CASE WHEN GROUPING(P.cPersNombre)=1 THEN 'TOTAL' " & _
           " ELSE ISNULL(P.cPersNombre, 'ANALISTA DESCONOCIDA')   END AS sAnalista, " & _
           " CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL' " & _
           " ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDA')   END AS cCtaCod, " & _
           " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = Prd.cCtaCod), " & _
           " sDireccion = (select P2.cPersDireccDomicilio From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = Prd.cCtaCod), " & _
            " SUM(C.nMontoCol) as nMontoDesemb, cClasificacion= dbo.ObtenerCalificacionDiasAtraso_Cuenta(Prd.cCtaCod), " & _
            " nAtrasoProm= dbo.ObtenerCalificacionPromedioDias_Cuenta (Prd.cCtaCod),  SUM(Prd.nSaldo) as nSaldoCap, " & _
            " sCuotaPend = CONVERT(varchar(5),(SELECT nNroProxCuota FROM ColocacCred WHERE cCtaCod=Prd.cCtaCod)), " & _
            " nSaldoCuota= dbo.ColocCredCuotaPendiente(Prd.cCtaCod), " & _
            " nPlazo= (SELECT nCuotas FROM ColocacEstado WHERE cCtaCod= Prd.cCtaCod AND nPrdEstado= 2002), " & _
            " dFecVencimiento = (Select dVenc From ColocCalendario CCal Where CCal.cCtaCod = Prd.cCtaCod And nNroCalen = (SELECT nNroCalen FROM ColocacCred WHERE cCtaCod= Prd.cCtaCod) " & _
            "       And nColocCalendApl = 1 And nCuota = (SELECT nNroProxCuota FROM ColocacCred WHERE cCtaCod=Prd.cCtaCod) ) " & _
            " From Producto Prd Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod " & _
            " Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2) " & _
            " INNER JOIN ProductoPersona PP ON Prd.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = 28 " & _
            " INNER JOIN Persona P ON P.cPersCod = PP.cPersCod " & _
            " Where Prd.nPrdEstado >= 2020 And Prd.nPrdEstado <= 2032 " & _
            " AND A.cAgeCod IN " & sCadAge & " AND PP.cPersCod IN " & sCadAna & " AND dbo.ObtenerCalificacionDiasAtraso_Cuenta(Prd.cCtaCod) IN " & sCadCalif & _
            " Group By A.cAgeDescripcion, P.cPersNombre, Prd.cCtaCod " & _
            " WITH ROLLUP ORDER BY sAgencia, sAnalista, cCtaCod "


    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosVigentesXCalificacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'ARCV 05-02-2007
Public Function RecuperaDatosListadoPoliza(ByVal pdFecSis As Date, pMatAgencias As Variant) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim i As Integer
Dim sCadAge As String

    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    sSql = " SELECT  Prd.cCtaCod," & _
           " cCodPolizaAseg= ISNULL(cCodPolizaAseg,''), " & _
            " cCiaSeguros= ASEG.cPersNombre, " & _
            "dVigenciaAseg= ISNULL(dVigenciaAseg,0), " & _
            "dVencAseg= ISNULL(dVencAseg,0), " & _
            "cCondicionPoliza=CASE WHEN dVencAseg<= '" & Format(pdFecSis, "mm/dd/yyyy") & "' THEN 'VIGENTE'ELSE 'VENCIDO'END, " & _
            "cDireccionInmueble=ISNULL(cDirGarRegPubli,''), " & _
            "nSumaAsegurada = IsNull(nSumaAsegurada, 0) " & _
            "Into #ColocPoliza " & _
            "FROM    Producto Prd INNER JOIN ColocGarantia CG ON Prd.cCtaCod = CG.cCtaCod " & _
            "INNER JOIN Garantias G ON CG.cNumGarant = G.cNumGarant " & _
            "INNER JOIN GarantReal GR ON GR.cNumGarant = G.cNumGarant " & _
            "INNER JOIN GarantRealTasacion GRT ON GRT.cNumGarant = GR.cNumGarant " & _
            "AND GRT.dTasacion = (SELECT MAX(dTasacion)FROM GarantRealTasacion WHERE cNumGarant=GRT.cNumGarant) " & _
            "INNER JOIN GarantPoliza GP ON GP.cNumGarant = GRT.cNumGarant AND GP.dTasacion = GRT.dTasacion " & _
            "INNER JOIN Poliza Pol ON GP.cNumPoliza = Pol.cNumPoliza " & _
            "INNER JOIN Persona ASEG ON ASEG.cPersCod = Pol.cPersCodAseg " & _
            "Where Prd.nPrdEstado >= 2020 And Prd.nPrdEstado <= 2032 " & _
            "AND SUBSTRING(Prd.cCtaCod,4,2) IN " & sCadAge & " AND GP.nEstado= 1"
            
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
        
    sSql = " SELECT CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL' " & _
            "ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')   END AS sAgencia, " & _
            "CASE WHEN GROUPING(P.cPersNombre)=1 THEN 'TOTAL' " & _
            "ELSE ISNULL(P.cPersNombre, 'CLIENTE DESCONOCIDO')   END AS sCliente, " & _
            "CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL' " & _
            "ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDA')   END AS cCtaCod, " & _
            "SUM(C.nMontoCol) as nMontoDesemb, SUM(Prd.nSaldo) as nSaldoCap, " & _
            "cMoneda=CASE WHEN SUBSTRING(Prd.cCtaCod,9,1)='1'THEN 'MN'ELSE 'ME'END, " & _
            "dDesembolso=(SELECT dVigencia FROM Colocaciones WHERE cCtaCod =Prd.cCtaCod ), " & _
            "dTermino = (SELECT MAX(dVenc) FROM ColocCalendario WHERE cCtaCod= Prd.cCtaCod AND nColocCalendApl=1 " & _
            "    AND nNroCalen = (SELECT nNroCalen FROM ColocacCred WHERE cCtaCod=Prd.cCtaCod)), " & _
            "cCodPolizaAseg= ( SELECT TOP 1 cCodPolizaAseg FROM #ColocPoliza WHERE cCtaCod = Prd.cCtaCod), " & _
            "cCiaSeguros= ( SELECT TOP 1 cCiaSeguros FROM #ColocPoliza WHERE cCtaCod = Prd.cCtaCod), " & _
            "dVigenciaAseg= ( SELECT TOP 1 dVigenciaAseg FROM #ColocPoliza WHERE cCtaCod = Prd.cCtaCod), " & _
            "dVencAseg= ( SELECT  TOP 1 dVencAseg FROM #ColocPoliza WHERE cCtaCod = Prd.cCtaCod), " & _
            "cCondicionPoliza= ( SELECT TOP 1 cCondicionPoliza FROM #ColocPoliza WHERE cCtaCod = Prd.cCtaCod), " & _
            "cDireccionInmueble= ( SELECT TOP 1 cDireccionInmueble FROM #ColocPoliza WHERE cCtaCod = Prd.cCtaCod), " & _
            "nSumaAsegurada = ( SELECT TOP 1 nSumaAsegurada FROM #ColocPoliza WHERE cCtaCod = Prd.cCtaCod) " & _
            "From Producto Prd Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod " & _
            "Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2) " & _
            "INNER JOIN ProductoPersona PP ON Prd.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = 20 " & _
            "INNER JOIN Persona P ON P.cPersCod = PP.cPersCod INNER JOIN #ColocPoliza CP ON CP.cCtaCod = Prd.cCtaCod " & _
            "Where Prd.nPrdEstado >= 2020 And Prd.nPrdEstado <= 2032 " & _
            " AND A.cAgeCod IN " & sCadAge & " Group By A.cAgeDescripcion, P.cPersNombre, Prd.cCtaCod" & _
            " WITH ROLLUP ORDER BY sAgencia, sCliente, cCtaCod"

    Set RecuperaDatosListadoPoliza = oConecta.CargaRecordSet(sSql)
    
    sSql = "DROP TABLE #ColocPoliza"
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'ARCV 07-02-2007
Public Function RecuperaDatosGruposEconomicos() As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim i As Integer
    
    sSql = " SELECT nGrupoCod=PG.nGrupoCod,cGrupoEcon=G.cConsDescripcion, " & _
            "cPersCodCli=P.cPersCod,    cPersNombreCli=P.cPersNombre ," & _
            "cTipoPer= CASE WHEN PID.cPersIDTPO = 1 THEN 'NAT'ELSE'JUR'END, " & _
            "cTipoDoc= CASE WHEN PID.cPersIDTPO = 1 THEN 'DNI'ELSE'RUC'END, " & _
            "cNumeroDoc= PID.cPersIDNro, cDomicilio= P.cPersDireccDomicilio, " & _
            "cVinculacion= ISNULL(V.cConsDescripcion,''), cPersCodOtro=ISNULL(cPersCodOtro,''), " & _
            "cPersNombreOtro=ISNULL(OTRO.cPersNombre,''), cNumeroDocOtro= ISNULL(OID.cPersIDNro,''), " & _
            "cDomicilioOtro= ISNULL(OTRO.cPersDireccDomicilio,''),  " & _
            "nPorcenOtro=ISNULL(nPorcenOtro,0), cCargoOtro=ISNULL(cCargoOtro,''),  " & _
            "nOrden = 2 " & _
            "INTO #GrupoEcon " & _
            "FROM PersGrupoEconomico PG " & _
            "INNER JOIN Constante G ON PG.nGrupoCod= G.nConsValor AND G.nConsCod= 9069 " & _
            "INNER JOIN Persona P ON PG.cPersCod = P.cPersCod " & _
            "AND P.cPersCod IN ( " & _
            "SELECT cPersCod FROM ProductoPersona PP INNER JOIN Producto Prd ON " & _
            "    PP.cCtaCod = Prd.cCtaCod And PP.NPRDPERSRELAC = 20 " & _
            "WHERE Prd.nPrdEstado IN (2020,2021,2022,2030,2031,2032) " & _
            " )  " & _
            "INNER JOIN PersID PID ON PID.cPersCod = P.cPersCod " & _
            "LEFT JOIN Persona OTRO ON PG.cPersCodOtro=OTRO.cPersCod " & _
            "LEFT JOIN PersID OID ON OID.cPersCod = OTRO.cPersCod " & _
            "LEFT JOIN Constante V ON PG.nTipoVinculacion = V.nConsValor AND V.nConsCod = 9070 "
            
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    
    sSql = "INSERT INTO #GrupoEcon " & _
           " SELECT DISTINCT nGrupoCod,cGrupoEcon,'','','','','','','','','','','',0,'',1 " & _
           " FROM #GrupoEcon "
    
    oConecta.Ejecutar (sSql)
    
    sSql = " SELECT nOrden,* FROM #GrupoEcon ORDER BY nGrupoCod,nOrden "

    Set RecuperaDatosGruposEconomicos = oConecta.CargaRecordSet(sSql)
    
    sSql = "DROP TABLE #GrupoEcon"
    
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'ARCV 07-02-2007
Public Function RecuperaDetalleGarantiasEnSoles(ByVal pMatAgencias As Variant, _
                                                ByVal pdFechaIni As Date, _
                                                ByVal pdFechaFin As Date) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim i As Integer
Dim sCadAge As String
Dim sCadAna As String
Dim sCadCalif As String

    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    ' BRGO Modificado 06/06/2010 Basilea II
    sSql = "SELECT * FROM (SELECT CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'        " & _
           " ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')   END AS sAgencia, " & _
            "CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL' " & _
            "ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDA')   END AS cCtaCod, " & _
            "dFecha = (SELECT dVigencia FROM Colocaciones WHERE cCtaCod=Prd.cCtaCod), " & _
            "sAnalista= ISNULL((SELECT RRHH.cUser From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 28 INNER JOIN RRHH ON RRHH.cPersCod=PP.cPersCod Where PP.cCtaCod = Prd.cCtaCod),''), " & _
            "sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = Prd.cCtaCod), " & _
            "nMontoDesemb= dbo.ColocObtieneMontoEnSoles(C.nMontoCol,SUBSTRING(Prd.cCtaCod,9,1)), " & _
            "nPreferidas=SUM(CASE WHEN G.IdSupGarant= 1 THEN dbo.ColocObtieneMontoEnSoles(G.nRealizacion,G.nMoneda) ELSE 0 END), " & _
            "nNoPreferidas=SUM(CASE WHEN G.IdSupGarant= 3 THEN  dbo.ColocObtieneMontoEnSoles(G.nRealizacion,G.nMoneda) ELSE 0 END), " & _
            "nAutoliquidables=SUM(CASE WHEN G.IdSupGarant= 2 THEN  dbo.ColocObtieneMontoEnSoles(G.nRealizacion,G.nMoneda) ELSE 0 END), " & _
            "nTotal = Sum(dbo.ColocObtieneMontoEnSoles(G.nRealizacion,G.nMoneda)) " & _
        "From Garantias G INNER JOIN ColocGarantia CG ON G.cNumGarant = CG.cNumGarant " & _
            "INNER JOIN Producto Prd ON CG.cCtaCod = Prd.cCtaCod " & _
            "Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod " & _
            "Inner join Agencias A ON A.cAgeCod = C.cAgeCodAct " & _
            "Where Prd.nPrdEstado >= 2020 And Prd.nPrdEstado <= 2032 " & _
            "AND A.cAgeCod IN " & sCadAge & _
            " AND DATEDIFF(day,C.dVigencia,'" & Format(pdFechaIni, "mm-dd-yyyy") & "')<=0 AND  DATEDIFF(day,C.dVigencia,'" & Format(pdFechaFin, "mm-dd-yyyy") & "')>=0" & _
            " GROUP BY A.cAgeDescripcion, Prd.cCtaCod,C.nMontoCol " & _
            "WITH ROLLUP)X WHERE ISNULL(nMontoDesemb,0)<>0 OR cCtaCod='TOTAL'" & _
            "ORDER BY sAgencia, cCtaCod "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDetalleGarantiasEnSoles = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaGarantiasXCliente(ByVal pMatAgencias As Variant) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim i As Integer
Dim sCadAge As String

    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    sSql = "SELECT cTitular=P.cPersCod, cPersNombre=P.cPersNombre, cCtaCod=Prd.cCtaCod, " & _
            " cEstado= E.cConsDescripcion," & _
            "cMonedaC=CASE WHEN SUBSTRING(Prd.cCtaCod,9,1)='1'THEN'SOLES'ELSE'DOLARES'END, " & _
            "nMontoCol,dVigencia,cUser=ANA.cUser, " & _
            "cPersCodG=TITU.cPersCod,    cNumGarant=G.cNumGarant, " & _
            "cMonedaG=CASE WHEN G.nMoneda=1 THEN'SOLES'ELSE'DOLARES'END, " & _
            "cTipo= REPLACE(T.cConsDescripcion,'GARANTIAS',''), " & _
            "sSubTipo=S.cConsDescripcion, " & _
            "nMontoSoles= CASE WHEN G.nMoneda=1 THEN CG.nGravado ELSE 0 END, " & _
            "nMontoDolares= CASE WHEN G.nMoneda=2 THEN CG.nGravado ELSE 0 END, " & _
            "nOrden = 3 " & _
            "INTO #GarantiasRep " & _
            "FROM Persona P INNER JOIN ProductoPersona PP ON  P.cPersCod = PP.cPersCod AND PP.nPrdPersRelac = 20 " & _
            "INNER JOIN Producto Prd ON Prd.cCtaCod = PP.cCtaCod " & _
            "INNER JOIN Colocaciones C ON C.cCtaCod = Prd.cCtaCod " & _
            "INNER JOIN ProductoPersona PP2 ON  PP2.cCtaCod = Prd.cCtaCod AND PP2.nPrdPersRelac = 28 " & _
            "INNER JOIN RRHH ANA ON ANA.cPersCod = PP2.cPersCod " & _
            "INNER JOIN ColocGarantia CG ON Prd.cCtaCod= CG.cCtaCod " & _
            "INNER JOIN Garantias G ON G.cNumGarant = CG.cNumGarant " & _
            "INNER JOIN Constante T ON T.nConsValor = G.IdSupGarant AND T.nConsCod=1048 " & _
            "INNER JOIN Constante S ON S.nConsValor = G.nTpoGarantia AND S.nConsCod=1027 " & _
            "INNER JOIN PersGarantia PG ON PG.cNumGarant = G.cNumGarant AND PG.nRelacion=1 " & _
            "INNER JOIN Persona TITU ON TITU.cPersCod = PG.cPersCod " & _
            "INNER JOIN Constante E ON E.nConsValor = Prd.nPrdEstado AND E.nConsCod = 3001 " & _
            "WHERE Prd.nPrdEstado IN (2020,2021,2022,2030,2031,2032) AND G.nestado not in (5) AND SUBSTRING(Prd.cCtaCod,4,2) IN " & sCadAge

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    
    sSql = " INSERT INTO #GarantiasRep " & _
            "SELECT DISTINCT cTitular, cPersNombre='', cCtaCod, " & _
            "cEstado,cMonedaC,nMontoCol, dVigencia,cUser, " & _
            "cPersCodG='',cNumGarant='', " & _
            "cMonedaG='', " & _
            "cTipo= '', " & _
            "sSubTipo='', " & _
            "nMontoSoles= 0.00, " & _
            "nMontoDolares= 0.00, " & _
            "nOrden = 2 " & _
            "FROM #GarantiasRep"
    oConecta.Ejecutar (sSql)
    
    sSql = "INSERT INTO #GarantiasRep " & _
            "SELECT DISTINCT cTitular, cPersNombre, cCtaCod='', " & _
            "cEstado='',cMonedaC='', " & _
            "nMontoCol=0,dVigencia=0 ,cUser='', " & _
            "cPersCodG='',   cNumGarant='', " & _
            "cMonedaG='', " & _
            "cTipo= '', " & _
            "sSubTipo='', " & _
            "nMontoSoles=0, " & _
            "nMontoDolares=0, " & _
            "nOrden = 1 " & _
            "FROM #GarantiasRep WHERE cPersNombre<>''"
            
    oConecta.Ejecutar (sSql)
    
    'ALPA 05/03/2008 se quito redundancia del campo norden
    
'    ssql = "SELECT nOrden,cTitular, cPersNombre=CASE WHEN nOrden=3 THEN ''ELSE cPersNombre END, " & _
'           "cCtaCod,cEstado,cMonedaC,nMontoCol,dVigencia,cUser, " & _
'           " cPersCodG , cNumGarant, cMonedaG, cTipo, sSubTipo, nMontoSoles, nMontoDolares, nOrden " & _
'           "FROM #GarantiasRep ORDER BY cTitular,cCtaCod,cNumGarant,nOrden"


    sSql = "SELECT nOrden,cTitular, cPersNombre=CASE WHEN nOrden=3 THEN ''ELSE cPersNombre END, " & _
           "cCtaCod,cEstado,cMonedaC,nMontoCol,dVigencia,cUser, " & _
           " cPersCodG , cNumGarant, cMonedaG, cTipo, sSubTipo, nMontoSoles, nMontoDolares " & _
           "FROM #GarantiasRep ORDER BY cTitular,cCtaCod,cNumGarant,nOrden"
    
    Set RecuperaGarantiasXCliente = oConecta.CargaRecordSet(sSql)
    
    sSql = "DROP TABLE #GarantiasRep"
    oConecta.Ejecutar (sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaGarantiasXGarantes(ByVal pMatAgencias As Variant) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim sCadAge As String
Dim i As Integer

    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    
    sSql = "SELECT G.cNumGarant,cCtaCod=CG.cCtaCod, " & _
            "cGarante = GAR.cPersNombre, " & _
            "cPersCod= GAR.cPersCod, " & _
            "Prd.nPrdEstado, " & _
            "ANA.cUser,dVigencia, " & _
            "cMoneda=CASE WHEN SUBSTRING(Prd.cCtaCod,9,1)='1'THEN'SOLES'ELSE'DOLARES'END, " & _
            "nMontoCol, " & _
            "cTipo= REPLACE(T.cConsDescripcion,'GARANTIAS','') " & _
            "INTO #GarantiasRep " & _
            "FROM Garantias G INNER JOIN ColocGarantia CG ON G.cNumGarant = CG.cNumGarant " & _
            "INNER JOIN Producto Prd ON Prd.cCtaCod = CG.cCtaCod " & _
            "INNER JOIN ProductoPersona PP ON Prd.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = 28 " & _
            "INNER JOIN RRHH ANA ON ANA.cPersCod = PP.cPersCod " & _
            "INNER JOIN ProductoPersona PP2 ON Prd.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = 25 " & _
            "INNER JOIN Persona GAR ON GAR.cPersCod = PP2.cPersCod " & _
            "INNER JOIN PersGarantia PG ON PG.cNumGarant = G.cNumGarant AND PG.cPersCod=GAR.cPersCod AND PG.nRelacion=1 " & _
            "INNER JOIN Colocaciones C ON Prd.cCtaCod = C.cCtaCod " & _
            "INNER JOIN Constante T ON T.nConsValor = G.IdSupGarant AND T.nConsCod=1048 " & _
            "WHERE Prd.nPrdEstado IN (2020,2021,2022,2030,2031,2032) AND G.nestado not in (5) AND C.cAgeCodAct IN " & sCadAge 'BRGO 06/06/2010

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    
    
    sSql = "SELECT  CASE WHEN GROUPING(M.cConsDescripcion)=1 THEN 'ZTOTAL' " & _
            "ELSE ISNULL(M.cConsDescripcion, 'MONEDA DESCONOCIDA')   END AS sMoneda, " & _
            "CASE WHEN GROUPING(GR.cGarante)=1 THEN 'ZTOTAL'        ELSE ISNULL(GR.cGarante, 'GARANTE DESCONOCIDO')   END AS sGarante, " & _
            "CASE WHEN GROUPING(G.cNumGarant)=1 THEN 'ZTOTAL'        ELSE ISNULL(G.cNumGarant, 'GARANTIA DESCONOCIDA')   END AS sGarantia, " & _
            "cCtaCod = ISNULL(GR.cCtaCod,'ZTOTAL'), " & _
            "cPersCod= (SELECT top 1 cPersCod FROM #GarantiasRep WHERE cNumGarant= G.cNumGarant AND cCtaCod=GR.cCtaCod), " & _
            "dDesembolso= (SELECT top 1 dVigencia FROM #GarantiasRep WHERE cNumGarant= G.cNumGarant AND cCtaCod=GR.cCtaCod), " & _
            "sAnalista= (SELECT top 1 cUser FROM #GarantiasRep WHERE cNumGarant= G.cNumGarant AND cCtaCod=GR.cCtaCod), " & _
            "nMontoCol= (SELECT top 1 nMontoCol FROM #GarantiasRep WHERE cNumGarant= G.cNumGarant AND cCtaCod=GR.cCtaCod), " & _
            "cMoneda= (SELECT top 1 cMoneda FROM #GarantiasRep WHERE cNumGarant= G.cNumGarant " & _
            "        AND cCtaCod=GR.cCtaCod), " & _
            "cTipo= (SELECT top 1 cTipo FROM #GarantiasRep WHERE cNumGarant= G.cNumGarant " & _
            "        AND cCtaCod=GR.cCtaCod), " & _
            "SUM(G.nRealizacion) as nMontoGarantia, " & _
            "SUM(CG.nGravado) As nMontoGarantizado " & _
            "FROM Garantias G " & _
            "INNER JOIN Constante M ON M.nConsValor = G.nMoneda AND M.nConsCod=1011 " & _
            "INNER JOIN ColocGarantia CG ON G.cNumGarant =  CG.cNumGarant " & _
            "INNER JOIN Producto Prd ON Prd.cCtaCod = CG.cCtaCod " & _
            "INNER JOIN #GarantiasRep GR ON G.cNumGarant = GR.cNumGarant AND Prd.cCtaCod = GR.cCtaCod " & _
            "WHERE Prd.nPrdEstado IN (2020,2021,2022,2030,2031,2032) AND G.nestado not in (5) AND SUBSTRING(Prd.cCtaCod,4,2) IN " & sCadAge & _
            "GROUP BY M.cConsDescripcion,GR.cGarante,G.cNumGarant,GR.cCtaCod " & _
            "With ROLLUP " & _
            "HAVING ISNULL(GR.cCtaCod,'')<>'' OR ISNULL(G.cNumGarant,'')=''-- AND ISNULL(GR.cCtaCod,'')='') " & _
            "ORDER BY sMoneda,sGarante,ISNULL(G.cNumGarant,'ZTOTAL'),ISNULL(GR.cCtaCod,'ZTOTAL') "
    
    Set RecuperaGarantiasXGarantes = oConecta.CargaRecordSet(sSql)
    
    sSql = "DROP TABLE #GarantiasRep"
    oConecta.Ejecutar (sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaGarantiasXTipoDetallado(ByVal pMatAgencias As Variant) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim i As Integer
Dim sCadAge As String

    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    ' BRGO Modificado el 06/06/2010 BASILEA II
    sSql = "SELECT cNumGarant,cCtaCod=CG.cCtaCod,Prd.nPrdEstado, " & _
            "P.cPersCod,P.cPersNombre,dVigencia,cMoneda=CASE WHEN SUBSTRING(Prd.cCtaCod,9,1)='1'THEN'SOLES'ELSE'DOLARES'END, " & _
            "nMontoCol, C.cAgeCodAct " & _
            "INTO #GarantiasRep " & _
            "FROM ColocGarantia CG INNER JOIN Producto Prd ON Prd.cCtaCod = CG.cCtaCod " & _
            "INNER JOIN ProductoPersona PP ON Prd.cCtaCod = PP.cCtaCod AND nPrdPersRelac = 20 " & _
            "INNER JOIN Persona P ON P.cPersCod = PP.cPersCod " & _
            "INNER JOIN Colocaciones C ON Prd.cCtaCod = C.cCtaCod " & _
            "WHERE Prd.nPrdEstado IN (2020,2021,2022,2030,2031,2032) AND C.cAgeCodAct IN " & sCadAge

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    
    ' BRGO Modificado el 06/06/2010 BASILEA II
    sSql = "SELECT CASE WHEN (GROUPING(T.cConsDescripcion) = 1) THEN 'ZTOTAL'        " & _
            "ELSE ISNULL(T.cConsDescripcion, 'TIPO DESCONOCIDO')   END AS sTipo, " & _
            "CASE WHEN GROUPING(M.cConsDescripcion)=1 THEN 'ZTOTAL' ELSE ISNULL(M.cConsDescripcion, 'MONEDA DESCONOCIDA')   END AS sMoneda, " & _
            "CASE WHEN GROUPING(S.cConsDescripcion)=1 THEN 'ZTOTAL' ELSE ISNULL(S.cConsDescripcion, 'SUBTIPO DESCONOCIDO')   END AS sSubTipo, " & _
            "CASE WHEN GROUPING(G.cNumGarant)=1 THEN 'ZTOTAL' " & _
            "ELSE ISNULL(G.cNumGarant, 'GARANTIA DESCONOCIDA')   END AS sGarantia, " & _
            "cCtaCod = ISNULL((SELECT cCtaCod FROM #GarantiasRep WHERE cNumGarant= G.cNumGarant AND cCtaCod=GR.cCtaCod),'ZTOTAL'), " & _
            "dGarantia = (SELECT dGarantia FROM Garantias WHERE cNumGarant = G.cNumGarant), " & _
            "SUM(G.nRealizacion) as nMontoGarantia, " & _
            "cPersCod = (SELECT cPersCod FROM #GarantiasRep WHERE cNumGarant= G.cNumGarant AND cCtaCod=GR.cCtaCod), " & _
            "cPersNombre= (SELECT cPersNombre FROM #GarantiasRep WHERE cNumGarant= G.cNumGarant AND cCtaCod=GR.cCtaCod), " & _
            "dVigencia= (SELECT dVigencia FROM #GarantiasRep WHERE cNumGarant= G.cNumGarant AND cCtaCod=GR.cCtaCod), " & _
            "cMoneda= (SELECT cMoneda FROM #GarantiasRep WHERE cNumGarant= G.cNumGarant AND cCtaCod=GR.cCtaCod), " & _
            "nMontoCol= (SELECT nMontoCol FROM #GarantiasRep WHERE cNumGarant= G.cNumGarant AND cCtaCod=GR.cCtaCod) " & _
            "FROM Garantias G " & _
            "INNER JOIN Constante T ON T.nConsValor = G.IdSupGarant AND T.nConsCod=1048 " & _
            "INNER JOIN Constante M ON M.nConsValor = G.nMoneda AND M.nConsCod=1011 " & _
            "INNER JOIN Constante S ON S.nConsValor = G.nTpoGarantia AND S.nConsCod = 1027 " & _
            "INNER JOIN ColocGarantia CG ON G.cNumGarant =  CG.cNumGarant " & _
            "INNER JOIN Producto P ON P.cCtaCod = CG.cCtaCod " & _
            "INNER JOIN #GarantiasRep GR ON G.cNumGarant = GR.cNumGarant AND P.cCtaCod = GR.cCtaCod " & _
            "WHERE P.nPrdEstado IN (2020,2021,2022,2030,2031,2032) AND G.nestado not in (5) AND GR.cAgeCodAct IN " & sCadAge & _
            "GROUP BY T.cConsDescripcion, M.cConsDescripcion, S.cConsDescripcion,G.cNumGarant,GR.cCtaCod   " & _
            "With ROLLUP " & _
            "ORDER BY sTipo,sMoneda,sSubTipo,sGarantia,ISNULL(GR.cCtaCod,'ZTOTAL')"
    
    Set RecuperaGarantiasXTipoDetallado = oConecta.CargaRecordSet(sSql)
    
    sSql = "DROP TABLE #GarantiasRep"
    oConecta.Ejecutar (sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function


Public Function RecuperaGarantiasXTipoResumido(ByVal pMatAgencias As Variant) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim sCadAge As String
Dim i As Integer
    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    
    
    sSql = "SELECT CASE WHEN (GROUPING(T.cConsDescripcion) = 1) THEN 'ZTOTAL'        " & _
            "ELSE ISNULL(T.cConsDescripcion, 'TIPO DESCONOCIDO')   END AS sTipo, " & _
            "CASE WHEN GROUPING(M.cConsDescripcion)=1 THEN 'ZTOTAL' " & _
            "ELSE ISNULL(M.cConsDescripcion, 'MONEDA DESCONOCIDA')   END AS sMoneda, " & _
            "CASE WHEN GROUPING(S.cConsDescripcion)=1 THEN 'ZTOTAL' " & _
            "ELSE ISNULL(S.cConsDescripcion, 'SUBTIPO DESCONOCIDO')   END AS sSubTipo, " & _
            "SUM(G.nRealizacion) As nMontoGarantia " & _
            "FROM Garantias G " & _
            "INNER JOIN Constante T ON T.nConsValor = G.IdSupGarant AND T.nConsCod=1048 " & _
            "INNER JOIN Constante M ON M.nConsValor = G.nMoneda AND M.nConsCod=1011 " & _
            "INNER JOIN Constante S ON S.nConsValor = G.nTpoGarantia AND S.nConsCod = 1027 " & _
            "INNER JOIN ColocGarantia CG ON G.cNumGarant = CG.cNumGarant " & _
            "INNER JOIN Producto P ON P.cCtaCod = CG.cCtaCod " & _
            "WHERE P.nPrdEstado IN (2020,2021,2022,2030,2031,2032) AND G.nestado not in (5) " & _
            "GROUP BY T.cConsDescripcion, M.cConsDescripcion, S.cConsDescripcion " & _
            "With ROLLUP " & _
            "ORDER BY sTipo,sMoneda,sSubTipo "
    
    Set RecuperaGarantiasXTipoResumido = oConecta.CargaRecordSet(sSql)
        
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaDatosListaReprogramados(ByVal pMatAgencias As Variant, _
                                ByVal pdFecFin As Date, ByVal pdFecIni As Date) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim i As Integer
Dim sCadAge As String

    sCadAge = "('"
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & "','"
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    
    '**Comentado por DAOR 20070828
'    ssql = "SELECT Prd.cCtaCod,Prd.nSaldo,nMontoRef=SUM(CR.nMontoRef) " & _
'            "INTO #ColocRefin " & _
'            "From Producto Prd Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod " & _
'            "INNER JOIN ColocacRefinanc CR ON CR.cCtaCod = C.cCtaCod AND nEstado=3 " & _
'            "WHERE Prd.nPrdEstado IN(2030,2031,2032) " & _
'            "AND SUBSTRING(Prd.cCtaCod,4,2) IN" & sCadAge & _
'            "and Prd.cCtaCod IN (SELECT cCtaCod FROM ColocacRefinanc WHERE nEstado=3) " & _
'            " AND DATEDIFF(day,C.dVigencia,'" & Format(pdFecFin, "mm-dd-yyyy") & "')<=0 AND  DATEDIFF(day,C.dVigencia,'" & Format(pdFecIni, "mm-dd-yyyy") & "')>=0 " & _
'            "Group By Prd.cCtaCod,Prd.nSaldo "
'
'    oconecta.Ejecutar ssql
'
'
'    ssql = "SELECT CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'        " & _
'            "ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')   END AS sAgencia, " & _
'            "CASE WHEN GROUPING(M.cConsDescripcion)=1 THEN 'TOTAL' " & _
'            "ELSE ISNULL(M.cConsDescripcion, 'MONEDA DESCONOCIDA')   END AS sMoneda, " & _
'            "CASE WHEN GROUPING(R.cUser)=1 THEN 'TOTAL' " & _
'            "ELSE ISNULL(R.cUser, 'ANALISTA DESCONOCIDO')   END AS sAnalista, " & _
'            "CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL' " & _
'            "ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDA')   END AS cCtaCod, " & _
'            "sCliente = (SELECT P.cPersNombre FROM ProductoPersona PP INNER JOIN Persona P ON P.cPersCod=PP.cPersCod " & _
'            "        AND nPrdPersRelac=20 WHERE PP.cCtaCod=Prd.cCtaCod), " & _
'            "dVigencia=(SELECT dVigencia FROM Colocaciones WHERE cCtaCod =Prd.cCtaCod ), " & _
'            "nMontoDesemb= SUM(CR.nMontoRef), nCapPagado= SUM(CR.nMontoRef-CR.nSaldo), " & _
'            "SUM(Prd.nSaldo) As nSaldoCap " & _
'            "From Producto Prd Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod " & _
'            "Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2) " & _
'            "INNER JOIN ProductoPersona PP ON Prd.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = 28 " & _
'            "INNER JOIN Persona P ON P.cPersCod = PP.cPersCod " & _
'            "INNER JOIN RRHH R ON R.cPersCod = P.cPersCod " & _
'            "INNER JOIN Constante M ON M.nConsValor = SUBSTRING(Prd.cCtaCod,9,1)AND M.nConsCod=1011 " & _
'            "INNER JOIN #ColocRefin CR ON Prd.cCtaCod= CR.cCtaCod " & _
'            "Where Prd.nPrdEstado IN(2030,2031,2032) AND A.cAgeCod IN" & sCadAge & _
'            " AND Prd.cCtaCod IN (SELECT cCtaCod FROM ColocacRefinanc WHERE nEstado=3) " & _
'            " AND DATEDIFF(day,C.dVigencia,'" & Format(pdFecFin, "mm-dd-yyyy") & "')<=0 AND  DATEDIFF(day,C.dVigencia,'" & Format(pdFecIni, "mm-dd-yyyy") & "')>=0 " & _
'            "Group By A.cAgeDescripcion,M.cConsDescripcion,R.cUser,Prd.cCtaCod " & _
'            "With ROLLUP ORDER BY sAgencia, sMoneda,sAnalista, cCtaCod"
'
'    Set RecuperaDatosListaReprogramados = oconecta.CargaRecordSet(ssql)
'
'    ssql = "DROP TABLE #ColocRefin"
'    oconecta.Ejecutar ssql

    '**DAOR 20070828
    sSql = "select case when (GROUPING(A.cAgeDescripcion) = 1) then 'TOTAL' ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ') END AS sAgencia, " & _
            "   case when GROUPING(M.cConsDescripcion)=1 then 'TOTAL' ELSE ISNULL(M.cConsDescripcion, 'MONEDA DESCONOCIDA')   END AS sMoneda, " & _
            "   case when GROUPING(R.cUser)=1 then 'TOTAL' else ISNULL(R.cUser, 'ANALISTA DESCONOCIDO')   END AS sAnalista, " & _
            "   case when GROUPING(P.cCtaCod)=1 then 'TOTAL' else ISNULL(P.cCtaCod, 'CUENTA DESCONOCIDA')   END AS cCtaCod, " & _
            "   sCliente=max(P1.cPersNombre), " & _
            "   dVigencia=max(C.dVigencia), " & _
            "   nMontoDesemb= SUM(C.nMontoCol), " & _
            "   nCapPagado= SUM(C.nMontoCol-P.nSaldo), " & _
            "   nSaldoCap = Sum(P.nSaldo) " & _
            "from Producto P inner join Colocaciones C ON C.cCtaCod = P.cCtaCod " & _
            "   inner join Agencias A ON A.cAgeCod = SUBSTRING(P.cCtaCod,4,2) " & _
            "   inner join ProductoPersona PP1 ON P.cCtaCod = PP1.cCtaCod AND PP1.nPrdPersRelac = 20 " & _
            "   inner join Persona P1 ON P1.cPersCod = PP1.cPersCod " & _
            "   inner join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = 28 " & _
            "   inner join RRHH R ON R.cPersCod = PP2.cPersCod " & _
            "   inner join Constante M ON M.nConsValor = SUBSTRING(P.cCtaCod,9,1)AND M.nConsCod=1011 " & _
            "where P.nPrdEstado in(2030,2031,2032) AND A.cAgeCod IN" & sCadAge & _
            "   and DATEDIFF(day,C.dVigencia,'" & Format(pdFecFin, "mm-dd-yyyy") & "')<=0 " & _
            "   and DATEDIFF(day,C.dVigencia,'" & Format(pdFecIni, "mm-dd-yyyy") & "')>=0 " & _
            "group by A.cAgeDescripcion,M.cConsDescripcion,R.cUser,P.cCtaCod " & _
            "with ROLLUP " & _
            "order by sAgencia, sMoneda,sAnalista,cCtaCod "

    Set RecuperaDatosListaReprogramados = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'**DAOR 20070313, Creditos Desembolsados en Banco de la Nación
Public Function RecuperaCreditosDesembolsadosEnBancoNacion(ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pMatAgencias As Variant) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim i As Integer
Dim sCadAge As String

sCadAge = "('"
For i = 0 To UBound(pMatAgencias) - 1
    sCadAge = sCadAge & pMatAgencias(i) & "','"
Next i
sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    

sSql = sSql + " select  AB.vAgencia,C.cCtaCod,A.cAgeDescripcion,P1.cPersCod,P1.cPersNombre,CL.cDescripcion as LineaCredito,convert(varchar(10),CE.dPrdEstado,110) as FechaAprob, "
sSql = sSql + "     CE.nCuotas,convert(varchar(10),CE2.dPrdEstado,110) as FechaDesem,CE2.nMonto,CE2.nPeriodoGracia,substring(C.cCtaCod,9,1) as Moneda, "
sSql = sSql + "         (case "
sSql = sSql + "        when CE2.nColocCalendCod >= 10 And CE2.nColocCalendCod <= 31 Then Round((CE2.nCuotas * CE2.nPlazo) / 30, 0) "
sSql = sSql + "        when CE2.nColocCalendCod >= 40 And CE2.nColocCalendCod <= 61 Then CE2.nCuotas "
sSql = sSql + "         end) as Plazo,PT.nTasaInteres,convert(varchar(10),CC.dVenc,110) as FechaPago, "
sSql = sSql + "     isnull((select case when G.cNumGarant is not null then 'SI' end  from ColocGarantia CG inner join Garantias G on CG.cNumGarant=G.cNumGarant "
sSql = sSql + "             where CG.cCtaCod=C.cCtaCod and G.nGarantReal=1),'NO') as GarantiaReal "
sSql = sSql + " from    Colocaciones C inner join ColocConvBcoNac CB on (C.cCtaCod=CB.cCtaCod) "
sSql = sSql + "     inner join AgenciasBancoNacion AB on (CB.cCodAgeBcoNac=AB.cCodAge) "
sSql = sSql + "     inner join ProductoTasaInteres PT on (C.cCtaCod=PT.cCtaCod and PT.nPrdTasaInteres=1) "
sSql = sSql + "     inner join ProductoPersona PP1 on (C.cCtaCod=PP1.cCtaCod and PP1.nPrdPersRelac=20) "
sSql = sSql + "     inner join Persona P1 on (PP1.cPersCod=P1.cPersCod) "
sSql = sSql + "     inner join ColocLineaCredito CL on (C.cLineaCred=CL.cLineaCred) "
sSql = sSql + "     inner join ColocacEstado CE on (C.cCtaCod=CE.cCtaCod and CE.nPrdEstado=2002) "
sSql = sSql + "     inner join ColocacEstado CE2 on (C.cCtaCod=CE2.cCtaCod and CE2.nPrdEstado=2020) "
sSql = sSql + "     inner join ColocCalendario CC on (C.cCtaCod=CC.cCtaCod and CC.nNroCalen=3 and CC.nColocCalendApl=1 and CC.nCuota=1) "
'sSQL = sSQL + "     inner join Agencias A on (substring(C.cCtaCod,4,2)=A.cAgeCod) "
sSql = sSql + "     inner join Agencias A on C.cAgeCodAct=A.cAgeCod " '''''''BRGO 06/06/2010
sSql = sSql + "     Inner Join Producto P on C.cCtaCod = P.cCtaCod And P.nPrdEstado <> 2050 " 'MAVM 20120322
sSql = sSql + " where  CE2.nCuotas is not null and CE2.dPrdEstado between '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & " 23:59:59' and A.cAgeCod in " & sCadAge
sSql = sSql + " order by AB.vAgencia "

Set oConecta = New DConecta
oConecta.AbreConexion
Set RecuperaCreditosDesembolsadosEnBancoNacion = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

'ARCV 14-03-2007
Public Function RecuperaDatosCreditosReprogramados(ByVal pMatAgencias As Variant, _
                                            ByVal pdFechaIni As Date, ByVal pdFechaFin As Date) As ADODB.Recordset
    
Dim sSql As String
Dim oConecta As DConecta
Dim i As Integer
Dim sCadAge As String

sCadAge = "('"
For i = 0 To UBound(pMatAgencias) - 1
    sCadAge = sCadAge & pMatAgencias(i) & "','"
Next i
sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    '***ALPA********20080605********************* ******************************************************
    '******************************************** ******************************************************
    '******************************************** ******************************************************
    sSql = "SELECT distinct substring(M.cMovNro,1,8) dFechaRepro,M.nMovNro,A.cAgeCod,cAgeDescripcion,P.cCtaCod, C.dVigencia,Per.cPersNombre ,max(C.nMontoCol) nMonCol , " & _
            "max(MC.nSaldoCap) nSaldoCap, CE.nCuotas,CAL.nCuota, RRHH.cUser, " & _
            "cMoneda= CASE WHEN SUBSTRING(P.cCtaCod,9,1)='1'THEN'SOLES'ELSE'DOLARES'END " & _
            "From " & _
            "Producto P INNER JOIN  MovCol MC ON P.cCtaCod = MC.cCtaCod " & _
            "INNER JOIN Mov M ON MC.nMovNro = M.nMovNro " & _
            "INNER JOIN Agencias A ON SUBSTRING(P.cCtaCod,4,2)=A.cAgeCod " & _
            "INNER JOIN ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac= 20 " & _
            "INNER JOIN Persona Per ON PP.cPersCod = Per.cPersCod " & _
            "INNER JOIN ProductoPersona PP2 ON PP2.cCtaCod = P.cCtaCod AND PP2.nPrdPersRelac=28 " & _
            "INNER JOIN RRHH ON RRHH.cPersCod =PP2.cPersCod " & _
            "INNER JOIN ColocacCred CC ON CC.cCtaCod = P.cCtaCod " & _
            "INNER JOIN Colocaciones C ON C.cCtaCod = P.cCtaCod " & _
            "INNER JOIN ColocacEstado CE ON CE.cCtaCod= P.cCtaCod AND DATEDIFF(d,CE.dPrdEstado,C.dVigencia)=0 " & _
            "    AND CE.nPrdEstado = P.nPrdEstado AND CE.dPrdEstado=( " & _
            "        SELECT MAX(dPrdEstado)FROM ColocacEstado WHERE cCtaCod=P.cCtaCod AND DATEDIFF(d,dPrdEstado,C.dVigencia)=0) " & _
            "INNER JOIN ColocCalendario CAL ON CAL.cCtaCod = P.cCtaCod AND CAL.nNroCalen = CC.nNroCalen-1 " & _
            "    AND CAL.nColocCalendApl = 1 AND CAL.nCuota = ( " & _
            "        SELECT MIN(nCuota)FROM ColocCalendario WHERE cCtaCod=P.cCtaCod AND nNroCalen=CC.nNroCalen-1 " & _
            "        AND CAL.nColocCalendApl=1 AND nColocCalendEstado=0) " & _
            "WHERE  M.cOpeCod='100902' and substring(M.cMovNro,1,8) between '" & Format(pdFechaIni, "yyyymmdd") & "' and " & _
            "'" & Format(pdFechaFin, "yyyymmdd") & "' " & _
            " AND CC.cMovNroRprg is not null AND A.cAgeCod IN" & sCadAge & _
            " group by  substring(M.cMovNro,1,8),M.nMovNro,A.cAgeCod,cAgeDescripcion,P.cCtaCod, C.dVigencia,Per.cPersNombre, CE.nCuotas,CAL.nCuota, RRHH.cUser" & _
            " ORDER BY A.cAgeCod,Per.cPersNombre"

        '" AND DATEDIFF(day,C.dVigencia,'" & Format(pdFechaIni, "ddmmyyyy") & "')<=0 AND  DATEDIFF(day,C.dVigencia,'" & Format(pdFechaFin, "ddmmyyyy") & "')>=0 "
        '************************************************************************************************
        '************************************************************************************************
        '************************************************************************************************
    Set oConecta = New DConecta
    oConecta.AbreConexion
    
    Set RecuperaDatosCreditosReprogramados = oConecta.CargaRecordSet(sSql)
    
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

'**DAOR 20070418, Creditos Desembolsados Con Aprobación de Exoneración de Reglamento
Public Function RecuperaCreditosDesembConAprobExoneraReglamento(ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pMatAgencias As Variant) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim i As Integer
Dim sCadAge As String

sCadAge = "('"
For i = 0 To UBound(pMatAgencias) - 1
    sCadAge = sCadAge & pMatAgencias(i) & "','"
Next i
sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    

sSql = " Select  P.cCtaCod,P1.cPersNombre,C.dVigencia,substring(C.cCtaCod,9,1) as Moneda,C.nMontoCol, "
sSql = sSql + " RH.cUser , P.nPrdEstado, A.cAgeDescripcion "
sSql = sSql + "From    Producto P inner join Colocaciones C on P.cCtaCod=C.cCtaCod "
sSql = sSql + " inner join ColocacCred CC on C.cCtaCod=CC.cCtaCod "
sSql = sSql + " inner join ProductoPersona PP1 on C.cCtaCod=PP1.cCtaCod and PP1.nPrdPersRelac=20 "
sSql = sSql + " inner join Persona P1 on PP1.cPersCod=P1.cPersCod "
sSql = sSql + " inner join ProductoPersona PP2 on C.cCtaCod=PP2.cCtaCod and PP2.nPrdPersRelac=29 "
sSql = sSql + " inner join RRHH RH on PP2.cPersCod=RH.cPersCod "
sSql = sSql + " inner join Agencias A on (substring(C.cCtaCod,4,2)=A.cAgeCod) "
sSql = sSql + "Where   CC.nAproReglamento=1 and C.dVigencia between '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "' and A.cAgeCod in " & sCadAge
sSql = sSql + " and P.nPrdEstado in (2020,2021,2022) "
sSql = sSql + "Order by P.cCtaCod "
        
Set oConecta = New DConecta
oConecta.AbreConexion
Set RecuperaCreditosDesembConAprobExoneraReglamento = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

'**DAOR 20070419, Montos Desembolsados por Linea
Public Function RecuperaMontosDesembolsadosPorLineas(ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pMatAgencias As Variant) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim i As Integer
Dim sCadAge As String

sCadAge = "('"
For i = 0 To UBound(pMatAgencias) - 1
    sCadAge = sCadAge & pMatAgencias(i) & "','"
Next i
sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
sSql = " Select  A.cAgeDescripcion as Agencia,CL3.cDescripcion as Fondo,CL2.cDescripcion as SubFondo,"
sSql = sSql + " CL.cDescripcion as Linea,C.cLineaCred, "
sSql = sSql + " isnull((case when substring(C.cLineaCred,5,1)='1' then count(C.cCtaCod) end),0) as  CasosMN, "
sSql = sSql + " isnull((case when substring(C.cLineaCred,5,1)='1' then sum(C.nMontoCol) end),0) as DesembolsoMN, "
sSql = sSql + " isnull((case when substring(C.cLineaCred,5,1)='1' then sum(P.nSaldo) end),0) as SaldoMN, "
sSql = sSql + " isnull((case when substring(C.cLineaCred,5,1)='2' then count(C.cCtaCod) end),0) as  CasosME, "
sSql = sSql + " isnull((case when substring(C.cLineaCred,5,1)='2' then sum(C.nMontoCol) end),0) as DesembolsoME, "
sSql = sSql + " isnull((case when substring(C.cLineaCred,5,1)='2' then sum(P.nSaldo) end),0) as SaldoME "
sSql = sSql + "From Producto P inner join Colocaciones C on P.cCtaCod=C.cCtaCod "
sSql = sSql + " inner join ColocLineaCredito CL on C.cLineaCred=CL.cLineaCred "
sSql = sSql + " inner join ColocLineaCredito CL2 on left(CL.cLineaCred,4)=CL2.cLineaCred "
sSql = sSql + " inner join ColocLineaCredito CL3 on left(CL.cLineaCred,2)=CL3.cLineaCred "
sSql = sSql + " inner join Agencias A on substring(C.cCtaCod,4,2)=A.cAgeCod "
sSql = sSql + "Where P.nPrdEstado in (2020,2021,2022) "
sSql = sSql + "     and C.dVigencia between '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "' and A.cAgeCod in " & sCadAge
sSql = sSql + "Group by A.cAgeDescripcion,CL3.cDescripcion,CL2.cDescripcion,CL.cDescripcion,C.cLineaCred "
sSql = sSql + "Order by A.cAgeDescripcion,C.cLineaCred"
      
Set oConecta = New DConecta
oConecta.AbreConexion
Set RecuperaMontosDesembolsadosPorLineas = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function

'**DAOR 20070717
Public Function RecuperaConsolidadoCarteraPorAnalista(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pMatProd As Variant, Optional ByVal pnTipCam As Double) As ADODB.Recordset 'DAOR 20070717, se agregó pnTipCam
    
Dim sSql As String
Dim oConecta As DConecta
Dim sCadAge As String
Dim i As Integer
Dim sCadCondicion As String
Dim sCadProd As String
Dim sAnalistas As String
Dim oGen As DGeneral


    sCadProd = " And SUBSTRING(P.cCtaCod,6,3) in ('"
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

    sSql = "Select  "
    sSql = sSql & " case when grouping(A.cAgeDescripcion)=1 then 'TOTAL' else isnull(A.cAgeDescripcion,'AGENCIA DESCONOCIDA') end  as cAgencia, "
    sSql = sSql & " case when grouping(RH.cUser)=1 then 'TOTAL' else isnull(RH.cUser ,'ANALISTA DESCONOCIDO') END as cUser, "
    sSql = sSql & "   sum(case CC.nColocCondicion when 1 then 1 else 0 end) as nCantNuevo, "
    sSql = sSql & "   sum(case CC.nColocCondicion when 1 then case substring(P.cCtaCod,9,1) when '1' then C.nMontoCol else C.nMontoCol*" & Str(pnTipCam) & "  end else 0 end) as nMontoNuevo, "
    sSql = sSql & "   sum(case CC.nColocCondicion when 2 then 1 else 0 end) as nCantRecurrente, "
    sSql = sSql & "   sum(case CC.nColocCondicion when 2 then case substring(P.cCtaCod,9,1) when '1' then C.nMontoCol else C.nMontoCol*" & Str(pnTipCam) & "  end else 0 end) as nMontoRecurrente, "
    sSql = sSql & "   sum(case CC.nColocCondicion when 3 then 1 else 0 end) as nCantParalelo, "
    sSql = sSql & "   sum(case CC.nColocCondicion when 3 then case substring(P.cCtaCod,9,1) when '1' then C.nMontoCol else C.nMontoCol*" & Str(pnTipCam) & "  end else 0 end) as nMontoParalelo, "
    sSql = sSql & "   sum(case CC.nColocCondicion when 4 then 1 else 0 end) as nCantRefinanciado, "
    sSql = sSql & "   sum(case CC.nColocCondicion when 4 then case substring(P.cCtaCod,9,1) when '1' then C.nMontoCol else C.nMontoCol*" & Str(pnTipCam) & "  end else 0 end) as nMontoRefinanciado, "
    sSql = sSql & "   sum(case CC.nColocCondicion when 5 then 1 else 0 end) as nCantAmpliado, "
    sSql = sSql & "   sum(case CC.nColocCondicion when 5 then case substring(P.cCtaCod,9,1) when '1' then C.nMontoCol else C.nMontoCol*" & Str(pnTipCam) & "  end else 0 end) as nMontoAmpliado, "
    sSql = sSql & "   sum(case CC.nColocCondicion when 6 then 1 else 0 end) as nCantAutomatico, "
    sSql = sSql & "   sum(case CC.nColocCondicion when 6 then case substring(P.cCtaCod,9,1) when '1' then C.nMontoCol else C.nMontoCol*" & Str(pnTipCam) & "  end else 0 end) as nMontoAutomatico "
    sSql = sSql & " From RRHH RH inner join ProductoPersona PP on RH.cPersCod=PP.cPersCod and PP.nPrdPersRelac = 28 "
    sSql = sSql & "   inner join Producto P on PP.cCtaCod=P.cCtaCod "
    sSql = sSql & "   inner join Colocaciones C on P.cCtaCod=C.cCtaCod "
    sSql = sSql & "   inner join ColocacCred CC on C.cCtaCod=CC.cCtaCod "
    sSql = sSql & "   inner Join Agencias A ON substring(P.cCtaCod,4,2)=A.cAgeCod "
    sSql = sSql & " where substring(P.cCtaCod,4,2) in " & sCadAge
    sSql = sSql & "   and ((P.nPrdEstado >= " & gColocEstVigNorm & " and P.nPrdEstado <= " & gColocEstRefMor & ") Or P.nPrdEstado in (" & gColocEstRecVigJud & ",2205))" 'Vigentes y judiciales
    sSql = sSql & "   and C.dVigencia between '" & Format(pdFecIni, "yyyymmdd") & "' and '" & Format(pdFecFin, "yyyymmdd") & "' "
    sSql = sSql & "        " & sCadProd
    sSql = sSql & " group by A.cAgeDescripcion,RH.cUser with rollup "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaConsolidadoCarteraPorAnalista = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

'MIOL 20120705 ***
Public Function YaRealizoDevBilletajePreCuadre(ByVal sUsuario As String, ByVal dFecha As String, Optional ByVal psCodAge As String = "") As Boolean
Dim sSql As String
Dim oConect As DConecta
Set oConect = New DConecta
Dim Rs As ADODB.Recordset

Set Rs = New ADODB.Recordset
If oConect.AbreConexion = False Then Exit Function

sSql = "Select M.cOpeCod From Mov M " _
    & "Where M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And cOpeCod = '901040'" _
    & "And Right(M.cMovNro, 4) = '" & sUsuario & "' And M.nMovFlag IN (" & gMovFlagVigente & ")"

If psCodAge <> "" Then
    sSql = sSql & " AND Substring(cMovNro, 18,2) = '" & psCodAge & "'"
End If

Set Rs = oConect.CargaRecordSet(sSql)
If Rs.EOF And Rs.BOF Then
    YaRealizoDevBilletajePreCuadre = False
Else
    YaRealizoDevBilletajePreCuadre = True
End If
Rs.Close
Set Rs = Nothing
Set oConect = Nothing
End Function
'***
