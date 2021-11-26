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
Dim sSQL As String
Dim oConecta As DConecta

    sSQL = "Select Pers.cPersCod, Pers.cPersNombre, PID.cPersIDNro DNI,PID2.cPersIDNro RUC, "
    sSQL = sSQL & " Pers.cPersDireccDomicilio, PNat.nPersnatHijos, Pers.dPersNacCreac, PNat.nPersNatEstCiv, "
    sSQL = sSQL & " CN.cConsDescripcion cEstadoCivil, PFI.cPersNombre cRazonSocial, PFI.cPersDireccDomicilio cDireccRazon, "
    sSQL = sSQL & " PFI.cPersCIIU, CI.cCIIUdescripcion, PFI.dPersNacCreac, PersAna.cPersNombre cAnalista, CN2.cConsDescripcion cDestino, "
    sSQL = sSQL & " CN3.cConsDescripcion cCondicion, CE.dPrdEstado dFecAsig, Inst.cPersNombre cNomInst, CV.cCodModular, CE.nMonto nMontoSol,  "
    sSQL = sSQL & " CN4.cConsDescripcion cMonedaSol, CE.nCuotas nCuotasSol, CE.nPlazo nPlazoSol, CN5.cConsDescripcion cTipoCred, CN6.cConsDescripcion cPersoneria, PF.cPersCod cFteInd"
    sSQL = sSQL & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtacod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "                 Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSQL = sSQL & "                 Left Join PersID PID ON PID.cPersCod = PP.cPersCod AND PID.cPersIDTpo = " & gPersIdDNI
    sSQL = sSQL & "                 Left Join PersID PID2 ON PID2.cPersCod = PP.cPersCod AND PID2.cPersIDTpo = " & gPersIdRUC
    sSQL = sSQL & "                 Left Join PersonaNat PNat ON PNat.cPersCod = PP.cPersCod"
    sSQL = sSQL & "                 Left Join Constante CN ON CN.nConsValor = PNat.nPersNatEstCiv AND CN.nConsCod = " & gPersEstadoCivil
    sSQL = sSQL & "                 Left Join ColocFteIngreso CF ON CF.cCtaCod = P.cCtaCod "
    sSQL = sSQL & "                 Left Join PersFteIngreso PF ON PF.cNumFuente = CF.cNumFuente "
    sSQL = sSQL & "                 Left Join Persona PFI ON PFI.cPersCod = PF.cPersCod"
    sSQL = sSQL & "                 Left Join CIIU CI ON PFI.cPersCIIU = CI.cCIIUcod "
    sSQL = sSQL & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAnalista
    sSQL = sSQL & "                 Inner Join Persona PersAna ON PersAna.cPersCod = PP2.cPersCod "
    sSQL = sSQL & "                 Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    sSQL = sSQL & "                 Left Join Constante CN2 ON CN2.nConsValor = CC.nColocDestino AND CN2.nConsCod = " & gColocDestino
    sSQL = sSQL & "                 Left Join Constante CN3 ON CN3.nConsValor = CC.nColocCondicion AND CN3.nConsCod = " & gColocCredCondicion
    sSQL = sSQL & "                 Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic
    sSQL = sSQL & "                 Left Join ColocacConvenio CV ON CV.cCtaCod = P.cCtaCod "
    sSQL = sSQL & "                 Left Join Persona Inst ON Inst.cPersCod = CV.cPersCod "
    sSQL = sSQL & "                 Left Join Constante CN4 ON Convert(int, SubString(P.cCtaCod,9,1)) = CN4.nConsValor AND CN4.nConsCod = " & gMoneda
    sSQL = sSQL & "                 Left Join Constante CN5 ON Convert(int, SubString(P.cCtaCod,6,3)) = CN5.nConsValor AND CN5.nConsCod = " & gProducto
    sSQL = sSQL & "                 Left Join Constante CN6 ON CN6.nConsValor = Pers.nPersPersoneria AND CN6.nConsCod = " & gPersPersoneria
    sSQL = sSQL & "                 Left Join PersFIDependiente PFDep ON PFDep.cNumFuente = CF.cNumFuente AND PFDep.dPersEval = CF.dPersEval "
    sSQL = sSQL & "                 Left Join PersFIIndependiente PFInd ON PFInd.cNumFuente = CF.cNumFuente AND PFInd.dPersEval = CF.dPersEval "
    sSQL = sSQL & " Where P.cCtaCod = '" & psCtaCod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosSolicitudCredito = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaDatosResumenComite(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    sSQL = "Select Pers.cPersCod, Pers.cPersNombre, PID.cPersIDNro DNI,PID2.cPersIDNro RUC, "
    sSQL = sSQL & " Pers.cPersDireccDomicilio, PNat.nPersnatHijos, Pers.dPersNacCreac, PNat.nPersNatEstCiv, "
    sSQL = sSQL & " CN.cConsDescripcion cEstadoCivil, CN2.cConsDescripcion cDestino, "
    sSQL = sSQL & " CN3.cConsDescripcion cCondicion, CN5.cConsDescripcion cTipoCred, L.cDescripcion cLinea, "
    sSQL = sSQL & " CN6.cConsDescripcion cPersoneria "
    sSQL = sSQL & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtacod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "                 Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSQL = sSQL & "                 Left Join PersID PID ON PID.cPersCod = PP.cPersCod AND PID.cPersIDTpo = " & gPersIdDNI
    sSQL = sSQL & "                 Left Join PersID PID2 ON PID2.cPersCod = PP.cPersCod AND PID2.cPersIDTpo = " & gPersIdRUC
    sSQL = sSQL & "                 Left Join PersonaNat PNat ON PNat.cPersCod = PP.cPersCod"
    sSQL = sSQL & "                 Left Join Constante CN ON CN.nConsValor = PNat.nPersNatEstCiv AND CN.nConsCod = " & gPersEstadoCivil
    sSQL = sSQL & "                 Inner Join Colocaciones C ON C.cCtaCod = P.cCtaCod "
    sSQL = sSQL & "                 Left Join ColocLineaCredito L ON C.cLineaCred = L.cLineaCred "
    sSQL = sSQL & "                 Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    sSQL = sSQL & "                 Left Join Constante CN2 ON CN2.nConsValor = CC.nColocDestino AND CN2.nConsCod = " & gColocDestino
    sSQL = sSQL & "                 Left Join Constante CN3 ON CN3.nConsValor = CC.nColocCondicion AND CN3.nConsCod = " & gColocCredCondicion
    sSQL = sSQL & "                 Left Join Constante CN5 ON Convert(int, SubString(P.cCtaCod,6,3)) = CN5.nConsValor AND CN5.nConsCod = " & gProducto
    sSQL = sSQL & "                 Left Join Constante CN6 ON CN6.nConsValor = Pers.nPersPersoneria AND CN6.nConsCod = " & gPersPersoneria
    sSQL = sSQL & " Where P.cCtaCod = '" & psCtaCod & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosResumenComite = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaDatosDocDesembolso(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    sSQL = "Select PersAna.cPersNombre cAnalista, CE.dPrdEstado dFecApr, Pers.cPersCod, "
    sSQL = sSQL & " Pers.cPersNombre, CE.dPrdEstado dFecDesemb, PersApod.cPersNombre cApoderado, CN.cConsDescripcion cMoneda,"
    sSQL = sSQL & " L.cDescripcion cLinea, P.nTasaInteres, CN2.cConsDescripcion cTipoDesemb, CN3.cConsDescripcion cTipoCred "
    sSQL = sSQL & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersAnalista
    sSQL = sSQL & "                Inner Join Persona PersAna ON PersAna.cPersCod = PP.cPersCod "
    sSQL = sSQL & "                Inner Join ColocacEstado CE ON P.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstAprob
    sSQL = sSQL & "                Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "                Inner Join Persona Pers ON Pers.cPersCod = PP2.cPersCod "
    sSQL = sSQL & "                Inner Join ColocacEstado CE2 ON P.cCtaCod = CE2.cCtaCod AND CE2.nPrdEstado = " & gColocEstVigNorm
    sSQL = sSQL & "                Inner Join ProductoPersona PP3 ON P.cCtaCod = PP3.cCtaCod AND PP3.nPrdPersRelac = " & gColRelPersApoderado
    sSQL = sSQL & "                Inner Join Persona PersApod ON PersApod.cPersCod = PP3.cPersCod "
    sSQL = sSQL & "                Inner Join Constante CN ON CN.nConsValor = Convert(Int, Substring(P.cCtaCod,9,1)) AND CN.nConsCod = " & gMoneda
    sSQL = sSQL & "                Inner Join Colocaciones C ON C.cCtacod = P.cCtaCod "
    sSQL = sSQL & "                Inner Join ColocLineaCredito L ON L.cLineaCred = C.cLineaCred "
    sSQL = sSQL & "                Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    sSQL = sSQL & "                Inner Join Constante CN2 ON CN2.nConsValor = CC.nTipoDesembolso AND CN2.nConsCod = " & gColocTiposDesembolso
    sSQL = sSQL & "                Left  Join Constante CN3 ON Convert(int, SubString(P.cCtaCod,6,3)) = CN3.nConsValor AND CN3.nConsCod = " & gProducto
    sSQL = sSQL & " Where P.cCtaCod = '" & psCtaCod & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosDocDesembolso = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaDatosDocPlanPagos(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    sSQL = "Select Pers.cPersNombre, PersAna.cPersNombre cAnalista, CN.cConsDescripcion cTipoCuota, "
    sSQL = sSQL & " nMontoCuota = (Select SUM(CD.nMonto) from ColocCalendDet CD where nNrocalen=(Select nNroCalen From ColocacCred Where cCtaCod = CD.cCtaCod) and cCtaCod = P.cCtaCod AND nColocCalendApl = 1 AND nCuota = 1), "
    sSQL = sSQL & " P.nTasaInteres, CN2.cConsDescripcion cMoneda, CE.nPlazo, CE.nMonto, CE2.dVigencia dFecVig "
    sSQL = sSQL & " From Producto P Inner Join ProductoPersona PP ON PP.cCtaCod = P.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "                INNER Join Colocaciones CE2 ON CE2.cCtaCod = P.cCtaCod "
    sSQL = sSQL & "                Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSQL = sSQL & "                Inner Join ProductoPersona PP2 ON PP2.cCtaCod = P.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAnalista
    sSQL = sSQL & "                Inner Join Persona PersAna ON PersAna.cPersCod = PP2.cPersCod"
    sSQL = sSQL & "                Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado = " & gColocEstAprob
    sSQL = sSQL & "                Inner Join Constante CN ON CN.nConsValor = CE.nColocCalendCod AND CN.nConsCod = " & gColocTipoCalend
    sSQL = sSQL & "                Inner Join Constante CN2 ON CN2.nConsValor = Convert(int,Substring(P.cCtaCod,9,1)) AND CN2.nConsCod = " & gMoneda
    sSQL = sSQL & " Where P.cCtaCod = '" & psCtaCod & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosDocPlanPagos = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaDatosBoletaDesembolso(ByVal pnNroMov As Long, ByVal psCtaCod As String, ByVal psNomAge As String) As ADODB.Recordset
Dim oConecta As DConecta
Dim sSQL As String

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
    
    sSQL = " Select P.cPersNombre,CASE WHEN SUBSTRING(CONVERT(VARCHAR(6),MC.nPrdConceptoCod),1,2) = 12 THEN SUM(MC.nMonto) END as nGasto, CASE WHEN SUBSTRING(CONVERT(VARCHAR(6),MC.nPrdConceptoCod),1,2) = 10 THEN SUM(MC.nMonto) END as nCapital, "
    sSQL = sSQL & "     FechaTran=dbo.FechaMov(M.cMovNro), HoraTran=dbo.HoraMov(M.cMovNro), CN.cConsDescripcion as sMoneda,       MC.nNrocuota, RIGHT(M.cMovNro,4) as sUsuario, Cal.dVenc,COUNT(MC.nMovNro) as nNumTran "
    sSQL = sSQL & " From ProductoPersona PP Inner Join Persona P ON PP.cPersCod = P.cPersCod And PP.nPrdPersRelac = 20 "
    sSQL = sSQL & "       Inner Join MovColDet MC ON MC.cCtaCod = PP.cCtacod "
    sSQL = sSQL & "       Inner Join Mov M ON M.nMovNro = MC.nMovnro "
    sSQL = sSQL & "       Inner Join Constante CN ON CN.nConsCod = " & gMoneda & " And CN.nConsValor = " & Mid(psCtaCod, 9, 1)
    sSQL = sSQL & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = PP.cCtaCod And nColocCalendApl = 1 And nCuota = 1 and Cal.nNroCalen = (Select nNroCalen from ColocacCred where cCtaCod = PP.cCtaCod) "
    sSQL = sSQL & " Where PP.cCtaCod = '" & psCtaCod & "' And MC.nMovNro = " & pnNroMov
    sSQL = sSQL & " Group By P.cPersNombre,  CN.cConsDescripcion, MC.nNrocuota, RIGHT(M.cMovNro,4), Cal.dVenc, MC.cOpeCod,M.cMovNro,SUBSTRING(CONVERT(VARCHAR(6),MC.nPrdConceptoCod),1,2)"
    
        
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosBoletaDesembolso = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function


Public Function RecuperaDatosBoletaGastosDesembolso(ByVal pnNroMov As Long, ByVal psCtaCod As String, ByVal psNomAge As String) As ADODB.Recordset
Dim oConecta As DConecta
Dim sSQL As String

    sSQL = "Select P.cPersNombre,  MC.nMonto, G.cDescripcion, "
    sSQL = sSQL & "       FechaTran=dbo.FechaMov(M.cMovnro), HoraTran=dbo.HoraMov(M.cMovnro), CN.cConsDescripcion as sMoneda, "
    sSQL = sSQL & "       MC.nNrocuota, RIGHT(M.cMovNro,4) as sUsuario, Cal.dVenc"
    sSQL = sSQL & " From ProductoPersona PP Inner Join Persona P ON PP.cPersCod = P.cPersCod And PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "       Inner Join MovColDet MC ON MC.cCtaCod = PP.cCtacod "
    sSQL = sSQL & "       Inner Join Mov M ON M.nMovNro = MC.nMovnro "
    sSQL = sSQL & "       Inner Join Constante CN ON CN.nConsCod = " & gMoneda & " And CN.nConsValor = " & Mid(psCtaCod, 9, 1)
    sSQL = sSQL & "       Inner Join ColocCalendario Cal ON Cal.cCtaCod = PP.cCtaCod And nColocCalendApl = " & gColocCalendAplDesembolso & " And nCuota = 1 and Cal.nNroCalen = (Select MAX(nNroCalen) from Coloccalendario where cCtaCod = PP.cCtaCod)"
    sSQL = sSQL & "       Inner Join ProductoConcepto G ON G.nPrdConceptoCod = MC.nPrdConceptoCod "
    sSQL = sSQL & " Where PP.cCtaCod = '" & psCtaCod & "' And MC.nMovNro = " & pnNroMov & " And SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2) = '12'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosBoletaGastosDesembolso = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaDatosBoletaCancelCredDesembolso(ByVal pnNroMov As Long, ByVal psCtaCod As String, ByVal psNomAge As String) As ADODB.Recordset
Dim oConecta As DConecta

Dim sSQL As String

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
    
    sSQL = "Select MCol.cCtacod as cCtacod,   "
    sSQL = sSQL & " P.cPersNombre,"
    sSQL = sSQL & " nCapital=(Select SUM(MC.nMonto) From MovColdet MC Where MC.cCtaCod = MCol.cCtaCod AND MC.nMovNro = MCol.nMovNro AND MC.nNroCalen = MCol.nNroCalen AND SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2)= '10'),"
    sSQL = sSQL & " ninteres=(Select SUM(MC.nMonto) From MovColdet MC Where MC.cCtaCod = MCol.cCtaCod AND MC.nMovNro = MCol.nMovNro AND MC.nNroCalen = MCol.nNroCalen AND SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2)= '11'),"
    sSQL = sSQL & " nGastos=(Select SUM(MC.nMonto) From MovColdet MC Where MC.cCtaCod = MCol.cCtaCod AND MC.nMovNro = MCol.nMovNro AND MC.nNroCalen = MCol.nNroCalen AND SUBSTRING(CONVERT(varchar(6),MC.nPrdConceptoCod),1,2)= '12'),"
    sSQL = sSQL & " FechaTran=dbo.FechaMov(M.cMovNro), HoraTran=dbo.HoraMov(M.cMovNro),"
    sSQL = sSQL & " CN.cConsDescripcion as sMoneda,"
    sSQL = sSQL & " COUNT(MCol.nMovNro) as nNumTran,"
    sSQL = sSQL & " RIGHT(m.cMovNro, 4) As sUsuario"
    sSQL = sSQL & " From"
    sSQL = sSQL & " MovCol MCol"
    sSQL = sSQL & " Inner Join ProductoPersona PP ON PP.cCtaCod = MCol.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & " Inner Join Persona P ON P.cPersCod = PP.cPersCod"
    sSQL = sSQL & " Inner Join Mov M ON M.nMovNro = MCol.nMovnro"
    sSQL = sSQL & " Inner Join Constante CN ON CN.nConsCod = 1011 And CN.nConsValor = SUBSTRING(MCol.cCtaCod,9,1)"
    sSQL = sSQL & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = PP.cCtaCod And nColocCalendApl = 1  and Cal.nNrocalen = (select MAX(nNrocalen) From ColocCalendario Where cCtaCod = PP.cCtaCod)"
    sSQL = sSQL & " Where MCol.nMovnro = " & pnNroMov & " AND MCol.cCtacod <> '" & psCtaCod & "'"
    sSQL = sSQL & " Group By MCol.cCtaCod, P.cPersNombre,M.cMovNro,CN.cConsDescripcion, MCol.nMovnro,MCol.nNroCalen "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosBoletaCancelCredDesembolso = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function


Public Function ReportePagoDeCreditos(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda) As ADODB.Recordset
Dim sSQL As String
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
    sSQL = " Select * From ("
    sSQL = sSQL & " SELECT"
    sSQL = sSQL & " CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'  ELSE  ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')   END AS sProducto,"
    sSQL = sSQL & "        CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')   END AS sFuenteFinan,"
    sSQL = sSQL & "        CASE WHEN (GROUPING(O.cOpeDesc) = 1) THEN 'TOTAL '        ELSE ISNULL(O.cOpeDesc, 'OPERACION DESCONOCIDA')   END AS cOpeCod,"
    sSQL = sSQL & "        CASE WHEN (GROUPING(MC.cCtaCod) = 1) THEN 'TOTAL'        ELSE ISNULL(MC.cCtaCod, 'DESCONOCIDO')   END AS cCtaCod,"
    sSQL = sSQL & " M.nMovNro,"
    sSQL = sSQL & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = MC.cCtaCod),"
    sSQL = sSQL & " sum(isnull(MCD.nMonto, 0)) as  nCapital,"
    sSQL = sSQL & " sum(isnull(MCD2.nMonto, 0)) as nInteres,"
    sSQL = sSQL & " sum(isnull(MCD4.nMonto, 0)) as nInteresMora,"
    sSQL = sSQL & " sum(isnull(MCD3.nMonto, 0)) as nInteresGracia,"
    sSQL = sSQL & " sum(isnull(MCD5.nMonto, 0)) as nInteresReprog,"
    sSQL = sSQL & " sum(isnull(MCD6.nMonto, 0)) as nInteresSuspenso,"
    sSQL = sSQL & " sum(IsNull(MCD7.nMonto, 0)) As nGastos"
    sSQL = sSQL & " FROM MovCol MC"
    sSQL = sSQL & " Inner Join Constante CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = 1001"
    sSQL = sSQL & " Inner Join OpeTpo O ON O.cOpeCod = MC.cOpeCod"
    sSQL = sSQL & " Inner Join Colocaciones C ON C.cCtaCod = MC.cCtaCod"
    sSQL = sSQL & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = C.cLineaCred"
    sSQL = sSQL & " Inner Join Persona P ON CL.cPersCod = P.cPersCod"
    sSQL = sSQL & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, nPrdConceptoCod, SUM(nMonto) as nMonto"
    sSQL = sSQL & "         From MovColDet Where nPrdConceptoCod = 1000"
    sSQL = sSQL & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD ON MCD.nMovNro = MC.nMovNro And MCD.cOpeCod = MC.cOpeCod And MCD.cCtaCod = MC.cCtaCod And MCD.nNroCalen = MC.nNroCalen And MCD.nPrdConceptoCod = 1000"
    sSQL = sSQL & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, 1100 as nPrdConceptoCod, SUM(nMonto) as nMonto"
    sSQL = sSQL & "         From MovColDet Where nPrdConceptoCod in (1100,1105) "
    sSQL = sSQL & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen) MCD2 ON MCD2.nMovNro = MC.nMovNro And MCD2.cOpeCod = MC.cOpeCod And MCD2.cCtaCod = MC.cCtaCod And MCD2.nNroCalen = MC.nNroCalen And MCD2.nPrdConceptoCod = 1100"
    sSQL = sSQL & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, nPrdConceptoCod, SUM(nMonto) as nMonto"
    sSQL = sSQL & "         From MovColDet Where nPrdConceptoCod = 1102"
    sSQL = sSQL & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD3 ON MCD3.nMovNro = MC.nMovNro And MCD3.cOpeCod = MC.cOpeCod And MCD3.cCtaCod = MC.cCtaCod And MCD3.nNroCalen = MC.nNroCalen And MCD3.nPrdConceptoCod = 1102"
    sSQL = sSQL & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, nPrdConceptoCod, SUM(nMonto) as nMonto"
    sSQL = sSQL & "         From MovColDet Where nPrdConceptoCod = 1101"
    sSQL = sSQL & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD4 ON MCD4.nMovNro = MC.nMovNro And MCD4.cOpeCod = MC.cOpeCod And MCD4.cCtaCod = MC.cCtaCod And MCD4.nNroCalen = MC.nNroCalen And MCD4.nPrdConceptoCod = 1101"
    sSQL = sSQL & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, nPrdConceptoCod, SUM(nMonto) as nMonto"
    sSQL = sSQL & "         From MovColDet Where nPrdConceptoCod = 1103"
    sSQL = sSQL & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD5 ON MCD5.nMovNro = MC.nMovNro And MCD5.cOpeCod = MC.cOpeCod And MCD5.cCtaCod = MC.cCtaCod And MCD5.nNroCalen = MC.nNroCalen And MCD5.nPrdConceptoCod = 1103"
    sSQL = sSQL & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, nPrdConceptoCod, SUM(nMonto) as nMonto"
    sSQL = sSQL & "         From MovColDet Where nPrdConceptoCod = 1104"
    sSQL = sSQL & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen,nPrdConceptoCod) MCD6 ON MCD6.nMovNro = MC.nMovNro And MCD6.cOpeCod = MC.cOpeCod And MCD6.cCtaCod = MC.cCtaCod And MCD6.nNroCalen = MC.nNroCalen And MCD6.nPrdConceptoCod = 1104"
    sSQL = sSQL & " Left Join ( Select nMovNro, cOpeCod, cCtaCod, nNroCalen, Convert(int, '1200') as nPrdConceptoCod, SUM(nMonto) as nMonto"
    sSQL = sSQL & "         From MovColDet Where SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2) = '12'"
    sSQL = sSQL & "         Group By nMovNro, cOpeCod, cCtaCod, nNroCalen, SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2))  MCD7  ON MCD7.nMovNro = MC.nMovNro And MCD7.cOpeCod = MC.cOpeCod And MCD7.cCtaCod = MC.cCtaCod And MCD7.nNroCalen = MC.nNroCalen And MCD7.nPrdConceptoCod = 1200"
    sSQL = sSQL & " Inner Join Mov M ON M.nMovNro = MC.nMovNro And M.nMovFlag = 0  "
    sSQL = sSQL & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'  and Len(MC.cCtaCod) = 18 And (O.cOpeCod >= '100200' and O.cOpeCod <= '100707')"
    sSQL = sSQL & " And SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And SUBSTRING(MC.cCtaCod,4,2) IN " & sCadAge
    sSQL = sSQL & " And SUBSTRING(M.cMovNro,1,8) =  '" & Format(pdFecFin, "yyyymmdd") & "' "
    'sSql = sSql & " And SUBSTRING(M.cMovNro,1,8) Between  '" & Format(pdFecIni, "yyyymmdd") & "' AND '" & Format(pdFecFin, "yyyymmdd") & "' "
    sSQL = sSQL & " GROUP BY CN.cConsDescripcion, P.cPersNombre, O.cOpeDesc, MC.cCtaCod,M.nMovNro With ROLLUP"
    sSQL = sSQL & " ) as T"
    sSQL = sSQL & " Where T.nMovNro is not NULL or  T.cCtaCod = 'TOTAL'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set ReportePagoDeCreditos = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosDesembolsados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda, ByVal pMatProducts As Variant) As ADODB.Recordset
Dim sSQL As String
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
    
    sSQL = "SELECT "
    sSQL = sSQL & "  CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSQL = sSQL & "   END AS sAgencia,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(CN.cAbrev) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(CN.cAbrev, 'PRODUCTO DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sProducto,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sFuenteFinan,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(O.cOpeDesc) = 1) THEN 'TOTAL '"
    sSQL = sSQL & "        ELSE ISNULL(O.cOpeDesc, 'OPERACION DESCONOCIDA')"
    sSQL = sSQL & "   END AS cOpeCod,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(MC.cCtaCod) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(MC.cCtaCod, 'CUENTA DESCONOCIDO ')"
    sSQL = sSQL & "   END AS nCredito,"
    sSQL = sSQL & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = " & gColRelPersTitular & " Where PP.cCtaCod = MC.cCtaCod),"
    'sSql = sSql & " SUM(CE.nMonto) as nMontoApr,"
    sSQL = sSQL & " SUM(isnull(MCD8.nMonto,0)) as nDesembolso,"
    sSQL = sSQL & " SUM(CASE WHEN MCD.nNroCuota >1 THEN 0 ELSE  CAL.nMonto END) as nIntApr, "
    sSQL = sSQL & " SUM(CASE WHEN MCD.nNroCuota >1 THEN 0 ELSE  CE.nMonto END) as nMontoApr, "
    sSQL = sSQL & " SUM(MCD7.nMonto) as Gasto,"
    sSQL = sSQL & " sAnalista = (Select Top 1 cUser From RRHH Where cPersCod = (select cPersCod From ProductoPersona PP2 Where PP2.cCtaCod = MC.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersAnalista & ")),"
    sSQL = sSQL & " nTasa = (Select P.nTasaInteres From Producto P Where P.cCtacod = MC.cCtaCod)"
    sSQL = sSQL & " FROM MovCol MC "
    sSQL = sSQL & " Inner Join ("
    sSQL = sSQL & "     Select nMovNro, cCtaCod, nNroCuota From MovColDet Where  cOpeCod in ('100101','100102','100103','100104','100105') Group by nMovNro, cCtaCod, nNroCuota"
    sSQL = sSQL & " ) MCD ON MCD.nMovNro = MC.nMovNro AND  MCD.cCtaCod = MC.cCtaCod"
    'sSql = sSql & " Inner Join Constante CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSQL = sSQL & " INNER JOIN (Select nConsCod,    nConsValor,  cConsDescripcion, (CONVERT(varchar(10),nConsValor)  + ' ' + cConsDescripcion ) as cAbrev "
    sSQL = sSQL & "    From Constante ) as CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSQL = sSQL & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(MC.cCtaCod,4,2)"
    sSQL = sSQL & " Inner Join OpeTpo O ON O.cOpeCod = MC.cOpeCod"
    sSQL = sSQL & " Inner Join Colocaciones CC ON CC.cCtaCod = MC.cCtaCod"
    sSQL = sSQL & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = CC.cLineaCred"
    sSQL = sSQL & " Inner Join Persona P ON CL.cPersCod = P.cPersCod "
    sSQL = sSQL & " Left Join ( Select nMovNro, cOpeCod, cCtaCod, nNroCalen, Convert(int, '1200') as nPrdConceptoCod, SUM(nMonto) as nMonto From MovColDet"
    sSQL = sSQL & "       Where SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2) = '12'"
    sSQL = sSQL & "       Group By nMovNro, cOpeCod, cCtaCod, nNroCalen, SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2))  MCD7"
    sSQL = sSQL & "    ON MCD7.nMovNro = MC.nMovNro And MCD7.cOpeCod = MC.cOpeCod And MCD7.cCtaCod = MC.cCtaCod And MCD7.nNroCalen = MC.nNroCalen And MCD7.nPrdConceptoCod = 1200 "
    sSQL = sSQL & " Left Join ( Select nMovNro, cOpeCod, cCtaCod, nNroCalen, Convert(int, '1000') as nPrdConceptoCod, SUM(nMonto) as nMonto From MovColDet"
    sSQL = sSQL & "       Where nPrdConceptoCod = 1000 AND (cOpeCod >= '" & gCredDesembEfec & "' and cOpeCod <= '" & gCredDesembCtaNuevaDOA & "') "
    sSQL = sSQL & "       Group By nMovNro, cOpeCod, cCtaCod, nNroCalen, SUBSTRING(CONVERT(Varchar(6),nPrdConceptoCod),1,2))  MCD8"
    sSQL = sSQL & "    ON MCD8.nMovNro = MC.nMovNro And MCD8.cOpeCod = MC.cOpeCod And MCD8.cCtaCod = MC.cCtaCod And MCD8.nNroCalen = MC.nNroCalen And MCD8.nPrdConceptoCod = 1000 "
    sSQL = sSQL & " Inner Join ColocacEstado CE ON CE.cCtaCod = MC.cCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSQL = sSQL & " Inner Join ColocacCred C ON C.cCtaCod = MC.cCtaCod"
    sSQL = sSQL & " Inner Join ( Select CD.cCtaCod, CD.nNrocalen, SUM(CD.nMonto) as nMonto From ColocCalendDet CD"
    sSQL = sSQL & "        Where CD.nColocCalendApl = 1 And CD.nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio
    sSQL = sSQL & "        GROUP BY CD.cCtaCod, CD.nNrocalen) CAL ON CAL.cCtaCod = MC.cCtaCod And CAL.nNroCalen = C.nNroCalen"
    sSQL = sSQL & " Inner Join Mov M ON M.nMovNro = MC.nMovNro "
    sSQL = sSQL & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'  and Len(MC.cCtaCod) = 18 And (MC.cOpeCod >= '" & gCredDesembEfec & "' and MC.cOpeCod <= '" & gCredDesembCtaNuevaDOA & "') And SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And SUBSTRING(MC.cCtaCod,4,2) IN " & sCadAge
    sSQL = sSQL & "     And SUBSTRING(M.cMovNro,1,8) Between  '" & Format(pdFecIni, "yyyymmdd") & "' AND '" & Format(pdFecFin, "yyyymmdd") & "' " & sCadProd
    sSQL = sSQL & " and  nMovFlag <> 2 "
    sSQL = sSQL & " GROUP BY A.cAgeDescripcion, CN.cAbrev, P.cPersNombre, O.cOpeDesc, MC.cCtaCod "
    sSQL = sSQL & " With ROLLUP "
    'sSql = sSql & " ORDER BY  CN.nConsValor "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosDesembolsados = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function


Public Function RecuperaSaldoCarteraVigente(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String)
Dim sSQL As String
Dim oConecta As DConecta
Dim I As Integer
Dim sCadAge As String
   
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"
    
    sSQL = " SELECT"
    sSQL = sSQL & " CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSQL = sSQL & "   END AS sAgencia,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sProducto,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(CN2.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(CN2.cConsDescripcion, 'PLAZO DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sPlazo,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sFuenteFinan,"
    sSQL = sSQL & " COUNT(*) as nCasos,"
    sSQL = sSQL & " SUM(CE.nMonto) as nMontoApr,"
    'sSql = sSql & " SUM(Cal.nMonto) as nDesembolsado,"
    sSQL = sSQL & " SUM(CC.nMontoCol) as nDesembolsado,"
    sSQL = sSQL & " SUM(MC.nSaldo) As nSaldoCapital"
    sSQL = sSQL & " FROM Producto MC Inner Join Constante CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSQL = sSQL & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(MC.cCtaCod,4,2)"
    sSQL = sSQL & " Inner Join Colocaciones CC ON CC.cCtaCod = MC.cCtaCod"
    sSQL = sSQL & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = CC.cLineaCred"
    sSQL = sSQL & " Inner Join Persona P ON CL.cPersCod = P.cPersCod"
    sSQL = sSQL & " Inner join ColocacEstado CE ON CE.cCtaCod = MC.cCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSQL = sSQL & " Inner join ColocacCred C ON C.cCtaCod = MC.cCtaCod"
    'sSql = sSql & " Inner Join ( "
    'sSql = sSql & "  Select cCtaCod, SUM(nMonto) as nMonto from MovColDet"
    'sSql = sSql & "  where cOpecod in ('100101','100102','100103','100104','100105')"
    'sSql = sSql & "  AND nPrdConceptoCod = 1000"
    'sSql = sSql & "  GROUP BY cCtaCod"
    'sSql = sSql & "  ) CAL ON CAL.cCtaCod = MC.cCtaCod "
    sSQL = sSQL & " Inner join Constante CN2 ON CN2.nConsValor = Convert(int,SUBSTRING(CC.cLineaCred,6,1)) And CN2.nConsCod = " & gColocLineaCredPlazo
    sSQL = sSQL & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'  and Len(MC.cCtaCod) = 18  And SUBSTRING(CC.cLineaCred,5,1) = '" & pnMoneda & "' And SUBSTRING(MC.cCtaCod,4,2) IN " & sCadAge
    sSQL = sSQL & " And MC.nPrdEstado in (2020,2021,2022,2030,2031,2032) "
    sSQL = sSQL & " GROUP BY A.cAgeDescripcion, CN.cConsDescripcion, CN2.cConsDescripcion, P.cPersNombre"
    sSQL = sSQL & " With ROLLUP"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaSaldoCarteraVigente = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
Public Function RecuperaCreditosCanceladosTodos(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda) As ADODB.Recordset
Dim sSQL As String
Dim sCadAge As String
Dim I As Integer
Dim oConecta As DConecta

    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sSQL = "SELECT "
    sSQL = sSQL & " CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSQL = sSQL & "   END AS sAgencia,"
    sSQL = sSQL & "  CASE WHEN (GROUPING(CN2.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(CN2.cConsDescripcion, 'PLAZO DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sPlazo,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sProducto,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sFuenteFinan,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(Prd.cCtaCod) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDO ')"
    sSQL = sSQL & "   END AS nCredito,"
    sSQL = sSQL & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = " & gColRelPersTitular & " Where PP.cCtaCod = Prd.cCtaCod),"
    sSQL = sSQL & " SUM(MC.nMonto) as nDesembolso,"
    sSQL = sSQL & " nCuotasApr = (Select nCuotas From ColocacEstado Where cCtaCod = Prd.cCtaCod And nPrdEstado = " & gColocEstAprob & "),"
    sSQL = sSQL & " nPlazoApr = (Select nPlazo From ColocacEstado Where cCtaCod = Prd.cCtaCod And nPrdEstado = " & gColocEstAprob & "),"
    sSQL = sSQL & " dVigencia = (Select dVigencia From Colocaciones Where cCtaCod = Prd.cCtaCod),"
    sSQL = sSQL & " sAnalista = (Select Top 1 cUser From RRHH Where cPersCod = (select cPersCod From ProductoPersona PP2 Where PP2.cCtaCod = Prd.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersAnalista & "))"
    sSQL = sSQL & " FROM Producto Prd Inner Join Constante CN ON Convert(Int,SUBSTRING(Prd.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSQL = sSQL & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSQL = sSQL & " Inner Join Colocaciones CC ON CC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = CC.cLineaCred"
    sSQL = sSQL & " Inner Join Persona P ON CL.cPersCod = P.cPersCod"
    sSQL = sSQL & " Inner join MovCol MC ON MC.cCtaCod = Prd.cCtaCod And (MC.cOpeCod >= '" & gCredDesembEfec & "' and MC.cOpeCod <= '" & gCredDesembCtaNuevaDOA & "')"
    sSQL = sSQL & " Inner join Constante CN2 ON CN2.nConsValor = Convert(int,SUBSTRING(CC.cLineaCred,6,1)) And CN2.nConsCod = " & gColocLineaCredPlazo
    sSQL = sSQL & " Where SUBSTRING(Prd.cCtaCod,6,3)<>'305'  and Len(Prd.cCtaCod) = 18 And Prd.nPrdEstado = " & gColocEstCancelado & " And SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And SUBSTRING(MC.cCtaCod,4,2) IN " & sCadAge
    sSQL = sSQL & "     And Prd.dPrdEstado Between  '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "' "
    sSQL = sSQL & " GROUP BY A.cAgeDescripcion, CN2.cConsDescripcion, CN.cConsDescripcion, P.cPersNombre, Prd.cCtaCod"
    sSQL = sSQL & " With ROLLUP"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosCanceladosTodos = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaCreditosCanceladosConsaldo(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
Dim sCadAge As String
Dim I As Integer
    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sSQL = "SELECT "
    sSQL = sSQL & "  CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSQL = sSQL & "   END AS sAgencia,"
    sSQL = sSQL & "  CASE WHEN (GROUPING(CN2.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(CN2.cConsDescripcion, 'PLAZO DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sPlazo,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sProducto,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sFuenteFinan,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(Prd.cCtaCod) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDO ')"
    sSQL = sSQL & "   END AS nCredito,"
    sSQL = sSQL & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = " & gColRelPersTitular & " Where PP.cCtaCod = Prd.cCtaCod),"
    sSQL = sSQL & " nSaldo=(Select Top 1 nMonto From MovCol Where cCtaCod = Prd.cCtaCod Order by nMovnro DESC),"
    sSQL = sSQL & " SUM(ISNULL(MCD.nMonto,0)) as nCapital,"
    sSQL = sSQL & " SUM(ISNULL(MCD2.nMonto,0)) as nInteres,"
    sSQL = sSQL & " SUM(ISNULL(MCD4.nMonto,0)) as nMora,"
    sSQL = sSQL & " SUM(Cal.nMonto) as nIntPerd,"
    sSQL = sSQL & " cLineaCred = (Select cLineaCred From Colocaciones Where cCtaCod = Prd.cCtaCod),"
    sSQL = sSQL & " sAnalista = (Select Top 1 cUser From RRHH Where cPersCod = (select cPersCod From ProductoPersona PP2 Where PP2.cCtaCod = Prd.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersAnalista & "))"
    sSQL = sSQL & " FROM    Producto Prd"
    sSQL = sSQL & " Inner Join Constante CN ON Convert(Int,SUBSTRING(Prd.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSQL = sSQL & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSQL = sSQL & " Inner Join Colocaciones CC ON CC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = CC.cLineaCred"
    sSQL = sSQL & " Inner Join Persona P ON CL.cPersCod = P.cPersCod"
    sSQL = sSQL & " Inner join MovCol MC ON MC.cCtaCod = Prd.cCtaCod And nMovNro = (Select MAX(nMovNro) From MovCol Where cCtaCod = MC.cCtaCod)"
    sSQL = sSQL & " Left Join ( Select nMovNro, cOpeCod, cCtaCod, nNroCalen, SUM(nMonto) as nMonto"
    sSQL = sSQL & "             From MovColdet"
    sSQL = sSQL & "        Where nPrdConceptoCod = " & gColocConceptoCodCapital
    sSQL = sSQL & "        Group By nMovNro, cOpeCod, cCtaCod, nNroCalen"
    sSQL = sSQL & "         ) as MCD ON MCD.nMovNro = MC.nMovNro And MCD.cOpeCod = MC.cOpeCod And MCD.cCtaCod = MC.cCtaCod And MCD.nNroCalen = MC.nNroCalen"
    sSQL = sSQL & " Left Join (Select nMovNro, cOpeCod, cCtaCod, nNroCalen, SUM(nMonto) as nMonto"
    sSQL = sSQL & "        From MovColdet"
    sSQL = sSQL & "        Where   nPrdConceptoCod in (" & gColocConceptoCodInteresCompensatorio & "," & gColocConceptoCodInteresGracia & gColocConceptoCodInteresReprogramado & gColocConceptoCodInteresSuspenso & ") "
    sSQL = sSQL & "        Group By nMovNro, cOpeCod, cCtaCod, nNroCalen"
    sSQL = sSQL & "         ) as MCD2 ON MCD2.nMovNro = MC.nMovNro And MCD2.cOpeCod = MC.cOpeCod And MCD2.cCtaCod = MC.cCtaCod And MCD2.nNroCalen = MC.nNroCalen"
    sSQL = sSQL & "    Left Join ( Select nMovNro, cOpeCod, cCtaCod, nNroCalen, SUM(nMonto) as nMonto"
    sSQL = sSQL & "        From MovColDet"
    sSQL = sSQL & "        Where nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio
    sSQL = sSQL & "        Group By nMovNro, cOpeCod, cCtaCod, nNroCalen"
    sSQL = sSQL & "         ) as MCD4 ON MCD4.nMovNro = MC.nMovNro And MCD4.cOpeCod = MC.cOpeCod And MCD4.cCtaCod = MC.cCtaCod And MCD4.nNroCalen = MC.nNroCalen"
    sSQL = sSQL & " Inner Join ColocacCred C ON C.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & " Inner Join ( Select CD.cCtaCod, CD.nNrocalen, SUM(CD.nMonto - ISNULL(CD.nMontoPagado,0)) as nMonto From ColocCalendDet CD"
    sSQL = sSQL & "        Where CD.nColocCalendApl = 1 And CD.nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio
    sSQL = sSQL & "        GROUP BY CD.cCtaCod, CD.nNrocalen) CAL ON CAL.cCtaCod = Prd.cCtaCod And CAL.nNroCalen = C.nNroCalen "
    sSQL = sSQL & " Inner join Constante CN2 ON CN2.nConsValor = Convert(int,SUBSTRING(CC.cLineaCred,6,1)) And CN2.nConsCod = " & gColocLineaCredPlazo
    sSQL = sSQL & " Where SUBSTRING(Prd.cCtaCod,6,3)<>'305'  and Len(Prd.cCtaCod) = 18 And Prd.nPrdEstado = " & gColocEstCancelado & " And SUBSTRING(MC.cCtaCod,9,1) = '" & pnMoneda & "' And SUBSTRING(MC.cCtaCod,4,2) IN " & sCadAge
    sSQL = sSQL & "     And Prd.dPrdEstado Between  '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "' "
    sSQL = sSQL & " GROUP BY A.cAgeDescripcion, CN2.cConsDescripcion, CN.cConsDescripcion, P.cPersNombre, Prd.cCtaCod"
    sSQL = sSQL & " With ROLLUP "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosCanceladosConsaldo = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaResumenSaldosCarteraPorAnalista(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal P1I As Integer, ByVal P1F As Integer, _
    ByVal P2I As Integer, ByVal P2F As Integer, ByVal P3I As Integer, ByVal P3F As Integer, ByVal P4F As Integer, _
    ByVal pMatCondicion As Variant, ByVal pMatProd As Variant) As ADODB.Recordset
    
Dim sSQL As String
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
    
    sSQL = "Select  "
    sSQL = sSQL & " CASE WHEN GROUPING(A.cAgeDescripcion)=1 THEN 'TOTAL'"
    sSQL = sSQL & "   ELSE ISNULL(A.cAgeDescripcion,'AGENCIA DESCONOCIDA')"
    sSQL = sSQL & " END  as cAgencia,"
    sSQL = sSQL & " CASE WHEN GROUPING(Q.cUser)=1 THEN 'TOTAL'"
    sSQL = sSQL & "   ELSE ISNULL(Q.cUser,'AGENCIA DESCONOCIDA')"
    sSQL = sSQL & " END as cUser,"
    sSQL = sSQL & " SUM(Q.nCar1Nro) as nCar1Nro, SUM(Q.nCar1Saldo) as nCar1Saldo,"
    sSQL = sSQL & " SUM(Q.nCar2Nro) as nCar2Nro, SUM(Q.nCar2Saldo) as nCar2Saldo,"
    sSQL = sSQL & " SUM(Q.nCar3Nro) as nCar3Nro, SUM(Q.nCar3Saldo) as nCar3Saldo,"
    sSQL = sSQL & " SUM(Q.nCar4Nro) as nCar4Nro, SUM(Q.nCar4Saldo) as nCar4Saldo,"
    sSQL = sSQL & " SUM(Q.nCar5Nro) as nCar5Nro, SUM(Q.nCar5Saldo) as nCar5Saldo"
    sSQL = sSQL & " From ("
    sSQL = sSQL & " select cAgencia, T.cUser + ' ' +  T.cPersNombre as cUser, SUM(T.nCar1Nro) as nCar1Nro, SUM(T.nCar1Saldo) as nCar1Saldo,"
    sSQL = sSQL & "    SUM(T.nCar2Nro) as nCar2Nro, SUM(T.nCar2Saldo) as nCar2Saldo,"
    sSQL = sSQL & "    SUM(T.nCar3Nro) as nCar3Nro, SUM(T.nCar3Saldo) as nCar3Saldo,"
    sSQL = sSQL & "    SUM(T.nCar4Nro) as nCar4Nro, SUM(T.nCar4Saldo) as nCar4Saldo,"
    sSQL = sSQL & "    SUM(T.nCar5Nro) as nCar5Nro, SUM(T.nCar5Saldo) as nCar5Saldo"
    sSQL = sSQL & " From ("
    sSQL = sSQL & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, COUNT(*) as nCar1Nro, SUM(Prd.nSaldo) as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo "
    sSQL = sSQL & "    From RRHH RH "
    sSQL = sSQL & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSQL = sSQL & "        Inner Join RHCargos RC ON RC.cpersCod = RH.cpersCod "
    sSQL = sSQL & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSQL = sSQL & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSQL = sSQL & "        Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & "    Where RC.cRHCargoCod in (" & sAnalistas & ")"
    sSQL = sSQL & "        AND CC.nDiasAtraso >=" & P1I & " AND CC.nDiasAtraso <=" & P1F
    sSQL = sSQL & "        And CC.nColocCondicion in " & sCadCondicion
    sSQL = sSQL & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSQL = sSQL & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSQL = sSQL & "    Union"
    sSQL = sSQL & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, COUNT(*) as nCar2Nro, SUM(Prd.nSaldo) as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo "
    sSQL = sSQL & "    From RRHH RH "
    sSQL = sSQL & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod "
    sSQL = sSQL & "        Inner Join RHCargos RC ON RC.cpersCod = RH.cpersCod "
    sSQL = sSQL & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSQL = sSQL & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSQL = sSQL & "        Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & "    Where RC.cRHCargoCod in (" & sAnalistas & ")"
    sSQL = sSQL & "        AND CC.nDiasAtraso >=" & P2I & " AND CC.nDiasAtraso <=" & P2F
    sSQL = sSQL & "        And CC.nColocCondicion in " & sCadCondicion
    sSQL = sSQL & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSQL = sSQL & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSQL = sSQL & "    Union"
    sSQL = sSQL & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, COUNT(*) as nCar3Nro, SUM(Prd.nSaldo) as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo "
    sSQL = sSQL & "    From RRHH RH "
    sSQL = sSQL & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSQL = sSQL & "        Inner Join RHCargos RC ON RC.cpersCod = RH.cpersCod "
    sSQL = sSQL & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSQL = sSQL & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSQL = sSQL & "        Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & "    Where RC.cRHCargoCod in (" & sAnalistas & ")"
    sSQL = sSQL & "        AND CC.nDiasAtraso >=" & P3I & " AND CC.nDiasAtraso <=" & P3F
    sSQL = sSQL & "        And CC.nColocCondicion in " & sCadCondicion
    sSQL = sSQL & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSQL = sSQL & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSQL = sSQL & "    Union"
    sSQL = sSQL & "    select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, COUNT(*) as nCar4Nro, SUM(Prd.nSaldo) as nCar4Saldo, 0 as nCar5Nro, 0 as nCar5Saldo "
    sSQL = sSQL & "    From RRHH RH"
    sSQL = sSQL & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSQL = sSQL & "        Inner Join RHCargos RC ON RC.cpersCod = RH.cpersCod "
    sSQL = sSQL & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSQL = sSQL & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSQL = sSQL & "        Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & "    Where RC.cRHCargoCod in (" & sAnalistas & ")"
    sSQL = sSQL & "        AND CC.nDiasAtraso > " & P4F
    sSQL = sSQL & "        And CC.nColocCondicion in " & sCadCondicion
    sSQL = sSQL & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSQL = sSQL & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSQL = sSQL & "    Union"
    sSQL = sSQL & "    Select substring(Prd.cCtaCod,4,2) as cAgencia, RH.cUser, P.cPersNombre, 0 as nCar1Nro, 0 as nCar1Saldo, 0 as nCar2Nro, 0 as nCar2Saldo, 0 as nCar3Nro, 0 as nCar3Saldo, 0 as nCar4Nro, 0 as nCar4Saldo , COUNT(*) as nCar5Nro, SUM(Prd.nSaldo) as nCar5Saldo "
    sSQL = sSQL & "    From RRHH RH"
    sSQL = sSQL & "        Inner Join Persona P ON P.cPersCod = RH.cPersCod"
    sSQL = sSQL & "        Inner Join RHCargos RC ON RC.cpersCod = RH.cpersCod "
    sSQL = sSQL & "        Inner Join ProductoPersona PrdPers ON PrdPers.cPersCod = RH.cPersCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSQL = sSQL & "        Inner Join Producto Prd ON Prd.cCtaCod = PrdPers.cCtaCod And Prd.nPrdEstado >= " & gColocEstVigNorm & " and Prd.nPrdEstado <= " & gColocEstRefMor
    sSQL = sSQL & "        Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod " & sCadProd
    sSQL = sSQL & "    Where RC.cRHCargoCod in (" & sAnalistas & ")"
    sSQL = sSQL & "        And CC.nColocCondicion in " & sCadCondicion
    sSQL = sSQL & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "'"
    sSQL = sSQL & "    Group by RH.cUser, P.cPersNombre, substring(Prd.cCtaCod,4,2)"
    sSQL = sSQL & " ) as T"
    sSQL = sSQL & " Group by  T.cUser + ' ' +  T.cPersNombre, cAgencia"
    sSQL = sSQL & " ) As Q Inner Join Agencias A ON A.cAgeCod = Q.cAgencia"
    sSQL = sSQL & " WHERE Q.cAgencia IN " & sCadAge
    sSQL = sSQL & " GROUP BY A.cAgeDescripcion, Q.cUser"
    sSQL = sSQL & " With ROLLUP"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaResumenSaldosCarteraPorAnalista = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaMoraInstitucional(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal P1I As Integer, ByVal P1F As Integer, _
    ByVal P2I As Integer, ByVal P2F As Integer, ByVal P3I As Integer, ByVal P3F As Integer, ByVal P4F As Integer, _
    ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, ByVal pbMoraAnt As Boolean) As ADODB.Recordset

Dim sSQL As String
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


    sSQL = " Select "
    sSQL = sSQL & "    CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSQL = sSQL & "    END AS sAgencia,"
    sSQL = sSQL & "    CASE WHEN GROUPING(T.cRango)=1 THEN 'TOTAL'"
    sSQL = sSQL & " ELSE ISNULL(T.cRango,'RANGO DESCONOCIDO')"
    sSQL = sSQL & " END as cRango,"
    sSQL = sSQL & " CASE WHEN GROUPING(T.cCtaCod)=1 THEN 'TOTAL'"
    sSQL = sSQL & " ELSE ISNULL(T.cCtaCod,'CUENTA DESCONOCIDA')"
    sSQL = sSQL & " END as cCtaCod,"
    sSQL = sSQL & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = T.cCtaCod),"
    sSQL = sSQL & " SUM(T.nMonto) as nMontoApr,"
    sSQL = sSQL & " SUM(T.nMontoCol) as nMontoDesemb,"
    sSQL = sSQL & " SUM(T.nSaldo) as nSaldoCap,"
    'sSql = sSql & " nCuotasApr = (Select nCuotas From ColocacEstado Where cCtaCod = T.cCtaCod And nPrdEstado = 2002),"
    sSQL = sSQL & " nCuotasApr = (Select COUNT(*) From ColocCalendario CC Where CC.cCtaCod = T.cCtaCod AND CC.nNroCalen = (Select nNroCalen from ColocacCred Where cCtaCod = T.cCtaCod) AND CC.nColocCalendApl = 1),"
    sSQL = sSQL & " nCuotaPend = (Select nNroProxCuota From ColocacCred Where cCtaCod = T.cCtaCod),"
    sSQL = sSQL & " SUM(T.nDiasAtraso) as nDiasAtraso,"
    sSQL = sSQL & " sAnalista = (Select Top 1 cUser From RRHH Where cPersCod = (select cPersCod From ProductoPersona PP2 Where PP2.cCtaCod = T.cCtaCod And PP2.nPrdPersRelac = 28)),"
    sSQL = sSQL & " nSaldoCapTotal = (Select SUM(nSaldo) From Producto Where nPrdEstado >= 2020 And nPrdEstado <= 2032 And SUBSTRING(cCtaCod,9,1) = '" & pnMoneda & "')"
    sSQL = sSQL & " From"
    sSQL = sSQL & " ( Select 'RANGO1' as cRango, Prd.cCtaCod, CE.nMonto, C.nMontoCol, Prd.nSaldo, CC.nDiasAtraso"
    sSQL = sSQL & " From  Producto Prd"
    sSQL = sSQL & "     Inner join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSQL = sSQL & "     Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & "     Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor
    sSQL = sSQL & "        AND CC.nDiasAtraso >=" & P1I & " AND CC.nDiasAtraso <=" & P1F
    sSQL = sSQL & "        And CC.nColocCondicion in " & sCadCondicion
    sSQL = sSQL & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select 'RANGO2' as cRango, Prd.cCtaCod, CE.nMonto, C.nMontoCol, Prd.nSaldo, CC.nDiasAtraso"
    sSQL = sSQL & " From  Producto Prd"
    sSQL = sSQL & "     Inner join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSQL = sSQL & "     Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & "     Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor
    sSQL = sSQL & "        AND CC.nDiasAtraso >=" & P2I & " AND CC.nDiasAtraso <=" & P2F
    sSQL = sSQL & "        And CC.nColocCondicion in " & sCadCondicion
    sSQL = sSQL & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select 'RANGO3' as cRango, Prd.cCtaCod, CE.nMonto, C.nMontoCol, Prd.nSaldo, CC.nDiasAtraso"
    sSQL = sSQL & " From  Producto Prd"
    sSQL = sSQL & "     Inner join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSQL = sSQL & "     Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & "     Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor
    sSQL = sSQL & "        AND CC.nDiasAtraso >=" & P3I & " AND CC.nDiasAtraso <=" & P3F
    sSQL = sSQL & "        And CC.nColocCondicion in " & sCadCondicion
    sSQL = sSQL & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select 'RANGO4' as cRango, Prd.cCtaCod, CE.nMonto, C.nMontoCol, Prd.nSaldo, CC.nDiasAtraso"
    sSQL = sSQL & " From  Producto Prd"
    sSQL = sSQL & "     Inner join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSQL = sSQL & "     Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & "     Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor
    sSQL = sSQL & "        AND CC.nDiasAtraso >=" & P4F
    sSQL = sSQL & "        And CC.nColocCondicion in " & sCadCondicion
    sSQL = sSQL & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd
    sSQL = sSQL & " ) as T Inner join Agencias A ON A.cAgeCod = SUBSTRING(T.cCtaCod,4,2)"
    sSQL = sSQL & " WHERE A.cAgeCod IN " & sCadAge
    sSQL = sSQL & " Group by A.cAgeDescripcion, T.cRango, T.cCtaCod"
    sSQL = sSQL & " With ROLLUP "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaMoraInstitucional = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaMoraxAnalista_AtrasoPagoCuotaLibre(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, _
    ByVal pMatAnalistas As Variant, ByVal pnConCuotaLibre As Boolean) As ADODB.Recordset
Dim sSQL As String
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




 sSQL = " Select CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
 sSQL = sSQL & "       ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')    END AS cAgencia,"
 sSQL = sSQL & "   CASE WHEN GROUPING(P2.cPersNombre)=1 THEN 'TOTAL'"
 sSQL = sSQL & "       ELSE ISNULL(P2.cPersNombre,'PERSONA DESCONOCIDA') END as cAnalista,"
 sSQL = sSQL & "   CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL'"
 sSQL = sSQL & "       ELSE ISNULL(Prd.cCtaCod,'CUENTA DESCONOCIDA') END as cCtaCod,"
 sSQL = sSQL & "   dbo.ColocCredTitular(Prd.cCtaCod) + ' ' + RIGHT(SPACE(10) + CONVERT(varchar(20),C.nMontoCol),10) +  ' ' + RIGHT(SPACE(10) + CONVERT(varchar(20),Prd.nSaldo),10) +  ' ' + RIGHT(SPACE(10) + CONVERT(varchar(15),dbo.CuotaAprobada(Prd.cCtaCod)),10) + '     ' + CONVERT(varchar(10),CCal.nCuota) + '    ' +  CONVERT(varchar(15),CCal.dvenc,103) as nDatos,"
 sSQL = sSQL & "   SUM(DATEDIFF(day,CCal.dVenc,'" & Format(pdFecIni, "mm/dd/yyyy") & "')) as nDiasAtraso,"
 sSQL = sSQL & "   SUM(dbo.ColocMoraCuota(Prd.cCtaCod,CCal.nCuota)) as nMora,"
 sSQL = sSQL & "   Sum (dbo.ColocGastoCuota(Prd.cCtaCod, CCal.nCuota)) as nGasto, "
 sSQL = sSQL & "   Sum (dbo.ColocSaldoCuota(Prd.cCtaCod, CCal.nCuota)) as nSaldoCuota "
 sSQL = sSQL & "From Producto Prd"
 sSQL = sSQL & "   Inner Join ProductoPersona PP ON PP.cCtaCod = Prd.cCtaCod And PP.nPrdPersRelac=28 And PP.cPersCod in ('1020100517046','1020100557056','1020100812412','1020100213047','1020100761231','1020100755362')"
 sSQL = sSQL & "   Inner Join Persona P2 ON P2.cPersCod = PP.cPersCod"
 sSQL = sSQL & "   Inner join Agencias A ON A.cAgeCod = SUBSTRING(PRD.cCtaCod,4,2)"
 sSQL = sSQL & "   Inner Join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
 sSQL = sSQL & "   Inner join ColocacCred CC ON CC.cCtaCod =  Prd.cCtaCod"
 sSQL = sSQL & "   Inner Join ColocCalendario CCal ON CCal.cCtaCod = Prd.cCtaCod And CCal.nNroCalen = CC.nNroCalen"
 sSQL = sSQL & "       And CCal.nColocCalendApl = 1 And CCal.dVenc <'08/01/2003' And CCal.nColocCalendEstado=0"
 sSQL = sSQL & " WHERE (Prd.nPrdEstado IN (" & gColocEstRefNorm & ", " & gColocEstRefVenc & ", " & gColocEstRefMor & ", " & gColocEstVigNorm & ", " & gColocEstVigVenc & ", " & gColocEstVigMor & ")) "

 sSQL = sSQL & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd & " And A.cAgeCod IN " & sCadAge
 sSQL = sSQL & "        And CC.nColocCondicion in " & sCadCondicion

    If pnConCuotaLibre = True Then
        sSQL = sSQL & " And CC.nColocCalendCod=" & gColocCalendCodCL & " "
    End If

 sSQL = sSQL & "  Group by A.cAgeDescripcion,P2.cPersNombre,Prd.cCtaCod,"
 sSQL = sSQL & " (dbo.ColocCredTitular(Prd.cCtaCod) + ' ' + RIGHT(SPACE(10) + CONVERT(varchar(20),C.nMontoCol),10) +  ' ' + RIGHT(SPACE(10) + CONVERT(varchar(20),Prd.nSaldo),10) +  ' ' + RIGHT(SPACE(10) + CONVERT(varchar(15),dbo.CuotaAprobada(Prd.cCtaCod)),10) + '     ' + CONVERT(varchar(10),CCal.nCuota) + '    ' +  CONVERT(varchar(15),CCal.dvenc,103))"
 sSQL = sSQL & " With ROLLUP"


    If pnConCuotaLibre = True Then
        'sSql = sSql & " CC.nColocCalendCod, "
    End If


    If pnConCuotaLibre = True Then
        'sSql = sSql & " And CC.nColocCalendCod=" & gColocCalendCodCL & " "
    End If

    If pnConCuotaLibre = True Then
        'sSql = sSql & ", CC.nColocCalendCod "
    End If

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaMoraxAnalista_AtrasoPagoCuotaLibre = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosProtestados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta
Dim sCadAge As String
Dim I As Integer

    sCadAge = "('"
    For I = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(I) & "','"
    Next I
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 2) & ")"

    sSQL = "Select A.cAgeDescripcion, CN.cConsDescripcion, Prd.cCtaCod, Pers.cPersNombre,"
    sSQL = sSQL & " Prd.nSaldo, SUBSTRING(M.cMovNro,1,8) as sFecha, CC.nNroProxCuota, CD.dVenc,"
    sSQL = sSQL & " CDet.nMonto as nCuotaApr, CDet.nMonto - CDet.nMontoPagado as nCuotaXAmort, CDet.nMontoPagado as nCuotaAmort,"
    sSQL = sSQL & " CC.nDiasAtraso"
    sSQL = sSQL & " From     Producto Prd"
    sSQL = sSQL & "  Inner Join Constante CN ON Convert(Int,SUBSTRING(Prd.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSQL = sSQL & "  Inner Join ProductoPersona PP ON PP.cCtaCod = Prd.cCtaCod And PP.nPrdPersrelac = " & gColRelPersTitular
    sSQL = sSQL & "  Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSQL = sSQL & "  LEFT Join (Select cCtaCod, MAX(nMovNro) as nMovNro  From MovCol Where SUBSTRING(cCtaCod,6,3) <> '305'  AND cOpeCod like '100[234567]%' Group by cCtaCod) as MC ON MC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & "  LEFT Join Mov M ON M.nMovnro = MC.nMovNro"
    sSQL = sSQL & "  Inner join ColocacCred CC ON CC.cCtacod = Prd.cCtaCod"
    sSQL = sSQL & "  Inner Join ColocCalendario CD ON CD.cCtaCod = Prd.cCtaCod And CD.nNroCalen = CC.nNroCalen And CD.nCuota=CC.nNroProxCuota And CD.nColocCalendApl = " & gColocCalendAplCuota
    sSQL = sSQL & "  Inner Join (Select cCtaCod, nNroCalen, nCuota, SUM(nMonto) as nMonto, SUM(nMontoPagado) as nMontoPagado From ColocCalendDet Where nColocCalendApl =" & gColocCalendAplCuota & " Group by cCtaCod, nNroCalen, nCuota) as CDet"
    sSQL = sSQL & "     On CDet.cCtaCod = Prd.cCtaCod And CDet.nNroCalen = CC.nNroCalen And CDet.nCuota = CC.nNroProxCuota"
    sSQL = sSQL & "  Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSQL = sSQL & " WHERE CC.cProtesto=1  and SUBSTRING(Prd.cCtaCod,6,3) in ('101','102','103','104','201','202','203','221','301','302','303','304','320','401','423')"
    sSQL = sSQL & "  And LEN(Prd.cCtacod)=18"
    sSQL = sSQL & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & " And A.cAgeCod IN " & sCadAge
    sSQL = sSQL & "        And Prd.nPrdEstado in (2020,2021,2022,2030,2031,2032) "
    sSQL = sSQL & " Order by A.cAgeDescripcion, CN.cConsDescripcion, Pers.cPersNombre "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosProtestados = oConecta.CargaRecordSet(sSQL)
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
Dim sSQL As String

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

    sSQL = " Select A.cAgeDescripcion, Prd.cCtaCod, P.cPersNombre, CE2.nMonto, CN.cConsDescripcion, CE2.dPrdEstado, RH.cUser"
    sSQL = sSQL & " From Producto Prd"
    sSQL = sSQL & " Inner join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & " Inner Join ProductoPersona PP ON PP.cCtaCod = Prd.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & " Inner Join Persona P ON P.cPersCod = PP.cPersCod"
    sSQL = sSQL & " Inner Join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = " & gColocEstRetirado
    sSQL = sSQL & " Inner Join ColocacEstado CE2 ON CE2.cCtaCod = Prd.cCtaCod And CE2.nPrdEstado = " & gColocEstAprob
    sSQL = sSQL & " Inner join Constante CN ON CN.nConsValor = CE.nMotivoRechazo And CN.nConsCod = " & gColocMotivRechazo
    sSQL = sSQL & " Inner Join ProductoPersona PrdPers ON PrdPers.cCtaCod = Prd.cCtaCod And PrdPers.nPrdPersRelac = " & gColRelPersAnalista
    sSQL = sSQL & " Inner join RRHH RH ON RH.cPersCod = PrdPers.cPersCod"
    sSQL = sSQL & " Inner Join Agencias A ON A.cAgeCod = substring(Prd.cCtaCod,4,2)"
    sSQL = sSQL & " Where Prd.nPrdEstado = " & gColocEstRetirado & " and Prd.dPrdEstado >= '" & Format(pdFecIni, "mm/dd/yyyy") & "' And Prd.dPrdEstado <= '" & Format(pdFecFin, "mm/dd/yyyy") & "' "
    sSQL = sSQL & "        And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd & " And A.cAgeCod IN " & sCadAge
    sSQL = sSQL & "        And CC.nColocCondicion in " & sCadCondicion
    sSQL = sSQL & " Order by A.cAgeDescripcion, P.cpersNombre "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosRetirados = oConecta.CargaRecordSet(sSQL)
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
Dim sSQL As String
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

    sSQL = "select A.cAgeDescripcion, Pers.cPersNombre, Pers.cPersDireccDomicilio, U.cUbiGeoDescripcion as Zona, Prd.cCtaCod, Cal.nCuota, Pers2.cPersNombre as cAnalista "
    sSQL = sSQL & " From Producto Prd"
    sSQL = sSQL & " Inner Join ProductoPersona PP ON PP.cCtaCod = Prd.cCtaCod And PP.nPrdPersRelac = 20"
    sSQL = sSQL & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSQL = sSQL & " Left Join UbicacionGeografica U ON U.cUbiGeoCod = Pers.cPersDireccUbigeo"
    sSQL = sSQL & " Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & " Inner Join (Select cCtaCod, nNroCalen, Max(nCuota) as nCuota From ColocCalendario"
    sSQL = sSQL & "         Where nColocCalendEstado = 0 And nColocCalendApl = 1 "
    sSQL = sSQL & "          Group By cCtaCod, nNroCalen"
    sSQL = sSQL & "        ) as Cal ON Cal.cCtaCod = Prd.cCtaCod And Cal.nNroCalen = CC.nNroCalen"
    sSQL = sSQL & " Inner Join ProductoPersona PP2 ON PP2.cCtaCod = Prd.cCtaCod And PP2.nPrdPersRelac = 28 And PP2.cPersCod in " & sCadAnalista
    sSQL = sSQL & " Inner Join Persona Pers2 ON Pers2.cPersCod = PP2.cPersCod"
    sSQL = sSQL & " Inner Join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSQL = sSQL & " Where SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd & " And A.cAgeCod IN " & sCadAge
    sSQL = sSQL & "        And CC.nColocCondicion in " & sCadCondicion
    sSQL = sSQL & "        And CC.nDiasAtraso >= " & pnDiasAtrIni & " And CC.nDiasAtraso <= " & pnDiasAtrFin
    sSQL = sSQL & " ORDER BY A.cAgeDescripcion, Pers.cPersNombre"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCartaCreditosMorosos = oConecta.CargaRecordSet(sSQL)
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
Dim sSQL As String
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

    sSQL = "select A.cAgeDescripcion, Pers.cPersNombre, Pers.cPersDireccDomicilio, U.cUbiGeoDescripcion as Zona, Prd.cCtaCod, Cal.nCuotas, Pers2.cPersNombre as cAnalista "
    sSQL = sSQL & " From Producto Prd"
    sSQL = sSQL & " Inner Join ProductoPersona PP ON PP.cCtaCod = Prd.cCtaCod And PP.nPrdPersRelac = 20"
    sSQL = sSQL & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSQL = sSQL & " Left Join UbicacionGeografica U ON U.cUbiGeoCod = Pers.cPersDireccUbigeo"
    sSQL = sSQL & " Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & " Inner Join (Select cCtaCod, nNroCalen, COUNT(nCuota) as nCuotas From ColocCalendario"
    sSQL = sSQL & "         Where nColocCalendEstado = 0 And nColocCalendApl = 1 "
    sSQL = sSQL & "          Group By cCtaCod, nNroCalen"
    sSQL = sSQL & "        ) as Cal ON Cal.cCtaCod = Prd.cCtaCod And Cal.nNroCalen = CC.nNroCalen"
    sSQL = sSQL & " Inner Join ProductoPersona PP2 ON PP2.cCtaCod = Prd.cCtaCod And PP2.nPrdPersRelac = 28 And PP2.cPersCod in " & sCadAnalista
    sSQL = sSQL & " Inner Join Persona Pers2 ON Pers2.cPersCod = PP2.cPersCod"
    sSQL = sSQL & " Inner Join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSQL = sSQL & " Inner Join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSQL = sSQL & " Left Join ( Select cCtaCod, nColocNota From ColocCalificacionAnalista CA Where dColocNotaFecha = (Select MAX(dColocNotaFecha) From ColocCalificacionAnalista Where cCtaCod = CA.cCtaCod)) as CA "
    sSQL = sSQL & " ON CA.cCtaCod = Prd.cCtaCod "
    sSQL = sSQL & " Where SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd & " And A.cAgeCod IN " & sCadAge
    sSQL = sSQL & "        And CC.nColocCondicion in " & sCadCondicion
    sSQL = sSQL & " And CA.nColocNota  >= " & pnNotaIni & " And CA.nColocNota  <= " & pnNotaFin
    If pbPorc <> 1 Then
        sSQL = sSQL & "    And Cal.nCuotas = " & pnUltCuotas
    Else
        sSQL = sSQL & "    And Prd.nSaldo <= CE.nMonto * " & Format(pnUltCuotas / 100, "#0.00")
    End If
    sSQL = sSQL & " ORDER BY A.cAgeDescripcion, Pers.cPersNombre"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCartaInvitacionCreditosParalelos = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosPorUbicacionGeo(ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, _
    ByVal psCodUbic As String) As ADODB.Recordset
Dim sSQL As String
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
    
    sSQL = "select A.cAgeDescripcion, Pers.cPersNombre, Pers.cPersDireccDomicilio, U.cUbiGeoDescripcion as Zona,"
    sSQL = sSQL & " Prd.cCtaCod, C.nMontoCol nMonto, Prd.nSaldo, CA.nColocNota, CC.nNroProxCuota, "
    sSQL = sSQL & " nCuotas = (SELECT Count(*) FROM ColocCalendario WHERE cCtaCod = C.cCtaCod AND nNroCalen = CC.nNroCalen AND nColocCalendApl = 1), "
    sSQL = sSQL & " RH.cUser From Producto Prd "
    sSQL = sSQL & " Inner Join ProductoPersona PP ON PP.cCtaCod = Prd.cCtaCod And PP.nPrdPersRelac=" & gColRelPersTitular & " "
    sSQL = sSQL & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod"
    sSQL = sSQL & " Inner Join UbicacionGeografica U ON U.cUbiGeoCod = Pers.cPersDireccUbigeo And U.cUbiGeoCod = '" & Trim(Right(psCodUbic, 15)) & "'"
    sSQL = sSQL & " Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod INNER JOIN COLOCACIONES C ON CC.cCtaCod = C.cCtaCod "
    sSQL = sSQL & " Left Join (Select cCtaCod, nColocNota From ColocCalificacionAnalista CAN Where dColocNotaFecha = (Select MAX(dColocNotaFecha) From ColocCalificacionAnalista Where cCtaCod = CAN.cCtaCod)) as CA"
    sSQL = sSQL & "     ON CA.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & " Inner Join ProductoPersona PP2 ON PP2.cCtaCod = Prd.cCtaCod And PP2.nPrdPersRelac=" & gColRelPersAnalista & " "
    sSQL = sSQL & " Inner Join RRHH RH ON RH.cPersCod = PP2.cPersCod"
    sSQL = sSQL & " Inner Join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSQL = sSQL & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor
    sSQL = sSQL & " And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd & " And A.cAgeCod IN " & sCadAge
    sSQL = sSQL & " And CC.nColocCondicion in " & sCadCondicion & " ORDER BY PP.cCtaCod "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosPorUbicacionGeo = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCredVig_CredVigConCuotLib(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, ByVal pnDiasAtrIni As Integer, _
    ByVal pnDiasAtrFin As Integer, ByVal pnTipoCambio As Double, ByVal pnConCuotaLibre As Boolean) As ADODB.Recordset
    
Dim sSQL As String
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
    
    sSQL = " SELECT "
    sSQL = sSQL & "  CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSQL = sSQL & "   END AS sAgencia,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sProducto,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(CN2.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(CN2.cConsDescripcion, 'PLAZO DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sPlazo,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sFuenteFinan,"
    sSQL = sSQL & "   CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL'"
    sSQL = sSQL & "    ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDA')"
    sSQL = sSQL & "   END AS cCtaCod,"
    sSQL = sSQL & " COUNT(*) as nCasos,"
    
    'If pnConCuotaLibre = True Then
    '    sSql = sSql & " cc.ncoloccalendcod, "
    'End If
    
    sSQL = sSQL & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac=" & gColRelPersTitular & " Where PP.cCtaCod = Prd.cCtaCod)," 'Decia 20
    sSQL = sSQL & " sInstitucion = (Select cPersNombre From Persona Where cPersCod = (Select Top 1 cPersCod From ColocacConvenio Where cCtaCod = Prd.cCtaCod)),"
    sSQL = sSQL & " dVigencia = (select dVigencia From Colocaciones Where cCtaCod = Prd.cCtaCod),"
    sSQL = sSQL & " sPlazoCuota = (Select CONVERT(varchar(6),nCuotas) + '/' +  CONVERT(varchar(6), Case nPlazo When 0 Then 30 Else nPlazo End) "
    sSQL = sSQL & " From ColocacEstado Where cCtaCod = Prd.cCtaCod And nPrdEstado = " & gColocEstAprob & "),"
    sSQL = sSQL & " SUM(isnull(CE.nMonto,0)) as nMontoApr,"
    sSQL = sSQL & " SUM(isnull(Prd.nSaldo,0)) as nSaldoCap,"
    sSQL = sSQL & " nCuotaMora = (select nNroProxCuota From ColocacCred Where cCtaCod = Prd.cCtaCod),"
    sSQL = sSQL & " nDiasAtraso = (select nDiasAtraso From ColocacCred Where cCtaCod = Prd.cCtaCod),"
    sSQL = sSQL & " nMontoCuota = (Select SUM(nMonto) from ColocCalendDet"
    sSQL = sSQL & "        Where cCtaCod = Prd.cCtaCod And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = Prd.cCtaCod)"
    sSQL = sSQL & "        And nCuota = (select nNroProxCuota From ColocacCred Where cCtaCod = Prd.cCtaCod)"
    sSQL = sSQL & "        And nColocCalendApl=" & gColocCalendAplCuota & "), " 'Decia 1
    sSQL = sSQL & " dFecUtlPago = (Select dbo.FechaMov(M.cMovnro) From Mov M Where nMovNro = (Select MAX(nMovNro) From MovCol where cCtaCod = Prd.cCtaCod)),"
    sSQL = sSQL & " dbo.InteresDevengado(prd.cCtaCod,'" & Format(pdFecIni, "mm/dd/yyyy") & "' ) as IntDeveng,"
    sSQL = sSQL & " dbo.DeudaTotal(Prd.cCtaCod) as nDeudaTotal,"
    sSQL = sSQL & " dbo.CapitalVencido(Prd.cCtaCod,'" & Format(pdFecIni, "mm/dd/yyyy") & "') as nCapVenc,"
    sSQL = sSQL & " dbo.CapitalXVencer(Prd.cCtaCod,'" & Format(pdFecIni, "mm/dd/yyyy") & "') as nCapXVencer,"
    sSQL = sSQL & " dbo.GarantiaHipotecaria(Prd.cCtaCod," & Format(pnTipoCambio, "#0.00") & ") as nMontoGarHip,"
    sSQL = sSQL & " dbo.OtrasGarantias(Prd.cCtaCod," & Format(pnTipoCambio, "#0.00") & ") as nMontoOtrasGarant,"
    sSQL = sSQL & " cEstadoDesc = (Select CN2.cConsDescripcion From Producto Prd2 Inner Join Constante CN2 ON CN2.nConsValor = Prd2.nPrdEstado And CN2.nConsCod = " & gColocEstado & " Where Prd2.cCtaCod = Prd.cCtaCod)"
    sSQL = sSQL & " FROM Producto Prd"
    sSQL = sSQL & " Inner Join Constante CN ON Convert(Int,SUBSTRING(Prd.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSQL = sSQL & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSQL = sSQL & " Inner Join Colocaciones C ON C.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & " Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod And CC.nDiasAtraso >= " & pnDiasAtrIni & " And CC.nDiasAtraso <=" & pnDiasAtrFin
    sSQL = sSQL & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = C.cLineaCred"
    sSQL = sSQL & " Inner join Constante CN2 ON CN2.nConsValor = Convert(int,SUBSTRING(C.cLineaCred,6,1)) And CN2.nConsCod = " & gColocLineaCredPlazo
    sSQL = sSQL & " Inner Join Persona P ON CL.cPersCod = P.cPersCod"
    sSQL = sSQL & " Inner Join ColocacEstado CE ON CE.cCtaCod = Prd.CCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSQL = sSQL & " WHERE (Prd.nPrdEstado IN (" & gColocEstRefNorm & ", " & gColocEstRefVenc & ", " & gColocEstRefMor & ", " & gColocEstVigNorm & ", " & gColocEstVigVenc & ", " & gColocEstVigMor & ")) "
    sSQL = sSQL & " And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' " & sCadProd & " And A.cAgeCod IN " & sCadAge
    sSQL = sSQL & "        And CC.nColocCondicion in " & sCadCondicion
    
    If pnConCuotaLibre = True Then
         sSQL = sSQL & " And CC.nColocCalendCod=" & gColocCalendCodCL & " "
    End If
    
    sSQL = sSQL & " GROUP BY A.cAgeDescripcion, CN.cConsDescripcion, CN2.cConsDescripcion, P.cPersNombre, Prd.cCtaCod"
    
    sSQL = sSQL & " WITH ROLLUP"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCredVig_CredVigConCuotLib = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaCreditosXInstitucion(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, _
    ByVal pdFecIni As Date, ByVal pnTipoOrden As Integer, pnIncluyeMora As Integer) As ADODB.Recordset
    
Dim sSQL As String
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
    
    sSQL = " SELECT "
    sSQL = sSQL & "  CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSQL = sSQL & "   END AS sAgencia,"
    sSQL = sSQL & "   CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL'"
    sSQL = sSQL & "    ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDA')"
    sSQL = sSQL & "   END AS cCtaCod,"
    sSQL = sSQL & "   CASE WHEN GROUPING(P.cPersNombre)=1 THEN 'TOTAL'"
    sSQL = sSQL & "    ELSE ISNULL(P.cPersNombre, 'INSTITUCION DESCONOCIDA')"
    sSQL = sSQL & "   END AS cInstitucion,"
    sSQL = sSQL & " COUNT(*) as nCasos, "
    sSQL = sSQL & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = Prd.cCtaCod),"
    sSQL = sSQL & " cCodModular = (Select cCodModular From ColocacConvenio CV Where CV.cCtaCod = Prd.cCtaCod),"
    sSQL = sSQL & " dFecDesemb = (Select dVigencia From Colocaciones Where cCtaCod = Prd.cCtaCod),"
    sSQL = sSQL & " nDiasAtraso = (Select nDiasAtraso From ColocacCred Where cCtaCod = Prd.cCtaCod), "
    sSQL = sSQL & " SUM(CE.nMonto) as nMontoApr,"
    sSQL = sSQL & " SUM(Prd.nSaldo) as nSaldoCap,"
    sSQL = sSQL & " sCuotaPend = (Select CONVERT(varchar(5),CC.nNroProxCuota) + '/' + CONVERT(varchar(5),CE.nCuotas) From ColocacCred CC Inner Join ColocacEstado CE ON CC.cCtaCod = CE.cCtaCod And CE.nPrdEstado = 2002"
    sSQL = sSQL & "        Where CC.cCtaCod = Prd.cCtaCod ),"
    sSQL = sSQL & " SUM(CDet.nMonto) as nCuota,"
    sSQL = sSQL & " dFecPago = (Select dVenc From ColocCalendario CCal Where CCal.cCtaCod = Prd.cCtaCod And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CCal.cCtaCod) And nColocCalendApl = 1 And nCuota = (Select nNroProxCuota From ColocacCred Where cCtaCod = CCal.cCtaCod))"
    sSQL = sSQL & " From Producto Prd"
    sSQL = sSQL & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSQL = sSQL & " Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & " Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSQL = sSQL & " Inner Join ColocacEstado CE ON CE.cCtaCod = Prd.cCtaCod And CE.nPrdEstado = 2002"
    sSQL = sSQL & " Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & " Inner Join (Select cCtaCod, nNroCalen, nCuota,  SUM(nMonto) as nMonto From ColocCalendDet CDet"
    sSQL = sSQL & "        Where nColocCalendApl = 1 " & sCadMora
    sSQL = sSQL & "        Group By cCtaCod, nNroCalen, nCuota"
    sSQL = sSQL & "        ) as CDet ON CDet.cCtaCod = Prd.cCtaCod And CDet.nNroCalen = CC.nNroCalen And CDet.nCuota = CC.nNroProxCuota"
    sSQL = sSQL & " Where Prd.nPrdEstado >= 2020 And Prd.nPrdEstado <= 2032"
    sSQL = sSQL & " And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' And A.cAgeCod IN " & sCadAge
    sSQL = sSQL & " Group By A.cAgeDescripcion, P.cPersNombre, Prd.cCtaCod"
    sSQL = sSQL & " WITH ROLLUP"
    sSQL = sSQL & " Order by sAgencia, cCtaCod, cInstitucion "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosXInstitucion = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaMoraXInstituciones(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, _
    ByVal pdFecIni As Date, pMatInstitucion As Variant) As ADODB.Recordset
    
Dim sSQL As String
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
    
    sSQL = " SELECT "
    sSQL = sSQL & "  CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSQL = sSQL & "   END AS sAgencia,"
    sSQL = sSQL & "   CASE WHEN GROUPING(P.cPersNombre)=1 THEN 'TOTAL'"
    sSQL = sSQL & "    ELSE ISNULL(P.cPersNombre, 'INSTITUCION DESCONOCIDA')"
    sSQL = sSQL & "   END AS cInstitucion,"
    sSQL = sSQL & "   CASE WHEN GROUPING(Prd.cCtaCod)=1 THEN 'TOTAL'"
    sSQL = sSQL & "    ELSE ISNULL(Prd.cCtaCod, 'CUENTA DESCONOCIDA')"
    sSQL = sSQL & "   END AS cCtaCod,"
    sSQL = sSQL & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = Prd.cCtaCod),"
    'sSql = sSql & " nMontoDesemb = (Select nMontoCol From Colocaciones Where cCtaCod = Prd.cCtaCod),"
    sSQL = sSQL & " SUM(C.nMontoCol) as nMontoDesemb, "
    sSQL = sSQL & " SUM(Prd.nSaldo) as nSaldoCap,"
    sSQL = sSQL & " sCuotaPend = (Select CONVERT(varchar(5),CC.nNroProxCuota) + '/' + CONVERT(varchar(5),CE.nCuotas) From ColocacCred CC Inner Join ColocacEstado CE ON CC.cCtaCod = CE.cCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSQL = sSQL & "        Where CC.cCtaCod = Prd.cCtaCod ),"
    sSQL = sSQL & " SUM(CDet.nMonto) as nSaldoCuota,"
    sSQL = sSQL & " nDiasAtraso = (Select nDiasAtraso From ColocacCred Where cCtaCod = Prd.cCtaCod),"
    sSQL = sSQL & " SUM(CDet2.nMonto) as nSaldoMora,"
    sSQL = sSQL & " dFecVencimiento = (Select dVenc From ColocCalendario CCal Where CCal.cCtaCod = Prd.cCtaCod And nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CCal.cCtaCod) And nColocCalendApl = 1 And nCuota = (Select nNroProxCuota From ColocacCred Where cCtaCod = CCal.cCtaCod)),"
    sSQL = sSQL & " COUNT(*) as nCasos"
    sSQL = sSQL & " From Producto Prd"
    sSQL = sSQL & " Inner join Colocaciones C ON C.cCtaCod = Prd.cCtaCod "
    sSQL = sSQL & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSQL = sSQL & " Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod And CV.cPersCod in " & sCadInstitucion
    sSQL = sSQL & " Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSQL = sSQL & " Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & " Inner Join (Select Cdet.cCtaCod, Cdet.nNroCalen, SUM(nMonto-nMontoPagado) as nMonto "
    sSQL = sSQL & "     From ColocCalendDet CDet"
    sSQL = sSQL & "     Inner Join ColocCalendario CC ON Cdet.cCtaCod =   CC.cCtaCod AND Cdet.nNroCalen =   CC.nNroCalen AND Cdet.nCuota =   CC.nCuota AND Cdet.nColocCalendApl =   CC.nColocCalendApl"
    sSQL = sSQL & "     Where Cdet.nColocCalendApl = 1  AND CC.dVenc <= '" & Format(pdFecIni, "mm/dd/yyyy") & "' "
    sSQL = sSQL & "     Group By Cdet.cCtaCod, Cdet.nNroCalen"
    sSQL = sSQL & "        ) as CDet ON CDet.cCtaCod = Prd.cCtaCod And CDet.nNroCalen = CC.nNroCalen "
    sSQL = sSQL & " Inner Join (Select cCtaCod, nNroCalen, SUM(nMonto-nMontoPagado) as nMonto "
    sSQL = sSQL & "     From ColocCalendDet CDet"
    sSQL = sSQL & "     Where nColocCalendApl = 1 And nPrdConceptoCod = 1101"
    sSQL = sSQL & "     Group By cCtaCod, nNroCalen "
    sSQL = sSQL & "        ) as CDet2 ON CDet2.cCtaCod = Prd.cCtaCod And CDet2.nNroCalen = CC.nNroCalen "
    sSQL = sSQL & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor
    sSQL = sSQL & " And SUBSTRING(Prd.cCtaCod,9,1) = '" & pnMoneda & "' And A.cAgeCod IN " & sCadAge
    sSQL = sSQL & " Group By A.cAgeDescripcion, P.cPersNombre, Prd.cCtaCod"
    sSQL = sSQL & " WITH ROLLUP"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaMoraXInstituciones = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaResumenSaldosCarteraPorInstitucionConsumo(ByVal pMatAgencias As Variant, _
    ByVal pMatProd As Variant, ByVal pdHoy As Date) As ADODB.Recordset
    
Dim sSQL As String
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

    sSQL = "SELECT "
    sSQL = sSQL & " CASE WHEN (GROUPING(T.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(T.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSQL = sSQL & "    END AS sAgencia,"
    sSQL = sSQL & " CASE WHEN (GROUPING(T.cPersNombre) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(T.cPersNombre, 'INSTITUCION DESCONOCIDA')"
    sSQL = sSQL & "    END AS sInstitucion,"
    sSQL = sSQL & " SUM(T.nNroCreditosMN) as nNroCreditosMN,"
    sSQL = sSQL & " SUM(T.nSaldoCapMN) as nSaldoCapMN,"
    sSQL = sSQL & " SUM(T.nDesembolsadoMN) as nDesembolsadoMN,"
    sSQL = sSQL & " SUM(T.nNroCreditosME) as nNroCreditosME,"
    sSQL = sSQL & " SUM(T.nSaldoCapME) as nSaldoCapME,"
    sSQL = sSQL & " SUM(T.nDesembolsadoME) as nDesembolsadoME,"
    sSQL = sSQL & " SUM(T.nNroCreditos30MN) as nNroCreditos30MN,"
    sSQL = sSQL & " SUM(T.nSaldoCap30MN) as nSaldoCap30MN,"
    sSQL = sSQL & " SUM(T.nNroCreditos30ME) as nNroCreditos30ME,"
    sSQL = sSQL & " SUM(T.nSaldoCap30ME) As nSaldoCap30ME"
    sSQL = sSQL & " From ("
    sSQL = sSQL & " select A.cAgeDescripcion, P.cPersNombre, SUM(Prd.nSaldo) as nSaldoCapMN, SUM(C.nMontoCol) as nDesembolsadoMN, Count(*) as nNroCreditosMN, 0.00 as nSaldoCapME, 0.00 as nDesembolsadoME, 0 as nNroCreditosME,"
    sSQL = sSQL & "    0 as nSaldoCap30MN, 0 as nNroCreditos30MN,"
    sSQL = sSQL & "    0 as nSaldoCap30ME, 0 as nNroCreditos30ME"
    sSQL = sSQL & " From Producto as Prd"
    sSQL = sSQL & "    Inner Join Colocaciones C ON Prd.cCtaCod = C.cCtaCod"
    sSQL = sSQL & "    Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSQL = sSQL & "    Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & "    Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSQL = sSQL & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor & " and SUBSTRING(Prd.cCtaCod,9,1)='1'"
    sSQL = sSQL & "    And Prd.cCtaCod <>  '305' And A.cAgecod in " & sCadAge & sCadProd
    sSQL = sSQL & " Group by P.cPersNombre, A.cAgeDescripcion"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select A.cAgeDescripcion, P.cPersNombre, 0.00 as nSaldoCapMN, 0.00 as nDesembolsadoMN, 0 as nNroCreditosMN, SUM(Prd.nSaldo/dbo.GetTipoCambio(1,'" & Format(pdHoy, "mm/dd/yyyy") & "')) as nSaldoCapME, SUM(C.nMontoCol/dbo.GetTipoCambio(1,'" & Format(pdHoy, "mm/dd/yyyy") & "')) as nDesembolsadoME, Count(*) as nNroCreditosME,"
    sSQL = sSQL & "    0 as nSaldoCap30MN, 0 as nNroCreditos30MN,"
    sSQL = sSQL & "    0 as nSaldoCap30ME, 0 as nNroCreditos30ME"
    sSQL = sSQL & " From Producto as Prd"
    sSQL = sSQL & "    Inner Join Colocaciones C ON Prd.cCtaCod = C.cCtaCod"
    sSQL = sSQL & "    Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSQL = sSQL & "    Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & "    Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSQL = sSQL & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor & " and SUBSTRING(Prd.cCtaCod,9,1)='2'"
    sSQL = sSQL & "    And Prd.cCtaCod <>  '305' And A.cAgecod in " & sCadAge & sCadProd
    sSQL = sSQL & " Group by P.cPersNombre, A.cAgeDescripcion"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select A.cAgeDescripcion, P.cPersNombre, 0.00 as nSaldoCapMN, 0.00 as nDesembolsadoMN, 0 as nNroCreditosMN, 0 as nSaldoCapME, 0 as nDesembolsadoME, 0 as nNroCreditosME,"
    sSQL = sSQL & "    SUM(Prd.nSaldo) as nSaldoCap30MN, Count(*) as nNroCreditos30MN,"
    sSQL = sSQL & "    0 as nSaldoCap30ME,  0 as nNroCreditos30ME"
    sSQL = sSQL & " From Producto as Prd"
    sSQL = sSQL & "    Inner Join Colocaciones C ON Prd.cCtaCod = C.cCtaCod"
    sSQL = sSQL & "    Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSQL = sSQL & "    Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & "    Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSQL = sSQL & "    Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSQL = sSQL & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor & " and SUBSTRING(Prd.cCtaCod,9,1)='1'"
    sSQL = sSQL & "    And Prd.cCtaCod <>  '305' and CC.nDiasAtraso >= 30 And A.cAgecod in " & sCadAge & sCadProd
    sSQL = sSQL & " Group by P.cPersNombre, A.cAgeDescripcion"
    sSQL = sSQL & "    Union"
    sSQL = sSQL & " Select A.cAgeDescripcion, P.cPersNombre, 0.00 as nSaldoCapMN, 0.00 as nDesembolsadoMN, 0 as nNroCreditosMN, 0 as nSaldoCapME, 0 as nDesembolsadoME, 0 as nNroCreditosME,"
    sSQL = sSQL & "    0 as nSaldoCap30MN, 0 as nNroCreditos30MN,"
    sSQL = sSQL & "    SUM(Prd.nSaldo/dbo.GetTipoCambio(1,'" & Format(pdHoy, "mm/dd/yyyy") & "')) as nSaldoCap30ME, Count(*) as nNroCreditos30ME"
    sSQL = sSQL & " From Producto as Prd"
    sSQL = sSQL & "    Inner Join Colocaciones C ON Prd.cCtaCod = C.cCtaCod"
    sSQL = sSQL & "    Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSQL = sSQL & "    Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & "    Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSQL = sSQL & "    Inner Join ColocacCred CC ON Prd.cCtaCod = CC.cCtaCod"
    sSQL = sSQL & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor & " and SUBSTRING(Prd.cCtaCod,9,1)='2'"
    sSQL = sSQL & "    And Prd.cCtaCod <>  '305' and CC.nDiasAtraso >= 30 And A.cAgecod in " & sCadAge & sCadProd
    sSQL = sSQL & " Group by P.cPersNombre, A.cAgeDescripcion"
    sSQL = sSQL & "   ) as T"
    sSQL = sSQL & " Group by T.cAgeDescripcion, T.cPersNombre"
    sSQL = sSQL & " With ROLLUP"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaResumenSaldosCarteraPorInstitucionConsumo = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaResumenSaldosCarteraPorAnalistaConsumo(ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal P1I As Integer, ByVal P1F As Integer, _
    ByVal P2I As Integer, ByVal P2F As Integer, ByVal P4F As Integer, _
    ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, pdFecha As Date) As ADODB.Recordset
    
Dim sSQL As String
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
    
    
    sSQL = "SELECT "
    sSQL = sSQL & "    CASE WHEN (GROUPING(T.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(T.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSQL = sSQL & "    END AS sAgencia,"
    sSQL = sSQL & " CASE WHEN (GROUPING(T.cPersNombre) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(T.cPersNombre, 'INSTITUCION DESCONOCIDA')"
    sSQL = sSQL & "    END AS sInstitucion,"
    sSQL = sSQL & " SUM(T.nNroCreditosMN) as nNroCreditosMN,"
    sSQL = sSQL & " SUM(T.nSaldoCapMN) as nSaldoCapMN,"
    sSQL = sSQL & " SUM(T.nDesembolsadoMN) as nDesembolsadoMN,"
    sSQL = sSQL & " SUM(T.nNroCreditosME) as nNroCreditosME,"
    sSQL = sSQL & " SUM(T.nSaldoCapME) as nSaldoCapMN,"
    sSQL = sSQL & " SUM(T.nDesembolsadoME) As nDesembolsadoME"
    sSQL = sSQL & " From ("
    sSQL = sSQL & " select A.cAgeDescripcion, P.cPersNombre, SUM(Prd.nSaldo) as nSaldoCapMN, SUM(C.nMontoCol) as nDesembolsadoMN, Count(*) as nNroCreditosMN, 0.00 as nSaldoCapME, 0.00 as nDesembolsadoME, 0 as nNroCreditosME"
    sSQL = sSQL & " From Producto as Prd"
    sSQL = sSQL & "    Inner Join Colocaciones C ON Prd.cCtaCod = C.cCtaCod"
    sSQL = sSQL & "    Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSQL = sSQL & "    Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & "    Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSQL = sSQL & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor & " and SUBSTRING(Prd.cCtaCod,9,1)='1'"
    sSQL = sSQL & "    And Prd.cCtaCod <>  '305'"
    sSQL = sSQL & " Group by P.cPersNombre, A.cAgeDescripcion"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select A.cAgeDescripcion, P.cPersNombre, 0.00 as nSaldoCapMN, 0.00 as nDesembolsadoMN, 0 as nNroCreditosMN, SUM(Prd.nSaldo/dbo.GetTipoCambio(1,'" & Format(pdFecha, "mm/dd/yyyy") & "')) as nSaldoCapME, SUM(C.nMontoCol/dbo.GetTipoCambio(1,'" & Format(pdFecha, "mm/dd/yyyy") & "')) as nDesembolsadoME, Count(*) as nNroCreditosME"
    sSQL = sSQL & " From Producto as Prd"
    sSQL = sSQL & "    Inner Join Colocaciones C ON Prd.cCtaCod = C.cCtaCod"
    sSQL = sSQL & "    Inner join Agencias A ON A.cAgeCod = SUBSTRING(Prd.cCtaCod,4,2)"
    sSQL = sSQL & "    Inner Join ColocacConvenio CV ON CV.cCtaCod = Prd.cCtaCod"
    sSQL = sSQL & "    Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSQL = sSQL & " Where Prd.nPrdEstado >= " & gColocEstVigNorm & " And Prd.nPrdEstado <= " & gColocEstRefMor & " and SUBSTRING(Prd.cCtaCod,9,1)='2'"
    sSQL = sSQL & "    And Prd.cCtaCod <>  '305'"
    sSQL = sSQL & " Group by P.cPersNombre, A.cAgeDescripcion"
    sSQL = sSQL & "   ) as T"
    sSQL = sSQL & " Group by T.cAgeDescripcion, T.cPersNombre"
    sSQL = sSQL & " With ROLLUP"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaResumenSaldosCarteraPorAnalistaConsumo = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaPagoConCheque(ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal pMatCondicion As Variant, ByVal pMatProd As Variant, _
    ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pnTipoBusq As Integer, ByVal psNroDoc As String) As ADODB.Recordset
    
Dim sSQL As String
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

    sSQL = " select A.cAgeDescripcion as sAgencia, P.cPersNombre as cInstitucion, MC.cCtaCod, sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = MC.cCtaCod),"
    sSQL = sSQL & " cOperacion = (Select CN.cConsDescripcion From Constante CN  where CN.nConsValor = MC.nCredEstado And CN.nConsCod = 3001),"
    sSQL = sSQL & "    cNroDoc = (Select cDocNro From MovDoc Where nMovNro = MC.nMovNro),"
    sSQL = sSQL & "    MC.nMonto as nMonto,"
    sSQL = sSQL & "    cUser = (Select RIGHT(cMovNro,4) From Mov Where nMovNro = MC.nMovNro),"
    sSQL = sSQL & "    dMov = (Select SUBSTRING(cMovNro,7,2) + '/' + SUBSTRING(cMovNro,5,2) + '/' + SUBSTRING(cMovNro,1,4) + ' ' + SUBSTRING(cMovNro,9,2) + ':' + SUBSTRING(cMovNro,11,2) + ':' + SUBSTRING(cMovNro,13,2) From Mov Where nMovNro = MC.nMovNro)"
    sSQL = sSQL & " From MovCol MC"
    sSQL = sSQL & "    Inner join Agencias A ON A.cAgeCod = SUBSTRING(MC.cCtaCod,4,2)"
    sSQL = sSQL & "    Inner Join ColocacConvenio CV ON CV.cCtaCod = MC.cCtaCod"
    sSQL = sSQL & "    Inner Join Persona P ON P.cPersCod = CV.cPersCod"
    sSQL = sSQL & "    Inner Join Mov M ON M.nMovNro = MC.nMovNro"
    sSQL = sSQL & "    Inner Join MovDoc MD ON MD.nMovNro = M.nMovNro" & sCadBusq
    sSQL = sSQL & "    Inner Join ColocacCred CC ON CC.cCtaCod = MC.cCtaCod"
    sSQL = sSQL & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'"
    sSQL = sSQL & "    And SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdFecIni, "yyyymmdd") & "' AND '" & Format(pdFecFin, "yyyymmdd") & "'"
    sSQL = sSQL & "    And A.cAgecod in " & sCadAge & sCadProd
    sSQL = sSQL & "    And CC.nColocCondicion in " & sCadCondicion
    sSQL = sSQL & "  And SUBSTRING(MC.cCtaCod, 9, 1) = '" & pnMoneda & "' "
    sSQL = sSQL & " Order by A.cAgeDescripcion, P.cPersNombre"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaPagoConCheque = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaPagoDeOtrasAgencias(ByVal pnMoneda As Moneda, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, psCodAge As String, ByVal psCodCmact As String) As ADODB.Recordset
    
Dim sSQL As String
Dim oConecta As DConecta

    sSQL = " Select A.cAgeDescripcion as sAgencia, MC.cCtaCod, "
    sSQL = sSQL & " cOperacion = (Select O.cOpeDesc From OpeTpo O  Where O.cOpeCod = MC.cOpeCod),"
    sSQL = sSQL & " cNroDoc = (Select cDocNro From MovDoc Where nMovNro = MC.nMovNro),"
    sSQL = sSQL & " MC.nMonto as nMonto,"
    sSQL = sSQL & " cUser = (Select RIGHT(cMovNro,4) From Mov Where nMovNro = MC.nMovNro),"
    sSQL = sSQL & " dMov = (Select SUBSTRING(cMovNro,7,2) + '/' + SUBSTRING(cMovNro,5,2) + '/' + SUBSTRING(cMovNro,1,4) + ' ' + SUBSTRING(cMovNro,9,2) + ':' + SUBSTRING(cMovNro,11,2) + ':' + SUBSTRING(cMovNro,13,2) From Mov Where nMovNro = MC.nMovNro)"
    sSQL = sSQL & " From MovCol MC"
    sSQL = sSQL & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(MC.cCtaCod,4,2)"
    sSQL = sSQL & " Inner Join Mov M ON M.nMovNro = MC.nMovNro And SUBSTRING(M.cMovNro,15,3) = '" & psCodCmact & "'"
    sSQL = sSQL & " Left Join MovDoc MD ON MD.nMovNro = M.nMovNro"
    sSQL = sSQL & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'"
    sSQL = sSQL & "    And LEN(MC.cCtaCod)=18"
    sSQL = sSQL & "    And SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdFecIni, "yyyymmdd") & "' AND '" & Format(pdFecFin, "yyyymmdd") & "'"
    sSQL = sSQL & "    And SUBSTRING(MC.cCtaCod,4,2) <> '" & psCodAge & "' And SUBSTRING(M.cMovNro,18,2) = '" & psCodAge & "' "
    sSQL = sSQL & "  And SUBSTRING(MC.cCtaCod, 9, 1) = '" & pnMoneda & "' "
    sSQL = sSQL & " AND MC.cOpeCod like '100[234567]%' "
    sSQL = sSQL & " Order by A.cAgeDescripcion, M.cMovNro"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaPagoDeOtrasAgencias = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaPagoENOtrasAgencias(ByVal pnMoneda As Moneda, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, psCodAge As String, ByVal psCodCmact As String, _
    ByVal pMatCondicion As Variant, ByVal pMatProd As Variant) As ADODB.Recordset
    
Dim sSQL As String
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

    sSQL = " Select A.cAgeDescripcion as sAgencia, MC.cCtaCod, P.cPersNombre,"
    sSQL = sSQL & " cOperacion = (Select O.cOpeDesc From OpeTpo O  Where O.cOpeCod = MC.cOpeCod),"
    sSQL = sSQL & " cNroDoc = (Select cDocNro From MovDoc Where nMovNro = MC.nMovNro),"
    sSQL = sSQL & " MC.nMonto as nMonto,"
    sSQL = sSQL & " cUser = (Select RIGHT(cMovNro,4) From Mov Where nMovNro = MC.nMovNro),"
    sSQL = sSQL & " dMov = (Select SUBSTRING(cMovNro,7,2) + '/' + SUBSTRING(cMovNro,5,2) + '/' + SUBSTRING(cMovNro,1,4) + ' ' + SUBSTRING(cMovNro,9,2) + ':' + SUBSTRING(cMovNro,11,2) + ':' + SUBSTRING(cMovNro,13,2) From Mov Where nMovNro = MC.nMovNro)"
    sSQL = sSQL & " From MovCol MC"
    sSQL = sSQL & " Inner Join Mov M ON M.nMovNro = MC.nMovNro And SUBSTRING(M.cMovNro,15,3) = '" & psCodCmact & "'"
    sSQL = sSQL & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(M.cMovNro,18,2) "
    sSQL = sSQL & " Inner join ProductoPersona PP ON PP.cCtaCod = MC.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & " Inner Join Persona P ON P.cPersCod = PP.cPersCod "
    sSQL = sSQL & " Left Join MovDoc MD ON MD.nMovNro = M.nMovNro"
    sSQL = sSQL & " Inner Join ColocacCred CC ON CC.cCtaCod = MC.cCtaCod "
    sSQL = sSQL & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'"
    sSQL = sSQL & "    And LEN(MC.cCtaCod)=18"
    sSQL = sSQL & "    And SUBSTRING(M.cMovNro,1,8) BETWEEN '" & Format(pdFecIni, "yyyymmdd") & "' AND '" & Format(pdFecFin, "yyyymmdd") & "'"
    sSQL = sSQL & "    And SUBSTRING(MC.cCtaCod,4,2) = '" & psCodAge & "' And SUBSTRING(M.cMovNro,18,2) <> '" & psCodAge & "' "
    sSQL = sSQL & "  And SUBSTRING(MC.cCtaCod, 9, 1) = '" & pnMoneda & "' " & sCadProd
    sSQL = sSQL & "    And CC.nColocCondicion in " & sCadCondicion
    sSQL = sSQL & " AND MC.cOpeCod like '100[234567]%' "
    sSQL = sSQL & " Order by A.cAgeDescripcion, M.cMovNro "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaPagoENOtrasAgencias = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function


Public Function RecuperaInteresesEnSuspenso(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatProd As Variant)
    
Dim sSQL As String
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
    
    sSQL = "SELECT "
    sSQL = sSQL & "   CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSQL = sSQL & "   END AS sAgencia,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sProducto,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(CN2.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(CN2.cConsDescripcion, 'PLAZO DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sPlazo,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sFuenteFinan,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(MC.cCtaCod) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(MC.cCtaCod, 'CREDITO DESCONOCIDO')"
    sSQL = sSQL & "   END AS cCtaCod,"
    sSQL = sSQL & " COUNT(*) as nCasos,"
    sSQL = sSQL & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = MC.cCtaCod),"
    sSQL = sSQL & " dRefinanciado = (Select CE.dPrdEstado From Producto CE Where CE.cCtaCod = MC.cCtaCod And CE.nPrdEstado = 2060),"
    sSQL = sSQL & " SUM(ISNULL(Cal.nMonto,0)) as nInteresComp,"
    sSQL = sSQL & " SUM(ISNULL(Cal2.nMonto,0)) as nInteresMor,"
    sSQL = sSQL & " SUM(IsNull(Cal3.nMonto, 0)) As nGastos"
    sSQL = sSQL & " FROM Producto MC Inner Join Constante CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSQL = sSQL & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(MC.cCtaCod,4,2)"
    sSQL = sSQL & " Inner Join Colocaciones CC ON CC.cCtaCod = MC.cCtaCod"
    sSQL = sSQL & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = CC.cLineaCred"
    sSQL = sSQL & " Inner Join Persona P ON CL.cPersCod = P.cPersCod "
    sSQL = sSQL & " Inner join Constante CN2 ON CN2.nConsValor = Convert(int,SUBSTRING(CC.cLineaCred,6,1)) And CN2.nConsCod = " & gColocLineaCredPlazo
    sSQL = sSQL & " Inner join ColocacRefinanc CR ON CR.cCtaCodRef = MC.cCtaCod And CR.nEstado = " & gColocEstadoRefinancAprobado
    sSQL = sSQL & " Left Join ( Select CD.cCtaCodRef, SUM(isnull(CD.nMonto,0) - isnull(CD.nMontoPagado,0)) as nMonto From ColocacRefinancDet CD"
    sSQL = sSQL & "        Where (CD.nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & " Or CD.nPrdConceptoCod = "
    sSQL = sSQL & gColocConceptoCodInteresGracia & " Or CD.nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado
    sSQL = sSQL & " Or CD.nPrdConceptoCod = " & gColocConceptoCodInteresCompVencido & ") And nEstado = 3 "
    sSQL = sSQL & "        GROUP BY CD.cCtaCodRef) CAL ON CAL.cCtaCodRef = MC.cCtaCod"
    sSQL = sSQL & " Left Join ( Select CD.cCtaCodRef, SUM(CD.nMonto - isnull(CD.nMontoPagado,0)) as nMonto From ColocacRefinancDet CD "
    sSQL = sSQL & "        Where CD.nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & " And nEstado = 3 "
    sSQL = sSQL & "        GROUP BY CD.cCtaCodRef) CAL2 ON CAL2.cCtaCodRef = MC.cCtaCod"
    sSQL = sSQL & " Left Join ( Select CD.cCtaCodRef, SUM(CD.nMonto - isnull(CD.nMontoPagado,0)) as nMonto From ColocacRefinancDet CD "
    sSQL = sSQL & "        Where CONVERT(varchar(6),CD.nPrdConceptoCod) like '12%' And nEstado = 3 "
    sSQL = sSQL & "        GROUP BY CD.cCtaCodRef) CAL3 ON CAL3.cCtaCodRef = MC.cCtaCod "
    sSQL = sSQL & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'  and Len(MC.cCtaCod) = 18"
    sSQL = sSQL & " And MC.nPrdEstado = " & gColocEstRefinanc
    sSQL = sSQL & "    And A.cAgecod in " & sCadAge & sCadProd
    sSQL = sSQL & "  And SUBSTRING(MC.cCtaCod, 9, 1) = '" & pnMoneda & "' "
    sSQL = sSQL & " GROUP BY A.cAgeDescripcion, CN.cConsDescripcion, CN2.cConsDescripcion, P.cPersNombre, MC.cCtaCod "
    sSQL = sSQL & " WITH ROLLUP "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaInteresesEnSuspenso = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
 
Public Function RecuperaCreditosRefinanciados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal pMatProd As Variant)
    
Dim sSQL As String
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
    
    sSQL = "SELECT DISTINCT "
    sSQL = sSQL & "   CASE WHEN (GROUPING(A.cAgeDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(A.cAgeDescripcion, 'AGENCIA DESCONOCIDA ')"
    sSQL = sSQL & "   END AS sAgencia,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(CN.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(CN.cConsDescripcion, 'PRODUCTO DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sProducto,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(CN2.cConsDescripcion) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(CN2.cConsDescripcion, 'PLAZO DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sPlazo,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(P.cPersNombre) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(P.cPersNombre, 'FFinan DESCONOCIDO ')"
    sSQL = sSQL & "   END AS sFuenteFinan,"
    sSQL = sSQL & "   CASE WHEN (GROUPING(MC.cCtaCod) = 1) THEN 'TOTAL'"
    sSQL = sSQL & "        ELSE ISNULL(MC.cCtaCod, 'CREDITO DESCONOCIDO')"
    sSQL = sSQL & "   END AS cCtaCod,"
    sSQL = sSQL & " COUNT(*) as nCasos,"
    sSQL = sSQL & " sTitular = (select P2.cPersNombre From Persona P2 Inner Join ProductoPersona PP ON P2.cPersCod = PP.cPersCod And PP.nPrdPersrelac = 20 Where PP.cCtaCod = MC.cCtaCod),"
    sSQL = sSQL & " dRefinanciado = (Select CE.dVigencia From Colocaciones CE Where CE.cCtaCod = MC.cCtaCod),"
    sSQL = sSQL & " SUM(ISNULL(Cal.nMonto,0)) as nInteresComp,"
    sSQL = sSQL & " SUM(ISNULL(Cal2.nMonto,0)) as nInteresMor,"
    sSQL = sSQL & " SUM(IsNull(Cal3.nMonto, 0)) As nGastos, "
    sSQL = sSQL & " SUM(IsNull(Cal4.nMonto, 0)) As nCapital, "
    sSQL = sSQL & " ISNULL(Cre.bRefCapInt,0) as bRefCapInt "
    sSQL = sSQL & " FROM Producto MC Inner Join Constante CN ON Convert(Int,SUBSTRING(MC.cCtaCod,6,3)) = CN.nConsValor And CN.nConsCod = " & gProducto
    sSQL = sSQL & " Inner join Agencias A ON A.cAgeCod = SUBSTRING(MC.cCtaCod,4,2)"
    sSQL = sSQL & " Inner Join Colocaciones CC ON CC.cCtaCod = MC.cCtaCod"
    sSQL = sSQL & " Inner Join ColocacCred Cre ON Cre.cCtaCod = MC.cCtaCod"
    sSQL = sSQL & " Inner Join ColocLineaCredito CL ON CL.cLineaCred = CC.cLineaCred"
    sSQL = sSQL & " Inner Join Persona P ON CL.cPersCod = P.cPersCod "
    sSQL = sSQL & " Inner join Constante CN2 ON CN2.nConsValor = Convert(int,SUBSTRING(CC.cLineaCred,6,1)) And CN2.nConsCod = " & gColocLineaCredPlazo
    sSQL = sSQL & " Inner join ColocacRefinanc CR ON CR.cCtaCod = MC.cCtaCod And CR.nEstado = " & gColocEstadoRefinancAprobado
    sSQL = sSQL & " Left Join ( Select CD.cCtaCod, SUM(isnull(CD.nMonto,0) - isnull(CD.nMontoPagado,0)) as nMonto From ColocacRefinancDet CD"
    sSQL = sSQL & "        Where ( CD.nPrdConceptoCod = " & gColocConceptoCodInteresCompensatorio & " Or CD.nPrdConceptoCod = "
    sSQL = sSQL & gColocConceptoCodInteresGracia & " Or CD.nPrdConceptoCod = " & gColocConceptoCodInteresReprogramado
    sSQL = sSQL & " Or CD.nPrdConceptoCod = " & gColocConceptoCodInteresCompVencido & ") AND nEstado=3 "
    sSQL = sSQL & "        GROUP BY CD.cCtaCod) CAL ON CAL.cCtaCod = MC.cCtaCod"
    sSQL = sSQL & " Left Join ( Select CD.cCtaCod, SUM(CD.nMonto - isnull(CD.nMontoPagado,0)) as nMonto From ColocacRefinancDet CD "
    sSQL = sSQL & "        Where (CD.nPrdConceptoCod = " & gColocConceptoCodInteresMoratorio & ") AND nEstado=3 "
    sSQL = sSQL & "        GROUP BY CD.cCtaCod) CAL2 ON CAL2.cCtaCod = MC.cCtaCod"
    sSQL = sSQL & " Left Join ( Select CD.cCtaCod, SUM(CD.nMonto - isnull(CD.nMontoPagado,0)) as nMonto From ColocacRefinancDet CD "
    sSQL = sSQL & "        Where (CONVERT(varchar(6),CD.nPrdConceptoCod) like '12%'" & " ) AND nEstado=3 "
    sSQL = sSQL & "        GROUP BY CD.cCtaCod) CAL3 ON CAL3.cCtaCod = MC.cCtaCod "
    sSQL = sSQL & " Left Join ( Select CD.cCtaCod, SUM(CD.nMonto - isnull(CD.nMontoPagado,0)) as nMonto From ColocacRefinancDet CD "
    sSQL = sSQL & "        Where ( CD.nPrdConceptoCod = " & gColocConceptoCodCapital & ") AND nEstado=3 "
    sSQL = sSQL & "        GROUP BY CD.cCtaCod) CAL4 ON CAL4.cCtaCod = MC.cCtaCod "
    sSQL = sSQL & " Where SUBSTRING(MC.cCtaCod,6,3)<>'305'  and Len(MC.cCtaCod) = 18"
    sSQL = sSQL & "    And A.cAgecod in " & sCadAge & sCadProd
    sSQL = sSQL & "  And SUBSTRING(MC.cCtaCod, 9, 1) = '" & pnMoneda & "' "
    sSQL = sSQL & " GROUP BY Cre.bRefCapInt, A.cAgeDescripcion, CN.cConsDescripcion, CN2.cConsDescripcion, P.cPersNombre, MC.cCtaCod "
    sSQL = sSQL & " WITH ROLLUP "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosRefinanciados = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
 

