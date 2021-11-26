VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************
'***     Rutina:           DCredito
'***     Descripcion:      Clase que permite Recuperar Datos del Credito
'***     Creado por:        NSSE
'***     Maquina:           07SISTDES03
'***     Fecha-Tiempo:         12/07/2001 12:06:08 AM
'***     Ultima Modificacion: Inicio de descripicion
'*****************************************************************************************

Option Explicit

Private gConsPersona As String
Private gConsComunes As String
Private gConsImagenes As String
Private Conn As DConecta

Public Function RecuperaGarantiasCreditoConsol(ByVal psPersCod As String, ByVal pdFecha As Date) As ADODB.Recordset
Dim oConecta As DConecta
Dim sSql As String
Dim oGeneral As DGeneral
Dim nTipoCambioFijo As Double

    On Error GoTo ErrorRecuperaGarantiasCredito
    
    Set oGeneral = New DGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
    nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))
    Set oGeneral = Nothing
    
    sSql = "Select SUM( "
    sSql = sSql & "         CASE WHEN  G.nMoneda = 1 THEN G.nPorGravar "
    sSql = sSql & "              WHEN  G.nMoneda = 2 THEN G.nPorGravar * " & Format(nTipoCambioFijo, "#0.00") & " END "
    sSql = sSql & "   ) nPorGravar,  "
    sSql = sSql & "    SUM( "
    sSql = sSql & "         CASE WHEN  G.nMoneda = 1 THEN G.nRealizacion "
    sSql = sSql & "              WHEN  G.nMoneda = 2 THEN G.nRealizacion * " & Format(nTipoCambioFijo, "#0.00") & " END "
    sSql = sSql & "  ) nRealizacion,  "
    sSql = sSql & "    SUM( "
    sSql = sSql & "         CASE WHEN  G.nMoneda = 1 THEN G.nTasacion "
    sSql = sSql & "              WHEN  G.nMoneda = 2 THEN G.nTasacion * " & Format(nTipoCambioFijo, "#0.00") & " END "
    sSql = sSql & "  ) nTasacion  "
    sSql = sSql & " From Garantias G  "
    sSql = sSql & "                       Inner Join Constante CN ON G.nTpoGarantia = CN.nConsvalor And CN.nConsCod = " & Trim(Str(gPersGarantia))
    sSql = sSql & "                       Inner Join PersGarantia PG ON G.cNumGarant = PG.cNumGarant AND PG.nRelacion = " & Trim(Str(gPersRelGarantiaTitular))
    sSql = sSql & "                       Inner Join Persona P ON PG.cPersCod = P.cPersCod "
    sSql = sSql & " WHERE PG.cPersCod = '" & psPersCod & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantiasCreditoConsol = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaGarantiasCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function NumerosCredEnJudicial(ByVal psPersCod As String)
Dim sSql As String
Dim oCon As DConecta
Dim R As ADODB.Recordset
    sSql = " select ISNULL(COUNT(P.cCtaCod),0) as NumCredJud from Producto P"
    sSql = sSql & "     Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod"
    sSql = sSql & " where PP.nPrdPersRelac = 20 AND P.nPrdEstado in (2201,2202)"
    sSql = sSql & " AND PP.cPersCod = '" & psPersCod & "'"
    
    Set oCon = New DConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    If R.RecordCount > 0 Then
        NumerosCredEnJudicial = R!NumCredJud
    Else
        NumerosCredEnJudicial = 0
    End If
    R.Close
    
    
End Function

Public Function RecuperaUltimoMovimiento(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    sSql = " Select MC.nMovNro, MC.cOpeCod, dbo.FechaMov(M.cMovNro) as dFecPago,"
    sSql = sSql & " nNroCuota = (Select MAX(nNroCuota) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro),"
    sSql = sSql & " nMontoPagado = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro),"
    sSql = sSql & " nCapital = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod = 1000 ),"
    sSql = sSql & " nInteres = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in (1100,1102,1103,1104,1105)),"
    sSql = sSql & " nMora = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod = 1101),"
    sSql = sSql & " nGastos = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod like  '12%'),"
    sSql = sSql & " MC.nDiasMora , MC.nSaldoCap"
    sSql = sSql & " from MovCol MC Inner Join Mov M ON M.nMovNro = MC.nMovNro"
    sSql = sSql & " where MC.cOpeCod like '100%' AND MC.cCtaCod = '" & psctacod & "' "
    sSql = sSql & " AND (SUBSTRING(MC.cOpeCod,1,4) in ('1002','1003','1004','1005','1006','1007','1070'))"
    sSql = sSql & " AND MC.nMovNro = (Select MAX(nMovNro) From MovCol Where cCtaCod = '" & psctacod & "' AND (SUBSTRING(cOpeCod,1,4) in ('1002','1003','1004','1005','1006','1007','1070')) ) "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaUltimoMovimiento = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaCreditosVigentesGrabados(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim i As Integer

    
    sSql = "Select cCtaCod, cCtaCodRef From ColocCredCredVig Where cCtaCod = '" & psctacod & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosVigentesGrabados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    
End Function


Public Function RecuperaChequeCreditos(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pnMoneda As Integer) As ADODB.Recordset
    
Dim sSql As String
Dim oConec As DConecta
Dim sFecha As String

    sSql = "select  '' , '.' as OPT,P.cPersNombre, DR.cIFCta, DR.cNroDoc,  DR.nMonto, C3.cConsDescripcion cProducto, "
    sSql = sSql & " C2.cConsDescripcion cEstado, DR.dValorizacion, DR.nEstado,DR.cPersCod "
    sSql = sSql & " from DocRec DR"
    sSql = sSql & " Inner Join Persona P ON P.cPersCod = DR.cPersCod"
    sSql = sSql & " Inner Join Constante C2 ON C2.nConsValor = DR.nEstado AND C2.nConsCod = 1007"
    sSql = sSql & " Inner Join ColocCredConsFiltro CF ON CF.nConsValor = DR.nProducto AND CF.nConsCod = 1001"
    sSql = sSql & " Inner Join Constante C3 ON C3.nConsCod = 1001 AND C3.nConsValor = CF.nConsValor"
    sSql = sSql & " Inner Join DocRecEst  RE on RE.cNroDoc=DR.cNroDoc and RE.nTpoDoc=DR.nTpoDoc and RE.cPersCod=DR.cPersCod and RE.cIFTpo=DR.cIFTpo"
    sSql = sSql & " Where  dbo.FechaMov(RE.cMovNro) >= '" & Format(pdFecIni, "dd/mm/yyyy") & "' AND dbo.FechaMov(RE.cMovNro) <= '" & Format(pdFecFin, "dd/mm/yyyy") & "' "
    sSql = sSql & " AND DR.nMoneda = " & pnMoneda & " AND DR.nTpoDoc = 47 "
    sSql = sSql & " Order by P.cPersNombre,DR.dValorizacion"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set RecuperaChequeCreditos = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
End Function

Public Function RecuperaPagosRealizados(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

'    sSQL = " Select MC.nMovNro, MC.cOpeCod, dbo.FechaMov(M.cMovNro) as dFecPago,"
'    sSQL = sSQL & " nNroCuota = (Select MAX(nNroCuota) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro),"
'    sSQL = sSQL & " nMontoPagado = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro),"
'    sSQL = sSQL & " nCapital = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod = 1000 ),"
'    sSQL = sSQL & " nInteres = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in (1100,1102,1103,1104,1105)),"
'    sSQL = sSQL & " nMora = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod = 1101),"
'    sSQL = sSQL & " nGastos = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod like  '12%'),"
'    sSQL = sSQL & " nITF = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in  (20, 21)),  "
'    sSQL = sSQL & " MC.nDiasMora , MC.nSaldoCap"
'    sSQL = sSQL & " from MovCol MC Inner Join Mov M ON M.nMovNro = MC.nMovNro"
'    sSQL = sSQL & " where MC.cOpeCod like '100%' AND MC.cCtaCod = '" & psCtaCod & "' "
'    sSQL = sSQL & " AND (SUBSTRING(MC.cOpeCod,1,4) in ('1002','1003','1004','1005','1006','1007','1070'))"
'    sSQL = sSQL & " Order by MC.nMovNro "

 sSql = "Select MC.nMovNro, MC.cOpeCod, dbo.FechaMov(M.cMovNro) as dFecPago, "
 sSql = sSql & " nNroCuota = (Select MAX(nNroCuota) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro),"
 sSql = sSql & " nMontoPagado = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro and cOpeCod not like '107[123456789]%')+(Select isnull(Sum(nMovImporte),0) From movOpeVarias Where nMovNro= MC.nMovNro),"
 sSql = sSql & " nCapital = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod in(1000,3000) and cOpeCod not like '107[123456789]%'),"
 sSql = sSql & " nInteres = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in (1100,1102,1103,1104,1105,3100)),"
 sSql = sSql & " nMora = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod in(1101,3101)),"
 sSql = sSql & " nGastos = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod like  '12%'),"
 sSql = sSql & " nITF = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in  (20, 21)),"
 sSql = sSql & " MC.nDiasMora , MC.nSaldoCap from MovCol MC Inner Join Mov M ON M.nMovNro = MC.nMovNro"
 sSql = sSql & " where MC.cOpeCod like '100%' AND MC.cCtaCod = '" & psctacod & "'  AND M.nMovFlag<>2 AND "
 sSql = sSql & "     (SUBSTRING(MC.cOpeCod,1,4) in ('1002','1003','1004','1005','1006','1007','1070')) AND "
 sSql = sSql & " (Select MAX(nNroCuota) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro) IS NOT NULL"
 sSql = sSql & " Order by (Select MAX(nNroCuota) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro)"
 
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaPagosRealizados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function UltimaCuotaCanceladaOriginal(ByVal psctacod As String) As Integer
Dim sSql As String
Dim oCon As DConecta
Dim R As ADODB.Recordset

    sSql = "Select ISNULL(MAX(nCuotaOrig),0) as nUltCuota From ColocCalifMiViv where cCtaCod = '" & psctacod & "' AND nColocCalendEstado = " & gColocCalendEstadoPagado
    Set oCon = New DConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    UltimaCuotaCanceladaOriginal = R!nUltCuota
    R.Close
    Set R = Nothing
End Function

Public Function SaldoPactadoCredito(ByVal psctacod As String) As Double
Dim sSql As String
Dim oCon As DConecta
Dim R As ADODB.Recordset

    sSql = "Select ISNULL(SUM(CD.nMonto),0) as nSaldo "
    sSql = sSql & " from ColocCalendDet CD"
    sSql = sSql & " Where CD.nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CD.cCtaCod )"
    sSql = sSql & " AND CD.nPrdConceptoCod = 1000 AND CD.nColocCalendApl = 1"
    sSql = sSql & " AND CD.cCtaCod = '" & psctacod & "'"

    Set oCon = New DConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    SaldoPactadoCredito = R!nSaldo
    R.Close
End Function

Public Function CuotaAprobada(ByVal psctacod As String) As Double
Dim sSql As String
Dim oCon As DConecta
Dim R As ADODB.Recordset

    
    sSql = "Select dbo.CuotaAprobada('" & psctacod & "') as nCuotaAprobada "
    Set oCon = New DConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    CuotaAprobada = IIf(IsNull(R!nCuotaAprobada), 0, R!nCuotaAprobada)
    
End Function

Public Function ExisteMetaAnalista(ByVal psPersCod As String, ByVal pdFechaIni As Date, _
    ByVal pdFechaFin As Date, ByVal pnMoneda As Integer, ByVal pnTipoMeta As Integer) As Boolean

Dim sSql As String
Dim oCon As DConecta
Dim R As ADODB.Recordset

    sSql = " Select * from ColocMetasAnalista "
    sSql = sSql & "  Where cPersCod = '" & psPersCod & "' AND nTipoMeta = " & pnTipoMeta
    sSql = sSql & "  AND dInicial <= '" & Format(pdFechaIni, "mm/dd/yyyy") & "' AND dFinal >= '" & Format(pdFechaIni, "mm/dd/yyyy") & "' "
    sSql = sSql & "  and nMoneda = " & pnMoneda
    
    
    
    Set oCon = New DConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    If R.RecordCount > 0 Then
        ExisteMetaAnalista = True
    Else
        ExisteMetaAnalista = False
    End If
    R.Close
End Function

Public Function CreditosCanceladoConDesembolso(ByVal psctacod As String, ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim oConec As DConecta

    sSql = "Select cCtaCod, nMonto from MovCol where nMovnro = " & pnMovNro & " and cCtaCod <> '" & psctacod & "' and copecod like '100[23456]%'"
    Set oConec = New DConecta
    oConec.AbreConexion
    Set CreditosCanceladoConDesembolso = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function PerteneceADesembolsoConCancelacion(ByVal pnMovNro As Long, ByVal psctacod As String) As Boolean
Dim oConecta As DConecta
Dim sSql As String
Dim R As ADODB.Recordset
    sSql = "Select cCtaCod,cOpeCod from MovCol where nMovnro = " & pnMovNro & " and cOpeCod in ('" & gCredDesembEfec & "','" & gCredDesembCtaNueva & "','" & gCredDesembCtaExist & "','" & gCredDesembCtaExistDOA & "','" & gCredDesembCtaNuevaDOA & "') and cCtaCod <> '" & psctacod & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If R.RecordCount > 0 Then
        PerteneceADesembolsoConCancelacion = True
    Else
        PerteneceADesembolsoConCancelacion = False
    End If
    R.Close
    Set R = Nothing
    
End Function

Public Function RecuperaCuentasAho(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConec As DConecta
    
    sSql = "Select cCodCtaAho from ColocCargoAutoma Where cCtaCod = '" & psctacod & "' "
    sSql = sSql & " AND cestado = '1' Order by nOrden "
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set RecuperaCuentasAho = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing

End Function

Public Function RecuperaCreditosCargoAutomatico(ByVal pdHoy As Date, ByVal pnMoneda As Moneda) As ADODB.Recordset
Dim sSql As String
Dim oConec As DConecta
    
    
    sSql = "        Select  ' ', C.cCtaCod, P.cPersNombre,  "
    sSql = sSql & "         nMonto = (Select SUM(nMonto-nMontoPagado) From ColocCalendDet CD"
    sSql = sSql & "                   Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen And nColocCalendApl = CC.nColocCalendApl "
    sSql = sSql & "                         AND nCuota = CC.nCuota), "
    sSql = sSql & "         Prd.nPrdEstado, C.cMetLiquidacion, CG.cCodCtaAho  "
    sSql = sSql & " from    Producto Prd Inner Join ColocacCred C ON Prd.cCtaCod = C.cCtaCod "
    sSql = sSql & "         INNER Join ColocCalendario CC ON CC.cCtaCod = C.cCtaCod AND CC.nNroCalen = C.nNroCalen"
    sSql = sSql & "         INNER Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "         INNER Join Persona P ON PP.cPersCod = P.cPersCod"
    sSql = sSql & "         INNER Join ColocCargoAutoma CG ON CG.cCtaCod = C.cCtaCod "
    sSql = sSql & " Where   CC.nColocCalendEstado = " & gColocCalendEstadoPendiente & " AND CC.nColocCalendApl = " & gColocCalendAplCuota & " AND CC.dVenc = '" & Format(pdHoy, "mm/dd/yyyy") & "' "
    sSql = sSql & "         AND SUBSTRING(C.cCtaCod,9,1)='" & pnMoneda & "' "
    sSql = sSql & " ORDER BY P.cPersNombre "
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set RecuperaCreditosCargoAutomatico = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
End Function

Public Function CodigosModulares(ByVal psPersCodTitular As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    sSql = "Select C.cCodModular From ColocacConvenio C Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular & " AND PP.cPersCod = '" & psPersCodTitular & "' Order by C.cCodModular"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set CodigosModulares = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function DefineCondicionCredito(ByVal psPersCod As String, Optional ByVal pnProducto As Producto = 0, _
Optional ByVal pdFecha As Date) As ColocCredCondicion
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConn As DConecta
Dim nCredNumVig As Integer
Dim nCredNumCancelados As Integer

    'Define la Condicion del Credito por Producto
    If pnProducto <> 0 Then
        sSql = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod Where PP.cPersCod = '" & psPersCod & "' "
        sSql = sSql & " AND PP.nPrdPersRelac = " & gColRelPersTitular
        sSql = sSql & " AND SUBSTRING(P.cCtaCod,6,2) = '" & Trim(Str(pnProducto)) & "' "
        sSql = sSql & " AND P.nPrdEstado in (" & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & ")"
        sSql = sSql & " AND SubString(PP.cCtaCod,6,3)<>'305'"
    Else
        sSql = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod  Where cPersCod = '" & psPersCod & "' "
        sSql = sSql & " AND PP.nPrdPersRelac = " & gColRelPersTitular
        sSql = sSql & " AND P.nPrdEstado in (" & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & ")"
        sSql = sSql & " AND SubString(PP.cCtaCod,6,3)<>'305'"
    End If
    Set oConn = New DConecta
    oConn.AbreConexion
    Set R = oConn.CargaRecordSet(sSql)
    nCredNumVig = IIf(IsNull(R!nTotal), 0, R!nTotal)
    R.Close
    If pnProducto <> 0 Then
        sSql = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod "
        sSql = sSql & " inner join colocaciones c on c.cctacod=pp.cctacod"
        sSql = sSql & " Where cPersCod = '" & psPersCod & "' "
        sSql = sSql & " AND PP.nPrdPersRelac = " & gColRelPersTitular
        sSql = sSql & " AND SUBSTRING(P.cCtaCod,6,2) = '" & Trim(Str(pnProducto)) & "' "
        sSql = sSql & " AND P.nPrdEstado not in (" & gColocEstSolic & "," & gColocEstSug & "," & gColocEstAprob & "," & gColocEstRetirado & "," & gColocEstRech & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & ")"
        sSql = sSql & " AND DATEDIFF(day,C.dVenc,'" & CStr(Format(pdFecha, "MM/dd/yyyy")) & "')<=360"
        sSql = sSql & " AND SubString(PP.cCtaCod,6,3)<>'305'"
    Else
        sSql = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod "
        sSql = sSql & " Inner Join Colocaciones C on C.cCtaCod=PP.cCtaCod"
        sSql = sSql & " Where cPersCod = '" & psPersCod & "' "
        sSql = sSql & " AND PP.nPrdPersRelac = " & gColRelPersTitular
        sSql = sSql & " AND P.nPrdEstado not in (" & gColocEstSolic & "," & gColocEstSug & "," & gColocEstAprob & "," & gColocEstRetirado & "," & gColocEstRech & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & ")"
        sSql = sSql & " AND DATEDIFF(day,C.dVenc,'" & CStr(Format(pdFecha, "MM/dd/yyyy")) & "')<=360"
        
    End If
    Set R = oConn.CargaRecordSet(sSql)
    nCredNumCancelados = IIf(IsNull(R!nTotal), 0, R!nTotal)
    oConn.CierraConexion
    Set oConn = Nothing
        
    'si no tiene ningun credito es normal
    If nCredNumVig = 0 And nCredNumCancelados = 0 Then
        DefineCondicionCredito = gColocCredCondNormal
    Else
        'si tiene creditos vigentes
        If nCredNumVig > 0 And nCredNumCancelados = 0 Or nCredNumVig > 0 And nCredNumCancelados > 0 Then
            DefineCondicionCredito = gColocCredCondParalelo
        Else
            DefineCondicionCredito = gColocCredCondRecurrente
        End If
    End If

End Function


Public Function VerificaCombinacionProd(ByVal psPersCod As String, Optional ByVal pnProducto As Producto = 0, _
Optional ByVal pbSustitucion As Boolean = False) As String
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConn As DConecta
Dim nCredNumProd As Integer
Dim nProductoCom As Integer
Dim sVerificaCombinacionProd As String
Dim nCheck As Boolean
Dim rs As ADODB.Recordset


    If pnProducto = 20 Then
       nProductoCom = 10
       sVerificaCombinacionProd = "El Titular presenta en su historia Créditos Comerciales"
       
       'Define la Condicion del Credito por Producto '
           sSql = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod Where PP.cPersCod = '" & psPersCod & "' "
           sSql = sSql & " AND PP.nPrdPersRelac = " & gColRelPersTitular
           sSql = sSql & " AND SUBSTRING(P.cCtaCod,6,2) = '" & Trim(Str(nProductoCom)) & "' "
           sSql = sSql & " AND P.nPrdEstado in(2020,2021,2022,2030,2032,2031,2032)"
        
       Set oConn = New DConecta
       oConn.AbreConexion
       Set R = oConn.CargaRecordSet(sSql)
       nCredNumProd = IIf(IsNull(R!nTotal), 0, R!nTotal)
       R.Close
       oConn.CierraConexion
       Set oConn = Nothing
           
     If pbSustitucion = False Then
            'si no tiene ningun credito c    on el producto Combinacion '
            If nCredNumProd = 0 Then
               VerificaCombinacionProd = ""
            Else 'si tiene creditos con el producto Combinacion '
               VerificaCombinacionProd = sVerificaCombinacionProd
            End If
      Else
          sSql = "Select nParamValor From ColocParametro Where nParamVar = 3102"
          Set oConn = New DConecta
          oConn.AbreConexion
          Set rs = oConn.CargaRecordSet(sSql)
          oConn.CierraConexion
          Set oConn = Nothing
          
          If Not rs.EOF And Not rs.BOF Then
             If rs!nParamValor = 1 Then
                nCheck = True
             Else
                nCheck = False
             End If
          End If
          Set rs = Nothing
           If nCheck = True Then ' tonces no se verifica el creditos de comercial a ames
                VerificaCombinacionProd = ""
           Else
                If nCredNumProd = 0 Then
                    VerificaCombinacionProd = ""
                Else 'si tiene creditos con el producto Combinacion '
                    VerificaCombinacionProd = sVerificaCombinacionProd
                 End If
           End If
      End If
    End If


End Function


Public Function RecuperaNivelesAprUsuario(ByVal psCodUser As String, ByVal pnMonto As Double, _
    psMoneda As String, ByVal psProducto As String, ByVal pbRefinanc As Boolean) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    
    sSql = " Select CN.cCodNiv, CA.cProduct, CA.cTpoCred, CA.nNivel, CA.nMontoMax, CA.nMontoMin"
    sSql = sSql & " From ColocCredPersNivelesApr CN Inner Join RRHH RH ON CN.cPersCod = RH.cPersCod"
    sSql = sSql & " Inner join ColocCredNivelesApr CA ON CN.cCodNiv = CA.cCodNiv"
    sSql = sSql & " Where RH.cUser = '" & psCodUser & "' AND CA.cProduct = '" & psProducto & "' AND CA.cTpoCred = '" & IIf(pbRefinanc, Trim(Str(gColocCredTipoRefinanciado)), Trim(Str(gColocCredTipoNormal))) & "' "
    sSql = sSql & " AND CA.nMontoMax >= " & Format(pnMonto, "#0.00") & " AND CA.nMontoMin <= " & Format(pnMonto, "#0.00") & " AND CA.nMoneda = " & psMoneda
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaNivelesAprUsuario = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaMovimientosAhorros(ByVal pnMov As Long, Optional ByVal pbSinMovRef As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConec As DConecta
Dim R As ADODB.Recordset

    Set oConec = New DConecta
    oConec.AbreConexion
    sSql = "Select * from Movref Where nMovNroRef = " & pnMov
    Set R = oConec.CargaRecordSet(sSql, adLockReadOnly)

    If R.RecordCount > 0 And Not pbSinMovRef Then
        sSql = "Select MC.nMovNro,MC.cOpeCod, MC.nMonto, MC.cCtaCod "
        'sSql = "Select MC.cOpeCod, SUM(MC.nMonto) nMonto "
        sSql = sSql & " From MovCap MC"
        sSql = sSql & " Inner Join MovRef MR ON MC.nMovNro = MR.nMovNro"
        sSql = sSql & " Where MR.nMovNroRef = " & pnMov
        sSql = sSql & " Order By MC.nMovNro "
        'sSql = sSql & " Group By MC.cOpeCod"
    Else
        sSql = "Select MC.nMovNro,MC.cOpeCod, MC.nMonto, MC.cCtaCod "
        'sSql = "Select MC.cOpeCod, SUM(MC.nMonto) nMonto "
        sSql = sSql & " From MovCap MC"
        sSql = sSql & " Where MC.nMovNro = " & pnMov
        sSql = sSql & " Order By MC.nMovNro "
        
    End If
    R.Close
    
    
    Set RecuperaMovimientosAhorros = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function RecuperaDatosExtorno(ByVal pnMovNro As Long, ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    sSql = "Select M.nCredEstado, SUM(MD.nMonto) as nCapital, "
    sSql = sSql & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
    sSql = sSql & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro AND MD.nPrdConceptoCod = 1000 Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psctacod & "' AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psctacod & "' and M.nCredEstado<>0  and MD.COPECOD LIKE '100[123456]%' and M.cOpeCod not Like '107[123456789]%'"
    sSql = sSql & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSql)
    
    If RecuperaDatosExtorno.RecordCount = 0 Then
        sSql = "Select M.nCredEstado, 0 as nCapital, "
        sSql = sSql & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
        sSql = sSql & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro AND MD.nPrdConceptoCod = 1100 Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psctacod & "' AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psctacod & "' and M.cOpeCod not Like '107[123456789]%'"
        sSql = sSql & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
        Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSql)
    End If
    
    If RecuperaDatosExtorno.RecordCount = 0 Then
        sSql = "Select M.nCredEstado, 0 as nCapital, "
        sSql = sSql & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
        sSql = sSql & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro AND MD.nPrdConceptoCod = 1101 Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psctacod & "' AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psctacod & "' and M.cOpeCod not Like '107[123456789]%'"
        sSql = sSql & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
        Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSql)
    End If
    
   ' Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSQL)
    If RecuperaDatosExtorno.RecordCount = 0 Then
        sSql = "Select M.nCredEstado, 0 as nCapital, "
        sSql = sSql & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
        sSql = sSql & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro AND Left(Cast(nPrdConceptoCod as Varchar(10)),2)='12' Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psctacod & "' AND Left(Cast(nPrdConceptoCod as Varchar(10)),2)='12' AND M.cCtaCod = '" & psctacod & "' and M.cOpeCod not Like '107[123456789]%'"
        sSql = sSql & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
        Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSql)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function


Public Function RecuperaDatosMetasAnalistaReporte(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaDatosMetasAnalistaReporte
    sSql = "Select DISTINCT M.cPersCod,M.dInicial,M.dFinal,M.nTipoMeta, C.cConsDescripcion as cTipoMeta, M.nMonto, M.nTipoCred, C2.cConsDescripcion cTipoCred, M.nMontoAlc, C3.cConsDescripcion as sMoneda "
    sSql = sSql & " From ColocMetasAnalista M Inner Join Constante C ON M.nTipoMeta = C.nConsValor AND C.nConsCod = " & gColocTipoMetas
    sSql = sSql & "                           Inner Join Constante C2 ON M.nTipoCred = C2.nConsValor  AND C2.nConsCod = 1001"
    sSql = sSql & "                           Inner Join Constante C3 ON C3.nConsValor = M.nMoneda AND C3.nConsCod = " & gMoneda
    sSql = sSql & " Inner Join ColocCredConsFiltro CF ON CF.nConsCod = C2.nConsCod AND CF.nCodFiltro =1 "
    sSql = sSql & " Where M.cPersCod = '" & psPersCod & "'  ORDER BY M.dInicial,cTipoMeta,sMoneda,nTipoCred"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosMetasAnalistaReporte = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosMetasAnalistaReporte:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaColocFteIngreso(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaColocacFteIngreso
    sSql = "Select * from ColocFteIngreso"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocFteIngreso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocacFteIngreso:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function


Public Function RecuperaDatosMetasAnalistaCab(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaDatosMetasAnalistaCab
    sSql = " Select M.cPersCod,M.dInicial,M.dFinal,M.nTipoMeta, C.cConsDescripcion as cTipoMeta, CN.cConsDescripcion as sMoneda, M.nMoneda "
    sSql = sSql & " From ColocMetasAnalista M Inner Join Constante C ON M.nTipoMeta = C.nConsValor AND C. nConsCod = " & gColocTipoMetas
    sSql = sSql & "     Inner Join Constante CN ON CN.nConsValor = M.nMoneda AND CN.nConsCod = " & gMoneda
    sSql = sSql & " Where M.cPersCod = '" & psPersCod & "'"
    sSql = sSql & "Group By M.cPersCod,M.dInicial,M.dFinal,M.nTipoMeta,C.cConsDescripcion,CN.cConsDescripcion,M.nMoneda "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosMetasAnalistaCab = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosMetasAnalistaCab:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function MontoColocadoCredCondicion(ByVal pdFechaIni As Date, _
    ByVal pdFechaFin As Date, ByVal pnMoneda As Integer, ByVal psProducto As String, _
    ByVal psPersCod As String, ByVal pnCondicion As ColocCredCondicion, Optional ByVal pnTodos As Boolean = False) As Double
Dim oCon As DConecta
Dim sSql As String
Dim R As ADODB.Recordset
    sSql = "Select SUM(C.nMontoCol) as nMontoCol from Colocaciones C "
    sSql = sSql & " Inner Join ColocacCred CC ON C.cCtaCod = CC.cCtaCod "
    sSql = sSql & " Inner Join ProductoPersona PP ON PP.cCtaCod = C.cCtaCod AND PP.nPrdPersRelac = 28"
    If pnTodos Then
        sSql = sSql & " Where SUBSTRING(C.cCtaCod,6,3) = '" & psProducto & "' "
    Else
        sSql = sSql & " Where CC.nColocCondicion = " & pnCondicion & " AND SUBSTRING(C.cCtaCod,6,3) = '" & psProducto & "' "
    End If
    sSql = sSql & " AND SUBSTRING(C.cCtaCod,9,1) = " & pnMoneda & " AND ( C.dVigencia >= '" & Format(pdFechaIni, "mm/dd/yyyy") & "' AND C.dVigencia <= '" & Format(pdFechaFin, "mm/dd/yyyy") & "') "
    sSql = sSql & " AND PP.cPersCod = '" & psPersCod & "'"
    
    Set oCon = New DConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    MontoColocadoCredCondicion = IIf(IsNull(R!nMontoCol), 0, R!nMontoCol)
    

End Function


Public Function RecuperaDatosMetasAnalistaDetalle(ByVal psPersCod As String, ByVal pnTipo As Integer, _
    ByVal pdInicio As Date, ByVal pdFinal As Date, ByVal pnMoneda As Integer) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaDatosMetasAnalistaDetalle
    
    'sSql = "Select nTipoCred,dInicial,dFinal,nMonto"
    'sSql = sSql & " From ColocMetasAnalista Where cPersCod = '" & psPersCod & "'"
    'sSql = sSql & " AND nTipoMeta = " & pnTipo & " AND  dInicial = '" & Format(pdInicio, "mm/dd/yyyy")
    'sSql = sSql & "' AND dFinal = '" & Format(pdFinal, "mm/dd/yyyy") & "' AND nMoneda = " & pnMoneda
    
    'sSql = "Select DISTINCT nTipoCred,dInicial,dFinal,nMonto"
    sSql = "Select DISTINCT nTipoCred,nMonto"
    sSql = sSql & " From ColocMetasAnalista Where cPersCod = '" & psPersCod & "'"
    sSql = sSql & " AND nTipoMeta = " & pnTipo & " AND  dInicial = '" & Format(pdInicio, "mm/dd/yyyy") & "' "
    sSql = sSql & " AND nMoneda = " & pnMoneda & " AND dFinal = '" & Format(pdFinal, "mm/dd/yyyy") & "' "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosMetasAnalistaDetalle = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaDatosMetasAnalistaDetalle:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function


Public Function RecuperaCreditosParaAsignarGasto(ByVal pnRanIni As Double, ByVal pnRanFin As Double, _
    ByVal pnAplicado As ColocCalendApl, ByVal psTipoprod As String, ByVal psMoneda As String, _
    ByVal psCodGasto As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCreditosParaAsignarGasto
    sSql = "Select distinct '',P.cCtaCod, RTRIM(Pers.cPersNombre), P.nSaldo"
    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                 Inner Join Persona Pers ON PP.cPersCod = Pers.cPersCod"
    sSql = sSql & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod"
    sSql = sSql & "                 Inner Join ColocCalendario CL ON P.cCtaCod = CL.cCtaCod AND CL.nNroCalen=CC.nNroCalen and CL.nColocCalendApl = " & pnAplicado & " AND CL.nColocCalendEstado = " & gColocCalendEstadoPendiente
    sSql = sSql & "                 Inner Join ColocCalendDet CD ON P.cCtaCod=CD.cCtaCod AND CD.nNroCalen = CC.nNroCalen AND CD.nColocCalendApl = " & pnAplicado
    sSql = sSql & "  WHERE substring(P.cCtaCod,6,1) = '" & psTipoprod & "' AND P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ")"
    sSql = sSql & "  AND substring(P.cCtaCod,9,1)='" & psMoneda & "' And nPrdConceptoCod <> " & CLng(psCodGasto)
    sSql = sSql & " Group By P.cCtaCod, Pers.cPersNombre, P.nSaldo, CD.nCuota, CC.nNroCalen "
    sSql = sSql & " HAVING SUM(CD.nMonto-isnull(CD.nMontopagado,0)) >= " & Format(pnRanIni, "#0.00") & " AND SUM(CD.nMonto-isnull(CD.nMontopagado,0)) <= " & Format(pnRanFin, "#0.00")
    sSql = sSql & " and CD.nCuota not in (select nCuota from ColocCalenddet where nCuota = CD.nCuota and nNroCalen = CC.nNrocalen And cCtaCod = P.cCtaCod AND nPrdConceptoCod =" & psCodGasto & ")"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosParaAsignarGasto = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosParaAsignarGasto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaCuotasParaAsignarGasto(ByVal pnRanIni As Double, ByVal pnRanFin As Double, _
    ByVal pnAplicado As ColocCalendApl, ByVal psctacod As String, ByVal psCodGasto As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCuotasParaAsignarGasto
    sSql = "Select distinct '',P.cCtaCod, Pers.cPersNombre, P.nSaldo, CD.nCuota, CC.nNroCalen, SUM(CD.nMonto-isnull(CD.nMontopagado,0)) as nMontoCuota "
    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                 Inner Join Persona Pers ON PP.cPersCod = Pers.cPersCod"
    sSql = sSql & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod"
    sSql = sSql & "                 Inner Join ColocCalendario CL ON P.cCtaCod = CL.cCtaCod AND CL.nNroCalen=CC.nNroCalen and CL.nColocCalendApl = " & pnAplicado & " AND CL.nColocCalendEstado = " & gColocCalendEstadoPendiente
    sSql = sSql & "                 Inner Join ColocCalendDet CD ON P.cCtaCod=CD.cCtaCod AND CD.nNroCalen = CC.nNroCalen AND CD.nColocCalendApl = " & pnAplicado
    sSql = sSql & "  WHERE P.cCtaCod = '" & psctacod & "'"
    sSql = sSql & " Group By P.cCtaCod, Pers.cPersNombre, P.nSaldo, CD.nCuota, CC.nNroCalen "
    sSql = sSql & " HAVING SUM(CD.nMonto-isnull(CD.nMontopagado,0)) >= " & Format(pnRanIni, "#0.00") & " AND SUM(CD.nMonto-isnull(CD.nMontopagado,0)) <= " & Format(pnRanFin, "#0.00")
    sSql = sSql & " and CD.nCuota not in (select nCuota from ColocCalenddet where nCuota = CD.nCuota and nNroCalen = CC.nNrocalen And cCtaCod = P.cCtaCod AND nPrdConceptoCod =" & psCodGasto & ")"
        
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCuotasParaAsignarGasto = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCuotasParaAsignarGasto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function



Public Function RecuperaTiposCredito() As ADODB.Recordset
Dim oConecta As DConecta
Dim sSql As String
    Set oConecta = New DConecta
    sSql = "Select nConsCod, nConsValor, cConsDescripcion from Constante where nConsCod = " & gProducto & " and (nConsValor % 100) = 0  order by  cConsDescripcion"
    oConecta.AbreConexion
    Set RecuperaTiposCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaCreditosParaJudicalTotal(Optional ByVal psPersCod As String = "", Optional psctacod As String = "") As ADODB.Recordset
Dim oParam As DParametro
Dim nValor As Double
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCreditosParaJudicalTotal
    Set oParam = New DParametro
    nValor = oParam.RecuperaValorParametro(gColocEstJudicial)
    Set oParam = Nothing
    
    If psPersCod <> "" Then
        sSql = "Select P.cCtaCod, C.nMontoCol, P.nSaldo, P.nTasaInteres, P.nPrdEstado, CC.nNroCalen, "
        sSql = sSql & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
        sSql = sSql & " cTitular=(Select cPersNombre From Persona Where cPersCod = PP2.cPersCod), CC.nDiasAtraso "
        sSql = sSql & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
        sSql = sSql & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
        sSql = sSql & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersAnalista
        sSql = sSql & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersTitular
        sSql = sSql & " WHERE PP2.cPersCod = '" & psPersCod & "'"
        sSql = sSql & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & ")"
        'sSql = sSql & " AND CC.nDiasAtraso > " & nValor
    Else
        If psctacod <> "" Then
            sSql = "Select P.cCtaCod, C.nMontoCol, P.nSaldo, P.nTasaInteres, P.nPrdEstado, CC.nNroCalen, "
            sSql = sSql & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
            sSql = sSql & " cTitular=(Select cPersNombre From Persona Where cPersCod = PP2.cPersCod), CC.nDiasAtraso "
            sSql = sSql & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
            sSql = sSql & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
            sSql = sSql & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersAnalista
            sSql = sSql & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersTitular
            sSql = sSql & " WHERE P.cCtaCod = '" & psctacod & "'"
            sSql = sSql & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & ")"
        Else
            sSql = "Select P.cCtaCod, C.nMontoCol, P.nSaldo, P.nTasaInteres, P.nPrdEstado, CC.nNroCalen, "
            sSql = sSql & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
            sSql = sSql & " cTitular=(Select cPersNombre From Persona Where cPersCod = PP2.cPersCod), CC.nDiasAtraso "
            sSql = sSql & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
            sSql = sSql & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
            sSql = sSql & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersAnalista
            sSql = sSql & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersTitular
            sSql = sSql & " WHERE CC.nDiasAtraso > " & nValor
            sSql = sSql & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & ")"
        End If
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosParaJudicalTotal = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaCreditosParaJudicalTotal:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaCreditosRefinanciados(ByVal psctacod As String, Optional ByVal pbSolicitados As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCreditosRefinanciados
    
    If Not pbSolicitados Then
        sSql = "Select Distinct CR.cCtaCodRef, "
        sSql = sSql & "nCapitalRef = (Select nMonto from ColocacRefinancDet Where cCtaCod = CR.cCtaCod AND cCtaCodRef = CR.cCtaCodRef AND nEstado = " & gColocEstadoRefinancAprobado & " AND nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
        sSql = sSql & "nInteresRef = (Select SUM(nMonto) from ColocacRefinancDet Where cCtaCod = CR.cCtaCod AND cCtaCodRef = CR.cCtaCodRef AND nEstado = " & gColocEstadoRefinancAprobado & " AND nPrdConceptoCod <> " & gColocConceptoCodCapital & ")"
        sSql = sSql & " from ColocacRefinanc CR "
        sSql = sSql & " WHERE CR.cCtaCod = '" & psctacod & "'"
    Else
        sSql = "Select Distinct CR.cCtaCodRef, "
        sSql = sSql & "nCapitalRef = (Select nMonto from ColocacRefinancDet Where cCtaCod = CR.cCtaCod AND cCtaCodRef = CR.cCtaCodRef AND nEstado = " & gColocEstadoRefinancSolicitado & " AND nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
        sSql = sSql & "nInteresRef = (Select SUM(nMonto) from ColocacRefinancDet Where cCtaCod = CR.cCtaCod AND cCtaCodRef = CR.cCtaCodRef AND nEstado = " & gColocEstadoRefinancSolicitado & " AND nPrdConceptoCod <> " & gColocConceptoCodCapital & ")"
        sSql = sSql & " from ColocacRefinanc CR "
        sSql = sSql & " WHERE CR.cCtaCod = '" & psctacod & "'"
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosRefinanciados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaCreditosRefinanciados:
                  Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaColocacRefinanc(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaColocacRefinanc
    sSql = "Select * from ColocacRefinanc where cCtacod = '" & psctacod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocacRefinanc = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaColocacRefinanc:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
Public Function RecuperaColocaciones(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaColocaciones
    sSql = "Select * from Colocaciones where cCtaCod = '" & psctacod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocaciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaColocaciones:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaGastosExonerados(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    sSql = "select * from ColocCredGastosExon Where cCtaCod = '" & psctacod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGastosExonerados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaDatosComunes(ByVal psctacod As String, Optional pnTodos As Boolean = True, Optional ByVal MatEst As Variant = Nothing) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim i As Integer

    On Error GoTo ErrorRecuperaDatosComunes
''''    sSQL = "Select P.nPrdEstado, CC.nCalendDinamico, CC.bCuotaCom, CL.nTasaIni as nTasaMora, CC.nCalendDinamTipo, CC.bPrepago, CC.bMiVivienda, P.cCtaCod, CE.nMonto nMontoSol, C.dVigencia,C.nMontoCol, CC.cMetLiquidacion, P.nSaldo, P.nTasaInteres, CC.nNroProxCuota, "
''''    sSQL = sSQL & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
''''    sSQL = sSQL & " Pers.cPersNombre cTitular, Pers.cPersCod,"
''''    sSQL = sSQL & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = Pers.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), "
''''    sSQL = sSQL & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = Pers.cPersCod AND cPersIDTpo = '" & gPersIdRUC & "'), "
''''    sSQL = sSQL & " CC.nExoPenalidad, nNota = (select nColocNota  From ColocCalificacionAnalista Where cTipo='A' AND cCtaCod = C.cCtaCod And dColocNotaFecha = (select Max(dColocNotaFecha) from ColocCalificacionAnalista Where cCtaCod = C.cCtaCod AND cTipo='A')), "
''''    sSQL = sSQL & " CN2.cConsDescripcion as cTipoCredDescrip, "
''''    sSQL = sSQL & " CN.cConsDescripcion as cDestinoDescripcion, CC.nDiasAtrasoAcum, "
''''    sSQL = sSQL & " CN3.cConsDescripcion as cMoneda, C.nMontoCol, CN4.cConsDescripcion cEstado, CN5.cConsDescripcion cCondicion, "
''''    sSQL = sSQL & " cLineaCred = (Select cLineaCred from Colocaciones Where cCtaCod = C.cCtaCod), "
''''    sSQL = sSQL & " dFechaSol = (Select dPrdEstado from ColocacEstado where cCtaCod = P.cCtaCod And nPrdEstado = " & gColocEstSolic & "), "
''''    sSQL = sSQL & " dbo.ColocNotaCredito(P.cCtaCod) as nNotaSist "
''''    sSQL = sSQL & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
''''    sSQL = sSQL & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
''''    sSQL = sSQL & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersAnalista
''''    sSQL = sSQL & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersTitular
''''    sSQL = sSQL & "                 Inner Join Persona Pers ON PP2.cPersCod = Pers.cPersCod"
''''    sSQL = sSQL & "                 Left Join Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = " & gProducto
''''    sSQL = sSQL & "                 Left Join Constante CN ON CC.nColocDestino = CN.nConsValor AND CN.nConsCod = " & Trim(Str(gColocDestino))
''''    sSQL = sSQL & "                 Left Join Constante CN3 ON convert(int,substring(C.cCtaCod,9,1)) = CN3.nConsValor AND CN3.nConsCod = " & gMoneda
''''    sSQL = sSQL & "                 Left Join Constante CN4 ON P.nPrdEstado = CN4.nConsValor AND CN4.nConsCod = " & gColocEstado
''''    sSQL = sSQL & "                 Left Join Constante CN5 ON CN5.nConsValor = CC.nColocCondicion AND CN5.nConsCod = " & gColocCredCondicion
''''    sSQL = sSQL & "                 Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado =" & gColocEstSolic
''''    sSQL = sSQL & "                 Left Join ColocLineaCreditoTasa CL ON CL.cLineaCred = C.cLineaCred AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal
''''    sSQL = sSQL & " WHERE P.cCtaCod = '" & psCtaCod & "' AND CC.nCalendDinamico = 0 " 'AND P.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
''''    If Not pnTodos Then
''''        sSQL = sSQL & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & ")"
''''    Else
''''        If IsArray(MatEst) Then
''''            sSQL = sSQL & " AND P.nPrdEstado in ("
''''            For I = 0 To UBound(MatEst)
''''                sSQL = sSQL & MatEst(I) & ","
''''            Next I
''''            sSQL = Mid(sSQL, 1, Len(sSQL) - 1) & ")"
''''        End If
''''    End If

    sSql = "Select P.nPrdEstado, CC.nCalendDinamico, CC.bCuotaCom, CL.nTasaIni as nTasaMora, CC.nCalendDinamTipo, CC.bPrepago, CC.bMiVivienda, P.cCtaCod, CE.nMonto nMontoSol, C.dVigencia,C.nMontoCol, CC.cMetLiquidacion, P.nSaldo, P.nTasaInteres, CC.nNroProxCuota, "
    sSql = sSql & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
    sSql = sSql & " Pers.cPersNombre cTitular, Pers.cPersCod,"
    sSql = sSql & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = Pers.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), "
    sSql = sSql & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = Pers.cPersCod AND cPersIDTpo = '" & gPersIdRUC & "'), "
    sSql = sSql & " CC.nExoPenalidad, nNota = (select nColocNota  From ColocCalificacionAnalista Where cTipo='A' AND cCtaCod = C.cCtaCod And dColocNotaFecha = (select Max(dColocNotaFecha) from ColocCalificacionAnalista Where cCtaCod = C.cCtaCod AND cTipo='A')), "
    sSql = sSql & " CC.nExoPenalidad, nNota = (select nColocNota  From ColocCalificacionAnalista Where cTipo='A' AND cCtaCod = C.cCtaCod And dColocNotaFecha = (select Max(dColocNotaFecha) from ColocCalificacionAnalista Where cCtaCod = C.cCtaCod AND cTipo='A')), "
    sSql = sSql & " CN2.cConsDescripcion as cTipoCredDescrip, "
    sSql = sSql & " CN.cConsDescripcion as cDestinoDescripcion, CC.nDiasAtrasoAcum, "
    sSql = sSql & " CN3.cConsDescripcion as cMoneda, C.nMontoCol, CN4.cConsDescripcion cEstado, CN5.cConsDescripcion cCondicion, "
    sSql = sSql & " cLineaCred = (Select cLineaCred from Colocaciones Where cCtaCod = C.cCtaCod), "
    sSql = sSql & " dFechaSol = (Select dPrdEstado from ColocacEstado where cCtaCod = P.cCtaCod And nPrdEstado = " & gColocEstSolic & "), "
    sSql = sSql & " dbo.ColocNotaCredito(P.cCtaCod) as nNotaSist "
    sSql = sSql & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
    sSql = sSql & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
    sSql = sSql & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                 Inner Join Persona Pers ON PP2.cPersCod = Pers.cPersCod"
    sSql = sSql & "                 Left Join Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = " & gProducto
    sSql = sSql & "                 Left Join Constante CN ON CC.nColocDestino = CN.nConsValor AND CN.nConsCod = " & Trim(Str(gColocDestino))
    sSql = sSql & "                 Left Join Constante CN3 ON convert(int,substring(C.cCtaCod,9,1)) = CN3.nConsValor AND CN3.nConsCod = " & gMoneda
    sSql = sSql & "                 Left Join Constante CN4 ON P.nPrdEstado = CN4.nConsValor AND CN4.nConsCod = " & gColocEstado
    sSql = sSql & "                 Left Join Constante CN5 ON CN5.nConsValor = CC.nColocCondicion AND CN5.nConsCod = " & gColocCredCondicion
    sSql = sSql & "                 Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado =" & gColocEstSolic
    sSql = sSql & "                 Left Join ColocLineaCreditoTasa CL ON CL.cLineaCred = C.cLineaCred AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal
    sSql = sSql & " WHERE P.cCtaCod = '" & psctacod & "' " 'AND P.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
    If Not pnTodos Then
        sSql = sSql & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & ")"
    Else
        If IsArray(MatEst) Then
            sSql = sSql & " AND P.nPrdEstado in ("
            For i = 0 To UBound(MatEst)
                sSql = sSql & MatEst(i) & ","
            Next i
            sSql = Mid(sSql, 1, Len(sSql) - 1) & ")"
        End If
    End If


    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosComunes = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ErrorRecuperaDatosComunes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaCreditosVigentesRelacCred(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCreditosVigentesRelacCred
    sSql = "Select  Pers.cPersNombre, CN.cConsDescripcion cRelacion, P.cCtaCod, CC.nDiasAtraso"
    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac not in (" & gColRelPersAnalista & "," & gColRelPersApoderado & "," & gColRelPersTitular & ")"
    sSql = sSql & "                 Inner Join Constante CN ON CN.nConsValor = PP.nPrdPersRelac AND CN.nConsCod = " & gColocRelacPers
    sSql = sSql & "                 Inner join Persona Pers ON Pers.cPersCod = PP.cPersCod "
    sSql = sSql & "                 Inner join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    sSql = sSql & " Where Pers.cPersCod In (Select cPersCod from ProductoPersona Where cCtaCod = '" & psctacod & "')"
    sSql = sSql & "AND P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ") "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosVigentesRelacCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosVigentesRelacCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description


End Function

Public Function RecuperaEstadosdelCredito(ByVal psctacod As String, ByVal pbSolicitud As Boolean, _
    ByVal pbSugerencia As Boolean, ByVal pbAprobacion As Boolean) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaEstadosdelCredito
    sSql = ""
    If pbSolicitud Then
        sSql = sSql & "Select CE.nPrdEstado, CN.cConsDescripcion cEstado, CE.nMonto, CE.nCuotas, CE.nPlazo, NULL as nMontoCuota "
        sSql = sSql & " From ColocacEstado CE Inner Join Constante CN ON CE.nPrdEstado = CN.nConsValor AND CN.nConsCod = " & gColocEstado
        sSql = sSql & " Where CE.cCtaCod = '" & psctacod & "' AND CE.nPrdEstado = " & gColocEstSolic
    End If
    If pbSugerencia Then
        If Len(sSql) > 0 Then
            sSql = sSql & " Union "
        End If
        sSql = sSql & "Select CE.nPrdEstado,CN.cConsDescripcion cEstado, CE.nMonto, CE.nCuotas, CE.nPlazo, "
        sSql = sSql & " nMontoCuota = (Select SUM(CD.nMonto) from ColocCalendDet CD where nNrocalen=1 and cCtaCod = CE.cCtaCod AND nColocCalendApl = 1 AND nCuota = 1) "
        sSql = sSql & " From ColocacEstado CE Inner Join Constante CN ON CE.nPrdEstado = CN.nConsValor AND CN.nConsCod = " & gColocEstado
        sSql = sSql & " Where CE.cCtaCod = '" & psctacod & "' AND  CE.nPrdEstado = " & gColocEstSug
    End If
    If pbAprobacion Then
        If Len(sSql) > 0 Then
            sSql = sSql & " Union "
        End If
        sSql = sSql & "Select CE.nPrdEstado, CN.cConsDescripcion cEstado, CE.nMonto, CE.nCuotas, CE.nPlazo, "
        sSql = sSql & " nMontoCuota = (Select SUM(CD.nMonto) from ColocCalendDet CD where nNrocalen=(Select nNroCalen From ColocacCred Where cCtaCod = CE.cCtaCod) and cCtaCod = CE.cCtaCod AND nColocCalendApl = 1 AND nCuota = 1) "
        sSql = sSql & " From ColocacEstado CE Inner Join Constante CN ON CE.nPrdEstado = CN.nConsValor AND CN.nConsCod = " & gColocEstado
        sSql = sSql & " Where CE.cCtaCod = '" & psctacod & "' AND  CE.nPrdEstado = " & gColocEstAprob
    End If
    sSql = sSql & " Order By Ce.nprdEstado"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaEstadosdelCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaEstadosdelCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaColocacEstado(ByVal psctacod As String, ByVal pnEstado As ColocEstado, Optional pbTodos As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaColocacEstado
    If Not pbTodos Then
        sSql = "Select * from ColocacEstado Where cCtaCod = '" & psctacod & "' AND nPrdEstado =  " & pnEstado
    Else
        sSql = "Select * from ColocacEstado Where cCtaCod = '" & psctacod & "' Order By nPrdEstado"
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocacEstado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocacEstado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
    
End Function

Public Function RecuperaMontoGarantiaCredito(ByVal psctacod As String, ByVal pdFecha As Date, Optional ByVal bAmpliado As Boolean) As Double
Dim sSql As String
Dim oConecta As DConecta
Dim oGeneral As DGeneral
Dim nTipoCambioFijo As Double
Dim nMontoMN As Double
Dim nMontoME As Double
Dim oAmpliado As DAmpliacion
Dim R As ADODB.Recordset
Dim cCtaCodOld As String

    On Error GoTo ErrorRecuperaMontoGarantiaCredito
    Set oGeneral = New DGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
    nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))
    Set oGeneral = Nothing
    
    'Moneda Nacional
    If bAmpliado = True Then
        ' se verifica que si el un credito ampliado o no?
        Set oAmpliado = New DAmpliacion
        
'            cCtaCodOld = oAmpliado.GetcCtaCodPorAmpliar(psCtaCod)
'            nMontoMN = oAmpliado.EmularGarantia(cCtaCodOld, gMonedaNacional)
'            nMontoMN = Format(nMontoMN, "#0.00")
             nMontoMN = oAmpliado.ObtenerMontoGarantAmpliado(psctacod, gMonedaNacional)
             Set oAmpliado = Nothing
             nMontoMN = Format(nMontoMN, "#0.00")
    End If
    
    
        sSql = "Select sum(nGravado) as nTotal from ColocGarantia Where cCtaCod = '" & psctacod & "' And nMoneda = " & gMonedaNacional
    
    
        Set oConecta = New DConecta
        oConecta.AbreConexion
        Set R = oConecta.CargaRecordSet(sSql)
        If Not R.BOF And Not R.EOF Then
            nMontoMN = IIf(IsNull(R!nTotal), 0, R!nTotal) + nMontoMN
            nMontoMN = CDbl(Format(nMontoMN, "#0.00"))
        End If
        R.Close
        Set R = Nothing
   
   
    'Moneda Extrangera
    If bAmpliado = True Then
        ' se verifica que si el un credito ampliado o no?
        Set oAmpliado = New DAmpliacion
            'cCtaCodOld = oAmpliado.GetcCtaCodPorAmpliar(psCtaCod)
           ' nMontoME = oAmpliado.EmularGarantia(cCtaCodOld, gMonedaExtranjera)
            nMontoME = oAmpliado.ObtenerMontoGarantAmpliado(psctacod, gMonedaExtranjera)
            nMontoME = Format(nMontoME, "#0.00")
        Set oAmpliado = Nothing
    End If
    
        sSql = "Select sum(nGravado) as nTotal from ColocGarantia Where cCtaCod = '" & psctacod & "' And nMoneda = " & gMonedaExtranjera
        Set R = oConecta.CargaRecordSet(sSql)
        If Not R.BOF And Not R.EOF Then
            nMontoME = IIf(IsNull(R!nTotal), 0, R!nTotal) + nMontoME
            nMontoME = CDbl(Format(nMontoME, "#0.00"))
        End If
        R.Close
        Set R = Nothing
        oConecta.CierraConexion
    
    
    Set oConecta = Nothing
    If CInt(Mid(psctacod, 9, 1)) = gMonedaNacional Then
        RecuperaMontoGarantiaCredito = 0
        If nMontoMN > 0 Then
            RecuperaMontoGarantiaCredito = RecuperaMontoGarantiaCredito + nMontoMN
        End If
        If nMontoME > 0 Then
            RecuperaMontoGarantiaCredito = RecuperaMontoGarantiaCredito + (nMontoME * nTipoCambioFijo)
        End If
    End If
    
    If CInt(Mid(psctacod, 9, 1)) = gMonedaExtranjera Then
        RecuperaMontoGarantiaCredito = 0
        If nMontoMN > 0 Then
            RecuperaMontoGarantiaCredito = RecuperaMontoGarantiaCredito + (nMontoMN / nTipoCambioFijo)
        End If
        If nMontoME > 0 Then
            RecuperaMontoGarantiaCredito = RecuperaMontoGarantiaCredito + nMontoME
        End If
    End If
    
    Exit Function

ErrorRecuperaMontoGarantiaCredito:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaDatosPrepago(ByVal psctacod As String, ByVal pdHoy As Date) As Double
Dim sSql As String
Dim oConec As DConecta
Dim R As ADODB.Recordset

    sSql = " Select SUM(MCD.nMonto) as nMontoPagado from MovCol MC Inner Join Mov M ON MC.nMovNro = M.nMovNro "
    sSql = sSql & " Inner Join MovColDet MCD ON MC.nMovNro = MCD.nMovNro "
    sSql = sSql & " Where MC.nPrepago = 1 AND MC.cCtaCod = '" & psctacod & "' AND "
    sSql = sSql & " M.cMovnro like '" & Year(pdHoy) & "%' "
    sSql = sSql & " AND MCD.nPrdConceptoCod = 1000 AND MCD.cOpeCod not in ('" & gCredDesembEfec & "','" & gCredDesembCtaNueva & "','" & gCredDesembCtaExist & "','" & gCredDesembCtaExistDOA & "','" & gCredDesembCtaNuevaDOA & "')"
    sSql = sSql & " AND MC.nNroCalen = (select nNroCalen From ColocacCred Where cCtaCod = MC.cCtaCod)"
    Set oConec = New DConecta
    oConec.AbreConexion
    Set R = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    RecuperaDatosPrepago = Format(IIf(IsNull(R!nMontoPagado), 0, R!nMontoPagado), "0.00")
    
End Function
Public Function RecuperaDatosCreditoVigente(ByVal psctacod As String, Optional ByVal pdHoy As Date = CDate("01/01/1900"), _
    Optional ByVal pbIncluirbAprobados As Boolean = False) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaDatosCreditoVigente
    sSql = "Select PP.cPersCod, CC.cMetLiquidacion, dbo.CuotaAprobada(P.cCtaCod) as CuotaAprobada, CC.bCuotaCom, CC.bMiVivienda, CC.nCalendDinamTipo, CC.bPrepago, CC.nCalPago, Pers.cPersNombre, C.cLineaCred,C.nMontoCol, P.nSaldo, "
    sSql = sSql & " cMoneda = (Select cConsDescripcion from Constante where nConsCod = " & gMoneda & " AND nConsValor = " & CInt(Mid(psctacod, 9, 1)) & "),"
    sSql = sSql & " CE.nCuotas nCuotasApr, CE.nColocCalendCod " & ", "
    sSql = sSql & " CC.nDiasAtraso, CC.cMetLiquidacion, P.nTransacc, CC.nCalendDinamico, CC.nIntPend, CC.nNroCalen, CC.nNroProxCuota, CC.nNroProxDesemb "
    If pdHoy <> CDate("01/01/1900") Then
        sSql = sSql & ", DATEDIFF(year,C.dVigencia,'" & Format(pdHoy, "mm/dd/yyyy") & "') as nPlazoTranscurrido "
    End If
    sSql = sSql & " From Producto P Inner join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                 Inner Join Persona Pers ON PP.cPersCod = Pers.cPersCod "
    sSql = sSql & "                 Inner Join Colocaciones C ON C.cCtaCod =  P.cCtaCod "
    sSql = sSql & "                 Inner Join ColocacCred CC ON CC.cCtaCod =  P.cCtaCod "
    sSql = sSql & "                 Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSql = sSql & " WHERE P.cCtaCod = '" & psctacod & "'"
    
    If pbIncluirbAprobados Then
        sSql = sSql & " AND P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstAprob & ")"
    Else
        sSql = sSql & " AND P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ")"
    End If
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosCreditoVigente = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaDatosCreditoVigente:
    Err.Raise Err.Number, "Error En Proceso RecuperaDatosCreditoVigente", Err.Description
        
End Function

Public Sub ActualizaCreditoAnalista(ByVal psctacod As String, ByVal psPersCodAnalista As String)
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorActualizaCreditoAnalista
    Set oConecta = New DConecta
    sSql = "UPDATE ProductoPersona SET cPersCod = '" & psPersCodAnalista & "' "
    sSql = sSql & " Where nPrdPersRelac = " & Trim(Str(gColRelPersAnalista)) & " AND cCtaCod = '" & psctacod & "'"
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorActualizaCreditoAnalista:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso ActualizaCreditoAnalista", Err.Description
    
End Sub
Public Function CarteraInstitucion(ByVal psPersCod As String)
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorCarteraInstitucion
    sSql = "Select C.cCtaCod, P.cPersNombre from ColocacConvenio C inner join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod and PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "             Inner Join Persona P ON PP.cPersCod = P.cPersCod"
    sSql = sSql & " where C.cPersCod = '" & Trim(psPersCod) & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set CarteraInstitucion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorCarteraInstitucion:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function CarteraAnalista(ByVal psPersCod As String, Optional ByVal pCodAgen As String, _
Optional ByVal psCodUbiGeografica As String = "", Optional ByVal psCodTipoProducto As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorCarteraAnalista
    
    Set oConecta = New DConecta
    sSql = "Select C.cCtaCod, PER.cPersNombre "
    sSql = sSql & " from    ColocacCred C inner join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod "
    sSql = sSql & " inner join Producto PRD ON C.cCtaCod = PRD.cCtaCod "
    sSql = sSql & " inner join ProductoPersona PP2 ON C.cCtaCod = PP2.cCtaCod "
    sSql = sSql & " inner join Persona PER ON PP2.cPersCod = PER.cPersCod "
    
    If Len(psCodUbiGeografica) > 0 Then
        If Mid(psCodUbiGeografica, 1, 1) = "3" Then
            sSql = sSql & " Inner Join UbicacionGeoGrafica U on SubString(U.cUbiGeoCod,1,1)='3' and SubString(U.cUbiGeoCod,2,6)=SubString(Per.cPersDireccUbiGeo,2,6) and Substring(U.cUbiGeoCod,8,5)='00000'"
        Else
            sSql = sSql & "Inner Join UbicacionGeoGrafica U on U.cUbiGeoCod=Per.cPersDireccUbiGeo"
        End If
    End If
    
    sSql = sSql & " Where PP.nPrdPersRelac = 28 and PRD.nPrdEstado IN('2020','2021','2022','2030','2031','2032','2201','2205') and PP2.nPrdPersRelac = " & gColRelPersTitular & " and PP.cPersCod = '" & psPersCod & "'"
    sSql = sSql & " and substring(prd.cctacod,4,2)='" & pCodAgen & "'"
    
    If Len(psCodUbiGeografica) > 0 Then
        sSql = sSql & " and U.cUbiGeoCod='" & psCodUbiGeografica & "'"
    End If
    
    If Len(psCodTipoProducto) > 0 Then
       sSql = sSql & " and SubString(Prd.cCtaCod,6,3)='" & psCodTipoProducto & "'"
    End If
    oConecta.AbreConexion
    Set CarteraAnalista = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorCarteraAnalista:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    Set oConecta = Nothing
    
End Function
Public Sub EliminaSolicitud(ByVal psctacod As String)
Dim sSql As String
Dim Conn As DConecta
Dim bTran As Boolean

    On Error GoTo ErrorEliminaSolicitud
    bTran = False
    Set Conn = New DConecta
    Conn.AbreConexion
    Conn.ConexionActiva.BeginTrans
    bTran = True
    
    'Elimina ProductoPersona
    sSql = "Delete ProductoPersona Where cCtaCod = '" & psctacod & "'"
    Conn.ConexionActiva.Execute sSql
    
    'Elimina ColocEstado
    sSql = "Delete ColocacEstado Where cCtaCod = '" & psctacod & "'"
    Conn.ConexionActiva.Execute sSql
    
    'Elimina ColocFteIngreso
    sSql = "Delete ColocFteIngreso Where cCtaCod = '" & psctacod & "'"
    Conn.ConexionActiva.Execute sSql
    
    'Elimina ColocacConvenio
    sSql = "Delete ColocacConvenio Where cCtaCod = '" & psctacod & "'"
    Conn.ConexionActiva.Execute sSql
    
    'Elimina ColocacCred
    sSql = "Delete ColocacCred Where cCtaCod = '" & psctacod & "'"
    Conn.ConexionActiva.Execute sSql
    
    'Elimina Colocaciones
    sSql = "Delete Colocaciones Where cCtaCod = '" & psctacod & "'"
    Conn.ConexionActiva.Execute sSql
    
    'Elimina Producto
    sSql = "Delete Producto Where cCtaCod = '" & psctacod & "'"
    Conn.ConexionActiva.Execute sSql
    
    'ELimina si es credito ampliado
    sSql = "Delete ColocacAmpliado Where cCtaCod='" & psctacod & "'"
    Conn.ConexionActiva.Execute sSql
    
    Conn.ConexionActiva.CommitTrans
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Sub

ErrorEliminaSolicitud:
    If bTran Then
        Conn.RollbackTrans
        Set Conn = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso EliminaSolicitud", Err.Description
End Sub
Public Sub ActualizaSolicitud(ByVal oRelPersCred As UCredRelacion, ByVal psctacod As String, _
                          ByVal psCondCred As String, ByVal psCondCredOtra As String, ByVal popersona As DPersona, ByVal psFuenteIng As String, ByVal pnMontoSol As Double, _
                          ByVal pnCuotasSol As Integer, ByVal pnPlazo As Integer, ByVal psDestCred As String, _
                          ByVal psPersCodAnalista As String, ByVal psPersCodInst As String, ByVal psCodModular As String, _
                          ByVal dFecSol As Date, ByVal psMovAct As String, ByVal pdFIFecEval As Date, ByVal bRefinanciar As Boolean, ByVal pnCondProd As Integer, Optional ByVal MatCredRef As Variant, Optional pnCapitalInt As Boolean = False, Optional bSustiDeudor As Boolean = False, Optional ByVal nCampanaCod As Integer)

Dim bTran As Boolean
Dim i As Integer
Dim sSql As String
Dim Conn As DConecta
Dim oBase As DCredActualizaBD

    On Error GoTo ErrorActualizaSolicitud
    
    bTran = False
    Set Conn = New DConecta
    Conn.AbreConexion
    Conn.ConexionActiva.BeginTrans
    bTran = True
    
    'Actualiza Producto
    sSql = "UPDATE Producto SET "
    sSql = sSql & " nPrdEstado = " & Trim(Str(gColocEstSolic)) & ", "
    sSql = sSql & " dPrdEstado = '" & Format(dFecSol, "mm/dd/yyyy") & "'"
    sSql = sSql & " WHERE cCtaCod = '" & psctacod & "'"
    Conn.ConexionActiva.Execute sSql
    
    'Actualiza Nuevas Personas Relacionadas con el Credito
    oRelPersCred.IniciarMatriz
    
    Do While Not oRelPersCred.EOF
            sSql = "DELETE ProductoPersona "
            sSql = sSql & " WHERE cCtaCod = '" & psctacod & "' AND cPersCod = '" & oRelPersCred.ObtenerCodigo & "' AND nPrdPersRelac = " & oRelPersCred.ObtenerValorRelacAnt
            Conn.ConexionActiva.Execute sSql
            
            sSql = "INSERT INTO ProductoPersona(cCtaCod,cPersCod,nPrdPersRelac) "
            sSql = sSql & " VALUES('" & psctacod & "','" & oRelPersCred.ObtenerCodigo & "'," & oRelPersCred.ObtenerValorRelac & ")"
            Conn.ConexionActiva.Execute sSql
            
        oRelPersCred.siguiente
    Loop
    
    'Actualiza Analista
    sSql = "UPDATE ProductoPersona SET "
    sSql = sSql & " cPersCod = '" & psPersCodAnalista & "' "
    sSql = sSql & " WHERE cCtaCod = '" & psctacod & "' AND nPrdPersRelac = " & gColRelPersAnalista
    
    Conn.ConexionActiva.Execute sSql
    
    'Actualiza Registro de Colocaciones
    sSql = "UPDATE Colocaciones SET "
    sSql = sSql & " nPlazo = " & Trim(Str(pnPlazo)) & ","
    sSql = sSql & " nMontoCol = " & Format(pnMontoSol, "#0.00") & ","
    sSql = sSql & " cUltimaActualizacion = '" & psMovAct & "' "
    sSql = sSql & " WHERE cCtaCod = '" & psctacod & "'"
    Conn.ConexionActiva.Execute sSql
    
    'Actualiza Registro de ColocacCred pnCondProd
    sSql = "UPDATE ColocacCred SET "
    sSql = sSql & " nColocCondicion = " & psCondCred & ", "
    sSql = sSql & " nColocCondicionProd = " & pnCondProd & ", "
    sSql = sSql & " nColocDestino = '" & psDestCred & "', "
    sSql = sSql & " bRefCapInt = " & IIf(pnCapitalInt, "1 ", "0 ") & ","
    sSql = sSql & " IdCampana=" & IIf(nCampanaCod = 0, "NULL", nCampanaCod)
    sSql = sSql & " WHERE cCtaCod = '" & psctacod & "'"
        
    Conn.ConexionActiva.Execute sSql
    
    'Actualiza Registro en ColocEstado
    Set oBase = New DCredActualizaBD
    sSql = "UPDATE ColocacEstado SET "
    sSql = sSql & " dPrdEstado = '" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
    sSql = sSql & " nCuotas = " & Trim(Str(pnCuotasSol)) & ","
    sSql = sSql & " nMonto = " & Format(pnMontoSol, "#0.00") & ","
    sSql = sSql & " nPlazo = " & Trim(Str(pnPlazo))
    sSql = sSql & " WHERE cCtaCod = '" & psctacod & "' AND nPrdEstado = " & Trim(Str(gColocEstSolic))
    Set oBase = Nothing
    Conn.ConexionActiva.Execute sSql
    
    'Actualiza Registro en ColocFteIngreso
            
    sSql = "UPDATE ColocFteIngreso SET "
    sSql = sSql & " cNumFuente = '" & psFuenteIng & "', "
    sSql = sSql & " dPersEval = '" & Format(pdFIFecEval, "mm/dd/yyyy") & "' "
    sSql = sSql & " WHERE cCtaCod = '" & psctacod & "' AND cNumFuente = '" & psFuenteIng & "'"
    Conn.ConexionActiva.Execute sSql
    
    'Nuevo Registro de ColocConvenio
    If Len(Trim(psPersCodInst)) > 0 Then
        sSql = "DELETE ColocacConvenio WHERE cCtaCod = '" & psctacod & "' AND cPersCod = '" & psPersCodInst & "'"
        Conn.ConexionActiva.Execute sSql
        
        sSql = "INSERT INTO ColocacConvenio(cCtaCod,cPersCod,cCodModular)"
        sSql = sSql & " VALUES('" & psctacod & "',"
        sSql = sSql & "'" & psPersCodInst & "',"
        sSql = sSql & "'" & psCodModular & "')"
        Conn.ConexionActiva.Execute sSql
    End If
    
    'Si es una Refinanciancion
    'Si es una Refinanciacion Guardar los Datos
    If bRefinanciar Then
        If IsArray(MatCredRef) Then
            Set oBase = New DCredActualizaBD
            If UBound(MatCredRef) > 0 Then
                sSql = "Delete ColocacRefinancDet where cCtaCod = '" & psctacod & "'"
                Conn.ConexionActiva.Execute sSql
                sSql = "Delete ColocacRefinanc where cCtaCod = '" & psctacod & "'"
                Conn.ConexionActiva.Execute sSql
                For i = 0 To UBound(MatCredRef) - 1
                    sSql = "INSERT INTO ColocacRefinanc(cCtaCod,cCtaCodRef,nEstado,dEstado,nMontoRef)"
                    sSql = sSql & " VALUES('" & psctacod & "','"
                    sSql = sSql & MatCredRef(i, 0) & "',"
                    sSql = sSql & gColocEstadoRefinancSolicitado & ","
                    sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                    sSql = sSql & MatCredRef(i, 1) & ")"
                    Conn.ConexionActiva.Execute sSql
                    
                    'Ingreso del Detalle
                    'Capital
                    sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                    sSql = sSql & " VALUES('" & psctacod & "','"
                    sSql = sSql & MatCredRef(i, 0) & "',"
                    sSql = sSql & gColocEstadoRefinancSolicitado & ","
                    sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                    sSql = sSql & gColocConceptoCodCapital & ","
                    sSql = sSql & MatCredRef(i, 2) & ","
                    sSql = sSql & "0.00)"
                    Conn.ConexionActiva.Execute sSql
                    'Interes Compensatorio
                    If CDbl(MatCredRef(i, 3)) > 0 Then
                        sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSql = sSql & " VALUES('" & psctacod & "','"
                        sSql = sSql & MatCredRef(i, 0) & "',"
                        sSql = sSql & gColocEstadoRefinancSolicitado & ","
                        sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                        sSql = sSql & gColocConceptoCodInteresCompensatorio & ","
                        sSql = sSql & MatCredRef(i, 3) & ","
                        sSql = sSql & "0.00)"
                        Conn.ConexionActiva.Execute sSql
                    End If
                    
                    'Interes Moratorio
                    If CDbl(MatCredRef(i, 4)) > 0 Then
                        sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSql = sSql & " VALUES('" & psctacod & "','"
                        sSql = sSql & MatCredRef(i, 0) & "',"
                        sSql = sSql & gColocEstadoRefinancSolicitado & ","
                        sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                        sSql = sSql & gColocConceptoCodInteresMoratorio & ","
                        sSql = sSql & MatCredRef(i, 4) & ","
                        sSql = sSql & "0.00)"
                        Conn.ConexionActiva.Execute sSql
                    End If
                    
                    'Interes Gracia
                    If CDbl(MatCredRef(i, 5)) > 0 Then
                        sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSql = sSql & " VALUES('" & psctacod & "','"
                        sSql = sSql & MatCredRef(i, 0) & "',"
                        sSql = sSql & gColocEstadoRefinancSolicitado & ","
                        sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                        sSql = sSql & gColocConceptoCodInteresGracia & ","
                        sSql = sSql & MatCredRef(i, 5) & ","
                        sSql = sSql & "0.00)"
                        Conn.ConexionActiva.Execute sSql
                    End If
                    
                    'Interes Suspenso
                    If CDbl(MatCredRef(i, 6)) > 0 Then
                        sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSql = sSql & " VALUES('" & psctacod & "','"
                        sSql = sSql & MatCredRef(i, 0) & "',"
                        sSql = sSql & gColocEstadoRefinancSolicitado & ","
                        sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                        sSql = sSql & gColocConceptoCodInteresSuspenso & ","
                        sSql = sSql & MatCredRef(i, 6) & ","
                        sSql = sSql & "0.00)"
                        Conn.ConexionActiva.Execute sSql
                    End If
                    
                    'Interes Reprogramado
                    If CDbl(MatCredRef(i, 7)) > 0 Then
                        sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSql = sSql & " VALUES('" & psctacod & "','"
                        sSql = sSql & MatCredRef(i, 0) & "',"
                        sSql = sSql & gColocEstadoRefinancSolicitado & ","
                        sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                        sSql = sSql & gColocConceptoCodInteresReprogramado & ","
                        sSql = sSql & MatCredRef(i, 7) & ","
                        sSql = sSql & "0.00)"
                        Conn.ConexionActiva.Execute sSql
                    End If
                    
                    'Gastos
                    If CDbl(MatCredRef(i, 8)) > 0 Then
                        sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSql = sSql & " VALUES('" & psctacod & "','"
                        sSql = sSql & MatCredRef(i, 0) & "',"
                        sSql = sSql & gColocEstadoRefinancSolicitado & ","
                        sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                        sSql = sSql & gColocConceptoCodGastoVarios & ","
                        sSql = sSql & MatCredRef(i, 8) & ","
                        sSql = sSql & "0.00)"
                        Conn.ConexionActiva.Execute sSql
                    End If
                Next i
            End If
            Set oBase = Nothing
        End If
    End If
    
    ' un credito automatico
        oRelPersCred.IniciarMatriz
        Do While Not oRelPersCred.EOF
                If oRelPersCred.ObtenerValorRelac = 20 Then ' si es titular
                    If psCondCredOtra = "3" Then
                        sSql = "Update Persona set bPreferencial=1 Where cPersCod='" & oRelPersCred.ObtenerCodigo & "'"
                    Else
                        sSql = "Update Persona set bPreferencial=0 Where cPersCod='" & oRelPersCred.ObtenerCodigo & "'"
                    End If
                    Conn.ConexionActiva.Execute sSql
                End If
            oRelPersCred.siguiente
        Loop
        
    Conn.ConexionActiva.CommitTrans
    Conn.CierraConexion
    Set Conn = Nothing
    
    Exit Sub

ErrorActualizaSolicitud:
    If bTran Then
        Conn.ConexionActiva.RollbackTrans
        Set Conn = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso ActualizaSolicitud", Err.Description

End Sub


Public Sub NuevaSolicitud(ByVal oRelPersCred As UCredRelacion, ByVal psctacod As String, _
                          ByVal psCondCred As String, ByVal psCondCredOtra As String, ByVal popersona As DPersona, ByVal psFteIngreso As String, ByVal pnMontoSol As Double, _
                          ByVal pnCuotasSol As Integer, ByVal pnPlazo As Integer, ByVal psDestCred As String, _
                          ByVal psPersCodAnalista As String, ByVal psPersCodInst As String, ByVal psCodModular As String, _
                          ByVal dFecSol As Date, ByVal psMovAct As String, ByVal bRefinanciar As Boolean, _
                          ByVal pnCondProd As Integer, ByVal dFecEvalFte As Date, Optional ByVal MatCredRef As Variant, _
                          Optional pnCapitalInt As Boolean = False, Optional bSustiDeudor As Boolean = False, _
                          Optional ByVal bAmpliado As Boolean = False, Optional rsAmpliado As ADODB.Recordset, _
                          Optional ByVal nCampanaCod As Integer)
Dim pbTran As Boolean
Dim i As Integer
Dim sSql As String
Dim Conn As DConecta
Dim oBase As DCredActualizaBD
Dim dFecGrab As Date
Dim dGar As DGarantia
Dim R As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim lcInstitucionCmac As String
Dim LsNumGarant As String

    On Error GoTo ErrorNuevaSolicitud
    
    pbTran = False
    Set Conn = New DConecta
    Conn.AbreConexion
    
    ' se obtiene el valor de la instucion de  la cmac ica para creditos de los trabajadores
    If Mid(psctacod, 6, 3) = "320" Then
        ' si es credito de trabajadores y funcionarios
        sSql = "Select nConsSisValor From ConstSistema where nConsSisCod=204"
        Set rs = Conn.CargaRecordSet(sSql)
        If Not rs.EOF And Not rs.BOF Then
            lcInstitucionCmac = rs!nConssisValor
        End If
        Set rs = Nothing
    End If
    
    Conn.ConexionActiva.BeginTrans
    pbTran = True
    
    'Ingresa Nuevo Producto
    sSql = "INSERT INTO Producto(cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc)"
    sSql = sSql & " VALUES('" & psctacod & "',"
    sSql = sSql & "0.00,"
    sSql = sSql & "0.00,"
    sSql = sSql & Trim(Str(gColocEstSolic)) & ","
    sSql = sSql & "'" & Format(dFecSol, "mm/dd/yyyy") & "',"
    sSql = sSql & "0)"
    Conn.ConexionActiva.Execute sSql

    'Ingresa Nuevas Personas Relacionadas con el Credito
    oRelPersCred.IniciarMatriz
    Do While Not oRelPersCred.EOF
        sSql = "INSERT INTO ProductoPersona(cCtaCod,cPersCod,nPrdPersRelac) "
        sSql = sSql & " VALUES('" & psctacod & "','" & oRelPersCred.ObtenerCodigo & "'," & oRelPersCred.ObtenerValorRelac & ")"
        Conn.ConexionActiva.Execute sSql
        oRelPersCred.siguiente
    Loop

    'Ingresa el Analista
    sSql = "INSERT INTO ProductoPersona(cCtaCod,cPersCod,nPrdPersRelac) "
    sSql = sSql & " VALUES('" & psctacod & "','" & psPersCodAnalista & "'," & gColRelPersAnalista & ")"
    Conn.ConexionActiva.Execute sSql
    
    'Nuevo Registro de Colocaciones
    sSql = "INSERT INTO Colocaciones(cCtaCod,nPlazo,dVenc,nMontoCol,nColocCalendCod,cUltimaActualizacion,cLineaCred)"
    sSql = sSql & " VALUES('" & psctacod & "',"
    sSql = sSql & Trim(Str(pnPlazo)) & ","
    sSql = sSql & "NULL" & ","
    sSql = sSql & Format(pnMontoSol, "#0.00") & ","
    sSql = sSql & "0,"
    sSql = sSql & "'" & psMovAct & "',"
    sSql = sSql & "'')"
    Conn.ConexionActiva.Execute sSql
    
    'Nuevo Registro de ColocacCred
    sSql = "INSERT INTO ColocacCred (cCtaCod,nDiasAtraso,nColocCondicion,nColocCondicion2,nColocDestino,"
    sSql = sSql & "cProtesto,bCargoAuto,cMetLiquidacion,bRefCapInt,nNroProxCuota,nIntPend,"
    sSql = sSql & "nExoPenalidad, nColocCondicionProd,IdCampana)"
    sSql = sSql & "VALUES('" & psctacod & "',"
    sSql = sSql & "0,"
    sSql = sSql & psCondCred & ","
    sSql = sSql & psCondCredOtra & ","
    sSql = sSql & psDestCred & ","
    sSql = sSql & "0,"
    sSql = sSql & "NULL,"
    sSql = sSql & "'GMIC',"
    sSql = sSql & IIf(pnCapitalInt, "1,", "0,")
    sSql = sSql & "1,"
    sSql = sSql & "0,"
    sSql = sSql & "NULL," & pnCondProd & ","
    sSql = sSql & IIf(nCampanaCod = 0, "NULL", nCampanaCod) & ")"
    Conn.ConexionActiva.Execute sSql
    
    'Nuevo Registro en ColocEstado
    Set oBase = New DCredActualizaBD
    sSql = "INSERT INTO ColocacEstado(cCtaCod,dPrdEstado,nPrdEstado,nCuotas,nMonto,cDescripcion,"
    sSql = sSql & "nColocCalendCod,nPeriodoFechaFija,nPeriodoGracia,nPlazo,nTipoGracia)"
    sSql = sSql & " VALUES('" & psctacod & "',"
    sSql = sSql & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
    sSql = sSql & Trim(Str(gColocEstSolic)) & ","
    sSql = sSql & Trim(Str(pnCuotasSol)) & ","
    sSql = sSql & Format(pnMontoSol, "#0.00") & ","
    sSql = sSql & "'Nueva Solicitud de Credito',"
    sSql = sSql & "NULL,"
    sSql = sSql & "NULL,"
    sSql = sSql & "NULL,"
    sSql = sSql & Trim(Str(pnPlazo)) & ","
    sSql = sSql & "NULL)"
    Set oBase = Nothing
    Conn.ConexionActiva.Execute sSql
    
    'Nuevo Registro en ColocFteIngreso
    sSql = "INSERT INTO ColocFteIngreso(cNumFuente,dPersEval,cCtaCod)"
            sSql = sSql & " VALUES('" & psFteIngreso & "',"
            sSql = sSql & "'" & Format(dFecEvalFte, "mm/dd/yyyy") & "', "
            sSql = sSql & "'" & psctacod & "')"
            Conn.ConexionActiva.Execute sSql

    'Nuevo Registro de ColocConvenio
    If Len(Trim(psPersCodInst)) > 0 Then
        sSql = "INSERT INTO ColocacConvenio(cCtaCod,cPersCod,cCodModular)"
        sSql = sSql & " VALUES('" & psctacod & "',"
        sSql = sSql & "'" & psPersCodInst & "',"
        sSql = sSql & "'" & psCodModular & "')"
        Conn.ConexionActiva.Execute sSql
    End If
    
    'Si es credito de funcionarios de cmac ica
    If Mid(psctacod, 6, 3) = "320" And Len(lcInstitucionCmac) = 13 Then
        sSql = "INSERT INTO ColocacConvenio(cCtaCod,cPersCod,cCodModular)"
        sSql = sSql & " VALUES('" & psctacod & "',"
        sSql = sSql & "'" & lcInstitucionCmac & "',"
        sSql = sSql & "'" & " " & "')"
        Conn.ConexionActiva.Execute sSql
    End If
    'Si es una Refinanciacion Guardar los Datos
    If bRefinanciar Then
        If IsArray(MatCredRef) Then
            Set oBase = New DCredActualizaBD
            dFecGrab = CDate(Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "dd/mm/yyyy hh:mm:ss"))
            Set oBase = Nothing
            If UBound(MatCredRef) > 0 Then
                For i = 0 To UBound(MatCredRef) - 1
                    sSql = "INSERT INTO ColocacRefinanc(cCtaCod,cCtaCodRef,nEstado,dEstado,nMontoRef)"
                    sSql = sSql & " VALUES('" & psctacod & "','"
                    sSql = sSql & MatCredRef(i, 0) & "',"
                    sSql = sSql & gColocEstadoRefinancSolicitado & ","
                    sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                    sSql = sSql & MatCredRef(i, 1) & ")"
                    Conn.ConexionActiva.Execute sSql
                    
                    'Ingreso del Detalle
                    'Capital
                    sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                    sSql = sSql & " VALUES('" & psctacod & "','"
                    sSql = sSql & MatCredRef(i, 0) & "',"
                    sSql = sSql & gColocEstadoRefinancSolicitado & ","
                    sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                    sSql = sSql & gColocConceptoCodCapital & ","
                    sSql = sSql & MatCredRef(i, 2) & ","
                    sSql = sSql & "0.00)"
                    Conn.ConexionActiva.Execute sSql
                    'Interes Compensatorio
                    If CDbl(MatCredRef(i, 3)) > 0 Then
                        sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSql = sSql & " VALUES('" & psctacod & "','"
                        sSql = sSql & MatCredRef(i, 0) & "',"
                        sSql = sSql & gColocEstadoRefinancSolicitado & ","
                        sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                        sSql = sSql & gColocConceptoCodInteresCompensatorio & ","
                        sSql = sSql & MatCredRef(i, 3) & ","
                        sSql = sSql & "0.00)"
                        Conn.ConexionActiva.Execute sSql
                    End If
                    
                    'Interes Moratorio
                    If CDbl(MatCredRef(i, 4)) > 0 Then
                        sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSql = sSql & " VALUES('" & psctacod & "','"
                        sSql = sSql & MatCredRef(i, 0) & "',"
                        sSql = sSql & gColocEstadoRefinancSolicitado & ","
                        sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                        sSql = sSql & gColocConceptoCodInteresMoratorio & ","
                        sSql = sSql & MatCredRef(i, 4) & ","
                        sSql = sSql & "0.00)"
                        Conn.ConexionActiva.Execute sSql
                    End If
                    
                    'Interes Gracia
                    If CDbl(MatCredRef(i, 5)) > 0 Then
                        sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSql = sSql & " VALUES('" & psctacod & "','"
                        sSql = sSql & MatCredRef(i, 0) & "',"
                        sSql = sSql & gColocEstadoRefinancSolicitado & ","
                        sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                        sSql = sSql & gColocConceptoCodInteresGracia & ","
                        sSql = sSql & MatCredRef(i, 5) & ","
                        sSql = sSql & "0.00)"
                        Conn.ConexionActiva.Execute sSql
                    End If
                    
                    'Interes Suspenso
                    If CDbl(MatCredRef(i, 6)) > 0 Then
                        sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSql = sSql & " VALUES('" & psctacod & "','"
                        sSql = sSql & MatCredRef(i, 0) & "',"
                        sSql = sSql & gColocEstadoRefinancSolicitado & ","
                        sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                        sSql = sSql & gColocConceptoCodInteresSuspenso & ","
                        sSql = sSql & MatCredRef(i, 6) & ","
                        sSql = sSql & "0.00)"
                        Conn.ConexionActiva.Execute sSql
                    End If
                    
                    'Interes Reprogramado
                    If CDbl(MatCredRef(i, 7)) > 0 Then
                        sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSql = sSql & " VALUES('" & psctacod & "','"
                        sSql = sSql & MatCredRef(i, 0) & "',"
                        sSql = sSql & gColocEstadoRefinancSolicitado & ","
                        sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                        sSql = sSql & gColocConceptoCodInteresReprogramado & ","
                        sSql = sSql & MatCredRef(i, 7) & ","
                        sSql = sSql & "0.00)"
                        Conn.ConexionActiva.Execute sSql
                    End If
                    
                    'Gastos
                    If CDbl(MatCredRef(i, 8)) > 0 Then
                        sSql = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSql = sSql & " VALUES('" & psctacod & "','"
                        sSql = sSql & MatCredRef(i, 0) & "',"
                        sSql = sSql & gColocEstadoRefinancSolicitado & ","
                        sSql = sSql & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                        sSql = sSql & gColocConceptoCodGastoVarios & ","
                        sSql = sSql & MatCredRef(i, 8) & ","
                        sSql = sSql & "0.00)"
                        Conn.ConexionActiva.Execute sSql
                    End If
                Next i
            End If
        End If
    End If
    
    'SI es Refinanciado Ingresar las Mismas Garantias de los Creditos Anteriores
    Dim cNumGaranOld As String
    Dim nNum As Integer
    If bRefinanciar Then
'        Set dGar = New DGarantia
'        For i = 0 To UBound(MatCredRef) - 1
'                sSql = "Select CN.nConsValor as nTpoGar, G.cNroDoc, CN.cConsDescripcion cTpoGarantia, G.cDescripcion, D.cDocDesc, CG.nGravado, CN2.cConsDescripcion cMoneda, G.nPorGravar, G.nGravaMent, CN3.cConsDescripcion cMonedaGar,G.cNumgarant, CG.nMoneda, CG.nEstado"
'                sSql = sSql & " From ColocGarantia CG Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant "
'                sSql = sSql & "                       Inner Join Constante CN ON G.nTpoGarantia = CN.nConsValor AND CN.nConsCod = " & gPersGarantia
'                sSql = sSql & "                       Inner Join Constante CN2 ON CN2.nConsValor = CG.nMoneda AND CN2.nConsCod = " & gMoneda
'                sSql = sSql & "                       Inner Join Constante CN3 ON CN3.nConsValor = G.nMoneda AND CN3.nConsCod = " & gMoneda
'                sSql = sSql & "                       Inner Join Documento D ON D.nDocTpo = Convert(int,G.cTpoDoc) "
'                sSql = sSql & "                       Inner Join GarantDoc O ON D.nDocTpo = O.nDocTpo  and O.nConsValor=G.nTpoGarantia"
'                sSql = sSql & " WHERE CG.cCtaCod = '" & MatCredRef(i, 0) & "'"
'                sSql = sSql & " Order By G.cNumgarant"
'            'Set R = dGar.RecuperaGarantiaCreditoDatos(MatCredRef(i, 0))
'             Set R = Conn.CargaRecordSet(sSql)
'            Do While Not R.EOF
'                If nNum = 0 Then
'                    sSql = "INSERT INTO COLOCGARANTIA(cNumGarant, cCtaCod, nMoneda, nGravado, nEstado) "
'                    sSql = sSql & " VALUES('" & R!cNumGarant & "','" & psCtaCod & "'," & R!nMoneda & "," & Format(R!nGravado, "#0.00") & "," & IIf(IsNull(R!nEstado), 1, R!nEstado) & ") "
'                    Conn.ConexionActiva.Execute sSql
'                Else
'                    If cNumGaranOld = R!cNumGarant Then
'                        sSql = "Update ColocGarantia"
'                        sSql = sSql & " Set nGravado=nGravado+ " & R!nGravado
'                        sSql = sSql & " Where cNumGarant='" & R!cNumGarant & "'"
'                        Conn.ConexionActiva.Execute sSql
'                    Else
'                        sSql = "INSERT INTO COLOCGARANTIA(cNumGarant, cCtaCod, nMoneda, nGravado, nEstado) "
'                        sSql = sSql & " VALUES('" & R!cNumGarant & "','" & psCtaCod & "'," & R!nMoneda & "," & Format(R!nGravado, "#0.00") & "," & IIf(IsNull(R!nEstado), 1, R!nEstado) & ") "
'                        Conn.ConexionActiva.Execute sSql
'                    End If
'                End If
'                nNum = nNum + 1
'                cNumGaranOld = R!cNumGarant
'                R.MoveNext
'            Loop
'        Next i
'        Set dGar = Nothing
      For i = 0 To UBound(MatCredRef) - 1
            If i = 0 Then
                LsNumGarant = "'" & MatCredRef(i, 0) & "'"
            Else
                LsNumGarant = LsNumGarant & ",'" & MatCredRef(i, 0) & "'"
            End If
      Next i
      
      sSql = "Select CN.nConsValor as nTpoGar, G.cNroDoc, CN.cConsDescripcion cTpoGarantia, G.cDescripcion, D.cDocDesc, Sum(CG.nGravado) as nGravado, CN2.cConsDescripcion cMoneda, SUM(G.nPorGravar) as nPorGravar, Sum(G.nGravaMent) as nGravaMent, CN3.cConsDescripcion cMonedaGar,G.cNumgarant, CG.nMoneda, CG.nEstado "
      sSql = sSql & " From ColocGarantia CG"
      sSql = sSql & " Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant"
      sSql = sSql & " Inner Join Constante CN ON G.nTpoGarantia = CN.nConsValor AND CN.nConsCod = 1027"
      sSql = sSql & " Inner Join Constante CN2 ON CN2.nConsValor = CG.nMoneda AND CN2.nConsCod = 1011"
      sSql = sSql & " Inner Join Constante CN3 ON CN3.nConsValor = G.nMoneda AND CN3.nConsCod = 1011"
      sSql = sSql & " Inner Join Documento D ON D.nDocTpo = Convert(int,G.cTpoDoc)"
      sSql = sSql & " Inner Join GarantDoc O ON D.nDocTpo = O.nDocTpo  and O.nConsValor=G.nTpoGarantia"
      sSql = sSql & " WHERE CG.cCtaCod in(" & LsNumGarant & ")"
      sSql = sSql & " GROUP BY CG.nEstado,CN.nConsValor,G.cNroDoc,CN.cConsDescripcion,G.cDescripcion,D.cDocDesc,CN2.cConsDescripcion,CN3.cConsDescripcion,G.cNumgarant,CG.nMoneda"
      sSql = sSql & " Order By G.cNumgarant"
    
        Set R = Conn.CargaRecordSet(sSql)
        Do Until R.EOF
        sSql = "INSERT INTO COLOCGARANTIA(cNumGarant, cCtaCod, nMoneda, nGravado, nEstado) "
        sSql = sSql & " VALUES('" & R!cNumGarant & "','" & psctacod & "'," & R!nMoneda & "," & Format(R!nGravado, "#0.00") & "," & IIf(IsNull(R!nEstado), 1, R!nEstado) & ") "
        Conn.ConexionActiva.Execute sSql
        R.MoveNext
        Loop
        Set R = Nothing
    End If
        ' un credito automatico
        oRelPersCred.IniciarMatriz
        Do While Not oRelPersCred.EOF
                If oRelPersCred.ObtenerValorRelac = 20 Then ' si es titular
                    If psCondCredOtra = "3" Then
                        sSql = "Update Persona set bPreferencial=1 Where cPersCod='" & oRelPersCred.ObtenerCodigo & "'"
                    Else
                        sSql = "Update Persona set bPreferencial=0 Where cPersCod='" & oRelPersCred.ObtenerCodigo & "'"
                    End If
                    Conn.ConexionActiva.Execute sSql
                End If
            oRelPersCred.siguiente
        Loop
    
      ' AUTOR:LMMD
      ' SI SE TRATA DE UN CREDITO AMPLIADO
      If bAmpliado = True Then
        
            Dim nEstado As Integer
            
            rsAmpliado.MoveFirst
            Do Until rsAmpliado.EOF
                
                sSql = "Select nPrdEstado From Producto Where cCtaCod='" & rsAmpliado(0) & "'"
                Set rs = New ADODB.Recordset
                Set rs = Conn.CargaRecordSet(sSql)
                 
                'SE OBTIENE EL ESTADO
                If Not rs.EOF And Not rs.BOF Then
                    nEstado = rs!nPrdEstado
                End If
                
                sSql = "Insert Into ColocacAmpliado Values('" & psctacod & "','" & rsAmpliado(0) & "'," & nEstado & ","
                sSql = sSql & "'" & Format(gdFecSis, "MM/dd/yyyy") & "'," & rsAmpliado(3) & "," & pnMontoSol & ")"
                
                Conn.ConexionActiva.Execute (sSql)
             rsAmpliado.MoveNext
          Loop
      End If
    
    
    Conn.ConexionActiva.CommitTrans
    Conn.CierraConexion
    Set Conn = Nothing
    
    Exit Sub

ErrorNuevaSolicitud:
    If pbTran Then
        Conn.ConexionActiva.RollbackTrans
        Set Conn = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso NuevaSolicitud", Err.Description
End Sub

Public Sub ActuializaPersRelacCred(ByVal poRelPersCred As UCredRelacion, ByVal psctacod As String, ByVal pbTransac As Boolean)
Dim sSql As String
    
    On Error GoTo ErrorActuializaPersRelacCred
    If pbTransac Then
        Set Conn = New DConecta
        Conn.AbreConexion
        Conn.ConexionActiva.BeginTrans
    End If
    poRelPersCred.IniciarMatriz
    If Not poRelPersCred.EOF Then
        sSql = "DELETE ProductoPersona Where cCtaCod = '" & psctacod & "' AND nPrdPersRelac not in (" & gColRelPersAnalista & "," & gColRelPersApoderado & ")"
        Conn.ConexionActiva.Execute sSql
    End If
    Do While Not poRelPersCred.EOF
        sSql = "INSERT INTO ProductoPersona(cCtaCod,cPersCod,nPrdPersRelac) "
        sSql = sSql & "       VALUES('" & psctacod & "','" & poRelPersCred.ObtenerCodigo & "'," & poRelPersCred.ObtenerValorRelac & ")"
        Conn.ConexionActiva.Execute sSql
        poRelPersCred.siguiente
    Loop
    If pbTransac Then
        Conn.ConexionActiva.CommitTrans
        Conn.CierraConexion
        Set Conn = Nothing
    End If
    Exit Sub

ErrorActuializaPersRelacCred:
    If pbTransac Then
        Conn.ConexionActiva.RollbackTrans
        Set Conn = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso ActuializaPersRelacCred", Err.Description
        
End Sub

'*****************************************************************************************
'***     Rutina:            RecuperaCodigosModulares
'***     Descripcion:       Permite Obtener los Codigos Modulares de una persona
'***     Creado por:        NSSE
'***     Maquina:           07SIST_08
'***     Fecha-Tiempo:      31/05/2001 09:10:04 AM
'***     Ultima Modificacion: Creacion
'*****************************************************************************************

Public Function RecuperaRelacPers(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaRelacPers
    
    sSql = "Select P.dPersNacCreac, PN.cPersNatSexo, P.nPersPersoneria, Prd.nPrdEstado, PP.cPersCod, RTRIM(P.cPersNombre) cPersNombre, C.cConsDescripcion, C.nConsValor, DNI = (Select Top 1 cPersIDnro from PersID Where cPersCod = PP.cPersCod  AND cPersIDTpo = '" & gPersIdDNI & "'),"
    sSql = sSql & " RUC = (Select Top 1 cPersIDnro from PersID Where cPersCod = PP.cPersCod AND cPersIDTpo = '" & gPersIdRUC & "'), "
    sSql = sSql & " CASE WHEN PP.nPrdPersRelac = 20 THEN 1 WHEN PP.nPrdPersRelac = 22 THEN 2 WHEN PP.nPrdPersRelac = 25 THEN 3 ELSE 99 END nOrden   "
    sSql = sSql & " From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod "
    sSql = sSql & "                         LEFT JOIN PersonaNat PN ON PN.cPersCod = P.cPersCod "
    sSql = sSql & "                         INNER JOIN Producto Prd ON Prd.cCtaCod = PP.cCtaCod "
    sSql = sSql & "                         INNER JOIN Constante C ON PP.nPrdPersRelac = C.nConsValor "
    sSql = sSql & " WHERE C.nConsCod = " & gColocRelacPers & " And C.nConsValor not in (" & gColRelPersAnalista & "," & gColRelPersApoderado & "," & gColRelPersAcreedor & ") AND PP.cCtaCod = '" & psctacod & "' "
    sSql = sSql & " ORDER BY nOrden "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaRelacPers = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaRelacPers:
    Err.Raise Err.Number, "Error En Proceso de Recuperar Relaciones de Personas con el Credito", Err.Description
End Function

Public Function RecuperaProductoTasaInteres(ByVal psctacod As String, ByVal pnPrdTasaInteres As ColocTipoTasa) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaProductoTasaInteres
    sSql = "Select * from ProductoTasaInteres Where cCtaCod = '" & psctacod & "' AND nPrdTasaInteres = " & pnPrdTasaInteres
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaProductoTasaInteres = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaProductoTasaInteres:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaProducto(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaProducto
    sSql = "Select * from Producto where cCtacod = '" & psctacod & "' "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaProducto = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaProducto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaColocacCred(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaColocacCred
    sSql = "Select * from ColocacCred where cCtacod = '" & psctacod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocacCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocacCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaSugerencia(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaSugerencia
    
'    sSQL = "Select ISNULL(CE.cTipoGasto,'F') as cTipoGasto,CE.nPrdEstado, CE.cDescripcion, PRD.nTransacc, C.nColocCalendCod, CN2.cConsDescripcion as cTipoCredDescrip, C.nColocDestino, CN.cConsDescripcion as cDestinoDescripcion, P.cPersCod, P.cPersNombre, "
'    sSQL = sSQL & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), "
'    sSQL = sSQL & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdRUC & "'), "
'    sSQL = sSQL & " CE.nProxMes,CE.nCalendDinamico, CE.nMonto, CE.nCuotas, CE.nPlazo, CE.dPrdEstado,  CE.nPeriodoFechaFija, CE.nPeriodoGracia, CE.nPlazo, CE.nTipoGracia, CE.nTipoDesembolso, "
'    sSQL = sSQL & " cAnalista = (select Top 1 cPersNombre from persona where cPersCod = (Select cPersCod From ProductoPersona Where cCtaCod = C.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ")), "
'    sSQL = sSQL & " nMontoSol = (Select nMonto from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "), "
'    sSQL = sSQL & " nCuotasSol = (Select nCuotas from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
'    sSQL = sSQL & " nPlazoSol = (Select nPlazo from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
'    sSQL = sSQL & " nTasaComp = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & "),"
'    sSQL = sSQL & " nTasaGracia = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntGracia & "),"
'    sSQL = sSQL & " cLineaCred = (Select cLineaCred from Colocaciones Where cCtaCod = C.cCtaCod), "
'    sSQL = sSQL & " C.bMiVivienda, C.bCuotaCom, C.nNroCalPar, ISNULL(PT.nTasaInteres,0) as nTasaMora, ISNULL(PT2.nTasaInteres,0) as nTasaICVen "
'    sSQL = sSQL & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
'    sSQL = sSQL & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
'    sSQL = sSQL & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado in (" & gColocEstSug & "," & gColocEstSolic & ")"
'    sSQL = sSQL & "                    INNER JOIN Constante CN ON C.nColocDestino = CN.nConsValor AND CN.nConsCod = " & Trim(Str(gColocDestino))
'    sSQL = sSQL & "                    INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = " & gProducto
'    sSQL = sSQL & "                    INNER JOIN Producto PRD ON C.cCtaCod = PRD.cCtaCod"
'    sSQL = sSQL & "                    LEFT JOIN ProductoTasaInteres PT ON C.cCtaCod = PT.cCtaCod AND PT.nPrdTasaInteres = 3 "
'    sSQL = sSQL & "                    LEFT JOIN ProductoTasaInteres PT2 ON C.cCtaCod = PT2.cCtaCod AND PT2.nPrdTasaInteres = 6 "
'    sSQL = sSQL & " WHERE C.cCtaCod = '" & psCtaCod & "' Order By CE.nPrdEstado DESC"
    
    sSql = " Select ISNULL(CE.cTipoGasto,'F') as cTipoGasto,CE.nPrdEstado, CE.cDescripcion, PRD.nTransacc, C.nColocCalendCod, CN2.cConsDescripcion as cTipoCredDescrip, C.nColocDestino, CN.cConsDescripcion as cDestinoDescripcion, P.cPersCod, P.cPersNombre, "
    sSql = sSql & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), "
    sSql = sSql & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdRUC & "'), "
    sSql = sSql & " CE.nProxMes,CE.nCalendDinamico, CE.nMonto, CE.nCuotas, CE.nPlazo, CE.dPrdEstado,  CE.nPeriodoFechaFija, CE.nPeriodoGracia, CE.nPlazo, CE.nTipoGracia, CE.nTipoDesembolso, "
    sSql = sSql & " cAnalista = (select Top 1 cPersNombre from persona where cPersCod = (Select cPersCod From ProductoPersona Where cCtaCod = C.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ")), "
    sSql = sSql & " nMontoSol = (Select nMonto from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "), "
    sSql = sSql & " nCuotasSol = (Select nCuotas from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
    sSql = sSql & " nPlazoSol = (Select nPlazo from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
    sSql = sSql & " nTasaComp = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & "),"
    sSql = sSql & " nTasaGracia = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntGracia & "),"
    sSql = sSql & " cLineaCred = (Select cLineaCred from Colocaciones Where cCtaCod = C.cCtaCod), "
    'CMACICA_CSTS - 26/11/2003 --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    sSql = sSql & " nPersFIDIngCli = (SELECT PFD.nPersIngCli FROM COLOCFTEINGRESO CFI INNER JOIN PERSFIDEPENDIENTE PFD ON CFI.cNumFuente = PFD.cNumFuente Where CFI.cCtaCod = C.cCtaCod AND PFD.dPersEval = ( SELECT MAX(PFD2.dPersEval) FROM COLOCFTEINGRESO CFI2 INNER JOIN PERSFIDEPENDIENTE PFD2 ON CFI2.cNumFuente = PFD2.cNumFuente WHERE CFI2.cCtaCod = CFI.cCtaCod )),"
    sSql = sSql & " cPersFIMoneda = (SELECT     PF.cPersFIMoneda FROM COLOCFTEINGRESO CFI INNER JOIN PERSFTEINGRESO PF ON CFI.cNumFuente = PF.cNumFuente Where CFI.cCtaCod = C.cCtaCod),"
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    sSql = sSql & " C.bMiVivienda, C.bCuotaCom, C.nNroCalPar, ISNULL(PT.nTasaInteres,0) as nTasaMora, ISNULL(PT2.nTasaInteres,0) as nTasaICVen "
    sSql = sSql & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
    sSql = sSql & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado in (" & gColocEstSug & "," & gColocEstSolic & ")"
    sSql = sSql & "                    INNER JOIN Constante CN ON C.nColocDestino = CN.nConsValor AND CN.nConsCod = " & Trim(Str(gColocDestino))
    sSql = sSql & "                    INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = " & gProducto
    sSql = sSql & "                    INNER JOIN Producto PRD ON C.cCtaCod = PRD.cCtaCod"
    sSql = sSql & "                    LEFT JOIN ProductoTasaInteres PT ON C.cCtaCod = PT.cCtaCod AND PT.nPrdTasaInteres = 3 "
    sSql = sSql & "                    LEFT JOIN ProductoTasaInteres PT2 ON C.cCtaCod = PT2.cCtaCod AND PT2.nPrdTasaInteres = 6 "
    sSql = sSql & " WHERE C.cCtaCod = '" & psctacod & "' Order By CE.nPrdEstado DESC"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaSugerencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaSugerencia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
End Function

Public Function RecuperaDatosDesembolso(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaDatosDesembolso
    sSql = "Select CC.nNroCalen, C.cLineaCred, L.cDescripcion, P.cPersCod, P.cPersNombre, P.cPersDireccDomicilio,CN.cConsDescripcion cMoneda,  "
    sSql = sSql & " CC.nNroproxDesemb, nPrestamo = (Select nMonto From ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstAprob & "),"
    sSql = sSql & " nMontoADesemb = (Select CD.nMonto from ColocCalenddet CD "
    sSql = sSql & "                 where CD.cCtaCod = C.cCtaCod and CD.nColocCalendApl=0 and nPrdConceptoCod=1000 and nCuota = CC.nNroproxDesemb "
    sSql = sSql & "                 and CD.nnrocalen=(Select MAX(nNroCalen) From ColocCalendario Where cCtaCod = CD.cCtaCod AND nColocCalendApl=0) ),  "
    sSql = sSql & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), "
    sSql = sSql & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdRUC & "'), "
    sSql = sSql & " nTotalDesemb=(Select Count(cCtaCod) from ColocCalendario Where cCtaCod = C.cCtaCod AND nNroCalen = CC.nNroCalen AND nColocCalendApl = " & gColocCalendAplDesembolso & "), "
    sSql = sSql & " CC.bMiVivienda, CC.bCuotaCom "
    sSql = sSql & " from Colocaciones C Inner Join ColocLineaCredito L ON C.cLineaCred = L.cLineaCred "
    sSql = sSql & "                       Inner Join ColocacCred CC ON C.cCtaCod = CC.cCtaCod"
    sSql = sSql & "                       Inner Join ProductoPersona PP ON PP.cCtaCod = C.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                       Inner Join Persona P ON PP.cPersCod = P.cPersCod"
    sSql = sSql & "                       Inner Join Constante CN ON CN.nConsValor = Convert(int,substring(C.cCtaCod,9,1)) And CN.nConsCod = " & gMoneda
    sSql = sSql & " WHERE C.cCtaCod = '" & psctacod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosDesembolso = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosDesembolso:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaDatosAprobacion(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaDatosAprobacion
    
    sSql = "Select CC.cLineaCred as cLineaCredCod,  CE.nPrdEstado, C.nColocCalendCod, CN2.cConsDescripcion as cTipoCredDescrip, C.nColocDestino, CN.cConsDescripcion as cDestinoDescripcion, "
    sSql = sSql & " ISNULL(CE.cTipogasto,'F') as cTipogasto, CE.nMonto, CE.nCuotas, CE.nPlazo, CE.nPeriodoGracia, CE.nPeriodoFechaFija, PRD.nTasaInteres, CE.nProxMes, CE.nCalendDinamico, CE.nTipoGracia, CE.nTipoDesembolso, "
    sSql = sSql & " cAnalista = (select Top 1 cPersNombre from persona where cPersCod = (Select cPersCod From ProductoPersona Where cCtaCod = C.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ")), "
    sSql = sSql & " nMontoSol = (Select nMonto from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "), "
    sSql = sSql & " nCuotasSol = (Select nCuotas from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
    sSql = sSql & " nPlazoSol = (Select nPlazo from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
    sSql = sSql & " nTasaComp = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & "),"
    sSql = sSql & " nTasaGracia = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntGracia & "),"
    sSql = sSql & " cLineaCred = (Select (cDescripcion +  space(150) + cLineaCred) from ColocLineaCredito Where cLineaCred = CC.cLineaCred), "
    sSql = sSql & " P2.cPersNombre cFteIngreso, PF.cPersFIPersCod, P.cPersCod, P.cPersNombre, P.cPersDireccDomicilio, "
    sSql = sSql & " CN3.cConsDescripcion as cCondicionCred, C.nColocCondicion, C.bRefCapInt, "
    sSql = sSql & " nMaxCalend = (Select MAX(nNroCalen) From ColocCalendario Cal Where cCtaCod = C.cCtaCod), "
    sSql = sSql & " C.bMiVivienda, C.bCuotaCom, "
    sSql = sSql & " nTasaIntComp = (Select nTasaInteres from ProductoTasaInteres where cCtaCod = C.cCtaCod AND nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & "), "
    sSql = sSql & " nTasaIntMora = (Select nTasaInteres from ProductoTasaInteres where cCtaCod = C.cCtaCod AND nPrdTasaInteres = " & gColocLineaCredTasasIntMoratNormal & ") "
    sSql = sSql & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                    INNER JOIN Colocaciones CC ON CC.cCtaCod = C.cCtaCod "
    sSql = sSql & "                    INNER JOIN PersFteIngreso PF ON PF.cPersCod = PP.cPersCod "
    sSql = sSql & "                    INNER JOIN Persona P2 ON P2.cPersCod = PF.cPersFIPersCod "
    sSql = sSql & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
    sSql = sSql & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSug
    sSql = sSql & "                    INNER JOIN Constante CN ON C.nColocDestino = CN.nConsValor AND CN.nConsCod = " & Trim(Str(gColocDestino))
    sSql = sSql & "                    INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = " & gProducto
    sSql = sSql & "                    INNER JOIN Constante CN3 ON C.nColocCondicion = CN3.nConsValor AND CN3.nConsCod = " & Trim(Str(gColocCredCondicion))
    sSql = sSql & "                    INNER JOIN Producto PRD ON C.cCtaCod = PRD.cCtaCod"
    sSql = sSql & " WHERE C.cCtaCod = '" & psctacod & "' And Prd.nPrdestado = " & gColocEstSug
    sSql = "Select CC.cLineaCred as cLineaCredCod,  CE.nPrdEstado, C.nColocCalendCod, CN2.cConsDescripcion as cTipoCredDescrip, C.nColocDestino, CN.cConsDescripcion as cDestinoDescripcion, "
    sSql = sSql & " ISNULL(CE.cTipogasto,'F') as cTipogasto, CE.nMonto, CE.nCuotas, CE.nPlazo, CE.nPeriodoGracia, CE.nPeriodoFechaFija, PRD.nTasaInteres, CE.nProxMes, CE.nCalendDinamico, CE.nTipoGracia, CE.nTipoDesembolso, "
    sSql = sSql & " cAnalista = (select Top 1 cPersNombre from persona where cPersCod = (Select cPersCod From ProductoPersona Where cCtaCod = C.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ")), "
    sSql = sSql & " nMontoSol = (Select nMonto from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "), "
    sSql = sSql & " nCuotasSol = (Select nCuotas from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
    sSql = sSql & " nPlazoSol = (Select nPlazo from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
    sSql = sSql & " nTasaComp = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & "),"
    sSql = sSql & " nTasaGracia = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntGracia & "),"
    sSql = sSql & " cLineaCred = (Select (cDescripcion +  space(150) + cLineaCred) from ColocLineaCredito Where cLineaCred = CC.cLineaCred), "
    sSql = sSql & " P2.cPersNombre cFteIngreso, PF.cPersFIPersCod, PF.nPersTipFte, P.cPersCod, P.cPersNombre, P.cPersDireccDomicilio, "
    sSql = sSql & " CN3.cConsDescripcion as cCondicionCred, CN4.cConsDescripcion as cCondicionOtraCred, C.nColocCondicion, C.nColocCondicion2, C.bRefCapInt, "
    sSql = sSql & " nMaxCalend = (Select MAX(nNroCalen) From ColocCalendario Cal Where cCtaCod = C.cCtaCod), "
    sSql = sSql & " C.bMiVivienda, C.bCuotaCom, "
    sSql = sSql & " nTasaIntComp = (Select nTasaInteres from ProductoTasaInteres where cCtaCod = C.cCtaCod AND nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & "), "
    sSql = sSql & " nTasaIntMora = (Select nTasaInteres from ProductoTasaInteres where cCtaCod = C.cCtaCod AND nPrdTasaInteres = " & gColocLineaCredTasasIntMoratNormal & ") "
'    sSQL = sSQL & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
''    sSQL = sSQL & "                    INNER JOIN Colocaciones CC ON CC.cCtaCod = C.cCtaCod "
''    sSQL = sSQL & "                    INNER JOIN PersFteIngreso PF ON PF.cPersCod = PP.cPersCod "
''    sSQL = sSQL & "                    INNER JOIN Persona P2 ON P2.cPersCod = PF.cPersFIPersCod "
''    sSQL = sSQL & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
''    sSQL = sSQL & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSug
''    sSQL = sSQL & "                    INNER JOIN Constante CN ON C.nColocDestino = CN.nConsValor AND CN.nConsCod = " & Trim(Str(gColocDestino))
''    sSQL = sSQL & "                    INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = " & gProducto
''    sSQL = sSQL & "                    INNER JOIN Constante CN3 ON C.nColocCondicion = CN3.nConsValor AND CN3.nConsCod = " & Trim(Str(gColocCredCondicion))
''    sSQL = sSQL & "                    INNER JOIN Constante CN4 ON C.nColocCondicion2 = CN4.nConsValor AND CN4.nConsCod = " & Trim(Str(gColocCredCondicionOtra))
''    sSQL = sSQL & "                    INNER JOIN Producto PRD ON C.cCtaCod = PRD.cCtaCod"
    sSql = sSql & " From ColocacCred C "
    sSql = sSql & " INNER JOIN Producto PRD ON C.cCtaCod = PRD.cCtaCod"
    sSql = sSql & " INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & " INNER JOIN Colocaciones CC ON CC.cCtaCod = C.cCtaCod"
    sSql = sSql & " INNER JOIN ColocFteIngreso CF on CF.cCtaCod=PRD.cCtaCod"
    sSql = sSql & " INNER JOIN PersFteIngreso PF ON PF.cPersCod = PP.cPersCod and PF.cNumFuente=CF.cNumFuente"
    sSql = sSql & " INNER JOIN Persona P2 ON P2.cPersCod = PF.cPersFIPersCod"
    sSql = sSql & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod"
    sSql = sSql & " INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = 2001"
    sSql = sSql & " INNER JOIN Constante CN ON C.nColocDestino = CN.nConsValor AND CN.nConsCod = 3016"
    sSql = sSql & " INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = 1001"
    sSql = sSql & " INNER JOIN Constante CN3 ON C.nColocCondicion = CN3.nConsValor AND CN3.nConsCod = 3015"
    sSql = sSql & " INNER JOIN Constante CN4 ON C.nColocCondicion2 = CN4.nConsValor AND CN4.nConsCod = 3030"
    sSql = sSql & " WHERE C.cCtaCod = '" & psctacod & "' And Prd.nPrdestado = " & gColocEstSug
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosAprobacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaDatosAprobacion:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
        
End Function

' CMACICA_CSTS - 06/11/2003 ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function RecuperaDatosAprobados(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaDatosAprobados
    
    sSql = "Select C.nColocCalendCod, C.nTipoDesembolso, nMaxCalend = (  Select MAX(nNroCalen) From ColocCalendario Cal Where cCtaCod = C.cCtaCod), CC.cLineaCred as cLineaCredCod, "
    sSql = sSql & " CC.dVigencia, CC.nMontoCol, CC.nPlazo, cLineaCred = (  Select (cDescripcion +  space(150) + cLineaCred) From ColocLineaCredito Where cLineaCred = CC.cLineaCred), "
    sSql = sSql & " nMonto     = (  Select  CE.nMonto From  ColocacEstado CE Where  CE.cCtaCod = C.cCtaCod AND CE.nPrdEstado = " & gColocEstAprob & "), PRD.nPrdEstado, P.cPersCod, P.cPersNombre "
    sSql = sSql & " From ColocacCred C   INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                      INNER JOIN Colocaciones CC ON CC.cCtaCod = C.cCtaCod "
    sSql = sSql & "                      INNER JOIN Persona P ON PP.cPersCod = P.cPersCod "
    sSql = sSql & "                      INNER JOIN Producto PRD ON C.cCtaCod = PRD.cCtaCod "
    sSql = sSql & " WHERE C.cCtaCod = '" & psctacod & "'"
    
'    ssql = "Select CC.cLineaCred as cLineaCredCod, CC.dVigencia, CE.nPrdEstado, C.nColocCalendCod, CN2.cConsDescripcion as cTipoCredDescrip, C.nColocDestino, CN.cConsDescripcion as cDestinoDescripcion, "
'    ssql = ssql & " ISNULL(CE.cTipogasto,'F') as cTipogasto, CE.nMonto, CE.nCuotas, CE.nPlazo, CE.nPeriodoGracia, CE.nPeriodoFechaFija, PRD.nTasaInteres, CE.nProxMes, CE.nCalendDinamico, CE.nTipoGracia, CE.nTipoDesembolso, "
'    ssql = ssql & " cAnalista = (select Top 1 cPersNombre from persona where cPersCod = (Select cPersCod From ProductoPersona Where cCtaCod = C.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ")), "
'    ssql = ssql & " nMontoSol = (Select nMonto from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "), "
'    ssql = ssql & " nCuotasSol = (Select nCuotas from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
'    ssql = ssql & " nPlazoSol = (Select nPlazo from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
'    ssql = ssql & " nTasaComp = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & "),"
'    ssql = ssql & " nTasaGracia = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntGracia & "),"
'    ssql = ssql & " cLineaCred = (Select (cDescripcion +  space(150) + cLineaCred) from ColocLineaCredito Where cLineaCred = CC.cLineaCred), "
'    ssql = ssql & " P2.cPersNombre cFteIngreso, PF.cPersFIPersCod, P.cPersCod, P.cPersNombre, P.cPersDireccDomicilio, "
'    ssql = ssql & " CN3.cConsDescripcion as cCondicionCred, CN4.cConsDescripcion as cCondicionOtraCred, C.nColocCondicion, C.nColocCondicion2, C.bRefCapInt, "
'    ssql = ssql & " nMaxCalend = (Select MAX(nNroCalen) From ColocCalendario Cal Where cCtaCod = C.cCtaCod), "
'    ssql = ssql & " C.bMiVivienda, C.bCuotaCom, "
'    ssql = ssql & " nTasaIntComp = (Select nTasaInteres from ProductoTasaInteres where cCtaCod = C.cCtaCod AND nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & "), "
'    ssql = ssql & " nTasaIntMora = (Select nTasaInteres from ProductoTasaInteres where cCtaCod = C.cCtaCod AND nPrdTasaInteres = " & gColocLineaCredTasasIntMoratNormal & ") "
'    ssql = ssql & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
'    ssql = ssql & "                    INNER JOIN Colocaciones CC ON CC.cCtaCod = C.cCtaCod "
'    ssql = ssql & "                    INNER JOIN PersFteIngreso PF ON PF.cPersCod = PP.cPersCod "
'    ssql = ssql & "                    INNER JOIN Persona P2 ON P2.cPersCod = PF.cPersFIPersCod "
'    ssql = ssql & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
'    ssql = ssql & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstAprob
'    ssql = ssql & "                    INNER JOIN Constante CN ON C.nColocDestino = CN.nConsValor AND CN.nConsCod = " & Trim(Str(gColocDestino))
'    ssql = ssql & "                    INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = " & gProducto
'    ssql = ssql & "                    INNER JOIN Constante CN3 ON C.nColocCondicion = CN3.nConsValor AND CN3.nConsCod = " & Trim(Str(gColocCredCondicion))
'    ssql = ssql & "                    INNER JOIN Constante CN4 ON C.nColocCondicion2 = CN4.nConsValor AND CN4.nConsCod = " & Trim(Str(gColocCredCondicionOtra))
'    ssql = ssql & "                    INNER JOIN Producto PRD ON C.cCtaCod = PRD.cCtaCod"
'    ssql = ssql & " WHERE C.cCtaCod = '" & psCtaCod & "' And Prd.nPrdestado = " & gColocEstAprob
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosAprobados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaDatosAprobados:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function RecuperaSolicitud(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaSolicitud
    sSql = "Select C.nColocDestino, C.nColocCondicion, C.nColocCondicion2, C.nColocCondicionProd, P.cPersCod, P.cPersNombre, "
    sSql = sSql & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), "
    sSql = sSql & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdRUC & "'), "
    sSql = sSql & " CE.nMonto, CE.nCuotas, CE.nPlazo, CE.dPrdEstado, "
    sSql = sSql & " cAnalista = (Select cPersCod From ProductoPersona Where cCtaCod = C.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & "),"
    sSql = sSql & " CC.cPersCod as cPersConvenio, CC.cCodModular, CF.cNumFuente, CF.dPersEval dPersFIEvaluacion, C.bRefCapint "
    sSql = sSql & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                    INNER JOIN Producto Prd ON Prd.cCtaCod = C.cCtaCod "
    sSql = sSql & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
    sSql = sSql & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic
    sSql = sSql & "                    LEFT JOIN ColocacConvenio CC ON C.cCtaCod = CC.cCtaCod"
    sSql = sSql & "                    LEFT join ColocFteIngreso CF ON C.cCtaCod = CF.cCtaCod "
    sSql = sSql & " WHERE C.cCtaCod = '" & psctacod & "' AND Prd.nPrdEstado = " & gColocEstSolic
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaSolicitud = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaSolicitud:
    Err.Raise Err.Number, "Recuperar Solicitud", Err.Description
End Function

Public Function ExisteGasto(ByVal psctacod As String, ByVal pnNroCalen As Integer, _
        pnColocCalendApl As Integer, pnCuota As Integer, ByVal nPrdConceptoCod As Long) As Boolean
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    
    sSql = "Select nCuota From colocCalendDet Where cCtaCod = '" & psctacod & "' AND nNroCalen = " & pnNroCalen
    sSql = sSql & " AND nColocCalendApl = " & pnColocCalendApl & " AND nCuota = " & pnCuota & " AND nPrdConceptoCod = " & nPrdConceptoCod
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not R.BOF And Not R.EOF Then
        ExisteGasto = True
    Else
        ExisteGasto = False
    End If
    R.Close
    Set R = Nothing
    
    
End Function

Public Function ExisteGarantia(ByVal psctacod As String, ByVal psNumGarant As String) As Boolean
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorExisteGarantia
    sSql = "Select cNumGarant From ColocGarantia Where cCtaCod = '" & psctacod & "' AND cNumGarant = '" & psNumGarant & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not R.BOF And Not R.EOF Then
        ExisteGarantia = True
    Else
        ExisteGarantia = False
    End If
    R.Close
    Set R = Nothing
    Exit Function

ErrorExisteGarantia:
                  Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'*****************************************************************************************
'***     Rutina:           RecuperaDatosGarantiaCred
'***     Descripcion:      Recupera los datos de la solicitud del Credito
'                          a la Cual se le va a registrar sus Garantias.
'***     Modificado por:        NSSE
'***     Fecha-Tiempo:         11/06/2001 11:30:40 AM
'***     Ultima Modificacion: Lo Ultimo que se Modifico
'*****************************************************************************************

Public Function RecuperaDatosGarantiaCred(ByVal psctacod As String) As ADODB.Recordset
Dim oConecta As DConecta
Dim sSql As String

    On Error GoTo ErrorRecuperaDatosGarantiaCred
    sSql = "Select CN3.cConsDescripcion as cTipoCredDescrip, C.nColocDestino, CN.cConsDescripcion as cDestinoDescripcion , P.cPersCod, P.cPersNombre, CN2.cConsDescripcion as cMonedaDesc, "
    sSql = sSql & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), "
    sSql = sSql & " CE.nMonto "
    sSql = sSql & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
    sSql = sSql & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = '" & gColocEstSolic & "' "
    sSql = sSql & "                    INNER JOIN Constante CN ON C.nColocDestino = CN.nConsValor AND CN.nConsCod = " & gColocDestino
    sSql = sSql & "                    INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,9,1)) = CN2.nConsValor AND CN2.nConsCod = " & gMoneda
    sSql = sSql & "                    INNER JOIN Constante CN3 ON convert(int,substring(C.cCtaCod,6,3)) = CN3.nConsValor AND CN3.nConsCod = " & gProducto
    sSql = sSql & " WHERE C.cCtaCod = '" & psctacod & "' "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosGarantiaCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaDatosGarantiaCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'*****************************************************************************************
'***     Rutina:           RecuperaGarantiasCredito
'***     Descripcion:      Recupera las Garantias del Credito Registradas anteriormente
'***     Modificado por:        NSSE
'***     Fecha-Tiempo:         11/06/2001 11:31:46 AM
'***     Ultima Modificacion: Lo Ultimo que se Modifico
'*****************************************************************************************

Public Function EstadoActual(ByVal psctacod As String) As Integer
Dim R As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorEstadoActual
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = "Select nPrdEstado from Producto Where cCtaCod = '" & psctacod & "'"
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        EstadoActual = R!nPrdEstado
    Else
        EstadoActual = vbNull
    End If
    R.Close
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorEstadoActual:
                  Err.Raise Err.Number, "Error En Proceso EstadoActual", Err.Description
End Function

Public Function GarantiaDeOtroCredito(ByVal psctacod As String, ByVal psNumGarant As String) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConec As DConecta

    sSql = " Select CG.cNumGarant "
    sSql = sSql & " from ColocacRefinanc CR Inner Join ColocGarantia CG ON CR.cCtaCodRef = CG.cCtaCod "
    sSql = sSql & " where CR.cCtaCod = '" & psctacod & "' And CG.cNumGarant = '" & psNumGarant & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set R = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If R.RecordCount > 0 Then
        GarantiaDeOtroCredito = True
    Else
        GarantiaDeOtroCredito = False
    End If
    R.Close
    
End Function

Public Sub EliminarGarantia(ByVal psctacod As String, ByVal psNumGarant As String, _
     ByVal pnMontoGarantia As Double)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminarGarantia
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.BeginTrans
    
    'Si es garantia de otro credito por Refinanciacion
    If Not GarantiaDeOtroCredito(psctacod, psNumGarant) Then
        sSql = "Update Garantias set nGravament = nGravament - " & Format(pnMontoGarantia, "#0.00") & " where cNumgarant = '" & psNumGarant & "' "
        oConecta.ConexionActiva.Execute sSql
    End If
    
    sSql = "DELETE ColocGarantia WHERE cCtaCod = '" & psctacod & "' AND cNumgarant = '" & psNumGarant & "' "
    oConecta.ConexionActiva.Execute sSql
    
    oConecta.ConexionActiva.CommitTrans
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Sub

ErrorEliminarGarantia:
                  Err.Raise Err.Number, "Error En Proceso EliminarGarantia", Err.Description
    
    
End Sub

Public Function RecuperaColocGarantia(ByVal psctacod As String, Optional poConex As ADODB.Connection) As ADODB.Recordset
Dim oConecta As DConecta
Dim sSql As String
        
    On Error GoTo ErrorRecuperaColocGarantia
    sSql = "Select G.cTpoDoc, G.cNroDoc, G.nTpoGarantia, G.IdSupGarant, CG.cNumGarant, CG.cCtaCod, CG.nMoneda, CG.nGravado, ISNULL(G.nGravament,0) AS nGravament "
    sSql = sSql & " from ColocGarantia CG Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant "
    sSql = sSql & " Where cCtaCod = '" & psctacod & "'"
    
   ' Set oConecta = New DConecta
    'oConecta.AbreConexion
   ' If Not poConex Is Nothing Then
  '      oConecta.ConexionActiva = poConex
  '  End If
    If poConex Is Nothing Then
        Set oConecta = New DConecta
        oConecta.AbreConexion
        Set RecuperaColocGarantia = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
    Else
        Set RecuperaColocGarantia = poConex.Execute(sSql)
    End If
    
    'oConecta.CierraConexion
    'Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocGarantia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaMov(ByVal psctacod As String, ByVal psOpeCod As String) As Long
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    sSql = "Select nMovNro from MovCol Where cCtaCod = '" & psctacod & "'"
    sSql = sSql & " And cOpeCod = '" & psOpeCod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If R.RecordCount > 0 Then
        RecuperaMov = IIf(IsNull(R!nMovNro), 0, R!nMovNro)
    Else
        RecuperaMov = 0
    End If
    R.Close
    Set R = Nothing
    
End Function

Public Function RecuperaGarantiasCredito(ByVal psctacod As String) As ADODB.Recordset
Dim oConecta As DConecta
Dim sSql As String

    On Error GoTo ErrorRecuperaGarantiasCredito
    sSql = "Select CG.cNumGarant, CG.nGravado, CG.nMoneda, G.nTpoGarantia, CN.cConsDescripcion cTpoGarDescripcion, "
    sSql = sSql & " G.nMoneda, G.nTasacion, G.nRealizacion, G.nPorGravar, G.cDescripcion, P.cPersNombre, G.cTpoDoc, G.cNroDoc "
    sSql = sSql & " From ColocGarantia CG Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant "
    sSql = sSql & "                       Inner Join Constante CN ON G.nTpoGarantia = CN.nConsvalor And CN.nConsCod = " & Trim(Str(gPersGarantia))
    sSql = sSql & "                       Inner Join PersGarantia PG ON CG.cNumGarant = PG.cNumGarant AND PG.nRelacion = " & Trim(Str(gPersRelGarantiaTitular))
    sSql = sSql & "                       Inner Join Persona P ON PG.cPersCod = P.cPersCod "
    sSql = sSql & " WHERE CG.cCtaCod = '" & psctacod & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantiasCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaGarantiasCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function GarantiaPerteneceACreditoAprobado(ByVal psctacod As String, ByVal psNumagarant As String) As Boolean
Dim oConecta As DConecta
Dim sSql As String
Dim R As ADODB.Recordset

    sSql = "select P.nPrdestado from ColocGarantia CG Inner Join Producto P ON CG.cCtaCod = P.cCtaCod "
    sSql = sSql & " Where P.cCtaCod = '" & psctacod & "' AND CG.cNumgarant='" & psNumagarant & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    
    If R.RecordCount > 0 Then
        If R!nPrdEstado >= gColocEstAprob Then
            GarantiaPerteneceACreditoAprobado = True
        Else
            GarantiaPerteneceACreditoAprobado = False
        End If
    Else
        GarantiaPerteneceACreditoAprobado = False
    End If
End Function

Public Sub NuevaGarantia(ByVal psctacod As String, ByVal psNumGarant As String, psMoneda As String, ByVal pnGravado As Double)
Dim oConecta As DConecta
Dim sSql As String
Dim bTran As Boolean
Dim dGar As DGarantia
Dim R As ADODB.Recordset
Dim nGravado As Double

    On Error GoTo ErrorNuevaGarantia
    bTran = False
    
    
    Set dGar = New DGarantia
    Set R = dGar.RecuperaGarantia(psNumGarant)
    Set dGar = Nothing
    If R.RecordCount > 0 Then
        nGravado = R!nGravament
    End If
    R.Close
    
    'Nuevo Registro en ColocGarantia
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.BeginTrans
    
    bTran = True
    sSql = "INSERT INTO ColocGarantia(cNumGarant, cCtaCod, nMoneda, nGravado, nEstado) "
    sSql = sSql & " VALUES('" & psNumGarant & "','" & psctacod & "'," & psMoneda & "," & Format(pnGravado, "#0.00") & ", 1)"
    Call oConecta.ConexionActiva.Execute(sSql)
    
    If nGravado = 0 Then
        sSql = "UPDATE Garantias SET "
        sSql = sSql & " nGravament = nGravament + " & Format(pnGravado, "#0.00") & ", nEstado = " & gPersGarantEstadoAsignado
        sSql = sSql & " WHERE cNumgarant = '" & psNumGarant & "'"
    Else
        sSql = "UPDATE Garantias SET "
        sSql = sSql & " nGravament = nGravament + " & Format(pnGravado, "#0.00") & ", nEstado = " & gPersGarantEstadoContabilizado
        sSql = sSql & " WHERE cNumgarant = '" & psNumGarant & "'"
    End If
    Call oConecta.ConexionActiva.Execute(sSql)
    
    Call oConecta.ConexionActiva.CommitTrans
    Call oConecta.CierraConexion
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorNuevaGarantia:
    If bTran Then
        Call oConecta.ConexionActiva.RollbackTrans
        Set oConecta = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso NuevaGarantia", Err.Description
End Sub


Public Sub ActualizaGarantias(ByVal psctacod As String, ByVal psNumGarant As String, psMoneda As String, ByVal pnGravado As Double, ByVal pnGravadoAnt As Double)
Dim oConecta As DConecta
Dim sSql As String
Dim bTran As Boolean

    On Error GoTo ErrorNuevaGarantia
    bTran = False
    'Actualiza Registro en ColocGarantia
    sSql = "UPDATE ColocGarantia SET "
    sSql = sSql & " nGravado = " & Format(pnGravado, "#0.00")
    sSql = sSql & " WHERE cNumGarant" & psNumGarant & "' AND cCtaCod = '" & psctacod & "' "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.BeginTrans
    bTran = True
    Call oConecta.ConexionActiva.Execute(sSql)
    
    sSql = "UPDATE Garantias SET "
    sSql = sSql & " nGravament = nGravament - " & Format(pnGravadoAnt, "#0.00")
    sSql = sSql & " WHERE cNumgarant = '" & psNumGarant & "'"
    Call oConecta.ConexionActiva.Execute(sSql)
    sSql = "UPDATE Garantias SET "
    sSql = sSql & " nGravament = nGravament + " & Format(pnGravado, "#0.00") & ", "
    sSql = sSql & " nEstado = " & gPersGarantEstadoAsignado
    sSql = sSql & " WHERE cNumgarant = '" & psNumGarant & "'"
    Call oConecta.ConexionActiva.Execute(sSql)
    Call oConecta.ConexionActiva.CommitTrans
    Call oConecta.CierraConexion
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorNuevaGarantia:
    If bTran Then
        Call oConecta.ConexionActiva.RollbackTrans
        Set oConecta = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso NuevaGarantia", Err.Description
End Sub
Public Function RecuperaAgencia(ByVal psAge As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaAgencia
    sSql = "Select * from Agencias Where cAgeCod = '" & psAge & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaAgencia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaAgencia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
    
End Function

Public Function RecuperaCreditosVigentesCtas(ByVal psPersCod As String)
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCreditosVigentes
    sSql = "Select Prd.cCtaCod "
    sSql = sSql & " FROM Producto Prd Inner Join ProductoPersona PP ON Prd.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & " WHERE PP.cPersCod = '" & psPersCod & "' And Prd.nPrdEstado  in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosVigentesCtas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosVigentes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
Public Function RecuperaCreditosVigentes(ByVal psPersCod As String, Optional ByVal psctacod As String = "", _
    Optional pEstadosEspecificos As Variant = Nothing, Optional ByVal psMoneda As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
Dim i As Integer

    On Error GoTo ErrorRecuperaCreditosVigentes
    sSql = "Select Prd.cCtaCod, CC.nDiasAtraso, Prd.nSaldo, CN.cConsDescripcion cEstado, Prd.nPrdEstado,"
    sSql = sSql & " nPrestamo = (Select nMonto From ColocacEstado Where cCtaCod = Prd.cCtaCod And nPrdEstado = " & gColocEstAprob & "), C.dVigencia, Prd.nTasaInteres"
    sSql = sSql & " FROM Producto Prd Inner Join ProductoPersona PP ON Prd.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "                   Inner Join Persona P ON PP.cPersCod = P.cPersCod "
    sSql = sSql & "                   Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod "
    sSql = sSql & "                   Inner Join Colocaciones C ON C.cCtaCod = Prd.cCtaCod "
    sSql = sSql & "                   Inner Join Constante CN ON CN.nConsValor = Prd.nPrdEstado And CN.nConsCod = " & gColocEstado
    sSql = sSql & " WHERE P.cPersCod = '" & psPersCod & "' "
    
    If psMoneda <> "" Then
        sSql = sSql & " AND SUBSTRING(Prd.cCtaCod,9,1) = '" & psMoneda & "' "
    End If
    If psctacod <> "" Then
        sSql = sSql & " AND Prd.cCtaCod <> '" & psctacod & "' "
    End If
    
    If IsArray(pEstadosEspecificos) Then
        If UBound(pEstadosEspecificos) <> 0 Then
            sSql = sSql & " AND Prd.nPrdEstado in ("
            For i = 0 To UBound(pEstadosEspecificos)
                sSql = sSql & pEstadosEspecificos(i) & ","
            Next i
            sSql = Mid(sSql, 1, Len(sSql) - 1)
            sSql = sSql & ")"
        End If
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosVigentes = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosVigentes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaConsultaCred(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaConsultaCred
    sSql = "Select ISNULL(CC.nCalendDinamico,0) as nCalendDinamico, P.nPrdEstado, ISNULL(CC.bMiVivienda,0) as bMiVivienda, ISNULL(CC.bCuotaCom,0) as bCuotaCom, CC.nDiasAtraso, PFI.cPersNombre cFteIngreso, L.cDescripcion cLineaDesc, PersAna.cPersNombre cAnalista, PersApod.cPersNombre cApoderado, CN.cConsDescripcion cCondicion, Lin.nTasaIni as nTasaGracia, "
    sSql = sSql & " CN2.cConsDescripcion cDestino, P.nTasaInteres, CN3.cConsDescripcion cTipoCuota, C.dVigencia,"
    sSql = sSql & " nNota = (select nColocNota  From ColocCalificacionAnalista Where cCtaCod = P.cCtaCod And dColocNotaFecha = (select Max(dColocNotaFecha) from ColocCalificacionAnalista Where cCtaCod = P.cCtaCod)), "
    sSql = sSql & " CE.nMonto nMontoSol, CE.dPrdEstado dFecSol, CE.nCuotas nCuotasSol, CE.nPlazo nPlazoSol,"
    sSql = sSql & " CE2.nMonto nMontoSug, CE2.dPrdEstado dFecSug, CE2.nCuotas nCuotasSug, CE2.nPlazo nPlazoSug, CE2.nPeriodoGracia nPeriodoGraciaSug,"
    sSql = sSql & " CE3.nMonto nMontoApr, CE3.dPrdEstado dFecApr, CE3.nCuotas nCuotasApr, CE3.nPlazo nPlazoApr, CE3.nPeriodoGracia nPeriodoGraciaApr, CN8.cConsDescripcion cTipoGracia,"
    sSql = sSql & " CN4.cConsDescripcion cMotivoRech, CE5.dPrdEstado dFecCancel, CE6.dPrdEstado dFecJud, CC.cMetLiquidacion, CC.cProtesto,"
    sSql = sSql & " CE7.nPrdEstado nEstRefin, CC.bCargoAuto, CN5.cConsDescripcion cEstActual, CN6.cConsDescripcion as cTipoCredDescrip, "
    sSql = sSql & " CN7.cConsDescripcion cTipoDesemb, CC.nNroProxDesemb, "
    sSql = sSql & " nCuotaSug = (Select sum(nMonto) from ColocCalendDet Cal Where Cal.cCtaCod = P.cCtaCod AND Cal.nNrocalen =1 AND Cal.nColocCalendApl=1 AND Cal.nCuota=1) " & " ,"
    sSql = sSql & " nCuotaApr = (Select sum(nMonto) from ColocCalendDet Cal Where Cal.cCtaCod = P.cCtaCod AND Cal.nNrocalen =(Select nNroCalen From ColocacCred Where cCtaCod = Cal.cCtaCod) AND Cal.nColocCalendApl=1 AND Cal.nCuota=2), "
    sSql = sSql & " nTasaComp = (Select nTasaInteres From ProductoTasaInteres Where nPrdTasaInteres  = 1 AND cCtaCod = P.cCtaCod), "
    sSql = sSql & " nTasaCompVen = (Select nTasaInteres From ProductoTasaInteres Where nPrdTasaInteres  = 6 AND cCtaCod = P.cCtaCod), "
    sSql = sSql & " nTasaMor = (Select nTasaInteres From ProductoTasaInteres Where nPrdTasaInteres  = 3 AND cCtaCod = P.cCtaCod) "
    If CInt(Mid(psctacod, 6, 3)) = gColConsuDctoPlan Then
        sSql = sSql & ", Conv.cCodModular, PConv.cPersNombre cConvenio"
    End If
    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdpersRelac = " & gColRelPersTitular
    sSql = sSql & "                 Left Join ColocFteIngreso FI ON  FI.cCtaCod = P.cCtaCod"
    sSql = sSql & "                 Left Join PersFteIngreso PF ON PF.cNumFuente =  FI.cNumFuente "
    sSql = sSql & "                 Left Join Persona PFI ON PFI.cPersCod = PF.cPersFIPersCod"
    sSql = sSql & "                 Inner Join Colocaciones C ON C.cCtaCod = P.cCtaCod "
    sSql = sSql & "                 Left Join ColocLineaCredito L ON L.cLineaCred = C.cLineaCred"
    sSql = sSql & "                 Left Join ProductoPersona PP2 ON PP2.cCtaCod = P.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "                 Left Join Persona PersAna ON PersAna.cPersCod = PP2.cPersCod"
    sSql = sSql & "                 Left Join ProductoPersona PP3 ON PP3.cCtaCod = P.cCtaCod AND PP3.nPrdPersRelac = " & gColRelPersApoderado
    sSql = sSql & "                 Left Join Persona PersApod ON PersApod.cPersCod = PP3.cPersCod"
    sSql = sSql & "                 Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod"
    sSql = sSql & "                 Left Join Constante CN ON CN.nConsValor = CC.nColocCondicion AND CN.nConsCod = " & gColocCredCondicion
    sSql = sSql & "                 Left Join Constante CN2 ON CN2.nConsValor = CC.nColocDestino AND CN2.nConsCod = " & gColocDestino
    sSql = sSql & "                 Left Join Constante CN3 ON CN3.nConsValor = CC.nColocCalendCod AND CN3.nConsCod = " & gColocTipoCalend
    sSql = sSql & "                 Left Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic
    sSql = sSql & "                 Left Join ColocacEstado CE2 ON CE2.cCtaCod = P.cCtaCod AND CE2.nPrdEstado = " & gColocEstSug
    sSql = sSql & "                 Left Join ColocacEstado CE3 ON CE3.cCtaCod = P.cCtaCod AND CE3.nPrdEstado = " & gColocEstAprob
    sSql = sSql & "                 Left Join ColocacEstado CE4 ON CE4.cCtaCod = P.cCtaCod AND CE4.nPrdEstado = " & gColocEstRech
    sSql = sSql & "                 Left Join Constante CN4 ON CN4.nConsValor = CE4.nMotivoRechazo AND CN4.nConsCod = " & gColocMotivRechazo
    sSql = sSql & "                 Left Join ColocacEstado CE5 ON CE5.cCtaCod = P.cCtaCod AND CE5.nPrdEstado = " & gColocEstCancelado
    sSql = sSql & "                 Left Join ColocacEstado CE6 ON CE6.cCtaCod = P.cCtaCod AND CE6.nPrdEstado = " & gColocEstJudicial
    sSql = sSql & "                 Left Join ColocacEstado CE7 ON CE7.cCtaCod = P.cCtaCod AND CE7.nPrdEstado = " & gColocEstRefNorm
    sSql = sSql & "                 Left Join Constante CN5 ON CN5.nConsValor = P.nPrdEstado AND CN5.nConsCod = " & gColocEstado
    sSql = sSql & "                 Left Join Constante CN6 ON convert(int,substring(P.cCtaCod,6,3)) = CN6.nConsValor AND CN6.nConsCod = " & gProducto
    sSql = sSql & "                 Left Join Constante CN7 ON CN7.nConsValor = CC.nTipoDesembolso AND CN7.nConsCod = " & gColocTiposDesembolso
    sSql = sSql & "                 Left Join Constante CN8 ON CN8.nConsValor = CE3.nTipoGracia AND CN8.nConsCod = " & gColocTiposGracia
    sSql = sSql & "                 Left Join ColocLineaCreditoTasa Lin ON C.cLineaCred = Lin.cLineaCred AND Lin.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntGracia
    If CInt(Mid(psctacod, 6, 3)) = gColConsuDctoPlan Then
        sSql = sSql & "             Left Join ColocacConvenio Conv ON P.cCtaCod = Conv.cCtaCod"
        sSql = sSql & "             Left Join Persona PConv ON PConv.cPersCod = Conv.cPersCod "
    End If
    sSql = sSql & " WHERE P.cCtaCod = '" & psctacod & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaConsultaCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaConsultaCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaDacionesPago() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    sSql = " Select C.nNroGarantRec, C.cCtaCod, P.cPersNombre "
    sSql = sSql & " From ColocGarantRec C Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "             Inner Join Persona P ON P.cPersCod = PP.cPersCod "
    sSql = sSql & "    WHERE C.nEstado = " & gColocGarantRecEstadoRegistrado
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDacionesPago = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaDacionPago(ByVal pnNroDacion As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    sSql = "Select C.*, P.cPersNombre, P.cPersCod From ColocGarantRec C Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "     Inner Join Persona P ON PP.cPersCod = P.cPersCod "
    sSql = sSql & " Where nNroGarantRec = " & pnNroDacion
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDacionPago = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaDacionPagoDetalle(ByVal pnNroDacion As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    sSql = "Select * From ColocGarantRecdet Where nNroGarantRec = " & pnNroDacion
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDacionPagoDetalle = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaDatosDacionPagoCredito(ByVal pnDacion As Long) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    sSql = "Select CG.cCtaCod, CG.nValorTotal, P.cPersNombre, CN.cConsDescripcion as cMoneda, C.nMontoCol,  "
    sSql = sSql & " Prd.nSaldo, CE.nCuotas, CC.nDiasAtraso, CC.cMetLiquidacion, Prd.nTransacc, CC.nCalendDinamico "
    sSql = sSql & " From ColocGarantRec CG Inner Join ProductoPersona PP ON CG.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular
    sSql = sSql & "             Inner Join Persona P ON P.cPersCod = PP.cPersCod "
    sSql = sSql & "             Inner Join Constante CN ON Convert(int,SUBSTRING(CG.cCtaCod,9,1)) = CN.nConsValor And CN.nConsCod = " & gMoneda
    sSql = sSql & "             Inner Join Colocaciones C ON CG.cCtaCod = C.cCtaCod "
    sSql = sSql & "             Inner Join Producto Prd ON Prd.cCtaCod = CG.cCtaCod "
    sSql = sSql & "             Inner Join ColocacEstado CE ON CE.cCtaCod = CG.cCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSql = sSql & "             Inner Join ColocacCred CC ON CC.cCtaCod = CG.cCtaCod "
    sSql = sSql & " Where CG.nNroGarantRec = " & pnDacion
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosDacionPagoCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaProductosDeCredito(Optional ByVal psCab As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    'sSQL = " Select cConsDescripcion, nConsValor from Constante"
    'sSQL = sSQL & " where nConscod = 1001 and (nConsValor % 10) <> 0 and (nConsValor < 232 or nConsValor >= 300)"
    'sSQL = sSQL & " Order by cConsDescripcion DESC"
    
    sSql = " Select C.cConsDescripcion, C.nConsValor, CF.cAbrev from Constante C "
    sSql = sSql & " Inner Join ColocCredConsFiltro CF ON C.nConsCod = CF.nConsCod AND C.nConsValor = CF.nConsValor"
    sSql = sSql & " Where CF.nCodFiltro = " & gCredFiltroProd
    
    If psCab <> "" Then
        sSql = sSql & " AND C.nConsValor like '" & Mid(psCab, 1, 1) & "%'"
    End If
    sSql = sSql & " Order by C.nConsValor  "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaProductosDeCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
Public Function RecuperaProductosDeSolicitudDeCredito() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    
    sSql = " Select C.cConsDescripcion, C.nConsValor from Constante C "
    sSql = sSql & " Inner Join ColocCredConsFiltro CF ON C.nConsCod = CF.nConsCod AND C.nConsValor = CF.nConsValor"
    sSql = sSql & " Where CF.nCodFiltro = " & gCredFiltroProdCab
    sSql = sSql & " Order by C.cConsDescripcion DESC"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaProductosDeSolicitudDeCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaUltimaCuotaEvaluada(ByVal psctacod As String) As Integer
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset
Dim nCuota1 As Integer

    sSql = "select ISNULL(MAX(nCuotaOrig),-1)+7 as nCuota from ColocCalifMiViv "
    sSql = sSql & "     where cCtaCod = '" & psctacod & "'"
    sSql = sSql & "   AND cColocMiVivEval='E'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    nCuota1 = R!nCuota
    
    sSql = "select ISNULL(MAX(nCuotaOrig),0) as nCuota from ColocCalifMiViv "
    sSql = sSql & "     where cCtaCod = '" & psctacod & "'"
    sSql = sSql & "   AND nColocCalendEstado=" & gColocCalendEstadoPagado
    Set R = oConecta.CargaRecordSet(sSql)
    RecuperaUltimaCuotaEvaluada = nCuota1 - R!nCuota
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaProductosDeSolicitudDeCF() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    sSql = " Select cConsDescripcion, nConsValor from Constante"
    sSql = sSql & " where nConscod = 1001 and (nConsValor % 10) <> 0 and (nConsValor < 232 or nConsValor >= 300) AND nConsValor In (" & gColCFComercial & "," & gColCFPYME & ")"
    sSql = sSql & " Order by cConsDescripcion DESC"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaProductosDeSolicitudDeCF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaTitularCredito(ByVal scCod As String) As String
Dim sSql As String
Dim oConecta As DConecta
Dim rs As New ADODB.Recordset

    sSql = " select cperscod from productopersona where cctacod='" & Trim(scCod) & "' and nprdpersrelac=" & gColRelPersTitular
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    If rs.BOF Then
        RecuperaTitularCredito = ""
    Else
        RecuperaTitularCredito = rs!cPersCod
    End If
    rs.Close
    Set rs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Private Sub Class_Initialize()
Dim CIni As COMConecta.DCOMClasIni
    Set CIni = New COMConecta.DCOMClasIni
    gConsPersona = CIni.BasePersonas
    gConsComunes = CIni.BaseComunes
    gConsImagenes = CIni.BaseImagenes
    Set CIni = Nothing
End Sub

Public Function RecuperaCorresponCred(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaCorresponCred
    sSql = "Select * from Corresponsalia where cCtacod = '" & psctacod & "' and nFlag=0"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCorresponCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCorresponCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Function ObtieneMonto_Validate(ByVal psctacod As String, ByVal pnMontoPago As Currency) As Boolean
    'Function que valida el monto no sobre pasa el pendiente
    Dim sSql As String
    Dim oConec As DConecta
    Dim rs_Monto As ADODB.Recordset
    Dim nMonto As Currency
    On Error GoTo ErrHandler
        
    sSql = sSql & " SELECT SUM(CD.NMONTO-CD.NMONTOPAGADO) AS RESULT"
    sSql = sSql & " FROM COLOCCALENDDET CD"
    sSql = sSql & " INNER JOIN COLOCCALENDARIO C ON C.CCTACOD=CD.CCTACOD AND CD.NNROCALEN=C.NNROCALEN AND"
    sSql = sSql & " C.NCOLOCCALENDAPL = CD.NCOLOCCALENDAPL And C.nCuota = CD.nCuota"
    sSql = sSql & " WHERE C.CCTACOD='" & psctacod & "' AND C.NCOLOCCALENDAPL=1 AND C.NCOLOCCALENDESTADO=0"
        Set oConec = New DConecta
        oConec.AbreConexion
        Set rs_Monto = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        
        If Not rs_Monto.EOF And Not rs_Monto.BOF Then
            nMonto = rs_Monto!result
        End If
        Set rs_Monto = Nothing
        
        If pnMontoPago > nMonto Then
           ObtieneMonto_Validate = False
        Else
            ObtieneMonto_Validate = True
        End If
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    If Not rs_Monto Is Nothing Then Set rs_Monto = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function InsertaColocacConvenioRegDevolucion(ByVal psPersCodIns As String, ByVal psPersCod As String, _
ByVal pdRegistro As Date, ByVal psCodAge As String, ByVal pnMoneda As Moneda, ByVal pnMonto As Currency, ByVal psNroDoc As String, Optional ByVal pnMovNroReg As Long)
Dim sql As String
Dim oCon As DConecta
Set oCon = New DConecta

oCon.AbreConexion

sql = "INSERT INTO ColocacConvenioRegDev (cPersCodIns, cPersCod, dRegistro, cCodAge, nMoneda, nMonto, cNroDoc, nMovNroReg)" _
    & " Values ('" & psPersCodIns & "','" & psPersCod & "','" & Format(pdRegistro, "mm/dd/yyyy") & "','" _
    & psCodAge & "'," & pnMoneda & "," & pnMonto & ",'" & psNroDoc & "'," & pnMovNroReg & ")"

oCon.Ejecutar sql

oCon.CierraConexion
End Function
Public Function ActualizaColocacConvenioRegDevolucion(ByVal psPersCodIns As String, ByVal psPersCod As String, _
ByVal pdRegistro As Date, ByVal psCodAge As String, ByVal pnMoneda As Moneda, ByVal pnMonto As Currency, ByVal psNroDoc As String)
Dim sql As String
Dim oCon As DConecta
Set oCon = New DConecta

oCon.AbreConexion

sql = "UPDATE ColocacConvenioRegDev SET nMoneda=" & pnMoneda & ", nMonto = " & pnMonto & ", cNroDoc = '" & psNroDoc & "' " _
    & " WHERE cPersCodIns ='" & psPersCodIns & "' and cPersCod='" & psPersCod & "' and dRegistro='" & Format(pdRegistro, "mm/dd/yyyy") & "'"
    
oCon.Ejecutar sql

oCon.CierraConexion
End Function

Public Function ActualizaColocacConvenioRegDevOpe(ByVal psPersCodIns As String, ByVal psPersCod As String, _
ByVal pdRegistro As Date, ByVal pnMovNro As Currency)
Dim sql As String
Dim oCon As DConecta
Set oCon = New DConecta

oCon.AbreConexion

sql = "UPDATE ColocacConvenioRegDev SET nMovNro=" & pnMovNro & ", nPendiente = 1  " _
    & " WHERE cPersCodIns ='" & psPersCodIns & "' and cPersCod='" & psPersCod & "' and dRegistro='" & Format(pdRegistro, "mm/dd/yyyy") & "'"
    
oCon.Ejecutar sql

oCon.CierraConexion
End Function

Public Function ExtornaRegDevConvOpe(ByVal pnMovNro As Currency)
Dim sql As String
Dim oCon As DConecta
Set oCon = New DConecta

oCon.AbreConexion

sql = "UPDATE ColocacConvenioRegDev SET nMovNro=NULL, nPendiente = 0  " _
    & " WHERE nMovNro = " & pnMovNro & ""
    
oCon.Ejecutar sql

oCon.CierraConexion
End Function

Public Function EliminaColocacConvenioRegDevolucion(ByVal psPersCodIns As String, ByVal psPersCod As String, ByVal pdRegistro As Date)
Dim sql As String
Dim oCon As DConecta
Set oCon = New DConecta

oCon.AbreConexion

sql = "DELETE FROM ColocacConvenioRegDev " _
    & " WHERE cPersCodIns ='" & psPersCodIns & "' and cPersCod='" & psPersCod & "' and dRegistro='" & Format(pdRegistro, "mm/dd/yyyy") & "'"
    
oCon.Ejecutar sql

oCon.CierraConexion
End Function

Public Function GetColocacConvRegDevolucion(ByVal psCodPersIns As String) As ADODB.Recordset
Dim sql As String
Dim oCon As DConecta

sql = "select  C.cPersCod, cPersNombre, CONVERT(CHAR(10), c.dRegistro,103) as dRegistro, c.nMonto, c.cNroDoc, " _
    & "        RTRIM(cConsDescripcion) + SPACE(75) + CONVERT(CHAR(1),nMoneda) AS cMoneda, '1' as cNuevo, nMovNroReg " _
    & " from    ColocacConvenioRegDev C " _
    & "         JOIN PERSONA P ON C.cPerscod = P.cPersCod " _
    & "         JOIN CONSTANTE CC ON CC.nConsValor = C.nMoneda and CC.NCONSCOD =1011 " _
    & " WHERE   (C.nPendiente IS NULL OR C.nPendiente<>'1')  and C.cPersCodIns='" & psCodPersIns & "' " _
    & "         and exists (SELECT M.nMovNro FROM MOV M WHERE M.nMovNro = nMovNroReg and M.nMovflag=0 ) " _
    & " ORDER BY P.CPERSNOMBRE "

Set oCon = New DConecta
oCon.AbreConexion
Set GetColocacConvRegDevolucion = oCon.CargaRecordSet(sql)

oCon.CierraConexion
Set oCon = Nothing

End Function
Public Function GetPersDevConv(ByVal psPersCod As String, ByVal pnMoneda As Moneda) As ADODB.Recordset
Dim sql As String
Dim oCon As DConecta

Set oCon = New DConecta

sql = "select   C.cPersCod, C.cPersCodIns, p.cPersNombre," _
    & "         CONVERT(CHAR(10), c.dRegistro,103) AS dRegistro, c.cNroDoc, c.nMonto,C.nMovNro as nMovNroReg " _
    & " from    ColocacConvenioRegDev C " _
    & "         JOIN PERSONA P ON C.cPersCodIns = P.cPersCod " _
    & "         JOIN CONSTANTE CC ON CC.nConsValor = C.nMoneda and CC.NCONSCOD =1011 " _
    & " WHERE   (C.nPendiente IS NULL OR C.nPendiente<>'1')  and C.cPersCod ='" & psPersCod & "' and C.nMoneda=" & pnMoneda & " " _
    & " order by C.dRegistro"

oCon.AbreConexion

Set GetPersDevConv = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function RecuperaEstadoCredito(ByVal psctacod As String) As Integer
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "Select nPrdEstado From Producto Where cCtaCod='" & psctacod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        RecuperaEstadoCredito = rs!nPrdEstado
    End If
    Set rs = Nothing
End Function


Public Function ListaKardex(ByVal psctacod As String) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
    
    sSql = "Select C.nCuota, "
    sSql = sSql & " MontoCuota=(Select Sum(nMontoPagado) From ColocCalendDet Where cCtaCod=C.cCtaCod and nNroCalen=C.nNroCalen and nColocCalendApl=C.nColocCalendApl and nCuota=C.nCuota),"
    sSql = sSql & " Capital=(Select Sum(nMontoPagado) From ColocCalendDet Where cCtaCod=C.cCtaCod and nNroCalen=C.nNroCalen and nColocCalendApl=C.nColocCalendApl and nCuota=C.nCuota and nPrdConceptoCod in (1000,1010)),"
    sSql = sSql & " Interes=(Select Sum(nMontoPagado) From ColocCalendDet Where cCtaCod=C.cCtaCod and nNroCalen=C.nNroCalen and nColocCalendApl=C.nColocCalendApl and nCuota=C.nCuota and nPrdConceptoCod in (1100,1107)),"
    sSql = sSql & " Mora=(Select Sum(nMontoPagado) From ColocCalendDet Where cCtaCod=C.cCtaCod and nNroCalen=C.nNroCalen and nColocCalendApl=C.nColocCalendApl and nCuota=C.nCuota and nPrdConceptoCod in (1101,1108)),"
    sSql = sSql & " Gastos=(Select Sum(nMontoPagado) From ColocCalendDet Where cCtaCod=C.cCtaCod and nNroCalen=C.nNroCalen and nColocCalendApl=C.nColocCalendApl and nCuota=C.nCuota and nPrdConceptoCod like '12%'),"
    sSql = sSql & " Atraso = DateDiff(Day, C.dPago, dVenc)"
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join ColocCalendario C on C.cCtaCod=P.cCtaCod"
    sSql = sSql & " Where P.cCtaCod='" & psctacod & "' and C.nColocCalendEstado=1 and C.nColocCalendApl=1"
    sSql = sSql & " and C.nNroCalen=(Select Max(nNroCalen) From ColocCalendario Where cCtaCod='" & psctacod & "')"
    sSql = sSql & " Order By C.nCuota"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaKardex = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function CargarProductos() As Recordset
    Dim sSql As String
    Dim oConec As DConecta
    
    
    sSql = "Select nConsValor,cConsDescripcion"
    sSql = sSql & " From Constante"
    sSql = sSql & " Where nconsCod = 1001 And nConsValor <> 1001"
    sSql = sSql & " Order By nConsValor"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set CargarProductos = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function VerificarVencimientoCredito(ByVal psctacod As String) As Boolean
    ' funcion que verifica si el credito esta vencido o no
    Dim sSql As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim nDiasAtraso As Integer
    Dim nDiasMinimo As Integer
    
    sSql = "Select nParamValor From ColocParametro Where nParamVar=3101"
    Set oConec = New DConecta
    oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        nDiasMinimo = rs!nParamValor
    End If
    Set rs = Nothing
    VerificarVencimientoCredito = False
    sSql = "Select nDiasAtraso From ColocacCred Where cCtaCod='" & psctacod & "'"
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    If Not rs.EOF And Not rs.BOF Then
        If rs!nDiasAtraso < nDiasMinimo Then
            VerificarVencimientoCredito = True
        Else
            VerificarVencimientoCredito = False
        End If
    End If
    Set rs = Nothing
End Function

Public Function RecuperaDetallePago(ByVal psctacod As String) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
    
    sSql = "Select nMovNro,cUsuario,cOpeCod,dFecPago,nNroCuota,nMontoPagado,nCapital,nInteres,nMora,"
    sSql = sSql & "  nGastos , nITF, nDiasMora, nSaldoCap"
    sSql = sSql & " From ("
    sSql = sSql & " Select MC.nMovNro, right(M.cMovNro,4) as cUsuario, MC.cOpeCod, dbo.FechaMov(M.cMovNro) as dFecPago,"
    sSql = sSql & " MD.nNroCuota,"
    sSql = sSql & " nMontoPagado = (Select ISNULL(SUM(nMonto),0)"
    sSql = sSql & " From MovColDet"
    sSql = sSql & " Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro and cOpeCod not like '107[123456789]%' and nNroCuota=MD.nNroCuota)+(Select isnull(Sum(nMovImporte),0) From movOpeVarias Where nMovNro= MC.nMovNro),"
    sSql = sSql & " nCapital = (Select ISNULL(SUM(nMonto),0)"
    sSql = sSql & " From MovColDet"
    sSql = sSql & " Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod in(1000,3000) and cOpeCod not like '107[123456789]%' and nNroCuota=MD.nNroCuota),"
    sSql = sSql & " nInteres = (Select ISNULL(SUM(nMonto),0)"
    sSql = sSql & " From MovColDet"
    sSql = sSql & " Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in (1100,1102,1103,1104,1105,3100) and nNroCuota=MD.nNroCuota),"
    sSql = sSql & " nMora = (Select ISNULL(SUM(nMonto),0)"
    sSql = sSql & " From MovColDet"
    sSql = sSql & " Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod in(1101,3101) and nNroCuota=MD.nNroCuota),"
    sSql = sSql & " nGastos = (Select ISNULL(SUM(nMonto),0)"
    sSql = sSql & " From MovColDet"
    sSql = sSql & " Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod like  '12%' and nNroCuota=MD.nNroCuota),"
    sSql = sSql & " nITF = (Select ISNULL(SUM(nMonto),0)"
    sSql = sSql & " From MovColDet"
    'sSQL = sSQL & " Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in  (20, 21) and nNroCuota=MD.nNroCuota), MC.nDiasMora , MC.nSaldoCap"
     sSql = sSql & " Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in  (20, 21) and "
     sSql = sSql & " (Select Max(nNroCuota) From MovColDet Where nMovNro=MC.nMovNro)=MD.nNroCuota), MC.nDiasMora , MC.nSaldoCap "
    sSql = sSql & " from MovCol MC"
    sSql = sSql & " Inner Join Mov M ON M.nMovNro = MC.nMovNro"
    sSql = sSql & " Inner Join MovColDet MD on MD.nMovNro=M.nMovNro and MC.cCtaCod=MD.cCtaCod"
    sSql = sSql & " where MC.cOpeCod like '100%' AND MC.cCtaCod = '" & psctacod & "'  AND M.nMovFlag<>2 AND"
    sSql = sSql & " (SUBSTRING(MC.cOpeCod,1,4) in ('1002','1003','1004','1005','1006','1007','1070'))"
    sSql = sSql & " ) X"
    sSql = sSql & " Where nNroCuota <> 0"
    sSql = sSql & " Group By nMovNro,cUsuario,cOpeCod,dFecPago,nNroCuota,nMontoPagado,nCapital,nInteres,nMora,"
    sSql = sSql & " nGastos , nITF, nDiasMora, nSaldoCap"
    sSql = sSql & " Order by nMovNro"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set RecuperaDetallePago = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ValidaSugAprobacion(ByVal psctacod As String) As Boolean
'Funcion que valida que en el momento de sugerir pueda aprobar
    Dim sSql As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim bValor As Boolean
    
    bValor = True
    
    sSql = "Select  Count(*) as nCantidad"
    sSql = sSql & " From ColocacEstado"
    sSql = sSql & " where cctacod='" & psctacod & "' and nPrdEstado=2002"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        If rs!nCantidad > 0 Then
            bValor = False
        Else
            bValor = True
        End If
    End If
    Set rs = Nothing
    ValidaSugAprobacion = bValor
End Function


Public Function GetDatosCheque(ByVal psCodCheque As String) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
    
    sSql = "Select MD.nMovNro,MD.cDocNro as cCheque,Pers.cPersNombre,AG.cAgeDescripcion as cAgencia,           "
    sSql = sSql & "            MOV.nMovImporte as nMonto,MOV.nMoneda,Ag.cAgeCod,MOI.cCtaIfCod,Pers.cPersCod"
    sSql = sSql & " From MovDoc MD"
    sSql = sSql & "           Inner Join MovObjIF MOI on MOI.nMovNro=MD.nMovNro"
    sSql = sSql & "           Inner Join Persona Pers on Pers.cPersCod=MOI.cPersCod"
    sSql = sSql & "           Inner Join MovObjAreaAgencia MOA on MOA.nMovNro=MD.nMovNro"
    sSql = sSql & "           Inner Join Agencias  AG on AG.cAgeCod=MOA.cAgeCod"
    sSql = sSql & "           Inner Join MovOpeVarias MOV on MOV.nMovNro=MD.nMovNro and MOV.cNroDoc= MD.cDocNro"
    sSql = sSql & "           Inner Join Mov M on M.nMovNro=MD.nMovNro"
    sSql = sSql & "  Where MD.cDocNro='" & psCodCheque & "' and MD.nDocTpo=47 and M.nMovFlag=0 and Not Exists(Select * From MovRef Where nMovNroRef=M.nMovNro)"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set GetDatosCheque = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


