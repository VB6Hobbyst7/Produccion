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

Public Function RecuperaChequeCreditos(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pnMoneda As Integer) As ADODB.Recordset
    
Dim sSQL As String
Dim oConec As DConecta

    sSQL = "select '',P.cPersNombre, DR.cIFCta, DR.cNroDoc,  DR.nMonto, C3.cConsDescripcion cProducto, "
    sSQL = sSQL & " C2.cConsDescripcion cEstado, DR.dValorizacion, DR.nEstado,DR.cPersCod "
    sSQL = sSQL & " from DocRec DR"
    sSQL = sSQL & " Inner Join Persona P ON P.cPersCod = DR.cPersCod"
    sSQL = sSQL & " Inner Join Constante C2 ON C2.nConsValor = DR.nEstado AND C2.nConsCod = 1007"
    sSQL = sSQL & " Inner Join ColocCredConsFiltro CF ON CF.nConsValor = DR.nProducto AND CF.nConsCod = 1001"
    sSQL = sSQL & " Inner Join Constante C3 ON C3.nConsCod = 1001 AND C3.nConsValor = CF.nConsValor"
    sSQL = sSQL & " Where  DR.dValorizacion >= '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND DR.dValorizacion <= '" & Format(pdFecFin, "mm/dd/yyyy") & "' "
    sSQL = sSQL & "  AND DR.nMoneda = " & pnMoneda & " AND DR.nTpoDoc = 47 "
    sSQL = sSQL & " Order by P.cPersNombre,DR.dValorizacion"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set RecuperaChequeCreditos = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
End Function

Public Function RecuperaPagosRealizados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    sSQL = " Select MC.nMovNro, MC.cOpeCod, dbo.FechaMov(M.cMovNro) as dFecPago,"
    sSQL = sSQL & " nNroCuota = (Select MAX(nNroCuota) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro),"
    sSQL = sSQL & " nMontoPagado = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro),"
    sSQL = sSQL & " nCapital = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod = 1000 ),"
    sSQL = sSQL & " nInteres = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in (1100,1102,1103,1104,1105)),"
    sSQL = sSQL & " nMora = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod = 1101),"
    sSQL = sSQL & " nGastos = (Select ISNULL(SUM(nMonto),0) from MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod like  '12%'),"
    sSQL = sSQL & " MC.nDiasMora , MC.nSaldoCap"
    sSQL = sSQL & " from MovCol MC Inner Join Mov M ON M.nMovNro = MC.nMovNro"
    sSQL = sSQL & " where MC.cOpeCod like '100%' AND MC.cCtaCod = '" & psCtaCod & "' "
    sSQL = sSQL & " AND (SUBSTRING(MC.cOpeCod,1,4) in ('1002','1003','1004','1005','1006','1007','1070'))"
    sSQL = sSQL & " Order by MC.nMovNro "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaPagosRealizados = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function UltimaCuotaCanceladaOriginal(ByVal psCtaCod As String) As Integer
Dim sSQL As String
Dim oCon As DConecta
Dim R As ADODB.Recordset

    sSQL = "Select ISNULL(MAX(nCuotaOrig),0) as nUltCuota From ColocCalifMiViv where cCtaCod = '" & psCtaCod & "' AND nColocCalendEstado = " & gColocCalendEstadoPagado
    Set oCon = New DConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    UltimaCuotaCanceladaOriginal = R!nUltCuota
    R.Close
    Set R = Nothing
End Function

Public Function SaldoPactadoCredito(ByVal psCtaCod As String) As Double
Dim sSQL As String
Dim oCon As DConecta
Dim R As ADODB.Recordset

    sSQL = "Select ISNULL(SUM(CD.nMonto),0) as nSaldo "
    sSQL = sSQL & " from ColocCalendDet CD"
    sSQL = sSQL & " Where CD.nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CD.cCtaCod )"
    sSQL = sSQL & " AND CD.nPrdConceptoCod = 1000 AND CD.nColocCalendApl = 1"
    sSQL = sSQL & " AND CD.cCtaCod = '" & psCtaCod & "'"

    Set oCon = New DConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    SaldoPactadoCredito = R!nSaldo
    R.Close
End Function

Public Function CuotaAprobada(ByVal psCtaCod As String) As Double
Dim sSQL As String
Dim oCon As DConecta
Dim R As ADODB.Recordset

    
    sSQL = "Select dbo.CuotaAprobada('" & psCtaCod & "') as nCuotaAprobada "
    Set oCon = New DConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    
    CuotaAprobada = IIf(IsNull(R!nCuotaAprobada), 0, R!nCuotaAprobada)
    
End Function

Public Function ExisteMetaAnalista(ByVal psPersCod As String, ByVal pdFechaIni As Date, _
    ByVal pdFechaFin As Date, ByVal pnMoneda As Integer, ByVal pnTipoMeta As Integer) As Boolean

Dim sSQL As String
Dim oCon As DConecta
Dim R As ADODB.Recordset

    sSQL = " Select * from ColocMetasAnalista "
    sSQL = sSQL & "  Where cPersCod = '" & psPersCod & "' AND nTipoMeta = " & pnTipoMeta
    sSQL = sSQL & "  AND dInicial <= '" & Format(pdFechaIni, "mm/dd/yyyy") & "' AND dFinal >= '" & Format(pdFechaIni, "mm/dd/yyyy") & "' "
    sSQL = sSQL & "  and nMoneda = " & pnMoneda
    
    
    
    Set oCon = New DConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    If R.RecordCount > 0 Then
        ExisteMetaAnalista = True
    Else
        ExisteMetaAnalista = False
    End If
    R.Close
End Function

Public Function CreditosCanceladoConDesembolso(ByVal psCtaCod As String, ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSQL As String
Dim oConec As DConecta

    sSQL = "Select cCtaCod, nMonto from MovCol where nMovnro = " & pnMovNro & " and cCtaCod <> '" & psCtaCod & "'"
    Set oConec = New DConecta
    oConec.AbreConexion
    Set CreditosCanceladoConDesembolso = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function PerteneceADesembolsoConCancelacion(ByVal pnMovNro As Long, ByVal psCtaCod As String) As Boolean
Dim oConecta As DConecta
Dim sSQL As String
Dim R As ADODB.Recordset
    sSQL = "Select cCtaCod,cOpeCod from MovCol where nMovnro = " & pnMovNro & " and cOpeCod in ('" & gCredDesembEfec & "','" & gCredDesembCtaNueva & "','" & gCredDesembCtaExist & "','" & gCredDesembCtaExistDOA & "','" & gCredDesembCtaNuevaDOA & "') and cCtaCod <> '" & psCtaCod & "'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
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

Public Function RecuperaCuentasAho(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConec As DConecta
    
    sSQL = "Select cCodCtaAho from ColocCargoAutoma Where cCtaCod = '" & psCtaCod & "' "
    sSQL = sSQL & " AND cestado = '1' Order by nOrden "
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set RecuperaCuentasAho = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing

End Function

Public Function RecuperaCreditosCargoAutomatico(ByVal pdHoy As Date) As ADODB.Recordset
Dim sSQL As String
Dim oConec As DConecta
    
    
    sSQL = "Select ' ',C.cCtaCod, P.cPersNombre,  "
    sSQL = sSQL & " nMonto = (Select SUM(nMonto-nMontoPagado) From ColocCalendDet CD"
    sSQL = sSQL & "     Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen And nColocCalendApl = CC.nColocCalendApl"
    sSQL = sSQL & "     AND nCuota = CC.nCuota), Prd.nPrdEstado, C.cMetLiquidacion "
    sSQL = sSQL & " from Producto Prd Inner Join ColocacCred C ON Prd.cCtaCod = C.cCtaCod "
    sSQL = sSQL & " INNER Join ColocCalendario CC ON CC.cCtaCod = C.cCtaCod AND CC.nNroCalen = C.nNroCalen"
    sSQL = sSQL & " INNER Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & " INNER Join Persona P ON PP.cPersCod = P.cPersCod"
    sSQL = sSQL & " Where CC.nColocCalendEstado = " & gColocCalendEstadoPendiente & " AND CC.nColocCalendApl = " & gColocCalendAplCuota & " AND CC.dVenc = '" & Format(pdHoy, "mm/dd/yyyy") & "' "
    sSQL = sSQL & " AND C.cCtaCod in (Select cCtaCod from ColocCargoAutoma)"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set RecuperaCreditosCargoAutomatico = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
End Function

Public Function CodigosModulares(ByVal psPersCodTitular As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    sSQL = "Select C.cCodModular From ColocacConvenio C Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular & " AND PP.cPersCod = '" & psPersCodTitular & "' Order by C.cCodModular"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set CodigosModulares = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function DefineCondicionCredito(ByVal psPersCod As String, Optional ByVal pnProducto As Producto = 0) As ColocCredCondicion
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConn As DConecta
Dim nCredNumVig As Integer
Dim nCredNumCancelados As Integer

    'Define la Condicion del Credito por Producto
    If pnProducto <> 0 Then
        sSQL = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod Where PP.cPersCod = '" & psPersCod & "' "
        sSQL = sSQL & " AND PP.nPrdPersRelac = " & gColRelPersTitular
        sSQL = sSQL & " AND SUBSTRING(P.cCtaCod,6,2) = '" & Trim(Str(pnProducto)) & "' "
        sSQL = sSQL & " AND P.nPrdEstado in (" & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & ")"
    Else
        sSQL = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod  Where cPersCod = '" & psPersCod & "' "
        sSQL = sSQL & " AND PP.nPrdPersRelac = " & gColRelPersTitular
        sSQL = sSQL & " AND P.nPrdEstado in (" & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & ")"
    End If
    Set oConn = New DConecta
    oConn.AbreConexion
    Set R = oConn.CargaRecordSet(sSQL)
    nCredNumVig = IIf(IsNull(R!nTotal), 0, R!nTotal)
    R.Close
    If pnProducto <> 0 Then
        sSQL = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod Where cPersCod = '" & psPersCod & "' "
        sSQL = sSQL & " AND PP.nPrdPersRelac = " & gColRelPersTitular
        sSQL = sSQL & " AND SUBSTRING(P.cCtaCod,6,2) = '" & Trim(Str(pnProducto)) & "' "
        sSQL = sSQL & " AND P.nPrdEstado not in (" & gColocEstSolic & "," & gColocEstSug & "," & gColocEstAprob & "," & gColocEstRetirado & "," & gColocEstRech & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & ")"
    Else
        sSQL = "Select Count(PP.cctaCod) as nTotal from ProductoPersona PP Inner Join Producto P ON P.cCtaCod = PP.cCtaCod Where cPersCod = '" & psPersCod & "' "
        sSQL = sSQL & " AND PP.nPrdPersRelac = " & gColRelPersTitular
        sSQL = sSQL & " AND P.nPrdEstado not in (" & gColocEstSolic & "," & gColocEstSug & "," & gColocEstAprob & "," & gColocEstRetirado & "," & gColocEstRech & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & ")"
    End If
    Set R = oConn.CargaRecordSet(sSQL)
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

Public Function RecuperaNivelesAprUsuario(ByVal psCodUser As String, ByVal pnMonto As Double, _
    psMoneda As String, ByVal psProducto As String, ByVal pbRefinanc As Boolean) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    
    sSQL = " Select CN.cCodNiv, CA.cProduct, CA.cTpoCred, CA.nNivel, CA.nMontoMax, CA.nMontoMin"
    sSQL = sSQL & " From ColocCredPersNivelesApr CN Inner Join RRHH RH ON CN.cPersCod = RH.cPersCod"
    sSQL = sSQL & " Inner join ColocCredNivelesApr CA ON CN.cCodNiv = CA.cCodNiv"
    sSQL = sSQL & " Where RH.cUser = '" & psCodUser & "' AND CA.cProduct = '" & psProducto & "' AND CA.cTpoCred = '" & IIf(pbRefinanc, Trim(Str(gColocCredTipoRefinanciado)), Trim(Str(gColocCredTipoNormal))) & "' "
    sSQL = sSQL & " AND CA.nMontoMax >= " & Format(pnMonto, "#0.00") & " AND CA.nMontoMin <= " & Format(pnMonto, "#0.00") & " AND CA.nMoneda = " & psMoneda
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaNivelesAprUsuario = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaMovimientosAhorros(ByVal pnMov As Long, Optional ByVal pbSinMovRef As Boolean = False) As ADODB.Recordset
Dim sSQL As String
Dim oConec As DConecta

    If Not pbSinMovRef Then
        sSQL = "Select MC.nMovNro,MC.cOpeCod, MC.nMonto, MC.cCtaCod "
        'sSql = "Select MC.cOpeCod, SUM(MC.nMonto) nMonto "
        sSQL = sSQL & " From MovCap MC"
        sSQL = sSQL & " Inner Join MovRef MR ON MC.nMovNro = MR.nMovNro"
        sSQL = sSQL & " Where MR.nMovNroRef = " & pnMov
        sSQL = sSQL & " Order By MC.nMovNro "
        'sSql = sSql & " Group By MC.cOpeCod"
    Else
        sSQL = "Select MC.nMovNro,MC.cOpeCod, MC.nMonto, MC.cCtaCod "
        'sSql = "Select MC.cOpeCod, SUM(MC.nMonto) nMonto "
        sSQL = sSQL & " From MovCap MC"
        sSQL = sSQL & " Where MC.nMovNro = " & pnMov
        sSQL = sSQL & " Order By MC.nMovNro "
        
    End If
    Set oConec = New DConecta
    oConec.AbreConexion
    Set RecuperaMovimientosAhorros = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function RecuperaDatosExtorno(ByVal pnMovNro As Long, ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    sSQL = "Select M.nCredEstado, SUM(MD.nMonto) as nCapital, "
    sSQL = sSQL & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
    sSQL = sSQL & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro AND MD.nPrdConceptoCod = 1000 Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psCtaCod & "' AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psCtaCod & "' "
    sSQL = sSQL & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSQL)
    If RecuperaDatosExtorno.RecordCount = 0 Then
        sSQL = "Select M.nCredEstado, 0 as nCapital, "
        sSQL = sSQL & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
        sSQL = sSQL & " From MovCol M Inner Join MovColDet MD ON M.nMovNro = MD.nMovNro AND MD.nPrdConceptoCod = 1100 Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psCtaCod & "' AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psCtaCod & "' "
        sSQL = sSQL & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
        Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function


Public Function RecuperaDatosMetasAnalistaReporte(ByVal psPersCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaDatosMetasAnalistaReporte
    sSQL = "Select DISTINCT M.cPersCod,M.dInicial,M.dFinal,M.nTipoMeta, C.cConsDescripcion as cTipoMeta, M.nMonto, M.nTipoCred, C2.cConsDescripcion cTipoCred, M.nMontoAlc, C3.cConsDescripcion as sMoneda "
    sSQL = sSQL & " From ColocMetasAnalista M Inner Join Constante C ON M.nTipoMeta = C.nConsValor AND C.nConsCod = " & gColocTipoMetas
    sSQL = sSQL & "                           Inner Join Constante C2 ON M.nTipoCred = C2.nConsValor  AND C2.nConsCod = 1001"
    sSQL = sSQL & "                           Inner Join Constante C3 ON C3.nConsValor = M.nMoneda AND C3.nConsCod = " & gMoneda
    sSQL = sSQL & " Inner Join ColocCredConsFiltro CF ON CF.nConsCod = C2.nConsCod AND CF.nCodFiltro =1 "
    sSQL = sSQL & " Where M.cPersCod = '" & psPersCod & "'  ORDER BY M.dInicial,cTipoMeta,sMoneda,nTipoCred"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosMetasAnalistaReporte = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosMetasAnalistaReporte:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaColocFteIngreso(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaColocacFteIngreso
    sSQL = "Select * from ColocFteIngreso"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocFteIngreso = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocacFteIngreso:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaDatosMetasAnalistaCab(ByVal psPersCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaDatosMetasAnalistaCab
    sSQL = " Select M.cPersCod,M.dInicial,M.dFinal,M.nTipoMeta, C.cConsDescripcion as cTipoMeta, CN.cConsDescripcion as sMoneda, M.nMoneda "
    sSQL = sSQL & " From ColocMetasAnalista M Inner Join Constante C ON M.nTipoMeta = C.nConsValor AND C. nConsCod = " & gColocTipoMetas
    sSQL = sSQL & "     Inner Join Constante CN ON CN.nConsValor = M.nMoneda AND CN.nConsCod = " & gMoneda
    sSQL = sSQL & " Where M.cPersCod = '" & psPersCod & "'"
    sSQL = sSQL & "Group By M.cPersCod,M.dInicial,M.dFinal,M.nTipoMeta,C.cConsDescripcion,CN.cConsDescripcion,M.nMoneda "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosMetasAnalistaCab = oConecta.CargaRecordSet(sSQL)
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
Dim sSQL As String
Dim R As ADODB.Recordset
    sSQL = "Select SUM(C.nMontoCol) as nMontoCol from Colocaciones C "
    sSQL = sSQL & " Inner Join ColocacCred CC ON C.cCtaCod = CC.cCtaCod "
    sSQL = sSQL & " Inner Join ProductoPersona PP ON PP.cCtaCod = C.cCtaCod AND PP.nPrdPersRelac = 28"
    If pnTodos Then
        sSQL = sSQL & " Where SUBSTRING(C.cCtaCod,6,3) = '" & psProducto & "' "
    Else
        sSQL = sSQL & " Where CC.nColocCondicion = " & pnCondicion & " AND SUBSTRING(C.cCtaCod,6,3) = '" & psProducto & "' "
    End If
    sSQL = sSQL & " AND SUBSTRING(C.cCtaCod,9,1) = " & pnMoneda & " AND ( C.dVigencia >= '" & Format(pdFechaIni, "mm/dd/yyyy") & "' AND C.dVigencia <= '" & Format(pdFechaFin, "mm/dd/yyyy") & "') "
    sSQL = sSQL & " AND PP.cPersCod = '" & psPersCod & "'"
    
    Set oCon = New DConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    
    MontoColocadoCredCondicion = IIf(IsNull(R!nMontoCol), 0, R!nMontoCol)
    

End Function


Public Function RecuperaDatosMetasAnalistaDetalle(ByVal psPersCod As String, ByVal pnTipo As Integer, _
    ByVal pdInicio As Date, ByVal pdFinal As Date, ByVal pnMoneda As Integer) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaDatosMetasAnalistaDetalle
    
    'sSql = "Select nTipoCred,dInicial,dFinal,nMonto"
    'sSql = sSql & " From ColocMetasAnalista Where cPersCod = '" & psPersCod & "'"
    'sSql = sSql & " AND nTipoMeta = " & pnTipo & " AND  dInicial = '" & Format(pdInicio, "mm/dd/yyyy")
    'sSql = sSql & "' AND dFinal = '" & Format(pdFinal, "mm/dd/yyyy") & "' AND nMoneda = " & pnMoneda
    
    'sSql = "Select DISTINCT nTipoCred,dInicial,dFinal,nMonto"
    sSQL = "Select DISTINCT nTipoCred,nMonto"
    sSQL = sSQL & " From ColocMetasAnalista Where cPersCod = '" & psPersCod & "'"
    sSQL = sSQL & " AND nTipoMeta = " & pnTipo & " AND  dInicial = '" & Format(pdInicio, "mm/dd/yyyy") & "' "
    sSQL = sSQL & " AND nMoneda = " & pnMoneda & " AND dFinal = '" & Format(pdFinal, "mm/dd/yyyy") & "' "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosMetasAnalistaDetalle = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaDatosMetasAnalistaDetalle:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function


Public Function RecuperaCreditosParaAsignarGasto(ByVal pnRanIni As Double, ByVal pnRanFin As Double, _
    ByVal pnAplicado As ColocCalendApl, ByVal psTipoprod As String, ByVal psMoneda As String, _
    ByVal psCodGasto As String) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCreditosParaAsignarGasto
    sSQL = "Select distinct '',P.cCtaCod, RTRIM(Pers.cPersNombre), P.nSaldo"
    sSQL = sSQL & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "                 Inner Join Persona Pers ON PP.cPersCod = Pers.cPersCod"
    sSQL = sSQL & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod"
    sSQL = sSQL & "                 Inner Join ColocCalendario CL ON P.cCtaCod = CL.cCtaCod AND CL.nNroCalen=CC.nNroCalen and CL.nColocCalendApl = " & pnAplicado & " AND CL.nColocCalendEstado = " & gColocCalendEstadoPendiente
    sSQL = sSQL & "                 Inner Join ColocCalendDet CD ON P.cCtaCod=CD.cCtaCod AND CD.nNroCalen = CC.nNroCalen AND CD.nColocCalendApl = " & pnAplicado
    sSQL = sSQL & "  WHERE substring(P.cCtaCod,6,1) = '" & psTipoprod & "' AND P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ")"
    sSQL = sSQL & "  AND substring(P.cCtaCod,9,1)='" & psMoneda & "' And nPrdConceptoCod <> " & CLng(psCodGasto)
    sSQL = sSQL & " Group By P.cCtaCod, Pers.cPersNombre, P.nSaldo, CD.nCuota, CC.nNroCalen "
    sSQL = sSQL & " HAVING SUM(CD.nMonto-isnull(CD.nMontopagado,0)) >= " & Format(pnRanIni, "#0.00") & " AND SUM(CD.nMonto-isnull(CD.nMontopagado,0)) <= " & Format(pnRanFin, "#0.00")
    sSQL = sSQL & " and CD.nCuota not in (select nCuota from ColocCalenddet where nCuota = CD.nCuota and nNroCalen = CC.nNrocalen And cCtaCod = P.cCtaCod AND nPrdConceptoCod =" & psCodGasto & ")"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosParaAsignarGasto = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosParaAsignarGasto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaCuotasParaAsignarGasto(ByVal pnRanIni As Double, ByVal pnRanFin As Double, _
    ByVal pnAplicado As ColocCalendApl, ByVal psCtaCod As String, ByVal psCodGasto As String) As ADODB.Recordset

Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCuotasParaAsignarGasto
    sSQL = "Select distinct '',P.cCtaCod, Pers.cPersNombre, P.nSaldo, CD.nCuota, CC.nNroCalen, SUM(CD.nMonto-isnull(CD.nMontopagado,0)) as nMontoCuota "
    sSQL = sSQL & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "                 Inner Join Persona Pers ON PP.cPersCod = Pers.cPersCod"
    sSQL = sSQL & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod"
    sSQL = sSQL & "                 Inner Join ColocCalendario CL ON P.cCtaCod = CL.cCtaCod AND CL.nNroCalen=CC.nNroCalen and CL.nColocCalendApl = " & pnAplicado & " AND CL.nColocCalendEstado = " & gColocCalendEstadoPendiente
    sSQL = sSQL & "                 Inner Join ColocCalendDet CD ON P.cCtaCod=CD.cCtaCod AND CD.nNroCalen = CC.nNroCalen AND CD.nColocCalendApl = " & pnAplicado
    sSQL = sSQL & "  WHERE P.cCtaCod = '" & psCtaCod & "'"
    sSQL = sSQL & " Group By P.cCtaCod, Pers.cPersNombre, P.nSaldo, CD.nCuota, CC.nNroCalen "
    sSQL = sSQL & " HAVING SUM(CD.nMonto-isnull(CD.nMontopagado,0)) >= " & Format(pnRanIni, "#0.00") & " AND SUM(CD.nMonto-isnull(CD.nMontopagado,0)) <= " & Format(pnRanFin, "#0.00")
    sSQL = sSQL & " and CD.nCuota not in (select nCuota from ColocCalenddet where nCuota = CD.nCuota and nNroCalen = CC.nNrocalen And cCtaCod = P.cCtaCod AND nPrdConceptoCod =" & psCodGasto & ")"
        
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCuotasParaAsignarGasto = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCuotasParaAsignarGasto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function



Public Function RecuperaTiposCredito() As ADODB.Recordset
Dim oConecta As DConecta
Dim sSQL As String
    Set oConecta = New DConecta
    sSQL = "Select nConsCod, nConsValor, cConsDescripcion from Constante where nConsCod = " & gProducto & " and (nConsValor % 100) = 0  order by  cConsDescripcion"
    oConecta.AbreConexion
    Set RecuperaTiposCredito = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaCreditosParaJudicalTotal(Optional ByVal psPersCod As String = "", Optional psCtaCod As String = "") As ADODB.Recordset
Dim oParam As DParametro
Dim nValor As Double
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCreditosParaJudicalTotal
    Set oParam = New DParametro
    nValor = oParam.RecuperaValorParametro(gColocEstJudicial)
    Set oParam = Nothing
    
    If psPersCod <> "" Then
        sSQL = "Select P.cCtaCod, C.nMontoCol, P.nSaldo, P.nTasaInteres, P.nPrdEstado, CC.nNroCalen, "
        sSQL = sSQL & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
        sSQL = sSQL & " cTitular=(Select cPersNombre From Persona Where cPersCod = PP2.cPersCod), CC.nDiasAtraso "
        sSQL = sSQL & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
        sSQL = sSQL & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
        sSQL = sSQL & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersAnalista
        sSQL = sSQL & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersTitular
        sSQL = sSQL & " WHERE PP2.cPersCod = '" & psPersCod & "'"
        sSQL = sSQL & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & ")"
        sSQL = sSQL & " AND CC.nDiasAtraso > " & nValor
    Else
        If psCtaCod <> "" Then
            sSQL = "Select P.cCtaCod, C.nMontoCol, P.nSaldo, P.nTasaInteres, P.nPrdEstado, CC.nNroCalen, "
            sSQL = sSQL & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
            sSQL = sSQL & " cTitular=(Select cPersNombre From Persona Where cPersCod = PP2.cPersCod), CC.nDiasAtraso "
            sSQL = sSQL & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
            sSQL = sSQL & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
            sSQL = sSQL & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersAnalista
            sSQL = sSQL & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersTitular
            sSQL = sSQL & " WHERE P.cCtaCod = '" & psCtaCod & "'"
            sSQL = sSQL & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & ")"
        Else
            sSQL = "Select P.cCtaCod, C.nMontoCol, P.nSaldo, P.nTasaInteres, P.nPrdEstado, CC.nNroCalen, "
            sSQL = sSQL & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
            sSQL = sSQL & " cTitular=(Select cPersNombre From Persona Where cPersCod = PP2.cPersCod), CC.nDiasAtraso "
            sSQL = sSQL & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
            sSQL = sSQL & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
            sSQL = sSQL & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersAnalista
            sSQL = sSQL & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersTitular
            sSQL = sSQL & " WHERE CC.nDiasAtraso > " & nValor
            sSQL = sSQL & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & ")"
        End If
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosParaJudicalTotal = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaCreditosParaJudicalTotal:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaCreditosRefinanciados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCreditosRefinanciados
    sSQL = "Select Distinct CR.cCtaCodRef, "
    sSQL = sSQL & "nCapitalRef = (Select nMonto from ColocacRefinancDet Where cCtaCod = CR.cCtaCod AND cCtaCodRef = CR.cCtaCodRef AND nEstado = " & gColocEstadoRefinancAprobado & " AND nPrdConceptoCod = " & gColocConceptoCodCapital & "),"
    sSQL = sSQL & "nInteresRef = (Select SUM(nMonto) from ColocacRefinancDet Where cCtaCod = CR.cCtaCod AND cCtaCodRef = CR.cCtaCodRef AND nEstado = " & gColocEstadoRefinancAprobado & " AND nPrdConceptoCod <> " & gColocConceptoCodCapital & ")"
    sSQL = sSQL & " from ColocacRefinanc CR "
    sSQL = sSQL & " WHERE CR.cCtaCod = '" & psCtaCod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosRefinanciados = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaCreditosRefinanciados:
                  Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaColocacRefinanc(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaColocacRefinanc
    sSQL = "Select * from ColocacRefinanc where cCtacod = '" & psCtaCod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocacRefinanc = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaColocacRefinanc:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
Public Function RecuperaColocaciones(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaColocaciones
    sSQL = "Select * from Colocaciones where cCtaCod = '" & psCtaCod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocaciones = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaColocaciones:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaGastosExonerados(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    sSQL = "select * from ColocCredGastosExon Where cCtaCod = '" & psCtaCod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGastosExonerados = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaDatosComunes(ByVal psCtaCod As String, Optional pnTodos As Boolean = True, Optional ByVal MatEst As Variant = Nothing) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
Dim I As Integer

    On Error GoTo ErrorRecuperaDatosComunes
    sSQL = "Select P.nPrdEstado, CC.nCalendDinamico, CC.bCuotaCom, CL.nTasaIni as nTasaMora, CC.nCalendDinamTipo, CC.bPrepago, CC.bMiVivienda, P.cCtaCod, CE.nMonto nMontoSol, C.dVigencia,C.nMontoCol, CC.cMetLiquidacion, P.nSaldo, P.nTasaInteres, CC.nNroProxCuota, "
    sSQL = sSQL & " cAnalista=(Select cPersNombre From Persona Where cPersCod = PP.cPersCod),"
    sSQL = sSQL & " Pers.cPersNombre cTitular, Pers.cPersCod,"
    sSQL = sSQL & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = Pers.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), "
    sSQL = sSQL & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = Pers.cPersCod AND cPersIDTpo = '" & gPersIdRUC & "'), "
    sSQL = sSQL & " CC.nExoPenalidad, nNota = (select nColocNota  From ColocCalificacionAnalista Where cTipo='A' AND cCtaCod = C.cCtaCod And dColocNotaFecha = (select Max(dColocNotaFecha) from ColocCalificacionAnalista Where cCtaCod = C.cCtaCod AND cTipo='A')), "
    sSQL = sSQL & " CN2.cConsDescripcion as cTipoCredDescrip, "
    sSQL = sSQL & " CN.cConsDescripcion as cDestinoDescripcion, CC.nDiasAtrasoAcum, "
    sSQL = sSQL & " CN3.cConsDescripcion as cMoneda, C.nMontoCol, CN4.cConsDescripcion cEstado, CN5.cConsDescripcion cCondicion, "
    sSQL = sSQL & " cLineaCred = (Select cLineaCred from Colocaciones Where cCtaCod = C.cCtaCod), "
    sSQL = sSQL & " dFechaSol = (Select dPrdEstado from ColocacEstado where cCtaCod = P.cCtaCod And nPrdEstado = " & gColocEstSolic & "), "
    sSQL = sSQL & " dbo.ColocNotaCredito(P.cCtaCod) as nNotaSist "
    sSQL = sSQL & " From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod "
    sSQL = sSQL & "                 Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod "
    sSQL = sSQL & "                 Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersAnalista
    sSQL = sSQL & "                 Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "                 Inner Join Persona Pers ON PP2.cPersCod = Pers.cPersCod"
    sSQL = sSQL & "                 Left Join Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = " & gProducto
    sSQL = sSQL & "                 Left Join Constante CN ON CC.nColocDestino = CN.nConsValor AND CN.nConsCod = " & Trim(Str(gColocDestino))
    sSQL = sSQL & "                 Left Join Constante CN3 ON convert(int,substring(C.cCtaCod,9,1)) = CN3.nConsValor AND CN3.nConsCod = " & gMoneda
    sSQL = sSQL & "                 Left Join Constante CN4 ON P.nPrdEstado = CN4.nConsValor AND CN4.nConsCod = " & gColocEstado
    sSQL = sSQL & "                 Left Join Constante CN5 ON CN5.nConsValor = CC.nColocCondicion AND CN5.nConsCod = " & gColocCredCondicion
    sSQL = sSQL & "                 Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado =" & gColocEstSolic
    sSQL = sSQL & "                 Left Join ColocLineaCreditoTasa CL ON CL.cLineaCred = C.cLineaCred AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal
    sSQL = sSQL & " WHERE P.cCtaCod = '" & psCtaCod & "' " 'AND P.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
    If Not pnTodos Then
        sSQL = sSQL & " AND P.nPrdEstado in (" & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & ")"
    Else
        If IsArray(MatEst) Then
            sSQL = sSQL & " AND P.nPrdEstado in ("
            For I = 0 To UBound(MatEst)
                sSQL = sSQL & MatEst(I) & ","
            Next I
            sSQL = Mid(sSQL, 1, Len(sSQL) - 1) & ")"
        End If
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosComunes = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ErrorRecuperaDatosComunes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaCreditosVigentesRelacCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCreditosVigentesRelacCred
    sSQL = "Select  Pers.cPersNombre, CN.cConsDescripcion cRelacion, P.cCtaCod, CC.nDiasAtraso"
    sSQL = sSQL & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac not in (" & gColRelPersAnalista & "," & gColRelPersApoderado & "," & gColRelPersTitular & ")"
    sSQL = sSQL & "                 Inner Join Constante CN ON CN.nConsValor = PP.nPrdPersRelac AND CN.nConsCod = " & gColocRelacPers
    sSQL = sSQL & "                 Inner join Persona Pers ON Pers.cPersCod = PP.cPersCod "
    sSQL = sSQL & "                 Inner join ColocacCred CC ON CC.cCtaCod = P.cCtaCod "
    sSQL = sSQL & " Where Pers.cPersCod In (Select cPersCod from ProductoPersona Where cCtaCod = '" & psCtaCod & "')"
    sSQL = sSQL & "AND P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ") "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosVigentesRelacCred = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosVigentesRelacCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description


End Function

Public Function RecuperaEstadosdelCredito(ByVal psCtaCod As String, ByVal pbSolicitud As Boolean, _
    ByVal pbSugerencia As Boolean, ByVal pbAprobacion As Boolean) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaEstadosdelCredito
    sSQL = ""
    If pbSolicitud Then
        sSQL = sSQL & "Select CE.nPrdEstado, CN.cConsDescripcion cEstado, CE.nMonto, CE.nCuotas, CE.nPlazo, NULL as nMontoCuota "
        sSQL = sSQL & " From ColocacEstado CE Inner Join Constante CN ON CE.nPrdEstado = CN.nConsValor AND CN.nConsCod = " & gColocEstado
        sSQL = sSQL & " Where CE.cCtaCod = '" & psCtaCod & "' AND CE.nPrdEstado = " & gColocEstSolic
    End If
    If pbSugerencia Then
        If Len(sSQL) > 0 Then
            sSQL = sSQL & " Union "
        End If
        sSQL = sSQL & "Select CE.nPrdEstado,CN.cConsDescripcion cEstado, CE.nMonto, CE.nCuotas, CE.nPlazo, "
        sSQL = sSQL & " nMontoCuota = (Select SUM(CD.nMonto) from ColocCalendDet CD where nNrocalen=1 and cCtaCod = CE.cCtaCod AND nColocCalendApl = 1 AND nCuota = 1) "
        sSQL = sSQL & " From ColocacEstado CE Inner Join Constante CN ON CE.nPrdEstado = CN.nConsValor AND CN.nConsCod = " & gColocEstado
        sSQL = sSQL & " Where CE.cCtaCod = '" & psCtaCod & "' AND  CE.nPrdEstado = " & gColocEstSug
    End If
    If pbAprobacion Then
        If Len(sSQL) > 0 Then
            sSQL = sSQL & " Union "
        End If
        sSQL = sSQL & "Select CE.nPrdEstado, CN.cConsDescripcion cEstado, CE.nMonto, CE.nCuotas, CE.nPlazo, "
        sSQL = sSQL & " nMontoCuota = (Select SUM(CD.nMonto) from ColocCalendDet CD where nNrocalen=(Select nNroCalen From ColocacCred Where cCtaCod = CE.cCtaCod) and cCtaCod = CE.cCtaCod AND nColocCalendApl = 1 AND nCuota = 1) "
        sSQL = sSQL & " From ColocacEstado CE Inner Join Constante CN ON CE.nPrdEstado = CN.nConsValor AND CN.nConsCod = " & gColocEstado
        sSQL = sSQL & " Where CE.cCtaCod = '" & psCtaCod & "' AND  CE.nPrdEstado = " & gColocEstAprob
    End If
    sSQL = sSQL & " Order By Ce.nprdEstado"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaEstadosdelCredito = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaEstadosdelCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaColocacEstado(ByVal psCtaCod As String, ByVal pnEstado As ColocEstado, Optional pbTodos As Boolean = False) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaColocacEstado
    If Not pbTodos Then
        sSQL = "Select * from ColocacEstado Where cCtaCod = '" & psCtaCod & "' AND nPrdEstado =  " & pnEstado
    Else
        sSQL = "Select * from ColocacEstado Where cCtaCod = '" & psCtaCod & "' Order By nPrdEstado"
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocacEstado = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocacEstado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
    
End Function

Public Function RecuperaMontoGarantiaCredito(ByVal psCtaCod As String, ByVal pdFecha As Date) As Double
Dim sSQL As String
Dim oConecta As DConecta
Dim oGeneral As DGeneral
Dim nTipoCambioFijo As Double
Dim nMontoMN As Double
Dim nMontoME As Double
Dim R As ADODB.Recordset

    On Error GoTo ErrorRecuperaMontoGarantiaCredito
    Set oGeneral = New DGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
    nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))
    Set oGeneral = Nothing
    'Moneda Nacional
    sSQL = "Select sum(nGravado) as nTotal from ColocGarantia Where cCtaCod = '" & psCtaCod & "' And nMoneda = " & gMonedaNacional
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    If Not R.BOF And Not R.EOF Then
        nMontoMN = IIf(IsNull(R!nTotal), 0, R!nTotal)
        nMontoMN = CDbl(Format(nMontoMN, "#0.00"))
    End If
    R.Close
    Set R = Nothing
    
    'Moneda Extrangera
    sSQL = "Select sum(nGravado) as nTotal from ColocGarantia Where cCtaCod = '" & psCtaCod & "' And nMoneda = " & gMonedaExtranjera
    Set R = oConecta.CargaRecordSet(sSQL)
    If Not R.BOF And Not R.EOF Then
        nMontoME = IIf(IsNull(R!nTotal), 0, R!nTotal)
        nMontoME = CDbl(Format(nMontoME, "#0.00"))
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    If CInt(Mid(psCtaCod, 9, 1)) = gMonedaNacional Then
        RecuperaMontoGarantiaCredito = 0
        If nMontoMN > 0 Then
            RecuperaMontoGarantiaCredito = RecuperaMontoGarantiaCredito + nMontoMN
        End If
        If nMontoME > 0 Then
            RecuperaMontoGarantiaCredito = RecuperaMontoGarantiaCredito + (nMontoME * nTipoCambioFijo)
        End If
    End If
    
    If CInt(Mid(psCtaCod, 9, 1)) = gMonedaExtranjera Then
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

Public Function RecuperaDatosPrepago(ByVal psCtaCod As String, ByVal pdHoy As Date) As Double
Dim sSQL As String
Dim oConec As DConecta
Dim R As ADODB.Recordset

    sSQL = " Select SUM(MCD.nMonto) as nMontoPagado from MovCol MC Inner Join Mov M ON MC.nMovNro = M.nMovNro "
    sSQL = sSQL & " Inner Join MovColDet MCD ON MC.nMovNro = MCD.nMovNro "
    sSQL = sSQL & " Where MC.nPrepago = 1 AND MC.cCtaCod = '" & psCtaCod & "' AND "
    sSQL = sSQL & " M.cMovnro like '" & Year(pdHoy) & "%' "
    sSQL = sSQL & " AND MCD.nPrdConceptoCod = 1000 AND MCD.cOpeCod not in ('" & gCredDesembEfec & "','" & gCredDesembCtaNueva & "','" & gCredDesembCtaExist & "','" & gCredDesembCtaExistDOA & "','" & gCredDesembCtaNuevaDOA & "')"
    sSQL = sSQL & " AND MC.nNroCalen = (select nNroCalen From ColocacCred Where cCtaCod = MC.cCtaCod)"
    Set oConec = New DConecta
    oConec.AbreConexion
    Set R = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    RecuperaDatosPrepago = Format(IIf(IsNull(R!nMontoPagado), 0, R!nMontoPagado), "0.00")
    
End Function
Public Function RecuperaDatosCreditoVigente(ByVal psCtaCod As String, Optional ByVal pdHoy As Date = CDate("01/01/1900"), _
    Optional ByVal pbIncluirbAprobados As Boolean = False) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaDatosCreditoVigente
    sSQL = "Select CC.cMetLiquidacion, dbo.CuotaAprobada(P.cCtaCod) as CuotaAprobada, CC.bCuotaCom, CC.bMiVivienda, CC.nCalendDinamTipo, CC.bPrepago, CC.nCalPago, Pers.cPersNombre, C.cLineaCred,C.nMontoCol, P.nSaldo, "
    sSQL = sSQL & " cMoneda = (Select cConsDescripcion from Constante where nConsCod = " & gMoneda & " AND nConsValor = " & CInt(Mid(psCtaCod, 9, 1)) & "),"
    sSQL = sSQL & " CE.nCuotas nCuotasApr, CE.nColocCalendCod " & ", "
    sSQL = sSQL & " CC.nDiasAtraso, CC.cMetLiquidacion, P.nTransacc, CC.nCalendDinamico, CC.nIntPend, CC.nNroCalen, CC.nNroProxCuota, CC.nNroProxDesemb "
    If pdHoy <> CDate("01/01/1900") Then
        sSQL = sSQL & ", DATEDIFF(year,C.dVigencia,'" & Format(pdHoy, "mm/dd/yyyy") & "') as nPlazoTranscurrido "
    End If
    sSQL = sSQL & " From Producto P Inner join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "                 Inner Join Persona Pers ON PP.cPersCod = Pers.cPersCod "
    sSQL = sSQL & "                 Inner Join Colocaciones C ON C.cCtaCod =  P.cCtaCod "
    sSQL = sSQL & "                 Inner Join ColocacCred CC ON CC.cCtaCod =  P.cCtaCod "
    sSQL = sSQL & "                 Inner Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSQL = sSQL & " WHERE P.cCtaCod = '" & psCtaCod & "'"
    
    If pbIncluirbAprobados Then
        sSQL = sSQL & " AND P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstAprob & ")"
    Else
        sSQL = sSQL & " AND P.nPrdEstado in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ")"
    End If
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosCreditoVigente = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaDatosCreditoVigente:
    Err.Raise Err.Number, "Error En Proceso RecuperaDatosCreditoVigente", Err.Description
        
End Function

Public Sub ActualizaCreditoAnalista(ByVal psCtaCod As String, ByVal psPersCodAnalista As String)
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorActualizaCreditoAnalista
    Set oConecta = New DConecta
    sSQL = "UPDATE ProductoPersona SET cPersCod = '" & psPersCodAnalista & "' "
    sSQL = sSQL & " Where nPrdPersRelac = " & Trim(Str(gColRelPersAnalista)) & " AND cCtaCod = '" & psCtaCod & "'"
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSQL
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorActualizaCreditoAnalista:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso ActualizaCreditoAnalista", Err.Description
    
End Sub
Public Function CarteraInstitucion(ByVal psPersCod As String)
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorCarteraInstitucion
    sSQL = "Select C.cCtaCod, P.cPersNombre from ColocacConvenio C inner join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod and PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "             Inner Join Persona P ON PP.cPersCod = P.cPersCod"
    sSQL = sSQL & " where C.cPersCod = '" & Trim(psPersCod) & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set CarteraInstitucion = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorCarteraInstitucion:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function CarteraAnalista(ByVal psPersCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorCarteraAnalista
    
    Set oConecta = New DConecta
    sSQL = "Select C.cCtaCod, PER.cPersNombre "
    sSQL = sSQL & " from    ColocacCred C inner join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod "
    sSQL = sSQL & " inner join Producto PRD ON C.cCtaCod = PRD.cCtaCod "
    sSQL = sSQL & " inner join ProductoPersona PP2 ON C.cCtaCod = PP2.cCtaCod "
    sSQL = sSQL & " inner join Persona PER ON PP2.cPersCod = PER.cPersCod "
    sSQL = sSQL & " Where PP.nPrdPersRelac = 28 and PRD.nPrdEstado like '202%' and PP2.nPrdPersRelac = " & gColRelPersTitular & " and PP.cPersCod = '" & psPersCod & "'"
    oConecta.AbreConexion
    Set CarteraAnalista = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorCarteraAnalista:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    Set oConecta = Nothing
    
End Function
Public Sub EliminaSolicitud(ByVal psCtaCod As String)
Dim sSQL As String
Dim Conn As DConecta
Dim bTran As Boolean

    On Error GoTo ErrorEliminaSolicitud
    bTran = False
    Set Conn = New DConecta
    Conn.AbreConexion
    Conn.ConexionActiva.BeginTrans
    bTran = True
    
    'Elimina ProductoPersona
    sSQL = "Delete ProductoPersona Where cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSQL
    
    'Elimina ColocEstado
    sSQL = "Delete ColocacEstado Where cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSQL
    
    'Elimina ColocFteIngreso
    sSQL = "Delete ColocFteIngreso Where cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSQL
    
    'Elimina ColocacConvenio
    sSQL = "Delete ColocacConvenio Where cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSQL
    
    'Elimina ColocacCred
    sSQL = "Delete ColocacCred Where cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSQL
    
    'Elimina Colocaciones
    sSQL = "Delete Colocaciones Where cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSQL
    
    'Elimina Producto
    sSQL = "Delete Producto Where cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSQL
    
    Conn.ConexionActiva.CommitTrans
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Sub

ErrorEliminaSolicitud:
    If bTran Then
        Conn.RollBackTrans
        Set Conn = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso EliminaSolicitud", Err.Description
End Sub
Public Sub ActualizaSolicitud(ByVal oRelPersCred As UCredRelacion, ByVal psCtaCod As String, _
                          ByVal psCondCred As String, ByVal popersona As DPersona, ByVal psFuenteIng As String, ByVal pnMontoSol As Double, _
                          ByVal pnCuotasSol As Integer, ByVal pnPlazo As Integer, ByVal psDestCred As String, _
                          ByVal psPersCodAnalista As String, ByVal psPersCodInst As String, ByVal psCodModular As String, _
                          ByVal dFecSol As Date, ByVal psMovAct As String, ByVal pdFIFecEval As Date, ByVal bRefinanciar As Boolean, ByVal pnCondProd As Integer, Optional ByVal MatCredRef As Variant, Optional pnCapitalInt As Boolean = False)

Dim bTran As Boolean
Dim I As Integer
Dim sSQL As String
Dim Conn As DConecta
Dim oBase As DCredActualizaBD

    On Error GoTo ErrorActualizaSolicitud
    
    bTran = False
    Set Conn = New DConecta
    Conn.AbreConexion
    Conn.ConexionActiva.BeginTrans
    bTran = True
    
    'Actualiza Producto
    sSQL = "UPDATE Producto SET "
    sSQL = sSQL & " nPrdEstado = " & Trim(Str(gColocEstSolic)) & ", "
    sSQL = sSQL & " dPrdEstado = '" & Format(dFecSol, "mm/dd/yyyy") & "'"
    sSQL = sSQL & " WHERE cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSQL
    
    'Actualiza Nuevas Personas Relacionadas con el Credito
    oRelPersCred.IniciarMatriz
    
    Do While Not oRelPersCred.EOF
            sSQL = "DELETE ProductoPersona "
            sSQL = sSQL & " WHERE cCtaCod = '" & psCtaCod & "' AND cPersCod = '" & oRelPersCred.ObtenerCodigo & "' AND nPrdPersRelac = " & oRelPersCred.ObtenerValorRelacAnt
            Conn.ConexionActiva.Execute sSQL
            
            sSQL = "INSERT INTO ProductoPersona(cCtaCod,cPersCod,nPrdPersRelac) "
            sSQL = sSQL & " VALUES('" & psCtaCod & "','" & oRelPersCred.ObtenerCodigo & "'," & oRelPersCred.ObtenerValorRelac & ")"
            Conn.ConexionActiva.Execute sSQL
            
        oRelPersCred.siguiente
    Loop
    
    'Actualiza Analista
    sSQL = "UPDATE ProductoPersona SET "
    sSQL = sSQL & " cPersCod = '" & psPersCodAnalista & "' "
    sSQL = sSQL & " WHERE cCtaCod = '" & psCtaCod & "' AND nPrdPersRelac = " & gColRelPersAnalista
    
    Conn.ConexionActiva.Execute sSQL
    
    'Actualiza Registro de Colocaciones
    sSQL = "UPDATE Colocaciones SET "
    sSQL = sSQL & " nPlazo = " & Trim(Str(pnPlazo)) & ","
    sSQL = sSQL & " nMontoCol = " & Format(pnMontoSol, "#0.00") & ","
    sSQL = sSQL & " cUltimaActualizacion = '" & psMovAct & "' "
    sSQL = sSQL & " WHERE cCtaCod = '" & psCtaCod & "'"
    Conn.ConexionActiva.Execute sSQL
    
    'Actualiza Registro de ColocacCred pnCondProd
    sSQL = "UPDATE ColocacCred SET "
    sSQL = sSQL & " nColocCondicion = " & psCondCred & ", "
    sSQL = sSQL & " nColocCondicionProd = " & pnCondProd & ", "
    sSQL = sSQL & " nColocDestino = '" & psDestCred & "', "
    sSQL = sSQL & " bRefCapInt = " & IIf(pnCapitalInt, "1 ", "0 ")
    sSQL = sSQL & " WHERE cCtaCod = '" & psCtaCod & "'"
        
    Conn.ConexionActiva.Execute sSQL
    
    'Actualiza Registro en ColocEstado
    Set oBase = New DCredActualizaBD
    sSQL = "UPDATE ColocacEstado SET "
    sSQL = sSQL & " dPrdEstado = '" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
    sSQL = sSQL & " nCuotas = " & Trim(Str(pnCuotasSol)) & ","
    sSQL = sSQL & " nMonto = " & Format(pnMontoSol, "#0.00") & ","
    sSQL = sSQL & " nPlazo = " & Trim(Str(pnPlazo))
    sSQL = sSQL & " WHERE cCtaCod = '" & psCtaCod & "' AND nPrdEstado = " & Trim(Str(gColocEstSolic))
    Set oBase = Nothing
    Conn.ConexionActiva.Execute sSQL
    
    'Actualiza Registro en ColocFteIngreso
            
    sSQL = "UPDATE ColocFteIngreso SET "
    sSQL = sSQL & " cNumFuente = '" & psFuenteIng & "', "
    sSQL = sSQL & " dPersEval = '" & Format(pdFIFecEval, "mm/dd/yyyy") & "' "
    sSQL = sSQL & " WHERE cCtaCod = '" & psCtaCod & "' AND cNumFuente = '" & psFuenteIng & "'"
    Conn.ConexionActiva.Execute sSQL
    
    'Nuevo Registro de ColocConvenio
    If Len(Trim(psPersCodInst)) > 0 Then
        sSQL = "DELETE ColocacConvenio WHERE cCtaCod = '" & psCtaCod & "' AND cPersCod = '" & psPersCodInst & "'"
        Conn.ConexionActiva.Execute sSQL
        
        sSQL = "INSERT INTO ColocacConvenio(cCtaCod,cPersCod,cCodModular)"
        sSQL = sSQL & " VALUES('" & psCtaCod & "',"
        sSQL = sSQL & "'" & psPersCodInst & "',"
        sSQL = sSQL & "'" & psCodModular & "')"
        Conn.ConexionActiva.Execute sSQL
    End If
    
    'Si es una Refinanciancion
    'Si es una Refinanciacion Guardar los Datos
    If bRefinanciar Then
        If IsArray(MatCredRef) Then
            Set oBase = New DCredActualizaBD
            If UBound(MatCredRef) > 0 Then
                sSQL = "Delete ColocacRefinancDet where cCtaCod = '" & psCtaCod & "'"
                Conn.ConexionActiva.Execute sSQL
                sSQL = "Delete ColocacRefinanc where cCtaCod = '" & psCtaCod & "'"
                Conn.ConexionActiva.Execute sSQL
                For I = 0 To UBound(MatCredRef) - 1
                    sSQL = "INSERT INTO ColocacRefinanc(cCtaCod,cCtaCodRef,nEstado,dEstado,nMontoRef)"
                    sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                    sSQL = sSQL & MatCredRef(I, 0) & "',"
                    sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                    sSQL = sSQL & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                    sSQL = sSQL & MatCredRef(I, 1) & ")"
                    Conn.ConexionActiva.Execute sSQL
                    
                    'Ingreso del Detalle
                    'Capital
                    sSQL = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                    sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                    sSQL = sSQL & MatCredRef(I, 0) & "',"
                    sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                    sSQL = sSQL & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                    sSQL = sSQL & gColocConceptoCodCapital & ","
                    sSQL = sSQL & MatCredRef(I, 2) & ","
                    sSQL = sSQL & "0.00)"
                    Conn.ConexionActiva.Execute sSQL
                    'Interes Compensatorio
                    If CDbl(MatCredRef(I, 3)) > 0 Then
                        sSQL = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                        sSQL = sSQL & MatCredRef(I, 0) & "',"
                        sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                        sSQL = sSQL & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                        sSQL = sSQL & gColocConceptoCodInteresCompensatorio & ","
                        sSQL = sSQL & MatCredRef(I, 3) & ","
                        sSQL = sSQL & "0.00)"
                        Conn.ConexionActiva.Execute sSQL
                    End If
                    
                    'Interes Moratorio
                    If CDbl(MatCredRef(I, 4)) > 0 Then
                        sSQL = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                        sSQL = sSQL & MatCredRef(I, 0) & "',"
                        sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                        sSQL = sSQL & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                        sSQL = sSQL & gColocConceptoCodInteresMoratorio & ","
                        sSQL = sSQL & MatCredRef(I, 4) & ","
                        sSQL = sSQL & "0.00)"
                        Conn.ConexionActiva.Execute sSQL
                    End If
                    
                    'Interes Gracia
                    If CDbl(MatCredRef(I, 5)) > 0 Then
                        sSQL = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                        sSQL = sSQL & MatCredRef(I, 0) & "',"
                        sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                        sSQL = sSQL & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                        sSQL = sSQL & gColocConceptoCodInteresGracia & ","
                        sSQL = sSQL & MatCredRef(I, 5) & ","
                        sSQL = sSQL & "0.00)"
                        Conn.ConexionActiva.Execute sSQL
                    End If
                    
                    'Interes Suspenso
                    If CDbl(MatCredRef(I, 6)) > 0 Then
                        sSQL = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                        sSQL = sSQL & MatCredRef(I, 0) & "',"
                        sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                        sSQL = sSQL & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                        sSQL = sSQL & gColocConceptoCodInteresSuspenso & ","
                        sSQL = sSQL & MatCredRef(I, 6) & ","
                        sSQL = sSQL & "0.00)"
                        Conn.ConexionActiva.Execute sSQL
                    End If
                    
                    'Interes Reprogramado
                    If CDbl(MatCredRef(I, 7)) > 0 Then
                        sSQL = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                        sSQL = sSQL & MatCredRef(I, 0) & "',"
                        sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                        sSQL = sSQL & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                        sSQL = sSQL & gColocConceptoCodInteresReprogramado & ","
                        sSQL = sSQL & MatCredRef(I, 7) & ","
                        sSQL = sSQL & "0.00)"
                        Conn.ConexionActiva.Execute sSQL
                    End If
                    
                    'Gastos
                    If CDbl(MatCredRef(I, 8)) > 0 Then
                        sSQL = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                        sSQL = sSQL & MatCredRef(I, 0) & "',"
                        sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                        sSQL = sSQL & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
                        sSQL = sSQL & gColocConceptoCodGastoVarios & ","
                        sSQL = sSQL & MatCredRef(I, 8) & ","
                        sSQL = sSQL & "0.00)"
                        Conn.ConexionActiva.Execute sSQL
                    End If
                Next I
            End If
            Set oBase = Nothing
        End If
    End If
    
    Conn.ConexionActiva.CommitTrans
    Conn.CierraConexion
    Set Conn = Nothing
    
    Exit Sub

ErrorActualizaSolicitud:
    If bTran Then
        Conn.ConexionActiva.RollBackTrans
        Set Conn = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso ActualizaSolicitud", Err.Description

End Sub


Public Sub NuevaSolicitud(ByVal oRelPersCred As UCredRelacion, ByVal psCtaCod As String, _
                          ByVal psCondCred As String, ByVal popersona As DPersona, ByVal psFteIngreso As String, ByVal pnMontoSol As Double, _
                          ByVal pnCuotasSol As Integer, ByVal pnPlazo As Integer, ByVal psDestCred As String, _
                          ByVal psPersCodAnalista As String, ByVal psPersCodInst As String, ByVal psCodModular As String, _
                          ByVal dFecSol As Date, ByVal psMovAct As String, ByVal bRefinanciar As Boolean, _
                          ByVal pnCondProd As Integer, ByVal dFecEvalFte As Date, Optional ByVal MatCredRef As Variant, _
                          Optional pnCapitalInt As Boolean = False)
Dim pbTran As Boolean
Dim I As Integer
Dim sSQL As String
Dim Conn As DConecta
Dim oBase As DCredActualizaBD
Dim dFecGrab As Date
Dim dGar As DGarantia
Dim R As ADODB.Recordset

    On Error GoTo ErrorNuevaSolicitud
    
    pbTran = False
    Set Conn = New DConecta
    Conn.AbreConexion
    Conn.ConexionActiva.BeginTrans
    pbTran = True
    
    'Ingresa Nuevo Producto
    sSQL = "INSERT INTO Producto(cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc)"
    sSQL = sSQL & " VALUES('" & psCtaCod & "',"
    sSQL = sSQL & "0.00,"
    sSQL = sSQL & "0.00,"
    sSQL = sSQL & Trim(Str(gColocEstSolic)) & ","
    sSQL = sSQL & "'" & Format(dFecSol, "mm/dd/yyyy") & "',"
    sSQL = sSQL & "0)"
    Conn.ConexionActiva.Execute sSQL

    'Ingresa Nuevas Personas Relacionadas con el Credito
    oRelPersCred.IniciarMatriz
    Do While Not oRelPersCred.EOF
        sSQL = "INSERT INTO ProductoPersona(cCtaCod,cPersCod,nPrdPersRelac) "
        sSQL = sSQL & " VALUES('" & psCtaCod & "','" & oRelPersCred.ObtenerCodigo & "'," & oRelPersCred.ObtenerValorRelac & ")"
        Conn.ConexionActiva.Execute sSQL
        oRelPersCred.siguiente
    Loop

    'Ingresa el Analista
    sSQL = "INSERT INTO ProductoPersona(cCtaCod,cPersCod,nPrdPersRelac) "
    sSQL = sSQL & " VALUES('" & psCtaCod & "','" & psPersCodAnalista & "'," & gColRelPersAnalista & ")"
    Conn.ConexionActiva.Execute sSQL
    
    'Nuevo Registro de Colocaciones
    sSQL = "INSERT INTO Colocaciones(cCtaCod,nPlazo,dVenc,nMontoCol,nColocCalendCod,cUltimaActualizacion,cLineaCred)"
    sSQL = sSQL & " VALUES('" & psCtaCod & "',"
    sSQL = sSQL & Trim(Str(pnPlazo)) & ","
    sSQL = sSQL & "NULL" & ","
    sSQL = sSQL & Format(pnMontoSol, "#0.00") & ","
    sSQL = sSQL & "0,"
    sSQL = sSQL & "'" & psMovAct & "',"
    sSQL = sSQL & "'')"
    Conn.ConexionActiva.Execute sSQL
    
    'Nuevo Registro de ColocacCred
    sSQL = "INSERT INTO ColocacCred (cCtaCod,nDiasAtraso,nColocCondicion,nColocDestino,"
    sSQL = sSQL & "cProtesto,bCargoAuto,cMetLiquidacion,bRefCapInt,nNroProxCuota,nIntPend,"
    sSQL = sSQL & "nExoPenalidad, nColocCondicionProd)"
    sSQL = sSQL & "VALUES('" & psCtaCod & "',"
    sSQL = sSQL & "0,"
    sSQL = sSQL & psCondCred & ","
    sSQL = sSQL & psDestCred & ","
    sSQL = sSQL & "0,"
    sSQL = sSQL & "NULL,"
    sSQL = sSQL & "'GMIC',"
    sSQL = sSQL & IIf(pnCapitalInt, "1,", "0,")
    sSQL = sSQL & "1,"
    sSQL = sSQL & "0,"
    sSQL = sSQL & "NULL," & pnCondProd & ")"
    Conn.ConexionActiva.Execute sSQL
    
    'Nuevo Registro en ColocEstado
    Set oBase = New DCredActualizaBD
    sSQL = "INSERT INTO ColocacEstado(cCtaCod,dPrdEstado,nPrdEstado,nCuotas,nMonto,cDescripcion,"
    sSQL = sSQL & "nColocCalendCod,nPeriodoFechaFija,nPeriodoGracia,nPlazo,nTipoGracia)"
    sSQL = sSQL & " VALUES('" & psCtaCod & "',"
    sSQL = sSQL & "'" & Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
    sSQL = sSQL & Trim(Str(gColocEstSolic)) & ","
    sSQL = sSQL & Trim(Str(pnCuotasSol)) & ","
    sSQL = sSQL & Format(pnMontoSol, "#0.00") & ","
    sSQL = sSQL & "'Nueva Solicitud de Credito',"
    sSQL = sSQL & "NULL,"
    sSQL = sSQL & "NULL,"
    sSQL = sSQL & "NULL,"
    sSQL = sSQL & Trim(Str(pnPlazo)) & ","
    sSQL = sSQL & "NULL)"
    Set oBase = Nothing
    Conn.ConexionActiva.Execute sSQL
    
    'Nuevo Registro en ColocFteIngreso
    sSQL = "INSERT INTO ColocFteIngreso(cNumFuente,dPersEval,cCtaCod)"
            sSQL = sSQL & " VALUES('" & psFteIngreso & "',"
            sSQL = sSQL & "'" & Format(dFecEvalFte, "mm/dd/yyyy") & "', "
            sSQL = sSQL & "'" & psCtaCod & "')"
            Conn.ConexionActiva.Execute sSQL

    'Nuevo Registro de ColocConvenio
    If Len(Trim(psPersCodInst)) > 0 Then
        sSQL = "INSERT INTO ColocacConvenio(cCtaCod,cPersCod,cCodModular)"
        sSQL = sSQL & " VALUES('" & psCtaCod & "',"
        sSQL = sSQL & "'" & psPersCodInst & "',"
        sSQL = sSQL & "'" & psCodModular & "')"
        Conn.ConexionActiva.Execute sSQL
    End If
    
    'Si es una Refinanciacion Guardar los Datos
    If bRefinanciar Then
        If IsArray(MatCredRef) Then
            Set oBase = New DCredActualizaBD
            dFecGrab = CDate(Format(CDate(Format(dFecSol, "dd/mm/yyyy") & Space(1) & Format(oBase.dFechaHora, "hh:mm:ss")), "dd/mm/yyyy hh:mm:ss"))
            Set oBase = Nothing
            If UBound(MatCredRef) > 0 Then
                For I = 0 To UBound(MatCredRef) - 1
                    sSQL = "INSERT INTO ColocacRefinanc(cCtaCod,cCtaCodRef,nEstado,dEstado,nMontoRef)"
                    sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                    sSQL = sSQL & MatCredRef(I, 0) & "',"
                    sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                    sSQL = sSQL & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                    sSQL = sSQL & MatCredRef(I, 1) & ")"
                    Conn.ConexionActiva.Execute sSQL
                    
                    'Ingreso del Detalle
                    'Capital
                    sSQL = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                    sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                    sSQL = sSQL & MatCredRef(I, 0) & "',"
                    sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                    sSQL = sSQL & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                    sSQL = sSQL & gColocConceptoCodCapital & ","
                    sSQL = sSQL & MatCredRef(I, 2) & ","
                    sSQL = sSQL & "0.00)"
                    Conn.ConexionActiva.Execute sSQL
                    'Interes Compensatorio
                    If CDbl(MatCredRef(I, 3)) > 0 Then
                        sSQL = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                        sSQL = sSQL & MatCredRef(I, 0) & "',"
                        sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                        sSQL = sSQL & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                        sSQL = sSQL & gColocConceptoCodInteresCompensatorio & ","
                        sSQL = sSQL & MatCredRef(I, 3) & ","
                        sSQL = sSQL & "0.00)"
                        Conn.ConexionActiva.Execute sSQL
                    End If
                    
                    'Interes Moratorio
                    If CDbl(MatCredRef(I, 4)) > 0 Then
                        sSQL = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                        sSQL = sSQL & MatCredRef(I, 0) & "',"
                        sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                        sSQL = sSQL & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                        sSQL = sSQL & gColocConceptoCodInteresMoratorio & ","
                        sSQL = sSQL & MatCredRef(I, 4) & ","
                        sSQL = sSQL & "0.00)"
                        Conn.ConexionActiva.Execute sSQL
                    End If
                    
                    'Interes Gracia
                    If CDbl(MatCredRef(I, 5)) > 0 Then
                        sSQL = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                        sSQL = sSQL & MatCredRef(I, 0) & "',"
                        sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                        sSQL = sSQL & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                        sSQL = sSQL & gColocConceptoCodInteresGracia & ","
                        sSQL = sSQL & MatCredRef(I, 5) & ","
                        sSQL = sSQL & "0.00)"
                        Conn.ConexionActiva.Execute sSQL
                    End If
                    
                    'Interes Suspenso
                    If CDbl(MatCredRef(I, 6)) > 0 Then
                        sSQL = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                        sSQL = sSQL & MatCredRef(I, 0) & "',"
                        sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                        sSQL = sSQL & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                        sSQL = sSQL & gColocConceptoCodInteresSuspenso & ","
                        sSQL = sSQL & MatCredRef(I, 6) & ","
                        sSQL = sSQL & "0.00)"
                        Conn.ConexionActiva.Execute sSQL
                    End If
                    
                    'Interes Reprogramado
                    If CDbl(MatCredRef(I, 7)) > 0 Then
                        sSQL = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                        sSQL = sSQL & MatCredRef(I, 0) & "',"
                        sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                        sSQL = sSQL & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                        sSQL = sSQL & gColocConceptoCodInteresReprogramado & ","
                        sSQL = sSQL & MatCredRef(I, 7) & ","
                        sSQL = sSQL & "0.00)"
                        Conn.ConexionActiva.Execute sSQL
                    End If
                    
                    'Gastos
                    If CDbl(MatCredRef(I, 8)) > 0 Then
                        sSQL = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
                        sSQL = sSQL & " VALUES('" & psCtaCod & "','"
                        sSQL = sSQL & MatCredRef(I, 0) & "',"
                        sSQL = sSQL & gColocEstadoRefinancSolicitado & ","
                        sSQL = sSQL & "'" & Format(dFecGrab, "mm/dd/yyyy hh:mm:ss") & "',"
                        sSQL = sSQL & gColocConceptoCodGastoVarios & ","
                        sSQL = sSQL & MatCredRef(I, 8) & ","
                        sSQL = sSQL & "0.00)"
                        Conn.ConexionActiva.Execute sSQL
                    End If
                Next I
            End If
        End If
    End If
    
    'SI es Refinanciado Ingresar las Mismas Garantias de los Creditos Anteriores
    
    If bRefinanciar Then
        Set dGar = New DGarantia
        For I = 0 To UBound(MatCredRef) - 1
            Set R = dGar.RecuperaGarantiaCreditoDatos(MatCredRef(I, 0))
            Do While Not R.EOF
                sSQL = "INSERT INTO COLOCGARANTIA(cNumGarant, cCtaCod, nMoneda, nGravado, nEstado) "
                sSQL = sSQL & " VALUES('" & R!cNumgarant & "','" & psCtaCod & "'," & R!nMoneda & "," & Format(R!nGravado, "#0.00") & "," & IIf(IsNull(R!nEstado), 1, R!nEstado) & ") "
                Conn.ConexionActiva.Execute sSQL
                R.MoveNext
            Loop
        Next I
        Set dGar = Nothing
    End If
    
    
    Conn.ConexionActiva.CommitTrans
    Conn.CierraConexion
    Set Conn = Nothing
    
    Exit Sub

ErrorNuevaSolicitud:
    If pbTran Then
        Conn.ConexionActiva.RollBackTrans
        Set Conn = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso NuevaSolicitud", Err.Description
End Sub

Public Sub ActuializaPersRelacCred(ByVal poRelPersCred As UCredRelacion, ByVal psCtaCod As String, ByVal pbTransac As Boolean)
Dim sSQL As String
    
    On Error GoTo ErrorActuializaPersRelacCred
    If pbTransac Then
        Set Conn = New DConecta
        Conn.AbreConexion
        Conn.ConexionActiva.BeginTrans
    End If
    poRelPersCred.IniciarMatriz
    If Not poRelPersCred.EOF Then
        sSQL = "DELETE ProductoPersona Where cCtaCod = '" & psCtaCod & "' AND nPrdPersRelac not in (" & gColRelPersAnalista & "," & gColRelPersApoderado & ")"
        Conn.ConexionActiva.Execute sSQL
    End If
    Do While Not poRelPersCred.EOF
        sSQL = "INSERT INTO ProductoPersona(cCtaCod,cPersCod,nPrdPersRelac) "
        sSQL = sSQL & "       VALUES('" & psCtaCod & "','" & poRelPersCred.ObtenerCodigo & "'," & poRelPersCred.ObtenerValorRelac & ")"
        Conn.ConexionActiva.Execute sSQL
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
        Conn.ConexionActiva.RollBackTrans
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

Public Function RecuperaRelacPers(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaRelacPers
    
    sSQL = "Select P.dPersNacCreac, PN.cPersNatSexo, P.nPersPersoneria, Prd.nPrdEstado, PP.cPersCod, RTRIM(P.cPersNombre) cPersNombre, C.cConsDescripcion, C.nConsValor, DNI = (Select Top 1 cPersIDnro from PersID Where cPersCod = PP.cPersCod  AND cPersIDTpo = '" & gPersIdDNI & "'),"
    sSQL = sSQL & " RUC = (Select Top 1 cPersIDnro from PersID Where cPersCod = PP.cPersCod AND cPersIDTpo = '" & gPersIdRUC & "'), "
    sSQL = sSQL & " CASE WHEN PP.nPrdPersRelac = 20 THEN 1 WHEN PP.nPrdPersRelac = 22 THEN 2 WHEN PP.nPrdPersRelac = 25 THEN 3 ELSE 99 END nOrden   "
    sSQL = sSQL & " From ProductoPersona PP INNER JOIN Persona P ON PP.cPersCod = P.cPersCod "
    sSQL = sSQL & "                         LEFT JOIN PersonaNat PN ON PN.cPersCod = P.cPersCod "
    sSQL = sSQL & "                         INNER JOIN Producto Prd ON Prd.cCtaCod = PP.cCtaCod "
    sSQL = sSQL & "                         INNER JOIN Constante C ON PP.nPrdPersRelac = C.nConsValor "
    sSQL = sSQL & " WHERE C.nConsCod = " & gColocRelacPers & " And C.nConsValor not in (" & gColRelPersAnalista & "," & gColRelPersApoderado & "," & gColRelPersAcreedor & ") AND PP.cCtaCod = '" & psCtaCod & "' "
    sSQL = sSQL & " ORDER BY nOrden "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaRelacPers = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaRelacPers:
    Err.Raise Err.Number, "Error En Proceso de Recuperar Relaciones de Personas con el Credito", Err.Description
End Function

Public Function RecuperaProductoTasaInteres(ByVal psCtaCod As String, ByVal pnPrdTasaInteres As ColocTipoTasa) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaProductoTasaInteres
    sSQL = "Select * from ProductoTasaInteres Where cCtaCod = '" & psCtaCod & "' AND nPrdTasaInteres = " & pnPrdTasaInteres
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaProductoTasaInteres = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaProductoTasaInteres:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaProducto(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaProducto
    sSQL = "Select * from Producto where cCtacod = '" & psCtaCod & "' "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaProducto = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaProducto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaColocacCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaColocacCred
    sSQL = "Select * from ColocacCred where cCtacod = '" & psCtaCod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocacCred = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocacCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaSugerencia(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaSugerencia
    
    sSQL = "Select ISNULL(CE.cTipoGasto,'F') as cTipoGasto,CE.nPrdEstado, PRD.nTransacc, C.nColocCalendCod, CN2.cConsDescripcion as cTipoCredDescrip, C.nColocDestino, CN.cConsDescripcion as cDestinoDescripcion, P.cPersCod, P.cPersNombre, "
    sSQL = sSQL & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), "
    sSQL = sSQL & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdRUC & "'), "
    sSQL = sSQL & " CE.nProxMes,CE.nCalendDinamico, CE.nMonto, CE.nCuotas, CE.nPlazo, CE.dPrdEstado,  CE.nPeriodoFechaFija, CE.nPeriodoGracia, CE.nPlazo, CE.nTipoGracia, CE.nTipoDesembolso, "
    sSQL = sSQL & " cAnalista = (select Top 1 cPersNombre from persona where cPersCod = (Select cPersCod From ProductoPersona Where cCtaCod = C.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ")), "
    sSQL = sSQL & " nMontoSol = (Select nMonto from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "), "
    sSQL = sSQL & " nCuotasSol = (Select nCuotas from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
    sSQL = sSQL & " nPlazoSol = (Select nPlazo from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
    sSQL = sSQL & " nTasaComp = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & "),"
    sSQL = sSQL & " nTasaGracia = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntGracia & "),"
    sSQL = sSQL & " cLineaCred = (Select cLineaCred from Colocaciones Where cCtaCod = C.cCtaCod), "
    sSQL = sSQL & " C.bMiVivienda, C.bCuotaCom, C.nNroCalPar, ISNULL(PT.nTasaInteres,0) as nTasaMora, ISNULL(PT2.nTasaInteres,0) as nTasaICVen "
    sSQL = sSQL & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
    sSQL = sSQL & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado in (" & gColocEstSug & "," & gColocEstSolic & ")"
    sSQL = sSQL & "                    INNER JOIN Constante CN ON C.nColocDestino = CN.nConsValor AND CN.nConsCod = " & Trim(Str(gColocDestino))
    sSQL = sSQL & "                    INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = " & gProducto
    sSQL = sSQL & "                    INNER JOIN Producto PRD ON C.cCtaCod = PRD.cCtaCod"
    sSQL = sSQL & "                    LEFT JOIN ProductoTasaInteres PT ON C.cCtaCod = PT.cCtaCod AND PT.nPrdTasaInteres = 3 "
    sSQL = sSQL & "                    LEFT JOIN ProductoTasaInteres PT2 ON C.cCtaCod = PT2.cCtaCod AND PT2.nPrdTasaInteres = 6 "
    sSQL = sSQL & " WHERE C.cCtaCod = '" & psCtaCod & "' Order By CE.nPrdEstado DESC"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaSugerencia = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaSugerencia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
End Function

Public Function RecuperaDatosDesembolso(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaDatosDesembolso
    sSQL = "Select CC.nNroCalen, C.cLineaCred, L.cDescripcion, P.cPersCod, P.cPersNombre, P.cPersDireccDomicilio,CN.cConsDescripcion cMoneda,  "
    sSQL = sSQL & " CC.nNroproxDesemb, nPrestamo = (Select nMonto From ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstAprob & "),"
    sSQL = sSQL & " nMontoADesemb = (Select CD.nMonto from ColocCalenddet CD "
    sSQL = sSQL & "                 where CD.cCtaCod = C.cCtaCod and CD.nColocCalendApl=0 and nPrdConceptoCod=1000 and nCuota = CC.nNroproxDesemb "
    sSQL = sSQL & "                 and CD.nnrocalen=(Select MAX(nNroCalen) From ColocCalendario Where cCtaCod = CD.cCtaCod AND nColocCalendApl=0) ),  "
    sSQL = sSQL & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), "
    sSQL = sSQL & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdRUC & "'), "
    sSQL = sSQL & " nTotalDesemb=(Select Count(cCtaCod) from ColocCalendario Where cCtaCod = C.cCtaCod AND nNroCalen = CC.nNroCalen AND nColocCalendApl = " & gColocCalendAplDesembolso & "), "
    sSQL = sSQL & " CC.bMiVivienda, CC.bCuotaCom "
    sSQL = sSQL & " from Colocaciones C Inner Join ColocLineaCredito L ON C.cLineaCred = L.cLineaCred "
    sSQL = sSQL & "                       Inner Join ColocacCred CC ON C.cCtaCod = CC.cCtaCod"
    sSQL = sSQL & "                       Inner Join ProductoPersona PP ON PP.cCtaCod = C.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "                       Inner Join Persona P ON PP.cPersCod = P.cPersCod"
    sSQL = sSQL & "                       Inner Join Constante CN ON CN.nConsValor = Convert(int,substring(C.cCtaCod,9,1)) And CN.nConsCod = " & gMoneda
    sSQL = sSQL & " WHERE C.cCtaCod = '" & psCtaCod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosDesembolso = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosDesembolso:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaDatosAprobacion(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaDatosAprobacion
    
    sSQL = "Select CC.cLineaCred as cLineaCredCod,  CE.nPrdEstado, C.nColocCalendCod, CN2.cConsDescripcion as cTipoCredDescrip, C.nColocDestino, CN.cConsDescripcion as cDestinoDescripcion, "
    sSQL = sSQL & " ISNULL(CE.cTipogasto,'F') as cTipogasto, CE.nMonto, CE.nCuotas, CE.nPlazo, CE.nPeriodoGracia, CE.nPeriodoFechaFija, PRD.nTasaInteres, CE.nProxMes, CE.nCalendDinamico, CE.nTipoGracia, CE.nTipoDesembolso, "
    sSQL = sSQL & " cAnalista = (select Top 1 cPersNombre from persona where cPersCod = (Select cPersCod From ProductoPersona Where cCtaCod = C.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & ")), "
    sSQL = sSQL & " nMontoSol = (Select nMonto from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "), "
    sSQL = sSQL & " nCuotasSol = (Select nCuotas from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
    sSQL = sSQL & " nPlazoSol = (Select nPlazo from ColocacEstado Where cCtaCod = C.cCtaCod And nPrdEstado = " & gColocEstSolic & "),"
    sSQL = sSQL & " nTasaComp = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & "),"
    sSQL = sSQL & " nTasaGracia = (Select nTasaInteres from ProductoTasaInteres Where cCtaCod = C.cCtaCod And nPrdTasaInteres = " & gColocLineaCredTasasIntGracia & "),"
    sSQL = sSQL & " cLineaCred = (Select (cDescripcion +  space(150) + cLineaCred) from ColocLineaCredito Where cLineaCred = CC.cLineaCred), "
    sSQL = sSQL & " P2.cPersNombre cFteIngreso, PF.cPersFIPersCod, P.cPersCod, P.cPersNombre, P.cPersDireccDomicilio, "
    sSQL = sSQL & " CN3.cConsDescripcion as cCondicionCred, C.nColocCondicion, C.bRefCapInt, "
    sSQL = sSQL & " nMaxCalend = (Select MAX(nNroCalen) From ColocCalendario Cal Where cCtaCod = C.cCtaCod), "
    sSQL = sSQL & " C.bMiVivienda, C.bCuotaCom, "
    sSQL = sSQL & " nTasaIntComp = (Select nTasaInteres from ProductoTasaInteres where cCtaCod = C.cCtaCod AND nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & "), "
    sSQL = sSQL & " nTasaIntMora = (Select nTasaInteres from ProductoTasaInteres where cCtaCod = C.cCtaCod AND nPrdTasaInteres = " & gColocLineaCredTasasIntMoratNormal & ") "
    sSQL = sSQL & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "                    INNER JOIN Colocaciones CC ON CC.cCtaCod = C.cCtaCod "
    sSQL = sSQL & "                    INNER JOIN PersFteIngreso PF ON PF.cPersCod = PP.cPersCod "
    sSQL = sSQL & "                    INNER JOIN Persona P2 ON P2.cPersCod = PF.cPersFIPersCod "
    sSQL = sSQL & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
    sSQL = sSQL & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSug
    sSQL = sSQL & "                    INNER JOIN Constante CN ON C.nColocDestino = CN.nConsValor AND CN.nConsCod = " & Trim(Str(gColocDestino))
    sSQL = sSQL & "                    INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,6,3)) = CN2.nConsValor AND CN2.nConsCod = " & gProducto
    sSQL = sSQL & "                    INNER JOIN Constante CN3 ON C.nColocCondicion = CN3.nConsValor AND CN3.nConsCod = " & Trim(Str(gColocCredCondicion))
    sSQL = sSQL & "                    INNER JOIN Producto PRD ON C.cCtaCod = PRD.cCtaCod"
    sSQL = sSQL & " WHERE C.cCtaCod = '" & psCtaCod & "' And Prd.nPrdestado = " & gColocEstSug
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosAprobacion = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaDatosAprobacion:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
        
End Function

Public Function RecuperaSolicitud(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaSolicitud
    sSQL = "Select C.nColocDestino, C.nColocCondicion, C.nColocCondicionProd, P.cPersCod, P.cPersNombre, "
    sSQL = sSQL & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), "
    sSQL = sSQL & " RUC=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdRUC & "'), "
    sSQL = sSQL & " CE.nMonto, CE.nCuotas, CE.nPlazo, CE.dPrdEstado, "
    sSQL = sSQL & " cAnalista = (Select cPersCod From ProductoPersona Where cCtaCod = C.cCtaCod AND nPrdPersRelac = " & gColRelPersAnalista & "),"
    sSQL = sSQL & " CC.cPersCod as cPersConvenio, CC.cCodModular, CF.cNumFuente, CF.dPersEval dPersFIEvaluacion, C.bRefCapint "
    sSQL = sSQL & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "                    INNER JOIN Producto Prd ON Prd.cCtaCod = C.cCtaCod "
    sSQL = sSQL & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
    sSQL = sSQL & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic
    sSQL = sSQL & "                    LEFT JOIN ColocacConvenio CC ON C.cCtaCod = CC.cCtaCod"
    sSQL = sSQL & "                    LEFT join ColocFteIngreso CF ON C.cCtaCod = CF.cCtaCod "
    sSQL = sSQL & " WHERE C.cCtaCod = '" & psCtaCod & "' AND Prd.nPrdEstado = " & gColocEstSolic
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaSolicitud = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorRecuperaSolicitud:
    Err.Raise Err.Number, "Recuperar Solicitud", Err.Description
End Function

Public Function ExisteGasto(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
        pnColocCalendApl As Integer, pnCuota As Integer, ByVal nPrdConceptoCod As Long) As Boolean
Dim sSQL As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    
    sSQL = "Select nCuota From colocCalendDet Where cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & pnNroCalen
    sSQL = sSQL & " AND nColocCalendApl = " & pnColocCalendApl & " AND nCuota = " & pnCuota & " AND nPrdConceptoCod = " & nPrdConceptoCod
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
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

Public Function ExisteGarantia(ByVal psCtaCod As String, ByVal psNumGarant As String) As Boolean
Dim sSQL As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorExisteGarantia
    sSQL = "Select cNumGarant From ColocGarantia Where cCtaCod = '" & psCtaCod & "' AND cNumGarant = '" & psNumGarant & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
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
Public Function RecuperaDatosGarantiaCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim oConecta As DConecta
Dim sSQL As String

    On Error GoTo ErrorRecuperaDatosGarantiaCred
    sSQL = "Select CN3.cConsDescripcion as cTipoCredDescrip, C.nColocDestino, CN.cConsDescripcion as cDestinoDescripcion , P.cPersCod, P.cPersNombre, CN2.cConsDescripcion as cMonedaDesc, "
    sSQL = sSQL & " DNI=(Select Top 1 cPersIDNro from " & gConsPersona & "PersID Where cPersCod = P.cPersCod AND cPersIDTpo = '" & gPersIdDNI & "'), "
    sSQL = sSQL & " CE.nMonto "
    sSQL = sSQL & " From ColocacCred C INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "                    INNER JOIN " & gConsPersona & "Persona P ON PP.cPersCod = P.cPersCod"
    sSQL = sSQL & "                    INNER JOIN ColocacEstado CE ON C.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = '" & gColocEstSolic & "' "
    sSQL = sSQL & "                    INNER JOIN Constante CN ON C.nColocDestino = CN.nConsValor AND CN.nConsCod = " & gColocDestino
    sSQL = sSQL & "                    INNER JOIN Constante CN2 ON convert(int,substring(C.cCtaCod,9,1)) = CN2.nConsValor AND CN2.nConsCod = " & gMoneda
    sSQL = sSQL & "                    INNER JOIN Constante CN3 ON convert(int,substring(C.cCtaCod,6,3)) = CN3.nConsValor AND CN3.nConsCod = " & gProducto
    sSQL = sSQL & " WHERE C.cCtaCod = '" & psCtaCod & "' "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosGarantiaCred = oConecta.CargaRecordSet(sSQL)
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

Public Function EstadoActual(ByVal psCtaCod As String) As Integer
Dim R As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorEstadoActual
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSQL = "Select nPrdEstado from Producto Where cCtaCod = '" & psCtaCod & "'"
    Set R = oConecta.CargaRecordSet(sSQL)
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

Public Function GarantiaDeOtroCredito(ByVal psCtaCod As String, ByVal psNumGarant As String) As Boolean
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConec As DConecta

    sSQL = " Select CG.cNumGarant "
    sSQL = sSQL & " from ColocacRefinanc CR Inner Join ColocGarantia CG ON CR.cCtaCodRef = CG.cCtaCod "
    sSQL = sSQL & " where CR.cCtaCod = '" & psCtaCod & "' And CG.cNumGarant = '" & psNumGarant & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set R = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If R.RecordCount > 0 Then
        GarantiaDeOtroCredito = True
    Else
        GarantiaDeOtroCredito = False
    End If
    R.Close
    
End Function

Public Sub EliminarGarantia(ByVal psCtaCod As String, ByVal psNumGarant As String, _
     ByVal pnMontoGarantia As Double)
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorEliminarGarantia
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.BeginTrans
    sSQL = "DELETE ColocGarantia WHERE cCtaCod = '" & psCtaCod & "' AND cNumgarant = '" & psNumGarant & "' "
    oConecta.ConexionActiva.Execute sSQL
    
    'Si es garantia de otro credito por Refinanciacion
    If Not GarantiaDeOtroCredito(psCtaCod, psNumGarant) Then
        sSQL = "Update Garantias set nGravament = nGravament - " & Format(pnMontoGarantia, "#0.00") & " where cNumgarant = '" & psNumGarant & "' "
        oConecta.ConexionActiva.Execute sSQL
    End If
    oConecta.ConexionActiva.CommitTrans
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Sub

ErrorEliminarGarantia:
                  Err.Raise Err.Number, "Error En Proceso EliminarGarantia", Err.Description
    
    
End Sub

Public Function RecuperaColocGarantia(ByVal psCtaCod As String) As ADODB.Recordset
Dim oConecta As DConecta
Dim sSQL As String
        
    On Error GoTo ErrorRecuperaColocGarantia
    sSQL = "Select G.cTpoDoc, G.cNroDoc, G.nTpoGarantia, CG.cNumGarant, CG.cCtaCod, CG.nMoneda, CG.nGravado "
    sSQL = sSQL & " from ColocGarantia CG Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant "
    sSQL = sSQL & " Where cCtaCod = '" & psCtaCod & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocGarantia = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocGarantia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaMov(ByVal psCtaCod As String, ByVal psOpeCod As String) As Long
Dim sSQL As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    sSQL = "Select nMovNro from MovCol Where cCtaCod = '" & psCtaCod & "'"
    sSQL = sSQL & " And cOpeCod = '" & psOpeCod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
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

Public Function RecuperaGarantiasCredito(ByVal psCtaCod As String) As ADODB.Recordset
Dim oConecta As DConecta
Dim sSQL As String

    On Error GoTo ErrorRecuperaGarantiasCredito
    sSQL = "Select CG.cNumGarant, CG.nGravado, CG.nMoneda, G.nTpoGarantia, CN.cConsDescripcion cTpoGarDescripcion, "
    sSQL = sSQL & " G.nMoneda, G.nTasacion, G.nRealizacion, G.nPorGravar, G.cDescripcion, P.cPersNombre, G.cTpoDoc, G.cNroDoc "
    sSQL = sSQL & " From ColocGarantia CG Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant "
    sSQL = sSQL & "                       Inner Join Constante CN ON G.nTpoGarantia = CN.nConsvalor And CN.nConsCod = " & Trim(Str(gPersGarantia))
    sSQL = sSQL & "                       Inner Join PersGarantia PG ON CG.cNumGarant = PG.cNumGarant AND PG.nRelacion = " & Trim(Str(gPersRelGarantiaTitular))
    sSQL = sSQL & "                       Inner Join Persona P ON PG.cPersCod = P.cPersCod "
    sSQL = sSQL & " WHERE CG.cCtaCod = '" & psCtaCod & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantiasCredito = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaGarantiasCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function GarantiaPerteneceACreditoAprobado(ByVal psCtaCod As String, ByVal psNumagarant As String) As Boolean
Dim oConecta As DConecta
Dim sSQL As String
Dim R As ADODB.Recordset

    sSQL = "select P.nPrdestado from ColocGarantia CG Inner Join Producto P ON CG.cCtaCod = P.cCtaCod "
    sSQL = sSQL & " Where P.cCtaCod = '" & psCtaCod & "' AND CG.cNumgarant='" & psNumagarant & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
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

Public Sub NuevaGarantia(ByVal psCtaCod As String, ByVal psNumGarant As String, psMoneda As String, ByVal pnGravado As Double)
Dim oConecta As DConecta
Dim sSQL As String
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
    sSQL = "INSERT INTO ColocGarantia(cNumGarant, cCtaCod, nMoneda, nGravado, nEstado) "
    sSQL = sSQL & " VALUES('" & psNumGarant & "','" & psCtaCod & "'," & psMoneda & "," & Format(pnGravado, "#0.00") & ", 1)"
    Call oConecta.ConexionActiva.Execute(sSQL)
    
    If nGravado = 0 Then
        sSQL = "UPDATE Garantias SET "
        sSQL = sSQL & " nGravament = nGravament + " & Format(pnGravado, "#0.00") & ", nEstado = " & gPersGarantEstadoAsignado
        sSQL = sSQL & " WHERE cNumgarant = '" & psNumGarant & "'"
    Else
        sSQL = "UPDATE Garantias SET "
        sSQL = sSQL & " nGravament = nGravament + " & Format(pnGravado, "#0.00") & ", nEstado = " & gPersGarantEstadoContabilizado
        sSQL = sSQL & " WHERE cNumgarant = '" & psNumGarant & "'"
    End If
    Call oConecta.ConexionActiva.Execute(sSQL)
    
    Call oConecta.ConexionActiva.CommitTrans
    Call oConecta.CierraConexion
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorNuevaGarantia:
    If bTran Then
        Call oConecta.ConexionActiva.RollBackTrans
        Set oConecta = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso NuevaGarantia", Err.Description
End Sub


Public Sub ActualizaGarantias(ByVal psCtaCod As String, ByVal psNumGarant As String, psMoneda As String, ByVal pnGravado As Double, ByVal pnGravadoAnt As Double)
Dim oConecta As DConecta
Dim sSQL As String
Dim bTran As Boolean

    On Error GoTo ErrorNuevaGarantia
    bTran = False
    'Actualiza Registro en ColocGarantia
    sSQL = "UPDATE ColocGarantia SET "
    sSQL = sSQL & " nGravado = " & Format(pnGravado, "#0.00")
    sSQL = sSQL & " WHERE cNumGarant" & psNumGarant & "' AND cCtaCod = '" & psCtaCod & "' "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.BeginTrans
    bTran = True
    Call oConecta.ConexionActiva.Execute(sSQL)
    
    sSQL = "UPDATE Garantias SET "
    sSQL = sSQL & " nGravament = nGravament - " & Format(pnGravadoAnt, "#0.00")
    sSQL = sSQL & " WHERE cNumgarant = '" & psNumGarant & "'"
    Call oConecta.ConexionActiva.Execute(sSQL)
    sSQL = "UPDATE Garantias SET "
    sSQL = sSQL & " nGravament = nGravament + " & Format(pnGravado, "#0.00") & ", "
    sSQL = sSQL & " nEstado = " & gPersGarantEstadoAsignado
    sSQL = sSQL & " WHERE cNumgarant = '" & psNumGarant & "'"
    Call oConecta.ConexionActiva.Execute(sSQL)
    Call oConecta.ConexionActiva.CommitTrans
    Call oConecta.CierraConexion
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorNuevaGarantia:
    If bTran Then
        Call oConecta.ConexionActiva.RollBackTrans
        Set oConecta = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso NuevaGarantia", Err.Description
End Sub
Public Function RecuperaAgencia(ByVal psAge As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaAgencia
    sSQL = "Select * from Agencias Where cAgeCod = '" & psAge & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaAgencia = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaAgencia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
    
End Function

Public Function RecuperaCreditosVigentesCtas(ByVal psPersCod As String)
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaCreditosVigentes
    sSQL = "Select Prd.cCtaCod "
    sSQL = sSQL & " FROM Producto Prd Inner Join ProductoPersona PP ON Prd.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & " WHERE PP.cPersCod = '" & psPersCod & "' And Prd.nPrdEstado  in (" & gColocEstRefMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstVigMor & "," & gColocEstVigNorm & "," & gColocEstVigVenc & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosVigentesCtas = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosVigentes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
Public Function RecuperaCreditosVigentes(ByVal psPersCod As String, Optional ByVal psCtaCod As String = "", _
    Optional pEstadosEspecificos As Variant = Nothing, Optional ByVal psMoneda As String = "") As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
Dim I As Integer

    On Error GoTo ErrorRecuperaCreditosVigentes
    sSQL = "Select Prd.cCtaCod, CC.nDiasAtraso, Prd.nSaldo, CN.cConsDescripcion cEstado, Prd.nPrdEstado,"
    sSQL = sSQL & " nPrestamo = (Select nMonto From ColocacEstado Where cCtaCod = Prd.cCtaCod And nPrdEstado = " & gColocEstAprob & "), C.dVigencia, Prd.nTasaInteres"
    sSQL = sSQL & " FROM Producto Prd Inner Join ProductoPersona PP ON Prd.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "                   Inner Join Persona P ON PP.cPersCod = P.cPersCod "
    sSQL = sSQL & "                   Inner Join ColocacCred CC ON CC.cCtaCod = Prd.cCtaCod "
    sSQL = sSQL & "                   Inner Join Colocaciones C ON C.cCtaCod = Prd.cCtaCod "
    sSQL = sSQL & "                   Inner Join Constante CN ON CN.nConsValor = Prd.nPrdEstado And CN.nConsCod = " & gColocEstado
    sSQL = sSQL & " WHERE P.cPersCod = '" & psPersCod & "' "
    
    If psMoneda <> "" Then
        sSQL = sSQL & " AND SUBSTRING(Prd.cCtaCod,9,1) = '" & psMoneda & "' "
    End If
    If psCtaCod <> "" Then
        sSQL = sSQL & " AND Prd.cCtaCod <> '" & psCtaCod & "' "
    End If
    
    If IsArray(pEstadosEspecificos) Then
        If UBound(pEstadosEspecificos) <> 0 Then
            sSQL = sSQL & " AND Prd.nPrdEstado in ("
            For I = 0 To UBound(pEstadosEspecificos)
                sSQL = sSQL & pEstadosEspecificos(I) & ","
            Next I
            sSQL = Mid(sSQL, 1, Len(sSQL) - 1)
            sSQL = sSQL & ")"
        End If
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCreditosVigentes = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosVigentes:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaConsultaCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaConsultaCred
    sSQL = "Select ISNULL(CC.nCalendDinamico,0) as nCalendDinamico, P.nPrdEstado, ISNULL(CC.bMiVivienda,0) as bMiVivienda, ISNULL(CC.bCuotaCom,0) as bCuotaCom, CC.nDiasAtraso, PFI.cPersNombre cFteIngreso, L.cDescripcion cLineaDesc, PersAna.cPersNombre cAnalista, PersApod.cPersNombre cApoderado, CN.cConsDescripcion cCondicion, Lin.nTasaIni as nTasaGracia, "
    sSQL = sSQL & " CN2.cConsDescripcion cDestino, P.nTasaInteres, CN3.cConsDescripcion cTipoCuota, C.dVigencia,"
    sSQL = sSQL & " nNota = (select nColocNota  From ColocCalificacionAnalista Where cCtaCod = P.cCtaCod And dColocNotaFecha = (select Max(dColocNotaFecha) from ColocCalificacionAnalista Where cCtaCod = P.cCtaCod)), "
    sSQL = sSQL & " CE.nMonto nMontoSol, CE.dPrdEstado dFecSol, CE.nCuotas nCuotasSol, CE.nPlazo nPlazoSol,"
    sSQL = sSQL & " CE2.nMonto nMontoSug, CE2.dPrdEstado dFecSug, CE2.nCuotas nCuotasSug, CE2.nPlazo nPlazoSug, CE2.nPeriodoGracia nPeriodoGraciaSug,"
    sSQL = sSQL & " CE3.nMonto nMontoApr, CE3.dPrdEstado dFecApr, CE3.nCuotas nCuotasApr, CE3.nPlazo nPlazoApr, CE3.nPeriodoGracia nPeriodoGraciaApr, CN8.cConsDescripcion cTipoGracia,"
    sSQL = sSQL & " CN4.cConsDescripcion cMotivoRech, CE5.dPrdEstado dFecCancel, CE6.dPrdEstado dFecJud, CC.cMetLiquidacion, CC.cProtesto,"
    sSQL = sSQL & " CE7.nPrdEstado nEstRefin, CC.bCargoAuto, CN5.cConsDescripcion cEstActual, CN6.cConsDescripcion as cTipoCredDescrip, "
    sSQL = sSQL & " CN7.cConsDescripcion cTipoDesemb, CC.nNroProxDesemb, "
    sSQL = sSQL & " nCuotaSug = (Select sum(nMonto) from ColocCalendDet Cal Where Cal.cCtaCod = P.cCtaCod AND Cal.nNrocalen =1 AND Cal.nColocCalendApl=1 AND Cal.nCuota=1) " & " ,"
    sSQL = sSQL & " nCuotaApr = (Select sum(nMonto) from ColocCalendDet Cal Where Cal.cCtaCod = P.cCtaCod AND Cal.nNrocalen =(Select nNroCalen From ColocacCred Where cCtaCod = Cal.cCtaCod) AND Cal.nColocCalendApl=1 AND Cal.nCuota=2), "
    sSQL = sSQL & " nTasaComp = (Select nTasaInteres From ProductoTasaInteres Where nPrdTasaInteres  = 1 AND cCtaCod = P.cCtaCod), "
    sSQL = sSQL & " nTasaMor = (Select nTasaInteres From ProductoTasaInteres Where nPrdTasaInteres  = 3 AND cCtaCod = P.cCtaCod) "
    If CInt(Mid(psCtaCod, 6, 3)) = gColConsuDctoPlan Then
        sSQL = sSQL & ", Conv.cCodModular, PConv.cPersNombre cConvenio"
    End If
    sSQL = sSQL & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdpersRelac = " & gColRelPersTitular
    sSQL = sSQL & "                 Left Join ColocFteIngreso FI ON  FI.cCtaCod = P.cCtaCod"
    sSQL = sSQL & "                 Left Join PersFteIngreso PF ON PF.cNumFuente =  FI.cNumFuente "
    sSQL = sSQL & "                 Left Join Persona PFI ON PFI.cPersCod = PF.cPersFIPersCod"
    sSQL = sSQL & "                 Inner Join Colocaciones C ON C.cCtaCod = P.cCtaCod "
    sSQL = sSQL & "                 Left Join ColocLineaCredito L ON L.cLineaCred = C.cLineaCred"
    sSQL = sSQL & "                 Left Join ProductoPersona PP2 ON PP2.cCtaCod = P.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAnalista
    sSQL = sSQL & "                 Left Join Persona PersAna ON PersAna.cPersCod = PP2.cPersCod"
    sSQL = sSQL & "                 Left Join ProductoPersona PP3 ON PP3.cCtaCod = P.cCtaCod AND PP3.nPrdPersRelac = " & gColRelPersApoderado
    sSQL = sSQL & "                 Left Join Persona PersApod ON PersApod.cPersCod = PP3.cPersCod"
    sSQL = sSQL & "                 Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod"
    sSQL = sSQL & "                 Left Join Constante CN ON CN.nConsValor = CC.nColocCondicion AND CN.nConsCod = " & gColocCredCondicion
    sSQL = sSQL & "                 Left Join Constante CN2 ON CN2.nConsValor = CC.nColocDestino AND CN2.nConsCod = " & gColocDestino
    sSQL = sSQL & "                 Left Join Constante CN3 ON CN3.nConsValor = CC.nColocCalendCod AND CN3.nConsCod = " & gColocTipoCalend
    sSQL = sSQL & "                 Left Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic
    sSQL = sSQL & "                 Left Join ColocacEstado CE2 ON CE2.cCtaCod = P.cCtaCod AND CE2.nPrdEstado = " & gColocEstSug
    sSQL = sSQL & "                 Left Join ColocacEstado CE3 ON CE3.cCtaCod = P.cCtaCod AND CE3.nPrdEstado = " & gColocEstAprob
    sSQL = sSQL & "                 Left Join ColocacEstado CE4 ON CE4.cCtaCod = P.cCtaCod AND CE4.nPrdEstado = " & gColocEstRech
    sSQL = sSQL & "                 Left Join Constante CN4 ON CN4.nConsValor = CE4.nMotivoRechazo AND CN4.nConsCod = " & gColocMotivRechazo
    sSQL = sSQL & "                 Left Join ColocacEstado CE5 ON CE5.cCtaCod = P.cCtaCod AND CE5.nPrdEstado = " & gColocEstCancelado
    sSQL = sSQL & "                 Left Join ColocacEstado CE6 ON CE6.cCtaCod = P.cCtaCod AND CE6.nPrdEstado = " & gColocEstJudicial
    sSQL = sSQL & "                 Left Join ColocacEstado CE7 ON CE7.cCtaCod = P.cCtaCod AND CE7.nPrdEstado = " & gColocEstRefNorm
    sSQL = sSQL & "                 Left Join Constante CN5 ON CN5.nConsValor = P.nPrdEstado AND CN5.nConsCod = " & gColocEstado
    sSQL = sSQL & "                 Left Join Constante CN6 ON convert(int,substring(P.cCtaCod,6,3)) = CN6.nConsValor AND CN6.nConsCod = " & gProducto
    sSQL = sSQL & "                 Left Join Constante CN7 ON CN7.nConsValor = CC.nTipoDesembolso AND CN7.nConsCod = " & gColocTiposDesembolso
    sSQL = sSQL & "                 Left Join Constante CN8 ON CN8.nConsValor = CE3.nTipoGracia AND CN8.nConsCod = " & gColocTiposGracia
    sSQL = sSQL & "                 Left Join ColocLineaCreditoTasa Lin ON C.cLineaCred = Lin.cLineaCred AND Lin.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntGracia
    If CInt(Mid(psCtaCod, 6, 3)) = gColConsuDctoPlan Then
        sSQL = sSQL & "             Left Join ColocacConvenio Conv ON P.cCtaCod = Conv.cCtaCod"
        sSQL = sSQL & "             Left Join Persona PConv ON PConv.cPersCod = Conv.cPersCod "
    End If
    sSQL = sSQL & " WHERE P.cCtaCod = '" & psCtaCod & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaConsultaCred = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaConsultaCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaDacionesPago() As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    sSQL = " Select C.nNroGarantRec, C.cCtaCod, P.cPersNombre "
    sSQL = sSQL & " From ColocGarantRec C Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "             Inner Join Persona P ON P.cPersCod = PP.cPersCod "
    sSQL = sSQL & "    WHERE C.nEstado = " & gColocGarantRecEstadoRegistrado
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDacionesPago = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaDacionPago(ByVal pnNroDacion As Long) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    sSQL = "Select C.*, P.cPersNombre, P.cPersCod From ColocGarantRec C Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "     Inner Join Persona P ON PP.cPersCod = P.cPersCod "
    sSQL = sSQL & " Where nNroGarantRec = " & pnNroDacion
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDacionPago = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaDacionPagoDetalle(ByVal pnNroDacion As Long) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    sSQL = "Select * From ColocGarantRecdet Where nNroGarantRec = " & pnNroDacion
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDacionPagoDetalle = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaDatosDacionPagoCredito(ByVal pnDacion As Long) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    sSQL = "Select CG.cCtaCod, CG.nValorTotal, P.cPersNombre, CN.cConsDescripcion as cMoneda, C.nMontoCol,  "
    sSQL = sSQL & " Prd.nSaldo, CE.nCuotas, CC.nDiasAtraso, CC.cMetLiquidacion, Prd.nTransacc, CC.nCalendDinamico "
    sSQL = sSQL & " From ColocGarantRec CG Inner Join ProductoPersona PP ON CG.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular
    sSQL = sSQL & "             Inner Join Persona P ON P.cPersCod = PP.cPersCod "
    sSQL = sSQL & "             Inner Join Constante CN ON Convert(int,SUBSTRING(CG.cCtaCod,9,1)) = CN.nConsValor And CN.nConsCod = " & gMoneda
    sSQL = sSQL & "             Inner Join Colocaciones C ON CG.cCtaCod = C.cCtaCod "
    sSQL = sSQL & "             Inner Join Producto Prd ON Prd.cCtaCod = CG.cCtaCod "
    sSQL = sSQL & "             Inner Join ColocacEstado CE ON CE.cCtaCod = CG.cCtaCod And CE.nPrdEstado = " & gColocEstAprob
    sSQL = sSQL & "             Inner Join ColocacCred CC ON CC.cCtaCod = CG.cCtaCod "
    sSQL = sSQL & " Where CG.nNroGarantRec = " & pnDacion
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosDacionPagoCredito = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaProductosDeCredito(Optional ByVal psCab As String = "") As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    'sSQL = " Select cConsDescripcion, nConsValor from Constante"
    'sSQL = sSQL & " where nConscod = 1001 and (nConsValor % 10) <> 0 and (nConsValor < 232 or nConsValor >= 300)"
    'sSQL = sSQL & " Order by cConsDescripcion DESC"
    
    sSQL = " Select C.cConsDescripcion, C.nConsValor, CF.cAbrev from Constante C "
    sSQL = sSQL & " Inner Join ColocCredConsFiltro CF ON C.nConsCod = CF.nConsCod AND C.nConsValor = CF.nConsValor"
    sSQL = sSQL & " Where CF.nCodFiltro = " & gCredFiltroProd
    
    If psCab <> "" Then
        sSQL = sSQL & " AND C.nConsValor like '" & Mid(psCab, 1, 1) & "%'"
    End If
    sSQL = sSQL & " Order by C.nConsValor  "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaProductosDeCredito = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
Public Function RecuperaProductosDeSolicitudDeCredito() As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    
    sSQL = " Select C.cConsDescripcion, C.nConsValor from Constante C "
    sSQL = sSQL & " Inner Join ColocCredConsFiltro CF ON C.nConsCod = CF.nConsCod AND C.nConsValor = CF.nConsValor"
    sSQL = sSQL & " Where CF.nCodFiltro = " & gCredFiltroProdCab
    sSQL = sSQL & " Order by C.cConsDescripcion DESC"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaProductosDeSolicitudDeCredito = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaUltimaCuotaEvaluada(ByVal psCtaCod As String) As Integer
Dim sSQL As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset
Dim nCuota1 As Integer

    sSQL = "select ISNULL(MAX(nCuotaOrig),-1)+7 as nCuota from ColocCalifMiViv "
    sSQL = sSQL & "     where cCtaCod = '" & psCtaCod & "'"
    sSQL = sSQL & "   AND cColocMiVivEval='E'"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    nCuota1 = R!nCuota
    
    sSQL = "select ISNULL(MAX(nCuotaOrig),0) as nCuota from ColocCalifMiViv "
    sSQL = sSQL & "     where cCtaCod = '" & psCtaCod & "'"
    sSQL = sSQL & "   AND nColocCalendEstado=" & gColocCalendEstadoPagado
    Set R = oConecta.CargaRecordSet(sSQL)
    RecuperaUltimaCuotaEvaluada = nCuota1 - R!nCuota
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

Public Function RecuperaProductosDeSolicitudDeCF() As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    sSQL = " Select cConsDescripcion, nConsValor from Constante"
    sSQL = sSQL & " where nConscod = 1001 and (nConsValor % 10) <> 0 and (nConsValor < 232 or nConsValor >= 300) AND nConsValor In (" & gColCFComercial & "," & gColCFPYME & ")"
    sSQL = sSQL & " Order by cConsDescripcion DESC"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaProductosDeSolicitudDeCF = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

'EJVG20111219 *********************************************************************
Public Function ObtenerVecesCreditoyAhorroPersona(ByVal psPersCod As String) As Integer
    Dim sSQL As String
    Dim oConecta As DConecta
    Dim rs As ADODB.Recordset
    On Error GoTo ErrorObtenerVecesCreditoyAhorroPersona
    sSQL = " exec stp_sel_CuentaCreditosAhorroPersona '" & psPersCod & "'"
    
    Set oConecta = New DConecta
    Set rs = New ADODB.Recordset
    
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSQL)
    ObtenerVecesCreditoyAhorroPersona = rs!nVeces
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerVecesCreditoyAhorroPersona:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Private Sub Class_Initialize()
Dim CIni As ClsIni.ClasIni
    Set CIni = New ClsIni.ClasIni
    gConsPersona = CIni.BasePersonas
    gConsComunes = CIni.BaseComunes
    gConsImagenes = CIni.BaseImagenes
    Set CIni = Nothing
End Sub





