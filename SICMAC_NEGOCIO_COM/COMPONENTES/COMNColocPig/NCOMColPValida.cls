VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMColPValida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************
'* Clase que contiene las validaciones de las operaciones de Credito Pignoraticio
'* LAYG - 01/07/2001
'****************************
Option Explicit
Dim oCon As New COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String
Public Function nValidaAnulacionCredPignoraticio(ByVal psCtaCod As String, ByVal pdFecSis As Date, Optional ByVal cUser As String, Optional ByRef psmensaje As String, Optional pnTipoBusqueda As Integer = 0) As ADODB.Recordset
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones

Dim lrValida As ADODB.Recordset

Set lrValida = New ADODB.Recordset


'    lsSQL = " select m.nmovnro,m.cmovnro,mc.copecod, mcd.nprdconceptocod, montodet=mcd.nmonto,montot=mc.nmonto "
'    lsSQL = lsSQL & " into #tmpPrend" & cUser & " from mov m"
'    lsSQL = lsSQL & " join movcol mc on mc.nmovnro=m.nmovnro "
'    lsSQL = lsSQL & " join movcoldet mcd on mcd.nmovnro=mc.nmovnro and mcd.copecod=mc.copecod "
'    lsSQL = lsSQL & " where mc.cctacod='" & psCtaCod & "' and m.nmovflag=0 and m.nmovestado=10 "
'    oCon.Ejecutar (lsSQL)
'
'    lsSQL = " SELECT P.cCtaCod, P.nPrdestado,  " _
'          & " dRegistrado = (SELECT top 1 CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
'          & "              FROM #tmpPrend" & cUser & "" _
'          & "              WHERE cOpeCod = " & gColPOpeRegContrato & " )" _
'          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
'          & " Inner Join ColocPignoraticio CP on C.cctacod = CP.cctacod " _
'          & " WHERE P.cCtaCod ='" & psCtaCod & "' "
'
'    lrValida.CursorLocation = adUseClient
'    lrValida.Open lsSQL, oCon, adOpenStatic, adLockReadOnly, adCmdText
'    Set lrValida.ActiveConnection = Nothing
'
'
'    lsSQL = " drop table #tmpPrend" & cUser
'    oCon.Ejecutar lsSQL

    lsSQL = " SELECT P.cCtaCod, P.nPrdestado,  " _
          & " dRegistrado = (SELECT CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
          & "              FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
          & "              WHERE MC.cCtaCod = C.cCtaCod AND m.nmovflag=0"
    'If pnTipoBusqueda = 0 Then
        lsSQL = lsSQL & " and MC.cOpeCod in ( " & gColPOpeRegContrato & ")  )"
    'Else
      '  lsSQL = lsSQL & ")"
   'End If
          
 lsSQL = lsSQL & _
            " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " Inner Join ColocPignoraticio CP on C.cctacod = CP.cctacod " _
          & " WHERE P.cCtaCod ='" & psCtaCod & "' "

    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    
    If lrValida Is Nothing Or (lrValida.BOF And lrValida.EOF) Then
        psmensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
  
    If pnTipoBusqueda = 0 Then
        If lrValida!nPrdEstado <> gColPEstRegis Then
            psmensaje = "Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado)
            Exit Function
        End If
        If DateDiff("d", lrValida!dRegistrado, pdFecSis) <> 0 Then
            psmensaje = " Contrato se encuentra fuera de su Fecha Limite "
            Exit Function
        End If
   End If
        Set nValidaAnulacionCredPignoraticio = lrValida
        Set lrValida = Nothing
End Function

Public Function nValidaRescateCredPignoraticio(ByVal psCtaCod As String, ByVal psCodAge As String, _
            ByVal pdFecSis As Date, ByVal psOperacion As Long, Optional ByVal cUser As String, Optional ByRef psmensaje As String) As ADODB.Recordset

'************************************
' ************************************
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset
Dim lnCustodiaDiferida As Currency
Dim loCalculos As NCOMColPCalculos
Dim lsOpeCancel As String
  Set lrValida = New ADODB.Recordset
'lsSQL = " select m.nmovnro,m.cmovnro,mc.copecod, mcd.nprdconceptocod, montodet=mcd.nmonto,montot=mc.nmonto "
'lsSQL = lsSQL & " into #tmpPrend" & cUser & " from mov m"
'lsSQL = lsSQL & " join movcol mc on mc.nmovnro=m.nmovnro "
'lsSQL = lsSQL & " join movcoldet mcd on mcd.nmovnro=mc.nmovnro and mcd.copecod=mc.copecod "
'lsSQL = lsSQL & " where mc.cctacod='" & psCtaCod & "' and m.nmovflag=0 and m.nmovestado=10 "
'oCon.Ejecutar lsSQL
'
'lsOpeCancel = " ( '" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" & gColPOpeCancelNorCHQ & "','" & gColPOpeCancelMorCHQ & "','" & gColPOpeCanceNorEnOtCjEFE & "','" & gColPOpeCanceMorEnOtCjEFE & "') "
'
'lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, CP.cAgeBoveda , CP.nNroDuplic, " _
'      & " dCancelado = (SELECT top 1 CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
'      & "              FROM #tmpPrend" & cUser & "" _
'      & "              WHERE cOpeCod in " & lsOpeCancel & " ) ,  " _
'      & " nDiasTranscurridos = DateDiff ( d, (SELECT top 1 CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
'      & "              FROM #tmpPrend" & cUser & "" _
'      & "              WHERE cOpeCod in " & lsOpeCancel & " ), '" & Format(pdFecSis, "yyyy/mm/dd") & " ') , " _
'      & " nCustodiaPag = ISNULL( (SELECT SUM( montodet) FROM #tmpPrend" & cUser & "" _
'      & "                       WHERE cOpeCod =  " & geColPCobCusDiferida & "  ) , 0 ) , " _
'      & " cBloqueo = (SELECT CASE WHEN COUNT(nBlqMotivo) > 0 THEN 'S' ELSE 'N' END FROM ProductoBloqueos PB LEFT JOIN Constante Cons " _
'      & "             ON PB.nBlqMotivo = Cons.nConsCod WHERE PB.cCtaCod = '" & psCtaCod & "' AND cMovNroDbl Is Null ) , " _
'      & " nMaxDiasCustodia = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & gConsColPMaxDiasCustodiaDiferida & " ) , " _
'      & " nTasaIGV = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & gConsColPTasaIGV & " ) , " _
'      & " nPorcentajeCustodia = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & gConsColPPorcentajeCustodiaDiferida & " ) ," _
'      & " nCustodiaDiferida = 0 FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = c.cCtaCod " _
'      & " INNER JOIN ColocPignoraticio CP ON c.cCtaCod = cp.cCtaCod " _
'      & " WHERE P.cCtaCod = '" & psCtaCod & "' "
'
'     lrValida.CursorLocation = adUseClient
'    lrValida.Open lsSQL, oCon, adOpenStatic, adLockReadOnly, adCmdText
'    Set lrValida.ActiveConnection = Nothing
'
'
'
' lsSQL = " drop table #tmpPrend" & cUser
' oCon.Ejecutar lsSQL

'Comentado por DAOR 20070827
'lsOpeCancel = " ( '" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" & gColPOpeCancelNorCHQ & "','" & gColPOpeCancelMorCHQ & "','" & gColPOpeCanceNorEnOtCjEFE & "','" & gColPOpeCanceMorEnOtCjEFE & "') "
'
'lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, CP.cAgeBoveda , CP.nNroDuplic, " _
'      & " dCancelado = (SELECT CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
'      & "              FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
'      & "              WHERE MC.cCtaCod = C.cCtaCod AND M.NMOVFLAG=0 AND MC.cOpeCod in " & lsOpeCancel & " ) ,  " _
'      & " nDiasTranscurridos = DateDiff ( d, (SELECT CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
'      & "              FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
'      & "              WHERE MC.cCtaCod = C.cCtaCod AND MC.cOpeCod in " & lsOpeCancel & " AND M.nMovFlag <> " & gMovFlagExtornado & " ), '" & Format(pdFecSis, "yyyy/mm/dd") & " ') , " _
'      & " nCustodiaPag = ISNULL( (SELECT SUM(nMonto) FROM MovColDet MCD " _
'      & "                       INNER JOIN Mov M ON M.nMovNro = MCD.nMovNro " _
'      & "                       WHERE MCD.nMovNro = M.nMovNro AND M.NMOVFLAG=0 AND MCD.cOpeCod =  " & geColPCobCusDiferida & " AND MCD.cCtaCod = P.cCtaCod ) , 0 ) , " _
'      & " cBloqueo = (SELECT CASE WHEN COUNT(nBlqMotivo) > 0 THEN 'S' ELSE 'N' END FROM ProductoBloqueos PB LEFT JOIN Constante Cons " _
'      & "             ON PB.nBlqMotivo = Cons.nConsCod WHERE PB.cCtaCod = '" & psctacod & "' AND cMovNroDbl Is Null ) , " _
'      & " nMaxDiasCustodia = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & gConsColPMaxDiasCustodiaDiferida & " ) , " _
'      & " nTasaIGV = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & gConsColPTasaIGV & " ) , " _
'      & " nPorcentajeCustodia = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & gConsColPPorcentajeCustodiaDiferida & " ) ," _
'      & " nCustodiaDiferida = 0 FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = c.cCtaCod " _
'      & " INNER JOIN ColocPignoraticio CP ON c.cCtaCod = cp.cCtaCod " _
'      & " WHERE P.cCtaCod = '" & psctacod & "' "

'*** PEAC - 20170929 - COMENTADO PAR ENVIAR A UN STORE
'**Modificado por DAOR 20070827
'lsOpeCancel = " ( '" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" & gColPOpeCancelNorCHQ & "','" & gColPOpeCancelMorCHQ & "','" & gColPOpeCanceNorEnOtCjEFE & "','" & gColPOpeCanceMorEnOtCjEFE & "') "

'lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, CP.cAgeBoveda , CP.nNroDuplic, " _
      & " dCancelado = (SELECT top 1 CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
      & "              FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
      & "              WHERE MC.cCtaCod = C.cCtaCod AND M.NMOVFLAG=0 AND MC.cOpeCod in " & lsOpeCancel & " order by M.nMovNro desc ) ,  " _
      & " nDiasTranscurridos = DateDiff ( d, (SELECT top 1 CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
      & "              FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
      & "              WHERE MC.cCtaCod = C.cCtaCod AND MC.cOpeCod in " & lsOpeCancel & " AND M.nMovFlag <> " & gMovFlagExtornado & " order by M.nMovNro desc ), '" & Format(pdFecSis, "yyyy/mm/dd") & " ') , " _
      & " nCustodiaPag = ISNULL( (SELECT SUM(nMonto) FROM MovColDet MCD " _
      & "                       INNER JOIN Mov M ON M.nMovNro = MCD.nMovNro " _
      & "                       WHERE MCD.nMovNro = M.nMovNro AND M.NMOVFLAG=0 AND MCD.cOpeCod =  " & geColPCobCusDiferida & " AND MCD.cCtaCod = P.cCtaCod ) , 0 ) , " _
      & " cBloqueo = (SELECT CASE WHEN COUNT(nBlqMotivo) > 0 THEN 'S' ELSE 'N' END FROM ProductoBloqueos PB LEFT JOIN Constante Cons " _
      & "             ON PB.nBlqMotivo = Cons.nConsCod WHERE PB.cCtaCod = '" & psCtaCod & "' AND cMovNroDbl Is Null ) , " _
      & " nMaxDiasCustodia = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & gConsColPMaxDiasCustodiaDiferida & " ) , " _
      & " nTasaIGV = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & gConsColPTasaIGV & " ) , " _
      & " nPorcentajeCustodia = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & gConsColPPorcentajeCustodiaDiferida & " ) ," _
      & " nCustodiaDiferida = 0 FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = c.cCtaCod " _
      & " INNER JOIN ColocPignoraticio CP ON c.cCtaCod = cp.cCtaCod " _
      & " WHERE P.cCtaCod = '" & psCtaCod & "' "

lsSQL = "exec stp_sel_ValidaRescateCredPignoraticio '" & Format(pdFecSis, "yyyymmdd") & "','" & psCtaCod & "' "
'*** FIN PEAC

    Set loRegValida = New COMDColocPig.DCOMColPFunciones
        Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    Set loRegValida = Nothing
    If lrValida Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    If lrValida.BOF And lrValida.EOF Then
        psmensaje = "ERROR: al Buscar datos del credito "
        Exit Function
    End If
    
    If psOperacion = gColPOpeDevJoyas Then ' Para Rescate Joya
        If lrValida!nPrdEstado <> gColPEstDifer Then
            psmensaje = "Contrato se encuentra en estado " & mfgEstadoCredPigDesc(lrValida!nPrdEstado)
            Exit Function
        End If
    ElseIf psOperacion = gColPOpeDevJoyasNoDesemb Then  ' Para Rescate Joya No Desembolsada
        If lrValida!nPrdEstado <> gColPEstAnula Then
            psmensaje = "Contrato se encuentra en estado " & mfgEstadoCredPigDesc(lrValida!nPrdEstado)
            Exit Function
        End If
    End If
    ' Si esta bloqueada
    If lrValida!cBloqueo = "S" Then
        psmensaje = " Contrato se encuentra Bloqueado "
        Exit Function
    End If
    ' Ubicacion de la Joya
    If Trim(lrValida!cAgeBovedA) <> Mid(psCtaCod, 1, 3) & psCodAge Then
        psmensaje = " Garantía Prendaria se encuentra en Boveda de Agencia " & lrValida!cAgeBovedA
        Exit Function
    End If
    
    'Calcula el Costo de Custodia Extra
    If DateDiff("d", lrValida!dCancelado, pdFecSis) > lrValida!nMaxDiasCustodia Then
        Set loCalculos = New NCOMColPCalculos
            lnCustodiaDiferida = loCalculos.nCalculaCostoCustodiaDiferida(lrValida!nTasacion, lrValida!nDiasTranscurridos, lrValida!nPorcentajeCustodia, lrValida!nTasaIGV)
        Set loCalculos = Nothing
        If lnCustodiaDiferida - lrValida!nCustodiaPag > 0 Then
            psmensaje = "Contrato tiene Custodia Diferida, cobrar primero la custodia "
            Exit Function
        End If
    Else
        lnCustodiaDiferida = 0
    End If
    
    Set nValidaRescateCredPignoraticio = lrValida
End Function


Public Function nValidaDesembolsoCredPignoraticio(ByVal psCtaCod As String, ByVal pdFecSis As Date, _
     ByVal pnMaxDiasDesembolso As Integer, ByRef psmensaje As String) As ADODB.Recordset

'************************************

Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset

    'SE AGREGO (max(C.dVenc)dVenc) POR PEAC 20070813
    
    '*** PEAC 20090505
    'lsSQL = " SELECT P.cCtaCod, max(C.dVenc)dVenc, P.nPrdEstado, dVigencia, " _
         & " cBloqueo = (SELECT CASE WHEN COUNT(nBlqMotivo) > 0 THEN 'S' ELSE 'N' END FROM ProductoBloqueos PB LEFT JOIN Constante Cons " _
         & "             ON PB.nBlqMotivo = Cons.nConsCod WHERE PB.cCtaCod = '" & psctacod & "' AND cMovNroDbl Is Null ) , " _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodInteresCompensatorio & "   then nMonto else 0 end )  as nInteres," _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodImpuesto & "  then nMonto else 0 end )  as nImpuesto," _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodTasacion & "  then nMonto else 0 end )  as nTasacion," _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodCustodia & "  then nMonto else 0 end )  as nCustodia " _
         & " FROM Colocaciones c Inner join Producto P on P.cCtaCod = c.cCtaCod " _
         & " Inner join ColocCalendario CCal on CCal.cCtaCod = C.cCtaCod " _
         & " Inner join ColoccalendDetPig CCalD on CCal.cCtaCod = CCalD.cCtaCod and CCal.nNroCalen = CCalD.nNroCalen " _
         & " And CCal.nColocCAlendApl = CCalD.nColocCalendApl " _
         & " Where P.cCtaCod  = '" & psctacod & "' AND CCalD.nColocCalendApl =  " & gColocCalendAplDesembolso & " " _
         & " GROUP BY P.cCtaCod, P.nPrdEstado, dVigencia "
    
    lsSQL = " exec stp_sel_ValidaDesembolsoCredPignoraticio '" & psCtaCod & "'"
    
    'Set lrValida = New ADODB.Recordset
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    If lrValida.BOF And lrValida.EOF Then
        psmensaje = " No se ha encontrado Contrato "
        Exit Function
    End If
    If lrValida!nPrdEstado <> gColPEstRegis Then
        psmensaje = " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado)
        Exit Function
    ElseIf DateDiff("d", lrValida!dVigencia, pdFecSis) > pnMaxDiasDesembolso Then
        psmensaje = " Contrato está fuera de la Fecha para su Desembolso "
        Exit Function
    End If
    If lrValida!cBloqueo = "S" Then
        psmensaje = " Contrato se encuentra Bloqueado "
        Exit Function
    End If
    Set nValidaDesembolsoCredPignoraticio = lrValida
    Set lrValida = Nothing
End Function


Public Function nValidaRenovacionCredPignoraticio(ByVal psCtaCod As String, ByVal pdFecSis As Date, _
     ByVal pnMaxDiasDesembolso As Integer, Optional ByVal cUser As String, Optional ByRef psmensaje As String, Optional ByRef pnCredRevolAntNue As Boolean = False) As ADODB.Recordset

'*** PEAC - se agregó parámetro pnCredRevolAntNue

'************************************
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset
Set lrValida = New ADODB.Recordset
    
'    lsSQL = " select m.nmovnro,m.cmovnro,mc.copecod, mcd.nprdconceptocod, montodet=mcd.nmonto,montot=mc.nmonto "
'    lsSQL = lsSQL & " into #tmpPrend" & cUser & " from mov m"
'    lsSQL = lsSQL & " join movcol mc on mc.nmovnro=m.nmovnro "
'    lsSQL = lsSQL & " join movcoldet mcd on mcd.nmovnro=mc.nmovnro and mcd.copecod=mc.copecod "
'    lsSQL = lsSQL & " where mc.cctacod='" & psCtaCod & "' and m.nmovflag=0 and m.nmovestado=10 "
'    oCon.Ejecutar lsSQL
'
'    lsSQL = " SELECT codigoant=isnull(r.cctacodant,''),P.cCtaCod, P.nPrdestado, C.dVenc, C.nPlazo, CP.nNroRenov,  " _
'          & " P.nSaldo, CP.nTasacion, P.nTasaInteres, " _
'          & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
'          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ),  " _
'          & " dDesembolso = (SELECT MAX (CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
'          & "              FROM #tmpPrend" & cUser & " " _
'          & "              WHERE  cOpeCod= " & gColPOpeDesembolsoEFE & " ),  " _
'          & " dRenovacion = (SELECT MAX (CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
'          & "              FROM #tmpPrend" & cUser & " " _
'          & "              WHERE cOpeCod  in ( " & gColPOpeRenovNorEFE & "," & gColPOpeRenovMorEFE & "," _
'          & gColPOpeRenovNorEnOtCjEFE & "," & gColPOpeRenovMorEnOtCjEFE & ")  ),  " _
'          & " cInicRemate = (SELECT CASE WHEN COUNT(RG.nRGEstado)> 0  THEN 'S' ELSE 'N' END " _
'          & "                FROM ColocPigRecGar RG Join ColocPigRGDet RGD ON RG.cTpoProceso = RGD.cTpoProceso And RG.cNroProceso = RGD.cNroProceso  " _
'          & "                WHERE RG.cTpoProceso ='R' AND DATEDIFF(D, RG.dProceso,'" & Format(pdFecSis, "mm/dd/yyyy") & "') = 0 " _
'          & "                AND RG.nRGEStado ='" & gColPRecGarEstIniciado & "' AND RGD.cCtaCod = C.cCtaCod AND RG.cCodAge = Substring(C.cCtaCod,1,5)  )" _
'          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
'          & " Inner Join ColocPignoraticio CP on C.cctacod = CP.cctacod " _
'          & " Left join Relcuentas r on r.cctacod=p.cctacod  " _
'          & " WHERE P.cCtaCod ='" & psCtaCod & "'"
'
'    lrValida.CursorLocation = adUseClient
'    lrValida.Open lsSQL, oCon, adOpenStatic, adLockReadOnly, adCmdText
'    Set lrValida.ActiveConnection = Nothing
'
'    lsSQL = " drop table #tmpPrend" & cUser
'    oCon.Ejecutar lsSQL
          
    '**Comentado por DAOR 20070714
'    lsSQL = lsSQL & " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, C.nPlazo, CP.nNroRenov,  " _
'          & " P.nSaldo, CP.nTasacion, P.nTasaInteres," _
'          & " nTasaIntMora = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
'          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ),  " _
'          & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
'          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntCompVencida & " ),  " _
'          & " dDesembolso = (SELECT MAX (CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
'          & "              FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
'          & "              WHERE MC.cCtaCod = C.cCtaCod AND MC.cOpeCod  in ( " & gColPOpeDesembolsoEFE & ") AND M.nMovFlag <> " & gMovFlagExtornado & " ),  " _
'          & " dRenovacion = (SELECT MAX (CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
'          & "              FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
'          & "              WHERE MC.cCtaCod = C.cCtaCod AND MC.cOpeCod  in ( " & gColPOpeRenovNorEFE & "," & gColPOpeRenovMorEFE & "," _
'          & gColPOpeRenovNorEnOtCjEFE & "," & gColPOpeRenovMorEnOtCjEFE & ") AND M.nMovFlag <> " & gMovFlagExtornado & " ),  " _
'          & " cInicRemate = (SELECT CASE WHEN COUNT(RG.nRGEstado)> 0  THEN 'S' ELSE 'N' END " _
'          & "                FROM ColocPigRecGar RG Join ColocPigRGDet RGD ON RG.cTpoProceso = RGD.cTpoProceso And RG.cNroProceso = RGD.cNroProceso  " _
'          & "                WHERE RG.cTpoProceso ='R' AND DATEDIFF(D, RG.dProceso,'" & Format(pdFecSis, "mm/dd/yyyy") & "') = 0 " _
'          & "                AND RG.nRGEStado ='" & gColPRecGarEstIniciado & "' AND RGD.cCtaCod = C.cCtaCod AND RG.cCodAge = Substring(C.cCtaCod,1,5)  ) " _
'          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
'          & " Inner Join ColocPignoraticio CP on C.cctacod = CP.cctacod " _
'          & " WHERE P.cCtaCod ='" & psctacod & "' "
          
    '**DAOR 20070714***************************************
    
    ' se agrego P.dPrdestado,C.nMontoCol y cp.cCredB peac 20070820
'    lsSQL = lsSQL & " SELECT P.cCtaCod, P.nPrdestado, P.dPrdestado,C.nMontoCol, C.dVenc, C.nPlazo, CP.nNroRenov,isnull(cp.cCredB,'')cCredB,  " _
'          & " P.nSaldo, CP.nTasacion, P.nTasaInteres," _
'          & " nTasaIntMora = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
'          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ),  " _
'          & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
'          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntCompVencida & " ),  " _
'          & " dDesembolso = (SELECT MAX (CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
'          & "              FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
'          & "              WHERE MC.cCtaCod = C.cCtaCod AND MC.cOpeCod  in ( " & gColPOpeDesembolsoEFE & ") AND M.nMovFlag <> " & gMovFlagExtornado & " ),  " _
'          & " dRenovacion = (SELECT MAX (CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
'          & "              FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
'          & "              WHERE MC.cCtaCod = C.cCtaCod AND MC.cOpeCod  in ( " & gColPOpeRenovNorEFE & "," & gColPOpeRenovMorEFE & "," _
'          & gColPOpeRenovNorEnOtCjEFE & "," & gColPOpeRenovMorEnOtCjEFE & ") AND M.nMovFlag <> " & gMovFlagExtornado & " ),  " _
'          & " cInicRemate = (SELECT CASE WHEN COUNT(RG.nRGEstado)> 0  THEN 'S' ELSE 'N' END " _
'          & "                FROM ColocPigRecGar RG Join ColocPigRGDet RGD ON RG.cTpoProceso = RGD.cTpoProceso And RG.cNroProceso = RGD.cNroProceso  " _
'          & "                WHERE RG.cTpoProceso ='R' AND DATEDIFF(D, RG.dProceso,'" & Format(pdFecSis, "mm/dd/yyyy") & "') = 0 " _
'          & "                AND RG.nRGEStado ='" & gColPRecGarEstIniciado & "' AND RGD.cCtaCod = C.cCtaCod AND RG.cCodAge = Substring(C.cCtaCod,1,5)  ), " _
'          & " nEstUltProcRem=isnull((select top 1 nRGDetEstado from ColocPigRGDET " _
'          & "                       where cCodAge=Substring(C.cCtaCod,1,5) and cTpoProceso='R' and cCtaCod=C.cCtaCod " _
'          & "                       order by cNroProceso desc),0) " _
'          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
'          & " Inner Join ColocPignoraticio CP on C.cctacod = CP.cctacod " _
'          & " WHERE P.cCtaCod ='" & psctacod & "' "
    '******************************************************
    
    '*** PEAC 20080208
    lsSQL = " exec stp_sel_ObtieneDatosParaRenovacionPigno '" & Format(pdFecSis, "yyyymmdd") & "','" & _
    psCtaCod & "'"
    
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)

    If lrValida Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    If lrValida.BOF And lrValida.EOF Then
        psmensaje = "ERROR: No se encontro el Credito "
        Exit Function
    End If
    If (lrValida!nPrdEstado <> gColPEstDesem) And (lrValida!nPrdEstado <> gColPEstVenci) And _
       (lrValida!nPrdEstado <> gColPEstPRema) And (lrValida!nPrdEstado <> gColPEstRenov) Then
        psmensaje = " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado)
        Exit Function
    End If
    ' Si ha desembolsado el mismo dia
    If DateDiff("d", pdFecSis, lrValida!dDesembolso) = 0 Then
        psmensaje = "Contrato ha sido Desembolsado el dia de hoy día "
            Exit Function
    End If
    ' Si ha renovado el mismo dia
    If Not IsNull(lrValida!dRenovacion) And DateDiff("d", pdFecSis, lrValida!dRenovacion) = 0 Then
         psmensaje = "Contrato ha sido Renovado el dia de hoy"
         Exit Function
    End If

    ' Si No se ha definido la tasa de Interes Moratoria
    If IsNull(lrValida!nTasaIntVenc) Then
        psmensaje = " No se ha definido la Tasa de Interes Vencida para linea de Credito "
        Exit Function
    End If
        
    ' Si el Proceso Remate se ha iniciado ' LAYG (21/07/2003)
    If lrValida!cInicRemate = "S" Then
        psmensaje = " Contrato se encuentra EN EL PROCESO DE REMATE, no se puede Renovar "
        Exit Function
    End If

    '*** PEAC 20161108
    If pnCredRevolAntNue = False Then 'true=antiguo, false=nuevo
        If lrValida!nVigenciaCredPigno = 1 Then '*** 1 vencido / 0 no vence
            psmensaje = "El crédito debe Cancelarse ya que superó el límite de " & Trim(Str(lrValida!nVigenciaParam)) & " meses de vigencia."
            Exit Function
        ElseIf lrValida!nVigenciaParaAlerta = 1 Then
            psmensaje = lrValida!cMensajeParaAlerta
        End If
    End If
    
    Set nValidaRenovacionCredPignoraticio = lrValida
    Set lrValida = Nothing
    
End Function


Public Function nValidaCancelacionCredPignoraticio(ByVal psCtaCod As String, ByVal pdFecSis As Date, _
     ByVal pnMaxDiasDesembolso As Integer, ByRef psmensaje As String) As ADODB.Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset


    'Comentado por DAOR 20070716
'    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, C.nPlazo, CP.nNroRenov,  " _
'          & " P.nSaldo, CP.nTasacion, P.nTasaInteres, " _
'          & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
'          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntCompVencida & " ) ,  " _
'          & " nTasaIntMora = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
'          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ) ,  " _
'          & " cInicRemate = (SELECT CASE WHEN COUNT(RG.nRGEstado)> 0  THEN 'S' ELSE 'N' END " _
'          & "                FROM ColocPigRecGar RG Join ColocPigRGDet RGD ON RG.cTpoProceso = RGD.cTpoProceso And RG.cNroProceso = RGD.cNroProceso " _
'          & "                WHERE RG.cTpoProceso ='R' AND DATEDIFF(D, RG.dProceso,'" & Format(pdFecSis, "mm/dd/yyyy") & "') = 0  " _
'          & "                AND RG.nRGEStado ='" & gColPRecGarEstIniciado & "' AND RGD.cCtaCod = C.cCtaCod AND RG.cCodAge = Substring(C.cCtaCod,1,5)  ) " _
'          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
'          & " Inner Join ColocPignoraticio CP on C.cctacod = CP.cctacod " _
'          & " WHERE P.cCtaCod ='" & psctacod & "' "

    '**DAOR 20070716*****************************************************
    
    'Se agrego (P.dPrdestado, C.nMontoCol) PEAC 20070813
'    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, P.dPrdestado, C.dVenc, C.nPlazo,isnull(cp.cCredB,'')cCredB, CP.nNroRenov,  " _
'          & " P.nSaldo, C.nMontoCol, CP.nTasacion, P.nTasaInteres, " _
'          & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
'          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntCompVencida & " ) ,  " _
'          & " nTasaIntMora = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
'          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ) ,  " _
'          & " cInicRemate = (SELECT CASE WHEN COUNT(RG.nRGEstado)> 0  THEN 'S' ELSE 'N' END " _
'          & "                FROM ColocPigRecGar RG Join ColocPigRGDet RGD ON RG.cTpoProceso = RGD.cTpoProceso And RG.cNroProceso = RGD.cNroProceso " _
'          & "                WHERE RG.cTpoProceso ='R' AND DATEDIFF(D, RG.dProceso,'" & Format(pdFecSis, "mm/dd/yyyy") & "') = 0  " _
'          & "                AND RG.nRGEStado ='" & gColPRecGarEstIniciado & "' AND RGD.cCtaCod = C.cCtaCod AND RG.cCodAge = Substring(C.cCtaCod,1,5)  ), " _
'          & " nEstUltProcRem=isnull((select top 1 nRGDetEstado from ColocPigRGDET " _
'          & "                       where cCodAge=Substring(C.cCtaCod,1,5) and cTpoProceso='R' and cCtaCod=C.cCtaCod " _
'          & "                       order by cNroProceso desc),0) " _
'          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
'          & " Inner Join ColocPignoraticio CP on C.cctacod = CP.cctacod " _
'          & " WHERE P.cCtaCod ='" & psctacod & "' "

'*** PEAC 20080208
lsSQL = " exec stp_sel_ObtieneDatosParaCancelacionPigno '" & _
        Format(pdFecSis, "yyyymmdd") & "','" & psCtaCod & "'"

    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    
    If lrValida.RecordCount = 0 Then
         psmensaje = "ERROR: No existen Datos "
        Exit Function
    End If
    
    If (lrValida!nPrdEstado <> gColPEstDesem) And (lrValida!nPrdEstado <> gColPEstVenci) And _
       (lrValida!nPrdEstado <> gColPEstPRema) And (lrValida!nPrdEstado <> gColPEstRenov) Then
        psmensaje = " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado)
        Exit Function
    End If
    ' Si el Proceso Remate se ha iniciado ' LAYG (21/07/2003)
    If lrValida!cInicRemate = "S" Then
        psmensaje = " Contrato se encuentra EN EL PROCESO DE REMATE, no se puede Cancelar "
        Exit Function
    End If
    
    Set nValidaCancelacionCredPignoraticio = lrValida
    Set lrValida = Nothing
End Function


Public Function nValidaCustodiaDiferidaCredPignoraticio(ByVal psCtaCod As String, ByVal pdFecSis As Date, _
     ByVal pnMaxDiasDesembolso As Integer, Optional ByVal cUser As String, Optional ByRef psmensaje As String) As ADODB.Recordset

'************************************

Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset
Set lrValida = New ADODB.Recordset

'    lsSQL = " select m.nmovnro,m.cmovnro,mc.copecod, mcd.nprdconceptocod, montodet=mcd.nmonto,montot=mc.nmonto "
'    lsSQL = lsSQL & " into #tmpPrend" & cUser & " from mov m"
'    lsSQL = lsSQL & " join movcol mc on mc.nmovnro=m.nmovnro "
'    lsSQL = lsSQL & " join movcoldet mcd on mcd.nmovnro=mc.nmovnro and mcd.copecod=mc.copecod "
'    lsSQL = lsSQL & " where mc.cctacod='" & psCtaCod & "' and m.nmovflag=0 and m.nmovestado=10 "
'    oCon.Ejecutar lsSQL
'
'
'   lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, C.nPlazo, cp.nTasacion , " _
'          & " dCancelado = (SELECT top 1 CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
'          & "              FROM #tmpPrend" & cUser & " " _
'          & "              WHERE cOpeCod  in ( '" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" & gColPOpeCanceNorEnOtCjEFE & "','" & gColPOpeCanceMorEnOtCjEFE & "')  ),  " _
'          & " nCustodiaPag = ISNULL( (SELECT SUM(montodet) FROM #tmpPrend" & cUser & " " _
'          & "                       WHERE cOpeCod = " & geColPCobCusDiferida & "  ) , 0 ) " _
'          & " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = c.cCtaCod " _
'          & " INNER JOIN ColocPignoraticio CP ON c.cCtaCod = cp.cCtaCod " _
'          & " WHERE P.cCtaCod = '" & psCtaCod & "' "
'    lrValida.CursorLocation = adUseClient
'    lrValida.Open lsSQL, oCon, adOpenStatic, adLockReadOnly, adCmdText
'    Set lrValida.ActiveConnection = Nothing
'
'   lsSQL = " drop table #tmpPrend" & cUser
'   oCon.Ejecutar lsSQL
'ALPA 20100609 B2*********************************************************************
'    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, C.nPlazo, cp.nTasacion , " _
'          & " dCancelado = (SELECT CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
'          & "              FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
'          & "              WHERE MC.cCtaCod = C.cCtaCod AND MC.cOpeCod  in ( '" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" & gColPOpeCanceNorEnOtCjEFE & "','" & gColPOpeCanceMorEnOtCjEFE & "') AND M.nMovFlag <> " & gMovFlagExtornado & " ),  " _
'          & " nCustodiaPag = ISNULL( (SELECT SUM(nMonto) FROM MovColDet MCD " _
'          & "                       INNER JOIN Mov M ON M.nMovNro = MCD.nMovNro " _
'          & "                       WHERE MCD.nMovNro = M.nMovNro  AND m.nmovflag=0  AND MCD.cOpeCod = " & geColPCobCusDiferida & " AND MCD.cCtaCod = P.cCtaCod ) , 0 ) " _
'          & " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = c.cCtaCod " _
'          & " INNER JOIN ColocPignoraticio CP ON c.cCtaCod = cp.cCtaCod " _
'          & " WHERE P.cCtaCod = '" & psctacod & "' "
 lsSQL = " exec stp_sel_ObtieneDatosParaValidaCustodiaDiferidaCredPignoraticio '" & psCtaCod & "'"
 '*************************************************************************************
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    
    If Not (lrValida.EOF And lrValida.BOF) Then
        If (lrValida!nPrdEstado <> gColPEstDifer) Then
            psmensaje = " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado)
            Exit Function
        End If
    End If
    Set nValidaCustodiaDiferidaCredPignoraticio = lrValida
    Set lrValida = Nothing
    
End Function



Public Function nValidaDuplicadoContratoCredPignoraticio(ByVal psCtaCod As String, Optional ByRef psmensaje As String) As ADODB.Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset
'ALPA 20100610 B2*****************************************************************
'    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, dVigencia, P.nTasaInteres, CP.nNroDuplic, " _
'         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodInteresCompensatorio & "   then nMonto else 0 end )  as nInteres," _
'         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodImpuesto & "  then nMonto else 0 end )  as nImpuesto," _
'         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodTasacion & "  then nMonto else 0 end )  as nTasacion," _
'         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodCustodia & "  then nMonto else 0 end )  as nCustodia " _
'         & " FROM Colocaciones c Inner join Producto P on P.cCtaCod = c.cCtaCod " _
'         & " Inner Join ColocPignoraticio CP On C.cCtaCod = CP.cCtaCod Inner join ColocCalendario CCal on CCal.cCtaCod = C.cCtaCod " _
'         & " Inner join ColoccalendDetPig CCalD on CCal.cCtaCod = CCalD.cCtaCod and CCal.nNroCalen = CCalD.nNroCalen " _
'         & " And CCal.nColocCalendApl = CCalD.nColocCalendApl " _
'         & " Where P.cCtaCod  = '" & psctacod & "' AND CCalD.nColocCalendApl =  " & gColocCalendAplDesembolso & " " _
'         & " GROUP BY P.cCtaCod, P.nPrdEstado, dVigencia, P.nTasaInteres, CP.nNroDuplic "
    
    'Set lrValida = New ADODB.Recordset
    lsSQL = "exec stp_sel_ObtieneValidaDuplicadoContratoCredPignoraticio '" & psCtaCod & "'"
'*********************************************************************************
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    If lrValida.BOF And lrValida.EOF Then
        psmensaje = "ERROR: al Buscar datos de Contrato "
        Exit Function
    End If
    If lrValida!nPrdEstado <> gColPEstRegis And lrValida!nPrdEstado <> gColPEstDesem _
        And lrValida!nPrdEstado <> gColPEstDifer And lrValida!nPrdEstado <> gColPEstCance _
        And lrValida!nPrdEstado <> gColPEstPRema And lrValida!nPrdEstado <> gColPEstRenov And lrValida!nPrdEstado <> gColPEstVenci _
        And lrValida!nPrdEstado <> gColPEstAdjud Then 'JGPA20170822 AGREGO gColPEstAdjud - SEGUN RFC1705120001
        psmensaje = " Contrato se encuentra " '& ContratoEstado(lrvalida!ncEstado)
        Exit Function
    End If
    Set nValidaDuplicadoContratoCredPignoraticio = lrValida
    Set lrValida = Nothing
End Function

Public Function nValidaRegVentaRemateCredPignoraticio(ByVal psCtaCod As String, ByVal psTipoProceso As String, _
    ByVal psNroProceso As String, Optional ByRef psmensaje As String) As ADODB.Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset
'ALPA 20100609 B2 *************************************************************
'    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, CRGDET.nRGDetEstado, CRGDET.nRemSubBaseVta,  " _
'          & " P.nSaldo, CP.nTasacion , C.dVenc, " _
'          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
'          & "               WHERE cLineaCred = C.cLineaCred  " _
'          & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ) " _
'          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
'          & " LEFT JOIN ColocPigRGDet CRGDET ON CP.cCtaCod = CRGDet.cCtaCod " _
'          & " WHERE CRGDET.cTpoProceso ='" & psTipoProceso & "' AND CRGDET.cNroProceso ='" & psNroProceso & "' " _
'          & " AND CRGDET.cCtaCod ='" & psctacod & "' "
lsSQL = "exec stp_sel_ObtieneDatosParaValidaRegVentaRemateCredPignoraticio '" & psCtaCod & "','" & psTipoProceso & "','" & psNroProceso & "'"
'******************************************************************************
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida.BOF And lrValida.EOF Then
        psmensaje = "El Credito no se encuetra en el Proceso de Remate "
        Exit Function
    End If
    
    If (lrValida!nPrdEstado <> gColPEstRemat) And (lrValida!nPrdEstado <> gColPEstPRema) And _
       (lrValida!nPrdEstado <> gColPEstVenci) And (lrValida!nPrdEstado <> gColPEstRenov) Then
        psmensaje = " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado)
        Exit Function
    End If
    If lrValida!nPrdEstado = gColPEstRemat And lrValida!nRgDetEstado = gColPRecGarVentaVendido Then
        psmensaje = " Contrato ya fue rematado " & psCtaCod
        Exit Function
    End If
    
    Set nValidaRegVentaRemateCredPignoraticio = lrValida
    Set lrValida = Nothing
End Function


Public Function nValidaPagoSobranteRemateCredPignoraticio(ByVal psCtaCod As String, ByVal psTipoProceso As String, Optional ByRef psmensaje As String) As ADODB.Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset
'ALPA 20100609 B2*************************************************************************
'    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, CRGDET.nRGDetEstado, CRGDET.nMontoProceso, CRGDET.nRemSubBaseVta,  " _
'          & " P.nSaldo, CP.nTasacion, CRGDET.cNroProceso, CVta.nSobrante, CVta.nComision " _
'          & " FROM Producto P Inner Join ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
'          & " LEFT JOIN ColocPigRGDet CRGDET ON CP.cCtaCod = CRGDet.cCtaCod  " _
'          & " LEFT JOIN ColocPigRecupVta CVta ON CRGDet.cCtaCod = CVta.cCtaCod  " _
'          & " WHERE CRGDET.cTpoProceso ='" & psTipoProceso & "' " _
'          & " AND CRGDET.cCtaCod ='" & psctacod & "' " _
'          & " ORDER BY CRGDET.cNroProceso DESC "
'*****************************************************************************************
    lsSQL = "exec stp_sel_ObtieneDatosParaValidaPagoSobranteRemateCredPignoraticio '" & psCtaCod & "','" & psTipoProceso & "'"
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
'
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        psmensaje = "AVISO: Contrato Pignoraticio No ha sido Rematado  "
        Exit Function
    End If
    If lrValida.BOF And lrValida.EOF Then
        psmensaje = "AVISO: Contrato Pignoraticio No ha sido Rematado  "
        Exit Function
    End If
    If lrValida!nRgDetEstado = geColPRecGarVentaSobranteCobrado Then
        psmensaje = " Ya fue cobrado el sobrante de Contrato " & psCtaCod
        Exit Function
    ElseIf lrValida!nRgDetEstado = gColPRecGarVentaNoVendido Then
        psmensaje = " Contrato No Fue Vendido en Remate " & psCtaCod
        Exit Function
    End If
    
    Set nValidaPagoSobranteRemateCredPignoraticio = lrValida
    Set lrValida = Nothing
End Function


Public Function nValidaRegVentaSubastaCredPignoraticio(ByVal psCtaCod As String, ByVal psTipoProceso As String, _
    ByVal psNroProceso As String, ByRef psmensaje As String) As ADODB.Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset
'ALPA 201006009 B2***************************************************************
'    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, CRGDET.nRGDetEstado,  " _
'          & " P.nSaldo, CP.nTasacion, CRGDET.nRemSubBaseVta,c.dprdestado,  " _
'          & " nValRegistroAdj = ISNULL( (Select nAdjValRegistro FROM ColocPigRGDET  " _
'          & "                            WHERE cTpoProceso = 'A' AND cCtaCod = P.cCtaCod ) , 0 ),bExcepVta=isnull(cp.bExcepVta,0)  " _
'          & " FROM Producto P Inner Join ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
'          & " LEFT JOIN ColocPigRGDet CRGDET ON CP.cCtaCod = CRGDet.cCtaCod " _
'          & " JOIN ColocaCestado C on C.cctacod=P.cctacod  and c.nprdestado='" & gColPEstAdjud & "'" _
'          & " WHERE p.cCtaCod ='" & psctacod & "' and p.nprdestado='" & gColPEstAdjud & "'"

          ' " and CRGDET.cTpoProceso ='" & psTipoProceso & "'"
          'AND CRGDET.cNroProceso ='" & psNroProceso & "'
          
lsSQL = "exec stp_sel_ObtieneDatosParaValidaRegVentaSubastaCredPignoraticio '" & psCtaCod & "','" & psTipoProceso & "','" & psNroProceso & "'"
'********************************************************************************
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida.BOF And lrValida.EOF Then
        psmensaje = "El Credito no se encuetra en el Proceso de Subasta "
        Exit Function
    End If
    
    If (lrValida!nPrdEstado <> gColPEstAdjud) Then
        psmensaje = " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado)
        Exit Function
    End If
    If lrValida!nRgDetEstado = gColPRecGarVentaVendido Then
        psmensaje = " Contrato ya fue Subastado " & psCtaCod
        Exit Function
    End If
    
    Set nValidaRegVentaSubastaCredPignoraticio = lrValida
    Set lrValida = Nothing
End Function

Public Function nValidaBloqueoCredPignoraticio(ByVal psCtaCod As String, ByRef psmensaje As String) As ADODB.Recordset

'************************************

Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset

lsSQL = "SELECT P.nPrdEstado, PBloq.nBlqTpo, PBloq.nBlqMotivo, ISNULL(PBloq.cComentario,'') cComentario, PBloq.cMovNro as cMovNroBloqueo , " _
    & " cBloqueo = (SELECT CASE WHEN COUNT(nBlqMotivo) > 0 THEN 'S' ELSE 'N' END FROM ProductoBloqueos PB WHERE PB.cCtaCod = '" & psCtaCod & "' AND cMovNroDbl Is Null ) " _
    & " FROM Producto P " _
    & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod  " _
    & " LEFT JOIN ProductoBloqueos PBloq ON CP.cCtaCod = PBloq.cCtaCod " _
    & " WHERE P.cCtaCod ='" & psCtaCod & "' ORDER BY cMovNro DESC " ' & "' AND cMovNroDbl Is Null "
     
    Set lrValida = New ADODB.Recordset
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Or (lrValida.BOF And lrValida.EOF) Then
        psmensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    
    If (lrValida!nPrdEstado <> gColPEstRegis) And (lrValida!nPrdEstado <> gColPEstDesem) And _
       (lrValida!nPrdEstado <> gColPEstDifer) And (lrValida!nPrdEstado <> gColPEstVenci) And _
       (lrValida!nPrdEstado <> gColPEstPRema) And (lrValida!nPrdEstado <> gColPEstVenci) Then
        psmensaje = " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado)
        Exit Function
    End If
    
    Set nValidaBloqueoCredPignoraticio = lrValida
    Set lrValida = Nothing
    
End Function


Public Function nBuscaOperacionesCredPigParaExtorno(ByVal psFecTrans As String, ByVal psOpeExtor As String, _
        Optional ByVal psCtaCod As String = "@", Optional ByVal pbOrdenAsc As Boolean = False, Optional ByRef psmensaje As String) As ADODB.Recordset

Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset

'CTI4 ERS0112020 Comento
'    lsSQL = " SELECT M.cMovNro, M.nMovNro, M.cOpeCod, M.nMovFlag, " _
'          & " MC.cCtaCod, MC.nMonto, Ope.cOpeDesc,  " _
'          & " (Select Max(P.cPersNombre) From Persona p Join ProductoPersona pp on p.cPersCod = pp.cPersCod " _
'          & "    Where pp.cCtaCod = MC.cCtaCod And pp.nPrdPersRelac = " & gColRelPersTitular & "  ) cCliente , " _
'          & " Isnull((Select mcm.cPersCod from movcmac mcm where mcm.nmovnro = m.nmovnro),'' ) cCMACOpe  " _
'          & " FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
'          & " LEFT JOIN OpeTpo Ope ON MC.cOpeCod = Ope.cOpeCod " _
'          & " WHERE M.cMovNro LIKE '" & psFecTrans & "%' " _
'          & " AND M.nMovFlag = " & gMovFlagVigente & " " _
'          & " AND Mc.cOpeCod in " & psOpeExtor & "  " _
'          & " AND Substring(M.cOpeCod,3,1) <> '" & Mid(geColPExtDupli, 3, 1) & "' "
'    If psCtaCod <> "@" Then
'        lsSQL = lsSQL & " AND MC.cCtaCod ='" & psCtaCod & "' "
'    End If
'    lsSQL = lsSQL & " ORDER BY M.nMovNro DESC  "

'CTI4 ERS0112020
    lsSQL = "exec stp_sel_BuscaOperacionesCredPigParaExtorno '" & psFecTrans & "','" & Replace(psOpeExtor, "'", "''") & "','" & psCtaCod & "'"
'END CTI4
     
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Extorno "
        Exit Function
    End If
    
    Set nBuscaOperacionesCredPigParaExtorno = lrValida
    Set lrValida = Nothing
End Function
'
Public Function nBuscaOperacionesCredPigParaExtornoLLamadaCMAC(ByVal psFecTrans As String, Optional ByVal psCtaCod As String = "@", _
        Optional ByVal pbOrdenAsc As Boolean = False, Optional ByRef psmensaje As String) As ADODB.Recordset

Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = " SELECT M.cMovNro, M.nMovNro, M.cOpeCod, M.nMovFlag, " _
          & " MC.cCtaIF as cCtaCod, MC.nMonto, Ope.cOpeDesc  " _
          & " FROM Mov M INNER JOIN MovCMAC MC ON M.nMovNro = MC.nMovNro " _
          & " LEFT JOIN OpeTpo Ope ON MC.cOpeCod = Ope.cOpeCod " _
          & " WHERE M.cMovNro LIKE '" & psFecTrans & "%' " _
          & " AND M.nMovFlag = " & gMovFlagVigente & " " _
          & " AND M.cOpeCod LIKE '" & Mid(geColPExtDesemb, 1, 2) & "%' " _
          & " AND Substring(M.cOpeCod,3,1) <> '" & Mid(geColPExtDupli, 3, 1) & "' "
    'If psCtaCod <> "@" Then
    '    lsSQL = lsSQL & " AND MC.cCtaCod ='" & psCtaCod & "' "
    'End If
    lsSQL = lsSQL & " ORDER BY M.nMovNro DESC  "
    
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Extorno "
        Exit Function
    End If
    
    Set nBuscaOperacionesCredPigParaExtornoLLamadaCMAC = lrValida
    Set lrValida = Nothing
End Function


Public Function nDocumentoEmitido(ByVal pnDocTpo As Integer, ByVal psDocNro As String, _
            ByVal psOpeCod As String) As Boolean
'Verifica si un documento ha sido emitido
Dim loValida As COMDColocPig.DCOMColPFunciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

On Error GoTo dError

lsSQL = "SELECT Count(MDoc.nDocTpo) as Cuenta " _
    & " FROM MovDoc MDoc INNER JOIN Mov M ON MDoc.nMovNro = M.nMovNro " _
    & " WHERE MDoc.nDocTpo = " & pnDocTpo & " AND MDoc.cDocNro ='" & psDocNro & "' " _
    & " AND M.cOpeCod IN ( " & psOpeCod & " ) "

    Set loValida = New COMDColocPig.DCOMColPFunciones
        Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    If lrs!Cuenta > 0 Then
        nDocumentoEmitido = True
    Else
        nDocumentoEmitido = False
    End If
Set lrs = Nothing
Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
End Function

Public Function nValidaAmortizacionCredPignoraticio(ByVal psCtaCod As String, ByVal pdFecSis As Date, _
     ByVal pnMaxDiasDesembolso As Integer, Optional ByVal cUser As String, Optional ByRef psmensaje As String, Optional ByRef pnCredRevolAntNue As Boolean = False) As ADODB.Recordset
    
    '*** PEAC 20161105 - se agregó pnCredRevolAntNue

'************************************
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones

Dim lrValida As ADODB.Recordset

Set lrValida = New ADODB.Recordset

'    lsSQL = " select m.nmovnro,m.cmovnro,mc.copecod, mcd.nprdconceptocod, montodet=mcd.nmonto,montot=mc.nmonto "
'    lsSQL = lsSQL & " into #tmpPrend" & cUser & " from mov m"
'    lsSQL = lsSQL & " join movcol mc on mc.nmovnro=m.nmovnro "
'    lsSQL = lsSQL & " join movcoldet mcd on mcd.nmovnro=mc.nmovnro and mcd.copecod=mc.copecod "
'    lsSQL = lsSQL & " where mc.cctacod='" & psCtaCod & "' and m.nmovflag=0 and m.nmovestado=10 "
'    oCon.Ejecutar lsSQL
'
'    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, C.nPlazo, CP.nNroRenov,  " _
'          & " P.nSaldo, CP.nTasacion, P.nTasaInteres, " _
'          & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
'          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ),  " _
'          & " dDesembolso = (SELECT MAX (CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
'          & "                FROM #tmpPrend" & cUser & " " _
'          & "                WHERE cOpeCod = " & gColPOpeDesembolsoEFE & " ),  " _
'          & " dRenovacion = isnull((SELECT MAX (CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
'          & "              FROM #tmpPrend" & cUser & "" _
'          & "              WHERE cOpeCod  in ( " & gColPOpeRenovNorEFE & "," & gColPOpeRenovMorEFE & "," _
'          & gColPOpeRenovNorEnOtCjEFE & "," & gColPOpeRenovMorEnOtCjEFE & ") ),'01/01/1900') ,  " _
'          & " cInicRemate = (SELECT CASE WHEN COUNT(RG.nRGEstado)> 0  THEN 'S' ELSE 'N' END " _
'          & "                FROM ColocPigRecGar RG Join ColocPigRGDet RGD ON RG.cTpoProceso = RGD.cTpoProceso And RG.cNroProceso = RGD.cNroProceso " _
'          & "                WHERE RG.cTpoProceso ='R' AND DATEDIFF(D, RG.dProceso,'" & Format(pdFecSis, "mm/dd/yyyy") & "') = 0  " _
'          & "                AND RG.nRGEStado ='" & gColPRecGarEstIniciado & "' AND RGD.cCtaCod = C.cCtaCod AND RG.cCodAge = Substring(C.cCtaCod,1,5)  )" _
'          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
'          & " Inner Join ColocPignoraticio CP on C.cctacod = CP.cctacod " _
'          & " WHERE P.cCtaCod ='" & psCtaCod & "' "
'
'    lrValida.CursorLocation = adUseClient
'    lrValida.Open lsSQL, oCon, adOpenStatic, adLockReadOnly, adCmdText
'    Set lrValida.ActiveConnection = Nothing
'
'
'    lsSQL = " drop table #tmpPrend" & cUser
'    oCon.Ejecutar lsSQL

    ' se agrego P.dPrdestado,C.nMontoCol PEAC 20070813
'    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, P.dPrdestado, C.nMontoCol, C.dVenc, C.nPlazo, isnull(cp.cCredB,'')cCredB, CP.nNroRenov,  " _
'          & " P.nSaldo, CP.nTasacion, P.nTasaInteres, " _
'          & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
'          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ),  " _
'          & " dDesembolso = (SELECT MAX (CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
'          & "              FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
'          & "              WHERE MC.cCtaCod = C.cCtaCod AND MC.cOpeCod  in ( " & gColPOpeDesembolsoEFE & ") AND M.nMovFlag <> " & gMovFlagExtornado & " ),  " _
'          & " dRenovacion = isnull((SELECT MAX (CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
'          & "              FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
'          & "              WHERE MC.cCtaCod = C.cCtaCod AND MC.cOpeCod  in ( " & gColPOpeRenovNorEFE & "," & gColPOpeRenovMorEFE & "," _
'          & gColPOpeRenovNorEnOtCjEFE & "," & gColPOpeRenovMorEnOtCjEFE & ") AND M.nMovFlag <> " & gMovFlagExtornado & " ),'01/01/1900') ,  " _
'          & " cInicRemate = (SELECT CASE WHEN COUNT(RG.nRGEstado)> 0  THEN 'S' ELSE 'N' END " _
'          & "                FROM ColocPigRecGar RG Join ColocPigRGDet RGD ON RG.cTpoProceso = RGD.cTpoProceso And RG.cNroProceso = RGD.cNroProceso " _
'          & "                WHERE RG.cTpoProceso ='R' AND DATEDIFF(D, RG.dProceso,'" & Format(pdFecSis, "mm/dd/yyyy") & "') = 0  " _
'          & "                AND RG.nRGEStado ='" & gColPRecGarEstIniciado & "' AND RGD.cCtaCod = C.cCtaCod AND RG.cCodAge = Substring(C.cCtaCod,1,5)  )" _
'          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
'          & " Inner Join ColocPignoraticio CP on C.cctacod = CP.cctacod " _
'          & " WHERE P.cCtaCod ='" & psctacod & "' "

    '*** PEAC 20080208
    lsSQL = " exec stp_sel_ObtieneDatosParaAmortizarPigno '" & _
        Format(pdFecSis, "yyyymmdd") & "','" & psCtaCod & "'"

    Set loRegValida = New COMDColocPig.DCOMColPFunciones

    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)


    If lrValida Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    ElseIf lrValida.BOF And lrValida.EOF Then
        psmensaje = "ERROR: No se encontro el Credito "
        Exit Function
    'End If
    ElseIf (lrValida!nPrdEstado <> gColPEstDesem) And (lrValida!nPrdEstado <> gColPEstVenci) And _
       (lrValida!nPrdEstado <> gColPEstPRema) And (lrValida!nPrdEstado <> gColPEstRenov) Then
        psmensaje = " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado)
        Exit Function
    End If
    
    ' Si ha desembolsado el mismo dia
    'If DateDiff("d", pdFecSis, lrValida!dDesembolso) = 0 Then
    '   If MsgBox("Contrato ha sido Desembolsado el dia de hoy día " & vbCr & "Desea continuar ? ", vbQuestion + vbYesNo, " Aviso ") = vbNo Then
    '        Exit Function
    '   End If
    'End If
    ' Si ha renovado el mismo dia
    'If Not IsNull(lrValida!dRenovacion) And DateDiff("d", pdFecSis, lrValida!dRenovacion) = 0 Then
    '   If MsgBox("Contrato ha sido Renovado el dia de hoy día " & vbCr & "Desea continuar ? ", vbQuestion + vbYesNo, " Aviso ") = vbNo Then
    '        Exit Function
    '   End If
    'End If
       
    ' Si No se ha definido la tasa de Interes Moratoria
    If IsNull(lrValida!nTasaIntVenc) Then
        psmensaje = " No se ha definido la Tasa de Interes Vencida para linea de Credito "
        Exit Function
    End If
        
    ' Si el Proceso Remate se ha iniciado ' LAYG (21/07/2001)
    If lrValida!cInicRemate = "S" Then
        psmensaje = " Contrato se encuentra EN EL PROCESO DE REMATE, no se puede Amortizar "
        Exit Function
    End If

    '*** PEAC 20160916
    If pnCredRevolAntNue = False Then 'true=antiguo, false=nuevo
        If lrValida!nVigenciaCredPigno = 1 Then '*** 1 vencido / 0 no vence
            psmensaje = "El crédito debe Cancelarse ya que superó el límite de " & Trim(Str(lrValida!nVigenciaParam)) & " meses de vigencia."
            Exit Function
        ElseIf lrValida!nVigenciaParaAlerta = 1 Then
            psmensaje = lrValida!cMensajeParaAlerta
        End If
    End If

    Set nValidaAmortizacionCredPignoraticio = lrValida

    Set lrValida = Nothing

End Function
'*** PEAC
Public Function nValidaPagoParcialCredPignoraticio(ByVal psCtaCod As String, ByVal pdFecSis As Date, _
     ByVal pnMaxDiasDesembolso As Integer, Optional ByVal cUser As String, Optional ByRef psmensaje As String, Optional ByRef pnCredRevolAntNue As Boolean = False) As ADODB.Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones

Dim lrValida As ADODB.Recordset

Set lrValida = New ADODB.Recordset

    lsSQL = " exec stp_sel_ObtieneDatosParaPagoParcial '" & _
        Format(pdFecSis, "yyyymmdd") & "','" & psCtaCod & "'"

    Set loRegValida = New COMDColocPig.DCOMColPFunciones

    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)

    If lrValida Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Validación "
        Exit Function
    End If
    If lrValida.BOF And lrValida.EOF Then
        psmensaje = "ERROR: No se encontro el Credito "
        Exit Function
    End If
    If (lrValida!nPrdEstado <> gColPEstDesem) And (lrValida!nPrdEstado <> gColPEstVenci) And _
       (lrValida!nPrdEstado <> gColPEstPRema) And (lrValida!nPrdEstado <> gColPEstRenov) Then
        psmensaje = " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado)
        Exit Function
    End If
       
    ' Si No se ha definido la tasa de Interes Moratoria
    If IsNull(lrValida!nTasaIntVenc) Then
        psmensaje = " No se ha definido la Tasa de Interes Vencida para linea de Credito "
        Exit Function
    End If
        
    ' Si el Proceso Remate se ha iniciado ' LAYG (21/07/2001)
    If lrValida!cInicRemate = "S" Then
        psmensaje = " Contrato se encuentra EN EL PROCESO DE REMATE, no se puede Amortizar "
        Exit Function
    End If

    '*** PEAC 20160916
'    If pnCredRevolAntNue = False Then 'true=antiguo, false=nuevo
'        If lrValida!nVigenciaCredPigno = 1 Then '*** 1 vencido / 0 no vence
'            psmensaje = "El crédito debe Cancelarse ya que superó el límite de " & Trim(Str(lrValida!nVigenciaParam)) & " meses de vigencia."
'            Exit Function
'        End If
'    End If

    If pnCredRevolAntNue = False Then 'true=antiguo, false=nuevo
        If lrValida!nVigenciaCredPigno = 1 Then '*** 1 vencido / 0 no vence
            psmensaje = "El crédito debe Cancelarse ya que superó el límite de " & Trim(Str(lrValida!nVigenciaParam)) & " meses de vigencia."
            Exit Function
        ElseIf lrValida!nVigenciaParaAlerta = 1 Then
            psmensaje = lrValida!cMensajeParaAlerta
        End If
    End If


    Set nValidaPagoParcialCredPignoraticio = lrValida

    Set lrValida = Nothing

End Function

Public Function nValidaRegVentaSubastaCredPignoraticioSINSubasta(ByVal psCtaCod As String, ByVal psTipoProceso As String, _
    ByVal psNroProceso As String, ByRef psmensaje As String) As ADODB.Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset

    '*** PEAC 20080901
'    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado,p.dprdestado, " _
          & " P.nSaldo, CP.nTasacion,nValRegistroAdj = isnull(RGDET.NADJVALREGISTRO,0),CnroProceso=isnull(RGDET.CnroProceso,'00'),CCODAGE=isnull(RGDET.CCODAGE,'00'),CNROPROCESOREM=isnull(RGDET.CNROPROCESOREM,'00'),DPROCESOREM=isnull(RGDET.DPROCESOREM,'01/01/1900'),AGEREM=isnull(RGDET.AGEREM,'00'),bExcepVta=isnull(cp.bExcepVta,0)  " _
          & " FROM Producto P Inner Join ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN (select RG.CTPOPROCESO,RG.CNROPROCESO,RG.DPROCESO,RG.CCODAGE,D.CCTACOD,D.NREMSUBBASEVTA,D.NADJVALREGISTRO,D.NDEUDA,REM.CNROPROCESOREM,REM.DPROCESOREM,REM.AGEREM " _
          & "            from colocpigrecgar RG " _
          & "            JOIN colocpigrgDET D ON D.CNROPROCESO=RG.CNROPROCESO AND D.CTPOPROCESO=RG.CTPOPROCESO " _
          & "            JOIN (Select CNROPROCESOREM=CNROPROCESO,DPROCESOREM=DPROCESO,AGEREM=CCODAGE FROM colocpigrecgar where CTPOPROCESO='R' ) REM on CONVERT(CHAR(8),REM.DPROCESOREM,112)=CONVERT(CHAR(8),RG.DPROCESO,112) " _
          & "            WHERE RG.CTPOPROCESO='A') RGDET ON RGDET.CCTACOD=P.CCTACOD  " _
          & " WHERE  P.cCtaCod ='" & psctacod & "'"
    
    '*** PEAC 20080901
    lsSQL = " exec stp_sel_ObtieneDatosParaRecupDeContratoAdjudicado '" & psCtaCod & "'"
          
          '& " nValRegistroAdj = ISNULL( (Select nAdjValRegistro FROM ColocPigRGDET  " _
          '& "                            WHERE cTpoProceso = 'A' AND cCtaCod = P.cCtaCod ) , 0 ), " _
          '& " CnroProceso= ISNULL( (Select cNroProceso FROM ColocPigRGDET  " _
          '& "                            WHERE cTpoProceso = 'A' AND cCtaCod = P.cCtaCod ) , '' ), " _
          '& " Ccodage= ISNULL( (Select Ccodage FROM ColocPigRGDET  " _
          '& "                            WHERE cTpoProceso = 'A' AND cCtaCod = P.cCtaCod ) , '' ) "
          

    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida.BOF And lrValida.EOF Then
        psmensaje = "El Credito no se encuetra en el Proceso de Subasta "
        Exit Function
    End If
    If (lrValida!nPrdEstado <> gColPEstAdjud) Then
        psmensaje = " Contrato se encuentra " & lrValida!cEstado '& mfgEstadoCredPigDesc(lrValida!nPrdEstado)
        Exit Function
    End If
    'If lrValida!nRGDetEstado = gColPRecGarVentaVendido Then
    '    MsgBox " Contrato ya fue Subastado " & psCtaCod, vbInformation, " Aviso "
    '    Exit Function
    'End If
    
    Set nValidaRegVentaSubastaCredPignoraticioSINSubasta = lrValida
    Set lrValida = Nothing
End Function

'RECO20160513 **************************************************************
Public Sub PignoActualizaDirecCliente(ByVal psPersCod As String, ByVal psCtaCod As String, ByVal pnActualiza As String)
    Dim loRegValida As COMDColocPig.DCOMColPFunciones
    Dim lrValida As ADODB.Recordset
    Dim lsSQL As String

    lsSQL = " exec stp_ins_ActualizaDirecCliente '" & psPersCod & "','" & psCtaCod & "'," & pnActualiza

    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    
    loRegValida.dObtieneRecordSet (lsSQL)
    
End Sub

'RECO FIN*******************************************************************
'Muestra para Remate
Public Function CalcDeudaRemate(ByVal psFecRemate As String, ByVal psAgencia As String, ByVal psNroRemate As String, ByVal psCtaCod As String, Optional ByRef pnSaldoDeuda As Double, Optional ByRef pnIntVencido As Double, Optional ByRef pnComision As Double) As Double
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NCOMColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
Dim pnPreOro14 As Double, pnPreOro16 As Double, pnPreOro21 As Double, pnPreOro18 As Double
Dim pnTasaCustodiaVencida As Double, pnTasaPreparacionRemate As Double, pnTasaImpuesto As Double, pnFactorPrecioVenta As Double

Dim lrCredPigJoyasDet As ADODB.Recordset
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim item As Integer, itemTMP As Integer
Dim lmDetalle(20) As String

pnPreOro14 = 0: pnPreOro16 = 0: pnPreOro21 = 0: pnPreOro18 = 0

    lsSQL = "select * from colocpigrgprecio where ctpoproceso='R' and  cnroproceso='" & psNroRemate & "' and ccodage='" & psAgencia & "'"
    
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    While Not lrDataRep.EOF
            Select Case lrDataRep!ctpoprecio
            
                Case 14
                        pnPreOro14 = lrDataRep!nPrecio
                Case 16
                        pnPreOro16 = lrDataRep!nPrecio
                Case 18
                        pnPreOro18 = lrDataRep!nPrecio
                Case 21
                        pnPreOro21 = lrDataRep!nPrecio
                Case 3005
                        pnTasaPreparacionRemate = lrDataRep!nPrecio
                Case 3004
                        pnTasaCustodiaVencida = lrDataRep!nPrecio
                Case 3003
                        pnTasaImpuesto = lrDataRep!nPrecio
                Case 3041
                        pnFactorPrecioVenta = lrDataRep!nPrecio
                
       'pnTasaCustodiaVencida As Double, pnTasaPreparacionRemate As Double, pnTasaImpuesto As Double, pnFactorPrecioVenta As Double
                
            End Select
            lrDataRep.MoveNext
    Wend
    Set lrDataRep = Nothing
    Set loDataRep = Nothing

    'ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, psFecRemate)

    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, "
    lsSQL = lsSQL & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote,   "
    lsSQL = lsSQL & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , "
    lsSQL = lsSQL & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , "
    lsSQL = lsSQL & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , "
    lsSQL = lsSQL & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , "
    lsSQL = lsSQL & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  "
    lsSQL = lsSQL & "               WHERE cLineaCred = C.cLineaCred  "
    lsSQL = lsSQL & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " )    "
    lsSQL = lsSQL & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod "
    lsSQL = lsSQL & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  "
    lsSQL = lsSQL & " WHERE p.cctacod='" & psCtaCod & "'"
    lsSQL = lsSQL & " GROUP BY P.cCtaCod, P.nPrdestado ,C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, "
    lsSQL = lsSQL & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote  "

'& " AND Substring(P.cCtaCod,4,2) = '" & psAgencia & "' AND DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  " _

    Set lrDataRep = New ADODB.Recordset
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If Not lrDataRep Is Nothing Then
                
        fnVarIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)

        Dim pnDeudaCosto As Double, pn15DeVenta As Double

        
        With lrDataRep
            If Not lrDataRep.EOF Then
                
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                vCuenta = 0
                
                pn15DeVenta = Round(((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * 0.15, 2)
              
                lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21))
                'lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * (fnVarIGV + 1)
                
                'lnPreBase = (v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)
                
                
                lnGramos = v14 + v16 + v18 + v21
                
                lnPreBase = Round(lnPreBase, 2)
                
                vdiasAtraso = DateDiff("d", !dVenc, psFecRemate)
                
                Set loCalculos = New NCOMColPCalculos
                    If vdiasAtraso <= 0 Then
                        vdiasAtraso = 0
                        vInteresVencido = 0
                        vCostoCustodiaMoratorio = 0
                        vImpuesto = 0
                    Else
                        vInteresVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vdiasAtraso)
                        vInteresVencido = Round(vInteresVencido, 2)
                        vCostoCustodiaMoratorio = loCalculos.nCalculaCostoCustodiaMoratorio(!nTasacion, pnTasaCustodiaVencida, vdiasAtraso)
                        vCostoCustodiaMoratorio = Round(vCostoCustodiaMoratorio, 2)
                        vImpuesto = (vInteresVencido + vCostoCustodiaMoratorio) * pnTasaImpuesto
                        vImpuesto = Round(vImpuesto, 2)
                    End If
                        vCostoPreparacionRemate = 0
                    'If !nPrdEstado = gColPEstPRema Then    ' Si esta en via de Remate
                        vCostoPreparacionRemate = pnTasaPreparacionRemate * !nTasacion
                        vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
                    'End If
                    lnDeuda = !nSaldo + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate
                    
                    pnSaldoDeuda = !nSaldo
                    pnIntVencido = vInteresVencido
                    
                    'pnSaldoDeuda pnIntVencido pnComision
                    
                   ' pnDeudaCosto = lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta
                     'pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta) * (fnVarIGV + 1), 2)
                    'Calcula el Precio Venta para Remate
                    'lnPreVenta = loCalculos.nCalculaPrecioVentaRemate(pnDeudaCosto, lnPreBase, pnFactorPrecioVenta)
                   ' lnPreVenta = Round(lnPreVenta, 2)
                   pnDeudaCosto = Round((lnDeuda + Round((lnDeuda * pnFactorPrecioVenta), 2) + pn15DeVenta), 2)
                   pnComision = Round((lnDeuda * pnFactorPrecioVenta), 2) + pn15DeVenta
                   
                   CalcDeudaRemate = pnDeudaCosto
                                     
                   Set loCalculos = Nothing
                
                
            End If
        End With
            
    End If
    
End Function

'Muestra para Remate
Public Function CalcDeudaRemateSINPROCESO(ByVal psFecEstado As String, ByVal psCtaCod As String, Optional ByRef pnSaldoDeuda As Double, Optional ByRef pnIntVencido As Double, Optional ByRef pnComision As Double, Optional ByRef psFecRemate As String) As Double
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NCOMColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
Dim pnPreOro14 As Double, pnPreOro16 As Double, pnPreOro21 As Double, pnPreOro18 As Double
Dim pnTasaCustodiaVencida As Double, pnTasaPreparacionRemate As Double, pnTasaImpuesto As Double, pnFactorPrecioVenta As Double

Dim lrCredPigJoyasDet As ADODB.Recordset
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim item As Integer, itemTMP As Integer, psNroRemate As String
Dim lmDetalle(20) As String

Select Case DatePart("m", psFecEstado)

        Case 8
                psFecRemate = "2004-08-27"
                        'Agosto 27
        Case 9
                psFecRemate = "2004-09-29"
                        'Setiembre 29
        Case 10
                psFecRemate = "2004-10-29"
                        'Octubre 29
        Case 11
                psFecRemate = "2004-11-30"
                        'Noviembre 30
        Case 12
                psFecRemate = "2004-12-29"
                        'Diciembre 29
        Case Else
                psFecRemate = psFecEstado
End Select

If DatePart("m", psFecEstado) <= 10 Then
        psNroRemate = "0064"
ElseIf DatePart("m", psFecEstado) >= 11 Then
        psNroRemate = "0065"
End If

pnPreOro14 = 0: pnPreOro16 = 0: pnPreOro21 = 0: pnPreOro18 = 0

    lsSQL = "select * from colocpigrgprecio where ctpoproceso='R' and  cnroproceso='" & psNroRemate & "'"
    'and ccodage='" & psAgencia & "'
    
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    While Not lrDataRep.EOF
            Select Case lrDataRep!ctpoprecio
            
                Case 14
                        pnPreOro14 = lrDataRep!nPrecio
                Case 16
                        pnPreOro16 = lrDataRep!nPrecio
                Case 18
                        pnPreOro18 = lrDataRep!nPrecio
                Case 21
                        pnPreOro21 = lrDataRep!nPrecio
                Case 3005
                        pnTasaPreparacionRemate = lrDataRep!nPrecio
                Case 3004
                        pnTasaCustodiaVencida = lrDataRep!nPrecio
                Case 3003
                        pnTasaImpuesto = lrDataRep!nPrecio
                Case 3041
                        pnFactorPrecioVenta = lrDataRep!nPrecio
                
       'pnTasaCustodiaVencida As Double, pnTasaPreparacionRemate As Double, pnTasaImpuesto As Double, pnFactorPrecioVenta As Double
                
            End Select
            lrDataRep.MoveNext
    Wend
    Set lrDataRep = Nothing
    Set loDataRep = Nothing

    'ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, psFecRemate)

    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, "
    lsSQL = lsSQL & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote,   "
    lsSQL = lsSQL & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , "
    lsSQL = lsSQL & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , "
    lsSQL = lsSQL & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , "
    lsSQL = lsSQL & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , "
    lsSQL = lsSQL & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  "
    lsSQL = lsSQL & "               WHERE cLineaCred = C.cLineaCred  "
    lsSQL = lsSQL & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " )    "
    lsSQL = lsSQL & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod "
    lsSQL = lsSQL & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  "
    lsSQL = lsSQL & " WHERE p.cctacod='" & psCtaCod & "'"
    lsSQL = lsSQL & " GROUP BY P.cCtaCod, P.nPrdestado ,C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, "
    lsSQL = lsSQL & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote  "

'& " AND Substring(P.cCtaCod,4,2) = '" & psAgencia & "' AND DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  " _

    Set lrDataRep = New ADODB.Recordset
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If Not lrDataRep Is Nothing Then
                
        fnVarIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)

        Dim pnDeudaCosto As Double, pn15DeVenta As Double

        
        With lrDataRep
            If Not lrDataRep.EOF Then
                
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                vCuenta = 0
                
                pn15DeVenta = Round(((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * 0.15, 2)
              
                lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21))
                'lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * (fnVarIGV + 1)
                
                'lnPreBase = (v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)
                
                
                lnGramos = v14 + v16 + v18 + v21
                
                lnPreBase = Round(lnPreBase, 2)
                
                vdiasAtraso = DateDiff("d", !dVenc, psFecRemate)
                
                Set loCalculos = New NCOMColPCalculos
                    If vdiasAtraso <= 0 Then
                        vdiasAtraso = 0
                        vInteresVencido = 0
                        vCostoCustodiaMoratorio = 0
                        vImpuesto = 0
                    Else
                        vInteresVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vdiasAtraso)
                        vInteresVencido = Round(vInteresVencido, 2)
                        vCostoCustodiaMoratorio = loCalculos.nCalculaCostoCustodiaMoratorio(!nTasacion, pnTasaCustodiaVencida, vdiasAtraso)
                        vCostoCustodiaMoratorio = Round(vCostoCustodiaMoratorio, 2)
                        vImpuesto = (vInteresVencido + vCostoCustodiaMoratorio) * pnTasaImpuesto
                        vImpuesto = Round(vImpuesto, 2)
                    End If
                        vCostoPreparacionRemate = 0
                    'If !nPrdEstado = gColPEstPRema Then    ' Si esta en via de Remate
                        vCostoPreparacionRemate = pnTasaPreparacionRemate * !nTasacion
                        vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
                    'End If
                    lnDeuda = !nSaldo + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate
                    
                    pnSaldoDeuda = !nSaldo
                    pnIntVencido = vInteresVencido
                    
                    'pnSaldoDeuda pnIntVencido pnComision
                    
                   ' pnDeudaCosto = lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta
                     'pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta) * (fnVarIGV + 1), 2)
                    'Calcula el Precio Venta para Remate
                    'lnPreVenta = loCalculos.nCalculaPrecioVentaRemate(pnDeudaCosto, lnPreBase, pnFactorPrecioVenta)
                   ' lnPreVenta = Round(lnPreVenta, 2)
                   pnDeudaCosto = Round((lnDeuda + Round((lnDeuda * pnFactorPrecioVenta), 2) + pn15DeVenta), 2)
                   pnComision = Round((lnDeuda * pnFactorPrecioVenta), 2) + pn15DeVenta
                   
                   CalcDeudaRemateSINPROCESO = pnDeudaCosto
                                     
                   Set loCalculos = Nothing
                
                
            End If
        End With
            
    End If
    
End Function




'**CALCULA PRECIO VENTA SEGUN PRECIO ORO
Public Function CalcPrecioVta(ByVal pdFecSis As Date, ByVal psCtaCod As String) As Double
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NCOMColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
Dim pnPreOro14 As Double, pnPreOro16 As Double, pnPreOro21 As Double, pnPreOro18 As Double
Dim pnTasaCustodiaVencida As Double, pnTasaPreparacionRemate As Double, pnTasaImpuesto As Double, pnFactorPrecioVenta As Double

Dim lrCredPigJoyasDet As ADODB.Recordset
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim item As Integer, itemTMP As Integer
Dim lmDetalle(20) As String

pnPreOro14 = 0: pnPreOro16 = 0: pnPreOro21 = 0: pnPreOro18 = 0

    lsSQL = " SELECT nmaterial,nvalor FROM ColocPigPrecioMaterial "
    lsSQL = lsSQL & " WHERE nTipoValor = 1 AND nmaterial in (14,16,18,21) and nMoneda = '1' AND DATEDIFF(d, dInicial, '" & Format(pdFecSis, "mm/dd/yyyy") & "') > 0 "
    lsSQL = lsSQL & " ORDER BY dInicial Desc "
    
    
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    While Not lrDataRep.EOF
            Select Case lrDataRep!nmaterial
            
                Case 14
                        pnPreOro14 = lrDataRep!nValor
                Case 16
                        pnPreOro16 = lrDataRep!nValor
                Case 18
                        pnPreOro18 = lrDataRep!nValor
                Case 21
                        pnPreOro21 = lrDataRep!nValor
       
                
            End Select
            lrDataRep.MoveNext
    Wend
    Set lrDataRep = Nothing
    Set loDataRep = Nothing
    

    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, "
    lsSQL = lsSQL & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote,   "
    lsSQL = lsSQL & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , "
    lsSQL = lsSQL & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , "
    lsSQL = lsSQL & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , "
    lsSQL = lsSQL & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , "
    lsSQL = lsSQL & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  "
    lsSQL = lsSQL & "               WHERE cLineaCred = C.cLineaCred  "
    lsSQL = lsSQL & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " )    "
    lsSQL = lsSQL & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod "
    lsSQL = lsSQL & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  "
    lsSQL = lsSQL & " WHERE p.cctacod='" & psCtaCod & "'"
    lsSQL = lsSQL & " GROUP BY P.cCtaCod, P.nPrdestado ,C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, "
    lsSQL = lsSQL & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote  "

    Set lrDataRep = New ADODB.Recordset
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If Not lrDataRep Is Nothing Then
                
        fnVarIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)

        Dim pnDeudaCosto As Double, pn15DeVenta As Double

        
        With lrDataRep
            If Not lrDataRep.EOF Then
                
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                vCuenta = 0
                
                'pn15DeVenta = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * 0.15
              
                lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * (fnVarIGV + 1)
                
                
'                lnGramos = v14 + v16 + v18 + v21
'
'                lnPreBase = Round(lnPreBase, 2)
'
'                vdiasAtraso = DateDiff("d", !dVenc, psFecRemate)
'
'                Set loCalculos = New NCOMColPCalculos
'                    If vdiasAtraso <= 0 Then
'                        vdiasAtraso = 0
'                        vInteresVencido = 0
'                        vCostoCustodiaMoratorio = 0
'                        vImpuesto = 0
'                    Else
'                        vInteresVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vdiasAtraso)
'                        vInteresVencido = Round(vInteresVencido, 2)
'                        vCostoCustodiaMoratorio = loCalculos.nCalculaCostoCustodiaMoratorio(!nTasacion, pnTasaCustodiaVencida, vdiasAtraso)
'                        vCostoCustodiaMoratorio = Round(vCostoCustodiaMoratorio, 2)
'                        vImpuesto = (vInteresVencido + vCostoCustodiaMoratorio) * pnTasaImpuesto
'                        vImpuesto = Round(vImpuesto, 2)
'                    End If
'                        vCostoPreparacionRemate = 0
'                    'If !nPrdEstado = gColPEstPRema Then    ' Si esta en via de Remate
'                        vCostoPreparacionRemate = pnTasaPreparacionRemate * !nTasacion
'                        vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
'                    'End If
'                    lnDeuda = !nSaldo + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate
'
'
'                   ' pnDeudaCosto = lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta
'                     'pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta) * (fnVarIGV + 1), 2)
'                    'Calcula el Precio Venta para Remate
'                    'lnPreVenta = loCalculos.nCalculaPrecioVentaRemate(pnDeudaCosto, lnPreBase, pnFactorPrecioVenta)
'                   ' lnPreVenta = Round(lnPreVenta, 2)
'                   pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta), 2)
'
                   CalcPrecioVta = lnPreBase
                   
                   
                    Set loCalculos = Nothing
                
                
            End If
        End With
            
    End If
    
End Function

Public Function ValidaCredCancelado(ByVal psCtaCod As String) As Boolean
    Dim sql As String
    Dim Rs As ADODB.Recordset
    
    sql = "select count(cctacod) from Producto where nPrdEstado='2103' and cctacod= '" & Trim(psCtaCod) & "'"
    Set Rs = oCon.CargaRecordSet(sql)
    If Rs(0) > 0 Then
        ValidaCredCancelado = True
    Else
        ValidaCredCancelado = False
    End If
    Set Rs = Nothing
End Function

Private Sub Class_Initialize()
    Dim loIni As New COMConecta.DCOMClasIni
    Dim csConexion As String
    Dim csNegocio As String
    Dim csCentralPer As String
    Dim csCentralCom As String
    Dim csCentralImg As String
    Dim csAdminist As String

    csConexion = loIni.CadenaConexion
    csNegocio = loIni.BaseNegocio
    csCentralPer = loIni.BasePersonas
    csCentralCom = loIni.BaseComunes
    csCentralImg = loIni.BaseImagenes
    csAdminist = loIni.BaseAdministracion
    If oCon.AbreConexion(csConexion) = False Then
        Call oError.RaiseError(oError.MyUnhandledError, "DColPContrato:Initialize. Error en Conexion a Base de datos")
    End If
End Sub

Private Sub Class_Terminate()
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

'*** PEAC 20090312
Public Function nValidaPagoSobranteAdjudicadoCredPignoraticio(ByVal psCtaCod As String, ByVal psTipoProceso As String, Optional ByRef psmensaje As String) As ADODB.Recordset

Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset

    'lsSQL = " SELECT P.cCtaCod, P.nPrdestado, CRGDET.nRGDetEstado, CRGDET.nMontoProceso, CRGDET.nRemSubBaseVta,  " _
          & " P.nSaldo, CP.nTasacion, CRGDET.cNroProceso, CVta.nSobrante, CVta.nComision " _
          & " FROM Producto P Inner Join ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigRGDet CRGDET ON CP.cCtaCod = CRGDet.cCtaCod  " _
          & " LEFT JOIN ColocPigRecupVta CVta ON CRGDet.cCtaCod = CVta.cCtaCod  " _
          & " WHERE CRGDET.cTpoProceso ='" & psTipoProceso & "' " _
          & " AND CRGDET.cCtaCod ='" & psctacod & "' " _
          & " ORDER BY CRGDET.cNroProceso DESC "

    lsSQL = " exec stp_sel_ObtieneDatosDeCredSobranteAdjudicado '" & psCtaCod & "','" & psTipoProceso & "' "

    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        psmensaje = "AVISO: Contrato Pignoraticio No ha sido Adjudicado  "
        Exit Function
    End If
    If lrValida.BOF And lrValida.EOF Then
        psmensaje = "AVISO: Contrato Pignoraticio No ha sido Adjudicado  "
        Exit Function
    End If
    If lrValida!nRgDetEstado = geColPRecGarVentaSobranteCobrado Then
        psmensaje = " Ya fue cobrado el sobrante de Contrato " & psCtaCod
        Exit Function
'    ElseIf lrValida!nRgDetEstado = gColPRecGarVentaNoVendido Then
'        psmensaje = " Contrato No Fue Vendido en Remate " & psctacod
'        Exit Function
    End If
    
    Set nValidaPagoSobranteAdjudicadoCredPignoraticio = lrValida
    Set lrValida = Nothing
End Function

'*** PEAC 20090507
Public Function CargarDatosCtaAho(ByVal psCtaCod As String, ByVal pdFecSis As Date, _
                                    ByVal psPersCod As String, ByRef prsCuentas As ADODB.Recordset, _
                                    ByRef pvAgencias As Variant, ByRef pnContPersCtaAho As Integer, _
                                    ByRef pRsPersCtaAho As Variant, ByRef pnMontoITF As Double, ByVal psOperacion As String)

'Public Function CargarDatosCtaAho(ByVal psCtaCod As String, ByVal pdFecSis As Date, ByRef pnNroProxDesemb As Integer, ByRef prsDesemb As ADODB.Recordset, ByRef prsCredVig As ADODB.Recordset, ByRef pvTotalesVig As Variant,ByRef prsVigGra As ADODB.Recordset,ByRef prsCuentas As ADODB.Recordset, ByRef pvAgencias As Variant, ByRef pnContPersCtaAho As Integer, ByRef pRsPersCtaAho As Variant, ByRef prsGastos As ADODB.Recordset,  ByRef pnMontoITF As Double, ByRef prsDesemPar As ADODB.Recordset, ByRef pbAmpliacion As Boolean, ByVal psOperacion As String, ByRef pbOperacionEfectivo As Boolean, ByRef pnMontoLavDinero As Double, ByRef pnTC As Double, ByRef pbExoneradaLavado As Boolean, ByRef psPersCodRep As String, ByRef psPersNombreRep As String, ByRef prBancos As ADODB.Recordset) 'CUSCO

Dim oCredito As COMNColoCPig.NCOMColPContrato
Dim oCaptacion As COMNCaptaGenerales.NCOMCaptaGenerales
Dim oITF As COMDConstSistema.FCOMITF
Dim RTemp As ADODB.Recordset
        
Dim i As Integer
Dim nMontoA As Double
Dim nInteresA As Double
Dim nTotal As Double
        
Dim nNumCtasAhorro As Integer
Dim vCtaAhorro As Variant
Dim nMontoPrestamo As Double
Dim sPersCod As String
        
On Error GoTo ErrorCargarDatosCtaAho

    Set oCredito = New COMNColoCPig.NCOMColPContrato
        
        '''*******************************************************************
        'Carga Cuentas de Ahorro de Clientes
        Set oCaptacion = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set prsCuentas = oCaptacion.GetCuentasPersona(psPersCod, gCapAhorros, True, True, CInt(Mid(psCtaCod, 9, 1)))
        Set oCaptacion = Nothing
        
        '''*******************************************************************
        
        If Not (prsCuentas.EOF And prsCuentas.BOF) Then
        
        ReDim pvAgencias(prsCuentas.RecordCount)
        nNumCtasAhorro = prsCuentas.RecordCount
        ReDim vCtaAhorro(prsCuentas.RecordCount)
        Do While Not prsCuentas.EOF
            vCtaAhorro(prsCuentas.Bookmark) = prsCuentas!cCtaCod
            Set RTemp = oCredito.RecuperaAgencia(Mid(prsCuentas!cCtaCod, 4, 2))
            pvAgencias(prsCuentas.Bookmark) = Trim(RTemp!cAgeDescripcion)
            RTemp.Close
            prsCuentas.MoveNext
        Loop
        
        'Carga Clientes Relacionados con las Ctas de Ahorros
        pnContPersCtaAho = 0
        For i = 1 To nNumCtasAhorro
            pnContPersCtaAho = pnContPersCtaAho + 1
            
            ReDim Preserve pRsPersCtaAho(pnContPersCtaAho)
            
            Set oCaptacion = New COMNCaptaGenerales.NCOMCaptaGenerales
            Set pRsPersCtaAho(pnContPersCtaAho - 1) = oCaptacion.GetPersonaCuenta(vCtaAhorro(i))
            Set oCaptacion = Nothing
        Next i
            
        Set oCredito = Nothing
         
        '''*******************************************************************
    
    End If
    
    Exit Function

ErrorCargarDatosCtaAho:
    Err.Raise Err.Number, "Cargar Datos Cta. Aho", Err.Description
End Function

'MAVM 20100723 ***
Public Function nRescateJoyasGarant(ByVal psCtaCod As String, ByVal psCodAge As String, _
            ByVal pdFecSis As Date, ByVal psOperacion As Long, Optional ByVal cUser As String, _
            Optional ByVal pscJoyGarCod As String, Optional ByVal iTipo As Integer, Optional ByRef psmensaje As String) As ADODB.Recordset

Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset
Dim lnCustodiaDiferida As Currency
Dim loCalculos As NCOMColPCalculos
Dim lsOpeCancel As String
Set lrValida = New ADODB.Recordset
lsOpeCancel = " ( '" & gCredPagNorNorEfec & "') "

lsSQL = "SELECT J.cJoyGarCod, IsNull(P.cCtaCod, '')cCtaCod, IsNull(P.nPrdEstado, '') nPrdEstado" _
      & ", J.nTasacion, cAgeBoveda = SUBSTRING(J.cMovNro, 15, 5)" _
      & ", dCancelado = (SELECT top 1 CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ) FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro WHERE MC.cCtaCod = P.cCtaCod AND M.NMOVFLAG=0 AND MC.cOpeCod in  ( '100201')  order by M.nMovNro desc)" _
      & ", nDiasTranscurridos = IsNull(DateDiff(d, (SELECT top 1 CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ) FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro WHERE MC.cCtaCod = P.cCtaCod AND MC.cOpeCod in ( '100201')  AND M.nMovFlag <> 2 order by M.nMovNro desc ), '2010/11/22 '), '')" _
      & ", nCustodiaPag = ISNULL( (SELECT SUM(nMonto) FROM MovColDet MCD INNER JOIN Mov M ON M.nMovNro = MCD.nMovNro WHERE MCD.nMovNro = M.nMovNro AND M.NMOVFLAG=0 AND MCD.cOpeCod =  121900 AND MCD.cCtaCod = P.cCtaCod ) , 0 )" _
      & ", cBloqueo = (SELECT CASE WHEN COUNT(nBlqMotivo) > 0 THEN 'S' ELSE 'N' END FROM ProductoBloqueos PB LEFT JOIN Constante Cons ON PB.nBlqMotivo = Cons.nConsCod WHERE PB.cCtaCod = '109015051001516914' AND cMovNroDbl Is Null )" _
      & ", nMaxDiasCustodia = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = 3033 ) ,  nTasaIGV = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = 3007 )" _
      & ", nPorcentajeCustodia = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = 3034 )" _
      & ", nCustodiaDiferida = 0, nPiezas, J.nTasacion, nPesoNeto, nPesoBruto" _
      & " FROM JoyaGarantia J Left Join Garantias G on J.cJoyGarCod = G.cNumGarant and cTpoDoc = 145" _
      & " Left Join PersGarantia PG on PG.cPersCod = J.cPersCod and PG.nRelacion = 1 Left Join ProductoPersona PP on PG.cPersCod = PP.cPersCod and PP.nPrdPersRelac = 20 " _
      & " Left Join Producto P On PP.cCtaCod = P.cCtaCod and P.nPrdEstado = 2050"
If iTipo = 1 Then
    lsSQL = lsSQL & " WHERE P.cCtaCod = '" & psCtaCod & "' "
Else
    lsSQL = lsSQL & " WHERE J.cJoyGarCod = '" & pscJoyGarCod & "' "
End If
    lsSQL = lsSQL & " And J.nJoyGarEstado in(0, 2)"
    
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
        Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    Set loRegValida = Nothing
    If lrValida Is Nothing Then
        psmensaje = "No Existen Datos"
        Exit Function
    End If
    If lrValida.BOF And lrValida.EOF Then
        psmensaje = "No Existen Datos"
        Exit Function
    End If
    
    If psOperacion = gColPOpeDevJoyas Then ' Para Rescate Joya
        If lrValida!nPrdEstado <> gColocEstCancelado Then
            psmensaje = "Contrato se encuentra en estado " & mfgEstadoCredPigDesc(lrValida!nPrdEstado)
            Exit Function
        End If
    ElseIf psOperacion = gColPOpeDevJoyasNoDesemb Then  ' Para Rescate Joya No Desembolsada
        If lrValida!nPrdEstado <> gColPEstAnula Then
            psmensaje = "Contrato se encuentra en estado " & mfgEstadoCredPigDesc(lrValida!nPrdEstado)
            Exit Function
        End If
    End If
    
    ' Si esta bloqueada
'    If lrValida!cBloqueo = "S" Then
'        psmensaje = " Contrato se encuentra Bloqueado "
'        Exit Function
'    End If
    
    ' Ubicacion de la Joya
    If iTipo = 1 Then
        If Trim(lrValida!cAgeBovedA) <> Mid(psCtaCod, 1, 3) & psCodAge Then
            psmensaje = " Garantía Prendaria se encuentra en Boveda de Agencia " & lrValida!cAgeBovedA
            Exit Function
        End If
    End If
    
    'Calcula el Costo de Custodia Extra
    If iTipo = 1 Then
        If DateDiff("d", lrValida!dCancelado, pdFecSis) > lrValida!nMaxDiasCustodia Then
            Set loCalculos = New NCOMColPCalculos
                lnCustodiaDiferida = loCalculos.nCalculaCostoCustodiaDiferida(lrValida!nTasacion, lrValida!nDiasTranscurridos, lrValida!nPorcentajeCustodia, lrValida!nTasaIGV)
            Set loCalculos = Nothing
            If lnCustodiaDiferida - lrValida!nCustodiaPag > 0 Then
                psmensaje = "Contrato tiene Custodia Diferida, cobrar primero la custodia "
                Exit Function
            End If
        Else
            lnCustodiaDiferida = 0
        End If
    End If
    Set nRescateJoyasGarant = lrValida
End Function

'*** PEAC 20180323
Public Function nCobroPrecioVentaLoteAdjudicado(ByVal psCtaCod As String, ByVal pnMontoDeuda As Double, ByVal pcFecSis As String, ByRef psmensaje As String) As ADODB.Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset
Dim lrValorVenta As ADODB.Recordset

    lsSQL = " exec stp_sel_ObtieneDatosParaVentaDeContratoAdjudicado '" & psCtaCod & "'," & pnMontoDeuda & ",'" & pcFecSis & "' "

    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValorVenta = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValorVenta.BOF And lrValorVenta.EOF Then
        psmensaje = "El Credito no se encuetra en el Proceso de Subasta "
        Exit Function
    End If
    
    Set nCobroPrecioVentaLoteAdjudicado = lrValorVenta
    Set lrValorVenta = Nothing
End Function

'*** PEAC 20180905
Public Function BuscaDevCredPignoPorAbonar(ByVal psCtaCod As String, ByRef psmensaje As String) As ADODB.Recordset

Dim lsSQL As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As ADODB.Recordset
Dim lrValorVenta As ADODB.Recordset

    lsSQL = " exec stp_sel_BuscaDevCredPorAbonar '" & psCtaCod & "'"

    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValorVenta = loRegValida.dObtieneRecordSet(lsSQL)
    If Not (lrValorVenta.BOF And lrValorVenta.EOF) Then
        psmensaje = "Este Crédito tiene Pendiente realizar la Devolución de Sobrante de Adjudicado, contáctese con el Coordinador de Operaciones."
        Exit Function
    End If
    
    Set BuscaDevCredPignoPorAbonar = lrValorVenta
    Set lrValorVenta = Nothing
End Function
'*** JAOR - 20210224****************************************************************************************
Public Function BuscaCredPignoCampana(ByVal psCtaCod As String, ByRef pnEstado As Integer) As ADODB.Recordset

Dim lsSQL As String
Dim rsBuscaPigCamp As COMDColocPig.DCOMColPFunciones

Dim rsResult As ADODB.Recordset

    lsSQL = " exec stp_sel_BuscaCredPignoCampana '" & psCtaCod & "'"

    Set rsBuscaPigCamp = New COMDColocPig.DCOMColPFunciones
    Set rsResult = rsBuscaPigCamp.dObtieneRecordSet(lsSQL)
    If Not (rsResult.BOF And rsResult.EOF) Then
        pnEstado = CInt(Abs(rsResult!bEstado))
        'Exit Function
    Else
         pnEstado = 0
    End If
    
    Set BuscaCredPignoCampana = rsResult
    Set rsResult = Nothing

End Function
'***********************************************************************************************************

'CTI4 ERS0112020
Public Function EsMismoTitulardeCuentaPignoYAhorro(ByVal psCtaCodAhorro As String, ByVal psCtaCodPigno As String) As Boolean
    Dim lsSQL As String
    Dim loRegValida As COMDColocPig.DCOMColPFunciones
    Dim lrValida As ADODB.Recordset
    
    lsSQL = " exec stp_sel_EsMismoTitulardeCuentaPignoYAhorro '" & psCtaCodAhorro & "','" & psCtaCodPigno & "'"
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If Not (lrValida.EOF And lrValida.BOF) Then
        EsMismoTitulardeCuentaPignoYAhorro = IIf(lrValida!bexito, True, False)
        Exit Function
    End If
    EsMismoTitulardeCuentaPignoYAhorro = False
End Function
'end CTI4

