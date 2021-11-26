VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nColPValida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************
'* Clase que contiene las validaciones de las operaciones de Credito Pignoraticio
'* LAYG - 01/07/2001
'****************************
Option Explicit
Dim oCon As Connection
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String
Public Function nValidaAnulacionCredPignoraticio(ByVal psCtacod As String, ByVal pdFecsis As Date, Optional ByVal cUser As String) As Recordset
Dim lsSQL As String
'Dim loRegValida As dColPFunciones

Dim lrValida As ADODB.Recordset

Set lrValida = New ADODB.Recordset


    lsSQL = " select m.nmovnro,m.cmovnro,mc.copecod, mcd.nprdconceptocod, montodet=mcd.nmonto,montot=mc.nmonto "
    lsSQL = lsSQL & " into #tmpPrend" & cUser & " from mov m"
    lsSQL = lsSQL & " join movcol mc on mc.nmovnro=m.nmovnro "
    lsSQL = lsSQL & " join movcoldet mcd on mcd.nmovnro=mc.nmovnro and mcd.copecod=mc.copecod "
    lsSQL = lsSQL & " where mc.cctacod='" & psCtacod & "' and m.nmovflag=0 and m.nmovestado=10 "
    oCon.Execute lsSQL
 
    lsSQL = " SELECT P.cCtaCod, P.nPrdestado,  " _
          & " dRegistrado = (SELECT top 1 CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
          & "              FROM #tmpPrend" & cUser & "" _
          & "              WHERE cOpeCod = " & gColPOpeRegContrato & " )" _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " Inner Join ColocPignoraticio CP on C.cctacod = CP.cctacod " _
          & " WHERE P.cCtaCod ='" & psCtacod & "' "
          
    lrValida.CursorLocation = adUseClient
    lrValida.Open lsSQL, oCon, adOpenStatic, adLockReadOnly, adCmdText
    Set lrValida.ActiveConnection = Nothing
         
          
    lsSQL = " drop table #tmpPrend" & cUser
    oCon.Execute lsSQL

'    lsSQL = " SELECT P.cCtaCod, P.nPrdestado,  " _
'          & " dRegistrado = (SELECT CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
'          & "              FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
'          & "              WHERE MC.cCtaCod = C.cCtaCod AND m.nmovflag=0 and MC.cOpeCod in ( " & gColPOpeRegContrato & ")  )" _
'          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
'          & " Inner Join ColocPignoraticio CP on C.cctacod = CP.cctacod " _
'          & " WHERE P.cCtaCod ='" & psCtaCod & "' "
'
'    Set loRegValida = New dColPFunciones
'    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    
    If lrValida Is Nothing Or (lrValida.BOF And lrValida.EOF) Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
  
    If lrValida!nPrdEstado <> gColPEstRegis Then
        MsgBox "Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
        Exit Function
    End If
    If DateDiff("d", lrValida!dRegistrado, pdFecsis) <> 0 Then
        MsgBox " Contrato se encuentra fuera de su Fecha Limite ", vbInformation, " Aviso "
        Exit Function
    End If
    
    Set nValidaAnulacionCredPignoraticio = lrValida
    Set lrValida = Nothing
End Function

Public Function nValidaRescateCredPignoraticio(ByVal psCtacod As String, ByVal psCodAge As String, _
            ByVal pdFecsis As Date, ByVal psOperacion As Long, Optional ByVal cUser As String) As Recordset

'************************************
' ************************************
Dim lsSQL As String
'Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset
Dim lnCustodiaDiferida As Currency
Dim loCalculos As NColPCalculos
Dim lsOpeCancel As String
  Set lrValida = New ADODB.Recordset
lsSQL = " select m.nmovnro,m.cmovnro,mc.copecod, mcd.nprdconceptocod, montodet=mcd.nmonto,montot=mc.nmonto "
lsSQL = lsSQL & " into #tmpPrend" & cUser & " from mov m"
lsSQL = lsSQL & " join movcol mc on mc.nmovnro=m.nmovnro "
lsSQL = lsSQL & " join movcoldet mcd on mcd.nmovnro=mc.nmovnro and mcd.copecod=mc.copecod "
lsSQL = lsSQL & " where mc.cctacod='" & psCtacod & "' and m.nmovflag=0 and m.nmovestado=10 "
oCon.Execute lsSQL

lsOpeCancel = " ( '" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" & gColPOpeCancelNorCHQ & "','" & gColPOpeCancelMorCHQ & "','" & gColPOpeCanceNorEnOtCjEFE & "','" & gColPOpeCanceMorEnOtCjEFE & "') "

lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, CP.cAgeBoveda , CP.nNroDuplic, " _
      & " dCancelado = (SELECT top 1 CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
      & "              FROM #tmpPrend" & cUser & "" _
      & "              WHERE cOpeCod in " & lsOpeCancel & " ) ,  " _
      & " nDiasTranscurridos = DateDiff ( d, (SELECT top 1 CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
      & "              FROM #tmpPrend" & cUser & "" _
      & "              WHERE cOpeCod in " & lsOpeCancel & " ), '" & Format(pdFecsis, "yyyy/mm/dd") & " ') , " _
      & " nCustodiaPag = ISNULL( (SELECT SUM( montodet) FROM #tmpPrend" & cUser & "" _
      & "                       WHERE cOpeCod =  " & geColPCobCusDiferida & "  ) , 0 ) , " _
      & " cBloqueo = (SELECT CASE WHEN COUNT(nBlqMotivo) > 0 THEN 'S' ELSE 'N' END FROM ProductoBloqueos PB LEFT JOIN Constante Cons " _
      & "             ON PB.nBlqMotivo = Cons.nConsCod WHERE PB.cCtaCod = '" & psCtacod & "' AND cMovNroDbl Is Null ) , " _
      & " nMaxDiasCustodia = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & gConsColPMaxDiasCustodiaDiferida & " ) , " _
      & " nTasaIGV = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & gConsColPTasaIGV & " ) , " _
      & " nPorcentajeCustodia = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & gConsColPPorcentajeCustodiaDiferida & " ) ," _
      & " nCustodiaDiferida = 0 FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = c.cCtaCod " _
      & " INNER JOIN ColocPignoraticio CP ON c.cCtaCod = cp.cCtaCod " _
      & " WHERE P.cCtaCod = '" & psCtacod & "' "
      
     lrValida.CursorLocation = adUseClient
    lrValida.Open lsSQL, oCon, adOpenStatic, adLockReadOnly, adCmdText
    Set lrValida.ActiveConnection = Nothing
    
 

 lsSQL = " drop table #tmpPrend" & cUser
 oCon.Execute lsSQL

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
'      & "             ON PB.nBlqMotivo = Cons.nConsCod WHERE PB.cCtaCod = '" & psCtaCod & "' AND cMovNroDbl Is Null ) , " _
'      & " nMaxDiasCustodia = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & gConsColPMaxDiasCustodiaDiferida & " ) , " _
'      & " nTasaIGV = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & gConsColPTasaIGV & " ) , " _
'      & " nPorcentajeCustodia = (SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & gConsColPPorcentajeCustodiaDiferida & " ) ," _
'      & " nCustodiaDiferida = 0 FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = c.cCtaCod " _
'      & " INNER JOIN ColocPignoraticio CP ON c.cCtaCod = cp.cCtaCod " _
'      & " WHERE P.cCtaCod = '" & psCtaCod & "' "

'
'    Set loRegValida = New dColPFunciones
'        Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
'    Set loRegValida = Nothing
    If lrValida Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrValida.BOF And lrValida.EOF Then
        MsgBox "ERROR: al Buscar datos del credito ", vbInformation, "Aviso"
        Exit Function
    End If
    
    If psOperacion = gColPOpeDevJoyas Then ' Para Rescate Joya
        If lrValida!nPrdEstado <> gColPEstDifer Then
            MsgBox "Contrato se encuentra en estado " & mfgEstadoCredPigDesc(lrValida!nPrdEstado), vbInformation, "Aviso"
            Exit Function
        End If
    ElseIf psOperacion = gColPOpeDevJoyasNoDesemb Then  ' Para Rescate Joya No Desembolsada
        If lrValida!nPrdEstado <> gColPEstAnula Then
            MsgBox "Contrato se encuentra en estado " & mfgEstadoCredPigDesc(lrValida!nPrdEstado), vbInformation, "Aviso"
            Exit Function
        End If
    End If
    ' Si esta bloqueada
    If lrValida!cBloqueo = "S" Then
        MsgBox " Contrato se encuentra Bloqueado ", vbInformation, " Aviso "
        Exit Function
    End If
    ' Ubicacion de la Joya
    If Trim(lrValida!cAgeBoveda) <> Mid(psCtacod, 1, 3) & psCodAge Then
        MsgBox " Garantía Prendaria se encuentra en Boveda de Agencia " & lrValida!cAgeBoveda, vbInformation, " Aviso "
        Exit Function
    End If
    
    'Calcula el Costo de Custodia Extra
    If DateDiff("d", lrValida!dCancelado, pdFecsis) > lrValida!nMaxDiasCustodia Then
        Set loCalculos = New NColPCalculos
            lnCustodiaDiferida = loCalculos.nCalculaCostoCustodiaDiferida(lrValida!nTasacion, lrValida!nDiasTranscurridos, lrValida!nPorcentajeCustodia, lrValida!nTasaIGV)
        Set loCalculos = Nothing
        If lnCustodiaDiferida - lrValida!nCustodiaPag > 0 Then
            MsgBox "Contrato tiene Custodia Diferida, cobrar primero la custodia ", vbInformation, "Aviso"
            Exit Function
        End If
    Else
        lnCustodiaDiferida = 0
    End If
    
    Set nValidaRescateCredPignoraticio = lrValida
End Function


Public Function nValidaDesembolsoCredPignoraticio(ByVal psCtacod As String, ByVal pdFecsis As Date, _
     ByVal pnMaxDiasDesembolso As Integer) As Recordset

'************************************

Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, dVigencia, " _
         & " cBloqueo = (SELECT CASE WHEN COUNT(nBlqMotivo) > 0 THEN 'S' ELSE 'N' END FROM ProductoBloqueos PB LEFT JOIN Constante Cons " _
         & "             ON PB.nBlqMotivo = Cons.nConsCod WHERE PB.cCtaCod = '" & psCtacod & "' AND cMovNroDbl Is Null ) , " _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodInteresCompensatorio & "   then nMonto else 0 end )  as nInteres," _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodImpuesto & "  then nMonto else 0 end )  as nImpuesto," _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodTasacion & "  then nMonto else 0 end )  as nTasacion," _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodCustodia & "  then nMonto else 0 end )  as nCustodia " _
         & " FROM Colocaciones c Inner join Producto P on P.cCtaCod = c.cCtaCod " _
         & " Inner join ColocCalendario CCal on CCal.cCtaCod = C.cCtaCod " _
         & " Inner join ColoccalendDetPig CCalD on CCal.cCtaCod = CCalD.cCtaCod and CCal.nNroCalen = CCalD.nNroCalen " _
         & " And CCal.nColocCAlendApl = CCalD.nColocCalendApl " _
         & " Where P.cCtaCod  = '" & psCtacod & "' AND CCalD.nColocCalendApl =  " & gColocCalendAplDesembolso & " " _
         & " GROUP BY P.cCtaCod, P.nPrdEstado, dVigencia "
    
    'Set lrValida = New ADODB.Recordset
    Set loRegValida = New dColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, " Aviso"
        Exit Function
    End If
    If lrValida.BOF And lrValida.EOF Then
        MsgBox " No se ha encontrado Contrato ", vbInformation, " Aviso "
        Exit Function
    End If
    If lrValida!nPrdEstado <> gColPEstRegis Then
        MsgBox " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
        Exit Function
    ElseIf DateDiff("d", lrValida!dVigencia, pdFecsis) > pnMaxDiasDesembolso Then
        MsgBox " Contrato está fuera de la Fecha para su Desembolso ", vbInformation, " Aviso "
        Exit Function
    End If
    If lrValida!cBloqueo = "S" Then
        MsgBox " Contrato se encuentra Bloqueado ", vbInformation, " Aviso "
        Exit Function
    End If
    Set nValidaDesembolsoCredPignoraticio = lrValida
    Set lrValida = Nothing
End Function


Public Function nValidaRenovacionCredPignoraticio(ByVal psCtacod As String, ByVal pdFecsis As Date, _
     ByVal pnMaxDiasDesembolso As Integer, Optional ByVal cUser As String) As Recordset

'************************************
Dim lsSQL As String
'Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset
Set lrValida = New ADODB.Recordset
    
    lsSQL = " select m.nmovnro,m.cmovnro,mc.copecod, mcd.nprdconceptocod, montodet=mcd.nmonto,montot=mc.nmonto "
    lsSQL = lsSQL & " into #tmpPrend" & cUser & " from mov m"
    lsSQL = lsSQL & " join movcol mc on mc.nmovnro=m.nmovnro "
    lsSQL = lsSQL & " join movcoldet mcd on mcd.nmovnro=mc.nmovnro and mcd.copecod=mc.copecod "
    lsSQL = lsSQL & " where mc.cctacod='" & psCtacod & "' and m.nmovflag=0 and m.nmovestado=10 "
    oCon.Execute lsSQL
 
    lsSQL = " SELECT codigoant=isnull(r.cctacodant,''),P.cCtaCod, P.nPrdestado, C.dVenc, C.nPlazo, CP.nNroRenov,  " _
          & " P.nSaldo, CP.nTasacion, P.nTasaInteres, " _
          & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ),  " _
          & " dDesembolso = (SELECT MAX (CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
          & "              FROM #tmpPrend" & cUser & " " _
          & "              WHERE  cOpeCod= " & gColPOpeDesembolsoEFE & " ),  " _
          & " dRenovacion = (SELECT MAX (CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
          & "              FROM #tmpPrend" & cUser & " " _
          & "              WHERE cOpeCod  in ( " & gColPOpeRenovNorEFE & "," & gColPOpeRenovMorEFE & "," _
          & gColPOpeRenovNorEnOtCjEFE & "," & gColPOpeRenovMorEnOtCjEFE & ")  ),  " _
          & " cInicRemate = (SELECT CASE WHEN COUNT(RG.nRGEstado)> 0  THEN 'S' ELSE 'N' END " _
          & "                FROM ColocPigRecGar RG Join ColocPigRGDet RGD ON RG.cTpoProceso = RGD.cTpoProceso And RG.cNroProceso = RGD.cNroProceso  " _
          & "                WHERE RG.cTpoProceso ='R' AND DATEDIFF(D, RG.dProceso,'" & Format(pdFecsis, "mm/dd/yyyy") & "') = 0 " _
          & "                AND RG.nRGEStado ='" & gColPRecGarEstIniciado & "' AND RGD.cCtaCod = C.cCtaCod AND RG.cCodAge = Substring(C.cCtaCod,1,5)  )" _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " Inner Join ColocPignoraticio CP on C.cctacod = CP.cctacod " _
          & " Left join Relcuentas r on r.cctacod=p.cctacod  " _
          & " WHERE P.cCtaCod ='" & psCtacod & "'"
          
    lrValida.CursorLocation = adUseClient
    lrValida.Open lsSQL, oCon, adOpenStatic, adLockReadOnly, adCmdText
    Set lrValida.ActiveConnection = Nothing
    
    lsSQL = " drop table #tmpPrend" & cUser
    oCon.Execute lsSQL
          
'    lsSQL = lsSQL & " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, C.nPlazo, CP.nNroRenov,  " _
'          & " P.nSaldo, CP.nTasacion, P.nTasaInteres, " _
'          & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
'          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ),  " _
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
'          & "                AND RG.nRGEStado ='" & gColPRecGarEstIniciado & "' AND RGD.cCtaCod = C.cCtaCod AND RG.cCodAge = Substring(C.cCtaCod,1,5)  )" _
'          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
'          & " Inner Join ColocPignoraticio CP on C.cctacod = CP.cctacod " _
'          & " WHERE P.cCtaCod ='" & psCtaCod & "' "

    
'    Set loRegValida = New dColPFunciones
'    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)

    If lrValida Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrValida.BOF And lrValida.EOF Then
        MsgBox "ERROR: No se encontro el Credito ", vbInformation, "Aviso"
        Exit Function
    End If
    If (lrValida!nPrdEstado <> gColPEstDesem) And (lrValida!nPrdEstado <> gColPEstVenci) And _
       (lrValida!nPrdEstado <> gColPEstPRema) And (lrValida!nPrdEstado <> gColPEstRenov) Then
        MsgBox " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
        Exit Function
    End If
    ' Si ha desembolsado el mismo dia
    If DateDiff("d", pdFecsis, lrValida!dDesembolso) = 0 Then
        MsgBox "Contrato ha sido Desembolsado el dia de hoy día ", vbInformation, " Aviso "
            Exit Function
    End If
    ' Si ha renovado el mismo dia
    If Not IsNull(lrValida!dRenovacion) And DateDiff("d", pdFecsis, lrValida!dRenovacion) = 0 Then
       If MsgBox("Contrato ha sido Renovado el dia de hoy día " & vbCr & "Desea continuar ? ", vbQuestion + vbYesNo, " Aviso ") = vbNo Then
            Exit Function
       End If
    End If
       
    ' Si No se ha definido la tasa de Interes Moratoria
    If IsNull(lrValida!nTasaIntVenc) Then
        MsgBox " No se ha definido la Tasa de Interes Vencida para linea de Credito ", vbInformation, " Aviso "
        Exit Function
    End If
        
    ' Si el Proceso Remate se ha iniciado ' LAYG (21/07/2003)
    If lrValida!cInicRemate = "S" Then
        MsgBox " Contrato se encuentra EN EL PROCESO DE REMATE, no se puede Renovar ", vbInformation, " Aviso "
        Exit Function
    End If
    
    Set nValidaRenovacionCredPignoraticio = lrValida
    Set lrValida = Nothing
    
End Function


Public Function nValidaCancelacionCredPignoraticio(ByVal psCtacod As String, ByVal pdFecsis As Date, _
     ByVal pnMaxDiasDesembolso As Integer) As Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset


        
    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, C.nPlazo, CP.nNroRenov,  " _
          & " P.nSaldo, CP.nTasacion, P.nTasaInteres, " _
          & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ) ,  " _
          & " cInicRemate = (SELECT CASE WHEN COUNT(RG.nRGEstado)> 0  THEN 'S' ELSE 'N' END " _
          & "                FROM ColocPigRecGar RG Join ColocPigRGDet RGD ON RG.cTpoProceso = RGD.cTpoProceso And RG.cNroProceso = RGD.cNroProceso " _
          & "                WHERE RG.cTpoProceso ='R' AND DATEDIFF(D, RG.dProceso,'" & Format(pdFecsis, "mm/dd/yyyy") & "') = 0  " _
          & "                AND RG.nRGEStado ='" & gColPRecGarEstIniciado & "' AND RGD.cCtaCod = C.cCtaCod AND RG.cCodAge = Substring(C.cCtaCod,1,5)  ) " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " Inner Join ColocPignoraticio CP on C.cctacod = CP.cctacod " _
          & " WHERE P.cCtaCod ='" & psCtacod & "' "


    
    Set loRegValida = New dColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    
    If (lrValida!nPrdEstado <> gColPEstDesem) And (lrValida!nPrdEstado <> gColPEstVenci) And _
       (lrValida!nPrdEstado <> gColPEstPRema) And (lrValida!nPrdEstado <> gColPEstRenov) Then
        MsgBox " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
        Exit Function
    End If
    ' Si el Proceso Remate se ha iniciado ' LAYG (21/07/2003)
    If lrValida!cInicRemate = "S" Then
        MsgBox " Contrato se encuentra EN EL PROCESO DE REMATE, no se puede Cancelar ", vbInformation, " Aviso "
        Exit Function
    End If
    
    Set nValidaCancelacionCredPignoraticio = lrValida
    Set lrValida = Nothing
End Function


Public Function nValidaCustodiaDiferidaCredPignoraticio(ByVal psCtacod As String, ByVal pdFecsis As Date, _
     ByVal pnMaxDiasDesembolso As Integer, Optional ByVal cUser As String) As Recordset

'************************************

Dim lsSQL As String
'Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset
Set lrValida = New ADODB.Recordset

    lsSQL = " select m.nmovnro,m.cmovnro,mc.copecod, mcd.nprdconceptocod, montodet=mcd.nmonto,montot=mc.nmonto "
    lsSQL = lsSQL & " into #tmpPrend" & cUser & " from mov m"
    lsSQL = lsSQL & " join movcol mc on mc.nmovnro=m.nmovnro "
    lsSQL = lsSQL & " join movcoldet mcd on mcd.nmovnro=mc.nmovnro and mcd.copecod=mc.copecod "
    lsSQL = lsSQL & " where mc.cctacod='" & psCtacod & "' and m.nmovflag=0 and m.nmovestado=10 "
    oCon.Execute lsSQL
 

   lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, C.nPlazo, cp.nTasacion , " _
          & " dCancelado = (SELECT top 1 CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
          & "              FROM #tmpPrend" & cUser & " " _
          & "              WHERE cOpeCod  in ( '" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" & gColPOpeCanceNorEnOtCjEFE & "','" & gColPOpeCanceMorEnOtCjEFE & "')  ),  " _
          & " nCustodiaPag = ISNULL( (SELECT SUM(montodet) FROM #tmpPrend" & cUser & " " _
          & "                       WHERE cOpeCod = " & geColPCobCusDiferida & "  ) , 0 ) " _
          & " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = c.cCtaCod " _
          & " INNER JOIN ColocPignoraticio CP ON c.cCtaCod = cp.cCtaCod " _
          & " WHERE P.cCtaCod = '" & psCtacod & "' "
    lrValida.CursorLocation = adUseClient
    lrValida.Open lsSQL, oCon, adOpenStatic, adLockReadOnly, adCmdText
    Set lrValida.ActiveConnection = Nothing
    
   lsSQL = " drop table #tmpPrend" & cUser
   oCon.Execute lsSQL

'    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, C.nPlazo, cp.nTasacion , " _
'          & " dCancelado = (SELECT CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )  " _
'          & "              FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
'          & "              WHERE MC.cCtaCod = C.cCtaCod AND MC.cOpeCod  in ( '" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" & gColPOpeCanceNorEnOtCjEFE & "','" & gColPOpeCanceMorEnOtCjEFE & "') AND M.nMovFlag <> " & gMovFlagExtornado & " ),  " _
'          & " nCustodiaPag = ISNULL( (SELECT SUM(nMonto) FROM MovColDet MCD " _
'          & "                       INNER JOIN Mov M ON M.nMovNro = MCD.nMovNro " _
'          & "                       WHERE MCD.nMovNro = M.nMovNro  AND m.nmovflag=0  AND MCD.cOpeCod = " & geColPCobCusDiferida & " AND MCD.cCtaCod = P.cCtaCod ) , 0 ) " _
'          & " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = c.cCtaCod " _
'          & " INNER JOIN ColocPignoraticio CP ON c.cCtaCod = cp.cCtaCod " _
'          & " WHERE P.cCtaCod = '" & psCtaCod & "' "
    
'    Set loRegValida = New dColPFunciones
'    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    
    If (lrValida!nPrdEstado <> gColPEstDifer) Then
        MsgBox " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
        Exit Function
    End If
    
    Set nValidaCustodiaDiferidaCredPignoraticio = lrValida
    Set lrValida = Nothing
    
End Function



Public Function nValidaDuplicadoContratoCredPignoraticio(ByVal psCtacod As String) As Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, dVigencia, P.nTasaInteres, CP.nNroDuplic, " _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodInteresCompensatorio & "   then nMonto else 0 end )  as nInteres," _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodImpuesto & "  then nMonto else 0 end )  as nImpuesto," _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodTasacion & "  then nMonto else 0 end )  as nTasacion," _
         & " SUM (CASE WHEN nPrdConceptoCod = " & gColPConceptoCodCustodia & "  then nMonto else 0 end )  as nCustodia " _
         & " FROM Colocaciones c Inner join Producto P on P.cCtaCod = c.cCtaCod " _
         & " Inner Join ColocPignoraticio CP On C.cCtaCod = CP.cCtaCod Inner join ColocCalendario CCal on CCal.cCtaCod = C.cCtaCod " _
         & " Inner join ColoccalendDetPig CCalD on CCal.cCtaCod = CCalD.cCtaCod and CCal.nNroCalen = CCalD.nNroCalen " _
         & " And CCal.nColocCalendApl = CCalD.nColocCalendApl " _
         & " Where P.cCtaCod  = '" & psCtacod & "' AND CCalD.nColocCalendApl =  " & gColocCalendAplDesembolso & " " _
         & " GROUP BY P.cCtaCod, P.nPrdEstado, dVigencia, P.nTasaInteres, CP.nNroDuplic "
    
    'Set lrValida = New ADODB.Recordset
    Set loRegValida = New dColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrValida.BOF And lrValida.EOF Then
        MsgBox "ERROR: al Buscar datos de Contrato ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrValida!nPrdEstado <> gColPEstRegis And lrValida!nPrdEstado <> gColPEstDesem _
        And lrValida!nPrdEstado <> gColPEstDifer And lrValida!nPrdEstado <> gColPEstCance _
        And lrValida!nPrdEstado <> gColPEstPRema And lrValida!nPrdEstado <> gColPEstRenov And lrValida!nPrdEstado <> gColPEstVenci Then
        MsgBox " Contrato se encuentra " '& ContratoEstado(lrvalida!ncEstado), vbInformation, " Aviso "
        Exit Function
    End If
    Set nValidaDuplicadoContratoCredPignoraticio = lrValida
    Set lrValida = Nothing
End Function

Public Function nValidaRegVentaRemateCredPignoraticio(ByVal psCtacod As String, ByVal psTipoProceso As String, _
    ByVal psNroProceso As String) As Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, CRGDET.nRGDetEstado, CRGDET.nRemSubBaseVta,  " _
          & " P.nSaldo, CP.nTasacion , C.dVenc, " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ) " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigRGDet CRGDET ON CP.cCtaCod = CRGDet.cCtaCod " _
          & " WHERE CRGDET.cTpoProceso ='" & psTipoProceso & "' AND CRGDET.cNroProceso ='" & psNroProceso & "' " _
          & " AND CRGDET.cCtaCod ='" & psCtacod & "' "

    Set loRegValida = New dColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida.BOF And lrValida.EOF Then
        MsgBox "El Credito no se encuetra en el Proceso de Remate ", vbInformation, "Aviso"
        Exit Function
    End If
    
    If (lrValida!nPrdEstado <> gColPEstRemat) And (lrValida!nPrdEstado <> gColPEstPRema) And _
       (lrValida!nPrdEstado <> gColPEstVenci) And (lrValida!nPrdEstado <> gColPEstRenov) Then
        MsgBox " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
        Exit Function
    End If
    If lrValida!nPrdEstado = gColPEstRemat And lrValida!nRgDetEstado = gColPRecGarVentaVendido Then
        MsgBox " Contrato ya fue rematado " & psCtacod, vbInformation, " Aviso "
        Exit Function
    End If
    
    Set nValidaRegVentaRemateCredPignoraticio = lrValida
    Set lrValida = Nothing
End Function


Public Function nValidaPagoSobranteRemateCredPignoraticio(ByVal psCtacod As String, ByVal psTipoProceso As String) As Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, CRGDET.nRGDetEstado, CRGDET.nMontoProceso, CRGDET.nRemSubBaseVta,  " _
          & " P.nSaldo, CP.nTasacion, CRGDET.cNroProceso, CVta.nSobrante, CVta.nComision " _
          & " FROM Producto P Inner Join ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigRGDet CRGDET ON CP.cCtaCod = CRGDet.cCtaCod  " _
          & " LEFT JOIN ColocPigRecupVta CVta ON CRGDet.cCtaCod = CVta.cCtaCod  " _
          & " WHERE CRGDET.cTpoProceso ='" & psTipoProceso & "' " _
          & " AND CRGDET.cCtaCod ='" & psCtacod & "' " _
          & " ORDER BY CRGDET.cNroProceso DESC "

    Set loRegValida = New dColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        MsgBox "AVISO: Contrato Pignoraticio No ha sido Rematado  ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrValida.BOF And lrValida.EOF Then
        MsgBox "AVISO: Contrato Pignoraticio No ha sido Rematado  ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrValida!nRgDetEstado = geColPRecGarVentaSobranteCobrado Then
        MsgBox " Ya fue cobrado el sobrante de Contrato " & psCtacod, vbInformation, " Aviso "
        Exit Function
    ElseIf lrValida!nRgDetEstado = gColPRecGarVentaNoVendido Then
        MsgBox " Contrato No Fue Vendido en Remate " & psCtacod, vbInformation, " Aviso "
        Exit Function
    End If
    
    Set nValidaPagoSobranteRemateCredPignoraticio = lrValida
    Set lrValida = Nothing
End Function


Public Function nValidaRegVentaSubastaCredPignoraticio(ByVal psCtacod As String, ByVal psTipoProceso As String, _
    ByVal psNroProceso As String) As Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, CRGDET.nRGDetEstado,  " _
          & " P.nSaldo, CP.nTasacion, CRGDET.nRemSubBaseVta,c.dprdestado,  " _
          & " nValRegistroAdj = ISNULL( (Select nAdjValRegistro FROM ColocPigRGDET  " _
          & "                            WHERE cTpoProceso = 'A' AND cCtaCod = P.cCtaCod ) , 0 ),bExcepVta=isnull(cp.bExcepVta,0)  " _
          & " FROM Producto P Inner Join ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigRGDet CRGDET ON CP.cCtaCod = CRGDet.cCtaCod " _
          & " JOIN ColocaCestado C on C.cctacod=P.cctacod  and c.nprdestado='" & gColPEstAdjud & "'" _
          & " WHERE p.cCtaCod ='" & psCtacod & "' and p.nprdestado='" & gColPEstAdjud & "'"
          
          ' " and CRGDET.cTpoProceso ='" & psTipoProceso & "'"
          'AND CRGDET.cNroProceso ='" & psNroProceso & "'
          

    Set loRegValida = New dColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida.BOF And lrValida.EOF Then
        MsgBox "El Credito no se encuetra en el Proceso de Subasta ", vbInformation, "Aviso"
        Exit Function
    End If
    
    If (lrValida!nPrdEstado <> gColPEstAdjud) Then
        MsgBox " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
        Exit Function
    End If
    If lrValida!nRgDetEstado = gColPRecGarVentaVendido Then
        MsgBox " Contrato ya fue Subastado " & psCtacod, vbInformation, " Aviso "
        Exit Function
    End If
    
    Set nValidaRegVentaSubastaCredPignoraticio = lrValida
    Set lrValida = Nothing
End Function

Public Function nValidaBloqueoCredPignoraticio(ByVal psCtacod As String) As Recordset

'************************************

Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

lsSQL = "SELECT P.nPrdEstado, PBloq.nBlqTpo, PBloq.nBlqMotivo, ISNULL(PBloq.cComentario,'') cComentario, PBloq.cMovNro as cMovNroBloqueo , " _
    & " cBloqueo = (SELECT CASE WHEN COUNT(nBlqMotivo) > 0 THEN 'S' ELSE 'N' END FROM ProductoBloqueos PB WHERE PB.cCtaCod = '" & psCtacod & "' AND cMovNroDbl Is Null ) " _
    & " FROM Producto P " _
    & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod  " _
    & " LEFT JOIN ProductoBloqueos PBloq ON CP.cCtaCod = PBloq.cCtaCod " _
    & " WHERE P.cCtaCod ='" & psCtacod & "' ORDER BY cMovNro DESC " ' & "' AND cMovNroDbl Is Null "
     
    Set lrValida = New ADODB.Recordset
    Set loRegValida = New dColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Or (lrValida.BOF And lrValida.EOF) Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    
    If (lrValida!nPrdEstado <> gColPEstRegis) And (lrValida!nPrdEstado <> gColPEstDesem) And _
       (lrValida!nPrdEstado <> gColPEstDifer) And (lrValida!nPrdEstado <> gColPEstVenci) And _
       (lrValida!nPrdEstado <> gColPEstPRema) And (lrValida!nPrdEstado <> gColPEstVenci) Then
        MsgBox " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
        Exit Function
    End If
    
    Set nValidaBloqueoCredPignoraticio = lrValida
    Set lrValida = Nothing
    
End Function


Public Function nBuscaOperacionesCredPigParaExtorno(ByVal psFecTrans As String, ByVal psOpeExtor As String, _
        Optional ByVal psCtacod As String = "@", Optional ByVal pbOrdenAsc As Boolean = False) As Recordset

Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = " SELECT M.cMovNro, M.nMovNro, M.cOpeCod, M.nMovFlag, " _
          & " MC.cCtaCod, MC.nMonto, Ope.cOpeDesc,  " _
          & " (Select Max(P.cPersNombre) From Persona p Join ProductoPersona pp on p.cPersCod = pp.cPersCod " _
          & "    Where pp.cCtaCod = MC.cCtaCod And pp.nPrdPersRelac = " & gColRelPersTitular & "  ) cCliente , " _
          & " Isnull((Select mcm.cPersCod from movcmac mcm where mcm.nmovnro = m.nmovnro),'' ) cCMACOpe  " _
          & " FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
          & " LEFT JOIN OpeTpo Ope ON MC.cOpeCod = Ope.cOpeCod " _
          & " WHERE M.cMovNro LIKE '" & psFecTrans & "%' " _
          & " AND M.nMovFlag = " & gMovFlagVigente & " " _
          & " AND Mc.cOpeCod in " & psOpeExtor & "  " _
          & " AND Substring(M.cOpeCod,3,1) <> '" & Mid(geColPExtDupli, 3, 1) & "' "
    If psCtacod <> "@" Then
        lsSQL = lsSQL & " AND MC.cCtaCod ='" & psCtacod & "' "
    End If
    lsSQL = lsSQL & " ORDER BY M.nMovNro DESC  "
    
    Set loRegValida = New dColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Extorno ", vbInformation, "Aviso"
        Exit Function
    End If
    
    Set nBuscaOperacionesCredPigParaExtorno = lrValida
    Set lrValida = Nothing
End Function
'
Public Function nBuscaOperacionesCredPigParaExtornoLLamadaCMAC(ByVal psFecTrans As String, Optional ByVal psCtacod As String = "@", _
        Optional ByVal pbOrdenAsc As Boolean = False) As Recordset

Dim lsSQL As String
Dim loRegValida As dColPFunciones
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
    
    Set loRegValida = New dColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Extorno ", vbInformation, "Aviso"
        Exit Function
    End If
    
    Set nBuscaOperacionesCredPigParaExtornoLLamadaCMAC = lrValida
    Set lrValida = Nothing
End Function


Public Function nDocumentoEmitido(ByVal pnDocTpo As Integer, ByVal psDocNro As String, _
            ByVal psOpecod As String) As Boolean
'Verifica si un documento ha sido emitido
Dim loValida As dColPFunciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

On Error GoTo dError

lsSQL = "SELECT Count(MDoc.nDocTpo) as Cuenta " _
    & " FROM MovDoc MDoc INNER JOIN Mov M ON MDoc.nMovNro = M.nMovNro " _
    & " WHERE MDoc.nDocTpo = " & pnDocTpo & " AND MDoc.cDocNro ='" & psDocNro & "' " _
    & " AND M.cOpeCod IN ( " & psOpecod & " ) "

    Set loValida = New dColPFunciones
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

Public Function nValidaAmortizacionCredPignoraticio(ByVal psCtacod As String, ByVal pdFecsis As Date, _
     ByVal pnMaxDiasDesembolso As Integer, Optional ByVal cUser As String) As Recordset

'************************************
Dim lsSQL As String
'Dim loRegValida As dColPFunciones

Dim lrValida As ADODB.Recordset

Set lrValida = New ADODB.Recordset

    lsSQL = " select m.nmovnro,m.cmovnro,mc.copecod, mcd.nprdconceptocod, montodet=mcd.nmonto,montot=mc.nmonto "
    lsSQL = lsSQL & " into #tmpPrend" & cUser & " from mov m"
    lsSQL = lsSQL & " join movcol mc on mc.nmovnro=m.nmovnro "
    lsSQL = lsSQL & " join movcoldet mcd on mcd.nmovnro=mc.nmovnro and mcd.copecod=mc.copecod "
    lsSQL = lsSQL & " where mc.cctacod='" & psCtacod & "' and m.nmovflag=0 and m.nmovestado=10 "
    oCon.Execute lsSQL
 
    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, C.nPlazo, CP.nNroRenov,  " _
          & " P.nSaldo, CP.nTasacion, P.nTasaInteres, " _
          & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
          & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ),  " _
          & " dDesembolso = (SELECT MAX (CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
          & "                FROM #tmpPrend" & cUser & " " _
          & "                WHERE cOpeCod = " & gColPOpeDesembolsoEFE & " ),  " _
          & " dRenovacion = isnull((SELECT MAX (CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
          & "              FROM #tmpPrend" & cUser & "" _
          & "              WHERE cOpeCod  in ( " & gColPOpeRenovNorEFE & "," & gColPOpeRenovMorEFE & "," _
          & gColPOpeRenovNorEnOtCjEFE & "," & gColPOpeRenovMorEnOtCjEFE & ") ),'01/01/1900') ,  " _
          & " cInicRemate = (SELECT CASE WHEN COUNT(RG.nRGEstado)> 0  THEN 'S' ELSE 'N' END " _
          & "                FROM ColocPigRecGar RG Join ColocPigRGDet RGD ON RG.cTpoProceso = RGD.cTpoProceso And RG.cNroProceso = RGD.cNroProceso " _
          & "                WHERE RG.cTpoProceso ='R' AND DATEDIFF(D, RG.dProceso,'" & Format(pdFecsis, "mm/dd/yyyy") & "') = 0  " _
          & "                AND RG.nRGEStado ='" & gColPRecGarEstIniciado & "' AND RGD.cCtaCod = C.cCtaCod AND RG.cCodAge = Substring(C.cCtaCod,1,5)  )" _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " Inner Join ColocPignoraticio CP on C.cctacod = CP.cctacod " _
          & " WHERE P.cCtaCod ='" & psCtacod & "' "
   
    lrValida.CursorLocation = adUseClient
    lrValida.Open lsSQL, oCon, adOpenStatic, adLockReadOnly, adCmdText
    Set lrValida.ActiveConnection = Nothing
    

    lsSQL = " drop table #tmpPrend" & cUser
    oCon.Execute lsSQL

'    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, C.nPlazo, CP.nNroRenov,  " _
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
'          & " WHERE P.cCtaCod ='" & psCtaCod & "' "
'
'    Set loRegValida = New dColPFunciones
    
  '  Set lrValida = loRegValida
   
      
      
  
   
    If lrValida Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrValida.BOF And lrValida.EOF Then
        MsgBox "ERROR: No se encontro el Credito ", vbInformation, "Aviso"
        Exit Function
    End If
    If (lrValida!nPrdEstado <> gColPEstDesem) And (lrValida!nPrdEstado <> gColPEstVenci) And _
       (lrValida!nPrdEstado <> gColPEstPRema) And (lrValida!nPrdEstado <> gColPEstRenov) Then
        MsgBox " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
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
        MsgBox " No se ha definido la Tasa de Interes Vencida para linea de Credito ", vbInformation, " Aviso "
        Exit Function
    End If
        
    ' Si el Proceso Remate se ha iniciado ' LAYG (21/07/2001)
    If lrValida!cInicRemate = "S" Then
        MsgBox " Contrato se encuentra EN EL PROCESO DE REMATE, no se puede Amortizar ", vbInformation, " Aviso "
        Exit Function
    End If
    
    Set nValidaAmortizacionCredPignoraticio = lrValida
    
    Set lrValida = Nothing
   
    
End Function

Public Function nValidaRegVentaSubastaCredPignoraticioSINSubasta(ByVal psCtacod As String, ByVal psTipoProceso As String, _
    ByVal psNroProceso As String) As Recordset

'************************************
Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado,p.dprdestado, " _
          & " P.nSaldo, CP.nTasacion,nValRegistroAdj = isnull(RGDET.NADJVALREGISTRO,0),CnroProceso=isnull(RGDET.CnroProceso,'00'),CCODAGE=isnull(RGDET.CCODAGE,'00'),CNROPROCESOREM=isnull(RGDET.CNROPROCESOREM,'00'),DPROCESOREM=isnull(RGDET.DPROCESOREM,'01/01/1900'),AGEREM=isnull(RGDET.AGEREM,'00'),bExcepVta=isnull(cp.bExcepVta,0)  " _
          & " FROM Producto P Inner Join ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN (select RG.CTPOPROCESO,RG.CNROPROCESO,RG.DPROCESO,RG.CCODAGE,D.CCTACOD,D.NREMSUBBASEVTA,D.NADJVALREGISTRO,D.NDEUDA,REM.CNROPROCESOREM,REM.DPROCESOREM,REM.AGEREM " _
          & "            from colocpigrecgar RG " _
          & "            JOIN colocpigrgDET D ON D.CNROPROCESO=RG.CNROPROCESO AND D.CTPOPROCESO=RG.CTPOPROCESO " _
          & "            JOIN (Select CNROPROCESOREM=CNROPROCESO,DPROCESOREM=DPROCESO,AGEREM=CCODAGE FROM colocpigrecgar where CTPOPROCESO='R' ) REM on CONVERT(CHAR(8),REM.DPROCESOREM,112)=CONVERT(CHAR(8),RG.DPROCESO,112) " _
          & "            WHERE RG.CTPOPROCESO='A') RGDET ON RGDET.CCTACOD=P.CCTACOD  " _
          & " WHERE  P.cCtaCod ='" & psCtacod & "'"
          
          '& " nValRegistroAdj = ISNULL( (Select nAdjValRegistro FROM ColocPigRGDET  " _
          '& "                            WHERE cTpoProceso = 'A' AND cCtaCod = P.cCtaCod ) , 0 ), " _
          '& " CnroProceso= ISNULL( (Select cNroProceso FROM ColocPigRGDET  " _
          '& "                            WHERE cTpoProceso = 'A' AND cCtaCod = P.cCtaCod ) , '' ), " _
          '& " Ccodage= ISNULL( (Select Ccodage FROM ColocPigRGDET  " _
          '& "                            WHERE cTpoProceso = 'A' AND cCtaCod = P.cCtaCod ) , '' ) "
          

    Set loRegValida = New dColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida.BOF And lrValida.EOF Then
        MsgBox "El Credito no se encuetra en el Proceso de Subasta ", vbInformation, "Aviso"
        Exit Function
    End If
    If (lrValida!nPrdEstado <> gColPEstAdjud) Then
        MsgBox " Contrato se encuentra " & mfgEstadoCredPigDesc(lrValida!nPrdEstado), vbInformation, " Aviso "
        Exit Function
    End If
    'If lrValida!nRGDetEstado = gColPRecGarVentaVendido Then
    '    MsgBox " Contrato ya fue Subastado " & psCtaCod, vbInformation, " Aviso "
    '    Exit Function
    'End If
    
    Set nValidaRegVentaSubastaCredPignoraticioSINSubasta = lrValida
    Set lrValida = Nothing
End Function


'Muestra para Remate
Public Function CalcDeudaRemate(ByVal psFecRemate As String, ByVal psAgencia As String, ByVal psNroRemate As String, ByVal psCtacod As String, Optional ByRef pnSaldoDeuda As Double, Optional ByRef pnIntVencido As Double, Optional ByRef pnComision As Double) As Double
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As DColPCalculos
Set loParam = New DColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NColPCalculos

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
Dim loMuestraContrato As DColPContrato
Dim item As Integer, itemTMP As Integer
Dim lmDetalle(20) As String

pnPreOro14 = 0: pnPreOro16 = 0: pnPreOro21 = 0: pnPreOro18 = 0

    lsSQL = "select * from colocpigrgprecio where ctpoproceso='R' and  cnroproceso='" & psNroRemate & "' and ccodage='" & psAgencia & "'"
    
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    While Not lrDataRep.EOF
            Select Case lrDataRep!ctpoprecio
            
                Case 14
                        pnPreOro14 = lrDataRep!nprecio
                Case 16
                        pnPreOro16 = lrDataRep!nprecio
                Case 18
                        pnPreOro18 = lrDataRep!nprecio
                Case 21
                        pnPreOro21 = lrDataRep!nprecio
                Case 3005
                        pnTasaPreparacionRemate = lrDataRep!nprecio
                Case 3004
                        pnTasaCustodiaVencida = lrDataRep!nprecio
                Case 3003
                        pnTasaImpuesto = lrDataRep!nprecio
                Case 3041
                        pnFactorPrecioVenta = lrDataRep!nprecio
                
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
    lsSQL = lsSQL & " WHERE p.cctacod='" & psCtacod & "'"
    lsSQL = lsSQL & " GROUP BY P.cCtaCod, P.nPrdestado ,C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, "
    lsSQL = lsSQL & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote  "

'& " AND Substring(P.cCtaCod,4,2) = '" & psAgencia & "' AND DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  " _

    Set lrDataRep = New ADODB.Recordset
    Set loDataRep = New dColPFunciones
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
                
                Set loCalculos = New NColPCalculos
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
Public Function CalcDeudaRemateSINPROCESO(ByVal psFecEstado As String, ByVal psCtacod As String, Optional ByRef pnSaldoDeuda As Double, Optional ByRef pnIntVencido As Double, Optional ByRef pnComision As Double, Optional ByRef psFecRemate As String) As Double
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As DColPCalculos
Set loParam = New DColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NColPCalculos

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
Dim loMuestraContrato As DColPContrato
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
    
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    While Not lrDataRep.EOF
            Select Case lrDataRep!ctpoprecio
            
                Case 14
                        pnPreOro14 = lrDataRep!nprecio
                Case 16
                        pnPreOro16 = lrDataRep!nprecio
                Case 18
                        pnPreOro18 = lrDataRep!nprecio
                Case 21
                        pnPreOro21 = lrDataRep!nprecio
                Case 3005
                        pnTasaPreparacionRemate = lrDataRep!nprecio
                Case 3004
                        pnTasaCustodiaVencida = lrDataRep!nprecio
                Case 3003
                        pnTasaImpuesto = lrDataRep!nprecio
                Case 3041
                        pnFactorPrecioVenta = lrDataRep!nprecio
                
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
    lsSQL = lsSQL & " WHERE p.cctacod='" & psCtacod & "'"
    lsSQL = lsSQL & " GROUP BY P.cCtaCod, P.nPrdestado ,C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, "
    lsSQL = lsSQL & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote  "

'& " AND Substring(P.cCtaCod,4,2) = '" & psAgencia & "' AND DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  " _

    Set lrDataRep = New ADODB.Recordset
    Set loDataRep = New dColPFunciones
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
                
                Set loCalculos = New NColPCalculos
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
Public Function CalcPrecioVta(ByVal pdFecsis As Date, ByVal psCtacod As String) As Double
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As DColPCalculos
Set loParam = New DColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NColPCalculos

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
Dim loMuestraContrato As DColPContrato
Dim item As Integer, itemTMP As Integer
Dim lmDetalle(20) As String

pnPreOro14 = 0: pnPreOro16 = 0: pnPreOro21 = 0: pnPreOro18 = 0

    lsSQL = " SELECT nmaterial,nvalor FROM ColocPigPrecioMaterial "
    lsSQL = lsSQL & " WHERE nTipoValor = 1 AND nmaterial in (14,16,18,21) and nMoneda = '1' AND DATEDIFF(d, dInicial, '" & Format(pdFecsis, "mm/dd/yyyy") & "') > 0 "
    lsSQL = lsSQL & " ORDER BY dInicial Desc "
    
    
    Set loDataRep = New dColPFunciones
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
    lsSQL = lsSQL & " WHERE p.cctacod='" & psCtacod & "'"
    lsSQL = lsSQL & " GROUP BY P.cCtaCod, P.nPrdestado ,C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, "
    lsSQL = lsSQL & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote  "

    Set lrDataRep = New ADODB.Recordset
    Set loDataRep = New dColPFunciones
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
'                Set loCalculos = New NColPCalculos
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



Private Sub Class_Initialize()
Dim sConn As String
Dim ClsIni As ClsIni.ClasIni
Set ClsIni = New ClsIni.ClasIni
sConn = ClsIni.CadenaConexion
sDBComunes = ClsIni.BaseComunes
sDBPersona = ClsIni.BasePersonas
sDBImagenes = ClsIni.BaseImagenes
Set ClsIni = Nothing
Set oCon = New Connection
oCon.Open sConn
oCon.Execute "SET DATEFORMAT MDY"
End Sub

Private Sub Class_Terminate()
oCon.Close
Set oCon = Nothing
End Sub
