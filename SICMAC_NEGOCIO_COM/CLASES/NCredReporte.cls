VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCredReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim csNomCMAC As String
Dim csNomAgencia As String
Dim csCodUser As String
Dim csFechaSis As String
 
Public Event ShowProgress()
Public Event CloseProgress()
Public Event Progress(pnValor As Long, pnTotal As Long)
 
Public Sub inicio(ByVal psNomCmac As String, ByVal psNomAgencia As String, _
            ByVal psCodUser As String, ByVal psFechaSis As String)
    csNomCMAC = psNomCmac
    csNomAgencia = psNomAgencia
    csCodUser = psCodUser
    csFechaSis = psFechaSis
End Sub
        
Public Function nRepoCabecera(ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        ByVal pnCodReporte As String) As String
        
Dim lsCadImp As String
Dim loImprimeCab As NColPImpre
Dim lsNegritaOn As String
Dim lsNegritaOff As String

lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

    Set loImprimeCab = New NColPImpre
        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta, pnCodReporte)
    Set loImprimeCab = Nothing
        
    If pnCodReporte = "108604" Or pnCodReporte = "108605" Or pnCodReporte = "108201" Or pnCodReporte = "108203" Or pnCodReporte = "138302" Or pnCodReporte = "138303" Or pnCodReporte = "108304" Or pnCodReporte = "108204" Then
        lsCadImp = lsCadImp & Chr(10) & String(pnAnchoLinea, "-") & Chr(10)
    End If
    Select Case pnCodReporte
        Case 108604
            lsCadImp = lsCadImp & lsNegritaOn & "      |      |      C    A    R    T    E    R    A          |     TOTAL     | COLOCACIONES  |  I N D I C E  " & lsNegritaOff & Chr(10)
            lsCadImp = lsCadImp & lsNegritaOn & " ITEM | ANA. | REFINANCIADA  |    VENCIDA    |   JUDICIAL    |     C.A.R.    |    VIGENTES   |     C.A.R.    " & lsNegritaOff & Chr(10)
            lsCadImp = lsCadImp & lsNegritaOn & "      |      |     (A)       |      (B)      |     (C)       |(D)=(A)+(B)+(C)|      (E)      |    (D)/(E)    " & lsNegritaOff & Chr(10)
        Case 108605
            lsCadImp = lsCadImp & lsNegritaOn & " FUENTE FINANCIAMI. |     C  A  R  T  E  R  A   D E   C  R  E  D  I  T  O  S    |        C  A  R  T  E  R  A    V  E  N  C  I  D  A         |" & lsNegritaOff & Chr(10)
            lsCadImp = lsCadImp & lsNegritaOn & "   A G E N C I A    |  MONEDA NACIONAL  | MONEDA EXTRANJERA |TOTAL(Nuevos Soles)|  MONEDA NACIONAL  | MONEDA EXTRANJERA |TOTAL(Nuevos Soles)|" & lsNegritaOff & Chr(10)
            lsCadImp = lsCadImp & lsNegritaOn & "                    | NUM.|  M O N T O  | NUM.|  M O N T O  | NUM.|  M O N T O  | NUM.|  M O N T O  | NUM.|  M O N T O  | NUM.|  M O N T O  |" & lsNegritaOff & Chr(10)
        Case 108203
            lsCadImp = lsCadImp & " CREDITO            CREDITO ANTIGUO           CLIENTE                     MONEDA     M.APROBADO     S.CAPITAL DESEMBOLSO  TASA  ANA. ATRASO" & Chr(10)
        Case 108304
            lsCadImp = lsCadImp & " FECHA      CREDITO            CLIENTE                             APLICADO   CONCEPTO                               MONTO ANA." & Chr(10)
'        Case 108202
'            lsCadImp = lsCadImp & "           |            D  E  S  E  M  B  O  L  S  O  S            |           P       A       G       O       S           |    TOTAL    |    SALDO    |    |    |    |    |    |    |      " & Chr(10)
'            lsCadImp = lsCadImp & "   FECHA   |    NORMAL     REFINANCIA. | M.REFINANC.    M. JUDIC.  |   CAPITAL   |   INTERES   |    MORA     |    OTROS    |   COBRADO   |    CAPITAL  |NORM|REF.|CANC|REF.|JUD.|SAL.|PAG.  " & Chr(10)
        Case 108204
            lsCadImp = lsCadImp & " COD. ACTUAL        CLIENTE                             MONEDA  TASA INT. DESEMBOLSADO  VIGENCIA     S. CAPITAL ANA. ATRASO GARANTIAS    " & Chr(10)
    End Select
    If pnCodReporte = "108604" Or pnCodReporte = "108605" Or pnCodReporte = "108203" Or pnCodReporte = "108304" Or pnCodReporte = "108202" Or pnCodReporte = "108204" Then
        lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & Chr(10)
    End If
nRepoCabecera = lsCadImp
End Function
 
Public Function nRepo108201_CreditosVigentes_DiasAtraso(ByVal cSubTit As String, ByVal pdFecFin As String, ByVal pnDiaIni As Integer, ByVal pnDiaFin As Integer, ByVal pcMoneda As String, ByVal pcProductos As String, ByVal pcCondicion As String, ByVal pcAnalistas) As String
  
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nitem As Integer
Dim nTNacionalAprob As Double
Dim nTExtranjeroAprob As Double
Dim nTNacionalSCapital As Double
Dim nTExtranjeroSCapital As Double
Dim loImprimeCab As NColPImpre
Dim sCliente As String * 35
Dim sMoneda As String * 7
Dim lsCondiciones As String
Dim sCriterio As String
Dim i As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
 
Dim oNCred As NCredito
Dim MatCalend As Variant
Dim MatCalendPar As Variant
Dim ntempoDeuda As Currency
   
Dim sTempoMoneda As String
Dim sTempoPlazo As String
Dim sTempoProducto As String
Dim sTempoFinanciamiento As String
Dim sTempoAnalista  As String

'Totales por Plazo
Dim sTotalP1 As Currency
Dim sTotalP2 As Currency
Dim sTotalP3 As Currency
Dim sTotalP4 As Currency
Dim sTotalP5 As Currency
Dim sTotalP0 As Long

'Totales por Moneda
Dim sTotalM1 As Currency
Dim sTotalM2 As Currency
Dim sTotalM3 As Currency
Dim sTotalM4 As Currency
Dim sTotalM5 As Currency
Dim sTotalM0 As Long

'Totales por Producto
Dim sTotalR1 As Currency
Dim sTotalR2 As Currency
Dim sTotalR3 As Currency
Dim sTotalR4 As Currency
Dim sTotalR5 As Currency
Dim sTotalR0 As Long

'Totales por Fuente de Financiamiento
Dim sTotalF1 As Currency
Dim sTotalF2 As Currency
Dim sTotalF3 As Currency
Dim sTotalF4 As Currency
Dim sTotalF5 As Currency
Dim sTotalF0 As Long

'Totales Generales
Dim sTotalA As Currency
Dim sTotalB As Currency
Dim sTotalC As Currency
Dim sTotalD As Currency
Dim sTotalE As Currency
Dim sTotal0 As Long

   
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

lsCondiciones = "De " & pnDiaIni & " a " & pnDiaFin & " Dias de Atraso"

lsSQL = " SELECT C1.cCtaCod as cCredito, C1.nPlazo, C1.cPlazo, C1.nProducto, C1.cProducto, C1.cCliente, C1.nSaldoCapital, C1.nMoneda, C1.cMoneda, " & _
        " C2.nDesembolsado, C3.nNroProxCuota as nCuotaPagar, C4.nTotalCuotas, C6.Cuota, C5.dVenc, C1.nDiasAtraso, C1.dVigencia, C1.cRefinan, C2.nPlazoDias, C8.nIntMoratorio, C1.nTasaInteres, C1.Analista, C1.CodAnalista, C1.cNomAnalista " & _
        " , dbo.InteresDevengado(C1.cCtaCod,'" & Format(pdFecFin, "mm/dd/yyyy") & "' ) as nIntDevengado From ( " & _
        " SELECT Producto.nTasaInteres, Producto.nSaldo as nSaldoCapital,  dbo.ColocacCred.cCtaCod, SUBSTRING(dbo.Colocaciones.cLineaCred, 1, 2) AS nFuente,  " & _
        " SUBSTRING(dbo.Colocaciones.cLineaCred, 6, 1) AS nPlazo, (SELECT CO.cconsdescripcion FROM Constante CO " & _
        " WHERE Co.nConsCod=" & gColocLineaCredPlazo & _
        " AND CO.nConsValor = SUBSTRING(dbo.Colocaciones.cLineaCred, 6, 1)) AS cPlazo, " & _
        " SUBSTRING(dbo.Colocaciones.cCtaCod, 6, 3) AS nProducto, (SELECT COO.cConsDescripcion FROM Constante COO " & _
        " WHERE Substring(Colocaciones.cCtaCod, 6, 3) = COO.nConsValor " & _
        " AND COO.nConsCod=" & gProducto & ") AS cProducto, dbo.Persona.cPersNombre AS cCLiente, dbo.Producto.nSaldo AS nCapital, " & _
        " (SELECT RHH.CUSER FROM RRHH RHH INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD " & _
        " WHERE PRP.CCTACOD = ProductoPersona.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS ANALISTA, (SELECT PRP.CPERSCOD " & _
        " FROM PRODUCTOPERSONA PRP WHERE PRP.CCTACOD = ProductoPersona.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS CODANALISTA, " & _
        " (SELECT PERSO.cPersNombre FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD " & _
        " WHERE PRP.CCTACOD = ProductoPersona.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS cNomAnalista, dbo.Constante.cConsDescripcion AS cMoneda, SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) AS nMoneda, dbo.ColocacCred.nColocCondicion, " & _
        " dbo.ColocacCred.nDiasAtraso , dbo.Colocaciones.dVigencia , " & _
        " cRefinan=case when Producto.nPrdEstado IN('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "') Then 'S' else 'N'  END " & _
        " FROM dbo.ColocacCred INNER JOIN dbo.Colocaciones ON dbo.ColocacCred.cCtaCod = dbo.Colocaciones.cCtaCod INNER JOIN " & _
        " dbo.Producto ON dbo.Colocaciones.cCtaCod = dbo.Producto.cCtaCod INNER JOIN dbo.ProductoPersona ON dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod " & _
        " AND dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod " & _
        " INNER JOIN dbo.Persona ON dbo.ProductoPersona.cPersCod = dbo.Persona.cPersCod AND dbo.ProductoPersona.cPersCod = dbo.Persona.cPersCod " & _
        " INNER JOIN dbo.Constante ON SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) = dbo.Constante.nConsValor " & _
        " WHERE (dbo.ProductoPersona.nPrdPersRelac=" & gColRelPersTitular & ") " & _
        " AND (dbo.ColocacCred.nDiasAtraso>=" & pnDiaIni & " and dbo.ColocacCred.nDiasAtraso<=" & pnDiaFin & ") "
        
        lsSQL = lsSQL & " AND (dbo.Constante.nConsCod= " & gMoneda & ") " & _
        " AND (Producto.nPrdEstado IN ('" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "')) "

        If Len(Trim(pcProductos)) > 0 Then
            lsSQL = lsSQL & " AND (SUBSTRING(dbo.Colocaciones.cCtaCod, 6, 3) IN (" & pcProductos & "))"
        End If
        
        If Len(pcAnalistas) > 0 Then
            lsSQL = lsSQL & " AND ((SELECT PRP.CPERSCOD FROM PRODUCTOPERSONA PRP WHERE PRP.CCTACOD = ProductoPersona.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") IN (" & pcAnalistas & ")) "
        End If
        
        If Len(Trim(pcMoneda)) > 0 Then
            lsSQL = lsSQL & " AND (SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) IN (" & pcMoneda & ")) "
        End If
        
        If Len(Trim(pcCondicion)) > 0 Then
            lsSQL = lsSQL & " AND (dbo.ColocacCred.nColocCondicion IN (" & pcCondicion & ")) "
        End If
        
        lsSQL = lsSQL & " ) C1 Inner Join " & _
        " ( SELECT dbo.ColocacCred.cCtaCod, dbo.Colocaciones.nPlazo as nPlazoDias, dbo.Colocaciones.nMontoCol AS nDesembolsado " & _
        " FROM dbo.ColocacCred INNER JOIN dbo.Colocaciones ON dbo.ColocacCred.cCtaCod = dbo.Colocaciones.cCtaCod " & _
        " ) C2 ON C1.cCtaCod=C2.cCtaCod " & _
        " Inner Join " & _
        " ( SELECT dbo.ColocacCred.cCtaCod, dbo.ColocacCred.nNroCalen, dbo.ColocacCred.nNroProxCuota " & _
        " FROM dbo.ColocacCred INNER JOIN dbo.Colocaciones ON dbo.ColocacCred.cCtaCod = dbo.Colocaciones.cCtaCod " & _
        " ) C3 " & _
        " on C2.cCtaCod=C3.cCtaCod Inner Join " & _
        " ( SELECT cCtaCod, nNroCalen, nCuota, dVenc From dbo.ColocCalendario Where nColocCalendApl=" & gColocCalendAplCuota & _
        " ) C5 ON " & _
        " C3.cCtaCod = C5.cCtaCod And C3.nNrocalen = C5.nNrocalen And C3.nNroProxCuota = C5.nCuota " & _
        " Inner Join (" & _
        " SELECT  cCtaCod, nNroCalen, COUNT(nCuota) AS nTotalCuotas From dbo.ColocCalendario " & _
        " Where nColocCalendApl = 1 GROUP BY cCtaCod, nNroCalen " & _
        " ) C4 ON C3.cCtaCod = C4.cCtaCod AND C3.nNroCalen=C4.nNroCalen " & _
        " Inner Join " & _
        " (SELECT cCtaCod, nNroCalen, nCuota,  SUM(nMonto) AS Cuota From dbo.ColocCalendDet Where nColocCalendApl = 1 " & _
        " GROUP BY cCtaCod, nNroCalen, nCuota " & _
        " ) C6 ON " & _
        " C3.cCtaCod = C6.cCtaCod And C3.nNrocalen = C6.nNrocalen And C3.nNroProxCuota = C6.nCuota Inner Join " & _
        " ( "
        lsSQL = lsSQL & " SELECT cCtaCod, nNroCalen, nCuota,  SUM(nMonto) AS nIntMoratorio " & _
        " From dbo.ColocCalendDet Where nColocCalendApl=" & gColocCalendAplCuota & " And nPrdConceptoCod=" & gColocConceptoCodInteresMoratorio & _
        " GROUP BY cCtaCod, nNroCalen, nCuota " & _
        " ) C8 ON " & _
        " C3.cCtaCod = C8.cCtaCod And C3.nNrocalen = C8.nNrocalen And C3.nNroProxCuota = C8.nCuota "
 
    lsSQL = lsSQL & " ORDER BY nMoneda, nProducto, C1.cCliente  "
    
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
          
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CREDITOS VIGENTES AL " & Format(pdFecFin, "dd/MM/YYYY"), lsCondiciones, lnPage, 184, cSubTit, 108201)
        lsCadImp = lsCadImp & Chr(10)
        
        lnIndice = 8:  lnLineas = 8
        
        sTempoMoneda = lrDataRep!nmoneda
        'sTempoPlazo = lrDataRep!nPlazo
        sTempoProducto = lrDataRep!nProducto
        'sTempoFinanciamiento = lrDataRep!nFuente
        'sTempoAnalista = lrDataRep!Analista
        
         
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA         : " & lrDataRep!cMoneda & lsNegritaOff & Chr(10) & Chr(10)
        'Plazo
        'lsCadImp = lsCadImp & lsNegritaOn & "PLAZO          : " & lrDataRep!cPlazo & lsNegritaOff & Chr(10) & Chr(10)
        'Producto
        lsCadImp = lsCadImp & lsNegritaOn & "PRODUCTO       : " & lrDataRep!cProducto & " : " & lrDataRep!nProducto & lsNegritaOff & Chr(10) & Chr(10)
        'Fuente de Financiamiento
        'lsCadImp = lsCadImp & lsNegritaOn & "FINANCIAMIENTO : " & lrDataRep!cFuente & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(184, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(1) & "CREDITO            CLIENTE                             CUOT/PLAZO     DESEMBOLSO       CUOTA      SALDO CAP. ATRA.        MORA       INTERES     DEUDA TOTAL  ANA. REF " & Chr(10)
        lsCadImp = lsCadImp & String(184, "-") & Chr(10)
        
        lnLineas = lnLineas + 11
        lnIndice = lnIndice + 11
          
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nitem = nitem + 1
                
                ntempoDeuda = 0
                
                Set oNCred = New NCredito
                MatCalend = oNCred.RecuperaMatrizCalendarioPendiente(!cCredito)
                MatCalendPar = oNCred.RecuperaMatrizCalendarioPendiente(!cCredito, True)
                MatCalend = UnirMatricesMiViviendaAmortizacion(MatCalend, MatCalendPar)
                
                
                ntempoDeuda = oNCred.MatrizDeudaAlaFecha(!cCredito, MatCalend, pdFecFin)
                Set oNCred = Nothing
                
                'Totalizar por Moneda
                                   
                If sTempoMoneda <> !nmoneda Then
                     
                    lsCadImp = lsCadImp & String(184, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL MONEDA               :  CASOS:  " & Space(1) & ImpreFormat(sTotalM0, 5, 0) & _
                                Space(23) & ImpreFormat(sTotalM1, 10, 2) & Space(15) & ImpreFormat(sTotalM2, 10, 2) & _
                                Space(7) & ImpreFormat(sTotalM3, 10, 2) & Space(1) & ImpreFormat(sTotalM4, 10, 2) & Space(1) & ImpreFormat(sTotalM5, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                                
                    lsCadImp = lsCadImp & String(184, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL PRODUCTO             :  CASOS:  " & Space(1) & ImpreFormat(sTotalR0, 5, 0) & _
                                Space(23) & ImpreFormat(sTotalR1, 10, 2) & Space(15) & ImpreFormat(sTotalR2, 10, 2) & _
                                Space(7) & ImpreFormat(sTotalR3, 10, 2) & Space(1) & ImpreFormat(sTotalR4, 10, 2) & Space(1) & ImpreFormat(sTotalR5, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                                                              
                                  
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA         : " & lrDataRep!cMoneda & lsNegritaOff & Chr(10) & Chr(10)
                    'Producto
                    lsCadImp = lsCadImp & lsNegritaOn & "PRODUCTO       : " & lrDataRep!cProducto & " : " & lrDataRep!nProducto & lsNegritaOff & Chr(10) & Chr(10)
                    
                    
                    lsCadImp = lsCadImp & String(184, "-") & Chr(10)
                    lsCadImp = lsCadImp & Space(1) & "CREDITO            CLIENTE                             CUOT/PLAZO     DESEMBOLSO       CUOTA      SALDO CAP. ATRA.        MORA       INTERES     DEUDA TOTAL  ANA. REF " & Chr(10)
                    lsCadImp = lsCadImp & String(184, "-") & Chr(10)
                    
                    sTotalM0 = 1
                    sTotalM1 = !nDesembolsado
                    sTotalM2 = Round(!nSaldoCapital, 2)
                    sTotalM3 = Round(!nIntMoratorio, 2)
                    sTotalM4 = Round(!nIntDevengado, 2)
                    sTotalM5 = Round(ntempoDeuda, 2)
                    
                    sTotalP0 = 1
                    sTotalP1 = !nDesembolsado
                    sTotalP2 = Round(!nSaldoCapital, 2)
                    sTotalP3 = Round(!nIntMoratorio, 2)
                    sTotalP4 = Round(!nIntDevengado, 2)
                    sTotalP5 = Round(ntempoDeuda, 2)
                    
                    sTotalR0 = 1
                    sTotalR1 = !nDesembolsado
                    sTotalR2 = Round(!nSaldoCapital, 2)
                    sTotalR3 = Round(!nIntMoratorio, 2)
                    sTotalR4 = Round(!nIntDevengado, 2)
                    sTotalR5 = Round(ntempoDeuda, 2)
                    
                    sTotalF0 = 1
                    sTotalF1 = !nDesembolsado
                    sTotalF2 = Round(!nSaldoCapital, 2)
                    sTotalF3 = Round(!nIntMoratorio, 2)
                    sTotalF4 = Round(!nIntDevengado, 2)
                    sTotalF5 = Round(ntempoDeuda, 2)
                        
                    nitem = 1
                    
                    sTempoMoneda = lrDataRep!nmoneda
                    'sTempoPlazo = lrDataRep!nPlazo
                    sTempoProducto = lrDataRep!nProducto
                    'sTempoFinanciamiento = lrDataRep!nFuente
                      
                    lnLineas = lnLineas + 9
                    lnIndice = lnIndice + 9
                     
                Else 'si es la misma moneda
                 
                    sTotalM0 = sTotalM0 + 1
                    sTotalM1 = sTotalM1 + !nDesembolsado
                    sTotalM2 = sTotalM2 + Round(!nSaldoCapital, 2)
                    sTotalM3 = sTotalM3 + Round(!nIntMoratorio, 2)
                    sTotalM4 = sTotalM4 + Round(IIf(IsNull(!nIntDevengado), 0, !nIntDevengado), 2)
                    sTotalM5 = sTotalM5 + Round(ntempoDeuda, 2)
                                                            
                    
                    If sTempoProducto <> !nProducto Then
                                        
                         
                        lsCadImp = lsCadImp & String(184, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL PRODUCTO             :  CASOS:  " & Space(1) & ImpreFormat(sTotalR0, 5, 0) & _
                                    Space(23) & ImpreFormat(sTotalR1, 10, 2) & Space(15) & ImpreFormat(sTotalR2, 10, 2) & _
                                    Space(7) & ImpreFormat(sTotalR3, 10, 2) & Space(1) & ImpreFormat(sTotalR4, 10, 2) & Space(1) & ImpreFormat(sTotalR5, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                         
                        'Producto
                        lsCadImp = lsCadImp & lsNegritaOn & "PRODUCTO       : " & lrDataRep!cProducto & " : " & lrDataRep!nProducto & lsNegritaOff & Chr(10) & Chr(10)
                    
                        lsCadImp = lsCadImp & String(184, "-") & Chr(10)
                        lsCadImp = lsCadImp & Space(1) & "CREDITO            CLIENTE                             CUOT/PLAZO     DESEMBOLSO       CUOTA      SALDO CAP. ATRA.        MORA       INTERES     DEUDA TOTAL  ANA. REF " & Chr(10)
                        lsCadImp = lsCadImp & String(184, "-") & Chr(10)
                                                
                         
                        sTotalP0 = 1
                        sTotalP1 = !nDesembolsado
                        sTotalP2 = Round(!nSaldoCapital, 2)
                        sTotalP3 = Round(!nIntMoratorio, 2)
                        sTotalP4 = Round(IIf(IsNull(!nIntDevengado), 0, !nIntDevengado), 2)
                        sTotalP5 = Round(ntempoDeuda, 2)
                                                 
                        sTotalF0 = 1
                        sTotalF1 = !nDesembolsado
                        sTotalF2 = Round(!nSaldoCapital, 2)
                        sTotalF3 = Round(!nIntMoratorio, 2)
                        sTotalF4 = Round(IIf(IsNull(!nIntDevengado), 0, !nIntDevengado), 2)
                        sTotalF5 = Round(ntempoDeuda, 2)
                        
                        sTotalR0 = 1
                        sTotalR1 = !nDesembolsado
                        sTotalR2 = Round(!nSaldoCapital, 2)
                        sTotalR3 = Round(!nIntMoratorio, 2)
                        sTotalR4 = Round(IIf(IsNull(!nIntDevengado), 0, !nIntDevengado), 2)
                        sTotalR5 = Round(ntempoDeuda, 2)
                        
                        nitem = 1
                        
                        'sTempoPlazo = lrDataRep!nPlazo
                        sTempoProducto = lrDataRep!nProducto
                        'sTempoFinanciamiento = lrDataRep!nFuente
                    
                        lnLineas = lnLineas + 7
                        lnIndice = lnIndice + 7
                    Else
                        sTotalR0 = sTotalR0 + 1
                        sTotalR1 = sTotalR1 + !nDesembolsado
                        sTotalR2 = sTotalR2 + Round(!nSaldoCapital, 2)
                        sTotalR3 = sTotalR3 + Round(!nIntMoratorio, 2)
                        sTotalR4 = sTotalR4 + Round(IIf(IsNull(!nIntDevengado), 0, !nIntDevengado), 2)
                        sTotalR5 = sTotalR5 + Round(ntempoDeuda, 2)
                                            
                    End If
                    
                End If
                
                sTotal0 = sTotal0 + 1
                sTotalA = sTotalA + !nDesembolsado
                sTotalB = sTotalB + Round(!nSaldoCapital, 2)
                sTotalC = sTotalC + Round(!nIntMoratorio, 2)
                sTotalD = sTotalD + Round(!nIntDevengado, 2)
                sTotalE = sTotalE + Round(ntempoDeuda, 2)
                                
                sCliente = !cCliente
                 
                lsCadImp = lsCadImp & Space(1) & !cCredito & Space(1) & sCliente & Space(1) & _
                          ImpreFormat(!nCuotaPagar, 4, 0) & "/" & ImpreFormat(!nTotalCuotas, 4, 0) & Space(1) & ImpreFormat(!nDesembolsado, 10, 2) & Space(1) & _
                          ImpreFormat(Round(!Cuota, 2), 10, 2) & Space(1) & ImpreFormat(Round(!nSaldoCapital, 2), 10, 2) & Space(1) & ImpreFormat(!nDiasAtraso, 5, 0) & Space(1) & _
                          ImpreFormat(Round(!nIntMoratorio, 2), 10, 2) & Space(1) & ImpreFormat(Round(!nIntDevengado, 2), 10, 2) & Space(1) & ImpreFormat(Round(ntempoDeuda, 2), 10, 2) & Space(6) & _
                          !ANALISTA & Space(1) & !cRefinan & Chr(10)
                
                 
               'lnLineas = lnLineas + 1
               
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 50 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & Chr(10)
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS VIGENTES AL " & Format(pdFecFin, "dd/MM/YYYY"), lsCondiciones, lnPage, 184, cSubTit, 108201)
                    lnLineas = 8
                    lnIndice = lnIndice + 8
                End If
                
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
                
            Loop
            RaiseEvent CloseProgress
                
        End With
                
         
        lsCadImp = lsCadImp & String(184, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL PRODUCTO             :  CASOS:  " & Space(1) & ImpreFormat(sTotalR0, 5, 0) & _
                    Space(23) & ImpreFormat(sTotalR1, 10, 2) & Space(15) & ImpreFormat(sTotalR2, 10, 2) & _
                    Space(7) & ImpreFormat(sTotalR3, 10, 2) & Space(1) & ImpreFormat(sTotalR4, 10, 2) & Space(1) & ImpreFormat(sTotalR5, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
         
         
        lsCadImp = lsCadImp & String(184, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL MONEDA               :  CASOS:  " & Space(1) & ImpreFormat(sTotalM0, 5, 0) & _
                    Space(23) & ImpreFormat(sTotalM1, 10, 2) & Space(15) & ImpreFormat(sTotalM2, 10, 2) & _
                    Space(7) & ImpreFormat(sTotalM3, 10, 2) & Space(1) & ImpreFormat(sTotalM4, 10, 2) & Space(1) & ImpreFormat(sTotalM5, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                     
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
    nRepo108201_CreditosVigentes_DiasAtraso = lsCadBuffer

End Function

Public Function nRepo108202_EstadisticaMensualCreditos(ByVal pnTipoReporte As Byte, ByVal cSubTit As String, _
            ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal pcMoneda As String, ByVal pcProductos As String, _
            ByVal pcLineaCredito As String, ByVal psAgencias As String) As String

Dim lsSQL As String
Dim rs As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nitem As Long
Dim lnSaldoCap As Currency
Dim lnCredSalCap As Long
Dim lnSaldoCapTot As Currency
Dim lnCredSalCapTot As Long
Dim loImprimeCab As NColPImpre
Dim lsCondiciones As String
Dim lsNegritaOn As String
Dim lsNegritaOff As String

Dim lnTotalCobrado As Currency

lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

lnPage = 1
lnSaldoCap = 0: lnSaldoCapTot = 0
lnCredSalCap = 0: lnCredSalCapTot = 0

lsCondiciones = "Del " & pdFecIni & " Al " & pdFecFin

   lsSQL = "SELECT " _
        & " CASE WHEN (GROUPING(cAgeDescripcion) = 1) THEN 'TOTAL' ELSE ISNULL(cAgeDescripcion, 'AGENCIA DESCONOCIDO ') END AS sAgencia, "

    If pnTipoReporte = 2 Or pnTipoReporte = 3 Then
        lsSQL = lsSQL & " CASE WHEN (GROUPING(cLineaCred) = 1) THEN 'TOTAL' ELSE ISNULL(cLineaCred, 'LINEA DESCONOCIDO ') END AS sLinea, "
    End If

    lsSQL = lsSQL & " CASE WHEN (GROUPING(dEstad) = 1) THEN 'TOTAL' ELSE ISNULL(dEstad, 'DESCONOCIDO ') END AS sEstad, " _
        & " Sum(nMontoDesembN) nMontoDesembN, sum(nMontoDesembR) nMontoDesembR, " _
        & " Sum(nMontoRefinan) nMontoRefinan, sum(nMontoJudic) nMontoJudic, Sum(nCapPag) nCapPag, " _
        & " Sum(nIntPag) nIntPag, Sum(nMoraPag) nMoraPag, Sum(nOtrPag) nOtrPag, Sum(nSaldoCap) nSaldoCap, " _
        & " Sum(nNumDesembN) nNumDesembN, Sum(nNumDesembR) nNumDesembR, Sum(nNumCanc) nNumCanc, " _
        & " Sum(nNumRef) nNumRef, Sum(nNumJud) nNumJud, Sum(nNumSaldos) nNumSaldos, Sum(nNumPagos) nNumPagos " _
        & " From " _
        & " (SELECT substring(cLineaCred,5,1) nMoneda, cCodAge, cLineaCred, convert(varchar(10), dEstad, 101) dEstad, " _
        & " nMontoDesembN,nMontoDesembR, nMontoRefinan, nMontoJudic, nCapPag, nIntPag, nMoraPag, nOtrPag, nSaldoCap, " _
        & " nNumDesembN , nNumDesembR, nNumCanc, nNumRef, nNumJud, nNumSaldos, nNumPagos " _
        & " From ColocEstadDiaCred " _
        & " WHERE datediff(dd, destad, '" & Format(pdFecIni, "mm/dd/yyyy") & "') <= 0 " _
        & " AND datediff(dd, destad, '" & Format(pdFecFin, "mm/dd/yyyy") & "') >= 0) Det " _
        & " INNER JOIN AGENCIAS A ON DET.cCodAge = A.cAgeCod " _
        & " WHERE cCodAge in (" & psAgencias & ") AND SUBSTRING(cLineaCred, 5,1) = '" & pcMoneda & "' "

    If pnTipoReporte = 3 Then
        lsSQL = lsSQL & " AND cLineaCred = '" & pcLineaCredito & "' "
    End If

    lsSQL = lsSQL & " GROUP BY cAgeDescripcion, "

    If pnTipoReporte = 2 Or pnTipoReporte = 3 Then
        lsSQL = lsSQL & " cLineaCred, "
    End If

    lsSQL = lsSQL & " dEstad With ROLLUP "

    Set loDataRep = New dColPFunciones
    Set rs = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If Not (rs.EOF And rs.BOF) Then
            
        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("ESTADISTICA MENSUAL DE CREDITOS", lsCondiciones, lnPage, 186, cSubTit, 108202)
        lnLineas = 5
        
        lsCadImp = lsCadImp & "MONEDA : " & IIf(pcMoneda = "1", "SOLES", "DOLARES") & Chr(10)
        lsCadImp = lsCadImp & "AGENCIA : " & rs!sAgencia & Chr(10)
        If pnTipoReporte = 2 Or pnTipoReporte = 3 Then
            lsCadImp = lsCadImp & "LINEA   : " & rs!sLinea & Chr(10)
            lnLineas = lnLineas + 1
        End If
        lsCadImp = lsCadImp & String(191, "-") & Chr(10)
        lsCadImp = lsCadImp & "                |            D  E  S  E  M  B  O  L  S  O  S            |           P       A       G       O       S           |    TOTAL    |    SALDO    |    |    |    |    |    |    |      " & Chr(10)
        lsCadImp = lsCadImp & "   FECHA        |    NORMAL     REFINANCIA. | M.REFINANC.    M. JUDIC.  |   CAPITAL   |   INTERES   |    MORA     |    OTROS    |   COBRADO   |    CAPITAL  |NORM|REF.|CANC|REF.|JUD.|SAL.|PAG.  " & Chr(10)
        lsCadImp = lsCadImp & String(191, "-") & Chr(10)
        lnLineas = lnLineas + 6

        Do While Not rs.EOF
            
            'RaiseEvent Progress(rs.boomark, rs.RecordCount)
            If pnTipoReporte = 1 Then
                '**** AGENCIAS
                If rs!sEstad = "TOTAL" Then
                    lnTotalCobrado = rs!nCapPag + rs!nIntPag + rs!nMoraPag + rs!nOtrPag
                    lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                    lsCadImp = lsCadImp & ImpreFormat("SUBTOT AGENCIA", 14) & ImpreFormat(rs!nMontoDesembN, 10, 2) & _
                              Space(1) & ImpreFormat(rs!nMontoDesembR, 10, 2) & Space(1) & ImpreFormat(rs!nMontoRefinan, 10, 2) & _
                              Space(1) & ImpreFormat(rs!nMontoJudic, 10, 2) & Space(1) & ImpreFormat(rs!nCapPag, 10, 2) & _
                              Space(1) & ImpreFormat(rs!nIntPag, 10, 2) & Space(1) & ImpreFormat(rs!nMoraPag, 10, 2) & _
                              Space(1) & ImpreFormat(rs!nOtrPag, 10, 2) & Space(1) & ImpreFormat(lnTotalCobrado, 10, 2) & _
                              Space(1) & ImpreFormat(lnSaldoCap, 10, 2) & Space(1) & ImpreFormat(rs!nNumDesembN, 4, 0) & _
                              Space(1) & ImpreFormat(rs!nNumDesembR, 4, 0) & Space(1) & ImpreFormat(rs!nNumCanc, 4, 0) & _
                              Space(1) & ImpreFormat(rs!nNumRef, 4, 0) & Space(1) & ImpreFormat(rs!nNumJud, 4, 0) & _
                              Space(1) & ImpreFormat(lnCredSalCap, 4, 0) & Space(1) & ImpreFormat(rs!nNumPagos, 4, 0) & Chr(10)
                              lnSaldoCapTot = lnSaldoCapTot + lnSaldoCap
                              lnCredSalCapTot = lnCredSalCapTot + lnCredSalCap

                    lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                    lsCadImp = lsCadImp & Chr(10)
                    lnLineas = lnLineas + 4
                    rs.MoveNext
                    If rs.EOF Then
                        Exit Do
                    End If
                    If rs!sAgencia = "TOTAL" Then
                        lnTotalCobrado = rs!nCapPag + rs!nIntPag + rs!nMoraPag + rs!nOtrPag
                        lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                        lsCadImp = lsCadImp & ImpreFormat("     TOTAL ", 14) & ImpreFormat(rs!nMontoDesembN, 10, 2) & _
                                  Space(1) & ImpreFormat(rs!nMontoDesembR, 10, 2) & Space(1) & ImpreFormat(rs!nMontoRefinan, 10, 2) & _
                                  Space(1) & ImpreFormat(rs!nMontoJudic, 10, 2) & Space(1) & ImpreFormat(rs!nCapPag, 10, 2) & _
                                  Space(1) & ImpreFormat(rs!nIntPag, 10, 2) & Space(1) & ImpreFormat(rs!nMoraPag, 10, 2) & _
                                  Space(1) & ImpreFormat(rs!nOtrPag, 10, 2) & Space(1) & ImpreFormat(lnTotalCobrado, 10, 2) & _
                                  Space(1) & ImpreFormat(lnSaldoCapTot, 10, 2) & Space(1) & ImpreFormat(rs!nNumDesembN, 4, 0) & _
                                  Space(1) & ImpreFormat(rs!nNumDesembR, 4, 0) & Space(1) & ImpreFormat(rs!nNumCanc, 4, 0) & _
                                  Space(1) & ImpreFormat(rs!nNumRef, 4, 0) & Space(1) & ImpreFormat(rs!nNumJud, 4, 0) & _
                                  Space(1) & ImpreFormat(lnCredSalCapTot, 4, 0) & Space(1) & ImpreFormat(rs!nNumPagos, 4, 0) & Chr(10)
                        lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                        lsCadImp = lsCadImp & Chr(10)
                        lnLineas = lnLineas + 4
                        Exit Do
                    Else
                        lsCadImp = lsCadImp & "AGENCIA : " & rs!sAgencia & Chr(10)
                        lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                        lsCadImp = lsCadImp & "                |            D  E  S  E  M  B  O  L  S  O  S            |           P       A       G       O       S           |    TOTAL    |    SALDO    |    |    |    |    |    |    |      " & Chr(10)
                        lsCadImp = lsCadImp & "   FECHA        |    NORMAL     REFINANCIA. | M.REFINANC.    M. JUDIC.  |   CAPITAL   |   INTERES   |    MORA     |    OTROS    |   COBRADO   |    CAPITAL  |NORM|REF.|CANC|REF.|JUD.|SAL.|PAG.  " & Chr(10)
                        lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                        lnLineas = lnLineas + 5
                    End If
                End If
                
            Else
                '**** AGENCIAS
                If rs!sEstad = "TOTAL" Then
                    lnTotalCobrado = rs!nCapPag + rs!nIntPag + rs!nMoraPag + rs!nOtrPag
                    lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                    lsCadImp = lsCadImp & ImpreFormat("SUBTOT LINEA", 14) & ImpreFormat(rs!nMontoDesembN, 10, 2) & _
                              Space(1) & ImpreFormat(rs!nMontoDesembR, 10, 2) & Space(1) & ImpreFormat(rs!nMontoRefinan, 10, 2) & _
                              Space(1) & ImpreFormat(rs!nMontoJudic, 10, 2) & Space(1) & ImpreFormat(rs!nCapPag, 10, 2) & _
                              Space(1) & ImpreFormat(rs!nIntPag, 10, 2) & Space(1) & ImpreFormat(rs!nMoraPag, 10, 2) & _
                              Space(1) & ImpreFormat(rs!nOtrPag, 10, 2) & Space(1) & ImpreFormat(lnTotalCobrado, 10, 2) & _
                              Space(1) & ImpreFormat(lnSaldoCap, 10, 2) & Space(1) & ImpreFormat(rs!nNumDesembN, 4, 0) & _
                              Space(1) & ImpreFormat(rs!nNumDesembR, 4, 0) & Space(1) & ImpreFormat(rs!nNumCanc, 4, 0) & _
                              Space(1) & ImpreFormat(rs!nNumRef, 4, 0) & Space(1) & ImpreFormat(rs!nNumJud, 4, 0) & _
                              Space(1) & ImpreFormat(lnCredSalCap, 4, 0) & Space(1) & ImpreFormat(rs!nNumPagos, 4, 0) & Chr(10)
                              lnSaldoCapTot = lnSaldoCapTot + lnSaldoCap
                              lnCredSalCapTot = lnCredSalCapTot + lnCredSalCap
                    lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                    lsCadImp = lsCadImp & Chr(10)
                    lnLineas = lnLineas + 4
                    rs.MoveNext
                    If rs.EOF Then
                        Exit Do
                    End If
                
                    '**** LINEA DE CREDITO
                    If rs!sLinea = "TOTAL" Then
                        lnTotalCobrado = rs!nCapPag + rs!nIntPag + rs!nMoraPag + rs!nOtrPag
                        lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                        lsCadImp = lsCadImp & ImpreFormat("SUBTOT AGENCIA", 14) & ImpreFormat(rs!nMontoDesembN, 10, 2) & _
                                  Space(1) & ImpreFormat(rs!nMontoDesembR, 10, 2) & Space(1) & ImpreFormat(rs!nMontoRefinan, 10, 2) & _
                                  Space(1) & ImpreFormat(rs!nMontoJudic, 10, 2) & Space(1) & ImpreFormat(rs!nCapPag, 10, 2) & _
                                  Space(1) & ImpreFormat(rs!nIntPag, 10, 2) & Space(1) & ImpreFormat(rs!nMoraPag, 10, 2) & _
                                  Space(1) & ImpreFormat(rs!nOtrPag, 10, 2) & Space(1) & ImpreFormat(lnTotalCobrado, 10, 2) & _
                                  Space(1) & ImpreFormat(lnSaldoCap, 10, 2) & Space(1) & ImpreFormat(rs!nNumDesembN, 4, 0) & _
                                  Space(1) & ImpreFormat(rs!nNumDesembR, 4, 0) & Space(1) & ImpreFormat(rs!nNumCanc, 4, 0) & _
                                  Space(1) & ImpreFormat(rs!nNumRef, 4, 0) & Space(1) & ImpreFormat(rs!nNumJud, 4, 0) & _
                                  Space(1) & ImpreFormat(lnCredSalCap, 4, 0) & Space(1) & ImpreFormat(rs!nNumPagos, 4, 0) & Chr(10)
                                  lnSaldoCapTot = lnSaldoCapTot + lnSaldoCap
                                  lnCredSalCapTot = lnCredSalCapTot + lnCredSalCap
                        lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                        lsCadImp = lsCadImp & Chr(10)
                        lnLineas = lnLineas + 4
                        rs.MoveNext
                        If rs.EOF Then
                            Exit Do
                        End If
                        
                        If rs!sAgencia = "TOTAL" Then
                            lnTotalCobrado = rs!nCapPag + rs!nIntPag + rs!nMoraPag + rs!nOtrPag
                            lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                            lsCadImp = lsCadImp & ImpreFormat("    TOTAL ", 14) & ImpreFormat(rs!nMontoDesembN, 10, 2) & _
                                      Space(1) & ImpreFormat(rs!nMontoDesembR, 10, 2) & Space(1) & ImpreFormat(rs!nMontoRefinan, 10, 2) & _
                                      Space(1) & ImpreFormat(rs!nMontoJudic, 10, 2) & Space(1) & ImpreFormat(rs!nCapPag, 10, 2) & _
                                      Space(1) & ImpreFormat(rs!nIntPag, 10, 2) & Space(1) & ImpreFormat(rs!nMoraPag, 10, 2) & _
                                      Space(1) & ImpreFormat(rs!nOtrPag, 10, 2) & Space(1) & ImpreFormat(lnTotalCobrado, 10, 2) & _
                                      Space(1) & ImpreFormat(lnSaldoCapTot, 10, 2) & Space(1) & ImpreFormat(rs!nNumDesembN, 4, 0) & _
                                      Space(1) & ImpreFormat(rs!nNumDesembR, 4, 0) & Space(1) & ImpreFormat(rs!nNumCanc, 4, 0) & _
                                      Space(1) & ImpreFormat(rs!nNumRef, 4, 0) & Space(1) & ImpreFormat(rs!nNumJud, 4, 0) & _
                                      Space(1) & ImpreFormat(lnCredSalCapTot, 4, 0) & Space(1) & ImpreFormat(rs!nNumPagos, 4, 0) & Chr(10)
                            lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                            lsCadImp = lsCadImp & Chr(10)
                            lnLineas = lnLineas + 4
                            Exit Do
                        Else
                            lsCadImp = lsCadImp & "AGENCIA : " & rs!sAgencia & Chr(10)
                            lsCadImp = lsCadImp & "LINEA   : " & rs!sLinea & Chr(10)
                            lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                            lsCadImp = lsCadImp & "                |            D  E  S  E  M  B  O  L  S  O  S            |           P       A       G       O       S           |    TOTAL    |    SALDO    |    |    |    |    |    |    |      " & Chr(10)
                            lsCadImp = lsCadImp & "   FECHA        |    NORMAL     REFINANCIA. | M.REFINANC.    M. JUDIC.  |   CAPITAL   |   INTERES   |    MORA     |    OTROS    |   COBRADO   |    CAPITAL  |NORM|REF.|CANC|REF.|JUD.|SAL.|PAG.  " & Chr(10)
                            lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                            lnLineas = lnLineas + 6
                        End If
                    Else
                        lsCadImp = lsCadImp & "AGENCIA : " & rs!sAgencia & Chr(10)
                        lsCadImp = lsCadImp & "LINEA   : " & rs!sLinea & Chr(10)
                        lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                        lsCadImp = lsCadImp & "                |            D  E  S  E  M  B  O  L  S  O  S            |           P       A       G       O       S           |    TOTAL    |    SALDO    |    |    |    |    |    |    |      " & Chr(10)
                        lsCadImp = lsCadImp & "   FECHA        |    NORMAL     REFINANCIA. | M.REFINANC.    M. JUDIC.  |   CAPITAL   |   INTERES   |    MORA     |    OTROS    |   COBRADO   |    CAPITAL  |NORM|REF.|CANC|REF.|JUD.|SAL.|PAG.  " & Chr(10)
                        lsCadImp = lsCadImp & String(191, "-") & Chr(10)
                        lnLineas = lnLineas + 5
                    End If
                End If
            
            End If
                            
            lnTotalCobrado = rs!nCapPag + rs!nIntPag + rs!nMoraPag + rs!nOtrPag
            lsCadImp = lsCadImp & rs!sEstad & Space(6) & ImpreFormat(rs!nMontoDesembN, 10, 2) & _
                      Space(1) & ImpreFormat(rs!nMontoDesembR, 10, 2) & Space(1) & ImpreFormat(rs!nMontoRefinan, 10, 2) & _
                      Space(1) & ImpreFormat(rs!nMontoJudic, 10, 2) & Space(1) & ImpreFormat(rs!nCapPag, 10, 2) & _
                      Space(1) & ImpreFormat(rs!nIntPag, 10, 2) & Space(1) & ImpreFormat(rs!nMoraPag, 10, 2) & _
                      Space(1) & ImpreFormat(rs!nOtrPag, 10, 2) & Space(1) & ImpreFormat(lnTotalCobrado, 10, 2) & _
                      Space(1) & ImpreFormat(rs!nSaldoCap, 10, 2) & Space(1) & ImpreFormat(rs!nNumDesembN, 4, 0) & _
                      Space(1) & ImpreFormat(rs!nNumDesembR, 4, 0) & Space(1) & ImpreFormat(rs!nNumCanc, 4, 0) & _
                      Space(1) & ImpreFormat(rs!nNumRef, 4, 0) & Space(1) & ImpreFormat(rs!nNumJud, 4, 0) & _
                      Space(1) & ImpreFormat(rs!nNumSaldos, 4, 0) & Space(1) & ImpreFormat(rs!nNumPagos, 4, 0) & Chr(10)
                      lnSaldoCap = rs!nSaldoCap
                      lnCredSalCap = rs!nNumSaldos
                      lnLineas = lnLineas + 1
                      
                      If lnLineas >= 60 Then
                            lnPage = lnPage + 1
                            lsCadImp = lsCadImp & Chr(12)
                            lsCadImp = lsCadImp & nRepoCabecera("ESTADISTICA MENSUAL DE CREDITOS", lsCondiciones, lnPage, 186, cSubTit, 108202)
                            lnLineas = 5
                            lsCadImp = lsCadImp & "                |            D  E  S  E  M  B  O  L  S  O  S            |           P       A       G       O       S           |    TOTAL    |    SALDO    |    |    |    |    |    |    |      " & Chr(10)
                            lsCadImp = lsCadImp & "   FECHA        |    NORMAL     REFINANCIA. | M.REFINANC.    M. JUDIC.  |   CAPITAL   |   INTERES   |    MORA     |    OTROS    |   COBRADO   |    CAPITAL  |NORM|REF.|CANC|REF.|JUD.|SAL.|PAG.  " & Chr(10)
                            lsCadImp = lsCadImp & String(186, "-") & Chr(10)
                            lnLineas = lnLineas + 4
                      End If
            rs.MoveNext
        Loop
    End If

    Set rs = Nothing
    
    nRepo108202_EstadisticaMensualCreditos = lsCadImp 'lsCadBuffer

End Function

Public Function nRepo108203_CreditosVigentes_Arqueo(ByVal nTipoReporte As Byte, ByVal cSubTit As String, ByVal pdFecFin As String, ByVal cMoneda As String, ByVal cProductos As String, ByVal cCondicion As String, ByVal psAnalistas, _
ByVal psAgencias As String, ByVal nProtesto As Integer, ByVal bOrdenCreditos As Boolean) As String
  
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nitem As Integer
Dim nTNacionalAprob As Double
Dim nTExtranjeroAprob As Double
Dim nTNacionalSCapital As Double
Dim nTExtranjeroSCapital As Double
Dim loImprimeCab As NColPImpre
Dim sCliente As String * 35
Dim sMoneda As String * 7
Dim lsCondiciones As String
Dim sCriterio As String
Dim sTempoAnalista As String
Dim i As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
   
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

lsCondiciones = "Al " & pdFecFin

'**Modificado por DAOR 20070417, se agregó el join con la tabla ProductoTasaInteres
'**para obtener la tasa de interes compensatorio, además al cCtaCodAnt
'**en caso que sea vacio reemplazar por "9" para que se visualizen al ultimo

'---ARCV se modifico para mejorar el Performance no hacer (SELECT...) as Dato

lsSQL = "SELECT PP.cCtaCod,case dbo.GetCtaAntigua(Producto.cCtaCod) when '' then '9' else dbo.GetCtaAntigua(Producto.cCtaCod) end as cCtaCodAnt,PP.cPersCod, dbo.Persona.cPersNombre, SUBSTRING(dbo.Producto.cCtaCod, 9, 1) AS cMoneda, " & _
            " Constante.cConsDescripcion AS cDesMoneda, Colocaciones.nMontoCol, Producto.nSaldo AS nSaldoCapital, Colocaciones.dVigencia, " & _
            " ANA.cPersCod AS cAnalista,PERAN.cPersNombre AS cNomAnalista,RH.cUser AS cUserAnalista,ColocacCred.nDiasAtraso,PT.nTasaInteres " & _
            " FROM ProductoPersona PP INNER JOIN Producto ON PP.cCtaCod = Producto.cCtaCod AND PP.nPrdPersRelac=20 INNER JOIN " & _
            " Persona ON PP.cPersCod = Persona.cPersCod INNER JOIN Constante ON SUBSTRING(Producto.cCtaCod, 9, 1) = Constante.nConsValor INNER JOIN " & _
            " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod INNER JOIN " & _
            " ColocacCred ON Colocaciones.cCtaCod = ColocacCred.cCtaCod " & _
            " inner join ProductoTasaInteres PT on Producto.cCtaCod=PT.cCtaCod and PT.nPrdTasaInteres=" & gColocLineaCredTasasIntCompNormal & _
            "  " & _
            " INNER JOIN ProductoPersona ANA ON ANA.cCtaCod = Producto.cCtaCod AND ANA.nPrdPersRelac=28" & _
            " INNER JOIN Persona PERAN ON PERAN.cPersCod = ANA.cPersCod " & _
            " INNER JOIN RRHH RH ON RH.cPersCod = ANA.cPersCod " & _
            " WHERE (Producto.nPrdEstado IN ('" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "')) " & _
            " " '" AND (PP.nPrdPersRelac=" & gColRelPersTitular & ") "
        
            'Left Join RelCuentas RC on RC.cCtaCod=Producto.cCtaCod

            If Len(Trim(cMoneda)) > 0 Then
                lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & cMoneda & ")) "
            End If
            
            lsSQL = lsSQL & " AND (Constante.nConsCod=" & gMoneda & ") "
 
            If Len(Trim(cCondicion)) > 0 Then
                lsSQL = lsSQL & " AND (ColocacCred.nColocCondicion IN (" & cCondicion & ")) "
            End If
         
            If Len(Trim(cProductos)) > 0 Then
                lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 6, 3) IN (" & cProductos & "))"
            End If
 
            If Len(psAnalistas) > 0 Then
            'ARCV 22-06-2007
            '     lsSQL = lsSQL & " AND ((SELECT PRP.cPersCod FROM ProductoPersona PRP " & _
                               " WHERE PRP.CCTACOD = PP.cCtaCod AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") IN (" & psAnalistas & ")) "
                    lsSQL = lsSQL & " AND ANA.cPersCod IN (" & psAnalistas & ") "
            '----------
            End If
        
            'asgnando las areas
            lsSQL = lsSQL & " and Substring(Producto.cCtaCod,4,2) in ('" & psAgencias & "')"
            
            'asignado si existe o no protesto
            If nProtesto = 1 Then
                lsSQL = lsSQL & " and  Exists(Select *"
                lsSQL = lsSQL & " From ColocCalendDet"
                lsSQL = lsSQL & " Where cCtaCod=Producto.cCtaCod and nPrdConceptoCod in (1219,1220))"
            ElseIf nProtesto = 2 Then
                lsSQL = lsSQL & " and not Exists(Select *"
                lsSQL = lsSQL & " From ColocCalendDet"
                lsSQL = lsSQL & " Where cCtaCod=Producto.cCtaCod and nPrdConceptoCod in (1219,1220))"
            End If
            
        If nTipoReporte = 1 Then 'Ordenado por Codigo
            If bOrdenCreditos = True Then
                lsSQL = lsSQL & " ORDER BY PP.cCtaCod ASC "
            Else
                lsSQL = lsSQL & " ORDER BY cCtaCodAnt ASC "
            End If
            'lsSQL = lsSQL & " ORDER BY PP.cCtaCod ASC "
        ElseIf nTipoReporte = 2 Then 'Agrupado por Analista
            'Comentado por DAOR 20070418
            'lsSQL = lsSQL & " ORDER BY (SELECT Pers.cPersNombre FROM Persona Pers INNER JOIN " & _
            '                " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERS.CPERSCOD " & _
            '                " WHERE PRP.CCTACOD = PP.cCtaCod AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & "), " & _
            '                " Persona.cPersNombre"
            
            '**DAOR 20070418***********************************************
            'ARCV ------- NO USAR (SELECT ....)
            'lsSQL = lsSQL & " ORDER BY (SELECT Pers.cPersNombre FROM Persona Pers INNER JOIN " & _
                " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERS.CPERSCOD " & _
                " WHERE PRP.CCTACOD = PP.cCtaCod AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") "
            lsSQL = lsSQL & " ORDER BY PERAN.cPersNombre  "
            
            If bOrdenCreditos = True Then
                lsSQL = lsSQL & ", PP.cCtaCod ASC "
            Else
                lsSQL = lsSQL & ", cCtaCodAnt ASC "
            End If
            '**************************************************************
        ElseIf nTipoReporte = 3 Then 'Analista Esp
            'Comentado por DAOR 20070418
            'lsSQL = lsSQL & " ORDER BY (SELECT Pers.cPersNombre FROM Persona Pers INNER JOIN " & _
            '                " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERS.CPERSCOD " & _
            '                " WHERE PRP.CCTACOD = PP.cCtaCod AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & "), " & _
            '                " Persona.cPersNombre"
            
            '**DAOR 20070418***********************************************
            'ARCV 22-06-2007
            'lsSQL = lsSQL & " ORDER BY (SELECT Pers.cPersNombre FROM Persona Pers INNER JOIN " & _
                            " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERS.CPERSCOD " & _
                            " WHERE PRP.CCTACOD = PP.cCtaCod AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") "
            lsSQL = lsSQL & " ORDER BY PERAN.cPersNombre "
            If bOrdenCreditos = True Then
                lsSQL = lsSQL & ", PP.cCtaCod ASC "
            Else
                lsSQL = lsSQL & ", cCtaCodAnt ASC "
            End If
            '**************************************************************
        End If
 
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
        nTNacionalAprob = 0
        nTNacionalSCapital = 0
        nTExtranjeroAprob = 0
        nTExtranjeroSCapital = 0
        lnLineas = 0
        lnPage = 1
        nitem = 0

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CREDITOS VIGENTES - ARQUEO", lsCondiciones, lnPage, 139, cSubTit, 108203)
  
        lnIndice = 5:  lnLineas = 5
         
        sTempoAnalista = IIf(IsNull(lrDataRep!cAnalista), "", lrDataRep!cAnalista)

        If nTipoReporte = 2 Or nTipoReporte = 3 Then
            'Analista
            lsCadImp = lsCadImp & lsNegritaOn & " ANALISTA         : " & lrDataRep!cNomAnalista & " - " & lrDataRep!cUserAnalista & lsNegritaOff & Chr(10) & Chr(10)
            lnLineas = lnLineas + 2
            lnIndice = lnIndice + 2
        End If

        With lrDataRep
            RaiseEvent ShowProgress

            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
              
                'Totalizar por Analista
               If ((nTipoReporte = 2 Or nTipoReporte = 3) And sTempoAnalista <> lrDataRep!cAnalista) Then
                    lsCadImp = lsCadImp & String(139, "-") & Chr(10)
                    'Comentado por DAOR 20070418
                    'lsCadImp = lsCadImp & " TOTAL DE CASOS: " & ImpreFormat(nitem, 5, 0) & Space(34) & "R E S U M E N   P O R   M O N E D A" & Chr(10) & Space(56) & "SOLES  " & Space(1) & ImpreFormat(nTNacionalAprob, 10, 2) & Space(1) & ImpreFormat(nTNacionalSCapital, 10, 2) & Chr(10) & _
                    '            Space(56) & "DOLARES" & Space(1) & ImpreFormat(nTExtranjeroAprob, 10, 2) & Space(1) & ImpreFormat(nTExtranjeroSCapital, 10, 2) & Chr(10)
                    
                    '**DAOR 20070418******************************************************
                    lsCadImp = lsCadImp & " TOTAL DE CASOS: " & ImpreFormat(nitem, 5, 0) & Space(52) & "R E S U M E N   P O R   M O N E D A" & Chr(10)
                    lsCadImp = lsCadImp & Space(74) & "SOLES  " & Space(1) & ImpreFormat(nTNacionalAprob, 10, 2) & Space(1) & ImpreFormat(nTNacionalSCapital, 10, 2) & Chr(10)
                    lsCadImp = lsCadImp & Space(74) & "DOLARES" & Space(1) & ImpreFormat(nTExtranjeroAprob, 10, 2) & Space(1) & ImpreFormat(nTExtranjeroSCapital, 10, 2) & Chr(10)
                    '*********************************************************************

                    lsCadImp = lsCadImp & String(139, "-") & Chr(10)

                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & Chr(10)
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS VIGENTES - ARQUEO", lsCondiciones, lnPage, 139, cSubTit, 108203)

                    lnLineas = 5
                    lnIndice = lnIndice + 5

                    'Analista
                    lsCadImp = lsCadImp & lsNegritaOn & " ANALISTA         : " & lrDataRep!cNomAnalista & " - " & lrDataRep!cUserAnalista & lsNegritaOff & Chr(10) & Chr(10)
                    lnLineas = lnLineas + 2
                    lnIndice = lnIndice + 2
 
                    nitem = 0

                    nTNacionalAprob = 0
                    nTNacionalSCapital = 0
                    nTExtranjeroAprob = 0
                    nTExtranjeroSCapital = 0
                    sTempoAnalista = lrDataRep!cAnalista
                      
                    lnLineas = lnLineas + 6
                    lnIndice = lnIndice + 6
                    
                Else

                    If !cMoneda = gMonedaNacional Then
                        nTNacionalAprob = nTNacionalAprob + !nMontoCol
                        nTNacionalSCapital = nTNacionalSCapital + !nSaldoCapital
                    Else
                        nTExtranjeroAprob = nTExtranjeroAprob + !nMontoCol
                        nTExtranjeroSCapital = nTExtranjeroSCapital + !nSaldoCapital
                    End If

                End If
 
                nitem = nitem + 1

                sCliente = !cPersNombre
                sMoneda = !cDesMoneda
                
                'Comentado por DAOR 20070418
                'lsCadImp = lsCadImp & Space(1) & !cCtaCod & Space(1) & ImpreFormat(IIf(IsNull(!CctacodAnt) Or !CctacodAnt = "", "NUEVO CREDITO", !CctacodAnt), 23) & sCliente & Space(1) & sMoneda & Space(1) & _
                '           ImpreFormat(!nMontoCol, 10, 2) & Space(1) & ImpreFormat(!nSaldoCapital, 10, 2) & Space(1) & _
                '           Format(!dVigencia, "dd/MM/YYYY") & Space(1) & !cUserAnalista & Space(1) & ImpreFormat(!nDiasAtraso, 6, 0) & Chr(10)
                
                '**DAOR 20070418*********************************************************
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(!cCtaCod, 18, 0)
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(IIf(IsNull(!CctacodAnt) Or !CctacodAnt = "9", "NUEVO CREDITO", !CctacodAnt), 18, 0)
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(sCliente, 34, 0)
                lsCadImp = lsCadImp & Space(1) & sMoneda
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(!nMontoCol, 10, 2)
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(!nSaldoCapital, 10, 2)
                lsCadImp = lsCadImp & Space(1) & Format(!dVigencia, "dd/MM/YYYY")
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(!nTasaInteres, 2, 4)
                lsCadImp = lsCadImp & Space(1) & !cUserAnalista
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(!nDiasAtraso, 5, 0) & Chr(10)
                '************************************************************************


                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
 
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & Chr(10)
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS VIGENTES - ARQUEO", lsCondiciones, lnPage, 139, cSubTit, 108203)
                    lnLineas = 5
                    lnIndice = lnIndice + 5
                End If
 
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
 
            Loop

            RaiseEvent CloseProgress
 
        End With
 
        lsCadImp = lsCadImp & String(139, "-") & Chr(10)

        'Comentado por DAOR
        'lsCadImp = lsCadImp & " TOTAL DE CASOS: " & ImpreFormat(nitem, 5, 0) & Space(34) & "R E S U M E N   P O R   M O N E D A" & Chr(10) & Space(56) & "SOLES  " & Space(1) & ImpreFormat(nTNacionalAprob, 10, 2) & Space(1) & ImpreFormat(nTNacionalSCapital, 10, 2) & Chr(10) & _
        '        Space(56) & "DOLARES" & Space(1) & ImpreFormat(nTExtranjeroAprob, 10, 2) & Space(1) & ImpreFormat(nTExtranjeroSCapital, 10, 2) & Chr(10)
        
        '**DAOR 20070418******************************************************
        lsCadImp = lsCadImp & " TOTAL DE CASOS: " & ImpreFormat(nitem, 5, 0) & Space(52) & "R E S U M E N   P O R   M O N E D A" & Chr(10)
        lsCadImp = lsCadImp & Space(74) & "SOLES  " & Space(1) & ImpreFormat(nTNacionalAprob, 10, 2) & Space(1) & ImpreFormat(nTNacionalSCapital, 10, 2) & Chr(10)
        lsCadImp = lsCadImp & Space(74) & "DOLARES" & Space(1) & ImpreFormat(nTExtranjeroAprob, 10, 2) & Space(1) & ImpreFormat(nTExtranjeroSCapital, 10, 2) & Chr(10)
        '*********************************************************************
                 
        lsCadImp = lsCadImp & String(139, "-") & Chr(10)
 
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)

    End If

    nRepo108203_CreditosVigentes_Arqueo = lsCadBuffer

End Function
 
Public Function nRepo108204_CreditosDesembolsadosVigentes(ByVal cSubTit As String, ByVal pdFecIni As String, _
            ByVal pdFecFin As String, ByVal pnMontoTope As Currency, ByVal pcMoneda As String, _
            ByVal pcProductos As String, ByVal pcCondicion As String, ByVal pcAgencias As String, _
            ByVal pnTipoCambio As Currency, ByVal pdFec As String) As String
  
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nitem As Integer
Dim loImprimeCab As NColPImpre
Dim sCliente As String * 35
Dim sMoneda As String * 7
Dim lsCondiciones As String
Dim sCriterio As String
Dim i As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
      
Dim nTempoMoneda As Integer

 'Totales Generales
Dim sTotalA As Currency
Dim sTotalB As Currency
    
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)
 

lsSQL = "SELECT DISTINCT C1.cCtaCod as cCredito, C1.nTasaInteres, C1.cCliente, C1.nSaldoCapital, " & _
        " C1.nMoneda, C1.cMoneda, C2.nDesembolsado, isnull(C3.nTasacion,0) as nGarantiaP, " & _
        " C1.DiasAtraso, C1.dVigencia, C1.nSaldoCapital, C1.Analista " & _
        " FROM " & _
        " ( SELECT P.nTasaInteres, P.nSaldo as nSaldoCapital, CC.cCtaCod, SUBSTRING(C.cCtaCod, 6, 3) nProducto, " & _
        " (SELECT COO.cConsDescripcion FROM Constante COO " & _
        " WHERE Substring(C.cCtaCod, 6, 3) = COO.nConsValor AND COO.nConsCod=" & gProducto & ") AS cProducto, " & _
        " Persona.cPersNombre AS cCLiente, (SELECT RHH.CUSER FROM RRHH RHH INNER JOIN " & _
        " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD " & _
        " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS ANALISTA, " & _
        " (SELECT PRP.CPERSCOD FROM PRODUCTOPERSONA PRP " & _
        " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS CODANALISTA, " & _
        " (SELECT PERSO.cPersNombre FROM PERSONA PERSO INNER JOIN " & _
        " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD " & _
        " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.nPrdPersRelac =" & gColRelPersAnalista & ") AS cNomAnalista, " & _
        " Constante.cConsDescripcion AS cMoneda, SUBSTRING(C.cCtaCod, 9, 1) AS nMoneda, " & _
        " CC.nColocCondicion, CC.nDiasAtraso, C.dVigencia, " & _
        " DiasAtraso = (SELECT datediff(dd, dVenc, '" & Format(pdFec, "mm/dd/yyyy") & "') From ColocacCred C1 " & _
        " INNER JOIN ColocCalendario CC ON CC.cCtaCod = C1.cCtaCod And CC.nNroCalen = C1.nNroCalen " & _
        " AND CC.nCuota = C1.nNroProxCuota Where C1.cCtaCod = C.cCtaCod and nColocCalendApl = 1) "
 lsSQL = lsSQL & " FROM ColocacCred CC INNER JOIN Colocaciones C ON CC.cCtaCod = C.cCtaCod " & _
        " INNER JOIN Producto P ON C.cCtaCod = P.cCtaCod INNER JOIN ProductoPersona PP ON P.cCtaCod = PP.cCtaCod " & _
        " AND P.cCtaCod = PP.cCtaCod INNER JOIN ColocLineaCredito CL ON C.cLineaCred = CL.cLineaCred " & _
        " INNER JOIN Persona ON PP.cPersCod = Persona.cPersCod AND PP.cPersCod = Persona.cPersCod INNER JOIN " & _
        " Constante ON SUBSTRING(C.cCtaCod, 9, 1) = Constante.nConsValor " & _
        " WHERE (PP.nPrdPersRelac=" & gColRelPersTitular & ") "
        
        lsSQL = lsSQL & " AND (Constante.nConsCod= " & gMoneda & ") " & _
        " AND (P.nPrdEstado IN ('" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "')) "

        If Len(Trim(pcProductos)) > 0 Then
            lsSQL = lsSQL & " AND (SUBSTRING(C.cCtaCod, 6, 3) IN (" & pcProductos & "))"
        End If
        
        If Len(Trim(pcMoneda)) > 0 Then
            lsSQL = lsSQL & " AND (SUBSTRING(C.cCtaCod, 9, 1) IN (" & pcMoneda & ")) "
        End If

        If Len(Trim(pcAgencias)) > 0 Then
            lsSQL = lsSQL & " AND (SUBSTRING(C.cCtaCod, 4, 2) IN (" & pcAgencias & ")) "
        End If
            
        If Len(Trim(pcCondicion)) > 0 Then
            lsSQL = lsSQL & " AND (CC.nColocCondicion IN (" & pcCondicion & ")) "
        End If
        
        lsSQL = lsSQL & " ) C1 " & _
        " INNER JOIN " & _
            " (SELECT Distinct CC.cCtaCod, C.nPlazo, C.nMontoCol AS nDesembolsado " & _
            " FROM ColocacCred CC INNER JOIN Colocaciones C ON CC.cCtaCod = C.cCtaCod " & _
            " WHERE C.nMontoCol >=" & pnMontoTope & " AND C.dVigencia >= '" & Format(pdFecIni, "mm/dd/yyyy") & _
            "' AND C.dVigencia <= '" & Format(pdFecFin, "mm/dd/yyyy 23:59:59") & "' " & _
            " ) C2 ON C1.cCtaCod=C2.cCtaCod "
            
        lsSQL = lsSQL & " LEFT OUTER JOIN  (SELECT Distinct CC.cCtaCod, " & _
            " SUM(CASE WHEN SUBSTRING(CG.cCtaCod, 9,1) <> CG.nMoneda THEN " & _
            " (CASE CG.nMoneda WHEN 1 THEN CG.nGravado / " & pnTipoCambio & _
            " WHEN 2 THEN CG.nGravado * " & pnTipoCambio & " END) ELSE CG.nGravado END) nTasacion " & _
                " FROM ColocGarantia CG INNER JOIN Garantias G ON CG.cNumGarant = G.cNumGarant " & _
                " INNER JOIN Colocaciones C ON CG.cCtaCod = C.cCtaCod INNER JOIN " & _
                " ColocacCred CC ON C.cCtaCod = CC.cCtaCod Group By CC.cCtaCod " & _
                " " & _
                " ) C3 ON C1.cCtaCod=C3.cCtaCod " ' Where (dbo.Garantias.nTpoGarantia = 1)
 
    lsSQL = lsSQL & " ORDER BY C1.cMoneda, C2.nDesembolsado DESC, C1.dVigencia, C1.cCliente "
    
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        nTempoMoneda = 99
        nitem = 0
        
        With lrDataRep
            RaiseEvent ShowProgress
            
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nitem = nitem + 1
                     
                If nTempoMoneda <> !nmoneda Then
                    nitem = nitem - 1
                    If nTempoMoneda <> 99 Then
                        lsCadImp = lsCadImp & String(138, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & Space(20) & "*** CASOS POR MONEDA: " & Space(2) & _
                                    ImpreFormat(nitem, 7, 0) & Space(22) & ImpreFormat(sTotalA, 10, 2) & _
                                    Space(12) & ImpreFormat(sTotalB, 10, 2) & _
                                    lsNegritaOff & Chr(10) & Chr(10)
                        lsCadImp = lsCadImp & String(138, "-") & Chr(10)
                        lnPage = lnPage + 1
                        lsCadImp = lsCadImp & Chr(12)
                    End If
                    
                    lsCondiciones = "MAYORES A " & IIf(!nmoneda = gMonedaNacional, "S/. ", "US $") & ImpreFormat(pnMontoTope, 10, 2)
                    
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS DESEMBOLSADOS VIGENTES DEL " & pdFecIni & " AL " & pdFecFin, lsCondiciones, lnPage, 138, cSubTit, 108204)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                    
                    sTotalA = !nDesembolsado
                    sTotalB = !nSaldoCapital
                    
                    nTempoMoneda = !nmoneda
            
                    nitem = 1
                    
                Else
                    sTotalA = sTotalA + !nDesembolsado
                    sTotalB = sTotalB + !nSaldoCapital
                End If
                nTempoMoneda = lrDataRep!nmoneda
                
                sCliente = !cCliente
                sMoneda = !cMoneda
                
                lsCadImp = lsCadImp & Space(1) & !cCredito & Space(1) & sCliente & _
                           Space(1) & sMoneda & Space(1) & ImpreFormat(!nTasaInteres, 5, 4) & _
                           Space(1) & ImpreFormat(!nDesembolsado, 10, 2) & Space(1) & Format(!dVigencia, "dd/MM/YYYY") & _
                           Space(1) & ImpreFormat(!nSaldoCapital, 10, 2) & _
                           Space(1) & !ANALISTA & Space(1) & ImpreFormat(!DiasAtraso, 6, 0) & _
                           Space(1) & ImpreFormat(Format(!nGarantiaP, "#0.00"), 10, 2) & Chr(10)
                             
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCondiciones = "MAYORES A " & IIf(!nmoneda = gMonedaNacional, "S/. ", "US $") & ImpreFormat(pnMontoTope, 10, 2)
                    
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS DESEMBOLSADOS VIGENTES DEL " & pdFecIni & " AL " & pdFecFin, lsCondiciones, lnPage, 138, cSubTit, 108204)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
                
            Loop
            RaiseEvent CloseProgress
        End With
         
        lsCadImp = lsCadImp & String(138, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & Space(20) & "*** CASOS POR MONEDA: " & Space(2) & _
                    ImpreFormat(nitem, 7, 0) & Space(22) & ImpreFormat(sTotalA, 10, 2) & _
                    Space(12) & ImpreFormat(sTotalB, 10, 2) & _
                    lsNegritaOff & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & String(138, "-") & Chr(10)
        lnPage = lnPage + 1
        lsCadImp = lsCadImp & Chr(12)

        lsCadBuffer = lsCadBuffer & lsCadImp

     End If
    nRepo108204_CreditosDesembolsadosVigentes = lsCadBuffer

End Function


Public Function nRepo108301_ListadoProgramacionPagosCuota(ByVal nTipoReporte As Byte, ByVal cSubTit As String, ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal cMoneda As String, ByVal cProductos As String, ByVal cCondicion As String, ByVal psAnalistas, Optional psCadAge As Variant = "") As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lsCadImp2 As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nitem As Integer
Dim loImprimeCab As NColPImpre
        
Dim sCliente As String * 30
Dim lsCondiciones As String
Dim sCriterio As String
 
Dim sTempoAnalista As String
Dim sTempoGracia As String

Dim sTempoMoneda As Integer
Dim sTempoPlazo As Integer
Dim sTempoProducto As String
Dim sTempoFinanciamiento As String

Dim lsNegritaOn As String
Dim lsNegritaOff As String

'Totales por Moneda
Dim sTotalM1 As Currency
Dim sTotalM2 As Currency
Dim sTotalM3 As Currency
Dim sTotalM0 As Long

'Totales por Plazo
Dim sTotalP1 As Currency
Dim sTotalP2 As Currency
Dim sTotalP3 As Currency
Dim sTotalP0 As Long

'Totales por Producto
Dim sTotalR1 As Currency
Dim sTotalR2 As Currency
Dim sTotalR3 As Currency
Dim sTotalR0 As Long

'Totales por Fuente de Financiamiento
Dim sTotalF1 As Currency
Dim sTotalF2 As Currency
Dim sTotalF3 As Currency
Dim sTotalF0 As Long

'Totales Generales
Dim sTotalA As Currency
Dim sTotalB As Currency
Dim sTotalC As Currency
Dim sTotal0 As Long
Dim sMoneda As String
Dim nSalto As Integer
nSalto = 0
 
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)
 
    sCriterio = "  WHERE (C5.dVenc BETWEEN '" & Format(pdFecIni, "mm/dd/YYYY") & "' and '" & Format(pdFecFin, "MM/dd/YYYY") & "') "
                            
    lsCondiciones = "Del " & pdFecIni & " al " & pdFecFin
  

        lsSQL = " SELECT C1.cCtaCod as cCredito, C1.nFuente, C1.cFuente, C1.nPlazo, C1.cPlazo, C1.nProducto, C1.cProducto, C1.cCliente, C7.nCapital, C1.nMoneda, C1.cMoneda, C1.Analista, C1.CodAnalista, C1.cNomAnalista, " & _
              " C2.nDesembolsado AS nDesembolso, C3.nNroProxCuota as nCuotaPagar, C4.nTotalCuotas as nCuotasAprobadas, C6.Cuota as nCuota, C5.dVenc,C8.Periodo " & _
              " FROM " & _
              " ( " & _
              " SELECT     dbo.ColocacCred.cCtaCod, SUBSTRING(dbo.Colocaciones.cLineaCred, 1, 2) AS nFuente, dbo.ColocLineaCredito.cDescripcion AS cFuente, " & _
              " SUBSTRING(dbo.Colocaciones.cLineaCred, 6, 1) AS nPlazo, (SELECT CO.cconsdescripcion FROM Constante CO " & _
              " WHERE Co.nConsCOd=" & gColocLineaCredPlazo & " AND CO.nConsValor = SUBSTRING(dbo.Colocaciones.cLineaCred, 6, 1)) AS cPlazo, " & _
              " SUBSTRING(dbo.Colocaciones.cCtaCod, 6, 3) AS nProducto, (SELECT COO.cConsDescripcion FROM Constante COO " & _
              " WHERE Substring(Colocaciones.cCtaCod, 6, 3) = COO.nConsValor AND COO.nConsCod=" & gProducto & ") AS cProducto, " & _
              " dbo.Persona.cPersNombre AS cCLiente, (SELECT RHH.CUSER FROM RRHH RHH INNER JOIN " & _
              " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD WHERE PRP.CCTACOD = ProductoPersona.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS ANALISTA, " & _
              " (SELECT PRP.CPERSCOD FROM PRODUCTOPERSONA PRP WHERE PRP.CCTACOD = ProductoPersona.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS CODANALISTA, " & _
              " (SELECT PERSO.cPersNombre FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD WHERE PRP.CCTACOD = ProductoPersona.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS cNomAnalista, " & _
              " dbo.Constante.cConsDescripcion AS cMoneda, SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) AS nMoneda, dbo.ColocacCred.nColocCondicion FROM dbo.ColocacCred INNER JOIN " & _
              " dbo.Colocaciones ON dbo.ColocacCred.cCtaCod = dbo.Colocaciones.cCtaCod INNER JOIN dbo.Producto ON dbo.Colocaciones.cCtaCod = dbo.Producto.cCtaCod INNER JOIN " & _
              " dbo.ProductoPersona ON dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod AND dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod INNER JOIN " & _
              " dbo.ColocLineaCredito ON dbo.Colocaciones.cLineaCred = dbo.ColocLineaCredito.cLineaCred INNER JOIN dbo.Persona ON dbo.ProductoPersona.cPersCod = dbo.Persona.cPersCod AND dbo.ProductoPersona.cPersCod = dbo.Persona.cPersCod INNER JOIN " & _
              " dbo.Constante ON SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) = dbo.Constante.nConsValor " & _
              " WHERE (dbo.ProductoPersona.nPrdPersRelac=" & gColRelPersTitular & ") " & _
              " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
              " AND (dbo.Constante.nConsCod=" & gMoneda & ") "
              
              If Len(Trim(cProductos)) > 0 Then
                    lsSQL = lsSQL & " AND (SUBSTRING(dbo.Colocaciones.cCtaCod, 6, 3) IN (" & cProductos & ")) "
              End If
              
              If Len(Trim(psCadAge)) > 0 Then
                lsSQL = lsSQL & " AND (SUBSTRING(dbo.Colocaciones.cCtaCod, 4, 2) IN (" & psCadAge & ")) "
              End If
                
              If Len(Trim(psAnalistas)) > 0 Then
                    lsSQL = lsSQL & " AND ((SELECT PRP.CPERSCOD FROM PRODUCTOPERSONA PRP " & _
                        " WHERE PRP.CCTACOD = ProductoPersona.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") IN (" & psAnalistas & ")) "
              End If
              
              If Len(Trim(cMoneda)) > 0 Then
                    lsSQL = lsSQL & " AND (SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) IN (" & cMoneda & ")) "
              End If
              
              If Len(Trim(cCondicion)) > 0 Then
                    lsSQL = lsSQL & " AND (dbo.ColocacCred.nColocCondicion IN (" & cCondicion & ")) "
              End If
        
        lsSQL = lsSQL & " ) C1 " & _
                " Inner Join " & _
                " ( " & _
                " SELECT dbo.ColocacCred.cCtaCod, dbo.Colocaciones.nMontoCol AS nDesembolsado " & _
                " FROM dbo.ColocacCred INNER JOIN " & _
                " dbo.Colocaciones ON dbo.ColocacCred.cCtaCod = dbo.Colocaciones.cCtaCod " & _
                " ) C2 ON C1.cCtaCod=C2.cCtaCod " & _
                " Inner Join " & _
                " ( " & _
                " SELECT dbo.ColocacCred.cCtaCod, dbo.ColocacCred.nNroCalen, dbo.ColocacCred.nNroProxCuota " & _
                " FROM dbo.ColocacCred INNER JOIN dbo.Colocaciones ON dbo.ColocacCred.cCtaCod = dbo.Colocaciones.cCtaCod " & _
                " ) C3 " & _
                " on C2.cCtaCod=C3.cCtaCod " & _
                " Inner Join " & _
                " ( " & _
                " SELECT cCtaCod, nNroCalen, COUNT(nCuota) AS nTotalCuotas " & _
                " From dbo.ColocCalendario Where nColocCalendApl=" & gColocCalendAplCuota & " GROUP BY cCtaCod, nNroCalen " & _
                " ) C4 " & _
                " ON C3.cCtaCod = C4.cCtaCod AND C3.nNroCalen=C4.nNroCalen " & _
                " Inner Join " & _
                " ( " & _
                " SELECT cCtaCod, nNroCalen, nCuota, dVenc " & _
                " From dbo.ColocCalendario Where nColocCalendApl=" & gColocCalendAplCuota & _
                " ) C5 " & _
                " ON C3.cCtaCod = C5.cCtaCod And C3.nNrocalen = C5.nNrocalen And C3.nNroProxCuota = C5.nCuota "
        
        lsSQL = lsSQL & " Inner Join " & _
                " ( " & _
                " SELECT cCtaCod, nNroCalen, nCuota,  SUM(nMonto-nMontoPagado) AS Cuota From dbo.ColocCalendDet " & _
                " Where nColocCalendApl=" & gColocCalendAplCuota & _
                " GROUP BY cCtaCod, nNroCalen, nCuota " & _
                " ) C6 ON C3.cCtaCod = C6.cCtaCod And C3.nNrocalen = C6.nNrocalen And C3.nNroProxCuota = C6.nCuota "
                  
        lsSQL = lsSQL & "Inner Join (SELECT cCtaCod, dPrdEstado,nPrdEstado,nPeriodoGracia as Periodo From ColocacEstado WHERE nPrdEstado IN ('2030', '2031', '2032', '2020', '2021', '2022') ) C8 on C3.cCtaCod=C8.cCtaCod  AND C8.dPrdEstado=(SELECT MAX(dPrdEstado)FROM ColocacEstado WHERE cCtaCod = C8.cCtaCod AND nPrdEstado=C8.nPrdEstado) " & _
                " Inner Join (SELECT cCtaCod, nNroCalen, nCuota,  SUM(nMonto) AS nCapital " & _
                " From dbo.ColocCalendDet " & _
                " Where nColocCalendApl=" & gColocCalendAplCuota & " AND nPrdConceptoCod = " & gColocConceptoCodCapital & _
                " GROUP BY cCtaCod, nNroCalen, nCuota " & _
                " ) C7 ON C3.cCtaCod = C7.cCtaCod And C3.nNrocalen = C7.nNrocalen And C3.nNroProxCuota = C7.nCuota "
                
        lsSQL = lsSQL & sCriterio
             
    ''''''''''''''''''''''''''''''
             
            If nTipoReporte = 1 Then
                lsSQL = lsSQL & " order by c1.nMoneda, c1.nPlazo, c1.nProducto, c1.nFuente, c5.dvenc, c1.cCtaCod "
            ElseIf nTipoReporte = 2 Or nTipoReporte = 3 Then
                lsSQL = lsSQL & " order by c1.Analista, c1.nMoneda, c1.nPlazo, c1.nProducto, c1.nFuente, c5.dvenc, c1.cCtaCod "
            End If
            
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
          
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("PROGRAMACION DE PAGOS DE CUOTAS", lsCondiciones, lnPage, 133, cSubTit, 108301)
        lsCadImp = lsCadImp & Chr(10)
        
        lnIndice = 8:  lnLineas = 8
        
        sTempoMoneda = lrDataRep!nmoneda
        sTempoPlazo = lrDataRep!nPlazo
        sTempoProducto = lrDataRep!nProducto
        sTempoFinanciamiento = lrDataRep!nFuente
        sTempoAnalista = lrDataRep!ANALISTA
        sTempoGracia = IIf(IsNull(lrDataRep!Periodo), 0, lrDataRep!Periodo)
        
        If nTipoReporte = 2 Then
            'Analista
            lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA       : " & lrDataRep!ANALISTA & " - " & lrDataRep!cNomAnalista & lsNegritaOff & Chr(10) & Chr(10)
            lnLineas = lnLineas + 2
            lnIndice = lnIndice + 2
        End If
        
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA         : " & lrDataRep!cMoneda & lsNegritaOff & Chr(10) & Chr(10)
        'Plazo
        lsCadImp = lsCadImp & lsNegritaOn & "PLAZO          : " & lrDataRep!cPlazo & lsNegritaOff & Chr(10) & Chr(10)
        'Producto
        lsCadImp = lsCadImp & lsNegritaOn & "PRODUCTO       : " & lrDataRep!cProducto & " : " & lrDataRep!nProducto & lsNegritaOff & Chr(10) & Chr(10)
        'Fuente de Financiamiento
        lsCadImp = lsCadImp & lsNegritaOn & "FINANCIAMIENTO : " & lrDataRep!cFuente & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(141, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(1) & "ITEM CREDITO            CLIENTE                         DESEMBOLSADO         CUOTA       CAPITAL CUOTAS  CUOTA  FECHA DE  ANALISTA  GRACIA" & Chr(10)
        lsCadImp = lsCadImp & Space(1) & "                                                                                                 APROB.  PAGAR VENCIMIENT" & Chr(10)
        lsCadImp = lsCadImp & String(141, "-") & Chr(10)
        sMoneda = lrDataRep!cMoneda
        lnLineas = lnLineas + 11
        
        lnIndice = lnIndice + 11
          If nTipoReporte = 3 Then
            'If sTempoAnalista <> lrDataRep!Analista Then
            lsCadImp2 = lsCadImp2 & nRepoCabecera("CAIDA DE CARTERA DEL " & Format(pdFecIni, "mm/YYYY"), lsCondiciones, lnPage, 133, cSubTit, 108301)
            lsCadImp2 = lsCadImp2 & Chr(10)
            lsCadImp2 = lsCadImp2 & ImpreFormat("CODANA", 10) & ImpreFormat("ANALISTA", 30) & ImpreFormat("CASOS S/.", 10) & ImpreFormat("MONTO S/.", 10) & ImpreFormat("CASOS $", 10) & ImpreFormat("MONTO $", 15) & Chr(10)
            'lsCadImp2 = lsCadImp2 & ImpreFormat("CODANA", 10) & lrDataRep!Analista & " - " & lrDataRep!cNomAnalista & Chr(10)
            'End If
        End If
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nitem = nitem + 1
                
                'Totalizar por Moneda
                                   
                If (nTipoReporte = 1 And sTempoMoneda <> !nmoneda) Or (((nTipoReporte = 2 Or nTipoReporte = 3) And sTempoAnalista <> lrDataRep!ANALISTA) Or ((nTipoReporte = 2 Or nTipoReporte = 3) And sTempoMoneda <> !nmoneda)) Then
                
                    lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL FUENTE FINANCIAMIENTO:  CASOS: " & Space(5) & ImpreFormat(sTotalF0, 3, 0) & _
                                Space(1) & ImpreFormat(sTotalF1, 10, 2) & Space(1) & ImpreFormat(sTotalF2, 10, 2) & _
                                Space(1) & ImpreFormat(sTotalF3, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                
                    lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL PRODUCTO             :  CASOS: " & Space(5) & ImpreFormat(sTotalR0, 3, 0) & _
                                Space(1) & ImpreFormat(sTotalR1, 10, 2) & Space(1) & ImpreFormat(sTotalR2, 10, 2) & _
                                Space(1) & ImpreFormat(sTotalR3, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                    
                    lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL PLAZO                :  CASOS: " & Space(5) & ImpreFormat(sTotalP0, 3, 0) & _
                                Space(1) & ImpreFormat(sTotalP1, 10, 2) & Space(1) & ImpreFormat(sTotalP2, 10, 2) & _
                                Space(1) & ImpreFormat(sTotalP3, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                     
                    lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL MONEDA               :  CASOS: " & Space(5) & ImpreFormat(sTotalM0, 3, 0) & _
                                Space(1) & ImpreFormat(sTotalM1, 10, 2) & Space(1) & ImpreFormat(sTotalM2, 10, 2) & _
                                Space(1) & ImpreFormat(sTotalM3, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
'''''                    If !nmoneda = "" Then
'''''
'''''                    Else
'''''                    End If
                    'ALPA 26062008***********************************
                    If sTempoAnalista <> lrDataRep!ANALISTA Then
''                    lsCadImp2 = lsCadImp2 & Chr(10)
''                    lsCadImp2 = lsCadImp2 & lsNegritaOn & "ANALISTA       : " & lrDataRep!Analista & " - " & lrDataRep!cNomAnalista & Chr(10)
                     If lrDataRep!nmoneda = 1 Then
                            If nSalto = 1 Then
                               lsCadImp2 = lsCadImp2 & Chr(10)
                            End If
                            lsCadImp2 = lsCadImp2 & ImpreFormat(lrDataRep!ANALISTA, 10) & ImpreFormat(lrDataRep!cNomAnalista, 30) & ImpreFormat(sTotalM0, 10) & ImpreFormat(sTotalM3, 10)
                            nSalto = 1
                      Else
                            If nSalto = 1 Then
                               lsCadImp2 = lsCadImp2 & Chr(10)
                            End If
                            lsCadImp2 = lsCadImp2 & ImpreFormat(lrDataRep!ANALISTA, 10) & ImpreFormat(lrDataRep!cNomAnalista, 30) & ImpreFormat("", 10) & ImpreFormat(" ", 10) & ImpreFormat(sTotalM0, 10) & ImpreFormat(sTotalM3, 15)
                            nSalto = 0
                      End If
                    Else
                    If lrDataRep!nmoneda = 2 Then
                      lsCadImp2 = lsCadImp2 & ImpreFormat(sTotalM0, 10) & ImpreFormat(sTotalM3, 15) & Chr(10)
                    End If
                    End If
                    
'                    lsCadImp2 = lsCadImp2 & ImpreFormat("CASOS :", 10) & ImpreFormat(sTotalM0, 3, 0) & Space(1) & ImpreFormat(sTotalM3, 10, 2) & ImpreFormat(lrDataRep!cMoneda, 15) & Chr(10)
                    '************************************************
                    If (nTipoReporte = 2 And sTempoAnalista <> lrDataRep!ANALISTA) Then
                        'Analista
                        lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA       : " & lrDataRep!ANALISTA & " - " & lrDataRep!cNomAnalista & lsNegritaOff & Chr(10) & Chr(10)
                        lnLineas = lnLineas + 2
                        lnIndice = lnIndice + 2
                    End If
                                
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA         : " & lrDataRep!cMoneda & lsNegritaOff & Chr(10) & Chr(10)
                    'Plazo
                    lsCadImp = lsCadImp & lsNegritaOn & "PLAZO          : " & lrDataRep!cPlazo & lsNegritaOff & Chr(10) & Chr(10)
                    'Producto
                    lsCadImp = lsCadImp & lsNegritaOn & "PRODUCTO       : " & lrDataRep!cProducto & " : " & lrDataRep!nProducto & lsNegritaOff & Chr(10) & Chr(10)
                    'Fuente de Financiamiento
                    lsCadImp = lsCadImp & lsNegritaOn & "FINANCIAMIENTO: " & lrDataRep!cFuente & lsNegritaOff & Chr(10)
                    lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                    lsCadImp = lsCadImp & Space(1) & "ITEM CREDITO            CLIENTE                         DESEMBOLSADO         CUOTA       CAPITAL CUOTAS  CUOTA  FECHA DE  ANALISTA" & Chr(10)
                    lsCadImp = lsCadImp & Space(1) & "                                                                                                 APROB.  PAGAR VENCIMIENT" & Chr(10)
                    lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                    
                    sTotalM0 = 1
                    sTotalM1 = !nDesembolso
                    sTotalM2 = !nCuota
                    sTotalM3 = !nCapital
                     
                    sTotalP0 = 1
                    sTotalP1 = !nDesembolso
                    sTotalP2 = !nCuota
                    sTotalP3 = !nCapital
                    
                    sTotalR0 = 1
                    sTotalR1 = !nDesembolso
                    sTotalR2 = !nCuota
                    sTotalR3 = !nCapital
                    
                    sTotalF0 = 1
                    sTotalF1 = !nDesembolso
                    sTotalF2 = !nCuota
                    sTotalF3 = !nCapital
                         
                    nitem = 1
                    
                    sTempoMoneda = lrDataRep!nmoneda
                    sTempoPlazo = lrDataRep!nPlazo
                    sTempoProducto = lrDataRep!nProducto
                    sTempoFinanciamiento = lrDataRep!nFuente
                      
                    lnLineas = lnLineas + 19
                    lnIndice = lnIndice + 19
                     
                Else 'si es la misma moneda
                    sTotalM0 = sTotalM0 + 1
                    sTotalM1 = sTotalM1 + !nDesembolso
                    sTotalM2 = sTotalM2 + !nCuota
                    sTotalM3 = sTotalM3 + !nCapital
                    
                    If sTempoPlazo <> !nPlazo Then
                
                        lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL FUENTE FINANCIAMIENTO:  CASOS: " & Space(5) & ImpreFormat(sTotalF0, 3, 0) & _
                                    Space(3) & ImpreFormat(sTotalF1, 10, 2) & Space(1) & ImpreFormat(sTotalF2, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotalF3, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                    
                        lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL PRODUCTO             :  CASOS: " & Space(5) & ImpreFormat(sTotalR0, 3, 0) & _
                                    Space(3) & ImpreFormat(sTotalR1, 10, 2) & Space(1) & ImpreFormat(sTotalR2, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotalR3, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                        
                        lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL PLAZO                :  CASOS: " & Space(5) & ImpreFormat(sTotalP0, 3, 0) & _
                                    Space(3) & ImpreFormat(sTotalP1, 10, 2) & Space(1) & ImpreFormat(sTotalP2, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotalP3, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                        
                        'Plazo
                        lsCadImp = lsCadImp & lsNegritaOn & "PLAZO          : " & lrDataRep!cPlazo & lsNegritaOff & Chr(10) & Chr(10)
                        'Producto
                        lsCadImp = lsCadImp & lsNegritaOn & "PRODUCTO       : " & lrDataRep!cProducto & " : " & lrDataRep!nProducto & lsNegritaOff & Chr(10) & Chr(10)
                        'Fuente de Financiamiento
                        lsCadImp = lsCadImp & lsNegritaOn & "FINANCIAMIENTO: " & lrDataRep!cFuente & lsNegritaOff & Chr(10)
                        lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                        lsCadImp = lsCadImp & Space(1) & "ITEM CREDITO            CLIENTE                         DESEMBOLSADO         CUOTA       CAPITAL CUOTAS  CUOTA  FECHA DE  ANALISTA   GRACIA" & Chr(10)
                        lsCadImp = lsCadImp & Space(1) & "                                                                                                 APROB.  PAGAR VENCIMIENT" & Chr(10)
                        lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                        
                        sTotalP0 = 1
                        sTotalP1 = !nDesembolso
                        sTotalP2 = !nCuota
                        sTotalP3 = !nCapital
                        
                        sTotalR0 = 1
                        sTotalR1 = !nDesembolso
                        sTotalR2 = !nCuota
                        sTotalR3 = !nCapital
                        
                        sTotalF0 = 1
                        sTotalF1 = !nDesembolso
                        sTotalF2 = !nCuota
                        sTotalF3 = !nCapital
                        
                        nitem = 1
                        
                        sTempoPlazo = lrDataRep!nPlazo
                        sTempoProducto = lrDataRep!nProducto
                        sTempoFinanciamiento = lrDataRep!nFuente
                    
                        lnLineas = lnLineas + 15
                        lnIndice = lnIndice + 15
                    Else 'si la moneda y el plazo son iguales
                          
                        sTotalP0 = sTotalP0 + 1
                        sTotalP1 = sTotalP1 + !nDesembolso
                        sTotalP2 = sTotalP2 + !nCuota
                        sTotalP3 = sTotalP3 + !nCapital
                        
                        If sTempoProducto <> !nProducto Then
                    
                            lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                            lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL FUENTE FINANCIAMIENTO:  CASOS: " & Space(5) & ImpreFormat(sTotalF0, 3, 0) & _
                                        Space(3) & ImpreFormat(sTotalF1, 10, 2) & Space(1) & ImpreFormat(sTotalF2, 10, 2) & _
                                        Space(1) & ImpreFormat(sTotalF3, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                        
                            lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                            lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL PRODUCTO             :  CASOS: " & Space(5) & ImpreFormat(sTotalR0, 3, 0) & _
                                        Space(3) & ImpreFormat(sTotalR1, 10, 2) & Space(1) & ImpreFormat(sTotalR2, 10, 2) & _
                                        Space(1) & ImpreFormat(sTotalR3, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                            
                    
                            'Producto
                            lsCadImp = lsCadImp & lsNegritaOn & "PRODUCTO       : " & lrDataRep!cProducto & " : " & lrDataRep!nProducto & lsNegritaOff & Chr(10) & Chr(10)
                            'Fuente de Financiamiento
                            lsCadImp = lsCadImp & lsNegritaOn & "FINANCIAMIENTO: " & lrDataRep!cFuente & lsNegritaOff & Chr(10)
                            lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                            lsCadImp = lsCadImp & Space(1) & "ITEM CREDITO            CLIENTE                         DESEMBOLSADO         CUOTA       CAPITAL CUOTAS  CUOTA  FECHA DE  ANALISTA" & Chr(10)
                            lsCadImp = lsCadImp & Space(1) & "                                                                                                 APROB.  PAGAR VENCIMIENT" & Chr(10)
                            lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                            
                            sTotalR0 = 1
                            sTotalR1 = !nDesembolso
                            sTotalR2 = !nCuota
                            sTotalR3 = !nCapital
                            
                            sTotalF0 = 1
                            sTotalF1 = !nDesembolso
                            sTotalF2 = !nCuota
                            sTotalF3 = !nCapital
                            
                            nitem = 1
                            
                            sTempoProducto = lrDataRep!nProducto
                            sTempoFinanciamiento = lrDataRep!nFuente
                        
                            lnLineas = lnLineas + 11
                            lnIndice = lnIndice + 11
                        Else
                            sTotalR0 = sTotalR0 + 1
                            sTotalR1 = sTotalR1 + !nDesembolso
                            sTotalR2 = sTotalR2 + !nCuota
                            sTotalR3 = sTotalR3 + !nCapital
                            
                            If sTempoFinanciamiento <> !nFuente Then
                        
                                lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                                lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL FUENTE FINANCIAMIENTO:  CASOS: " & Space(5) & ImpreFormat(sTotalF0, 3, 0) & _
                                            Space(3) & ImpreFormat(sTotalF1, 10, 2) & Space(1) & ImpreFormat(sTotalF2, 10, 2) & _
                                            Space(1) & ImpreFormat(sTotalF3, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)

                                 
                                'Fuente de Financiamiento
                                lsCadImp = lsCadImp & lsNegritaOn & "FINANCIAMIENTO: " & lrDataRep!cFuente & lsNegritaOff & Chr(10)
                                lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                                lsCadImp = lsCadImp & Space(1) & "ITEM CREDITO            CLIENTE                         DESEMBOLSADO         CUOTA       CAPITAL CUOTAS  CUOTA  FECHA DE  ANALISTA" & Chr(10)
                                lsCadImp = lsCadImp & Space(1) & "                                                                                                 APROB.  PAGAR VENCIMIENT" & Chr(10)
                                lsCadImp = lsCadImp & String(141, "-") & Chr(10)
                                 
                                sTotalF0 = 1
                                sTotalF1 = !nDesembolso
                                sTotalF2 = !nCuota
                                sTotalF3 = !nCapital
                                nitem = 1
                                
                                sTempoFinanciamiento = lrDataRep!nFuente
                            
                                lnLineas = lnLineas + 7
                                lnIndice = lnIndice + 7
                            Else
                                sTotalF0 = sTotalF0 + 1
                                sTotalF1 = sTotalF1 + !nDesembolso
                                sTotalF2 = sTotalF2 + !nCuota
                                sTotalF3 = sTotalF3 + !nCapital
                            End If
                        End If
                    End If
                End If
                
                sTotal0 = sTotal0 + 1
                sTotalA = sTotalA + !nDesembolso
                sTotalB = sTotalB + !nCuota
                sTotalC = sTotalC + !nCapital
                sTempoAnalista = !ANALISTA
                sTempoGracia = IIf(IsNull(!Periodo), 0, !Periodo)
                sCliente = PstaNombre(!cCliente)
                 
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(nitem, 4, 0) & Space(1) & !cCredito & Space(1) & _
                           sCliente & Space(1) & ImpreFormat(!nDesembolso, 10, 2) & Space(1) & ImpreFormat(!nCuota, 10, 2) & _
                           Space(1) & ImpreFormat(!nCapital, 10, 2) & Space(1) & ImpreFormat(!nCuotasAprobadas, 6, 0) & _
                           Space(1) & ImpreFormat(!nCuotaPagar, 6, 0) & Space(1) & Format(!dVenc, "dd/MM/YYYY") & _
                           Space(3) & !ANALISTA & Space(8) & !Periodo & Chr(10)
                
                 
                'lnLineas = lnLineas + 1
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & Chr(10)
                    lsCadImp = lsCadImp & nRepoCabecera("PROGRAMACION DE PAGOS DE CUOTAS", lsCondiciones, lnPage, 133, cSubTit, 108301)
                    lnLineas = 8
                    lnIndice = lnIndice + 8
                End If
                
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
                
            Loop
            RaiseEvent CloseProgress
                
        End With
        
        lsCadImp = lsCadImp & String(141, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL FUENTE FINANCIAMIENTO:  CASOS: " & Space(3) & ImpreFormat(sTotalF0, 3, 0) & _
                    Space(3) & ImpreFormat(sTotalF1, 10, 2) & Space(1) & ImpreFormat(sTotalF2, 10, 2) & _
                    Space(1) & ImpreFormat(sTotalF3, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
    
        lsCadImp = lsCadImp & String(141, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL PRODUCTO             :  CASOS: " & Space(3) & ImpreFormat(sTotalR0, 3, 0) & _
                    Space(3) & ImpreFormat(sTotalR1, 10, 2) & Space(1) & ImpreFormat(sTotalR2, 10, 2) & _
                    Space(1) & ImpreFormat(sTotalR3, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
        
        lsCadImp = lsCadImp & String(141, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL PLAZO                :  CASOS: " & Space(3) & ImpreFormat(sTotalP0, 3, 0) & _
                    Space(3) & ImpreFormat(sTotalP1, 10, 2) & Space(1) & ImpreFormat(sTotalP2, 10, 2) & _
                    Space(1) & ImpreFormat(sTotalP3, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
         
        lsCadImp = lsCadImp & String(141, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & " ** SUB TOTAL MONEDA               :  CASOS: " & Space(3) & ImpreFormat(sTotalM0, 3, 0) & _
                    Space(3) & ImpreFormat(sTotalM1, 10, 2) & Space(1) & ImpreFormat(sTotalM2, 10, 2) & _
                    Space(1) & ImpreFormat(sTotalM3, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
        'ALPA 26062008***********************************
        lsCadImp2 = lsCadImp2 & ImpreFormat("CASOS :", 10) & ImpreFormat(sTotalM0, 3, 0) & Space(1) & ImpreFormat(sTotalM3, 10, 2) & ImpreFormat(sMoneda, 15) & Chr(10)
        '************************************************
               
        If nTipoReporte = 3 Then
        lsCadBuffer = ""
        lsCadBuffer = lsCadBuffer & lsCadImp2 & Chr(12)
        Else
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        End If
    End If
    nRepo108301_ListadoProgramacionPagosCuota = lsCadBuffer
End Function
Public Function nRepo108301_ListadoProgramacionResumen(ByVal nTipoReporte As Byte, ByVal cSubTit As String, ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal cMoneda As String, ByVal cProductos As String, ByVal cCondicion As String, ByVal psAnalistas, Optional psCadAge As Variant = "") As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lsCadImp2 As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nitem As Integer
Dim loImprimeCab As NColPImpre
        
Dim sCliente As String * 30
Dim lsCondiciones As String
Dim sCriterio As String
 
Dim sTempoAnalista As String
Dim sTempoGracia As String

Dim sTempoMoneda As Integer
Dim sTempoPlazo As Integer
Dim sTempoProducto As String
Dim sTempoFinanciamiento As String

Dim lsNegritaOn As String
Dim lsNegritaOff As String

'Totales por Moneda
Dim sTotalM1 As Currency
Dim sTotalM2 As Currency
Dim sTotalM3 As Currency
Dim sTotalM0 As Long

'Totales por Plazo
Dim sTotalP1 As Currency
Dim sTotalP2 As Currency
Dim sTotalP3 As Currency
Dim sTotalP0 As Long

'Totales por Producto
Dim sTotalR1 As Currency
Dim sTotalR2 As Currency
Dim sTotalR3 As Currency
Dim sTotalR0 As Long

'Totales por Fuente de Financiamiento
Dim sTotalF1 As Currency
Dim sTotalF2 As Currency
Dim sTotalF3 As Currency
Dim sTotalF0 As Long

'Totales Generales
Dim sTotalA As Currency
Dim sTotalB As Currency
Dim sTotalC As Currency
Dim sTotal0 As Long
Dim sMoneda As String
Dim nSalto As Integer
nSalto = 0
 
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)
 
    sCriterio = "  WHERE (C5.dVenc BETWEEN '" & Format(pdFecIni, "mm/dd/YYYY") & "' and '" & Format(pdFecFin, "MM/dd/YYYY") & "') "
                            
    lsCondiciones = "Del " & pdFecIni & " al " & pdFecFin
  
            lsSQL = "select x.Analista, x.CodAnalista, x.cNomAnalista,"
            lsSQL = lsSQL & "sum(x.nCantSoles) nCantSoles,"
            lsSQL = lsSQL & "sum(x.nMonedaSoles) nMonedaSoles,"
            lsSQL = lsSQL & "sum(x.nCantDolar) nCantDolar,"
            lsSQL = lsSQL & "sum(x.nMonedaDolares) nMonedaDolares "
            lsSQL = lsSQL & "from ("
            lsSQL = lsSQL & " SELECT C1.Analista, C1.CodAnalista, C1.cNomAnalista, " & _
              " case C1.nMoneda when 1 then count(C1.cCtaCod) else 0 end nCantSoles," & _
              " sum(case C1.nMoneda when 1 then C7.nCapital else 0 end) nMonedaSoles," & _
              " case C1.nMoneda when 2 then count(C1.cCtaCod) else 0 end nCantDolar," & _
              " sum(case C1.nMoneda when 2 then C7.nCapital else 0 end) nMonedaDolares " & _
              " FROM " & _
              " ( " & _
              " SELECT     dbo.ColocacCred.cCtaCod, SUBSTRING(dbo.Colocaciones.cLineaCred, 1, 2) AS nFuente, dbo.ColocLineaCredito.cDescripcion AS cFuente, " & _
              " SUBSTRING(dbo.Colocaciones.cLineaCred, 6, 1) AS nPlazo, (SELECT CO.cconsdescripcion FROM Constante CO " & _
              " WHERE Co.nConsCOd=" & gColocLineaCredPlazo & " AND CO.nConsValor = SUBSTRING(dbo.Colocaciones.cLineaCred, 6, 1)) AS cPlazo, " & _
              " SUBSTRING(dbo.Colocaciones.cCtaCod, 6, 3) AS nProducto, (SELECT COO.cConsDescripcion FROM Constante COO " & _
              " WHERE Substring(Colocaciones.cCtaCod, 6, 3) = COO.nConsValor AND COO.nConsCod=" & gProducto & ") AS cProducto, " & _
              " dbo.Persona.cPersNombre AS cCLiente, (SELECT RHH.CUSER FROM RRHH RHH INNER JOIN " & _
              " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD WHERE PRP.CCTACOD = ProductoPersona.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS ANALISTA, " & _
              " (SELECT PRP.CPERSCOD FROM PRODUCTOPERSONA PRP WHERE PRP.CCTACOD = ProductoPersona.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS CODANALISTA, " & _
              " (SELECT PERSO.cPersNombre FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD WHERE PRP.CCTACOD = ProductoPersona.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS cNomAnalista, " & _
              " dbo.Constante.cConsDescripcion AS cMoneda, SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) AS nMoneda, dbo.ColocacCred.nColocCondicion FROM dbo.ColocacCred INNER JOIN " & _
              " dbo.Colocaciones ON dbo.ColocacCred.cCtaCod = dbo.Colocaciones.cCtaCod INNER JOIN dbo.Producto ON dbo.Colocaciones.cCtaCod = dbo.Producto.cCtaCod INNER JOIN " & _
              " dbo.ProductoPersona ON dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod AND dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod INNER JOIN " & _
              " dbo.ColocLineaCredito ON dbo.Colocaciones.cLineaCred = dbo.ColocLineaCredito.cLineaCred INNER JOIN dbo.Persona ON dbo.ProductoPersona.cPersCod = dbo.Persona.cPersCod AND dbo.ProductoPersona.cPersCod = dbo.Persona.cPersCod INNER JOIN " & _
              " dbo.Constante ON SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) = dbo.Constante.nConsValor " & _
              " WHERE (dbo.ProductoPersona.nPrdPersRelac=" & gColRelPersTitular & ") " & _
              " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
              " AND (dbo.Constante.nConsCod=" & gMoneda & ") "
              
              If Len(Trim(cProductos)) > 0 Then
                    lsSQL = lsSQL & " AND (SUBSTRING(dbo.Colocaciones.cCtaCod, 6, 3) IN (" & cProductos & ")) "
              End If
              
              If Len(Trim(psCadAge)) > 0 Then
                lsSQL = lsSQL & " AND (SUBSTRING(dbo.Colocaciones.cCtaCod, 4, 2) IN (" & psCadAge & ")) "
              End If
                
              If Len(Trim(psAnalistas)) > 0 Then
                    lsSQL = lsSQL & " AND ((SELECT PRP.CPERSCOD FROM PRODUCTOPERSONA PRP " & _
                        " WHERE PRP.CCTACOD = ProductoPersona.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") IN (" & psAnalistas & ")) "
              End If
              
              If Len(Trim(cMoneda)) > 0 Then
                    lsSQL = lsSQL & " AND (SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) IN (" & cMoneda & ")) "
              End If
              
              If Len(Trim(cCondicion)) > 0 Then
                    lsSQL = lsSQL & " AND (dbo.ColocacCred.nColocCondicion IN (" & cCondicion & ")) "
              End If
        
        lsSQL = lsSQL & " ) C1 " & _
                " Inner Join " & _
                " ( " & _
                " SELECT dbo.ColocacCred.cCtaCod, dbo.Colocaciones.nMontoCol AS nDesembolsado " & _
                " FROM dbo.ColocacCred INNER JOIN " & _
                " dbo.Colocaciones ON dbo.ColocacCred.cCtaCod = dbo.Colocaciones.cCtaCod " & _
                " ) C2 ON C1.cCtaCod=C2.cCtaCod " & _
                " Inner Join " & _
                " ( " & _
                " SELECT dbo.ColocacCred.cCtaCod, dbo.ColocacCred.nNroCalen, dbo.ColocacCred.nNroProxCuota " & _
                " FROM dbo.ColocacCred INNER JOIN dbo.Colocaciones ON dbo.ColocacCred.cCtaCod = dbo.Colocaciones.cCtaCod " & _
                " ) C3 " & _
                " on C2.cCtaCod=C3.cCtaCod " & _
                " Inner Join " & _
                " ( " & _
                " SELECT cCtaCod, nNroCalen, COUNT(nCuota) AS nTotalCuotas " & _
                " From dbo.ColocCalendario Where nColocCalendApl=" & gColocCalendAplCuota & " GROUP BY cCtaCod, nNroCalen " & _
                " ) C4 " & _
                " ON C3.cCtaCod = C4.cCtaCod AND C3.nNroCalen=C4.nNroCalen " & _
                " Inner Join " & _
                " ( " & _
                " SELECT cCtaCod, nNroCalen, nCuota, dVenc " & _
                " From dbo.ColocCalendario Where nColocCalendApl=" & gColocCalendAplCuota & _
                " ) C5 " & _
                " ON C3.cCtaCod = C5.cCtaCod And C3.nNrocalen = C5.nNrocalen And C3.nNroProxCuota = C5.nCuota "
        
        lsSQL = lsSQL & " Inner Join " & _
                " ( " & _
                " SELECT cCtaCod, nNroCalen, nCuota,  SUM(nMonto-nMontoPagado) AS Cuota From dbo.ColocCalendDet " & _
                " Where nColocCalendApl=" & gColocCalendAplCuota & _
                " GROUP BY cCtaCod, nNroCalen, nCuota " & _
                " ) C6 ON C3.cCtaCod = C6.cCtaCod And C3.nNrocalen = C6.nNrocalen And C3.nNroProxCuota = C6.nCuota "
                  
        lsSQL = lsSQL & "Inner Join (SELECT cCtaCod, dPrdEstado,nPrdEstado,nPeriodoGracia as Periodo From ColocacEstado WHERE nPrdEstado IN ('2030', '2031', '2032', '2020', '2021', '2022') ) C8 on C3.cCtaCod=C8.cCtaCod  AND C8.dPrdEstado=(SELECT MAX(dPrdEstado)FROM ColocacEstado WHERE cCtaCod = C8.cCtaCod AND nPrdEstado=C8.nPrdEstado) " & _
                " Inner Join (SELECT cCtaCod, nNroCalen, nCuota,  SUM(nMonto) AS nCapital " & _
                " From dbo.ColocCalendDet " & _
                " Where nColocCalendApl=" & gColocCalendAplCuota & " AND nPrdConceptoCod = " & gColocConceptoCodCapital & _
                " GROUP BY cCtaCod, nNroCalen, nCuota " & _
                " ) C7 ON C3.cCtaCod = C7.cCtaCod And C3.nNrocalen = C7.nNrocalen And C3.nNroProxCuota = C7.nCuota "
                
        lsSQL = lsSQL & sCriterio
        lsSQL = lsSQL & " group by C1.Analista, C1.CodAnalista, C1.cNomAnalista,  C1.nMoneda "
        lsSQL = lsSQL & " ) x "
        lsSQL = lsSQL & "group by x.Analista, x.CodAnalista, x.cNomAnalista "
        lsSQL = lsSQL & "order by x.Analista "
    
            
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
          
        lnLineas = 0
        lnPage = 1
        
        lnIndice = 8:  lnLineas = 8
        
        lnLineas = lnLineas + 11
        
        lnIndice = lnIndice + 11
            lsCadImp2 = lsCadImp2 & nRepoCabecera("CAIDA DE CARTERA DEL " & Format(pdFecIni, "mm/YYYY"), lsCondiciones, lnPage, 133, cSubTit, 108301)
            lsCadImp2 = lsCadImp2 & Chr(10)
            lsCadImp2 = lsCadImp2 & ImpreFormat("CODANA", 10) & ImpreFormat("ANALISTA", 30) & ImpreFormat("CASOS S/.", 10) & ImpreFormat("MONTO S/.", 10) & ImpreFormat("CASOS $", 10) & ImpreFormat("MONTO $", 15) & Chr(10)
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nitem = nitem + 1
                    lsCadImp2 = lsCadImp2 & ImpreFormat(lrDataRep!ANALISTA, 10) & ImpreFormat(lrDataRep!cNomAnalista, 30) & ImpreFormat(lrDataRep!nCantSoles, 10, 0) & ImpreFormat(lrDataRep!nMonedaSoles, 10) & ImpreFormat(lrDataRep!nCantDolar, 10, 0) & ImpreFormat(lrDataRep!nMonedaDolares, 15) & Chr(10)
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp2 = lsCadImp2 & Chr(12)
                    lsCadImp2 = lsCadImp2 & Chr(10)
                    lsCadImp2 = lsCadImp2 & nRepoCabecera("PROGRAMACION DE PAGOS DE CUOTAS", lsCondiciones, lnPage, 133, cSubTit, 108301)
                    lnLineas = 8
                    lnIndice = lnIndice + 8
                End If
                
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            RaiseEvent CloseProgress
        End With
       
        lsCadBuffer = lsCadBuffer & lsCadImp2 & Chr(12)
    End If
    nRepo108301_ListadoProgramacionResumen = lsCadBuffer
End Function

Public Function nRepo108302_VisitaCobroCuotas(ByVal cSubTit As String, ByVal pdFecFin As String, ByVal cMoneda As String, _
                ByVal cProductos As String, ByVal cCondicion As String, ByVal psAnalistas, Optional psCodRepo As String) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nitem As Integer
Dim nTNacionalAprob As Double
Dim nTExtranjeroAprob As Double
Dim nTNacionalSCapital As Double
Dim nTExtranjeroSCapital As Double
Dim loImprimeCab As NColPImpre
        
Dim sCliente As String * 30
Dim sDirecNeg As String * 25
Dim sZonaCli As String * 15

Dim lsCondiciones As String
Dim sCriterio As String
 
Dim sTempoAnalista As String
Dim sTempoMoneda As String

Dim sTempoCred As String

Dim sSCapitalAnalista As Double
Dim sCuotaAnalista As Double
Dim sMoraAnalista As Double

Dim sSCapitalMoneda As Double
Dim sCuotaMoneda As Double
Dim sMoraMoneda As Double

Dim lsNegritaOn As String
Dim lsNegritaOff As String
   
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)
    
    lsCondiciones = "Al " & pdFecFin
  
    lsSQL = "select c1.cCtaCod, c1.cMoneda, c1.cConsDescripcion, c1.cPersNombre, c1.cAnalista, c1.cNomAnalista, " & _
            " c1.cUserAnalista, c1.cRazSocDirecc, c1.cUbiGeoDescripcion, c1.nSaldo, c2.nCuota , C2.sCuota, C2.sMora, " & _
            " C2.nDiasAtraso, C2.dVenc From " & _
            "  (  SELECT P.cCtaCod, SUBSTRING(P.cCtaCod, 6, 3) AS cProducto, SUBSTRING(P.cCtaCod, 9, 1) AS cMoneda, " & _
            " Cons.cConsDescripcion, Per.cPersNombre, (SELECT PRP.cPersCod FROM ProductoPersona PRP " & _
            " WHERE PRP.CCTACOD = PPer.cCtaCod AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS cAnalista, (SELECT Pers.cPersNombre " & _
            " FROM Persona Pers INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERS.CPERSCOD " & _
            " WHERE PRP.CCTACOD = PPer.cCtaCod AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS cNomAnalista, (SELECT RHH.CUSER " & _
            " FROM RRHH RHH INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD " & _
            " WHERE PRP.CCTACOD = PPer.cCtaCod AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS cUserAnalista, " & _
            " ISNULL(FteIng.cRazSocDirecc, '') AS cRazSocDirecc, ISNULL(Ubi.cUbiGeoDescripcion, '') AS cUbiGeoDescripcion, " & _
            " P.nSaldo FROM  "
    
'ARCV 25-06-2007
'    lsSQL = lsSQL & " UbicacionGeografica Ubi INNER JOIN dbo.PersFteIngreso FteIng ON Ubi.cUbiGeoCod = FteIng.cRazSocUbiGeo " & _
            " " & _
            " RIGHT OUTER JOIN ProductoPersona PPer INNER JOIN Producto P ON PPer.cCtaCod = P.cCtaCod AND PPer.cCtaCod = P.cCtaCod INNER JOIN " & _
            " Constante Cons ON SUBSTRING(P.cCtaCod, 9, 1) = Cons.nConsValor INNER JOIN Persona Per ON PPer.cPersCod = Per.cPersCod " & _
            " AND PPer.cPersCod = Per.cPersCod INNER JOIN Colocaciones Coloc ON P.cCtaCod = Coloc.cCtaCod INNER JOIN " & _
            " ColocacCred ColocCre ON Coloc.cCtaCod = ColocCre.cCtaCod ON FteIng.cNumFuente = ColocCre.cNumFte "
    lsSQL = lsSQL & "     Producto P INNER JOIN ProductoPersona PPer ON PPer.cCtaCod = P.cCtaCod  AND nPrdPersRelac= 20 " & _
                "INNER JOIN  ColocFteIngreso CFE ON CFE.cCtaCod = P.cCtaCod " & _
        "AND CFE.dPersEval=(SELECT MAX(dPersEval)FROM ColocFteIngreso WHERE cCtaCod=P.cCtaCod ) " & _
        "INNER JOIN PersFteIngreso FteIng ON FteIng.cNumFuente = CFE.cNumFuente " & _
        "INNER JOIN  UbicacionGeografica Ubi ON Ubi.cUbiGeoCod = FteIng.cRazSocUbiGeo " & _
        "INNER JOIN  Constante Cons ON SUBSTRING(P.cCtaCod, 9, 1) = Cons.nConsValor " & _
        "INNER JOIN Persona Per ON PPer.cPersCod = Per.cPersCod  AND PPer.cPersCod = Per.cPersCod " & _
        "INNER JOIN Colocaciones Coloc ON P.cCtaCod = Coloc.cCtaCod " & _
        "INNER JOIN  ColocacCred ColocCre ON Coloc.cCtaCod = ColocCre.cCtaCod  "

'---------
lsSQL = lsSQL & " Where (Cons.nConsCod =" & gMoneda & ") " & _
                " AND (P.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
                " And (PPer.nPrdPersRelac =" & gColRelPersTitular & ") "
            
            If Len(Trim(cProductos)) > 0 Then
                lsSQL = lsSQL & " and (SUBSTRING(P.cCtaCod, 6, 3) IN (" & cProductos & "))"
            End If
             
            If Len(psAnalistas) > 0 Then
                lsSQL = lsSQL & " AND ((SELECT PRP.cPersCod FROM ProductoPersona PRP " & _
                " WHERE PRP.CCTACOD = PPer.cCtaCod AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") IN (" & psAnalistas & ")) "
            End If
            
            If Len(Trim(cMoneda)) > 0 Then
                lsSQL = lsSQL & " AND (SUBSTRING(P.cCtaCod, 9, 1) IN (" & cMoneda & ")) "
            End If
            
            If Len(Trim(cCondicion)) > 0 Then
                lsSQL = lsSQL & " AND (ColocCre.nColocCondicion IN (" & cCondicion & ") "
            End If
                                  
            lsSQL = lsSQL & "  ))   C1 "
    
    lsSQL = lsSQL & " Inner Join " & _
            "  (  SELECT    CCal.cCtaCod, CCal.nCuota, (select sum(nMonto - nMontoPagado) as sCuota from coloccalenddet CCD " & _
            " Where CCD.cCtaCod = CCal.cCtaCod and CCD.nNroCalen=CCal.nNroCalen and CCD.nCUota=CCal.nCuota " & _
            " AND nColocCalendApl=" & gColocCalendAplCuota & _
            " group by cCtaCod) as sCuota, " & _
            " (select sum(nMonto - nMontoPagado) as sCuota from coloccalenddet CCD " & _
            " Where CCD.cCtaCod = CCal.cCtaCod " & _
            " and CCD.nNroCalen=CCal.nNroCalen " & _
            " and CCD.nCUota=CCal.nCuota " & _
            " and CCD.nPrdConceptoCod IN('" & gColocConceptoCodInteresMoratorio & "')" & _
            " AND nColocCalendApl=" & gColocCalendAplCuota & _
            " group by cCtaCod) as sMora, " & _
            " DATEDIFF(day, CCal.dVenc, '" & Format(pdFecFin, "MM/dd/YYYY") & "') AS nDiasAtraso, CCal.dVenc " & _
            " FROM dbo.ColocCalendario CCal INNER JOIN " & _
            " dbo.ColocacCred CCred ON CCal.cCtaCod = CCred.cCtaCod AND CCal.nNroCalen = CCred.nNroCalen " & _
            " Where (CCal.nColocCalendEstado=" & gColocCalendEstadoPendiente & ") AND CCal.dVenc < '" & Format(pdFecFin, "mm/dd/yyyy") & "' " & _
            "  ) C2 " & _
            " ON C1.cCtaCod=C2.cCtaCod ORDER BY  C1.cNomAnalista, SUBSTRING(C1.cCtaCod,9,1), C1.cCtaCod, C2.nCuota"
     
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO PARA VISITA A CLIENTES - COBRO DE CUOTAS PENDIENTES ", lsCondiciones, lnPage, 160, cSubTit, psCodRepo)

        lnIndice = 7:  lnLineas = 7
        sTempoAnalista = lrDataRep!cAnalista
        sTempoMoneda = lrDataRep!cMoneda
        
        sTempoCred = "xx"
        
        'Analista
        lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA: " & lrDataRep!cNomAnalista & " - " & lrDataRep!cUserAnalista & lsNegritaOff & Chr(10) & Chr(10)
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cConsDescripcion & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(160, "-") & Chr(10)
        lsCadImp = lsCadImp & " CREDITO            CLIENTE                             DIR. DEL NEGOCIO          ZONA            SALDO      SALDO     CUOTA  DIAS    SALDO MORA   FECHA DE " & Chr(10)
        lsCadImp = lsCadImp & "                                                                                                 CAPITAL     CUOTA     MORA  ATRASO               VENCIMIENT" & Chr(10)
        lsCadImp = lsCadImp & String(160, "-") & Chr(10)
        lnLineas = lnLineas + 7
        lnIndice = lnIndice + 7
        
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nitem = nitem + 1
                
                If sTempoAnalista <> !cAnalista Then
                
                    lsCadImp = lsCadImp & String(160, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(75) & ImpreFormat(sSCapitalAnalista, 10, 2) & _
                                Space(1) & ImpreFormat(sCuotaAnalista, 8, 2) & Space(14) & ImpreFormat(sMoraAnalista, 8, 2) & _
                                  lsNegritaOff & Chr(10) & Chr(10)
                                        
 '                   lsCadImp = lsCadImp & String(160, "-") & Chr(10)
 '                   lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ANALISTA" & Space(73) & ImpreFormat(sSCapitalMoneda, 10, 2) & _
 '                               Space(1) & ImpreFormat(sCuotaMoneda, 8, 2) & Space(14) & ImpreFormat(sMoraMoneda, 8, 2) & lsNegritaOff & Chr(10) & Chr(10)
                      
                    'Analista
                    lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA: " & lrDataRep!cNomAnalista & " - " & lrDataRep!cUserAnalista & lsNegritaOff & Chr(10) & Chr(10)
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cConsDescripcion & lsNegritaOff & Chr(10)
                    lsCadImp = lsCadImp & String(160, "-") & Chr(10)
                    lsCadImp = lsCadImp & " CREDITO            CLIENTE                             DIR. DEL NEGOCIO          ZONA            SALDO      SALDO     CUOTA  DIAS    SALDO MORA   FECHA DE " & Chr(10)
                    lsCadImp = lsCadImp & "                                                                                                 CAPITAL     CUOTA     MORA  ATRASO               VENCIMIENT" & Chr(10)
                    lsCadImp = lsCadImp & String(160, "-") & Chr(10)
                      
                    sSCapitalAnalista = !nSaldo
                    sCuotaAnalista = !sCuota
                    sMoraAnalista = !sMora
                     
                    sSCapitalMoneda = !nSaldo
                    sCuotaMoneda = !sCuota
                    sMoraMoneda = !sMora
                     
                    nitem = 1
                    
                    sTempoAnalista = lrDataRep!cAnalista
                    sTempoMoneda = lrDataRep!cMoneda
                              
                    lnLineas = lnLineas + 10
                    lnIndice = lnIndice + 10
                    
                Else
                    If sTempoCred <> lrDataRep!cCtaCod Then
                        sSCapitalMoneda = sSCapitalMoneda + !nSaldo
                    End If
                    sCuotaMoneda = sCuotaMoneda + IIf(IsNull(!nCuota), 0, !nCuota)
                    sMoraMoneda = sMoraMoneda + IIf(IsNull(!sMora), 0, !sMora)
                     
                    If sTempoMoneda <> !cMoneda Then
                    
                        lsCadImp = lsCadImp & String(160, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(75) & ImpreFormat(sSCapitalAnalista, 10, 2) & _
                                Space(1) & ImpreFormat(sCuotaAnalista, 8, 2) & Space(14) & ImpreFormat(sMoraAnalista, 8, 2) & _
                                  lsNegritaOff & Chr(10) & Chr(10)
                     
                        'Moneda
                        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cConsDescripcion & lsNegritaOff & Chr(10) & Chr(10)
                        lsCadImp = lsCadImp & String(160, "-") & Chr(10)
                         
                        sSCapitalAnalista = !nSaldo
                        sCuotaAnalista = IIf(IsNull(!sCuota), 0, !sCuota)
                        sMoraAnalista = IIf(IsNull(!sMora), 0, !sMora)
                        nitem = 1
                        
                        sTempoMoneda = lrDataRep!cMoneda
                        
                        lnLineas = lnLineas + 6
                        lnIndice = lnIndice + 6
                        
                    Else
                        If sTempoCred <> lrDataRep!cCtaCod Then  'Aca Cambio
                            sSCapitalAnalista = sSCapitalAnalista + !nSaldo
                        End If
                        sCuotaAnalista = sCuotaAnalista + IIf(IsNull(!sCuota), 0, !sCuota)
                        sMoraAnalista = sMoraAnalista + IIf(IsNull(!sMora), 0, !sMora)
                    End If
                End If
                
                 
                sCliente = PstaNombre(!cPersNombre)
                sDirecNeg = !cRazSocDirecc
                sZonaCli = !cUbiGeoDescripcion
                     
                If sTempoCred = lrDataRep!cCtaCod Then
                    lsCadImp = lsCadImp & Space(105) & _
                        ImpreFormat(IIf(IsNull(!sCuota), 0, !sCuota), 8, 2) & Space(1) & ImpreFormat(!nCuota, 5, 0) & Space(2) & ImpreFormat(!nDiasAtraso, 4, 0) & _
                        Space(3) & ImpreFormat(IIf(IsNull(!sMora), 0, !sMora), 7, 2) & Space(4) & Format(!dVenc, "dd/MM/YYYY") & Chr(10)
                
                Else
                    sTempoCred = lrDataRep!cCtaCod
                    
                    lsCadImp = lsCadImp & !cCtaCod & Space(1) & sCliente & Space(1) & sDirecNeg & _
                            Space(1) & sZonaCli & Space(1) & ImpreFormat(!nSaldo, 9, 2) & Space(1) & ImpreFormat(IIf(IsNull(!sCuota), 0, !sCuota), 8, 2) & _
                            Space(1) & ImpreFormat(!nCuota, 5, 0) & Space(2) & ImpreFormat(!nDiasAtraso, 4, 0) & _
                            Space(3) & ImpreFormat(IIf(IsNull(!sMora), 0, !sMora), 7, 2) & Space(4) & Format(!dVenc, "dd/MM/YYYY") & Chr(10)
                            
                End If
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO PARA VISITA A CLIENTES - COBRO DE CUOTAS PENDIENTES ", lsCondiciones, lnPage, 173, cSubTit, psCodRepo)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            RaiseEvent CloseProgress
        End With
         
        lsCadImp = lsCadImp & String(160, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(75) & ImpreFormat(sSCapitalAnalista, 10, 2) & _
                    Space(1) & ImpreFormat(sCuotaAnalista, 8, 2) & Space(14) & ImpreFormat(sMoraAnalista, 8, 2) & _
                      lsNegritaOff & Chr(10) & Chr(10)
                            
 '       lsCadImp = lsCadImp & String(160, "-") & Chr(10)
  '      lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ANALISTA" & Space(73) & ImpreFormat(sSCapitalMoneda, 10, 2) & _
  '                  Space(1) & ImpreFormat(sCuotaMoneda, 8, 2) & Space(14) & ImpreFormat(sMoraMoneda, 8, 2) & lsNegritaOff & Chr(10) & Chr(10)
        
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
     
    nRepo108302_VisitaCobroCuotas = lsCadBuffer
End Function

Public Function nRepo108303_ClientesCuotasPend(ByVal cSubTit As String, ByVal pcMoneda As String, ByVal pcProductos As String, _
        ByVal pcCondicion As String, ByVal pcAnalistas, ByVal pnNota1 As Integer, ByVal pnNota2 As Integer, _
        ByVal pnTipoCuotas As Integer, ByVal pnCuotas As Integer, Optional ByVal psCodRepo As String = "", Optional nPorPagado As Double = 0) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nitem As Integer
Dim nTNacionalAprob As Double
Dim nTExtranjeroAprob As Double
Dim nTNacionalSCapital As Double
Dim nTExtranjeroSCapital As Double
Dim loImprimeCab As NColPImpre
Dim sCliente As String * 40
Dim sDirecNeg As String * 35
Dim sZonaCli As String * 15
Dim lsCondiciones As String
Dim sCriterio As String
Dim sTempoAnalista As String
Dim sTempoMoneda As String
Dim sTempoCred As String
Dim i As Integer

Dim sSCapitalAnalista As Double
Dim sCuotaAnalista As Double
Dim sMoraAnalista As Double
 
Dim sSCapitalMoneda As Double
 
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim nPorPaga As Double
Dim nPorPaga2 As Double
nPorPaga = 0
   
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)
    
    lsCondiciones = " CON NOTAS " & pnNota1 & " - " & pnNota2
  
    lsSQL = "SELECT P.cCtaCod, P.cPersNombre, P.cRazSocDirecc, P.cRazSocTelef, P.cUbiGeoDescripcion, P.nSaldo, "
     
    lsSQL = lsSQL & " P.nDiasAtraso,  " & _
      " P.cMoneda, P.cDesMoneda,   "
        
    lsSQL = lsSQL & " Cuot.nCuotasTotal, Cuot.nCuotasPend, cNotas.nColocNota, "
       
    lsSQL = lsSQL & " An.cAnalista, An.cDesAnalista ,Cuot.NroCuotas,(Cuot.nNroProxCuota-1) nNroProxCuota FROM " & _
        " (" & _
        " SELECT DISTINCT " & _
        " dbo.Producto.cCtaCod, Producto.nSaldo, dbo.ProductoPersona.cPersCod, dbo.ColocacCred.nDiasAtraso, dbo.Persona.cPersNombre, SUBSTRING(dbo.Producto.cCtaCod, 9, 1) AS cMoneda, dbo.Constante.cConsDescripcion AS cDesMoneda, SUBSTRING(dbo.Producto.cCtaCod, 6, 3) AS cProducto, " & _
        " dbo.ColocacCred.nNroCalen, ISNULL(dbo.PersFteIngreso.cRazSocDirecc, '') AS cRazSocDirecc, ISNULL(dbo.PersFteIngreso.cRazSocTelef, '') AS cRazSocTelef, ISNULL(dbo.UbicacionGeografica.cUbiGeoDescripcion, '') AS cUbiGeoDescripcion "
        
'ARCV 25-06-2007
'lsSQL = lsSQL & " FROM dbo.UbicacionGeografica INNER JOIN dbo.PersFteIngreso ON dbo.UbicacionGeografica.cUbiGeoCod = dbo.PersFteIngreso.cRazSocUbiGeo RIGHT OUTER JOIN " & _
            " dbo.Producto INNER JOIN dbo.ProductoPersona ON dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod INNER JOIN dbo.Colocaciones ON dbo.Producto.cCtaCod = dbo.Colocaciones.cCtaCod INNER JOIN " & _
            " dbo.ColocacCred ON dbo.Colocaciones.cCtaCod = dbo.ColocacCred.cCtaCod INNER JOIN dbo.Persona ON dbo.ProductoPersona.cPersCod = dbo.Persona.cPersCod INNER JOIN dbo.Constante ON SUBSTRING(dbo.Producto.cCtaCod, 9, 1) = dbo.Constante.nConsValor ON " & _
            " dbo.PersFteIngreso.cNumFuente = dbo.ColocacCred.cNumFte "
    lsSQL = lsSQL & "  FROM Producto INNER JOIN ProductoPersona ON ProductoPersona.cCtaCod = Producto.cCtaCod  " & _
            "AND ProductoPersona.nPrdPersRelac= 20 " & _
        "INNER JOIN  ColocFteIngreso ON ColocFteIngreso.cCtaCod = Producto.cCtaCod " & _
        "AND ColocFteIngreso.dPersEval=(SELECT MAX(dPersEval)FROM ColocFteIngreso WHERE cCtaCod=Producto.cCtaCod ) " & _
        "INNER JOIN PersFteIngreso ON PersFteIngreso.cNumFuente = ColocFteIngreso.cNumFuente " & _
        "INNER JOIN  UbicacionGeografica ON UbicacionGeografica.cUbiGeoCod = PersFteIngreso.cRazSocUbiGeo " & _
        "INNER JOIN  Constante ON SUBSTRING(Producto.cCtaCod, 9, 1) = Constante.nConsValor " & _
        "INNER JOIN Persona ON ProductoPersona.cPersCod = Persona.cPersCod " & _
        "INNER JOIN Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod " & _
        "INNER JOIN  ColocacCred ON Colocaciones.cCtaCod = ColocacCred.cCtaCod  "
'----------
lsSQL = lsSQL & " WHERE (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
            " AND (ProductoPersona.nPrdPersRelac=" & gColRelPersTitular & ") "
         
    If Len(Trim(pcMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & pcMoneda & ")) "
    End If
    If Len(Trim(pcProductos)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 6, 3) IN (" & pcProductos & ")) "
    End If
    
    lsSQL = lsSQL & " AND (ColocacCred.nColocCondicion IN (" & pcCondicion & ")) " & _
     " AND (Constante.nConsCod=" & gMoneda & ")) P Inner Join " & _
     " ( SELECT  distinct PRP.cCtaCod, RHH.CUSER as cAnalista, PERSO.cPersNombre as cDesAnalista " & _
     " FROM RRHH RHH INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD " & _
     " INNER JOIN PERSONA PERSO ON PRP.CPERSCOD = PERSO.CPERSCOD Where PRP.nPrdPersRelac =" & gColRelPersAnalista & _
     " AND PRP.cPersCod IN(" & pcAnalistas & ") " & _
     " ) as An ON P.cCtaCod=An.cCtaCod "
     
     
    lsSQL = lsSQL & " INNER JOIN (Select cCtaCod, nNroCalen, Max(nCuota) as nCuota From ColocCalendario " & _
       " Where nColocCalendEstado=" & gColocCalendEstadoPendiente & " AND nColocCalendApl=" & gColocCalendAplCuota & " Group By cCtaCod, nNroCalen) " & _
       " as Cal ON Cal.cCtaCod = P.cCtaCod And Cal.nNroCalen = P.nNroCalen "
    
    'Cuotas solo para el reporte paralelo
        lsSQL = lsSQL & " Inner Join ( SELECT dbo.ColocacCred.cCtaCod, ColocacEstado.nCuotas as nCuotasTotal, "
        lsSQL = lsSQL & " (abs(ColocacEstado.nCuotas - ColocacCred.nNroProxCuota) + 1) AS nCuotasPend "
        lsSQL = lsSQL & ",(select count(*)  from coloccalendario where nColocCalendApl = 1 and  cCtaCod=dbo.ColocacCred.cCtaCod and nNroCalen=(select max(nNroCalen) from coloccalendario where cCtaCod=dbo.ColocacCred.cCtaCod)) NroCuotas,"
        lsSQL = lsSQL & "  ColocacCred.nNroProxCuota "
        lsSQL = lsSQL & " FROM dbo.ColocacCred INNER JOIN dbo.ColocacEstado ON dbo.ColocacCred.cCtaCod = "
        lsSQL = lsSQL & " dbo.ColocacEstado.cCtaCod "
        lsSQL = lsSQL & " WHERE (dbo.ColocacEstado.nPrdEstado IN ('" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "')) "

       If pnTipoCuotas = 0 Then
           lsSQL = lsSQL & " AND (ColocacEstado.nCuotas - ColocacCred.nNroProxCuota + 1) <= " & CInt(pnCuotas)
       ElseIf pnTipoCuotas = 1 Then
           lsSQL = lsSQL & " AND round( (  100 * (ColocacEstado.nCuotas - ColocacCred.nNroProxCuota + 1) / ColocacEstado.nCuotas ) , 0) <= " & CInt(pnCuotas)
       End If
               
               
      lsSQL = lsSQL & " ) Cuot ON P.cCtaCod=Cuot.cCtaCod " & _
          " LEFT OUTER JOIN " & _
          " ( "
           
      lsSQL = lsSQL & " SELECT  cCtaCod, MAX(ColocCalificacionAnalista.dColocNotaFecha) AS dFechaMax, " & _
          " ISNULL((Select Max(nColocNota) FROM ColocCalificacionAnalista CCA " & _
          " Where CCA.cCtaCod = ColocCalificacionAnalista.cCtaCod And CCA.dColocNotaFecha = Max(dbo.ColocCalificacionAnalista.dColocNotaFecha) " & _
          " ),0) as nColocNota FROM dbo.ColocCalificacionAnalista Where (nColocNota >=" & pnNota1 & " and nColocNota<=" & pnNota2 & ") " & _
          " GROUP BY cCtaCod ) cNotas ON P.cCtaCod=cNotas.cCtaCod   "
      
    lsSQL = lsSQL & " ORDER BY An.cAnalista, Substring(P.cCtaCod, 9,1), P.cPersNombre, P.cCtaCod "
     
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CLIENTES CON " & pnCuotas & " CUOTAS PENDIENTES PARA TERMINAR", lsCondiciones, lnPage, 161, cSubTit, psCodRepo)

        lnIndice = 7:  lnLineas = 7
        sTempoAnalista = IIf(IsNull(lrDataRep!cAnalista), "", lrDataRep!cAnalista)
        sTempoMoneda = lrDataRep!cMoneda
        
        'Analista
        lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA: " & lrDataRep!cDesAnalista & " - " & lrDataRep!cAnalista & lsNegritaOff & Chr(10) & Chr(10)
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cDesMoneda & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(170, "-") & Chr(10)
        lsCadImp = lsCadImp & " CREDITO                    CLIENTE                             DIR. DEL NEGOCIO                    ZONA               SALDO  CUOT. PEND     DIAS   NOTA   ANA.   %PAG." & Chr(10)
        lsCadImp = lsCadImp & "                                                                                                                      CAPITAL               ATRASO             " & Chr(10)
        lsCadImp = lsCadImp & String(170, "-") & Chr(10)
        lnLineas = lnLineas + 7
        lnIndice = lnIndice + 7
        
        sTempoCred = "xx"
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                nPorPaga2 = 100 - (!nNroProxCuota / !NroCuotas) * 100
                'ALPA 20080620**********************************************************
                If nPorPagado <= nPorPaga2 Or pnTipoCuotas = 0 Then
                '***********************************************************************
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nitem = nitem + 1
                
                If sTempoAnalista <> !cAnalista Then
                
                    lsCadImp = lsCadImp & String(170, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(97) & ImpreFormat(sSCapitalAnalista, 10, 2) & _
                                  lsNegritaOff & Chr(10) & Chr(10)
                                        
 '                   lsCadImp = lsCadImp & String(160, "-") & Chr(10)
 '                   lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ANALISTA" & Space(95) & ImpreFormat(sSCapitalMoneda, 10, 2) & _
 '                               lsNegritaOff & Chr(10) & Chr(10)
                      
                    'Analista
                    lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA: " & lrDataRep!cDesAnalista & " - " & lrDataRep!cAnalista & lsNegritaOff & Chr(10) & Chr(10)
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cDesMoneda & lsNegritaOff & Chr(10)
                    lsCadImp = lsCadImp & String(170, "-") & Chr(10)
                    lsCadImp = lsCadImp & " CREDITO                    CLIENTE                             DIR. DEL NEGOCIO                    ZONA               SALDO  CUOT. PEND     DIAS   NOTA   ANA.   %PAG." & Chr(10)
                    lsCadImp = lsCadImp & "                                                                                                                      CAPITAL               ATRASO             " & Chr(10)
                    lsCadImp = lsCadImp & String(170, "-") & Chr(10)
                      
                    sSCapitalAnalista = !nSaldo
                     
                    sSCapitalMoneda = !nSaldo
                     
                    nitem = 1
                    
                    sTempoAnalista = lrDataRep!cAnalista
                    sTempoMoneda = lrDataRep!cMoneda
                              
                    lnLineas = lnLineas + 13
                    lnIndice = lnIndice + 13
                    
                Else
                    If sTempoCred <> lrDataRep!cCtaCod Then
                        sSCapitalMoneda = sSCapitalMoneda + !nSaldo
                    End If
                    
                    If sTempoMoneda <> !cMoneda Then
                    
                        lsCadImp = lsCadImp & String(170, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(97) & ImpreFormat(sSCapitalAnalista, 10, 2) & _
                                  lsNegritaOff & Chr(10) & Chr(10)
                     
                        'Moneda
                        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cDesMoneda & lsNegritaOff & Chr(10) & Chr(10)
                        lsCadImp = lsCadImp & String(170, "-") & Chr(10)
                         
                        sSCapitalAnalista = !nSaldo
                        nitem = 1
                        
                        sTempoMoneda = lrDataRep!cMoneda
                        
                        lnLineas = lnLineas + 6
                        lnIndice = lnIndice + 6
                        
                    Else
                        If sTempoCred <> lrDataRep!cCtaCod Then
                            sSCapitalAnalista = sSCapitalAnalista + !nSaldo
                        End If
                    End If
                End If
                
                sCliente = PstaNombre(!cPersNombre)
                sDirecNeg = !cRazSocDirecc
                sZonaCli = Trim(!cUbiGeoDescripcion) & " " & Trim(!cRazSocTelef)
                nPorPaga = 100 - (!nNroProxCuota / !NroCuotas) * 100
                If sTempoCred = lrDataRep!cCtaCod Then
                
                   lsCadImp = lsCadImp & Space(127) & _
                            ImpreFormat(!nCuotasPend, 4, 0) & " / " & ImpreFormat(!nCuotasTotal, 4, 0) & Space(1) & ImpreFormat(!nDiasAtraso, 6, 0) & Space(1) & _
                            ImpreFormat(!nColocNota, 6, 0) & Space(3) & !cAnalista & ImpreFormat(nPorPaga, 8, 2) & "%" & Chr(10)
                            
                Else
                    sTempoCred = lrDataRep!cCtaCod
                
                    lsCadImp = lsCadImp & Space(1) & !cCtaCod & Space(1) & sCliente & Space(1) & sDirecNeg & _
                            Space(1) & sZonaCli & Space(1) & ImpreFormat(!nSaldo, 10, 2) & Space(1) & _
                            ImpreFormat(!nCuotasPend, 4, 0) & " / " & ImpreFormat(!nCuotasTotal, 4, 0) & Space(1) & ImpreFormat(!nDiasAtraso, 6, 0) & Space(1) & _
                            ImpreFormat(!nColocNota, 6, 0) & Space(3) & !cAnalista & ImpreFormat(nPorPaga, 8, 2) & "%" & Chr(10)
                End If

                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("CLIENTES CON " & pnCuotas & " CUOTAS PENDIENTES PARA TERMINAR", lsCondiciones, lnPage, 187, cSubTit, psCodRepo)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                End If
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            RaiseEvent CloseProgress
        End With
         
        lsCadImp = lsCadImp & String(170, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(97) & ImpreFormat(sSCapitalAnalista, 10, 2) & _
                      lsNegritaOff & Chr(10) & Chr(10)
                            
'        lsCadImp = lsCadImp & String(160, "-") & Chr(10)
'        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ANALISTA" & Space(95) & ImpreFormat(sSCapitalMoneda, 10, 2) & _
'                    lsNegritaOff & Chr(10) & Chr(10)
        
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
     
    nRepo108303_ClientesCuotasPend = lsCadBuffer
End Function


Public Function nRepo108304_IngresosxGastos(ByVal pdFechaIni As String, ByVal pdFechaFin As String, ByVal cSubTit As String, _
            ByVal pcMoneda As String, ByVal pcProductos As String, ByVal pcCondicion As String, ByVal pcAnalistas As String, _
            ByVal pcAgencias As String, Optional ByVal psCodRepo As String) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nitem As Integer
Dim nTNacionalAprob As Double
Dim nTExtranjeroAprob As Double
Dim nTNacionalSCapital As Double
Dim nTExtranjeroSCapital As Double
Dim loImprimeCab As NColPImpre
        
Dim sCliente As String * 35
Dim sAplicado As String * 10
Dim sConcepto As String * 30

Dim lsCondiciones As String
Dim sCriterio As String
 
Dim sTempoMoneda As String

Dim i As Integer

Dim sSCapitalAnalista As Double
Dim sCuotaAnalista As Double
Dim sMoraAnalista As Double
 
Dim sSCapitalMoneda As Double
 
Dim lsNegritaOn As String
Dim lsNegritaOff As String
   
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)
    
lsCondiciones = " Del " & pdFechaIni & " Al " & pdFechaFin
  
   sCriterio = " WHERE (CONVERT(datetime, SUBSTRING(C2.cMovNro, 1, 8), 103) Between '" & Format(pdFechaIni, "mm/dd/yyyy") & "' AND '" & Format(pdFechaFin, "mm/dd/yyyy") & "') "
 
   lsSQL = "SELECT  C2.cMovNro, C1.cMoneda, C1.cDesMoneda, C2.dFecha, C1.cCtaCod, C1.cPersNombre, C1.cDestino, C2.cDescripcion, ABS(C2.nMonto) as nMonto, AN.cAnalista " & _
            " From ( " & _
            " SELECT dbo.ColocCalendario.cCtaCod, dbo.ColocCalendario.nNroCalen, dbo.ColocCalendario.nCuota, dbo.ColocCalendario.nColocCalendApl, " & _
            " dbo.Persona.cPersNombre, SUBSTRING(dbo.Constante.cConsDescripcion, 1, 1) + ' ' + CONVERT(char(3), dbo.ColocCalendario.nCuota) AS cDestino, " & _
            " SUBSTRING(dbo.ColocCalendario.cCtaCod, 9, 1) AS cMoneda, " & _
            " (SELECT CONS.cConsDescripcion FROM Constante CONS " & _
            " WHERE CONS.nConsCod=" & gMoneda & " AND CONS.nConsValor = SUBSTRING(dbo.ColocCalendario.cCtaCod, 9, 1)) AS cDesMoneda, " & _
            " dbo.ColocacCred.nColocCondicion FROM dbo.ColocCalendario INNER JOIN " & _
            " dbo.Colocaciones ON dbo.ColocCalendario.cCtaCod = dbo.Colocaciones.cCtaCod INNER JOIN " & _
            " dbo.Producto ON dbo.Colocaciones.cCtaCod = dbo.Producto.cCtaCod INNER JOIN " & _
            " dbo.ProductoPersona ON dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod AND " & _
            " dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod INNER JOIN " & _
            " dbo.Persona ON dbo.ProductoPersona.cPersCod = dbo.Persona.cPersCod AND dbo.ProductoPersona.cPersCod = dbo.Persona.cPersCod INNER JOIN " & _
            " dbo.Constante ON dbo.ColocCalendario.nColocCalendApl = dbo.Constante.nConsValor INNER JOIN " & _
            " dbo.ColocacCred ON dbo.Colocaciones.cCtaCod = dbo.ColocacCred.cCtaCod " & _
            " WHERE (dbo.ProductoPersona.nPrdPersRelac=" & gColRelPersTitular & ") " & _
            " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
            " AND (dbo.Constante.nConsCod=" & gColocCalendApl & ") "
            
            If Len(Trim(pcProductos)) > 0 Then
                lsSQL = lsSQL & " AND (SUBSTRING(dbo.ColocCalendario.cCtaCod, 6, 3) IN (" & pcProductos & ")) "
            End If
            
            If Len(Trim(pcCondicion)) > 0 Then
                lsSQL = lsSQL & " AND (dbo.ColocacCred.nColocCondicion IN (" & pcCondicion & ")) "
            End If
            
            If Len(Trim(pcMoneda)) > 0 Then
                lsSQL = lsSQL & " AND (SUBSTRING(dbo.ColocCalendario.cCtaCod, 9, 1) IN (" & pcMoneda & ")) "
            End If
            
            If Len(Trim(pcAgencias)) > 0 Then
                lsSQL = lsSQL & " AND (SUBSTRING(dbo.ColocCalendario.cCtaCod, 4, 2) IN (" & pcAgencias & ")) "
            End If
            
            lsSQL = lsSQL & " ) C1 Inner Join "
            
        lsSQL = lsSQL & " ( " & _
            " SELECT CONVERT(DATETIME, SUBSTRING(dbo.Mov.cMovNro, 1, 8),103) AS dFecha, dbo.Mov.cMovNro, dbo.MovColDet.cCtaCod, dbo.MovColDet.nNroCalen, " & _
            " dbo.MovColDet.nPrdConceptoCod, dbo.MovColDet.nNroCuota, SUM(dbo.MovColDet.nMonto) AS nMonto, dbo.ProductoConcepto.cDescripcion, CASE WHEN dbo.MovColDet.cOpeCod like '1001%' THEN 0 ELSE  1 END as nAplicado " & _
            " FROM dbo.Mov INNER JOIN dbo.MovCol ON dbo.Mov.nMovNro = dbo.MovCol.nMovNro INNER JOIN " & _
            " dbo.MovColDet ON dbo.MovCol.nMovNro = dbo.MovColDet.nMovNro AND dbo.MovCol.cOpeCod = dbo.MovColDet.cOpeCod AND " & _
            " dbo.MovCol.cCtaCod = dbo.MovColDet.cCtaCod AND dbo.MovCol.nNroCalen = dbo.MovColDet.nNroCalen INNER JOIN " & _
            " dbo.ProductoConcepto ON dbo.MovColDet.nPrdConceptoCod = dbo.ProductoConcepto.nPrdConceptoCod WHERE nMovFlag = 0 " & _
            " GROUP BY dbo.Mov.cMovNro, dbo.MovColDet.cCtaCod, dbo.MovColDet.nNroCalen, dbo.MovColDet.nPrdConceptoCod, dbo.MovColDet.nNroCuota, " & _
            " dbo.ProductoConcepto.cDescripcion,dbo.MovColDet.cOpeCod  HAVING (dbo.MovColDet.nPrdConceptoCod LIKE '12%') " & _
            " ) c2 " & _
            " ON C1.cCtaCod = C2.cCtaCod AND C1.nNroCalen=C2.nNroCalen AND C1.nCuota=C2.nNroCuota AND C1.nColocCalendApl = C2.nAplicado "

        lsSQL = lsSQL & " Inner Join " & _
            " ( SELECT  distinct PRP.cCtaCod, RHH.CUSER as cAnalista, PERSO.cPersNombre as cDesAnalista " & _
            " FROM RRHH RHH INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD " & _
            " INNER JOIN PERSONA PERSO ON PRP.CPERSCOD = PERSO.CPERSCOD Where PRP.nPrdPersRelac =" & gColRelPersAnalista & _
            " AND PRP.cPersCod IN(" & pcAnalistas & ") " & _
            " ) as An ON c1.cCtaCod=An.cCtaCod "

        lsSQL = lsSQL & sCriterio
        
        lsSQL = lsSQL & " ORDER BY C1.cMoneda, C2.dFecha, C1.cCtaCod "
     
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
        
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE INGRESOS POR GASTOS", lsCondiciones, lnPage, 127, cSubTit, 108304)

        lnIndice = 7:  lnLineas = 7
        sTempoMoneda = lrDataRep!cMoneda
        
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cDesMoneda & lsNegritaOff & Chr(10)
        lnLineas = lnLineas + 1
        lnIndice = lnIndice + 1
        
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nitem = nitem + 1
                
                If sTempoMoneda <> !cMoneda Then
                
                    lsCadImp = lsCadImp & String(122, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(93) & ImpreFormat(sSCapitalAnalista, 10, 2) & _
                                  lsNegritaOff & Chr(10) & Chr(10)
                        
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cDesMoneda & lsNegritaOff & Chr(10)
                      
                    sSCapitalMoneda = !nMonto
                    sSCapitalAnalista = !nMonto
                    nitem = 1
                    
                    sTempoMoneda = lrDataRep!cMoneda
                              
                    lnLineas = lnLineas + 4
                    lnIndice = lnIndice + 4
                    
                Else
                    sSCapitalAnalista = sSCapitalAnalista + !nMonto
                      
                End If
                
                sCliente = !cPersNombre
                sAplicado = !cDestino
                sConcepto = !CDescripcion
               
                
                lsCadImp = lsCadImp & Space(1) & Format(!dFecha, "dd/MM/YYYY") & Space(1) & !cCtaCod & Space(1) & _
                           sCliente & Space(1) & sAplicado & Space(1) & sConcepto & Space(1) & _
                           ImpreFormat(!nMonto, 10, 2) & Space(1) & !cAnalista & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE INGRESOS POR GASTOS", lsCondiciones, lnPage, 127, cSubTit, 108304)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            RaiseEvent CloseProgress
        End With
        lsCadImp = lsCadImp & String(127, "-") & Chr(10)
                   lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(93) & ImpreFormat(sSCapitalAnalista, 10, 2) & _
                                 lsNegritaOff & Chr(10) & Chr(10)
                                 
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
     
    nRepo108304_IngresosxGastos = lsCadBuffer
End Function
 
 
Public Function nRepo108210_MoraxAnalista(ByVal cSubTit As String, ByVal pdFecFin As String, ByVal pnDia1I As Integer, ByVal pnDia1F As Integer, ByVal pnDia2I As Integer, ByVal pnDia2F As Integer, ByVal pnDia3I As Integer, ByVal pcMoneda As String, ByVal pcProductos As String, ByVal pcCondicion As String, ByVal pcAgencias As String, ByVal pcAnalistas As String) As String

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim xlHojaP As Excel.Worksheet
Dim lnFil As Integer, lnCol As Integer


Dim fs As Scripting.FileSystemObject



Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
xlAplicacion.DisplayAlerts = False


Dim lsNameArchivo As String

lsNameArchivo = "MoraXAnalista" & Format(Now, "yyyymmddhhmmss") & ".xls"
If fs.FileExists(App.path & "\SPOOLER\" & lsNameArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsNameArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If


Dim lsCodigoAgencia As String
Dim lsNombreAgencia As String
Dim lsrt As ADODB.Recordset
Set lsrt = dObtieneAgencias(True, "Desc")

Do While Not lsrt.EOF
    Set xlHoja1 = xlLibro.Worksheets.Add
    lsCodigoAgencia = lsrt("cAgeCod")
    lsNombreAgencia = lsrt("cAgeDescripcion")
    xlHoja1.Name = lsNombreAgencia
    ImprimeCabecera xlHoja1, lsNombreAgencia
    CargaMoraAnalistas xlHoja1, lsCodigoAgencia, 6
    lsrt.MoveNext
Loop

xlHoja1.SaveAs App.path & "\SPOOLER\" & lsNameArchivo

xlAplicacion.DisplayAlerts = True
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
Screen.MousePointer = 0
nRepo108210_MoraxAnalista = "Se ha Generado el Archivo MoraXAnalista.xls Satisfactoriamente"
MsgBox "Se ha Generado el Archivo MoraXAnalista.xls Satisfactoriamente", vbInformation, "Aviso"


Exit Function

ErrorExcel:
    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
    xlLibro.Close
    ' Cierra Microsoft Excel con el método Quit.
    xlAplicacion.Quit
    'Libera los objetos.
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja1 = Nothing

End Function


Public Function nRepo108601_ConsolidadoColocacionesxAgencia(ByVal pnTipoFecha As Integer, ByVal cSubTit As String, ByVal pnTipoCambio As Currency, ByVal pdFecFin As String, ByVal psMoneda As String, ByVal psTipoCredito As String, ByVal psAgencias As String) As String
Dim matAgencia() As String

Dim matVMonN1() As Long
Dim matVMonN2() As Currency
Dim matVMonE1() As Long
Dim matVMonE2() As Currency
 
Dim matVTot1() As Long
Dim matVTot2() As Currency
Dim matVSal1() As Long
Dim matVSal2() As Currency

Dim matMMonN1() As Long
Dim matMMonN2() As Currency
Dim matMMonE1() As Long
Dim matMMonE2() As Currency

Dim matMTot1() As Long
Dim matMTot2() As Currency
Dim matMSal1() As Long
Dim matMSal2() As Currency
Dim matIngRecup()  As Currency

'Los totales de la parte superior
Dim matTot1(11) As Currency
'Los totales de la parte inferior
Dim matTot2(11) As Currency

Dim nContador As Integer
 
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim loImprimeCab As NColPImpre
        
Dim oTipCambio As nTipoCambio
 
'Obtencion de la fecha del mes pasado
Dim dFechaPasada As String
Dim nTipoCambioPasado As Currency

Dim i As Integer
Dim sTempoAgencia As String
Dim cAgencia As String * 15
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsCondiciones As String
Dim lsBaseConsol As String

    
            
    lsSQL = "Select nConsSisValor FROM ConstSistema WHERE nConsSisCod = 43"

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
        lsBaseConsol = lrDataRep!nConsSisValor
    Set loDataRep = Nothing
    Set lrDataRep = Nothing

    lsNegritaOn = Chr$(27) + Chr$(69)
    lsNegritaOff = Chr$(27) + Chr$(70)

    lsCondiciones = "Al " & pdFecFin
    
    dFechaPasada = DateAdd("d", -1, "01" & Mid(pdFecFin, 3, 8))
     
    'La fecha actual es la que pasan como parametro
      
    Set oTipCambio = New nTipoCambio
    'nTipoCambioPasado = Format(oTipCambio.EmiteTipoCambio(dFechaPasada, TCFijoMes), "0.000")
    nTipoCambioPasado = Format(oTipCambio.EmiteTipoCambio(DateAdd("d", 1, dFechaPasada), TCFijoMes), "0.000")
    Set oTipCambio = Nothing
    
    'Obtengo los totales para vigente
    'Saco 10 Nacional y 11 Extranjero el 1 que es Total se sacara multiplicando por el tipo de cambio

If pnTipoFecha = 1 Then 'Fecha Actual Busco en producto
    
    'No se multiplica por el tipo de cambio porque son las cuentas en su propia denominacion
    lsSQL = "SELECT case when substring(Producto.cCtaCod,9,1) = '1' then '10' else '11' end as cLugar, COUNT(Producto.cCtaCod) AS Cantidad, SUM(Producto.nSaldo) AS Total, Colocaciones.cAgeCodAct AS cAgencia, " & _
            "Agencias.cAgeDescripcion FROM Producto INNER JOIN " & _
            "Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod INNER JOIN " & _
            "Agencias ON Colocaciones.cAgeCodAct = Agencias.cAgeCod " & _
            " WHERE  " & _
            " (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
            " AND (Agencias.cAgeCod IN (" & psAgencias & ")) "
            
        If Len(Trim(psMoneda)) > 0 Then
            lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
        End If
        If Len(Trim(psTipoCredito)) > 0 Then
            lsSQL = lsSQL & " AND (Colocaciones.cTpoCredCod IN (" & psTipoCredito & ")) "
        End If
        
        lsSQL = lsSQL & " GROUP BY case when substring(Producto.cCtaCod,9,1) = '1' then '10' else '11' end, Colocaciones.cAgeCodAct, Agencias.cAgeDescripcion "
         
ElseIf pnTipoFecha = 2 Then 'Fecha pasada Busco en Estadisticas
    'No se multiplica por el tipo de cambio porque son las cuentas en su propia denominacion
    lsSQL = "SELECT case when substring(CSC.cCtaCod,9,1) = '1' then '10' else '11' end as cLugar, COUNT(CSC.cCtaCod) AS Cantidad, SUM(CSC.nSaldoCap) AS Total, CCT.cAgeCodAct AS cAgencia, " & _
            " A.cAgeDescripcion FROM " & lsBaseConsol & "CreditoSaldoConsol CSC " & _
            " INNER JOIN" & lsBaseConsol & "CreditoConsolTotal CCT On CCT.cCtaCod = CSC.cCtaCod " & _
            " Agencias A ON CCT.cAgeCodAct = A.cAgeCod " & _
            " WHERE  " & _
            " (CSC.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
            " AND (CCT.cAgeCodAct IN (" & psAgencias & ")) "
            
        If Len(Trim(psMoneda)) > 0 Then
            lsSQL = lsSQL & " AND (SUBSTRING(CSC.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
        End If
        If Len(Trim(psTipoCredito)) > 0 Then
            lsSQL = lsSQL & " AND (CCT.cTpoCredCod IN (" & psTipoCredito & ")) "
        End If
            
        lsSQL = lsSQL & " AND (convert(varchar(8), CSC.dFecha,112)='" & Format(pdFecFin, "yyyyMMdd") & "') "
        lsSQL = lsSQL & " GROUP BY case when substring(CSC.cCtaCod,9,1) = '1' then '10' else '11' end, CCT.cAgeCodAct , A.cAgeDescripcion "
             
End If
  
    'Obtengo los saldos del mes pasado para la cartera vigente

    lsSQL = lsSQL & " UNION ALL "
     
    lsSQL = lsSQL & "SELECT 2 AS Lugar, COUNT(CSC.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaNacional & "' then CSC.nSaldoCap " & _
    " when SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then CSC.nSaldoCap * " & nTipoCambioPasado & "  " & _
    " End) AS Total, CCT.cAgeCodAct AS cAgencia, A.cAgeDescripcion " & _
    " FROM " & lsBaseConsol & "CreditoSaldoConsol CSC " & _
    " INNER JOIN " & lsBaseConsol & "CreditoConsolTotal CCT ON CCT.cCtaCod = CSC.cCtaCod " & _
    " INNER JOIN Agencias A ON CCT.cCtaCod = A.cAgeCod " & _
    " WHERE " & _
    " (CSC.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
    " AND (CCT.cAgeCodAct IN (" & psAgencias & ")) "
    
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CSC.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (CCT.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    
    lsSQL = lsSQL & " AND (CONVERT(VARCHAR(8), CSC.dFecha, 112)='" & Format(dFechaPasada, "YYYYMMdd") & "') "
    lsSQL = lsSQL & " GROUP BY CCT.cAgeCodAct, A.cAgeDescripcion "
  
''''MOROSOS
 
'20) Obtengo vigentes en moneda nacional y extranjera para los morosos
    'saco 20 y 21 juntos . El 6 se obtendra multiplicando
    
    lsSQL = lsSQL & " UNION ALL "
     
If pnTipoFecha = 1 Then 'Fecha Actual
      
    'Para sacar los morosos a la fecha:
        'Para el vigente vencido y Refinanciado Vencido se saca de Producto
        'Para los Vigentes Normales y Morosos y Refinanciados y Morosos
            'Si es producto 1 o 2 se saca de producto
            'Si es producto 3 o 4 se saca de las cuotas pendientes con fecha de vencimiento menor a la del reporte
            'Considerar Calendario Normal y Calendario Paralelo
      
    lsSQL = lsSQL & " select T.cLugar as Lugar, SUM(T.nCantidad) as Cantidad, SUM(T.nMonto) as Total, T.cAgencia, T.cAgeDescripcion "
    lsSQL = lsSQL & " From "
    lsSQL = lsSQL & " ( "
    ' -- Solo los vencidos
    lsSQL = lsSQL & " Select case when substring(P.cCtaCod,9,1) = '1' then '20' else '21' end as cLugar, "
    lsSQL = lsSQL & " C.cAgeCodAct as cAgencia, A.cAgeDescripcion, count(p.cCtaCod) as nCantidad, sum(p.nSaldo) as nMonto "
    lsSQL = lsSQL & " From producto p INNER JOIN Colocaciones C ON C.cCtaCod = P.cCtaCod"
    lsSQL = lsSQL & " INNER JOIN Agencias A ON C.cAgeCodAct = A.cAgeCod"
    lsSQL = lsSQL & " Where p.nPrdEstado  in ('" & gColocEstVigVenc & "', '" & gColocEstRefVenc & "') " '  -- Vencidos
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ") "
    End If
    lsSQL = lsSQL & " And C.cAgeCodAct IN (" & psAgencias & ")  " '-- Agencia
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND C.cTpoCredCod IN (" & psTipoCredito & ") " '--Productos
    End If
    lsSQL = lsSQL & " Group By case when substring(P.cCtaCod,9,1) = '1' then '20' else '21' end, C.cAgeCodAct, A.cAgeDescripcion "
    
    ' ========= MOROSOS CALENDARIO 1
    lsSQL = lsSQL & " Union All "
    lsSQL = lsSQL & " Select case when substring(P.cCtaCod,9,1) = '1' then '20' else '21' end as cLugar, "
    lsSQL = lsSQL & " C.cAgeCodAct as cAgencia, A.cAgeDescripcion, COUNT(DISTINCT P.cCtaCod) as nCantidad, SUM(CCD.nMonto - CCD.nMontoPagado) as nMonto "
    lsSQL = lsSQL & " FROM Producto P "
    lsSQL = lsSQL & " INNER JOIN Colocaciones C ON C.cCtaCod = P.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Agencias A ON c.cAgeCodAct = A.cAgeCod "
    lsSQL = lsSQL & " INNER JOIN ColocacCred CCR ON P.cCtaCod=CCR.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Coloccalendario CC ON CC.cCtaCod=CCR.cCtaCod AND CC.nNroCalen = CCR.nNroCalen "
    lsSQL = lsSQL & " INNER JOIN ColocCalendDet CCD ON CC.cCtaCod=CCD.cCtaCod "
    lsSQL = lsSQL & " AND CC.nNroCalen = CCD.nNroCalen AND CC.nColocCalendApl=CCD.nCOlocCalendApl "
    lsSQL = lsSQL & " AND CC.nCuota = CCD.nCuota "
    lsSQL = lsSQL & " Where CC.ncoloccalendapl =" & gColocCalendAplCuota & " " '- -1 Cuota
    lsSQL = lsSQL & " And datediff(dd, CC.dvenc, '" & Format(pdFecFin, "mm/dd/yyyy") & "') Between 30 AND 91 " '-- Fecha de Vencimiento <='fecha de reporte
    lsSQL = lsSQL & " And CC.nColocCalendEstado=" & gColocCalendEstadoPendiente & " " '-- Vigente 0
    lsSQL = lsSQL & " AND CCD.nPrdConceptoCod=" & gColocConceptoCodCapital & " " '-- 1000 Capital
    lsSQL = lsSQL & " AND P.nPrdEstado in('" & gColocEstVigMor & "', '" & gColocEstRefMor & "') " '-- Normal y Moroso
    lsSQL = lsSQL & " And substring(C.cTpoCredCod,1,1) IN ('7','8') " '-- No vencidos pero Consumo e Hipotecario

    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ") "
    End If
    
    lsSQL = lsSQL & " And C.cAgeCodAct IN (" & psAgencias & ") " ' -- Agencia
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND C.cTpoCredCod IN (" & psTipoCredito & ") " ' --Productos
    End If
    lsSQL = lsSQL & " Group By case when substring(P.cCtaCod,9,1) = '1' then '20' else '21' end, C.cAgeCodAct, A.cAgeDescripcion "
    lsSQL = lsSQL & " Union All "
    
    ' -- Final para morosos calendario 2 paralelo
    lsSQL = lsSQL & " Select case when substring(P.cCtaCod,9,1) = '1' then '20' else '21' end as cLugar, "
    lsSQL = lsSQL & " C.cAgeCodAct as cAgencia, A.cAgeDescripcion, COUNT(DISTINCT P.cCtaCod) as nCantidad, SUM(CCD.nMonto - CCD.nMontoPagado) as nMonto "
    lsSQL = lsSQL & " FROM Producto P "
    lsSQL = lsSQL & " INNER JOIN Colocaciones C ON C.cCtaCod = P.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Agencias A ON C.cAgeCodAct = A.cAgeCod "
    lsSQL = lsSQL & " INNER JOIN ColocacCred CCR ON P.cCtaCod=CCR.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Coloccalendario CC ON CC.cCtaCod=CCR.cCtaCod AND CC.nNroCalen = CCR.nNroCalPar "
    lsSQL = lsSQL & " INNER JOIN ColocCalendDet CCD ON CC.cCtaCod=CCD.cCtaCod "
    lsSQL = lsSQL & " AND CC.nNroCalen = CCD.nNroCalen AND CC.nColocCalendApl=CCD.nCOlocCalendApl "
    lsSQL = lsSQL & " AND CC.nCuota = CCD.nCuota "
    lsSQL = lsSQL & " Where CC.ncoloccalendapl = " & gColocCalendAplCuota & " " '- -Cuota 1
    lsSQL = lsSQL & " And datediff(dd, CC.dvenc, '" & Format(pdFecFin, "mm/dd/yyyy") & "') Between 30 AND 91 " '-- Fecha de Vencimiento <='fecha de reporte
    lsSQL = lsSQL & " And CC.nColocCalendEstado=" & gColocCalendEstadoPendiente & " " '-- Vigente 0
    lsSQL = lsSQL & " AND CCD.nPrdConceptoCod=" & gColocConceptoCodCapital & " " '-- Capital 1000
    lsSQL = lsSQL & " AND P.nPrdEstado in('" & gColocEstVigMor & "', '" & gColocEstRefMor & "') " '-- Normal y Moroso
    lsSQL = lsSQL & " And substring(C.cTpoCredCod,1,1) IN ('6','7','8') " '-- No vencidos pero Consumo e Hipotecario
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ") "
    End If
    lsSQL = lsSQL & " And C.cAgeCodAct IN (" & psAgencias & ") " ' -- Agencia
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND C.cTpoCredCod IN (" & psTipoCredito & ") " ' --Productos
    End If
    lsSQL = lsSQL & " Group By case when substring(P.cCtaCod,9,1) = '1' then '20' else '21' end, C.cAgeCodAct, A.cAgeDescripcion "
    lsSQL = lsSQL & " ) T "
    lsSQL = lsSQL & " Group By T.cLugar, T.cAgencia, T.cAgeDescripcion"
         
ElseIf pnTipoFecha = 2 Then 'Fecha pasada
    'Fecha pasada Busco en Estadisticas
    'No se multiplica por el tipo de cambio porque son las cuentas en su propia denominacion
    'ARCV 25-07-2006    lsSQL = lsSQL & "
    lsSQL = lsSQL & " SELECT case when substring(CSC.cCtaCod,9,1) = '1' then '20' else '21' end as cLugar, COUNT(CSC.cCtaCod) AS Cantidad, SUM(CSC.nSaldoCap) AS Total, CCT.cAgeCodAct AS cAgencia, " & _
            " A.cAgeDescripcion  " & _
            " FROM " & lsBaseConsol & "CreditoSaldoConsol CSC " & _
            " INNER JOIN " & lsBaseConsol & "CreditoConsolTotal CCT ON CCT.cCtaCod = CSC.cCtaCod " & _
            " INNER JOIN Agencias A ON CCT.cAgeCodAct = A.cAgeCod " & _
            " WHERE  " & _
            " (CSC.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
            " AND (CCT.cAgeCodAct IN (" & psAgencias & ")) "
            
        If Len(Trim(psMoneda)) > 0 Then
            lsSQL = lsSQL & " AND (SUBSTRING(CSC.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
        End If
        If Len(Trim(psTipoCredito)) > 0 Then
            lsSQL = lsSQL & " AND (CCT.cTpoCredcod IN (" & psTipoCredito & ")) "
        End If
            
        lsSQL = lsSQL & " AND (convert(varchar(8), CSC.dFecha,112)='" & Format(dFechaPasada, "yyyyMMdd") & "') "
        lsSQL = lsSQL & " GROUP BY case when substring(CSC.cCtaCod,9,1) = '1' then '20' else '21' end, CCT.cAgeCodAct, A.cAgeDescripcion "

    End If
 
'7) Obtengo los saldos del mes pasado (Para todos no inportando el tipo de fecha)
'    para la cartera morosa

    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & "SELECT 7 AS Lugar, COUNT(CSC.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaNacional & "' then CSC.nCapVencido " & _
    " when SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then CSC.nCapVencido * " & nTipoCambioPasado & "  " & _
    " End) AS Total, CCT.cAgeCodAct AS cAgencia, A.cAgeDescripcion " & _
    " FROM " & lsBaseConsol & "CreditoSaldoConsol CSC " & _
    " INNER JOIN " & lsBaseConsol & "CreditoConsolTotal CCT ON CCT.cCtaCod = CSC.cCtaCod " & _
    " INNER JOIN Agencias A ON CCT.cAgeCodAct = A.cAgeCod " & _
    " WHERE " & _
    " (CSC.nPrdEstado IN ('" & gColocEstVigMor & "', '" & gColocEstRefMor & "','" & gColocEstVigVenc & "','" & gColocEstRefVenc & "')) " & _
    " AND (CCT.cAgeCodAct IN (" & psAgencias & ")) "
    
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CSC.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (CCT.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    
    lsSQL = lsSQL & " AND (CONVERT(VARCHAR(8), CSC.dFecha, 112)='" & Format(dFechaPasada, "YYYYMMdd") & "') "
    lsSQL = lsSQL & " GROUP BY CCT.cAgeCodAct, A.cAgeDescripcion "
         
' 6) Obtengo los ingresos a recuperaciones          '--Considero la suma de COlocacEstado.Monto para ColocacEstado.nPrdEstado=2201

     lsSQL = lsSQL & " UNION ALL "
      
    lsSQL = lsSQL & " Select 8 as Lugar, COUNT(T.cCtaCod) as Cantidad, "
    lsSQL = lsSQL & " SUM( "
    lsSQL = lsSQL & " case when SUBSTRING(T.cCtaCod, 9, 1)='1' then T.nCapital "
    lsSQL = lsSQL & " when SUBSTRING(T.cCtaCod, 9, 1)='2' then T.nCapital * " & pnTipoCambio & " end "
    lsSQL = lsSQL & " ) as Total, C.cAgeCodAct as cAgencia , A.cAgeDescripcion  "
    lsSQL = lsSQL & " From ( "
    lsSQL = lsSQL & " Select CR.cCtaCod, "
    lsSQL = lsSQL & " IsNull( "
    lsSQL = lsSQL & " ( SELECT SUM(nmonto)  FROM movcoldet MCD INNER JOIN MOV M ON MCD.nMovNro = M.nMovNro "
    lsSQL = lsSQL & " WHERE M.nMovFlag=0 and MCD.cCtaCod = CR.cCtaCod AND MCD.nPrdConceptoCod like '14%' "
    lsSQL = lsSQL & " And MCD.cOpecod = '130100'), 0)  AS nCapital "
    lsSQL = lsSQL & " FROM ColocRecup CR "
    lsSQL = lsSQL & " where convert(varchar(8),CR.dIngRecup,112)>'" & Format(dFechaPasada, "YYYYMMdd") & "' And convert(varchar(8), CR.dIngRecup,112)<='" & Format(pdFecFin, "YYYYMMdd") & "' "
    lsSQL = lsSQL & " ) T "
    lsSQL = lsSQL & " INNER JOIN Colocaciones C ON T.cCtaCod=C.cCtaCod"
    lsSQL = lsSQL & " INNER JOIN ColocRecup CR ON T.cCtaCod = CR.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Producto PC ON T.cCtaCod = PC.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Agencias A ON C.cAgeCodAct = A.cAgeCod "
    lsSQL = lsSQL & " WHERE (PC.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "', '" & gColocEstRecVigCast & "', '" & gColocEstRecCanCast & "')) "
    lsSQL = lsSQL & " and convert(varchar(8), CR.dIngRecup, 112) >'" & Format(dFechaPasada, "YYYYMMdd") & "' and convert(varchar(8), CR.dIngRecup, 112)<='" & Format(pdFecFin, "YYYYMMdd") & "' "
    lsSQL = lsSQL & " AND (C.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(T.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND C.cTpoCredCod IN (" & psTipoCredito & ") "
    End If
    
    lsSQL = lsSQL & " group by C.cAgeCodAct, A.cAgeDescripcion "
    
    lsSQL = lsSQL & " ORDER BY cAgeDescripcion, cLugar "
 
    'Lugar: 1 si es total, 2 si es saldos , 6 total para la cartera morosa, 7 saldos hasta el mes pasado
    '       para la cartera morosa, 8 ingresos a recuperacioness 10 moneda nacional de los vigentes
    '       11 para la moneda extranjera de los vigentes 20 para la moneda nacional de los morosos
    '       21 Para la moneda extranjera de los morosos
 
    Dim oConn As COMConecta.DCOMConecta
    Dim sqlConn As String
    Dim RRep As ADODB.Recordset
    Set oConn = New COMConecta.DCOMConecta
    oConn.AbreConexion
    
'ARCV 12-06-2007
'    Set RRep = oConn.CargaRecordSet("Select * from constsistema where  nconssiscod = 74")
'    sqlConn = RRep!nConssisValor
'    oConn.CierraConexion
'    Call oConn.AbreConexion(sqlConn)
    
    Set loDataRep = New dColPFunciones
        'Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
        Set lrDataRep = oConn.CargaRecordSet(lsSQL)
    Set loDataRep = Nothing

    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        lrDataRep.Close
        Set lrDataRep = Nothing
        Exit Function
    Else
        nContador = 0
        sTempoAgencia = lrDataRep!cAgencia
        
        ReDim Preserve matAgencia(nContador) As String
        
        ReDim Preserve matVMonN1(nContador) As Long
        ReDim Preserve matVMonN2(nContador) As Currency
        ReDim Preserve matVMonE1(nContador) As Long
        ReDim Preserve matVMonE2(nContador) As Currency
         
        ReDim Preserve matVTot1(nContador) As Long
        ReDim Preserve matVTot2(nContador) As Currency
        ReDim Preserve matVSal1(nContador) As Long
        ReDim Preserve matVSal2(nContador) As Currency
        
        ReDim Preserve matMMonN1(nContador) As Long
        ReDim Preserve matMMonN2(nContador) As Currency
        ReDim Preserve matMMonE1(nContador) As Long
        ReDim Preserve matMMonE2(nContador) As Currency
        
        ReDim Preserve matMTot1(nContador) As Long
        ReDim Preserve matMTot2(nContador) As Currency
        ReDim Preserve matMSal1(nContador) As Long
        ReDim Preserve matMSal2(nContador) As Currency
        ReDim Preserve matIngRecup(nContador) As Currency
 
        matAgencia(0) = lrDataRep!cAgeDescripcion
        RaiseEvent ShowProgress
        Do While Not lrDataRep.EOF
            If sTempoAgencia <> lrDataRep!cAgencia Then
                nContador = nContador + 1
                
                ReDim Preserve matAgencia(nContador) As String
                ReDim Preserve matVMonN1(nContador) As Long
                ReDim Preserve matVMonN2(nContador) As Currency
                ReDim Preserve matVMonE1(nContador) As Long
                ReDim Preserve matVMonE2(nContador) As Currency
                 
                ReDim Preserve matVTot1(nContador) As Long
                ReDim Preserve matVTot2(nContador) As Currency
                ReDim Preserve matVSal1(nContador) As Long
                ReDim Preserve matVSal2(nContador) As Currency
                
                ReDim Preserve matMMonN1(nContador) As Long
                ReDim Preserve matMMonN2(nContador) As Currency
                ReDim Preserve matMMonE1(nContador) As Long
                ReDim Preserve matMMonE2(nContador) As Currency
                
                ReDim Preserve matMTot1(nContador) As Long
                ReDim Preserve matMTot2(nContador) As Currency
                ReDim Preserve matMSal1(nContador) As Long
                ReDim Preserve matMSal2(nContador) As Currency
                ReDim Preserve matIngRecup(nContador) As Currency

                matAgencia(nContador) = lrDataRep!cAgeDescripcion
                sTempoAgencia = lrDataRep!cAgencia
            End If
            
            'pnTipoCambio
            
            With lrDataRep
                Select Case !cLugar
'                Case 1
'                    matVTot1(nContador) = !cantidad
'                    matVTot2(nContador) = !Total
'                Case 6
'                    matMTot1(nContador) = !cantidad
'                    matMTot2(nContador) = !Total
                  
                'Vigente
                Case 10 'Moneda Nacional Vigentes
                    matVMonN1(nContador) = !cantidad
                    matVMonN2(nContador) = Format(!Total, "#0.00")
                    
                    'Sumo en total de vigente
                    matVTot1(nContador) = matVTot1(nContador) + !cantidad
                    matVTot2(nContador) = matVTot2(nContador) + Format(!Total, "#0.00")
                    
                Case 11 'Moneda Extranjera Vigentes
                    matVMonE1(nContador) = !cantidad
                    matVMonE2(nContador) = Format(!Total, "#0.00")
                
                    'Sumo en total de vigente
                    matVTot1(nContador) = matVTot1(nContador) + !cantidad
                    matVTot2(nContador) = matVTot2(nContador) + Format((!Total * pnTipoCambio), "#0.00")
                    
                Case 2 'Saldo Vigente mes pasado
                    matVSal1(nContador) = !cantidad
                    matVSal2(nContador) = Format(!Total, "#0.00")
                
                'Moroso
                Case 20 'Moneda Nacional Morosa
                    matMMonN1(nContador) = !cantidad
                    matMMonN2(nContador) = Format(!Total, "#0.00")
                    
                    'Sumo en total de vigente
                    matMTot1(nContador) = matMTot1(nContador) + !cantidad
                    matMTot2(nContador) = matMTot2(nContador) + Format(!Total, "#0.00")
                    
                Case 21 'Moneda Extranjera Morosa
                    matMMonE1(nContador) = !cantidad
                    matMMonE2(nContador) = Format(!Total, "#0.00")
                
                    'Sumo en total de vigente
                    matMTot1(nContador) = matMTot1(nContador) + !cantidad
                    matMTot2(nContador) = matMTot2(nContador) + Format((!Total * pnTipoCambio), "#0.00")
                
                Case 7 'Saldo Moroso mes pasado
                    matMSal1(nContador) = !cantidad
                    matMSal2(nContador) = Format(!Total, "#0.00")
                
                'Ingreso a Judicial
                Case 8
                    matIngRecup(nContador) = Format(!Total, "#0.00")
                
                End Select
            End With
            
            RaiseEvent Progress(lrDataRep.Bookmark, lrDataRep.RecordCount)
            lrDataRep.MoveNext
        Loop
        
        lrDataRep.Close
        Set lrDataRep = Nothing
        
        RaiseEvent CloseProgress
        
        'Calculo los totales
        For i = 0 To nContador
           
            'Vigentes
            '========
            'Total de Numero para Moneda Nacional de Vigentes
            matTot1(0) = matTot1(0) + matVMonN1(i)
            'Total de Total para Moneda Nacional de Vigentes
            matTot1(1) = matTot1(1) + matVMonN2(i)
            'Total de Numero para Moneda Nacional de Vigentes
            matTot1(2) = matTot1(2) + matVMonE1(i)
            'Total de Total para Moneda Nacional de Vigentes
            matTot1(3) = matTot1(3) + matVMonE2(i)
            
            'Total de Numero para Totales
            matTot1(4) = matTot1(4) + matVTot1(i)
            'Total de Totales para Totales
            matTot1(5) = matTot1(5) + matVTot2(i)
            'Total de Numero para Saldos
            matTot1(6) = matTot1(6) + matVSal1(i)
            'Total de Totales para Saldos
            matTot1(7) = matTot1(7) + matVSal2(i)
            
            'Variacion a la fecha para vigente de numero
            matTot1(8) = matTot1(8) + matVTot1(i) - matVSal1(i)
            'Variacion a la fecha para vigente de monto
            matTot1(9) = matTot1(9) + matVTot2(i) - matVSal2(i)
            
            'Total de Ingresos a Judiciales
            matTot1(10) = matTot1(10) + matIngRecup(i)
            
            'Variacion Real para Vigentes
            matTot1(11) = matTot1(11) + matVTot2(i) - matVSal2(i) + matIngRecup(i)

            'Morosos
            '=======
            'Total de Numero para Moneda Nacional de Morosos
            matTot2(0) = matTot2(0) + matMMonN1(i)
            'Total de Total para Moneda Nacional de Morosos
            matTot2(1) = matTot2(1) + matMMonN2(i)
            'Total de Numero para Moneda Nacional de Morosos
            matTot2(2) = matTot2(2) + matMMonE1(i)
            'Total de Total para Moneda Nacional de Morosos
            matTot2(3) = matTot2(3) + matMMonE2(i)
            
            'Total de Numero para Totales
            matTot2(4) = matTot2(4) + matMTot1(i)
            'Total de Totales para Totales
            matTot2(5) = matTot2(5) + matMTot2(i)
            'Total de Numero para Saldos
            matTot2(6) = matTot2(6) + matMSal1(i)
            'Total de Totales para Saldos
            matTot2(7) = matTot2(7) + matMSal2(i)
            
            'Variacion a la fecha para vigente de numero
            matTot2(8) = matTot2(8) + matMTot1(i) - matMSal1(i)
            'Variacion a la fecha para vigente de monto
            matTot2(9) = matTot2(9) + matMTot2(i) - matMSal2(i)
             
            'SUma de Ingresos a Recuperaciones
            matTot2(10) = matTot2(10) + matIngRecup(i)
            'Variacion Real
            matTot2(11) = matTot2(11) + matMTot2(i) - matVSal2(i) + matIngRecup(i)
             
        Next
                
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CUADRO CONSOLIDADO DE COLOCACIONES X AGENCIA", lsCondiciones, lnPage, 165, cSubTit & Chr(10) & "TCF. = " & Format(pnTipoCambio, "0.000") & " // TCF. (" & dFechaPasada & ") = " & Format(nTipoCambioPasado, "0.000"), 108601)
        lsCadImp = lsCadImp & Chr(10)
        
        lsCadImp = lsCadImp & lsNegritaOn & "C A R T E R A  V I G E N T E" & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(165, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA            |   MONEDA  NACIONAL   |  MONEDA  EXTRANJERA  | TOTAL (Nuevos Soles) | SALDO AL " & dFechaPasada & "  |  V A R I A C I O N   |    INGRESO    |   VARIACION   " & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "                   | NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |   M O N T O   |   COB.  JUD.  |  R  E  A  L   " & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(165, "-") & Chr(10)
        
        lnIndice = 13:  lnLineas = 13
             
        For i = 0 To nContador
            lnIndice = lnIndice + 1
            lnLineas = lnLineas + 1
 
            cAgencia = matAgencia(i)
            lsCadImp = lsCadImp & Space(1) & cAgencia & Space(2) & _
                    " | " & ImpreFormat(matVMonN1(i), 4, 0) & " | " & ImpreFormat(matVMonN2(i), 10, 2) & _
                    " | " & ImpreFormat(matVMonE1(i), 4, 0) & " | " & ImpreFormat(matVMonE2(i), 10, 2) & _
                    " | " & ImpreFormat(matVTot1(i), 4, 0) & " | " & ImpreFormat(matVTot2(i), 10, 2) & _
                    " | " & ImpreFormat(matVSal1(i), 4, 0) & " | " & ImpreFormat(matVSal2(i), 10, 2) & _
                    " | " & ImpreFormat(matVTot1(i) - matVSal1(i), 4, 0) & " | " & ImpreFormat(matVTot2(i) - matVSal2(i), 10, 2) & _
                    " | " & ImpreFormat(matIngRecup(i), 10, 2) & _
                    " | " & ImpreFormat(matVTot2(i) - matVSal2(i) + matIngRecup(i), 10, 2) & _
                    Chr(10)
          
            If lnIndice Mod 300 = 0 Then
                lsCadBuffer = lsCadBuffer & lsCadImp
                lsCadImp = ""
            End If
        
            If lnLineas >= 55 Then
                lnPage = lnPage + 1
                lsCadImp = lsCadImp & Chr(12)
                lsCadImp = lsCadImp & nRepoCabecera("CUADRO CONSOLIDADO DE COLOCACIONES X AGENCIA", lsCondiciones, lnPage, 165, cSubTit & Chr(10) & "TCF. = " & Format(pnTipoCambio, "0.000") & " // TCF. (" & dFechaPasada & ") = " & Format(nTipoCambioPasado, "0.000"), 108601)
                lnLineas = 7
                lnIndice = lnIndice + 7
            End If
        Next
        
        'Imprimo los totales de vigentes
        lsCadImp = lsCadImp & String(165, "=") & Chr(10)
        
        lsCadImp = lsCadImp & lsNegritaOn & Space(18) & " | " & ImpreFormat(matTot1(0), 4, 0) & " | " & ImpreFormat(matTot1(1), 10, 2) & _
                " | " & ImpreFormat(matTot1(2), 4, 0) & " | " & ImpreFormat(matTot1(3), 10, 2) & _
                " | " & ImpreFormat(matTot1(4), 4, 0) & " | " & ImpreFormat(matTot1(5), 10, 2) & _
                " | " & ImpreFormat(matTot1(6), 4, 0) & " | " & ImpreFormat(matTot1(7), 10, 2) & _
                " | " & ImpreFormat(matTot1(8), 4, 0) & " | " & ImpreFormat(matTot1(9), 10, 2) & _
                " | " & ImpreFormat(matTot1(10), 10, 2) & " | " & ImpreFormat(matTot1(11), 10, 2) & _
                 lsNegritaOff & Chr(10)
        
        lsCadImp = lsCadImp & String(165, "_") & Chr(10) & Chr(10) & Chr(10)
        
        lsCadImp = lsCadImp & lsNegritaOn & "C A R T E R A  M O R O S A" & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(165, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA            |   MONEDA  NACIONAL   |  MONEDA  EXTRANJERA  | TOTAL (Nuevos Soles) | SALDO AL " & dFechaPasada & "  |  V A R I A C I O N   |    INGRESO    |   VARIACION   " & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "                   | NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |   M O N T O   |   COB.  JUD.  |  R  E  A  L   " & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(165, "-") & Chr(10)
        
        lnIndice = lnIndice + 9:  lnLineas = lnLineas + 9
             
        For i = 0 To nContador
            lnIndice = lnIndice + 1
            lnLineas = lnLineas + 1
            cAgencia = matAgencia(i)
             
            lsCadImp = lsCadImp & Space(1) & cAgencia & Space(2) & _
                    " | " & ImpreFormat(matMMonN1(i), 4, 0) & " | " & ImpreFormat(matMMonN2(i), 10, 2) & _
                    " | " & ImpreFormat(matMMonE1(i), 4, 0) & " | " & ImpreFormat(matMMonE2(i), 10, 2) & _
                    " | " & ImpreFormat(matMTot1(i), 4, 0) & " | " & ImpreFormat(matMTot2(i), 10, 2) & _
                    " | " & ImpreFormat(matMSal1(i), 4, 0) & " | " & ImpreFormat(matMSal2(i), 10, 2) & _
                    " | " & ImpreFormat(matMTot1(i) - matMSal1(i), 4, 0) & " | " & ImpreFormat(matMTot2(i) - matMSal2(i), 10, 2) & _
                    " | " & ImpreFormat(matIngRecup(i), 10, 2) & _
                    " | " & ImpreFormat(matMTot2(i) - matMSal2(i) + matIngRecup(i), 10, 2) & _
                    Chr(10)
          
            If lnIndice Mod 300 = 0 Then
                lsCadBuffer = lsCadBuffer & lsCadImp
                lsCadImp = ""
            End If
        
            If lnLineas >= 55 Then
                lnPage = lnPage + 1
                lsCadImp = lsCadImp & Chr(12)
                lsCadImp = lsCadImp & nRepoCabecera("CUADRO CONSOLIDADO DE COLOCACIONES X AGENCIA", lsCondiciones, lnPage, 165, "TCF. = " & Format(pnTipoCambio, "0.000") & " // TCF. (" & dFechaPasada & ") = " & Format(nTipoCambioPasado, "0.000"), 108601)
                lnLineas = 7
                lnIndice = lnIndice + 7
            End If
        Next
        
        'Imprimo los totales de morosos
        lsCadImp = lsCadImp & String(165, "=") & Chr(10)
        
        lsCadImp = lsCadImp & lsNegritaOn & Space(18) & " | " & ImpreFormat(matTot2(0), 4, 0) & " | " & ImpreFormat(matTot2(1), 10, 2) & _
                " | " & ImpreFormat(matTot2(2), 4, 0) & " | " & ImpreFormat(matTot2(3), 10, 2) & _
                " | " & ImpreFormat(matTot2(4), 4, 0) & " | " & ImpreFormat(matTot2(5), 10, 2) & _
                " | " & ImpreFormat(matTot2(6), 4, 0) & " | " & ImpreFormat(matTot2(7), 10, 2) & _
                " | " & ImpreFormat(matTot2(8), 4, 0) & " | " & ImpreFormat(matTot2(9), 10, 2) & _
                " | " & ImpreFormat(matTot2(10), 10, 0) & " | " & ImpreFormat(matTot2(5) - matTot2(7) + matTot2(10), 10, 2) & _
                 lsNegritaOff & Chr(10)
        
        lsCadImp = lsCadImp & String(165, "_") & Chr(10) & Chr(10)

        lnLineas = lnLineas + 3
        lnIndice = lnIndice + 3
        
        If lnLineas >= 52 Then
            lnPage = lnPage + 1
            lsCadImp = lsCadImp & Chr(12)
            lsCadImp = lsCadImp & nRepoCabecera("CUADRO CONSOLIDADO DE COLOCACIONES X AGENCIA", lsCondiciones, lnPage, 165, cSubTit & Chr(10) & "TCF. = " & Format(pnTipoCambio, "0.000") & " // TCF. (" & dFechaPasada & ") = " & Format(nTipoCambioPasado, "0.000"), 108601)
            lnLineas = 7
            lnIndice = lnIndice + 7
        End If
        '*********************************************
        'Se agrego la cartera judicial actual
        '*********************************************
        '***** BRGO - BASILEA II
        lsSQL = "Select SUM(case when substring(P.cctacod,9,1)='1' then P.nsaldo  else P.nsaldo*" & pnTipoCambio & "  end ) as nSaldoJud"
        lsSQL = lsSQL & " from producto P"
        lsSQL = lsSQL & " inner join Colocaciones C on C.cCtaCod = P.cCtaCod "
        lsSQL = lsSQL & " Where C.cAgeCodAct IN (" & psAgencias & ")"
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psTipoCredito & ")) "
        lsSQL = lsSQL & " and P.nprdestado in (2201) "
        '************
        Set loDataRep = New dColPFunciones
        Set lrDataRep = oConn.CargaRecordSet(lsSQL)
        Set loDataRep = Nothing

        Dim nSaldoJud As Double
        nSaldoJud = 0
        If lrDataRep.RecordCount > 0 Then
            nSaldoJud = Format(lrDataRep!nSaldoJud, "#0.00")
        End If
        
        lsCadImp = lsCadImp & Space(10) & "--- *******     RESUMEN    ******* ---" & Chr(10)
        lsCadImp = lsCadImp & Space(10) & "CREDITOS VIG. NORMALES : " & ImpreFormat(matTot1(5), 10, 2) & Space(3) & "(A)" & Chr(10)
        lsCadImp = lsCadImp & Space(10) & "CREDITOS VIG. MOROSOS  : " & ImpreFormat(matTot2(5), 10, 2) & Space(3) & "(B)" & Chr(10)
        'ARCV 05-07-2007
        'lsCadImp = lsCadImp & Space(10) & "CARTERA EN JUDICIAL    : " & ImpreFormat(nSaldoJud, 10, 2) & Space(3) & "(E)" & Chr(10)
        '----
        lsCadImp = lsCadImp & Space(10) & "CREDITOS JUD INGRESOS. : " & ImpreFormat(matTot1(10), 10, 2) & Space(3) & "(C)" & Chr(10)
        lsCadImp = lsCadImp & Space(10) & "======================================" & Chr(10)
        If matTot1(5) + matTot2(5) + matTot1(10) = 0 Then
            'ARCV 05-07-2007
            lsCadImp = lsCadImp & Space(10) & "TOTAL CARTERA CRED.    : " & ImpreFormat(matTot1(5) + matTot2(5) + matTot1(10), 10, 2) & Space(3) & "(D)" & Space(20) & "MORA (B) + (C) / (D) : " & ImpreFormat(0, 10, 2) & " % " & Chr(10)
            'lsCadImp = lsCadImp & Space(10) & "TOTAL CARTERA CRED.    : " & ImpreFormat(matTot1(5) + nSaldoJud, 10, 2) & Space(3) & "(D)" & Space(20) & "MORA (B) + (C) / (D) : " & ImpreFormat(0, 10, 2) & " % " & Chr(10)
            '----
        Else
            'ARCV 05-07-2007
            lsCadImp = lsCadImp & Space(10) & "TOTAL CARTERA CRED.    : " & ImpreFormat(matTot1(5) + matTot2(5) + matTot1(10), 10, 2) & Space(3) & "(D)" & Space(20) & "MORA MES (B) + (C) / (D)       : " & ImpreFormat((matTot2(5) + matTot1(10)) / (matTot1(5) + matTot2(5) + matTot1(10)), 10, 2) & " % " & Chr(10)
            'lsCadImp = lsCadImp & Space(10) & "TOTAL CARTERA CRED.    : " & ImpreFormat(matTot1(5) + nSaldoJud, 10, 2) & Space(3) & "(D)" & Space(20) & "MORA MES (B) + (C) / (D)       : " & ImpreFormat(Format((matTot2(5) + matTot1(10)) * 100 / (matTot1(5) + nSaldoJud), "#0.00"), 10, 2) & " % " & Chr(10)
            '---
        End If
        'ARCV 05-07-2007
        'lsCadImp = lsCadImp & Space(74) & "MORA TOTAL (B) + (E) / (A + E) : " & ImpreFormat(Format((matTot2(5) + nSaldoJud) * 100 / (matTot1(5) + nSaldoJud), "#0.00"), 10, 2) & " % " & Chr(10)
        lsCadImp = lsCadImp & Space(74) & "MORA TOTAL (B) / (A) : " & ImpreFormat(Format((matTot2(5)) * 100 / (matTot1(5)), "#0.00"), 10, 2) & " % " & Chr(10)
        '---------
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
 
    End If
    
    oConn.CierraConexion
    Set oConn = Nothing

      nRepo108601_ConsolidadoColocacionesxAgencia = lsCadBuffer
    
End Function

Public Function nRepo108602_ConsolidadoColocacionesxAnalista(ByVal pnTipoFecha As Integer, ByVal cSubTit As String, ByVal pnTipoCambio As Currency, ByVal pdFecFin As String, ByVal psMoneda As String, ByVal psTipoCredito As String, ByVal psAgencias As String, ByVal psAnalistas) As String
Dim matAnalista() As String
Dim matVTot1() As Long
Dim matVTot2() As Currency
Dim matVSal1() As Long
Dim matVSal2() As Currency
Dim matVDesemNue1() As Long
Dim matVDesemNue2() As Currency
Dim matVDesemRep1() As Long
Dim matVDesemRep2() As Currency
Dim matVDesemRef1() As Long
Dim matVDesemRef2() As Currency
Dim matMTot1() As Long
Dim matMTot2() As Currency
Dim matMSal1() As Long
Dim matMSal2() As Currency
Dim matMIngRecup()  As Currency

'ARCV 23-06-2007
Dim matVDesemRec1() As Long
Dim matVDesemRec2() As Currency
Dim matVDesemAut1() As Long
Dim matVDesemAut2() As Currency
Dim matVDesemAmp1() As Long
Dim matVDesemAmp2() As Currency
'-------
'WIOR 20130627 **************************
Dim matVDesemAdi1() As Long
Dim matVDesemAdi2() As Currency
'WIOR FIN **************************

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim loImprimeCab As NColPImpre
        
Dim lsCondiciones As String

Dim oTipCambio As nTipoCambio
 
'Obtencion de la fecha del mes pasado
 
Dim dFechaPasada As String
Dim nTipoCambioPasado As Currency

Dim i As Integer
Dim sTempoAnalista As String
Dim cUser As String * 4
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim ntempo1 As Currency

'Los totales de la parte superior

'Dim matTot1(13) As Currency
Dim matTot1(21) As Currency 'ARCV 23-06-2007'WIOR 20130627 SE CAMBIO DE 19 A 21

'Los totales de la parte inferior
Dim matTot2(8) As Currency
Dim nContador As Integer
Dim lsBaseConsol As String

    lsSQL = "Select nConsSisValor FROM ConstSistema WHERE nConsSisCod = 43"

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
        lsBaseConsol = lrDataRep!nConsSisValor
    Set loDataRep = Nothing
    Set lrDataRep = Nothing
    
    lsNegritaOn = Chr$(27) + Chr$(69)
    lsNegritaOff = Chr$(27) + Chr$(70)

    lsCondiciones = "Al " & pdFecFin

    dFechaPasada = DateAdd("d", -1, "01" & Mid(pdFecFin, 3, 8))
     
    'La fecha actual es la que pasan como parametro
      
    Set oTipCambio = New nTipoCambio
    'nTipoCambioPasado = Format(oTipCambio.EmiteTipoCambio(dFechaPasada, TCFijoMes), "0.000")
    nTipoCambioPasado = Format(oTipCambio.EmiteTipoCambio(DateAdd("d", 1, dFechaPasada), TCFijoMes), "0.000")
    Set oTipCambio = Nothing
    
    'Obtengo los totales para vigente
    'Sacar 1
    
    'BRGO 07/06/2010 BASILEA II Adecuación de los nuevos tipos de créditos y agencias
If pnTipoFecha = 1 Then 'Fecha Actual
    lsSQL = " SELECT 1 AS Lugar, COUNT(Producto.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Producto.nSaldo " & _
    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Producto.nSaldo * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN Colocaciones ON Colocaciones.cCtaCod = ProductoPersona.cCtaCod " & _
    " INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
    " AND (Colocaciones.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (Colocaciones.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
    
ElseIf pnTipoFecha = 2 Then 'Fecha pasada Busco en Estadisticas
    lsSQL = "SELECT 1 AS Lugar, COUNT(CSC.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaNacional & "' then CSC.nSaldoCap " & _
    " when SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then CSC.nSaldoCap * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN  " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod " & _
    " INNER JOIN " & lsBaseConsol & "CreditoSaldoConsol CSC ON ProductoPersona.cCtaCod = CSC.cCtaCod " & _
    " INNER JOIN " & lsBaseConsol & "CreditoConsolTotal CCT ON CCT.cCtaCod = CSC.cCtaCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (convert(varchar(8), CSC.dFecha,112)='" & Format(pdFecFin, "yyyyMMdd") & "') " & _
    " AND (CSC.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
    " AND (CCT.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CSC.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (CCT.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
End If

    'Obtengo los saldos del mes pasado para la cartera vigente

    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 2 AS Lugar, COUNT(CSC.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaNacional & "' then CSC.nSaldoCap " & _
    " when SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then CSC.nSaldoCap * " & nTipoCambioPasado & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod " & _
    " INNER JOIN " & lsBaseConsol & "CreditoSaldoConsol CSC ON ProductoPersona.cCtaCod = CSC.cCtaCod " & _
    " INNER JOIN " & lsBaseConsol & "CreditoConsolTotal CCT ON CCT.cCtaCod = CSC.cCtaCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (CONVERT(VARCHAR(8), CSC.dFecha, 112)='" & Format(dFechaPasada, "YYYYMMdd") & "') " & _
    " AND (CSC.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
    " AND (CCT.cAgeCodAct IN (" & psAgencias & "))  "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CSC.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (CCT.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    
    lsSQL = lsSQL & " GROUP BY RRHH.cUser"

'3) Obtengo desembolsos nuevos      La fecha es el rango entre el 31 del mes pasado a la fecha consultada y la fecha consultada
                                   ' Aca busco los coloccondicion=1
' para la cartera vigente

    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 3 AS Lugar, COUNT(Colocaciones.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Colocaciones.nMontoCol " & _
    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Colocaciones.nMontoCol * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod INNER JOIN " & _
    " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod INNER JOIN ColocacCred ON Colocaciones.cCtaCod = ColocacCred.cCtaCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (Colocaciones.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (Colocaciones.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    
    lsSQL = lsSQL & " AND (ColocacCred.nColocCondicion=" & gColocCredCondNormal & ") "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)>'" & Format(dFechaPasada, "YYYYMMdd") & "') "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)<= '" & Format(pdFecFin, "YYYYmmdd") & "') "
    lsSQL = lsSQL & " AND nPrdEstado Not In (" & gColocEstSug & "," & gColocEstAprob & ","
    lsSQL = lsSQL & gColocEstRech & "," & gColocEstRetirado & ") GROUP BY RRHH.cUser "

'ARCV 23-06-2007
  
'CREDITOS RECURRENTES
    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 4 AS Lugar, COUNT(Colocaciones.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Colocaciones.nMontoCol " & _
    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Colocaciones.nMontoCol * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod INNER JOIN " & _
    " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod INNER JOIN ColocacCred ON Colocaciones.cCtaCod = ColocacCred.cCtaCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (Colocaciones.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (Colocaciones.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    
    lsSQL = lsSQL & " AND (ColocacCred.nColocCondicion=" & gColocCredCondRecurrente & ") "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)>'" & Format(dFechaPasada, "YYYYMMdd") & "') "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)<= '" & Format(pdFecFin, "YYYYmmdd") & "') "
    lsSQL = lsSQL & " AND nPrdEstado Not In (" & gColocEstSug & "," & gColocEstAprob & ","
    lsSQL = lsSQL & gColocEstRech & "," & gColocEstRetirado & ") GROUP BY RRHH.cUser "
  
'CREDITOS PARALELOS
    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 5 AS Lugar, COUNT(Colocaciones.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Colocaciones.nMontoCol " & _
    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Colocaciones.nMontoCol * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod INNER JOIN " & _
    " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod INNER JOIN ColocacCred ON Colocaciones.cCtaCod = ColocacCred.cCtaCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (Colocaciones.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (Colocaciones.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    
    lsSQL = lsSQL & " AND (ColocacCred.nColocCondicion=" & gColocCredCondParalelo & ") "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)>'" & Format(dFechaPasada, "YYYYMMdd") & "') "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)<= '" & Format(pdFecFin, "YYYYmmdd") & "') "
    lsSQL = lsSQL & " AND nPrdEstado Not In (" & gColocEstSug & "," & gColocEstAprob & ","
    lsSQL = lsSQL & gColocEstRech & "," & gColocEstRetirado & ") GROUP BY RRHH.cUser "
  
'CREDITOS REFINANCIADOS
    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 6 AS Lugar, COUNT(Colocaciones.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Colocaciones.nMontoCol " & _
    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Colocaciones.nMontoCol * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod INNER JOIN " & _
    " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod INNER JOIN ColocacCred ON Colocaciones.cCtaCod = ColocacCred.cCtaCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (Colocaciones.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (Colocaciones.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    
    lsSQL = lsSQL & " AND (ColocacCred.nColocCondicion=" & gColocCredCondRefinan & ") "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)>'" & Format(dFechaPasada, "YYYYMMdd") & "') "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)<= '" & Format(pdFecFin, "YYYYmmdd") & "') "
    lsSQL = lsSQL & " AND nPrdEstado Not In (" & gColocEstSug & "," & gColocEstAprob & ","
    lsSQL = lsSQL & gColocEstRech & "," & gColocEstRetirado & ") GROUP BY RRHH.cUser "
  
'CREDITOS AMPLIADOS
    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 7 AS Lugar, COUNT(Colocaciones.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Colocaciones.nMontoCol " & _
    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Colocaciones.nMontoCol * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod INNER JOIN " & _
    " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod INNER JOIN ColocacCred ON Colocaciones.cCtaCod = ColocacCred.cCtaCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (Colocaciones.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (Colocaciones.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    
    lsSQL = lsSQL & " AND (ColocacCred.nColocCondicion=" & 5 & ") "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)>'" & Format(dFechaPasada, "YYYYMMdd") & "') "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)<= '" & Format(pdFecFin, "YYYYmmdd") & "') "
    lsSQL = lsSQL & " AND nPrdEstado Not In (" & gColocEstSug & "," & gColocEstAprob & ","
    lsSQL = lsSQL & gColocEstRech & "," & gColocEstRetirado & ") GROUP BY RRHH.cUser "
  
'CREDITOS AUTOMATICOS
    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 8 AS Lugar, COUNT(Colocaciones.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Colocaciones.nMontoCol " & _
    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Colocaciones.nMontoCol * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod INNER JOIN " & _
    " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod INNER JOIN ColocacCred ON Colocaciones.cCtaCod = ColocacCred.cCtaCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (Colocaciones.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (Colocaciones.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    
    lsSQL = lsSQL & " AND (ColocacCred.nColocCondicion=" & 6 & ") "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)>'" & Format(dFechaPasada, "YYYYMMdd") & "') "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)<= '" & Format(pdFecFin, "YYYYmmdd") & "') "
    lsSQL = lsSQL & " AND nPrdEstado Not In (" & gColocEstSug & "," & gColocEstAprob & ","
    lsSQL = lsSQL & gColocEstRech & "," & gColocEstRetirado & ") GROUP BY RRHH.cUser "
    
    'WIOR 20130627 **************************************************************************
    'CREDITOS ADICIONALES
    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 9 AS Lugar, COUNT(Colocaciones.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Colocaciones.nMontoCol " & _
    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Colocaciones.nMontoCol * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod INNER JOIN " & _
    " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod INNER JOIN ColocacCred ON Colocaciones.cCtaCod = ColocacCred.cCtaCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (Colocaciones.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (Colocaciones.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    
    lsSQL = lsSQL & " AND (ColocacCred.nColocCondicion=" & 7 & ") "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)>'" & Format(dFechaPasada, "YYYYMMdd") & "') "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)<= '" & Format(pdFecFin, "YYYYmmdd") & "') "
    lsSQL = lsSQL & " AND nPrdEstado Not In (" & gColocEstSug & "," & gColocEstAprob & ","
    lsSQL = lsSQL & gColocEstRech & "," & gColocEstRetirado & ") GROUP BY RRHH.cUser "
    'WIOR FIN *******************************************************************************
  
' 4) Obtengo desembolosos represtamos                La fecha es el rango entre el 31 del mes pasado a la fecha consultada y la fecha consultada
'                                                    'Aca busco los coloccondicion<>1
' para la cartera vigente

'    lsSQL = lsSQL & " UNION ALL "
'
'    lsSQL = lsSQL & " SELECT 4 AS Lugar, COUNT(Colocaciones.cCtaCod) AS Cantidad, " & _
'    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Colocaciones.nMontoCol " & _
'    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Colocaciones.nMontoCol * " & pnTipoCambio & " " & _
'    " End) AS Total, RRHH.cUser AS Analista " & _
'    " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN " & _
'    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod INNER JOIN " & _
'    " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod INNER JOIN ColocacCred ON Colocaciones.cCtaCod = ColocacCred.cCtaCod " & _
'    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
'    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
'    " AND (SUBSTRING(Producto.cCtaCod, 4, 2) IN (" & psAgencias & ")) "
'    If Len(Trim(psMoneda)) > 0 Then
'        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
'    End If
'
'    If Len(Trim(psProductos)) > 0 Then
'        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 6, 3) IN (" & psProductos & ")) "
'    End If
'
'    lsSQL = lsSQL & " AND (ColocacCred.nColocCondicion<>" & gColocCredCondNormal & ") "
'    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)>'" & Format(dFechaPasada, "YYYYMMdd") & "') "
'    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)<= '" & Format(pdFecFin, "YYYYmmdd") & "') "
'    lsSQL = lsSQL & " AND nPrdEstado Not In (" & gColocEstSug & "," & gColocEstAprob & ","
'    lsSQL = lsSQL & gColocEstRech & "," & gColocEstRetirado & "," & gColocEstRefMor & "," & gColocEstRefVenc & "," & gColocEstRefNorm & ") GROUP BY RRHH.cUser "
'
''5) Obtengo los desembolsos Refinanciados            La fecha es el rango entre el 31 del mes pasado a la fecha consultada y la fecha consultada
'                                                    'Aca busco el estado 2030 a 2032
'' para la cartera vigente
'    lsSQL = lsSQL & " UNION ALL "
'    lsSQL = lsSQL & " SELECT 5 AS Lugar, COUNT(Colocaciones.cCtaCod) AS Cantidad, " & _
'    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Colocaciones.nMontoCol " & _
'    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Colocaciones.nMontoCol * " & pnTipoCambio & " " & _
'    " End) AS Total, RRHH.cUser AS Analista " & _
'    " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN " & _
'    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod INNER JOIN " & _
'    " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod " & _
'    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
'    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
'    " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "')) " & _
'    " AND (SUBSTRING(Producto.cCtaCod, 4, 2) IN (" & psAgencias & ")) "
'    If Len(Trim(psMoneda)) > 0 Then
'        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
'    End If
'    If Len(Trim(psProductos)) > 0 Then
'        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 6, 3) IN (" & psProductos & ")) "
'    End If
'
'    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)>'" & Format(dFechaPasada, "YYYYMMdd") & "') "
'    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)<= '" & Format(pdFecFin, "YYYYmmdd") & "') "
'    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
    
'----------------------------------------------------
    
'6) ======= CARTERA MOROSA ========

    lsSQL = lsSQL & " UNION ALL "
     
If pnTipoFecha = 1 Then 'Fecha Actual
    'Para sacar los morosos a la fecha:
      'Para el vigente vencido y Refinanciado Vencido se saca de Producto
      'Para los Vigentes Normales y Morosos y Refinanciados y Morosos
         'Si es producto 1 o 2 se saca de producto
         'Si es producto 3 o 4 se saca de las cuotas pendientes con fecha de vencimiento menor a la del reporte
         'Considerar Calendario Normal y Calendario Paralelo
    'WIOR 20130627 SE CAMBIO DE 9 AS LUGAR A 10 AS LUGAR
    lsSQL = lsSQL & " Select 10 AS Lugar, SUM(T.nCantidad) as Cantidad, SUM(T.nMonto) as Total, T.Analista "
    lsSQL = lsSQL & " From "
    lsSQL = lsSQL & " ( " ' -- Solo los vencidos
    lsSQL = lsSQL & " Select RH.cUser AS Analista, count(p.cCtaCod) as nCantidad, SUM("
    lsSQL = lsSQL & " CASE WHEN Substring(P.cCtaCod, 9,1) = '1' THEN P.nSaldo ELSE P.nSaldo * "
    lsSQL = lsSQL & pnTipoCambio & " END) as nMonto "
    lsSQL = lsSQL & " From producto P INNER JOIN ProductoPersona PP ON P.cCtaCod=PP.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN RRHH RH ON PP.cPersCod = RH.cPersCod "
    lsSQL = lsSQL & " Where "
    lsSQL = lsSQL & " PP.nPrdPersRelac='" & gColRelPersAnalista & "' "
    lsSQL = lsSQL & " AND PP.cPersCod IN (" & psAnalistas & ") " '--Analistas
    lsSQL = lsSQL & " AND p.nPrdEstado  in ('" & gColocEstVigVenc & "', '" & gColocEstRefVenc & "') " ' -- Vencidos
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ") "
    End If
    lsSQL = lsSQL & " And C.cAgeCodAct IN (" & psAgencias & ")  " '-- Agencia
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND C.cTpoCredCod IN (" & psTipoCredito & ") " '--Productos
    End If
    lsSQL = lsSQL & " Group By RH.cUser "
   
    lsSQL = lsSQL & " Union All "
    
    '-- Final para morosos
    lsSQL = lsSQL & " Select RH.cUser AS Analista, COUNT(DISTINCT P.cCtaCod) as nCantidad, "
    lsSQL = lsSQL & " SUM(CASE WHEN Substring(P.cCtaCod, 9,1) = '1' THEN (CCD.nMonto - CCD.nMontoPagado) "
    lsSQL = lsSQL & " ELSE (CCD.nMonto - CCD.nMontoPagado) * " & pnTipoCambio & " END ) as nMonto "
    lsSQL = lsSQL & " FROM Producto P INNER JOIN ProductoPersona PP ON P.cCtaCod=PP.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Colocaciones C ON C.cCtaCod = P.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN RRHH RH ON PP.cPersCod = RH.cPersCod "
    lsSQL = lsSQL & " INNER JOIN ColocacCred CCR ON P.cCtaCod=CCR.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Coloccalendario CC ON CC.cCtaCod=CCR.cCtaCod AND CC.nNroCalen = CCR.nNroCalen "
    lsSQL = lsSQL & " INNER JOIN ColocCalendDet CCD ON CC.cCtaCod=CCD.cCtaCod AND CC.nNroCalen = CCD.nNroCalen "
    lsSQL = lsSQL & " AND CC.nColocCalendApl=CCD.nCOlocCalendApl AND CC.nCuota = CCD.nCuota "
    lsSQL = lsSQL & " Where"
    lsSQL = lsSQL & " PP.nPrdPersRelac='" & gColRelPersAnalista & "' "
    lsSQL = lsSQL & " AND PP.cPersCod IN (" & psAnalistas & ") " '--Analistas
    lsSQL = lsSQL & " and CC.ncoloccalendapl=" & gColocCalendAplCuota & " " '-- Cuota = 1
'    lsSQL = lsSQL & " And convert(varchar(8), CC.dvenc, 112)<='" & Format(pdFecFin, "YYYYMMdd") & "' " '-- Fecha de Vencimiento <='fecha de reporte
    lsSQL = lsSQL & " And datediff(dd, CC.dvenc, '" & Format(pdFecFin, "mm/dd/yyyy") & "') Between 30 AND 91 " '-- Fecha de Vencimiento <='fecha de reporte
    lsSQL = lsSQL & " And CC.nColocCalendEstado=" & gColocCalendEstadoPendiente & " " '-- Vigente 0
    lsSQL = lsSQL & " AND CCD.nPrdConceptoCod=" & gColocConceptoCodCapital & " " '1000 -- Capital
    lsSQL = lsSQL & " AND P.nPrdEstado in('" & gColocEstVigMor & "', '" & gColocEstRefMor & "') " '-- Normal y Moroso
    lsSQL = lsSQL & " And substring(C.cTpoCredCod,1,1) IN ('6','7','8') " '-- No vencidos pero Consumo e Hipotecario
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ") "
    End If
    lsSQL = lsSQL & " And C.cAgeCodAct IN (" & psAgencias & ") " ' -- Agencia
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND C.cTpoCredCod IN (" & psTipoCredito & ") " ' --Productos
    End If
    lsSQL = lsSQL & " Group By RH.cUser "
    
    lsSQL = lsSQL & " Union All "
    
    lsSQL = lsSQL & " Select RH.cUser AS Analista, COUNT(DISTINCT P.cCtaCod) as nCantidad, "
    lsSQL = lsSQL & " SUM( CASE WHEN Substring(P.cCtaCod, 9,1) = '1' THEN (CCD.nMonto - CCD.nMontoPagado) "
    lsSQL = lsSQL & " ELSE (CCD.nMonto - CCD.nMontoPagado) * " & pnTipoCambio & " END) as nMonto "
    lsSQL = lsSQL & " FROM Producto P "
    lsSQL = lsSQL & " INNER JOIN Colocaciones C ON C.cCtaCod = P.cCtaCod"
    lsSQL = lsSQL & " INNER JOIN ProductoPersona PP ON P.cCtaCod=PP.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN RRHH RH ON PP.cPersCod = RH.cPersCod "
    lsSQL = lsSQL & " INNER JOIN ColocacCred CCR ON P.cCtaCod=CCR.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Coloccalendario CC ON CC.cCtaCod=CCR.cCtaCod AND CC.nNroCalen = CCR.nNroCalPar "
    lsSQL = lsSQL & " INNER JOIN ColocCalendDet CCD ON CC.cCtaCod=CCD.cCtaCod AND CC.nNroCalen = CCD.nNroCalen "
    lsSQL = lsSQL & " AND CC.nColocCalendApl=CCD.nCOlocCalendApl AND CC.nCuota = CCD.nCuota "
    lsSQL = lsSQL & " Where "
    lsSQL = lsSQL & " PP.nPrdPersRelac='" & gColRelPersAnalista & "' "
    lsSQL = lsSQL & " AND PP.cPersCod IN (" & psAnalistas & ") " '--Analistas
    lsSQL = lsSQL & " And CC.ncoloccalendapl=" & gColocCalendAplCuota & " " '1 -- Cuota
'    lsSQL = lsSQL & " And convert(varchar(8), CC.dvenc, 112) <='" & Format(pdFecFin, "YYYYMMdd") & "' " '-- Fecha de Vencimiento <='fecha de reporte
    lsSQL = lsSQL & " And datediff(dd, CC.dvenc, '" & Format(pdFecFin, "mm/dd/yyyy") & "') Between 30 AND 91 " '-- Fecha de Vencimiento <='fecha de reporte
    lsSQL = lsSQL & " And CC.nColocCalendEstado=" & gColocCalendEstadoPendiente & " " '-- Vigente 0
    lsSQL = lsSQL & " AND CCD.nPrdConceptoCod=" & gColocConceptoCodCapital & " "   '1000 " '-- Capital
    lsSQL = lsSQL & " AND P.nPrdEstado in('" & gColocEstVigMor & "', '" & gColocEstRefMor & "') " '-- Normal y Moroso
    lsSQL = lsSQL & " AND substring(C.cTpoCredCod,1,1) IN ('6','7','8') " '-- No vencidos pero Consumo e Hipotecario
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ") "
    End If
    lsSQL = lsSQL & " And C.cAgeCodAct IN (" & psAgencias & ") " ' -- Agencia
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND C.cTpoCredCod IN (" & psTipoCredito & ") " ' --Productos
    End If
    lsSQL = lsSQL & " Group By RH.cUser "
    
    lsSQL = lsSQL & " ) T "
    lsSQL = lsSQL & " Group By T.Analista "

ElseIf pnTipoFecha = 2 Then 'Fecha pasada
    'WIOR 20130627 SE CAMBIO DE 9 AS LUGAR A 10 AS LUGAR
    lsSQL = lsSQL & " SELECT 10 AS Lugar, COUNT(CSC.cCtaCod) AS Cantidad, " & _
    " SUM(CASE WHEN SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaNacional & "' then CSC.nSaldoCap " & _
    " when SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then CSC.nSaldoCap * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod=Producto.cCtaCod INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod " & _
    " INNER JOIN " & lsBaseConsol & "CreditoSaldoConsol CSC ON ProductoPersona.cCtaCod = CSC.cCtaCod " & _
    " INNER JOIN " & lsBaseConsol & "CreditoConsolTotal CCT ON CreditoConsolTotal.cCtaCod = CSC.cCtaCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (CONVERT(VARCHAR(8), CSC.dFecha, 112)='" & Format(pdFecFin, "YYYYMMdd") & "') " & _
    " AND (Producto.nPrdEstado IN ('" & gColocEstVigMor & "', '" & gColocEstRefMor & "','" & gColocEstVigVenc & "','" & gColocEstRefVenc & "')) " & _
    " AND (CCT.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CSC.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (CCT.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
End If

'7) Obtengo los saldos del mes pasado (Para todos no inportando el tipo de fecha)
'    para la cartera morosa
    'WIOR 20130627 SE CAMBIO DE 10 AS LUGAR A 11 AS LUGAR
    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 11 AS Lugar, COUNT(CSC.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaNacional & "' then CSC.nCapVencido " & _
    " when SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then CSC.nCapVencido * " & nTipoCambioPasado & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod " & _
    " INNER JOIN " & lsBaseConsol & "CreditoSaldoConsol CSC ON ProductoPersona.cCtaCod = CSC.cCtaCod " & _
    " INNER JOIN " & lsBaseConsol & "CreditoConsolTotal CCT ON CCT.cCtaCod = CSC.cCtaCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (CONVERT(VARCHAR(8), CSC.dFecha, 112)='" & Format(dFechaPasada, "YYYYMMdd") & "') " & _
    " AND (CSC.nPrdEstado IN ('" & gColocEstVigMor & "', '" & gColocEstRefMor & "','" & gColocEstVigVenc & "','" & gColocEstRefVenc & "')) " & _
    " AND (CCT.cAgeCodAct IN (" & psAgencias & "))  "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CSC.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (CCT.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    
    lsSQL = lsSQL & " GROUP BY RRHH.cUser"
  
    ' 6) Obtengo los ingresos a recuperaciones
    'WIOR 20130627 SE CAMBIO DE 11 AS LUGAR A 12 AS LUGAR
    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " Select 12 as Lugar, COUNT(T.cCtaCod) as Cantidad, "
    lsSQL = lsSQL & " SUM( "
    lsSQL = lsSQL & " case when SUBSTRING(T.cCtaCod, 9, 1)='1' then T.nCapital "
    lsSQL = lsSQL & " when SUBSTRING(T.cCtaCod, 9, 1)='2' then T.nCapital * " & pnTipoCambio & " end "
    lsSQL = lsSQL & " ) as Total, RH.cUser  "
    lsSQL = lsSQL & " From ( "
    lsSQL = lsSQL & " Select CR.cCtaCod, "
    lsSQL = lsSQL & " IsNull( "
    lsSQL = lsSQL & " ( SELECT SUM(nmonto)  FROM movcoldet MCD INNER JOIN MOV M ON MCD.nMovNro = M.nMovNro "
    lsSQL = lsSQL & " WHERE M.nMovFlag=0 and MCD.cCtaCod = CR.cCtaCod AND MCD.nPrdConceptoCod like '14%' "
    lsSQL = lsSQL & " And MCD.cOpecod = '130100'), 0)  AS nCapital "
    lsSQL = lsSQL & " FROM ColocRecup CR "
    lsSQL = lsSQL & " where convert(varchar(8),CR.dIngRecup,112)>='" & Format(dFechaPasada, "YYYYMMdd") & "' And convert(varchar(8), CR.dIngRecup,112)<='" & Format(pdFecFin, "YYYYMMdd") & "' "
    lsSQL = lsSQL & " ) T "
    lsSQL = lsSQL & " INNER JOIN Colocaciones C ON C.cCtaCod = T.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN ColocRecup CR ON T.cCtaCod = CR.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Producto PC ON T.cCtaCod = PC.cCtaCod "
    lsSQL = lsSQL & " "
    lsSQL = lsSQL & " INNER JOIN ProductoPersona PP ON PP.cCtaCod=PC.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN RRHH RH ON PP.cPersCod = RH.cPersCod "
    lsSQL = lsSQL & " WHERE "
    
    lsSQL = lsSQL & " (PP.nPrdPersRelac='" & gColRelPersAnalista & "') "
    lsSQL = lsSQL & " AND (PP.cPersCod IN (" & psAnalistas & ")) "

    lsSQL = lsSQL & " AND (PC.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "', '" & gColocEstRecVigCast & "', '" & gColocEstRecCanCast & "')) "
    lsSQL = lsSQL & " AND convert(varchar(8), CR.dIngRecup, 112) >'" & Format(dFechaPasada, "YYYYMMdd") & "' and convert(varchar(8), CR.dIngRecup, 112)<='" & Format(pdFecFin, "YYYYMMdd") & "' "
    lsSQL = lsSQL & " AND (C.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(T.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    
    lsSQL = lsSQL & " group by RH.cUser "
  
    lsSQL = lsSQL & " ORDER BY analista, lugar"
    
    'Lugar: 1 si es total, 2 si es saldos , 3 si es desembolsos nuevos, 4 si son desembolsos represtamos
    '       5 si son desembolsos refinanciados, 6 total para la cartera morosa, 7 saldos hasta el mes pasado
    '       para la cartera morosa, 8 ingresos a recuperacioness
 
 
    
    Dim oConn As COMConecta.DCOMConecta
    Dim sqlConn As String
    Dim RRep As ADODB.Recordset
    Set oConn = New COMConecta.DCOMConecta
    oConn.AbreConexion
    
'ARCV 12-06-2007
'    Set RRep = oConn.CargaRecordSet("Select * from constsistema where  nconssiscod = 74")
'    sqlConn = RRep!nConssisValor
'    oConn.CierraConexion
'    Call oConn.AbreConexion(sqlConn)
 
    Set loDataRep = New dColPFunciones
        'Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
        Set lrDataRep = oConn.CargaRecordSet(lsSQL)
    Set loDataRep = Nothing

      oConn.CierraConexion
       Set oConn = Nothing
 

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        lrDataRep.Close
        Set lrDataRep = Nothing
        Exit Function
    Else
        nContador = 0
        sTempoAnalista = IIf(IsNull(lrDataRep!ANALISTA), "", lrDataRep!ANALISTA)

        ReDim Preserve matAnalista(nContador) As String
        ReDim Preserve matVTot1(nContador) As Long
        ReDim Preserve matVTot2(nContador) As Currency
        ReDim Preserve matVSal1(nContador) As Long
        ReDim Preserve matVSal2(nContador) As Currency
        ReDim Preserve matVDesemNue1(nContador) As Long
        ReDim Preserve matVDesemNue2(nContador) As Currency
        ReDim Preserve matVDesemRep1(nContador) As Long
        ReDim Preserve matVDesemRep2(nContador) As Currency
        ReDim Preserve matVDesemRef1(nContador) As Long
        ReDim Preserve matVDesemRef2(nContador) As Currency
        ReDim Preserve matMTot1(nContador) As Long
        ReDim Preserve matMTot2(nContador) As Currency
        ReDim Preserve matMSal1(nContador) As Long
        ReDim Preserve matMSal2(nContador) As Currency
        ReDim Preserve matMIngRecup(nContador) As Currency
 
        'ARCV 23-06-2007
        ReDim Preserve matVDesemRec1(nContador) As Long
        ReDim Preserve matVDesemRec2(nContador) As Currency
        ReDim Preserve matVDesemAut1(nContador) As Long
        ReDim Preserve matVDesemAut2(nContador) As Currency
        ReDim Preserve matVDesemAmp1(nContador) As Long
        ReDim Preserve matVDesemAmp2(nContador) As Currency
        '-------
        'WIOR 20130627 ***************************************
        ReDim Preserve matVDesemAdi1(nContador) As Long
        ReDim Preserve matVDesemAdi2(nContador) As Currency
        'WIOR FIN ********************************************
 
        matAnalista(0) = IIf(IsNull(lrDataRep!ANALISTA), "", lrDataRep!ANALISTA)

        RaiseEvent ShowProgress
        
        Do While Not lrDataRep.EOF
            If sTempoAnalista <> lrDataRep!ANALISTA Then
                nContador = nContador + 1
                ReDim Preserve matAnalista(nContador) As String
                ReDim Preserve matVTot1(nContador) As Long
                ReDim Preserve matVTot2(nContador) As Currency
                ReDim Preserve matVSal1(nContador) As Long
                ReDim Preserve matVSal2(nContador) As Currency
                ReDim Preserve matVDesemNue1(nContador) As Long
                ReDim Preserve matVDesemNue2(nContador) As Currency
                ReDim Preserve matVDesemRep1(nContador) As Long
                ReDim Preserve matVDesemRep2(nContador) As Currency
                ReDim Preserve matVDesemRef1(nContador) As Long
                ReDim Preserve matVDesemRef2(nContador) As Currency
                ReDim Preserve matMTot1(nContador) As Long
                ReDim Preserve matMTot2(nContador) As Currency
                ReDim Preserve matMSal1(nContador) As Long
                ReDim Preserve matMSal2(nContador) As Currency
                ReDim Preserve matMIngRecup(nContador) As Currency

                'ARCV 23-06-2007
                ReDim Preserve matVDesemRec1(nContador) As Long
                ReDim Preserve matVDesemRec2(nContador) As Currency
                ReDim Preserve matVDesemAut1(nContador) As Long
                ReDim Preserve matVDesemAut2(nContador) As Currency
                ReDim Preserve matVDesemAmp1(nContador) As Long
                ReDim Preserve matVDesemAmp2(nContador) As Currency
                '-------
                'WIOR 20130627 ***************************************
                ReDim Preserve matVDesemAdi1(nContador) As Long
                ReDim Preserve matVDesemAdi2(nContador) As Currency
                'WIOR FIN ********************************************
        
                matAnalista(nContador) = lrDataRep!ANALISTA
                sTempoAnalista = lrDataRep!ANALISTA
            End If
            
            With lrDataRep
                Select Case !Lugar
                Case 1
                    matVTot1(nContador) = !cantidad
                    matVTot2(nContador) = Format(!Total, "#0.00")
                Case 2
                    matVSal1(nContador) = !cantidad
                    matVSal2(nContador) = Format(!Total, "#0.00")
                Case 3
                    matVDesemNue1(nContador) = !cantidad
                    matVDesemNue2(nContador) = Format(!Total, "#0.00")
                Case 4
                    'matVDesemRep1(nContador) = !cantidad
                    'matVDesemRep2(nContador) = Format(!Total, "#0.00")
                    matVDesemRec1(nContador) = !cantidad
                    matVDesemRec2(nContador) = Format(!Total, "#0.00")
                Case 5
                    'matVDesemRef1(nContador) = !cantidad
                    'matVDesemRef2(nContador) = Format(!Total, "#0.00")
                    matVDesemRep1(nContador) = !cantidad
                    matVDesemRep2(nContador) = Format(!Total, "#0.00")
                Case 6
                    'matMTot1(nContador) = !cantidad
                    'matMTot2(nContador) = Format(!Total, "#0.00")
                    matVDesemRef1(nContador) = !cantidad
                    matVDesemRef2(nContador) = Format(!Total, "#0.00")
                Case 7
                    'matMSal1(nContador) = !cantidad
                    'matMSal2(nContador) = Format(!Total, "#0.00")
                    matVDesemAmp1(nContador) = !cantidad
                    matVDesemAmp2(nContador) = Format(!Total, "#0.00")
                Case 8
                    'matMIngRecup(nContador) = Format(!Total, "#0.00")
                    matVDesemAut1(nContador) = !cantidad
                    matVDesemAut2(nContador) = Format(!Total, "#0.00")
                'ARCV 23-06-2007
                'WIOR 20130627
                 Case 9
                    matVDesemAdi1(nContador) = !cantidad
                    matVDesemAdi2(nContador) = Format(!Total, "#0.00")
                'WIOR FIN
                Case 10 'WIOR 20130627 SE CAMBIO DE 9 A 10
                    'matMIngRecup(nContador) = Format(!Total, "#0.00")
                    matMTot1(nContador) = !cantidad
                    matMTot2(nContador) = Format(!Total, "#0.00")
                Case 11 'WIOR 20130627 SE CAMBIO DE 10 A 11
                    'matMIngRecup(nContador) = Format(!Total, "#0.00")
                    matMSal1(nContador) = !cantidad
                    matMSal2(nContador) = Format(!Total, "#0.00")
                Case 12 'WIOR 20130627 SE CAMBIO DE 11 A 12
                    'matMIngRecup(nContador) = Format(!Total, "#0.00")
                    matMIngRecup(nContador) = Format(!Total, "#0.00")
                '--------------
               
                End Select
            End With
            
            RaiseEvent Progress(lrDataRep.Bookmark, lrDataRep.RecordCount)
            lrDataRep.MoveNext
            
        Loop
        lrDataRep.Close
        Set lrDataRep = Nothing
            
        RaiseEvent CloseProgress
            
        'Calculo los totales
        For i = 0 To nContador
           
            'Vigentes
            '========
            
            'Total de Numero para Totales
            matTot1(0) = matTot1(0) + matVTot1(i)
            'Total de Totales para Totales
            matTot1(1) = matTot1(1) + matVTot2(i)
            'Total de Numero para Saldos
            matTot1(2) = matTot1(2) + matVSal1(i)
            'Total de Totales para Saldos
            matTot1(3) = matTot1(3) + matVSal2(i)
            'Variacion a la fecha para vigente de numero
            matTot1(4) = matTot1(4) + matVTot1(i) - matVSal1(i)
            'Variacion a la fecha para vigente de monto
            matTot1(5) = matTot1(5) + matVTot2(i) - matVSal2(i)
            'Total de Desembolsos Nuevos para cantidad
            matTot1(6) = matTot1(6) + matVDesemNue1(i)
            'Total de Desembolsos Nuevos para Total
            matTot1(7) = matTot1(7) + matVDesemNue2(i)
            'Total de Desembolsos Represtamos para cantidad
            matTot1(8) = matTot1(8) + matVDesemRep1(i)
            'Total de Desembolsos Represtamos para Total
            matTot1(9) = matTot1(9) + matVDesemRep2(i)
            'Total de Desembolsos Refinanciados para cantidad
            matTot1(10) = matTot1(10) + matVDesemRef1(i)
            'Total de Desembolsos Refinanciados para Total
            matTot1(11) = matTot1(11) + matVDesemRef2(i)
            
            'ARCV 22-06-2007
            matTot1(14) = matTot1(14) + matVDesemRec1(i)
            matTot1(15) = matTot1(15) + matVDesemRec2(i)
            matTot1(16) = matTot1(16) + matVDesemAmp1(i)
            matTot1(17) = matTot1(17) + matVDesemAmp2(i)
            matTot1(18) = matTot1(18) + matVDesemAut1(i)
            matTot1(19) = matTot1(19) + matVDesemAut2(i)
            '---------
            'WIOR 20130627 *****************************
            matTot1(20) = matTot1(20) + matVDesemAdi1(i)
            matTot1(21) = matTot1(21) + matVDesemAdi2(i)
            'WIOR FIN **********************************
            
            'ARCV 22-06-2007
            'total de Desembolsos cantidad
            'matTot1(12) = matTot1(12) + matVDesemNue1(i) + matVDesemRep1(i) + matVDesemRef1(i)
            'total de Desembolsos Total
            'matTot1(13) = matTot1(13) + matVDesemNue2(i) + matVDesemRep2(i) + matVDesemRef2(i)
            'total de Desembolsos cantidad
            'WIOR 20130627 SE AGREGO EN EL CONTEO matVDesemAdi1(i)
            matTot1(12) = matTot1(12) + matVDesemNue1(i) + matVDesemRep1(i) + matVDesemRef1(i) + matVDesemRec1(i) + matVDesemAmp1(i) + matVDesemAut1(i) + matVDesemAdi1(i)
            'total de Desembolsos Total
            'WIOR 20130627 SE AGREGO EN EL CONTEO matVDesemAdi2(i)
            matTot1(13) = matTot1(13) + matVDesemNue2(i) + matVDesemRep2(i) + matVDesemRef2(i) + matVDesemRec2(i) + matVDesemAmp2(i) + matVDesemAut2(i) + matVDesemAdi2(i)
            '------------

            'Morosos
            '=======
             
            'Total de Numero para Totales
            matTot2(0) = matTot2(0) + matMTot1(i)
            'Total de Totales para Totales
            matTot2(1) = matTot2(1) + matMTot2(i)
            'Total de Numero para Saldos
            matTot2(2) = matTot2(2) + matMSal1(i)
            'Total de Totales para Saldos
            matTot2(3) = matTot2(3) + matMSal2(i)
            'Variacion a la fecha para vigente de numero
            matTot2(4) = matTot2(4) + matMTot1(i) - matMSal1(i)
            'Variacion a la fecha para vigente de monto
            matTot2(5) = matTot2(5) + matMTot2(i) - matMSal2(i)
            'SUma de Ingresos a Recuperaciones
            matTot2(6) = matTot2(6) + matMIngRecup(i)
            'Variacion Real
            matTot2(7) = matTot2(7) + matMTot2(i) - matMSal2(i) + matMIngRecup(i)
            '% de Mora
            If matMTot2(i) = 0 Then
            Else
                matTot2(8) = matTot2(8) + (matMTot2(i) / matVTot2(i) * 100)
            End If
        Next
                
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CUADRO CONSOLIDADO DE COLOCACIONES X ANALISTA", lsCondiciones, lnPage, 168, cSubTit & Chr(10) & "TCF. = " & Format(pnTipoCambio, "0.000") & " // TCF. (" & dFechaPasada & ") = " & Format(nTipoCambioPasado, "0.000"), 108602)
        lsCadImp = lsCadImp & Chr(10)
        
        lsCadImp = lsCadImp & "C A R T E R A      V I G E N T E" & Chr(10)
        lsCadImp = lsCadImp & String(260, "-") & Chr(10)
                  lsCadImp = lsCadImp & "        |    T  O  T  A  L     |   S A L D O S  A L   |  V A R I A C I O N   |                            D  E  S  E  M  B  O  L  S  O  S             D  E  L                 M  E  S                                                         |   T   O   T   A   L  " & Chr(10)    'WIOR 20130627 SE AMPLIO LA COLUMNA
        lsCadImp = lsCadImp & "ANALISTA|  (En Nuevos Soles)   |      " & dFechaPasada & "      |  A  L A  F E C H A   |       N U E V O S    |       RECURRENTES    |         PARALELOS    |     REFINANCIADOS    |       AMPLIADOS      |     AUTOMATICOS      |      ADICIONALES     |        DESEMBOLSOS    " & Chr(10)   'WIOR 20130627 SE AGREGO LA COLUMNA DE ADICIONALES
                  lsCadImp = lsCadImp & "        | NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |    M O N T O  | NUM. |   M O N T O   | NUM. |    M O N T O  | NUM. |    M O N T O  | NUM. |    M O N T O  | NUM. |    M O N T O  | NUM. |    M O N T O  | NUM. |     M O N T O  " & Chr(10)   'WIOR 20130627 SE AGREGO LA COLUMNA DE ADICIONALES
        lsCadImp = lsCadImp & String(260, "-") & Chr(10)
        
        lnIndice = 14:  lnLineas = 14
             
        For i = 0 To nContador
            lnIndice = lnIndice + 1
            lnLineas = lnLineas + 1
 
            cUser = matAnalista(i)
            'WIOR 20130627 SE AGREGO EN EL CONTEO matVDesemAdi1(i) y matVDesemAdi2(i)
            lsCadImp = lsCadImp & Space(1) & cUser & Space(2) & _
                    " | " & ImpreFormat(matVTot1(i), 4, 0) & " | " & ImpreFormat(matVTot2(i), 10, 2) & _
                    " | " & ImpreFormat(matVSal1(i), 4, 0) & " | " & ImpreFormat(matVSal2(i), 10, 2) & _
                    " | " & ImpreFormat(matVTot1(i) - matVSal1(i), 4, 0) & " | " & ImpreFormat(matVTot2(i) - matVSal2(i), 10, 2) & _
                    " | " & ImpreFormat(matVDesemNue1(i), 4, 0) & " | " & ImpreFormat(matVDesemNue2(i), 10, 2) & _
                    " | " & ImpreFormat(matVDesemRec1(i), 4, 0) & " | " & ImpreFormat(matVDesemRec2(i), 10, 2) & _
                    " | " & ImpreFormat(matVDesemRep1(i), 4, 0) & " | " & ImpreFormat(matVDesemRep2(i), 10, 2) & _
                    " | " & ImpreFormat(matVDesemRef1(i), 4, 0) & " | " & ImpreFormat(matVDesemRef2(i), 10, 2) & _
                    " | " & ImpreFormat(matVDesemAmp1(i), 4, 0) & " | " & ImpreFormat(matVDesemAmp2(i), 10, 2) & _
                    " | " & ImpreFormat(matVDesemAut1(i), 4, 0) & " | " & ImpreFormat(matVDesemAut2(i), 10, 2) & _
                    " | " & ImpreFormat(matVDesemAdi1(i), 4, 0) & " | " & ImpreFormat(matVDesemAdi2(i), 10, 2) & _
                    " | " & ImpreFormat(matVDesemNue1(i) + matVDesemRep1(i) + matVDesemRef1(i) + matVDesemRec1(i) + matVDesemAmp1(i) + matVDesemAut1(i) + matVDesemAdi1(i), 4, 0) & " | " & ImpreFormat(matVDesemNue2(i) + matVDesemRep2(i) + matVDesemRef2(i) + matVDesemRec2(i) + matVDesemAmp2(i) + matVDesemAut2(i) + matVDesemAdi2(i), 10, 2) & _
                    Chr(10)
          
            If lnIndice Mod 300 = 0 Then
                lsCadBuffer = lsCadBuffer & lsCadImp
                lsCadImp = ""
            End If
        
            If lnLineas >= 55 Then
                lnPage = lnPage + 1
                lsCadImp = lsCadImp & Chr(12)
                lsCadImp = lsCadImp & nRepoCabecera("CUADRO CONSOLIDADO DE COLOCACIONES X ANALISTA", lsCondiciones, lnPage, 168, cSubTit & Chr(10) & "TCF. = " & Format(pnTipoCambio, "0.000") & " // TCF. (" & dFechaPasada & ") = " & Format(nTipoCambioPasado, "0.000"), 108602)
                lnLineas = 7
                lnIndice = lnIndice + 7
            End If
        Next
        
        'Imprimo los totales de vigentes
        lsCadImp = lsCadImp & String(260, "=") & Chr(10)
        'WIOR 20130627 SE AGREGO EN EL CONTEO matTot1(8), matTot1(9), matTot1(20) y matTot1(21)
        lsCadImp = lsCadImp & Space(7) & " | " & ImpreFormat(matTot1(0), 4, 0) & " | " & ImpreFormat(matTot1(1), 10, 2) & _
                " | " & ImpreFormat(matTot1(2), 4, 0) & " | " & ImpreFormat(matTot1(3), 10, 2) & _
                " | " & ImpreFormat(matTot1(4), 4, 0) & " | " & ImpreFormat(matTot1(5), 10, 2) & _
                " | " & ImpreFormat(matTot1(6), 4, 0) & " | " & ImpreFormat(matTot1(7), 10, 2) & _
                " | " & ImpreFormat(matTot1(14), 4, 0) & " | " & ImpreFormat(matTot1(15), 10, 2) & _
                " | " & ImpreFormat(matTot1(8), 4, 0) & " | " & ImpreFormat(matTot1(9), 10, 2) & _
                " | " & ImpreFormat(matTot1(10), 4, 0) & " | " & ImpreFormat(matTot1(11), 10, 2) & _
                " | " & ImpreFormat(matTot1(16), 4, 0) & " | " & ImpreFormat(matTot1(17), 10, 2) & _
                " | " & ImpreFormat(matTot1(18), 4, 0) & " | " & ImpreFormat(matTot1(19), 10, 2) & _
                " | " & ImpreFormat(matTot1(20), 4, 0) & " | " & ImpreFormat(matTot1(21), 10, 2) & _
                " | " & ImpreFormat(matTot1(12), 4, 0) & " | " & ImpreFormat(matTot1(13), 10, 2) & _
                 Chr(10)
        
        lsCadImp = lsCadImp & String(260, "_") & Chr(10) & Chr(10) & Chr(10)
        
        lsCadImp = lsCadImp & "C A R T E R A        M O R O S A" & Chr(10)
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lsCadImp = lsCadImp & "        |    T  O  T  A  L     |   S A L D O S  A L   |  V A R I A C I O N   |    INGRESO    |   VARIACION   |               " & Chr(10)
        lsCadImp = lsCadImp & "ANALISTA|  (En Nuevos Soles)   |      " & dFechaPasada & "      |  A  L A  F E C H A   | RECUPERACIONES|  R  E  A  L   |   % M O R A   " & Chr(10)
        lsCadImp = lsCadImp & "        | NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |   M O N T O   | (Nuevos Soles)|               |               " & Chr(10)
        
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        
        lnIndice = lnIndice + 10:  lnLineas = lnLineas + 10
             
        For i = 0 To nContador
            lnIndice = lnIndice + 1
            lnLineas = lnLineas + 1
            cUser = matAnalista(i)
            
            If matVTot2(i) = 0 Then
                ntempo1 = 0
            Else
                ntempo1 = ((matMTot2(i) / matVTot2(i)) * 100)
            End If
            
            lsCadImp = lsCadImp & Space(1) & cUser & Space(2) & _
                    " | " & ImpreFormat(matMTot1(i), 4, 0) & " | " & ImpreFormat(matMTot2(i), 10, 2) & _
                    " | " & ImpreFormat(matMSal1(i), 4, 0) & " | " & ImpreFormat(matMSal2(i), 10, 2) & _
                    " | " & ImpreFormat(matMTot1(i) - matMSal1(i), 4, 0) & " | " & ImpreFormat(matMTot2(i) - matMSal2(i), 10, 2) & _
                    " | " & ImpreFormat(matMIngRecup(i), 10, 2) & _
                    " | " & ImpreFormat(matMSal2(i) + matMIngRecup(i), 10, 2) & _
                    " | " & ImpreFormat(ntempo1, 10, 2) & _
                    Chr(10)
           
          
          
            If lnIndice Mod 300 = 0 Then
                lsCadBuffer = lsCadBuffer & lsCadImp
                lsCadImp = ""
            End If
        
            If lnLineas >= 55 Then
                lnPage = lnPage + 1
                lsCadImp = lsCadImp & Chr(12)
                lsCadImp = lsCadImp & nRepoCabecera("CUADRO CONSOLIDADO DE COLOCACIONES X ANALISTA", lsCondiciones, lnPage, 168, cSubTit & Chr(10) & "TCF. = " & Format(pnTipoCambio, "0.000") & " // TCF. (" & dFechaPasada & ") = " & Format(nTipoCambioPasado, "0.000"), 108602)
                lnLineas = 7
                lnIndice = lnIndice + 7
            End If
        Next
        
        'Imprimo los totales de morosos
        lsCadImp = lsCadImp & String(130, "=") & Chr(10)
        
        '**Modificado por DAOR 20081001********************************
        'ntempo1 = Round(matTot2(1) / matTot1(1) * 100, 2)
        If (matTot1(1) > 0) Then
            ntempo1 = Round(matTot2(1) / matTot1(1) * 100, 2)
        Else
            ntempo1 = 0
        End If
        '**************************************************************
        
        lsCadImp = lsCadImp & Space(7) & " | " & ImpreFormat(matTot2(0), 4, 0) & " | " & ImpreFormat(matTot2(1), 10, 2) & _
                    " | " & ImpreFormat(matTot2(2), 4, 0) & " | " & ImpreFormat(matTot2(3), 10, 2) & _
                    " | " & ImpreFormat(matTot2(4), 4, 0) & " | " & ImpreFormat(matTot2(5), 10, 2) & _
                    " | " & ImpreFormat(matTot2(6), 10, 2) & " | " & ImpreFormat(matTot2(7), 10, 2) & _
                    " | " & ImpreFormat(ntempo1, 10, 2) & _
                     Chr(10)
        
        lsCadImp = lsCadImp & String(130, "_") & Chr(10) & Chr(10)

        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
 
    End If
      nRepo108602_ConsolidadoColocacionesxAnalista = lsCadBuffer
    
End Function

Public Function nRepo108603_CuadroMetasAlcanzadasxAnalista(ByVal pnTipoFecha As Integer, ByVal cSubTit As String, ByVal pnTipoCambio As Currency, ByVal pdFecFin As String, ByVal psMoneda As String, ByVal psProductos As String, ByVal psAgencias As String, ByVal psAnalistas) As String

Dim matAnalista() As String
Dim matVTot1() As Long
Dim matVTot2() As Currency
Dim matVSal1() As Long
Dim matVSal2() As Currency
Dim matVMetaP() As Currency
Dim matVDesem1() As Long
Dim matVDesem2() As Currency
  
'Los totales de la parte superior
Dim matTot1(9) As Currency
Dim nContador As Integer
  
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim lrDataTemp As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim loImprimeCab As NColPImpre
        
Dim lsCondiciones As String

Dim oTipCambio As nTipoCambio
 
'Obtencion de la fecha del mes pasado
Dim dUltimoDia As DCredReporte
Dim nDiaPasado As Integer
Dim nMesPasado As Integer
Dim dFechaPasada As String
Dim nTipoCambioPasado As Currency

Dim i As Integer
Dim sTempoAnalista As String
Dim cUser As String * 4
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim ntempo1 As Currency
Dim lsBaseConsol As String

    lsSQL = "Select nConsSisValor FROM ConstSistema WHERE nConsSisCod = 43"

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
        lsBaseConsol = lrDataRep!nConsSisValor
    Set loDataRep = Nothing
    Set lrDataRep = Nothing

lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

    lsCondiciones = "Al " & pdFecFin

    Set dUltimoDia = New DCredReporte
    nMesPasado = IIf(Val(Mid(Format(pdFecFin, "dd/MM/YYYY"), 4, 2)) = 1, 12, Val(Mid(Format(pdFecFin, "dd/MM/YYYY"), 4, 2)) - 1)
    dFechaPasada = "01/" & IIf(nMesPasado < 10, "0" & Trim(Str(nMesPasado)), Trim(Str(nMesPasado))) & "/" & IIf(nMesPasado = 1, Trim(Str(Val(Mid(Format(pdFecFin, "dd/MM/YYYY"), 7, 4)) - 1)), Trim(Str(Val(Mid(Format(pdFecFin, "dd/MM/YYYY"), 7, 4)))))
    
    nDiaPasado = dUltimoDia.RecuperaUltimoDiaMes(dFechaPasada)
    dFechaPasada = IIf(nDiaPasado < 10, "0" & Trim(Str(nDiaPasado)), Trim(Str(nDiaPasado))) & "/" & Mid(dFechaPasada, 4, 7)
    'Ya tengo la fecha del ultimo dia del mes pasado
    'La fecha actual es la que pasan como parametro
    dFechaPasada = Format(dFechaPasada, "dd/MM/YYYY")
      
    Set oTipCambio = New nTipoCambio
    nTipoCambioPasado = Format(oTipCambio.EmiteTipoCambio(dFechaPasada, TCFijoMes), "0.00")
    Set oTipCambio = Nothing
    
'1) Obtengo el total para la cartera vigente
If pnTipoFecha = 1 Then 'Fecha Actual
    lsSQL = " SELECT 1 AS Lugar, COUNT(Producto.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Producto.nSaldo " & _
    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Producto.nSaldo * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
    " AND (SUBSTRING(Producto.cCtaCod, 4, 2) IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 6, 3) IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
ElseIf pnTipoFecha = 2 Then 'Fecha pasada(que llega como parametro)
    lsSQL = "SELECT 1 AS Lugar, COUNT(CSC.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaNacional & "' then CSC.nSaldoCap " & _
    " when SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then CSC.nSaldoCap * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod INNER JOIN " & lsBaseConsol & "CreditoSaldoConsol CSC ON ProductoPersona.cCtaCod = CSC.cCtaCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (convert(varchar(8), CSC.dFecha,112)='" & Format(pdFecFin, "yyyyMMdd") & "') " & _
    " AND (CSC.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
    " AND (SUBSTRING(CSC.cCtaCod, 4, 2) IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CSC.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CSC.cCtaCod, 6, 3) IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
End If


 
'2) Obtengo los saldos del mes pasado (Para todos no inportando el tipo de fecha)
'    para la cartera vigente

    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 2 AS Lugar, COUNT(CSC.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaNacional & "' then CSC.nSaldoCap " & _
    " when SUBSTRING(CSC.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then CSC.nSaldoCap * " & nTipoCambioPasado & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod INNER JOIN " & lsBaseConsol & "CreditoSaldoConsol CSC ON ProductoPersona.cCtaCod = CSC.cCtaCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (CONVERT(VARCHAR(8), CSC.dFecha, 112)='" & Format(dFechaPasada, "YYYYMMdd") & "') " & _
    " AND (CSC.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
    " AND (SUBSTRING(CSC.cCtaCod, 4, 2) IN (" & psAgencias & "))  "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CSC.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CSC.cCtaCod, 6, 3) IN (" & psProductos & ")) "
    End If
    
    lsSQL = lsSQL & " GROUP BY RRHH.cUser"

'3) Obtengo desembolsos nuevos      La fecha es el rango entre el 31 del mes pasado a la fecha consultada y la fecha consultada
'    y los represtamos              ' Aca busco los coloccondicion=1 y coloccondicion<>1 para los represtamos
' para la cartera vigente

    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 3 AS Lugar, COUNT(Colocaciones.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Colocaciones.nMontoCol " & _
    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Colocaciones.nMontoCol * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod INNER JOIN " & _
    " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod INNER JOIN ColocacCred ON Colocaciones.cCtaCod = ColocacCred.cCtaCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (SUBSTRING(Producto.cCtaCod, 4, 2) IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 6, 3) IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)>'" & Format(dFechaPasada, "YYYYMMdd") & "') "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)<= '" & Format(pdFecFin, "YYYYmmdd") & "') "
    lsSQL = lsSQL & " AND nPrdEstado Not In (" & gColocEstSug & "," & gColocEstAprob & ","
    lsSQL = lsSQL & gColocEstRech & "," & gColocEstRetirado & ") GROUP BY RRHH.cUser "
  
    ''" AND (ColocacCred.nColocCondicion=" & gColocCredCondNormal & ") " & _

  
'5) Obtengo los desembolsos Refinanciados            La fecha es el rango entre el 31 del mes pasado a la fecha consultada y la fecha consultada
                                                    'Aca busco el estado 2030 a 2032
' para la cartera vigente

    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 5 AS Lugar, COUNT(Colocaciones.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Colocaciones.nMontoCol " & _
    " when SUBSTRING(Producto.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Colocaciones.nMontoCol * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM ProductoPersona INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod INNER JOIN " & _
    " RRHH ON ProductoPersona.cPersCod = RRHH.cPersCod INNER JOIN " & _
    " Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod " & _
    " WHERE (ProductoPersona.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (ProductoPersona.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "')) " & _
    " AND (SUBSTRING(Producto.cCtaCod, 4, 2) IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(Producto.cCtaCod, 6, 3) IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)>'" & Format(dFechaPasada, "YYYYMMdd") & "') "
    lsSQL = lsSQL & " AND (CONVERT(varchar(8), Colocaciones.dVigencia, 112)<= '" & Format(pdFecFin, "YYYYmmdd") & "') "
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "

    
'30) Obtengo la meta propuesta

    lsSQL = lsSQL & " UNION ALL "
 
    lsSQL = lsSQL & " SELECT 30 AS Lugar, 0 AS Cantidad, Sum(Case nMoneda WHEN 1 THEN CM.nMonto WHEN 2 THEN CM.nMonto * " & pnTipoCambio & " END) AS Total,  cUser as Analista "
    lsSQL = lsSQL & " FROM ColocMetasAnalista CM INNER JOIN RRHH RH ON CM.cPersCod = RH.cPersCod "
    lsSQL = lsSQL & " Where CM.nTipoCred IN (" & psProductos & ") "
    lsSQL = lsSQL & " AND CM.nTipoMeta= " & gColocTipoMetasMenSual & " And CONVERT(VARCHAR(8), CM.dInicial, 112)<='" & Format(pdFecFin, "YYYYmmdd")
    lsSQL = lsSQL & " ' And CONVERT(VARCHAR(8), CM.dFinal, 112)>='" & Format(DateAdd("d", 1, dFechaPasada), "YYYYmmdd") & "'"
    lsSQL = lsSQL & " AND CM.cPersCod in (Select distinct cPersCod From ProductoPersona Where nPrdPersRelac = '" & gColRelPersAnalista & "')"
    lsSQL = lsSQL & " GROUP BY cUser"
    lsSQL = lsSQL & " ORDER BY cUser"
    
    'Lugar: 1 si es total, 2 si es saldos vigentes , 3 si es desembolsos nuevos y represtamos
    '       5 si son desembolsos refinanciados, 30 es meta propuesta
 
 
 Dim oConn As COMConecta.DCOMConecta
    Dim sqlConn As String
    Dim RRep As ADODB.Recordset
    Set oConn = New COMConecta.DCOMConecta
    oConn.AbreConexion
    
'ARCV 12-06-2007
'    Set RRep = oConn.CargaRecordSet("Select * from constsistema where  nconssiscod = 74")
'    sqlConn = RRep!nConssisValor
'    oConn.CierraConexion
'    Call oConn.AbreConexion(sqlConn)
 
    Set loDataRep = New dColPFunciones
        'Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
        Set lrDataRep = oConn.CargaRecordSet(lsSQL)
    Set loDataRep = Nothing
    
      oConn.CierraConexion
       Set oConn = Nothing
 

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        lrDataRep.Close
        Set lrDataRep = Nothing
        Exit Function
    Else
        nContador = 0
        sTempoAnalista = IIf(IsNull(lrDataRep!ANALISTA), "", lrDataRep!ANALISTA)

        ReDim Preserve matAnalista(nContador) As String
        ReDim Preserve matVTot1(nContador) As Long
        ReDim Preserve matVTot2(nContador) As Currency
        ReDim Preserve matVSal1(nContador) As Long
        ReDim Preserve matVSal2(nContador) As Currency
        ReDim Preserve matVMetaP(nContador) As Currency
        ReDim Preserve matVDesem1(nContador) As Long
        ReDim Preserve matVDesem2(nContador) As Currency
         
        matAnalista(0) = IIf(IsNull(lrDataRep!ANALISTA), "", lrDataRep!ANALISTA)
        
        RaiseEvent ShowProgress
        
        Do While Not lrDataRep.EOF
            If sTempoAnalista <> lrDataRep!ANALISTA Then
                nContador = nContador + 1
                ReDim Preserve matAnalista(nContador) As String
                ReDim Preserve matVTot1(nContador) As Long
                ReDim Preserve matVTot2(nContador) As Currency
                ReDim Preserve matVSal1(nContador) As Long
                ReDim Preserve matVSal2(nContador) As Currency
                ReDim Preserve matVMetaP(nContador) As Currency
                ReDim Preserve matVDesem1(nContador) As Long
                ReDim Preserve matVDesem2(nContador) As Currency

                matAnalista(nContador) = lrDataRep!ANALISTA
                sTempoAnalista = lrDataRep!ANALISTA
            End If
            
            With lrDataRep
                Select Case !Lugar
                Case 1
                    matVTot1(nContador) = !cantidad
                    matVTot2(nContador) = Format(!Total, "#0.00")
                Case 2
                    matVSal1(nContador) = !cantidad
                    matVSal2(nContador) = Format(!Total, "#0.00")
                Case 3, 5
                    matVDesem1(nContador) = !cantidad
                    matVDesem2(nContador) = Format(!Total, "#0.00")
                Case 30
                    matVMetaP(nContador) = Format(!Total, "#0.00")
                End Select
            End With
            
            RaiseEvent Progress(lrDataRep.Bookmark, lrDataRep.RecordCount)
            lrDataRep.MoveNext
            
        Loop
        lrDataRep.Close
        Set lrDataRep = Nothing
        
        RaiseEvent CloseProgress
                
        'Calculo los totales
        For i = 0 To nContador
 
            'Vigentes
            '========
            'Total de Numero para Totales
            matTot1(0) = matTot1(0) + matVTot1(i)
            'Total de Totales para Totales
            matTot1(1) = matTot1(1) + matVTot2(i)
            'Total de Numero para Saldos
            matTot1(2) = matTot1(2) + matVSal1(i)
            'Total de Totales para Saldos
            matTot1(3) = matTot1(3) + matVSal2(i)
            'Variacion a la fecha para vigente de numero
            matTot1(4) = matTot1(4) + matVTot1(i) - matVSal1(i)
            'Variacion a la fecha para vigente de monto
            matTot1(5) = matTot1(5) + matVTot2(i) - matVSal2(i)
            'Meta Propuesta
            matTot1(6) = matTot1(6) + matVMetaP(i)
            '% de Alcance
            If matVMetaP(i) = 0 Then
            Else
                matTot1(7) = matTot1(7) + Format((matVTot2(i) - matVSal2(i)) / matVMetaP(i) * 100, "#0.0000")
            End If
            'Desembolso del mes
           'Cantidad de Desembolsos
            matTot1(8) = matTot1(8) + matVDesem1(i)
            'Total de Desembolsos
            matTot1(9) = matTot1(9) + matVDesem2(i)
        Next
                
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CUADRO METAS ALCANZADAS X ANALISTA", lsCondiciones, lnPage, 134, cSubTit & Chr(10) & "TCF. = " & Format(pnTipoCambio, "0.000") & " // TCF. (" & dFechaPasada & ") = " & Format(nTipoCambioPasado, "0.000"), 108603)
        lsCadImp = lsCadImp & Chr(10)
        
        lsCadImp = lsCadImp & lsNegritaOn & "C A R T E R A  V I G E N T E" & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(134, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "        |   CARTERA VIGENTE    | SALDOS AL " & dFechaPasada & " | VARIACION A LA FECHA |META PROPUESTA |    %  ALCANCE    | DESEMBOLSO DEL MES   " & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA| NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |   M O N T O   |               |                  | NUM. |   M O N T O   " & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "        |        (A)           |        (B)           |   (C) = (A) - (B)    |     (D)       |     (C) / (D)    |                      " & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(134, "-") & Chr(10)
        
        lnIndice = 14:  lnLineas = 14
             
        For i = 0 To nContador
            lnIndice = lnIndice + 1
            lnLineas = lnLineas + 1
 
            If matVMetaP(i) = 0 Then
                ntempo1 = 0
            Else
                ntempo1 = Format((matVTot2(i) - matVSal2(i)) / matVMetaP(i) * 100, "#0.0000")
            End If
 
            cUser = matAnalista(i)
            lsCadImp = lsCadImp & Space(1) & cUser & Space(2) & _
                    " | " & ImpreFormat(matVTot1(i), 4, 0) & " | " & ImpreFormat(matVTot2(i), 10, 2) & _
                    " | " & ImpreFormat(matVSal1(i), 4, 0) & " | " & ImpreFormat(matVSal2(i), 10, 2) & _
                    " | " & ImpreFormat(matVTot1(i) - matVSal1(i), 4, 0) & " | " & ImpreFormat(matVTot2(i) - matVSal2(i), 10, 2) & _
                    " | " & ImpreFormat(matVMetaP(i), 10, 2) & " | " & IIf(ntempo1 = 0, " NO  REGISTRADA  ", ImpreFormat(Format(ntempo1, "#0.0000"), 10, 4) & " % ") & _
                    " | " & ImpreFormat(matVDesem1(i), 4, 0) & " | " & ImpreFormat(matVDesem2(i), 10, 2) & _
                    Chr(10)
          
            If lnIndice Mod 300 = 0 Then
                lsCadBuffer = lsCadBuffer & lsCadImp
                lsCadImp = ""
            End If
        
            If lnLineas >= 55 Then
                lnPage = lnPage + 1
                lsCadImp = lsCadImp & Chr(12)
                lsCadImp = lsCadImp & nRepoCabecera("CUADRO METAS ALCANZADAS X ANALISTA", lsCondiciones, lnPage, 134, cSubTit & Chr(10) & "TCF. = " & Format(pnTipoCambio, "0.000") & " // TCF. (" & dFechaPasada & ") = " & Format(nTipoCambioPasado, "0.000"), 108603)
                lnLineas = 7
                lnIndice = lnIndice + 7
            End If
        Next
        
        'Imprimo los totales de vigentes
        lsCadImp = lsCadImp & String(134, "=") & Chr(10)
        
        lsCadImp = lsCadImp & lsNegritaOn & Space(7) & " | " & ImpreFormat(matTot1(0), 4, 0) & " | " & ImpreFormat(matTot1(1), 10, 2) & _
                " | " & ImpreFormat(matTot1(2), 4, 0) & " | " & ImpreFormat(matTot1(3), 10, 2) & _
                " | " & ImpreFormat(matTot1(4), 4, 0) & " | " & ImpreFormat(matTot1(5), 10, 2) & _
                " | " & ImpreFormat(matTot1(6), 10, 2) & " | " & ImpreFormat(matTot1(7), 10, 2) & " % " & _
                " | " & ImpreFormat(matTot1(8), 4, 0) & " | " & ImpreFormat(matTot1(9), 10, 2) & _
                lsNegritaOff & Chr(10)
        
        lsCadImp = lsCadImp & String(134, "_") & Chr(10) & Chr(10) & Chr(10)
         
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
 
    End If
     
    nRepo108603_CuadroMetasAlcanzadasxAnalista = lsCadBuffer
End Function
  
Public Function nRepo108604_CarteraAltoRiesgoxAnalista(ByVal cSubTit As String, ByVal pnTipoCambio As Currency, ByVal pdFecFin As String, ByVal psMoneda As String, ByVal psTipoCredito As String, ByVal psAgencias As String, ByVal psAnalistas) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nitem As Integer
Dim loImprimeCab As NColPImpre
        
Dim lsCondiciones As String
Dim sCriterio As String
 
Dim sTempoAnalista As String
 
Dim lsNegritaOn As String
Dim lsNegritaOff As String

Dim matAnalista() As String
Dim matCarRefinanciada() As Currency
Dim matCarVencida() As Currency
Dim matCarJudicial() As Currency
Dim matColocVigentes() As Currency

Dim sTempo1 As Currency
Dim sTempo2 As Currency

Dim cantidad As Integer
Dim i As Integer

'Totales Generales
Dim sTotalA As Currency
Dim sTotalB As Currency
Dim sTotalC As Currency
Dim sTotalD As Currency
Dim sTotalE As Currency
Dim sTotalF As Currency
 
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)
 
    sCriterio = ""
                          
    lsCondiciones = "Al " & pdFecFin
   
    'Se sacan dos selects, al primero le pongo cTipo 'N' y el Segundo con cTipo 'V'
    
    'El primero saca de producto agrupado para todos los estados Vigente Normal, Vencido y Moroso
    '                                                            Refinanciado Normal, Vencido y Moroso y Judicial
    'Con estos se llenaran las columnas Refinanciadas (Los 3 refinanciados) Judicial (Judicial)
    '                                   y la Columna Colocaciones Vigentes con los 6 vigentes + Judicial
    
    'El segundo saca los Vencidos (Solo se utilizara para la columna vencidos)
    'Se consideran: Los Vigentes Vencidos(2021) => Se sacara de Producto
    '               Para los Vigentes Morosos(2022) => Para los productos 3 y 4 Se sacara las cuotas Pendientes con fecha vencida
    '               No se considera los productos  1 y 2 porque no existen con esas condiciones (Se obvia)
    '               * De considerarse los productos 1 y 2 se sacarian de la tabla Producto
  
    lsSQL = "SELECT 'N' as cTipo, Producto.nPrdEstado, "
    lsSQL = lsSQL & " SUM(case when SUBSTRING(ProductoPersona.cCtaCod, 9, 1)='" & gMonedaNacional & "' then Producto.nSaldo "
    lsSQL = lsSQL & " when SUBSTRING(ProductoPersona.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then Producto.nSaldo * " & pnTipoCambio & " End) AS Total, RRHH.cUser "
    lsSQL = lsSQL & " From ProductoPersona "
    lsSQL = lsSQL & " INNER JOIN Colocaciones ON ProductoPersona.cCtaCod = Colocaciones.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Persona ON ProductoPersona.cPersCod = Persona.cPersCod "
    lsSQL = lsSQL & " INNER JOIN  RRHH ON Persona.cPersCod = RRHH.cPersCod "
    lsSQL = lsSQL & " Where (ProductoPersona.nPrdPersRelac =" & gColRelPersAnalista & ") "
    lsSQL = lsSQL & " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRecVigJud & "')) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(ProductoPersona.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (Colocaciones.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    lsSQL = lsSQL & " AND (Colocaciones.cAgeCodAct IN (" & psAgencias & ")) "
    lsSQL = lsSQL & " AND (ProductoPersona.cPersCod "
    lsSQL = lsSQL & " IN (" & psAnalistas & ")) "
    lsSQL = lsSQL & " GROUP BY Producto.nPrdEstado, RRHH.cUser "
 
    lsSQL = lsSQL & " Union All "

    lsSQL = lsSQL & " Select 'V' as cTipo, 9999 as nPrdEstado, SUM(T.nMonto) as Total, T.cUser "
    lsSQL = lsSQL & " From "
    lsSQL = lsSQL & " ( "
    '-- Solo los vencidos
    lsSQL = lsSQL & " Select RH.cUser, "
    
    lsSQL = lsSQL & " SUM(case when SUBSTRING(P.cCtaCod, 9, 1)='" & gMonedaNacional & "' then P.nSaldo "
    lsSQL = lsSQL & " when SUBSTRING(P.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then P.nSaldo * " & pnTipoCambio & " End)  as nMonto "
 
    lsSQL = lsSQL & " From producto P "
    lsSQL = lsSQL & " INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN ProductoPersona PP ON P.cCtaCod=PP.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN RRHH RH ON PP.cPersCod = RH.cPersCod "
    lsSQL = lsSQL & " Where PP.nPrdPersRelac=" & gColRelPersAnalista & " "
    lsSQL = lsSQL & " AND PP.cPersCod IN (" & psAnalistas & ") "
    lsSQL = lsSQL & " AND p.nPrdEstado  in ('" & gColocEstVigVenc & "') "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod,9,1) IN (" & psMoneda & ") "
    End If
    lsSQL = lsSQL & " AND C.cAgeCodAct IN (" & psAgencias & ")"
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psTipoCredito & ")) "
    End If
    lsSQL = lsSQL & " Group By RH.cUser "

    lsSQL = lsSQL & " Union All "

    '--Solo los que no son vencidos pero son 1 y 2
    'No se considera porque no existiran
    
    '-- Final para morosos  tipo 3 y 4
    lsSQL = lsSQL & " select RH.cUser, "
         
    lsSQL = lsSQL & " SUM(case when SUBSTRING(p.cCtaCod, 9, 1)='" & gMonedaNacional & "' then (CCD.nMonto - CCD.nMontoPagado) "
    lsSQL = lsSQL & " when SUBSTRING(P.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then ((CCD.nMonto * " & pnTipoCambio & ") - (CCD.nMontoPagado * " & pnTipoCambio & ")) End)  as nMonto "
    
    lsSQL = lsSQL & " FROM Producto P "
    lsSQL = lsSQL & " INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN ProductoPersona PP ON P.cCtaCod=PP.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN RRHH RH ON PP.cPersCod = RH.cPersCod "
    lsSQL = lsSQL & " INNER JOIN ColocacCred CCR ON P.cCtaCod=CCR.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Coloccalendario CC ON CC.cCtaCod=CCR.cCtaCod AND CC.nNroCalen = CCR.nNroCalen "
    lsSQL = lsSQL & " INNER JOIN ColocCalendDet CCD ON CC.cCtaCod=CCD.cCtaCod "
    lsSQL = lsSQL & " AND CC.nNroCalen = CCD.nNroCalen "
    lsSQL = lsSQL & " AND CC.nColocCalendApl=CCD.nCOlocCalendApl "
    lsSQL = lsSQL & " AND CC.nCuota = CCD.nCuota "
    lsSQL = lsSQL & " Where "
    lsSQL = lsSQL & " PP.nPrdPersRelac=" & gColRelPersAnalista & " "
    lsSQL = lsSQL & " AND PP.cPersCod IN (" & psAnalistas & ") "
    lsSQL = lsSQL & " and CC.ncoloccalendapl=" & gColocCalendAplCuota & " " ' 1 Cuota
    lsSQL = lsSQL & " and datediff(dd, CC.dvenc, '" & Format(pdFecFin, "mm/dd/yyyy") & "') between 30 and 91 "
    lsSQL = lsSQL & " and CC.nColocCalendEstado=" & gColocCalendEstadoPendiente & " " '-- Vigente 0
    lsSQL = lsSQL & " AND CCD.nPrdConceptoCod=" & gColocConceptoCodCapital & " " '-- 1000 Capital
    lsSQL = lsSQL & " AND P.nPrdEstado in('" & gColocEstVigMor & "') "
    lsSQL = lsSQL & " and substring(C.cTpoCredCod,1,1) IN ('6','7','8') "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod,9,1) IN (" & psMoneda & ") "
    End If
    lsSQL = lsSQL & " and C.cAgeCodAct IN (" & psAgencias & ") "
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND C.cTpoCredCod IN (" & psTipoCredito & ") "
    End If
    lsSQL = lsSQL & " Group By RH.cUser "

    lsSQL = lsSQL & " Union All "
    lsSQL = lsSQL & " Select RH.cUser, "
    lsSQL = lsSQL & " SUM(case when SUBSTRING(p.cCtaCod, 9, 1)='" & gMonedaNacional & "' then (CCD.nMonto - CCD.nMontoPagado) "
    lsSQL = lsSQL & " when SUBSTRING(P.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then ((CCD.nMonto * " & pnTipoCambio & ") - (CCD.nMontoPagado * " & pnTipoCambio & ")) End)  as nMonto "

    lsSQL = lsSQL & " FROM Producto P "
    lsSQL = lsSQL & " INNER JOIN Colocaciones C ON C.cCtaCod = P.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN ProductoPersona PP ON P.cCtaCod=PP.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN RRHH RH ON PP.cPersCod = RH.cPersCod "
    lsSQL = lsSQL & " INNER JOIN ColocacCred CCR ON P.cCtaCod=CCR.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Coloccalendario CC ON CC.cCtaCod=CCR.cCtaCod AND CC.nNroCalen = CCR.nNroCalPar "
    lsSQL = lsSQL & " INNER JOIN ColocCalendDet CCD ON CC.cCtaCod=CCD.cCtaCod "
    lsSQL = lsSQL & " AND CC.nNroCalen = CCD.nNroCalen "
    lsSQL = lsSQL & " AND CC.nColocCalendApl=CCD.nCOlocCalendApl "
    lsSQL = lsSQL & " AND CC.nCuota = CCD.nCuota "
    lsSQL = lsSQL & " Where "
    lsSQL = lsSQL & " PP.nPrdPersRelac=" & gColRelPersAnalista & " "
    lsSQL = lsSQL & " AND PP.cPersCod IN (" & psAnalistas & ") "
    lsSQL = lsSQL & " and CC.ncoloccalendapl=" & gColocCalendAplCuota & " " '- -1 Cuota
    lsSQL = lsSQL & " and datediff(dd, CC.dvenc, '" & Format(pdFecFin, "mm/dd/yyyy") & "') Between 30 AND 91"
    lsSQL = lsSQL & " and CC.nColocCalendEstado=" & gColocCalendEstadoPendiente & " " '-- Vigente 0
    lsSQL = lsSQL & " AND CCD.nPrdConceptoCod=" & gColocConceptoCodCapital & " " '-- 1000 Capital
    lsSQL = lsSQL & " AND P.nPrdEstado in('" & gColocEstVigMor & "') "
    lsSQL = lsSQL & " and substring(C.cTpoCredCod,1,1) IN ('6','7','8') "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod,9,1) IN (" & psMoneda & ") "
    End If
    lsSQL = lsSQL & " and C.cAgeCodAct IN (" & psAgencias & ") "
    If Len(Trim(psTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND C.cTpoCredCod IN (" & psTipoCredito & ") "
    End If
    lsSQL = lsSQL & " Group By RH.cUser "

    lsSQL = lsSQL & " ) T "
    lsSQL = lsSQL & " Group By T.cUser "
    lsSQL = lsSQL & " ORDER BY cUser "



 Dim oConn As COMConecta.DCOMConecta
    Dim sqlConn As String
    Dim RRep As ADODB.Recordset
    Set oConn = New COMConecta.DCOMConecta
    oConn.AbreConexion
    
'ARCV 12-06-2007
'    Set RRep = oConn.CargaRecordSet("Select * from constsistema where  nconssiscod = 74")
'    sqlConn = RRep!nConssisValor
'    oConn.CierraConexion
'    Call oConn.AbreConexion(sqlConn)

    Set loDataRep = New dColPFunciones
        'Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
        Set lrDataRep = oConn.CargaRecordSet(lsSQL)
    Set loDataRep = Nothing

     oConn.CierraConexion
       Set oConn = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        'MsgBox " No Existen Datos para el reporte", vbInformation, " Aviso"
        Exit Function
    Else
        cantidad = 0
        sTempoAnalista = IIf(IsNull(lrDataRep!cUser), "", lrDataRep!cUser)
        
        ReDim Preserve matAnalista(cantidad) As String
        ReDim Preserve matCarRefinanciada(cantidad) As Currency
        ReDim Preserve matCarVencida(cantidad) As Currency
        ReDim Preserve matCarJudicial(cantidad) As Currency
        ReDim Preserve matColocVigentes(cantidad) As Currency
        
        RaiseEvent ShowProgress
        
        matAnalista(0) = IIf(IsNull(lrDataRep!cUser), "", lrDataRep!cUser)
         
        Do While Not lrDataRep.EOF
            If sTempoAnalista <> lrDataRep!cUser Then
                cantidad = cantidad + 1
                ReDim Preserve matAnalista(cantidad) As String
                ReDim Preserve matCarRefinanciada(cantidad) As Currency
                ReDim Preserve matCarVencida(cantidad) As Currency
                ReDim Preserve matCarJudicial(cantidad) As Currency
                ReDim Preserve matColocVigentes(cantidad) As Currency
                
                matAnalista(cantidad) = lrDataRep!cUser
                sTempoAnalista = lrDataRep!cUser
            End If
            
            If (lrDataRep!cTipo = "N") And (lrDataRep!nPrdEstado = gColocEstRefNorm Or lrDataRep!nPrdEstado = gColocEstRefVenc Or lrDataRep!nPrdEstado = gColocEstRefMor) Then
                matCarRefinanciada(cantidad) = matCarRefinanciada(cantidad) + Format(lrDataRep!Total, "#0.00")
            Else
                If (lrDataRep!cTipo = "N") And (lrDataRep!nPrdEstado = gColocEstRecVigJud) Then
                    matCarJudicial(cantidad) = matCarJudicial(cantidad) + Format(lrDataRep!Total, "#0.00")
                Else
                    If (lrDataRep!cTipo = "V") Then
                        matCarVencida(cantidad) = matCarVencida(cantidad) + Format(lrDataRep!Total, "#0.00")
                    End If
                End If
            End If
            
            'En vigentes incluyo judicial
            If lrDataRep!cTipo = "N" Then
                matColocVigentes(cantidad) = matColocVigentes(cantidad) + Format(lrDataRep!Total, "#0.00")
            End If
            
            RaiseEvent Progress(lrDataRep.Bookmark, lrDataRep.RecordCount)
            lrDataRep.MoveNext
        Loop
        lrDataRep.Close
        Set lrDataRep = Nothing
        
        RaiseEvent CloseProgress
        
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CARTERA DE ALTO RIESGO X ANALISTA", lsCondiciones, lnPage, 109, cSubTit & Chr(10) & "TCF." & ImpreFormat(pnTipoCambio, 10, 3), 108604)

        lnIndice = 0:  lnLineas = 7
          
        For i = 0 To cantidad
             
            lnIndice = lnIndice + 1
            lnLineas = lnLineas + 1
            nitem = nitem + 1
               
            sTempo1 = matCarRefinanciada(i) + matCarVencida(i) + matCarJudicial(i)
            
            If (matColocVigentes(i) + matCarJudicial(i)) = 0 Or sTempo1 = 1 Then
                sTempo2 = 0
            Else
                sTempo2 = Round(sTempo1 / (matColocVigentes(i)), 2)
            End If
               
            sTotalA = sTotalA + matCarRefinanciada(i)
            sTotalB = sTotalB + matCarVencida(i)
            sTotalC = sTotalC + matCarJudicial(i)
            sTotalD = sTotalD + sTempo1
            sTotalE = sTotalE + matColocVigentes(i)
            sTotalF = sTotalF + sTempo2
              
            lsCadImp = lsCadImp & Space(1) & ImpreFormat(nitem, 4, 0) & " | " & matAnalista(i) & " | " & _
                       ImpreFormat(matCarRefinanciada(i), 10, 2) & " | " & ImpreFormat(matCarVencida(i), 10, 2) & _
                       " | " & ImpreFormat(matCarJudicial(i), 10, 2) & " | " & ImpreFormat(sTempo1, 10, 2) & _
                       " | " & ImpreFormat(matColocVigentes(i), 10, 2) & " | " & ImpreFormat(sTempo2, 10, 2) & _
                        Chr(10)
              
            'lnLineas = lnLineas + 1
            If lnIndice Mod 300 = 0 Then
                lsCadBuffer = lsCadBuffer & lsCadImp
                lsCadImp = ""
            End If
            
            If lnLineas >= 55 Then
                lnPage = lnPage + 1
                lsCadImp = lsCadImp & Chr(12)
                lsCadImp = lsCadImp & nRepoCabecera("CARTERA DE ALTO RIESGO X ANALISTA", lsCondiciones, lnPage, 109, cSubTit & Chr(10) & "TCF." & ImpreFormat(pnTipoCambio, 10, 3), 108604)
                lnLineas = 7
                lnIndice = lnIndice + 7
            End If
              
        Next
          
        lsCadImp = lsCadImp & String(109, "=") & Chr(10)
        sTotalF = Round(sTotalD / (sTotalE + sTotalC), 2)
        lsCadImp = lsCadImp & lsNegritaOn & Space(15) & ImpreFormat(sTotalA, 10, 2) & _
                    " | " & ImpreFormat(sTotalB, 10, 2) & " | " & ImpreFormat(sTotalC, 10, 2) & _
                    " | " & ImpreFormat(sTotalD, 10, 2) & " | " & ImpreFormat(sTotalE, 10, 2) & _
                    " | " & ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & Chr(10)
                    
        lsCadImp = lsCadImp & String(109, "_") & Chr(10)
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    
    End If
    nRepo108604_CarteraAltoRiesgoxAnalista = lsCadBuffer
End Function

Public Function nRepo108605_ConsolidadoColocxFteFinanciamiento(ByVal cSubTit As String, ByVal pnTipoCambio As Currency, ByVal pdFecFin As String, ByVal psMoneda As String, ByVal psProductos As String, ByVal psAgencias As String) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nitem As Integer
Dim loImprimeCab As NColPImpre
        
Dim lsCondiciones As String
  
Dim sTempoAgencia As String
Dim sTempoFuente As String
 
Dim lsNegritaOn As String
Dim lsNegritaOff As String

Dim cMatFuente() As String
Dim matFuente() As String
Dim cMatAgencia() As String
Dim matAgencia() As String
Dim matCarVigMonN1() As Long
Dim matCarVigMonN2() As Currency
Dim matCarVigMonE1() As Long
Dim matCarVigMonE2() As Currency
Dim matCarVigTot1() As Long
Dim matCarVigTot2() As Currency
Dim matCarVenMonN1() As Long
Dim matCarVenMonN2() As Currency
Dim matCarVenMonE1() As Long
Dim matCarVenMonE2() As Currency
Dim matCarVenTot1() As Long
Dim matCarVenTot2() As Currency

Dim sSub(11) As Currency
Dim sGen(11) As Currency
Dim nContador As Integer

Dim cAgencia As String * 20
 
Dim i As Integer
Dim J As Integer
  
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)
 
                           
    lsCondiciones = "Al " & pdFecFin
   
    'Se sacan dos selects, al primero le pongo cTipo 'N' y el Segundo con cTipo 'V'
    
    'El primero saca de producto agrupado para todos los estados Vigente Normal, Vencido y Moroso
    '                                                            Refinanciado Normal, Vencido y Moroso Sin Judicial
    'Con estos se llenara la columna de Cartera Vigente con los 6 vigentes sin Judicial
    
    'El segundo (V) saca la Cartera Vencida (Solo se utilizara para la columna vencidos)
    'Se consideran: Los Vigentes Vencidos(2021) => Se sacara de Producto
    '               Para los Vigentes Morosos(2022) => Para los productos 3 y 4 Se sacara las cuotas Pendientes con fecha vencida
    '               No se considera los productos  1 y 2 porque no existen con esas condiciones (Se obvia)
   
    ' La columna de cartera vencido sera la suma de los 3 refinanciados + los vig. vencidos (2021) + los cTipo 'V'
    
    lsSQL = " SELECT 'N' as cTipo, SUBSTRING(ProductoPersona.cCtaCod, 9, 1) AS cMoneda, "
    lsSQL = lsSQL & " SUBSTRING(ProductoPersona.cCtaCod, 4, 2) AS cCodAgencia, "
    lsSQL = lsSQL & " Agencias.cAgeDescripcion AS cDesAgencia, SUBSTRING(Colocaciones.cLineaCred, 1, 2) AS cCodFuente, "
    lsSQL = lsSQL & "  ColocLineaCredito.cDescripcion AS cDesFuente, Producto.nPrdEstado, COUNT(Producto.cCtaCod) as Cantidad, "
    lsSQL = lsSQL & " SUM(Producto.nSaldo) As Total "
    lsSQL = lsSQL & " From ProductoPersona "
    lsSQL = lsSQL & " INNER JOIN Producto ON ProductoPersona.cCtaCod = Producto.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN  Colocaciones ON Producto.cCtaCod = Colocaciones.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN ColocLineaCredito ON SUBSTRING(Colocaciones.cLineaCred, 1, 2) = ColocLineaCredito.cLineaCred "
    lsSQL = lsSQL & " INNER JOIN  Agencias ON SUBSTRING(ProductoPersona.cCtaCod, 4, 2) = Agencias.cAgeCod "
    lsSQL = lsSQL & " Where (ProductoPersona.nPrdPersRelac =" & gColRelPersTitular & ") "
    lsSQL = lsSQL & " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(ProductoPersona.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(ProductoPersona.cCtaCod, 6, 3) IN (" & psProductos & ")) "
    End If

    lsSQL = lsSQL & " AND (SUBSTRING(ProductoPersona.cCtaCod, 4, 2) IN (" & psAgencias & ")) "
    lsSQL = lsSQL & " GROUP BY SUBSTRING(ProductoPersona.cCtaCod, 9, 1), SUBSTRING(ProductoPersona.cCtaCod, 4, 2), "
    lsSQL = lsSQL & " SUBSTRING(Colocaciones.cLineaCred, 1, 2),  ColocLineaCredito.cDescripcion , Producto.nPrdEstado, "
    lsSQL = lsSQL & " Agencias.cAgeDescripcion "

    lsSQL = lsSQL & " Union All "
 
    lsSQL = lsSQL & " Select 'V' as cTipo, T.cMoneda, "
    lsSQL = lsSQL & " T.cCodAgencia, T.cDesAgencia, T.cCodFuente, T.cDesFuente, 9999 as nPrdEstado, SUM(T.Cantidad) as Cantidad, SUM(T.Total) as Total "
    lsSQL = lsSQL & " From  ( "
    lsSQL = lsSQL & " Select SUBSTRING(PP.cCtaCod, 9, 1) AS cMoneda, "
    lsSQL = lsSQL & " COUNT(P.cCtaCod) as Cantidad, SUM(P.nSaldo)  as Total, "
    lsSQL = lsSQL & " SUBSTRING(PP.cCtaCod, 4, 2) AS cCodAgencia, "
    lsSQL = lsSQL & " A.cAgeDescripcion AS cDesAgencia, SUBSTRING(C.cLineaCred, 1, 2) AS cCodFuente, "
    lsSQL = lsSQL & " CLC.cDescripcion AS cDesFuente, P.nPrdEstado "
    lsSQL = lsSQL & " From producto P "
    lsSQL = lsSQL & " INNER JOIN ProductoPersona PP ON P.cCtaCod=PP.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN  Colocaciones C ON P.cCtaCod = C.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN ColocLineaCredito CLC ON SUBSTRING(C.cLineaCred, 1, 2) = CLC.cLineaCred "
    lsSQL = lsSQL & " INNER JOIN  Agencias A ON SUBSTRING(PP.cCtaCod, 4, 2) = A.cAgeCod "
    lsSQL = lsSQL & " Where PP.nPrdPersRelac =" & gColRelPersTitular & " "
    lsSQL = lsSQL & " AND p.nPrdEstado  in ('" & gColocEstVigVenc & "','" & gColocEstRefVenc & "') "
    
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ") "
    End If
    
    lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 4, 2) IN (" & psAgencias & ") "
    
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 6, 3) IN (" & psProductos & ") "
    End If

    lsSQL = lsSQL & " Group By SUBSTRING(PP.cCtaCod, 9, 1), SUBSTRING(PP.cCtaCod, 4, 2), "
    lsSQL = lsSQL & " SUBSTRING(C.cLineaCred, 1, 2),  CLC.cDescripcion , P.nPrdEstado, A.cAgeDescripcion "
    lsSQL = lsSQL & " Union All "
    lsSQL = lsSQL & " select SUBSTRING(PP.cCtaCod, 9, 1) AS cMoneda, COUNT(DISTINCT P.cCtaCod) as nCantidad, "
    lsSQL = lsSQL & " SUM (CCD.nMonto - CCD.nMontoPagado) as Total, SUBSTRING(PP.cCtaCod, 4, 2) AS cCodAgencia, "
    lsSQL = lsSQL & " A.cAgeDescripcion AS cDesAgencia, SUBSTRING(C.cLineaCred, 1, 2) AS cCodFuente, "
    lsSQL = lsSQL & " CLC.cDescripcion AS cDesFuente, P.nPrdEstado "
    lsSQL = lsSQL & " FROM Producto P "
    lsSQL = lsSQL & " INNER JOIN ProductoPersona PP ON P.cCtaCod=PP.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN ColocacCred CCR ON P.cCtaCod=CCR.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Coloccalendario CC ON CC.cCtaCod=CCR.cCtaCod AND CC.nNroCalen = CCR.nNroCalen "
    lsSQL = lsSQL & " INNER JOIN ColocCalendDet CCD ON CC.cCtaCod=CCD.cCtaCod  AND CC.nNroCalen = CCD.nNroCalen "
    lsSQL = lsSQL & " AND CC.nColocCalendApl=CCD.nCOlocCalendApl AND CC.nCuota = CCD.nCuota "
    lsSQL = lsSQL & " INNER JOIN  Colocaciones C ON P.cCtaCod = C.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN ColocLineaCredito CLC ON SUBSTRING(C.cLineaCred, 1, 2) = CLC.cLineaCred "
    lsSQL = lsSQL & " INNER JOIN  Agencias A ON SUBSTRING(PP.cCtaCod, 4, 2) = A.cAgeCod "
    lsSQL = lsSQL & " Where PP.nPrdPersRelac = " & gColRelPersTitular & " "
    lsSQL = lsSQL & " and CC.ncoloccalendapl=" & gColocCalendAplCuota & " " '- -1 Cuota
'    lsSQL = lsSQL & " and convert(varchar(8), CC.dvenc, 112)<='" & Format(pdFecFin, "YYYYMMdd") & "' "
    lsSQL = lsSQL & " And datediff(dd, CC.dvenc, '" & Format(pdFecFin, "mm/dd/yyyy") & "') Between 30 AND 91 " '-- Fecha de Vencimiento <='fecha de reporte
    lsSQL = lsSQL & " and CC.nColocCalendEstado=" & gColocCalendEstadoPendiente & " " '-- Vigente 0
    lsSQL = lsSQL & " AND CCD.nPrdConceptoCod=" & gColocConceptoCodCapital & " " '-- 1000 Capital
    lsSQL = lsSQL & " AND P.nPrdEstado in('" & gColocEstVigMor & "','" & gColocEstRefMor & "') "
    lsSQL = lsSQL & " and substring(P.cCtaCod,6,1) IN ('3','4') "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ") "
    End If
    lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 4, 2) IN (" & psAgencias & ") "
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 6, 3) IN (" & psProductos & ") "
    End If
    lsSQL = lsSQL & " Group By SUBSTRING(PP.cCtaCod, 9, 1), SUBSTRING(PP.cCtaCod, 4, 2), "
    lsSQL = lsSQL & " SUBSTRING(C.cLineaCred, 1, 2),  CLC.cDescripcion , P.nPrdEstado, A.cAgeDescripcion "
    lsSQL = lsSQL & " Union All "
    lsSQL = lsSQL & " select SUBSTRING(PP.cCtaCod, 9, 1) AS cMoneda, COUNT(DISTINCT P.cCtaCod) as nCantidad, "
    lsSQL = lsSQL & " SUM (CCD.nMonto - CCD.nMontoPagado) as Total, SUBSTRING(PP.cCtaCod, 4, 2) AS cCodAgencia, "
    lsSQL = lsSQL & " A.cAgeDescripcion AS cDesAgencia, SUBSTRING(C.cLineaCred, 1, 2) AS cCodFuente, "
    lsSQL = lsSQL & " CLC.cDescripcion AS cDesFuente, P.nPrdEstado "
    lsSQL = lsSQL & " FROM Producto P "
    lsSQL = lsSQL & " INNER JOIN ProductoPersona PP ON P.cCtaCod=PP.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN RRHH RH ON PP.cPersCod = RH.cPersCod "
    lsSQL = lsSQL & " INNER JOIN ColocacCred CCR ON P.cCtaCod=CCR.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN Coloccalendario CC ON CC.cCtaCod=CCR.cCtaCod AND CC.nNroCalen = CCR.nNroCalPar "
    lsSQL = lsSQL & " INNER JOIN ColocCalendDet CCD ON CC.cCtaCod=CCD.cCtaCod  AND CC.nNroCalen = CCD.nNroCalen "
    lsSQL = lsSQL & " AND CC.nColocCalendApl=CCD.nCOlocCalendApl  AND CC.nCuota = CCD.nCuota "
    lsSQL = lsSQL & " INNER JOIN  Colocaciones C ON P.cCtaCod = C.cCtaCod "
    lsSQL = lsSQL & " INNER JOIN ColocLineaCredito CLC ON SUBSTRING(C.cLineaCred, 1, 2) = CLC.cLineaCred "
    lsSQL = lsSQL & " INNER JOIN  Agencias A ON SUBSTRING(PP.cCtaCod, 4, 2) = A.cAgeCod "
    lsSQL = lsSQL & " Where PP.nPrdPersRelac =" & gColRelPersTitular & " "
    lsSQL = lsSQL & " and CC.ncoloccalendapl=" & gColocCalendAplCuota & " " '- -1 Cuota
    lsSQL = lsSQL & " And datediff(dd, CC.dvenc, '" & Format(pdFecFin, "mm/dd/yyyy") & "') Between 30 AND 91 " '-- Fecha de Vencimiento <='fecha de reporte
    lsSQL = lsSQL & " and CC.nColocCalendEstado=" & gColocCalendEstadoPendiente & " " '-- Vigente 0
    lsSQL = lsSQL & " AND CCD.nPrdConceptoCod=" & gColocConceptoCodCapital & " " '-- 1000 Capital
    lsSQL = lsSQL & " AND P.nPrdEstado in('" & gColocEstVigMor & "','" & gColocEstRefMor & "')  and substring(P.cCtaCod,6,1) IN ('3','4') "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ") "
    End If
    lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 4, 2) IN (" & psAgencias & ") "
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND SUBSTRING(P.cCtaCod, 6, 3) IN (" & psProductos & ") "
    End If
    lsSQL = lsSQL & " Group By SUBSTRING(PP.cCtaCod, 9, 1), SUBSTRING(PP.cCtaCod, 4, 2), SUBSTRING(C.cLineaCred, 1, 2),  CLC.cDescripcion , P.nPrdEstado, "
    lsSQL = lsSQL & " A.cAgeDescripcion "
    lsSQL = lsSQL & " ) T  Group By "
    lsSQL = lsSQL & " T.cMoneda, T.cCodAgencia , T.cDesAgencia, T.cCodFuente, T.cDesFuente "
    lsSQL = lsSQL & " ORDER BY ColocLineaCredito.cDescripcion, Agencias.cAgeDescripcion "
 
 
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
        nContador = 0
        sTempoAgencia = lrDataRep!cCodAgencia
        sTempoFuente = lrDataRep!cCodFuente
        
        ReDim Preserve cMatAgencia(nContador) As String
        ReDim Preserve cMatFuente(nContador) As String
        ReDim Preserve matAgencia(nContador) As String
        ReDim Preserve matFuente(nContador) As String
        ReDim Preserve matCarVigMonN1(nContador) As Long
        ReDim Preserve matCarVigMonN2(nContador) As Currency
        ReDim Preserve matCarVigMonE1(nContador) As Long
        ReDim Preserve matCarVigMonE2(nContador) As Currency
        ReDim Preserve matCarVenMonN1(nContador) As Long
        ReDim Preserve matCarVenMonN2(nContador) As Currency
        ReDim Preserve matCarVenMonE1(nContador) As Long
        ReDim Preserve matCarVenMonE2(nContador) As Currency
        ReDim Preserve matCarVigTot1(nContador) As Long
        ReDim Preserve matCarVigTot2(nContador) As Currency
        ReDim Preserve matCarVenTot1(nContador) As Long
        ReDim Preserve matCarVenTot2(nContador) As Currency
        
        cMatAgencia(0) = lrDataRep!cCodAgencia
        cMatFuente(0) = lrDataRep!cCodFuente
        matAgencia(0) = lrDataRep!cDesAgencia
        matFuente(0) = lrDataRep!cDesFuente
        
        RaiseEvent ShowProgress
        
        Do While Not lrDataRep.EOF
            If sTempoAgencia <> lrDataRep!cCodAgencia Or sTempoFuente <> lrDataRep!cCodFuente Then
                nContador = nContador + 1
                ReDim Preserve cMatAgencia(nContador) As String
                ReDim Preserve cMatFuente(nContador) As String
                ReDim Preserve matAgencia(nContador) As String
                ReDim Preserve matFuente(nContador) As String
                ReDim Preserve matCarVigMonN1(nContador) As Long
                ReDim Preserve matCarVigMonN2(nContador) As Currency
                ReDim Preserve matCarVigMonE1(nContador) As Long
                ReDim Preserve matCarVigMonE2(nContador) As Currency
                ReDim Preserve matCarVenMonN1(nContador) As Long
                ReDim Preserve matCarVenMonN2(nContador) As Currency
                ReDim Preserve matCarVenMonE1(nContador) As Long
                ReDim Preserve matCarVenMonE2(nContador) As Currency
                ReDim Preserve matCarVigTot1(nContador) As Long
                ReDim Preserve matCarVigTot2(nContador) As Currency
                ReDim Preserve matCarVenTot1(nContador) As Long
                ReDim Preserve matCarVenTot2(nContador) As Currency
                
                cMatAgencia(nContador) = lrDataRep!cCodAgencia
                cMatFuente(nContador) = lrDataRep!cCodFuente
                matAgencia(nContador) = lrDataRep!cDesAgencia
                matFuente(nContador) = lrDataRep!cDesFuente
                
                sTempoAgencia = lrDataRep!cCodAgencia
                sTempoFuente = lrDataRep!cCodFuente
            End If
            
            If lrDataRep!cTipo = "N" Then
                If lrDataRep!nPrdEstado = gColocEstVigNorm Or lrDataRep!nPrdEstado = gColocEstVigVenc Or lrDataRep!nPrdEstado = gColocEstVigMor Or lrDataRep!nPrdEstado = gColocEstRefNorm Or lrDataRep!nPrdEstado = gColocEstRefVenc Or lrDataRep!nPrdEstado = gColocEstRefMor Then
                    'Saco todos
                    'Vigente
                    If lrDataRep!cMoneda = gMonedaNacional Then 'Moneda nacional
                        'Cantidad Moneda Nacional Vigente
                        matCarVigMonN1(nContador) = matCarVigMonN1(nContador) + lrDataRep!cantidad
                        'Total Moneda Nacional Vigente
                        matCarVigMonN2(nContador) = matCarVigMonN2(nContador) + Format(lrDataRep!Total, "#0.00")
                        matCarVigTot1(nContador) = matCarVigTot1(nContador) + lrDataRep!cantidad
                        matCarVigTot2(nContador) = matCarVigTot2(nContador) + Format(lrDataRep!Total, "#0.00")
                    ElseIf lrDataRep!cMoneda = gMonedaExtranjera Then
                        'Cantidad Moneda Extranjera Vigente
                        matCarVigMonE1(nContador) = matCarVigMonE1(nContador) + lrDataRep!cantidad
                        'Total Moneda Extranjera Vigente
                        matCarVigMonE2(nContador) = matCarVigMonE2(nContador) + Format(lrDataRep!Total, "#0.00")
                        matCarVigTot1(nContador) = matCarVigTot1(nContador) + lrDataRep!cantidad
                        matCarVigTot2(nContador) = matCarVigTot2(nContador) + Format((lrDataRep!Total * pnTipoCambio), "#0.00")
                    End If
                End If
            Else
                If lrDataRep!cTipo = "V" Then
                    'Vencido
                    If lrDataRep!cMoneda = gMonedaNacional Then 'Moneda nacional
                        'Cantidad Moneda Nacional Vencida
                        matCarVenMonN1(nContador) = matCarVenMonN1(nContador) + lrDataRep!cantidad
                        'Total Moneda Nacional Vencida
                        matCarVenMonN2(nContador) = matCarVenMonN2(nContador) + Format(lrDataRep!Total, "#0.00")
                        matCarVenTot1(nContador) = matCarVenTot1(nContador) + lrDataRep!cantidad
                        matCarVenTot2(nContador) = matCarVenTot2(nContador) + Format(lrDataRep!Total, "#0.00")
                        
                    ElseIf lrDataRep!cMoneda = gMonedaExtranjera Then
                        'Cantidad Moneda Extranjera Vencida
                        matCarVenMonE1(nContador) = matCarVenMonE1(nContador) + lrDataRep!cantidad
                        'Total Moneda Extranjera Vencida
                        matCarVenMonE2(nContador) = matCarVenMonE2(nContador) + Format(lrDataRep!Total, "#0.00")
                        matCarVenTot1(nContador) = matCarVenTot1(nContador) + lrDataRep!cantidad
                        matCarVenTot2(nContador) = matCarVenTot2(nContador) + Format((lrDataRep!Total * pnTipoCambio), "#0.00")
                    End If
                End If
            End If
            
            RaiseEvent Progress(lrDataRep.Bookmark, lrDataRep.RecordCount)
            lrDataRep.MoveNext
            
        Loop
        lrDataRep.Close
        Set lrDataRep = Nothing
  
        RaiseEvent CloseProgress
  
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CUADRO CONSOLIDADO DE COLOCACIONES X FUENTE DE FINANCIAMIENTO(Sin inc. Judicial)", lsCondiciones, lnPage, 140, cSubTit & Chr(10) & "TCF." & ImpreFormat(pnTipoCambio, 10, 3), 108605)

        'Fuente de Financiamiento
        lsCadImp = lsCadImp & lsNegritaOn & matFuente(i) & lsNegritaOff & Chr(10)

        lnIndice = 0:  lnLineas = 7
          
        sTempoFuente = ""
        sTempoAgencia = ""
        
        sTempoFuente = cMatFuente(0)
        sTempoAgencia = cMatAgencia(0)
        
          
        For i = 0 To nContador
             
            lnIndice = lnIndice + 1
            lnLineas = lnLineas + 1
            nitem = nitem + 1
               
            If sTempoFuente <> cMatFuente(i) Then
  
               lsCadImp = lsCadImp & String(140, "-") & Chr(10)
               lsCadImp = lsCadImp & Space(21) & _
                           ImpreFormat(sSub(0), 5, 0) & Space(1) & ImpreFormat(sSub(1), 10, 2) & Space(1) & _
                           ImpreFormat(sSub(2), 5, 0) & Space(1) & ImpreFormat(sSub(3), 10, 2) & Space(1) & _
                           ImpreFormat(sSub(4), 5, 0) & Space(1) & ImpreFormat(sSub(5), 10, 2) & Space(1) & _
                           ImpreFormat(sSub(6), 5, 0) & Space(1) & ImpreFormat(sSub(7), 10, 2) & Space(1) & _
                           ImpreFormat(sSub(8), 5, 0) & Space(1) & ImpreFormat(sSub(9), 10, 2) & Space(1) & _
                           ImpreFormat(sSub(10), 5, 0) & Space(1) & ImpreFormat(sSub(11), 10, 2) & Chr(10) & Chr(10)
                           
               'Fuente de Financiamiento
               lsCadImp = lsCadImp & lsNegritaOn & matFuente(i) & lsNegritaOff & Chr(10)
                
               'Vigente
               sSub(0) = matCarVigMonN1(i)
               sSub(1) = matCarVigMonN2(i)
               sSub(2) = matCarVigMonE1(i)
               sSub(3) = matCarVigMonE2(i)
               sSub(4) = matCarVigMonN1(i) + Format(matCarVigMonE1(i), "#0.00")
               sSub(5) = matCarVigMonN2(i) + Format((matCarVigMonE2(i) * pnTipoCambio), "#0.00")
               
               'Vencida
               sSub(6) = matCarVenMonN1(i)
               sSub(7) = matCarVenMonN2(i)
               sSub(8) = matCarVenMonE1(i)
               sSub(9) = matCarVenMonE2(i)
               sSub(10) = matCarVenMonN1(i) + Format(matCarVenMonE1(i), "#0.00")
               sSub(11) = matCarVenMonN2(i) + Format((matCarVenMonE2(i) * pnTipoCambio), "#0.00")
    
               nitem = 1
               sTempoFuente = cMatFuente(i)
               sTempoAgencia = cMatAgencia(i)
                
               lnLineas = lnLineas + 4
               lnIndice = lnIndice + 4
                
           Else
               'Vigente
               sSub(0) = sSub(0) + matCarVigMonN1(i)
               sSub(1) = sSub(1) + matCarVigMonN2(i)
               sSub(2) = sSub(2) + matCarVigMonE1(i)
               sSub(3) = sSub(3) + matCarVigMonE2(i)
               sSub(4) = sSub(4) + matCarVigMonN1(i) + Format(matCarVigMonE1(i), "#0.00")
               sSub(5) = sSub(5) + matCarVigMonN2(i) + Format((matCarVigMonE2(i) * pnTipoCambio), "#0.00")
               
               'Vencida
               sSub(6) = sSub(6) + matCarVenMonN1(i)
               sSub(7) = sSub(7) + matCarVenMonN2(i)
               sSub(8) = sSub(8) + matCarVenMonE1(i)
               sSub(9) = sSub(9) + matCarVenMonE2(i)
               sSub(10) = sSub(10) + matCarVenMonN1(i) + Format(matCarVenMonE1(i), "#0.00")
               sSub(11) = sSub(11) + matCarVenMonN2(i) + Format((matCarVenMonE2(i) * pnTipoCambio), "#0.00")
             
           End If
        
            'Vigente
            sGen(0) = sGen(0) + matCarVigMonN1(i)
            sGen(1) = sGen(1) + matCarVigMonN2(i)
            sGen(2) = sGen(2) + matCarVigMonE1(i)
            sGen(3) = sGen(3) + matCarVigMonE2(i)
            sGen(4) = sGen(4) + matCarVigMonN1(i) + Format(matCarVigMonE1(i), "#0.00")
            sGen(5) = sGen(5) + matCarVigMonN2(i) + Format((matCarVigMonE2(i) * pnTipoCambio), "#0.00")
            
            'Vencida
            sGen(6) = sGen(6) + matCarVenMonN1(i)
            sGen(7) = sGen(7) + matCarVenMonN2(i)
            sGen(8) = sGen(8) + matCarVenMonE1(i)
            sGen(9) = sGen(9) + matCarVenMonE2(i)
            sGen(10) = sGen(10) + matCarVenMonN1(i) + Format(matCarVenMonE1(i), "#0.00")
            sGen(11) = sGen(11) + matCarVenMonN2(i) + Format((matCarVenMonE2(i) * pnTipoCambio), "#0.00")
            
            cAgencia = matAgencia(i)
            lsCadImp = lsCadImp & Space(1) & cAgencia & _
                       ImpreFormat(matCarVigMonN1(i), 5, 0) & Space(1) & ImpreFormat(matCarVigMonN2(i), 10, 2) & Space(1) & _
                       ImpreFormat(matCarVigMonE1(i), 5, 0) & Space(1) & ImpreFormat(matCarVigMonE2(i), 10, 2) & Space(1) & _
                       ImpreFormat(matCarVigTot1(i), 5, 0) & Space(1) & ImpreFormat(matCarVigTot2(i), 10, 2) & Space(1) & _
                       ImpreFormat(matCarVenMonN1(i), 5, 0) & Space(1) & ImpreFormat(matCarVenMonN2(i), 10, 2) & Space(1) & _
                       ImpreFormat(matCarVenMonE1(i), 5, 0) & Space(1) & ImpreFormat(matCarVenMonE2(i), 10, 2) & Space(1) & _
                       ImpreFormat(matCarVenTot1(i), 5, 0) & Space(1) & ImpreFormat(matCarVenTot2(i), 10, 2) & Space(1) & _
                       Chr(10)
              
            'lnLineas = lnLineas + 1
            If lnIndice Mod 300 = 0 Then
                lsCadBuffer = lsCadBuffer & lsCadImp
                lsCadImp = ""
            End If
            
            If lnLineas >= 55 Then
                lnPage = lnPage + 1
                lsCadImp = lsCadImp & Chr(12)
                lsCadImp = lsCadImp & nRepoCabecera("CUADRO CONSOLIDADO DE COLOCACIONES X FUENTE DE FINANCIAMIENTO(Sin inc. Judicial)", lsCondiciones, lnPage, 140, cSubTit & Chr(10) & "TCF." & ImpreFormat(pnTipoCambio, 10, 3), 108605)
                lnLineas = 7
                lnIndice = lnIndice + 7
            End If
        Next
          
               lsCadImp = lsCadImp & String(140, "-") & Chr(10)
               lsCadImp = lsCadImp & Space(21) & _
                           ImpreFormat(sSub(0), 5, 0) & Space(1) & ImpreFormat(sSub(1), 10, 2) & Space(1) & _
                           ImpreFormat(sSub(2), 5, 0) & Space(1) & ImpreFormat(sSub(3), 10, 2) & Space(1) & _
                           ImpreFormat(sSub(4), 5, 0) & Space(1) & ImpreFormat(sSub(5), 10, 2) & Space(1) & _
                           ImpreFormat(sSub(6), 5, 0) & Space(1) & ImpreFormat(sSub(7), 10, 2) & Space(1) & _
                           ImpreFormat(sSub(8), 5, 0) & Space(1) & ImpreFormat(sSub(9), 10, 2) & Space(1) & _
                           ImpreFormat(sSub(10), 5, 0) & Space(1) & ImpreFormat(sSub(11), 10, 2) & Chr(10) & Chr(10)
         
               lsCadImp = lsCadImp & String(140, "=") & Chr(10)
               lsCadImp = lsCadImp & "TOTAL" & Space(16) & _
                           ImpreFormat(sGen(0), 5, 0) & Space(1) & ImpreFormat(sGen(1), 10, 2) & Space(1) & _
                           ImpreFormat(sGen(2), 5, 0) & Space(1) & ImpreFormat(sGen(3), 10, 2) & Space(1) & _
                           ImpreFormat(sGen(4), 5, 0) & Space(1) & ImpreFormat(sGen(5), 10, 2) & Space(1) & _
                           ImpreFormat(sGen(6), 5, 0) & Space(1) & ImpreFormat(sGen(7), 10, 2) & Space(1) & _
                           ImpreFormat(sGen(8), 5, 0) & Space(1) & ImpreFormat(sGen(9), 10, 2) & Space(1) & _
                           ImpreFormat(sGen(10), 5, 0) & Space(1) & ImpreFormat(sGen(11), 10, 2) & Chr(10)
               lsCadImp = lsCadImp & String(140, "_") & Chr(10)
               
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    
    End If
    nRepo108605_ConsolidadoColocxFteFinanciamiento = lsCadBuffer
End Function
 
Public Function nRepo108606_ConsolidadoColocacionesxMoraxAnalista(ByVal pnTipoFecha As Integer, ByVal cSubTit As String, ByVal pnTipoCambio As Currency, ByVal pdFecFin As String, ByVal psMoneda As String, ByVal psProductos As String, ByVal psAgencias As String, ByVal psAnalistas) As String
Dim matAnalista() As String
Dim matVMonN1() As Long
Dim matVMonN2() As Currency
Dim matVMonE1() As Long
Dim matVMonE2() As Currency
Dim matVTot1() As Long
Dim matVTot2() As Currency
Dim matVSal1() As Long
Dim matVSal2() As Currency
Dim matMMonN1() As Long
Dim matMMonN2() As Currency
Dim matMMonE1() As Long
Dim matMMonE2() As Currency
Dim matMTot1() As Long
Dim matMTot2() As Currency
Dim matMSal1() As Long
Dim matMSal2() As Currency
Dim matIngRecup()  As Currency

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim loImprimeCab As NColPImpre
        
Dim lsCondiciones As String

Dim oTipCambio As nTipoCambio
 
'Obtencion de la fecha del mes pasado
Dim dUltimoDia As DCredReporte
Dim nDiaPasado As Integer
Dim nMesPasado As Integer
Dim dFechaPasada As String
Dim nTipoCambioPasado As Currency

Dim i As Integer
Dim sTempoAnalista As String
Dim cUser As String * 4
Dim lsNegritaOn As String
Dim lsNegritaOff As String

'Los totales de la parte superior
Dim matTot1(11) As Currency
'Los totales de la parte inferior
Dim matTot2(11) As Currency

Dim nContador As Integer

lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

    lsCondiciones = "Al " & pdFecFin

    Set dUltimoDia = New DCredReporte
    nMesPasado = IIf(Val(Mid(Format(pdFecFin, "dd/MM/YYYY"), 4, 2)) = 1, 12, Val(Mid(Format(pdFecFin, "dd/MM/YYYY"), 4, 2)) - 1)
    dFechaPasada = "01/" & IIf(nMesPasado < 10, "0" & Trim(Str(nMesPasado)), Trim(Str(nMesPasado))) & "/" & IIf(nMesPasado = 1, Trim(Str(Val(Mid(Format(pdFecFin, "dd/MM/YYYY"), 7, 4)) - 1)), Trim(Str(Val(Mid(Format(pdFecFin, "dd/MM/YYYY"), 7, 4)))))
    
    nDiaPasado = dUltimoDia.RecuperaUltimoDiaMes(dFechaPasada)
    dFechaPasada = IIf(nDiaPasado < 10, "0" & Trim(Str(nDiaPasado)), Trim(Str(nDiaPasado))) & "/" & Mid(dFechaPasada, 4, 7)
    'Ya tengo la fecha del ultimo dia del mes pasado
    'La fecha actual es la que pasan como parametro
    dFechaPasada = Format(dFechaPasada, "dd/MM/YYYY")
      
    Set oTipCambio = New nTipoCambio
    'nTipoCambioPasado = Format(oTipCambio.EmiteTipoCambio(dFechaPasada, TCFijoMes), "0.00")
    nTipoCambioPasado = Format(oTipCambio.EmiteTipoCambio(DateAdd("d", 1, dFechaPasada), TCFijoMes), "0.000")
    Set oTipCambio = Nothing

'****** By BRGO BASILEA II ***************************

'10) Moneda Nacional Vigente
If pnTipoFecha = 1 Then 'Fecha Actual
    lsSQL = " SELECT 10 AS Lugar, COUNT(P.cCtaCod) AS Cantidad, " & _
    " SUM(P.nSaldo) AS Total, RRHH.cUser AS Analista " & _
    " FROM Colocaciones C " & _
    " INNER JOIN Producto P ON C.cCtaCod = P.cCtaCod " & _
    " INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod " & _
    " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod  " & _
    " INNER JOIN RRHH ON PE.cPersCod = RRHH.cPersCod " & _
    " WHERE (PP.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (PP.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (P.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
    " AND (C.cAgeCodAct IN (" & psAgencias & ")) " & _
    " AND (SUBSTRING(P.cCtaCod, 9, 1) =('" & gMonedaNacional & "')) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
ElseIf pnTipoFecha = 2 Then 'Fecha pasada(que llega como parametro)
    lsSQL = "SELECT 10 AS Lugar, COUNT(CS.cCtaCod) AS Cantidad, " & _
    " SUM(CS.nSaldoCap) AS Total, RRHH.cUser AS Analista " & _
    " FROM Colocaciones C " & _
    " INNER JOIN ProductoPersona PP ON PP.cCtaCod = C.cCtaCod " & _
    " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " & _
    " INNER JOIN RRHH ON PE.cPersCod = RRHH.cPersCod " & _
    " INNER JOIN ColocacSaldo CS ON PP.cCtaCod = CS.cCtaCod " & _
    " WHERE (PP.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (PP.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (DATEDIFF(day, CS.dFecha, '" & Format(dFechaPasada, "mm/dd/yyyy") & "') = 0) " & _
    " AND (CS.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
    " AND (C.cAgeCodAct IN (" & psAgencias & ")) " & _
    " AND (SUBSTRING(CS.cCtaCod, 9, 1) =('" & gMonedaNacional & "')) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(C.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
End If

lsSQL = lsSQL & " UNION ALL "

'11) Moneda Extranjera Vigente
If pnTipoFecha = 1 Then 'Fecha Actual
    lsSQL = lsSQL & " SELECT 11 AS Lugar, COUNT(P.cCtaCod) AS Cantidad, " & _
    " SUM(P.nSaldo) AS Total, RRHH.cUser AS Analista " & _
    " FROM Colocaciones C " & _
    " INNER JOIN Producto P ON C.cCtaCod = P.cCtaCod " & _
    " INNER JOIN ProductoPersona PP ON PP.cCtaCod = P.cCtaCod " & _
    " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " & _
    " INNER JOIN RRHH ON PE.cPersCod = RRHH.cPersCod " & _
    " WHERE (PP.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (PP.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (P.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
    " AND (C.cAgeCodAct IN (" & psAgencias & ")) " & _
    " AND (SUBSTRING(P.cCtaCod, 9, 1) =('" & gMonedaExtranjera & "')) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
ElseIf pnTipoFecha = 2 Then 'Fecha pasada(que llega como parametro)
    lsSQL = lsSQL & " SELECT 11 AS Lugar, COUNT(CS.cCtaCod) AS Cantidad, " & _
    " SUM(CS.nSaldoCap) AS Total, RRHH.cUser AS Analista " & _
    " FROM Colocaciones C " & _
    " INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod" & _
    " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " & _
    " INNER JOIN RRHH ON PE.cPersCod = RRHH.cPersCod " & _
    " INNER JOIN ColocacSaldo CS ON PP.cCtaCod = CS.cCtaCod " & _
    " WHERE (PP.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (PP.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (DATEDIFF(day, CS.dFecha, '" & Format(dFechaPasada, "mm/dd/yyyy") & "') = 0) " & _
    " AND (CS.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
    " AND (C.cAgeCodAct IN (" & psAgencias & ")) " & _
    " AND (SUBSTRING(CS.cCtaCod, 9, 1) =('" & gMonedaExtranjera & "')) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CS.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
End If

lsSQL = lsSQL & " UNION ALL "
 
'1) Obtengo el total para la cartera vigente
If pnTipoFecha = 1 Then 'Fecha Actual
    lsSQL = lsSQL & " SELECT 1 AS Lugar, COUNT(P.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(P.cCtaCod, 9, 1)='" & gMonedaNacional & "' then P.nSaldo " & _
    " when SUBSTRING(P.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then P.nSaldo * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM Colocaciones C " & _
    " INNER JOIN ProductoPersona PP ON PP.cCtaCod = C.cCtaCod " & _
    " INNER JOIN Producto P ON PP.cCtaCod = P.cCtaCod " & _
    " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " & _
    " INNER JOIN RRHH ON PE.cPersCod = RRHH.cPersCod " & _
    " WHERE (PP.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (PP.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (P.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
    " AND (C.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
ElseIf pnTipoFecha = 2 Then 'Fecha pasada(que llega como parametro)
    lsSQL = lsSQL & "SELECT 1 AS Lugar, COUNT(CS.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(CS.cCtaCod, 9, 1)='" & gMonedaNacional & "' then CS.nSaldoCap " & _
    " when SUBSTRING(CS.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then CS.nSaldoCap * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM Colocaciones C " & _
    " INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod" & _
    " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " & _
    " INNER JOIN RRHH ON PE.cPersCod = RRHH.cPersCod " & _
    " INNER JOIN ColocacSaldo CS ON PP.cCtaCod = CS.cCtaCod " & _
    " WHERE (PP.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (PP.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (DATEDIFF(day, CS.dFecha, '" & Format(dFechaPasada, "mm/dd/yyyy") & "') = 0) " & _
    " AND (CS.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
    " AND (C.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CS.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
End If

'2) Obtengo los saldos del mes pasado (Para todos no inportando el tipo de fecha)
'    para la cartera vigente

    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 2 AS Lugar, COUNT(CS.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(CS.cCtaCod, 9, 1)='" & gMonedaNacional & "' then CS.nSaldoCap " & _
    " when SUBSTRING(CS.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then CS.nSaldoCap * " & nTipoCambioPasado & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM Colocaciones C " & _
    " INNER JOIN ProductoPersona PP ON C.cCtaCod = PP.cCtaCod " & _
    " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " & _
    " INNER JOIN RRHH ON PE.cPersCod = RRHH.cPersCod  " & _
    " INNER JOIN ColocacSaldo CS ON PP.cCtaCod = CS.cCtaCod " & _
    " WHERE (PP.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (PP.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (DATEDIFF(day, CS.dFecha, '" & Format(dFechaPasada, "mm/dd/yyyy") & "') = 0) " & _
    " AND (CS.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) " & _
    " AND (C.cAgeCodAct IN (" & psAgencias & "))  "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CS.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser"

'20) Obtengo la moneda nacional para la cartera morosa

lsSQL = lsSQL & " UNION ALL "
     
If pnTipoFecha = 1 Then 'Fecha Actual
    lsSQL = lsSQL & " SELECT 20 AS Lugar, COUNT(P.cCtaCod) AS Cantidad, " & _
    " SUM(P.nSaldo) AS Total, RRHH.cUser AS Analista " & _
    " FROM Colocaciones C " & _
    " INNER JOIN Producto P ON C.cCtaCod = P.cCtaCod " & _
    " INNER JOIN ProductoPersona PP ON PP.cCtaCod = P.cCtaCod " & _
    " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " & _
    " INNER JOIN RRHH ON PE.cPersCod = RRHH.cPersCod " & _
    " WHERE (PP.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (PP.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (P.nPrdEstado IN (" & gColocEstVigMor & "," & gColocEstRefMor & "," & gColocEstVigVenc & "," & gColocEstRefVenc & ")) " & _
    " AND (C.cAgeCodAct IN (" & psAgencias & ")) " & _
    " AND (SUBSTRING(P.cCtaCod, 9, 1) IN (" & gMonedaNacional & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
    
ElseIf pnTipoFecha = 2 Then 'Fecha pasada
    lsSQL = lsSQL & " SELECT 20 AS Lugar, COUNT(CS.cCtaCod) AS Cantidad, " & _
    " SUM(CS.nSaldoCap) AS Total, RRHH.cUser AS Analista " & _
    " FROM Colocaciones C " & _
    " INNER JOIN Producto P ON C.cCtaCod=P.cCtaCod " & _
    " INNER JOIN ProductoPersona PP ON PP.cCtaCod = P.cCtaCod " & _
    " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " & _
    " INNER JOIN RRHH ON P.cPersCod = RRHH.cPersCod " & _
    " INNER JOIN ColocacSaldo CS ON PP.cCtaCod = CS.cCtaCod " & _
    " WHERE (PP.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (PP.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (DATEDIFF(day, CS.dFecha, '" & Format(dFechaPasada, "mm/dd/yyyy") & "') = 0) " & _
    " AND (P.nPrdEstado IN (" & gColocEstVigMor & "," & gColocEstRefMor & "," & gColocEstVigVenc & "," & gColocEstRefVenc & ")) " & _
    " AND (C.cAgeCodAct IN (" & psAgencias & ")) " & _
    " AND (SUBSTRING(CS.cCtaCod, 9, 1) IN (" & gMonedaNacional & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CS.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
End If

'21) Obtengo la moneda nacional para la cartera morosa

lsSQL = lsSQL & " UNION ALL "
     
If pnTipoFecha = 1 Then 'Fecha Actual
    lsSQL = lsSQL & " SELECT 21 AS Lugar, COUNT(P.cCtaCod) AS Cantidad, " & _
    " SUM(P.nSaldo) AS Total, RRHH.cUser AS Analista " & _
    " FROM Colocaciones C " & _
    " INNER JOIN Producto P ON C.cCtaCod = P.cCtaCod " & _
    " INNER JOIN ProductoPersona PP ON PP.cCtaCod = P.cCtaCod " & _
    " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " & _
    " INNER JOIN RRHH ON PE.cPersCod = RRHH.cPersCod " & _
    " WHERE (PP.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (PP.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (P.nPrdEstado IN (" & gColocEstVigMor & "," & gColocEstRefMor & "," & gColocEstVigVenc & "," & gColocEstRefVenc & ")) " & _
    " AND (C.cAgeCodAct IN (" & psAgencias & ")) " & _
    " AND (SUBSTRING(P.cCtaCod, 9, 1) IN (" & gMonedaExtranjera & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
    
ElseIf pnTipoFecha = 2 Then 'Fecha pasada
    lsSQL = lsSQL & " SELECT 21 AS Lugar, COUNT(CS.cCtaCod) AS Cantidad, " & _
    " SUM(CS.nSaldoCap) AS Total, RRHH.cUser AS Analista " & _
    " FROM Producto P" & _
    " INNER JOIN ProductoPersona PP ON PP.cCtaCod=P.cCtaCod " & _
    " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " & _
    " INNER JOIN RRHH ON PE.cPersCod = RRHH.cPersCod " & _
    " INNER JOIN ColocacSaldo CS ON PP.cCtaCod = CS.cCtaCod " & _
    " WHERE (PP.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (PP.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (DATEDIFF(day, CS.dFecha, '" & Format(dFechaPasada, "mm/dd/yyyy") & "') = 0) " & _
    " AND (P.nPrdEstado IN (" & gColocEstVigMor & "," & gColocEstRefMor & "," & gColocEstVigVenc & "," & gColocEstRefVenc & ")) " & _
    " AND (C.cAgeCodAct IN (" & psAgencias & ")) " & _
    " AND (SUBSTRING(CS.cCtaCod, 9, 1) IN (" & gMonedaExtranjera & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CS.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
End If
 
'6) Obtengo el total para la cartera morosa

    lsSQL = lsSQL & " UNION ALL "
     
If pnTipoFecha = 1 Then 'Fecha Actual
    lsSQL = lsSQL & " SELECT 6 AS Lugar, COUNT(P.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(P.cCtaCod, 9, 1)='" & gMonedaNacional & "' then P.nSaldo " & _
    " when SUBSTRING(P.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then P.nSaldo * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM Colocaciones C " & _
    " INNER JOIN Producto P ON C.cCtaCod = P.cCtaCod " & _
    " INNER JOIN ProductoPersona PP ON PP.cCtaCod = P.cCtaCod " & _
    " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " & _
    " INNER JOIN RRHH ON PE.cPersCod = RRHH.cPersCod " & _
    " WHERE (PP.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (PP.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (P.nPrdEstado IN ('" & gColocEstVigMor & "', '" & gColocEstRefMor & "', '" & gColocEstRefVenc & "', '" & gColocEstVigVenc & "')) " & _
    " AND (C.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
ElseIf pnTipoFecha = 2 Then 'Fecha pasada
    lsSQL = lsSQL & " SELECT 6 AS Lugar, COUNT(CS.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(CS.cCtaCod, 9, 1)='" & gMonedaNacional & "' then CS.nSaldoCap " & _
    " when SUBSTRING(CS.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then CS.nSaldoCap * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM Colocaciones C " & _
    " INNER JOIN Producto P ON P.cCtaCod=C.cCtaCod " & _
    " INNER JOIN ProductoPersona PP ON PP.cCtaCod = P.cCtaCod " & _
    " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " & _
    " INNER JOIN RRHH ON PE.cPersCod = RRHH.cPersCod " & _
    " INNER JOIN ColocacSaldo CS ON PP.cCtaCod = CS.cCtaCod " & _
    " WHERE (PP.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (PP.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (DATEDIFF(day, CS.dFecha, '" & Format(dFechaPasada, "mm/dd/yyyy") & "') = 0) " & _
    " AND (P.nPrdEstado IN ('" & gColocEstVigMor & "', '" & gColocEstRefMor & "', '" & gColocEstVigVenc & "', '" & gColocEstRefVenc & "')) " & _
    " AND (C.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CS.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser "
End If

'7) Obtengo los saldos del mes pasado (Para todos no inportando el tipo de fecha)
'    para la cartera morosa

    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 7 AS Lugar, COUNT(CS.cCtaCod) AS Cantidad, " & _
    " SUM(case when SUBSTRING(CS.cCtaCod, 9, 1)='" & gMonedaNacional & "' then CS.nSaldoCap " & _
    " when SUBSTRING(CS.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then CS.nSaldoCap * " & nTipoCambioPasado & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM Colocaciones C " & _
    " INNER JOIN ProductoPersona PP ON PP.cCtaCod = C.cCtaCod " & _
    " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " & _
    " INNER JOIN RRHH ON PE.cPersCod = RRHH.cPersCod " & _
    " INNER JOIN ColocacSaldo CS ON PP.cCtaCod = CS.cCtaCod " & _
    " WHERE (PP.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND (PP.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (DATEDIFF(day, CS.dFecha, '" & Format(dFechaPasada, "mm/dd/yyyy") & "') = 0) " & _
    " AND (CS.nPrdEstado IN ('" & gColocEstVigMor & "', '" & gColocEstRefMor & "', '" & gColocEstRefVenc & "', '" & gColocEstVigVenc & "')) " & _
    " AND (C.cAgeCodAct IN (" & psAgencias & "))  "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(CS.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " GROUP BY RRHH.cUser"
  
' 6) Obtengo los ingresos a recuperaciones          '--Considero la suma de COlocacEstado.Monto para ColocacEstado.nPrdEstado=2201
'Para la cartera morosa

    lsSQL = lsSQL & " UNION ALL "
    lsSQL = lsSQL & " SELECT 8 AS Lugar, 0 as Cantidad, " & _
    " SUM(case when SUBSTRING(P.cCtaCod, 9, 1)='" & gMonedaNacional & "' then CE.nMonto " & _
    " WHEN SUBSTRING(P.cCtaCod, 9, 1)='" & gMonedaExtranjera & "' then CE.nMonto * " & pnTipoCambio & " " & _
    " End) AS Total, RRHH.cUser AS Analista " & _
    " FROM Colocaciones C " & _
    " INNER JOIN Producto P ON C.cCtaCod = P.cCtaCod " & _
    " INNER JOIN ProductoPersona PP ON PP.cCtaCod = P.cCtaCod " & _
    " INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod " & _
    " INNER JOIN RRHH ON PE.cPersCod = RRHH.cPersCod " & _
    " INNER JOIN ColocRecup CR ON P.cCtaCod = CR.cCtaCod " & _
    " INNER JOIN ColocacEstado CE ON CR.cCtaCod = CE.cCtaCod " & _
    " WHERE (PP.nPrdPersRelac='" & gColRelPersAnalista & "') " & _
    " AND CE.nPrdEstado='" & gColocEstRecVigJud & "' " & _
    " AND (PP.cPersCod IN (" & psAnalistas & ")) " & _
    " AND (C.cAgeCodAct IN (" & psAgencias & ")) "
    If Len(Trim(psMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(P.cCtaCod, 9, 1) IN (" & psMoneda & ")) "
    End If
    If Len(Trim(psProductos)) > 0 Then
        lsSQL = lsSQL & " AND (C.cTpoCredCod IN (" & psProductos & ")) "
    End If
    lsSQL = lsSQL & " and (convert(varchar(8), CR.dIngRecup,112) BETWEEN '" & Format(dFechaPasada, "YYYYMMdd") & "' and '" & Format(pdFecFin, "YYYYMMdd") & "') " & _
    " GROUP BY RRHH.cUser "
  
    lsSQL = lsSQL & " ORDER BY RRHH.cUser"
    
    'Lugar: 1 si es total, 2 si es saldos , 3 si es desembolsos nuevos, 4 si son desembolsos represtamos
    '       5 si son desembolsos refinanciados, 6 total para la cartera morosa, 7 saldos hasta el mes pasado
    '       para la cartera morosa, 8 ingresos a recuperacioness
 
    Set loDataRep = New dColPFunciones
        
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        lrDataRep.Close
        Set lrDataRep = Nothing
        Exit Function
    Else
        nContador = 0
        sTempoAnalista = IIf(IsNull(lrDataRep!ANALISTA), "", lrDataRep!ANALISTA)

        ReDim Preserve matAnalista(nContador) As String
        ReDim Preserve matVMonN1(nContador) As Long
        ReDim Preserve matVMonN2(nContador) As Currency
        ReDim Preserve matVMonE1(nContador) As Long
        ReDim Preserve matVMonE2(nContador) As Currency
        ReDim Preserve matVTot1(nContador) As Long
        ReDim Preserve matVTot2(nContador) As Currency
        ReDim Preserve matVSal1(nContador) As Long
        ReDim Preserve matVSal2(nContador) As Currency
        ReDim Preserve matMMonN1(nContador) As Long
        ReDim Preserve matMMonN2(nContador) As Currency
        ReDim Preserve matMMonE1(nContador) As Long
        ReDim Preserve matMMonE2(nContador) As Currency
        ReDim Preserve matMTot1(nContador) As Long
        ReDim Preserve matMTot2(nContador) As Currency
        ReDim Preserve matMSal1(nContador) As Long
        ReDim Preserve matMSal2(nContador) As Currency
        ReDim Preserve matIngRecup(nContador) As Currency
 
        matAnalista(0) = IIf(IsNull(lrDataRep!ANALISTA), "", lrDataRep!ANALISTA)

        RaiseEvent ShowProgress
 
        Do While Not lrDataRep.EOF
            If sTempoAnalista <> lrDataRep!ANALISTA Then
                nContador = nContador + 1
                ReDim Preserve matAnalista(nContador) As String
                ReDim Preserve matVMonN1(nContador) As Long
                ReDim Preserve matVMonN2(nContador) As Currency
                ReDim Preserve matVMonE1(nContador) As Long
                ReDim Preserve matVMonE2(nContador) As Currency
                ReDim Preserve matVTot1(nContador) As Long
                ReDim Preserve matVTot2(nContador) As Currency
                ReDim Preserve matVSal1(nContador) As Long
                ReDim Preserve matVSal2(nContador) As Currency
                ReDim Preserve matMMonN1(nContador) As Long
                ReDim Preserve matMMonN2(nContador) As Currency
                ReDim Preserve matMMonE1(nContador) As Long
                ReDim Preserve matMMonE2(nContador) As Currency
                ReDim Preserve matMTot1(nContador) As Long
                ReDim Preserve matMTot2(nContador) As Currency
                ReDim Preserve matMSal1(nContador) As Long
                ReDim Preserve matMSal2(nContador) As Currency
                ReDim Preserve matIngRecup(nContador) As Currency
 
                matAnalista(nContador) = lrDataRep!ANALISTA
                sTempoAnalista = lrDataRep!ANALISTA
            End If
            
            With lrDataRep
                Select Case !Lugar
                Case 1
                    matVTot1(nContador) = !cantidad
                    matVTot2(nContador) = !Total
                Case 2
                    matVSal1(nContador) = !cantidad
                    matVSal2(nContador) = !Total
                Case 6
                    matMTot1(nContador) = !cantidad
                    matMTot2(nContador) = !Total
                Case 7
                    matMSal1(nContador) = !cantidad
                    matMSal2(nContador) = !Total
                Case 8
                    matIngRecup(nContador) = !Total
                Case 10
                    matVMonN1(nContador) = !cantidad
                    matVMonN2(nContador) = !Total
                Case 11
                    matVMonE1(nContador) = !cantidad
                    matVMonE2(nContador) = !Total
                Case 20
                    matMMonN1(nContador) = !cantidad
                    matMMonN2(nContador) = !Total
                Case 21
                    matMMonE1(nContador) = !cantidad
                    matMMonE2(nContador) = !Total
                End Select
            End With
            
            RaiseEvent Progress(lrDataRep.Bookmark, lrDataRep.RecordCount)
            lrDataRep.MoveNext
        Loop
        lrDataRep.Close
        Set lrDataRep = Nothing
        
        RaiseEvent CloseProgress
        
        'Calculo los totales
        For i = 0 To nContador
           
            'Vigentes
            '========
            'cantidad de moneda nacional para vigente
            matTot1(0) = matTot1(0) + matVMonN1(i)
            'total de moneda nacional para vigente
            matTot1(1) = matTot1(1) + matVMonN2(i)
            'cantidad de moneda extranjera para vigente
            matTot1(2) = matTot1(2) + matVMonE1(i)
            'total de moneda extranjera para vigente
            matTot1(3) = matTot1(3) + matVMonE2(i)
            
            'Total de Numero para Totales
            matTot1(4) = matTot1(4) + matVTot1(i)
            'Total de Totales para Totales
            matTot1(5) = matTot1(5) + matVTot2(i)
            'Total de Numero para Saldos
            matTot1(6) = matTot1(6) + matVSal1(i)
            'Total de Totales para Saldos
            matTot1(7) = matTot1(7) + matVSal2(i)
            'Variacion a la fecha para vigente de numero
            matTot1(8) = matTot1(8) + matVTot1(i) - matVSal1(i)
            'Variacion a la fecha para vigente de monto
            matTot1(9) = matTot1(9) + matVTot2(i) - matVSal2(i)
            'Ingresos a Judicial
            matTot1(10) = matTot1(10) + matIngRecup(i)
            'variacion real
            matTot1(11) = matTot1(11) + matVTot2(i) - matVSal2(i) + matIngRecup(i)
            
            'Morosos
            '=======
             
            'cantidad de moneda nacional para moroso
            matTot2(0) = matTot2(0) + matMMonN1(i)
            'total de moneda nacional para moroso
            matTot2(1) = matTot2(1) + matMMonN2(i)
            'cantidad de moneda extranjera para moroso
            matTot2(2) = matTot2(2) + matMMonE1(i)
            'total de moneda extranjera para moroso
            matTot2(3) = matTot2(3) + matMMonE2(i)
             
            'Total de Numero para Totales
            matTot2(4) = matTot2(4) + matMTot1(i)
            'Total de Totales para Totales
            matTot2(5) = matTot2(5) + matMTot2(i)
            'Total de Numero para Saldos
            matTot2(6) = matTot2(6) + matMSal1(i)
            'Total de Totales para Saldos
            matTot2(7) = matTot2(7) + matMSal2(i)
            'Variacion a la fecha para vigente de numero
            matTot2(8) = matTot2(8) + matMTot1(i) - matMSal1(i)
            'Variacion a la fecha para vigente de monto
            matTot2(9) = matTot2(9) + matMTot2(i) - matMSal2(i)
             
        Next
                
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CUADRO CONSOLIDADO DE COLOCACIONES X MORA Y ANALISTA", lsCondiciones, lnPage, 154, cSubTit & Chr(10) & "TCF. = " & Format(pnTipoCambio, "0.000") & " // TCF. (" & dFechaPasada & ") = " & Format(nTipoCambioPasado, "0.000"), 108606)
        lsCadImp = lsCadImp & Chr(10)
        
        lsCadImp = lsCadImp & lsNegritaOn & "C A R T E R A  V I G E N T E" & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(154, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "        |  C  A  R  T  E  R  A   V  I  G  E  N  T  E  |    T  O  T  A  L     |   S A L D O S  A L   |  V A R I A C I O N   |   INGRESO A   |   VARIACION   " & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA|  MONEDA  NACIONAL    |  MONEDA EXTRANJERA   |  (En Nuevos Soles)   |      " & dFechaPasada & "      |  A  L A  F E C H A   | COBR. JUDIC.  |  R  E  A  L   " & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "        | NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |   M O N T O   |               |               " & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(154, "-") & Chr(10)
        
        lnIndice = 14:  lnLineas = 14
             
        For i = 0 To nContador
            lnIndice = lnIndice + 1
            lnLineas = lnLineas + 1
 
            cUser = matAnalista(i)
            lsCadImp = lsCadImp & Space(1) & cUser & Space(2) & _
                    " | " & ImpreFormat(matVMonN1(i), 4, 0) & " | " & ImpreFormat(matVMonN2(i), 10, 2) & _
                    " | " & ImpreFormat(matVMonE1(i), 4, 0) & " | " & ImpreFormat(matVMonE2(i), 10, 2) & _
                    " | " & ImpreFormat(matVTot1(i), 4, 0) & " | " & ImpreFormat(matVTot2(i), 10, 2) & _
                    " | " & ImpreFormat(matVSal1(i), 4, 0) & " | " & ImpreFormat(matVSal2(i), 10, 2) & _
                    " | " & ImpreFormat(matVTot1(i) - matVSal1(i), 4, 0) & " | " & ImpreFormat(matVTot2(i) - matVSal2(i), 10, 2) & _
                    " | " & ImpreFormat(matIngRecup(i), 10, 2) & " | " & ImpreFormat(matVTot2(i) - matVSal2(i) + matIngRecup(i), 10, 2) & _
                    Chr(10)
          
            If lnIndice Mod 300 = 0 Then
                lsCadBuffer = lsCadBuffer & lsCadImp
                lsCadImp = ""
            End If
        
            If lnLineas >= 55 Then
                lnPage = lnPage + 1
                lsCadImp = lsCadImp & Chr(12)
                lsCadImp = lsCadImp & nRepoCabecera("CUADRO CONSOLIDADO DE COLOCACIONES X MORA Y ANALISTA", lsCondiciones, lnPage, 154, cSubTit & Chr(10) & "TCF. = " & Format(pnTipoCambio, "0.00") & " // TCF. (" & dFechaPasada & ") = " & Format(nTipoCambioPasado, "0.00"), 108606)
                lnLineas = 7
                lnIndice = lnIndice + 7
            End If
        Next
        
        'Imprimo los totales de vigentes
        lsCadImp = lsCadImp & String(154, "=") & Chr(10)
        
        lsCadImp = lsCadImp & lsNegritaOn & Space(7) & " | " & ImpreFormat(matTot1(0), 4, 0) & " | " & ImpreFormat(matTot1(1), 10, 2) & _
                " | " & ImpreFormat(matTot1(2), 4, 0) & " | " & ImpreFormat(matTot1(3), 10, 2) & _
                " | " & ImpreFormat(matTot1(4), 4, 0) & " | " & ImpreFormat(matTot1(5), 10, 2) & _
                " | " & ImpreFormat(matTot1(6), 4, 0) & " | " & ImpreFormat(matTot1(7), 10, 2) & _
                " | " & ImpreFormat(matTot1(8), 4, 0) & " | " & ImpreFormat(matTot1(9), 10, 2) & _
                " | " & ImpreFormat(matTot1(10), 10, 2) & " | " & ImpreFormat(matTot1(11), 10, 2) & _
                 lsNegritaOff & Chr(10)
        
        lsCadImp = lsCadImp & String(154, "_") & Chr(10) & Chr(10) & Chr(10)
        
        lsCadImp = lsCadImp & lsNegritaOn & "C A R T E R A  M O R O S A" & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(154, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "        |  C  A  R  T  E  R  A   V  I  G  E  N  T  E  |    T  O  T  A  L     |   S A L D O S  A L   |  V A R I A C I O N   " & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA|  MONEDA  NACIONAL    |  MONEDA EXTRANJERA   |  (En Nuevos Soles)   |      " & dFechaPasada & "      |  A  L A  F E C H A   " & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "        | NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |   M O N T O   | NUM. |   M O N T O   " & lsNegritaOff & Chr(10)
        
        lsCadImp = lsCadImp & String(154, "-") & Chr(10)
        
        lnIndice = lnIndice + 10:  lnLineas = lnLineas + 10
             
        For i = 0 To nContador
            lnIndice = lnIndice + 1
            lnLineas = lnLineas + 1
            cUser = matAnalista(i)
            
            lsCadImp = lsCadImp & Space(1) & cUser & Space(2) & _
                    " | " & ImpreFormat(matMMonN1(i), 4, 0) & " | " & ImpreFormat(matMMonN2(i), 10, 2) & _
                    " | " & ImpreFormat(matMMonE1(i), 4, 0) & " | " & ImpreFormat(matMMonE2(i), 10, 2) & _
                    " | " & ImpreFormat(matMTot1(i), 4, 0) & " | " & ImpreFormat(matMTot2(i), 10, 2) & _
                    " | " & ImpreFormat(matMSal1(i), 4, 0) & " | " & ImpreFormat(matMSal2(i), 10, 2) & _
                    " | " & ImpreFormat(matMTot1(i) - matMSal1(i), 4, 0) & " | " & ImpreFormat(matMTot2(i) - matMSal2(i), 10, 2) & _
                    Chr(10)
            
            If lnIndice Mod 300 = 0 Then
                lsCadBuffer = lsCadBuffer & lsCadImp
                lsCadImp = ""
            End If
        
            If lnLineas >= 55 Then
                lnPage = lnPage + 1
                lsCadImp = lsCadImp & Chr(12)
                lsCadImp = lsCadImp & nRepoCabecera("CUADRO CONSOLIDADO DE COLOCACIONES X MORA Y ANALISTA", lsCondiciones, lnPage, 154, cSubTit & Chr(10) & "TCF. = " & Format(pnTipoCambio, "0.00") & " // TCF. (" & dFechaPasada & ") = " & Format(nTipoCambioPasado, "0.00"), 108606)
                lnLineas = 7
                lnIndice = lnIndice + 7
            End If
        Next
        
        'Imprimo los totales de morosos
        lsCadImp = lsCadImp & String(154, "=") & Chr(10)
        
        lsCadImp = lsCadImp & lsNegritaOn & Space(7) & " | " & ImpreFormat(matTot2(0), 4, 0) & " | " & ImpreFormat(matTot2(1), 10, 2) & _
                    " | " & ImpreFormat(matTot2(2), 4, 0) & " | " & ImpreFormat(matTot2(3), 10, 2) & _
                    " | " & ImpreFormat(matTot2(4), 4, 0) & " | " & ImpreFormat(matTot2(5), 10, 2) & _
                    " | " & ImpreFormat(matTot2(6), 4, 0) & " | " & ImpreFormat(matTot2(7), 10, 2) & _
                    " | " & ImpreFormat(matTot2(8), 4, 0) & " | " & ImpreFormat(matTot2(9), 10, 2) & _
                    lsNegritaOff & Chr(10)
        
        lsCadImp = lsCadImp & String(154, "_") & Chr(10) & Chr(10)

        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
 
    End If
      nRepo108606_ConsolidadoColocacionesxMoraxAnalista = lsCadBuffer
    
End Function


Public Function nRepo108404_SaldosCarteraxAnalista(ByVal cSubTit As String, ByVal pdFecFin As String, ByVal pnDia1I As Integer, ByVal pnDia1F As Integer, ByVal pnDia2I As Integer, ByVal pnDia2F As Integer, ByVal pnDia3I As Integer, ByVal pcMoneda As String, ByVal pcTipoCredito As String, ByVal pcCondicion As String, ByVal pcAgencias As String, ByVal pcAnalistas As String) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nitem As Long
Dim loImprimeCab As NColPImpre
Dim sCliente As String * 35
Dim sMoneda As String * 7
Dim lsCondiciones As String
Dim sCriterio As String
Dim i As Integer
Dim J As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim matTNac As Variant
Dim matTExtr As Variant
Dim nTopeNac As Long
Dim nTopeExtr As Long
Dim nTexto1 As String * 9
Dim nTexto2 As String * 9
Dim nTexto3 As String * 9
Dim sNombreAnalista As String * 15
Dim nTempoMoneda As Integer

 'Totales Generales
Dim matTotales(1, 13) As Currency

lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

lsSQL = "select C1.nMOneda, C1.cPersNombre, C1.cUser, C1.cPersCod, sum(C1.nCantidadA) as CN1, sum(C1.nTotalA) as CT1, sum(C1.nCantidadB) as CN2, sum(C1.nTotalB) as CT2, sum(C1.nCantidadC) AS CN3, sum(C1.nTotalC) as CT3, sum(C1.nCantidadD) as CN4, sum(C1.nTotalD) as CT4  FROM( "

'BRGO 07/06/2010 BASILEA II
lsSQL = lsSQL & " SELECT SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) AS nMoneda, COUNT(dbo.Colocaciones.cCtaCod) AS nCantidadA, SUM(dbo.Producto.nSaldo) AS nTotalA, " & _
        " 0 as nCantidadB, 0 as nTotalB, 0 as nCantidadC, 0 as nTotalC, 0 as nCantidadD, 0 as nTotalD, " & _
        " dbo.ProductoPersona.cPersCod , dbo.RRHH.cUser, dbo.Persona.cPersNombre " & _
        " FROM dbo.Colocaciones INNER JOIN dbo.ColocacCred ON dbo.Colocaciones.cCtaCod = dbo.ColocacCred.cCtaCod " & _
        " INNER JOIN dbo.Producto ON dbo.Colocaciones.cCtaCod = dbo.Producto.cCtaCod INNER JOIN " & _
        " dbo.ProductoPersona ON dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod AND " & _
        " dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod INNER JOIN " & _
        " dbo.RRHH ON dbo.ProductoPersona.cPersCod = dbo.RRHH.cPersCod INNER JOIN " & _
        " dbo.Persona ON dbo.RRHH.cPersCod = dbo.Persona.cPersCod " & _
        " WHERE (dbo.ProductoPersona.nPrdPersRelac=" & gColRelPersAnalista & ") " & _
        " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) "

    If Len(Trim(pcCondicion)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.ColocacCred.nColocCondicion IN (" & pcCondicion & ")) "
    End If

    If Len(Trim(pcTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.Colocaciones.cTpoCredCod IN (" & pcTipoCredito & ")) "
    End If

    If Len(Trim(pcMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) IN (" & pcMoneda & ")) "
    End If

    If Len(Trim(pcAgencias)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.Colocaciones.cAgeCodAct IN (" & pcAgencias & ")) "
    End If

    If Len(Trim(pcAnalistas)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.ProductoPersona.cPersCod in (" & pcAnalistas & ")) "
    End If

    lsSQL = lsSQL & " GROUP BY dbo.ProductoPersona.cPersCod, dbo.RRHH.cUser, SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1), dbo.Persona.cPersNombre "

lsSQL = lsSQL & " Union All "

'--Cartera Vencida A

lsSQL = lsSQL & " SELECT SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) AS nMoneda, 0 as nCantidadA, 0 as nTotalA, COUNT(dbo.Colocaciones.cCtaCod) AS nCantidadB, SUM(dbo.Producto.nSaldo) AS nTotalB, " & _
        " 0 as nCantidadC, 0 as nTotalC, 0 as nCantidadD, 0 as nTotalD, " & _
        " dbo.ProductoPersona.cPersCod , dbo.RRHH.cUser, dbo.Persona.cPersNombre " & _
        " FROM dbo.Colocaciones INNER JOIN dbo.ColocacCred ON dbo.Colocaciones.cCtaCod = dbo.ColocacCred.cCtaCod " & _
        " INNER JOIN dbo.Producto ON dbo.Colocaciones.cCtaCod = dbo.Producto.cCtaCod INNER JOIN " & _
        " dbo.ProductoPersona ON dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod AND " & _
        " dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod INNER JOIN " & _
        " dbo.RRHH ON dbo.ProductoPersona.cPersCod = dbo.RRHH.cPersCod INNER JOIN " & _
        " dbo.Persona ON dbo.RRHH.cPersCod = dbo.Persona.cPersCod " & _
        " WHERE (dbo.ProductoPersona.nPrdPersRelac=" & gColRelPersAnalista & ") " & _
        " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) "

    lsSQL = lsSQL & " AND (dbo.ColocacCred.nDiasAtraso>=" & pnDia1I & ") AND (dbo.ColocacCred.nDiasAtraso<=" & pnDia1F & ") "

    If Len(Trim(pcCondicion)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.ColocacCred.nColocCondicion IN (" & pcCondicion & ")) "
    End If

    If Len(Trim(pcTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.Colocaciones.cTpoCredCod IN (" & pcTipoCredito & ")) "
    End If

    If Len(Trim(pcMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) IN (" & pcMoneda & ")) "
    End If

    If Len(Trim(pcAgencias)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.Colocaciones.cAgeCodAct IN (" & pcAgencias & ")) "
    End If

    If Len(Trim(pcAnalistas)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.ProductoPersona.cPersCod in (" & pcAnalistas & ")) "
    End If

    lsSQL = lsSQL & " GROUP BY dbo.ProductoPersona.cPersCod, dbo.RRHH.cUser, SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1), dbo.Persona.cPersNombre "

lsSQL = lsSQL & " Union All "

'--Cartera Vencida B
lsSQL = lsSQL & " SELECT SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) AS nMoneda, 0 as nCantidadA, 0 as nTotalA, 0 as nCantidadB, 0 as nTotalB, COUNT(dbo.Colocaciones.cCtaCod) AS nCantidadC, SUM(dbo.Producto.nSaldo) AS nTotalC, " & _
        " 0 as nCantidadD, 0 as nTotalD, " & _
        " dbo.ProductoPersona.cPersCod , dbo.RRHH.cUser, dbo.Persona.cPersNombre " & _
        " FROM dbo.Colocaciones INNER JOIN dbo.ColocacCred ON dbo.Colocaciones.cCtaCod = dbo.ColocacCred.cCtaCod " & _
        " INNER JOIN dbo.Producto ON dbo.Colocaciones.cCtaCod = dbo.Producto.cCtaCod INNER JOIN " & _
        " dbo.ProductoPersona ON dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod AND " & _
        " dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod INNER JOIN " & _
        " dbo.RRHH ON dbo.ProductoPersona.cPersCod = dbo.RRHH.cPersCod INNER JOIN " & _
        " dbo.Persona ON dbo.RRHH.cPersCod = dbo.Persona.cPersCod " & _
        " WHERE (dbo.ProductoPersona.nPrdPersRelac=" & gColRelPersAnalista & ") " & _
        " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) "

    lsSQL = lsSQL & " AND (dbo.ColocacCred.nDiasAtraso>=" & pnDia2I & ") AND (dbo.ColocacCred.nDiasAtraso<=" & pnDia2F & ") "

    If Len(Trim(pcCondicion)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.ColocacCred.nColocCondicion IN (" & pcCondicion & ")) "
    End If

    If Len(Trim(pcTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.Colocaciones.cTpoCredCod IN (" & pcTipoCredito & ")) "
    End If

    If Len(Trim(pcMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) IN (" & pcMoneda & ")) "
    End If

    If Len(Trim(pcAgencias)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.Colocaciones.cAgeCodAct IN (" & pcAgencias & ")) "
    End If

    If Len(Trim(pcAnalistas)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.ProductoPersona.cPersCod in (" & pcAnalistas & ")) "
    End If

    lsSQL = lsSQL & " GROUP BY dbo.ProductoPersona.cPersCod, dbo.RRHH.cUser, SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1), dbo.Persona.cPersNombre "

    lsSQL = lsSQL & " UNION ALL "

'--Cartera Vencida C

lsSQL = lsSQL & " SELECT SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) AS nMoneda, 0 as nCantidadA, 0 as nTotalA, 0 as nCantidadB, 0 as nTotalB, 0 as nCantidadC, 0 as nTotalC, COUNT(dbo.Colocaciones.cCtaCod) AS nCantidadD, SUM(dbo.Producto.nSaldo) AS nTotalD, " & _
        " dbo.ProductoPersona.cPersCod , dbo.RRHH.cUser, dbo.Persona.cPersNombre " & _
        " FROM dbo.Colocaciones INNER JOIN dbo.ColocacCred ON dbo.Colocaciones.cCtaCod = dbo.ColocacCred.cCtaCod " & _
        " INNER JOIN dbo.Producto ON dbo.Colocaciones.cCtaCod = dbo.Producto.cCtaCod INNER JOIN " & _
        " dbo.ProductoPersona ON dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod AND " & _
        " dbo.Producto.cCtaCod = dbo.ProductoPersona.cCtaCod INNER JOIN " & _
        " dbo.RRHH ON dbo.ProductoPersona.cPersCod = dbo.RRHH.cPersCod INNER JOIN " & _
        " dbo.Persona ON dbo.RRHH.cPersCod = dbo.Persona.cPersCod " & _
        " WHERE (dbo.ProductoPersona.nPrdPersRelac=" & gColRelPersAnalista & ") " & _
        " AND (Producto.nPrdEstado IN ('" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "')) "

    lsSQL = lsSQL & " AND (dbo.ColocacCred.nDiasAtraso>=" & pnDia3I & ") "

    If Len(Trim(pcCondicion)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.ColocacCred.nColocCondicion IN (" & pcCondicion & ")) "
    End If

    If Len(Trim(pcTipoCredito)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.Colocaciones.cTpoCredCod IN (" & pcTipoCredito & ")) "
    End If

    If Len(Trim(pcMoneda)) > 0 Then
        lsSQL = lsSQL & " AND (SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1) IN (" & pcMoneda & ")) "
    End If

    If Len(Trim(pcAgencias)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.Colocaciones.cAgeCodAct IN (" & pcAgencias & ")) "
    End If

    If Len(Trim(pcAnalistas)) > 0 Then
        lsSQL = lsSQL & " AND (dbo.ProductoPersona.cPersCod in (" & pcAnalistas & ")) "
    End If

    lsSQL = lsSQL & " GROUP BY dbo.ProductoPersona.cPersCod, dbo.RRHH.cUser, SUBSTRING(dbo.Colocaciones.cCtaCod, 9, 1), dbo.Persona.cPersNombre "


    lsSQL = lsSQL & ") C1 GROUP BY C1.cPersCod, cUser, nMoneda, cPersNombre   ORDER BY C1.nMoneda, C1.cUser"

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        lsCondiciones = "Al " & pdFecFin

        nTexto1 = ImpreFormat(pnDia1I, 4, 0) & "-" & ImpreFormat(pnDia1F, 4, 0)
        nTexto2 = ImpreFormat(pnDia2I, 4, 0) & "-" & ImpreFormat(pnDia2F, 4, 0)
        nTexto3 = ">" & ImpreFormat(pnDia3I, 4, 0)

        lrDataRep.Filter = "nMoneda=" & gMonedaNacional

        If lrDataRep.BOF Then
        Else
            Do While Not lrDataRep.EOF
                nTopeNac = nTopeNac + 1
                lrDataRep.MoveNext
            Loop

            ReDim matTNac(nTopeNac, 15)

            lrDataRep.MoveFirst
            i = 0
            Do While Not lrDataRep.EOF
                matTNac(i, 0) = lrDataRep!cUser
                matTNac(i, 1) = lrDataRep!cPersNombre
                matTNac(i, 2) = lrDataRep!CN1
                matTNac(i, 3) = lrDataRep!CT1
                matTNac(i, 4) = lrDataRep!CN2
                matTNac(i, 5) = lrDataRep!CT2
                matTNac(i, 7) = lrDataRep!CN3
                matTNac(i, 8) = lrDataRep!CT3
                matTNac(i, 10) = lrDataRep!CN4
                matTNac(i, 11) = lrDataRep!CT4
                i = i + 1
                lrDataRep.MoveNext
            Loop
        End If

        lrDataRep.Filter = "nMoneda=" & gMonedaExtranjera

        If lrDataRep.BOF Then
        Else
            Do While Not lrDataRep.EOF
                nTopeExtr = nTopeExtr + 1
                lrDataRep.MoveNext
            Loop

            ReDim matTExtr(nTopeExtr, 15)

            lrDataRep.MoveFirst
            i = 0
            Do While Not lrDataRep.EOF
                matTExtr(i, 0) = lrDataRep!cUser
                matTExtr(i, 1) = lrDataRep!cPersNombre
                matTExtr(i, 2) = lrDataRep!CN1
                matTExtr(i, 3) = lrDataRep!CT1
                matTExtr(i, 4) = lrDataRep!CN2
                matTExtr(i, 5) = lrDataRep!CT2
                matTExtr(i, 7) = lrDataRep!CN3
                matTExtr(i, 8) = lrDataRep!CT3
                matTExtr(i, 10) = lrDataRep!CN4
                matTExtr(i, 11) = lrDataRep!CT4
                i = i + 1
                lrDataRep.MoveNext
            Loop
        End If

        lrDataRep.Close
        Set lrDataRep = Nothing

        For i = 0 To nTopeNac - 1

            If matTNac(i, 3) = 0 Then
                matTNac(i, 6) = 0
                matTNac(i, 9) = 0
                matTNac(i, 12) = 0
            Else
                matTNac(i, 6) = matTNac(i, 5) / matTNac(i, 3)
                matTNac(i, 9) = matTNac(i, 8) / matTNac(i, 3)
                matTNac(i, 12) = Format(matTNac(i, 11) / matTNac(i, 3), "#0.00")
            End If


            matTNac(i, 13) = matTNac(i, 4) + matTNac(i, 7) + matTNac(i, 10)
            matTNac(i, 14) = matTNac(i, 5) + matTNac(i, 8) + matTNac(i, 11)
            matTNac(i, 15) = Format(matTNac(i, 14) / matTNac(i, 3), "#0.00")

            For J = 0 To 13
                matTotales(0, J) = matTotales(0, J) + matTNac(i, J + 2)
            Next

        Next

        For i = 0 To nTopeExtr - 1
            If matTExtr(i, 3) = 0 Then
                matTExtr(i, 6) = 0
                matTExtr(i, 9) = 0
                matTExtr(i, 12) = 0
                matTExtr(i, 15) = 0
            Else
                matTExtr(i, 6) = matTExtr(i, 5) / matTExtr(i, 3)
                matTExtr(i, 9) = matTExtr(i, 8) / matTExtr(i, 3)
                matTExtr(i, 12) = Format(matTExtr(i, 11) / matTExtr(i, 3), "#0.00")

            End If
            matTExtr(i, 13) = matTExtr(i, 4) + matTExtr(i, 7) + matTExtr(i, 10)
            matTExtr(i, 14) = matTExtr(i, 5) + matTExtr(i, 8) + matTExtr(i, 11)
            matTExtr(i, 15) = Format(matTExtr(i, 14) / matTExtr(i, 3), "#0.00")

            For J = 0 To 13
                matTotales(1, J) = matTotales(1, J) + matTExtr(i, J + 2)
            Next

        Next

        'With lrDataRep
            RaiseEvent ShowProgress
            'Do While Not lrDataRep.EOF
                nitem = nitem + 1
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1

                If nTopeNac > 0 Then
                    lsCadImp = lsCadImp & nRepoCabecera("RESUMEN DE SALDOS DE CARTERA POR ANALISTA - CREDITOS CONSUMO", lsCondiciones & " MONEDA: SOLES", lnPage, 160, cSubTit, 108404)
                    lsCadImp = lsCadImp & String(160, "-") & Chr(10)
                    lsCadImp = lsCadImp & "COD.|ANALISTA       | CARTERA TOTAL  |   CARTERA VENCIDA " & nTexto1 & " |  CARTERA VENCIDA " & nTexto2 & "|   CARTERA VENCIDA " & nTexto3 & "| TOTAL CARTERA VENCIDA" & Chr(10)
                    lsCadImp = lsCadImp & "                    |NRO.|     SALDO |NRO.|      SALDO|       % MORA|NRO.|    SALDO|     % MORA|NRO.|      SALDO|     % MORA|NRO.|      SALDO|     % MORA" & Chr(10)
                    lsCadImp = lsCadImp & String(160, "-") & Chr(10)

                    lnLineas = 7
                    lnIndice = lnIndice + 7

                    For i = 0 To nTopeNac - 1
                        lnLineas = lnLineas + 1
                        sNombreAnalista = matTNac(i, 1)
                        lsCadImp = lsCadImp & matTNac(i, 0) & Space(1) & sNombreAnalista & Space(1)
                        For J = 2 To 15
                            If J = 2 Or J = 4 Or J = 7 Or J = 10 Or J = 13 Then
                                lsCadImp = lsCadImp & ImpreFormat(matTNac(i, J), 4, 0) & Space(1)
                            Else
                                lsCadImp = lsCadImp & ImpreFormat(matTNac(i, J), 8, 2) & Space(1)
                            End If
                        Next
                        lsCadImp = lsCadImp & Chr(10)

                        If lnLineas >= 55 Then
                            lnPage = lnPage + 1
                            lsCadImp = lsCadImp & Chr(12)
                            lsCadImp = lsCadImp & nRepoCabecera("RESUMEN DE SALDOS DE CARTERA POR ANALISTA - CREDITOS CONSUMO", lsCondiciones & " MONEDA: SOLES", lnPage, 160, cSubTit, 108404)
                            lsCadImp = lsCadImp & String(160, "-") & Chr(10)
                            lsCadImp = lsCadImp & "COD.|ANALISTA       | CARTERA TOTAL  |   CARTERA VENCIDA " & nTexto1 & " |  CARTERA VENCIDA " & nTexto2 & "|   CARTERA VENCIDA " & nTexto3 & "| TOTAL CARTERA VENCIDA" & Chr(10)
                            lsCadImp = lsCadImp & "                    |NRO.|     SALDO |NRO.|      SALDO|       % MORA|NRO.|    SALDO|     % MORA|NRO.|      SALDO|     % MORA|NRO.|      SALDO|     % MORA" & Chr(10)
                            lsCadImp = lsCadImp & String(160, "-") & Chr(10)

                            lnLineas = 7
                            lnIndice = lnIndice + 7
                        End If
                        RaiseEvent Progress(nitem, nTopeNac + nTopeExtr)
                    Next
                    lsCadImp = lsCadImp & String(160, "-") & Chr(10)

                    lsCadImp = lsCadImp & Space(21)
                    For i = 0 To 13
                        If i = 0 Or i = 2 Or i = 5 Or i = 8 Or i = 11 Then
                            lsCadImp = lsCadImp & ImpreFormat(matTotales(0, i), 4, 0) & Space(1)
                        Else
                            lsCadImp = lsCadImp & ImpreFormat(matTotales(0, i), 8, 2) & Space(1)
                        End If
                    Next
                    lsCadImp = lsCadImp & Chr(10)

                    lsCadImp = lsCadImp & String(160, "-") & Chr(10)
                    lnLineas = lnLineas + 3
                End If

                If nTopeExtr > 0 Then

                    If nTopeNac > 0 Then
                        lnPage = lnPage + 1
                        lsCadImp = lsCadImp & Chr(12)
                    End If

                    lsCadImp = lsCadImp & nRepoCabecera("RESUMEN DE SALDOS DE CARTERA POR ANALISTA - CREDITOS CONSUMO", lsCondiciones & " MONEDA: DOLARES", lnPage, 160, cSubTit, 108404)
                    lsCadImp = lsCadImp & String(160, "-") & Chr(10)
                    lsCadImp = lsCadImp & "COD.|ANALISTA       | CARTERA TOTAL  |   CARTERA VENCIDA " & nTexto1 & " |  CARTERA VENCIDA " & nTexto2 & "|   CARTERA VENCIDA " & nTexto3 & "| TOTAL CARTERA VENCIDA" & Chr(10)
                    lsCadImp = lsCadImp & "                    |NRO.|     SALDO |NRO.|      SALDO|       % MORA|NRO.|    SALDO|     % MORA|NRO.|      SALDO|     % MORA|NRO.|      SALDO|     % MORA" & Chr(10)
                    lsCadImp = lsCadImp & String(160, "-") & Chr(10)
                    lnLineas = 7
                    lnIndice = lnIndice + 7

                   ' For I = 0 To nTopeNac - 1
                    For i = 0 To nTopeExtr - 1
                        lnLineas = lnLineas + 1
                        sNombreAnalista = matTExtr(i, 1)
                        lsCadImp = lsCadImp & matTExtr(i, 0) & Space(1) & sNombreAnalista & Space(1)
                        For J = 2 To 15
                            If J = 2 Or J = 4 Or J = 7 Or J = 10 Or J = 13 Then
                                lsCadImp = lsCadImp & ImpreFormat(matTExtr(i, J), 4, 0) & Space(1)
                            Else
                                lsCadImp = lsCadImp & ImpreFormat(matTExtr(i, J), 8, 2) & Space(1)
                            End If
                        Next
                        lsCadImp = lsCadImp & Chr(10)

                        If lnLineas >= 55 Then
                            lnPage = lnPage + 1
                            lsCadImp = lsCadImp & Chr(12)
                            lsCadImp = lsCadImp & nRepoCabecera("RESUMEN DE SALDOS DE CARTERA POR ANALISTA - CREDITOS CONSUMO", lsCondiciones & " MONEDA: SOLES", lnPage, 160, cSubTit, 108404)
                            lsCadImp = lsCadImp & String(160, "-") & Chr(10)
                            lsCadImp = lsCadImp & "COD.|ANALISTA       | CARTERA TOTAL  |   CARTERA VENCIDA " & nTexto1 & " |  CARTERA VENCIDA " & nTexto2 & "|   CARTERA VENCIDA " & nTexto3 & "| TOTAL CARTERA VENCIDA" & Chr(10)
                            lsCadImp = lsCadImp & "                    |NRO.|     SALDO |NRO.|      SALDO|       % MORA|NRO.|    SALDO|     % MORA|NRO.|      SALDO|     % MORA|NRO.|      SALDO|     % MORA" & Chr(10)
                            lsCadImp = lsCadImp & String(160, "-") & Chr(10)
                            lnLineas = 7
                            lnIndice = lnIndice + 7
                        End If
                        RaiseEvent Progress(nitem, nTopeNac + nTopeExtr)
                    Next

                    lsCadImp = lsCadImp & String(160, "-") & Chr(10)

                    lsCadImp = lsCadImp & Space(21)
                    For i = 0 To 13
                        If i = 0 Or i = 2 Or i = 5 Or i = 8 Or i = 11 Then
                            lsCadImp = lsCadImp & ImpreFormat(matTotales(1, i), 4, 0) & Space(1)
                        Else
                            lsCadImp = lsCadImp & ImpreFormat(matTotales(1, i), 8, 2) & Space(1)
                        End If
                    Next
                    lsCadImp = lsCadImp & Chr(10)

                    lsCadImp = lsCadImp & String(160, "-") & Chr(10)
                    lnLineas = lnLineas + 3

                End If

            RaiseEvent CloseProgress
        'End With

        lsCadBuffer = lsCadBuffer & lsCadImp

     End If
    nRepo108404_SaldosCarteraxAnalista = lsCadBuffer

End Function


Private Sub CuadroExcel(xlHoja1 As Excel.Worksheet, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, Optional lbLineasVert As Boolean = False)
Dim i, J As Integer

For i = X1 To X2
    xlHoja1.Range(xlHoja1.Cells(Y1, i), xlHoja1.Cells(Y1, i)).Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(Y2, i), xlHoja1.Cells(Y2, i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Next i
If lbLineasVert = False Then
    For i = X1 To X2
        For J = Y1 To Y2
            xlHoja1.Range(xlHoja1.Cells(J, i), xlHoja1.Cells(J, i)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Next J
    Next i
End If
If lbLineasVert Then
    For J = Y1 To Y2
        xlHoja1.Range(xlHoja1.Cells(J, X1), xlHoja1.Cells(J, X1)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Next J
End If

For J = Y1 To Y2
    xlHoja1.Range(xlHoja1.Cells(J, X2), xlHoja1.Cells(J, X2)).Borders(xlEdgeRight).LineStyle = xlContinuous
Next J
End Sub

Private Sub ImprimeCabecera(xlHoja As Excel.Worksheet, psNombreAgencia As String)

    '----------------------------------"RECORD POR ANALISTAS DE CREDITOS"-----------------------------------------
    Dim Titulo As String
    
    Titulo = "RECORD POR ANALISTAS DE CREDITOS"
    
    xlHoja.Range("A1:I1").MergeCells = True
    xlHoja.Range("A1").HorizontalAlignment = xlHAlignCenter
    xlHoja.Range("A1").Cells.Font.FontStyle = "Bookman Old Style"
    xlHoja.Range("A1").Cells.Font.Size = 10
    xlHoja.Range("A1").Cells.Font.Bold = True
    xlHoja.Cells(1, 1) = Titulo
    
    
    
    
    '----------------------------------"INFORMACION AL "-----------------------------------------
    
    xlHoja.Range("A3").MergeCells = True
    xlHoja.Range("A3").HorizontalAlignment = xlHAlignLeft
    xlHoja.Range("A3").Cells.Font.FontStyle = "Bookman Old Style"
    xlHoja.Range("A3").Cells.Font.Size = 9
    xlHoja.Range("A3").Cells.Font.Bold = True
    xlHoja.Cells(3, 1) = psNombreAgencia
        
    '----------------------------------"INFORMACION AL "-----------------------------------------
    Dim a As Variant
    a = xlHoja.Range("A2:I2").BorderAround(xlContinuous, xlThick, xlColorIndexAutomatic, 800)
    xlHoja.Range("A2:I2").MergeCells = True
    xlHoja.Range("A2").HorizontalAlignment = xlHAlignCenter
    xlHoja.Range("A2").Cells.Font.FontStyle = "Bookman Old Style"
    xlHoja.Range("A2").Cells.Font.Size = 8
    xlHoja.Range("A2").Cells.Font.Bold = True
    xlHoja.Cells(2, 1) = "INFORMACION AL " & Format(gdFecSis, "dd/mm/yyyy")
    
    
    '----------------------------------"CABECERA DE COLUMNA "-----------------------------------------
    
    xlHoja.Range("A4:I4").HorizontalAlignment = xlHAlignCenter
    xlHoja.Range("A4:I4").Cells.Font.FontStyle = "Bookman Old Style"
    xlHoja.Range("A4:I4").Cells.Font.Size = 6
    xlHoja.Range("A4:I4").Cells.Font.Bold = True
    xlHoja.Range("A4:I4").Cells.Font.Color = RGB(10, 600, 10)
    
    xlHoja.Range("A5:I5").HorizontalAlignment = xlHAlignCenter
    xlHoja.Range("A5:I5").Cells.Font.FontStyle = "Bookman Old Style"
    xlHoja.Range("A5:I5").Cells.Font.Size = 6
    xlHoja.Range("A5:I5").Cells.Font.Bold = True
    xlHoja.Range("A5:I5").Cells.Font.Color = RGB(10, 1000, 10)
    
    xlHoja.Cells(4, 1) = "COD"
    xlHoja.Cells(4, 2) = "NOMBRE DEL ANALISTA"
    xlHoja.Cells(4, 3) = "Nª CREDITOS"
    xlHoja.Cells(4, 4) = "Nª CLIENTES"
    xlHoja.Cells(4, 5) = "COLOCACIONES"
    xlHoja.Cells(4, 6) = "CARTERA ATRASADA (%)"
    xlHoja.Cells(4, 7) = "CARTERA ALTO RIESGO (%)"
    xlHoja.Cells(4, 8) = "MOROSIDAD >1DIA (%)"
    xlHoja.Cells(4, 9) = "CARTERA EN RIESGO(%)"
    
    a = xlHoja.Range("A4:I5").BorderAround(xlContinuous, xlMedium, xlColorIndexAutomatic, 800)
    
    xlHoja.Range("A4:I5").Interior.Color = RGB(10, 100, 10)
    
    xlHoja.Range("A4").Justify
    xlHoja.Range("B4").Justify
    xlHoja.Range("C4").Justify
    xlHoja.Range("D4").Justify
    xlHoja.Range("E4").Justify
    xlHoja.Range("F4").Justify
    xlHoja.Range("G4").Justify
    xlHoja.Range("H4").Justify
    xlHoja.Range("I4").Justify


End Sub

Public Function dObtieneAgencias(Optional pbOrdenCodigo As Boolean = True, Optional lsOrder As String) As Recordset
    Dim lsSQL As String
    Dim lsOrdena As String
    Dim loDataRep As dColPFunciones
    Dim lrst As ADODB.Recordset
    
    On Error GoTo ErrorAge
    
    
    Set loDataRep = New dColPFunciones
    
    
    
    
    lsOrdena = " ORDER BY " & IIf(pbOrdenCodigo = True, " cAgeCod ", " cAgeDescripcion ") & Space(2) & lsOrder
    lsSQL = "SELECT cAgeCod, cAgeDescripcion FROM Agencias" & lsOrdena

    Set lrst = loDataRep.dObtieneRecordSet(lsSQL)
    
    Set dObtieneAgencias = lrst
    
    
Exit Function
ErrorAge:
    Call RaiseError(MyUnhandledError, "DConstantes:GetConstante Method")
End Function

Private Sub CargaMoraAnalistas(xlHoja As Excel.Worksheet, lsCodigoAgencia As String, nFilaInicio As Integer)

    Dim oConect As DConecta
    Dim lsSQL As String
    Dim lrst As ADODB.Recordset
    
    
    Set oConect = New DConecta
    oConect.AbreConexion
        
    lsSQL = "SP_COL_CALIFICACION_MORAxANALISTA '" & lsCodigoAgencia & "'"
    
    oConect.Ejecutar lsSQL
    
    lsSQL = "Select * from MoraAnalistaAux order by CarteraRiesgo"
    
    Set lrst = New ADODB.Recordset
    Set lrst = oConect.CargaRecordSet(lsSQL)

    Dim lnCol As Integer
    
    Dim i As Integer
    
    'MsgBox lrst("analista")
    i = 0
'    MsgBox lrst.RecordCount
    Do While Not lrst.EOF
        xlHoja.Cells(nFilaInicio + i, 1) = IIf(IsNull(lrst("UserAnalistaAnt")), "SSUU", lrst("UserAnalistaAnt"))
        xlHoja.Cells(nFilaInicio + i, 2) = PstaNombre(lrst("analista"), True)
        xlHoja.Cells(nFilaInicio + i, 3) = lrst("creditos")
        xlHoja.Cells(nFilaInicio + i, 4) = lrst("clientes")
        xlHoja.Cells(nFilaInicio + i, 5) = lrst("TotalColocacion")
        xlHoja.Cells(nFilaInicio + i, 6) = lrst("CarteraAtrasada")
        xlHoja.Cells(nFilaInicio + i, 7) = lrst("CarteraAltoRiesgo")
        xlHoja.Cells(nFilaInicio + i, 8) = lrst("CarteraMorosa")
        xlHoja.Cells(nFilaInicio + i, 9) = lrst("CarteraRiesgo")
        i = i + 1
        lrst.MoveNext
    Loop
    
    Dim a As Variant
    a = xlHoja.Range(xlHoja.Cells(nFilaInicio, 1), xlHoja.Cells(nFilaInicio + i, 9)).BorderAround(xlContinuous, xlMedium, xlColorIndexAutomatic, 800)
    xlHoja.Range(xlHoja.Cells(nFilaInicio, 1), xlHoja.Cells(nFilaInicio + i, 2)).HorizontalAlignment = xlHAlignLeft
    xlHoja.Range(xlHoja.Cells(nFilaInicio, 3), xlHoja.Cells(nFilaInicio + i, 9)).HorizontalAlignment = xlHAlignRight
    xlHoja.Range(xlHoja.Cells(nFilaInicio, 1), xlHoja.Cells(nFilaInicio + i, 9)).Cells.Font.Size = 8
    xlHoja.Range(xlHoja.Cells(nFilaInicio, 1), xlHoja.Cells(nFilaInicio + i, 9)).Cells.Font.FontStyle = "Bookman Old Style"
    xlHoja.Range(xlHoja.Cells(nFilaInicio, 1), xlHoja.Cells(nFilaInicio + i, 9)).Cells.Font.Bold = False
    
    'xlHoja.Range(xlHoja.Cells(nFilaInicio, 3), xlHoja.Cells(nFilaInicio + I, 9)).FormatConditions(1) = "#,###,##0.00"

End Sub


