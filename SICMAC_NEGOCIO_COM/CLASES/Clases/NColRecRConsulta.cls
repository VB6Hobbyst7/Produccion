VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NColRecRConsulta"
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
        ByVal pnCodReporte As Long) As String
        
Dim lsCadImp As String
Dim loImprimeCab As NColPImpre
    
    Set loImprimeCab = New NColPImpre
        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    Set loImprimeCab = Nothing
    If pnCodReporte <> 150000 And pnCodReporte <> 138025 Then
        lsCadImp = lsCadImp & Chr(10) & String(pnAnchoLinea, "-") & Chr(10)
    End If
    Select Case pnCodReporte
        
        Case 138001 ' Listado de Pagos x Abogado
            lsCadImp = lsCadImp & "ITEM  FECHA      CLIENTE                               MONTO           A           A              A              A           SALDO       COMISION " & Chr(10)
            lsCadImp = lsCadImp & "      PAGO                                             PAGADO       CAPITAL    INT. COMP.     INT. MORA        GASTOS       CAPITAL        S.K.   " & Chr(10)
             
        Case 138002 ' Listado de Pagos x Credito
            lsCadImp = lsCadImp & "ITEM  FECHA      FECHA      CLIENTE                               MONTO           A           A              A              A           SALDO       COMISION " & Chr(10)
            lsCadImp = lsCadImp & "      PAGO       JUD.                                             PAGADO       CAPITAL    INT. COMP.     INT. MORA        GASTOS       CAPITAL        S.K.   " & Chr(10)
          
        Case 138003 ' Listado de Pagos x Fecha
            lsCadImp = lsCadImp & "ITEM  CUENTA             FECHA       CLIENTE                               MONTO           A           A              A              A           SALDO       COMISION USER AG" & Chr(10)
            lsCadImp = lsCadImp & "                         PAGO                                              PAGADO       CAPITAL    INT. COMP.     INT. MORA        GASTOS       CAPITAL        S.K.   " & Chr(10)
           
        Case 138004 ' Listado de Pagos x Analista
            lsCadImp = lsCadImp & "ITEM  FECHA      CLIENTE                               MONTO           A           A              A              A           SALDO       COMISION " & Chr(10)
            lsCadImp = lsCadImp & "      PAGO                                             PAGADO       CAPITAL    INT. COMP.     INT. MORA        GASTOS       CAPITAL        S.K.   " & Chr(10)
            
        Case 138011 ' Listado de Creditos Nuevos
            lsCadImp = lsCadImp & "ITEM  CUENTA             FEC.VIGEN. FEC.ING.   LINEA CRED. CLIENTE                        ANA.          MONTO       CAPITAL     INT. COMP    INT. MORAT     GASTOS        TOTAL " & Chr(10)
 
        Case 138012 ' Listado de Creditos En Judicial
            lsCadImp = lsCadImp & "ITEM CUENTA             FEC.ING.   CLIENTE                        ANA.       PRESTAMO     SALDO CAP      INTERES         MORA       GASTOS      TOTAL    EXP" & Chr(10)
            
        Case 138014 ' Listado de Creditos x Analista
            lsCadImp = lsCadImp & "ITEM  CUENTA             F.ING.REC. CLIENTE                             PRESTAMO     SALDO CAP       INTERES          MORA        GASTOS       TOTAL   ABOGADO" & Chr(10)
            
            
        Case 138021 'Listado de Consolidado por Abogado
            lsCadImp = lsCadImp & "DESCRIPCION                    PRESTAMO         SALDO       INTERES          MORA        GASTOS         TOTAL " & Chr(10)
        
        Case 138023 ' Listado Segun Saldo de capital
            lsCadImp = lsCadImp & " ITEM CUENTA              CLIENTE                        FECHA      FECHA ING.  TIPO DE      MONTO       SALDO    ESTADO" & Chr(10)
            lsCadImp = lsCadImp & "                                                         PRESTAMO   COB. JUD.   MONEDA     PRESTAMO   CAPITAL S/." & Chr(10)
            
        Case 138024 'Listado de Expedientes por Abogado
            'lsCadImp = lsCadImp & "ITEM  CUENTA             LINEA CRED. CLIENTE                             PRESTAMO         SALDO       INTERES          MORA        GASTOS       TOTAL   ANA. SITUACION CASO" & Chr(10)
            lsCadImp = lsCadImp & "ITEM  CUENTA             F.ING.REC. CLIENTE                             PRESTAMO     SALDO CAP       INTERES          MORA        GASTOS       TOTAL   ANALISTA" & Chr(10)


         Case 150000
             
    End Select
    If pnCodReporte <> 150000 And pnCodReporte <> 138025 Then
        lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & Chr(10)
    End If
nRepoCabecera = lsCadImp
End Function

 
Public Function nRepo138001_ListadoPagosxAbogado(ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal psConAbogado As Byte, ByVal psCodAbogado As String) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As NColPImpre
        
Dim sCliente As String * 30
Dim lsCondiciones As String
Dim sCriterio As String
 
Dim sTempoAbogado As String
Dim sTempoMoneda As Integer
Dim sTempoTipoRecuperacion As Integer

Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim sOpeEst As String
'Totales por Abogados
Dim sSubTotal1 As Currency
Dim sSubTotal2 As Currency
Dim sSubTotal3 As Currency
Dim sSubTotal4 As Currency
Dim sSubTotal5 As Currency
Dim sSubTotal6 As Currency
Dim sSubTotal7 As Currency

'Totales por Moneda
Dim sTotal1 As Currency
Dim sTotal2 As Currency
Dim sTotal3 As Currency
Dim sTotal4 As Currency
Dim sTotal5 As Currency
Dim sTotal6 As Currency
Dim sTotal7 As Currency

'Totales por Tipo de Recuperacion
Dim sTotal11 As Currency
Dim sTotal21 As Currency
Dim sTotal31 As Currency
Dim sTotal41 As Currency
Dim sTotal51 As Currency
Dim sTotal61 As Currency
Dim sTotal71 As Currency
 
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

    
    sOpeEst = " (MOC.cOpeCod in(" & gColRecOpePasoARecup & ", " & gColRecOpePagJudSD & ", " & gColRecOpePagJudSDEfe & ", " & gColRecOpePagJudSDChq & ", " & gColRecOpePagJudCD & ", " & gColRecOpePagJudCDEfe & ", " & gColRecOpePagJudCDChq & ", " & gColRecOpePagCast & ", " & gColRecOpePagCastEfe & ", " & gColRecOpePagCastChq & ", " & gColRecOpePagJudSDEnOtCj & ", " & gColRecOpePagJudSDEnOtCjEfe & ", " & gColRecOpePagJudCDEnOtCj & ", " & gColRecOpePagJudCDEnOtCjEfe & ", " & gColRecOpePagJudCastEnOtCj & ", " & gColRecOpePagJudCastEnOtCjEfe & ")) AND "
    
    sCriterio = " AND (CONVERT(datetime, SUBSTRING(Mov.cMovNro, 1, 8), 103) Between '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "') "
        
    lsCondiciones = "Del " & pdFecIni & " al " & pdFecFin
 
    'Si se busca ABOGADO
    If psConAbogado = 1 Then
         
        sCriterio = sCriterio & " AND ((SELECT PRP.CPERSCOD FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON " & _
        " PRP.CPERSCOD = PERSO.CPERSCOD WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ")='" & Trim(psCodAbogado) & "') "
 
        lsCondiciones = lsCondiciones & " - Abogado: " & psCodAbogado & " : "
    End If

lsSQL = "SELECT MC.nMonto AS Pagado, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodCapital & "'), 0) AS Capital, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & "  mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodInteresCompensatorio & "'), 0) AS InteresCom, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & "  mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodInteresMoratorio & "'), 0) AS InteresMor, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & "  mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND (MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoCobranza & "' OR MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoComisionAbog & "' OR MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoAdmin & "')), 0) AS Gastos, " & _
        " isnull(MC.nSaldoCap,0) AS SaldoCapital, ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & "  mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoComisionAbog & "'), 0) AS ComisionAbog, Persona.cPersNombre, " & _
        " (SELECT PRP.CPERSCOD FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD " & _
        " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ") AS CODABOGADO, (SELECT PERSO.cPersNombre " & _
        " FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ") AS ABOGADO, " & _
        " CONVERT(varchar(10), CONVERT(datetime, SUBSTRING(Mov.cMovNro, 1, 8), 103), 103) AS Fecha, SUBSTRING(Colocaciones.cLineaCred, 5, 1) AS Moneda, Constante.cConsDescripcion AS sMoneda, " & _
        " nTipoRecuperacion=case when Producto.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "') THEN 1 ELSE 2 END, " & _
        " nDesTipoRecuperacion=case when Producto.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "') THEN 'JUDICIAL' ELSE 'CASTIGADO' END " & _
        " FROM ColocRecup INNER JOIN Colocaciones ON ColocRecup.cCtaCod = Colocaciones.cCtaCod INNER JOIN MovCol MC INNER JOIN ProductoPersona PP ON MC.cCtaCod = PP.cCtaCod INNER JOIN " & _
        " Mov ON MC.nMovNro = Mov.nMovNro ON ColocRecup.cCtaCod = PP.cCtaCod INNER JOIN Persona ON PP.cPersCod = Persona.cPersCod INNER JOIN Producto ON " & _
        " Colocaciones.cCtaCod = Producto.cCtaCod AND PP.cCtaCod = Producto.cCtaCod INNER JOIN " & _
        " Constante ON SUBSTRING(Colocaciones.cLineaCred, 5, 1) = Constante.nConsValor " & _
        " INNER Join  OpeTpo Ope on ope.cOpeCod = MC.cOpeCod " & _
        " WHERE (PP.nPrdPersRelac = 20) AND (Mov.nMovFlag=" & gMovFlagVigente & ") " & _
        " AND (Producto.nPrdEstado IN ('" & gColocEstRecVigJud & "', '" & gColocEstRecVigCast & "', '" & gColocEstRecCanJud & "', '" & gColocEstRecCanCast & "')) AND (Constante.nConsCod = '1011') " & _
        sCriterio
        
        'nTiporecuperacion=1 si es VigenteJudicial(2201) o CanceladoJudicial(2203)
        'nTiporecuperacion=2 si es VigenteCastigado(2201) o CanceladoCastigado(2203)

lsSQL = lsSQL & " and Ope.cOpeCod like '130%' and Ope.nOpeNiv<>1"
lsSQL = lsSQL & " Order By (SELECT PERSO.cPersNombre FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON " & _
        " PRP.cPersCod = PERSO.cPersCod WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & "), " & _
        " SUBSTRING(Colocaciones.cLineaCred, 5, 1), case when Producto.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "') THEN 1 ELSE 2 END, SUBSTRING(Mov.cMovNro, 1, 8) "
 
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
     
        If psConAbogado = 1 Then
            lsCondiciones = lsCondiciones & "" & lrDataRep!Abogado
        End If
        
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE PAGOS DE CREDITOS POR ABOGADO", lsCondiciones, lnPage, 145, "", 138001)

        lnIndice = 0:  lnLineas = 7
        
        sTempoAbogado = "" & lrDataRep!CODABOGADO
        sTempoMoneda = lrDataRep!Moneda
        sTempoTipoRecuperacion = lrDataRep!nTipoRecuperacion
        
        'Abogado
        lsCadImp = lsCadImp & lsNegritaOn & "ABOGADO: " & lrDataRep!Abogado & " : " & lrDataRep!CODABOGADO & lsNegritaOff & Chr(10) & Chr(10)
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & Chr(10) & Chr(10)
        'Tipo de Recuperacion
        lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
        
        lnLineas = lnLineas + 6
        lnIndice = lnIndice + 6
          
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                
                'Totalizar por Abogado
                If sTempoAbogado <> !CODABOGADO Then
                
                    lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & ImpreFormat(sTotal11, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal41, 10, 2) & Space(1) & ImpreFormat(sTotal51, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal61, 10, 2) & Space(1) & ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                
                    lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(23) & ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal6, 10, 2) & Space(1) & ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                    
                    lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ABOGADO " & Space(30) & ImpreFormat(sSubTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sSubTotal2, 10, 2) & Space(1) & ImpreFormat(sSubTotal3, 10, 2) & _
                                Space(1) & ImpreFormat(sSubTotal4, 10, 2) & Space(1) & ImpreFormat(sSubTotal5, 10, 2) & _
                                Space(1) & ImpreFormat(sSubTotal6, 10, 2) & Space(1) & ImpreFormat(sSubTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                     
                    'Abogado
                    lsCadImp = lsCadImp & lsNegritaOn & "ABOGADO: " & lrDataRep!Abogado & " : " & lrDataRep!CODABOGADO & lsNegritaOff & Chr(10) & Chr(10)
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & Chr(10) & Chr(10)
                    'Tipo de Recuperacion
                    lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & Chr(10)
                    lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                    
                    sSubTotal1 = !Pagado
                    sSubTotal2 = !Capital
                    sSubTotal3 = !InteresCom
                    sSubTotal4 = !InteresMor
                    sSubTotal5 = !Gastos
                    sSubTotal6 = !SaldoCapital
                    sSubTotal7 = !ComisionAbog
                    
                    sTotal1 = !Pagado
                    sTotal2 = !Capital
                    sTotal3 = !InteresCom
                    sTotal4 = !InteresMor
                    sTotal5 = !Gastos
                    sTotal6 = !SaldoCapital
                    sTotal7 = !ComisionAbog
                        
                    sTotal11 = !Pagado
                    sTotal21 = !Capital
                    sTotal31 = !InteresCom
                    sTotal41 = !InteresMor
                    sTotal51 = !Gastos
                    sTotal61 = !SaldoCapital
                    sTotal71 = !ComisionAbog
                    
                    nItem = 1
                    
                    sTempoAbogado = "" & lrDataRep!CODABOGADO
                    sTempoMoneda = lrDataRep!Moneda
                    sTempoTipoRecuperacion = lrDataRep!nTipoRecuperacion
                      
                    lnLineas = lnLineas + 15
                    lnIndice = lnIndice + 15
                     
                Else
                    sSubTotal1 = sSubTotal1 + !Pagado
                    sSubTotal2 = sSubTotal2 + !Capital
                    sSubTotal3 = sSubTotal3 + !InteresCom
                    sSubTotal4 = sSubTotal4 + !InteresMor
                    sSubTotal5 = sSubTotal5 + !Gastos
                    sSubTotal6 = sSubTotal6 + !SaldoCapital
                    sSubTotal7 = sSubTotal7 + !ComisionAbog
                    
                    If sTempoMoneda <> !Moneda Then
                
                        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & ImpreFormat(sTotal11, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal41, 10, 2) & Space(1) & ImpreFormat(sTotal51, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal61, 10, 2) & Space(1) & ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                    
                        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(23) & ImpreFormat(sTotal1, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal6, 10, 2) & Space(1) & ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                        
                        'Moneda
                        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & Chr(10) & Chr(10)
                        'Tipo de Recuperacion
                        lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & Chr(10)
                        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                        
                        sTotal1 = !Pagado
                        sTotal2 = !Capital
                        sTotal3 = !InteresCom
                        sTotal4 = !InteresMor
                        sTotal5 = !Gastos
                        sTotal6 = !SaldoCapital
                        sTotal7 = !ComisionAbog
                            
                        sTotal11 = !Pagado
                        sTotal21 = !Capital
                        sTotal31 = !InteresCom
                        sTotal41 = !InteresMor
                        sTotal51 = !Gastos
                        sTotal61 = !SaldoCapital
                        sTotal71 = !ComisionAbog
                        
                        nItem = 1
                        
                        sTempoMoneda = lrDataRep!Moneda
                        sTempoTipoRecuperacion = lrDataRep!nTipoRecuperacion
                    
                        lnLineas = lnLineas + 10
                        lnIndice = lnIndice + 10
                    Else
                        sTotal1 = sTotal1 + !Pagado
                        sTotal2 = sTotal2 + !Capital
                        sTotal3 = sTotal3 + !InteresCom
                        sTotal4 = sTotal4 + !InteresMor
                        sTotal5 = sTotal5 + !Gastos
                        sTotal6 = sTotal6 + !SaldoCapital
                        sTotal7 = sTotal7 + !ComisionAbog
                        
                        If sTempoTipoRecuperacion <> !nTipoRecuperacion Then
                    
                            lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                            lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & ImpreFormat(sTotal11, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal41, 10, 2) & Space(1) & ImpreFormat(sTotal51, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal61, 10, 2) & Space(1) & ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                    
                            'Tipo de Recuperacion
                            lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & Chr(10)
                            lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                            
                            sTotal11 = !Pagado
                            sTotal21 = !Capital
                            sTotal31 = !InteresCom
                            sTotal41 = !InteresMor
                            sTotal51 = !Gastos
                            sTotal61 = !SaldoCapital
                            sTotal71 = !ComisionAbog
                                 
                            nItem = 1
                            
                            sTempoTipoRecuperacion = lrDataRep!nTipoRecuperacion
                        
                            lnLineas = lnLineas + 5
                            lnIndice = lnIndice + 5
                        Else
                            sTotal11 = sTotal11 + !Pagado
                            sTotal21 = sTotal21 + !Capital
                            sTotal31 = sTotal31 + !InteresCom
                            sTotal41 = sTotal41 + !InteresMor
                            sTotal51 = sTotal51 + !Gastos
                            sTotal61 = sTotal61 + !SaldoCapital
                            sTotal71 = sTotal71 + !ComisionAbog
                        End If
                    End If
                End If
                 
                sCliente = !cpersNombre
                 
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(nItem, 4, 0) & Space(1) & !Fecha & Space(1) & _
                           sCliente & Space(1) & ImpreFormat(!Pagado, 10, 2) & Space(1) & ImpreFormat(!Capital, 10, 2) & _
                           Space(1) & ImpreFormat(!InteresCom, 10, 2) & Space(1) & ImpreFormat(!InteresMor, 10, 2) & _
                           Space(1) & ImpreFormat(!Gastos, 10, 2) & Space(1) & ImpreFormat(!SaldoCapital, 10, 2) & _
                           Space(1) & ImpreFormat(!ComisionAbog, 10, 2) & Chr(10)
                 
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE PAGOS DE CREDITOS POR ABOGADO", lsCondiciones, lnPage, 145, "", 138001)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & ImpreFormat(sTotal11, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal41, 10, 2) & Space(1) & ImpreFormat(sTotal51, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal61, 10, 2) & Space(1) & ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
    
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(23) & ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal6, 10, 2) & Space(1) & ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
        
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ABOGADO " & Space(30) & ImpreFormat(sSubTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal2, 10, 2) & Space(1) & ImpreFormat(sSubTotal3, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal4, 10, 2) & Space(1) & ImpreFormat(sSubTotal5, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal6, 10, 2) & Space(1) & ImpreFormat(sSubTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
        
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
    nRepo138001_ListadoPagosxAbogado = lsCadBuffer
End Function


Public Function nRepo138002_ListadoPagosdeCreditos(ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal psEstados As String, ByVal pnMoneda As Byte) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As NColPImpre
        
Dim sCliente As String * 30
Dim lsCondiciones As String
Dim sCriterio As String
Dim sOpeEst As String

Dim sTempoMoneda As Integer
Dim sTempoAgencia  As String
Dim sTempoLineaCredito As String

Dim lsNegritaOn As String
Dim lsNegritaOff As String

'Totales por Moneda
Dim sSubTotal1 As Currency
Dim sSubTotal2 As Currency
Dim sSubTotal3 As Currency
Dim sSubTotal4 As Currency
Dim sSubTotal5 As Currency
Dim sSubTotal6 As Currency
Dim sSubTotal7 As Currency

'Totales por Agencia
Dim sTotal1 As Currency
Dim sTotal2 As Currency
Dim sTotal3 As Currency
Dim sTotal4 As Currency
Dim sTotal5 As Currency
Dim sTotal6 As Currency
Dim sTotal7 As Currency

'Totales por Linea de Credito
Dim sTotal11 As Currency
Dim sTotal21 As Currency
Dim sTotal31 As Currency
Dim sTotal41 As Currency
Dim sTotal51 As Currency
Dim sTotal61 As Currency
Dim sTotal71 As Currency
 
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)
     
    lsCondiciones = "Del " & pdFecIni & " al " & pdFecFin
    
    If psEstados = "'" & gColocEstRecVigJud & "', '" & gColocEstRecCanJud & "'" Then
        lsCondiciones = lsCondiciones & " - CREDITOS JUDICIALES"
    Else
        lsCondiciones = lsCondiciones & " - CREDITOS CASTIGADOS"
    End If
     
    'Busco por estados en la tabla Producto
    sCriterio = " AND (P.nPrdEstado IN(" & psEstados & ")) "
    
    'Si se busca por MONEDA
    If pnMoneda <> 3 Then
        sCriterio = sCriterio & " AND (SUBSTRING(C.cLineaCred, 5, 1)=" & pnMoneda & ")"
       
    End If
     
    sCriterio = sCriterio & " AND (CONVERT(datetime, SUBSTRING(M.cMovNro, 1, 8), 103) Between '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "') "
         
    sOpeEst = " (MOC.cOpeCod in(" & gColRecOpePasoARecup & ", " & gColRecOpePagJudSD & ", " & gColRecOpePagJudSDEfe & ", " & gColRecOpePagJudSDChq & ", " & gColRecOpePagJudCD & ", " & gColRecOpePagJudCDEfe & ", " & gColRecOpePagJudCDChq & ", " & gColRecOpePagCast & ", " & gColRecOpePagCastEfe & ", " & gColRecOpePagCastChq & ", " & gColRecOpePagJudSDEnOtCj & ", " & gColRecOpePagJudSDEnOtCjEfe & ", " & gColRecOpePagJudCDEnOtCj & ", " & gColRecOpePagJudCDEnOtCjEfe & ", " & gColRecOpePagJudCastEnOtCj & ", " & gColRecOpePagJudCastEnOtCjEfe & ")) AND "
         
lsSQL = "SELECT DISTINCT MC.nMonto AS nPagado, " & _
        " ISNULL ((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO = MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO = MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro = MC.nMovNro AND MOC.cCtaCod = MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodCapital & "'), 0) AS nCapital, " & _
        " ISNULL ((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO = MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO = MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro = MC.nMovNro AND MOC.cCtaCod = MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodInteresCompensatorio & "'), 0) AS nInteresCom, " & _
        " ISNULL ((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO = MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO = MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro = MC.nMovNro AND MOC.cCtaCod = MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodInteresMoratorio & "'), 0) AS nInteresMor, " & _
        " ISNULL ((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO = MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO = MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro = MC.nMovNro AND MOC.cCtaCod = MC.cCtaCod AND (MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoCobranza & "' OR MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoComisionAbog & "' OR MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoAdmin & "')), 0) AS nGastos, ISNULL(MC.nSaldoCap, 0) AS nSaldoCapital, " & _
        " ISNULL ((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO = MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO = MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro = MC.nMovNro AND MOC.cCtaCod = MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoComisionAbog & "'), 0) AS nComisionAbog, " & _
        " ISNULL(MC.nSaldoCap, 0) AS nSaldoCapital, Persona.cPersNombre, CONVERT(varchar(10), CONVERT(datetime, SUBSTRING(M.cMovNro, 1, 8), 103), 103) AS dFechaPago, SUBSTRING(C.cLineaCred, 5, 1) AS nMoneda, " & _
        " K.cConsDescripcion AS cMoneda, C.cLineaCred AS cLineaCredito, CLC.cDescripcion AS cDesLineaCredito, C.cCtaCod, SUBSTRING(C.cCtaCod, 4, 2) AS cCodAgencia, A.cAgeDescripcion, CR.dIngRecup AS dFechaJudicial " & _
        " FROM Producto P INNER JOIN ColocRecup CR INNER JOIN Colocaciones C ON CR.cCtaCod = C.cCtaCod ON P.cCtaCod = C.cCtaCod INNER JOIN Persona INNER JOIN MovCol MC INNER JOIN " & _
        " ProductoPersona PP ON MC.cCtaCod = PP.cCtaCod ON Persona.cPersCod = PP.cPersCod ON P.cCtaCod = PP.cCtaCod INNER JOIN Constante K ON SUBSTRING(C.cLineaCred, 5, 1) = K.nConsValor INNER JOIN " & _
        " ColocLineaCredito CLC ON C.cLineaCred = CLC.cLineaCred INNER JOIN Agencias A ON SUBSTRING(C.cCtaCod, 4, 2) = A.cAgeCod " & _
         " INNER JOIN MOV M ON MC.nMovNro=M.nMovNro " & _
         " INNER Join  OpeTpo Ope on ope.cOpeCod = MC.cOpeCod " & _
         " WHERE (PP.nPrdPersRelac=" & gColRelPersTitular & ") AND  (K.nConsCod='" & gMoneda & "') " & _
         sCriterio & _
         " and Ope.cOpeCod like '130%' and Ope.nOpeNiv<>1 " & _
         " ORDER BY SUBSTRING(C.cLineaCred,5,1), A.cAgeDescripcion, C.cLineaCred "
          
        'nTiporecuperacion=1 si es VigenteJudicial(2201) o CanceladoJudicial(2203)
        'nTiporecuperacion=2 si es VigenteCastigado(2201) o CanceladoCastigado(2203)
         
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
     
        If pnMoneda <> 3 Then
            lsCondiciones = lsCondiciones & " - Moneda: " & lrDataRep!cMoneda
        End If
        
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE PAGOS DE CREDITOS", lsCondiciones, lnPage, 156, "", 138002)

        lnIndice = 0:  lnLineas = 7
        
        sTempoAgencia = "" & lrDataRep!cCodAgencia
        sTempoMoneda = lrDataRep!nMoneda
        sTempoLineaCredito = "" & lrDataRep!cLineaCredito
        
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cMoneda & lsNegritaOff & Chr(10) & Chr(10)
        'Agencia
        lsCadImp = lsCadImp & lsNegritaOn & "**  AGENCIA: " & lrDataRep!cAgeDescripcion & " : " & lrDataRep!cCodAgencia & lsNegritaOff & Chr(10) & Chr(10)
        'Linea de Credito
        lsCadImp = lsCadImp & lsNegritaOn & "*** LINEA DE CREDITO: " & lrDataRep!cDesLineaCredito & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(156, "-") & Chr(10)
        
        lnLineas = lnLineas + 6
        lnIndice = lnIndice + 6
          
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                
                'Totalizar por Moneda
                
                If sTempoMoneda <> !nMoneda Then
                
                    lsCadImp = lsCadImp & String(156, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X LINEA DE CRED.  " & Space(23) & ImpreFormat(sTotal11, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal41, 10, 2) & Space(1) & ImpreFormat(sTotal51, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal61, 10, 2) & Space(1) & ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
    
                    lsCadImp = lsCadImp & String(156, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X AGENCIA         " & Space(23) & ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal6, 10, 2) & Space(1) & ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
        
                    lsCadImp = lsCadImp & String(156, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X MONEDA   " & Space(30) & ImpreFormat(sSubTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal2, 10, 2) & Space(1) & ImpreFormat(sSubTotal3, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal4, 10, 2) & Space(1) & ImpreFormat(sSubTotal5, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal6, 10, 2) & Space(1) & ImpreFormat(sSubTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                     
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cMoneda & lsNegritaOff & Chr(10) & Chr(10)
                    'Agencia
                    lsCadImp = lsCadImp & lsNegritaOn & "**  AGENCIA: " & lrDataRep!cAgeDescripcion & " : " & lrDataRep!cCodAgencia & lsNegritaOff & Chr(10) & Chr(10)
                    'Linea de Credito
                    lsCadImp = lsCadImp & lsNegritaOn & "*** LINEA DE CREDITO: " & lrDataRep!cDesLineaCredito & lsNegritaOff & Chr(10)
                    lsCadImp = lsCadImp & String(156, "-") & Chr(10)
                     
                    sSubTotal1 = !nPagado
                    sSubTotal2 = !nCapital
                    sSubTotal3 = !nInteresCom
                    sSubTotal4 = !nInteresMor
                    sSubTotal5 = !nGastos
                    sSubTotal6 = !nSaldoCapital
                    sSubTotal7 = !nComisionAbog
                    
                    sTotal1 = !nPagado
                    sTotal2 = !nCapital
                    sTotal3 = !nInteresCom
                    sTotal4 = !nInteresMor
                    sTotal5 = !nGastos
                    sTotal6 = !nSaldoCapital
                    sTotal7 = !nComisionAbog
                        
                    sTotal11 = !nPagado
                    sTotal21 = !nCapital
                    sTotal31 = !nInteresCom
                    sTotal41 = !nInteresMor
                    sTotal51 = !nGastos
                    sTotal61 = !nSaldoCapital
                    sTotal71 = !nComisionAbog
                    
                    nItem = 1
                    
                    sTempoMoneda = lrDataRep!nMoneda
                    sTempoAgencia = "" & lrDataRep!cCodAgencia
                    sTempoLineaCredito = "" & lrDataRep!cLineaCredito
                      
                    lnLineas = lnLineas + 15
                    lnIndice = lnIndice + 15
                     
                Else
                    sSubTotal1 = sSubTotal1 + !nPagado
                    sSubTotal2 = sSubTotal2 + !nCapital
                    sSubTotal3 = sSubTotal3 + !nInteresCom
                    sSubTotal4 = sSubTotal4 + !nInteresMor
                    sSubTotal5 = sSubTotal5 + !nGastos
                    sSubTotal6 = sSubTotal6 + !nSaldoCapital
                    sSubTotal7 = sSubTotal7 + !nComisionAbog
                    
                    If sTempoAgencia <> !cCodAgencia Then
                
                        lsCadImp = lsCadImp & String(156, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X LINEA DE CRED.  " & Space(23) & ImpreFormat(sTotal11, 10, 2) & _
                        Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 10, 2) & _
                        Space(1) & ImpreFormat(sTotal41, 10, 2) & Space(1) & ImpreFormat(sTotal51, 10, 2) & _
                        Space(1) & ImpreFormat(sTotal61, 10, 2) & Space(1) & ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
        
                        lsCadImp = lsCadImp & String(156, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X AGENCIA         " & Space(23) & ImpreFormat(sTotal1, 10, 2) & _
                        Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                        Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                        Space(1) & ImpreFormat(sTotal6, 10, 2) & Space(1) & ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                            
                        'Agencia
                        lsCadImp = lsCadImp & lsNegritaOn & "**  AGENCIA: " & lrDataRep!cAgeDescripcion & " : " & lrDataRep!cCodAgencia & lsNegritaOff & Chr(10) & Chr(10)
                        'Linea de Credito
                        lsCadImp = lsCadImp & lsNegritaOn & "*** LINEA DE CREDITO: " & lrDataRep!cDesLineaCredito & lsNegritaOff & Chr(10)
                        lsCadImp = lsCadImp & String(156, "-") & Chr(10)
                        
                        sTotal1 = !nPagado
                        sTotal2 = !nCapital
                        sTotal3 = !nInteresCom
                        sTotal4 = !nInteresMor
                        sTotal5 = !nGastos
                        sTotal6 = !nSaldoCapital
                        sTotal7 = !nComisionAbog
                            
                        sTotal11 = !nPagado
                        sTotal21 = !nCapital
                        sTotal31 = !nInteresCom
                        sTotal41 = !nInteresMor
                        sTotal51 = !nGastos
                        sTotal61 = !nSaldoCapital
                        sTotal71 = !nComisionAbog
                        
                        nItem = 1
                        
                        sTempoAgencia = "" & lrDataRep!cCodAgencia
                        sTempoLineaCredito = "" & lrDataRep!cLineaCredito
                     
                        lnLineas = lnLineas + 10
                        lnIndice = lnIndice + 10
                    Else
                        sTotal1 = sTotal1 + !nPagado
                        sTotal2 = sTotal2 + !nCapital
                        sTotal3 = sTotal3 + !nInteresCom
                        sTotal4 = sTotal4 + !nInteresMor
                        sTotal5 = sTotal5 + !nGastos
                        sTotal6 = sTotal6 + !nSaldoCapital
                        sTotal7 = sTotal7 + !nComisionAbog
                        
                        If sTempoLineaCredito <> !cLineaCredito Then
                    
                            lsCadImp = lsCadImp & String(156, "-") & Chr(10)
                            lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X LINEA DE CRED.  " & Space(23) & ImpreFormat(sTotal11, 10, 2) & _
                            Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 10, 2) & _
                            Space(1) & ImpreFormat(sTotal41, 10, 2) & Space(1) & ImpreFormat(sTotal51, 10, 2) & _
                            Space(1) & ImpreFormat(sTotal61, 10, 2) & Space(1) & ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                     
                            'Linea de Credito
                            lsCadImp = lsCadImp & lsNegritaOn & "*** LINEA DE CREDITO: " & lrDataRep!cDesLineaCredito & lsNegritaOff & Chr(10)
                            lsCadImp = lsCadImp & String(156, "-") & Chr(10)
                            
                            sTotal11 = !nPagado
                            sTotal21 = !nCapital
                            sTotal31 = !nInteresCom
                            sTotal41 = !nInteresMor
                            sTotal51 = !nGastos
                            sTotal61 = !nSaldoCapital
                            sTotal71 = !nComisionAbog
                                 
                            nItem = 1
                            
                            sTempoLineaCredito = "" & lrDataRep!cLineaCredito
                        
                            lnLineas = lnLineas + 5
                            lnIndice = lnIndice + 5
                        Else
                            sTotal11 = sTotal11 + !nPagado
                            sTotal21 = sTotal21 + !nCapital
                            sTotal31 = sTotal31 + !nInteresCom
                            sTotal41 = sTotal41 + !nInteresMor
                            sTotal51 = sTotal51 + !nGastos
                            sTotal61 = sTotal61 + !nSaldoCapital
                            sTotal71 = sTotal71 + !nComisionAbog
                        End If
                    End If
                End If
                 
                sCliente = !cpersNombre
                 
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(nItem, 4, 0) & Space(1) & !dFechaPago & Space(1) & _
                           Format(!dFechaJudicial, "dd/MM/YYYY") & _
                           Space(1) & sCliente & Space(1) & ImpreFormat(!nPagado, 10, 2) & Space(1) & ImpreFormat(!nCapital, 10, 2) & _
                           Space(1) & ImpreFormat(!nInteresCom, 10, 2) & Space(1) & ImpreFormat(!nInteresMor, 10, 2) & _
                           Space(1) & ImpreFormat(!nGastos, 10, 2) & Space(1) & ImpreFormat(!nSaldoCapital, 10, 2) & _
                           Space(1) & ImpreFormat(!nComisionAbog, 10, 2) & Chr(10)
                 
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE PAGOS DE CREDITOS", lsCondiciones, lnPage, 156, "", 138002)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(156, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X LINEA DE CRED.  " & Space(23) & ImpreFormat(sTotal11, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal41, 10, 2) & Space(1) & ImpreFormat(sTotal51, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal61, 10, 2) & Space(1) & ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
    
        lsCadImp = lsCadImp & String(156, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X AGENCIA         " & Space(23) & ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal6, 10, 2) & Space(1) & ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
        
        lsCadImp = lsCadImp & String(156, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X MONEDA   " & Space(30) & ImpreFormat(sSubTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal2, 10, 2) & Space(1) & ImpreFormat(sSubTotal3, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal4, 10, 2) & Space(1) & ImpreFormat(sSubTotal5, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal6, 10, 2) & Space(1) & ImpreFormat(sSubTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
      
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
    nRepo138002_ListadoPagosdeCreditos = lsCadBuffer
End Function


Public Function nRepo138003_ListadoPagosxFechas(ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal psEstados As String, ByVal pnMoneda As Byte) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As NColPImpre
        
Dim sCliente As String * 30
Dim lsCondiciones As String
Dim sCriterio As String
  
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim sOpe As String

Dim sTempoMoneda As Integer

'Totales x Moneda
Dim sTotal1 As Currency
Dim sTotal2 As Currency
Dim sTotal3 As Currency
Dim sTotal4 As Currency
Dim sTotal5 As Currency
Dim sTotal6 As Currency
Dim sTotal7 As Currency
  
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)
     
    lsCondiciones = "Del " & pdFecIni & " al " & pdFecFin
    
    If psEstados = "'" & gColocEstRecVigJud & "', '" & gColocEstRecCanJud & "'" Then
        lsCondiciones = lsCondiciones & " - CREDITOS JUDICIALES"
    Else
        lsCondiciones = lsCondiciones & " - CREDITOS CASTIGADOS"
    End If
     
    'Busco por estados en la tabla Producto
    sCriterio = " AND (P.nPrdEstado IN(" & psEstados & ")) "
    
    'Si se busca por MONEDA
    If pnMoneda <> 3 Then
        sCriterio = sCriterio & " AND (SUBSTRING(C.cLineaCred, 5, 1)=" & pnMoneda & ")"
       
    End If
     
    sOpe = " (MOC.cOpeCod in(" & gColRecOpePasoARecup & ", " & gColRecOpePagJudSD & ", " & gColRecOpePagJudSDEfe & ", " & gColRecOpePagJudSDChq & ", " & gColRecOpePagJudCD & ", " & gColRecOpePagJudCDEfe & ", " & gColRecOpePagJudCDChq & ", " & gColRecOpePagCast & ", " & gColRecOpePagCastEfe & ", " & gColRecOpePagCastChq & ", " & gColRecOpePagJudSDEnOtCj & ", " & gColRecOpePagJudSDEnOtCjEfe & ", " & gColRecOpePagJudCDEnOtCj & ", " & gColRecOpePagJudCDEnOtCjEfe & ", " & gColRecOpePagJudCastEnOtCj & ", " & gColRecOpePagJudCastEnOtCjEfe & ")) AND "
     
    sCriterio = sCriterio & " AND (CONVERT(datetime, SUBSTRING(M.cMovNro, 1, 8), 103) Between '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "') "
           
    lsSQL = "SELECT MC.nMonto AS nPagado, right(M.cmovnro,4) as cUsuarioMov, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpe & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodCapital & "'), 0) AS nCapital, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpe & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodInteresCompensatorio & "'), 0) AS nInteresCom, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpe & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodInteresMoratorio & "'), 0) AS nInteresMor, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpe & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND (MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoCobranza & "' OR MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoComisionAbog & "' OR MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoAdmin & "')), 0) AS nGastos, " & _
        " ISNULL(MC.nSaldoCap,0) AS nSaldoCapital, ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpe & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoComisionAbog & "'), 0) AS nComisionAbog, Persona.cPersNombre, " & _
        " CONVERT(varchar(10), CONVERT(datetime, SUBSTRING(M.cMovNro, 1, 8), 103), 103) AS dFechaPago, SUBSTRING(C.cLineaCred, 5, 1) AS nMoneda, " & _
        " K.cConsDescripcion AS cMoneda, C.cLineaCred AS cLineaCredito, CLC.cDescripcion AS cDesLineaCredito, C.cCtaCod, SUBSTRING(C.cCtaCod, 4, 2) " & _
        " AS cCodAgencia, A.cAgeDescripcion, (SELECT RHH.CUSER FROM RRHH RHH INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD " & _
        " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS ANALISTA FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod INNER JOIN Persona INNER JOIN MovCol MC INNER JOIN " & _
        " ProductoPersona PP ON MC.cCtaCod = PP.cCtaCod INNER JOIN Mov M ON MC.nMovNro = M.nMovNro ON Persona.cPersCod = PP.cPersCod ON P.cCtaCod = PP.cCtaCod INNER JOIN " & _
        " Constante K ON SUBSTRING(C.cLineaCred, 5, 1) = K.nConsValor INNER JOIN ColocLineaCredito CLC ON C.cLineaCred = CLC.cLineaCred INNER JOIN " & _
        " Agencias A ON SUBSTRING(C.cCtaCod, 4, 2) = A.cAgeCod " & _
        " INNER Join  OpeTpo Ope on ope.cOpeCod = MC.cOpeCod " & _
        " WHERE  (PP.nPrdPersRelac =" & gColRelPersTitular & ") AND (M.nMovFlag=" & gMovFlagVigente & ") AND (K.nConsCod='" & gMoneda & "') " & _
         sCriterio & _
         " and Ope.cOpeCod like '130%' and Ope.nOpeNiv<>1 " & _
         " ORDER BY SUBSTRING(C.cLineaCred, 5, 1), SUBSTRING(M.cMovNro, 1, 8), C.cCtaCod "
 
        'nTiporecuperacion=1 si es VigenteJudicial(2201) o CanceladoJudicial(2203)
        'nTiporecuperacion=2 si es VigenteCastigado(2201) o CanceladoCastigado(2203)
         
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
     
        If pnMoneda <> 3 Then
            lsCondiciones = lsCondiciones & " - Moneda: " & lrDataRep!cMoneda
        End If
        
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE PAGOS POR FECHAS", lsCondiciones, lnPage, 173, "", 138003)

        lnIndice = 0:  lnLineas = 7
           
        sTempoMoneda = lrDataRep!nMoneda
        
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cMoneda & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(173, "-") & Chr(10)
        lnLineas = lnLineas + 2
        lnIndice = lnIndice + 2
        
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                 
                If sTempoMoneda <> !nMoneda Then
                
                    lsCadImp = lsCadImp & String(173, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X MONEDA          " & Space(32) & ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal6, 10, 2) & Space(1) & ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
 
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cMoneda & lsNegritaOff & Chr(10)
                    lsCadImp = lsCadImp & String(173, "-") & Chr(10)
                    
                    nItem = 1
                    
                    sTempoMoneda = lrDataRep!nMoneda
                    
                    lnLineas = lnLineas + 5
                    lnIndice = lnIndice + 5
                    
                    sTotal1 = !nPagado
                    sTotal2 = !nCapital
                    sTotal3 = !nInteresCom
                    sTotal4 = !nInteresMor
                    sTotal5 = !nGastos
                    sTotal6 = !nSaldoCapital
                    sTotal7 = !nComisionAbog
                Else
                    sTotal1 = sTotal1 + !nPagado
                    sTotal2 = sTotal2 + !nCapital
                    sTotal3 = sTotal3 + !nInteresCom
                    sTotal4 = sTotal4 + !nInteresMor
                    sTotal5 = sTotal5 + !nGastos
                    sTotal6 = sTotal6 + !nSaldoCapital
                    sTotal7 = sTotal7 + !nComisionAbog
                End If
                
                sCliente = !cpersNombre
                 
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(nItem, 4, 0) & Space(1) & !cCtaCod & Space(1) & _
                           !dFechaPago & Space(1) & _
                           Space(1) & sCliente & Space(1) & ImpreFormat(!nPagado, 10, 2) & Space(1) & ImpreFormat(!nCapital, 10, 2) & _
                           Space(1) & ImpreFormat(!nInteresCom, 10, 2) & Space(1) & ImpreFormat(!nInteresMor, 10, 2) & _
                           Space(1) & ImpreFormat(!nGastos, 10, 2) & Space(1) & ImpreFormat(!nSaldoCapital, 10, 2) & _
                           Space(1) & ImpreFormat(!nComisionAbog, 10, 2) & Space(1) & !cUsuarioMov & Space(1) & !cCodAgencia & Chr(10)
                 
                 
                '!analista
                 
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE PAGOS POR FECHAS", lsCondiciones, lnPage, 173, "", 138003)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(173, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X MONEDA          " & Space(32) & ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal6, 10, 2) & Space(1) & ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
  
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
    nRepo138003_ListadoPagosxFechas = lsCadBuffer
End Function
 
 
Public Function nRepo138004_ListadoPagosxAnalista(ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal psConAnalista As Byte, ByVal psCodAnalista As String) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As NColPImpre
        
Dim sCliente As String * 30
Dim lsCondiciones As String
Dim sCriterio As String
Dim sOpeEst As String

Dim sTempoAnalista As String
Dim sTempoMoneda As Integer
Dim sTempoTipoRecuperacion As Integer

Dim lsNegritaOn As String
Dim lsNegritaOff As String

'Totales por Analistas
Dim sSubTotal1 As Currency
Dim sSubTotal2 As Currency
Dim sSubTotal3 As Currency
Dim sSubTotal4 As Currency
Dim sSubTotal5 As Currency
Dim sSubTotal6 As Currency
Dim sSubTotal7 As Currency

'Totales por Moneda
Dim sTotal1 As Currency
Dim sTotal2 As Currency
Dim sTotal3 As Currency
Dim sTotal4 As Currency
Dim sTotal5 As Currency
Dim sTotal6 As Currency
Dim sTotal7 As Currency

'Totales por Tipo de Recuperacion
Dim sTotal11 As Currency
Dim sTotal21 As Currency
Dim sTotal31 As Currency
Dim sTotal41 As Currency
Dim sTotal51 As Currency
Dim sTotal61 As Currency
Dim sTotal71 As Currency
 
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

    
    sOpeEst = " (MOC.cOpeCod in(" & gColRecOpePasoARecup & ", " & gColRecOpePagJudSD & ", " & gColRecOpePagJudSDEfe & ", " & gColRecOpePagJudSDChq & ", " & gColRecOpePagJudCD & ", " & gColRecOpePagJudCDEfe & ", " & gColRecOpePagJudCDChq & ", " & gColRecOpePagCast & ", " & gColRecOpePagCastEfe & ", " & gColRecOpePagCastChq & ", " & gColRecOpePagJudSDEnOtCj & ", " & gColRecOpePagJudSDEnOtCjEfe & ", " & gColRecOpePagJudCDEnOtCj & ", " & gColRecOpePagJudCDEnOtCjEfe & ", " & gColRecOpePagJudCastEnOtCj & ", " & gColRecOpePagJudCastEnOtCjEfe & ")) AND "
        
    sCriterio = " AND (CONVERT(datetime, SUBSTRING(Mov.cMovNro, 1, 8), 103) Between '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "') "
        
    lsCondiciones = "Del " & pdFecIni & " al " & pdFecFin
 
    'Si se busca analista
    If psConAnalista = 1 Then
        
        sCriterio = " AND ((SELECT RHH.CUSER FROM RRHH RHH INNER JOIN " & _
                    " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD " & _
                    " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ")='" & Trim(psCodAnalista) & "')"
        lsCondiciones = lsCondiciones & " Analista: " & psCodAnalista & " : "

    End If
         
lsSQL = "SELECT  MC.nMonto AS Pagado, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodCapital & "'), 0) AS Capital, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodInteresCompensatorio & "'), 0) AS InteresCom, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodInteresMoratorio & "'), 0) AS InteresMor, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND (MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoCobranza & "' OR MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoComisionAbog & "' OR MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoAdmin & "')), 0) AS Gastos, " & _
        " ISNULL(MC.nSaldoCap,0) AS SaldoCapital, ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoComisionAbog & "'), 0) AS ComisionAbog, Persona.cPersNombre, " & _
        " (SELECT PRP.CPERSCOD FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS CODANALISTA, (SELECT PERSO.cPersNombre " & _
        " FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS ANALISTA, (SELECT RHH.CUSER FROM RRHH RHH INNER JOIN " & _
        " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS SIGANALISTA, " & _
        " CONVERT(varchar(10), CONVERT(datetime, SUBSTRING(Mov.cMovNro, 1, 8), 103), 103) AS Fecha, SUBSTRING(Colocaciones.cLineaCred, 5, 1) AS Moneda, Constante.cConsDescripcion AS sMoneda, " & _
        " nTipoRecuperacion=case when Producto.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "') THEN 1 ELSE 2 END, " & _
        " nDesTipoRecuperacion=case when Producto.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "') THEN 'JUDICIAL' ELSE 'CASTIGADO' END " & _
        " FROM ColocRecup INNER JOIN Colocaciones ON ColocRecup.cCtaCod = Colocaciones.cCtaCod INNER JOIN MovCol MC INNER JOIN ProductoPersona PP ON MC.cCtaCod = PP.cCtaCod INNER JOIN " & _
        " Mov ON MC.nMovNro = Mov.nMovNro ON ColocRecup.cCtaCod = PP.cCtaCod INNER JOIN Persona ON PP.cPersCod = Persona.cPersCod INNER JOIN Producto ON " & _
        " Colocaciones.cCtaCod = Producto.cCtaCod AND PP.cCtaCod = Producto.cCtaCod INNER JOIN " & _
        " Constante ON SUBSTRING(Colocaciones.cLineaCred, 5, 1) = Constante.nConsValor " & _
        " INNER Join  OpeTpo Ope on ope.cOpeCod = MC.cOpeCod " & _
        " WHERE (PP.nPrdPersRelac=" & gColRelPersTitular & ") AND (Mov.nMovFlag=" & gMovFlagVigente & ") " & _
        " AND (Producto.nPrdEstado IN ('" & gColocEstRecVigJud & "', '" & gColocEstRecVigCast & "', '" & gColocEstRecCanJud & "', '" & gColocEstRecCanCast & "')) AND (Constante.nConsCod='" & gMoneda & "') " & _
        sCriterio
        
        'nTiporecuperacion=1 si es VigenteJudicial(2201) o CanceladoJudicial(2203)
        'nTiporecuperacion=2 si es VigenteCastigado(2201) o CanceladoCastigado(2203)
lsSQL = lsSQL & " and Ope.cOpeCod like '130%' and Ope.nOpeNiv<>1 "
lsSQL = lsSQL & " Order By (SELECT PERSO.cPersNombre FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON " & _
        " PRP.cPersCod = PERSO.cPersCod WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & "), " & _
        " SUBSTRING(Colocaciones.cLineaCred, 5, 1), case when Producto.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "') THEN 1 ELSE 2 END, SUBSTRING(Mov.cMovNro, 1, 8) "
 
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
     
        If psConAnalista = 1 Then
            lsCondiciones = lsCondiciones & " - " & lrDataRep!Analista & " : " & lrDataRep!CodAnalista
        End If
        
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE PAGOS DE CREDITOS POR ANALISTA", lsCondiciones, lnPage, 145, "", 138004)

        lnIndice = 0:  lnLineas = 7
        
        sTempoAnalista = "" & lrDataRep!CodAnalista
        sTempoMoneda = lrDataRep!Moneda
        sTempoTipoRecuperacion = lrDataRep!nTipoRecuperacion
        
        'Abogado
        lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA: " & lrDataRep!Analista & " - " & lrDataRep!SigAnalista & " : " & lrDataRep!CodAnalista & lsNegritaOff & Chr(10) & Chr(10)
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & Chr(10) & Chr(10)
        'Tipo de Recuperacion
        lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
        
        lnLineas = lnLineas + 6
        lnIndice = lnIndice + 6
          
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                
                'Totalizar por Analista
                
                If sTempoAnalista <> !CodAnalista Then
                
                    lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & ImpreFormat(sTotal11, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal41, 10, 2) & Space(1) & ImpreFormat(sTotal51, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal61, 10, 2) & Space(1) & ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                
                    lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(23) & ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal6, 10, 2) & Space(1) & ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                    
                    lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X  ANALISTA" & Space(29) & ImpreFormat(sSubTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sSubTotal2, 10, 2) & Space(1) & ImpreFormat(sSubTotal3, 10, 2) & _
                                Space(1) & ImpreFormat(sSubTotal4, 10, 2) & Space(1) & ImpreFormat(sSubTotal5, 10, 2) & _
                                Space(1) & ImpreFormat(sSubTotal6, 10, 2) & Space(1) & ImpreFormat(sSubTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                     
                    'Analista
                    lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA: " & lrDataRep!Analista & " - " & lrDataRep!SigAnalista & " : " & lrDataRep!CodAnalista & lsNegritaOff & Chr(10) & Chr(10)
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & Chr(10) & Chr(10)
                    'Tipo de Recuperacion
                    lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & Chr(10)
                    lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                    
                    sSubTotal1 = !Pagado
                    sSubTotal2 = !Capital
                    sSubTotal3 = !InteresCom
                    sSubTotal4 = !InteresMor
                    sSubTotal5 = !Gastos
                    sSubTotal6 = !SaldoCapital
                    sSubTotal7 = !ComisionAbog
                    
                    sTotal1 = !Pagado
                    sTotal2 = !Capital
                    sTotal3 = !InteresCom
                    sTotal4 = !InteresMor
                    sTotal5 = !Gastos
                    sTotal6 = !SaldoCapital
                    sTotal7 = !ComisionAbog
                        
                    sTotal11 = !Pagado
                    sTotal21 = !Capital
                    sTotal31 = !InteresCom
                    sTotal41 = !InteresMor
                    sTotal51 = !Gastos
                    sTotal61 = !SaldoCapital
                    sTotal71 = !ComisionAbog
                    
                    nItem = 1
                    
                    sTempoAnalista = "" & lrDataRep!CodAnalista
                    sTempoMoneda = lrDataRep!Moneda
                    sTempoTipoRecuperacion = lrDataRep!nTipoRecuperacion
                      
                    lnLineas = lnLineas + 15
                    lnIndice = lnIndice + 15
                     
                Else
                    sSubTotal1 = sSubTotal1 + !Pagado
                    sSubTotal2 = sSubTotal2 + !Capital
                    sSubTotal3 = sSubTotal3 + !InteresCom
                    sSubTotal4 = sSubTotal4 + !InteresMor
                    sSubTotal5 = sSubTotal5 + !Gastos
                    sSubTotal6 = sSubTotal6 + !SaldoCapital
                    sSubTotal7 = sSubTotal7 + !ComisionAbog
                    
                    If sTempoMoneda <> !Moneda Then
                
                        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & ImpreFormat(sTotal11, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal41, 10, 2) & Space(1) & ImpreFormat(sTotal51, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal61, 10, 2) & Space(1) & ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                    
                        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(23) & ImpreFormat(sTotal1, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal6, 10, 2) & Space(1) & ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                        
                        'Moneda
                        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & Chr(10) & Chr(10)
                        'Tipo de Recuperacion
                        lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & Chr(10)
                        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                        
                        sTotal1 = !Pagado
                        sTotal2 = !Capital
                        sTotal3 = !InteresCom
                        sTotal4 = !InteresMor
                        sTotal5 = !Gastos
                        sTotal6 = !SaldoCapital
                        sTotal7 = !ComisionAbog
                            
                        sTotal11 = !Pagado
                        sTotal21 = !Capital
                        sTotal31 = !InteresCom
                        sTotal41 = !InteresMor
                        sTotal51 = !Gastos
                        sTotal61 = !SaldoCapital
                        sTotal71 = !ComisionAbog
                        
                        nItem = 1
                        
                        sTempoMoneda = lrDataRep!Moneda
                        sTempoTipoRecuperacion = lrDataRep!nTipoRecuperacion
                    
                        lnLineas = lnLineas + 10
                        lnIndice = lnIndice + 10
                    Else
                        sTotal1 = sTotal1 + !Pagado
                        sTotal2 = sTotal2 + !Capital
                        sTotal3 = sTotal3 + !InteresCom
                        sTotal4 = sTotal4 + !InteresMor
                        sTotal5 = sTotal5 + !Gastos
                        sTotal6 = sTotal6 + !SaldoCapital
                        sTotal7 = sTotal7 + !ComisionAbog
                        
                        If sTempoTipoRecuperacion <> !nTipoRecuperacion Then
                    
                            lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                            lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & ImpreFormat(sTotal11, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal41, 10, 2) & Space(1) & ImpreFormat(sTotal51, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal61, 10, 2) & Space(1) & ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                    
                            'Tipo de Recuperacion
                            lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & Chr(10)
                            lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                            
                            sTotal11 = !Pagado
                            sTotal21 = !Capital
                            sTotal31 = !InteresCom
                            sTotal41 = !InteresMor
                            sTotal51 = !Gastos
                            sTotal61 = !SaldoCapital
                            sTotal71 = !ComisionAbog
                                 
                            nItem = 1
                            
                            sTempoTipoRecuperacion = lrDataRep!nTipoRecuperacion
                        
                            lnLineas = lnLineas + 5
                            lnIndice = lnIndice + 5
                        Else
                            sTotal11 = sTotal11 + !Pagado
                            sTotal21 = sTotal21 + !Capital
                            sTotal31 = sTotal31 + !InteresCom
                            sTotal41 = sTotal41 + !InteresMor
                            sTotal51 = sTotal51 + !Gastos
                            sTotal61 = sTotal61 + !SaldoCapital
                            sTotal71 = sTotal71 + !ComisionAbog
                        End If
                    End If
                End If
                
                sCliente = !cpersNombre
                 
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(nItem, 4, 0) & Space(1) & !Fecha & Space(1) & _
                           sCliente & Space(1) & ImpreFormat(!Pagado, 10, 2) & Space(1) & ImpreFormat(!Capital, 10, 2) & _
                           Space(1) & ImpreFormat(!InteresCom, 10, 2) & Space(1) & ImpreFormat(!InteresMor, 10, 2) & _
                           Space(1) & ImpreFormat(!Gastos, 10, 2) & Space(1) & ImpreFormat(!SaldoCapital, 10, 2) & _
                           Space(1) & ImpreFormat(!ComisionAbog, 10, 2) & Chr(10)
                
                 
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE PAGOS DE CREDITOS POR ANALISTA", lsCondiciones, lnPage, 145, "", 138004)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & ImpreFormat(sTotal11, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal41, 10, 2) & Space(1) & ImpreFormat(sTotal51, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal61, 10, 2) & Space(1) & ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
    
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(23) & ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal6, 10, 2) & Space(1) & ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
        
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X  ANALISTA" & Space(29) & ImpreFormat(sSubTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal2, 10, 2) & Space(1) & ImpreFormat(sSubTotal3, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal4, 10, 2) & Space(1) & ImpreFormat(sSubTotal5, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal6, 10, 2) & Space(1) & ImpreFormat(sSubTotal7, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
      
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
    nRepo138004_ListadoPagosxAnalista = lsCadBuffer
End Function
  
Public Function nRepo138011_ListadoCreditosNuevos(ByVal pdFecIni As String, ByVal pdFecFin As String) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lsNegritaOn As String
Dim lsNegritaOff As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim nItem As Integer
Dim lnPage As Integer
Dim loImprimeCab As NColPImpre
        
Dim sCliente As String * 30
Dim lsCondiciones As String
Dim sCriterio As String
Dim sTempoMoneda As Integer
Dim sTempoAgencia As String

'Dim sSubTotal1 As Currency
'Dim sSubTotal2 As Currency
'Dim sSubTotal3 As Currency
'Dim sSubTotal4 As Currency
'Dim sSubTotal5 As Currency
'Dim sSubTotal31 As Currency

Dim sTotal1 As Currency
Dim sTotal2 As Currency
Dim sTotal3 As Currency
Dim sTotal4 As Currency
Dim sTotal5 As Currency
Dim sTotal31 As Currency
 
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

'sCriterio = " AND (CONVERT (DATETIME, SUBSTRING(CONVERT (VARCHAR,  CR.dIngRecup), 1, 11)) Between '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "') "
 
lsCondiciones = "Del " & pdFecIni & " Al " & pdFecFin
         

lsSQL = " Select T.cCtaCod, Substring(T.cCtaCod, 4,2) Agencia, A.cAgeDescripcion, T.nCapital, T.nInteres, "
lsSQL = lsSQL & " T.nInteresMor, T.nGastos, C.dVigencia, CR.dIngRecup, C.cLineaCred, P.cPersNombre,"
lsSQL = lsSQL & " SUBSTRING(T.cCtaCod, 9, 1) AS Moneda, CON.cConsDescripcion AS sMoneda,  "
lsSQL = lsSQL & " (  SELECT RH.cUser FROM RRHH RH INNER JOIN PRODUCTOPERSONA PP ON PP.cPersCod = RH.cPersCod "
lsSQL = lsSQL & " Where PP.cCtaCod = T.cCtaCod And PP.nPrdPersRelac = 28) AS ANALISTA, C.nMontoCol "
lsSQL = lsSQL & " From "
lsSQL = lsSQL & " ( Select CR.cCtaCod, "
lsSQL = lsSQL & " IsNull( "
lsSQL = lsSQL & " ( SELECT SUM(nmonto) "
lsSQL = lsSQL & " FROM movcoldet MCD INNER JOIN MOV M ON MCD.nMovNro = M.nMovNro "
lsSQL = lsSQL & " WHERE M.nMovFlag=0 and MCD.cCtaCod = CR.cCtaCod AND MCD.nPrdConceptoCod = '3000' And MCD.cOpecod = '130100' "
lsSQL = lsSQL & " ), 0)  AS nCapital, "
lsSQL = lsSQL & " IsNull( "
lsSQL = lsSQL & " ( SELECT SUM(nmonto) "
lsSQL = lsSQL & " FROM movcoldet MCD INNER JOIN MOV M ON MCD.nMovNro = M.nMovNro "
lsSQL = lsSQL & " WHERE M.nMovFlag=0 and MCD.cCtaCod = CR.cCtaCod AND  MCD.nPrdConceptoCod = '3100' And MCD.cOpecod = '130100' "
lsSQL = lsSQL & " ), 0)  AS nInteres, "
lsSQL = lsSQL & " IsNull( "
lsSQL = lsSQL & " ( SELECT SUM(nmonto) "
lsSQL = lsSQL & " FROM MovColDet MCD INNER JOIN MOV M ON MCD.nMovNro = M.nMovNro "
lsSQL = lsSQL & " WHERE M.nMovFlag=0 and MCD.cCtaCod = CR.cCtaCod AND MCD.nPrdConceptoCod = '3101' And MCD.cOpecod = '130100' "
lsSQL = lsSQL & " ), 0)  AS nInteresMor, "
lsSQL = lsSQL & " IsNull( "
lsSQL = lsSQL & " ( SELECT SUM(nmonto) "
lsSQL = lsSQL & " FROM movcoldet MCD INNER JOIN MOV M ON MCD.nMovNro = M.nMovNro "
lsSQL = lsSQL & " WHERE M.nMovFlag=0 and MCD.cCtaCod = CR.cCtaCod AND MCD.nPrdConceptoCod NOT IN('3000', '3100', '3101') "
lsSQL = lsSQL & " and MCD.cOpecod = '130100' "
lsSQL = lsSQL & " ), 0)  AS nGastos "
lsSQL = lsSQL & " FROM ColocRecup CR "
lsSQL = lsSQL & " where convert(varchar(8),CR.dIngRecup,112)>='" & Format(pdFecIni, "YYYYMMdd") & "' And convert(varchar(8), CR.dIngRecup,112)<='" & Format(pdFecFin, "YYYYMMdd") & "'"
lsSQL = lsSQL & " ) T INNER JOIN Colocaciones C ON C.cCtaCod = T.cCtaCod "
lsSQL = lsSQL & " INNER JOIN ColocRecup CR ON T.cCtaCod = CR.cCtaCod "
lsSQL = lsSQL & " INNER JOIN  ProductoPersona PP ON T.cCtaCod = PP.cCtaCod "
lsSQL = lsSQL & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod "
lsSQL = lsSQL & " INNER JOIN Agencias A ON A.cAgeCod =  Substring(T.cCtaCod, 4,2) "
lsSQL = lsSQL & " INNER JOIN Producto PC ON T.cCtaCod = PC.cCtaCod "
lsSQL = lsSQL & " INNER JOIN  Constante CON ON SUBSTRING(T.cCtaCod, 9, 1) = CON.nConsValor "
lsSQL = lsSQL & " Where(PP.nPrdPersRelac =" & gColRelPersTitular & ") And (CON.nConsCod=" & gMoneda & ") "
lsSQL = lsSQL & " AND (PC.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud
lsSQL = lsSQL & "', '" & gColocEstRecVigCast & "', '" & gColocEstRecCanCast & "'))"
lsSQL = lsSQL & " Order By Substring(T.cCtaCod, 4,2), SUBSTRING(T.cCtaCod, 9, 1)"
  
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS NUEVOS ", lsCondiciones, lnPage, 175, "", 138011)

        lnIndice = 0:  lnLineas = 7
        sTempoMoneda = lrDataRep!Moneda
        sTempoAgencia = lrDataRep!Agencia
        
        'Agencia
        lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & Chr(10) & Chr(10)
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & Chr(10)
                lsCadImp = lsCadImp & String(175, "-") & Chr(10)
        
        lnLineas = lnLineas + 4
        lnIndice = lnIndice + 4
        
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                
                If sTempoAgencia <> !Agencia Then
                        
                    lsCadImp = lsCadImp & String(175, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA " & Space(79) & ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 8, 2) & _
                                Space(1) & ImpreFormat(sTotal31, 8, 2) & Space(1) & ImpreFormat(sTotal4, 8, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
  
                    'Agencia
                    lsCadImp = lsCadImp & lsNegritaOn & "Agencia: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & Chr(10) & Chr(10)
                    
                      
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "Moneda: " & lrDataRep!sMoneda & lsNegritaOff & Chr(10)
                    lsCadImp = lsCadImp & String(175, "-") & Chr(10)
                    
                    sTotal1 = !nMontoCol
                    sTotal2 = !nCapital
                    sTotal3 = !nInteres
                    sTotal31 = !nInteresMor
                    sTotal4 = !nGastos
                    sTotal5 = !nCapital + !nInteres + !nInteresMor + !nGastos
                    
                    nItem = 1
                    
                    sTempoAgencia = lrDataRep!Agencia
                    sTempoMoneda = lrDataRep!Moneda
                                         
                    lnLineas = lnLineas + 9 - 3
                    lnIndice = lnIndice + 9 - 3
                    
                Else
                    If sTempoMoneda <> !Moneda Then
                    
                        lsCadImp = lsCadImp & String(175, "-") & Chr(10)
 '
                        'Moneda
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA " & Space(79) & ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 8, 2) & _
                                Space(1) & ImpreFormat(sTotal31, 8, 2) & Space(1) & ImpreFormat(sTotal4, 8, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
  

                        lsCadImp = lsCadImp & lsNegritaOn & "Moneda: " & lrDataRep!sMoneda & lsNegritaOff & Chr(10) '& Chr(10)
                        lsCadImp = lsCadImp & String(175, "-") & Chr(10)
                         
                        sTotal1 = !nMontoCol
                        sTotal2 = !nCapital
                        sTotal3 = !nInteres
                        sTotal31 = !nInteresMor
                        sTotal4 = !nGastos
                        sTotal5 = !nCapital + !nInteres + !nInteresMor + !nGastos
                        nItem = 1
                        
'                        sTempoAgencia = lrDataRep!Agencia
                        sTempoMoneda = lrDataRep!Moneda
                        
                        lnLineas = lnLineas + 5
                        lnIndice = lnIndice + 5
                        
                    Else
                        sTotal1 = sTotal1 + !nMontoCol
                        sTotal2 = sTotal2 + !nCapital
                        sTotal3 = sTotal3 + !nInteres
                        sTotal31 = sTotal31 + !nInteresMor
                        sTotal4 = sTotal4 + !nGastos
                        sTotal5 = sTotal5 + !nCapital + !nInteres + !nInteresMor + !nGastos
                    End If
                End If
                
                sCliente = !cpersNombre
                       
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(nItem, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dVigencia, "dd/MM/YYYY") & _
                            Space(1) & Format(!dIngRecup, "dd/MM/YYYY") & Space(1) & !cLineaCred & Space(1) & _
                            sCliente & Space(1) & !Analista & Space(1) & Space(1) & _
                            ImpreFormat(!nMontoCol, 10, 2) & Space(1) & ImpreFormat(!nCapital, 10, 2) & Space(1) & _
                            ImpreFormat(!nInteres, 8, 2) & Space(1) & ImpreFormat(!nInteresMor, 8, 2) & Space(1) & _
                            ImpreFormat(!nGastos, 8, 2) & Space(1) & _
                            ImpreFormat(!nCapital + !nInteres + !nInteresMor + !nGastos, 10, 2) & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS NUEVOS ", lsCondiciones, lnPage, 165, "", 138011)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            RaiseEvent CloseProgress
        End With
       
        lsCadImp = lsCadImp & String(175, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA " & Space(79) & ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 8, 2) & _
                    Space(1) & ImpreFormat(sTotal31, 8, 2) & Space(1) & ImpreFormat(sTotal4, 8, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & lsNegritaOff & Chr(10)

        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
    nRepo138011_ListadoCreditosNuevos = lsCadBuffer
End Function

Public Function nRepo138013_(ByVal sBuscarExp As Byte, ByVal sConSinExp As Byte, ByVal sMoneda As Byte, _
ByVal sEstados As String, ByVal sMasCero As Byte, ByVal sCero As Byte, _
ByVal pdf1 As Date, pdf2 As Date) As String


Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As NColPImpre
         
Dim lsNegritaOn As String
Dim lsNegritaOff As String
        
Dim sCliente As String * 30
Dim lsCondiciones As String
Dim sCriterio As String
Dim sTempoMoneda As Integer
Dim sTempoAgencia As String
Dim sTempoFteFinancia As String

'TOTALES por Fuente de Financiamiento
Dim sTotal11 As Currency
Dim sTotal21 As Currency
Dim sTotal31 As Currency
Dim sTotal41 As Currency
Dim sTotal51 As Currency
Dim sTotal61 As Currency

'TOTALES por MONEDA
Dim sSubTotal1 As Currency
Dim sSubTotal2 As Currency
Dim sSubTotal3 As Currency
Dim sSubTotal4 As Currency
Dim sSubTotal5 As Currency
Dim sSubTotal6 As Currency

'TOTALES por Agencia
Dim sTotal1 As Currency
Dim sTotal2 As Currency
Dim sTotal3 As Currency
Dim sTotal4 As Currency
Dim sTotal5 As Currency
Dim sTotal6 As Currency
 
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

    'Si se busca expediente
    If sBuscarExp = 1 Then
        sCriterio = " AND (CR.nDemanda=" & sConSinExp & ")"
    End If
    
    'Busco por estados en la tabla Producto
    sCriterio = sCriterio & " AND (P.nPrdEstado IN(" & sEstados & ")) "
    
    'Si se busca por MONEDA
    If sMoneda <> 3 Then
        sCriterio = sCriterio & " AND (SUBSTRING(C.cLineaCred, 5, 1)=" & sMoneda & ")"
    End If
    
    'Montos mayores o iguales a cero
    If sMasCero = 1 And sCero = 1 Then
        sCriterio = sCriterio & " AND (P.nSaldo>=0) "
    Else
        If sMasCero = 1 Then
            sCriterio = sCriterio & " AND (P.nSaldo>0) "
        End If
        If sCero = 1 Then
            sCriterio = sCriterio & " AND (P.nSaldo=0) "
        End If
    End If
  
    lsSQL = " SELECT P.cCtaCod, SUBSTRING(P.cCtaCod, 4, 2) AS Agencia, A.cAgeDescripcion, " & _
            " C.dVigencia, CR.dIngRecup, C.cLineaCred, CLC.cDescripcion, Persona.cPersNombre, " & _
            " (SELECT RHH.CUSER FROM RRHH RHH INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD " & _
            " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS ANALISTA, C.nMontoCol, P.nSaldo AS SaldoCapital, " & _
            " CR.nSaldoIntComp As Interes, CR.nSaldoIntMor AS Mora, CR.nSaldoGasto AS Gasto, SUBSTRING(C.cLineaCred, 5, 1) AS Moneda, " & _
            " CON.cConsDescripcion AS sMoneda, SUBSTRING(C.cLineaCred, 1, 2) AS cCodFuente, (SELECT CLC.cDescripcion " & _
            " FROM ColocLineaCredito CLC WHERE substring(C.clineacred, 1, 2) = CLC.cLineaCred) AS cDesFuente " & _
            " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod INNER JOIN " & _
            " ColocRecup CR ON C.cCtaCod = CR.cCtaCod INNER JOIN ColocLineaCredito CLC ON C.cLineaCred = CLC.cLineaCred INNER JOIN " & _
            " ProductoPersona PP ON P.cCtaCod = PP.cCtaCod INNER JOIN Persona ON PP.cPersCod = Persona.cPersCod INNER JOIN " & _
            " Constante CON ON SUBSTRING(C.cLineaCred, 5, 1) = CON.nConsValor INNER JOIN Agencias A ON A.cAgeCod = SUBSTRING(P.cCtaCod, 4, 2) " & _
            " Where (PP.nPrdPersRelac=" & gColRelPersTitular & ") And (CON.nConsCod=" & gMoneda & ") " & _
            " AND datediff(d,dPrdEstado,'" & Format(pdf1, "mm/dd/yyyy") & "')<=0 and  datediff(d,dPrdEstado,'" & Format(pdf2, "mm/dd/yyyy") & "')>=0 " & _
            sCriterio & " ORDER BY SUBSTRING(C.cLineaCred, 5, 1), SUBSTRING(P.cCtaCod, 4, 2), SUBSTRING(C.cLineaCred, 1, 2) "
        
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CREDITOS CASTIGADOS", lsCondiciones, lnPage, 175, "", 138012)

        lnIndice = 0:  lnLineas = 7
        sTempoMoneda = lrDataRep!Moneda
        sTempoAgencia = lrDataRep!Agencia
        sTempoFteFinancia = lrDataRep!cCodFuente
        
        'MONEDA
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & Chr(10) & Chr(10)
        'Agencia
        lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & Chr(10) & Chr(10)
        'Fuente de Financiamiento
        lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & Chr(10)
        
        lsCadImp = lsCadImp & String(175, "-") & Chr(10)
        
        lnLineas = lnLineas + 6
        lnIndice = lnIndice + 6
         
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                                
                'TOTALizar por MONEDA
                If sTempoMoneda <> !Moneda Then
                
                    lsCadImp = lsCadImp & String(175, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61) & ImpreFormat(sTotal11, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 9, 2) & _
                                Space(1) & ImpreFormat(sTotal41, 9, 2) & Space(1) & ImpreFormat(sTotal51, 9, 2) & _
                                Space(1) & ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & Chr(10)
                    
                    lsCadImp = lsCadImp & String(175, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X AGENCIA" & Space(78) & ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 9, 2) & _
                                Space(1) & ImpreFormat(sTotal4, 9, 2) & Space(1) & ImpreFormat(sTotal5, 9, 2) & _
                                Space(1) & ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & Chr(10)
                    
                    lsCadImp = lsCadImp & String(175, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(79) & ImpreFormat(sSubTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sSubTotal2, 10, 2) & Space(1) & ImpreFormat(sSubTotal3, 9, 2) & _
                                Space(1) & ImpreFormat(sSubTotal4, 9, 2) & Space(1) & ImpreFormat(sSubTotal5, 9, 2) & _
                                Space(1) & ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)

                    'MONEDA
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & Chr(10) & Chr(10)
                    'Agencia
                    lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & Chr(10) & Chr(10)
                    'Fuente de Financiamiento
                    lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & Chr(10)
                    lsCadImp = lsCadImp & String(175, "-") & Chr(10)
                    
                    sSubTotal1 = !nMontoCol
                    sSubTotal2 = !SaldoCapital
                    sSubTotal3 = !Interes
                    sSubTotal4 = !Mora
                    sSubTotal5 = !Gasto
                    sSubTotal6 = !SaldoCapital + !Interes + !Mora + !Gasto
                    
                    sTotal1 = !nMontoCol
                    sTotal2 = !SaldoCapital
                    sTotal3 = !Interes
                    sTotal4 = !Mora
                    sTotal5 = !Gasto
                    sTotal6 = !SaldoCapital + !Interes + !Mora + !Gasto
                    
                    sTotal11 = !nMontoCol
                    sTotal21 = !SaldoCapital
                    sTotal31 = !Interes
                    sTotal41 = !Mora
                    sTotal51 = !Gasto
                    sTotal61 = !SaldoCapital + !Interes + !Mora + !Gasto
                    
                    nItem = 1
                    
                    lnLineas = lnLineas + 13
                    lnIndice = lnIndice + 13
                    
                    sTempoMoneda = lrDataRep!Moneda
                    sTempoAgencia = lrDataRep!Agencia
                    sTempoFteFinancia = lrDataRep!cCodFuente
                    
                Else
                    sSubTotal1 = sSubTotal1 + !nMontoCol
                    sSubTotal2 = sSubTotal2 + !SaldoCapital
                    sSubTotal3 = sSubTotal3 + !Interes
                    sSubTotal4 = sSubTotal4 + !Mora
                    sSubTotal5 = sSubTotal5 + !Gasto
                    sSubTotal6 = sSubTotal6 + !SaldoCapital + !Interes + !Mora + !Gasto
                   
                    If sTempoAgencia <> !Agencia Then
                
                        lsCadImp = lsCadImp & String(175, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61) & ImpreFormat(sTotal11, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 9, 2) & _
                                    Space(1) & ImpreFormat(sTotal41, 9, 2) & Space(1) & ImpreFormat(sTotal51, 9, 2) & _
                                    Space(1) & ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & Chr(10)
                        
                        lsCadImp = lsCadImp & String(175, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X AGENCIA" & Space(78) & ImpreFormat(sTotal1, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 9, 2) & _
                                    Space(1) & ImpreFormat(sTotal4, 9, 2) & Space(1) & ImpreFormat(sTotal5, 9, 2) & _
                                    Space(1) & ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                         
                        sTotal1 = !nMontoCol
                        sTotal2 = !SaldoCapital
                        sTotal3 = !Interes
                        sTotal4 = !Mora
                        sTotal5 = !Gasto
                        sTotal6 = !SaldoCapital + !Interes + !Mora + !Gasto
                    
                        sTotal11 = !nMontoCol
                        sTotal21 = !SaldoCapital
                        sTotal31 = !Interes
                        sTotal41 = !Mora
                        sTotal51 = !Gasto
                        sTotal61 = !SaldoCapital + !Interes + !Mora + !Gasto
                                             
                        nItem = 1
                                             
                        sTempoAgencia = lrDataRep!Agencia
                        sTempoFteFinancia = lrDataRep!cCodFuente
                        
                        'Agencia
                        lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & Chr(10) & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & Chr(10)
                        lsCadImp = lsCadImp & String(175, "-") & Chr(10)
                        lnLineas = lnLineas + 9
                        lnIndice = lnIndice + 9
                     
                    Else
                        sTotal1 = sTotal1 + !nMontoCol
                        sTotal2 = sTotal2 + !SaldoCapital
                        sTotal3 = sTotal3 + !Interes
                        sTotal4 = sTotal4 + !Mora
                        sTotal5 = sTotal5 + !Gasto
                        sTotal6 = sTotal6 + !SaldoCapital + !Interes + !Mora + !Gasto
                     
                        'TOTALizar por fuente de financiamiento
                        If sTempoFteFinancia <> !cCodFuente Then
                    
                            lsCadImp = lsCadImp & String(175, "-") & Chr(10)
                            lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61) & ImpreFormat(sTotal11, 10, 2) & _
                                        Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 9, 2) & _
                                        Space(1) & ImpreFormat(sTotal41, 9, 2) & Space(1) & ImpreFormat(sTotal51, 9, 2) & _
                                        Space(1) & ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                            
                            'Fuente de Financiamiento
                            lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & Chr(10)
                            lsCadImp = lsCadImp & String(175, "-") & Chr(10)
                             
                            sTotal11 = 0
                            sTotal21 = 0
                            sTotal31 = 0
                            sTotal41 = 0
                            sTotal51 = 0
                            sTotal61 = 0
                            
                            nItem = 1
                            
                            sTempoFteFinancia = lrDataRep!cCodFuente
                            lnLineas = lnLineas + 5
                                
                        Else
                            sTotal11 = sTotal11 + !nMontoCol
                            sTotal21 = sTotal21 + !SaldoCapital
                            sTotal31 = sTotal31 + !Interes
                            sTotal41 = sTotal41 + !Mora
                            sTotal51 = sTotal51 + !Gasto
                            sTotal61 = sTotal61 + !SaldoCapital + !Interes + !Mora + !Gasto
                         End If
                    End If
                End If
                
                sCliente = !cpersNombre
                
                lsCadImp = lsCadImp & ImpreFormat(nItem, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dVigencia, "dd/MM/YYYY") & _
                            Space(1) & Format(!dIngRecup, "dd/MM/YYYY") & Space(1) & !cLineaCred & Space(1) & _
                            sCliente & Space(1) & !Analista & Space(1) & Space(1) & _
                            ImpreFormat(!nMontoCol, 10, 2) & Space(1) & ImpreFormat(!SaldoCapital, 10, 2) & Space(1) & _
                            ImpreFormat(!Interes, 9, 2) & Space(1) & ImpreFormat(!Mora, 9, 2) & Space(1) & _
                            ImpreFormat(!Gasto, 9, 2) & Space(1) & ImpreFormat(!SaldoCapital + !Interes + !Mora + !Gasto, 10, 2) & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS CASTIGADOS", lsCondiciones, lnPage, 175, "", 138012)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            RaiseEvent CloseProgress
        End With
        lsCadImp = lsCadImp & String(175, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61) & ImpreFormat(sTotal11, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 9, 2) & _
                    Space(1) & ImpreFormat(sTotal41, 9, 2) & Space(1) & ImpreFormat(sTotal51, 9, 2) & _
                    Space(1) & ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(175, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(79) & ImpreFormat(sSubTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal2, 10, 2) & Space(1) & ImpreFormat(sSubTotal3, 9, 2) & _
                    Space(1) & ImpreFormat(sSubTotal4, 9, 2) & Space(1) & ImpreFormat(sSubTotal5, 9, 2) & _
                    Space(1) & ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(175, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X AGENCIA" & Space(78) & ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 9, 2) & _
                    Space(1) & ImpreFormat(sTotal4, 9, 2) & Space(1) & ImpreFormat(sTotal5, 9, 2) & _
                    Space(1) & ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & Chr(10)
        
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
   nRepo138013_ = lsCadBuffer


End Function
'
Public Function nRepo138012_ListadoCreditosEnJudicial(ByVal sBuscarExp As Byte, ByVal sConSinExp As Byte, ByVal sMoneda As Byte, _
                                                        ByVal sEstados As String, ByVal sMasCero As Byte, ByVal sCero As Byte, _
                                                        sRotulo As String) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As NColPImpre
         
Dim lsNegritaOn As String
Dim lsNegritaOff As String
        
Dim sCliente As String * 30
Dim lsCondiciones As String
Dim sCriterio As String
Dim sTempoMoneda As Integer
Dim sTempoAgencia As String
Dim sTempoFteFinancia As String

'TOTALES por Fuente de Financiamiento
Dim sTotal11 As Currency
Dim sTotal21 As Currency
Dim sTotal31 As Currency
Dim sTotal41 As Currency
Dim sTotal51 As Currency
Dim sTotal61 As Currency

'TOTALES por MONEDA
Dim sSubTotal1 As Currency
Dim sSubTotal2 As Currency
Dim sSubTotal3 As Currency
Dim sSubTotal4 As Currency
Dim sSubTotal5 As Currency
Dim sSubTotal6 As Currency

'TOTALES por Agencia
Dim sTotal1 As Currency
Dim sTotal2 As Currency
Dim sTotal3 As Currency
Dim sTotal4 As Currency
Dim sTotal5 As Currency
Dim sTotal6 As Currency
 
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

    'Si se busca expediente
    If sBuscarExp = 1 Then
        sCriterio = " AND (CR.nDemanda=" & sConSinExp & ")"
    End If
    
    'Busco por estados en la tabla Producto
    sCriterio = sCriterio & " AND (P.nPrdEstado IN(" & sEstados & ")) "
    
    'Si se busca por MONEDA
    If sMoneda <> 3 Then
        sCriterio = sCriterio & " AND (SUBSTRING(C.cLineaCred, 5, 1)=" & sMoneda & ")"
    End If
    
    'Montos mayores o iguales a cero
    If sMasCero = 1 And sCero = 1 Then
        sCriterio = sCriterio & " AND (P.nSaldo>=0) "
    Else
        If sMasCero = 1 Then
            sCriterio = sCriterio & " AND (P.nSaldo>0) "
        End If
        If sCero = 1 Then
            sCriterio = sCriterio & " AND (P.nSaldo=0) "
        End If
    End If
  
    lsSQL = " SELECT P.cCtaCod, sExpe = case when CR.nDemanda = 1 then 'S' when CR.nDemanda =2 then 'N' end, SUBSTRING(P.cCtaCod, 4, 2) AS Agencia, A.cAgeDescripcion, " & _
            " C.dVigencia, CR.dIngRecup, C.cLineaCred, CLC.cDescripcion, Persona.cPersNombre, " & _
            " (SELECT RHH.CUSER FROM RRHH RHH INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD " & _
            " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS ANALISTA, C.nMontoCol, P.nSaldo AS SaldoCapital, " & _
            " CR.nSaldoIntComp As Interes, CR.nSaldoIntMor AS Mora, CR.nSaldoGasto AS Gasto, SUBSTRING(P.cCtaCod, 9, 1) AS Moneda, " & _
            " CON.cConsDescripcion AS sMoneda, SUBSTRING(C.cLineaCred, 1, 2) AS cCodFuente, (SELECT CLC.cDescripcion " & _
            " FROM ColocLineaCredito CLC WHERE substring(C.clineacred, 1, 2) = CLC.cLineaCred) AS cDesFuente " & _
            " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod INNER JOIN " & _
            " ColocRecup CR ON C.cCtaCod = CR.cCtaCod INNER JOIN ColocLineaCredito CLC ON C.cLineaCred = CLC.cLineaCred INNER JOIN " & _
            " ProductoPersona PP ON P.cCtaCod = PP.cCtaCod INNER JOIN Persona ON PP.cPersCod = Persona.cPersCod INNER JOIN " & _
            " Constante CON ON SUBSTRING(P.cCtaCod, 9, 1) = CON.nConsValor INNER JOIN Agencias A ON A.cAgeCod = SUBSTRING(P.cCtaCod, 4, 2) " & _
            " Where (PP.nPrdPersRelac=" & gColRelPersTitular & ") And (CON.nConsCod=" & gMoneda & ") " & _
            sCriterio & " ORDER BY SUBSTRING(P.cCtaCod, 9, 1), SUBSTRING(P.cCtaCod, 4, 2), SUBSTRING(C.cLineaCred, 1, 2) "
        
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS EN JUDICIAL ", lsCondiciones, lnPage, 175 - 19, "", 138012)

        lnIndice = 0:  lnLineas = 7
        sTempoMoneda = lrDataRep!Moneda
        sTempoAgencia = lrDataRep!Agencia
        sTempoFteFinancia = lrDataRep!cCodFuente
        
        'MONEDA
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & Chr(10) & Chr(10)
        'Agencia
        lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & Chr(10) & Chr(10)
        'Fuente de Financiamiento
        lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & Chr(10)
        
        lsCadImp = lsCadImp & lsNegritaOn & sRotulo & Chr(10)
        
        lsCadImp = lsCadImp & String(175 - 19, "-") & lsNegritaOff & Chr(10)
        
        
        lnLineas = lnLineas + 7
        lnIndice = lnIndice + 7
         
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                                
                'TOTALizar por MONEDA
                If sTempoMoneda <> !Moneda Then
                
                    lsCadImp = lsCadImp & String(175 - 19, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61 - 23) & ImpreFormat(sTotal11, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 9, 2) & _
                                Space(1) & ImpreFormat(sTotal41, 9, 2) & Space(1) & ImpreFormat(sTotal51, 9, 2) & _
                                Space(1) & ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & Chr(10)
                    
                    lsCadImp = lsCadImp & String(175 - 19, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X AGENCIA" & Space(78 - 23) & ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 9, 2) & _
                                Space(1) & ImpreFormat(sTotal4, 9, 2) & Space(1) & ImpreFormat(sTotal5, 9, 2) & _
                                Space(1) & ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & Chr(10)
                    
                    lsCadImp = lsCadImp & String(175 - 19, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(79 - 23) & ImpreFormat(sSubTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sSubTotal2, 10, 2) & Space(1) & ImpreFormat(sSubTotal3, 9, 2) & _
                                Space(1) & ImpreFormat(sSubTotal4, 9, 2) & Space(1) & ImpreFormat(sSubTotal5, 9, 2) & _
                                Space(1) & ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)

                    'MONEDA
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & Chr(10) & Chr(10)
                    'Agencia
                    lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & Chr(10) & Chr(10)
                    'Fuente de Financiamiento
                    lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & sRotulo & Chr(10)
                    
                    lsCadImp = lsCadImp & String(175 - 19, "-") & Chr(10)
                    
                    sSubTotal1 = !nMontoCol
                    sSubTotal2 = !SaldoCapital
                    sSubTotal3 = !Interes
                    sSubTotal4 = !Mora
                    sSubTotal5 = !Gasto
                    sSubTotal6 = !SaldoCapital + !Interes + !Mora + !Gasto
                    
                    sTotal1 = !nMontoCol
                    sTotal2 = !SaldoCapital
                    sTotal3 = !Interes
                    sTotal4 = !Mora
                    sTotal5 = !Gasto
                    sTotal6 = !SaldoCapital + !Interes + !Mora + !Gasto
                    
                    sTotal11 = !nMontoCol
                    sTotal21 = !SaldoCapital
                    sTotal31 = !Interes
                    sTotal41 = !Mora
                    sTotal51 = !Gasto
                    sTotal61 = !SaldoCapital + !Interes + !Mora + !Gasto
                    
                    nItem = 1
                    
                    lnLineas = lnLineas + 13
                    lnIndice = lnIndice + 13
                    
                    sTempoMoneda = lrDataRep!Moneda
                    sTempoAgencia = lrDataRep!Agencia
                    sTempoFteFinancia = lrDataRep!cCodFuente
                    
                Else
                    sSubTotal1 = sSubTotal1 + !nMontoCol
                    sSubTotal2 = sSubTotal2 + !SaldoCapital
                    sSubTotal3 = sSubTotal3 + !Interes
                    sSubTotal4 = sSubTotal4 + !Mora
                    sSubTotal5 = sSubTotal5 + !Gasto
                    sSubTotal6 = sSubTotal6 + !SaldoCapital + !Interes + !Mora + !Gasto
                   
                    If sTempoAgencia <> !Agencia Then
                
                        lsCadImp = lsCadImp & String(175 - 19, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61 - 23) & ImpreFormat(sTotal11, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 9, 2) & _
                                    Space(1) & ImpreFormat(sTotal41, 9, 2) & Space(1) & ImpreFormat(sTotal51, 9, 2) & _
                                    Space(1) & ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & Chr(10)
                        
                        lsCadImp = lsCadImp & String(175 - 19, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X AGENCIA" & Space(78 - 23) & ImpreFormat(sTotal1, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 9, 2) & _
                                    Space(1) & ImpreFormat(sTotal4, 9, 2) & Space(1) & ImpreFormat(sTotal5, 9, 2) & _
                                    Space(1) & ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                         
                        sTotal1 = !nMontoCol
                        sTotal2 = !SaldoCapital
                        sTotal3 = !Interes
                        sTotal4 = !Mora
                        sTotal5 = !Gasto
                        sTotal6 = !SaldoCapital + !Interes + !Mora + !Gasto
                    
                        sTotal11 = !nMontoCol
                        sTotal21 = !SaldoCapital
                        sTotal31 = !Interes
                        sTotal41 = !Mora
                        sTotal51 = !Gasto
                        sTotal61 = !SaldoCapital + !Interes + !Mora + !Gasto
                                             
                        nItem = 1
                                             
                        sTempoAgencia = lrDataRep!Agencia
                        sTempoFteFinancia = lrDataRep!cCodFuente
                        
                        'Agencia
                        lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & Chr(10) & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & sRotulo & Chr(10)
                        lsCadImp = lsCadImp & String(175 - 19, "-") & Chr(10)
                        lnLineas = lnLineas + 10
                        lnIndice = lnIndice + 10
                     
                    Else
                        sTotal1 = sTotal1 + !nMontoCol
                        sTotal2 = sTotal2 + !SaldoCapital
                        sTotal3 = sTotal3 + !Interes
                        sTotal4 = sTotal4 + !Mora
                        sTotal5 = sTotal5 + !Gasto
                        sTotal6 = sTotal6 + !SaldoCapital + !Interes + !Mora + !Gasto
                     
                        'TOTALizar por fuente de financiamiento
                        If sTempoFteFinancia <> !cCodFuente Then
                    
                            lsCadImp = lsCadImp & String(175 - 19, "-") & Chr(10)
                            lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61 - 23) & ImpreFormat(sTotal11, 10, 2) & _
                                        Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 9, 2) & _
                                        Space(1) & ImpreFormat(sTotal41, 9, 2) & Space(1) & ImpreFormat(sTotal51, 9, 2) & _
                                        Space(1) & ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                            
                            'Fuente de Financiamiento
                            lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & Chr(10)
                            lsCadImp = lsCadImp & lsNegritaOn & sRotulo & Chr(10)
                            lsCadImp = lsCadImp & String(175 - 19, "-") & Chr(10)
                             
                            sTotal11 = 0
                            sTotal21 = 0
                            sTotal31 = 0
                            sTotal41 = 0
                            sTotal51 = 0
                            sTotal61 = 0
                            
                            nItem = 1
                            
                            sTempoFteFinancia = lrDataRep!cCodFuente
                            lnLineas = lnLineas + 6
                                
                        Else
                            sTotal11 = sTotal11 + !nMontoCol
                            sTotal21 = sTotal21 + !SaldoCapital
                            sTotal31 = sTotal31 + !Interes
                            sTotal41 = sTotal41 + !Mora
                            sTotal51 = sTotal51 + !Gasto
                            sTotal61 = sTotal61 + !SaldoCapital + !Interes + !Mora + !Gasto
                         End If
                    End If
                End If
                
                sCliente = !cpersNombre
                
                lsCadImp = lsCadImp & ImpreFormat(nItem, 4, 0) & Space(1) & !cCtaCod & _
                            Space(1) & Format(!dIngRecup, "dd/MM/YYYY") & Space(1) & _
                            sCliente & Space(1) & !Analista & Space(1) & Space(1) & _
                            ImpreFormat(!nMontoCol, 10, 2) & Space(1) & ImpreFormat(!SaldoCapital, 10, 2) & Space(1) & _
                            ImpreFormat(!Interes, 9, 2) & Space(1) & ImpreFormat(!Mora, 9, 2) & Space(1) & _
                            ImpreFormat(!Gasto, 9, 2) & Space(1) & ImpreFormat(!SaldoCapital + !Interes + !Mora + !Gasto, 10, 2) & Space(2) & !sExpe & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS EN JUDICIAL ", lsCondiciones, lnPage, 175 - 19, "", 138012)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            RaiseEvent CloseProgress
        End With
    lsCadImp = lsCadImp & String(175 - 19, "-") & Chr(10)
    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61 - 23) & ImpreFormat(sTotal11, 10, 2) & _
                Space(1) & ImpreFormat(sTotal21, 10, 2) & Space(1) & ImpreFormat(sTotal31, 9, 2) & _
                Space(1) & ImpreFormat(sTotal41, 9, 2) & Space(1) & ImpreFormat(sTotal51, 9, 2) & _
                Space(1) & ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & Chr(10)
    
    lsCadImp = lsCadImp & String(175 - 19, "-") & Chr(10)
    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X AGENCIA" & Space(78 - 23) & ImpreFormat(sTotal1, 10, 2) & _
                Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 9, 2) & _
                Space(1) & ImpreFormat(sTotal4, 9, 2) & Space(1) & ImpreFormat(sTotal5, 9, 2) & _
                Space(1) & ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & Chr(10)
    
    lsCadImp = lsCadImp & String(175 - 19, "-") & Chr(10)
    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(79 - 23) & ImpreFormat(sSubTotal1, 10, 2) & _
                Space(1) & ImpreFormat(sSubTotal2, 10, 2) & Space(1) & ImpreFormat(sSubTotal3, 9, 2) & _
                Space(1) & ImpreFormat(sSubTotal4, 9, 2) & Space(1) & ImpreFormat(sSubTotal5, 9, 2) & _
                Space(1) & ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)

        
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
    nRepo138012_ListadoCreditosEnJudicial = lsCadBuffer
End Function

Public Function nRepo138014_ListadoCreditosxAnalista(ByVal sConAnalista As Byte, ByVal sCodAnalista As String, ByVal sTipC As Double) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As NColPImpre
        
Dim sCliente As String * 30
Dim sAbogado As String * 30
Dim lsCondiciones As String
Dim sCriterio As String
Dim sTempoTipoCredito As String
Dim sTempoAnalista As String
Dim sTempoMoneda As String
Dim sProducto As String * 20

Dim lsNegritaOn As String
Dim lsNegritaOff As String

'Totales por Analista
Dim sSubTotal1 As Currency
Dim sSubTotal2 As Currency
Dim sSubTotal3 As Currency
Dim sSubTotal4 As Currency
Dim sSubTotal5 As Currency
Dim sSubTotal6 As Currency

'Totales por Tipo de Credito
Dim sTotal1 As Currency
Dim sTotal2 As Currency
Dim sTotal3 As Currency
Dim sTotal4 As Currency
Dim sTotal5 As Currency
Dim sTotal6 As Currency
   
'Totales por Moneda
Dim sTotalA As Currency
Dim sTotalB As Currency
Dim sTotalC As Currency
Dim sTotalD As Currency
Dim sTotalE As Currency
Dim sTotalF As Currency
  
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

    'Si se busca analista
    If sConAnalista = 1 Then
        
        sCriterio = " AND ((SELECT RHH.CUSER FROM RRHH RHH INNER JOIN " & _
                    " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD " & _
                    " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ")='" & Trim(sCodAnalista) & "')"
        lsCondiciones = "Analista: " & sCodAnalista & " : "
        
    End If
    
  '  lsSQL = "SELECT P.cCtaCod, substring(P.cCtaCod, 6,1) as cMoneda, C.dVigencia, CR.dIngRecup, C.cLineaCred, CLC.cDescripcion, SUBSTRING(P.cCtaCod, 6,3) AS cCodTipoCredito, Persona.cPersCod, " & _
            " Persona.cPersNombre, Persona.cPersDireccUbiGeo, (SELECT PrP.cPersCod FROM ProductoPersona PrP WHERE PrP.cCtaCod = PP.cCtaCod AND nPrdPersRelac=" & gColRelPersAnalista & ") AS CodAnalista, (SELECT PERSO.cPersNombre " & _
            " FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS ANALISTA, " & _
            " (SELECT PERSO.cPersNombre FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD " & _
            " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ") AS ABOGADO, C.nMontoCol AS Prestamo, P.nSaldo AS Saldo, " & _
            " CR.nSaldoIntComp AS Interes, CR.nSaldoIntMor AS Mora, CR.nSaldoGasto AS Gasto, SUBSTRING(C.cLineaCred, 5, 1) AS Moneda, " & _
            " CON.cConsDescripcion AS sMoneda, (SELECT COO.cConsDescripcion FROM Constante COO " & _
            " WHERE substring(P.cCtaCod, 6,3) = COO.nConsValor AND COO.nConsCod=" & gProducto & ") AS cDesTipoCredito, PP.nPrdPersRelac " & _
            " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod INNER JOIN " & _
            " ColocRecup CR ON C.cCtaCod = CR.cCtaCod INNER JOIN ColocLineaCredito CLC ON C.cLineaCred = CLC.cLineaCred INNER JOIN " & _
            " ProductoPersona PP ON P.cCtaCod = PP.cCtaCod INNER JOIN " & _
            " Persona ON PP.cPersCod = Persona.cPersCod INNER JOIN Constante CON ON SUBSTRING(C.cLineaCred, 5, 1) = CON.nConsValor " & _
            " Where (PP.NPRDPERSRELAC=" & gColRelPersTitular & ") And (CON.nconscod=" & gMoneda & ") AND (P.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "', '" & gColocEstRecVigCast & "', '" & gColocEstRecCanCast & "')) " & _
            sCriterio & " ORDER BY (SELECT PrP.cPersCod FROM ProductoPersona PrP WHERE PrP.cCtaCod = PP.cCtaCod AND nPrdPersRelac=" & gColRelPersAnalista & "),  SUBSTRING(P.cCtaCod, 6,3), SUBSTRING(P.cCtaCod, 6, 1), P.cCtaCod"

lsSQL = "SELECT P.cCtaCod, substring(P.cCtaCod, 9,1) as cMoneda, C.dVigencia, CR.dIngRecup, C.cLineaCred, CLC.cDescripcion, SUBSTRING(P.cCtaCod, 6,3) AS cCodTipoCredito, Persona.cPersCod, " & _
            " Persona.cPersNombre, Persona.cPersDireccUbiGeo, (SELECT PrP.cPersCod FROM ProductoPersona PrP WHERE PrP.cCtaCod = PP.cCtaCod AND nPrdPersRelac=" & gColRelPersAnalista & ") AS CodAnalista, (SELECT PERSO.cPersNombre " & _
            " FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS ANALISTA, " & _
            " (SELECT PERSO.cPersNombre FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD " & _
            " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ") AS ABOGADO, C.nMontoCol AS Prestamo, P.nSaldo AS Saldo, " & _
            " CR.nSaldoIntComp AS Interes, CR.nSaldoIntMor AS Mora, CR.nSaldoGasto AS Gasto, SUBSTRING(P.cCtaCod, 9,1) AS Moneda, " & _
            " CON.cConsDescripcion AS sMoneda, (SELECT COO.cConsDescripcion FROM Constante COO " & _
            " WHERE substring(P.cCtaCod, 6,3) = COO.nConsValor AND COO.nConsCod=" & gProducto & ") AS cDesTipoCredito, PP.nPrdPersRelac " & _
            " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod INNER JOIN " & _
            " ColocRecup CR ON C.cCtaCod = CR.cCtaCod INNER JOIN ColocLineaCredito CLC ON C.cLineaCred = CLC.cLineaCred INNER JOIN " & _
            " ProductoPersona PP ON P.cCtaCod = PP.cCtaCod INNER JOIN " & _
            " Persona ON PP.cPersCod = Persona.cPersCod INNER JOIN Constante CON ON  substring(P.cCtaCod, 9,1) = CON.nConsValor " & _
            " Where (PP.NPRDPERSRELAC=" & gColRelPersTitular & ") And (CON.nconscod=" & gMoneda & ") AND (P.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "', '" & gColocEstRecVigCast & "', '" & gColocEstRecCanCast & "')) " & _
            sCriterio & " ORDER BY (SELECT PrP.cPersCod FROM ProductoPersona PrP WHERE PrP.cCtaCod = PP.cCtaCod AND nPrdPersRelac=" & gColRelPersAnalista & "),  SUBSTRING(P.cCtaCod, 6,3), SUBSTRING(P.cCtaCod, 9, 1), P.cCtaCod"


    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
     
        If sConAnalista = 1 Then
            lsCondiciones = lsCondiciones & "" & lrDataRep!Analista
        End If
        
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS X ANALISTA ", lsCondiciones, lnPage, 193 - 12, "//T.C. :" & Format(sTipC, "0.00"), 138014)

        lnIndice = 0:  lnLineas = 7
        sTempoTipoCredito = lrDataRep!cCodTipoCredito
        sTempoAnalista = "" & lrDataRep!CodAnalista
        sTempoMoneda = "" & lrDataRep!cMoneda
        
        'Analista
        lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA: " & lrDataRep!Analista & " : " & lrDataRep!CodAnalista & lsNegritaOff & Chr(10) & Chr(10)
        'Producto
        lsCadImp = lsCadImp & lsNegritaOn & "TIPO DE CREDITO: " & lrDataRep!cDesTipoCredito & lsNegritaOff & Chr(10) & Chr(10)
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
        
        lnLineas = lnLineas + 5
        lnIndice = lnIndice + 5
          
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                'Totalizar por analista
                
                If sTempoAnalista <> !CodAnalista Then
                
                    lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & ImpreFormat(sTotalA, 10, 2) & _
                                Space(1) & ImpreFormat(sTotalB, 10, 2) & Space(1) & ImpreFormat(sTotalC, 10, 2) & _
                                Space(1) & ImpreFormat(sTotalD, 10, 2) & Space(1) & ImpreFormat(sTotalE, 10, 2) & _
                                Space(1) & ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                    
                    lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO DE CREDITO" & Space(54 - 12) & ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                    
                    lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ANALISTA" & Space(61 - 12) & ImpreFormat(sSubTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sSubTotal2, 10, 2) & Space(1) & ImpreFormat(sSubTotal3, 10, 2) & _
                                Space(1) & ImpreFormat(sSubTotal4, 10, 2) & Space(1) & ImpreFormat(sSubTotal5, 10, 2) & _
                                Space(1) & ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                     
                    'Analista
                    lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA: " & lrDataRep!Analista & " : " & lrDataRep!CodAnalista & lsNegritaOff & Chr(10) & Chr(10)
                    
                    'Producto
                    lsCadImp = lsCadImp & lsNegritaOn & "TIPO DE CREDITO: " & !cDesTipoCredito & lsNegritaOff & Chr(10) & Chr(10)
                    
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & Chr(10)
                    lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                    
                    
                    sSubTotal1 = (!Prestamo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal2 = (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal3 = (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal4 = (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal5 = (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal6 = (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                    
                    sTotal1 = (!Prestamo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal2 = (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal3 = (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal4 = (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal5 = (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal6 = (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                    
                    sTotalA = !Prestamo
                    sTotalB = !Saldo
                    sTotalC = !Interes
                    sTotalD = !Mora
                    sTotalE = !Gasto
                    sTotalF = !Saldo + !Interes + !Mora + !Gasto
                    
                    nItem = 1
                    sTempoTipoCredito = lrDataRep!cCodTipoCredito
                    sTempoAnalista = "" & lrDataRep!CodAnalista
                    sTempoMoneda = "" & lrDataRep!cMoneda
                      
                    lnLineas = lnLineas + 12
                    lnIndice = lnIndice + 12
                     
                Else
                                        
                    sSubTotal1 = sSubTotal1 + (!Prestamo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal2 = sSubTotal2 + (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal3 = sSubTotal3 + (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal4 = sSubTotal4 + (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal5 = sSubTotal5 + (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal6 = sSubTotal6 + (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                    
                    
                    If sTempoTipoCredito <> lrDataRep!cCodTipoCredito Then
                        lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & ImpreFormat(sTotalA, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotalB, 10, 2) & Space(1) & ImpreFormat(sTotalC, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotalD, 10, 2) & Space(1) & ImpreFormat(sTotalE, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                        
                        lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO DE CREDITO" & Space(54 - 12) & ImpreFormat(sTotal1, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                        
                        'Producto
                        lsCadImp = lsCadImp & lsNegritaOn & "TIPO DE CREDITO: " & !cDesTipoCredito & lsNegritaOff & Chr(10) & Chr(10)
                        
                        'Moneda
                        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & Chr(10)
                        lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                        
                        
                        sTotal1 = (!Prestamo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal2 = (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal3 = (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal4 = (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal5 = (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal6 = (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                        
                        sTotalA = !Prestamo
                        sTotalB = !Saldo
                        sTotalC = !Interes
                        sTotalD = !Mora
                        sTotalE = !Gasto
                        sTotalF = !Saldo + !Interes + !Mora + !Gasto
                        
                        nItem = 1
                        sTempoTipoCredito = lrDataRep!cCodTipoCredito
                        sTempoMoneda = "" & lrDataRep!cMoneda
                          
                        lnLineas = lnLineas + 9
                        lnIndice = lnIndice + 9
                    Else
                    
                    '***************************
                
                        sTotal1 = sTotal1 + (!Prestamo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal2 = sTotal2 + (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal3 = sTotal3 + (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal4 = sTotal4 + (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal5 = sTotal5 + (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal6 = sTotal6 + ((!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                 
                        If sTempoMoneda <> !cMoneda Then
                
                            lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                            lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & ImpreFormat(sTotalA, 10, 2) & _
                                        Space(1) & ImpreFormat(sTotalB, 10, 2) & Space(1) & ImpreFormat(sTotalC, 10, 2) & _
                                        Space(1) & ImpreFormat(sTotalD, 10, 2) & Space(1) & ImpreFormat(sTotalE, 10, 2) & _
                                        Space(1) & ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                        
                            'Moneda
                            lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & Chr(10)
                            lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
 
                            sTotalA = !Prestamo
                            sTotalB = !Saldo
                            sTotalC = !Interes
                            sTotalD = !Mora
                            sTotalE = !Gasto
                            sTotalF = !Saldo + !Interes + !Mora + !Gasto
                            nItem = 1
                            sTempoMoneda = lrDataRep!cMoneda
                            lnLineas = lnLineas + 3
                            lnIndice = lnIndice + 3
                        Else
                            sTotalA = sTotalA + !Prestamo
                            sTotalB = sTotalB + !Saldo
                            sTotalC = sTotalC + !Interes
                            sTotalD = sTotalD + !Mora
                            sTotalE = sTotalE + !Gasto
                            sTotalF = sTotalF + !Saldo + !Interes + !Mora + !Gasto
                        End If
                    End If
                End If
                
                sCliente = !cpersNombre
                sAbogado = "" & !Abogado
                sProducto = !cDesTipoCredito
                
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(nItem, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dIngRecup, "dd/MM/YYYY") & _
                              Space(1) & _
                            sCliente & Space(1) & _
                            ImpreFormat(!Prestamo, 10, 2) & Space(1) & ImpreFormat(!Saldo, 10, 2) & Space(1) & _
                            ImpreFormat(!Interes, 10, 2) & Space(1) & ImpreFormat(!Mora, 10, 2) & Space(1) & _
                            ImpreFormat(!Gasto, 10, 2) & Space(1) & ImpreFormat(!Saldo + !Interes + !Mora + !Gasto, 10, 2) & Space(1) & _
                            sAbogado & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 53 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS X ANALISTA ", lsCondiciones, lnPage, 193 - 12, "//T.C. :" & Format(sTipC, "0.00"), 138014)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & ImpreFormat(sTotalA, 10, 2) & _
                    Space(1) & ImpreFormat(sTotalB, 10, 2) & Space(1) & ImpreFormat(sTotalC, 10, 2) & _
                    Space(1) & ImpreFormat(sTotalD, 10, 2) & Space(1) & ImpreFormat(sTotalE, 10, 2) & _
                    Space(1) & ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
        
        lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO DE CREDITO" & Space(54 - 12) & ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
        
        lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ANALISTA" & Space(61 - 12) & ImpreFormat(sSubTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal2, 10, 2) & Space(1) & ImpreFormat(sSubTotal3, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal4, 10, 2) & Space(1) & ImpreFormat(sSubTotal5, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
 
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
    nRepo138014_ListadoCreditosxAnalista = lsCadBuffer
End Function

Public Function nRepo138021_ListadoCreditosEnCJxAbogado(ByVal sConAbogado As Byte, ByVal sCodAbogado As String, ByVal sTipC As Double) As String
  
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
Dim sCriterio As String
Dim sTempoAbogado As String
Dim sDescripcion As String * 25

Dim lsNegritaOn As String
Dim lsNegritaOff As String

'Totales por Abogado
Dim sTotal1 As Currency
Dim sTotal2 As Currency
Dim sTotal3 As Currency
Dim sTotal4 As Currency
Dim sTotal5 As Currency
Dim sTotal6 As Currency

'Totales Generales
Dim sTotalA As Currency
Dim sTotalB As Currency
Dim sTotalC As Currency
Dim sTotalD As Currency
Dim sTotalE As Currency
Dim sTotalF As Currency
Dim nTipoCambio As Double
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

    'Si se busca abogado
    If sConAbogado = 1 Then
        
        sCriterio = " AND ((SELECT PRP.CPERSCOD FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD " & _
                  " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ")='" & Trim(sCodAbogado) & "')"
 
        lsCondiciones = "Abogado:  " & sCodAbogado & " : "
        
    End If

lsSQL = " select sum(qr.prestamo) as Prestamo, sum(qr.Saldo) as Saldo, sum(qr.Interes) as Interes, " & _
        " sum(qr.Mora) as Mora, sum(qr.Gasto) as Gasto, sum(qr.Total) as Total, qr.Codabogado , qr.abogado, " & _
        " qr.nPrdEstado, qr.desestadorecuperacion from( " & _
        " SELECT C.nMontoCol * case when substring(c.cCtacod,9,1)=1 then 1 else " & sTipC & " end AS Prestamo, P.nSaldo * case when substring(c.cCtacod,9,1)=1 then 1 else " & sTipC & " end AS Saldo, CR.nSaldoIntComp * case when substring(c.cCtacod,9,1)=1 then 1 else " & sTipC & " end AS Interes, CR.nSaldoIntMor * case when substring(c.cCtacod,9,1)=1 then 1 else " & sTipC & " end AS Mora, " & _
        " CR.nSaldoGasto * case when substring(c.cCtacod,9,1)=1 then 1 else " & sTipC & " end AS Gasto, (P.nSaldo + CR.nSaldoIntComp + CR.nSaldoIntMor + CR.nSaldoGasto) * case when substring(c.cCtacod,9,1)=1 then 1 else " & sTipC & " end AS Total, " & _
        " (SELECT PRP.CPERSCOD FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD " & _
        " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ") AS CODABOGADO, (SELECT PERSO.cPersNombre " & _
        " FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD " & _
        " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ") AS ABOGADO, P.nPrdEstado, (SELECT K.cConsDescripcion " & _
        " FROM constante K WHERE k.nconscod = '3001' AND k.nconsvalor = P.nPrdEstado) AS DesEstadoRecuperacion " & _
        " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod INNER JOIN ColocRecup CR ON " & _
        " C.cCtaCod = CR.cCtaCod INNER JOIN ProductoPersona PP ON P.cCtaCod = PP.cCtaCod  "
        '" INNER JOIN Persona ON PP.cPersCod = Persona.cPersCod LEFT OUTER JOIN " & _
        " ColocRecupExpedienteInf ON P.cCtaCod = ColocRecupExpedienteInf.cCtaCod "

lsSQL = lsSQL & " WHERE (PP.nPrdPersRelac=" & gColRelPersTitular & ") AND (P.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "', '" & gColocEstRecVigCast & "', '" & gColocEstRecCanCast & "')) " & _
        sCriterio & _
        " ) QR  group by qr.Codabogado, qr.Abogado, qr.nPrdEstado, qr.DesEstadoRecuperacion " & _
        " Order By qr.abogado, qr.desestadorecuperacion desc "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
     
        If sConAbogado = 1 Then
            lsCondiciones = lsCondiciones & "" & lrDataRep!Abogado
        End If
        
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CREDITOS EN C.J. POR ABOGADO(Consolidado)", lsCondiciones, lnPage, 109, "//T.C. :" & Format(sTipC, "0.00"), 138021)

        lnIndice = 0:  lnLineas = 7
        sTempoAbogado = "" & lrDataRep!CODABOGADO
        
        'Analista
        lsCadImp = lsCadImp & lsNegritaOn & "ABOGADO:  " & lrDataRep!Abogado & " : " & lrDataRep!CODABOGADO & lsNegritaOff & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & String(109, "-") & Chr(10)
        
        lnLineas = lnLineas + 3
        lnIndice = lnIndice + 3
          
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                  
                'Totalizar por analista
                
                If sTempoAbogado <> !CODABOGADO Then
                
                        lsCadImp = lsCadImp & String(109, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ABOG." & Space(11) & ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                    
                    'Abogado
                    lsCadImp = lsCadImp & lsNegritaOn & "ABOGADO:  " & lrDataRep!Abogado & " : " & lrDataRep!CODABOGADO & lsNegritaOff & Chr(10) & Chr(10)
                    lsCadImp = lsCadImp & String(109, "-") & Chr(10)
                      
                    sTotal1 = !Prestamo
                    sTotal2 = !Saldo
                    sTotal3 = !Interes
                    sTotal4 = !Mora
                    sTotal5 = !Gasto
                    sTotal6 = !Saldo + !Interes + !Mora + !Gasto
                    
                    sTempoAbogado = lrDataRep!CODABOGADO
                      
                    lnLineas = lnLineas + 6
                    lnIndice = lnIndice + 6
                Else
                     
                    sTotal1 = sTotal1 + !Prestamo
                    sTotal2 = sTotal2 + !Saldo
                    sTotal3 = sTotal3 + !Interes
                    sTotal4 = sTotal4 + !Mora
                    sTotal5 = sTotal5 + !Gasto
                    sTotal6 = sTotal6 + !Saldo + !Interes + !Mora + !Gasto
                End If
                
                sTotalA = sTotalA + !Prestamo
                sTotalB = sTotalB + !Saldo
                sTotalC = sTotalC + !Interes
                sTotalD = sTotalD + !Mora
                sTotalE = sTotalE + !Gasto
                sTotalF = sTotalF + !Saldo + !Interes + !Mora + !Gasto
                
                sDescripcion = Replace(!DesEstadoRecuperacion, "VIGENTE ", "")
                
                lsCadImp = lsCadImp & sDescripcion & Space(1) & _
                            ImpreFormat(!Prestamo, 10, 2) & Space(1) & ImpreFormat(!Saldo, 10, 2) & Space(1) & _
                            ImpreFormat(!Interes, 10, 2) & Space(1) & ImpreFormat(!Mora, 10, 2) & Space(1) & _
                            ImpreFormat(!Gasto, 10, 2) & Space(1) & ImpreFormat(!Saldo + !Interes + !Mora + !Gasto, 10, 2) & Chr(10)
                
                'lnLineas = lnLineas + 1
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS EN C.J. POR ABOGADO(Consolidado)", lsCondiciones, lnPage, 109, "//T.C. :" & Format(sTipC, "0.00"), 138021)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(109, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ABOG." & Space(11) & ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & Chr(10)
        
        lsCadImp = lsCadImp & String(109, "=") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTAL          " & Space(11) & ImpreFormat(sTotalA, 10, 2) & _
                    Space(1) & ImpreFormat(sTotalB, 10, 2) & Space(1) & ImpreFormat(sTotalC, 10, 2) & _
                    Space(1) & ImpreFormat(sTotalD, 10, 2) & Space(1) & ImpreFormat(sTotalE, 10, 2) & _
                    Space(1) & ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & Chr(10)
    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
    nRepo138021_ListadoCreditosEnCJxAbogado = lsCadBuffer
End Function
Public Function nRepo138024_ListadoExpedientesxAbogado(ByVal sConAbogado As Byte, ByVal sCodAbogado As String, ByVal sTipC As Double) As String
 
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As NColPImpre
        
Dim sCliente As String * 30
Dim sAbogado As String * 30
Dim lsCondiciones As String
Dim sCriterio As String
Dim sTempoTipoCredito As String
Dim sTempoAnalista As String
Dim sTempoMoneda As String
Dim sProducto As String * 20

Dim lsNegritaOn As String
Dim lsNegritaOff As String

'Totales por Analista
Dim sSubTotal1 As Currency
Dim sSubTotal2 As Currency
Dim sSubTotal3 As Currency
Dim sSubTotal4 As Currency
Dim sSubTotal5 As Currency
Dim sSubTotal6 As Currency

'Totales por Tipo de Credito
Dim sTotal1 As Currency
Dim sTotal2 As Currency
Dim sTotal3 As Currency
Dim sTotal4 As Currency
Dim sTotal5 As Currency
Dim sTotal6 As Currency
   
'Totales por Moneda
Dim sTotalA As Currency
Dim sTotalB As Currency
Dim sTotalC As Currency
Dim sTotalD As Currency
Dim sTotalE As Currency
Dim sTotalF As Currency

   
  
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

    'Si se busca analista
    If sConAbogado = 1 Then
        
        sCriterio = " AND ((SELECT RHH.CUSER FROM RRHH RHH INNER JOIN " & _
                    " PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD " & _
                    " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ")='" & Trim(sCodAbogado) & "')"
        lsCondiciones = "Estudio: " & sCodAbogado & " : "
        
    End If
    
    lsSQL = "SELECT P.cCtaCod, substring(P.cCtaCod, 9,1) as cMoneda, C.dVigencia, CR.dIngRecup, C.cLineaCred, CLC.cDescripcion, SUBSTRING(P.cCtaCod, 6,3) AS cCodTipoCredito, Persona.cPersCod, " & _
            " Persona.cPersNombre, Persona.cPersDireccUbiGeo, (SELECT PrP.cPersCod FROM ProductoPersona PrP WHERE PrP.cCtaCod = PP.cCtaCod AND nPrdPersRelac=" & gColRelPersEstudioJuridico & ") AS CodAnalista, (SELECT PERSO.cPersNombre " & _
            " FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ") AS ANALISTA, " & _
            " (SELECT PERSO.cPersNombre FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD " & _
            " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS ABOGADO, C.nMontoCol AS Prestamo, P.nSaldo AS Saldo, " & _
            " CR.nSaldoIntComp AS Interes, CR.nSaldoIntMor AS Mora, CR.nSaldoGasto AS Gasto, SUBSTRING(P.cCtaCod, 9,1) AS Moneda, " & _
            " CON.cConsDescripcion AS sMoneda, (SELECT COO.cConsDescripcion FROM Constante COO " & _
            " WHERE substring(P.cCtaCod, 6,3) = COO.nConsValor AND COO.nConsCod=" & gProducto & ") AS cDesTipoCredito, PP.nPrdPersRelac " & _
            " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod INNER JOIN " & _
            " ColocRecup CR ON C.cCtaCod = CR.cCtaCod INNER JOIN ColocLineaCredito CLC ON C.cLineaCred = CLC.cLineaCred INNER JOIN " & _
            " ProductoPersona PP ON P.cCtaCod = PP.cCtaCod INNER JOIN " & _
            " Persona ON PP.cPersCod = Persona.cPersCod INNER JOIN Constante CON ON  substring(P.cCtaCod, 9,1) = CON.nConsValor " & _
            " Where (PP.NPRDPERSRELAC=" & gColRelPersTitular & ") And (CON.nconscod=" & gMoneda & ") AND (P.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "', '" & gColocEstRecVigCast & "', '" & gColocEstRecCanCast & "')) " & _
            sCriterio & " ORDER BY (SELECT PrP.cPersCod FROM ProductoPersona PrP WHERE PrP.cCtaCod = PP.cCtaCod AND nPrdPersRelac=" & gColRelPersEstudioJuridico & "),  SUBSTRING(P.cCtaCod, 6,3), SUBSTRING(P.cCtaCod, 9, 1), P.cCtaCod"


    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
     
        If sConAbogado = 1 Then
            lsCondiciones = lsCondiciones & "" & lrDataRep!Analista
        End If
        
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CREDITOS EN C.J. POR ABOGADO(Detalle)", lsCondiciones, lnPage, 193 - 12, "//T.C. :" & Format(sTipC, "0.00"), 138024)
                                             
        lnIndice = 0:  lnLineas = 7
        sTempoTipoCredito = lrDataRep!cCodTipoCredito
        sTempoAnalista = "" & lrDataRep!CodAnalista
        sTempoMoneda = "" & lrDataRep!cMoneda
        
        'Analista
        lsCadImp = lsCadImp & lsNegritaOn & "ABOGADO: " & lrDataRep!Analista & " : " & lrDataRep!CodAnalista & lsNegritaOff & Chr(10) & Chr(10)
        'Producto
        lsCadImp = lsCadImp & lsNegritaOn & "TIPO DE CREDITO: " & lrDataRep!cDesTipoCredito & lsNegritaOff & Chr(10) & Chr(10)
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
        
        lnLineas = lnLineas + 5
        lnIndice = lnIndice + 5
          
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                'Totalizar por analista
                
                If sTempoAnalista <> !CodAnalista Then
                
                    lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & ImpreFormat(sTotalA, 10, 2) & _
                                Space(1) & ImpreFormat(sTotalB, 10, 2) & Space(1) & ImpreFormat(sTotalC, 10, 2) & _
                                Space(1) & ImpreFormat(sTotalD, 10, 2) & Space(1) & ImpreFormat(sTotalE, 10, 2) & _
                                Space(1) & ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                    
                    lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO DE CREDITO" & Space(54 - 12) & ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                                Space(1) & ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                    
                    lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ABOGADO " & Space(61 - 12) & ImpreFormat(sSubTotal1, 10, 2) & _
                                Space(1) & ImpreFormat(sSubTotal2, 10, 2) & Space(1) & ImpreFormat(sSubTotal3, 10, 2) & _
                                Space(1) & ImpreFormat(sSubTotal4, 10, 2) & Space(1) & ImpreFormat(sSubTotal5, 10, 2) & _
                                Space(1) & ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                     
                    'Analista
                    lsCadImp = lsCadImp & lsNegritaOn & "ABOGADO: " & lrDataRep!Analista & " : " & lrDataRep!CodAnalista & lsNegritaOff & Chr(10) & Chr(10)
                    
                    'Producto
                    lsCadImp = lsCadImp & lsNegritaOn & "TIPO DE CREDITO: " & !cDesTipoCredito & lsNegritaOff & Chr(10) & Chr(10)
                    
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & Chr(10)
                    lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                    
                    
                    sSubTotal1 = (!Prestamo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal2 = (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal3 = (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal4 = (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal5 = (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal6 = (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                    
                    sTotal1 = (!Prestamo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal2 = (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal3 = (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal4 = (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal5 = (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal6 = (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                    
                    sTotalA = !Prestamo
                    sTotalB = !Saldo
                    sTotalC = !Interes
                    sTotalD = !Mora
                    sTotalE = !Gasto
                    sTotalF = !Saldo + !Interes + !Mora + !Gasto
                    
                    nItem = 1
                    sTempoTipoCredito = lrDataRep!cCodTipoCredito
                    sTempoAnalista = "" & lrDataRep!CodAnalista
                    sTempoMoneda = "" & lrDataRep!cMoneda
                      
                    lnLineas = lnLineas + 12
                    lnIndice = lnIndice + 12
                     
                Else
                                        
                    sSubTotal1 = sSubTotal1 + (!Prestamo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal2 = sSubTotal2 + (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal3 = sSubTotal3 + (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal4 = sSubTotal4 + (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal5 = sSubTotal5 + (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal6 = sSubTotal6 + (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                    
                    
                    If sTempoTipoCredito <> lrDataRep!cCodTipoCredito Then
                        lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & ImpreFormat(sTotalA, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotalB, 10, 2) & Space(1) & ImpreFormat(sTotalC, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotalD, 10, 2) & Space(1) & ImpreFormat(sTotalE, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                        
                        lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO DE CREDITO" & Space(54 - 12) & ImpreFormat(sTotal1, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                                    Space(1) & ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                        
                        'Producto
                        lsCadImp = lsCadImp & lsNegritaOn & "TIPO DE CREDITO: " & !cDesTipoCredito & lsNegritaOff & Chr(10) & Chr(10)
                        
                        'Moneda
                        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & Chr(10)
                        lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                        
                        
                        sTotal1 = (!Prestamo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal2 = (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal3 = (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal4 = (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal5 = (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal6 = (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                        
                        sTotalA = !Prestamo
                        sTotalB = !Saldo
                        sTotalC = !Interes
                        sTotalD = !Mora
                        sTotalE = !Gasto
                        sTotalF = !Saldo + !Interes + !Mora + !Gasto
                        
                        nItem = 1
                        sTempoTipoCredito = lrDataRep!cCodTipoCredito
                        sTempoMoneda = "" & lrDataRep!cMoneda
                          
                        lnLineas = lnLineas + 9
                        lnIndice = lnIndice + 9
                    Else
                    
                    '***************************
                
                        sTotal1 = sTotal1 + (!Prestamo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal2 = sTotal2 + (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal3 = sTotal3 + (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal4 = sTotal4 + (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal5 = sTotal5 + (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal6 = sTotal6 + ((!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                 
                        If sTempoMoneda <> !cMoneda Then
                
                            lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
                            lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & ImpreFormat(sTotalA, 10, 2) & _
                                        Space(1) & ImpreFormat(sTotalB, 10, 2) & Space(1) & ImpreFormat(sTotalC, 10, 2) & _
                                        Space(1) & ImpreFormat(sTotalD, 10, 2) & Space(1) & ImpreFormat(sTotalE, 10, 2) & _
                                        Space(1) & ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                        
                            'Moneda
                            lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & Chr(10)
                            lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
 
                            sTotalA = !Prestamo
                            sTotalB = !Saldo
                            sTotalC = !Interes
                            sTotalD = !Mora
                            sTotalE = !Gasto
                            sTotalF = !Saldo + !Interes + !Mora + !Gasto
                            nItem = 1
                            sTempoMoneda = lrDataRep!cMoneda
                            lnLineas = lnLineas + 3
                            lnIndice = lnIndice + 3
                        Else
                            sTotalA = sTotalA + !Prestamo
                            sTotalB = sTotalB + !Saldo
                            sTotalC = sTotalC + !Interes
                            sTotalD = sTotalD + !Mora
                            sTotalE = sTotalE + !Gasto
                            sTotalF = sTotalF + !Saldo + !Interes + !Mora + !Gasto
                        End If
                    End If
                End If
                
                sCliente = !cpersNombre
                sAbogado = "" & !Abogado
                sProducto = !cDesTipoCredito
                
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(nItem, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dIngRecup, "dd/MM/YYYY") & _
                              Space(1) & _
                            sCliente & Space(1) & _
                            ImpreFormat(!Prestamo, 10, 2) & Space(1) & ImpreFormat(!Saldo, 10, 2) & Space(1) & _
                            ImpreFormat(!Interes, 10, 2) & Space(1) & ImpreFormat(!Mora, 10, 2) & Space(1) & _
                            ImpreFormat(!Gasto, 10, 2) & Space(1) & ImpreFormat(!Saldo + !Interes + !Mora + !Gasto, 10, 2) & Space(1) & _
                            sAbogado & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 53 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS EN C.J. POR ABOGADO(Detalle)", lsCondiciones, lnPage, 193 - 12, "//T.C. :" & Format(sTipC, "0.00"), 138024)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                RaiseEvent Progress(.Bookmark, .RecordCount)
                 
                .MoveNext
                
            Loop
            RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & ImpreFormat(sTotalA, 10, 2) & _
                    Space(1) & ImpreFormat(sTotalB, 10, 2) & Space(1) & ImpreFormat(sTotalC, 10, 2) & _
                    Space(1) & ImpreFormat(sTotalD, 10, 2) & Space(1) & ImpreFormat(sTotalE, 10, 2) & _
                    Space(1) & ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
        
        lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO DE CREDITO" & Space(54 - 12) & ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal2, 10, 2) & Space(1) & ImpreFormat(sTotal3, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal4, 10, 2) & Space(1) & ImpreFormat(sTotal5, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
        
        lsCadImp = lsCadImp & String(193 - 12, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ABOGADO " & Space(61 - 12) & ImpreFormat(sSubTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal2, 10, 2) & Space(1) & ImpreFormat(sSubTotal3, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal4, 10, 2) & Space(1) & ImpreFormat(sSubTotal5, 10, 2) & _
                    Space(1) & ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
 
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
    nRepo138024_ListadoExpedientesxAbogado = lsCadBuffer
End Function
 
Public Function nRepo138023_ListadoSaldoCapital(ByVal sTope As Long, ByVal nTipoCambio As Currency, psEstados As String) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim nItem As Long
Dim lnPage As Integer
Dim loImprimeCab As NColPImpre
        
Dim cCliente As String * 30
Dim cEstado As String * 30
Dim lsCondiciones As String
  
Dim cMoneda As String

Dim sCriterio As String
 
Dim lsNegritaOn As String
Dim lsNegritaOff As String
  
'Totales
Dim sTotal1 As Currency
Dim sTotal2 As Currency
 
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)
     
    lsCondiciones = "SEGUN SALDO DE CAPITAL : TOP " & sTope
                
    'Busco por estados en la tabla Producto
    sCriterio = " AND (P.nPrdEstado IN(" & psEstados & ")) "
                
                
    lsSQL = " SELECT TOP " & sTope & "  C.cCtaCod AS cCuenta, PERS.cPersNombre, SUBSTRING(P.cCtaCod, 9,1) AS nMoneda, CON.cConsDescripcion AS cMoneda, " & _
            " (SELECT K.cConsDescripcion FROM constante K WHERE K.nconscod='" & gColocEstado & "' AND " & _
            " k.nconsvalor = P.nPrdEstado) AS cEstadoRecuperacion, C.dVigencia AS dFecVigencia, " & _
            " CR.dIngRecup AS dFecIngJud, C.nMontoCol AS nPrestamo, nSaldoCapitalSoles=case when SUBSTRING(P.cCtaCod, 9,1)=" & gMonedaNacional & " then" & _
            " P.nSaldo ELSE P.nSaldo*" & Val(nTipoCambio) & " END FROM ColocRecup CR INNER JOIN Colocaciones C ON CR.cCtaCod = C.cCtaCod INNER JOIN " & _
            " ProductoPersona PP ON CR.cCtaCod = PP.cCtaCod INNER JOIN Persona PERS ON PP.cPersCod = PERS.cPersCod INNER JOIN Producto P ON C.cCtaCod = " & _
            " P.cCtaCod AND PP.cCtaCod = P.cCtaCod INNER JOIN Constante CON ON SUBSTRING(P.cCtaCod, 9,1) = CON.nConsValor " & _
            " WHERE (PP.nPrdPersRelac=" & gColRelPersTitular & ") AND (CON.nConsCod='" & gMoneda & "') " & _
            sCriterio & _
            " ORDER BY substring(C.cCtaCod,9,1), case when SUBSTRING(P.cCtaCod, 9,1)=" & gMonedaNacional & " then" & _
            " P.nSaldo ELSE P.nSaldo*" & Val(nTipoCambio) & " END DESC, PERS.cPersNombre, c.cCtaCod "
 
        'nTiporecuperacion=1 si es VigenteJudicial(2201) o CanceladoJudicial(2203)
        'nTiporecuperacion=2 si es VigenteCastigado(2201) o CanceladoCastigado(2203)
         
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
         
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CREDITOS EN COBRANZA JUDICIAL CON SALDO DE CAPITAL MAYOR AL " & Format(gdFecSis, "dd/MM/YYYY"), lsCondiciones, lnPage, 144 - 15, "", 138023)

        cMoneda = lrDataRep!nMoneda
        lsCadImp = lsCadImp & "Moneda: " & IIf(lrDataRep!nMoneda = "1", "Soles", "Dolares") & Chr(10)

        lnIndice = 0:  lnLineas = 7
         
        With lrDataRep
            RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                
                If cMoneda <> !nMoneda And Len(Trim(cMoneda)) > 0 Then
                    lsCadImp = lsCadImp & String(144 - 15, "-") & Chr(10)
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES           " & Space(68) & ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal2, 10, 2) & lsNegritaOff & Chr(10) & Chr(10)
                    
                    sTotal1 = 0
                    sTotal2 = 0
                    nItem = 0
                    cMoneda = !nMoneda
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS EN COBRANZA JUDICIAL CON SALDO DE CAPITAL MAYOR AL " & Format(gdFecSis, "dd/MM/YYYY"), lsCondiciones, lnPage, 144 - 15, "", 138023)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                    lsCadImp = lsCadImp & "Moneda: " & IIf(!nMoneda = "1", "Soles", "Dolares") & Chr(10)
                End If
                
                nItem = nItem + 1
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                sTotal1 = sTotal1 + lrDataRep!nPrestamo
                sTotal2 = sTotal2 + lrDataRep!nSaldoCapitalSoles
                
                cEstado = Replace(!cEstadoRecuperacion, "Vigente ", "")
                cCliente = !cpersNombre
                 
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(nItem, 4, 0) & Space(1) & !cCuenta & Space(1) & _
                           Space(1) & cCliente & Space(1) & !dFecVigencia & Space(1) & !dFecIngJud & Space(4) & _
                           Mid(!cMoneda, 1, 1) & Space(3) & ImpreFormat(!nPrestamo, 10, 2) & Space(1) & _
                           ImpreFormat(!nSaldoCapitalSoles, 10, 2) & Space(1) & cEstado & Chr(10)
                 
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS EN COBRANZA JUDICIAL CON SALDO DE CAPITAL MAYOR AL " & Format(gdFecSis, "dd/MM/YYYY"), lsCondiciones, lnPage, 144 - 15, "", 138023)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(144 - 15, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES           " & Space(68) & ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & ImpreFormat(sTotal2, 10, 2) & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(144 - 15, "=") & Chr(10)
    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
    nRepo138023_ListadoSaldoCapital = lsCadBuffer
End Function


Public Function nRepo138025_ResumenMovimientoCJ(ByVal psFecha1 As String, ByVal psFecha2 As String, ByVal pnTipoCambio As Currency) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As NColPImpre
        
Dim sCliente As String * 30
Dim lsCondiciones As String
Dim sCriterio As String
 
Dim lsNegritaOn As String
Dim lsNegritaOff As String

Dim MatMonN(0 To 5, 1 To 2) As Currency 'Moneda Nacional 0=Saldo Inicial 1=Ingresos(Nuevos) 2=Pagos(Amortizaciones + CAncel)
                                        '                3=SAldo Final Total 4=Saldo Final con demanda  5=Saldo Final Sin Demanda
                                        '                1=Cantidad 2=Monto Sumado
Dim MatMonE(0 To 5, 1 To 2) As Currency 'Moneda Extranjera 0=Saldo Inicial 1=Ingresos(Nuevos) 2=Pagos(Amortizaciones + CAncel)
                                        '                3=SAldo Final Total 4=Saldo Final con demanda  5=Saldo Final Sin Demanda
                                        '                1=Cantidad 2=Monto Sumado

lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)
lsCondiciones = "Del " & psFecha1 & " al " & psFecha2

    '--Ingresos
    lsSQL = "SELECT  1 as sTipo, SUBSTRING(MovColDet.cCtaCod, 9, 1) AS sMoneda, " & _
    " COUNT(MovColDet.nMovNro) As Cantidad, SUM(MovColDet.nMonto) AS Total " & _
    " FROM MovColDet INNER JOIN " & _
    " MovCol ON MovColDet.nMovNro = MovCol.nMovNro AND MovColDet.cOpeCod = MovCol.cOpeCod AND " & _
    " MovColDet.cCtaCod = MovCol.cCtaCod AND MovColDet.nNroCalen = MovCol.nNroCalen INNER JOIN " & _
    " Mov ON MovCol.nMovNro = Mov.nMovNro INNER JOIN ColocRecup ON MovColDet.cCtaCod = ColocRecup.cCtaCod " & _
    " Where (Mov.nMovFlag=" & gMovFlagVigente & ") AND (MovColDet.cOpeCod='" & gColRecOpePasoARecup & "') AND (MovColDet.nPrdConceptoCod='" & gColRecConceptoCodCapital & "') " & _
    " And (CONVERT(VARCHAR(8),ColocRecup.dIngRecup,112)>='" & Format(psFecha1, "YYYYMMdd") & "' and CONVERT(VARCHAR(8),ColocRecup.dIngRecup,112)<='" & Format(psFecha2, "YYYYMMdd") & "') " & _
    " GROUP BY  SUBSTRING(MovColDet.cCtaCod, 9, 1) "
    
    '--Pagos
    lsSQL = lsSQL & " Union All " & _
    " SELECT 2 as sTipo, SUBSTRING(MovColDet.cCtaCod, 9, 1) AS sMoneda, " & _
    " COUNT(MovColDet.nMovNro) As Cantidad, SUM(MovColDet.nMonto) AS Total " & _
    " FROM MovColDet INNER JOIN " & _
    " MovCol ON MovColDet.nMovNro = MovCol.nMovNro AND MovColDet.cOpeCod = MovCol.cOpeCod AND " & _
    " MovColDet.cCtaCod = MovCol.cCtaCod AND MovColDet.nNroCalen = MovCol.nNroCalen INNER JOIN " & _
    " Mov ON MovCol.nMovNro = Mov.nMovNro " & _
    " Where (Mov.nMovFlag=" & gMovFlagVigente & ") AND (MovColDet.cOpeCod like '1302%' or MovColDet.cOpeCod like '1303%') And (MovColDet.nPrdConceptoCod='" & gColRecConceptoCodCapital & "') " & _
    " AND (SUBSTRING(MOV.cMovNro, 1, 8)>='" & Format(psFecha1, "YYYYMMdd") & "' and SUBSTRING(MOV.cMovNro, 1, 8)<='" & Format(psFecha2, "YYYYMMdd") & "') " & _
    " GROUP BY  SUBSTRING(MovColDet.cCtaCod, 9, 1) "
    
    '--Saldo
    lsSQL = lsSQL & " Union All " & _
    " SELECT 3 as sTipo, substring(cCtaCod,9,1) as sMoneda, COUNT(cCtaCod) AS Cantidad, SUM(nSaldo) AS Total From Producto " & _
    " WHERE (nPrdEstado IN('" & gColocEstRecVigJud & "', '" & gColocEstRecVigCast & "', '" & gColocEstRecCanJud & "', '" & gColocEstRecCanCast & "')) AND (Producto.nSaldo > 0) " & _
    " GROUP BY substring(cCtaCod,9,1) "

    '--Saldo Detallado CON demanda
    lsSQL = lsSQL & " Union All " & _
    " SELECT 4 as sTipo, substring(Producto.cCtaCod,9,1) as sMoneda, COUNT(Producto.cCtaCod) AS Cantidad, SUM(Producto.nSaldo) AS Total " & _
    " FROM Producto INNER JOIN ColocRecup ON Producto.cCtaCod = ColocRecup.cCtaCod " & _
    " WHERE (Producto.nPrdEstado IN('" & gColocEstRecVigJud & "', '" & gColocEstRecVigCast & "', '" & gColocEstRecCanJud & "', '" & gColocEstRecCanCast & "')) AND " & _
    " (Producto.nSaldo > 0) and (ColocRecup.nDemanda=" & gColRecDemandaSi & ") " & _
    " GROUP BY substring(Producto.cCtaCod,9,1), ColocRecup.nDemanda "
    
    '--Saldo Detallado SIN demanda
    lsSQL = lsSQL & " Union All " & _
    " SELECT 5 as sTipo, substring(Producto.cCtaCod,9,1) as sMoneda, COUNT(Producto.cCtaCod) AS Cantidad, SUM(Producto.nSaldo) AS Total " & _
    " FROM Producto INNER JOIN " & _
    " ColocRecup ON Producto.cCtaCod = ColocRecup.cCtaCod " & _
    " WHERE (Producto.nPrdEstado IN('" & gColocEstRecVigJud & "', '" & gColocEstRecVigCast & "', '" & gColocEstRecCanJud & "', '" & gColocEstRecCanCast & "')) AND (Producto.nSaldo > 0) and (ColocRecup.nDemanda=" & gColRecDemandaNo & ") " & _
    " GROUP BY substring(Producto.cCtaCod,9,1), ColocRecup.nDemanda "
   
    lsSQL = lsSQL & " order by sTipo , sMoneda "
   
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
      
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("RESUMEN DEL MOVIMIENTO DE COBRANZA JUDICIAL", lsCondiciones, lnPage, 99, "", 138025)

        lnIndice = 0:  lnLineas = 7
                 
        With lrDataRep
        RaiseEvent ShowProgress
        Do While Not .EOF
            'Saldo Inicial = SaldoFinal + Pagos - Ingresos
            If !sTipo = 1 Then
                If !sMoneda = gMonedaNacional Then
                    MatMonN(1, 1) = !cantidad
                    MatMonN(1, 2) = !Total
                    
                    MatMonN(0, 1) = MatMonN(0, 1) - !cantidad
                    MatMonN(0, 2) = MatMonN(0, 2) - !Total
                ElseIf !sMoneda = gMonedaExtranjera Then
                    MatMonE(1, 1) = !cantidad
                    MatMonE(1, 2) = !Total
                    
                    MatMonE(0, 1) = MatMonN(0, 1) - !cantidad
                    MatMonE(0, 2) = MatMonN(0, 2) - !Total
                End If
            ElseIf !sTipo = 2 Then
                If !sMoneda = gMonedaNacional Then
                    MatMonN(2, 1) = !cantidad
                    MatMonN(2, 2) = !Total
                    
                    MatMonN(0, 1) = MatMonN(0, 1) + !cantidad
                    MatMonN(0, 2) = MatMonN(0, 2) + !Total
                ElseIf !sMoneda = gMonedaExtranjera Then
                    MatMonE(2, 1) = !cantidad
                    MatMonE(2, 2) = !Total
                    
                    MatMonE(0, 1) = MatMonN(0, 1) + !cantidad
                    MatMonE(0, 2) = MatMonN(0, 2) + !Total
                End If
            ElseIf !sTipo = 3 Then
                If !sMoneda = gMonedaNacional Then
                    MatMonN(3, 1) = !cantidad
                    MatMonN(3, 2) = !Total
                    
                    MatMonN(0, 1) = MatMonN(0, 1) + !cantidad
                    MatMonN(0, 2) = MatMonN(0, 2) + !Total
                ElseIf !sMoneda = gMonedaExtranjera Then
                    MatMonE(3, 1) = !cantidad
                    MatMonE(3, 2) = !Total
                    
                    MatMonE(0, 1) = MatMonN(0, 1) + !cantidad
                    MatMonE(0, 2) = MatMonN(0, 2) + !Total
                End If
            ElseIf !sTipo = 4 Then
                If !sMoneda = gMonedaNacional Then
                    MatMonN(4, 1) = !cantidad
                    MatMonN(4, 2) = !Total
                     
                ElseIf !sMoneda = gMonedaExtranjera Then
                    MatMonE(4, 1) = !cantidad
                    MatMonE(4, 2) = !Total
                    
                End If
            ElseIf !sTipo = 5 Then
                If !sMoneda = gMonedaNacional Then
                    MatMonN(5, 1) = !cantidad
                    MatMonN(5, 2) = !Total
                     
                ElseIf !sMoneda = gMonedaExtranjera Then
                    MatMonE(5, 1) = !cantidad
                    MatMonE(5, 2) = !Total
                End If
            End If
            RaiseEvent Progress(.Bookmark, .RecordCount)
            lrDataRep.MoveNext
        Loop
        RaiseEvent CloseProgress
        End With
        lrDataRep.Close
        Set lrDataRep = Nothing
         
        lsCadImp = lsCadImp & String(99, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "CONCEPTO                        MONEDA   NACIONAL        MONEDA EXTRANJERA           CONSOLIDADO" & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "                                 NRO       MONTO          NRO       MONTO          NRO       MONTO" & lsNegritaOff & Chr(10)
        lsCadImp = lsCadImp & String(99, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "SALDO INICIAL          : " & lsNegritaOff & ImpreFormat(MatMonN(0, 1), 10, 0) & Space(1) & ImpreFormat(MatMonN(0, 2), 10, 2) & _
                    Space(1) & ImpreFormat(MatMonE(0, 1), 10, 0) & Space(1) & ImpreFormat(MatMonE(0, 2), 10, 2) & _
                    Space(1) & ImpreFormat(MatMonN(0, 1) + MatMonE(0, 1), 10, 0) & Space(1) & _
                    ImpreFormat(MatMonN(0, 2) + (MatMonE(0, 2) * pnTipoCambio), 10, 2) & _
                    Chr(10)
        lsCadImp = lsCadImp & String(99, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "INGRESOS NUEVOS        : " & lsNegritaOff & ImpreFormat(MatMonN(1, 1), 10, 0) & Space(1) & ImpreFormat(MatMonN(1, 2), 10, 2) & _
                    Space(1) & ImpreFormat(MatMonE(1, 1), 10, 0) & Space(1) & ImpreFormat(MatMonE(1, 2), 10, 2) & _
                    Space(1) & ImpreFormat(MatMonN(1, 1) + MatMonE(1, 1), 10, 0) & Space(1) & _
                    ImpreFormat(MatMonN(1, 2) + (MatMonE(1, 2) * pnTipoCambio), 10, 2) & _
                    Chr(10)
        lsCadImp = lsCadImp & String(99, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "PAGOS(AMORTIZ.+CANCEL.): " & lsNegritaOff & ImpreFormat(MatMonN(2, 1), 10, 0) & Space(1) & ImpreFormat(MatMonN(2, 2), 10, 2) & _
                    Space(1) & ImpreFormat(MatMonE(2, 1), 10, 0) & Space(1) & ImpreFormat(MatMonE(2, 2), 10, 2) & _
                    Space(1) & ImpreFormat(MatMonN(2, 1) + MatMonE(2, 1), 10, 0) & Space(1) & _
                    ImpreFormat(MatMonN(2, 2) + (MatMonE(2, 2) * pnTipoCambio), 10, 2) & _
                    Chr(10)
        lsCadImp = lsCadImp & String(99, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "SALDO FINAL TOTAL      : " & lsNegritaOff & ImpreFormat(MatMonN(3, 1), 10, 0) & Space(1) & ImpreFormat(MatMonN(3, 2), 10, 2) & _
                    Space(1) & ImpreFormat(MatMonE(3, 1), 10, 0) & Space(1) & ImpreFormat(MatMonE(3, 2), 10, 2) & _
                    Space(1) & ImpreFormat(MatMonN(3, 1) + MatMonE(3, 1), 10, 0) & Space(1) & _
                    ImpreFormat(MatMonN(3, 2) + (MatMonE(3, 2) * pnTipoCambio), 10, 2) & _
                    Chr(10)
        lsCadImp = lsCadImp & String(99, "=") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "SALDO FINAL CON DEMANDA: " & lsNegritaOff & ImpreFormat(MatMonN(4, 1), 10, 0) & Space(1) & ImpreFormat(MatMonN(4, 2), 10, 2) & _
                    Space(1) & ImpreFormat(MatMonE(4, 1), 10, 0) & Space(1) & ImpreFormat(MatMonE(4, 2), 10, 2) & _
                    Space(1) & ImpreFormat(MatMonN(4, 1) + MatMonE(4, 1), 10, 0) & Space(1) & _
                    ImpreFormat(MatMonN(4, 2) + (MatMonE(4, 2) * pnTipoCambio), 10, 2) & _
                    Chr(10)
        lsCadImp = lsCadImp & String(99, "-") & Chr(10)
        lsCadImp = lsCadImp & lsNegritaOn & "SALDO FINAL SIN DEMANDA: " & lsNegritaOff & ImpreFormat(MatMonN(5, 1), 10, 0) & Space(1) & ImpreFormat(MatMonN(5, 2), 10, 2) & _
                    Space(1) & ImpreFormat(MatMonE(5, 1), 10, 0) & Space(1) & ImpreFormat(MatMonE(5, 2), 10, 2) & _
                    Space(1) & ImpreFormat(MatMonN(5, 1) + MatMonE(5, 1), 10, 0) & Space(1) & _
                    ImpreFormat(MatMonN(5, 2) + (MatMonE(5, 2) * pnTipoCambio), 10, 2) & _
                    Chr(10)
         
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
    nRepo138025_ResumenMovimientoCJ = lsCadBuffer
End Function
 
Public Function nRepo150000_ConsultaCreditoenCobranzaJudicial(ByVal psCtaCod As String) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim loImprimeCab As NColPImpre
        
Dim lrDatCredito As ADODB.Recordset
Dim lrDatTotales As ADODB.Recordset
Dim lrDatGastos As ADODB.Recordset
Dim lrDatTotalesGastos As ADODB.Recordset
Dim lrDatAmortizaciones As ADODB.Recordset
Dim lrDatDatosGenerales As ADODB.Recordset

Dim loValCred As DColRecRConsulta
Dim I As Integer
Dim sOperacion As String * 30
Dim sImporte As Currency
Dim sSaldo As Currency

Dim lsNegritaOn As String
Dim lsNegritaOff As String

Dim lsCondiciones As String

lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)
 
    'Carga Datos
    Set loValCred = New DColRecRConsulta
        Set lrDatCredito = loValCred.dObtieneDatosCabeceraRecuperacion(psCtaCod)
        Set lrDatTotales = loValCred.dObtieneDatosTotalesRecuperacion(psCtaCod)
        Set lrDatGastos = loValCred.dObtieneGastosRecuperacion(psCtaCod)
        Set lrDatTotalesGastos = loValCred.dObtieneTotalesGastosRecuperacion(psCtaCod)
        Set lrDatAmortizaciones = loValCred.dObtieneListaAmortizaciones(psCtaCod)
        Set lrDatDatosGenerales = loValCred.dObtieneDatosGenerales(psCtaCod)
    Set loValCred = Nothing
    
    If lrDatCredito Is Nothing Or (lrDatCredito.BOF And lrDatCredito.EOF) Then
        Set lrDatCredito = Nothing
        Set lrDatTotales = Nothing
        Set lrDatGastos = Nothing
        Set lrDatTotalesGastos = Nothing
        Set lrDatAmortizaciones = Nothing
        Set lrDatDatosGenerales = Nothing
        Exit Function
    End If
    
    lsCondiciones = lsNegritaOn & "Crédito Nro. " & psCtaCod & "               Cliente: " & lrDatCredito!cpersNombre & lsNegritaOff
    
    lnLineas = 0
    lnPage = 1

    'CABECERA
    lsCadImp = lsCadImp & nRepoCabecera("CONSULTA DE CREDITOS EN COBRANZA JUDICIAL", "", lnPage, 125, lsCondiciones, 150000)
    lsCadImp = lsCadImp & Chr(10)
    lnIndice = 0:  lnLineas = 8
 
    lsCadImp = lsCadImp & String(125, "-") & Chr(10)
    lsCadImp = lsCadImp & lsNegritaOn & "DATOS DEL CREDITO" & lsNegritaOff & Chr(10)
    lsCadImp = lsCadImp & String(125, "-") & Chr(10)
    lsCadImp = lsCadImp & Space(43) & "Cond. Credito:  " & lrDatCredito!sCondicion & Chr(10)
    lsCadImp = lsCadImp & "Fecha Prestamo : " & lrDatDatosGenerales!FechaPrestamo & Space(15) & " Estad. Credito: " & lrDatCredito!ssEstado & Chr(10)
    lsCadImp = lsCadImp & "Monto Prestamo : " & Format(lrDatDatosGenerales!nMontoCol, "0.00") & Chr(10)
    lsCadImp = lsCadImp & "Linea Credito  : " & lrDatDatosGenerales!cDescripcion & Chr(10)
    lsCadImp = lsCadImp & "Analista       : " & lrDatDatosGenerales!Analista & Chr(10)
    lsCadImp = lsCadImp & String(125, "-") & Chr(10)
    lsCadImp = lsCadImp & lsNegritaOn
    lsCadImp = lsCadImp & "Saldo Capital  : " & ImpreFormat(lrDatTotales!CapitalActual, 10, 2) & Chr(10)
    lsCadImp = lsCadImp & "Saldo Int. Com.: " & ImpreFormat(lrDatTotales!InteresActual, 10, 2) & Chr(10)
    lsCadImp = lsCadImp & "Saldo Int. Mor.: " & ImpreFormat(lrDatTotales!MoraActual, 10, 2) & Chr(10)
    lsCadImp = lsCadImp & "Saldo Gastos   : " & ImpreFormat(lrDatTotales!GastoActual, 10, 2) & Chr(10)
    lsCadImp = lsCadImp & "------------------------------" & Chr(10)
    lsCadImp = lsCadImp & "Deuda          : " & ImpreFormat(lrDatTotales!CapitalActual + lrDatTotales!InteresActual + lrDatTotales!MoraActual + lrDatTotales!GastoActual, 10, 2) & lsNegritaOff & Chr(10)
    lsCadImp = lsCadImp & String(125, "-") & Chr(10)
    lsCadImp = lsCadImp & lsNegritaOn & "GASTOS ASIGNADOS" & lsNegritaOff & Chr(10)
    lsCadImp = lsCadImp & String(125, "-") & Chr(10)
    lsCadImp = lsCadImp & "FECHA      DESCR. GASTO                         IMPORTE" & Chr(10)
    lsCadImp = lsCadImp & String(125, "-") & Chr(10)
    sOperacion = ""
    
    lnLineas = lnLineas + 20
    lnIndice = lnIndice + 20
    
    Set lrDatCredito = Nothing
    Set lrDatDatosGenerales = Nothing
    RaiseEvent ShowProgress
    Do While Not lrDatGastos.EOF
        sOperacion = lrDatGastos!cDescripcion
        lsCadImp = lsCadImp & lrDatGastos!Fecha & Space(1) & sOperacion & Space(1) & ImpreFormat(lrDatGastos!Importe, 10, 2) & Chr(10)
        
        lnLineas = lnLineas + 1
        lnIndice = lnIndice + 1
        
        If lnIndice Mod 300 = 0 Then
            lsCadBuffer = lsCadBuffer & lsCadImp
            lsCadImp = ""
        End If
        
        If lnLineas >= 55 Then
            lsCadImp = lsCadImp & String(122, "-") & "..." & Chr(10)
            lnPage = lnPage + 1
            lsCadImp = lsCadImp & Chr(12)
            lsCadImp = lsCadImp & nRepoCabecera("CONSULTA DE CREDITOS EN COBRANZA JUDICIAL", "", lnPage, 125, lsCondiciones, 150000)
            lsCadImp = lsCadImp & Chr(10)
            lnLineas = 8
            lnIndice = lnIndice + 8
            lsCadImp = lsCadImp & lsNegritaOn & "GASTOS ASIGNADOS" & lsNegritaOff & Chr(10)
            lsCadImp = lsCadImp & String(125, "-") & Chr(10)
            lsCadImp = lsCadImp & "FECHA      DESCR. GASTO                         IMPORTE" & Chr(10)
            lsCadImp = lsCadImp & String(125, "-") & Chr(10)
            lnLineas = lnLineas + 4
            lnIndice = lnIndice + 4
        End If
        RaiseEvent Progress(lrDatGastos.Bookmark, lrDatGastos.RecordCount)
        lrDatGastos.MoveNext
    Loop
    Set lrDatGastos = Nothing
    RaiseEvent CloseProgress
    
    lsCadImp = lsCadImp & String(125, "-") & Chr(10)
    lsCadImp = lsCadImp & Space(11) & "GASTOS ASIGNADOS: " & Space(13) & ImpreFormat(lrDatTotalesGastos!Importe, 10, 2) & Chr(10) & Chr(10)
    
    lnLineas = lnLineas + 3
    lnIndice = lnIndice + 3
    
    Set lrDatTotalesGastos = Nothing
    
    lsCadImp = lsCadImp & lsNegritaOn & "MOVIMIENTOS REALIZADOS" & lsNegritaOff & Chr(10)
    lsCadImp = lsCadImp & String(125, "-") & Chr(10)
    lsCadImp = lsCadImp & "FECHA      OPERACION                              MONTO       CAPITAL       INTERES          MORA        GASTOS      SALDO K. " & Chr(10)
    lsCadImp = lsCadImp & String(125, "-") & Chr(10)
    
    
    sSaldo = CDbl(lrDatTotales!CapitalPagado + lrDatTotales!CapitalActual)
             
    sOperacion = ""
    RaiseEvent ShowProgress
    Do While Not lrDatAmortizaciones.EOF
        sImporte = lrDatAmortizaciones!Capital + lrDatAmortizaciones!Interes + lrDatAmortizaciones!Mora + lrDatAmortizaciones!Gastos
        sSaldo = sSaldo - lrDatAmortizaciones!Capital
        sOperacion = lrDatAmortizaciones!Operacion
        lsCadImp = lsCadImp & lrDatAmortizaciones!Fecha & Space(1) & sOperacion & Space(1) & ImpreFormat(sImporte, 10, 2) & Space(1) & ImpreFormat(lrDatAmortizaciones!Capital, 10, 2) & Space(1) & ImpreFormat(lrDatAmortizaciones!Interes, 10, 2) & Space(1) & ImpreFormat(lrDatAmortizaciones!Mora, 10, 2) & Space(1) & ImpreFormat(lrDatAmortizaciones!Gastos, 10, 2) & Space(1) & ImpreFormat(sSaldo, 10, 2) & Chr(10)
         
        lnLineas = lnLineas + 1
        lnIndice = lnIndice + 1
        
        If lnIndice Mod 300 = 0 Then
            lsCadBuffer = lsCadBuffer & lsCadImp
            lsCadImp = ""
        End If
        
        If lnLineas >= 55 Then
            lsCadImp = lsCadImp & String(122, "-") & "..." & Chr(10)
            lnPage = lnPage + 1
            lsCadImp = lsCadImp & Chr(12)
            lsCadImp = lsCadImp & nRepoCabecera("CONSULTA DE CREDITOS EN COBRANZA JUDICIAL", "", lnPage, 125, lsCondiciones, 150000)
            lsCadImp = lsCadImp & Chr(10)
            lnLineas = 8
            lnIndice = lnIndice + 8
            lsCadImp = lsCadImp & lsNegritaOn & "MOVIMIENTOS REALIZADOS" & lsNegritaOff & Chr(10)
            lsCadImp = lsCadImp & String(125, "-") & Chr(10)
            lsCadImp = lsCadImp & "FECHA      OPERACION                              MONTO       CAPITAL       INTERES          MORA        GASTOS      SALDO K. " & Chr(10)
            lsCadImp = lsCadImp & String(125, "-") & Chr(10)
            
            lnLineas = lnLineas + 4
            lnIndice = lnIndice + 4
        End If
        RaiseEvent Progress(lrDatAmortizaciones.Bookmark, lrDatAmortizaciones.RecordCount)
        lrDatAmortizaciones.MoveNext
    Loop
    RaiseEvent CloseProgress
    Set lrDatAmortizaciones = Nothing
    
    lsCadImp = lsCadImp & String(125, "-") & Chr(10)
    lsCadImp = lsCadImp & Space(11) & "TOTAL PAGADO:" & Space(18) & ImpreFormat(lrDatTotales!CapitalPagado + lrDatTotales!InteresPagado + lrDatTotales!MoraPagado + lrDatTotales!GastoPagado, 10, 2)
    lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
     
    Set lrDatTotales = Nothing
      
    nRepo150000_ConsultaCreditoenCobranzaJudicial = lsCadBuffer
     
End Function
