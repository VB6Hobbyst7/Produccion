VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMColRecRConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim csNomCMAC As String
Dim csNomAgencia As String
Dim csCodUser As String
Dim csFechaSis As String

Dim oFun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion

Public Sub inicio(ByVal psNomCmac As String, ByVal psNomAgencia As String, _
            ByVal psCodUser As String, ByVal psFechaSis As String)
    csNomCMAC = psNomCmac
    csNomAgencia = psNomAgencia
    csCodUser = psCodUser
    csFechaSis = psFechaSis
End Sub

        
Public Function nRepoCabecera(ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        ByVal pnCodReporte As Long, Optional ByVal psImpresoras As Impresoras = gEPSON) As String
        
Dim lsCadImp As String
Dim loImprimeCab As New COMNColoCPig.NCOMColPImpre

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresoras

    
    Set loImprimeCab = New COMNColoCPig.NCOMColPImpre
        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
        'lsCadImp = loImprimeCab.nImprimeCabeceraReportes(csNomCMAC, csNomAgencia, csCodUser, gdFecSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    Set loImprimeCab = Nothing
    If pnCodReporte <> 150000 And pnCodReporte <> 138025 Then
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
    End If
    Select Case pnCodReporte
        
        Case 138001 ' Listado de Pagos x Abogado
            lsCadImp = lsCadImp & "ITEM  FECHA      CLIENTE                               MONTO           A           A              A              A           SALDO       COMISION " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "      PAGO                                             PAGADO       CAPITAL    INT. COMP.     INT. MORA        GASTOS       CAPITAL        S.K.   " & oFunI.gPrnSaltoLinea
             
        Case 138002 ' Listado de Pagos x Credito
            lsCadImp = lsCadImp & "ITEM  FECHA      FECHA      CLIENTE                               MONTO           A           A              A              A           SALDO       COMISION " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "      PAGO       JUD.                                             PAGADO       CAPITAL    INT. COMP.     INT. MORA        GASTOS       CAPITAL        S.K.   " & oFunI.gPrnSaltoLinea
          
        Case 138003 ' Listado de Pagos x Fecha
            lsCadImp = lsCadImp & "ITEM  CUENTA             FECHA       CLIENTE                           MONTO        A         A           A           A        SALDO    COMISION USER AG" & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "                         PAGO                                          PAGADO    CAPITAL  INT. COMP.  INT. MORA     GASTOS    CAPITAL      S.K.   " & oFunI.gPrnSaltoLinea
           
        Case 138004 ' Listado de Pagos x Analista
            lsCadImp = lsCadImp & "ITEM  FECHA      CLIENTE                               MONTO           A           A              A              A           SALDO       COMISION " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "      PAGO                                             PAGADO       CAPITAL    INT. COMP.     INT. MORA        GASTOS       CAPITAL        S.K.   " & oFunI.gPrnSaltoLinea
            
        Case 138011 ' Listado de Creditos Nuevos
            lsCadImp = lsCadImp & "ITEM  CUENTA             FEC.VIGEN. FEC.ING. CLIENTE                        ANA.          MONTO       CAPITAL     INT. COMP    INT. MORAT     GASTOS        TOTAL " & oFunI.gPrnSaltoLinea
 
        Case 138012 ' Listado de Creditos En Judicial
            lsCadImp = lsCadImp & "ITEM CUENTA             FEC.ING.   CLIENTE                        ANA.       PRESTAMO     SALDO CAP      INTERES         MORA       GASTOS      TOTAL    EXP" & oFunI.gPrnSaltoLinea
            
        Case 138014 ' Listado de Creditos x Analista
            lsCadImp = lsCadImp & "ITEM  CUENTA             F.ING.REC. CLIENTE                             PRESTAMO     SALDO CAP       INTERES          MORA        GASTOS       TOTAL   ABOGADO" & oFunI.gPrnSaltoLinea
            
        Case 138015 ' REPORTE DE CARTERA DE CREDI JUDI Y CAST PEAC 20071025
            lsCadImp = lsCadImp & "ITEM  CUENTA           CLIENTE                            DIRECCION                          DIAS ATRASO  CAPITAL      INTERES    MORA       GASTOS       TOTAL" & oFunI.gPrnSaltoLinea
            
        Case 138021 'Listado de Consolidado por Abogado
            lsCadImp = lsCadImp & "DESCRIPCION                    PRESTAMO      CAPITAL       INTERES          MORA        GASTOS         TOTAL " & oFunI.gPrnSaltoLinea
        
        Case 138023 ' Listado Segun Saldo de capital
            lsCadImp = lsCadImp & " ITEM CUENTA              CLIENTE                        FECHA      FECHA ING.  TIPO DE      MONTO       SALDO    ESTADO" & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "                                                         PRESTAMO   COB. JUD.   MONEDA     PRESTAMO   CAPITAL S/." & oFunI.gPrnSaltoLinea
            
        Case 138024 'Listado de Expedientes por Abogado
            'lsCadImp = lsCadImp & "ITEM  CUENTA             LINEA CRED. CLIENTE                             PRESTAMO         SALDO       INTERES          MORA        GASTOS       TOTAL   ANA. SITUACION CASO" & ofunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "ITEM  CUENTA             F.ING.REC. CLIENTE                             PRESTAMO     SALDO CAP       INTERES          MORA        GASTOS       TOTAL   ANALISTA" & oFunI.gPrnSaltoLinea
        Case 138036 'DAOR 20070124, Listado de Creditos Castigados
            lsCadImp = lsCadImp & "COD. CUENTA        COD. CLIENTE  NOMBRE DEL CLIENTE             FECHA      DESEMBOLSO     CAP. PAGADO   SALDO         ANALISTA D. ATRASO AGENCIA" & oFunI.gPrnSaltoLinea
         
        Case 138038 'PEAC 20071010 reporte de gastos de cred jud y cast
         lsCadImp = lsCadImp & "                                    MONTO      FECHA            MONTO     FECHA               " & oFunI.gPrnSaltoLinea
         lsCadImp = lsCadImp & "    ITEM   CONCEPTO                 INGRESADO  INGRESADA        PAGADO    PAGADA         SALDO" & oFunI.gPrnSaltoLinea
         
        Case 138039 'PEAC 20071011 reporte de creditos con saldos condonados
         lsCadImp = lsCadImp & "       CREDITO           CREDITO                                                                                                                    FECHA             " & oFunI.gPrnSaltoLinea
         lsCadImp = lsCadImp & "ITEM   ANTIGUO           ACTUAL              NOMBRE                          ANALISTA  CAPITAL      INTERES         MORA       GASTOS        TOTAL  CONDONACI.  ESTADO" & oFunI.gPrnSaltoLinea
         
        Case 150000
             
    End Select
    If pnCodReporte <> 150000 And pnCodReporte <> 138025 Then
        lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
    End If
nRepoCabecera = lsCadImp
End Function

 
Public Function nRepo138001_ListadoPagosxAbogado(ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal psConAbogado As Byte, ByVal psCodAbogado As String, _
                                                 Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As New COMNColoCPig.NCOMColPImpre
        
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


Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

'ARCV 30-06-2007
Dim oCon As COMConecta.DCOMConecta

lsSQL = "SELECT * INTO #MOVRECUP FROM Mov " & _
        " Where nMovFlag = 0 " & _
        " AND CONVERT(datetime, SUBSTRING(Mov.cMovNro, 1, 8), 103) BETWEEN '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "' " & _
        " and cOpeCod like '130%'"
Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
oCon.Ejecutar lsSQL
'-----
    
    sOpeEst = " (MOC.cOpeCod in(" & gColRecOpePasoARecup & ", " & gColRecOpePagJudSD & ", " & gColRecOpePagJudSDEfe & ", " & gColRecOpePagJudSDChq & ", " & gColRecOpePagJudCD & ", " & gColRecOpePagJudCDEfe & ", " & gColRecOpePagJudCDChq & ", " & gColRecOpePagCast & ", " & gColRecOpePagCastEfe & ", " & gColRecOpePagCastChq & ", " & gColRecOpePagJudSDEnOtCj & ", " & gColRecOpePagJudSDEnOtCjEfe & ", " & gColRecOpePagJudCDEnOtCj & ", " & gColRecOpePagJudCDEnOtCjEfe & ", " & gColRecOpePagJudCastEnOtCj & ", " & gColRecOpePagJudCastEnOtCjEfe & ")) AND "
    
    'ARCV 30-06-2007
    'sCriterio = " AND (CONVERT(datetime, SUBSTRING(Mov.cMovNro, 1, 8), 103) Between '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "') "
    sCriterio = " AND (CONVERT(datetime, SUBSTRING(M.cMovNro, 1, 8), 103) Between '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "') "
    '----
    
    lsCondiciones = "Del " & pdFecIni & " al " & pdFecFin
 
    'Si se busca ABOGADO
    If psConAbogado = 1 Then
         
        sCriterio = sCriterio & " AND ((SELECT PRP.CPERSCOD FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON " & _
        " PRP.CPERSCOD = PERSO.CPERSCOD WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ")='" & Trim(psCodAbogado) & "') "
 
        lsCondiciones = lsCondiciones & " - Abogado: " & psCodAbogado & " : "
    End If

'ARCV 30-06-2007
'lsSQL = "SELECT MC.nMonto AS Pagado, " & _
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


'ANDE20160206 agregó AND MOC.nNroCalen = MCD.nNroCalen'
lsSQL = "SELECT MC.nMonto AS Pagado, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodCapital & "'), 0) AS Capital, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO AND MOC.nNroCalen = MCD.nNroCalen INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & "  mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodInteresCompensatorio & "'), 0) AS InteresCom, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & "  mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodInteresMoratorio & "'), 0) AS InteresMor, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & "  mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND (MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoCobranza & "' OR MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoComisionAbog & "' OR MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoAdmin & "')), 0) AS Gastos, " & _
        " isnull(MC.nSaldoCap,0) AS SaldoCapital, ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & "  mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoComisionAbog & "'), 0) AS ComisionAbog, Persona.cPersNombre, " & _
        " (SELECT PRP.CPERSCOD FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD " & _
        " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ") AS CODABOGADO, (SELECT PERSO.cPersNombre " & _
        " FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ") AS ABOGADO, " & _
        " CONVERT(varchar(10), CONVERT(datetime, SUBSTRING(M.cMovNro, 1, 8), 103), 103) AS Fecha, SUBSTRING(Colocaciones.cLineaCred, 5, 1) AS Moneda, Constante.cConsDescripcion AS sMoneda, " & _
        " nTipoRecuperacion=case when Producto.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "') THEN 1 ELSE 2 END, " & _
        " nDesTipoRecuperacion=case when Producto.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "') THEN 'JUDICIAL' ELSE 'CASTIGADO' END " & _
        " FROM ColocRecup INNER JOIN Colocaciones ON ColocRecup.cCtaCod = Colocaciones.cCtaCod INNER JOIN MovCol MC INNER JOIN ProductoPersona PP ON MC.cCtaCod = PP.cCtaCod INNER JOIN " & _
        " #MOVRECUP M ON MC.nMovNro = M.nMovNro ON ColocRecup.cCtaCod = PP.cCtaCod INNER JOIN Persona ON PP.cPersCod = Persona.cPersCod INNER JOIN Producto ON " & _
        " Colocaciones.cCtaCod = Producto.cCtaCod AND PP.cCtaCod = Producto.cCtaCod INNER JOIN " & _
        " Constante ON SUBSTRING(Colocaciones.cLineaCred, 5, 1) = Constante.nConsValor " & _
        " INNER Join  OpeTpo Ope on ope.cOpeCod = MC.cOpeCod " & _
        " WHERE (PP.nPrdPersRelac = 20)  " & _
        " AND (Producto.nPrdEstado IN ('" & gColocEstRecVigJud & "', '" & gColocEstRecVigCast & "', '" & gColocEstRecCanJud & "', '" & gColocEstRecCanCast & "')) AND (Constante.nConsCod = '1011') " & _
        sCriterio
'----------------

lsSQL = lsSQL & " and Ope.cOpeCod like '130%' and Ope.nOpeNiv<>1"
lsSQL = lsSQL & " Order By (SELECT PERSO.cPersNombre FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON " & _
        " PRP.cPersCod = PERSO.cPersCod WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & "), " & _
        " SUBSTRING(Colocaciones.cLineaCred, 5, 1), case when Producto.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "') THEN 1 ELSE 2 END, SUBSTRING(M.cMovNro, 1, 8) "
 
'ARCV 30-06-2007
    'Set loDataRep = New COMDColocPig.DCOMColPFunciones
    '    Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    'Set loDataRep = Nothing
Set lrDataRep = oCon.CargaRecordSet(lsSQL)
oCon.CierraConexion
Set oCon = Nothing
'-------

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
        lsCadImp = lsCadImp & lsNegritaOn & "ABOGADO: " & lrDataRep!Abogado & " : " & lrDataRep!CODABOGADO & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'Tipo de Recuperacion
        lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
        
        lnLineas = lnLineas + 6
        lnIndice = lnIndice + 6
          
        With lrDataRep
            'RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                
                'Totalizar por Abogado
                If sTempoAbogado <> !CODABOGADO Then
                
                    lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal41, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                
                    lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(23) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                    lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ABOGADO " & Space(30) & oFun.ImpreFormat(sSubTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal3, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal5, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal6, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal7, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                     
                    'Abogado
                    lsCadImp = lsCadImp & lsNegritaOn & "ABOGADO: " & lrDataRep!Abogado & " : " & lrDataRep!CODABOGADO & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    'Tipo de Recuperacion
                    lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                    
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
                
                        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal41, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(23) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        
                        'Moneda
                        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        'Tipo de Recuperacion
                        lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                        
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
                    
                            lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal41, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                            'Tipo de Recuperacion
                            lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                            
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
                 
                sCliente = !cPersNombre
                 
                lsCadImp = lsCadImp & Space(1) & oFun.ImpreFormat(nItem, 4, 0) & Space(1) & !Fecha & Space(1) & _
                           sCliente & Space(1) & oFun.ImpreFormat(!Pagado, 10, 2) & Space(1) & oFun.ImpreFormat(!Capital, 10, 2) & _
                           Space(1) & oFun.ImpreFormat(!InteresCom, 10, 2) & Space(1) & oFun.ImpreFormat(!InteresMor, 10, 2) & _
                           Space(1) & oFun.ImpreFormat(!Gastos, 10, 2) & Space(1) & oFun.ImpreFormat(!SaldoCapital, 10, 2) & _
                           Space(1) & oFun.ImpreFormat(!ComisionAbog, 10, 2) & oFunI.gPrnSaltoLinea
                 
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE PAGOS DE CREDITOS POR ABOGADO", lsCondiciones, lnPage, 145, "", 138001)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                'RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            'RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal41, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    
        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(23) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ABOGADO " & Space(30) & oFun.ImpreFormat(sSubTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal3, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal5, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal6, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal7, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
    nRepo138001_ListadoPagosxAbogado = lsCadBuffer
End Function


Public Function nRepo138002_ListadoPagosdeCreditos(ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal psEstados As String, ByVal pnMoneda As Byte, _
                                                   Optional ByVal psJudicial As String = "", Optional ByVal psTipoC As String = "", Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As New COMNColoCPig.NCOMColPImpre
        
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
 

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
 
'ARCV 30-06-2007
Dim oCon As COMConecta.DCOMConecta

lsSQL = "SELECT * INTO #MOVRECUP FROM Mov " & _
        " Where nMovFlag = 0 " & _
        " AND CONVERT(datetime, SUBSTRING(Mov.cMovNro, 1, 8), 103) BETWEEN '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "' " & _
        " and cOpeCod like '130%'"
Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
oCon.Ejecutar lsSQL
'-----
     

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
         
    '**** BRGO BASILEA II
    'ANDE20160206 agregó AND MOC.nNroCalen = MCD.nNroCalen'
    lsSQL = " SELECT DISTINCT MC.nMonto AS nPagado, " & _
             " ISNULL ((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO = MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO = MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro = MC.nMovNro AND MOC.cCtaCod = MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodCapital & "'), 0) AS nCapital, " & _
             " ISNULL ((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO = MOC.NMOVNRO AND MOC.nNroCalen = MCD.nNroCalen INNER JOIN MOV MV ON MV.NMOVNRO = MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro = MC.nMovNro AND MOC.cCtaCod = MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodInteresCompensatorio & "'), 0) AS nInteresCom, " & _
             " ISNULL ((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO = MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO = MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro = MC.nMovNro AND MOC.cCtaCod = MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodInteresMoratorio & "'), 0) AS nInteresMor, " & _
             " ISNULL ((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO = MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO = MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro = MC.nMovNro AND MOC.cCtaCod = MC.cCtaCod AND (MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoCobranza & "' OR MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoComisionAbog & "' OR MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoAdmin & "')), 0) AS nGastos, ISNULL(MC.nSaldoCap, 0) AS nSaldoCapital, " & _
             " ISNULL ((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO = MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO = MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro = MC.nMovNro AND MOC.cCtaCod = MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodGastoComisionAbog & "'), 0) AS nComisionAbog, " & _
             " ISNULL(MC.nSaldoCap, 0) AS nSaldoCapital, Persona.cPersNombre, CONVERT(varchar(10), CONVERT(datetime, SUBSTRING(M.cMovNro, 1, 8), 103), 103) AS dFechaPago, SUBSTRING(C.cLineaCred, 5, 1) AS nMoneda, " & _
             " K.cConsDescripcion AS cMoneda, C.cLineaCred AS cLineaCredito, CLC.cDescripcion AS cDesLineaCredito, C.cCtaCod, C.cAgeCodAct AS cCodAgencia, A.cAgeDescripcion, CR.dIngRecup AS dFechaJudicial " & _
             " FROM Producto P INNER JOIN ColocRecup CR INNER JOIN Colocaciones C ON CR.cCtaCod = C.cCtaCod ON P.cCtaCod = C.cCtaCod INNER JOIN Persona INNER JOIN MovCol MC INNER JOIN " & _
             " ProductoPersona PP ON MC.cCtaCod = PP.cCtaCod ON Persona.cPersCod = PP.cPersCod ON P.cCtaCod = PP.cCtaCod INNER JOIN Constante K ON SUBSTRING(C.cLineaCred, 5, 1) = K.nConsValor INNER JOIN " & _
             " ColocLineaCredito CLC ON C.cLineaCred = CLC.cLineaCred INNER JOIN Agencias A ON C.cAgeCodAct = A.cAgeCod " & _
             " INNER JOIN #MOVRECUP M ON MC.nMovNro=M.nMovNro " & _
             " INNER Join  OpeTpo Ope on ope.cOpeCod = MC.cOpeCod " & _
             " WHERE (PP.nPrdPersRelac=" & gColRelPersTitular & ") AND  (K.nConsCod='" & gMoneda & "') " & _
               sCriterio & _
             " and Ope.cOpeCod like '130%' and Ope.nOpeNiv<>1 "
    '********************
If Trim(psTipoC) <> "" Then
   lsSQL = lsSQL & "and nTipCj in " & psTipoC
End If
   lsSQL = lsSQL & " ORDER BY SUBSTRING(C.cLineaCred,5,1), A.cAgeDescripcion, C.cLineaCred "

Set lrDataRep = oCon.CargaRecordSet(lsSQL)
oCon.CierraConexion
Set oCon = Nothing
'-------

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
        sTempoMoneda = lrDataRep!nmoneda
        sTempoLineaCredito = "" & lrDataRep!cLineaCredito
        
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'Agencia
        lsCadImp = lsCadImp & lsNegritaOn & "**  AGENCIA: " & lrDataRep!cAgeDescripcion & " : " & lrDataRep!cCodAgencia & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'Linea de Credito
        lsCadImp = lsCadImp & lsNegritaOn & "*** LINEA DE CREDITO: " & lrDataRep!cDesLineaCredito & lsNegritaOff & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(156, "-") & oFunI.gPrnSaltoLinea
        
        lnLineas = lnLineas + 6
        lnIndice = lnIndice + 6
          
        With lrDataRep
            'RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                
                'Totalizar por Moneda
                
                If sTempoMoneda <> !nmoneda Then
                
                    lsCadImp = lsCadImp & String(156, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X LINEA DE CRED.  " & Space(23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal41, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    
                    lsCadImp = lsCadImp & String(156, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X AGENCIA         " & Space(23) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
                    lsCadImp = lsCadImp & String(156, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X MONEDA   " & Space(30) & oFun.ImpreFormat(sSubTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal3, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal5, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal6, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal7, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                     
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    'Agencia
                    lsCadImp = lsCadImp & lsNegritaOn & "**  AGENCIA: " & lrDataRep!cAgeDescripcion & " : " & lrDataRep!cCodAgencia & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    'Linea de Credito
                    lsCadImp = lsCadImp & lsNegritaOn & "*** LINEA DE CREDITO: " & lrDataRep!cDesLineaCredito & lsNegritaOff & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & String(156, "-") & oFunI.gPrnSaltoLinea
                     
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
                    
                    sTempoMoneda = lrDataRep!nmoneda
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
                
                        lsCadImp = lsCadImp & String(156, "-") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X LINEA DE CRED.  " & Space(23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                        Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 10, 2) & _
                        Space(1) & oFun.ImpreFormat(sTotal41, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 10, 2) & _
                        Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
                        lsCadImp = lsCadImp & String(156, "-") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X AGENCIA         " & Space(23) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                        Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                        Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                        Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                            
                        'Agencia
                        lsCadImp = lsCadImp & lsNegritaOn & "**  AGENCIA: " & lrDataRep!cAgeDescripcion & " : " & lrDataRep!cCodAgencia & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        'Linea de Credito
                        lsCadImp = lsCadImp & lsNegritaOn & "*** LINEA DE CREDITO: " & lrDataRep!cDesLineaCredito & lsNegritaOff & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & String(156, "-") & oFunI.gPrnSaltoLinea
                        
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
                    
                            lsCadImp = lsCadImp & String(156, "-") & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X LINEA DE CRED.  " & Space(23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                            Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 10, 2) & _
                            Space(1) & oFun.ImpreFormat(sTotal41, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 10, 2) & _
                            Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                     
                            'Linea de Credito
                            lsCadImp = lsCadImp & lsNegritaOn & "*** LINEA DE CREDITO: " & lrDataRep!cDesLineaCredito & lsNegritaOff & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & String(156, "-") & oFunI.gPrnSaltoLinea
                            
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
                 
                sCliente = !cPersNombre
                 
                lsCadImp = lsCadImp & Space(1) & oFun.ImpreFormat(nItem, 4, 0) & Space(1) & !dFechaPago & Space(1) & _
                           Format(!dFechaJudicial, "dd/MM/YYYY") & _
                           Space(1) & sCliente & Space(1) & oFun.ImpreFormat(!nPagado, 10, 2) & Space(1) & oFun.ImpreFormat(!nCapital, 10, 2) & _
                           Space(1) & oFun.ImpreFormat(!nInteresCom, 10, 2) & Space(1) & oFun.ImpreFormat(!nInteresMor, 10, 2) & _
                           Space(1) & oFun.ImpreFormat(!nGastos, 10, 2) & Space(1) & oFun.ImpreFormat(!nSaldoCapital, 10, 2) & _
                           Space(1) & oFun.ImpreFormat(!nComisionAbog, 10, 2) & oFunI.gPrnSaltoLinea
                 
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE PAGOS DE CREDITOS", lsCondiciones, lnPage, 156, "", 138002)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                'RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            'RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(156, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X LINEA DE CRED.  " & Space(23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal41, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    
        lsCadImp = lsCadImp & String(156, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X AGENCIA         " & Space(23) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & String(156, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X MONEDA   " & Space(30) & oFun.ImpreFormat(sSubTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal3, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal5, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal6, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal7, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
      
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
    nRepo138002_ListadoPagosdeCreditos = lsCadBuffer
End Function


Public Function nRepo138003_ListadoPagosxFechas(ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal psEstados As String, ByVal pnMoneda As Byte, Optional ByVal psTipoC As String = "", Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As New COMNColoCPig.NCOMColPImpre
        
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
Dim cMoneda As String '*** PEAC 20080725
Dim cTipo As String
Dim sEstados As String

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)
     
     
'    Set oDCred = New COMDColocRec.DCOMColRecCredito
'        Set R = oDCred.dObtienegRepCredJudToCastigar(psCodAge, gdFecSis, pnAtrIni, pnAtrFin, psListaAgencias)
'        Set R = oDCred.dObtienegRepCredJudToCastigar(psCodAge, gdFecSis, pnAtrIni, pnAtrFin, psListaAgencias)
'    Set oDCred = Nothing
     
     
    lsCondiciones = "Del " & pdFecIni & " al " & pdFecFin
    
    If psEstados = "'" & gColocEstRecVigJud & "', '2205', '" & gColocEstRecCanJud & "'" Then
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
     
    If pnMoneda = 3 Then
        cMoneda = "%"
    Else
        cMoneda = pnMoneda
    End If

    sOpe = " (MOC.cOpeCod in(" & gColRecOpePasoARecup & ", " & gColRecOpePagJudSD & ", " & gColRecOpePagJudSDEfe & ", " & gColRecOpePagJudSDChq & ", " & gColRecOpePagJudCD & ", " & gColRecOpePagJudCDEfe & ", " & gColRecOpePagJudCDChq & ", " & gColRecOpePagCast & ", " & gColRecOpePagCastEfe & ", " & gColRecOpePagCastChq & ", " & gColRecOpePagJudSDEnOtCj & ", " & gColRecOpePagJudSDEnOtCjEfe & ", " & gColRecOpePagJudCDEnOtCj & ", " & gColRecOpePagJudCDEnOtCjEfe & ", " & gColRecOpePagJudCastEnOtCj & ", " & gColRecOpePagJudCastEnOtCjEfe & ")) AND "
     
    sCriterio = sCriterio & " AND (CONVERT(datetime, SUBSTRING(M.cMovNro, 1, 8), 103) Between '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "') "
           
    sEstados = Replace(psEstados, "'", "")
  

    If Trim(psTipoC) <> "" Then
        cTipo = Replace(Replace(Replace(Trim(psTipoC), "(", ""), ")", ""), "'", "")
        lsSQL = "exec stp_sel_ListadoPagosPorFechas '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & sEstados & "','" & cMoneda & "','" & cTipo & "'"
    Else
        lsSQL = "exec stp_sel_ListadoPagosPorFechas '" & pdFecIni & "','" & pdFecIni & "','" & sEstados & "','" & cMoneda & "'"
    End If
         
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
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
        lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE PAGOS POR FECHAS", lsCondiciones, lnPage, 160, "", 138003)

        lnIndice = 0:  lnLineas = 7
           
        sTempoMoneda = lrDataRep!nmoneda
        
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(160, "-") & oFunI.gPrnSaltoLinea
        lnLineas = lnLineas + 2
        lnIndice = lnIndice + 2
        
        With lrDataRep
            'RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                 
                If sTempoMoneda <> !nmoneda Then
                
                    lsCadImp = lsCadImp & String(160, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X MONEDA          " & Space(28) & oFun.ImpreFormat(sTotal1, 8, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal2, 6, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 5, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal4, 7, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 7, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal6, 8, 2) & Space(1) & oFun.ImpreFormat(sTotal7, 7, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
 
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!cMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & String(160, "-") & oFunI.gPrnSaltoLinea
                    
                    nItem = 1
                    
                    sTempoMoneda = lrDataRep!nmoneda
                    
                    lnLineas = lnLineas + 5
                    lnIndice = lnIndice + 5
                    
                    sTotal1 = !MONTO_PGDO
                    sTotal2 = !A_CAPITAL
                    sTotal3 = !A_INT_COMP
                    sTotal4 = !A_INT_MORAT
                    sTotal5 = !A_GASTOS
                    sTotal6 = !SALDO_CAPITAL
                    sTotal7 = !COMISION_SK
                Else
                    sTotal1 = sTotal1 + !MONTO_PGDO
                    sTotal2 = sTotal2 + !A_CAPITAL
                    sTotal3 = sTotal3 + !A_INT_COMP
                    sTotal4 = sTotal4 + !A_INT_MORAT
                    sTotal5 = sTotal5 + !A_GASTOS
                    sTotal6 = sTotal6 + !SALDO_CAPITAL
                    sTotal7 = sTotal7 + !COMISION_SK
                End If
                
                sCliente = !CLIENTE
                 
                lsCadImp = lsCadImp & Space(1) & oFun.ImpreFormat(nItem, 4, 0) & Space(1) & !CUENTA & Space(1) & _
                           !FECHA_PAGO & Space(1) & _
                           Space(1) & sCliente & Space(1) & oFun.ImpreFormat(!MONTO_PGDO, 8, 2) & Space(1) & oFun.ImpreFormat(!A_CAPITAL, 6, 2) & _
                           Space(1) & oFun.ImpreFormat(!A_INT_COMP, 5, 2) & Space(1) & oFun.ImpreFormat(!A_INT_MORAT, 7, 2) & _
                           Space(1) & oFun.ImpreFormat(!A_GASTOS, 7, 2) & Space(1) & oFun.ImpreFormat(!SALDO_CAPITAL, 8, 2) & _
                           Space(1) & oFun.ImpreFormat(!COMISION_SK, 7, 2) & Space(1) & !USUARIO & Space(1) & !cCodAgencia & oFunI.gPrnSaltoLinea
                 
                 
                '!analista
                 
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE PAGOS POR FECHAS", lsCondiciones, lnPage, 173, "", 138003)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                'RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            'RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(160, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "      *** TOTALES X MONEDA          " & Space(28) & oFun.ImpreFormat(sTotal1, 8, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal2, 6, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 5, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal4, 7, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 7, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal6, 8, 2) & Space(1) & oFun.ImpreFormat(sTotal7, 7, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
  
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
    nRepo138003_ListadoPagosxFechas = lsCadBuffer
End Function
 
 
Public Function nRepo138004_ListadoPagosxAnalista(ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal psConAnalista As Byte, ByVal psCodAnalista As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As New COMNColoCPig.NCOMColPImpre
        
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
 

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

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
'ANDE20160206 agregó AND MOC.nNroCalen = MCD.nNroCalen'
lsSQL = "SELECT  MC.nMonto AS Pagado, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodCapital & "'), 0) AS Capital, " & _
        " ISNULL((SELECT SUM(MCD.nMonto) FROM MOVCOLDET MCD INNER JOIN MOVCOL MOC ON MCD.NMOVNRO=MOC.NMOVNRO AND MOC.nNroCalen = MCD.nNroCalen INNER JOIN MOV MV ON MV.NMOVNRO=MOC.NMOVNRO WHERE " & sOpeEst & " mV.nMovFlag=" & gMovFlagVigente & " AND MOC.nMovNro=MC.nMovNro AND MOC.cCtaCod=MC.cCtaCod AND MCD.nPrdConceptoCod='" & gColRecConceptoCodInteresCompensatorio & "'), 0) AS InteresCom, " & _
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

lsSQL = lsSQL & " and  SUBSTRING(Mov.cMovNro, 1, 8) BETWEEN  CONVERT(VARCHAR(8), CONVERT(datetime, '" & Format(pdFecIni, "MM/DD/YYYY") & "') ,112) AND CONVERT(VARCHAR(8), CONVERT(datetime, '" & Format(pdFecFin, "MM/DD/YYYY") & "') ,112) "

lsSQL = lsSQL & " Order By (SELECT PERSO.cPersNombre FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON " & _
        " PRP.cPersCod = PERSO.cPersCod WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & "), " & _
        " SUBSTRING(Colocaciones.cLineaCred, 5, 1), case when Producto.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "') THEN 1 ELSE 2 END, SUBSTRING(Mov.cMovNro, 1, 8) "
 
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
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
        lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA: " & lrDataRep!Analista & " - " & lrDataRep!SigAnalista & " : " & lrDataRep!CodAnalista & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'Tipo de Recuperacion
        lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
        
        lnLineas = lnLineas + 6
        lnIndice = lnIndice + 6
          
        With lrDataRep
            'RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                
                'Totalizar por Analista
                
                If sTempoAnalista <> !CodAnalista Then
                
                    lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal41, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                
                    lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(23) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                    lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X  ANALISTA" & Space(29) & oFun.ImpreFormat(sSubTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal3, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal5, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal6, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal7, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                     
                    'Analista
                    lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA: " & lrDataRep!Analista & " - " & lrDataRep!SigAnalista & " : " & lrDataRep!CodAnalista & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    'Tipo de Recuperacion
                    lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                    
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
                
                        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal41, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(23) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        
                        'Moneda
                        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        'Tipo de Recuperacion
                        lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                        
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
                    
                            lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal41, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                            'Tipo de Recuperacion
                            lsCadImp = lsCadImp & lsNegritaOn & lrDataRep!nDesTipoRecuperacion & lsNegritaOff & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                            
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
                
                sCliente = !cPersNombre
                 
                lsCadImp = lsCadImp & Space(1) & oFun.ImpreFormat(nItem, 4, 0) & Space(1) & !Fecha & Space(1) & _
                           sCliente & Space(1) & oFun.ImpreFormat(!Pagado, 10, 2) & Space(1) & oFun.ImpreFormat(!Capital, 10, 2) & _
                           Space(1) & oFun.ImpreFormat(!InteresCom, 10, 2) & Space(1) & oFun.ImpreFormat(!InteresMor, 10, 2) & _
                           Space(1) & oFun.ImpreFormat(!Gastos, 10, 2) & Space(1) & oFun.ImpreFormat(!SaldoCapital, 10, 2) & _
                           Space(1) & oFun.ImpreFormat(!ComisionAbog, 10, 2) & oFunI.gPrnSaltoLinea
                
                 
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE PAGOS DE CREDITOS POR ANALISTA", lsCondiciones, lnPage, 145, "", 138004)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                'RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            'RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO RECUPERAC." & Space(23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal41, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal71, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    
        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(23) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal7, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X  ANALISTA" & Space(29) & oFun.ImpreFormat(sSubTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal3, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal5, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal6, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal7, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
      
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
    nRepo138004_ListadoPagosxAnalista = lsCadBuffer
End Function
  
Public Function nRepo138011_ListadoCreditosNuevos(ByVal pdFecIni As String, ByVal pdFecFin As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lsNegritaOn As String
Dim lsNegritaOff As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim nItem As Integer
Dim lnPage As Integer
Dim loImprimeCab As New COMNColoCPig.NCOMColPImpre
        
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
 

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
 
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

lsCondiciones = "Del " & pdFecIni & " Al " & pdFecFin
         
lsSQL = " Select T.cCtaCod, C.cAgeCodAct Agencia, A.cAgeDescripcion, T.nCapital, T.nInteres, "
lsSQL = lsSQL & " T.nInteresMor, T.nGastos, C.dVigencia, CR.dIngRecup, C.cLineaCred, P.cPersNombre,"
lsSQL = lsSQL & " SUBSTRING(T.cCtaCod, 9, 1) AS Moneda, CON.cConsDescripcion AS sMoneda,  "
lsSQL = lsSQL & " (  SELECT RH.cUser FROM RRHH RH INNER JOIN PRODUCTOPERSONA PP ON PP.cPersCod = RH.cPersCod "
lsSQL = lsSQL & " Where PP.cCtaCod = T.cCtaCod And PP.nPrdPersRelac = 28) AS ANALISTA, C.nMontoCol "
lsSQL = lsSQL & " From "
lsSQL = lsSQL & " ( Select CR.cCtaCod, "
lsSQL = lsSQL & " IsNull( "
lsSQL = lsSQL & " ( SELECT SUM(nmonto) "
lsSQL = lsSQL & " FROM movcoldet MCD INNER JOIN MOV M ON MCD.nMovNro = M.nMovNro "
lsSQL = lsSQL & " WHERE M.nMovFlag=0 and MCD.cCtaCod = CR.cCtaCod AND (MCD.nPrdConceptoCod = '3000' OR MCD.nPrdConceptoCod LIKE '14%')And MCD.cOpecod = '130100' "
'----------
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
lsSQL = lsSQL & " INNER JOIN ProductoPersona PP ON T.cCtaCod = PP.cCtaCod "
lsSQL = lsSQL & " INNER JOIN Persona P ON PP.cPersCod = P.cPersCod "
lsSQL = lsSQL & " INNER JOIN Agencias A ON A.cAgeCod =  C.cAgeCodAct " 'BRGO BASILEA II
lsSQL = lsSQL & " INNER JOIN Producto PC ON T.cCtaCod = PC.cCtaCod "
lsSQL = lsSQL & " INNER JOIN Constante CON ON SUBSTRING(T.cCtaCod, 9, 1) = CON.nConsValor "
lsSQL = lsSQL & " Where(PP.nPrdPersRelac =" & gColRelPersTitular & ") And (CON.nConsCod=" & gMoneda & ") "
lsSQL = lsSQL & " AND (PC.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud
lsSQL = lsSQL & "', '" & gColocEstRecVigCast & "', '" & gColocEstRecCanCast & "'))"
lsSQL = lsSQL & " Order By C.cAgeCodAct, SUBSTRING(T.cCtaCod, 9, 1)"
  
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
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
        lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
        
        lnLineas = lnLineas + 4
        lnIndice = lnIndice + 4
        
        With lrDataRep
            'RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                
                If sTempoAgencia <> !Agencia Then
                        
                    lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA " & Space(65) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 8, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal31, 8, 2) & Space(1) & oFun.ImpreFormat(sTotal4, 8, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
  
                    'Agencia
                    lsCadImp = lsCadImp & lsNegritaOn & "Agencia: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                      
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "Moneda: " & lrDataRep!sMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                    
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
                    
                        lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
 '
                        'Moneda
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA " & Space(65) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 8, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal31, 8, 2) & Space(1) & oFun.ImpreFormat(sTotal4, 8, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
  

                        lsCadImp = lsCadImp & lsNegritaOn & "Moneda: " & lrDataRep!sMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea '& ofunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                         
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
                
                sCliente = !cPersNombre
                       
                lsCadImp = lsCadImp & Space(1) & oFun.ImpreFormat(nItem, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dVigencia, "dd/MM/YYYY") & _
                            Space(1) & Format(!dIngRecup, "dd/MM/YYYY") & Space(1) & _
                            sCliente & Space(1) & !Analista & Space(1) & Space(1) & _
                            oFun.ImpreFormat(!nMontoCol, 10, 2) & Space(1) & oFun.ImpreFormat(!nCapital, 10, 2) & Space(1) & _
                            oFun.ImpreFormat(!nInteres, 8, 2) & Space(1) & oFun.ImpreFormat(!nInteresMor, 8, 2) & Space(1) & _
                            oFun.ImpreFormat(!nGastos, 8, 2) & Space(1) & _
                            oFun.ImpreFormat(!nCapital + !nInteres + !nInteresMor + !nGastos, 10, 2) & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS NUEVOS ", lsCondiciones, lnPage, 165, "", 138011)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                'RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            'RaiseEvent CloseProgress
        End With
       
        lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA " & Space(65) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 8, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal31, 8, 2) & Space(1) & oFun.ImpreFormat(sTotal4, 8, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea

        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
    nRepo138011_ListadoCreditosNuevos = lsCadBuffer
End Function

Public Function nRepo138013_(ByVal sBuscarExp As Byte, ByVal sConSinExp As Byte, ByVal sMoneda As Byte, _
ByVal sEstados As String, ByVal sMasCero As Byte, ByVal sCero As Byte, _
ByVal pdf1 As Date, pdf2 As Date, Optional ByVal psImpresora As Impresoras = gEPSON) As String


Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As New COMNColoCPig.NCOMColPImpre
         
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


Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
 
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
        
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
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
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'Agencia
        lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'Fuente de Financiamiento
        lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
        
        lnLineas = lnLineas + 6
        lnIndice = lnIndice + 6
         
        With lrDataRep
            'RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                                
                'TOTALizar por MONEDA
                If sTempoMoneda <> !Moneda Then
                
                    lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 9, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal41, 9, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 9, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea
                    
                    lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X AGENCIA" & Space(78) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 9, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal4, 9, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 9, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea
                    
                    lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(79) & oFun.ImpreFormat(sSubTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal3, 9, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal4, 9, 2) & Space(1) & oFun.ImpreFormat(sSubTotal5, 9, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea

                    'MONEDA
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    'Agencia
                    lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    'Fuente de Financiamiento
                    lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                    
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
                
                        lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 9, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal41, 9, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 9, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea
                        
                        lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X AGENCIA" & Space(78) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 9, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal4, 9, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 9, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                         
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
                        lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
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
                    
                            lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                                        Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 9, 2) & _
                                        Space(1) & oFun.ImpreFormat(sTotal41, 9, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 9, 2) & _
                                        Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                            
                            'Fuente de Financiamiento
                            lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                             
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
                
                sCliente = !cPersNombre
                
                lsCadImp = lsCadImp & oFun.ImpreFormat(nItem, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dVigencia, "dd/MM/YYYY") & _
                            Space(1) & Format(!dIngRecup, "dd/MM/YYYY") & Space(1) & !cLineaCred & Space(1) & _
                            sCliente & Space(1) & !Analista & Space(1) & Space(1) & _
                            oFun.ImpreFormat(!nMontoCol, 10, 2) & Space(1) & oFun.ImpreFormat(!SaldoCapital, 10, 2) & Space(1) & _
                            oFun.ImpreFormat(!Interes, 9, 2) & Space(1) & oFun.ImpreFormat(!Mora, 9, 2) & Space(1) & _
                            oFun.ImpreFormat(!Gasto, 9, 2) & Space(1) & oFun.ImpreFormat(!SaldoCapital + !Interes + !Mora + !Gasto, 10, 2) & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS CASTIGADOS", lsCondiciones, lnPage, 175, "", 138012)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                'RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            'RaiseEvent CloseProgress
        End With
        lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 9, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal41, 9, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 9, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(79) & oFun.ImpreFormat(sSubTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal3, 9, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal4, 9, 2) & Space(1) & oFun.ImpreFormat(sSubTotal5, 9, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X AGENCIA" & Space(78) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 9, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal4, 9, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 9, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea
        
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
   nRepo138013_ = lsCadBuffer


End Function
'
Public Function nRepo138012_ListadoCreditosEnJudicial(ByVal sBuscarExp As Byte, ByVal sConSinExp As Byte, ByVal sMoneda As Byte, _
                                                        ByVal sEstados As String, ByVal sMasCero As Byte, ByVal sCero As Byte, _
                                                        sRotulo As String, Optional ByVal sAgencias As String = "", Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As New COMNColoCPig.NCOMColPImpre
         
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


Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
 
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

    'Si se busca expediente
    If sBuscarExp = 1 Then
        sCriterio = " AND (CR.nDemanda=" & sConSinExp & ")"
    End If
    
    'Busco por estados en la tabla Producto
    sCriterio = sCriterio & " AND (P.nPrdEstado IN(" & sEstados & ")) "
    
    'buscar MONEDA
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
        
        If sAgencias <> "" Then
            'sCriterio = sCriterio & " AND Substring(P.cCtaCod,4,2) in (" & sAgencias & ")"
            '*** BRGO BASILEA II
            sCriterio = sCriterio & " AND C.cAgeCodAct in (" & sAgencias & ")"
        End If
    End If
    
    '*** BRGO BASILEA II
    lsSQL = " SELECT P.cCtaCod, sExpe = case when CR.nDemanda = 1 then 'S' when CR.nDemanda =2 then 'N' end, C.cAgeCodAct AS Agencia, A.cAgeDescripcion, " & _
            " C.dVigencia, CR.dIngRecup, C.cLineaCred, CLC.cDescripcion, Persona.cPersNombre, " & _
            " (SELECT RHH.CUSER FROM RRHH RHH INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = RHH.CPERSCOD " & _
            " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS ANALISTA, C.nMontoCol, P.nSaldo AS SaldoCapital, " & _
            " CR.nSaldoIntComp As Interes, CR.nSaldoIntMor AS Mora, CR.nSaldoGasto AS Gasto, SUBSTRING(P.cCtaCod, 9, 1) AS Moneda, " & _
            " CON.cConsDescripcion AS sMoneda, SUBSTRING(C.cLineaCred, 1, 2) AS cCodFuente, (SELECT CLC.cDescripcion " & _
            " FROM ColocLineaCredito CLC WHERE substring(C.clineacred, 1, 2) = CLC.cLineaCred) AS cDesFuente " & _
            " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod INNER JOIN " & _
            " ColocRecup CR ON C.cCtaCod = CR.cCtaCod INNER JOIN ColocLineaCredito CLC ON C.cLineaCred = CLC.cLineaCred INNER JOIN " & _
            " ProductoPersona PP ON P.cCtaCod = PP.cCtaCod INNER JOIN Persona ON PP.cPersCod = Persona.cPersCod INNER JOIN " & _
            " Constante CON ON SUBSTRING(P.cCtaCod, 9, 1) = CON.nConsValor INNER JOIN Agencias A ON A.cAgeCod = C.cAgeCodAct " & _
            " Where (PP.nPrdPersRelac=" & gColRelPersTitular & ") And (CON.nConsCod=" & gMoneda & ") " & _
            sCriterio & " ORDER BY SUBSTRING(C.cLineaCred, 1, 2),  SUBSTRING(P.cCtaCod, 9, 1), C.cAgeCodAct  "
    '******************************
        
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
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
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'Agencia
        lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'Fuente de Financiamiento
        lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & lsNegritaOn & sRotulo & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & String(175 - 19, "-") & lsNegritaOff & oFunI.gPrnSaltoLinea
        
        
        lnLineas = lnLineas + 7
        lnIndice = lnIndice + 7
         
        With lrDataRep
            'RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                                
                'TOTALizar por MONEDA
                If sTempoMoneda <> !Moneda Then
                
                    lsCadImp = lsCadImp & String(175 - 19, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61 - 23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 9, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal41, 9, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 9, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea
                    
                    lsCadImp = lsCadImp & String(175 - 19, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X AGENCIA" & Space(78 - 23) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 9, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal4, 9, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 9, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea
                    
                    lsCadImp = lsCadImp & String(175 - 19, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(79 - 23) & oFun.ImpreFormat(sSubTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal3, 9, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal4, 9, 2) & Space(1) & oFun.ImpreFormat(sSubTotal5, 9, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea

                    'MONEDA
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & lrDataRep!sMoneda & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    'Agencia
                    lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    'Fuente de Financiamiento
                    lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & sRotulo & oFunI.gPrnSaltoLinea
                    
                    lsCadImp = lsCadImp & String(175 - 19, "-") & oFunI.gPrnSaltoLinea
                    
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
                
                        lsCadImp = lsCadImp & String(175 - 19, "-") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61 - 23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 9, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal41, 9, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 9, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea
                        
                        lsCadImp = lsCadImp & String(175 - 19, "-") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X AGENCIA" & Space(78 - 23) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 9, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal4, 9, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 9, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                         
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
                        lsCadImp = lsCadImp & lsNegritaOn & "AGENCIA: " & lrDataRep!Agencia & " : " & lrDataRep!cAgeDescripcion & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & sRotulo & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & String(175 - 19, "-") & oFunI.gPrnSaltoLinea
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
                    
                            lsCadImp = lsCadImp & String(175 - 19, "-") & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61 - 23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                                        Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 9, 2) & _
                                        Space(1) & oFun.ImpreFormat(sTotal41, 9, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 9, 2) & _
                                        Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                            
                            'Fuente de Financiamiento
                            lsCadImp = lsCadImp & lsNegritaOn & "FUENTE DE FINANCIAMIENTO: " & lrDataRep!cCodFuente & " : " & lrDataRep!cDesFuente & lsNegritaOff & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & lsNegritaOn & sRotulo & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & String(175 - 19, "-") & oFunI.gPrnSaltoLinea
                             
'                            sTotal11 = 0
'                            sTotal21 = 0
'                            sTotal31 = 0
'                            sTotal41 = 0
'                            sTotal51 = 0
'                            sTotal61 = 0
                            
                            sTotal11 = !nMontoCol
                            sTotal21 = !SaldoCapital
                            sTotal31 = !Interes
                            sTotal41 = !Mora
                            sTotal51 = !Gasto
                            sTotal61 = !SaldoCapital + !Interes + !Mora + !Gasto
                            
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
                
                sCliente = !cPersNombre
                
                lsCadImp = lsCadImp & oFun.ImpreFormat(nItem, 4, 0) & Space(1) & !cCtaCod & _
                            Space(1) & Format(!dIngRecup, "dd/MM/YYYY") & Space(1) & _
                            sCliente & Space(1) & IIf(!Analista = "", Space(4), !Analista) & Space(1) & Space(1) & _
                            oFun.ImpreFormat(!nMontoCol, 10, 2) & Space(1) & oFun.ImpreFormat(!SaldoCapital, 10, 2) & Space(1) & _
                            oFun.ImpreFormat(!Interes, 9, 2) & Space(1) & oFun.ImpreFormat(!Mora, 9, 2) & Space(1) & _
                            oFun.ImpreFormat(!Gasto, 9, 2) & Space(1) & oFun.ImpreFormat(!SaldoCapital + !Interes + !Mora + !Gasto, 10, 2) & Space(2) & !sExpe & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS EN JUDICIAL ", lsCondiciones, lnPage, 175 - 19, "", 138012)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                'RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            'RaiseEvent CloseProgress
        End With
    lsCadImp = lsCadImp & String(175 - 19, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X FUENTE DE FINANCIAMIENTO" & Space(61 - 23) & oFun.ImpreFormat(sTotal11, 10, 2) & _
                Space(1) & oFun.ImpreFormat(sTotal21, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal31, 9, 2) & _
                Space(1) & oFun.ImpreFormat(sTotal41, 9, 2) & Space(1) & oFun.ImpreFormat(sTotal51, 9, 2) & _
                Space(1) & oFun.ImpreFormat(sTotal61, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea
    
    lsCadImp = lsCadImp & String(175 - 19, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X AGENCIA" & Space(78 - 23) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 9, 2) & _
                Space(1) & oFun.ImpreFormat(sTotal4, 9, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 9, 2) & _
                Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea
    
    lsCadImp = lsCadImp & String(175 - 19, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA" & Space(79 - 23) & oFun.ImpreFormat(sSubTotal1, 10, 2) & _
                Space(1) & oFun.ImpreFormat(sSubTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal3, 9, 2) & _
                Space(1) & oFun.ImpreFormat(sSubTotal4, 9, 2) & Space(1) & oFun.ImpreFormat(sSubTotal5, 9, 2) & _
                Space(1) & oFun.ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea

        
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
    nRepo138012_ListadoCreditosEnJudicial = lsCadBuffer
End Function

Public Function nRepo138014_ListadoCreditosxAnalista(ByVal sConAnalista As Byte, ByVal sCodAnalista As String, ByVal sTipC As Double, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As New COMNColoCPig.NCOMColPImpre
        
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

Dim sTipoBusqueda As String

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
  

sTipoBusqueda = Right(Trim(sCodAnalista), 1)
sCodAnalista = Left(sCodAnalista, (Len(sCodAnalista) - 1))

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

'*** BRGO BASILEA II
      lsSQL = "SELECT P.cCtaCod, substring(P.cCtaCod, 9,1) as cMoneda, C.dVigencia, CR.dIngRecup, C.cLineaCred, CLC.cDescripcion, C.cTpoCredCod AS cCodTipoCredito, Persona.cPersCod, " & _
                  " Persona.cPersNombre, Persona.cPersDireccUbiGeo, (SELECT PrP.cPersCod FROM ProductoPersona PrP WHERE PrP.cCtaCod = PP.cCtaCod AND nPrdPersRelac=" & gColRelPersAnalista & ") AS CodAnalista, (SELECT PERSO.cPersNombre " & _
                  " FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS ANALISTA, " & _
                  " (SELECT PERSO.cPersNombre FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD " & _
                  " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ") AS ABOGADO, C.nMontoCol AS Prestamo, P.nSaldo AS Saldo, " & _
                  " CR.nSaldoIntComp AS Interes, CR.nSaldoIntMor AS Mora, CR.nSaldoGasto AS Gasto, SUBSTRING(P.cCtaCod, 9,1) AS Moneda, " & _
                  " CON.cConsDescripcion AS sMoneda, (SELECT COO.cConsDescripcion FROM Constante COO " & _
                  " WHERE C.cTpoCredCod = COO.nConsValor AND COO.nConsCod = 3034) AS cDesTipoCredito, PP.nPrdPersRelac " & _
                  " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod INNER JOIN " & _
                  " ColocRecup CR ON C.cCtaCod = CR.cCtaCod INNER JOIN ColocLineaCredito CLC ON C.cLineaCred = CLC.cLineaCred INNER JOIN " & _
                  " ProductoPersona PP ON P.cCtaCod = PP.cCtaCod INNER JOIN " & _
                  " Persona ON PP.cPersCod = Persona.cPersCod INNER JOIN Constante CON ON  substring(P.cCtaCod, 9,1) = CON.nConsValor " & _
                  " Where (PP.NPRDPERSRELAC=" & gColRelPersTitular & ") And (CON.nconscod=" & gMoneda & ") "
      If sTipoBusqueda = "1" Then
          lsSQL = lsSQL & " AND (P.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecVigCast & "','2205')) "
          '-------
      ElseIf sTipoBusqueda = "2" Then
          lsSQL = lsSQL & " AND (P.nPrdEstado IN('" & gColocEstRecVigJud & "','2205')) "
          '------
      ElseIf sTipoBusqueda = "3" Then
          lsSQL = lsSQL & " AND (P.nPrdEstado IN('" & gColocEstRecVigCast & "')) "
      End If
      lsSQL = lsSQL & sCriterio & " ORDER BY (SELECT PrP.cPersCod FROM ProductoPersona PrP WHERE PrP.cCtaCod = PP.cCtaCod AND nPrdPersRelac=" & gColRelPersAnalista & "),  C.cTpoCredCod, SUBSTRING(P.cCtaCod, 9, 1), P.cCtaCod"

'************************
    
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
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
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS X ANALISTA ", lsCondiciones, lnPage, 172 - 12, "//T.C. :" & Format(sTipC, "0.00"), 138014)

        lnIndice = 0:  lnLineas = 7
        sTempoTipoCredito = lrDataRep!cCodTipoCredito
        sTempoAnalista = "" & lrDataRep!CodAnalista
        sTempoMoneda = "" & lrDataRep!cMoneda
        
        'Analista
        lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA: " & lrDataRep!Analista & " : " & lrDataRep!CodAnalista & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'Producto
        lsCadImp = lsCadImp & lsNegritaOn & "TIPO DE CREDITO: " & lrDataRep!cDesTipoCredito & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
        
        lnLineas = lnLineas + 5
        lnIndice = lnIndice + 5
          
        With lrDataRep
            'RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                'Totalizar por analista
                
                If sTempoAnalista <> !CodAnalista Then
                
                    lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & oFun.ImpreFormat(sTotalA, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotalB, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalC, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotalD, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalE, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                    lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO DE CREDITO" & Space(54 - 12) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                    lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ANALISTA" & Space(61 - 12) & oFun.ImpreFormat(sSubTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal3, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal5, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                     
                    'Analista
                    lsCadImp = lsCadImp & lsNegritaOn & "ANALISTA: " & lrDataRep!Analista & " : " & lrDataRep!CodAnalista & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                    'Producto
                    lsCadImp = lsCadImp & lsNegritaOn & "TIPO DE CREDITO: " & !cDesTipoCredito & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                    
                    
                    sSubTotal1 = (!PRESTAMO * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal2 = (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal3 = (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal4 = (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal5 = (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal6 = (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                    
                    sTotal1 = (!PRESTAMO * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal2 = (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal3 = (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal4 = (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal5 = (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal6 = (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                    
                    sTotalA = !PRESTAMO
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
                                        
                    sSubTotal1 = sSubTotal1 + (!PRESTAMO * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal2 = sSubTotal2 + (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal3 = sSubTotal3 + (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal4 = sSubTotal4 + (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal5 = sSubTotal5 + (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal6 = sSubTotal6 + (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                    
                    
                    If sTempoTipoCredito <> lrDataRep!cCodTipoCredito Then
                        lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & oFun.ImpreFormat(sTotalA, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotalB, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalC, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotalD, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalE, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        
                        lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO DE CREDITO" & Space(54 - 12) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        
                        'Producto
                        lsCadImp = lsCadImp & lsNegritaOn & "TIPO DE CREDITO: " & !cDesTipoCredito & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        
                        'Moneda
                        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                        
                        
                        sTotal1 = (!PRESTAMO * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal2 = (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal3 = (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal4 = (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal5 = (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal6 = (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                        
                        sTotalA = !PRESTAMO
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
                
                        sTotal1 = sTotal1 + (!PRESTAMO * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal2 = sTotal2 + (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal3 = sTotal3 + (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal4 = sTotal4 + (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal5 = sTotal5 + (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal6 = sTotal6 + ((!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                 
                        If sTempoMoneda <> !cMoneda Then
                
                            lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & oFun.ImpreFormat(sTotalA, 10, 2) & _
                                        Space(1) & oFun.ImpreFormat(sTotalB, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalC, 10, 2) & _
                                        Space(1) & oFun.ImpreFormat(sTotalD, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalE, 10, 2) & _
                                        Space(1) & oFun.ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        
                            'Moneda
                            lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
 
                            sTotalA = !PRESTAMO
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
                            sTotalA = sTotalA + !PRESTAMO
                            sTotalB = sTotalB + !Saldo
                            sTotalC = sTotalC + !Interes
                            sTotalD = sTotalD + !Mora
                            sTotalE = sTotalE + !Gasto
                            sTotalF = sTotalF + !Saldo + !Interes + !Mora + !Gasto
                        End If
                    End If
                End If
                
                sCliente = !cPersNombre
                sAbogado = "" & !Abogado
                sProducto = !cDesTipoCredito
                
                lsCadImp = lsCadImp & Space(1) & oFun.ImpreFormat(nItem, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dIngRecup, "dd/MM/YYYY") & _
                              Space(1) & _
                            sCliente & Space(1) & _
                            oFun.ImpreFormat(!PRESTAMO, 10, 2) & Space(1) & oFun.ImpreFormat(!Saldo, 10, 2) & Space(1) & _
                            oFun.ImpreFormat(!Interes, 10, 2) & Space(1) & oFun.ImpreFormat(!Mora, 10, 2) & Space(1) & _
                            oFun.ImpreFormat(!Gasto, 10, 2) & Space(1) & oFun.ImpreFormat(!Saldo + !Interes + !Mora + !Gasto, 10, 2) & Space(1) & _
                            Trim(Left(sAbogado, 10)) & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 53 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS X ANALISTA ", lsCondiciones, lnPage, 172 - 12, "//T.C. :" & Format(sTipC, "0.00"), 138014)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                'RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            'RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & oFun.ImpreFormat(sTotalA, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotalB, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalC, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotalD, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalE, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO DE CREDITO" & Space(54 - 12) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ANALISTA" & Space(61 - 12) & oFun.ImpreFormat(sSubTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal3, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal5, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
 
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
    nRepo138014_ListadoCreditosxAnalista = lsCadBuffer
End Function

Public Function nRepo138021_ListadoCreditosEnCJxAbogado(ByVal sConAbogado As Byte, ByVal sCodAbogado As String, ByVal sTipC As Double, Optional ByVal psImpresora As Impresoras = gEPSON) As String
  
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim loImprimeCab As New COMNColoCPig.NCOMColPImpre
         
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


Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

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

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
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
        lsCadImp = lsCadImp & lsNegritaOn & "ABOGADO:  " & lrDataRep!Abogado & " : " & lrDataRep!CODABOGADO & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(109, "-") & oFunI.gPrnSaltoLinea
        
        lnLineas = lnLineas + 3
        lnIndice = lnIndice + 3
          
        With lrDataRep
            'RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                  
                'Totalizar por analista
                
                If sTempoAbogado <> !CODABOGADO Then
                
                        lsCadImp = lsCadImp & String(109, "-") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ABOG." & Space(11) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                    'Abogado
                    lsCadImp = lsCadImp & lsNegritaOn & "ABOGADO:  " & lrDataRep!Abogado & " : " & lrDataRep!CODABOGADO & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & String(109, "-") & oFunI.gPrnSaltoLinea
                      
                    sTotal1 = !PRESTAMO
                    sTotal2 = !Saldo
                    sTotal3 = !Interes
                    sTotal4 = !Mora
                    sTotal5 = !Gasto
                    sTotal6 = !Saldo + !Interes + !Mora + !Gasto
                    
                    sTempoAbogado = lrDataRep!CODABOGADO
                      
                    lnLineas = lnLineas + 6
                    lnIndice = lnIndice + 6
                Else
                     
                    sTotal1 = sTotal1 + !PRESTAMO
                    sTotal2 = sTotal2 + !Saldo
                    sTotal3 = sTotal3 + !Interes
                    sTotal4 = sTotal4 + !Mora
                    sTotal5 = sTotal5 + !Gasto
                    sTotal6 = sTotal6 + !Saldo + !Interes + !Mora + !Gasto
                End If
                
                sTotalA = sTotalA + !PRESTAMO
                sTotalB = sTotalB + !Saldo
                sTotalC = sTotalC + !Interes
                sTotalD = sTotalD + !Mora
                sTotalE = sTotalE + !Gasto
                sTotalF = sTotalF + !Saldo + !Interes + !Mora + !Gasto
                
                sDescripcion = Replace(!DesEstadoRecuperacion, "VIGENTE ", "")
                
                lsCadImp = lsCadImp & sDescripcion & Space(1) & _
                            oFun.ImpreFormat(!PRESTAMO, 10, 2) & Space(1) & oFun.ImpreFormat(!Saldo, 10, 2) & Space(1) & _
                            oFun.ImpreFormat(!Interes, 10, 2) & Space(1) & oFun.ImpreFormat(!Mora, 10, 2) & Space(1) & _
                            oFun.ImpreFormat(!Gasto, 10, 2) & Space(1) & oFun.ImpreFormat(!Saldo + !Interes + !Mora + !Gasto, 10, 2) & oFunI.gPrnSaltoLinea
                
                'lnLineas = lnLineas + 1
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS EN C.J. POR ABOGADO(Consolidado)", lsCondiciones, lnPage, 109, "//T.C. :" & Format(sTipC, "0.00"), 138021)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                'RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            'RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(109, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ABOG." & Space(11) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & String(109, "=") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTAL          " & Space(11) & oFun.ImpreFormat(sTotalA, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotalB, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalC, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotalD, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalE, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea
    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
    nRepo138021_ListadoCreditosEnCJxAbogado = lsCadBuffer
End Function
Public Function nRepo138024_ListadoExpedientesxAbogado(ByVal sConAbogado As Byte, ByVal sCodAbogado As String, ByVal sTipC As Double, Optional psImpresora As Impresoras = gEPSON) As String
 
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As New COMNColoCPig.NCOMColPImpre
        
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


Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
   
  
lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)

    'Si se busca analista
    If sConAbogado = 1 Then
        
        sCriterio = " AND ((Select CPERSCOD FROM  " & _
                    " PRODUCTOPERSONA PRP  " & _
                    " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ")='" & Trim(sCodAbogado) & "')"
        lsCondiciones = "Estudio: " & sCodAbogado & " : "
        
    End If
    
    'ARCV se agregaron las funciones para calculo de mora e interes actual
    Dim oCons As COMDConstSistema.NCOMConstSistema
    Dim dFecSisTmp As Date

    Set oCons = New COMDConstSistema.NCOMConstSistema
    dFecSisTmp = CDate(oCons.LeeConstSistema(gConstSistFechaSistema))
    Set oCons = Nothing
    '-----------------
    
    lsSQL = "SELECT P.cCtaCod, substring(P.cCtaCod, 9,1) as cMoneda, C.dVigencia, CR.dIngRecup, C.cLineaCred, CLC.cDescripcion, C.cTpoCredCod AS cCodTipoCredito, Persona.cPersCod, " & _
            " Persona.cPersNombre, Persona.cPersDireccUbiGeo, (SELECT PrP.cPersCod FROM ProductoPersona PrP WHERE PrP.cCtaCod = PP.cCtaCod AND nPrdPersRelac=" & gColRelPersEstudioJuridico & ") AS CodAnalista, (SELECT PERSO.cPersNombre " & _
            " FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersEstudioJuridico & ") AS ABOGADO, " & _
            " (SELECT cUser FROM PERSONA PERSO   INNER JOIN RRHH RH ON PERSO.CPERSCOD=RH.CPERSCOD  INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD " & _
            " WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=" & gColRelPersAnalista & ") AS ANALISTA, C.nMontoCol AS Prestamo, P.nSaldo AS Saldo, " & _
            " CR.nSaldoIntComp + dbo.ColocRecupCalculoInteresActual(P.cCtaCod,'" & Format(dFecSisTmp, "mm/dd/yyyy") & "') AS Interes, CR.nSaldoIntMor + dbo.ColocRecupCalculoMoraActual(P.cCtaCod,'" & Format(dFecSisTmp, "mm/dd/yyyy") & "') AS Mora, CR.nSaldoGasto AS Gasto, SUBSTRING(P.cCtaCod, 9,1) AS Moneda, " & _
            " CON.cConsDescripcion AS sMoneda, (SELECT COO.cConsDescripcion FROM Constante COO " & _
            " WHERE C.cTpoCredCod = COO.nConsValor AND COO.nConsCod=3034) AS cDesTipoCredito, PP.nPrdPersRelac " & _
            " FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod INNER JOIN " & _
            " ColocRecup CR ON C.cCtaCod = CR.cCtaCod INNER JOIN ColocLineaCredito CLC ON C.cLineaCred = CLC.cLineaCred INNER JOIN " & _
            " ProductoPersona PP ON P.cCtaCod = PP.cCtaCod INNER JOIN " & _
            " Persona ON PP.cPersCod = Persona.cPersCod INNER JOIN Constante CON ON  substring(P.cCtaCod, 9,1) = CON.nConsValor " & _
            " Where (PP.NPRDPERSRELAC=" & gColRelPersTitular & ") And (CON.nconscod=" & gMoneda & ") AND (P.nPrdEstado IN('" & gColocEstRecVigJud & "','" & gColocEstRecCanJud & "', '" & gColocEstRecVigCast & "', '" & gColocEstRecCanCast & "')) " & _
            sCriterio & " ORDER BY (SELECT PrP.cPersCod FROM ProductoPersona PrP WHERE PrP.cCtaCod = PP.cCtaCod AND nPrdPersRelac=" & gColRelPersEstudioJuridico & "),  C.cTpoCredCod, SUBSTRING(P.cCtaCod, 9, 1), P.cCtaCod"


    Set loDataRep = New COMDColocPig.DCOMColPFunciones
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
        lsCadImp = lsCadImp & nRepoCabecera("CREDITOS EN C.J. POR ABOGADO(Detalle)", lsCondiciones, lnPage, 172 - 12, "//T.C. :" & Format(sTipC, "0.00"), 138024)
                                             
        lnIndice = 0:  lnLineas = 7
        sTempoTipoCredito = lrDataRep!cCodTipoCredito
        sTempoAnalista = "" & lrDataRep!CodAnalista
        sTempoMoneda = "" & lrDataRep!cMoneda
        
        'Analista
        lsCadImp = lsCadImp & lsNegritaOn & "ABOGADO: " & lrDataRep!Abogado & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'Producto
        lsCadImp = lsCadImp & lsNegritaOn & "TIPO DE CREDITO: " & lrDataRep!cDesTipoCredito & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'Moneda
        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
        
        lnLineas = lnLineas + 5
        lnIndice = lnIndice + 5
          
        With lrDataRep
            'RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                'Totalizar por analista
                
                If sTempoAnalista <> !CodAnalista Then
                
                    lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & oFun.ImpreFormat(sTotalA, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotalB, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalC, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotalD, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalE, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                    lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO DE CREDITO" & Space(54 - 12) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                    lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ABOGADO " & Space(61 - 12) & oFun.ImpreFormat(sSubTotal1, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal3, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal5, 10, 2) & _
                                Space(1) & oFun.ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                     
                    'Analista
                    lsCadImp = lsCadImp & lsNegritaOn & "ABOGADO: " & lrDataRep!Analista & " : " & lrDataRep!CodAnalista & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                    'Producto
                    lsCadImp = lsCadImp & lsNegritaOn & "TIPO DE CREDITO: " & !cDesTipoCredito & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                    'Moneda
                    lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                    
                    
                    sSubTotal1 = (!PRESTAMO * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal2 = (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal3 = (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal4 = (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal5 = (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal6 = (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                    
                    sTotal1 = (!PRESTAMO * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal2 = (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal3 = (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal4 = (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal5 = (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sTotal6 = (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                    
                    sTotalA = !PRESTAMO
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
                                        
                    sSubTotal1 = sSubTotal1 + (!PRESTAMO * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal2 = sSubTotal2 + (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal3 = sSubTotal3 + (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal4 = sSubTotal4 + (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal5 = sSubTotal5 + (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                    sSubTotal6 = sSubTotal6 + (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                    
                    
                    If sTempoTipoCredito <> lrDataRep!cCodTipoCredito Then
                        lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & oFun.ImpreFormat(sTotalA, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotalB, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalC, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotalD, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalE, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        
                        lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO DE CREDITO" & Space(54 - 12) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                                    Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        
                        'Producto
                        lsCadImp = lsCadImp & lsNegritaOn & "TIPO DE CREDITO: " & !cDesTipoCredito & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        
                        'Moneda
                        lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & oFunI.gPrnSaltoLinea
                        lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                        
                        
                        sTotal1 = (!PRESTAMO * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal2 = (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal3 = (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal4 = (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal5 = (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal6 = (!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC)
                        
                        sTotalA = !PRESTAMO
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
                
                        sTotal1 = sTotal1 + (!PRESTAMO * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal2 = sTotal2 + (!Saldo * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal3 = sTotal3 + (!Interes * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal4 = sTotal4 + (!Mora * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal5 = sTotal5 + (!Gasto * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                        sTotal6 = sTotal6 + ((!Saldo + !Interes + !Mora + !Gasto) * IIf(lrDataRep!cMoneda = 1, 1, sTipC))
                 
                        If sTempoMoneda <> !cMoneda Then
                
                            lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & oFun.ImpreFormat(sTotalA, 10, 2) & _
                                        Space(1) & oFun.ImpreFormat(sTotalB, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalC, 10, 2) & _
                                        Space(1) & oFun.ImpreFormat(sTotalD, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalE, 10, 2) & _
                                        Space(1) & oFun.ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                        
                            'Moneda
                            lsCadImp = lsCadImp & lsNegritaOn & "MONEDA: " & IIf(lrDataRep!cMoneda = 1, "Soles", "Dolares") & lsNegritaOff & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
 
                            sTotalA = !PRESTAMO
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
                            sTotalA = sTotalA + !PRESTAMO
                            sTotalB = sTotalB + !Saldo
                            sTotalC = sTotalC + !Interes
                            sTotalD = sTotalD + !Mora
                            sTotalE = sTotalE + !Gasto
                            sTotalF = sTotalF + !Saldo + !Interes + !Mora + !Gasto
                        End If
                    End If
                End If
                
                sCliente = !cPersNombre
                sAbogado = "" & !Abogado
                sProducto = !cDesTipoCredito
                
                lsCadImp = lsCadImp & Space(1) & oFun.ImpreFormat(nItem, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dIngRecup, "dd/MM/YYYY") & _
                              Space(1) & _
                            sCliente & Space(1) & _
                            oFun.ImpreFormat(!PRESTAMO, 10, 2) & Space(1) & oFun.ImpreFormat(!Saldo, 10, 2) & Space(1) & _
                            oFun.ImpreFormat(!Interes, 10, 2) & Space(1) & oFun.ImpreFormat(!Mora, 10, 2) & Space(1) & _
                            oFun.ImpreFormat(!Gasto, 10, 2) & Space(1) & oFun.ImpreFormat(!Saldo + !Interes + !Mora + !Gasto, 10, 2) & Space(1) & _
                            !Analista & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 53 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS EN C.J. POR ABOGADO(Detalle)", lsCondiciones, lnPage, 172 - 12, "//T.C. :" & Format(sTipC, "0.00"), 138024)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                'RaiseEvent Progress(.Bookmark, .RecordCount)
                 
                .MoveNext
                
            Loop
            'RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X MONEDA         " & Space(54 - 12) & oFun.ImpreFormat(sTotalA, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotalB, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalC, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotalD, 10, 2) & Space(1) & oFun.ImpreFormat(sTotalE, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotalF, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X TIPO DE CREDITO" & Space(54 - 12) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal3, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sTotal5, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & String(172 - 12, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES X ABOGADO " & Space(61 - 12) & oFun.ImpreFormat(sSubTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal2, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal3, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal4, 10, 2) & Space(1) & oFun.ImpreFormat(sSubTotal5, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sSubTotal6, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
 
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
    nRepo138024_ListadoExpedientesxAbogado = lsCadBuffer
End Function
 
Public Function nRepo138023_ListadoSaldoCapital(ByVal sTope As Long, ByVal nTipoCambio As Currency, psEstados As String, ByVal gdFecSis As Date, Optional ByVal psTipoC As String = "", Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim nItem As Long
Dim lnPage As Integer
Dim loImprimeCab As New COMNColoCPig.NCOMColPImpre
        
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


Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

 
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
            sCriterio
     If Trim(psTipoC) <> "" Then
      lsSQL = lsSQL & "and nTipCj in " & psTipoC
     End If
     lsSQL = lsSQL & " ORDER BY substring(C.cCtaCod,9,1), case when SUBSTRING(P.cCtaCod, 9,1)=" & gMonedaNacional & " then" & _
                 " P.nSaldo ELSE P.nSaldo*" & Val(nTipoCambio) & " END DESC, PERS.cPersNombre, c.cCtaCod "
    
        'nTiporecuperacion=1 si es VigenteJudicial(2201) o CanceladoJudicial(2203)
        'nTiporecuperacion=2 si es VigenteCastigado(2201) o CanceladoCastigado(2203)
         
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        Exit Function
    Else
         
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CREDITOS EN COBRANZA JUDICIAL CON SALDO DE CAPITAL MAYOR AL " & Format(gdFecSis, "dd/MM/YYYY"), lsCondiciones, lnPage, 144 - 15, "", 138023)

        cMoneda = lrDataRep!nmoneda
        lsCadImp = lsCadImp & "Moneda: " & IIf(lrDataRep!nmoneda = "1", "Soles", "Dolares") & oFunI.gPrnSaltoLinea

        lnIndice = 0:  lnLineas = 7
         
        With lrDataRep
            'RaiseEvent ShowProgress
            Do While Not lrDataRep.EOF
                
                If cMoneda <> !nmoneda And Len(Trim(cMoneda)) > 0 Then
                    lsCadImp = lsCadImp & String(144 - 15, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & lsNegritaOn & "TOTALES           " & Space(68) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                    sTotal1 = 0
                    sTotal2 = 0
                    nItem = 0
                    cMoneda = !nmoneda
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS EN COBRANZA JUDICIAL CON SALDO DE CAPITAL MAYOR AL " & Format(gdFecSis, "dd/MM/YYYY"), lsCondiciones, lnPage, 144 - 15, "", 138023)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                    lsCadImp = lsCadImp & "Moneda: " & IIf(!nmoneda = "1", "Soles", "Dolares") & oFunI.gPrnSaltoLinea
                End If
                
                nItem = nItem + 1
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                sTotal1 = sTotal1 + lrDataRep!nPrestamo
                sTotal2 = sTotal2 + lrDataRep!nSaldoCapitalSoles
                
                cEstado = Replace(!cEstadoRecuperacion, "Vigente ", "")
                cCliente = !cPersNombre
                 
                lsCadImp = lsCadImp & Space(1) & oFun.ImpreFormat(nItem, 4, 0) & Space(1) & !CCuenta & Space(1) & _
                           Space(1) & cCliente & Space(1) & !dFecVigencia & Space(1) & !dFecIngJud & Space(4) & _
                           Mid(!cMoneda, 1, 1) & Space(3) & oFun.ImpreFormat(!nPrestamo, 10, 2) & Space(1) & _
                           oFun.ImpreFormat(!nSaldoCapitalSoles, 10, 2) & Space(1) & cEstado & oFunI.gPrnSaltoLinea
                 
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS EN COBRANZA JUDICIAL CON SALDO DE CAPITAL MAYOR AL " & Format(gdFecSis, "dd/MM/YYYY"), lsCondiciones, lnPage, 144 - 15, "", 138023)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                'RaiseEvent Progress(.Bookmark, .RecordCount)
                .MoveNext
            Loop
            'RaiseEvent CloseProgress
        End With
        
        lsCadImp = lsCadImp & String(144 - 15, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "TOTALES           " & Space(68) & oFun.ImpreFormat(sTotal1, 10, 2) & _
                    Space(1) & oFun.ImpreFormat(sTotal2, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(144 - 15, "=") & oFunI.gPrnSaltoLinea
    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
    nRepo138023_ListadoSaldoCapital = lsCadBuffer
End Function


Public Function nRepo138025_ResumenMovimientoCJ(ByVal psFecha1 As String, ByVal psFecha2 As String, _
                                                ByVal pnTipoCambio As Currency, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim loImprimeCab As New COMNColoCPig.NCOMColPImpre
        
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


Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

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
   
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
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
        'RaiseEvent ShowProgress
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
            'RaiseEvent Progress(.Bookmark, .RecordCount)
            lrDataRep.MoveNext
        Loop
        'RaiseEvent CloseProgress
        End With
        lrDataRep.Close
        Set lrDataRep = Nothing
         
        lsCadImp = lsCadImp & String(99, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "CONCEPTO                        MONEDA   NACIONAL        MONEDA EXTRANJERA           CONSOLIDADO" & lsNegritaOff & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "                                 NRO       MONTO          NRO       MONTO          NRO       MONTO" & lsNegritaOff & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(99, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "SALDO INICIAL          : " & lsNegritaOff & oFun.ImpreFormat(MatMonN(0, 1), 10, 0) & Space(1) & oFun.ImpreFormat(MatMonN(0, 2), 10, 2) & _
                    Space(1) & oFun.ImpreFormat(MatMonE(0, 1), 10, 0) & Space(1) & oFun.ImpreFormat(MatMonE(0, 2), 10, 2) & _
                    Space(1) & oFun.ImpreFormat(MatMonN(0, 1) + MatMonE(0, 1), 10, 0) & Space(1) & _
                    oFun.ImpreFormat(MatMonN(0, 2) + (MatMonE(0, 2) * pnTipoCambio), 10, 2) & _
                    oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(99, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "INGRESOS NUEVOS        : " & lsNegritaOff & oFun.ImpreFormat(MatMonN(1, 1), 10, 0) & Space(1) & oFun.ImpreFormat(MatMonN(1, 2), 10, 2) & _
                    Space(1) & oFun.ImpreFormat(MatMonE(1, 1), 10, 0) & Space(1) & oFun.ImpreFormat(MatMonE(1, 2), 10, 2) & _
                    Space(1) & oFun.ImpreFormat(MatMonN(1, 1) + MatMonE(1, 1), 10, 0) & Space(1) & _
                    oFun.ImpreFormat(MatMonN(1, 2) + (MatMonE(1, 2) * pnTipoCambio), 10, 2) & _
                    oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(99, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "PAGOS(AMORTIZ.+CANCEL.): " & lsNegritaOff & oFun.ImpreFormat(MatMonN(2, 1), 10, 0) & Space(1) & oFun.ImpreFormat(MatMonN(2, 2), 10, 2) & _
                    Space(1) & oFun.ImpreFormat(MatMonE(2, 1), 10, 0) & Space(1) & oFun.ImpreFormat(MatMonE(2, 2), 10, 2) & _
                    Space(1) & oFun.ImpreFormat(MatMonN(2, 1) + MatMonE(2, 1), 10, 0) & Space(1) & _
                    oFun.ImpreFormat(MatMonN(2, 2) + (MatMonE(2, 2) * pnTipoCambio), 10, 2) & _
                    oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(99, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "SALDO FINAL TOTAL      : " & lsNegritaOff & oFun.ImpreFormat(MatMonN(3, 1), 10, 0) & Space(1) & oFun.ImpreFormat(MatMonN(3, 2), 10, 2) & _
                    Space(1) & oFun.ImpreFormat(MatMonE(3, 1), 10, 0) & Space(1) & oFun.ImpreFormat(MatMonE(3, 2), 10, 2) & _
                    Space(1) & oFun.ImpreFormat(MatMonN(3, 1) + MatMonE(3, 1), 10, 0) & Space(1) & _
                    oFun.ImpreFormat(MatMonN(3, 2) + (MatMonE(3, 2) * pnTipoCambio), 10, 2) & _
                    oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(99, "=") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "SALDO FINAL CON DEMANDA: " & lsNegritaOff & oFun.ImpreFormat(MatMonN(4, 1), 10, 0) & Space(1) & oFun.ImpreFormat(MatMonN(4, 2), 10, 2) & _
                    Space(1) & oFun.ImpreFormat(MatMonE(4, 1), 10, 0) & Space(1) & oFun.ImpreFormat(MatMonE(4, 2), 10, 2) & _
                    Space(1) & oFun.ImpreFormat(MatMonN(4, 1) + MatMonE(4, 1), 10, 0) & Space(1) & _
                    oFun.ImpreFormat(MatMonN(4, 2) + (MatMonE(4, 2) * pnTipoCambio), 10, 2) & _
                    oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(99, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lsNegritaOn & "SALDO FINAL SIN DEMANDA: " & lsNegritaOff & oFun.ImpreFormat(MatMonN(5, 1), 10, 0) & Space(1) & oFun.ImpreFormat(MatMonN(5, 2), 10, 2) & _
                    Space(1) & oFun.ImpreFormat(MatMonE(5, 1), 10, 0) & Space(1) & oFun.ImpreFormat(MatMonE(5, 2), 10, 2) & _
                    Space(1) & oFun.ImpreFormat(MatMonN(5, 1) + MatMonE(5, 1), 10, 0) & Space(1) & _
                    oFun.ImpreFormat(MatMonN(5, 2) + (MatMonE(5, 2) * pnTipoCambio), 10, 2) & _
                    oFunI.gPrnSaltoLinea
         
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
    nRepo138025_ResumenMovimientoCJ = lsCadBuffer
End Function
 
Public Function nRepo150000_ConsultaCreditoenCobranzaJudicial(ByVal psCtaCod As String, Optional ByVal psImpresora As Impresoras = gEPSON, Optional ByVal gdFecSis As Date) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim loImprimeCab As New COMNColoCPig.NCOMColPImpre
        
Dim lrDatCredito As ADODB.Recordset
Dim lrDatTotales As ADODB.Recordset
Dim lrDatGastos As ADODB.Recordset
Dim lrDatTotalesGastos As ADODB.Recordset
Dim lrDatAmortizaciones As ADODB.Recordset
Dim lrDatDatosGenerales As ADODB.Recordset

Dim loValCred As COMDColocRec.DCOMColRecRConsulta
Dim I As Integer
Dim sOperacion As String * 30
Dim sImporte As Currency
Dim sSaldo As Currency

Dim lsNegritaOn As String
Dim lsNegritaOff As String

Dim lsCondiciones As String
Dim ldFecha As Date


Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora




lsNegritaOn = Chr$(27) + Chr$(69)
lsNegritaOff = Chr$(27) + Chr$(70)
 
    'Carga Datos
    Set loValCred = New COMDColocRec.DCOMColRecRConsulta
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
        
    
        
       
    Dim loCalcula As NCOMColRecCalculos
    Dim lnIntCompGenerado As Double, lnIntMoraGenerado As Double
    Dim lnDiasUltTrans As Integer
    Dim oFunF As New COMFunciones.FCOMFechas
    Dim fsFecUltPago As String
'    ldFecha = Format(Now(), "dd/mm/yyyy")
    ldFecha = Format(gdFecSis, "dd/mm/yyyy")
       
    fsFecUltPago = CDate(oFunF.fgFechaHoraGrab(lrDatDatosGenerales!cUltimaActualizacion))
    lnDiasUltTrans = CDate(Format(ldFecha, "dd/mm/yyyy")) - CDate(Format(fsFecUltPago, "dd/mm/yyyy"))

    
    Set loCalcula = New NCOMColRecCalculos
        'INTERES SIMPLE
        lnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, lrDatDatosGenerales!nTasaInt, lrDatTotales!CapitalActual)
        'INTERES SIMPLE
        lnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, lrDatDatosGenerales!nTasaIntMor, lrDatTotales!CapitalActual)
    Set loCalcula = Nothing
       
    lsCondiciones = lsNegritaOn & "Crédito Nro. " & psCtaCod & "               Cliente: " & lrDatCredito!cPersNombre & lsNegritaOff
    
    lnLineas = 0
    lnPage = 1

    'CABECERA
    lsCadImp = lsCadImp & nRepoCabecera("CONSULTA DE CREDITOS EN COBRANZA JUDICIAL", "", lnPage, 125, lsCondiciones, 150000)
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lnIndice = 0:  lnLineas = 8
 
    lsCadImp = lsCadImp & String(125, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & lsNegritaOn & "DATOS DEL CREDITO" & lsNegritaOff & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & String(125, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(43) & "Cond. Credito:  " & lrDatCredito!sCondicion & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Fecha Prestamo : " & lrDatDatosGenerales!FechaPrestamo & Space(15) & " Estad. Credito: " & lrDatCredito!ssEstado & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Monto Prestamo : " & Format(lrDatDatosGenerales!nMontoCol, "0.00") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Linea Credito  : " & lrDatDatosGenerales!CDescripcion & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Analista       : " & lrDatDatosGenerales!Analista & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & String(125, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & lsNegritaOn
    lsCadImp = lsCadImp & "Saldo Capital  : " & oFun.ImpreFormat(lrDatTotales!CapitalActual, 10, 2) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Saldo Int. Com.: " & oFun.ImpreFormat(lrDatTotales!InteresActual + lnIntCompGenerado, 10, 2) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Saldo Int. Mor.: " & oFun.ImpreFormat(lrDatTotales!MoraActual + lnIntMoraGenerado, 10, 2) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Saldo Gastos   : " & oFun.ImpreFormat(lrDatTotales!GastoActual, 10, 2) & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "------------------------------" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Deuda          : " & oFun.ImpreFormat(lrDatTotales!CapitalActual + lrDatTotales!InteresActual + lrDatTotales!MoraActual + lrDatTotales!GastoActual + lnIntCompGenerado + lnIntMoraGenerado, 10, 2) & lsNegritaOff & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & String(125, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & lsNegritaOn & "GASTOS ASIGNADOS" & lsNegritaOff & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & String(125, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "FECHA      DESCR. GASTO                         IMPORTE" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & String(125, "-") & oFunI.gPrnSaltoLinea
    sOperacion = ""
    
    lnLineas = lnLineas + 20
    lnIndice = lnIndice + 20
    
    Set lrDatCredito = Nothing
    Set lrDatDatosGenerales = Nothing
    'RaiseEvent ShowProgress
    Do While Not lrDatGastos.EOF
        sOperacion = lrDatGastos!CDescripcion
        lsCadImp = lsCadImp & lrDatGastos!Fecha & Space(1) & sOperacion & Space(1) & oFun.ImpreFormat(lrDatGastos!Importe, 10, 2) & oFunI.gPrnSaltoLinea
        
        lnLineas = lnLineas + 1
        lnIndice = lnIndice + 1
        
        If lnIndice Mod 300 = 0 Then
            lsCadBuffer = lsCadBuffer & lsCadImp
            lsCadImp = ""
        End If
        
        If lnLineas >= 55 Then
            lsCadImp = lsCadImp & String(122, "-") & "..." & oFunI.gPrnSaltoLinea
            lnPage = lnPage + 1
            lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
            lsCadImp = lsCadImp & nRepoCabecera("CONSULTA DE CREDITOS EN COBRANZA JUDICIAL", "", lnPage, 125, lsCondiciones, 150000)
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
            lnLineas = 8
            lnIndice = lnIndice + 8
            lsCadImp = lsCadImp & lsNegritaOn & "GASTOS ASIGNADOS" & lsNegritaOff & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & String(125, "-") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "FECHA      DESCR. GASTO                         IMPORTE" & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & String(125, "-") & oFunI.gPrnSaltoLinea
            lnLineas = lnLineas + 4
            lnIndice = lnIndice + 4
        End If
        'RaiseEvent Progress(lrDatGastos.Bookmark, lrDatGastos.RecordCount)
        lrDatGastos.MoveNext
    Loop
    Set lrDatGastos = Nothing
    'RaiseEvent CloseProgress
    
    lsCadImp = lsCadImp & String(125, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(11) & "GASTOS ASIGNADOS: " & Space(13) & oFun.ImpreFormat(lrDatTotalesGastos!Importe, 10, 2) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    
    lnLineas = lnLineas + 3
    lnIndice = lnIndice + 3
    
    Set lrDatTotalesGastos = Nothing
    
    lsCadImp = lsCadImp & lsNegritaOn & "MOVIMIENTOS REALIZADOS" & lsNegritaOff & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & String(125, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "FECHA      OPERACION                              MONTO       CAPITAL       INTERES          MORA        GASTOS      SALDO K. " & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & String(125, "-") & oFunI.gPrnSaltoLinea
    
    
    sSaldo = CDbl(lrDatTotales!CapitalPagado + lrDatTotales!CapitalActual)
             
    sOperacion = ""
    'RaiseEvent ShowProgress
    Do While Not lrDatAmortizaciones.EOF
        sImporte = lrDatAmortizaciones!Capital + lrDatAmortizaciones!Interes + lrDatAmortizaciones!Mora + lrDatAmortizaciones!Gastos
        sSaldo = sSaldo - lrDatAmortizaciones!Capital
        sOperacion = lrDatAmortizaciones!Operacion
        lsCadImp = lsCadImp & lrDatAmortizaciones!Fecha & Space(1) & sOperacion & Space(1) & oFun.ImpreFormat(sImporte, 10, 2) & Space(1) & oFun.ImpreFormat(lrDatAmortizaciones!Capital, 10, 2) & Space(1) & oFun.ImpreFormat(lrDatAmortizaciones!Interes, 10, 2) & Space(1) & oFun.ImpreFormat(lrDatAmortizaciones!Mora, 10, 2) & Space(1) & oFun.ImpreFormat(lrDatAmortizaciones!Gastos, 10, 2) & Space(1) & oFun.ImpreFormat(sSaldo, 10, 2) & oFunI.gPrnSaltoLinea
         
        lnLineas = lnLineas + 1
        lnIndice = lnIndice + 1
        
        If lnIndice Mod 300 = 0 Then
            lsCadBuffer = lsCadBuffer & lsCadImp
            lsCadImp = ""
        End If
        
        If lnLineas >= 55 Then
            lsCadImp = lsCadImp & String(122, "-") & "..." & oFunI.gPrnSaltoLinea
            lnPage = lnPage + 1
            lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
            lsCadImp = lsCadImp & nRepoCabecera("CONSULTA DE CREDITOS EN COBRANZA JUDICIAL", "", lnPage, 125, lsCondiciones, 150000)
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
            lnLineas = 8
            lnIndice = lnIndice + 8
            lsCadImp = lsCadImp & lsNegritaOn & "MOVIMIENTOS REALIZADOS" & lsNegritaOff & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & String(125, "-") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "FECHA      OPERACION                              MONTO       CAPITAL       INTERES          MORA        GASTOS      SALDO K. " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & String(125, "-") & oFunI.gPrnSaltoLinea
            
            lnLineas = lnLineas + 4
            lnIndice = lnIndice + 4
        End If
        'RaiseEvent Progress(lrDatAmortizaciones.Bookmark, lrDatAmortizaciones.RecordCount)
        lrDatAmortizaciones.MoveNext
    Loop
    'RaiseEvent CloseProgress
    Set lrDatAmortizaciones = Nothing
    
    lsCadImp = lsCadImp & String(125, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(11) & "TOTAL PAGADO:" & Space(18) & oFun.ImpreFormat(lrDatTotales!CapitalPagado + lrDatTotales!InteresPagado + lrDatTotales!MoraPagado + lrDatTotales!GastoPagado, 10, 2)
    lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
     
    Set lrDatTotales = Nothing
      
    nRepo150000_ConsultaCreditoenCobranzaJudicial = lsCadBuffer
     
End Function


Public Function Repo_138031(ByVal psFiltroEstado As Integer, ByVal pMatAgencias As Variant, _
ByVal pMatAnalistas As Variant, ByVal pnMoneda As Integer, Optional ByVal psTipoC As String = "", Optional ByVal psImpresora As Impresoras = gEPSON, _
Optional ByVal dFechaI As Date, Optional ByVal dFechaF As Date) As String



    Dim Rs As ADODB.Recordset
    Dim rsFiltro As ADODB.Recordset
    Dim ssql As String
    Dim oConec As New COMConecta.DCOMConecta
    Dim odRec  As COMDColocRec.DCOMColRecRConsulta
    
    Dim sNomAgencia As String
    Dim sNomAnalista As String
    Dim sMoneda As String
    Dim sEstado As String
    Dim sCadImp As String
    
    Dim I As Integer
    Dim j As Integer
    
    Dim nCont As Integer
    
    Dim sCodCtaTemp As String
    
    Dim nLinea As Integer
    
    Dim lnPage  As Integer
    
    
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    oFunI.Inicia psImpresora

    
    lnPage = 1
    
    
    
    If pnMoneda = 1 Then
        sMoneda = "SOLES"
    Else
        sMoneda = "DOLARES"
    End If
    
    If psFiltroEstado = "2" Then
        sEstado = "('2202','2206')"
    Else
        sEstado = "('2201','2205')"
    End If
    
    For I = 0 To UBound(pMatAgencias) - 1
        ssql = "Select cAgeDescripcion From Agencias Where cAgeCod='" & pMatAgencias(I) & "'"
        oConec.AbreConexion
        Set rsFiltro = oConec.CargaRecordSet(ssql)
        oConec.CierraConexion
        Set oConec = Nothing
        
        If Not rsFiltro.EOF And Not rsFiltro.BOF Then
            sNomAgencia = rsFiltro!cAgeDescripcion
        End If
        Set rsFiltro = Nothing
        
        For j = 0 To UBound(pMatAnalistas) - 1
             ssql = "Select cPersNombre From Persona Where cPerscod='" & pMatAnalistas(j) & "'"
             oConec.AbreConexion
             Set rsFiltro = oConec.CargaRecordSet(ssql)
             oConec.CierraConexion
             Set oConec = Nothing
             
             If Not rsFiltro.EOF And Not rsFiltro.BOF Then
                sNomAnalista = rsFiltro!cPersNombre
             End If
             Set oConec = Nothing
             Set rsFiltro = Nothing
             
             If j > 0 Then
                sCadImp = sCadImp & oFunI.gPrnSaltoPagina
                sCadImp = sCadImp & nRepoCabecera("Lista de Creditos en Judicial", "Filtrado por Analista y Agencia", lnPage, 145, "", 138031)
             Else
                sCadImp = sCadImp & nRepoCabecera("Lista de Creditos en Judicial", "Filtrado por Analista y Agencia", lnPage, 145, "", 138031)
             End If
             
             Set odRec = New COMDColocRec.DCOMColRecRConsulta
             Set Rs = odRec.ListaCreditosXAnalistaXAgencia(pMatAgencias(I), pMatAnalistas(j), pnMoneda, psTipoC, dFechaI, dFechaF)
             
             If Rs.RecordCount > 0 Then
                 'RaiseEvent ShowProgre
             End If
             
             'Armando estructua de Reporte
             sCadImp = sCadImp & oFunI.gPrnSaltoLinea
             sCadImp = sCadImp & "Agencia: " & csNomAgencia & oFunI.gPrnSaltoLinea
             sCadImp = sCadImp & "Usuario: " & csCodUser & oFunI.gPrnSaltoLinea
             sCadImp = sCadImp & "Agencia Filtro: " & sNomAgencia & oFunI.gPrnSaltoLinea
             sCadImp = sCadImp & "Analista: " & sNomAnalista
             
             sCadImp = sCadImp & oFunI.gPrnSaltoLinea
             sCadImp = sCadImp & oFunI.gPrnSaltoLinea
             
             'Armando las Columnas
             nLinea = 15
             sCadImp = sCadImp & String(120, "-") & oFunI.gPrnSaltoLinea
             sCadImp = sCadImp & String(120, "-") & oFunI.gPrnSaltoLinea
             sCadImp = sCadImp & oFun.ImpreFormat("Credito", 20) & oFun.ImpreFormat("Titular", 30) & oFun.ImpreFormat("Telefono1", 10) & oFun.ImpreFormat("Telefono2", 10) & oFun.ImpreFormat("Direccion", 35) & oFun.ImpreFormat("Direc.Negocio", 30) & oFunI.gPrnSaltoLinea
             
             nCont = 1
             Do Until Rs.EOF
                  If sCodCtaTemp <> Rs!cCtaCod Then
                    If nCont > 1 Then
                        sCadImp = sCadImp & String(120, "-") & oFunI.gPrnSaltoLinea
                        sCadImp = sCadImp & String(120, "-") & oFunI.gPrnSaltoLinea
                        sCadImp = sCadImp & oFun.ImpreFormat("Credito", 20) & oFun.ImpreFormat("Titular", 30) & oFun.ImpreFormat("Telefono1", 10) & oFun.ImpreFormat("Telefono2", 10) & oFun.ImpreFormat("Direccion", 35) & oFun.ImpreFormat("Direc.Negocio", 30) & oFunI.gPrnSaltoLinea
                    End If
                    sCadImp = sCadImp & oFun.ImpreFormat(Rs!cCtaCod, 20) & oFun.ImpreFormat(Rs!cTitular, 30) & oFun.ImpreFormat(Rs!cTitularTelefono1, 10) & oFun.ImpreFormat(Rs!cTitularTelefono2, 10) & oFun.ImpreFormat(Rs!cTitularDomicilio, 35) & oFun.ImpreFormat(Rs!cTitularDirecNegocio, 30) & oFunI.gPrnSaltoLinea
                    sCadImp = sCadImp & oFunI.gPrnSaltoLinea
                  End If
                  'Armando los datos del Credito
                  If sCodCtaTemp <> Rs!cCtaCod Then
                    sCadImp = sCadImp & oFun.ImpreFormat("SK", 10, 2) & oFun.ImpreFormat("Int.Comp", 10, 2) & oFun.ImpreFormat("Int.Mor", 10, 2) & oFun.ImpreFormat("Gasto", 10, 2) & oFun.ImpreFormat("Fec.Pag", 12) & oFun.ImpreFormat("Monto.Pag", 10, 2) & oFunI.gPrnSaltoLinea
                    sCadImp = sCadImp & oFun.ImpreFormat(Rs!nSK, 10, 2) & oFun.ImpreFormat(Rs!nMontoComp, 10, 2) & oFun.ImpreFormat(Rs!nMontoMor, 10, 2) & oFun.ImpreFormat(Rs!nGasto, 10, 2) & oFun.ImpreFormat(Rs!FechaPago, 12) & oFun.ImpreFormat(Rs!MontoPago, 10, 2) & oFunI.gPrnSaltoLinea
                    sCadImp = sCadImp & oFunI.gPrnSaltoLinea
                    'Armando las columnas del Aval
                    sCadImp = sCadImp & oFun.ImpreFormat("Nombre Aval", 20) & oFun.ImpreFormat("Domicilio Aval", 20) & oFun.ImpreFormat("Telefono1", 10) & oFun.ImpreFormat("Telefono2", 10) & oFunI.gPrnSaltoLinea
                  End If
                  sCadImp = sCadImp & oFun.ImpreFormat(Rs!cNombreAval, 20) & oFun.ImpreFormat(Rs!cAvalDomicilio, 20) & oFun.ImpreFormat(Rs!cAvalTelefono1, 10) & oFun.ImpreFormat(Rs!cAvalTelefono2, 10)
                  sCadImp = sCadImp & oFunI.gPrnSaltoLinea
                   
                 sCodCtaTemp = Rs!cCtaCod
                 nCont = nCont + 1
                 'RaiseEvent Progress(Rs.Bookmark, Rs.RecordCount)
                 Rs.MoveNext
             Loop
        Next j
    Next I
    'RaiseEvent CloseProgress
    Repo_138031 = sCadImp
End Function
'*********** Nuevo Reporte 31/03/2005 ******************
Public Function nRepo138034_ListadoActuacionesProcesales(ByVal pdFecActPro As Date, ByRef psmensaje As String) As ADODB.Recordset

Dim lsSQL As String
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lrDataRep As New ADODB.Recordset

    lsSQL = " select CA.cCtaCod,P.cPersCod,P.cPersNombre Cliente,nMontoCol,nSaldoIntComp,nSaldoIntMor,nSaldoGasto,cComenta,dFechaAviso,dFechaVencimiento,P1.cPersNombre Abogado " & _
            " from Producto PR" & _
            " JOIN ProductoPersona PP ON  PR.cCtaCod=PP.cCtaCod AND nPrdPersRelac='20'" & _
            " JOIN Persona P ON P.cPersCod=PP.cPersCod" & _
            " JOIN ProductoPersona PP1 ON  PR.cCtaCod=PP1.cCtaCod AND PP1.nPrdPersRelac='30'" & _
            " JOIN Persona P1 ON P1.cPersCod=PP1.cPersCod" & _
            " JOIN Colocaciones CO ON CO.cCtaCod=PR.cCtaCod" & _
            " JOIN ColocRecup CR ON CR.cCtaCod=CO.cCtaCod" & _
            " JOIN ColocRecupActProcesaLes CA on CA.cCtaCod=CR.cCtaCod" & _
            " WHERE dFechaVencimiento = '" & Format(pdFecActPro, "YYYYMMDD") & "'" & _
            " ORDER BY P.cPersCod "
         
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = New ADODB.Recordset
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If (lrDataRep.BOF And lrDataRep.EOF) Then
       psmensaje = "No Existe datos para el Reporte"
    Else
        Set nRepo138034_ListadoActuacionesProcesales = lrDataRep
    End If
    
End Function

'** DAOR 20070122, Reporte de Actuaciones Procesales
Public Function nRepo138035_ListadoActuacionesProcesales(ByVal pdFecActProDe As Date, pdFecActProHasta As Date, ByVal pnConAbogado As Byte, ByVal psCodAbogado As String, ByRef psmensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lrDataRep As New ADODB.Recordset
Dim whereAbogado As String
    If pnConAbogado = 1 Then
        whereAbogado = " and PP2.cPersCod='" & psCodAbogado & "' "
    Else
        whereAbogado = ""
    End If
    lsSQL = " Select CA.cCtaCod,P.cPersCod,P.cPersNombre Cliente,P2.cPersNombre Abogado, SUBSTRING(CO.cLineaCred, 5, 1) AS Moneda, " & _
            " isnull((select P1.cPersNombre from ProductoPersona PP1 JOIN Persona P1 " & _
            " ON P1.cPersCod=PP1.cPersCod and PR.cCtaCod=PP1.cCtaCod AND PP1.nPrdPersRelac='32'),'') Juzgado, " & _
            " PR.nSaldo,CR.dIngRecup,CE.cNumExp,TD.cConsDescripcion Materia,CA.dFechaVencimiento,CA.cComenta, " & _
            " TA.cConsDescripcion TipoAct,PE.cConsDescripcion as Estado, " & _
            " isnull((select top 1 MD.nMonto from MovColDet MD  where MD.cCtaCod=PR.cCtaCod and MD.cOpeCod='130100' and nPrdConceptoCod=3000),0) nSaldoTrans " & _
            " From Producto PR" & _
            " JOIN ProductoPersona PP ON  PR.cCtaCod=PP.cCtaCod AND nPrdPersRelac='20'" & _
            " JOIN Persona P ON P.cPersCod=PP.cPersCod" & _
            " JOIN ProductoPersona PP2 on PR.cCtaCod=PP2.cCtaCod and PP2.nPrdPersRelac='30' " & _
            " JOIN Persona P2 on PP2.cPersCod=P2.cPersCod " & _
            " JOIN Colocaciones CO ON CO.cCtaCod=PR.cCtaCod" & _
            " JOIN ColocRecup CR ON CR.cCtaCod=CO.cCtaCod " & _
            " JOIN ColocRecupExpediente CE On CE.cCtaCod=CR.cCtaCod " & _
            " JOIN ColocRecupActProcesaLes CA on CA.cCtaCod=CR.cCtaCod" & _
            " JOIN Constante TD on TD.nConsValor=CE.nViaProce and TD.nConsCod=3303 " & _
            " JOIN Constante TA on TA.nConsValor=CA.nTipoAct and TA.nConsCod=3310 " & _
            " JOIN Constante PE on PE.nConsValor=PR.nPrdEstado and PE.nConsCod=3001 " & _
            " WHERE dFechaVencimiento between '" & Format(pdFecActProDe, "YYYYMMDD") & "' and '" & Format(pdFecActProHasta, "YYYYMMDD") & "' " & _
            whereAbogado & _
            " ORDER BY P2.cPersNombre, SUBSTRING(CO.cLineaCred, 5, 1),TD.cConsDescripcion,P.cPersNombre "
         
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = New ADODB.Recordset
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If (lrDataRep.BOF And lrDataRep.EOF) Then
       psmensaje = "No Existe datos para el Reporte"
    Else
        Set nRepo138035_ListadoActuacionesProcesales = lrDataRep
    End If
    
End Function

'**DAOR 20070124, Reporte de créditos castigados
Public Function nRepo138036_CreditosCastigados( _
ByVal pdFecCastDe As Date, _
pdFecCastHasta As Date, _
ByVal nTipoCambio As Currency, _
psCodAge As String, _
ByVal psListaAgencias As String, _
ByVal pMatProd As Variant, _
Optional ByRef psmensaje As String, _
Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim oFun As New COMFunciones.FCOMImpresion
Dim P As String
Dim oConec As New COMConecta.DCOMConecta
Dim whereAgencias As String
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
Dim nTotalD As Double, nTotalS As Double
Dim nTotalCas As Integer, I As Integer, j As Integer
Dim sTempMoneda As String
Dim sCadProd As String
Dim vmone As String

Screen.MousePointer = 11

'**** BRGO - BASILEA II 20100615
    sCadProd = " And CO.cTpoCredCod in ('"
    For I = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    If UBound(pMatProd) = 0 Then
        sCadProd = ""
    End If

    If psListaAgencias = "" Then
        whereAgencias = " and CO.cAgeCodAct = '" & psCodAge & "' "
    Else
        whereAgencias = " and CO.cAgeCodAct in (" & psListaAgencias & " ) "
    End If
    
'''lsSQL = " select  Money=(case SUBSTRING(cLineaCred, 5, 1) when '1' then 'NUEVOS SOLES' else 'DOLARES US' end),*," 'marg ers044-2016
lsSQL = " select  Money=(case SUBSTRING(cLineaCred, 5, 1) when '1' then 'SOLES' else 'DOLARES US' end),*,"
lsSQL = lsSQL & "     TotalDCast=(nSaldo+Interes+Mora+Gastos),"
lsSQL = lsSQL & "     TotalDCastMN=case substring(cctacod,9,1) when '2' then (nSaldo+Interes+Mora+Gastos)*" & nTipoCambio & " else (nSaldo+Interes+Mora+Gastos) end,"
lsSQL = lsSQL & "     TotalACast=(CapiCas+InteCas+MoraCas+GastCas),"
lsSQL = lsSQL & "     TotalACastMN=case substring(cctacod,9,1) when '2' then (CapiCas+InteCas+MoraCas+GastCas)*" & nTipoCambio & " else (CapiCas+InteCas+MoraCas+GastCas) end"
lsSQL = lsSQL & " From"
lsSQL = lsSQL & " (Select CR.cCtaCod,"
lsSQL = lsSQL & "     P.cPersCod,"
lsSQL = lsSQL & "     P.cPersNombre Cliente,"
lsSQL = lsSQL & "     isnull(P2.cPersNombre,'') Abogado,"
lsSQL = lsSQL & "     RH.cUser Analista,"
lsSQL = lsSQL & "     SUBSTRING(CO.cLineaCred, 5, 1) AS Moneda,"
lsSQL = lsSQL & "     CO.cLineaCred,"
lsSQL = lsSQL & "     CO.nMontoCol,"
lsSQL = lsSQL & "     CC.nDiasAtraso,"
lsSQL = lsSQL & "     PR.nSaldo,"
lsSQL = lsSQL & "     Interes=isnull(cr.nSaldoIntComp + cr.nIntCompGen,0),"
lsSQL = lsSQL & "     Mora=isnull(cr.nSaldoIntMor,0),"
lsSQL = lsSQL & "     Gastos=isnull(cr.nSaldoGasto,0),"
lsSQL = lsSQL & "     CapiCas=isnull((select nmonto from movcoldet where nprdconceptocod = 3503 and cctacod=pr.cctacod),0), "
lsSQL = lsSQL & "     InteCas=isnull((select nmonto from movcoldet where nprdconceptocod = 3504 and cctacod=pr.cctacod),0), "
lsSQL = lsSQL & "     MoraCas=isnull((select nmonto from movcoldet where nprdconceptocod = 3505 and cctacod=pr.cctacod),0), "
lsSQL = lsSQL & "     GastCas=isnull((select nmonto from movcoldet where nprdconceptocod = 3506 and cctacod=pr.cctacod),0), "
lsSQL = lsSQL & "     CE.dPrdEstado Fecha,"
lsSQL = lsSQL & "     CtaAnt=isnull(rc.cctacodant,''),"
lsSQL = lsSQL & "     direccion=p.cPersDireccDomicilio,"
lsSQL = lsSQL & "     DocID=isnull(pid.cpersIDNro,''),"
lsSQL = lsSQL & "     AG.cAgeDescripcion NombAgencia,"
lsSQL = lsSQL & "     CR.dIngRecup"
lsSQL = lsSQL & " From Producto PR"
lsSQL = lsSQL & " inner join ProductoPersona PP on  PR.cCtaCod=PP.cCtaCod AND nPrdPersRelac='20'"
lsSQL = lsSQL & " inner join Persona P ON P.cPersCod=PP.cPersCod"
lsSQL = lsSQL & " left join ProductoPersona PP2 on PR.cCtaCod=PP2.cCtaCod and PP2.nPrdPersRelac='30'"
lsSQL = lsSQL & " left join Persona P2 on PP2.cPersCod=P2.cPersCod"
lsSQL = lsSQL & " inner join ProductoPersona PP3 on PR.cCtaCod=PP3.cCtaCod and PP3.nPrdPersRelac='28'"
lsSQL = lsSQL & " inner join RRHH RH on PP3.cPersCod=RH.cPersCod"
lsSQL = lsSQL & " inner join Colocaciones CO on CO.cCtaCod=PR.cCtaCod"
lsSQL = lsSQL & " inner join ColocacCred CC on CO.cCtaCod=CC.cCtaCod"
lsSQL = lsSQL & " inner join ColocacEstado CE on CC.cCtaCod=CE.cCtaCod and CE.nPrdEstado in (2202,2204)"
lsSQL = lsSQL & " inner join ColocRecup CR ON CR.cCtaCod=CO.cCtaCod"
lsSQL = lsSQL & " inner join Agencias AG on CO.cAgeCodAct = AG.cAgeCod"
lsSQL = lsSQL & " left join relcuentas rc on cr.cctacod = rc.cctacodant"
lsSQL = lsSQL & " left join persID pid on pp.cperscod=pid.cperscod and pid.cPersIDTpo='1') lw"
lsSQL = lsSQL & " WHERE Fecha between '" & Format(pdFecCastDe, "YYYYMMDD") & "' and '" & Format(pdFecCastHasta, "YYYYMMDD") & "' "
lsSQL = lsSQL & " " & whereAgencias
lsSQL = lsSQL & " " & sCadProd
lsSQL = lsSQL & " ORDER BY SUBSTRING(cLineaCred, 5, 1),Cliente"

'***** End BRGO
        
        oConec.AbreConexion
        Set lrDataRep = oConec.CargaRecordSet(lsSQL)
        oConec.CierraConexion
        Set oConec = Nothing
        
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia "
        Exit Function
    Else
'---------------------------****************************** inicia excel

    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "REPORTE DE CREDITOS CASTIGADOS DEL " & Format(pdFecCastDe, "dd/MM/YYYY") & " AL " & Format(pdFecCastHasta, "dd/mm/yyyy")
    'ApExcel.cells(2, 2).Font.Size = 12
    ApExcel.RANGE("B2", "R2").mergecells = True
    ApExcel.Cells(2, 2).Font.Bold = True
    ApExcel.Cells(2, 2).horizontalalignment = 3
    
    ApExcel.Cells(5, 2).Formula = "TIPO CAMBIO : " & oFun.ImpreFormat(nTipoCambio, 5, 3)
    ApExcel.Cells(5, 2).Font.Bold = True
    
    ApExcel.Cells(8, 2).Formula = "Nº"
    ApExcel.Cells(8, 3).Formula = "CREDITO"
    ApExcel.Cells(7, 4).Formula = "CUENTA" '**
    ApExcel.Cells(8, 4).Formula = "ANTIGUA"
    ApExcel.Cells(7, 5).Formula = "CODIGO" '**
    ApExcel.Cells(8, 5).Formula = "CLIENTE"
    ApExcel.Cells(8, 6).Formula = "NOMBRE"
    ApExcel.Cells(8, 7).Formula = "DIRECCION"
    ApExcel.Cells(8, 8).Formula = "DOCUMENTO"
    ApExcel.Cells(7, 9).Formula = "FECHA"
    ApExcel.Cells(8, 9).Formula = "CASTIGO"
    ApExcel.Cells(7, 10).Formula = "MONTO"
    ApExcel.Cells(8, 10).Formula = "COLOCADO"
    ApExcel.Cells(7, 11).Formula = "MONTO"
    ApExcel.Cells(8, 11).Formula = "PAGADO"
    ApExcel.Cells(7, 12).Formula = "SALDO"
    ApExcel.Cells(8, 12).Formula = "CAPITAL"
    ApExcel.Cells(8, 13).Formula = "INTERES"
    ApExcel.Cells(8, 14).Formula = "MORA"
    ApExcel.Cells(8, 15).Formula = "GASTOS"
    ApExcel.Cells(8, 16).Formula = "TOTAL"
    ApExcel.Cells(8, 17).Formula = "TOTAL MN"
    ApExcel.Cells(7, 18).Formula = "CAPITAL"
    ApExcel.Cells(8, 18).Formula = "CASTIGADO"
    ApExcel.Cells(7, 19).Formula = "INTERES"
    ApExcel.Cells(8, 19).Formula = "CASTIGADO"
    ApExcel.Cells(7, 20).Formula = "MORA"
    ApExcel.Cells(8, 20).Formula = "CASTIGADO"
    ApExcel.Cells(7, 21).Formula = "GASTO"
    ApExcel.Cells(8, 21).Formula = "CASTIGADO"
    ApExcel.Cells(7, 22).Formula = "TOTAL"
    ApExcel.Cells(8, 22).Formula = "CASTIGADO"
    ApExcel.Cells(7, 23).Formula = "TOTAL"
    ApExcel.Cells(8, 23).Formula = "CAST. MN"
    ApExcel.Cells(8, 24).Formula = "ANALISTA"
    ApExcel.Cells(7, 25).Formula = "DIAS"
    ApExcel.Cells(8, 25).Formula = "ATRASO"
    ApExcel.Cells(8, 26).Formula = "AGENCIA"

    ApExcel.RANGE("B7", "Z8").Interior.Color = RGB(10, 190, 160)
    ApExcel.RANGE("B7", "Z8").Font.Bold = True
    ApExcel.RANGE("B7", "Z8").horizontalalignment = 3

I = 8

    Do While Not lrDataRep.EOF

        I = I + 1
        ApExcel.Cells(I, 2).Formula = "MONEDA : " & lrDataRep!Money
        ApExcel.Cells(I, 2).Font.Bold = True
        vmone = lrDataRep!Moneda
             j = 0
            Do While lrDataRep!Moneda = vmone
                j = j + 1
                I = I + 1
                ApExcel.Cells(I, 2).Formula = j
                ApExcel.Cells(I, 3).Formula = "'" & lrDataRep!cCtaCod
                ApExcel.Cells(I, 4).Formula = "'" & lrDataRep!CtaAnt
                ApExcel.Cells(I, 5).Formula = lrDataRep!cPersCod
                ApExcel.Cells(I, 6).Formula = lrDataRep!CLIENTE
                ApExcel.Cells(I, 7).Formula = lrDataRep!Direccion
                ApExcel.Cells(I, 8).Formula = "'" & lrDataRep!DocID
                ApExcel.Cells(I, 9).Formula = Format(lrDataRep!Fecha, "mm/dd/yyyy")
                ApExcel.Cells(I, 10).Formula = lrDataRep!nMontoCol
                ApExcel.Cells(I, 11).Formula = lrDataRep!nMontoCol - lrDataRep!nSaldo
                ApExcel.Cells(I, 12).Formula = lrDataRep!nSaldo
                ApExcel.Cells(I, 13).Formula = lrDataRep!Interes
                ApExcel.Cells(I, 14).Formula = lrDataRep!Mora
                ApExcel.Cells(I, 15).Formula = lrDataRep!Gastos
                ApExcel.Cells(I, 16).Formula = lrDataRep!TotalDCast
                ApExcel.Cells(I, 17).Formula = lrDataRep!TotalACastMN
                ApExcel.Cells(I, 18).Formula = lrDataRep!CapiCas
                ApExcel.Cells(I, 19).Formula = lrDataRep!InteCas
                ApExcel.Cells(I, 20).Formula = lrDataRep!MoraCas
                ApExcel.Cells(I, 21).Formula = lrDataRep!GastCas
                ApExcel.Cells(I, 22).Formula = lrDataRep!TotalACast
                ApExcel.Cells(I, 23).Formula = lrDataRep!TotalACastMN
                ApExcel.Cells(I, 24).Formula = lrDataRep!Analista
                ApExcel.Cells(I, 25).Formula = lrDataRep!nDiasAtraso
                ApExcel.Cells(I, 26).Formula = lrDataRep!NombAgencia
                                
                ApExcel.RANGE("J" & Trim(Str(I)) & ":" & "W" & Trim(Str(I))).NUMBERFORMAT = "#,##0.00"
'                ApExcel.RANGE("B" & Trim(Str(I)) & ":" & "R" & Trim(Str(I))).borders.linestyle = 1

                
                lrDataRep.MoveNext
                If lrDataRep.EOF Then
                    Exit Do
                End If
                
        Loop
'        If lrDataRep.EOF Then
'            Exit Do
'        End If
    Loop
    
    lrDataRep.Close
    Set lrDataRep = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.RANGE("B2").Select

    ApExcel.Visible = True
    Set ApExcel = Nothing
       
    End If
    'nRepo138036_CreditosCastigados = lsCadBuffer
Set oFunI = Nothing
      
       
'----------------------------****************************** finaliza excel
'        lnLineas = 0
'        lnPage = 1
'        sTempMoneda = "0"
'        nTotalCas = 0
'        '**Configuración de página**********************************************************
'        'lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
'        'lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(66)   'Longitud de página a 22 líneas'
'        lsCadImp = lsCadImp & oFunI.gPrnTamLetra10CPI    'Tamaño 10 cpi
'        'lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
'        'lsCadImp = lsCadImp & Chr$(27) + Chr$(18)  ' cancela condensada
'        lsCadImp = lsCadImp & oFunI.gPrnBoldOFF  ' desactiva negrita
'        '***********************************************************************************
'
        lsCadImp = lsCadImp & nRepoCabecera(" LISTADO DE CRÉDITOS CASTIGADOS ", " DEL " & Format(pdFecCastDe, "dd/mm/yyyy") & " AL " & Format(pdFecCastHasta, "dd/mm/yyyy"), lnPage, 150, "", 138036)
'
'        lnIndice = 0:  lnLineas = 7
'
'        With lrDataRep
'            Do While Not lrDataRep.EOF
'                lnIndice = lnIndice + 1
'                lnLineas = lnLineas + 1
'                If sTempMoneda <> !Moneda And sTempMoneda <> "0" Then
'                    lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea
'                    lsCadImp = lsCadImp & oFun.ImpreFormat(">TOTAL " & IIf(sTempMoneda = "1", "SOLES", "DOLARES"), 20, 0)
'                    lsCadImp = lsCadImp & Space(55) & oFun.ImpreFormat(nTotalD, 12, 2, True) & Space(1)
'                    lsCadImp = lsCadImp & oFun.ImpreFormat(nTotalD - nTotalS, 12, 2, True) & Space(1)
'                    lsCadImp = lsCadImp & oFun.ImpreFormat(nTotalS, 12, 2, True) & oFunI.gPrnSaltoLinea
'                    lsCadImp = lsCadImp & ">Total de Creditos Castigados en  " & IIf(sTempMoneda = "1", "SOLES", "DOLARES") & ": " & nTotalCas & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
'                End If
'                If sTempMoneda <> !Moneda Then
'                    sTempMoneda = !Moneda
'                    lsCadImp = lsCadImp & ">MONEDA : " & IIf(sTempMoneda = "1", "SOLES", "DOLARES") & oFunI.gPrnSaltoLinea
'                    lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea
'                    nTotalD = 0
'                    nTotalS = 0
'                    nTotalCas = 0
'                End If
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!cCtaCod, 18, 0) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!CtaAnt, 18, 0) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!cPersCod, 13, 0) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!Cliente, 30, 0) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!Direccion, 30, 0) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!DocID, 15, 0) & Space(1)
'                lsCadImp = lsCadImp & !Fecha & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!nMontoCol, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!nMontoCol - !nSaldo, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!nSaldo, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!Interes, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!Mora, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!Gastos, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!TotalDCast, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!TotalDCastMN, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!CapiCas, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!InteCas, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!MoraCas, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!GastCas, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!TotalACast, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!TotalACastMN, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!Analista, 4, 0) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!nDiasAtraso, 8, 0) & Space(1)
'                If psListaAgencias <> "" Then
'                    lsCadImp = lsCadImp & Space(1) & oFun.ImpreFormat(!NombAgencia, 12, 0)
'                End If
'                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'                nTotalD = nTotalD + !nMontoCol
'                nTotalS = nTotalS + !nSaldo
'                nTotalCas = nTotalCas + 1
'                If lnIndice Mod 300 = 0 Then
'                    lsCadBuffer = lsCadBuffer & lsCadImp
'                    lsCadImp = ""
'                End If
'
'                If lnLineas >= 55 Then
'                    lnPage = lnPage + 1
'                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
'                    lsCadImp = lsCadImp & nRepoCabecera(" LISTADO DE CRÉDITOS CASTIGADOS ", " DEL " & Format(pdFecCastDe, "dd/mm/yyyy") & " AL " & Format(pdFecCastHasta, "dd/mm/yyyy"), lnPage, 150, "", 138036)
'                    lnLineas = 7
'                End If
'                .MoveNext
'            Loop
'        End With
'        lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & oFun.ImpreFormat(">TOTAL " & IIf(sTempMoneda = "1", "SOLES", "DOLARES"), 20, 0)
'        lsCadImp = lsCadImp & Space(55) & oFun.ImpreFormat(nTotalD, 12, 2, True) & Space(1)
'        lsCadImp = lsCadImp & oFun.ImpreFormat(nTotalD - nTotalS, 12, 2, True) & Space(1)
'        lsCadImp = lsCadImp & oFun.ImpreFormat(nTotalS, 12, 2, True) & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & ">Total de Creditos Castigados en  " & IIf(sTempMoneda = "1", "SOLES", "DOLARES") & ": " & nTotalCas & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
'
'
'        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & "Total de Creditos Castigados: " & oFun.ImpreFormat(lnIndice, 5, 0)
'        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
'
'    End If
'    nRepo138036_CreditosCastigados = lsCadBuffer
'Set oFunI = Nothing

Screen.MousePointer = 0

End Function


Function SaltaLinea(ByRef pnLinea As Integer) As String
    
  pnLinea = pnLinea + 1
  If pnLinea > 40 Then
      pnLinea = 1
      SaltaLinea = oFunI.gPrnSaltoPagina
  Else
     SaltaLinea = oFunI.gPrnSaltoLinea
  End If
End Function

'--PEAC 20070912, Reporte de créditos JUDICALES PARA CASTIGAR

Public Function nRepo138037_RepCredJudToCastigar( _
ByVal psCodAge As String, _
ByVal gdFecSis As Date, _
ByVal pnAtrIni As Integer, _
ByVal pnAtrFin As Integer, _
ByVal nTipoCambio As Currency, _
ByVal psListaAgencias As String) As String

'gsCodAge, gdFecSis, Me.TxtDiaAtrIni.Text, Me.TxtDiaAtrFin.Text, Val(txtTipoCambio.Text)

Dim oDCred As COMDColocRec.DCOMColRecCredito

Dim oFun As New COMFunciones.FCOMImpresion
Dim R As ADODB.Recordset
Dim I As Integer, j As Integer
Dim vmone As String, vage As String

'Dim nTipoCambio As Currency
'nTipoCambio = 3.161

Screen.MousePointer = 11

    Set oDCred = New COMDColocRec.DCOMColRecCredito
    Set R = oDCred.dObtienegRepCredJudToCastigar(psCodAge, gdFecSis, pnAtrIni, pnAtrFin, psListaAgencias)
    Set oDCred = Nothing

    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "REPORTE DE CREDITOS JUDICIALES PARA PASE A CASTIGO AL " & Format(gdFecSis, "dd/MM/YYYY")
    'ApExcel.cells(2, 2).Font.Size = 12
    ApExcel.RANGE("B2", "V2").mergecells = True
    ApExcel.Cells(2, 2).Font.Bold = True
    ApExcel.Cells(2, 2).horizontalalignment = 3
    
    ApExcel.Cells(5, 2).Formula = "TIPO CAMBIO : " & oFun.ImpreFormat(nTipoCambio, 5, 3)
    ApExcel.Cells(5, 2).Font.Bold = True
    
    ApExcel.Cells(8, 2).Formula = "Nº"
    ApExcel.Cells(8, 3).Formula = "CREDITO"
    ApExcel.Cells(8, 4).Formula = "ESTADO"
    ApExcel.Cells(8, 5).Formula = "ANALISTA"
    ApExcel.Cells(8, 6).Formula = "LINEA"
    ApExcel.Cells(7, 7).Formula = "CODIGO" '**
    ApExcel.Cells(8, 7).Formula = "CLIENTE"
    ApExcel.Cells(8, 8).Formula = "CLIENTE"
    ApExcel.Cells(8, 9).Formula = "DOC. ID."
    ApExcel.Cells(8, 10).Formula = "DIRECCION"
    ApExcel.Cells(8, 11).Formula = "COD. S.B.S."
    ApExcel.Cells(7, 12).Formula = "FECHA" '**
    ApExcel.Cells(8, 12).Formula = "VIGENCIA"
    ApExcel.Cells(8, 13).Formula = "ATRASO"
    ApExcel.Cells(7, 14).Formula = "MONTO" '**
    ApExcel.Cells(8, 14).Formula = "DESEMBOLSADO"
    ApExcel.Cells(7, 15).Formula = "MONTO PROV." '**
    ApExcel.Cells(8, 15).Formula = "CONSTITUIDA"
    ApExcel.Cells(7, 16).Formula = "SALDO" '**
    ApExcel.Cells(8, 16).Formula = "CAPITAL"
    ApExcel.Cells(8, 17).Formula = "INTERES"
    ApExcel.Cells(8, 18).Formula = "MORA"
    ApExcel.Cells(8, 19).Formula = "GASTOS"
    ApExcel.Cells(8, 20).Formula = "TOTAL"
    ApExcel.Cells(8, 21).Formula = "TOTAL MN"
    ApExcel.Cells(8, 22).Formula = "CALIFICACION"
    
    ApExcel.RANGE("B7", "V8").Interior.Color = RGB(10, 190, 160)
    ApExcel.RANGE("B7", "V8").Font.Bold = True
    ApExcel.RANGE("B7", "V8").horizontalalignment = 3

I = 8

    Do While Not R.EOF
    I = I + 1
    ApExcel.Cells(I, 2).Formula = "AGENCIA : " & R!NomAge
    ApExcel.Cells(I, 2).Font.Bold = True
    vage = R!NomAge
    
        Do While R!NomAge = vage
        I = I + 1
        ApExcel.Cells(I, 2).Formula = "MONEDA : " & R!cMoneda
        ApExcel.Cells(I, 2).Font.Bold = True
        vmone = R!cMoney
             j = 0
            Do While R!cMoney = vmone
                j = j + 1
                I = I + 1
                ApExcel.Cells(I, 2).Formula = j
                ApExcel.Cells(I, 3).Formula = "'" & R!Credito
                ApExcel.Cells(I, 4).Formula = R!Estado
                ApExcel.Cells(I, 5).Formula = R!Analista
                ApExcel.Cells(I, 6).Formula = R!Linea
                ApExcel.Cells(I, 7).Formula = "'" & R!CodCliente
                ApExcel.Cells(I, 8).Formula = R!CLIENTE
                ApExcel.Cells(I, 9).Formula = R!Dni
                ApExcel.Cells(I, 10).Formula = R!Direccion
                ApExcel.Cells(I, 11).Formula = R!CodSbs
                ApExcel.Cells(I, 12).Formula = Format(R!FechaVige, "mm/dd/yyyy")
                ApExcel.Cells(I, 13).Formula = R!Atraso
                ApExcel.Cells(I, 14).Formula = R!MontoDesem
                ApExcel.Cells(I, 15).Formula = R!ProviConst
                ApExcel.Cells(I, 16).Formula = R!SaldoCap
                ApExcel.Cells(I, 17).Formula = R!IntCom + ((((1 + (R!TasaIntComp / 100)) ^ (R!nDiasTransUltPgo / 30)) - 1) * R!SaldoCap)
                ApExcel.Cells(I, 18).Formula = R!IntMor + ((((1 + (R!TasaIntMora / 100)) ^ (R!nDiasTransUltPgo / 30)) - 1) * R!SaldoCap)
                ApExcel.Cells(I, 19).Formula = R!Gastos
                ApExcel.Cells(I, 20).Formula = "=+RC[-4]+RC[-3]+RC[-2]+RC[-1]"
                ApExcel.Cells(I, 21).Formula = IIf(R!cMoney = "2", "=+RC[-1]*" & nTipoCambio, 0)
                ApExcel.Cells(I, 22).Formula = "'" & R!Calificacion
                
                ApExcel.RANGE("N" & Trim(Str(I)) & ":" & "U" & Trim(Str(I))).NUMBERFORMAT = "#,##0.00"
                ApExcel.RANGE("B" & Trim(Str(I)) & ":" & "V" & Trim(Str(I))).borders.linestyle = 1
                
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                               
            Loop
            
            I = I + 1
            ApExcel.Cells(I, 14).Formula = "=SUM(R[-" & Trim(Str(Int(j))) & "]C:R[-1]C)"
            ApExcel.Cells(I, 15).Formula = "=SUM(R[-" & Trim(Str(Int(j))) & "]C:R[-1]C)"
            ApExcel.Cells(I, 16).Formula = "=SUM(R[-" & Trim(Str(Int(j))) & "]C:R[-1]C)"
            ApExcel.Cells(I, 17).Formula = "=SUM(R[-" & Trim(Str(Int(j))) & "]C:R[-1]C)"
            ApExcel.Cells(I, 18).Formula = "=SUM(R[-" & Trim(Str(Int(j))) & "]C:R[-1]C)"
            ApExcel.Cells(I, 19).Formula = "=SUM(R[-" & Trim(Str(Int(j))) & "]C:R[-1]C)"
            ApExcel.Cells(I, 20).Formula = "=SUM(R[-" & Trim(Str(Int(j))) & "]C:R[-1]C)"
            ApExcel.Cells(I, 21).Formula = "=SUM(R[-" & Trim(Str(Int(j))) & "]C:R[-1]C)"
            I = I + 1
            
            If R.EOF Then
                Exit Do
            End If
            
        Loop
    Loop
    
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.RANGE("B2").Select

    ApExcel.Visible = True
    Set ApExcel = Nothing

    Set oFunI = Nothing

Screen.MousePointer = 0

End Function

''--PEAC 20070912, Reporte de cobranza de gestores

Public Function nRepo138038_RepCobranzaGestores( _
ByVal psCodAge As String, _
ByVal pnAtrIni As Integer, _
ByVal pnAtrFin As Integer, _
ByVal pdDel As Date, _
ByVal pdAl As Date, _
ByVal psNomAge As String, _
ByVal pdFecSis As Date, _
ByVal psCodUser As String, _
ByVal psListaAgencias As String, _
Optional ByRef psmensaje As String, _
Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim oDCred As COMDColocRec.DCOMColRecCredito
Dim R As ADODB.Recordset
Dim I As Integer, j As Integer
Dim vmone As String, vage As String, vana As String
Dim nTipoCambio As Currency
Dim lsCadImp As String
Dim lnPage As Integer
Dim lnLineas As Integer
Dim lnTotPago As Double, lnTotCapi As Double, lnTotInte As Double
Dim lnTotMora As Double, lnTotGtos As Double

Dim lnTotGenPago As Double, lnTotGenCapi As Double, lnTotGenInte As Double
Dim lnTotGenMora As Double, lnTotGenGtos As Double

Dim oFunI As New COMFunciones.FCOMVarImpresion

Screen.MousePointer = 11

oFunI.Inicia psImpresora

    Set oDCred = New COMDColocRec.DCOMColRecCredito
    Set R = oDCred.dObtienegRepCobranzaGestores(psCodAge, psListaAgencias, pnAtrIni, pnAtrFin, pdDel, pdAl)
    Set oDCred = Nothing

If R.RecordCount = 0 Then
    'Dim lpmes As String
    'lpmes = MsgBox("No existen Datos para este Reporte.", vbOKOnly + 48, "Atención")
    MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
    Exit Function
End If



lnPage = 1

    lsCadImp = ""
        lsCadImp = lsCadImp & nRepoCabecera(" INGRESOS DE CAJA ", " DEL " & Format(pdDel, "dd/mm/yyyy") & " AL " & Format(pdAl, "dd/mm/yyyy"), lnPage, 165, "", 138038)

    lnLineas = 7

    Do While Not R.EOF
        lsCadImp = lsCadImp & "AGENCIA  : " & R!Agencia & oFunI.gPrnSaltoLinea
        vage = R!Agencia

        Do While R!Agencia = vage
            lsCadImp = lsCadImp & "      MONEDA : " & R!Moneey & oFunI.gPrnSaltoLinea
            vmone = R!Moneda

            lnTotGenPago = 0
            lnTotGenCapi = 0
            lnTotGenInte = 0
            lnTotGenMora = 0
            lnTotGenGtos = 0

            Do While R!Moneda = vmone
                lsCadImp = lsCadImp & "           ANALISTA : " & R!cUser & oFunI.gPrnSaltoLinea
                vana = R!cUser

                    j = 0
                    lnTotPago = 0
                    lnTotCapi = 0
                    lnTotInte = 0
                    lnTotMora = 0
                    lnTotGtos = 0

                    Do While R!cUser = vana
                        j = j + 1
                        lnLineas = lnLineas + 1

                        lsCadImp = lsCadImp & oFun.ImpreFormat(j, 4, 0)
                        lsCadImp = lsCadImp & oFun.ImpreFormat(R!cCtaCod, 18)
                        lsCadImp = lsCadImp & oFun.ImpreFormat(R!nCuota, 6, 0)
                        lsCadImp = lsCadImp & oFun.ImpreFormat(R!cPersNombre, 30)
                        lsCadImp = lsCadImp & oFun.ImpreFormat(R!Capital + R!Interes + R!Mora + R!Gastos, 12, 2, True)
                        lsCadImp = lsCadImp & oFun.ImpreFormat(R!Capital, 12, 2, True)
                        lsCadImp = lsCadImp & oFun.ImpreFormat(R!Interes, 12, 2, True)
                        lsCadImp = lsCadImp & oFun.ImpreFormat(R!Mora, 12, 2, True)
                        lsCadImp = lsCadImp & oFun.ImpreFormat(R!Gastos, 12, 2, True) & Space(2)
                        lsCadImp = lsCadImp & oFun.ImpreFormat(R!cUser, 4, 0)
                        lsCadImp = lsCadImp & oFun.ImpreFormat(Format(R!dPago, "dd/mm/yyyy"), 10)
                        lsCadImp = lsCadImp & oFun.ImpreFormat(R!Atraso, 6, 0)
                        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                        lnTotPago = lnTotPago + R!Capital + R!Interes + R!Mora + R!Gastos
                        lnTotCapi = lnTotCapi + R!Capital
                        lnTotInte = lnTotInte + R!Interes
                        lnTotMora = lnTotMora + R!Mora
                        lnTotGtos = lnTotGtos + R!Gastos

                        If lnLineas >= 55 Then
                            lnPage = lnPage + 1
                            lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                            lsCadImp = lsCadImp & nRepoCabecera(" INGRESOS DE CAJA ", " DEL " & Format(pdDel, "dd/mm/yyyy") & " AL " & Format(pdAl, "dd/mm/yyyy"), lnPage, 165, "", 138038)
                            lnLineas = 7
                        End If

                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                    Loop

                    lsCadImp = lsCadImp & String(165, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & Space(35) & "TOT. ANALISTA :" & Space(12)
                    lsCadImp = lsCadImp & oFun.ImpreFormat(lnTotPago, 12, 2, True)
                    lsCadImp = lsCadImp & oFun.ImpreFormat(lnTotCapi, 12, 2, True)
                    lsCadImp = lsCadImp & oFun.ImpreFormat(lnTotInte, 12, 2, True)
                    lsCadImp = lsCadImp & oFun.ImpreFormat(lnTotMora, 12, 2, True)
                    lsCadImp = lsCadImp & oFun.ImpreFormat(lnTotGtos, 12, 2, True)
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea

                    lnTotGenPago = lnTotGenPago + lnTotPago
                    lnTotGenCapi = lnTotGenCapi + lnTotCapi
                    lnTotGenInte = lnTotGenInte + lnTotInte
                    lnTotGenMora = lnTotGenMora + lnTotMora
                    lnTotGenGtos = lnTotGenGtos + lnTotGtos

                    If R.EOF Then
                        Exit Do
                    End If
            Loop
            lsCadImp = lsCadImp & String(165, "-") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Space(35) & "TOT. MONEDA :" & Space(14)
            lsCadImp = lsCadImp & oFun.ImpreFormat(lnTotGenPago, 12, 2, True)
            lsCadImp = lsCadImp & oFun.ImpreFormat(lnTotGenCapi, 12, 2, True)
            lsCadImp = lsCadImp & oFun.ImpreFormat(lnTotGenInte, 12, 2, True)
            lsCadImp = lsCadImp & oFun.ImpreFormat(lnTotGenMora, 12, 2, True)
            lsCadImp = lsCadImp & oFun.ImpreFormat(lnTotGenGtos, 12, 2, True)
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
            If R.EOF Then
                Exit Do
            End If
        Loop

    Loop

    R.Close
    Set R = Nothing
    nRepo138038_RepCobranzaGestores = lsCadImp
    Set oFunI = Nothing

    Screen.MousePointer = 0
End Function

'--PEAC 20071010, Reporte de gastos de cred jud y cas

Public Function nRepo138038_RepGastosJudCas( _
ByVal psCodAge As String, _
ByVal pEstado As String, _
ByVal pMoneda As String, _
ByVal pdDel As Date, _
ByVal pdAl As Date, _
ByVal psNomAge As String, _
ByVal pdFecSis As Date, _
ByVal psCodUser As String, _
ByVal psListaAgencias As String, _
ByVal pMatProd As Variant, _
Optional ByRef psmensaje As String, _
Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim oDCred As COMDColocRec.DCOMColRecCredito
Dim R As ADODB.Recordset
Dim I As Integer, j As Integer
Dim vmone As String, vage As String, vana As String, vEstado As String, vmoneda As String, vproducto As String
Dim vCuenta As String, vNombre As String
Dim nTipoCambio As Currency
Dim lsCadImp As String
Dim lnPage As Integer
Dim lnLineas As Integer
Dim lnTotPago As Double, lnTotCapi As Double, lnTotInte As Double
Dim lnTotMora As Double, lnTotGtos As Double

Dim lnMonto As Double, lnMontoPag As Double, lnSaldo As Double
Dim lnMonto1 As Double, lnMontoPag1 As Double, lnSaldo1 As Double
Dim lnMonto2 As Double, lnMontoPag2 As Double, lnSaldo2 As Double

Dim oFunI As New COMFunciones.FCOMVarImpresion

oFunI.Inicia psImpresora

Set oDCred = New COMDColocRec.DCOMColRecCredito
Set R = oDCred.dObtienegRepGastosJudCas(psCodAge, pEstado, pMoneda, psListaAgencias, pdDel, pdAl, pMatProd)
Set oDCred = Nothing
    
If R.RecordCount = 0 Then
    MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
    Exit Function
End If
        
lnPage = 1

    lsCadImp = ""
          
    lsCadImp = lsCadImp & nRepoCabecera(" REPORTE DE GASTOS DE CREDITOS (" & R!Estado & ")", " DEL " & Format(pdDel, "dd/mm/yyyy") & " AL " & Format(pdAl, "dd/mm/yyyy"), lnPage, 165, "", 138038)
      
    lsCadImp = lsCadImp & "MONEDA : " & R!Moneda & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "ESTADO : " & R!Estado & oFunI.gPrnSaltoLinea
    
    lnLineas = 7
    
    Do While Not R.EOF
        lsCadImp = lsCadImp & "AGENCIA  : " & R!Agencia & oFunI.gPrnSaltoLinea
        vage = R!Agencia
        
        lnMonto2 = 0
        lnMontoPag2 = 0
        lnSaldo2 = 0
                
            Do While R!Agencia = vage
                lsCadImp = lsCadImp & "   PRODUCTO : " & R!Producto & oFunI.gPrnSaltoLinea
                vproducto = R!Producto
                                               
                    Do While R!Producto = vproducto
                    lsCadImp = lsCadImp & "CREDITO ANTIGUO: " & R!cCtaCodAnt & "CREDITO NUEVO: " & R!cCtaCod & "NOMBRES: " & R!cPersNombre & oFunI.gPrnSaltoLinea
                    vCuenta = R!cCtaCod
                        
                        lnMonto = 0
                        lnMontoPag = 0
                        lnSaldo = 0
                                                
                        j = 0
                        
                            Do While R!cCtaCod = vCuenta
                                                    
                                j = j + 1
                                lnLineas = lnLineas + 1
                        
                                lsCadImp = lsCadImp & Space(5) & oFun.ImpreFormat(j, 4, 0)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!cMotivoGasto, 20)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!nMonto, 12, 2, True)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(Format(R!dAsigna, "dd/mm/yyyy"), 10)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!nMontoPagado, 12, 2, True)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Pago, 8)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Saldo, 12, 2, True)
                                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                                                        
                                 lnMonto = lnMonto + R!nMonto
                                 lnMontoPag = lnMontoPag + R!nMontoPagado
                                 lnSaldo = lnSaldo + R!Saldo
                                                        
                                If lnLineas >= 55 Then
                                    lnPage = lnPage + 1
                                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                                    lsCadImp = lsCadImp & nRepoCabecera(" REPORTE DE GASTOS DE CREDITOS (" & R!Estado & ")", " DEL " & Format(pdDel, "dd/mm/yyyy") & " AL " & Format(pdAl, "dd/mm/yyyy"), lnPage, 165, "", 138038)
                                    lnLineas = 7
                                End If
                        
                                R.MoveNext
                                If R.EOF Then
                                    Exit Do
                                End If
                            Loop

                            lsCadImp = lsCadImp & String(165, "-") & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & Space(15) & "TOT. CUENTA :" & Space(14)
                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnMonto, 12, 2, True)
                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnMontoPag, 12, 2, True)
                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnSaldo, 12, 2, True)
                            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                                   
                            lnMonto1 = lnMonto1 + lnMonto
                            lnMontoPag1 = lnMontoPag1 + lnMontoPag
                            lnSaldo1 = lnSaldo1 + lnSaldo

                    If R.EOF Then
                        Exit Do
                    End If
                Loop
            
                lsCadImp = lsCadImp & String(165, "-") & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & Space(15) & "TOT. PRODUCTO :" & Space(14)
                lsCadImp = lsCadImp & oFun.ImpreFormat(lnMonto1, 12, 2, True)
                lsCadImp = lsCadImp & oFun.ImpreFormat(lnMontoPag1, 12, 2, True)
                lsCadImp = lsCadImp & oFun.ImpreFormat(lnSaldo1, 12, 2, True)
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                       
                lnMonto2 = lnMonto2 + lnMonto1
                lnMontoPag2 = lnMontoPag2 + lnMontoPag1
                lnSaldo2 = lnSaldo2 + lnSaldo1
                       
            If R.EOF Then
                Exit Do
            End If
        Loop
        
        lsCadImp = lsCadImp & String(165, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(15) & "TOT. AGENCIA  :" & Space(14)
        lsCadImp = lsCadImp & oFun.ImpreFormat(lnMonto2, 12, 2, True)
        lsCadImp = lsCadImp & oFun.ImpreFormat(lnMontoPag2, 12, 2, True)
        lsCadImp = lsCadImp & oFun.ImpreFormat(lnSaldo2, 12, 2, True) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(165, "-") & oFunI.gPrnSaltoLinea
        
    Loop
    
    R.Close
    Set R = Nothing
    nRepo138038_RepGastosJudCas = lsCadImp
    Set oFunI = Nothing
    
End Function

'--PEAC 20071011, Reporte de creditos con saldos condonados

Public Function nRepo138039_RepCredSaldosCondo( _
ByVal psCodAge As String, _
ByVal pdDel As Date, _
ByVal pdAl As Date, _
ByVal psNomAge As String, _
ByVal pdFecSis As Date, _
ByVal psCodUser As String, _
ByVal psListaAgencias As String, _
Optional ByRef psmensaje As String, _
Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim oDCred As COMDColocRec.DCOMColRecCredito
Dim R As ADODB.Recordset
Dim I As Integer, j As Integer

Dim vmone As String, vage As String
Dim vmoneda As String

Dim vCuenta As String
Dim nTipoCambio As Currency
Dim lsCadImp As String
Dim lnPage As Integer
Dim lnLineas As Integer

Dim lnCapital As Double, lnInteres As Double, lnMora As Double, lnGastos As Double, lnTotal As Double

Dim oFunI As New COMFunciones.FCOMVarImpresion

oFunI.Inicia psImpresora

Set oDCred = New COMDColocRec.DCOMColRecCredito
Set R = oDCred.dObtienegRepCredSaldosCondo(psCodAge, psListaAgencias, pdDel, pdAl)
Set oDCred = Nothing
    
If R.RecordCount = 0 Then
    MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
    Exit Function
End If
    
lnPage = 1

    lsCadImp = ""
    
    lsCadImp = lsCadImp & nRepoCabecera(" REPORTE DE CREDITOS CON SALDOS CONDONADOS ", " DEL " & Format(pdDel, "dd/mm/yyyy") & " AL " & Format(pdAl, "dd/mm/yyyy"), lnPage, 175, "", 138039)
        
    lnLineas = 7
    
    Do While Not R.EOF
        lsCadImp = lsCadImp & "AGENCIA  : " & R!Agencia & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        vage = R!CodAge
                        
            Do While R!CodAge = vage
                lsCadImp = lsCadImp & "MONEDA : " & R!Moneda & oFunI.gPrnSaltoLinea
                vmoneda = R!Moneda
                                                
                        lnCapital = 0
                        lnInteres = 0
                        lnMora = 0
                        lnGastos = 0
                        lnTotal = 0

                        j = 0
                        
                        Do While R!Moneda = vmoneda
                                                   
                                j = j + 1
                                lnLineas = lnLineas + 1
                        
                                lsCadImp = lsCadImp & oFun.ImpreFormat(j, 3, 0)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!CtaAnt, 18)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!CtaAct, 18)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!cPersNombre, 30)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Analista, 4)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Capital, 10, 2, True)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Interes, 10, 2, True)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Mora, 10, 2, True)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Gastos, 10, 2, True)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Total, 10, 2, True)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(Format(R!Fecha, "dd/mm/yyyy"), 10)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Estado, 20)
                                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                                
                                lnCapital = lnCapital + R!Capital
                                lnInteres = lnInteres + R!Interes
                                lnMora = lnMora + R!Mora
                                lnGastos = lnGastos + R!Gastos
                                lnTotal = lnTotal + R!Total
                        
                                If lnLineas >= 55 Then
                                    lnPage = lnPage + 1
                                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                                    lsCadImp = lsCadImp & nRepoCabecera(" REPORTE DE CREDITOS CON SALDOS CONDONADOS ", " DEL " & Format(pdDel, "dd/mm/yyyy") & " AL " & Format(pdAl, "dd/mm/yyyy"), lnPage, 175, "", 138039)
                                    lnLineas = 7
                                End If
                        
                                R.MoveNext
                                If R.EOF Then
                                    Exit Do
                                End If
                            Loop
                                                                                
                            lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & Space(27) & "TOT. MONEDA   :" & Space(39)
                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnCapital, 10, 2, True)
                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnInteres, 10, 2, True)
                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnMora, 10, 2, True)
                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnGastos, 10, 2, True)
                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnTotal, 10, 2, True) & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                                                        
                    If R.EOF Then
                        Exit Do
                    End If
                Loop
    Loop
    
    R.Close
    Set R = Nothing
    nRepo138039_RepCredSaldosCondo = lsCadImp
    Set oFunI = Nothing
    
End Function

'--PEAC 20071023, Reporte de cartera de ceditos judiciales y castigados

Public Function nRepo138015_RepCarteraJudCas( _
ByVal psCodAge As String, _
ByVal pEstado As String, _
ByVal pMoneda As String, _
ByVal pdDel As Date, _
ByVal pdAl As Date, _
ByVal psNomAge As String, _
ByVal pdFecSis As Date, _
ByVal psCodUser As String, _
ByVal psListaAgencias As String, _
ByVal pMatProd As Variant, _
Optional ByRef psmensaje As String, _
Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim oDCred As COMDColocRec.DCOMColRecCredito
Dim R As ADODB.Recordset
Dim I As Integer, j As Integer
Dim vmone As String, vage As String, vana As String, vEstado As String, vmoneda As String, vproducto As String
Dim vCuenta As String, vNombre As String
Dim nTipoCambio As Currency
Dim lsCadImp As String
Dim lnPage As Integer
Dim lnLineas As Integer

Dim lnCapi2 As Double, lnInte2 As Double, lnMora2 As Double, lnGast2 As Double, lnTota2 As Double
Dim lnCapi1 As Double, lnInte1 As Double, lnMora1 As Double, lnGast1 As Double, lnTota1 As Double
Dim lnCapi0 As Double, lnInte0 As Double, lnMora0 As Double, lnGast0 As Double, lnTota0 As Double

Dim oFunI As New COMFunciones.FCOMVarImpresion

oFunI.Inicia psImpresora

Set oDCred = New COMDColocRec.DCOMColRecCredito
Set R = oDCred.dObtienegRepCarteraJudCas(psCodAge, pEstado, pMoneda, psListaAgencias, pdDel, pdAl, pMatProd)
Set oDCred = Nothing
    
If R.RecordCount = 0 Then
    MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
    Exit Function
End If
        
lnPage = 1

    lsCadImp = ""
    
    lsCadImp = lsCadImp & nRepoCabecera(" REPORTE DE CARTERA DE CREDITOS ", " DEL " & Format(pdDel, "dd/mm/yyyy") & " AL " & Format(pdAl, "dd/mm/yyyy"), lnPage, 175, "", 138015)
      
    lsCadImp = lsCadImp & "MONEDA : " & R!Moneda & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    
    lnLineas = 9
    
    Do While Not R.EOF
        lnLineas = lnLineas + 1
        lsCadImp = lsCadImp & "AGENCIA: " & R!Agencia & oFunI.gPrnSaltoLinea
        vage = R!Agencia
        
        lnCapi2 = 0: lnInte2 = 0: lnMora2 = 0: lnGast2 = 0: lnTota2 = 0
                
            Do While R!Agencia = vage
                lnLineas = lnLineas + 1
                lsCadImp = lsCadImp & "  PRODUCTO: " & R!Producto & oFunI.gPrnSaltoLinea
                vproducto = R!Producto
                                               
                lnCapi1 = 0: lnInte1 = 0: lnMora1 = 0: lnGast1 = 0: lnTota1 = 0
                                               
                    Do While R!Producto = vproducto
                    lnLineas = lnLineas + 1
                    lsCadImp = lsCadImp & "    ESTADO: " & R!Estado & oFunI.gPrnSaltoLinea
                    vEstado = R!Estado
                        
                        lnCapi0 = 0: lnInte0 = 0: lnMora0 = 0: lnGast0 = 0: lnTota0 = 0
                                                
                        j = 0
                        
                            Do While R!Estado = vEstado

                                j = j + 1
                                lnLineas = lnLineas + 1
                                lsCadImp = lsCadImp & Space(5) & oFun.ImpreFormat(j, 4, 0)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!cCtaCod, 18)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!cPersNombre, 30)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!cPersDireccDomicilio, 30)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!nDiasAtraso, 6, 0)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Capital, 12, 2, True)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Interes, 12, 2, True)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Mora, 12, 2, True)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Gastos, 12, 2, True)
                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Total, 12, 2, True) & oFunI.gPrnSaltoLinea
                                'lsCadImp = lsCadImp & String(165, "-") & oFunI.gPrnSaltoLinea
                                                        
                                'lsCadImp = lsCadImp & oFun.ImpreFormat(Format(R!dAsigna, "dd/mm/yyyy"), 10)
                                
                                 lnCapi0 = lnCapi0 + R!Capital
                                 lnInte0 = lnInte0 + R!Interes
                                 lnMora0 = lnMora0 + R!Mora
                                 lnGast0 = lnGast0 + R!Gastos
                                 lnTota0 = lnTota0 + R!Total
                                                                                         
                                If lnLineas >= 55 Then
                                    lnPage = lnPage + 1
                                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                                    lsCadImp = lsCadImp & nRepoCabecera(" REPORTE DE CARTERA DE CREDITOS ", " DEL " & Format(pdDel, "dd/mm/yyyy") & " AL " & Format(pdAl, "dd/mm/yyyy"), lnPage, 175, "", 138015)
                                    lnLineas = 9
                                End If
                        
                                R.MoveNext
                                If R.EOF Then
                                    Exit Do
                                End If
                            Loop
                            lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & Space(15) & "TOT. ESTADO :" & Space(71)
                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnCapi0, 12, 2, True)
                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnInte0, 12, 2, True)
                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnMora0, 12, 2, True)
                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnGast0, 12, 2, True)
                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnTota0, 12, 2, True) & oFunI.gPrnSaltoLinea
                            lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                            
                            lnCapi1 = lnCapi1 + lnCapi0
                            lnInte1 = lnInte1 + lnInte0
                            lnMora1 = lnMora1 + lnMora0
                            lnGast1 = lnGast1 + lnGast0
                            lnTota1 = lnTota1 + lnTota0
                            
                                If lnLineas >= 55 Then
                                    lnPage = lnPage + 1
                                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                                    lsCadImp = lsCadImp & nRepoCabecera(" REPORTE DE CARTERA DE CREDITOS ", " DEL " & Format(pdDel, "dd/mm/yyyy") & " AL " & Format(pdAl, "dd/mm/yyyy"), lnPage, 175, "", 138015)
                                    lnLineas = 9
                                End If
                            
                    If R.EOF Then
                        Exit Do
                    End If
                Loop
            
                lsCadImp = lsCadImp & Space(15) & "TOT. PRODUCTO :" & Space(69)
                lsCadImp = lsCadImp & oFun.ImpreFormat(lnCapi1, 12, 2, True)
                lsCadImp = lsCadImp & oFun.ImpreFormat(lnInte1, 12, 2, True)
                lsCadImp = lsCadImp & oFun.ImpreFormat(lnMora1, 12, 2, True)
                lsCadImp = lsCadImp & oFun.ImpreFormat(lnGast1, 12, 2, True)
                lsCadImp = lsCadImp & oFun.ImpreFormat(lnTota1, 12, 2, True) & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
                
                'lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                       
                lnCapi2 = lnCapi2 + lnCapi1
                lnInte2 = lnInte2 + lnInte1
                lnMora2 = lnMora2 + lnMora1
                lnGast2 = lnGast2 + lnGast1
                lnTota2 = lnTota2 + lnTota1
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera(" REPORTE DE CARTERA DE CREDITOS ", " DEL " & Format(pdDel, "dd/mm/yyyy") & " AL " & Format(pdAl, "dd/mm/yyyy"), lnPage, 175, "", 138015)
                    lnLineas = 9
                End If
                       
            If R.EOF Then
                Exit Do
            End If
        Loop
        
        lsCadImp = lsCadImp & Space(15) & "TOT. AGENCIA  :" & Space(69)
        lsCadImp = lsCadImp & oFun.ImpreFormat(lnCapi2, 12, 2, True)
        lsCadImp = lsCadImp & oFun.ImpreFormat(lnInte2, 12, 2, True)
        lsCadImp = lsCadImp & oFun.ImpreFormat(lnMora2, 12, 2, True)
        lsCadImp = lsCadImp & oFun.ImpreFormat(lnGast2, 12, 2, True)
        lsCadImp = lsCadImp & oFun.ImpreFormat(lnTota2, 12, 2, True) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
        
    Loop
    
    R.Close
    Set R = Nothing
    nRepo138015_RepCarteraJudCas = lsCadImp
    Set oFunI = Nothing
    
End Function
