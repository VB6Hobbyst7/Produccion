VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMColPRepo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim csNomCMAC As String
Dim csNomAgencia As String
Dim csCodUser As String
Dim csFechaSis As String

'Dim xlAplicacion As Excel.Application
'Dim xlLibro As Excel.Workbook
'Dim xlHoja1 As Excel.Worksheet
'Dim lsArchivoN As String, lbLibroOpen As Boolean

Public Sub inicio(ByVal psNomCmac As String, ByVal psNomAgencia As String, _
            ByVal psCodUser As String, ByVal psFechaSis As String)
    csNomCMAC = psNomCmac
    csNomAgencia = psNomAgencia
    csCodUser = psCodUser
    csFechaSis = psFechaSis
End Sub
        
Public Function nRepoCabecera(ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        ByVal pnCodReporte As Long, Optional ByVal psImpresora As Impresoras = gEPSON) As String
        
Dim lsCadImp As String
Dim loImprimeCab As NCOMColPImpre
    
Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


    Set loImprimeCab = New NCOMColPImpre
        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    Set loImprimeCab = Nothing
    If pnCodReporte <> 128000 Then '** Juez 20120529
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
    End If
    Select Case pnCodReporte
        Case 128001  ' Contratos Registrados
            lsCadImp = lsCadImp & Space(1) & "Item    ContratoAnt     Contrato            Cliente                             DOI        N°Holog. Segment.  Piezas  P.Bruto  P.Neto    Prestamo    Costo Tasac.  Costo Custod.  Impuesto   Usuario " & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        Case 128002 ' Contratos Desembolsados
            'lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt            Contrato            Prestamo    Interes      CostoTasacion   Costo Custodia    Impuesto   Neto Pagado Usuario " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt      Contrato            Cliente                             DOI        N°Holog. Segment.  Piezas  P.Bruto  P.Neto    Prestamo     Neto Pagado Usu.Reg  Usu.Desem. " & oFunI.gPrnSaltoLinea
        Case 128003 ' Contratos Anulados
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt      Contrato            Cliente                              DOI        N°Holog. Segment.  Piezas  P.Bruto  P.Neto     Prestamo  Usuario Motivo" & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        Case 128004 ' Contratos Renovados
            lsCadImp = lsCadImp & Space(1) & "Item  ContratoAnt       Contrato           Cliente                               DOI         N°Holog. Segment.  Piezas  P.Bruto  P.Neto  Tasación     Prestamo    Monto Pago     Capital     Interes     SaldCap Usuario" & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        Case 128005 ' Contratos Cancelados
            lsCadImp = lsCadImp & Space(1) & "Item  ContratoAnt       Contrato           Cliente                              DOI         N°Holog. Segment.  Piezas  P.Bruto  P.Neto  Tasación     Prestamo    Monto Pago   Capital      Interes  Usuario" & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        Case 128006 ' Contratos Rescatados
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt      Contrato            N°Holog.  Prestamo    Monto Pago   Capital     Interes     Impuesto  CostoCustod  CostoPrepRem SaldoCap  Usuario" & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        Case 128007 ' Prendas en Condicion de Recuperadas
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt            Contrato      Cliente                              DOI         N°Holog. Segment.     P.Bruto  P.Neto     Prestamo    Monto Pago DeudaVencida  Int.Mor.    Itf         Usuario Fecha Recup." & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        Case 128009 ' Prendas en Condicion
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt            Contrato       N°Holog.    Prestamo       Monto Pago       Deuda          IGV      Usuario" & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        'RECO20141028 ERS125-2014**********************
        Case 128011
            lsCadImp = lsCadImp & Space(1) & " Item     Contrato             Cliente                        Fec. Gestión     Usuario     Resultado de Llamada                    Fec. Compromiso     Motivo de llamada no concretada " & oFunI.gPrnSaltoLinea
        'RECO FIN *************************************
        ' CROB20180612 ERS076-2017 Contratos Rechazados
        Case 128012
            'lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt            Contrato            Cliente                            Prestamo  Usuario Motivo" & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt            Contrato      N°Holog.      Cliente                            Prestamo  Usuario Motivo" & oFunI.gPrnSaltoLinea
        ' CROB20180612 ERS076-2017 Contratos Rechazados
        Case 128008 ' Prendas en Condicion de Diferidas
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt            Contrato            Prestamo    Monto Pago      Capital     Interes     Impuesto   Costo Custod  Costo PrepRem     Saldo Cap Usuario" & oFunI.gPrnSaltoLinea
            
        Case 128021 ' Prendas Nuevas
            'lsCadImp = lsCadImp & Space(1) & "Item   Contrato            Val Tasac   Piezas  Oro Neto  Boveda Usuario Bloq " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Space(1) & "Item   Contrato         Cliente                               DOI         N°Holog. Segment.  Piezas  P.Bruto  P.Neto   Val Tasac  Prestamo   Boveda  Usuario Bloq " & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        Case 128022 ' Prendas Rescatadas
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt          Contrato       Cliente                               DOI         N°Holog. Segment.  Piezas  P.Bruto  P.Neto   Val Tasac    Prestamo   Boveda  Usuario Bloq " & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        Case 128023 ' Prendas en Nuevas en Condicion de Diferidas
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt          Contrato         Cliente                               DOI         N°Holog. Segment.  Piezas  P.Bruto  P.Neto   Val Tasac  Prestamo   Boveda  Usuario Bloq " & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        Case 128024 ' Prendas Recuperadas en Condicion de Adjudicadas
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt          Contrato          Cliente                               DOI         N°Holog. Segment.  Piezas  P.Bruto  P.Neto   Val Tasac  Prestamo   Boveda  Usu.  Fecha Recup." & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        Case 128025 ' Prendas Recuperadas en Condicion de Adjudicadas
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt          Contrato         N°Holog.   Val Tasac  Piezas  Oro Neto Boveda Usuario" & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        'RECO20140220***************************************************
        Case 128026
            lsCadImp = lsCadImp & Space(1) & " Item          Cliente                                Fecha        ContratoAnt       UsuIni    TasacAnt       ContratoNew     N°HologNew     TasacNew   UsuFin     Nº Piez.  OroNeto    Agencia" & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        'RECO FIN********************************************************
        Case 128031 ' Listado de Creditos Vigentes
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt           Contrato           Cliente                      Fec.Prestamo Plazo    Prestamo      SaldoCap      Tasacion  OroNeto  OroBruto  Piezas  Segmento  N° Holog.  Estado    Tasador  " & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        Case 128032 ' Listado de Creditos Diferidos
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt           Contrato           Cliente                               Tasacion  OroNeto  OroBruto  Piezas   N°Holog.  Segment.  Prestamo   Cancelac. Diferido Bov  Bloq    Tasador " & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        Case 128033 ' Listado de Creditos x Dias de Atraso
            'lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt           Contrato           Cliente             Fec.Prestamo     Tasacion  OroNeto  nPlazo   nSaldo  Dias Atraso  Interes Dev " & oFunI.gPrnSaltoLinea
            'lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt           Contrato           Cliente               Telf.   Fec.Prestamo Tasacion  OroNeto Plazo  Saldo  D.Atraso  Int.Dev.  Mora  Interes  Int.Venc." & oFunI.gPrnSaltoLinea '** PASI 20130923 RFC1304040003
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt           Contrato           Cliente               N° Holog.  Telf.   Fec.Prestamo Tasacion  OroNeto Plazo  Saldo  D.Atraso  Int.Dev.  Mora  Interes  Int.Venc." & oFunI.gPrnSaltoLinea '** APRI20190515 ERS005-2019
        Case 128034
            'lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt           Contrato           Cliente                               Tasacion  OroNeto  Fecha Adj. Bov  Bloq    Tasador " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt           Contrato           Cliente                               Tasacion  OroNeto  OroBruto  Piezas   N°Holog.  Segment.  Prestamo   Fecha Adj. Bov  Bloq    Tasador " & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
        'PEAC 20071018
        Case 128035 ' Listado de Creditos Vigentes DETALLADO POR LOTES
            lsCadImp = lsCadImp & Space(1) & "      ITEM  CANT.  DESCRIPCION                KILATE   PESO BRUTO     PESO NETO" & oFunI.gPrnSaltoLinea
        
        'PEAC 20090423
        Case 128038 ' Listado de Creditos Adjudicados DETALLADOS POR LOTES
            lsCadImp = lsCadImp & Space(1) & "      ITEM  CANT.  DESCRIPCION                KILATE   PESO BRUTO     PESO NETO" & oFunI.gPrnSaltoLinea
               
        'PEAC 20090521
        Case 128039 ' Listado de Creditos Diferdos DETALLADOS POR LOTES
            lsCadImp = lsCadImp & Space(1) & "      ITEM  CANT.  DESCRIPCION                KILATE   PESO BRUTO     PESO NETO" & oFunI.gPrnSaltoLinea

        Case 128080 ' Listado de Movimientos por Contrato
            lsCadImp = lsCadImp & "DATO       DESCRIPCION" & oFunI.gPrnSaltoLinea
        
        Case 128035 ' Listado de Contratos por Persona
            lsCadImp = lsCadImp & "DATO       DESCRIPCION" & oFunI.gPrnSaltoLinea
        
        Case 128036
            lsCadImp = lsCadImp & "Item Cuenta             Cliente                            Saldo Cap Dias Atr      Tasacion  Oro Neto" & oFunI.gPrnSaltoLinea
            
         '*** PEAC 20090312
        Case 128037 ' REPORTE DE SOBRANTE DE ADJUDICADO
            lsCadImp = lsCadImp & Space(1) & "      ITEM  CUENTA                N°HOLOG.  CLIENTE                          DOMICILIO                       TELEFONO            DEUDA      TASACION   DEVOLVER" & oFunI.gPrnSaltoLinea 'APRI20190515 ERS005-2019
            
        Case 128043
            lsCadImp = lsCadImp & "  Fecha          Saldo        Monto          Monto         Monto         Monto         Saldo " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & "                Inicial   Desembolsado     Cancelado      Renovado    Amortizado       Final  " & oFunI.gPrnSaltoLinea
        
        Case 128000 '** Juez 20120529
    End Select
    If pnCodReporte <> 128000 Then '** Juez 20120529
        lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
    End If
nRepoCabecera = lsCadImp
Set oFunI = Nothing
End Function

Public Function nRepo128001_ContratoRegis(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
        Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim lnTotPrestamo As Currency, lnTotCostoTas As Currency, lnTotCostoCus As Currency, lnTotImp As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


lsOperaciones = " = '" & gColPOpeRegContrato & "' "

    psBovedasSelect = Replace(Replace(Replace(Replace(psBovedasSelect, "'", ""), " ", ""), "(", ""), ")", "")
    psAgencia = Replace(Replace(Replace(Replace(psAgencia, "'", ""), " ", ""), "(", ""), ")", "")
    lsSQL = "exec sel_Repo128001_ContratoRegis '" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "','" & psAgencia & "','" & psUsuarioRep & "','" & psBovedasSelect & "'" 'APRI20190515 ERS005-2019
    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS REGISTRADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 200, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotPrestamo = lnTotPrestamo + !prestamo
                lnTotCostoTas = lnTotCostoTas + !CostoTas
                lnTotCostoCus = lnTotCostoCus + !CostoCus
                lnTotImp = lnTotImp + !Impuesto
                
                'Cadena de Impresion
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!NomCliente, 35, 0) & Space(1) & ofun.ImpreFormat(!cDOI, 12, 0) & Space(1) _
                    & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(6) & ofun.ImpreFormat(!cSegmento, 6, 0) & Space(1) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) & ofun.ImpreFormat(!nOroBruto, 6, 2) & Space(1) & ofun.ImpreFormat(!nOroNeto, 6, 2) & ofun.ImpreFormat(!prestamo, 10, 2, True) & Space(1) _
                    & ofun.ImpreFormat(!CostoTas, 10, 2, True) & Space(1) & ofun.ImpreFormat(!CostoCus, 10, 2, True) & Space(1) & ofun.ImpreFormat(!Impuesto, 10, 2, True) _
                    & Space(3) & ofun.ImpreFormat(!Usuario, 6, 0) _
                    & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS REGISTRADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(200, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(135) & ofun.ImpreFormat(lnTotPrestamo, 10, 2, True) & Space(1) _
            & ofun.ImpreFormat(lnTotCostoTas, 10, 2, True) & Space(1) & ofun.ImpreFormat(lnTotCostoCus, 10, 2, True) & Space(1) _
            & ofun.ImpreFormat(lnTotImp, 10, 2, True) & oFunI.gPrnSaltoLinea

        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo128001_ContratoRegis = lsCadBuffer
    Set oFunI = Nothing
End Function


Public Function nRepo128002_ContratoDesemb(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
        Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lcAges As String, lcBovedas As String '*** PEAC 20090704

Dim lnTotPrestamo As Currency, lnTotInteresAdel As Currency, lnTotCostoTas As Currency, lnTotCostoCus As Currency
Dim lnTotImp As Currency, lnTotPagado As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

'*** PEAC 20090704
'lsOperaciones = " in ('" & gColPOpeDesembolso & "','" & gColPOpeDesembolsoEFE & "') "

'On Error GoTo dError

'    lsSQL = " SELECT codigoant=isnull(r.cctacodant,'                  '),mc.cCtaCod, m.cMovNro,  mc.cOpeCod, mc.nPlazo, c.nMontoCol, " _
'        & " InteresAdel = (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
'        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodInteresCompensatorio & " ) , " _
'        & " CostoTas = (Select ISNULL(SUM(nMonto),0) From Movcoldet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
'        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodTasacion & " ) , " _
'        & " CostoCus = (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
'        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodCustodia & " ) , " _
'        & " Impuesto =  (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
'        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodImpuesto & " ) , " _
'        & " NetoPagado =  (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
'        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodMontoEntregar & " ) , " _
'        & " Right(m.cMovNro, 4) As Usuario " _
'        & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro and M.nMovFlag=0 " _
'        & " INNER JOIN Colocaciones c On  c.cCtaCod = mc.cCtaCod " _
'        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = c.cCtaCod " _
'        & " LEFT JOIN RELCUENTAS R ON R.CCTACOD=mc.CCTACOD " _
'        & " WHERE m.cOpecod " & lsOperaciones & "  " _
'        & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "' " _
'        & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' "

'        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
'            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
'        End If
        
'        If Len(Trim(psBovedasSelect)) > 0 Then
'            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
'        End If
'        lsSQL = lsSQL & " ORDER BY mc.cCtaCod "

'*** FIN PEAC
    
        lcBovedas = Replace(Replace(Replace(Replace(psBovedasSelect, "'", ""), " ", ""), "(", ""), ")", "")
        lcAges = Replace(Replace(Replace(Replace(psAgencia, "'", ""), " ", ""), "(", ""), ")", "")

        lsSQL = " exec stp_sel_ObtieneContratosPignoDesembolsados '" & lcAges & "','" & lcBovedas & "','" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "' "

    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS DESEMBOLSADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 182, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotPrestamo = lnTotPrestamo + !nMontoCol
                lnTotInteresAdel = lnTotInteresAdel + !InteresAdel
                lnTotCostoTas = lnTotCostoTas + !CostoTas
                lnTotCostoCus = lnTotCostoCus + !CostoCus
                lnTotImp = lnTotImp + !Impuesto
                lnTotPagado = lnTotPagado + !NetoPagado
                
                'Cadena de Impresion
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!NomCliente, 35, 0) & Space(1) & ofun.ImpreFormat(!cDOI, 12, 0) & Space(1) _
                    & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(6) & ofun.ImpreFormat(!cSegmento, 6, 0) & Space(1) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) & ofun.ImpreFormat(!nOroBruto, 6, 2) & Space(1) & ofun.ImpreFormat(!nOroNeto, 6, 2) _
                    & ofun.ImpreFormat(!nMontoCol, 10, 2, True) & Space(0) & Space(0) & ofun.ImpreFormat(!NetoPagado, 10, 2, True) & Space(4) & ofun.ImpreFormat(!cUserReg, 6, 0) & Space(4) & ofun.ImpreFormat(!Usuario, 6, 0) _
                    & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS DESEMBOLSADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 182, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(182, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(134) & ofun.ImpreFormat(lnTotPrestamo, 10, 2, True) & Space(0) & ofun.ImpreFormat(lnTotPagado, 10, 2, True) & oFunI.gPrnSaltoLinea

        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo128002_ContratoDesemb = lsCadBuffer
    Set oFunI = Nothing
End Function


Public Function nRepo128003_ContratoAnula(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
        Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lnTotPrestamo As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

Dim sBovedas As String

sBovedas = Replace(Replace(Replace(psBovedasSelect, "'", ""), "(", ""), ")", "")

'lsOperaciones = " = '" & gColPOpeAnula & "' "
'On Error GoTo dError

'    lsSQL = " SELECT codigoant=isnull(r.cctacodant,'                  '),mc.cCtaCod, m.cMovNro,  mc.cOpeCod, c.nMontoCol,  " _
'        & " NomCliente = (Select ISNULL(MAX(Per.cPersNombre),'') From ProductoPersona PP INNER JOIN Persona Per On PP.cPersCod = Per.cPersCod Where PP.cCtaCod = mc.cCtaCod ) , " _
'        & " Right(m.cMovNro, 4) As Usuario , m.cMovDesc " _
'        & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro " _
'        & " INNER JOIN Colocaciones c On  c.cCtaCod = mc.cCtaCod " _
'        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = mc.cCtaCod " _
'        & " LEFT JOIN RELCUENTAS R ON R.CCTACOD=mc.CCTACOD " _
'        & " WHERE m.cOpecod " & lsOperaciones & "  " _
'        & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "' " _
'        & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' "
'
'        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
'            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
'        End If
'        If Len(Trim(psBovedasSelect)) > 0 Then
'            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
'        End If
'        lsSQL = lsSQL & " ORDER BY mc.cCtaCod "

        lsSQL = " exec stp_sel_Repo128003ContratoAnula '" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "','" & psUsuarioRep & "','" & sBovedas & "','" & psAgencia & "' "

    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS ANULADOS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 165, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7

        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotPrestamo = lnTotPrestamo + !nMontoCol

                'Cadena de Impresion
                '*** PEAC 20170314-se agrego cMovDesc (motivo anulacion)
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!NomCliente, 35, 0) & Space(2) _
                    & Space(1) & ofun.ImpreFormat(!cDOI, 12, 0) & Space(1) & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(6) & ofun.ImpreFormat(!cSegmento, 6, 0) & Space(1) _
                    & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) & ofun.ImpreFormat(!nOroBruto, 6, 2) & Space(1) & ofun.ImpreFormat(!nOroNeto, 6, 2) _
                    & ofun.ImpreFormat(!nMontoCol, 10, 2, True) & Space(2) & ofun.ImpreFormat(!Usuario, 6, 0) & Space(2) & Trim(!cMovDesc) _
                    & oFunI.gPrnSaltoLinea
                
                    
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If

                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS REGISTRADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 165, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With

        lsCadImp = lsCadImp & String(165, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(137) & ofun.ImpreFormat(lnTotPrestamo, 10, 2, True) & oFunI.gPrnSaltoLinea

        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina

    End If
    nRepo128003_ContratoAnula = lsCadBuffer
    Set oFunI = Nothing
End Function

Public Function nRepo128004_ContratoRenov(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
        Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim lnTotPrestamo As Currency, lnTotPagado As Currency, lnTotCapital As Currency, lnTotInteres As Currency
Dim lnTotImp As Currency, lnTotCostoCus As Currency, lnTotCostoPrepRemate As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    Set loDataRep = New DCOMColPFunciones
    psBovedasSelect = Replace(Replace(Replace(Replace(psBovedasSelect, "'", ""), " ", ""), "(", ""), ")", "")
    psAgencia = Replace(Replace(Replace(Replace(psAgencia, "'", ""), " ", ""), "(", ""), ")", "")
    lsSQL = "exec sel_nRepo128004_ContratoRenov '" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "','" & psAgencia & "','" & psUsuarioRep & "','" & psBovedasSelect & "'" 'APRI20190515 ERS005-2019
    Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS RENOVADOS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 220, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotPrestamo = lnTotPrestamo + !nMontoCol
                lnTotPagado = lnTotPagado + !MontoPag
                lnTotCapital = lnTotCapital + !Capital
                lnTotInteres = lnTotInteres + !Interes
                lnTotImp = lnTotImp + !Impuesto
                lnTotCostoCus = lnTotCostoCus + !CostoCus
                lnTotCostoPrepRemate = lnTotCostoPrepRemate + !CostoPrepRemate

                'Cadena de Impresion
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!NomCliente, 35, 0) & Space(2) _
                    & ofun.ImpreFormat(!cDOI, 12, 0) & Space(1) & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(6) & ofun.ImpreFormat(!cSegmento, 6, 0) & Space(1) _
                    & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) & ofun.ImpreFormat(!nOroBruto, 6, 2) & Space(1) & ofun.ImpreFormat(!nOroNeto, 6, 2) & Space(1) _
                    & ofun.ImpreFormat(!nTasacion, 8, 2, True) & Space(1) & ofun.ImpreFormat(!nMontoCol, 8, 2, True) & Space(1) _
                    & ofun.ImpreFormat(!MontoPag, 8, 2, True) & Space(1) & ofun.ImpreFormat(!Capital, 8, 2, True) & Space(1) & ofun.ImpreFormat(!Interes, 8, 2, True) & Space(1) _
                    & Space(1) & ofun.ImpreFormat(!NewSaldoCap, 8, 2, True) & Space(2) & ofun.ImpreFormat(!Usuario, 4, 0) _
                    & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS RENOVADOS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 220, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(220, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(147) & ofun.ImpreFormat(lnTotPrestamo, 10, 2, True) & Space(0) _
            & ofun.ImpreFormat(lnTotPagado, 10, 2, True) & Space(0) & ofun.ImpreFormat(lnTotCapital, 10, 2, True) & Space(0) & ofun.ImpreFormat(lnTotInteres, 10, 2, True) & Space(0) _
            & oFunI.gPrnSaltoLinea
        
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo128004_ContratoRenov = lsCadBuffer
    Set oFunI = Nothing
End Function



Public Function nRepo128005_ContratoCancel(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
        Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim lnTotPrestamo As Currency, lnTotPagado As Currency, lnTotCapital As Currency, lnTotInteres As Currency
Dim lnTotImp As Currency, lnTotCostoCus As Currency, lnTotCostoPrepRemate As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


'*** PEAC 20140421
'lsOperaciones = " in ('" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" & gColPOpeCancelNorCHQ & "','" _
                & gColPOpeCancelMorCHQ & "','" & gColPOpeCanceNorEnOtCjEFE & "','" & gColPOpeCanceMorEnOtCjEFE & "' ) "
'On Error GoTo dError

'    lsSQL = " SELECT codigoant=isnull(r.cctacodant,'                  '),mc.cCtaCod, m.cMovNro,  mc.cOpeCod, mc.nPlazo, c.nMontoCol, mc.nMonto as MontoPag, mc.nSaldoCap as NewSaldoCap,  " _
        & " Capital = (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodCapital & " ) , " _
        & " Interes = (Select ISNULL(SUM(nMonto),0) From Movcoldet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod in ( " & gColPConceptoCodInteresCompensatorio & "," & gColPConceptoCodInteresMoratorio & ")  ) , " _
        & " CostoCus = (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod in ( " & gColPConceptoCodCustodia & "," & gColPConceptoCodCustodiaVencida & ")  ) , " _
        & " CostoPrepRemate =  (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodPreparaRemate & " ) , " _
        & " Impuesto =  (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodImpuesto & " ) , " _
        & " Right(m.cMovNro, 4) As Usuario " _
        & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro " _
        & " INNER JOIN Colocaciones c On  c.cCtaCod = mc.cCtaCod " _
        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = mc.cCtaCod " _
        & " LEFT JOIN RELCUENTAS R ON R.CCTACOD=mc.CCTACOD " _
        & " WHERE m.cOpecod " & lsOperaciones & "  and mc.copecod not like '99%' " _
        & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "' " _
        & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' "

'        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
'            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
'        End If

'        If Len(Trim(psBovedasSelect)) > 0 Then
'            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
'        End If
        
'        lsSQL = lsSQL & " ORDER BY mc.cCtaCod "

    psUsuarioRep = Replace(Replace(Replace(Trim(psUsuarioRep), "'", ""), "(", ""), ")", "")
    psBovedasSelect = Replace(Replace(Replace(Trim(psBovedasSelect), "'", ""), "(", ""), ")", "")

    lsSQL = "exec stp_sel_Repo128005_ContratoPigCancel '" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "','" & psAgencia & "','" & IIf(psUsuarioRep <> "<Consolidado>", psUsuarioRep, "") & "','" & IIf(Len(Trim(psBovedasSelect)) > 0, Trim(psBovedasSelect), "") & "' "

    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS CANCELADOS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 205, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotPrestamo = lnTotPrestamo + !nMontoCol
                lnTotPagado = lnTotPagado + !MontoPag
                lnTotCapital = lnTotCapital + !Capital
                lnTotInteres = lnTotInteres + !Interes
                lnTotImp = lnTotImp + !Impuesto
                lnTotCostoCus = lnTotCostoCus + !CostoCus
                lnTotCostoPrepRemate = lnTotCostoPrepRemate + !CostoPrepRemate

                'Cadena de Impresion
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 4, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!NomCliente, 35, 0) & Space(2) _
                    & ofun.ImpreFormat(!cDOI, 12, 0) & Space(1) & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(6) & ofun.ImpreFormat(!cSegmento, 6, 0) & Space(1) _
                    & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) & ofun.ImpreFormat(!nOroBruto, 6, 2) & Space(1) & ofun.ImpreFormat(!nOroNeto, 6, 2) & Space(1) _
                    & ofun.ImpreFormat(!nTasacion, 8, 2, True) & Space(1) & ofun.ImpreFormat(!nMontoCol, 8, 2, True) & Space(1) _
                    & ofun.ImpreFormat(!MontoPag, 8, 2, True) & Space(1) & ofun.ImpreFormat(!Capital, 8, 2, True) & Space(1) & ofun.ImpreFormat(!Interes, 8, 2, True) & Space(3) _
                    & ofun.ImpreFormat(!Usuario, 4, 0) _
                    & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS CANCELADOS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 251, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(251, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(148) & ofun.ImpreFormat(lnTotPrestamo, 8, 2, True) & Space(1) _
            & ofun.ImpreFormat(lnTotPagado, 8, 2, True) & Space(1) & ofun.ImpreFormat(lnTotCapital, 8, 2, True) & Space(1) & ofun.ImpreFormat(lnTotInteres, 8, 2, True) & Space(1) _
            & oFunI.gPrnSaltoLinea
        
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo128005_ContratoCancel = lsCadBuffer
    Set oFunI = Nothing
End Function


Public Function nRepo128008_PagoSobrantes(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
        Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lnTotSobrante As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


lsOperaciones = " = '" & gColPOpePagSobrante & "' "
'On Error GoTo dError

    lsSQL = " SELECT mc.cCtaCod, m.cMovNro,  mc.cOpeCod, c.nMontoCol,  " _
        & " NomCliente = (Select ISNULL(MAX(Per.cPersNombre),'') From ProductoPersona PP INNER JOIN Persona Per On PP.cPersCod = Per.cPersCod Where PP.cCtaCod = mc.cCtaCod ) , " _
        & " Right(m.cMovNro, 4) As Usuario " _
        & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro " _
        & " INNER JOIN Colocaciones c On  c.cCtaCod = mc.cCtaCod " _
        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = mc.cCtaCod " _
        & " WHERE m.cOpecod " & lsOperaciones & "  " _
        & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "' " _
        & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' And m.nMovFlag <> " & gMovFlagExtornado & " "

        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
        End If
        If Len(Trim(psBovedasSelect)) > 0 Then
            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
        End If
        lsSQL = lsSQL & " ORDER BY mc.cCtaCod "

    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE PAGO DE SOBRANTES", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotSobrante = lnTotSobrante + !nMontoCol
                
                'Cadena de Impresion
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!NomCliente, 35, 0) & Space(2) _
                    & ofun.ImpreFormat(!nMontoCol, 10, 2, True) & Space(2) & ofun.ImpreFormat(!Usuario, 6, 0) _
                    & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE PAGOS DE SOBRANTES", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(120, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(62) & ofun.ImpreFormat(lnTotSobrante, 10, 2) & oFunI.gPrnSaltoLinea

        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo128008_PagoSobrantes = lsCadBuffer
    Set oFunI = Nothing
End Function



'Public Function nRepo128008_PagoSobrantes(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
'        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
'        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String
'
'Dim lsSQL As String
'Dim lrDataRep As ADODB.Recordset
'Dim loDataRep As DCOMColPFunciones
'Dim lsCadImp As String
'Dim lsCadBuffer As String
'
'Dim lnIndice As Long
'Dim lnLineas As Integer
'Dim lnPage As Integer
'Dim lsOperaciones As String
'Dim lnTotSobrante As Currency
'
'lsOperaciones = " = '" & gColPOpePagSobrante & "' "
''On Error GoTo dError
'
'    lsSQL = " SELECT mc.cCtaCod, m.cMovNro,  mc.cOpeCod, c.nMontoCol,  " _
'        & " NomCliente = (Select ISNULL(MAX(Per.cPersNombre),'') From ProductoPersona PP INNER JOIN Persona Per On PP.cPersCod = Per.cPersCod Where PP.cCtaCod = mc.cCtaCod ) , " _
'        & " Right(m.cMovNro, 4) As Usuario " _
'        & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro " _
'        & " INNER JOIN Colocaciones c On  c.cCtaCod = mc.cCtaCod " _
'        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = mc.cCtaCod " _
'        & " WHERE m.cOpecod " & lsOperaciones & "  " _
'        & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "' " _
'        & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' "
'
'        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
'            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
'        End If
'        If Len(Trim(psBovedasSelect)) > 0 Then
'            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
'        End If
'        lsSQL = lsSQL & " ORDER BY mc.cCtaCod "
'
'    Set loDataRep = New DCOMColPFunciones
'        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
'    Set loDataRep = Nothing
'
'    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
'        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
'        Exit Function
'    Else
'        lnLineas = 0
'        lnPage = 1
'
'        'CABECERA
'        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE PAGO DE SOBRANTES", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)
'
'        lnIndice = 0:  lnLineas = 7
'
'        With lrDataRep
'            Do While Not lrDataRep.EOF
'                lnIndice = lnIndice + 1
'                lnLineas = lnLineas + 1
'                'Asigno los totales
'                lnTotSobrante = lnTotSobrante + !nMontoCol
'
'                'Cadena de Impresion
'                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!NomCliente, 35, 0) & Space(2) _
'                    & ofun.ImpreFormat(!nMontoCol, 10, 2) & Space(2) & ofun.ImpreFormat(!Usuario, 6, 0) _
'                    & oFunI.gPrnSaltoLinea
'
'                If lnIndice Mod 300 = 0 Then
'                    lsCadBuffer = lsCadBuffer & lsCadImp
'                    lsCadImp = ""
'                End If
'
'                If lnLineas >= 55 Then
'                    lnPage = lnPage + 1
'                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
'                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE PAGOS DE SOBRANTES", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)
'                    lnLineas = 7
'                End If
'                .MoveNext
'            Loop
'        End With
'
'        lsCadImp = lsCadImp & String(120, "-") & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & Space(62) & ofun.ImpreFormat(lnTotSobrante, 10, 2) & oFunI.gPrnSaltoLinea
'
'        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
'
'    End If
'    nRepo128008_PagoSobrantes = lsCadBuffer
'
'End Function

Public Function nRepo128007_ContratoAmortiz(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
        Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim lnTotPrestamo As Currency, lnTotInteresAdel As Currency, lnTotCostoTas As Currency, lnTotCostoCus As Currency
Dim lnTotImp As Currency, lnTotPagado As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora



  lsOperaciones = " in ('" & gColPOpeAmortNorEFE & "','" & gColPOpeAmortNorCHQ & "','" & gColPOpeAmortNorEnOtCjEFE & "' ) "

    Set loDataRep = New DCOMColPFunciones
    psBovedasSelect = Replace(Replace(Replace(Replace(psBovedasSelect, "'", ""), " ", ""), "(", ""), ")", "")
    psAgencia = Replace(Replace(Replace(Replace(psAgencia, "'", ""), " ", ""), "(", ""), ")", "")
    lsSQL = "exec sel_nRepo128007_ContratoAmortiz '" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "','" & psAgencia & "','" & psUsuarioRep & "','" & psBovedasSelect & "'" 'APRI20190515 ERS005-2019
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CREDITOS PIGNORATICIOS AMORTIZADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy"), lnPage, 155, "", pnOpeCod)
        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotPrestamo = lnTotPrestamo + !nMontoCol
                lnTotInteresAdel = lnTotInteresAdel + !InteresAdel
                lnTotCostoTas = lnTotCostoTas + !CostoTas
                lnTotCostoCus = lnTotCostoCus + !CostoCus
                lnTotImp = lnTotImp + !Impuesto
                lnTotPagado = lnTotPagado + !nMonto
                'Cadena de Impresion
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(1) _
                    & ofun.ImpreFormat(!nMontoCol, 8, 2, True) & Space(1) & ofun.ImpreFormat(!nMonto, 8, 2, True) & Space(1) & ofun.ImpreFormat(!Capital, 8, 2, True) & Space(1) _
                    & ofun.ImpreFormat(!InteresAdel, 8, 2, True) & Space(1) & ofun.ImpreFormat(!CostoTas, 8, 2, True) & Space(1) & ofun.ImpreFormat(!CostoCus, 8, 2) & Space(1) _
                    & ofun.ImpreFormat(!Impuesto, 8, 2, True) & Space(1) & ofun.ImpreFormat(!NewSaldoCap, 8, 2, True) & Space(3) & ofun.ImpreFormat(!Usuario, 6, 0) _
                    & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS PIGNORATICIOS AMORTIZADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy"), lnPage, 155, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(155, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(61) & ofun.ImpreFormat(lnTotPrestamo, 8, 2, True) & Space(2) & ofun.ImpreFormat(lnTotPagado, 8, 2, True) & Space(2) _
            & ofun.ImpreFormat(lnTotInteresAdel, 8, 2, True) & Space(1) & ofun.ImpreFormat(lnTotCostoTas, 8, 2, True) & Space(1) & ofun.ImpreFormat(lnTotCostoCus, 8, 2, True) & Space(1) _
            & ofun.ImpreFormat(lnTotImp, 8, 2, True) & oFunI.gPrnSaltoLinea

        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo128007_ContratoAmortiz = lsCadBuffer
    Set oFunI = Nothing
End Function

Public Function nRepo128007_ContratoRecup(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
       ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
       ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
       Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim lnTotPrestamo As Currency, lnTotDeudaVenc As Currency, lnTotIntMoratorio As Currency, lnTotItf As Currency
Dim lnTotImp As Currency, lnTotPagado As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


lsOperaciones = "='122900' "

    Set loDataRep = New DCOMColPFunciones
    psBovedasSelect = Replace(Replace(Replace(Replace(psBovedasSelect, "'", ""), " ", ""), "(", ""), ")", "")
    psAgencia = Replace(Replace(Replace(Replace(psAgencia, "'", ""), " ", ""), "(", ""), ")", "")
    lsSQL = "exec sel_nRepo128007_ContratoRecup '" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "','" & psAgencia & "','" & psUsuarioRep & "','" & psBovedasSelect & "'"  'APRI20190515 ERS005-2019
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CREDITOS PIGNORATICIOS ADJUDICADOS RECUPERADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy"), lnPage, 215, "", pnOpeCod)
        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotPrestamo = lnTotPrestamo + !nMontoCol
                lnTotDeudaVenc = lnTotDeudaVenc + !DeudaVenc
                lnTotIntMoratorio = lnTotIntMoratorio + !IntMoratorio
                'lnTotCostoCus = lnTotCostoCus + !CostoCus
                lnTotItf = lnTotItf + !ITF
                lnTotPagado = lnTotPagado + !nMonto
                'Cadena de Impresion
                '"Item   ContratoAnt            Contrato            Prestamo    Monto Pago      DeudaVencida  Interes   Itf        Usuario" & oFunI.gPrnSaltoLinea
                
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!NomCliente, 35, 0) & Space(1) _
                    & ofun.ImpreFormat(!cDOI, 12, 0) & Space(1) & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(6) & ofun.ImpreFormat(!cSegmento, 6, 0) & Space(1) _
                    & ofun.ImpreFormat(!nOroBruto, 6, 2) & Space(1) & ofun.ImpreFormat(!nOroNeto, 6, 2) & Space(1) & ofun.ImpreFormat(!nMontoCol, 8, 2, True) & Space(1) _
                    & ofun.ImpreFormat(!nMonto, 8, 2, True) & Space(1) & ofun.ImpreFormat(!DeudaVenc, 8, 2, True) & Space(1) & ofun.ImpreFormat(!IntMoratorio, 8, 2, True) & Space(1) & ofun.ImpreFormat(!ITF, 8, 2, True) _
                    & Space(7) & ofun.ImpreFormat(!Usuario, 6, 0) & Space(2) & !cFechaRec _
                    & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS PIGNORATICIOS ADJUDICADOS RECUPERADOS", "DEL " & Format(pdDesde, "dd/mm/yyyy"), lnPage, 215, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(215, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(132) & ofun.ImpreFormat(lnTotPrestamo, 8, 2, True) & Space(1) _
                            & ofun.ImpreFormat(lnTotPagado, 8, 2, True) & Space(1) & ofun.ImpreFormat(lnTotDeudaVenc, 8, 2, True) & Space(1) & ofun.ImpreFormat(lnTotIntMoratorio, 8, 2, True) & Space(2) _
                            & ofun.ImpreFormat(lnTotItf, 8, 2, True) & Space(2) & oFunI.gPrnSaltoLinea

        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo128007_ContratoRecup = lsCadBuffer

End Function

Public Function nRepo128009_ContratoVend(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
       ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
       ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
       Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim lnTotPrestamo As Currency, lnTotDeudaVenc As Currency, lnTotIntMoratorio As Currency, lnTotIGV As Currency
Dim lnTotImp As Currency, lnTotPagado As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

lsOperaciones = "='" & gColPOpeVtaSubasta & "'"

    Set loDataRep = New DCOMColPFunciones
    psBovedasSelect = Replace(Replace(Replace(Replace(psBovedasSelect, "'", ""), " ", ""), "(", ""), ")", "")
    psAgencia = Replace(Replace(Replace(Replace(psAgencia, "'", ""), " ", ""), "(", ""), ")", "")
    lsSQL = "exec sel_nRepo128009_ContratoVend '" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "','" & psAgencia & "','" & psUsuarioRep & "','" & psBovedasSelect & "'"  'APRI20190515 ERS005-2019
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CREDITOS PIGNORATICIOS ADJUDICADOS VENDIDOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotPrestamo = lnTotPrestamo + !nMontoCol
                lnTotDeudaVenc = lnTotDeudaVenc + !Deuda
                'lnTotIntMoratorio = lnTotIntMoratorio + !IntMoratorio
                'lnTotCostoCus = lnTotCostoCus + !CostoCus
                lnTotIGV = lnTotIGV + !IGV
                lnTotPagado = lnTotPagado + !nMonto
                'Cadena de Impresion
                '"Item   ContratoAnt            Contrato            Prestamo    Monto Pago      DeudaVencida  Interes   Itf        Usuario" & oFunI.gPrnSaltoLinea
                
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(2) & ofun.ImpreFormat(!nMontoCol, 10, 2, True) & Space(2) _
                    & ofun.ImpreFormat(!nMonto, 10, 2, True) & Space(2) & ofun.ImpreFormat(!Deuda, 10, 2, True) & Space(2) & ofun.ImpreFormat(!IGV, 10, 2, True) & Space(4) _
                    & ofun.ImpreFormat(!Usuario, 6, 0) _
                    & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS PIGNORATICIOS ADJUDICADOS VENDIDOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(52) & ofun.ImpreFormat(lnTotPrestamo, 10, 2, True) & Space(2) _
                            & ofun.ImpreFormat(lnTotPagado, 10, 2, True) & Space(2) & ofun.ImpreFormat(lnTotDeudaVenc, 10, 2, True) & Space(2) _
                            & ofun.ImpreFormat(lnTotIGV, 10, 2, True) & Space(2) & oFunI.gPrnSaltoLinea

        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo128009_ContratoVend = lsCadBuffer

End Function

Public Function nRepo12802_MovimientoJoyas(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
        Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lsTitulo As String

Dim lnTotValTasac As Currency, lnTotOroNeto As Currency
Dim lnTotPiezas As Long

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


Select Case pnOpeCod
    Case 128021  ' Prendas Nuevas
        lsOperaciones = "'" & gColPOpeRegContrato & "' "
        lsTitulo = "LISTADO DE PRENDAS NUEVAS"
    Case 128022  ' Prendas Rescatadas
        lsOperaciones = "'" & gColPOpeDevJoyas & "' "
        lsTitulo = "LISTADO DE PRENDAS RESCATADAS"
    Case 128023  ' Prendas Nuevas en condicion de Diferidas
        lsOperaciones = "'" & gColPOpeCancelNorEFE & "," & gColPOpeCancelMorEFE & "," _
                & gColPOpeCancelNorCHQ & "," & gColPOpeCancelMorCHQ & "'"
        lsTitulo = "LISTADO DE PRENDAS NUEVAS EN CONDICION DE DIFERIDAS"
        
    Case 128024   'Prendas Recuperadas en condición de Adjudicadas
        lsOperaciones = "'122900' "
        lsTitulo = "LISTADO DE PRENDAS RECUPERADAS EN CONDICION DE ADJUDICADAS"
        
    Case 128025   'Prendas Recuperadas en condición de Adjudicadas
        lsOperaciones = "'" & gColPOpeVtaSubasta & "'"
        lsTitulo = "LISTADO DE PRENDAS VENDIDAS EN CONDICION DE ADJUDICADAS"
        
End Select
    
    Set loDataRep = New DCOMColPFunciones
    psBovedasSelect = Replace(Replace(Replace(Replace(psBovedasSelect, "'", ""), " ", ""), "(", ""), ")", "")
    psAgencia = Replace(Replace(Replace(Replace(psAgencia, "'", ""), " ", ""), "(", ""), ")", "")
    lsSQL = "exec sel_nRepo12802_MovimientoJoyas '" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "','" & psAgencia & "'," & lsOperaciones & ",'" & psUsuarioRep & "','" & psBovedasSelect & "'"  'APRI20190515 ERS005-2019
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera(lsTitulo, "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 170, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotValTasac = lnTotValTasac + !nTasacion
                lnTotPiezas = lnTotPiezas + !nPiezas
                lnTotOroNeto = lnTotOroNeto + !nOroNeto
                
                If pnOpeCod = "128021" Then
                'Cadena de Impresion
                    lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!NomCliente, 35, 0) & Space(2) _
                    & ofun.ImpreFormat(!cDOI, 12, 0) & Space(1) & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(6) & ofun.ImpreFormat(!cSegmento, 6, 0) & Space(1) _
                    & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) & ofun.ImpreFormat(!nOroBruto, 6, 2) & Space(1) & ofun.ImpreFormat(!nOroNeto, 6, 2) & Space(1) _
                    & ofun.ImpreFormat(!nTasacion, 8, 2, True) & Space(1) & ofun.ImpreFormat(!nMontoCol, 8, 2, True) & Space(4) & ofun.ImpreFormat(!cAgeBovedA, 5, 0) _
                        & Space(2) & ofun.ImpreFormat(!Usuario, 6, 0) _
                        & oFunI.gPrnSaltoLinea
                ElseIf pnOpeCod = "128025" Then
                        lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(1) & ofun.ImpreFormat(!nTasacion, 10, 2) & Space(1) _
                        & ofun.ImpreFormat(!nPiezas, 5, 0) & Space(1) & ofun.ImpreFormat(!nOroNeto, 6, 2) & Space(4) & ofun.ImpreFormat(!cAgeBovedA, 5, 0) _
                        & Space(2) & ofun.ImpreFormat(!Usuario, 6, 0) _
                        & oFunI.gPrnSaltoLinea
                Else
                       
                    lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!NomCliente, 35, 0) & Space(2) _
                    & ofun.ImpreFormat(!cDOI, 12, 0) & Space(1) & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(6) & ofun.ImpreFormat(!cSegmento, 6, 0) & Space(1) _
                    & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) & ofun.ImpreFormat(!nOroBruto, 6, 2) & Space(1) & ofun.ImpreFormat(!nOroNeto, 6, 2) & Space(1) _
                    & ofun.ImpreFormat(!nTasacion, 8, 2, True) & Space(1) & ofun.ImpreFormat(!nMontoCol, 8, 2, True) & Space(4) & ofun.ImpreFormat(!cAgeBovedA, 5, 0) _
                        & Space(2) & ofun.ImpreFormat(!Usuario, 6, 0) & Space(1) & IIf(pnOpeCod = "128024", !FechaMov, "") _
                        & oFunI.gPrnSaltoLinea

                End If
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera(lsTitulo, "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 170, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        lsCadImp = lsCadImp & String(170, "-") & oFunI.gPrnSaltoLinea
        If pnOpeCod <> "128021" And pnOpeCod <> "128025" Then
        
'            lsCadImp = lsCadImp & Space(44) & ofun.ImpreFormat(lnTotValTasac, 10, 2) & Space(1) _
'                                        & ofun.ImpreFormat(lnTotPiezas, 5, 0) & Space(1) & ofun.ImpreFormat(lnTotOroNeto, 6, 2) & Space(1) _
'                                        & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(112) & ofun.ImpreFormat(lnTotPiezas, 5, 0) & Space(11) _
                                            & ofun.ImpreFormat(lnTotOroNeto, 6, 2) & Space(0) & ofun.ImpreFormat(lnTotValTasac, 8, 2, True) & Space(1) _
                                            & oFunI.gPrnSaltoLinea
        ElseIf pnOpeCod = "128025" Then
        
        lsCadImp = lsCadImp & Space(51) & ofun.ImpreFormat(lnTotValTasac, 10, 2, True) & Space(1) _
                    & ofun.ImpreFormat(lnTotPiezas, 5, 0) & Space(1) & ofun.ImpreFormat(lnTotOroNeto, 6, 2) & Space(1) _
                    & oFunI.gPrnSaltoLinea

        Else
            lsCadImp = lsCadImp & Space(92) & ofun.ImpreFormat(lnTotPiezas, 5, 0) & Space(11) _
                                            & ofun.ImpreFormat(lnTotOroNeto, 6, 2) & Space(0) & ofun.ImpreFormat(lnTotValTasac, 10, 2, True) & Space(1) _
                                            & oFunI.gPrnSaltoLinea
        End If
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo12802_MovimientoJoyas = lsCadBuffer

End Function

'CROB20180612 ERS076-2017 ****
Public Function nRepo128012_ContratoRechazo(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
        Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lnTotPrestamo As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

Dim sBovedas As String

sBovedas = Replace(Replace(Replace(psBovedasSelect, "'", ""), "(", ""), ")", "")

    lsSQL = " exec stp_sel_Repo128012ContratoRechazo '" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "','" & psUsuarioRep & "','" & sBovedas & "','" & psAgencia & "' "

    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS RECHAZADOS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7

        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotPrestamo = lnTotPrestamo + !nMontoCol

                'Cadena de Impresion
                '*** PEAC 20170314-se agrego cMovDesc (motivo anulacion)
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(2) & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(2) & ofun.ImpreFormat(!NomCliente, 35, 0) & Space(1) _
                    & ofun.ImpreFormat(!nMontoCol, 10, 2, True) & Space(2) & ofun.ImpreFormat(!Usuario, 6, 0) & Space(2) & Trim(!cMovDesc) _
                    & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If

                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS RECHAZADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With

        lsCadImp = lsCadImp & String(120, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(62) & ofun.ImpreFormat(lnTotPrestamo, 10, 2, True) & oFunI.gPrnSaltoLinea

        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina

    End If
    nRepo128012_ContratoRechazo = lsCadBuffer
    Set oFunI = Nothing
End Function
'CROB20180612 ERS076-2017 ****

Public Function nRepo128031_ListadoCredVigentes( _
        ByVal pdIni As Date, ByVal pdFin As Date, _
        ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal psRenovado As String, ByVal psEstado As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
        Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsRenovad As String
Dim lsEstados As String
Dim lsCondiciones  As String ' Cabecera Impresion

Dim lnTotPrestamo As Currency, lnTotSaldo As Currency, lnTotTasacion As Currency, lnTotOroNeto  As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


If psRenovado = "00" Or psRenovado = "11" Then ' Ambos NoRenov/Renov
    lsRenovad = " "
    lsCondiciones = "NO RENOVADO/RENOVADO"
ElseIf psRenovado = "10" Then  ' No Renov
    lsRenovad = " AND cp.nNroRenov = 0 "
    lsCondiciones = "NO RENOVADO"
ElseIf psRenovado = "01" Then  ' Renov
    lsRenovad = " AND cp.nNroRenov > 0 "
    lsCondiciones = "RENOVADO"
End If

If psEstado = "000" Or psEstado = "111" Then
    lsEstados = gColPEstDesem & "," & gColPEstRenov & "," & gColPEstVenci & "," & gColPEstPRema
    lsCondiciones = lsCondiciones & " // VIGENTES "
Else
    lsCondiciones = lsCondiciones & " // "
    If Mid(psEstado, 1, 1) = "1" Then ' Normal
        lsEstados = lsEstados & gColPEstDesem & "," & gColPEstRenov & ","
        lsCondiciones = lsCondiciones & " NORMALES/"
    End If
    If Mid(psEstado, 2, 1) = "1" Then ' Vencido
        lsEstados = lsEstados & gColPEstVenci & ","
        lsCondiciones = lsCondiciones & " VENCIDOS/"
    End If
    If Mid(psEstado, 3, 1) = "1" Then ' Para Remate
        lsEstados = lsEstados & gColPEstPRema & ","
        lsCondiciones = lsCondiciones & " PARA REMATE/"
    End If
    lsEstados = Mid(lsEstados, 1, Len(lsEstados) - 1)
    lsCondiciones = Mid(lsCondiciones, 1, Len(lsCondiciones) - 1)
End If

'On Error GoTo dError
'**ARLO20180102 COMENTO
'    lsSQL = " SELECT codigoant=isnull(r.cctacodant,'                  '),P.cCtaCod, c.dVigencia, c.nPlazo, c.nMontoCol, P.nSaldo, "
'    lsSQL = lsSQL & " cp.nTasacion, cp.nOroNeto, cp.cAgeBoveda,  p.nPrdEstado, Substring(Cons1.cConsDescripcion,1,10) as DescEstado,  "
'    lsSQL = lsSQL & " NomCliente = (Select ISNULL(MAX(Per.cPersNombre),'') From ProductoPersona PP INNER JOIN Persona Per "
'    lsSQL = lsSQL & "                         On PP.cPersCod = Per.cPersCod Where PP.cCtaCod = c.cCtaCod )  "
'    lsSQL = lsSQL & ", right(m.cMovNro,4) cUser" ' RIRO 20130702 SEGUN TI-ERS084-2013
'    'lsSQL = lsSQL & " FROM Producto P INNER JOIN Colocaciones c On p.cCtaCod = c.cCtaCod and c.dVigencia between '" & Format(pdIni, "yyyymmdd") & "' AND '" & Format(pdFin, "yyyymmdd") & "' "
'    lsSQL = lsSQL & " FROM Producto P INNER JOIN Colocaciones c On p.cCtaCod = c.cCtaCod and CONVERT(VARCHAR(8),c.dVigencia,112) between '" & Format(pdIni, "yyyymmdd") & "' AND '" & Format(pdFin, "yyyymmdd") & "' " 'EJVG20120103
'    lsSQL = lsSQL & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = c.cCtaCod "
'    lsSQL = lsSQL & " INNER JOIN Constante Cons1 On P.nPrdEstado = Cons1.nConsValor "
'    lsSQL = lsSQL & " LEFT  JOIN relcuentas r on r.cctacod=p.cctacod    "
'    lsSQL = lsSQL & " LEFT  JOIN movcol mc on mc.cCtaCod = p.cCtaCod and mc.cOpeCod = '120100'" ' RIRO 20130702 SEGUN TI-ERS084-2013
'    lsSQL = lsSQL & " LEFT  JOIN mov m on mc.nMovNro = m.nMovNro" ' RIRO 20130702 SEGUN TI-ERS084-2013
'    lsSQL = lsSQL & " WHERE P.nPrdEstado in ( " & lsEstados & " )  "
'    lsSQL = lsSQL & " AND Substring (p.cCtaCod, 4,2) = '" & psAgencia & "' "
'    lsSQL = lsSQL & " AND Cons1.nConsCod = " & gColocEstado
'
'        If Len(Trim(psBovedasSelect)) > 0 Then
'            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
'        End If
'
'        lsSQL = lsSQL & lsRenovad   ' Renovaciones
'        lsSQL = lsSQL & " ORDER BY p.cCtaCod "
'END ARLO

    '**ARLO20180102
    Dim sBovedasSelect As String
    sBovedasSelect = Replace(Replace(Replace(psBovedasSelect, "'", ""), "(", ""), ")", "")
    

    lsSQL = "exec stp_sel_128031_ListadoCredVigentes '" & Format(pdIni, "yyyymmdd") & "','" & Format(pdFin, "yyyymmdd") & "','" & lsEstados & "','" & psAgencia & "','" & sBovedasSelect & "'"
    'END ARLO
    
    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS", lsCondiciones, lnPage, 205, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotPrestamo = lnTotPrestamo + !nMontoCol
                lnTotSaldo = lnTotSaldo + !nSaldo
                lnTotTasacion = lnTotTasacion + !nTasacion
                lnTotOroNeto = lnTotOroNeto + !nOroNeto

                'Cadena de Impresion
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!NomCliente, 35, 0) & Space(1) _
                    & Format(!dVigencia, "dd/mm/yyyy") & Space(1) & ofun.ImpreFormat(!nPlazo, 3, 0) & Space(1) & ofun.ImpreFormat(!nMontoCol, 10, 2, True) & Space(1) _
                    & ofun.ImpreFormat(!nSaldo, 10, 2, True) & Space(1) & ofun.ImpreFormat(!nTasacion, 10, 2, True) & Space(1) & ofun.ImpreFormat(!nOroNeto, 5, 2) _
                    & Space(1) & ofun.ImpreFormat(!nOroBruto, 5, 2) & Space(3) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(8) & ofun.ImpreFormat(!cSegmento, 6, 0) & Space(1) & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(3) & ofun.ImpreFormat(!DescEstado, 10, 0) & Space(2) & !cUser _
                    & oFunI.gPrnSaltoLinea
                    
                    ' ofun.ImpreFormat(!cAgeBoveda, 5, 0)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS", lsCondiciones, lnPage, 205, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(205, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(96) & ofun.ImpreFormat(lnTotPrestamo, 10, 2, True) & Space(1) _
            & ofun.ImpreFormat(lnTotSaldo, 10, 2, True) & Space(1) & ofun.ImpreFormat(lnTotTasacion, 10, 2, True) & Space(1) & ofun.ImpreFormat(lnTotOroNeto, 5, 2) & Space(1) _
            & oFunI.gPrnSaltoLinea
        
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo128031_ListadoCredVigentes = lsCadBuffer

End Function

Public Function nRepo128032_ListadoCredDiferidos(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, Optional ByRef psmensaje As String, _
        Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsEstados As String
Dim lsCondiciones  As String ' Cabecera Impresion

Dim lnTotTasacion As Currency, lnTotOroNeto  As Currency
Dim lsFecCancel As String

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


lsEstados = gColPEstDifer
lsCondiciones = ""
'**ARLO20180102 COMENTO
'On Error GoTo dError
    'Se agregó campo cUser, RIRO 20130702 SEGUN TI-ERS084-2013
'    lsSQL = " SELECT codigoant=isnull(R.cctacodAnt,'                  '),P.cCtaCod, c.dVigencia, c.nPlazo, c.nMontoCol, P.nSaldo, " _
'        & " cp.nTasacion, cp.nOroNeto, cp.cAgeBoveda,  p.nPrdEstado, Cons1.cConsDescripcion as DescEstado,  " _
'        & " NomCliente = (Select ISNULL(MAX(Per.cPersNombre),'') From ProductoPersona PP INNER JOIN Persona Per " _
'        & "                         On PP.cPersCod = Per.cPersCod Where PP.cCtaCod = c.cCtaCod ) ,  " _
'        & " dCancelado = (SELECT Max(CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
'        & "            FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
'        & "            WHERE MC.cCtaCod = C.cCtaCod AND MC.cOpeCod in ('" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" & gColPOpeCanceNorEnOtCjEFE & "','" & gColPOpeCanceMorEnOtCjEFE & "') " _
'        & "                 AND M.nMovFlag <> " & gMovFlagExtornado & " ),  " _
'        & " cUser = ISNULL(RIGHT(m.cMovNro,4), '//')" _
'        & " FROM Producto P INNER JOIN Colocaciones c On  p.cCtaCod = c.cCtaCod " _
'        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = c.cCtaCod " _
'        & " INNER JOIN Constante Cons1 On P.nPrdEstado = Cons1.nConsValor " _
'        & " left join relcuentas r on r.cctacod=p.cctacod " _
'        & " LEFT  JOIN movcol mc on mc.cCtaCod = p.cCtaCod and mc.cOpeCod = '120100'" _
'        & " LEFT  JOIN mov m on mc.nMovNro = m.nMovNro " _
'        & " WHERE P.nPrdEstado in ( " & lsEstados & " )  " _
'        & " AND Substring (p.cCtaCod, 4,2) = '" & psAgencia & "' " _
'        & " AND Cons1.nConsCod = " & gColocEstado
'        If Len(Trim(psBovedasSelect)) > 0 Then
'            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
'        End If
'        lsSQL = lsSQL & " ORDER BY p.cCtaCod "
'END ARLO

    '**ARLO20180102
    Dim sBovedasSelect As String
    sBovedasSelect = Replace(Replace(Replace(psBovedasSelect, "'", ""), "(", ""), ")", "")

    lsSQL = "exec stp_sel_128032_ListadoCredDiferidos '" & lsEstados & "','" & psAgencia & "','" & sBovedasSelect & "'"
    'END ARLO
    
    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS EN CONDICION DE DIFERIDOS", lsCondiciones, lnPage, 195, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        Dim lsAux As String
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotTasacion = lnTotTasacion + !nTasacion
                lnTotOroNeto = lnTotOroNeto + !nOroNeto
                lsFecCancel = IIf(IsNull(!dCancelado), "   / /         ", Format(!dCancelado, "dd/mm/yyyy"))
                
                If IsDate(lsFecCancel) Then
                    lsAux = ofun.ImpreFormat(DateDiff("d", Format(lsFecCancel, "dd/mm/yyyy"), Format(csFechaSis, "dd/mm/yyyy")), 5, 0)
                End If

'ofun.ImpreFormat(!nPlazo, 3, 0) & Space(1) &
                'Cadena de Impresion
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!NomCliente, 35, 0) & Space(1) _
                    & ofun.ImpreFormat(!nTasacion, 10, 2, True) & Space(1) & ofun.ImpreFormat(!nOroNeto, 5, 2) & Space(3) _
                    & ofun.ImpreFormat(!nOroBruto, 5, 2) & Space(3) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(4) _
                    & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(8) & ofun.ImpreFormat(!cSegmento, 3, 0) & Space(0) & ofun.ImpreFormat(!nMontoCol, 10, 2, True) & Space(2) _
                    & lsFecCancel & Space(1) & IIf(IsDate(lsFecCancel), lsAux, "") _
                    & Space(3) & ofun.ImpreFormat(!cAgeBovedA, 5, 0) & Space(11) & !cUser & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS EN CONDICION DE DIFERIDOS", lsCondiciones, lnPage, 195, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(195, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(80) & ofun.ImpreFormat(lnTotTasacion, 10, 2, True) & Space(1) _
            & ofun.ImpreFormat(lnTotOroNeto, 5, 2) & Space(1) _
            & oFunI.gPrnSaltoLinea
        
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo128032_ListadoCredDiferidos = lsCadBuffer

End Function

Public Function nRepo128034_ListadoCredAdjudicados(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, Optional ByRef psmensaje As String, _
        Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer, lcBovedas As String
Dim lsEstados As String
Dim lsCondiciones  As String ' Cabecera Impresion

Dim lnTotTasacion As Currency, lnTotOroNeto  As Currency
Dim lsFecCancel As String

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


lsEstados = gColPEstAdjud
lsCondiciones = ""
'On Error GoTo dError

'*** PEAC 20091123
'    lsSQL = " SELECT codigoant=isnull(R.cctacodAnt,'                  '),P.cCtaCod, c.dVigencia, c.nPlazo, c.nMontoCol, P.nSaldo, " _
'        & " cp.nTasacion, cp.nOroNeto, cp.cAgeBoveda,  p.nPrdEstado, Cons1.cConsDescripcion as DescEstado,  " _
'        & " NomCliente = (Select ISNULL(MAX(Per.cPersNombre),'') From ProductoPersona PP INNER JOIN Persona Per " _
'        & "                         On PP.cPersCod = Per.cPersCod Where PP.cCtaCod = c.cCtaCod ) ,  " _
'        & " dCancelado = p.dprdestado  " _
'        & " FROM Producto P INNER JOIN Colocaciones c On  p.cCtaCod = c.cCtaCod " _
'        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = c.cCtaCod " _
'        & " INNER JOIN Constante Cons1 On P.nPrdEstado = Cons1.nConsValor " _
'        & " left join relcuentas r on r.cctacod=p.cctacod " _
'        & " WHERE P.nPrdEstado = " & lsEstados & "" _
'        & " AND Substring (p.cCtaCod, 4,2) = '" & psAgencia & "' " _
'        & " AND Cons1.nConsCod = " & gColocEstado
'        If Len(Trim(psBovedasSelect)) > 0 Then
'            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
'        End If
'        lsSQL = lsSQL & " ORDER BY p.cCtaCod "

        '*** PEAC 20091209 - SE QUITO LOS PARENTESIS DE LAS BOVEDAS
        If Len(Trim(psBovedasSelect)) > 0 Then
            lcBovedas = Replace(Replace(Replace(Replace(psBovedasSelect, "'", ""), " ", ""), "(", ""), ")", "")
        Else
            lcBovedas = ""
        End If
                        
        lsSQL = " exec stp_sel_ObtieneListadoCredAdjudicado '" & psAgencia & "','" & lcBovedas & "'"
'*** FIN PEAC

    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS EN CONDICION DE ADJUDICADOS", lsCondiciones, lnPage, 185, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        Dim lsAux As String
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotTasacion = lnTotTasacion + !nTasacion
                lnTotOroNeto = lnTotOroNeto + !nOroNeto
                lsFecCancel = IIf(IsNull(!dCancelado), "  /  /    ", Format(!dCancelado, "dd/mm/yyyy"))
                
                If IsDate(lsFecCancel) Then
                    lsAux = ofun.ImpreFormat(DateDiff("d", Format(lsFecCancel, "dd/mm/yyyy"), Format(csFechaSis, "dd/mm/yyyy")), 5, 0)
                End If

'ofun.ImpreFormat(!nPlazo, 3, 0) & Space(1) &
                'Cadena de Impresion
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!NomCliente, 35, 0) & Space(1) _
                    & ofun.ImpreFormat(!nTasacion, 10, 2, True) & Space(1) & ofun.ImpreFormat(!nOroNeto, 5, 2) & Space(3) _
                    & ofun.ImpreFormat(!nOroBruto, 5, 2) & Space(3) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(4) _
                    & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(8) & ofun.ImpreFormat(!cSegmento, 3, 0) & Space(0) & ofun.ImpreFormat(!nMontoCol, 10, 2, True) & Space(2) _
                    & lsFecCancel & Space(1) _
                    & Space(3) & ofun.ImpreFormat(!cAgeBovedA, 5, 0) & Space(8) & !cUser & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS EN CONDICION DE ADJUDICADOS", lsCondiciones, lnPage, 185, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(185, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(80) & ofun.ImpreFormat(lnTotTasacion, 10, 2, True) & Space(1) _
            & ofun.ImpreFormat(lnTotOroNeto, 5, 2) & Space(1) _
            & oFunI.gPrnSaltoLinea
        
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo128034_ListadoCredAdjudicados = lsCadBuffer

End Function

' PEAC 20071018 REPORTE DE CREDITOS VIGENTES DETALLADO POR LOTES

Public Function nRepo128035_ListadoCredVigentes( _
        ByVal pdIni As Date, ByVal pdFin As Date, _
        ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal psRenovado As String, ByVal psEstado As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
        Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsRenovad As String
Dim lsEstados As String
Dim lsCondiciones  As String ' Cabecera Impresion

Dim vCta As String

Dim lnTotPrestamo As Currency, lnTotSaldo As Currency, lnTotTasacion As Currency, lnTotOroNeto  As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
      
''*** PEAC 20090521
'lsSQL = " select  a.cctacod,"
'lsSQL = lsSQL & "     isnull(i.cctacodant,space(18)) cctacodant,"
'lsSQL = lsSQL & "     h.cpersnombre,"
'lsSQL = lsSQL & "     d.dVigencia,"
'lsSQL = lsSQL & "     f.cconsdescripcion Condicion,"
'lsSQL = lsSQL & "     c.nPiezas,"
'lsSQL = lsSQL & "     c.cDescrip,"
'lsSQL = lsSQL & "     c.cKilataje,"
'lsSQL = lsSQL & "     c.nPesoBruto,"
'lsSQL = lsSQL & "     c.nPesoNeto"
'lsSQL = lsSQL & " from colocpignoraticio a"
'lsSQL = lsSQL & " left join relcuentas i on a.cctacod=i.cctacod"
'lsSQL = lsSQL & " inner join colocPigJoyaDet c on a.cctacod=c.cctacod and c.cKilataje > 0"
'lsSQL = lsSQL & " inner join producto b on a.cctacod=B.CCTACOD AND b.nprdestado in (2101,2104,2106,2107)"
'lsSQL = lsSQL & " inner join Constante f on f.nconscod=3001 and b.nprdestado=f.nconsvalor"
'lsSQL = lsSQL & " inner join colocaciones d on b.cctacod=d.cctacod and d.dvigencia between '" & Format(pdIni, "yyyymmdd") & "' AND '" & Format(pdFin, "yyyymmdd") & "' "
'lsSQL = lsSQL & " inner join productopersona g on a.cctacod=g.cctacod and nprdpersrelac=20"
'lsSQL = lsSQL & " inner join persona h on g.cperscod=h.cperscod"
'lsSQL = lsSQL & " where substring(a.cctacod,4,2) = '" & psAgencia & "' "
'lsSQL = lsSQL & " order by a.cctacod,c.nitem"

''*** PEAC 20090521
lsSQL = " exec stp_sel_DetalleCredVigentesPorLote '" & psAgencia & "','" & Format(pdIni, "yyyymmdd") & "','" & Format(pdFin, "yyyymmdd") & "' "
                                
    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para esta consulta." & psAgencia
        Exit Function
    End If
    
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS VIGENTES DETALLADO POR LOTE", lsCondiciones, lnPage, 175, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnLineas = lnLineas + 1
                
                lsCadImp = lsCadImp & "CREDITO ACTUAL: " & !cCtaCod & Space(1) & "HOLOGRAMA: " & !nHolograma & Space(1) & _
                                    "CREDITO ANTIGUO: " & !cctacodant & Space(1) & _
                                    "NOMBRE: " & ofun.ImpreFormat(!cPersNombre, 30) & Space(1) & _
                                    "DESEMBOLSO: " & Format(!dVigencia, "dd/mm/yyyy") & Space(1) & _
                                    "CONDICION: " & ofun.ImpreFormat(!Condicion, 20) & _
                                    oFunI.gPrnSaltoLinea
                vCta = !cCtaCod
                lnIndice = 0
                
                    Do While !cCtaCod = vCta
                    
                    lnIndice = lnIndice + 1
                    lnLineas = lnLineas + 1
                    
                    lsCadImp = lsCadImp & Space(5) & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & _
                                        ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) & _
                                        ofun.ImpreFormat(UCase(!cDescrip), 30) & Space(1) & _
                                        ofun.ImpreFormat(!cKilataje, 4, 0) & Space(1) & _
                                        ofun.ImpreFormat(!nPesoBruto, 10, 2) & Space(1) & _
                                        ofun.ImpreFormat(!npesoneto, 10, 2) & Space(1) & _
                                        oFunI.gPrnSaltoLinea
    
                    If lnIndice Mod 300 = 0 Then
                        lsCadBuffer = lsCadBuffer & lsCadImp
                        lsCadImp = ""
                    End If
                    
                    If lnLineas >= 55 Then
                        lnPage = lnPage + 1
                        lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS VIGENTES DETALLADO POR LOTE", lsCondiciones, lnPage, 175, "", pnOpeCod)
                        lnLineas = 7
                    End If
                    .MoveNext
                    If .EOF Then
                        Exit Do
                    End If
                Loop
            Loop
        End With
        
        '***Totales
        lsCadImp = lsCadImp & String(175, "-") & oFunI.gPrnSaltoLinea
        'lsCadImp = lsCadImp & Space(76) & ofun.ImpreFormat(lnTotPrestamo, 10, 2) & Space(1) _
            & ofun.ImpreFormat(lnTotSaldo, 10, 2) & Space(1) & ofun.ImpreFormat(lnTotTasacion, 10, 2) & Space(1) & ofun.ImpreFormat(lnTotOroNeto, 5, 2) & Space(1) _
            & oFunI.gPrnSaltoLinea
        
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    nRepo128035_ListadoCredVigentes = lsCadBuffer

End Function


Public Function nRepo128036_Creditos_prendarios_fecha_determinada(ByVal psAgencia As String, ByVal pdfecha As String, ByRef psmensaje As String, _
                                                                  Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim cCliente As String * 30
Dim lsEstados As String
Dim lsCondiciones  As String ' Cabecera Impresion
Dim nItem As Long

Dim lnTotalSaldo As Double
Dim lnTotalTasacion As Double
Dim lnTotalOroNeto As Double

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


lsEstados = gColPEstDesem & "," & gColPEstRenov & "," & gColPEstVenci & "," & gColPEstPRema

'On Error GoTo dError

    lsSQL = "select "
    lsSQL = lsSQL & " SUBSTRING(P.cCtaCod, 4, 2) AS cAgencia, A.cAgeDescripcion, SUBSTRING(P.cCtaCod, 6, 3) AS cProducto, "
    lsSQL = lsSQL & " (SELECT  cConsDescripcion "
    lsSQL = lsSQL & " FROM    constante C "
    lsSQL = lsSQL & " WHERE   C.nconscod = 1001 AND C.nConsValor = substring(P.cCtaCod, 6, 3)) AS cProdDescripcion, "
    lsSQL = lsSQL & " SUBSTRING(P.cCtaCod, 9, 1) AS cMoneda, "
    lsSQL = lsSQL & " (SELECT  cConsDescripcion "
    lsSQL = lsSQL & " FROM    constante C "
    lsSQL = lsSQL & " WHERE   C.nconscod = 1011 AND C.nConsValor = substring(P.cCtaCod, 9, 1)) AS cMonDescripcion, "
    lsSQL = lsSQL & " cs.cCtaCod, cs.nSaldocap, cs.nprdestado, cs.ndiasatraso, "
    lsSQL = lsSQL & " cp.ntasacion, cp.noroneto, "
    lsSQL = lsSQL & " (select max (p.cpersnombre) from persona p join productopersona pp "
    lsSQL = lsSQL & " on p.cperscod = pp.cperscod where pp.cCtacod = cs.cCtacod and pp.nprdpersrelac = 20 ) as cCliente "
    lsSQL = lsSQL & " from Producto p "
    lsSQL = lsSQL & " INNER JOIN colocpignoraticio cp on p.cctacod  = cp.cctacod "
    lsSQL = lsSQL & " INNER JOIN colocacsaldo cs on p.cctacod = cs.cctacod "
    lsSQL = lsSQL & " INNER JOIN Agencias A ON A.cAgeCod = SUBSTRING(P.cCtaCod, 4, 2) "
    lsSQL = lsSQL & " where convert(varchar(8), cs.dfecha, 112)='" & Format(pdfecha, "YYYYMMdd") & "' "
    lsSQL = lsSQL & " and cs.nprdestado in (" & lsEstados & ") " '2101,2104,2106,2107) "
    lsSQL = lsSQL & " and SUBSTRING(P.cCtaCod, 4, 2)='" & psAgencia & "' "
    lsSQL = lsSQL & "  ORDER BY SUBSTRING(P.cCtaCod, 4, 2), SUBSTRING(P.cCtaCod, 9, 1), SUBSTRING(P.cCtaCod, 6, 3) "

    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS PRENDARIOS PARA FECHA DETERMINADA - Agencia " & NombreAgencia(psAgencia), lsCondiciones, lnPage, 145, "", 128036)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                
                nItem = nItem + 1
                
                cCliente = lrDataRep!cCliente
                
                lnTotalSaldo = lnTotalSaldo + lrDataRep!nSaldoCap
                lnTotalTasacion = lnTotalTasacion + lrDataRep!nTasacion
                lnTotalOroNeto = lnTotalOroNeto + lrDataRep!nOroNeto

                'Cadena de Impresion
                lsCadImp = lsCadImp & ofun.ImpreFormat(nItem, 4, 0) & Space(1) & lrDataRep!cCtaCod & Space(1) _
                                    & cCliente & Space(1) & ofun.ImpreFormat(lrDataRep!nSaldoCap, 10, 2) _
                                    & Space(1) & ofun.ImpreFormat(lrDataRep!nDiasAtraso, 8, 0) _
                                    & Space(1) & ofun.ImpreFormat(lrDataRep!nTasacion, 10, 2) _
                                    & Space(1) & ofun.ImpreFormat(lrDataRep!nOroNeto, 6, 2) & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS PRENDARIOS PARA FECHA DETERMINADA - Agencia " & NombreAgencia(psAgencia), lsCondiciones, lnPage, 145, "", 128036)
                    lnLineas = 7
                    
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(55) & ofun.ImpreFormat(lnTotalSaldo, 10, 2) & Space(10) _
            & ofun.ImpreFormat(lnTotalTasacion, 10, 2) & Space(1) & ofun.ImpreFormat(lnTotalOroNeto, 6, 2) & oFunI.gPrnSaltoLinea
        
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo128036_Creditos_prendarios_fecha_determinada = lsCadBuffer

End Function

Public Function NombreAgencia(ByVal lsAge As String) As String
Dim Rs As New ADODB.Recordset
Dim sql As String
Dim oCon As DCOMConecta
Set oCon = New DCOMConecta
oCon.AbreConexion

sql = "select cAgeDescripcion from Agencias where cAgeCod='" & lsAge & "'"

Set Rs = oCon.CargaRecordSet(sql)
oCon.CierraConexion
NombreAgencia = IIf(IsNull(Rs!cAgeDescripcion), "", Rs!cAgeDescripcion)
End Function


Public Function nRepo128041_Estadistica_Saldos(ByVal pdIni As Date, ByVal pdFin As Date, _
ByVal psCodAge As String, ByVal pdFecSis As Date, ByVal psNomCmac As String, _
ByVal psNomAge As String, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim pCadena As String
Dim vSaldo As Currency
Dim vNumero As Currency
Dim vLineas As Integer
Dim vPage As Integer
Dim lsSQL As String
Dim lrReg As New ADODB.Recordset
Dim lsPrestamos As String, lsCancelacion As String, lsRenovacion As String
Dim paso1 As Boolean
Dim Co As DCOMConecta
Dim gdHoraGrab As String
Dim n1 As Integer
Dim n2 As Integer
Dim n3 As Integer
Dim n4 As Integer
Dim n5 As Integer
Dim sum1 As Double
Dim sum2 As Double
Dim sum3 As Double
Dim sum4 As Double
Dim sum5 As Double

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

n1 = 0
n2 = 0
n3 = 0
n4 = 0
n5 = 0
sum1 = 0
sum2 = 0
sum3 = 0
sum4 = 0
sum5 = 0

Set Co = New DCOMConecta

If (DateDiff("d", pdIni, pdFin) + 1) > 32650 Then
    psmensaje = " Rango es demasiado amplio "
    Exit Function
End If

lsSQL = " Select CE.dEstad Fecha,"
lsSQL = lsSQL & " sum (CE.nNumPres) NroPres,"
lsSQL = lsSQL & " sum (CE.nMonPres) MonPres,"
'lsSQL = lsSQL & " sum (CE.nNumDesem) NroPres,"
'lsSQL = lsSQL & " sum (CE.nMonDesem) MonPres,"
lsSQL = lsSQL & " sum (CE.nNumCanc) NroCanc,"
lsSQL = lsSQL & " sum (CE.nMonCanc) MonCanc,"
lsSQL = lsSQL & " sum (CE.nNumReno) NroRenov,"
lsSQL = lsSQL & " sum (CE.nMonReno) MonRenov,"
lsSQL = lsSQL & " sum (CE.nNumITF) as NroItf,"
lsSQL = lsSQL & " sum (CE.nMonITF) as MonITF, "
lsSQL = lsSQL & " sum (CP.nNumCredVig) as nCredVig, "
lsSQL = lsSQL & " sum (CP.nCapVig) as nSalVig, "
lsSQL = lsSQL & " sum (CE.nNumAmort) as NroAmort,"
lsSQL = lsSQL & " sum (CE.nMonAmort) as MonAmort "
lsSQL = lsSQL & " From  ColocEstadDiaPrendMov CE "
lsSQL = lsSQL & "       JOIN ColocEstadDiaPrenda CP ON  Convert(Char(10), CE.dEstad, 112) = Convert(Char(10), CP.dEstad, 112) "
lsSQL = lsSQL & "                                       AND CE.cCodAge = CP.cCodAge and CE.cLineaCred = CP.cLineaCred"
lsSQL = lsSQL & " Where Convert(varchar(8), CE.destad, 112)>='" & Format(pdIni, "YYYYMMdd") & "' "
lsSQL = lsSQL & " And Convert(varchar(8), CE.destad, 112)<='" & Format(pdFin, "YYYYMMdd") & "' "

lsSQL = lsSQL & " and right(CE.cCodAge,2) ='" & Right(psCodAge, 2) & "' "
lsSQL = lsSQL & " Group by CE.dEstad "

Co.AbreConexion
Set lrReg = Co.CargaRecordSet(lsSQL)
Co.CierraConexion

gdHoraGrab = Format(pdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss")
vPage = 1
pCadena = pCadena & oFunI.gPrnSaltoLinea
pCadena = pCadena & ofun.ImpreFormat(psNomCmac, 35, 5) & Space(100) & "Página : " & ofun.ImpreFormat(vPage, 6, 0) & oFunI.gPrnSaltoLinea
pCadena = pCadena & ofun.ImpreFormat(psNomAge, 35, 5) & Space(99) & Format(pdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
'pCadena = pCadena & Space(4) & "Bovedas " & psBoveda & oFunI.gPrnSaltoLinea
pCadena = pCadena & Space(55) & " ESTADISTICA DE SALDOS DEL " & pdIni & " - AL " & pdFin & oFunI.gPrnSaltoLinea
pCadena = pCadena & Space(55) & String(54, "=") & oFunI.gPrnSaltoLinea
pCadena = pCadena & oFunI.gPrnSaltoLinea
pCadena = pCadena & Space(2) & String(130, "-") & oFunI.gPrnSaltoLinea
pCadena = pCadena & Space(7) & "FECHA     NRO. PREST.NUEVOS  NRO. MONTO.CANCEL.   NRO.  MONTO.RENOVA.  NRO. MONTO.AMORT.   NRO. MONTO ITF  Nro    SALDO.VIG" & oFunI.gPrnSaltoLinea
pCadena = pCadena & Space(2) & String(130, "-") & oFunI.gPrnSaltoLinea

vLineas = 13
While Not lrReg.EOF
        With lrReg
            pCadena = pCadena & Space(2) & Format(!Fecha, "DD/MM/YYYY") & ofun.ImpreFormat(!NroPres, 6, 0) & ofun.ImpreFormat(!MonPres, 12, , True) _
                    & ofun.ImpreFormat(!NroCanc, 6, 0) & ofun.ImpreFormat(!MonCanc, 12, , True) & ofun.ImpreFormat(!NroRenov, 6, 0) _
                    & ofun.ImpreFormat(!MonRenov, 12, , True) & ofun.ImpreFormat(!NroAmort, 6, 0) & ofun.ImpreFormat(!MonAmort, 12, , True) & ofun.ImpreFormat(!NroITF, 5, 0) & ofun.ImpreFormat(!MonITF, 7, , True) _
                    & ofun.ImpreFormat(!nCredVig, 6, 0) & ofun.ImpreFormat(!nSalVig, 12, , True) & oFunI.gPrnSaltoLinea
                    
            n1 = n1 + !NroPres
            n2 = n2 + !NroCanc
            n3 = n3 + !NroRenov
            n4 = n4 + !NroITF
            n5 = n5 + !NroAmort
            sum1 = sum1 + !MonPres
            sum2 = sum2 + !MonCanc
            sum3 = sum3 + !MonRenov
            sum4 = sum4 + !MonITF
            sum5 = sum5 + !MonAmort
            
        End With
        If vLineas > 66 Then  'pLineasMax
          vPage = vPage + 1
          pCadena = pCadena & oFunI.gPrnSaltoPagina
          pCadena = pCadena & Space(135) & "Página : " & ofun.ImpreFormat(vPage, 6, 0) & oFunI.gPrnSaltoLinea
          pCadena = pCadena & Space(2) & String(130, "-") & oFunI.gPrnSaltoLinea
          pCadena = pCadena & Space(7) & "FECHA      NRO. PREST.NUEVOS  NRO. MONTO.CANCEL.   NRO. MONTO.RENOVA.  NRO. MONTO.AMORT.   NRO. MONTO ITF  Nro    SALDO.VIG " & oFunI.gPrnSaltoLinea
          pCadena = pCadena & Space(2) & String(130, "-") & oFunI.gPrnSaltoLinea
          vLineas = 5
        End If
        vLineas = vLineas + 1
    lrReg.MoveNext
Wend

pCadena = pCadena & Space(2) & String(130, "-") & oFunI.gPrnSaltoLinea
pCadena = pCadena & Space(12) & ofun.ImpreFormat(n1, 6, 0) & ofun.ImpreFormat(sum1, 12, , True) & ofun.ImpreFormat(n2, 6, 0) & ofun.ImpreFormat(sum2, 12, , True) & ofun.ImpreFormat(n3, 6, 0) & ofun.ImpreFormat(sum3, 12, , True) & ofun.ImpreFormat(n5, 6, 0) & ofun.ImpreFormat(sum5, 12, , True) & ofun.ImpreFormat(n4, 6, 0) & ofun.ImpreFormat(sum4, 8, , True) & oFunI.gPrnSaltoLinea
pCadena = pCadena & Space(2) & String(130, "-") & oFunI.gPrnSaltoLinea

'Devuelve cadena con Impresión
nRepo128041_Estadistica_Saldos = pCadena & oFunI.gPrnSaltoPagina
Set Co = Nothing
Set lrReg = Nothing
End Function


'Nuevo
Public Function nRepo128049_Estadistica_Oro(sCodage, psFecIni As String, psFecFin As String, pdFecSis As Date, ByVal psNomCmac As String, _
ByVal psNomAge As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


Dim pCadena As String
Dim vSaldo As Currency
Dim vNumero As Currency
Dim vLineas As Integer
Dim vPage As Integer
Dim lsSQL As String
Dim lrReg As New ADODB.Recordset
Dim lsPrestamos As String, lsCancelacion As String, lsRenovacion As String
Dim paso1 As Boolean
Dim Co As New COMConecta.DCOMConecta, CLSAGE As New COMDConstantes.DCOMAgencias
Dim gdHoraGrab As String
Dim sumsald As Double, CANT As Integer
Dim sumpres As Double
Dim n14 As Double
Dim n16 As Double
Dim n18 As Double
Dim n21 As Double

Dim b14 As Double
Dim b16 As Double
Dim b18 As Double
Dim b21 As Double



sumsald = 0
sumpres = 0
CANT = 0

n14 = 0
n16 = 0
n18 = 0
n21 = 0

b14 = 0
b16 = 0
b18 = 0
b21 = 0


'lsSQL = lsSQL & " ISNULL((SELECT  sum(NPESOBRUTO) FROM COLOCPIGJOYADET WHERE CCTACOD=P.CCTACOD AND CKILATAJE='14' ),0) as bk14, "
'lsSQL = lsSQL & " ISNULL((SELECT  sum(NPESOBRUTO) FROM COLOCPIGJOYADET WHERE CCTACOD=P.CCTACOD AND CKILATAJE='16' ),0) as bk16, "
'lsSQL = lsSQL & " ISNULL((SELECT  sum(NPESOBRUTO) FROM COLOCPIGJOYADET WHERE CCTACOD=P.CCTACOD AND CKILATAJE='18' ),0) as bk18, "
'lsSQL = lsSQL & " ISNULL((SELECT  sum(NPESOBRUTO) FROM COLOCPIGJOYADET WHERE CCTACOD=P.CCTACOD AND CKILATAJE='21' ),0) as bk21, "
lsSQL = " Select DFECHA=convert(char(8),dEstad,112),cCodage, nCantVig,    nk14Vig ,    nk16Vig ,              nk18Vig,               nk21Vig, " _
        & " nCantDif,    nk14Dif,     nk16Dif,               nk18Dif,              nk21Dif, " _
        & " nCantAdj , nk14Adj, nk16Adj, nk18Adj, nk21Adj " _
        & " From colocestadoro " _
        & " WHERE CONVERT(CHAR(8),DESTAD,112)>='" & psFecIni & "'" _
        & "  AND CONVERT(CHAR(8),DESTAD,112)<'" & psFecFin & "' and ccodage='" & sCodage & "'"

If psFecFin = Format(pdFecSis, "yyyymmdd") Then
        lsSQL = lsSQL & " Union "
        lsSQL = lsSQL & " SELECT dfecha=convert(char(8),cast('" & Format(pdFecSis, "yyyymmdd") & "' as datetime),112) ,CAGECOD=SUBSTRING(CCTACOD,4,2) , "
        lsSQL = lsSQL & " CantidadVig=Sum(CantidadVig), "
        lsSQL = lsSQL & " VK14=SUM(nK14Vig),VK16=SUM(nK16Vig),VK18=SUM(nK18Vig),VK21=SUM(nK21Vig),  "
        lsSQL = lsSQL & " CantidadDif= sum(CantidadDif), "
        lsSQL = lsSQL & " DK14=SUM(nK14Dif),DK16=SUM(nK16Dif),DK18=SUM(nK18Dif),DK21=SUM(nK21Dif),  "
        lsSQL = lsSQL & " CantidadAdj= sum(CantidadAdj), "
        lsSQL = lsSQL & " AK14=SUM(nK14Adj),AK16=SUM(nK16Adj),AK18=SUM(nK18Adj),AK21=SUM(nK21Adj) "
        lsSQL = lsSQL & " From ( SELECT distinct CODESIAF=ISNULL(RE.CCTACODANT,'                  '), "
        lsSQL = lsSQL & " P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, "
        lsSQL = lsSQL & " CP.cAgeBoveda, C.cLineaCred,CantidadVig=case when p.nprdestado<>2108 and p.nprdestado<>2102 then 1 else  0 end , "
        lsSQL = lsSQL & " CantidadAdj=case when p.nprdestado=2108 then 1 else  0 end,  CantidadDif=case when p.nprdestado=2102 then 1 else  0 end, "
        lsSQL = lsSQL & " SUM( case when cKilataje ='14' and p.nprdestado<>2108 and p.nprdestado<>2102 then nPesoOro else 0  end ) as  nK14Vig , "
        lsSQL = lsSQL & " SUM( case when cKilataje ='16' and p.nprdestado<>2108 and p.nprdestado<>2102 then nPesoOro else 0  end ) as  nK16Vig , "
        lsSQL = lsSQL & " SUM( case when cKilataje ='18' and p.nprdestado<>2108 and p.nprdestado<>2102 then nPesoOro else 0  end ) as  nK18Vig , "
        lsSQL = lsSQL & " SUM( case when cKilataje ='21' and p.nprdestado<>2108 and p.nprdestado<>2102 then nPesoOro else 0  end ) as  nK21Vig , "
        lsSQL = lsSQL & " SUM( case when cKilataje ='14' and p.nprdestado=2108 then nPesoOro else 0  end ) as  nK14Adj , "
        lsSQL = lsSQL & " SUM( case when cKilataje ='16' and p.nprdestado=2108 then nPesoOro else 0  end ) as  nK16Adj , "
        lsSQL = lsSQL & " SUM( case when cKilataje ='18' and p.nprdestado=2108 then nPesoOro else 0  end ) as  nK18Adj , "
        lsSQL = lsSQL & " SUM( case when cKilataje ='21' and p.nprdestado=2108 then nPesoOro else 0  end ) as  nK21Adj , "
        lsSQL = lsSQL & " SUM( case when cKilataje ='14' and p.nprdestado=2102 then nPesoOro else 0  end ) as  nK14Dif , "
        lsSQL = lsSQL & " SUM( case when cKilataje ='16' and p.nprdestado=2102 then nPesoOro else 0  end ) as  nK16Dif , "
        lsSQL = lsSQL & " SUM( case when cKilataje ='18' and p.nprdestado=2102 then nPesoOro else 0  end ) as  nK18Dif , "
        lsSQL = lsSQL & " SUM( case when cKilataje ='21' and p.nprdestado=2102 then nPesoOro else 0  end ) as  nK21Dif , "
        lsSQL = lsSQL & " cNomCliente = (SELECT MAX(cPersNombre) FROM Persona Per INNER JOIN ProductoPersona PPer "
        lsSQL = lsSQL & " ON Per.cPersCod = PPer.cPersCod WHERE PPer.cCtaCod = P.cCtaCod ) "
        lsSQL = lsSQL & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod "
        lsSQL = lsSQL & " LEFT JOIN RELCUENTAS RE ON RE.CCTACOD=P.CCTACOD "
        lsSQL = lsSQL & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod "
        lsSQL = lsSQL & " inner JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod "
        lsSQL = lsSQL & " Where  P.nPrdEstado IN (2101,2104,2107,2106,2108,2102 )  and SUBSTRING(p.CCTACOD,4,2)='" & sCodage & "'"
        lsSQL = lsSQL & " GROUP BY P.cCtaCod,ISNULL(RE.CCTACODANT,'                  ') ,P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, "
        lsSQL = lsSQL & " CP.cAgeBoveda , C.cLineaCred "
        lsSQL = lsSQL & " ) J "
        lsSQL = lsSQL & " GROUP BY SUBSTRING(CCTACOD,4,2) "
End If
        lsSQL = lsSQL & " order by dfecha asc "
Co.AbreConexion
Set lrReg = Co.CargaRecordSet(lsSQL)
Co.CierraConexion

gdHoraGrab = Format(pdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss")
vPage = 1
pCadena = pCadena & oFunI.gPrnSaltoLinea
pCadena = pCadena & ofun.ImpreFormat("CMAC MAYNAS", 35, 5) & Space(100) & "Página : " & ofun.ImpreFormat(vPage, 6, 0) & oFunI.gPrnSaltoLinea
pCadena = pCadena & ofun.ImpreFormat(psNomAge, 35, 5) & Space(99) & Format(pdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
'pCadena = pCadena & Space(4) & "Bovedas " & psBoveda & oFunI.gPrnSaltoLinea
pCadena = pCadena & Space(55) & " ESTADISTICA DE ORO DEL " & Right(psFecIni, 2) & "/" & Mid(psFecIni, 5, 2) & "/" & Left(psFecIni, 4) & " AL " & Right(psFecFin, 2) & "/" & Mid(psFecFin, 5, 2) & "/" & Left(psFecFin, 4) & " - " & CLSAGE.NombreAgencia(sCodage) & oFunI.gPrnSaltoLinea
pCadena = pCadena & Space(55) & String(54, "=") & oFunI.gPrnSaltoLinea
pCadena = pCadena & oFunI.gPrnSaltoLinea
pCadena = pCadena & Space(2) & String(140, "-") & oFunI.gPrnSaltoLinea
pCadena = pCadena & Space(2) & "FECHA   CANTIDAD VIG.    VIG14K     VIG16K    VIG18K    VIG21K  TOTALVIG   CANTIDAD ADJ.  ADJ14     ADJ16     ADJ18     ADJ21  TOTALADJ " & oFunI.gPrnSaltoLinea
pCadena = pCadena & Space(2) & String(140, "-") & oFunI.gPrnSaltoLinea

vLineas = 13
While Not lrReg.EOF
        With lrReg
            pCadena = pCadena & Space(2) & Right(!dfecha, 2) & "/" & Mid(!dfecha, 5, 2) & "/" & Left(!dfecha, 4) & Space(3) & ofun.ImpreFormat(!nCantVig, 6, 0) & Space(2) & ofun.ImpreFormat(!nk14Vig, 6, 2, False) & Space(2) & ofun.ImpreFormat(!nk16Vig, 6, 2, False) & Space(2) & ofun.ImpreFormat(!nk18Vig, 6, 2, False) & Space(2) & ofun.ImpreFormat(!nk21Vig, 6, 2, False) & Space(2) & ofun.ImpreFormat(!nk14Vig + !nk16Vig + !nk18Vig + !nk21Vig, 6, 2, False) & Space(2) _
                      & ofun.ImpreFormat(!nCantAdj, 6, 0) & Space(2) & ofun.ImpreFormat(!nk14Adj, 6, 2, False) & Space(2) & ofun.ImpreFormat(!nk16Adj, 6, 2, False) & Space(2) & ofun.ImpreFormat(!nk18Adj, 6, 2, False) & Space(2) & ofun.ImpreFormat(!nk21Adj, 6, 2, False) & Space(2) & ofun.ImpreFormat(!nk14Adj + !nk16Adj + !nk18Adj + !nk21Adj, 6, 2, False) & oFunI.gPrnSaltoLinea
                    
'            sumsald = sumsald + !SALDOS
'            sumpres = sumpres + !Prestamo
'            CANT = !cantidad + CANT

            n14 = n14 + !nk14Vig
            n16 = n16 + !nk16Vig
            n18 = n18 + !nk18Vig
            n21 = n21 + !nk21Vig

            b14 = b14 + !nk14Adj
            b16 = b16 + !nk16Adj
            b18 = b18 + !nk18Adj
            b21 = b21 + !nk21Adj

            
        End With
        If vLineas > 66 Then  'pLineasMax
          vPage = vPage + 1
          pCadena = pCadena & oFunI.gPrnSaltoPagina
          pCadena = pCadena & Space(135) & "Página : " & ofun.ImpreFormat(vPage, 6, 0) & oFunI.gPrnSaltoLinea
          pCadena = pCadena & Space(2) & String(130, "-") & oFunI.gPrnSaltoLinea
          pCadena = pCadena & Space(55) & " ESTADISTICA DE ORO DEL " & Right(psFecIni, 2) & "/" & Mid(psFecIni, 5, 2) & "/" & Left(psFecIni, 4) & " AL " & Right(psFecFin, 2) & "/" & Mid(psFecFin, 5, 2) & "/" & Left(psFecFin, 4) & " - " & CLSAGE.NombreAgencia(sCodage) & oFunI.gPrnSaltoLinea
          pCadena = pCadena & Space(55) & String(54, "=") & oFunI.gPrnSaltoLinea
          pCadena = pCadena & oFunI.gPrnSaltoLinea
          pCadena = pCadena & Space(2) & String(130, "-") & oFunI.gPrnSaltoLinea
          pCadena = pCadena & Space(2) & "FECHA   CANTIDAD VIG.    VIG14K     VIG16K    VIG18K    VIG21K  TOTALVIG   CANTIDAD ADJ.  ADJ14     ADJ16     ADJ18     ADJ21  TOTALADJ " & oFunI.gPrnSaltoLinea
          pCadena = pCadena & Space(2) & String(130, "-") & oFunI.gPrnSaltoLinea
          vLineas = 13
        End If
        vLineas = vLineas + 1
    lrReg.MoveNext
Wend

'pCadena = pCadena & Space(2) & String(140, "-") & oFunI.gPrnSaltoLinea
'pCadena = pCadena & Space(8) & "TOTAL: " & Space(15) & ofun.ImpreFormat(n14, 8, , True) & ofun.ImpreFormat(n16, 8, , True) & ofun.ImpreFormat(n18, 8, , True) & ofun.ImpreFormat(n21, 8, , True) & ofun.ImpreFormat(b14, 8, , True) & ofun.ImpreFormat(b16, 8, , True) & ofun.ImpreFormat(b18, 8, , True) & ofun.ImpreFormat(b21, 8, , True) & oFunI.gPrnSaltoLinea
'pCadena = pCadena & Space(2) & String(140, "-") & oFunI.gPrnSaltoLinea
'
'Devuelve cadena con Impresión
nRepo128049_Estadistica_Oro = pCadena & oFunI.gPrnSaltoPagina
Set Co = Nothing
Set lrReg = Nothing



End Function

'** PASI 20130923 RFC130404003
Private Function fgCalculaDeuda(ByRef pnDiasAtraso As Double, ByRef pdFecVencimiento As Date, ByRef pnDiasAtrasoReal As Double, _
                            ByRef pcCredAntiguo As String, ByRef pnInteresAdel As Double, ByRef pnDiasAdel As Integer, _
                            ByRef pdFecEstado As Date, ByRef pnSaldoCapital As Double, ByRef pnTasaInteres As Double, _
                            ByRef pnInteresVencido As Double, ByRef pnInteresMoratorio As Double, ByRef pnTasaInteresVencido As Double, _
                            ByRef pnTasaInteresMoratorio As Double) As Boolean
Dim loCalculos As COMNColoCPig.NCOMColPCalculos
Dim lsmensaje As String
pnDiasAtraso = DateDiff("d", Format(pdFecVencimiento, "dd/mm/yyyy"), Format(csFechaSis, "dd/mm/yyyy"))
pnDiasAtrasoReal = pnDiasAtraso
If pnDiasAtraso <= 0 Then
    If pcCredAntiguo = "A" Then
        pnInteresAdel = Round(0, 2)
    Else
        Set loCalculos = New COMNColoCPig.NCOMColPCalculos
        pnDiasAdel = DateDiff("d", Format(pdFecEstado, "dd/mm/yyyy"), Format(csFechaSis, "dd/mm/yyyy"))
        pnInteresAdel = loCalculos.nCalculaInteresAlVencimiento(pnSaldoCapital, pnTasaInteres, pnDiasAdel)
        pnInteresAdel = Round(pnInteresAdel, 2)
        Set loCalculos = Nothing
    End If
    pnDiasAtraso = 0
    pnInteresVencido = 0
    pnInteresMoratorio = 0
Else
    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
    If pcCredAntiguo = "A" Then
        pnInteresAdel = Round(pnInteresAdel, 2)
    Else
        pnDiasAdel = DateDiff("d", Format(pdFecEstado, "dd/mm/yyyy"), Format(pdFecVencimiento, "dd/mm/yyyy"))
        pnInteresAdel = loCalculos.nCalculaInteresAlVencimiento(pnSaldoCapital, pnTasaInteres, pnDiasAdel)
        pnInteresAdel = Round(pnInteresAdel, 2)
    End If
    
    pnInteresVencido = loCalculos.nCalculaInteresMoratorio(pnSaldoCapital, pnTasaInteresVencido, pnDiasAtraso)
    pnInteresVencido = Round(pnInteresVencido, 2)
    pnInteresMoratorio = loCalculos.nCalculaInteresMoratorio(pnSaldoCapital, pnTasaInteresMoratorio, pnDiasAtraso)
    pnInteresMoratorio = Round(pnInteresMoratorio, 2)
    If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Function
    End If
    Set loCalculos = Nothing
End If
fgCalculaDeuda = True
End Function
'** FIN PASI

Public Function nRepo128033_ListadoCredxDiasAtraso(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pnDiasAIni As Integer, ByVal pnDiasAFin As Integer, ByVal pdFecSis As Date, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
        Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsRenovad As String
Dim lsEstados As String
Dim lsCondiciones  As String ' Cabecera Impresion

Dim lnTotPrestamo As Currency, lnTotSaldo As Currency, lnTotTasacion As Currency, lnTotOroNeto  As Currency

Dim vIntDev As Currency, vIntAct As Currency, vIntPro As Currency, vIntAde     As Currency
Dim vAtraso  As Integer, vDiaAct As Integer, vDiaPro As Integer
Dim vFecVen As Date

'** PASI 20130923 RFC1304040003
Dim loPigFunc As COMDColocPig.DCOMColPFunciones
Dim lsFecVenTemp As String

Dim lnDiasAtraso As Double
Dim ldFecEstado As Date
Dim lnDiasAtrasoReal As Double
Dim lcCredAntiguo As String
Dim lnDiasAdel As Integer, lnInteresAdel As Double
Dim lnSaldoCapital As Double
Dim lnTasaInteres As Double
Dim lnInteresVencido As Double
Dim lnTasaInteresVencido As Double
Dim lnInteresMoratorio As Double
Dim ldFecVencimiento As Date
Dim lnTasaInteresMoratorio As Double
Dim resultado As Boolean
'**FIN PASI

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

lsEstados = gColPEstDesem & "," & gColPEstRenov & "," & gColPEstVenci & "," & gColPEstPRema
lsCondiciones = "( " & Str(pnDiasAIni) & " - " & Str(pnDiasAFin) & " ) "


'On Error GoTo dError
'**ARLO20180102 COMENTO
'    lsSQL = " SELECT codigoant=isnull(R.cctacodAnt,'                  '),P.cCtaCod, c.dVigencia, c.nPlazo, P.nSaldo, P.nTasaInteres, C.dVenc,  " _
'        & " cp.nTasacion, cp.nOroNeto, cp.cAgeBoveda,  p.nPrdEstado,  " _
'        & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
'        & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & 1 & " ),  " _
'        & " NomCliente = (Select ISNULL(MAX(Per.cPersNombre),'') From ProductoPersona PP INNER JOIN Persona Per " _
'        & "                         On PP.cPersCod = Per.cPersCod Where PP.cCtaCod = c.cCtaCod ),  "
'
'        '**PASI 20132309
'        lsSQL = lsSQL & " TelfCliente= (SELECT ISNULL(CASE WHEN Per1.cPersTelefono = '' THEN Per1.cPersTelefono2 ELSE Per1.cPersTelefono END,'') " _
'                        & " FROM ProductoPersona PP1 " _
'                        & " INNER JOIN Persona Per1 " _
'                        & " on PP1.cPersCod=Per1.cPersCod " _
'                        & " WHERE PP1.cCtaCod=c.cCtaCod), " _
'        & " dPrdEstado=(SELECT MAX (CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 )) " _
'                    & " FROM Mov M " _
'                    & " INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
'                    & " Where MC.cCtaCod = C.cCtaCod " _
'                    & " AND MC.cOpeCod  in (120201,120202,120203,120204,121001,121011,126301,126311,121101,121102,121111,121112,126101,126102,126111,126112) AND M.nMovFlag <> 2), " _
'        & " isnull(cp.cCredB,'')cCredB, " _
'        & " nTasaIntMora=(select isnull(PT.nTasaInteres,0) " _
'                    & " from PRODUCTOTASAINTERES PT " _
'                    & " Where PT.cCtaCod = P.cCtaCod " _
'                    & " and PT.nPrdTasaInteres =3) , " _
'        & " c.nMontoCol "
'        '** END PASI
'
'        lsSQL = lsSQL & " FROM Producto P INNER JOIN Colocaciones c On  p.cCtaCod = c.cCtaCod " _
'        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = c.cCtaCod " _
'        & " LEFT JOIN RELCUENTAS R ON R.CCTACOD=P.CCTACOD " _
'        & " WHERE P.nPrdEstado in ( " & lsEstados & " )  " _
'        & " AND Substring (p.cCtaCod, 4,2) = '" & psAgencia & "' " _
'        & " AND Datediff(dd, c.dVenc, '" & Format(pdFecSis, "mm/dd/yyyy") & "') > " & pnDiasAIni & " " _
'        & " AND Datediff(dd, c.dVenc, '" & Format(pdFecSis, "mm/dd/yyyy") & "') < " & pnDiasAFin & " "
'        If Len(Trim(psBovedasSelect)) > 0 Then
'            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
'        End If
'        lsSQL = lsSQL & " ORDER BY p.cCtaCod "
'END ARLO

    '**ARLO20180102
    Dim sBovedasSelect As String
    sBovedasSelect = Replace(Replace(Replace(psBovedasSelect, "'", ""), "(", ""), ")", "")

    lsSQL = "exec stp_sel_128032_ListadoCredXDiasAtraso '" & lsEstados & "','" & psAgencia & "','" & sBovedasSelect & "','" & Format(pdFecSis, "mm/dd/yyyy") & "','" & Format(pdFecSis, "mm/dd/yyyy") & "'," & pnDiasAIni & "," & pnDiasAFin & ""
    'END ARLO
    
    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS X DIAS DE ATRASO", lsCondiciones, lnPage, 170, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                
                '** PASI 20130923 RFC1304040003
                'Asignamos valores a las variables
                lnTasaInteresVencido = lrDataRep!nTasaIntVenc
                lnTasaInteresMoratorio = lrDataRep!nTasaIntMora
                lcCredAntiguo = lrDataRep!cCredB
                ldFecEstado = lrDataRep!dPrdEstado
                lnSaldoCapital = Format(lrDataRep!nSaldo, "#0.00")
                lnTasaInteres = lrDataRep!nTasaInteres
                ldFecVencimiento = Format(lrDataRep!dVenc, "dd/mm/yyyy")
        
                'Fecha de Vencimiento es feriado
                lsFecVenTemp = ldFecVencimiento
                Set loPigFunc = New COMDColocPig.DCOMColPFunciones
                If loPigFunc.dVerSiFeriado(lsFecVenTemp, psmensaje) = True Then
                    If Trim(psmensaje) <> "" Then
                        MsgBox psmensaje, vbInformation, "Aviso"
                        Exit Function
                    End If
                    Do While True
                        lsFecVenTemp = DateAdd("d", 1, lsFecVenTemp)
                        If Not loPigFunc.dVerSiFeriado(lsFecVenTemp, psmensaje) = True Then
                            If Trim(psmensaje) <> "" Then
                                MsgBox psmensaje, vbInformation, "Aviso"
                            End If
                            Exit Do
                        End If
                    Loop
                    If lsFecVenTemp = csFechaSis Then
                        ldFecVencimiento = lsFecVenTemp
                    End If
                End If
                Set loPigFunc = Nothing
        
                resultado = fgCalculaDeuda(lnDiasAtraso, ldFecVencimiento, lnDiasAtrasoReal, lcCredAntiguo, lnInteresAdel, lnDiasAdel, ldFecEstado, lnSaldoCapital, lnTasaInteres, lnInteresVencido, lnInteresMoratorio, lnTasaInteresVencido, lnTasaInteresMoratorio)
                '** FIN PASI
                
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                
                
                vIntDev = 0:    vIntAct = 0:    vIntPro = 0:    vIntAde = 0
                vAtraso = 0:    vDiaAct = 0:    vDiaPro = 0
                vFecVen = Format(!dVenc, "dd/mm/yyyy")

                If DateDiff("d", vFecVen, pdFecSis) > 0 Then
                    vIntDev = ((1 + !nTasaIntVenc) ^ (DateDiff("d", vFecVen, pdFecSis) / 30) - 1) * !nSaldo
                    vIntDev = Round(vIntDev, 2)
                End If
                vAtraso = IIf(DateDiff("d", vFecVen, pdFecSis) > 0, DateDiff("d", vFecVen, pdFecSis), 0)
                If vAtraso = 0 Then
                    vDiaPro = IIf(DateDiff("d", pdFecSis, vFecVen) > 0, DateDiff("d", pdFecSis, vFecVen), 0)
                    vDiaAct = !nPlazo - vDiaPro
                    'vIntAde = CalculaInteresAdelantado(!nSaldoCap, !nTasaInt, !nPlazo)
                    vIntAde = !nSaldo * (1 - (1 / ((1 + !nTasaInteres) ^ (!nPlazo / 30))))
                    'vIntAde = Round(vIntAde, 2)
                    vIntAct = !nSaldo * ((((vIntAde / !nSaldo + 1) ^ (1 / !nPlazo)) ^ vDiaAct) - 1)
                    vIntAct = Round(vIntAct, 2)
                    vIntPro = (!nSaldo + vIntAct) * ((((vIntAde / !nSaldo + 1) ^ (1 / !nPlazo)) ^ vDiaPro) - 1)
                    vIntPro = Round(vIntPro, 2)
                End If
                
                    'vRTFImp = vRTFImp & ofun.ImpreFormat(vIndice, 8, 0) & Space(1) & FormatoContratro(!cCodCta) & Space(1) & _
                        Format(!dfecpres, "dd/mm/yyyy") & ofun.ImpreFormat(vNombre, 21, 1) & _
                        ofun.ImpreFormat(!nSaldoCap, 10) & ofun.ImpreFormat(!nvaltasac, 10) & _
                        ofun.ImpreFormat(!nPlazo, 3, 0) & _
                        ofun.ImpreFormat(vIntDev, 6) & ofun.ImpreFormat(vAtraso, 5, 0) & _
                        ofun.ImpreFormat(vIntAct, 9) & ofun.ImpreFormat(vIntPro, 9) & _
                        ofun.ImpreFormat(vIntAct + vIntPro, 9) & ofun.ImpreFormat(!nSaldoCap + vIntDev, 10) & _
                        ofun.ImpreFormat(!nOroNeto, 7) & pCorte

                'Asigno los totales
                'lnTotPrestamo = lnTotPrestamo + !nMontoCol
                lnTotSaldo = lnTotSaldo + !nSaldo
                lnTotTasacion = lnTotTasacion + !nTasacion
                lnTotOroNeto = lnTotOroNeto + !nOroNeto

                'Cadena de Impresion
'                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!NomCliente, 25, 0) & Space(1) _
'                    & Format(!dVigencia, "dd/mm/yyyy") & Space(1) & ofun.ImpreFormat(!nTasacion, 10, 2) & Space(1) & ofun.ImpreFormat(!noroneto, 5, 2) & Space(1) _
'                    & ofun.ImpreFormat(!nPlazo, 3, 0) & Space(1) & ofun.ImpreFormat(!nSaldo, 10, 2) & Space(1) _
'                    & ofun.ImpreFormat(vAtraso, 5, 2) & Space(1) & ofun.ImpreFormat(vIntDev, 6, 2) & Space(1) & ofun.ImpreFormat(vIntAct, 8, 2) _
'                    & Space(1) & ofun.ImpreFormat(vIntPro, 8, 2) & Space(1) & ofun.ImpreFormat(vIntAct + vIntPro, 8, 2) _
'                    & Space(1) & ofun.ImpreFormat(!nSaldo + vIntDev, 10, 2) & oFunI.gPrnSaltoLinea
                
                ''** Comentado por PASI 20130923
'                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cctacod & Space(1) & ofun.ImpreFormat(!NomCliente, 25, 0) & Space(1) _
'                    & Format(!dVigencia, "dd/mm/yyyy") & Space(1) & ofun.ImpreFormat(!nTasacion, 10, 2) & Space(1) & ofun.ImpreFormat(!noroneto, 5, 2) & Space(1) _
'                    & ofun.ImpreFormat(!nPlazo, 3, 0) & Space(1) & ofun.ImpreFormat(!nSaldo, 10, 2) & Space(1) _
'                    & ofun.ImpreFormat(vAtraso, 5, 2) & Space(1) & ofun.ImpreFormat(vIntDev, 6, 2) & Space(1) & oFunI.gPrnSaltoLinea _

                ''** Agregado por PASI 20130923
                    lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!NomCliente, 25, 0) & Space(1) _
                    & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(3) & ofun.ImpreFormat(!TelfCliente, 10, 0) & Format(!dVigencia, "dd/mm/yyyy") & Space(1) & ofun.ImpreFormat(!nTasacion, 6, 2) & Space(1) & ofun.ImpreFormat(!nOroNeto, 4, 2) & Space(1) _
                    & ofun.ImpreFormat(!nPlazo, 4, 0) & Space(1) & ofun.ImpreFormat(!nSaldo, 5, 2) & Space(1) _
                     & ofun.ImpreFormat(vAtraso, 3, 2) & Space(1) & ofun.ImpreFormat(vIntDev, 7, 2) & Space(1) _
                    & ofun.ImpreFormat(lnInteresMoratorio, 4, 2) & ofun.ImpreFormat(lnInteresAdel, 5, 2) & ofun.ImpreFormat(lnInteresVencido, 7, 2) & oFunI.gPrnSaltoLinea _
            
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS", lsCondiciones, lnPage, 170, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(65) & ofun.ImpreFormat(lnTotPrestamo, 10, 2) & Space(1) _
            & ofun.ImpreFormat(lnTotTasacion, 10, 2) & Space(2) & ofun.ImpreFormat(lnTotOroNeto, 5, 2) & Space(6) & ofun.ImpreFormat(lnTotSaldo, 10, 2) & Space(1) _
            & oFunI.gPrnSaltoLinea
        
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo128033_ListadoCredxDiasAtraso = lsCadBuffer

End Function


Public Function nRepo128080_ListadoMovsxContrato(ByVal psPersCod As String, ByVal psPersNombre As String, _
        ByVal psPersDireccDomicilio As String, ByVal psTelefono As String, ByVal psDocumento As String, ByVal psNroContrato As String, _
        ByVal sDesde As String, ByVal sHasta As String, ByVal sConFecha As Byte, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Integer
Dim lnLineas As Integer
Dim nItem As Integer
Dim lnPage As Integer
Dim loImprimeCab As NCOMColPImpre
        
Dim sOperacion As String * 30
Dim lsCondiciones As String
Dim sCriterio As String

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

'If sConFecha = 1 Then
'    sCriterio = " AND Substring (Mov.cMovNro,1,8) BETWEEN '" & Format(sDesde, "yyyymmdd") & "' AND '" & Format(sHasta, "yyyymmdd") & "' "
'    lsCondiciones = "Del " & sDesde & " Al " & sHasta
'End If

If sConFecha = 1 Then
    lsCondiciones = "Del " & sDesde & " Al " & sHasta
    '*** PEAC 20080122
    lsSQL = " exec stp_sel_ReporteKardexContratoPig '" & Trim(psNroContrato) & "','" & Format(sDesde, "yyyymmdd") & "','" & Format(sHasta, "yyyymmdd") & "'"
Else
    '*** PEAC 20080122
    lsSQL = " exec stp_sel_ReporteKardexContratoPig '" & Trim(psNroContrato) & "'"
End If

'*** PEAC 20080122
'lsSQL = "SELECT MovCol.cCtaCod, SUBSTRING(Mov.cMovNro, 1, 8) AS Fecha, OpeTpo.cOpeDesc, MovCol.nPlazo,  " & _
        " MovCol.nMonto, SUBSTRING(Mov.cMovNro,18,2) AS Agencia, SUBSTRING(Mov.cMovNro, 22,4) AS Usuario  " & _
        " FROM MovCol INNER JOIN OpeTpo ON MovCol.cOpeCod = OpeTpo.cOpeCod INNER JOIN " & _
        " Mov ON MovCol.nMovNro = Mov.nMovNro " & _
        " WHERE (MovCol.cCtaCod = '" & Trim(psNroContrato) & "') " & _
        sCriterio & _
        " ORDER BY MovCol.cCtaCod, SUBSTRING(Mov.cMovNro, 1, 8) "
    
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("MOVIMIENTOS DEL CONTRATO " & psNroContrato, lsCondiciones, lnPage, 143, "", 128034)

        lsCadImp = lsCadImp & "Codigo   : " & psPersCod & oFunI.gPrnSaltoLinea & _
                                "Nombre   : " & ofun.ImpreCarEsp(psPersNombre) & oFunI.gPrnSaltoLinea & _
                                "Direccion: " & ofun.ImpreCarEsp(psPersDireccDomicilio) & oFunI.gPrnSaltoLinea & _
                                "Telefono : " & psTelefono & oFunI.gPrnSaltoLinea & _
                                "Documento: " & psDocumento & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & String(153, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(1) & "                                                                       CAPITAL INTERES  INTERES INTERES            SALDO   DIAS                FECHA" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(1) & "ITEM  FECHA       OPERACION                       PLAZO       MONTO    PAGADO  COMPENS. VENCIDO MORATOR.  GASTOS    CAPITAL ATRASO AGE  USU     VNCTO" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(153, "-") & oFunI.gPrnSaltoLinea
        'lnLineas = lnLineas + 7

        lnIndice = 0:  lnLineas = 15
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                
                sOperacion = lrDataRep!cOpeDesc
                
                
                lsCadImp = lsCadImp & Space(1) & ofun.ImpreFormat(nItem, 4, 0) & Space(2) & Mid(lrDataRep!Fecha, 7, 2) & "/" & Mid(lrDataRep!Fecha, 5, 2) & _
                            "/" & Mid(lrDataRep!Fecha, 1, 4) & Space(2) & sOperacion & Space(2) & _
                            ofun.ImpreFormat(lrDataRep!nPlazo, 5, 0) & ofun.ImpreFormat(lrDataRep!nMonto, 10, 2) & ofun.ImpreFormat(lrDataRep!CapitalPagado, 8, 2) & ofun.ImpreFormat(lrDataRep!InteresComp, 5, 2) & ofun.ImpreFormat(lrDataRep!InteresVencido, 5, 2) & ofun.ImpreFormat(lrDataRep!InteresMora, 5, 2) & ofun.ImpreFormat(lrDataRep!nMontoNotif, 5, 2) & ofun.ImpreFormat(lrDataRep!SaldoK, 8, 2) & ofun.ImpreFormat(lrDataRep!DiasAtraso, 4, 0) & Space(4) & lrDataRep!Agencia & Space(3) & lrDataRep!Usuario & Space(3) & IIf(Format(lrDataRep!vence, "yyyymmdd") <= "19801231", Space(10), lrDataRep!vence) & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("MOVIMIENTOS DEL CONTRATO " & psNroContrato, lsCondiciones, lnPage, 143, "", 128034)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                .MoveNext
            Loop
        End With
            lsCadImp = lsCadImp & String(153, "-") & oFunI.gPrnSaltoLinea
                   
            lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
    nRepo128080_ListadoMovsxContrato = lsCadBuffer
End Function



Public Function nRepo128035_ListadoContratosxCliente(ByVal sCodAgen As String, ByVal psPersCod As String, ByVal psPersNombre As String, _
        ByVal psPersDireccDomicilio As String, ByVal psTelefono As String, ByVal psDocumento As String, ByVal psEstados As String, _
        ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal sConFecha As Byte, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim nItem As Integer
Dim lnPage As Integer
Dim loImprimeCab As NCOMColPImpre
        
Dim lsCondiciones As String
Dim sCriterio As String
Dim sMoneda As String * 3
'No total porque pueden ser en varias monedas
Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


If sConFecha = 1 Then
    'sCriterio = " AND (Producto.dPrdEstado Between '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "') "
    
    
    sCriterio = " AND (CONVERT (DATETIME, SUBSTRING(CONVERT (VARCHAR, Producto.dPrdEstado), 1, 11)) Between '" & Format(pdFecIni, "mm/dd/yyyy") & "' AND '" & Format(pdFecFin, "mm/dd/yyyy") & "') "
 
    lsCondiciones = "Del " & pdFecIni & " Al " & pdFecFin
End If
 

lsSQL = "SELECT Colocaciones.cCtaCod, Producto.dPrdEstado AS cFechaPrestamo, Colocaciones.cUltimaActualizacion, " & _
        " Colocaciones.nMontoCol, UPPER(Constante.cConsDescripcion) AS sEstado , ProductoPersona.cPersCod, dbo.FechaHoraMov(cUltimaActualizacion) AS cFechaHora " & _
        " FROM Colocaciones INNER JOIN " & _
        " ColocPignoraticio ON Colocaciones.cCtaCod = ColocPignoraticio.cCtaCod INNER JOIN " & _
        " Producto ON Colocaciones.cCtaCod = Producto.cCtaCod INNER JOIN " & _
        " Constante ON Producto.nPrdEstado = Constante.nConsValor INNER JOIN " & _
        " ProductoPersona ON Producto.cCtaCod = ProductoPersona.cCtaCod " & _
        " WHERE (Constante.nConsCod = 3001) AND (ProductoPersona.cPersCod ='" & Trim(psPersCod) & "')" & _
        " and (substring(Colocaciones.cCtaCod,4,2)='" & Trim(sCodAgen) & "') " & _
        " and (dbo.Constante.nConsValor in(" & psEstados & ")) " & _
        sCriterio & _
        " ORDER BY Colocaciones.cCtaCod "
    
    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1
        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS POR CLIENTE ", lsCondiciones, lnPage, 145, "", 128035)

        lsCadImp = lsCadImp & "Codigo   : " & psPersCod & oFunI.gPrnSaltoLinea & _
                                "Nombre   : " & ofun.ImpreCarEsp(psPersNombre) & oFunI.gPrnSaltoLinea & _
                                "Direccion: " & ofun.ImpreCarEsp(psPersDireccDomicilio) & oFunI.gPrnSaltoLinea & _
                                "Telefono : " & psTelefono & oFunI.gPrnSaltoLinea & _
                                "Documento: " & psDocumento & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(1) & "ITEM  CONTRATO            FECHA PRESTAMO          ULTIMA MODIFICACION                  MONTO  ESTADO" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea

        lnIndice = 0:  lnLineas = 15
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nItem = nItem + 1
                                 
                If CLng(Mid(!cCtaCod, 9, 1)) = gMonedaNacional Then
                    sMoneda = "S/."
                Else
                    sMoneda = "$"
                End If
                
                       
                lsCadImp = lsCadImp & Space(1) & ofun.ImpreFormat(nItem, 4, 0) & Space(2) & Trim(!cCtaCod) & Space(2) & UCase(!cFechaPrestamo) & _
                            Space(2) & Format(!cFechaHora, "dd/MM/YYYY") & " " & UCase(Format(!cFechaHora, "hh:mm:ss AMPM")) & Space(2) & _
                            sMoneda & Space(2) & ofun.ImpreFormat(!nMontoCol, 10, 2) & Space(2) & !sEstado & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS POR CLIENTE ", lsCondiciones, lnPage, 145, "", 128035)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                .MoveNext
            Loop
        End With
            lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
                   
            lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
    nRepo128035_ListadoContratosxCliente = lsCadBuffer
End Function

'*** PEAC 20090312 - REPORTE DE SOBRANTE DE ADJUDICADO

Public Function nRepo128037_ReporteSobranteAdjudicados( _
        ByVal pnOpeCod As Long, _
        ByVal psAgencia As String, _
        Optional ByRef psmensaje As String, _
        Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsRenovad As String
Dim lsEstados As String
Dim lsCondiciones  As String ' Cabecera Impresion

Dim vCta As String

Dim lnTotPrestamo As Currency, lnTotSaldo As Currency, lnTotTasacion As Currency, lnTotOroNeto  As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
      
'lsSQL = " select  a.cctacod,"
'lsSQL = lsSQL & "     isnull(i.cctacodant,space(18)) cctacodant,"
'lsSQL = lsSQL & "     h.cpersnombre,"
'lsSQL = lsSQL & "     d.dVigencia,"
'lsSQL = lsSQL & "     f.cconsdescripcion Condicion,"
'lsSQL = lsSQL & "     c.nPiezas,"
'lsSQL = lsSQL & "     c.cDescrip,"
'lsSQL = lsSQL & "     c.cKilataje,"
'lsSQL = lsSQL & "     c.nPesoBruto,"
'lsSQL = lsSQL & "     c.nPesoNeto"
'lsSQL = lsSQL & " from colocpignoraticio a"
'lsSQL = lsSQL & " left join relcuentas i on a.cctacod=i.cctacod"
'lsSQL = lsSQL & " inner join colocPigJoyaDet c on a.cctacod=c.cctacod and c.cKilataje > 0"
'lsSQL = lsSQL & " inner join producto b on a.cctacod=B.CCTACOD AND b.nprdestado in (2101,2104,2106,2107)"
'lsSQL = lsSQL & " inner join Constante f on f.nconscod=3001 and b.nprdestado=f.nconsvalor"
'lsSQL = lsSQL & " inner join colocaciones d on b.cctacod=d.cctacod and d.dvigencia between '" & Format(pdIni, "yyyymmdd") & "' AND '" & Format(pdFin, "yyyymmdd") & "' "
'lsSQL = lsSQL & " inner join productopersona g on a.cctacod=g.cctacod and nprdpersrelac=20"
'lsSQL = lsSQL & " inner join persona h on g.cperscod=h.cperscod"
'lsSQL = lsSQL & " where substring(a.cctacod,4,2) = '" & psAgencia & "' "
'lsSQL = lsSQL & " order by a.cctacod,c.nitem"

lsSQL = " exec stp_sel_ObtieneSobrantesDeAdjudicado '" & psAgencia & "' "

    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para esta consulta." & psAgencia
        Exit Function
    End If
    
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE SOBRANTE DE ADJUDICADOS", lsCondiciones, lnPage, 165, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
            
'                lnLineas = lnLineas + 1
'
'                lsCadImp = lsCadImp & "CREDITO ACTUAL: " & !cctacod & Space(1) & _
'                                    "CREDITO ANTIGUO: " & !cctacodant & Space(1) & _
'                                    "NOMBRE: " & ofun.ImpreFormat(!cPersNombre, 30) & Space(1) & _
'                                    "DESEMBOLSO: " & Format(!dVigencia, "dd/mm/yyyy") & Space(1) & _
'                                    "CONDICION: " & ofun.ImpreFormat(!Condicion, 20) & _
'                                    oFunI.gPrnSaltoLinea
'                vCta = !cctacod

'                    lnIndice = 0
'
'                    Do While !cctacod = vCta
                    
                    lnIndice = lnIndice + 1
                    lnLineas = lnLineas + 1
                    
                    lsCadImp = lsCadImp & Space(5) & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & _
                                        ofun.ImpreFormat(!cCtaCod, 18) & Space(1) & ofun.ImpreFormat(!nHolograma, 8, 0) & Space(2) & _
                                        ofun.ImpreFormat(UCase(!cPersNombre), 30) & Space(1) & _
                                        ofun.ImpreFormat(!cPersDireccDomicilio, 30) & Space(1) & _
                                        ofun.ImpreFormat(!cPersTelefono, 10) & Space(1) & _
                                        ofun.ImpreFormat(!nDeuda, 10, 2) & Space(1) & _
                                        ofun.ImpreFormat(!nTasacion, 10, 2) & Space(1) & _
                                        ofun.ImpreFormat(!nDevolver, 10, 2) & Space(1) & _
                                        oFunI.gPrnSaltoLinea
    
                    If lnIndice Mod 300 = 0 Then
                        lsCadBuffer = lsCadBuffer & lsCadImp
                        lsCadImp = ""
                    End If
                    
                    If lnLineas >= 55 Then
                        lnPage = lnPage + 1
                        lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                        lsCadImp = lsCadImp & nRepoCabecera("REPORTE DE SOBRANTE DE ADJUDICADOS", lsCondiciones, lnPage, 165, "", pnOpeCod)
                        lnLineas = 7
                    End If
                    .MoveNext
                    If .EOF Then
                        Exit Do
                    End If
                Loop
'            Loop
        End With
        
        '***Totales
        lsCadImp = lsCadImp & String(165, "-") & oFunI.gPrnSaltoLinea
        'lsCadImp = lsCadImp & Space(76) & ofun.ImpreFormat(lnTotPrestamo, 10, 2) & Space(1) _
            & ofun.ImpreFormat(lnTotSaldo, 10, 2) & Space(1) & ofun.ImpreFormat(lnTotTasacion, 10, 2) & Space(1) & ofun.ImpreFormat(lnTotOroNeto, 5, 2) & Space(1) _
            & oFunI.gPrnSaltoLinea
        
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    nRepo128037_ReporteSobranteAdjudicados = lsCadBuffer

End Function

'*** PEAC 20090416 REPORTE DE CREDITOS ADJUDICADOS DETALLADO POR LOTE

Public Function nRepo128038_ListadoCredAdjdicadosXLote(ByVal pdIni As Date, ByVal pdFin As Date, _
        ByVal pnOpeCod As Long, ByVal psAgencia As String, ByVal psRenovado As String, ByVal psEstado As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsRenovad As String
Dim lsEstados As String
Dim lsCondiciones  As String ' Cabecera Impresion

Dim vCta As String
Dim lcTitu As String
Dim lnTotPrestamo As Currency, lnTotSaldo As Currency, lnTotTasacion As Currency, lnTotOroNeto  As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim lcAge As String

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
           
lcTitu = "LISTADO DE CONTRATOS ADJUDICADOS DETALLADOS POR LOTE DEL " & CStr(Format(pdIni, "dd/mm/yyyy")) & " AL " & CStr(Format(pdFin, "dd/mm/yyyy"))

lcAge = Replace(Replace(psAgencia, "'", ""), " ", "")

lsSQL = " exec stp_sel_DetalleCredAdjudicadosPorLote '" & lcAge & "','" & Format(pdIni, "yyyymmdd") & "','" & Format(pdFin, "yyyymmdd") & "'"

    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para esta consulta." & psAgencia
        Screen.MousePointer = 0
        Exit Function
    End If
    
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera(lcTitu, lsCondiciones, lnPage, 155, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnLineas = lnLineas + 1
                
                lsCadImp = lsCadImp & "CREDITO ACTUAL: " & !cCtaCod & Space(1) & "HOLOGRAMA: " & !nHolograma & Space(1) & _
                                    "CREDITO ANTIGUO: " & IIf(Len(!cctacodant) = 0, Space(18), !cctacodant) & Space(1) & _
                                    "NOMBRE: " & ofun.ImpreFormat(!cPersNombre, 30) & Space(1) & _
                                    "ADJUDICADO: " & Format(!dProceso, "dd/mm/yyyy") & Space(1) & _
                                    oFunI.gPrnSaltoLinea
                vCta = !cCtaCod
                lnIndice = 0
                
                    Do While !cCtaCod = vCta
                    
                    lnIndice = lnIndice + 1
                    lnLineas = lnLineas + 1
                    
                    lsCadImp = lsCadImp & Space(5) & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & _
                                        ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) & _
                                        ofun.ImpreFormat(UCase(!cDescrip), 30) & Space(1) & _
                                        ofun.ImpreFormat(!cKilataje, 4, 0) & Space(1) & _
                                        ofun.ImpreFormat(!nPesoBruto, 10, 2) & Space(1) & _
                                        ofun.ImpreFormat(!npesoneto, 10, 2) & Space(1) & _
                                        oFunI.gPrnSaltoLinea
    
                    If lnIndice Mod 300 = 0 Then
                        lsCadBuffer = lsCadBuffer & lsCadImp
                        lsCadImp = ""
                    End If
                    
                    If lnLineas >= 55 Then
                        lnPage = lnPage + 1
                        lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                        lsCadImp = lsCadImp & nRepoCabecera(lcTitu, lsCondiciones, lnPage, 155, "", pnOpeCod)
                        lnLineas = 7
                    End If
                    .MoveNext
                    If .EOF Then
                        Exit Do
                    End If
                Loop
            Loop
        End With
        
        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
        
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    nRepo128038_ListadoCredAdjdicadosXLote = lsCadBuffer

End Function

''*** PEAC 20090521
Public Function nRepo128039_ListadoCredDiferidosXLote(ByVal pdIni As Date, ByVal pdFin As Date, _
        ByVal pnOpeCod As Long, ByVal psAgencia As String, ByVal psRenovado As String, ByVal psEstado As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsRenovad As String
Dim lsEstados As String
Dim lsCondiciones  As String ' Cabecera Impresion

Dim vCta As String
Dim lcTitu As String
Dim lnTotPrestamo As Currency, lnTotSaldo As Currency, lnTotTasacion As Currency, lnTotOroNeto  As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim lcAge As String

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
           
'lcTitu = "LISTADO DE CONTRATOS DIFERIDOS DETALLADOS POR LOTE DEL " & CStr(Format(pdIni, "dd/mm/yyyy")) & " AL " & CStr(Format(pdFin, "dd/mm/yyyy"))
lcTitu = "LISTADO DE CONTRATOS DIFERIDOS DETALLADOS POR LOTE"

lcAge = Replace(Replace(psAgencia, "'", ""), " ", "")

'lsSQL = " exec stp_sel_DetalleCredAdjudicadosPorLote '" & lcAge & "','" & Format(pdIni, "yyyymmdd") & "','" & Format(pdFin, "yyyymmdd") & "'"
lsSQL = " exec stp_sel_DetalleCredDiferidosPorLote '" & lcAge & "'"

    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para esta consulta." & psAgencia
        Screen.MousePointer = 0
        Exit Function
    End If
    
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera(lcTitu, lsCondiciones, lnPage, 155, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
        
        lsCadImp = lsCadImp & "===== AGENCIA: " & Trim(!cAgeDescripcion) & " =====" & Space(1) & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
            Do While Not lrDataRep.EOF
                lnLineas = lnLineas + 1
                
                lsCadImp = lsCadImp & "CREDITO ACTUAL: " & !cCtaCod & Space(1) & "HOLOGRAMA: " & !nHolograma & Space(1) & _
                                    "CREDITO ANTIGUO: " & IIf(Len(!cctacodant) = 0, Space(18), !cctacodant) & Space(1) & _
                                    "NOMBRE: " & ofun.ImpreFormat(!cPersNombre, 30) & Space(1) & _
                                    oFunI.gPrnSaltoLinea
                vCta = !cCtaCod
                lnIndice = 0
                
                    Do While !cCtaCod = vCta
                    
                    lnIndice = lnIndice + 1
                    lnLineas = lnLineas + 1
                    
                    lsCadImp = lsCadImp & Space(5) & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & _
                                        ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) & _
                                        ofun.ImpreFormat(UCase(!cDescrip), 30) & Space(1) & _
                                        ofun.ImpreFormat(!cKilataje, 4, 0) & Space(1) & _
                                        ofun.ImpreFormat(!nPesoBruto, 10, 2) & Space(1) & _
                                        ofun.ImpreFormat(!npesoneto, 10, 2) & Space(1) & _
                                        oFunI.gPrnSaltoLinea
    
                    If lnIndice Mod 300 = 0 Then
                        lsCadBuffer = lsCadBuffer & lsCadImp
                        lsCadImp = ""
                    End If
                    
                    If lnLineas >= 55 Then
                        lnPage = lnPage + 1
                        lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                        lsCadImp = lsCadImp & nRepoCabecera(lcTitu, lsCondiciones, lnPage, 155, "", pnOpeCod)
                        lnLineas = 7
                    End If
                    .MoveNext
                    If .EOF Then
                        Exit Do
                    End If
                Loop
            Loop
        End With
        
        lsCadImp = lsCadImp & String(145, "-") & oFunI.gPrnSaltoLinea
        
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    nRepo128039_ListadoCredDiferidosXLote = lsCadBuffer

End Function

'** Juez 20120529 ***************************************************************
Public Function ImprimeConsultaCreditoPignoraticio(ByVal psNroContrato As String, Optional ByVal psImpresora As Impresoras = gEPSON, Optional ByVal gdFecSis As Date) As String
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim oFunC As COMFunciones.FCOMCadenas

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim loImprimeCab As New COMNColoCPig.NCOMColPImpre
        
Dim lrCredPig As ADODB.Recordset
Dim lrCredPigJoyas As ADODB.Recordset
Dim lrCredPigPersonas As ADODB.Recordset
Dim lrCredPigJoyasDet As ADODB.Recordset
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnJoyasDet As Integer

Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
        
Dim i As Integer
Dim sKilate As String * 12
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
 
Set oFunC = New COMFunciones.FCOMCadenas

    'Carga Datos
    Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
        Set lrCredPig = loMuestraContrato.dObtieneDatosCreditoPignoraticio(psNroContrato)
        Set lrCredPigJoyas = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyas(psNroContrato)
        Set lrCredPigPersonas = loMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psNroContrato)
        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(psNroContrato)
    Set loMuestraContrato = Nothing

    If lrCredPig Is Nothing Or (lrCredPig.BOF And lrCredPig.EOF) Then
        Set lrCredPig = Nothing
        Set lrCredPigJoyas = Nothing
        Set lrCredPigPersonas = Nothing
        Set lrCredPigJoyasDet = Nothing
        Exit Function
    End If
       
    lsCondiciones = lsNegritaOn & "Crédito Nro. " & psNroContrato & "               Cliente(s): "
    
    If lrCredPigPersonas.RecordCount > 1 Then
        Do While Not lrCredPigPersonas.EOF
            lsCondiciones = lsCondiciones & "- " & lrCredPigPersonas!cPersApellido & " " & lrCredPigPersonas!cPersNombre & oFunI.gPrnSaltoLinea & String(58, " ")
            lrCredPigPersonas.MoveNext
        Loop
        lsCondiciones = lsCondiciones & lsNegritaOff
    Else
        lsCondiciones = lsCondiciones & "- " & lrCredPigPersonas!cPersApellido & " " & lrCredPigPersonas!cPersNombre & lsNegritaOff & oFunI.gPrnSaltoLinea
    End If
    lsCondiciones = lsCondiciones & lsNegritaOn & "Tipo de Contrato: " & lrCredPig!cTipCta & lsNegritaOff
    
    lnLineas = 0
    lnPage = 1

    'CABECERA
    lsCadImp = lsCadImp & nRepoCabecera("CONSULTA DE CREDITOS PIGNORATICIO", "", lnPage, 125, lsCondiciones, 128000)
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lnIndice = 0:  lnLineas = 8
    
    lsCadImp = lsCadImp & String(93, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & lsNegritaOn & "DATOS DE LAS PIEZAS" & lsNegritaOff & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & String(93, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Piezas: " & Left(lrCredPig!nPiezas & String(18, " "), 18) '& oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Tasación: " & Left(Format(lrCredPig!nTasacion, "#0.00") & String(15, " "), 15) '& oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Fec. Préstamo: " & Format(lrCredPig!dVigencia, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Oro Bruto(gr): " & Left(Format(lrCredPig!nOroBruto, "#0.00") & String(11, " "), 11) '& oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Préstamo: " & Left(Format(lrCredPig!nMontoCol, "#0.00") & String(15, " "), 15) '& oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Fec. Vencim: " & Format(lrCredPig!dVenc, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Oro Neto(gr): " & Left(Format(lrCredPig!nOroNeto, "#0.00") & String(12, " "), 12) '& oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Saldo Cap: " & Left(Format(lrCredPig!nSaldo, "#0.00") & String(14, " "), 14) '& oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "Estado: " & lrCredPig!cEstado & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    
    lsCadImp = lsCadImp & String(93, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & lsNegritaOn & "KILATE TOTAL" & lsNegritaOff & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & String(93, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "14K: " & Left(Format(IIf(IsNull(lrCredPigJoyas!nK14), "0.00", lrCredPigJoyas!nK14), "#0.00") & String(10, " "), 10) '& oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "16K: " & Format(IIf(IsNull(lrCredPigJoyas!nK16), "0.00", lrCredPigJoyas!nK16), "#0.00") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "18K: " & Left(Format(IIf(IsNull(lrCredPigJoyas!nK18), "0.00", lrCredPigJoyas!nK18), "#0.00") & String(10, " "), 10) '& oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & "21K: " & Format(IIf(IsNull(lrCredPigJoyas!nK21), "0.00", lrCredPigJoyas!nK21), "#0.00") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    
    lsCadImp = lsCadImp & lsNegritaOn & "DESCRIPCION DE LOTE" & lsNegritaOff & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & String(93, "-") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & Space(4) & "ITEM  CANT.  DESCRIPCION                         KILATE  ORO BRUTO  ORO NETO   TASACION" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & String(93, "-") & oFunI.gPrnSaltoLinea
    Do While Not lrCredPigJoyasDet.EOF
        Dim cDescrip As String * 35
        cDescrip = lrCredPigJoyasDet!cDescrip
        lsCadImp = lsCadImp & Space(4) & Right(String(4, " ") & Trim(lrCredPigJoyasDet!nItem), 4) & Space(2)
        lsCadImp = lsCadImp & Right(String(4, " ") & lrCredPigJoyasDet!nPiezas, 4) & Space(3)
        'Right(String(4, " ") & Trim(lrCredPigJoyasDet!nItem), 4)
        lsCadImp = lsCadImp & cDescrip & Space(2)
        lsCadImp = lsCadImp & Right(String(5, " ") & lrCredPigJoyasDet!cKilataje & "K", 5) & Space(2)
        lsCadImp = lsCadImp & Right(String(9, " ") & Format(lrCredPigJoyasDet!nPesoBruto, "#0.00"), 9) & Space(1)
        lsCadImp = lsCadImp & Right(String(9, " ") & Format(lrCredPigJoyasDet!npesoneto, "#0.00"), 9) & Space(1)
        lsCadImp = lsCadImp & Right(String(10, " ") & Format(lrCredPigJoyasDet!nValTasac, "#0.00"), 10)
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lrCredPigJoyasDet.MoveNext
    Loop
    lsCadImp = lsCadImp & String(93, "-") & oFunI.gPrnSaltoLinea
    lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    
    ImprimeConsultaCreditoPignoraticio = lsCadBuffer
End Function
'** End Juez ********************************************************************
'RECO20140208 ERS002*************************************************************
Public Function nRS128026_PrendasModalidadAmpliados(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
        Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lsTitulo As String

Dim lnTotValTasac As Currency, lnTotOroNeto As Currency
Dim lnTotPiezas As Long

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

 lsOperaciones = " = '128026' "
 lsTitulo = "REPORTE DE CREDITOS PIGNORATICIOS EN LA MODALIDAD DE AMPLIADOS"

'On Error GoTo dError

    lsSQL = "stp_sel_RS128026_PrendasModalidadAmpliados '" & pdDesde & "','" & pdHasta & "','" & psAgencia & "'"

    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera(lsTitulo, "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 193, "", pnOpeCod)
        'lsCadImp = lsCadImp & "Item"
        lnIndice = 0:  lnLineas = 10
        lsCadImp = lsCadImp & vbNewLine
        
            
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotValTasac = lnTotValTasac + !nTasacion
                lnTotPiezas = lnTotPiezas + !nPiezas
                lnTotOroNeto = lnTotOroNeto + !nOroNeto
                
                'Cadena de Impresion
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(2) & ofun.ImpreFormat(!cPersNombre & Space(50), 40) & Space(2) _
                        & ofun.ImpreFormat(!dFecAmpl, 10) & Space(2) & !cCtaCodAmp & Space(4) _
                        & !cUserAnt & ofun.ImpreFormat(!nTasacion, 10) & Space(5) & !cCtaCod & Space(1) & ofun.ImpreFormat(!nHolograma, 6, 0) & Space(1) _
                        & ofun.ImpreFormat(!nTasacionNew, 10) & Space(4) & !cUserAct & Space(2) & Space(2) & ofun.ImpreFormat(!nPiezas, 5) _
                        & Space(2) & ofun.ImpreFormat(!nOroNeto, 5) & Space(2) & ofun.ImpreFormat(!cAgeDescripcion & Space(30), 30) & Space(2) _
                        & oFunI.gPrnSaltoLinea
      
                    'lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!nTasacion, 10, 2) & Space(1) _
                        & ofun.ImpreFormat(!nPiezas, 5, 0) & Space(1) & ofun.ImpreFormat(!nOroNeto, 6, 2) & Space(4) & ofun.ImpreFormat(!cAgeBovedA, 5, 0) _
                        & Space(2) & ofun.ImpreFormat(!Usuario, 6, 0) _
                        & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera(lsTitulo, "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 193, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        lsCadImp = lsCadImp & String(193, "-") & oFunI.gPrnSaltoLinea
        If pnOpeCod <> "128021" Then
        
            'lsCadImp = lsCadImp & Space(44) & ofun.ImpreFormat(lnTotValTasac, 10, 2) & Space(1) _
                                        & ofun.ImpreFormat(lnTotPiezas, 5, 0) & Space(1) & ofun.ImpreFormat(lnTotOroNeto, 6, 2) & Space(1) _
                                        & oFunI.gPrnSaltoLinea
        Else
            'lsCadImp = lsCadImp & Space(25) & ofun.ImpreFormat(lnTotValTasac, 10, 2) & Space(1) _
                                            & ofun.ImpreFormat(lnTotPiezas, 5, 0) & Space(1) & ofun.ImpreFormat(lnTotOroNeto, 6, 2) & Space(1) _
                                            & oFunI.gPrnSaltoLinea
        End If
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRS128026_PrendasModalidadAmpliados = lsCadBuffer

End Function
'RECO FIN************************************************************************
'RECO20141028 ERS125-2014****************************
Public Function nRepo128011_SeguimientoCampCallCenter(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False, _
        Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lnTotSobrante As Currency

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


    lsSQL = "stp_sel_128011_SeguimientoCampCallCenter '" & Format(pdDesde, "yyyy/MM/dd") & "','" & Format(pdHasta, "yyyy/MM/dd") & "','" & psUsuarioRep & "','" & psAgencia & "'"

    Set loDataRep = New DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("SEGUIMIENTO CAMPAÑA CALL CENTER", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 185, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                'lnTotSobrante = lnTotSobrante + !nMontoCol
                
                'Cadena de Impresion
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!cPersNombre, 35, 0) & Space(2) _
                    & ofun.ImpreFormat(IIf(!dFechaReg = "1900-01-01", " ", !dFechaReg), 10, 2) & Space(7) _
                    & ofun.ImpreFormat(IIf(!cUser = "--", "", !cUser), 6, 0) & Space(2) _
                    & ofun.ImpreFormat(IIf(!cResultGestion = "--", "", !cResultGestion), 40, 2) & Space(3) _
                    & ofun.ImpreFormat(IIf(!dFechaComp = "1900-01-01", "", !dFechaComp), 10, 0) & Space(8) _
                    & ofun.ImpreFormat(IIf(!cMotivNoAtendida = "--", "", !cMotivNoAtendida), 35, 0) _
                    & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("SEGUIMIENTO CAMPAÑA CALL CENTER", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 185, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(120, "-") & oFunI.gPrnSaltoLinea
        'lsCadImp = lsCadImp & Space(62) & ofun.ImpreFormat(lnTotSobrante, 10, 2) & oFunI.gPrnSaltoLinea

        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo128011_SeguimientoCampCallCenter = lsCadBuffer
    Set oFunI = Nothing
End Function
'RECO FIN *******************************************

''*** PEAC 20170330
Public Function nObtieneCredPigAnulados(ByVal psAgencias As String, ByVal psBovedas As String, ByVal pdFecIni As Date, ByVal pdFecFin As Date) As ADODB.Recordset

Dim loValida As COMDColocPig.DCOMColPFunciones
Dim lrs As New ADODB.Recordset
Dim lsSQL As String
Dim lsAgencia As String

On Error GoTo dError

     lsSQL = "exec stp_sel_Repo128003ContratoAnula '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & "<CONSOLIDADO>" & "','" & psBovedas & "','" & psAgencias & "' "

    Set loValida = New COMDColocPig.DCOMColPFunciones
        Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    Set nObtieneCredPigAnulados = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, " <<nObtieneCredPigAnulados>> ", Err.Description
End Function


''*** PEAC 20170929
Public Function nObtieneClientesPignoDejadosDeAtender(ByVal pdfecha As Date, ByVal psAges As String) As ADODB.Recordset

Dim loValida As COMDColocPig.DCOMColPFunciones
Dim lrs As New ADODB.Recordset
Dim lsSQL As String
Dim lsAgencia As String

On Error GoTo dError

     lsSQL = "exec stp_sel_ObtieneNoVueltosAtenderPigno '" & Format(pdfecha, "yyyymmdd") & "','" & psAges & "' "

    Set loValida = New COMDColocPig.DCOMColPFunciones
        Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    Set nObtieneClientesPignoDejadosDeAtender = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, " <<nObtieneClientesPignoDejadosDeAtender>> ", Err.Description
End Function



