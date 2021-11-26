VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NColPRepo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim csNomCMAC As String
Dim csNomAgencia As String
Dim csCodUser As String
Dim csFechaSis As String

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lsArchivoN As String, lbLibroOpen As Boolean

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
    lsCadImp = lsCadImp & Chr(10) & String(pnAnchoLinea, "-") & Chr(10)
    Select Case pnCodReporte
        Case 128001  ' Contratos Registrados
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt            Contrato            Prestamo    Costo Tasac.  Costo Custod.    Impuesto   Usuario " & Chr(10)
        Case 128002 ' Contratos Desembolsados
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt            Contrato            Prestamo    Interes      CostoTasacion   Costo Custodia    Impuesto   Neto Pagado Usuario " & Chr(10)
        Case 128003 ' Contratos Anulados
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt            Contrato            Cliente                            Prestamo  Usuario " & Chr(10)
        Case 128004 ' Contratos Renovados
            lsCadImp = lsCadImp & Space(1) & "Item  ContratoAnt       Contrato           Prestamo    Monto Pago   Capital    Interes    Impuesto   Costo Cust.  Costo PrepRem  SaldoCap Usuario" & Chr(10)
        Case 128005 ' Contratos Cancelados
            lsCadImp = lsCadImp & Space(1) & "Item  ContratoAnt       Contrato           Prestamo    Monto Pago   Capital    Interes    Impuesto   Costo Cust.  Costo PrepRem  SaldoCap Usuario" & Chr(10)
        Case 128006 ' Contratos Rescatados
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt            Contrato            Prestamo    Monto Pago      Capital     Interes     Impuesto   Costo Custod  Costo PrepRem     Saldo Cap Usuario" & Chr(10)
        Case 128007 ' Prendas en Condicion de Recuperadas
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt            Contrato            Prestamo    Monto Pago      DeudaVencida  Int. Moratorio    Itf         Usuario" & Chr(10)
        Case 128009 ' Prendas en Condicion
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt            Contrato            Prestamo    Monto Pago        Deuda          IGV       Usuario" & Chr(10)
            
        Case 128008 ' Prendas en Condicion de Diferidas
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt            Contrato            Prestamo    Monto Pago      Capital     Interes     Impuesto   Costo Custod  Costo PrepRem     Saldo Cap Usuario" & Chr(10)
            
        Case 128021 ' Prendas Nuevas
            lsCadImp = lsCadImp & Space(1) & "Item   Contrato            Val Tasac   Piezas  Oro Neto  Boveda Usuario Bloq " & Chr(10)
        Case 128022 ' Prendas Rescatadas
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt          Contrato             Val Tasac  Piezas  Oro Neto  Boveda Usuario  Bloq " & Chr(10)
        Case 128023 ' Prendas en Nuevas en Condicion de Diferidas
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt          Contrato             Val Tasac  Piezas  Oro Neto  Boveda Usuario  Bloq " & Chr(10)
        Case 128024 ' Prendas Recuperadas en Condicion de Adjudicadas
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt          Contrato             Val Tasac  Piezas  Oro Neto Boveda Usuario" & Chr(10)
        Case 128025 ' Prendas Recuperadas en Condicion de Adjudicadas
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt          Contrato             Val Tasac  Piezas  Oro Neto Boveda Usuario" & Chr(10)
        Case 128031 ' Listado de Creditos Vigentes
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt           Contrato           Cliente                      Fec.Prestamo Plazo    Prestamo      SaldoCap      Tasacion  OroNeto Estado    " & Chr(10)
        Case 128032 ' Listado de Creditos Diferidos
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt           Contrato           Cliente                               Tasacion  OroNeto  Cancelac. Diferido Bov  Bloq " & Chr(10)
        Case 128033 ' Listado de Creditos x Dias de Atraso
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt           Contrato           Cliente                      Fec.Prestamo Plazo      Tasacion  OroNeto  Cancelac. Diferido Bov  Bloq " & Chr(10)
        Case 128034
            lsCadImp = lsCadImp & Space(1) & "Item   ContratoAnt           Contrato           Cliente                               Tasacion  OroNeto  Fecha Adj. Bov  Bloq " & Chr(10)
        
        Case 128080 ' Listado de Movimientos por Contrato
            lsCadImp = lsCadImp & "DATO       DESCRIPCION" & Chr(10)
        
        Case 128035 ' Listado de Contratos por Persona
            lsCadImp = lsCadImp & "DATO       DESCRIPCION" & Chr(10)
        
        Case 128036
            lsCadImp = lsCadImp & "Item Cuenta             Cliente                            Saldo Cap Dias Atr      Tasacion  Oro Neto" & Chr(10)
            
        Case 128043
            lsCadImp = lsCadImp & "  Fecha          Saldo        Monto          Monto         Monto         Monto         Saldo " & Chr(10)
            lsCadImp = lsCadImp & "                Inicial   Desembolsado     Cancelado      Renovado    Amortizado       Final  " & Chr(10)
            
    End Select
    lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & Chr(10)
nRepoCabecera = lsCadImp
End Function

Public Function nRepo128001_ContratoRegis(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim lnTotPrestamo As Currency, lnTotCostoTas As Currency, lnTotCostoCus As Currency, lnTotImp As Currency

lsOperaciones = " = '" & gColPOpeRegContrato & "' "
'On Error GoTo dError

    lsSQL = " SELECT codigoant=isnull(r.cctacodant,'                  '),mc.cCtaCod, m.cMovNro,  mc.cOpeCod, mc.nPlazo, " _
        & " Prestamo = (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodCapital & " ) , " _
        & " CostoTas = (Select ISNULL(SUM(nMonto),0) From Movcoldet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodTasacion & " ) , " _
        & " CostoCus = (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodCustodia & " ) , " _
        & " Impuesto =  (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodImpuesto & " ) , " _
        & " Right(m.cMovNro, 4) As Usuario " _
        & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro " _
        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = mc.cCtaCod " _
        & " left join relcuentas r on r.cctacod=mc.cctacod " _
        & " WHERE m.cOpecod " & lsOperaciones & " " _
        & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "' " _
        & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' "

        If psUsuarioRep <> "<Consolidado>" Then         'Todos los Usuarios
            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
        End If
        If Len(Trim(psBovedasSelect)) > 0 Then
            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
        End If
        lsSQL = lsSQL & " ORDER BY mc.cCtaCod "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS REGISTRADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotPrestamo = lnTotPrestamo + !Prestamo
                lnTotCostoTas = lnTotCostoTas + !CostoTas
                lnTotCostoCus = lnTotCostoCus + !CostoCus
                lnTotImp = lnTotImp + !Impuesto
                
                'Cadena de Impresion
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!Prestamo, 10, 2) & Space(2) _
                    & ImpreFormat(!CostoTas, 10, 2) & Space(2) & ImpreFormat(!CostoCus, 10, 2) & Space(2) & ImpreFormat(!Impuesto, 10, 2) _
                    & Space(2) & ImpreFormat(!Usuario, 6, 0) _
                    & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS REGISTRADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(120, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(44) & ImpreFormat(lnTotPrestamo, 10, 2) & Space(2) _
            & ImpreFormat(lnTotCostoTas, 10, 2) & Space(2) & ImpreFormat(lnTotCostoCus, 10, 2) & Space(2) _
            & ImpreFormat(lnTotImp, 10, 2) & Chr(10)

        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo128001_ContratoRegis = lsCadBuffer

End Function


Public Function nRepo128002_ContratoDesemb(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim lnTotPrestamo As Currency, lnTotInteresAdel As Currency, lnTotCostoTas As Currency, lnTotCostoCus As Currency
Dim lnTotImp As Currency, lnTotPagado As Currency

lsOperaciones = " in ('" & gColPOpeDesembolso & "','" & gColPOpeDesembolsoEFE & "') "
'On Error GoTo dError

    lsSQL = " SELECT codigoant=isnull(r.cctacodant,'                  '),mc.cCtaCod, m.cMovNro,  mc.cOpeCod, mc.nPlazo, c.nMontoCol, " _
        & " InteresAdel = (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodInteresCompensatorio & " ) , " _
        & " CostoTas = (Select ISNULL(SUM(nMonto),0) From Movcoldet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodTasacion & " ) , " _
        & " CostoCus = (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodCustodia & " ) , " _
        & " Impuesto =  (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodImpuesto & " ) , " _
        & " NetoPagado =  (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodMontoEntregar & " ) , " _
        & " Right(m.cMovNro, 4) As Usuario " _
        & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro " _
        & " INNER JOIN Colocaciones c On  c.cCtaCod = mc.cCtaCod " _
        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = c.cCtaCod " _
        & " LEFT JOIN RELCUENTAS R ON R.CCTACOD=mc.CCTACOD " _
        & " WHERE m.cOpecod " & lsOperaciones & "  " _
        & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "' " _
        & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' "

        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
        End If
        If Len(Trim(psBovedasSelect)) > 0 Then
            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
        End If
        lsSQL = lsSQL & " ORDER BY mc.cCtaCod "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS DESEMBOLSADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 138, "", pnOpeCod)

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
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!nMontoCol, 10, 2) & Space(2) _
                    & ImpreFormat(!InteresAdel, 10, 2) & Space(2) & ImpreFormat(!CostoTas, 10, 2) & Space(2) & ImpreFormat(!CostoCus, 10, 2) & Space(2) & ImpreFormat(!Impuesto, 10, 2) _
                    & Space(2) & ImpreFormat(!NetoPagado, 10, 2) & Space(2) & ImpreFormat(!Usuario, 6, 0) _
                    & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS DESEMBOLSADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 138, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(138, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(43) & ImpreFormat(lnTotPrestamo, 10, 2) & Space(2) _
            & ImpreFormat(lnTotInteresAdel, 10, 2) & Space(2) & ImpreFormat(lnTotCostoTas, 10, 2) & Space(2) & ImpreFormat(lnTotCostoCus, 10, 2) & Space(2) _
            & ImpreFormat(lnTotImp, 10, 2) & Space(2) & ImpreFormat(lnTotPagado, 10, 2) & Chr(10)

        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo128002_ContratoDesemb = lsCadBuffer

End Function


Public Function nRepo128003_ContratoAnula(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lnTotPrestamo As Currency

lsOperaciones = " = '" & gColPOpeAnula & "' "
'On Error GoTo dError

    lsSQL = " SELECT codigoant=isnull(r.cctacodant,'                  '),mc.cCtaCod, m.cMovNro,  mc.cOpeCod, c.nMontoCol,  " _
        & " NomCliente = (Select ISNULL(MAX(Per.cPersNombre),'') From ProductoPersona PP INNER JOIN Persona Per On PP.cPersCod = Per.cPersCod Where PP.cCtaCod = mc.cCtaCod ) , " _
        & " Right(m.cMovNro, 4) As Usuario " _
        & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro " _
        & " INNER JOIN Colocaciones c On  c.cCtaCod = mc.cCtaCod " _
        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = mc.cCtaCod " _
        & " LEFT JOIN RELCUENTAS R ON R.CCTACOD=mc.CCTACOD " _
        & " WHERE m.cOpecod " & lsOperaciones & "  " _
        & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "' " _
        & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' "

        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
        End If
        If Len(Trim(psBovedasSelect)) > 0 Then
            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
        End If
        lsSQL = lsSQL & " ORDER BY mc.cCtaCod "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS ANULADOS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotPrestamo = lnTotPrestamo + !nMontoCol
                
                'Cadena de Impresion
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!NomCliente, 35, 0) & Space(2) _
                    & ImpreFormat(!nMontoCol, 10, 2) & Space(2) & ImpreFormat(!Usuario, 6, 0) _
 _
                    & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS REGISTRADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(120, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(62) & ImpreFormat(lnTotPrestamo, 10, 2) & Chr(10)

        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo128003_ContratoAnula = lsCadBuffer

End Function

Public Function nRepo128004_ContratoRenov(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim lnTotPrestamo As Currency, lnTotPagado As Currency, lnTotCapital As Currency, lnTotInteres As Currency
Dim lnTotImp As Currency, lnTotCostoCus As Currency, lnTotCostoPrepRemate As Currency

lsOperaciones = " in ('" & gColPOpeRenovNorEFE & "','" & gColPOpeRenovMorEFE & "','" & gColPOpeRenovNorCHQ & "','" _
                & gColPOpeRenovMorCHQ & "','" & gColPOpeRenovNorEnOtCjEFE & "','" & gColPOpeRenovMorEnOtCjEFE & "' ) "
'On Error GoTo dError

    lsSQL = " SELECT codigoant=isnull(r.cctacodant,'                  '),mc.cCtaCod, m.cMovNro,  mc.cOpeCod, mc.nPlazo, c.nMontoCol, mc.nMonto as MontoPag, mc.nSaldoCap as NewSaldoCap,  " _
        & " Capital = (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodCapital & " ) , " _
        & " Interes = (Select ISNULL(SUM(nMonto),0) From Movcoldet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod in ( " & gColPConceptoCodInteresMoratorio & "," & gColPConceptoCodTasacion & ")  ) , " _
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
        & " LEFT JOIN RELCUENTAS R ON R.CCTACOD=mC.CCTACOD " _
        & " WHERE mc.cOpecod " & lsOperaciones & " and mc.copecod not like '99%' " _
        & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "' " _
        & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' "

        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
        End If
        If Len(Trim(psBovedasSelect)) > 0 Then
            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
        End If
        lsSQL = lsSQL & " ORDER BY mc.cCtaCod "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS RENOVADOS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 147, "", pnOpeCod)

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
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!nMontoCol, 8, 2) & Space(1) _
                    & ImpreFormat(!MontoPag, 8, 2) & Space(1) & ImpreFormat(!Capital, 8, 2) & Space(1) & ImpreFormat(!Interes, 8, 2) & Space(1) _
                    & ImpreFormat(!Impuesto, 8, 2) & Space(1) & ImpreFormat(!CostoCus, 8, 2) & Space(1) & ImpreFormat(!CostoPrepRemate, 8, 2) _
                    & Space(1) & ImpreFormat(!NewSaldoCap, 8, 2) & Space(1) & ImpreFormat(!Usuario, 4, 0) _
                    & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS RENOVADOS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 147, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(147, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(44) & ImpreFormat(lnTotPrestamo, 10, 2) & Space(1) _
            & ImpreFormat(lnTotPagado, 10, 2) & Space(1) & ImpreFormat(lnTotCapital, 10, 2) & Space(1) & ImpreFormat(lnTotInteres, 10, 2) & Space(1) _
            & ImpreFormat(lnTotImp, 10, 2) & Space(1) & ImpreFormat(lnTotCostoCus, 10, 2) & Space(1) & ImpreFormat(lnTotCostoPrepRemate, 10, 2) _
            & Chr(10)
        
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo128004_ContratoRenov = lsCadBuffer

End Function



Public Function nRepo128005_ContratoCancel(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim lnTotPrestamo As Currency, lnTotPagado As Currency, lnTotCapital As Currency, lnTotInteres As Currency
Dim lnTotImp As Currency, lnTotCostoCus As Currency, lnTotCostoPrepRemate As Currency

lsOperaciones = " in ('" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" & gColPOpeCancelNorCHQ & "','" _
                & gColPOpeCancelMorCHQ & "','" & gColPOpeCanceNorEnOtCjEFE & "','" & gColPOpeCanceMorEnOtCjEFE & "' ) "
'On Error GoTo dError

    lsSQL = " SELECT codigoant=isnull(r.cctacodant,'                  '),mc.cCtaCod, m.cMovNro,  mc.cOpeCod, mc.nPlazo, c.nMontoCol, mc.nMonto as MontoPag, mc.nSaldoCap as NewSaldoCap,  " _
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

        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
        End If
        If Len(Trim(psBovedasSelect)) > 0 Then
            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
        End If
        lsSQL = lsSQL & " ORDER BY mc.cCtaCod "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS CANCELADOS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 145, "", pnOpeCod)

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
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 4, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!nMontoCol, 8, 2) & Space(1) _
                    & ImpreFormat(!MontoPag, 8, 2) & Space(1) & ImpreFormat(!Capital, 8, 2) & Space(1) & ImpreFormat(!Interes, 8, 2) & Space(1) _
                    & ImpreFormat(!Impuesto, 8, 2) & Space(1) & ImpreFormat(!CostoCus, 8, 2) & Space(1) & ImpreFormat(!CostoPrepRemate, 8, 2) _
                    & Space(1) & ImpreFormat(!NewSaldoCap, 8, 2) & Space(1) & ImpreFormat(!Usuario, 4, 0) _
                    & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS CANCELADOS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 145, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(43) & ImpreFormat(lnTotPrestamo, 8, 2) & Space(1) _
            & ImpreFormat(lnTotPagado, 8, 2) & Space(1) & ImpreFormat(lnTotCapital, 8, 2) & Space(1) & ImpreFormat(lnTotInteres, 8, 2) & Space(1) _
            & ImpreFormat(lnTotImp, 8, 2) & Space(1) & ImpreFormat(lnTotCostoCus, 8, 2) & Space(1) & ImpreFormat(lnTotCostoPrepRemate, 8, 2) _
            & Chr(10)
        
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo128005_ContratoCancel = lsCadBuffer

End Function


Public Function nRepo128008_PagoSobrantes(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lnTotSobrante As Currency

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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
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
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!NomCliente, 35, 0) & Space(2) _
                    & ImpreFormat(!nMontoCol, 10, 2) & Space(2) & ImpreFormat(!Usuario, 6, 0) _
                    & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE PAGOS DE SOBRANTES", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(120, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(62) & ImpreFormat(lnTotSobrante, 10, 2) & Chr(10)

        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo128008_PagoSobrantes = lsCadBuffer

End Function



'Public Function nRepo128008_PagoSobrantes(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
'        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
'        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String
'
'Dim lsSQL As String
'Dim lrDataRep As ADODB.Recordset
'Dim loDataRep As dColPFunciones
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
'    Set loDataRep = New dColPFunciones
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
'                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!NomCliente, 35, 0) & Space(2) _
'                    & ImpreFormat(!nMontoCol, 10, 2) & Space(2) & ImpreFormat(!Usuario, 6, 0) _
'                    & Chr(10)
'
'                If lnIndice Mod 300 = 0 Then
'                    lsCadBuffer = lsCadBuffer & lsCadImp
'                    lsCadImp = ""
'                End If
'
'                If lnLineas >= 55 Then
'                    lnPage = lnPage + 1
'                    lsCadImp = lsCadImp & Chr(12)
'                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE PAGOS DE SOBRANTES", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)
'                    lnLineas = 7
'                End If
'                .MoveNext
'            Loop
'        End With
'
'        lsCadImp = lsCadImp & String(120, "-") & Chr(10)
'        lsCadImp = lsCadImp & Space(62) & ImpreFormat(lnTotSobrante, 10, 2) & Chr(10)
'
'        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
'
'    End If
'    nRepo128008_PagoSobrantes = lsCadBuffer
'
'End Function

Public Function nRepo128007_ContratoAmortiz(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim lnTotPrestamo As Currency, lnTotInteresAdel As Currency, lnTotCostoTas As Currency, lnTotCostoCus As Currency
Dim lnTotImp As Currency, lnTotPagado As Currency

lsOperaciones = " in ('" & gColPOpeAmortNorEFE & "','" & gColPOpeAmortNorCHQ & "','" & gColPOpeAmortNorEnOtCjEFE & "' ) "
'On Error GoTo dError

     lsSQL = " SELECT codigoant=isnull(r.cctacodant,'                  '), mc.cCtaCod, m.cMovNro,  mc.cOpeCod, mc.nPlazo, c.nMontoCol, mc.nMonto,  " _
        & " InteresAdel = (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodInteresCompensatorio & " ) , " _
        & " CostoTas = (Select ISNULL(SUM(nMonto),0) From Movcoldet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodTasacion & " ) , " _
        & " CostoCus = (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodCustodia & " ) , " _
        & " Impuesto =  (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
        & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodImpuesto & " ) , " _
        & " " _
        & " " _
        & " Right(m.cMovNro, 4) As Usuario " _
        & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro " _
        & " INNER JOIN Colocaciones c On  c.cCtaCod = mc.cCtaCod " _
        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = c.cCtaCod " _
        & " LEFT JOIN RELCUENTAS R ON R.CCTACOD=mc.CCTACOD " _
        & " WHERE m.cOpecod " & lsOperaciones & "  " _
        & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "' " _
        & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' "

        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
        End If
        If Len(Trim(psBovedasSelect)) > 0 Then
            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
        End If
        lsSQL = lsSQL & " ORDER BY mc.cCtaCod "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CREDITOS PIGNORATICIOS AMORTIZADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
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
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!nMontoCol, 10, 2) & Space(2) _
                    & ImpreFormat(!InteresAdel, 10, 2) & Space(2) & ImpreFormat(!CostoTas, 10, 2) & Space(2) & ImpreFormat(!CostoCus, 10, 2) & Space(2) & ImpreFormat(!Impuesto, 10, 2) _
                    & Space(2) & ImpreFormat(!nMonto, 10, 2) & Space(2) & ImpreFormat(!Usuario, 6, 0) _
                    & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS PIGNORATICIOS AMORTIZADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(150, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(44) & ImpreFormat(lnTotPrestamo, 10, 2) & Space(2) _
            & ImpreFormat(lnTotInteresAdel, 10, 2) & Space(2) & ImpreFormat(lnTotCostoTas, 10, 2) & Space(2) & ImpreFormat(lnTotCostoCus, 10, 2) & Space(2) _
            & ImpreFormat(lnTotImp, 10, 2) & Space(2) & ImpreFormat(lnTotPagado, 10, 2) & Chr(10)

        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo128007_ContratoAmortiz = lsCadBuffer

End Function

Public Function nRepo128007_ContratoRecup(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
       ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim lnTotPrestamo As Currency, lnTotDeudaVenc As Currency, lnTotIntMoratorio As Currency, lnTotItf As Currency
Dim lnTotImp As Currency, lnTotPagado As Currency

lsOperaciones = "='122900' "
'On Error GoTo dError

     lsSQL = " SELECT codigoant=isnull(r.cctacodant,'                  '), mc.cCtaCod, m.cMovNro,  mc.cOpeCod, mc.nPlazo, c.nMontoCol, mc.nMonto,  " _
           & " DeudaVenc = (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
           & "                 and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodCapitalVencido & " ) , " _
           & " IntMoratorio = (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
           & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodInteresMoratorio & " ) , " _
           & " Itf = (Select ISNULL(SUM(nMonto),0) From Movcoldet Where nMovNro = mc.nMovNro and cOpeCod ='" & gITFCobroEfectivoPrend & "'" _
           & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen  ) , " _
           & " Right(m.cMovNro, 4) As Usuario " _
           & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro " _
           & " INNER JOIN Colocaciones c On  c.cCtaCod = mc.cCtaCod " _
           & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = c.cCtaCod " _
           & " LEFT JOIN RELCUENTAS R ON R.CCTACOD=mc.CCTACOD " _
           & " WHERE mc.cOpecod " & lsOperaciones & "  " _
           & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "' " _
           & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' "

        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
        End If
        If Len(Trim(psBovedasSelect)) > 0 Then
            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
        End If
        lsSQL = lsSQL & " ORDER BY mc.cCtaCod "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("CREDITOS PIGNORATICIOS ADJUDICADOS RECUPERADOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
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
                lnTotItf = lnTotItf + !Itf
                lnTotPagado = lnTotPagado + !nMonto
                'Cadena de Impresion
                '"Item   ContratoAnt            Contrato            Prestamo    Monto Pago      DeudaVencida  Interes   Itf        Usuario" & Chr(10)
                
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!nMontoCol, 10, 2) & Space(2) _
                    & ImpreFormat(!nMonto, 10, 2) & Space(2) & ImpreFormat(!DeudaVenc, 10, 2) & Space(2) & ImpreFormat(!IntMoratorio, 10, 2) & Space(2) & ImpreFormat(!Itf, 10, 2) _
                    & Space(7) & ImpreFormat(!Usuario, 6, 0) _
                    & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS PIGNORATICIOS ADJUDICADOS RECUPERADOS", "DEL " & Format(pdDesde, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(150, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(44) & ImpreFormat(lnTotPrestamo, 10, 2) & Space(2) _
                            & ImpreFormat(lnTotPagado, 10, 2) & Space(2) & ImpreFormat(lnTotDeudaVenc, 10, 2) & Space(2) & ImpreFormat(lnTotIntMoratorio, 10, 2) & Space(2) _
                            & ImpreFormat(lnTotItf, 10, 2) & Space(2) & Chr(10)

        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo128007_ContratoRecup = lsCadBuffer

End Function

Public Function nRepo128009_ContratoVend(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
       ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim lnTotPrestamo As Currency, lnTotDeudaVenc As Currency, lnTotIntMoratorio As Currency, lnTotIGV As Currency
Dim lnTotImp As Currency, lnTotPagado As Currency

lsOperaciones = "='" & gColPOpeVtaSubasta & "'"
'On Error GoTo dError

     lsSQL = " SELECT codigoant=isnull(r.cctacodant,'                  '), mc.cCtaCod, m.cMovNro,  mc.cOpeCod, mc.nPlazo, c.nMontoCol, mc.nMonto,  " _
           & " Deuda = (Select case when SUM(MD.nMonto)=0 or SUM(MD.nMonto) is null  then sum(p.nsaldo) else SUM(MD.nMonto) end From MovColDet MD JOIN PRODUCTO P ON P.CCTACOD=MD.CCTACOD Where MD.nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
           & "                 and MD.cCtaCod = mc.cCtaCod and MD.nNroCalen = mc.nNroCalen and MD.nPrdConceptoCod = " & gColPConceptoCodValorAdjudica & " ) , " _
           & " IGV = (Select ISNULL(SUM(nMonto),0) From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod " _
           & "              and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodImpuestoVtaSubasta & " ) , " _
           & " Right(m.cMovNro, 4) As Usuario " _
           & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro " _
           & " INNER JOIN Colocaciones c On  c.cCtaCod = mc.cCtaCod " _
           & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = c.cCtaCod " _
           & " LEFT JOIN RELCUENTAS R ON R.CCTACOD=mc.CCTACOD " _
           & " WHERE mc.cOpecod " & lsOperaciones & "  " _
           & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "' " _
           & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' "

        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
        End If
        If Len(Trim(psBovedasSelect)) > 0 Then
            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
        End If
        lsSQL = lsSQL & " ORDER BY mc.cCtaCod "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
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
                '"Item   ContratoAnt            Contrato            Prestamo    Monto Pago      DeudaVencida  Interes   Itf        Usuario" & Chr(10)
                
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!nMontoCol, 10, 2) & Space(2) _
                    & ImpreFormat(!nMonto, 10, 2) & Space(2) & ImpreFormat(!Deuda, 10, 2) & Space(2) & ImpreFormat(!IGV, 10, 2) & Space(4) _
                    & ImpreFormat(!Usuario, 6, 0) _
                    & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("CREDITOS PIGNORATICIOS ADJUDICADOS VENDIDOS ", "DEL " & Format(pdDesde, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(150, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(44) & ImpreFormat(lnTotPrestamo, 10, 2) & Space(2) _
                            & ImpreFormat(lnTotPagado, 10, 2) & Space(2) & ImpreFormat(lnTotDeudaVenc, 10, 2) & Space(2) _
                            & ImpreFormat(lnTotIGV, 10, 2) & Space(2) & Chr(10)

        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo128009_ContratoVend = lsCadBuffer

End Function






Public Function nRepo12802_MovimientoJoyas(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lsTitulo As String

Dim lnTotValTasac As Currency, lnTotOroNeto As Currency
Dim lnTotPiezas As Long

Select Case pnOpeCod
    Case 128021  ' Prendas Nuevas
        lsOperaciones = " = '" & gColPOpeRegContrato & "' "
        lsTitulo = "LISTADO DE PRENDAS NUEVAS"
    Case 128022  ' Prendas Rescatadas
        lsOperaciones = " = '" & gColPOpeDevJoyas & "' "
        lsTitulo = "LISTADO DE PRENDAS RESCATADAS"
    Case 128023  ' Prendas Nuevas en condicion de Diferidas
        lsOperaciones = " in ('" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" _
                & gColPOpeCancelNorCHQ & "','" & gColPOpeCancelMorCHQ & "' ) "
        lsTitulo = "LISTADO DE PRENDAS NUEVAS EN CONDICION DE DIFERIDAS"
        
    Case 128024   'Prendas Recuperadas en condición de Adjudicadas
        lsOperaciones = " = '122900' "
        lsTitulo = "LISTADO DE PRENDAS RECUPERADAS EN CONDICION DE ADJUDICADAS"
        
    Case 128025   'Prendas Recuperadas en condición de Adjudicadas
        lsOperaciones = " = '" & gColPOpeVtaSubasta & "'"
        lsTitulo = "LISTADO DE PRENDAS VENDIDAS EN CONDICION DE ADJUDICADAS"
        
End Select
'On Error GoTo dError

    lsSQL = " SELECT codigoant=isnull(R.cctacodAnt,'                  '),mc.cCtaCod, m.cMovNro,  mc.cOpeCod, cp.cAgeBoveda, " _
        & " cp.nOroNeto, cp.nPiezas, cp.nTasacion, " _
        & " Right(m.cMovNro, 4) As Usuario " _
        & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro " _
        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = mc.cCtaCod " _
        & " LEFT JOIN RELCUENTAS R ON R.CCTACOD=mc.CCTACOD " _
        & " WHERE m.nmovflag=0 and m.cOpecod " & lsOperaciones & "  and mc.copecod not like '99%'  " _
        & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND  '" & Format(pdHasta, "yyyymmdd") & "' " _
        & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' "

        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
        End If
        If Len(Trim(psBovedasSelect)) > 0 Then
            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
        End If
        lsSQL = lsSQL & " ORDER BY mc.cCtaCod "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera(lsTitulo, "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotValTasac = lnTotValTasac + !nTasacion
                lnTotPiezas = lnTotPiezas + !npiezas
                lnTotOroNeto = lnTotOroNeto + !noroneto
                
                If pnOpeCod = "128021" Then
                'Cadena de Impresion
                    lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!nTasacion, 10, 2) & Space(1) _
                        & ImpreFormat(!npiezas, 5, 0) & Space(1) & ImpreFormat(!noroneto, 6, 2) & Space(4) & ImpreFormat(!cAgeBoveda, 5, 0) _
                        & Space(2) & ImpreFormat(!Usuario, 6, 0) _
                        & Chr(10)
                Else
                    lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!nTasacion, 10, 2) & Space(1) _
                        & ImpreFormat(!npiezas, 5, 0) & Space(1) & ImpreFormat(!noroneto, 6, 2) & Space(4) & ImpreFormat(!cAgeBoveda, 5, 0) _
                        & Space(2) & ImpreFormat(!Usuario, 6, 0) _
                        & Chr(10)
                    
                End If
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera(lsTitulo, "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        lsCadImp = lsCadImp & String(120, "-") & Chr(10)
        If pnOpeCod <> "128021" Then
        
            lsCadImp = lsCadImp & Space(44) & ImpreFormat(lnTotValTasac, 10, 2) & Space(1) _
                                        & ImpreFormat(lnTotPiezas, 5, 0) & Space(1) & ImpreFormat(lnTotOroNeto, 6, 2) & Space(1) _
                                        & Chr(10)
        Else
            lsCadImp = lsCadImp & Space(25) & ImpreFormat(lnTotValTasac, 10, 2) & Space(1) _
                                            & ImpreFormat(lnTotPiezas, 5, 0) & Space(1) & ImpreFormat(lnTotOroNeto, 6, 2) & Space(1) _
                                            & Chr(10)
        End If
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo12802_MovimientoJoyas = lsCadBuffer

End Function

Public Function nRepo128031_ListadoCredVigentes(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal psRenovado As String, ByVal psEstado As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsRenovad As String
Dim lsEstados As String
Dim lsCondiciones  As String ' Cabecera Impresion

Dim lnTotPrestamo As Currency, lnTotSaldo As Currency, lnTotTasacion As Currency, lnTotOroNeto  As Currency

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

    lsSQL = " SELECT codigoant=isnull(r.cctacodant,'                  '),P.cCtaCod, c.dVigencia, c.nPlazo, c.nMontoCol, P.nSaldo, " _
        & " cp.nTasacion, cp.nOroNeto, cp.cAgeBoveda,  p.nPrdEstado, Substring(Cons1.cConsDescripcion,1,10) as DescEstado,  " _
        & " NomCliente = (Select ISNULL(MAX(Per.cPersNombre),'') From ProductoPersona PP INNER JOIN Persona Per " _
        & "                         On PP.cPersCod = Per.cPersCod Where PP.cCtaCod = c.cCtaCod )  " _
        & " FROM Producto P INNER JOIN Colocaciones c On  p.cCtaCod = c.cCtaCod " _
        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = c.cCtaCod " _
        & " INNER JOIN Constante Cons1 On P.nPrdEstado = Cons1.nConsValor " _
        & " LEFT join relcuentas r on r.cctacod=p.cctacod    " _
        & " WHERE P.nPrdEstado in ( " & lsEstados & " )  " _
        & " AND Substring (p.cCtaCod, 4,2) = '" & psAgencia & "' " _
        & " AND Cons1.nConsCod = " & gColocEstado
        If Len(Trim(psBovedasSelect)) > 0 Then
            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
        End If
        lsSQL = lsSQL & lsRenovad   ' Renovaciones
        lsSQL = lsSQL & " ORDER BY p.cCtaCod "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS", lsCondiciones, lnPage, 155, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotPrestamo = lnTotPrestamo + !nMontoCol
                lnTotSaldo = lnTotSaldo + !nSaldo
                lnTotTasacion = lnTotTasacion + !nTasacion
                lnTotOroNeto = lnTotOroNeto + !noroneto

                'Cadena de Impresion
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!NomCliente, 35, 0) & Space(1) _
                    & Format(!dVigencia, "dd/mm/yyyy") & Space(1) & ImpreFormat(!nPlazo, 3, 0) & Space(1) & ImpreFormat(!nMontoCol, 10, 2) & Space(1) _
                    & ImpreFormat(!nSaldo, 10, 2) & Space(1) & ImpreFormat(!nTasacion, 10, 2) & Space(1) & ImpreFormat(!noroneto, 5, 2) _
                    & Space(1) & ImpreFormat(!DescEstado, 10, 0) & Space(1) _
                    & Chr(10)
                    
                    ' ImpreFormat(!cAgeBoveda, 5, 0)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS", lsCondiciones, lnPage, 155, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(76) & ImpreFormat(lnTotPrestamo, 10, 2) & Space(1) _
            & ImpreFormat(lnTotSaldo, 10, 2) & Space(1) & ImpreFormat(lnTotTasacion, 10, 2) & Space(1) & ImpreFormat(lnTotOroNeto, 5, 2) & Space(1) _
            & Chr(10)
        
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo128031_ListadoCredVigentes = lsCadBuffer

End Function

Public Function nRepo128032_ListadoCredDiferidos(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsEstados As String
Dim lsCondiciones  As String ' Cabecera Impresion

Dim lnTotTasacion As Currency, lnTotOroNeto  As Currency
Dim lsFecCancel As String

lsEstados = gColPEstDifer
lsCondiciones = ""
'On Error GoTo dError

    lsSQL = " SELECT codigoant=isnull(R.cctacodAnt,'                  '),P.cCtaCod, c.dVigencia, c.nPlazo, c.nMontoCol, P.nSaldo, " _
        & " cp.nTasacion, cp.nOroNeto, cp.cAgeBoveda,  p.nPrdEstado, Cons1.cConsDescripcion as DescEstado,  " _
        & " NomCliente = (Select ISNULL(MAX(Per.cPersNombre),'') From ProductoPersona PP INNER JOIN Persona Per " _
        & "                         On PP.cPersCod = Per.cPersCod Where PP.cCtaCod = c.cCtaCod ) ,  " _
        & " dCancelado = (SELECT Max(CONVERT(DATETIME, SUBSTRING(CMOVNRO,1,8),  103 ))  " _
        & "            FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
        & "            WHERE MC.cCtaCod = C.cCtaCod AND MC.cOpeCod in ('" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" & gColPOpeCanceNorEnOtCjEFE & "','" & gColPOpeCanceMorEnOtCjEFE & "') " _
        & "                 AND M.nMovFlag <> " & gMovFlagExtornado & " )  " _
        & " FROM Producto P INNER JOIN Colocaciones c On  p.cCtaCod = c.cCtaCod " _
        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = c.cCtaCod " _
        & " INNER JOIN Constante Cons1 On P.nPrdEstado = Cons1.nConsValor " _
        & " left join relcuentas r on r.cctacod=p.cctacod " _
        & " WHERE P.nPrdEstado in ( " & lsEstados & " )  " _
        & " AND Substring (p.cCtaCod, 4,2) = '" & psAgencia & "' " _
        & " AND Cons1.nConsCod = " & gColocEstado
        If Len(Trim(psBovedasSelect)) > 0 Then
            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
        End If
        lsSQL = lsSQL & " ORDER BY p.cCtaCod "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS EN CONDICION DE DIFERIDOS", lsCondiciones, lnPage, 145, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        Dim lsAux As String
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotTasacion = lnTotTasacion + !nTasacion
                lnTotOroNeto = lnTotOroNeto + !noroneto
                lsFecCancel = IIf(IsNull(!dCancelado), "  /  /    ", Format(!dCancelado, "dd/mm/yyyy"))
                
                If IsDate(lsFecCancel) Then
                    lsAux = ImpreFormat(DateDiff("d", Format(lsFecCancel, "dd/mm/yyyy"), Format(csFechaSis, "dd/mm/yyyy")), 5, 0)
                End If

'ImpreFormat(!nPlazo, 3, 0) & Space(1) &
                'Cadena de Impresion
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!NomCliente, 35, 0) & Space(1) _
                    & ImpreFormat(!nTasacion, 10, 2) & Space(1) & ImpreFormat(!noroneto, 5, 2) & Space(3) _
                    & lsFecCancel & Space(1) _
                    & IIf(IsDate(lsFecCancel), lsAux, "") _
                    & Space(3) & ImpreFormat(!cAgeBoveda, 5, 0) & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS EN CONDICION DE DIFERIDOS", lsCondiciones, lnPage, 145, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(65) & ImpreFormat(lnTotTasacion, 10, 2) & Space(1) _
            & ImpreFormat(lnTotOroNeto, 5, 2) & Space(1) _
            & Chr(10)
        
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo128032_ListadoCredDiferidos = lsCadBuffer

End Function

Public Function nRepo128034_ListadoCredAdjudicados(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsEstados As String
Dim lsCondiciones  As String ' Cabecera Impresion

Dim lnTotTasacion As Currency, lnTotOroNeto  As Currency
Dim lsFecCancel As String

lsEstados = gColPEstAdjud
lsCondiciones = ""
'On Error GoTo dError

    lsSQL = " SELECT codigoant=isnull(R.cctacodAnt,'                  '),P.cCtaCod, c.dVigencia, c.nPlazo, c.nMontoCol, P.nSaldo, " _
        & " cp.nTasacion, cp.nOroNeto, cp.cAgeBoveda,  p.nPrdEstado, Cons1.cConsDescripcion as DescEstado,  " _
        & " NomCliente = (Select ISNULL(MAX(Per.cPersNombre),'') From ProductoPersona PP INNER JOIN Persona Per " _
        & "                         On PP.cPersCod = Per.cPersCod Where PP.cCtaCod = c.cCtaCod ) ,  " _
        & " dCancelado = p.dprdestado  " _
        & " FROM Producto P INNER JOIN Colocaciones c On  p.cCtaCod = c.cCtaCod " _
        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = c.cCtaCod " _
        & " INNER JOIN Constante Cons1 On P.nPrdEstado = Cons1.nConsValor " _
        & " left join relcuentas r on r.cctacod=p.cctacod " _
        & " WHERE P.nPrdEstado = " & lsEstados & "" _
        & " AND Substring (p.cCtaCod, 4,2) = '" & psAgencia & "' " _
        & " AND Cons1.nConsCod = " & gColocEstado
        If Len(Trim(psBovedasSelect)) > 0 Then
            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
        End If
        lsSQL = lsSQL & " ORDER BY p.cCtaCod "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS EN CONDICION DE ADJUDICADOS", lsCondiciones, lnPage, 145, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        Dim lsAux As String
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                'Asigno los totales
                lnTotTasacion = lnTotTasacion + !nTasacion
                lnTotOroNeto = lnTotOroNeto + !noroneto
                lsFecCancel = IIf(IsNull(!dCancelado), "  /  /    ", Format(!dCancelado, "dd/mm/yyyy"))
                
                If IsDate(lsFecCancel) Then
                    lsAux = ImpreFormat(DateDiff("d", Format(lsFecCancel, "dd/mm/yyyy"), Format(csFechaSis, "dd/mm/yyyy")), 5, 0)
                End If

'ImpreFormat(!nPlazo, 3, 0) & Space(1) &
                'Cadena de Impresion
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!NomCliente, 35, 0) & Space(1) _
                    & ImpreFormat(!nTasacion, 10, 2) & Space(1) & ImpreFormat(!noroneto, 5, 2) & Space(3) _
                    & lsFecCancel & Space(1) _
                    & Space(3) & ImpreFormat(!cAgeBoveda, 5, 0) & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS EN CONDICION DE ADJUDICADOS", lsCondiciones, lnPage, 145, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(65) & ImpreFormat(lnTotTasacion, 10, 2) & Space(1) _
            & ImpreFormat(lnTotOroNeto, 5, 2) & Space(1) _
            & Chr(10)
        
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo128034_ListadoCredAdjudicados = lsCadBuffer

End Function



Public Function nRepo128036_Creditos_prendarios_fecha_determinada(ByVal psAgencia As String, ByVal pdFecha As String) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim cCliente As String * 30
Dim lsEstados As String
Dim lsCondiciones  As String ' Cabecera Impresion
Dim nitem As Long

Dim lnTotalSaldo As Double
Dim lnTotalTasacion As Double
Dim lnTotalOroNeto As Double

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
    lsSQL = lsSQL & " where convert(varchar(8), cs.dfecha, 112)='" & Format(pdFecha, "YYYYMMdd") & "' "
    lsSQL = lsSQL & " and cs.nprdestado in (" & lsEstados & ") " '2101,2104,2106,2107) "
    lsSQL = lsSQL & " and SUBSTRING(P.cCtaCod, 4, 2)='" & psAgencia & "' "
    lsSQL = lsSQL & "  ORDER BY SUBSTRING(P.cCtaCod, 4, 2), SUBSTRING(P.cCtaCod, 9, 1), SUBSTRING(P.cCtaCod, 6, 3) "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
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
                
                nitem = nitem + 1
                
                cCliente = lrDataRep!cCliente
                
                lnTotalSaldo = lnTotalSaldo + lrDataRep!nSaldoCap
                lnTotalTasacion = lnTotalTasacion + lrDataRep!nTasacion
                lnTotalOroNeto = lnTotalOroNeto + lrDataRep!noroneto

                'Cadena de Impresion
                lsCadImp = lsCadImp & ImpreFormat(nitem, 4, 0) & Space(1) & lrDataRep!cCtaCod & Space(1) _
                                    & cCliente & Space(1) & ImpreFormat(lrDataRep!nSaldoCap, 10, 2) _
                                    & Space(1) & ImpreFormat(lrDataRep!nDiasAtraso, 8, 0) _
                                    & Space(1) & ImpreFormat(lrDataRep!nTasacion, 10, 2) _
                                    & Space(1) & ImpreFormat(lrDataRep!noroneto, 6, 2) & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS PRENDARIOS PARA FECHA DETERMINADA - Agencia " & NombreAgencia(psAgencia), lsCondiciones, lnPage, 145, "", 128036)
                    lnLineas = 7
                    
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(55) & ImpreFormat(lnTotalSaldo, 10, 2) & Space(10) _
            & ImpreFormat(lnTotalTasacion, 10, 2) & Space(1) & ImpreFormat(lnTotalOroNeto, 6, 2) & Chr(10)
        
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo128036_Creditos_prendarios_fecha_determinada = lsCadBuffer

End Function

Public Function NombreAgencia(ByVal lsAge As String) As String
Dim rs As New ADODB.Recordset
Dim sql As String
Dim oCon As DConecta
Set oCon = New DConecta
oCon.AbreConexion

sql = "select cAgeDescripcion from Agencias where cAgeCod='" & lsAge & "'"

Set rs = oCon.CargaRecordSet(sql)
oCon.CierraConexion
NombreAgencia = IIf(IsNull(rs!cAgeDescripcion), "", rs!cAgeDescripcion)
End Function

'''Public Function nRepo128041_Estadistica_Saldos(ByVal pdIni As Date, ByVal pdFin As Date, _
'''ByVal psCodAge As String, ByVal pdFecSis As Date, ByVal psNomCmac As String, _
'''ByVal psNomAge As String) As String
''''On Error GoTo ControlError
'''Dim pCadena As String
'''Dim vSaldo As Currency
'''Dim vNumero As Currency
'''Dim vLineas As Integer
'''Dim vPage As Integer
'''Dim lsSQL As String
'''Dim lrReg As New ADODB.Recordset
'''Dim lsPrestamos As String, lsCancelacion As String, lsRenovacion As String
'''Dim paso1 As Boolean
'''Dim Co As DConecta
'''Dim gdHoraGrab As String
'''Dim n1 As Integer
'''Dim n2 As Integer
'''Dim n3 As Integer
'''Dim sum1 As Double
'''Dim sum2 As Double
'''Dim sum3 As Double
'''
'''n1 = 0
'''n2 = 0
'''n3 = 0
'''sum1 = 0
'''sum2 = 0
'''sum3 = 0
'''
'''
'''Set Co = New DConecta
'''
'''If (DateDiff("d", pdIni, pdFin) + 1) > 32650 Then
'''    MsgBox " Rango es demasiado amplio ", vbInformation, " Aviso "
'''    Exit Function
'''End If
'''
'''lsSQL = " select dEstad Fecha,"
'''lsSQL = lsSQL & " sum (nNumPres) NroPres,"
'''lsSQL = lsSQL & " sum (nMonPres) MonPres,"
'''lsSQL = lsSQL & " sum (nNumCanc) NroCanc,"
'''lsSQL = lsSQL & " sum (nMonCanc) MonCanc,"
'''lsSQL = lsSQL & " sum (nNumReno) NroRenov,"
'''lsSQL = lsSQL & " sum (nMonReno) MonRenov"
'''lsSQL = lsSQL & " From ColocEstadDiaPrend"
'''lsSQL = lsSQL & " where datediff(day,dEstad,'" & Format(pdIni, "YYYY/MM/DD") & "')<=0  "
'''lsSQL = lsSQL & " and   datediff(day,dEstad,'" & Format(pdFin, "YYYY/MM/DD") & "')>=0 "
'''lsSQL = lsSQL & " and cCodAge='" & psCodAge & "' "
'''lsSQL = lsSQL & " Group by dEstad"
'''
'''Co.AbreConexion
'''Set lrReg = Co.CargaRecordSet(lsSQL)
'''Co.CierraConexion
'''
'''gdHoraGrab = Format(pdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss")
'''vPage = 1
'''pCadena = pCadena & Chr(10)
'''pCadena = pCadena & ImpreFormat(psNomCmac, 35, 5) & Space(100) & "Página : " & ImpreFormat(vPage, 6, 0) & Chr(10)
'''pCadena = pCadena & ImpreFormat(psNomAge, 35, 5) & Space(99) & Format(pdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & Chr(10)
''''pCadena = pCadena & Space(4) & "Bovedas " & psBoveda & Chr(10)
'''pCadena = pCadena & Space(55) & " ESTADISTICA DE SALDOS DEL " & pdIni & " - AL " & pdFin & Chr(10)
'''pCadena = pCadena & Space(55) & String(54, "=") & Chr(10)
'''pCadena = pCadena & Chr(10)
'''pCadena = pCadena & Space(2) & String(120, "-") & Chr(10)
'''pCadena = pCadena & Space(7) & "FECHA        NRO. PREST.NUEVOS  NRO. MONTO.CANCEL.   NRO. MONTO.RENOVA." & Chr(10)
'''pCadena = pCadena & Space(2) & String(120, "-") & Chr(10)
'''
'''vLineas = 13
'''While Not lrReg.EOF
'''        With lrReg
'''            pCadena = pCadena & Space(4) & Format(!Fecha, "DD/MM/YYYY") & ImpreFormat(!NroPres, 8, 0) & ImpreFormat(!MonPres, 12, , True) & ImpreFormat(!NroCanc, 12, , True) & ImpreFormat(!MonCanc, 12, , True) & ImpreFormat(!NroRenov, 6, 0) & ImpreFormat(!MonRenov, 12, , True) & Chr(10)
'''            n1 = !NroPres
'''            n2 = !NroCanc
'''            n3 = !NroRenov
'''            sum1 = !MonPres
'''            sum2 = !MonCanc
'''            sum3 = !MonRenov
'''        End With
'''        If vLineas > 66 Then  'pLineasMax
'''          vPage = vPage + 1
'''          pCadena = pCadena & Chr(12)
'''          pCadena = pCadena & Space(135) & "Página : " & ImpreFormat(vPage, 6, 0) & Chr(10)
'''          pCadena = pCadena & Space(2) & String(120, "-") & Chr(10)
'''          pCadena = pCadena & Space(7) & "FECHA        NRO. PREST.NUEVOS  NRO. MONTO.CANCEL.   NRO. MONTO.RENOVA." & Chr(10)
'''          pCadena = pCadena & Space(2) & String(120, "-") & Chr(10)
'''          vLineas = 5
'''        End If
'''        vLineas = vLineas + 1
'''    lrReg.MoveNext
'''Wend
'''
'''pCadena = pCadena & Space(2) & String(120, "-") & Chr(10)
'''pCadena = pCadena & Space(14) & ImpreFormat(n1, 8, 0) & ImpreFormat(sum1, 12, , True) & ImpreFormat(n2, 6, 0) & ImpreFormat(sum2, 12, , True) & ImpreFormat(n3, 6, 0) & ImpreFormat(sum3, 12, , True) & Chr(10)
'''pCadena = pCadena & Space(2) & String(120, "-") & Chr(10)
'''
''''Devuelve cadena con Impresión
'''nRepo128041_Estadistica_Saldos = pCadena & Chr(12)
'''Set Co = Nothing
'''Set lrReg = Nothing
'''End Function
'''Public Function nRepo128042_Estadistica_x_Lote(ByVal pdIni As Date, _
'''ByVal pdFin As Date, ByVal psBoveda As String, _
'''ByVal psNomAge As String, ByVal pdFecSis As Date) As String
'''
'''On Error GoTo ControlError
'''Dim pCadena As String
'''Dim vVigSal As Currency, vAdjSal As Currency
'''Dim vVigNum As Currency, vAdjNum As Currency
'''Dim x As Integer, Y As Integer
'''Dim vFecIni As Date, vFecFin As Date
'''Dim Arreglo() As Currency
'''Dim vLineas As Integer
'''Dim vPage As Integer
'''Dim gdHoraGrab As Date
'''Erase Arreglo
'''If (DateDiff("d", pdIni, pdFin) + 1) > 32650 Then
'''    MsgBox " Rango es demasiado amplio ", vbInformation, " Aviso "
'''    nRepo128042_Estadistica_x_Lote = ""
'''    Exit Function
'''End If
'''x = DateDiff("d", pdIni, pdFin) + 1
'''If x > 365 Then
'''    If MsgBox("Rango es muy amplio" & vbCr & "Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton1, " Aviso ") = vbNo Then
'''        nRepo128042_Estadistica_x_Lote = ""
'''        Exit Function
'''    End If
'''End If
'''ReDim Arreglo(x, 12)
'''vFecIni = pdIni: vFecFin = pdFin
'''Y = 1
'''vVigNum = SalNroEst(psBoveda, "2100", "2101", "2102", "2104", "2106", "2107")
'''vVigSal = SalSumEst(psBoveda, "nOroNeto", "2100", "2101", "2102", "2104", "2106", "2107")
'''vAdjNum = SalNroEst(psBoveda, "2108")
'''vAdjSal = SalSumEst(psBoveda, "nOroNeto", "2108")
'''
'''
'''
'''Dim paso1 As Boolean
'''paso1 = True
'''Do While DateDiff("d", vFecIni, vFecFin) >= 0
'''    'Vigentes : Ingresos
'''   ' Arreglo(x, 1) = DiaNroEst(psBoveda, vFecFin, 2100) - DiaNroEst(psBoveda, vFecFin, 2110)
'''   ' Arreglo(x, 2) = DiaSumCP(psBoveda, vFecFin, "nOroNeto", 2100) - DiaSumCP(psBoveda, pConexPrend, vFecFin, "nOroNeto", 2110)
'''   'Vigentes : Egresos
'''    Arreglo(x, 3) = DiaNroEst(psBoveda, vFecFin, gsDevJoyas, gsDevJoyasEOA, gsVtaRemate)
'''    Arreglo(x, 4) = DiaSumCP(psBoveda, vFecFin, "nOroNeto", gsDevJoyas, gsDevJoyasEOA, gsVtaRemate, gsVtaRemEnOtAg)
'''    'Adjudicados : Ingresos
'''    Arreglo(x, 5) = DiaNroEst(psBoveda, vFecFin, gsAdjudica)
'''    Arreglo(x, 6) = DiaSumCP(psBoveda, vFecFin, "nOroNeto", gsAdjudica)
'''    'Adjudicados : Egresos
'''    Arreglo(x, 7) = DiaNroEst(psBoveda, vFecFin, gsVtaSubasta)
'''    Arreglo(x, 8) = DiaSumCP(psBoveda, vFecFin, "nOroNeto", gsVtaSubasta, gsVtaSubEnOtAg)
'''
'''    'Saldos Vigentes  -  Adjudicados
'''    If paso1 Then
'''        Arreglo(x, 9) = vVigNum
'''        Arreglo(x, 10) = vVigSal
'''        Arreglo(x, 11) = vAdjNum
'''        Arreglo(x, 12) = vAdjSal
'''        paso1 = False
'''    Else
'''        Arreglo(x, 9) = Abs(Arreglo(x + 1, 9) - Arreglo(x + 1, 1)) + Arreglo(x + 1, 3)
'''        Arreglo(x, 10) = Abs(Arreglo(x + 1, 10) - Arreglo(x + 1, 2)) + Arreglo(x + 1, 4)
'''        Arreglo(x, 11) = Abs(Arreglo(x + 1, 11) - Arreglo(x + 1, 5)) + Arreglo(x + 1, 7)
'''        Arreglo(x, 12) = Abs(Arreglo(x + 1, 12) - Arreglo(x + 1, 6)) + Arreglo(x + 1, 8)
'''    End If
'''    Arreglo(0, 1) = Arreglo(0, 1) + Arreglo(x, 1)
'''    Arreglo(0, 2) = Arreglo(0, 2) + Arreglo(x, 2)
'''    Arreglo(0, 3) = Arreglo(0, 3) + Arreglo(x, 3)
'''    Arreglo(0, 4) = Arreglo(0, 4) + Arreglo(x, 4)
'''    Arreglo(0, 5) = Arreglo(0, 5) + Arreglo(x, 5)
'''    Arreglo(0, 6) = Arreglo(0, 6) + Arreglo(x, 6)
'''    Arreglo(0, 7) = Arreglo(0, 7) + Arreglo(x, 7)
'''    Arreglo(0, 8) = Arreglo(0, 8) + Arreglo(x, 8)
'''
'''    'poProgress.value = y
'''    Y = Y + 1
'''    x = x - 1: vFecFin = DateAdd("d", -1, vFecFin)
'''Loop
'''
''''Carga fecha y hora de grabación
'''gdHoraGrab = Format(pdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss")
'''vPage = 1
'''pCadena = pCadena & Chr(10)
'''pCadena = pCadena & ImpreFormat(gsNomCmac, 35, 5) & Space(85) & "Página : " & ImpreFormat(vPage, 6, 0) & Chr(10)
'''pCadena = pCadena & Space(5) & "Credito Pignoraticio" & Space(100) & Format(pdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & Chr(10)
'''pCadena = pCadena & ImpreFormat(psNomAge, 35, 5) & Space(97) & gsCodUser & Chr(10)
'''pCadena = pCadena & Space(50) & " ESTADISTICA DE LOTES DEL : " & pdIni & " AL " & pdFin & Chr(10)
'''pCadena = pCadena & Space(50) & String(53, "=") & Chr(10)
'''pCadena = pCadena & Space(5) & " Bovedas " & psBoveda & Chr(10)
'''pCadena = pCadena & Space(1) & String(144, "-") & Chr(10)
'''pCadena = pCadena & Space(5) & "       ¦              V  I  G  E  N  T  E  S                 ¦          A  D  J  U  D  I  C  A  D  O  S            ¦      T  O  T  A  L" & Chr(10)
'''pCadena = pCadena & Space(5) & "       ¦     INGRESO         EGRESO           S T O C K      ¦     INGRESO         EGRESO           S T O C K      ¦ S T O C K  B O V E D A" & Chr(10)
'''pCadena = pCadena & Space(5) & "FECHA  ¦ LOTES   GRAMOS  LOTES   GRAMOS    LOTES      GRAMOS ¦ LOTES   GRAMOS  LOTES   GRAMOS    LOTES      GRAMOS ¦    LOTES      GRAMOS" & Chr(10)
'''pCadena = pCadena & Space(1) & String(144, "-") & Chr(10)
'''
'''x = 1
'''vLineas = 13
'''vFecIni = pdIni: vFecFin = pdFin
'''Do While DateDiff("d", vFecIni, vFecFin) >= 0
'''    pCadena = pCadena & Space(2) & vFecIni & "¦" & ImpreFormat(Arreglo(x, 1), 6, 0, True) & ImpreFormat(Arreglo(x, 2), 7, , True) & ImpreFormat(Arreglo(x, 3), 6, 0, True) & ImpreFormat(Arreglo(x, 4), 7, , True) & ImpreFormat(Arreglo(x, 9), 8, 0, True) & ImpreFormat(Arreglo(x, 10), 9, , True) & " ¦" & _
'''                                            ImpreFormat(Arreglo(x, 5), 6, 0, True) & ImpreFormat(Arreglo(x, 6), 7, , True) & ImpreFormat(Arreglo(x, 7), 6, 0, True) & ImpreFormat(Arreglo(x, 8), 7, , True) & ImpreFormat(Arreglo(x, 11), 8, 0, True) & ImpreFormat(Arreglo(x, 12), 9, , True) & " ¦" & ImpreFormat(Arreglo(x, 9) + Arreglo(x, 11), 9, 0, True) & ImpreFormat(Arreglo(x, 10) + Arreglo(x, 12), 11, , True) & Chr(10)
'''    If vLineas > 66 Then 'pLineasMax
'''        vPage = vPage + 1
'''        pCadena = pCadena & gPrnSaltoPagina
'''        pCadena = pCadena & ImpreFormat(gsNomCmac, 35, 5) & Space(85) & "Página : " & ImpreFormat(vPage, 6, 0) & Chr(10)
'''        pCadena = pCadena & Space(5) & "Credito Pignoraticio" & Space(100) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & Chr(10)
'''        pCadena = pCadena & ImpreFormat(psNomAge, 35, 5) & Space(97) & gsCodUser & Chr(10)
'''        pCadena = pCadena & Space(50) & " ESTADISTICA DE LOTES DEL : " & pdIni & " AL " & pdFin & Chr(10)
'''        pCadena = pCadena & Space(50) & String(53, "=") & Chr(10)
'''        pCadena = pCadena & Space(5) & " Bovedas " & psBoveda & Chr(10)
'''        pCadena = pCadena & Space(1) & String(144, "-") & Chr(10)
'''        pCadena = pCadena & Space(5) & "       ¦              V  I  G  E  N  T  E  S                 ¦          A  D  J  U  D  I  C  A  D  O  S            ¦      T  O  T  A  L" & Chr(10)
'''        pCadena = pCadena & Space(5) & "       ¦     INGRESO         EGRESO           S T O C K      ¦     INGRESO         EGRESO           S T O C K      ¦ S T O C K  B O V E D A" & Chr(10)
'''        pCadena = pCadena & Space(5) & "FECHA  ¦ LOTES   GRAMOS  LOTES   GRAMOS    LOTES      GRAMOS ¦ LOTES   GRAMOS  LOTES   GRAMOS    LOTES      GRAMOS ¦    LOTES      GRAMOS" & Chr(10)
'''        pCadena = pCadena & Space(1) & String(144, "-") & Chr(10)
'''        vLineas = 13
'''    End If
'''    vLineas = vLineas + 1
'''    x = x + 1: vFecIni = DateAdd("d", 1, vFecIni)
'''Loop
'''pCadena = pCadena & Space(1) & String(144, "-") & Chr(10)
'''pCadena = pCadena & Space(13) & ImpreFormat(Arreglo(0, 1), 6, 0, True) & ImpreFormat(Arreglo(0, 2), 7, , True) & ImpreFormat(Arreglo(0, 3), 6, 0, True) & ImpreFormat(Arreglo(0, 4), 7, , True) & _
'''                    Space(22) & ImpreFormat(Arreglo(0, 5), 6, 0, True) & ImpreFormat(Arreglo(0, 6), 7, , True) & ImpreFormat(Arreglo(0, 7), 6, 0, True) & ImpreFormat(Arreglo(0, 8), 7, , True) & Chr(10)
'''pCadena = pCadena & Space(1) & String(144, "-") & Chr(10)
''''poProgress.value = 0
''''poProgress.Visible = False
'''
'''nRepo128042_Estadistica_x_Lote = pCadena
'''Exit Function
'''
'''ControlError:   ' Rutina de control de errores.
'''    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
'''        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'''End Function
'''Public Function SalNroEst(ByVal pBoveda As String, ByVal pEst1 As String, _
'''    Optional ByVal pEst2 As String = "", Optional ByVal pEst3 As String = "", _
'''    Optional ByVal pEst4 As String = "", Optional ByVal pEst5 As String = "", _
'''    Optional ByVal pEst6 As String = "", Optional ByVal pEst7 As String = "") As Currency
''''Sumas generales de saldos y creditos para la estadística
'''Dim tmpReg As New ADODB.Recordset
'''Dim tmpSql As String
'''Dim Conecta As DConecta
'''
'''Set Conecta = New DConecta
'''
'''tmpSql = " select Count(Pro.cCtaCod) AS Nume from ColocPignoraticio CP" & _
'''        " Inner Join Producto Pro on Pro.cCtaCod = CP.cCtaCod " & _
'''        " where Pro.nPrdEstado IN ('" & pEst1 & "','" & pEst2 & "','" & pEst3 & "','" & pEst4 & "'," & _
'''        " '" & pEst5 & "','" & pEst6 & "','" & pEst7 & "') "
'''
'''If Len(Trim(pBoveda)) > 0 Then  ' Boveda
'''    tmpSql = tmpSql & " AND cAgeBoveda in ('" & pBoveda & "') "
'''End If
'''
'''Conecta.AbreConexion
'''Set tmpReg = Conecta.CargaRecordSet(tmpSql)
'''Conecta.CierraConexion
'''
'''If (tmpReg.BOF Or tmpReg.EOF) Then
'''    SalNroEst = Format(0, "#0")
'''Else
'''    SalNroEst = IIf(IsNull(tmpReg!nume) = True, 0, Format(tmpReg!nume, "#0"))
'''End If
'''tmpReg.Close
'''Set tmpReg = Nothing
'''End Function
'''
'''Public Function SalSumEst(ByVal pBoveda As String, ByVal pNomReg As String, ByVal pEst1 As String, _
'''    Optional ByVal pEst2 As String = "", Optional ByVal pEst3 As String = "", _
'''    Optional ByVal pEst4 As String = "", Optional ByVal pEst5 As String = "", _
'''    Optional ByVal pEst6 As String = "", Optional ByVal pEst7 As String = "") As Currency
'''Dim tmpReg As New ADODB.Recordset
'''Dim tmpSql As String
'''Dim Conecta As DConecta
'''Set Conecta = New DConecta
'''
'''tmpSql = " select Count(Pro.cCtaCod) AS suma from ColocPignoraticio CP" & _
'''        " Inner Join Producto Pro on Pro.cCtaCod = CP.cCtaCod " & _
'''        " where Pro.nPrdEstado IN ('" & pEst1 & "','" & pEst2 & "','" & pEst3 & "','" & pEst4 & "'," & _
'''        " '" & pEst5 & "','" & pEst6 & "','" & pEst7 & "') "
'''
'''
'''If Len(Trim(pBoveda)) > 0 Then
'''    tmpSql = tmpSql & " AND cAgeBoveda in ('" & pBoveda & "') "
'''End If
'''
'''Conecta.AbreConexion
'''Set tmpReg = Conecta.CargaRecordSet(tmpSql)
'''Conecta.CierraConexion
'''
'''If (tmpReg.BOF Or tmpReg.EOF) Then
'''    SalSumEst = Format(0, "#0.00")
'''Else
'''    SalSumEst = IIf(IsNull(tmpReg!suma) = True, 0, Format(tmpReg!suma, "#0.00"))
'''End If
'''tmpReg.Close
'''Set tmpReg = Nothing
'''End Function
'''Private Function DiaNroEst(ByVal pBoveda As String, ByVal pFecha As String, ByVal pCodOpe As String, _
'''    Optional ByVal pCodOpe2 As String = "", Optional ByVal pCodOpe3 As String = "", _
'''    Optional ByVal pCodOpe4 As String = "", Optional ByVal pCodOpe5 As String = "", _
'''    Optional ByVal pCodOpe6 As String = "", Optional ByVal pCodOpe7 As String = "") As Double
''''Permite obtener el número de operaciones
''''de la tabla de TransPrenda de acuerdo a la condición deseada
'''Dim tmpReg As New ADODB.Recordset
'''Dim tmpSql As String
'''Dim Co As DConecta
'''Set Co = New DConecta
'''
'''If Len(Trim(pBoveda)) > 0 Then  ' Boveda
'''
'''    tmpSql = "select Count(MC.cCtaCod) AS Cuenta"
'''    tmpSql = tmpSql & " from Mov M"
'''    tmpSql = tmpSql & " Inner Join MovCol MC on Mc.nMovNro = M.nMovNro"
'''    tmpSql = tmpSql & " Inner Join ColocPignoraticio CP on CP.cCtaCod = MC.cCtaCod"
'''    tmpSql = tmpSql & " where M.nMovFlag=0 and CP.cAgeBoveda in ('" & pBoveda & "')  "
'''    tmpSql = tmpSql & " and  datediff(day,convert(datetime,substring(M.cMovNro,1,8),103),'" & Format(pFecha, "YYYY/MM/DD") & "')=0"
'''    tmpSql = tmpSql & " and  MC.nCredEstado in ('" & pCodOpe & "','" & pCodOpe2 & "','" & pCodOpe3 & "','" & pCodOpe4 & "','" & pCodOpe5 & "','" & pCodOpe6 & "','" & pCodOpe7 & "')   "
'''
'''Else
'''
'''    tmpSql = "select Count(MC.cCtaCod) AS Cuenta"
'''    tmpSql = tmpSql & " from Mov M"
'''    tmpSql = tmpSql & " Inner Join MovCol MC on Mc.nMovNro = M.nMovNro"
'''    tmpSql = tmpSql & " where M.nMovFlag=0  "
'''    tmpSql = tmpSql & " and  datediff(day,convert(datetime,substring(M.cMovNro,1,8),103),'" & Format(pFecha, "YYYY/MM/DD") & "')=0"
'''    tmpSql = tmpSql & " and  MC.nCredEstado in ('" & pCodOpe & "','" & pCodOpe2 & "','" & pCodOpe3 & "','" & pCodOpe4 & "','" & pCodOpe5 & "','" & pCodOpe6 & "','" & pCodOpe7 & "')   "
'''
'''End If
'''Co.AbreConexion
'''Set tmpReg = Co.CargaRecordSet(tmpSql)
'''If (tmpReg.BOF Or tmpReg.EOF) Then
'''    DiaNroEst = Format(0, "#0")
'''Else
'''    DiaNroEst = IIf(IsNull(tmpReg!Cuenta) = True, 0, Format(tmpReg!Cuenta, "#0"))
'''End If
'''tmpReg.Close
'''Co.CierraConexion
'''Set tmpReg = Nothing
'''End Function
'''Private Function DiaSumCP(ByVal pBoveda As String, ByVal pFecha As Date, ByVal pNomReg As String, ByVal pCodOpe As String, _
'''    Optional ByVal pCodOpe2 As String = "", Optional ByVal pCodOpe3 As String = "", _
'''    Optional ByVal pCodOpe4 As String = "", Optional ByVal pCodOpe5 As String = "", _
'''    Optional ByVal pCodOpe6 As String = "", Optional ByVal pCodOpe7 As String = "") As Currency
'''Dim tmpReg As New ADODB.Recordset
'''Dim tmpSql As String
'''Dim Co As DConecta
'''Set Co = New DConecta
'''Co.AbreConexion
'''
'''
'''    tmpSql = " select sum(abs(CP." & pNomReg & ")) AS Suma"
'''    tmpSql = tmpSql & " from Mov M"
'''    tmpSql = tmpSql & " Inner Join MovCol MC on Mc.nMovNro = M.nMovNro"
'''    tmpSql = tmpSql & " Inner Join ColocPignoraticio CP on CP.cCtaCod = MC.cCtaCod"
'''    tmpSql = tmpSql & " where M.nMovFlag=0 "
'''    tmpSql = tmpSql & " and  datediff(day,convert(datetime,substring(cMovNro,1,8),103),'" & Format(pFecha, "YYYY/MM/DD") & "')=0"
'''    tmpSql = tmpSql & " and  MC.cOpeCod in ('" & pCodOpe & "','" & pCodOpe2 & "','" & pCodOpe3 & "','" & pCodOpe4 & "','" & pCodOpe5 & "','" & pCodOpe6 & "','" & pCodOpe7 & "')"
'''
'''If Len(Trim(pBoveda)) > 0 Then  ' Boveda
'''    tmpSql = tmpSql & " AND cp.cAgeBoveda in ('" & pBoveda & "')  "
'''End If
'''
'''Set tmpReg = Co.CargaRecordSet(tmpSql)
'''
'''If (tmpReg.BOF Or tmpReg.EOF) Then
'''    DiaSumCP = 0
'''Else
'''    DiaSumCP = IIf(IsNull(tmpReg!suma), 0, tmpReg!suma)
'''End If
'''tmpReg.Close
'''Set tmpReg = Nothing
'''Co.CierraConexion
'''End Function



Public Function nRepo128041_Estadistica_Saldos(ByVal pdIni As Date, ByVal pdFin As Date, _
ByVal psCodAge As String, ByVal pdFecsis As Date, ByVal psNomCmac As String, _
ByVal psNomAge As String) As String

Dim pCadena As String
Dim vSaldo As Currency
Dim vNumero As Currency
Dim vLineas As Integer
Dim vPage As Integer
Dim lsSQL As String
Dim lrReg As New ADODB.Recordset
Dim lsPrestamos As String, lsCancelacion As String, lsRenovacion As String
Dim paso1 As Boolean
Dim Co As DConecta
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

Set Co = New DConecta

If (DateDiff("d", pdIni, pdFin) + 1) > 32650 Then
    MsgBox " Rango es demasiado amplio ", vbInformation, " Aviso "
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

gdHoraGrab = Format(pdFecsis & " " & Time, "dd/mm/yyyy hh:mm:ss")
vPage = 1
pCadena = pCadena & Chr(10)
pCadena = pCadena & ImpreFormat(psNomCmac, 35, 5) & Space(100) & "Página : " & ImpreFormat(vPage, 6, 0) & Chr(10)
pCadena = pCadena & ImpreFormat(psNomAge, 35, 5) & Space(99) & Format(pdFecsis & " " & Time, "dd/mm/yyyy hh:mm") & Chr(10)
'pCadena = pCadena & Space(4) & "Bovedas " & psBoveda & Chr(10)
pCadena = pCadena & Space(55) & " ESTADISTICA DE SALDOS DEL " & pdIni & " - AL " & pdFin & Chr(10)
pCadena = pCadena & Space(55) & String(54, "=") & Chr(10)
pCadena = pCadena & Chr(10)
pCadena = pCadena & Space(2) & String(130, "-") & Chr(10)
pCadena = pCadena & Space(7) & "FECHA     NRO. PREST.NUEVOS  NRO. MONTO.CANCEL.   NRO.  MONTO.RENOVA.  NRO. MONTO.AMORT.   NRO. MONTO ITF  Nro    SALDO.VIG" & Chr(10)
pCadena = pCadena & Space(2) & String(130, "-") & Chr(10)

vLineas = 13
While Not lrReg.EOF
        With lrReg
            pCadena = pCadena & Space(2) & Format(!Fecha, "DD/MM/YYYY") & ImpreFormat(!NroPres, 6, 0) & ImpreFormat(!MonPres, 12, , True) _
                    & ImpreFormat(!NroCanc, 6, 0) & ImpreFormat(!MonCanc, 12, , True) & ImpreFormat(!NroRenov, 6, 0) _
                    & ImpreFormat(!MonRenov, 12, , True) & ImpreFormat(!NroAmort, 6, 0) & ImpreFormat(!MonAmort, 12, , True) & ImpreFormat(!NroITF, 5, 0) & ImpreFormat(!MonITF, 7, , True) _
                    & ImpreFormat(!nCredVig, 6, 0) & ImpreFormat(!nSalVig, 12, , True) & Chr(10)
                    
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
          pCadena = pCadena & Chr(12)
          pCadena = pCadena & Space(135) & "Página : " & ImpreFormat(vPage, 6, 0) & Chr(10)
          pCadena = pCadena & Space(2) & String(130, "-") & Chr(10)
          pCadena = pCadena & Space(7) & "FECHA      NRO. PREST.NUEVOS  NRO. MONTO.CANCEL.   NRO. MONTO.RENOVA.  NRO. MONTO.AMORT.   NRO. MONTO ITF  Nro    SALDO.VIG " & Chr(10)
          pCadena = pCadena & Space(2) & String(130, "-") & Chr(10)
          vLineas = 5
        End If
        vLineas = vLineas + 1
    lrReg.MoveNext
Wend

pCadena = pCadena & Space(2) & String(130, "-") & Chr(10)
pCadena = pCadena & Space(12) & ImpreFormat(n1, 6, 0) & ImpreFormat(sum1, 12, , True) & ImpreFormat(n2, 6, 0) & ImpreFormat(sum2, 12, , True) & ImpreFormat(n3, 6, 0) & ImpreFormat(sum3, 12, , True) & ImpreFormat(n5, 6, 0) & ImpreFormat(sum5, 12, , True) & ImpreFormat(n4, 6, 0) & ImpreFormat(sum4, 8, , True) & Chr(10)
pCadena = pCadena & Space(2) & String(130, "-") & Chr(10)

'Devuelve cadena con Impresión
nRepo128041_Estadistica_Saldos = pCadena & Chr(12)
Set Co = Nothing
Set lrReg = Nothing
End Function

'Nuevo
Public Function nRepo128042_Estadistica_x_Lotes(ByVal pdIni As Date, ByVal pdFin As Date, ByVal psBoveda As String, _
                                                ByVal psAgencias As String, ByVal pdFecsis As Date) As String

Dim oCon As New DConecta
Dim rs As New ADODB.Recordset
Dim sql As String

Dim lnFila As Integer
Dim sAgeTempo As String
Dim n1 As Integer
Dim n2 As Integer

sql = "Select CDP.dEstad, CDP.cCodAge, "
'sql = sql & " CDP.cLineaCred, "
sql = sql & " CDP.cAgeBoveda, SUM(CDP.nnumCredVig) nnumCredVig, SUM(CDP.nOroVig) nOroVig, "
sql = sql & " SUM(CDP.nCapVig) nCapVig, SUM(CDP.nNumCredAdj) nNumCredAdj, SUM(CDP.nOroAdj) nOroAdj, SUM(CDP.nCapadj) nCapadj, "
sql = sql & " SUM(CDP.nNumCredDif) nNumCredDif, SUM(CDP.nOroDif) nOroDif"
sql = sql & " from ColocEstadDiaPrendaBov CDP "
sql = sql & " where convert(varchar(8), CDP.dEstad, 112) >='" & Format(pdIni, "YYYYMMdd") & "' "
sql = sql & " And convert(varchar(8), CDP.dEstad, 112) <='" & Format(pdFin, "YYYYMMdd") & "' "
sql = sql & " and CDP.cAgeBoveda IN(" & psBoveda & ") "
sql = sql & " and CDP.cCodAge IN(" & psAgencias & ") "
sql = sql & " Group By CDP.dEstad, CDP.cCodAge, CDP.cAgeBoveda "
sql = sql & " Order By cDP.cCodAge, CDP.dEstad, CDP.cAgeBoveda "

oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sql)
If rs.BOF Then
Else
    
    lsArchivoN = App.path & "\Spooler\EstLotes" & Format(pdFin & " " & Time, "yyyymmddhhmmss") & gsCodUser & ".xls"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    
    If lbLibroOpen Then
        Set xlHoja1 = xlLibro.Worksheets(1)
        ExcelAddHoja "Estadistica", xlLibro, xlHoja1
    
        xlHoja1.Cells(2, 2) = "Caja Municipal de Ica"
        xlHoja1.Cells(2, 10) = "Fecha"
        xlHoja1.Cells(2, 11) = "'" & Format(pdFecsis, "dd/MM/YYYY")
        xlHoja1.Range("B2:D2").MergeCells = True
        
        xlHoja1.Cells(3, 2) = "Crédito Pignoraticio"
        
        xlHoja1.Cells(5, 2) = "ESTADISTICA DE LOTES DEL " & Format(pdIni, "dd/MM/YYYY") & " Al " & Format(pdFin, "dd/MM/YYYY")
        xlHoja1.Range("B5:K5").MergeCells = True
        xlHoja1.Range("B5:K5").Font.Underline = True
        xlHoja1.Range("B5:K5").HorizontalAlignment = xlCenter
                
        xlHoja1.Range("B2:K5").Font.Bold = True
        
        lnFila = 7
        
        Do While Not rs.EOF
                 
            If sAgeTempo <> rs!CCodage Then
                
                If Len(Trim(sAgeTempo)) > 0 Then
                    ExcelCuadro xlHoja1, 2, n1, 11, lnFila
                    '
                    lnFila = lnFila + 1
                    xlHoja1.Cells(lnFila, 2) = "Total"
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 4)).Formula = "=SUM(D" & n1 & ":D" & lnFila - 1 & ")"
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Formula = "=SUM(E" & n1 & ":E" & lnFila - 1 & ")"
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=SUM(F" & n1 & ":F" & lnFila - 1 & ")"
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=SUM(G" & n1 & ":G" & lnFila - 1 & ")"
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).Formula = "=SUM(H" & n1 & ":H" & lnFila - 1 & ")"
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila, 9)).Formula = "=SUM(I" & n1 & ":I" & lnFila - 1 & ")"
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 10), xlHoja1.Cells(lnFila, 10)).Formula = "=SUM(J" & n1 & ":J" & lnFila - 1 & ")"
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 11), xlHoja1.Cells(lnFila, 11)).Formula = "=SUM(K" & n1 & ":K" & lnFila - 1 & ")"
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 11)).Font.Bold = True
                    ExcelCuadro xlHoja1, 2, lnFila, 11, lnFila
                    '
                End If
                
                lnFila = lnFila + 2
                xlHoja1.Cells(lnFila, 2) = "Agencia " & rs!CCodage
                xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 11)).Font.Bold = True
                
                lnFila = lnFila + 1
                xlHoja1.Cells(lnFila, 2) = "Fecha"
                xlHoja1.Cells(lnFila, 3) = "Bóveda"
                xlHoja1.Cells(lnFila, 4) = "Créditos Vigentes"
                xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 6)).MergeCells = True
                xlHoja1.Cells(lnFila, 7) = "Créditos Adjudicados"
                xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 9)).MergeCells = True
                xlHoja1.Cells(lnFila, 10) = "Créditos Diferidos"
                xlHoja1.Range(xlHoja1.Cells(lnFila, 10), xlHoja1.Cells(lnFila, 11)).MergeCells = True
                
                lnFila = lnFila + 1
                xlHoja1.Cells(lnFila, 4) = "Gramos Oro"
                xlHoja1.Cells(lnFila, 5) = "Cantidad"
                xlHoja1.Cells(lnFila, 6) = "Monto"
                xlHoja1.Cells(lnFila, 7) = "Gramos Oro"
                xlHoja1.Cells(lnFila, 8) = "Cantidad"
                xlHoja1.Cells(lnFila, 9) = "Monto"
                xlHoja1.Cells(lnFila, 10) = "Gramos Oro"
                xlHoja1.Cells(lnFila, 11) = "Cantidad"
                
                xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila, 11)).HorizontalAlignment = xlCenter
                xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 2), xlHoja1.Cells(lnFila, 11)).Font.Bold = True
                
                ExcelCuadro xlHoja1, 2, lnFila - 1, 11, lnFila, , True
                
                n1 = lnFila + 1
                
                sAgeTempo = rs!CCodage
                
            End If
                 
            lnFila = lnFila + 1
                 
            xlHoja1.Cells(lnFila, 2) = "'" & Format(rs!DESTAD, "dd/MM/YYYY")
            xlHoja1.Cells(lnFila, 3) = rs!cAgeBoveda
            xlHoja1.Cells(lnFila, 4) = Format(rs!nOroVig, "0.00")
            xlHoja1.Cells(lnFila, 5) = Format(rs!nNumCredVig, "0")
            xlHoja1.Cells(lnFila, 6) = Format(rs!nCapVig, "0.00")
            xlHoja1.Cells(lnFila, 7) = Format(rs!nOroAdj, "0.00")
            xlHoja1.Cells(lnFila, 8) = Format(rs!nNumCredAdj, "0")
            xlHoja1.Cells(lnFila, 9) = Format(rs!nCapadj, "0.00")
            xlHoja1.Cells(lnFila, 10) = Format(rs!nOroDif, "0.00")
            xlHoja1.Cells(lnFila, 11) = Format(rs!nNumCredDif, "0")
            
            n2 = lnFila
            
            rs.MoveNext
        Loop
        
        If Len(Trim(sAgeTempo)) > 0 Then
            ExcelCuadro xlHoja1, 2, n1, 11, lnFila
            '
            lnFila = lnFila + 1
            xlHoja1.Cells(lnFila, 2) = "Total"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 4)).Formula = "=SUM(D" & n1 & ":D" & lnFila - 1 & ")"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Formula = "=SUM(E" & n1 & ":E" & lnFila - 1 & ")"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=SUM(F" & n1 & ":F" & lnFila - 1 & ")"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=SUM(G" & n1 & ":G" & lnFila - 1 & ")"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).Formula = "=SUM(H" & n1 & ":H" & lnFila - 1 & ")"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila, 9)).Formula = "=SUM(I" & n1 & ":I" & lnFila - 1 & ")"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 10), xlHoja1.Cells(lnFila, 10)).Formula = "=SUM(J" & n1 & ":J" & lnFila - 1 & ")"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 11), xlHoja1.Cells(lnFila, 11)).Formula = "=SUM(K" & n1 & ":K" & lnFila - 1 & ")"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 11)).Font.Bold = True
            ExcelCuadro xlHoja1, 2, lnFila, 11, lnFila
            '
            xlHoja1.Range(xlHoja1.Cells(6, 4), xlHoja1.Cells(lnFila, 4)).NumberFormat = "###,###.00"
            xlHoja1.Range(xlHoja1.Cells(6, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "###,###.00"
            xlHoja1.Range(xlHoja1.Cells(6, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "###,###.00"
            xlHoja1.Range(xlHoja1.Cells(6, 9), xlHoja1.Cells(lnFila, 9)).NumberFormat = "###,###.00"
            xlHoja1.Range(xlHoja1.Cells(6, 10), xlHoja1.Cells(lnFila, 10)).NumberFormat = "###,###.00"
            
            xlHoja1.Range(xlHoja1.Cells(6, 5), xlHoja1.Cells(lnFila, 5)).NumberFormat = "###,###"
            xlHoja1.Range(xlHoja1.Cells(6, 8), xlHoja1.Cells(lnFila, 8)).NumberFormat = "###,###"
            xlHoja1.Range(xlHoja1.Cells(6, 11), xlHoja1.Cells(lnFila, 11)).NumberFormat = "###,###"
            
            '
        End If
    
        xlHoja1.Cells.Select
        xlHoja1.Cells.Font.Name = "Arial"
        xlHoja1.Cells.Font.Size = 9
        xlHoja1.Cells.EntireColumn.AutoFit
        ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
    
       ' CargaArchivo lsArchivoN, App.path & "\SPOOLER\"
    
    End If
End If
Set rs = Nothing

End Function

'Fin Nuevo


'Nuevo
Public Function nRepo128049_Estadistica_Oro(sCodage, psFecIni As String, psFecFin As String, pdFecsis As Date, ByVal psNomCmac As String, _
ByVal psNomAge As String) As String


Dim pCadena As String
Dim vSaldo As Currency
Dim vNumero As Currency
Dim vLineas As Integer
Dim vPage As Integer
Dim lsSQL As String
Dim lrReg As New ADODB.Recordset
Dim lsPrestamos As String, lsCancelacion As String, lsRenovacion As String
Dim paso1 As Boolean
Dim Co As DConecta, CLSAGE As DAgencias
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

Set CLSAGE = New DAgencias
Set Co = New DConecta

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

If psFecFin = Format(pdFecsis, "yyyymmdd") Then
        lsSQL = lsSQL & " Union "
        lsSQL = lsSQL & " SELECT dfecha=convert(char(8),cast('" & Format(pdFecsis, "yyyymmdd") & "' as datetime),112) ,CAGECOD=SUBSTRING(CCTACOD,4,2) , "
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

gdHoraGrab = Format(pdFecsis & " " & Time, "dd/mm/yyyy hh:mm:ss")
vPage = 1
pCadena = pCadena & Chr(10)
pCadena = pCadena & ImpreFormat("CMAC ICA", 35, 5) & Space(100) & "Página : " & ImpreFormat(vPage, 6, 0) & Chr(10)
pCadena = pCadena & ImpreFormat(psNomAge, 35, 5) & Space(99) & Format(pdFecsis & " " & Time, "dd/mm/yyyy hh:mm") & Chr(10)
'pCadena = pCadena & Space(4) & "Bovedas " & psBoveda & Chr(10)
pCadena = pCadena & Space(55) & " ESTADISTICA DE ORO DEL " & Right(psFecIni, 2) & "/" & Mid(psFecIni, 5, 2) & "/" & Left(psFecIni, 4) & " AL " & Right(psFecFin, 2) & "/" & Mid(psFecFin, 5, 2) & "/" & Left(psFecFin, 4) & " - " & CLSAGE.NombreAgencia(sCodage) & Chr(10)
pCadena = pCadena & Space(55) & String(54, "=") & Chr(10)
pCadena = pCadena & Chr(10)
pCadena = pCadena & Space(2) & String(140, "-") & Chr(10)
pCadena = pCadena & Space(2) & "FECHA   CANTIDAD VIG.    VIG14K     VIG16K    VIG18K    VIG21K  TOTALVIG   CANTIDAD ADJ.  ADJ14     ADJ16     ADJ18     ADJ21  TOTALADJ " & Chr(10)
pCadena = pCadena & Space(2) & String(140, "-") & Chr(10)

vLineas = 13
While Not lrReg.EOF
        With lrReg
            pCadena = pCadena & Space(2) & Right(!dFecha, 2) & "/" & Mid(!dFecha, 5, 2) & "/" & Left(!dFecha, 4) & Space(3) & ImpreFormat(!nCantVig, 6, 0) & Space(2) & ImpreFormat(!nk14Vig, 6, 2, False) & Space(2) & ImpreFormat(!nk16Vig, 6, 2, False) & Space(2) & ImpreFormat(!nk18Vig, 6, 2, False) & Space(2) & ImpreFormat(!nk21Vig, 6, 2, False) & Space(2) & ImpreFormat(!nk14Vig + !nk16Vig + !nk18Vig + !nk21Vig, 6, 2, False) & Space(2) _
                      & ImpreFormat(!nCantAdj, 6, 0) & Space(2) & ImpreFormat(!nk14Adj, 6, 2, False) & Space(2) & ImpreFormat(!nk16Adj, 6, 2, False) & Space(2) & ImpreFormat(!nk18Adj, 6, 2, False) & Space(2) & ImpreFormat(!nk21Adj, 6, 2, False) & Space(2) & ImpreFormat(!nk14Adj + !nk16Adj + !nk18Adj + !nk21Adj, 6, 2, False) & Chr(10)
                    
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
          pCadena = pCadena & Chr(12)
          pCadena = pCadena & Space(135) & "Página : " & ImpreFormat(vPage, 6, 0) & Chr(10)
          pCadena = pCadena & Space(2) & String(130, "-") & Chr(10)
          pCadena = pCadena & Space(55) & " ESTADISTICA DE ORO DEL " & Right(psFecIni, 2) & "/" & Mid(psFecIni, 5, 2) & "/" & Left(psFecIni, 4) & " AL " & Right(psFecFin, 2) & "/" & Mid(psFecFin, 5, 2) & "/" & Left(psFecFin, 4) & " - " & CLSAGE.NombreAgencia(sCodage) & Chr(10)
          pCadena = pCadena & Space(55) & String(54, "=") & Chr(10)
          pCadena = pCadena & Chr(10)
          pCadena = pCadena & Space(2) & String(130, "-") & Chr(10)
          pCadena = pCadena & Space(2) & "FECHA   CANTIDAD VIG.    VIG14K     VIG16K    VIG18K    VIG21K  TOTALVIG   CANTIDAD ADJ.  ADJ14     ADJ16     ADJ18     ADJ21  TOTALADJ " & Chr(10)
          pCadena = pCadena & Space(2) & String(130, "-") & Chr(10)
          vLineas = 13
        End If
        vLineas = vLineas + 1
    lrReg.MoveNext
Wend

'pCadena = pCadena & Space(2) & String(140, "-") & Chr(10)
'pCadena = pCadena & Space(8) & "TOTAL: " & Space(15) & ImpreFormat(n14, 8, , True) & ImpreFormat(n16, 8, , True) & ImpreFormat(n18, 8, , True) & ImpreFormat(n21, 8, , True) & ImpreFormat(b14, 8, , True) & ImpreFormat(b16, 8, , True) & ImpreFormat(b18, 8, , True) & ImpreFormat(b21, 8, , True) & Chr(10)
'pCadena = pCadena & Space(2) & String(140, "-") & Chr(10)
'
'Devuelve cadena con Impresión
nRepo128049_Estadistica_Oro = pCadena & Chr(12)
Set Co = Nothing
Set lrReg = Nothing



End Function



''''Public Function nRepo128042_Estadistica_x_Lotes(ByVal pdIni As Date, _
''''ByVal pdFin As Date, ByVal psBoveda As String, _
''''ByVal psNomAge As String, ByVal pdFecSis As Date) As String
''''
''''On Error GoTo ControlError
''''Dim pCadena As String
''''Dim vVigSal As Currency, vAdjSal As Currency
''''Dim vVigNum As Currency, vAdjNum As Currency
''''Dim X As Integer, Y As Integer
''''Dim vFecIni As Date, vFecFin As Date
''''Dim Arreglo() As Currency
''''Dim vLineas As Integer
''''Dim vPage As Integer
''''Dim gdHoraGrab As Date
''''Erase Arreglo
''''If (DateDiff("d", pdIni, pdFin) + 1) > 32650 Then
''''    MsgBox " Rango es demasiado amplio ", vbInformation, " Aviso "
''''    nRepo128042_Estadistica_x_Lote = ""
''''    Exit Function
''''End If
''''X = DateDiff("d", pdIni, pdFin) + 1
''''If X > 365 Then
''''    If MsgBox("Rango es muy amplio" & vbCr & "Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton1, " Aviso ") = vbNo Then
''''        nRepo128042_Estadistica_x_Lote = ""
''''        Exit Function
''''    End If
''''End If
''''ReDim Arreglo(X, 12)
''''vFecIni = pdIni: vFecFin = pdFin
''''Y = 1
''''vVigNum = SalNroEst(psBoveda, "2100", "2101", "2102", "2104", "2106", "2107")
''''vVigSal = SalSumEst(psBoveda, "nOroNeto", "2100", "2101", "2102", "2104", "2106", "2107")
''''vAdjNum = SalNroEst(psBoveda, "2108")
''''vAdjSal = SalSumEst(psBoveda, "nOroNeto", "2108")
''''
''''
''''
''''Dim paso1 As Boolean
''''paso1 = True
''''Do While DateDiff("d", vFecIni, vFecFin) >= 0
''''    'Vigentes : Ingresos
''''    Arreglo(X, 1) = DiaNroEst(psBoveda, vFecFin, 2100) - DiaNroEst(psBoveda, vFecFin, 2110)
''''    Arreglo(X, 2) = DiaSumCP(psBoveda, vFecFin, "nOroNeto", 2100) - DiaSumCP(psBoveda, vFecFin, "nOroNeto", 2110)
''''   'Vigentes : Egresos
''''    Arreglo(X, 3) = DiaNroEst(psBoveda, vFecFin, gsDevJoyas, gsDevJoyasEOA, gsVtaRemate)
''''    Arreglo(X, 4) = DiaSumCP(psBoveda, vFecFin, "nOroNeto", gsDevJoyas, gsDevJoyasEOA, gsVtaRemate, gsVtaRemEnOtAg)
''''    'Adjudicados : Ingresos
''''    Arreglo(X, 5) = DiaNroEst(psBoveda, vFecFin, gsAdjudica)
''''    Arreglo(X, 6) = DiaSumCP(psBoveda, vFecFin, "nOroNeto", gsAdjudica)
''''    'Adjudicados : Egresos
''''    Arreglo(X, 7) = DiaNroEst(psBoveda, vFecFin, gsVtaSubasta)
''''    Arreglo(X, 8) = DiaSumCP(psBoveda, vFecFin, "nOroNeto", gsVtaSubasta, gsVtaSubEnOtAg)
''''
''''    'Saldos Vigentes  -  Adjudicados
''''    If paso1 Then
''''        Arreglo(X, 9) = vVigNum
''''        Arreglo(X, 10) = vVigSal
''''        Arreglo(X, 11) = vAdjNum
''''        Arreglo(X, 12) = vAdjSal
''''        paso1 = False
''''    Else
''''        Arreglo(X, 9) = Abs(Arreglo(X + 1, 9) - Arreglo(X + 1, 1)) + Arreglo(X + 1, 3)
''''        Arreglo(X, 10) = Abs(Arreglo(X + 1, 10) - Arreglo(X + 1, 2)) + Arreglo(X + 1, 4)
''''        Arreglo(X, 11) = Abs(Arreglo(X + 1, 11) - Arreglo(X + 1, 5)) + Arreglo(X + 1, 7)
''''        Arreglo(X, 12) = Abs(Arreglo(X + 1, 12) - Arreglo(X + 1, 6)) + Arreglo(X + 1, 8)
''''    End If
''''    Arreglo(0, 1) = Arreglo(0, 1) + Arreglo(X, 1)
''''    Arreglo(0, 2) = Arreglo(0, 2) + Arreglo(X, 2)
''''    Arreglo(0, 3) = Arreglo(0, 3) + Arreglo(X, 3)
''''    Arreglo(0, 4) = Arreglo(0, 4) + Arreglo(X, 4)
''''    Arreglo(0, 5) = Arreglo(0, 5) + Arreglo(X, 5)
''''    Arreglo(0, 6) = Arreglo(0, 6) + Arreglo(X, 6)
''''    Arreglo(0, 7) = Arreglo(0, 7) + Arreglo(X, 7)
''''    Arreglo(0, 8) = Arreglo(0, 8) + Arreglo(X, 8)
''''
''''    'poProgress.value = y
''''    Y = Y + 1
''''    X = X - 1: vFecFin = DateAdd("d", -1, vFecFin)
''''Loop
''''
'''''Carga fecha y hora de grabación
''''gdHoraGrab = Format(pdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss")
''''vPage = 1
''''pCadena = pCadena & Chr(10)
''''pCadena = pCadena & ImpreFormat(gsNomCmac, 35, 5) & Space(85) & "Página : " & ImpreFormat(vPage, 6, 0) & Chr(10)
''''pCadena = pCadena & Space(5) & "Credito Pignoraticio" & Space(100) & Format(pdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & Chr(10)
''''pCadena = pCadena & ImpreFormat(psNomAge, 35, 5) & Space(97) & gsCodUser & Chr(10)
''''pCadena = pCadena & Space(50) & " ESTADISTICA DE LOTES DEL : " & pdIni & " AL " & pdFin & Chr(10)
''''pCadena = pCadena & Space(50) & String(53, "=") & Chr(10)
''''pCadena = pCadena & Space(5) & " Bovedas " & psBoveda & Chr(10)
''''pCadena = pCadena & Space(1) & String(144, "-") & Chr(10)
''''pCadena = pCadena & Space(5) & "       ¦              V  I  G  E  N  T  E  S                 ¦          A  D  J  U  D  I  C  A  D  O  S            ¦      T  O  T  A  L" & Chr(10)
''''pCadena = pCadena & Space(5) & "       ¦     INGRESO         EGRESO           S T O C K      ¦     INGRESO         EGRESO           S T O C K      ¦ S T O C K  B O V E D A" & Chr(10)
''''pCadena = pCadena & Space(5) & "FECHA  ¦ LOTES   GRAMOS  LOTES   GRAMOS    LOTES      GRAMOS ¦ LOTES   GRAMOS  LOTES   GRAMOS    LOTES      GRAMOS ¦    LOTES      GRAMOS" & Chr(10)
''''pCadena = pCadena & Space(1) & String(144, "-") & Chr(10)
''''
''''X = 1
''''vLineas = 13
''''vFecIni = pdIni: vFecFin = pdFin
''''Do While DateDiff("d", vFecIni, vFecFin) >= 0
''''    pCadena = pCadena & Space(2) & vFecIni & "¦" & ImpreFormat(Arreglo(X, 1), 6, 0, True) & ImpreFormat(Arreglo(X, 2), 7, , True) & ImpreFormat(Arreglo(X, 3), 6, 0, True) & ImpreFormat(Arreglo(X, 4), 7, , True) & ImpreFormat(Arreglo(X, 9), 8, 0, True) & ImpreFormat(Arreglo(X, 10), 9, , True) & " ¦" & _
''''                                            ImpreFormat(Arreglo(X, 5), 6, 0, True) & ImpreFormat(Arreglo(X, 6), 7, , True) & ImpreFormat(Arreglo(X, 7), 6, 0, True) & ImpreFormat(Arreglo(X, 8), 7, , True) & ImpreFormat(Arreglo(X, 11), 8, 0, True) & ImpreFormat(Arreglo(X, 12), 9, , True) & " ¦" & ImpreFormat(Arreglo(X, 9) + Arreglo(X, 11), 9, 0, True) & ImpreFormat(Arreglo(X, 10) + Arreglo(X, 12), 11, , True) & Chr(10)
''''    If vLineas > 66 Then 'pLineasMax
''''        vPage = vPage + 1
''''        pCadena = pCadena & gPrnSaltoPagina
''''        pCadena = pCadena & ImpreFormat(gsNomCmac, 35, 5) & Space(85) & "Página : " & ImpreFormat(vPage, 6, 0) & Chr(10)
''''        pCadena = pCadena & Space(5) & "Credito Pignoraticio" & Space(100) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & Chr(10)
''''        pCadena = pCadena & ImpreFormat(psNomAge, 35, 5) & Space(97) & gsCodUser & Chr(10)
''''        pCadena = pCadena & Space(50) & " ESTADISTICA DE LOTES DEL : " & pdIni & " AL " & pdFin & Chr(10)
''''        pCadena = pCadena & Space(50) & String(53, "=") & Chr(10)
''''        pCadena = pCadena & Space(5) & " Bovedas " & psBoveda & Chr(10)
''''        pCadena = pCadena & Space(1) & String(144, "-") & Chr(10)
''''        pCadena = pCadena & Space(5) & "       ¦              V  I  G  E  N  T  E  S                 ¦          A  D  J  U  D  I  C  A  D  O  S            ¦      T  O  T  A  L" & Chr(10)
''''        pCadena = pCadena & Space(5) & "       ¦     INGRESO         EGRESO           S T O C K      ¦     INGRESO         EGRESO           S T O C K      ¦ S T O C K  B O V E D A" & Chr(10)
''''        pCadena = pCadena & Space(5) & "FECHA  ¦ LOTES   GRAMOS  LOTES   GRAMOS    LOTES      GRAMOS ¦ LOTES   GRAMOS  LOTES   GRAMOS    LOTES      GRAMOS ¦    LOTES      GRAMOS" & Chr(10)
''''        pCadena = pCadena & Space(1) & String(144, "-") & Chr(10)
''''        vLineas = 13
''''    End If
''''    vLineas = vLineas + 1
''''    X = X + 1: vFecIni = DateAdd("d", 1, vFecIni)
''''Loop
''''pCadena = pCadena & Space(1) & String(144, "-") & Chr(10)
''''pCadena = pCadena & Space(13) & ImpreFormat(Arreglo(0, 1), 6, 0, True) & ImpreFormat(Arreglo(0, 2), 7, , True) & ImpreFormat(Arreglo(0, 3), 6, 0, True) & ImpreFormat(Arreglo(0, 4), 7, , True) & _
''''                    Space(22) & ImpreFormat(Arreglo(0, 5), 6, 0, True) & ImpreFormat(Arreglo(0, 6), 7, , True) & ImpreFormat(Arreglo(0, 7), 6, 0, True) & ImpreFormat(Arreglo(0, 8), 7, , True) & Chr(10)
''''pCadena = pCadena & Space(1) & String(144, "-") & Chr(10)
'''''poProgress.value = 0
'''''poProgress.Visible = False
''''
''''nRepo128042_Estadistica_x_Lote = pCadena
''''Exit Function
''''
''''ControlError:   ' Rutina de control de errores.
''''    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
''''        " Avise al Area de Sistemas ", vbInformation, " Aviso "
''''End Function
''''Public Function SalNroEst(ByVal pBoveda As String, ByVal pEst1 As String, _
''''    Optional ByVal pEst2 As String = "", Optional ByVal pEst3 As String = "", _
''''    Optional ByVal pEst4 As String = "", Optional ByVal pEst5 As String = "", _
''''    Optional ByVal pEst6 As String = "", Optional ByVal pEst7 As String = "") As Currency
'''''Sumas generales de saldos y creditos para la estadística
''''Dim tmpReg As New ADODB.Recordset
''''Dim tmpSql As String
''''Dim Conecta As DConecta
''''
''''Set Conecta = New DConecta
''''
''''tmpSql = " select Count(Pro.cCtaCod) AS Nume from ColocPignoraticio CP" & _
''''        " Inner Join Producto Pro on Pro.cCtaCod = CP.cCtaCod " & _
''''        " where Pro.nPrdEstado IN ('" & pEst1 & "','" & pEst2 & "','" & pEst3 & "','" & pEst4 & "'," & _
''''        " '" & pEst5 & "','" & pEst6 & "','" & pEst7 & "') "
''''
''''If Len(Trim(pBoveda)) > 0 Then  ' Boveda
''''    tmpSql = tmpSql & " AND cAgeBoveda in ('" & pBoveda & "') "
''''End If
''''
''''Conecta.AbreConexion
''''Set tmpReg = Conecta.CargaRecordSet(tmpSql)
''''Conecta.CierraConexion
''''
''''If (tmpReg.BOF Or tmpReg.EOF) Then
''''    SalNroEst = Format(0, "#0")
''''Else
''''    SalNroEst = IIf(IsNull(tmpReg!nume) = True, 0, Format(tmpReg!nume, "#0"))
''''End If
''''tmpReg.Close
''''Set tmpReg = Nothing
''''End Function
''''
''''Public Function SalSumEst(ByVal pBoveda As String, ByVal pNomReg As String, ByVal pEst1 As String, _
''''    Optional ByVal pEst2 As String = "", Optional ByVal pEst3 As String = "", _
''''    Optional ByVal pEst4 As String = "", Optional ByVal pEst5 As String = "", _
''''    Optional ByVal pEst6 As String = "", Optional ByVal pEst7 As String = "") As Currency
''''Dim tmpReg As New ADODB.Recordset
''''Dim tmpSql As String
''''Dim Conecta As DConecta
''''Set Conecta = New DConecta
''''
''''tmpSql = " select Sum(Pro.nSaldo) AS suma from ColocPignoraticio CP" & _
''''        " Inner Join Producto Pro on Pro.cCtaCod = CP.cCtaCod " & _
''''        " where Pro.nPrdEstado IN ('" & pEst1 & "','" & pEst2 & "','" & pEst3 & "','" & pEst4 & "'," & _
''''        " '" & pEst5 & "','" & pEst6 & "','" & pEst7 & "') "
''''
''''
''''If Len(Trim(pBoveda)) > 0 Then
''''    tmpSql = tmpSql & " AND cAgeBoveda in ('" & pBoveda & "') "
''''End If
''''
''''Conecta.AbreConexion
''''Set tmpReg = Conecta.CargaRecordSet(tmpSql)
''''Conecta.CierraConexion
''''
''''If (tmpReg.BOF Or tmpReg.EOF) Then
''''    SalSumEst = Format(0, "#0.00")
''''Else
''''    SalSumEst = IIf(IsNull(tmpReg!suma) = True, 0, Format(tmpReg!suma, "#0.00"))
''''End If
''''tmpReg.Close
''''Set tmpReg = Nothing
''''End Function
''''Private Function DiaNroEst(ByVal pBoveda As String, ByVal pFecha As String, ByVal pCodOpe As String, _
''''    Optional ByVal pCodOpe2 As String = "", Optional ByVal pCodOpe3 As String = "", _
''''    Optional ByVal pCodOpe4 As String = "", Optional ByVal pCodOpe5 As String = "", _
''''    Optional ByVal pCodOpe6 As String = "", Optional ByVal pCodOpe7 As String = "") As Double
'''''Permite obtener el número de operaciones
'''''de la tabla de TransPrenda de acuerdo a la condición deseada
''''Dim tmpReg As New ADODB.Recordset
''''Dim tmpSql As String
''''Dim Co As DConecta
''''Set Co = New DConecta
''''
''''If Len(Trim(pBoveda)) > 0 Then  ' Boveda
''''
''''    tmpSql = "select Count(MC.cCtaCod) AS Cuenta"
''''    tmpSql = tmpSql & " from Mov M"
''''    tmpSql = tmpSql & " Inner Join MovCol MC on Mc.nMovNro = M.nMovNro"
''''    tmpSql = tmpSql & " Inner Join ColocPignoraticio CP on CP.cCtaCod = MC.cCtaCod"
''''    tmpSql = tmpSql & " where M.nMovFlag=0 and CP.cAgeBoveda in ('" & pBoveda & "')  "
''''    tmpSql = tmpSql & " and  datediff(day,convert(datetime,substring(M.cMovNro,1,8),103),'" & Format(pFecha, "YYYY/MM/DD") & "')=0"
''''    tmpSql = tmpSql & " and  MC.nCredEstado in ('" & pCodOpe & "','" & pCodOpe2 & "','" & pCodOpe3 & "','" & pCodOpe4 & "','" & pCodOpe5 & "','" & pCodOpe6 & "','" & pCodOpe7 & "')   "
''''
''''Else
''''
''''    tmpSql = "select Count(MC.cCtaCod) AS Cuenta"
''''    tmpSql = tmpSql & " from Mov M"
''''    tmpSql = tmpSql & " Inner Join MovCol MC on Mc.nMovNro = M.nMovNro"
''''    tmpSql = tmpSql & " where M.nMovFlag=0  "
''''    tmpSql = tmpSql & " and  datediff(day,convert(datetime,substring(M.cMovNro,1,8),103),'" & Format(pFecha, "YYYY/MM/DD") & "')=0"
''''    tmpSql = tmpSql & " and  MC.nCredEstado in ('" & pCodOpe & "','" & pCodOpe2 & "','" & pCodOpe3 & "','" & pCodOpe4 & "','" & pCodOpe5 & "','" & pCodOpe6 & "','" & pCodOpe7 & "')   "
''''
''''End If
''''Co.AbreConexion
''''Set tmpReg = Co.CargaRecordSet(tmpSql)
''''If (tmpReg.BOF Or tmpReg.EOF) Then
''''    DiaNroEst = Format(0, "#0")
''''Else
''''    DiaNroEst = IIf(IsNull(tmpReg!Cuenta) = True, 0, Format(tmpReg!Cuenta, "#0"))
''''End If
''''tmpReg.Close
''''Co.CierraConexion
''''Set tmpReg = Nothing
''''End Function
''''Private Function DiaSumCP(ByVal pBoveda As String, ByVal pFecha As Date, ByVal pNomReg As String, ByVal pCodOpe As String, _
''''    Optional ByVal pCodOpe2 As String = "", Optional ByVal pCodOpe3 As String = "", _
''''    Optional ByVal pCodOpe4 As String = "", Optional ByVal pCodOpe5 As String = "", _
''''    Optional ByVal pCodOpe6 As String = "", Optional ByVal pCodOpe7 As String = "") As Currency
''''Dim tmpReg As New ADODB.Recordset
''''Dim tmpSql As String
''''Dim Co As DConecta
''''Set Co = New DConecta
''''Co.AbreConexion
''''
''''
''''    tmpSql = " select sum(abs(CP." & pNomReg & ")) AS Suma"
''''    tmpSql = tmpSql & " from Mov M"
''''    tmpSql = tmpSql & " Inner Join MovCol MC on Mc.nMovNro = M.nMovNro"
''''    tmpSql = tmpSql & " Inner Join ColocPignoraticio CP on CP.cCtaCod = MC.cCtaCod"
''''    tmpSql = tmpSql & " where M.nMovFlag=0 "
''''    tmpSql = tmpSql & " and  datediff(day,convert(datetime,substring(cMovNro,1,8),103),'" & Format(pFecha, "YYYY/MM/DD") & "')=0"
''''    tmpSql = tmpSql & " and  MC.cOpeCod in ('" & pCodOpe & "','" & pCodOpe2 & "','" & pCodOpe3 & "','" & pCodOpe4 & "','" & pCodOpe5 & "','" & pCodOpe6 & "','" & pCodOpe7 & "')"
''''
''''If Len(Trim(pBoveda)) > 0 Then  ' Boveda
''''    tmpSql = tmpSql & " AND cp.cAgeBoveda in ('" & pBoveda & "')  "
''''End If
''''
''''Set tmpReg = Co.CargaRecordSet(tmpSql)
''''
''''If (tmpReg.BOF Or tmpReg.EOF) Then
''''    DiaSumCP = 0
''''Else
''''    DiaSumCP = IIf(IsNull(tmpReg!suma), 0, tmpReg!suma)
''''End If
''''tmpReg.Close
''''Set tmpReg = Nothing
''''Co.CierraConexion
''''End Function
''''
'''''Fin Cambio JHVP
''''
''''Public Function nRepo128043_SaldosResumen(ByVal pnOpeCod As Long, ByVal psAgencia As String, ByVal pdDesde As Date, ByVal pdHasta As Date, _
''''                                   ByVal psUsuarioRep As String, ByVal psBovedasSelect As String)
''''
'''''Modificado JHVP 06.02.2005
''''
''''Dim lsSQL As String
''''Dim lsSqlT As String
''''
''''Dim lsDesembolsos As String
''''Dim lsRenovaciones As String
''''Dim lsCancelaciones As String
''''Dim lsAmortizaciones As String
''''
''''Dim lrDataRep As ADODB.Recordset
''''Dim loDataRep As dColPFunciones
''''Dim lsCadImp As String
''''Dim lsCadBuffer As String
''''
''''Dim lnIndice As Long
''''Dim lnLineas As Integer
''''Dim lnPage As Integer
''''Dim lsOperaciones As String
''''
''''Dim Fecha() As String
''''Dim Matriz() As Double
''''Dim nCant As Integer
''''Dim I As Integer
''''Dim nSaldoFinal As Double
''''
''''lsDesembolsos = " in ('" & gColPOpeDesembolso & "','" & gColPOpeDesembolsoEFE & "') "
''''lsRenovaciones = " in ('" & gColPOpeRenovNorEFE & "','" & gColPOpeRenovMorEFE & "','" & gColPOpeRenovNorCHQ & "','" _
''''                & gColPOpeRenovMorCHQ & "','" & gColPOpeRenovNorEnOtCjEFE & "','" & gColPOpeRenovMorEnOtCjEFE & "' ) "
''''lsCancelaciones = " in ('" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" & gColPOpeCancelNorCHQ & "','" _
''''                & gColPOpeCancelMorCHQ & "','" & gColPOpeCanceNorEnOtCjEFE & "','" & gColPOpeCanceMorEnOtCjEFE & "' ) "
''''
''''lsAmortizaciones = " in ('" & gColPOpeAmortNorEFE & "','" & gColPOpeAmortNorCHQ & "','" & gColPOpeAmortNorEnOtCjEFE & "' ) "
''''
''''lsSQL = "SELECT C0.cFecha, isnull(C1.nMonto,0) as nDesembolsos, isnull(c2.nMonto,0) as nRenovados, isnull(C3.nMonto,0) as nCancelados, isnull(C4.nMonto,0) as nAmortizados "
''''lsSQL = lsSQL & " From ( "
''''
''''lsSQL = lsSQL & " Select convert(varchar(8), dfecha, 112) cFecha from FechaTmp('" & Format(pdHasta, gsFormatoFecha) & "') "
''''lsSQL = lsSQL & " ) C0 LEFT JOIN ("
''''
''''lsSQL = lsSQL & " SELECT substring(m.cMovNro,1,8) as cFecha,'DESEMBOLSOS' As cDescripcion, sum(c.nMontoCol) as nMonto "
''''lsSQL = lsSQL & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro "
''''lsSQL = lsSQL & " INNER JOIN Colocaciones c On  c.cCtaCod = mc.cCtaCod "
''''lsSQL = lsSQL & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = c.cCtaCod "
''''lsSQL = lsSQL & " WHERE m.cOpecod " & lsDesembolsos
''''lsSQL = lsSQL & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "YYYYMMdd") & "' AND '" & Format(pdHasta, "YYYYMMdd") & "' "
''''lsSQL = lsSQL & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' And m.nMovFlag <> " & gMovFlagExtornado & " "
''''
''''        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
''''            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
''''        End If
''''        If Len(Trim(psBovedasSelect)) > 0 Then
''''            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
''''        End If
''''
''''lsSQL = lsSQL & " group by substring(m.cMovNro,1,8) "
''''lsSQL = lsSQL & " ) C1 ON C0.cFecha=C1.cFecha"
''''
''''lsSQL = lsSQL & " Left Join ( "
''''lsSQL = lsSQL & " Select R.cFecha, 'RENOVADOS' As cDescripcion, sum(R.capital) as nMonto "
''''lsSQL = lsSQL & " From ( "
''''lsSQL = lsSQL & " SELECT  substring(m.cMovNro,1,8) as cFecha, "
''''lsSQL = lsSQL & " Capital = (Select ISNULL(SUM(nMonto),0) "
''''lsSQL = lsSQL & " From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod "
''''lsSQL = lsSQL & " and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodCapital & ") "
''''lsSQL = lsSQL & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro "
''''lsSQL = lsSQL & " INNER JOIN Colocaciones c On  c.cCtaCod = mc.cCtaCod "
''''lsSQL = lsSQL & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = mc.cCtaCod "
''''lsSQL = lsSQL & " WHERE m.cOpecod  " & lsRenovaciones
''''lsSQL = lsSQL & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "YYYYMMdd") & "' AND '" & Format(pdHasta, "YYYYMMdd") & "' "
''''lsSQL = lsSQL & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' And m.nMovFlag <> " & gMovFlagExtornado & " "
''''
''''        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
''''            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
''''        End If
''''        If Len(Trim(psBovedasSelect)) > 0 Then
''''            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
''''        End If
''''
''''lsSQL = lsSQL & " ) R "
''''lsSQL = lsSQL & " group by R.cFecha "
''''lsSQL = lsSQL & " )C2 ON C0.cFecha = C2.cFecha "
''''lsSQL = lsSQL & " Left Join ( "
''''lsSQL = lsSQL & " Select C.cFecha, 'CANCELADOS' As cDescripcion, sum(C.capital) as nMonto "
''''lsSQL = lsSQL & " From ( "
''''lsSQL = lsSQL & " SELECT  substring(m.cMovNro,1,8) as cFecha, "
''''lsSQL = lsSQL & " Capital = (Select ISNULL(SUM(nMonto),0) "
''''lsSQL = lsSQL & " From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod "
''''lsSQL = lsSQL & " and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod =" & gColPConceptoCodCapital & ") "
''''lsSQL = lsSQL & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro "
''''lsSQL = lsSQL & " INNER JOIN Colocaciones c On  c.cCtaCod = mc.cCtaCod "
''''lsSQL = lsSQL & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = mc.cCtaCod "
''''lsSQL = lsSQL & " Where m.cOpeCod " & lsCancelaciones
''''lsSQL = lsSQL & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "YYYYMMdd") & "' AND '" & Format(pdHasta, "YYYYMMdd") & "' "
''''lsSQL = lsSQL & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' And m.nMovFlag <> " & gMovFlagExtornado & " "
''''
''''        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
''''            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
''''        End If
''''        If Len(Trim(psBovedasSelect)) > 0 Then
''''            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
''''        End If
''''
''''lsSQL = lsSQL & " ) C "
''''lsSQL = lsSQL & " group by C.cFecha "
''''lsSQL = lsSQL & " ) C3 on C0.cFecha=C3.cFecha "
''''lsSQL = lsSQL & " Left Join ( "
''''lsSQL = lsSQL & " select A.cFecha, 'AMORTIZADOS' As cDescripcion, sum(A.capital) as nMonto "
''''lsSQL = lsSQL & " From ( "
''''lsSQL = lsSQL & " SELECT  substring(m.cMovNro,1,8) as cFecha, "
''''lsSQL = lsSQL & " Capital = (Select ISNULL(SUM(nMonto),0) "
''''lsSQL = lsSQL & " From MovColDet Where nMovNro = mc.nMovNro and cOpeCod = mc.cOpeCod "
''''lsSQL = lsSQL & " and cCtaCod = mc.cCtaCod and nNroCalen = mc.nNroCalen and nPrdConceptoCod = " & gColPConceptoCodCapital & ") "
''''lsSQL = lsSQL & " FROM Mov m INNER JOIN Movcol mc On  m.nMovNro = mc.nMovNro "
''''lsSQL = lsSQL & " INNER JOIN Colocaciones c On  c.cCtaCod = mc.cCtaCod "
''''lsSQL = lsSQL & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = mc.cCtaCod "
''''lsSQL = lsSQL & " Where m.cOpeCod " & lsAmortizaciones
''''lsSQL = lsSQL & " AND Substring (m.cMovNro,1,8) BETWEEN '" & Format(pdDesde, "YYYYMMdd") & "' AND '" & Format(pdHasta, "YYYYMMdd") & "' "
''''lsSQL = lsSQL & " AND Substring (mc.cCtaCod, 4,2) = '" & psAgencia & "' And m.nMovFlag <> " & gMovFlagExtornado & " "
''''
''''        If psUsuarioRep <> "<Consolidado>" Then       'Todos los Usuarios
''''            lsSQL = lsSQL & " AND Right(m.cMovNro, 4) = '" & psUsuarioRep & "' "
''''        End If
''''        If Len(Trim(psBovedasSelect)) > 0 Then
''''            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
''''        End If
''''
''''lsSQL = lsSQL & " ) A "
''''lsSQL = lsSQL & " group by A.cFecha "
''''lsSQL = lsSQL & " ) C4 "
''''lsSQL = lsSQL & "  ON C0.cFecha=C4.cFecha "
''''
''''lsSQL = lsSQL & " where c0.cFecha<='" & Format(pdHasta, "YYYYmmdd") & "' "
''''lsSQL = lsSQL & " order by C0.cFecha "
''''
''''
''''lsSqlT = " select Sum(nSaldo) nSaldo From Producto "
''''lsSqlT = lsSqlT & " where nprdestado in(" & gColPEstDesem & "," & gColPEstRenov & "," & gColPEstVenci & "," & gColPEstPRema & ") "
''''lsSqlT = lsSqlT & " and substring(cCtaCod, 4,2)='" & psAgencia & "'"
''''
''''
''''    Set loDataRep = New dColPFunciones
''''        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSqlT)
''''        If lrDataRep.BOF Then
''''            nSaldoFinal = 0
''''        Else
''''            nSaldoFinal = lrDataRep!nSaldo
''''        End If
''''        Set lrDataRep = Nothing
''''
''''        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
''''    Set loDataRep = Nothing
''''
''''    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
''''        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
''''        Exit Function
''''    Else
''''        nCant = lrDataRep.RecordCount
''''        I = 0
''''        ReDim Preserve Matriz(1 To nCant, 1 To 6) As Double
''''        ReDim Preserve Fecha(1 To nCant) As String
''''        Do While Not lrDataRep.EOF
''''            I = I + 1
''''            Fecha(I) = Mid(lrDataRep!cFecha, 7, 2) & "/" & Mid(lrDataRep!cFecha, 5, 2) & "/" & Mid(lrDataRep!cFecha, 1, 4)
''''            Matriz(I, 2) = lrDataRep!nDesembolsos
''''            Matriz(I, 3) = lrDataRep!nCancelados
''''            Matriz(I, 4) = lrDataRep!nRenovados
''''            Matriz(I, 5) = lrDataRep!nAmortizados
''''            lrDataRep.MoveNext
''''        Loop
''''
''''        Matriz(nCant, 6) = nSaldoFinal
''''        Matriz(nCant, 1) = Matriz(nCant, 6) - Matriz(nCant, 2) + Matriz(nCant, 3) + Matriz(nCant, 4) + Matriz(nCant, 5)
''''
''''        For I = nCant - 1 To 1 Step -1
''''            Matriz(I, 6) = Matriz(I + 1, 1)
''''            Matriz(I, 1) = Matriz(I, 6) - Matriz(I, 2) + Matriz(I, 3) + Matriz(I, 4) + Matriz(I, 5)
''''        Next
''''
''''        lnLineas = 0
''''        lnPage = 1
''''
''''        'CABECERA
''''        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE SALDOS DE PRENDARIO", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 145, "Agencia - " & psAgencia, pnOpeCod)
''''
''''        lnIndice = 0:  lnLineas = 7
''''
''''        For I = 1 To nCant
''''
''''            lnIndice = lnIndice + 1
''''            lnLineas = lnLineas + 1
''''
''''            'Cadena de Impresion
''''
''''            lsCadImp = lsCadImp & Fecha(I) & Space(1) & ImpreFormat(Matriz(I, 1), 10, 2)
''''            lsCadImp = lsCadImp & Space(1) & ImpreFormat(Matriz(I, 2), 10, 2)
''''            lsCadImp = lsCadImp & Space(1) & ImpreFormat(Matriz(I, 3), 10, 2)
''''            lsCadImp = lsCadImp & Space(1) & ImpreFormat(Matriz(I, 4), 10, 2)
''''            lsCadImp = lsCadImp & Space(1) & ImpreFormat(Matriz(I, 5), 10, 2)
''''            lsCadImp = lsCadImp & Space(1) & ImpreFormat(Matriz(I, 6), 10, 2) & Chr(10)
''''
''''            If lnIndice Mod 300 = 0 Then
''''                lsCadBuffer = lsCadBuffer & lsCadImp
''''                lsCadImp = ""
''''            End If
''''
''''            If lnLineas >= 55 Then
''''                lnPage = lnPage + 1
''''                lsCadImp = lsCadImp & Chr(12)
''''                lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE SALDOS DE PRENDARIO", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 145, "Agencia - " & psAgencia, pnOpeCod)
''''                lnLineas = 7
''''            End If
''''
''''        Next I
''''
''''
''''        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
''''
''''        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
''''
''''    End If
''''    nRepo128043_SaldosResumen = lsCadBuffer
''''End Function
''''
'''''Fin Cambio JHVP


Public Function nRepo128033_ListadoCredxDiasAtraso(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pnDiasAIni As Integer, ByVal pnDiasAFin As Integer, ByVal pdFecsis As Date, _
        ByVal psBovedasSelect As String, Optional lbHistorica As Boolean = False) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
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

lsEstados = gColPEstDesem & "," & gColPEstRenov & "," & gColPEstVenci & "," & gColPEstPRema
lsCondiciones = "( " & Str(pnDiasAIni) & " - " & Str(pnDiasAFin) & " ) "


'On Error GoTo dError

    lsSQL = " SELECT codigoant=isnull(R.cctacodAnt,'                  '),P.cCtaCod, c.dVigencia, c.nPlazo, P.nSaldo, P.nTasaInteres, C.dVenc,  " _
        & " cp.nTasacion, cp.nOroNeto, cp.cAgeBoveda,  p.nPrdEstado,  " _
        & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " _
        & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = " & 1 & " ),  " _
        & " NomCliente = (Select ISNULL(MAX(Per.cPersNombre),'') From ProductoPersona PP INNER JOIN Persona Per " _
        & "                         On PP.cPersCod = Per.cPersCod Where PP.cCtaCod = c.cCtaCod )  " _
        & " FROM Producto P INNER JOIN Colocaciones c On  p.cCtaCod = c.cCtaCod " _
        & " INNER JOIN ColocPignoraticio cp On  cp.cCtaCod = c.cCtaCod " _
        & " LEFT JOIN RELCUENTAS R ON R.CCTACOD=P.CCTACOD " _
        & " WHERE P.nPrdEstado in ( " & lsEstados & " )  " _
        & " AND Substring (p.cCtaCod, 4,2) = '" & psAgencia & "' " _
        & " AND Datediff(dd, c.dVenc, '" & Format(pdFecsis, "mm/dd/yyyy") & "') > " & pnDiasAIni & " " _
        & " AND Datediff(dd, c.dVenc, '" & Format(pdFecsis, "mm/dd/yyyy") & "') < " & pnDiasAFin & " "
        If Len(Trim(psBovedasSelect)) > 0 Then
            lsSQL = lsSQL & " AND Right(cp.cAgeBoveda,2) in " & psBovedasSelect & " "
        End If
        lsSQL = lsSQL & " ORDER BY p.cCtaCod "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CREDITOS X DIAS DE ATRASO", lsCondiciones, lnPage, 145, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                
                
                vIntDev = 0:    vIntAct = 0:    vIntPro = 0:    vIntAde = 0
                vAtraso = 0:    vDiaAct = 0:    vDiaPro = 0
                vFecVen = Format(!dVenc, "dd/mm/yyyy")

                If DateDiff("d", vFecVen, pdFecsis) > 0 Then
                    vIntDev = ((1 + !nTasaIntVenc) ^ (DateDiff("d", vFecVen, pdFecsis) / 30) - 1) * !nSaldo
                    vIntDev = Round(vIntDev, 2)
                End If
                vAtraso = IIf(DateDiff("d", vFecVen, pdFecsis) > 0, DateDiff("d", vFecVen, pdFecsis), 0)
                If vAtraso = 0 Then
                    vDiaPro = IIf(DateDiff("d", pdFecsis, vFecVen) > 0, DateDiff("d", pdFecsis, vFecVen), 0)
                    vDiaAct = !nPlazo - vDiaPro
                    'vIntAde = CalculaInteresAdelantado(!nSaldoCap, !nTasaInt, !nPlazo)
                    vIntAde = !nSaldo * (1 - (1 / ((1 + !nTasaInteres) ^ (!nPlazo / 30))))
                    'vIntAde = Round(vIntAde, 2)
                    vIntAct = !nSaldo * ((((vIntAde / !nSaldo + 1) ^ (1 / !nPlazo)) ^ vDiaAct) - 1)
                    vIntAct = Round(vIntAct, 2)
                    vIntPro = (!nSaldo + vIntAct) * ((((vIntAde / !nSaldo + 1) ^ (1 / !nPlazo)) ^ vDiaPro) - 1)
                    vIntPro = Round(vIntPro, 2)
                End If
                
                    'vRTFImp = vRTFImp & ImpreFormat(vIndice, 8, 0) & Space(1) & FormatoContratro(!cCodCta) & Space(1) & _
                        Format(!dfecpres, "dd/mm/yyyy") & ImpreFormat(vNombre, 21, 1) & _
                        ImpreFormat(!nSaldoCap, 10) & ImpreFormat(!nvaltasac, 10) & _
                        ImpreFormat(!nPlazo, 3, 0) & _
                        ImpreFormat(vIntDev, 6) & ImpreFormat(vAtraso, 5, 0) & _
                        ImpreFormat(vIntAct, 9) & ImpreFormat(vIntPro, 9) & _
                        ImpreFormat(vIntAct + vIntPro, 9) & ImpreFormat(!nSaldoCap + vIntDev, 10) & _
                        ImpreFormat(!nOroNeto, 7) & pCorte

                'Asigno los totales
                'lnTotPrestamo = lnTotPrestamo + !nMontoCol
                lnTotSaldo = lnTotSaldo + !nSaldo
                lnTotTasacion = lnTotTasacion + !nTasacion
                lnTotOroNeto = lnTotOroNeto + !noroneto

                'Cadena de Impresion
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODIGOANT & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!NomCliente, 25, 0) & Space(1) _
                    & Format(!dVigencia, "dd/mm/yyyy") & Space(1) & ImpreFormat(!nTasacion, 10, 2) & Space(1) & ImpreFormat(!noroneto, 5, 2) & Space(1) _
                    & ImpreFormat(!nPlazo, 3, 0) & Space(1) & ImpreFormat(!nSaldo, 10, 2) & Space(1) _
                    & ImpreFormat(vAtraso, 5, 2) & Space(1) & ImpreFormat(vIntDev, 6, 2) & Space(1) & ImpreFormat(vIntAct, 8, 2) _
                    & Space(1) & ImpreFormat(vIntPro, 8, 2) & Space(1) & ImpreFormat(vIntAct + vIntPro, 8, 2) _
                    & Space(1) & ImpreFormat(!nSaldo + vIntDev, 10, 2) & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS", lsCondiciones, lnPage, 145, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(76) & ImpreFormat(lnTotPrestamo, 10, 2) & Space(1) _
            & ImpreFormat(lnTotSaldo, 10, 2) & Space(1) & ImpreFormat(lnTotTasacion, 10, 2) & Space(1) & ImpreFormat(lnTotOroNeto, 5, 2) & Space(1) _
            & Chr(10)
        
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo128033_ListadoCredxDiasAtraso = lsCadBuffer

End Function


Public Function nRepo128080_ListadoMovsxContrato(ByVal psPersCod As String, ByVal psPersNombre As String, _
        ByVal psPersDireccDomicilio As String, ByVal psTelefono As String, ByVal psDocumento As String, ByVal psNroContrato As String, _
        ByVal sDesde As String, ByVal sHasta As String, ByVal sConFecha As Byte) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Integer
Dim lnLineas As Integer
Dim nitem As Integer
Dim lnPage As Integer
Dim loImprimeCab As NColPImpre
        
Dim sOperacion As String * 30
Dim lsCondiciones As String
Dim sCriterio As String

If sConFecha = 1 Then
    sCriterio = " AND Substring (Mov.cMovNro,1,8) BETWEEN '" & Format(sDesde, "yyyymmdd") & "' AND '" & Format(sHasta, "yyyymmdd") & "' "
    lsCondiciones = "Del " & sDesde & " Al " & sHasta
End If

lsSQL = "SELECT MovCol.cCtaCod, SUBSTRING(Mov.cMovNro, 1, 8) AS Fecha, OpeTpo.cOpeDesc, MovCol.nPlazo,  " & _
        " MovCol.nMonto, SUBSTRING(Mov.cMovNro,18,2) AS Agencia, SUBSTRING(Mov.cMovNro, 22,4) AS Usuario  " & _
        " FROM MovCol INNER JOIN OpeTpo ON MovCol.cOpeCod = OpeTpo.cOpeCod INNER JOIN " & _
        " Mov ON MovCol.nMovNro = Mov.nMovNro " & _
        " WHERE (MovCol.cCtaCod = '" & Trim(psNroContrato) & "') " & _
        sCriterio & _
        " ORDER BY MovCol.cCtaCod, SUBSTRING(Mov.cMovNro, 1, 8) "
    
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        'MsgBox " No Existen Datos para el reporte", vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("MOVIMIENTOS DEL CONTRATO " & psNroContrato, lsCondiciones, lnPage, 130, "", 128034)

        lsCadImp = lsCadImp & "Codigo   : " & psPersCod & Chr(10) & _
                                "Nombre   : " & ImpreCarEsp(psPersNombre) & Chr(10) & _
                                "Direccion: " & ImpreCarEsp(psPersDireccDomicilio) & Chr(10) & _
                                "Telefono : " & psTelefono & Chr(10) & _
                                "Documento: " & psDocumento & Chr(10)
        
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(1) & "ITEM  CONTRATO            FECHA       OPERACION                       PLAZO            MONTO   AGENCIA USUARIO  " & Chr(10)
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        'lnLineas = lnLineas + 7

        lnIndice = 0:  lnLineas = 15
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nitem = nitem + 1
                
                sOperacion = lrDataRep!cOpeDesc
                
                
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(nitem, 4, 0) & Space(2) & Trim(psNroContrato) & Space(2) & Mid(lrDataRep!Fecha, 7, 2) & "/" & Mid(lrDataRep!Fecha, 5, 2) & _
                            "/" & Mid(lrDataRep!Fecha, 1, 4) & Space(2) & sOperacion & Space(2) & _
                            ImpreFormat(lrDataRep!nPlazo, 5, 0) & ImpreFormat(lrDataRep!nMonto, 14, 2) & Space(6) & lrDataRep!Agencia & Space(3) & lrDataRep!Usuario & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("MOVIMIENTOS DEL CONTRATO " & psNroContrato, lsCondiciones, lnPage, 130, "", 128034)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                .MoveNext
            Loop
        End With
            lsCadImp = lsCadImp & String(130, "-") & Chr(10)
                   
            lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
    nRepo128080_ListadoMovsxContrato = lsCadBuffer
End Function



Public Function nRepo128035_ListadoContratosxCliente(ByVal sCodAgen As String, ByVal psPersCod As String, ByVal psPersNombre As String, _
        ByVal psPersDireccDomicilio As String, ByVal psTelefono As String, ByVal psDocumento As String, ByVal psEstados As String, _
        ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal sConFecha As Byte) As String

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim nitem As Integer
Dim lnPage As Integer
Dim loImprimeCab As NColPImpre
        
Dim lsCondiciones As String
Dim sCriterio As String
Dim sMoneda As String * 3
'No total porque pueden ser en varias monedas

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
    
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        'MsgBox " No Existen Datos para el reporte", vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1
        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS POR CLIENTE ", lsCondiciones, lnPage, 145, "", 128035)

        lsCadImp = lsCadImp & "Codigo   : " & psPersCod & Chr(10) & _
                                "Nombre   : " & ImpreCarEsp(psPersNombre) & Chr(10) & _
                                "Direccion: " & ImpreCarEsp(psPersDireccDomicilio) & Chr(10) & _
                                "Telefono : " & psTelefono & Chr(10) & _
                                "Documento: " & psDocumento & Chr(10)
        
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(1) & "ITEM  CONTRATO            FECHA PRESTAMO          ULTIMA MODIFICACION                  MONTO  ESTADO" & Chr(10)
        lsCadImp = lsCadImp & String(145, "-") & Chr(10)

        lnIndice = 0:  lnLineas = 15
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                nitem = nitem + 1
                                 
                If CLng(Mid(!cCtaCod, 9, 1)) = gMonedaNacional Then
                    sMoneda = "S/."
                Else
                    sMoneda = "$"
                End If
                
                       
                lsCadImp = lsCadImp & Space(1) & ImpreFormat(nitem, 4, 0) & Space(2) & Trim(!cCtaCod) & Space(2) & UCase(!cFechaPrestamo) & _
                            Space(2) & Format(!cFechaHora, "dd/MM/YYYY") & " " & UCase(Format(!cFechaHora, "hh:mm:ss AMPM")) & Space(2) & _
                            sMoneda & Space(2) & ImpreFormat(!nMontoCol, 10, 2) & Space(2) & !sEstado & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS POR CLIENTE ", lsCondiciones, lnPage, 145, "", 128035)
                    lnLineas = 7
                    lnIndice = lnIndice + 7
                End If
                .MoveNext
            Loop
        End With
            lsCadImp = lsCadImp & String(145, "-") & Chr(10)
                   
            lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
    nRepo128035_ListadoContratosxCliente = lsCadBuffer
End Function

