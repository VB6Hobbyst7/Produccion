VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMGasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private MatGastos() As String
Private nNumMatGastos As Integer
Dim bCumpleEspecifGastoEstCta As Boolean 'By Capi 28122007

Dim nCuotaAplica1217() As String 'madm 20100903
Dim nMontoPrima As Double 'BRGO 20111205

'------ madm 09112009 ------
Private Function getValor1224(ByVal pnPlazo As Integer, ByVal pValor As Integer) As String
       
   Dim diapas1 As Integer
   Dim nu As Integer
   Dim nu1 As Integer
   
       getValor1224 = ""
       diapas1 = pnPlazo Mod 7
       nu = (pnPlazo - diapas1) / 7
              
              If pnPlazo = 0 Then
                    getValor1224 = Format(pValor, "#0.00")
              ElseIf pnPlazo <= 28 Then
                                      
                    If (nu = 4) Or (nu = 3 And diapas1 > 0) Then
                               getValor1224 = Format(pValor, "#0.00")
                    ElseIf (nu = 3) Or (nu = 2 And diapas1 > 0) Then
                               getValor1224 = Format(pValor / 4 * 3, "#0.00")
                    ElseIf (nu = 2) Or (nu = 1 And diapas1 > 0) Then
                               getValor1224 = Format(pValor / 2, "#0.00")
                    ElseIf (nu = 1) Or (nu = 0) Then
                               getValor1224 = Format(pValor / 4, "#0.00")
                    End If
              ElseIf pnPlazo > 28 And pnPlazo <= 31 Then
                    getValor1224 = Format(pValor, "#0.00")
              ElseIf (pnPlazo > 31) Then
                    getValor1224 = Format(pValor * (pnPlazo - (pnPlazo Mod 30)) / 30 + (pValor / 4) * IIf((pnPlazo Mod 30) = 0, 0, IIf((pnPlazo Mod 30) >= 7, ((pnPlazo Mod 30) - ((pnPlazo Mod 30) Mod 7)) / 7, 0)), "#0.00")
              End If
End Function
'ALPA 20100607 B2 ********************************
Private Function getAplicaGastoACuotas(psCtaCod As String, pnPrdConceptoCod As Integer, pcInstitucion As String, RFiltro As ADODB.Recordset, RFiltroInst As ADODB.Recordset, RFiltroGar As ADODB.Recordset, RGar As ADODB.Recordset, Optional ByVal psAgeCodAct As String = "", Optional ByVal psTpoProdCod As String = "", Optional ByVal psTpoCredCod As String = "") As Boolean
'**************************************************
Dim bAplicaGasto As Boolean
    bAplicaGasto = False
    RFiltro.Filter = " nPrdConceptoCod = " & pnPrdConceptoCod & " AND cTpoFiltro = 'P'"
    If RFiltro.RecordCount = 0 Then
       bAplicaGasto = True
    Else
        'ALPA 20100607 B2***************************************
        'If Mid(psCtaCod, 6, 3) = "301" Then
        If psTpoCredCod = "751" Then
        '*******************************************************
            RFiltro.Filter = " nPrdConceptoCod = " & pnPrdConceptoCod & " AND cTpoFiltro = 'P'"
            If RFiltro.RecordCount = 0 Then
                bAplicaGasto = True
            Else
                'RFiltroInst.Filter = " nPrdConceptoCod = " & pnPrdConceptoCod & " And nProdCod = " & Mid(psCtaCod, 6, 3) & " AND cAgeCod = '" & Mid(psCtaCod, 4, 2) & "' AND cTpoFiltro = 'P' AND cIntitucion='" & pcInstitucion & "'"
                RFiltroInst.Filter = " nPrdConceptoCod = " & pnPrdConceptoCod & " And nProdCod = " & psTpoCredCod & " AND cAgeCod = '" & psAgeCodAct & "' AND cTpoFiltro = 'P' AND cIntitucion='" & pcInstitucion & "'"
                If RFiltroInst.RecordCount > 0 Then
                    bAplicaGasto = True
                Else
                    bAplicaGasto = False
                End If
            End If
        Else
            'ALPA 20100607 B2*******************************************
            'RFiltro.Filter = " nPrdConceptoCod = " & pnPrdConceptoCod & " And nProdCod = " & Mid(psCtaCod, 6, 3) & " AND cAgeCod = '" & Mid(psCtaCod, 4, 2) & "' AND cTpoFiltro = 'P'"
            RFiltro.Filter = " nPrdConceptoCod = " & pnPrdConceptoCod & " And nProdCod = " & psTpoCredCod & " AND cAgeCod = '" & psAgeCodAct & "' AND cTpoFiltro = 'P'"
            '***********************************************************
            
            If RFiltro.RecordCount > 0 Then
                bAplicaGasto = True
            Else
                bAplicaGasto = False
            End If
        End If
    End If
             
    '************************************************************
    'Si tiene Filtro Por Tipos de Garantias
    '************************************************************
    If bAplicaGasto Then
        bAplicaGasto = False
        RFiltro.Filter = " nPrdConceptoCod = " & pnPrdConceptoCod & " AND cTpoFiltro = 'G'"
        If RFiltro.RecordCount = 0 Then
            bAplicaGasto = True
        End If
        
        If Not bAplicaGasto Then
            RGar.MoveFirst
            Do While Not RGar.EOF
                If Not bAplicaGasto Then
                    'ALPA 20100607 B2*********************************************************************************
                    'RFiltro.Filter = " nPrdConceptoCod = " & pnPrdConceptoCod & " And nProdCod = " & RGar!nTpoGar & " AND cAgeCod = '" & Mid(psCtaCod, 4, 2) & "' AND cTpoFiltro = 'G'"
                    RFiltro.Filter = " nPrdConceptoCod = " & pnPrdConceptoCod & " And nProdCod = " & RGar!nTpoGar & " AND cAgeCod = '" & psAgeCodAct & "' AND cTpoFiltro = 'G'"
                    '*************************************************************************************************
                    If RFiltro.RecordCount > 0 Then
                        bAplicaGasto = True
                        Exit Do
                    Else
                        bAplicaGasto = False
                    End If
                End If
                RGar.MoveNext
            Loop
        End If
    End If
    
    getAplicaGastoACuotas = bAplicaGasto
End Function

'ARCV 14-11-2006
Private Function RedimensionaMatriz(ByVal MatDatos As Variant, _
                                    ByVal nNumDatos As Integer) As Variant
Dim MatTempo As Variant
Dim i As Integer

ReDim MatTempo(nNumDatos, 5)

For i = 0 To nNumDatos - 1
    MatTempo(i, 0) = MatDatos(i, 0) 'Cuota...Aplicado
    MatTempo(i, 1) = MatDatos(i, 1) 'Num Gasto
    MatTempo(i, 2) = MatDatos(i, 2) 'Descripcion + Codigo
    MatTempo(i, 3) = MatDatos(i, 3) 'Valor
    MatTempo(i, 4) = MatDatos(i, 4) 'Valor
Next i
RedimensionaMatriz = MatTempo
End Function

'DAOR 20061215, Se modificó para que tome en cuenta gastos a la primera cuota
Private Function FormateaMatrizGastos(ByVal MatGastos As Variant, ByVal nNumGastos As Integer) As Variant
Dim i As Integer
Dim k As Integer
Dim MatTempo As Variant
Dim MatIndices() As Integer

Dim nPrdConceptoCod As Double
Dim MatPromedio() As Double
Dim MatPrdConceptoCod() As Double
Dim J As Integer
Dim nNumGastosporConcepto As Integer

    k = 0
    If nNumGastos = 0 Then
        FormateaMatrizGastos = MatGastos
        Exit Function
    End If

    ReDim MatPrdConceptoCod(1)
    MatPrdConceptoCod(0) = CInt(Trim(Right(MatGastos(0, 2), 8)))
    nPrdConceptoCod = CInt(Trim(Right(MatGastos(0, 2), 8)))
    For i = 0 To nNumGastos - 1
        If CInt(Trim(Right(MatGastos(i, 2), 8))) <> nPrdConceptoCod Then 'ARCV 07-07-2006
        'If BuscarGastosRepetidos(CStr(nPrdConceptoCod), MatGastos, i) = False Then
            ReDim Preserve MatPrdConceptoCod(UBound(MatPrdConceptoCod) + 1)
            MatPrdConceptoCod(UBound(MatPrdConceptoCod) - 1) = CInt(Trim(Right(MatGastos(i, 2), 8)))
            nPrdConceptoCod = CInt(Trim(Right(MatGastos(i, 2), 8)))
        End If
    Next i

    ReDim MatPromedio(UBound(MatPrdConceptoCod))

    For i = 0 To UBound(MatPrdConceptoCod) - 1
        MatPromedio(i) = 0
        nNumGastosporConcepto = 0
        For J = 0 To nNumGastos - 1
            If CInt(Trim(Right(MatGastos(J, 2), 8))) = MatPrdConceptoCod(i) Then
                MatPromedio(i) = MatPromedio(i) + MatGastos(J, 3)
                nNumGastosporConcepto = nNumGastosporConcepto + 1
            End If
        Next J
        MatPromedio(i) = MatPromedio(i) / nNumGastosporConcepto
    Next i

    '**CAPI 20080104 ************************************
    Dim nMaxCuotasSegunGastos As Integer
    Dim nTempoMax As Integer

    nMaxCuotasSegunGastos = 0
    
    For i = 0 To UBound(MatPrdConceptoCod) - 1
        nTempoMax = getNumGastosDePrdConcepto(CInt(MatPrdConceptoCod(i)), MatGastos)
        If nTempoMax > nMaxCuotasSegunGastos Then
            nMaxCuotasSegunGastos = nTempoMax
        End If
    Next i
    '****************************************************
    
    For i = 0 To nNumGastos - 1
        For J = 0 To UBound(MatPrdConceptoCod) - 1
            '**Modificado por CAPI 20080104 *********************
            'If (Trim(Right(MatGastos(i, 0), 2)) = "0" And CInt(Trim(Right(MatGastos(i, 2), 8))) = MatPrdConceptoCod(J)) _
            '    Or ((Trim(Right(MatGastos(i, 0), 2)) = "1" And _
            '    CInt(Trim(Right(MatGastos(i, 2), 8))) = MatPrdConceptoCod(J) And (CDbl(MatGastos(i, 3)) <> MatPromedio(J) Or (CDbl(MatGastos(i, 3)) = MatPromedio(J)) And CInt(MatGastos(i, 1)) = 1))) Then
            'COMENTADO POR APRI20180925 R-2018
            'If (Trim(Right(MatGastos(i, 0), 2)) = "0" And CInt(Trim(Right(MatGastos(i, 2), 8))) = MatPrdConceptoCod(J)) _
            '    Or (Trim(Right(MatGastos(i, 0), 2)) = "1" And _
            '    CInt(Trim(Right(MatGastos(i, 2), 8))) = MatPrdConceptoCod(J) And (CDbl(MatGastos(i, 3)) <> MatPromedio(J) Or (CDbl(MatGastos(i, 3)) = MatPromedio(J) And (CInt(MatGastos(i, 1)) = 1 Or getNumGastosDePrdConcepto(CInt(Trim(Right(MatGastos(i, 2), 8))), MatGastos) <> nMaxCuotasSegunGastos)))) Then
            'APRI20180924 ERS061-2018
            If (Trim(Right(MatGastos(i, 0), 2)) = "0" And CInt(Trim(Right(MatGastos(i, 2), 8))) = MatPrdConceptoCod(J)) _
                Or (Trim(Right(MatGastos(i, 0), 2)) = "1" And _
                CInt(Trim(Right(MatGastos(i, 2), 8))) = MatPrdConceptoCod(J) And (Round(CDbl(MatGastos(i, 3)), 2) <> Round(MatPromedio(J), 2) Or (Round(CDbl(MatGastos(i, 3)), 2) = Round(MatPromedio(J), 2) And (CInt(MatGastos(i, 1)) = 1 Or getNumGastosDePrdConcepto(CInt(Trim(Right(MatGastos(i, 2), 8))), MatGastos) <> nMaxCuotasSegunGastos))) Or CInt(Trim(Right(MatGastos(i, 2), 8))) = 1272) Then
            'Lógica
            'If (x = "0" And x = y) _
            '   Or (x = "1" And x = y And (x <> y Or (x = y And (x = 1 Or x <> y))) Or x = 1272) Then
            'END APRI
                k = k + 1
                ReDim Preserve MatIndices(k + 1)
                MatIndices(k - 1) = i
            End If
        Next J
    Next i

    ReDim MatTempo(k, 4)
    For i = 0 To k - 1
        For J = 0 To UBound(MatPrdConceptoCod) - 1
            MatTempo(i, 0) = MatGastos(MatIndices(i), 0)  'Aplicado
            If CInt(Trim(Right(MatGastos(MatIndices(i), 2), 8))) = MatPrdConceptoCod(J) Then
                'If CDbl(MatGastos(MatIndices(i), 3)) = MatPromedio(J) Then
                If Round(CDbl(MatGastos(MatIndices(i), 3)), 2) = Round(MatPromedio(J), 2) Then 'APRI20180924 ERS061-2018
                    '**Modificado por CAPI 20080104
                    'If getNumGastosDePrdConcepto(CInt(Trim(Right(MatGastos(i, 2), 8))), MatGastos) > 1 Then
                    'If getNumGastosDePrdConcepto(CInt(Trim(Right(MatGastos(MatIndices(i), 2), 8))), MatGastos) = nMaxCuotasSegunGastos Then
                    'Solucion
                    'If getNumGastosDePrdConcepto(CInt(Trim(Right(MatGastos(MatIndices(i), 2), 8))), MatGastos) = nMaxCuotasSegunGastos And CInt(Trim(Right(MatGastos(MatIndices(i), 2), 8))) <> 1217 Then
                    If getNumGastosDePrdConcepto(CInt(Trim(Right(MatGastos(MatIndices(i), 2), 8))), MatGastos) = nMaxCuotasSegunGastos And CInt(Trim(Right(MatGastos(MatIndices(i), 2), 8))) <> 1217 And CInt(Trim(Right(MatGastos(MatIndices(i), 2), 8))) <> 1272 Then 'APRI20180924 ERS061-2018
                       MatTempo(i, 1) = "*"
                    Else
                        MatTempo(i, 1) = MatGastos(MatIndices(i), 1) 'Numero
                    End If
                Else
                    MatTempo(i, 1) = MatGastos(MatIndices(i), 1) 'Numero
                End If
                MatTempo(i, 2) = MatGastos(MatIndices(i), 2) 'Gasto
                MatTempo(i, 3) = MatGastos(MatIndices(i), 3) 'Monto
            End If
        Next J
    Next i

    FormateaMatrizGastos = MatTempo
End Function

Private Function getNumGastosDePrdConcepto(nPrdConcepto As Integer, MatGastos As Variant) As Integer
Dim i As Integer
Dim acum As Integer
    acum = 0
    For i = 0 To UBound(MatGastos) - 1
        If CInt(Trim(Right(MatGastos(i, 2), 8))) = nPrdConcepto Then
            acum = acum + 1
        End If
    Next i
    getNumGastosDePrdConcepto = acum
End Function

'madm 20111206 - ALPA 20091202*******************************
'Private Sub GetChangeDatos(ByVal R As ADODB.Recordset, ByRef nMontoPrimaTotal As Double, ByVal psCtaCod As String, ByRef nValorUltCuota As Double, ByRef nTotalCuotas As Integer, ByRef nValorCuota As Double, ByVal MatCalend As Variant)
'Dim valor As Double
'valor = 0
'If R.RecordCount > 0 Then
'    If R!iTipoPago = 2 Then
'        If Mid(psCtaCod, 9, 1) = 1 And R!nEstado <> 1 Then
'            nMontoPrimaTotal = R!nMontoPrimaTotTC
'            nValorCuota = Round(nMontoPrimaTotal / 12, 2)
'            valor = ValorMinimoPrima(102740)
'            nValorCuota = IIf(nValorCuota < valor, valor, nValorCuota)
'        Else
'            nMontoPrimaTotal = R!nMontoPrimaTotal
'            nValorCuota = Round(nMontoPrimaTotal / 12, 2)
'            valor = ValorMinimoPrima(102741)
'            nValorCuota = IIf(nValorCuota < valor, valor, nValorCuota)
'        End If
'        nValorUltCuota = nValorCuota
'        nTotalCuotas = UBound(MatCalend)
'    End If
'End If
'End Sub

'WIOR 20120319 ***
Private Sub GetChangeDatos2(ByRef nMontoPrimaTotal As Double, ByVal psCtaCod As String, ByRef nValorUltCuota As Double, ByRef nTotalCuotas As Integer, ByRef nValorCuota As Double, ByVal MatCalend As Variant)

Dim R As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String

sSql = "stp_sel_MontoPagarPoliza '" & psCtaCod & "'"
Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        If Not R.EOF Then
            nMontoPrimaTotal = IIf(IsNull(R!nMontoPrimaTotTC), 0, R!nMontoPrimaTotTC)
            nValorCuota = IIf(IsNull(R!MontoPagar), 0, R!MontoPagar)
        End If
        nValorUltCuota = nValorCuota
        nTotalCuotas = UBound(MatCalend)
    End If
    oCon.CierraConexion
    Set oCon = Nothing
    nTotalCuotas = UBound(MatCalend) 'EJVG20150702
End Sub
'***

'DAOR 20061214, Se consideran nuevos criterios (Restricciones de Gastos):
Public Function GeneraCalendarioGastos(ByRef MatCalend As Variant, _
    ByRef MatDesemb As Variant, ByRef nNumGastos As Integer, ByVal pdHoy As Date, _
    ByVal psCtaCod As String, ByVal pnColocCred As Integer, ByVal psAplicaProceso As String, _
    Optional ByVal psGastoFijo As String = "F", Optional ByVal pnCuotaApr As Double = -1, _
    Optional ByVal pnPrestamo As Double = -1, Optional ByVal pnMontoCapital As Double = 0, _
    Optional ByVal pnCuotaPend As Integer = -1, Optional ByVal pHabRecorsed As Boolean = False, _
    Optional ByVal pR As ADODB.Recordset = Nothing, Optional ByVal pRFiltro As ADODB.Recordset = Nothing, _
    Optional ByVal pRFiltroGar As ADODB.Recordset = Nothing, Optional pnDiasAtraso As Integer = -1, _
    Optional ByVal pbSoloMostrar As Boolean = False, Optional ByVal pnTipoPeriodo As Integer = 0, _
    Optional ByVal pnPlazo As Integer = 0, Optional ByVal pnExoSeguroDes As Integer = 0, _
    Optional ByVal psAgeCodAct As String = "", Optional ByVal psTpoProdCod As String = "", _
    Optional ByVal psTpoCredCod As String = "", Optional ByVal pdFechaSis As Date = "01/01/1900", _
    Optional ByVal psTipoSegDes As String = "", Optional ByVal pnValorInmueble As Double = 0, _
    Optional ByVal pdFecSim As Date, Optional pbLogicoBF As Integer = 0, Optional pdFechaBF As Date = CDate("1900-01-01"), _
    Optional ByRef pMatCalendSegDes As Variant, Optional ByVal pnCuotaMivienda As Integer = -1, _
    Optional ByVal pbPagoAnticipado As Boolean = False, Optional ByVal pnExoSegMYPE As Integer = 0) As Variant
    'WIOR 20120521 SE AGREGO EL PARAMETRO  pdFechaSis, MAVM 20120625 Se agrego el parametro psTipoSegDes (MIVivienda)
    'WIOR 20140825 AGREGO pMatCalendSegDes
    'LUCV20180424, Agregó pbPagoAnticipado
    'APRI20180801 ERS061-2018 AGREGÓ pnExoSegMYPE
    
Dim oGastos As COMDCredito.DCOMGasto
Dim R As ADODB.Recordset
Dim i As Integer
Dim RFiltro As ADODB.Recordset
Dim bAplicaGasto As Boolean
Dim nRes As Double
'Dim nGastos1224 As Double
Dim oCred As COMDCredito.DCOMCredito
Dim oNCred As COMNCredito.NCOMCredito
Dim nPrestamo, nCuotaApr, nMontoGarantia As Double
Dim oGar As COMDCredito.DCOMGarantia
Dim nDiasAtraso As Integer
Dim sGastoFijoVariable, sTipoRelacGarantia As String
Dim oCredRel As COMDCredito.UCOMCredRela
Dim RFiltroGar As ADODB.Recordset
Dim RGar As ADODB.Recordset
Dim bCumpleRestricc As Boolean
Dim nMontoTemp As Double
Dim RCredVig As ADODB.Recordset
Dim MatCalTemp As Variant
Dim MontoPrePagado As Double
Dim MatGastosExon() As String
Dim RGastosExon As ADODB.Recordset
Dim odGastos1 As COMDCredito.DCOMGasto
Dim cInstitucion As String
Dim nMontoConstruccion As Double
Dim nValor As Double
Dim bExisteMasDeUnTitular As Boolean
Dim dFecDesemb As Date
Dim nMunMesesPorDia, nNumMesesPorMes, nNumConCer As Integer
Dim nNumConMicr As Integer
Dim nNroProxCuota As Integer 'DAOR 20070411
Dim RTemp As ADODB.Recordset
Dim RFiltroInst As ADODB.Recordset
Dim dFecPivot As Date
Dim bEsParalelo As Boolean
Dim nMontoPrimaTotal As Double, nTotalCuotas As Integer, nValorCuota As Double, nValorUltCuota As Double
'madm 20100908 *********************************************
Dim lnAcumula As Integer, lnCuenta As Integer, lnLimite As Integer
Dim nAplicaRestriTit1217xUno As Boolean, nAplicaRestriTit1217xDos As Boolean
Dim bValidar As Boolean, bEsRecurrentePol As Boolean, bEsRefinanciadoPol As Boolean 'MADM 20110425
'BRGO 20111205
Dim nSumaAsegurada As Double
Dim oPol As COMDCredito.DCOMPoliza
Dim rsPol As ADODB.Recordset
Dim sNumPoliza As String
'END BRGO
'WIOR 20140822 ***********************
Dim MatPersSegDes As Variant, nEdadMinSegDes As Long, nEdadMaxSegDes As Long, nCoberMinSegDes As Double, nCoberMaxSegDes As Double
Call CargaValoreSegDes(psCtaCod, pdHoy, MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, nCoberMinSegDes, nCoberMaxSegDes, pbPagoAnticipado) 'APRI20181025 ADD pbPagoAnticipado ERS071-2018
'WIOR FIN ****************************
Dim bEnvioEstCta As Boolean, nCostoEnvioEstCta As Double 'JUEZ 20130529
    'Recuperando el Monto de prestamo y la cuota aprobada
    Set oCred = New COMDCredito.DCOMCredito
    Set oPol = New COMDCredito.DCOMPoliza 'BRGO 20111209
    If psAplicaProceso <> "CD" Then
        Set R = oCred.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
        If pnCuotaApr = -1 Then
            sGastoFijoVariable = IIf(IsNull(R!cTipoGasto), "F", "V")
            Set oNCred = New COMNCredito.NCOMCredito
            MatCalTemp = oNCred.RecuperaMatrizCalendarioInicial(psCtaCod)
            Set oNCred = Nothing
            nCuotaApr = Format(CDbl(MatCalTemp(0, 3)) + CDbl(MatCalTemp(0, 4)) + CDbl(MatCalTemp(0, 5)) + CDbl(MatCalTemp(0, 6)) + CDbl(MatCalTemp(0, 7)), "#0.00")
        Else
            sGastoFijoVariable = psGastoFijo
            nCuotaApr = pnCuotaApr
        End If
        If pnPrestamo = -1 Then
            nPrestamo = Format(R!nMonto, "#0.000000")
        Else
            nPrestamo = pnPrestamo
        End If
        If psAplicaProceso = "DE" Or psAplicaProceso = "RP" Then 'DAOR 20070413, se aumentó condicional RP (Reprogramación)
            dFecDesemb = CDate(MatDesemb(0, 0)) 'DAOR 20061219, Fecha de Desembolso
        End If
        If psAplicaProceso = "PA" Or psAplicaProceso = "CA" Or psAplicaProceso = "PP" Then
            ReDim MatDesemb(1, 1)
            MatDesemb(0, 0) = Format(R!nMonto, "#0.000000")
        End If
        
        R.Close
        Set R = Nothing
        '*** PEAC 20080815
        If psAplicaProceso = "SI" Then
            nDiasAtraso = 0
            nNumConCer = 0
            nNroProxCuota = 1
            bEsParalelo = False
        '*** FIN PEAC
        Else
'            Set R = oCred.RecuperaColocacCred(psCtaCod)
'            nDiasAtraso = IIf(IsNull(R!nDiasAtraso), 0, R!nDiasAtraso)
'            nNumConCer = IIf(IsNull(R!nNumConCer), 0, R!nNumConCer)
'            nNumConMicr = IIf(IsNull(R!nNumConMic), 0, R!nNumConMic)
'            nNroProxCuota = R!nNroProxCuota
'            'By Capi 28122007
'            bEsParalelo = IIf(R!nColocCondicion = 3, True, False)
'            bEsRecurrentePol = IIf(R!nColocCondicion = 2, True, False) 'MADM 20110425
'            bEsRefinanciadoPol = IIf(R!nColocCondicion = 4, True, False) 'MADM 20110425
'            R.Close
'            Set oCred = Nothing
            'MAVM 20121113
            ResetVariables psCtaCod, nDiasAtraso, nNumConCer, nNumConMicr, nNroProxCuota, bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol
            '***
        End If
        'MAVM 20120625 ***
        If psTipoSegDes <> "" Then
            DesgravCredHipot sTipoRelacGarantia, bExisteMasDeUnTitular, psTipoSegDes
        Else
            Set oCredRel = New COMDCredito.UCOMCredRela
            Call oCredRel.CargaRelacPersCred(psCtaCod)
            sTipoRelacGarantia = IIf(oCredRel.ExisteMasDeUnTitular, "M", "I")
            DesgravCredHipot sTipoRelacGarantia, oCredRel.ExisteMasDeUnTitular, psTipoSegDes, psCtaCod
            bExisteMasDeUnTitular = oCredRel.ExisteTitularYCodeudor(psCtaCod) 'WIOR 20130419 AGREGO psCtaCod
        End If
                
    Set oGar = New COMDCredito.DCOMGarantia
    nMontoGarantia = oGar.RecuperaMontoGarantiaCredito(psCtaCod)
    nMontoConstruccion = oGar.RecuperaMontoGarantiaCredito(psCtaCod, True)
    ObtieneValoresEnvioEstadoCta bEnvioEstCta, nCostoEnvioEstCta, psCtaCod 'JUEZ 20130529
            
    If psAplicaProceso <> "CA" And psAplicaProceso <> "PA" Then
        Set R = oGar.RecuperaPolizaParaGenerarGastos(psCtaCod)
        'WIOR 20120329
        Call GetChangeDatos2(nMontoPrimaTotal, psCtaCod, nValorUltCuota, nTotalCuotas, nValorCuota, MatCalend)
        '*** BRGO 20111205 ****************************
        If psTpoProdCod = "517" Then
            If R.RecordCount <> 0 Then nSumaAsegurada = R!nSumaAsegurada
            Set rsPol = oPol.RecuperaPolizasCuenta(psCtaCod)
            If rsPol.RecordCount > 0 Then sNumPoliza = IIf(rsPol.RecordCount > 0, rsPol!cNumPoliza, "")
            Set rsPol = Nothing
            Set oPol = Nothing
        End If
        '*** END BRGO ************************************
        'bValidar = Validar(bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol, nTotalCuotas, nValorUltCuota, psCtaCod)'EJVG20150702
        R.Close
        Set R = Nothing
    End If
    Set RGar = oGar.RecuperaGarantiaCredito(psCtaCod)
    Set oGar = Nothing
    Else
        nPrestamo = pnPrestamo
    End If
    
    If psAplicaProceso = "DE" Then
        pdHoy = CDate(MatCalend(UBound(MatCalend) - 1, 0))
        nNumGastos = 0
        ReDim MatGastos(10000, 4)
        Set oGastos = New COMDCredito.DCOMGasto
        Set R = oGastos.RecuperaGastosAplicablesDesembolso(CInt(Mid(psCtaCod, 9, 1)), , psAplicaProceso, sGastoFijoVariable)
        If psTpoCredCod <> "751" Then
            Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P", True)
            Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G", True)
        Else
            Set odGastos1 = New COMDCredito.DCOMGasto
            cInstitucion = odGastos1.ObtenerCodInstitucionByCredito(psCtaCod)
            Set odGastos1 = Nothing
            Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P", True, cInstitucion)
            Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G", True, cInstitucion)
        End If
        
        Set oGastos = Nothing
        'Aplicar Gastos al Desembolso
        If IsArray(MatDesemb) Then
            ReDim Preserve MatDesemb(UBound(MatDesemb), 3)
            Do While Not R.EOF
                For i = 0 To UBound(MatDesemb) - 1
                    bAplicaGasto = False
                    If RFiltro.RecordCount = 0 Then
                        bAplicaGasto = True
                    Else
                        RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'P'" & IIf(cInstitucion <> "", " AND cIntitucion='" & cInstitucion & "'", "")
                        If RFiltro.RecordCount = 0 Then bAplicaGasto = True
                        If Not bAplicaGasto Then
                            RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & psTpoCredCod & " AND cAgeCod = '" & psAgeCodAct & "' AND cTpoFiltro = 'P'" & IIf(cInstitucion <> "", " AND cIntitucion='" & cInstitucion & "'", "")
                            bAplicaGasto = IIf(RFiltro.RecordCount > 0, True, False)
                        Else
                            If psTpoCredCod = "751" Then
                                RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & psTpoCredCod & " AND cAgeCod = '" & psAgeCodAct & "' AND cTpoFiltro = 'P' AND cIntitucion='" & cInstitucion & "'"
                                bAplicaGasto = IIf(RFiltro.RecordCount > 0, True, False)
                            End If
                        End If
                    End If
                    'Si tiene Filtro Por Tipos de Garantias
                    If bAplicaGasto Then
                        bAplicaGasto = False
                        RFiltroGar.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'G'"
                        If RFiltroGar.RecordCount = 0 Then bAplicaGasto = True
                        If Not bAplicaGasto Then
                            Dim oAmpliado As COMDCredito.DCOMAmpliacion
                            Dim bAmpliado As Boolean
                            
                            Set oAmpliado = New COMDCredito.DCOMAmpliacion
                            bAmpliado = oAmpliado.ValidaCreditoaAmpliar(psCtaCod)
                            Set oAmpliado = Nothing
                            If bAmpliado = False Then RGar.MoveFirst
                            'End If 'WIOR 20140823 COMENTO
                            Do While Not RGar.EOF
                                If Not bAplicaGasto Then
                                    RFiltroGar.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & RGar!nTpoGar & " AND cAgeCod = '" & psAgeCodAct & "' AND cTpoFiltro = 'G'"
                                    If RFiltroGar.RecordCount > 0 Then
                                        bAplicaGasto = True
                                        Exit Do
                                    Else
                                        bAplicaGasto = False
                                    End If
                                End If
                                RGar.MoveNext
                            Loop
                        End If
                    End If
                    If bAplicaGasto Then
                        If CDbl(MatDesemb(i, 1)) >= R!nInicial And CDbl(MatDesemb(i, 1)) <= R!nFinal Then
                            'Si el gasto se aplica por relaciones de credito
                            If R!cFiltro = "I" Or R!cFiltro = "M" Then
                                If sTipoRelacGarantia = "I" Then 'Si el Credito es de Tipo Individual
                                    If R!cFiltro = "I" Then
                                        'Aplico a todos el mismo porcentaje
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                            'Filtramos por Edad
                                                bCumpleRestricc = False
                                                bCumpleRestricc = CumpleRestriccion(1, R, oCredRel, pdHoy) 'WIOR 20140823 - RESTRICCION 1
                                                'WIOR 20140823 COMENTO
                                                'If Not IsNull(R!nEdadOper) Then
                                                '    Select Case R!nEdadOper
                                                '        Case 2 'Mayor
                                                '            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then bCumpleRestricc = True
                                                '            'End If
                                                '        Case 3 'Mayor Igual
                                                '            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then bCumpleRestricc = True
                                                '            'End If
                                                '        Case 4 'Menor
                                                '            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then bCumpleRestricc = True
                                                '            'End If
                                                '        Case 5 'Menor Igual
                                                '            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then bCumpleRestricc = True
                                                '            'End If
                                                '        Case 6 'Igual a
                                                '            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then bCumpleRestricc = True
                                                '            'End If
                                                '    End Select
                                                'Else
                                                '    bCumpleRestricc = True
                                                'End If
                                                'WIOR 20140823 FIN DE COMENTARIO
                                                
                                                'Si es Fijo
                                                If bCumpleRestricc Then
                                                    If R!nTpoValor = 1 Then nMontoTemp = nMontoTemp + Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                                    'End If
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                        'End If
                                                        If R!cAplicaMonto = "D" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(i, 1)), "#0.000000")
                                                        'End If
                                                        If R!cAplicaMonto = "H" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        'End If
                                                        If R!cAplicaMonto = "T" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                                        'End If
                                                        
                                                        If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                                        'End If
                                                        If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                                        'End If
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                End If
                                            oCredRel.siguiente
                                        Loop
                                        MatDesemb(i, 2) = Format(nMontoTemp, "#0.000000")
                                    End If
                                Else 'Si es Mancomunada cobrar el porcentaje mancomunado al Titular y al Conyuge
                                     'y al resto el gasto de Individual
                                     If R!cFiltro = "M" Then
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                            If oCredRel.ObtenerValorRelac = gColRelPersConyugue Or _
                                                oCredRel.ObtenerValorRelac = gColRelPersTitular Then
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then bCumpleRestricc = True
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then bCumpleRestricc = True
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then bCumpleRestricc = True
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then bCumpleRestricc = True
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then bCumpleRestricc = True
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If
                                                'Si es Fijo
                                                If bCumpleRestricc Then
                                                    If R!nTpoValor = 1 Then nMontoTemp = nMontoTemp + Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                                    'End If 'WIOR 20140823 COMENTO
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                        If R!cAplicaMonto = "D" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(i, 1)), "#0.000000")
                                                        If R!cAplicaMonto = "H" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        If R!cAplicaMonto = "T" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                                        If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                                        If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                End If
                                            End If
                                            oCredRel.siguiente
                                        Loop
                                        MatDesemb(i, 2) = Format(nMontoTemp, "#0.000000")
                                     End If
                                     If R!cFiltro = "I" Then
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                            If oCredRel.ObtenerValorRelac <> gColRelPersConyugue And _
                                                oCredRel.ObtenerValorRelac <> gColRelPersTitular Then
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then bCumpleRestricc = True
                                                            'End If
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then bCumpleRestricc = True
                                                            'End If
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then bCumpleRestricc = True
                                                            'End If
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then bCumpleRestricc = True
                                                            'End If
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then bCumpleRestricc = True
                                                            'End If
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If
                                                'Si es Fijo
                                                If bCumpleRestricc Then
                                                    If R!nTpoValor = 1 Then nMontoTemp = nMontoTemp + Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                                    'End If
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                        'End If
                                                        If R!cAplicaMonto = "D" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(i, 1)), "#0.000000")
                                                        'End If
                                                        If R!cAplicaMonto = "H" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        'End If
                                                        If R!cAplicaMonto = "T" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                                        'End If
                                                        If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                                        'End If
                                                        If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                                        'End If
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                    
                                                End If
                                            End If
                                            oCredRel.siguiente
                                        Loop
                                        MatDesemb(i, 2) = Format(nMontoTemp, "#0.000000")
                                     End If
                                End If
                            Else 'Si no se considera las Relaciones entonces
                                'Filtramos por Edad
                                bCumpleRestricc = False
                                If Not IsNull(R!nEdadOper) Then
                                    Select Case R!nEdadOper
                                        Case 2 'Mayor
                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then bCumpleRestricc = True
                                            'End If
                                        Case 3 'Mayor Igual
                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then bCumpleRestricc = True
                                            'End If
                                        Case 4 'Menor
                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then bCumpleRestricc = True
                                            'End If
                                        Case 5 'Menor Igual
                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then bCumpleRestricc = True
                                            'End If
                                        Case 6 'Igual a
                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then bCumpleRestricc = True
                                            'End If
                                    End Select
                                Else
                                    bCumpleRestricc = True
                                End If
                                'Si es Fijo
                                If bCumpleRestricc Then
                                    If R!nTpoValor = 1 Then MatDesemb(i, 2) = Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.00")
                                    'End If
                                    'Si es Porcentaje
                                    If R!nTpoValor = 2 Then
                                        If R!cAplicaMonto = "C" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.00")
                                        'End If
                                        If R!cAplicaMonto = "D" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(i, 1)), "#0.00")
                                        'End If
                                        If R!cAplicaMonto = "H" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.00")
                                        'End If
                                        If R!cAplicaMonto = "T" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                        'End If
                                        If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                        'End If
                                        If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                        'End If
                                        MatDesemb(i, 2) = Format(nRes, "#0.000000")
                                    End If
                                End If
                            End If
                            nNumGastos = nNumGastos + 1
                            MatGastos(nNumGastos - 1, 0) = "Desembolso" & Space(50) & gColocCalendAplDesembolso
                            MatGastos(nNumGastos - 1, 1) = i + 1
                            MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod
                            
                            If R!nMontoMensual = "1" Then
                                MatGastos(nNumGastos - 1, 3) = Format(CDbl(MatDesemb(i, 2)) * UBound(MatCalend), "#0.00")
                            Else
                                MatGastos(nNumGastos - 1, 3) = MatCalend(i, 2)
                            End If
                            
                            If R!nAplicado = gColocConceptoAplDesembolso Then Exit For
                            'End If 'WIOR 20140823 COMENTO
                        End If
                    End If
                Next i
                R.MoveNext
            Loop
            R.Close
            Set R = Nothing
        End If

        Set oGastos = New COMDCredito.DCOMGasto
        Set R = oGastos.RecuperaGastosAplicablesCuotas(CInt(Mid(psCtaCod, 9, 1)), psAplicaProceso, sGastoFijoVariable)
        Set oGastos = Nothing
        dFecPivot = dFecDesemb
        Set oGastos = New COMDCredito.DCOMGasto
        If psTpoCredCod <> "751" Then
            Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P", False)
            Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G", True)
        Else
            Set odGastos1 = New COMDCredito.DCOMGasto
            cInstitucion = odGastos1.ObtenerCodInstitucionByCredito(psCtaCod)
            Set odGastos1 = Nothing
            Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P", False)
            Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G", True, cInstitucion)
            Set RFiltroInst = oGastos.RecuperaFiltroAplicadoCuenta("P", True, cInstitucion)
        End If
        Set RTemp = RFiltro.Clone   'ARCV 01-02-2007
        Do While Not R.EOF
        Set RFiltro = RTemp.Clone 'ARCV 01-02-2007

        If R!bAplValorDosTit Then 'Verificar si la restricción de Creditos con mas de un titular está activo
            nValor = IIf(bExisteMasDeUnTitular = True, IIf(IsNull(R!nValorDosTit), 0, R!nValorDosTit), IIf(IsNull(R!nValor), 0, R!nValor))
            'madm 20101130 - 20100903 ********************************************************************************
            If R!nPrdConceptoCod = 1217 Then nCuotaAplica1217 = GetCuotaAplicaExoneraSegDes(80, MatCalend, oCredRel, psCtaCod) 'WIOR 20130903 AGREGO psCtaCod
            'end madm **************************************************************************************
        Else
            nValor = IIf(IsNull(R!nValor), 0, R!nValor)
        End If
        'EJVG20150702 *** Si o si la ultima cuota me devolverá Monto Prima Poliza(Necesario para determinar variable bValor)
        'If R!nPrdConceptoCod = 1231 Or R!nPrdConceptoCod = 1243 Then 'LUCV20180601, Comentó según ERS022-2018
        If R!nPrdConceptoCod = 1243 Then 'LUCV20180601, Agregó según ERS022-2018
            Set oPol = New COMDCredito.DCOMPoliza
            nValorCuota = oPol.DamePrimaPolizaxCuota(psCtaCod, nTotalCuotas, nTotalCuotas, R!nPrdConceptoCod)
            nValorUltCuota = nValorCuota
            bValidar = Validar(bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol, nTotalCuotas, nValorUltCuota, psCtaCod)
            Set oPol = Nothing
        End If

        'END EJVG *******
            For i = 0 To UBound(MatCalend) - 1
                    bAplicaGasto = getAplicaGastoACuotas(psCtaCod, R!nPrdConceptoCod, cInstitucion, RFiltro, RFiltroInst, RFiltroGar, RGar, psAgeCodAct, psTpoProdCod, psTpoCredCod)
                    'MAVM 20120801 ***
                    'LUCV20180601, Agregó ERS022-2018. condición para el gasto en la sugerencia
                    If R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipoteca And MatCalend(i, 15) > 0 Then
                        bAplicaGasto = True
                    End If
                    'Fin LUCV20180601
                    
                    If psTpoCredCod = "853" Then If R!nPrdConceptoCod = "1243" Or R!nPrdConceptoCod = "1245" Or R!nPrdConceptoCod = "1246" Then bAplicaGasto = False
                    '***
                    If bAplicaGasto Then
                        If CDbl(MatDesemb(0, 1)) >= R!nInicial And CDbl(MatDesemb(0, 1)) <= R!nFinal Then '-> Filtro por Monto de Prestamo
                            If R!cFiltro = "I" Or R!cFiltro = "M" Then
                                If sTipoRelacGarantia = "I" Then 'Si el Credito es de Tipo Individual
                                    If R!cFiltro = "I" Then
                                        'Aplico a todos el mismo porcentaje
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then bCumpleRestricc = True
                                                            'End If'WIOR 20140822 COMENTO
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then bCumpleRestricc = True
                                                            'End If 'WIOR 20140822 COMENTO
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then bCumpleRestricc = True
                                                            'End If 'WIOR 20140822 COMENTO
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then bCumpleRestricc = True
                                                            'End If 'WIOR 20140822 COMENTO
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then bCumpleRestricc = True
                                                            'End If 'WIOR 20140822 COMENTO
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If
                                                
                                                If bCumpleRestricc Then
                                                'Si es Fijo
                                                    If R!nTpoValor = 1 Then
                                                         'By Capi 28122007  Acta 051-2007
                                                        If R!nPrdConceptoCod = gColocConceptoCodGastoEstCta Then
                                                            bCumpleEspecifGastoEstCta = obtenerEspecifGastoEstCta(MatCalend(i, 0), dFecPivot)
                                                            'By Capi 01022008 porque es en base al numero de cuota segun acta 010-2008
                                                            If bCumpleEspecifGastoEstCta = True And nTotalCuotas > 1 Then
                                                                dFecPivot = MatCalend(i, 0)
                                                                nMontoTemp = nMontoTemp + Format(nValor, "#0.000000")
                                                            End If
                                                        Else
                                                                nMontoTemp = nMontoTemp + Format(nValor, "#0.000000")
                                                        End If
                                                    End If
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then nRes = Format(Format(nValor / 100, "#0.000000") * CDbl(MatCalend(i, 2)), "#0.000000")
                                                        'End If
                                                        If R!cAplicaMonto = "D" Then nRes = Format(Format(nValor / 100, "#0.000000") * nPrestamo, "#0.000000")
                                                        'End If
                                                        If R!cAplicaMonto = "H" Then nRes = Format(Format(nValor / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        'End If
                                                        If R!cAplicaMonto = "S" Then nRes = Format(Format(nValor / 100, "#0.000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                                        'End If
                                                        If R!cAplicaMonto = "T" Then nRes = Format(Format(nValor / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                                        'End If
                                                        'DAOR 20061213, Para el caso del seguro de desgravamen
                                                        'JUEZ 20140310 ****************************************************
                                                        'If R!cAplicaMonto = "I" Then
                                                        '    If i = 0 Then
                                                        '        nRes = Format(Format(nValor / 100, "#0.000000") * (nPrestamo + CDbl(MatCalend(i, 4))), "#0.000000")
                                                        '    Else
                                                        '        nRes = Format(Format(nValor / 100, "#0.000000") * (CDbl(MatCalend(i - 1, 7)) + CDbl(MatCalend(i, 4))), "#0.000000")
                                                        '    End If
                                                        'End If
                                                        If R!cAplicaMonto = "I" Or R!cAplicaMonto = "K" Then
                                                            '->*****LUCV20180601, Comentó y agregó(El parámetro de retorno pMatCalendSegDes, se obtendra en el calendario)
                                                            'nRes = CalculaSeguroDesgravamen(i, R!cAplicaMonto, nValor, nPrestamo, CDbl(MatCalend(i, 4)), CDbl(MatCalend(IIf(i = 0, 1, i) - 1, 7)))'WIOR 20140822 COMENTO
                                                            'nRes = CalculaSeguroDesgravamen(i, R!cAplicaMonto, nValor, nPrestamo, CDbl(MatCalend(i, 4)), CDbl(MatCalend(IIf(i = 0, 1, i) - 1, 7)), CDate(MatCalend(i, 0)), CDbl(R!nValor), CDbl(R!nValorDosTit), MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, nCoberMinSegDes, nCoberMaxSegDes, pMatCalendSegDes, psCtaCod) 'WIOR 20140822 'JUEZ 20150714 psCtaCod
                                                            nRes = MatCalend(i, 8) 'El seguro desgravamen ya se incluye en el cálculo del calendario
                                                            '<-***** Fin LUCV20180601
                                                        End If
                                                        'END JUEZ *********************************************************
                                                        '-------------------------------------------------------
                                                        If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                                        'End If
                                                        If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                                        'End If
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                                                                     
                                                    '4**PEAC 20080110, Si es personalizada **********************
                                                    If R!nTpoValor = 3 Then
                                                        'ALPA20140312*************************
                                                        'MatCalend(i, 6) = ObtieneValorCuotaPoliza(psCtaCod, psTpoProdCod, i, nTotalCuotas, nValorUltCuota, nValorCuota, R!nPrdConceptoCod, pdHoy, sNumPoliza, nSumaAsegurada, bValidar, bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol, pdFechaSis)
                                                        'Se agregó pbLogicoBF, pdFechaBF, MatCalend(i, 0)
                                                        If R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipoteca Then 'LUCV20180601, Agregó según ERS022-2018
                                                            MatCalend(i, 6) = MatCalend(i, 15)
                                                        ElseIf R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipotecaGracia Then 'LUCV20180601, Agregó según ERS022-2018
                                                            MatCalend(i, 6) = MatCalend(i, 16)
                                                        Else
                                                            MatCalend(i, 6) = ObtieneValorCuotaPoliza(psCtaCod, psTpoProdCod, i, nTotalCuotas, nValorUltCuota, nValorCuota, R!nPrdConceptoCod, pdHoy, sNumPoliza, nSumaAsegurada, bValidar, bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol, pdFechaSis, pbLogicoBF, pdFechaBF, MatCalend(i, 0), MatCalend(i, 1), pnCuotaMivienda, pnPlazo, pnTipoPeriodo, pnExoSegMYPE) 'pnPlazo,pnTipoPeriodo APRI 20170227
                                                        'WIOR 20120521 SE AGREGO EL PARAMETRO pdFechaSis
                                                        'APRI20180821 ERS061-2018 AGREGÓ pnExoSegMYPE
                                                        End If
                                                    End If
                                                    '**********************************************************
                                                End If
                                            oCredRel.siguiente
                                        Loop
                                        MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                    Else
                                        Exit For
                                    End If
                                Else 'Si es Mancomunada cobrar el porcentaje mancomunado todos menos al Titular
                                     'y al resto el gasto de Mancomunado
                                     'Para desiguales al Titular
                                     If R!cFiltro = "M" Then
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                            If oCredRel.ObtenerValorRelac <> gColRelPersTitular Then
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then bCumpleRestricc = True
                                                            'End If
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then bCumpleRestricc = True
                                                            'End If
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then bCumpleRestricc = True
                                                            'End If
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then bCumpleRestricc = True
                                                            'End If
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then bCumpleRestricc = True
                                                            'End If
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If
                                            
                                                'Si es Fijo
                                                If bCumpleRestricc Then
                                                    If R!nTpoValor = 1 Then
                                                        'By Capi 28122007  Acta 051-2007
                                                        If R!nPrdConceptoCod = gColocConceptoCodGastoEstCta Then
                                                            bCumpleEspecifGastoEstCta = obtenerEspecifGastoEstCta(MatCalend(i, 0), dFecPivot)
                                                            'If bCumpleEspecifGastoEstCta = True And Not bEsParalelo Then
                                                            If bCumpleEspecifGastoEstCta = True And nTotalCuotas > 1 Then
                                                                dFecPivot = MatCalend(i, 0)
                                                                nMontoTemp = nMontoTemp + Format(nValor, "#0.000000")
                                                            End If
                                                        Else
                                                                nMontoTemp = nMontoTemp + Format(nValor, "#0.000000")
                                                        End If
                                                        'End By
                                                    End If
                                                   'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then nRes = Format(Format(nValor / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                        'End If
                                                        If R!cAplicaMonto = "D" Then nRes = Format(Format(nValor / 100, "#0.000000") * nPrestamo, "#0.000000")
                                                        'End If
                                                        If R!cAplicaMonto = "H" Then nRes = Format(Format(nValor / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        'End If
                                                        If R!cAplicaMonto = "T" Then nRes = Format(Format(nValor / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                                        'End If
                                                        If R!cAplicaMonto = "S" Then nRes = Format(Format(nValor / 100, "#0.000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                                        'End If
                                                        'DAOR 20061213, Para el caso del seguro de desgravamen
                                                        'JUEZ 20140310 ****************************************************
                                                        'If R!cAplicaMonto = "I" Then
                                                        '    If i = 0 Then
                                                        '        nRes = Format(Format(nValor / 100, "#0.000000") * (nPrestamo + CDbl(MatCalend(i, 4))), "#0.000000")
                                                        '    Else
                                                        '        nRes = Format(Format(nValor / 100, "#0.000000") * (CDbl(MatCalend(i - 1, 7)) + CDbl(MatCalend(i, 4))), "#0.000000")
                                                        '    End If
                                                        'End If
                                                        If R!cAplicaMonto = "I" Or R!cAplicaMonto = "K" Then
                                                            '->*****LUCV20180601, Comentó y agregó(El parámetro de retorno pMatCalendSegDes, se obtendra en el calendario)
                                                            'nRes = CalculaSeguroDesgravamen(i, R!cAplicaMonto, nValor, nPrestamo, CDbl(MatCalend(i, 4)), CDbl(MatCalend(IIf(i = 0, 1, i) - 1, 7))) 'WIOR 20140822 COMENTO
                                                            'nRes = CalculaSeguroDesgravamen(i, R!cAplicaMonto, nValor, nPrestamo, CDbl(MatCalend(i, 4)), CDbl(MatCalend(IIf(i = 0, 1, i) - 1, 7)), CDate(MatCalend(i, 0)), CDbl(R!nValor), CDbl(R!nValorDosTit), MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, nCoberMinSegDes, nCoberMaxSegDes, pMatCalendSegDes, psCtaCod) 'WIOR 20140822 'JUEZ 20150714 psCtaCod
                                                            nRes = MatCalend(i, 8) 'El seguro desgravamen ya se incluye en el cálculo del calendario
                                                            '<-***** Fin LUCV20180601
                                                        End If
                                                        'END JUEZ *********************************************************
                                                        If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                                        'End If
                                                        If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                                        'End If
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                    '3 **PEAC 20080110, Si es personalizada **********************
                                                    If R!nTpoValor = 3 Then
                                                        'BRGO 20111205
                                                        'ALPA20140312
                                                        'MatCalend(i, 6) = ObtieneValorCuotaPoliza(psCtaCod, psTpoProdCod, i, nTotalCuotas, nValorUltCuota, nValorCuota, R!nPrdConceptoCod, pdHoy, sNumPoliza, nSumaAsegurada, bValidar, bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol, pdFechaSis)
                                                        'Se agregó pbLogicoBF, pdFechaBF, MatCalend(i, 0)
                                                        If R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipoteca Then 'LUCV20180601, Agregó según ERS022-2018
                                                            MatCalend(i, 6) = MatCalend(i, 15)
                                                        ElseIf R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipotecaGracia Then 'LUCV20180601, Agregó según ERS022-2018
                                                            MatCalend(i, 6) = MatCalend(i, 16)
                                                        Else
                                                            MatCalend(i, 6) = ObtieneValorCuotaPoliza(psCtaCod, psTpoProdCod, i, nTotalCuotas, nValorUltCuota, nValorCuota, R!nPrdConceptoCod, pdHoy, sNumPoliza, nSumaAsegurada, bValidar, bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol, pdFechaSis, pbLogicoBF, pdFechaBF, MatCalend(i, 0), MatCalend(i, 1), pnCuotaMivienda, pnPlazo, pnTipoPeriodo, pnExoSegMYPE) 'pnPlazo,pnTipoPeriodo APRI 20170227
                                                            'WIOR 20120521 SE AGREGO EL PARAMETRO pdFechaSis
                                                            'APRI20180821 ERS061-2018 AGREGÓ pnExoSegMYPE
                                                        End If
                                                    End If
                                                    '***********************************************************
                                                    
                                                End If
                                            End If
                                            oCredRel.siguiente
                                        Loop
                                            MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                     End If
                                     
                                     If R!cFiltro = "I" Then
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                            If oCredRel.ObtenerValorRelac = gColRelPersTitular Then
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then bCumpleRestricc = True
                                                            'End If 'WIOR 20140822 COMENTO
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then bCumpleRestricc = True
                                                            'End If 'WIOR 20140822 COMENTO
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then bCumpleRestricc = True
                                                            'End If 'WIOR 20140822 COMENTO
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then bCumpleRestricc = True
                                                            'End If 'WIOR 20140822 COMENTO
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then bCumpleRestricc = True
                                                            'End If 'WIOR 20140822 COMENTO
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If
                                            
                                                'Si es Fijo
                                                If bCumpleRestricc Then
                                                    If R!nTpoValor = 1 Then
                                                        'By Capi 28122007  Acta 051-2007
                                                        If R!nPrdConceptoCod = gColocConceptoCodGastoEstCta Then
                                                            bCumpleEspecifGastoEstCta = obtenerEspecifGastoEstCta(MatCalend, dFecPivot)
                                                            'If bCumpleEspecifGastoEstCta = True And Not bEsParalelo Then
                                                            If bCumpleEspecifGastoEstCta = True And nTotalCuotas > 1 Then
                                                                dFecPivot = MatCalend(i, 0)
                                                                nMontoTemp = nMontoTemp + Format(nValor, "#0.000000")
                                                            End If
                                                        Else
                                                            nMontoTemp = nMontoTemp + Format(nValor, "#0.000000")
                                                        End If
                                                    End If
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then nRes = Format(Format(nValor / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                        If R!cAplicaMonto = "D" Then nRes = Format(Format(nValor / 100, "#0.000000") * nPrestamo, "#0.000000")
                                                        If R!cAplicaMonto = "H" Then nRes = Format(Format(nValor / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        If R!cAplicaMonto = "S" Then nRes = Format(Format(nValor / 100, "#0.0000000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                                        If R!cAplicaMonto = "T" Then nRes = Format(Format(nValor / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                                        
                                                        'DAOR 20061213, Para el caso del seguro de desgravamen
                                                        'JUEZ 20140310 **********************************************
                                                        'If R!cAplicaMonto = "I" Then
                                                        '    If i = 0 Then
                                                        '        nRes = Format(Format(nValor / 100, "#0.000000") * (nPrestamo + CDbl(MatCalend(i, 4))), "#0.000000")
                                                        '    Else
                                                        '        nRes = Format(Format(nValor / 100, "#0.000000") * (CDbl(MatCalend(i - 1, 7)) + CDbl(MatCalend(i, 4))), "#0.000000")
                                                        '    End If
                                                        'End If
                                                        If R!cAplicaMonto = "I" Or R!cAplicaMonto = "K" Then
                                                            '->*****LUCV20180601, Comentó y agregó(El parámetro de retorno pMatCalendSegDes, se obtendra en el calendario)
                                                            'nRes = CalculaSeguroDesgravamen(i, R!cAplicaMonto, nValor, nPrestamo, CDbl(MatCalend(i, 4)), CDbl(MatCalend(IIf(i = 0, 1, i) - 1, 7))) 'WIOR 20140822 COMENTO
                                                            'nRes = CalculaSeguroDesgravamen(i, R!cAplicaMonto, nValor, nPrestamo, CDbl(MatCalend(i, 4)), CDbl(MatCalend(IIf(i = 0, 1, i) - 1, 7)), CDate(MatCalend(i, 0)), CDbl(R!nValor), CDbl(R!nValorDosTit), MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, nCoberMinSegDes, nCoberMaxSegDes, pMatCalendSegDes, psCtaCod) 'WIOR 20140822 'JUEZ 20150714 psCtaCod
                                                            nRes = MatCalend(i, 8) 'El seguro desgravamen ya se incluye en el cálculo del calendario
                                                            '<-***** Fin LUCV20180601
                                                        End If
                                                        'END JUEZ ***************************************************
                                                        If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                                        If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                    
                                                    '2 MADM 20110425 **PEAC 20080110, Si es personalizada **********************
                                                    If R!nTpoValor = 3 Then
                                                        'BRGO 20111205
                                                        'MatCalend(i, 6) = ObtieneValorCuotaPoliza(psCtaCod, psTpoProdCod, i, nTotalCuotas, nValorUltCuota, nValorCuota, R!nPrdConceptoCod, pdHoy, sNumPoliza, nSumaAsegurada, bValidar, bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol, pdFechaSis)
                                                        'Se agregó pbLogicoBF, pdFechaBF, MatCalend(i, 0)
                                                        If R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipoteca Then 'LUCV20180601, Agregó según ERS022-2018
                                                            MatCalend(i, 6) = MatCalend(i, 15)
                                                        ElseIf R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipotecaGracia Then 'LUCV20180601, Agregó según ERS022-2018
                                                            MatCalend(i, 6) = MatCalend(i, 16)
                                                        Else
                                                            MatCalend(i, 6) = ObtieneValorCuotaPoliza(psCtaCod, psTpoProdCod, i, nTotalCuotas, nValorUltCuota, nValorCuota, R!nPrdConceptoCod, pdHoy, sNumPoliza, nSumaAsegurada, bValidar, bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol, pdFechaSis, pbLogicoBF, pdFechaBF, MatCalend(i, 0), MatCalend(i, 1), pnCuotaMivienda, pnPlazo, pnTipoPeriodo, pnExoSegMYPE) 'pnPlazo,pnTipoPeriodo APRI 20170227
                                                            'WIOR 20120521 SE AGREGO EL PARAMETRO pdFechaSis
                                                            'APRI20180821 ERS061-2018 AGREGÓ pnExoSegMYPE
                                                        End If
                                                    End If
                                                    '***********************************************************
                                                End If
                                            End If
                                            oCredRel.siguiente
                                        Loop
                                        MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                     End If
                                
                                End If
                            Else    'Si no se considera las Relaciones entonces
                                'Filtramos por Edad
                                bCumpleRestricc = False
                                oCredRel.nPuntMat = oCredRel.PosicionTitular
                                bCumpleRestricc = CumpleRestriccion(2, R, oCredRel, pdHoy) 'WIOR 20140823 -RESTRICCION 2
                                'WIOR 20140823 COMENTO
                                'If Not IsNull(R!nEdadOper) Then
                                '    Select Case R!nEdadOper
                                '        Case 2 'Mayor
                                '            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then bCumpleRestricc = True
                                '        Case 3 'Mayor Igual
                                '            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then bCumpleRestricc = True
                                '        Case 4 'Menor
                                '            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then bCumpleRestricc = True
                                '        Case 5 'Menor Igual
                                '            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then bCumpleRestricc = True
                                '        Case 6 'Igual a
                                '            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then bCumpleRestricc = True
                                '    End Select
                                'Else
                                '    bCumpleRestricc = True
                                'End If
                                'WIOR 20140823 FIN COMENTARIO
                            
                                'Si es Fijo
                                If bCumpleRestricc Then
                                    If R!nTpoValor = 1 Then
                                        'By Capi 28122007  Acta 051-2007
                                        If R!nPrdConceptoCod = gColocConceptoCodGastoEstCta Then
                                            bCumpleEspecifGastoEstCta = obtenerEspecifGastoEstCta(MatCalend(i, 0), dFecPivot)
                                            MatCalend(i, 6) = 0
                                            If bCumpleEspecifGastoEstCta = True And nTotalCuotas > 1 Then
                                                dFecPivot = MatCalend(i, 0)
                                                MatCalend(i, 6) = Format(nValor, "#0.00")
                                            End If
                                        Else
                                            If psTpoCredCod <> "851" Then
                                                If R!nPrdConceptoCod = 1224 Then MatCalend(i, 6) = IIf(R!nPrdConceptoCod = 1224, Format(getValor1224(pnPlazo, nValor), "#0.00"), Format(nValor, "#0.00"))
                                                'End If 'WIOR 20140823 COMENTO
                                                'JUEZ 20130529 ********************
                                                If R!nPrdConceptoCod = gColocConceptoCodGastoComisionEnvioEstCta Then
                                                    'If bEnvioEstCta Then
                                                    MatCalend(i, 6) = Format(IIf(bEnvioEstCta, nCostoEnvioEstCta, 0), "#0.00") 'JUEZ 20130827
                                                    'End If
                                                End If 'END JUEZ *************************
                                            Else
                                                'If R!nPrdConceptoCod = 1224 Then
                                                '    nGastos1224 = IIf(R!nPrdConceptoCod = 1224, Format(getValor1224(pnPlazo, nValor), "#0.00"), Format(nValor, "#0.00"))
                                                '    MatCalend(i, 6) = Format(nGastos1224, "#0.00")
                                                'Else
                                                If R!nPrdConceptoCod = 1247 Then MatCalend(i, 6) = IIf(R!nPrdConceptoCod = 1247, Format(nValor, "#0.00"), 0)                                                     '+ CDbl(nGastos1224)
                                                'End If 'WIOR 20140823 COMENTO
                                                'JUEZ 20130529 ********************
                                                If R!nPrdConceptoCod = gColocConceptoCodGastoComisionEnvioEstCta Then
                                                    'If bEnvioEstCta Then
                                                    MatCalend(i, 6) = Format(IIf(bEnvioEstCta, nCostoEnvioEstCta, 0), "#0.00") 'JUEZ 20130827
                                                    'End If
                                                End If 'END JUEZ *************************
                                            End If
                                        
                                        End If
                                    End If
                                    'Si es Porcentaje
                                    If R!nTpoValor = 2 Then
                                        If R!cAplicaMonto = "C" Then nRes = Format(Format(nValor / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                        If R!cAplicaMonto = "D" Then nRes = Format(Format(nValor / 100, "#0.000000") * nPrestamo, "#0.000000")
                                        If R!cAplicaMonto = "H" Then nRes = Format(Format(nValor / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                        If R!cAplicaMonto = "S" Then nRes = Format(Format(nValor / 100, "#0.000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                        If R!cAplicaMonto = "T" Then nRes = Format(Format(nValor / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                        'DAOR 20061213, Para el caso del seguro de desgravamen
                                        'If R!cAplicaMonto = "I" Then
                                        If R!cAplicaMonto = "I" Or R!cAplicaMonto = "K" Then 'JUEZ 20140310
                                        'MADM 20100908 ---------------------------------------------------------
                                        lnLimite = UBound(nCuotaAplica1217) - 1
                                        For lnCuenta = 0 To lnLimite
                                            If (i >= CInt(nCuotaAplica1217(lnCuenta)) And (CInt(nCuotaAplica1217(lnCuenta)) <> -1)) Then lnAcumula = lnAcumula + 1
                                            'End If
                                        Next
                                        
                                        If bExisteMasDeUnTitular Then
                                            If lnAcumula = 1 Then
                                                nValor = R!nValor
                                            ElseIf lnAcumula > 1 Then
                                                nAplicaRestriTit1217xDos = True
                                            End If
                                        Else
                                            If lnAcumula = 1 Then nAplicaRestriTit1217xUno = True
                                            'End If
                                        End If
                                        
                                        lnAcumula = 0
                                        lnLimite = 0
                                        'END MADM ******************************************************************
                                            'JUEZ 20140310 **********************************************
                                            'If i = 0 Then
                                            '    nRes = Format(Format(nValor / 100, "#0.000000") * (nPrestamo + CDbl(MatCalend(i, 4))), "#0.000000")
                                            'Else
                                            '    nRes = Format(Format(nValor / 100, "#0.000000") * (CDbl(MatCalend(i - 1, 7)) + CDbl(MatCalend(i, 4))), "#0.000000")
                                            'End If
                                            'nRes = CalculaSeguroDesgravamen(i, R!cAplicaMonto, nValor, nPrestamo, CDbl(MatCalend(i, 4)), CDbl(MatCalend(IIf(i = 0, 1, i) - 1, 7))) 'WIOR 20140822 COMENTO
                                            '->*****LUCV20180601, Comentó y agregó(El parámetro de retorno pMatCalendSegDes, se obtendra en el calendario)
                                            'nRes = CalculaSeguroDesgravamen(i, R!cAplicaMonto, nValor, nPrestamo, CDbl(MatCalend(i, 4)), CDbl(MatCalend(IIf(i = 0, 1, i) - 1, 7)), CDate(MatCalend(i, 0)), CDbl(R!nValor), CDbl(R!nValorDosTit), MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, nCoberMinSegDes, nCoberMaxSegDes, pMatCalendSegDes, psCtaCod) 'WIOR 20140822 'JUEZ 20150714 psCtaCod
                                            nRes = MatCalend(i, 8) 'El seguro desgravamen ya se incluye en el cálculo del calendario
                                            '<-***** Fin LUCV20180601
                                            'END JUEZ ***************************************************
                                        End If
                                        If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                        If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                       'madm 20100908 ********************************************
                                        
                                        '->***** LUCV20180601, Comentó y Agregó según ERS022-2018
                                        'If (nAplicaRestriTit1217xDos) Or (nAplicaRestriTit1217xUno And bExisteMasDeUnTitular = False) Then
                                        '    MatCalend(i, 6) = Format(0, "#0.00")
                                        'Else
                                        '    MatCalend(i, 6) = Format(nRes, "#0.00")
                                        'End If
                                            MatCalend(i, 6) = Format(nRes, "#0.00") 'Seg. Desgravamen
                                        '<-***** Fin LUCV20180601
                                        '***********************************************************
                                    End If
                                    '1 MADM 20110425 - AplicaGastoPoliza **PEAC 20080110, Si es personalizada **********************
                                    If R!nTpoValor = 3 Then
                                       'ALPA20140312**********************************************
                                       'MatCalend(i, 6) = ObtieneValorCuotaPoliza(psCtaCod, psTpoProdCod, i, nTotalCuotas, nValorUltCuota, nValorCuota, R!nPrdConceptoCod, pdHoy, sNumPoliza, nSumaAsegurada, bValidar, bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol, pdFechaSis)
                                       'Se agregó pbLogicoBF, pdFechaBF, MatCalend(i, 0)
                                       'RECO20160407**************************************************
                                        If R!nPrdConceptoCod = 1272 Then
                                            MatCalend(i, 7) = ObtieneValorCuotaPoliza(psCtaCod, psTpoProdCod, i, nTotalCuotas, nValorUltCuota, nValorCuota, R!nPrdConceptoCod, pdHoy, sNumPoliza, nSumaAsegurada, bValidar, bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol, pdFechaSis, pbLogicoBF, pdFechaBF, MatCalend(i, 0), MatCalend(i, 1), pnCuotaMivienda, pnPlazo, pnTipoPeriodo, pnExoSegMYPE) 'pnPlazo,pnTipoPeriodo APRI 20170227
                                            'APRI20180821 ERS061-2018 AGREGÓ pnExoSegMYPE
                                        Else
                                            If R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipoteca Then 'LUCV20180601, Agregó según ERS022-2018
                                                MatCalend(i, 6) = MatCalend(i, 15)
                                            ElseIf R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipotecaGracia Then 'LUCV20180601, Agregó según ERS022-2018
                                                MatCalend(i, 6) = MatCalend(i, 16)
                                            Else
                                                MatCalend(i, 6) = ObtieneValorCuotaPoliza(psCtaCod, psTpoProdCod, i, nTotalCuotas, nValorUltCuota, nValorCuota, R!nPrdConceptoCod, pdHoy, sNumPoliza, nSumaAsegurada, bValidar, bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol, pdFechaSis, pbLogicoBF, pdFechaBF, MatCalend(i, 0), MatCalend(i, 1), pnCuotaMivienda, pnPlazo, pnTipoPeriodo) 'pnPlazo,pnTipoPeriodo APRI 20170227
                                            End If
                                        End If
                                       'RECO FIN *******************************************************
                                       'WIOR 20120521 SE AGREGO EL PARAMETRO pdFechaSis
                                    End If
                                    '***********************************************************
                                End If
                            End If
                            If R!bAplNumConCer Then 'Multiplicar por el numero de consultas a la central de riesgos
                                MatCalend(i, 6) = MatCalend(i, 6) * nNumConCer
                            End If
                            'If R!bAplNumConMic Then  'Multiplicar por el numero de consultas Score Microfinanzas gitu 03-06-2009
                            If IIf(IsNull(R!bAplNumConMic), 0, R!bAplNumConMic) Then  'Multiplicar por el numero de consultas Score Microfinanzas gitu 03-06-2009
                                MatCalend(i, 6) = MatCalend(i, 6) * nNumConMicr
                            End If
                            '******************COMENTADO POR APRI 20170417
'                            If i = 0 Then
'                                If R!bAplNumMeses Or R!nPrdConceptoCod = "1231" Then  'Multiplicar por el numero de meses para la primera cuota
'                                    nMunMesesPorDia = Round(DateDiff("d", dFecDesemb, MatCalend(i, 0)) / 30, 0)
'                                    nNumMesesPorMes = DateDiff("m", dFecDesemb, MatCalend(i, 0))
'                                    MatCalend(i, 6) = MatCalend(i, 6) * IIf(nMunMesesPorDia >= nNumMesesPorMes, nMunMesesPorDia, nNumMesesPorMes)
'
'                                    'MatCalend(i, 6) = MatCalend(i, 6) * IIf(DateDiff("m", dFecDesemb, MatCalend(i, 0)) = 0, 1, DateDiff("m", dFecDesemb, MatCalend(i, 0)))
'
'                                End If
'                            'MAVM 20130410 *** Seg Desg Por Periodo
'                            ElseIf pnPlazo >= 30 Then
'                                nMunMesesPorDia = Round(pnPlazo / 30, 0)
'                                MatCalend(i, 6) = MatCalend(i, 6) * nMunMesesPorDia
'                            '***
'                            End If
                            'APRI20180924 ERS061-2018
                            Dim nMontoGasto As Currency
                            nMontoGasto = 0
                            If R!nPrdConceptoCod = 1272 Then
                                nMontoGasto = CDbl(MatCalend(i, 7))
                            Else
                                nMontoGasto = CDbl(MatCalend(i, 6))
                            End If
                            'END APRI
                            'MatCalend(i, 6) = OperacionPrimeraCuota(i, R!bAplNumMeses, R!nPrdConceptoCod, dFecDesemb, CDate(MatCalend(i, 0)), pnPlazo, CDbl(MatCalend(i, 6))) 'APRI 20170417
                            
                            If R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipoteca Then 'LUCV20180601, Agregó según ERS022-2018
                                MatCalend(i, 6) = MatCalend(i, 15)
                            ElseIf R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipotecaGracia Then 'LUCV20180601, Agregó según ERS022-2018
                                MatCalend(i, 6) = MatCalend(i, 16)
                            Else
                                MatCalend(i, 6) = OperacionPrimeraCuota(i, R!bAplNumMeses, R!nPrdConceptoCod, dFecDesemb, CDate(MatCalend(i, 0)), pnPlazo, nMontoGasto) 'APRI20180924 ERS061-2018
                            End If
                            '*****************END APRI*****************
                            
                            If bCumpleRestricc And (CDbl(MatCalend(i, 6)) > 0 Or (CDbl(MatCalend(i, 6)) = 0 And i = 0) Or CDbl(MatCalend(i, 7)) > 0) Then
                            'If bCumpleRestricc And (CDbl(MatCalend(i, 6)) > 0 Or (CDbl(MatCalend(i, 6)) = 0 And i = 0)) Then
                                nNumGastos = nNumGastos + 1
                                MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                MatGastos(nNumGastos - 1, 1) = i + 1
                                MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod
                                                                                                                                
                                If R!nMontoMensual = "1" Then
                                    If pnTipoPeriodo <> 2 And pnPlazo <> 30 Then 'Si no es fechafija
                                        MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6) * (pnPlazo / 30)
                                    Else
                                        MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
                                    End If
                                Else
                                    'COMENTADO POR APRI20180925 ERS061-2018
                                    'RECO20160407 ***************************************
                                    'If R!nPrdConceptoCod = 1272 Then
                                    '    MatGastos(nNumGastos - 1, 3) = MatCalend(i, 7)
                                    'Else
                                    '    MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
                                    'End If
                                    'RECO FIN ********************************************
                                    MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6) 'APRI20180925 ERS061-2018
                                End If
                                If R!nPrdConceptoCod = gColocConceptoCodGastoSeguro7 Then
                                    If pnExoSeguroDes = 1 Then MatGastos(nNumGastos - 1, 3) = Format(0, "#0.00")
                                    'End If
                                End If
                                If R!nAplicado = gColocConceptoAplCuota Then Exit For
                                'End If
                            End If
                        End If
                    End If
            Next i
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
    End If
        
    'Para los Gastos que se cargan en los Cancelaciones y Pagos
    If psAplicaProceso = "PA" Or psAplicaProceso = "CA" Or psAplicaProceso = "PP" Then
        
        Set oCred = New COMDCredito.DCOMCredito
        Set RGastosExon = oCred.RecuperaGastosExonerados(psCtaCod)
        Set oCred = Nothing
        ReDim MatGastosExon(0)
        Do While Not RGastosExon.EOF
            ReDim Preserve MatGastosExon(UBound(MatGastosExon) + 1)
            MatGastosExon(UBound(MatGastosExon) - 1) = Str(Trim(RGastosExon!nPrdConceptoCod))
            RGastosExon.MoveNext
        Loop
        
        nNumGastos = 0
        ReDim MatGastos(1000, 5)
        Set oGastos = New COMDCredito.DCOMGasto
        Set R = oGastos.RecuperaGastosAplicablesCuotas(CInt(Mid(psCtaCod, 9, 1)), psAplicaProceso, sGastoFijoVariable, , MatGastosExon)
        
        'ALPA 20100607 B2***************
        'If Mid(psCtaCod, 6, 3) <> "301" Then
        If psTpoCredCod <> "751" Then
            Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P", True)
            Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G", True)
        Else
            'verifico la institucion
            Set odGastos1 = New COMDCredito.DCOMGasto
            cInstitucion = odGastos1.ObtenerCodInstitucionByCredito(psCtaCod)
            Set odGastos1 = Nothing
            Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P", True, cInstitucion)
            Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G", True, cInstitucion)
        End If
        Set oGastos = Nothing
        
        Do While Not R.EOF
            bAplicaGasto = False
            RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod
            If RFiltro.RecordCount = 0 Then bAplicaGasto = True
            'End If
            '************************************************************
            'Si tiene Filtro Por Tipos de Garantias
            '************************************************************
            If bAplicaGasto Then
                bAplicaGasto = False
                RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'G'"
                If RFiltro.RecordCount = 0 Then bAplicaGasto = True
                'End If
                
                If Not bAplicaGasto Then
                    RGar.MoveFirst
                    Do While Not RGar.EOF
                        If Not bAplicaGasto Then
                            RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & RGar!nTpoGar & " AND cAgeCod = '" & psAgeCodAct & "' AND cTpoFiltro = 'G'"
                            If RFiltro.RecordCount > 0 Then
                                bAplicaGasto = True
                                Exit Do
                            Else
                                bAplicaGasto = False
                            End If
                        End If
                        RGar.MoveNext
                    Loop
                End If
            End If
            If bAplicaGasto Then
                Set oCred = New COMDCredito.DCOMCredito
                Set RCredVig = oCred.RecuperaDatosCreditoVigente(psCtaCod)
                Set oCred = Nothing
                'Si se encuentra en el Rango Permitido
                If nPrestamo >= R!nInicial And nPrestamo <= R!nFinal Then
                    bCumpleRestricc = False
                    'Si esta sujeto al vencimiento de dias de atraso
                    If Not IsNull(R!nDiasApl) Then
                       If pnDiasAtraso = R!nDiasApl Then
                            bCumpleRestricc = True
                       Else
                            bCumpleRestricc = False
                       End If
                    Else
                        bCumpleRestricc = True
                    End If
                    
                    If bCumpleRestricc Then
                        'Si esta sujeto a restriccion de Operacion
                        bCumpleRestricc = False
                        If Not IsNull(R!nOperador) Then
                            Select Case R!cOperMonto
                                Case "P" 'Prestamo
                                    Select Case R!nEdadOper
                                        Case 2 'Mayor
                                            If nPrestamo > CDbl(Format((R!nOperPorc / 100) * RCredVig!nMontoCol)) Then bCumpleRestricc = True
                                        Case 3 'Mayor Igual
                                            If nPrestamo >= CDbl(Format((R!nOperPorc / 100) * RCredVig!nMontoCol)) Then bCumpleRestricc = True
                                        Case 4 'Menor
                                            If nPrestamo < CDbl(Format((R!nOperPorc / 100) * RCredVig!nMontoCol)) Then bCumpleRestricc = True
                                        Case 5 'Menor Igual
                                            If nPrestamo <= CDbl(Format((R!nOperPorc / 100) * RCredVig!nMontoCol)) Then bCumpleRestricc = True
                                        Case 6 'Igual a
                                            If nPrestamo = CDbl(Format((R!nOperPorc / 100) * RCredVig!nMontoCol)) Then bCumpleRestricc = True
                                    End Select
                                Case "C" 'Cuota
                                    Select Case R!nEdadOper
                                        Case 2 'Mayor
                                            If nPrestamo > CDbl(Format((R!nOperPorc / 100) * nCuotaApr)) Then bCumpleRestricc = True
                                            'End If
                                        Case 3 'Mayor Igual
                                            If nPrestamo >= CDbl(Format((R!nOperPorc / 100) * nCuotaApr)) Then bCumpleRestricc = True
                                            'End If
                                        Case 4 'Menor
                                            If nPrestamo < CDbl(Format((R!nOperPorc / 100) * nCuotaApr)) Then bCumpleRestricc = True
                                            'End If
                                        Case 5 'Menor Igual
                                            If nPrestamo <= CDbl(Format((R!nOperPorc / 100) * nCuotaApr)) Then bCumpleRestricc = True
                                            'End If
                                        Case 6 'Igual a
                                            If nPrestamo = CDbl(Format((R!nOperPorc / 100) * nCuotaApr)) Then bCumpleRestricc = True
                                            'End If
                                    End Select
                                Case "G" 'Garantia
                                    Select Case R!nEdadOper
                                        Case 2 'Mayor
                                            If nPrestamo > CDbl(Format((R!nOperPorc / 100) * nMontoGarantia)) Then bCumpleRestricc = True
                                            'End If
                                        Case 3 'Mayor Igual
                                            If nPrestamo >= CDbl(Format((R!nOperPorc / 100) * nMontoGarantia)) Then bCumpleRestricc = True
                                            'End If
                                        Case 4 'Menor
                                            If nPrestamo < CDbl(Format((R!nOperPorc / 100) * nMontoGarantia)) Then bCumpleRestricc = True
                                            'End If
                                        Case 5 'Menor Igual
                                            If nPrestamo <= CDbl(Format((R!nOperPorc / 100) * nMontoGarantia)) Then bCumpleRestricc = True
                                            'End If
                                        Case 6 'Igual a
                                            If nPrestamo = CDbl(Format((R!nOperPorc / 100) * nMontoGarantia)) Then bCumpleRestricc = True
                                            'End If
                                    End Select
                                Case "S" 'Saldo
                                    Select Case R!nEdadOper
                                        Case 2 'Mayor
                                            If nPrestamo > CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then bCumpleRestricc = True
                                            'End If
                                        Case 3 'Mayor Igual
                                            If nPrestamo >= CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then bCumpleRestricc = True
                                            'End If
                                        Case 4 'Menor
                                            If nPrestamo < CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then bCumpleRestricc = True
                                            'End If
                                        Case 5 'Menor Igual
                                            If nPrestamo <= CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then bCumpleRestricc = True
                                            'End If
                                        Case 6 'Igual a
                                            If nPrestamo = CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then bCumpleRestricc = True
                                            'End If
                                    End Select
                                Case "H" 'Hipoteca
                                    Set oCred = New COMDCredito.DCOMCredito
                                    MontoPrePagado = oCred.RecuperaDatosPrepago(psCtaCod, pdHoy)
                                    Set oCred = Nothing
                                    Select Case R!nOperador
                                        Case 2 'Mayor
                                            If pnMontoCapital + MontoPrePagado > CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then bCumpleRestricc = True
                                            'End If
                                        Case 3 'Mayor Igual
                                            If pnMontoCapital + MontoPrePagado >= CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then bCumpleRestricc = True
                                            'End If
                                        Case 4 'Menor
                                            If pnMontoCapital + MontoPrePagado < CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then bCumpleRestricc = True
                                            'End If
                                        Case 5 'Menor Igual
                                            If pnMontoCapital + MontoPrePagado <= CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then bCumpleRestricc = True
                                            'End If
                                        Case 6 'Igual a
                                            If pnMontoCapital + MontoPrePagado = CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo)) Then bCumpleRestricc = True
                                            'End If
                                    End Select
                            End Select
                        Else
                            bCumpleRestricc = True
                        End If
                    End If
                    'Si cumple la primera Restriccion
                    'Buecar Por Restriccion por edad del Titular
                    If bCumpleRestricc Then
                        bCumpleRestricc = False
                        oCredRel.nPuntMat = oCredRel.PosicionTitular
                        If Not IsNull(R!nEdadOper) Then
                            Select Case R!nEdadOper
                                Case 2 'Mayor
                                    If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then bCumpleRestricc = True
                                    'End If
                                Case 3 'Mayor Igual
                                    If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then bCumpleRestricc = True
                                    'End If
                                Case 4 'Menor
                                    If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then bCumpleRestricc = True
                                    'End If
                                Case 5 'Menor Igual
                                    If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then bCumpleRestricc = True
                                    'End If
                                Case 6 'Igual a
                                    If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then bCumpleRestricc = True
                                    'End If
                            End Select
                        Else
                            bCumpleRestricc = True
                        End If
                        
                        If bCumpleRestricc Then
                            If Not IsNull(R!nOperDiasVenc) Then
                                bCumpleRestricc = False
                                nDiasAtraso = pnDiasAtraso
                                Select Case R!nOperDiasVenc
                                    Case 2 'Mayor
                                        If nDiasAtraso > R!nDiasVenc Then bCumpleRestricc = True
                                        'End If
                                    Case 3 'Mayor Igual
                                        If nDiasAtraso >= R!nDiasVenc Then bCumpleRestricc = True
                                        'End If
                                    Case 4 'Menor
                                        If nDiasAtraso < R!nDiasVenc Then bCumpleRestricc = True
                                        'End If
                                    Case 5 'Menor Igual
                                        If nDiasAtraso <= R!nDiasVenc Then bCumpleRestricc = True
                                        'End If
                                    Case 6 'Igual a
                                        If nDiasAtraso = R!nDiasVenc Then bCumpleRestricc = True
                                        'End If
                                End Select
                                
                            End If
                        End If
                        
                        nMontoTemp = 0
                        'Si cumple la restriccion
                        If bCumpleRestricc Then
                            If R!nTpoValor = 1 Then
                                nMontoTemp = Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                If nMontoTemp > 0 Then
                                    nNumGastos = nNumGastos + 1
                                    MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                    MatGastos(nNumGastos - 1, 1) = pnCuotaPend
                                    'MatGastos(nNumGastos - 1, 2) = "GASTO POR CANCELACION, PAGO O PREPAGO " & Space(150) & R!nPrdConceptoCod
                                    MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod
                                    'Validar Exoneracion del Seguro - AVMM - 16-12-2006
                                    If R!nPrdConceptoCod = gColocConceptoCodGastoSeguro7 Then
                                        If pnExoSeguroDes = 1 Then
                                            MatGastos(nNumGastos - 1, 3) = Format(0, "#0.00")
                                        Else
                                            MatGastos(nNumGastos - 1, 3) = Format(nMontoTemp, "#0.00")
                                        End If
                                    Else
                                        MatGastos(nNumGastos - 1, 3) = Format(nMontoTemp, "#0.00")
                                    End If
                                    
                                    MatGastos(nNumGastos - 1, 4) = Trim(R!cAplicaProceso)
                                End If
                            End If
                            'Si es Porcentaje
                            If R!nTpoValor = 2 Then
                                If Not IsNull(R!cOperMonto) Then
                                    If R!cOperMonto = "H" Then
                                        nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * ((pnMontoCapital + MontoPrePagado) - CDbl(Format((R!nOperPorc / 100) * RCredVig!nSaldo))), "#0.000000")
                                    Else
                                        If R!cAplicaMonto = "C" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nPrestamo, "#0.000000")
                                        'End If
                                        If R!cAplicaMonto = "D" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(0, 0)), "#0.000000")
                                        'End If
                                        If R!cAplicaMonto = "H" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                        'End If
                                        If R!cAplicaMonto = "S" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * RCredVig!nSaldo, "#0.000000")
                                        'End If
                                        If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                        'End If
                                        If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                        'End If
                                    End If
                                Else
                                    If R!cAplicaMonto = "C" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nPrestamo, "#0.000000")
                                    'End If
                                    If R!cAplicaMonto = "D" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(0, 0)), "#0.000000")
                                    'End If
                                    If R!cAplicaMonto = "H" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                    'End If
                                    If R!cAplicaMonto = "S" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * RCredVig!nSaldo, "#0.000000")
                                    'End If
                                    If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                    'End If
                                    If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                    'End If
                                End If
                                nMontoTemp = Format(nRes, "#0.00")
                                nNumGastos = nNumGastos + 1
                                MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                MatGastos(nNumGastos - 1, 1) = pnCuotaPend
                                MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod
                                'Validacion de Exoneracion de Seguro --- AVMM -- 16-12-2006
                                If R!nPrdConceptoCod = gColocConceptoCodGastoSeguro7 Then
                                    If pnExoSeguroDes = 1 Then
                                        MatGastos(nNumGastos - 1, 3) = Format(0, "#0.00")
                                    Else
                                        MatGastos(nNumGastos - 1, 3) = Format(nMontoTemp, "#0.00")
                                    End If
                                Else
                                    MatGastos(nNumGastos - 1, 3) = Format(nMontoTemp, "#0.00")
                                End If
                                MatGastos(nNumGastos - 1, 4) = Trim(R!cAplicaProceso)
                            End If
                        End If
                    End If
                End If
            End If
            
            R.MoveNext
        Loop
    End If
    'Para Cierre de Dia
    If psAplicaProceso = "CD" Then
        nNumGastos = 0
        ReDim MatGastos(1000, 4)
        If Not pHabRecorsed Then
            Set oGastos = New COMDCredito.DCOMGasto
            Set R = oGastos.RecuperaGastosAplicablesCuotas(CInt(Mid(psCtaCod, 9, 1)), psAplicaProceso, , True)
            If psTpoCredCod <> "751" Then
                 Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P", True)
                 Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G", True)
            Else
                'verifico la institucion
                Set odGastos1 = New COMDCredito.DCOMGasto
                cInstitucion = odGastos1.ObtenerCodInstitucionByCredito(psCtaCod)
              Set odGastos1 = Nothing
                 Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P", True, cInstitucion)
                    Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G", True, cInstitucion)
            End If
            Set oGastos = Nothing
        Else
            Set R = pR
            Set RFiltro = pRFiltro
            Set RFiltroGar = pRFiltroGar
            If R.RecordCount > 0 Then R.MoveFirst
            'End If
            If RFiltro.RecordCount > 0 Then RFiltro.MoveFirst
            'End If
            If RFiltroGar.RecordCount > 0 Then RFiltroGar.MoveFirst
            'End If
        End If
        
        Do While Not R.EOF
            For i = 0 To UBound(MatCalend) - 1
                bAplicaGasto = False
                RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod
                If RFiltro.RecordCount = 0 Then bAplicaGasto = True
                'End If
                If Not bAplicaGasto Then
                    RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & psTpoCredCod & " AND cAgeCod = '" & psAgeCodAct & "' AND cTpoFiltro = 'P'" & IIf(cInstitucion <> "", " AND cIntitucion='" & cInstitucion & "'", "")
                    If RFiltro.RecordCount > 0 Then
                        bAplicaGasto = True
                    Else
                        bAplicaGasto = False
                    End If
                End If
                'Si tiene Filtro Por Tipos de Garantias
                If bAplicaGasto Then
                    bAplicaGasto = False
                    RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'G'"
                    If RFiltro.RecordCount = 0 Then bAplicaGasto = True
                    'End If 'WIOR 20140823 COMENTO

                    If Not bAplicaGasto Then
                        RGar.MoveFirst
                        Do While Not RGar.EOF
                            If Not bAplicaGasto Then
                                RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & RGar!nTpoGar & " AND cAgeCod = '" & psAgeCodAct & "' AND cTpoFiltro = 'G'"
                                If RFiltro.RecordCount > 0 Then
                                    bAplicaGasto = True
                                    Exit Do
                                Else
                                    bAplicaGasto = False
                                End If
                            End If
                            RGar.MoveNext
                        Loop
                    End If
                End If

                If bAplicaGasto Then
                    If nPrestamo >= R!nInicial And nPrestamo <= R!nFinal Then
                        If Not IsNull(R!nDiasApl) Then
                            nDiasAtraso = (pdHoy - CDate(MatCalend(i, 0))) + 1
                            If R!nDiasApl = nDiasAtraso Then
                                If R!nTpoValor = 1 Then
                                    nMontoTemp = Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                    MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                End If
                                'Si es Porcentaje
                                If R!nTpoValor = 2 Then
                                    If R!cAplicaMonto = "C" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 2)), "#0.000000")
                                    'End If 'WIOR 20140822 COMENTO
                                    If R!cAplicaMonto = "D" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nPrestamo, "#0.000000")
                                    'End If 'WIOR 20140822 COMENTO
                                    If R!cAplicaMonto = "H" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                    'End If 'WIOR 20140822 COMENTO
                                    If R!cAplicaMonto = "S" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 3)), "#0.000000")
                                    'End If 'WIOR 20140822 COMENTO
                                    If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                    'End If 'WIOR 20140822 COMENTO
                                    If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                    'End If 'WIOR 20140822 COMENTO
                                    MatCalend(i, 6) = Format(nRes, "#0.00")
                                End If
                                If CDbl(MatCalend(i, 6)) > 0 Then
                                        nNumGastos = nNumGastos + 1
                                        MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                        MatGastos(nNumGastos - 1, 1) = i + 1
                                        MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod
                                        MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
                                End If
                            End If

                        End If

                        'Para el otro gasto de Cierre de Dia
                        If Not IsNull(R!nOperDiasVenc) Then
                            bCumpleRestricc = False
                            nDiasAtraso = pdHoy - CDate(MatCalend(i, 0))
                            Select Case R!nOperDiasVenc
                                Case 2 'Mayor
                                    If nDiasAtraso > R!nDiasVenc Then bCumpleRestricc = True
                                    'End If 'WIOR 20140822 COMENTO
                                Case 3 'Mayor Igual
                                    If nDiasAtraso >= R!nDiasVenc Then bCumpleRestricc = True
                                    'End If 'WIOR 20140822 COMENTO
                                Case 4 'Menor
                                    If nDiasAtraso < R!nDiasVenc Then bCumpleRestricc = True
                                    'End If 'WIOR 20140822 COMENTO
                                Case 5 'Menor Igual
                                    If nDiasAtraso <= R!nDiasVenc Then bCumpleRestricc = True
                                    'End If 'WIOR 20140822 COMENTO
                                Case 6 'Igual a
                                    If nDiasAtraso = R!nDiasVenc Then bCumpleRestricc = True
                                    'End If 'WIOR 20140822 COMENTO
                            End Select

                            If bCumpleRestricc Then
                                If R!nTpoValor = 1 Then
                                    nMontoTemp = Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                    nRes = nMontoTemp
                                    MatCalend(i, 6) = Format(nRes, "#0.00")
                                End If
                                'Si es Porcentaje
                                If R!nTpoValor = 2 Then
                                    If R!cAplicaMonto = "C" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 2)), "#0.000000")
                                    'End If 'WIOR 20140822 COMENTO
                                    If R!cAplicaMonto = "D" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nPrestamo, "#0.000000")
                                    'End If 'WIOR 20140822 COMENTO
                                    If R!cAplicaMonto = "H" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                    'End If 'WIOR 20140822 COMENTO
                                    If R!cAplicaMonto = "S" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 3)), "#0.000000")
                                    'End If 'WIOR 20140822 COMENTO
                                    If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                    'End If 'WIOR 20140822 COMENTO
                                    If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                    'End If 'WIOR 20140822 COMENTO
                                    MatCalend(i, 6) = Format(nRes, "#0.00")
                                End If
                                If CDbl(MatCalend(i, 6)) > 0 Then
                                        nNumGastos = nNumGastos + 1
                                        MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                        MatGastos(nNumGastos - 1, 1) = i + 1
                                        MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod
                                        MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
            R.MoveNext
        Loop
    End If
    If psAplicaProceso = "RE" Then
        Set oGastos = New COMDCredito.DCOMGasto
        Set R = oGastos.RecuperaGastosAplicablesCuotas(CInt(Mid(psCtaCod, 9, 1)))
        If psTpoCredCod <> "751" Then
            Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P", True)
            Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G", True)
        Else
            Set odGastos1 = New COMDCredito.DCOMGasto
            cInstitucion = odGastos1.ObtenerCodInstitucionByCredito(psCtaCod)
            Set odGastos1 = Nothing
            Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P", True, cInstitucion)
            Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G", True, cInstitucion)
        End If
        Set oGastos = Nothing
        Do While Not R.EOF
            For i = 0 To UBound(MatCalend) - 1
                bAplicaGasto = False
                    RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod
                    If RFiltro.RecordCount = 0 Then bAplicaGasto = True
                    If Not bAplicaGasto Then
                        RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And cProdCod = '" & psTpoCredCod & "' And nMoneda = " & Mid(psCtaCod, 9, 1) & " AND cAgeCod = '" & psAgeCodAct & "'" & IIf(cInstitucion <> "", " AND cIntitucion='" & cInstitucion & "'", "")
                        bAplicaGasto = IIf(RFiltro.RecordCount > 0, True, False)
                    End If

                    If bAplicaGasto Then
                        RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'G'"
                        If RFiltro.RecordCount = 0 Then bAplicaGasto = True
                        If Not bAplicaGasto Then
                            RGar.MoveFirst
                            Do While Not RGar.EOF
                                If Not bAplicaGasto Then
                                    RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & RGar!nTpoGar & " AND cAgeCod = '" & psAgeCodAct & "' AND cTpoFiltro = 'G'"
                                    If RFiltro.RecordCount > 0 Then
                                        bAplicaGasto = True
                                        Exit Do
                                    Else
                                        bAplicaGasto = False
                                    End If
                                End If
                                RGar.MoveNext
                            Loop
                        End If
                    End If

                    If bAplicaGasto Then
                        If CDbl(MatCalend(i, 2)) >= R!nInicial And CDbl(MatCalend(i, 2)) <= R!nFinal Then
                            'Si el gasto se aplica por relaciones de credito
                            If R!cFiltro = "I" Or R!cFiltro = "M" Then
                                If sTipoRelacGarantia = "I" Then 'Si el Credito es de Tipo Individual
                                    If R!cFiltro = "I" Then
                                        'Aplico a todos el mismo porcentaje
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then bCumpleRestricc = True
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then bCumpleRestricc = True
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then bCumpleRestricc = True
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then bCumpleRestricc = True
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then bCumpleRestricc = True
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If

                                                If bCumpleRestricc Then
                                                'Si es Fijo
                                                    If R!nTpoValor = 1 Then nMontoTemp = nMontoTemp + Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                        If R!cAplicaMonto = "C" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 2)), "#0.000000")
                                                        If R!cAplicaMonto = "D" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nPrestamo, "#0.000000")
                                                        If R!cAplicaMonto = "H" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                        If R!cAplicaMonto = "S" Then nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                                        If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                                        If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                End If
                                            oCredRel.siguiente
                                        Loop
                                        MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                    End If
                                Else 'Si es Mancomunada cobrar el porcentaje mancomunado al Titular y al Conyuge
                                     'y al resto el gasto de Individual
                                     'Para Titular y Conyuge
                                     If R!cFiltro = "M" Then
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                            If oCredRel.ObtenerValorRelac = gColRelPersConyugue Or oCredRel.ObtenerValorRelac = gColRelPersTitular Then
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then bCumpleRestricc = True
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then bCumpleRestricc = True
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then bCumpleRestricc = True
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then bCumpleRestricc = True
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then bCumpleRestricc = True
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If

                                                'Si es Fijo
                                                If bCumpleRestricc Then
                                                    If R!nTpoValor = 1 Then nMontoTemp = nMontoTemp + Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then

                                                    'MAVM 20110502 ***
                                                    Select Case R!cAplicaMonto
                                                        Case "C"
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                        Case "D"
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(i, 1)), "#0.000000")
                                                        Case "H"
                                                          nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                    End Select

                                                        If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                                        If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                End If
                                            End If
                                            oCredRel.siguiente
                                        Loop
                                        MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                     End If

                                     If R!cFiltro = "I" Then
                                        nMontoTemp = 0
                                        oCredRel.IniciarMatriz
                                        Do While Not oCredRel.EOF
                                            If oCredRel.ObtenerValorRelac <> gColRelPersConyugue And _
                                                oCredRel.ObtenerValorRelac <> gColRelPersTitular Then
                                                'Filtramos por Edad
                                                bCumpleRestricc = False
                                                If Not IsNull(R!nEdadOper) Then
                                                    Select Case R!nEdadOper
                                                        Case 2 'Mayor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then bCumpleRestricc = True
                                                        Case 3 'Mayor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then bCumpleRestricc = True
                                                        Case 4 'Menor
                                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then bCumpleRestricc = True
                                                        Case 5 'Menor Igual
                                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then bCumpleRestricc = True
                                                        Case 6 'Igual a
                                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then bCumpleRestricc = True
                                                    End Select
                                                Else
                                                    bCumpleRestricc = True
                                                End If

                                                'Si es Fijo
                                                If bCumpleRestricc Then
                                                    If R!nTpoValor = 1 Then nMontoTemp = nMontoTemp + Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                                    'Si es Porcentaje
                                                    If R!nTpoValor = 2 Then
                                                    'MAVM 20110502 ***
                                                    Select Case R!cAplicaMonto
                                                        Case "C"
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                        Case "D"
                                                            nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(i, 1)), "#0.000000")
                                                        Case "H"
                                                          nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                    End Select
                                                        If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                                        If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                                        nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                    End If
                                                End If
                                             End If
                                            oCredRel.siguiente
                                       Loop
                                        MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                    End If
                                End If
                            Else    'Si no se considera las Relaciones entonces
                                'Filtramos por Edad
                                bCumpleRestricc = False
                                oCredRel.nPuntMat = oCredRel.PosicionTitular
                                If Not IsNull(R!nEdadOper) Then
                                    Select Case R!nEdadOper
                                        Case 2 'Mayor
                                            If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then bCumpleRestricc = True
                                        Case 3 'Mayor Igual
                                            If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then bCumpleRestricc = True
                                        Case 4 'Menor
                                            If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then bCumpleRestricc = True
                                        Case 5 'Menor Igual
                                            If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then bCumpleRestricc = True
                                        Case 6 'Igual a
                                            If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then bCumpleRestricc = True
                                    End Select
                                Else
                                    bCumpleRestricc = True
                                End If

                                'Si es Fijo
                                If bCumpleRestricc Then
                                    If R!nTpoValor = 1 Then MatCalend(i, 6) = Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.000000")
                                    'Si es Porcentaje
                                    If R!nTpoValor = 2 Then
                                        'MAVM 20110502 ***
                                        Select Case R!cAplicaMonto
                                            Case "C"
                                                nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                            Case "D"
                                                nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * CDbl(MatDesemb(i, 1)), "#0.000000")
                                            Case "H"
                                              nRes = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                        End Select

                                        If nRes < R!nMontoMin Then nRes = R!nMontoMin
                                        If nRes > R!nMontoMax Then nRes = R!nMontoMax
                                        MatCalend(i, 6) = Format(nRes, "#0.00")
                                    End If
                                End If
                            End If
                            nNumGastos = nNumGastos + 1
                            MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                            MatGastos(nNumGastos - 1, 1) = i + 1
                            MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod
                            MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
                            If R!nPrdConceptoCod = gColocConceptoCodGastoSeguro7 Then If pnExoSeguroDes = 1 Then MatGastos(nNumGastos - 1, 3) = Format(0, "#0.00")
                                'End If
                            'End If
                            If R!nAplicado = gColocConceptoAplCuota Then Exit For
                            'End If
                        End If
                    End If
            Next i
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
    End If
                                   
    '**DAOR 20070411, Gastos para proceso de Reprogramación***********************************
    If psAplicaProceso = "RP" Then Call GeneraCalendarioGastosReprogramacion(MatCalend, MatDesemb, nNumGastos, psCtaCod, _
                                    nNumConCer, dFecDesemb, pnColocCred, nMontoGarantia, nMontoConstruccion, psGastoFijo, _
                                    nCuotaApr, nPrestamo, pnTipoPeriodo, pnPlazo, pnExoSeguroDes, nValorUltCuota, MatPersSegDes, _
                                    nEdadMinSegDes, nEdadMaxSegDes, nCoberMinSegDes, nCoberMaxSegDes, pMatCalendSegDes, pbPagoAnticipado)
    'WIOR 20140826 AGREGO MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, nCoberMinSegDes, nCoberMaxSegDes, pMatCalendSegDes
    'LUCV20180424, Agregó pbPagoAnticipado
    'End If
    If psAplicaProceso = "SI" Then
        dFecDesemb = pdFecSim 'MAVM 20130430
        Call GeneraCalendarioGastosParaSimulador(MatCalend, MatDesemb, nNumGastos, psCtaCod, nNumConCer, dFecDesemb, pnColocCred, nMontoGarantia, nMontoConstruccion, psGastoFijo, nCuotaApr, nPrestamo, pnTipoPeriodo, pnPlazo, pnExoSeguroDes, psTipoSegDes, pnValorInmueble, pnCuotaMivienda)
    End If
       'Call GeneraCalendarioGastosParaSimulador(MatCalend, MatDesemb, nNumGastos, psCtaCod, nNumConCer, dFecDesemb, pnColocCred, nMontoGarantia, nMontoConstruccion, psGastoFijo, nCuotaApr, nPrestamo, pnTipoPeriodo, pnPlazo, pnExoSeguroDes) 'MAVM 250120625 ***
    'End If
    Set oCredRel = Nothing
        
    If pbSoloMostrar Then
        GeneraCalendarioGastos = FormateaMatrizGastos(RedimensionaMatriz(MatGastos, nNumGastos), nNumGastos)
    Else
        GeneraCalendarioGastos = RedimensionaMatriz(MatGastos, nNumGastos)
    End If
End Function
'*** BRGO 20111205 ******************************************************************
''Se agregó para reducir procedimiento
'Private Function ObtieneValorCuotaPoliza(ByVal psCtaCod As String, ByVal psTpoProdCod As String, ByVal i As Integer, ByVal pnTotalCuotas As Integer, ByVal pnValorUltCuota As Double, _
'ByVal pnValorCuota As Double, ByVal pnPrdConceptoCod As Integer, ByVal pdHoy As Date, Optional ByRef psNumPoliza As String = 0, _
'    Optional ByRef pnSumaAsegurada As Double = 0, Optional ByVal bValidar As Boolean = False, Optional ByVal bEsParalelo As Boolean = False, Optional ByVal bEsRecurrentePol As Boolean = False, Optional ByVal bEsRefinanciadoPol As Boolean = False, _
'    Optional ByVal pdFechaSis As Date = "01/01/1900", Optional pbLogicoBF As Integer = 0, Optional pdFechaBF As Date = CDate("1900-01-01"), Optional ByVal pdFechaCuota As Date = CDate("1900-01-01"), Optional ByVal pnNroCuota As Integer = 0, Optional ByVal pCuotaMIVIENDA As Integer = -1) As Double
'    'WIOR 20120521 SE AGREGO EL PARAMETRO pdFechaSis
'    Dim nCuotaPoliza, nFactor As Double
'    Dim nCuotasAnual As Integer
'    Dim rsPol As ADODB.Recordset
'    Dim nCuotasGracia As Integer
'    'WIOR 20120521********************************
'    Dim oCredito As COMDCredito.DCOMCredito
'    Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
'    Dim nTC As Double
'    Set oCredito = New COMDCredito.DCOMCredito
'    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
'    nTC = oTipoCam.EmiteTipoCambio(pdFechaSis, TCFijoDia)
'    'WIOR -FIN ***********************************
'    If psTpoProdCod <> "517" Then
'        Select Case pnPrdConceptoCod
'            Case gColocConceptoCodGastoPolizaIncendioHipoteca
'            'MatCalend(i, 6) = IIf(AplicaGastoPoliza(bValidar, bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol), Format(0, "#0.00"), Format(IIf(i + 1 = nTotalCuotas, nValorUltCuota, nValorCuota), "#0.00"))
'            nCuotaPoliza = IIf(AplicaGastoPoliza(bValidar, bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol), Format(0, "#0.00"), Format(IIf(i + 1 = pnTotalCuotas, pnValorUltCuota, pnValorCuota), "#0.00"))
'            'ALPA20140210******************************************************
'            If pbLogicoBF = 1 Then
'                If DateDiff("D", pdFechaBF, pdFechaCuota) <= 0 Then
'                    nCuotaPoliza = 0
'                End If
'            End If
'            If pnNroCuota < pCuotaMIVIENDA Then
'                nCuotaPoliza = 0
'            End If
'            '******************************************************************
'            'WIOR 20120521*****************************************************
'            Case gColocConceptoCodGastoMicroseguros
'                Set rsPol = oCredito.ObtenerMicroseguro(psCtaCod)
'                If rsPol.RecordCount > 0 Then
'                    If IsNull(rsPol!nTipo) Then
'                        nCuotaPoliza = Format(0, "#0.00")
'                    Else
'                        If Trim(rsPol!nTipo) = "1" Then
'                            nCuotaPoliza = Format(ValorMinimoPrima(102743), "#0.00")
'                        Else
'                            nCuotaPoliza = Format(ValorMinimoPrima(102744), "#0.00")
'                        End If
'
'                        If Mid(psCtaCod, 9, 1) = "2" Then
'                            nCuotaPoliza = Format(Val(nCuotaPoliza) / Val(nTC), "#,#0.00")
'                        End If
'                    End If
'                Else
'                    nCuotaPoliza = Format(0, "#0.00")
'                End If
'            Case gColocConceptoCodGastoSeguroMultiriesgo
'                Set rsPol = oCredito.ObtenerMultiriesgo(psCtaCod)
'                If rsPol.RecordCount > 0 Then
'                    nCuotaPoliza = Format((CDbl(rsPol!nValorTotal) * 0.004) / 12, "#0.00")
'                    If CDbl(nCuotaPoliza) < ValorMinimoPrima(102745) Then
'                        nCuotaPoliza = Format(ValorMinimoPrima(102745), "#0.00")
'                    End If
'
'                    If Mid(psCtaCod, 9, 1) = "2" Then
'                        nCuotaPoliza = Format(Val(nCuotaPoliza) / Val(nTC), "#,#0.00")
'                    End If
'                Else
'                    nCuotaPoliza = Format(0, "#0.00")
'                End If
'             'WIOR FIN *********************************************************
'             'RECO2015052015 ERS023-2015****************************************
'             Case gColocConceptoCodGastoSeguroMultiriesgoMYPE
'                Set rsPol = oCredito.ObtenerMultiriesgoMYPE(psCtaCod)
'                Dim nValorAseg As Double
'                Dim nTasaAnual As Double
'                Dim nPrimaNetaAnual As Double
'                Dim nDerechoEmision As Double
'                Dim nIGV As Double
'                If rsPol.RecordCount > 0 Then
'                    nValorAseg = rsPol!nValorAseg
'                    nTasaAnual = rsPol!nValor
'                    nPrimaNetaAnual = nValorAseg * nTasaAnual
'                    nDerechoEmision = (nValorAseg * nTasaAnual) * 0.03
'                    nPrimaNetaAnual = nPrimaNetaAnual + nDerechoEmision
'                    nIGV = nPrimaNetaAnual * 0.18
'                    nPrimaNetaAnual = nPrimaNetaAnual + nIGV
'                    nPrimaNetaAnual = nPrimaNetaAnual / 12
'                    nCuotaPoliza = Round(nPrimaNetaAnual, 2)
'                Else
'                    nCuotaPoliza = Format(0, "#0.00")
'                End If
'             'RECO FIN *********************************************************
'        End Select
'    Else
'        Select Case pnPrdConceptoCod
'            Case gColocConceptoCodGastoPolizaVehiculoMultiriesgo
'                Dim oPol As COMDCredito.DCOMPoliza
'                Set oPol = New COMDCredito.DCOMPoliza
'
'                nCuotasAnual = 12
'
'                If (i + 1) Mod nCuotasAnual = 1 And (i + 1) > nCuotasAnual Then
'                    'WIOR 20130716 *******************
'                    If (i + 1) > 24 Then
'                        pnSumaAsegurada = pnSumaAsegurada * 0.9
'                    Else
'                    'WIOR FIN *******************
'                        pnSumaAsegurada = pnSumaAsegurada * 0.85
'                    End If
'
'                    Set rsPol = oPol.CargaDatosPrimaNetaPoliza(psNumPoliza, Round(pnSumaAsegurada, 0), "06/14/2011")
'                    nMontoPrima = rsPol!Total
'
'                End If
'
'                If i = 0 Then
'                    Set rsPol = oPol.CargaDatosPrimaNetaPoliza(psNumPoliza, Round(pnSumaAsegurada, 0), "06/14/2011")
'                    nMontoPrima = rsPol!Total
'                End If
'
'                nFactor = 0.01 / (1 - (1 / (1 + 0.01)) ^ (nCuotasAnual))
'                pnValorCuota = nMontoPrima * nFactor
'                pnValorUltCuota = pnValorCuota
'
'                nCuotaPoliza = Format(IIf(i + 1 = pnTotalCuotas, pnValorUltCuota, pnValorCuota), "#0.00")
'        End Select
'    End If
'    ObtieneValorCuotaPoliza = nCuotaPoliza
'End Function
''*** END BRGO ***********************************************************************
'EJVG20150713 ***
Private Function ObtieneValorCuotaPoliza(ByVal psCtaCod As String, ByVal psTpoProdCod As String, ByVal i As Integer, ByVal pnTotalCuotas As Integer, ByVal pnValorUltCuota As Double, _
ByVal pnValorCuota As Double, ByVal pnPrdConceptoCod As Integer, ByVal pdHoy As Date, Optional ByRef psNumPoliza As String = 0, _
    Optional ByRef pnSumaAsegurada As Double = 0, Optional ByVal bValidar As Boolean = False, Optional ByVal bEsParalelo As Boolean = False, Optional ByVal bEsRecurrentePol As Boolean = False, Optional ByVal bEsRefinanciadoPol As Boolean = False, _
    Optional ByVal pdFechaSis As Date = "01/01/1900", Optional pbLogicoBF As Integer = 0, Optional pdFechaBF As Date = CDate("1900-01-01"), Optional ByVal pdFechaCuota As Date = CDate("1900-01-01"), Optional ByVal pnNroCuota As Integer = 0, Optional ByVal pCuotaMIVIENDA As Integer = -1, _
    Optional ByVal pnPlazo As Integer, Optional ByVal pnTipoPeriodo As Integer, Optional ByVal pnExoSegMYPE As Integer = 0) As Double
     'pnPlazo,pnTipoPeriodo AGREGADO POR APRI 20170227
    'WIOR 20120521 SE AGREGO EL PARAMETRO pdFechaSis
    'APRI20180821 ERS061-2018 AGREGÓ pnExoSegMYPE
    Dim nCuotaPoliza, nFactor As Double
    Dim nCuotasAnual As Integer
    Dim rsPol As ADODB.Recordset
    Dim nCuotasGracia As Integer
    Dim oPol As COMDCredito.DCOMPoliza
    'WIOR 20120521********************************
    Dim oCredito As COMDCredito.DCOMCredito
    Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
    Dim nTC As Double
    Set oCredito = New COMDCredito.DCOMCredito
    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
    nTC = oTipoCam.EmiteTipoCambio(pdFechaSis, TCFijoDia)
    'WIOR -FIN ***********************************
    'APRI 20170227
    Dim nTotal As Single
    Dim dFecha1 As Date
    Dim dFecha2 As Date
    dFecha1 = DateAdd("m", DateDiff("m", 0, pdFechaCuota), 1)
    dFecha2 = DateAdd("m", DateDiff("m", 0, DateAdd("d", -pnPlazo, pdFechaCuota)), 1)
    If pnTipoPeriodo <> 2 Then 'And pnPlazo <> 30
        nTotal = (dFecha1 - dFecha2) / 30
        nTotal = Round(nTotal, 0)
    Else
        nTotal = 1
    End If
    'END APRI
    'If psTpoProdCod <> "517" Then
        Select Case pnPrdConceptoCod
            'EJVG20150702 ***
            'Case gColocConceptoCodGastoPolizaIncendioHipoteca
            'Case gColocConceptoCodGastoPolizaIncendioHipoteca, gColocConceptoCodGastoPolizaVehiculoMultiriesgo 'LUCV20180601, Comentó según ERS022-2018
            Case gColocConceptoCodGastoPolizaVehiculoMultiriesgo 'LUCV20180601, Agregó según ERS022-2018
            Set oPol = New COMDCredito.DCOMPoliza
            pnValorCuota = oPol.DamePrimaPolizaxCuota(psCtaCod, i + 1, pnTotalCuotas, pnPrdConceptoCod)
            pnValorUltCuota = pnValorCuota
            'END EJVG *******
            'MatCalend(i, 6) = IIf(AplicaGastoPoliza(bValidar, bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol), Format(0, "#0.00"), Format(IIf(i + 1 = nTotalCuotas, nValorUltCuota, nValorCuota), "#0.00"))
            nCuotaPoliza = IIf(AplicaGastoPoliza(bValidar, bEsParalelo, bEsRecurrentePol, bEsRefinanciadoPol), Format(0, "#0.00"), Format(IIf(i + 1 = pnTotalCuotas, pnValorUltCuota, pnValorCuota), "#0.00"))
            'EJVG20150702 ***
            If pnPrdConceptoCod = gColocConceptoCodGastoPolizaVehiculoMultiriesgo And psTpoProdCod = "517" Then
                Set oPol = New COMDCredito.DCOMPoliza

                nCuotasAnual = 12

                If (i + 1) Mod nCuotasAnual = 1 And (i + 1) > nCuotasAnual Then
                    If (i + 1) > 24 Then
                        pnSumaAsegurada = pnSumaAsegurada * 0.9
                    Else
                        pnSumaAsegurada = pnSumaAsegurada * 0.85
                    End If
                    
                    Set rsPol = oPol.CargaDatosPrimaNetaPoliza(psNumPoliza, Round(pnSumaAsegurada, 0), "06/14/2011")
                    nMontoPrima = rsPol!Total
                    
                End If
                
                If i = 0 Then
                    Set rsPol = oPol.CargaDatosPrimaNetaPoliza(psNumPoliza, Round(pnSumaAsegurada, 0), "06/14/2011")
                    nMontoPrima = rsPol!Total
                End If
                
                nFactor = 0.01 / (1 - (1 / (1 + 0.01)) ^ (nCuotasAnual))
                pnValorCuota = nMontoPrima * nFactor
                pnValorUltCuota = pnValorCuota
            
                nCuotaPoliza = Format(IIf(i + 1 = pnTotalCuotas, pnValorUltCuota, pnValorCuota), "#0.00")
            End If
            'END EJVG
            'ALPA20140210******************************************************
            If pbLogicoBF = 1 Then
                If DateDiff("D", pdFechaBF, pdFechaCuota) <= 0 Then
                    nCuotaPoliza = 0
                End If
            End If
            If pnNroCuota < pCuotaMIVIENDA Then
                nCuotaPoliza = 0
            End If
            '******************************************************************
            'WIOR 20120521*****************************************************
            Case gColocConceptoCodGastoMicroseguros
                Set rsPol = oCredito.ObtenerMicroseguro(psCtaCod)
                If rsPol.RecordCount > 0 Then
                    If IsNull(rsPol!nTipo) Then
                        nCuotaPoliza = Format(0, "#0.00")
                    Else
                        If Trim(rsPol!nTipo) = "1" Then
                            nCuotaPoliza = Format(ValorMinimoPrima(102743), "#0.00")
                        Else
                            nCuotaPoliza = Format(ValorMinimoPrima(102744), "#0.00")
                        End If
                       
                        If Mid(psCtaCod, 9, 1) = "2" Then
                            nCuotaPoliza = Format(Val(nCuotaPoliza) / Val(nTC), "#,#0.00")
                        End If
                    End If
                Else
                    nCuotaPoliza = Format(0, "#0.00")
                End If
            Case gColocConceptoCodGastoSeguroMultiriesgo
                Set rsPol = oCredito.ObtenerMultiriesgo(psCtaCod)
                If rsPol.RecordCount > 0 Then
                    nCuotaPoliza = Format((CDbl(rsPol!nValorTotal) * 0.004) / 12, "#0.00")
                    If CDbl(nCuotaPoliza) < ValorMinimoPrima(102745) Then
                        nCuotaPoliza = Format(ValorMinimoPrima(102745), "#0.00")
                    End If
                    
                    If Mid(psCtaCod, 9, 1) = "2" Then
                        nCuotaPoliza = Format(Val(nCuotaPoliza) / Val(nTC), "#,#0.00")
                    End If
                Else
                    nCuotaPoliza = Format(0, "#0.00")
                End If
             'WIOR FIN *********************************************************
             'RECO2015052015 ERS023-2015****************************************
                Case gColocConceptoCodGastoSeguroMultiriesgoMYPE
                If pnExoSegMYPE = 0 Then 'APRI20180821 ERS061-2018
                    Set rsPol = oCredito.ObtenerMultiriesgoMYPE(psCtaCod)
                    Dim nValorAseg As Double
                    Dim nTasaAnual As Double
                    Dim nPrimaNetaAnual As Double
                    Dim nDerechoEmision As Double
                    Dim nIGV As Double
                    If rsPol.RecordCount > 0 Then
                      'APRI20180821 ERS061-2018 COMENTÓ
                        'nValorAseg = rsPol!nValorAseg
                        'nTasaAnual = rsPol!nValor
                        'nPrimaNetaAnual = nValorAseg * nTasaAnual
                        'nDerechoEmision = (nValorAseg * nTasaAnual) * 0.03
                        'nPrimaNetaAnual = nPrimaNetaAnual + nDerechoEmision
                        'nIGV = nPrimaNetaAnual * 0.18
                        'nPrimaNetaAnual = nPrimaNetaAnual + nIGV
                        'nPrimaNetaAnual = nPrimaNetaAnual / 12
                        
                        'nCuotaPoliza = Round(nPrimaNetaAnual, 2)
                        nCuotaPoliza = Round(rsPol!nPrimaTrama, 2)
                    Else
                        nCuotaPoliza = Format(0, "#0.00")
                    End If
                Else
                    nCuotaPoliza = Format(0, "#0.00")
                End If
             'RECO FIN *********************************************************
        End Select
    'Else
    '    Select Case pnPrdConceptoCod
    '        Case gColocConceptoCodGastoPolizaVehiculoMultiriesgo
    '            Dim oPol As COMDCredito.DCOMPoliza
    '            Set oPol = New COMDCredito.DCOMPoliza
    '
    '            nCuotasAnual = 12
    '
    '            If (i + 1) Mod nCuotasAnual = 1 And (i + 1) > nCuotasAnual Then
    '                'WIOR 20130716 *******************
    '                If (i + 1) > 24 Then
    '                    pnSumaAsegurada = pnSumaAsegurada * 0.9
    '                Else
    '                'WIOR FIN *******************
    '                    pnSumaAsegurada = pnSumaAsegurada * 0.85
    '                End If
    '
    '                Set rsPol = oPol.CargaDatosPrimaNetaPoliza(psNumPoliza, Round(pnSumaAsegurada, 0), "06/14/2011")
    '                nMontoPrima = rsPol!Total
    '
    '            End If
    '
    '            If i = 0 Then
    '                Set rsPol = oPol.CargaDatosPrimaNetaPoliza(psNumPoliza, Round(pnSumaAsegurada, 0), "06/14/2011")
    '                nMontoPrima = rsPol!Total
    '            End If
    '
    '            nFactor = 0.01 / (1 - (1 / (1 + 0.01)) ^ (nCuotasAnual))
    '            pnValorCuota = nMontoPrima * nFactor
    '            pnValorUltCuota = pnValorCuota
    '
    '            nCuotaPoliza = Format(IIf(i + 1 = pnTotalCuotas, pnValorUltCuota, pnValorCuota), "#0.00")
    '    End Select
    'End If
    ObtieneValorCuotaPoliza = nCuotaPoliza * nTotal 'APRI 20170227
        
    Set oPol = Nothing
    Set oCredito = Nothing
    Set oTipoCam = Nothing
End Function
'END EJVG *******
Public Function GeneraCalendarioGastos_NEW(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Integer, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnTasaGracia As Double, ByVal pnDiaFijo As Integer, _
                ByVal bProxMes As Boolean, ByVal pMatGracia As Variant, ByVal pnMiViv As Integer, _
                ByVal pnCuotaCom As Integer, ByRef MatMiViv_2 As Variant, _
                ByRef MatDesemb As Variant, ByRef nNumGastos As Integer, _
                ByVal pdHoy As Date, ByVal psCtaCod As String, ByVal pnColocCred As Integer, _
                ByVal psAplicaProceso As String, _
                Optional ByVal psGastoFijo As String = "F", Optional ByVal pnCuotaApr As Double = -1, _
                Optional ByVal pnPrestamo As Double = -1, _
                Optional ByVal pnMontoCapital As Double = 0, Optional ByVal pnCuotaPend As Integer = -1, _
                Optional ByVal pHabRecorsed As Boolean = False, _
                Optional ByVal pR As ADODB.Recordset = Nothing, Optional ByVal pRFiltro As ADODB.Recordset = Nothing, _
                Optional ByVal pRFiltroGar As ADODB.Recordset = Nothing, Optional pnDiasAtraso As Integer = -1, _
                Optional ByVal pbSoloMostrar As Boolean = False, _
                Optional ByVal pnSugerAprob As Integer = 0, Optional ByVal pbNoMostrarCalendario As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
                Optional ByVal pMatDesPar As Variant = "", Optional ByVal bQuincenal As Boolean, _
                Optional ByRef pbErrorValidacion As Boolean = False, Optional ByVal pnDiaFijo2 As Integer = 0, _
                Optional ByVal pbIncrementaCapital As Boolean = False, Optional ByVal pdInicioPago As Date = 0, Optional ByVal pnCuotaMivienda As Double = -1, _
                Optional ByVal gnITFPorcent As Double = 0, Optional ByVal gbITFAplica As Double = 0, Optional ByVal pnExoSeguroDesgravamen As Integer = 0, _
                Optional ByVal psAgeCodAct As String = "", Optional ByVal psTpoProdCod As String = "", Optional ByVal psTpoCredCod As String = "", _
                Optional ByVal psTipoSegDes As String = "", Optional ByVal pnValorInmueble As Double = 0, _
                Optional ByVal pnMontoGra As Double = 0, Optional ByRef MatCalendIC As Variant, Optional ByRef MatCalendIC2 As Variant, Optional ByVal pnExoSegMYPE As Integer = 0, Optional pbLogicoBF As Integer = 0, Optional pdFechaBF As Date = CDate("1900-01-01"), _
                Optional ByRef pMatCalendSegDes As Variant, Optional ByRef pArrDatos As Variant = Nothing) As Variant
                'ALPA 20140312 pbLogicoBF , pdFechaBF
                'MAVM 20120625 Se agrego el parametro psTipoSegDes, pnValorInmueble (MIVivienda)
                'MAVM 20120309
                'WIOR 20131111 AGREGO pnCuotaBalon
                'WIOR 20140514 AGREGO pbReprogConvenio
                'WIOR 20140528 CAMBIO (Optional ByVal pbReprogConvenio As Boolean = False) POR (Optional ByRef pMatCalendSegDes As Variant)
                'APRI20180821 ERS061-2018 REEMPLAZÓ pnCuotaBalon POR pnExoSegMYPE
                'LUCV20180601, Reemplazó: [Optional pnMontoMivivienda As Currency = 0#] = [Optional ByRef pArrDatos As Variant = Nothing]
Dim MatGastos As Variant
Dim MatCalend As Variant
Dim bReprogConvenio As Boolean 'WIOR 20140515
Dim C As NCOMCalendario
'WIOR 20140528 *******************************
Dim pbReprogConvenio As Boolean
    pbReprogConvenio = False
    
    If IsArray(pMatCalendSegDes) Then
        If pMatCalendSegDes(0, 0) = "X" Then
            pbReprogConvenio = True
        End If
    End If
'WIOR FIN ************************************

'->***** LUCV20180601, Agregó según ERS022-2018
Dim nCont As Integer
Dim pnMontoMivivienda As Currency
    If IsArray(pArrDatos) Then
        For nCont = 0 To UBound(pArrDatos)
            Select Case nCont
                Case 0: pnMontoMivivienda = pArrDatos(nCont)
                Case 1: MatCalend = pArrDatos(nCont)
            End Select
        Next
    End If
'<-***** Fin LUCV20180601

'WIOR 20140515 ****************************
    bReprogConvenio = False
    If pbReprogConvenio Then
        bReprogConvenio = IIf(pnTipoGracia = 6, True, False)
    End If
'WIOR FIN *********************************

    Set C = New NCOMCalendario
    If Not IsArray(MatCalend) Then 'LUCV20180601, Según ERS022-2018
         MatCalend = C.GenerarCalendarioPagos(pnMonto, pnTasaInt, pnNroCuotas, _
                                pnPeriodo, pdFecDesemb, pnTipoCuota, _
                                pnTipoPeriodo, pnTipoGracia, _
                                pnDiasGracia, pnTasaGracia, pnDiaFijo, _
                                bProxMes, pMatGracia, pnMiViv, _
                                pnCuotaCom, MatMiViv_2, _
                                pnSugerAprob, pbNoMostrarCalendario, pbDesemParcial, _
                                pMatDesPar, bQuincenal, psCtaCod, _
                                pbErrorValidacion, pnDiaFijo2, _
                                pbIncrementaCapital, _
                                pdInicioPago, , , , pnMontoGra, , _
                                bReprogConvenio, pnMontoMivivienda)
                                'WIOR 20131111 AGREGO
                                'WIOR 20140514 AGREGO pbReprogConvenio
                                'APRI20180821 ERS061-2018 QUITÓ pnCuotaBalon
    End If
    Set C = Nothing
    
    'MAVM 20130209 ***
    If pnTipoGracia = 6 Or pbReprogConvenio Then 'WIOR 20140515 AGREGO pbReprogConvenio
        MatCalendIC = MatCalend
    End If
    '***
    
    'MAVM 20130305 ***
    If pnTipoGracia = 6 Then
        Dim y As Integer
        Dim i As Integer
        Dim MatCalendTemp() As String
        ReDim MatCalendTemp(UBound(MatCalend) - 1, 13)
        For i = 0 To UBound(MatCalend) - 1
            For y = 0 To 13
                MatCalendTemp(i, y) = MatCalend(i + 1, y)
            Next y
        Next i
        Erase MatCalend
        ReDim MatCalend(UBound(MatCalendTemp), 13)
        
        For i = 0 To UBound(MatCalendTemp)
            For y = 0 To 13
                MatCalend(i, y) = MatCalendTemp(i, y)
            Next y
        Next i
        Erase MatCalendTemp
    End If
    '***
    
    'MatGastos = GeneraCalendarioGastos(MatCalend, MatDesemb, nNumGastos, pdHoy, _
    '            psCtaCod, pnColocCred, psAplicaProceso, psGastoFijo, _
    '            pnCuotaApr, pnPrestamo, pnMontoCapital, pnCuotaPend, pHabRecorsed, pR, pRFiltro, _
    '            pRFiltroGar, pnDiasAtraso, pbSoloMostrar, pnTipoPeriodo, pnPeriodo, pnExoSeguroDesgravamen, psAgeCodAct, psTpoProdCod, psTpoCredCod, pdHoy, psTipoSegDes)
    '            'WIOR 20120521 SE AGREGO EL PARAMENTRO pdHoy para enviar el dato a pdFechaSis
       
    MatGastos = GeneraCalendarioGastos(MatCalend, MatDesemb, nNumGastos, pdHoy, _
                psCtaCod, pnColocCred, psAplicaProceso, psGastoFijo, _
                pnCuotaApr, pnPrestamo, pnMontoCapital, pnCuotaPend, pHabRecorsed, pR, pRFiltro, _
                pRFiltroGar, pnDiasAtraso, pbSoloMostrar, pnTipoPeriodo, pnPeriodo, pnExoSeguroDesgravamen, psAgeCodAct, psTpoProdCod, psTpoCredCod, pdHoy, _
                psTipoSegDes, pnValorInmueble, pdFecDesemb, pbLogicoBF, pdFechaBF, pMatCalendSegDes, pnCuotaMivienda, , pnExoSegMYPE)
                'MAVM 20120625 Se agrego el parametro psTipoSegDes, pnValorInmueble (MiVivienda)
                'MAVM 20130430 pdFecDesemb
                'WIOR 20140528 AGREGO pMatCalendSegDes
                
    Set C = Nothing
    
    GeneraCalendarioGastos_NEW = MatGastos
    
End Function

'**DAOR 20061214, Función que genera el calendario de gastos de una Reprogramación
Private Sub GeneraCalendarioGastosReprogramacion(ByRef MatCalend As Variant, ByRef MatDesemb As Variant, _
    ByRef nNumGastos As Integer, ByVal psCtaCod As String, ByVal pnNumConCer As Integer, _
    pdFecDesemb As Date, ByVal pnColocCred As Integer, _
    pnMontoGarantia As Double, pnMontoConstruccion As Double, Optional ByVal psGastoFijo As String = "F", _
    Optional ByVal pnCuotaApr As Double = -1, Optional ByVal pnPrestamo As Double = -1, _
    Optional ByVal pnTipoPeriodo As Integer = 0, Optional ByVal pnPlazo As Integer = 0, _
    Optional ByVal pnExoSeguroDes As Integer = 0, Optional ByRef pnValorUltCuota As Double = -1, _
    Optional ByVal MatPersSegDes As Variant, Optional ByVal pnEdadMinSegDes As Long, Optional ByVal pnEdadMaxSegDes As Long, _
    Optional ByVal pnCoberMinSegDes As Double, Optional ByVal pnCoberMaxSegDes As Double, Optional ByRef pMatCalendSegDes As Variant, _
    Optional ByVal pbPagoAnticipado As Boolean = False)
    'Optional ByVal pnExoSeguroDes As Integer = 0) 'MAVM 05022010 Se agrego el parametro pnValorUltCuota para la poliza
    'WIOR 20140826 AGREGO MatPersSegDes, pnEdadMinSegDes, pnEdadMaxSegDes, pnCoberMinSegDes, pnCoberMaxSegDes y pMatCalendSegDes
    'LUCV20180424, Agregó pbPagoAnticipado
    
Dim oGastos As COMDCredito.DCOMGasto
Dim R As ADODB.Recordset
Dim i As Integer
Dim RFiltro As ADODB.Recordset
Dim bAplicaGasto As Boolean
Dim nRes As Double
Dim oCred As COMDCredito.DCOMCredito
Dim oNCred As COMNCredito.NCOMCredito
Dim nPrestamo As Double
Dim nCuotaApr As Double
Dim nMontoGarantia As Double
Dim nMontoConstruccion As Double
Dim oGar As COMDCredito.DCOMGarantia
Dim nDiasAtraso As Integer
Dim sGastoFijoVariable As String
Dim oCredRel As COMDCredito.UCOMCredRela
Dim RFiltroGar As ADODB.Recordset
Dim RGar As ADODB.Recordset
Dim sTipoRelacGarantia As String
Dim bCumpleRestricc As Boolean
Dim nMontoTemp As Double
Dim MatCalTemp As Variant
Dim odGastos1 As COMDCredito.DCOMGasto
Dim cInstitucion As String
Dim nValor As Double
Dim bExisteMasDeUnTitular As Boolean
Dim dFecDesemb As Date
Dim nMunMesesPorDia As Integer
Dim nNumMesesPorMes As Integer
Dim RTemp As ADODB.Recordset
Dim RFiltroInst As ADODB.Recordset
Dim nNumConCer As Integer
Dim pdHoy As Date
Dim nNumConMicr As Integer
'MADM 20101120 -------------------------------------------
Dim RCol As ADODB.Recordset
Dim pcTpoProdCod As String
Dim pcAgeCodAct As String
'JUEZ 20130529 *****************
Dim bEnvioEstCta As Boolean
Dim nCostoEnvioEstCta As Double
'END JUEZ **********************

    Set oCred = New COMDCredito.DCOMCredito
    Set RCol = oCred.RecuperaColocaciones(psCtaCod)
    If Not (RCol.BOF Or RCol.EOF) Then
        pcTpoProdCod = RCol!cTpoCredCod
        pcAgeCodAct = RCol!cAgeCodAct
    End If
'END MADM ------------------------------------------------
'WIOR 20120529 ********************************************
Dim MatGastosRP As Variant
Dim oCalendario As COMDCredito.DCOMCalendario
Dim rsCalendario As ADODB.Recordset
Dim nMicroseguro As Double
Dim nMultiriesgo As Double
Dim nEcoTaxi As Double
Dim nSegMYPE As Double 'APRI20180827 ERS061-2017
'ReDim MatGastosRP(1000, 4)
'ReDim MatGastosRP(1000, 5) 'APRI20180827 ERS061-2017 'LUCV20180601. Comentó según ERS022-2018
ReDim MatGastosRP(1000, 6) 'APRI20180827 ERS061-2017  'LUCV20180601. Agregó según ERS022-2018
nEcoTaxi = 0
nMicroseguro = 0
nMultiriesgo = 0
pnValorUltCuota = 0
nSegMYPE = 0 'APRI20180827 ERS061-2017
Set oCalendario = New COMDCredito.DCOMCalendario
Set rsCalendario = oCalendario.RecuperaCalendarioPagos(psCtaCod, , False, , True, , , pbPagoAnticipado) 'LUCV20180424, Agregó pbPagoAnticipado
If rsCalendario.RecordCount > 0 Then
    If Mid(psCtaCod, 6, 3) <> "517" Then
        'pnValorUltCuota = CDbl(IIf(IsNull(rsCalendario!nGastoPolizaIncendio), 0, rsCalendario!nGastoPolizaIncendio))
        'APRI20170825 MEJORA INC1708210001
        Do While Not rsCalendario.EOF
            If CInt(rsCalendario!ncoloccalendestado) = 0 Then 'APRI20180827 ERS061-2018 MEJORA
            pnValorUltCuota = CDbl(IIf(IsNull(rsCalendario!nGastoPolizaIncendio), 0, rsCalendario!nGastoPolizaIncendio))
            
            If IIf(IsNull(rsCalendario!nGastoMicroseguros), 0, rsCalendario!nGastoMicroseguros) > 0 Then
                nMicroseguro = IIf(IsNull(rsCalendario!nGastoMicroseguros), 0, rsCalendario!nGastoMicroseguros)
            End If
        
            If IIf(IsNull(rsCalendario!nGastoSeguroMultiriesgo), 0, rsCalendario!nGastoSeguroMultiriesgo) > 0 Then
                nMultiriesgo = IIf(IsNull(rsCalendario!nGastoSeguroMultiriesgo), 0, rsCalendario!nGastoSeguroMultiriesgo)
            End If
            'APRI20180827 ERS061-2018
            If IIf(IsNull(rsCalendario!nGastoMultMYPE), 0, rsCalendario!nGastoMultMYPE) > 0 Then
                nSegMYPE = IIf(IsNull(rsCalendario!nGastoMultMYPE), 0, rsCalendario!nGastoMultMYPE)
            End If
            'END APRI
            If pnValorUltCuota > 0 Then
                Exit Do
            End If
            End If
            rsCalendario.MoveNext
        Loop
        'END APRI
    Else
        nEcoTaxi = CDbl(IIf(IsNull(rsCalendario!nGastoPolizaVehiculo), 0, rsCalendario!nGastoPolizaVehiculo))
    End If
'COMENTADO POR APRI20171019 POR INCIDENTE
'    If IIf(IsNull(rsCalendario!nGastoMicroseguros), 0, rsCalendario!nGastoMicroseguros) > 0 Then
'        nMicroseguro = IIf(IsNull(rsCalendario!nGastoMicroseguros), 0, rsCalendario!nGastoMicroseguros)
'    End If
'
'    If IIf(IsNull(rsCalendario!nGastoSeguroMultiriesgo), 0, rsCalendario!nGastoSeguroMultiriesgo) > 0 Then
'        nMultiriesgo = IIf(IsNull(rsCalendario!nGastoSeguroMultiriesgo), 0, rsCalendario!nGastoSeguroMultiriesgo)
'    End If
End If
Set oCalendario = Nothing
Set rsCalendario = Nothing
'WIOR FIN *************************************************

ObtieneValoresEnvioEstadoCta bEnvioEstCta, nCostoEnvioEstCta, psCtaCod 'JUEZ 20130529
                                                   
    Set oCredRel = New COMDCredito.UCOMCredRela
    Call oCredRel.CargaRelacPersCred(psCtaCod)
    If oCredRel.ExisteMasDeUnTitular Then
        sTipoRelacGarantia = "M"
    Else
        sTipoRelacGarantia = "I"
    End If
    bExisteMasDeUnTitular = oCredRel.ExisteTitularYCodeudor(psCtaCod) 'WIOR 20130419 AGREGO psCtaCod

    nCuotaApr = pnCuotaApr
    sGastoFijoVariable = psGastoFijo
    nPrestamo = pnPrestamo
    dFecDesemb = pdFecDesemb
    nMontoGarantia = pnMontoGarantia
    nMontoConstruccion = pnMontoConstruccion
    nNumConCer = pnNumConCer
    pdHoy = CDate(MatCalend(UBound(MatCalend) - 1, 0))
    nNumGastos = 0
    ReDim MatGastos(10000, 4)
                   
    '**Aplicar Gastos a la Cuota********************************************
    Set oGastos = New COMDCredito.DCOMGasto
    Set R = oGastos.RecuperaGastosAplicablesCuotas(CInt(Mid(psCtaCod, 9, 1)), "DE", sGastoFijoVariable)
    If pcTpoProdCod <> "751" Then  'MADM 20101120
        Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P", False)
        Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G", True)
    Else
        'verifico la institucion
        Set odGastos1 = New COMDCredito.DCOMGasto
        cInstitucion = odGastos1.ObtenerCodInstitucionByCredito(psCtaCod)
        Set odGastos1 = Nothing
        Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P", False)
        Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G", True, cInstitucion)
        Set RFiltroInst = oGastos.RecuperaFiltroAplicadoCuenta("P", True, cInstitucion)
    End If
    Set oGastos = Nothing
    
    Set RTemp = RFiltro.Clone   'ARCV 01-02-2007
    
    Do While Not R.EOF
        If R!nAplicado <> 5 Then 'No incluye los conceptos que se aplican solo a la primera cuota
            Set RFiltro = RTemp.Clone 'ARCV 01-02-2007
            If R!bAplValorDosTit Then 'Verificar si la restricción de Creditos con mas de un titular está activo
                nValor = IIf(bExisteMasDeUnTitular = True, IIf(IsNull(R!nValorDosTit), 0, R!nValorDosTit), IIf(IsNull(R!nValor), 0, R!nValor))
            Else
                nValor = IIf(IsNull(R!nValor), 0, R!nValor)
            End If
            
            For i = 0 To UBound(MatCalend) - 1
                bAplicaGasto = False
                RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'P'"
                If RFiltro.RecordCount = 0 Then
                   bAplicaGasto = True
                Else
                    If pcTpoProdCod = "751" Then   'MADM 20101120
                        RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'P'"
                        If RFiltro.RecordCount = 0 Then
                            bAplicaGasto = True
                        Else
                            'MADM 20101120    'RFiltroInst.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & Mid(psCtaCod, 6, 3) & " AND cAgeCod = '" & Mid(psCtaCod, 4, 2) & "' AND cTpoFiltro = 'P' AND cIntitucion='" & cInstitucion & "'"
                            RFiltroInst.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & pcTpoProdCod & " AND cAgeCod = '" & pcAgeCodAct & "' AND cTpoFiltro = 'P' AND cIntitucion='" & cInstitucion & "'"
                            If RFiltroInst.RecordCount > 0 Then
                                bAplicaGasto = True
                            Else
                                bAplicaGasto = False
                            End If
                        End If
                     Else
                        RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & pcTpoProdCod & " AND cAgeCod = '" & pcAgeCodAct & "' AND cTpoFiltro = 'P'"
                            If RFiltro.RecordCount > 0 Then
                                bAplicaGasto = True
                            Else
                                bAplicaGasto = False
                            End If
                     End If
                End If
                    
                '************************************************************
                'Si tiene Filtro Por Tipos de Garantias
                '************************************************************
                If bAplicaGasto Then
                    bAplicaGasto = False
                    RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'G'"
                    If RFiltro.RecordCount = 0 Then
                        bAplicaGasto = True
                    End If
                    
                    If Not bAplicaGasto Then
                        RGar.MoveFirst
                        Do While Not RGar.EOF
                            If Not bAplicaGasto Then
                                RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & RGar!nTpoGar & " AND cAgeCod = '" & pcAgeCodAct & "' AND cTpoFiltro = 'G'"
                                If RFiltro.RecordCount > 0 Then
                                    bAplicaGasto = True
                                    Exit Do
                                Else
                                    bAplicaGasto = False
                                End If
                            End If
                            RGar.MoveNext
                        Loop
                    End If
                End If
                    
                If bAplicaGasto Then
                    If CDbl(MatDesemb(0, 1)) >= R!nInicial And CDbl(MatDesemb(0, 1)) <= R!nFinal Then '-> Filtro por Monto de Prestamo
                        'Si el gasto se aplica por relaciones de credito
                        If R!cFiltro = "I" Or R!cFiltro = "M" Then
                            If sTipoRelacGarantia = "I" Then 'Si el Credito es de Tipo Individual
                                If R!cFiltro = "I" Then
                                    'Aplico a todos el mismo porcentaje
                                    nMontoTemp = 0
                                    oCredRel.IniciarMatriz
                                    Do While Not oCredRel.EOF
                                            'Filtramos por Edad
                                            bCumpleRestricc = False
                                            If Not IsNull(R!nEdadOper) Then
                                                Select Case R!nEdadOper
                                                    Case 2 'Mayor
                                                        If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 3 'Mayor Igual
                                                        If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 4 'Menor
                                                        If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 5 'Menor Igual
                                                        If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 6 'Igual a
                                                        If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                End Select
                                            Else
                                                bCumpleRestricc = True
                                            End If
                                            
                                            If bCumpleRestricc Then
                                            'Si es Fijo
                                                If R!nTpoValor = 1 Then
                                                    nMontoTemp = nMontoTemp + Format(nValor, "#0.000000")
                                                End If
                                                'Si es Porcentaje
                                                If R!nTpoValor = 2 Then
                                                    If R!cAplicaMonto = "C" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * CDbl(MatCalend(i, 2)), "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "D" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nPrestamo, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "H" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "S" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "T" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                                    End If
                                                    'DAOR 20061213, Para el caso del seguro de desgravamen
                                                    If R!cAplicaMonto = "I" Or R!cAplicaMonto = "K" Then 'JUEZ 20140310 Se agregó R!cAplicaMonto = "K"
                                                        'WIOR 20140826 COMENTO
                                                        'If i = 0 Then
                                                        '    nRes = Format(Format(nValor / 100, "#0.000000") * (nPrestamo + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                                        'Else
                                                        '    nRes = Format(Format(nValor / 100, "#0.000000") * (CDbl(MatCalend(i - 1, 7)) + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                                        'End If
                                                        'nRes = CalculaSeguroDesgravamen(i, R!cAplicaMonto, nValor, nPrestamo, CDbl(MatCalend(i, 4)), CDbl(MatCalend(IIf(i = 0, 1, i) - 1, 7)), CDate(MatCalend(i, 0)), CDbl(R!nValor), CDbl(R!nValorDosTit), MatPersSegDes, pnEdadMinSegDes, pnEdadMaxSegDes, pnCoberMinSegDes, pnCoberMaxSegDes, pMatCalendSegDes, psCtaCod) 'WIOR 20140826 'JUEZ 20150714 psCtaCod'LUCV20180601- Comentó. Según ERS022-2018
                                                        nRes = MatCalend(i, 8) 'LUCV20180601: El seguro desgravamen ya se incluye en el cálculo del calendario
                                                    End If
                                                    '-------------------------------------------------------
                                                    If nRes < R!nMontoMin Then
                                                        nRes = R!nMontoMin
                                                    End If
                                                    If nRes > R!nMontoMax Then
                                                       nRes = R!nMontoMax
                                                    End If
                                                    nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                End If
                                                'AQUI
                                                If R!nTpoValor = 3 Then
                                                    Select Case R!nPrdConceptoCod
                                                        Case gColocConceptoCodGastoPolizaIncendioHipoteca
                                                        'MatCalend(i, 6) = Format(pnValorUltCuota, "#0.00")
                                                        'WIOR 20120529 *************************************
                                                            If pnValorUltCuota > 0 Then
                                                                MatGastosRP(i, 1) = Format(pnValorUltCuota, "#0.00")
                                                            End If
                                                        Case gColocConceptoCodGastoPolizaVehiculoMultiriesgo
                                                            If nEcoTaxi > 0 Then
                                                                MatGastosRP(i, 2) = Format(nEcoTaxi, "#0.00")
                                                            End If
                                                        Case gColocConceptoCodGastoMicroseguros
                                                            If nMicroseguro > 0 Then
                                                                MatGastosRP(i, 3) = Format(nMicroseguro, "#0.00")
                                                            End If
                                                        Case gColocConceptoCodGastoSeguroMultiriesgo
                                                            If nMultiriesgo > 0 Then
                                                                MatGastosRP(i, 4) = Format(nMultiriesgo, "#0.00")
                                                            End If
                                                        'WIOR FIN ******************************************
                                                        'APRI20180827 ERS061-2018
                                                        Case gColocConceptoCodGastoSeguroMultiriesgoMYPE
                                                            If nSegMYPE > 0 Then
                                                                MatGastosRP(i, 5) = Format(nSegMYPE, "#0.00")
                                                            End If
                                                        'END APRI
                                                         'LUCV20180601, Según ERS022-2018
                                                         Case gColocConceptoCodGastoPolizaIncendioHipotecaGracia
                                                            If MatCalend(i, 16) > 0 Then
                                                                MatGastosRP(i, 6) = Format(MatCalend(i, 16), "#0.00")
                                                            End If
                                                        'Fin LUCV20180601
                                                    End Select
                                                End If
                                                
                                            End If
                                        oCredRel.siguiente
                                    Loop
                                    MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                Else
                                    Exit For
                                End If
                            Else 'Si es Mancomunada cobrar el porcentaje mancomunado todos menos al Titular
                                 'y al resto el gasto de Mancomunado
                                 'Para desiguales al Titular
                                 If R!cFiltro = "M" Then
                                    nMontoTemp = 0
                                    oCredRel.IniciarMatriz
                                    Do While Not oCredRel.EOF
                                        If oCredRel.ObtenerValorRelac <> gColRelPersTitular Then
                                            'Filtramos por Edad
                                            bCumpleRestricc = False
                                            If Not IsNull(R!nEdadOper) Then
                                                Select Case R!nEdadOper
                                                    Case 2 'Mayor
                                                        If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 3 'Mayor Igual
                                                        If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 4 'Menor
                                                        If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 5 'Menor Igual
                                                        If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 6 'Igual a
                                                        If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                End Select
                                            Else
                                                bCumpleRestricc = True
                                            End If
                                        
                                            'Si es Fijo
                                            If bCumpleRestricc Then
                                                If R!nTpoValor = 1 Then
                                                    nMontoTemp = nMontoTemp + Format(nValor, "#0.000000")
                                                End If
                                                'Si es Porcentaje
                                                If R!nTpoValor = 2 Then
                                                    If R!cAplicaMonto = "C" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "D" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nPrestamo, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "H" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "T" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "S" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                                    End If
                                                    'DAOR 20061213, Para el caso del seguro de desgravamen
                                                    If R!cAplicaMonto = "I" Or R!cAplicaMonto = "K" Then 'JUEZ 20140310 Se agregó R!cAplicaMonto = "K"
                                                        'WIOR 20140826 COMENTO
                                                        'If i = 0 Then
                                                        '    nRes = Format(Format(nValor / 100, "#0.000000") * (nPrestamo + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                                        'Else
                                                        '    nRes = Format(Format(nValor / 100, "#0.000000") * (CDbl(MatCalend(i - 1, 7)) + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                                        'End If
                                                        'nRes = CalculaSeguroDesgravamen(i, R!cAplicaMonto, nValor, nPrestamo, CDbl(MatCalend(i, 4)), CDbl(MatCalend(IIf(i = 0, 1, i) - 1, 7)), CDate(MatCalend(i, 0)), CDbl(R!nValor), CDbl(R!nValorDosTit), MatPersSegDes, pnEdadMinSegDes, pnEdadMaxSegDes, pnCoberMinSegDes, pnCoberMaxSegDes, pMatCalendSegDes, psCtaCod) 'WIOR 20140826 'JUEZ 20150714 psCtaCod 'LUCV20180601, Comentó Según ERS022-2018
                                                        nRes = MatCalend(i, 8) 'LUCV20180601:El seguro desgravamen ya se incluye en el cálculo del calendario
                                                    End If
                                                    If nRes < R!nMontoMin Then
                                                        nRes = R!nMontoMin
                                                    End If
                                                    If nRes > R!nMontoMax Then
                                                       nRes = R!nMontoMax
                                                    End If
                                                    nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                End If
                                                
                                                'AQUI
                                                If R!nTpoValor = 3 Then
                                                    Select Case R!nPrdConceptoCod
                                                        Case gColocConceptoCodGastoPolizaIncendioHipoteca
                                                        'MatCalend(i, 6) = Format(pnValorUltCuota, "#0.00")
                                                        'WIOR 20120529 *************************************
                                                            If pnValorUltCuota > 0 Then
                                                                MatGastosRP(i, 1) = Format(pnValorUltCuota, "#0.00")
                                                            End If
                                                        Case gColocConceptoCodGastoPolizaVehiculoMultiriesgo
                                                            If nEcoTaxi > 0 Then
                                                                MatGastosRP(i, 2) = Format(nEcoTaxi, "#0.00")
                                                            End If
                                                        Case gColocConceptoCodGastoMicroseguros
                                                            If nMicroseguro > 0 Then
                                                                MatGastosRP(i, 3) = Format(nMicroseguro, "#0.00")
                                                            End If
                                                        Case gColocConceptoCodGastoSeguroMultiriesgo
                                                            If nMultiriesgo > 0 Then
                                                                MatGastosRP(i, 4) = Format(nMultiriesgo, "#0.00")
                                                            End If
                                                        'WIOR FIN ******************************************
                                                        'APRI20180827 ERS061-2018
                                                        Case gColocConceptoCodGastoSeguroMultiriesgoMYPE
                                                            If nSegMYPE > 0 Then
                                                                MatGastosRP(i, 5) = Format(nSegMYPE, "#0.00")
                                                            End If
                                                        'END APRI
                                                        'LUCV20180601, Según ERS022-2018
                                                         Case gColocConceptoCodGastoPolizaIncendioHipotecaGracia
                                                            If MatCalend(i, 16) > 0 Then
                                                                MatGastosRP(i, 6) = Format(MatCalend(i, 16), "#0.00")
                                                            End If
                                                        'Fin LUCV20180601
                                                    End Select
                                                End If
                                                
                                            End If
                                        End If
                                        oCredRel.siguiente
                                    Loop
                                                                            
                                    MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                    
                                 End If
                                 
                                 If R!cFiltro = "I" Then
                                    nMontoTemp = 0
                                    oCredRel.IniciarMatriz
                                    Do While Not oCredRel.EOF
                                        If oCredRel.ObtenerValorRelac = gColRelPersTitular Then
                                            'Filtramos por Edad
                                            bCumpleRestricc = False
                                            If Not IsNull(R!nEdadOper) Then
                                                Select Case R!nEdadOper
                                                    Case 2 'Mayor
                                                        If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 3 'Mayor Igual
                                                        If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 4 'Menor
                                                        If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 5 'Menor Igual
                                                        If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 6 'Igual a
                                                        If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                End Select
                                            Else
                                                bCumpleRestricc = True
                                            End If
                                        
                                            'Si es Fijo
                                            If bCumpleRestricc Then
                                                If R!nTpoValor = 1 Then
                                                    nMontoTemp = nMontoTemp + Format(nValor, "#0.000000")
                                                End If
                                                'Si es Porcentaje
                                                If R!nTpoValor = 2 Then
                                                    If R!cAplicaMonto = "C" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "D" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nPrestamo, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "H" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "S" Then
                                                        nRes = Format(Format(nValor / 100, "#0.0000000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "T" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                                    End If
                                                    'DAOR 20061213, Para el caso del seguro de desgravamen
                                                    If R!cAplicaMonto = "I" Or R!cAplicaMonto = "K" Then 'JUEZ 20140310 Se agregó R!cAplicaMonto = "K"
                                                        'WIOR 20140826 COMENTO
                                                        'If i = 0 Then
                                                        '    nRes = Format(Format(nValor / 100, "#0.000000") * (nPrestamo + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                                        'Else
                                                        '    nRes = Format(Format(nValor / 100, "#0.000000") * (CDbl(MatCalend(i - 1, 7)) + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                                        'End If
                                                        'nRes = CalculaSeguroDesgravamen(i, R!cAplicaMonto, nValor, nPrestamo, CDbl(MatCalend(i, 4)), CDbl(MatCalend(IIf(i = 0, 1, i) - 1, 7)), CDate(MatCalend(i, 0)), CDbl(R!nValor), CDbl(R!nValorDosTit), MatPersSegDes, pnEdadMinSegDes, pnEdadMaxSegDes, pnCoberMinSegDes, pnCoberMaxSegDes, pMatCalendSegDes, psCtaCod) 'WIOR 20140826 'JUEZ 20150714 psCtaCod 'LUCV20180601, Comentó seún ERS022-2018
                                                        nRes = MatCalend(i, 8) 'LUCV20180601:El seguro desgravamen ya se incluye en el cálculo del calendario
                                                    End If
                                                    If nRes < R!nMontoMin Then
                                                        nRes = R!nMontoMin
                                                    End If
                                                    If nRes > R!nMontoMax Then
                                                       nRes = R!nMontoMax
                                                    End If
                                                    nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                End If
                                                
                                                'AQUI
                                                If R!nTpoValor = 3 Then
                                                    Select Case R!nPrdConceptoCod
                                                        Case gColocConceptoCodGastoPolizaIncendioHipoteca
                                                        'MatCalend(i, 6) = Format(pnValorUltCuota, "#0.00")
                                                        'WIOR 20120529 *************************************
                                                            If pnValorUltCuota > 0 Then
                                                                MatGastosRP(i, 1) = Format(pnValorUltCuota, "#0.00")
                                                            End If
                                                        Case gColocConceptoCodGastoPolizaVehiculoMultiriesgo
                                                            If nEcoTaxi > 0 Then
                                                                MatGastosRP(i, 2) = Format(nEcoTaxi, "#0.00")
                                                            End If
                                                        Case gColocConceptoCodGastoMicroseguros
                                                            If nMicroseguro > 0 Then
                                                                MatGastosRP(i, 3) = Format(nMicroseguro, "#0.00")
                                                            End If
                                                        Case gColocConceptoCodGastoSeguroMultiriesgo
                                                            If nMultiriesgo > 0 Then
                                                                MatGastosRP(i, 4) = Format(nMultiriesgo, "#0.00")
                                                            End If
                                                        'WIOR FIN ******************************************
                                                        'APRI20180827 ERS061-2018
                                                        Case gColocConceptoCodGastoSeguroMultiriesgoMYPE
                                                            If nSegMYPE > 0 Then
                                                                MatGastosRP(i, 5) = Format(nSegMYPE, "#0.00")
                                                            End If
                                                        'END APRI
                                                        'LUCV20180601. Comentó según ERS022-2018
                                                        Case gColocConceptoCodGastoPolizaIncendioHipotecaGracia
                                                            If MatCalend(i, 16) > 0 Then
                                                                MatGastosRP(i, 6) = Format(MatCalend(i, 16), "#0.00")
                                                            End If
                                                        'Fin LUCV20180601
                                                    End Select
                                                End If
                                                
                                            End If
                                        End If
                                        oCredRel.siguiente
                                    Loop
                                    MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                 End If
                            End If
                        Else    'Si no se considera las Relaciones entonces
                            'Filtramos por Edad
                            bCumpleRestricc = False
                            oCredRel.nPuntMat = oCredRel.PosicionTitular
                            If Not IsNull(R!nEdadOper) Then
                                Select Case R!nEdadOper
                                    Case 2 'Mayor
                                        If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                            bCumpleRestricc = True
                                        End If
                                    Case 3 'Mayor Igual
                                        If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                            bCumpleRestricc = True
                                        End If
                                    Case 4 'Menor
                                        If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                            bCumpleRestricc = True
                                        End If
                                    Case 5 'Menor Igual
                                        If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                            bCumpleRestricc = True
                                        End If
                                    Case 6 'Igual a
                                        If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                            bCumpleRestricc = True
                                        End If
                                End Select
                            Else
                                bCumpleRestricc = True
                            End If
                        
                            'Si es Fijo
                            If bCumpleRestricc Then
                                If R!nTpoValor = 1 Then
                                    'MatCalend(i, 6) = Format(nValor, "#0.00")
                                    'madm 20091110
                                    MatCalend(i, 6) = IIf(R!nPrdConceptoCod = 1224, Format(getValor1224(pnPlazo, nValor), "#0.00"), Format(nValor, "#0.00"))
                                    'end madm
                                End If
                                'JUEZ 20130529 ********************
                                If R!nPrdConceptoCod = gColocConceptoCodGastoComisionEnvioEstCta Then
                                    'If bEnvioEstCta Then
                                    MatCalend(i, 6) = Format(IIf(bEnvioEstCta, nCostoEnvioEstCta, 0), "#0.00") 'JUEZ 20130827
                                    'End If
                                End If
                                'END JUEZ *************************
                                'Si es Porcentaje
                                If R!nTpoValor = 2 Then
                                    If R!cAplicaMonto = "C" Then
                                        nRes = Format(Format(nValor / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "D" Then
                                        nRes = Format(Format(nValor / 100, "#0.000000") * nPrestamo, "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "H" Then
                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "S" Then
                                        nRes = Format(Format(nValor / 100, "#0.000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "T" Then
                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                    End If
                                    'DAOR 20061213, Para el caso del seguro de desgravamen
                                    If R!cAplicaMonto = "I" Or R!cAplicaMonto = "K" Then 'JUEZ 20140310 Se agregó R!cAplicaMonto = "K"
                                        'WIOR 20140826 COMENTO
                                        'If i = 0 Then
                                        '    nRes = Format(Format(nValor / 100, "#0.000000") * (nPrestamo + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                        'Else
                                        '    nRes = Format(Format(nValor / 100, "#0.000000") * (CDbl(MatCalend(i - 1, 7)) + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                        'End If
                                        'nRes = CalculaSeguroDesgravamen(i, R!cAplicaMonto, nValor, nPrestamo, CDbl(MatCalend(i, 4)), CDbl(MatCalend(IIf(i = 0, 1, i) - 1, 7)), CDate(MatCalend(i, 0)), CDbl(R!nValor), CDbl(R!nValorDosTit), MatPersSegDes, pnEdadMinSegDes, pnEdadMaxSegDes, pnCoberMinSegDes, pnCoberMaxSegDes, pMatCalendSegDes, psCtaCod) 'WIOR 20140826 'JUEZ 20150714 psCtaCod 'LUCV20180601, Comentó según ERS022-2018
                                        nRes = MatCalend(i, 8) 'LUCV20180601: El seguro desgravamen ya se incluye en el cálculo del calendario
                                    End If
                                    If nRes < R!nMontoMin Then
                                        nRes = R!nMontoMin
                                    End If
                                    If nRes > R!nMontoMax Then
                                       nRes = R!nMontoMax
                                    End If
                                    MatCalend(i, 6) = Format(nRes, "#0.00")
                                End If
                                'MAVM 05022010 Para poliza
                                If R!nTpoValor = 3 Then
                                    Select Case R!nPrdConceptoCod
                                        Case gColocConceptoCodGastoPolizaIncendioHipoteca
                                        'MatCalend(i, 6) = Format(pnValorUltCuota, "#0.00")
                                        'WIOR 20120529 *************************************
                                            'If pnValorUltCuota > 0 Then
                                            If MatCalend(i, 15) > 0 Then
                                                MatGastosRP(i, 1) = Format(MatCalend(i, 15), "#0.00")
                                            End If
                                        Case gColocConceptoCodGastoPolizaVehiculoMultiriesgo
                                            If nEcoTaxi > 0 Then
                                                MatGastosRP(i, 2) = Format(nEcoTaxi, "#0.00")
                                            End If
                                        Case gColocConceptoCodGastoMicroseguros
                                            If nMicroseguro > 0 Then
                                                MatGastosRP(i, 3) = Format(nMicroseguro, "#0.00")
                                            End If
                                        Case gColocConceptoCodGastoSeguroMultiriesgo
                                            If nMultiriesgo > 0 Then
                                                MatGastosRP(i, 4) = Format(nMultiriesgo, "#0.00")
                                            End If
                                        'WIOR FIN ******************************************
                                         'APRI20180827 ERS061-2018
                                        Case gColocConceptoCodGastoSeguroMultiriesgoMYPE
                                            If nSegMYPE > 0 Then
                                                MatGastosRP(i, 5) = Format(nSegMYPE, "#0.00")
                                            End If
                                        'END APRI
                                        'LUCV20180601, Agregó según
                                        Case gColocConceptoCodGastoPolizaIncendioHipotecaGracia
                                         If MatCalend(i, 16) > 0 Then
                                                MatGastosRP(i, 6) = Format(MatCalend(i, 16), "#0.00")
                                         End If
                                        'Fin LUCV20190601
                                    End Select
                                End If
                                '**
                            End If
                        End If
                        If R!bAplNumConCer Then 'Multiplicar por el numero de consultas a la central de riesgos
                            MatCalend(i, 6) = MatCalend(i, 6) * nNumConCer
                        End If
                        'If R!bAplNumConMic Then  'Multiplicar por el numero de consultas Score Microfinanzas gitu 03-06-2009
                        If IIf(IsNull(R!bAplNumConMic), 0, R!bAplNumConMic) Then  'Multiplicar por el numero de consultas Score Microfinanzas gitu 03-06-2009
                            MatCalend(i, 6) = MatCalend(i, 6) * nNumConMicr
                        End If
                        If i = 0 Then
                            If R!bAplNumMeses Then 'Multiplicar por el numero de meses para la primera cuota
                                nMunMesesPorDia = Round(DateDiff("d", dFecDesemb, MatCalend(i, 0)) / 30, 0)
                                nNumMesesPorMes = DateDiff("m", dFecDesemb, MatCalend(i, 0))
                                MatCalend(i, 6) = MatCalend(i, 6) * IIf(nMunMesesPorDia >= nNumMesesPorMes, nMunMesesPorDia, nNumMesesPorMes)
                                'MatCalend(i, 6) = MatCalend(i, 6) * IIf(DateDiff("m", dFecDesemb, MatCalend(i, 0)) = 0, 1, DateDiff("m", dFecDesemb, MatCalend(i, 0)))
                            End If
                        End If
                        If bCumpleRestricc And (CDbl(MatCalend(i, 6)) > 0 Or (CDbl(MatCalend(i, 6)) = 0 And i = 0)) Then
                           'WIOR 20120529 *************************************
                           If R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipoteca Then
                                    'If pnValorUltCuota > 0 Then
                                    If MatCalend(i, 15) > 0 Then
                                        nNumGastos = nNumGastos + 1
                                        MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                        MatGastos(nNumGastos - 1, 1) = i + 1
                                        MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod

                                        'LUCV20180601. Comentó Según ERS022-2018
'                                        If R!nMontoMensual = "1" Then
'                                            If pnTipoPeriodo <> 2 And pnPlazo <> 30 Then 'Si no es fechafija
'                                                MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 1) * (pnPlazo / 30)
'                                            Else
'                                                MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 1)
'                                            End If
'                                        Else
'                                            MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 1)
'                                        End If
'                                        If R!nPrdConceptoCod = gColocConceptoCodGastoSeguro7 Then
'                                            If pnExoSeguroDes = 1 Then
'                                                MatGastos(nNumGastos - 1, 3) = Format(0, "#0.00")
'                                            End If
'                                        End If
                                        MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 1)
                                        'Fin LUCV20180601
                                        
                                        If R!nAplicado = gColocConceptoAplCuota Then
                                            Exit For
                                        End If
                                    End If
                           ElseIf R!nPrdConceptoCod = gColocConceptoCodGastoPolizaVehiculoMultiriesgo Then
                                    If nEcoTaxi > 0 Then
                                        nNumGastos = nNumGastos + 1
                                        MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                        MatGastos(nNumGastos - 1, 1) = i + 1
                                        MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod

                                        If R!nMontoMensual = "1" Then
                                            If pnTipoPeriodo <> 2 And pnPlazo <> 30 Then 'Si no es fechafija
                                                MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 2) * (pnPlazo / 30)
                                            Else
                                                MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 2)
                                            End If
                                        Else
                                            MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 2)
                                        End If
                                        If R!nPrdConceptoCod = gColocConceptoCodGastoSeguro7 Then
                                            If pnExoSeguroDes = 1 Then
                                                MatGastos(nNumGastos - 1, 3) = Format(0, "#0.00")
                                            End If
                                        End If
                                        If R!nAplicado = gColocConceptoAplCuota Then
                                            Exit For
                                        End If
                                    End If
                           ElseIf R!nPrdConceptoCod = gColocConceptoCodGastoMicroseguros Then
                                    If nMicroseguro > 0 Then
                                        nNumGastos = nNumGastos + 1
                                        MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                        MatGastos(nNumGastos - 1, 1) = i + 1
                                        MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod

                                        If R!nMontoMensual = "1" Then
                                            If pnTipoPeriodo <> 2 And pnPlazo <> 30 Then 'Si no es fechafija
                                                MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 3) * (pnPlazo / 30)
                                            Else
                                                MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 3)
                                            End If
                                        Else
                                            MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 3)
                                        End If
                                        If R!nPrdConceptoCod = gColocConceptoCodGastoSeguro7 Then
                                            If pnExoSeguroDes = 1 Then
                                                MatGastos(nNumGastos - 1, 3) = Format(0, "#0.00")
                                            End If
                                        End If
                                        If R!nAplicado = gColocConceptoAplCuota Then
                                            Exit For
                                        End If
                                    End If
                           ElseIf R!nPrdConceptoCod = gColocConceptoCodGastoSeguroMultiriesgo Then
                                    If nMultiriesgo > 0 Then
                                        nNumGastos = nNumGastos + 1
                                        MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                        MatGastos(nNumGastos - 1, 1) = i + 1
                                        MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod

                                        If R!nMontoMensual = "1" Then
                                            If pnTipoPeriodo <> 2 And pnPlazo <> 30 Then 'Si no es fechafija
                                                MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 4) * (pnPlazo / 30)
                                            Else
                                                MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 4)
                                            End If
                                        Else
                                            MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 4)
                                        End If
                                        If R!nPrdConceptoCod = gColocConceptoCodGastoSeguro7 Then
                                            If pnExoSeguroDes = 1 Then
                                                MatGastos(nNumGastos - 1, 3) = Format(0, "#0.00")
                                            End If
                                        End If
                                        If R!nAplicado = gColocConceptoAplCuota Then
                                            Exit For
                                        End If
                                    End If
                           'WIOR 20130710 **************************************************************
                           'APRI20180827 ERS061-2018****************************************************
                            ElseIf R!nPrdConceptoCod = gColocConceptoCodGastoSeguroMultiriesgoMYPE Then
                                    If nSegMYPE > 0 Then
                                        nNumGastos = nNumGastos + 1
                                        MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                        MatGastos(nNumGastos - 1, 1) = i + 1
                                        MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod

                                        If R!nMontoMensual = "1" Then
                                            If pnTipoPeriodo <> 2 And pnPlazo <> 30 Then 'Si no es fechafija
                                                MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 5) * (pnPlazo / 30)
                                            Else
                                                MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 5)
                                            End If
                                        Else
                                            MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 5)
                                        End If
                                        If R!nAplicado = gColocConceptoAplCuota Then
                                            Exit For
                                        End If
                                    End If
                           'END APRI20180827************************************************************
                           'LUCV20180601. Agregó según ERS022-2018
                            ElseIf R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipotecaGracia Then
                                    If MatCalend(i, 16) > 0 Then
                                        nNumGastos = nNumGastos + 1
                                        MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                        MatGastos(nNumGastos - 1, 1) = i + 1
                                        MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod
                                        MatGastos(nNumGastos - 1, 3) = MatGastosRP(i, 6)
                                        
                                        If R!nAplicado = gColocConceptoAplCuota Then
                                            Exit For
                                        End If
                                    End If
                           'Fin LUCV20180601
                           ElseIf R!nPrdConceptoCod = gColocConceptoCodGastoComisionEnvioEstCta Then
                                    If bEnvioEstCta And nCostoEnvioEstCta > 0 Then
                                        nNumGastos = nNumGastos + 1
                                        MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                        MatGastos(nNumGastos - 1, 1) = i + 1
                                        MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod
        
                                        If R!nMontoMensual = "1" Then
                                            If pnTipoPeriodo <> 2 And pnPlazo <> 30 Then 'Si no es fechafija
                                                MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6) * (pnPlazo / 30)
                                            Else
                                                MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
                                            End If
                                        Else
                                            MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
                                        End If
                                        If R!nPrdConceptoCod = gColocConceptoCodGastoSeguro7 Then
                                            If pnExoSeguroDes = 1 Then
                                                MatGastos(nNumGastos - 1, 3) = Format(0, "#0.00")
                                            End If
                                        End If
                                        If R!nAplicado = gColocConceptoAplCuota Then
                                            Exit For
                                        End If
                                    End If
                           'WIOR FIN ********************************************************************
                           Else
                                nNumGastos = nNumGastos + 1
                                MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                                MatGastos(nNumGastos - 1, 1) = i + 1
                                MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod

                                If R!nMontoMensual = "1" Then
                                    If pnTipoPeriodo <> 2 And pnPlazo <> 30 Then 'Si no es fechafija
                                        MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6) * (pnPlazo / 30)
                                    Else
                                        MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
                                    End If
                                Else
                                    MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
                                End If
                                If R!nPrdConceptoCod = gColocConceptoCodGastoSeguro7 Then
                                    If pnExoSeguroDes = 1 Then
                                        MatGastos(nNumGastos - 1, 3) = Format(0, "#0.00")
                                    End If
                                End If
                                If R!nAplicado = gColocConceptoAplCuota Then
                                    Exit For
                                End If
                           End If
                            'WIOR FIN ******************************************
'                            nNumGastos = nNumGastos + 1
'                            MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
'                            MatGastos(nNumGastos - 1, 1) = i + 1
'                            MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod
'
'                            If R!nMontoMensual = "1" Then
'                                If pnTipoPeriodo <> 2 And pnPlazo <> 30 Then 'Si no es fechafija
'                                    MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6) * (pnPlazo / 30)
'                                Else
'                                    MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
'                                End If
'                            Else
'                                MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
'                            End If
'                            If R!nPrdConceptoCod = gColocConceptoCodGastoSeguro7 Then
'                                If pnExoSeguroDes = 1 Then
'                                    MatGastos(nNumGastos - 1, 3) = Format(0, "#0.00")
'                                End If
'                            End If
'                            If R!nAplicado = gColocConceptoAplCuota Then
'                                Exit For
'                            End If
                        Else
                            Exit For
                        End If
                    End If
                End If
            Next i
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set oCredRel = Nothing
End Sub

'**PEAC 20080815, Función que genera el calendario de gastos con el desgravamen para el simulador
Private Sub GeneraCalendarioGastosParaSimulador(ByRef MatCalend As Variant, ByRef MatDesemb As Variant, _
    ByRef nNumGastos As Integer, ByVal psCtaCod As String, ByVal pnNumConCer As Integer, _
    pdFecDesemb As Date, ByVal pnColocCred As Integer, _
    pnMontoGarantia As Double, pnMontoConstruccion As Double, Optional ByVal psGastoFijo As String = "F", _
    Optional ByVal pnCuotaApr As Double = -1, Optional ByVal pnPrestamo As Double = -1, _
    Optional ByVal pnTipoPeriodo As Integer = 0, Optional ByVal pnPlazo As Integer = 0, _
    Optional ByVal pnExoSeguroDes As Integer = 0, Optional ByVal psTipoSegDes As String = "", _
    Optional ByVal pnValorInmueble As Double = 0, Optional ByVal pnCuotaMivienda As Integer = -1) 'MAVM 20120625 Se agrego el parametro psTipoSegDes, pnValorInmueble (MIVivienda)
        
Dim oGastos As COMDCredito.DCOMGasto
Dim R As ADODB.Recordset
Dim i As Integer
Dim RFiltro As ADODB.Recordset
Dim bAplicaGasto As Boolean
Dim nRes As Double
Dim oCred As COMDCredito.DCOMCredito
Dim oNCred As COMNCredito.NCOMCredito
Dim nPrestamo As Double
Dim nCuotaApr As Double
Dim nMontoGarantia As Double
Dim nMontoConstruccion As Double
Dim oGar As COMDCredito.DCOMGarantia
Dim nDiasAtraso As Integer
Dim sGastoFijoVariable As String
Dim oCredRel As COMDCredito.UCOMCredRela
Dim RFiltroGar As ADODB.Recordset
Dim RGar As ADODB.Recordset
Dim sTipoRelacGarantia As String
Dim bCumpleRestricc As Boolean
Dim nMontoTemp As Double
Dim MatCalTemp As Variant
Dim odGastos1 As COMDCredito.DCOMGasto
Dim cInstitucion As String
Dim nValor As Double
Dim bExisteMasDeUnTitular As Boolean
Dim dFecDesemb As Date
Dim nMunMesesPorDia As Integer
Dim nNumMesesPorMes As Integer
Dim RTemp As ADODB.Recordset
Dim RFiltroInst As ADODB.Recordset
Dim nNumConCer As Integer
Dim pdHoy As Date
Dim nNumConMicr As Integer
'MADM 20101120 -------
Dim RCol As ADODB.Recordset
Dim pcTpoProdCod As String
Dim pcAgeCodAct As String
Dim nGastos1224 As Double

    Set oCred = New COMDCredito.DCOMCredito
    Set RCol = oCred.RecuperaColocaciones(psCtaCod)
    If Not (RCol.BOF Or RCol.EOF) Then
        pcTpoProdCod = RCol!cTpoCredCod
        pcAgeCodAct = RCol!cAgeCodAct
    End If
'END MADM -------

    'MAVM 20120625 ***
    Set oCredRel = New COMDCredito.UCOMCredRela
    If psTipoSegDes <> "" Then
        Call oCredRel.CargaRelacPersCred(psCtaCod)
        If psTipoSegDes = "I" Then
            sTipoRelacGarantia = "I"
            bExisteMasDeUnTitular = False
        Else
            sTipoRelacGarantia = "M"
            bExisteMasDeUnTitular = True
        End If
    Else
        '***
        If oCredRel.ExisteMasDeUnTitular Then
            sTipoRelacGarantia = "M"
        Else
            sTipoRelacGarantia = "I"
        End If
        bExisteMasDeUnTitular = oCredRel.ExisteTitularYCodeudor(psCtaCod) 'WIOR 20130419 AGREGO psCtaCod
        'MAVM 20120625 ***
    End If
    '***

    nCuotaApr = pnCuotaApr
    sGastoFijoVariable = psGastoFijo
    nPrestamo = pnPrestamo
    dFecDesemb = pdFecDesemb
    nMontoGarantia = pnMontoGarantia
    nMontoConstruccion = pnMontoConstruccion
    nNumConCer = pnNumConCer
    pdHoy = CDate(MatCalend(UBound(MatCalend) - 1, 0))
    nNumGastos = 0
    ReDim MatGastos(10000, 4)
                   
    '**Aplicar Gastos a la Cuota********************************************
    Set oGastos = New COMDCredito.DCOMGasto
    Set R = oGastos.RecuperaGastosAplicablesCuotas("1", "DE", sGastoFijoVariable)
    If pcTpoProdCod <> "751" Then  'MADM 20101120
        Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P", False)
        Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G", True)
    Else
        'verifico la institucion
        Set odGastos1 = New COMDCredito.DCOMGasto
        cInstitucion = odGastos1.ObtenerCodInstitucionByCredito(psCtaCod)
        Set odGastos1 = Nothing
        Set RFiltro = oGastos.RecuperaFiltroAplicadoCuenta("P", False)
        Set RFiltroGar = oGastos.RecuperaFiltroAplicadoCuenta("G", True, cInstitucion)
        Set RFiltroInst = oGastos.RecuperaFiltroAplicadoCuenta("P", True, cInstitucion)
    End If
    Set oGastos = Nothing
    Set RTemp = RFiltro.Clone   'ARCV 01-02-2007
    Do While Not R.EOF
        If R!nAplicado <> 5 Then 'No incluye los conceptos que se aplican solo a la primera cuota
            Set RFiltro = RTemp.Clone 'ARCV 01-02-2007
            If R!bAplValorDosTit Then 'Verificar si la restricción de Creditos con mas de un titular está activo
                nValor = IIf(bExisteMasDeUnTitular = True, IIf(IsNull(R!nValorDosTit), 0, R!nValorDosTit), IIf(IsNull(R!nValor), 0, R!nValor))
            Else
                nValor = IIf(IsNull(R!nValor), 0, R!nValor)
            End If
            
            For i = 0 To UBound(MatCalend) - 1
                'WIOR 20130904 *****************
                If i > 0 Then MatCalend(i, 6) = 0
                'WIOR FIN **********************
                bAplicaGasto = False
                RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'P'"
                If RFiltro.RecordCount = 0 Then
                    'MAVM 20120625 ***
                    'bAplicaGasto = True
                        If (R!nPrdConceptoCod = "1243" Or R!nPrdConceptoCod = "1245" Or R!nPrdConceptoCod = "1246") And psTipoSegDes <> "" Then
                            bAplicaGasto = False
                        Else
                            bAplicaGasto = True
                        End If
                    '***
                Else
                    If pcTpoProdCod = "751" Then   'MADM 20101120
                        RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'P'"
                        If RFiltro.RecordCount = 0 Then
                            bAplicaGasto = True
                        Else
                            RFiltroInst.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & pcTpoProdCod & " AND cAgeCod = '" & pcAgeCodAct & "' AND cTpoFiltro = 'P' AND cIntitucion='" & cInstitucion & "'"
                            If RFiltroInst.RecordCount > 0 Then
                                bAplicaGasto = True
                            Else
                                bAplicaGasto = False
                            End If
                        End If
                     Else
                        RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & "001" & " AND cAgeCod = '" & pcAgeCodAct & "' AND cTpoFiltro = 'P'"
                            If RFiltro.RecordCount > 0 Then
                                bAplicaGasto = True
                            Else
                                'MAVM 20120625 ***
                                'bAplicaGasto = False
                                If R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipoteca And psTipoSegDes <> "" Then
                                    bAplicaGasto = True
                                Else
                                    bAplicaGasto = False
                                End If
                                '***
                            End If
                     End If
                End If
                'Si tiene Filtro Por Tipos de Garantias
                If bAplicaGasto Then
                    bAplicaGasto = False
                    RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " AND cTpoFiltro = 'G'"
                    If RFiltro.RecordCount = 0 Then
                        bAplicaGasto = True
                    End If
                    
                    If Not bAplicaGasto Then
                        RGar.MoveFirst
                        Do While Not RGar.EOF
                            If Not bAplicaGasto Then
                                RFiltro.Filter = " nPrdConceptoCod = " & R!nPrdConceptoCod & " And nProdCod = " & RGar!nTpoGar & " AND cAgeCod = '" & pcAgeCodAct & "' AND cTpoFiltro = 'G'"
                                If RFiltro.RecordCount > 0 Then
                                    bAplicaGasto = True
                                    Exit Do
                                Else
                                    bAplicaGasto = False
                                End If
                            End If
                            RGar.MoveNext
                        Loop
                    End If
                End If
                    
                If bAplicaGasto Then
                    If CDbl(MatDesemb(0, 1)) >= R!nInicial And CDbl(MatDesemb(0, 1)) <= R!nFinal Then '-> Filtro por Monto de Prestamo
                        'Si el gasto se aplica por relaciones de credito
                        If R!cFiltro = "I" Or R!cFiltro = "M" Then
                            If sTipoRelacGarantia = "I" Then 'Si el Credito es de Tipo Individual
                                If R!cFiltro = "I" Then
                                    'Aplico a todos el mismo porcentaje
                                    nMontoTemp = 0
                                    oCredRel.IniciarMatriz
                                    Do While Not oCredRel.EOF
                                            'Filtramos por Edad
                                            bCumpleRestricc = False
                                            If Not IsNull(R!nEdadOper) Then
                                                Select Case R!nEdadOper
                                                    Case 2 'Mayor
                                                        If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 3 'Mayor Igual
                                                        If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 4 'Menor
                                                        If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 5 'Menor Igual
                                                        If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 6 'Igual a
                                                        If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                End Select
                                            Else
                                                bCumpleRestricc = True
                                            End If
                                            
                                            If bCumpleRestricc Then
                                            'Si es Fijo
                                                If R!nTpoValor = 1 Then
                                                    nMontoTemp = nMontoTemp + Format(nValor, "#0.000000")
                                                End If
                                                'Si es Porcentaje
                                                If R!nTpoValor = 2 Then
                                                    If R!cAplicaMonto = "C" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * CDbl(MatCalend(i, 2)), "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "D" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nPrestamo, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "H" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "S" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "T" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                                    End If
                                                    'DAOR 20061213, Para el caso del seguro de desgravamen
                                                    If R!cAplicaMonto = "I" Or R!cAplicaMonto = "K" Then 'JUEZ 20140310 Se agregó R!cAplicaMonto = "K"
                                                        If i = 0 Then
                                                            nRes = Format(Format(nValor / 100, "#0.000000") * (nPrestamo + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                                        Else
                                                            nRes = Format(Format(nValor / 100, "#0.000000") * (CDbl(MatCalend(i - 1, 7)) + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                                        End If
                                                    End If
                                                    '-------------------------------------------------------
                                                    If nRes < R!nMontoMin Then
                                                        nRes = R!nMontoMin
                                                    End If
                                                    If nRes > R!nMontoMax Then
                                                       nRes = R!nMontoMax
                                                    End If
                                                    nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                End If
                                            End If
                                        oCredRel.siguiente
                                    Loop
                                    MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                Else
                                    Exit For
                                End If
                            Else 'Si es Mancomunada cobrar el porcentaje mancomunado todos menos al Titular y al resto el gasto de Mancomunado
                                 'Para desiguales al Titular
                                 If R!cFiltro = "M" Then
                                    nMontoTemp = 0
                                    oCredRel.IniciarMatriz
                                    Do While Not oCredRel.EOF
                                        If oCredRel.ObtenerValorRelac <> gColRelPersTitular Then
                                            'Filtramos por Edad
                                            bCumpleRestricc = False
                                            If Not IsNull(R!nEdadOper) Then
                                                Select Case R!nEdadOper
                                                    Case 2 'Mayor
                                                        If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 3 'Mayor Igual
                                                        If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 4 'Menor
                                                        If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 5 'Menor Igual
                                                        If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 6 'Igual a
                                                        If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                End Select
                                            Else
                                                bCumpleRestricc = True
                                            End If
                                        
                                            'Si es Fijo
                                            If bCumpleRestricc Then
                                                If R!nTpoValor = 1 Then
                                                    nMontoTemp = nMontoTemp + Format(nValor, "#0.000000")
                                                End If
                                                'Si es Porcentaje
                                                If R!nTpoValor = 2 Then
                                                    If R!cAplicaMonto = "C" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "D" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nPrestamo, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "H" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "T" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "S" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                                    End If
                                                    'DAOR 20061213, Para el caso del seguro de desgravamen
                                                    If R!cAplicaMonto = "I" Or R!cAplicaMonto = "K" Then 'JUEZ 20140310 Se agregó R!cAplicaMonto = "K"
                                                        If i = 0 Then
                                                            nRes = Format(Format(nValor / 100, "#0.000000") * (nPrestamo + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                                        Else
                                                            nRes = Format(Format(nValor / 100, "#0.000000") * (CDbl(MatCalend(i - 1, 7)) + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                                        End If
                                                    End If
                                                    If nRes < R!nMontoMin Then
                                                        nRes = R!nMontoMin
                                                    End If
                                                    If nRes > R!nMontoMax Then
                                                       nRes = R!nMontoMax
                                                    End If
                                                    nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                End If
                                            End If
                                        End If
                                        oCredRel.siguiente
                                    Loop
                                                                            
                                    MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                    
                                 End If
                                 
                                 If R!cFiltro = "I" Then
                                    nMontoTemp = 0
                                    oCredRel.IniciarMatriz
                                    Do While Not oCredRel.EOF
                                        If oCredRel.ObtenerValorRelac = gColRelPersTitular Then
                                            'Filtramos por Edad
                                            bCumpleRestricc = False
                                            If Not IsNull(R!nEdadOper) Then
                                                Select Case R!nEdadOper
                                                    Case 2 'Mayor
                                                        If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 3 'Mayor Igual
                                                        If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 4 'Menor
                                                        If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 5 'Menor Igual
                                                        If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                    Case 6 'Igual a
                                                        If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                                            bCumpleRestricc = True
                                                        End If
                                                End Select
                                            Else
                                                bCumpleRestricc = True
                                            End If
                                        
                                            'Si es Fijo
                                            If bCumpleRestricc Then
                                                If R!nTpoValor = 1 Then
                                                    nMontoTemp = nMontoTemp + Format(nValor, "#0.000000")
                                                End If
                                                'Si es Porcentaje
                                                If R!nTpoValor = 2 Then
                                                    If R!cAplicaMonto = "C" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "D" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nPrestamo, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "H" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "S" Then
                                                        nRes = Format(Format(nValor / 100, "#0.0000000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                                    End If
                                                    If R!cAplicaMonto = "T" Then
                                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                                    End If
                                                    'DAOR 20061213, Para el caso del seguro de desgravamen
                                                    If R!cAplicaMonto = "I" Or R!cAplicaMonto = "K" Then 'JUEZ 20140310 Se agregó R!cAplicaMonto = "K"
                                                        If i = 0 Then
                                                            nRes = Format(Format(nValor / 100, "#0.000000") * (nPrestamo + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                                        Else
                                                            nRes = Format(Format(nValor / 100, "#0.000000") * (CDbl(MatCalend(i - 1, 7)) + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                                        End If
                                                    End If
                                                    If nRes < R!nMontoMin Then
                                                        nRes = R!nMontoMin
                                                    End If
                                                    If nRes > R!nMontoMax Then
                                                       nRes = R!nMontoMax
                                                    End If
                                                    nMontoTemp = nMontoTemp + Format(nRes, "#0.000000")
                                                End If
                                            End If
                                        End If
                                        oCredRel.siguiente
                                    Loop
                                    MatCalend(i, 6) = Format(nMontoTemp, "#0.00")
                                 End If
                            End If
                        Else    'Si no se considera las Relaciones entonces
                            'Filtramos por Edad
                            bCumpleRestricc = False
                            oCredRel.nPuntMat = oCredRel.PosicionTitular
                            If Not IsNull(R!nEdadOper) Then
                                Select Case R!nEdadOper
                                    Case 2 'Mayor
                                        If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then
                                            bCumpleRestricc = True
                                        End If
                                    Case 3 'Mayor Igual
                                        If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then
                                            bCumpleRestricc = True
                                        End If
                                    Case 4 'Menor
                                        If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then
                                            bCumpleRestricc = True
                                        End If
                                    Case 5 'Menor Igual
                                        If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then
                                            bCumpleRestricc = True
                                        End If
                                    Case 6 'Igual a
                                        If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then
                                            bCumpleRestricc = True
                                        End If
                                End Select
                            Else
                                bCumpleRestricc = True
                            End If
                        
                            'Si es Fijo
                            If bCumpleRestricc Then
                                If R!nTpoValor = 1 Then
                                   'madm 20091110 MatCalend(i, 6) = Format(nValor, "#0.00")
                                    If psTipoSegDes = "" Then
                                        If R!nPrdConceptoCod = 1224 Then
                                            MatCalend(i, 6) = IIf(R!nPrdConceptoCod = 1224, Format(getValor1224(pnPlazo, nValor), "#0.00"), Format(nValor, "#0.00"))
                                        End If
                                    Else
                                        If R!nPrdConceptoCod = 1224 Then
                                            nGastos1224 = IIf(R!nPrdConceptoCod = 1224, Format(getValor1224(pnPlazo, nValor), "#0.00"), Format(nValor, "#0.00"))
                                            MatCalend(i, 6) = Format(nGastos1224, "#0.00")
                                        ElseIf R!nPrdConceptoCod = 1247 Then
                                            MatCalend(i, 6) = IIf(R!nPrdConceptoCod = 1247, Format(nValor, "#0.00"), 0) '+ CDbl(nGastos1224)
                                        End If
                                    End If
                                   'end madm

                                End If
                                'Si es Porcentaje
                                If R!nTpoValor = 2 Then
                                    If R!cAplicaMonto = "C" Then
                                        nRes = Format(Format(nValor / 100, "#0.000000") * nCuotaApr, "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "D" Then
                                        nRes = Format(Format(nValor / 100, "#0.000000") * nPrestamo, "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "H" Then
                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoGarantia, "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "S" Then
                                        nRes = Format(Format(nValor / 100, "#0.000000") * CDbl(MatCalend(i, 7)), "#0.000000")
                                    End If
                                    If R!cAplicaMonto = "T" Then
                                        nRes = Format(Format(nValor / 100, "#0.000000") * nMontoConstruccion, "#0.000000")
                                    End If
                                    'DAOR 20061213, Para el caso del seguro de desgravamen
                                    If R!cAplicaMonto = "I" Or R!cAplicaMonto = "K" Then 'JUEZ 20140310 Se agregó R!cAplicaMonto = "K"
                                        If i = 0 Then
                                            nRes = Format(Format(nValor / 100, "#0.000000") * (nPrestamo + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                        Else
                                            nRes = Format(Format(nValor / 100, "#0.000000") * (CDbl(MatCalend(i - 1, 7)) + IIf(R!cAplicaMonto = "K", 0, CDbl(MatCalend(i, 4)))), "#0.000000") 'JUEZ 20140310 Se modificó para no sumar Int. Comp. en caso sea K
                                        End If
                                    End If
                                    If nRes < R!nMontoMin Then
                                        nRes = R!nMontoMin
                                    End If
                                    If nRes > R!nMontoMax Then
                                       nRes = R!nMontoMax
                                    End If
                                    MatCalend(i, 6) = Format(nRes, "#0.00")
                                End If
                                
                                'MAVM 20120626 ***
                                If R!nTpoValor = 3 Then
                                    If psTipoSegDes <> "" Then
                                        If R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipoteca Then
                                            MatCalend(i, 6) = IIf(Round(0.022 / 100 * (pnValorInmueble * 0.76), 2) < ValorMinimoPrima(102740), ValorMinimoPrima(102740), Round(0.022 / 100 * (pnValorInmueble * 0.76), 2))
                                            If MatCalend(i, 1) < pnCuotaMivienda Then
                                                MatCalend(i, 6) = 0
                                            End If
                                        End If
                                    End If
                                End If
                                '***
                                
                            End If
                        End If
                        If R!bAplNumConCer Then 'Multiplicar por el numero de consultas a la central de riesgos
                            MatCalend(i, 6) = MatCalend(i, 6) * nNumConCer
                        End If
                        'If R!bAplNumConMicr Then  'Multiplicar por el numero de consultas Score Microfinanzas gitu 03-06-2009
                        If IIf(IsNull(R!bAplNumConMic), 0, R!bAplNumConMic) Then  'Multiplicar por el numero de consultas Score Microfinanzas gitu 03-06-2009
                            MatCalend(i, 6) = MatCalend(i, 6) * nNumConMicr
                        End If
                        If i = 0 Then
                            If R!bAplNumMeses Then 'Multiplicar por el numero de meses para la primera cuota
                                nMunMesesPorDia = Round(DateDiff("d", dFecDesemb, MatCalend(i, 0)) / 30, 0)
                                nNumMesesPorMes = DateDiff("m", dFecDesemb, MatCalend(i, 0))
                                MatCalend(i, 6) = MatCalend(i, 6) * IIf(nMunMesesPorDia >= nNumMesesPorMes, nMunMesesPorDia, nNumMesesPorMes)
                            End If
                        ElseIf pnPlazo >= 30 Then
                            nMunMesesPorDia = Round(pnPlazo / 30, 0)
                            MatCalend(i, 6) = MatCalend(i, 6) * nMunMesesPorDia
                        End If
                                               
                        If bCumpleRestricc And (CDbl(MatCalend(i, 6)) > 0 Or (CDbl(MatCalend(i, 6)) = 0 And i = 0) Or (R!nPrdConceptoCod = gColocConceptoCodGastoPolizaIncendioHipoteca)) Then
                            nNumGastos = nNumGastos + 1
                            MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                            MatGastos(nNumGastos - 1, 1) = i + 1
                            MatGastos(nNumGastos - 1, 2) = R!CDescripcion & Space(150) & R!nPrdConceptoCod
                                                                                                                            
                            If R!nMontoMensual = "1" Then
                                If pnTipoPeriodo <> 2 And pnPlazo <> 30 Then 'Si no es fechafija
                                    MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6) * (pnPlazo / 30)
                                Else
                                    MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
                                End If
                            Else
                                MatGastos(nNumGastos - 1, 3) = MatCalend(i, 6)
                            End If
                            If R!nPrdConceptoCod = gColocConceptoCodGastoSeguro7 Then
                                If pnExoSeguroDes = 1 Then
                                    MatGastos(nNumGastos - 1, 3) = Format(0, "#0.00")
                                End If
                            End If
                            
                            If R!nAplicado = gColocConceptoAplCuota Then
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    End If
                End If
            Next i
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set oCredRel = Nothing
End Sub

Private Function obtenerEspecifGastoEstCta(ByVal pdFecVenc As Date, ByVal pdFecPivot As Date) As Boolean
    obtenerEspecifGastoEstCta = False
    If DateDiff("d", Format(pdFecPivot, "dd/MM/yyyy"), Format(pdFecVenc, "dd/MM/yyyy")) >= 90 Then
         obtenerEspecifGastoEstCta = True
    End If
End Function

'MADM 20100706/20100903
Private Function GetCuotaAplicaExoneraSegDes(ByVal pnEdadRequerida As Integer, ByVal pMatCalend As Variant, ByVal pMatrizRel As COMDCredito.UCOMCredRela, Optional psCtaCod As String = "") As Variant 'WIOR 20130903 AGREGO  Optional psCtaCod As String = ""
    Dim lnMes, lnedad As Integer
    Dim i, it, ix As Integer
    Dim lnIndicadorNumeroCuota() As String
    Dim lnFecFuturaAplica, dFecDesembolso, dFecUltimaCuota As Date
    Dim lbAplicaSeguro As Boolean
    'WIOR 20130903 *********************************
    Dim oDCredito As COMDCredito.DCOMCredito
    Dim bExo As Boolean
    Dim rsSegDesDet  As ADODB.Recordset
    Set oDCredito = New COMDCredito.DCOMCredito
    'WIOR FIN **************************************
    lbAplicaSeguro = False
    pMatrizRel.IniciarMatriz
    i = 0
    it = 0
    ReDim lnIndicadorNumeroCuota(pMatrizRel.NroRelaciones)
    
    dFecDesembolso = CDate(pMatCalend(0, 0))
    dFecUltimaCuota = CDate(pMatCalend(UBound(pMatCalend) - 1, 0))
    
    Do While Not pMatrizRel.EOF
            lnIndicadorNumeroCuota(i) = -1
            'WIOR 20130903 *********************************
            bExo = False
            Set rsSegDesDet = oDCredito.ListaPersonaExonerasSegDesg(psCtaCod, Trim(pMatrizRel.ObtenerCodigo))
            If Not (rsSegDesDet.EOF And rsSegDesDet.BOF) Then
                bExo = rsSegDesDet!bExonera
            End If
            Set rsSegDesDet = Nothing
            If Not bExo Then
            'WIOR FIN **************************************
                If CInt(pMatrizRel.ObtenerValorRelac) = COMDConstantes.gColRelPersTitular Or CInt(pMatrizRel.ObtenerValorRelac) = COMDConstantes.gColRelPersCodeudor Then
                        lnedad = Round(DateDiff("yyyy", CDate(pMatrizRel.ObtenerValorFechaNac), CDate(Date)), 0)
                        lnMes = Round((DateDiff("d", dFecDesembolso, dFecUltimaCuota) / 30), 0)
                        lnFecFuturaAplica = DateAdd("YYYY", pnEdadRequerida, CDate(pMatrizRel.ObtenerValorFechaNac))
                        
                        If DateAdd("m", lnMes, DateAdd("YYYY", lnedad, CDate(pMatrizRel.ObtenerValorFechaNac))) >= lnFecFuturaAplica Then
                           lbAplicaSeguro = True
                        End If
                    
                        If lbAplicaSeguro Then
                            For ix = 0 To UBound(pMatCalend) - 1
                                If CDate(pMatCalend(ix, 0)) >= lnFecFuturaAplica Then
                                    lnIndicadorNumeroCuota(it) = pMatCalend(ix, 1)
                                    Exit For
                                End If
                            Next ix
                        End If
                    it = it + 1
                End If
            End If 'WIOR 20130903
        i = i + 1
        pMatrizRel.siguiente
    Loop
    GetCuotaAplicaExoneraSegDes = lnIndicadorNumeroCuota
End Function
'MADM 20100903/20110425 - polizas
Private Function ValMismaGarPoliza(ByVal psCtaCod As String) As Boolean
    Dim oGarPol As COMDCredito.DCOMGarantia
    Dim oGar As COMDCredito.DCOMGarantia
    Dim RGarPol As ADODB.Recordset
    Dim cCtaCodPrincipal As String
    Dim cNumGarantParaPolParalelo As String
    Dim cNumGarantParaPol As String

    ValMismaGarPoliza = False

    Set oGarPol = New COMDCredito.DCOMGarantia
    Set RGarPol = New ADODB.Recordset
    Set RGarPol = oGarPol.RecuperaGarantiaCredito(psCtaCod)
    Set oGarPol = Nothing

    If Not (RGarPol.EOF Or RGarPol.BOF) Then
        cNumGarantParaPol = RGarPol!cNumGarantP
        RGarPol.Close
        Set RGarPol = Nothing
    Else
        Exit Function
    End If

    Dim oAmpliadoPol As COMDCredito.DCOMAmpliacion
    Dim Rt As ADODB.Recordset
    Set Rt = New ADODB.Recordset
    Set oAmpliadoPol = New COMDCredito.DCOMAmpliacion
    Set Rt = oAmpliadoPol.ObtenerCreditoPrincipalPol(psCtaCod)
    Set oAmpliadoPol = Nothing

    If Not (Rt.EOF Or Rt.BOF) Then
       cCtaCodPrincipal = Rt!cCtaCod
       Rt.Close
       Set Rt = Nothing
    Else
        Exit Function
    End If

    Dim oGarPolT As COMDCredito.DCOMGarantia
    Dim Rtp As ADODB.Recordset
    Set oGarPolT = New COMDCredito.DCOMGarantia
    Set Rtp = New ADODB.Recordset
    Set Rtp = oGarPolT.RecuperaGarantiaCredito(cCtaCodPrincipal)
    Set oGarPolT = Nothing
    If Not (Rtp.EOF Or Rtp.BOF) Then
       cNumGarantParaPolParalelo = Rtp!cNumGarantP
       Rtp.Close
       Set Rtp = Nothing
    Else
        Exit Function
    End If

    If cNumGarantParaPol = cNumGarantParaPolParalelo Then
        ValMismaGarPoliza = True
        Exit Function
    End If
    
End Function

Private Function AplicaGastoPoliza(ByVal bMismaPol As Boolean, ByVal pbEsParalelo As Boolean, ByVal pbEsRecurrentePol As Boolean, pbEsRefinanciadoPol As Boolean) As Boolean
   AplicaGastoPoliza = False
'   If (pbEsParalelo = True Or pbEsRecurrentePol = True Or pbEsRefinanciadoPol = True) And (bMismaPol) Then
    If (pbEsParalelo = True) And (bMismaPol) Then
        AplicaGastoPoliza = True
        Exit Function
    End If
End Function

Private Function Validar(ByVal bEsParalelo As Boolean, ByVal bEsRecurrentePol As Boolean, ByVal bEsRefinanciadoPol As Boolean, ByVal nTotalCuotas As Integer, ByVal nValorUltCuota As Double, ByVal psCtaCod As String) As Boolean
    Dim bMismaPol As Boolean
    bMismaPol = False
    'If bEsParalelo Or bEsRecurrentePol Or bEsRefinanciadoPol Then
    If bEsParalelo Then
         If nTotalCuotas > 0 And nValorUltCuota > 0 Then
              bMismaPol = ValMismaGarPoliza(psCtaCod)
         End If
    End If
    Validar = bMismaPol
End Function

'MADM 20110701
Public Function DevolverEstadoProducto(ByVal psCtaCod As String) As String
    Dim R As ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "SELECT nPrdEstado FROM Producto WHERE cCtaCod='" & psCtaCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
        If Not R.EOF Then
            DevolverEstadoProducto = Trim(R!nPrdEstado)
        End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'END MADM

'MADM 20111209
Public Function ValorMinimoPrima(ByVal psConsGarantia As String) As Double
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset

    On Error GoTo ErrorValorMinimoPrima
    
    sSql = "Select nParamValor from ColocParametro Where nParamVar = " & psConsGarantia
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        ValorMinimoPrima = R!nParamValor
    Else
        ValorMinimoPrima = 0
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    R.Close
    Set R = Nothing
    Exit Function

ErrorValorMinimoPrima:
    Set oConecta = Nothing
    Set R = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'END MADM

Private Sub DesgravCredHipot(ByRef sTipoRelacGarantia As String, ByVal bExisteMasDeUnTitular As Boolean, ByRef sTipoSegDes As String, Optional ByVal sCtaCod As String)
    If sCtaCod = "" Then
        If sTipoSegDes = "I" Then
            sTipoRelacGarantia = "I"
            bExisteMasDeUnTitular = False
        ElseIf sTipoSegDes = "M" Then
            sTipoRelacGarantia = "M"
            bExisteMasDeUnTitular = True
        End If
    Else
        If sTipoRelacGarantia = "I" Then
            If Mid(sCtaCod, 6, 3) = "801" Then sTipoSegDes = "I"
        End If
        If sTipoRelacGarantia = "M" Then
            If Mid(sCtaCod, 6, 3) = "801" Then sTipoSegDes = "M"
        End If
    End If
End Sub

Private Sub ResetVariables(ByVal psCtaCod As String, ByRef nDiasAtraso As Integer, ByRef nNumConCer As Integer, ByRef nNumConMicr As Integer, ByRef nNroProxCuota As Integer, ByRef bEsParalelo As Boolean, ByRef bEsRecurrentePol As Boolean, ByRef bEsRefinanciadoPol As Boolean)
    Dim oCred As COMDCredito.DCOMCredito
    Set oCred = New COMDCredito.DCOMCredito
    Dim R As ADODB.Recordset
    Set R = oCred.RecuperaColocacCred(psCtaCod)
    nDiasAtraso = IIf(IsNull(R!nDiasAtraso), 0, R!nDiasAtraso)
    nNumConCer = IIf(IsNull(R!nNumConCer), 0, R!nNumConCer)
    nNumConMicr = IIf(IsNull(R!nNumConMic), 0, R!nNumConMic)
    nNroProxCuota = R!nNroProxCuota
    bEsParalelo = IIf(R!nColocCondicion = 3, True, False)
    bEsRecurrentePol = IIf(R!nColocCondicion = 2, True, False)
    bEsRefinanciadoPol = IIf(R!nColocCondicion = 4, True, False)
    R.Close
    Set oCred = Nothing
End Sub

'JUEZ 20130529 *************************************************
Public Sub ObtieneValoresEnvioEstadoCta(ByRef pbEnvioEstCta As Boolean, ByRef pnCostoEnvioEstCta As Double, ByVal psCtaCod As String)
    Dim rsEnvio As ADODB.Recordset
    Dim pnCosto As Double
    Dim oEnvEst As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oEnvEst = New COMDCaptaGenerales.DCOMCaptaGenerales
    pbEnvioEstCta = False
    pnCostoEnvioEstCta = 0
    Set rsEnvio = oEnvEst.RecuperaDatosEnvioEstadoCta(psCtaCod)
    
    Dim rsCosto As ADODB.Recordset
    Dim oCostConc As COMDCredito.DCOMCredito
    Set oCostConc = New COMDCredito.DCOMCredito
    Set rsCosto = oCostConc.RecuperaProductoConcepto(gColocConceptoCodGastoComisionEnvioEstCta)
    pnCosto = CDbl(rsCosto!nValor)
    Set oCostConc = Nothing
    
    Do While Not rsEnvio.EOF
        If IIf(IsNull(rsEnvio!nModoEnvio), 0, rsEnvio!nModoEnvio) = 2 Then
            pbEnvioEstCta = True
            pnCostoEnvioEstCta = pnCostoEnvioEstCta + pnCosto
        End If
        rsEnvio.MoveNext
    Loop
    Set oEnvEst = Nothing
End Sub
'END JUEZ ******************************************************
'JUEZ 20140310 *************************************************
Public Function CalculaSeguroDesgravamen(ByVal i As Integer, ByVal psAplicaMonto As String, ByVal pnValor As Double, _
                                        ByVal pnPrestamo As Double, ByVal pnIntComp As Double, pnSaldoCapAnt, _
                                        Optional ByVal pdFechaCuota As Date = "01/01/1900", _
                                        Optional ByVal pnValorUnoTit As Double = 0, _
                                        Optional ByVal pnValorDosTit As Double = 0, Optional ByVal pMatPersSegDes As Variant, _
                                        Optional ByVal pnEdadMin As Long = 0, Optional ByVal pnEdadMax As Long = 0, _
                                        Optional ByVal pnMontoMin As Double = 0, Optional ByVal pnMontoMax As Double = 0, _
                                        Optional ByRef pMatCalendSegDes As Variant, _
                                        Optional ByVal psCtaCod As String = "") As Double
                                        'WIOR 20140822 AGREGO pdFechaCuota,pnValorDosTit, pMatPersSegDes, pnEdadMin, pnEdadMax, pnMontoMin, pnMontoMax y pMatCalendSegDes
                                        'JUEZ 20150714 psCtaCod
                                         
    'WIOR 20140822 ***************************
    Dim nEdadCuota As Long
    Dim Cant As Integer
    Dim Contador As Integer
    Dim bAplicaSegDes As Boolean
    Dim nSaldoCalc As Double
    Dim nSaldoCober As Double
    Dim nTasaCalc As Double
    
    'APRI20171122 ERS028-2017
    Dim CumpleEdadMinima As Boolean
    Dim CumpleEdadMaxima As Boolean
    Dim oParametro As COMDCredito.DCOMParametro
    Set oParametro = New COMDCredito.DCOMParametro
    'END APRI
    
    'JUEZ 20150714 ********************************************
    Dim RGas As ADODB.Recordset
    Dim oGas As COMDCredito.DCOMGasto
    Set oGas = New COMDCredito.DCOMGasto
    Set RGas = oGas.RecuperaProductoConceptoHist(1217, psCtaCod)
    If Not (RGas.EOF And RGas.BOF) Then
        pnValorUnoTit = RGas!nTasa1
        pnValorDosTit = RGas!nTasa2
    End If
    Set oGas = Nothing
    'END JUEZ *************************************************
    
    If IsArray(pMatPersSegDes) And Trim(pMatPersSegDes(0, 0)) <> "X" Then 'LUCV20180601, Modificó (Or por And) según ERS022-2018
        '0 - Orden
        '1 - cPersCod
        '2 - cPersNombre
        '3 - dPersNacCreac
        
        If i = 0 Then
            nSaldoCalc = pnPrestamo
        Else
            nSaldoCalc = pnSaldoCapAnt
        End If
        
        bAplicaSegDes = True
        
        If Trim(pMatPersSegDes(0, 0)) = "" Then
            Contador = 0
            CalculaSeguroDesgravamen = 0
            Exit Function
        Else
            Contador = 0
            nSaldoCober = 0
            
            For Cant = 0 To UBound(pMatPersSegDes, 2)
                nEdadCuota = EdadPersona(CDate(pMatPersSegDes(3, Cant)), pdFechaCuota)
                CumpleEdadMinima = oParametro.CumpleCriterioEdad(pMatPersSegDes(1, Cant), pdFechaCuota, 2) 'APRI20171122 ERS028-2017
               'If nEdadCuota >= pnEdadMin Then
                If Not CumpleEdadMinima Then 'APRI20171122 ERS028-2017
                    CumpleEdadMaxima = oParametro.CumpleCriterioEdad(pMatPersSegDes(1, Cant), pdFechaCuota, 3) 'APRI20171122 ERS028-2017
                    'If nEdadCuota >= pnEdadMax Then
                    If Not CumpleEdadMaxima Then
                        bAplicaSegDes = False
                    Else
                        Contador = Contador + 1
                    End If
                
                    If CInt(pMatPersSegDes(0, Cant)) = 1 Then
                        If Not bAplicaSegDes Then
                            Contador = 0
                            CalculaSeguroDesgravamen = 0
                            Exit Function
                        Else
                            nSaldoCober = pnMontoMin
                        End If
                    End If
                Else
                    Contador = Contador + 1
                    If CInt(pMatPersSegDes(0, Cant)) = 1 Then
                        nSaldoCober = pnMontoMax
                    End If
                End If
            Next Cant
        End If
        
        If Contador > 1 Then
            nTasaCalc = pnValorDosTit
        Else
            nTasaCalc = pnValorUnoTit
        End If
        
        If nSaldoCalc > nSaldoCober Then
            nSaldoCalc = nSaldoCober
        End If
        
        If IsArray(pMatCalendSegDes) Then
            ReDim Preserve pMatCalendSegDes(4, 0 To i)
            pMatCalendSegDes(0, i) = i + 1
            pMatCalendSegDes(1, i) = Contador
            pMatCalendSegDes(2, i) = nSaldoCalc
            pMatCalendSegDes(3, i) = Round(nTasaCalc, 6)
        End If
        
        CalculaSeguroDesgravamen = Format(Format(nTasaCalc / 100, "#0.000000") * (nSaldoCalc + IIf(psAplicaMonto = "K", 0, pnIntComp)), "#0.000000")
    Else
        If i = 0 Then
            CalculaSeguroDesgravamen = Format(Format(pnValor / 100, "#0.000000") * (pnPrestamo + IIf(psAplicaMonto = "K", 0, pnIntComp)), "#0.000000")
        Else
            CalculaSeguroDesgravamen = Format(Format(pnValor / 100, "#0.000000") * (pnSaldoCapAnt + IIf(psAplicaMonto = "K", 0, pnIntComp)), "#0.000000")
        End If
    End If
    'WIOR FIN ********************************
End Function
'END JUEZ ******************************************************

'WIOR 20140822 *****************************************
Public Function EdadPersona(ByVal pdFecNac As Date, Optional ByVal pdFecha As Date = "01/01/1900") As Long
Dim nEdad As Integer

nEdad = DateDiff("yyyy", pdFecNac, pdFecha)

If Month(pdFecNac) >= Month(pdFecha) Then
    If Month(pdFecNac) = Month(pdFecha) Then
        If Day(pdFecNac) > Day(pdFecha) Then
            nEdad = nEdad - 1
        End If
    Else
        nEdad = nEdad - 1
    End If
End If

EdadPersona = nEdad
End Function

Public Function CargaPersRelacSegDes(ByVal psCtaCod As String, ByVal pdFecha As Date, ByVal pnEdadMin As Long, Optional ByVal pbReprogramado As Boolean = False) As Variant 'Add JOEP20200414 covid
'Public Function CargaPersRelacSegDes(ByVal psCtaCod As String, ByVal pdFecha As Date, ByVal pnEdadMin As Long) As Variant 'Comento JOEP20200414 covid
Dim oCredito As COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Dim nEdadActual As Long
Dim i As Integer
Dim Contador As Integer
Dim MatPersSegDes As Variant
'APRI20171122 ERS028-2017
Dim CumpleEdadMinima As Boolean
Dim oParametro As COMDCredito.DCOMParametro
'END APRI
Contador = -1
ReDim MatPersSegDes(4, 0)
If Len(Trim(psCtaCod)) = 18 Then
    Set oCredito = New COMDCredito.DCOMCredito
    Set oParametro = New COMDCredito.DCOMParametro 'APRI20171122 ERS028-2017
    Set rsCredito = oCredito.ListaPersonaExonerasSegDesg(psCtaCod, "%", "0,1")
    
    If Not (rsCredito.EOF And rsCredito.BOF) Then
        For i = 0 To rsCredito.RecordCount - 1
            'CumpleEdadMinima = oParametro.CumpleCriterioEdad(rsCredito!cPersCod, pdFecha, 2) 'APRI20171122 ERS028-2017 'Comento JOEP20200414 covid
            'add JOEP20200414 covid
            If pbReprogramado = True Then
                CumpleEdadMinima = True
            Else
                CumpleEdadMinima = oParametro.CumpleCriterioEdad(rsCredito!cPersCod, pdFecha, 2)
            End If
            'add JOEP20200414 covid
            If CBool(rsCredito!bExonera) = False And CumpleEdadMinima Then 'APRI20171122 ERS028-2017
            'nEdadActual = EdadPersona(CDate(rsCredito!dPersNacCreac), pdFecha) 'APRI20171012 MEJORA CORREO20171011-SEOR
            'If CBool(rsCredito!bExonera) = False And nEdadActual < pnEdadMin Then 'APRI20171012 MEJORA CORREO20171011-SEOR ADD And nEdadActual < pnEdadMin
                Contador = Contador + 1
                ReDim Preserve MatPersSegDes(4, 0 To Contador)
                MatPersSegDes(0, Contador) = CInt(rsCredito!Orden)
                MatPersSegDes(1, Contador) = Trim(rsCredito!cPersCod)
                MatPersSegDes(2, Contador) = Trim(rsCredito!cPersNombre)
                MatPersSegDes(3, Contador) = CDate(rsCredito!dPersNacCreac)
            End If
            rsCredito.MoveNext
        Next i
    Else
        Set rsCredito = Nothing
        Set rsCredito = oCredito.RecuperaRelacPersSegDes(psCtaCod)

        If Not (rsCredito.EOF And rsCredito.BOF) Then
            For i = 0 To rsCredito.RecordCount - 1
                If Contador >= 1 Then
                    Call oCredito.ExoneraPersSegDesgOpe(1, Trim(psCtaCod), Trim(rsCredito!cPersCod), True, "Persona exonerada por el sistema, ya que no se puede asegurar a mas de 2 personas o si el titular del seguro fue exonerado.", True)
                Else
                    CumpleEdadMinima = oParametro.CumpleCriterioEdad(rsCredito!cPersCod, pdFecha, 2) 'APRI20171122 ERS028-2017
                    If Not CumpleEdadMinima Then 'APRI20171122 ERS028-2017
                    'nEdadActual = EdadPersona(CDate(rsCredito!dPersNacCreac), pdFecha)
                    'If nEdadActual >= pnEdadMin Then
                        If CInt(rsCredito!Orden) = 1 Then
                            ReDim Preserve MatPersSegDes(4, 0 To 0)
                            Contador = 1
                            Call oCredito.ExoneraPersSegDesgOpe(1, Trim(psCtaCod), Trim(rsCredito!cPersCod), True, "Credito exonerado por el sistema, ya que el Titular del Seguro supera el límite de edad de ingreso al Seg. Desgravamen que es " & pnEdadMin - 1 & " años 11 meses con 29 días.", True)
                        Else
                            Call oCredito.ExoneraPersSegDesgOpe(1, Trim(psCtaCod), Trim(rsCredito!cPersCod), True, "Persona exonerada por el sistema, por superar el límite de edad de ingreso al Seg. Desgravamen que es " & pnEdadMin - 1 & " años 11 meses con 29 días.", True)
                        End If
                    Else
                        Contador = Contador + 1
                        ReDim Preserve MatPersSegDes(4, 0 To Contador)
                        MatPersSegDes(0, Contador) = CInt(rsCredito!Orden)
                        MatPersSegDes(1, Contador) = Trim(rsCredito!cPersCod)
                        MatPersSegDes(2, Contador) = Trim(rsCredito!cPersNombre)
                        MatPersSegDes(3, Contador) = CDate(rsCredito!dPersNacCreac)
                        Call oCredito.ExoneraPersSegDesgOpe(1, Trim(psCtaCod), Trim(rsCredito!cPersCod))
                    End If
                End If
                rsCredito.MoveNext
            Next i
        Else
            MatPersSegDes(0, 0) = "X"
        End If
    End If
Else
    MatPersSegDes(0, 0) = "X"
End If

CargaPersRelacSegDes = MatPersSegDes
Set oCredito = Nothing
Set rsCredito = Nothing
End Function

Public Sub CargaValoreSegDes(ByVal pcCtaCod As String, ByVal pdFechaSis As Date, ByRef MatPersSegDes As Variant, ByRef nEdadMinSegDes As Long, ByRef nEdadMaxSegDes As Long, ByRef nCoberMinSegDes As Double, ByRef nCoberMaxSegDes As Double, Optional ByVal pbPagoAnticipado As Boolean = False, Optional ByVal pbReprogramado As Boolean = False) 'Add 20200414 covid
'Public Sub CargaValoreSegDes(ByVal pcCtaCod As String, ByVal pdFechaSis As Date, ByRef MatPersSegDes As Variant, ByRef nEdadMinSegDes As Long, ByRef nEdadMaxSegDes As Long, ByRef nCoberMinSegDes As Double, ByRef nCoberMaxSegDes As Double, Optional ByVal pbPagoAnticipado As Boolean = False)'Comento 20200414 covid
'APRI20181025 ADD pbPagoAnticipado ERS-071-2108
    Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
    Dim oParam As COMDCredito.DCOMParametro
    Dim nTC As Double

    Set oParam = New COMDCredito.DCOMParametro
    nEdadMinSegDes = oParam.RecuperaValorParametro(3213)
    nEdadMaxSegDes = oParam.RecuperaValorParametro(3204)
    'nCoberMaxSegDes = oParam.RecuperaValorParametro(3214)
    'nCoberMinSegDes = oParam.RecuperaValorParametro(3215)
    nCoberMaxSegDes = oParam.RecuperaValorParametroDesg(1) 'APRI20171121 ERS028-2017
    nCoberMinSegDes = oParam.RecuperaValorParametroDesg(2) 'APRI20171121 ERS028-2017
    
    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
    nTC = oTipoCam.EmiteTipoCambio(pdFechaSis, TCFijoMes)
    
    If Mid(pcCtaCod, 9, 1) = "1" Then
        nCoberMinSegDes = nCoberMinSegDes * nTC
        nCoberMaxSegDes = nCoberMaxSegDes * nTC
    End If
    
    'MatPersSegDes = CargaPersRelacSegDes(pcCtaCod, pdFechaSis, nEdadMinSegDes)
    'APRI20181025 ERS071-2018
     Dim dVigencia As Date
     Dim oCred As COMDCredito.DCOMCredito
     Dim RCredVig As ADODB.Recordset
     
     Set oCred = New COMDCredito.DCOMCredito
        Set RCredVig = oCred.RecuperaColocaciones(pcCtaCod)
        If Not (RCredVig.BOF Or RCredVig.EOF) Then
            dVigencia = IIf(IsNull(RCredVig!dVigencia), pdFechaSis, RCredVig!dVigencia)
        End If
     Set RCredVig = Nothing
     Set oCred = Nothing
     'MatPersSegDes = CargaPersRelacSegDes(pcCtaCod, IIf(pbPagoAnticipado = True, dVigencia, pdFechaSis), nEdadMinSegDes)'Comento JOEP20200414 covid
     MatPersSegDes = CargaPersRelacSegDes(pcCtaCod, IIf(pbPagoAnticipado = True, dVigencia, pdFechaSis), nEdadMinSegDes, pbReprogramado)
    'END APRI
    Set oParam = Nothing
    Set oTipoCam = Nothing
End Sub


Public Function CumpleRestriccion(ByVal pnTipo As Integer, ByVal R As ADODB.Recordset, ByRef oCredRel As COMDCredito.UCOMCredRela, ByVal pdHoy As Date) As Boolean
Dim bCumpleRestricc As Boolean
bCumpleRestricc = False

Select Case pnTipo
    Case 1:
        If Not IsNull(R!nEdadOper) Then
            Select Case R!nEdadOper
                Case 2 'Mayor
                    If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then bCumpleRestricc = True
                    'End If
                Case 3 'Mayor Igual
                    If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then bCumpleRestricc = True
                    'End If
                Case 4 'Menor
                    If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then bCumpleRestricc = True
                    'End If
                Case 5 'Menor Igual
                    If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then bCumpleRestricc = True
                    'End If
                Case 6 'Igual a
                    If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then bCumpleRestricc = True
                    'End If
            End Select
        Else
            bCumpleRestricc = True
        End If
    Case 2:
        If Not IsNull(R!nEdadOper) Then
            Select Case R!nEdadOper
                Case 2 'Mayor
                    If oCredRel.ObtenerValorEdad(pdHoy) > R!nEdad Then bCumpleRestricc = True
                Case 3 'Mayor Igual
                    If oCredRel.ObtenerValorEdad(pdHoy) >= R!nEdad Then bCumpleRestricc = True
                Case 4 'Menor
                    If oCredRel.ObtenerValorEdad(pdHoy) < R!nEdad Then bCumpleRestricc = True
                Case 5 'Menor Igual
                    If oCredRel.ObtenerValorEdad(pdHoy) <= R!nEdad Then bCumpleRestricc = True
                Case 6 'Igual a
                    If oCredRel.ObtenerValorEdad(pdHoy) = R!nEdad Then bCumpleRestricc = True
            End Select
        Else
            bCumpleRestricc = True
        End If
End Select


CumpleRestriccion = bCumpleRestricc
End Function
'WIOR FIN *********************************************
'RECO20150526 ERS023-2015******************************
Public Function RecuperaGastosMYPE(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oGastos As New COMDCredito.DCOMGasto
    Set RecuperaGastosMYPE = oGastos.RecuperaGastosMYPE(psCtaCod)
End Function
'RECO FIN *********************************************
'APRI 20170417
Public Function OperacionPrimeraCuota(ByVal pCuota As Integer, ByVal pbAplicaNumMeses, ByVal pnPrdConceptoCod As Integer, pdFecDesem As Date, pdFecPrimerCuota As Date, pnPlazo As Integer, pnMonto As Currency) As Currency
'Public Function OperacionPrimeraCuota(ByVal pCuota As Integer, ByVal pbAplicaNumMeses, ByVal pnPrdConceptoCod As Integer, pdFecDesem As Date, pdFecPrimerCuota As Date, pnPlazo As Integer, ByVal pnMonto As Currency) As Currency 'APRI20180925 ERS061-2018
Dim nMunMesesPorDia, nNumMesesPorMes  As Integer
Dim dFechaFinPrimeraCuotas As Date
 If pCuota = 0 Then
        'If pbAplicaNumMeses Or pnPrdConceptoCod = "1231" Or pnPrdConceptoCod = "1243" Or pnPrdConceptoCod = "1272" Then  'Multiplicar por el numero de meses para la primera cuota 'LUCV20180601, Comentó Según ERS022-2018
        If pnPrdConceptoCod = "1243" Or pnPrdConceptoCod = "1272" Then  'Multiplicar por el numero de meses para la primera cuota 'LUCV20180601, Agregó Según ERS022-2018
            nMunMesesPorDia = Round(DateDiff("d", pdFecDesem, pdFecPrimerCuota) / 30, 0)
            nNumMesesPorMes = DateDiff("m", pdFecDesem, pdFecPrimerCuota)
            
            'APRI 20170417
            If pnPrdConceptoCod <> "1217" Then
                If DateDiff("d", pdFecDesem, pdFecPrimerCuota) > pnPlazo Then
                    Dim nTotal As Single
                    
                    nTotal = pnPlazo / 30
                    If Round(nTotal, 0) < nTotal Then
                        nTotal = Round(nTotal, 0) + 1
                    Else
                        nTotal = Round(nTotal, 0)
                    End If
                    
                    pnMonto = pnMonto / nTotal
                    dFechaFinPrimeraCuotas = DateAdd("m", DateDiff("m", 0, pdFecPrimerCuota), 1)
                    nMunMesesPorDia = DateDiff("d", pdFecDesem, dFechaFinPrimeraCuotas)
                    nTotal = nMunMesesPorDia / 30
                    nTotal = Round(nTotal, 0)
                    pnMonto = pnMonto * nTotal
                End If
            'Else 'LUCV20180601, Comentó Según ERS022-2018 (El segDesg. Se calcula en el calendario)
            '    pnMonto = pnMonto * IIf(nMunMesesPorDia >= nNumMesesPorMes, nMunMesesPorDia, nNumMesesPorMes)'LUCV20180601, Comentó según ERS022-2018
            End If
             'END APRI
        End If
    'MAVM 20130410 *** Seg Desg Por Periodo
'->***** LUCV20180601, Comentó según ERS022-2018 (Se calcula en la generación del calendario)
'ElseIf pnPlazo >= 30 And pnPrdConceptoCod = "1217" Then 'APRI20170417 AGREGO R!nPrdConceptoCod = "1217"
'    nMunMesesPorDia = Round(pnPlazo / 30, 0)
'    pnMonto = pnMonto * nMunMesesPorDia
'***
'<-***** Fin LUCV20180601
End If
 OperacionPrimeraCuota = pnMonto
End Function
' EN APRI

'->***** LUCV20180601, Agregó según ERS022-2018
Public Function ObtenerTasaSeguroDesgPagoAnt(Optional ByVal pdFechaPago As Date = "01/01/1900", _
                                             Optional ByVal pMatPersSegDes As Variant, _
                                             Optional ByVal pnEdadMin As Long = 0, _
                                             Optional ByVal pnEdadMax As Long = 0, _
                                             Optional ByVal pnMontoMin As Double = 0, _
                                             Optional ByVal pnMontoMax As Double = 0, _
                                             Optional ByVal psCtaCod As String = "") As Double
   
                                         
    Dim nEdadCuota As Long
    Dim Cant As Integer
    Dim Contador As Integer
    Dim bAplicaSegDes As Boolean
    
    Dim nTasaCalc As Double
    
    'APRI20171122 ERS028-2017
    Dim CumpleEdadMinima As Boolean
    Dim CumpleEdadMaxima As Boolean
    Dim oParametro As COMDCredito.DCOMParametro
    Set oParametro = New COMDCredito.DCOMParametro
    'END APRI
    
    'LUCV20180601
    Dim pnValorUnoTit As Double
    Dim pnValorDosTit As Double
    pnValorUnoTit = 0
    pnValorDosTit = 0
    'Fin LUCV20180601
    
    Dim RGas As ADODB.Recordset
    Dim oGas As COMDCredito.DCOMGasto
    Set oGas = New COMDCredito.DCOMGasto
    Set RGas = oGas.RecuperaProductoConceptoHist(1217, psCtaCod)
    If Not (RGas.EOF And RGas.BOF) Then
        pnValorUnoTit = RGas!nTasa1
        pnValorDosTit = RGas!nTasa2
    End If
    Set oGas = Nothing
    
    If IsArray(pMatPersSegDes) And Trim(pMatPersSegDes(0, 0)) <> "X" Then 'LUCV20180601, Modificó (Or por And) según ERS022-2018
        '0 - Orden
        '1 - cPersCod
        '2 - cPersNombre
        '3 - dPersNacCreac
        bAplicaSegDes = True

        If Trim(pMatPersSegDes(0, 0)) = "" Then
            Contador = 0
            ObtenerTasaSeguroDesgPagoAnt = 0
            Exit Function
        Else
            Contador = 0

            For Cant = 0 To UBound(pMatPersSegDes, 2)
                nEdadCuota = EdadPersona(CDate(pMatPersSegDes(3, Cant)), pdFechaPago)
                CumpleEdadMinima = oParametro.CumpleCriterioEdad(pMatPersSegDes(1, Cant), pdFechaPago, 2) 'APRI20171122 ERS028-2017
               'If nEdadCuota >= pnEdadMin Then
                If Not CumpleEdadMinima Then 'APRI20171122 ERS028-2017
                    CumpleEdadMaxima = oParametro.CumpleCriterioEdad(pMatPersSegDes(1, Cant), pdFechaPago, 3) 'APRI20171122 ERS028-2017
                    'If nEdadCuota >= pnEdadMax Then
                    If Not CumpleEdadMaxima Then
                        bAplicaSegDes = False
                    Else
                        Contador = Contador + 1
                    End If
                
                    If CInt(pMatPersSegDes(0, Cant)) = 1 Then
                        If Not bAplicaSegDes Then
                            Contador = 0
                            ObtenerTasaSeguroDesgPagoAnt = 0
                            Exit Function
                        End If
                    End If
                Else
                    Contador = Contador + 1
                End If
            Next Cant
        End If
        
        If Contador > 1 Then
            nTasaCalc = pnValorDosTit
        Else
            nTasaCalc = pnValorUnoTit
        End If
        
        ObtenerTasaSeguroDesgPagoAnt = nTasaCalc
    Else
        ObtenerTasaSeguroDesgPagoAnt = 0
    End If
End Function

Public Function RecuperaMontoPoliza(ByVal psCtaCod As String, _
                                    ByVal pnTotalCuotas As Integer, _
                                    ByVal pnPrdConceptoCod As Integer, _
                                    Optional ByRef pnTasaSegInc As Double, _
                                    Optional ByVal pbEsParalelo As Boolean = False, _
                                    Optional ByVal pbEsRecurrentePol As Boolean = False, _
                                    Optional ByVal bEsRefinanciadoPol As Boolean = False, _
                                    Optional ByVal pbModReprogramado As Integer = 0) As Double
 'JOEP20200413 add pbModReprogramado para identificar si viene del modulo de reprogramacion
                                    
Dim oDCOMPoliza As COMDCredito.DCOMPoliza
Set oDCOMPoliza = New COMDCredito.DCOMPoliza
Dim bValidar As Boolean
Dim nValorCuota As Double
Dim nTasaSegInc As Double
nValorCuota = 0
nTasaSegInc = 0

Dim rPolizaTasaPrimaCuota As ADODB.Recordset
Set rPolizaTasaPrimaCuota = New ADODB.Recordset

   Set rPolizaTasaPrimaCuota = oDCOMPoliza.DamePrimaPolizaxCuotaTasaSegInc(psCtaCod, pnTotalCuotas, pnTotalCuotas, pnPrdConceptoCod)
    
    If Not rPolizaTasaPrimaCuota.EOF Then
    'JOEP20200413 Comento New Calculo reprogramado
        'nValorCuota = rPolizaTasaPrimaCuota!nMontoPagar
        'nTasaSegInc = rPolizaTasaPrimaCuota!nTasaSegInc
    'JOEP20200413 Comento New Calculo reprogramado
    'JOEP20200413 New Calculo reprogramado
        If pbModReprogramado = 1 Then
            nValorCuota = rPolizaTasaPrimaCuota!PrimaReprogramado
            nTasaSegInc = rPolizaTasaPrimaCuota!nTasaSegInc
        Else
            nValorCuota = rPolizaTasaPrimaCuota!nMontoPagar
            nTasaSegInc = rPolizaTasaPrimaCuota!nTasaSegInc
        End If
    'JOEP20200413 New Calculo reprogramado
    End If
    
    'Validaciones
    'Valida: si la garantía posee la misma poliza
    bValidar = Validar(pbEsParalelo, pbEsRecurrentePol, bEsRefinanciadoPol, pnTotalCuotas, pnTotalCuotas, psCtaCod)
    
    'Valida: Condicion Crédito
    If (AplicaGastoPoliza(bValidar, pbEsParalelo, pbEsRecurrentePol, bEsRefinanciadoPol)) Then
        nValorCuota = Format(0, "#0.00")
        pnTasaSegInc = Format(0, "#0.0000")
    Else
        nValorCuota = Format(nValorCuota, "#0.00")
        pnTasaSegInc = Format(nTasaSegInc, "#0.0000")
    End If
    
    Set oDCOMPoliza = Nothing
    rPolizaTasaPrimaCuota.Close
    Set rPolizaTasaPrimaCuota = Nothing
    
    RecuperaMontoPoliza = nValorCuota
End Function
'<-***** Fin LUCV20180601

