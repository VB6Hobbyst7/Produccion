VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NRFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Public Sub LiberaGarantiaPago(ByVal pnMovNro As Long, ByVal poBase As DCredActualizaBD, ByVal psCtaCod As String, ByVal psCodAge As String, ByVal psCodUser As String, _
    ByVal pdFecha As Date, ByVal pnPrestamo As Double, ByVal pnMontoCapital As Double, ByVal bCancelado As Boolean)

Dim oCred As DCredito
Dim R As ADODB.Recordset
Dim R2 As ADODB.Recordset
Dim oBase As DCredActualizaBD
Dim sSQL As String
Dim sMovNro As String
Dim nMovNro As Long
Dim sOpeCod As String
Dim nSumaTotalGarantia As Double
Dim nPago As Double
Dim nMontoLibera As Double
Dim nMontoLiberaTotal As Double ' Total a liberar

    Set oBase = poBase
    Set oCred = New DCredito
    Set R = oCred.RecuperaColocGarantia(psCtaCod, oBase.coConex.ConexionActiva)
    nSumaTotalGarantia = 0
    Do While Not R.EOF
        nSumaTotalGarantia = nSumaTotalGarantia + R!nGravado
        R.MoveNext
    Loop
    nSumaTotalGarantia = Format(nSumaTotalGarantia, "#0.00")
    nPago = Format(pnMontoCapital * nSumaTotalGarantia / pnPrestamo, "#0.00")
    R.Close
    
    nMontoLiberaTotal = 0
    
    Set R = oCred.RecuperaColocGarantia(psCtaCod, oBase.coConex.ConexionActiva)
    Do While Not R.EOF
        sSQL = "Select cOpeCod from GarantiaOpe where nTpoGarantia = " & R!nTpoGarantia & " and nMoneda = " & R!nmoneda & " AND bReversion = 1 "
        Set R2 = New ADODB.Recordset
        R2.Open sSQL, oBase.coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
        If R2.RecordCount > 0 Then
            sOpeCod = R2!copecod
        Else
            sOpeCod = ""
        End If
        R2.Close
        
        nMontoLibera = nPago * (R!nGravado / nSumaTotalGarantia)
        nMontoLiberaTotal = nMontoLiberaTotal + nMontoLibera
        
        'oBase.dBeginTrans
        
        sSQL = "UPDATE GARANTIAS SET nPorGravar = nPorGravar + " & Format(nMontoLibera, "#0.00") & ", "
            sSQL = sSQL & " nGravament = nGravament -  " & Format(nMontoLibera, "#0.00")
            sSQL = sSQL & " WHERE cNumGarant = '" & R!cNumGarant & "'"
            oBase.coConex.ConexionActiva.Execute sSQL
        If bCancelado Then
            sSQL = " UPDATE ColocGarantia Set nEstado = 0 WHERE cNumGarant = '" & R!cNumGarant & "' "
            sSQL = sSQL & " AND  cCtaCod = '" & psCtaCod & "'"
            oBase.coConex.ConexionActiva.Execute sSQL
        End If
        
        'sMovNro = oBase.GeneraMovNro(pdFecha, psCodAge, psCodUser)
        'Call oBase.dInsertMov(sMovNro, sOpeCod, "Asiento de Garantia", gMovEstContabMovContable, gMovFlagVigente, False)
        'nMovNro = oBase.dGetnMovNro(sMovNro)
        'nMovNro = pnMovNro
        
        'Call oBase.dInsertMovCol(nMovNro, sOpeCod, psCtaCod, 0, R!nGravado, 0, "", 0, 0, 0, False)
        'Call oBase.dInsertMovColDet(nMovNro, sOpeCod, psCtaCod, 0, 1000, 0, R!nGravado, False)
        
        'oBase.dCommitTrans
        
        R.MoveNext
    Loop
    ' LAYG - 2004/09/22   luis
    'sMovNro = oBase.GeneraMovNro(pdFecha, psCodAge, psCodUser)
    'Call oBase.dInsertMov(sMovNro, sOpeCod, "Asiento de Garantia", gMovEstContabMovContable, gMovFlagVigente, False)
    'nMovNro = oBase.dGetnMovNro(sMovNro)
    nMovNro = pnMovNro
    Call oBase.dInsertMovCol(nMovNro, sOpeCod, psCtaCod, 0, nMontoLiberaTotal, 0, "", 0, 0, 0, False)
    Call oBase.dInsertMovColDet(nMovNro, sOpeCod, psCtaCod, 0, 1000, 0, nMontoLiberaTotal, False)
    
    R.Close
    Set oCred = Nothing

End Sub

Public Sub AsientoGarantia(ByVal psCtaCod As String, ByVal psCodAge As String, ByVal psCodUser As String, _
    ByVal pdFecha As Date)
Dim oCred As DCredito
Dim R As ADODB.Recordset
Dim R2 As ADODB.Recordset
Dim oBase As DCredActualizaBD
Dim sSQL As String
Dim sMovNro As String
Dim nMovNro As Long
Dim sOpeCod As String

    Set oBase = New DCredActualizaBD
    Set oCred = New DCredito
    Set R = oCred.RecuperaColocGarantia(psCtaCod)
    Set oCred = Nothing
    
    Do While Not R.EOF
        sSQL = "Select cOpeCod from GarantiaOpe where nTpoGarantia = " & R!nTpoGarantia & " and nMoneda = " & R!nmoneda & " AND bReversion = 0 "
        Set R2 = New ADODB.Recordset
        R2.Open sSQL, oBase.coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
        If R2.RecordCount > 0 Then
            sOpeCod = R2!copecod
        End If
        sSQL = "Select cOpeCod from GarantiaOpe where nTpoGarantia = " & R!nTpoGarantia & " and nMoneda = " & R!nmoneda & " AND bReversion = 0 "
        Set R2 = New ADODB.Recordset
        R2.Open sSQL, oBase.coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
        If R2.RecordCount > 0 Then
            sOpeCod = R2!copecod
        Else
            sOpeCod = ""
        End If
        R2.Close
        
        sMovNro = oBase.GeneraMovNro(pdFecha, psCodAge, psCodUser)
        oBase.dBeginTrans
        
        Call oBase.dInsertMov(sMovNro, sOpeCod, "Asiento de Garantia", gMovEstContabMovContable, gMovFlagVigente, False)
        nMovNro = oBase.dGetnMovNro(sMovNro)
        Call oBase.dInsertMovCol(nMovNro, sOpeCod, psCtaCod, 0, R!nGravado, 0, "", 0, 0, 0, False)
        Call oBase.dInsertMovColDet(nMovNro, sOpeCod, psCtaCod, 0, 1000, 0, R!nGravado, False)
        
        oBase.dCommitTrans
        
        R.MoveNext
    Loop
    R.Close
End Sub

Public Sub ActualizaCreditosVigentes(ByVal pMatCredVig As Variant)
Dim i As Integer
Dim oBase As DCredActualizaBD

    If IsArray(pMatCredVig) Then
        If UBound(pMatCredVig) > 0 Then
            Set oBase = New DCredActualizaBD
            oBase.coConex.ConexionActiva.Execute "DELETE ColocCredCredVig Where cCtaCod = '" & pMatCredVig(0, 0) & "'"
            
            For i = 0 To UBound(pMatCredVig) - 1
                Call oBase.dInsertColocCredCredVig(pMatCredVig(i, 0), pMatCredVig(i, 1), False)
            Next i
        End If
    End If
    Set oBase = Nothing
End Sub

'**********************************************************************
'FUNCION QUE ME DEVUELVE LA TASA EFECTIVA
'**********************************************************************
Function CreditosTasaEfectiva(ByVal pnCuota As Double, ByVal pnCuotas As Integer, ByVal pnMonto As Double, ByVal pnTasaIni As Double, ByVal pnPeriodo As Double) As Double
Dim i As Double
Dim nTasaTemp As Double
Dim nTasaTemp2 As Double
Dim nMontoAprox1 As Double
Dim nMontoAprox2 As Double

    CreditosTasaEfectiva = 0
    nMontoAprox1 = -999999
    'pnTasaIni = pnTasaIni / 100
    nTasaTemp = pnTasaIni
    For i = pnTasaIni / 100 To 1 Step 0.0000001
        nTasaTemp2 = i
        nMontoAprox2 = CDbl(Format(CuotaFija(nTasaTemp2 * 100, pnCuotas, pnMonto, pnPeriodo), "#0.00"))
        If nMontoAprox2 = pnCuota Then
            CreditosTasaEfectiva = i
            Exit For
        Else
            If nMontoAprox2 - pnCuota > 0 Then
                If Abs(nMontoAprox2 - pnCuota) > Abs(nMontoAprox1 - pnCuota) Then
                    CreditosTasaEfectiva = i - 0.0000001
                    Exit For
                Else
                    CreditosTasaEfectiva = i
                    Exit For
                End If
            End If
        End If
        nMontoAprox1 = nMontoAprox2
    Next i
End Function




'*****************************************************************
'FUNCION HALLA LA CUOTA FIJA CON FECHA FIJA
'*****************************************************************
Function CFijaFechaFija(ByVal pTasa As Double, ByVal pPlazo As Integer, ByVal pnMonto As Double, _
    ByVal pnPer As Double, ByVal pdFecDes As Date, ByVal pnDiaFijo As Integer, _
    ByVal pnDiasGracia As Integer, ByVal pnProxMes As Integer, Optional pnDeltaMes As Integer = 1, _
    Optional ByVal pbMiViv As Boolean = False) As Double
    
Dim i As Integer
Dim SumTotal As Double
Dim nTasaDiaria As Double
Dim nValorActual As Double
Dim nDiasTrans As Integer
Dim dFecInicio, dFecVenc, dFecTemp As Date
Dim nDia, nMes, nAnio As Integer
Dim nProxMesTmp As Boolean

    nTasaDiaria = ((1 + (pTasa / 100)) ^ (1 / 30) - 1) * 100
    SumTotal = 0
    dFecInicio = pdFecDes + pnDiasGracia
    
    nDia = pnDiaFijo
    nMes = Month(dFecInicio)
    nAnio = Year(dFecInicio)
    
        If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
            If nDia = 31 Then
                nDia = 30
            End If
        End If
        If nMes = 2 Then
            If nDia >= 29 Then
                If nAnio Mod 4 = 0 Then 'si es biciesto
                    nDia = 29
                Else
                    nDia = 28
                End If
            End If
        End If
    
    dFecTemp = CDate(Format(Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Trim(Str(nAnio)), "dd/mm/yyyy"))
    If dFecTemp < dFecInicio Then
        nProxMesTmp = True
    Else
        nProxMesTmp = False
    End If
    
    If pnProxMes Or nProxMesTmp Or pnDeltaMes = 6 Then
        nMes = nMes + pnDeltaMes
    End If
    For i = 1 To pPlazo
        If nMes > 12 Then
            nMes = nMes - 12
            nAnio = nAnio + 1
        End If
        If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
            If nDia = 31 Then
                nDia = 30
            End If
        End If
        If nMes = 2 Then
            If nDia >= 29 Then
                If nAnio Mod 4 = 0 Then 'si es biciesto
                    nDia = 29
                Else
                    nDia = 28
                End If
            End If
        End If
        
        dFecVenc = CDate(Format(Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Format(nAnio, "0000"), "dd/mm/yyyy"))
        If i = 1 And pbMiViv Then
            dFecInicio = dFecInicio - pnDiasGracia
        End If
        nDiasTrans = dFecVenc - dFecInicio
        nValorActual = (1 / (1 + (nTasaDiaria / 100))) ^ (nDiasTrans)
        SumTotal = SumTotal + nValorActual
        nDia = pnDiaFijo
        nMes = nMes + pnDeltaMes
    Next i
    
    CFijaFechaFija = pnMonto / SumTotal
    CFijaFechaFija = CDbl(Format(CFijaFechaFija, "#0.00"))
    
End Function

'***********************************************
'*  FUNCION QUE HALLA EL INTERES DE UN PERIODO DE DIAS TRANACURRIDOS
'***********************************************
Public Function TasaIntPerDias(ByVal pnTasaInter As Double, ByVal pnDiasTrans As Integer) As Double
    TasaIntPerDias = ((1 + pnTasaInter / 100) ^ (pnDiasTrans / 30)) - 1
End Function

Public Function MontoIntPerDias(ByVal pnTasaInter As Double, ByVal pnDiasTrans As Integer, ByVal pnMonto As Double) As Double
    MontoIntPerDias = (((1 + pnTasaInter / 100) ^ (pnDiasTrans / 30)) - 1) * pnMonto
    MontoIntPerDias = CDbl(Format(MontoIntPerDias, "#0.00"))
End Function

'*********************************************************
'FUNCION QUE HALLA LA CUOTA FIJA APARTIR DE LA TASA DE INTERES Y EL PLAZO
'*********************************************************
Public Function CuotaFija(ByVal pnTasa As Double, ByVal pnCuotas As Integer, ByVal pnMonto As Double, ByVal pnPeriodo As Double) As Double
Dim Pot1 As Double
Dim nTasaTmp As Double
    nTasaTmp = TasaIntPerDias(pnTasa, pnPeriodo)
'Obtengo la cuota de pago
    Pot1 = (1 + nTasaTmp) ^ pnCuotas
    CuotaFija = ((Pot1 * nTasaTmp) / (Pot1 - 1)) * pnMonto
    CuotaFija = CDbl(Format(CuotaFija, "#0.00"))
End Function
 
'*********************************************************
'FUNCION QUE HALLA EL PLAZO APROXIMADO PARA UNA CUOTA FIJA
'*********************************************************
Public Function PlazoDeCuotaFija(ByVal pnTasa As Double, ByVal pnCuotasMax As Integer, _
    ByVal pnMonto As Double, ByVal pnPeriodo As Double, ByVal pnCuotaObjetivo As Double) As Integer
Dim nCuotaTemp As Double
Dim nCuotaResult As Double
Dim nCuotaTemp2  As Double
Dim i As Integer
    For i = pnCuotasMax To 1 Step -1
        nCuotaTemp = CuotaFija(pnTasa, i, pnMonto, pnPeriodo)
        If nCuotaTemp <= pnCuotaObjetivo Then
            nCuotaResult = nCuotaTemp
            Exit For
        End If
    Next i
    If i = 0 Then
        PlazoDeCuotaFija = pnCuotasMax
    Else
        PlazoDeCuotaFija = i
    End If
    
End Function

'*****************************************************************************************
'FUNCION QUE HALLA EL NUMERO DE CUOTAS APROXIMADO MANTENER LA FECHA DE VENC. DEL CREDITO
'*****************************************************************************************
Public Function CuotasAproximada(ByVal pdHoy As Date, ByVal pdFecVenc As Date, _
    ByVal pnPeriodo As Integer, Optional ByVal pnDia As Integer = 0, Optional ByVal pbProxMes As Boolean = True) As Integer
Dim i As Integer
Dim nValor As Integer
Dim dFecHoy As Date
Dim dFecTemp As Date
Dim nMes, nAnio, nDia As Integer
Dim bProxMes As Boolean

    bProxMes = pbProxMes

    'If Day(pdHoy) <= pnDia Then
    '    bProxMes = True
    'Else
    '    bProxMes = False
    'End If
    If CInt((pdFecVenc - pdHoy) / IIf(pnPeriodo = 0, 30, pnPeriodo)) = 1 Then
        nValor = 1
    Else
        nValor = CInt((pdFecVenc - pdHoy) / IIf(pnPeriodo = 0, 30, pnPeriodo)) + 1
    End If
    dFecHoy = pdHoy
    CuotasAproximada = nValor
    
    If pnPeriodo = 0 Then
        dFecTemp = dFecHoy
            nMes = Month(pdHoy)
            nAnio = Year(pdHoy)
            nDia = pnDia
            
            For i = 0 To nValor - 1
                nDia = pnDia
                If Not (i = 0 And pnDia > Day(pdHoy) And (Not bProxMes)) Then
                    nMes = nMes + 1
                    If nMes > 12 Then
                        nAnio = nAnio + 1
                        nMes = 1
                    End If
                Else
                    If nMes = 2 Then
                        If nDia >= 29 Then
                            If nAnio Mod 4 <> 0 Then
                                nMes = nMes + 1
                            End If
                        End If
                    Else
                        If nDia > 30 Then
                            If nMes = 4 Or nMes = 6 Or nMes = 9 Or 11 Then
                                nMes = nMes + 1
                            End If
                        End If
                    End If
                End If
                If nMes = 2 Then
                    If nDia > 28 Then
                        If nAnio Mod 4 = 0 Then
                            nDia = 29
                        Else
                            nDia = 28
                        End If
                    End If
                Else
                    If nDia > 30 Then
                        If nMes = 4 Or nMes = 6 Or nMes = 9 Or 11 Then
                            nDia = 30
                        End If
                    End If
                End If
                dFecTemp = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)))
                If dFecTemp > pdFecVenc Then
                    
                    Exit For
                End If
                CuotasAproximada = i + 1
        Next i
        
    Else
        For i = nValor To 1 Step -1
            If DateAdd("d", i * pnPeriodo, dFecHoy) <= pdFecVenc Then
                CuotasAproximada = i
                Exit For
            End If
        Next i
    End If
End Function

Public Function PlazoDeCuotaFijaFechaFija(ByVal pnCuotaObjetivo As Double, ByVal pTasa As Double, ByVal pPlazo As Integer, ByVal pnMonto As Double, _
    ByVal pnPer As Double, ByVal pdFecDes As Date, ByVal pnDiaFijo As Integer, _
    ByVal pnDiasGracia As Integer, ByVal pnProxMes As Integer, Optional pnDeltaMes As Integer = 1) As Integer
Dim nCuotaTemp As Double
Dim nCuotaResult As Double
Dim i As Integer
Dim nCuota As Double
    nCuota = 0
    For i = pPlazo To 1 Step -1
        nCuotaTemp = CFijaFechaFija(pTasa, i, pnMonto, pnPer, pdFecDes, pnDiaFijo, pnDiasGracia, pnProxMes, pnDeltaMes)
        If i = pPlazo Then
            nCuota = nCuotaTemp
        End If
        If nCuotaTemp > pnCuotaObjetivo Then
            nCuotaResult = nCuota
            Exit For
        End If
        nCuota = nCuotaTemp
    Next i
    PlazoDeCuotaFijaFechaFija = i + 1
    
End Function

Public Function PorcentajeCuotaCredConsumo(ByVal psConsCuotaCredConsumo As String) As Double
Dim oConecta As DConecta
Dim sSQL As String
Dim R As ADODB.Recordset

    On Error GoTo ErrorPorcentajeCuotaCredConsumo
    sSQL = "Select nParamValor from ColocParametro Where nParamVar = " & psConsCuotaCredConsumo
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    If Not R.BOF And Not R.EOF Then
        PorcentajeCuotaCredConsumo = R!nParamValor / 100
    Else
        PorcentajeCuotaCredConsumo = 0
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    R.Close
    Set R = Nothing
    Exit Function

ErrorPorcentajeCuotaCredConsumo:
    Set oConecta = Nothing
    Set R = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function


'CMACICA_CSTS - 26/11/2003 - ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ValidaCuotaParaTipoCredito(ByVal pnMonedaCod As Integer, ByVal pnMontoCuota As Double, ByVal pnMonedaFteCod As Integer, ByVal pnMontoIngCliFte As Double, ByVal pnValor As Double, ByVal pdFecha As Date) As String
Dim oGeneral As DGeneral
Dim nTipoCambioFijo As Double
Dim nMontoMN As Double
Dim nMontoME As Double
Dim nMontoIngCliFte As Double

    On Error GoTo ErrorValidaCuotaParaTipoCredito

    nMontoME = 0
    nMontoMN = 0
    ValidaCuotaParaTipoCredito = ""

    Set oGeneral = New DGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
    nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))
    Set oGeneral = Nothing

    'Moneda de la Fte de Ingreso
    If pnMonedaFteCod = gMonedaNacional Then 'Moneda Nacional '
       nMontoMN = pnMontoIngCliFte
       nMontoMN = CDbl(Format(nMontoMN, "#0.00"))

       nMontoME = nMontoME + (nMontoMN / nTipoCambioFijo)
    Else 'Moneda Extrangera '
       nMontoME = nMontoME + CDbl(Format(pnMontoIngCliFte, "#0.00"))
       nMontoMN = nMontoMN + (nMontoME * nTipoCambioFijo)
    End If

    'Moneda del Prestamo
    If pnMonedaCod = gMonedaNacional Then
        nMontoIngCliFte = 0
        If nMontoMN > 0 Then
            nMontoIngCliFte = nMontoIngCliFte + nMontoMN
        End If
    End If
    
    If pnMonedaCod = gMonedaExtranjera Then
        nMontoIngCliFte = 0
        If nMontoME > 0 Then
            nMontoIngCliFte = nMontoIngCliFte + nMontoME
        End If
    End If

    'Verifica la Condicion
    If pnMontoCuota > Format(pnValor * nMontoIngCliFte, "#0.00") Then
       ValidaCuotaParaTipoCredito = "Monto de Cuota No Puede Ser Mayor que  el " & Format(pnValor * 100, "#0.00") & "% del Sueldo Neto del Cliente para este Tipo de Crédito"
    End If

    Exit Function

ErrorValidaCuotaParaTipoCredito:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


 Public Function ValidaDatosGarantiaCredito(ByVal pnMonto As Double, ByVal pnMontoDisponible As Double) As String
    
    ValidaDatosGarantiaCredito = ""
    'Valida Existencia de Disponible
    If pnMontoDisponible = 0 Then
        ValidaDatosGarantiaCredito = "No Existe Disponible para Esta Garantia"
        Exit Function
    End If
    'Valida que Monto del Gravament no Exeda al Disponible
    If pnMonto > pnMontoDisponible Then
        ValidaDatosGarantiaCredito = "Monto de Garantia No debe Exeder al Disponible"
        Exit Function
    End If

 End Function

Public Function ValidaGrabarDesembolso(ByVal psCtaCod As String, Optional ByVal pnPrestamo As Double, Optional ByVal pbDesembCC As Boolean, Optional ByVal psCtaAbo As String, Optional ByVal pbCtaAhoNueva As Boolean = False) As String
Dim oBase As DCredActualizaBD
Dim oCredito As DCredito
Dim R As ADODB.Recordset
Dim sCodLinea As String

    On Error GoTo ErrorValidaGrabarDesembolso
    ValidaGrabarDesembolso = ""
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaColocaciones(psCtaCod)
    Set oCredito = Nothing
    sCodLinea = R!cLineaCred
    R.Close
    Set R = Nothing
    
    Set oBase = New DCredActualizaBD
    Set R = oBase.RecuperaLineasCreditoSaldo(sCodLinea)
    Set oBase = Nothing
    
    If R.BOF And R.EOF Then
        ValidaGrabarDesembolso = "Falta definir los saldos de la Linea de Credito, Consulte con Caja General"
        R.Close
        Set R = Nothing
        Exit Function
    End If
    R.Close
    Set R = Nothing
    
    If pnPrestamo <= 0 Then
        ValidaGrabarDesembolso = "Monto de Prestamo debe ser mayor que cero"
        Exit Function
    End If
    
    If pbDesembCC And Not pbCtaAhoNueva Then
        If Trim(psCtaAbo) = "" Then
            ValidaGrabarDesembolso = "Falta seleccionar la Cuenta de Ahorro a la que se va a depositar el Desembolso"
            Exit Function
        End If
    End If
    Exit Function

ErrorValidaGrabarDesembolso:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function
Private Function DameTipoCuota(ByVal pnTipoCuota As Integer) As Integer
        If pnTipoCuota = gColocCalendCodFFCC Or pnTipoCuota = gColocCalendCodFFCCPG Or pnTipoCuota = gColocCalendCodPFCC Or pnTipoCuota = gColocCalendCodPFCCPG Then
            DameTipoCuota = 2
        End If
        If pnTipoCuota = gColocCalendCodFFCF Or pnTipoCuota = gColocCalendCodFFCFPG Or pnTipoCuota = gColocCalendCodPFCF Or pnTipoCuota = gColocCalendCodPFCFPG Then
            DameTipoCuota = 1
        End If
        If pnTipoCuota = gColocCalendCodFFCD Or pnTipoCuota = gColocCalendCodFFCDPG Or pnTipoCuota = gColocCalendCodPFCD Or pnTipoCuota = gColocCalendCodPFCDPG Then
            DameTipoCuota = 3
        End If
End Function

Private Function DameTipoPeriodo(ByVal pnTipoPeriodo As Integer) As Integer
    If pnTipoPeriodo = gColocCalendCodFFCC Or pnTipoPeriodo = gColocCalendCodFFCCPG Or pnTipoPeriodo = gColocCalendCodFFCD Or pnTipoPeriodo = gColocCalendCodFFCCPG _
          Or pnTipoPeriodo = gColocCalendCodFFCCPG Or pnTipoPeriodo = gColocCalendCodFFCD Or pnTipoPeriodo = gColocCalendCodFFCDPG Or pnTipoPeriodo = gColocCalendCodFFCF Or pnTipoPeriodo = gColocCalendCodFFCFPG Then
            DameTipoPeriodo = 2
        End If
        If pnTipoPeriodo = gColocCalendCodPFCC Or pnTipoPeriodo = gColocCalendCodPFCCPG Or pnTipoPeriodo = gColocCalendCodPFCD Or pnTipoPeriodo = gColocCalendCodPFCCPG _
          Or pnTipoPeriodo = gColocCalendCodPFCCPG Or pnTipoPeriodo = gColocCalendCodPFCD Or pnTipoPeriodo = gColocCalendCodPFCDPG Or pnTipoPeriodo = gColocCalendCodPFCF Or pnTipoPeriodo = gColocCalendCodPFCFPG Then
            DameTipoPeriodo = 1
        End If
End Function

Public Function DesembolsarCredito(ByRef pnMovNro As Long, ByRef pMatDatosAho As Variant, ByVal psCtaCod As String, ByVal pnPrestamo As Double, ByVal pnMontoDesembolso As Double, _
        ByVal pMatGastos As Variant, ByVal MatCargoAutom As Variant, ByVal pdHoy As Date, ByVal psCodAge As String, psCodUser As String, _
        ByVal pbDesembCC As Boolean, Optional ByVal pbCtaAhoNueva As Boolean = False, _
        Optional ByVal MatCredCanc As Variant = Nothing, Optional ByVal pnContMatCredCanc As Integer = -1, Optional ByVal psCtaAbo As String = "", _
        Optional ByVal pRSRela As ADODB.Recordset = Nothing, Optional ByVal pnTasa As Double = 0#, Optional ByVal pnPersoneria As Integer = 0, _
        Optional ByVal pnTipoCuenta As Integer = 0, Optional ByVal pnTipoTasa As Integer = 0, Optional ByVal pbDocumento As Boolean = False, _
        Optional ByVal psNroDoc As String = "", Optional ByVal psCodIF As String = "", _
        Optional ByVal pnMontoGastos As Double = 0#, Optional ByVal pnMontoCancelaciones As Double = 0#, Optional psPersLavDinero As String = "", Optional ByVal pnITF As Double = 0#) As String

Dim R As ADODB.Recordset
Dim oDCredito As DCredito
'Dim oFunciones As NContFunciones
Dim oBase As DCredActualizaBD
Dim oLinea As DLineaCredito
Dim oCalend As NCalendario
Dim oDCalend As Dcalendario
Dim oGastos As NGasto
Dim sLineaCod As String
Dim nMontoReservado As Double
Dim nMontoColocado As Double
Dim MatCalendPagos As Variant
Dim MatCalendPagos_2 As Variant
Dim MatGracia As Variant
Dim nPeriodoGracia As Integer
Dim nTasaInteres As Double
Dim nNroCalen As Integer
Dim nNroProxDesemb As Integer
Dim nCuotaCom As Integer
Dim nMiViv As Integer
Dim i As Integer
Dim MatGastos As Variant
Dim MatDesemb() As String
Dim ContDesemb As Integer
'Dim MatGastosDesemb(1000, 3) As String
'Dim ContGastosDesemb As Integer
Dim nNumGastos As Integer
Dim nNumTransac As Long
Dim sMovNro As String
Dim nMovNro As Long
Dim nMovNroAnt As Long
Dim nConsCred As String
Dim sMetLiquid As String
Dim bTransac As Boolean
Dim MatDesParcial As Variant

'Para MiVivienda
Dim oParam As DParametro
Dim nTramoNoConsMonto As Double
Dim nTramoConsMonto As Double
Dim nTramoNoConsPorcen As Double
Dim nTramoConsTasa As Double
Dim nMontoGarantia As Double
Dim oGar As DGarantia
Dim nCuotaTemp As String
Dim dFecTran As Date

'Para Cancelacion de Creditos
Dim MatCalendTemp As Variant
Dim MatCalendDistrTemp As Variant
Dim sMetLiqTmp As String
Dim RTmp As ADODB.Recordset
Dim MatRSCancel() As ADODB.Recordset
Dim MatRSCancelCalPend() As Variant
Dim RCalendActual As ADODB.Recordset
Dim RCalendDesembActual As ADODB.Recordset
Dim RColocCalendDesembActual As ADODB.Recordset

'Para Apertura de Cuenta de Ahorros
Dim sCtaApe As String
Dim nMivivMontoCondonadoCapital As Double
Dim nMivivMontoCondonadoInteres As Double
Dim sTipoGasto As String
Dim sConsAhoRetGas As String
Dim sConsAhoRetCan As String

        On Error GoTo ErrorDesembolsarCredito
        
        'inicializa Datos de Ahorros
        pMatDatosAho(0) = "" 'Cuenta de Ahorros
        pMatDatosAho(1) = "0.00" 'Monto de Apertura
        pMatDatosAho(2) = "0.00" 'Interes Ganado de Abono
        pMatDatosAho(3) = "0.00" 'Interes Ganado de Retiro Gastos
        pMatDatosAho(4) = "0.00" 'Interes Ganado de Retiro Cancelaciones
        pMatDatosAho(5) = "0.00" 'Monto de Abono
        pMatDatosAho(6) = "0.00" 'Monto de Retiro de Gastos
        pMatDatosAho(7) = "0.00" 'Monto de Retiro de Cancelaciones
        pMatDatosAho(8) = "0.00" 'Saldo Disponible Abono
        pMatDatosAho(9) = "0.00" 'Saldo Contable Abono
        pMatDatosAho(10) = "0.00" 'Saldo Disponible Retiro de Gastos
        pMatDatosAho(11) = "0.00" 'Saldo Contable Retiro de Gastos
        pMatDatosAho(12) = "0.00" 'Saldo Disponible Retiro de Cancelaciones
        pMatDatosAho(13) = "0.00" 'Saldo Contable Retiro de Cancelaciones
        
        bTransac = False
        'Valida datos antes de Desembolsar
        DesembolsarCredito = ValidaGrabarDesembolso(psCtaCod, pnPrestamo, pbDesembCC, psCtaAbo, pbCtaAhoNueva)
        If DesembolsarCredito <> "" Then
            Exit Function
        End If
        
        If pbDesembCC Then
            If pbCtaAhoNueva Then
                'If Mid(psCtaAbo, 4, 2) <> psCodAge Then 'Cuenta de Otra Age
                '    nConsCred = gCredDesembCtaNuevaDOA
                'Else
                    nConsCred = gCredDesembCtaNueva
                    sConsAhoRetGas = gCredDesembCtaRetiroGastos
                    sConsAhoRetCan = gCredDesembCtaRetiroCancelacion
                'End If
            Else
                If Mid(psCtaAbo, 4, 2) <> psCodAge Then 'Cuenta de Otra Age
                    nConsCred = gCredDesembCtaExistDOA
                    sConsAhoRetGas = gCredDesembCtaRetiroGastosDOA
                    sConsAhoRetCan = gCredDesembCtaRetiroCancelacionDOA
                Else
                    nConsCred = gCredDesembCtaExist
                    sConsAhoRetGas = gCredDesembCtaRetiroGastos
                    sConsAhoRetCan = gCredDesembCtaRetiroCancelacion
                End If
            End If
        Else
            nConsCred = gCredDesembEfec
        End If
        
        
        Set oDCredito = New DCredito
        Set R = oDCredito.RecuperaColocaciones(psCtaCod)
        sLineaCod = R!cLineaCred
        R.Close
        Set R = Nothing
        
        
        Set R = oDCredito.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
        sTipoGasto = R!cTipoGasto
        R.Close
        Set R = Nothing
        
        Set R = oDCredito.RecuperaProducto(psCtaCod)
        nTasaInteres = Format(R!nTasaInteres, "#0.0000")
        nNumTransac = IIf(IsNull(R!nTransacc), 0, R!nTransacc)
        R.Close
        Set R = Nothing
        Set R = oDCredito.RecuperaColocacCred(psCtaCod)
        nNroCalen = R!nNroCalen
        nNroProxDesemb = R!nNroProxDesemb
        sMetLiquid = R!cMetLiquidacion
        nCuotaCom = R!bCuotaCom
        nMiViv = R!bMiVivienda
        If nMiViv = 1 Then
            nNroCalen = R!nNroCalPar
        End If
        R.Close
        Set R = Nothing
        Set oDCredito = Nothing
        
        
        Set oBase = Nothing
        Set oBase = New DCredActualizaBD
        sMovNro = oBase.GeneraMovNro(pdHoy, psCodAge, psCodUser)
                
        'Recupera Calendario de Desembolsos
        ContDesemb = 0
        Set oDCalend = New Dcalendario
        Set R = oDCalend.RecuperaCalendarioDesemb(psCtaCod)
        Do While Not R.EOF
            ContDesemb = ContDesemb + 1
            R.MoveNext
        Loop
        If ContDesemb > 0 Then
            ReDim MatDesemb(ContDesemb, 3)
            ContDesemb = 0
            R.MoveFirst
            Do While Not R.EOF
                ContDesemb = ContDesemb + 1
                If R!nColocCalendEstado = 1 Then
                    MatDesemb(ContDesemb - 1, 0) = Format(R!dPago, "dd/mm/yyyy")
                Else
                    MatDesemb(ContDesemb - 1, 0) = Format(R!DVENC, "dd/mm/yyyy")
                End If
                MatDesemb(ContDesemb - 1, 1) = Format(R!nCapital, "#0.00")
                MatDesemb(ContDesemb - 1, 2) = Format(R!DVENC, "dd/mm/yyyy")
                R.MoveNext
            Loop
        End If
        R.Close
        MatDesParcial = MatDesemb
        
        'Actualiza MatDesembolsos
        If ContDesemb > 1 Then
            For i = 0 To UBound(MatDesParcial) - 1
                If i = nNroProxDesemb - 1 Then
                    MatDesParcial(i, 0) = Format(pdHoy, "dd/mm/yyyy")
                Else
                    If i > nNroProxDesemb - 1 Then
                        MatDesParcial(i, 1) = "0.00"
                    End If
                End If
            Next i
            
        End If
        
        Set R = Nothing
        
        Set RCalendActual = oDCalend.RecuperaColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota)
        Set RCalendDesembActual = oDCalend.RecuperaColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplDesembolso)
        Set RColocCalendDesembActual = oDCalend.RecuperaColocCalendario(psCtaCod, nNroCalen, gColocCalendAplDesembolso)
        Set oDCalend = Nothing
        
        
        '********************************************************************
        'Genera Nuevo Calendario
        '********************************************************************
        Set oDCredito = New DCredito
        Set R = oDCredito.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
        Set oDCredito = Nothing
        If R!nColocCalendCod <> gColocCalendCodCL Then
            Set oCalend = New NCalendario
            nPeriodoGracia = IIf(IsNull(R!nPeriodoGracia), 0, R!nPeriodoGracia)
            If nPeriodoGracia > 0 Then
                MatGracia = oCalend.GeneraGracia(R!nTipoGracia, CDbl(Format(TasaIntPerDias(nTasaInteres, nPeriodoGracia) * pnPrestamo, "#0.00")), R!nCuotas)
            End If
            
            If nMiViv = 1 Then

                'Porcentaje Real sin dividir enter 100
                Set oParam = New DParametro
                nTramoNoConsPorcen = oParam.RecuperaValorParametro(gColocMiVivTramo)
                Set oParam = Nothing

                nTramoNoConsMonto = Format((nTramoNoConsPorcen / 100) * pnPrestamo, "#0.00")
                nTramoConsMonto = Format(pnPrestamo - nTramoNoConsMonto, "#0.00")
                MatCalendPagos = oCalend.GeneraCalendario(nTramoConsMonto, nTasaInteres, R!nCuotas, _
                            R!nPlazo, pdHoy, DameTipoCuota(R!nColocCalendCod), DameTipoPeriodo(R!nColocCalendCod), _
                            IIf(IsNull(R!nTipoGracia), 0, R!nTipoGracia), nPeriodoGracia, IIf(IsNull(R!nPeriodoFechaFija), 0, R!nPeriodoFechaFija), _
                            IIf(R!nProxMes = 0, False, True), MatGracia, True)
                MatCalendPagos_2 = oCalend.GeneraCalendario(nTramoNoConsMonto, nTasaInteres, R!nCuotas / 6, _
                            180, pdHoy, Fija, FechaFija, _
                            Exonerada, 0, IIf(IsNull(R!nPeriodoFechaFija), 0, R!nPeriodoFechaFija), IIf(R!nProxMes = 0, False, True), MatGracia, True, , , , 6)
                MatCalendPagos_2 = UnirMatricesMiVivienda(MatCalendPagos, MatCalendPagos_2, pnPrestamo)
                nCuotaTemp = CDbl(MatCalendPagos_2(7, 2))
                MatCalendPagos_2 = DiferencialMatricesMiVivienda(MatCalendPagos, MatCalendPagos_2)
            Else
                MatCalendPagos = oCalend.GeneraCalendario(pnPrestamo, nTasaInteres, R!nCuotas, R!nPlazo, _
                    pdHoy, DameTipoCuota(R!nColocCalendCod), DameTipoPeriodo(R!nColocCalendCod), _
                    IIf(IsNull(R!nTipoGracia), 0, R!nTipoGracia), nPeriodoGracia, IIf(IsNull(R!nPeriodoFechaFija), 0, R!nPeriodoFechaFija), _
                    IIf(R!nProxMes = 0, False, True), MatGracia, True, IIf(nCuotaCom = 1, True, False), IIf(ContDesemb > 1, True, False), MatDesParcial)
                nCuotaTemp = CDbl(MatCalendPagos(0, 2))
            End If
            
            Set oCalend = Nothing
        
        R.Close
        Set R = Nothing
        '********************************************************************
        'GENERACION DE GASTOS
        '********************************************************************
        Set oGar = New DGarantia
        nMontoGarantia = oGar.RecuperaMontoGarantiaCredito(psCtaCod)
        Set oGar = Nothing
        Set oGastos = New NGasto
        MatGastos = oGastos.GeneraCalendarioGastos(MatCalendPagos, MatDesemb, nNumGastos, pdHoy, psCtaCod, 1, "DE", sTipoGasto, CDbl(MatCalendPagos(0, 2)), pnPrestamo)
        Set oGastos = Nothing
        
        
        '********************************************************************
        'Para Cancelacion de Creditos
        '********************************************************************
        If nNroProxDesemb = 1 Then
            ReDim MatRSCancel(pnContMatCredCanc)
            ReDim MatRSCancelCalPend(pnContMatCredCanc)
            Set oDCredito = New DCredito
            Set RTmp = New ADODB.Recordset
            For i = 0 To pnContMatCredCanc - 1
                Set MatRSCancel(i) = oDCredito.RecuperaDatosCreditoVigente(MatCredCanc(i, 0))
                MatRSCancelCalPend(i) = RecuperaMatrizCalendarioPendiente(MatCredCanc(i, 0))
            Next i
            Set oDCredito = Nothing
            
        End If
        '********************************************************************
        'INICIO DE TRANSACCION
        '********************************************************************
        Call oBase.dBeginTrans
        bTransac = True
                                       
        '********************************************************************
        'Inserta Movimiento
        Call oBase.dInsertMov(sMovNro, nConsCred, "", gMovEstContabMovContable, gMovFlagVigente, False)
        nMovNro = oBase.dGetnMovNro(sMovNro)
        pnMovNro = nMovNro
        
        '**********************************************************
        'Apertura de Cuenta de Ahorro
        '**********************************************************
        If pbDesembCC And pbCtaAhoNueva Then
                sCtaApe = oBase.CapAperturaCuenta(gCapAhorros, CInt(Mid(psCtaCod, 9, 1)), _
                        pRSRela, sMovNro, psCodAge, nConsCred, pnTasa, pnPrestamo, pdHoy, 1, _
                        pnPersoneria, "", pnTipoCuenta, pnTipoTasa, pbDocumento, psNroDoc, psCodIF)
                psCtaAbo = sCtaApe
                pMatDatosAho(1) = Format(pnPrestamo, "#0.00")
            End If
        pMatDatosAho(0) = psCtaAbo
        '**********************************************************
        
        '***********************************************************
        'Proceso de Cancelacion de Creditos
        '**********************************************************
        If nNroProxDesemb = 1 Then
            Set oDCredito = New DCredito
            Set RTmp = New ADODB.Recordset
            For i = 0 To pnContMatCredCanc - 1
                Set RTmp = MatRSCancel(i)
                sMetLiqTmp = RTmp!cMetLiquidacion
                RTmp.Close
                'MatCalendTemp = RecuperaMatrizCalendarioPendiente(MatCredCanc(I, 0))
                MatCalendTemp = MatRSCancelCalPend(i)
                MatCalendDistrTemp = MatrizDistribuirCancelacion(MatCredCanc(i, 0), MatCalendTemp, fgITFCalculaImpuestoIncluido(CDbl(MatCredCanc(i, 1))), sMetLiqTmp, pdHoy)
                Call AmortizarCredito(MatCredCanc(i, 0), MatCalendTemp, MatCalendDistrTemp, fgITFCalculaImpuestoIncluido(CDbl(MatCredCanc(i, 1))), pdHoy, sMetLiqTmp, IIf(pbDesembCC, gColocTipoPagoCargoCta, gColocTipoPagoEfectivo), psCodAge, psCodUser, , oBase, nMovNro, , , , , , sMovNro, , , , , , , , CDbl(MatCredCanc(i, 1)) - fgITFCalculaImpuestoIncluido(CDbl(MatCredCanc(i, 1))))
            Next i
            Set oDCredito = Nothing
        End If
        '**********************************************************
        'Final de Proceso de Cancelacion
        '**********************************************************
        
        'Actualiza la Saldos de Linea
        Set oLinea = New DLineaCredito
        'Set R = oBase.RecuperaLineasCreditoSaldo(sLineaCod, True)
        'Set oLinea = Nothing
        'nMontoReservado = CDbl(Format(R!nMontoReservado, "#0.00"))
        'nMontoColocado = CDbl(Format(R!nMontoColocado, "#0.00"))
        'R.Close
        'Set R = Nothing
        'nMontoReservado = nMontoReservado - pnPrestamo
        'nMontoColocado = nMontoColocado + pnPrestamo
        'Call oBase.dUpdateLineaCreditoSaldo(sLineaCod, , , nMontoColocado, , nMontoReservado, False)
            
            'Nuevo Calendario de Desembolsos
            For i = 0 To ContDesemb - 1
                Call oBase.dInsertColocCalendario(psCtaCod, nNroCalen + 1, gColocCalendAplDesembolso, i + 1, CDate(MatDesemb(i, 0)), gColocCalendEstadoPendiente, "Calendario de Desembolsos -  Desembolso", gColocCalendConceptoProcDesembolsado, False)
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplDesembolso, i + 1, gColocConceptoCodCapital, CDbl(MatDesemb(i, 1)), 0, "", False)
            Next i
            
            'Nuevos Gastos de Desembolsos
            For i = 0 To nNumGastos - 1
                If CInt(Trim(Right(MatGastos(i, 0), 10))) = gColocCalendAplDesembolso Then
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplDesembolso, CInt(Trim(MatGastos(i, 1))), CLng(Trim(Right(MatGastos(i, 2), 20))), CDbl(MatGastos(i, 3)), 0, "", False)
                End If
            Next i
                                    
                                    
            'Para Desembolsos Parciales Actualizo del Calendario Anterior
            'ColoCalendario
            If nNroProxDesemb <> 1 Then
                Do While Not RColocCalendDesembActual.EOF
                    Call oBase.dUpdateColocCalendario(psCtaCod, nNroCalen + 1, RColocCalendDesembActual!nCuota, gColocCalendAplDesembolso, , RColocCalendDesembActual!nColocCalendEstado, , , False, , IIf(IsNull(RColocCalendDesembActual!dPago), CDate("01/01/1900"), RColocCalendDesembActual!dPago))
                    RColocCalendDesembActual.MoveNext
                Loop
            End If
            RColocCalendDesembActual.Close
            'ColoCalendDet
            If nNroProxDesemb <> 1 Then
                Do While Not RCalendDesembActual.EOF
                    Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplDesembolso, RCalendDesembActual!nCuota, RCalendDesembActual!nPrdConceptoCod, , RCalendDesembActual!nMontoPagado, "", False)
                    RCalendDesembActual.MoveNext
                Loop
            End If
            RCalendDesembActual.Close
            
            
            'Nuevo Calendario de Pagos
            For i = 0 To UBound(MatCalendPagos) - 1
                
                    Call oBase.dInsertColocCalendario(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, i + 1, CDate(MatCalendPagos(i, 0)), gColocCalendEstadoPendiente, "Calendario de Pagos - Desembolso", gColocCalendConceptoProcDesembolsado, False)
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, i + 1, gColocConceptoCodCapital, CDbl(MatCalendPagos(i, 3)), 0, "", False)
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalendPagos(i, 4)), 0, "", False)
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", False)
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresCompVencido, 0, 0, "", False)
                    If CDbl(MatCalendPagos(i, 5)) > 0 Then
                        Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresGracia, CDbl(MatCalendPagos(i, 5)), 0, "", False)
                    End If
                
            Next i
            
            'Nuevos Gastos de Cuota
            For i = 0 To nNumGastos - 1
                If CInt(Trim(Right(MatGastos(i, 0), 10))) = gColocCalendAplCuota Then
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, CInt(Trim(Right(MatGastos(i, 0), 3))), CInt(Trim(MatGastos(i, 1))), CLng(Trim(Right(MatGastos(i, 2), 20))), CDbl(MatGastos(i, 3)), 0, "", False)
                End If
            Next i
            
            If nNroProxDesemb <> 1 Then
                Do While Not RCalendActual.EOF
                    Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, RCalendActual!nCuota, RCalendActual!nPrdConceptoCod, , RCalendActual!nMontoPagado, "", False)
                    RCalendActual.MoveNext
                Loop
            End If
            RCalendActual.Close
            
            If nMiViv = 1 Then
                
                For i = 0 To UBound(MatCalendPagos_2) - 1
                    Call oBase.dInsertColocCalendario(psCtaCod, nNroCalen + 2, gColocCalendAplCuota, i + 1, CDate(MatCalendPagos_2(i, 0)), gColocCalendEstadoPendiente, "Calendario de Pagos - Desembolso", gColocCalendConceptoProcDesembolsado, False)
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 2, gColocCalendAplCuota, i + 1, gColocConceptoCodCapital, CDbl(MatCalendPagos_2(i, 3)), 0, "", False)
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 2, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalendPagos_2(i, 4)), 0, "", False)
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 2, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", False)
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 2, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresCompVencido, 0, 0, "", False)
                    If CDbl(MatCalendPagos_2(i, 5)) > 0 Then
                        Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 2, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresGracia, CDbl(MatCalendPagos_2(i, 5)), 0, "", False)
                    End If
                Next i
                'Llena Tabla de Calificacion (ColocCalifMiViv) Buen Pagador o Mal Pagador
                For i = 0 To UBound(MatCalendPagos) - 1
                    Call oBase.dInsertColocCalifMiViv(psCtaCod, i + 1, nNroCalen + 1, i + 1, CDate(MatCalendPagos(i, 0)), gColocCalendEstadoPendiente)
                Next i
                
            End If
            
        Else ' Si es Cuota Libre
            MatCalendPagos = RecuperaMatrizCalendarioPendiente(psCtaCod)
            Call oBase.dBeginTrans
            bTransac = True
            'Inserta Movimiento
            Call oBase.dInsertMov(sMovNro, nConsCred, "", gMovEstContabMovContable, gMovFlagVigente, False)
            nMovNro = oBase.dGetnMovNro(sMovNro)
            pnMovNro = nMovNro
            nNroCalen = nNroCalen - 1
        End If
        
        Set oDCredito = New DCredito
        Set R = oDCredito.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
        Set oDCredito = Nothing
        
        'Actualiza Estado
        Call oBase.dInsertColocacEstado(psCtaCod, pdHoy, gColocEstVigNorm, R!nCuotas, pnPrestamo, "Desembolso de Credito", R!nColocCalendCod, IIf(IsNull(R!nPeriodoFechaFija), 0, R!nPeriodoFechaFija), nPeriodoGracia, R!nPlazo, R!nTipoGracia, R!nTipoDesembolso, R!nProxMes, R!nCalendDinamico, False)
        'Actualiza Colocaciones
        If nNroProxDesemb = 1 Then
            Call oBase.dUpdateColocaciones(psCtaCod, CDate(MatCalendPagos(UBound(MatCalendPagos) - 1, 0)) - CDate(MatDesemb(0, 0)) _
                , CDate(MatCalendPagos(UBound(MatCalendPagos) - 1, 0)), pnPrestamo, , , pdHoy, False)
        Else
            Call oBase.dUpdateColocaciones(psCtaCod, CDate(MatCalendPagos(UBound(MatCalendPagos) - 1, 0)) - CDate(MatDesemb(0, 0)) _
                , CDate(MatCalendPagos(UBound(MatCalendPagos) - 1, 0)), pnPrestamo, , , , False, True)
        End If
        'Actualiza ColocacCred
        Call oBase.dUpdateColocacCred(psCtaCod, DateDiff("d", CDate(MatCalendPagos(0, 0)), pdHoy), , , , IIf(UBound(MatCargoAutom) > 0, 1, 0), , 0, 1, , , , , , nNroCalen + 1, nNroProxDesemb + 1, , , False, nCuotaCom, nMiViv, nNroCalen + 2, IIf(nMiViv = 1, 1, 0))
        
        'Actualiza Producto
        If nNroProxDesemb = 1 Then
            Call oBase.dUpdateProducto(psCtaCod, , pnPrestamo, gColocEstVigNorm, pdHoy, nNumTransac + 1, False)
        Else
            Call oBase.dUpdateProducto(psCtaCod, , pnPrestamo, , pdHoy, nNumTransac + 1, False, , True)
        End If
        R.Close
        Set R = Nothing
        
        'Inserta Movimientos en MovCol
        'Call oBase.dInsertMovCol(nMovNro, nConsCred, psCtaCod, nNroCalen + 1, CDbl(Format(pnPrestamo - (pnMontoGastos + pnMontoCancelaciones), "#0.00")), 0, sMetLiquid, CDate(MatCalendPagos(UBound(MatCalendPagos) - 1, 0)) - CDate(MatCalendPagos(0, 0)), False, gColocEstAprob)
        Call oBase.dInsertMovCol(nMovNro, nConsCred, psCtaCod, nNroCalen + 1, CDbl(Format(pnPrestamo - (pnMontoGastos + pnITF), "#0.00")), 0, sMetLiquid, CDate(MatCalendPagos(UBound(MatCalendPagos) - 1, 0)) - CDate(MatCalendPagos(0, 0)), False, gColocEstAprob)
        
        Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, nNroCalen + 1, gColocConceptoCodCapital, nNroProxDesemb, (pnPrestamo - pnITF), False)
        
        '***************    ITF *********************************
        Call oBase.dInsertMovCol(nMovNro, gCredITFDesemb, psCtaCod, 0, pnITF, 0, sMetLiquid, 0, False, gColocEstAprob)
        Call oBase.dInsertMovColDet(nMovNro, gCredITFDesemb, psCtaCod, 0, gConcITFCliente, nNroProxDesemb, pnITF, False)
        
        'Ingresa el Monto de Todos los Gastos del desembolso
        For i = 0 To UBound(pMatGastos) - 1
            Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, nNroCalen + 1, CLng(pMatGastos(i, 2)), CInt(pMatGastos(i, 0)), Format(-1 * pMatGastos(i, 1), "#0.00"), False)
        Next i
                
        '********************************************************************
        'Actualiza Calendario de Desembolsos
        Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplDesembolso, nNroProxDesemb, gColocConceptoCodCapital, , pnPrestamo, , False)
        
        '********************************************************************
        'Actualiza Gastos de Desembolso
        If nNumGastos > 0 Then
            For i = 0 To UBound(pMatGastos) - 1
                'If CInt(Trim(Right(pMatGastos(i, 0), 10))) = gColocCalendAplDesembolso Then
                    Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplDesembolso, nNroProxDesemb, CLng(Trim(Right(pMatGastos(i, 2), 20))), , CDbl(pMatGastos(i, 1)), , False)
                'End If
            Next i
        End If
        'Cambia a estado Cancelado el Desembolso
        Call oBase.dUpdateColocCalendario(psCtaCod, nNroCalen + 1, nNroProxDesemb, gColocCalendAplDesembolso, , gColocCalendEstadoPagado, , , False, , pdHoy)
        
        '*****************************************************************
        'Actualiza Cuentas de Ahorro en caso de Abono Cta
        '*****************************************************************
        If pbDesembCC And Not pbCtaAhoNueva Then
            oBase.CapAbonoCuentaAho pMatDatosAho, psCtaAbo, pnPrestamo, nConsCred, sMovNro, "", , , , , , , pdHoy, ""
            pMatDatosAho(5) = pnPrestamo
        End If
        
        'Actualiza CargoAutomatico
        For i = 0 To UBound(MatCargoAutom) - 1
            Call oBase.dInsertColocCargoAutoma(psCtaCod, i + 1, MatCargoAutom(i), False)
        Next i
        
        ''''''''''''''''''''
        ' Lavado de Dinero
        If psPersLavDinero <> "" Then
            Call oBase.dInsertaMovLavDinero(nMovNro, psPersLavDinero)
        End If
        '
        ''''''''''''''''''''
        
        
        Call oBase.dCommitTrans
        
        'Retiro de Gastos y Cancelaciones
        
       
            If pbDesembCC Then
                nMovNroAnt = nMovNro
                If pnMontoGastos > 0 Then
                    Call oBase.dBeginTrans
                    sMovNro = oBase.GeneraMovNro(pdHoy, psCodAge, psCodUser)
                    Call oBase.dInsertMov(sMovNro, nConsCred, "", gMovEstContabMovContable, gMovFlagVigente, False)
                    nMovNro = oBase.dGetnMovNro(sMovNro)
                    Call oBase.dInsertMovRef(nMovNro, nMovNroAnt)
                    oBase.CapCargoCuentaAho pMatDatosAho, psCtaAbo, pnMontoGastos, sConsAhoRetGas, sMovNro, "", TpoDocBolDep, "", "", False
                    pMatDatosAho(6) = Format(pnMontoGastos, "#0.00")
                    Call oBase.dCommitTrans
                End If
                If pnMontoCancelaciones > 0 Then
                        Sleep 1000
                        Call oBase.dBeginTrans
                        sMovNro = oBase.GeneraMovNro(pdHoy, psCodAge, psCodUser)
                        Call oBase.dInsertMov(sMovNro, nConsCred, "", gMovEstContabMovContable, gMovFlagVigente, False)
                        nMovNro = oBase.dGetnMovNro(sMovNro)
                        Call oBase.dInsertMovRef(nMovNro, nMovNroAnt)
                        oBase.CapCargoCuentaAho pMatDatosAho, psCtaAbo, pnMontoCancelaciones, sConsAhoRetCan, sMovNro, "", TpoDocBolDep, "", "", False, True, , , , , , , True
                        pMatDatosAho(7) = Format(pnMontoCancelaciones, "#0.00")
                        Call oBase.dCommitTrans
                End If
            End If
           
        '**************************************************************************
        'Para Creditos MiVivienda Cancelar Primeras 6 Cuotas del Tramo Concesional
        '**************************************************************************
        If nMiViv Then
            Call oBase.dBeginTrans
            nMivivMontoCondonadoCapital = 0
            nMivivMontoCondonadoInteres = 0
            For i = 0 To 5
                nMivivMontoCondonadoCapital = nMivivMontoCondonadoCapital + CDbl(MatCalendPagos_2(i, 3))
                nMivivMontoCondonadoInteres = nMivivMontoCondonadoInteres + CDbl(MatCalendPagos_2(i, 4)) + CDbl(MatCalendPagos_2(i, 5))
                Call oBase.dUpdateColocCalendario(psCtaCod, nNroCalen + 2, i + 1, gColocCalendAplCuota, , gColocCalendEstadoPagado, "Calendario de Pagos - Desembolso", gColocCalendConceptoProcDesembolsado, False)
                Call oBase.dUpdateColocCalifMiViv(psCtaCod, nNroCalen + 1, i + 1, , , , pdHoy)
            Next i
            If (nMivivMontoCondonadoCapital + nMivivMontoCondonadoInteres) > 0 Then
                Call oBase.dInsertMovCol(nMovNro, gCredMiVivCondonacion, psCtaCod, nNroCalen + 2, CDbl(Format(nMivivMontoCondonadoCapital + nMivivMontoCondonadoInteres, "#0.00")), 0, sMetLiquid, 0, 0, 0, False)
                Call oBase.dInsertMovColDet(nMovNro, gCredMiVivCondonacion, psCtaCod, nNroCalen + 2, gColocConceptoCodCapital, 0, nMivivMontoCondonadoCapital, False)
                Call oBase.dInsertMovColDet(nMovNro, gCredMiVivCondonacion, psCtaCod, nNroCalen + 2, gColocConceptoCodInteresCompensatorio, 0, nMivivMontoCondonadoInteres, False)
            End If
            Call oBase.dCommitTrans
        End If
        
        If nNroProxDesemb = 1 Then
            'Asiento de Garantias
            Call AsientoGarantia(psCtaCod, psCodAge, psCodUser, pdHoy)
        End If
        
        Set oBase = Nothing
        

        
        Exit Function

ErrorDesembolsarCredito:
    If bTransac Then
        Call oBase.dRollbackTrans
        Set oBase = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
        
End Function

Public Function ValidaCargaDatosDesembolso(ByVal psCtaCod As String, ByVal dHoy As Date) As String
Dim oCredito As Dcalendario
Dim oCredDatos As DCredito
Dim oCredGen As DCredGeneral
Dim R As ADODB.Recordset
Dim RTemp As ADODB.Recordset
Dim nDias As Integer
Dim nDiasMin As Integer
Dim dTemp As Date
Dim nCuota As Integer
Dim nNroCalen As Integer
Dim nTipoDesemb As Integer

    On Error GoTo ErrorValidaCargaDatosDesembolso
    ValidaCargaDatosDesembolso = ""
    
    Set oCredDatos = New DCredito
    Set R = oCredDatos.RecuperaColocacCred(psCtaCod)
    If R.EOF And R.BOF Then
        Exit Function
    End If
    nTipoDesemb = R!nTipoDesembolso
    Set R = oCredDatos.RecuperaProducto(psCtaCod)
    Set oCredDatos = Nothing
    If R.RecordCount = 0 Then
        ValidaCargaDatosDesembolso = "Credito No ha sido Aprobado"
        R.Close
        Set R = Nothing
        Exit Function
    Else
        If (R!nPrdEstado = gColocEstVigMor Or R!nPrdEstado = gColocEstVigNorm Or R!nPrdEstado = gColocEstVigVenc _
            Or R!nPrdEstado = gColocEstRefMor Or R!nPrdEstado = gColocEstRefNorm Or R!nPrdEstado = gColocEstRefVenc) And nTipoDesemb <> 1 Then
            ValidaCargaDatosDesembolso = "Credito ya Esta Vigente"
            R.Close
            Set R = Nothing
            Exit Function
            
        Else
        
            If R!nPrdEstado <> gColocEstAprob And nTipoDesemb <> 1 Then
                ValidaCargaDatosDesembolso = "Credito No ha sido Aprobado"
                R.Close
                Set R = Nothing
                Exit Function
            End If
        End If
    End If
    R.Close
    Set R = Nothing
    
    Set oCredDatos = New DCredito
    Set R = oCredDatos.RecuperaColocacCred(psCtaCod)
    Set oCredDatos = Nothing
    nCuota = IIf(IsNull(R!nNroProxDesemb), 0, R!nNroProxDesemb)
    nNroCalen = IIf(IsNull(R!nNroCalen), 0, R!nNroCalen)
    R.Close
    Set R = Nothing
    
    Set oCredito = New Dcalendario
    Set R = oCredito.RecuperaColocCalendario(psCtaCod, nNroCalen, gColocCalendAplDesembolso, nCuota)
    Set oCredito = Nothing
    dTemp = CDate(Format(R!DVENC, "dd/mm/yyyy"))
    R.Close
    Set R = Nothing
    
    Set oCredGen = New DCredGeneral
    Set RTemp = oCredGen.RecuperaParametro(gColocTiposDesembolso)
    Set oCredGen = Nothing
    nDias = IIf(IsNull(RTemp!nParamValor), 0, RTemp!nParamValor)
    RTemp.Close
    Set RTemp = Nothing
        
    If DateDiff("d", dTemp, dHoy) > nDias Then
        ValidaCargaDatosDesembolso = "La Fecha del Desembolso " & Format(dTemp, "dd/mm/yyyy") & " a rebazado el Minimo de Tolerancia de dias para Desembolsar "
        Exit Function
    End If
    
    '
    Set oCredGen = New DCredGeneral
    Set RTemp = oCredGen.RecuperaParametro(3019)
    Set oCredGen = Nothing
    nDiasMin = IIf(IsNull(RTemp!nParamValor), 0, RTemp!nParamValor)
    RTemp.Close
    Set RTemp = Nothing
        
    If DateDiff("d", dTemp, dHoy) > nDiasMin Then
        ValidaCargaDatosDesembolso = "La Fecha del Desembolso " & Format(dTemp, "dd/mm/yyyy") & " es inferior al Minimo de Tolerancia de dias para Desembolsar "
        Exit Function
    End If
    
    
    Exit Function

ErrorValidaCargaDatosDesembolso:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
    
End Function

Public Function SugerenciaCredito(ByVal psCtaCod As String, ByVal pnEstadoActual As ColocEstado, ByVal pnEstado As ColocEstado, ByVal pdEstado As Date, ByVal pnNroTransac As Long, ByVal psLineaCred As String, ByVal pnTasa As Double, ByVal pnMonto As Double, _
        ByVal pnNroCuotas As Integer, ByVal pnPlazo As Integer, ByVal pnTipoCuota As ColocTipoCalend, ByVal pnDiaFijo As Integer, _
        ByVal pnProxMes As Integer, ByVal pnTipoDesemb As ColocTiposDesembolso, ByVal pnCalendDinamico As Integer, _
        ByVal pnGraciaSug As Integer, ByVal pnTasaGracia As Double, ByVal pnTipoGracia As ColocTiposGracia, _
        ByVal nNroCalend As Integer, ByVal MatCalendDes As Variant, ByVal MatCalendPagos As Variant, _
        ByVal MatCalendPagos_2 As Variant, ByVal pnCuotaCom As Integer, ByVal pnMiViv As Integer, _
        ByVal pnNroCalenPar As Integer, ByVal psTipoGasto As String, ByVal pnTasaMora As Double, ByVal pMatCredVig As Variant, ByVal psDescripcion As String) As String
        
Dim oDCredito As DCredActualizaBD
Dim oLinea As DLineaCredito
Dim R As ADODB.Recordset
Dim i As Integer


    On Error GoTo ErrorSugerenciaCredito
    

    
    SugerenciaCredito = ValidaGarantia(psCtaCod, pdEstado, pnMonto)
    If SugerenciaCredito <> "" Then
        Exit Function
    End If
    
    Set oDCredito = New DCredActualizaBD
    oDCredito.dBeginTrans
    Call oDCredito.dUpdateProducto(psCtaCod, pnTasa, 0#, pnEstado, pdEstado, pnNroTransac, False)
    
    'Inserta Tasa Compensatoria
    If pnEstadoActual <> pnEstado Then
        Call oDCredito.dInsertProductoTasaInteres(psCtaCod, gColocLineaCredTasasIntCompNormal, pnTasa, False)
    Else
        Call oDCredito.dUpdateProductoTasaInteres(psCtaCod, gColocLineaCredTasasIntCompNormal, pnTasa, False)
    End If
    
    'Inserta TasaGracia
    If pnGraciaSug > 0 Then
        Call oDCredito.dDeleteProductoTasaInteres(psCtaCod, gColocLineaCredTasasIntGracia, False)
        Call oDCredito.dInsertProductoTasaInteres(psCtaCod, gColocLineaCredTasasIntGracia, pnTasaGracia, False)
    Else
        Call oDCredito.dDeleteProductoTasaInteres(psCtaCod, gColocLineaCredTasasIntGracia, False)
    End If
    
    'Interes Compenstorio Vencido
    Call oDCredito.dDeleteProductoTasaInteres(psCtaCod, 6, False)
    Call oDCredito.dInsertProductoTasaInteres(psCtaCod, 6, Format(((1 + pnTasa / 100) ^ (1 / 30) - 1) * 100, "#0.0000"), False)
    'Interes Compenstorio Mora
    Call oDCredito.dDeleteProductoTasaInteres(psCtaCod, gColocLineaCredTasasIntMoratNormal, False)
    Call oDCredito.dInsertProductoTasaInteres(psCtaCod, gColocLineaCredTasasIntMoratNormal, pnTasaMora, False)
    
    Call oDCredito.dUpdateColocaciones(psCtaCod, pnPlazo, CDate(MatCalendPagos(UBound(MatCalendPagos) - 1, 0)), 0, pnTipoCuota, psLineaCred, pdEstado, False)
    If pnEstadoActual <> pnEstado Then
        Call oDCredito.dInsertColocacEstado(psCtaCod, pdEstado, pnEstado, pnNroCuotas, pnMonto, Replace(psDescripcion, "'", "''"), pnTipoCuota, pnDiaFijo, pnGraciaSug, pnPlazo, pnTipoGracia, pnTipoDesemb, pnProxMes, pnCalendDinamico, False, , psTipoGasto)
    Else
        Call oDCredito.dUpdateColocacEstado(psCtaCod, pdEstado, pnEstado, pnNroCuotas, pnMonto, Replace(psDescripcion, "'", "''"), pnTipoCuota, pnDiaFijo, pnGraciaSug, pnPlazo, pnTipoGracia, pnTipoDesemb, pnProxMes, pnCalendDinamico, False, psTipoGasto)
    End If
    
    Set oLinea = New DLineaCredito
    Set R = oLinea.RecuperaLineadeCredito(psLineaCred)
    Set oLinea = Nothing
    Call oDCredito.dUpdateColocacCred(psCtaCod, , , , , , , , , , , pnTipoCuota, pnCalendDinamico, pnTipoDesemb, nNroCalend, , R!cPersCod, IIf(IsNull(R!nmoneda), CInt(Mid(psLineaCred, 5, 1)), R!nmoneda), False, pnCuotaCom, pnMiViv, pnNroCalenPar)
    R.Close
    Set R = Nothing
    'Borrado de Calendario
    Call oDCredito.dDeleteColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplDesembolso, , , False)
    Call oDCredito.dDeleteColocCalendario(psCtaCod, nNroCalend, gColocCalendAplDesembolso, , False)

    Call oDCredito.dDeleteColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, , , False)
    Call oDCredito.dDeleteColocCalendario(psCtaCod, nNroCalend, gColocCalendAplCuota, , False)
        
    
    'Nuevo Calendario
    For i = 0 To UBound(MatCalendDes) - 1
        Call oDCredito.dInsertColocCalendario(psCtaCod, nNroCalend, gColocCalendAplDesembolso, i + 1, CDate(MatCalendDes(i, 0)), gColocCalendEstadoPendiente, "Calendario de Sugerencia Desembolsos", gColocCalendConceptoProcSugerido, False)
        Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplDesembolso, i + 1, gColocConceptoCodCapital, CDbl(MatCalendDes(i, 1)), 0, "", False)
    Next i
    For i = 0 To UBound(MatCalendPagos) - 1
        Call oDCredito.dInsertColocCalendario(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, CDate(MatCalendPagos(i, 0)), gColocCalendEstadoPendiente, "Calendario de Sugerencia Pagos", gColocCalendConceptoProcSugerido, False)
        If pnTipoCuota <> gColocCalendCodCL Then
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, gColocConceptoCodCapital, CDbl(MatCalendPagos(i, 3)), 0, "", False)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalendPagos(i, 4)), 0, "", False)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", False)
            If CDbl(MatCalendPagos(i, 5)) > 0 Then
                Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresGracia, CDbl(MatCalendPagos(i, 5)), 0, "", False)
            End If
        Else
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, gColocConceptoCodCapital, CDbl(MatCalendPagos(i, 3)), 0, "", False)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalendPagos(i, 4)), 0, "", False)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", False)
        End If
    Next i
    
    If pnMiViv = 1 Then
        Call oDCredito.dDeleteColocCalendDet(psCtaCod, pnNroCalenPar, gColocCalendAplCuota, , , False)
        Call oDCredito.dDeleteColocCalendario(psCtaCod, pnNroCalenPar, gColocCalendAplCuota, , False)
        For i = 0 To UBound(MatCalendPagos_2) - 1
            Call oDCredito.dInsertColocCalendario(psCtaCod, pnNroCalenPar, gColocCalendAplCuota, i + 1, CDate(MatCalendPagos_2(i, 0)), gColocCalendEstadoPendiente, "Calendario de Sugerencia Pagos", gColocCalendConceptoProcSugerido, False)
            If pnTipoCuota <> gColocCalendCodCL Then
                Call oDCredito.dInsertColocCalendDet(psCtaCod, pnNroCalenPar, gColocCalendAplCuota, i + 1, gColocConceptoCodCapital, CDbl(MatCalendPagos_2(i, 3)), 0, "", False)
                Call oDCredito.dInsertColocCalendDet(psCtaCod, pnNroCalenPar, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalendPagos_2(i, 4)), 0, "", False)
                Call oDCredito.dInsertColocCalendDet(psCtaCod, pnNroCalenPar, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", False)
                If CDbl(MatCalendPagos_2(i, 5)) > 0 Then
                    Call oDCredito.dInsertColocCalendDet(psCtaCod, pnNroCalenPar, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresGracia, CDbl(MatCalendPagos_2(i, 5)), 0, "", False)
                End If
            End If
        Next i
    End If
    
    Call ActualizaCreditosVigentes(pMatCredVig)
    
    'oDCredito.dEjecutaBatch
    oDCredito.dCommitTrans
    Set oDCredito = Nothing
    
    Exit Function

ErrorSugerenciaCredito:
    oDCredito.dRollbackTrans
    Err.Raise Err.Number, "Error En Proceso SugerenciaCredito", Err.Description
    
End Function
Public Function EsRefinanciado(ByVal psCtaCod As String, Optional ByVal PregCapInt As Boolean = False) As Boolean
Dim oCred As DCredito
Dim R As ADODB.Recordset

    On Error GoTo ErrorEsRefinanciado
    Set oCred = New DCredito
    Set R = oCred.RecuperaColocacRefinanc(psCtaCod)
    
    
    Set oCred = Nothing
    If R.RecordCount > 0 Then
        If PregCapInt Then
            Set oCred = New DCredito
            Set R = oCred.RecuperaColocacCred(psCtaCod)
            Set oCred = Nothing
            If R!bRefCapInt = True Then
                EsRefinanciado = True
            Else
                EsRefinanciado = False
            End If
            R.Close
        Else
            EsRefinanciado = True
        End If
        
    Else
        EsRefinanciado = False
    End If
    Exit Function

ErrorEsRefinanciado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Function EsSustiDeudor(ByVal psCtaCod As String, Optional ByVal PregCapInt As Boolean = False) As Boolean
Dim oCred As DCredito
Dim R As ADODB.Recordset

    On Error GoTo ErrorEsSustiDeudor
    Set oCred = New DCredito
    Set R = oCred.RecuperaColocacRefinanc(psCtaCod)
    
    Set oCred = Nothing
    If R.RecordCount > 0 Then
        If R!bSustiDeudor Then
            'Set oCred = New DCredito
            'Set R = oCred.RecuperaColocacCred(psCtaCod)
            'Set oCred = Nothing
            'If R!bRefCapInt = True Then
            '    EsSustiDeudor = True
            'Else
            '    EsSustiDeudor = False
            'End If
            R.Close
            EsSustiDeudor = True
        Else
            EsSustiDeudor = False
        End If
        
    Else
        EsSustiDeudor = False
    End If
    Exit Function

ErrorEsSustiDeudor:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function MontosARefinanciar(ByVal psCtaCod As String, ByVal pdHoy As Date) As Variant
Dim R As ADODB.Recordset
Dim MatCalend As Variant
Dim MatCalendTempo As Variant
Dim nMontoInt As Double
Dim nMontoGastos As Double
Dim nMontoCap As Double
Dim oDCredito As DCredito

    On Error GoTo ErrorMontosARefinanciar
    nMontoInt = 0
    nMontoGastos = 0
    nMontoCap = 0
    
    Set oDCredito = New DCredito
    Set R = oDCredito.RecuperaCreditosRefinanciados(psCtaCod)
    Set oDCredito = Nothing
    Do While Not R.EOF
        MatCalend = RecuperaMatrizCalendarioPendiente(R!cCtaCodRef)
        MatCalendTempo = RecuperaMatrizCalendarioPendiente(R!cCtaCodRef, True)
        MatCalend = UnirMatricesMiViviendaAmortizacion(MatCalend, MatCalendTempo)
        
        nMontoInt = nMontoInt + MatrizInteresGastosAFecha(R!cCtaCodRef, MatCalend, pdHoy) - MatrizGastosFecha(R!cCtaCodRef, MatCalend)
        nMontoGastos = nMontoGastos + MatrizGastosFecha(R!cCtaCodRef, MatCalend)
        nMontoCap = nMontoCap + MatrizCapitalAFecha(R!cCtaCodRef, MatCalend)
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    MontosARefinanciar = Array(Format(nMontoInt, "#0.00"), Format(nMontoGastos, "#0.00"), Format(nMontoCap, "#0.00"))
    Exit Function

ErrorMontosARefinanciar:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Private Function ValidaGarantia(ByVal psCtaCod As String, ByVal pdGarantia As Date, ByVal pnMonto As Double) As String
Dim oCred As DCredito
Dim nMontoGarantia As Double

    On Error GoTo ErrorValidaSugerenciaAprobacion
    Set oCred = New DCredito
    nMontoGarantia = oCred.RecuperaMontoGarantiaCredito(psCtaCod, pdGarantia)
    Set oCred = Nothing
    ValidaGarantia = ""
    If nMontoGarantia < pnMonto Then
        ValidaGarantia = "Garantia No Cubre el Monto del Credito"
        Exit Function
    End If
    Exit Function

ErrorValidaSugerenciaAprobacion:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function CalendarioDesembolsos(ByVal psCtaCod As String, ByVal pnMonto As Double, _
        ByVal pnTipoDesemb As ColocTiposDesembolso, ByVal nNroCalend As Integer, _
        ByVal MatCalendDes As Variant) As String
        
Dim oDCredito As DCredActualizaBD
Dim oDCalend As Dcalendario
Dim i As Integer
Dim bTransac As Boolean

    On Error GoTo ErrorCalendarioDesembolsos
    bTransac = False

    CalendarioDesembolsos = ValidaGrabarDesembolso(psCtaCod, pnMonto)

    If CalendarioDesembolsos <> "" Then
        Exit Function
    End If
        
    Set oDCredito = New DCredActualizaBD
    Call oDCredito.dBeginTrans
    bTransac = True
        
    Call oDCredito.dUpdateColocacCred(psCtaCod, , , , , , , , , , , , , , nNroCalend, , , , False, , , -1)
     
    'NUEVO CALENDARIO

    'Desembolsos
    Dim nMontoPagado As Double
    
    For i = 0 To UBound(MatCalendDes) - 1
        
        If CInt(MatCalendDes(i, 2)) = 1 Then
           nMontoPagado = CDbl(MatCalendDes(i, 1))
        Else
            nMontoPagado = 0
        End If
        
        Call oDCredito.dInsertColocCalendario(psCtaCod, nNroCalend, gColocCalendAplDesembolso, i + 1, CDate(MatCalendDes(i, 0)), CInt(MatCalendDes(i, 2)), "Nuevo Calendario de Desembolsos", gColocCalendConceptoProcOtro, False)
        Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplDesembolso, i + 1, gColocConceptoCodCapital, CDbl(MatCalendDes(i, 1)), nMontoPagado, "", False)
    Next i


    'Pagos
    Dim RColocCalendPagosActual As ADODB.Recordset
    Dim RCalendPagosDetActual As ADODB.Recordset

    Set oDCalend = New Dcalendario
    
        Set RColocCalendPagosActual = oDCalend.RecuperaColocCalendario(psCtaCod, nNroCalend - 1, gColocCalendAplCuota)
        Set RCalendPagosDetActual = oDCalend.RecuperaColocCalendDet(psCtaCod, nNroCalend - 1, gColocCalendAplCuota)

    'Inserta ColocCalendario
        Do While Not RColocCalendPagosActual.EOF
           Call oDCredito.dInsertColocCalendario(psCtaCod, nNroCalend, RColocCalendPagosActual!nColocCalendApl, RColocCalendPagosActual!nCuota, RColocCalendPagosActual!DVENC, RColocCalendPagosActual!nColocCalendEstado, RColocCalendPagosActual!CDescripcion, RColocCalendPagosActual!nCalendProc, False)
           RColocCalendPagosActual.MoveNext
        Loop
        RColocCalendPagosActual.Close
        
        'Inserta ColocCalendario ColocCalendDet
        Do While Not RCalendPagosDetActual.EOF
           Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, RCalendPagosDetActual!nColocCalendApl, RCalendPagosDetActual!nCuota, RCalendPagosDetActual!nPrdConceptoCod, RCalendPagosDetActual!nMonto, RCalendPagosDetActual!nMontoPagado, RCalendPagosDetActual!cFlag, False)
           RCalendPagosDetActual.MoveNext
        Loop
        RCalendPagosDetActual.Close

  
    oDCredito.dCommitTrans
    Set oDCredito = Nothing
    
    Exit Function

ErrorCalendarioDesembolsos:
    If bTransac Then
        oDCredito.dRollbackTrans
        Set oDCredito = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso CalendarioDesembolsos", Err.Description
    
End Function

Public Sub ImprimeConsolidadoCred(ByVal psCtaCod As String, ByVal psOrigen As String, ByVal pdFecha As Date)
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim sql  As String
Set oCon = New DConecta
Set rs = New ADODB.Recordset

oCon.AbreConexion
sql = "SELECT GETDATE() "
Set rs = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing
' dependiendo del tipo de credito

With DRSugerencia
      '.Orientation =
      Set .DataSource = rs
      .DataMember = ""
      .inicio psCtaCod, psOrigen, pdFecha
      .Refresh
      .Show vbModal
End With
End Sub

Public Sub ImprimeConsolidadoConsumo(ByVal psCtaCod As String, ByVal psOrigen As String, ByVal pdFecha As Date)
    Dim rs As ADODB.Recordset
    Dim oCon As DConecta
    Dim sql As String
    Set oCon = New DConecta
    
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    sql = "Select GETDATE() "
    Set rs = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    Set oCon = Nothing
    
'    With DRConsumo
'        Set .DataSource = Rs
'           .DataMember = ""
'        .inicio psCtaCod, psOrigen, pdFecha
'        .Refresh
'        .Show vbModal
'    End With
End Sub
' Creado para Imprimer contrato Agricola
Public Sub ImprimeConsolidadoCredAgricola(ByVal pstrCtaCod As String, ByVal psOrigen As String, pdFecha As Date)
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim sql  As String
Set oCon = New DConecta
Set rs = New ADODB.Recordset

oCon.AbreConexion
sql = "SELECT GETDATE() "
Set rs = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing
' dependiendo del tipo de credito

With DRPrestamoAgricola
      Set .DataSource = rs
      .DataMember = ""
      .inicio pstrCtaCod, psOrigen, pdFecha
     ' .Orientation = rptOrientLandscape
    
      .Refresh
      .Show vbModal
End With



End Sub


Public Function AprobacionCredito(ByVal psCtaCod As String, ByVal pnEstado As ColocEstado, ByVal pdEstado As Date, ByVal pnNroTransac As Long, ByVal psLineaCred As String, ByVal pnTasa As Double, ByVal pnMonto As Double, _
        ByVal pnNroCuotas As Integer, ByVal pnPlazo As Integer, ByVal pnTipoCuota As ColocTipoCalend, ByVal pnDiaFijo As Integer, _
        ByVal pnProxMes As Integer, ByVal pnTipoDesemb As ColocTiposDesembolso, ByVal pnCalendDinamico As Integer, _
        ByVal pnGraciaSug As Integer, ByVal pnTasaGracia As Double, ByVal pnTipoGracia As ColocTiposGracia, _
        ByVal nNroCalend As Integer, ByVal MatCalendDes As Variant, ByVal MatCalendPagos As Variant, _
        ByVal MatGastos As Variant, ByVal pNumGastos As Integer, ByVal psApoderado As String, _
        ByVal pnCuotaCom As Integer, ByVal pnMiViv As Integer, ByVal pnNroCalenPar As Integer, _
        Optional ByVal MatCalendPagos_2 As Variant = "", _
        Optional pbRefinanc As Boolean = False, Optional pbCapitalInter As Boolean = False, _
        Optional ByVal psCodAge As String = "", Optional ByVal psCodUsu As String = "", Optional ByVal psTipoGasto As String = "F", Optional ByVal MatCredVig As Variant = "") As String
        
Dim oDCredito As DCredActualizaBD
Dim R As ADODB.Recordset
Dim RRef As ADODB.Recordset
Dim i As Integer
Dim nMontoReservado As Double
Dim bTransac As Boolean
Dim MatCredRef As Variant
Dim MatCredRefTempo As Variant
Dim MatCalend As Variant
Dim oCredito As DCredito
Dim oFunciones As NContFunciones
Dim sMovNro As String
Dim nMovNro As Long
Dim nMontoRefTmp As Double
Dim nMontoRefTmpCapital As Double
Dim nMontoRefTmpInteres As Double
Dim nMontoRefTmpInterSusp As Double
Dim nMontoRefTmpGastosSusp As Double
Dim nMontoRefTmpNew As Double
Dim nMontoRefTmpCapitalNew As Double
Dim nMontoRefTmpInteresNew As Double
Dim nMontoRefTmpInterSuspNew As Double
Dim nMontoRefTmpGastosSuspNew As Double
Dim nMontoRefTempo As Double

Dim oGeneral As DGeneral
Dim nEstadCredRef As Integer
Dim sConsCred As String
Dim sConsCred2 As String
Dim nTipoCambioFijo As Double
Dim nIntSuspTot As Double
Dim nGastoSuspTot As Double

Dim MatrizMatCredRef() As Variant
Dim MatrizGastosCredRef() As Variant
Dim MatrizGastosCredRefTempo As Variant
Dim MatrizGastosCredGrp As Variant
Dim k As Integer

Dim nEstado As Integer


Dim sCadCtaGastoTemp As String


    On Error GoTo ErrorSugerenciaCredito
    nIntSuspTot = 0
    nGastoSuspTot = 0
    bTransac = False
    AprobacionCredito = ValidaGarantia(psCtaCod, pdEstado, pnMonto)
    If AprobacionCredito <> "" Then
        Exit Function
    End If
    AprobacionCredito = ValidaGrabarDesembolso(psCtaCod, pnMonto)
    If AprobacionCredito <> "" Then
        Exit Function
    End If
        
    If pbRefinanc Then
        '************************************************
        'Recupera los calendarios de Pagos de Todos los Creditos Pendientes
        '************************************************
        Set oCredito = New DCredito
        Set R = oCredito.RecuperaColocacRefinanc(psCtaCod)
        ReDim MatrizMatCredRef(0)
        ReDim MatrizGastosCredRef(0)
        sCadCtaGastoTemp = "('"
        Do While Not R.EOF
            MatCredRef = RecuperaMatrizCalendarioPendiente(R!cCtaCodRef)
            MatCredRefTempo = RecuperaMatrizCalendarioPendiente(R!cCtaCodRef, True)
            MatCredRef = UnirMatricesMiViviendaAmortizacion(MatCredRef, MatCredRefTempo)
            
            ReDim Preserve MatrizMatCredRef(R.Bookmark)
            MatrizMatCredRef(R.Bookmark - 1) = MatCredRef
            
            ReDim Preserve MatrizGastosCredRef(R.Bookmark)
            MatrizGastosCredRef(R.Bookmark - 1) = RecuperaMatrizGastosPendiente(R!cCtaCodRef)
            
            sCadCtaGastoTemp = sCadCtaGastoTemp & R!cCtaCodRef & "',"
            R.MoveNext
        Loop
        R.Close
        Set oCredito = Nothing
        
        MatCredRef = RecuperaMatrizRefinanciados(psCtaCod)
        
        'Obtiene Gastos Agrupados de Todos los Calendarios
        sCadCtaGastoTemp = Mid(sCadCtaGastoTemp, 1, Len(sCadCtaGastoTemp) - 1) & ")"
        MatrizGastosCredGrp = RecuperaMatrizGastosPendienteAgrupado(sCadCtaGastoTemp)
        
    End If
    
    Set oFunciones = New NContFunciones
    sMovNro = oFunciones.GeneraMovNro(pdEstado, psCodAge, psCodUsu)
    Set oFunciones = Nothing
    
    Set oDCredito = New DCredActualizaBD
    Call oDCredito.dBeginTrans
    bTransac = True
    Call oDCredito.dUpdateProducto(psCtaCod, pnTasa, 0#, pnEstado, pdEstado, pnNroTransac, False)
    
    'Ingresa Apoderado
    Call oDCredito.dInsertProductoPersona(psCtaCod, psApoderado, gColRelPersApoderado, False)
    'Reserva Monto de Prestamo en Saldos de Linea de Credito
    Set R = oDCredito.RecuperaLineasCreditoSaldo(psLineaCred, True)
    nMontoReservado = IIf(IsNull(R!nMontoReservado), 0, CDbl(Format(R!nMontoReservado, "#0.00")))
    nMontoReservado = nMontoReservado + pnMonto
    R.Close
    Set R = Nothing
    Call oDCredito.dUpdateLineaCreditoSaldo(psLineaCred, , , , , nMontoReservado, False)
    
    'Inserta TasaGracia
    If pnGraciaSug = 0 Then
        Call oDCredito.dDeleteProductoTasaInteres(psCtaCod, gColocTasaCompGraCia)
    End If
        
    Call oDCredito.dUpdateColocaciones(psCtaCod, pnPlazo, CDate(MatCalendPagos(UBound(MatCalendPagos) - 1, 0)), 0, pnTipoCuota, Trim(psLineaCred), pdEstado, False)
    Call oDCredito.dInsertColocacEstado(psCtaCod, pdEstado, pnEstado, pnNroCuotas, pnMonto, "Aprobacion de Credito", pnTipoCuota, pnDiaFijo, pnGraciaSug, pnPlazo, pnTipoGracia, pnTipoDesemb, pnProxMes, pnCalendDinamico, False, , psTipoGasto)
    
    Call oDCredito.dUpdateColocacCred(psCtaCod, , , , , , , , , , , pnTipoCuota, pnCalendDinamico, pnTipoDesemb, nNroCalend, 1, , , False, pnCuotaCom, pnMiViv, pnNroCalenPar)
    
    
    'Nuevo Calendario
    'Desembolsos
    For i = 0 To UBound(MatCalendDes) - 1
        Call oDCredito.dInsertColocCalendario(psCtaCod, nNroCalend, gColocCalendAplDesembolso, i + 1, CDate(MatCalendDes(i, 0)), gColocCalendEstadoPendiente, "Calendario de Aprobacion Desembolsos", gColocCalendConceptoProcAprobado, False)
        Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplDesembolso, i + 1, gColocConceptoCodCapital, CDbl(MatCalendDes(i, 1)), 0, "", False)
    Next i
    'Pagos
    For i = 0 To UBound(MatCalendPagos) - 1
        Call oDCredito.dInsertColocCalendario(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, CDate(MatCalendPagos(i, 0)), gColocCalendEstadoPendiente, "Calendario de Aprobacion Pagos", gColocCalendConceptoProcAprobado, False)
        If pnTipoCuota <> gColocCalendCodCL Then
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, gColocConceptoCodCapital, CDbl(MatCalendPagos(i, 3)), 0, "", False)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalendPagos(i, 4)), 0, "", False)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", False)
            If CDbl(MatCalendPagos(i, 5)) > 0 Then
                Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresGracia, CDbl(MatCalendPagos(i, 5)), 0, "", False)
            End If
        Else
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, gColocConceptoCodCapital, CDbl(MatCalendPagos(i, 3)), 0, "", False)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalendPagos(i, 4)), 0, "", False)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", False)
        End If
    Next i
    
    'Gastos
    If Not pbRefinanc Then
        For i = 0 To pNumGastos - 1
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, CInt(Trim(Right(MatGastos(i, 0), 3))), CInt(Trim(MatGastos(i, 1))), CLng(Trim(Right(MatGastos(i, 2), 20))), CDbl(MatGastos(i, 3)), 0, "", False)
        Next i
    End If
    
    '*********************************************************
    '************* REFINANCIACION DE CREDITO *****************
    '*********************************************************
    If pbRefinanc Then
        
        Set oGeneral = New DGeneral
        nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdEstado, TCFijoMes)
        Set oGeneral = Nothing
        
        Call oDCredito.dInsertColocacEstado(psCtaCod, pdEstado, gColocEstRefNorm, pnNroCuotas, pnMonto, "Refinanciacion de Credito", pnTipoCuota, pnDiaFijo, pnGraciaSug, pnPlazo, pnTipoGracia, pnTipoDesemb, pnProxMes, pnCalendDinamico, False, , psTipoGasto)
        nMontoRefTmpNew = 0
        nMontoRefTmpCapitalNew = 0
        nMontoRefTmpInteresNew = 0
        Call oDCredito.dInsertMov(sMovNro, gCredRefinanciacion, "", gMovEstContabMovContable, gMovFlagVigente, False)
        nMovNro = oDCredito.dGetnMovNro(sMovNro)
        
        Set oCredito = New DCredito
        Set R = oCredito.RecuperaColocacRefinanc(psCtaCod)
        Do While Not R.EOF
            Set RRef = oCredito.RecuperaProducto(R!cCtaCodRef)
            nEstadCredRef = RRef!nPrdEstado
            RRef.Close
            Set RRef = Nothing
            MatCredRef = MatrizMatCredRef(R.Bookmark - 1)
            
            If pbCapitalInter Then
                
                If Mid(R!cCtaCodRef, 9, 1) = Mid(R!cCtaCod, 9, 1) Then
                    nMontoRefTempo = MatrizInteresGastosAFecha(R!cCtaCodRef, MatCredRef, pdEstado) + MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)
                Else
                    If Mid(R!cCtaCod, 9, 1) = "1" Then
                        nMontoRefTempo = (MatrizInteresGastosAFecha(R!cCtaCodRef, MatCredRef, pdEstado) + MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)) * nTipoCambioFijo
                        nMontoRefTempo = Format(nMontoRefTempo, "#0.00")
                    Else
                        nMontoRefTempo = (MatrizInteresGastosAFecha(R!cCtaCodRef, MatCredRef, pdEstado) + MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)) / nTipoCambioFijo
                        nMontoRefTempo = Format(nMontoRefTempo, "#0.00")
                    End If
                End If
                
                Call oDCredito.dInsertColocacRefinanc(psCtaCod, R!cCtaCodRef, nMontoRefTempo, gColocEstadoRefinancAprobado, pdEstado, False, R!bSustiDeudor)
            Else
                If Mid(R!cCtaCodRef, 9, 1) = Mid(R!cCtaCod, 9, 1) Then
                    nMontoRefTempo = MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)
                Else
                    If Mid(R!cCtaCod, 9, 1) = "1" Then
                        nMontoRefTempo = MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef) * nTipoCambioFijo
                        nMontoRefTempo = Format(nMontoRefTempo, "#0.00")
                    Else
                        nMontoRefTempo = MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef) / nTipoCambioFijo
                        nMontoRefTempo = Format(nMontoRefTempo, "#0.00")
                    End If
                End If
                ' CMACICA_CSTS - 11/11/2003 ----------------------------
                Call oDCredito.dInsertColocacRefinanc(psCtaCod, R!cCtaCodRef, nMontoRefTempo, gColocEstadoRefinancAprobado, pdEstado, False, R!bSustiDeudor)
                '-------------------------------------------------------
            End If
                                    
            'Inserta en ColocacCredRefinancDet y en el Calendario
                If pbCapitalInter Then
                
                    If Mid(R!cCtaCodRef, 9, 1) = Mid(R!cCtaCod, 9, 1) Then
                        nMontoRefTempo = MatrizInteresGastosAFecha(R!cCtaCodRef, MatCredRef, pdEstado) + MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)
                    Else
                        If Mid(R!cCtaCod, 9, 1) = "1" Then
                            nMontoRefTempo = (MatrizInteresGastosAFecha(R!cCtaCodRef, MatCredRef, pdEstado) + MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)) * nTipoCambioFijo
                            nMontoRefTempo = Format(nMontoRefTempo, "#0.00")
                        Else
                            nMontoRefTempo = (MatrizInteresGastosAFecha(R!cCtaCodRef, MatCredRef, pdEstado) + MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)) / nTipoCambioFijo
                            nMontoRefTempo = Format(nMontoRefTempo, "#0.00")
                        End If
                    End If
                
                    Call oDCredito.dInsertColocacRefinancDet(psCtaCod, R!cCtaCodRef, nMontoRefTempo, 0, gColocEstadoRefinancAprobado, gColocConceptoCodCapital, pdEstado, False)
                Else
                    If Mid(R!cCtaCodRef, 9, 1) = Mid(R!cCtaCod, 9, 1) Then
                        nMontoRefTempo = MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)
                    Else
                        If Mid(R!cCtaCod, 9, 1) = "1" Then
                            nMontoRefTempo = MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef) * nTipoCambioFijo
                            nMontoRefTempo = Format(nMontoRefTempo, "#0.00")
                        Else
                            nMontoRefTempo = MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef) / nTipoCambioFijo
                            nMontoRefTempo = Format(nMontoRefTempo, "#0.00")
                        End If
                    End If
                    
                    Call oDCredito.dInsertColocacRefinancDet(psCtaCod, R!cCtaCodRef, nMontoRefTempo, 0, gColocEstadoRefinancAprobado, gColocConceptoCodCapital, pdEstado, False)
                End If
                
                
                    If MatrizInteresReprogramadoFecha(R!cCtaCodRef, MatCredRef) > 0 Then
                        Call oDCredito.dInsertColocacRefinancDet(psCtaCod, R!cCtaCodRef, MatrizInteresReprogramadoFecha(R!cCtaCodRef, MatCredRef), 0, gColocEstadoRefinancAprobado, gColocConceptoCodInteresReprogramado, pdEstado, False)
                    End If
                    If MatrizInteresSuspensoFecha(R!cCtaCodRef, MatCredRef) > 0 Then
                        Call oDCredito.dInsertColocacRefinancDet(psCtaCod, R!cCtaCodRef, MatrizInteresSuspensoFecha(R!cCtaCodRef, MatCredRef), 0, gColocEstadoRefinancAprobado, gColocConceptoCodInteresSuspenso, pdEstado, False)
                    End If
                    If MatrizInteresGraciaFecha(R!cCtaCodRef, MatCredRef, pdEstado) > 0 Then
                        Call oDCredito.dInsertColocacRefinancDet(psCtaCod, R!cCtaCodRef, MatrizInteresGraciaFecha(R!cCtaCodRef, MatCredRef, pdEstado), 0, gColocEstadoRefinancAprobado, gColocConceptoCodInteresGracia, pdEstado, False)
                    End If
                    If MatrizInteresMorFecha(R!cCtaCodRef, MatCredRef) > 0 Then
                        Call oDCredito.dInsertColocacRefinancDet(psCtaCod, R!cCtaCodRef, MatrizInteresMorFecha(R!cCtaCodRef, MatCredRef), 0, gColocEstadoRefinancAprobado, gColocConceptoCodInteresMoratorio, pdEstado, False)
                    End If
                    If MatrizInteresCompAFecha(R!cCtaCodRef, MatCredRef, pdEstado) > 0 Then
                        Call oDCredito.dInsertColocacRefinancDet(psCtaCod, R!cCtaCodRef, MatrizInteresCompAFecha(R!cCtaCodRef, MatCredRef, pdEstado), 0, gColocEstadoRefinancAprobado, gColocConceptoCodInteresCompensatorio, pdEstado, False)
                    End If
                    If MatrizInteresCompVencidoFecha(R!cCtaCodRef, MatCredRef) > 0 Then
                        Call oDCredito.dInsertColocacRefinancDet(psCtaCod, R!cCtaCodRef, MatrizInteresCompVencidoFecha(R!cCtaCodRef, MatCredRef), 0, gColocEstadoRefinancAprobado, gColocConceptoCodInteresCompVencido, pdEstado, False)
                    End If
                    If MatrizGastosFecha(R!cCtaCodRef, MatCredRef) > 0 Then
                        Call oDCredito.dInsertColocacRefinancDet(psCtaCod, R!cCtaCodRef, MatrizGastosFecha(R!cCtaCodRef, MatCredRef), 0, gColocEstadoRefinancAprobado, gColocConceptoCodGastoVarios, pdEstado, False)
                    End If
                
            'Halla la Constante
            If pbCapitalInter Then
                Select Case nEstadCredRef
                    Case gColocEstRefMor
                        If CInt(Mid(R!cCtaCodRef, 9, 1)) = gMonedaNacional Then
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefRefCapIntDMMorSoles
                                sConsCred2 = gCredRefRefCapIntIngDMDolares
                            Else
                                sConsCred = gCredRefRefCapIntMorSoles
                                sConsCred2 = gCredRefRefCapIntIngSoles
                            End If
                        Else
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefRefCapIntDMMorDolares
                                sConsCred2 = gCredRefRefCapIntIngDMSoles
                            Else
                                sConsCred = gCredRefRefCapIntMorDolares
                                sConsCred2 = gCredRefRefCapIntIngDolares
                            End If
                        End If
                    Case gColocEstRefNorm
                        If CInt(Mid(R!cCtaCodRef, 9, 1)) = gMonedaNacional Then
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefRefCapIntDMNormSoles
                                sConsCred2 = gCredRefRefCapIntIngDMDolares
                            Else
                                sConsCred = gCredRefRefCapIntNormSoles
                                sConsCred2 = gCredRefRefCapIntIngSoles
                            End If
                        Else
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefRefCapIntDMNormDolares
                                sConsCred2 = gCredRefRefCapIntIngDMSoles
                            Else
                                sConsCred = gCredRefRefCapIntNormDolares
                                sConsCred2 = gCredRefRefCapIntIngDolares
                            End If
                        End If
                    Case gColocEstRefVenc
                        If CInt(Mid(R!cCtaCodRef, 9, 1)) = gMonedaNacional Then
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefRefCapIntDMVencSoles
                                sConsCred2 = gCredRefRefCapIntIngDMDolares
                            Else
                                sConsCred = gCredRefRefCapIntVencSoles
                                sConsCred2 = gCredRefRefCapIntIngSoles
                            End If
                        Else
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefRefCapIntDMVencDolares
                                sConsCred2 = gCredRefRefCapIntIngDMSoles
                            Else
                                sConsCred = gCredRefRefCapIntVencDolares
                                sConsCred2 = gCredRefRefCapIntIngDMDolares
                            End If
                        End If
                    Case gColocEstVigMor
                        If CInt(Mid(R!cCtaCodRef, 9, 1)) = gMonedaNacional Then
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefCapIntDMMorSoles
                                sConsCred2 = gCredRefCapIntIngDMDolares
                            Else
                                sConsCred = gCredRefCapIntMorSoles
                                sConsCred2 = gCredRefCapIntIngSoles
                            End If
                        Else
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefCapIntDMMorDolares
                                sConsCred2 = gCredRefCapIntIngDMSoles
                            Else
                                sConsCred = gCredRefCapIntMorDolares
                                sConsCred2 = gCredRefCapIntIngDolares
                            End If
                        End If
                    Case gColocEstVigNorm
                        If CInt(Mid(R!cCtaCodRef, 9, 1)) = gMonedaNacional Then
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefCapIntDMNormSoles
                                sConsCred2 = gCredRefCapIntIngDMDolares
                            Else
                                sConsCred = gCredRefCapIntNormSoles
                                sConsCred2 = gCredRefCapIntIngSoles
                            End If
                        Else
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefCapIntDMNormDolares
                                sConsCred2 = gCredRefCapIntIngDMSoles
                            Else
                                sConsCred = gCredRefCapIntNormDolares
                                sConsCred2 = gCredRefCapIntIngDolares
                            End If
                        End If
                    Case gColocEstVigVenc
                        If CInt(Mid(R!cCtaCodRef, 9, 1)) = gMonedaNacional Then
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefCapIntDMVencSoles
                                sConsCred2 = gCredRefCapIntIngDMDolares
                            Else
                                sConsCred = gCredRefCapIntVencSoles
                                sConsCred2 = gCredRefCapIntIngSoles
                            End If
                        Else
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefCapIntDMVencDolares
                                sConsCred2 = gCredRefCapIntIngDMSoles
                            Else
                                sConsCred = gCredRefCapIntVencDolares
                                sConsCred2 = gCredRefCapIntIngDolares
                            End If
                        End If
                End Select
            Else
                Select Case nEstadCredRef
                    Case gColocEstRefMor
                        If CInt(Mid(R!cCtaCodRef, 9, 1)) = gMonedaNacional Then
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefRefDMMorSoles
                                sConsCred2 = gCredRefRefIngDMDolares
                            Else
                                sConsCred = gCredRefRefMorSoles
                                sConsCred2 = gCredRefRefIngSoles
                            End If
                        Else
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefRefDMMorDolares
                                sConsCred2 = gCredRefRefIngDMSoles
                            Else
                                sConsCred = gCredRefRefMorDolares
                                sConsCred2 = gCredRefRefIngDolares
                            End If
                        End If
                    Case gColocEstRefNorm
                        If CInt(Mid(R!cCtaCodRef, 9, 1)) = gMonedaNacional Then
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefRefDMNormSoles
                                sConsCred2 = gCredRefRefIngDMDolares
                            Else
                                sConsCred = gCredRefRefNormSoles
                                sConsCred2 = gCredRefRefIngSoles
                            End If
                        Else
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefRefDMNormDolares
                                sConsCred2 = gCredRefRefIngDMSoles
                            Else
                                sConsCred = gCredRefRefNormDolares
                                sConsCred2 = gCredRefRefIngDolares
                            End If
                        End If
                    Case gColocEstRefVenc
                        If CInt(Mid(R!cCtaCodRef, 9, 1)) = gMonedaNacional Then
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefRefDMVencSoles
                                sConsCred2 = gCredRefRefIngDMDolares
                            Else
                                sConsCred = gCredRefRefVencSoles
                                sConsCred2 = gCredRefRefIngSoles
                            End If
                        Else
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefRefDMVencDolares
                                sConsCred2 = gCredRefRefIngDMSoles
                            Else
                                sConsCred = gCredRefRefVencDolares
                                sConsCred2 = gCredRefRefIngDolares
                            End If
                        End If
                    Case gColocEstVigMor
                        If CInt(Mid(R!cCtaCodRef, 9, 1)) = gMonedaNacional Then
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefDMMorSoles
                                sConsCred2 = gCredRefIngDMDolares
                            Else
                                sConsCred = gCredRefMorSoles
                                sConsCred2 = gCredRefIngSoles
                            End If
                        Else
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefDMMorDolares
                                sConsCred2 = gCredRefIngDMSoles
                            Else
                                sConsCred = gCredRefMorDolares
                                sConsCred2 = gCredRefIngDolares
                            End If
                        End If
                    Case gColocEstVigNorm
                        If CInt(Mid(R!cCtaCodRef, 9, 1)) = gMonedaNacional Then
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefDMNormSoles
                                sConsCred2 = gCredRefIngDMDolares
                            Else
                                sConsCred = gCredRefNormSoles
                                sConsCred2 = gCredRefIngSoles
                            End If
                        Else
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefDMNormDolares
                                sConsCred2 = gCredRefIngDMSoles
                            Else
                                sConsCred = gCredRefNormDolares
                                sConsCred2 = gCredRefIngDolares
                            End If
                        End If
                    Case gColocEstVigVenc
                        If CInt(Mid(R!cCtaCodRef, 9, 1)) = gMonedaNacional Then
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefDMVencSoles
                                sConsCred2 = gCredRefIngDMDolares
                            Else
                                sConsCred = gCredRefVencSoles
                                sConsCred2 = gCredRefIngSoles
                            End If
                        Else
                            If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                                sConsCred = gCredRefDMVencDolares
                                sConsCred2 = gCredRefIngDMSoles
                            Else
                                sConsCred = gCredRefVencDolares
                                sConsCred2 = gCredRefIngDolares
                            End If
                        End If
                End Select
            End If
            
            'Movimiento para Credito Antiguo
            If pbCapitalInter Then
                nMontoRefTmp = MatrizInteresGastosAFecha(R!cCtaCodRef, MatCredRef, pdEstado) + MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)
                nMontoRefTmpCapital = MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)
                nMontoRefTmpInteres = MatrizInteresGastosAFecha(R!cCtaCodRef, MatCredRef, pdEstado) - MatrizGastosFecha(R!cCtaCodRef, MatCredRef)
                nMontoRefTmpGastosSusp = MatrizGastosFecha(R!cCtaCodRef, MatCredRef)
                nMontoRefTmpInterSusp = nMontoRefTmpInteres
                'Insertar Movimientos
                Call oDCredito.dInsertMovCol(nMovNro, sConsCred, R!cCtaCodRef, 0, nMontoRefTmp, 0, 0, 0, False, gColocEstAprob)
                Call oDCredito.dInsertMovColDet(nMovNro, sConsCred, R!cCtaCodRef, 0, gColocConceptoCodCapital, 0, nMontoRefTmpCapital, False)
                If nMontoRefTmpInteres > 0 Then
                    Call oDCredito.dInsertMovColDet(nMovNro, sConsCred, R!cCtaCodRef, 0, gColocConceptoCodInteresSuspenso, 0, nMontoRefTmpInteres, False)
                End If
                
                'Inserta Movimiento para Gastos Capitalizados al Detalle
                If nMontoRefTmpGastosSusp > 0 Then
                    'Call oDCredito.dInsertMovCol(nMovNro, gCredRefGastosSuspenso, R!cCtaCodRef, 0, nMontoRefTmpGastosSusp, 0, 0, 0, False, gColocEstAprob)
                    'Call oDCredito.dInsertMovColDet(nMovNro, gCredRefGastosSuspenso, R!cCtaCodRef, 0, gColocConceptoCodGastoSuspenso, 0, nMontoRefTmpGastosSusp, False)
                    k = -1
                    For i = 0 To UBound(MatrizGastosCredRef) - 1
                        MatrizGastosCredRefTempo = MatrizGastosCredRef(i)
                        If UBound(MatrizGastosCredRefTempo) > 0 Then
                            If MatrizGastosCredRefTempo(0, 3) = R!cCtaCodRef Then
                                k = i
                                Exit For
                            End If
                        End If
                    Next i
                    
                    If k <> -1 Then
                        MatrizGastosCredRefTempo = MatrizGastosCredRef(k)
                        For i = 0 To UBound(MatrizGastosCredRefTempo) - 1
                            Call oDCredito.dInsertMovColDet(nMovNro, sConsCred, R!cCtaCodRef, 0, CLng(MatrizGastosCredRefTempo(i, 1)), CLng(MatrizGastosCredRefTempo(i, 0)), CDbl(MatrizGastosCredRefTempo(i, 2)), False)
                        Next i
                    End If
                    
                    nGastoSuspTot = nGastoSuspTot + nMontoRefTmpGastosSusp
                    
                End If
                
            Else
                'Solo capital, Intereses y Gastos en suspenso son de otro Movimiento
                nMontoRefTmp = MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)
                nMontoRefTmpCapital = MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)
                nMontoRefTmpInterSusp = MatrizInteresGastosAFecha(R!cCtaCodRef, MatCredRef, pdEstado) - MatrizGastosFecha(R!cCtaCodRef, MatCredRef)
                nMontoRefTmpGastosSusp = MatrizGastosFecha(R!cCtaCodRef, MatCredRef)
                
                'Insertar Movimientos
                Call oDCredito.dInsertMovCol(nMovNro, sConsCred, R!cCtaCodRef, 0, nMontoRefTmp, 0, 0, 0, False, gColocEstAprob)
                Call oDCredito.dInsertMovColDet(nMovNro, sConsCred, R!cCtaCodRef, 0, gColocConceptoCodCapital, 0, nMontoRefTmpCapital, False)
                
                'Insertar Movimientos Interes en suspenso
                If nMontoRefTmpInterSusp > 0 Then
                    'Actualiza Calendario
                    nIntSuspTot = nIntSuspTot + nMontoRefTmpInterSusp
                    'Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, UBound(MatCalendPagos), gColocConceptoCodInteresSuspenso, nMontoRefTmpInterSusp, 0, "", False)
                End If
                If nMontoRefTmpGastosSusp > 0 Then
                    nGastoSuspTot = nGastoSuspTot + nMontoRefTmpGastosSusp
                    'Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, UBound(MatCalendPagos), gColocConceptoCodGastoVarios, nMontoRefTmpGastosSusp, 0, "", False)
                End If
            End If
            
            'Insertar Movimientos Intereses en suspenso
            If nMontoRefTmpInterSusp > 0 Then
                Call oDCredito.dInsertMovCol(nMovNro, gCredRefIntSuspenso, R!cCtaCodRef, 0, nMontoRefTmpInterSusp, 0, 0, 0, False, gColocEstAprob)
                Call oDCredito.dInsertMovColDet(nMovNro, gCredRefIntSuspenso, R!cCtaCodRef, 0, gColocConceptoCodInteresSuspenso, 0, nMontoRefTmpInterSusp, False)
            End If
            
            'Insertar Movimientos Gastos en suspenso
            If nMontoRefTmpGastosSusp > 0 Then
                Call oDCredito.dInsertMovCol(nMovNro, gCredRefGastosSuspenso, R!cCtaCodRef, 0, nMontoRefTmpGastosSusp, 0, 0, 0, False, gColocEstAprob)
                'Call oDCredito.dInsertMovColDet(nMovNro, gCredRefGastosSuspenso, R!cCtaCodRef, 0, gColocConceptoCodGastoSuspenso, 0, nMontoRefTmpGastosSusp, False)
                
                k = -1
                For i = 0 To UBound(MatrizGastosCredRef) - 1
                    MatrizGastosCredRefTempo = MatrizGastosCredRef(i)
                    If UBound(MatrizGastosCredRefTempo) > 0 Then
                        If MatrizGastosCredRefTempo(0, 3) = R!cCtaCodRef Then
                            k = i
                            Exit For
                        End If
                    End If
                Next i
                
                If k <> -1 Then
                    MatrizGastosCredRefTempo = MatrizGastosCredRef(k)
                    For i = 0 To UBound(MatrizGastosCredRefTempo) - 1
                        Call oDCredito.dInsertMovColDet(nMovNro, gCredRefGastosSuspenso, R!cCtaCodRef, 0, CLng(MatrizGastosCredRefTempo(i, 1)), CLng(MatrizGastosCredRefTempo(i, 0)), CDbl(MatrizGastosCredRefTempo(i, 2)), False)
                    Next i
                End If
                
            End If
                
            'Movimiento para Credito Nuevo
            If pbCapitalInter Then
                nMontoRefTmp = MatrizInteresGastosAFecha(R!cCtaCodRef, MatCredRef, pdEstado) + MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)
                nMontoRefTmpCapital = MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)
                nMontoRefTmpInteres = MatrizInteresGastosAFecha(R!cCtaCodRef, MatCredRef, pdEstado) - MatrizGastosFecha(R!cCtaCodRef, MatCredRef)
                'Si es Diferente Moneda
                If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                    Set oGeneral = New DGeneral
                    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdEstado, TCFijoMes)
                    Set oGeneral = Nothing
                    If CInt(Mid(psCtaCod, 9, 1)) = gMonedaNacional Then 'de Dolares a Soles
                        nMontoRefTmp = nMontoRefTmp * nTipoCambioFijo
                        nMontoRefTmp = CDbl(Format(nMontoRefTmp, "#0.00"))
                        nMontoRefTmpCapital = nMontoRefTmpCapital * nTipoCambioFijo
                        nMontoRefTmpCapital = CDbl(Format(nMontoRefTmpCapital, "#0.00"))
                        nMontoRefTmpInteres = nMontoRefTmpInteres * nTipoCambioFijo
                        nMontoRefTmpInteres = CDbl(Format(nMontoRefTmpInteres, "#0.00"))
                    Else 'De Soles a Dolares
                        nMontoRefTmp = nMontoRefTmp / nTipoCambioFijo
                        nMontoRefTmp = CDbl(Format(nMontoRefTmp, "#0.00"))
                        nMontoRefTmpCapital = nMontoRefTmpCapital / nTipoCambioFijo
                        nMontoRefTmpCapital = CDbl(Format(nMontoRefTmpCapital, "#0.00"))
                        nMontoRefTmpInteres = nMontoRefTmpInteres / nTipoCambioFijo
                        nMontoRefTmpInteres = CDbl(Format(nMontoRefTmpInteres, "#0.00"))
                    End If
                End If
                
                nMontoRefTmpNew = nMontoRefTmpNew + nMontoRefTmp
                nMontoRefTmpNew = CDbl(Format(nMontoRefTmpNew, "#0.00"))
                nMontoRefTmpCapitalNew = nMontoRefTmpCapitalNew + nMontoRefTmpCapital
                nMontoRefTmpCapitalNew = CDbl(Format(nMontoRefTmpCapitalNew, "#0.00"))
                nMontoRefTmpInteresNew = nMontoRefTmpInteresNew + nMontoRefTmpInteres
                nMontoRefTmpInteresNew = CDbl(Format(nMontoRefTmpInteresNew, "#0.00"))
                
            Else
                nMontoRefTmp = MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)
                nMontoRefTmpCapital = MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)
                'Si es Diferente Moneda
                If Mid(psCtaCod, 9, 1) <> Mid(R!cCtaCodRef, 9, 1) Then
                    Set oGeneral = New DGeneral
                    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdEstado, TCFijoMes)
                    Set oGeneral = Nothing
                    If CInt(Mid(psCtaCod, 9, 1)) = gMonedaNacional Then 'de Dolares a Soles
                        nMontoRefTmp = nMontoRefTmp * nTipoCambioFijo
                        nMontoRefTmp = CDbl(Format(nMontoRefTmp, "#0.00"))
                        nMontoRefTmpCapital = nMontoRefTmpCapital * nTipoCambioFijo
                        nMontoRefTmpCapital = CDbl(Format(nMontoRefTmpCapital, "#0.00"))
                    Else 'De Soles a Dolares
                        nMontoRefTmp = nMontoRefTmp / nTipoCambioFijo
                        nMontoRefTmp = CDbl(Format(nMontoRefTmp, "#0.00"))
                        nMontoRefTmpCapital = nMontoRefTmpCapital / nTipoCambioFijo
                        nMontoRefTmpCapital = CDbl(Format(nMontoRefTmpCapital, "#0.00"))
                    End If
                End If
                
                nMontoRefTmpNew = nMontoRefTmpNew + nMontoRefTmp
                nMontoRefTmpNew = CDbl(Format(nMontoRefTmpNew, "#0.00"))
                nMontoRefTmpCapitalNew = nMontoRefTmpCapitalNew + nMontoRefTmpCapital
                nMontoRefTmpCapitalNew = CDbl(Format(nMontoRefTmpCapitalNew, "#0.00"))
            End If
            
            
            'Actualiza Estado
            If R!bSustiDeudor Then
               nEstado = gColocEstCancelado
            Else
                nEstado = gColocEstRefinanc
            End If
            Call oDCredito.dUpdateProducto_Estado(R!cCtaCodRef, nEstado, pdEstado, , False, True)
            
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing

        'Graba Movimientos de Credito Nuevo Totalizado de Todos los Creditos Antiguos
        If nGastoSuspTot > 0 And Not pbCapitalInter Then
            For i = 0 To UBound(MatrizGastosCredGrp) - 1
                Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, UBound(MatCalendPagos), CLng(MatrizGastosCredGrp(i, 0)), CDbl(MatrizGastosCredGrp(i, 1)), 0, "", False)
            Next i
        End If
        If nIntSuspTot > 0 And Not pbCapitalInter Then
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, UBound(MatCalendPagos), gColocConceptoCodInteresSuspenso, nIntSuspTot, 0, "", False)
        End If
        
        Set oCredito = Nothing
        
        'Ingresar Movimientos para Nuevo Credito
        If pbCapitalInter Then
            'Inserta Movimientos
            Call oDCredito.dInsertMovCol(nMovNro, sConsCred2, psCtaCod, 0, nMontoRefTmpNew, 0, 0, 0, False, gColocEstAprob)
            Call oDCredito.dInsertMovColDet(nMovNro, sConsCred2, psCtaCod, 0, gColocConceptoCodCapital, 0, nMontoRefTmpCapitalNew, False)
            If nMontoRefTmpInteresNew > 0 Then
                Call oDCredito.dInsertMovColDet(nMovNro, sConsCred2, psCtaCod, 0, gColocConceptoCodInteresSuspenso, 0, nMontoRefTmpInteresNew, False)
            End If
            If nGastoSuspTot > 0 Then
                For i = 0 To UBound(MatrizGastosCredGrp) - 1
                    Call oDCredito.dInsertMovColDet(nMovNro, sConsCred2, psCtaCod, 0, CLng(MatrizGastosCredGrp(i, 0)), 0, CDbl(MatrizGastosCredGrp(i, 1)), False)
                Next i
            End If
        Else
            'Insertar Movimientos
            Call oDCredito.dInsertMovCol(nMovNro, sConsCred2, psCtaCod, 0, nMontoRefTmpNew, 0, 0, 0, False, gColocEstAprob)
            Call oDCredito.dInsertMovColDet(nMovNro, sConsCred2, psCtaCod, 0, gColocConceptoCodCapital, 0, nMontoRefTmpCapitalNew, False)
        End If
        
        'Actualiza Estado de Producto
        Call oDCredito.dUpdateProducto(psCtaCod, , pnMonto, gColocEstRefNorm, pdEstado, 0, False)
        'Actualiza Colocaciones
        Call oDCredito.dUpdateColocaciones(psCtaCod, , , pnMonto, , , pdEstado, False)
        
    Else
        'Call oDCredito.dInsertMov(sMovNro, gCredAprobacion, "Aprobacion de Credito", gMovEstContabNoContable, gMovFlagVigente, False)
        'nMovNro = oDCredito.dGetnMovNro(sMovNro)
        'Call oDCredito.dInsertMovCol(nMovNro, gCredAprobacion, psCtaCod, 0, pnMonto, 0, "", 0, False, gColocEstAprob)
    End If
    
    
    If pnMiViv = 1 Then
        Call oDCredito.dDeleteColocCalendDet(psCtaCod, pnNroCalenPar, gColocCalendAplCuota, , , False)
        Call oDCredito.dDeleteColocCalendario(psCtaCod, pnNroCalenPar, gColocCalendAplCuota, , False)
        For i = 0 To UBound(MatCalendPagos_2) - 1
            Call oDCredito.dInsertColocCalendario(psCtaCod, pnNroCalenPar, gColocCalendAplCuota, i + 1, CDate(MatCalendPagos_2(i, 0)), gColocCalendEstadoPendiente, "Calendario de Sugerencia Pagos", gColocCalendConceptoProcSugerido, False)
            If pnTipoCuota <> gColocCalendCodCL Then
                Call oDCredito.dInsertColocCalendDet(psCtaCod, pnNroCalenPar, gColocCalendAplCuota, i + 1, gColocConceptoCodCapital, CDbl(MatCalendPagos_2(i, 3)), 0, "", False)
                Call oDCredito.dInsertColocCalendDet(psCtaCod, pnNroCalenPar, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalendPagos_2(i, 4)), 0, "", False)
                Call oDCredito.dInsertColocCalendDet(psCtaCod, pnNroCalenPar, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", False)
                If CDbl(MatCalendPagos_2(i, 5)) > 0 Then
                    Call oDCredito.dInsertColocCalendDet(psCtaCod, pnNroCalenPar, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresGracia, CDbl(MatCalendPagos_2(i, 5)), 0, "", False)
                End If
            End If
        Next i
    End If
    
    Call ActualizaCreditosVigentes(MatCredVig)
    
    oDCredito.dCommitTrans
    Set oDCredito = Nothing
    
    Exit Function

ErrorSugerenciaCredito:
    If bTransac Then
        oDCredito.dRollbackTrans
        Set oDCredito = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso SugerenciaCredito", Err.Description
    
End Function

'Public Function InteresGastosAFecha(ByVal psCtaCod As String, ByVal pdHoy As Date, ByVal pnMonto As Double, ByVal pdVig As Date, ByVal pnTasa As Double) As Double
'Dim R As ADODB.Recordset
'Dim oCalend As Dcalendario
'Dim nMontoGanado As Double
'Dim nMontoAFecha As Double
'Dim dFecIni As Date
'Dim nSaldoCap As Double
'Dim nSaldoCal As Double
'
'    'Calculo de Montos Ya Ganados
'    Set oCalend = New Dcalendario
'    Set R = oCalend.RecuperaCalendarioPagosDeuda(psCtaCod)
'    nMontoGanado = 0
'    nSaldoCap = 0
'    Do While Not R.EOF
'        nMontoGanado = nMontoGanado + CDbl(Format(IIf(IsNull(R!nIntGracia), 0, R!nIntGracia), "#0.00"))
'        nMontoGanado = nMontoGanado + CDbl(Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00"))
'        nMontoGanado = nMontoGanado + CDbl(Format(IIf(IsNull(R!nIntSuspenso), 0, R!nIntSuspenso), "#0.00"))
'        nMontoGanado = nMontoGanado + CDbl(Format(IIf(IsNull(R!nIntReprog), 0, R!nIntReprog), "#0.00"))
'        nMontoGanado = nMontoGanado + CDbl(Format(IIf(IsNull(R!NGasto), 0, R!NGasto), "#0.00"))
'        If pdHoy >= R!dvenc Then
'            nMontoGanado = nMontoGanado + CDbl(Format(R!nIntComp, "#0.00"))
'        End If
'        nSaldoCap = nSaldoCap + CDbl(Format(R!nCapital, "#0.00"))
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'
'    dFecIni = pdVig
'    nSaldoCal = pnMonto
'
'    'Interes a la fecha
'    nMontoAFecha = 0
'    Set R = oCalend.RecuperaCalendarioPagos(psCtaCod)
'    Do While Not R.EOF
'        If pdHoy >= R!dvenc Then
'            nSaldoCal = nSaldoCal - R!nCapital
'            nSaldoCal = CDbl(Format(nSaldoCal, "#0.00"))
'            dFecIni = R!dvenc
'        Else
'            nMontoAFecha = TasaIntPerDias(pnTasa, pdHoy - dFecIni) * nSaldoCal
'        End If
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'    Set oCalend = Nothing
'
'    'Total Calculado es
'    InteresGastosAFecha = CDbl(Format(nMontoGanado + nMontoAFecha, "#0.00"))
'
'End Function
Public Function InteresGastosAFecha(ByVal psCtaCod As String, ByVal pdHoy As Date) As Double
Dim MatCalend As Variant

    'Calculo de Montos Ya Ganados
    MatCalend = RecuperaMatrizCalendarioPendiente(psCtaCod)
    'Total Calculado es
    InteresGastosAFecha = CDbl(Format(MatrizInteresGastosAFecha(psCtaCod, MatCalend, pdHoy), "#0.00"))
    
End Function

Public Function MatrizGastosFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 9)) 'Gastos
    Next i
    
    'Total Calculado es
    MatrizGastosFecha = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function


Public Function MatrizInteresReprogramadoFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 7)) 'Interes Reprogramado
    Next i
    
    'Total Calculado es
    MatrizInteresReprogramadoFecha = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function

Public Function MatrizInteresCompensatorioVencido(ByVal MatCalend As Variant) As Double
Dim nMontoGanado As Double
Dim i As Integer

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 11)) 'Interes Suspenso
    Next i
    
    'Total Calculado es
    MatrizInteresCompensatorioVencido = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function

Public Function MatrizInteresSuspensoFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 8)) 'Interes Suspenso
    Next i
    
    'Total Calculado es
    MatrizInteresSuspensoFecha = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function


Public Function MatrizInteresGraciaFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date, _
    Optional pnMontoCol As Double = -1) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim i As Integer
Dim nSaldoCal As Double
Dim pnTasa As Double
Dim nDiasGracia As Integer
Dim oCalend As Dcalendario

    If pnMontoCol = -1 Then
        Set oCredito = New DCredito
        Set R = oCredito.RecuperaColocaciones(psCtaCod)
        Set oCredito = Nothing
        nSaldoCal = R!nMontoCol
        R.Close
        Set R = Nothing
    Else
        nSaldoCal = pnMontoCol
    End If
    
    Set oCalend = New Dcalendario
    dFecIni = oCalend.RecuperaFechaInicioCuota(psCtaCod, CInt(MatCalend(0, 1)), gColocCalendAplCuota)
    Set oCalend = Nothing
    
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaProductoTasaInteres(psCtaCod, gColocLineaCredTasasIntGracia)
    Set oCredito = Nothing
    If Not R.BOF And Not R.EOF Then
        pnTasa = R!nTasaInteres
    Else
        MatrizInteresGraciaFecha = 0
        R.Close
        Set R = Nothing
        Exit Function
    End If
    R.Close
    Set R = Nothing
    
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
    Set oCredito = Nothing
    nDiasGracia = R!nPeriodoGracia
    R.Close
    Set R = Nothing
    
    nMontoGanado = 0
    'Si es Cierto ya Vencio la Fecha de Gracia
    If Abs(CDate(Format(pdHoy, "dd/mm/yyyy")) - CDate(Format(dFecIni, "dd/mm/yyyy"))) >= nDiasGracia Then
        'Calculo de Montos Ya Ganados
        For i = 0 To UBound(MatCalend) - 1
            nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 5)) 'Interes Moratorio
        Next i
    Else
        nMontoGanado = TasaIntPerDias(pnTasa, CDate(Format(pdHoy, "dd/mm/yyyy")) - CDate(Format(dFecIni, "dd/mm/yyyy"))) * nSaldoCal
    End If
    'Total Calculado es
    MatrizInteresGraciaFecha = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function



Public Function MatrizCapitalAFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant) As Double
Dim nMontoGanado As Double
Dim i As Integer
    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 3)) 'Capital
    Next i
    
    'Total Calculado es
    MatrizCapitalAFecha = CDbl(Format(nMontoGanado, "#0.00"))

End Function
Private Sub MatrizActualizarEstadoCalendCancelado(ByRef MatCalendDistrib As Variant, Optional ByVal pbVerficaSolosCapital As Boolean = False, _
    Optional ByVal pbNoCancelar As Boolean = True, Optional ByVal MatCalend As Variant = "")
Dim i As Integer

    For i = 0 To UBound(MatCalendDistrib) - 1
        If pbNoCancelar Then
            
            If pbVerficaSolosCapital Then
                If CInt(MatCalendDistrib(i, 2)) = gColocCalendEstadoPendiente Then
                    If CDbl(MatCalend(i, 3)) = CDbl(MatCalendDistrib(i, 3)) Then
                        MatCalendDistrib(i, 2) = Trim(Str(gColocCalendEstadoPagado))
                    Else
                        MatCalendDistrib(i, 2) = Trim(Str(gColocCalendEstadoPendiente))
                    End If
                End If
            Else
                If CDbl(MatCalend(i, 8)) = CDbl(MatCalendDistrib(i, 8)) And _
                    CDbl(MatCalend(i, 7)) = CDbl(MatCalendDistrib(i, 7)) And _
                    CDbl(MatCalend(i, 6)) = CDbl(MatCalendDistrib(i, 6)) And _
                    CDbl(MatCalend(i, 5)) = CDbl(MatCalendDistrib(i, 5)) And _
                    CDbl(MatCalend(i, 3)) = CDbl(MatCalendDistrib(i, 3)) Then
                    MatCalendDistrib(i, 2) = Trim(Str(gColocCalendEstadoPagado))
                End If
            End If
        Else
            MatCalendDistrib(i, 2) = Trim(Str(gColocCalendEstadoPagado))
        End If
    Next i
End Sub
Public Function MatrizInteresMorFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 6)) 'Interes Moratorio
    Next i
    
    'Total Calculado es
    MatrizInteresMorFecha = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function

Public Function MatrizInteresCompVencidoFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 11)) 'Interes Moratorio
    Next i
    
    'Total Calculado es
    MatrizInteresCompVencidoFecha = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function

Public Function MatrizCapitalCalendario(ByVal MatCalend As Variant) As Double
Dim i As Integer
    MatrizCapitalCalendario = 0
    For i = 0 To UBound(MatCalend) - 1
        MatrizCapitalCalendario = MatrizCapitalCalendario + CDbl(MatCalend(i, 3))
    Next i
End Function

Public Function CalculaGastoPenalidadCancelacion(ByVal pnSaldoCap As Double, ByVal pnMoneda As Integer) As Double
Dim oGasto As DGastos
Dim R As ADODB.Recordset
Dim nValor As Double
    Set oGasto = New DGastos
    nValor = oGasto.RecuperaGastoValorPenalidadCancelacion(pnMoneda)
    Set oGasto = Nothing
    nValor = CDbl(Format(nValor / 100, "#0.00"))
    CalculaGastoPenalidadCancelacion = CDbl(Format(nValor * pnSaldoCap, "#0.00"))
End Function

Public Function MatrizIntCompCalendario(ByVal MatCalend As Variant) As Double
Dim i As Integer
    MatrizIntCompCalendario = 0
    For i = 0 To UBound(MatCalend) - 1
        MatrizIntCompCalendario = MatrizIntCompCalendario + CDbl(MatCalend(i, 4))
    Next i
End Function

Public Function MatrizIntGraciaCalendario(ByVal MatCalend As Variant) As Double
Dim i As Integer
    MatrizIntGraciaCalendario = 0
    For i = 0 To UBound(MatCalend) - 1
        MatrizIntGraciaCalendario = MatrizIntGraciaCalendario + CDbl(MatCalend(i, 5))
    Next i
End Function

Public Function MatrizIntComVencCalendario(ByVal MatCalend As Variant) As Double
Dim i As Integer
    MatrizIntComVencCalendario = 0
    For i = 0 To UBound(MatCalend) - 1
        MatrizIntComVencCalendario = MatrizIntComVencCalendario + CDbl(MatCalend(i, 11))
    Next i
End Function

Public Function MatrizIntMoratorioCalendario(ByVal MatCalend As Variant) As Double
Dim i As Integer
    MatrizIntMoratorioCalendario = 0
    For i = 0 To UBound(MatCalend) - 1
        MatrizIntMoratorioCalendario = MatrizIntMoratorioCalendario + CDbl(MatCalend(i, 6))
    Next i
End Function

Public Function MatrizIntReprogCalendario(ByVal MatCalend As Variant) As Double
Dim i As Integer
    MatrizIntReprogCalendario = 0
    For i = 0 To UBound(MatCalend) - 1
        MatrizIntReprogCalendario = MatrizIntReprogCalendario + CDbl(MatCalend(i, 7))
    Next i
End Function

Public Function MatrizIntSuspensoCalendario(ByVal MatCalend As Variant) As Double
Dim i As Integer
    MatrizIntSuspensoCalendario = 0
    For i = 0 To UBound(MatCalend) - 1
        MatrizIntSuspensoCalendario = MatrizIntSuspensoCalendario + CDbl(MatCalend(i, 8))
    Next i
End Function

Public Function MatrizIntGastosCalendario(ByVal MatCalend As Variant) As Double
Dim i As Integer
    MatrizIntGastosCalendario = 0
    For i = 0 To UBound(MatCalend) - 1
        MatrizIntGastosCalendario = MatrizIntGastosCalendario + CDbl(MatCalend(i, 9))
    Next i
End Function

Public Function MatrizInteresCompAFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date, _
    Optional pnMontoCol As Double = -1, Optional pnTasaInteres As Double = -1, Optional ByVal pbDesagio As Boolean = False) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim oCalend As Dcalendario
Dim nMontoGanado As Double
Dim nMontoAFecha As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double
Dim nColocCalendCod As Integer
Dim sSQL As String
Dim oConec As DConecta

    'Calculo de Montos Ya Ganados


    Set oCredito = New DCredito
    Set R = oCredito.RecuperaColocacCred(psCtaCod)
    Set oCredito = Nothing
    nColocCalendCod = R!nColocCalendCod
    R.Close



    nMontoGanado = 0
    nSaldoCap = 0
    nMontoAFecha = 0

    If nColocCalendCod <> 70 Then
        For i = 0 To UBound(MatCalend) - 1
            If pdHoy >= CDate(MatCalend(i, 0)) Then
                nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 4))
            End If
        Next i
    Else
        nMontoGanado = nMontoGanado + CDbl(MatCalend(0, 4))
    End If

    If nColocCalendCod <> 70 Then
        If pnMontoCol = -1 Then
            Set oCredito = New DCredito
            Set R = oCredito.RecuperaColocaciones(psCtaCod)
            Set oCredito = Nothing
            nSaldoCal = R!nMontoCol
            R.Close
            Set R = Nothing
        Else
            nSaldoCal = pnMontoCol
        End If


        Set oCalend = New Dcalendario
        dFecIni = oCalend.RecuperaFechaInicioCuota(psCtaCod, CInt(MatCalend(0, 1)), gColocCalendAplCuota)
        Set oCalend = Nothing

        If pnTasaInteres = -1 Then
            Set oCredito = New DCredito
            Set R = oCredito.RecuperaProducto(psCtaCod)
            Set oCredito = Nothing
            pnTasa = R!nTasaInteres
            R.Close
            Set R = Nothing
        Else
            pnTasa = pnTasaInteres
        End If

        'Interes a la fecha
        nMontoAFecha = 0
        For i = 0 To UBound(MatCalend) - 1
            If pdHoy >= CDate(MatCalend(i, 0)) Then
                dFecIni = CDate(MatCalend(i, 0))
            Else
                If (pdHoy - dFecIni) > 0 Then
                    If MatCalend(i, 1) = "1" Then
                        nMontoAFecha = TasaIntPerDias(pnTasa, pdHoy - dFecIni) * nSaldoCal
                        'nMontoAFecha = TasaIntPerDias(pnTasa, pdHoy -dFecIni) * (CDbl(MatCalend(I, 10)) + CDbl(MatCalend(I, 3)))
                    Else
                        nMontoAFecha = TasaIntPerDias(pnTasa, pdHoy - dFecIni) * CDbl(SaldoK(psCtaCod, pdHoy))
                        'nMontoAFecha = TasaIntPerDias(pnTasa, pdHoy - dFecIni) * CDbl(MatCalend(I, 10))
                    End If
                End If
                sSQL = "Select nMontoPagado From ColocCalendDet Where cCtaCod = '" & psCtaCod & "' AND nColocCalendApl = 1 "
                sSQL = sSQL & " AND nNroCalen =  (Select nNroCalen from ColocacCred Where cctaCod = '" & psCtaCod & "')"
                sSQL = sSQL & " AND nPrdConceptoCod = 1100 AND nCuota = " & MatCalend(i, 1)
                sSQL = "Select nMontoPagado From ColocCalendDet Where cCtaCod = '" & psCtaCod & "' AND nColocCalendApl = 1 "
                sSQL = sSQL & " AND nNroCalen =  (Select nNroCalen from ColocacCred Where cctaCod = '" & psCtaCod & "')"
                sSQL = sSQL & " AND nPrdConceptoCod = 1100 AND nCuota = " & MatCalend(i, 1)
                Set oConec = New DConecta
                oConec.AbreConexion
                Set R = oConec.CargaRecordSet(sSQL)
                
                nMontoAFecha = nMontoAFecha - R!nMontoPagado
                If Not pbDesagio And nMontoAFecha < 0 Then
                    nMontoAFecha = 0
                End If
                oConec.CierraConexion
                Set oConec = Nothing
                Exit For
            End If
        Next i
    End If
    'Total Calculado es
    MatrizInteresCompAFecha = CDbl(Format(nMontoGanado + nMontoAFecha, "#0.00"))

End Function
Public Function SaldoK(ByVal pcCtaCod As String, ByVal dPago As Date) As Currency
    Dim oConec As DConecta
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    'Calculo del SaldoCapital
'        sSQL = "Select Top 1 (Select Sum(nMonto) From ColocCalendDet CD Where  C.cCtaCod=CD.cCtaCod and"
'        sSQL = sSQL & " C.nNroCalen = CD.nNroCalen And CD.cCtaCod = C.cCtaCod"
'        sSQL = sSQL & " and CD.nPrdConceptoCod=1000 and CD.nColocCalendApl=0)-"
'        sSQL = sSQL & " (Select Sum(nMontoPagado) From ColocCalendDet CD Where  C.cCtaCod=CD.cCtaCod and"
'        sSQL = sSQL & " C.nNroCalen = CD.nNroCalen And CD.cCtaCod = C.cCtaCod"
'        sSQL = sSQL & " and CD.nPrdConceptoCod=1000 and CD.nColocCalendApl=1) as SK"
'        sSQL = sSQL & " From colocCalendario C"
'        sSQL = sSQL & " Where C.cCtaCod='" & pcCtaCod & "'"

    sSQL = "Select nSaldo as SK From Producto Where cCtaCod='" & pcCtaCod & "'"
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    
    If Not rs.EOF And Not rs.BOF Then
        SaldoK = rs!SK
    End If
    'En el Caso pague cuotas atrasada y cancele el credito
    sSQL = "Select isnull(Sum(nMonto),0) as Capital"
    sSQL = sSQL & " From ColocCalendDet a "
    sSQL = sSQL & " inner join ColocCalendario b on a.cctacod=b.cctacod and a.nnrocalen=b.nnrocalen and "
    sSQL = sSQL & " a.nColocCalendApl=b.nColocCalendApl and a.ncuota=b.nCuota "
    sSQL = sSQL & " Where a.nPrdConceptoCod=1000 and b.dVenc<='" & Format(CStr(dPago), "MM/dd/yyyy") & "' and b.nColocCalendEstado=0 "
    sSQL = sSQL & " and a.cctacod='" & pcCtaCod & "'"
    
    Set rs = New ADODB.Recordset
    Set rs = oConec.CargaRecordSet(sSQL)
    If Not rs.EOF And Not rs.BOF Then
        SaldoK = SaldoK - IIf(IsNull(rs!Capital), 0, rs!Capital)
    End If
    Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function MatrizInteresTotalesAFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaColocaciones(psCtaCod)
    Set oCredito = Nothing
    
    dFecIni = R!dVigencia
    nSaldoCal = R!nMontoCol
    R.Close
    Set R = Nothing
        
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaProducto(psCtaCod)
    Set oCredito = Nothing
    pnTasa = R!nTasaInteres
    R.Close
    Set R = Nothing
    
    'Total Calculado a la fecha
    MatrizInteresTotalesAFecha = CDbl(Format(MatrizInteresReprogramadoFecha(psCtaCod, MatCalend) + _
                MatrizInteresSuspensoFecha(psCtaCod, MatCalend) + _
                MatrizInteresGraciaFecha(psCtaCod, MatCalend, pdHoy, nSaldoCal) + _
                MatrizInteresMorFecha(psCtaCod, MatCalend) + _
                MatrizInteresCompAFecha(psCtaCod, MatCalend, pdHoy, nSaldoCal, pnTasa) + MatrizInteresCompensatorioVencido(MatCalend), "#0.00"))
End Function

Public Function MatrizInteresTotalesAFechaSinMora(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaColocaciones(psCtaCod)
    Set oCredito = Nothing
    
    dFecIni = R!dVigencia
    nSaldoCal = R!nMontoCol
    R.Close
    Set R = Nothing
        
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaProducto(psCtaCod)
    Set oCredito = Nothing
    pnTasa = R!nTasaInteres
    R.Close
    Set R = Nothing
    
    'Total Calculado a la fecha
    MatrizInteresTotalesAFechaSinMora = CDbl(Format(MatrizInteresReprogramadoFecha(psCtaCod, MatCalend) + _
                MatrizInteresSuspensoFecha(psCtaCod, MatCalend) + _
                MatrizInteresGraciaFecha(psCtaCod, MatCalend, pdHoy, nSaldoCal) + _
                MatrizInteresCompAFecha(psCtaCod, MatCalend, pdHoy, nSaldoCal, pnTasa) + MatrizInteresCompensatorioVencido(MatCalend), "#0.00"))
End Function


Public Function MatrizInteresGastosAFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date, Optional ByVal pbDesagio As Boolean = False) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaColocaciones(psCtaCod)
    Set oCredito = Nothing
    
    dFecIni = R!dVigencia
    nSaldoCal = R!nMontoCol
    R.Close
    Set R = Nothing
        
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaProducto(psCtaCod)
    Set oCredito = Nothing
    pnTasa = R!nTasaInteres
    R.Close
    Set R = Nothing
    
    'Total Calculado a la fecha
    MatrizInteresGastosAFecha = CDbl(Format(MatrizGastosFecha(psCtaCod, MatCalend) + _
                MatrizInteresReprogramadoFecha(psCtaCod, MatCalend) + _
                MatrizInteresSuspensoFecha(psCtaCod, MatCalend) + _
                MatrizInteresGraciaFecha(psCtaCod, MatCalend, pdHoy, nSaldoCal) + _
                MatrizInteresMorFecha(psCtaCod, MatCalend) + _
                MatrizInteresCompAFecha(psCtaCod, MatCalend, pdHoy, nSaldoCal, pnTasa, pbDesagio) + MatrizInteresCompensatorioVencido(MatCalend), "#0.00"))
    
End Function

'Actualiza la Tabla de Calificacion para Mi Vivienda con el ultimo calendario
Public Sub ActualizarCalificacionMIVivienda(ByVal psCtaCod As String)
Dim MatCalend As Variant
Dim i As Integer
Dim oBase As DCredActualizaBD
Dim k As Integer
Dim oCred As DCredito
Dim nNroCalen As Integer
Dim R As ADODB.Recordset

    Set oBase = New DCredActualizaBD
    MatCalend = RecuperaMatrizCalendarioPendiente(psCtaCod)
    
    oBase.dBeginTrans
    
    Set oCred = New DCredito
    k = oCred.UltimaCuotaCanceladaOriginal(psCtaCod)
    k = k + 1
    Call oBase.dEliminaCuotasPendientesCalifMiViv(psCtaCod)
    Set R = oCred.RecuperaColocacCred(psCtaCod)
    nNroCalen = R!nNroCalen
    Set oCred = Nothing
    
    For i = 0 To UBound(MatCalend) - 1
        Call oBase.dInsertColocCalifMiViv(psCtaCod, k, nNroCalen, MatCalend(i, 1), CDate(MatCalend(i, 0)), gColocCalendEstadoPendiente)
        k = k + 1
    Next i
    oBase.dCommitTrans
    Set oBase = Nothing
End Sub

Public Function ReprogramarCreditoenMemoriaTotalMiVivienda(ByVal psCtaCod As String, ByVal pdHoy As Date, _
    ByRef pMatTramo_2 As Variant, Optional ByVal pbMantieneCuota As Boolean = True) As Variant
    
Dim oCred As DCredito
Dim oCalend As NCalendario
Dim oDCalend As Dcalendario
Dim RAprob As ADODB.Recordset
Dim MatCalendPagos_2, MatCalendPagos As Variant
Dim R As ADODB.Recordset
Dim RColoc As ADODB.Recordset
Dim nInteresReprog As Double
Dim nMontoPrestamo As Double
Dim nCuotasPendientes As Integer
Dim nTipocalend As Integer
Dim oParam As DParametro
Dim nTramoNoConsPorcen, nTramoNoConsMonto As Double
Dim nTramoConsMonto As Double
Dim nCuotaTemp As Integer
Dim nCuotaObjetivo As Double
Dim nCuotaFinal As Integer
Dim nCalifBuenPagador As Integer
Dim RCred As ADODB.Recordset
Dim MatCal As Variant
Dim nCuotaObjetivoTemp As Double
Dim i As Integer

    ReDim MatCalend(0, 0)
    MatCal = RecuperaMatrizCalendarioPendiente(psCtaCod)
    nCuotaObjetivo = MatrizMontoCuota(MatCal, CInt(MatCal(UBound(MatCal) - 2, 1)))
    MatCal = RecuperaMatrizCalendarioPendiente(psCtaCod, True)
    nCuotaObjetivo = nCuotaObjetivo + MatrizMontoCuota(MatCal, CInt(MatCal(UBound(MatCal) - 2, 1)))
    
    Set oCred = New DCredito
    nCuotaFinal = oCred.RecuperaUltimaCuotaEvaluada(psCtaCod)
    Set RCred = oCred.RecuperaColocacCred(psCtaCod)
    nCalifBuenPagador = RCred!nCalPago
    Set RAprob = oCred.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
    Set RColoc = oCred.RecuperaColocaciones(psCtaCod)
    Set R = oCred.RecuperaProducto(psCtaCod)
    nMontoPrestamo = CDbl(Format(R!nSaldo, "#0.00"))
    
    Set oCred = Nothing
    nInteresReprog = 0#
    
    Set oDCalend = New Dcalendario
    nCuotasPendientes = oDCalend.CuotasPendientes(psCtaCod)
    Set oDCalend = Nothing
    
    If RAprob!nColocCalendCod = gColocCalendCodFFCF Or RAprob!nColocCalendCod = gColocCalendCodFFCFPG _
        Or RAprob!nColocCalendCod = gColocCalendCodPFCFPG Or RAprob!nColocCalendCod = gColocCalendCodPFCF Then
        nTipocalend = 1
    End If
    
    If RAprob!nColocCalendCod = gColocCalendCodFFCC Or RAprob!nColocCalendCod = gColocCalendCodFFCCPG _
        Or RAprob!nColocCalendCod = gColocCalendCodPFCCPG Or RAprob!nColocCalendCod = gColocCalendCodPFCC Then
        nTipocalend = 2
    End If
    
    If RAprob!nColocCalendCod = gColocCalendCodFFCD Or RAprob!nColocCalendCod = gColocCalendCodFFCDPG _
        Or RAprob!nColocCalendCod = gColocCalendCodPFCDPG Or RAprob!nColocCalendCod = gColocCalendCodPFCD Then
        nTipocalend = 3
    End If
    Set oCalend = New NCalendario
            
        'Si Mantiene la Cuota entonces Halla Nuevo Plazo
        If Not pbMantieneCuota Then
            nCuotasPendientes = CuotasAproximada(pdHoy, RColoc!DVENC, RAprob!nPlazo, RAprob!nPeriodoFechaFija)
        Else
            'Halla el plazo aproximado para mantener la cuota
            'nCuotasPendientes = PlazoDeCuotaFijaFechaFija(nCuotaObjetivo, R!nTasaInteres, _
            '    RAprob!nCuotas, nMontoPrestamo, 0, RColoc!dVigencia, RAprob!nPeriodoFechaFija, _
            '    RAprob!nPeriodoGracia, RAprob!nProxMes)
            
            For i = RAprob!nCuotas To 1 Step -1
                'Porcentaje Real sin dividir enter 100
                Set oParam = New DParametro
                nTramoNoConsPorcen = oParam.RecuperaValorParametro(gColocMiVivTramo)
                Set oParam = Nothing
                
                nCuotasPendientes = i
                
                nTramoNoConsMonto = Format((nTramoNoConsPorcen / 100) * nMontoPrestamo, "#0.00")
                nTramoConsMonto = Format(nMontoPrestamo - nTramoNoConsMonto, "#0.00")
                MatCalendPagos = oCalend.GeneraCalendario(nTramoConsMonto, R!nTasaInteres, nCuotasPendientes, _
                            IIf(IsNull(RAprob!nPlazo), 0, RAprob!nPlazo), pdHoy, nTipocalend, IIf(RAprob!nPeriodoFechaFija > 0, 2, 1), _
                            Exonerada, 0, RAprob!nPeriodoFechaFija, _
                            IIf(RAprob!nProxMes = 0, False, True), , True)
                '(nCuotasPendientes / 6) + 0.5 redondeo hacia arriba
                MatCalendPagos_2 = oCalend.GeneraCalendario(nTramoNoConsMonto, R!nTasaInteres, (nCuotasPendientes / 6) + 0.5, _
                            180, pdHoy, Fija, PeriodoFijo, _
                            Exonerada, 0, RAprob!nPeriodoFechaFija, _
                            IIf(RAprob!nProxMes = 0, False, True), , True)
                
                MatCalendPagos_2 = UnirMatricesMiViviendaReprogramado(MatCalendPagos, MatCalendPagos_2, nMontoPrestamo, nCalifBuenPagador, nCuotaFinal)
                nCuotaTemp = CDbl(MatCalendPagos_2(7, 2))
                MatCalendPagos_2 = DiferencialMatricesMiVivienda(MatCalendPagos, MatCalendPagos_2)
                
                nCuotaObjetivoTemp = MatrizMontoCuota(MatCalendPagos, CInt(MatCalendPagos(7, 1)))
                nCuotaObjetivoTemp = nCuotaObjetivoTemp + MatrizMontoCuota(MatCalendPagos_2, CInt(MatCalendPagos_2(7, 1)))
            
                If nCuotaObjetivoTemp > nCuotaObjetivo Then
                    Exit For
                End If
            
            Next i
            
            nCuotasPendientes = i + 1
            
            
        End If
        
        'Porcentaje Real sin dividir enter 100
        Set oParam = New DParametro
        nTramoNoConsPorcen = oParam.RecuperaValorParametro(gColocMiVivTramo)
        Set oParam = Nothing

        
        nTramoNoConsMonto = Format((nTramoNoConsPorcen / 100) * nMontoPrestamo, "#0.00")
        nTramoConsMonto = Format(nMontoPrestamo - nTramoNoConsMonto, "#0.00")
        MatCalendPagos = oCalend.GeneraCalendario(nTramoConsMonto, R!nTasaInteres, nCuotasPendientes, _
                    IIf(IsNull(RAprob!nPlazo), 0, RAprob!nPlazo), pdHoy, nTipocalend, IIf(RAprob!nPeriodoFechaFija > 0, 2, 1), _
                    Exonerada, 0, RAprob!nPeriodoFechaFija, _
                    IIf(RAprob!nProxMes = 0, False, True), , True)
                    
        MatCalendPagos_2 = oCalend.GeneraCalendario(nTramoNoConsMonto, R!nTasaInteres, nCuotasPendientes / 6, _
                    180, pdHoy, Fija, PeriodoFijo, _
                    Exonerada, 0, RAprob!nPeriodoFechaFija, _
                    IIf(RAprob!nProxMes = 0, False, True), , True)
        
        MatCalendPagos_2 = UnirMatricesMiViviendaReprogramado(MatCalendPagos, MatCalendPagos_2, nMontoPrestamo, nCalifBuenPagador, nCuotaFinal)
        nCuotaTemp = CDbl(MatCalendPagos_2(7, 2))
        MatCalendPagos_2 = DiferencialMatricesMiVivienda(MatCalendPagos, MatCalendPagos_2)
    
        nCuotaObjetivoTemp = MatrizMontoCuota(MatCalendPagos, CInt(MatCalendPagos(7, 1)))
        nCuotaObjetivoTemp = nCuotaObjetivoTemp + MatrizMontoCuota(MatCalendPagos_2, CInt(MatCalendPagos_2(7, 1)))


    Set oCalend = Nothing
    R.Close
    Set R = Nothing
    RAprob.Close
    Set RAprob = Nothing
    
    pMatTramo_2 = MatCalendPagos_2
    ReprogramarCreditoenMemoriaTotalMiVivienda = MatCalendPagos
    
End Function



Public Function ReprogramarCreditoenMemoriaTotal(ByVal psCtaCod As String, ByVal pdHoy As Date, _
    Optional ByVal pbMantieneCuota As Boolean = True) As Variant
    
Dim oCred As DCredito
Dim oCalend As NCalendario
Dim oDCalend As Dcalendario
Dim RAprob As ADODB.Recordset
Dim MatCalend As Variant
Dim R As ADODB.Recordset
Dim RColoc As ADODB.Recordset
Dim nInteresReprog As Double
Dim nMontoPrestamo As Double
Dim nCuotasPendientes As Integer
Dim nTipocalend As Integer
Dim nCuotaObjetivo As Double

    ReDim MatCalend(0, 0)
    Set oCred = New DCredito
    Set RAprob = oCred.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
    Set RColoc = oCred.RecuperaColocaciones(psCtaCod)
    Set R = oCred.RecuperaProducto(psCtaCod)
    nMontoPrestamo = CDbl(Format(R!nSaldo, "#0.00"))
    nCuotaObjetivo = oCred.CuotaAprobada(psCtaCod)
    Set oCred = Nothing
    'nInteresReprog = InteresGastosAFecha(pscatcod, pdHoy, nMontoPrestamo, R!dPrdestado, R!nTasaInteres) + (TasaIntPerDias(R!nTasaInteres, dFecReprog - pdHoy) * R!nSaldo)
    nInteresReprog = 0#
    
    Set oDCalend = New Dcalendario
    nCuotasPendientes = oDCalend.CuotasPendientes(psCtaCod)
    Set oDCalend = Nothing
    
    If RAprob!nColocCalendCod = gColocCalendCodFFCF Or RAprob!nColocCalendCod = gColocCalendCodFFCFPG _
        Or RAprob!nColocCalendCod = gColocCalendCodPFCFPG Or RAprob!nColocCalendCod = gColocCalendCodPFCF Then
        nTipocalend = 1
    End If
    
    If RAprob!nColocCalendCod = gColocCalendCodFFCC Or RAprob!nColocCalendCod = gColocCalendCodFFCCPG _
        Or RAprob!nColocCalendCod = gColocCalendCodPFCCPG Or RAprob!nColocCalendCod = gColocCalendCodPFCC Then
        nTipocalend = 2
    End If
    
    If RAprob!nColocCalendCod = gColocCalendCodFFCD Or RAprob!nColocCalendCod = gColocCalendCodFFCDPG _
        Or RAprob!nColocCalendCod = gColocCalendCodPFCDPG Or RAprob!nColocCalendCod = gColocCalendCodPFCD Then
        nTipocalend = 3
    End If
    Set oCalend = New NCalendario
    
    If pbMantieneCuota Then
        'Halla las cuotas aproximado para mantener la fecha de vencimiento del credito
        If RAprob!nPlazo = 0 Then
            nCuotasPendientes = PlazoDeCuotaFijaFechaFija(nCuotaObjetivo, R!nTasaInteres, _
                nCuotasPendientes, nMontoPrestamo, 0, RColoc!dVigencia, RAprob!nPeriodoFechaFija, _
                RAprob!nPeriodoGracia, RAprob!nProxMes)
        Else
            nCuotasPendientes = PlazoDeCuotaFija(R!nTasaInteres, nCuotasPendientes, nMontoPrestamo, _
                        IIf(IsNull(RAprob!nPlazo), 0, RAprob!nPlazo), nCuotaObjetivo)
        End If
    Else
        nCuotasPendientes = CuotasAproximada(pdHoy, RColoc!DVENC, RAprob!nPlazo, RAprob!nPeriodoFechaFija, RAprob!nProxMes)
    End If
    
    MatCalend = oCalend.GeneraCalendario(nMontoPrestamo, R!nTasaInteres, nCuotasPendientes, IIf(IsNull(RAprob!nPlazo), 0, RAprob!nPlazo), _
        pdHoy, nTipocalend, IIf(RAprob!nPeriodoFechaFija > 0, 2, 1), Exonerada, 0, RAprob!nPeriodoFechaFija, RAprob!nProxMes, , True)
        
    
    Set oCalend = Nothing
    R.Close
    Set R = Nothing
    RAprob.Close
    Set RAprob = Nothing
    
    ReprogramarCreditoenMemoriaTotal = MatCalend
    
End Function

Public Function ReprogramarCreditoSoloFechasMemoria(ByVal pnCuota As Integer, ByVal pnFechaFija As Integer, _
    ByVal pnPlazo As Integer, ByVal pMatCal As Variant) As Variant

Dim i As Integer
Dim dFecTemp As Date
Dim pnDiaFijo As Integer
Dim nMes, nAnio, nDia As Integer
Dim MatCalTmp As Variant
        MatCalTmp = pMatCal
        'Recalculo de fechas de Vencimiento
            dFecTemp = CDate(MatCalTmp(pnCuota - 1, 0))
            pnDiaFijo = pnFechaFija
            If pnFechaFija <> 0 Then
                nMes = Month(dFecTemp)
                nAnio = Year(dFecTemp)
                nDia = pnDiaFijo
                For i = pnCuota - 1 To UBound(MatCalTmp) - 1
                        nDia = pnDiaFijo
                        If Not (i = 0 And pnDiaFijo > Day(dFecTemp)) Then
                            nMes = nMes + 1
                            If nMes > 12 Then
                                nAnio = nAnio + 1
                                nMes = 1
                            End If
                        Else
                            If nMes = 2 Then
                                If nDia >= 29 Then
                                    If nAnio Mod 4 <> 0 Then
                                        nMes = nMes + 1
                                    End If
                                End If
                            Else
                                If nDia > 30 Then
                                    If nMes = 4 Or nMes = 6 Or nMes = 9 Or 11 Then
                                        nMes = nMes + 1
                                    End If
                                End If
                            End If
                        End If
                        If nMes = 2 Then
                            If nDia > 28 Then
                                If nAnio Mod 4 = 0 Then
                                    nDia = 29
                                Else
                                    nDia = 28
                                End If
                            End If
                        Else
                            If nDia > 30 Then
                                If nMes = 4 Or nMes = 6 Or nMes = 9 Or 11 Then
                                    nDia = 30
                                End If
                            End If
                        End If
                        dFecTemp = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)))
                        MatCalTmp(i, 0) = Format(dFecTemp, "dd/mm/yyyy")
                Next i
            Else
                For i = pnCuota - 1 To UBound(MatCalTmp) - 1
                    dFecTemp = dFecTemp + pnPlazo
                    MatCalTmp(i, 0) = Format(dFecTemp, "dd/mm/yyyy")
                Next i
            End If
        ReprogramarCreditoSoloFechasMemoria = MatCalTmp
End Function

Public Function ReprogramarCreditoMonto(ByVal pMatCalend As Variant, ByVal pnFila As Integer, ByVal pnMonto As Double)
Dim i As Integer
Dim nMontoCal As Double
Dim nDiferencia As Double

    nMontoCal = CDbl(pMatCalend(pnFila, 2)) + CDbl(pMatCalend(pnFila, 4)) + CDbl(pMatCalend(pnFila, 6)) + CDbl(pMatCalend(pnFila, 8)) + CDbl(pMatCalend(pnFila, 10)) + CDbl(pMatCalend(pnFila, 12)) + CDbl(pMatCalend(pnFila, 14))
    If nMontoCal <> pnMonto Then
        nDiferencia = pMatCalend(pnFila, 2)
        
        If pnMonto > CDbl(pMatCalend(pnFila, 4)) + CDbl(pMatCalend(pnFila, 6)) + CDbl(pMatCalend(pnFila, 8)) + CDbl(pMatCalend(pnFila, 10)) + CDbl(pMatCalend(pnFila, 12)) + CDbl(pMatCalend(pnFila, 14)) Then
            pMatCalend(pnFila, 2) = Format(pnMonto - (CDbl(pMatCalend(pnFila, 4)) + CDbl(pMatCalend(pnFila, 6)) + CDbl(pMatCalend(pnFila, 8)) + CDbl(pMatCalend(pnFila, 10)) + CDbl(pMatCalend(pnFila, 12)) + CDbl(pMatCalend(pnFila, 14))), "#0.00")
            nDiferencia = nDiferencia - CDbl(pMatCalend(pnFila, 2))
        End If
    End If
    
    pMatCalend(UBound(pMatCalend) - 1, 2) = Format(CDbl(pMatCalend(UBound(pMatCalend) - 1, 2)) + nDiferencia, "#0.00")
    
    ReprogramarCreditoMonto = pMatCalend
    
End Function


Public Function ReprogramarCreditoenMemoria(ByVal psCtaCod As String, ByVal pnTasa As Double, _
    ByVal pdFecOrig As Date, ByVal pdFecReprog As Date, ByVal nFilaReprog As Integer, ByVal pnTipoReprog As Integer, _
    Optional ByVal pbNoCargarMatriz As Boolean = False, Optional ByVal pMatCalend As Variant = "") As Variant

Dim MontoAReprogramar As Double
Dim MontoAProratear As Double
Dim MontoIntTemp As Double
Dim nDias As Integer
Dim i As Integer
Dim J As Integer
Dim nCapitalFusionado As Double
Dim MatCalTmp() As String
Dim nNumFilasReal As Integer
Dim nPunt As Integer
Dim nMontoCredito As Double
Dim nMontoMovido(7, 2) As Double
Dim nMontoMovidoProrateado(7, 2) As Double
Dim nMontoSaldo As Double
Dim MatCalend() As String
Dim oCal As Dcalendario
Dim oCredito As DCredito
Dim R As ADODB.Recordset
Dim nMontoDesemb As Double
Dim dFecTmp As Date
Dim nPlazoApr As Integer
Dim nFechaFija As Integer
Dim pnDiaFijo As Integer
Dim nMes As Integer
Dim nDia As Integer
Dim nAnio As Integer
Dim dFecTemp As Date
Dim nCuotaApr As Double

    'Carga Calendario A Reprogramar
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
    
    nMontoCredito = IIf(IsNull(R!nMonto), 0, R!nMonto)
    nMontoSaldo = IIf(IsNull(R!nMonto), 0, R!nMonto)
    nFechaFija = IIf(IsNull(R!nPeriodoFechaFija), 0, R!nPeriodoFechaFija)
    nPlazoApr = IIf(IsNull(R!nPlazo), 0, R!nPlazo)
    nMontoDesemb = nMontoCredito
    R.Close
    Set R = Nothing
    
    nMontoCredito = oCredito.SaldoPactadoCredito(psCtaCod)
    nMontoSaldo = nMontoCredito
    Set oCredito = Nothing
    Set oCal = New Dcalendario
    
    If Not pbNoCargarMatriz Or Not IsArray(pMatCalend) Then
        Set R = oCal.RecuperaCalendarioPagos(psCtaCod)
        ReDim MatCalend(R.RecordCount, 20)
        Do While Not R.EOF
            MatCalend(R.Bookmark - 1, 0) = Format(R!DVENC, "dd/mm/yyyy") 'Fecha de Vencimiento
            MatCalend(R.Bookmark - 1, 1) = Trim(Str(R!nCuota)) 'NroCuota
            MatCalend(R.Bookmark - 1, 2) = Format(IIf(IsNull(R!nCapital), 0, R!nCapital), "#0.00") 'Capital
            MatCalend(R.Bookmark - 1, 3) = Format(IIf(IsNull(R!nCapitalPag), 0, R!nCapitalPag), "#0.00") 'Capital Pagado
            MatCalend(R.Bookmark - 1, 4) = Format(IIf(IsNull(R!nIntComp), 0, R!nIntComp), "#0.00") 'Interes Compensatorio
            MatCalend(R.Bookmark - 1, 5) = Format(IIf(IsNull(R!nIntCompPag), 0, R!nIntCompPag), "#0.00") 'Interes Compensatorio Pagado
            MatCalend(R.Bookmark - 1, 6) = Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00") 'Interes Moratorio
            MatCalend(R.Bookmark - 1, 7) = Format(IIf(IsNull(R!nIntMorPag), 0, R!nIntMorPag), "#0.00") 'Interes Moratorio Pagado
            MatCalend(R.Bookmark - 1, 8) = Format(IIf(IsNull(R!nIntReprog), 0, R!nIntReprog), "#0.00") 'Interes Reprogramado
            MatCalend(R.Bookmark - 1, 9) = Format(IIf(IsNull(R!nIntReprogPag), 0, R!nIntReprogPag), "#0.00") 'Interes Reprogramado Pagado
            MatCalend(R.Bookmark - 1, 10) = Format(IIf(IsNull(R!nIntGracia), 0, R!nIntGracia), "#0.00") 'Interes Gracia
            MatCalend(R.Bookmark - 1, 11) = Format(IIf(IsNull(R!nIntGraciaPag), 0, R!nIntGraciaPag), "#0.00") 'Interes Gracia Pagado
            MatCalend(R.Bookmark - 1, 12) = Format(IIf(IsNull(R!nIntSuspenso), 0, R!nIntSuspenso), "#0.00") 'Interes Suspenso
            MatCalend(R.Bookmark - 1, 13) = Format(IIf(IsNull(R!nIntSuspensoPag), 0, R!nIntSuspensoPag), "#0.00") 'Interes Suspenso Pagado
            MatCalend(R.Bookmark - 1, 14) = Format(IIf(IsNull(R!NGasto), 0, R!NGasto), "#0.00") 'Gastos
            MatCalend(R.Bookmark - 1, 15) = Format(IIf(IsNull(R!nGastoPag), 0, R!nGastoPag), "#0.00") 'Gastos Pagados
            nMontoCredito = nMontoCredito - IIf(IsNull(R!nCapital), 0, R!nCapital)
            nMontoCredito = CDbl(Format(nMontoCredito, "#0.00"))
            MatCalend(R.Bookmark - 1, 16) = Format(nMontoCredito, "#0.00") 'Saldo Capital
            MatCalend(R.Bookmark - 1, 17) = Trim(Str(R!nColocCalendEstado))
            MatCalend(R.Bookmark - 1, 18) = Format(IIf(IsNull(R!nIntCompVenc), 0, R!nIntCompVenc), "#0.00") 'Interes Compensatorio Vencido
            MatCalend(R.Bookmark - 1, 19) = Format(IIf(IsNull(R!nIntCompVencPag), 0, R!nIntCompVencPag), "#0.00") 'Interes Compensatorio Vencido Pagado
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
    End If
    
    
    
    If IsArray(pMatCalend) Then
        MatCalend = pMatCalend
    End If
        
    If CDate(MatCalend(nFilaReprog, 0)) = pdFecReprog Then
        ReprogramarCreditoenMemoria = MatCalend
        Exit Function
    End If
    
    MatCalend(nFilaReprog, 0) = Format(pdFecReprog, "dd/mm/yyyy")
    'Mover la Fechas
    nCapitalFusionado = 0
    For i = 0 To 6
        nMontoMovido(i, 0) = 0#
        nMontoMovido(i, 1) = 0#
    Next i
    
    If pnTipoReprog = 1 Then 'Proratear
        nNumFilasReal = UBound(MatCalend)
    Else
        nNumFilasReal = UBound(MatCalend) + 1
    End If
    ReDim MatCalTmp(nNumFilasReal, 20)
    nPunt = 0
    J = 0
    For i = 0 To UBound(MatCalend) - 1
        If MatCalend(i, 0) <> "" Then
            MatCalTmp(nPunt, 0) = MatCalend(i, 0)
            MatCalTmp(nPunt, 1) = nPunt + 1
            MatCalTmp(nPunt, 2) = MatCalend(i, 2)
            MatCalTmp(nPunt, 3) = MatCalend(i, 3)
            MatCalTmp(nPunt, 4) = MatCalend(i, 4)
            MatCalTmp(nPunt, 5) = MatCalend(i, 5)
            MatCalTmp(nPunt, 6) = MatCalend(i, 6)
            MatCalTmp(nPunt, 7) = MatCalend(i, 7)
            MatCalTmp(nPunt, 8) = MatCalend(i, 8)
            MatCalTmp(nPunt, 9) = MatCalend(i, 9)
            MatCalTmp(nPunt, 10) = MatCalend(i, 10)
            MatCalTmp(nPunt, 11) = MatCalend(i, 11)
            MatCalTmp(nPunt, 12) = MatCalend(i, 12)
            MatCalTmp(nPunt, 13) = MatCalend(i, 13)
            MatCalTmp(nPunt, 14) = MatCalend(i, 14)
            MatCalTmp(nPunt, 15) = MatCalend(i, 15)
            MatCalTmp(nPunt, 17) = MatCalend(i, 17)
            MatCalTmp(nPunt, 18) = MatCalend(i, 18)
            MatCalTmp(nPunt, 19) = MatCalend(i, 19)
            nPunt = nPunt + 1
            'Solo Pendientes
            If CInt(Trim(MatCalTmp(nPunt - 1, 17))) = gColocCalendEstadoPendiente Then
                J = J + 1 'Total Pendientes
            End If
        End If
    Next i
        
    'Recalculo de Saldos
    For i = 0 To UBound(MatCalend) - 1
        nMontoSaldo = nMontoSaldo - CDbl(MatCalTmp(i, 2))
        nMontoSaldo = CDbl(Format(nMontoSaldo, "#0.00"))
        MatCalTmp(i, 16) = Format(nMontoSaldo, "#0.00")
    Next i
    
    If pnTipoReprog <> 3 Then 'Proratear
    
        'Calulo por Dias Reprogramados
        nDias = CDate(MatCalTmp(nFilaReprog, 0)) - pdFecOrig
        If nFilaReprog = 0 Then
            MontoAReprogramar = TasaIntPerDias(pnTasa, nDias) * nMontoDesemb
        Else
            MontoAReprogramar = TasaIntPerDias(pnTasa, nDias) * CDbl(MatCalTmp(nFilaReprog - 1, 16))
        End If
        MontoAReprogramar = CDbl(Format(MontoAReprogramar, "#0.00"))
    Else
        nFilaReprog = UBound(MatCalTmp) - 2
    End If
        'Recalculo de fechas de Vencimiento
        dFecTmp = CDate(MatCalTmp(nFilaReprog, 0))
        pnDiaFijo = nFechaFija
        If nFechaFija <> 0 Then
            nMes = Month(dFecTmp)
            nAnio = Year(dFecTmp)
            nDia = pnDiaFijo
            If MatCalTmp(nFilaReprog + 1, 2) <> "" Then
                nCuotaApr = CDbl(Format(CDbl(MatCalTmp(nFilaReprog + 1, 2)) + CDbl(MatCalTmp(nFilaReprog + 1, 4)), "#0.00"))
            End If
            For i = nFilaReprog + 1 To UBound(MatCalTmp) - 1
                    nDia = pnDiaFijo
                    If Not (i = 0 And pnDiaFijo > Day(dFecTmp)) Then
                        nMes = nMes + 1
                        If nMes > 12 Then
                            nAnio = nAnio + 1
                            nMes = 1
                        End If
                    Else
                        If nMes = 2 Then
                            If nDia >= 29 Then
                                If nAnio Mod 4 <> 0 Then
                                    nMes = nMes + 1
                                End If
                            End If
                        Else
                            If nDia > 30 Then
                                If nMes = 4 Or nMes = 6 Or nMes = 9 Or 11 Then
                                    nMes = nMes + 1
                                End If
                            End If
                        End If
                    End If
                    If nMes = 2 Then
                        If nDia > 28 Then
                            If nAnio Mod 4 = 0 Then
                                nDia = 29
                            Else
                                nDia = 28
                            End If
                        End If
                    Else
                        If nDia > 30 Then
                            If nMes = 4 Or nMes = 6 Or nMes = 9 Or 11 Then
                                nDia = 30
                            End If
                        End If
                    End If
                    dFecTemp = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)))
                    MatCalTmp(i, 0) = Format(dFecTemp, "dd/mm/yyyy")
                    If i = nFilaReprog + 1 Then
                        nDias = CDate(MatCalTmp(i, 0)) - dFecTmp
                    Else
                        nDias = CDate(MatCalTmp(i, 0)) - CDate(MatCalTmp(i - 1, 0))
                    End If
                    MontoIntTemp = TasaIntPerDias(pnTasa, nDias) * CDbl(MatCalTmp(i - 1, 16))
                    MatCalTmp(i, 4) = Format(MontoIntTemp, "#0.00")
                    MatCalTmp(i, 2) = CDbl(Format(nCuotaApr - CDbl(MatCalTmp(i, 4)), "#0.00"))
                    
                    If UBound(MatCalTmp) - 1 = i Then
                        If CDbl(MatCalTmp(i - 1, 16)) < CDbl(MatCalTmp(i, 2)) Then
                            MatCalTmp(i, 2) = MatCalTmp(i - 1, 16)
                        End If
                    End If
                    MatCalTmp(i, 16) = Format(CDbl(MatCalTmp(i - 1, 16)) - CDbl(MatCalTmp(i, 2)), "#0.00")
            Next i
            nMontoSaldo = MatCalTmp(nFilaReprog, 16)
            For i = nFilaReprog + 1 To UBound(MatCalTmp) - 1
                If pnTipoReprog <> 2 Then
                    If i = UBound(MatCalTmp) - 1 Then
                        MatCalTmp(i, 2) = Format(nMontoSaldo, "#0.00")
                    End If
                Else
                    If i = UBound(MatCalTmp) - 2 Then
                        MatCalTmp(i, 2) = Format(nMontoSaldo, "#0.00")
                    End If
                End If
                nMontoSaldo = nMontoSaldo - CDbl(MatCalTmp(i, 2))
                nMontoSaldo = Format(nMontoSaldo, "#0.00")
                MatCalTmp(i, 16) = Format(nMontoSaldo, "#0.00")
            Next i
        Else
            For i = nFilaReprog + 1 To UBound(MatCalTmp) - 1
                dFecTmp = dFecTmp + nPlazoApr
                MatCalTmp(i, 0) = Format(dFecTmp, "dd/mm/yyyy")
            Next i
        End If
    
    If pnTipoReprog = 1 Then 'Proratear
        MontoAProratear = Format((MontoAReprogramar / (UBound(MatCalTmp) - nFilaReprog)), "#0.00")
    
        For i = nFilaReprog To UBound(MatCalTmp) - 1
            MatCalTmp(i, 8) = Format(CDbl(MatCalTmp(i, 8)) + CDbl(MontoAProratear), "#0.00")
        Next i
    
    Else    'A la Ultima Cuota
        
        If pnTipoReprog = 3 Then 'Pasar Interes a Ultima Cuota (Especial)
            MontoIntTemp = 0
            For i = 0 To UBound(MatCalTmp) - 2
                If CDbl(MatCalTmp(i, 2)) > 0 And MatCalTmp(i, 17) = "0" Then
                    MontoIntTemp = MontoIntTemp + CDbl(MatCalTmp(i, 4))
                    MatCalTmp(i, 4) = "0.00"
                End If
            Next i
            
            MatCalTmp(UBound(MatCalTmp) - 1, 1) = Trim(Str(UBound(MatCalTmp)))
            MatCalTmp(UBound(MatCalTmp) - 1, 2) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 3) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 4) = Format(CDbl(MontoIntTemp), "#0.00")
            MatCalTmp(UBound(MatCalTmp) - 1, 5) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 6) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 7) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 8) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 9) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 10) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 11) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 12) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 13) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 14) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 15) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 16) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 17) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 18) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 19) = "0.00"
        Else
            MatCalTmp(UBound(MatCalTmp) - 1, 1) = Trim(Str(UBound(MatCalTmp)))
            MatCalTmp(UBound(MatCalTmp) - 1, 2) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 3) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 4) = Format(CDbl(MontoAReprogramar), "#0.00")
            MatCalTmp(UBound(MatCalTmp) - 1, 5) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 6) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 7) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 8) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 9) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 10) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 11) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 12) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 13) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 14) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 15) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 16) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 17) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 18) = "0.00"
            MatCalTmp(UBound(MatCalTmp) - 1, 19) = "0.00"
        End If
    End If
    Set oCal = Nothing
    ReprogramarCreditoenMemoria = MatCalTmp
End Function

Public Function ReprogramarCredito(ByVal psCtaCod As String, ByVal MatCalend As Variant, _
    ByVal pnTipoReprog As Integer, Optional pnMiViv As Boolean = False, Optional MatCalend_2 As Variant = "", Optional pdHoy As Date = CDate("01/01/1900")) As String
'pnTipoReprog = 1 Reprogramacion con Edicion de Fecha
'pnTipoReprog = 2 Reprogramacion Total Otorgamiento de un Nuevo
'                 Credito bajo el Saldo de capital Existente y las mismas condiciones que el anterior
Dim oCredito As DCredito
Dim oBase As DCredActualizaBD
Dim oCalend As Dcalendario
Dim nNroCalen As Integer
Dim R As ADODB.Recordset
Dim RAprob As ADODB.Recordset
Dim i As Integer
Dim oGastos As NGasto
Dim nNumGastos As Integer
Dim MatDesemb As Variant
Dim MatGastos As Variant
Dim nMontoPrestamoRep As Double
Dim sTipoGasto As String

    On Error GoTo ErrorReprogramarCredito
    ReDim MatDesemb(0, 0)
    ReprogramarCredito = ""
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaColocacCred(psCtaCod)
    Set RAprob = oCredito.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
    sTipoGasto = IIf(IsNull(RAprob!cTipoGasto), "F", RAprob!cTipoGasto)
    If pnMiViv Then
        nNroCalen = R!nNroCalPar + 1
    Else
        nNroCalen = R!nNroCalen + 1
    End If
    R.Close
    Set R = Nothing
    
    If pnMiViv Then
        Set R = oCredito.RecuperaProducto(psCtaCod)
        nMontoPrestamoRep = R!nSaldo
        R.Close
        Set oGastos = New NGasto
        MatGastos = oGastos.GeneraCalendarioGastos(MatCalend, MatDesemb, nNumGastos, pdHoy, psCtaCod, 1, "DE", sTipoGasto, CDbl(MatCalend(0, 2)), nMontoPrestamoRep)
        Set oGastos = Nothing
    End If
    
    Set oBase = New DCredActualizaBD
    oBase.dBeginTrans
    
    If pnTipoReprog = 1 Then
        
        For i = 0 To UBound(MatCalend) - 1
            Call oBase.dInsertColocCalendario(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(i, 1)), CDate(MatCalend(i, 0)), CInt(MatCalend(i, 17)), "Calendario Reprogramado", gColocCalendConceptoProcReprogramado, False)
            If CDbl(MatCalend(i, 2)) > 0 Then
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(i, 1)), gColocConceptoCodCapital, CDbl(MatCalend(i, 2)), CDbl(MatCalend(i, 3)), "", False)
            End If
            
            If CDbl(MatCalend(i, 4)) > 0 Then
            Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(i, 1)), gColocConceptoCodInteresCompensatorio, CDbl(MatCalend(i, 4)), CDbl(MatCalend(i, 5)), "", False)
            End If
            
            'Tiene que insertarle si o si para cobrar mora
            Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(i, 1)), gColocConceptoCodInteresMoratorio, CDbl(MatCalend(i, 6)), CDbl(MatCalend(i, 7)), "", False)
            
            'Intereses Reprogramados
            'Graba Como Nuevo Concepto
            'If CDbl(MatCalend(I, 8)) > 0 Then
            '    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(I, 1)), gColocConceptoCodInteresReprogramado, CDbl(MatCalend(I, 8)), CDbl(MatCalend(I, 9)), "", False)
            'End If
            'Graba Acumulandolo en Int Comp Vencido por Presentar Problemas al Calcular Deuda a la Fecha
            If CDbl(MatCalend(i, 8)) > 0 Then
                Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(i, 1)), gColocConceptoCodInteresCompensatorio, CDbl(MatCalend(i, 8)), , "", False, , True)
            End If
            
            If CDbl(MatCalend(i, 10)) > 0 Then
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(i, 1)), gColocConceptoCodInteresGracia, CDbl(MatCalend(i, 10)), CDbl(MatCalend(i, 11)), "", False)
            End If
            If CDbl(MatCalend(i, 12)) > 0 Then
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(i, 1)), gColocConceptoCodInteresSuspenso, CDbl(MatCalend(i, 12)), CDbl(MatCalend(i, 13)), "", False)
            End If
            
            If CDbl(MatCalend(i, 14)) > 0 Then
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(i, 1)), gColocConceptoCodGastoVarios, CDbl(MatCalend(i, 14)), CDbl(MatCalend(i, 15)), "", False)
            End If
                  
            'If CDbl(MatCalend(I, 18)) > 0 Then
            'Tiene que insertarle si o si para cobrar Interes Comp vencido
            Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(i, 1)), gColocConceptoCodInteresCompVencido, CDbl(MatCalend(i, 18)), CDbl(MatCalend(i, 19)), "", False)
            'End If
            
        Next i
        
    Else
              
        For i = 0 To UBound(MatCalend) - 1
            Call oBase.dInsertColocCalendario(psCtaCod, nNroCalen, gColocCalendAplCuota, i + 1, CDate(MatCalend(i, 0)), gColocCalendEstadoPendiente, "Calendario Reporgramado de Pagos", gColocCalendConceptoProcReprogramado, False)
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, i + 1, gColocConceptoCodCapital, CDbl(MatCalend(i, 3)), 0, "", False)
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalend(i, 4)), 0, "", False)
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", False)
                If CDbl(MatCalend(i, 5)) > 0 Then
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresGracia, CDbl(MatCalend(i, 5)), 0, "", False)
                End If
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresCompVencido, 0, 0, "", False)
        Next i
        
        If pnMiViv Then
            
            For i = 0 To nNumGastos - 1
                If CInt(Trim(Right(MatGastos(i, 0), 10))) = gColocCalendAplCuota Then
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, CInt(Trim(Right(MatGastos(i, 0), 3))), CInt(Trim(MatGastos(i, 1))), CLng(Trim(Right(MatGastos(i, 2), 20))), CDbl(MatGastos(i, 3)), 0, "", False)
                End If
            Next i
            
            For i = 0 To UBound(MatCalend_2) - 1
                If i >= 0 And i <= 5 Then
                    Call oBase.dInsertColocCalendario(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, i + 1, CDate(MatCalend_2(i, 0)), gColocCalendEstadoPendiente, "Calendario Reporgramado de Pagos", gColocCalendConceptoProcReprogramado, False)
                Else
                    Call oBase.dInsertColocCalendario(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, i + 1, CDate(MatCalend_2(i, 0)), gColocCalendEstadoPendiente, "Calendario Reporgramado de Pagos", gColocCalendConceptoProcReprogramado, False)
                End If
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, i + 1, gColocConceptoCodCapital, CDbl(MatCalend_2(i, 3)), 0, "", False)
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalend_2(i, 4)), 0, "", False)
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", False)
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresCompVencido, 0, 0, "", False)
                    If CDbl(MatCalend_2(i, 5)) > 0 Then
                        Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, i + 1, gColocConceptoCodInteresGracia, CDbl(MatCalend_2(i, 5)), 0, "", False)
                    End If
            Next i
        End If
    End If
    
    Call oBase.dUpdateColocacCred(psCtaCod, , , , , , , , , , , , , , nNroCalen, , , , False, , , nNroCalen + 1)
    oBase.dCommitTrans
    
    Set oBase = Nothing
    Exit Function

ErrorReprogramarCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Private Function MatrizDistribuirGastos(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, ByRef pnMonto As Double, _
        Optional ByVal DistVert As Boolean = False)
Dim i As Integer

    For i = 0 To UBound(MatCalend) - 1
        If i > 0 And Not DistVert Then
            'Si Aun queda pendiente capital, interes, gastos, mora de la cuota anterior
            If (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 3)) - CDbl(MatCalendDistrib(i - 1, 3))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 8)) - CDbl(MatCalendDistrib(i - 1, 8))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 7)) - CDbl(MatCalendDistrib(i - 1, 7))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 5)) - CDbl(MatCalendDistrib(i - 1, 5))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 4)) - CDbl(MatCalendDistrib(i - 1, 4))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 6)) - CDbl(MatCalendDistrib(i - 1, 6))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 And _
                (CDbl(MatCalend(i - 1, 9)) - CDbl(MatCalendDistrib(i - 1, 9))) > 0) Then
                
                Exit Function
            End If
        End If
    
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 And (CDbl(MatCalend(i, 9)) - CDbl(MatCalendDistrib(i, 9))) > 0 Then
            If pnMonto > (CDbl(MatCalend(i, 9)) - CDbl(MatCalendDistrib(i, 9))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 9)) - CDbl(MatCalendDistrib(i, 9)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 9) = MatCalend(i, 9)
            Else
                MatCalendDistrib(i, 9) = Format(CDbl(MatCalendDistrib(i, 9)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
            If Not DistVert Then
                Exit For
            End If
        End If
    Next i
    
End Function
'*********************************************************
'** Cubrir Verticalmente Todas las Moras de las Cuotas
'*********************************************************
Private Function MatrizDistribuirMora(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, _
ByRef pnMonto As Double, Optional ByVal DistVert As Boolean = False)
Dim i As Integer

    For i = 0 To UBound(MatCalend) - 1
        
        If i > 0 And Not DistVert Then
            'Si Aun queda pendiente capital, interes, gastos, mora de la cuota anterior
            If (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 3)) - CDbl(MatCalendDistrib(i - 1, 3))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 8)) - CDbl(MatCalendDistrib(i - 1, 8))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 7)) - CDbl(MatCalendDistrib(i - 1, 7))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 5)) - CDbl(MatCalendDistrib(i - 1, 5))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 4)) - CDbl(MatCalendDistrib(i - 1, 4))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 6)) - CDbl(MatCalendDistrib(i - 1, 6))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 And _
                (CDbl(MatCalend(i - 1, 9)) - CDbl(MatCalendDistrib(i - 1, 9))) > 0) Then
                
                Exit Function
            End If
        End If
        
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(i, 6)) - CDbl(MatCalendDistrib(i, 6))) > 0 Then
            If pnMonto > (CDbl(MatCalend(i, 6)) - CDbl(MatCalendDistrib(i, 6))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 6)) - CDbl(MatCalendDistrib(i, 6)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 6) = MatCalend(i, 6)
            Else
                MatCalendDistrib(i, 6) = Format(CDbl(MatCalendDistrib(i, 6)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
    Next i
End Function

Private Function MatrizDistribuirInteresI(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, _
    ByRef pnMonto As Double, Optional ByVal DistVert As Boolean = False)
Dim i As Integer
Dim bSiCubrio As Boolean
    bSiCubrio = False
    For i = 0 To UBound(MatCalend) - 1
        If i > 0 And Not DistVert Then
            'Si Aun queda pendiente capital, interes, gastos, mora de la cuota anterior
            If (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 3)) - CDbl(MatCalendDistrib(i - 1, 3))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 8)) - CDbl(MatCalendDistrib(i - 1, 8))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 7)) - CDbl(MatCalendDistrib(i - 1, 7))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 5)) - CDbl(MatCalendDistrib(i - 1, 5))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 4)) - CDbl(MatCalendDistrib(i - 1, 4))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 6)) - CDbl(MatCalendDistrib(i - 1, 6))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 And _
                (CDbl(MatCalend(i - 1, 9)) - CDbl(MatCalendDistrib(i - 1, 9))) > 0) Then
                
                Exit Function
            End If
        End If
    
        'Cubre Interes en Suspenso
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(i, 8)) - CDbl(MatCalendDistrib(i, 8))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(i, 8)) - CDbl(MatCalendDistrib(i, 8))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 8)) - CDbl(MatCalendDistrib(i, 8)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 8) = MatCalend(i, 8)
            Else
                MatCalendDistrib(i, 8) = Format(CDbl(MatCalendDistrib(i, 8)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        'Cubre Intereses Reprogramados
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(i, 7)) - CDbl(MatCalendDistrib(i, 7))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(i, 7)) - CDbl(MatCalendDistrib(i, 7))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 7)) - CDbl(MatCalendDistrib(i, 7)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 7) = MatCalend(i, 7)
            Else
                MatCalendDistrib(i, 7) = Format(CDbl(MatCalendDistrib(i, 7)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        'Intereses de Gracia
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(i, 5)) - CDbl(MatCalendDistrib(i, 5))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(i, 5)) - CDbl(MatCalendDistrib(i, 5))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 5)) - CDbl(MatCalendDistrib(i, 5)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 5) = MatCalend(i, 5)
            Else
                MatCalendDistrib(i, 5) = Format(CDbl(MatCalendDistrib(i, 5)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        'Intereses Compensatorio Vencido
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(i, 11)) - CDbl(MatCalendDistrib(i, 11))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(i, 11)) - CDbl(MatCalendDistrib(i, 11))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 11)) - CDbl(MatCalendDistrib(i, 11)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 11) = MatCalend(i, 11)
            Else
                MatCalendDistrib(i, 11) = Format(CDbl(MatCalendDistrib(i, 11)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        'Intereses Compensatorios
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(i, 4)) - CDbl(MatCalendDistrib(i, 4))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(i, 4)) - CDbl(MatCalendDistrib(i, 4))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 4)) - CDbl(MatCalendDistrib(i, 4)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 4) = MatCalend(i, 4)
            Else
                MatCalendDistrib(i, 4) = Format(CDbl(MatCalendDistrib(i, 4)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        If bSiCubrio Then
            If Not DistVert Then
                Exit For
            End If
        End If
    Next i
End Function

Private Function MatrizDistribuirCapital(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, _
    ByRef pnMonto As Double, Optional ByVal DistVert As Boolean = False)
Dim i As Integer

    For i = 0 To UBound(MatCalend) - 1
    
        If i > 0 And Not DistVert Then
            'Si Aun queda pendiente capital, interes, gastos, mora de la cuota anterior
            If (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 3)) - CDbl(MatCalendDistrib(i - 1, 3))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 8)) - CDbl(MatCalendDistrib(i - 1, 8))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 7)) - CDbl(MatCalendDistrib(i - 1, 7))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 5)) - CDbl(MatCalendDistrib(i - 1, 5))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 4)) - CDbl(MatCalendDistrib(i - 1, 4))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 6)) - CDbl(MatCalendDistrib(i - 1, 6))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 And _
                (CDbl(MatCalend(i - 1, 9)) - CDbl(MatCalendDistrib(i - 1, 9))) > 0) Then
                
                Exit Function
            End If
        End If
    
        If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(i, 3)) - CDbl(MatCalendDistrib(i, 3))) > 0 Then
            If pnMonto > (CDbl(MatCalend(i, 3)) - CDbl(MatCalendDistrib(i, 3))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 3)) - CDbl(MatCalendDistrib(i, 3)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 3) = MatCalend(i, 3)
            Else
                MatCalendDistrib(i, 3) = Format(CDbl(MatCalendDistrib(i, 3)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
            If Not DistVert Then
                Exit For
            End If
        End If
    Next i
End Function


Private Sub MatrizActualizarEstadoCuota(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant)
Dim i As Integer
    For i = 0 To UBound(MatCalend) - 1
        If CInt(MatCalendDistrib(i, 2)) = gColocCalendEstadoPendiente Then
                If CDbl(MatCalend(i, 8)) = CDbl(MatCalendDistrib(i, 8)) And _
                    CDbl(MatCalend(i, 7)) = CDbl(MatCalendDistrib(i, 7)) And _
                    CDbl(MatCalend(i, 6)) = CDbl(MatCalendDistrib(i, 6)) And _
                    CDbl(MatCalend(i, 5)) = CDbl(MatCalendDistrib(i, 5)) And _
                    CDbl(MatCalend(i, 4)) = CDbl(MatCalendDistrib(i, 4)) And _
                    CDbl(MatCalend(i, 3)) = CDbl(MatCalendDistrib(i, 3)) And CDbl(MatCalend(i, 11)) = CDbl(MatCalendDistrib(i, 11)) Then
                    MatCalendDistrib(i, 2) = Trim(Str(gColocCalendEstadoPagado))
                End If
        End If
    Next i
End Sub

Public Function MatrizDistribuirCancelacion(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pnMontoPago As Double, _
                 ByVal psMetLiquidacion As String, ByVal pdHoy As Date, Optional ByVal pbNoCancelar As Boolean = False) As Variant

Dim MatCalendDistrib As Variant
Dim nMontoGastos As Double
Dim nMontoMora As Double
Dim nMontoInteres As Double
Dim nMontoCapital As Double
Dim nMontoTotalTmp As Double
Dim nMontoPagoTmp As Double
Dim J As Integer

        MatCalendDistrib = CrearMatrizparaAmortizacion(MatCalend)
        nMontoTotalTmp = 0
        nMontoPagoTmp = pnMontoPago
        For J = 1 To 4
                Select Case Mid(psMetLiquidacion, J, 1)
                    Case "G"
                        nMontoGastos = MatrizGastosFecha(psCtaCod, MatCalend)
                        If nMontoPagoTmp > nMontoGastos Then
                            nMontoPagoTmp = nMontoPagoTmp - nMontoGastos
                        Else
                            nMontoGastos = nMontoPagoTmp
                            nMontoPagoTmp = 0#
                        End If
                        nMontoTotalTmp = nMontoTotalTmp + nMontoGastos
                        Call MatrizDistribuirGastos(MatCalend, MatCalendDistrib, nMontoGastos, True)
                    Case "M"
                        nMontoMora = MatrizInteresMorFecha(psCtaCod, MatCalend)
                        If nMontoPagoTmp > nMontoMora Then
                            nMontoPagoTmp = nMontoPagoTmp - nMontoMora
                        Else
                            nMontoMora = nMontoPagoTmp
                            nMontoPagoTmp = 0#
                        End If
                        nMontoTotalTmp = nMontoTotalTmp + nMontoMora
                        Call MatrizDistribuirMora(MatCalend, MatCalendDistrib, nMontoMora, True)
                    Case "I", "i", "Y"
                        nMontoInteres = CDbl(Format(MatrizInteresTotalesAFechaSinMora(psCtaCod, MatCalend, pdHoy), "#0.00"))
                        If nMontoPagoTmp > nMontoInteres Then
                            nMontoPagoTmp = nMontoPagoTmp - nMontoInteres
                        Else
                            nMontoInteres = nMontoPagoTmp
                            nMontoPagoTmp = 0#
                        End If
                        nMontoTotalTmp = nMontoTotalTmp + nMontoInteres
                        Call MatrizDistribuirInteresI(MatCalend, MatCalendDistrib, nMontoInteres, True)
                    Case "C"
                        nMontoCapital = MatrizCapitalAFecha(psCtaCod, MatCalend)
                        If nMontoPagoTmp > nMontoCapital Then
                            nMontoPagoTmp = nMontoPagoTmp - nMontoCapital
                        Else
                            nMontoCapital = nMontoPagoTmp
                            nMontoPagoTmp = 0#
                        End If
                        nMontoTotalTmp = nMontoTotalTmp + nMontoCapital
                        Call MatrizDistribuirCapital(MatCalend, MatCalendDistrib, nMontoCapital, True)
                End Select
        Next J
        If (Mid(psMetLiquidacion, 3, 1) = "i" Or Mid(psMetLiquidacion, 3, 1) = "Y") And pbNoCancelar Then
            If Mid(psMetLiquidacion, 3, 1) = "Y" Then
                Call MatrizActualizarEstadoCalendCancelado(MatCalendDistrib, True, pbNoCancelar, MatCalend)
            Else
                Call MatrizActualizarEstadoCalendCancelado(MatCalendDistrib, False, pbNoCancelar, MatCalend)
            End If
        Else
            Call MatrizActualizarEstadoCalendCancelado(MatCalendDistrib, , False)
        End If
        MatrizDistribuirCancelacion = MatCalendDistrib
End Function

Public Function MatrizDistribuirCalendCuotaLibre(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pnMontoPago As Double, _
                 ByVal psMetLiquidacion As String, ByVal pdHoy As Date, ByVal pnIntPend As Double, _
                 ByRef pnIntPendientePagado As Double) As Variant

Dim MatCalendDistrib As Variant
Dim nMontoGastos As Double
Dim nMontoMora As Double
Dim nMontoInteres As Double
Dim nMontoCapital As Double
Dim nMontoTotalTmp As Double
Dim nMontoPago As Double
Dim J As Integer

        MatCalendDistrib = CrearMatrizparaAmortizacion(MatCalend)
        nMontoTotalTmp = 0
        nMontoPago = pnMontoPago
        
        'Cubre Gastos
        If nMontoPago > 0 Then
            Call MatrizDistribuirGastos(MatCalend, MatCalendDistrib, nMontoPago, True)
        End If
        'Cubre Mora
        If nMontoPago > 0 Then
            Call MatrizDistribuirMora(MatCalend, MatCalendDistrib, nMontoPago)
        End If
        'Cubre Interes
        If nMontoPago > 0 Then
            If nMontoPago > pnIntPend Then
                nMontoPago = nMontoPago - pnIntPend
                pnIntPendientePagado = pnIntPend
            Else
                pnIntPendientePagado = nMontoPago
                nMontoPago = 0
                Exit Function
            End If
            Call MatrizDistribuirInteresI(MatCalend, MatCalendDistrib, nMontoPago, True)
        End If
        'Cubrir Capital
        If nMontoPago > 0 Then
            Call MatrizDistribuirCapital(MatCalend, MatCalendDistrib, nMontoPago, True)
        End If
        MatrizDistribuirCalendCuotaLibre = MatCalendDistrib
        
End Function


Public Function MatrizDistribuirCalendDinamico(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pnMontoPago As Double, _
                 ByVal psMetLiquidacion As String, ByVal pdHoy As Date) As Variant

Dim MatCalendDistrib As Variant
Dim nMontoGastos As Double
Dim nMontoMora As Double
Dim nMontoInteres As Double
Dim nMontoCapital As Double
Dim nMontoTotalTmp As Double
Dim nMontoPago As Double
Dim J As Integer
Dim nMivivienda As Integer
Dim oCred As DCredito
Dim R As ADODB.Recordset

        Set oCred = New DCredito
        Set R = oCred.RecuperaColocacCred(psCtaCod)
        Set oCred = Nothing
        nMivivienda = IIf(IsNull(R!bMiVivienda), 0, R!bMiVivienda)
        R.Close
        
        MatCalendDistrib = CrearMatrizparaAmortizacion(MatCalend)
        nMontoTotalTmp = 0
        nMontoPago = pnMontoPago
        For J = 1 To 4
            Select Case Mid(psMetLiquidacion, J, 1)
                Case "G"
                    If nMivivienda = 1 Then
                        nMontoGastos = MatrizGastosVencidos(MatCalend, pdHoy)
                    Else
                        nMontoGastos = MatrizGastosFecha(psCtaCod, MatCalend)
                    End If
                    nMontoPago = nMontoPago - nMontoGastos
                    Call MatrizDistribuirGastos(MatCalend, MatCalendDistrib, nMontoGastos, True)
                Case "M"
                    nMontoMora = MatrizInteresMorFecha(psCtaCod, MatCalend)
                    nMontoPago = nMontoPago - nMontoMora
                    Call MatrizDistribuirMora(MatCalend, MatCalendDistrib, nMontoMora, True)
                Case "I"
                    nMontoInteres = CDbl(Format(MatrizInteresTotalesAFechaSinMora(psCtaCod, MatCalend, pdHoy), "#0.00"))
                    nMontoPago = nMontoPago - nMontoInteres
                    Call MatrizDistribuirInteresI(MatCalend, MatCalendDistrib, nMontoInteres, True)
                Case "C"
                    nMontoCapital = MatrizCapitalVencido(MatCalend, pdHoy)
                    nMontoPago = nMontoPago - nMontoCapital
                    If nMontoCapital > 0 Then
                        Call MatrizDistribuirCapital(MatCalend, MatCalendDistrib, nMontoCapital, True)
                    End If
            End Select
            nMontoPago = CDbl(Format(nMontoPago, "#0.00"))
            Call MatrizActualizarEstadoCuota(MatCalend, MatCalendDistrib)
        Next J
        If nMontoPago > 0 Then
            Call MatrizDistribuirCapital(MatCalend, MatCalendDistrib, nMontoPago, True)
            Call MatrizActualizarEstadoCuota(MatCalend, MatCalendDistrib)
        End If
        MatrizDistribuirCalendDinamico = MatCalendDistrib
End Function

Public Function MatrizDistribuirMontoMiVivienda(ByVal MatCalend As Variant, ByVal pnMontoPago As Double, _
                    ByVal psMetLiquidacion As String) As Variant
Dim i As Integer
Dim J As Integer
Dim MatCalendDistrib As Variant
Dim nMonto As Double

        nMonto = pnMontoPago
        MatCalendDistrib = CrearMatrizparaAmortizacion(MatCalend)
        Do While nMonto > 0
            For J = 1 To 4
                Select Case Mid(psMetLiquidacion, J, 1)
                    Case "G"
                        Call MatrizDistribuirGastos(MatCalend, MatCalendDistrib, nMonto)
                    Case "M"
                        Call MatrizDistribuirMora(MatCalend, MatCalendDistrib, nMonto)
                    Case "I"
                        Call MatrizDistribuirInteresI(MatCalend, MatCalendDistrib, nMonto)
                    Case "C"
                        Call MatrizDistribuirCapital(MatCalend, MatCalendDistrib, nMonto)
                End Select
                Call MatrizActualizarEstadoCuota(MatCalend, MatCalendDistrib)
            Next J
        Loop
        MatrizDistribuirMontoMiVivienda = MatCalendDistrib
End Function



Public Function MatrizDistribuirMonto(ByVal MatCalend As Variant, ByVal pnMontoPago As Double, _
                    ByVal psMetLiquidacion As String, Optional ByVal pnMontoIntFecha As Double = 0#) As Variant
Dim i As Integer
Dim J As Integer
Dim MatCalendDistrib As Variant
Dim nMonto As Double
Dim nMontoTemp As Double

        nMonto = pnMontoPago
        MatCalendDistrib = CrearMatrizparaAmortizacion(MatCalend)
        Do While nMonto > 0
            For J = 1 To 4
                Select Case Mid(psMetLiquidacion, J, 1)
                    Case "G"
                        Call MatrizDistribuirGastos(MatCalend, MatCalendDistrib, nMonto)
                    Case "M"
                        Call MatrizDistribuirMora(MatCalend, MatCalendDistrib, nMonto)
                    Case "I"
                        Call MatrizDistribuirInteresI(MatCalend, MatCalendDistrib, nMonto)
                    
                    Case "C"
                        Call MatrizDistribuirCapital(MatCalend, MatCalendDistrib, nMonto)
                        
                End Select
                Call MatrizActualizarEstadoCuota(MatCalend, MatCalendDistrib)
            Next J
        Loop
        MatrizDistribuirMonto = MatCalendDistrib
End Function

Public Function MatrizSaldoCapital(ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant) As Double
Dim i As Integer
    MatrizSaldoCapital = 0
    For i = 0 To UBound(MatCalend) - 1
        MatrizSaldoCapital = MatrizSaldoCapital + (CDbl(MatCalend(i, 3)) - CDbl(MatCalendDistrib(i, 3)))
        MatrizSaldoCapital = CDbl(Format(MatrizSaldoCapital, "#0.00"))
    Next i
End Function

Public Function MatrizSaldoCapitalMiVivienda(ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant, ByVal MatCalendDistribParalelo As Variant) As Double
Dim i As Integer

    MatrizSaldoCapitalMiVivienda = 0
    For i = 0 To UBound(MatCalend) - 1
        MatrizSaldoCapitalMiVivienda = MatrizSaldoCapitalMiVivienda + (CDbl(MatCalend(i, 3)) - CDbl(MatCalendDistrib(i, 3)))
        MatrizSaldoCapitalMiVivienda = CDbl(Format(MatrizSaldoCapitalMiVivienda, "#0.00"))
    Next i
    MatrizSaldoCapitalMiVivienda = MatrizSaldoCapitalMiVivienda - MatrizCapitalPagado(MatCalendDistribParalelo)
    
End Function

Public Function MatrizMontoCuota(ByVal MatCalend As Variant, ByVal pnNroCuota As Integer) As Double
Dim i As Integer
    MatrizMontoCuota = 0
    For i = 0 To UBound(MatCalend) - 1
        If CInt(MatCalend(i, 1)) = pnNroCuota Then
            MatrizMontoCuota = MatrizMontoCuota + CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5))
            MatrizMontoCuota = CDbl(Format(MatrizMontoCuota, "#0.00"))
            Exit Function
        End If
    Next i
End Function


Public Function MatrizCuotaPendiente(ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant) As Integer
Dim i As Integer
    MatrizCuotaPendiente = 0
    For i = 0 To UBound(MatCalend) - 1
        If CInt(MatCalendDistrib(i, 2)) = gColocCalendEstadoPendiente Then
            MatrizCuotaPendiente = CInt(MatCalendDistrib(i, 1))
            Exit For
        End If
    Next i
End Function
Public Function MatrizCapitalPagado(ByVal MatCalendDistrib As Variant) As Double
Dim i As Integer
    MatrizCapitalPagado = 0
    For i = 0 To UBound(MatCalendDistrib) - 1
        MatrizCapitalPagado = MatrizCapitalPagado + CDbl(MatCalendDistrib(i, 3))
    Next i
End Function

Public Function MatrizIntCompPagado(ByVal MatCalendDistrib As Variant) As Double
Dim i As Integer
    MatrizIntCompPagado = 0
    For i = 0 To UBound(MatCalendDistrib) - 1
        MatrizIntCompPagado = MatrizIntCompPagado + CDbl(MatCalendDistrib(i, 4))
    Next i
End Function

Public Function MatrizIntCompVencPagado(ByVal MatCalendDistrib As Variant) As Double
Dim i As Integer
    MatrizIntCompVencPagado = 0
    For i = 0 To UBound(MatCalendDistrib) - 1
        MatrizIntCompVencPagado = MatrizIntCompVencPagado + CDbl(MatCalendDistrib(i, 11))
    Next i
End Function

Public Function MatrizIntGraciaPagado(ByVal MatCalendDistrib As Variant) As Double
Dim i As Integer
    MatrizIntGraciaPagado = 0
    For i = 0 To UBound(MatCalendDistrib) - 1
        MatrizIntGraciaPagado = MatrizIntGraciaPagado + CDbl(MatCalendDistrib(i, 5))
    Next i
End Function

Public Function MatrizIntMorPagado(ByVal MatCalendDistrib As Variant) As Double
Dim i As Integer
    MatrizIntMorPagado = 0
    For i = 0 To UBound(MatCalendDistrib) - 1
        MatrizIntMorPagado = MatrizIntMorPagado + CDbl(MatCalendDistrib(i, 6))
    Next i
End Function

Public Function MatrizIntReprogPag(ByVal MatCalendDistrib As Variant) As Double
Dim i As Integer
    MatrizIntReprogPag = 0
    For i = 0 To UBound(MatCalendDistrib) - 1
        MatrizIntReprogPag = MatrizIntReprogPag + CDbl(MatCalendDistrib(i, 7))
    Next i
End Function

Public Function MatrizIntSuspensoPag(ByVal MatCalendDistrib As Variant) As Double
Dim i As Integer
    MatrizIntSuspensoPag = 0
    For i = 0 To UBound(MatCalendDistrib) - 1
        MatrizIntSuspensoPag = MatrizIntSuspensoPag + CDbl(MatCalendDistrib(i, 8))
    Next i
End Function

Public Function MatrizGastoPag(ByVal MatCalendDistrib As Variant) As Double
Dim i As Integer
    MatrizGastoPag = 0
    For i = 0 To UBound(MatCalendDistrib) - 1
        MatrizGastoPag = MatrizGastoPag + CDbl(MatCalendDistrib(i, 9))
    Next i
End Function

Public Function MatrizCuotasPagadas(ByVal MatCalendDistrib As Variant) As String
Dim i As Integer
    MatrizCuotasPagadas = ""
    For i = 0 To UBound(MatCalendDistrib) - 1
        If CInt(MatCalendDistrib(i, 2)) = gColocCalendEstadoPagado Then
            MatrizCuotasPagadas = MatrizCuotasPagadas & MatCalendDistrib(i, 1) & ","
        End If
    Next i
    If Len(MatrizCuotasPagadas) > 0 Then
        MatrizCuotasPagadas = Mid(MatrizCuotasPagadas, 1, Len(MatrizCuotasPagadas) - 1)
    Else
        MatrizCuotasPagadas = "0"
    End If
    
End Function
Public Function MatrizFechaCuotaPendiente(ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant) As Date
Dim i As Integer
    For i = 0 To UBound(MatCalend) - 1
        If CInt(MatCalendDistrib(i, 2)) = gColocCalendEstadoPendiente Then
            MatrizFechaCuotaPendiente = CDate(MatCalendDistrib(i, 0))
            Exit For
        End If
    Next i
End Function

Public Function MatrizEstadoCalendario(ByVal MatCalendDistrib As Variant) As Integer
Dim i As Integer
    MatrizEstadoCalendario = gColocCalendEstadoPagado
    For i = 0 To UBound(MatCalendDistrib) - 1
        If CInt(MatCalendDistrib(i, 2)) = gColocCalendEstadoPendiente Then
            MatrizEstadoCalendario = gColocCalendEstadoPendiente
            Exit For
        End If
    Next i
End Function

Public Function CrearMatrizparaAmortizacion(ByVal MatCalend As Variant) As Variant
Dim MatCalendAmortiz() As String
Dim i As Integer
    ReDim MatCalendAmortiz(UBound(MatCalend), 12)
    
    For i = 0 To UBound(MatCalend) - 1
        MatCalendAmortiz(i, 0) = MatCalend(i, 0)
        MatCalendAmortiz(i, 1) = MatCalend(i, 1)
        MatCalendAmortiz(i, 2) = MatCalend(i, 2)
        MatCalendAmortiz(i, 3) = "0.00"
        MatCalendAmortiz(i, 4) = "0.00"
        MatCalendAmortiz(i, 5) = "0.00"
        MatCalendAmortiz(i, 6) = "0.00"
        MatCalendAmortiz(i, 7) = "0.00"
        MatCalendAmortiz(i, 8) = "0.00"
        MatCalendAmortiz(i, 9) = "0.00"
        MatCalendAmortiz(i, 11) = "0.00"
        MatCalendAmortiz(i, 10) = MatCalend(i, 10)
    Next i
    CrearMatrizparaAmortizacion = MatCalendAmortiz
End Function

Function RecuperaMatrizRefinanciados(ByVal psCtaCod As String) As Variant
Dim oCalend As Dcalendario
Dim R As ADODB.Recordset
Dim MatCalend() As String
    On Error GoTo ErrorRecuperaMatrizRefinanciados
    Set oCalend = New Dcalendario
    Set R = oCalend.RecuperaRefinanciados(psCtaCod)
    If Not R.BOF And Not R.EOF Then
        ReDim MatCalend(R.RecordCount, 17)
    Else
        ReDim MatCalend(0, 0)
    End If
    Do While Not R.EOF
        MatCalend(R.Bookmark - 1, 0) = R!cCtaCod
        MatCalend(R.Bookmark - 1, 1) = R!cCtaCodRef
        MatCalend(R.Bookmark - 1, 2) = R!nMontoRef
        MatCalend(R.Bookmark - 1, 3) = Format(IIf(IsNull(R!nCapital), 0, R!nCapital), "#0.00")
        MatCalend(R.Bookmark - 1, 4) = Format(IIf(IsNull(R!nCapitalPag), 0, R!nCapitalPag), "#0.00")
        MatCalend(R.Bookmark - 1, 5) = Format(IIf(IsNull(R!nIntComp), 0, R!nIntComp), "#0.00")
        MatCalend(R.Bookmark - 1, 6) = Format(IIf(IsNull(R!nIntCompPag), 0, R!nIntCompPag), "#0.00")
        MatCalend(R.Bookmark - 1, 7) = Format(IIf(IsNull(R!nIntGracia), 0, R!nIntGracia), "#0.00")
        MatCalend(R.Bookmark - 1, 8) = Format(IIf(IsNull(R!nIntGraciaPag), 0, R!nIntGraciaPag), "#0.00")
        MatCalend(R.Bookmark - 1, 9) = Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00")
        MatCalend(R.Bookmark - 1, 10) = Format(IIf(IsNull(R!nIntMorPag), 0, R!nIntMorPag), "#0.00")
        MatCalend(R.Bookmark - 1, 11) = Format(IIf(IsNull(R!nIntReprog), 0, R!nIntReprog), "#0.00")
        MatCalend(R.Bookmark - 1, 12) = Format(IIf(IsNull(R!nIntReprogPag), 0, R!nIntReprogPag), "#0.00")
        MatCalend(R.Bookmark - 1, 13) = Format(IIf(IsNull(R!nIntSuspenso), 0, R!nIntSuspenso), "#0.00")
        MatCalend(R.Bookmark - 1, 14) = Format(IIf(IsNull(R!nIntSuspensoPag), 0, R!nIntSuspensoPag), "#0.00")
        MatCalend(R.Bookmark - 1, 15) = Format(IIf(IsNull(R!nGastos), 0, R!nGastos), "#0.00")
        MatCalend(R.Bookmark - 1, 16) = Format(IIf(IsNull(R!nGastosPag), 0, R!nGastosPag), "#0.00")
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set oCalend = Nothing
    RecuperaMatrizRefinanciados = MatCalend
    Exit Function

ErrorRecuperaMatrizRefinanciados:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Function RecuperaMatrizCalendarioInicial(ByVal psCtaCod As String, Optional nMontoInicial As Double = -1, _
    Optional ByVal pbCalendParalelo As Boolean = False, Optional ByVal pbParaImpresion As Boolean = False) As Variant
Dim R As ADODB.Recordset
Dim RCred As ADODB.Recordset
Dim oCalend As Dcalendario
Dim oCredito As DCredito
Dim MatCalend() As String
Dim nMontoSaldo As Double
Dim i As Integer

    On Error GoTo ErrorRecuperaMatrizCalendarioInicial
    Set oCalend = New Dcalendario
    Set R = oCalend.RecuperaCalendarioPagos(psCtaCod, , pbCalendParalelo, , True)
    Set oCalend = Nothing
    ReDim MatCalend(R.RecordCount, 11)
    Set oCredito = New DCredito
    Set RCred = oCredito.RecuperaColocaciones(psCtaCod)
    If nMontoInicial <> -1 Then
        nMontoSaldo = nMontoInicial
    Else
        nMontoSaldo = CDbl(Format(RCred!nMontoCol, "#0.00"))
    End If
    Set oCredito = Nothing
    Do While Not R.EOF
        MatCalend(R.Bookmark - 1, 0) = Format(R!DVENC, "dd/mm/yyyy")
        MatCalend(R.Bookmark - 1, 1) = Trim(Str(R!nCuota))
        MatCalend(R.Bookmark - 1, 2) = Trim(Str(R!nColocCalendEstado))
        MatCalend(R.Bookmark - 1, 3) = Format(IIf(IsNull(R!nCapital), 0, R!nCapital), "#0.00")
        MatCalend(R.Bookmark - 1, 4) = Format(IIf(IsNull(R!nIntComp), 0, R!nIntComp), "#0.00")
        MatCalend(R.Bookmark - 1, 5) = Format(IIf(IsNull(R!nIntGracia), 0, R!nIntGracia), "#0.00")
        'MatCalend(R.Bookmark - 1, 6) = Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00")
        MatCalend(R.Bookmark - 1, 6) = "0.00"
        MatCalend(R.Bookmark - 1, 7) = Format(IIf(IsNull(R!nIntReprog), 0, R!nIntReprog), "#0.00")
        'MatCalend(R.Bookmark - 1, 8) = Format(IIf(IsNull(R!nIntSuspenso), 0, R!nIntSuspenso), "#0.00")
        MatCalend(R.Bookmark - 1, 8) = "0.00"
        MatCalend(R.Bookmark - 1, 9) = Format(IIf(IsNull(R!NGasto), 0, R!NGasto), "#0.00")
        nMontoSaldo = nMontoSaldo - IIf(IsNull(R!nCapital), 0, R!nCapital)
        nMontoSaldo = CDbl(Format(nMontoSaldo, "#0.00"))
        MatCalend(R.Bookmark - 1, 10) = Format(nMontoSaldo, "#0.00")
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    If pbParaImpresion And pbCalendParalelo Then
        Set oCredito = New DCredito
        Set RCred = oCredito.RecuperaDatosComunes(psCtaCod, True, Array(gColocEstSolic, gColocEstSug, gColocEstAprob, gColocEstVigMor, gColocEstVigNorm, gColocEstVigVenc, gColocEstRefMor, gColocEstRefNorm, gColocEstRefVenc))
        If Not IsNull(RCred!nMontoCol) Then
            nMontoSaldo = CDbl(Format(RCred!nMontoCol, "#0.00"))
        End If
        If nMontoSaldo = 0 Then
            nMontoSaldo = RCred!nMontoSol
        End If
        Set oCredito = Nothing
        
        Set oCalend = New Dcalendario
        Set R = oCalend.RecuperaCalendarioPagos(psCtaCod)
        Set oCalend = Nothing
        Do While Not R.EOF
            MatCalend(R.Bookmark - 1, 0) = Format(R!DVENC, "dd/mm/yyyy")
            MatCalend(R.Bookmark - 1, 1) = Trim(Str(R!nCuota))
            MatCalend(R.Bookmark - 1, 2) = Trim(Str(R!nColocCalendEstado))
            MatCalend(R.Bookmark - 1, 3) = Format(CDbl(MatCalend(R.Bookmark - 1, 3)) + IIf(IsNull(R!nCapital), 0, R!nCapital), "#0.00")
            MatCalend(R.Bookmark - 1, 4) = Format(CDbl(MatCalend(R.Bookmark - 1, 4)) + IIf(IsNull(R!nIntComp), 0, R!nIntComp), "#0.00")
            MatCalend(R.Bookmark - 1, 5) = Format(CDbl(MatCalend(R.Bookmark - 1, 5)) + IIf(IsNull(R!nIntGracia), 0, R!nIntGracia), "#0.00")
            MatCalend(R.Bookmark - 1, 6) = Format(CDbl(MatCalend(R.Bookmark - 1, 6)) + IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00")
            MatCalend(R.Bookmark - 1, 7) = Format(CDbl(MatCalend(R.Bookmark - 1, 7)) + IIf(IsNull(R!nIntReprog), 0, R!nIntReprog), "#0.00")
            'MatCalend(R.Bookmark - 1, 8) = Format(IIf(IsNull(R!nIntSuspenso), 0, R!nIntSuspenso), "#0.00")
            MatCalend(R.Bookmark - 1, 8) = "0.00"
            MatCalend(R.Bookmark - 1, 9) = Format(CDbl(MatCalend(R.Bookmark - 1, 8)) + IIf(IsNull(R!NGasto), 0, R!NGasto), "#0.00")
            nMontoSaldo = nMontoSaldo - CDbl(MatCalend(R.Bookmark - 1, 3))
            nMontoSaldo = CDbl(Format(nMontoSaldo, "#0.00"))
            MatCalend(R.Bookmark - 1, 10) = Format(nMontoSaldo, "#0.00")
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        
        
    End If
    
    RecuperaMatrizCalendarioInicial = MatCalend
    Exit Function

ErrorRecuperaMatrizCalendarioInicial:
    ReDim MatCalend(0, 0)
    RecuperaMatrizCalendarioInicial = MatCalend
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function


Function RecuperaMatrizGastosPendiente(ByVal psCtaCod As String) As Variant
Dim oCredito As Dcalendario
Dim MatGastos() As String
Dim R As ADODB.Recordset

    Set oCredito = New Dcalendario
    Set R = oCredito.RecuperaCalendarioGastosPendientes(psCtaCod, gColocCalendAplCuota)
    Set oCredito = Nothing
    'Cuota, Concepto, Monto
    ReDim MatGastos(0)
    If R.RecordCount > 0 Then
        ReDim MatGastos(R.RecordCount, 4)
    End If
    Do While Not R.EOF
        MatGastos(R.Bookmark - 1, 0) = Trim(Str(R!nCuota))
        MatGastos(R.Bookmark - 1, 1) = Trim(Str(R!nPrdConceptoCod))
        MatGastos(R.Bookmark - 1, 2) = Format(R!nMonto - R!nMontoPagado, "#0.00")
        MatGastos(R.Bookmark - 1, 3) = psCtaCod
        R.MoveNext
    Loop
    R.Close
    
    RecuperaMatrizGastosPendiente = MatGastos
    
End Function

Function RecuperaMatrizGastosPendienteAgrupado(ByVal psCtaCod As String) As Variant
Dim oCredito As Dcalendario
Dim MatGastos() As String
Dim R As ADODB.Recordset

    Set oCredito = New Dcalendario
    Set R = oCredito.RecuperaCalendarioGastosPendientesAgrupado(psCtaCod)
    Set oCredito = Nothing
    'Concepto, Monto
    ReDim MatGastos(0)
    If R.RecordCount > 0 Then
        ReDim MatGastos(R.RecordCount, 2)
    End If
    Do While Not R.EOF
        MatGastos(R.Bookmark - 1, 0) = Trim(Str(R!nPrdConceptoCod))
        MatGastos(R.Bookmark - 1, 1) = Format(R!nMonto, "#0.00")
        
        R.MoveNext
    Loop
    R.Close
    
    RecuperaMatrizGastosPendienteAgrupado = MatGastos
    
End Function

Function RecuperaMatrizCalendarioPendiente(ByVal psCtaCod As String, Optional ByVal pbParalelo As Boolean = False) As Variant
Dim R As ADODB.Recordset
Dim RTmp As ADODB.Recordset
Dim RCred As ADODB.Recordset
Dim oCalend As Dcalendario
Dim oCredito As DCredito
Dim MatCalend() As String
Dim nMontoSaldo As Double
Dim oGen As dFeriado
Dim nNumDiasfer As Integer
Dim dFecSisTmp As Date
Dim oCons As NConstSistemas
Dim bFeriado As Boolean
Dim i As Integer
Dim nIntPend, nColocCalendCod, nIntFecha  As Double
Dim dVigencia As Date
Dim nSaldo As Double
Dim nDiasTranscurridos As Integer
Dim nTasaInt As Double


    On Error GoTo ErrorRecuperaMatrizCalendarioPendiente
    
    Set oGen = New dFeriado
    Set oCons = New NConstSistemas
    dFecSisTmp = CDate(oCons.LeeConstSistema(gConstSistFechaSistema))
    Set oCons = Nothing
    
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaColocacCred(psCtaCod)
        
    nIntPend = R!nIntPend
    nColocCalendCod = R!nColocCalendCod
    R.Close
    
    Set R = oCredito.RecuperaColocaciones(psCtaCod)
    dVigencia = R!dVigencia
    R.Close
    
    Set R = oCredito.RecuperaProducto(psCtaCod)
    nSaldo = R!nSaldo
    nTasaInt = R!nTasaInteres
    Set oCredito = Nothing
    R.Close
    
    Set oCalend = New Dcalendario
    Set R = oCalend.RecuperaCalendarioPagosPendiente(psCtaCod, pbParalelo)
    Set oCalend = Nothing
    ReDim MatCalend(R.RecordCount, 13)
    If R.RecordCount > 0 Then
        Set oCredito = New DCredito
        'Set RCred = oCredito.RecuperaColocaciones(psCtaCod)
        'nMontoSaldo = CDbl(Format(RCred!nMontoCol, "#0.00"))
        nMontoSaldo = CDbl(Format(R!nMontoPrestamo, "#0.00"))
        'Set oCredito = Nothing
        Do While Not R.EOF
            
            MatCalend(R.Bookmark - 1, 0) = Format(R!DVENC, "dd/mm/yyyy")
            MatCalend(R.Bookmark - 1, 1) = Trim(Str(R!nCuota))
            MatCalend(R.Bookmark - 1, 2) = Trim(Str(R!nColocCalendEstado))
            MatCalend(R.Bookmark - 1, 3) = Format(IIf(IsNull(R!nCapital), 0, R!nCapital), "#0.00")
            MatCalend(R.Bookmark - 1, 4) = Format(IIf(IsNull(R!nIntComp), 0, R!nIntComp), "#0.00")
            MatCalend(R.Bookmark - 1, 5) = Format(IIf(IsNull(R!nIntGracia), 0, R!nIntGracia), "#0.00")
            If dFecSisTmp > R!DVENC Then
               bFeriado = False
                'Si se le cobra Mora
                For i = 0 To DateDiff("d", R!DVENC, dFecSisTmp) - 1
                                        
                    Set RTmp = oGen.DetallaFeriado(R!DVENC + i)
                    If RTmp.RecordCount > 0 Then
                        bFeriado = True
                    Else
                        bFeriado = False
                        Exit For
                    End If
                    RTmp.Close
                    
                    If i = 2 Then
                        Exit For
                    End If
                Next i
                
                If bFeriado Then
                    MatCalend(R.Bookmark - 1, 6) = "0.00"
                Else
                    MatCalend(R.Bookmark - 1, 6) = Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00")
                End If
            Else
                MatCalend(R.Bookmark - 1, 6) = Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00")
            End If
            MatCalend(R.Bookmark - 1, 7) = Format(IIf(IsNull(R!nIntReprog), 0, R!nIntReprog), "#0.00")
            MatCalend(R.Bookmark - 1, 8) = Format(IIf(IsNull(R!nIntSuspenso), 0, R!nIntSuspenso), "#0.00")
            MatCalend(R.Bookmark - 1, 9) = Format(IIf(IsNull(R!NGasto), 0, R!NGasto), "#0.00")
            MatCalend(R.Bookmark - 1, 12) = Format(nMontoSaldo, "#0.00")
            nMontoSaldo = nMontoSaldo - IIf(IsNull(R!nSaldoCap), 0, R!nSaldoCap)
            nMontoSaldo = CDbl(Format(nMontoSaldo, "#0.00"))
            MatCalend(R.Bookmark - 1, 10) = Format(nMontoSaldo, "#0.00")
            MatCalend(R.Bookmark - 1, 11) = Format(IIf(IsNull(R!nIntCompVenc), 0, R!nIntCompVenc), "#0.00")
            MatCalend(R.Bookmark - 1, 12) = Format(IIf(IsNull(R!nITF), 0, R!nITF), "#0.00")
            R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
    Set oGen = Nothing
    
    'Si es cuota libre hallar Interes a la Fecha
    If nColocCalendCod = 70 Then
        Set oCredito = New DCredito
        Set R = oCredito.RecuperaUltimoMovimiento(psCtaCod)
        Set oCredito = Nothing
        If R.RecordCount > 0 Then
            dVigencia = R!dFecPago
        End If
        R.Close
        nDiasTranscurridos = DateDiff("d", dVigencia, dFecSisTmp)
        If nDiasTranscurridos > 0 Then
            nIntFecha = (TasaIntPerDias(nTasaInt, nDiasTranscurridos) * nSaldo) + nIntPend
        Else
            nIntFecha = nIntPend
        End If
        
        For i = 0 To UBound(MatCalend) - 1
            If i = 0 Then
                MatCalend(i, 4) = Format(nIntFecha, "#0.00")
            Else
                MatCalend(i, 4) = "0.00"
            End If
        Next i
        
    End If
    
    RecuperaMatrizCalendarioPendiente = MatCalend
    Exit Function

ErrorRecuperaMatrizCalendarioPendiente:
    ReDim MatCalend(0, 0)
    RecuperaMatrizCalendarioPendiente = MatCalend
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function MatrizGastosVencidos(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim i As Integer
    MatrizGastosVencidos = 0
    If MatCalend(0, 9) = "" Then
        MatCalend(0, 9) = "0.00"
    End If
    
    MatrizGastosVencidos = MatrizGastosVencidos + CDbl(MatCalend(0, 9))
    For i = 1 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(i, 0)) Then
            MatrizGastosVencidos = MatrizGastosVencidos + CDbl(MatCalend(i, 9))
        End If
    Next i
End Function

Public Function MatrizIntCompVencido(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim i As Integer
    MatrizIntCompVencido = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(i, 0)) Then
            MatrizIntCompVencido = MatrizIntCompVencido + CDbl(MatCalend(i, 4))
        End If
    Next i
End Function

Public Function MatrizIntCompVencVencido(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim i As Integer
    MatrizIntCompVencVencido = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(i, 0)) Then
            MatrizIntCompVencVencido = MatrizIntCompVencVencido + CDbl(MatCalend(i, 11))
        End If
    Next i
End Function

Public Function MatrizIntReprogVencido(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim i As Integer
    MatrizIntReprogVencido = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(i, 0)) Then
            MatrizIntReprogVencido = MatrizIntReprogVencido + CDbl(MatCalend(i, 7))
        End If
    Next i
End Function

Public Function MatrizIntSuspensoVencido(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim i As Integer
    MatrizIntSuspensoVencido = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(i, 0)) Then
            MatrizIntSuspensoVencido = MatrizIntSuspensoVencido + CDbl(MatCalend(i, 8))
        End If
    Next i
End Function

Public Function MatrizIntGraciaVencido(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim i As Integer
    MatrizIntGraciaVencido = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(i, 0)) Then
            MatrizIntGraciaVencido = MatrizIntGraciaVencido + CDbl(MatCalend(i, 5))
        End If
    Next i
End Function

Public Function MatrizIntCompVencidoCalendCierre(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim i As Integer
    MatrizIntCompVencidoCalendCierre = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(i, 0)) Then
            MatrizIntCompVencidoCalendCierre = MatrizIntCompVencidoCalendCierre + CDbl(MatCalend(i, 11))
        End If
    Next i
End Function

Public Function MatrizIntCompVencidoCierre(ByRef MatCalend As Variant, ByVal pdHoy As Date, ByVal pnTasa As Double, ByVal pnMontoAcum As Double) As Double
Dim i As Integer
Dim nIntTempo As Double

    MatrizIntCompVencidoCierre = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(i, 0)) Then
            nIntTempo = CDbl(Format((((1 + (pnTasa / 100)) ^ (1 / 30)) - 1), "#0.000000")) * CDbl(MatCalend(i, 3))
            MatCalend(i, 3) = Format(CDbl(MatCalend(i, 3)) + nIntTempo, "#0.000000")
            MatrizIntCompVencidoCierre = MatrizIntCompVencidoCierre + nIntTempo
        End If
    Next i
    MatrizIntCompVencidoCierre = MatrizIntCompVencidoCierre + pnMontoAcum
End Function


Public Function MatrizCapitalVencido(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim i As Integer
    MatrizCapitalVencido = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(i, 0)) Then
            MatrizCapitalVencido = MatrizCapitalVencido + CDbl(MatCalend(i, 3))
        End If
    Next i
End Function

Public Function MatrizDeudaAlaFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double

    'MatrizDeudaAlaFecha = MatrizCapitalVencido(MatCalend, pdHoy) + MatrizInteresGastosAFecha(psCtaCod, MatCalend, pdHoy)
    MatrizDeudaAlaFecha = MatrizCapitalCalendario(MatCalend) + MatrizInteresGastosAFecha(psCtaCod, MatCalend, pdHoy)
End Function

Public Function MatrizMoraTotal(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim i As Integer
    MatrizMoraTotal = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(0, 0)) Then
            MatrizMoraTotal = MatrizMoraTotal + CDbl(MatCalend(i, 6))
        End If
    Next i
End Function

Public Function MatrizCuotasEnMora(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Integer
Dim i As Integer
    MatrizCuotasEnMora = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy > CDate(MatCalend(i, 0)) Then
            MatrizCuotasEnMora = MatrizCuotasEnMora + 1
        End If
    Next i
End Function

Public Function MatrizMontoCalendDinamico(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date, Optional ByVal pnMivivienda As Integer = 0)
    
    MatrizMontoCalendDinamico = 0
    'MatrizMontoCalendDinamico = MatrizMontoCalendDinamico + MatrizCapitalVencido(MatCalend, pdHoy) _
    '    + MatrizGastosVencidos(MatCalend, pshoy) + MatrizInteresTotalesAFecha(psCtaCod, MatCalend, pdHoy)
    If pnMivivienda = 1 Then
        MatrizMontoCalendDinamico = MatrizMontoCalendDinamico + MatrizCapitalVencido(MatCalend, pdHoy) _
            + MatrizGastosVencidos(MatCalend, gdFecSis) + MatrizInteresTotalesAFecha(psCtaCod, MatCalend, pdHoy)
    Else
        MatrizMontoCalendDinamico = MatrizMontoCalendDinamico + MatrizCapitalVencido(MatCalend, pdHoy) _
            + MatrizGastosFecha(psCtaCod, MatCalend) + MatrizInteresTotalesAFecha(psCtaCod, MatCalend, pdHoy)
    End If
End Function

'**********************************************************************
'* Funcion que trabaja conla Matriz de Calendario ya cargado
'*  y calcula va  a Capital de lo que tiene que pagar el Cliente para Ponerse al Dia
'***********************************************************************
Public Function MatrizMontoCapitalAPagar(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim i As Integer
Dim J As Integer
    MatrizMontoCapitalAPagar = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(i, 0)) Then
            For J = 3 To 9
                MatrizMontoCapitalAPagar = MatrizMontoCapitalAPagar + CDbl(MatCalend(i, 3))
            Next J
        End If
    Next i
    If MatrizMontoCapitalAPagar = 0 Then
        MatrizMontoCapitalAPagar = MatrizMontoCapitalAPagar + CDbl(MatCalend(0, 3))
    End If
End Function
'**********************************************************************
'* Funcion que trabaja conla Matriz de Calendario ya cargado
'*  y calcula lo que tiene que pagar el Cliente para Ponerse al Dia
'***********************************************************************

Public Function DistribuirMatrizMiVivEnDosCalendarios(ByRef pMatCalendDistribParalelo As Variant, _
    ByRef pMatCalendDistrib As Variant, ByVal pMatResult As Variant, ByVal pMatCalendParalelo As Variant, _
    ByVal pMatCalend As Variant, ByVal nCalPago As Integer)
    
Dim i, J As Integer
Dim k As Integer
Dim m As Integer
Dim nMontoPago As Double

    '*******************************************************************
    'Si es Mi vivienda Distribuir Amortizacion
    '*******************************************************************
    
    'Primero gastos al calandario no consecional
        
    For i = 0 To UBound(pMatResult) - 1
        For m = 3 To 11
            If m <> 10 Then
                nMontoPago = nMontoPago + CDbl(pMatResult(i, m))
                If nMontoPago > 0 Then
                    
                    'If nCalPago = 0 Then 'Si es Mal Pagador
                        'Capital, e Intereses del paralelo
                        For k = 0 To UBound(pMatCalendDistribParalelo) - 1
                            If pMatCalendDistribParalelo(k, 1) = pMatResult(i, 1) Then
                                For J = 3 To 11
                                    If J = m Then
                                        If nMontoPago > CDbl(pMatCalendParalelo(k, J)) Then
                                            pMatCalendDistribParalelo(k, J) = pMatCalendParalelo(k, J)
                                            nMontoPago = nMontoPago - CDbl(pMatCalendParalelo(k, J))
                                        Else
                                            pMatCalendDistribParalelo(k, J) = Format(nMontoPago, "#0.00")
                                            nMontoPago = 0
                                        End If
                                    End If
                                Next J
                                Exit For
                            End If
                        Next k
                    'End If
                
                    'Capital, e Intereses
                    For k = 0 To UBound(pMatCalendDistrib) - 1
                        If pMatCalendDistrib(k, 1) = pMatResult(i, 1) Then
                            For J = 3 To 11
                                If J = m Then
                                    If nMontoPago > CDbl(pMatCalend(k, J)) Then
                                        pMatCalendDistrib(k, J) = pMatCalend(k, J)
                                        nMontoPago = nMontoPago - CDbl(pMatCalend(k, J))
                                    Else
                                        pMatCalendDistrib(k, J) = Format(nMontoPago, "#0.00")
                                        nMontoPago = 0
                                    End If
                                    If nMontoPago > 0 Then
                                        pMatCalendDistrib(k, J) = Format(CDbl(pMatCalendDistrib(k, J)) + nMontoPago, "#0.00")
                                        nMontoPago = 0
                                    End If
                                End If
                            Next J
                            Exit For
                        End If
                    Next k
                End If
            End If
        Next m
    Next i
    Call MatrizActualizarEstadoCuota(pMatCalendParalelo, pMatCalendDistribParalelo)
    Call MatrizActualizarEstadoCuota(pMatCalend, pMatCalendDistrib)
    
End Function

Public Function MatrizMontoPagado(ByVal MatCalendDistrib As Variant, _
    Optional ByVal pnCuota As Integer = -1) As Double
Dim i, J As Integer
    MatrizMontoPagado = 0
    
    If pnCuota <> -1 Then
        For i = 0 To UBound(MatCalendDistrib) - 1
            If CInt(MatCalendDistrib(i, 1)) = pnCuota Then
                For J = 3 To 9
                    MatrizMontoPagado = MatrizMontoPagado + MatCalendDistrib(i, J)
                Next J
                MatrizMontoPagado = MatrizMontoPagado + MatCalendDistrib(i, 11)
                Exit For
            End If
        Next i
    Else
        For i = 0 To UBound(MatCalendDistrib) - 1
            For J = 3 To 9
                MatrizMontoPagado = MatrizMontoPagado + MatCalendDistrib(i, J)
            Next J
            MatrizMontoPagado = MatrizMontoPagado + MatCalendDistrib(i, 11)
        Next i
    End If
End Function

Public Function MatrizCalculoMoraSimuladorPagos(ByVal MatCalend As Variant, ByVal pdHoy As Date, _
    ByVal pdFin As Date, ByVal pnTasaintMor As Double) As Double
Dim i As Integer
    MatrizCalculoMoraSimuladorPagos = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(i, 0)) Then
            If pdFin > pdHoy Then
                MatrizCalculoMoraSimuladorPagos = MatrizCalculoMoraSimuladorPagos + Format(((pnTasaintMor / 100) * (pdFin - pdHoy)) * MatCalend(i, 3), "#0.00")
            Else
                MatrizCalculoMoraSimuladorPagos = 0
            End If
        Else
            If pdFin >= CDate(MatCalend(i, 0)) Then
                MatrizCalculoMoraSimuladorPagos = MatrizCalculoMoraSimuladorPagos + Format(((pnTasaintMor / 100) * (pdFin - CDate(MatCalend(i, 0)))) * MatCalend(i, 3), "#0.00")
            End If
        End If
    Next i
End Function

Public Function MatrizMontoAPagar(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim i As Integer
Dim J As Integer
    MatrizMontoAPagar = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(i, 0)) Then
            For J = 3 To 9
                MatrizMontoAPagar = MatrizMontoAPagar + CDbl(MatCalend(i, J))
            Next J
            MatrizMontoAPagar = MatrizMontoAPagar + CDbl(MatCalend(i, 11))
        End If
    Next i
    If MatrizMontoAPagar = 0 Then
        For J = 3 To 9
            MatrizMontoAPagar = MatrizMontoAPagar + CDbl(MatCalend(0, J))
        Next J
        MatrizMontoAPagar = MatrizMontoAPagar + CDbl(MatCalend(0, 11))
    End If
End Function


Public Function MatrizMontoAPagarCuotaPendiente(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim i As Integer
Dim J As Integer
    MatrizMontoAPagarCuotaPendiente = 0
    For J = 3 To 9
            MatrizMontoAPagarCuotaPendiente = MatrizMontoAPagarCuotaPendiente + CDbl(MatCalend(0, J))
    Next J
    MatrizMontoAPagarCuotaPendiente = MatrizMontoAPagarCuotaPendiente + CDbl(MatCalend(0, 11))
    
End Function

Private Function ValidaAmortizacionCredito(ByVal pnMonto As Double) As String
    ValidaAmortizacionCredito = ""
    If pnMonto = 0 Then
        ValidaAmortizacionCredito = "Monto de Pago debe ser Mayor que Cero"
    End If
End Function

Public Function ValidaMontoPrestamo(ByVal psCtaCod As String, ByVal pnMontoPrestamo As Double, ByVal pdFecha As Date) As String
Dim oGeneral As DGeneral
Dim nTipoCambioFijo As Double
Dim nMontoMN As Double
Dim nMontoME As Double

Dim nMontoMaxRiesgosExt As Double
Dim nMontoMaxPolizaSegExt As Double
Dim oCredGen As DCredGeneral
Dim RTemp As ADODB.Recordset

    On Error GoTo ErrorValidaMontoPrestamo
    
    ' Obtiene el Parametro del Monto Max del Credito, para solicitar el reporte de Riesgos
    Set oCredGen = New DCredGeneral
    Set RTemp = oCredGen.RecuperaParametro(3054)
    Set oCredGen = Nothing
    nMontoMaxRiesgosExt = IIf(IsNull(RTemp!nParamValor), 0, RTemp!nParamValor)
    RTemp.Close
    Set RTemp = Nothing
        
    ' Obtiene el Parametro del Monto Max del Credito, para solicitar Poliza de Seguros
    Set oCredGen = New DCredGeneral
    Set RTemp = oCredGen.RecuperaParametro(3055)
    Set oCredGen = Nothing
    nMontoMaxPolizaSegExt = IIf(IsNull(RTemp!nParamValor), 0, RTemp!nParamValor)
    RTemp.Close
    Set RTemp = Nothing

    ValidaMontoPrestamo = ""

    Set oGeneral = New DGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
    nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))
    Set oGeneral = Nothing

    'Moneda Nacional '
    If CInt(Mid(psCtaCod, 9, 1)) = gMonedaNacional Then
       nMontoMN = pnMontoPrestamo
       nMontoMN = CDbl(Format(pnMontoPrestamo, "#0.00"))

       nMontoME = nMontoMN / nTipoCambioFijo
    End If
    
    'Moneda Extrangera '
    If CInt(Mid(psCtaCod, 9, 1)) = gMonedaExtranjera Then
        nMontoME = pnMontoPrestamo
        nMontoME = CDbl(Format(nMontoME, "#0.00"))

    End If

    If nMontoME > nMontoMaxRiesgosExt Then
       ValidaMontoPrestamo = "El Monto del Crédito supera los $" & CStr(nMontoMaxRiesgosExt) & ", deberá presentar Informe de Riesgos." & Chr(13)
    End If
    
    If nMontoME > nMontoMaxPolizaSegExt Then
       ValidaMontoPrestamo = ValidaMontoPrestamo & "El Monto del Crédito supera los $" & CStr(nMontoMaxPolizaSegExt) & ", deberá presentar Póliza de Seguros."
    End If

        
    Exit Function

ErrorValidaMontoPrestamo:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function


Public Function ValidaMontoParaTipoCredito(ByVal psTipoCred As String, ByVal psMonedaCod As String, ByVal pnMontoPrestamo As Double, ByVal psMonedaFteCod As String, ByVal pnMontoFte As Double, ByVal pdFecha As Date) As String
Dim oGeneral As DGeneral
Dim nTipoCambioFijo As Double
Dim nMontoMN As Double
Dim nMontoME As Double
Dim nMontoMaxExt As Double
Dim sTipoCred As String

    On Error GoTo ErrorValidaMontoParaTipoCredito

    nMontoMaxExt = 30000
    nMontoME = 0
    
    ValidaMontoParaTipoCredito = ""

    Set oGeneral = New DGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
    nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))
    Set oGeneral = Nothing

    If CInt(psMonedaCod) = gMonedaNacional Then 'Moneda Nacional '
       nMontoMN = pnMontoPrestamo
       nMontoMN = CDbl(Format(nMontoMN, "#0.00"))

       nMontoME = nMontoME + (nMontoMN / nTipoCambioFijo)
    Else 'Moneda Extrangera '
       nMontoME = nMontoME + CDbl(Format(pnMontoPrestamo, "#0.00"))
    End If
    
    If CInt(psMonedaFteCod) = gMonedaNacional Then 'Moneda Nacional '
       nMontoMN = pnMontoFte
       nMontoMN = CDbl(Format(nMontoMN, "#0.00"))

       nMontoME = nMontoME + (nMontoMN / nTipoCambioFijo)
    Else 'Moneda Extrangera '
       nMontoME = nMontoME + pnMontoFte
    End If

    
    If nMontoME > nMontoMaxExt And psTipoCred = "20" Then 'De MES a COMERCIAL
       ValidaMontoParaTipoCredito = "El Monto del Crédito supera los $" & CStr(Format(nMontoMaxExt, "#0.00")) & ", deberia considerarlo como Tipo de Crédito COMERCIAL"
    End If
        
    If nMontoME < nMontoMaxExt And psTipoCred = "10" Then 'De COMERCIAL a MES
       ValidaMontoParaTipoCredito = "El Monto del Crédito no supera los $" & CStr(Format(nMontoMaxExt, "#0.00")) & ", deberia considerarlo como Tipo de Crédito MES"
    End If
        
    Exit Function

ErrorValidaMontoParaTipoCredito:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function


Private Function DevuelveMatrizGastosCredito(ByRef pnNumreg As Integer, ByVal psCtaCod As String, ByVal nNroCalen As Integer) As Variant
Dim R As ADODB.Recordset
Dim oCalend As Dcalendario
Dim MatGastosCred() As String
        Set oCalend = New Dcalendario
        Set R = oCalend.RecuperaCalendarioGastos(psCtaCod, nNroCalen, 0, gColocCalendAplCuota, True)
        Set oCalend = Nothing
        pnNumreg = 0
        pnNumreg = R.RecordCount
        If R.RecordCount > 0 Then
            ReDim MatGastosCred(R.RecordCount, 3)
            Do While Not R.EOF
                MatGastosCred(R.Bookmark - 1, 0) = Trim(Str(R!nCuota))
                MatGastosCred(R.Bookmark - 1, 1) = Trim(Str(R!nPrdConceptoCod))
                MatGastosCred(R.Bookmark - 1, 2) = Format(R!nMonto - R!nMontoPagado, "#0.00")
                R.MoveNext
            Loop
        End If
        R.Close

        DevuelveMatrizGastosCredito = MatGastosCred

End Function

Private Function DevuelveMatrizGastosCreditoCuota(pnNumreg As Integer, ByVal pnCuota As Integer, MatGastosCred As Variant, ByVal pnNumRegGasCred As Integer) As Variant
Dim i As Integer
Dim nCont As Integer
Dim MatGastosCredCuota() As String
    nCont = 0
    For i = 0 To pnNumRegGasCred - 1
        If pnCuota = CInt(MatGastosCred(i, 0)) Then
           nCont = nCont + 1
        End If
    Next i

    ReDim MatGastosCredCuota(nCont, 3)
    nCont = 0
    For i = 0 To pnNumRegGasCred - 1
        If pnCuota = CInt(MatGastosCred(i, 0)) Then
           nCont = nCont + 1
           MatGastosCredCuota(nCont - 1, 0) = MatGastosCred(i, 0)
           MatGastosCredCuota(nCont - 1, 1) = MatGastosCred(i, 1)
           MatGastosCredCuota(nCont - 1, 2) = MatGastosCred(i, 2)
        End If
    Next i
    pnNumreg = nCont
    DevuelveMatrizGastosCreditoCuota = MatGastosCredCuota
End Function

Public Function AmortizarCredito(ByVal psCtaCod, ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant, _
            ByVal pnMonto As Double, ByVal pdHoy As Date, ByVal psMetLiquid As String, _
            ByVal pnTipoPago As ColocTipoPago, ByVal psCodAge As String, ByVal psCodUser As String, _
            Optional psNroDoc As String = "", Optional ByVal pConnBase As DCredActualizaBD = Nothing, _
            Optional ByVal pnMovNro As Long = -1, Optional ByVal pnNroDacion As Integer = -1, Optional pbEnOtraCmac As Boolean = False, _
            Optional psPersCmac As String = "", Optional ByVal pnIntPend As Double = 0, Optional ByVal pnIntPendPagado As Double = 0, _
            Optional psMovnroTemp As String = "", Optional ByVal pMatGastosGen As Variant = "", Optional ByVal pnNumGastosGen As Integer = -1, _
            Optional ByVal MatCalendDistribParalelo As Variant = "", Optional ByVal pnCalPago As Integer = 1, Optional ByVal MatCalendParalelo As Variant = "", _
            Optional ByVal pnPrepago As Integer = 0, Optional psPersLavDinero As String = "", Optional pnITF As Double = 0#, Optional ByVal pnMontoDesagio As Double = 0, _
            Optional ByVal pBrfa As Boolean = False, Optional ByVal pcRFA As String) As String
            
Dim oBase As DCredActualizaBD
Dim nEstadoCred As Integer
Dim nEstadoCredTemp As Integer
Dim R As ADODB.Recordset
Dim oCred As DCredito
Dim oCalend As Dcalendario
Dim nTransacc As Long
Dim sLineaCred As String
Dim nMontoColocado As Double
Dim dFecPend As Date
Dim nDiasAtraso As Integer
Dim i, k As Integer
Dim nNroCalen As Integer
Dim nMontoGasto As Double
Dim sMovNro As String
Dim nMovNro As Long
Dim nConsCred As String
Dim pnPlazo As Integer
Dim bTran As Boolean
'Dim oFunciones As New NContFunciones
Dim dFechaTran As Date
Dim nIntPend As Double
Dim nMontoPago As Double
Dim nMontoPago_2 As Double
Dim nNroCalPar As Integer
Dim nMivivienda As Integer
Dim MatGastosCred As Variant
Dim NumregGastosCred As Integer
Dim MatGastosCuota As Variant
Dim NumRegGastosCuota As Integer
Dim nBuenPagador As Integer
Dim nPrestamo As Double
Dim CapitalPagado As Double
Dim nDiasAtrasoMov As Integer
Dim nMontoK As Double
Dim nIntDif As Double
Static bDif As Integer

Dim cLRFA As String
    On Error GoTo ErrorAmortizarPago
    AmortizarCredito = ""
    AmortizarCredito = ValidaAmortizacionCredito(pnMonto)
    If AmortizarCredito <> "" Then
        Exit Function
    End If
    bTran = False
    Set oCred = New DCredito
    Set R = oCred.RecuperaProducto(psCtaCod)
    Set oCred = Nothing
    nEstadoCred = R!nPrdEstado
    nEstadoCredTemp = R!nPrdEstado
    nTransacc = R!nTransacc
    R.Close
    
    Set oCred = New DCredito
    Set R = oCred.RecuperaColocaciones(psCtaCod)
    Set oCred = Nothing
    nPrestamo = R!nMontoCol
    R.Close
    Set R = Nothing
    
    If pnTipoPago <> gColocTipoPagoDacionPago Then
        If pnTipoPago <> gColocTipoPagoCargoCta Then
            'Definir Codigo de Operacion de Pago
            Select Case nEstadoCred
                'Si es Credito refinanciado
                Case gColocEstRefMor
                    If Not pbEnOtraCmac Then
                        nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefMorEfec, gCredPagRefMorCh)
                    Else
                        nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefMorEOCEfec, gCredPagRefMorEOCCh)
                    End If
                Case gColocEstRefNorm
                    If Not pbEnOtraCmac Then
                        nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefNorEfec, gCredPagRefNorCh)
                    Else
                        nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefNorEOCEfec, gCredPagRefNorEOCCh)
                    End If
                Case gColocEstRefVenc
                    If Not pbEnOtraCmac Then
                        nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefVenEfec, gCredPagRefVenCh)
                    Else
                        nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefVenEOCEfec, gCredPagRefVenEOCCh)
                    End If
                'si es Credito Normal
                Case gColocEstVigMor
                    If Not pbEnOtraCmac Then
                        nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorMorEfec, gCredPagNorMorCh)
                    Else
                        nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorMorEOCEfec, gCredPagNorMorEOCCh)
                    End If
                Case gColocEstVigNorm
                    If Not pbEnOtraCmac Then
                        nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorNorEfec, gCredPagNorNorCh)
                    Else
                        nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorNorEOCEfec, gCredPagNorNorEOCCh)
                    End If
                Case gColocEstVigVenc
                    If Not pbEnOtraCmac Then
                        nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorVenEfec, gCredPagNorVenCh)
                    Else
                        nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorVenEOCEfec, gCredPagNorVenEOCCh)
                    End If
            End Select
        Else
            Select Case nEstadoCred
            'Si es Credito refinanciado
            Case gColocEstRefMor
                If Not pbEnOtraCmac Then
                    nConsCred = gCredPagRefMorCC
                Else
                    nConsCred = 100205
                End If
            Case gColocEstRefNorm
                If Not pbEnOtraCmac Then
                    nConsCred = gCredPagRefNorCC
                 Else
                    nConsCred = 100205
                 End If
                
            Case gColocEstRefVenc
                If Not pbEnOtraCmac Then
                    nConsCred = gCredPagRefVenCC
                Else
                    nConsCred = 100205
                End If
                
            'si es Credito Normal
            Case gColocEstVigMor
                If Not pbEnOtraCmac Then
                    nConsCred = gCredPagNorMorCC
                Else
                    nConsCred = 100205
                End If
                
            Case gColocEstVigNorm
                If Not pbEnOtraCmac Then
                    nConsCred = gCredPagNorNorCC
                Else
                    nConsCred = 100205
                End If
            Case gColocEstVigVenc
                If Not pbEnOtraCmac Then
                    nConsCred = gCredPagNorVenCC
                Else
                    nConsCred = 100205
                End If
            End Select
        End If
    Else
        'Definir Codigo de Operacion de Pago
        Select Case nEstadoCred
            'Si es Credito refinanciado
            Case gColocEstRefMor
                nConsCred = gCredPagRefMorDacion
            Case gColocEstRefNorm
                nConsCred = gCredPagRefNorDacion
            Case gColocEstRefVenc
                nConsCred = gCredPagRefVenDacion
            'si es Credito Normal
            Case gColocEstVigMor
                nConsCred = gCredPagNorMorDacion
            Case gColocEstVigNorm
                nConsCred = gCredPagNorNorDacion
            Case gColocEstVigVenc
                nConsCred = gCredPagNorVenDacion
        End Select
    End If
    Set oCred = New DCredito
    Set R = oCred.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
    Set oCred = Nothing
    pnPlazo = IIf(IsNull(R!nPlazo), 0, R!nPlazo)
    
    R.Close
    Set R = Nothing
    
    Set oCred = New DCredito
    Set R = oCred.RecuperaColocacCred(psCtaCod)
    Set oCred = Nothing
    nNroCalen = R!nNroCalen
    nNroCalPar = R!nNroCalPar
    nMivivienda = IIf(IsNull(R!bMiVivienda), 0, R!bMiVivienda)
    nBuenPagador = IIf(IsNull(R!nCalPago), 0, R!nCalPago)
    nDiasAtrasoMov = IIf(IsNull(R!nDiasAtraso), 0, R!nDiasAtraso)
    R.Close
    Set R = Nothing
    
    Set oBase = New DCredActualizaBD
    
    If psMovnroTemp <> "" Then
        sMovNro = psMovnroTemp
    Else
        
        sMovNro = oBase.GeneraMovNro(pdHoy, psCodAge, psCodUser)
    End If
    
    '*******************************************************************
    'Inserta o Actualiza Gastos Generados por la Operacion
    '*******************************************************************
    Call oBase.dBeginTrans
    If IsArray(pMatGastosGen) Then
        Set oCred = New DCredito
        For i = 0 To pnNumGastosGen - 1
            If (MatrizEstadoCalendario(MatCalendDistrib) = gColocCalendEstadoPagado And pMatGastosGen(i, 4) = "CA") Or (pnPrepago = 1 And pMatGastosGen(i, 4) = "PP") Or pMatGastosGen(i, 4) = "PA" Then
                If oCred.ExisteGasto(psCtaCod, nNroCalen, CInt(Trim(Right(pMatGastosGen(i, 0), 2))), CInt(pMatGastosGen(i, 1)), CLng(Trim(Right(pMatGastosGen(i, 2), 10)))) Then
                    Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, CInt(Trim(Right(pMatGastosGen(i, 0), 2))), CInt(pMatGastosGen(i, 1)), CLng(Trim(Right(pMatGastosGen(i, 2), 10))), CDbl(pMatGastosGen(i, 3)), , "", False, , True)
                Else
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, CInt(Trim(Right(pMatGastosGen(i, 0), 2))), CInt(pMatGastosGen(i, 1)), CLng(Trim(Right(pMatGastosGen(i, 2), 10))), CDbl(pMatGastosGen(i, 3)), 0, "", False)
                End If
            End If
        Next i
        Set oCred = Nothing
    End If
    Call oBase.dCommitTrans
    Set oBase = Nothing
    '**********************************************************************
    '****** En Caso de Que sea cancelacion Con un Desembolso
    '**********************************************************************
    bTran = True
    If pConnBase Is Nothing Then
        Set oBase = New DCredActualizaBD
        Call oBase.dBeginTrans
    Else
        Set oBase = pConnBase
    End If
    dFechaTran = CDate(Format(Format(pdHoy, "dd/mm/yyyy") & " " & Format(oBase.dFechaHora, "hh:mm:ss"), "dd/mm/yyyy hh:mm:ss"))
    
    'Actualiza Producto
    If MatrizEstadoCalendario(MatCalendDistrib) = gColocCalendEstadoPagado Then
        nEstadoCred = gColocEstCancelado
    End If
    
    If nMivivienda = 1 Then
        Call oBase.dUpdateProducto(psCtaCod, , MatrizSaldoCapitalMiVivienda(MatCalend, MatCalendDistrib, MatCalendDistribParalelo), nEstadoCred, pdHoy, nTransacc + 1, False)
    Else
        Call oBase.dUpdateProducto(psCtaCod, , MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCred, pdHoy, nTransacc + 1, False)
    End If
    
    Set oCred = New DCredito
    Set R = oCred.RecuperaColocaciones(psCtaCod)
    Set oCred = Nothing
    sLineaCred = R!cLineaCred
    R.Close
    Set R = Nothing
    '********* ACTIVAR PARA CONTRO DEL SALDOS DE LINEA
    'Set R = oBase.RecuperaLineasCreditoSaldo(sLineaCred, True)
    'nMontoColocado = R!nMontoColocado
    'R.Close
    'Set R = Nothing
    'nMontoColocado = nMontoColocado - pnMonto
    'Call oBase.dUpdateLineaCreditoSaldo(sLineaCred, , , nMontoColocado, , , False)
    
    'Actualiza ColocacCred
    dFecPend = MatrizFechaCuotaPendiente(MatCalend, MatCalendDistrib)
    If MatrizCuotaPendiente(MatCalend, MatCalendDistrib) = 0 Then
        nDiasAtraso = 0
    Else
        nDiasAtraso = pdHoy - dFecPend
    End If
    
    nIntPend = pnIntPend - pnIntPendPagado
    Call oBase.dUpdateColocacCred(psCtaCod, nDiasAtraso, , , , , , , MatrizCuotaPendiente(MatCalend, MatCalendDistrib), nIntPend, , , , , , , , , False)
    
    'Insert Movimientos
    If psMovnroTemp <> "" Then
        nMovNro = oBase.dGetnMovNro(psMovnroTemp)
        If nMovNro = 0 Then
            Call oBase.dInsertMov(psMovnroTemp, nConsCred, "", gMovEstContabMovContable, gMovFlagVigente, False)
            nMovNro = oBase.dGetnMovNro(psMovnroTemp)
        End If
    Else
        Call oBase.dInsertMov(sMovNro, nConsCred, "", gMovEstContabMovContable, gMovFlagVigente, False)
        nMovNro = oBase.dGetnMovNro(sMovNro)
    End If
    '************************************************************************************
    'Si es una Pago en Otra Cmac
    '************************************************************************************
    If pbEnOtraCmac Then
        ' determinar si el pago es un pago de rfa
        Dim oDRFA As DRFA
        Dim bRfa As Boolean
        
        Set oDRFA = New DRFA
        bRfa = oDRFA.VerificaCred(psCtaCod)
        Set oDRFA = Nothing
        If bRfa = True Then
            If bDif = 0 Then
                bDif = 1
                Call oBase.dInsertMovCMAC(nMovNro, psPersCmac, Format$(gTpoIFCmac, "00"), CInt(Mid(psCtaCod, 9, 1)), "", psNroDoc, nConsCred, pnMonto, False)
            Else
                Call oBase.dActualizarMovCmac(nMovNro, pnMonto, False)
            End If
        Else
            Call oBase.dInsertMovCMAC(nMovNro, psPersCmac, Format$(gTpoIFCmac, "00"), CInt(Mid(psCtaCod, 9, 1)), "", psNroDoc, nConsCred, pnMonto, False)
        End If
        
    End If
    '************************************************************************************
    'Si es Cancelado con un Desembolso Insertar su Referencia
    '************************************************************************************
    'If Not pConnBase Is Nothing Then
    '    Call oBase.dInsertMovRef(pnMovnro, nMovNro, False)
    'End If
    
    If pnTipoPago = gColocTipoPagoCheque Then
        Call oBase.dInsertaMovDoc(nMovNro, TpoDocCheque, psNroDoc, dFechaTran, False)
    End If
    
    If pnTipoPago <> gColocTipoPagoDacionPago Then
        If nMivivienda = 1 Then
            nMontoPago = MatrizMontoPagado(MatCalendDistrib)
            Call oBase.dInsertMovCol(nMovNro, nConsCred, psCtaCod, nNroCalen, IIf(pnMontoDesagio > 0, nMontoPago - pnMontoDesagio, nMontoPago), nDiasAtrasoMov, psMetLiquid, pnPlazo, MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCredTemp, False, , pnPrepago)
        Else
            Call oBase.dInsertMovCol(nMovNro, nConsCred, psCtaCod, nNroCalen, IIf(pnMontoDesagio > 0, pnMonto - pnMontoDesagio, pnMonto), nDiasAtrasoMov, psMetLiquid, pnPlazo, MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCredTemp, False, , pnPrepago)
        End If
    Else
        Call oBase.dInsertMovCol(nMovNro, nConsCred, psCtaCod, nNroCalen, pnMonto, nDiasAtrasoMov, psMetLiquid, pnPlazo, MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCredTemp, False, pnNroDacion, pnPrepago)
    End If
    
    '*********  ITF  *****************
    If Not pbEnOtraCmac Then
        Call oBase.dInsertMovCol(nMovNro, gCredITF, psCtaCod, nNroCalen, pnITF, 0, "", 0, 0#, nEstadoCredTemp, False)
        Call oBase.dInsertMovColDet(nMovNro, gCredITF, psCtaCod, CLng(nNroCalen), gConcITFCliente, 0, pnITF, False)
    Else
        Call oBase.dInsertMovCol(nMovNro, gCredITFEOC, psCtaCod, nNroCalen, pnITF, 0, "", 0, 0#, nEstadoCredTemp, False)
        Call oBase.dInsertMovColDet(nMovNro, gCredITFEOC, psCtaCod, CLng(nNroCalen), gConcITFCliente, 0, pnITF, False)
    End If
    'Carga Gastos en Memoria para Evitar Bloqueo
    MatGastosCred = DevuelveMatrizGastosCredito(NumregGastosCred, psCtaCod, nNroCalen)
    
    'Actualiza calendario (ColocCalendario y ColocCalendDet)
    For i = 0 To UBound(MatCalendDistrib) - 1
        If MatrizMontoPagado(MatCalendDistrib, CInt(MatCalendDistrib(i, 1))) > 0 Then
            Call oBase.dUpdateColocCalendario(psCtaCod, nNroCalen, CInt(MatCalendDistrib(i, 1)), gColocCalendAplCuota, , CInt(MatCalendDistrib(i, 2)), "Pago de Cuota", gColocCalendConceptoProcAprobado, False, , dFechaTran)
        Else
            Call oBase.dUpdateColocCalendario(psCtaCod, nNroCalen, CInt(MatCalendDistrib(i, 1)), gColocCalendAplCuota, , CInt(MatCalendDistrib(i, 2)), "Pago de Cuota", gColocCalendConceptoProcAprobado, False)
        End If
        'If nMivivienda = 1 Then  'Si es buen pagador se cancela sus otras cuotas
        '    If pnCalPago = 1 Then
        '        If CInt(MatCalendDistrib(i, 2)) = gColocCalendEstadoPagado Then
        '            Call oBase.dUpdateColocCalendario(psCtaCod, nNroCalPar, CInt(MatCalendDistrib(i, 1)), gColocCalendAplCuota, , CInt(MatCalendDistrib(i, 2)), "Pago de Cuota", gColocCalendConceptoProcOtro, False)
        '        End If
        '    End If
        'End If
        
        ' se verifica de que tipo de credito se esta amortizando (RFC-DIF-RFA)
         'cLRFA = GetCRFA(psCtaCod)
        'Amortizando Capital
        If CDbl(MatCalendDistrib(i, 3)) > 0 Then
            'Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocRFACapital, , CDbl(MatCalendDistrib(i, 3)), , False, True)
            'Inserta Detalle Movimiento Capital
             Call CapitalDistRFA(CDbl(MatCalendDistrib(i, 3)), psCtaCod, nMontoK, nIntDif)
             
             If pcRFA = "RFA" Then
                Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocRFACapital, , CDbl(MatCalendDistrib(i, 3)), , False, True)
                Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocRFACapital, CInt(MatCalendDistrib(i, 1)), nMontoK, False)
                Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocRFAInteresDiferidoRFA, CInt(MatCalendDistrib(i, 1)), nIntDif, False)
             
             Else
                Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocConceptoCodCapital, , CDbl(MatCalendDistrib(i, 3)), , False, True)
                Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodCapital, CInt(MatCalendDistrib(i, 1)), nMontoK, False)
                Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocRFAInteresDiferido, CInt(MatCalendDistrib(i, 1)), nIntDif, False)
             End If
            
            
            'Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodCapital, CInt(MatCalendDistrib(I, 1)), nMontoK, False)
            'Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), "124353", CInt(MatCalendDistrib(I, 1)), nIntDif, False)
            
        End If
        'Amortizando Interes Compensatorio
        If CDbl(MatCalendDistrib(i, 4)) > 0 Then
            
            'Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocRFAInteresCompesatorio, , CDbl(MatCalendDistrib(i, 4)), , False, True)
            'Inserta Detalle Movimiento Interes Compensatorio
            If pcRFA = "RFA" Then
                Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocRFAInteresCompesatorio, , CDbl(MatCalendDistrib(i, 4)), , False, True)
                Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocRFAInteresCompesatorio, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 4)), False)
            Else
                Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocConceptoCodInteresCompensatorio, , CDbl(MatCalendDistrib(i, 4)), , False, True)
                Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresCompensatorio, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 4)), False)
            End If
        End If
        'Amortizando Interes Gracia
        If CDbl(MatCalendDistrib(i, 5)) > 0 Then
            Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocConceptoCodInteresGracia, , CDbl(MatCalendDistrib(i, 5)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresGracia, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 5)), False)
        End If
        'Amortizando Interes Moratorio
        If CDbl(MatCalendDistrib(i, 6)) > 0 Then
            If pcRFA = "RFA" Then
                Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocRFAInteresMoratorio, , CDbl(MatCalendDistrib(i, 6)), , False, True)
                Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocRFAInteresMoratorio, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 6)), False)
            Else
                Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocConceptoCodInteresMoratorio, , CDbl(MatCalendDistrib(i, 6)), , False, True)
                Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresMoratorio, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 6)), False)
            End If
        
            'Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocRFAInteresMoratorio, , CDbl(MatCalendDistrib(i, 6)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
           ' Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocRFAInteresMoratorio, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 6)), False)
        End If
        'Amortizando Interes Reprog
        If CDbl(MatCalendDistrib(i, 7)) > 0 Then
            Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocConceptoCodInteresReprogramado, , CDbl(MatCalendDistrib(i, 7)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresReprogramado, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 7)), False)
        End If
        'Amortizando Interes Suspenso
        If CDbl(MatCalendDistrib(i, 8)) > 0 Then
            Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocConceptoCodInteresSuspenso, , CDbl(MatCalendDistrib(i, 8)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresSuspenso, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 7)), False)
        End If
        
        'Amortizando Interes Compensatorio Vencido
        If CDbl(MatCalendDistrib(i, 11)) > 0 Then
            Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocConceptoCodInteresCompVencido, , CDbl(MatCalendDistrib(i, 11)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresCompVencido, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 11)), False)
        End If
        
        
        'Amortizando Gastos
        If CDbl(MatCalendDistrib(i, 9)) > 0 Then
            nMontoGasto = CDbl(MatCalendDistrib(i, 9))
            'Set oCalend = New Dcalendario
            'Set R = oCalend.RecuperaCalendarioGastos(psCtaCod, nNroCalen, CInt(MatCalendDistrib(i, 1)), gColocCalendAplCuota)
            'Set oCalend = Nothing
            'Dim MatGastosCuota As Variant
            'Dim NumRegGastosCuota As Integer
            
            MatGastosCuota = DevuelveMatrizGastosCreditoCuota(NumRegGastosCuota, CInt(MatCalendDistrib(i, 1)), MatGastosCred, NumregGastosCred)
            
            For k = 0 To NumRegGastosCuota - 1
                If nMontoGasto >= CDbl(MatGastosCuota(k, 2)) Then
                    Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), CLng(MatGastosCuota(k, 1)), , CDbl(MatGastosCuota(k, 2)), , False, True)
                    'Inserta Detalle Movimiento Gastos
                    Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), CLng(MatGastosCuota(k, 1)), CInt(MatCalendDistrib(i, 1)), CDbl(MatGastosCuota(k, 2)), False)
                    nMontoGasto = nMontoGasto - CDbl(MatGastosCuota(k, 2))
                Else
                    Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), CLng(MatGastosCuota(k, 1)), , nMontoGasto, , False, True)
                    'Inserta Detalle Movimiento Gastos
                    Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), CLng(MatGastosCuota(k, 1)), CInt(MatCalendDistrib(i, 1)), CDbl(Format(nMontoGasto, "#0.00")), False)
                    nMontoGasto = 0
                End If
                
                nMontoGasto = CDbl(Format(nMontoGasto, "#0.00"))
                If nMontoGasto = 0 Then
                    Exit For
                End If
            Next k
            ' R.Close
             Set R = Nothing
        End If
    Next i
    
    'Amortizando Desagio si Hubiere
    If pnMontoDesagio > 0 Then
        Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresCompDesagio, CInt(MatCalendDistrib(0, 1)), pnMontoDesagio, False)
    End If
    
    '*****************************************************
    'Para Amortizar Segundo Tramo de MiVivienda
    '*****************************************************
    If nMivivienda = 1 Then
        nMontoPago_2 = MatrizMontoPagado(MatCalendDistribParalelo)
        If nMontoPago_2 > 0 Then
            
            Call oBase.dInsertMovCol(nMovNro, nConsCred, psCtaCod, nNroCalPar, nMontoPago_2, 0, psMetLiquid, pnPlazo, MatrizSaldoCapital(MatCalendParalelo, MatCalendDistribParalelo), nEstadoCredTemp, False, , pnPrepago)
            
            For i = 0 To UBound(MatCalendDistribParalelo) - 1
                If MatrizMontoPagado(MatCalendDistribParalelo, CInt(MatCalendDistribParalelo(i, 1))) > 0 Then
                    Call oBase.dUpdateColocCalendario(psCtaCod, nNroCalPar, CInt(MatCalendDistribParalelo(i, 1)), gColocCalendAplCuota, , CInt(MatCalendDistribParalelo(i, 2)), "Pago de Cuota", gColocCalendConceptoProcOtro, False, , dFechaTran)
                Else
                    Call oBase.dUpdateColocCalendario(psCtaCod, nNroCalPar, CInt(MatCalendDistribParalelo(i, 1)), gColocCalendAplCuota, , CInt(MatCalendDistribParalelo(i, 2)), "Pago de Cuota", gColocCalendConceptoProcOtro, False)
                End If
                'Amortizando Capital
                If CDbl(MatCalendDistribParalelo(i, 3)) > 0 Then
                    Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalPar, gColocCalendAplCuota, CInt(MatCalendDistribParalelo(i, 1)), gColocConceptoCodCapital, , CDbl(MatCalendDistribParalelo(i, 3)), , False, True)
                    'Inserta Detalle Movimiento Capital
                    Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalPar), gColocConceptoCodCapital, CInt(MatCalendDistribParalelo(i, 1)), CDbl(MatCalendDistribParalelo(i, 3)), False)
                End If
                'Amortizando Interes Compensatorio
                If CDbl(MatCalendDistribParalelo(i, 4)) > 0 Then
                    Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalPar, gColocCalendAplCuota, CInt(MatCalendDistribParalelo(i, 1)), gColocConceptoCodInteresCompensatorio, , CDbl(MatCalendDistribParalelo(i, 4)), , False, True)
                    'Inserta Detalle Movimiento Interes Compensatorio
                    Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalPar), gColocConceptoCodInteresCompensatorio, CInt(MatCalendDistribParalelo(i, 1)), CDbl(MatCalendDistribParalelo(i, 4)), False)
                End If
                'Amortizando Interes Gracia
                If CDbl(MatCalendDistribParalelo(i, 5)) > 0 Then
                    Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalPar, gColocCalendAplCuota, CInt(MatCalendDistribParalelo(i, 1)), gColocConceptoCodInteresGracia, , CDbl(MatCalendDistribParalelo(i, 5)), , False, True)
                    'Inserta Detalle Movimiento Interes Gracia
                    Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalPar), gColocConceptoCodInteresGracia, CInt(MatCalendDistribParalelo(i, 1)), CDbl(MatCalendDistribParalelo(i, 5)), False)
                End If
                'Amortizando Interes Moratorio
                If CDbl(MatCalendDistribParalelo(i, 6)) > 0 Then
                    Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalPar, gColocCalendAplCuota, CInt(MatCalendDistribParalelo(i, 1)), gColocConceptoCodInteresMoratorio, , CDbl(MatCalendDistribParalelo(i, 6)), , False, True)
                    'Inserta Detalle Movimiento Interes Gracia
                    Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalPar), gColocConceptoCodInteresMoratorio, CInt(MatCalendDistribParalelo(i, 1)), CDbl(MatCalendDistribParalelo(i, 6)), False)
                End If
                'Amortizando Interes Reprog
                If CDbl(MatCalendDistribParalelo(i, 7)) > 0 Then
                    Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalPar, gColocCalendAplCuota, CInt(MatCalendDistribParalelo(i, 1)), gColocConceptoCodInteresReprogramado, , CDbl(MatCalendDistribParalelo(i, 7)), , False, True)
                    'Inserta Detalle Movimiento Interes Gracia
                    Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalPar), gColocConceptoCodInteresReprogramado, CInt(MatCalendDistribParalelo(i, 1)), CDbl(MatCalendDistribParalelo(i, 7)), False)
                End If
                'Amortizando Interes Suspenso
                If CDbl(MatCalendDistribParalelo(i, 8)) > 0 Then
                    Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalPar, gColocCalendAplCuota, CInt(MatCalendDistribParalelo(i, 1)), gColocConceptoCodInteresSuspenso, , CDbl(MatCalendDistribParalelo(i, 8)), , False, True)
                    'Inserta Detalle Movimiento Interes Gracia
                    Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalPar), gColocConceptoCodInteresSuspenso, CInt(MatCalendDistribParalelo(i, 1)), CDbl(MatCalendDistribParalelo(i, 8)), False)
                End If
                
                'Amortizando Interes Vencido
                If CDbl(MatCalendDistribParalelo(i, 11)) > 0 Then
                    Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalPar, gColocCalendAplCuota, CInt(MatCalendDistribParalelo(i, 1)), gColocConceptoCodInteresCompVencido, , CDbl(MatCalendDistribParalelo(i, 11)), , False, True)
                    'Inserta Detalle Movimiento Interes Gracia
                    Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalPar), gColocConceptoCodInteresCompVencido, CInt(MatCalendDistribParalelo(i, 1)), CDbl(MatCalendDistribParalelo(i, 11)), False)
                End If
                
            Next i
        End If
                
        'Actualiza Tabla de Calificacion Mivivienda
        If pnPrepago <> 1 Then
            For i = 0 To UBound(MatCalendDistrib) - 1
                If MatCalendDistrib(i, 2) = "1" Then
                    Call oBase.dUpdateColocCalifMiViv(psCtaCod, nNroCalen, CInt(MatCalendDistrib(i, 1)), dFechaTran, , CInt(MatCalendDistrib(i, 2)))
                End If
            Next i
        End If
    End If
    
    
    If pnTipoPago = gColocTipoPagoDacionPago Then
        Call oBase.dAnularColocGarantRec(pnNroDacion, gColocGarantRecEstadoCancelado, False)
    End If
    
    ''''''''''''''''''''
    ' Lavado de Dinero
    If psPersLavDinero <> "" Then
        Call oBase.dInsertaMovLavDinero(nMovNro, psPersLavDinero)
    End If
    ''''''''''''''''''''
    
    ''''''''''''''''''''
    ' Reversion de Garantia
    ''''''''''''''''''''
    Call LiberaGarantiaPago(nMovNro, oBase, psCtaCod, psCodAge, psCodUser, pdHoy, nPrestamo, CapitalPagado, IIf(nEstadoCred = gColocEstCancelado, True, False))
    
    ''''''''''''''''''''
    
    If pConnBase Is Nothing Then
        Call oBase.dCommitTrans
    End If
       
    If pConnBase Is Nothing Then
        If nMivivienda = 1 Then
            
            Call EvaluacionBuenPagador(psCtaCod, pdHoy, psCodAge, psCodUser, nMovNro)
            
        End If
    End If
    If pConnBase Is Nothing Then
        Set oBase = Nothing
    End If
    
    CapitalPagado = 0
    CapitalPagado = CapitalPagado + MatrizCapitalPagado(MatCalendDistrib)
    If nMivivienda = 1 Then
        CapitalPagado = CapitalPagado + MatrizCapitalPagado(MatCalendDistribParalelo)
    End If
    
    
    Exit Function

ErrorAmortizarPago:
    If bTran Then
        Call oBase.dRollbackTrans
        Set oBase = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Sub ExtornarTransfARecuperaciones(ByVal psCtaCod As String, ByVal pnSaldoCap As Double, _
    ByVal pnDemanda As Boolean, ByVal pnDiasAtraso As Integer, ByVal pdHoy As Date, _
    ByVal psCodAge As String, ByVal psCodUser As String)
    
    
End Sub

Public Sub TransferirARecuperaciones(ByVal psCtaCod As String, ByVal pnSaldoCap As Double, _
    ByVal pnDemanda As Integer, ByVal pnDiasAtraso As Integer, ByVal pdHoy As Date, _
    ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnNroCalen As Integer, _
    ByVal pnPrdEstado As Integer, Optional ByVal pnPorcGastoTransferencia As Double = 0, _
    Optional ByVal psMetLiquid As String = "GCIM")
Dim MatCalend As Variant
Dim oBase As DCredActualizaBD
Dim lnInteres As Double
Dim lnMora As Double
Dim lnGasto As Double
Dim dFechaEst As Date
Dim lsFechaHoraGrab As String
Dim sMovNro As String
Dim oFunciones As NContFunciones
Dim pbTran As Boolean
Dim nMovNro As Long
Dim lsOpeConceptoCapital As Long

Select Case pnPrdEstado
    Case gColocEstVigNorm
        If pnDemanda = gColRecDemandaSi Then
            lsOpeConceptoCapital = gColocConceptoCodCambCarteraNorm_JuCD   ' "1405"
        ElseIf pnDemanda = gColRecDemandaNo Then
            lsOpeConceptoCapital = gColocConceptoCodCambCarteraNorm_JuSD  '  "1408"
        End If
    Case gColocEstVigMor
        If pnDemanda = gColRecDemandaSi Then
            lsOpeConceptoCapital = gColocConceptoCodCambCarteraMoro_JuCD  '"1403"
        ElseIf pnDemanda = gColRecDemandaNo Then
            lsOpeConceptoCapital = gColocConceptoCodCambCarteraMoro_JuSD  '"1406"
        End If
    Case gColocEstVigVenc
        If pnDemanda = gColRecDemandaSi Then
            lsOpeConceptoCapital = gColocConceptoCodCambCarteraVenc_JuCD   ' "1404"
        ElseIf pnDemanda = gColRecDemandaNo Then
            lsOpeConceptoCapital = gColocConceptoCodCambCarteraVenc_JuSD   ' "1407"
        End If
    Case gColocEstRefNorm
        If pnDemanda = gColRecDemandaSi Then
            lsOpeConceptoCapital = gColocConceptoCodCambCarteraNorR_JuCD   ' "1411"
        ElseIf pnDemanda = gColRecDemandaNo Then
            lsOpeConceptoCapital = gColocConceptoCodCambCarteraNorR_JuSD   '"1414"
        End If
    Case gColocEstRefMor
        If pnDemanda = gColRecDemandaSi Then
            lsOpeConceptoCapital = gColocConceptoCodCambCarteraMorR_JuCD  ' "1409"
        ElseIf pnDemanda = gColRecDemandaNo Then
            lsOpeConceptoCapital = gColocConceptoCodCambCarteraMorR_JuSD  '"1412"
        End If
    Case gColocEstRefVenc
        If pnDemanda = gColRecDemandaSi Then
            lsOpeConceptoCapital = gColocConceptoCodCambCarteraVenR_JuSD  '"1410"
        ElseIf pnDemanda = gColRecDemandaNo Then
            lsOpeConceptoCapital = gColocConceptoCodCambCarteraVenR_JuSD  '"1413"
        End If
End Select

    On Error GoTo ErrorTransferirARecuperaciones
    pbTran = False
    MatCalend = RecuperaMatrizCalendarioPendiente(psCtaCod)
    lnInteres = MatrizInteresGastosAFecha(psCtaCod, MatCalend, pdHoy) - MatrizGastosFecha(psCtaCod, MatCalend) - MatrizInteresMorFecha(psCtaCod, MatCalend)
    lnMora = MatrizInteresMorFecha(psCtaCod, MatCalend)
    lnGasto = MatrizGastosFecha(psCtaCod, MatCalend)
    
    Set oFunciones = New NContFunciones
    sMovNro = oFunciones.GeneraMovNro(pdHoy, psCodAge, psCodUser)
    Set oFunciones = Nothing
    
    lsFechaHoraGrab = fgFechaHoraGrab(sMovNro)
    
    Set oBase = New DCredActualizaBD

    pbTran = True
    oBase.dBeginTrans
    Call oBase.dUpdateProducto(psCtaCod, , , gColocEstRecVigJud, lsFechaHoraGrab, , False)
    Call oBase.dUpdateColocaciones(psCtaCod, , , , , , , , , sMovNro)
    
    Call oBase.dInsertColocRecup(psCtaCod, lsFechaHoraGrab, 0, lnInteres, lnMora, lnGasto, _
                 0, psMetLiquid, 0, pnDemanda, pnNroCalen, False)
    Call oBase.dInsertColocacEstado(psCtaCod, lsFechaHoraGrab, gColocEstRecVigJud, 0, pnSaldoCap, "Transferencia a Recuperaciones", _
            0, 0, 0, 0, 0, 0, 0, 0)
    ' Insert ColocRecupExpediente
    Call oBase.dInsertColocRecupExpediente(psCtaCod, "", 0, Mid(psCtaCod, 9, 1), gColRecViaProcNoEspecificado, 0, False)
    ' Insert Mov
    Call oBase.dInsertMov(sMovNro, gColRecOpePasoARecup, "Transferencia a Recuperaciones", gMovEstContabMovContable, gMovFlagVigente, False)
    nMovNro = oBase.dGetnMovNro(sMovNro)
    ' Insert MovCol
    Call oBase.dInsertMovCol(nMovNro, gCredPasoARecup, psCtaCod, 0, pnSaldoCap, pnDiasAtraso, "", 0, 0, False)
    ' Insert MovColDet
    'Capital
    Call oBase.dInsertMovColDet(nMovNro, gCredPasoARecup, psCtaCod, 0, lsOpeConceptoCapital, 0, pnSaldoCap, False)
    'Int Compensatorio
    If lnInteres > 0 Then
        Call oBase.dInsertMovColDet(nMovNro, gCredPasoARecup, psCtaCod, 0, gColRecConceptoCodInteresCompensatorio, 0, lnInteres, False)
    End If
    'Int Moratorio
    If lnMora > 0 Then
        Call oBase.dInsertMovColDet(nMovNro, gCredPasoARecup, psCtaCod, 0, gColRecConceptoCodInteresMoratorio, 0, lnMora, False)
    End If
    'Gasto
    If lnGasto > 0 Then
        Call oBase.dInsertColocRecupGastos(psCtaCod, 1, gColRecConceptoCodGastoCobranza, dFechaEst, lnGasto, 0, gColRecGastoEstPendiente, "Gasto Transferido Recuperaciones")
        Call oBase.dInsertMovColDet(nMovNro, gCredPasoARecup, psCtaCod, 0, gColRecConceptoCodGastoCobranza, 0, lnGasto, False)
    End If
    
    Call oBase.dCommitTrans
    Set oBase = Nothing
    Exit Sub

ErrorTransferirARecuperaciones:
    If pbTran Then
        Call oBase.dRollbackTrans
    End If
    Err.Raise Err.Number, "Error En Proceso", Err.Description
        
End Sub
Public Sub AsignarGastoLoteCredito(ByVal pnRanIni As Double, ByVal pnRanFin As Double, _
    ByVal pnAplicado As ColocCalendApl, ByVal psCtaCod As String, ByVal psCodGasto As String, _
    ByVal pnValorGasto As Double, ByVal pnTipoValor As Integer)

Dim oDCred As DCredito
Dim oBase As DCredActualizaBD
Dim R As ADODB.Recordset
Dim nMonto As Double
Dim pbTran As Boolean
    'Tipo Valor 1:Valor
    '           2: Porcentaje
    On Error GoTo ErrorAsignarGastoLoteCredito
    pbTran = False
    Set oDCred = New DCredito
    Set R = oDCred.RecuperaCuotasParaAsignarGasto(pnRanIni, pnRanFin, pnAplicado, psCtaCod, psCodGasto)
    Set oDCred = Nothing
    Set oBase = New DCredActualizaBD
    Call oBase.dBeginTrans
    pbTran = True
    Do While Not R.EOF
        If pnTipoValor = 1 Then
            nMonto = pnValorGasto
        Else
            nMonto = CDbl(Format((pnValorGasto / 100) * R!nMontoCuota))
        End If
        If nMonto > 0 Then
            Call oBase.dInsertColocCalendDet(psCtaCod, R!nNroCalen, pnAplicado, R!nCuota, CLng(psCodGasto), nMonto, 0, "", False)
        End If
        R.MoveNext
    Loop
    Call oBase.dCommitTrans
    R.Close
    Set R = Nothing
    Set oBase = Nothing
    Exit Sub

ErrorAsignarGastoLoteCredito:
    If pbTran Then
        Call oBase.dRollbackTrans
    End If
    Err.Raise Err.Number, "Error En Proceso AsignarGastoLoteCredito", Err.Description
    
End Sub

Public Function EvaluaNota(ByVal pnDiasAtrasoAcum As Integer) As Integer
    
    If pnDiasAtrasoAcum = 0 Then
        EvaluaNota = 1
    End If
    If pnDiasAtrasoAcum > 0 And pnDiasAtrasoAcum <= 3 Then
        EvaluaNota = 2
    End If
    If pnDiasAtrasoAcum > 3 And pnDiasAtrasoAcum <= 5 Then
        EvaluaNota = 3
    End If
    If pnDiasAtrasoAcum > 5 And pnDiasAtrasoAcum <= 7 Then
        EvaluaNota = 4
    End If
    If pnDiasAtrasoAcum > 7 Then
        EvaluaNota = 5
    End If
End Function

Public Sub ActualizarNotaCredito(ByVal psCtaCod As String, ByVal pnNota As Integer, ByVal pdHoy As Date, ByVal psComent As String)
Dim oBase As DCredActualizaBD
Dim dFecha As Date
    On Error GoTo ErrorActualizarNotaCredito
    Set oBase = New DCredActualizaBD
    dFecha = CDate(Format(Format(pdHoy, "dd/mm/yyyy") & " " & Format(oBase.dFechaHora, "hh:mm:ss"), "dd/mm/yyyy hh:mm:ss"))
    Call oBase.dInsertColocCalificacionAnalista(psCtaCod, dFecha, pnNota, psComent, "A", False)
    Set oBase = Nothing
    
    Exit Sub

ErrorActualizarNotaCredito:
        Err.Raise Err.Number, "Error En Proceso ActualizarNotaCredito", Err.Description
    
End Sub

Public Sub ActualizaMetasAnalista(ByVal psCodPers As String, ByVal pnTipoMeta As ColocTipoMetas, _
     ByVal pdInicial As Date, ByVal pdFinal As Date, ByVal pnTipoAct As Integer, _
     ByVal pnMoneda As Integer, Optional pMatMontos As Variant)

Dim i As Integer
Dim oBase As DCredActualizaBD
Dim pbTran As Boolean
'Tipo de Actualizacion 1:Nuevo
'                      2: Modificacion
'                      3: Eliminar

    On Error GoTo ErrorActualizaMetasAnalista
    pbTran = False
    Set oBase = New DCredActualizaBD
    Select Case pnTipoAct
        Case 1
            oBase.dBeginTrans
            pbTran = True
            For i = 0 To UBound(pMatMontos) - 1
                Call oBase.dInsertColocMetasAnalista(psCodPers, pnTipoMeta, CInt(pMatMontos(i, 1)), pdInicial, pdFinal, CDbl(pMatMontos(i, 0)), pnMoneda, False)
            Next i
            oBase.dCommitTrans
        Case 2
            oBase.dBeginTrans
            pbTran = True
            For i = 0 To UBound(pMatMontos) - 1
                Call oBase.dUpdateColocMetasAnalista(psCodPers, pnTipoMeta, CInt(pMatMontos(i, 1)), pdInicial, pdFinal, CDbl(pMatMontos(i, 0)), pnMoneda, False)
            Next i
            oBase.dCommitTrans
        Case 3
            Call oBase.dDeleteColocMetasAnalista(psCodPers, pnTipoMeta, pdInicial, pdFinal, pnMoneda, False)
    End Select
    Set oBase = Nothing

    Exit Sub

ErrorActualizaMetasAnalista:
    If pbTran Then
        oBase.dRollbackTrans
        Set oBase = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso ActualizaMetasAnalista", Err.Description
    
End Sub
    
Public Function ValidaRechazoRetiroCredito(ByVal psCtaCod As String, ByVal pbRefinanc As Boolean, ByVal pnTipoAcc As Integer) As String
Dim oDCred As DCredito
Dim R As ADODB.Recordset
    
    ValidaRechazoRetiroCredito = ""
    If pnTipoAcc = 2 Then
        'Que no se haya pagado cuotas de una refinanciacion
        If pbRefinanc Then
            Set oDCred = New DCredito
            Set R = oDCred.RecuperaProducto(psCtaCod)
            Set oDCred = Nothing
            If R!nTransacc >= 1 Then
                ValidaRechazoRetiroCredito = "No se puede retirar este credito porque ya se han realizados pagos"
            End If
        End If
    End If
End Function

Public Function RechazoRetiroCredito(ByVal psCtaCod As String, ByVal psComent As String, _
        ByVal pnMotivo As ColocMotivRechazo, ByVal pdHoy As Date, ByVal psCodAge As String, _
        ByVal psCodUser As String, ByVal pnMonto As Double, Optional ByVal pnTipoAcc As Integer = 1, _
        Optional pbRefinanc As Boolean = False) As String
        
Dim oBase As DCredActualizaBD
Dim R As ADODB.Recordset
Dim RRef As ADODB.Recordset
Dim oDCred As DCredito
Dim nMov As Long
Dim oFunciones As NContFunciones
Dim sMovNro As String
Dim nMovNro As Long
Dim bTran As Boolean
    '1:Rechazo
    '2:Retiro
    On Error GoTo ErrorRechazoRetiroCredito
    bTran = False
    RechazoRetiroCredito = ValidaRechazoRetiroCredito(psCtaCod, pbRefinanc, pnTipoAcc)
    
    If RechazoRetiroCredito <> "" Then
        Exit Function
    End If
    Set oDCred = New DCredito
    
    If pnTipoAcc = 2 Then
        If pbRefinanc Then
            nMov = oDCred.RecuperaMov(psCtaCod, gCredRefinanciacion)
        Else
            nMov = oDCred.RecuperaMov(psCtaCod, gCredAprobacion)
        End If
    End If
    If pnTipoAcc = 2 And pbRefinanc Then
        Set RRef = oDCred.RecuperaColocacRefinanc(psCtaCod)
    End If
    'Set R = oDCred.RecuperaCorresponCred(psCtaCod)
    Set oDCred = Nothing
    'If R.RecordCount > 0 Then
    '    Set oBase = New DCredActualizaBD
    '    Call oBase.dUpdatecorresponsalia(psCtaCod)
    '    R.Close
    '    Set R = Nothing
    '    Set oBase = Nothing
    'Else
    '    R.Close
    '    Set R = Nothing
    'End If
    '-------
    Set oFunciones = New NContFunciones
    sMovNro = oFunciones.GeneraMovNro(pdHoy, psCodAge, psCodUser)
    Set oFunciones = Nothing
    
    Set oDCred = New DCredito
    Set R = oDCred.RecuperaColocGarantia(psCtaCod)
    Set oDCred = Nothing
    
    Set oBase = New DCredActualizaBD
    'Actualiza ColocacEstado
    oBase.dBeginTrans
    bTran = True
    If pnTipoAcc = 1 Then
        Call oBase.dInsertColocacEstado(psCtaCod, oBase.dFechaHora(Format(pdHoy, "dd/mm/yyyy")), _
            gColocEstRech, 0, pnMonto, psComent, 0, 0, 0, 0, 0, 0, 0, 0, False, pnMotivo)
    Else
        Call oBase.dInsertColocacEstado(psCtaCod, oBase.dFechaHora(Format(pdHoy, "dd/mm/yyyy")), _
            gColocEstRetirado, 0, pnMonto, psComent, 0, 0, 0, 0, 0, 0, 0, 0, False, pnMotivo)
    End If
    
    'Actualiza Producto
    If pnTipoAcc = 1 Then
        Call oBase.dUpdateProducto(psCtaCod, , , gColocEstRech, oBase.dFechaHora(Format(pdHoy, "dd/mm/yyyy")), , False)
    Else
        Call oBase.dUpdateProducto(psCtaCod, , , gColocEstRetirado, oBase.dFechaHora(Format(pdHoy, "dd/mm/yyyy")), , False)
    End If
    
    'Actualiza Garantias
    Do While Not R.EOF
        Call oBase.dUpdateGarantias(R!cTpoDoc, R!cNroDoc, R!nGravado, 2, False)
        Call oBase.dDeleteColocGarantia(R!cNumGarant, psCtaCod, False)
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    'ActualizaMov
    If pnTipoAcc = 2 Then
        Call oBase.dUpdateMov(nMov, , , , gMovFlagExtornado, False)
        Call oBase.dInsertMov(sMovNro, gCredExtAprobacion, "Extorno de Aprobacion Con Refinanciacion", gMovEstContabMovContable, gMovFlagVigente, False)
        nMovNro = oBase.GetnMovNro(sMovNro)
        Call oBase.dInsertMovCol(nMovNro, gCredExtAprobacion, psCtaCod, 0, pnMonto, 0, "", 0, 0, False)
        'Call oBase.dInsertMovRef(nMovNro, nMov, False)
    End If
    If pnTipoAcc = 2 And pbRefinanc Then 'Si es Retiro de Refinanciacion
        Do While Not RRef.EOF
            
            Call oBase.dUpdateProducto(RRef!cCtaCodRef, , , gColocEstVigVenc, , , False)
            RRef.MoveNext
        Loop
        RRef.Close
        Set RRef = Nothing
    End If
    
    oBase.dCommitTrans
    Set oBase = Nothing
    
    Exit Function

ErrorRechazoRetiroCredito:
    If bTran Then
        oBase.dRollbackTrans
        Set oBase = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function AmortizarPagoLote(ByVal MatPagos As Variant, ByVal pnTipoPago As ColocTipoPago, ByVal pdHoy As Date, _
    ByVal psNomAge As String, ByVal psCodAge As String, ByVal psCodUsu As String, _
    ByVal sLpt As String, ByVal psNomInstitucion As String, Optional psNroDoc As String = "", Optional ByVal pbImpBoletas As Boolean = False, _
    Optional ByVal pbCobrarMora As Boolean = False) As String
    
Dim MatCalend() As Variant
Dim MatGastosCancelacion() As Variant
Dim MatCalendDistribuido() As Variant
Dim pMatGastosGen As Variant
Dim pMatCalend As Variant
Dim i, k As Integer
Dim oBase As DCredActualizaBD
Dim oConstante As DConstante
Dim oDoc As NCredDoc
Dim sMovNro As String
Dim nMovNro As Long
Dim oFunciones As New NContFunciones
Dim nConsCred As String
Dim sError As String
Dim sCadImp As String
Dim MatTemp As Variant
Dim dFechaTran As Date
Dim oGastos As NGasto
Dim oCred As DCredito
Dim nNumGastosCancel As Integer
Dim nMontoGastoCargo As Double

    On Error GoTo ErrorAmortizarPagoLote
    
    Set oConstante = New DConstante
    Set oDoc = New NCredDoc
    Set oFunciones = New NContFunciones
    sMovNro = oFunciones.GeneraMovNro(Format(pdHoy, "mm/dd/yyyy"), psCodAge, psCodUsu)
    Set oFunciones = Nothing
    
    Set oBase = New DCredActualizaBD
    
    ReDim MatCalend(UBound(MatPagos))
    ReDim MatCalendDistribuido(UBound(MatPagos))
    ReDim MatGastosCancelacion(UBound(MatPagos))
    Set oGastos = New NGasto

    For i = 0 To UBound(MatPagos) - 1
        MatCalend(i) = RecuperaMatrizCalendarioPendiente(MatPagos(i, 0))
        If Not pbCobrarMora Then
            MatTemp = MatCalend(i)
            For k = 0 To UBound(MatTemp) - 1
                MatTemp(k, 6) = 0
            Next k
            MatCalend(i) = MatTemp
        End If
        pMatCalend = MatCalend(i)
        MatGastosCancelacion(i) = oGastos.GeneraCalendarioGastos(Array(0), Array(0), nNumGastosCancel, pdHoy, MatPagos(i, 0), 1, "PA", , , CDbl(MatPagos(i, 6)), MatrizMontoCapitalAPagar(pMatCalend, pdHoy), pMatCalend(0, 1), , , , , CInt(MatPagos(i, 7)))
        nMontoGastoCargo = MontoTotalGastosGenerado(MatGastosCancelacion(i), nNumGastosCancel, Array("CA", "PA", ""))
        'FEPagoLote.TextMatrix(I, 5) = Format(CDbl(FEPagoLote.TextMatrix(I, 5)) + nMontoGastoCargo, "#0.00")
        'FEPagoLote.TextMatrix(I, 10) = Format(CDbl(FEPagoLote.TextMatrix(I, 10)) + nMontoGastoCargo, "#0.00")
       
        pMatCalend(0, 9) = Format(CDbl(pMatCalend(0, 9)) + nMontoGastoCargo, "#0.00")
        MatCalend(i) = pMatCalend
        MatCalendDistribuido(i) = MatrizDistribuirMonto(MatCalend(i), CDbl(Format(MatPagos(i, 1), "#0.00")), MatPagos(i, 4))
        
    Next i
    Set oGastos = Nothing
     
    dFechaTran = CDate(Format(pdHoy, "dd/mm/yyyy hh:mm:ss"))
    
    '*******************************************************************
    'Inserta o Actualiza Gastos Generados por la Operacion
    '*******************************************************************
    For i = 0 To UBound(MatPagos) - 1
        Call oBase.dBeginTrans
        If IsArray(MatGastosCancelacion(i)) Then
            Set oCred = New DCredito
            pMatGastosGen = MatGastosCancelacion(i)
            For k = 0 To nNumGastosCancel - 1
                If (MatrizEstadoCalendario(MatCalendDistribuido(i)) = gColocCalendEstadoPagado And pMatGastosGen(i, 4) = "CA") Or pMatGastosGen(k, 4) = "PA" Then
                    If oCred.ExisteGasto(MatPagos(i, 0), CInt(MatPagos(i, 8)), CInt(Trim(Right(pMatGastosGen(k, 0), 2))), CInt(pMatGastosGen(k, 1)), CLng(Trim(Right(pMatGastosGen(k, 2), 10)))) Then
                        Call oBase.dUpdateColocCalendDet(MatPagos(i, 0), CInt(MatPagos(i, 8)), CInt(Trim(Right(pMatGastosGen(k, 0), 2))), CInt(pMatGastosGen(k, 1)), CLng(Trim(Right(pMatGastosGen(k, 2), 10))), CDbl(pMatGastosGen(k, 3)), , "", False, , True)
                    Else
                        Call oBase.dInsertColocCalendDet(MatPagos(i, 0), CInt(MatPagos(i, 8)), CInt(Trim(Right(pMatGastosGen(k, 0), 2))), CInt(pMatGastosGen(k, 1)), CLng(Trim(Right(pMatGastosGen(k, 2), 10))), CDbl(pMatGastosGen(k, 3)), 0, "", False)
                    End If
                End If
            Next k
            Set oCred = Nothing
        End If
        Call oBase.dCommitTrans
    Next i
    
    '**********************************************************************
    
    
    oBase.dBeginTrans
    'Insert Movimiento
    Call oBase.dInsertMov(sMovNro, gCredPagLote, "PAGO EN LOTE", gMovEstContabMovContable, gMovFlagVigente, False)
    nMovNro = oBase.dGetnMovNro(sMovNro)
    If pnTipoPago = gColocTipoPagoCheque Then
        Call oBase.dInsertaMovDoc(nMovNro, TpoDocCheque, psNroDoc, dFechaTran, False)
    End If
    For i = 0 To UBound(MatPagos) - 1
        Select Case CInt(MatPagos(i, 5))
            'Si es Credito refinanciado
            Case gColocEstRefMor
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefMorEfec, gCredPagRefMorCh)
            Case gColocEstRefNorm
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefNorEfec, gCredPagRefNorCh)
            Case gColocEstRefVenc
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefVenEfec, gCredPagRefVenCh)
            'si es Credito Normal
            Case gColocEstVigMor
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorNorEfec, gCredPagNorNorCh)
            Case gColocEstVigNorm
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorMorEfec, gCredPagNorMorCh)
            Case gColocEstVigVenc
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorVenEfec, gCredPagNorVenCh)
        End Select
        
        sError = oBase.AmortizarCreditoLote(MatPagos(i, 0), MatCalend(i), MatCalendDistribuido(i), CDbl(Format(MatPagos(i, 1), "#0.00")), pdHoy, MatPagos(i, 4), pnTipoPago, psCodAge, psCodUsu, nMovNro, CInt(MatPagos(i, 5)), CLng(MatPagos(i, 2)), IIf(pnTipoPago = gColocTipoPagoCheque, psNroDoc, ""), CDbl(MatPagos(i, 9)))
        If pbImpBoletas Then
            Call oDoc.ImprimeBoleta(MatPagos(i, 0), MatPagos(i, 3), psNomAge, oConstante.DameDescripcionConstante(gMoneda, CInt(Mid(MatPagos(i, 0), 9, 1))), _
                MatrizCuotasPagadas(MatCalendDistribuido(i)), Format(pdHoy, "dd/mm/yyyy"), Format(pdHoy, "hh:mm:ss"), Trim(Str(CInt(MatPagos(i, 2)) + 1)), "Credito Consumo", _
                MatrizCapitalPagado(MatCalendDistribuido(i)), MatrizIntCompPagado(MatCalendDistribuido(i)), _
                MatrizIntCompVencPagado(MatCalendDistribuido(i)), _
                MatrizIntMorPagado(MatCalendDistribuido(i)), MatrizGastoPag(MatCalendDistribuido(i)), _
                MatrizIntGraciaPagado(MatCalendDistribuido(i)), _
                MatrizIntSuspensoPag(MatCalendDistribuido(i)) + MatrizIntReprogPag(MatCalendDistribuido(i)), _
                MatrizSaldoCapital(MatCalend(i), MatCalendDistribuido(i)), MatrizFechaCuotaPendiente(MatCalend(i), MatCalendDistribuido(i)), _
                psCodUsu, sLpt, , , , , CDbl(MatPagos(i, 9)))
        End If
    Next i
    oBase.dCommitTrans
    Set oBase = Nothing
    Set oConstante = Nothing
    sCadImp = oDoc.ImprimePagoLote(MatPagos, MatCalend, MatCalendDistribuido, psNomAge, pdHoy, psCodUsu, psNomInstitucion, IIf(Mid(MatPagos(0, 0), 9, 1) = "1", "SOLES", "DOLARES"), gsNomCmac)
    Set oDoc = Nothing
    AmortizarPagoLote = sCadImp
    Exit Function

ErrorAmortizarPagoLote:
    Err.Raise Err.Number, "Error En Proceso AmortizarPagoLote", Err.Description
    
End Function

Public Sub AnularRegistroDacionEnPago(ByVal pnNroDacion As Long, ByVal psCtaCod As String)
Dim oBase As DCredActualizaBD
Dim nMovNro As Long

    Set oBase = New DCredActualizaBD
    oBase.dBeginTrans
    Call oBase.dAnularColocGarantRec(pnNroDacion, gColocGarantRecEstadoAnulado, False)
    nMovNro = oBase.dRecuperaMovRegistroDacion(pnNroDacion, psCtaCod)
    Call oBase.dUpdateMov(nMovNro, , , , gMovFlagExtornado, False)
    oBase.dCommitTrans
    Set oBase = Nothing
End Sub

Public Sub ExtornarCredito(ByVal psCtaCod As String, ByVal pdFecSis As Date, ByVal psCodUser As String, _
    ByVal psCodAge As String, ByVal pnMov As Long, psCodOpe As String, ByVal pnMonto As Double, Optional ByVal pnNroDacion As Long = -1, _
    Optional ByVal pbExtornoEnLote As Boolean = False, Optional ByVal poBase As DCredActualizaBD = Nothing, Optional ByVal pnMovNro As Long = -1, _
    Optional ByVal pnPrepago As Integer = -1)
    
Dim oDCred As DCredito
Dim oBase As DCredActualizaBD
Dim oFun As NContFunciones
Dim nMovNro As Long
Dim sMovNro, sMovNroCap As String
Dim MatsMovNro() As String
Dim R, RCap, RCapMov, RCancel As ADODB.Recordset
Dim dCap As Date
Dim sSQL As String
Dim nCuotaMin As Integer
Dim nNroCalen As Integer
Dim nNroCalenPar As Integer
Dim i As Integer
Dim nMontoCond As Double
Dim nMovNroCond As Long

    Set oDCred = New DCredito
    Set R = oDCred.RecuperaDatosExtorno(pnMov, psCtaCod)
    If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA _
        Or psCodOpe = gCredPagNorNorCC Or psCodOpe = gCredPagNorMorCC Or psCodOpe = gCredPagNorVenCC Or psCodOpe = gCredPagRefNorCC Or psCodOpe = gCredPagRefMorCC Or psCodOpe = gCredPagRefVenCC Then
        If psCodOpe = gCredPagNorNorCC Or psCodOpe = gCredPagNorMorCC Or psCodOpe = gCredPagNorVenCC Or psCodOpe = gCredPagRefNorCC Or psCodOpe = gCredPagRefMorCC Or psCodOpe = gCredPagRefVenCC Then
            Set RCapMov = oDCred.RecuperaMovimientosAhorros(pnMov, True)
        Else
            Set RCapMov = oDCred.RecuperaMovimientosAhorros(pnMov)
        End If
    End If
    Set oDCred = Nothing
    
    If pbExtornoEnLote Then
        Set oBase = poBase
    Else
        Set oFun = New NContFunciones
        sMovNro = oFun.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
        Set oFun = Nothing
        Set oBase = New DCredActualizaBD
        ReDim MatsMovNro(0)
        If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Then
            
            Set oFun = New NContFunciones
            Do While Not RCapMov.EOF
                'If RCapMov!cOpeCod = "200301" Then '********* Aca hice la modificacion CAFF
                If RCapMov!copecod = gCredDesembCtaRetiroGastos Or RCapMov!copecod = gCredDesembCtaRetiroGastosDOA Or RCapMov!copecod = gCredDesembCtaRetiroCancelacion Or RCapMov!copecod = gCredDesembCtaRetiroCancelacionDOA Then
                    Sleep 1000
                    ReDim Preserve MatsMovNro(RCapMov.Bookmark)
                    MatsMovNro(RCapMov.Bookmark - 1) = oFun.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
                End If
                RCapMov.MoveNext
            Loop
            RCapMov.MoveFirst
            Set oFun = Nothing
        End If
        Call oBase.dBeginTrans
    End If
    
    Sleep 1000
    Set oFun = New NContFunciones
    If pbExtornoEnLote Then
        sMovNroCap = poBase.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
    Else
        sMovNroCap = oFun.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
    End If
    Set oFun = Nothing
    'Extorna Movimientos
    If Not pbExtornoEnLote Then
        Call oBase.dUpdateMov(pnMov, , , , gMovFlagExtornado, False)
        Call oBase.dInsertMov(sMovNro, gCredExtPago, "Extorno de Pago", gMovEstContabNoContable, gMovFlagDeExtorno, False)
        nMovNro = oBase.dGetnMovNro(sMovNro)
    Else
        nMovNro = pnMovNro
    End If
    
    'Extorna Movimiento de Colocaciones Insertando sus Filas Correspondientes
    If Not pbExtornoEnLote Then
        Call oBase.dInsertMovCol(nMovNro, gCredExtPago, psCtaCod, 0, pnMonto, 0, "", 0, 0, 0, False)
    Else
        Call oBase.dInsertMovCol(nMovNro, gCredExtPagoLote, psCtaCod, 0, pnMonto, 0, "", 0, 0, 0, False)
    End If
    
    
    If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Or psCodOpe = gCredDesembEfec Or psCodOpe = gCredPagNorNorCC Or psCodOpe = gCredPagNorMorCC Or psCodOpe = gCredPagNorVenCC Or psCodOpe = gCredPagRefNorCC Or psCodOpe = gCredPagRefMorCC Or psCodOpe = gCredPagRefVenCC Then
        If psCodOpe = gCredPagNorNorCC Or psCodOpe = gCredPagNorMorCC Or psCodOpe = gCredPagNorVenCC Or psCodOpe = gCredPagRefNorCC Or psCodOpe = gCredPagRefMorCC Or psCodOpe = gCredPagRefVenCC Then
            Call oBase.dExtornaSaldosCalendario(pnMov, gColocCalendAplCuota, psCtaCod, False)
        Else
            'Extorna Saldos Calendario
            Call oBase.dExtornaSaldosCalendario(pnMov, gColocCalendAplDesembolso, psCtaCod, False)
        End If
        If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Then
            Set RCap = oBase.RecuperaMovimientoCapataciones(pnMov)
            'Extorna Retiros por Gastos
            Do While Not RCapMov.EOF
                'If RCapMov!cOpeCod = "200301" Then '********* Aca hice la modificacion CAFF
                If RCapMov!copecod = gCredDesembCtaRetiroGastos Or RCapMov!copecod = gCredDesembCtaRetiroGastosDOA Or RCapMov!copecod = gCredDesembCtaRetiroCancelacion Or RCapMov!copecod = gCredDesembCtaRetiroCancelacionDOA Then
                    'Call oBase.dInsertMov(MatsMovNro(RCapMov.boomark - 1), gCredExtPago, "Extorno de Pago", gMovEstContabNoContable, gMovFlagDeExtorno, False)
                    Call oBase.CapExtornoCargoAho(nMovNro, RCapMov!nMovNro, gAhoExtRetEfec, RCap!cCtaCod, MatsMovNro(RCapMov.Bookmark - 1), "Extorno de Desembolso Abono Cuenta", RCapMov!nMonto)
                End If
                RCapMov.MoveNext
            Loop
            'Recupera DatosExtorno
            dCap = CDate(Mid(RCap!cMovNro, 7, 2) & "/" & Mid(RCap!cMovNro, 5, 2) & "/" & Mid(RCap!cMovNro, 1, 4) & " " & Mid(RCap!cMovNro, 9, 2) & ":" & Mid(RCap!cMovNro, 11, 2) & ":" & Mid(RCap!cMovNro, 13, 2))
            'Extorno de Apertura
            'If RCap!cOpeCod = "200101" Then
                        
            If RCap!copecod = "100102" Then
                Call oBase.CapExtornoApertura(sMovNroCap, nMovNro, pnMovNro, gAhoExtApeEfec, RCap!cCtaCod, sMovNro, "Extorno de Desembolso Abono Cuenta", RCap!nMonto)
            Else
                Call oBase.CapExtornoAbonoAho(sMovNroCap, nMovNro, pnMovNro, gAhoExtDepEfec, RCap!cCtaCod, sMovNro, "Extorno de Desembolso Abono Cuenta", RCap!nMonto)
            End If
            
        ElseIf psCodOpe = gCredPagNorNorCC Or psCodOpe = gCredPagNorMorCC Or psCodOpe = gCredPagNorVenCC Or psCodOpe = gCredPagRefNorCC Or psCodOpe = gCredPagRefMorCC Or psCodOpe = gCredPagRefVenCC Then
            Do While Not RCapMov.EOF
                If RCapMov!copecod = gCredDesembCtaRetiroGastos Or RCapMov!copecod = gCredDesembCtaRetiroGastosDOA Or RCapMov!copecod = gCredDesembCtaRetiroCancelacion Or RCapMov!copecod = gCredDesembCtaRetiroCancelacionDOA Then
                    'Call oBase.CapExtornoCargoAho(nMovNro, RCapMov!nMovNro, gCredDesembCtaRetiroCancelacion, RCapMov!cCtaCod, MatsMovNro(RCapMov.Bookmark - 1), "Extorno de Desembolso Abono Cuenta", RCapMov!nMonto)
                    Call oBase.CapExtornoCargoAho(nMovNro, RCapMov!nMovNro, gCredDesembCtaRetiroCancelacion, RCapMov!cCtaCod, sMovNro, "Extorno de Desembolso Abono Cuenta", RCapMov!nMonto)
                End If
                RCapMov.MoveNext
            Loop
            
        End If
    Else
        'Extorna Saldos Calendario
        Call oBase.dExtornaSaldosCalendario(pnMov, gColocCalendAplCuota, psCtaCod, False)
    End If
    
    'Extorna Saldos de Maestros
    'Colocaccred
    Call oBase.dUpdateColocacCred(psCtaCod, R!nDiasMora, , , , , , , R!nMinCuota, , , , , , , R!nMinCuota, , , False)
    
    'Producto
    If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Or psCodOpe = gCredDesembEfec Then
        Call oBase.dUpdateProducto(psCtaCod, , R!nCapital, R!nCredEstado, , , False, 1)
    Else
        Call oBase.dUpdateProducto(psCtaCod, , R!nCapital, R!nCredEstado, , , False, 1, True)
    End If
        
    '**************************************************************
    'En Caso que el pago halla sido con una Dacion en Pago
    '**************************************************************
    If psCodOpe = gCredPagNorNorDacion Or psCodOpe = gCredPagNorMorDacion Or psCodOpe = gCredPagNorVenDacion _
        Or psCodOpe = gCredPagRefNorDacion Or psCodOpe = gCredPagNorNorDacion Or psCodOpe = gCredPagRefMorDacion Or psCodOpe = gCredPagRefVenDacion Then
        Call oBase.dAnularColocGarantRec(R!nFlag, gColocGarantRecEstadoRegistrado, False)
    End If
    
    '**************************************************************
    'Si es Desembolso Con Cancelacion de Creditos
    '**************************************************************
    If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Or psCodOpe = gCredDesembEfec Then
        Set oDCred = New DCredito
        Set RCancel = oDCred.CreditosCanceladoConDesembolso(psCtaCod, pnMov)
        Do While Not RCancel.EOF
            Set R = oDCred.RecuperaDatosExtorno(pnMov, RCancel!cCtaCod)
            Call oBase.dInsertMovCol(nMovNro, gCredExtPago, RCancel!cCtaCod, 0, RCancel!nMonto, 0, "", 0, 0, 0, False)
            Call oBase.dExtornaSaldosCalendario(pnMov, gColocCalendAplCuota, RCancel!cCtaCod, False)
            Call oBase.dUpdateColocacCred(RCancel!cCtaCod, R!nDiasMora, , , , , , , R!nMinCuota, , , , , , , R!nMinCuota, , , False)
            Call oBase.dUpdateProducto(RCancel!cCtaCod, , R!nCapital, R!nCredEstado, , , False, 0)
            RCancel.MoveNext
        Loop
    End If
    '***************************************************************
    'Si es Desembolso es un Hipotecario ---- CAFF
    '***************************************************************
    If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Or psCodOpe = gCredDesembEfec Then
        If Mid(psCtaCod, 6, 3) = "402" Then
            Call oBase.dDeleteColocCalifMiViv(psCtaCod)
        End If
    End If
    
    
            
    If Not pbExtornoEnLote Then
        Call oBase.dCommitTrans
    End If
    
    
    R.Close
    Set R = Nothing
                    
    
    '***************************************************************
    'Si es Prepago
    '***************************************************************
    Dim dCred As Dcalendario
    
    If Not (psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Or psCodOpe = gCredDesembEfec) Then
        If pnPrepago = 1 Then
    
            Call oBase.dBeginTrans
                        
            sSQL = "Select MIN(nNroCuota) as nCuotaMin from MovColDet where nMovNro = " & pnMov
            Set R = oBase.coConex.CargaRecordSet(sSQL)
            If R.RecordCount > 0 Then
                nCuotaMin = R!nCuotaMin
            End If
            R.Close
            
            Set R = oBase.RecuperaColocacCred(psCtaCod)
            nNroCalen = R!nNroCalen
            R.Close
            
            sSQL = "DELETE ColocCalifMiViv Where nCuota >= " & nCuotaMin & " AND nNroCalen = " & nNroCalen & " AND cCtaCod = '" & psCtaCod & "'"
            oBase.coConex.Ejecutar (sSQL)
            
            Set R = oBase.RecuperaColocacCred(psCtaCod)
            nNroCalen = R!nNroCalen - 2
            
            sSQL = "Select ISNULL(MAX(nCuotaOrig),0) as nCuotaMax from ColocCalifMiViv where cCtaCod = '" & psCtaCod & "'"
            Set R = oBase.coConex.CargaRecordSet(sSQL)
            If R.RecordCount > 0 Then
                i = R!nCuotaMax + 1
            End If
            R.Close
            
            Set dCred = New Dcalendario
            Set R = dCred.RecuperaColocCalendario(psCtaCod, nNroCalen, gColocCalendAplCuota)
            Set dCred = Nothing
            
            Do While Not R.EOF
                If R!nCuota >= nCuotaMin Then
                    Call oBase.dInsertColocCalifMiViv(psCtaCod, i, nNroCalen, R!nCuota, R!DVENC, R!nColocCalendEstado)
                End If
                i = i + 1
                R.MoveNext
            Loop
            
            Call oBase.dUpdateColocacCred(psCtaCod, , , , , , , , , , , , , , nNroCalen, , , , False, , , nNroCalen + 1)
            
            sSQL = "DELETE ColocCalendDet Where nNroCalen = " & nNroCalen + 2 & " AND cCtaCod =  '" & psCtaCod & "'"
            Call oBase.coConex.Ejecutar(sSQL)
            sSQL = "DELETE ColocCalendDet Where nNroCalen = " & nNroCalen + 3 & " AND cCtaCod =  '" & psCtaCod & "'"
            Call oBase.coConex.Ejecutar(sSQL)
            
            sSQL = "DELETE ColocCalendario Where nNroCalen = " & nNroCalen + 2 & " AND cCtaCod =  '" & psCtaCod & "'"
            Call oBase.coConex.Ejecutar(sSQL)
            sSQL = "DELETE ColocCalendario Where nNroCalen = " & nNroCalen + 3 & " AND cCtaCod =  '" & psCtaCod & "'"
            Call oBase.coConex.Ejecutar(sSQL)
            
                nMovNroCond = 0
                sSQL = "Select * From MovRef Where nMovNroRef = " & pnMov
                Set R = oBase.coConex.CargaRecordSet(sSQL)
                If R.RecordCount > 0 Then
                    nMovNroCond = R!nMovNro
                End If
                R.Close
                
                If nMovNroCond <> 0 Then
                    sSQL = "Select * From MovCol Where nMovNro = " & nMovNroCond
                    Set R = oBase.coConex.CargaRecordSet(sSQL)
                    If R.RecordCount > 0 Then
                        nMontoCond = R!nMonto
                    End If
                    R.Close
                End If
            
                If nMontoCond <> 0 Then
                    sSQL = "Update Producto Set nSaldo = nSaldo + " & Format(nMontoCond, "#0.00") & " Where cCtaCod = '" & psCtaCod & "'"
                    Call oBase.coConex.Ejecutar(sSQL)
                    
                    sSQL = "Update Mov Set nMovFlag = 1 Where nMovNro = " & nMovNroCond
                    Call oBase.coConex.Ejecutar(sSQL)
                    
                End If
                
            Call oBase.dCommitTrans
        Else 'Si no es un prepago verificar si ha sido una cuota con evaluacion de Mivivienda
            If Mid(psCtaCod, 6, 3) = "402" Then
                sSQL = "Select MIN(nNroCuota) as nCuotaMin from MovColDet where nMovNro = " & pnMov
                Set R = oBase.coConex.CargaRecordSet(sSQL)
                If R.RecordCount > 0 Then
                    nCuotaMin = R!nCuotaMin
                End If
                R.Close
                                                
                nMovNroCond = 0
                sSQL = "Select * From MovRef Where nMovNroRef = " & pnMov
                Set R = oBase.coConex.CargaRecordSet(sSQL)
                If R.RecordCount > 0 Then
                    nMovNroCond = R!nMovNro
                End If
                R.Close
                
                If nMovNroCond <> 0 Then
                    sSQL = "Select * From MovCol Where nMovNro = " & nMovNroCond
                    Set R = oBase.coConex.CargaRecordSet(sSQL)
                    If R.RecordCount > 0 Then
                        nMontoCond = R!nMonto
                    End If
                    R.Close
                End If
                
                Set R = oBase.RecuperaColocacCred(psCtaCod)
                nNroCalen = R!nNroCalen
                nNroCalenPar = R!nNroCalPar
                R.Close
                
                If nCuotaMin = 5 Or (((nCuotaMin - 5) Mod 6) = 0) Then
                    
                    sSQL = " UPDATE ColocCalifMiViv Set cColocMiVivEval = NULL Where nCuota >= " & Trim(Str(nCuotaMin - 5)) & " AND nCuota <= " & nCuotaMin
                    sSQL = sSQL & " AND cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & nNroCalen
                    Call oBase.coConex.Ejecutar(sSQL)
                    
                    sSQL = " UPDATE ColocCalifMiViv Set nColocCalendEstado = 0 Where nCuota = " & nCuotaMin
                    sSQL = sSQL & " AND cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & nNroCalen
                    Call oBase.coConex.Ejecutar(sSQL)
                
                    sSQL = " UPDATE ColocCalendario Set nColocCalendEstado = 0  Where nCuota >= " & nCuotaMin + 2 & " AND nCuota <= " & nCuotaMin + 7
                    sSQL = sSQL & " AND cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & nNroCalenPar & " AND nColocCalendApl = 1 "
                    Call oBase.coConex.Ejecutar(sSQL)
                    
                    If nMontoCond <> 0 Then
                        sSQL = "Update Producto Set nSaldo = nSaldo + " & Format(nMontoCond, "#0.00") & " Where cCtaCod = '" & psCtaCod & "'"
                        Call oBase.coConex.Ejecutar(sSQL)
                        
                        sSQL = "Update Mov Set nMovFlag = 1 Where nMovNro = " & nMovNroCond
                        Call oBase.coConex.Ejecutar(sSQL)
                        
                    End If
                
                End If
            End If
        End If
    End If
    
    Set oBase = Nothing
    
End Sub

Public Sub ExtornarPagoEnLote(ByVal pnMovNro As Long, ByVal pdFecSis As Date, ByVal psCodUser As String, ByVal psCodAge As String)
Dim R As ADODB.Recordset
Dim oDCreditos As DCreditos
Dim oBase As DCredActualizaBD
Dim sMovNro As String
Dim nMovNro As Long
Dim oFun As NContFunciones
    Set oFun = New NContFunciones
    sMovNro = oFun.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
    Set oFun = Nothing
    
    Set oDCreditos = New DCreditos
    Set R = oDCreditos.RecuperaCreditosdePagoEnLote(pnMovNro)
    Set oBase = New DCredActualizaBD
    oBase.dBeginTrans
    Call oBase.dUpdateMov(pnMovNro, , , , gMovFlagExtornado, False)
    Call oBase.dInsertMov(sMovNro, gCredExtPagoLote, "Extorno de Pago", gMovEstContabNoContable, gMovFlagDeExtorno, False)
    nMovNro = oBase.dGetnMovNro(sMovNro)
    Do While Not R.EOF
        Call ExtornarCredito(R!cCtaCod, pdFecSis, psCodUser, psCodAge, pnMovNro, R!copecod, R!nMonto, , True, oBase, nMovNro)
        R.MoveNext
    Loop
    oBase.dCommitTrans
    Set oBase = Nothing
    R.Close
    Set R = Nothing
End Sub

Public Function ValidaNivelAprUsuario(ByVal psCodUser As String, ByVal pnMonto As Double, psMoneda As String, ByVal psProducto As String, ByVal pbRefinanc As Boolean) As Boolean
Dim oCred As DCredito
Dim R As ADODB.Recordset
    
    Set oCred = New DCredito
    Set R = oCred.RecuperaNivelesAprUsuario(psCodUser, pnMonto, psMoneda, psProducto, pbRefinanc)
    If R.RecordCount > 0 Then
        ValidaNivelAprUsuario = True
    Else
        ValidaNivelAprUsuario = False
    End If
    R.Close
    Set R = Nothing
End Function

Public Sub EvaluacionBuenPagador(ByVal psCtaCod As String, ByVal pdHoy As Date, _
    ByVal psCodAge As String, ByVal psCodUser As String, ByVal pnMov As Long)
    
Dim sSQL As String
Dim oConec As DConecta
Dim R As ADODB.Recordset
Dim nUltimaEval As Integer
Dim nAplicaFinal As Integer
Dim nAplicaInicial As Integer
Dim nEstado As Integer
Dim nEvaluaIni As Integer
Dim nEvaluaFin As Integer
Dim nTotalPagosMalos As Integer
Dim nEvaluaBuenPagador  As Integer
Dim oCredDB As DCredActualizaBD
Dim nNroCalen As Integer
Dim nNrocalenParalelo As Integer
Dim i As Integer
Dim nMontoCondonado As Double
Dim oFunciones As NContFunciones
Dim sMovNro As String
Dim nMovNro As Long

    Set oConec = New DConecta
    oConec.AbreConexion
    
    'sSql = "Select MAX(CC.nCuota) as nCuota from ColocCalendario CC Where CC.cColocMiVivEval = 'E' "
    'sSql = sSql & " AND CC.cCtaCod = '" & psCtaCod & "' AND CC.nNroCalen = (Select nNroCalen From ColocacCred Where cCtaCod = CC.cCtaCod) "
    'sSql = sSql & " AND nColocCalendApl = 1 "
    
    sSQL = "Select MAX(CC.nCuotaOrig) as nCuota from ColocCalifMiViv CC Where CC.cColocMiVivEval = 'E' "
    sSQL = sSQL & " AND CC.cCtaCod = '" & psCtaCod & "' "
    
    Set R = oConec.CargaRecordSet(sSQL)
    nUltimaEval = IIf(IsNull(R!nCuota), 0, R!nCuota)
    R.Close
    If nUltimaEval = 0 Then
        nAplicaFinal = 5
    Else
        nAplicaFinal = nUltimaEval + 6
    End If
    
    'Verfico si ya esta Cancelada
    
    'sSql = "Select nColocCalendEstado  from ColocCalendario CC "
    'sSql = sSql & " Where CC.cCtaCod = '" & psCtaCod & "' AND CC.nNroCalen = (Select nNroCalen From ColocacCred wHERE cCtaCod = CC.cCtaCod) "
    'sSql = sSql & " AND nColocCalendApl = 1 AND nCuota = " & nAplicaFinal
    
    sSQL = "Select nColocCalendEstado  from ColocCalifMiViv CC "
    sSQL = sSQL & " Where CC.cCtaCod = '" & psCtaCod & "' "
    sSQL = sSQL & "     AND nCuotaOrig = " & nAplicaFinal
    
    Set R = oConec.CargaRecordSet(sSQL)
    nEstado = R!nColocCalendEstado
    R.Close
    If nEstado = gColocCalendEstadoPagado Then
        If nAplicaFinal = 5 Then
            nAplicaInicial = nAplicaFinal + 2
            nAplicaFinal = nAplicaFinal + 7
            nEvaluaIni = 1
            nEvaluaFin = 5
        Else
            nAplicaInicial = nAplicaFinal + 2
            nAplicaFinal = nAplicaFinal + 7
            nEvaluaIni = nUltimaEval + 1
            nEvaluaFin = nUltimaEval + 6
        End If
        
        '*****************************************************************
        'Evaluamos si fue mal pagador es decir si se atraso en algun pago
        '*****************************************************************
        
        'sSql = "Select COUNT(CC.cCtaCod) as nTotal From ColocCalendario CC "
        'sSql = sSql & " Where CC.cCtaCod = '" & psCtaCod & "' AND CC.nNroCalen  = (Select nNroCalen From ColocacCred Where cCtaCod = CC.cCtaCod)"
        'sSql = sSql & " AND CC.nCuota >= " & nEvaluaIni & " AND CC.nCuota <= " & nEvaluaFin
        'sSql = sSql & " AND nColocCalendApl = 1 "
        'sSql = sSql & " AND (CC.dPago >= DATEADD(day,30,CC.dVenc) OR CC.dPago is null) "
        
        sSQL = "Select COUNT(CC.cCtaCod) as nTotal From ColocCalifMiViv CC "
        sSQL = sSQL & " Where CC.cCtaCod = '" & psCtaCod & "' "
        sSQL = sSQL & " AND CC.nCuotaOrig >= " & nEvaluaIni & " AND CC.nCuotaOrig <= " & nEvaluaFin
        sSQL = sSQL & " AND (CC.dPago >= DATEADD(day,30,CC.dVenc) OR CC.dPago is null) "
        
        Set R = oConec.CargaRecordSet(sSQL)
        nTotalPagosMalos = IIf(IsNull(R!nTotal), 0, R!nTotal)
        R.Close
        
        sSQL = "Select nNrocalen,nNroCalPar from ColocacCred Where cCtaCod = '" & psCtaCod & "'"
        Set R = oConec.CargaRecordSet(sSQL)
        nNroCalen = R!nNroCalen
        nNrocalenParalelo = R!nNroCalPar
        R.Close
        
        If nTotalPagosMalos = 0 Then
            nEvaluaBuenPagador = 1
        Else
            nEvaluaBuenPagador = 0
        End If
        
        If nEvaluaBuenPagador = 1 Then 'Reduzco el saldo de la deuda con el monto condonado
            'nAplicaInicial    nAplicaFinal
            sSQL = "Select ISNULL(SUM(nMonto),0) as nMontoCondonado From ColocCalendDet CD  "
            sSQL = sSQL & " Where cCtaCod = '" & psCtaCod & "' "
            sSQL = sSQL & "  AND   CD.nColocCalendApl = " & gColocCalendAplCuota
            sSQL = sSQL & "  AND   CD.nCuota >= " & nAplicaInicial & " AND CD.nCuota <= " & nAplicaFinal
            sSQL = sSQL & "  AND CD.nPrdConceptoCod = 1000 "
            sSQL = sSQL & "  AND nNroCalen = " & nNrocalenParalelo
            
            Set R = oConec.CargaRecordSet(sSQL)
            nMontoCondonado = R!nMontoCondonado
            R.Close
            
        End If
        
        'Actualizo Tablas
        Set oCredDB = New DCredActualizaBD
        oCredDB.dBeginTrans
        
        Call oCredDB.dUpdateColocacCred(psCtaCod, , , , , , , , , , , , , , , , , , , , , , nEvaluaBuenPagador)
        For i = nEvaluaIni To nEvaluaFin
            Call oCredDB.dUpdateColocCalendario(psCtaCod, nNroCalen, i, gColocCalendAplCuota, , , , gColocCalendConceptoProcOtro, False, "E")
            Call oCredDB.dUpdateColocCalifMiViv(psCtaCod, nNroCalen, i, , "E")
        Next i
        
        If nEvaluaBuenPagador = 1 Then
        
            Set oFunciones = New NContFunciones
            sMovNro = oFunciones.GeneraMovNro(pdHoy, psCodAge, psCodUser)
            Set oFunciones = Nothing
            Call oCredDB.dInsertMov(sMovNro, gCredMiVivCondonacion, "CONDONACION DE DEUDA", gMovEstContabMovContable, gMovFlagVigente, False)
            nMovNro = oCredDB.dGetnMovNro(sMovNro)
            Call oCredDB.dInsertMovCol(nMovNro, gCredMiVivCondonacion, psCtaCod, _
                nNrocalenParalelo, nMontoCondonado, 0, "", 0, 0, 0, False)
            Call oCredDB.dInsertMovColDet(nMovNro, gCredMiVivCondonacion, psCtaCod, CLng(nNrocalenParalelo), _
                gColocConceptoCodCapital, 0, nMontoCondonado, False)
            Call oCredDB.dUpdateProducto(psCtaCod, , nMontoCondonado, , , , False, 1)
            Call oCredDB.dInsertMovRef(nMovNro, pnMov, False)
            
            'Cancelo las Cuotas de Segundo Tramo
            For i = nAplicaInicial To nAplicaFinal
                Call oCredDB.dUpdateColocCalendario(psCtaCod, nNrocalenParalelo, i, gColocCalendAplCuota, , gColocCalendEstadoPagado, , gColocCalendConceptoProcOtro, False, "E")
                Call oCredDB.dUpdateColocCalifMiViv(psCtaCod, nNroCalen, i, , , , CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Mid(sMovNro, 1, 4) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)))
            Next i
            
        End If
        
        oCredDB.dCommitTrans
        Set oCredDB = Nothing
    End If
    
    oConec.CierraConexion
    Set oConec = Nothing

End Sub


Sub CapitalDistRFA(ByVal pnMonto, ByVal psCtaCod As String, ByRef pnMontoK As Double, pnIntDif As Double)
    Dim objDRFA As DRFA
    Dim nFactor As Double
    
    Set objDRFA = New DRFA
    nFactor = objDRFA.ObtenerFactor(psCtaCod)
    Set objDRFA = Nothing
    pnMontoK = Format(pnMonto * nFactor, "#0.00")
    'pnIntDif = pnMonto * (1 - nFactor)
    pnIntDif = pnMonto - pnMontoK
End Sub

Public Function GetCreditosRFA(ByVal psPersCod As String) As ADODB.Recordset
Dim oRFA As DRFA
Set oRFA = New DRFA
Set GetCreditosRFA = oRFA.GetCredRFA(psPersCod)
Set oRFA = Nothing
End Function
Public Function GetCalendariosRFA(ByVal psPersCod As String, ByVal psRFA As String, _
Optional ByVal pbReprogramacion As Boolean = False, Optional ByVal pbRefinan As Boolean) As ADODB.Recordset
Dim oRFA As DRFA
Set oRFA = New DRFA
If pbReprogramacion = True Then
    Set GetCalendariosRFA = oRFA.GetCalendariosRFA(psPersCod, psRFA, True, pbRefinan)
Else
    Set GetCalendariosRFA = oRFA.GetCalendariosRFA(psPersCod, psRFA, False)
End If
Set oRFA = Nothing
End Function

Public Function ReprogramarCalendario(ByVal MatCalendDIF As Variant, ByVal MatCalendRFC As Variant, _
ByVal MatCalendRFA As Variant, ByVal pnNroCalen As Integer, ByVal psCtaCodDIF As String, _
ByVal psCtaCodRFC As String, ByVal psCtaCodRFA As String) As Boolean
    'Definiendo los conceptos
    Dim oDRFA As DRFA
    Dim i As Integer
    Dim oBase As DCredActualizaBD
    Set oBase = New DCredActualizaBD
    
    oBase.coConex.BeginTrans
          'insertando los desembolsoso
          Call oBase.dInsertaDesembolsosReprogramado(psCtaCodRFC, pnNroCalen)
          Call oBase.dInsertaDesembolsosReprogramado(psCtaCodDIF, pnNroCalen)
          Call oBase.dInsertaDesembolsosReprogramado(psCtaCodRFA, pnNroCalen)
          
          'insertandolo demas de las cuotas
          '**********************************************************************
          'RFC
          For i = 1 To UBound(MatCalendRFC)
            'insertando los calendario
            
             Call oBase.dInsertColocCalendario(psCtaCodRFC, pnNroCalen, gColocCalendAplCuota, _
                             MatCalendRFC(i, 1), MatCalendRFC(i, 2), MatCalendRFC(i, 3), "Reprogramacion RFA", _
                             gColocCalendConceptoProcReprogramado, False)
          Next i
          
          'insertanndo el detalle
          For i = 1 To UBound(MatCalendRFC)
          
          '   If MatCalendRFC(i, 4) > 0 Then 'Capital
                Call oBase.dInsertColocCalendDet(psCtaCodRFC, pnNroCalen, gColocCalendAplCuota, MatCalendRFC(i, 1), _
                         "1000", MatCalendRFC(i, 4), MatCalendRFC(i, 8), "", False)
           '  End If
             
             
           ' If MatCalendRFC(i, 5) > 0 Then 'Interes Compesatorio
                Call oBase.dInsertColocCalendDet(psCtaCodRFC, pnNroCalen, gColocCalendAplCuota, MatCalendRFC(i, 1), _
                         "1100", MatCalendRFC(i, 5), MatCalendRFC(i, 9), "", False)
            ' End If
             
             If MatCalendRFC(i, 6) > 0 Then 'Cofide
                If MatCalendRFC(i, 2) = 1 Then 'cancelado
                    Call oBase.dInsertColocCalendDet(psCtaCodRFC, pnNroCalen, gColocCalendAplCuota, MatCalendRFC(i, 1), _
                             "124352", MatCalendRFC(i, 6), MatCalendRFC(i, 6), "", False)
                Else
                    Call oBase.dInsertColocCalendDet(psCtaCodRFC, pnNroCalen, gColocCalendAplCuota, MatCalendRFC(i, 1), _
                             "124352", MatCalendRFC(i, 6), 0, "", False)
                End If
             End If
             
             'If MatCalendRFC(i, 7) > 0 Then 'Mora
                Call oBase.dInsertColocCalendDet(psCtaCodRFC, pnNroCalen, gColocCalendAplCuota, MatCalendRFC(i, 1), _
                         "1101", MatCalendRFC(i, 7), MatCalendRFC(i, 10), "", False)
            ' End If
          Next i
          
          '******DIF*******************************************
          'DIF
          For i = 1 To UBound(MatCalendDIF)
            'insertando los calendario
            
             Call oBase.dInsertColocCalendario(psCtaCodDIF, pnNroCalen, gColocCalendAplCuota, _
                             MatCalendDIF(i, 1), MatCalendDIF(i, 2), MatCalendDIF(i, 3), "Reprogramacion RFA", _
                             gColocCalendConceptoProcReprogramado, False)
          Next i
          
          'insertanndo el detalle
          For i = 1 To UBound(MatCalendDIF)
          
             If MatCalendDIF(i, 4) > 0 Then 'Capital
                Call oBase.dInsertColocCalendDet(psCtaCodDIF, pnNroCalen, gColocCalendAplCuota, MatCalendDIF(i, 1), _
                         "1000", MatCalendDIF(i, 4), MatCalendDIF(i, 8), "", False)
             End If
             
             
            If MatCalendDIF(i, 5) > 0 Then 'Interes Compesatorio
                Call oBase.dInsertColocCalendDet(psCtaCodDIF, pnNroCalen, gColocCalendAplCuota, MatCalendDIF(i, 1), _
                         "1100", MatCalendDIF(i, 5), MatCalendDIF(i, 9), "", False)
             End If
             
             If MatCalendDIF(i, 6) > 0 Then 'Cofide
                If MatCalendDIF(i, 2) = 1 Then
                    Call oBase.dInsertColocCalendDet(psCtaCodDIF, pnNroCalen, gColocCalendAplCuota, MatCalendDIF(i, 1), _
                             "124350", MatCalendDIF(i, 6), MatCalendDIF(i, 6), "", False)
                Else
                    Call oBase.dInsertColocCalendDet(psCtaCodDIF, pnNroCalen, gColocCalendAplCuota, MatCalendDIF(i, 1), _
                             "124350", MatCalendDIF(i, 6), 0, "", False)
                End If
             End If
             
             'If MatCalendDIF(i, 7) > 0 Then 'Gastos
                Call oBase.dInsertColocCalendDet(psCtaCodDIF, pnNroCalen, gColocCalendAplCuota, MatCalendDIF(i, 1), _
                         "1101", MatCalendDIF(i, 7), MatCalendDIF(i, 10), "", False)
             'End If
          Next i
          
          '*********************************************
          '********************RFA********************************************
          For i = 1 To UBound(MatCalendRFA)
            'insertando los calendario
            
             Call oBase.dInsertColocCalendario(psCtaCodRFA, pnNroCalen, gColocCalendAplCuota, _
                             MatCalendRFA(i, 1), MatCalendRFA(i, 2), MatCalendRFA(i, 3), "Reprogramacion RFA", _
                             gColocCalendConceptoProcReprogramado, False)
          Next i
          
          'insertanndo el detalle
          For i = 1 To UBound(MatCalendRFA)
          
             If MatCalendRFA(i, 4) > 0 Then 'Capital
                Call oBase.dInsertColocCalendDet(psCtaCodRFA, pnNroCalen, gColocCalendAplCuota, MatCalendDIF(i, 1), _
                         "1010", MatCalendRFA(i, 4), MatCalendRFA(i, 8), "", False)
             End If
             
             
            If MatCalendRFA(i, 5) > 0 Then 'Interes Compesatorio
                Call oBase.dInsertColocCalendDet(psCtaCodRFA, pnNroCalen, gColocCalendAplCuota, MatCalendRFA(i, 1), _
                         "1107", MatCalendRFA(i, 5), MatCalendRFA(i, 9), "", False)
             End If
             
             If MatCalendRFA(i, 6) > 0 Then 'Cofide
                If MatCalendRFA(i, 2) = 1 Then
                    Call oBase.dInsertColocCalendDet(psCtaCodRFA, pnNroCalen, gColocCalendAplCuota, MatCalendRFA(i, 1), _
                             "124350", MatCalendRFA(i, 6), MatCalendRFA(i, 6), "", False)
                Else
                    Call oBase.dInsertColocCalendDet(psCtaCodRFA, pnNroCalen, gColocCalendAplCuota, MatCalendRFA(i, 1), _
                             "124350", MatCalendRFA(i, 6), 0, "", False)
                End If
             End If
             
             'If MatCalendRFA(i, 7) > 0 Then 'Gastos
                Call oBase.dInsertColocCalendDet(psCtaCodRFA, pnNroCalen, gColocCalendAplCuota, MatCalendRFA(i, 1), _
                         "1108", MatCalendRFA(i, 7), MatCalendRFA(i, 11), "", False)
             'End If
          Next i
          
          
          Call oBase.dUpdateColocacCred(psCtaCodRFA, , , , , , , , , , , , , , pnNroCalen, , , , False)
          Call oBase.dUpdateColocacCred(psCtaCodDIF, , , , , , , , , , , , , , pnNroCalen, , , , False)
          Call oBase.dUpdateColocacCred(psCtaCodRFC, , , , , , , , , , , , , , pnNroCalen, , , , False)
          
          
    oBase.coConex.CommitTrans
    Set oBase = Nothing
    
    Set oDRFA = New DRFA
    Call oDRFA.ActualizaCalendarioReprog(psCtaCodDIF)
    Call oDRFA.ActualizaCalendarioReprog(psCtaCodRFC)
    Call oDRFA.ActualizaCalendarioReprog(psCtaCodRFA)
    Set oDRFA = Nothing
End Function

    
