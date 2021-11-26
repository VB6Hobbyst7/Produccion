Attribute VB_Name = "gCalendario"
Option Explicit

'Tipos de Gracia
Enum TCalendTipoGracia
    PrimeraCuota = 1
    UltimaCuota = 2
    Prorateada = 3
    Configurable = 4
    Exonerada = 5
    EnCuotas = 6
    Capitalizada = 7
End Enum
'Tipos de Periodo
Enum TCalendTipoPeriodo
    PeriodoFijo = 1
    FechaFija = 2
End Enum

'Tipos de Cuota
Enum TCalendTipoCuota
    Fija = 1
    Creciente = 2
    Decreciente = 3
End Enum

Private Type TCalendario
    dFecha As Date
    NroCuota As Integer
    Cuota As Double
    Captital As Double
    IntComp As Double
    IntGra As Double
    Gasto As Double
    SegDes As Double
    CuotaPrimaPoliza As Double 'LUCV20180601, Según ERS022-2018
    CuotaPrimaPolizaGracia As Double 'LUCV20180601, Según ERS022-2018
    SaldoCap As Double
    SegBien As Double 'MAVM 20121113
    'MAVM 20130209 ***
    CuotaGra As Double
    IntCompGra As Double
    SaldoCapGra As Double
    '***
    'RECO20150512***********
    SegMultMype As Double
    'RECO FIN **************
End Type

'Para Reporte no Borrar
Dim csNomCMAC As String
Dim csNomAgencia As String
Dim csCodUser As String
Dim csFechaSis As String

Dim Calendario() As TCalendario
Dim NroCuotas As Integer

Dim oImpre As COMFunciones.FCOMImpresion
Dim oITF As COMDConstSistema.FCOMITF

Public Function GeneraGracia(ByVal pOptGracia As TCalendTipoGracia, ByVal pnInteres As Double, ByVal pnNroCuotas As Integer) As Variant
Dim nValorProrateado As Double
Dim i As Integer
Dim MatGracia() As Double
Dim sMatGracia() As String

    nValorProrateado = pnInteres / pnNroCuotas
    nValorProrateado = CDbl(Format(nValorProrateado, "#0.00"))
    ReDim MatGracia(pnNroCuotas)
    If pOptGracia = PrimeraCuota Then
        ReDim Preserve MatGracia(pnNroCuotas + 1)
        MatGracia(0) = pnInteres
    End If
    If pOptGracia = UltimaCuota Then
        ReDim Preserve MatGracia(pnNroCuotas + 1)
        MatGracia(pnNroCuotas) = pnInteres
    End If
    If pOptGracia = Prorateada Then
        For i = 0 To pnNroCuotas - 1
            If i = pnNroCuotas - 1 Then
                MatGracia(i) = pnInteres - CDbl(Format((nValorProrateado * (pnNroCuotas - 1)), "#0.00"))
            Else
                MatGracia(i) = nValorProrateado
            End If
        Next i
    End If
    
    If pOptGracia = Exonerada Then
        ReDim MatGracia(pnNroCuotas)
    End If
    
    ReDim sMatGracia(UBound(MatGracia))
    For i = 0 To UBound(MatGracia) - 1
        sMatGracia(i) = Format(MatGracia(i), "#0.00")
    Next i
    GeneraGracia = sMatGracia
End Function

Private Sub CalendarioCreciente(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, Optional ByVal MatGracia As Variant, _
                Optional ByVal pnCuotaIni As Integer = 0, Optional ByVal pnCuotaFin As Integer = 0, Optional ByVal pnDiaFijo2 As Integer = 0, _
                Optional pnMontoCapInicial As Double = 0)

Dim nSaldoCapital As Double
Dim dDesembolso As Date
Dim i As Integer
Dim oCredito As NCOMCredito
Dim dFecTemp As Date
Dim nMes As Integer
Dim nAnio As Integer
Dim nDia As Integer
Dim nSumCuotas As Double

'Para llevar el control del Capital en el caso de capitalizar la gracia
'pero sin incrementar el capital
Dim nSumCapital As Double

    nSaldoCapital = pnMonto
    dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
    nSumCuotas = 0
    nSumCapital = 0
    
    For i = 1 To pnNroCuotas
        nSumCuotas = nSumCuotas + i
    Next i
        
    If pnTipoPeriodo = PeriodoFijo Then
        Set oCredito = New NCOMCredito
        'For i = 0 To pnNroCuotas - 1
        For i = pnCuotaIni To pnCuotaFin
            Calendario(i).dFecha = dDesembolso + pnPeriodo
            Calendario(i).NroCuota = i + 1
            Calendario(i).IntGra = 0#
            Calendario(i).Gasto = 0#
            Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital)
            'If i = pnNroCuotas - 1 Then
            If i = pnCuotaFin Then
                Calendario(i).Captital = nSaldoCapital
            Else
                Calendario(i).Captital = CDbl(Format(((Calendario(i).NroCuota - pnCuotaIni) / nSumCuotas) * pnMonto, "#0.00"))
            End If
            nSumCapital = nSumCapital + Calendario(i).Captital
            Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
            Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
            nSaldoCapital = nSaldoCapital - Calendario(i).Captital
            dDesembolso = dDesembolso + pnPeriodo
            'Verificar en el caso de capitalizar la gracia
            If pnMontoCapInicial > 0 Then
                'Recalculamos los resultados
                If nSumCapital > pnMontoCapInicial Then
                    Calendario(i).Captital = Calendario(i).Captital - (nSumCapital - pnMontoCapInicial)
                    Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                    'Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                End If
                If Calendario(i).Captital < 0 Then
                    Calendario(i).IntGra = 0#
                    Calendario(i).Gasto = 0#
                    Calendario(i).IntComp = 0#
                    Calendario(i).Captital = 0#
                    Calendario(i).Cuota = 0#
                    Calendario(i).SaldoCap = 0#
                End If
            End If
            '*********************************************
        Next i
               
    Else
        'Si es Fecha Fija
        If pnTipoPeriodo = FechaFija Then
            Set oCredito = New COMNCredito.NCOMCredito
            nMes = Month(dDesembolso)
            nAnio = Year(dDesembolso)
            'nDia = pnDiaFijo
            '**************************************
            'Se agrego para manejar 2 Fechas Fijas
            If pnDiaFijo2 = 0 Then
                nDia = pnDiaFijo
            Else
                If Day(dDesembolso) <= pnDiaFijo2 - 8 Then
                    nDia = pnDiaFijo2
                Else
                    nDia = pnDiaFijo
                End If
            End If
            '**************************************

            'For i = 0 To pnNroCuotas - 1
            For i = pnCuotaIni To pnCuotaFin
                'nDia = pnDiaFijo
                '**************************************
                'Se agrego para manejar 2 Fechas Fijas
                If pnDiaFijo2 = 0 Then
                    nDia = pnDiaFijo
                Else
                    If i > pnCuotaIni Then
                        If nDia = pnDiaFijo Then
                            nDia = pnDiaFijo2
                        Else
                            nDia = pnDiaFijo
                        End If
                    End If
                End If
                '**************************************
                'If Not (i = 0 And pnDiaFijo > Day(dDesembolso) And Not bProxMes) Then
                If Not (i = 0 And nDia > Day(dDesembolso) And Not bProxMes) Then
                    '**************************************
                    'Se modifico para manejar 2 Fechas Fijas
                    'nMes = nMes + 1
                    'If nMes > 12 Then
                    '    nAnio = nAnio + 1
                    '    nMes = 1
                    'End If
                    If nDia = pnDiaFijo Then nMes = nMes + 1
                    If nMes > 12 Then
                        nAnio = nAnio + 1
                        nMes = 1
                    End If
                    '**************************************
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
                        If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
                            nDia = 30
                        End If
                    End If
                End If
                dFecTemp = CDate(Right("0" & Trim(str(nDia)), 2) & "/" & Right("0" & Trim(str(nMes)), 2) & "/" & Trim(str(nAnio)))
                Calendario(i).dFecha = dFecTemp
                Calendario(i).NroCuota = i + 1
                Calendario(i).IntGra = 0#
                Calendario(i).Gasto = 0
                
                If i = 0 Then
                    Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(i).dFecha), nSaldoCapital)
                    'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNrocuotas, pnMonto, DateDiff("d", dDesembolso, Calendario(i).dFecha))
                Else
                    Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha), nSaldoCapital)
                    'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNrocuotas, pnMonto, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha))
                End If
                
                'If i = pnNroCuotas - 1 Then
                If i = pnCuotaFin Then
                    Calendario(i).Captital = nSaldoCapital
                Else
                    Calendario(i).Captital = CDbl(Format((Calendario(i).NroCuota / nSumCuotas) * pnMonto, "#0.00"))
                End If
                nSumCapital = nSumCapital + Calendario(i).Captital
                Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                nSaldoCapital = nSaldoCapital - Calendario(i).Captital
                
                'Verificar en el caso de capitalizar la gracia
                If pnMontoCapInicial > 0 Then
                    'Recalculamos los resultados
                    If nSumCapital > pnMontoCapInicial Then
                        Calendario(i).Captital = Calendario(i).Captital - (nSumCapital - pnMontoCapInicial)
                        Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                        'Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                    End If
                    If Calendario(i).Captital < 0 Then
                        Calendario(i).IntGra = 0#
                        Calendario(i).Gasto = 0#
                        Calendario(i).IntComp = 0#
                        Calendario(i).Captital = 0#
                        Calendario(i).Cuota = 0#
                        Calendario(i).SaldoCap = 0#
                    End If
                End If
                '*********************************************

            Next i
        End If
    End If
    
    'Actualizar si existe Periodo de Gracia
    If pnDiasGracia > 0 _
       And pnCuotaIni = 0 And pnMontoCapInicial = 0 Then
       '(25-03-06) y no se adicionaron cuotas de gracia y
       ' no se Capitalizo la gracia sin incrementar el Capital
        If pnTipoGracia = PrimeraCuota Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            For i = pnNroCuotas To 1 Step -1
                Calendario(i) = Calendario(i - 1)
                Calendario(i).NroCuota = Calendario(i).NroCuota + 1
            Next i
            Calendario(0).dFecha = pdFecDesemb + pnDiasGracia
            Calendario(0).NroCuota = 1
            Calendario(0).IntGra = MatGracia(0)
            Calendario(0).Gasto = 0
            Calendario(0).IntComp = 0#
            Calendario(0).Cuota = Calendario(0).IntGra
            Calendario(0).Captital = 0#
            Calendario(0).SaldoCap = pnMonto
        End If
        If pnTipoGracia = Prorateada Or pnTipoGracia = Configurable Then
            For i = 0 To pnNroCuotas - 1
                Calendario(i).IntGra = MatGracia(i)
                Calendario(i).Cuota = Calendario(i).Cuota + MatGracia(i)
            Next i
        End If
        If pnTipoGracia = UltimaCuota Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            Calendario(pnNroCuotas).dFecha = Calendario(pnNroCuotas - 1).dFecha + pnDiasGracia
            Calendario(pnNroCuotas).NroCuota = pnNroCuotas + 1
            Calendario(pnNroCuotas).IntGra = MatGracia(pnNroCuotas)
            Calendario(pnNroCuotas).Gasto = 0
            Calendario(pnNroCuotas).IntComp = 0#
            Calendario(pnNroCuotas).Cuota = Calendario(pnNroCuotas).IntGra
            Calendario(pnNroCuotas).Captital = 0#
            Calendario(pnNroCuotas).SaldoCap = 0#
        End If
    End If

End Sub

Private Sub CalendarioDecreciente(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, Optional ByVal MatGracia As Variant, _
                Optional ByVal pnCuotaIni As Integer = 0, Optional ByVal pnCuotaFin As Integer = 0, Optional ByVal pnDiaFijo2 As Integer = 0, _
                Optional pnMontoCapInicial As Double = 0)

Dim nSaldoCapital As Double
Dim dDesembolso As Date
Dim i As Integer
Dim oCredito As COMNCredito.NCOMCredito
Dim dFecTemp As Date
Dim nMes As Integer
Dim nAnio As Integer
Dim nDia As Integer
'Para llevar el control del Capital en el caso de capitalizar la gracia
'pero sin incrementar el capital
Dim nSumCapital As Double

    nSumCapital = 0
    nSaldoCapital = pnMonto
    dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
    
            
    If pnTipoPeriodo = PeriodoFijo Then
        Set oCredito = New COMNCredito.NCOMCredito
        'For i = 0 To pnNroCuotas - 1
        For i = pnCuotaIni To pnCuotaFin
            Calendario(i).dFecha = dDesembolso + pnPeriodo
            Calendario(i).NroCuota = i + 1
            Calendario(i).IntGra = 0#
            Calendario(i).Gasto = 0#
            Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital)
            
            'If i = pnNroCuotas - 1 Then
            If i = pnCuotaFin Then
                Calendario(i).Captital = nSaldoCapital
            Else
                Calendario(i).Captital = CDbl(Format((pnMonto / pnNroCuotas), "#0.00"))
            End If
            nSumCapital = nSumCapital + Calendario(i).Captital
            Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
            Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
            nSaldoCapital = nSaldoCapital - Calendario(i).Captital
            dDesembolso = dDesembolso + pnPeriodo
            
            'Verificar en el caso de capitalizar la gracia
            If pnMontoCapInicial > 0 Then
                'Recalculamos los resultados
                If nSumCapital > pnMontoCapInicial Then
                    Calendario(i).Captital = Calendario(i).Captital - (nSumCapital - pnMontoCapInicial)
                    Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                    'Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                End If
                If Calendario(i).Captital < 0 Then
                    Calendario(i).IntGra = 0#
                    Calendario(i).Gasto = 0#
                    Calendario(i).IntComp = 0#
                    Calendario(i).Captital = 0#
                    Calendario(i).Cuota = 0#
                    Calendario(i).SaldoCap = 0#
                End If

            End If
            '*********************************************

        Next i
    Else
        'Si es Fecha Fija
        If pnTipoPeriodo = FechaFija Then
            Set oCredito = New COMNCredito.NCOMCredito
            nMes = Month(dDesembolso)
            nAnio = Year(dDesembolso)
            'nDia = pnDiaFijo
            '**************************************
            'Se agrego para manejar 2 Fechas Fijas
            If pnDiaFijo2 = 0 Then
                nDia = pnDiaFijo
            Else
                If Day(dDesembolso) <= pnDiaFijo2 - 8 Then
                    nDia = pnDiaFijo2
                Else
                    nDia = pnDiaFijo
                End If
            End If
            '**************************************

            'For i = 0 To pnNroCuotas - 1
            For i = pnCuotaIni To pnCuotaFin
                '**************************************
                'Se agrego para manejar 2 Fechas Fijas
                If pnDiaFijo2 = 0 Then
                    nDia = pnDiaFijo
                Else
                    If i > pnCuotaIni Then
                        If nDia = pnDiaFijo Then
                            nDia = pnDiaFijo2
                        Else
                            nDia = pnDiaFijo
                        End If
                    End If
                End If
                '**************************************
                'If Not (i = 0 And pnDiaFijo > Day(dDesembolso) And Not bProxMes) Then
                If Not (i = 0 And nDia > Day(dDesembolso) And Not bProxMes) Then
                '**************************************
                    'Se modifico para manejar 2 Fechas Fijas
                    'nMes = nMes + 1
                    'If nMes > 12 Then
                    '    nAnio = nAnio + 1
                    '    nMes = 1
                    'End If
                    If nDia = pnDiaFijo Then nMes = nMes + 1
                    If nMes > 12 Then
                        nAnio = nAnio + 1
                        nMes = 1
                    End If
                End If
                '**************************************
                
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
                dFecTemp = CDate(Right("0" & Trim(str(nDia)), 2) & "/" & Right("0" & Trim(str(nMes)), 2) & "/" & Trim(str(nAnio)))
                Calendario(i).dFecha = dFecTemp
                Calendario(i).NroCuota = i + 1
                Calendario(i).IntGra = 0#
                Calendario(i).Gasto = 0
                'If i = 0 Then
                If i = pnCuotaIni Then
                    Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(i).dFecha), nSaldoCapital)
                    'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNrocuotas, pnMonto, DateDiff("d", dDesembolso, Calendario(i).dFecha))
                Else
                    Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha), nSaldoCapital)
                    'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNrocuotas, pnMonto, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha))
                End If
                'If i = pnNroCuotas - 1 Then
                If i = pnCuotaFin Then
                    Calendario(i).Captital = nSaldoCapital
                Else
                    Calendario(i).Captital = CDbl(Format((pnMonto / pnNroCuotas), "#0.00"))
                End If
                nSumCapital = nSumCapital + Calendario(i).Captital
                Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                nSaldoCapital = nSaldoCapital - Calendario(i).Captital
                
                'Verificar en el caso de capitalizar la gracia
                If pnMontoCapInicial > 0 Then
                    'Recalculamos los resultados
                    If nSumCapital > pnMontoCapInicial Then
                        Calendario(i).Captital = Calendario(i).Captital - (nSumCapital - pnMontoCapInicial)
                        Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                        'Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                    End If
                    If Calendario(i).Captital < 0 Then
                        Calendario(i).IntGra = 0#
                        Calendario(i).Gasto = 0#
                        Calendario(i).IntComp = 0#
                        Calendario(i).Captital = 0#
                        Calendario(i).Cuota = 0#
                        Calendario(i).SaldoCap = 0#
                    End If
                End If
                '*********************************************

            Next i
        End If
    End If
    
    'Actualizar si existe Periodo de Gracia
    If pnDiasGracia > 0 _
       And pnCuotaIni = 0 Then
        If pnTipoGracia = PrimeraCuota Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            For i = pnNroCuotas To 1 Step -1
                Calendario(i) = Calendario(i - 1)
                Calendario(i).NroCuota = Calendario(i).NroCuota + 1
            Next i
            Calendario(0).dFecha = pdFecDesemb + pnDiasGracia
            Calendario(0).NroCuota = 1
            Calendario(0).IntGra = MatGracia(0)
            Calendario(0).Gasto = 0
            Calendario(0).IntComp = 0#
            Calendario(0).Cuota = Calendario(0).IntGra
            Calendario(0).Captital = 0#
            Calendario(0).SaldoCap = pnMonto
        End If
        If pnTipoGracia = Prorateada Or pnTipoGracia = Configurable Then
            For i = 0 To pnNroCuotas - 1
                Calendario(i).IntGra = MatGracia(i)
                Calendario(i).Cuota = Calendario(i).Cuota + MatGracia(i)
            Next i
        End If
        If pnTipoGracia = UltimaCuota Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            Calendario(pnNroCuotas).dFecha = Calendario(pnNroCuotas - 1).dFecha + pnDiasGracia
            Calendario(pnNroCuotas).NroCuota = pnNroCuotas + 1
            Calendario(pnNroCuotas).IntGra = MatGracia(pnNroCuotas)
            Calendario(pnNroCuotas).Gasto = 0
            Calendario(pnNroCuotas).IntComp = 0#
            Calendario(pnNroCuotas).Cuota = Calendario(pnNroCuotas).IntGra
            Calendario(pnNroCuotas).Captital = 0#
            Calendario(pnNroCuotas).SaldoCap = 0#
        End If
    End If
End Sub
'
'Private Sub CalendarioCuotaFija(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
'                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
'                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
'                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
'                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
'                Optional ByVal pMatDesPar As Variant = "", Optional ByVal pnNumMes As Integer = 1, Optional ByVal pbMiViv As Boolean = False)
'
'Dim nSaldoCapital As Double
'Dim dDesembolso As Date
'Dim I As Integer
'Dim oCredito As NCredito
'Dim dFecTemp As Date
'Dim nMes As Integer
'Dim nAnio As Integer
'Dim nDia As Integer
'Dim nMontoCuotaTemp As Double
'Dim nCDIntPend As Double
'Dim nPeriodoDesPar As Integer
'Dim oCalend As Dcalendario
'
'    If pbDesemParcial Then
'        Set oCredito = New NCredito
'        ReDim Calendario(1)
'        Set oCalend = New Dcalendario
'        Calendario(0).dFecha = CDate(pMatDesPar(UBound(pMatDesPar) - 1, 2)) + pnPeriodo
'        Set oCalend = Nothing
'        Calendario(0).IntComp = 0
'        Calendario(0).Captital = 0
'        nSaldoCapital = 0
'        For I = 0 To UBound(pMatDesPar) - 1
'            nSaldoCapital = CDbl(pMatDesPar(I, 1))
'            Calendario(0).Captital = Calendario(0).Captital + nSaldoCapital
'            nPeriodoDesPar = DateDiff("d", CDate(pMatDesPar(I, 0)), Calendario(0).dFecha)
'            Calendario(0).IntComp = Calendario(0).IntComp + oCredito.MontoIntPerDias(pnTasaInt, nPeriodoDesPar, nSaldoCapital)
'        Next I
'        Calendario(0).NroCuota = 1
'
'        Calendario(0).IntGra = 0#
'        Calendario(0).Gasto = 0#
'        Calendario(0).Cuota = CDbl(Format(pnMonto + Calendario(0).IntComp, "#0.00"))
'        Set oCredito = Nothing
'    Else
'
'        nSaldoCapital = pnMonto
'        dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
'        If pnTipoPeriodo = PeriodoFijo Then
'            Set oCredito = New NCredito
'            For I = 0 To pnNroCuotas - 1
'                Calendario(I).dFecha = dDesembolso + pnPeriodo
'                Calendario(I).NroCuota = I + 1
'                Calendario(I).IntGra = 0#
'                Calendario(I).Gasto = 0#
'                Calendario(I).IntComp = oCredito.MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital)
'                Calendario(I).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo)
'                If I = pnNroCuotas - 1 Then
'                    Calendario(I).Captital = nSaldoCapital
'                    If (Calendario(I).IntComp + Calendario(I).Captital) > Calendario(I).Cuota Then
'                        Calendario(I).Cuota = Calendario(I).Captital + Calendario(I).IntComp
'                    Else
'                        Calendario(I).IntComp = Calendario(I).Cuota - Calendario(I).Captital
'                    End If
'                Else
'                    Calendario(I).Captital = Calendario(I).Cuota - Calendario(I).IntComp
'                End If
'                Calendario(I).SaldoCap = nSaldoCapital - Calendario(I).Captital
'                nSaldoCapital = nSaldoCapital - Calendario(I).Captital
'                nSaldoCapital = CDbl(Format(nSaldoCapital, "#0.00"))
'                dDesembolso = dDesembolso + pnPeriodo
'            Next I
'
'        Else
'            If pbCuotaFijaFechaFija Then
'                Set oCredito = New NCredito
'                nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv)
'                Set oCredito = Nothing
'            End If
'            'Si es Fecha Fija
'            If pnTipoPeriodo = FechaFija Then
'                Set oCredito = New NCredito
'                nMes = Month(dDesembolso)
'                nAnio = Year(dDesembolso)
'                nDia = pnDiaFijo
'                nCDIntPend = 0
'                For I = 0 To pnNroCuotas - 1
'                    nDia = pnDiaFijo
'                    If Not (I = 0 And pnDiaFijo > Day(dDesembolso) And (Not bProxMes)) Or pnNumMes = 6 Then
'                        nMes = nMes + pnNumMes
'                        If nMes > 12 Then
'                            nAnio = nAnio + 1
'                            nMes = nMes - 12
'                        End If
'                    Else
'                        If nMes = 2 Then
'                            If nDia >= 29 Then
'                                If nAnio Mod 4 <> 0 Then
'                                    nMes = nMes + pnNumMes
'                                End If
'                            End If
'                        Else
'                            If nDia > 30 Then
'                                If nMes = 4 Or nMes = 6 Or nMes = 9 Or 11 Then
'                                    nMes = nMes + pnNumMes
'                                End If
'                            End If
'                        End If
'                    End If
'                    If nMes = 2 Then
'                        If nDia > 28 Then
'                            If nAnio Mod 4 = 0 Then
'                                nDia = 29
'                            Else
'                                nDia = 28
'                            End If
'                        End If
'                    Else
'                        If nDia > 30 Then
'                            If nMes = 4 Or nMes = 6 Or nMes = 9 Or 11 Then
'                                nDia = 30
'                            End If
'                        End If
'                    End If
'                    dFecTemp = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)))
'                    Calendario(I).dFecha = dFecTemp
'                    Calendario(I).NroCuota = I + 1
'                    Calendario(I).IntGra = 0#
'                    Calendario(I).Gasto = 0
'                    If I = 0 Then
'                        If pbCuotaFijaFechaFija And pbMiViv Then
'                            dDesembolso = dDesembolso - pnDiasGracia
'                        End If
'                        Calendario(I).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(I).dFecha), nSaldoCapital)
'                        If Not pbCuotaFijaFechaFija Then
'                            Calendario(I).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, DateDiff("d", dDesembolso, Calendario(I).dFecha))
'                        Else
'                            Calendario(I).Cuota = nMontoCuotaTemp
'                        End If
'
'                    Else
'                        Calendario(I).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(I - 1).dFecha, Calendario(I).dFecha), nSaldoCapital)
'                        If Not pbCuotaFijaFechaFija Then
'                            Calendario(I).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, DateDiff("d", Calendario(I - 1).dFecha, Calendario(I).dFecha))
'                        Else
'                            Calendario(I).Cuota = nMontoCuotaTemp
'                        End If
'                    End If
'
'                    If pbCuotaFijaFechaFija Then
'                        Calendario(I).IntComp = Calendario(I).IntComp + nCDIntPend
'                        nCDIntPend = 0#
'                        If Calendario(I).IntComp > nMontoCuotaTemp Then
'                            nCDIntPend = Calendario(I).IntComp - nMontoCuotaTemp
'                            Calendario(I).IntComp = nMontoCuotaTemp
'                        End If
'                    End If
'
'                    If I = pnNroCuotas - 1 Then
'                        Calendario(I).Captital = nSaldoCapital
'                        If (Calendario(I).IntComp + Calendario(I).Captital) > Calendario(I).Cuota Then
'                            Calendario(I).Cuota = Calendario(I).Captital + Calendario(I).IntComp
'                        Else
'                            Calendario(I).IntComp = Calendario(I).Cuota - Calendario(I).Captital
'                        End If
'                    Else
'                        Calendario(I).Captital = Calendario(I).Cuota - Calendario(I).IntComp
'                    End If
'
'                    Calendario(I).SaldoCap = nSaldoCapital - Calendario(I).Captital
'                    nSaldoCapital = nSaldoCapital - Calendario(I).Captital
'                Next I
'            End If
'        End If
'    End If
'
'    'Actualizar si existe Periodo de Gracia
'    If pnDiasGracia > 0 Then
'        If pnTipoGracia = PrimeraCuota Then
'            ReDim Preserve Calendario(pnNroCuotas + 1)
'            For I = pnNroCuotas To 1 Step -1
'                Calendario(I) = Calendario(I - 1)
'                Calendario(I).NroCuota = Calendario(I).NroCuota + 1
'            Next I
'            Calendario(0).dFecha = pdFecDesemb + pnDiasGracia
'            Calendario(0).NroCuota = 1
'            Calendario(0).IntGra = MatGracia(0)
'            Calendario(0).Gasto = 0
'            Calendario(0).IntComp = 0#
'            Calendario(0).Cuota = Calendario(0).IntGra
'            Calendario(0).Captital = 0#
'            Calendario(0).SaldoCap = pnMonto
'        End If
'        If pnTipoGracia = Prorateada Or pnTipoGracia = Configurable Then
'            For I = 0 To pnNroCuotas - 1
'                Calendario(I).IntGra = MatGracia(I)
'                Calendario(I).Cuota = Calendario(I).Cuota + MatGracia(I)
'            Next I
'        End If
'        If pnTipoGracia = UltimaCuota Then
'            ReDim Preserve Calendario(pnNroCuotas + 1)
'            Calendario(pnNroCuotas).dFecha = Calendario(pnNroCuotas - 1).dFecha + pnDiasGracia
'            Calendario(pnNroCuotas).NroCuota = pnNroCuotas + 1
'            Calendario(pnNroCuotas).IntGra = MatGracia(pnNroCuotas)
'            Calendario(pnNroCuotas).Gasto = 0
'            Calendario(pnNroCuotas).IntComp = 0#
'            Calendario(pnNroCuotas).Cuota = Calendario(pnNroCuotas).IntGra
'            Calendario(pnNroCuotas).Captital = 0#
'            Calendario(pnNroCuotas).SaldoCap = 0#
'        End If
'    End If
'End Sub

'->***** LUCV20180601, Deshabilitado. según ERS022-2018
Private Sub CalendarioCuotaFija(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
                Optional ByVal pMatDesPar As Variant = "", Optional ByVal pnNumMes As Integer = 1, Optional ByVal pbMiViv As Boolean = False, _
                Optional ByVal pnCuotaIni As Integer = 0, Optional ByVal pnCuotaFin As Integer = 0, Optional ByVal pnDiaFijo2 As Integer = 0, _
                Optional pnMontoCapInicial As Double = 0, Optional ByVal pbRenovarCredito As Boolean = False, Optional ByVal pnInteresAFecha As Double = 0, _
                Optional ByVal pnMontoGra As Double = 0, Optional ByVal pnCuotaBalon As Integer = 0)
                'MAVM 20130302: pnMontoGra
                'WIOR 20131111 agregó pnCuotaBalon

Dim nSaldoCapital As Double
Dim dDesembolso As Date
Dim i As Integer
Dim oCredito As COMNCredito.NCOMCredito
Dim dFecTemp As Date
Dim nMes As Integer
Dim nAnio As Integer
Dim nDia As Integer
Dim nMontoCuotaTemp As Double
Dim nMontoGraciaTemp As Double 'MAVM 20130209
Dim nCDIntPend As Double
Dim nPeriodoDesPar As Integer
'Para llevar el control del Capital en el caso de capitalizar la gracia
'pero sin incrementar el capital
Dim nSumCapital As Double
'MAVM 20130209 ***
Dim nSumCapitalGra As Double
Dim nSaldoCapitalGra As Double
'***
Dim nNroCuotasAfec As Integer 'WIOR 20131127
nSumCapital = 0
'MAVM 20130209 ***
nMontoGraciaTemp = 0
nSumCapitalGra = 0
nSaldoCapitalGra = 0
'***

    If pbDesemParcial Then
        Set oCredito = New NCOMCredito
        ReDim Calendario(1)

        'Calendario(0).dFecha = CDate(pMatDesPar(UBound(pMatDesPar) - 1, 0)) + pnPeriodo

        'Cambiado para Santa -- LAYG
        'Valida que la fecha sea menor que el ultimo desemboloso
        'If CDate(pdFecDesemb) + pnPeriodo > CDate(pMatDesPar(UBound(pMatDesPar) - 1, 0)) Then
        '    MsgBox "Fecha de Ultimo desembolso es Posterior al Plazo del credito", vbInformation, "Aviso"
        'End If
        Calendario(0).dFecha = CDate(pdFecDesemb) + pnPeriodo

        Calendario(0).IntComp = 0
        Calendario(0).Captital = 0
        nSaldoCapital = 0
        For i = 0 To UBound(pMatDesPar) - 1
            nSaldoCapital = CDbl(pMatDesPar(i, 1))
            Calendario(0).Captital = Calendario(0).Captital + nSaldoCapital
            nPeriodoDesPar = DateDiff("d", CDate(pMatDesPar(i, 0)), Calendario(0).dFecha)
            Calendario(0).IntComp = Calendario(0).IntComp + oCredito.MontoIntPerDias(pnTasaInt, nPeriodoDesPar, nSaldoCapital)
        Next i
        Calendario(0).NroCuota = 1

        Calendario(0).IntGra = 0#
        Calendario(0).Gasto = 0#
        Calendario(0).Cuota = CDbl(Format(pnMonto + Calendario(0).IntComp, "#0.00"))
        Set oCredito = Nothing
    Else

        nSaldoCapital = pnMonto
        nSaldoCapitalGra = pnMontoGra 'MAVM 20130430
        dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia

        '25-05-2006
        If pnCuotaIni > 0 Then
            dDesembolso = Calendario(pnCuotaIni - 1).dFecha
        End If

        If pnTipoPeriodo = PeriodoFijo Then
            Set oCredito = New NCOMCredito
            'For i = 0 To pnNroCuotas - 1
            For i = pnCuotaIni To pnCuotaFin
                'WIOR 20131111 **********************************
                If i < pnCuotaBalon Then
                    nSaldoCapital = pnMonto
                End If
                'WIOR FIN ***************************************
                Calendario(i).dFecha = dDesembolso + pnPeriodo
                Calendario(i).NroCuota = i + 1
                Calendario(i).IntGra = 0#
                Calendario(i).Gasto = 0#
                Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital)

                'MAVM 20130430***
                If pnTipoGracia = 6 Then
                    Calendario(i).IntCompGra = oCredito.MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapitalGra)
                End If
                '***

                'WIOR 20131111 **********************************
                nNroCuotasAfec = pnNroCuotas - pnCuotaBalon
                If i < pnCuotaBalon Then
                    Calendario(i).Cuota = Calendario(i).IntComp
                Else
                    Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, nNroCuotasAfec, pnMonto, pnPeriodo)
                End If
                'WIOR FIN ***************************************
                'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo)'WIOR 20131111 COMENTO
                'MAVM 20130430***
                If pnTipoGracia = 6 Then
                    Calendario(i).CuotaGra = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMontoGra, pnPeriodo)
                End If
                '***

                'If i = pnNroCuotas - 1 Then
                If i = pnCuotaFin Then
                    Calendario(i).Captital = nSaldoCapital

                    'MAVM 20130430***
                    If pnTipoGracia = 6 Then
                        Calendario(i).IntGra = nSaldoCapitalGra
                    End If
                    '***

                    If (Calendario(i).IntComp + Calendario(i).Captital) > Calendario(i).Cuota Then
                        Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                    Else
                        Calendario(i).IntComp = Calendario(i).Cuota - Calendario(i).Captital

                        'MAVM 20130430***
                        If pnTipoGracia = 6 Then
                            Calendario(i).IntCompGra = Calendario(i).CuotaGra - Calendario(i).IntGra
                        End If
                        '***

                    End If
                Else
                    Calendario(i).Captital = Calendario(i).Cuota - Calendario(i).IntComp

                    'MAVM 20130430***
                    If pnTipoGracia = 6 Then
                        Calendario(i).IntGra = Calendario(i).CuotaGra - Calendario(i).IntCompGra
                    End If
                    '***

                End If
                nSumCapital = nSumCapital + Calendario(i).Captital
                Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                nSaldoCapital = nSaldoCapital - Calendario(i).Captital
                nSaldoCapital = CDbl(Format(nSaldoCapital, "#0.00"))

                'MAVM 20130430***
                If pnTipoGracia = 6 Then
                    nSumCapitalGra = nSumCapitalGra + Calendario(i).IntGra
                    Calendario(i).SaldoCapGra = nSaldoCapitalGra - Calendario(i).IntGra
                    nSaldoCapitalGra = nSaldoCapitalGra - Calendario(i).IntGra
                End If

                'Agrega la parte decimal a la primera cuota gitu
                'Descomentar cuando esten seguros de los cambios GITU
'                If I > 0 Then
'                    If InStr(Trim(Calendario(I).IntComp), ".") <> 0 Then
'                        Calendario(0).IntComp = Calendario(0).IntComp + (Round(Calendario(I).IntComp - IIf(InStr(Trim(Calendario(I).IntComp), ".") <> 0, Val(Left(Trim(Str(Calendario(I).IntComp)), InStr(Trim(Str(Calendario(I).IntComp)), ".") - 1)), 0), 2))
'                        Calendario(I).IntComp = Val(Left(Trim(Str(Calendario(I).IntComp)), InStr(Trim(Str(Calendario(I).IntComp)), ".") - 1))
'                    End If
'                    If InStr(Trim(Calendario(I).Cuota), ".") <> 0 Then
'                        Calendario(0).Cuota = Calendario(0).Cuota + (Round(Calendario(I).Cuota - IIf(InStr(Trim(Calendario(I).Cuota), ".") <> 0, Val(Left(Trim(Str(Calendario(I).Cuota)), InStr(Trim(Str(Calendario(I).Cuota)), ".") - 1)), 0), 2))
'                        Calendario(I).Cuota = Val(Left(Trim(Str(Calendario(I).Cuota)), InStr(Trim(Str(Calendario(I).Cuota)), ".") - 1))
'                    End If
'                    If InStr(Trim(Calendario(I).Captital), ".") <> 0 Then
'                        Calendario(0).Captital = Calendario(0).Captital + (Round(Calendario(I).Captital - IIf(InStr(Trim(Calendario(I).Captital), ".") <> 0, Val(Left(Trim(Str(Calendario(I).Captital)), InStr(Trim(Str(Calendario(I).Captital)), ".") - 1)), 0), 2))
'                        Calendario(I).Captital = Val(Left(Trim(Str(Calendario(I).Captital)), InStr(Trim(Str(Calendario(I).Captital)), ".") - 1))
'                    End If
'                    If InStr(Trim(Calendario(I).SaldoCap), ".") <> 0 Then
'                        Calendario(0).SaldoCap = Calendario(0).SaldoCap + (Round(Calendario(I).SaldoCap - IIf(InStr(Trim(Calendario(I).SaldoCap), ".") <> 0, Val(Left(Trim(Str(Calendario(I).SaldoCap)), InStr(Trim(Str(Calendario(I).SaldoCap)), ".") - 1)), 0), 2))
'                        Calendario(I).SaldoCap = Val(Left(Trim(Str(Calendario(I).SaldoCap)), InStr(Trim(Str(Calendario(I).SaldoCap)), ".") - 1))
'                    End If
'
''                    Calendario(0).IntComp = Calendario(0).IntComp + (Round(Calendario(I).IntComp - CInt(Calendario(I).IntComp), 2))
''                    Calendario(0).Cuota = Calendario(0).Cuota + ((Calendario(I).Cuota - CInt(Calendario(I).Cuota)))
''                    Calendario(0).Captital = Calendario(0).Captital + (Round(Calendario(I).Captital - CInt(Calendario(I).Captital), 2))
''                    Calendario(0).SaldoCap = Calendario(0).SaldoCap + (Round(Calendario(I).SaldoCap - CInt(Calendario(I).SaldoCap), 2))
'
''                    Calendario(I).IntComp = CInt(Calendario(I).IntComp)
''                    Calendario(I).Cuota = CInt(Calendario(I).Cuota)
''                    Calendario(I).Captital = CInt(Calendario(I).Captital)
''                    Calendario(I).SaldoCap = CInt(Calendario(I).SaldoCap)
'                End If

                dDesembolso = dDesembolso + pnPeriodo

                'Verificar en el caso de capitalizar la gracia
                If pnMontoCapInicial > 0 Then
                    'Recalculamos los resultados
                    If nSumCapital > pnMontoCapInicial Then
                        Calendario(i).Captital = Calendario(i).Captital - (nSumCapital - pnMontoCapInicial)
                        Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                        'Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                    End If
                    If Calendario(i).Captital < 0 Then
                        Calendario(i).IntGra = 0#
                        Calendario(i).Gasto = 0#
                        Calendario(i).IntComp = 0#
                        Calendario(i).Captital = 0#
                        Calendario(i).Cuota = 0#
                        Calendario(i).SaldoCap = 0#
                    End If
                End If
                '*********************************************
            Next i

        Else
            '11-05-2006
            'If pbCuotaFijaFechaFija Then
            '    Set oCredito = New COMNCredito.NCOMCredito
            '    nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas)
            '    Set oCredito = Nothing
            'End If
            '******************************
            'Si es Fecha Fija
            If pnTipoPeriodo = FechaFija Then

                dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia

                '23-05-2006
                If pnCuotaIni > 0 Then
                    dDesembolso = Calendario(pnCuotaIni - 1).dFecha
                End If

                Set oCredito = New COMNCredito.NCOMCredito
                nMes = Month(dDesembolso)
                nAnio = Year(dDesembolso)
                'nDia = pnDiaFijo
                '**************************************
                'Se agrego para manejar 2 Fechas Fijas
                If pnDiaFijo2 = 0 Then
                    nDia = pnDiaFijo
                Else
                    If Day(dDesembolso) <= pnDiaFijo2 - 8 Then
                        nDia = pnDiaFijo2
                    Else
                        nDia = pnDiaFijo
                    End If
                End If

            '**************************************
                nCDIntPend = 0
                'For i = 0 To pnNroCuotas - 1
                For i = pnCuotaIni To pnCuotaFin
                    'nDia = pnDiaFijo
                    '**************************************
                    'Se agrego para manejar 2 Fechas Fijas
                    If pnDiaFijo2 = 0 Then
                        nDia = pnDiaFijo
                    Else
                        If i > pnCuotaIni Then
                            If nDia = pnDiaFijo Then
                                nDia = pnDiaFijo2
                            Else
                                nDia = pnDiaFijo
                            End If
                        End If
                    End If
                    '**************************************

                    'Se modifico para manejar 2 dias fijos
                    'If Not (i = 0 And pnDiaFijo > Day(dDesembolso) And (Not bProxMes)) Or pnNumMes = 6 Then
                    '    nMes = nMes + pnNumMes
                    '    If nMes > 12 Then
                    '        nAnio = nAnio + 1
                    '        nMes = 1
                    '    End If
                    '15-05-2006
                    If Not (i = 0 And nDia >= Day(dDesembolso) And (Not bProxMes)) Or pnNumMes = 6 Then

                        'MAVM 20101201 ***
                        If i = 0 And pnDiasGracia <> 0 Then
                            dFecTemp = dDesembolso + 30
                            nDia = Day(dFecTemp)
                            nMes = Month(dFecTemp)
                            nAnio = Year(dFecTemp)
                        Else

                            If nDia = pnDiaFijo Then nMes = nMes + pnNumMes

                            If nMes > 12 Then
                                nAnio = nAnio + 1
                                'nMes = 1
                                nMes = nMes - 12
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
                                    If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
                                        nDia = 30
                                    End If
                                End If
                            End If
                            
                            If CDate(Right("0" & Trim(str(nDia)), 2) & "/" & Right("0" & Trim(str(nMes)), 2) & "/" & Trim(str(nAnio))) - dDesembolso < 5 Then
                                nMes = nMes + pnNumMes
                            End If
                        End If
                        '***

'                        If nMes > 12 Then
'                            nAnio = nAnio + 1
'                            nMes = 1
'                        End If
'                    Else
'                        If nMes = 2 Then
'                            If nDia >= 29 Then
'                                If nAnio Mod 4 <> 0 Then
'                                    nMes = nMes + pnNumMes
'                                End If
'                            End If
'                        Else
'                            If nDia > 30 Then
'                                If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
'                                    nMes = nMes + pnNumMes
'                                End If
'                            End If
'                        End If
'Modificacion por LMMD
'                            If nDia > 30 Then
'                                If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
'                                    nMes = nMes + pnNumMes
'                                End If
'                            End If
'                        End If

                    End If
'                    If nMes = 2 Then
'                        If nDia > 28 Then
'                            If nAnio Mod 4 = 0 Then
'                                nDia = 29
'                            Else
'                                nDia = 28
'                            End If
'                        End If
'                    Else
'                        If nDia > 30 Then
'                            If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
'                                nDia = 30
'                            End If
'                        End If
'                    End If

                    dFecTemp = CDate(Right("0" & Trim(str(nDia)), 2) & "/" & Right("0" & Trim(str(nMes)), 2) & "/" & Trim(str(nAnio)))

                    'If (dFecTemp > pdFecDesemb + CDate(pnDiasGracia) And bProxMes) Then
                    'If (dFecTemp > dDesembolso And bProxMes) Then
                    '    Calendario(i).dFecha = DateAdd("m", -Fix(pnDiasGracia / 30), dFecTemp)
                    'Else
                        Calendario(i).dFecha = dFecTemp
                    'End If
                    Calendario(i).NroCuota = i + 1
                    Calendario(i).IntGra = 0#
                    Calendario(i).Gasto = 0
                    'If i = 0 Then
                    'WIOR 20131111 **********************************
                    If i < pnCuotaBalon Then
                        nSaldoCapital = pnMonto
                    End If
                    nNroCuotasAfec = pnNroCuotas - pnCuotaBalon
                    'WIOR FIN ***************************************

                    If i = pnCuotaIni Then
                    '11-05-2006
                        If pbCuotaFijaFechaFija Then
                            '29-05-2006
                            Dim pbGraciaEnCuotas As Boolean

                            pbGraciaEnCuotas = IIf(pnCuotaIni > 0, True, False)

                            '23-05-2006
                            If pnCuotaIni > 0 Then
                                'nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas) 'WIOR 20131127 COMENTÓ
                                nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, nNroCuotasAfec, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas) 'WIOR 20131127
                            Else
                                'nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas)
                                '27-05-2006 + 29-05-2006
                                'nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas) 'WIOR 20131127 COMENTÓ
                                nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, nNroCuotasAfec, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas) 'WIOR 20131127

                                'MAVM 20130430***
                                If pnTipoGracia = 6 Then
                                    nMontoGraciaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMontoGra, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas)
                                End If
                                '***

                            End If
                        End If
                    '**************************************
                        If pbCuotaFijaFechaFija And pbMiViv Then
                            dDesembolso = dDesembolso - pnDiasGracia
                        End If
                        Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(i).dFecha), nSaldoCapital)


                        'MAVM 20130430***
                        If pnTipoGracia = 6 Then
                            Calendario(i).IntCompGra = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(i).dFecha), pnMontoGra)
                        End If
                        '***
                        If Not pbCuotaFijaFechaFija Then
                            'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, DateDiff("d", dDesembolso, Calendario(i).dFecha)) 'WIOR 20131111 COMENTÓ
                            'WIOR 20131111 **********************************
                            If i < pnCuotaBalon Then
                                Calendario(i).Cuota = Calendario(i).IntComp
                            Else
                                Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, nNroCuotasAfec, pnMonto, DateDiff("d", dDesembolso, Calendario(i).dFecha))
                            End If
                            'WIOR FIN ***************************************
                        Else
                            If nMontoCuotaTemp = 0 Then
                                'Calendario(i).Cuota = nMontoCuotaTemp'WIOR 20131111 COMENTÓ
                                'WIOR 20131111 **********************************
                                If i < pnCuotaBalon Then
                                    Calendario(i).Cuota = Calendario(i).IntComp
                                Else
                                    Calendario(i).Cuota = nMontoCuotaTemp
                                End If
                                'WIOR FIN ***************************************
                            Else
                                'Calendario(i).Cuota = nMontoCuotaTemp 'WIOR 20131111 COMENTÓ
                                'WIOR 20131111 **********************************
                                If i < pnCuotaBalon Then
                                    Calendario(i).Cuota = Calendario(i).IntComp
                                Else
                                    Calendario(i).Cuota = nMontoCuotaTemp
                                End If
                                'WIOR FIN ***************************************

                                'MAVM 20130430***
                                If pnTipoGracia = 6 Then
                                    Calendario(i).CuotaGra = nMontoGraciaTemp
                                End If
                                '***

                            End If
                        End If

                    Else
                        Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha), nSaldoCapital)

                        'MAVM 20130430***
                        If pnTipoGracia = 6 Then
                            Calendario(i).IntCompGra = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha), nSaldoCapitalGra)
                        End If

                        If Not pbCuotaFijaFechaFija Then
                            'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha)) 'WIOR 20131111 COMENTÓ
                            'WIOR 20131111 **********************************
                            If i < pnCuotaBalon Then
                                Calendario(i).Cuota = Calendario(i).IntComp
                            Else
                                Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, nNroCuotasAfec, pnMonto, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha))
                            End If
                            'WIOR FIN ***************************************
                        Else
                            'Calendario(i).Cuota = nMontoCuotaTemp 'WIOR 20131111 COMENTÓ
                            'WIOR 20131111 **********************************
                            If i < pnCuotaBalon Then
                                Calendario(i).Cuota = Calendario(i).IntComp
                            Else
                                Calendario(i).Cuota = nMontoCuotaTemp
                            End If
                            'WIOR FIN ***************************************

                            'MAVM 20130430***
                            If pnTipoGracia = 6 Then
                                Calendario(i).CuotaGra = nMontoGraciaTemp
                            End If
                            '***

                        End If
                    End If

                    If pbCuotaFijaFechaFija Then
                        Calendario(i).IntComp = Calendario(i).IntComp + nCDIntPend
                        nCDIntPend = 0#
                        If Calendario(i).IntComp > nMontoCuotaTemp Then
                            nCDIntPend = Calendario(i).IntComp - nMontoCuotaTemp
                            Calendario(i).IntComp = nMontoCuotaTemp
                        End If
                    End If

                    'If i = pnNroCuotas - 1 Then
                    If i = pnCuotaFin Then
                        Calendario(i).Captital = nSaldoCapital

                        'MAVM 20130430***
                        If pnTipoGracia = 6 Then
                            Calendario(i).IntGra = nSaldoCapitalGra
                        End If
                        '***

                        If (Calendario(i).IntComp + Calendario(i).Captital) > Calendario(i).Cuota Then
                            Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                        Else
                            Calendario(i).IntComp = Calendario(i).Cuota - Calendario(i).Captital
                            'MAVM 20130430***
                            If pnTipoGracia = 6 Then
                                Calendario(i).IntCompGra = Calendario(i).CuotaGra - Calendario(i).IntGra
                            End If
                            '***
                        End If
                    Else
                        Calendario(i).Captital = Calendario(i).Cuota - Calendario(i).IntComp

                        'MAVM 20130430***
                        If pnTipoGracia = 6 Then
                            Calendario(i).IntGra = Calendario(i).CuotaGra - Calendario(i).IntCompGra
                        End If
                        '***

                    End If

                    nSumCapital = nSumCapital + Calendario(i).Captital
                    Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                    nSaldoCapital = nSaldoCapital - Calendario(i).Captital

                    'MAVM 20130430***
                    If pnTipoGracia = 6 Then
                        nSumCapitalGra = nSumCapitalGra + Calendario(i).IntGra
                        Calendario(i).SaldoCapGra = nSaldoCapitalGra - Calendario(i).IntGra
                        nSaldoCapitalGra = nSaldoCapitalGra - Calendario(i).IntGra
                    End If

                    'Agrega la parte decimal a la primera cuota GITU
                    'Descomentar cuando esten seguros de los cambios GITU
'                    If I > 0 Then
'                        If InStr(Trim(Calendario(I).IntComp), ".") <> 0 Then
'                            Calendario(0).IntComp = Calendario(0).IntComp + (Round(Calendario(I).IntComp - IIf(InStr(Trim(Calendario(I).IntComp), ".") <> 0, Val(Left(Trim(Str(Calendario(I).IntComp)), InStr(Trim(Str(Calendario(I).IntComp)), ".") - 1)), 0), 2))
'                            Calendario(I).IntComp = Val(Left(Trim(Str(Calendario(I).IntComp)), InStr(Trim(Str(Calendario(I).IntComp)), ".") - 1))
'                        End If
'                        If InStr(Trim(Calendario(I).Cuota), ".") <> 0 Then
'                            Calendario(0).Cuota = Calendario(0).Cuota + (Round(Calendario(I).Cuota - IIf(InStr(Trim(Calendario(I).Cuota), ".") <> 0, Val(Left(Trim(Str(Calendario(I).Cuota)), InStr(Trim(Str(Calendario(I).Cuota)), ".") - 1)), 0), 2))
'                            Calendario(I).Cuota = Val(Left(Trim(Str(Calendario(I).Cuota)), InStr(Trim(Str(Calendario(I).Cuota)), ".") - 1))
'                        End If
'                        If InStr(Trim(Calendario(I).Captital), ".") <> 0 Then
'                            Calendario(0).Captital = Calendario(0).Captital + (Round(Calendario(I).Captital - IIf(InStr(Trim(Calendario(I).Captital), ".") <> 0, Val(Left(Trim(Str(Calendario(I).Captital)), InStr(Trim(Str(Calendario(I).Captital)), ".") - 1)), 0), 2))
'                            Calendario(I).Captital = Val(Left(Trim(Str(Calendario(I).Captital)), InStr(Trim(Str(Calendario(I).Captital)), ".") - 1))
'                        End If
'                        If InStr(Trim(Calendario(I).SaldoCap), ".") <> 0 Then
'                            Calendario(0).SaldoCap = Calendario(0).SaldoCap + (Round(Calendario(I).SaldoCap - IIf(InStr(Trim(Calendario(I).SaldoCap), ".") <> 0, Val(Left(Trim(Str(Calendario(I).SaldoCap)), InStr(Trim(Str(Calendario(I).SaldoCap)), ".") - 1)), 0), 2))
'                            Calendario(I).SaldoCap = Val(Left(Trim(Str(Calendario(I).SaldoCap)), InStr(Trim(Str(Calendario(I).SaldoCap)), ".") - 1))
'                        End If
'
''                    Calendario(0).IntComp = Calendario(0).IntComp + (Round(Calendario(I).IntComp - CInt(Calendario(I).IntComp), 2))
''                    Calendario(0).Cuota = Calendario(0).Cuota + ((Calendario(I).Cuota - CInt(Calendario(I).Cuota)))
''                    Calendario(0).Captital = Calendario(0).Captital + (Round(Calendario(I).Captital - CInt(Calendario(I).Captital), 2))
''                    Calendario(0).SaldoCap = Calendario(0).SaldoCap + (Round(Calendario(I).SaldoCap - CInt(Calendario(I).SaldoCap), 2))
'
''                    Calendario(I).IntComp = CInt(Calendario(I).IntComp)
''                    Calendario(I).Cuota = CInt(Calendario(I).Cuota)
''                    Calendario(I).Captital = CInt(Calendario(I).Captital)
''                    Calendario(I).SaldoCap = CInt(Calendario(I).SaldoCap)
'                    End If
                    ' End GITU

                    'Verificar en el caso de capitalizar la gracia
                    If pnMontoCapInicial > 0 Then
                        'Recalculamos los resultados
                        If nSumCapital > pnMontoCapInicial Then
                            Calendario(i).Captital = Calendario(i).Captital - (nSumCapital - pnMontoCapInicial)
                            Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                            'Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                        End If
                        If Calendario(i).Captital < 0 Then
                            Calendario(i).IntGra = 0#
                            Calendario(i).Gasto = 0#
                            Calendario(i).IntComp = 0#
                            Calendario(i).Captital = 0#
                            Calendario(i).Cuota = 0#
                            Calendario(i).SaldoCap = 0#
                        End If
                    End If
                    '*********************************************

                Next i
            End If
        End If
    End If

    'Actualizar si existe Periodo de Gracia
    If pnDiasGracia > 0 _
       And pnCuotaIni = 0 Then
        If pnTipoGracia = PrimeraCuota Then
            'MAVM 20130216 ***
            'ReDim Preserve Calendario(pnNroCuotas + 1)
            'For i = pnNroCuotas To 1 Step -1
            '    Calendario(i) = Calendario(i - 1)
            '    Calendario(i).NroCuota = Calendario(i).NroCuota + 1
            'Next i
            'Calendario(0).dFecha = pdFecDesemb + pnDiasGracia
            'Calendario(0).NroCuota = 1
            Calendario(0).IntGra = MatGracia(0)
            Calendario(0).Cuota = Calendario(0).Cuota + MatGracia(0)
            'Calendario(0).Gasto = 0
            'Calendario(0).IntComp = 0#
            'Calendario(0).Cuota = Calendario(0).IntGra
            'Calendario(0).Captital = 0#
            'Calendario(0).SaldoCap = pnMonto
            '***
        End If
        If pnTipoGracia = Prorateada Or pnTipoGracia = Configurable Then
            For i = 0 To pnNroCuotas - 1
                Calendario(i).IntGra = MatGracia(i)
                Calendario(i).Cuota = Calendario(i).Cuota + MatGracia(i)
                'GITU
                'Se agrega los decimales de las demas cuotas a la primera cuota
                'Descomentar cuando esten seguros de los cambios GITU
'                If I > 0 Then
'                    If InStr(Trim(Str(Calendario(I).IntGra)), ".") <> 0 Then
'                        Calendario(0).IntGra = Calendario(0).IntGra + (Round(Calendario(I).IntGra - Val(Left(Trim(Str(Calendario(I).IntGra)), InStr(Trim(Str(Calendario(I).IntGra)), ".") - 1)), 2))
'                        Calendario(I).IntGra = Val(Left(Trim(Str(Calendario(I).IntGra)), InStr(Trim(Str(Calendario(I).IntGra)), ".") - 1))
'                    End If
'                    If InStr(Trim(Str(MatGracia(I))), ".") <> 0 Then
'                        Calendario(0).Cuota = Calendario(0).Cuota + (Round(MatGracia(I) - Val(Left(Trim(Str(MatGracia(I))), InStr(Trim(Str(MatGracia(I))), ".") - 1)), 2))
'                        Calendario(I).Cuota = Val(Left(Trim(Str(Calendario(I).Cuota)), InStr(Trim(Str(Calendario(I).Cuota)), ".") - 1))
'                    End If
'
''                    Calendario(0).IntGra = Calendario(0).IntGra + (Round(Calendario(I).IntGra - CLng(Calendario(I).IntGra), 2))
''                    Calendario(0).Cuota = Calendario(0).Cuota + (Round(MatGracia(I) - CLng(MatGracia(I)), 2))
'
''                    Calendario(I).IntGra = CLng(Calendario(I).IntGra)
''                    Calendario(I).Cuota = CLng(Calendario(I).Cuota)
'                End If
                'End GITU
            Next i
        End If
        If pnTipoGracia = UltimaCuota Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            Calendario(pnNroCuotas).dFecha = Calendario(pnNroCuotas - 1).dFecha + pnDiasGracia
            Calendario(pnNroCuotas).NroCuota = pnNroCuotas + 1
            Calendario(pnNroCuotas).IntGra = MatGracia(pnNroCuotas)
            Calendario(pnNroCuotas).Gasto = 0
            Calendario(pnNroCuotas).IntComp = 0#
            Calendario(pnNroCuotas).Cuota = Calendario(pnNroCuotas).IntGra
            Calendario(pnNroCuotas).Captital = 0#
            Calendario(pnNroCuotas).SaldoCap = 0#
        End If

        'MAVM 20130209 ***
        If pnTipoGracia = gColocTiposGraciaCapitalizada Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            For i = pnNroCuotas To 1 Step -1
                Calendario(i) = Calendario(i - 1)
            Next i
            Calendario(0).dFecha = pdFecDesemb + pnDiasGracia
            Calendario(0).NroCuota = 0
            Calendario(0).IntGra = pnMontoGra
            Calendario(0).Gasto = 0
            Calendario(0).IntComp = 0#
            Calendario(0).Cuota = 0#
            Calendario(0).Captital = pnMonto
            Calendario(0).SaldoCap = pnMonto
            Calendario(0).SaldoCapGra = 0#
            Calendario(0).CuotaGra = 0#
            Calendario(0).IntCompGra = 0#
        End If
        '***
    End If

    'ARCV 24-10-2006
    'ARCV 01-03-2007
    If pbRenovarCredito And pnInteresAFecha > 0 Then
    '        If pnDiasGracia > 0 And pnTipoGracia = UltimaCuota Then
    '            ReDim Preserve Calendario(pnNroCuotas + 2)
    '            Calendario(pnNroCuotas).dFecha = Calendario(pnNroCuotas).dFecha + pnDiasGracia
    '            Calendario(pnNroCuotas).NroCuota = pnNroCuotas + 2
    '        Else
    '            ReDim Preserve Calendario(pnNroCuotas + 1)
    '            Calendario(pnNroCuotas).dFecha = Calendario(pnNroCuotas - 1).dFecha + pnDiasGracia
    '            Calendario(pnNroCuotas).NroCuota = pnNroCuotas + 1
    '        End If
    '        Calendario(pnNroCuotas).IntGra = 0
    '        Calendario(pnNroCuotas).Gasto = 0
    '        Calendario(pnNroCuotas).IntComp = pnInteresAFecha
    '        Calendario(pnNroCuotas).Cuota = pnInteresAFecha
    '        Calendario(pnNroCuotas).Captital = 0#
    '        Calendario(pnNroCuotas).SaldoCap = 0#
        If pnDiasGracia > 0 And pnTipoGracia = UltimaCuota Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            Calendario(pnNroCuotas).dFecha = Calendario(pnNroCuotas).dFecha + pnDiasGracia
            Calendario(pnNroCuotas).NroCuota = pnNroCuotas + 1
        Else
            Calendario(pnNroCuotas).dFecha = Calendario(pnNroCuotas).dFecha + pnDiasGracia
        End If
        For i = 0 To pnNroCuotas - 1
            Calendario(i).IntComp = Calendario(i).IntComp + Round(pnInteresAFecha / pnNroCuotas, 2)
            Calendario(i).Cuota = Calendario(i).Cuota + Round(pnInteresAFecha / pnNroCuotas, 2)
            'ADD By GITU 07-08-2008
            'Descomentar cuando esten seguros de los cambios GITU
'            If I > 0 Then
'                If InStr(Trim(Str(Calendario(I).IntComp)), ".") > 0 Then
'                    Calendario(0).IntComp = Calendario(0).IntComp + (Round(Calendario(I).IntComp - Val(Left(Trim(Str(Calendario(I).IntComp)), InStr(Trim(Str(Calendario(I).IntComp)), ".") - 1)), 2))
'                    Calendario(I).IntComp = Val(Left(Trim(Str(Calendario(I).IntComp)), InStr(Trim(Str(Calendario(I).IntComp)), ".") - 1))
'                End If
'                If InStr(Trim(Str(Calendario(I).Cuota)), ".") <> 0 Then
'                    Calendario(0).Cuota = Calendario(0).Cuota + (Round(Calendario(I).Cuota - Val(Left(Trim(Str(Calendario(I).Cuota)), InStr(Trim(Str(Calendario(I).Cuota)), ".") - 1)), 2))
'                    Calendario(I).Cuota = Val(Left(Trim(Str(Calendario(I).Cuota)), InStr(Trim(Str(Calendario(I).Cuota)), ".") - 1))
'                End If
'            End If
            'End GITU
        Next i
    End If
    '---------------
End Sub
'<-***** Fin LUCV20180601

'->***** LUCV20180601, Según ERS022-2018 [Reemplazará al método: CalendarioCuotaFija]
Private Sub CalendarioCuotaFijaNuevo(ByVal pnMonto As Double, _
                                    ByVal pnTasaInt As Double, _
                                    ByVal pnNroCuotas As Integer, _
                                    ByVal pnPeriodo As Double, _
                                    ByVal pdFecDesemb As Date, _
                                    ByVal pnTipoPeriodo As TCalendTipoPeriodo, _
                                    ByVal pnTipoGracia As TCalendTipoGracia, _
                                    ByVal pnDiasGracia As Integer, _
                                    ByVal pnDiaFijo As Integer, _
                                    Optional ByVal pnNumMes As Integer = 1, _
                                    Optional ByVal pnCuotaIni As Integer = 0, _
                                    Optional ByVal pnCuotaFin As Integer = 0, _
                                    Optional ByVal psCtaCod As String = "", _
                                    Optional ByVal pnTasaSegDes As Double = 0, _
                                    Optional ByRef pMatCalendSegDes As Variant, _
                                    Optional ByVal pnExoSeguroDesgravamen As Integer = 0, _
                                    Optional ByVal pnMontoPoliza As Double, _
                                    Optional ByVal pnTasaSegInc As Double)
                                    'LUCV20180601, Agregó psCtaCod, pnTasaSegDes, pMatCalendSegDes, pnExoSeguroDesgravamen, pnMontoPoliza, pnTasaSegInc según ERS022-2018
            
Dim nSaldoCapital As Double
Dim nSumCapital As Double
Dim dDesembolso As Date
Dim i As Integer
Dim oNCOMCredito As COMNCredito.NCOMCredito
Dim dFecTemp As Date
Dim nMes As Integer
Dim nAnio As Integer
Dim nDia As Integer
Dim nMontoCuotaTemp As Double

'->***** LUCV20180601, Según ERS022-2018
Dim nMontoCuotaPrevia As Double
Dim nTEMTotal As Double            'Tasa Efectiva Mensual Total
Dim nTEDTotal As Double            'Tasa Efectiva Diaria Total
'Para Gastos
Dim oNCOMGasto As COMNCredito.NCOMGasto
Dim nEEMSegDesg As Double          'Equivalente Efectivo Mensual del Seguro de Desgravamen
Dim nEEMSegInc As Double           'Equivalente Efectivo Mensual del Seguro de Incendio
Dim nMontoSegDes As Double
Dim MatPersSegDes As Variant
Dim nEdadMinSegDes As Long
Dim nEdadMaxSegDes As Long
Dim nCoberMinSegDes As Double
Dim nCoberMaxSegDes As Double
Dim nPrimaPerGracia As Double      'Prima del periodo de Gracia
'Para Int. Gracia
Dim nIntPeriodoGracia As Double    'IPG: Interes por Periodo de Gracia
Dim nGraciaGenerada As Double      'GG:  Gracia Generada
Dim nGraciaCapitalizada As Double  'GC:  Gracia Capitalizada
Dim nGraciaNoPagada As Double      'GNP: Gracia No Pagada
'Método Iterativo
Dim nSumaFAS As Double             'Factor Acumulado [Sumatoria FAS]
Dim nFVAS As Double                'Factor Valor Actual Saldo
Dim nSKU As Double                 'Saldo capital de la última cuota (final)
Dim nVASKU As Double               'Valor Actual Saldo Capital
Dim nCVASKU As Double              'Cuota del Valor Actual Saldo Capital
Dim nTotalDias As Integer
Dim nValorAjusteIteracion As Double
Dim nValorAjusteIteracionTotal As Double
Dim nNumIteracion As Integer
Dim nContadorSKU As Integer
Dim nAjusteCuotaGracia As Double
'Ajuste Ultima Cuota
Dim nSKPenultimaCuota As Double
'Calculo Distribución Int. Comp.
Dim nIntGraciaAsignado As Double
Dim nIntCompCalculado As Double 'Interes Calculado(A1)
Dim nIntCompCalculadoAcumulado As Double
Dim nIntCompCalculadoSumatoria As Double
Dim nIntCompAsignado As Double  'Interes Asignado (B1)
Dim nIntCompAsignadoAcumulado As Double
Dim nDiferencia As Double       'Vericación Diferencia Calc. - Asig.
Dim nDiferenciaAcumulada As Double
Dim nIntCapitalizadoDiferencia As Double
'<-***** fin LUCV20180601
    
    '1.- Inicializa variables
    nIntPeriodoGracia = 0: nGraciaGenerada = 0: nGraciaNoPagada = 0
    nVASKU = 0: nSKU = 1.5
    nSumCapital = 0
    nValorAjusteIteracion = 0: nValorAjusteIteracionTotal = 0: nNumIteracion = 0: nContadorSKU = 0
    nSKPenultimaCuota = 0
    nPrimaPerGracia = 0
    
    '2.- Fecha de desembolso + los días del periodo de gracia
    dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
    
    '3.- Realiza Cálculos de parametros para el met. iterativo
    '3.1.- Equivalente Efectivo Mensual SegDes. y Tasa Efectiva (Mensual/Diaria) Total
    Set oNCOMCredito = New NCOMCredito
    nEEMSegDesg = oNCOMCredito.ObtieneEquivalenteEfectivoMensual(pnTasaSegDes)
    nEEMSegInc = oNCOMCredito.ObtieneEquivalenteEfectivoMensual(pnTasaSegInc)
    nTEMTotal = (pnTasaInt + Round(nEEMSegDesg, 6)) '+ Round(nEEMSegInc, 6)) 'Comentó a Peticion GELU
    nTEDTotal = (((1 + nTEMTotal / 100) ^ (1 / 30)) - 1)
    Set oNCOMCredito = Nothing
    
    '3.2.- Obtiene datos de Seg. Desg. según condiciones
    Set oNCOMGasto = New COMNCredito.NCOMGasto
    Call oNCOMGasto.CargaValoreSegDes(psCtaCod, pdFecDesemb, MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, nCoberMinSegDes, nCoberMaxSegDes)
    Set oNCOMGasto = Nothing
    
    '3.3.- Cuota Fija Mensual Inicial o temporal (para calendario FechaFija/PeriodoFijo)
    Set oNCOMCredito = New NCOMCredito
    nMontoCuotaTemp = oNCOMCredito.CuotaFijaTipoPeriodo(nTEMTotal, pdFecDesemb, pnNroCuotas, pnMonto, pnTipoPeriodo, pnPeriodo, pnDiaFijo, _
                                                        pnDiasGracia, pnNumMes, nTotalDias, nSumaFAS)
    '3.4.- Factor Valor Actual Saldo
    nFVAS = Round(((1 + nTEDTotal) ^ nTotalDias), 6)
    
    '3.5.- Interes de gracia por periodo (por días de gracia)
    nIntPeriodoGracia = oNCOMCredito.MontoIntPerDias(pnTasaInt, pnDiasGracia, pnMonto)
    nAjusteCuotaGracia = Round(nIntPeriodoGracia / nSumaFAS, 2)
    Set oNCOMCredito = Nothing
    
    '3.6.- Obtiene la Prima del periodo de gracia de manera diaria. (prorrateada según numero de cuotas)
    nPrimaPerGracia = (pnMontoPoliza * (pnDiasGracia / 30)) / pnNroCuotas

    '4.- Aplicacion del metodo iterativo
    Do While (Not (nSKU >= -1.01 And nSKU <= 1.01)) Or (nContadorSKU <= 1)
        Set oNCOMCredito = New NCOMCredito
        nSaldoCapital = pnMonto
        nSumCapital = 0
        'Seteo Var. Ajuste Interes Compensatorio
        nIntCapitalizadoDiferencia = 0
        nIntCompCalculadoAcumulado = 0
        nIntCompCalculadoSumatoria = 0
        nIntCompAsignadoAcumulado = 0
        nDiferenciaAcumulada = 0
        nIntCompCalculado = 0

        '4.1.- Recorrido desde la cuota Inicial hasta la Final
        For i = pnCuotaIni To pnCuotaFin
            '4.1.1.- Administración de fechas
            If pnTipoPeriodo = PeriodoFijo Then
                If i = 0 Then dFecTemp = dDesembolso
                Calendario(i).dFecha = dFecTemp + pnPeriodo
            Else
                nDia = pnDiaFijo
                If i = 0 Then
                    dFecTemp = dDesembolso + 30
                    nDia = Day(dFecTemp): nMes = Month(dFecTemp): nAnio = Year(dFecTemp)
                    
                    'Artificio para cuando el desembolso sea a fines de enero [Febrero]
                    If CDate(Right("0" & Trim(str(nDia)), 2) & "/" & Right("0" & Trim(str(nMes)), 2) & "/" & Trim(str(nAnio))) - dDesembolso < 5 Then
                        nMes = nMes + pnNumMes
                        oNCOMCredito.ValidaFechasFijasCuota nDia, nMes, nAnio
                    End If
                Else
                    If nDia = pnDiaFijo Then nMes = nMes + pnNumMes
                    oNCOMCredito.ValidaFechasFijasCuota nDia, nMes, nAnio
                End If
                
                dFecTemp = CDate(Right("0" & Trim(str(nDia)), 2) & "/" & Right("0" & Trim(str(nMes)), 2) & "/" & Trim(str(nAnio)))
                Calendario(i).dFecha = dFecTemp
            End If
            
            '4.1.2.- Seteo de variables del calendario
            Calendario(i).NroCuota = i + 1
            Calendario(i).IntComp = 0#
            Calendario(i).IntGra = 0#
            Calendario(i).SegDes = 0#
            Calendario(i).CuotaPrimaPoliza = 0#
            Calendario(i).CuotaPrimaPolizaGracia = 0#
            Calendario(i).Gasto = 0#
            
            '4.1.3.- Condiciones Cuota Inicial
            If i = pnCuotaIni Then
                'Monto de la cuota Fija + Ajuste Cuota Gracia + Ajuste por Iteracion
                nMontoCuotaPrevia = (nMontoCuotaTemp + nAjusteCuotaGracia) + nValorAjusteIteracionTotal
                Calendario(i).Cuota = nMontoCuotaPrevia
                
                'Interes Compensatorio Calc.
                nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", dDesembolso, Calendario(i).dFecha)), nSaldoCapital)
                
                'Seguro Desgravamen
                Set oNCOMGasto = New COMNCredito.NCOMGasto
                nMontoSegDes = oNCOMGasto.CalculaSeguroDesgravamen(i, "K", pnTasaSegDes, pnMonto, Calendario(i).IntComp, nSaldoCapital, _
                                            Calendario(i).dFecha, 0, 0, MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, _
                                            nCoberMinSegDes, nCoberMaxSegDes, pMatCalendSegDes, psCtaCod)
                                            
                'Exoneración Seguro de Desgravamen(Cuota Inicial)
                If pnExoSeguroDesgravamen = 1 Then
                    Calendario(i).SegDes = Format(0, "#0.00")
                Else
                    Calendario(i).SegDes = Round((nMontoSegDes / 30) * (DateDiff("d", pdFecDesemb, Calendario(i).dFecha)), 2)
                End If
                
                'Gasto: Póliza contra incendio
                Calendario(i).CuotaPrimaPoliza = pnMontoPoliza
                
                'Gasto: Póliza contra incendio por periodo de gracia (prorrateado)
                Calendario(i).CuotaPrimaPolizaGracia = Round(nPrimaPerGracia, 2)
                
                'Interés de Gracia Cuota Inicial
                nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(pnCuotaIni, pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", dDesembolso, Calendario(i).dFecha)), _
                                                                         nIntPeriodoGracia, nMontoCuotaPrevia, Calendario(i).IntComp, Calendario(i).SegDes, (Calendario(i).CuotaPrimaPoliza + Calendario(i).CuotaPrimaPolizaGracia), _
                                                                         nGraciaGenerada, nGraciaCapitalizada, nGraciaNoPagada)
                'Ajuste Interés Compensario
                nIntCompCalculadoAcumulado = nIntCompCalculado
                nIntCompCalculadoSumatoria = nIntCompCalculadoAcumulado
                
                If nIntGraciaAsignado > 0 Then
                    nIntCompAsignado = nMontoCuotaPrevia - nIntCapitalizadoDiferencia - nIntGraciaAsignado - Calendario(i).SegDes
                Else
                    If (nIntCompCalculado + Abs(nDiferenciaAcumulada) + Calendario(i).SegDes) > nMontoCuotaPrevia Then
                        nIntCompAsignado = nMontoCuotaPrevia - Calendario(i).SegDes - nIntCapitalizadoDiferencia
                    Else
                        nIntCompAsignado = nIntCompCalculado + Abs(nDiferenciaAcumulada)
                    End If
                End If
                    
                If nIntCompAsignado < 0 Then
                    nIntCompAsignado = 0
                End If
                  
                If nIntCompAsignado > (nIntCompCalculado + Abs(nDiferenciaAcumulada)) Then
                    nIntCompAsignado = (nIntCompCalculado + Abs(nDiferenciaAcumulada))
                End If
                
                nIntCompAsignadoAcumulado = nIntCompAsignado + nIntCompAsignadoAcumulado
                nDiferenciaAcumulada = Round(nIntCompCalculadoAcumulado - nIntCompAsignadoAcumulado, 2)
                
                'Interés Compensatorio
                Calendario(i).IntComp = nIntCompAsignado
                
                'Interés de Gracia
                Calendario(i).IntGra = nIntGraciaAsignado
                
                'Amortizacion Capital
                If Round(nMontoCuotaPrevia - (nIntCompAsignado + nIntGraciaAsignado + Calendario(i).SegDes) - nIntCapitalizadoDiferencia, 4) < 0 Then
                    Calendario(i).Captital = 0
                Else
                    Calendario(i).Captital = Round(nMontoCuotaPrevia - (nIntCompAsignado + nIntGraciaAsignado + Calendario(i).SegDes) - nIntCapitalizadoDiferencia, 4)
                End If
                Set oNCOMGasto = Nothing
            Else
                'Monto de la Cuota Fija Final
                Calendario(i).Cuota = nMontoCuotaPrevia
                
                'Interés Compensatorio
                nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha)), nSaldoCapital)
                
                'Seguro Desgravamen
                Set oNCOMGasto = New COMNCredito.NCOMGasto
                nMontoSegDes = oNCOMGasto.CalculaSeguroDesgravamen(i, "K", pnTasaSegDes, pnMonto, Calendario(i).IntComp, nSaldoCapital, _
                                            Calendario(i).dFecha, 0, 0, MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, _
                                            nCoberMinSegDes, nCoberMaxSegDes, pMatCalendSegDes, psCtaCod)
                'Exoneración Seguro de Desgravamen
                If pnExoSeguroDesgravamen = 1 Then
                    Calendario(i).SegDes = Format(0, "#0.00")
                Else
                    Calendario(i).SegDes = Round(nMontoSegDes * (IIf(pnPeriodo > 30, Round(pnPeriodo / 30, 0), 1)), 2)
                End If
                
                'Gasto: Póliza contra incendio
                Calendario(i).CuotaPrimaPoliza = pnMontoPoliza
                
                'Gasto: Póliza contra incendio por periodo de gracia (prorrateado)
                Calendario(i).CuotaPrimaPolizaGracia = Round(nPrimaPerGracia, 2)
                
                'Interés de Gracia
                nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(i, pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha)), _
                                                                         nIntPeriodoGracia, nMontoCuotaPrevia, Calendario(i).IntComp, Calendario(i).SegDes, (Calendario(i).CuotaPrimaPoliza + Calendario(i).CuotaPrimaPolizaGracia), _
                                                                         nGraciaGenerada, nGraciaCapitalizada, nGraciaNoPagada)
                                                                         
                'Ajuste Interés Compensario
                nIntCapitalizadoDiferencia = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha)), nDiferenciaAcumulada)
                
                If nIntGraciaAsignado > 0 Then
                    nIntCompAsignado = nMontoCuotaPrevia - nIntCapitalizadoDiferencia - nIntGraciaAsignado - Calendario(i).SegDes
                Else
                    If (nIntCompCalculado + Abs(nDiferenciaAcumulada) + Calendario(i).SegDes) > nMontoCuotaPrevia Then
                        nIntCompAsignado = nMontoCuotaPrevia - Calendario(i).SegDes - nIntCapitalizadoDiferencia
                    Else
                        nIntCompAsignado = nIntCompCalculado + Abs(nDiferenciaAcumulada)
                    End If
                End If
                    
                If nIntCompAsignado < 0 Then
                    nIntCompAsignado = 0
                End If
                
                nIntCompCalculadoSumatoria = nIntCompCalculado + nIntCompCalculadoSumatoria
                If nIntCompAsignado = 0 Then
                    nIntCompCalculadoAcumulado = (nIntCompCalculadoSumatoria) + nIntCapitalizadoDiferencia
                Else
                    nIntCompCalculadoAcumulado = (nIntCompCalculadoSumatoria)
                End If
                
                If nIntCompAsignado > (nIntCompCalculado + Abs(nDiferenciaAcumulada)) Then
                    nIntCompAsignado = (nIntCompCalculado + Abs(nDiferenciaAcumulada))
                End If
                     
                nIntCompAsignadoAcumulado = nIntCompAsignado + nIntCompAsignadoAcumulado
                nDiferenciaAcumulada = Abs(Round(nIntCompCalculadoAcumulado - nIntCompAsignadoAcumulado, 2))
                 
                'Interés Compensatorio
                If nIntCompAsignado = 0 Then
                    Calendario(i).IntComp = nIntCompAsignado
                Else
                    Calendario(i).IntComp = nIntCompAsignado + nIntCapitalizadoDiferencia
                End If
                
                'Interés de Gracia
                Calendario(i).IntGra = nIntGraciaAsignado
                
                 'Amortizacion Capital
                If Round(nMontoCuotaPrevia - (nIntCompAsignado + nIntGraciaAsignado + Calendario(i).SegDes) - nIntCapitalizadoDiferencia, 4) < 0 Then
                    Calendario(i).Captital = 0
                Else
                    Calendario(i).Captital = Round(nMontoCuotaPrevia - (nIntCompAsignado + nIntGraciaAsignado + Calendario(i).SegDes) - nIntCapitalizadoDiferencia, 4)
                End If
                
                Set oNCOMGasto = Nothing
            End If
            
            '4.1.4.- Ajuste de Saldo Capital en cada cuota
            nSumCapital = nSumCapital + Calendario(i).Captital
            Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
            nSaldoCapital = nSaldoCapital - Calendario(i).Captital
            
            '4.1.5.- Saldo Capital de la Penultima Cuota
            If i = pnCuotaFin - 1 Then
                nSKPenultimaCuota = Calendario(i).SaldoCap
            End If
            
            If pnTipoPeriodo = PeriodoFijo Then
                 dFecTemp = dFecTemp + pnPeriodo
            End If
        Next i
        
        nVASKU = Round((nSaldoCapital / nFVAS), 2)
        nCVASKU = Round((nVASKU * (1 / Round(nSumaFAS, 6))), 2)
        nSKU = Round(nSaldoCapital, 2)
        
        nValorAjusteIteracion = MetodoAjustePorIteracion(nSKU, nFVAS, nSumaFAS, nContadorSKU)
        nValorAjusteIteracionTotal = nValorAjusteIteracionTotal + nValorAjusteIteracion
        
        If nNumIteracion > 17 Then 'Recomendación GELU
            Exit Do
        End If
        nNumIteracion = nNumIteracion + 1
        nContadorSKU = nContadorSKU + 1
        nSaldoCapital = 0
    Loop
        
    '5.- Ajuste Saldo Capital en la ultima Cuota
    If nSumCapital <> pnMonto Then
        Calendario(pnCuotaFin).Captital = Calendario(pnCuotaFin).Captital + Round(Calendario(pnCuotaFin).SaldoCap, 2)
        If pnCuotaFin = 0 Then 'Cuando Calend. es a una cuota
            Calendario(pnCuotaFin).SaldoCap = Calendario(pnCuotaFin).Captital - Calendario(pnCuotaFin).Captital
        Else
            Calendario(pnCuotaFin).SaldoCap = nSKPenultimaCuota - Calendario(pnCuotaFin).Captital
        End If
    End If
    Set oNCOMCredito = Nothing
End Sub
Public Function MetodoAjustePorIteracion(ByRef pnSKU As Double, ByVal pnFVAS As String, _
                            ByVal pnSumaFAS As Double, ByVal pnContadorSKU As Integer) As Double
    Dim nValorAjusteIteracion As Double
    nValorAjusteIteracion = ((pnSKU / pnFVAS) * (1 / pnSumaFAS))
    
    If ((pnSKU >= -1.01 And pnSKU <= 1.01) And pnContadorSKU > 0) Then 'Según Rango Excell - GELU
            pnSKU = 0
    End If
    MetodoAjustePorIteracion = Round(nValorAjusteIteracion, 2)
End Function
'<-***** Fin LUCV20180601

Private Sub CalendarioQuincena(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
                Optional ByVal pMatDesPar As Variant = "", Optional ByVal pnNumMes As Integer = 1, Optional ByVal pbMiViv As Boolean = False)
                
Dim nSaldoCapital As Double
Dim dDesembolso As Date
Dim i As Integer
Dim oCredito As COMNCredito.NCOMCredito
Dim dFecTemp As Date
Dim nMes As Integer
Dim nAnio As Integer
Dim nDia As Integer
Dim nMontoCuotaTemp As Double
Dim nCDIntPend As Double
Dim nPeriodoDesPar As Integer
'Considerando que tiene 2 fechas el 16 y el 1 de cada mes
Dim nDia1 As Integer
Dim nDia2 As Integer
Dim bCambioFechas As Boolean

        nDia1 = 1
        nDia2 = 16
       ' nSaldoCapital = nMonto
        dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
        
        
        Set oCredito = New COMNCredito.NCOMCredito
        nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, nDia1, pnDiasGracia, bProxMes, pnNumMes, pbMiViv)
        nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, nDia2, pnDiasGracia, bProxMes, pnNumMes, pbMiViv)
        Set oCredito = Nothing
        
        If pnTipoPeriodo = FechaFija Then
            Set oCredito = New COMNCredito.NCOMCredito
                nMes = Month(dDesembolso)
                nAnio = Year(dDesembolso)
                nDia = pnDiaFijo
                nCDIntPend = 0
                
                bCambioFechas = False
            
                For i = 0 To pnNroCuotas - 1
                    
                    If bCambioFechas = False Then
                        ' el menor dia
                        bCambioFechas = True
                    Else
                        ' el mayor dia
                        bCambioFechas = False
                    End If
                Next i
        End If
        
End Sub

Private Sub ProcesarCalendario(ByVal pnMonto As Double, _
                ByVal pnTasaInt As Double, _
                ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, _
                ByVal pdFecDesemb As Date, _
                ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, _
                ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, _
                ByVal pnDiaFijo As Integer, _
                ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, _
                Optional ByVal pbCuotaFijaFechaFija As Boolean = False, _
                Optional ByVal pbDesemParcial As Boolean = False, _
                Optional ByVal pMatDesPar As Variant = "", _
                Optional ByVal pnNumMes As Integer = 1, _
                Optional ByVal pbMiViv As Boolean = False, Optional ByVal bQuincena As Boolean, _
                Optional ByVal pnCuotaIni As Integer = 0, Optional ByVal pnCuotaFin As Integer = 0, _
                Optional ByVal pnDiaFijo2 As Integer = 0, Optional ByVal pnMontoCapInicial As Double = 0, _
                Optional ByVal pbRenovarCredito As Boolean = False, Optional ByVal pnInteresAFecha As Double = 0, _
                Optional ByVal pnMontoGra As Double = 0, Optional ByVal pnCuotaBalon As Integer = 0, _
                Optional ByVal psCtaCod As String = "", Optional ByVal pnTasaSegDes As Double = 0, _
                Optional ByRef pMatCalendSegDes As Variant, Optional ByVal pbEsSimulador As Boolean = False, _
                Optional ByVal pnExoSeguroDesgravamen As Integer = 0, Optional ByVal pnMontoPoliza As Double, _
                Optional ByVal pnTasaSegInc As Double)
                'MAVM 20130305: pnMontoGra
                'WIOR 20131111 AGREGO pnCuotaBalon
                'LUCV20180601, psCtaCod, pnTasaSegDes, pMatCalendSegDes, pbEsSimulador, pnExoSeguroDesgravamen, pnMontoPoliza,pnTasaSegInc. Según ERS022-2018
                If bQuincena = False Then
                    Select Case pnTipoCuota
                        Case Creciente
                            Call CalendarioCreciente(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pnCuotaIni, pnCuotaFin, pnDiaFijo2, pnMontoCapInicial)
                        Case Decreciente
                            Call CalendarioDecreciente(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pnCuotaIni, pnCuotaFin, pnDiaFijo2, pnMontoCapInicial)
                        Case Fija
                        '->***** LUCV20180601, Comentó y agregó. según ERS022-2018
                            If pbEsSimulador Then 'Según Adecuaciones ARLO
                                    Call CalendarioCuotaFija(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, _
                                                pnDiaFijo, bProxMes, MatGracia, pbCuotaFijaFechaFija, pbDesemParcial, pMatDesPar, pnNumMes, pbMiViv, pnCuotaIni, _
                                                pnCuotaFin, pnDiaFijo2, pnMontoCapInicial, pbRenovarCredito, pnInteresAFecha, pnMontoGra, pnCuotaBalon)
                                                'MAVM 20130209: pnMontoGra
                                                'WIOR 20131111 AGREGO pnCuotaBalon
                            Else
                                    Call CalendarioCuotaFijaNuevo(pnMonto, _
                                                                pnTasaInt, _
                                                                pnNroCuotas, _
                                                                pnPeriodo, _
                                                                pdFecDesemb, _
                                                                pnTipoPeriodo, _
                                                                pnTipoGracia, _
                                                                pnDiasGracia, _
                                                                pnDiaFijo, _
                                                                pnNumMes, _
                                                                pnCuotaIni, _
                                                                pnCuotaFin, _
                                                                psCtaCod, _
                                                                pnTasaSegDes, _
                                                                pMatCalendSegDes, _
                                                                pnExoSeguroDesgravamen, _
                                                                pnMontoPoliza, _
                                                                pnTasaSegInc)
                                                                'LUCV20180601, Agregó: psCtaCod, pnTasaSegDes, pMatCalendSegDes,pnExoSeguroDesgravamen,pnMontoPoliza,pnTasaMensualSegInc  según ERS022-2018
                            End If
                        '<-***** Fin LUCV20180601
                    End Select
                Else
                    If bQuincena = True Then
                        Call CalendarioTrabajadoresDirectores(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, bProxMes, pnDiasGracia, MatGracia, pnTipoGracia, pnCuotaIni, pnCuotaFin)
                    End If
                End If
End Sub
Public Sub Inicio(ByVal psNomCmac As String, ByVal psNomAgencia As String, _
            ByVal psCodUser As String, ByVal psFechaSis As String)
    csNomCMAC = psNomCmac
    csNomAgencia = psNomAgencia
    csCodUser = psCodUser
    csFechaSis = psFechaSis

End Sub

'Public Function GeneraCalendario(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
'                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
'                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
'                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
'                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, _
'                Optional ByVal pbCuotaComodin As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
'                Optional ByVal pMatDesPar As Variant = "", Optional ByVal pnNumMes As Integer = 1, Optional ByVal pbMiViv As Boolean = False, _
'                Optional ByVal bQuincena As Boolean, Optional ByVal pbGraciaEnCuotas As Boolean = False, Optional ByVal pnTasaGracia As Double = 0, _
'                Optional ByVal pnDiaFijo2 As Integer = 0, Optional pnMontoCapInicial As Double = 0, _
'                Optional ByVal pbPagoInteresEnGracia As Boolean = False, _
'                Optional ByVal pdFechaInicioPago As Date = 0) As Variant
'
''Si pnMontoCapInicial >0 entonces se Capitalizo la Gracia sin aumentar el capital
'
'Dim sCalendPagos() As String
'Dim i As Integer
'Dim TasaComTemp As Double
'Dim nPlazoCom As Integer
'Dim nCuotasCom As Integer
'Dim nCuotaMontoCom As Double
'Dim oCred As COMNCredito.NCOMCredito
'
''Gracia en Cuotas
'Dim nCuotasGracia As Integer
'Dim nCuotaIni As Integer
'Dim nCuotaFin As Integer
'Dim dDesembolso As Date
'
'        On Error GoTo ErrorGeneraCalendario
'
'
'        If pbCuotaComodin Then
'            TasaComTemp = pnTasaInt
'            If pnTipoPeriodo = PeriodoFijo Then
'                nPlazoCom = pnPeriodo
'            Else
'                nPlazoCom = 30
'            End If
'            nCuotasCom = pnNroCuotas
'            Set oCred = New COMNCredito.NCOMCredito
'            nCuotaMontoCom = oCred.CuotaFija(TasaComTemp, nCuotasCom + 1, pnMonto, nPlazoCom)
'            nCuotaMontoCom = CDbl(Format((nCuotaMontoCom * (pnNroCuotas + 1)) / pnNroCuotas, "#0.00"))
'            TasaComTemp = CDbl(Format(oCred.CreditosTasaEfectiva(nCuotaMontoCom, pnNroCuotas, pnMonto, pnTasaInt, nPlazoCom), "#0.0000"))
'            pnTasaInt = Format(TasaComTemp * 100, "#0.0000")
'            Set oCred = Nothing
'        End If
'
'
'    '******************************************
'    '******** Gracia repartida en Cuotas ******
'
'    nCuotasGracia = 0
'    nCuotaIni = 0
'    nCuotaFin = pnNroCuotas - 1
'    ReDim Calendario(pnNroCuotas)
'
'    If pbGraciaEnCuotas And pnTipoPeriodo = PeriodoFijo Then
''        nCuotasGracia = Fix(pnDiasGracia / pnPeriodo)   'Entero inmediato inferior
''        dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy"))
''        ReDim Calendario(pnNroCuotas + nCuotasGracia)
''        Set oCred = New COMNCredito.NCOMCredito
''        For i = 0 To nCuotasGracia - 1
''            Calendario(i).dFecha = IIf(i = nCuotasGracia - 1, CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia, dDesembolso + pnPeriodo)
''            Calendario(i).NroCuota = i + 1
''            'Agregado para interes Compensatorio
''            If pbPagoInteresEnGracia Then
''                If i = nCuotasGracia - 1 And nCuotasGracia > 1 Then
''                    Calendario(i).IntComp = oCred.MontoIntPerDias(pnTasaInt, Calendario(i).dFecha - (Calendario(i - 1).dFecha), pnMonto)
''                Else
''                    Calendario(i).IntComp = oCred.MontoIntPerDias(pnTasaInt, pnPeriodo, pnMonto)
''                End If
''            Else
''                Calendario(i).IntComp = 0#
''            End If
''            '****************************************
''            Calendario(i).Gasto = 0#
''            If i = nCuotasGracia - 1 And nCuotasGracia > 1 Then
''                Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, Calendario(i).dFecha - (Calendario(i - 1).dFecha), pnMonto)
''            Else
''                If nCuotasGracia = 1 Then
''                    Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, Calendario(i).dFecha - pdFecDesemb, pnMonto)
''                Else
''                    Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, pnPeriodo, pnMonto)
''                End If
''            End If
''
''            Calendario(i).Captital = 0#
''            'Agregado para interes de Gracia
''            'Calendario(i).Cuota = Calendario(i).IntGra
''            Calendario(i).Cuota = Calendario(i).IntGra + Calendario(i).IntComp
''            '******************************************
''            Calendario(i).SaldoCap = pnMonto
''            dDesembolso = dDesembolso + pnPeriodo
''        Next i
''        nCuotaIni = nCuotasGracia
''        nCuotaFin = pnNroCuotas + nCuotasGracia - 1
''        ReDim Preserve Calendario(nCuotaFin + 1)
''        Set oCred = Nothing
''    End If
'        nCuotasGracia = pnDiasGracia
'
'
'        dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy"))
'
'        ReDim Calendario(pnNroCuotas + nCuotasGracia)
'        Set oCred = New COMNCredito.NCOMCredito
'        For i = 0 To nCuotasGracia - 1
'            'Calendario(i).dFecha = IIf(i = nCuotasGracia - 1, CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia, dDesembolso + pnPeriodo)
'            Calendario(i).dFecha = dDesembolso + pnPeriodo
'            Calendario(i).NroCuota = i + 1
'            'Agregado para interes Compensatorio
'            If pbPagoInteresEnGracia Then
'                If i = nCuotasGracia - 1 And nCuotasGracia > 1 Then
'                    Calendario(i).IntComp = oCred.MontoIntPerDias(pnTasaInt, Calendario(i).dFecha - (Calendario(i - 1).dFecha), pnMonto)
'                Else
'                    Calendario(i).IntComp = oCred.MontoIntPerDias(pnTasaInt, pnPeriodo, pnMonto)
'                End If
'            Else
'                Calendario(i).IntComp = 0#
'            End If
'            '****************************************
'            Calendario(i).Gasto = 0#
'            If i = nCuotasGracia - 1 And nCuotasGracia > 1 Then
'                Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, Calendario(i).dFecha - (Calendario(i - 1).dFecha), pnMonto)
'            Else
'                If nCuotasGracia = 1 Then
'                    Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, Calendario(i).dFecha - pdFecDesemb, pnMonto)
'                Else
'                    Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, pnPeriodo, pnMonto)
'                End If
'            End If
'
'            Calendario(i).Captital = 0#
'            'Agregado para interes de Gracia
'            'Calendario(i).Cuota = Calendario(i).IntGra
'            Calendario(i).Cuota = Calendario(i).IntGra + Calendario(i).IntComp
'            '******************************************
'            Calendario(i).SaldoCap = pnMonto
'            dDesembolso = dDesembolso + pnPeriodo
'        Next i
'        nCuotaIni = nCuotasGracia
'        nCuotaFin = pnNroCuotas + nCuotasGracia - 1
'        ReDim Preserve Calendario(nCuotaFin + 1)
'        Set oCred = Nothing
'    End If
'
'    '************************************
'
''***10-05-2006 *******************************************
'    If pbGraciaEnCuotas And pnTipoPeriodo = FechaFija Then
'        Dim nMes As Integer
'        Dim nAnio As Integer
'        Dim nDia As Integer
'        Dim dFecTemp As Date
'        Dim nDifDiasGracia As Integer
'
'        'nCuotasGracia = Fix(pnDiasGracia / 30)   'Entero inmediato inferior
'        'nDifDiasGracia = pnDiasGracia Mod 30    'Dias faltantes
'
'        nCuotasGracia = pnDiasGracia
'
'        'If nDifDiasGracia = 0 Then
'        '    dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy"))
'            'nMes = Month(dDesembolso)
'       ' Else
'            'If bProxMes Then
'       '         dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + CDate(nDifDiasGracia)
'                'nMes = Month(dDesembolso)
'            '    nMes = Month(dDesembolso) + 1
'            'End If
'        'End If
'
'        If pdFechaInicioPago <> 0 Then
'
'            dDesembolso = pdFechaInicioPago
'        Else
'            dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + CDate(pnDiasGracia)
'        End If
'
'        nMes = Month(dDesembolso)
'
'        ReDim Calendario(pnNroCuotas + nCuotasGracia)
'
'        nAnio = Year(dDesembolso)
'        nDia = pnDiaFijo
'
'        Set oCred = New COMNCredito.NCOMCredito
'        For i = 0 To nCuotasGracia - 1
'            'Calendario(i).dFecha = IIf(i = nCuotasGracia - 1, CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia, dDesembolso + pnPeriodo)
'            If (Not (i = 0 And nDia > Day(dDesembolso) And (Not bProxMes)) Or pnNumMes = 6) And (pdFechaInicioPago <> 0 And i > 0) Then
'                If nDia = pnDiaFijo Then nMes = nMes + pnNumMes
'                If nMes > 12 Then
'                    nAnio = nAnio + 1
'                    nMes = 1
'                End If
'            Else
'                If nDia > 30 Then
'                    If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
'                        nMes = nMes + pnNumMes
'                    End If
'                End If
'            End If
'            If nMes = 2 Then
'                If nDia > 28 Then
'                    If nAnio Mod 4 = 0 Then
'                        nDia = 29
'                    Else
'                        nDia = 28
'                    End If
'                End If
'            Else
'                If nDia > 30 Then
'                    If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
'                        nDia = 30
'                    End If
'                End If
'            End If
'
'            Calendario(i).NroCuota = i + 1
'
'            If pdFechaInicioPago <> 0 And i = 0 Then
'
'                    dFecTemp = pdFechaInicioPago
'
'            Else
'
'                dFecTemp = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)))
'
'            End If
'            Calendario(i).dFecha = dFecTemp
'            'Agregado para interes Compensatorio
'            'If pbPagoInteresEnGracia Then
'            '    If i = nCuotasGracia - 1 Then
'            '        Calendario(i).IntComp = oCred.MontoIntPerDias(pnTasaInt, Calendario(i).dFecha - (Calendario(i - 1).dFecha), pnMonto)
'            '    Else
'            '        Calendario(i).IntComp = oCred.MontoIntPerDias(pnTasaInt, pnPeriodo, pnMonto)
'            '    End If
'            'Else
'            '    Calendario(i).IntComp = 0#
'            'End If
'            '****************************************
'            Calendario(i).Gasto = 0#
'
'            If i = 0 Then
'                Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, DateDiff("d", pdFecDesemb, Calendario(i).dFecha), pnMonto)
'            Else
'            '   If i = nCuotasGracia - 1 And nCuotasGracia > 1 Then
'
'                    'Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, (Calendario(i).dFecha + CDate(nDifDiasGracia)) - (Calendario(i - 1).dFecha), pnMonto)
'            '        Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, (pdFecDesemb + CDate(pnDiasGracia)) - (Calendario(i - 1).dFecha), pnMonto)
'
'            '    Else
'
'                    Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, DateDiff("d", CDate(Calendario(i - 1).dFecha), Calendario(i).dFecha), pnMonto)
'
'           '     End If
'            End If
'
'            Calendario(i).Captital = 0#
'            'Agregado para interes de Gracia
'            Calendario(i).Cuota = Calendario(i).IntGra + Calendario(i).IntComp
'            '******************************************
'            Calendario(i).SaldoCap = pnMonto
'            'dDesembolso = dDesembolso + pnPeriodo
'        Next i
'        nCuotaIni = nCuotasGracia
'        nCuotaFin = pnNroCuotas + nCuotasGracia - 1
'        ReDim Preserve Calendario(nCuotaFin + 1)
'        Set oCred = Nothing
'
''        If bProxMes Then
'            'pdFecDesemb = DateAdd("m", nCuotasGracia, pdFecDesemb)
'
'        'If bProxMes Then
'        '    pdFecDesemb = DateAdd("m", 1, pdFecDesemb)
'        'End If
''            If nDifDiasGracia > 0 Then 'And bProxMes
''                pdFecDesemb = pdFecDesemb + CDate(nDifDiasGracia)
'
'        '    pdFecDesemb = DateAdd("m", 1, pdFecDesemb)
'        '    pdFecDesemb = DateAdd("d", -nDifDiasGracia, pdFecDesemb)
'
''            End If
''            pdFecDesemb = DateAdd("m", nCuotasGracia, pdFecDesemb)
''            pdFecDesemb = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(Month(pdFecDesemb))), 2) & "/" & Trim(Str(Year(pdFecDesemb))))
''        End If
'        'If bProxMes Then
'        '    pdFecDesemb = DateAdd("m", 1, pdFecDesemb)
'        'End If
'        'pdFecDesemb = pdFecDesemb - pnDiasGracia
'        'pdFecDesemb = DateAdd("m", -nCuotasGracia, pdFecDesemb)
'End If
''**************************************************
'
'        'ReDim Calendario(pnNroCuotas )
'        Call ProcesarCalendario(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoCuota, _
'                pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pbCuotaFijaFechaFija, _
'                pbDesemParcial, pMatDesPar, pnNumMes, pbMiViv, bQuincena, nCuotaIni, nCuotaFin, pnDiaFijo2, pnMontoCapInicial)
'
'        ReDim sCalendPagos(UBound(Calendario), 8)
'
'        For i = 0 To UBound(Calendario) - 1
'            sCalendPagos(i, 0) = Format(Calendario(i).dFecha, "dd/mm/yyyy")
'            sCalendPagos(i, 1) = Trim(Str(Calendario(i).NroCuota))
'            sCalendPagos(i, 2) = Format(Calendario(i).Cuota, "#0.00")
'            sCalendPagos(i, 3) = Format(Calendario(i).Captital, "#0.00")
'            sCalendPagos(i, 4) = Format(Calendario(i).IntComp, "#0.00")
'            sCalendPagos(i, 5) = Format(Calendario(i).IntGra, "#0.00")
'            sCalendPagos(i, 6) = Format(Calendario(i).Gasto, "#0.00")
'            sCalendPagos(i, 7) = Format(Calendario(i).SaldoCap, "#0.00")
'        Next i
'        GeneraCalendario = sCalendPagos
'        Exit Function
'
'ErrorGeneraCalendario:
'        'MsgBox Err.Description, vbCritical, "Aviso"
'        Err.Raise Err.Number, "Error", Err.Description
'End Function

'29-05-2006
Public Function ReporteCalendario(ByVal pTipoRep As Byte, ByVal MatP As Variant, ByVal MatMalo As Variant, _
        ByVal pTCuota As String, ByVal pInteres As Double, ByVal pMonto, _
        ByVal pCuota As Double, ByVal pPlazo As Integer, ByVal pVigencia As Date, _
        Optional ByVal pnTipoProceso As Integer = 0, Optional ByVal pMatDesPar As Variant = "", Optional GBITFAPLICAt As Boolean, Optional gnITFPorcentt As Double, Optional gnITFMontoMint As Double = 0, _
        Optional ByVal psCtaCod As String, Optional ByVal pDiasGracia As Integer = 0, Optional ByVal psTipoPeriodo As String = "") As String

Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer

Dim cTitulo As String
Dim i As Integer
Dim cCuota As String * 35
Dim sCapital As Double
Dim sInteres As Double
 
Dim ssql As String
Dim oConec As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim sNombre As String
 
'29-05-2006
Dim sInteresGracia As Double
Dim sGastos As Double
Dim sCuotaConIGV As Double
'-----------
oITF.gbITFAplica = GBITFAPLICAt
oITF.gnITFPorcent = gnITFPorcentt
oITF.gnITFMontoMin = gnITFMontoMint
 
cCuota = pTCuota
sCapital = 0
sInteres = 0
  
    If pTipoRep = 1 Then
        If pnTipoProceso = 0 Then
            cTitulo = "SIMULACION DE CALENDARIO DE PAGOS"
        ElseIf pnTipoProceso = 1 Then
            cTitulo = "SUGERENCIA DE ANALISTA - CALENDARIO DE PAGOS"
        ElseIf pnTipoProceso = 2 Then
            cTitulo = "APROBACION DE CREDITO - CALENDARIO DE PAGOS"
        End If
    ElseIf pTipoRep = 2 Then
        If pnTipoProceso = 0 Then
            cTitulo = "SIMULACION DE CALENDARIO DE PAGOS MI VIVIENDA"
        ElseIf pnTipoProceso = 1 Then
            cTitulo = "SUGERENCIA DE ANALISTA - CALENDARIO DE PAGOS MI VIVIENDA"
        ElseIf pnTipoProceso = 2 Then
            cTitulo = "APROBACION DE CREDITO - CALENDARIO DE PAGOS MI VIVIENDA"
        End If
    End If

    lnLineas = 0
    lnPage = 1

    'obteniendo titular
    ssql = "Select Pers.cPersNombre"
    ssql = ssql & " From ProductoPersona PP"
    ssql = ssql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    ssql = ssql & " Where PP.cCtaCod='" & psCtaCod & "' And PP.nPrdPersRelac=20"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(ssql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        sNombre = rs!cPersNombre
    End If
    Set rs = Nothing
    
    'CABECERA
    If pTipoRep = 1 Then
        lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "PLAN DE PAGOS", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, psCtaCod, Trim(sNombre), pDiasGracia, psTipoPeriodo)
    ElseIf pTipoRep = 2 Then
        lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "BUEN PAGADOR", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, psCtaCod, Trim(sNombre))
    End If
    
    
    If IsArray(pMatDesPar) Then
        lsCadImp = lsCadImp & Space(5) & "DESEMBOLSOS PARCIALES" & Chr(10)
        lsCadImp = lsCadImp & Space(5) & String(63, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(5) & oImpre.ImpreFormat("FECHA", 16, 0) & oImpre.ImpreFormat("MONTO", 10, 0) & Chr(10)
        lsCadImp = lsCadImp & Space(5) & String(63, "-") & Chr(10)
        For i = 0 To UBound(pMatDesPar) - 1
            lsCadImp = lsCadImp & Space(5) & oImpre.ImpreFormat(pMatDesPar(i, 0), 10, 0) & oImpre.ImpreFormat(CDbl(pMatDesPar(i, 1)), 10, 2, True) & Chr(10)
        Next i
        lsCadImp = lsCadImp & Space(5) & String(63, "-") & Chr(10)
    End If
    
    lsCadImp = lsCadImp & String(123, "-") & Chr(10)
    lsCadImp = lsCadImp & "ITEM  FECHA            NO.CUOTA        CUOTA      CAPITAL      INTERES   INT.GRACIA       GASTOS   SALDO CAP.   CUOTA + ITF" & Chr(10)
    lsCadImp = lsCadImp & String(123, "-") & Chr(10)
    
    'lsCadImp = lsCadImp & Chr(10)
    
    lnIndice = 10:  lnLineas = 10
    nItem = 0
    For i = 0 To UBound(MatP) - 1
    
        nItem = nItem + 1
        lnLineas = lnLineas + 1
        lnIndice = lnIndice + 1
        lsCadImp = lsCadImp & oImpre.ImpreFormat(nItem, 4, 0) & Space(2)
        lsCadImp = lsCadImp & oImpre.ImpreFormat(Format(MatP(i, 0), "ddd, dd mmm yyyy"), 16, 0) & Space(2)
        lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatP(i, 1)), 7, 0) & Space(2)
        lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatP(i, 2)), 8, 2) & Space(2)
        lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatP(i, 3)), 8, 2) & Space(2)
        sCapital = sCapital + MatP(i, 3)
        lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatP(i, 4)), 8, 2) & Space(2)
        sInteres = sInteres + MatP(i, 4)
        lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatP(i, 5)), 8, 2) & Space(2)
        
        sInteresGracia = sInteresGracia + MatP(i, 5)
        
        lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatP(i, 6)), 8, 2) & Space(2)
        
        sGastos = sGastos + MatP(i, 6)
        
        lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatP(i, 7)), 8, 2) & Space(2)
        lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatP(i, 2)) + oITF.fgITFCalculaImpuesto(val(MatP(i, 2))), 8, 2) & Space(2) & Chr(10)
        'sCuotaConIGV = sCuotaConIGV + Val(MatP(i, 2)) + oITF.fgITFCalculaImpuesto(Val(MatP(i, 2)))
 
        If lnIndice Mod 300 = 0 Then
            lsCadBuffer = lsCadBuffer & lsCadImp
            lsCadImp = ""
        End If
        
        If lnLineas >= 55 Then
            lnPage = lnPage + 1
            lsCadImp = lsCadImp & Chr(12)
            lsCadImp = lsCadImp & Chr(10)
            If pTipoRep = 1 Then
                lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "PLAN DE PAGOS", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia)
            ElseIf pTipoRep = 2 Then
                lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "BUEN PAGADOR", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia)
            End If
        
            lnLineas = 8
            lnIndice = lnIndice + 8
        End If
                
    Next

    lsCadImp = lsCadImp & Chr(10) & String(123, "=") & Chr(10)
    lsCadImp = lsCadImp & "Totales :" & Space(37) & oImpre.ImpreFormat(sCapital, 8, 2)
    lsCadImp = lsCadImp & Space(2) & oImpre.ImpreFormat(sInteres, 8, 2)
    '29-05
    lsCadImp = lsCadImp & Space(2) & oImpre.ImpreFormat(sInteresGracia, 8, 2)
    lsCadImp = lsCadImp & Space(2) & oImpre.ImpreFormat(sGastos, 8, 2)
    '--
    lsCadImp = lsCadImp & Space(2) & oImpre.ImpreFormat(sCapital + sInteres, 8, 2)
    
    '29-05
    'lsCadImp = lsCadImp & Space(2) & oImpre.ImpreFormat(sCuotaConIGV, 8, 2)
    'Mal Pagador Mi Vivienda
    
    sCapital = 0
    sInteres = 0
     
    If pTipoRep = 2 Then
        lsCadImp = lsCadImp & Chr(12)
        lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "MAL PAGADOR", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia)
     
        lsCadImp = lsCadImp & Chr(10)
        
        lnIndice = 10:  lnLineas = 10
        nItem = 0
        For i = 0 To UBound(MatMalo) - 1
            nItem = nItem + 1
            lnLineas = lnLineas + 1
            lnIndice = lnIndice + 1
            lsCadImp = lsCadImp & oImpre.ImpreFormat(nItem, 4, 0) & Space(2)
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Format(MatMalo(i, 0), "ddd, dd mmm yyyy"), 16, 0) & Space(2)
            lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatMalo(i, 1)), 7, 0) & Space(2)
            lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatMalo(i, 2)), 8, 2) & Space(2)
            lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatMalo(i, 3)), 8, 2) & Space(2)
            sCapital = sCapital + MatMalo(i, 3)
            lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatMalo(i, 4)), 8, 2) & Space(2)
            sInteres = sInteres + MatMalo(i, 4)
            lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatMalo(i, 5)), 8, 2) & Space(2)
        
            sInteresGracia = sInteresGracia + MatP(i, 5)
        
            lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatMalo(i, 6)), 8, 2) & Space(2)
        
            sGastos = sGastos + MatP(i, 6)
            lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatMalo(i, 7)), 8, 2) & Space(2)
            lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatMalo(i, 8)), 8, 2) & Space(2) & Chr(10)
              
            '29-05-2006
            lsCadImp = lsCadImp & oImpre.ImpreFormat(val(MatP(i, 2)) + oITF.fgITFCalculaImpuesto(val(MatP(i, 2))), 8, 2) & Space(2) & Chr(10)
            'sCuotaConIGV = sCuotaConIGV + Val(MatP(i, 2)) + oITF.fgITFCalculaImpuesto(Val(MatP(i, 2)))
        
            'lnLineas = lnLineas + 1
            If lnIndice Mod 300 = 0 Then
                lsCadBuffer = lsCadBuffer & lsCadImp
                lsCadImp = ""
            End If
            
            If lnLineas >= 55 Then
                lnPage = lnPage + 1
                lsCadImp = lsCadImp & Chr(12)
                lsCadImp = lsCadImp & Chr(10)
                lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "MAL PAGADOR", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia)
             
                lnLineas = 8
                lnIndice = lnIndice + 8
            End If
                    
        Next
        
        lsCadImp = lsCadImp & Chr(10) & String(109, "=") & Chr(10)
        lsCadImp = lsCadImp & "Totales :" & Space(37) & oImpre.ImpreFormat(sCapital, 8, 2)
        lsCadImp = lsCadImp & Space(2) & oImpre.ImpreFormat(sInteres, 8, 2)
        
        lsCadImp = lsCadImp & Space(2) & oImpre.ImpreFormat(sInteresGracia, 8, 2)
        lsCadImp = lsCadImp & Space(2) & oImpre.ImpreFormat(sGastos, 8, 2)
        
        lsCadImp = lsCadImp & Space(2) & oImpre.ImpreFormat(sCapital + sInteres, 8, 2)
        '29-05
        'lsCadImp = lsCadImp & Space(2) & oImpre.ImpreFormat(sCuotaConIGV, 8, 2)

    End If
ReporteCalendario = lsCadBuffer & lsCadImp
 
 End Function

'Modificado 29-05
Public Function nRepoCabecera(ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        ByVal pnCodReporte As String, ByVal pTCuota As String, ByVal pInteres As Double, ByVal pMonto As Double, _
        ByVal pCuota As Double, ByVal pPlazo As Integer, ByVal pVigencia As Date, _
        Optional ByVal psCtaCod As String, Optional ByVal psNomCli As String, _
        Optional ByVal pCuotasGracia As Integer = -1, Optional ByVal psTipoPeriodo As String = "") As String
        
Dim lsCadImp As String
'Dim loImprimeCab As NColPImpre

'    Set loImprimeCab = New NColPImpre
'        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
        lsCadImp = nImprimeCabeceraReportes(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
'    Set loImprimeCab = Nothing
    lsCadImp = lsCadImp & Chr(10) & String(pnAnchoLinea, "-") & Chr(10)
    Select Case pnCodReporte
        Case "CalenPagos"
            lsCadImp = lsCadImp & "    CREDITO         : " & psCtaCod & Chr(10)
            lsCadImp = lsCadImp & "    CLIENTE         : " & psNomCli & Chr(10)
            lsCadImp = lsCadImp & "    TIPO DE CUOTA   : " & pTCuota & Space(20) & "CUOTA :      " & oImpre.ImpreFormat(pMonto, 10, 2) & Chr(10)
            lsCadImp = lsCadImp & "    INTERES         : " & oImpre.ImpreFormat(pInteres, 8, 4) & Space(42) & "PLAZO :      " & oImpre.ImpreFormat(pPlazo, 10, 2) & Chr(10)
            lsCadImp = lsCadImp & "    MONTO           :   " & oImpre.ImpreFormat(pMonto, 8, 2) & Space(42) & "VIGENCIA: " & Format(pVigencia, "ddd, dd mmm yyyy") & Chr(10)
            
            If pCuotasGracia <> -1 Then
                lsCadImp = lsCadImp & "    PERIODO GRACIA  :   " & oImpre.ImpreFormat(pCuotasGracia, 8, 0) & " " & psTipoPeriodo & Chr(10)
            End If
            lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & Chr(10)
            
       End Select
       
nRepoCabecera = lsCadImp
End Function

Public Function nImprimeCabeceraReportes(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, _
        ByVal psFechaSis As String, ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        Optional ByVal psCodRepo As String) As String
        
  Dim lsCabe01 As String, lsCabe02 As String
  Dim lsCabe03 As String, lsCabe04 As String
  Dim lsCabRepo As String
  
  lsCabRepo = ""
  ' Cabecera 1
  lsCabe01 = oImpre.FillText(Trim(UCase(psNomCmac)), 55, " ")
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 55 - 25)
  lsCabe01 = lsCabe01 & "Pag.  : " & str(pnPagina) & "  -  " & psCodUser & Chr(10)
  'lsCabe01 = lsCabe01 & IIf(pbCiereDia = True, IIf(VerifSiCierreDia(), "DC", "AC"), "") & chr$(10)
  ' Cabecera 2
  lsCabe01 = lsCabe01 & oImpre.FillText(Trim(UCase(psNomAgencia)), 35, " ")
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 35 - 25)
  lsCabe01 = lsCabe01 & "Fecha : " & Format(psFechaSis & " " & Time, "dd/mm/yyyy hh:mm") & Chr$(10)
  ' Titulo
  'psTitulo = psTitulo
  lsCabe02 = psCodRepo & String(Int((pnAnchoLinea - Len(psTitulo)) / 2), " ") & psTitulo & Chr$(10)
  ' SubTitulo
  lsCabe03 = String(Int((pnAnchoLinea - Len(psSubTitulo)) / 2), " ") & psSubTitulo & Chr$(10)
  ' Comenta
  lsCabe04 = IIf(Len(psComenta) > 0, psComenta & Chr$(10), "")
  ' ***
  lsCabRepo = lsCabRepo & lsCabe01 & lsCabe02
  lsCabRepo = lsCabRepo & lsCabe03 & lsCabe04
  nImprimeCabeceraReportes = lsCabRepo
End Function


Private Sub CalendarioTrabajadoresDirectores(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, ByVal bProxMes As Boolean, Optional ByVal pnDiasGracia As Integer, _
                Optional ByVal MatGracia As Variant, Optional ByVal pnTipoGracia As TCalendTipoGracia, _
                Optional ByVal pnCuotaIni As Integer = 0, Optional ByVal pnCuotaFin As Integer = 0)

' El Calendario de Trabajadores y Directores tiene lo siguiente 1 y los 16 fechas de Pago


Dim dDesembolso As Date

Dim nMontoCuotaTemp1 As Double
Dim nMontoCuotaTemp16 As Double
Dim nSaldoCapital As Double
Dim nSaldoCapitalTmp As Double

Dim nDiaDesemb As Integer
Dim nCuotas1 As Integer
Dim nCuotas16 As Integer
Dim nMes As Integer
Dim nDia As Integer
Dim nAno As Integer
Dim i As Integer

Dim MatFechas() As Variant

Dim nMontTemp As Double
Dim oCred As COMNCredito.NCOMCredito

        
        nSaldoCapital = pnMonto
        dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
        ' considerando dos Fechas Fijas los 1 y 16 de Cada Mes
        
        ' Se verifica quien empieza
        nDiaDesemb = Day(dDesembolso)
        nAno = Year(dDesembolso)
        If nDiaDesemb > 0 And nDiaDesemb <= 8 Then
           nDia = 16
        Else
           nDia = 1
        End If
        
        nMes = Month(dDesembolso)
        
        ' se verifica cuando empieza el calendario el mismo mes o el mes siguiente
        If nDiaDesemb > nDia Then
            If nMes = 12 Then
                nAno = nAno + 1
                nMes = nMes + 1
            Else
                nMes = nMes + 1
            End If
        End If
        
        'Armando Matriz de Fechas
        ReDim MatFechas(pnNroCuotas)
        
        For i = 0 To pnNroCuotas - 1
            MatFechas(i) = CDate(nDia & "/" & nMes & "/" & nAno)
            
            If nDia = 1 Then
                nDia = 16
            Else
                nDia = 1
                nMes = nMes + 1
            End If
            
            If nMes > 12 Then
                nMes = 1
                nAno = nAno + 1
            End If
        Next i
        
        If nDiaDesemb > 0 And nDiaDesemb <= 8 Then
           nDia = 16
        Else
           nDia = 1
        End If
        
        Set oCred = New COMNCredito.NCOMCredito
        nMontTemp = oCred.CFijaFechaFijaTrabajadores(pnTasaInt, pnNroCuotas, pnMonto)
        nSaldoCapitalTmp = nSaldoCapital
        For i = 0 To pnNroCuotas - 1
            
            Calendario(i).dFecha = MatFechas(i)
            Calendario(i).NroCuota = i + 1
            Calendario(i).IntGra = 0#
            Calendario(i).Gasto = 0
            
            Set oCred = New COMNCredito.NCOMCredito
            If i = 0 Then
                Calendario(i).IntComp = oCred.MontoIntPersDiasQuincena(pnTasaInt, DateDiff("d", dDesembolso, Calendario(i).dFecha), nSaldoCapital)
                Calendario(i).Cuota = nMontTemp
                Calendario(i).Captital = nMontTemp - Calendario(i).IntComp
                Calendario(i).SaldoCap = nSaldoCapitalTmp - Calendario(i).Captital
                nSaldoCapitalTmp = nSaldoCapitalTmp - Calendario(i).Captital
            Else
                If i = pnNroCuotas - 1 Then
                    Calendario(i).IntComp = oCred.MontoIntPersDiasQuincena(pnTasaInt, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha), nSaldoCapitalTmp)
                    Calendario(i).Captital = nSaldoCapitalTmp
                    Calendario(i).Cuota = nMontTemp + Calendario(i).IntComp
                    Calendario(i).SaldoCap = nSaldoCapitalTmp - Calendario(i).Captital
                    nSaldoCapitalTmp = nSaldoCapitalTmp - Calendario(i).Captital
                Else
                    Calendario(i).IntComp = oCred.MontoIntPersDiasQuincena(pnTasaInt, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha), nSaldoCapitalTmp)
                    Calendario(i).Cuota = nMontTemp
                    Calendario(i).Captital = nMontTemp - Calendario(i).IntComp
                    Calendario(i).SaldoCap = nSaldoCapitalTmp - Calendario(i).Captital
                    nSaldoCapitalTmp = nSaldoCapitalTmp - Calendario(i).Captital
                End If
            End If
       Next i
       
       
       If pnDiasGracia > 0 Then
        If pnTipoGracia = PrimeraCuota Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            For i = pnNroCuotas To 1 Step -1
                Calendario(i) = Calendario(i - 1)
                Calendario(i).NroCuota = Calendario(i).NroCuota + 1
            Next i
            Calendario(0).dFecha = pdFecDesemb + pnDiasGracia
            Calendario(0).NroCuota = 1
            Calendario(0).IntGra = MatGracia(0)
            Calendario(0).Gasto = 0
            Calendario(0).IntComp = 0#
            Calendario(0).Cuota = Calendario(0).IntGra
            Calendario(0).Captital = 0#
            Calendario(0).SaldoCap = pnMonto
        End If
        If pnTipoGracia = Prorateada Or pnTipoGracia = Configurable Then
            For i = 0 To pnNroCuotas - 1
                Calendario(i).IntGra = MatGracia(i)
                Calendario(i).Cuota = Calendario(i).Cuota + MatGracia(i)
            Next i
        End If
        If pnTipoGracia = UltimaCuota Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            Calendario(pnNroCuotas).dFecha = Calendario(pnNroCuotas - 1).dFecha + pnDiasGracia
            Calendario(pnNroCuotas).NroCuota = pnNroCuotas + 1
            Calendario(pnNroCuotas).IntGra = MatGracia(pnNroCuotas)
            Calendario(pnNroCuotas).Gasto = 0
            Calendario(pnNroCuotas).IntComp = 0#
            Calendario(pnNroCuotas).Cuota = Calendario(pnNroCuotas).IntGra
            Calendario(pnNroCuotas).Captital = 0#
            Calendario(pnNroCuotas).SaldoCap = 0#
        End If
    End If
End Sub

Public Function VerCalendarioD(ByVal psCtaCod As String, dFechaDesembolso As Date) As Recordset
    Dim i As Integer
    Dim dFecha As Date
    Dim nMonto As Double
    
    Dim rs  As ADODB.Recordset
'    Dim oCalend As NCalendario
    
    Dim MatCalendPagos As Variant
    Dim MatGracia As Variant
    
    Dim ssql As String
    Dim oConec As COMConecta.DCOMConecta

    Dim rs_Pagos As ADODB.Recordset
    
    Dim rs_Temp As ADODB.Recordset
    
    Dim nCuotas As Integer
    Dim nPeriodoGracia As Integer
    On Error GoTo ErrorDesembolsarCredito
    
            ssql = "Select PTI.nTasaInteres,CE.nCuotas,CE.nPlazo,CE.nColocCalendCod,CE.nTipoGracia,CE.nPeriodoFechaFija,CE.nProxMes"
            ssql = ssql & " From ProductoTasaInteres PTI"
            ssql = ssql & " Inner Join ColocacEstado CE on PTI.cCtaCod=CE.cCtaCod and CE.nPrdEstado=2002"
            ssql = ssql & " Inner Join ColocacCred CC on CC.cCtaCod=CE.cCtaCod"
            ssql = ssql & " Where PTI.nPrdTasaInteres=1 and PTI.cCtaCod='" & psCtaCod & "'"
            
            Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
            Set rs = oConec.CargaRecordSet(ssql)
            oConec.CierraConexion
            Set oConec = Nothing
    
    If rs!nColocCalendCod <> gColocCalendCodCL Then
                ' se verifica la fecha de desembolso sea diferente la fecha de aprobacion
            ' obteniendo la  fecha de desembolso
            ssql = "Select dPrdEstado,nMonto"
            ssql = ssql & " From ColocacEstado"
            ssql = ssql & " Where cCtacod='" & psCtaCod & "' and nPrdEstado=2002"
            
            Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
            Set rs = oConec.CargaRecordSet(ssql)
            oConec.CierraConexion
            Set oConec = Nothing
            
            If Not rs.EOF And Not rs.BOF Then
               dFecha = Format(rs!dPrdEstado, "dd/MM/yyyy")
               nMonto = rs!nMonto
            End If
            Set rs = Nothing
            
            ' Obteniendo la tasa de interes
            ssql = "Select PTI.nTasaInteres,CE.nCuotas,CE.nPlazo,CE.nColocCalendCod,CE.nTipoGracia,CE.nPeriodoFechaFija,CE.nProxMes,"
            ssql = ssql & " CE.nPeriodoGracia"
            ssql = ssql & " From ProductoTasaInteres PTI"
            ssql = ssql & " Inner Join ColocacEstado CE on PTI.cCtaCod=CE.cCtaCod and CE.nPrdEstado=2002"
            ssql = ssql & " Inner Join ColocacCred CC on CC.cCtaCod=CE.cCtaCod"
            ssql = ssql & " Where PTI.nPrdTasaInteres=1 and PTI.cCtaCod='" & psCtaCod & "'"
            
            Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
            Set rs = oConec.CargaRecordSet(ssql)
            oConec.CierraConexion
            Set oConec = Nothing
            
 '           Set oCalend = New NCalendario
            nPeriodoGracia = IIf(IsNull(rs!nPeriodoGracia), 0, rs!nPeriodoGracia)
            If nPeriodoGracia > 0 Then
            '    MatGracia = oCalend.GeneraGracia(Rs!nTipoGracia, CDbl(Format(TasaIntPerDias(Rs!nTasaInteres, nPeriodoGracia) * nMonto, "#0.00")), Rs!nCuotas)
                MatGracia = GeneraGracia(rs!nTipoGracia, CDbl(Format(TasaIntPerDias(rs!nTasaInteres, nPeriodoGracia) * nMonto, "#0.00")), rs!nCuotas)
            End If
            
            ssql = "Select nCuotas"
            ssql = ssql & " From ColocacEstado"
            ssql = ssql & " Where cCtaCod='" & psCtaCod & "' and nPrdEstado=2002"
            
            Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
            Set rs_Temp = oConec.CargaRecordSet(ssql)
            oConec.CierraConexion
            Set oConec = Nothing
            
            If Not rs_Temp.EOF And Not rs_Temp.BOF Then
                nCuotas = rs_Temp!nCuotas
            End If
            Set rs_Temp = Nothing
            
'            Set oCalend = New NCalendario
            If dFecha <> dFechaDesembolso Then
               'MatCalendPagos = oCalend.GeneraCalendario(nMonto, Rs!nTasaInteres, Rs!nCuotas, _
                                    Rs!nPlazo, dFechaDesembolso, DameTipoCuota(Rs!nColocCalendCod), DameTipoPeriodo(Rs!nColocCalendCod), _
                                    IIf(IsNull(Rs!nTipoGracia), 0, Rs!nTipoGracia), nPeriodoGracia, IIf(IsNull(Rs!nPeriodoFechaFija), 0, Rs!nPeriodoFechaFija), _
                                    IIf(Rs!nProxMes = 0, False, True), MatGracia, , , , , , , IIf(nCuotas > 1 And Mid(psCtaCod, 6, 3) = "320", True, False))
                
                MatCalendPagos = GeneraCalendario(nMonto, rs!nTasaInteres, rs!nCuotas, _
                                    rs!nPlazo, dFechaDesembolso, DameTipoCuota(rs!nColocCalendCod), DameTipoPeriodo(rs!nColocCalendCod), _
                                    IIf(IsNull(rs!nTipoGracia), 0, rs!nTipoGracia), nPeriodoGracia, IIf(IsNull(rs!nPeriodoFechaFija), 0, rs!nPeriodoFechaFija), _
                                    IIf(rs!nProxMes = 0, False, True), MatGracia, , , , , , , IIf(nCuotas > 1 And Mid(psCtaCod, 6, 3) = "320", True, False))
                                
             Set rs_Pagos = New ADODB.Recordset
             
             With rs_Pagos.Fields
                .Append "nCuota", adInteger
                .Append "Tipo", adVarChar, 20
                .Append "dVenc", adDate
                .Append "Capital", adDouble
                .Append "Interes", adDouble
             End With
             
             rs_Pagos.Open
             
             rs_Pagos.AddNew
             rs_Pagos(0) = "1"
             rs_Pagos(1) = "Desembolso"
             rs_Pagos(2) = Format(dFechaDesembolso, "dd/MM/yyyy")
             rs_Pagos(3) = nMonto
             rs_Pagos(4) = 0#
             
             rs_Pagos.Update
             
             For i = 0 To UBound(MatCalendPagos) - 1
                rs_Pagos.AddNew
                rs_Pagos(0) = MatCalendPagos(i, 1)
                rs_Pagos(1) = "Pago"
                rs_Pagos(2) = Format(MatCalendPagos(i, 0), "dd/MM/yyyy") ' fecha
                rs_Pagos(3) = MatCalendPagos(i, 3)
                rs_Pagos(4) = MatCalendPagos(i, 4)
                rs_Pagos.Update
             Next i
            End If
    End If
    Set VerCalendarioD = rs_Pagos
    Exit Function

ErrorDesembolsarCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function TasaIntPerDias(ByVal pnTasaInter As Double, ByVal pnDiasTrans As Integer) As Double
    TasaIntPerDias = ((1 + pnTasaInter / 100) ^ (pnDiasTrans / 30)) - 1
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

'Funcion que genera doble Calendario

Public Function GeneraDobleCalendario(ByRef pMatCalend_1 As Variant, ByRef pMatCalend_2 As Variant, _
                                    ByVal pnTramoConsMonto As Double, ByVal pnTramoNoConsMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                                    ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                                    ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                                    ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                                    Optional ByVal MatGracia As Variant, Optional ByVal pbMiViv As Boolean = False, _
                                    Optional ByVal pbGraciaCusco As Boolean = False, Optional ByVal pnTasaGracia As Double = 0) As Variant

        On Error GoTo ErrorGeneraDobleCalendario
        'MAVM 20121113 *** Se cambio: pMatCalend_1 (1ra variable) y pMatCalend_2 (9na variable)
        pMatCalend_1 = GeneraCalendario(pnTramoNoConsMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, _
                        pnTipoCuota, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, _
                        bProxMes, MatGracia, True, , , , , pbMiViv, , pbGraciaCusco, pnTasaGracia)

        'pMatCalend_2 = GeneraCalendario(pnTramoConsMonto, pnTasaInt, pnNroCuotas / 6, 180, pdFecDesemb, _
        '                 Fija, FechaFija, Exonerada, 0, pnDiaFijo, _
        '                bProxMes, MatGracia, True, , , , 6, , , pbGraciaCusco, pnTasaGracia)

        pMatCalend_2 = GeneraCalendario(pnTramoConsMonto, 0.534, pnNroCuotas / 6, 180, pdFecDesemb, _
                         Fija, FechaFija, Exonerada, pnDiasGracia, pnDiaFijo, _
                        bProxMes, MatGracia, True, , , , 6, , , pbGraciaCusco, pnTasaGracia)
        '***
        
        Exit Function

ErrorGeneraDobleCalendario:
        'MsgBox Err.Description, vbCritical, "Aviso"
        Err.Raise Err.Number, "Error", Err.Description
End Function

'Private Sub Class_Initialize()
'
'Set oImpre = New COMFunciones.FCOMImpresion
'Set oITF = New COMDConstSistema.FCOMITF
'End Sub
'
'Private Sub Class_Terminate()
'    Set oImpre = Nothing
'    Set oITF = Nothing
'End Sub

'->***** LUCV20180601, Comentó según ERS022-2018
Public Function GeneraCalendario(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, _
                Optional ByVal pbCuotaComodin As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
                Optional ByVal pMatDesPar As Variant = "", Optional ByVal pnNumMes As Integer = 1, Optional ByVal pbMiViv As Boolean = False, _
                Optional ByVal bQuincena As Boolean, Optional ByVal pbGraciaEnCuotas As Boolean = False, Optional ByVal pnTasaGracia As Double = 0, _
                Optional ByVal pnDiaFijo2 As Integer = 0, Optional pnMontoCapInicial As Double = 0, _
                Optional ByVal pbPagoInteresEnGracia As Boolean = False, _
                Optional ByVal pbRenovarCredito As Boolean = False, Optional ByVal pnInteresAFecha As Double = 0, _
                Optional ByVal pnMontoGra As Double = 0, Optional ByVal pnCuotaBalon As Integer = 0) As Variant
                'MAVM 20130305: pnMontoGra
                'WIOR 20131111 AGREGO pnCuotaBalon

'Si pnMontoCapInicial >0 entonces se Capitalizo la Gracia sin aumentar el capital
                
Dim sCalendPagos() As String
Dim i As Integer
Dim TasaComTemp As Double
Dim nPlazoCom As Integer
Dim nCuotasCom As Integer
Dim nCuotaMontoCom As Double
Dim oCred As COMNCredito.NCOMCredito

'Gracia en Cuotas
Dim nCuotasGracia As Integer
Dim nCuotaIni As Integer
Dim nCuotaFin As Integer
Dim dDesembolso As Date
        
        On Error GoTo ErrorGeneraCalendario
        
        If pbCuotaComodin Then
            TasaComTemp = pnTasaInt
            If pnTipoPeriodo = PeriodoFijo Then
                nPlazoCom = pnPeriodo
            Else
                nPlazoCom = 30
            End If
            nCuotasCom = pnNroCuotas
            Set oCred = New COMNCredito.NCOMCredito
            nCuotaMontoCom = oCred.CuotaFija(TasaComTemp, nCuotasCom + 1, pnMonto, nPlazoCom)
            nCuotaMontoCom = CDbl(Format((nCuotaMontoCom * (pnNroCuotas + 1)) / pnNroCuotas, "#0.00"))
            TasaComTemp = CDbl(Format(oCred.CreditosTasaEfectiva(nCuotaMontoCom, pnNroCuotas, pnMonto, pnTasaInt, nPlazoCom), "#0.0000"))
            pnTasaInt = Format(TasaComTemp * 100, "#0.0000")
            Set oCred = Nothing
        End If

    '******************************************
    '******** Gracia repartida en Cuotas ******
    
    nCuotasGracia = 0
    nCuotaIni = 0
    nCuotaFin = pnNroCuotas - 1
    ReDim Calendario(pnNroCuotas)
    
    If pbGraciaEnCuotas And pnTipoPeriodo = PeriodoFijo Then
        nCuotasGracia = Fix(pnDiasGracia / pnPeriodo)   'Entero inmediato inferior
        dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy"))
        ReDim Calendario(pnNroCuotas + nCuotasGracia)
        Set oCred = New COMNCredito.NCOMCredito
        For i = 0 To nCuotasGracia - 1
            Calendario(i).dFecha = IIf(i = nCuotasGracia - 1, CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia, dDesembolso + pnPeriodo)
            Calendario(i).NroCuota = i + 1
            'Agregado para interes Compensatorio
            If pbPagoInteresEnGracia Then
                If i = nCuotasGracia - 1 And nCuotasGracia > 1 Then
                    Calendario(i).IntComp = oCred.MontoIntPerDias(pnTasaInt, Calendario(i).dFecha - (Calendario(i - 1).dFecha), pnMonto)
                Else
                    Calendario(i).IntComp = oCred.MontoIntPerDias(pnTasaInt, pnPeriodo, pnMonto)
                End If
            Else
                Calendario(i).IntComp = 0#
            End If
            '****************************************
            Calendario(i).Gasto = 0#
            If i = nCuotasGracia - 1 And nCuotasGracia > 1 Then
                Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, Calendario(i).dFecha - (Calendario(i - 1).dFecha), pnMonto)
            Else
                If nCuotasGracia = 1 Then
                    Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, Calendario(i).dFecha - pdFecDesemb, pnMonto)
                Else
                    Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, pnPeriodo, pnMonto)
                End If
            End If
            
            Calendario(i).Captital = 0#
            'Agregado para interes de Gracia
            'Calendario(i).Cuota = Calendario(i).IntGra
            Calendario(i).Cuota = Calendario(i).IntGra + Calendario(i).IntComp
            '******************************************
            Calendario(i).SaldoCap = pnMonto
            dDesembolso = dDesembolso + pnPeriodo
        Next i
        nCuotaIni = nCuotasGracia
        nCuotaFin = pnNroCuotas + nCuotasGracia - 1
        ReDim Preserve Calendario(nCuotaFin + 1)
        Set oCred = Nothing
    End If
    
    '************************************
    
'***10-05-2006 *******************************************
    If pbGraciaEnCuotas And pnTipoPeriodo = FechaFija Then
        Dim nMes As Integer
        Dim nAnio As Integer
        Dim nDia As Integer
        Dim dFecTemp As Date
        Dim nDifDiasGracia As Integer
        
        nCuotasGracia = Fix(pnDiasGracia / 30)   'Entero inmediato inferior
        nDifDiasGracia = pnDiasGracia Mod 30    'Dias faltantes
        If nDifDiasGracia = 0 Then
            dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy"))
            'nMes = Month(dDesembolso)
        Else
            'If bProxMes Then
                dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + CDate(nDifDiasGracia)
                'nMes = Month(dDesembolso)
            '    nMes = Month(dDesembolso) + 1
            'End If
        End If
        nMes = Month(dDesembolso)
        
        ReDim Calendario(pnNroCuotas + nCuotasGracia)
        
        nAnio = Year(dDesembolso)
        nDia = pnDiaFijo
        
        Set oCred = New COMNCredito.NCOMCredito
        For i = 0 To nCuotasGracia - 1
            'Calendario(i).dFecha = IIf(i = nCuotasGracia - 1, CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia, dDesembolso + pnPeriodo)
            If Not (i = 0 And nDia > Day(dDesembolso) And (Not bProxMes)) Or pnNumMes = 6 Then
                If nDia = pnDiaFijo Then nMes = nMes + pnNumMes
                    If nMes > 12 Then
                        nAnio = nAnio + 1
                        nMes = 1
                    End If
                Else
                    If nDia > 30 Then
                        If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
                            nMes = nMes + pnNumMes
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
                    If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
                        nDia = 30
                    End If
                End If
            End If
            
            Calendario(i).NroCuota = i + 1
            dFecTemp = CDate(Right("0" & Trim(str(nDia)), 2) & "/" & Right("0" & Trim(str(nMes)), 2) & "/" & Trim(str(nAnio)))
            Calendario(i).dFecha = dFecTemp
            'Agregado para interes Compensatorio
            'If pbPagoInteresEnGracia Then
            '    If i = nCuotasGracia - 1 Then
            '        Calendario(i).IntComp = oCred.MontoIntPerDias(pnTasaInt, Calendario(i).dFecha - (Calendario(i - 1).dFecha), pnMonto)
            '    Else
            '        Calendario(i).IntComp = oCred.MontoIntPerDias(pnTasaInt, pnPeriodo, pnMonto)
            '    End If
            'Else
            '    Calendario(i).IntComp = 0#
            'End If
            '****************************************
            Calendario(i).Gasto = 0#
            
            If i = 0 Then
                Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, DateDiff("d", pdFecDesemb, Calendario(i).dFecha), pnMonto)
            Else
                Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, DateDiff("d", CDate(Calendario(i - 1).dFecha), Calendario(i).dFecha), pnMonto)
            End If
            
            Calendario(i).Captital = 0#
            'Agregado para interes de Gracia
            Calendario(i).Cuota = Calendario(i).IntGra + Calendario(i).IntComp
            '******************************************
            Calendario(i).SaldoCap = pnMonto
            'dDesembolso = dDesembolso + pnPeriodo
        Next i
        nCuotaIni = nCuotasGracia
        nCuotaFin = pnNroCuotas + nCuotasGracia - 1
        ReDim Preserve Calendario(nCuotaFin + 1)
        Set oCred = Nothing
        
End If
'**************************************************
        
        'ReDim Calendario(pnNroCuotas )
        Call ProcesarCalendario(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoCuota, _
                pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pbCuotaFijaFechaFija, _
                pbDesemParcial, pMatDesPar, pnNumMes, pbMiViv, bQuincena, nCuotaIni, nCuotaFin, pnDiaFijo2, pnMontoCapInicial, _
                pbRenovarCredito, pnInteresAFecha, pnMontoGra, pnCuotaBalon) 'MAVM 20130302: pnMontoGra
                'WIOR 20131111 AGREGO pnCuotaBalon
        
        'MAVM 20100320
        'ReDim sCalendPagos(UBound(Calendario), 8)
        'MAVM 20121113 ***
        'ReDim sCalendPagos(UBound(Calendario), 10)
        'ReDim sCalendPagos(UBound(Calendario), 13)
        ReDim sCalendPagos(UBound(Calendario), 14) 'RECO20150512
        '***
        
       For i = 0 To UBound(Calendario) - 1
            sCalendPagos(i, 0) = Format(Calendario(i).dFecha, "dd/mm/yyyy")
            sCalendPagos(i, 1) = Trim(str(Calendario(i).NroCuota))
            sCalendPagos(i, 2) = Format(Calendario(i).Cuota, "#0.00")
            sCalendPagos(i, 3) = Format(Calendario(i).Captital, "#0.00")
            sCalendPagos(i, 4) = Format(Calendario(i).IntComp, "#0.00")
            'MAVM 20130209 ***
            'sCalendPagos(i, 5) = Format(Calendario(i).IntGra, "#0.00")
            If Not (pnTipoGracia = 1 And i = 0) Then
                sCalendPagos(i, 5) = Format(Calendario(i).IntGra, "#0.00")
            Else
                Set oCred = New COMNCredito.NCOMCredito
                sCalendPagos(i, 5) = Format(Calendario(i).IntGra + oCred.MontoIntPerDias(CDbl(pnTasaGracia), IIf(pnTipoPeriodo = 1, CInt(pnPeriodo), 30), CDbl(Trim(Calendario(i).IntGra))), "#0.00")
                sCalendPagos(i, 2) = sCalendPagos(i, 2) + oCred.MontoIntPerDias(CDbl(pnTasaGracia), IIf(pnTipoPeriodo = 1, CInt(pnPeriodo), 30), CDbl(Trim(Calendario(i).IntGra)))
                Set oCred = Nothing
            End If
            '***
            sCalendPagos(i, 6) = Format(Calendario(i).Gasto, "#0.00")
            sCalendPagos(i, 7) = Format(Calendario(i).SaldoCap, "#0.00")
            sCalendPagos(i, 8) = Format(Calendario(i).SegDes, "#0.00")
            sCalendPagos(i, 9) = Format(Calendario(i).SegBien, "#0.00") 'MAVM 20121113
            'MAVM 20130209 ***
            sCalendPagos(i, 10) = Format(Calendario(i).CuotaGra, "#0.00")
            sCalendPagos(i, 11) = Format(Calendario(i).IntCompGra, "#0.00")
            sCalendPagos(i, 12) = Format(Calendario(i).SaldoCapGra, "#0.00")
            '***
            sCalendPagos(i, 13) = Format(Calendario(i).SegMultMype, "#0.00") 'RECO
            sCalendPagos(i, 14) = Format(0#, "#0.00") 'RECO
        Next i
        GeneraCalendario = sCalendPagos
        Exit Function

ErrorGeneraCalendario:
        'MsgBox Err.Description, vbCritical, "Aviso"
        Err.Raise Err.Number, "Error", Err.Description
End Function
'<-***** Fin LUCV20180601

'->***** LUCV20180601, Agregó según ERS022-2018 [Esto reemplazará el met. GeneraCalendario]
Public Function GeneraCalendarioNuevo(ByVal pnMonto As Double, _
                                ByVal pnTasaInt As Double, _
                                ByVal pnNroCuotas As Integer, _
                                ByVal pnPeriodo As Double, _
                                ByVal pdFecDesemb As Date, _
                                ByVal pnTipoCuota As TCalendTipoCuota, _
                                ByVal pnTipoPeriodo As TCalendTipoPeriodo, _
                                ByVal pnTipoGracia As TCalendTipoGracia, _
                                ByVal pnDiasGracia As Integer, _
                                ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, _
                                Optional ByVal pbCuotaComodin As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
                                Optional ByVal pMatDesPar As Variant = "", Optional ByVal pnNumMes As Integer = 1, _
                                Optional ByVal pbMiViv As Boolean = False, Optional ByVal bQuincena As Boolean, _
                                Optional ByVal pbGraciaEnCuotas As Boolean = False, Optional ByVal pnTasaGracia As Double = 0, _
                                Optional ByVal pnDiaFijo2 As Integer = 0, Optional pnMontoCapInicial As Double = 0, _
                                Optional ByVal pbPagoInteresEnGracia As Boolean = False, Optional ByVal pbRenovarCredito As Boolean = False, _
                                Optional ByVal pnInteresAFecha As Double = 0, Optional ByVal pnMontoGra As Double = 0, _
                                Optional ByVal pnCuotaBalon As Integer = 0, _
                                Optional ByVal psCtaCod As String, _
                                Optional ByVal pnTasaSegDes As Double = 0, _
                                Optional ByRef pMatCalendSegDes As Variant, _
                                Optional ByVal pnExoSeguroDesgravamen As Integer = 0, _
                                Optional ByVal pnMontoPoliza As Double, _
                                Optional ByVal pnTasaSegInc As Double) As Variant
                                'LUCV20180601, Agregó psCtaCod, pnTasaSegDes, pMatCalendSegDes, pnExoSeguroDesgravamen, pnMontoPoliza, pnTasaSegInc según ERS022-2018
                
Dim sCalendPagos() As String
Dim i As Integer
Dim oNCOMCredito As COMNCredito.NCOMCredito
Dim nCuotaIni As Integer
Dim nCuotaFin As Integer
'Dim nEEMSegDesg As Double 'LUCV20180601:Equivalente Efectivo Mensual del Seguro de Desgravamen

    On Error GoTo ErrorGeneraCalendarioNuevo
    Set oNCOMCredito = New COMNCredito.NCOMCredito
    
    nCuotaIni = 0
    nCuotaFin = pnNroCuotas - 1
    ReDim Calendario(pnNroCuotas)
    Call ProcesarCalendario(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, _
                        pnTipoCuota, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, _
                        bProxMes, MatGracia, pbCuotaFijaFechaFija, pbDesemParcial, pMatDesPar, _
                        pnNumMes, pbMiViv, bQuincena, nCuotaIni, nCuotaFin, pnDiaFijo2, pnMontoCapInicial, _
                        pbRenovarCredito, pnInteresAFecha, pnMontoGra, pnCuotaBalon, _
                        psCtaCod, pnTasaSegDes, pMatCalendSegDes, , pnExoSeguroDesgravamen, pnMontoPoliza, pnTasaSegInc)
                        'LUCV20180601, Agregó: pnTasaSegDes, pMatCalendSegDes, pnExoSeguroDesgravamen, pnMontoPoliza, pnTasaSegInc según ERS022-2018

    ReDim sCalendPagos(UBound(Calendario), 16) 'LUCV20180601, Cambió 14 por 16. Según ERS022-2018
    
    For i = 0 To UBound(Calendario) - 1
         sCalendPagos(i, 0) = Format(Calendario(i).dFecha, "dd/mm/yyyy")
         sCalendPagos(i, 1) = Trim(str(Calendario(i).NroCuota))
         sCalendPagos(i, 2) = Format(Calendario(i).Cuota, "#0.00")
         sCalendPagos(i, 3) = Format(Calendario(i).Captital, "#0.00")
         sCalendPagos(i, 4) = Format(Calendario(i).IntComp, "#0.00")
         sCalendPagos(i, 5) = Format(Calendario(i).IntGra, "#0.00") 'LUCV20180601, Según ERS022-2018
         sCalendPagos(i, 6) = Format(Calendario(i).Gasto, "#0.00") 'LUCV20180601, Según ERS022-2018
         sCalendPagos(i, 7) = Format(Calendario(i).SaldoCap, "#0.00")
         sCalendPagos(i, 8) = Format(Calendario(i).SegDes, "#0.00")
         sCalendPagos(i, 9) = Format(Calendario(i).SegBien, "#0.00")
         sCalendPagos(i, 10) = Format(Calendario(i).CuotaGra, "#0.00")
         sCalendPagos(i, 11) = Format(Calendario(i).IntCompGra, "#0.00")
         sCalendPagos(i, 12) = Format(Calendario(i).SaldoCapGra, "#0.00")
         sCalendPagos(i, 13) = Format(Calendario(i).SegMultMype, "#0.00")
         sCalendPagos(i, 14) = Format(0#, "#0.00")
         sCalendPagos(i, 15) = Format(Calendario(i).CuotaPrimaPoliza, "#0.00") 'LUCV20180601, Según ERS022-2018
         sCalendPagos(i, 16) = Format(Calendario(i).CuotaPrimaPolizaGracia, "#0.00") 'LUCV20180601, Según ERS022-2018
     Next i
        GeneraCalendarioNuevo = sCalendPagos
    Exit Function

ErrorGeneraCalendarioNuevo:
        Err.Raise Err.Number, "Error", Err.Description
End Function
'<-***** Fin LUCV20180601

Public Function GeneraCalendarioLeasing(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, _
                Optional ByVal pbCuotaComodin As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
                Optional ByVal pMatDesPar As Variant = "", Optional ByVal pnNumMes As Integer = 1, Optional ByVal pbMiViv As Boolean = False, _
                Optional ByVal bQuincena As Boolean, Optional ByVal pbGraciaEnCuotas As Boolean = False, Optional ByVal pnTasaGracia As Double = 0, _
                Optional ByVal pnDiaFijo2 As Integer = 0, Optional pnMontoCapInicial As Double = 0, _
                Optional ByVal pbPagoInteresEnGracia As Boolean = False, _
                Optional ByVal pbRenovarCredito As Boolean = False, Optional ByVal pnInteresAFecha As Double = 0, Optional ByVal lsCtaCodLeasing As String = "") As Variant

'Si pnMontoCapInicial >0 entonces se Capitalizo la Gracia sin aumentar el capital
                
Dim sCalendPagos() As String
Dim i As Integer
Dim TasaComTemp As Double
Dim nPlazoCom As Integer
Dim nCuotasCom As Integer
Dim nCuotaMontoCom As Double
Dim oCred As COMNCredito.NCOMCredito
Dim oDCred As COMDCredito.DCOMleasing
Dim oRS As ADODB.Recordset
'Gracia en Cuotas
Dim nCuotasGracia As Integer
Dim nCuotaIni As Integer
Dim nCuotaFin As Integer
Dim dDesembolso As Date
        
        On Error GoTo ErrorGeneraCalendarioLeasing

        Set oDCred = New COMDCredito.DCOMleasing
        Set oRS = New ADODB.Recordset
        Set oRS = oDCred.RecuperaCalendarioLeasing(lsCtaCodLeasing)
'        ReDim sCalendPagos(UBound(Calendario), 9)
'        If (oRs.EOF Or oRs.BOF) Then
'            GeneraCalendarioLeasing = sCalendPagos
'            Exit Function
'        End If
        i = 0
        ReDim sCalendPagos(oRS.RecordCount, 12)
        Do While Not oRS.EOF
     
            sCalendPagos(i, 0) = Format(oRS!dVec, "dd/mm/yyyy")
            sCalendPagos(i, 1) = Trim(str(oRS!nNroCuota))
            sCalendPagos(i, 2) = Format(oRS!nCuotaTotal, "#0.00")
            sCalendPagos(i, 3) = Format(oRS!nCapital, "#0.00")
            sCalendPagos(i, 4) = Format(oRS!nInteres, "#0.00")
            sCalendPagos(i, 5) = Format(0, "#0.00") 'Interes de gracia
            sCalendPagos(i, 6) = Format(oRS!nGasto, "#0.00") 'gasto
            sCalendPagos(i, 7) = Format(oRS!nSaldoCap, "#0.00")
            sCalendPagos(i, 8) = Format(oRS!nSeguroDesgravamen, "#0.00") 'Seguro desgravament
            'ALPA20130923********************
            sCalendPagos(i, 9) = 0#
            sCalendPagos(i, 10) = 0#
            sCalendPagos(i, 11) = 0#
            sCalendPagos(i, 12) = 0#
            '********************************

        'Next i
        i = i + 1
        oRS.MoveNext
        Loop
        GeneraCalendarioLeasing = sCalendPagos
        Exit Function

ErrorGeneraCalendarioLeasing:
        'MsgBox Err.Description, vbCritical, "Aviso"
        Err.Raise Err.Number, "Error", Err.Description
End Function

