VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
    'MAVM 20100320
    SegDes As Double
    CuotaPrimaPoliza As Double 'LUCV20180601, Según ERS022-2018
    CuotaPrimaPolizaGracia As Double 'LUCV20180601, Según ERS022-2018
    SaldoCap As Double
    SegInm As Double 'MAVM 20120625
    'MAVM 20130209 ***
    CuotaGra As Double
    IntCompGra As Double
    SaldoCapGra As Double
    '***
    
    'RIRO 20200825 Mejora en Liquidación ****
    IntCompCalculado As Double
    DiasCalculo As Integer
    IntCompDiferenciaCapitalizado As Double
    IntGraciaGenerado As Double
    IntGraciaCapitalizado As Double
    IntGraciaAsignado As Double
    'RIRO 20200825 Mejora en Liquidación ****
    
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
Public oImp As COMFunciones.FCOMVarImpresion

Public Function GeneraGracia(ByVal pOptGracia As TCalendTipoGracia, ByVal pnInteres As Double, ByVal pnNroCuotas As Integer) As Variant
Dim nValorProrateado As Double
Dim i As Integer
Dim MatGracia() As Double
Dim sMatGracia() As String

    nValorProrateado = pnInteres / pnNroCuotas
    nValorProrateado = CDbl(Format(nValorProrateado, "#0.00"))
    ReDim MatGracia(pnNroCuotas)
    If pOptGracia = PrimeraCuota Then
        'ReDim Preserve MatGracia(pnNroCuotas + 1)
        ReDim Preserve MatGracia(pnNroCuotas)
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
                dFecTemp = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)))
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
                dFecTemp = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)))
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

'->*****LUCV20180601, Deshabilitó para ser reemplazado por [CalendarioCuotaFijaNuevo]
Private Sub CalendarioCuotaFija(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
                Optional ByVal pMatDesPar As Variant = "", Optional ByVal pnNumMes As Integer = 1, Optional ByVal pbMiViv As Boolean = False, _
                Optional ByVal pnCuotaIni As Integer = 0, Optional ByVal pnCuotaFin As Integer = 0, Optional ByVal pnDiaFijo2 As Integer = 0, _
                Optional pnMontoCapInicial As Double = 0, _
                Optional ByVal psCtaCod As String, Optional ByVal psMensFecAprob As Date, Optional ByVal lnDescripCuota As Variant = "", _
                Optional ByVal pnMontoGra As Double = 0, Optional ByVal psIncluyeIntCap As String = "S", _
                Optional ByVal pnCuotaBalon As Integer = 0, _
                Optional ByVal pbReprogConvenio As Boolean = False) 'MAVM 11092010 Se agrego el par: psCtaCod (Prepago)
                'MADM 20110409 - FECHA APRO - lnDescripCuota
                'WIOR 20131111 AGREGO pnCuotaBalon
                'WIOR 20140514 AGREGO pbReprogConvenio

Dim nSaldoCapital As Double
Dim dDesembolso As Date
Dim dFecAproGracia As Date
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
Dim nPeriodoPrimerPF As Integer 'MADM 20110503

'MAVM 20130209 ***
Dim nSumCapitalGra As Double
Dim nSaldoCapitalGra As Double
'***
Dim nNroCuotasAfec As Integer 'WIOR 20131127

nSumCapital = 0
nPeriodoPrimerPF = 0

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

        'MADM 20120221-PAGOS -----------------------------------------------------------
        If psMensFecAprob = "12:00:00 AM" And psCtaCod <> "" Then
            Dim oCredT As COMNCredito.NCOMCredito
            Dim RMantPP As ADODB.Recordset
            Dim DfechaPre As Date
            Dim PlazoPre As Integer
            Dim indPre As Boolean

            Set oCredT = New COMNCredito.NCOMCredito
            Set RMantPP = oCredT.RecuperaValidacionCredMantPrepago(psCtaCod, Format(pdFecDesemb, "dd/mm/yyyy"))
            Set oCredT = Nothing
            indPre = False

            If Not RMantPP.BOF And Not RMantPP.EOF Then
                indPre = True
                DfechaPre = RMantPP!dVenc 'CC.nCuota , Ce.nNroProxCuota, CC.nNroCalen,
                PlazoPre = RMantPP!DifPlazo
                RMantPP.Close
            End If
            Set RMantPP = Nothing
        End If
        '-------------------------------------------------------------------------------

        nSaldoCapital = pnMonto
        nSaldoCapitalGra = pnMontoGra 'MAVM 20130209
        dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia

        If pnTipoPeriodo = PeriodoFijo Then
            Set oCredito = New NCOMCredito
            'For i = 0 To pnNroCuotas - 1
            For i = pnCuotaIni To pnCuotaFin

            'MADM 20110518 - No entre cuando es aprobacion o PAGO
              If i = 0 And psMensFecAprob <> "12:00:00 AM" Then
                'MADM 20110420 - Verifica Si tiene Gracia / Diferencia es -
                    If i = 0 And (dDesembolso > psMensFecAprob) Then
                          If lnDescripCuota(0) = "0" Then
                            nPeriodoPrimerPF = pnPeriodo - CInt(DateDiff("d", psMensFecAprob, dDesembolso))
                          Else
                                If lnDescripCuota(1) <> "0" Then
                                    nPeriodoPrimerPF = pnPeriodo + Val(lnDescripCuota(1))
                                End If
                          End If
                    End If
              End If
              'END MADM
               'MADM 20120221 ----------------------------------------------------------
                If indPre = True Then
                    Calendario(i).dFecha = dDesembolso + IIf(i = 0, PlazoPre, pnPeriodo)
                Else
                    Calendario(i).dFecha = dDesembolso + IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo)
                End If
                '-------------------------------------------------------------------------
                Calendario(i).NroCuota = i + 1
                Calendario(i).IntGra = 0#
                Calendario(i).Gasto = 0#
                'MADM 20110518 - Periodo
                If indPre = True Then
                    Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, IIf(i = 0, PlazoPre, pnPeriodo), nSaldoCapital)
                    Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, IIf(i = 0, PlazoPre, pnPeriodo))
                Else
                     'WIOR 20131111 **********************************
                    If i < pnCuotaBalon Then
                        nSaldoCapital = pnMonto
                    End If
                    'WIOR FIN ***************************************
                    Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), nSaldoCapital)

                    'MAVM 20130430***
                    If pnTipoGracia = 6 Then
                        Calendario(i).IntCompGra = oCredito.MontoIntPerDias(pnTasaInt, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), nSaldoCapitalGra)
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
                    'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo)

                    'MAVM 20130430***
                    If pnTipoGracia = 6 Then
                        Calendario(i).CuotaGra = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMontoGra, pnPeriodo)
                    End If
                    '***

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
                nSaldoCapital = CDbl(Format(nSaldoCapital, "#0.00"))

                'MAVM 20130430***
                If pnTipoGracia = 6 Then
                    nSumCapitalGra = nSumCapitalGra + Calendario(i).IntGra
                    Calendario(i).SaldoCapGra = nSaldoCapitalGra - Calendario(i).IntGra
                    nSaldoCapitalGra = nSaldoCapitalGra - Calendario(i).IntGra
                End If

                'MADM 20120221 ----------------------------------------------------------
                If indPre = True Then
                    dDesembolso = dDesembolso + IIf(i = 0, PlazoPre, pnPeriodo)
                Else
                    dDesembolso = dDesembolso + IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo)
                End If
                '-------------------------------------------------------------------------
                'MADM 20110518 - + IIf(nPeriodoPrimerPF <> 0, nPeriodoPrimerPF, pnPeriodo)

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
                nPeriodoPrimerPF = 0
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

                        'MAVM 20110108 ***
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

                            If CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy")) - dDesembolso < 5 Then
                                nMes = nMes + pnNumMes
                            End If
                        End If
                        '***

'                        If nMes > 12 Then
'                            nAnio = nAnio + 1
'                            nMes = 1
'                        End If

                    '*****************************************
                    Else
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
                            If nDia > 30 Then
                                If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
                                    nMes = nMes + pnNumMes
                                End If
                            End If
'                        End If


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
                    'By Capi 06122008 formateo de fecha

                    'MAVM 20100820 ***
                    If psCtaCod <> "" Then
                        If i = 0 Then
                            If CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy")) - dDesembolso < 20 Then
                                pnDiasGracia = CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy")) - dDesembolso
                                Dim pnMontPeridExtra As Double
                                pnMontPeridExtra = oCredito.MontoIntPerDias(pnTasaInt, pnDiasGracia, nSaldoCapital)
                                dDesembolso = CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy"))
                                nMes = nMes + pnNumMes
                                If nMes > 12 Then
                                    nAnio = nAnio + 1
                                    nMes = 1
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
                            End If
                        End If
                    End If
                        '***

                    dFecTemp = CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy"))

                    'If (dFecTemp > pdFecDesemb + CDate(pnDiasGracia) And bProxMes) Then
                    'If (dFecTemp > dDesembolso And bProxMes) Then
                    '    Calendario(i).dFecha = DateAdd("m", -Fix(pnDiasGracia / 30), dFecTemp)
                    'Else
                        Calendario(i).dFecha = dFecTemp
                    'End If
                    Calendario(i).NroCuota = i + 1
                    Calendario(i).IntGra = 0#
                    Calendario(i).Gasto = 0
                    'WIOR 20131111 **********************************
                    If i < pnCuotaBalon Then
                        nSaldoCapital = pnMonto
                    End If
                    nNroCuotasAfec = pnNroCuotas - pnCuotaBalon
                    'WIOR FIN ***************************************
                    'If i = 0 Then
                    If i = pnCuotaIni Then
                    '11-05-2006
                        If pbCuotaFijaFechaFija Then
                            'Set oCredito = New COMNCredito.NCOMCredito
                            'nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas)
                            'dFecTemp - dDesembolso +
                            'If pnDiasGracia = 0 Then
                                'nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas) 'WIOR 20131127 COMENTÓ
                                nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, nNroCuotasAfec, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas) 'WIOR 20131127 COMENTÓ
                            'Else
                            '    nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, dFecTemp - dDesembolso + pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas)
                            'End If
                            'Set oCredito = Nothing

                            'MAVM 20130209***
                            If pnTipoGracia = 6 Then
                                nMontoGraciaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMontoGra, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas)
                            End If
                            '***

                        End If
                    '**************************************
                        If pbCuotaFijaFechaFija And pbMiViv Then
                            dDesembolso = dDesembolso - pnDiasGracia
                        End If
                        Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(i).dFecha), nSaldoCapital)

                        'MAVM 20130209***
                        If pnTipoGracia = 6 Then
                            Calendario(i).IntCompGra = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(i).dFecha), pnMontoGra)
                        End If

                        If Not pbCuotaFijaFechaFija Then
                            If DateDiff("d", dDesembolso, Calendario(i).dFecha) = 0 And Len(Trim(psCtaCod)) = 0 Then
                                'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, 30) 'WIOR 20131111 COMENTÓ
                                'WIOR 20131111 **********************************
                                If i < pnCuotaBalon Then
                                    Calendario(i).Cuota = Calendario(i).IntComp
                                Else
                                    Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, nNroCuotasAfec, pnMonto, 30)
                                End If
                                'WIOR FIN ***************************************

                            Else
                                'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, DateDiff("d", dDesembolso, Calendario(i).dFecha)) 'WIOR 20131111 COMENTÓ
                                'WIOR 20131111 **********************************
                                If i < pnCuotaBalon Then
                                    Calendario(i).Cuota = Calendario(i).IntComp
                                Else
                                    Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, nNroCuotasAfec, pnMonto, DateDiff("d", dDesembolso, Calendario(i).dFecha))
                                End If
                                'WIOR FIN ***************************************

                            End If
                        Else
                            'Calendario(i).Cuota = nMontoCuotaTemp 'WIOR 20131111 COMENTÓ
                            'WIOR 20131111 **********************************
                            If i < pnCuotaBalon Then
                                Calendario(i).Cuota = Calendario(i).IntComp
                            Else
                                Calendario(i).Cuota = nMontoCuotaTemp
                            End If
                            'WIOR FIN ***************************************

                             'MAVM 20130209***
                            If pnTipoGracia = 6 Then
                                Calendario(i).CuotaGra = nMontoGraciaTemp
                            End If
                            '***

                        End If

                    Else
                        Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha), nSaldoCapital)
                        'MAVM 20130209***
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
                            'MAVM 20130209***
                            If pnTipoGracia = 6 Then
                                Calendario(i).CuotaGra = nMontoGraciaTemp
                            End If
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
                        'MAVM 20130209***
                        If pnTipoGracia = 6 Then
                            Calendario(i).IntGra = nSaldoCapitalGra
                        End If
                        '***
                        If (Calendario(i).IntComp + Calendario(i).Captital) > Calendario(i).Cuota Then
                            Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                        Else
                            Calendario(i).IntComp = Calendario(i).Cuota - Calendario(i).Captital
                             'MAVM 20130209***
                            If pnTipoGracia = 6 Then
                                Calendario(i).IntCompGra = Calendario(i).CuotaGra - Calendario(i).IntGra
                            End If
                            '***
                        End If
                    Else
                        Calendario(i).Captital = Calendario(i).Cuota - Calendario(i).IntComp
                        'MAVM 20130209***
                        If pnTipoGracia = 6 Then
                            Calendario(i).IntGra = Calendario(i).CuotaGra - Calendario(i).IntCompGra
                        End If
                        '***
                    End If

                    nSumCapital = nSumCapital + Calendario(i).Captital
                    Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                    nSaldoCapital = nSaldoCapital - Calendario(i).Captital

                    'MAVM 20130209***
                    If pnTipoGracia = 6 Then
                        nSumCapitalGra = nSumCapitalGra + Calendario(i).IntGra
                        Calendario(i).SaldoCapGra = nSaldoCapitalGra - Calendario(i).IntGra
                        nSaldoCapitalGra = nSaldoCapitalGra - Calendario(i).IntGra
                    End If

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
    If (pnDiasGracia > 0 _
       And pnCuotaIni = 0) Or (pbReprogConvenio And pnCuotaIni = 0) Then 'WIOR 20140514 AGREGO (pbReprogConvenio And pnCuotaIni = 0)
        If pnTipoGracia = PrimeraCuota Then
            'ReDim Preserve Calendario(pnNroCuotas + 1)
            'For i = pnNroCuotas To 1 Step -1
            '    Calendario(i) = Calendario(i - 1)
            '    Calendario(i).NroCuota = Calendario(i).NroCuota + 1
            'Next i
            'Calendario(0).Dfecha = pdFecDesemb + pnDiasGracia
            'Calendario(0).NroCuota = 1
            'Calendario(0).IntGra = MatGracia(0)
            'Calendario(0).Gasto = 0
            'Calendario(0).IntComp = 0#
            'Calendario(0).Cuota = Calendario(0).IntGra
            'Calendario(0).Captital = 0#
            'Calendario(0).SaldoCap = pnMonto

            'MAVM 20130209 ***
            Calendario(0).IntGra = MatGracia(0)
            Calendario(0).Cuota = Calendario(0).Cuota + MatGracia(0)
            '***
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

        'MAVM 20120209 ***
        If pnTipoGracia = gColocTiposGraciaCapitalizada And psIncluyeIntCap = "S" Then
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
            Calendario(0).SaldoCap = pnMonto '+ pnMontoGra
            Calendario(0).SaldoCapGra = 0#
            Calendario(0).CuotaGra = 0#
            Calendario(0).IntCompGra = 0#
        End If
        '***

        'MAVM 20100820 ***
        If pnMontPeridExtra <> 0 Then
            Dim nCuotaNormalPago As Double
            Dim nFactor As Double
            Dim oCred As COMDCredito.DCOMCredito
            Set oCred = New COMDCredito.DCOMCredito

            nCuotaNormalPago = oCred.CuotaNormalPagoSinGastos(psCtaCod)

            If (nCuotaNormalPago > Calendario(0).Cuota) And ((nCuotaNormalPago - Calendario(0).Cuota) * pnNroCuotas > pnMontPeridExtra) Then
                nFactor = nCuotaNormalPago - Calendario(0).Cuota

                For i = 0 To pnNroCuotas - 1
                    If pnMontPeridExtra > 0 Then
                        If pnMontPeridExtra > nFactor Then
                            Calendario(i).IntComp = Calendario(i).IntComp + nFactor
                            Calendario(i).Cuota = Calendario(i).Cuota + nFactor
                            pnMontPeridExtra = pnMontPeridExtra - nFactor
                        Else
                            Calendario(i).IntComp = Calendario(i).IntComp + pnMontPeridExtra
                            Calendario(i).Cuota = Calendario(i).Cuota + pnMontPeridExtra
                            pnMontPeridExtra = pnMontPeridExtra - pnMontPeridExtra
                        End If
                    End If
                Next i

            Else
                nFactor = Round(pnMontPeridExtra / (pnNroCuotas), 2)
                For i = 0 To pnNroCuotas - 1
                    If i <> 0 Then
                        Calendario(i).IntComp = Calendario(i).IntComp + nFactor
                        Calendario(i).Cuota = Calendario(i).Cuota + nFactor
                    End If
                Next i
            End If

        End If
        '***

    End If
End Sub
'<-*****Fin LUCV20180601

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

Private Sub ProcesarCalendario(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, _
                Optional ByVal pbDesemParcial As Boolean = False, Optional ByVal pMatDesPar As Variant = "", _
                Optional ByVal pnNumMes As Integer = 1, _
                Optional ByVal pbMiViv As Boolean = False, _
                Optional ByVal bQuincena As Boolean, _
                Optional ByVal pnCuotaIni As Integer = 0, _
                Optional ByVal pnCuotaFin As Integer = 0, _
                Optional ByVal pnDiaFijo2 As Integer = 0, _
                Optional pnMontoCapInicial As Double = 0, _
                Optional psCtaCod As String, _
                Optional ByVal psMensFecAprob As Date, _
                Optional ByVal lnDescripCuota As Variant = "", _
                Optional ByVal pnMontoGra As Double = 0, _
                Optional ByVal psIncluyeIntCap As String = "S", _
                Optional ByVal pnCuotaBalon As Integer = 0, _
                Optional ByVal pbReprogConvenio As Boolean = False, _
                Optional ByVal pbReduccionCuotaPagAnt As Boolean = False, _
                Optional ByVal pnTasaSegDes As Double = 0, Optional ByRef pMatCalendSegDes As Variant, Optional ByVal pnExoSeguroDesgravamen As Integer = 0, _
                Optional ByVal pnMontoPoliza As Double = 0, Optional ByVal pnTasaSegInc As Double = 0, Optional ByRef pArrayDatos As Variant = Nothing)
                'MAVM 11092010 Se agrego el par: psCtaCod (Prepago)
                'madm 20110531 - lnDescripCuota
                'MAVM 20130209
                'WIOR 20131111 AGREGO pnCuotaBalon
                'WIOR 20140514 AGREGO pbReprogConvenio
                'LUCV20180320, Agregó: pbReduccionCuotaPagAnt
                'LUCV20180601, Agregó : pnTasaSegDes, pnExoneraSegDes, pMatCalendSegDes, pnExoSeguroDesgravamen, pnMontoPoliza,pnTasaSegInc, pArrayDatos
                
                If bQuincena = False Then
                    Select Case pnTipoCuota
                        Case Creciente
                            Call CalendarioCreciente(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pnCuotaIni, pnCuotaFin, pnDiaFijo2, pnMontoCapInicial)
                        Case Decreciente
                            Call CalendarioDecreciente(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pnCuotaIni, pnCuotaFin, pnDiaFijo2, pnMontoCapInicial)
                        Case Fija
                            '--->***** LUCV20180320, Según ERS087-2017
'                            Call CalendarioCuotaFija(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pbCuotaFijaFechaFija, pbDesemParcial, _
'                                                pMatDesPar, pnNumMes, pbMiViv, pnCuotaIni, pnCuotaFin, pnDiaFijo2, pnMontoCapInicial, psCtaCod, psMensFecAprob, lnDescripCuota, pnMontoGra, psIncluyeIntCap, pnCuotaBalon, pbReprogConvenio)
'                            'WIOR 20131111 AGREGO pnCuotaBalon
'                            'WIOR 20140514 AGREGO pbReprogConvenio

'->***** LUCV20180601, Comentó según ERS022-2018
'                            If Not pbReduccionCuotaPagAnt Then
'                            'comment by marg20180514 pag.ant.---------
''                                Call CalendarioCuotaFija(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, _
''                                            pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pbCuotaFijaFechaFija, pbDesemParcial, _
''                                            pMatDesPar, pnNumMes, pbMiViv, pnCuotaIni, pnCuotaFin, pnDiaFijo2, pnMontoCapInicial, _
''                                            psCtaCod, psMensFecAprob, lnDescripCuota, pnMontoGra, psIncluyeIntCap, pnCuotaBalon, _
''                                            pbReprogConvenio)
'                            'marg20180514----------------------------
                            
'                            'add by marg20180514 pag.ant.---------
'                                Call CalendarioCuotaFijaNew(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, _
'                                            pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pbCuotaFijaFechaFija, pbDesemParcial, _
'                                            pMatDesPar, pnNumMes, pbMiViv, pnCuotaIni, pnCuotaFin, pnDiaFijo2, pnMontoCapInicial, _
'                                            psCtaCod, psMensFecAprob, lnDescripCuota, pnMontoGra, psIncluyeIntCap, pnCuotaBalon, _
'                                            pbReprogConvenio)
'                            'marg20180514----------------------------
                            
                                Call CalendarioCuotaFijaNuevo(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, _
                                                                pdFecDesemb, pnTipoPeriodo, pnTipoGracia, _
                                                                pnDiasGracia, pnDiaFijo, pnNumMes, _
                                                                pnCuotaIni, pnCuotaFin, psCtaCod, _
                                                                psMensFecAprob, lnDescripCuota, pnMontoGra, _
                                                                pnTasaSegDes, pMatCalendSegDes, _
                                                                pnExoSeguroDesgravamen, pnMontoPoliza, pnTasaSegInc, _
                                                                pbReduccionCuotaPagAnt, pArrayDatos)
                                'LUCV20180601. Agregó pnTasaSegDes, pMatCalendSegDes, pnExoneraSegSDes, pnMontoPoliza, pnMontoPoliza, pnTasaSegInc . Según ERS022-2018
'                            Else
'                                Call CalendarioCuotaFijaPagoAnticipado(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, _
'                                            pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pbCuotaFijaFechaFija, pbDesemParcial, _
'                                            pMatDesPar, pnNumMes, pbMiViv, pnCuotaIni, pnCuotaFin, pnDiaFijo2, pnMontoCapInicial, psCtaCod, _
'                                            psMensFecAprob, lnDescripCuota, pnMontoGra, psIncluyeIntCap, pnCuotaBalon, pbReprogConvenio)
'                            End If
                             '<---***** Fin LUCV20180320
'<-***** Fin LUCV20180601
                    End Select
                Else
                    If bQuincena = True Then
                        Call CalendarioTrabajadoresDirectores(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, bProxMes, pnDiasGracia, MatGracia, pnTipoGracia, pnCuotaIni, pnCuotaFin)
                    End If
                End If
End Sub
Public Sub inicio(ByVal psNomCmac As String, ByVal psNomAgencia As String, _
            ByVal psCodUser As String, ByVal psFechaSis As String)
    csNomCMAC = psNomCmac
    csNomAgencia = psNomAgencia
    csCodUser = psCodUser
    csFechaSis = psFechaSis
End Sub

'->***** LUCV20180601, Comentó según ERS022-2018
'Public Function GeneraCalendario(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
'                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
'                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
'                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
'                Optional ByVal MatGracia As Variant, _
'                Optional ByVal pbCuotaFijaFechaFija As Boolean = False, _
'                Optional ByVal pbCuotaComodin As Boolean = False, _
'                Optional ByVal pbDesemParcial As Boolean = False, _
'                Optional ByVal pMatDesPar As Variant = "", _
'                Optional ByVal pnNumMes As Integer = 1, _
'                Optional ByVal pbMiViv As Boolean = False, _
'                Optional ByVal bQuincena As Boolean, _
'                Optional ByVal pbGraciaEnCuotas As Boolean = False, _
'                Optional ByVal pnTasaGracia As Double = 0, _
'                Optional ByVal pnDiaFijo2 As Integer = 0, _
'                Optional pnMontoCapInicial As Double = 0, _
'                Optional ByVal pbPagoInteresEnGracia As Boolean = False, _
'                Optional ByVal psCtaCod As String, _
'                Optional ByVal psMensFecAprob As Date, _
'                Optional ByVal lnDescripCuota As Variant = "", _
'                Optional ByVal pnMontoGra As Double = 0, _
'                Optional ByVal psIncluyeIntCap As String = "S", _
'                Optional ByVal pnCuotaBalon As Integer = 0, _
'                Optional ByVal pbReprogConvenio As Boolean = False, _
'                Optional ByVal pbReduccionCuotaPagAnt As Boolean = False) As Variant  'MAVM 11092010
'                'MADM 20110518 pbTieneGracia - lnDescripCuota
'                'Si pnMontoCapInicial >0 entonces se Capitalizo la Gracia sin aumentar el capital
'                'MAVM 20130209
'                'WIOR 20131111 AGREGO pnCuotaBalon
'                'WIOR 20140514 AGREGO pbReprogConvenio
'                'LUCV20180320, Agregó: pbReduccionCuotaPagAnt
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
'    On Error GoTo ErrorGeneraCalendario
'    If pbCuotaComodin Then
'        TasaComTemp = pnTasaInt
'        If pnTipoPeriodo = PeriodoFijo Then
'            nPlazoCom = pnPeriodo
'        Else
'            nPlazoCom = 30
'        End If
'        nCuotasCom = pnNroCuotas
'        Set oCred = New COMNCredito.NCOMCredito
'        nCuotaMontoCom = oCred.CuotaFija(TasaComTemp, nCuotasCom + 1, pnMonto, nPlazoCom)
'        nCuotaMontoCom = CDbl(Format((nCuotaMontoCom * (pnNroCuotas + 1)) / pnNroCuotas, "#0.00"))
'        TasaComTemp = CDbl(Format(oCred.CreditosTasaEfectiva(nCuotaMontoCom, pnNroCuotas, pnMonto, pnTasaInt, nPlazoCom), "#0.0000"))
'        pnTasaInt = Format(TasaComTemp * 100, "#0.0000")
'        Set oCred = Nothing
'    End If
'
'    '******************************************
'    '******** Gracia repartida en Cuotas ******
'    nCuotasGracia = 0
'    nCuotaIni = 0
'    nCuotaFin = pnNroCuotas - 1
'    ReDim Calendario(pnNroCuotas)
'
'    If pbGraciaEnCuotas And pnTipoPeriodo = PeriodoFijo Then
'        nCuotasGracia = Fix(pnDiasGracia / pnPeriodo)   'Entero inmediato inferior
'        dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy"))
'        ReDim Calendario(pnNroCuotas + nCuotasGracia)
'        Set oCred = New COMNCredito.NCOMCredito
'        For i = 0 To nCuotasGracia - 1
'            Calendario(i).dFecha = IIf(i = nCuotasGracia - 1, CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia, dDesembolso + pnPeriodo)
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
'    '************************************
'
'    '***10-05-2006 *******************************************
'    If pbGraciaEnCuotas And pnTipoPeriodo = FechaFija Then
'        Dim nMes As Integer
'        Dim nAnio As Integer
'        Dim nDia As Integer
'        Dim dFecTemp As Date
'        Dim nDifDiasGracia As Integer
'
'        nCuotasGracia = Fix(pnDiasGracia / 30)   'Entero inmediato inferior
'        nDifDiasGracia = pnDiasGracia Mod 30    'Dias faltantes
'        If nDifDiasGracia = 0 Then
'            dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy"))
'            'nMes = Month(dDesembolso)
'        Else
'            'If bProxMes Then
'                dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + CDate(nDifDiasGracia)
'                'nMes = Month(dDesembolso)
'            '    nMes = Month(dDesembolso) + 1
'            'End If
'        End If
'        nMes = Month(dDesembolso)
'        ReDim Calendario(pnNroCuotas + nCuotasGracia)
'        nAnio = Year(dDesembolso)
'        nDia = pnDiaFijo
'
'        Set oCred = New COMNCredito.NCOMCredito
'        For i = 0 To nCuotasGracia - 1
'            'Calendario(i).dFecha = IIf(i = nCuotasGracia - 1, CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia, dDesembolso + pnPeriodo)
'            If Not (i = 0 And nDia > Day(dDesembolso) And (Not bProxMes)) Or pnNumMes = 6 Then
'                If nDia = pnDiaFijo Then nMes = nMes + pnNumMes
'                    If nMes > 12 Then
'                        nAnio = nAnio + 1
'                        nMes = 1
'                    End If
'                Else
'                    If nDia > 30 Then
'                        If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
'                            nMes = nMes + pnNumMes
'                        End If
'                    End If
'                End If
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
'            dFecTemp = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)))
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
'                Calendario(i).IntGra = oCred.MontoIntPerDias(pnTasaGracia, DateDiff("d", CDate(Calendario(i - 1).dFecha), Calendario(i).dFecha), pnMonto)
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
'        'MAVM 11092010 ***
'        'Call ProcesarCalendario(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoCuota, _
'        '       pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pbCuotaFijaFechaFija, _
'        '       pbDesemParcial, pMatDesPar, pnNumMes, pbMiViv, bQuincena, nCuotaIni, nCuotaFin, pnDiaFijo2, pnMontoCapInicial)
'        Call ProcesarCalendario(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoCuota, _
'                pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pbCuotaFijaFechaFija, _
'                pbDesemParcial, pMatDesPar, pnNumMes, pbMiViv, bQuincena, nCuotaIni, nCuotaFin, _
'                pnDiaFijo2, pnMontoCapInicial, psCtaCod, psMensFecAprob, lnDescripCuota, pnMontoGra, psIncluyeIntCap, _
'                pnCuotaBalon, pbReprogConvenio, pbReduccionCuotaPagAnt)
'                'WIOR 20131111 AGREGO pnCuotaBalon
'                'WIOR 20140514 AGREGO pbReprogConvenio
'                'LUCV20180320, agregó bReduccionCuotaPagAnt Según ERS087-2017
'        '***
'        'MADM 20110531 lnDescripCuota
'       'MAVM 20100320
'       'ReDim sCalendPagos(UBound(Calendario), 8)
'       'ReDim sCalendPagos(UBound(Calendario), 9)
'       'ReDim sCalendPagos(UBound(Calendario), 13)
'       ReDim sCalendPagos(UBound(Calendario), 14)
'
'       For i = 0 To UBound(Calendario) - 1
'            sCalendPagos(i, 0) = Format(Calendario(i).dFecha, "dd/mm/yyyy")
'            sCalendPagos(i, 1) = Trim(Str(Calendario(i).NroCuota))
'            sCalendPagos(i, 2) = Format(Calendario(i).Cuota, "#0.00")
'            sCalendPagos(i, 3) = Format(Calendario(i).Captital, "#0.00")
'            sCalendPagos(i, 4) = Format(Calendario(i).IntComp, "#0.00")
'            'MAVM 20130209 ***
'            'sCalendPagos(i, 5) = Format(Calendario(i).IntGra, "#0.00")
'            If Not (pnTipoGracia = 1 And i = 0) Then
'                sCalendPagos(i, 5) = Format(Calendario(i).IntGra, "#0.00")
'            Else
'                Set oCred = New COMNCredito.NCOMCredito
'                sCalendPagos(i, 5) = Format(Calendario(i).IntGra, "#0.00") + oCred.MontoIntPerDias(CDbl(pnTasaGracia), IIf(pnTipoPeriodo = 1, CInt(pnPeriodo), 30), CDbl(Trim(Calendario(i).IntGra)))
'                sCalendPagos(i, 2) = sCalendPagos(i, 2) + oCred.MontoIntPerDias(CDbl(pnTasaGracia), IIf(pnTipoPeriodo = 1, CInt(pnPeriodo), 30), CDbl(Trim(Calendario(i).IntGra)))
'                Set oCred = Nothing
'            End If
'            '***
'            sCalendPagos(i, 6) = Format(Calendario(i).Gasto, "#0.00")
'            sCalendPagos(i, 7) = Format(Calendario(i).SaldoCap, "#0.00")
'
'            'MAVM 20100320
'            sCalendPagos(i, 8) = Format(Calendario(i).SegDes, "#0.00")
'            sCalendPagos(i, 9) = Format(Calendario(i).SegInm, "#0.00") 'MAVM 20120625
'            'MAVM 20130209 ***
'            sCalendPagos(i, 10) = Format(Calendario(i).CuotaGra, "#0.00")
'            sCalendPagos(i, 11) = Format(Calendario(i).IntCompGra, "#0.00")
'            sCalendPagos(i, 12) = Format(Calendario(i).SaldoCapGra, "#0.00")
'            '***
'            sCalendPagos(i, 14) = Format(0, "#0.00") 'RECO
'        Next i
'        GeneraCalendario = sCalendPagos
'        Exit Function
'
'ErrorGeneraCalendario:
'        'MsgBox Err.Description, vbCritical, "Aviso"
'        Err.Raise Err.Number, "Error", Err.Description
'End Function
'<-*****  Fin LUCV20180601

'->****LUCV20180601, [Reestructuración met. (GeneraCalendario)], según ERS022-2018
Public Function GeneraCalendario(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, ByVal pnPeriodo As Double, _
        ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
        ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
        Optional ByVal MatGracia As Variant = "", Optional ByVal pbCuotaFijaFechaFija As Boolean = False, _
        Optional ByVal pbCuotaComodin As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
        Optional ByVal pMatDesPar As Variant = "", _
        Optional ByVal pnNumMes As Integer = 1, _
        Optional ByVal pbMiViv As Boolean = False, _
        Optional ByVal bQuincena As Boolean, _
        Optional ByVal pbGraciaEnCuotas As Boolean = False, _
        Optional ByVal pnTasaGracia As Double = 0, _
        Optional ByVal pnDiaFijo2 As Integer = 0, _
        Optional ByVal pnMontoCapInicial As Double = 0, _
        Optional ByVal pbPagoInteresEnGracia As Boolean = False, _
        Optional ByVal psCtaCod As String, _
        Optional ByVal psMensFecAprob As Date, _
        Optional ByVal lnDescripCuota As Variant = "", _
        Optional ByVal pnMontoGra As Double = 0, _
        Optional ByVal psIncluyeIntCap As String = "S", _
        Optional ByVal pnCuotaBalon As Integer = 0, _
        Optional ByVal pbReprogConvenio As Boolean = False, _
        Optional ByVal pbReduccionCuotaPagAnt As Boolean = False, _
        Optional ByVal pnTasaSegDes As Double = 0, Optional ByRef pMatCalendSegDes As Variant, _
        Optional ByVal pnExoSeguroDesgravamen As Integer = 0, Optional ByVal pnMontoPoliza As Double = 0, _
        Optional ByVal pnTasaSegInc As Double = 0, Optional ByRef pArrayDatos As Variant = Nothing) As Variant
        'MADM 20110518 pbTieneGracia - lnDescripCuota
        'Si pnMontoCapInicial >0 entonces se Capitalizo la Gracia sin aumentar el capital
        'MAVM 20130209
        'WIOR 20131111 AGREGO pnCuotaBalon
        'WIOR 20140514 AGREGO pbReprogConvenio
        'LUCV20180320, Agregó: pbReduccionCuotaPagAnt
        'LUCV20180601, Agregó: pnTasaSegDes, pMatCalendSegDes,pnExoSeguroDesgravamen, pnMontoPoliza Según ERS022-2018
                                
Dim sCalendPagos() As String
Dim i As Integer
Dim TasaComTemp As Double
Dim nPlazoCom As Integer
Dim nCuotasCom As Integer
Dim nCuotaMontoCom As Double
Dim oCred As COMNCredito.NCOMCredito

'Gracia en Cuotas
'Dim nCuotasGracia As Integer
Dim nCuotaIni As Integer
Dim nCuotaFin As Integer
Dim dDesembolso As Date
        
    On Error GoTo ErrorGeneraCalendario
    nCuotaIni = 0
    nCuotaFin = pnNroCuotas - 1
    ReDim Calendario(pnNroCuotas)
    
    Call ProcesarCalendario(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, _
                            pnTipoCuota, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, _
                            bProxMes, MatGracia, pbCuotaFijaFechaFija, pbDesemParcial, pMatDesPar, _
                            pnNumMes, pbMiViv, bQuincena, nCuotaIni, nCuotaFin, pnDiaFijo2, pnMontoCapInicial, _
                            psCtaCod, psMensFecAprob, lnDescripCuota, pnMontoGra, psIncluyeIntCap, _
                            pnCuotaBalon, pbReprogConvenio, pbReduccionCuotaPagAnt, _
                            pnTasaSegDes, pMatCalendSegDes, pnExoSeguroDesgravamen, pnMontoPoliza, pnTasaSegInc, pArrayDatos)
                            'WIOR 20131111 AGREGO pnCuotaBalon
                            'WIOR 20140514 AGREGO pbReprogConvenio
                            'LUCV20180320, agregó bReduccionCuotaPagAnt Según ERS087-2017
                            'LUCV20180601, Agregó: pnTasaSegDes, pnExoneraSegDes, pMatCalendSegDes, pnMontoPoliza
       
       ReDim sCalendPagos(UBound(Calendario), 22) ' LUCV20180601 Según ERS022-2018, Modificó 14 por 16
       'RIRO20200825 Se cambió indice de 16 a 22
       For i = 0 To UBound(Calendario) - 1
            sCalendPagos(i, 0) = Format(Calendario(i).dFecha, "dd/mm/yyyy")
            sCalendPagos(i, 1) = Trim(Str(Calendario(i).NroCuota))
            sCalendPagos(i, 2) = Format(Calendario(i).Cuota, "#0.00")
            sCalendPagos(i, 3) = Format(Calendario(i).Captital, "#0.00")
            sCalendPagos(i, 4) = Format(Calendario(i).IntComp, "#0.00")
            sCalendPagos(i, 5) = Format(Calendario(i).IntGra, "#0.00") 'LUCV20180601, Según ERS022-2018
            sCalendPagos(i, 6) = Format(Calendario(i).Gasto, "#0.00") 'LUCV20180601, Según ERS022-2018
            sCalendPagos(i, 7) = Format(Calendario(i).SaldoCap, "#0.00")
            sCalendPagos(i, 8) = Format(Calendario(i).SegDes, "#0.00")
            sCalendPagos(i, 9) = Format(Calendario(i).SegInm, "#0.00")
            sCalendPagos(i, 10) = Format(Calendario(i).CuotaGra, "#0.00")
            sCalendPagos(i, 11) = Format(Calendario(i).IntCompGra, "#0.00")
            sCalendPagos(i, 12) = Format(Calendario(i).SaldoCapGra, "#0.00")
            'sCalendPagos(i, 13) = Format(Calendario(i).SegMultMype, "#0.00")
            sCalendPagos(i, 14) = Format(0, "#0.00")
            sCalendPagos(i, 15) = Format(Calendario(i).CuotaPrimaPoliza, "#0.00") 'LUCV20180601, Según ERS022-2018
            sCalendPagos(i, 16) = Format(Calendario(i).CuotaPrimaPolizaGracia, "#0.00") 'LUCV20180601, Según ERS022-2018
            
            'RIRO20200825 Mejora en Liquidación ******
            sCalendPagos(i, 17) = Format(Calendario(i).IntCompCalculado, "#0.00")
            sCalendPagos(i, 18) = Format(Calendario(i).DiasCalculo, "#0.00")
            sCalendPagos(i, 19) = Format(Calendario(i).IntCompDiferenciaCapitalizado, "#0.00")
            sCalendPagos(i, 20) = Format(Calendario(i).IntGraciaGenerado, "#0.00")
            sCalendPagos(i, 21) = Format(Calendario(i).IntGraciaCapitalizado, "#0.00")
            sCalendPagos(i, 22) = Format(Calendario(i).IntGraciaAsignado, "#0.00")
            'RIRO20200825 Mejora en Liquidación ******
            
        Next i
        GeneraCalendario = sCalendPagos
        Exit Function
ErrorGeneraCalendario:
        Err.Raise Err.Number, "Error", Err.Description
End Function
'<-***** Fin LUCV201806011

Public Function ReporteCalendario(ByVal pTipoRep As Byte, ByVal MatP As Variant, ByVal MatMalo As Variant, _
                ByVal pTCuota As String, ByVal pInteres As Double, ByVal pMonto, _
                ByVal pCuota As Double, ByVal pPlazo As Integer, ByVal pVigencia As Date, _
                Optional ByVal pnTipoProceso As Integer = 0, _
                Optional ByVal pMatDesPar As Variant = "", _
                Optional GBITFAPLICAt As Boolean, _
                Optional gnITFPorcentt As Double, _
                Optional gnITFMontoMint As Double = 0, _
                Optional ByVal psCtaCod As String, _
                Optional pnCuotas As Integer, _
                Optional ByVal pnTasaEfectivaAnual As Double, _
                Optional ByVal pnTasaCostoEfectivoAnual As Double, _
                Optional ByVal lsCtaCodLeasing As String = "", _
                Optional ByVal nTipoGracia As Integer, _
                Optional ByVal nIntGraInicial As Double, _
                Optional ByVal nDiasAtraso As Integer, _
                Optional psPersNom As String = "", _
                Optional ByVal ArrSimulador As Variant = Nothing) As String
                'DAOR 20070403, se agregó variable pnTasaEfectivaAnual, pnTasaCostoEfectivoAnual
                'MAVM 20130209
                'JUEZ 20140104 se agregó psPersNom
                'WIOR 20150211 AGREGO Optional ByVal ArrSimulador As Variant
        
    Dim lsCadImp As String
    Dim lsCadBuffer As String
    
    Dim lnIndice As Long
    Dim lnLineas As Integer
    Dim lnPage As Integer
    Dim nItem As Integer
    
    Dim cTitulo As String
    Dim cSubTitulo As String 'FRHU 20140722 ERS105-2014
    Dim i As Integer
    Dim cCuota As String * 35
    
    Dim sCuota As Double '** PEAC 20080723
    Dim sCapital As Double '** PEAC 20080723
    Dim sInteres As Double '** PEAC 20080723
    Dim sGastos As Double '** PEAC 20080723
    Dim sInteresGra As Double 'WIOR 20150211
    
    Dim sSegDes As Double '** MAVM 20100320
    Dim sSegInm As Double '** MAVM 20120825
     
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sNombre As String
    
    '**DAOR 20070331*********************************
    Dim oParametros As COMDCredito.DCOMParametro
    Dim nPorComiCorrespMN As Double
    Dim nPorComiCorrespME As Double
    Dim nMonMinComiCorrespMN As Double
    Dim nMonMaxComiCorrespMN As Double
    Dim nMonMinComiCorrespME As Double
    Dim nFactorCorresp As Double
    Dim nMonMinComi As Double
    Dim nMonMaxComi As Double
    Dim nComiCorresp As Double
    '*************************************************
    oITF.gbITFAplica = GBITFAPLICAt
    oITF.gnITFPorcent = gnITFPorcentt
    oITF.gnITFMontoMin = gnITFMontoMint
     
    cCuota = pTCuota
    sCapital = 0
    sInteres = 0
    'WIOR 20150211**************************
    Dim bSimulador As Boolean
    Dim bEnvioEstFis As Boolean
    Dim nTipoSegDes As Integer
    Dim nSegMYPE As Double 'APRI20180821 ERS061-2018
    bSimulador = False
    bEnvioEstFis = False
    nTipoSegDes = 0
    If IsArray(ArrSimulador) Then
        If ArrSimulador(0) = 1 Then
            bSimulador = True
            nTipoSegDes = ArrSimulador(2)
            If ArrSimulador(1) = 1 Then
                bEnvioEstFis = True
            End If
        End If
    End If
    'WIOR FIN ******************************
    
    'WIOR 20160121 ***
    Dim oCred As COMDCredito.DCOMCredito
    Dim rsCred As ADODB.Recordset
    Dim TpoProd As String
    Dim TpoCred As String
    Dim bEsMIVIVIENDAAnt As Boolean
    bEsMIVIVIENDAAnt = False
    
    If Not bSimulador Then
        If Len(Trim(psCtaCod)) >= 18 Then
            Set oCred = New COMDCredito.DCOMCredito
            Set rsCred = oCred.RecuperaColocaciones(psCtaCod)
            TpoProd = IIf(IsNull(rsCred!cTpoProdCod), "", rsCred!cTpoProdCod)
            TpoCred = IIf(IsNull(rsCred!cTpoCredCod), "", rsCred!cTpoCredCod)
            bEsMIVIVIENDAAnt = oCred.EsCredMIVIVENDA(TpoProd, TpoCred, 1)
            Set oCred = Nothing
        End If
    End If
    'WIOR FIN ********

    cSubTitulo = "" 'FRHU 20140722 ERS105-2014
    '**DAOR 20070331*************************************************
    nFactorCorresp = 1
    nMonMinComi = 0
    nMonMaxComi = 0
    If psCtaCod <> "" Then
        Set oParametros = New COMDCredito.DCOMParametro
        nPorComiCorrespMN = oParametros.RecuperaValorParametro(gColPParamPorComiCorresponsaliaMN)
        nPorComiCorrespME = oParametros.RecuperaValorParametro(gColPParamPorComiCorresponsaliaME)
        nMonMinComiCorrespMN = oParametros.RecuperaValorParametro(gColPParamMonMinCobroCorresponsaliaMN)
        nMonMaxComiCorrespMN = oParametros.RecuperaValorParametro(gColPParamMonMaxCobroCorresponsaliaMN)
        nMonMinComiCorrespME = oParametros.RecuperaValorParametro(gColPParamMonMinCobroCorresponsaliaME)
        If Mid$(psCtaCod, 9, 1) = "1" Then 'MN
            nFactorCorresp = nPorComiCorrespMN / 100
            nMonMinComi = nMonMinComiCorrespMN
            nMonMaxComi = nMonMaxComiCorrespMN
        Else
            nFactorCorresp = nPorComiCorrespME / 100
            nMonMinComi = nMonMinComiCorrespME
            nMonMaxComi = 9999999999#
        End If
    End If
    '***************************************************************
    
    If pTipoRep = 1 Then
        If pnTipoProceso = 0 Then
            cTitulo = "SIMULACION DE CALENDARIO DE PAGOS"
            cSubTitulo = "SÓLO PARA USO DE CARÁCTER INFORMATIVO" 'FRHU 20140722 ERS105-2014
        ElseIf pnTipoProceso = 1 Then
            'cTitulo = "SUGERENCIA DE ANALISTA - CALENDARIO DE PAGOS" 'LUCV20180601, Comentó según ERS022-2018
            cTitulo = "SUGERENCIA DE ANALISTA - CRONOGRAMA DE PAGOS" 'LUCV20180601 - según ERS022-2018
        ElseIf pnTipoProceso = 2 Then
            'cTitulo = "APROBACION DE CREDITO - CALENDARIO DE PAGOS" 'LUCV20180601- Comentó según ERS022-2018
            cTitulo = "APROBACION DE CREDITO - CRONOGRAMA DE PAGOS" 'LUCV20180601 - según ERS022-2018
        End If
    ElseIf pTipoRep = 2 Then
        If pnTipoProceso = 0 Then
            cTitulo = "SIMULACION DE CALENDARIO DE PAGOS MI VIVIENDA"
            cSubTitulo = "SÓLO PARA USO DE CARÁCTER INFORMATIVO" 'FRHU 20140722 ERS105-2014
        ElseIf pnTipoProceso = 1 Then
            'cTitulo = "SUGERENCIA DE ANALISTA - CALENDARIO DE PAGOS MI VIVIENDA" 'LUCV20180601, Comentó según ERS022-2018
            cTitulo = "SUGERENCIA DE ANALISTA - CRONOGRAMA DE PAGOS MI VIVIENDA" 'LUCV20180601 - según ERS022-2018
        ElseIf pnTipoProceso = 2 Then
            'cTitulo = "APROBACION DE CREDITO - CALENDARIO DE PAGOS MI VIVIENDA" 'LUCV20180601, Comentó según ERS022-2018
            cTitulo = "APROBACION DE CREDITO - CRONOGRAMA DE PAGOS MI VIVIENDA" 'LUCV20180601 - según ERS022-2018
        End If
    End If

    lnLineas = 0
    lnPage = 1

    'Obteniendo titular
    sSql = "Select Pers.cPersNombre"
    sSql = sSql & " From ProductoPersona PP"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Where PP.cCtaCod='" & psCtaCod & "' And PP.nPrdPersRelac=20"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        sNombre = rs!cPersNombre
    End If
    Set rs = Nothing
    
    'CABECERA
    If pTipoRep = 1 Then
        lsCadImp = lsCadImp & nRepoCabecera(cTitulo, cSubTitulo, lnPage, 116, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, IIf(psPersNom <> "", "Rapiflash con tu Plazo Fijo", psCtaCod), IIf(Trim(sNombre) = "", psPersNom, Trim(sNombre)), pnCuotas, pnTasaEfectivaAnual, pnTasaCostoEfectivoAnual, nTipoGracia, nDiasAtraso) 'APRI20180821 ERS061-2018
    ElseIf pTipoRep = 2 Then
        'lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "BUEN PAGADOR", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, psCtaCod, Trim(sNombre), pnCuotas)
        lsCadImp = lsCadImp & nRepoCabecera(cTitulo, cSubTitulo, lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, psCtaCod, Trim(sNombre), pnCuotas) 'FRHU 20140722 ERS105-2014
    End If
        
    If IsArray(pMatDesPar) Then
        lsCadImp = lsCadImp & Space(5) & "DESEMBOLSOS PARCIALES" & oImp.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(5) & String(63, "-") & oImp.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(5) & oImpre.ImpreFormat("FECHA", 16, 0) & oImpre.ImpreFormat("MONTO", 10, 0) & oImp.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(5) & String(63, "-") & oImp.gPrnSaltoLinea
        For i = 0 To UBound(pMatDesPar) - 1
            lsCadImp = lsCadImp & Space(5) & oImpre.ImpreFormat(pMatDesPar(i, 0), 10, 0) & oImpre.ImpreFormat(CDbl(pMatDesPar(i, 1)), 10, 2, True) & oImp.gPrnSaltoLinea
        Next i
        lsCadImp = lsCadImp & Space(5) & String(63, "-") & oImp.gPrnSaltoLinea
    End If
    
    If bSimulador Then
        If bEnvioEstFis Then
            lsCadImp = lsCadImp & " FECHA        NR.        CUOTA      CAPITAL      INTERES    INT.GRACIA   GAST/COM(2)  SEG. DES(1)  SALDO CAP." & oImp.gPrnSaltoLinea
        Else
            lsCadImp = lsCadImp & " FECHA        NR.        CUOTA      CAPITAL      INTERES    INT.GRACIA   GAST/COM     SEG. DES(1)  SALDO CAP." & oImp.gPrnSaltoLinea
        End If
    Else
    'WIOR FIN *************************
        If pTipoRep = 1 Then
            'lsCadImp = lsCadImp & " FECHA        NR.        CUOTA      CAPITAL      INTERES     GAST/COM     SEG. DES   SEG. MYPE   SALDO CAP." & oImp.gPrnSaltoLinea 'APRI20180821 ERS061-2018 'LUCV20180601, Comentó. Según ERS022-2018
            lsCadImp = lsCadImp & " FECHA        NR.        CUOTA      CAPITAL      INTERES     GAST/COM     SEG. DES   SEG. MULT.   SALDO CAP." & oImp.gPrnSaltoLinea 'APRI20180821 ERS061-2018 'LUCV20180601. Según ERS022-2018
        ElseIf pTipoRep = 2 Then
            lsCadImp = lsCadImp & " FECHA        NR.        CUOTA      CAPITAL      INTERES     GAST/COM     SEG. DES  SEG. INM    SALDO CAP." & oImp.gPrnSaltoLinea
        End If
    End If 'WIOR 20150211
    
    lsCadImp = lsCadImp & "              CUOTA                                                                        " & oImp.gPrnSaltoLinea
    
    'WIOR 20150211 ****************
    If bSimulador Then
        lsCadImp = lsCadImp & String(106, "-") & oImp.gPrnSaltoLinea
    Else
    'WIOR FIN *********************
        lsCadImp = lsCadImp & String(106, "-") & oImp.gPrnSaltoLinea 'APRI20180821 ERS061-2018 CHANGE 100 TO 106
    End If 'WIOR 20150211
    
    lnIndice = 10:  lnLineas = 10
    nItem = 0
    
    'Si la gracia es capitalizada
    If nTipoGracia = 6 And i = 0 Then
        lsCadImp = lsCadImp & " " & oImpre.ImpreFormat(Format(CDate(Format(pVigencia + nDiasAtraso, "dd/MM/yyyy")), "dd/mm/yy"), 10, 0) & Space(2) '** FECHA
        lsCadImp = lsCadImp & oImpre.ImpreFormat(0, 4, 0) & Space(15)
        lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(pMonto), 8, 2) & IIf(bSimulador, Space(13), Space(2)) 'WIOR 20150211 AGREGO IIf(bSimulador
        lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(nIntGraInicial), 8, 2) & IIf(bSimulador, Space(28), Space(28))   'WIOR 20150211 AGREGO IIf(bSimulador
        lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(pMonto), 8, 2)
        lsCadImp = lsCadImp & oImp.gPrnSaltoLinea
    End If
    
    'Recorrido del detalle del calendario
    For i = 0 To UBound(MatP) - 1
        nItem = nItem + 1
        lnLineas = lnLineas + 1
        lnIndice = lnIndice + 1
        lsCadImp = lsCadImp & " " & oImpre.ImpreFormat(Format(CDate(Format(MatP(i, 0), "dd/MM/yyyy")), "dd/mm/yy"), 10, 0) & Space(2) '** FECHA
        lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 1)), 4, 0) & Space(2) '** NRO CUOTA
        lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 2)), 8, 2) & Space(2) '** CUOTA
        sCuota = sCuota + MatP(i, 2)
        lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 3)), 8, 2) & Space(2) '** CAPITAL
        sCapital = sCapital + MatP(i, 3)
        
        If bSimulador Then
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 4)), 8, 2)  '** INTERES
            sInteres = sInteres + MatP(i, 4)
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 5)), 8, 2)  '** INT.GRACIA
            sInteresGra = sInteresGra + MatP(i, 5)
        Else
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 4)) + Val(MatP(i, 5)), 8, 2) & Space(2) '** INTERES (Compensatorio + Gracia)
            sInteres = sInteres + MatP(i, 4) + MatP(i, 5)
        End If
        
        'ALPA 20110528****************LEASING
        If Len(Trim(lsCtaCodLeasing)) = 0 Then
            'lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 8)), 8, 2) & Space(2) '** GASTOS 'LUCV20180601. Comentó
            'sGastos = sGastos + MatP(i, 8) 'LUCV20180601. Comentó
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 6)), 8, 2) & Space(2) '** GASTOS 'LUCV20180601. Según ERS022-2018
            sGastos = sGastos + MatP(i, 6) 'LUCV20180601. Según ERS022-2018
            
            'lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 6)), 8, 2) & Space(1) '** GASTOS SEG DES 'LUCV20180601. Comentó
            'sSegDes = sSegDes + MatP(i, 6) 'LUCV20180601. Comentó
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 8)), 8, 2) & Space(1) '** GASTOS SEG DES 'LUCV20180601. Según ERS022-2018
            sSegDes = sSegDes + MatP(i, 8) 'LUCV20180601. Según ERS022-2018
            
            'APRI20180821 ERS061-2018
            'lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 14)), 8, 2) & Space(1) '** GASTOS SEG MYPE ''LUCV20180601, Según ERS022-2018
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 14)) + (Val(MatP(i, 15)) + Val(MatP(i, 16))), 8, 2) & Space(1) '** GASTOS SEG MYPE 'LUCV20180601, Según ERS022-2018
            nSegMYPE = nSegMYPE + MatP(i, 14)
            'END APRI
        Else
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 15)) + Val(MatP(i, 16)), 8, 2) & Space(2) '** GASTOS
            sGastos = sGastos + MatP(i, 8)
            
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 8)), 8, 2) & Space(2) '** GASTOS SEG DES
            sSegDes = sSegDes + MatP(i, 6)
        End If
        '*************************************
        
        If pTipoRep = 2 Then 'Mi Vivienda
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 9)), 8, 2) & Space(1) '** GASTOS SEG INM
            sSegInm = sSegInm + MatP(i, 9)
        End If
        
        'WIOR 20150211**************************
        If bSimulador Then
            lsCadImp = lsCadImp & Space(2)
        End If
        'WIOR FIN ******************************
        
        lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 7)), 8, 2) & Space(2) & oImp.gPrnSaltoLinea '** SALDO K
        
        If lnIndice Mod 300 = 0 Then
            lsCadBuffer = lsCadBuffer & lsCadImp
            lsCadImp = ""
        End If
        
        If lnLineas >= 50 Then
            lnPage = lnPage + 1
            lsCadImp = lsCadImp & oImp.gPrnSaltoPagina
            lsCadImp = lsCadImp & oImp.gPrnSaltoLinea
            If pTipoRep = 1 Then
                lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "PLAN DE PAGOS", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, psCtaCod, Trim(sNombre), pnCuotas, pnTasaEfectivaAnual, pnTasaCostoEfectivoAnual) 'DAOR 20070403
            ElseIf pTipoRep = 2 Then
                lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "BUEN PAGADOR", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, psCtaCod, Trim(sNombre), pnCuotas)
            End If
            
            lnLineas = 8
            lnIndice = lnIndice + 8
        End If
    Next
    
    'LUCV20180601, Comentó y Agregó. Según ERS022-2018
    'WIOR 20150211 ****************
'    If bSimulador Then
'        lsCadImp = lsCadImp & String(106, "=") & oImp.gPrnSaltoLinea
'    Else
'    'WIOR FIN *********************
'        lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & String(106, "=") & oImp.gPrnSaltoLinea 'APRI20180821 ERS061-2018 CHANGE 100 TO 106
'    End If 'WIOR 20150211
'
'    If bSimulador Then
'        lsCadImp = lsCadImp & "Totales :" & Space(10) & oImpre.ImpreFormat(sCuota, 8, 2) & Space(2) & oImpre.ImpreFormat(sCapital, 8, 2) & Space(2) & oImpre.ImpreFormat(sInteres, 8, 2) & oImpre.ImpreFormat(sInteresGra, 8, 2) & oImpre.ImpreFormat(sGastos, 8, 2) & Space(4) & oImpre.ImpreFormat(sSegDes, 6, 2) & oImp.gPrnSaltoLinea
'    Else
'    'WIOR FIN *********************
'        If pTipoRep = 1 Then
'            lsCadImp = lsCadImp & "Totales :" & Space(10) & oImpre.ImpreFormat(sCuota, 8, 2) & Space(2) & oImpre.ImpreFormat(sCapital, 8, 2) & Space(2) & oImpre.ImpreFormat(sInteres, 8, 2) & Space(2) & oImpre.ImpreFormat(sGastos, 8, 2) & Space(4) & oImpre.ImpreFormat(sSegDes, 6, 2) & Space(3) & oImpre.ImpreFormat(nSegMYPE, 6, 2) & oImp.gPrnSaltoLinea 'APRI20180821 ERS061-2018
'        ElseIf pTipoRep = 2 Then
'            lsCadImp = lsCadImp & "Totales :" & Space(10) & oImpre.ImpreFormat(sCuota, 8, 2) & Space(2) & oImpre.ImpreFormat(sCapital, 8, 2) & Space(2) & oImpre.ImpreFormat(sInteres, 8, 2) & Space(2) & oImpre.ImpreFormat(sGastos, 8, 2) & Space(4) & oImpre.ImpreFormat(sSegDes, 6, 2) & Space(1) & oImpre.ImpreFormat(sSegInm, 9, 2) & oImp.gPrnSaltoLinea
'        End If
'    End If 'WIOR 20150211
'
'    'WIOR 20150211 ****************
'    If bSimulador Then
'        lsCadImp = lsCadImp & String(106, "=") & oImp.gPrnSaltoLinea
'    Else
'    'WIOR FIN *********************
'        lsCadImp = lsCadImp & String(106, "=") & oImp.gPrnSaltoLinea 'APRI20180821 ERS061-2018 CHANGE 100 TO 106
'    End If 'WIOR 20150211

   If bSimulador Then
        lsCadImp = lsCadImp & String(106, "=") & oImp.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Totales :" & Space(10) & oImpre.ImpreFormat(sCuota, 8, 2) & Space(2) & oImpre.ImpreFormat(sCapital, 8, 2) & Space(2) & oImpre.ImpreFormat(sInteres, 8, 2) & oImpre.ImpreFormat(sInteresGra, 8, 2) & oImpre.ImpreFormat(sGastos, 8, 2) & Space(4) & oImpre.ImpreFormat(sSegDes, 6, 2) & oImp.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(106, "=") & oImp.gPrnSaltoLinea
    Else
    'WIOR FIN *********************
        lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & String(106, "=") & oImp.gPrnSaltoLinea 'APRI20180821 ERS061-2018 CHANGE 100 TO 106
        If pTipoRep = 1 Then
            lsCadImp = lsCadImp & "Totales :" & Space(10) & oImpre.ImpreFormat(sCuota, 8, 2) & Space(2) & oImpre.ImpreFormat(sCapital, 8, 2) & Space(2) & oImpre.ImpreFormat(sInteres, 8, 2) & Space(2) & oImpre.ImpreFormat(sGastos, 8, 2) & Space(4) & oImpre.ImpreFormat(sSegDes, 6, 2) & Space(3) & oImpre.ImpreFormat(nSegMYPE, 6, 2) & oImp.gPrnSaltoLinea 'APRI20180821 ERS061-2018
        ElseIf pTipoRep = 2 Then
            lsCadImp = lsCadImp & "Totales :" & Space(10) & oImpre.ImpreFormat(sCuota, 8, 2) & Space(2) & oImpre.ImpreFormat(sCapital, 8, 2) & Space(2) & oImpre.ImpreFormat(sInteres, 8, 2) & Space(2) & oImpre.ImpreFormat(sGastos, 8, 2) & Space(4) & oImpre.ImpreFormat(sSegDes, 6, 2) & Space(1) & oImpre.ImpreFormat(sSegInm, 9, 2) & oImp.gPrnSaltoLinea
        End If
        lsCadImp = lsCadImp & String(106, "=") & oImp.gPrnSaltoLinea 'APRI20180821 ERS061-2018 CHANGE 100 TO 106
    End If 'WIOR 20150211
    'Fin LUCV20180601
        
    'Mal Pagador Mi Vivienda
    sCapital = 0
    sInteres = 0
    sSegDes = 0
    sCuota = 0
    sSegInm = 0
     
    If pTipoRep = 2 Then
        If bEsMIVIVIENDAAnt Then 'WIOR 20160121
            lsCadImp = lsCadImp & oImp.gPrnSaltoPagina
            lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "MAL PAGADOR", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia)
         
            lsCadImp = lsCadImp & oImp.gPrnSaltoLinea
            
            lnIndice = 10:  lnLineas = 10
            nItem = 0
            For i = 0 To UBound(MatMalo) - 1
                nItem = nItem + 1
                lnLineas = lnLineas + 1
                lnIndice = lnIndice + 1
                
                'MAVM 20120828
                'lsCadImp = lsCadImp & oImpre.ImpreFormat(nItem, 4, 0) & Space(2)
                'lsCadImp = lsCadImp & oImpre.ImpreFormat(Format(MatMalo(i, 0), "ddd, dd mmm yyyy"), 16, 0) & Space(2)
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Format(MatMalo(i, 0), "dd/MM/yyyy"), 10, 0) & Space(2) 'Fecha
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 1)), 4, 0) & Space(2) 'Nro Cuota
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 2)), 8, 2) & Space(2) 'Cuota
                sCuota = sCuota + MatMalo(i, 2) 'MAVM 20120828
                
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 3)), 8, 2) & Space(2)
                sCapital = sCapital + MatMalo(i, 3)
                
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 4)) + Val(MatMalo(i, 5)), 8, 2) & Space(2) 'Interes
                sInteres = sInteres + MatMalo(i, 4) + MatMalo(i, 5)
                'lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 5)), 8, 2) & Space(2)
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 8)), 8, 2) & Space(2) 'Gastos
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 6)), 8, 2) & Space(2) 'Seg Des
                sSegDes = sSegDes + MatMalo(i, 6)
                
                'MAVM 20120825 ***
                If pTipoRep = 2 Then
                    lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 9)), 8, 2) & Space(2) ' Seg Inm
                    sSegInm = sSegInm + MatMalo(i, 9)
                End If
                '***
                
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 7)), 8, 2) & Space(2) & oImp.gPrnSaltoLinea 'Saldo Cap
                'lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 8)), 8, 2) & Space(2)
                  
                'lnLineas = lnLineas + 1
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oImp.gPrnSaltoPagina
                    lsCadImp = lsCadImp & oImp.gPrnSaltoLinea
                    lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "MAL PAGADOR", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia)
                 
                    lnLineas = 8
                    lnIndice = lnIndice + 8
                End If
                        
            Next
            
            lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & String(100, "=") & oImp.gPrnSaltoLinea
            
            '*** PEAC 20080723
    '        lsCadImp = lsCadImp & "Totales :" & Space(37) & oImpre.ImpreFormat(sCapital, 8, 2)
    '        lsCadImp = lsCadImp & Space(2) & oImpre.ImpreFormat(sInteres, 8, 2)
    '        lsCadImp = lsCadImp & Space(28) & oImpre.ImpreFormat(sCapital + sInteres, 8, 2)
             lsCadImp = lsCadImp & "Totales :" & Space(10) & oImpre.ImpreFormat(sCuota, 8, 2) & Space(2) & oImpre.ImpreFormat(sCapital, 8, 2) & Space(2) & oImpre.ImpreFormat(sInteres, 8, 2) & Space(2) & oImpre.ImpreFormat(sGastos, 8, 2) & Space(2) & oImpre.ImpreFormat(sSegDes, 6, 2) & Space(2) & oImpre.ImpreFormat(sSegInm, 9, 2)
         End If 'WIOR 20160121
    End If
    
    'WIOR 20150211 ************************
    If bSimulador Then
        If nTipoSegDes > 0 Then
            lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & "(1) El seguro desgravamen corresponde a la tasa de " & nTipoSegDes & " titular" & IIf(nTipoSegDes > 1, "(es).", ".")
        End If
        
        If bEnvioEstFis Then
            lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & "(2) Comisión correspondiente al envío de estado de cuenta físico."
        End If
    End If
    'WIOR FIN *****************************
ReporteCalendario = lsCadBuffer & lsCadImp
 End Function

Public Function nRepoCabecera(ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        ByVal pnCodReporte As String, ByVal pTCuota As String, ByVal pInteres As Double, ByVal pMonto As Double, _
        ByVal pCuota As Double, ByVal pPlazo As Integer, ByVal pVigencia As Date, _
        Optional ByVal psCtaCod As String, Optional ByVal psNomCli As String, Optional ByVal pnNroCuotas As Integer, _
        Optional ByVal pnTasaEfectivaAnual As Double, Optional ByVal pnTasaCostoEfectivoAnual As Double, _
        Optional ByVal nTipoGracia As Integer, Optional ByVal pnPerGra As Integer = 0) As String 'DAOR 20070403
        'MAVM 20130209
        'WIOR 20150211 Optional ByVal pnPerGra As Integer
        
Dim lsCadImp As String
'Dim loImprimeCab As NColPImpre

'    Set loImprimeCab = New NColPImpre
'        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
        'psSubTitulo = "" 'FRHU 20140722 ERS105-2014
        lsCadImp = nImprimeCabeceraReportes(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
'    Set loImprimeCab = Nothing
    lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & String(pnAnchoLinea - 10, "-") & oImp.gPrnSaltoLinea
    Select Case pnCodReporte
        Case "CalenPagos"
            lsCadImp = lsCadImp & "    CREDITO         : " & psCtaCod & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    CLIENTE         : " & psNomCli & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    TIPO DE CUOTA   : " & pTCuota & Space(16) & "Nro Cuotas: " & Trim(oImpre.ImpreFormat(pnNroCuotas, 3, 0)) & oImp.gPrnSaltoLinea
            'MAVM 20120124 ***
            'lsCadImp = lsCadImp & "    TASA EFECT.MEN. : " & oImpre.ImpreFormat(pInteres, 8, 4) & Space(42) & "PLAZO :    " & oImpre.ImpreFormat(pPlazo, 3, 0) & oImp.gPrnSaltoLinea
            'lsCadImp = lsCadImp & "    MONTO           : " & oImpre.ImpreFormat(pMonto, 8, 2) & Space(42) & "VIGENCIA: " & Format(pVigencia, "ddd, dd mmm yyyy") & oImp.gPrnSaltoLinea
            'lsCadImp = lsCadImp & "    PLAZO           : " & Trim(oImpre.ImpreFormat(pPlazo, 3, 0)) & oImp.gPrnSaltoLinea 'LUCV20180601, Comentó Según ERS022-2018
            lsCadImp = lsCadImp & "    FRECUENCIA      : " & Trim(oImpre.ImpreFormat(pPlazo, 3, 0)) & oImp.gPrnSaltoLinea 'LUCV20180601, Según ERS022-2018
            '***
            '->***** LUCV20180601, Comentó según ERS022-2018
'            'WIOR 20150211******
'            If pnPerGra > 0 Then
'                lsCadImp = lsCadImp & "    DÍAS DE GRACIA  : " & Trim(oImpre.ImpreFormat(pnPerGra, 3, 0)) & oImp.gPrnSaltoLinea
'            End If
'            'WIOR FIN ***********
            '<-***** Fin LUCV20180601
            
            lsCadImp = lsCadImp & "    MONTO           : " & Trim(oImpre.ImpreFormat(pMonto, 8, 2)) & Space(45) & "VIGENCIA  : " & Format(pVigencia, "ddd, dd mmm yyyy") & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    MONEDA          : " & IIf(Mid(psCtaCod, 9, 1) = 1, "SOLES", IIf(Mid(psCtaCod, 9, 1) = 2, "DOLARES", " ")) & oImp.gPrnSaltoLinea
            If pnTasaEfectivaAnual > 0 Then lsCadImp = lsCadImp & "    TASA EFECTIVA ANUAL       : " & oImpre.ImpreFormat(pnTasaEfectivaAnual, 3, 2) & " %" & oImp.gPrnSaltoLinea
            'If pnTasaEfectivaAnual > 0 Then lsCadImp = lsCadImp & "    TASA COSTO EFECTIVO ANUAL : " & oImpre.ImpreFormat(pnTasaCostoEfectivoAnual, 3, 2) & " %" & oImp.gPrnSaltoLinea
            '*** PEAC 20080818
            If pnTasaCostoEfectivoAnual > 0 Then lsCadImp = lsCadImp & "    TASA COSTO EFECTIVO ANUAL : " & oImpre.ImpreFormat(pnTasaCostoEfectivoAnual, 3, 2) & " %" & oImp.gPrnSaltoLinea
            
            lsCadImp = lsCadImp & String(pnAnchoLinea - 10, "-") & oImp.gPrnSaltoLinea
            
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
  'lsCabe01 = oImpre.FillText(Trim(UCase(psNomCmac)), 55, " ")
  lsCabe01 = oImpre.FillText(Trim(UCase(psNomCmac)), 45, " ")
  
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 45 - 25)
  
  lsCabe01 = lsCabe01 & "Pag.  : " & Str(pnPagina) & "  -  " & psCodUser & oImp.gPrnSaltoLinea
  'lsCabe01 = lsCabe01 & IIf(pbCiereDia = True, IIf(VerifSiCierreDia(), "DC", "AC"), "") & chr$(10)
  ' Cabecera 2
  
  'lsCabe01 = lsCabe01 & oImpre.FillText(Trim(UCase(psNomAgencia)), 35, " ")
  lsCabe01 = lsCabe01 & oImpre.FillText(Trim(UCase(psNomAgencia)), 25, " ")
  
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 25 - 25)
  
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
    
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta

    Dim rs_Pagos As ADODB.Recordset
    
    Dim rs_Temp As ADODB.Recordset
    
    Dim nCuotas As Integer
    Dim nPeriodoGracia As Integer
    On Error GoTo ErrorDesembolsarCredito
    
            sSql = "Select PTI.nTasaInteres,CE.nCuotas,CE.nPlazo,CE.nColocCalendCod,CE.nTipoGracia,CE.nPeriodoFechaFija,CE.nProxMes"
            sSql = sSql & " From ProductoTasaInteres PTI"
            sSql = sSql & " Inner Join ColocacEstado CE on PTI.cCtaCod=CE.cCtaCod and CE.nPrdEstado=2002"
            sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=CE.cCtaCod"
            sSql = sSql & " Where PTI.nPrdTasaInteres=1 and PTI.cCtaCod='" & psCtaCod & "'"
            
            Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
            Set rs = oConec.CargaRecordSet(sSql)
            oConec.CierraConexion
            Set oConec = Nothing
    
    If rs!nColocCalendCod <> gColocCalendCodCL Then
                ' se verifica la fecha de desembolso sea diferente la fecha de aprobacion
            ' obteniendo la  fecha de desembolso
            sSql = "Select dPrdEstado,nMonto"
            sSql = sSql & " From ColocacEstado"
            sSql = sSql & " Where cCtacod='" & psCtaCod & "' and nPrdEstado=2002"
            
            Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
            Set rs = oConec.CargaRecordSet(sSql)
            oConec.CierraConexion
            Set oConec = Nothing
            
            If Not rs.EOF And Not rs.BOF Then
               dFecha = Format(rs!dPrdEstado, "dd/MM/yyyy")
               nMonto = rs!nMonto
            End If
            Set rs = Nothing
            
            ' Obteniendo la tasa de interes
            sSql = "Select PTI.nTasaInteres,CE.nCuotas,CE.nPlazo,CE.nColocCalendCod,CE.nTipoGracia,CE.nPeriodoFechaFija,CE.nProxMes,"
            sSql = sSql & " CE.nPeriodoGracia"
            sSql = sSql & " From ProductoTasaInteres PTI"
            sSql = sSql & " Inner Join ColocacEstado CE on PTI.cCtaCod=CE.cCtaCod and CE.nPrdEstado=2002"
            sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=CE.cCtaCod"
            sSql = sSql & " Where PTI.nPrdTasaInteres=1 and PTI.cCtaCod='" & psCtaCod & "'"
            
            Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
            Set rs = oConec.CargaRecordSet(sSql)
            oConec.CierraConexion
            Set oConec = Nothing
            
 '           Set oCalend = New NCalendario
            nPeriodoGracia = IIf(IsNull(rs!nPeriodoGracia), 0, rs!nPeriodoGracia)
            If nPeriodoGracia > 0 Then
            '    MatGracia = oCalend.GeneraGracia(Rs!nTipoGracia, CDbl(Format(TasaIntPerDias(Rs!nTasaInteres, nPeriodoGracia) * nMonto, "#0.00")), Rs!nCuotas)
                MatGracia = GeneraGracia(rs!nTipoGracia, CDbl(Format(TasaIntPerDias(rs!nTasaInteres, nPeriodoGracia) * nMonto, "#0.00")), rs!nCuotas)
            End If
            
            sSql = "Select nCuotas"
            sSql = sSql & " From ColocacEstado"
            sSql = sSql & " Where cCtaCod='" & psCtaCod & "' and nPrdEstado=2002"
            
            Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
            Set rs_Temp = oConec.CargaRecordSet(sSql)
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
                                
        pMatCalend_1 = GeneraCalendario(pnTramoNoConsMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, _
                        pnTipoCuota, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, _
                        bProxMes, MatGracia, True, , , , , pbMiViv, , pbGraciaCusco, pnTasaGracia)

        'pMatCalend_2 = GeneraCalendario(pnTramoConsMonto, pnTasaInt, pnNroCuotas / 6, 180, pdFecDesemb, _
        '                 Fija, FechaFija, Exonerada, 0, pnDiaFijo, _
        '                bProxMes, MatGracia, True, , , , 6, , , pbGraciaCusco, pnTasaGracia)
                        
        pMatCalend_2 = GeneraCalendario(pnTramoConsMonto, 0.534, pnNroCuotas / 6, 180, pdFecDesemb, _
                         Fija, FechaFija, Exonerada, pnDiasGracia, pnDiaFijo, _
                        bProxMes, MatGracia, True, , , , 6, , , pbGraciaCusco, pnTasaGracia)
        
        Exit Function

ErrorGeneraDobleCalendario:
        'MsgBox Err.Description, vbCritical, "Aviso"
        Err.Raise Err.Number, "Error", Err.Description
End Function

Public Sub IniciaImpresora(Optional ByVal nImpresora As COMFunciones.Impresoras = gEPSON)
Set oImp = New COMFunciones.FCOMVarImpresion
oImp.Inicia nImpresora
End Sub

'ARCV

Public Function GenerarCalendarioPagos(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Integer, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnTasaGracia As Double, ByVal pnDiaFijo As Integer, _
                ByVal bProxMes As Boolean, ByVal pMatGracia As Variant, ByVal pnMiViv As Integer, _
                ByVal pnCuotaCom As Integer, ByRef MatMiViv_2 As Variant, Optional ByVal pnSugerAprob As Integer = 0, _
                Optional ByVal pbNoMostrarCalendario As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
                Optional ByVal pMatDesPar As Variant = "", Optional ByVal bQuincenal As Boolean, Optional ByVal psCtaCod As String, _
                Optional ByRef pbErrorValidacion As Boolean = False, Optional ByVal pnDiaFijo2 As Integer = 0, _
                Optional ByVal pbIncrementaCapital As Boolean = False, Optional ByVal pdInicioPago As Date = 0, _
                Optional ByVal gnITFMontoMin As Double = 0, Optional ByVal gnITFPorcent As Double = 0, Optional ByVal gbITFAplica As Double = 0, _
                Optional ByVal pnMontoGra As Double = 0, Optional ByVal pnCuotaBalon As Integer = 0, _
                Optional ByVal pbReprogConvenio As Boolean = False, Optional pnMontoMivivienda As Currency = 0#) As Variant
                'MAVM 20120309: pnCapInicial, pnIntGraInicial
                'WIOR 20131111 AGREGO pnCuotaBalon
                'WIOR 20140514 AGREGO pbReprogConvenio
Dim i As Integer
Dim nTipoCuota As Integer
Dim nTipoPeriodo As Integer
Dim nTotalInteres As Double
Dim nTotalCapital As Double
Dim oParam As COMDCredito.DCOMParametro
Dim nTramoNoConsMonto As Double
Dim nTramoConsMonto As Double
Dim nTramoNoConsPorcen As Double
Dim nTramoConsTasa As Double
Dim nPlazoMiViv As Integer
Dim nTotalcuotasCONItF As Double
Dim boptTipoGracia_0 As Boolean
Dim boptTipoGracia_1 As Boolean
                
Dim bDesemParcial As Boolean
Dim MatDesPar As Variant
Dim cCtaCodG As String
Dim nSugerAprob As Integer
Dim nTipoGracia As Integer
Dim MatGracia As Variant
Dim MatResul As Variant
Dim MatResulDiff As Variant
Dim MatCalend As Variant
Dim MatCalend_2 As Variant
Dim nMontoCapInicial As Double
                
    bDesemParcial = pbDesemParcial
    MatDesPar = pMatDesPar
    cCtaCodG = psCtaCod
    nSugerAprob = pnSugerAprob
    nTipoGracia = pnTipoGracia
    Set MatGracia = Nothing
    MatGracia = pMatGracia
    
    'Cambios para las opciones de gracia
    If pnTipoGracia = EnCuotas - 1 Then
        boptTipoGracia_1 = True
    End If
    If pnTipoGracia = Capitalizada - 1 Then
        boptTipoGracia_0 = True
    End If
    '***********************************
    
    '*********************************************************************************
    '****  APLICAION DE CALENDARIO
    '*********************************************************************************
    MatResul = Array(0)
    MatResulDiff = Array(0)
    MatCalend = Array(0)
    
    If pnMiViv = 1 Then   'Esos Parametros se deben cargar solo para la Opcion MiVivienda
        'Porcentaje Real sin dividir enter 100
        Set oParam = New COMDCredito.DCOMParametro
        'nTramoNoConsPorcen = oParam.RecuperaValorParametro(gColocMiVivTramo)
        'nPlazoMiViv = oParam.RecuperaValorParametro(3056)
        Call oParam.RecuperaParametrosCalendario(nTramoNoConsPorcen, nPlazoMiViv)
        Set oParam = Nothing
    End If
    
    nTipoCuota = pnTipoCuota
    nTipoPeriodo = pnTipoPeriodo
    
    If pnMiViv = 1 Then
        'MAVM 20120401 ***
        'nTramoNoConsMonto = Format((nTramoNoConsPorcen / 100) * pnMonto, "#0.00")
        'nTramoConsMonto = Format(pnMonto - nTramoNoConsMonto, "#0.00")
'        nTramoNoConsMonto = Format(CDbl(pnMonto) - nTramoNoConsPorcen, "#0.00")
'        nTramoConsMonto = Format(nTramoNoConsPorcen, "#0.00")
        '***
        Set oParam = New COMDCredito.DCOMParametro
        If pnMontoMivivienda > oParam.RecuperaValorParametro(2001) * 50 Then
                nTramoConsMonto = oParam.RecuperaValorParametro(gColocMiVivBonoBuenPagador2)
        Else
                nTramoConsMonto = oParam.RecuperaValorParametro(gColocMiVivBonoBuenPagador)
        End If
        nTramoNoConsMonto = Format(CDbl(pnMonto) - nTramoConsMonto, "#0.00")
        Set oParam = Nothing

        Call GeneraDobleCalendario(MatCalend, MatCalend_2, nTramoConsMonto, nTramoNoConsMonto, pnTasaInt, pnNroCuotas, pnPeriodo, _
                                        pdFecDesemb, nTipoCuota, nTipoPeriodo, nTipoGracia, pnDiasGracia, _
                                         pnDiaFijo, IIf(bProxMes, True, False), MatGracia, IIf(pnMiViv = 1, True, False), _
                                         boptTipoGracia_0, pnTasaGracia)
        MatResul = UnirMatricesMiVivienda(MatCalend, MatCalend_2, pnMonto)
        MatResulDiff = DiferencialMatricesMiVivienda(MatCalend, MatResul)
            
        'Cargar Matrices a FlexGrid
        nTotalInteres = 0
        nTotalCapital = 0
        nTotalcuotasCONItF = 0
        For i = 0 To UBound(MatCalend) - 1
            nTotalCapital = nTotalCapital + CDbl(Trim(MatCalend(i, 3)))
            nTotalcuotasCONItF = nTotalCapital + Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2)), gnITFMontoMin, gnITFPorcent, gbITFAplica), "0.00")
            nTotalInteres = nTotalInteres + CDbl(Trim(MatCalend(i, 4))) + CDbl(Trim(MatCalend(i, 5)))
        Next i
    Else
        'Se Agrego para manejar la Capitalizacion de la Gracia
        If boptTipoGracia_0 Then
            Dim oCredito As COMNCredito.NCOMCredito
            Set oCredito = New COMNCredito.NCOMCredito
            'Para realizar los cálculos
            nMontoCapInicial = pnMonto
            'If pbIncrementaCapital Then
            'pnMonto = Format(pnMonto + oCredito.MontoIntPerDias(pnTasaGracia, pnDiasGracia, pnMonto), "#0.00")
            'End If
            Set oCredito = Nothing
            'Los calculos se hacen con el Nuevo Monto
            'If pbIncrementaCapital Then
            'nMontoCapInicial = 0
        End If
        '*********************************************
        Dim dFechaInicio As Date
        dFechaInicio = pdInicioPago
        
        MatCalend = GeneraCalendario(pnMonto, pnTasaInt, pnNroCuotas, _
                    pnPeriodo, pdFecDesemb, nTipoCuota, nTipoPeriodo, _
                    nTipoGracia, pnDiasGracia, pnDiaFijo, IIf(bProxMes, 1, 0), MatGracia, True, _
                    IIf(pnCuotaCom = 1, True, False), bDesemParcial, MatDesPar, , , , _
                    boptTipoGracia_1, pnTasaGracia, pnDiaFijo2, nMontoCapInicial, , , , , pnMontoGra, , pnCuotaBalon, pbReprogConvenio)
                    'WIOR 20131111 AGREGO pnCuotaBalon
                    'WIOR 20140514 AGREGO pbReprogConvenio
        
        'Carga Flex Edit
        nTotalInteres = 0
        nTotalCapital = 0
        nTotalcuotasCONItF = 0
        
        For i = 0 To UBound(MatCalend) - 1
            If Not (nTipoGracia = gColocTiposGraciaCapitalizada And i = 0) Then
                nTotalCapital = nTotalCapital + CDbl(Trim(MatCalend(i, 3)))
                'nTotalInteres = nTotalInteres + CDbl(Trim(MatCalend(I, 4)))
                nTotalInteres = nTotalInteres + CDbl(Trim(MatCalend(i, 4))) + CDbl(Trim(MatCalend(i, 5)))
                nTotalcuotasCONItF = nTotalcuotasCONItF + Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2)), gnITFMontoMin, gnITFPorcent, gbITFAplica), "0.00")
            End If
        Next i
    End If
                                                   
    '**********************************************************************************
    GenerarCalendarioPagos = MatCalend
    MatMiViv_2 = MatResulDiff
    pbErrorValidacion = True
        
End Function

Public Function fgITFCalculaImpuesto(ByVal pnMonto As Double, ByVal gnITFMontoMin As Double, _
    ByVal gnITFPorcent As Double, ByVal gbITFAplica As Double) As Double

Dim lnValor As Double
lnValor = pnMonto
If gbITFAplica = True Then
    If pnMonto > gnITFMontoMin Then
        
        lnValor = pnMonto * gnITFPorcent
        
        Dim aux As Double
        If InStr(1, CStr(lnValor), ".", vbTextCompare) > 0 Then
            aux = CDbl(CStr(Int(lnValor)) & "." & Mid(CStr(lnValor), InStr(1, CStr(lnValor), ".", vbTextCompare) + 1, 2))
        Else
            aux = CDbl(CStr(Int(lnValor)))
        End If
        lnValor = aux
        lnValor = Format(lnValor, "#0.00")
    End If
End If
fgITFCalculaImpuesto = lnValor
End Function

Public Function UnirMatricesMiVivienda(ByVal pMat1 As Variant, ByVal pMat2 As Variant, _
    ByVal pnMonto As Double) As Variant
Dim i As Integer
Dim MatResul() As String
Dim nNumProrat As Integer
Dim nMontoTemp As Double
Dim nCuotaTemp As Double
Dim nMontoTotal As Double
Dim nIndS As Integer
Dim nTramoCons As Double
        nMontoTemp = pnMonto
        'ReDim MatResul(UBound(pMat1), 8) 'MAVM 20120709
        ReDim MatResul(UBound(pMat1), 10) 'MAVM 20120709
        nTramoCons = 12500 'MAVM 20121113
        nNumProrat = 0
        nIndS = -1
        For i = 0 To UBound(pMat1) - 1
            If i < 6 Then
                MatResul(i, 0) = pMat1(i, 0) 'fecha
                MatResul(i, 1) = pMat1(i, 1) 'Cuota
                MatResul(i, 2) = pMat1(i, 2) 'Monto Cuota
                MatResul(i, 3) = pMat1(i, 3) 'Capital
                MatResul(i, 4) = pMat1(i, 4) 'Interes
                MatResul(i, 5) = pMat1(i, 5) 'Gracia
                'MatResul(i, 6) = pMat1(i, 6) 'Gasto
                'MatResul(i, 7) = pMat1(i, 7) 'Saldo
                MatResul(i, 6) = pMat1(i, 8) 'Gasto
                MatResul(i, 8) = pMat1(i, 6) 'Seg Desg
                MatResul(i, 9) = pMat1(i, 9) 'Seg Inm
            Else
                MatResul(i, 0) = pMat1(i, 0) 'fecha
                MatResul(i, 1) = pMat1(i, 1) 'Cuota
                MatResul(i, 3) = Format(CDbl(pMat1(i, 3)) + CDbl(pMat2(nIndS, 3)) / nNumProrat, "#0.00") 'Capital
                MatResul(i, 4) = Format(CDbl(pMat1(i, 4)) + CDbl(pMat2(nIndS, 4)) / nNumProrat, "#0.00") 'Interes
                MatResul(i, 5) = Format(CDbl(pMat1(i, 5)) + CDbl(pMat2(nIndS, 5)) / nNumProrat, "#0.00") 'Gracia
                'MatResul(i, 6) = Format(CDbl(pMat1(i, 6)) + CDbl(pMat2(nIndS, 6)) / nNumProrat, "#0.00") 'Gasto
                'MatResul(i, 2) = Format(CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)), "#0.00")
                MatResul(i, 6) = Format((CDbl(pMat1(i, 8)) + CDbl(pMat2(nIndS, 8)) / nNumProrat) + ((nTramoCons * 0.02081) / 100), "#0.00") 'Gasto
                nTramoCons = nTramoCons - Format(CDbl(pMat2(nIndS, 3)) / nNumProrat, "#0.00")
                MatResul(i, 8) = Format(CDbl(pMat1(i, 6)) + CDbl(pMat2(nIndS, 6)) / nNumProrat, "#0.00") 'Seg Desg
                MatResul(i, 9) = Format(CDbl(pMat1(i, 9)) + CDbl(pMat2(nIndS, 9)) / nNumProrat, "#0.00") 'Seg Inm
                MatResul(i, 2) = Format(CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)) + CDbl(MatResul(i, 8)) + CDbl(MatResul(i, 9)), "#0.00")
                'MatResul(i, 7) = pMat1(i, 7) 'Saldo
                
                'nCuotaTemp = Format(CDbl(pMat1(i, 2)) + CDbl(pMat2(nIndS, 2)) / nNumProrat, "#0.00") 'Monto Cuota
                'If nCuotaTemp <> (CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6))) Then
                'If nCuotaTemp <> (CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)) + CDbl(MatResul(i, 8)) + CDbl(MatResul(i, 9))) Then
                '    'MatResul(i, 4) = Format(CDbl(MatResul(i, 4)) + (nCuotaTemp - (CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)))), "#0.00")
                '    'MatResul(i, 2) = Format(CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)), "#0.00")
                '    MatResul(i, 4) = Format(CDbl(MatResul(i, 4)) + (nCuotaTemp - (CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)) + CDbl(MatResul(i, 8)) + CDbl(MatResul(i, 9)))), "#0.00")
                '    MatResul(i, 2) = Format(CDbl(MatResul(i, 3)) + CDbl(MatResul(i, 4)) + CDbl(MatResul(i, 5)) + CDbl(MatResul(i, 6)) + CDbl(MatResul(i, 8)) + CDbl(MatResul(i, 9)), "#0.00")
                'End If
                
            End If
                        
            
            If (i + 1) Mod 6 = 0 Then
                If i <> 0 Then
                    nIndS = nIndS + 1
                End If
                If (UBound(pMat1) - i) >= 6 Then
                    nNumProrat = 6
                Else
                    nNumProrat = UBound(pMat1) - i
                End If
            End If
        Next i
                
        'Agregando la ultima Cuota Semestral
        'MatResul(UBound(MatResul) - 1, 3) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(pMat2(UBound(pMat2) - 1, 3)), "#0.00")  'Capital
        'MatResul(UBound(MatResul) - 1, 4) = Format(CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(pMat2(UBound(pMat2) - 1, 4)), "#0.00")  'Interes
        'MatResul(UBound(MatResul) - 1, 5) = Format(CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(pMat2(UBound(pMat2) - 1, 5)), "#0.00")  'Gracia
        ''MatResul(UBound(MatResul) - 1, 6) = Format(CDbl(MatResul(UBound(MatResul) - 1, 6)) + CDbl(pMat2(UBound(pMat2) - 1, 6)), "#0.00")  'Gasto
        ''MatResul(UBound(MatResul) - 1, 2) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(MatResul(UBound(MatResul) - 1, 6)), "#0.00")
        'MatResul(UBound(MatResul) - 1, 6) = Format(CDbl(MatResul(UBound(MatResul) - 1, 8)) + CDbl(pMat2(UBound(pMat2) - 1, 8)), "#0.00")  'Gasto
        'MatResul(UBound(MatResul) - 1, 8) = Format(CDbl(MatResul(UBound(MatResul) - 1, 6)) + CDbl(pMat2(UBound(pMat2) - 1, 6)), "#0.00")  'Seg Desg
        'MatResul(UBound(MatResul) - 1, 9) = Format(CDbl(MatResul(UBound(MatResul) - 1, 9)) + CDbl(pMat2(UBound(pMat2) - 1, 9)), "#0.00")  'Seg Inm
        'MatResul(UBound(MatResul) - 1, 2) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(MatResul(UBound(MatResul) - 1, 6)) + CDbl(MatResul(UBound(MatResul) - 1, 8)) + CDbl(MatResul(UBound(MatResul) - 1, 9)), "#0.00")
                
        'Comprobando Capital Total sea Igual a Prestamo
        nMontoTotal = 0
        For i = 0 To UBound(MatResul) - 1
            nMontoTotal = nMontoTotal + CDbl(MatResul(i, 3))
        Next i
        
        If nMontoTotal <> pnMonto Then
            MatResul(UBound(MatResul) - 1, 3) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) - (nMontoTotal - pnMonto), "#0.00")
            'MatResul(UBound(MatResul) - 1, 2) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(MatResul(UBound(MatResul) - 1, 6)), "#0.00")
            MatResul(UBound(MatResul) - 1, 2) = Format(CDbl(MatResul(UBound(MatResul) - 1, 3)) + CDbl(MatResul(UBound(MatResul) - 1, 4)) + CDbl(pMat2(nIndS, 4)) + CDbl(MatResul(UBound(MatResul) - 1, 5)) + CDbl(MatResul(UBound(MatResul) - 1, 6)) + CDbl(MatResul(UBound(MatResul) - 1, 8)) + CDbl(MatResul(UBound(MatResul) - 1, 9)), "#0.00")
        End If
        
        nMontoTemp = pnMonto
        For i = 0 To UBound(MatResul) - 1
            nMontoTemp = nMontoTemp - CDbl(MatResul(i, 3))
            nMontoTemp = Format(nMontoTemp, "#0.00")
            MatResul(i, 7) = Format(nMontoTemp, "#0.00")
        Next i
        
        
        UnirMatricesMiVivienda = MatResul
End Function

Public Function DiferencialMatricesMiVivienda(ByVal pMat1 As Variant, ByVal pMat2 As Variant) As Variant
Dim i As Integer
Dim MatResul As Variant
    'ReDim MatResul(UBound(pMat1), 7)
    ReDim MatResul(UBound(pMat1), 9) 'MAVM 20120709
    For i = 0 To UBound(pMat1) - 1
        MatResul(i, 0) = pMat1(i, 0) 'fecha
        MatResul(i, 1) = pMat1(i, 1) 'Cuota
        MatResul(i, 2) = pMat1(i, 2) 'Monto Cuota
        MatResul(i, 3) = Format(CDbl(pMat2(i, 3)) - CDbl(pMat1(i, 3)), "#0.00") 'Capital
        MatResul(i, 4) = Format(CDbl(pMat2(i, 4)) - CDbl(pMat1(i, 4)), "#0.00") 'Interes
        'MatResul(i, 5) = Format(CDbl(pMat2(i, 5)) - CDbl(pMat1(i, 5)), "#0.00") 'Gracia
        'MatResul(i, 6) = Format(CDbl(pMat2(i, 6)) - CDbl(pMat1(i, 6)), "#0.00") 'Gasto
    Next i
    DiferencialMatricesMiVivienda = MatResul
End Function


Private Sub Class_Initialize()
Set oImpre = New COMFunciones.FCOMImpresion
Set oITF = New COMDConstSistema.FCOMITF
IniciaImpresora
End Sub

Private Sub Class_Terminate()
    Set oImpre = Nothing
    Set oITF = Nothing
End Sub

'MAVM 20120606 ***
Public Function ReporteResumenMiVivienda(ByVal pTipoRep As Byte, ByVal MatP As Variant, ByVal MatMalo As Variant, _
        ByVal pTCuota As String, ByVal pInteres As Double, ByVal pMonto, _
        ByVal pCuota As Double, ByVal pPlazo As Integer, ByVal pVigencia As Date, _
        Optional ByVal pnTipoProceso As Integer = 0, Optional ByVal pMatDesPar As Variant = "", Optional GBITFAPLICAt As Boolean, Optional gnITFPorcentt As Double, Optional gnITFMontoMint As Double = 0, _
        Optional ByVal psCtaCod As String, Optional pnCuotas As Integer, _
        Optional ByVal pnTasaEfectivaAnual As Double, Optional ByVal pnTasaCostoEfectivoAnual As Double, Optional ByVal lsCtaCodLeasing As String = "", _
        Optional ByVal pnValorInmueble As Double, Optional ByVal pnCuotaInicial As Double, Optional ByVal pnBonoBuenPagador As Double, _
        Optional ByVal pnDiasGracia As Integer, Optional ByVal psSeguroDesgravamen As String, Optional ByVal psGastoAdministracion As Double, _
        Optional ByVal pnCuotMens As Double, Optional ByVal pnCuotMensBono As Double) As String

Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim cTitulo As String
Dim i As Integer
Dim cCuota As String * 35
Dim sCuota As Double '** PEAC 20080723
Dim sCapital As Double '** PEAC 20080723
Dim sInteres As Double '** PEAC 20080723
Dim sGastos As Double '** PEAC 20080723
Dim sSegDes As Double '** MAVM 20100320
Dim sSql As String
Dim oConec As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim sNombre As String
Dim oParametros As COMDCredito.DCOMParametro
Dim nPorComiCorrespMN As Double
Dim nPorComiCorrespME As Double
Dim nMonMinComiCorrespMN As Double
Dim nMonMaxComiCorrespMN As Double
Dim nMonMinComiCorrespME As Double
Dim nFactorCorresp As Double
Dim nMonMinComi As Double
Dim nMonMaxComi As Double
Dim nComiCorresp As Double
'*************************************************
oITF.gbITFAplica = GBITFAPLICAt
oITF.gnITFPorcent = gnITFPorcentt
oITF.gnITFMontoMin = gnITFMontoMint
 
cCuota = pTCuota
sCapital = 0
sInteres = 0

    nFactorCorresp = 1
    nMonMinComi = 0
    nMonMaxComi = 0
    If psCtaCod <> "" Then
        Set oParametros = New COMDCredito.DCOMParametro
        nPorComiCorrespMN = oParametros.RecuperaValorParametro(gColPParamPorComiCorresponsaliaMN)
        nPorComiCorrespME = oParametros.RecuperaValorParametro(gColPParamPorComiCorresponsaliaME)
        nMonMinComiCorrespMN = oParametros.RecuperaValorParametro(gColPParamMonMinCobroCorresponsaliaMN)
        nMonMaxComiCorrespMN = oParametros.RecuperaValorParametro(gColPParamMonMaxCobroCorresponsaliaMN)
        nMonMinComiCorrespME = oParametros.RecuperaValorParametro(gColPParamMonMinCobroCorresponsaliaME)
        If Mid$(psCtaCod, 9, 1) = "1" Then 'MN
            nFactorCorresp = nPorComiCorrespMN / 100
            nMonMinComi = nMonMinComiCorrespMN
            nMonMaxComi = nMonMaxComiCorrespMN
        Else
            nFactorCorresp = nPorComiCorrespME / 100
            nMonMinComi = nMonMinComiCorrespME
            nMonMaxComi = 9999999999#
        End If
    End If
    
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
            cTitulo = "SIMULADOR DE CREDITO HIPOTECARIO MI VIVIENDA"
        ElseIf pnTipoProceso = 1 Then
            cTitulo = "SUGERENCIA DE ANALISTA - CALENDARIO DE PAGOS MI VIVIENDA"
        ElseIf pnTipoProceso = 2 Then
            cTitulo = "APROBACION DE CREDITO - CALENDARIO DE PAGOS MI VIVIENDA"
        End If
    End If

    lnLineas = 0
    lnPage = 1

    'obteniendo titular
    sSql = "Select Pers.cPersNombre"
    sSql = sSql & " From ProductoPersona PP"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Where PP.cCtaCod='" & psCtaCod & "' And PP.nPrdPersRelac=20"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        sNombre = rs!cPersNombre
    End If
    Set rs = Nothing
    
    'CABECERA
    If pTipoRep = 1 Then
        lsCadImp = lsCadImp & nRepoCabecera(cTitulo, " ", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, psCtaCod, Trim(sNombre), pnCuotas, pnTasaEfectivaAnual, pnTasaCostoEfectivoAnual) 'DAOR 20070403
    ElseIf pTipoRep = 2 Then
        lsCadImp = lsCadImp & nRepoCabeceraResumen(cTitulo, "BUEN PAGADOR", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, psCtaCod, Trim(sNombre), pnCuotas, pnTasaEfectivaAnual, pnTasaCostoEfectivoAnual, pnValorInmueble, pnCuotaInicial, pnBonoBuenPagador, pnDiasGracia, psSeguroDesgravamen, psGastoAdministracion, pnCuotMens, pnCuotMensBono)
    End If
    
ReporteResumenMiVivienda = lsCadBuffer & lsCadImp
End Function

Public Function nRepoCabeceraResumen(ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        ByVal pnCodReporte As String, ByVal pTCuota As String, ByVal pInteres As Double, ByVal pMonto As Double, _
        ByVal pCuota As Double, ByVal pPlazo As Integer, ByVal pVigencia As Date, _
        Optional ByVal psCtaCod As String, Optional ByVal psNomCli As String, Optional ByVal pnNroCuotas As Integer, _
        Optional ByVal pnTasaEfectivaAnual As Double, Optional ByVal pnTasaCostoEfectivoAnual As Double, _
        Optional ByVal pnValorInmueble As Double, Optional ByVal pnCuotaInicial As Double, Optional ByVal pnBonoBuenPagador As Double, _
        Optional ByVal pnDiasGracia As Integer, Optional ByVal psSeguroDesgravamen As String, Optional ByVal psGastoAdministracion As Double, _
        Optional ByVal pnCuotMens As Double, Optional ByVal pnCuotMensBono As Double) As String
        
    Dim lsCadImp As String

    psSubTitulo = ""
    lsCadImp = nImprimeCabeceraReportes(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & String(pnAnchoLinea - 10, "-") & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    Select Case pnCodReporte
        Case "CalenPagos"
            'lsCadImp = lsCadImp & "    CREDITO         : " & psCtaCod & oImp.gPrnSaltoLinea
            'lsCadImp = lsCadImp & "    CLIENTE         : " & psNomCli & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    VALOR DEL INMUEBLE    : " & "S/. " & Trim(oImpre.ImpreFormat(pnValorInmueble, 8, 2)) & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    CUOTA INICIAL         : " & "S/. " & Trim(oImpre.ImpreFormat(pnCuotaInicial, 8, 2)) & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    IMPORTE PRESTAMO      : " & "S/. " & Trim(oImpre.ImpreFormat(pnValorInmueble - pnCuotaInicial, 8, 2)) & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    BONO DEL BUEN PAGADOR : " & "S/. " & Trim(oImpre.ImpreFormat(pnBonoBuenPagador, 8, 2)) & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    TOTAL A FINANCIAR     : " & "S/.  " & Trim(oImpre.ImpreFormat((pnValorInmueble - pnCuotaInicial) - pnBonoBuenPagador, 8, 2)) & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    PLAZO                 : " & Trim(oImpre.ImpreFormat((CInt(pnNroCuotas) * 30) / 360, 3, 0)) & " Años" & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    NRO DE CUOTAS         : " & Trim(oImpre.ImpreFormat(pnNroCuotas, 3, 0)) & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    PERIODO DE GRACIA     : " & Trim(oImpre.ImpreFormat(pnDiasGracia, 3, 0)) & " dias" & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    SEGURO DE DESGRAVAMEN : " & Trim(psSeguroDesgravamen) & Space(16) & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    GASTOS                : " & "S/. " & Trim(oImpre.ImpreFormat(psGastoAdministracion, 8, 2)) & oImp.gPrnSaltoLinea
            
            'lsCadImp = lsCadImp & "    TIPO DE CUOTA   : " & pTCuota & Space(16) & "Nro Cuotas: " & Trim(oImpre.ImpreFormat(pnNroCuotas, 3, 0)) & oImp.gPrnSaltoLinea
            'lsCadImp = lsCadImp & "    PLAZO           : " & Trim(oImpre.ImpreFormat(pPlazo, 3, 0)) & oImp.gPrnSaltoLinea
            'lsCadImp = lsCadImp & "    MONTO           : " & Trim(oImpre.ImpreFormat(pMonto, 8, 2)) & Space(45) & "VIGENCIA  : " & Format(pVigencia, "ddd, dd mmm yyyy") & oImp.gPrnSaltoLinea
            'lsCadImp = lsCadImp & "    MONEDA          : " & IIf(Mid(psCtaCod, 9, 1) = 1, "SOLES", IIf(Mid(psCtaCod, 9, 1) = 2, "DOLARES", " ")) & oImp.gPrnSaltoLinea
            If pnTasaEfectivaAnual > 0 Then lsCadImp = lsCadImp & "    TASA EFECTIVA ANUAL   : " & Trim(oImpre.ImpreFormat(pnTasaEfectivaAnual, 3, 2)) & " %" & oImp.gPrnSaltoLinea
            If pnTasaEfectivaAnual > 0 Then lsCadImp = lsCadImp & "    TASA SEG DESGRAV MEN  : " & IIf(Left(psSeguroDesgravamen, 1) = "I", "0.0429", "0.0772") & " %" & oImp.gPrnSaltoLinea
            If pnTasaEfectivaAnual > 0 Then lsCadImp = lsCadImp & "    TASA SEG INMUEBL MEN  : " & "0.022 %" & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    FECHA DE SIMULACION   : " & Format(pVigencia, "ddd, dd mmm yyyy") & oImp.gPrnSaltoLinea
            'If pnTasaCostoEfectivoAnual > 0 Then lsCadImp = lsCadImp & "    TASA COSTO EFECTIVO ANUAL : " & oImpre.ImpreFormat(pnTasaCostoEfectivoAnual, 3, 2) & " %" & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & String(pnAnchoLinea - 10, "-") & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
            
            lsCadImp = lsCadImp & "    CUOTA TOTAL MENSUAL              : " & "S/. " & Trim(oImpre.ImpreFormat(pnCuotMens, 8, 2)) & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & oImpre.PrnSet("B+") & "    CUOTA CON PREMIO AL BUEN PAGADOR : " & "S/. " & Trim(oImpre.ImpreFormat(pnCuotMensBono, 8, 2)) & oImpre.PrnSet("B-") & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    INGRESO NETO SOLICITADO          : " & "S/. " & Trim(oImpre.ImpreFormat(pnCuotMens / 0.3999, 8, 2)) & oImp.gPrnSaltoLinea
            If pnTasaCostoEfectivoAnual > 0 Then lsCadImp = lsCadImp & "    TASA COSTO EFECTIVO ANUAL        : " & Trim(oImpre.ImpreFormat(pnTasaCostoEfectivoAnual, 3, 2)) & " %" & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & String(pnAnchoLinea - 10, "-") & oImp.gPrnSaltoLinea
            
       End Select
       
nRepoCabeceraResumen = lsCadImp
End Function
'***

'--->***** LUCV20180320, Según ERS087-2017
Private Sub CalendarioCuotaFijaPagoAnticipado(ByVal pnMonto As Double, _
                ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
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
                Optional ByVal pbMiViv As Boolean = False, _
                Optional ByVal pnCuotaIni As Integer = 0, _
                Optional ByVal pnCuotaFin As Integer = 0, _
                Optional ByVal pnDiaFijo2 As Integer = 0, _
                Optional pnMontoCapInicial As Double = 0, _
                Optional ByVal psCtaCod As String, _
                Optional ByVal psMensFecAprob As Date, _
                Optional ByVal lnDescripCuota As Variant = "", _
                Optional ByVal pnMontoGra As Double = 0, _
                Optional ByVal psIncluyeIntCap As String = "S", _
                Optional ByVal pnCuotaBalon As Integer = 0, _
                Optional ByVal pbReprogConvenio As Boolean = False)
                
Dim nSaldoCapital As Double
Dim dDesembolso As Date
Dim dFecAproGracia As Date
Dim i As Integer
Dim oCredito As COMNCredito.NCOMCredito
Dim dFecTemp As Date
Dim nMes As Integer
Dim nAnio As Integer
Dim nDia As Integer
Dim nMontoCuotaTemp As Double
Dim nMontoGraciaTemp As Double
Dim nCDIntPend As Double
Dim nPeriodoDesPar As Integer
Dim nSumCapital As Double
Dim nPeriodoPrimerPF As Integer

Dim nSumCapitalGra As Double
Dim nSaldoCapitalGra As Double
Dim nNroCuotasAfec As Integer

Dim dFechaCuotaPend As Date  'LUCV20180320, Agregó según ERS087-2017
Dim dDesembolsoNuevo As Date 'LUCV20180320, Agregó según ERS087-2017


nSumCapital = 0
nPeriodoPrimerPF = 0

nMontoGraciaTemp = 0
nSumCapitalGra = 0
nSaldoCapitalGra = 0

    'Datos de Pagos Anticipados-----------------------------------------------
    '***---> LUCV20180320, Agregó según ERS087-2017
    Dim nMontoCuotaDesembolso As Double
    Dim oDCOMCredito As COMDCredito.DCOMCredito
    Set oDCOMCredito = New COMDCredito.DCOMCredito

    nMontoCuotaDesembolso = oDCOMCredito.CuotaFijaMensualPago(psCtaCod)
    Set oDCOMCredito = Nothing
    '<---*** Fin LUCV20180320
    
    If psMensFecAprob = "12:00:00 AM" And psCtaCod <> "" Then
         Dim oCredT As COMNCredito.NCOMCredito
         Dim RMantPP As ADODB.Recordset
         Dim DfechaPre As Date
         Dim PlazoPre As Integer
         Dim indPre As Boolean
         
         Set oCredT = New COMNCredito.NCOMCredito
         Set RMantPP = oCredT.RecuperaValidacionCredMantPrepago(psCtaCod, Format(pdFecDesemb, "dd/mm/yyyy"))
         Set oCredT = Nothing
         indPre = False
         
         If Not RMantPP.BOF And Not RMantPP.EOF Then
             indPre = True
             DfechaPre = RMantPP!dVenc
             PlazoPre = RMantPP!DifPlazo
             dFechaCuotaPend = RMantPP!dFechaCuotaPend 'LUCV20180320, Agregó según ERS087-2017
             dDesembolsoNuevo = RMantPP!dDesembolsoNuevo 'LUCV20180320, Agregó según ERS087-2017
             RMantPP.Close
         End If
         Set RMantPP = Nothing
    End If
     '------------------------------------------------------------------------
     
     nSaldoCapital = pnMonto
     nSaldoCapitalGra = pnMontoGra
     dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
    
    '********************** Tipo Periodo : Periodo Fijo **********************
    '*************************************************************************
     If pnTipoPeriodo = PeriodoFijo Then
         Set oCredito = New NCOMCredito
         For i = pnCuotaIni To pnCuotaFin
            'No entre cuando es aprobacion o PAGO
            If i = 0 And psMensFecAprob <> "12:00:00 AM" Then
                'Verifica Si tiene Gracia / Diferencia es -
                 If i = 0 And (dDesembolso > psMensFecAprob) Then
                       If lnDescripCuota(0) = "0" Then
                         nPeriodoPrimerPF = pnPeriodo - CInt(DateDiff("d", psMensFecAprob, dDesembolso))
                       Else
                             If lnDescripCuota(1) <> "0" Then
                                 nPeriodoPrimerPF = pnPeriodo + Val(lnDescripCuota(1))
                             End If
                       End If
                 End If
            End If
        
            Calendario(i).dFecha = dDesembolso + IIf(i = 0, PlazoPre, pnPeriodo)
            Calendario(i).NroCuota = i + 1
            Calendario(i).IntGra = 0#
            Calendario(i).Gasto = 0#
            
            Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, IIf(i = 0, PlazoPre, pnPeriodo), nSaldoCapital)
            '->***** LUCV20180530, Según Anexo 002 del Acta066-2018
            If (nSaldoCapital >= nMontoCuotaDesembolso) Then
                Calendario(i).Cuota = nMontoCuotaDesembolso 'LUCV20180320
            Else
                Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, IIf(i = 0, PlazoPre, pnPeriodo))
            End If
            '<-***** Fin LUCV20180530
            If i = pnCuotaFin Then
                Calendario(i).Captital = nSaldoCapital
                If (Calendario(i).Captital) < Calendario(i).Cuota Then
                    Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                'Else 'LUCV20180320
                '   Calendario(i).Cuota = Calendario(i).Cuota - Calendario(i).Captital
                End If
            Else
                Calendario(i).Captital = Calendario(i).Cuota - Calendario(i).IntComp
            End If
             
             nSumCapital = nSumCapital + Calendario(i).Captital
             Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
             nSaldoCapital = nSaldoCapital - Calendario(i).Captital
             nSaldoCapital = CDbl(Format(nSaldoCapital, "#0.00"))
             
             dDesembolso = dDesembolso + IIf(i = 0, PlazoPre, pnPeriodo)
             nPeriodoPrimerPF = 0
         Next i
         
    '********************** Tipo Periodo : Fecha Fija **********************
    '*************************************************************************
     Else
         If pnTipoPeriodo = FechaFija Then
             Set oCredito = New COMNCredito.NCOMCredito
             nMes = Month(dDesembolso)
             nAnio = Year(dDesembolso)
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
                 
                 'Se modifico para manejar 2 dias fijos
                 If Not (i = 0 And nDia >= Day(dDesembolso) And (Not bProxMes)) Or pnNumMes = 6 Then
                     If i = 0 And pnDiasGracia <> 0 Then
                         dFecTemp = dDesembolso + 30
                         nDia = Day(dFecTemp)
                         nMes = Month(dFecTemp)
                         nAnio = Year(dFecTemp)
                     Else
                         If nDia = pnDiaFijo Then nMes = nMes + pnNumMes
                         If nMes > 12 Then
                             nAnio = nAnio + 1
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
                    
                 'Si el cliente realizó pagos, solo toma la fecha pendiente de pago
                 If psCtaCod <> "" Then
                     If i = 0 Then
                        'If dFechaCuotaPend > CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy")) Then
                        'dDesembolso = dDesembolsoNuevo
                        nDia = Day(dFechaCuotaPend)
                        nMes = Month(dFechaCuotaPend)
                        nAnio = Year(dFechaCuotaPend)
                         If (dFechaCuotaPend > dDesembolso) And (dDesembolsoNuevo > dDesembolso) Then
                                dDesembolso = dDesembolsoNuevo
                         End If
                     End If
                 End If
 
                 dFecTemp = CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy"))
                 Calendario(i).dFecha = dFecTemp
                 Calendario(i).NroCuota = i + 1
                 Calendario(i).IntGra = 0#
                 Calendario(i).Gasto = 0


                 nNroCuotasAfec = pnNroCuotas - pnCuotaBalon
                 If i = pnCuotaIni Then
                    '->***** LUCV20180530, Según Anexo 002 del Acta066-2018
                     If (nSaldoCapital >= nMontoCuotaDesembolso) Then
                        nMontoCuotaTemp = nMontoCuotaDesembolso 'LUCV20180320, Según ERS087-2017
                     Else
                        nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, nNroCuotasAfec, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas)
                     End If
                     '<-***** Fin LUCV20180530
                     Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(i).dFecha), nSaldoCapital)
                     Calendario(i).Cuota = nMontoCuotaTemp
                 Else
                     Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha), nSaldoCapital)
                     Calendario(i).Cuota = nMontoCuotaTemp
                 End If
                 
                 If pbCuotaFijaFechaFija Then
                     Calendario(i).IntComp = Calendario(i).IntComp + nCDIntPend
                     nCDIntPend = 0#
                     If Calendario(i).IntComp > nMontoCuotaTemp Then
                         nCDIntPend = Calendario(i).IntComp - nMontoCuotaTemp
                         Calendario(i).IntComp = nMontoCuotaTemp
                     End If
                 End If
                 
                 If i = pnCuotaFin Then
                     Calendario(i).Captital = nSaldoCapital
                     If (Calendario(i).Captital) < Calendario(i).Cuota Then
                         Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                     'Else 'LUCV20180320
                         'Calendario(i).Cuota = Calendario(i).Cuota - Calendario(i).Captital
                     End If
                 Else
                     Calendario(i).Captital = Calendario(i).Cuota - Calendario(i).IntComp
                 End If

                 nSumCapital = nSumCapital + Calendario(i).Captital
                 Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                 nSaldoCapital = nSaldoCapital - Calendario(i).Captital
             Next i
         End If
     End If
End Sub
 '<---***** Fin LUCV20180320
'ARLO20180723 ERS037-2018
Public Function ReporteCalendarioNEW(ByVal pTipoRep As Byte, Optional ByVal MatP As Variant = "", Optional ByVal MatMalo As Variant = "", _
        Optional ByVal pTCuota As String, Optional ByVal pInteres As Double, Optional ByVal pMonto, _
        Optional ByVal pCuota As Double, Optional ByVal pPlazo As Integer, Optional ByVal pVigencia As Date, _
        Optional ByVal pnTipoProceso As Integer = 0, Optional ByVal pMatDesPar As Variant = "", Optional GBITFAPLICAt As Boolean, Optional gnITFPorcentt As Double, Optional gnITFMontoMint As Double = 0, _
        Optional ByVal psCtaCod As String, Optional pnCuotas As Integer, _
        Optional ByVal pnTasaEfectivaAnual As Double, Optional ByVal pnTasaCostoEfectivoAnual As Double, Optional ByVal lsCtaCodLeasing As String = "", _
        Optional ByVal nTipoGracia As Integer, Optional ByVal nIntGraInicial As Double, Optional ByVal nDiasAtraso As Integer, Optional psPersNom As String = "", _
        Optional ByVal ArrSimulador As Variant = Nothing, _
        Optional ByVal psTpoSubPro As String, Optional ByVal psTpoPro As String, Optional ByVal psTpoMoneda As String, _
        Optional ByVal pdFechaVencPigno As Date, Optional ByVal pnTotalPigno As Double, Optional ByVal pnInteresPigno As Double, _
        Optional ByVal pnTasaNormal As Double, Optional ByVal pnTasaCPP As Double, Optional ByVal psTpoCliente As String) As String

        
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer

Dim cTitulo As String
Dim cSubTitulo As String
Dim i As Integer
Dim cCuota As String * 35

Dim sCuota As Double
Dim sCapital As Double
Dim sInteres As Double
Dim sGastos As Double
Dim sInteresGra As Double

Dim sSegDes As Double
Dim sSegInm As Double
Dim sSegMult As Double
 
Dim sSql As String
Dim oConec As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim sNombre As String

Dim oParametros As COMDCredito.DCOMParametro
Dim nPorComiCorrespMN As Double
Dim nPorComiCorrespME As Double
Dim nMonMinComiCorrespMN As Double
Dim nMonMaxComiCorrespMN As Double
Dim nMonMinComiCorrespME As Double
Dim nFactorCorresp As Double
Dim nMonMinComi As Double
Dim nMonMaxComi As Double
Dim nComiCorresp As Double
oITF.gbITFAplica = GBITFAPLICAt
oITF.gnITFPorcent = gnITFPorcentt
oITF.gnITFMontoMin = gnITFMontoMint
 
cCuota = pTCuota
sCapital = 0
sInteres = 0
Dim bSimulador As Boolean
Dim bEnvioEstFis As Boolean
Dim nTipoSegDes As Integer
Dim bSegDesGra As Boolean

bSimulador = False
bEnvioEstFis = False
bSegDesGra = False
nTipoSegDes = 0
If IsArray(ArrSimulador) Then
    If ArrSimulador(0) = 1 Then
        bSegDesGra = True
        bSimulador = True
        nTipoSegDes = ArrSimulador(2)
        If ArrSimulador(1) = 1 Then
            bEnvioEstFis = True
        End If
    End If
End If

If (nTipoGracia = 6) Then
    bSimulador = True
End If

Dim oCred As COMDCredito.DCOMCredito
Dim rsCred As ADODB.Recordset
Dim TpoProd As String
Dim TpoCred As String
Dim bEsMIVIVIENDAAnt As Boolean
bEsMIVIVIENDAAnt = False

If Not bSimulador Then
    If Len(Trim(psCtaCod)) >= 18 Then
        Set oCred = New COMDCredito.DCOMCredito
        Set rsCred = oCred.RecuperaColocaciones(psCtaCod)
        TpoProd = IIf(IsNull(rsCred!cTpoProdCod), "", rsCred!cTpoProdCod)
        TpoCred = IIf(IsNull(rsCred!cTpoCredCod), "", rsCred!cTpoCredCod)
        bEsMIVIVIENDAAnt = oCred.EsCredMIVIVENDA(TpoProd, TpoCred, 1)
        Set oCred = Nothing
    End If
End If

cSubTitulo = ""
    nFactorCorresp = 1
    nMonMinComi = 0
    nMonMaxComi = 0
    If psCtaCod <> "" Then
        Set oParametros = New COMDCredito.DCOMParametro
        nPorComiCorrespMN = oParametros.RecuperaValorParametro(gColPParamPorComiCorresponsaliaMN)
        nPorComiCorrespME = oParametros.RecuperaValorParametro(gColPParamPorComiCorresponsaliaME)
        nMonMinComiCorrespMN = oParametros.RecuperaValorParametro(gColPParamMonMinCobroCorresponsaliaMN)
        nMonMaxComiCorrespMN = oParametros.RecuperaValorParametro(gColPParamMonMaxCobroCorresponsaliaMN)
        nMonMinComiCorrespME = oParametros.RecuperaValorParametro(gColPParamMonMinCobroCorresponsaliaME)
        If Mid$(psCtaCod, 9, 1) = "1" Then 'MN
            nFactorCorresp = nPorComiCorrespMN / 100
            nMonMinComi = nMonMinComiCorrespMN
            nMonMaxComi = nMonMaxComiCorrespMN
        Else
            nFactorCorresp = nPorComiCorrespME / 100
            nMonMinComi = nMonMinComiCorrespME
            nMonMaxComi = 9999999999#
        End If
    End If
    '***************************************************************
    
    If pTipoRep = 1 Then
        If pnTipoProceso = 0 Then
            cTitulo = "SIMULACION REFERENCIAL DE CRONOGRAMA DE PAGOS"
            cSubTitulo = ""
        ElseIf pnTipoProceso = 1 Then
            cTitulo = "SUGERENCIA DE ANALISTA - CALENDARIO DE PAGOS"
        ElseIf pnTipoProceso = 2 Then
            cTitulo = "APROBACION DE CREDITO - CALENDARIO DE PAGOS"
        End If
    ElseIf pTipoRep = 2 Then
        If pnTipoProceso = 0 Then
            cTitulo = "SIMULACION DE CALENDARIO DE PAGOS MI VIVIENDA"
            cSubTitulo = "SÓLO PARA USO DE CARÁCTER INFORMATIVO"
        ElseIf pnTipoProceso = 1 Then
            cTitulo = "SUGERENCIA DE ANALISTA - CALENDARIO DE PAGOS MI VIVIENDA"
        ElseIf pnTipoProceso = 2 Then
            cTitulo = "APROBACION DE CREDITO - CALENDARIO DE PAGOS MI VIVIENDA"
        End If
    ElseIf pTipoRep = 3 Then
        If pnTipoProceso = 0 Then
            cTitulo = "SIMULACION REFERENCIAL DE CREDITO PRENDARIO"
            cSubTitulo = ""
        End If
    End If

    lnLineas = 0
    lnPage = 1

    'obteniendo titular
    sSql = "Select Pers.cPersNombre"
    sSql = sSql & " From ProductoPersona PP"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Where PP.cCtaCod='" & psCtaCod & "' And PP.nPrdPersRelac=20"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        sNombre = rs!cPersNombre
    End If
    Set rs = Nothing
    
    'CABECERA
    If pTipoRep = 1 Then
        lsCadImp = lsCadImp & nRepoCabeceraNEW(pTipoRep, cTitulo, cSubTitulo, lnPage, 120, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, IIf(psPersNom <> "", "Rapiflash con tu Plazo Fijo", psCtaCod), IIf(Trim(sNombre) = "", psPersNom, Trim(sNombre)), pnCuotas, pnTasaEfectivaAnual, pnTasaCostoEfectivoAnual, nTipoGracia, nDiasAtraso, psTpoSubPro, psTpoPro, psTpoMoneda, pdFechaVencPigno, 0, 0, 0, 0, "")
    ElseIf pTipoRep = 2 Then
        lsCadImp = lsCadImp & nRepoCabeceraNEW(pTipoRep, cTitulo, cSubTitulo, lnPage, 133, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, psCtaCod, Trim(sNombre), pnCuotas, pnTasaEfectivaAnual, pnTasaCostoEfectivoAnual, nTipoGracia, nDiasAtraso, psTpoSubPro, psTpoPro, psTpoMoneda, pdFechaVencPigno, 0, 0, 0, 0, "")
        ElseIf pTipoRep = 3 Then
        lsCadImp = lsCadImp & nRepoCabeceraNEW(pTipoRep, cTitulo, "", lnPage, 109, "", "CalenPagos", "0", 0, pMonto, 0, 0, pVigencia, "", "", 0, pnTasaEfectivaAnual, pnTasaCostoEfectivoAnual, nTipoGracia, nDiasAtraso, psTpoSubPro, psTpoPro, psTpoMoneda, pdFechaVencPigno, pnTotalPigno, pnInteresPigno, pnTasaNormal, pnTasaCPP, psTpoCliente)
    End If
        
    If IsArray(pMatDesPar) Then
        lsCadImp = lsCadImp & Space(5) & "DESEMBOLSOS PARCIALES" & oImp.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(5) & String(63, "-") & oImp.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(5) & oImpre.ImpreFormat("FECHA", 16, 0) & oImpre.ImpreFormat("MONTO", 10, 0) & oImp.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(5) & String(63, "-") & oImp.gPrnSaltoLinea
        For i = 0 To UBound(pMatDesPar) - 1
            lsCadImp = lsCadImp & Space(5) & oImpre.ImpreFormat(pMatDesPar(i, 0), 10, 0) & oImpre.ImpreFormat(CDbl(pMatDesPar(i, 1)), 10, 2, True) & oImp.gPrnSaltoLinea
        Next i
        lsCadImp = lsCadImp & Space(5) & String(63, "-") & oImp.gPrnSaltoLinea
    End If
    
    If bSimulador Then
        If bEnvioEstFis Then
            If (pTipoRep = 2) Then
                'lsCadImp = lsCadImp & " FECH. VCTO   N°CUOTA      CUOTA      CAPITAL     INT.COMP    INT.GRACIA  GAST/COM(2) SEG.DES(1)     SEG.INM    SEG.MULT   SALDO CAP." & oImp.gPrnSaltoLinea
                lsCadImp = lsCadImp & " FECH. VCTO   N°CUOTA      CUOTA      CAPITAL     INT.COMP    INT.GRACIA  GAST/COM(2) SEG.DES(1)     SEG.INM    SEG.MULT   SALDO CAP." & oImp.gPrnSaltoLinea
            Else
                'lsCadImp = lsCadImp & " FECH. VCTO   N°CUOTA      CUOTA      CAPITAL     INT.COMP    INT.GRACIA  GAST/COM(2)  SEG.DES(1)    SEG INM   SALDO CAP." & oImp.gPrnSaltoLinea
                lsCadImp = lsCadImp & " FECH. VCTO   N°CUOTA      CUOTA      CAPITAL     INT.COMP    INT.GRACIA  GAST/COM(2)  SEG.DES(1)    SEG.MULT   SALDO CAP." & oImp.gPrnSaltoLinea
            End If
        Else
            If nTipoGracia = 0 Then
                'lsCadImp = lsCadImp & " FECH. VCTO   N°CUOTA      CUOTA      CAPITAL     INT.COMP    GAST/COM     SEG.DES(1)   SEG INM     SEG.MULT   SALDO CAP." & oImp.gPrnSaltoLinea
                lsCadImp = lsCadImp & " FECH. VCTO   N°CUOTA      CUOTA      CAPITAL     INT.COMP    GAST/COM     INT.GRACIA    SEG.DES(1)  SEG.MULT   SALDO CAP." & oImp.gPrnSaltoLinea
            Else
                If (pTipoRep = 2 And bSegDesGra) Then
                    lsCadImp = lsCadImp & " FECH. VCTO   N°CUOTA      CUOTA      CAPITAL     INT.COMP    INT.GRACIA    GAST/COM  SEG.DES(1)     SEG INM    SEG.MULT    SALDO CAP." & oImp.gPrnSaltoLinea
                ElseIf (pTipoRep = 1 And bSegDesGra) Then
                    lsCadImp = lsCadImp & " FECH. VCTO   N°CUOTA      CUOTA      CAPITAL     INT.COMP    INT.GRACIA    GAST/COM   SEG.DES(1)   SEG.MULT   SALDO CAP." & oImp.gPrnSaltoLinea
                ElseIf (pTipoRep = 2 And Not bSegDesGra) Then
                    lsCadImp = lsCadImp & " FECH. VCTO   N°CUOTA      CUOTA      CAPITAL     INT.COMP    INT.GRACIA    GAST/COM     SEG.DES     SEG INM    SEG.MULT     SALDO CAP." & oImp.gPrnSaltoLinea
                ElseIf (pTipoRep = 1 And Not bSegDesGra) Then
                    lsCadImp = lsCadImp & " FECH. VCTO   N°CUOTA      CUOTA      CAPITAL     INT.COMP    INT.GRACIA    GAST/COM     SEG.DES    SEG.MULT   SALDO CAP." & oImp.gPrnSaltoLinea
                End If
            End If
        End If
    Else
        If pTipoRep = 1 Then
            lsCadImp = lsCadImp & " FECH. VCTO   N°CUOTA      CUOTA      CAPITAL     INT.COMP    GAST/COM      SEG.DES    SEG.MULT   SALDO CAP." & oImp.gPrnSaltoLinea
        ElseIf pTipoRep = 2 Then
            lsCadImp = lsCadImp & " FECH. VCTO   N°CUOTA      CUOTA      CAPITAL     INT.COMP   INT.GRACIA    GAST/COM     SEG.DES    SEG.MULT   SALDO CAP." & oImp.gPrnSaltoLinea
        End If
    End If

    If bSimulador Then
        If (pTipoRep = 1) Then
            lsCadImp = lsCadImp & String(118, "-") & oImp.gPrnSaltoLinea
        ElseIf (pTipoRep = 2) Then
            lsCadImp = lsCadImp & String(133, "-") & oImp.gPrnSaltoLinea
        End If
    Else
        lsCadImp = lsCadImp & String(118, "-") & oImp.gPrnSaltoLinea
    End If
    
    lnIndice = 10:  lnLineas = 10
    nItem = 0
    
    If nTipoGracia = 6 And i = 0 Then
        lsCadImp = lsCadImp & " " & oImpre.ImpreFormat(Format(CDate(Format(pVigencia + nDiasAtraso, "dd/MM/yyyy")), "dd/mm/yy"), 10, 0) & Space(2) '** FECHA
        lsCadImp = lsCadImp & oImpre.ImpreFormat(0, 4, 0) & Space(17)
        lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(pMonto), 8, 2) & IIf(bSimulador, Space(16), Space(2)) 'WIOR 20150211 AGREGO IIf(bSimulador
        If (pTipoRep = 2) Then
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(nIntGraInicial), 8, 2) & IIf(bSimulador, Space(51), Space(42))
        Else
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(nIntGraInicial), 8, 2) & IIf(bSimulador, Space(38), Space(42))
        End If
        lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(pMonto), 8, 2)
        lsCadImp = lsCadImp & oImp.gPrnSaltoLinea
    End If
    
    For i = 0 To UBound(MatP) - 1

        nItem = nItem + 1
        lnLineas = lnLineas + 1
        lnIndice = lnIndice + 1
        
        lsCadImp = lsCadImp & " " & oImpre.ImpreFormat(Format(CDate(Format(MatP(i, 0), "dd/MM/yyyy")), "dd/mm/yy"), 10, 0) & Space(2) '** FECHA
        lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 1)), 4, 0) & Space(4) '** NRO CUOTA
        lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 2)), 8, 2) & Space(2) '** CUOTA
        sCuota = sCuota + MatP(i, 2)

        lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 3)), 8, 2) & Space(2) '** CAPITAL
        sCapital = sCapital + MatP(i, 3)

         'WIOR 20150211 ************************
        If bSimulador Then
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 4)), 8, 2) & Space(3) '** INTERES
            sInteres = sInteres + MatP(i, 4)
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 5)), 8, 2) & Space(1)  '** INT.GRACIA
            sInteresGra = sInteresGra + MatP(i, 5)
        Else
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 4)) + Val(MatP(i, 5)), 8, 2) & Space(2) '** INTERES
            sInteres = sInteres + MatP(i, 4) + MatP(i, 5)
        End If
        If Len(Trim(lsCtaCodLeasing)) = 0 Then
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 8)), 8, 2) & Space(1) '** GASTOS
                sGastos = sGastos + MatP(i, 8)
            If pTipoRep = 1 Then
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 6)), 8, 2) & Space(1) '** GASTOS SEG DES
                sSegDes = sSegDes + MatP(i, 6)
            ElseIf pTipoRep = 2 Then
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 6)), 8, 2) & Space(1) '** GASTOS SEG DES
                sSegDes = sSegDes + MatP(i, 6)
            End If
        Else
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 8)), 8, 2) & Space(1) '** GASTOS
            sGastos = sGastos + MatP(i, 8)
            
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 6)), 8, 2) & Space(1) '** GASTOS SEG DES
            sSegDes = sSegDes + MatP(i, 6)
        End If

        If pTipoRep = 2 Then
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 9)), 8, 2) & Space(1) '** GASTOS SEG INM
            sSegInm = sSegInm + MatP(i, 9)
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 13)), 8, 2) & Space(2) '** GASTOS SEG MULTI
            sSegMult = sSegMult + Val(MatP(i, 13))
        ElseIf pTipoRep = 1 Then
            lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 13)), 8, 2) & Space(1) '** GASTOS SEG MULTI
            sSegMult = sSegMult + Val(MatP(i, 13))
        End If

        If bSimulador Then
         lsCadImp = lsCadImp & Space(1)
        End If

        lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatP(i, 7)), 8, 2) & Space(2) & oImp.gPrnSaltoLinea '** SALDO K
        
 
        If lnIndice Mod 300 = 0 Then
            lsCadBuffer = lsCadBuffer & lsCadImp
            lsCadImp = ""
        End If
        
        If lnLineas >= 50 Then
            lnPage = lnPage + 1
            lsCadImp = lsCadImp & oImp.gPrnSaltoPagina
            lsCadImp = lsCadImp & oImp.gPrnSaltoLinea
            If pTipoRep = 1 Then
                lsCadImp = lsCadImp & nRepoCabeceraNEW(pTipoRep, cTitulo, cSubTitulo, lnPage, 118, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, IIf(psPersNom <> "", "Rapiflash con tu Plazo Fijo", psCtaCod), IIf(Trim(sNombre) = "", psPersNom, Trim(sNombre)), pnCuotas, pnTasaEfectivaAnual, pnTasaCostoEfectivoAnual, nTipoGracia, nDiasAtraso, psTpoSubPro, psTpoPro, psTpoMoneda, pdFechaVencPigno, 0, 0, 0, 0, "")
            ElseIf pTipoRep = 2 Then
                lsCadImp = lsCadImp & nRepoCabeceraNEW(pTipoRep, cTitulo, cSubTitulo, lnPage, 133, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, psCtaCod, Trim(sNombre), pnCuotas, pnTasaEfectivaAnual, pnTasaCostoEfectivoAnual, nTipoGracia, nDiasAtraso, psTpoSubPro, psTpoPro, psTpoMoneda, pdFechaVencPigno, 0, 0, 0, 0, "")
            End If
            
            lnLineas = 8
            lnIndice = lnIndice + 8
        End If
                
    Next

    If bSimulador Then
        lsCadImp = lsCadImp & String(120, "=") & oImp.gPrnSaltoLinea
    Else
        If pTipoRep <> 3 Then
        lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & String(120, "=") & oImp.gPrnSaltoLinea
        End If
    End If
    
    If bSimulador Then
        
        If pTipoRep = 2 Then
            lsCadImp = lsCadImp & "Totales :" & Space(12) & oImpre.ImpreFormat(sCuota, 8, 2) & Space(2) & oImpre.ImpreFormat(sCapital, 8, 2) & Space(2) & oImpre.ImpreFormat(sInteres, 8, 2) & Space(3) & oImpre.ImpreFormat(sInteresGra, 8, 2) & Space(1) & oImpre.ImpreFormat(sGastos, 8, 2) & Space(3) & oImpre.ImpreFormat(sSegDes, 6, 2) & Space(0) & oImpre.ImpreFormat(sSegInm, 9, 2) & Space(3) & oImpre.ImpreFormat(sSegMult, 6, 2) & oImp.gPrnSaltoLinea
        Else
        lsCadImp = lsCadImp & "Totales :" & Space(12) & oImpre.ImpreFormat(sCuota, 8, 2) & Space(2) & oImpre.ImpreFormat(sCapital, 8, 2) & Space(2) & oImpre.ImpreFormat(sInteres, 8, 2) & Space(3) & oImpre.ImpreFormat(sInteresGra, 8, 2) & Space(1) & oImpre.ImpreFormat(sGastos, 8, 2) & Space(3) & oImpre.ImpreFormat(sSegDes, 6, 2) & Space(3) & oImpre.ImpreFormat(sSegMult, 6, 2) & oImp.gPrnSaltoLinea
        End If
    Else
        If pTipoRep = 1 Then
            lsCadImp = lsCadImp & "Totales :" & Space(12) & oImpre.ImpreFormat(sCuota, 8, 2) & Space(2) & oImpre.ImpreFormat(sCapital, 8, 2) & Space(2) & oImpre.ImpreFormat(sInteres, 8, 2) & Space(2) & oImpre.ImpreFormat(sGastos, 8, 2) & Space(3) & oImpre.ImpreFormat(sSegDes, 6, 2) & Space(3) & oImpre.ImpreFormat(sSegMult, 6, 2) & oImp.gPrnSaltoLinea
        ElseIf pTipoRep = 2 Then
            lsCadImp = lsCadImp & "Totales :" & Space(12) & oImpre.ImpreFormat(sCuota, 8, 2) & Space(2) & oImpre.ImpreFormat(sCapital, 8, 2) & Space(2) & oImpre.ImpreFormat(sInteres, 8, 2) & Space(2) & oImpre.ImpreFormat(sGastos, 8, 2) & Space(3) & oImpre.ImpreFormat(sSegDes, 6, 2) & Space(0) & oImpre.ImpreFormat(sSegInm, 9, 2) & Space(3) & oImpre.ImpreFormat(sSegMult, 6, 2) & oImp.gPrnSaltoLinea
        End If
    End If
    
    If bSimulador Then
        lsCadImp = lsCadImp & String(120, "=") & oImp.gPrnSaltoLinea
    Else
        If pTipoRep <> 3 Then
        lsCadImp = lsCadImp & String(120, "=") & oImp.gPrnSaltoLinea
        End If
    End If

    'Mal Pagador Mi Vivienda
    sCapital = 0
    sInteres = 0
    sSegDes = 0
    sCuota = 0
    sSegInm = 0
     
    If pTipoRep = 2 Then
        If bEsMIVIVIENDAAnt Then
            lsCadImp = lsCadImp & oImp.gPrnSaltoPagina
            lsCadImp = lsCadImp & nRepoCabeceraNEW(pTipoRep, cTitulo, "MAL PAGADOR", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, psTpoSubPro, psTpoPro, psTpoMoneda, "", 0, 0, 0, 0, "")
         
            lsCadImp = lsCadImp & oImp.gPrnSaltoLinea
            
            lnIndice = 10:  lnLineas = 10
            nItem = 0
            For i = 0 To UBound(MatMalo) - 1
                nItem = nItem + 1
                lnLineas = lnLineas + 1
                lnIndice = lnIndice + 1
                
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Format(MatMalo(i, 0), "dd/MM/yyyy"), 10, 0) & Space(2) 'Fecha
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 1)), 4, 0) & Space(2) 'Nro Cuota
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 2)), 8, 2) & Space(2) 'Cuota
                sCuota = sCuota + MatMalo(i, 2)
                
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 3)), 8, 2) & Space(2)
                sCapital = sCapital + MatMalo(i, 3)
                
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 4)) + Val(MatMalo(i, 5)), 8, 2) & Space(2) 'Interes
                sInteres = sInteres + MatMalo(i, 4) + MatMalo(i, 5)
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 8)), 8, 2) & Space(2) 'Gastos
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 6)), 8, 2) & Space(2) 'Seg Des
                sSegDes = sSegDes + MatMalo(i, 6)
        
                If pTipoRep = 2 Then
                    lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 9)), 8, 2) & Space(2) ' Seg Inm
                    sSegInm = sSegInm + MatMalo(i, 9)
                End If
                
                lsCadImp = lsCadImp & oImpre.ImpreFormat(Val(MatMalo(i, 7)), 8, 2) & Space(2) & oImp.gPrnSaltoLinea 'Saldo Cap

                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oImp.gPrnSaltoPagina
                    lsCadImp = lsCadImp & oImp.gPrnSaltoLinea
                    lsCadImp = lsCadImp & nRepoCabeceraNEW(pTipoRep, cTitulo, "MAL PAGADOR", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, psTpoSubPro, psTpoPro, psTpoMoneda, "", 0, 0, 0, 0, "")
                 
                    lnLineas = 8
                    lnIndice = lnIndice + 8
                End If
                        
            Next
            
            lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & String(100, "=") & oImp.gPrnSaltoLinea
            
            lsCadImp = lsCadImp & "Totales :" & Space(10) & oImpre.ImpreFormat(sCuota, 8, 2) & Space(2) & oImpre.ImpreFormat(sCapital, 8, 2) & Space(2) & oImpre.ImpreFormat(sInteres, 8, 2) & Space(2) & oImpre.ImpreFormat(sGastos, 8, 2) & Space(2) & oImpre.ImpreFormat(sSegDes, 6, 2) & Space(2) & oImpre.ImpreFormat(sSegInm, 9, 2)
         End If
    End If
    
    If bSimulador Then
        If nTipoSegDes > 0 Then
            lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & "(1) El seguro desgravamen corresponde a la tasa de " & nTipoSegDes & " titular" & IIf(nTipoSegDes > 1, "(es).", ".")
        End If
        
        If bEnvioEstFis Then
            lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & "(2) Comisión correspondiente al envío de estado de cuenta físico."
        End If
    End If
    lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & ""
    lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & "Nota : "
    lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & "Las operaciones realizadas se encuentran afectas al pago del Impuesto por Transacciones Financieras (ITF) del 0.005%"
ReporteCalendarioNEW = lsCadBuffer & lsCadImp
 
 End Function

Public Function nRepoCabeceraNEW(ByVal pTipoRep As Integer, ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        Optional ByVal pnCodReporte As String, Optional ByVal pTCuota As String, Optional ByVal pInteres As Double, Optional ByVal pMonto As Double, _
        Optional ByVal pCuota As Double, Optional ByVal pPlazo As Integer, Optional ByVal pVigencia As Date, _
        Optional ByVal psCtaCod As String, Optional ByVal psNomCli As String, Optional ByVal pnNroCuotas As Integer, _
        Optional ByVal pnTasaEfectivaAnual As Double, Optional ByVal pnTasaCostoEfectivoAnual As Double, _
        Optional ByVal nTipoGracia As Integer, Optional ByVal pnPerGra As Integer = 0, _
        Optional ByVal psTpoSubPro As String, Optional ByVal psTpoPro As String, Optional ByVal psTpoMoneda As String, _
        Optional ByVal pdFechaVencPigno As Date, Optional ByVal pnTotalPigno As Double, Optional ByVal pnInteresPigno As Double, _
        Optional ByVal pnTasaNormal As Double, Optional ByVal pnTasaCPP As Double, Optional ByVal psTpoCliente As String) As String

        
Dim lsCadImp As String

    lsCadImp = nImprimeCabeceraReportesNEW(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)

    lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & String(pnAnchoLinea - 0, "-") & oImp.gPrnSaltoLinea
    If (pTipoRep <> 3) Then
    Select Case pnCodReporte
        Case "CalenPagos"
            lsCadImp = lsCadImp & "    PRODUCTO        : " & psTpoPro & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    SUB PRODUCTO    : " & psTpoSubPro & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    TIPO DE CUOTA   : " & pTCuota & IIf(pTipoRep = 1, Space(11), Space(16)) & "PLAZO     : " & Trim(oImpre.ImpreFormat(pnNroCuotas, 3, 0)) & oImp.gPrnSaltoLinea
           
            lsCadImp = lsCadImp & "    FRECUENCIA(Días): " & Trim(oImpre.ImpreFormat(pPlazo, 3, 0)) & oImp.gPrnSaltoLinea

            If pnPerGra > 0 Then
                lsCadImp = lsCadImp & "    PERIODO GRACIA  : " & Trim(oImpre.ImpreFormat(pnPerGra, 3, 0)) & oImp.gPrnSaltoLinea
            End If

            lsCadImp = lsCadImp & "    MONTO           : " & Trim(oImpre.ImpreFormat(pMonto, 8, 2)) & IIf(pTipoRep = 1, Space(39), Space(42)) & "FECHA DESEMBOLSO  : " & Format(pVigencia, "ddd, dd mmm yyyy") & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    MONEDA          : " & psTpoMoneda & oImp.gPrnSaltoLinea
            If pnTasaEfectivaAnual > 0 Then lsCadImp = lsCadImp & "    TASA EFECTIVA ANUAL       : " & oImpre.ImpreFormat(pnTasaEfectivaAnual, 3, 2) & " %" & oImp.gPrnSaltoLinea
            If pnTasaCostoEfectivoAnual > 0 Then lsCadImp = lsCadImp & "    TASA COSTO EFECTIVO ANUAL : " & oImpre.ImpreFormat(pnTasaCostoEfectivoAnual, 3, 2) & " %" & oImp.gPrnSaltoLinea
            
            lsCadImp = lsCadImp & String(pnAnchoLinea - 0, "-") & oImp.gPrnSaltoLinea
            
       End Select
       Else
            lsCadImp = lsCadImp & "    MONTO                    : " & Trim(oImpre.ImpreFormat(pMonto, 8, 2)) & Space(40) & "MONEDA         : " & psTpoMoneda & oImp.gPrnSaltoLinea 'ARLO20181010
            lsCadImp = lsCadImp & "    PLAZO VENCIMIENTO(DÍAS)  : " & 30 & Space(40) & "     INTERÉS        : " & Trim(oImpre.ImpreFormat(pnInteresPigno, 8, 2)) & oImp.gPrnSaltoLinea 'ARLO20181010
            If pnTasaEfectivaAnual > 0 Then
            lsCadImp = lsCadImp & "    TASA EFECTIVA ANUAL      : " & Trim(oImpre.ImpreFormat(pnTasaEfectivaAnual, 3, 2)) & " %" & Space(39) & "TIPO CLIENTE   : " & psTpoCliente & oImp.gPrnSaltoLinea
            End If
            lsCadImp = lsCadImp & "    FECHA DESEMBOLSO         : " & pVigencia & Space(35) & "  FECHA DE PAGO  : " & pdFechaVencPigno & oImp.gPrnSaltoLinea
            
            lsCadImp = lsCadImp & String(pnAnchoLinea - 0, "-") & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    MONTO                                 : " & Trim(oImpre.ImpreFormat(pnTotalPigno, 8, 2)) & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    CALIFICACION SBS              " & Space(10) & "TEA" & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    Clientes Con Calificación Normal      " & oImpre.ImpreFormat(pnTasaNormal, 3, 2) & " %" & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    Clientes Con Calificación CPP         " & oImpre.ImpreFormat(pnTasaCPP, 3, 2) & " %" & oImp.gPrnSaltoLinea
       End If
       
       
nRepoCabeceraNEW = lsCadImp
End Function
Public Function nImprimeCabeceraReportesNEW(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, _
        ByVal psFechaSis As String, ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        Optional ByVal psCodRepo As String) As String
        
  Dim lsCabe01 As String, lsCabe02 As String
  Dim lsCabe03 As String, lsCabe04 As String
  Dim lsCabRepo As String
  
  lsCabRepo = ""
  ' Cabecera 1
  lsCabe01 = oImpre.FillText(Trim(UCase(psNomCmac)), 45, " ")
  
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 43 - 28)
  
  lsCabe01 = lsCabe01 & "Pag.  : " & Str(pnPagina) & "  -  " & psCodUser & oImp.gPrnSaltoLinea

  lsCabe01 = lsCabe01 & oImpre.FillText(Trim(UCase(psNomAgencia)), 25, " ")
  
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 23 - 28)
  
  lsCabe01 = lsCabe01 & "Fecha : " & Format(psFechaSis & " " & Time, "dd/mm/yyyy hh:mm") & Chr$(10)

  lsCabe02 = psCodRepo & String(Int((pnAnchoLinea - Len(psTitulo)) / 2), " ") & psTitulo & Chr$(10)

  lsCabe03 = String(Int((pnAnchoLinea - Len(psSubTitulo)) / 2), " ") & psSubTitulo & Chr$(10)

  lsCabe04 = IIf(Len(psComenta) > 0, psComenta & Chr$(10), "")
  
  lsCabRepo = lsCabRepo & lsCabe01 & lsCabe02
  lsCabRepo = lsCabRepo & lsCabe03 & lsCabe04
  nImprimeCabeceraReportesNEW = lsCabRepo
End Function

Public Function ReporteResumenMiViviendaNEW(ByVal pTipoRep As Byte, ByVal MatP As Variant, ByVal MatMalo As Variant, _
        ByVal pTCuota As String, ByVal pInteres As Double, ByVal pMonto, _
        ByVal pCuota As Double, ByVal pPlazo As Integer, ByVal pVigencia As Date, _
        Optional ByVal pnTipoProceso As Integer = 0, Optional ByVal pMatDesPar As Variant = "", Optional GBITFAPLICAt As Boolean, Optional gnITFPorcentt As Double, Optional gnITFMontoMint As Double = 0, _
        Optional ByVal psCtaCod As String, Optional pnCuotas As Integer, _
        Optional ByVal pnTasaEfectivaAnual As Double, Optional ByVal pnTasaCostoEfectivoAnual As Double, Optional ByVal lsCtaCodLeasing As String = "", _
        Optional ByVal pnValorInmueble As Double, Optional ByVal pnCuotaInicial As Double, Optional ByVal pnBonoBuenPagador As Double, _
        Optional ByVal pnDiasGracia As Integer, Optional ByVal psSeguroDesgravamen As String, Optional ByVal psGastoAdministracion As Double, _
        Optional ByVal pnCuotMens As Double, Optional ByVal pnCuotMensBono As Double, _
        Optional ByVal psTpoSubPro As String, Optional ByVal psTpoPro As String, Optional ByVal psTpoMoneda As String, _
        Optional ByVal cTpoHipo As String) As String

Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer
Dim cTitulo As String
Dim i As Integer
Dim cCuota As String * 35
Dim sCuota As Double
Dim sCapital As Double
Dim sInteres As Double
Dim sGastos As Double
Dim sSegDes As Double
Dim sSql As String
Dim oConec As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim sNombre As String
Dim oParametros As COMDCredito.DCOMParametro
Dim nPorComiCorrespMN As Double
Dim nPorComiCorrespME As Double
Dim nMonMinComiCorrespMN As Double
Dim nMonMaxComiCorrespMN As Double
Dim nMonMinComiCorrespME As Double
Dim nFactorCorresp As Double
Dim nMonMinComi As Double
Dim nMonMaxComi As Double
Dim nComiCorresp As Double
oITF.gbITFAplica = GBITFAPLICAt
oITF.gnITFPorcent = gnITFPorcentt
oITF.gnITFMontoMin = gnITFMontoMint
 
cCuota = pTCuota
sCapital = 0
sInteres = 0

    nFactorCorresp = 1
    nMonMinComi = 0
    nMonMaxComi = 0
    If psCtaCod <> "" Then
        Set oParametros = New COMDCredito.DCOMParametro
        nPorComiCorrespMN = oParametros.RecuperaValorParametro(gColPParamPorComiCorresponsaliaMN)
        nPorComiCorrespME = oParametros.RecuperaValorParametro(gColPParamPorComiCorresponsaliaME)
        nMonMinComiCorrespMN = oParametros.RecuperaValorParametro(gColPParamMonMinCobroCorresponsaliaMN)
        nMonMaxComiCorrespMN = oParametros.RecuperaValorParametro(gColPParamMonMaxCobroCorresponsaliaMN)
        nMonMinComiCorrespME = oParametros.RecuperaValorParametro(gColPParamMonMinCobroCorresponsaliaME)
        If Mid$(psCtaCod, 9, 1) = "1" Then 'MN
            nFactorCorresp = nPorComiCorrespMN / 100
            nMonMinComi = nMonMinComiCorrespMN
            nMonMaxComi = nMonMaxComiCorrespMN
        Else
            nFactorCorresp = nPorComiCorrespME / 100
            nMonMinComi = nMonMinComiCorrespME
            nMonMaxComi = 9999999999#
        End If
    End If
    
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
            cTitulo = "SIMULADOR REFERENCIAL DE CREDITO HIPOTECARIO"
        ElseIf pnTipoProceso = 1 Then
            cTitulo = "SUGERENCIA DE ANALISTA - CALENDARIO DE PAGOS MI VIVIENDA"
        ElseIf pnTipoProceso = 2 Then
            cTitulo = "APROBACION DE CREDITO - CALENDARIO DE PAGOS MI VIVIENDA"
        End If
    End If

    lnLineas = 0
    lnPage = 1

    'obteniendo titular
    sSql = "Select Pers.cPersNombre"
    sSql = sSql & " From ProductoPersona PP"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Where PP.cCtaCod='" & psCtaCod & "' And PP.nPrdPersRelac=20"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        sNombre = rs!cPersNombre
    End If
    Set rs = Nothing
    
    'CABECERA
    If pTipoRep = 1 Then
        lsCadImp = lsCadImp & nRepoCabeceraNEW(pTipoRep, cTitulo, "", lnPage, 120, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, psCtaCod, Trim(sNombre), pnCuotas, pnTasaEfectivaAnual, pnTasaCostoEfectivoAnual, 0, 0, psTpoSubPro, psTpoPro, psTpoMoneda, "01-06-2018", 0, 0, 0, 0, "")
    ElseIf pTipoRep = 2 Then
        lsCadImp = lsCadImp & nRepoCabeceraResumenNEW(cTitulo, "BUEN PAGADOR", lnPage, 120, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia, psCtaCod, Trim(sNombre), pnCuotas, pnTasaEfectivaAnual, pnTasaCostoEfectivoAnual, pnValorInmueble, pnCuotaInicial, pnBonoBuenPagador, pnDiasGracia, psSeguroDesgravamen, psGastoAdministracion, pnCuotMens, pnCuotMensBono, psTpoSubPro, psTpoPro, psTpoMoneda, cTpoHipo)
    End If
    
ReporteResumenMiViviendaNEW = lsCadBuffer & lsCadImp
End Function

Public Function nRepoCabeceraResumenNEW(ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        ByVal pnCodReporte As String, ByVal pTCuota As String, ByVal pInteres As Double, ByVal pMonto As Double, _
        ByVal pCuota As Double, ByVal pPlazo As Integer, ByVal pVigencia As Date, _
        Optional ByVal psCtaCod As String, Optional ByVal psNomCli As String, Optional ByVal pnNroCuotas As Integer, _
        Optional ByVal pnTasaEfectivaAnual As Double, Optional ByVal pnTasaCostoEfectivoAnual As Double, _
        Optional ByVal pnValorInmueble As Double, Optional ByVal pnCuotaInicial As Double, Optional ByVal pnBonoBuenPagador As Double, _
        Optional ByVal pnDiasGracia As Integer, Optional ByVal psSeguroDesgravamen As String, Optional ByVal psGastoAdministracion As Double, _
        Optional ByVal pnCuotMens As Double, Optional ByVal pnCuotMensBono As Double, _
        Optional ByVal psTpoSubPro As String, Optional ByVal psTpoPro As String, Optional ByVal psTpoMoneda As String, _
        Optional ByVal cTpoHipo As String) As String
        
    Dim lsCadImp As String

    psSubTitulo = ""
    lsCadImp = nImprimeCabeceraReportes(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & String(pnAnchoLinea - 10, "-") & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    Select Case pnCodReporte
        Case "CalenPagos"
            lsCadImp = lsCadImp & "    PRODUCTO              : " & psTpoPro & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    SUB PRODUCTO          : " & psTpoSubPro & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    INCENDIO              : " & cTpoHipo & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    VALOR DEL INMUEBLE    : " & "S/. " & Trim(oImpre.ImpreFormat(pnValorInmueble, 8, 2)) & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    MONTO                 : " & "S/.  " & Trim(oImpre.ImpreFormat((pnValorInmueble - pnCuotaInicial) - pnBonoBuenPagador, 8, 2)) & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    FRECUENCIA            : " & Trim(oImpre.ImpreFormat((CInt(pnNroCuotas) * 30) / 360, 3, 0)) & " Años" & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    PLAZO                 : " & Trim(oImpre.ImpreFormat(pnNroCuotas, 3, 0)) & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    PERIODO DE GRACIA     : " & Trim(oImpre.ImpreFormat(pnDiasGracia, 3, 0)) & " dias" & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    SEGURO DE DESGRAVAMEN : " & Trim(psSeguroDesgravamen) & Space(16) & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    GASTOS                : " & "S/. " & Trim(oImpre.ImpreFormat(psGastoAdministracion, 8, 2)) & oImp.gPrnSaltoLinea
            
            If pnTasaEfectivaAnual > 0 Then lsCadImp = lsCadImp & "    TASA EFECTIVA ANUAL   : " & Trim(oImpre.ImpreFormat(pnTasaEfectivaAnual, 3, 2)) & " %" & oImp.gPrnSaltoLinea
            If pnTasaEfectivaAnual > 0 Then lsCadImp = lsCadImp & "    TASA SEG DESGRAV MEN  : " & IIf(Left(psSeguroDesgravamen, 1) = "I", "0.0429", "0.0772") & " %" & oImp.gPrnSaltoLinea
            If pnTasaEfectivaAnual > 0 Then lsCadImp = lsCadImp & "    TASA SEG INMUEBL MEN  : " & "0.022 %" & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & "    FECHA DE SIMULACION   : " & Format(pVigencia, "ddd, dd mmm yyyy") & oImp.gPrnSaltoLinea
    
            lsCadImp = lsCadImp & oImp.gPrnSaltoLinea & String(pnAnchoLinea - 10, "-") & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
            If pnTasaCostoEfectivoAnual > 0 Then lsCadImp = lsCadImp & "    TASA COSTO EFECTIVO ANUAL        : " & Trim(oImpre.ImpreFormat(pnTasaCostoEfectivoAnual, 3, 2)) & " %" & oImp.gPrnSaltoLinea
            lsCadImp = lsCadImp & String(pnAnchoLinea - 10, "-") & oImp.gPrnSaltoLinea
            
       End Select
       
nRepoCabeceraResumenNEW = lsCadImp
End Function
'ARLO20180723 ERS037-2018

'---MARG20180513 pag.ant.----------------------------------------
Private Sub CalendarioCuotaFijaNew(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
                Optional ByVal pMatDesPar As Variant = "", Optional ByVal pnNumMes As Integer = 1, Optional ByVal pbMiViv As Boolean = False, _
                Optional ByVal pnCuotaIni As Integer = 0, Optional ByVal pnCuotaFin As Integer = 0, Optional ByVal pnDiaFijo2 As Integer = 0, _
                Optional pnMontoCapInicial As Double = 0, _
                Optional ByVal psCtaCod As String, Optional ByVal psMensFecAprob As Date, Optional ByVal lnDescripCuota As Variant = "", _
                Optional ByVal pnMontoGra As Double = 0, Optional ByVal psIncluyeIntCap As String = "S", _
                Optional ByVal pnCuotaBalon As Integer = 0, _
                Optional ByVal pbReprogConvenio As Boolean = False) 'MAVM 11092010 Se agrego el par: psCtaCod (Prepago)
                'MADM 20110409 - FECHA APRO - lnDescripCuota
                'WIOR 20131111 AGREGO pnCuotaBalon
                'WIOR 20140514 AGREGO pbReprogConvenio
                
Dim nSaldoCapital As Double
Dim dDesembolso As Date
Dim dFecAproGracia As Date
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
Dim nPeriodoPrimerPF As Integer 'MADM 20110503

'MAVM 20130209 ***
Dim nSumCapitalGra As Double
Dim nSaldoCapitalGra As Double
'***
Dim nNroCuotasAfec As Integer 'WIOR 20131127

Dim dFechaCuotaPend As Date  'add by marg2080523
Dim dDesembolsoNuevo As Date 'add by marg2080523
Dim nCuotasPendientes As Integer 'add by marg2080524

nSumCapital = 0
nPeriodoPrimerPF = 0

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
        
        'MADM 20120221-PAGOS -----------------------------------------------------------
        If psMensFecAprob = "12:00:00 AM" And psCtaCod <> "" Then
            Dim oCredT As COMNCredito.NCOMCredito
            Dim RMantPP As ADODB.Recordset
            Dim DfechaPre As Date
            Dim PlazoPre As Integer
            Dim indPre As Boolean
            
            Set oCredT = New COMNCredito.NCOMCredito
            Set RMantPP = oCredT.RecuperaValidacionCredMantPrepago(psCtaCod, Format(pdFecDesemb, "dd/mm/yyyy"))
            Set oCredT = Nothing
            indPre = False
            
            If Not RMantPP.BOF And Not RMantPP.EOF Then
                indPre = True
                DfechaPre = RMantPP!dVenc 'CC.nCuota , Ce.nNroProxCuota, CC.nNroCalen,
                PlazoPre = RMantPP!DifPlazo
                dFechaCuotaPend = RMantPP!dFechaCuotaPend 'add by marg20180523
                dDesembolsoNuevo = RMantPP!dDesembolsoNuevo 'add by marg20180523
                nCuotasPendientes = RMantPP!nCuotasPendientes 'add by marg20180524
                pnCuotaFin = nCuotasPendientes - 1
                RMantPP.Close
            End If
            Set RMantPP = Nothing
        End If
        '-------------------------------------------------------------------------------
        
        nSaldoCapital = pnMonto
        nSaldoCapitalGra = pnMontoGra 'MAVM 20130209
        dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
       
        If pnTipoPeriodo = PeriodoFijo Then
            Set oCredito = New NCOMCredito
            'For i = 0 To pnNroCuotas - 1
            For i = pnCuotaIni To pnCuotaFin
                
            'MADM 20110518 - No entre cuando es aprobacion o PAGO
              If i = 0 And psMensFecAprob <> "12:00:00 AM" Then
                'MADM 20110420 - Verifica Si tiene Gracia / Diferencia es -
                    If i = 0 And (dDesembolso > psMensFecAprob) Then
                          If lnDescripCuota(0) = "0" Then
                            nPeriodoPrimerPF = pnPeriodo - CInt(DateDiff("d", psMensFecAprob, dDesembolso))
                          Else
                                If lnDescripCuota(1) <> "0" Then
                                    nPeriodoPrimerPF = pnPeriodo + Val(lnDescripCuota(1))
                                End If
                          End If
                    End If
              End If
              'END MADM
               'MADM 20120221 ----------------------------------------------------------
                If indPre = True Then
                    Calendario(i).dFecha = dDesembolso + IIf(i = 0, PlazoPre, pnPeriodo)
                Else
                    Calendario(i).dFecha = dDesembolso + IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo)
                End If
                '-------------------------------------------------------------------------
                Calendario(i).NroCuota = i + 1
                Calendario(i).IntGra = 0#
                Calendario(i).Gasto = 0#
                'MADM 20110518 - Periodo
                If indPre = True Then
                    Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, IIf(i = 0, PlazoPre, pnPeriodo), nSaldoCapital)
                    Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, IIf(i = 0, PlazoPre, pnPeriodo))
                Else
                     'WIOR 20131111 **********************************
                    If i < pnCuotaBalon Then
                        nSaldoCapital = pnMonto
                    End If
                    'WIOR FIN ***************************************
                    Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), nSaldoCapital)

                    'MAVM 20130430***
                    If pnTipoGracia = 6 Then
                        Calendario(i).IntCompGra = oCredito.MontoIntPerDias(pnTasaInt, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), nSaldoCapitalGra)
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
                    'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo)
                    
                    'MAVM 20130430***
                    If pnTipoGracia = 6 Then
                        Calendario(i).CuotaGra = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMontoGra, pnPeriodo)
                    End If
                    '***
                    
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
                nSaldoCapital = CDbl(Format(nSaldoCapital, "#0.00"))
                
                'MAVM 20130430***
                If pnTipoGracia = 6 Then
                    nSumCapitalGra = nSumCapitalGra + Calendario(i).IntGra
                    Calendario(i).SaldoCapGra = nSaldoCapitalGra - Calendario(i).IntGra
                    nSaldoCapitalGra = nSaldoCapitalGra - Calendario(i).IntGra
                End If
                
                'MADM 20120221 ----------------------------------------------------------
                If indPre = True Then
                    dDesembolso = dDesembolso + IIf(i = 0, PlazoPre, pnPeriodo)
                Else
                    dDesembolso = dDesembolso + IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo)
                End If
                '-------------------------------------------------------------------------
                'MADM 20110518 - + IIf(nPeriodoPrimerPF <> 0, nPeriodoPrimerPF, pnPeriodo)
                
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
                nPeriodoPrimerPF = 0
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
                        
                        'MAVM 20110108 ***
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
                            
                            'comment by marg20180523-----
'                            If CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy")) - dDesembolso < 5 Then
'                                nMes = nMes + pnNumMes
'                            End If
                            'end marg-----------
                        End If
                        '***
                        
'                        If nMes > 12 Then
'                            nAnio = nAnio + 1
'                            nMes = 1
'                        End If

                    '*****************************************
                    Else
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
                            If nDia > 30 Then
                                If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
                                    nMes = nMes + pnNumMes
                                End If
                            End If
'                        End If


                    End If
                    
                    'add by marg20180514 pag.ant.---
                    If nMes > 12 Then
                        nAnio = nAnio + 1
                        nMes = nMes - 12
                    End If
                    'end marg-----------------------
                    
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
                    'By Capi 06122008 formateo de fecha
                    
                    'MAVM 20100820 ***
                    If psCtaCod <> "" Then
                        If i = 0 Then
                        'comment by mar20180523-------------------------
'                            If CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy")) - dDesembolso < 20 Then
'                                pnDiasGracia = CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy")) - dDesembolso
'                                Dim pnMontPeridExtra As Double
'                                pnMontPeridExtra = oCredito.MontoIntPerDias(pnTasaInt, pnDiasGracia, nSaldoCapital)
'                                dDesembolso = CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy"))
'                                nMes = nMes + pnNumMes
'                                If nMes > 12 Then
'                                    nAnio = nAnio + 1
'                                    nMes = 1
'                                End If
'                                If nMes = 2 Then
'                                    If nDia > 28 Then
'                                        If nAnio Mod 4 = 0 Then
'                                            nDia = 29
'                                        Else
'                                            nDia = 28
'                                        End If
'                                    End If
'                                Else
'                                    If nDia > 30 Then
'                                        If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
'                                            nDia = 30
'                                        End If
'                                    End If
'                                End If
'                            End If
                            'end marg----------------------------
                            'add by marg20180523--------------
                            'pnDiasGracia = CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy")) - dDesembolso EAAS20181029 CORRECCION SEGUN CORREO
                            Dim pnMontPeridExtra As Double
                            pnMontPeridExtra = oCredito.MontoIntPerDias(pnTasaInt, pnDiasGracia, nSaldoCapital)
                            nDia = Day(dFechaCuotaPend)
                            nMes = Month(dFechaCuotaPend)
                            nAnio = Year(dFechaCuotaPend)
                            If (dFechaCuotaPend > dDesembolso) And (dDesembolsoNuevo > dDesembolso) Then
                                   dDesembolso = dDesembolsoNuevo
                            End If
                            'end marg-------------------------
                        End If
                    End If
                        '***
                    
                    dFecTemp = CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy"))
                     
                    'If (dFecTemp > pdFecDesemb + CDate(pnDiasGracia) And bProxMes) Then
                    'If (dFecTemp > dDesembolso And bProxMes) Then
                    '    Calendario(i).dFecha = DateAdd("m", -Fix(pnDiasGracia / 30), dFecTemp)
                    'Else
                        Calendario(i).dFecha = dFecTemp
                    'End If
                    Calendario(i).NroCuota = i + 1
                    Calendario(i).IntGra = 0#
                    Calendario(i).Gasto = 0
                    'WIOR 20131111 **********************************
                    If i < pnCuotaBalon Then
                        nSaldoCapital = pnMonto
                    End If
                    nNroCuotasAfec = pnNroCuotas - pnCuotaBalon
                    'WIOR FIN ***************************************
                    'If i = 0 Then
                    If i = pnCuotaIni Then
                    '11-05-2006
                        If pbCuotaFijaFechaFija Then
                            'Set oCredito = New COMNCredito.NCOMCredito
                            'nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas)
                            'dFecTemp - dDesembolso +
                            'If pnDiasGracia = 0 Then
                                'nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas) 'WIOR 20131127 COMENTÓ
                                nMontoCuotaTemp = oCredito.CFijaFechaFijaNew(pnTasaInt, nNroCuotasAfec, pnMonto, pnPeriodo, dDesembolso, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas, dFechaCuotaPend) 'WIOR 20131127 COMENTÓ 'EAAS20181029 CORRECCION SEGUN CORREO SE CAMBIO PARAMETRO pdFecDesemb A dDesembolso y se agrego dFechaCuotaPend
                            'Else
                            '    nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, dFecTemp - dDesembolso + pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas)
                            'End If
                            'Set oCredito = Nothing
                            
                            'MAVM 20130209***
                            If pnTipoGracia = 6 Then
                                'nMontoGraciaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMontoGra, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas)'EAAS20181031 COMENTO
                                nMontoGraciaTemp = oCredito.CFijaFechaFijaNew(pnTasaInt, pnNroCuotas, pnMontoGra, pnPeriodo, dDesembolso, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv, pnDiaFijo2, pnNroCuotas, dFechaCuotaPend) 'EAAS20181029 CORRECCION SEGUN CORREO SE CAMBIO PARAMETRO pdFecDesemb A dDesembolso y se agrego dFechaCuotaPend
                            End If
                            '***
                            
                        End If
                    '**************************************
                        If pbCuotaFijaFechaFija And pbMiViv Then
                            dDesembolso = dDesembolso - pnDiasGracia
                        End If
                        Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(i).dFecha), nSaldoCapital)
                        
                        'MAVM 20130209***
                        If pnTipoGracia = 6 Then
                            Calendario(i).IntCompGra = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(i).dFecha), pnMontoGra)
                        End If
                        
                        If Not pbCuotaFijaFechaFija Then
                            If DateDiff("d", dDesembolso, Calendario(i).dFecha) = 0 And Len(Trim(psCtaCod)) = 0 Then
                                'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, 30) 'WIOR 20131111 COMENTÓ
                                'WIOR 20131111 **********************************
                                If i < pnCuotaBalon Then
                                    Calendario(i).Cuota = Calendario(i).IntComp
                                Else
                                    Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, nNroCuotasAfec, pnMonto, 30)
                                End If
                                'WIOR FIN ***************************************
                                
                            Else
                                'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, DateDiff("d", dDesembolso, Calendario(i).dFecha)) 'WIOR 20131111 COMENTÓ
                                'WIOR 20131111 **********************************
                                If i < pnCuotaBalon Then
                                    Calendario(i).Cuota = Calendario(i).IntComp
                                Else
                                    Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, nNroCuotasAfec, pnMonto, DateDiff("d", dDesembolso, Calendario(i).dFecha))
                                End If
                                'WIOR FIN ***************************************
                                
                            End If
                        Else
                        
                            'Calendario(i).Cuota = nMontoCuotaTemp 'WIOR 20131111 COMENTÓ
                            'WIOR 20131111 **********************************
                            If i < pnCuotaBalon Then
                                Calendario(i).Cuota = Calendario(i).IntComp
                            Else
                                Calendario(i).Cuota = nMontoCuotaTemp
                            End If
                            'WIOR FIN ***************************************
                            
                             'MAVM 20130209***
                            If pnTipoGracia = 6 Then
                                Calendario(i).CuotaGra = nMontoGraciaTemp
                            End If
                            '***
                            
                        End If
                        
                    Else
                        Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha), nSaldoCapital)
                        'MAVM 20130209***
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
                            'MAVM 20130209***
                            If pnTipoGracia = 6 Then
                                Calendario(i).CuotaGra = nMontoGraciaTemp
                            End If
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
                        'MAVM 20130209***
                        If pnTipoGracia = 6 Then
                            Calendario(i).IntGra = nSaldoCapitalGra
                        End If
                        '***
                        If (Calendario(i).IntComp + Calendario(i).Captital) > Calendario(i).Cuota Then
                            Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                        Else
                            Calendario(i).IntComp = Calendario(i).Cuota - Calendario(i).Captital
                             'MAVM 20130209***
                            If pnTipoGracia = 6 Then
                                Calendario(i).IntCompGra = Calendario(i).CuotaGra - Calendario(i).IntGra
                            End If
                            '***
                        End If
                    Else
                        Calendario(i).Captital = Calendario(i).Cuota - Calendario(i).IntComp
                        'MAVM 20130209***
                        If pnTipoGracia = 6 Then
                            Calendario(i).IntGra = Calendario(i).CuotaGra - Calendario(i).IntCompGra
                        End If
                        '***
                    End If
                    
                    nSumCapital = nSumCapital + Calendario(i).Captital
                    Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                    nSaldoCapital = nSaldoCapital - Calendario(i).Captital
                    
                    'MAVM 20130209***
                    If pnTipoGracia = 6 Then
                        nSumCapitalGra = nSumCapitalGra + Calendario(i).IntGra
                        Calendario(i).SaldoCapGra = nSaldoCapitalGra - Calendario(i).IntGra
                        nSaldoCapitalGra = nSaldoCapitalGra - Calendario(i).IntGra
                    End If
                    
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
    If (pnDiasGracia > 0 _
       And pnCuotaIni = 0) Or (pbReprogConvenio And pnCuotaIni = 0) Then 'WIOR 20140514 AGREGO (pbReprogConvenio And pnCuotaIni = 0)
        If pnTipoGracia = PrimeraCuota Then
            'ReDim Preserve Calendario(pnNroCuotas + 1)
            'For i = pnNroCuotas To 1 Step -1
            '    Calendario(i) = Calendario(i - 1)
            '    Calendario(i).NroCuota = Calendario(i).NroCuota + 1
            'Next i
            'Calendario(0).Dfecha = pdFecDesemb + pnDiasGracia
            'Calendario(0).NroCuota = 1
            'Calendario(0).IntGra = MatGracia(0)
            'Calendario(0).Gasto = 0
            'Calendario(0).IntComp = 0#
            'Calendario(0).Cuota = Calendario(0).IntGra
            'Calendario(0).Captital = 0#
            'Calendario(0).SaldoCap = pnMonto
            
            'MAVM 20130209 ***
            Calendario(0).IntGra = MatGracia(0)
            Calendario(0).Cuota = Calendario(0).Cuota + MatGracia(0)
            '***
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
        
        'MAVM 20120209 ***
        If pnTipoGracia = gColocTiposGraciaCapitalizada And psIncluyeIntCap = "S" Then
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
            Calendario(0).SaldoCap = pnMonto '+ pnMontoGra
            Calendario(0).SaldoCapGra = 0#
            Calendario(0).CuotaGra = 0#
            Calendario(0).IntCompGra = 0#
        End If
        '***
        
        'MAVM 20100820 ***
        If pnMontPeridExtra <> 0 Then
            Dim nCuotaNormalPago As Double
            Dim nFactor As Double
            Dim oCred As COMDCredito.DCOMCredito
            Set oCred = New COMDCredito.DCOMCredito
            
            nCuotaNormalPago = oCred.CuotaNormalPagoSinGastos(psCtaCod)
            
            If (nCuotaNormalPago > Calendario(0).Cuota) And ((nCuotaNormalPago - Calendario(0).Cuota) * pnNroCuotas > pnMontPeridExtra) Then
                nFactor = nCuotaNormalPago - Calendario(0).Cuota
                        
                For i = 0 To pnNroCuotas - 1
                    If pnMontPeridExtra > 0 Then
                        If pnMontPeridExtra > nFactor Then
                            Calendario(i).IntComp = Calendario(i).IntComp + nFactor
                            Calendario(i).Cuota = Calendario(i).Cuota + nFactor
                            pnMontPeridExtra = pnMontPeridExtra - nFactor
                        Else
                            Calendario(i).IntComp = Calendario(i).IntComp + pnMontPeridExtra
                            Calendario(i).Cuota = Calendario(i).Cuota + pnMontPeridExtra
                            pnMontPeridExtra = pnMontPeridExtra - pnMontPeridExtra
                        End If
                    End If
                Next i
        
            Else
                nFactor = Round(pnMontPeridExtra / (pnNroCuotas), 2)
                For i = 0 To pnNroCuotas - 1
                    If i <> 0 Then
                        Calendario(i).IntComp = Calendario(i).IntComp + nFactor
                        Calendario(i).Cuota = Calendario(i).Cuota + nFactor
                    End If
                Next i
            End If

        End If
        '***
    
    End If
End Sub
'---END MARG--------------------------------------------
'->*****LUCV20180601, Comentó según ERS022-2018 [Adecuaciones del met. CalendarioCuotaFija]
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
                        Optional ByVal psCtaCod As String, _
                        Optional ByVal psMensFecAprob As Date, _
                        Optional ByVal lnDescripCuota As Variant = "", _
                        Optional ByVal pnMontoGraAnt As Double = 0, _
                        Optional ByVal pnTasaSegDes As Double = 0, _
                        Optional ByRef pMatCalendSegDes As Variant, _
                        Optional ByVal pnExoSeguroDesgravamen As Integer = 0, _
                        Optional ByVal pnMontoPoliza As Double = 0, _
                        Optional ByVal pnTasaSegInc As Double = 0, _
                        Optional ByVal pbReduccionNroCuotaPagAnt As Boolean = False, _
                        Optional ByRef pvArrayDatos As Variant = Nothing)
                        'MADM 20110409 - FECHA APRO - lnDescripCuota
                        'LUCV20180601, Agregó: pnTasaSegDes, pMatCalendSegDes, pnExoSeguroDesgravamen, pnMontoPoliza, pnTasaSegInc, pArrayDatos. Según ERS022-2018
                        'pnMontoGraAnt, para gracia pendiente
                
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
    'Pagos Anticipados
    Dim dFechaCuotaPend As Date
    Dim dDesembolsoNuevo As Date
    Dim nCuotasPendientes As Integer 'add by marg2080524
    Dim nPeriodoPrimerPF As Integer
    Dim oDCOMCalendario As COMDCredito.DCOMCalendario
    Set oDCOMCalendario = New COMDCredito.DCOMCalendario
    'Pag. Ant. Reducción de Nro. Cuotas
    Dim nMontoCuotaDesembolso As Double
    Dim oDCOMCredito As COMDCredito.DCOMCredito
    'Reprogramados
    Dim rsDatosReprogramados As ADODB.Recordset
    Dim bEsReprogramado As Boolean
    Dim dFecVencUltimoPago As Date
    Dim iMat As Integer
    Dim nInteresCompAFecha As Double
    Dim nInteresGraciaAFecha As Double
    Dim nInteresMoratorio As Double
    Dim nInteresCompVencAFecha As Double
    Dim nSegDesgAnt As Double
    Dim nSegIncAnt As Double
    Dim nSegIncGraciaAnt As Double
    Dim nTotalIntLiquidados As Double
    Dim nDiasCalculo As Integer 'RIRO 20200904 Liquidacion
    Dim nTasaEspCovid As Double 'Joep20200910 Tasa Especial covid
    Dim nOpCovid As Integer 'Joep20200910 Tasa Especial covid
    'Dim dFecVencIni As Date
    '<-***** Fin LUCV20180601
    
    'RIRO 20210212 ********************
    Dim nPolizaMen As Double        ' Póliza mensual sin el prorrateo, concepto 1231
    Dim nPolizaCuotReprog As Double ' Póliza de la primera cuota, concepto 1231
    Dim nPolizaProrrateo As Double      ' Póliza prorrateada en cada cuota, concepto 1279
    'END RIRO
    
    '1.- Inicializa variables
    nPeriodoPrimerPF = 0
    nIntPeriodoGracia = 0: nGraciaGenerada = 0: nGraciaNoPagada = 0
    nVASKU = 0: nSKU = 1.5
    nSumCapital = 0
    nValorAjusteIteracion = 0: nValorAjusteIteracionTotal = 0: nNumIteracion = 0: nContadorSKU = 0
    
    nSKPenultimaCuota = 0
    nPrimaPerGracia = 0
    iMat = 0: bEsReprogramado = False: nInteresCompAFecha = 0: nInteresGraciaAFecha = 0: nInteresMoratorio = 0: nInteresCompVencAFecha = 0
    nSegDesgAnt = 0: nSegIncAnt = 0: nSegIncGraciaAnt = 0
    nOpCovid = 0: nTasaEspCovid = 0 'Joep20200910 Tasa Especial covid
    
    '2.- Fecha de desembolso + los días del periodo de gracia
    dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
    nMes = Month(dDesembolso): nAnio = Year(dDesembolso)
    
    '* Para: Pago Anticipado
    If psMensFecAprob = "12:00:00 AM" And psCtaCod <> "" Then
        Dim rsDatosPagoAnticipado As ADODB.Recordset
        Dim DfechaPre As Date
        Dim PlazoPre As Integer
        Dim bEsPagoAnticipado As Boolean
        'JOEP20211020 Mejora de pago anticipa
        Dim nDiasFijo As Integer
        nDiasFijo = 0
        'JOEP20211020 Mejora de pago anticipa
        Set rsDatosPagoAnticipado = New ADODB.Recordset
        Set oNCOMCredito = New COMNCredito.NCOMCredito
        Set rsDatosPagoAnticipado = oNCOMCredito.RecuperaValidacionCredMantPrepago(psCtaCod, Format(pdFecDesemb, "dd/mm/yyyy"), pnNroCuotas) 'RIRO20210222 pnNroCuotas
        Set oNCOMCredito = Nothing
        bEsPagoAnticipado = False

        If Not rsDatosPagoAnticipado.BOF And Not rsDatosPagoAnticipado.EOF Then
            bEsPagoAnticipado = True
            DfechaPre = rsDatosPagoAnticipado!dVenc
            PlazoPre = rsDatosPagoAnticipado!DifPlazo
            dFechaCuotaPend = rsDatosPagoAnticipado!dFechaCuotaPend     'Fecha de Prox. Cuota a pagar
            dDesembolsoNuevo = rsDatosPagoAnticipado!dDesembolsoNuevo   'Fecha Inicial, Anterior a dFechaCuotaPend
            nCuotasPendientes = rsDatosPagoAnticipado!nCuotasPendientes 'Cantidad de Cuotas Pend.
            pnCuotaFin = nCuotasPendientes - 1
            nPrimaPerGracia = rsDatosPagoAnticipado!nPrimaPerGracia
            
            'RIRO 20210222 *****
            nPolizaMen = rsDatosPagoAnticipado!nPolizaMen
            nPolizaCuotReprog = rsDatosPagoAnticipado!nPolizaMen
            nPolizaProrrateo = rsDatosPagoAnticipado!nPolizaProrrateo
            'END RIRO **********
            
            'JOEP20211020 Mejora de pago anticipa
            nDiasFijo = rsDatosPagoAnticipado!nDiasFijo
            If IsArray(pvArrayDatos) Then
                For iMat = 0 To UBound(pvArrayDatos) - 1
                    Select Case iMat
                        Case 1: nIntPeriodoGracia = pvArrayDatos(iMat)
                    End Select
                Next iMat
            End If
            'JOEP20211020 Mejora de pago anticipa
            
            rsDatosPagoAnticipado.Close
            
            'Cuota Objetivo (PagoAnt: Reduccion Plazo)
            Set oDCOMCredito = New COMDCredito.DCOMCredito
            If pbReduccionNroCuotaPagAnt Then
                nMontoCuotaDesembolso = oDCOMCredito.CuotaFijaMensualPago(psCtaCod)
                pnCuotaFin = pnNroCuotas - 1
            End If
            Set oDCOMCredito = Nothing
        End If
        Set rsDatosPagoAnticipado = Nothing
    End If
    
    '** Fechas: Periodo Fijo / Fecha Fija (**Pago Anticipado)
    If pnTipoPeriodo = PeriodoFijo Then
           'Para Pagos Anticipados
             If i = 0 And psMensFecAprob <> "12:00:00 AM" Then
               'Verifica Si tiene Gracia / Diferencia es -
                 If i = 0 And (dDesembolso > psMensFecAprob) Then
                    If lnDescripCuota(0) = "0" Then
                        nPeriodoPrimerPF = pnPeriodo - CInt(DateDiff("d", psMensFecAprob, dDesembolso))
                    Else
                        If lnDescripCuota(1) <> "0" Then
                            nPeriodoPrimerPF = pnPeriodo + Val(lnDescripCuota(1))
                        End If
                    End If
                 End If
             End If
    Else
        nMes = Month(dDesembolso): nAnio = Year(dDesembolso)
    End If
    
    '*** Para: Reprogramación de créditos
    Set oNCOMCredito = New COMNCredito.NCOMCredito
    Set rsDatosReprogramados = New ADODB.Recordset
    Set rsDatosReprogramados = oNCOMCredito.RecuperaDatosReprogramacionCalendario(psCtaCod)
    If Not rsDatosReprogramados.BOF And Not rsDatosReprogramados.EOF And psCtaCod <> "" Then
        bEsReprogramado = True
        dFecVencUltimoPago = rsDatosReprogramados!dFecVencUltimoPago 'Fecha ultima cuota pagada
        dFechaCuotaPend = rsDatosReprogramados!dFechaCuotaPend       'Fecha Cuota Pendiente
        dDesembolso = rsDatosReprogramados!dFecVencReprog            'Fecha Cuota Pendiente + "n" Dias Reprog.
        pnDiaFijo = Day(dDesembolso): nMes = Month(dDesembolso): nAnio = Year(dDesembolso)
        nPrimaPerGracia = rsDatosReprogramados!nPrimaPerGracia
        
        'Parametros de liquidación de deuda
        If IsArray(pvArrayDatos) Then
            For iMat = 0 To UBound(pvArrayDatos) - 1
                Select Case iMat
                    Case 0: nInteresCompAFecha = pvArrayDatos(iMat)
                    Case 1: nInteresGraciaAFecha = pvArrayDatos(iMat)
                    Case 2: nInteresMoratorio = pvArrayDatos(iMat)
                    Case 3: nInteresCompVencAFecha = pvArrayDatos(iMat)
                    Case 4: nSegDesgAnt = pvArrayDatos(iMat)
                    'Case 5: nSegIncAnt = pvArrayDatos(iMat)        RIRO 20210215 COMENTADO
                    'Case 6: nSegIncGraciaAnt = pvArrayDatos(iMat)  RIRO 20210215 COMENTADO
                    Case 5: nPolizaMen = pvArrayDatos(iMat) 'ADD RIRO 20210212
                    Case 6: nPolizaCuotReprog = pvArrayDatos(iMat) 'ADD RIRO 20210212
                    Case 7: nTasaEspCovid = pvArrayDatos(iMat) 'Joep20200910 Tasa Especial covid
                    Case 8: nOpCovid = pvArrayDatos(iMat) 'Joep20200910 Tasa Especial covid
                    Case 9: nPolizaProrrateo = pvArrayDatos(iMat) 'ADD RIRO 20210212
                End Select
            Next
        End If
        'Total Intereses Liquidados
        nTotalIntLiquidados = nInteresCompAFecha + nInteresGraciaAFecha + nInteresMoratorio + nInteresCompVencAFecha
    End If
    Set rsDatosReprogramados = Nothing
    Set oNCOMCredito = Nothing
    
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
    'Call oNCOMGasto.CargaValoreSegDes(psCtaCod, pdFecDesemb, MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, nCoberMinSegDes, nCoberMaxSegDes)'Comento JOEP20200414 covid
    Call oNCOMGasto.CargaValoreSegDes(psCtaCod, pdFecDesemb, MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, nCoberMinSegDes, nCoberMaxSegDes, , bEsReprogramado) 'Add JOEP20200414 covid
    Set oNCOMGasto = Nothing
    
    '3.3.- Cuota Fija Mensual Inicial o temporal (para calendario FechaFija/PeriodoFijo)
    Set oNCOMCredito = New NCOMCredito
    nMontoCuotaTemp = oNCOMCredito.CuotaFijaTipoPeriodo(nTEMTotal, pdFecDesemb, pnNroCuotas, pnMonto, pnTipoPeriodo, pnPeriodo, pnDiaFijo, _
                                                        pnDiasGracia, pnNumMes, nTotalDias, nSumaFAS, bEsReprogramado, bEsPagoAnticipado, dFechaCuotaPend, IIf(bEsReprogramado, dFecVencUltimoPago, dDesembolsoNuevo))
    '3.4.- Factor Valor Actual Saldo
    nFVAS = Round(((1 + nTEDTotal) ^ nTotalDias), 6)
    
    '3.5.- Interes de gracia por periodo (por días de gracia)
    If Not bEsReprogramado Then
    'JOEP20211020 Mejora de pago anticipa
        If bEsPagoAnticipado = True Then
            nIntPeriodoGracia = nIntPeriodoGracia
        Else
            nIntPeriodoGracia = oNCOMCredito.MontoIntPerDias(pnTasaInt, pnDiasGracia, pnMonto)
            nAjusteCuotaGracia = Round(nIntPeriodoGracia / nSumaFAS, 2)
        End If
    'JOEP20211020 Mejora de pago anticipa
    'Comento JOEP20211020 Mejora de pago anticipa
        'nIntPeriodoGracia = oNCOMCredito.MontoIntPerDias(pnTasaInt, pnDiasGracia, pnMonto)
        'nAjusteCuotaGracia = Round(nIntPeriodoGracia / nSumaFAS, 2)
    'Comento JOEP20211020 Mejora de pago anticipa
        Set oNCOMCredito = Nothing
    Else
        nIntPeriodoGracia = nTotalIntLiquidados
        nAjusteCuotaGracia = 0
    End If
    
    '3.6.- Obtiene la Prima del periodo de gracia de manera diaria. (prorrateada según numero de cuotas)
    If (Not bEsPagoAnticipado And Not bEsReprogramado And nPolizaProrrateo <= 0) Then 'RIRO 20210221 Entrará en esta linea en caso de ser desembolso
        nPrimaPerGracia = (pnMontoPoliza * (pnDiasGracia / 30)) / pnNroCuotas
        nPolizaProrrateo = nPrimaPerGracia
        nPolizaMen = pnMontoPoliza
        nPolizaCuotReprog = pnMontoPoliza
    End If
        
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
                If bEsPagoAnticipado = True Then
                    Calendario(i).dFecha = dFecTemp + IIf(i = 0, PlazoPre, pnPeriodo)
                ElseIf bEsReprogramado And (i = 0) Then
                    Calendario(i).dFecha = dFecTemp
                Else
                    Calendario(i).dFecha = dFecTemp + IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo)
                End If
             Else
                 nDia = pnDiaFijo
                 If i = 0 Then
                    If bEsReprogramado Then
                        dFecTemp = dDesembolso
                    Else
                        dFecTemp = dDesembolso + 30
                    End If
                    nDia = Day(dFecTemp): nMes = Month(dFecTemp): nAnio = Year(dFecTemp)
                    
                    'Artificio para cuando el desembolso sea a fines de enero [Febrero]
                    If Not bEsReprogramado Then
                        If CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio))) - dDesembolso < 5 Then
                            nMes = nMes + pnNumMes
                            oNCOMCredito.ValidaFechasFijasCuota nDia, nMes, nAnio
                        End If
                    End If
                 Else
                     If nDia = pnDiaFijo Then nMes = nMes + pnNumMes
                     
                'JOEP20211020 Mejora de pago anticipa
                    If bEsPagoAnticipado = True And pnDiaFijo <> nDiasFijo And i <> 0 Then
                        nDia = nDiasFijo
                    End If
                'JOEP20211020 Mejora de pago anticipa
                     
                     oNCOMCredito.ValidaFechasFijasCuota nDia, nMes, nAnio
                 End If
                 
                 'Si el cliente realizó pagos, solo toma la fecha pendiente de pago
                 If (psCtaCod <> "" And bEsPagoAnticipado = True) Then
                     If i = 0 Then
                        nDia = Day(dFechaCuotaPend)
                        nMes = Month(dFechaCuotaPend)
                        nAnio = Year(dFechaCuotaPend)
                         If (dFechaCuotaPend > dDesembolso) And (dDesembolsoNuevo > dDesembolso) Then
                                dDesembolso = dDesembolsoNuevo
                         End If
                     End If
                 End If
                 dFecTemp = CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy"))
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
                
                'PagoAnt. Reduccion Nro. Cuota o Plazo
                If pbReduccionNroCuotaPagAnt = True Then
                    If (nSaldoCapital >= nMontoCuotaDesembolso) Then
                        nMontoCuotaPrevia = nMontoCuotaDesembolso
                    End If
                End If
                Calendario(i).Cuota = nMontoCuotaPrevia
                
                nDiasCalculo = 0 'RIRO20200904
                
                'Interes Compensatorio Calc.
                If bEsPagoAnticipado = True Then 'Para Pago Anticipado
                    nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, PlazoPre, nSaldoCapital)
                    nDiasCalculo = PlazoPre 'RIRO 20200904
                Else
                    'JOEP20200320 comento Mejora Reprogramacion
                    'nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFecVencUltimoPago, dFechaCuotaPend), DateDiff("d", dDesembolso, Calendario(i).dFecha))), nSaldoCapital)
                    'JOEP20200320 comento Mejora Reprogramacion
                    
                    'nDiasCalculo = IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFechaCuotaPend, dDesembolso), DateDiff("d", dDesembolso, Calendario(i).dFecha))) 'RIRO 20200904
                    
                    'JOEP20200320 Add Mejora Reprogramacion
                    If bEsReprogramado = True Then
                        If nOpCovid = 3 Then
                            nIntCompCalculado = oNCOMCredito.MontoIntPerDias(nTasaEspCovid, DateDiff("d", dFechaCuotaPend, dDesembolso), nSaldoCapital) 'Cambio JOEP_RIRO_20200913
                            nDiasCalculo = DateDiff("d", dFechaCuotaPend, dDesembolso) 'RIRO 20200904
                        Else
                        'nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFechaCuotaPend, dDesembolso), DateDiff("d", dDesembolso, Calendario(i).dFecha))), nSaldoCapital)'COmento JOEP_RIRO_20200913
                            nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dFechaCuotaPend, dDesembolso), nSaldoCapital) 'Cambio JOEP_RIRO_20200913
                            nDiasCalculo = DateDiff("d", dFechaCuotaPend, dDesembolso) 'RIRO 20200904
                        End If
                    Else
                        nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFecVencUltimoPago, dFechaCuotaPend), DateDiff("d", dDesembolso, Calendario(i).dFecha))), nSaldoCapital)
                        nDiasCalculo = IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFechaCuotaPend, dDesembolso), DateDiff("d", dDesembolso, Calendario(i).dFecha))) 'RIRO 20200904
                    End If
                    'JOEP20200320 Add Mejora Reprogramacion
                End If
                
                'RIRO 20200825 mejora en Liquidación *************
                Calendario(i).IntCompCalculado = nIntCompCalculado
                Calendario(i).DiasCalculo = nDiasCalculo
                'END RIRO ****************************************
                 
                'Seguro Desgravamen
                Set oNCOMGasto = New COMNCredito.NCOMGasto
                nMontoSegDes = oNCOMGasto.CalculaSeguroDesgravamen(i, "K", pnTasaSegDes, pnMonto, Calendario(i).IntComp, nSaldoCapital, _
                                            Calendario(i).dFecha, 0, 0, MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, _
                                            nCoberMinSegDes, nCoberMaxSegDes, pMatCalendSegDes, psCtaCod)
                                          
                'ADD RIRO 20210221, Se aplica esta validación si es pago anticipado
                If bEsPagoAnticipado Then
                    'Si realizó pago de seg. des. en el pago anticipado no cobrarlo en el nuevo calend.
                    Set oDCOMCalendario = New COMDCredito.DCOMCalendario
                    If oDCOMCalendario.VerificaCuotaGastoPagadoCalendAnt(psCtaCod, Year(Calendario(i).dFecha), Month(Calendario(i).dFecha)) Then
                        nMontoSegDes = 0
                    End If
                End If
                'Exoneración seguro de desgravamen(Cuota Inicial)
                If pnExoSeguroDesgravamen = 1 Then
                    Calendario(i).SegDes = Format(0, "#0.00")
                Else
                    'Calendario(i).SegDes = Round((nMontoSegDes / 30) * (DateDiff("d", pdFecDesemb, Calendario(i).dFecha)), 2) 'JOEP20200320 Comento Reprogramacion
                    If bEsReprogramado = True Then
                        Calendario(i).SegDes = Round((nMontoSegDes / 30) * (DateDiff("d", dFechaCuotaPend, Calendario(i).dFecha)), 2) 'JOEP20200320 Add Comento Reprogramacion
                    Else
                        Calendario(i).SegDes = Round((nMontoSegDes / 30) * (DateDiff("d", pdFecDesemb, Calendario(i).dFecha)), 2)
                    End If
                    
                    If bEsReprogramado = True Then
                        Calendario(i).SegDes = Calendario(i).SegDes + nSegDesgAnt
                    End If
                End If
                
                'Gasto: Póliza contra incendio
                Calendario(i).CuotaPrimaPoliza = Round(nPolizaCuotReprog, 2) ' RIRO 20210215 se inserta el monto de la póliza del primer mes, concepto 1231
                'RIRO 20210215 COMENTADO ******************************
                'Calendario(i).CuotaPrimaPoliza = pnMontoPoliza
                'If bEsReprogramado = True Then
                '    Calendario(i).CuotaPrimaPoliza = Calendario(i).CuotaPrimaPoliza + nSegIncAnt
                'End If
                'RIRO 20210215 COMENTADO ******************************
                
                'Gasto: Póliza contra incendio por periodo de gracia (prorrateado)
                Calendario(i).CuotaPrimaPolizaGracia = Round(nPolizaProrrateo, 2) 'RIRO 20210215 Se inserta la póliza prorrateada, concepto  1279
                
                'RIRO 20210215 COMENTADO ***************************
                'Calendario(i).CuotaPrimaPolizaGracia = Round(nPrimaPerGracia, 2)
                'If bEsReprogramado = True Then
                '    Calendario(i).CuotaPrimaPolizaGracia = Calendario(i).CuotaPrimaPolizaGracia + nSegIncGraciaAnt
                'End If
                'RIRO 20210215 *************************************
                
                'Interés de Gracia Cuota Inicial
                'JOEP20200321 Comento Mejora Reprogramacion
                'nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(pnCuotaIni, pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", dDesembolso, Calendario(i).dFecha)), _
                '                                                        nIntPeriodoGracia, nMontoCuotaPrevia, Calendario(i).IntComp, Calendario(i).SegDes, (Calendario(i).CuotaPrimaPoliza + Calendario(i).CuotaPrimaPolizaGracia), _
                '                                                       nGraciaGenerada, nGraciaCapitalizada, nGraciaNoPagada)
                'JOEP20200321 Comento Mejora Reprogramacion
                'JOEP20200321 Add Mejora Reprogramacion
                If bEsReprogramado = True Then
                    'nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(pnCuotaIni, pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", dFechaCuotaPend, dDesembolso)),'Comento JOEP_RIRO_20200913
                    'JOEP_RIRO_20200913
                    nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(pnCuotaIni, pnTasaInt, DateDiff("d", dFechaCuotaPend, dDesembolso), _
                                                                         nIntPeriodoGracia, nMontoCuotaPrevia, Calendario(i).IntComp, Calendario(i).SegDes, (Calendario(i).CuotaPrimaPoliza + Calendario(i).CuotaPrimaPolizaGracia), _
                                                                         nGraciaGenerada, nGraciaCapitalizada, nGraciaNoPagada)
                    'JOEP_RIRO_20200913
                Else
                    nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(pnCuotaIni, pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", dDesembolso, Calendario(i).dFecha)), _
                                                                         nIntPeriodoGracia, nMontoCuotaPrevia, Calendario(i).IntComp, Calendario(i).SegDes, (Calendario(i).CuotaPrimaPoliza + Calendario(i).CuotaPrimaPolizaGracia), _
                                                                         nGraciaGenerada, nGraciaCapitalizada, nGraciaNoPagada)
                End If
                'JOEP20200321 Add Mejora Reprogramacion
                
                'RIRO 20200904 Grabando datos ******
                Calendario(i).IntGraciaGenerado = nGraciaGenerada
                Calendario(i).IntGraciaCapitalizado = nGraciaCapitalizada
                Calendario(i).IntGraciaAsignado = nIntGraciaAsignado
                'END RIRO 20200904 *****************
                
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
                Set oDCOMCalendario = Nothing
             Else
                 'PagoAnt. Reduccion Nro. Cuota o Plazo
                 If pbReduccionNroCuotaPagAnt = True Then
                    If (nSaldoCapital >= nMontoCuotaDesembolso) Then
                        nMontoCuotaPrevia = nMontoCuotaDesembolso
                    End If
                 End If
                 
                 'Monto de la Cuota Fija Final
                 Calendario(i).Cuota = nMontoCuotaPrevia
                 nDiasCalculo = 0 'RIRO20200904
                 
                 'Interés Compensatorio
                 If pnTipoPeriodo = PeriodoFijo Then
                    If bEsPagoAnticipado = True Then
                        nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital)
                        nDiasCalculo = pnPeriodo
                    Else
                        nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), nSaldoCapital)
                        nDiasCalculo = IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo)
                    End If
                 Else
                    nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha), nSaldoCapital)
                    nDiasCalculo = DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha)
                 End If
                 
                'RIRO 20200825 mejora en Liquidación *************
                Calendario(i).IntCompCalculado = nIntCompCalculado
                Calendario(i).DiasCalculo = nDiasCalculo
                'END RIRO ****************************************
             
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
                Calendario(i).CuotaPrimaPoliza = Round(nPolizaMen, 2) 'ADD RIRO 20210212, Concepto 1231
                'Calendario(i).CuotaPrimaPoliza = pnMontoPoliza
                
                'Gasto: Póliza contra incendio por periodo de gracia (prorrateado)
                Calendario(i).CuotaPrimaPolizaGracia = Round(nPolizaProrrateo, 2) 'ADD RIRO 20210212, Concepto 1279
                'Calendario(i).CuotaPrimaPolizaGracia = Round(nPrimaPerGracia, 2)
                
                'Interés de Gracia
                nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(i, pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha)), _
                                                                                            nIntPeriodoGracia, nMontoCuotaPrevia, Calendario(i).IntComp, Calendario(i).SegDes, (Calendario(i).CuotaPrimaPoliza + Calendario(i).CuotaPrimaPolizaGracia), _
                                                                                            nGraciaGenerada, nGraciaCapitalizada, nGraciaNoPagada)
                
                'RIRO 20200904 Grabando datos ******
                Calendario(i).IntGraciaGenerado = nGraciaGenerada
                Calendario(i).IntGraciaCapitalizado = nGraciaCapitalizada
                Calendario(i).IntGraciaAsignado = nIntGraciaAsignado
                'END RIRO 20200904 *****************
                
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
                Calendario(i).IntCompDiferenciaCapitalizado = nIntCapitalizadoDiferencia 'RIRO 20200904
                
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
                 If bEsPagoAnticipado = True Then 'Pago anticipado
                     dFecTemp = dFecTemp + IIf(i = 0, PlazoPre, pnPeriodo)
                 Else
                     dFecTemp = dFecTemp + IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo)
                 End If
                 nPeriodoPrimerPF = 0
            End If
         Next i
         
        nVASKU = Round((nSaldoCapital / nFVAS), 2)
        nCVASKU = Round((nVASKU * (1 / Round(nSumaFAS, 6))), 2)
        nSKU = Round(nSaldoCapital, 2)
        
        nValorAjusteIteracion = MetodoAjustePorIteracion(nSKU, nFVAS, nSumaFAS, nContadorSKU)
        nValorAjusteIteracionTotal = nValorAjusteIteracionTotal + nValorAjusteIteracion

        If nNumIteracion > 17 Or pbReduccionNroCuotaPagAnt Then 'Recomendación GELU
            Exit Do
        End If
        nContadorSKU = nContadorSKU + 1
        nNumIteracion = nNumIteracion + 1
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
'<-***** Fin LUCV0180601

'JOEP20200428 Covid cuotas iguales
Private Sub CovidCuotasIgualesCalendarioCuotaFijaNuevo(ByVal pnMonto As Double, _
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
                        Optional ByVal psCtaCod As String, _
                        Optional ByVal psMensFecAprob As Date, _
                        Optional ByVal lnDescripCuota As Variant = "", _
                        Optional ByVal pnMontoGraAnt As Double = 0, _
                        Optional ByVal pnTasaSegDes As Double = 0, _
                        Optional ByRef pMatCalendSegDes As Variant, _
                        Optional ByVal pnExoSeguroDesgravamen As Integer = 0, _
                        Optional ByVal pnMontoPoliza As Double = 0, _
                        Optional ByVal pnTasaSegInc As Double = 0, _
                        Optional ByVal pbReduccionNroCuotaPagAnt As Boolean = False, _
                        Optional ByRef pvArrayDatos As Variant = Nothing)
                        'MADM 20110409 - FECHA APRO - lnDescripCuota
                        'LUCV20180601, Agregó: pnTasaSegDes, pMatCalendSegDes, pnExoSeguroDesgravamen, pnMontoPoliza, pnTasaSegInc, pArrayDatos. Según ERS022-2018
                        'pnMontoGraAnt, para gracia pendiente
                
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
    'Pagos Anticipados
    Dim dFechaCuotaPend As Date
    Dim dDesembolsoNuevo As Date
    Dim nCuotasPendientes As Integer 'add by marg2080524
    Dim nPeriodoPrimerPF As Integer
    Dim oDCOMCalendario As COMDCredito.DCOMCalendario
    Set oDCOMCalendario = New COMDCredito.DCOMCalendario
    'Pag. Ant. Reducción de Nro. Cuotas
    Dim nMontoCuotaDesembolso As Double
    Dim oDCOMCredito As COMDCredito.DCOMCredito
    'Reprogramados
    Dim rsDatosReprogramados As ADODB.Recordset
    Dim bEsReprogramado As Boolean
    Dim dFecVencUltimoPago As Date
    Dim iMat As Integer
    Dim nInteresCompAFecha As Double
    Dim nInteresGraciaAFecha As Double
    Dim nInteresMoratorio As Double
    Dim nInteresCompVencAFecha As Double
    Dim nSegDesgAnt As Double
    Dim nSegIncAnt As Double
    Dim nSegIncGraciaAnt As Double
    Dim nTotalIntLiquidados As Double
    Dim nMontoCuota As Double
    Dim nDiasCalculo As Integer 'RIRO 20200904 Liquidacion
    'Dim dFecVencIni As Date
    '<-***** Fin LUCV20180601
    
    'RIRO 20210212 ********************
    Dim nPolizaMen As Double        ' Póliza mensual sin el prorrateo, concepto 1231
    Dim nPolizaCuotReprog As Double ' Póliza de la primera cuota, concepto 1231
    Dim nPolizaProrrateo As Double      ' Póliza prorrateada en cada cuota, concepto 1279
    'END RIRO
        
    '1.- Inicializa variables
    nPeriodoPrimerPF = 0
    nIntPeriodoGracia = 0: nGraciaGenerada = 0: nGraciaNoPagada = 0
    nVASKU = 0: nSKU = 1.5
    nSumCapital = 0
    nValorAjusteIteracion = 0: nValorAjusteIteracionTotal = 0: nNumIteracion = 0: nContadorSKU = 0
    
    nSKPenultimaCuota = 0
    nPrimaPerGracia = 0
    iMat = 0: bEsReprogramado = False: nInteresCompAFecha = 0: nInteresGraciaAFecha = 0: nInteresMoratorio = 0: nInteresCompVencAFecha = 0
    nSegDesgAnt = 0: nSegIncAnt = 0: nSegIncGraciaAnt = 0
    
    
    '2.- Fecha de desembolso + los días del periodo de gracia
    dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
    nMes = Month(dDesembolso): nAnio = Year(dDesembolso)
    
    '* Para: Pago Anticipado
    If psMensFecAprob = "12:00:00 AM" And psCtaCod <> "" Then
        Dim rsDatosPagoAnticipado As ADODB.Recordset
        Dim DfechaPre As Date
        Dim PlazoPre As Integer
        Dim bEsPagoAnticipado As Boolean
        
        'RIRO 20210221 Comentado porque este método no se utiliza en pagos anticipados, solo en reprogramaciones
        'Set rsDatosPagoAnticipado = New ADODB.Recordset
        'Set oNCOMCredito = New COMNCredito.NCOMCredito
        'Set rsDatosPagoAnticipado = oNCOMCredito.RecuperaValidacionCredMantPrepago(psCtaCod, Format(pdFecDesemb, "dd/mm/yyyy"))
        'Set oNCOMCredito = Nothing
        bEsPagoAnticipado = False
        
        'RIRO 20210221 Comentado porque este método no se utiliza en pagos anticipados, solo en reprogramaciones
        'If Not rsDatosPagoAnticipado.BOF And Not rsDatosPagoAnticipado.EOF Then
        '    bEsPagoAnticipado = True
        '    DfechaPre = rsDatosPagoAnticipado!dVenc
        '    PlazoPre = rsDatosPagoAnticipado!DifPlazo
        '    dFechaCuotaPend = rsDatosPagoAnticipado!dFechaCuotaPend     'Fecha de Prox. Cuota a pagar
        '    dDesembolsoNuevo = rsDatosPagoAnticipado!dDesembolsoNuevo   'Fecha Inicial, Anterior a dFechaCuotaPend
        '    nCuotasPendientes = rsDatosPagoAnticipado!nCuotasPendientes 'Cantidad de Cuotas Pend.
        '    pnCuotaFin = nCuotasPendientes - 1
        '    nPrimaPerGracia = rsDatosPagoAnticipado!nPrimaPerGracia
        '    rsDatosPagoAnticipado.Close
        '
        '    'Cuota Objetivo (PagoAnt: Reduccion Plazo)
        '    Set oDCOMCredito = New COMDCredito.DCOMCredito
        '    If pbReduccionNroCuotaPagAnt Then
        '        nMontoCuotaDesembolso = oDCOMCredito.CuotaFijaMensualPago(psCtaCod)
        '        pnCuotaFin = pnNroCuotas - 1
        '    End If
        '    Set oDCOMCredito = Nothing
        'End If
        'Set rsDatosPagoAnticipado = Nothing
    End If
    
    '** Fechas: Periodo Fijo / Fecha Fija (**Pago Anticipado)
    If pnTipoPeriodo = PeriodoFijo Then
           'Para Pagos Anticipados
             If i = 0 And psMensFecAprob <> "12:00:00 AM" Then
               'Verifica Si tiene Gracia / Diferencia es -
                 If i = 0 And (dDesembolso > psMensFecAprob) Then
                    If lnDescripCuota(0) = "0" Then
                        nPeriodoPrimerPF = pnPeriodo - CInt(DateDiff("d", psMensFecAprob, dDesembolso))
                    Else
                        If lnDescripCuota(1) <> "0" Then
                            nPeriodoPrimerPF = pnPeriodo + Val(lnDescripCuota(1))
                        End If
                    End If
                 End If
             End If
    Else
        nMes = Month(dDesembolso): nAnio = Year(dDesembolso)
    End If
    
    '*** Para: Reprogramación de créditos
    Set oNCOMCredito = New COMNCredito.NCOMCredito
    Set rsDatosReprogramados = New ADODB.Recordset
    Set rsDatosReprogramados = oNCOMCredito.RecuperaDatosReprogramacionCalendario(psCtaCod)
    If Not rsDatosReprogramados.BOF And Not rsDatosReprogramados.EOF And psCtaCod <> "" Then
        bEsReprogramado = True
        dFecVencUltimoPago = rsDatosReprogramados!dFecVencUltimoPago 'Fecha ultima cuota pagada
        dFechaCuotaPend = rsDatosReprogramados!dFechaCuotaPend       'Fecha Cuota Pendiente
        dDesembolso = rsDatosReprogramados!dFecVencReprog            'Fecha Cuota Pendiente + "n" Dias Reprog.
        pnDiaFijo = Day(dDesembolso): nMes = Month(dDesembolso): nAnio = Year(dDesembolso)
        nPrimaPerGracia = rsDatosReprogramados!nPrimaPerGracia
        
        'Parametros de liquidación de deuda
        If IsArray(pvArrayDatos) Then
            For iMat = 0 To UBound(pvArrayDatos) - 1
                Select Case iMat
                    Case 0: nInteresCompAFecha = pvArrayDatos(iMat)
                    Case 1: nInteresGraciaAFecha = pvArrayDatos(iMat)
                    Case 2: nInteresMoratorio = pvArrayDatos(iMat)
                    Case 3: nInteresCompVencAFecha = pvArrayDatos(iMat)
                    Case 4: nSegDesgAnt = pvArrayDatos(iMat)
                    'Case 5: nSegIncAnt = pvArrayDatos(iMat) RIRO 20210212 Comentado
                    'Case 6: nSegIncGraciaAnt = pvArrayDatos(iMat) RIRO 20210212 Comentado
                    Case 5: nPolizaMen = pvArrayDatos(iMat) 'ADD RIRO 20210212
                    Case 6: nPolizaCuotReprog = pvArrayDatos(iMat) 'ADD RIRO 20210212
                    Case 7: nMontoCuota = pvArrayDatos(iMat) 'JOEP20200429 Covid cuotas iguales
                    Case 9: nPolizaProrrateo = pvArrayDatos(iMat) 'ADD RIRO 20210212
                End Select
            Next
        End If
        'Total Intereses Liquidados
        nTotalIntLiquidados = nInteresCompAFecha + nInteresGraciaAFecha + nInteresMoratorio + nInteresCompVencAFecha
    End If
    Set rsDatosReprogramados = Nothing
    Set oNCOMCredito = Nothing
    
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
    'Call oNCOMGasto.CargaValoreSegDes(psCtaCod, pdFecDesemb, MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, nCoberMinSegDes, nCoberMaxSegDes)'Comento JOEP20200414 covid
    Call oNCOMGasto.CargaValoreSegDes(psCtaCod, pdFecDesemb, MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, nCoberMinSegDes, nCoberMaxSegDes, , bEsReprogramado) 'Add JOEP20200414 covid
    Set oNCOMGasto = Nothing
    
    '3.3.- Cuota Fija Mensual Inicial o temporal (para calendario FechaFija/PeriodoFijo)
    Set oNCOMCredito = New NCOMCredito
    nMontoCuotaTemp = nMontoCuota 'oNCOMCredito.CuotaFijaTipoPeriodo(nTEMTotal, pdFecDesemb, pnNroCuotas, pnMonto, pnTipoPeriodo, pnPeriodo, pnDiaFijo, _
                       '                                 pnDiasGracia, pnNumMes, nTotalDias, nSumaFAS, bEsReprogramado, bEsPagoAnticipado, dFechaCuotaPend, IIf(bEsReprogramado, dFecVencUltimoPago, dDesembolsoNuevo))
    '3.4.- Factor Valor Actual Saldo
    nFVAS = Round(((1 + nTEDTotal) ^ nTotalDias), 6)
    
    '3.5.- Interes de gracia por periodo (por días de gracia)
    If Not bEsReprogramado Then
        nIntPeriodoGracia = oNCOMCredito.MontoIntPerDias(pnTasaInt, pnDiasGracia, pnMonto)
        nAjusteCuotaGracia = Round(nIntPeriodoGracia / nSumaFAS, 2)
        Set oNCOMCredito = Nothing
    Else
        nIntPeriodoGracia = nTotalIntLiquidados
        nAjusteCuotaGracia = 0
    End If
    
    '3.6.- Obtiene la Prima del periodo de gracia de manera diaria. (prorrateada según numero de cuotas)
    If Not bEsPagoAnticipado And Not bEsReprogramado Then
        nPrimaPerGracia = (pnMontoPoliza * (pnDiasGracia / 30)) / pnNroCuotas
    End If
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
        
        'lari20210717**********
        Dim nCantCuotas As Integer
        nCantCuotas = 0
        '**********************
        
         '4.1.- Recorrido desde la cuota Inicial hasta la Final
         For i = pnCuotaIni To pnCuotaFin
            '4.1.1.- Administración de fechas
             If pnTipoPeriodo = PeriodoFijo Then
                If i = 0 Then dFecTemp = dDesembolso
                If bEsPagoAnticipado = True Then
                    Calendario(i).dFecha = dFecTemp + IIf(i = 0, PlazoPre, pnPeriodo)
                ElseIf bEsReprogramado And (i = 0) Then
                    Calendario(i).dFecha = dFecTemp
                Else
                    Calendario(i).dFecha = dFecTemp + IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo)
                End If
             Else
                 nDia = pnDiaFijo
                 If i = 0 Then
                    If bEsReprogramado Then
                        dFecTemp = dDesembolso
                    Else
                        dFecTemp = dDesembolso + 30
                    End If
                    nDia = Day(dFecTemp): nMes = Month(dFecTemp): nAnio = Year(dFecTemp)
                    
                    'Artificio para cuando el desembolso sea a fines de enero [Febrero]
                    If Not bEsReprogramado Then
                        If CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio))) - dDesembolso < 5 Then
                            nMes = nMes + pnNumMes
                            oNCOMCredito.ValidaFechasFijasCuota nDia, nMes, nAnio
                        End If
                    End If
                 Else
                     If nDia = pnDiaFijo Then nMes = nMes + pnNumMes
                     oNCOMCredito.ValidaFechasFijasCuota nDia, nMes, nAnio
                 End If
                 
                 'Si el cliente realizó pagos, solo toma la fecha pendiente de pago
                 If (psCtaCod <> "" And bEsPagoAnticipado = True) Then
                     If i = 0 Then
                        nDia = Day(dFechaCuotaPend)
                        nMes = Month(dFechaCuotaPend)
                        nAnio = Year(dFechaCuotaPend)
                         If (dFechaCuotaPend > dDesembolso) And (dDesembolsoNuevo > dDesembolso) Then
                                dDesembolso = dDesembolsoNuevo
                         End If
                     End If
                 End If
                 dFecTemp = CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy"))
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
                
                'PagoAnt. Reduccion Nro. Cuota o Plazo
                If pbReduccionNroCuotaPagAnt = True Then
                    If (nSaldoCapital >= nMontoCuotaDesembolso) Then
                        nMontoCuotaPrevia = nMontoCuotaDesembolso
                    End If
                End If
                Calendario(i).Cuota = nMontoCuotaPrevia
                
                'Interes Compensatorio Calc.
                If bEsPagoAnticipado = True Then 'Para Pago Anticipado
                    nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, PlazoPre, nSaldoCapital)
                    nDiasCalculo = PlazoPre
                Else
                    'JOEP20200320 comento Mejora Reprogramacion
                    'nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFecVencUltimoPago, dFechaCuotaPend), DateDiff("d", dDesembolso, Calendario(i).dFecha))), nSaldoCapital)
                    'JOEP20200320 comento Mejora Reprogramacion
                    
                    'JOEP20200320 Add Mejora Reprogramacion
                    If bEsReprogramado = True Then
                        'nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFechaCuotaPend, dDesembolso), DateDiff("d", dDesembolso, Calendario(i).dFecha))), nSaldoCapital)'Comento JOEP_RIRO_20200913
                        'nDiasCalculo = IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFechaCuotaPend, dDesembolso), DateDiff("d", dDesembolso, Calendario(i).dFecha)))'Comento JOEP_RIRO_20200913
                        nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dFechaCuotaPend, dDesembolso), nSaldoCapital) 'Cambio JOEP_RIRO_20200913
                        nDiasCalculo = DateDiff("d", dFechaCuotaPend, dDesembolso) 'Cambio JOEP_RIRO_20200913
                    Else
                        nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFecVencUltimoPago, dFechaCuotaPend), DateDiff("d", dDesembolso, Calendario(i).dFecha))), nSaldoCapital)
                        nDiasCalculo = IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFecVencUltimoPago, dFechaCuotaPend), DateDiff("d", dDesembolso, Calendario(i).dFecha)))
                    End If
                    'JOEP20200320 Add Mejora Reprogramacion
                End If
                 
                'RIRO 20200825 mejora en Liquidación *************
                nDiasCalculo = IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFechaCuotaPend, dDesembolso), DateDiff("d", dDesembolso, Calendario(i).dFecha)))
                Calendario(i).IntCompCalculado = nIntCompCalculado
                Calendario(i).DiasCalculo = nDiasCalculo
                'END RIRO ****************************************
                 
                'Seguro Desgravamen
                Set oNCOMGasto = New COMNCredito.NCOMGasto
                nMontoSegDes = oNCOMGasto.CalculaSeguroDesgravamen(i, "K", pnTasaSegDes, pnMonto, Calendario(i).IntComp, nSaldoCapital, _
                                            Calendario(i).dFecha, 0, 0, MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, _
                                            nCoberMinSegDes, nCoberMaxSegDes, pMatCalendSegDes, psCtaCod)
                                                          
                'RIRO 20210221 Comentado porque este método no se utiliza en pagos anticipados, solo en reprogramaciones
                'Si realizó pago de seg. des. en el pago anticipado no cobrarlo en el nuevo calend.
                'Set oDCOMCalendario = New COMDCredito.DCOMCalendario
                'If oDCOMCalendario.VerificaCuotaGastoPagadoCalendAnt(psCtaCod, Year(Calendario(i).dFecha), Month(Calendario(i).dFecha)) Then
                '    nMontoSegDes = 0
                'End If
                'Exoneración seguro de desgravamen(Cuota Inicial)
                If pnExoSeguroDesgravamen = 1 Then
                    Calendario(i).SegDes = Format(0, "#0.00")
                Else
                    'Calendario(i).SegDes = Round((nMontoSegDes / 30) * (DateDiff("d", pdFecDesemb, Calendario(i).dFecha)), 2) 'JOEP20200320 Comento Reprogramacion
                    If bEsReprogramado = True Then
                        Calendario(i).SegDes = Round((nMontoSegDes / 30) * (DateDiff("d", dFechaCuotaPend, Calendario(i).dFecha)), 2) 'JOEP20200320 Add Comento Reprogramacion
                    Else
                        Calendario(i).SegDes = Round((nMontoSegDes / 30) * (DateDiff("d", pdFecDesemb, Calendario(i).dFecha)), 2)
                    End If
                    
                    If bEsReprogramado = True Then
                        Calendario(i).SegDes = Calendario(i).SegDes + nSegDesgAnt
                    End If
                End If
                
                'Gasto: Póliza contra incendio
                'RIRO 20210212 Comentado
                'Calendario(i).CuotaPrimaPoliza = pnMontoPoliza
                'If bEsReprogramado = True Then
                '    Calendario(i).CuotaPrimaPoliza = Calendario(i).CuotaPrimaPoliza + nSegIncAnt
                'End If
                Calendario(i).CuotaPrimaPoliza = Round(nPolizaCuotReprog, 2) ' se inserta el monto de la póliza del primer mes, concepto 1231
                
                'Gasto: Póliza contra incendio por periodo de gracia (prorrateado)
                
                'RIRO 20210212 Comentado
                'Calendario(i).CuotaPrimaPolizaGracia = Round(nPrimaPerGracia, 2)
                'If bEsReprogramado = True Then
                '    Calendario(i).CuotaPrimaPolizaGracia = Calendario(i).CuotaPrimaPolizaGracia + nSegIncGraciaAnt
                'End If
                Calendario(i).CuotaPrimaPolizaGracia = Round(nPolizaProrrateo, 2) 'Se inserta la póliza prorrateada, concepto  1279
                                
                'Interés de Gracia Cuota Inicial
                'JOEP20200321 Comento Mejora Reprogramacion
                'nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(pnCuotaIni, pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", dDesembolso, Calendario(i).dFecha)), _
                '                                                        nIntPeriodoGracia, nMontoCuotaPrevia, Calendario(i).IntComp, Calendario(i).SegDes, (Calendario(i).CuotaPrimaPoliza + Calendario(i).CuotaPrimaPolizaGracia), _
                '                                                       nGraciaGenerada, nGraciaCapitalizada, nGraciaNoPagada)
                'JOEP20200321 Comento Mejora Reprogramacion
                'JOEP20200321 Add Mejora Reprogramacion
                If bEsReprogramado = True Then
                    'nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(pnCuotaIni, pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", dFechaCuotaPend, dDesembolso)),'Comento JOEP_RIRO_20200913
                    'Cambio JOEP_RIRO_20200913
                    nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(pnCuotaIni, pnTasaInt, DateDiff("d", dFechaCuotaPend, dDesembolso), _
                                                                         nIntPeriodoGracia, nMontoCuotaPrevia, Calendario(i).IntComp, Calendario(i).SegDes, (Calendario(i).CuotaPrimaPoliza + Calendario(i).CuotaPrimaPolizaGracia), _
                                                                         nGraciaGenerada, nGraciaCapitalizada, nGraciaNoPagada)
                    'Cambio JOEP_RIRO_20200913
                Else
                    nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(pnCuotaIni, pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", dDesembolso, Calendario(i).dFecha)), _
                                                                         nIntPeriodoGracia, nMontoCuotaPrevia, Calendario(i).IntComp, Calendario(i).SegDes, (Calendario(i).CuotaPrimaPoliza + Calendario(i).CuotaPrimaPolizaGracia), _
                                                                         nGraciaGenerada, nGraciaCapitalizada, nGraciaNoPagada)
                End If
                'JOEP20200321 Add Mejora Reprogramacion
                
                'RIRO 20200904 Grabando datos ******
                Calendario(i).IntGraciaGenerado = nGraciaGenerada
                Calendario(i).IntGraciaCapitalizado = nGraciaCapitalizada
                Calendario(i).IntGraciaAsignado = nIntGraciaAsignado
                'END RIRO 20200904 *****************
                
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
                Set oDCOMCalendario = Nothing
             Else
                 'PagoAnt. Reduccion Nro. Cuota o Plazo
                 If pbReduccionNroCuotaPagAnt = True Then
                    If (nSaldoCapital >= nMontoCuotaDesembolso) Then
                        nMontoCuotaPrevia = nMontoCuotaDesembolso
                    End If
                 End If
                 
                 'Monto de la Cuota Fija Final
                 Calendario(i).Cuota = nMontoCuotaPrevia
                 nDiasCalculo = 0 'RIRO20200904
                 
                 'Interés Compensatorio
                 If pnTipoPeriodo = PeriodoFijo Then
                    'RIRO 20210221 Comentado porque este método solo se utiliza en reprogramaciones y no en pagos anticipados
                    'If bEsPagoAnticipado = True Then
                    '    nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital)
                    '    nDiasCalculo = pnPeriodo
                    'Else
                    '    nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), nSaldoCapital)
                    '    nDiasCalculo = IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo)
                    'End If
                    nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), nSaldoCapital)
                    nDiasCalculo = IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo)
                 Else
                    nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha), nSaldoCapital)
                    nDiasCalculo = DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha)
                 End If
             
                'RIRO 20200825 mejora en Liquidación *************
                Calendario(i).IntCompCalculado = nIntCompCalculado
                Calendario(i).DiasCalculo = nDiasCalculo
                'END RIRO ****************************************
             
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
                'Calendario(i).CuotaPrimaPoliza = pnMontoPoliza RIRO 20210212 Comentado
                Calendario(i).CuotaPrimaPoliza = Round(nPolizaMen, 2) 'ADD RIRO 20210212, Concepto 1231
                
                'Gasto: Póliza contra incendio por periodo de gracia (prorrateado)
                'Calendario(i).CuotaPrimaPolizaGracia = Round(nPrimaPerGracia, 2) RIRO 20210212 Comentado
                Calendario(i).CuotaPrimaPolizaGracia = Round(nPolizaProrrateo, 2) 'ADD RIRO 20210212, Concepto 1279
                
                'Interés de Gracia
                nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(i, pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha)), _
                                                                                            nIntPeriodoGracia, nMontoCuotaPrevia, Calendario(i).IntComp, Calendario(i).SegDes, (Calendario(i).CuotaPrimaPoliza + Calendario(i).CuotaPrimaPolizaGracia), _
                                                                                            nGraciaGenerada, nGraciaCapitalizada, nGraciaNoPagada)
                
                'RIRO 20200904 Grabando datos ******
                Calendario(i).IntGraciaGenerado = nGraciaGenerada
                Calendario(i).IntGraciaCapitalizado = nGraciaCapitalizada
                Calendario(i).IntGraciaAsignado = nIntGraciaAsignado
                'END RIRO 20200904 *****************
                
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
            
            'LARI20210717*********************************************************
            If ((nSaldoCapital - Calendario(i).Captital) > 0) Then
                Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                Else
                Calendario(i).Captital = nSaldoCapital
                Calendario(i).SaldoCap = nSaldoCapital
                
                If (nCantCuotas = 0) Then
                    nCantCuotas = i
                End If
            End If
            '*********************************************************************
            
            nSaldoCapital = nSaldoCapital - Calendario(i).Captital
            
            '4.1.5.- Saldo Capital de la Penultima Cuota
            If i = pnCuotaFin - 1 Then
                nSKPenultimaCuota = Calendario(i).SaldoCap
            End If
            
            If pnTipoPeriodo = PeriodoFijo Then
                 If bEsPagoAnticipado = True Then 'Pago anticipado
                     dFecTemp = dFecTemp + IIf(i = 0, PlazoPre, pnPeriodo)
                 Else
                     dFecTemp = dFecTemp + IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo)
                 End If
                 nPeriodoPrimerPF = 0
            End If
         Next i
         
        If nSaldoCapital > 0 Then
            pnCuotaFin = pnCuotaFin + 1 'JOEP20200429 Covid Cuotas iguales
        End If
         
        nVASKU = Round((nSaldoCapital / nFVAS), 2)
        If nSumaFAS = 0 Then
            nCVASKU = 0
        Else
            nCVASKU = Round((nVASKU * (1 / Round(nSumaFAS, 6))), 2)
        End If
        
        If nSaldoCapital < 0 Then
            nSKU = Round(-1, 2)
        Else
            nSKU = Round(nSaldoCapital, 2)
        End If
        
        If nSumaFAS = 0 Then
            nValorAjusteIteracion = 0
        Else
            nValorAjusteIteracion = MetodoAjustePorIteracion(nSKU, nFVAS, nSumaFAS, nContadorSKU)
        End If
        
        nValorAjusteIteracionTotal = nValorAjusteIteracionTotal + nValorAjusteIteracion
        
        'Mejora reprogramacion mantener cuota JOEP20201020
            'If nSaldoCapital < 0 Then
            If nSaldoCapital <= 0 Then
                Exit Do
            End If
        'Mejora reprogramacion mantener cuota JOEP20201020
            'If nNumIteracion > 17 Or pbReduccionNroCuotaPagAnt Then 'Recomendación GELU
             '   Exit Do
            'End If
        
        
        nContadorSKU = nContadorSKU + 1
        nNumIteracion = nNumIteracion + 1
        nSaldoCapital = 0
    Loop
    'pvArrayDatos(8) = pnCuotaFin + 1 'Add JOEP-LUCV
    pvArrayDatos(8) = nCantCuotas + 1 'Add LARI
    
    '5.- Ajuste Saldo Capital en la ultima Cuota
    If nSumCapital <> pnMonto Then
        'Calendario(pnCuotaFin).Captital = Calendario(pnCuotaFin).Captital + Round(Calendario(pnCuotaFin).SaldoCap, 2)
        
        '2021-07-31 LARI - JOEP **********El duo terrorifico!!!*************************************************************
        If Round(Calendario(pnCuotaFin).Captital, 2) <> Round(Calendario(pnCuotaFin).SaldoCap, 2) Then
            Calendario(pnCuotaFin).Captital = Calendario(pnCuotaFin).Captital + Round(Calendario(pnCuotaFin).SaldoCap, 2)
        End If
        '*******************************************************************************************************************
        If pnCuotaFin = 0 Then 'Cuando Calend. es a una cuota
            Calendario(pnCuotaFin).SaldoCap = Calendario(pnCuotaFin).Captital - Calendario(pnCuotaFin).Captital
        Else
            Calendario(pnCuotaFin).SaldoCap = Round(nSKPenultimaCuota, 2) - Round(Calendario(pnCuotaFin).Captital, 2)
        End If
    End If
    Set oNCOMCredito = Nothing
End Sub

Public Function CovidCuotasIgualesGeneraCalendario(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, ByVal pnPeriodo As Double, _
        ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
        ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
        Optional ByVal MatGracia As Variant = "", Optional ByVal pbCuotaFijaFechaFija As Boolean = False, _
        Optional ByVal pbCuotaComodin As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
        Optional ByVal pMatDesPar As Variant = "", _
        Optional ByVal pnNumMes As Integer = 1, _
        Optional ByVal pbMiViv As Boolean = False, _
        Optional ByVal bQuincena As Boolean, _
        Optional ByVal pbGraciaEnCuotas As Boolean = False, _
        Optional ByVal pnTasaGracia As Double = 0, _
        Optional ByVal pnDiaFijo2 As Integer = 0, _
        Optional ByVal pnMontoCapInicial As Double = 0, _
        Optional ByVal pbPagoInteresEnGracia As Boolean = False, _
        Optional ByVal psCtaCod As String, _
        Optional ByVal psMensFecAprob As Date, _
        Optional ByVal lnDescripCuota As Variant = "", _
        Optional ByVal pnMontoGra As Double = 0, _
        Optional ByVal psIncluyeIntCap As String = "S", _
        Optional ByVal pnCuotaBalon As Integer = 0, _
        Optional ByVal pbReprogConvenio As Boolean = False, _
        Optional ByVal pbReduccionCuotaPagAnt As Boolean = False, _
        Optional ByVal pnTasaSegDes As Double = 0, Optional ByRef pMatCalendSegDes As Variant, _
        Optional ByVal pnExoSeguroDesgravamen As Integer = 0, Optional ByVal pnMontoPoliza As Double = 0, _
        Optional ByVal pnTasaSegInc As Double = 0, Optional ByRef pArrayDatos As Variant = Nothing) As Variant
        'MADM 20110518 pbTieneGracia - lnDescripCuota
        'Si pnMontoCapInicial >0 entonces se Capitalizo la Gracia sin aumentar el capital
        'MAVM 20130209
        'WIOR 20131111 AGREGO pnCuotaBalon
        'WIOR 20140514 AGREGO pbReprogConvenio
        'LUCV20180320, Agregó: pbReduccionCuotaPagAnt
        'LUCV20180601, Agregó: pnTasaSegDes, pMatCalendSegDes,pnExoSeguroDesgravamen, pnMontoPoliza Según ERS022-2018
                                
Dim sCalendPagos() As String
Dim i As Integer
Dim TasaComTemp As Double
Dim nPlazoCom As Integer
Dim nCuotasCom As Integer
Dim nCuotaMontoCom As Double
Dim oCred As COMNCredito.NCOMCredito

'Gracia en Cuotas
'Dim nCuotasGracia As Integer
Dim nCuotaIni As Integer
Dim nCuotaFin As Integer
Dim dDesembolso As Date
        
    On Error GoTo ErrorGeneraCalendario
    nCuotaIni = 0
    nCuotaFin = pnNroCuotas - 1
    'ReDim Calendario(pnNroCuotas) 'Comento JOEP-LUCV
    ReDim Calendario(1000) 'Add JOEP-LUCV
    Dim nCuotasNew As Integer 'Add JOEP-LUCV
    
    Call CovidCuotasIgualesProcesarCalendario(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, _
                            pnTipoCuota, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, _
                            bProxMes, MatGracia, pbCuotaFijaFechaFija, pbDesemParcial, pMatDesPar, _
                            pnNumMes, pbMiViv, bQuincena, nCuotaIni, nCuotaFin, pnDiaFijo2, pnMontoCapInicial, _
                            psCtaCod, psMensFecAprob, lnDescripCuota, pnMontoGra, psIncluyeIntCap, _
                            pnCuotaBalon, pbReprogConvenio, pbReduccionCuotaPagAnt, _
                            pnTasaSegDes, pMatCalendSegDes, pnExoSeguroDesgravamen, pnMontoPoliza, pnTasaSegInc, pArrayDatos)
                            'WIOR 20131111 AGREGO pnCuotaBalon
                            'WIOR 20140514 AGREGO pbReprogConvenio
                            'LUCV20180320, agregó bReduccionCuotaPagAnt Según ERS087-2017
                            'LUCV20180601, Agregó: pnTasaSegDes, pnExoneraSegDes, pMatCalendSegDes, pnMontoPoliza
    nCuotasNew = pArrayDatos(8) 'Add JOEP-LUCV
    'ReDim sCalendPagos(UBound(Calendario), 16) ' LUCV20180601 Según ERS022-2018, Modificó 14 por 16
    ReDim sCalendPagos(nCuotasNew, 22) 'Add JOEP-LUCV / RIRO 20200914 Se ingrementó campo al número 22
       
       'For i = 0 To UBound(Calendario) - 1'Comento JOEP-LUCV
       For i = 0 To nCuotasNew - 1
            sCalendPagos(i, 0) = Format(Calendario(i).dFecha, "dd/mm/yyyy")
            sCalendPagos(i, 1) = Trim(Str(Calendario(i).NroCuota))
            sCalendPagos(i, 2) = Format(Calendario(i).Cuota, "#0.00")
            sCalendPagos(i, 3) = Format(Calendario(i).Captital, "#0.00")
            sCalendPagos(i, 4) = Format(Calendario(i).IntComp, "#0.00")
            sCalendPagos(i, 5) = Format(Calendario(i).IntGra, "#0.00") 'LUCV20180601, Según ERS022-2018
            sCalendPagos(i, 6) = Format(Calendario(i).Gasto, "#0.00") 'LUCV20180601, Según ERS022-2018
            sCalendPagos(i, 7) = Format(Calendario(i).SaldoCap, "#0.00")
            sCalendPagos(i, 8) = Format(Calendario(i).SegDes, "#0.00")
            sCalendPagos(i, 9) = Format(Calendario(i).SegInm, "#0.00")
            sCalendPagos(i, 10) = Format(Calendario(i).CuotaGra, "#0.00")
            sCalendPagos(i, 11) = Format(Calendario(i).IntCompGra, "#0.00")
            sCalendPagos(i, 12) = Format(Calendario(i).SaldoCapGra, "#0.00")
            'sCalendPagos(i, 13) = Format(Calendario(i).SegMultMype, "#0.00")
            sCalendPagos(i, 14) = Format(0, "#0.00")
            sCalendPagos(i, 15) = Format(Calendario(i).CuotaPrimaPoliza, "#0.00") 'LUCV20180601, Según ERS022-2018
            sCalendPagos(i, 16) = Format(Calendario(i).CuotaPrimaPolizaGracia, "#0.00") 'LUCV20180601, Según ERS022-2018
            
            'RIRO20200825 Mejora en Liquidación ******
            sCalendPagos(i, 17) = Format(Calendario(i).IntCompCalculado, "#0.00")
            sCalendPagos(i, 18) = Format(Calendario(i).DiasCalculo, "#0.00")
            sCalendPagos(i, 19) = Format(Calendario(i).IntCompDiferenciaCapitalizado, "#0.00")
            sCalendPagos(i, 20) = Format(Calendario(i).IntGraciaGenerado, "#0.00")
            sCalendPagos(i, 21) = Format(Calendario(i).IntGraciaCapitalizado, "#0.00")
            sCalendPagos(i, 22) = Format(Calendario(i).IntGraciaAsignado, "#0.00")
            'RIRO20200825 Mejora en Liquidación ******
            
        Next i
        CovidCuotasIgualesGeneraCalendario = sCalendPagos
        Exit Function
ErrorGeneraCalendario:
        Err.Raise Err.Number, "Error", Err.Description
End Function

Private Sub CovidCuotasIgualesProcesarCalendario(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, _
                Optional ByVal pbDesemParcial As Boolean = False, Optional ByVal pMatDesPar As Variant = "", _
                Optional ByVal pnNumMes As Integer = 1, _
                Optional ByVal pbMiViv As Boolean = False, _
                Optional ByVal bQuincena As Boolean, _
                Optional ByVal pnCuotaIni As Integer = 0, _
                Optional ByVal pnCuotaFin As Integer = 0, _
                Optional ByVal pnDiaFijo2 As Integer = 0, _
                Optional pnMontoCapInicial As Double = 0, _
                Optional psCtaCod As String, _
                Optional ByVal psMensFecAprob As Date, _
                Optional ByVal lnDescripCuota As Variant = "", _
                Optional ByVal pnMontoGra As Double = 0, _
                Optional ByVal psIncluyeIntCap As String = "S", _
                Optional ByVal pnCuotaBalon As Integer = 0, _
                Optional ByVal pbReprogConvenio As Boolean = False, _
                Optional ByVal pbReduccionCuotaPagAnt As Boolean = False, _
                Optional ByVal pnTasaSegDes As Double = 0, Optional ByRef pMatCalendSegDes As Variant, Optional ByVal pnExoSeguroDesgravamen As Integer = 0, _
                Optional ByVal pnMontoPoliza As Double = 0, Optional ByVal pnTasaSegInc As Double = 0, Optional ByRef pArrayDatos As Variant = Nothing)
                'MAVM 11092010 Se agrego el par: psCtaCod (Prepago)
                'madm 20110531 - lnDescripCuota
                'MAVM 20130209
                'WIOR 20131111 AGREGO pnCuotaBalon
                'WIOR 20140514 AGREGO pbReprogConvenio
                'LUCV20180320, Agregó: pbReduccionCuotaPagAnt
                'LUCV20180601, Agregó : pnTasaSegDes, pnExoneraSegDes, pMatCalendSegDes, pnExoSeguroDesgravamen, pnMontoPoliza,pnTasaSegInc, pArrayDatos
                
                If bQuincena = False Then
                    Select Case pnTipoCuota
                        Case Creciente
                            Call CalendarioCreciente(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pnCuotaIni, pnCuotaFin, pnDiaFijo2, pnMontoCapInicial)
                        Case Decreciente
                            Call CalendarioDecreciente(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pnCuotaIni, pnCuotaFin, pnDiaFijo2, pnMontoCapInicial)
                        Case Fija
                                Call CovidCuotasIgualesCalendarioCuotaFijaNuevo(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, _
                                                                pdFecDesemb, pnTipoPeriodo, pnTipoGracia, _
                                                                pnDiasGracia, pnDiaFijo, pnNumMes, _
                                                                pnCuotaIni, pnCuotaFin, psCtaCod, _
                                                                psMensFecAprob, lnDescripCuota, pnMontoGra, _
                                                                pnTasaSegDes, pMatCalendSegDes, _
                                                                pnExoSeguroDesgravamen, pnMontoPoliza, pnTasaSegInc, _
                                                                pbReduccionCuotaPagAnt, pArrayDatos)
                                'LUCV20180601. Agregó pnTasaSegDes, pMatCalendSegDes, pnExoneraSegSDes, pnMontoPoliza, pnMontoPoliza, pnTasaSegInc . Según ERS022-2018
'<-***** Fin LUCV20180601
                    End Select
                Else
                    If bQuincena = True Then
                        Call CalendarioTrabajadoresDirectores(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, bProxMes, pnDiasGracia, MatGracia, pnTipoGracia, pnCuotaIni, pnCuotaFin)
                    End If
                End If
End Sub

'JOEP20201012 - Facilidades Crediticias.
Public Function SimuladorReprogGeneraCalendario(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, ByVal pnPeriodo As Double, _
        ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
        ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
        Optional ByVal MatGracia As Variant = "", Optional ByVal pbCuotaFijaFechaFija As Boolean = False, _
        Optional ByVal pbCuotaComodin As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
        Optional ByVal pMatDesPar As Variant = "", _
        Optional ByVal pnNumMes As Integer = 1, _
        Optional ByVal pbMiViv As Boolean = False, _
        Optional ByVal bQuincena As Boolean, _
        Optional ByVal pbGraciaEnCuotas As Boolean = False, _
        Optional ByVal pnTasaGracia As Double = 0, _
        Optional ByVal pnDiaFijo2 As Integer = 0, _
        Optional ByVal pnMontoCapInicial As Double = 0, _
        Optional ByVal pbPagoInteresEnGracia As Boolean = False, _
        Optional ByVal psCtaCod As String, _
        Optional ByVal psMensFecAprob As Date, _
        Optional ByVal lnDescripCuota As Variant = "", _
        Optional ByVal pnMontoGra As Double = 0, _
        Optional ByVal psIncluyeIntCap As String = "S", _
        Optional ByVal pnCuotaBalon As Integer = 0, _
        Optional ByVal pbReprogConvenio As Boolean = False, _
        Optional ByVal pbReduccionCuotaPagAnt As Boolean = False, _
        Optional ByVal pnTasaSegDes As Double = 0, Optional ByRef pMatCalendSegDes As Variant, _
        Optional ByVal pnExoSeguroDesgravamen As Integer = 0, Optional ByVal pnMontoPoliza As Double = 0, _
        Optional ByVal pnTasaSegInc As Double = 0, Optional ByRef pArrayDatos As Variant = Nothing) As Variant
        'MADM 20110518 pbTieneGracia - lnDescripCuota
        'Si pnMontoCapInicial >0 entonces se Capitalizo la Gracia sin aumentar el capital
        'MAVM 20130209
        'WIOR 20131111 AGREGO pnCuotaBalon
        'WIOR 20140514 AGREGO pbReprogConvenio
        'LUCV20180320, Agregó: pbReduccionCuotaPagAnt
        'LUCV20180601, Agregó: pnTasaSegDes, pMatCalendSegDes,pnExoSeguroDesgravamen, pnMontoPoliza Según ERS022-2018
                                
Dim sCalendPagos() As String
Dim i As Integer
Dim TasaComTemp As Double
Dim nPlazoCom As Integer
Dim nCuotasCom As Integer
Dim nCuotaMontoCom As Double
Dim oCred As COMNCredito.NCOMCredito

'Gracia en Cuotas
'Dim nCuotasGracia As Integer
Dim nCuotaIni As Integer
Dim nCuotaFin As Integer
Dim dDesembolso As Date
        
    On Error GoTo ErrorGeneraCalendario
    
    If (pArrayDatos(8) = 2 Or pArrayDatos(8) = 4) Then
        ReDim Calendario(1000)
    Else
        nCuotaIni = 0
        nCuotaFin = pnNroCuotas - 1
        ReDim Calendario(pnNroCuotas)
    End If
    
    Call SimuladorReprogProcesarCalendario(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, _
                            pnTipoCuota, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, _
                            bProxMes, MatGracia, pbCuotaFijaFechaFija, pbDesemParcial, pMatDesPar, _
                            pnNumMes, pbMiViv, bQuincena, nCuotaIni, nCuotaFin, pnDiaFijo2, pnMontoCapInicial, _
                            psCtaCod, psMensFecAprob, lnDescripCuota, pnMontoGra, psIncluyeIntCap, _
                            pnCuotaBalon, pbReprogConvenio, pbReduccionCuotaPagAnt, _
                            pnTasaSegDes, pMatCalendSegDes, pnExoSeguroDesgravamen, pnMontoPoliza, pnTasaSegInc, pArrayDatos)
                            'WIOR 20131111 AGREGO pnCuotaBalon
                            'WIOR 20140514 AGREGO pbReprogConvenio
                            'LUCV20180320, agregó bReduccionCuotaPagAnt Según ERS087-2017
                            'LUCV20180601, Agregó: pnTasaSegDes, pnExoneraSegDes, pMatCalendSegDes, pnMontoPoliza
    
    If (pArrayDatos(8) = 2 Or pArrayDatos(8) = 4) Then
        Dim nCuotasNew As Integer
        nCuotasNew = pArrayDatos(14)
        ReDim sCalendPagos(nCuotasNew, 22)
        
        For i = 0 To nCuotasNew - 1
            sCalendPagos(i, 0) = Format(Calendario(i).dFecha, "dd/mm/yyyy")
            sCalendPagos(i, 1) = Trim(Str(Calendario(i).NroCuota))
            sCalendPagos(i, 2) = Format(Calendario(i).Cuota, "#0.00")
            sCalendPagos(i, 3) = Format(Calendario(i).Captital, "#0.00")
            sCalendPagos(i, 4) = Format(Calendario(i).IntComp, "#0.00")
            sCalendPagos(i, 5) = Format(Calendario(i).IntGra, "#0.00") 'LUCV20180601, Según ERS022-2018
            sCalendPagos(i, 6) = Format(Calendario(i).Gasto, "#0.00") 'LUCV20180601, Según ERS022-2018
            sCalendPagos(i, 7) = Format(Calendario(i).SaldoCap, "#0.00")
            sCalendPagos(i, 8) = Format(Calendario(i).SegDes, "#0.00")
            sCalendPagos(i, 9) = Format(Calendario(i).SegInm, "#0.00")
            sCalendPagos(i, 10) = Format(Calendario(i).CuotaGra, "#0.00")
            sCalendPagos(i, 11) = Format(Calendario(i).IntCompGra, "#0.00")
            sCalendPagos(i, 12) = Format(Calendario(i).SaldoCapGra, "#0.00")
            'sCalendPagos(i, 13) = Format(Calendario(i).SegMultMype, "#0.00")
            sCalendPagos(i, 14) = Format(0, "#0.00")
            sCalendPagos(i, 15) = Format(Calendario(i).CuotaPrimaPoliza, "#0.00") 'LUCV20180601, Según ERS022-2018
            sCalendPagos(i, 16) = Format(Calendario(i).CuotaPrimaPolizaGracia, "#0.00") 'LUCV20180601, Según ERS022-2018
            
            'RIRO20200825 Mejora en Liquidación ******
            sCalendPagos(i, 17) = Format(Calendario(i).IntCompCalculado, "#0.00")
            sCalendPagos(i, 18) = Format(Calendario(i).DiasCalculo, "#0.00")
            sCalendPagos(i, 19) = Format(Calendario(i).IntCompDiferenciaCapitalizado, "#0.00")
            sCalendPagos(i, 20) = Format(Calendario(i).IntGraciaGenerado, "#0.00")
            sCalendPagos(i, 21) = Format(Calendario(i).IntGraciaCapitalizado, "#0.00")
            sCalendPagos(i, 22) = Format(Calendario(i).IntGraciaAsignado, "#0.00")
            'RIRO20200825 Mejora en Liquidación ******
        Next i
    Else
        ReDim sCalendPagos(UBound(Calendario), 22) ' LUCV20180601 Según ERS022-2018, Modificó 14 por 16
       'RIRO20200825 Se cambió indice de 16 a 22
       For i = 0 To UBound(Calendario) - 1
            sCalendPagos(i, 0) = Format(Calendario(i).dFecha, "dd/mm/yyyy")
            sCalendPagos(i, 1) = Trim(Str(Calendario(i).NroCuota))
            sCalendPagos(i, 2) = Format(Calendario(i).Cuota, "#0.00")
            sCalendPagos(i, 3) = Format(Calendario(i).Captital, "#0.00")
            sCalendPagos(i, 4) = Format(Calendario(i).IntComp, "#0.00")
            sCalendPagos(i, 5) = Format(Calendario(i).IntGra, "#0.00") 'LUCV20180601, Según ERS022-2018
            sCalendPagos(i, 6) = Format(Calendario(i).Gasto, "#0.00") 'LUCV20180601, Según ERS022-2018
            sCalendPagos(i, 7) = Format(Calendario(i).SaldoCap, "#0.00")
            sCalendPagos(i, 8) = Format(Calendario(i).SegDes, "#0.00")
            sCalendPagos(i, 9) = Format(Calendario(i).SegInm, "#0.00")
            sCalendPagos(i, 10) = Format(Calendario(i).CuotaGra, "#0.00")
            sCalendPagos(i, 11) = Format(Calendario(i).IntCompGra, "#0.00")
            sCalendPagos(i, 12) = Format(Calendario(i).SaldoCapGra, "#0.00")
            'sCalendPagos(i, 13) = Format(Calendario(i).SegMultMype, "#0.00")
            sCalendPagos(i, 14) = Format(0, "#0.00")
            sCalendPagos(i, 15) = Format(Calendario(i).CuotaPrimaPoliza, "#0.00") 'LUCV20180601, Según ERS022-2018
            sCalendPagos(i, 16) = Format(Calendario(i).CuotaPrimaPolizaGracia, "#0.00") 'LUCV20180601, Según ERS022-2018
            
            'RIRO20200825 Mejora en Liquidación ******
            sCalendPagos(i, 17) = Format(Calendario(i).IntCompCalculado, "#0.00")
            sCalendPagos(i, 18) = Format(Calendario(i).DiasCalculo, "#0.00")
            sCalendPagos(i, 19) = Format(Calendario(i).IntCompDiferenciaCapitalizado, "#0.00")
            sCalendPagos(i, 20) = Format(Calendario(i).IntGraciaGenerado, "#0.00")
            sCalendPagos(i, 21) = Format(Calendario(i).IntGraciaCapitalizado, "#0.00")
            sCalendPagos(i, 22) = Format(Calendario(i).IntGraciaAsignado, "#0.00")
            'RIRO20200825 Mejora en Liquidación ******
        Next i
    End If
       
       
        SimuladorReprogGeneraCalendario = sCalendPagos
        Exit Function
ErrorGeneraCalendario:
        Err.Raise Err.Number, "Error", Err.Description
End Function

Private Sub SimuladorReprogProcesarCalendario(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, _
                Optional ByVal pbDesemParcial As Boolean = False, Optional ByVal pMatDesPar As Variant = "", _
                Optional ByVal pnNumMes As Integer = 1, _
                Optional ByVal pbMiViv As Boolean = False, _
                Optional ByVal bQuincena As Boolean, _
                Optional ByVal pnCuotaIni As Integer = 0, _
                Optional ByVal pnCuotaFin As Integer = 0, _
                Optional ByVal pnDiaFijo2 As Integer = 0, _
                Optional pnMontoCapInicial As Double = 0, _
                Optional psCtaCod As String, _
                Optional ByVal psMensFecAprob As Date, _
                Optional ByVal lnDescripCuota As Variant = "", _
                Optional ByVal pnMontoGra As Double = 0, _
                Optional ByVal psIncluyeIntCap As String = "S", _
                Optional ByVal pnCuotaBalon As Integer = 0, _
                Optional ByVal pbReprogConvenio As Boolean = False, _
                Optional ByVal pbReduccionCuotaPagAnt As Boolean = False, _
                Optional ByVal pnTasaSegDes As Double = 0, Optional ByRef pMatCalendSegDes As Variant, Optional ByVal pnExoSeguroDesgravamen As Integer = 0, _
                Optional ByVal pnMontoPoliza As Double = 0, Optional ByVal pnTasaSegInc As Double = 0, Optional ByRef pArrayDatos As Variant = Nothing)
                'MAVM 11092010 Se agrego el par: psCtaCod (Prepago)
                'madm 20110531 - lnDescripCuota
                'MAVM 20130209
                'WIOR 20131111 AGREGO pnCuotaBalon
                'WIOR 20140514 AGREGO pbReprogConvenio
                'LUCV20180320, Agregó: pbReduccionCuotaPagAnt
                'LUCV20180601, Agregó : pnTasaSegDes, pnExoneraSegDes, pMatCalendSegDes, pnExoSeguroDesgravamen, pnMontoPoliza,pnTasaSegInc, pArrayDatos
                
                If bQuincena = False Then
                    Select Case pnTipoCuota
                        Case Creciente
                            Call CalendarioCreciente(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pnCuotaIni, pnCuotaFin, pnDiaFijo2, pnMontoCapInicial)
                        Case Decreciente
                            Call CalendarioDecreciente(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pnCuotaIni, pnCuotaFin, pnDiaFijo2, pnMontoCapInicial)
                        Case Fija
                                Call SimuladorReprogCalendarioCuotaFijaNuevo(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, _
                                                                pdFecDesemb, pnTipoPeriodo, pnTipoGracia, _
                                                                pnDiasGracia, pnDiaFijo, pnNumMes, _
                                                                pnCuotaIni, pnCuotaFin, psCtaCod, _
                                                                psMensFecAprob, lnDescripCuota, pnMontoGra, _
                                                                pnTasaSegDes, pMatCalendSegDes, _
                                                                pnExoSeguroDesgravamen, pnMontoPoliza, pnTasaSegInc, _
                                                                pbReduccionCuotaPagAnt, pArrayDatos)
                                'LUCV20180601. Agregó pnTasaSegDes, pMatCalendSegDes, pnExoneraSegSDes, pnMontoPoliza, pnMontoPoliza, pnTasaSegInc . Según ERS022-2018
'                            Else
'                                Call CalendarioCuotaFijaPagoAnticipado(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, _
'                                            pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pbCuotaFijaFechaFija, pbDesemParcial, _
'                                            pMatDesPar, pnNumMes, pbMiViv, pnCuotaIni, pnCuotaFin, pnDiaFijo2, pnMontoCapInicial, psCtaCod, _
'                                            psMensFecAprob, lnDescripCuota, pnMontoGra, psIncluyeIntCap, pnCuotaBalon, pbReprogConvenio)
'                            End If
                             '<---***** Fin LUCV20180320
'<-***** Fin LUCV20180601
                    End Select
                Else
                    If bQuincena = True Then
                        Call CalendarioTrabajadoresDirectores(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, bProxMes, pnDiasGracia, MatGracia, pnTipoGracia, pnCuotaIni, pnCuotaFin)
                    End If
                End If
End Sub

Private Sub SimuladorReprogCalendarioCuotaFijaNuevo(ByVal pnMonto As Double, _
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
                        Optional ByVal psCtaCod As String, _
                        Optional ByVal psMensFecAprob As Date, _
                        Optional ByVal lnDescripCuota As Variant = "", _
                        Optional ByVal pnMontoGraAnt As Double = 0, _
                        Optional ByVal pnTasaSegDes As Double = 0, _
                        Optional ByRef pMatCalendSegDes As Variant, _
                        Optional ByVal pnExoSeguroDesgravamen As Integer = 0, _
                        Optional ByVal pnMontoPoliza As Double = 0, _
                        Optional ByVal pnTasaSegInc As Double = 0, _
                        Optional ByVal pbReduccionNroCuotaPagAnt As Boolean = False, _
                        Optional ByRef pvArrayDatos As Variant = Nothing)
                        'MADM 20110409 - FECHA APRO - lnDescripCuota
                        'LUCV20180601, Agregó: pnTasaSegDes, pMatCalendSegDes, pnExoSeguroDesgravamen, pnMontoPoliza, pnTasaSegInc, pArrayDatos. Según ERS022-2018
                        'pnMontoGraAnt, para gracia pendiente
                
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
    'Pagos Anticipados
    Dim dFechaCuotaPend As Date
    Dim dDesembolsoNuevo As Date
    Dim nCuotasPendientes As Integer 'add by marg2080524
    Dim nPeriodoPrimerPF As Integer
    Dim oDCOMCalendario As COMDCredito.DCOMCalendario
    Set oDCOMCalendario = New COMDCredito.DCOMCalendario
    'Pag. Ant. Reducción de Nro. Cuotas
    Dim nMontoCuotaDesembolso As Double
    Dim oDCOMCredito As COMDCredito.DCOMCredito
    'Reprogramados
    Dim rsDatosReprogramados As ADODB.Recordset
    Dim bEsReprogramado As Boolean
    Dim dFecVencUltimoPago As Date
    Dim iMat As Integer
    Dim nInteresCompAFecha As Double
    Dim nInteresGraciaAFecha As Double
    Dim nInteresMoratorio As Double
    Dim nInteresCompVencAFecha As Double
    Dim nSegDesgAnt As Double
    Dim nSegIncAnt As Double
    Dim nSegIncGraciaAnt As Double
    Dim nTotalIntLiquidados As Double
    Dim nDiasCalculo As Integer 'RIRO 20200904 Liquidacion
    Dim nTasaEspCovid As Double 'Joep20200910 Tasa Especial covid
    Dim nOpCovid As Integer 'Joep20200910 Tasa Especial covid
    Dim nMontoCuota As Double
    'Dim dFecVencIni As Date
    '<-***** Fin LUCV20180601
    
    'RIRO 20210212 ********************
    Dim nPolizaMen As Double        ' Póliza mensual sin el prorrateo, concepto 1231
    Dim nPolizaCuotReprog As Double ' Póliza de la primera cuota, concepto 1231
    Dim nPolizaProrrateo As Double      ' Póliza prorrateada en cada cuota, concepto 1279
    'END RIRO
    
    '1.- Inicializa variables
    nPeriodoPrimerPF = 0
    nIntPeriodoGracia = 0: nGraciaGenerada = 0: nGraciaNoPagada = 0
    nVASKU = 0: nSKU = 1.5
    nSumCapital = 0
    nValorAjusteIteracion = 0: nValorAjusteIteracionTotal = 0: nNumIteracion = 0: nContadorSKU = 0
    
    nSKPenultimaCuota = 0
    nPrimaPerGracia = 0
    iMat = 0: bEsReprogramado = False: nInteresCompAFecha = 0: nInteresGraciaAFecha = 0: nInteresMoratorio = 0: nInteresCompVencAFecha = 0
    nSegDesgAnt = 0: nSegIncAnt = 0: nSegIncGraciaAnt = 0
    nOpCovid = 0: nTasaEspCovid = 0 'Joep20200910 Tasa Especial covid
    
    '2.- Fecha de desembolso + los días del periodo de gracia
    dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
    nMes = Month(dDesembolso): nAnio = Year(dDesembolso)
    
    Dim bEsPagoAnticipado As Boolean
    Dim PlazoPre As Integer
    bEsPagoAnticipado = False
    '* Para: Pago Anticipado
   
    '** Fechas: Periodo Fijo / Fecha Fija (**Pago Anticipado)
    If pnTipoPeriodo = PeriodoFijo Then
           'Para Pagos Anticipados
             If i = 0 And psMensFecAprob <> "12:00:00 AM" Then
               'Verifica Si tiene Gracia / Diferencia es -
                 If i = 0 And (dDesembolso > psMensFecAprob) Then
                    If lnDescripCuota(0) = "0" Then
                        nPeriodoPrimerPF = pnPeriodo - CInt(DateDiff("d", psMensFecAprob, dDesembolso))
                    Else
                        If lnDescripCuota(1) <> "0" Then
                            nPeriodoPrimerPF = pnPeriodo + Val(lnDescripCuota(1))
                        End If
                    End If
                 End If
             End If
    Else
        nMes = Month(dDesembolso): nAnio = Year(dDesembolso)
    End If
    
    '*** Para: Reprogramación de créditos
    'Set oNCOMCredito = New COMNCredito.NCOMCredito
    'Set rsDatosReprogramados = New ADODB.Recordset
    'Set rsDatosReprogramados = oNCOMCredito.RecuperaDatosReprogramacionCalendario(psCtaCod)
    'If Not rsDatosReprogramados.BOF And Not rsDatosReprogramados.EOF And psCtaCod <> "" Then
        bEsReprogramado = True
        'dFecVencUltimoPago = rsDatosReprogramados!dFecVencUltimoPago 'Fecha ultima cuota pagada
        'dFechaCuotaPend = rsDatosReprogramados!dFechaCuotaPend       'Fecha Cuota Pendiente
        'dDesembolso = rsDatosReprogramados!dFecVencReprog            'Fecha Cuota Pendiente + "n" Dias Reprog.
        'pnDiaFijo = Day(dDesembolso): nMes = Month(dDesembolso): nAnio = Year(dDesembolso)
        'nPrimaPerGracia = rsDatosReprogramados!nPrimaPerGracia
        
        'Parametros de liquidación de deuda
        If IsArray(pvArrayDatos) Then
            For iMat = 0 To UBound(pvArrayDatos) - 1
                Select Case iMat
                    Case 0: nInteresCompAFecha = pvArrayDatos(iMat)
                    Case 1: nInteresGraciaAFecha = pvArrayDatos(iMat)
                    Case 2: nInteresMoratorio = pvArrayDatos(iMat)
                    Case 3: nInteresCompVencAFecha = pvArrayDatos(iMat)
                    Case 4: nSegDesgAnt = pvArrayDatos(iMat)
                    
                    'Case 5: nSegIncAnt = pvArrayDatos(iMat)        RIRO 20210215 COMENTADO
                    'Case 6: nSegIncGraciaAnt = pvArrayDatos(iMat)  RIRO 20210215 COMENTADO
                    Case 5: nPolizaMen = pvArrayDatos(iMat) 'ADD RIRO 20210212
                    Case 6: nPolizaCuotReprog = pvArrayDatos(iMat) 'ADD RIRO 20210212
                    
                    Case 7: nTasaEspCovid = pvArrayDatos(iMat) 'Joep20200910 Tasa Especial covid
                    Case 8: nOpCovid = pvArrayDatos(iMat) 'Joep20200910 Tasa Especial covid
                    Case 9: nPolizaProrrateo = pvArrayDatos(iMat)
                    Case 10: dFecVencUltimoPago = pvArrayDatos(iMat)
                    Case 11: dFechaCuotaPend = pvArrayDatos(iMat)
                    Case 12: dDesembolso = pvArrayDatos(iMat)
                    Case 13: nPrimaPerGracia = pvArrayDatos(iMat)
                    Case 14: nMontoCuota = pvArrayDatos(iMat)
                End Select
            Next
        End If
        pnDiaFijo = Day(dDesembolso): nMes = Month(dDesembolso): nAnio = Year(dDesembolso)
        'Total Intereses Liquidados
        nTotalIntLiquidados = nInteresCompAFecha + nInteresGraciaAFecha + nInteresMoratorio + nInteresCompVencAFecha
    'End If
    'Set rsDatosReprogramados = Nothing
    'Set oNCOMCredito = Nothing
    
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
    'Call oNCOMGasto.CargaValoreSegDes(psCtaCod, pdFecDesemb, MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, nCoberMinSegDes, nCoberMaxSegDes)'Comento JOEP20200414 covid
    Call oNCOMGasto.CargaValoreSegDes(psCtaCod, pdFecDesemb, MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, nCoberMinSegDes, nCoberMaxSegDes, , bEsReprogramado) 'Add JOEP20200414 covid
    Set oNCOMGasto = Nothing
    
    '3.3.- Cuota Fija Mensual Inicial o temporal (para calendario FechaFija/PeriodoFijo)
    Set oNCOMCredito = New NCOMCredito
    If (nOpCovid = 2 Or nOpCovid = 4) Then
        nMontoCuotaTemp = nMontoCuota
    Else
        nMontoCuotaTemp = oNCOMCredito.CuotaFijaTipoPeriodo(nTEMTotal, pdFecDesemb, pnNroCuotas, pnMonto, pnTipoPeriodo, pnPeriodo, pnDiaFijo, _
                                                        pnDiasGracia, pnNumMes, nTotalDias, nSumaFAS, bEsReprogramado, bEsPagoAnticipado, dFechaCuotaPend, IIf(bEsReprogramado, dFecVencUltimoPago, dDesembolsoNuevo))
    End If
    '3.4.- Factor Valor Actual Saldo
    nFVAS = Round(((1 + nTEDTotal) ^ nTotalDias), 6)
    
    '3.5.- Interes de gracia por periodo (por días de gracia)
    If Not bEsReprogramado Then
        nIntPeriodoGracia = oNCOMCredito.MontoIntPerDias(pnTasaInt, pnDiasGracia, pnMonto)
        nAjusteCuotaGracia = Round(nIntPeriodoGracia / nSumaFAS, 2)
        Set oNCOMCredito = Nothing
    Else
        nIntPeriodoGracia = nTotalIntLiquidados
        nAjusteCuotaGracia = 0
    End If
    
'    '3.6.- Obtiene la Prima del periodo de gracia de manera diaria. (prorrateada según numero de cuotas)
'    If Not bEsPagoAnticipado And Not bEsReprogramado Then
'        nPrimaPerGracia = (pnMontoPoliza * (pnDiasGracia / 30)) / pnNroCuotas
'    End If
    
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
                If bEsPagoAnticipado = True Then
                    Calendario(i).dFecha = dFecTemp + IIf(i = 0, PlazoPre, pnPeriodo)
                ElseIf bEsReprogramado And (i = 0) Then
                    Calendario(i).dFecha = dFecTemp
                Else
                    Calendario(i).dFecha = dFecTemp + IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo)
                End If
             Else
                 nDia = pnDiaFijo
                 If i = 0 Then
                    If bEsReprogramado Then
                        dFecTemp = dDesembolso
                    Else
                        dFecTemp = dDesembolso + 30
                    End If
                    nDia = Day(dFecTemp): nMes = Month(dFecTemp): nAnio = Year(dFecTemp)
                    
                    'Artificio para cuando el desembolso sea a fines de enero [Febrero]
                    If Not bEsReprogramado Then
                        If CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio))) - dDesembolso < 5 Then
                            nMes = nMes + pnNumMes
                            oNCOMCredito.ValidaFechasFijasCuota nDia, nMes, nAnio
                        End If
                    End If
                 Else
                     If nDia = pnDiaFijo Then nMes = nMes + pnNumMes
                     oNCOMCredito.ValidaFechasFijasCuota nDia, nMes, nAnio
                 End If
                 
                 'Si el cliente realizó pagos, solo toma la fecha pendiente de pago
                 If (psCtaCod <> "" And bEsPagoAnticipado = True) Then
                     If i = 0 Then
                        nDia = Day(dFechaCuotaPend)
                        nMes = Month(dFechaCuotaPend)
                        nAnio = Year(dFechaCuotaPend)
                         If (dFechaCuotaPend > dDesembolso) And (dDesembolsoNuevo > dDesembolso) Then
                                dDesembolso = dDesembolsoNuevo
                         End If
                     End If
                 End If
                 dFecTemp = CDate(Format(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)), "dd/MM/yyyy"))
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
                
                'PagoAnt. Reduccion Nro. Cuota o Plazo
                If pbReduccionNroCuotaPagAnt = True Then
                    If (nSaldoCapital >= nMontoCuotaDesembolso) Then
                        nMontoCuotaPrevia = nMontoCuotaDesembolso
                    End If
                End If
                Calendario(i).Cuota = nMontoCuotaPrevia
                
                nDiasCalculo = 0 'RIRO20200904
                
                'Interes Compensatorio Calc.
                If bEsPagoAnticipado = True Then 'Para Pago Anticipado
                    nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, PlazoPre, nSaldoCapital)
                    nDiasCalculo = PlazoPre 'RIRO 20200904
                Else
                    'JOEP20200320 comento Mejora Reprogramacion
                    'nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFecVencUltimoPago, dFechaCuotaPend), DateDiff("d", dDesembolso, Calendario(i).dFecha))), nSaldoCapital)
                    'JOEP20200320 comento Mejora Reprogramacion
                    
                    'nDiasCalculo = IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFechaCuotaPend, dDesembolso), DateDiff("d", dDesembolso, Calendario(i).dFecha))) 'RIRO 20200904
                    
                    'JOEP20200320 Add Mejora Reprogramacion
                    If bEsReprogramado = True Then
                        If nOpCovid = 3 Then
                            nIntCompCalculado = oNCOMCredito.MontoIntPerDias(nTasaEspCovid, DateDiff("d", dFechaCuotaPend, dDesembolso), nSaldoCapital) 'Cambio JOEP_RIRO_20200913
                            nDiasCalculo = DateDiff("d", dFechaCuotaPend, dDesembolso) 'RIRO 20200904
                        Else
                        'nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFechaCuotaPend, dDesembolso), DateDiff("d", dDesembolso, Calendario(i).dFecha))), nSaldoCapital)'COmento JOEP_RIRO_20200913
                            nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dFechaCuotaPend, dDesembolso), nSaldoCapital) 'Cambio JOEP_RIRO_20200913
                            nDiasCalculo = DateDiff("d", dFechaCuotaPend, dDesembolso) 'RIRO 20200904
                        End If
                    Else
                        nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFecVencUltimoPago, dFechaCuotaPend), DateDiff("d", dDesembolso, Calendario(i).dFecha))), nSaldoCapital)
                        nDiasCalculo = IIf(pnTipoPeriodo = PeriodoFijo, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), IIf(bEsReprogramado, DateDiff("d", dFechaCuotaPend, dDesembolso), DateDiff("d", dDesembolso, Calendario(i).dFecha))) 'RIRO 20200904
                    End If
                    'JOEP20200320 Add Mejora Reprogramacion
                End If
                
                'RIRO 20200825 mejora en Liquidación *************
                Calendario(i).IntCompCalculado = nIntCompCalculado
                Calendario(i).DiasCalculo = nDiasCalculo
                'END RIRO ****************************************
                 
                'Seguro Desgravamen
                Set oNCOMGasto = New COMNCredito.NCOMGasto
                nMontoSegDes = oNCOMGasto.CalculaSeguroDesgravamen(i, "K", pnTasaSegDes, pnMonto, Calendario(i).IntComp, nSaldoCapital, _
                                            Calendario(i).dFecha, 0, 0, MatPersSegDes, nEdadMinSegDes, nEdadMaxSegDes, _
                                            nCoberMinSegDes, nCoberMaxSegDes, pMatCalendSegDes, psCtaCod)
                                          
                'Si realizó pago de seg. des. en el pago anticipado no cobrarlo en el nuevo calend.
                Set oDCOMCalendario = New COMDCredito.DCOMCalendario
                If oDCOMCalendario.VerificaCuotaGastoPagadoCalendAnt(psCtaCod, Year(Calendario(i).dFecha), Month(Calendario(i).dFecha)) Then
                    nMontoSegDes = 0
                End If
                'Exoneración seguro de desgravamen(Cuota Inicial)
                If pnExoSeguroDesgravamen = 1 Then
                    Calendario(i).SegDes = Format(0, "#0.00")
                Else
                    'Calendario(i).SegDes = Round((nMontoSegDes / 30) * (DateDiff("d", pdFecDesemb, Calendario(i).dFecha)), 2) 'JOEP20200320 Comento Reprogramacion
                    If bEsReprogramado = True Then
                        Calendario(i).SegDes = Round((nMontoSegDes / 30) * (DateDiff("d", dFechaCuotaPend, Calendario(i).dFecha)), 2) 'JOEP20200320 Add Comento Reprogramacion
                    Else
                        Calendario(i).SegDes = Round((nMontoSegDes / 30) * (DateDiff("d", pdFecDesemb, Calendario(i).dFecha)), 2)
                    End If
                    
                    If bEsReprogramado = True Then
                        Calendario(i).SegDes = Calendario(i).SegDes + nSegDesgAnt
                    End If
                End If
                
                'Gasto: Póliza contra incendio
                Calendario(i).CuotaPrimaPoliza = Round(nPolizaCuotReprog, 2) ' RIRO 20210215 se inserta el monto de la póliza del primer mes, concepto 1231
'                Calendario(i).CuotaPrimaPoliza = pnMontoPoliza
'                If bEsReprogramado = True Then
'                    Calendario(i).CuotaPrimaPoliza = Calendario(i).CuotaPrimaPoliza + nSegIncAnt
'                End If
                
                'Gasto: Póliza contra incendio por periodo de gracia (prorrateado)
                Calendario(i).CuotaPrimaPolizaGracia = Round(nPolizaProrrateo, 2) 'RIRO 20210215 Se inserta la póliza prorrateada, concepto  1279
'                Calendario(i).CuotaPrimaPolizaGracia = Round(nPrimaPerGracia, 2)
'                If bEsReprogramado = True Then
'                    Calendario(i).CuotaPrimaPolizaGracia = Calendario(i).CuotaPrimaPolizaGracia + nSegIncGraciaAnt
'                End If
                
                'Interés de Gracia Cuota Inicial
                'JOEP20200321 Comento Mejora Reprogramacion
                'nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(pnCuotaIni, pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", dDesembolso, Calendario(i).dFecha)), _
                '                                                        nIntPeriodoGracia, nMontoCuotaPrevia, Calendario(i).IntComp, Calendario(i).SegDes, (Calendario(i).CuotaPrimaPoliza + Calendario(i).CuotaPrimaPolizaGracia), _
                '                                                       nGraciaGenerada, nGraciaCapitalizada, nGraciaNoPagada)
                'JOEP20200321 Comento Mejora Reprogramacion
                'JOEP20200321 Add Mejora Reprogramacion
                If bEsReprogramado = True Then
                    'nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(pnCuotaIni, pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", dFechaCuotaPend, dDesembolso)),'Comento JOEP_RIRO_20200913
                    'JOEP_RIRO_20200913
                    nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(pnCuotaIni, pnTasaInt, DateDiff("d", dFechaCuotaPend, dDesembolso), _
                                                                         nIntPeriodoGracia, nMontoCuotaPrevia, Calendario(i).IntComp, Calendario(i).SegDes, (Calendario(i).CuotaPrimaPoliza + Calendario(i).CuotaPrimaPolizaGracia), _
                                                                         nGraciaGenerada, nGraciaCapitalizada, nGraciaNoPagada)
                    'JOEP_RIRO_20200913
                Else
                    nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(pnCuotaIni, pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", dDesembolso, Calendario(i).dFecha)), _
                                                                         nIntPeriodoGracia, nMontoCuotaPrevia, Calendario(i).IntComp, Calendario(i).SegDes, (Calendario(i).CuotaPrimaPoliza + Calendario(i).CuotaPrimaPolizaGracia), _
                                                                         nGraciaGenerada, nGraciaCapitalizada, nGraciaNoPagada)
                End If
                'JOEP20200321 Add Mejora Reprogramacion
                
                'RIRO 20200904 Grabando datos ******
                Calendario(i).IntGraciaGenerado = nGraciaGenerada
                Calendario(i).IntGraciaCapitalizado = nGraciaCapitalizada
                Calendario(i).IntGraciaAsignado = nIntGraciaAsignado
                'END RIRO 20200904 *****************
                
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
                Set oDCOMCalendario = Nothing
             Else
                 'PagoAnt. Reduccion Nro. Cuota o Plazo
                 If pbReduccionNroCuotaPagAnt = True Then
                    If (nSaldoCapital >= nMontoCuotaDesembolso) Then
                        nMontoCuotaPrevia = nMontoCuotaDesembolso
                    End If
                 End If
                 
                 'Monto de la Cuota Fija Final
                 Calendario(i).Cuota = nMontoCuotaPrevia
                 nDiasCalculo = 0 'RIRO20200904
                 
                 'Interés Compensatorio
                 If pnTipoPeriodo = PeriodoFijo Then
                    If bEsPagoAnticipado = True Then
                        nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital)
                        nDiasCalculo = pnPeriodo
                    Else
                        nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo), nSaldoCapital)
                        nDiasCalculo = IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo)
                    End If
                 Else
                    nIntCompCalculado = oNCOMCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha), nSaldoCapital)
                    nDiasCalculo = DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha)
                 End If
                 
                'RIRO 20200825 mejora en Liquidación *************
                Calendario(i).IntCompCalculado = nIntCompCalculado
                Calendario(i).DiasCalculo = nDiasCalculo
                'END RIRO ****************************************
             
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
                Calendario(i).CuotaPrimaPoliza = Round(nPolizaMen, 2) 'ADD RIRO 20210212, Concepto 1231
                'Calendario(i).CuotaPrimaPoliza = pnMontoPoliza
                
                'Gasto: Póliza contra incendio por periodo de gracia (prorrateado)
                Calendario(i).CuotaPrimaPolizaGracia = Round(nPolizaProrrateo, 2) 'ADD RIRO 20210212, Concepto 1279
                'Calendario(i).CuotaPrimaPolizaGracia = Round(nPrimaPerGracia, 2)
                                
                'Interés de Gracia
                nIntGraciaAsignado = oNCOMCredito.ObtieneValoresGraciaCalendarioIterativo(i, pnTasaInt, IIf(pnTipoPeriodo = PeriodoFijo, pnPeriodo, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha)), _
                                                                                            nIntPeriodoGracia, nMontoCuotaPrevia, Calendario(i).IntComp, Calendario(i).SegDes, (Calendario(i).CuotaPrimaPoliza + Calendario(i).CuotaPrimaPolizaGracia), _
                                                                                            nGraciaGenerada, nGraciaCapitalizada, nGraciaNoPagada)
                
                'RIRO 20200904 Grabando datos ******
                Calendario(i).IntGraciaGenerado = nGraciaGenerada
                Calendario(i).IntGraciaCapitalizado = nGraciaCapitalizada
                Calendario(i).IntGraciaAsignado = nIntGraciaAsignado
                'END RIRO 20200904 *****************
                
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
                Calendario(i).IntCompDiferenciaCapitalizado = nIntCapitalizadoDiferencia 'RIRO 20200904
                
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
                 If bEsPagoAnticipado = True Then 'Pago anticipado
                     dFecTemp = dFecTemp + IIf(i = 0, PlazoPre, pnPeriodo)
                 Else
                     dFecTemp = dFecTemp + IIf(nPeriodoPrimerPF > 0, nPeriodoPrimerPF, pnPeriodo)
                 End If
                 nPeriodoPrimerPF = 0
            End If
         Next i
        
        If (nOpCovid = 2 Or nOpCovid = 4) Then
            If nSaldoCapital > 0 Then
                pnCuotaFin = pnCuotaFin + 1 'JOEP20200429 Covid Cuotas iguales
            End If
            nVASKU = Round((nSaldoCapital / nFVAS), 2)
            If nSumaFAS = 0 Then
                nCVASKU = 0
            Else
                nCVASKU = Round((nVASKU * (1 / Round(nSumaFAS, 6))), 2)
            End If
            If nSaldoCapital < 0 Then
                nSKU = Round(-1, 2)
            Else
                nSKU = Round(nSaldoCapital, 2)
            End If
            If nSumaFAS = 0 Then
                nValorAjusteIteracion = 0
            Else
                nValorAjusteIteracion = MetodoAjustePorIteracion(nSKU, nFVAS, nSumaFAS, nContadorSKU)
            End If
            nValorAjusteIteracionTotal = nValorAjusteIteracionTotal + nValorAjusteIteracion
        Else
            nVASKU = Round((nSaldoCapital / nFVAS), 2)
            nCVASKU = Round((nVASKU * (1 / Round(nSumaFAS, 6))), 2)
            nSKU = Round(nSaldoCapital, 2)
            
            nValorAjusteIteracion = MetodoAjustePorIteracion(nSKU, nFVAS, nSumaFAS, nContadorSKU)
            nValorAjusteIteracionTotal = nValorAjusteIteracionTotal + nValorAjusteIteracion
            End If
            
            If (nOpCovid = 2 Or nOpCovid = 4) Then 'Mejora reprogramacion mantener cuota JOEP20201020
                If nSaldoCapital < 0 Then
                    Exit Do
                End If
            Else
                If nNumIteracion > 17 Or pbReduccionNroCuotaPagAnt Then 'Recomendación GELU
                    Exit Do
                End If
            End If
            nContadorSKU = nContadorSKU + 1
            nNumIteracion = nNumIteracion + 1
            nSaldoCapital = 0
    Loop
    
    If (nOpCovid = 2 Or nOpCovid = 4) Then
        pvArrayDatos(14) = pnCuotaFin + 1 'Add JOEP-LUCV
    End If
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
'JOEP20201012 - Facilidades Crediticias.
