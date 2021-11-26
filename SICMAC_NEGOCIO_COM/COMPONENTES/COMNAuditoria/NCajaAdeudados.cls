VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCajaAdeudados"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A89DACF0109"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit

Enum OpeCGReportes
    OpeCGRepFlujoDiarioResMN = 461001
    OpeCGRepFlujoDiarioDetMN = 461002
    OpeCGRepRepBancosMN = 461010
    OpeCGRepRepBancosFlujoMensMN = 461011
    OpeCGRepRepBancosFlujoPFMN = 461012
    OpeCGRepRepBancosResumenPFMN = 461013
    OpeCGRepRepBancosConcentFdos = 461014
    
    OpeCGRepRepCMACSMN = 461020
    OpeCGRepRepCMACSFlujoMensMN = 461021
    OpeCGRepRepCMACSFlujoPFMN = 461022
    OpeCGRepRepCMACSResumenPFMN = 461023
    'Ordenes de Pago
    OpeCGRepRepOPMN = 461030
    OpeCGRepRepOPGirMN = 461031
    'A Rendir Cuenta
    OpeCGRepArendirLibroAuxMN = 461041
    OpeCGRepArendirPendienteMN = 461042
    OpeCGRepArendirViaticosMN = 461043
    OpeCGRepArendirViaticoLibroAuxMN = 461044
    
    'Cheques
    OpeCGRepRepChqMN = 461100
    OpeCGRepRepChqRecMN = 461110
    OpeCGRepRepChqRecDetMN = 461111
    OpeCGRepRepChqRecResMN = 461112
    OpeCGRepRepChqValMN = 461120
    OpeCGRepRepChqValDetMN = 461121
    OpeCGRepRepChqValResMN = 461122
    OpeCGRepRepChqValorizadosMN = 461130
    OpeCGRepRepChqValorizadosDetMN = 461131
    OpeCGRepRepChqValorizadosResMN = 461132
    OpeCGRepRepChqAnulMN = 461140
    OpeCGRepRepChqAnulDetMN = 461141
    OpeCGRepRepChqAnulResMN = 461142
    OpeCGRepRepChqObsMN = 461150
    OpeCGRepRepChqObsDetMN = 461151
    OpeCGRepRepChqObsResMN = 461152
    OpeCGRepChqRecibidoCajaMN = 461160
    OpeCGRepChqDepositadoMN = 461170
    
    'Cheques Remesados
    OpeCGRepChequesEnvMN = 461210
    OpeCGRepChequesAnulMN = 461220
    OpeCGRepChequesCobMN = 461230
    
    
    OpeCGAdeudRepGeneralMN = 461091
    OpeCGAdeudRepDetalleMN = 461092
    OpeCGAdeudRepSaldLinFinancDescalceMN = 461093
    OpeCGAdeudRepCortoPlazoMN = 461094
    OpeCGAdeudRepxFecVenc = 461095
    OpeCGAdeudRepVinculadosMN = 461096

    
    OpeCGAdeudRepGeneralME = 462091
    OpeCGADeudRepDetalleME = 462092
    OpeCGADeudRepDetalleCofideME = 462093
    OpeCGAdeudRepCortoPlazoME = 462094
    OpeCGAdeudRepxFecVencME = 462095
    OpeCGAdeudRepVinculadosME = 462096

'Flujo diario de Caja
    OpeCGRepSaldoBcos = 461701
    OpeCGRepSaldoCajAge = 461702
    
'MONEDA EXTRANJERA
    OpeCGRepFlujoDiarioResME = 462001
    OpeCGRepFlujoDiarioDetME = 462002
    OpeCGRepRepBancosME = 462010
    OpeCGRepRepBancosFlujoMensME = 462011
    OpeCGRepRepBancosFlujoPFME = 462012
    OpeCGRepRepBancosResumenPFME = 462013
    OpeCGRepRepCMACSME = 462020
    OpeCGRepRepCMACSFlujoMensME = 462021
    OpeCGRepRepCMACSFlujoPFME = 462022
    OpeCGRepRepCMACSResumenPFME = 462023
    'Ordenes de Pago
    OpeCGRepRepOPME = 462030
    OpeCGRepRepOPGirME = 462031
    'A Rendir Cuenta
    OpeCGRepArendirLibroAuxME = 462041
    OpeCGRepArendirPendienteME = 462042
    OpeCGRepArendirViaticosME = 462043
    OpeCGRepArendirViaticoLibroAuxME = 462044
    
    'Cheques
    OpeCGRepRepChqME = 462100
    OpeCGRepRepChqRecME = 462110
    OpeCGRepRepChqRecDetME = 462111
    OpeCGRepRepChqRecResME = 462112
    OpeCGRepRepChqValME = 462120
    OpeCGRepRepChqValDetME = 462121
    OpeCGRepRepChqValResME = 462122
    OpeCGRepRepChqValorizadosME = 462130
    OpeCGRepRepChqValorizadosDetME = 462131
    OpeCGRepRepChqValorizadosResME = 462132
    OpeCGRepRepChqAnulME = 462140
    OpeCGRepRepChqAnulDetME = 462141
    OpeCGRepRepChqAnulResME = 462142
    OpeCGRepRepChqObsME = 462150
    OpeCGRepRepChqObsDetME = 462151
    OpeCGRepRepChqObsResME = 462152
    OpeCGRepChqRecibidoCajaME = 462160
    OpeCGRepChqDepositadoME = 462170
    'Cheques Remesados
    OpeCGRepChequesEnvME = 462210
    OpeCGRepChequesAnulME = 462220
    OpeCGRepChequesCobME = 462230
    
End Enum

'Tipos de Gracia
Enum TCalendTipoGracia
    PrimeraCuota = 1
    UltimaCuota = 2
    Prorateada = 3
    Configurable = 4
    Exonerada = 5
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
    SaldoCap As Double
End Type
 
Dim Calendario() As TCalendario
Dim NroCuotas As Integer
Dim oCon      As COMConecta.DCOMConecta
Dim sSQL      As String
Dim lbConexionPropia As Boolean  'Para cerrar Conexion al cerrar Clase

Public Function GeneraGracia(ByVal pOptGracia As TCalendTipoGracia, ByVal pnInteres As Double, ByVal pnNroCuotas As Integer) As Variant
Dim nValorProrateado As Double
Dim I As Integer
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
        For I = 0 To pnNroCuotas - 1
            If I = pnNroCuotas - 1 Then
                MatGracia(I) = pnInteres - CDbl(Format((nValorProrateado * (pnNroCuotas - 1)), "#0.00"))
            Else
                MatGracia(I) = nValorProrateado
            End If
        Next I
    End If
    
    If pOptGracia = Exonerada Then
        ReDim MatGracia(pnNroCuotas)
    End If
    
    ReDim sMatGracia(UBound(MatGracia))
    For I = 0 To UBound(MatGracia) - 1
        sMatGracia(I) = Format(MatGracia(I), "#0.00")
    Next I
    GeneraGracia = sMatGracia
End Function
 
Public Function MontoIntPerDias(ByVal pnTasaInter As Double, ByVal pnDiasTrans As Integer, ByVal pnMonto As Double) As Double
    MontoIntPerDias = (((1 + pnTasaInter / 100) ^ (pnDiasTrans / 30)) - 1) * pnMonto
    MontoIntPerDias = CDbl(Format(MontoIntPerDias, "#0.00"))
End Function

'*****************************************************************
'FUNCION HALLA LA CUOTA FIJA CON FECHA FIJA
'*****************************************************************
Function CFijaFechaFija(ByVal pTasa As Double, ByVal pPlazo As Integer, ByVal pnMonto As Double, _
    ByVal pnPer As Double, ByVal pdFecDes As Date, ByVal pnDiaFijo As Integer, _
    ByVal pnDiasGracia As Integer, ByVal pnProxmes As Integer, Optional pnDeltaMes As Integer = 1) As Double
Dim I As Integer
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
    
    dFecTemp = CDate(Format(Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Trim(Str(nAnio)), "dd/mm/yyyy"))
    If dFecTemp < dFecInicio Then
        nProxMesTmp = True
    Else
        nProxMesTmp = False
    End If
    
    If pnProxmes Or nProxMesTmp Then
        nMes = nMes + pnDeltaMes
    End If
    For I = 1 To pPlazo
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
        nDiasTrans = dFecVenc - dFecInicio
        nValorActual = (1 / (1 + (nTasaDiaria / 100))) ^ (nDiasTrans)
        SumTotal = SumTotal + nValorActual
        nDia = pnDiaFijo
        nMes = nMes + pnDeltaMes
    Next I
    
    CFijaFechaFija = pnMonto / SumTotal
    CFijaFechaFija = CDbl(Format(CFijaFechaFija, "#0.00"))
    
End Function
 
Private Sub CalendarioCuotaFija(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False)
                
Dim nSaldoCapital As Double
Dim dDesembolso As Date
Dim I As Integer
Dim dFecTemp As Date
Dim nMes As Integer
Dim nAnio As Integer
Dim nDia As Integer
Dim nMontoCuotaTemp As Double
Dim nCDIntPend As Double
 
    nSaldoCapital = pnMonto
    dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
    If pnTipoPeriodo = PeriodoFijo Then
        For I = 0 To pnNroCuotas - 1
            Calendario(I).dFecha = dDesembolso + pnPeriodo
            Calendario(I).NroCuota = I + 1
            Calendario(I).IntGra = 0#
            Calendario(I).Gasto = 0#
            Calendario(I).IntComp = MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital)
            Calendario(I).Cuota = CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo)
            If I = pnNroCuotas - 1 Then
                Calendario(I).Captital = nSaldoCapital
                If (Calendario(I).IntComp + Calendario(I).Captital) > Calendario(I).Cuota Then
                    Calendario(I).Cuota = Calendario(I).Captital + Calendario(I).IntComp
                Else
                    Calendario(I).IntComp = Calendario(I).Cuota - Calendario(I).Captital
                End If
            Else
                Calendario(I).Captital = Calendario(I).Cuota - Calendario(I).IntComp
            End If
            Calendario(I).SaldoCap = nSaldoCapital - Calendario(I).Captital
            nSaldoCapital = nSaldoCapital - Calendario(I).Captital
            nSaldoCapital = CDbl(Format(nSaldoCapital, "#0.00"))
            dDesembolso = dDesembolso + pnPeriodo
        Next I
        
    Else
        If pbCuotaFijaFechaFija Then
            nMontoCuotaTemp = CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes)
        End If
        'Si es Fecha Fija
        If pnTipoPeriodo = FechaFija Then
            nMes = Month(dDesembolso)
            nAnio = Year(dDesembolso)
            nDia = pnDiaFijo
            nCDIntPend = 0
            For I = 0 To pnNroCuotas - 1
                nDia = pnDiaFijo
                If Not (I = 0 And pnDiaFijo > Day(dDesembolso) And (Not bProxMes)) Then
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
                Calendario(I).dFecha = dFecTemp
                Calendario(I).NroCuota = I + 1
                Calendario(I).IntGra = 0#
                Calendario(I).Gasto = 0
                If I = 0 Then
                    Calendario(I).IntComp = MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(I).dFecha), nSaldoCapital)
                    If Not pbCuotaFijaFechaFija Then
                        Calendario(I).Cuota = CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, DateDiff("d", dDesembolso, Calendario(I).dFecha))
                    Else
                        Calendario(I).Cuota = nMontoCuotaTemp
                    End If
                    
                Else
                    Calendario(I).IntComp = MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(I - 1).dFecha, Calendario(I).dFecha), nSaldoCapital)
                    If Not pbCuotaFijaFechaFija Then
                        Calendario(I).Cuota = CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, DateDiff("d", Calendario(I - 1).dFecha, Calendario(I).dFecha))
                    Else
                        Calendario(I).Cuota = nMontoCuotaTemp
                    End If
                End If
                
                If pbCuotaFijaFechaFija Then
                    Calendario(I).IntComp = Calendario(I).IntComp + nCDIntPend
                    nCDIntPend = 0#
                    If Calendario(I).IntComp > nMontoCuotaTemp Then
                        nCDIntPend = Calendario(I).IntComp - nMontoCuotaTemp
                        Calendario(I).IntComp = nMontoCuotaTemp
                    End If
                End If
                
                If I = pnNroCuotas - 1 Then
                    Calendario(I).Captital = nSaldoCapital
                    If (Calendario(I).IntComp + Calendario(I).Captital) > Calendario(I).Cuota Then
                        Calendario(I).Cuota = Calendario(I).Captital + Calendario(I).IntComp
                    Else
                        Calendario(I).IntComp = Calendario(I).Cuota - Calendario(I).Captital
                    End If
                Else
                    Calendario(I).Captital = Calendario(I).Cuota - Calendario(I).IntComp
                End If
                
                Calendario(I).SaldoCap = nSaldoCapital - Calendario(I).Captital
                nSaldoCapital = nSaldoCapital - Calendario(I).Captital
            Next I
        End If
    End If
    'Actualizar si existe Periodo de Gracia
    If pnDiasGracia > 0 Then
        If pnTipoGracia = PrimeraCuota Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            For I = pnNroCuotas To 1 Step -1
                Calendario(I) = Calendario(I - 1)
                Calendario(I).NroCuota = Calendario(I).NroCuota + 1
            Next I
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
            For I = 0 To pnNroCuotas - 1
                Calendario(I).IntGra = MatGracia(I)
                Calendario(I).Cuota = Calendario(I).Cuota + MatGracia(I)
            Next I
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
 
Private Sub ProcesarCalendario(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False)
                
                Select Case pnTipoCuota
                    Case Creciente
                    Case Decreciente
                    Case Fija
                        Call CalendarioCuotaFija(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pbCuotaFijaFechaFija)
                End Select
End Sub

'**********************************************************************
'FUNCION QUE ME DEVUELVE LA TASA EFECTIVA
'**********************************************************************
Function CreditosTasaEfectiva(ByVal pnCuota As Double, ByVal pnCuotas As Integer, ByVal pnMonto As Double, ByVal pnTasaIni As Double, ByVal pnPeriodo As Double) As Double
Dim I As Double
Dim nTasaTemp As Double
Dim nTasaTemp2 As Double
Dim nMontoAprox1 As Double
Dim nMontoAprox2 As Double
 
    CreditosTasaEfectiva = 0
    nMontoAprox1 = -999999
    'pnTasaIni = pnTasaIni / 100
    nTasaTemp = pnTasaIni
    For I = pnTasaIni / 100 To 1 Step 0.0000001
        nTasaTemp2 = I
        nMontoAprox2 = CDbl(Format(CuotaFija(nTasaTemp2 * 100, pnCuotas, pnMonto, pnPeriodo), "#0.00"))
        If nMontoAprox2 = pnCuota Then
            CreditosTasaEfectiva = I
            Exit For
        Else
            If nMontoAprox2 - pnCuota > 0 Then
                If Abs(nMontoAprox2 - pnCuota) > Abs(nMontoAprox1 - pnCuota) Then
                    CreditosTasaEfectiva = I - 0.0000001
                    Exit For
                Else
                    CreditosTasaEfectiva = I
                    Exit For
                End If
            End If
        End If
        nMontoAprox1 = nMontoAprox2
    Next I
End Function
                
Public Function GeneraCalendario(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, _
                Optional ByVal pbCuotaComodin As Boolean = False) As Variant
Dim sCalendPagos() As String
Dim I As Integer
Dim TasaComTemp As Double
Dim nPlazoCom As Integer
Dim nCuotasCom As Integer
Dim nCuotaMontoCom As Double
 
        On Error GoTo ErrorGeneraCalendario
 
        If pbCuotaComodin Then
            TasaComTemp = pnTasaInt
            If pnTipoPeriodo = PeriodoFijo Then
                nPlazoCom = pnPeriodo
            Else
                nPlazoCom = 30
            End If
            nCuotasCom = pnNroCuotas
            nCuotaMontoCom = CuotaFija(TasaComTemp, nCuotasCom + 1, pnMonto, nPlazoCom)
            nCuotaMontoCom = CDbl(Format((nCuotaMontoCom * (pnNroCuotas + 1)) / pnNroCuotas, "#0.00"))
            TasaComTemp = CDbl(Format(CreditosTasaEfectiva(nCuotaMontoCom, pnNroCuotas, pnMonto, pnTasaInt, nPlazoCom), "#0.0000"))
            pnTasaInt = Format(TasaComTemp * 100, "#0.0000")
        End If
 
        ReDim Calendario(pnNroCuotas)
        Call ProcesarCalendario(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoCuota, _
                pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pbCuotaFijaFechaFija)
        ReDim sCalendPagos(UBound(Calendario), 8)
        For I = 0 To UBound(Calendario) - 1
            sCalendPagos(I, 0) = Format(Calendario(I).dFecha, "dd/mm/yyyy")
            sCalendPagos(I, 1) = Trim(Str(Calendario(I).NroCuota))
            sCalendPagos(I, 2) = Format(Calendario(I).Cuota, "#0.00")
            sCalendPagos(I, 3) = Format(Calendario(I).Captital, "#0.00")
            sCalendPagos(I, 4) = Format(Calendario(I).IntComp, "#0.00")
            sCalendPagos(I, 5) = Format(Calendario(I).IntGra, "#0.00")
            sCalendPagos(I, 6) = Format(Calendario(I).Gasto, "#0.00")
            sCalendPagos(I, 7) = Format(Calendario(I).SaldoCap, "#0.00")
        Next I
        GeneraCalendario = sCalendPagos
        Exit Function
 
ErrorGeneraCalendario:
        MsgBox Err.Description, vbCritical, "Aviso"
        
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

'***********************************************
'*  FUNCION QUE HALLA EL INTERES DE UN PERIODO DE DIAS TRANACURRIDOS
'***********************************************
Public Function TasaIntPerDias(ByVal pnTasaInter As Double, ByVal pnDiasTrans As Integer) As Double
    TasaIntPerDias = ((1 + pnTasaInter / 100) ^ (pnDiasTrans / 30)) - 1
End Function

Public Function CalculaInteres(lnDiasTrans As Long, lnPeriodo As Long, lnTasaInt As Currency, lnCapital As Currency) As Currency
    Dim lnFactor As Double
    Dim lnBaseInt As Double
    If lnPeriodo > 0 Then
        lnFactor = lnDiasTrans / lnPeriodo
        lnBaseInt = ((1 + (lnTasaInt / 100)) ^ lnFactor) - 1
        CalculaInteres = lnBaseInt * lnCapital
    Else
        CalculaInteres = 0
    End If
End Function

Public Function CargaDatosAdeudadosProyectados(lsPersCod As String, Optional lsPersCodFin As String = "", Optional lnIndiceVac As Double, Optional psFiltroCta As String = "") As ADODB.Recordset
Dim lsFiltro As String
On Error GoTo CargaDatosAdeudadosProyectadosErr
    
    If lsPersCod <> "" Then
        lsFiltro = " AND ci.cPersCod = '" & Mid(lsPersCod, 4, 13) & "' "
    End If
    If lsPersCodFin <> "" Then
        If Len(lsPersCodFin) > 16 Then
            lsFiltro = " AND ci.cPersCod BETWEEN '" & Mid(lsPersCod, 4, 13) & "' and '" & Mid(lsPersCodFin, 4, 13) & "' and ci.cCtaIFCod BETWEEN '" & Mid(lsPersCod, 18, 10) & "' and '" & Mid(lsPersCodFin, 18, 10) & "'"
        Else
            lsFiltro = " AND ci.cPersCod BETWEEN '" & Mid(lsPersCod, 4, 13) & "' and '" & Mid(lsPersCodFin, 4, 13) & "' "
        End If
    End If
   sSQL = "SELECT ci.cIFTpo, ci.cPersCod, ci.cCtaIFCod, P.cPersNombre, ci.cCtaIFDesc, ci.dCtaIFAper, dCtaIFVenc, cia.nMontoPrestado, ci.nCtaIFPlazo, cia.nCtaIFCuotas, cia.nPeriodoGracia, cii.nCtaIFIntPeriodo, cii.nCtaIFIntValor nTasaInteres, cia.nSaldoCap nSaldoCapCal, ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") nVac, cia.cMonedaPago, " _
        & "       Round( cia.nSaldoCap * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END,2) nSaldoCap, " _
        & "             ISNULL(cia.cCodLinCred,'') cCodLinCred, ISNULL(l.cDescripcion,'') cDesLinCred, cia.nCuotaPagoCap " _
        & "    FROM CtaIF ci LEFT JOIN CtaIfAdeudados cia ON cia.cIFTpo = ci.cIFTpo and cia.cPersCod = ci.cPersCod and cia.cCtaIFCod = ci.cCtaIFCod LEFT JOIN ColocLineaCredito l ON l.cLineaCred = cia.cCodLinCred JOIN Persona P ON ci.cPersCod = p.cPersCod " _
        & "         LEFT JOIN IndiceVac iv ON iv.dIndiceVac = ISNULL(cia.dCuotaUltPago, ci.dCtaIFAper) " _
        & "          JOIN CtaIFInteres cii ON cii.cIFTpo = ci.cIFTpo and cii.cPersCod = ci.cPersCod and cii.cCtaIFCod = ci.cCtaIFCod " _
        & "    WHERE ci.cCtaIFEstado in (" & gEstadoCtaIFActiva & "," & gEstadoCtaIFRegistrada & ") and  ci.cIFTpo+ci.cCtaIFCod LIKE '" & psFiltroCta & "' " & lsFiltro _
        & "    ORDER BY ci.cIFtpo, ci.cPersCod, ISNULL(cia.cCodLinCred,''), ci.cCtaIFDesc  "

  
    Set CargaDatosAdeudadosProyectados = oCon.CargaRecordSet(sSQL)
Exit Function
CargaDatosAdeudadosProyectadosErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CargaDatosADeudadosxEndeudamiento(psOpeCod As String, lsMoneda As String, ldFecha As Date, Optional lsPersCod As String, Optional psFecha1 As String, Optional psfecha2 As String, Optional psFecha3 As String) As ADODB.Recordset
Dim lsFiltro As String

On Error GoTo CargaDatosAdeudadosProyectadosErr
    
    If lsPersCod <> "" Then
        lsFiltro = " AND Per.cPersCod = '" & Mid(lsPersCod, 4, 13) & "' "
    End If
    
    sSQL = " select Per.cPersNombre, " & _
          " ISNULL(cia.cCodLinCred,'') cCodLinCred, ISNULL(l.cDescripcion,'') cDesLinCred," & _
          " CIF.cPersCod, CIF.cIFTpo, CIF.cCtaIFCod, CIF.cCtaIFDesc,CIF.dCtaIFAper," & _
          " case Substring(CIF.cCtaIFCod, 3,1) when '2' then 'N.DOLARES' when '1' then 'N.SOLES' END as cMoneda," & _
          " CIA.nMontoPrestado, CIA.nSaldoCap, dbo.AdeuGetFechaVencimiento(CIF.cPersCod, CIF.cIFTpo, CIF.cCtaIFCod) as cVencimiento, case cFinalidad when '1' then 'INTER.FINA'  else 'ACTIVO FIJO' END Finalidad," & _
          " datediff(m, CIF.dCtaIFAper, dbo.AdeuGetFechaVencimiento(CIF.cPersCod, CIF.cIFTpo, CIF.cCtaIFCod)) As nMeses, " & _
          " CIC2.nCantidadCuotas, cic3.nProxCuota, cii.nCtaIFIntValor nTasaInteres"
          
    sSQL = sSQL & " From ctaif CIF" & _
          " Inner Join CtaIfAdeudados CIA" & _
          " On CIF.cPersCod=CIA.cPersCod and CIF.cIFTpo= CIA.cIFTpo And CIF.cCtaIFCod= CIA.cCtaIFCod" & _
          " Inner Join Persona Per ON Per.cPersCod = CIF.cPersCod" & _
          " LEFT JOIN ColocLineaCredito l ON l.cLineaCred = cia.cCodLinCred"
   
    'Interes
    sSQL = sSQL & " JOIN CtaIFInteres cii ON cii.cIFTpo = CIF.cIFTpo and cii.cPersCod = CIF.cPersCod and cii.cCtaIFCod = CIF.cCtaIFCod "
    sSQL = sSQL & " and cii.dCtaIFIntRegistro = (SELECT Max(dCtaIFIntRegistro) "
    sSQL = sSQL & " FROM CtaIFInteres cii1 WHERE cii1.cIFTpo = cii.cIFTpo "
    sSQL = sSQL & " and cii1.cPersCod = cii.cPersCod and cii1.cCtaIFCod = cii.cCtaIFCod ) "
      
    'cic2.nCantidadCuotas
    sSQL = sSQL & " Inner Join ( Select CIC1.cPersCod, CIC1.cCtaIFCod, CIC1.cIFTpo, COUNT(*) as nCantidadCuotas"
    sSQL = sSQL & " From CtaIFCalendario CIC1 where ctpocuota='2' "
    sSQL = sSQL & " Group By CIC1.cPersCod, CIC1.cCtaIFCod, CIC1.cIFTpo "
    sSQL = sSQL & " ) CIC2  ON CIC2.cPersCod = cif.cPersCod and CIC2.cCtaIFCod = ciF.cCtaIFCod and CIC2.cIFTpo = ciF.cIFTpo "
    
    'cic3.nProxCuota
    sSQL = sSQL & " Inner Join ( Select CIC1.cPersCod, CIC1.cCtaIFCod, CIC1.cIFTpo, CIC1.nCapital + CIC1.nInteres + CIC1.nComision as nProxCuota "
    sSQL = sSQL & " From CtaIFCalendario CIC1 "
    sSQL = sSQL & " Where CIC1.cTpoCuota='2' and CIC1.nNroCuota = "
    sSQL = sSQL & " ( "
    sSQL = sSQL & " Select MIN(CIC1A.nNroCuota) as nProxCuota "
    sSQL = sSQL & " From CtaIFCalendario CIC1A "
    sSQL = sSQL & " where CIC1A.ctpocuota='2' and convert(varchar(8), dvencimiento, 112)>'" & Format(ldFecha, "YYYYMMdd") & "' "
    sSQL = sSQL & " and CIC1A.cPersCod=CIC1.cPersCod and  CIC1A.cCtaIFCod=CIC1.cCtaIFCod and  CIC1A.cIFTpo = CIC1.cIFTpo "
    sSQL = sSQL & " Group By CIC1A.cPersCod, CIC1A.cCtaIFCod, CIC1A.cIFTpo "
    sSQL = sSQL & " )"
    sSQL = sSQL & " ) CIC3  ON CIC3.cPersCod = cif.cPersCod and CIC3.cCtaIFCod = ciF.cCtaIFCod and CIC3.cIFTpo = ciF.cIFTpo "

    sSQL = sSQL & "         LEFT JOIN ( SELECT mif.cPersCod, mif.cCtaIFCod, mif.cIFTpo, SUM(" & IIf(Mid(psOpeCod, 3, 1) = "2", "me.nMovMEImporte", "mc.nMovImporte") & ") nMonto " _
        & "                     FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro " & IIf(Mid(psOpeCod, 3, 1) = "2", "JOIN MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
        & "                                JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' " _
        & "                     WHERE m.nMovFlag = 0 and LEFT(m.cmovnro,8) > '" & Format(ldFecha, gsFormatoMovFecha) & "' and oc.copecod = '" & psOpeCod & "' and oc.cOpeCtaOrden = '0' " _
        & "                     GROUP BY mif.cPersCod, mif.cCtaIFCod, mif.cIFTpo ) Canc ON canc.cPersCod = cif.cPersCod and canc.cCtaIFCod = ciF.cCtaIFCod and canc.cIFTpo = ciF.cIFTpo "
        
    sSQL = sSQL & " Where (CIF.cctaifestado in(" & gEstadoCtaIFActiva & "," & gEstadoCtaIFRegistrada & ") or not Canc.nMonto is NULL ) " & _
          " And datediff(d, '" & Format(ldFecha, "MM/dd/YYYY") & "', dbo.AdeuGetFechaVencimiento(CIF.cPersCod, CIF.cIFTpo, CIF.cCtaIFCod)) > 360" & _
          " and convert(varchar(8), dbo.AdeuGetFechaVencimiento(CIF.cPersCod, CIF.cIFTpo, CIF.cCtaIFCod), 112) >='" & Format(ldFecha, "YYYYMMdd") & "' and Substring(CIF.cCtaIFCod, 3,1)='" & lsMoneda & "' " & lsFiltro & _
          " and convert(varchar(8), CIF.dCtaIFAper, 112)<='" & Format(ldFecha, "YYYYMMdd") & "' ORDER BY cif.cIFtpo, cif.cPersCod, ISNULL(cia.cCodLinCred,''), cif.cCtaIFDesc "
          
 Set CargaDatosADeudadosxEndeudamiento = oCon.CargaRecordSet(sSQL)
Exit Function
CargaDatosAdeudadosProyectadosErr:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function CargaDatosGeneralesCtaIF(lsPersCod As String, lsFiltroCta As String, Optional lsPersCodFin As String = "", Optional lnIndiceVac As Double, Optional pdFecha As Date, Optional psOpeCod As String = "") As ADODB.Recordset
Dim lsFiltro As String
On Error GoTo CargaDatosGeneralesCtaIFErr
    
    If lsPersCod <> "" Then
        lsFiltro = " AND ci.cPersCod = '" & Mid(lsPersCod, 4, 13) & "' "
    End If
    If lsPersCodFin <> "" Then
        If Len(lsPersCodFin) > 16 Then
            lsFiltro = " AND ci.cPersCod BETWEEN '" & Mid(lsPersCod, 4, 13) & "' and '" & Mid(lsPersCodFin, 4, 13) & "' and ci.cCtaIFCod BETWEEN '" & Mid(lsPersCod, 18, 10) & "' and '" & Mid(lsPersCodFin, 18, 10) & "'"
        Else
            lsFiltro = " AND ci.cPersCod BETWEEN '" & Mid(lsPersCod, 4, 13) & "' and '" & Mid(lsPersCodFin, 4, 13) & "' "
        End If
    End If
    '
   sSQL = " SELECT cia.nSaldoCapLP SaldoLargoPlazo, dbo.AdeuGetFechaVencimiento(ci.cPersCod, ci.cIFTpo, ci.cCtaIFCod)  FechaVencimiento, ci.cIFTpo, ci.cPersCod, ci.cCtaIFCod, P.cPersNombre, ci.cCtaIFDesc, ci.dCtaIFAper, dCtaIFVenc, cia.nMontoPrestado, ci.nCtaIFPlazo, cia.nCtaIFCuotas, cia.nPeriodoGracia, cii.nCtaIFIntPeriodo, cii.nCtaIFIntValor nTasaInteres, cic.nNroCuota, ISNULL(cic.nInteresPagado,0) + " _
        & "             ISNULL( (SELECT SUM(" & IIf(Mid(psOpeCod, 3, 1) = "2", "me.nMovMEImporte", "mc.nMovImporte") & ") " _
        & "              FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro " & IIf(Mid(psOpeCod, 3, 1) = "2", "JOIN MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
        & "                       JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' " _
        & "              WHERE m.nMovFlag = 0 and LEFT(m.cmovnro,8) > '" & Format(pdFecha, gsFormatoMovFecha) & "' and oc.copecod = '" & psOpeCod & "' and oc.cOpeCtaOrden = '3' and mif.cPersCod = ci.cPersCod and mif.cCtaIFCod = ci.cCtaIFCod and mif.cIFTpo = ci.cIFTpo " _
        & "             ), 0) nInteresPagado, cic.dVencimiento, cia.nSaldoCap nSaldoCapCal, ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") nVac, cic.cMovNro, cia.cMonedaPago, " _
        & "       Round( cia.nSaldoCap * convert(decimal(15,5), CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END),2) + " _
        & "             ISNULL( Canc.nMonto, 0) nSaldoCap , ISNULL(cia.cCodLinCred,'') cCodLinCred, ISNULL(l.cDescripcion,'') cDesLinCred, cia.nCuotaPagoCap, " _
        & "            IsNull((SELECT SUM(me.nMovMEImporte)  FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro JOIN MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' WHERE m.nMovFlag = 0 and LEFT(m.cmovnro,8) < '" & Format(pdFecha, gsFormatoMovFecha) & "' and oc.copecod = '" & psOpeCod & "' and oc.cOpeCtaOrden = '3' and mif.cPersCod = ci.cPersCod and mif.cCtaIFCod = ci.cCtaIFCod and mif.cIFTpo = ci.cIFTpo),0) IntDevAnt" _
        & "    FROM CtaIF ci LEFT JOIN CtaIfAdeudados cia ON cia.cIFTpo = ci.cIFTpo and cia.cPersCod = ci.cPersCod and cia.cCtaIFCod = ci.cCtaIFCod LEFT JOIN ColocLineaCredito l ON l.cLineaCred = cia.cCodLinCred JOIN Persona P ON ci.cPersCod = p.cPersCod " _
        & "         LEFT JOIN IndiceVac iv ON iv.dIndiceVac = ISNULL(cia.dCuotaUltPago, ci.dCtaIFAper) JOIN CtaIFInteres cii ON cii.cIFTpo = ci.cIFTpo and cii.cPersCod = ci.cPersCod and cii.cCtaIFCod = ci.cCtaIFCod " _
        & "              and cii.dCtaIFIntRegistro = (SELECT Max(dCtaIFIntRegistro) " _
        & "                                     FROM CtaIFInteres cii1 WHERE cii1.cIFTpo = cii.cIFTpo " _
        & "                                      and cii1.cPersCod = cii.cPersCod and cii1.cCtaIFCod = cii.cCtaIFCod ) " _
        & "         LEFT JOIN CtaIFCalendario cic ON cic.cIFTpo = ci.cIFTpo and cic.cPersCod = ci.cPersCod and cic.cCtaIFCod = ci.cCtaIFCod " _
        & "              and cic.cTpoCuota = '2' and cic.nNroCuota = (SELECT Min(nNroCuota) FROM CtaIFCalendario cic1 " _
        & "                        Where cic1.cIFTpo = cic.cIFTpo And cic1.cPersCod = cic.cPersCod And cic1.cCtaIFCod = cic.cCtaIFCod " _
        & "                          and cic1.cTpoCuota = cic.cTpoCuota and cEstado = 0 ) " _
        & "         LEFT JOIN ( SELECT mif.cPersCod, mif.cCtaIFCod, mif.cIFTpo, SUM(" & IIf(Mid(psOpeCod, 3, 1) = "2", "me.nMovMEImporte", "mc.nMovImporte") & ") nMonto " _
        & "                     FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro " & IIf(Mid(psOpeCod, 3, 1) = "2", "JOIN MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
        & "                                JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' " _
        & "                     WHERE m.nMovFlag = 0 and LEFT(m.cmovnro,8) > '" & Format(pdFecha, gsFormatoMovFecha) & "' and oc.copecod = '" & psOpeCod & "' and oc.cOpeCtaOrden = '0' " _
        & "                     GROUP BY mif.cPersCod, mif.cCtaIFCod, mif.cIFTpo ) Canc ON canc.cPersCod = ci.cPersCod and canc.cCtaIFCod = ci.cCtaIFCod and canc.cIFTpo = ci.cIFTpo " _
        & "    WHERE   ci.cIFTpo+ci.cCtaIFCod LIKE '" & lsFiltroCta & "' and datediff(d,ci.dCtaIFAper,'" & Format(pdFecha, gsFormatoFecha) & "') >= 0 " & lsFiltro & " "
        
        '(ci.cCtaIFEstado in (" & gEstadoCtaIFActiva & "," & gEstadoCtaIFRegistrada & ") or not Canc.nMonto is NULL ) and
    sSQL = sSQL & "    ORDER BY ci.cIFtpo, ci.cPersCod, ISNULL(cia.cCodLinCred,''), ci.cCtaIFDesc  "

    'nSaldoCap
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set CargaDatosGeneralesCtaIF = oCon.CargaRecordSet(sSQL)
Exit Function
CargaDatosGeneralesCtaIFErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CargaDatosGeneralesCtaIFPlazoFijo(lsPersCod As String, lsFiltroCta As String, psOpeCod As String, Optional lsPersCodFin As String = "") As ADODB.Recordset
Dim lsFiltro As String
On Error GoTo CargaDatosGeneralesCtaIFPlazoFijoErr
    
    If lsPersCod <> "" Then
        lsFiltro = " AND ci.cPersCod = '" & Mid(lsPersCod, 4, 13) & "' "
    End If
    If lsPersCodFin <> "" Then
        If Len(lsPersCodFin) > 16 Then
            lsFiltro = " AND ci.cPersCod BETWEEN '" & Mid(lsPersCod, 4, 13) & "' and '" & Mid(lsPersCodFin, 4, 13) & "' and ci.cCtaIFCod BETWEEN '" & Mid(lsPersCod, 18, 10) & "' and '" & Mid(lsPersCodFin, 18, 10) & "'"
        Else
            lsFiltro = " AND ci.cPersCod BETWEEN '" & Mid(lsPersCod, 4, 13) & "' and '" & Mid(lsPersCodFin, 4, 13) & "' "
        End If
    End If
   sSQL = "SELECT ci.cIFTpo, ci.cPersCod, ci.cCtaIFCod, " _
        & "       P.cPersNombre, ci.cCtaIFDesc, ci.dCtaIFAper, ISNULL(dCtaIFVenc,ci.dCtaIFAper) dCtaIFVenc, cii.nCtaIFIntPeriodo, cii.nCtaIFIntValor nTasaInteres, ci.nInteres nInteresPagado, ci.dCtaIFVenc dVencimiento, cis.nSaldo nMontoIni, cis.nSaldoME nMontoIniME " _
        & "FROM CtaIF ci " _
        & "     JOIN Persona P ON ci.cPersCod = p.cPersCod " _
        & "     JOIN CtaIFInteres cii ON cii.cIFTpo = ci.cIFTpo and cii.cPersCod = ci.cPersCod and cii.cCtaIFCod = ci.cCtaIFCod " _
        & "      and cii.dCtaIFIntRegistro = (SELECT Max(dCtaIFIntRegistro) " _
        & "                                     FROM CtaIFInteres cii1 WHERE cii1.cIFTpo = cii.cIFTpo " _
        & "                                      and cii1.cPersCod = cii.cPersCod and cii1.cCtaIFCod = cii.cCtaIFCod ) " _
        & "     JOIN CtaIFSaldo cis ON cis.cIFTpo = ci.cIFTpo and cis.cPersCod = ci.cPersCod and cis.cCtaIFCod = ci.cCtaIFCod " _
        & "      and cis.dCtaIFSaldo = (SELECT Min(dCtaIFSaldo) " _
        & "                             FROM CtaIFSaldo cis1 WHERE cis1.cIFTpo = cis.cIFTpo " _
        & "                              and cis1.cPersCod = cis.cPersCod and cis1.cCtaIFCod = cis.cCtaIFCod and cis1.cCtaContCod = cis.cCtaContCod ) " _
        & "     JOIN OpeCta oc ON cis.cCtaContCod LIKE oc.cCtaContCod + '%' " _
        & " WHERE ci.cIFTpo+ci.cCtaIFCod LIKE '" & lsFiltroCta & "' " & lsFiltro & " and oc.cOpeCod = '" & psOpeCod & "' and cOpeCtaOrden = '0' and copectaopc=2 " _
        & " ORDER BY ci.cIFtpo, ci.cPersCod, ci.cCtaIFCod "

    Set CargaDatosGeneralesCtaIFPlazoFijo = oCon.CargaRecordSet(sSQL)
Exit Function
CargaDatosGeneralesCtaIFPlazoFijoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetLineaCredito(psMoneda As String, Optional psPersCod As String = "") As ADODB.Recordset
On Error GoTo GetLineaCreditoErr
Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
sSQL = "SELECT cLineaCred, RTRIM(p.cPersNombre) + ' - ' + cDesCripcion cDescripcion, len(cLineaCred) nObjNivel, c.cPersCod " _
     & "FROM colocLineaCredito c JOIN Persona p ON p.cPersCod = c.cPersCod " _
     & "WHERE bEstado = 1 and LEN(cLineaCred) <= 6 and NOT cLineaCred = '' and ( (LEN(cLineaCred) = 5 and cLineaCred LIKE '____" & psMoneda & "%') or LEN(cLineaCred) < 5) " _
     & IIf(psPersCod = "", "", " and c.cPersCod = '" & psPersCod & "' ") _
     & "ORDER BY cLineaCred "
Set GetLineaCredito = oCon.CargaRecordSet(sSQL)
oCon.CierraConexion
Exit Function
GetLineaCreditoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetLineaCreditoPersona(psLineaCod As String) As String
Dim rs As ADODB.Recordset
On Error GoTo GetLineaCreditoErr
sSQL = "SELECT c.cPersCod " _
     & "FROM colocLineaCredito c " _
     & "WHERE bEstado = 1 and cLineaCred = '" & psLineaCod & "'  "
Set rs = oCon.CargaRecordSet(sSQL)
If Not rs.EOF Then
    GetLineaCreditoPersona = rs!cPerscod
End If
RSClose rs
Exit Function
GetLineaCreditoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetAdeudadosCortoPlazo(psPersCod As String, psCtaIfCod As String, psIFTpo As String, pdFecha As Date) As Currency
Dim prs As ADODB.Recordset
On Error GoTo GetAdeudadosCortoPlazoErr
sSQL = "SELECT ISNULL(SUM(nCapital),0) nSaldoCapCorto " _
     & "FROM  CtaIFCalendario " _
     & "WHERE cPersCod = '" & psPersCod & "' and cCtaIFCod = '" & psCtaIfCod & "' and cIFTpo = '" & psIFTpo & "' and " _
     & "      dVencimiento > '" & Format(pdFecha, gsFormatoFecha) & "' and dVencimiento <= '" & Format(DateAdd("yyyy", 1, pdFecha), gsFormatoFecha) & "' "
Set prs = oCon.CargaRecordSet(sSQL)
If Not prs.EOF Then
    GetAdeudadosCortoPlazo = prs!nSaldoCapCorto
End If
Exit Function
GetAdeudadosCortoPlazoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function GetAdeudadosProvisionMes(psOpeCod As String, psMovNro As String) As Currency
Dim prs As ADODB.Recordset
On Error GoTo GetAdeudadosProvisionMesErr
sSQL = "SELECT " & IIf(Mid(psOpeCod, 3, 1) = "2", "me.nMovMEImporte", "mc.nMovImporte") & " nMovImporte " _
     & "FROM  Mov m Join MovCta mc on mc.nMovNro = m.nMovNro " _
     & IIf(Mid(psOpeCod, 3, 1) = "2", " join MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
     & "      JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' " _
     & "WHERE m.cMovNro = '" & psMovNro & "' and oc.cOpeCod = '" & psOpeCod & "' "
Set prs = oCon.CargaRecordSet(sSQL)
If Not prs.EOF Then
    GetAdeudadosProvisionMes = prs!nMovImporte
End If
Exit Function
GetAdeudadosProvisionMesErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetAdeudadosOperaciones(psOpeCod As String, pdFechaDel As Date, pdFechaAl As Date) As ADODB.Recordset

On Error GoTo GetAdeudadosOperacionesErr
Set oCon = New COMConecta.DCOMConecta
sSQL = "SELECT Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103) dFecha, p.cPersNombre, cif.cCtaIFDesc, SUM(ISNULL(me.nMovMEImporte,mc.nMovImporte)) * -1 as nImporte, m.cMovDesc, cif.cIFTpo, cif.cPersCod, cif.cCtaIFCod, m.cMovNro, m.nMovNro, CEst.dFecUltActAnt dFechaMod "
sSQL = sSQL & "FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro " _
     & "           LEFT JOIN MovMe  me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
     & "           JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem " _
     & "           JOIN CtaIF cif ON cif.cIFTpo = mif.cIFTpo and cif.cPersCod = mif.cPersCod and cif.cCtaIFCod = mif.cCtaIFCod " _
     & "           /*JOIN CtaIFcalendario ca on ca.cIFTpo = mif.cIFTpo and ca.cPersCod = mif.cPersCod and ca.cCtaIFCod = mif.cCtaIFCod  and ca.cestado=2*/ " _
     & "           JOIN Persona p ON p.cPersCod = mif.cPersCod " _
     & "           Join CtaIFEstadistica CEst On CEst.nMovNro = m.nMovNro " _
     & "WHERE m.cOpeCod = '" & psOpeCod & "' and m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag = " & gMovFlagVigente _
     & "  and LEFT(m.cMovNro,8) BETWEEN '" & Format(pdFechaDel, gsFormatoMovFecha) & "' and '" & Format(pdFechaAl, gsFormatoMovFecha) & "' "
     
     If Mid(psOpeCod, 3, 1) = "1" Then
        sSQL = sSQL & "  and mc.nMovImporte < 0 "
    Else
        sSQL = sSQL & " and ISNULL(me.nMovMEImporte,mc.nMovImporte) * -1 <0"
    End If
    
    sSQL = sSQL & " group by Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103), p.cPersNombre,"
    sSQL = sSQL & " cif.cCtaIFDesc,m.cMovDesc, cif.cIFTpo, cif.cPersCod,cif.cCtaIFCod, m.cMovNro, m.nMovNro, CEst.dFecUltActAnt  "
oCon.AbreConexion
Set GetAdeudadosOperaciones = oCon.CargaRecordSet(sSQL)
oCon.CierraConexion

Exit Function
GetAdeudadosOperacionesErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetAdeudadosPagoCuota(psOpeCod As String, pdFechaDel As Date, pdFechaAl As Date) As ADODB.Recordset
On Error GoTo GetAdeudadosPagoCuotaErr
Dim oCon As New COMConecta.DCOMConecta

sSQL = "SELECT Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103) dFecha, Adeud.cPersNombre, Adeud.cCtaIFDesc, nImporte, m.cMovDesc, Adeud.cIFTpo, Adeud.cPersCod, Adeud.cCtaIFCod, m.cMovNro, m.nMovNro, RTRIM(Pag.cPersNombre) + ' ' + Pag.cCtaIFDesc EntidadPago " _
     & "FROM Mov m JOIN " _
     & "  (SELECT DISTINCT mc.nMovNro, p.cPersNombre, cif.cCtaIFDesc, cif.cPersCod, cif.cIFTpo, cif.cCtaIFCod " _
     & "   FROM MovCta mc JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' " _
     & "        LEFT JOIN MovMe  me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
     & "        JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem " _
     & "        JOIN CtaIF cif ON cif.cIFTpo = mif.cIFTpo and cif.cPersCod = mif.cPersCod and cif.cCtaIFCod = mif.cCtaIFCod " _
     & "        JOIN Persona p ON p.cPersCod = mif.cPersCod " _
     & "   WHERE oc.cOpeCod = '" & psOpeCod & "' and oc.cOpeCtaDH = 'D' and oc.cOpeCtaOrden = '0' " _
     & "  ) Adeud ON Adeud.nMovNro = m.nMovNro " _
     & "  JOIN (SELECT mc.nMovNro, p.cPersNombre, cif.cCtaIfDesc, SUM( ISNULL(me.nMovMEImporte,mc.nMovImporte) * -1 ) as nImporte " _
     & "   FROM MovCta mc JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' " _
     & "        LEFT JOIN MovMe  me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
     & "        LEFT JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem " _
     & "        LEFT JOIN CtaIF cif ON cif.cIFTpo = mif.cIFTpo and cif.cPersCod = mif.cPersCod and cif.cCtaIFCod = mif.cCtaIFCod " _
     & "        LEFT JOIN Persona p ON p.cPersCod = mif.cPersCod " _
     & "   WHERE oc.cOpeCod = '" & psOpeCod & "' and oc.cOpeCtaDH = 'H'  GROUP BY mc.nMovNro, p.cPersNombre, cif.cCtaIfDesc " _
     & "  ) Pag ON Pag.nMovNro = m.nMovNro " _
     & "WHERE m.cOpeCod = '" & psOpeCod & "' and m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag = " & gMovFlagVigente _
     & "  and LEFT(m.cMovNro,8) BETWEEN '" & Format(pdFechaDel, gsFormatoMovFecha) & "' and '" & Format(pdFechaAl, gsFormatoMovFecha) & "' " _
     & " "
oCon.AbreConexion
Set GetAdeudadosPagoCuota = oCon.CargaRecordSet(sSQL)
oCon.CierraConexion

Exit Function
GetAdeudadosPagoCuotaErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CargaDatosAdeudoReclasifica(lsFiltroCta As String, Optional lnIndiceVac As Double, Optional pdFecha As Date, Optional psOpeCod As String = "", Optional pbSoloCortoPlazo As Boolean = False) As ADODB.Recordset
Dim lsFiltro As String
On Error GoTo CargaDatosGeneralesCtaIFErr
    
   sSQL = ""
   sSQL = sSQL & "SELECT p.cPersNombre, ci.cIFTpo, ci.cPersCod, ci.cCtaIFCod, ci.cCtaIFDesc, cia.nSaldoCap nSaldoCapCal, ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") nVac, cia.cMonedaPago,"
   sSQL = sSQL & "         round( cia.nSaldoCap * convert(decimal(12,5), CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1.0 END),2) +"
   sSQL = sSQL & "             ISNULL( Canc.nMonto, 0) nSaldoCap, ISNULL(cic.nSaldoCapCorto,0) nSaldoCapCortoCal,"
   sSQL = sSQL & "         round( ISNULL(cic.nSaldoCapCorto,0) * convert(decimal(12,5), CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1.0 END),2) nSaldoCapCorto, ISNULL(sdo.nSaldoCnt,0) nSaldoCnt "
   sSQL = sSQL & "FROM CtaIF ci JOIN CtaIfAdeudados cia ON cia.cIFTpo = ci.cIFTpo and cia.cPersCod = ci.cPersCod and cia.cCtaIFCod = ci.cCtaIFCod"
   sSQL = sSQL & "              LEFT JOIN Persona p ON p.cPersCod = ci.cPersCod "
   sSQL = sSQL & "              LEFT JOIN IndiceVac iv ON iv.dIndiceVac = ISNULL(cia.dCuotaUltPago, ci.dCtaIFAper)"
   sSQL = sSQL & "              LEFT JOIN ( SELECT cIFTpo, cPersCod, cCtaIFCod, ISNULL(SUM(nCapital),0) nSaldoCapCorto"
   sSQL = sSQL & "                From CtaIFCalendario"
   sSQL = sSQL & "                          WHERE dVencimiento > '" & Format(pdFecha, gsFormatoFecha) & "' and dVencimiento <= '" & Format(DateAdd("yyyy", 1, pdFecha), gsFormatoFecha) & "'"
   sSQL = sSQL & "                          GROUP BY cIFTpo, cPersCod, cCtaIFCod"
   sSQL = sSQL & "                        ) cic ON cic.cIFTpo = ci.cIFTpo And cic.cPersCod = ci.cPersCod And cic.cCtaIFCod = ci.cCtaIFCod"
   sSQL = sSQL & "              LEFT JOIN ( SELECT mif.cPersCod, mif.cCtaIFCod, mif.cIFTpo, SUM(" & IIf(Mid(psOpeCod, 3, 1) = "2", "me.nMovMEImporte", "mc.nMovImporte") & ") nMonto"
   sSQL = sSQL & "                          FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro " & IIf(Mid(psOpeCod, 3, 1) = "2", "JOIN MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "")
   sSQL = sSQL & "                                     JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%'"
   sSQL = sSQL & "                          WHERE m.nMovFlag = 0 and LEFT(m.cmovnro,8) > '" & Format(pdFecha, gsFormatoMovFecha) & "' and oc.copecod = '" & psOpeCod & "' and oc.cOpeCtaOrden = '0'"
   sSQL = sSQL & "                          GROUP BY mif.cPersCod, mif.cCtaIFCod, mif.cIFTpo"
   sSQL = sSQL & "                        ) Canc ON canc.cPersCod = ci.cPersCod and canc.cCtaIFCod = ci.cCtaIFCod and canc.cIFTpo = ci.cIFTpo "
   sSQL = sSQL & "             LEFT JOIN ( SELECT cs.cPersCod, cs.cCtaIFCod, cs.cIFTpo, SUM(" & IIf(Mid(psOpeCod, 3, 1) = 1, "cs.nSaldo", "cs.nSaldoME") & ") nSaldoCnt"
   sSQL = sSQL & "                         FROM CtaIFSaldo cs JOIN OpeCta oc ON cs.cCtaContCod LIKE oc.cCtaContCod + '%'"
   sSQL = sSQL & "                         WHERE oc.cOpeCod = '" & psOpeCod & "' and dCtaIFSaldo = ( SELECT Max(dCtaIFSaldo) FROM CtaIFSaldo cs1"
   sSQL = sSQL & "                                          Where cs1.cCtaContCod = cs.cCtaContCod And cs1.cPersCod = cs.cPersCod And cs1.cIFTpo = cs.cIFTpo And cs1.cCtaIFCod = cs.cCtaIFCod"
   sSQL = sSQL & "                                                and dCtaIFSaldo <= '" & Format(pdFecha, gsFormatoFecha) & "' )"
   sSQL = sSQL & IIf(pbSoloCortoPlazo, " and oc.cCtaContCod LIKE '24%'", "")
   sSQL = sSQL & "                         GROUP BY cs.cPersCod, cs.cCtaIFCod, cs.cIFTpo"
   sSQL = sSQL & "                       ) Sdo ON Sdo.cPersCod = ci.cPersCod and Sdo.cIFTpo = ci.cIFTpo and Sdo.cCtaIFCod = ci.cCtaIFCod "
   sSQL = sSQL & "WHERE (ci.cCtaIFEstado in (" & gEstadoCtaIFActiva & "," & gEstadoCtaIFRegistrada & ") or not Canc.nMonto is NULL )"
   sSQL = sSQL & "  and  ci.cIFTpo+ci.cCtaIFCod LIKE '" & lsFiltroCta & "' " & IIf(pbSoloCortoPlazo, "", "and cia.cMonedaPago = '2'")
   sSQL = sSQL & "  and datediff(d,ci.dCtaIFAper,'" & Format(pdFecha, gsFormatoFecha) & "') >= 0 "
   sSQL = sSQL & "ORDER BY p.cPersNombre, ci.cCtaIFCod "

    Set CargaDatosAdeudoReclasifica = oCon.CargaRecordSet(sSQL)
Exit Function
CargaDatosGeneralesCtaIFErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CargaDatosADeudadosPYMES(psOpeCod As String, lsMoneda As String, ldFecha As Date, Optional lnIndiceVac As Double, Optional lsPersCod As String, Optional psFecha As Date) As ADODB.Recordset
Dim lsFiltro As String
Dim rs As New ADODB.Recordset
Dim cFecha As String
Dim cCtaIF As String
Dim cCtaIFAdeudados As String
Dim cCtaIFInteres As String
Dim bRegistrado As Boolean

On Error GoTo CargaDatosAdeudadosProyectadosErr
    
sSQL = "select dFecDataConsol "
sSQL = sSQL & " From CGAdeudCierreConsol where convert(varchar(8), dfeccierre, 112)='" & Format(ldFecha, "YYYYMMdd") & "'"

Set rs = oCon.CargaRecordSet(sSQL)
If rs.BOF Then
    cCtaIF = "CtaIF"
    cCtaIFAdeudados = "CtaIFAdeudados"
    cCtaIFInteres = "CtaIFInteres"
    bRegistrado = False
    
Else
    cFecha = Format(rs!dFecDataConsol, "YYYYMMdd")
    cCtaIF = "CGCtaIFConsol"
    cCtaIFAdeudados = "CGCtaIFAdeudadosConsol"
    cCtaIFInteres = "CGCtaIFInteresConsol"
    bRegistrado = True
End If


If lsPersCod <> "" Then
    lsFiltro = " AND Per.cPersCod = '" & Mid(lsPersCod, 4, 13) & "' "
End If
    
sSQL = "  SELECT DISTINCT "
sSQL = sSQL & "     convert(integer, datediff(d, ci.dCtaIFAper, dbo.AdeuGetFechaVencimiento(ci.cPersCod, ci.cIFTpo, ci.cCtaIFCod))/360) as nAnno, "
sSQL = sSQL & "     convert(integer, convert(integer, datediff(d, ci.dCtaIFAper, dbo.AdeuGetFechaVencimiento(ci.cPersCod, ci.cIFTpo, ci.cCtaIFCod))%360) / 30) as nMes, "
sSQL = sSQL & "     ci.cIFTpo, ci.cPersCod, ci.cCtaIFCod, P.cPersNombre, "
sSQL = sSQL & "     cii.nCtaIFIntPeriodo, cii.nCtaIFIntValor nTasaInteres, "
sSQL = sSQL & "     cia.nSaldoCap nSaldoCapCal, ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") nVac, cia.cMonedaPago, "
sSQL = sSQL & "     Round( cia.nSaldoCap * convert(decimal(15,5), CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' "
sSQL = sSQL & "     and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END),2) nSaldoCap, "
'sSQL = sSQL & "     + ISNULL( Canc.nMonto, 0) nSaldoCap , "
sSQL = sSQL & "     ISNULL(cia.cCodLinCred,'') cCodLinCred, ISNULL(l.cDescripcion,'') cDesLinCred "
sSQL = sSQL & " FROM " & cCtaIF & " ci LEFT JOIN " & cCtaIFAdeudados & " cia ON cia.cIFTpo = ci.cIFTpo and cia.cPersCod = ci.cPersCod "
sSQL = sSQL & "     and cia.cCtaIFCod = ci.cCtaIFCod "
sSQL = sSQL & " INNER JOIN ColocLineaCredito l ON l.cLineaCred = cia.cCodLinCred JOIN Persona P ON ci.cPersCod = p.cPersCod "
sSQL = sSQL & " LEFT JOIN IndiceVac iv ON iv.dIndiceVac = ISNULL(cia.dCuotaUltPago, ci.dCtaIFAper) "
sSQL = sSQL & " JOIN " & cCtaIFInteres & " cii ON cii.cIFTpo = ci.cIFTpo and cii.cPersCod = ci.cPersCod and cii.cCtaIFCod = ci.cCtaIFCod "
sSQL = sSQL & "     and cii.dCtaIFIntRegistro = (   SELECT Max(dCtaIFIntRegistro) "
sSQL = sSQL & "                                     FROM " & cCtaIFInteres & " cii1 "
sSQL = sSQL & "                                     Where cii1.cIFTpo = cii.cIFTpo And cii1.cPersCod = cii.cPersCod And cii1.cCtaIFCod = cii.cCtaIFCod "

    If bRegistrado = True Then
        sSQL = sSQL & " And Convert(varchar(8), cii1.dfecha,112)='" & cFecha & "' "
    End If

sSQL = sSQL & "                                 )"

sSQL = sSQL & " WHERE (ci.cCtaIFEstado in (1,0) )" 'or not Canc.nMonto is NULL ) "
sSQL = sSQL & "     and  ci.cIFTpo+ci.cCtaIFCod LIKE '__05" & lsMoneda & "%'"
sSQL = sSQL & "     and datediff(d,ci.dCtaIFAper,'" & Format(ldFecha, "MM/dd/YYYY") & "') >= 0 "
sSQL = sSQL & lsFiltro

If bRegistrado = True Then
    sSQL = sSQL & " And Convert(varchar(8), ci.dfecha,112)='" & cFecha & "' "
    sSQL = sSQL & " And Convert(varchar(8), cia.dfecha,112)='" & cFecha & "' "
    sSQL = sSQL & " And Convert(varchar(8), cii.dfecha,112)='" & cFecha & "' "
End If

sSQL = sSQL & " ORDER BY ci.cPersCod, ISNULL(cia.cCodLinCred,''), convert(integer, datediff(d, ci.dCtaIFAper, dbo.AdeuGetFechaVencimiento(ci.cPersCod, ci.cIFTpo, ci.cCtaIFCod))/360) "
     
Set CargaDatosADeudadosPYMES = oCon.CargaRecordSet(sSQL)
Exit Function
CargaDatosAdeudadosProyectadosErr:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Public Function GetDatosAdeudadoPago(pdFecha As Date, pnMoneda As Moneda, Optional psBuscar As String = "", Optional bCancelacion As Boolean = False, Optional bRestringirFecha As Boolean = False, Optional pbControlFechaRep As Boolean = False) As ADODB.Recordset

    On Error GoTo GetDatosAdeudadoPagoErr
    Dim oCon As New COMConecta.DCOMConecta
    
    sSQL = "Select CIF.cPersCod, P.cPersNombre, CIF.cIFTpo, CIF.cCtaIFCod, CIF.cCtaIFDesc, "
    
    If bCancelacion = True Then
        sSQL = sSQL & " CICUltimo.dVencimiento, CalRef.nNroCuota, '2' as cTpoCuota, "
        sSQL = sSQL & " CICUltimo.cEstado as cEstadoCuota, "
    ElseIf bCancelacion = False Then
        sSQL = sSQL & " CIC.dVencimiento, CIC.nNroCuota, CIC.cTpoCuota, "
        sSQL = sSQL & " CIC.cEstado as cEstadoCuota, "
    End If
    
    sSQL = sSQL & " CII.nCtaIFIntPeriodo, CII.nCtaIFIntValor, CIA.cMonedaPago, CIA.nMontoPrestado, "
    sSQL = sSQL & " ISNULL(CIE.nSumaInteresMovido, 0) as nInteresProvisionado, "
    sSQL = sSQL & " ISNULL(CIE.nSumaInteresMovidoReal, 0) as nInteresProvisionadoReal, "
    sSQL = sSQL & " CIA.dCuotaUltModSaldos, CIA.dCuotaUltPago, "
    
    If bCancelacion = True Then
        sSQL = sSQL & " CIA.nSaldoCap, CIA.nSaldoCapLP, CIC.nCapital, IsNull(CIC.nInteres,0) + IsNull(CIC.nInteresGracia,0) As nInteres, CIC.nComision, CIC.nTotalCuota, "
    Else
        sSQL = sSQL & " CIA.nSaldoCap, CIA.nSaldoCapLP, CIC.nCapital + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,1) nCapital, IsNull(CIC.nInteres,0) +  dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,2) + IsNull(CIC.nInteresGracia,0) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,3) As nInteres, CIC.nComision + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,4) nComision, CIC.nTotalCuota + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,1) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,2) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,3) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,4) nTotalCuota, "
    End If
    
    sSQL = sSQL & " ISNULL(CIA.cCodLinCred, '') as cCodLinCred, ISNULL(CLC.cDescripcion, '') as cDesLinCred , "
    sSQL = sSQL & " iVacAper=(select nIndiceVac from indiceVac where dIndiceVAC=CIF.dctaifaper), (Select max(dVencimiento) from CtaIfCalendario Calen where Calen.cPersCod = CIF.cPersCod And Calen.cIFTpo = CIF.cIFTpo And Calen.cCtaIfcod = CIF.cCtaIFCod And bVigente = 1) dVencimientoFinal, (Select dCtaIFAper from CtaIf where CtaIf.cPersCod = CIF.cPersCod And CtaIf.cIFTpo = CIF.cIFTpo And CtaIf.cCtaIfcod = CIF.cCtaIFCod) dCtaIFAper "
    
    sSQL = sSQL & " From CtaIF CIF "
    sSQL = sSQL & " Inner Join CtaIFAdeudados CIA "
    sSQL = sSQL & "         On CIF.cPersCod=CIA.cPersCod And CIF.cIFTpo=CIA.cIFTpo And CIF.cCtaIFCod=CIA.cCtaIFCod "
    
    If pbControlFechaRep Then
        sSQL = sSQL & " Inner Join CtaIF CtaIF "
        sSQL = sSQL & "         On CIF.cPersCod = CtaIF.cPersCod And CIF.cIFTpo = CtaIF.cIFTpo And CIF.cCtaIFCod = CtaIF.cCtaIFCod And CtaIF.dCtaIfAper <= '" & Format(pdFecha, gsFormatoFecha) & "' "
    End If
    
    'Proxima Cuota Pendiente
    sSQL = sSQL & " Inner Join (    SELECT cPersCod, cIFTpo, cCtaIFCod, MIN(nNroCuota) nNroCuota "
    sSQL = sSQL & "                 From CtaIFCalendario "
    sSQL = sSQL & "                 WHERE bVigente=1 And (cEstado = '0' or cEstado='2') and cTpoCuota <> '6' "
    sSQL = sSQL & "                     And Substring(cCtaIFCod, 3,1)='" & Trim(Str(pnMoneda)) & "' " '& Replace(psBuscar, "CIF.", "")
    sSQL = sSQL & "                 GROUP BY cPersCod, cIFTpo, cCtaIFCod "
    sSQL = sSQL & "             ) CalRef ON CIF.cPersCod=CalRef.cPersCod And CIF.cIFTpo=CalRef.cIFTpo And CIF.cCtaIFCod=CalRef.cCtaIFCod "
    
    If bCancelacion = True Then
        
        'Sacamos todo el capital, intereses y comisiones pendientes Pero de lo no Concesional,
        'Porque lo concesional se PERDONA!!!
        
        sSQL = sSQL & " Inner Join ( "
        sSQL = sSQL & " select  CIK.cPersCod, CIK.cIFTpo, CIK.cCtaIFCod, Sum(IsNull(CIK.nCapital,0)) as nCapital, Sum(IsNull(CIK.nInteres,0)) AS nInteres, Sum(IsNull(CIK.nInteresGracia,0)) as nInteresGracia, "
        sSQL = sSQL & " sum(IsNull(CIK.nComision,0)) As nComision, Sum(IsNull(CIK.nCapital,0) + IsNull(CIK.nInteres,0) + IsNull(CIK.nInteresGracia,0) + IsNull(CIK.nComision,0)) as nTotalCuota "
        sSQL = sSQL & " from ctaifcalendario CIK "
        sSQL = sSQL & " Where CIK.bVigente=1 And CIK.cEstado in('0', '2')  " 'And CIK.cTpoCuota='2'
        sSQL = sSQL & " Group By CIK.cPersCod, CIK.cIFTpo, CIK.cCtaIFCod "
        sSQL = sSQL & " ) CIC On CIF.cPersCod=CIC.cPersCod And CIF.cIFTpo=CIC.cIFTpo And CIF.cCtaIFCod=CIC.cCtaIFCod "
    
        'Sacamos el estado de la ultima cuota pendiente
        sSQL = sSQL & " Inner Join ( "
        sSQL = sSQL & " Select CIK.cPersCod, CIK.cIFTpo, CIK.cCtaIFCod, CIK.nNroCuota, CIK.cEstado, CIK.dVencimiento "
        sSQL = sSQL & " From CtaIFCalendario CIK "
        sSQL = sSQL & " Where CIK.bVigente=1 And CIK.cEstado in('0', '2') "
        sSQL = sSQL & " And CIK.cTpoCuota='2' "
        sSQL = sSQL & " ) CICUltimo On CIF.cPersCod=CICUltimo.cPersCod And CIF.cIFTpo=CICUltimo.cIFTpo And CIF.cCtaIFCod=CICUltimo.cCtaIFCod "
        sSQL = sSQL & " And CalRef.nNroCuota=CICUltimo.nNroCuota "
    
    ElseIf bCancelacion = False Then
    
        'Sacamos el capital, intereses y comisiones de la prox cuota pendiente
        sSQL = sSQL & " Inner Join CtaIFCalendario CIC "
        sSQL = sSQL & "         ON CIF.cPersCod=CIC.cPersCod And CIF.cIFTpo=CIC.cIFTpo And CIF.cCtaIFCod=CIC.cCtaIFCod "
        sSQL = sSQL & "                 AND CIC.nNroCuota=CalRef.nNroCuota "
    
    End If
    
    sSQL = sSQL & " Inner Join CtaIFInteres CII "
    sSQL = sSQL & "         On CIF.cPersCod=CII.cPersCod And CIF.cIFTpo=CII.cIFTpo And CIF.cCtaIFCod=CII.cCtaIFCod "
    sSQL = sSQL & " Inner JOIN Persona P "
    sSQL = sSQL & "         ON CIF.cPersCod = P.cPersCod "
    
    'Ultimo Interes provisionado
    sSQL = sSQL & " Left Join ( "
    sSQL = sSQL & "       Select CIE.cPersCod, CIE.cIFTpo, CIE.cCtaIFCod, " & IIf(bCancelacion = True, "", " CIE.nNroCuota,")
    sSQL = sSQL & "           SUM(CIE.nInteres1Movido) as nSumaInteresMovido, SUM(CIED.nInteres1MovidoReal) as nSumaInteresMovidoReal "
    sSQL = sSQL & "       From CtaIFEstadistica CIE "
    sSQL = sSQL & "       Inner Join CtaIFEstadisticaDet CIED "
    sSQL = sSQL & "                       On CIE.nMovNro = CIED.nMovNro "
    sSQL = sSQL & "       Inner Join Mov m on   cie.nmovnro=m.nmovnro and m.nmovflag<>1 " & IIf(pbControlFechaRep, " And cMovNro < '" & Format(pdFecha + 1, gsFormatoMovFecha) & "' ", "")
    sSQL = sSQL & "       Left Join ( Select cPersCod,cIFTpo,cCtaIFCod,max(nNroCuota) nNroCuota"
    sSQL = sSQL & "                from CtaIFEstadistica a"
    sSQL = sSQL & "                Inner Join Mov b on a.nmovnro=b.nmovnro and b.nmovflag<>1 " & IIf(pbControlFechaRep, " And cMovNro < '" & Format(pdFecha + 1, gsFormatoMovFecha) & "' ", "")
    sSQL = sSQL & "                Where nTipoEstadistica = 3"
    sSQL = sSQL & "                group by cPersCod,cIFTpo,cCtaIFCod"
    sSQL = sSQL & "           ) X   on x.cperscod=CIE.cperscod and x.cCtaIFCod=CIE.cCtaIFCod "
    sSQL = sSQL & "       Where CIE.nTipoEstadistica = 4 and IsNull(x.nNroCuota,0)+1=CIE.nNroCuota"
    sSQL = sSQL & "       Group By CIE.cPersCod, CIE.cIFTpo, CIE.cCtaIFCod" & IIf(bCancelacion = True, "", ",CIE.nNroCuota")
    sSQL = sSQL & " ) CIE On CIF.cPersCod=CIE.cPersCod And CIF.cCtaIFCod=CIE.cCtaIFCod And CIF.cIFTpo=CIE.cIFTpo "
    
    If bCancelacion = False Then
        sSQL = sSQL & " and CIE.nNrocuota=CIC.nNroCuota "
    End If
    'Linea de Credito
    
    sSQL = sSQL & " Left Join ColocLineaCredito CLC "
    sSQL = sSQL & "         ON CIA.cCodLinCred=CLC.cLineaCred"

    sSQL = sSQL & " Where   CIF.cCtaIFCod Like '__" & pnMoneda & "%' and CIF.cCtaIFEstado = '" & gEstadoCtaIFActiva & "' "
    
    'Solo se Pagan las que no estan pagadas
    
    If bCancelacion = True Then
        
        If bRestringirFecha = True Then
            sSQL = sSQL & " And CICUltimo.dVencimiento >='" & Format(pdFecha, "MM/dd/YYYY") & "' And CICUltimo.dVencimiento < DateAdd(dd,1,'" & Format(pdFecha, "MM/dd/YYYY") & "') "
        End If
        
    ElseIf bCancelacion = False Then
        
        sSQL = sSQL & " And CIC.bVigente=1 And (CIC.cEstado = '0' or CIC.cEstado='2') and CIC.cTpoCuota <> '6' "
        
        If bRestringirFecha = True Then
            sSQL = sSQL & " And CIC.dVencimiento >='" & Format(pdFecha, "MM/dd/YYYY") & "' And CIC.dVencimiento < DateAdd(dd,1,'" & Format(pdFecha, "MM/dd/YYYY") & "') "
        End If
        
    End If
    
    sSQL = sSQL & " And CII.nVigente=1 "
    sSQL = sSQL & " " & psBuscar & " ORDER BY P.cPersNombre, CIF.cCtaIFCod "

    oCon.AbreConexion
    Set GetDatosAdeudadoPago = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Exit Function
    
GetDatosAdeudadoPagoErr:
    MsgBox Err.Description, vbCritical
End Function

Public Function CalculaMontoCuota(psPersCod As String, psIFTpo As String, psCtaIfCod As String, pscTpoCuota As String, pnNroCuota As Integer, Optional bCancelacion As Boolean = False) As Currency
Dim rs As New ADODB.Recordset

On Error GoTo CalculaMontoCuotaErr

    CalculaMontoCuota = 0
    
    If bCancelacion = True Then
        sSQL = "SELECT  SUM(nCapital) as nCapital "
    Else
        sSQL = "Select nCapital "
    End If
    
    sSQL = sSQL & " FROM CtaIFCalendario "
    sSQL = sSQL & " WHERE   cPersCod = '" & psPersCod & "' and cIFTpo = '" & psIFTpo & "' and cCtaIFCod = '" & psCtaIfCod & "' "
    sSQL = sSQL & "         And  cTpoCuota='" & pscTpoCuota & "' "
    
    If bCancelacion = True Then
        sSQL = sSQL & " And nNroCuota>=" & pnNroCuota
    Else
        sSQL = sSQL & " And nNroCuota=" & pnNroCuota
    End If
    
    Set rs = oCon.CargaRecordSet(sSQL)
    If Not rs.EOF Then
        CalculaMontoCuota = rs!nCapital
    End If
    RSClose rs
    
Exit Function

CalculaMontoCuotaErr:
    Err.Raise vbObjectError + 100, "CalculaSaldoAdeudadoLargoPlazo", Err.Description & " Adeudados "
End Function

Public Function GetSubCtaLineaCredito(ByVal pscLineaCred As String) As String

Dim rs As New ADODB.Recordset
Set oCon = New COMConecta.DCOMConecta
sSQL = "Select isnull(cCtaIFSubCta,'') cCtaIFSubCta "
sSQL = sSQL & " From ColocLineaCredito "
sSQL = sSQL & " Where cLineaCred='" & pscLineaCred & "'"

oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSQL)
If rs.BOF Then
    GetSubCtaLineaCredito = ""
Else
    GetSubCtaLineaCredito = rs!cCtaIFSubCta
End If

rs.Close
Set rs = Nothing
oCon.CierraConexion
End Function


Public Function InteresReal1(ByVal Ti As Double, ByVal nDiasTrans As Integer, Optional nPeriodo As Integer = 30) As Double
    Dim lnFactor As Double
    
    lnFactor = nDiasTrans / nPeriodo
    InteresReal1 = ((1 + (Ti / 100)) ^ lnFactor) - 1
     
End Function

Public Function GetOpeCta(ByVal psOpeCod As String, ByVal psDH As String, Optional psOpeCtaOrden As String = "0", _
                           Optional psPersCod As String = "", Optional psIFTpo As String = "", _
                           Optional pscCtaIFCod As String = "") As String

Dim rs As New ADODB.Recordset
Dim lsCuenta As String
Dim lsSubCuenta As String

sSQL = "Select Distinct cCtaContCod "
sSQL = sSQL & " from OpeCta where copecod='" & psOpeCod & "' and copectaorden='" & psOpeCtaOrden & "' "
sSQL = sSQL & " And cOpeCtaDH='" & psDH & "' " 'and cOpeCtaOpc='1'"
Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSQL)
If rs.BOF Then
Else
    If Len(Trim(psPersCod)) = 0 Then
        If rs.RecordCount = 1 Then
            lsCuenta = rs!cCtaContCod
        End If
    Else
        Do While Not rs.EOF
            If lsCuenta = "" Then
                lsCuenta = GetSubCtaIFFiltro(psPersCod, psIFTpo, pscCtaIFCod, rs!cCtaContCod)
                If lsCuenta <> "" Then
                    lsCuenta = rs!cCtaContCod & lsCuenta
                End If
            End If
            rs.MoveNext
        Loop
    End If
End If
Set rs = Nothing
oCon.CierraConexion
GetOpeCta = lsCuenta

End Function

Public Function GetSubCtaIFFiltro(cPerscod As String, cIFTpo As String, cCtaIFCod As String, cCtaContCod As String) As String
 
Dim rs As New ADODB.Recordset

sSQL = "Select cCtaIFSubCta "
sSQL = sSQL & " From CtaIFFiltro "
sSQL = sSQL & " Where cPersCod='" & cPerscod & "' And cIFTpo ='" & cIFTpo & "' "
sSQL = sSQL & " And cCtaIFCod='" & cCtaIFCod & "'"
 
If Len(Trim(cCtaContCod)) > 0 Then
    sSQL = sSQL & " And cCtaContCod='" & cCtaContCod & "'"
End If
Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSQL)
If rs.BOF Then
Else
    If rs.RecordCount = 1 Then
        If Len(Trim(cCtaContCod)) > 0 Then
            GetSubCtaIFFiltro = rs!cCtaIFSubCta
        Else
            GetSubCtaIFFiltro = rs!cCtaContCod
        End If
    End If
End If
Set rs = Nothing
oCon.CierraConexion
End Function


Public Function GetConvierteCuentaVac(ByVal cCtaContCod As String) As String
Dim cTempo As String
Dim gMonedaVAC As String
    ' ==================== CONVERTIR CUENTA EN SOLES A CUENTA VAC ====================
    gMonedaVAC = "3"
    
    If Mid(cCtaContCod, 3, 1) = Moneda.gMonedaNacional Then
        cTempo = Left(cCtaContCod, 2) & gMonedaVAC & Mid(cCtaContCod, 4, Len(cCtaContCod))
    End If
    GetConvierteCuentaVac = cTempo
End Function

Public Function GetbExtranjeroInteres(ByVal pscPersCod As String, pscIFTpo As String, pscLineaCred As String, pscMoneda As String) As Boolean

Dim rs As New ADODB.Recordset
sSQL = "Select bInteresExtranjero "
sSQL = sSQL & " From CGCtaIFInstituciones "
sSQL = sSQL & " Where cPersCod='" & pscPersCod & "' And cIFTpo='" & pscIFTpo & "' and cCodLinCred='" & pscLineaCred & "' And cMoneda='" & pscMoneda & "'"

Set rs = oCon.CargaRecordSet(sSQL)
If rs.BOF Then
    GetbExtranjeroInteres = False
Else
    If rs!bInteresExtranjero = 1 Then
        GetbExtranjeroInteres = True
    Else
        GetbExtranjeroInteres = False
    End If
End If

rs.Close
Set rs = Nothing
End Function


Public Function CalculaSaldoAdeudadoLargoPlazo(pdFecha As Date, psPersCod As String, psIFTpo As String, psCtaIfCod As String, Optional psTpoCuota As CGTipoCuotCalIF = gCGTipoCuotCalIFCuota, Optional pCon As Object) As Currency
On Error GoTo CalculaSaldoAdeudadoLargoPlazoErr
Dim sql As String
Dim oCon As New COMConecta.DCOMConecta
Dim rs As New ADODB.Recordset

CalculaSaldoAdeudadoLargoPlazo = 0

    sql = "SELECT ISNULL(SUM(nCapital),0) nCapital FROM CtaIFCalendario " _
        & "WHERE cPersCod = '" & psPersCod & "' and cIFTpo = '" & psIFTpo & "' and cCtaIFCod = '" & psCtaIfCod & "' and " _
        & "      datediff(d,'" & Format(pdFecha, gsFormatoFecha) & "',dVencimiento) > 360 and cEstado = '" & gTpoEstCuotaAdeudPend & "' "
    If Not pCon Is Nothing Then
        oCon.ConexionActiva = pCon
        lbConexionPropia = False
    End If
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    If Not rs.EOF Then
        CalculaSaldoAdeudadoLargoPlazo = rs!nCapital
    End If
    RSClose rs
Exit Function
CalculaSaldoAdeudadoLargoPlazoErr:
    Err.Raise vbObjectError + 100, "CalculaSaldoAdeudadoLargoPlazo", Err.Description & " Adeudados "
End Function

Private Sub Class_Initialize()
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
End Sub

Private Sub Class_Terminate()
If lbConexionPropia = True Then
    oCon.CierraConexion
    Set oCon = Nothing
End If
End Sub

Public Function GetAdeudadosSaldoCap(psPersCod As String, Optional psCtaIfCod As String = "", Optional psIFTpo As String = "", Optional pCon As ADODB.Connection, Optional pdFechaAl As Date = 0, Optional psFiltro As String = "", Optional psMoneda As String = "", Optional psTpoCuota As Integer = gCGTipoCuotCalIFCuota, Optional psFiltroNuevo As String = "") As Currency
Dim prs As ADODB.Recordset
On Error GoTo GetAdeudadosSaldoCapErr

If Not IsNull(pCon) And Not pCon Is Nothing Then
    lbConexionPropia = False
    oCon.ConexionActiva = pCon
Else
    lbConexionPropia = True
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
End If
If pdFechaAl = 0 Then
    
    If psFiltroNuevo = "" Then
        psFiltroNuevo = " and (cTpoCuota = '" & psTpoCuota & "') "
    Else
        psFiltroNuevo = psFiltroNuevo
    End If
    
    sSQL = "SELECT ISNULL(SUM(nCapital),0) nSaldoCap " _
     & "FROM  CtaIFCalendario " _
     & "WHERE (cPersCod = '" & psPersCod & "' and cCtaIFCod = '" & psCtaIfCod & "' and cIFTpo = '" & psIFTpo & "' and " _
     & "      cEstado = '" & gTpoEstCuotaAdeudPend & "') " & psFiltroNuevo
     
     
Else
    sSQL = "SELECT cPersCod, " & IIf(psCtaIfCod = "", "", " cIFTpo, cCtaIFCod,") & " ISNULL(SUM(nSaldoCap),0) nSaldoCap " _
     & "FROM (SELECT ci.cPersCod, " & IIf(psCtaIfCod = "", "", " ci.cIFTpo, ci.cCtaIFCod,") & " cia.nSaldoCap + " _
     & "       ISNULL( (SELECT SUM(" & IIf(psMoneda = "2", "me.nMovImporte", "mc.nMovImporte") & ") " _
     & "                FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro " & IIf(psMoneda = "2", " JOIN MovME me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
     & "                       JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem " _
     & "                       JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' " _
     & "                WHERE m.nMovFlag = " & gMovFlagVigente & " and LEFT(m.cmovnro,8) > '" & Format(pdFechaAl, gsFormatoMovFecha) & "' and oc.cOpeCod IN ('" & OpeCGAdeudRepGeneralMN & "', '" & OpeCGAdeudRepGeneralME & "') and oc.cOpeCtaOrden = '0' " _
     & "                  and not m.cOpeCod IN ('401804', '402804') " _
     & "                  and mif.cPersCod = ci.cPersCod and mif.cCtaIFCod = ci.cCtaIFCod and mif.cIFTpo = ci.cIFTpo " _
     & "       ), 0) as nSaldoCap " _
     & "FROM ctaif ci join ctaifadeudados cia on cia.cperscod = ci.cperscod " _
     & "                              and cia.cIFTpo = ci.cIFTpo and cia.cCtaIFCod = ci.cCtaIFCod " _
     & "    WHERE cia.cPersCod = '" & psPersCod & "' and cia.cIFTpo + cia.cCtaIFCod like '" & psFiltro & "%' and cCtaIFEstado = 1 ) a " _
     & "GROUP BY cPersCod " & IIf(psCtaIfCod = "", "", ", cIFTpo, cCtaIFCod ")
     
End If
Set prs = oCon.CargaRecordSet(sSQL)
If Not prs.EOF Then
    GetAdeudadosSaldoCap = prs!nSaldoCap
End If


Exit Function
GetAdeudadosSaldoCapErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
