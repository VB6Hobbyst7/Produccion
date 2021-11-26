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
Dim oCon      As DConecta
Dim sSql      As String
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
   sSql = "SELECT ci.cIFTpo, ci.cPersCod, ci.cCtaIFCod, P.cPersNombre, ci.cCtaIFDesc, ci.dCtaIFAper, dCtaIFVenc, cia.nMontoPrestado, ci.nCtaIFPlazo, cia.nCtaIFCuotas, cia.nPeriodoGracia, cii.nCtaIFIntPeriodo, cii.nCtaIFIntValor nTasaInteres, cia.nSaldoCap nSaldoCapCal, ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") nVac, cia.cMonedaPago, " _
        & "       Round( cia.nSaldoCap * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END,2) nSaldoCap, " _
        & "             ISNULL(cia.cCodLinCred,'') cCodLinCred, ISNULL(l.cDescripcion,'') cDesLinCred, cia.nCuotaPagoCap " _
        & "    FROM CtaIF ci LEFT JOIN CtaIfAdeudados cia ON cia.cIFTpo = ci.cIFTpo and cia.cPersCod = ci.cPersCod and cia.cCtaIFCod = ci.cCtaIFCod LEFT JOIN ColocLineaCredito l ON l.cLineaCred = cia.cCodLinCred JOIN Persona P ON ci.cPersCod = p.cPersCod " _
        & "         LEFT JOIN IndiceVac iv ON iv.dIndiceVac = ISNULL(cia.dCuotaUltPago, ci.dCtaIFAper) " _
        & "          JOIN CtaIFInteres cii ON cii.cIFTpo = ci.cIFTpo and cii.cPersCod = ci.cPersCod and cii.cCtaIFCod = ci.cCtaIFCod " _
        & "    WHERE ci.cCtaIFEstado in (" & gEstadoCtaIFActiva & "," & gEstadoCtaIFRegistrada & ") and  ci.cIFTpo+ci.cCtaIFCod LIKE '" & psFiltroCta & "' " & lsFiltro _
        & "    ORDER BY ci.cIFtpo, ci.cPersCod, ISNULL(cia.cCodLinCred,''), ci.cCtaIFDesc  "

  
    Set CargaDatosAdeudadosProyectados = oCon.CargaRecordSet(sSql)
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
    
    sSql = " select Per.cPersNombre, " & _
          " ISNULL(cia.cCodLinCred,'') cCodLinCred, ISNULL(l.cDescripcion,'') cDesLinCred," & _
          " CIF.cPersCod, CIF.cIFTpo, CIF.cCtaIFCod, CIF.cCtaIFDesc,CIF.dCtaIFAper," & _
          " case Substring(CIF.cCtaIFCod, 3,1) when '2' then 'N.DOLARES' when '1' then 'N.SOLES' END as cMoneda," & _
          " CIA.nMontoPrestado, CIA.nSaldoCap, dbo.AdeuGetFechaVencimiento(CIF.cPersCod, CIF.cIFTpo, CIF.cCtaIFCod) as cVencimiento, case cFinalidad when '1' then 'INTER.FINA'  else 'ACTIVO FIJO' END Finalidad," & _
          " datediff(m, CIF.dCtaIFAper, dbo.AdeuGetFechaVencimiento(CIF.cPersCod, CIF.cIFTpo, CIF.cCtaIFCod)) As nMeses, " & _
          " CIC2.nCantidadCuotas, cic3.nProxCuota, cii.nCtaIFIntValor nTasaInteres"
          
'    ssql=ssql & " , "
          
'    sSql = sSql & " case when CIA.nMontoPrestado - ISNULL(Sald1.nSaldo,0) >0 then CIA.nMontoPrestado - ISNULL(Sald1.nSaldo,0) else 0 end as nSaldo1, "
'    sSql = sSql & " case when CIA.nMontoPrestado - ISNULL(Sald2.nSaldo,0) >0 then CIA.nMontoPrestado - ISNULL(Sald2.nSaldo,0) else 0 end as nSaldo2, "
'    sSql = sSql & " case when CIA.nMontoPrestado - ISNULL(Sald3.nSaldo,0) >0 then CIA.nMontoPrestado - ISNULL(Sald3.nSaldo,0) else 0 end as nSaldo3 "
          
    sSql = sSql & " From ctaif CIF" & _
          " Inner Join CtaIfAdeudados CIA" & _
          " On CIF.cPersCod=CIA.cPersCod and CIF.cIFTpo= CIA.cIFTpo And CIF.cCtaIFCod= CIA.cCtaIFCod" & _
          " Inner Join Persona Per ON Per.cPersCod = CIF.cPersCod" & _
          " LEFT JOIN ColocLineaCredito l ON l.cLineaCred = cia.cCodLinCred"
    
' sSql = sSql & " Left Join ( "
'sSql = sSql & " SELECT  MO.cIFTpo, MO.cPersCod, MO.cCtaIFCod, "
'If Mid(psOpeCod, 3, 1) = "1" Then
'    sSql = sSql & " ABS(SUM(CASE WHEN oc.cOpeCtaOrden = 0 THEN  ISNULL(Mc.nMovImporte,0)  ELSE 0 END)) AS nSaldo "
'ElseIf Mid(psOpeCod, 3, 1) = "2" Then
'    sSql = sSql & " ABS(SUM(CASE WHEN oc.cOpeCtaOrden = 0 THEN  ISNULL(ME.nMovMEImporte,0)  ELSE 0 END)) AS nSaldo "
'End If
'sSql = sSql & " FROM    MOV M JOIN MOVCTA MC   ON MC.nMOVNRO= M.nMOVNRO "
'sSql = sSql & " LEFT JOIN MOVME ME    ON ME.nMOVNRO=MC.nMOVNRO AND MC.nMOVITEM=ME.nMOVITEM "
'sSql = sSql & " JOIN MOVOBJIF MO ON MO.nMOVNRO=MC.nMOVNRO AND MO.nMOVITEM=MC.nMOVITEM "
'sSql = sSql & " LEFT JOIN CtaIFAdeudados cia ON cia.cIFTpo = mo.cIFTpo and cia.cPersCod = mo.cPersCod and cia.cCtaIFCod = mo.cCtaIFCod "
'sSql = sSql & " LEFT JOIN CtaIF  cii ON cii.cIFTpo = mo.cIFTpo and cii.cPersCod = mo.cPersCod and cii.cCtaIFCod = mo.cCtaIFCod "
'sSql = sSql & " JOIN OpeCta OC   ON MC.cCtaContCod LIKE OC.cCtaContCod + '%' "
'sSql = sSql & " LEFT JOIN IndiceVac i ON i.dIndiceVac = Convert(datetime,LEFT(m.cMovNro,8) ) "
'sSql = sSql & " WHERE M.copecod='40" & Mid(psOpeCod, 3, 1) & "806' AND M.nMOVFLAG NOT IN(1, 2, 3, 5) and m.nMovEstado = 10 "
'sSql = sSql & " AND SUBSTRING(M.CMOVNRO,1,8)> convert(varchar(8), cii.dCtaIFAper, 112) "
'sSql = sSql & " AND SUBSTRING(M.CMOVNRO,1,8)<='" & Trim(psFecha1) & "' "
'sSql = sSql & " and oc.cOpeCod = '46" & Mid(psOpeCod, 3, 1) & "092' "
'sSql = sSql & " And M.nMovNro Not IN(682714, 683330, 692474, 718469, 718476) "
'sSql = sSql & " GROUP BY MO.cIFTpo, MO.cPersCod, MO.cCtaIFCod ) Sald1"
'sSql = sSql & " ON Sald1.cPersCod = cif.cPersCod and Sald1.cCtaIFCod = ciF.cCtaIFCod and Sald1.cIFTpo = ciF.cIFTpo "
'
'sSql = sSql & " Left Join ( "
'sSql = sSql & " SELECT  MO.cIFTpo, MO.cPersCod, MO.cCtaIFCod, "
'If Mid(psOpeCod, 3, 1) = "1" Then
'    sSql = sSql & " ABS(SUM(CASE WHEN oc.cOpeCtaOrden = 0 THEN  ISNULL(Mc.nMovImporte,0)  ELSE 0 END)) AS nSaldo "
'ElseIf Mid(psOpeCod, 3, 1) = "2" Then
'    sSql = sSql & " ABS(SUM(CASE WHEN oc.cOpeCtaOrden = 0 THEN  ISNULL(ME.nMovMEImporte,0)  ELSE 0 END)) AS nSaldo "
'End If
'sSql = sSql & " FROM    MOV M JOIN MOVCTA MC   ON MC.nMOVNRO= M.nMOVNRO "
'sSql = sSql & " LEFT JOIN MOVME ME    ON ME.nMOVNRO=MC.nMOVNRO AND MC.nMOVITEM=ME.nMOVITEM "
'sSql = sSql & " JOIN MOVOBJIF MO ON MO.nMOVNRO=MC.nMOVNRO AND MO.nMOVITEM=MC.nMOVITEM "
'sSql = sSql & " LEFT JOIN CtaIFAdeudados cia ON cia.cIFTpo = mo.cIFTpo and cia.cPersCod = mo.cPersCod and cia.cCtaIFCod = mo.cCtaIFCod "
'sSql = sSql & " LEFT JOIN CtaIF  cii ON cii.cIFTpo = mo.cIFTpo and cii.cPersCod = mo.cPersCod and cii.cCtaIFCod = mo.cCtaIFCod "
'sSql = sSql & " JOIN OpeCta OC   ON MC.cCtaContCod LIKE OC.cCtaContCod + '%' "
'sSql = sSql & " LEFT JOIN IndiceVac i ON i.dIndiceVac = Convert(datetime,LEFT(m.cMovNro,8) ) "
'sSql = sSql & " WHERE M.copecod='40" & Mid(psOpeCod, 3, 1) & "806' AND M.nMOVFLAG NOT IN(1, 2, 3, 5) and m.nMovEstado = 10 "
'sSql = sSql & " AND SUBSTRING(M.CMOVNRO,1,8)> convert(varchar(8), cii.dCtaIFAper, 112) "
'sSql = sSql & " AND SUBSTRING(M.CMOVNRO,1,8)<='" & Trim(psfecha2) & "' "
'sSql = sSql & " and oc.cOpeCod = '46" & Mid(psOpeCod, 3, 1) & "092' "
'sSql = sSql & " And M.nMovNro Not IN(682714, 683330, 692474, 718469, 718476) "
'sSql = sSql & " GROUP BY MO.cIFTpo, MO.cPersCod, MO.cCtaIFCod ) Sald2"
'sSql = sSql & " ON Sald2.cPersCod = cif.cPersCod and Sald2.cCtaIFCod = ciF.cCtaIFCod and Sald2.cIFTpo = ciF.cIFTpo "
'
'sSql = sSql & " Left Join ( "
'sSql = sSql & " SELECT  MO.cIFTpo, MO.cPersCod, MO.cCtaIFCod, "
'If Mid(psOpeCod, 3, 1) = "1" Then
'    sSql = sSql & " ABS(SUM(CASE WHEN oc.cOpeCtaOrden = 0 THEN  ISNULL(Mc.nMovImporte,0)  ELSE 0 END)) AS nSaldo "
'ElseIf Mid(psOpeCod, 3, 1) = "2" Then
'    sSql = sSql & " ABS(SUM(CASE WHEN oc.cOpeCtaOrden = 0 THEN  ISNULL(ME.nMovMEImporte,0)  ELSE 0 END)) AS nSaldo "
'End If
'sSql = sSql & " FROM    MOV M JOIN MOVCTA MC   ON MC.nMOVNRO= M.nMOVNRO "
'sSql = sSql & " LEFT JOIN MOVME ME    ON ME.nMOVNRO=MC.nMOVNRO AND MC.nMOVITEM=ME.nMOVITEM "
'sSql = sSql & " JOIN MOVOBJIF MO ON MO.nMOVNRO=MC.nMOVNRO AND MO.nMOVITEM=MC.nMOVITEM "
'sSql = sSql & " LEFT JOIN CtaIFAdeudados cia ON cia.cIFTpo = mo.cIFTpo and cia.cPersCod = mo.cPersCod and cia.cCtaIFCod = mo.cCtaIFCod "
'sSql = sSql & " LEFT JOIN CtaIF  cii ON cii.cIFTpo = mo.cIFTpo and cii.cPersCod = mo.cPersCod and cii.cCtaIFCod = mo.cCtaIFCod "
'sSql = sSql & " JOIN OpeCta OC   ON MC.cCtaContCod LIKE OC.cCtaContCod + '%' "
'sSql = sSql & " LEFT JOIN IndiceVac i ON i.dIndiceVac = Convert(datetime,LEFT(m.cMovNro,8) ) "
'sSql = sSql & " WHERE M.copecod='40" & Mid(psOpeCod, 3, 1) & "806' AND M.nMOVFLAG NOT IN(1, 2, 3, 5) and m.nMovEstado = 10 "
'sSql = sSql & " AND SUBSTRING(M.CMOVNRO,1,8)> convert(varchar(8), cii.dCtaIFAper, 112) "
'sSql = sSql & " AND SUBSTRING(M.CMOVNRO,1,8)<='" & Trim(psFecha3) & "' "
'sSql = sSql & " and oc.cOpeCod = '46" & Mid(psOpeCod, 3, 1) & "092' "
'sSql = sSql & " And M.nMovNro Not IN(682714, 683330, 692474, 718469, 718476) "
'sSql = sSql & " GROUP BY MO.cIFTpo, MO.cPersCod, MO.cCtaIFCod ) Sald3"
'sSql = sSql & " ON Sald3.cPersCod = cif.cPersCod and Sald3.cCtaIFCod = ciF.cCtaIFCod and Sald3.cIFTpo = ciF.cIFTpo "
     
    'Interes
    sSql = sSql & " JOIN CtaIFInteres cii ON cii.cIFTpo = CIF.cIFTpo and cii.cPersCod = CIF.cPersCod and cii.cCtaIFCod = CIF.cCtaIFCod "
    sSql = sSql & " and cii.dCtaIFIntRegistro = (SELECT Max(dCtaIFIntRegistro) "
    sSql = sSql & " FROM CtaIFInteres cii1 WHERE cii1.cIFTpo = cii.cIFTpo "
    sSql = sSql & " and cii1.cPersCod = cii.cPersCod and cii1.cCtaIFCod = cii.cCtaIFCod ) "
      
    'Meses y Cuotas
    'sSql = sSql & " Inner Join ( Select CIC1.cPersCod, CIC1.cCtaIFCod, CIC1.cIFTpo, COUNT(*) as nCantidadCuotas, MAX(nCapital) as nMaximaCuota "
    'sSql = sSql & " From CtaIFCalendario CIC1 Group By CIC1.cPersCod, CIC1.cCtaIFCod, CIC1.cIFTpo  ) CIC2 "
    'sSql = sSql & " ON CIC2.cPersCod = cif.cPersCod and CIC2.cCtaIFCod = ciF.cCtaIFCod and CIC2.cIFTpo = ciF.cIFTpo "
      
    'cic2.nCantidadCuotas
    sSql = sSql & " Inner Join ( Select CIC1.cPersCod, CIC1.cCtaIFCod, CIC1.cIFTpo, COUNT(*) as nCantidadCuotas"
    sSql = sSql & " From CtaIFCalendario CIC1 where ctpocuota='2' "
    sSql = sSql & " Group By CIC1.cPersCod, CIC1.cCtaIFCod, CIC1.cIFTpo "
    sSql = sSql & " ) CIC2  ON CIC2.cPersCod = cif.cPersCod and CIC2.cCtaIFCod = ciF.cCtaIFCod and CIC2.cIFTpo = ciF.cIFTpo "
    
    'cic3.nProxCuota
    sSql = sSql & " Inner Join ( Select CIC1.cPersCod, CIC1.cCtaIFCod, CIC1.cIFTpo, CIC1.nCapital + CIC1.nInteres + CIC1.nComision as nProxCuota "
    sSql = sSql & " From CtaIFCalendario CIC1 "
    sSql = sSql & " Where CIC1.cTpoCuota='2' and CIC1.nNroCuota = "
    sSql = sSql & " ( "
    sSql = sSql & " Select MIN(CIC1A.nNroCuota) as nProxCuota "
    sSql = sSql & " From CtaIFCalendario CIC1A "
    sSql = sSql & " where CIC1A.ctpocuota='2' and convert(varchar(8), dvencimiento, 112)>'" & Format(ldFecha, "YYYYMMdd") & "' "
    sSql = sSql & " and CIC1A.cPersCod=CIC1.cPersCod and  CIC1A.cCtaIFCod=CIC1.cCtaIFCod and  CIC1A.cIFTpo = CIC1.cIFTpo "
    sSql = sSql & " Group By CIC1A.cPersCod, CIC1A.cCtaIFCod, CIC1A.cIFTpo "
    sSql = sSql & " )"
    sSql = sSql & " ) CIC3  ON CIC3.cPersCod = cif.cPersCod and CIC3.cCtaIFCod = ciF.cCtaIFCod and CIC3.cIFTpo = ciF.cIFTpo "

  
  
      
    sSql = sSql & "         LEFT JOIN ( SELECT mif.cPersCod, mif.cCtaIFCod, mif.cIFTpo, SUM(" & IIf(Mid(psOpeCod, 3, 1) = "2", "me.nMovMEImporte", "mc.nMovImporte") & ") nMonto " _
        & "                     FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro " & IIf(Mid(psOpeCod, 3, 1) = "2", "JOIN MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
        & "                                JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' " _
        & "                     WHERE m.nMovFlag = 0 and LEFT(m.cmovnro,8) > '" & Format(ldFecha, gsFormatoMovFecha) & "' and oc.copecod = '" & psOpeCod & "' and oc.cOpeCtaOrden = '0' " _
        & "                     GROUP BY mif.cPersCod, mif.cCtaIFCod, mif.cIFTpo ) Canc ON canc.cPersCod = cif.cPersCod and canc.cCtaIFCod = ciF.cCtaIFCod and canc.cIFTpo = ciF.cIFTpo "
        
    sSql = sSql & " Where (CIF.cctaifestado in(" & gEstadoCtaIFActiva & "," & gEstadoCtaIFRegistrada & ") or not Canc.nMonto is NULL ) " & _
          " And datediff(d, '" & Format(ldFecha, "MM/dd/YYYY") & "', dbo.AdeuGetFechaVencimiento(CIF.cPersCod, CIF.cIFTpo, CIF.cCtaIFCod)) > 360" & _
          " and convert(varchar(8), dbo.AdeuGetFechaVencimiento(CIF.cPersCod, CIF.cIFTpo, CIF.cCtaIFCod), 112) >='" & Format(ldFecha, "YYYYMMdd") & "' and Substring(CIF.cCtaIFCod, 3,1)='" & lsMoneda & "' " & lsFiltro & _
          " and convert(varchar(8), CIF.dCtaIFAper, 112)<='" & Format(ldFecha, "YYYYMMdd") & "' ORDER BY cif.cIFtpo, cif.cPersCod, ISNULL(cia.cCodLinCred,''), cif.cCtaIFDesc "
          
 Set CargaDatosADeudadosxEndeudamiento = oCon.CargaRecordSet(sSql)
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
   sSql = " SELECT cia.nSaldoCapLP SaldoLargoPlazo, dbo.AdeuGetFechaVencimiento(ci.cPersCod, ci.cIFTpo, ci.cCtaIFCod)  FechaVencimiento, ci.cIFTpo, ci.cPersCod, ci.cCtaIFCod, P.cPersNombre, ci.cCtaIFDesc, ci.dCtaIFAper, dCtaIFVenc, cia.nMontoPrestado, ci.nCtaIFPlazo, cia.nCtaIFCuotas, cia.nPeriodoGracia, cii.nCtaIFIntPeriodo, cii.nCtaIFIntValor nTasaInteres, cic.nNroCuota, ISNULL(cic.nInteresPagado,0) + " _
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
    sSql = sSql & "    ORDER BY ci.cIFtpo, ci.cPersCod, ISNULL(cia.cCodLinCred,''), ci.cCtaIFDesc  "

    'nSaldoCap
    Set oCon = New DConecta
    oCon.AbreConexion
    Set CargaDatosGeneralesCtaIF = oCon.CargaRecordSet(sSql)
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
   sSql = "SELECT ci.cIFTpo, ci.cPersCod, ci.cCtaIFCod, " _
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

    Set CargaDatosGeneralesCtaIFPlazoFijo = oCon.CargaRecordSet(sSql)
Exit Function
CargaDatosGeneralesCtaIFPlazoFijoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetLineaCredito(psMoneda As String, Optional psPersCod As String = "") As ADODB.Recordset
On Error GoTo GetLineaCreditoErr
Set oCon = New DConecta
oCon.AbreConexion
sSql = "SELECT cLineaCred, RTRIM(p.cPersNombre) + ' - ' + cDesCripcion cDescripcion, len(cLineaCred) nObjNivel, c.cPersCod " _
     & "FROM colocLineaCredito c JOIN Persona p ON p.cPersCod = c.cPersCod " _
     & "WHERE bEstado = 1 and LEN(cLineaCred) <= 6 and NOT cLineaCred = '' and ( (LEN(cLineaCred) = 5 and cLineaCred LIKE '____" & psMoneda & "%') or LEN(cLineaCred) < 5) " _
     & IIf(psPersCod = "", "", " and c.cPersCod = '" & psPersCod & "' ") _
     & "ORDER BY cLineaCred "
Set GetLineaCredito = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Exit Function
GetLineaCreditoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetLineaCreditoPersona(psLineaCod As String) As String
Dim rs As ADODB.Recordset
On Error GoTo GetLineaCreditoErr
sSql = "SELECT c.cPersCod " _
     & "FROM colocLineaCredito c " _
     & "WHERE bEstado = 1 and cLineaCred = '" & psLineaCod & "'  "
Set rs = oCon.CargaRecordSet(sSql)
If Not rs.EOF Then
    GetLineaCreditoPersona = rs!cPerscod
End If
RSClose rs
Exit Function
GetLineaCreditoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetAdeudadosCortoPlazo(psPersCod As String, psCtaIFCod As String, psIFTpo As String, pdFecha As Date) As Currency
Dim prs As ADODB.Recordset
On Error GoTo GetAdeudadosCortoPlazoErr
sSql = "SELECT ISNULL(SUM(nCapital),0) nSaldoCapCorto " _
     & "FROM  CtaIFCalendario " _
     & "WHERE cPersCod = '" & psPersCod & "' and cCtaIFCod = '" & psCtaIFCod & "' and cIFTpo = '" & psIFTpo & "' and " _
     & "      dVencimiento > '" & Format(pdFecha, gsFormatoFecha) & "' and dVencimiento <= '" & Format(DateAdd("yyyy", 1, pdFecha), gsFormatoFecha) & "' "
Set prs = oCon.CargaRecordSet(sSql)
If Not prs.EOF Then
    GetAdeudadosCortoPlazo = prs!nSaldoCapCorto
End If
Exit Function
GetAdeudadosCortoPlazoErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'Public Function GetAdeudadosSaldoCap(psPersCod As String, Optional psCtaIfCod As String = "", Optional psIFTpo As String = "", Optional pCon As ADODB.Connection, Optional pdFechaAl As Date = 0, Optional psFiltro As String = "", Optional psMoneda As String = "", Optional psTpoCuota As Integer = gCGTipoCuotCalIFCuota, Optional psFiltroNuevo As String = "") As Currency
'Dim prs As ADODB.Recordset
'On Error GoTo GetAdeudadosSaldoCapErr
'
'If Not IsNull(pCon) And Not pCon Is Nothing Then
'    lbConexionPropia = False
'    oCon.ConexionActiva = pCon
'Else
'    lbConexionPropia = True
'    Set oCon = New DConecta
'    oCon.AbreConexion
'End If
'If pdFechaAl = 0 Then
'
'    If psFiltroNuevo = "" Then
'        psFiltroNuevo = " and (cTpoCuota = '" & psTpoCuota & "') "
'    Else
'        psFiltroNuevo = psFiltroNuevo
'    End If
'
'    sSQL = "SELECT ISNULL(SUM(nCapital),0) nSaldoCap " _
'     & "FROM  CtaIFCalendario " _
'     & "WHERE (cPersCod = '" & psPersCod & "' and cCtaIFCod = '" & psCtaIfCod & "' and cIFTpo = '" & psIFTpo & "' and " _
'     & "      cEstado = '" & gTpoEstCuotaAdeudPend & "') " & psFiltroNuevo
'
'
'Else
'    sSQL = "SELECT cPersCod, " & IIf(psCtaIfCod = "", "", " cIFTpo, cCtaIFCod,") & " ISNULL(SUM(nSaldoCap),0) nSaldoCap " _
'     & "FROM (SELECT ci.cPersCod, " & IIf(psCtaIfCod = "", "", " ci.cIFTpo, ci.cCtaIFCod,") & " cia.nSaldoCap + " _
'     & "       ISNULL( (SELECT SUM(" & IIf(psMoneda = "2", "me.nMovImporte", "mc.nMovImporte") & ") " _
'     & "                FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro " & IIf(psMoneda = "2", " JOIN MovME me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
'     & "                       JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem " _
'     & "                       JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' " _
'     & "                WHERE m.nMovFlag = " & gMovFlagVigente & " and LEFT(m.cmovnro,8) > '" & Format(pdFechaAl, gsFormatoMovFecha) & "' and oc.cOpeCod IN ('" & OpeCGAdeudRepGeneralMN & "', '" & OpeCGAdeudRepGeneralME & "') and oc.cOpeCtaOrden = '0' " _
'     & "                  and not m.cOpeCod IN ('401804', '402804') " _
'     & "                  and mif.cPersCod = ci.cPersCod and mif.cCtaIFCod = ci.cCtaIFCod and mif.cIFTpo = ci.cIFTpo " _
'     & "       ), 0) as nSaldoCap " _
'     & "FROM ctaif ci join ctaifadeudados cia on cia.cperscod = ci.cperscod " _
'     & "                              and cia.cIFTpo = ci.cIFTpo and cia.cCtaIFCod = ci.cCtaIFCod " _
'     & "    WHERE cia.cPersCod = '" & psPersCod & "' and cia.cIFTpo + cia.cCtaIFCod like '" & psFiltro & "%' and cCtaIFEstado = 1 ) a " _
'     & "GROUP BY cPersCod " & IIf(psCtaIfCod = "", "", ", cIFTpo, cCtaIFCod ")
'
'End If
'Set prs = oCon.CargaRecordSet(sSQL)
'If Not prs.EOF Then
'    GetAdeudadosSaldoCap = prs!nSaldoCap
'End If
'
'
'Exit Function
'GetAdeudadosSaldoCapErr:
'    Err.Raise Err.Number, Err.Source, Err.Description
'End Function

Public Function GetAdeudadosSaldoCap(psPersCod As String, Optional psCtaIFCod As String = "", Optional psIFTpo As String = "", Optional pCon As ADODB.Connection, Optional pdFechaAl As Date = 0, Optional psFiltro As String = "", Optional psMoneda As String = "", Optional psTpoCuota As Integer = gCGTipoCuotCalIFCuota, Optional psFiltroNuevo As String = "") As Currency
Dim prs As ADODB.Recordset
On Error GoTo GetAdeudadosSaldoCapErr

If Not IsNull(pCon) And Not pCon Is Nothing Then
    lbConexionPropia = False
    oCon.ConexionActiva = pCon
Else
    lbConexionPropia = True
    Set oCon = New DConecta
    oCon.AbreConexion
End If
If pdFechaAl = 0 Then
    
    If psFiltroNuevo = "" Then
        psFiltroNuevo = " and (cTpoCuota = '" & psTpoCuota & "') "
    Else
        psFiltroNuevo = psFiltroNuevo
    End If
    
    sSql = "SELECT ISNULL(SUM(nCapital),0) nSaldoCap " _
     & "FROM  CtaIFCalendario " _
     & "WHERE (cPersCod = '" & psPersCod & "' and cCtaIFCod = '" & psCtaIFCod & "' and cIFTpo = '" & psIFTpo & "' and " _
     & "      cEstado = '" & gTpoEstCuotaAdeudPend & "') " & psFiltroNuevo
     
     
Else
    sSql = "SELECT cPersCod, " & IIf(psCtaIFCod = "", "", " cIFTpo, cCtaIFCod,") & " ISNULL(SUM(nSaldoCap),0) nSaldoCap " _
     & "FROM (SELECT ci.cPersCod, " & IIf(psCtaIFCod = "", "", " ci.cIFTpo, ci.cCtaIFCod,") & " cia.nSaldoCap + " _
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
     & "GROUP BY cPersCod " & IIf(psCtaIFCod = "", "", ", cIFTpo, cCtaIFCod ")
     
End If
Set prs = oCon.CargaRecordSet(sSql)
If Not prs.EOF Then
    GetAdeudadosSaldoCap = prs!nSaldoCap
End If


Exit Function
GetAdeudadosSaldoCapErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function GetAdeudadosProvisionMes(psOpeCod As String, psMovNro As String) As Currency
Dim prs As ADODB.Recordset
On Error GoTo GetAdeudadosProvisionMesErr
sSql = "SELECT " & IIf(Mid(psOpeCod, 3, 1) = "2", "me.nMovMEImporte", "mc.nMovImporte") & " nMovImporte " _
     & "FROM  Mov m Join MovCta mc on mc.nMovNro = m.nMovNro " _
     & IIf(Mid(psOpeCod, 3, 1) = "2", " join MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
     & "      JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' " _
     & "WHERE m.cMovNro = '" & psMovNro & "' and oc.cOpeCod = '" & psOpeCod & "' "
Set prs = oCon.CargaRecordSet(sSql)
If Not prs.EOF Then
    GetAdeudadosProvisionMes = prs!nMovImporte
End If
Exit Function
GetAdeudadosProvisionMesErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'Public Function GetAdeudadosOperaciones(psOpeCod As String, pdFechaDel As Date, pdFechaAl As Date) As ADODB.Recordset
'
'On Error GoTo GetAdeudadosOperacionesErr
'Set oCon = New DConecta
'sSQL = "SELECT Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103) dFecha, p.cPersNombre, cif.cCtaIFDesc, SUM(ISNULL(me.nMovMEImporte,mc.nMovImporte)) * -1 as nImporte, m.cMovDesc, cif.cIFTpo, cif.cPersCod, cif.cCtaIFCod, m.cMovNro, m.nMovNro, CEst.dFecUltActAnt dFechaMod "
'sSQL = sSQL & "FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro " _
'     & "           LEFT JOIN MovMe  me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'     & "           JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem " _
'     & "           JOIN CtaIF cif ON cif.cIFTpo = mif.cIFTpo and cif.cPersCod = mif.cPersCod and cif.cCtaIFCod = mif.cCtaIFCod " _
'     & "           JOIN CtaIFcalendario ca on ca.cIFTpo = mif.cIFTpo and ca.cPersCod = mif.cPersCod and ca.cCtaIFCod = mif.cCtaIFCod  and ca.cestado= 1 " _
'     & "           JOIN Persona p ON p.cPersCod = mif.cPersCod " _
'     & "           Join CtaIFEstadistica CEst On CEst.nMovNro = m.nMovNro " _
'     & "WHERE m.cOpeCod = '" & psOpeCod & "' and m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag = " & gMovFlagVigente _
'     & "  and LEFT(m.cMovNro,8) BETWEEN '" & Format(pdFechaDel, gsFormatoMovFecha) & "' and '" & Format(pdFechaAl, gsFormatoMovFecha) & "' "
'
'     If Mid(psOpeCod, 3, 1) = "1" Then
'        sSQL = sSQL & "  and mc.nMovImporte < 0 "
'    Else
'        sSQL = sSQL & " and ISNULL(me.nMovMEImporte,mc.nMovImporte) * -1 <0"
'    End If
'
'    sSQL = sSQL & " group by Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103), p.cPersNombre,"
'    sSQL = sSQL & " cif.cCtaIFDesc,m.cMovDesc, cif.cIFTpo, cif.cPersCod,cif.cCtaIFCod, m.cMovNro, m.nMovNro, CEst.dFecUltActAnt  "
'oCon.AbreConexion
'Set GetAdeudadosOperaciones = oCon.CargaRecordSet(sSQL)
'oCon.CierraConexion
'
'Exit Function
'GetAdeudadosOperacionesErr:
'    Err.Raise Err.Number, Err.Source, Err.Description
'End Function

Public Function GetAdeudadosOperaciones(psOpeCod As String, pdFechaDel As Date, pdFechaAl As Date) As ADODB.Recordset

On Error GoTo GetAdeudadosOperacionesErr
Set oCon = New DConecta
sSql = "SELECT Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103) dFecha, p.cPersNombre, cif.cCtaIFDesc, SUM(ISNULL(me.nMovMEImporte,mc.nMovImporte)) * -1 as nImporte, m.cMovDesc, cif.cIFTpo, cif.cPersCod, cif.cCtaIFCod, m.cMovNro, m.nMovNro, CEst.dFecUltActAnt dFechaMod, "
sSql = sSql & "Cest.nNroCuota , cif.cPerscod, cif.cIFTpo, cif.cCtaIFCod "
sSql = sSql & "FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro "
sSql = sSql & "           LEFT JOIN MovMe  me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem "
sSql = sSql & "           JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem "
sSql = sSql & "           JOIN CtaIF cif ON cif.cIFTpo = mif.cIFTpo and cif.cPersCod = mif.cPersCod and cif.cCtaIFCod = mif.cCtaIFCod "
'sSql = sSql & "           /*JOIN CtaIFcalendario ca on ca.cIFTpo = mif.cIFTpo and ca.cPersCod = mif.cPersCod and ca.cCtaIFCod = mif.cCtaIFCod  and ca.cestado=2*/ "
sSql = sSql & "           JOIN Persona p ON p.cPersCod = mif.cPersCod "
sSql = sSql & "           Join CtaIFEstadistica CEst On CEst.nMovNro = m.nMovNro "
sSql = sSql & "WHERE m.cOpeCod = '" & psOpeCod & "' and m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag = " & gMovFlagVigente
sSql = sSql & "  and LEFT(m.cMovNro,8) BETWEEN '" & Format(pdFechaDel, gsFormatoMovFecha) & "' and '" & Format(pdFechaAl, gsFormatoMovFecha) & "' "
     
     If Mid(psOpeCod, 3, 1) = "1" Then
        sSql = sSql & "  and mc.nMovImporte < 0 "
    Else
        sSql = sSql & " and ISNULL(me.nMovMEImporte,mc.nMovImporte) * -1 <0"
    End If
    
    sSql = sSql & " group by Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103), p.cPersNombre,"
    sSql = sSql & " cif.cCtaIFDesc,m.cMovDesc, cif.cIFTpo, cif.cPersCod,cif.cCtaIFCod, m.cMovNro, m.nMovNro, CEst.dFecUltActAnt,Cest.nNroCuota  "
oCon.AbreConexion
Set GetAdeudadosOperaciones = oCon.CargaRecordSet(sSql)
oCon.CierraConexion

Exit Function
GetAdeudadosOperacionesErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function



Public Function GetAdeudadosPagoCuota(psOpeCod As String, pdFechaDel As Date, pdFechaAl As Date) As ADODB.Recordset
On Error GoTo GetAdeudadosPagoCuotaErr
Dim oCon As New DConecta

sSql = "SELECT Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103) dFecha, Adeud.cPersNombre, Adeud.cCtaIFDesc, nImporte, m.cMovDesc, Adeud.cIFTpo, Adeud.cPersCod, Adeud.cCtaIFCod, m.cMovNro, m.nMovNro, RTRIM(Pag.cPersNombre) + ' ' + Pag.cCtaIFDesc EntidadPago " _
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
Set GetAdeudadosPagoCuota = oCon.CargaRecordSet(sSql)
oCon.CierraConexion

Exit Function
GetAdeudadosPagoCuotaErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


'Public Function GetUltimosMovimientosExtorno(ByVal cOpeCodBuscada As String, ByVal sCadena As String, ByVal cFechaMov As String) As ADODB.Recordset
'
'On Error GoTo GetUltimosMovimientosExtornoErr
'
'    sSQL = "Select M.cMovNro, M.cMovDesc, M.cOpeCod, O.cOpeDesc, "
'    sSQL = sSQL & " CIE.nTipoEstadistica, C1.cConsDescripcion, "
'    sSQL = sSQL & " CIE.cPersCod, CIE.cIFTpo, CIE.cCtaIFCod, "
'    sSQL = sSQL & " P.cPersNombre, CIF.cCtaIFDesc, "
'    sSQL = sSQL & " CIE.nMovNro, CIE.cTpoCuota, CIE.nNroCuota, isnull(CIE.cEstAnt, '') cEstAnt, ISNULL(CIA.dCancelacion, '01/01/1900') as dCancelacion, "
'
'    'Solo en Extorno de Pagos y Provisiones Relacionar con Calendario
'    If cOpeCodBuscada = "401705" Or cOpeCodBuscada = "402705" Or cOpeCodBuscada = "401706" Or cOpeCodBuscada = "402706" Then
'        sSQL = sSQL & " CIC.dVencimiento , CIE.nTipoEstadistica, C1.cConsDescripcion, "
'    End If
'
'    sSQL = sSQL & " CIE.nCapitalMovido, CIE.nInteres1Movido, CIE.nInteres2Movido, CIE.nComisionMovida, CIE.nSaldoAnt, CIE.nSaldoNew, "
'    sSQL = sSQL & " CIE.nSaldoLPAnt, CIE.nSaldoLPNew, CIE.bVac, CIEDet.nCapitalMovidoReal, CIEDet.nInteres1MovidoReal, CIEDet.nInteres2MovidoReal, "
'    sSQL = sSQL & " CIEDet.nComisionMovidaReal, CIEDet.nSaldoAntReal, CIEDet.nSaldoNewReal, CIEDet.nSaldoLPAntReal, CIEDet.nSaldoLPNewReal, "
'    sSQL = sSQL & " CIE.dFecUltActAnt , CIE.dFecUltActNew, CIE.dFecUltPagoAnt, CIE.dFecUltPagoNew, "
'    sSQL = sSQL & " IsNull(CIA.cCodLinCred, '') as cCodLinCred "
'
'    sSQL = sSQL & " From CtaIFEstadistica CIE "
'
'            '** Cambio 2
'            sSQL = sSQL & " Inner Join CtaIFEstadisticaDet CIEDet On CIE.nMovNro=CIEDet.nMovNro"
'
'    sSQL = sSQL & "     Inner Join (   Select CIE.cPersCod, CIE.cIFTpo, CIE.cCtaIFCod, MAX(CIE.nMovNro) as nMovNro "
'    sSQL = sSQL & "             From CtaIFEstadistica CIE "
'    sSQL = sSQL & "                 Inner Join Mov M On CIE.nMovNro=M.nMovNro "
'    sSQL = sSQL & "             Where m.nMovFlag = 0 "
'    sSQL = sSQL & "             Group By CIE.cPersCod, CIE.cIFTpo, CIE.cCtaIFCod ) CIE1 "
'
'    sSQL = sSQL & "         On CIE.cPErsCod=CIE1.cPErsCod And CIE.cIFTpo=CIE1.cIFTpo And CIE.cCtaIFCod=CIE1.cCtaIFCod And CIE.nMovNro=CIE1.nMovNro"
'    sSQL = sSQL & "     Inner Join Mov M On CIE.nMovNro=M.nMovNro "
'    sSQL = sSQL & "     Inner Join CtaIF CIF On CIE.cPerscod=CIF.cPerscod And CIE.cIFTpo = CIF.cIFTpo And CIE.cCtaIFCod=CIF.cCtaIFCod "
'    sSQL = sSQL & "     Inner Join Persona P On CIF.cPersCod=P.cPErsCod "
'    sSQL = sSQL & "     Inner Join Constante C1 On CIE.nTipoEstadistica = C1.nConsValor "
'    sSQL = sSQL & "     Inner Join OpeTpo O On M.cOpeCod=O.cOpeCod "
'    sSQL = sSQL & "     Inner Join CtaIFAdeudados CIA "
'    sSQL = sSQL & "         On CIF.cPersCod=CIA.cPersCod And CIF.cIFTpo=CIA.cIFTpo And CIF.cCtaIFCod=CIA.cCtaIFCod "
'
'    'Solo en Extorno de Pagos y Provisiones Relacionar con Calendario
'    If cOpeCodBuscada = "401705" Or cOpeCodBuscada = "402705" Or cOpeCodBuscada = "401706" Or cOpeCodBuscada = "402706" Then
'        sSQL = sSQL & "     Inner Join CtaIFCalendario CIC "
'        sSQL = sSQL & "         On CIF.cPersCod=CIC.cPersCod And CIF.cIFTpo=CIC.cIFTpo And CIF.cCtaIFCod=CIC.cCtaIFCod AND CIE.cTpoCuota=CIC.cTpoCuota And CIE.nNroCuota=CIC.nNroCuota"
'    End If
'
'    sSQL = sSQL & "     Left Join ColocLineaCredito CLC "
'    sSQL = sSQL & "           On CIA.cCodLinCred=CLC.cLineaCred "
'
'    sSQL = sSQL & " Where M.nMovFlag=0 And C1.nConsCod = 4034 "
'    sSQL = sSQL & " And M.cOpeCod='" & cOpeCodBuscada & "' "
'
'    'Solo en Extorno de Pagos y Provisiones Relacionar con Calendario
'    If cOpeCodBuscada = "401705" Or cOpeCodBuscada = "402705" Or cOpeCodBuscada = "401706" Or cOpeCodBuscada = "402706" Then
'        sSQL = sSQL & " And CIC.bVigente=1 "
'    End If
'
'    sSQL = sSQL & " And M.cMovNro Like '" & cFechaMov & "%'"
'
'    If Len(Trim(sCadena)) > 0 Then
'        sSQL = sSQL & sCadena
'    End If
'
'    Set GetUltimosMovimientosExtorno = oCon.CargaRecordSet(sSQL)
'
'Exit Function
'GetUltimosMovimientosExtornoErr:
'    Call RaiseError(MyUnhandledError, "DACGAdeudados:GetUltimonMovNro Method")
'
'
'End Function

Public Function GetUltimosMovimientosExtorno(ByVal cOpeCodBuscada As String, ByVal sCadena As String, ByVal cFechaMov As String, ByVal cFechaMovFin As String) As ADODB.Recordset
On Error GoTo GetUltimosMovimientosExtornoErr
 
    sSql = "Select M.cMovNro, M.cMovDesc, M.cOpeCod, O.cOpeDesc, "
    sSql = sSql & " CIE.nTipoEstadistica, C1.cConsDescripcion, "
    sSql = sSql & " CIE.cPersCod, CIE.cIFTpo, CIE.cCtaIFCod, "
    sSql = sSql & " P.cPersNombre, CIF.cCtaIFDesc, "
    sSql = sSql & " CIE.nMovNro, CIE.cTpoCuota, CIE.nNroCuota, isnull(CIE.cEstAnt, '') cEstAnt, ISNULL(CIA.dCancelacion, '01/01/1900') as dCancelacion, "
    
    'Solo en Extorno de Pagos y Provisiones Relacionar con Calendario
    If cOpeCodBuscada = "401705" Or cOpeCodBuscada = "402705" Or cOpeCodBuscada = "401706" Or cOpeCodBuscada = "402706" Then
        sSql = sSql & " CIC.dVencimiento , CIE.nTipoEstadistica, C1.cConsDescripcion, "
    End If
    
    sSql = sSql & " CIE.nCapitalMovido, CIE.nInteres1Movido, CIE.nInteres2Movido, CIE.nComisionMovida, CIE.nSaldoAnt, CIE.nSaldoNew, "
    sSql = sSql & " CIE.nSaldoLPAnt, CIE.nSaldoLPNew, CIE.bVac, CIEDet.nCapitalMovidoReal, CIEDet.nInteres1MovidoReal, CIEDet.nInteres2MovidoReal, "
    sSql = sSql & " CIEDet.nComisionMovidaReal, CIEDet.nSaldoAntReal, CIEDet.nSaldoNewReal, CIEDet.nSaldoLPAntReal, CIEDet.nSaldoLPNewReal, "
    sSql = sSql & " CIE.dFecUltActAnt , CIE.dFecUltActNew, CIE.dFecUltPagoAnt, CIE.dFecUltPagoNew, "
    sSql = sSql & " IsNull(CIA.cCodLinCred, '') as cCodLinCred "
    
    sSql = sSql & " From CtaIFEstadistica CIE "
    
            '** Cambio 2
            sSql = sSql & " Inner Join CtaIFEstadisticaDet CIEDet On CIE.nMovNro=CIEDet.nMovNro"
    
    sSql = sSql & "     Inner Join (   Select CIE.cPersCod, CIE.cIFTpo, CIE.cCtaIFCod, MAX(CIE.nMovNro) as nMovNro "
    sSql = sSql & "             From CtaIFEstadistica CIE "
    sSql = sSql & "                 Inner Join Mov M On CIE.nMovNro=M.nMovNro "
    sSql = sSql & "             Where m.nMovFlag = 0 "
    sSql = sSql & "             Group By CIE.cPersCod, CIE.cIFTpo, CIE.cCtaIFCod ) CIE1 "
    
    sSql = sSql & "         On CIE.cPErsCod=CIE1.cPErsCod And CIE.cIFTpo=CIE1.cIFTpo And CIE.cCtaIFCod=CIE1.cCtaIFCod And CIE.nMovNro=CIE1.nMovNro"
    sSql = sSql & "     Inner Join Mov M On CIE.nMovNro=M.nMovNro "
    sSql = sSql & "     Inner Join CtaIF CIF On CIE.cPerscod=CIF.cPerscod And CIE.cIFTpo = CIF.cIFTpo And CIE.cCtaIFCod=CIF.cCtaIFCod "
    sSql = sSql & "     Inner Join Persona P On CIF.cPersCod=P.cPErsCod "
    sSql = sSql & "     Inner Join Constante C1 On CIE.nTipoEstadistica = C1.nConsValor "
    sSql = sSql & "     Inner Join OpeTpo O On M.cOpeCod=O.cOpeCod "
    sSql = sSql & "     Inner Join CtaIFAdeudados CIA "
    sSql = sSql & "         On CIF.cPersCod=CIA.cPersCod And CIF.cIFTpo=CIA.cIFTpo And CIF.cCtaIFCod=CIA.cCtaIFCod "
    
    'Solo en Extorno de Pagos y Provisiones Relacionar con Calendario
    If cOpeCodBuscada = "401705" Or cOpeCodBuscada = "402705" Or cOpeCodBuscada = "401706" Or cOpeCodBuscada = "402706" Then
        sSql = sSql & "     Inner Join CtaIFCalendario CIC "
        sSql = sSql & "         On CIF.cPersCod=CIC.cPersCod And CIF.cIFTpo=CIC.cIFTpo And CIF.cCtaIFCod=CIC.cCtaIFCod AND CIE.cTpoCuota=CIC.cTpoCuota And CIE.nNroCuota=CIC.nNroCuota"
    End If
    
    sSql = sSql & "     Left Join ColocLineaCredito CLC "
    sSql = sSql & "           On CIA.cCodLinCred=CLC.cLineaCred "
    
    sSql = sSql & " Where M.nMovFlag=0 And C1.nConsCod = 4034 "
    sSql = sSql & " And M.cOpeCod='" & cOpeCodBuscada & "' "
    
    'Solo en Extorno de Pagos y Provisiones Relacionar con Calendario
    If cOpeCodBuscada = "401705" Or cOpeCodBuscada = "402705" Or cOpeCodBuscada = "401706" Or cOpeCodBuscada = "402706" Then
        sSql = sSql & " And CIC.bVigente=1 "
    End If
    
    sSql = sSql & " And M.cMovNro Between '" & cFechaMov & "' And '" & cFechaMovFin & "'"
    
    If Len(Trim(sCadena)) > 0 Then
        sSql = sSql & sCadena
    End If
    
    Set oCon = New DConecta
    oCon.AbreConexion
    
    Set GetUltimosMovimientosExtorno = oCon.CargaRecordSet(sSql)

Exit Function
GetUltimosMovimientosExtornoErr:
    Call RaiseError(MyUnhandledError, "DACGAdeudados:GetUltimonMovNro Method")
 
 
End Function


Public Function GetUltimosMovimientosExtorno11(ByVal cOpeCodBuscada As String, ByVal sCadena As String, ByVal cFechaMov As String, ByVal cFechaMovFin As String) As ADODB.Recordset
On Error GoTo GetUltimosMovimientosExtornoErr
    
    sSql = "Select Convert(varchar(10),Convert(datetime,LEFT(m.cMovNro,8)),103) dFecha, P.cPersNombre, CIF.cCtaIFDesc, IsNull(CIE.nCapitalMovido,0) + IsNull(CIE.nInteres1Movido,0) + IsNull(CIE.nInteres2Movido,0) + IsNull(CIE.nComisionMovida,0) nImporte, M.cMovDesc, CIE.cPersCod, CIE.cIFTpo, CIE.cCtaIFCod, M.cMovNro, M.nMovNro,RP='', CIE.nNroCuota,cif.cPersCod,cif.cIFTpo,cif.cCtaIfCod"
    sSql = sSql & " From CtaIFEstadistica CIE "
    '** Cambio 2
    sSql = sSql & " Inner Join CtaIFEstadisticaDet CIEDet On CIE.nMovNro=CIEDet.nMovNro"
    sSql = sSql & "     Inner Join (   Select CIE.cPersCod, CIE.cIFTpo, CIE.cCtaIFCod, MAX(CIE.nMovNro) as nMovNro "
    sSql = sSql & "             From CtaIFEstadistica CIE "
    sSql = sSql & "                 Inner Join Mov M On CIE.nMovNro=M.nMovNro "
    sSql = sSql & "             Where m.nMovFlag = 0 "
    sSql = sSql & "             Group By CIE.cPersCod, CIE.cIFTpo, CIE.cCtaIFCod ) CIE1 "
    
    sSql = sSql & "         On CIE.cPErsCod=CIE1.cPErsCod And CIE.cIFTpo=CIE1.cIFTpo And CIE.cCtaIFCod=CIE1.cCtaIFCod And CIE.nMovNro=CIE1.nMovNro"
    sSql = sSql & "     Inner Join Mov M On CIE.nMovNro=M.nMovNro "
    sSql = sSql & "     Inner Join CtaIF CIF On CIE.cPerscod=CIF.cPerscod And CIE.cIFTpo = CIF.cIFTpo And CIE.cCtaIFCod=CIF.cCtaIFCod "
    sSql = sSql & "     Inner Join Persona P On CIF.cPersCod=P.cPErsCod "
    sSql = sSql & "     Inner Join Constante C1 On CIE.nTipoEstadistica = C1.nConsValor "
    sSql = sSql & "     Inner Join OpeTpo O On M.cOpeCod=O.cOpeCod "
    sSql = sSql & "     Inner Join CtaIFAdeudados CIA "
    sSql = sSql & "         On CIF.cPersCod=CIA.cPersCod And CIF.cIFTpo=CIA.cIFTpo And CIF.cCtaIFCod=CIA.cCtaIFCod "
    
    'Solo en Extorno de Pagos y Provisiones Relacionar con Calendario
    If cOpeCodBuscada = "401705" Or cOpeCodBuscada = "402705" Or cOpeCodBuscada = "401706" Or cOpeCodBuscada = "402706" Then
        sSql = sSql & "     Inner Join CtaIFCalendario CIC "
        sSql = sSql & "         On CIF.cPersCod=CIC.cPersCod And CIF.cIFTpo=CIC.cIFTpo And CIF.cCtaIFCod=CIC.cCtaIFCod AND CIE.cTpoCuota=CIC.cTpoCuota And CIE.nNroCuota=CIC.nNroCuota"
    End If
    
    sSql = sSql & "     Left Join ColocLineaCredito CLC "
    sSql = sSql & "           On CIA.cCodLinCred=CLC.cLineaCred "
    
    sSql = sSql & " Where M.nMovFlag=0 And C1.nConsCod = 4034 "
    sSql = sSql & " And M.cOpeCod='" & cOpeCodBuscada & "' "
    
    'Solo en Extorno de Pagos y Provisiones Relacionar con Calendario
    If cOpeCodBuscada = "401705" Or cOpeCodBuscada = "402705" Or cOpeCodBuscada = "401706" Or cOpeCodBuscada = "402706" Then
        sSql = sSql & " And CIC.bVigente=1 "
    End If
    
    sSql = sSql & " And M.cMovNro Between '" & cFechaMov & "' And '" & cFechaMovFin & "'"
    
    If Len(Trim(sCadena)) > 0 Then
        sSql = sSql & sCadena
    End If
    
    Set oCon = New DConecta
    oCon.AbreConexion
    
    Set GetUltimosMovimientosExtorno11 = oCon.CargaRecordSet(sSql)

Exit Function
GetUltimosMovimientosExtornoErr:
    Call RaiseError(MyUnhandledError, "DACGAdeudados:GetUltimonMovNro Method")
 
 
End Function



Public Function CargaDatosAdeudoReclasifica(lsFiltroCta As String, Optional lnIndiceVac As Double, Optional pdFecha As Date, Optional psOpeCod As String = "", Optional pbSoloCortoPlazo As Boolean = False) As ADODB.Recordset
Dim lsFiltro As String
On Error GoTo CargaDatosGeneralesCtaIFErr
    
   sSql = ""
   sSql = sSql & "SELECT p.cPersNombre, ci.cIFTpo, ci.cPersCod, ci.cCtaIFCod, ci.cCtaIFDesc, cia.nSaldoCap nSaldoCapCal, ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") nVac, cia.cMonedaPago,"
   sSql = sSql & "         round( cia.nSaldoCap * convert(decimal(12,5), CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1.0 END),2) +"
   sSql = sSql & "             ISNULL( Canc.nMonto, 0) nSaldoCap, ISNULL(cic.nSaldoCapCorto,0) nSaldoCapCortoCal,"
   sSql = sSql & "         round( ISNULL(cic.nSaldoCapCorto,0) * convert(decimal(12,5), CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1.0 END),2) nSaldoCapCorto, ISNULL(sdo.nSaldoCnt,0) nSaldoCnt "
   sSql = sSql & "FROM CtaIF ci JOIN CtaIfAdeudados cia ON cia.cIFTpo = ci.cIFTpo and cia.cPersCod = ci.cPersCod and cia.cCtaIFCod = ci.cCtaIFCod"
   sSql = sSql & "              LEFT JOIN Persona p ON p.cPersCod = ci.cPersCod "
   sSql = sSql & "              LEFT JOIN IndiceVac iv ON iv.dIndiceVac = ISNULL(cia.dCuotaUltPago, ci.dCtaIFAper)"
   sSql = sSql & "              LEFT JOIN ( SELECT cIFTpo, cPersCod, cCtaIFCod, ISNULL(SUM(nCapital),0) nSaldoCapCorto"
   sSql = sSql & "                From CtaIFCalendario"
   sSql = sSql & "                          WHERE dVencimiento > '" & Format(pdFecha, gsFormatoFecha) & "' and dVencimiento <= '" & Format(DateAdd("yyyy", 1, pdFecha), gsFormatoFecha) & "'"
   sSql = sSql & "                          GROUP BY cIFTpo, cPersCod, cCtaIFCod"
   sSql = sSql & "                        ) cic ON cic.cIFTpo = ci.cIFTpo And cic.cPersCod = ci.cPersCod And cic.cCtaIFCod = ci.cCtaIFCod"
   sSql = sSql & "              LEFT JOIN ( SELECT mif.cPersCod, mif.cCtaIFCod, mif.cIFTpo, SUM(" & IIf(Mid(psOpeCod, 3, 1) = "2", "me.nMovMEImporte", "mc.nMovImporte") & ") nMonto"
   sSql = sSql & "                          FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro " & IIf(Mid(psOpeCod, 3, 1) = "2", "JOIN MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "")
   sSql = sSql & "                                     JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%'"
   sSql = sSql & "                          WHERE m.nMovFlag = 0 and LEFT(m.cmovnro,8) > '" & Format(pdFecha, gsFormatoMovFecha) & "' and oc.copecod = '" & psOpeCod & "' and oc.cOpeCtaOrden = '0'"
   sSql = sSql & "                          GROUP BY mif.cPersCod, mif.cCtaIFCod, mif.cIFTpo"
   sSql = sSql & "                        ) Canc ON canc.cPersCod = ci.cPersCod and canc.cCtaIFCod = ci.cCtaIFCod and canc.cIFTpo = ci.cIFTpo "
   sSql = sSql & "             LEFT JOIN ( SELECT cs.cPersCod, cs.cCtaIFCod, cs.cIFTpo, SUM(" & IIf(Mid(psOpeCod, 3, 1) = 1, "cs.nSaldo", "cs.nSaldoME") & ") nSaldoCnt"
   sSql = sSql & "                         FROM CtaIFSaldo cs JOIN OpeCta oc ON cs.cCtaContCod LIKE oc.cCtaContCod + '%'"
   sSql = sSql & "                         WHERE oc.cOpeCod = '" & psOpeCod & "' and dCtaIFSaldo = ( SELECT Max(dCtaIFSaldo) FROM CtaIFSaldo cs1"
   sSql = sSql & "                                          Where cs1.cCtaContCod = cs.cCtaContCod And cs1.cPersCod = cs.cPersCod And cs1.cIFTpo = cs.cIFTpo And cs1.cCtaIFCod = cs.cCtaIFCod"
   sSql = sSql & "                                                and dCtaIFSaldo <= '" & Format(pdFecha, gsFormatoFecha) & "' )"
   sSql = sSql & IIf(pbSoloCortoPlazo, " and oc.cCtaContCod LIKE '24%'", "")
   sSql = sSql & "                         GROUP BY cs.cPersCod, cs.cCtaIFCod, cs.cIFTpo"
   sSql = sSql & "                       ) Sdo ON Sdo.cPersCod = ci.cPersCod and Sdo.cIFTpo = ci.cIFTpo and Sdo.cCtaIFCod = ci.cCtaIFCod "
   sSql = sSql & "WHERE (ci.cCtaIFEstado in (" & gEstadoCtaIFActiva & "," & gEstadoCtaIFRegistrada & ") or not Canc.nMonto is NULL )"
   sSql = sSql & "  and  ci.cIFTpo+ci.cCtaIFCod LIKE '" & lsFiltroCta & "' " & IIf(pbSoloCortoPlazo, "", "and cia.cMonedaPago = '2'")
   sSql = sSql & "  and datediff(d,ci.dCtaIFAper,'" & Format(pdFecha, gsFormatoFecha) & "') >= 0 "
   sSql = sSql & "ORDER BY p.cPersNombre, ci.cCtaIFCod "

    Set CargaDatosAdeudoReclasifica = oCon.CargaRecordSet(sSql)
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
    
sSql = "select dFecDataConsol "
sSql = sSql & " From CGAdeudCierreConsol where convert(varchar(8), dfeccierre, 112)='" & Format(ldFecha, "YYYYMMdd") & "'"

Set rs = oCon.CargaRecordSet(sSql)
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
    
sSql = "  SELECT DISTINCT "
sSql = sSql & "     convert(integer, datediff(d, ci.dCtaIFAper, dbo.AdeuGetFechaVencimiento(ci.cPersCod, ci.cIFTpo, ci.cCtaIFCod))/360) as nAnno, "
sSql = sSql & "     convert(integer, convert(integer, datediff(d, ci.dCtaIFAper, dbo.AdeuGetFechaVencimiento(ci.cPersCod, ci.cIFTpo, ci.cCtaIFCod))%360) / 30) as nMes, "
sSql = sSql & "     ci.cIFTpo, ci.cPersCod, ci.cCtaIFCod, P.cPersNombre, "
sSql = sSql & "     cii.nCtaIFIntPeriodo, cii.nCtaIFIntValor nTasaInteres, "
sSql = sSql & "     cia.nSaldoCap nSaldoCapCal, ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") nVac, cia.cMonedaPago, "
sSql = sSql & "     Round( cia.nSaldoCap * convert(decimal(15,5), CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' "
sSql = sSql & "     and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END),2) nSaldoCap, "
'sSQL = sSQL & "     + ISNULL( Canc.nMonto, 0) nSaldoCap , "
sSql = sSql & "     ISNULL(cia.cCodLinCred,'') cCodLinCred, ISNULL(l.cDescripcion,'') cDesLinCred "
sSql = sSql & " FROM " & cCtaIF & " ci LEFT JOIN " & cCtaIFAdeudados & " cia ON cia.cIFTpo = ci.cIFTpo and cia.cPersCod = ci.cPersCod "
sSql = sSql & "     and cia.cCtaIFCod = ci.cCtaIFCod "
sSql = sSql & " INNER JOIN ColocLineaCredito l ON l.cLineaCred = cia.cCodLinCred JOIN Persona P ON ci.cPersCod = p.cPersCod "
sSql = sSql & " LEFT JOIN IndiceVac iv ON iv.dIndiceVac = ISNULL(cia.dCuotaUltPago, ci.dCtaIFAper) "
sSql = sSql & " JOIN " & cCtaIFInteres & " cii ON cii.cIFTpo = ci.cIFTpo and cii.cPersCod = ci.cPersCod and cii.cCtaIFCod = ci.cCtaIFCod "
sSql = sSql & "     and cii.dCtaIFIntRegistro = (   SELECT Max(dCtaIFIntRegistro) "
sSql = sSql & "                                     FROM " & cCtaIFInteres & " cii1 "
sSql = sSql & "                                     Where cii1.cIFTpo = cii.cIFTpo And cii1.cPersCod = cii.cPersCod And cii1.cCtaIFCod = cii.cCtaIFCod "

    If bRegistrado = True Then
        sSql = sSql & " And Convert(varchar(8), cii1.dfecha,112)='" & cFecha & "' "
    End If

sSql = sSql & "                                 )"

'Cancelaciones
'sSQL = sSQL & " LEFT JOIN (     SELECT mif.cPersCod, mif.cCtaIFCod, mif.cIFTpo, SUM(mc.nMovImporte) nMonto "
'sSQL = sSQL & "                 FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro "
'If lsMoneda = "2" Then
'    sSQL = sSQL & "             JOIN MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem "
'End If
'sSQL = sSQL & "                 JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem "
'sSQL = sSQL & "                 JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' "
'sSQL = sSQL & "                 WHERE m.nMovFlag = 0 and LEFT(m.cmovnro,8) > '" & Format(ldFecha, "YYYYMMdd") & "' "
'sSQL = sSQL & "                       and oc.copecod = '46" & lsMoneda & "091' and oc.cOpeCtaOrden = '0' "
'sSQL = sSQL & "                 GROUP BY mif.cPersCod, mif.cCtaIFCod, mif.cIFTpo ) Canc "
'sSQL = sSQL & "     ON canc.cPersCod = ci.cPersCod and canc.cCtaIFCod = ci.cCtaIFCod and canc.cIFTpo = ci.cIFTpo "
'Fin Cancelaciones

sSql = sSql & " WHERE (ci.cCtaIFEstado in (1,0) )" 'or not Canc.nMonto is NULL ) "
sSql = sSql & "     and  ci.cIFTpo+ci.cCtaIFCod LIKE '__05" & lsMoneda & "%'"
sSql = sSql & "     and datediff(d,ci.dCtaIFAper,'" & Format(ldFecha, "MM/dd/YYYY") & "') >= 0 "
sSql = sSql & lsFiltro

If bRegistrado = True Then
    sSql = sSql & " And Convert(varchar(8), ci.dfecha,112)='" & cFecha & "' "
    sSql = sSql & " And Convert(varchar(8), cia.dfecha,112)='" & cFecha & "' "
    sSql = sSql & " And Convert(varchar(8), cii.dfecha,112)='" & cFecha & "' "
End If

sSql = sSql & " ORDER BY ci.cPersCod, ISNULL(cia.cCodLinCred,''), convert(integer, datediff(d, ci.dCtaIFAper, dbo.AdeuGetFechaVencimiento(ci.cPersCod, ci.cIFTpo, ci.cCtaIFCod))/360) "
     
Set CargaDatosADeudadosPYMES = oCon.CargaRecordSet(sSql)
Exit Function
CargaDatosAdeudadosProyectadosErr:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Public Function GrabaProvisionAdeudado(ByVal pcOpeCod As String, ByVal pcMoneda As String, ByVal pdFechaProvision As Date, ByVal lnVac As Double, _
                                    ByVal lstListadoProvision As Variant, ByVal cGlosa As String, ByVal psCodAge As String, ByVal psCodUser As String, _
                                    ByRef lsListaAsientos As String)
Dim oDMov As DMov
Dim lsMsgErr As String
Dim lsPersCod As String
Dim lsIFTpo As String
Dim lsCtaIFCod As String

'CALCULA CUENTA
Dim lbCalculaCuenta As Boolean
lbCalculaCuenta = True

Dim lscCodLinCred As String

Dim lsCtaDebe As String
Dim lsCtaHaber As String

Dim lsCuentaCorto As String
Dim lsCuentaComsionAdelantadaD As String
Dim lsCuentaComsionAdelantadaH As String
Dim lnComisionAdelantada As Currency
Dim lsCuentaLargo As String
Dim lsCuentaVac As String

Dim cMonedaPago As String

Dim I As Integer
Dim item As Integer
Dim lnMovNro As Long

Dim lsMovNro As String
Dim lbTrans As Boolean

Dim cTpoCuota As String 'Tipo de Cuota
Dim nNroCuota As Integer 'Numero de Cuota

Dim lnInteres As Double 'Interes Provisionado con VAC
Dim lnInteresReal As Double 'Interes Provisionado REAL SIN VAC

Dim lnSKAntReal As Double 'SK Anterior REAL SIN VAC

Dim lnSKLPAntReal As Double 'SK Largo Plazo REAL SIN VAC
Dim lnSKLPNewReal As Double 'SK Largo Plazo REAL NUEVO SIN VAC
  
Dim lnSKxVacAnt As Double 'SK * VAC Anterior
Dim lnSKxVacNew As Double 'SK * VAC Nuevo

Dim lnSKLPxVacAnt As Double 'SK Largo Plazo * VAC Anterior
Dim lnSKLPxVacNew As Double 'SK Largo Plazo * VAC Nuevo

Dim dFecUltActSaldos As Date 'Fecha de Ultima Provision o Pago Efectuado Anterior (Ultima Modificacion)
Dim dFecUltPago As Date 'Fecha de Ultima Pago Efectuado Anterior

Dim lnMontoVacCP As Double
Dim lnMontoVacLP As Double
Dim lnMontoVacTotal As Double

On Error GoTo CGGrabaProvisionAdeudadoErr
'
Set oDMov = New DMov
      
GrabaProvisionAdeudado = 1
                   
'
    For I = 1 To lstListadoProvision.ListItems.Count
    
        If lstListadoProvision.ListItems(I).Checked = True Then
            
            lbTrans = True
            'Transaccion
            oDMov.BeginTrans
            
            ' ==================== INICIALIZAMOS VALORES ====================
            
            lsPersCod = ""
            lsIFTpo = ""
            lsCtaIFCod = ""
            cMonedaPago = ""
            cTpoCuota = ""
            nNroCuota = 0
            dFecUltActSaldos = "01/01/1900"
            dFecUltPago = "01/01/1900"
            lscCodLinCred = ""
            lsCtaDebe = ""
            lsCtaHaber = ""
            lsCuentaCorto = ""
            lsCuentaLargo = ""
            lsCuentaComsionAdelantadaD = ""
            lsCuentaComsionAdelantadaH = ""
            lnComisionAdelantada = 0
            lsCuentaVac = ""
            lsMovNro = ""
            lnInteres = 0
            lnInteresReal = 0
            lnSKAntReal = 0
            lnSKLPAntReal = 0
            lnSKLPNewReal = 0
            lnSKxVacAnt = 0
            lnSKxVacNew = 0
            lnSKLPxVacAnt = 0
            lnSKLPxVacNew = 0
            
            lnMontoVacCP = 0
            lnMontoVacLP = 0
            lnMontoVacTotal = 0
            
            item = 0
            
            ' ==================== OBTENEMOS VALORES ====================
            lsPersCod = Mid(lstListadoProvision.ListItems(I).SubItems(14), 4, 13)
            lsIFTpo = Mid(lstListadoProvision.ListItems(I).SubItems(14), 1, 2)
            lsCtaIFCod = Mid(lstListadoProvision.ListItems(I).SubItems(14), 18, 10)
            cMonedaPago = lstListadoProvision.ListItems(I).SubItems(13)
            
            'Tipo de Cuota
            cTpoCuota = lstListadoProvision.ListItems(I).SubItems(24)
            'Numero de Cuota
            nNroCuota = Val(lstListadoProvision.ListItems(I).SubItems(6))
            
            'Fecha de Ultima Provision o Pago Efectuado Anterior
            dFecUltActSaldos = CDate(lstListadoProvision.ListItems(I).SubItems(23))
            
            'Fecha de Ultimo Pago Efectuado Anterior
            dFecUltPago = CDate(lstListadoProvision.ListItems(I).SubItems(5))
            
            'SubCuenta Contable asignada a la Linea de Credito
            lscCodLinCred = GetSubCtaLineaCredito(lstListadoProvision.ListItems(I).SubItems(17))
              
            'Cuentas Provision Interes
            lsCtaDebe = GetOpeCta(pcOpeCod, "D", "1", lsPersCod, lsIFTpo) & lscCodLinCred
            lsCtaHaber = GetOpeCta(pcOpeCod, "H", "1", lsPersCod, lsIFTpo)
            If DateDiff("d", pdFechaProvision, lstListadoProvision.ListItems(I).SubItems(15)) > 365 And Left(lsCtaHaber, 2) = "24" Then
                lsCtaHaber = Left(lsCtaHaber, 1) & "6" & Mid(lsCtaHaber, 3, Len(lsCtaHaber))
            End If
            'Comision
            lsCuentaComsionAdelantadaD = GetOpeCta(pcOpeCod, "D", "B", lsPersCod, lsIFTpo)
            lsCuentaComsionAdelantadaH = GetOpeCta(pcOpeCod, "H", "B", lsPersCod, lsIFTpo)
            'CP LP
            lsCuentaCorto = GetOpeCta(pcOpeCod, "D", "5", lsPersCod, lsIFTpo)
'            If lbCalculaCuenta Then
'                lsCuentaLargo = GetOpeCta(pcOpeCod, "H", "1", lsPersCod, lsIFTpo)
'            Else
'                lsCuentaLargo = "26" & Mid(lsCuentaCorto, 3, 50)
'            End If
            lsCuentaLargo = "26" & Mid(lsCuentaCorto, 3, 50)
            
            
            If pcMoneda = "1" And cMonedaPago = "2" Then
                'Cuenta VAC
                lsCuentaVac = GetOpeCta(pcOpeCod, "V", "2", lsPersCod, lsIFTpo) & lscCodLinCred
            End If
            
            lsMovNro = oDMov.GeneraMovNro(pdFechaProvision, psCodAge, psCodUser)
             
            
            ' ==================== VALORES REALES SIN VAC ====================
                        
            'Interes Provisionado REAL SIN VAC
            lnInteresReal = Val(lstListadoProvision.ListItems(I).SubItems(11))
            
            'SK Anterior REAL SIN VAC
            lnSKAntReal = Val(lstListadoProvision.ListItems(I).SubItems(12))
            
            'SK Largo Plazo REAL SIN VAC
            lnSKLPAntReal = Val(lstListadoProvision.ListItems(I).SubItems(20))
            
            'SK Largo Plazo Real Nuevo SIN VAC
            lnSKLPNewReal = CalculaSaldoAdeudadoLargoPlazo(GetFechaMov(lsMovNro, True), lsPersCod, lsIFTpo, lsCtaIFCod)
            
            ' ==================== VALORES INCLUIDO VAC ====================
            'Comision adelantada
            lnComisionAdelantada = Val(lstListadoProvision.ListItems(I).SubItems(25))
            'Interes Provisionado CON VAC
            lnInteres = Val(lstListadoProvision.ListItems(I).SubItems(10))
            
            'SK x Vac Anterior
            lnSKxVacAnt = Val(lstListadoProvision.ListItems(I).SubItems(19))
            
            'SK x Vac Nuevo
            lnSKxVacNew = Val(lstListadoProvision.ListItems(I).SubItems(4))
            
            'SK LP x Vac Anterior
            lnSKLPxVacAnt = Val(lstListadoProvision.ListItems(I).SubItems(21))
            
            'SK LP x Vac Nuevo
            lnSKLPxVacNew = Val(lstListadoProvision.ListItems(I).SubItems(22))
            
            ' ==================== GRABACION DE ASIENTOS ====================
            
            oDMov.InsertaMov lsMovNro, pcOpeCod, cGlosa, gMovEstContabMovContable, gMovFlagVigente
            lnMovNro = oDMov.GetnMovNro(lsMovNro)
            oDMov.InsertaMovCont lnMovNro, lnInteres, 0, ""
            Dim oCS As DCtaSaldo
            Set oCS = New DCtaSaldo
            Call oCS.RegistrarAdeudadoMov(lnMovNro, Mid(lstListadoProvision.ListItems(I).SubItems(14), 4, 13), Mid(lstListadoProvision.ListItems(I).SubItems(14), 1, 2), Mid(lstListadoProvision.ListItems(I).SubItems(14), 18, 7), nNroCuota)
            Set oCS = Nothing
            'Asiento de Provision
            item = item + 1
            oDMov.InsertaMovCta lnMovNro, item, lsCtaDebe, lnInteres
            oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
            oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod

            item = item + 1
            oDMov.InsertaMovCta lnMovNro, item, lsCtaHaber, lnInteres * -1
            oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
            oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
            If lnComisionAdelantada > 0 Then
            'lsCuentaComsionAdelantadaD
            item = item + 1
            oDMov.InsertaMovCta lnMovNro, item, lsCuentaComsionAdelantadaD, lnComisionAdelantada
            oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
            oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
            
            'lsCuentaComsionAdelantadaH
            item = item + 1
            oDMov.InsertaMovCta lnMovNro, item, lsCuentaComsionAdelantadaH, lnComisionAdelantada * -1
            oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
            oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
            End If
            'Traslado CP LP
'            If lnSKLPAntReal <> lnSKLPNewReal Then
'                item = item + 1
'                If pcMoneda = "1" And cMonedaPago = "2" Then
'                    oDMov.InsertaMovCta lnMovNro, item, lsCuentaCorto, Format((lnSKLPNewReal - lnSKLPAntReal) * lnVac, "0.00")
'                Else
'                    oDMov.InsertaMovCta lnMovNro, item, lsCuentaCorto, Format((lnSKLPNewReal - lnSKLPAntReal), "0.00")
'                End If
'                oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
'                oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                
'                item = item + 1
'                If pcMoneda = "1" And cMonedaPago = "2" Then
'                    oDMov.InsertaMovCta lnMovNro, item, lsCuentaLargo, Format((lnSKLPNewReal - lnSKLPAntReal) * -1 * lnVac, "0.00")
'                Else
'                    oDMov.InsertaMovCta lnMovNro, item, lsCuentaLargo, Format((lnSKLPNewReal - lnSKLPAntReal) * -1, "0.00")
'                End If
'                oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
'                oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
'            End If
            
            ' ==================== TRASLADO VAC ====================
            
            If pcMoneda = "1" And cMonedaPago = "2" Then

                'Si VAC subio/bajo
                lnMontoVacCP = Format(((lnSKxVacNew - lnSKxVacAnt) - (lnSKLPxVacNew - lnSKLPxVacAnt)), "0.00") * -1
                lnMontoVacLP = Format((lnSKLPxVacNew - lnSKLPxVacAnt), "0.00") * -1
                lnMontoVacTotal = Format((lnSKxVacNew - lnSKxVacAnt), "0.00")

                If lnMontoVacTotal - Abs(lnMontoVacCP) - Abs(lnMontoVacLP) <> 0 Then
                    If lnMontoVacTotal > 0 Then
                        lnMontoVacCP = Format(Abs(lnMontoVacTotal) - Abs(lnMontoVacLP), "0.00") * -1
                    ElseIf lnMontoVacTotal < 0 Then
                        lnMontoVacCP = Format(Abs(lnMontoVacTotal) - Abs(lnMontoVacLP), "0.00")
                    End If
                End If

                'Si VAC subio/bajo
                
                If lnMontoVacCP <> 0 Then
                    'Corto Plazo
                    item = item + 1
                    oDMov.InsertaMovCta lnMovNro, item, GetConvierteCuentaVac(lsCuentaCorto), Format(lnMontoVacCP, "0.00")
                    oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                    oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                End If
                
                If lnMontoVacLP <> 0 Then
                    'Largo Plazo
                    item = item + 1
                    oDMov.InsertaMovCta lnMovNro, item, GetConvierteCuentaVac(lsCuentaLargo), Format(lnMontoVacLP, "0.00")
                    oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                    oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                End If
                
                If lnMontoVacTotal <> 0 Then
                    'Cuenta Vac
                    item = item + 1
                    oDMov.InsertaMovCta lnMovNro, item, GetConvierteCuentaVac(lsCuentaVac), Format(lnMontoVacTotal, "0.00")
                    oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                    oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                End If
                
                If lnMontoVacCP <> 0 Or lnMontoVacLP <> 0 Or lnMontoVacTotal <> 0 Then
                    'Insertar Tipo de Cambio
                    oDMov.InsertaMovTpoCambio lnMovNro, lnVac
                End If
                
            End If
            
            ' ==================== ACTUALIZACION DE DATOS DE ADEUDADOS ====================
            
            oDMov.ActualizaCtaIF lsPersCod, lsIFTpo, lsCtaIFCod, , , , , , , , , lsMovNro
            oDMov.ActualizaDatosFAdeudado lsPersCod, lsIFTpo, lsCtaIFCod, False, "01/01/1900", True, pdFechaProvision, False, 0, False, 0
            'ALPA 20110815*************************
            'oDMov.ActualizaDatosCalendario lsPersCod, lsIFTpo, lsCtaIFCod, True, cTpoCuota, nNroCuota, False, "", 0, "2" 'lsMovNro
            
            oDMov.ActualizaDatosCalendario lsPersCod, lsIFTpo, lsCtaIFCod, True, cTpoCuota, nNroCuota, False, "", 0, "2", , , , , , CDbl(IIf(lstListadoProvision.ListItems(I).SubItems(27) = "", 0, lstListadoProvision.ListItems(I).SubItems(27))), CDbl(lnInteres) 'lsMovNro
            '**************************************
 
            ' ==================== GRABACION DE ESTADISTICAS ====================
            
            If pcMoneda = "1" And cMonedaPago = "2" Then
                'Si es VAC
                oDMov.GrabaCtaIFEstadistica lsPersCod, lsIFTpo, lsCtaIFCod, cTpoCuota, nNroCuota, lnMovNro, 4, _
                                            0, lnInteres, 0, 0, 0, Format((lnSKLPAntReal - lnSKLPNewReal) * lnVac, "0.00"), _
                                            Format(lnSKAntReal * lnVac, "0.00"), Format(lnSKAntReal * lnVac, "0.00"), _
                                            Format(lnSKLPAntReal * lnVac, "0.00"), Format(lnSKLPNewReal * lnVac, "0.00"), 1, _
                                            0, lnInteresReal, 0, 0, 0, Format((lnSKLPAntReal - lnSKLPNewReal), "0.00"), _
                                            lnSKAntReal, lnSKAntReal, lnSKLPAntReal, lnSKLPNewReal, _
                                            dFecUltActSaldos, pdFechaProvision, dFecUltPago, dFecUltPago, "0"
            Else
                'Si No Es Vac
                oDMov.GrabaCtaIFEstadistica lsPersCod, lsIFTpo, lsCtaIFCod, cTpoCuota, nNroCuota, lnMovNro, 4, _
                                            0, lnInteres, 0, 0, 0, Format((lnSKLPAntReal - lnSKLPNewReal), "0.00"), _
                                            lnSKAntReal, lnSKAntReal, _
                                            lnSKLPAntReal, lnSKLPNewReal, 0, _
                                            0, lnInteresReal, 0, 0, 0, Format((lnSKLPAntReal - lnSKLPNewReal), "0.00"), _
                                            lnSKAntReal, lnSKAntReal, lnSKLPAntReal, lnSKLPNewReal, _
                                            dFecUltActSaldos, pdFechaProvision, dFecUltPago, dFecUltPago, "0"
            End If
            
            If Mid(pcOpeCod, 3, 1) = "2" Then
               oDMov.GeneraMovME lnMovNro, lsMovNro
            End If
            oDMov.ActualizaSaldoMovimiento lsMovNro, "+"
              
            If Len(Trim(lsListaAsientos)) = 0 Then
                lsListaAsientos = "'" & lsMovNro & "'"
            Else
                lsListaAsientos = lsListaAsientos & ", '" & lsMovNro & "'"
            End If
            
            oDMov.CommitTrans
            lbTrans = False
            GrabaProvisionAdeudado = 2

        End If
    Next

GrabaProvisionAdeudado = 0
Set oDMov = Nothing

Exit Function
CGGrabaProvisionAdeudadoErr:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "CGGrabaProvisionAdeudado", lsMsgErr & " Error Graba Provision de Adeudados"
End Function



Public Function GrabaPagoAdeudado(ByVal pcOpeCod As String, ByVal pcMoneda As String, ByVal pdFechaPago As Date, ByVal lnVac As Double, _
                                    ByVal lstListadoPago As Variant, ByVal cGlosa As String, ByVal psCodAge As String, ByVal psCodUser As String, _
                                    ByRef lsListaAsientos As String, ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pdFechaDoc As Date, _
                                    ByVal psNroVoucher As String, ByVal cCodBco As String, ByVal bCancelacion As Boolean, _
                                    Optional ByVal bCanAdel As Boolean, Optional ByVal pnMontot As Currency, Optional ByRef sError As String = "", Optional ByRef pnMontoCapital As Currency = 0, Optional ByRef pnMontoInteres As Currency = 0, Optional ByRef pnMontoComision As Currency = 0, Optional ByRef pnMontoTotal As Currency = 0, Optional ByRef pnCambiarMonto As Integer = 0, _
                                    Optional ByVal pbConcesion As Boolean = False) As Integer
                                    
                                    
Dim oDMov As DMov

Dim lsPersCod As String
Dim lsIFTpo As String
Dim lsCtaIFCod As String

Dim lsPersCodBco As String
Dim lsIFTpoBco As String
Dim lscCtaIFCodBco As String

Dim lscCodLinCred As String

Dim lsCtaCapital As String
Dim lsCtaInteresProv As String
Dim lsCtaInteres As String
Dim lsCtaInteresGas As String
Dim lsCtaComision As String
Dim lsCtaBanco As String

Dim lsCuentaCorto As String
Dim lsCuentaLargo As String

Dim lsCuentaVac As String

Dim bExtranjero As Boolean
Dim lsCuentaExtranjeroD As String
Dim lsCuentaExtranjeroH As String
Dim lnTotalExtranjero As Double

Dim bRelacionarCredito As Boolean
Dim lsCuentaBuenPagadorD As String
Dim lsCuentaBuenPagadorH As String
Dim lnMontoBuenPagador As Double
Dim cCodCtaNeg As String

Dim cMonedaPago As String

Dim I As Integer
Dim item As Integer
Dim lnMovNro As Long

Dim lsMovNro As String
Dim lbTrans As Boolean

Dim cEstadoCuota As String
Dim cTpoCuota As String 'Tipo de Cuota
Dim nNroCuota As Integer 'Numero de Cuota
Dim nNroCuota6 As Integer 'Si hubiera cuota 6

Dim lnCapitalVac As Double
Dim lnCapitalReal As Double
Dim lnInteresProvVac As Double
Dim lnInteresProvReal As Double
Dim lnInteresVac As Double
Dim lnInteresGasVac As Double
Dim lnInteresGasReal As Double
Dim lnInteresReal As Double
Dim lnComisionVac As Double
Dim lnComisionReal As Double
Dim lnTotalVac As Double
Dim lnTotalReal As Double

Dim lnTotalInteresExtr As Double

Dim lnSKAntReal As Double 'SK Anterior REAL SIN VAC
Dim lnSKNewReal As Double 'SK que quedara REAL SIN VAC

Dim lnSKLPAntReal As Double 'SK Largo Plazo REAL SIN VAC
Dim lnSKLPNewReal As Double 'SK Largo Plazo REAL NUEVO SIN VAC
  
Dim lnSKxVacAnt As Double 'SK * VAC Anterior
Dim lnSKxVacNew As Double 'SK * VAC Nuevo

Dim lnSKLPxVacAnt As Double 'SK Largo Plazo * VAC Anterior
Dim lnSKLPxVacNew As Double 'SK Largo Plazo * VAC Nuevo

Dim dFecUltActSaldos As Date 'Fecha de Ultima Provision o Pago Efectuado Anterior
Dim dFecUltPago As Date 'Fecha de Ultimo Pago Efectuado Anterior

Dim lnMontoVacCP As Double
Dim lnMontoVacLP As Double
Dim lnMontoVacTotal As Double

'Nuevos Para PrePago
Dim lsCtaIFCodNew As String
Dim SKNew As Double
Dim SKNewLP As Double
Dim nMovNroNew As Long
Dim cMovNroNew As String
Dim oCaja As New nCajaGeneral
Dim lsMsgErr As String
Dim cDescripcionCuenta As String
Dim nSaldProvCLPlazo As Currency
Dim oCtaSaldo As DCtaSaldo
Dim nMontoCtaSaldo As Currency

'ALPA20130617********************************
Dim lsCuentaCapitalConcesionadoD As String
Dim lsCuentaInteresConcesionadoD As String
Dim lnInteresProvisionadoConcesi As Double
Dim lnCapitalProvisionadoConcesi As Double
Dim lnPagadoConcesion As Integer
'********************************************
'ALPA20130904********************************
Dim lnComisionConcesiado As Double
Dim lsCuentaComisionConcesionadoReal As String
'********************************************

On Error GoTo GrabaPagoAdeudadoErr
'
Set oDMov = New DMov
      
GrabaPagoAdeudado = 1

'Datos Banco
lsPersCodBco = Mid(Trim(cCodBco), 4, 13)
lsIFTpoBco = Mid(Trim(cCodBco), 1, 2)
lscCtaIFCodBco = Mid(Trim(cCodBco), 18, 7)

lsCtaBanco = GetOpeCta(pcOpeCod, "H", "0", lsPersCodBco, lsIFTpoBco, lscCtaIFCodBco)
                   
    For I = 1 To lstListadoPago.ListItems.Count
        
        If lstListadoPago.ListItems(I).Checked = True Then
                
            'Transaccion
            lbTrans = True
            oDMov.BeginTrans
            
            ' ==================== INICIALIZAMOS VALORES ====================
            
            lsPersCod = ""
            lsIFTpo = ""
            lsCtaIFCod = ""
            cMonedaPago = ""
            cTpoCuota = ""
            nNroCuota = 0
            nNroCuota6 = 0
            dFecUltActSaldos = "01/01/1900"
            dFecUltPago = "01/01/1900"
            lscCodLinCred = ""
            lsCtaCapital = ""
            lsCtaInteresProv = ""
            lsCtaInteres = ""
            lsCtaInteresGas = ""
            lsCtaComision = ""
            lsCuentaComisionConcesionadoReal = "" 'ALPA20130904
            'lsCtaBanco = ""
            lsCuentaCorto = ""
            lsCuentaLargo = ""
            lsCuentaVac = ""
            bExtranjero = False
            lsCuentaExtranjeroD = ""
            lsCuentaExtranjeroH = ""
            lnTotalExtranjero = 0
            bRelacionarCredito = False
            cCodCtaNeg = ""
            lsCuentaBuenPagadorD = ""
            lsCuentaBuenPagadorH = ""
            lnMontoBuenPagador = 0
            lsMovNro = ""
            lnCapitalReal = 0
            lnInteresProvReal = 0
            lnInteresReal = 0
            lnComisionReal = 0
            lnComisionConcesiado = 0 'ALPA20130904
            lnTotalReal = 0
            lnSKAntReal = 0
            lnSKNewReal = 0
            lnSKLPAntReal = 0
            lnSKLPNewReal = 0
            lnCapitalVac = 0
            lnInteresProvVac = 0
            lnInteresVac = 0
            lnComisionVac = 0
            lnTotalVac = 0
            lnTotalInteresExtr = 0
            lnSKxVacAnt = 0
            lnSKxVacNew = 0
            lnSKLPxVacAnt = 0
            lnSKLPxVacNew = 0
            
            lnMontoVacCP = 0
            lnMontoVacLP = 0
            lnMontoVacTotal = 0
             
            cDescripcionCuenta = ""
            
            lsCuentaCapitalConcesionadoD = ""
            lsCuentaInteresConcesionadoD = ""
            lnInteresProvisionadoConcesi = 0#
            lnPagadoConcesion = 0
             
            item = 0
             
            ' ==================== OBTENEMOS VALORES ====================
            lsPersCod = Mid(lstListadoPago.ListItems(I).SubItems(3), 4, 13)
            lsIFTpo = Mid(lstListadoPago.ListItems(I).SubItems(3), 1, 2)
            lsCtaIFCod = Mid(lstListadoPago.ListItems(I).SubItems(3), 18, 10)
            cMonedaPago = lstListadoPago.ListItems(I).SubItems(30)
            
            cDescripcionCuenta = Trim(lstListadoPago.ListItems(I).SubItems(4))
            
            'Estado Cuota
            cEstadoCuota = lstListadoPago.ListItems(I).SubItems(39)
            
            'Tipo de Cuota
            cTpoCuota = lstListadoPago.ListItems(I).SubItems(6)
            'Numero de Cuota
            nNroCuota = Val(lstListadoPago.ListItems(I).SubItems(7))
            nNroCuota6 = Val(lstListadoPago.ListItems(I).SubItems(33))

            'Fecha de Ultima Provision o Pago Efectuado Anterior
            dFecUltActSaldos = CDate(lstListadoPago.ListItems(I).SubItems(11))
            
            'Fecha de Ultimo Pago Efectuado Anterior
            dFecUltPago = CDate(lstListadoPago.ListItems(I).SubItems(10))
            
            'SubCuenta Contable asignada a la Linea de Credito
            lscCodLinCred = GetSubCtaLineaCredito(lstListadoPago.ListItems(I).SubItems(31))
              
            'Cuentas
            lsCtaCapital = GetOpeCta(pcOpeCod, "D", "0", lsPersCod, lsIFTpo)
            lsCtaInteresProv = GetOpeCta(pcOpeCod, "D", "1", lsPersCod, lsIFTpo)
            'lsCtaInteres = GetOpeCta(pcOpeCod, "D", "2", lsPersCod, lsIFTpo) & lscCodLinCred
            lsCtaInteres = GetOpeCta(pcOpeCod, "D", "2", lsPersCod, lsIFTpo) & lscCodLinCred 'SE QUITO LA SUBCTA DE LINEA POR ECONFIANZA
            lsCtaComision = GetOpeCta(pcOpeCod, "D", "3", lsPersCod, lsIFTpo)
            lsCuentaComisionConcesionadoReal = GetOpeCta(pcOpeCod, "D", "3", lsPersCod, lsIFTpo) 'ALPA20130904--14
            lsCtaInteresGas = GetOpeCta(pcOpeCod, "D", "A", lsPersCod, lsIFTpo) 'ALPA 20110704
            'lsCtaBanco = GetOpeCta(pcOpeCod, "D", "1", lsPersCod, lsIFTpo, lsCtaIFCod)

            'CP LP
            'lsCuentaCorto = GetOpeCta(pcOpeCod, "D", "5", lsPersCod, lsIFTpo)
            lsCuentaCorto = GetOpeCta(pcOpeCod, "H", "5", lsPersCod, lsIFTpo)
            'lsCuentaLargo = "26" & Mid(lsCuentaCorto, 3, 50)
            lsCuentaLargo = GetOpeCta(pcOpeCod, "H", "6", lsPersCod, lsIFTpo) 'EJVG20130204
            
            lsCuentaCapitalConcesionadoD = GetOpeCta(pcOpeCod, "D", "12", lsPersCod, lsIFTpo) 'ALPA20130204
            lsCuentaInteresConcesionadoD = GetOpeCta(pcOpeCod, "D", "13", lsPersCod, lsIFTpo) 'ALPA20130204
             
            If pcMoneda = "1" And cMonedaPago = "2" Then
                'Cuenta VAC
                lsCuentaVac = GetOpeCta(pcOpeCod, "V", "2", lsPersCod, lsIFTpo) & lscCodLinCred
            End If
           
            If lstListadoPago.ListItems(I).SubItems(21) = "13072" And nNroCuota6 > 0 Then
                'Cambio para Bono Buen Pagador
                lsCuentaBuenPagadorD = GetOpeCta(pcOpeCod, "D", "7", lsPersCod, lsIFTpo)
                lsCuentaBuenPagadorH = GetOpeCta(pcOpeCod, "H", "7", lsPersCod, lsIFTpo) & "01"
                lnMontoBuenPagador = CalculaMontoCuota(lsPersCod, lsIFTpo, lsCtaIFCod, "6", nNroCuota6, bCancelacion)
            End If

            lsMovNro = oDMov.GeneraMovNro(pdFechaPago, psCodAge, psCodUser)
             
            ' ==================== VALORES REALES SIN VAC ====================
             
            
            'Capital Real
            If pnCambiarMonto = 0 Then
                lnCapitalReal = Val(lstListadoPago.ListItems(I).SubItems(21))
                lnSKAntReal = Val(lstListadoPago.ListItems(I).SubItems(12))
                lnCapitalVac = Val(lstListadoPago.ListItems(I).SubItems(20))
            Else
                lnCapitalReal = pnMontoCapital
                lnSKAntReal = pnMontoCapital
                lnCapitalVac = pnMontoCapital
            End If
            
            'Interes Provisionado Real
            lnInteresProvReal = Val(lstListadoPago.ListItems(I).SubItems(19))
            
            'Interes Real
            lnInteresReal = Val(lstListadoPago.ListItems(I).SubItems(25))
            
            'Comision Real
            lnComisionReal = Val(lstListadoPago.ListItems(I).SubItems(27))
            
            'Comision Real Concesionado
            lnComisionConcesiado = Val(lstListadoPago.ListItems(I).SubItems(52)) 'ALPA20130904
            
            'Total Real
            lnTotalReal = Val(lstListadoPago.ListItems(I).SubItems(29))
                        
            'SK Anterior REAL SIN VAC
'            lnSKAntReal = Val(lstListadoPago.ListItems(I).SubItems(12))
             
            'SK Nuevo REAL SIN VAC
            lnSKNewReal = lnSKAntReal - lnCapitalReal
            
            'SK Largo Plazo Actual REAL SIN VAC
            lnSKLPAntReal = Val(lstListadoPago.ListItems(I).SubItems(15))
            
            'SK Largo Plazo Real Nuevo SIN VAC
            If bCancelacion = True Then
                lnSKLPNewReal = 0
            Else
                lnSKLPNewReal = oCaja.CalculaSaldoAdeudadoLargoPlazo(GetFechaMov(lsMovNro, True), lsPersCod, lsIFTpo, lsCtaIFCod)
            End If
             
            ' ==================== VALORES INCLUIDO VAC ====================
            
            'Capital Vac
'            lnCapitalVac = Val(lstListadoPago.ListItems(I).SubItems(20))
            
            'Interes Provisionado VAC
            'lnInteresProvVac = Val(lstListadoPago.ListItems(I).SubItems(18))
            lnInteresProvVac = Val(lstListadoPago.ListItems(I).SubItems(47))
            
            'Por Cancelacion Adelantada GITU
            'Interes VAC
            Dim nMontoProvision As Currency
            Set oCtaSaldo = New DCtaSaldo
            'lsCtaInteresGas
            
            If pnCambiarMonto = 0 Then
            nMontoCtaSaldo = oCtaSaldo.ObtenerDatosCtaSaldoMaxFecha(lsCtaInteresProv, pdFechaPago, 0, lstListadoPago.ListItems(I).SubItems(44), lstListadoPago.ListItems(I).SubItems(45), lstListadoPago.ListItems(I).SubItems(46), lstListadoPago.ListItems(I).SubItems(7))
            nMontoProvision = Val(lstListadoPago.ListItems(I).SubItems(47))
            If bCanAdel Then
                lnInteresVac = pnMontot - Val(lstListadoPago.ListItems(I).SubItems(20))
            Else
                'lnInteresVac = Val(lstListadoPago.ListItems(I).SubItems(24))
                'If (Val(lstListadoPago.ListItems(I).SubItems(24)) - nMontoCtaSaldo) > 0 And nMontoCtaSaldo > 0 Then
                If (lnVac * Val(lstListadoPago.ListItems(I).SubItems(24))) > 0 Then
'                    lnInteresGasVac = nMontoCtaSaldo
'                    lnInteresVac = Val(lstListadoPago.ListItems(I).SubItems(24)) - nMontoCtaSaldo
                    lnInteresGasVac = lnVac * Val(lstListadoPago.ListItems(I).SubItems(24))
                    lnInteresVac = Round(Val(lstListadoPago.ListItems(I).SubItems(23)) - Val(lstListadoPago.ListItems(I).SubItems(47)) - lnInteresGasVac, 2)
                Else
                    'lnInteresVac = Val(lstListadoPago.ListItems(I).SubItems(24))
                    lnInteresVac = Round(Val(lstListadoPago.ListItems(I).SubItems(23)) - Val(lstListadoPago.ListItems(I).SubItems(47)), 2)
                    lnInteresGasVac = 0
                End If
            End If
            Else
                lnInteresVac = pnMontoInteres
                lnInteresGasVac = 0
            End If
            
            'Comision VAC
            If pnCambiarMonto = 0 Then
                lnComisionVac = Val(lstListadoPago.ListItems(I).SubItems(26))
            Else
                lnComisionVac = pnMontoComision
            End If
                        
            'Total VAC GITU
            'If IsNumeric(lstListadoPago.ListItems(I).SubItems(42)) Then
            '    If bCandel Then
            '        lnTotalVac = pnMontot
            '    Else
            '        lnTotalVac = Val(lstListadoPago.ListItems(I).SubItems(28)) '+ Val(lstListadoPago.ListItems(i).SubItems(42))
            '    End If
            'Else
            If pnCambiarMonto = 0 Then
                If bCanAdel Then
                    lnTotalVac = pnMontot
                Else
                    lnTotalVac = Val(lstListadoPago.ListItems(I).SubItems(28))
                End If
            Else
                lnTotalVac = pnMontoTotal
            End If
            'End If
            
            'SK x Vac Anterior
            lnSKxVacAnt = Val(lstListadoPago.ListItems(I).SubItems(13))
            
            'SK x Vac Nuevo
            lnSKxVacNew = Val(lstListadoPago.ListItems(I).SubItems(14))
            
            'SK LP x Vac Anterior
            lnSKLPxVacAnt = Val(lstListadoPago.ListItems(I).SubItems(16))
            
            'SK LP x Vac Nuevo
            lnSKLPxVacNew = Val(lstListadoPago.ListItems(I).SubItems(17))
            
            'K Concesionado
            lnCapitalProvisionadoConcesi = Val(lstListadoPago.ListItems(I).SubItems(48))
            'I Concesionado
            lnInteresProvisionadoConcesi = Val(lstListadoPago.ListItems(I).SubItems(51))
           
            oDMov.InsertaMov lsMovNro, pcOpeCod, " Pagaré " & cDescripcionCuenta & " " & cGlosa, gMovEstContabMovContable, gMovFlagVigente

            
            lnMovNro = oDMov.GetnMovNro(lsMovNro)
            oDMov.InsertaMovCont lnMovNro, lnTotalVac, 0, ""

            'Asiento de Pago
            If lnCapitalVac <> 0 Then
                'ALPA 20110831
                If Trim(lsCtaCapital) = "" Then
                    sError = "Cuenta Capital no esta registrado (D (Capital))"
                    If lbTrans Then
                        oDMov.RollbackTrans
                        lbTrans = False
                    End If
                    Exit Function
                End If
                If Trim(lsCuentaCapitalConcesionadoD) = "" And pbConcesion = True Then
                    sError = "Cuenta Capital Concesionado no esta registrado (D (Capital Concesionado))"
                    If lbTrans Then
                        oDMov.RollbackTrans
                        lbTrans = False
                    End If
                    Exit Function
                End If
                
                If pbConcesion = False Then
                    item = item + 1
                    oDMov.InsertaMovCta lnMovNro, item, lsCtaCapital, lnCapitalVac
                    oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                    oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                    lnPagadoConcesion = 0
                Else
                    item = item + 1
                    oDMov.InsertaMovCta lnMovNro, item, lsCtaCapital, lnCapitalVac - lnCapitalProvisionadoConcesi
                    oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                    oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                    
                    item = item + 1
                    oDMov.InsertaMovCta lnMovNro, item, lsCuentaCapitalConcesionadoD, lnCapitalProvisionadoConcesi
                    oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                    oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                    lnPagadoConcesion = 1

                End If
                
                If pcMoneda = "1" And cMonedaPago = "2" Then
                    Dim lnCapitalVAC3 As Currency
                    lnCapitalVAC3 = Val(lstListadoPago.ListItems(I).SubItems(42))
                    'lnCapitalVAC3 = Format((lnCapitalReal * lstListadoPago.ListItems(i).SubItems(41)) - (lnCapitalReal * lnVac), "0.00")
                    
                    'item = item + 1
                    'oDMov.InsertaMovCta lnMovNro, item, GetConvierteCuentaVac(lsCtaCapital), lnCapitalVAC3 * -1
                    'oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                    'oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                
                    'item = item + 1
                    'oDMov.InsertaMovCta lnMovNro, item, lsCuentaVac, lnCapitalVAC3
                    'oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                    'oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                
                    item = item + 1
                    oDMov.InsertaMovCta lnMovNro, item, GetConvierteCuentaVac(lsCtaCapital), lnCapitalVAC3
                    oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                    oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                End If
                
            End If
            If pnCambiarMonto = 0 Then
            If lnInteresProvVac <> 0 Then
                'ALPA 20110831
                If Trim(lsCtaInteresProv) = "" Then
                    sError = "Cuenta Provisión de Interes no esta registrado (D (Provisión de Interes))"
                    If lbTrans Then
                        oDMov.RollbackTrans
                        lbTrans = False
                    End If
                    Exit Function
                End If
                
               If Trim(lsCuentaInteresConcesionadoD) = "" And pbConcesion = True Then
                    sError = "Cuenta Interes Concesionado no esta registrado (D (Interes Concesionado))"
                    If lbTrans Then
                        oDMov.RollbackTrans
                        lbTrans = False
                    End If
                    Exit Function
                End If

                If pbConcesion = False Then
                    item = item + 1
                    oDMov.InsertaMovCta lnMovNro, item, lsCtaInteresProv, lnInteresProvVac
                    oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                    oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                Else
                    item = item + 1
                    oDMov.InsertaMovCta lnMovNro, item, lsCtaInteresProv, lnInteresProvVac - lnInteresProvisionadoConcesi
                    oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                    oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                    
                    item = item + 1
                    oDMov.InsertaMovCta lnMovNro, item, lsCuentaInteresConcesionadoD, lnInteresProvisionadoConcesi
                    oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                    oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                End If
            End If
            End If
            If lnInteresVac <> 0 Then
                'ALPA 20110831
                If Trim(lsCtaInteres) = "" Then
                    sError = "Cuenta de Interes no esta registrado (D (Interes))"
                    If lbTrans Then
                        oDMov.RollbackTrans
                        lbTrans = False
                    End If
                    Exit Function
                End If

                item = item + 1
                oDMov.InsertaMovCta lnMovNro, item, lsCtaInteres, lnInteresVac
                oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
            End If
            If lnInteresGasVac <> 0 Then
                'ALPA 20110831
                If Trim(lsCtaInteresGas) = "" Then
                    sError = "Cuenta de Gastos de Inteses no esta registrado (D (Gastos de Inteses))"
                    If lbTrans Then
                        oDMov.RollbackTrans
                        lbTrans = False
                    End If
                    Exit Function
                End If
                item = item + 1
                oDMov.InsertaMovCta lnMovNro, item, lsCtaInteresGas, lnInteresGasVac
                oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod

            End If
            If lnComisionVac <> 0 Then
                item = item + 1
                If Trim(lsCtaComision) = "" Then
                    sError = "Cuenta de Comisión no esta registrado (D (Comisión))"
                    If lbTrans Then
                        oDMov.RollbackTrans
                        lbTrans = False
                    End If
                    Exit Function
                End If

                oDMov.InsertaMovCta lnMovNro, item, lsCtaComision, lnComisionVac
                oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
            End If
            
            'ALPA20130904********************************
            If lnComisionConcesiado <> 0 And pnCambiarMonto = 0 Then 'NAGL Agregó pnCambiarMonto = 0 Según Inc 20181017
                item = item + 1
                If Trim(lsCtaComision) = "" Then
                    sError = "Cuenta de Comisión no esta registrado (D (Comisión))"
                    If lbTrans Then
                        oDMov.RollbackTrans
                        lbTrans = False
                    End If
                    Exit Function
                End If

                oDMov.InsertaMovCta lnMovNro, item, lsCuentaComisionConcesionadoReal, lnComisionConcesiado
                oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
            End If
            '********************************************
            
            'Pago del Banco
            If Trim(lsCtaBanco) = "" Then
                sError = "Cuenta de Banco no está registrado (H (Capital))"
                If lbTrans Then
                    oDMov.RollbackTrans
                    lbTrans = False
                End If
                Exit Function
            End If
            item = item + 1
            oDMov.InsertaMovCta lnMovNro, item, lsCtaBanco, lnTotalVac * -1
            oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
            oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCodBco, lsIFTpoBco, lscCtaIFCodBco
                        
            'Traslado CP LP
            If lnSKLPAntReal <> lnSKLPNewReal Then
            If pnCambiarMonto = 0 Then
                If Trim(lsCuentaCorto) = "" Then
                    sError = "Cuenta Corto no está registrado (H (Plazo))"
                    If lbTrans Then
                        oDMov.RollbackTrans
                        lbTrans = False
                    End If
                    Exit Function
                End If

                item = item + 1
                If pcMoneda = "1" And cMonedaPago = "2" Then
                    oDMov.InsertaMovCta lnMovNro, item, lsCuentaCorto, Format((lnSKLPNewReal - lnSKLPAntReal) * lstListadoPago.ListItems(I).SubItems(41), "0.00")
                Else
                    oDMov.InsertaMovCta lnMovNro, item, lsCuentaCorto, Format((lnSKLPNewReal - lnSKLPAntReal), "0.00")
                End If
                oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                
                If Trim(lsCuentaLargo) = "" Then
                    sError = "Cuenta Largo no está registrado (H (Plazo))"
                    If lbTrans Then
                        oDMov.RollbackTrans
                        lbTrans = False
                    End If
                    Exit Function
                End If
                
                item = item + 1
                If pcMoneda = "1" And cMonedaPago = "2" Then
                    oDMov.InsertaMovCta lnMovNro, item, lsCuentaLargo, Format((lnSKLPNewReal - lnSKLPAntReal) * -1 * lstListadoPago.ListItems(I).SubItems(41), "0.00")
                Else
                    oDMov.InsertaMovCta lnMovNro, item, lsCuentaLargo, Format((lnSKLPNewReal - lnSKLPAntReal) * -1, "0.00")
                End If
                oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
             End If
            End If
            
            ' ==================== ASIENTO DE TRASLADO VAC ====================
            
            If pcMoneda = "1" And cMonedaPago = "2" Then
                    
                'Si VAC subio/bajo
                lnMontoVacCP = Format(((lnSKxVacNew - lnSKxVacAnt) - (lnSKLPxVacNew - lnSKLPxVacAnt)), "0.00") * -1
                lnMontoVacLP = Format((lnSKLPxVacNew - lnSKLPxVacAnt), "0.00") * -1
                lnMontoVacTotal = Format((lnSKxVacNew - lnSKxVacAnt), "0.00")
                
                If lnMontoVacTotal - Abs(lnMontoVacCP) - Abs(lnMontoVacLP) <> 0 Then
                    If lnMontoVacTotal > 0 Then
                        lnMontoVacCP = Format(Abs(lnMontoVacTotal) - Abs(lnMontoVacLP), "0.00") * -1
                    ElseIf lnMontoVacTotal < 0 Then
                        lnMontoVacCP = Format(Abs(lnMontoVacTotal) - Abs(lnMontoVacLP), "0.00")
                    End If
                End If
                
                If lnMontoVacCP <> 0 Then
                    'Corto Plazo
                    item = item + 1
                    oDMov.InsertaMovCta lnMovNro, item, GetConvierteCuentaVac(lsCuentaCorto), Format(lnMontoVacCP, "0.00")
                    oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                    oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                End If
                
                If lnMontoVacLP <> 0 Then
            
                    'Largo Plazo
                    item = item + 1
                    oDMov.InsertaMovCta lnMovNro, item, GetConvierteCuentaVac(lsCuentaLargo), Format(lnMontoVacLP, "0.00")
                    oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                    oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                End If
                
                If lnMontoVacTotal <> 0 Then
                    'Cuenta Vac
                    item = item + 1
                    'modificacion GCZA
                    oDMov.InsertaMovCta lnMovNro, item, GetConvierteCuentaVac(lsCuentaVac), Format(lnMontoVacTotal, "0.00")
                    oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                    oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                End If
                   
                If lnMontoVacCP <> 0 Or lnMontoVacLP <> 0 Or lnMontoVacTotal <> 0 Then
                    'Insertar Tipo de Cambio
                    oDMov.InsertaMovTpoCambio lnMovNro, lnVac
                End If
                
            End If
            
          
            ' ==================== MOVIMIENTO DE BONO BUEN PAGADOR ====================
            
            If lstListadoPago.ListItems(I).SubItems(21) = "13072" And nNroCuota6 > 0 Then
                If Trim(lsCuentaBuenPagadorD) = "" Then
                    sError = "Cuenta Buen pagador no está registrado (D (Buen Pagador))"
                    If lbTrans Then
                        oDMov.RollbackTrans
                        lbTrans = False
                    End If
                    Exit Function
                End If
                
                item = item + 1
                oDMov.InsertaMovCta lnMovNro, item, lsCuentaBuenPagadorD, lnMontoBuenPagador
                oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                
                If Trim(lsCuentaBuenPagadorH) = "" Then
                    sError = "Cuenta Buen pagador no está registrado (H (Buen Pagador))"
                    If lbTrans Then
                        oDMov.RollbackTrans
                        lbTrans = False
                    End If
                    Exit Function
                End If

                
                item = item + 1
                oDMov.InsertaMovCta lnMovNro, item, lsCuentaBuenPagadorH, lnMontoBuenPagador * -1
                oDMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
                oDMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
                
            End If
            
            
            ' ==================== ACTUALIZACION DE DATOS DE ADEUDADOS ====================
            If bCancelacion = True Then
                oDMov.ActualizaCtaIF lsPersCod, lsIFTpo, lsCtaIFCod, , , , , , , "2", , lsMovNro
            Else
                oDMov.ActualizaCtaIF lsPersCod, lsIFTpo, lsCtaIFCod, , , , , , , "1", , lsMovNro
            End If
            
'            If nNroCuota6 = 0 Then
'                oDMov.ActualizaCtaIF lsPersCod, lsIFTpo, lsCtaifCod, , , , , , , , , lsMovNro
'                oDMov.ActualizaAdeudadosPagoCuota lsPersCod, lsIFTpo, lsCtaifCod, CStr(pdFechaPago), CCur(lnSKNewReal), CCur(lnInteresReal), nNroCuota, lsMovNro, bCancelacion
'            ElseIf nNroCuota6 > 0 Then
'                oDMov.ActualizaCtaIF lsPersCod, lsIFTpo, lsCtaifCod, , , , , , , , , lsMovNro
'                oDMov.ActualizaAdeudadosPagoCuota lsPersCod, lsIFTpo, lsCtaifCod, CStr(pdFechaPago), CCur(lnSKNewReal), CCur(lnInteresReal), nNroCuota, lsMovNro, bCancelacion, CStr(nNroCuota6)
'            End If
            If nNroCuota6 = 0 Then
                oDMov.ActualizaDatosFAdeudado lsPersCod, lsIFTpo, lsCtaIFCod, True, pdFechaPago, True, pdFechaPago, True, lnSKNewReal, True, lnSKLPNewReal, bCancelacion, pdFechaPago, , False
                oDMov.ActualizaDatosCalendario lsPersCod, lsIFTpo, lsCtaIFCod, True, cTpoCuota, nNroCuota, False, "0", 0, "1", bCancelacion, , False, , lsMovNro, , 0, Val(lstListadoPago.ListItems(I).SubItems(47)), , lnPagadoConcesion
            
            ElseIf nNroCuota6 > 0 Then
                oDMov.ActualizaDatosFAdeudado lsPersCod, lsIFTpo, lsCtaIFCod, True, pdFechaPago, True, pdFechaPago, True, lnSKNewReal, True, lnSKLPNewReal, bCancelacion, pdFechaPago, , True
                oDMov.ActualizaDatosCalendario lsPersCod, lsIFTpo, lsCtaIFCod, True, cTpoCuota, nNroCuota, True, "6", nNroCuota6, "1", bCancelacion, , False, , lsMovNro, , 0, Val(lstListadoPago.ListItems(I).SubItems(13)), , lnPagadoConcesion
            End If

                
            ' ==================== GRABACION DE DOCUMENTOS ====================
            If psNroDoc <> "" Then
                oDMov.InsertaMovDoc lnMovNro, pnTpoDoc, psNroDoc, Format(pdFechaDoc, gsFormatoFecha)
            End If
            If psNroVoucher <> "" Then
                oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, Format(pdFechaDoc, gsFormatoFecha)
            End If

            ' ==================== GRABACION DE ESTADISTICAS ====================
             
            If pcMoneda = "1" And cMonedaPago = "2" Then
                'Si es VAC
                oDMov.GrabaCtaIFEstadistica lsPersCod, lsIFTpo, lsCtaIFCod, cTpoCuota, nNroCuota, lnMovNro, 3, _
                                            lnCapitalVac, lnInteresProvVac, lnInteresVac, lnComisionVac, 0, Format((lnSKLPAntReal - lnSKLPNewReal) * lnVac, "0.00"), _
                                            Format(lnSKxVacAnt, "0.00"), Format(lnSKNewReal * lnVac, "0.00"), _
                                            Format(lnSKLPxVacAnt, "0.00"), Format(lnSKLPNewReal * lnVac, "0.00"), 1, _
                                            lnCapitalReal, lnInteresProvReal, lnInteresReal, lnComisionReal, 0, Format((lnSKLPAntReal - lnSKLPNewReal), "0.00"), _
                                            lnSKAntReal, lnSKNewReal, lnSKLPAntReal, lnSKLPNewReal, _
                                            dFecUltActSaldos, pdFechaPago, dFecUltPago, pdFechaPago, cEstadoCuota
            Else
                'Si No Es Vac
                oDMov.GrabaCtaIFEstadistica lsPersCod, lsIFTpo, lsCtaIFCod, cTpoCuota, nNroCuota, lnMovNro, 3, _
                                            lnCapitalReal, lnInteresProvReal, lnInteresReal, lnComisionReal, 0, Format((lnSKLPAntReal - lnSKLPNewReal), "0.00"), _
                                            lnSKAntReal, lnSKNewReal, _
                                            lnSKLPAntReal, lnSKLPNewReal, 0, _
                                            lnCapitalReal, lnInteresProvReal, lnInteresReal, lnComisionReal, 0, Format((lnSKLPAntReal - lnSKLPNewReal), "0.00"), _
                                            lnSKAntReal, lnSKNewReal, lnSKLPAntReal, lnSKLPNewReal, _
                                            dFecUltActSaldos, pdFechaPago, dFecUltPago, pdFechaPago, cEstadoCuota
            End If
           
           
           
            If Mid(pcOpeCod, 3, 1) = "2" Then
                oDMov.GeneraMovME lnMovNro, lsMovNro
            End If
            
            oDMov.ActualizaSaldoMovimiento lsMovNro, "+"
            
            If Len(Trim(lsListaAsientos)) = 0 Then
                lsListaAsientos = "'" & lsMovNro & "'"
            Else
                lsListaAsientos = lsListaAsientos & ", '" & lsMovNro & "' "
            End If
            oDMov.CommitTrans
            lbTrans = False
            
            GrabaPagoAdeudado = 2
            
        End If
    Next

'

GrabaPagoAdeudado = 0
Set oDMov = Nothing

Exit Function
GrabaPagoAdeudadoErr:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaPagoAdeudado", lsMsgErr & " Error Graba Pago de Adeudados"
                                    
End Function


Public Function GetDatosAdeudadoPago(pdFecha As Date, pnMoneda As Moneda, Optional psBuscar As String = "", Optional bCancelacion As Boolean = False, Optional bRestringirFecha As Boolean = False, Optional pbControlFechaRep As Boolean = False) As ADODB.Recordset

    On Error GoTo GetDatosAdeudadoPagoErr
    Dim oCon As New DConecta
    
    'PEAC 20210722 - Se envio este scritp aun Store Procesure
'    sSql = "Select CIF.cPersCod, P.cPersNombre, CIF.cIFTpo, CIF.cCtaIFCod, CIF.cCtaIFDesc, "
'
'    If bCancelacion = True Then
'        sSql = sSql & " CICUltimo.dVencimiento, CalRef.nNroCuota, '2' as cTpoCuota, "
'        sSql = sSql & " CICUltimo.cEstado as cEstadoCuota, "
'    ElseIf bCancelacion = False Then
'        sSql = sSql & " CIC.dVencimiento, CIC.nNroCuota, CIC.cTpoCuota, "
'        sSql = sSql & " CIC.cEstado as cEstadoCuota, "
'    End If
'
'    sSql = sSql & " CII.nCtaIFIntPeriodo, CII.nCtaIFIntValor, CIA.cMonedaPago, CIA.nMontoPrestado, "
'    sSql = sSql & " ISNULL(CIE.nSumaInteresMovido, 0) as nInteresProvisionado, "
'    sSql = sSql & " ISNULL(CIE.nSumaInteresMovidoReal, 0) as nInteresProvisionadoReal, "
'    sSql = sSql & " CIA.dCuotaUltModSaldos, CIA.dCuotaUltPago, "
'
'    If bCancelacion = True Then
'        sSql = sSql & " CIA.nSaldoCap, CIA.nSaldoCapLP, CIC.nCapital, IsNull(CIC.nInteres,0) + IsNull(CIC.nInteresGracia,0) As nInteres, CIC.nComision, CIC.nTotalCuota, "
'    Else
'        sSql = sSql & " CIA.nSaldoCap, CIA.nSaldoCapLP, CIC.nCapital + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,1) nCapital, IsNull(CIC.nInteres,0) +  dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,2) + IsNull(CIC.nInteresGracia,0) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,3) As nInteres, CIC.nComision + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,4) nComision, CIC.nTotalCuota + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,1) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,2) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,3) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,4) nTotalCuota, "
'    End If
'
'    sSql = sSql & " ISNULL(CIA.cCodLinCred, '') as cCodLinCred, ISNULL(CLC.cDescripcion, '') as cDesLinCred , "
'    sSql = sSql & " iVacAper=(select nIndiceVac from indiceVac where dIndiceVAC=CIF.dctaifaper), (Select max(dVencimiento) from CtaIfCalendario Calen where Calen.cPersCod = CIF.cPersCod And Calen.cIFTpo = CIF.cIFTpo And Calen.cCtaIfcod = CIF.cCtaIFCod And bVigente = 1) dVencimientoFinal, (Select dCtaIFAper from CtaIf where CtaIf.cPersCod = CIF.cPersCod And CtaIf.cIFTpo = CIF.cIFTpo And CtaIf.cCtaIfcod = CIF.cCtaIFCod) dCtaIFAper,CIF.cPersCod,CIF.cIFTpo,CIF.cCtaIFCod, "
'    sSql = sSql & " IsNull(CalRef.nSaldoMes,0) nSaldoMes, "
'    'ALPA20130617******
'    sSql = sSql & " nCapitalConce=IsNull(CIC.nCapitalConce,0),"
'    sSql = sSql & " nInteresConce=IsNull(CIC.nInteresConce,0), "
'    sSql = sSql & " nSaldoCapConce=IsNull(CIA.nSaldoCapConce,0), "
'    sSql = sSql & " nProvisionConce= IsNull(CIC.nProvisionConce,0) "
'    '******************
'    'ALPA20130904******
'    sSql = sSql & " ,nComisionConce= IsNull(CIC.nComisionConce,0) "
'    '******************
'    sSql = sSql & " From CtaIF CIF "
'    sSql = sSql & " Inner Join CtaIFAdeudados CIA "
'    sSql = sSql & "         On CIF.cPersCod=CIA.cPersCod And CIF.cIFTpo=CIA.cIFTpo And CIF.cCtaIFCod=CIA.cCtaIFCod "
'
'    If pbControlFechaRep Then
'        sSql = sSql & " Inner Join CtaIF CtaIF "
'        sSql = sSql & "         On CIF.cPersCod = CtaIF.cPersCod And CIF.cIFTpo = CtaIF.cIFTpo And CIF.cCtaIFCod = CtaIF.cCtaIFCod And CtaIF.dCtaIfAper <= '" & Format(pdFecha, gsFormatoFecha) & "' "
'    End If
'
'    'Proxima Cuota Pendiente
'    sSql = sSql & " Inner Join (    SELECT cPersCod, cIFTpo, cCtaIFCod, MIN(nNroCuota) nNroCuota,nSaldoMes "
'    sSql = sSql & "                 From CtaIFCalendario "
'    sSql = sSql & "                 WHERE bVigente=1 And (cEstado = '0' or cEstado='2') and cTpoCuota <> '6' "
'    sSql = sSql & "                     And Substring(cCtaIFCod, 3,1)='" & Trim(Str(pnMoneda)) & "' " '& Replace(psBuscar, "CIF.", "")
'    sSql = sSql & "                 GROUP BY cPersCod, cIFTpo, cCtaIFCod,nSaldoMes "
'    sSql = sSql & "             ) CalRef ON CIF.cPersCod=CalRef.cPersCod And CIF.cIFTpo=CalRef.cIFTpo And CIF.cCtaIFCod=CalRef.cCtaIFCod "
'
'    If bCancelacion = True Then
'
'        'Sacamos todo el capital, intereses y comisiones pendientes Pero de lo no Concesional,
'        'Porque lo concesional se PERDONA!!!
'
'        sSql = sSql & " Inner Join ( "
'        sSql = sSql & " select  CIK.cPersCod, CIK.cIFTpo, CIK.cCtaIFCod, Sum(IsNull(CIK.nCapital,0)) as nCapital, Sum(IsNull(CIK.nInteres,0)) AS nInteres, Sum(IsNull(CIK.nInteresGracia,0)) as nInteresGracia, "
'        sSql = sSql & " sum(IsNull(CIK.nComision,0)) As nComision, Sum(IsNull(CIK.nCapital,0) + IsNull(CIK.nInteres,0) + IsNull(CIK.nInteresGracia,0) + IsNull(CIK.nComision,0)) as nTotalCuota, " 'Se agregó la "," para la continuación en lo descrito líneas Abajo
'        sSql = sSql & " Sum(IsNull(CIK.nCapitalConce,0)) as nCapitalConce, Sum(IsNull(CIK.nInteresConce,0)) AS nInteresConce, Sum(IsNull(CIK.nProvisionConce,0)) as nProvisionConce, Sum(IsNull(CIK.nComisionConce,0)) as nComisionConce" 'Agregado by NAGL 20180613 Según INC1806120013
'        sSql = sSql & " from ctaifcalendario CIK "
'        sSql = sSql & " Where CIK.bVigente=1 And CIK.cEstado in('0', '2')  " 'And CIK.cTpoCuota='2'
'        sSql = sSql & " Group By CIK.cPersCod, CIK.cIFTpo, CIK.cCtaIFCod "
'        sSql = sSql & " ) CIC On CIF.cPersCod=CIC.cPersCod And CIF.cIFTpo=CIC.cIFTpo And CIF.cCtaIFCod=CIC.cCtaIFCod "
'
'        'Sacamos el estado de la ultima cuota pendiente
'        sSql = sSql & " Inner Join ( "
'        sSql = sSql & " Select CIK.cPersCod, CIK.cIFTpo, CIK.cCtaIFCod, CIK.nNroCuota, CIK.cEstado, CIK.dVencimiento "
'        sSql = sSql & " From CtaIFCalendario CIK "
'        sSql = sSql & " Where CIK.bVigente=1 And CIK.cEstado in('0', '2') "
'        sSql = sSql & " And CIK.cTpoCuota='2' "
'        sSql = sSql & " ) CICUltimo On CIF.cPersCod=CICUltimo.cPersCod And CIF.cIFTpo=CICUltimo.cIFTpo And CIF.cCtaIFCod=CICUltimo.cCtaIFCod "
'        sSql = sSql & " And CalRef.nNroCuota=CICUltimo.nNroCuota "
'
'    ElseIf bCancelacion = False Then
'
'        'Sacamos el capital, intereses y comisiones de la prox cuota pendiente
'        sSql = sSql & " Inner Join CtaIFCalendario CIC "
'        sSql = sSql & "         ON CIF.cPersCod=CIC.cPersCod And CIF.cIFTpo=CIC.cIFTpo And CIF.cCtaIFCod=CIC.cCtaIFCod "
'        sSql = sSql & "                 AND CIC.nNroCuota=CalRef.nNroCuota "
'
'    End If
'
'    sSql = sSql & " Inner Join CtaIFInteres CII "
'    sSql = sSql & "         On CIF.cPersCod=CII.cPersCod And CIF.cIFTpo=CII.cIFTpo And CIF.cCtaIFCod=CII.cCtaIFCod "
'    sSql = sSql & " Inner JOIN Persona P "
'    sSql = sSql & "         ON CIF.cPersCod = P.cPersCod "
'
'    'Ultimo Interes provisionado
'    sSql = sSql & " Left Join ( "
'    sSql = sSql & "       Select CIE.cPersCod, CIE.cIFTpo, CIE.cCtaIFCod, " & IIf(bCancelacion = True, "", " CIE.nNroCuota,")
'    sSql = sSql & "           SUM(CIE.nInteres1Movido) as nSumaInteresMovido, SUM(CIED.nInteres1MovidoReal) as nSumaInteresMovidoReal "
'    sSql = sSql & "       From CtaIFEstadistica CIE "
'    sSql = sSql & "       Inner Join CtaIFEstadisticaDet CIED "
'    sSql = sSql & "                       On CIE.nMovNro = CIED.nMovNro "
'    sSql = sSql & "       Inner Join Mov m on   cie.nmovnro=m.nmovnro and m.nmovflag<>1 " & IIf(pbControlFechaRep, " And cMovNro < '" & Format(pdFecha + 1, gsFormatoMovFecha) & "' ", "")
'    sSql = sSql & "       Left Join ( Select cPersCod,cIFTpo,cCtaIFCod,max(nNroCuota) nNroCuota"
'    sSql = sSql & "                from CtaIFEstadistica a"
'    sSql = sSql & "                Inner Join Mov b on a.nmovnro=b.nmovnro and b.nmovflag<>1 " & IIf(pbControlFechaRep, " And cMovNro < '" & Format(pdFecha + 1, gsFormatoMovFecha) & "' ", "")
'    sSql = sSql & "                Where nTipoEstadistica = 3"
'    sSql = sSql & "                group by cPersCod,cIFTpo,cCtaIFCod"
'    sSql = sSql & "           ) X   on x.cperscod=CIE.cperscod and x.cCtaIFCod=CIE.cCtaIFCod "
'    sSql = sSql & "       Where CIE.nTipoEstadistica = 4 and IsNull(x.nNroCuota,0)+1=CIE.nNroCuota"
'    sSql = sSql & "       Group By CIE.cPersCod, CIE.cIFTpo, CIE.cCtaIFCod" & IIf(bCancelacion = True, "", ",CIE.nNroCuota")
'    sSql = sSql & " ) CIE On CIF.cPersCod=CIE.cPersCod And CIF.cCtaIFCod=CIE.cCtaIFCod And CIF.cIFTpo=CIE.cIFTpo "
'
'    If bCancelacion = False Then
'        sSql = sSql & " and CIE.nNrocuota=CIC.nNroCuota "
'    End If
'    'Linea de Credito
'
'    sSql = sSql & " Left Join ColocLineaCredito CLC "
'    sSql = sSql & "         ON CIA.cCodLinCred=CLC.cLineaCred"
'
'    sSql = sSql & " Where   CIF.cCtaIFCod Like '__" & pnMoneda & "%' and CIF.cCtaIFEstado = '" & gEstadoCtaIFActiva & "' "
'
'    'Solo se Pagan las que no estan pagadas
'
'    If bCancelacion = True Then
'
'        If bRestringirFecha = True Then
'            sSql = sSql & " And CICUltimo.dVencimiento >='" & Format(pdFecha, "MM/dd/YYYY") & "' And CICUltimo.dVencimiento < DateAdd(dd,1,'" & Format(pdFecha, "MM/dd/YYYY") & "') "
'        End If
'
'    ElseIf bCancelacion = False Then
'
'        sSql = sSql & " And CIC.bVigente=1 And (CIC.cEstado = '0' or CIC.cEstado='2') and CIC.cTpoCuota <> '6' "
'
'        If bRestringirFecha = True Then
'            sSql = sSql & " And CIC.dVencimiento >='" & Format(pdFecha, "MM/dd/YYYY") & "' And CIC.dVencimiento < DateAdd(dd,1,'" & Format(pdFecha, "MM/dd/YYYY") & "') "
'        End If
'
'    End If
'
'    sSql = sSql & " And CII.nVigente=1 "
'    sSql = sSql & " " & psBuscar & " ORDER BY P.cPersNombre, CIF.cCtaIFCod "

    sSql = "exec stp_sel_GetDatosAdeudadoPago '" & Format(pdFecha, "yyyy-MM-dd") & "'," & pnMoneda & ",'" & psBuscar & "'," & IIf(bCancelacion, 1, 0) & "," & IIf(bRestringirFecha, 1, 0) & "," & IIf(pbControlFechaRep, 1, 0) & "," & gEstadoCtaIFActiva

    oCon.AbreConexion
    Set GetDatosAdeudadoPago = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Exit Function
    
GetDatosAdeudadoPagoErr:
    Call RaiseError(MyUnhandledError, "DACGAdeudados: GetDatosAdeudadoPago Function")

End Function

Public Function CalculaMontoCuota(psPersCod As String, psIFTpo As String, psCtaIFCod As String, pscTpoCuota As String, pnNroCuota As Integer, Optional bCancelacion As Boolean = False) As Currency
Dim rs As New ADODB.Recordset

On Error GoTo CalculaMontoCuotaErr

    CalculaMontoCuota = 0
    
    If bCancelacion = True Then
        sSql = "SELECT  SUM(nCapital) as nCapital "
    Else
        sSql = "Select nCapital "
    End If
    
    sSql = sSql & " FROM CtaIFCalendario "
    sSql = sSql & " WHERE   cPersCod = '" & psPersCod & "' and cIFTpo = '" & psIFTpo & "' and cCtaIFCod = '" & psCtaIFCod & "' "
    sSql = sSql & "         And  cTpoCuota='" & pscTpoCuota & "' "
    
    If bCancelacion = True Then
        sSql = sSql & " And nNroCuota>=" & pnNroCuota
    Else
        sSql = sSql & " And nNroCuota=" & pnNroCuota
    End If
    
    Set rs = oCon.CargaRecordSet(sSql)
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
Set oCon = New DConecta
sSql = "Select isnull(cCtaIFSubCta,'') cCtaIFSubCta "
sSql = sSql & " From ColocLineaCredito "
sSql = sSql & " Where cLineaCred='" & pscLineaCred & "'"

oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
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
Public Function SacarComisionPorDias(ByVal pnDiasTrans As Integer, ByVal pnComision As Currency, ByVal pnPlazo As Integer) As Currency
    SacarComisionPorDias = (pnDiasTrans * pnComision) / pnPlazo
End Function



Public Function GetOpeCta(ByVal psOpeCod As String, ByVal psDH As String, Optional psOpeCtaOrden As String = "0", _
                           Optional psPersCod As String = "", Optional psIFTpo As String = "", _
                           Optional pscCtaIFCod As String = "") As String

Dim rs As New ADODB.Recordset
Dim lsCuenta As String
Dim lsSubCuenta As String
lsCuenta = ""
sSql = "Select Distinct cCtaContCod "
sSql = sSql & " from OpeCta where copecod='" & psOpeCod & "' and copectaorden='" & psOpeCtaOrden & "' "
sSql = sSql & " And cOpeCtaDH='" & psDH & "' " 'and cOpeCtaOpc='1'"
Set oCon = New DConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
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

sSql = "Select cCtaIFSubCta "
sSql = sSql & " From CtaIFFiltro "
sSql = sSql & " Where cPersCod='" & cPerscod & "' And cIFTpo ='" & cIFTpo & "' "
sSql = sSql & " And cCtaIFCod='" & cCtaIFCod & "'"
 
If Len(Trim(cCtaContCod)) > 0 Then
    sSql = sSql & " And cCtaContCod='" & cCtaContCod & "'"
End If
Set oCon = New DConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
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


Public Function GetDatosCuota6(ByVal psPersCod As String, ByVal pscIFTpo As String, ByVal pscCtaIFCod As String, _
                            ByVal pdVencimiento As Date, ByVal cEstados As String, Optional bCancelacion As Boolean = False, Optional bExtornoCancelacion As Boolean = False) As ADODB.Recordset

    On Error GoTo GetDatosCuota6Err
    Set oCon = New DConecta
    
    If bCancelacion = True Or bExtornoCancelacion = True Then
        sSql = "Select isnull(min(nNroCuota),0) as nNroCuota "
    Else
        sSql = "Select isnull(nNroCuota, 0) as nNroCuota "
    End If
    
    sSql = sSql & " From CtaIFCalendario "
    sSql = sSql & " Where cPersCod='" & psPersCod & "' And cIFTpo='" & pscIFTpo & "' And cCtaIFCod='" & pscCtaIFCod & "'  "

    
    If bExtornoCancelacion = True Then
    
    Else
        If Len(Trim(cEstados)) > 0 Then
            sSql = sSql & " And cEstado in(" & cEstados & ") "
        End If
    End If
    
    sSql = sSql & " and cTpoCuota ='6' "
    
    'Si es Cancelacion se coge todos
    If bCancelacion = True Then
    
    ElseIf bExtornoCancelacion = True Then
        sSql = sSql & " And dVencimiento >='" & Format(pdVencimiento, "MM/dd/YYYY") & "' "
    Else
        sSql = sSql & " And dVencimiento >='" & Format(pdVencimiento, "MM/dd/YYYY") & "' And dVencimiento < DateAdd(dd,1,'" & Format(pdVencimiento, "MM/dd/YYYY") & "') "
    End If
    oCon.AbreConexion
    Set GetDatosCuota6 = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Exit Function
    
GetDatosCuota6Err:
    Call RaiseError(MyUnhandledError, "DACGAdeudados: GetDatosCuota6Err Function")

End Function

Public Function GetbExtranjeroInteres(ByVal pscPersCod As String, pscIFTpo As String, pscLineaCred As String, pscMoneda As String) As Boolean

Dim rs As New ADODB.Recordset
sSql = "Select bInteresExtranjero "
sSql = sSql & " From CGCtaIFInstituciones "
sSql = sSql & " Where cPersCod='" & pscPersCod & "' And cIFTpo='" & pscIFTpo & "' and cCodLinCred='" & pscLineaCred & "' And cMoneda='" & pscMoneda & "'"

Set rs = oCon.CargaRecordSet(sSql)
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







''Public Function CGGrabaExtornoAdeudado(ByVal pcOpeCod As String, ByVal pcOpeCodOrigen As String, ByVal pcMoneda As String, ByVal pdFechaExtorno As Date, _
''                                    ByVal lstListadoExtorno As Variant, ByVal lnFilaActual As Long, ByVal cGlosa As String, ByVal psCodAge As String, ByVal psCodUser As String, ByVal lbEliminaMov As Boolean)
''
''
''Dim oDMov As DMov
''
''Dim lsPersCod As String
''Dim lsIFTpo As String
''Dim lsCtaIFCod As String
''
''Dim cTpoCuota As String
''Dim nNroCuota As Long
''Dim nNroCuota6 As Long
''Dim cEstAnt As String
''
''Dim lscCodLinCred As String
''Dim cMonedaPago As String
''
''Dim nSaldoCapAnt As Double
''Dim nSaldoCapNew As Double
''Dim nsaldoCapLPant As Double
''Dim nSaldoCapLPNew As Double
''
''Dim dFecUltAntActSaldos As Date
''Dim dFecUltNewActSaldos As Date
''Dim dfecultantpago As Date
''Dim dFecUltNewPago As Date
''
''Dim dCancelacion As Date
''Dim bCancelacion As Boolean
''
''Dim lbTrans As Boolean
''Dim i As Integer
''Dim lsMovNro As String
''Dim lsMovNroAnt As String
''Dim lnMovNro As Long
''Dim lnMovNroAnt As Long
''
''On Error GoTo CGGrabaExtornoAdeudadoErr
'''
''Set oDMov = New DMov
''
''CGGrabaExtornoAdeudado = 1
''
'''Transaccion
''oDMov.BeginTrans
''
''lbTrans = True
''
''
''    ' ==================== INICIALIZAMOS VALORES ====================
''    i = lnFilaActual
''
''    lsPersCod = ""
''    lsIFTpo = ""
''    lsCtaIFCod = ""
''    cMonedaPago = ""
''    lscCodLinCred = ""
''    cTpoCuota = ""
''    nNroCuota = 0
''    nNroCuota6 = 0
''    cEstAnt = ""
''    nSaldoCapAnt = 0
''    nSaldoCapNew = 0
''    nsaldoCapLPant = 0
''    nSaldoCapLPNew = 0
''    dFecUltAntActSaldos = "01/01/1900"
''    dFecUltNewActSaldos = "01/01/1900"
''    dfecultantpago = "01/01/1900"
''    dFecUltNewPago = "01/01/1900"
''    lsMovNroAnt = ""
''    lsMovNro = ""
''    bCancelacion = False
''
''    ' ==================== OBTENEMOS VALORES ====================
''
''    i = lnFilaActual
''
''    lsPersCod = Trim(lstListadoExtorno.ListItems(i).SubItems(7))
''    lsIFTpo = Trim(lstListadoExtorno.ListItems(i).SubItems(8))
''    lsCtaIFCod = Trim(lstListadoExtorno.ListItems(i).SubItems(9))
''    cMonedaPago = IIf(Trim(lstListadoExtorno.ListItems(i).SubItems(24)) = 1, "2", "1")
''    lscCodLinCred = Trim(lstListadoExtorno.ListItems(i).SubItems(37))
''
''    'Tipo de Cuota
''    cTpoCuota = lstListadoExtorno.ListItems(i).SubItems(12)
''    'Numero de Cuota
''    nNroCuota = Val(lstListadoExtorno.ListItems(i).SubItems(13))
''    nNroCuota6 = Val(lstListadoExtorno.ListItems(i).SubItems(38))
''
''    'Estado Anterior
''    cEstAnt = Trim(lstListadoExtorno.ListItems(i).SubItems(40))
''
''    'Saldo Cap Ant Real
''    nSaldoCapAnt = Val(lstListadoExtorno.ListItems(i).SubItems(29))
''
''    'Saldo Cap New Real
''    nSaldoCapNew = Val(lstListadoExtorno.ListItems(i).SubItems(30))
''
''    'Saldo Cap LP Ant Real
''    nsaldoCapLPant = Val(lstListadoExtorno.ListItems(i).SubItems(31))
''
''    'Saldo Cap LP New Real
''    nSaldoCapLPNew = Val(lstListadoExtorno.ListItems(i).SubItems(32))
''
''
''    'Fecha de Ultima Provision o Pago Efectuado Anterior
''    dFecUltAntActSaldos = CDate(lstListadoExtorno.ListItems(i).SubItems(33))
''
''    'Fecha de Ultima Provision o Pago Efectuado
''    dFecUltNewActSaldos = CDate(lstListadoExtorno.ListItems(i).SubItems(34))
''
''    'Fecha de Ultimo Pago Efectuado Anterior
''    dfecultantpago = CDate(lstListadoExtorno.ListItems(i).SubItems(35))
''
''    'Fecha de Ultimo Pago Efectuado
''    dFecUltNewPago = CDate(lstListadoExtorno.ListItems(i).SubItems(36))
''
''    'Cancelacion
''    If Format(CDate(lstListadoExtorno.ListItems(i).SubItems(41)), "YYYYMMdd") <> "19000101" Then
''        bCancelacion = True
''        dCancelacion = CDate(lstListadoExtorno.ListItems(i).SubItems(41))
''    Else
''        bCancelacion = False
''        dCancelacion = "01/01/1900"
''    End If
''
''    'lsMovNroAnt
''    lsMovNroAnt = Trim(lstListadoExtorno.ListItems(i).SubItems(39))
''
''    'lnMovNroAnt
''    lnMovNroAnt = Val(lstListadoExtorno.ListItems(i).SubItems(3))
''
''    lsMovNro = oDMov.GeneraMovNro(pdFechaExtorno, psCodAge, psCodUser)
''
''    oDMov.InsertaMov lsMovNro, pcOpeCod, cGlosa, gMovEstContabMovContable, gMovFlagVigente
''    lnMovNro = oDMov.GetnMovNro(lsMovNro)
''
''
''
''    ' ==================== ACTUALIZACION DE DATOS DE ADEUDADOS ====================
''
''    'Extorno de Registro
''    If pcOpeCodOrigen = "401701" Or pcOpeCodOrigen = "402701" Then
''
''        'Extorno Contable
''        oDMov.ExtornaMovimiento lsMovNro, lnMovNroAnt, pcOpeCod, cGlosa, lbEliminaMov, lsMovNro
''
''        'Registro se Anula
''        oDMov.ActualizaCtaIF lsPersCod, lsIFTpo, lsCtaIFCod, , , , , , , gEstadoCtaIFAnulada, , lsMovNro
''
''        'Si se extorna un registro, colocar fecha de cancelacion para q no salga en los reportes
''        'Fecha de Cancelacion
''        oDMov.CGActualizaDatosFAdeudado lsPersCod, lsIFTpo, lsCtaIFCod, False, "01/01/1900", False, "01/01/1900", False, 0, False, 0, True, pdFechaExtorno
''
''        'Referencia
''        oDMov.InsertaMovRef lnMovNro, lnMovNroAnt
''
''    'Extorno de Confirmacion
''    ElseIf pcOpeCodOrigen = "401702" Or pcOpeCodOrigen = "402702" Then 'Confirmacion
''
''        'Extorno Contable
''        oDMov.ExtornaMovimiento lsMovNro, lnMovNroAnt, pcOpeCod, cGlosa, lbEliminaMov, lsMovNro
''
''        'CtaIF
''        oDMov.ActualizaCtaIF lsPersCod, lsIFTpo, lsCtaIFCod, , , , , , , gEstadoCtaIFRegistrada, , lsMovNro
''
''        'CtaIFAdeudados
''        oDMov.CGActualizaDatosFAdeudado lsPersCod, lsIFTpo, lsCtaIFCod, False, "01/01/1900", False, "01/01/1900", False, 0, True, nsaldoCapLPant
''
''        'Extorno del CtaIFFiltro
''        oDMov.CGBorraCuentaIFFiltro lsPersCod, lsIFTpo, lsCtaIFCod
''
''        'Referencia
''        oDMov.InsertaMovRef lnMovNro, lnMovNroAnt
''
''    'Extorno de Provision
''    ElseIf pcOpeCodOrigen = "401705" Or pcOpeCodOrigen = "402705" Then 'Provision
''
''        'Extorno Contable
''        oDMov.ExtornaMovimiento lsMovNro, lnMovNroAnt, pcOpeCod, cGlosa, lbEliminaMov, lsMovNro
''
''        'CtaIF
''        oDMov.ActualizaCtaIF lsPersCod, lsIFTpo, lsCtaIFCod, , , , , , , gEstadoCtaIFActiva, , lsMovNro
''
''        'CtaIFAdeudados
''        oDMov.CGActualizaDatosFAdeudado lsPersCod, lsIFTpo, lsCtaIFCod, False, "01/01/1900", True, dFecUltAntActSaldos, False, 0, True, nsaldoCapLPant
''
''        'CtaIFCalendario
''
''        If nNroCuota6 = 0 Then
''            'La Provision para ambos tipos regresa a ser CERO
''            'Sin cuota 6
''            oDMov.CGActualizaDatosCalendario lsPersCod, lsIFTpo, lsCtaIFCod, True, cTpoCuota, nNroCuota, False, "", 0, "0"
''        ElseIf nNroCuota6 > 0 Then
''            'La Provision para ambos tipos regresa a ser CERO
''            'Con Cuota 6
''            oDMov.CGActualizaDatosCalendario lsPersCod, lsIFTpo, lsCtaIFCod, True, cTpoCuota, nNroCuota, True, "6", nNroCuota6, "0"
''        End If
''
''        'Referencia
''        oDMov.InsertaMovRef lnMovNro, lnMovNroAnt
''
''    'Extorno de Pago
''    ElseIf pcOpeCodOrigen = "401706" Or pcOpeCodOrigen = "402706" Then 'Pago
''
''        'Extorno Contable
''        oDMov.ExtornaMovimiento lsMovNro, lnMovNroAnt, pcOpeCod, cGlosa, lbEliminaMov, lsMovNro
''
''        'CtaIF
''        oDMov.ActualizaCtaIF lsPersCod, lsIFTpo, lsCtaIFCod, , , , , , , gEstadoCtaIFActiva, , lsMovNro
''
''        'CtaIFAdeudados
''        oDMov.CGActualizaDatosFAdeudado lsPersCod, lsIFTpo, lsCtaIFCod, True, dfecultantpago, True, dFecUltAntActSaldos, True, nSaldoCapAnt, True, nsaldoCapLPant, False, "01/01/1900", bCancelacion
''
''        'CtaIFCalendario
''
''        If nNroCuota6 = 0 Then
''            'La Provision para el tipo 2 regresa a ser CERO
''            'Cuota 2
''            oDMov.CGActualizaDatosCalendario lsPersCod, lsIFTpo, lsCtaIFCod, True, cTpoCuota, nNroCuota, False, "", 0, cEstAnt, False, bCancelacion
''        ElseIf nNroCuota6 > 0 Then
''            'La Provision para el tipo 2 regresa a ser CERO
''            'Cuota 2
''            oDMov.CGActualizaDatosCalendario lsPersCod, lsIFTpo, lsCtaIFCod, True, cTpoCuota, nNroCuota, False, "", 0, cEstAnt, False, bCancelacion
''            'La Cuota 6 al extornar el pago regresa a cero asi haya habido provision
''            oDMov.CGActualizaDatosCalendario lsPersCod, lsIFTpo, lsCtaIFCod, False, "", 0, True, "6", nNroCuota6, "0", False, bCancelacion
''        End If
''
''        'Referencia
''        oDMov.InsertaMovRef lnMovNro, lnMovNroAnt
''
''    End If
''
''    If Mid(gsOpeCod, 3, 1) = "2" Then
''       oDMov.GeneraMovME lnMovNro, lsMovNro
''    End If
''    oDMov.ActualizaSaldoMovimiento lsMovNro, "+"
''
''
'''
''
''oDMov.CommitTrans
''lbTrans = False
''CGGrabaExtornoAdeudado = 0
''Set oDMov = Nothing
''
''Exit Function
''CGGrabaExtornoAdeudadoErr:
''    lsMsgErr = Err.Description
''    If lbTrans Then
''        oDMov.RollbackTrans
''        lbTrans = False
''    End If
''    Err.Raise vbObjectError + 100, "CGGrabaExtornoAdeudado", lsMsgErr & " Error Graba Provision de Adeudados"
''
''
''
''Private Sub Class_Initialize()
''    Dim oImp As DImpresoras
''    Set oImp = New DImpresoras
''    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
''    Set oImp = Nothing
''
''    lbConexionPropia = True
''    Set oCon = New DConecta
''    oCon.AbreConexion
''
''End Sub

Public Function GetExisteCuentaContable(psCuenta As String, psSubCuenta As String) As Boolean
Dim rs As New ADODB.Recordset
'Dim oCon As New DConecta
On Error GoTo GetExisteCuentaContableErr

Dim oCon As New DConecta

'Verificamos Si existe la Cuenta Contable a Generar
oCon.AbreConexion
sSql = "SELECT cCtaContCod "
sSql = sSql & "FROM CtaCont "
sSql = sSql & " WHERE cCtaContCod = '" & psCuenta & psSubCuenta & "'"

Set rs = oCon.CargaRecordSet(sSql)
If rs.EOF Then
    RSClose rs
    GetExisteCuentaContable = False
Else
    RSClose rs
    GetExisteCuentaContable = True
End If

oCon.CierraConexion
Exit Function
GetExisteCuentaContableErr:
    Call RaiseError(MyUnhandledError, "DACGAdeudados:GetExisteCuentaContable Method")

End Function


Public Function GrabaConfirmacionAdeudado(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal psCtaDebe As String, ByVal psSubCtaDebe As String, ByVal psCtaHaber As String, _
                                ByVal pnImporte As Double, _
                                ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String, _
                                ByVal pnMovNroRef As Long, ByVal pdFecha As Date, ByVal pnTasaVac As Double, _
                                ByVal nMontoPrestado As Double, ByVal nMontoPrestadoReal As Double, ByVal cMonedaPago As String, _
                                Optional ByVal psCuentaCorto As String = "", Optional ByVal psCuentaLargo As String = "") As Integer

Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim oCtaIf As NCajaCtaIF
Dim lnSaldo As Currency
Dim ldFecha As Date
Dim lnFactor As Integer
Dim nImporteLP As Currency
Dim lsMsgErr As String

Set oCtaIf = New NCajaCtaIF
Set oDMov = New DMov

On Error GoTo GrabaConfirmacionAdeudadoErr
nImporteLP = 0
lnFactor = 1

If Val(Left(psCtaIFCod, 2)) = gTpoCtaIFCtaAdeud Then
   nImporteLP = CalculaSaldoAdeudadoLargoPlazo(GetFechaMov(psMovNro, True), psPersCod, psIFTpo, psCtaIFCod)
   lnFactor = -1
End If

GrabaConfirmacionAdeudado = 1
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc
lnMovNro = oDMov.GetnMovNro(psMovNro)
ldFecha = CDate(GetFechaMov(psMovNro, True))

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"

'guardamos la cuenta de Bancos en el debe : Primer Capital
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1

If pnImporte - nImporteLP = 0 Then
Else
    If Mid(psCtaIFCod, 3, 1) = "1" And cMonedaPago = "2" Then
        'ALPA20130709
        'oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe & psSubCtaDebe, (pnImporte - Format(nImporteLP * pnTasaVac, "0.00")) * lnFactor
        oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCuentaCorto, (pnImporte - Format(nImporteLP * pnTasaVac, "0.00")) * lnFactor
        
    Else
        'ALPA20130709
        'oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe & psSubCtaDebe, Format(pnImporte - nImporteLP, "0.00") * lnFactor
        oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCuentaCorto, Format(pnImporte - nImporteLP, "0.00") * lnFactor
    End If
End If

lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod

'Creamos la SubCuenta de la Inst.Financ en CtaIFFiltro
oDMov.InsertaCuentaIFFiltro psPersCod, psIFTpo, psCtaIFCod, psCtaDebe, psSubCtaDebe

'ADEUDADOS Largo Plazo
If nImporteLP > 0 Then
    lnMovOrden = 0
    lnMovItem = lnMovItem + 1
    
    If Mid(psCtaIFCod, 3, 1) = "1" And cMonedaPago = "2" Then
        oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCuentaLargo, Format(nImporteLP * pnTasaVac, "0.00") * lnFactor
    Else
        oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCuentaLargo, Format(nImporteLP, "0.00") * lnFactor
    End If
    
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod
    
    'oDMov.CGActualizaAdeudadosSaldoCap psPersCod, psIFTpo, psCtaIfCod, , nImporteLP
    oDMov.CGActualizaDatosFAdeudado psPersCod, psIFTpo, psCtaIFCod, False, "01/01/1900", False, "01/01/1900", False, 0, True, nImporteLP
    
End If

'Guardamos la Cuenta de la Pendiente que se elimina
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, (pnImporte * -1) * lnFactor
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCod, psIFTpo, psCtaIFCod

'Pasamos a activo el estado de la cuenta aperturada
oDMov.ActualizaCtaIF psPersCod, psIFTpo, psCtaIFCod, , , , , , , gEstadoCtaIFActiva, , psMovNro
oDMov.InsertaMovRef lnMovNro, pnMovNroRef

'Estadistica
If Mid(psCtaIFCod, 3, 1) = "1" And cMonedaPago = "2" Then
    oDMov.GrabaCtaIFEstadistica psPersCod, Format(psIFTpo, "00"), psCtaIFCod, _
                                 "0", 0, lnMovNro, 2, nMontoPrestado, 0, 0, 0, _
                                 Format(nImporteLP * pnTasaVac, "0.00"), 0, nMontoPrestado, nMontoPrestado, 0, _
                                Format(nImporteLP * pnTasaVac, "0.00"), 1, _
                                nMontoPrestadoReal, 0, 0, 0, nImporteLP, 0, nMontoPrestadoReal, nMontoPrestadoReal, 0, _
                                 nImporteLP, pdFecha, pdFecha, pdFecha, pdFecha, "0"
Else
    oDMov.GrabaCtaIFEstadistica psPersCod, Format(psIFTpo, "00"), psCtaIFCod, _
                                "0", 0, lnMovNro, 2, nMontoPrestadoReal, 0, 0, 0, _
                                Format(nImporteLP, "0.00"), 0, nMontoPrestadoReal, nMontoPrestadoReal, 0, _
                                nImporteLP, 0, _
                                nMontoPrestadoReal, 0, 0, 0, nImporteLP, 0, nMontoPrestadoReal, nMontoPrestadoReal, 0, _
                                nImporteLP, pdFecha, pdFecha, pdFecha, pdFecha, "0"

End If

If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"

oDMov.CommitTrans
lbTrans = False

GrabaConfirmacionAdeudado = 0
Set oDMov = Nothing


Exit Function
GrabaConfirmacionAdeudadoErr:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaConfApertura", lsMsgErr & " Error en Grabación de Confirmación de Apertura "
End Function

Public Function CalculaSaldoAdeudadoLargoPlazo(pdFecha As Date, psPersCod As String, psIFTpo As String, psCtaIFCod As String, Optional psTpoCuota As CGTipoCuotCalIF = gCGTipoCuotCalIFCuota, Optional pCon As Object) As Currency
On Error GoTo CalculaSaldoAdeudadoLargoPlazoErr
Dim sql As String
Dim oCon As New DConecta
Dim rs As New ADODB.Recordset

CalculaSaldoAdeudadoLargoPlazo = 0
'ALPA
'    sql = "SELECT ISNULL(SUM(nCapital),0) nCapital FROM CtaIFCalendario " _
'        & "WHERE cPersCod = '" & psPersCod & "' and cIFTpo = '" & psIFTpo & "' and cCtaIFCod = '" & psCtaIfCod & "' and " _
'        & "      datediff(d,'" & Format(pdFecha, gsFormatoFecha) & "',dVencimiento) > 360 and cEstado = '" & gTpoEstCuotaAdeudPend & "' "
    sql = "SELECT ISNULL(SUM(nCapital),0) nCapital FROM CtaIFCalendario "
    sql = sql & " WHERE cPersCod = '" & psPersCod & "' and cIFTpo = '" & psIFTpo & "' and cCtaIFCod = '" & psCtaIFCod & "' and "
    sql = sql & " datediff(d,'" & Format(pdFecha, gsFormatoFecha) & "',dVencimiento) > 360 and cEstado = '" & gTpoEstCuotaAdeudPend & "' "
    sql = sql & " and bVigente=1"

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
    Set oCon = New DConecta
    oCon.AbreConexion
End Sub

Private Sub Class_Terminate()
If lbConexionPropia = True Then
    oCon.CierraConexion
    Set oCon = Nothing
End If
End Sub

Public Function GetReporteDatosAdeudadoPago(pdFecha As Date, pnMoneda As Moneda, Optional psBuscar As String = "", Optional bCancelacion As Boolean = False, Optional bRestringirFecha As Boolean = False, Optional pbControlFechaRep As Boolean = False) As ADODB.Recordset

    On Error GoTo GetReporteDatosAdeudadoPagoErr
    Dim oCon As New DConecta
    
    sSql = "Select CIF.cPersCod, P.cPersNombre, CIF.cIFTpo, CIF.cCtaIFCod, CIF.cCtaIFDesc, "
    
    If bCancelacion = True Then
        sSql = sSql & " CICUltimo.dVencimiento, CalRef.nNroCuota, '2' as cTpoCuota, "
        sSql = sSql & " CICUltimo.cEstado as cEstadoCuota, "
    ElseIf bCancelacion = False Then
        sSql = sSql & " CIN.dVencimiento, CIC.nNroCuota, CIC.cTpoCuota, "
        sSql = sSql & " CIC.cEstado as cEstadoCuota, "
    End If
    
    sSql = sSql & " CII.nCtaIFIntPeriodo, CII.nCtaIFIntValor, CIA.cMonedaPago, CIA.nMontoPrestado, "
    sSql = sSql & " ISNULL(CIE.nSumaInteresMovido, 0) as nInteresProvisionado, "
    sSql = sSql & " ISNULL(CIE.nSumaInteresMovidoReal, 0) as nInteresProvisionadoReal, "
    sSql = sSql & " CIA.dCuotaUltModSaldos, CIA.dCuotaUltPago, "
    
    If bCancelacion = True Then
        sSql = sSql & " CIA.nSaldoCap, CIA.nSaldoCapLP, CIC.nCapital, IsNull(CIC.nInteres,0) + IsNull(CIC.nInteresGracia,0) As nInteres, CIC.nComision, CIC.nTotalCuota, "
    Else
        'sSql = sSql & " CIA.nSaldoCap, CIA.nSaldoCapLP, CIC.nCapital + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,1) nCapital, IsNull(CIC.nInteres,0) +  dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,2) + IsNull(CIC.nInteresGracia,0) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,3) As nInteres, CIC.nComision + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,4) nComision, CIC.nTotalCuota + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,1) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,2) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,3) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,4) nTotalCuota, "
        sSql = sSql & " CIA.nMontoPrestado - isnull(SalCal.nCapital,0) nSaldoCap, CIA.nSaldoCapLP, " 'NAGL 20170519
        sSql = sSql & " CIC.nCapital + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,1) nCapital, IsNull(CIC.nInteres,0) +  dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,2) + IsNull(CIC.nInteresGracia,0) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,3) As nInteres, CIC.nComision + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,4) nComision, CIC.nTotalCuota + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,1) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,2) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,3) + dbo.CGConceptosMalPagador(CIF.cPersCod,CIF.cIFTpo, CIF.cCtaIFCod,CIC.nNroCuota,4) nTotalCuota, "
    End If
    
    sSql = sSql & " ISNULL(CIA.cCodLinCred, '') as cCodLinCred, ISNULL(CLC.cDescripcion, '') as cDesLinCred , "
    sSql = sSql & " iVacAper=(select nIndiceVac from indiceVac where dIndiceVAC=CIF.dctaifaper), (Select max(dVencimiento) from CtaIfCalendario Calen where Calen.cPersCod = CIF.cPersCod And Calen.cIFTpo = CIF.cIFTpo And Calen.cCtaIfcod = CIF.cCtaIFCod And bVigente = 1) dVencimientoFinal, (Select dCtaIFAper from CtaIf where CtaIf.cPersCod = CIF.cPersCod And CtaIf.cIFTpo = CIF.cIFTpo And CtaIf.cCtaIfcod = CIF.cCtaIFCod) dCtaIFAper,CIF.cPersCod,CIF.cIFTpo,CIF.cCtaIFCod, "
    sSql = sSql & " case when CIF.nCtaIFPlazo = 30 then 'Mensual'"
    sSql = sSql & "     when CIF.nCtaIFPlazo = 60 then 'Bimestral'"
    sSql = sSql & "     when CIF.nCtaIFPlazo between 70 and 95 then 'Trimestral'"
    sSql = sSql & "     when CIF.nCtaIFPlazo = 180 then 'Sementral'"
    sSql = sSql & "     when CIF.nCtaIFPlazo between 1600 and 3000 then 'Quinquenal'"
    sSql = sSql & "         else convert(varchar(4),CIF.nCtaIFPlazo)"
    sSql = sSql & " end cCtaIFPlazo,"
    sSql = sSql & " isnull(CES.nCapitalPagado,0) nCapitalPagado,"
    sSql = sSql & " isnull(CES.nInteresPagado,0) nInteresPagado,"
    sSql = sSql & " isnull(CIN.nSumInteres,0) nSumInteres,"
    sSql = sSql & " isnull(CIN.nNroCuotas,0) nNroCuotas, "
    'sSql = sSql & " isnull(CIN.nSumInteresPen,0) nSumInteresPen"
    sSql = sSql & " isnull(SalCal.nInteres,0) nSumInteresPen"
    
    sSql = sSql & " From CtaIF CIF "
    
    sSql = sSql & " Inner Join CtaIFAdeudados CIA "
    sSql = sSql & "         On CIF.cPersCod=CIA.cPersCod And CIF.cIFTpo=CIA.cIFTpo And CIF.cCtaIFCod=CIA.cCtaIFCod "
    
    If pbControlFechaRep Then
        sSql = sSql & " Inner Join CtaIF CtaIF "
        sSql = sSql & "         On CIF.cPersCod = CtaIF.cPersCod And CIF.cIFTpo = CtaIF.cIFTpo And CIF.cCtaIFCod = CtaIF.cCtaIFCod And CtaIF.dCtaIfAper <= '" & Format(pdFecha, gsFormatoFecha) & "' "
    End If
    
    'Proxima Cuota Pendiente
    sSql = sSql & " Inner Join (    SELECT cPersCod, cIFTpo, cCtaIFCod, MIN(nNroCuota) nNroCuota "
    sSql = sSql & "                 From CtaIFCalendario "
    sSql = sSql & "                 WHERE bVigente=1 And (cEstado = '0' or cEstado='2') and cTpoCuota <> '6' "
    sSql = sSql & "                     And Substring(cCtaIFCod, 3,1)='" & Trim(Str(pnMoneda)) & "' " '& Replace(psBuscar, "CIF.", "")
    sSql = sSql & "                 GROUP BY cPersCod, cIFTpo, cCtaIFCod "
    sSql = sSql & "             ) CalRef ON CIF.cPersCod=CalRef.cPersCod And CIF.cIFTpo=CalRef.cIFTpo And CIF.cCtaIFCod=CalRef.cCtaIFCod "
    
    If bCancelacion = True Then
        
        'Sacamos todo el capital, intereses y comisiones pendientes Pero de lo no Concesional,
        'Porque lo concesional se PERDONA!!!
        
        sSql = sSql & " Inner Join ( "
        sSql = sSql & " select  CIK.cPersCod, CIK.cIFTpo, CIK.cCtaIFCod, Sum(IsNull(CIK.nCapital,0)) as nCapital, Sum(IsNull(CIK.nInteres,0)) AS nInteres, Sum(IsNull(CIK.nInteresGracia,0)) as nInteresGracia, "
        sSql = sSql & " sum(IsNull(CIK.nComision,0)) As nComision, Sum(IsNull(CIK.nCapital,0) + IsNull(CIK.nInteres,0) + IsNull(CIK.nInteresGracia,0) + IsNull(CIK.nComision,0)) as nTotalCuota "
        sSql = sSql & " from ctaifcalendario CIK "
        sSql = sSql & " Where CIK.bVigente=1 And CIK.cEstado in('0', '2')  " 'And CIK.cTpoCuota='2'
        sSql = sSql & " Group By CIK.cPersCod, CIK.cIFTpo, CIK.cCtaIFCod "
        sSql = sSql & " ) CIC On CIF.cPersCod=CIC.cPersCod And CIF.cIFTpo=CIC.cIFTpo And CIF.cCtaIFCod=CIC.cCtaIFCod "
    
        'Sacamos el estado de la ultima cuota pendiente
        sSql = sSql & " Inner Join ( "
        sSql = sSql & " Select CIK.cPersCod, CIK.cIFTpo, CIK.cCtaIFCod, CIK.nNroCuota, CIK.cEstado, CIK.dVencimiento "
        sSql = sSql & " From CtaIFCalendario CIK "
        sSql = sSql & " Where CIK.bVigente=1 And CIK.cEstado in('0', '2') "
        sSql = sSql & " And CIK.cTpoCuota='2' "
        sSql = sSql & " ) CICUltimo On CIF.cPersCod=CICUltimo.cPersCod And CIF.cIFTpo=CICUltimo.cIFTpo And CIF.cCtaIFCod=CICUltimo.cCtaIFCod "
        sSql = sSql & " And CalRef.nNroCuota=CICUltimo.nNroCuota "
    
    ElseIf bCancelacion = False Then
    
        'Sacamos el capital, intereses y comisiones de la prox cuota pendiente
        sSql = sSql & " Inner Join CtaIFCalendario CIC "
        sSql = sSql & "         ON CIF.cPersCod=CIC.cPersCod And CIF.cIFTpo=CIC.cIFTpo And CIF.cCtaIFCod=CIC.cCtaIFCod "
        sSql = sSql & "                 AND CIC.nNroCuota=CalRef.nNroCuota "
    
    End If
    
    sSql = sSql & " Inner Join CtaIFInteres CII "
    sSql = sSql & "         On CIF.cPersCod=CII.cPersCod And CIF.cIFTpo=CII.cIFTpo And CIF.cCtaIFCod=CII.cCtaIFCod "
    sSql = sSql & " Inner JOIN Persona P "
    sSql = sSql & "         ON CIF.cPersCod = P.cPersCod "
    
    'Ultimo Interes provisionado
    sSql = sSql & " Left Join ( "
    sSql = sSql & "       Select CIE.cPersCod, CIE.cIFTpo, CIE.cCtaIFCod, " & IIf(bCancelacion = True, "", " CIE.nNroCuota,")
    sSql = sSql & "           SUM(CIE.nInteres1Movido) as nSumaInteresMovido, SUM(CIED.nInteres1MovidoReal) as nSumaInteresMovidoReal "
    sSql = sSql & "       From CtaIFEstadistica CIE "
    sSql = sSql & "       Inner Join CtaIFEstadisticaDet CIED "
    sSql = sSql & "                       On CIE.nMovNro = CIED.nMovNro "
    sSql = sSql & "       Inner Join Mov m on   cie.nmovnro=m.nmovnro and m.nmovflag<>1 " & IIf(pbControlFechaRep, " And cMovNro < '" & Format(pdFecha + 1, gsFormatoMovFecha) & "' ", "")
    sSql = sSql & "       Left Join ( Select cPersCod,cIFTpo,cCtaIFCod,max(nNroCuota) nNroCuota"
    sSql = sSql & "                from CtaIFEstadistica a"
    sSql = sSql & "                Inner Join Mov b on a.nmovnro=b.nmovnro and b.nmovflag<>1 " & IIf(pbControlFechaRep, " And cMovNro < '" & Format(pdFecha + 1, gsFormatoMovFecha) & "' ", "")
    sSql = sSql & "                Where nTipoEstadistica = 3"
    sSql = sSql & "                group by cPersCod,cIFTpo,cCtaIFCod"
    sSql = sSql & "           ) X   on x.cperscod=CIE.cperscod and x.cCtaIFCod=CIE.cCtaIFCod "
    sSql = sSql & "       Where CIE.nTipoEstadistica = 4 and IsNull(x.nNroCuota,0)+1=CIE.nNroCuota"
    sSql = sSql & "       Group By CIE.cPersCod, CIE.cIFTpo, CIE.cCtaIFCod" & IIf(bCancelacion = True, "", ",CIE.nNroCuota")
    sSql = sSql & " ) CIE On CIF.cPersCod=CIE.cPersCod And CIF.cCtaIFCod=CIE.cCtaIFCod And CIF.cIFTpo=CIE.cIFTpo "
    
    If bCancelacion = False Then
        sSql = sSql & " and CIE.nNrocuota=CIC.nNroCuota "
    End If
    'Linea de Credito
    
    sSql = sSql & " Left Join ColocLineaCredito CLC "
    sSql = sSql & "         ON CIA.cCodLinCred=CLC.cLineaCred"
    
    'ALPA 20110812*****************************
    sSql = sSql & " left join ("
    sSql = sSql & " select cPersCod,cIFTpo,cCtaIFCod,sum(nInteres) nSumInteres,max(nNroCuota) nNroCuotas,"
    sSql = sSql & " sum(case cEstado when '0' then nInteres end) nSumInteresPen,max(dVencimiento) dVencimiento "
    sSql = sSql & " From CtaIFCalendario"
    sSql = sSql & " Where bvigente = 1"
    sSql = sSql & " group by cPersCod,cIFTpo,cCtaIFCod"
    sSql = sSql & " ) CIN on CIF.cPersCod=CIN.cPersCod"
    sSql = sSql & " and CIF.cIFTpo=CIN.cIFTpo"
    sSql = sSql & " and CIF.cCtaIFCod=CIN.cCtaIFCod"
    
    sSql = sSql & " Left Join"
    sSql = sSql & "             ("
    sSql = sSql & "             select CE.cPersCod,CE.cCtaIFCod,CE.cIFTpo,"
    sSql = sSql & "                 Sum(isnull(CC.nCapital,0)) nCapitalPagado,"
    sSql = sSql & "                 Sum(isnull(CC.nInteres,0)) nInteresPagado"
    sSql = sSql & "             from CtaIFEstadistica CE"
    sSql = sSql & "                 inner join Mov M on CE.nMovNro=M.nMovNro"
    sSql = sSql & "                 inner join CtaIFCalendario CC on"
    sSql = sSql & "                 CE.cIFTpo = CC.cIFTpo"
    sSql = sSql & "                 and CE.cPersCod=CC.cPersCod"
    sSql = sSql & "                 and CE.cCtaIFCod=CC.cCtaIFCod"
    sSql = sSql & "                 and CE.cTpoCuota=CC.cTpoCuota"
    sSql = sSql & "                 and CE.nNroCuota=CC.nNroCuota"
    sSql = sSql & "             where M.cOpeCod like '40_806'"
    sSql = sSql & "                 and DateDiff(m,convert(datetime,left(M.cMovNro,8)),'" & Format(pdFecha, gsFormatoMovFecha) & "')=0"
    sSql = sSql & "                 and CC.bvigente=1"
    sSql = sSql & "             group by CE.cPersCod,CE.cCtaIFCod,CE.cIFTpo"
    sSql = sSql & "             ) CES on CIF.cPersCod=CES.cPersCod"
    sSql = sSql & "                                     and CES.cIFTpo=CIN.cIFTpo"
    sSql = sSql & "                                     and CES.cCtaIFCod=CIN.cCtaIFCod"
    
    sSql = sSql & " Left Join"
    sSql = sSql & "            ("
    'ALPA20131120**********************************************************************************************
    'sSql = sSql & "                select cPersCod,cCtaIFCod,cIFTpo,sum(nCapital) nCapital,sum(nInteres) nInteres"
    sSql = sSql & "                select cPersCod,cCtaIFCod,cIFTpo,sum(nCapital+isnull(nCapitalConce,0)) nCapital,"
     sSql = sSql & "               nInteres =(select sum(nInteres+isnull(nComision,0)+isnull(nInteresConce,0)+isnull(nComisionConce,0)) from CtaIFCalendario where cPersCod = CFcal.cPersCod and cCtaIFCod = CFcal.cCtaIFCod and bVigente = 1 and dateDiff(d,dVencimiento,'" & Format(pdFecha, "YYYY/MM/DD") & "')<0)"
    sSql = sSql & "                From CtaIFCalendario CFcal" 'NAGL 20170519
    sSql = sSql & "                Where bVigente = 1"
     sSql = sSql & "                and dateDiff(d,dVencimiento,'" & Format(pdFecha, "YYYY/MM/DD") & "')>0" 'NAGL 20170519 ANTES < 0
    sSql = sSql & "                group by cPersCod,cCtaIFCod,cIFTpo"
    sSql = sSql & "            ) SalCal on CIF.cPersCod=SalCal.cPersCod and SalCal.cIFTpo=CIN.cIFTpo and SalCal.cCtaIFCod=CIN.cCtaIFCod"
    '*******************************************

    sSql = sSql & " Where   CIF.cCtaIFCod Like '__" & pnMoneda & "%' and CIF.cCtaIFEstado = '" & gEstadoCtaIFActiva & "' "
    
    'Solo se Pagan las que no estan pagadas
    
    If bCancelacion = True Then
        
        If bRestringirFecha = True Then
            sSql = sSql & " And CICUltimo.dVencimiento >='" & Format(pdFecha, "MM/dd/YYYY") & "' And CICUltimo.dVencimiento < DateAdd(dd,1,'" & Format(pdFecha, "MM/dd/YYYY") & "') "
        End If
        
    ElseIf bCancelacion = False Then
        
        sSql = sSql & " And CIC.bVigente=1 And (CIC.cEstado = '0' or CIC.cEstado='2') and CIC.cTpoCuota <> '6' "
        
        If bRestringirFecha = True Then
            sSql = sSql & " And CIC.dVencimiento >='" & Format(pdFecha, "MM/dd/YYYY") & "' And CIC.dVencimiento < DateAdd(dd,1,'" & Format(pdFecha, "MM/dd/YYYY") & "') "
        End If
        
    End If
    
    sSql = sSql & " And CII.nVigente=1 "
    sSql = sSql & " " & psBuscar & " ORDER BY P.cPersNombre,ISNULL(CLC.cDescripcion, ''), CIF.cCtaIFCod "

    oCon.AbreConexion
    Set GetReporteDatosAdeudadoPago = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Exit Function
    
GetReporteDatosAdeudadoPagoErr:
    Call RaiseError(MyUnhandledError, "DACGAdeudados: GetDatosAdeudadoPago Function")

End Function



