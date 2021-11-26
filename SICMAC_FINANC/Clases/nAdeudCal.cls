VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nAdeudCal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Type Calendario
    Cuota As Long
    FechaPago As Date
    Capital As Currency
    Interes As Currency
    Total As Currency
    Saldo As Currency
    cTpoCuota As CGTipoCuotCalIF
    cEstado As String
    nDiasPago As Integer
    nInteresPagado As Currency
    nComisionCuota As Currency
    cMovNro As String
    CapitalConc As Currency
    InteresConc As Currency
    nComisionCuotaConc As Currency
End Type
Dim ltCalendario() As Calendario
Dim ltcalendarioC() As Calendario
Dim ltCalendarioItera() As Calendario   'ANGC202102

Dim oNAdeud As NCajaAdeudados
Dim oDAdeud As DCaja_Adeudados
Dim vsCentralPers As String
Dim vsCentralCom As String
Dim vsCentralImg As String
Dim vnTotalCapital As Currency
Dim vnTotalInteres As Currency
'ALPA 20130614**
Dim vnTotalCapitalC As Currency
Dim vnTotalInteresC As Currency
Dim vnTotalComisioC As Currency
'***************
'NAGL 20170519
Dim vnTotalComision As Currency
Dim vnTotalComisionC As Currency
'********************
Dim vnTotalGeneral As Currency
Dim vsFormatoFecha As String

Public Sub Inicio(psFormatoFecha As String)
Attribute Inicio.VB_UserMemId = 0
    vsFormatoFecha = psFormatoFecha
End Sub

Public Sub GeneraCalendarioCuotaVariable(fg As Object, ByVal lnCapitalIni As Currency, ByVal lnNumCuotas As Integer, ByVal lnPlazoCuotas As Integer, _
       ByVal lnPeriodo As Long, ByVal ldFechaDesemb As Date, _
       ByVal lnPerGracia As Integer, ByVal lnInteres As Currency, Optional ByVal lnCuotaK As Integer = 0, Optional pnComisionCuota As Currency, _
       Optional ByVal pnCapitalConcesional As Currency = 0, Optional ByVal pnComisionCuotaConc As Currency = 0)
' Se agregó la variable pnCapitalConcesional ALPA 20130614)

On Error GoTo ErrorGeneraCal
    Dim lnTotalPagar As Currency
    Dim I As Integer
    Dim lnCapitalCuota As Currency
    Dim lnCapital As Currency
    Dim lnDiasTrans As Long
    Dim lnPorCentaje As Double
    Dim lnSaldoCapital As Double
    Dim lnSumaCapital As Currency
    Dim lnSumaCapitalConc As Currency
    
    Dim lnAcuCapital As Currency
    Dim lnAcuCapitalConc As Currency
    
    'ALPA 20130614***************
    lnSaldoCapital = lnCapitalIni
    lnPorCentaje = pnCapitalConcesional / lnSaldoCapital
    lnCapitalIni = lnCapitalIni - pnCapitalConcesional
'****************************

    If lnNumCuotas = 0 Then
        Exit Sub
    End If

    Set oNAdeud = New NCajaAdeudados
    'ALPA 20130614************************
    'lnCapitalCuota = Round(CCur(lnCapitalIni) / (lnNumCuotas - IIf(lnCuotaK = 0, 0, lnCuotaK - 1)), 2)
    lnCapitalCuota = Round(CCur(lnCapitalIni + pnCapitalConcesional) / (lnNumCuotas - IIf(lnCuotaK = 0, 0, lnCuotaK - 1)), 2)
    '*************************************
    ReDim ltCalendario(lnNumCuotas)
    ReDim ltcalendarioC(0)
    'ALPA 20130614****************************
    lnCapital = lnCapitalIni
    lnSumaCapitalConc = pnCapitalConcesional
    lnSumaCapital = lnCapitalIni
    '*****************************************
    ltCalendario(0).Interes = 0
    ltCalendario(0).Cuota = -1
    ltCalendario(0).Total = 0
    ltCalendario(0).Capital = 0
    ltCalendario(0).Saldo = 0
    'ALPA 20130614*******************
    ltCalendario(0).InteresConc = 0
    ltCalendario(0).CapitalConc = 0
    '********************************
    For I = 1 To UBound(ltCalendario)
        If I = 1 Then
            If lnPerGracia > 0 Then
                ltCalendario(I).FechaPago = DateAdd("d", CDbl(lnPerGracia), ldFechaDesemb)
                lnDiasTrans = lnPerGracia
            Else
                ltCalendario(I).FechaPago = DateAdd("D", Val(lnPlazoCuotas), ldFechaDesemb)
                lnDiasTrans = DateDiff("d", ldFechaDesemb, ltCalendario(I).FechaPago)
            End If
        Else
            ltCalendario(I).FechaPago = DateAdd("d", Val(lnPlazoCuotas), ltCalendario(I - 1).FechaPago)
            lnDiasTrans = DateDiff("d", ltCalendario(I - 1).FechaPago, ltCalendario(I).FechaPago)
        End If
        If I >= lnCuotaK Then
            'If (lnCapital - lnCapitalCuota) < 1 Then
            If ((lnSumaCapitalConc + lnSumaCapital) - lnCapitalCuota) < 1 Then
                lnCapitalCuota = lnSumaCapitalConc + lnSumaCapital 'lnCapital
            End If
            'ALPA 20130614**********************
            If I = UBound(ltCalendario) Then
                ltCalendario(I).Capital = lnCapitalIni - lnAcuCapital
                ltCalendario(I).CapitalConc = pnCapitalConcesional - lnAcuCapitalConc
            Else
                ltCalendario(I).Capital = Round(lnCapitalCuota * (1 - lnPorCentaje), 2)
                ltCalendario(I).CapitalConc = Round(lnCapitalCuota - ltCalendario(I).Capital, 2)
            End If
            '***********************************
        End If

        
        ltCalendario(I).Cuota = I
        'ALPA 20130613******************************
        'ltCalendario(I).Interes = Format(oNAdeud.CalculaInteres(lnDiasTrans, lnPeriodo, CCur(lnInteres), lnCapital), "#,#0.00")
        ltCalendario(I).Interes = Format(oNAdeud.CalculaInteres(lnDiasTrans, lnPeriodo, CCur(lnInteres), lnSumaCapital), "#,#0.00")
        ltCalendario(I).InteresConc = Format(oNAdeud.CalculaInteres(lnDiasTrans, lnPeriodo, CCur(lnInteres), lnSumaCapitalConc), "#,#0.00")
        '*******************************************
         If pnComisionCuota <> 0 Then
            If I = 1 Then
                ltCalendario(I).nComisionCuota = MontoIntPerDias(pnComisionCuota, DateDiff("d", ldFechaDesemb, ltCalendario(I).FechaPago), lnCapital, CInt(lnPeriodo))
            Else
                ltCalendario(I).nComisionCuota = MontoIntPerDias(pnComisionCuota, DateDiff("d", ltCalendario(I - 1).FechaPago, ltCalendario(I).FechaPago), lnCapital, CInt(lnPeriodo))
            End If
        End If
        'ALPA20130905*********
        If pnComisionCuotaConc <> 0 Then
            If I = 1 Then
                ltCalendario(I).nComisionCuotaConc = MontoIntPerDias(pnComisionCuotaConc, DateDiff("d", ldFechaDesemb, ltCalendario(I).FechaPago), pnCapitalConcesional, CInt(lnPeriodo))
            Else
                ltCalendario(I).nComisionCuotaConc = MontoIntPerDias(pnComisionCuotaConc, DateDiff("d", ltCalendario(I - 1).FechaPago, ltCalendario(I).FechaPago), pnCapitalConcesional, CInt(lnPeriodo))
            End If
        End If
        '*********************
        'ALPA 20130614******
        ltCalendario(I).Total = ltCalendario(I).Capital + ltCalendario(I).Interes + ltCalendario(I).nComisionCuota + ltCalendario(I).InteresConc + ltCalendario(I).CapitalConc
        '*******************
        ltCalendario(I).cTpoCuota = gCGTipoCuotCalIFCuota
        ltCalendario(I).cEstado = gTpoEstCuotaAdeudPend
        ltCalendario(I).nDiasPago = 0
        ltCalendario(I).nInteresPagado = 0
        ltCalendario(I).cMovNro = ""
        If I >= lnCuotaK Then
            'ALPA 20130614*********************************************************
            lnCapital = lnCapital - Format(lnCapitalCuota, "#,0.00")
            lnSumaCapitalConc = lnSumaCapitalConc - ltCalendario(I).CapitalConc
            lnSumaCapital = lnSumaCapital - ltCalendario(I).Capital
            lnAcuCapitalConc = lnAcuCapitalConc + ltCalendario(I).CapitalConc
            lnAcuCapital = lnAcuCapital + ltCalendario(I).Capital
            '**********************************************************************
        End If
        ltCalendario(I).Saldo = lnSumaCapitalConc + lnSumaCapital '- lnCapitalCuota
    Next I
    MuestraDatos fg
    Exit Sub
ErrorGeneraCal:
    Call RaiseError(MyUnhandledError, "DCaja_Adeudados:CargaIndiceVAC Method")
End Sub

Public Function CuotaFija(ByVal pnTasa As Double, ByVal pnCuotas As Integer, ByVal pnMonto As Double, ByVal pnPeriodo As Double, ByVal pnPeriodoInt As Integer) As Double
Dim Pot1 As Double
Dim nTasaTmp As Double
    nTasaTmp = TasaIntPerDias(pnTasa, pnPeriodo, pnPeriodoInt)
'Obtengo la cuota de pago
    Pot1 = (1 + nTasaTmp) ^ pnCuotas
    CuotaFija = ((Pot1 * nTasaTmp) / (Pot1 - 1)) * pnMonto
    CuotaFija = CDbl(Format(CuotaFija, "#0.00"))
End Function

'***********************************************
'*  FUNCION QUE HALLA EL INTERES DE UN PERIODO DE DIAS TRANACURRIDOS
'***********************************************
Public Function TasaIntPerDias(ByVal pnTasaInter As Double, ByVal pnDiasTrans As Integer, Optional pnPeriodoInt As Integer = 30) As Double
    TasaIntPerDias = ((1 + pnTasaInter / 100) ^ (pnDiasTrans / pnPeriodoInt)) - 1
End Function

Public Function MontoIntPerDias(ByVal pnTasaInter As Double, ByVal pnDiasTrans As Integer, ByVal pnMonto As Double, Optional pnPeriodoInt As Integer = 30) As Double
    MontoIntPerDias = (((1 + pnTasaInter / 100) ^ (pnDiasTrans / pnPeriodoInt)) - 1) * pnMonto
    MontoIntPerDias = CDbl(Format(MontoIntPerDias, "#0.00"))
End Function

Public Sub GeneraCalendarioCuotaFija(fg As Object, ByVal pnMonto As Double, ByRef pnMontoCuota As Double, ByVal pnTasaInt As Double, ByVal pnPeriodoInt As Integer, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, _
                Optional ByVal pnComisionCuota As Currency = 0, Optional ByVal pnTramoConcesional As Currency = 0, _
                Optional ByVal pnCapitalConcesional As Currency = 0, Optional ByVal pnComisionCuotaConc As Currency = 0)
' Se agregó la variable pnCapitalConcesional ALPA 20130614
                
Dim nSaldoCapital As Double
'Dim nSaldoCapitalConcesional As Double
Dim dDesembolso As Date
Dim I As Integer
Dim dFecTemp As Date
Dim nMes As Integer
Dim nAnio As Integer
Dim nDia As Integer
Dim nMontoCuotaTemp As Double
Dim nCDIntPend As Double
Dim nCDIntPendConc As Double
Dim oFun As New NContFunciones
Dim lnContC As Integer
Dim lnSaldoC As Double
Dim lnMontoCuotaC As Double
Dim lnPorCentaje As Double
ReDim ltcalendarioC(0)
lnContC = 0
'ALPA 20130614***************
lnPorCentaje = Round((pnCapitalConcesional / pnMonto), 4)
pnMonto = pnMonto - pnCapitalConcesional
'****************************

Dim nTotalNOConcesional As Integer
nTotalNOConcesional = Int(pnNroCuotas * pnTramoConcesional / 100)

    If pnTramoConcesional <> 0 Then
        nSaldoCapital = Round(pnMonto * pnTramoConcesional / 100, 2)
        lnSaldoC = pnMonto - pnCapitalConcesional 'Round(pnMonto * pnTramoConcesional / 100, 2)
        lnMontoCuotaC = CuotaFija(pnTasaInt, Int(pnNroCuotas / 6), lnSaldoC, pnPeriodo * 6, pnPeriodoInt)
    Else
        nSaldoCapital = pnMonto
        lnSaldoC = pnCapitalConcesional
    End If
    dDesembolso = pdFecDesemb + pnDiasGracia
    ReDim ltCalendario(pnNroCuotas)
    If pnTipoPeriodo = PeriodoFijo Then
        For I = 1 To pnNroCuotas
            ltCalendario(I).FechaPago = dDesembolso + pnPeriodo
            ltCalendario(I).Cuota = I
            ltCalendario(I).Interes = MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital, pnPeriodoInt)
            'ALPA 20130614****************
            ltCalendario(I).InteresConc = MontoIntPerDias(pnTasaInt, pnPeriodo, lnSaldoC, pnPeriodoInt)
            '*****************************
            If pnMontoCuota <> 0 Then
                pnMontoCuota = CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pnPeriodoInt)
                pnMontoCuota = pnMontoCuota + CuotaFija(pnTasaInt, pnNroCuotas, pnCapitalConcesional, pnPeriodo, pnPeriodoInt)
            End If
            ltCalendario(I).Total = pnMontoCuota

            If I = pnNroCuotas Then
                ltCalendario(I).Capital = nSaldoCapital
                ltCalendario(I).CapitalConc = lnSaldoC
                If (ltCalendario(I).Interes + ltCalendario(I).Capital + ltCalendario(I).InteresConc + ltCalendario(I).CapitalConc) > ltCalendario(I).Total Then
                    ltCalendario(I).Total = ltCalendario(I).Capital + ltCalendario(I).Interes + ltCalendario(I).InteresConc + ltCalendario(I).CapitalConc
                Else
                    ltCalendario(I).Interes = ltCalendario(I).Total - (ltCalendario(I).Capital + ltCalendario(I).CapitalConc)
                End If
            Else
                
                ltCalendario(I).Capital = Round((1 - lnPorCentaje) * (ltCalendario(I).Total - (ltCalendario(I).Interes + ltCalendario(I).InteresConc + ltCalendario(I).CapitalConc)), 2)
                ltCalendario(I).CapitalConc = ltCalendario(I).Total - (ltCalendario(I).Interes + ltCalendario(I).InteresConc + ltCalendario(I).Capital)
            End If
            If pnComisionCuota <> 0 Then
                If I = 1 Then
                    ltCalendario(I).nComisionCuota = MontoIntPerDias(pnComisionCuota, DateDiff("d", dDesembolso, ltCalendario(I).FechaPago), nSaldoCapital, pnPeriodoInt)
                Else
                    ltCalendario(I).nComisionCuota = MontoIntPerDias(pnComisionCuota, DateDiff("d", ltCalendario(I - 1).FechaPago, ltCalendario(I).FechaPago), nSaldoCapital, pnPeriodoInt)
                End If
            End If
            
            If pnComisionCuotaConc <> 0 Then
                If I = 1 Then
                    ltCalendario(I).nComisionCuotaConc = MontoIntPerDias(pnComisionCuotaConc, DateDiff("d", dDesembolso, ltCalendario(I).FechaPago), lnSaldoC, pnPeriodoInt)
                Else
                    ltCalendario(I).nComisionCuotaConc = MontoIntPerDias(pnComisionCuotaConc, DateDiff("d", ltCalendario(I - 1).FechaPago, ltCalendario(I).FechaPago), lnSaldoC, pnPeriodoInt)
                End If
            End If
            
            ltCalendario(I).Saldo = nSaldoCapital + lnSaldoC - (ltCalendario(I).Capital + ltCalendario(I).CapitalConc)
            
            'If i <= nTotalNOConcesional Then
            '    ltCalendario(i).cTpoCuota = gCGTipoCuotCalIFNoConcesional
            'Else
                'ltCalendario(i).cTpoCuota = gCGTipoCuotCalIFCuota
            'End If
            
            ltCalendario(I).cTpoCuota = gCGTipoCuotCalIFCuota
            ltCalendario(I).cEstado = gTpoEstCuotaAdeudPend
            ltCalendario(I).nDiasPago = 0
            ltCalendario(I).nInteresPagado = 0
            ltCalendario(I).cMovNro = ""
            
            nSaldoCapital = nSaldoCapital - (ltCalendario(I).Capital)
            lnSaldoC = lnSaldoC - (ltCalendario(I).CapitalConc)
            nSaldoCapital = CDbl(Format(nSaldoCapital, "#0.00"))
            lnSaldoC = CDbl(Format(lnSaldoC, "#0.00"))
            dDesembolso = dDesembolso + pnPeriodo
        Next I
    Else
        If pbCuotaFijaFechaFija Then
            If pnMontoCuota = 0 Then
                nMontoCuotaTemp = CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes)
            Else
                nMontoCuotaTemp = pnMontoCuota
            End If
        End If
        'Si es Fecha Fija
        If pnTipoPeriodo = FechaFija Then
            nMes = Month(dDesembolso)
            nAnio = Year(dDesembolso)
            nDia = pnDiaFijo
            nCDIntPend = 0
            nCDIntPendConc = 0
            For I = 1 To pnNroCuotas
                nDia = pnDiaFijo
                If Not (I = 1 And pnDiaFijo > Day(dDesembolso) And (Not bProxMes)) Then
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
                            If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
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
                        If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
                            nDia = 30
                        End If
                    End If
                End If
                dFecTemp = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)))
                Do While True
                    'Vemos si es sábado o Domingo
                    If Weekday(dFecTemp, vbMonday) = 6 Or Weekday(dFecTemp, vbMonday) = 7 Then
                        dFecTemp = dFecTemp - 1
                    Else
                        If oFun.EsFeriado(dFecTemp) Then
                            dFecTemp = dFecTemp - 1
                        Else
                            Exit Do
                        End If
                    End If
                Loop
                
                ltCalendario(I).FechaPago = dFecTemp
                ltCalendario(I).Cuota = I
                ltCalendario(I).Interes = 0#
                If I = 1 Then
                    ltCalendario(I).Interes = MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, ltCalendario(I).FechaPago), nSaldoCapital, pnPeriodoInt)
                    'ALPA 20130614***********
                    ltCalendario(I).InteresConc = MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, ltCalendario(I).FechaPago), lnSaldoC, pnPeriodoInt)
                    '************************
                    If pnComisionCuota <> 0 Then
                        ltCalendario(I).nComisionCuota = MontoIntPerDias(pnComisionCuota, DateDiff("d", dDesembolso, ltCalendario(I).FechaPago), nSaldoCapital + lnSaldoC, pnPeriodoInt)
                    End If
                    If Not pbCuotaFijaFechaFija Then
                        If pnMontoCuota = 0 Then  'Si el usuario definio la Cuota
                            ltCalendario(I).Total = CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, DateDiff("d", dDesembolso, ltCalendario(I).FechaPago), pnPeriodoInt)
                        Else
                            ltCalendario(I).Total = pnMontoCuota
                        End If
                    Else
                        ltCalendario(I).Total = nMontoCuotaTemp
                    End If
                    
                Else
                    ltCalendario(I).Interes = MontoIntPerDias(pnTasaInt, DateDiff("d", ltCalendario(I - 1).FechaPago, ltCalendario(I).FechaPago), nSaldoCapital, pnPeriodoInt)
                    'ALPA 20130614***********
                    ltCalendario(I).InteresConc = MontoIntPerDias(pnTasaInt, DateDiff("d", ltCalendario(I - 1).FechaPago, ltCalendario(I).FechaPago), lnSaldoC, pnPeriodoInt)
                    '************************
                    If pnComisionCuota <> 0 Then
                        ltCalendario(I).nComisionCuota = MontoIntPerDias(pnComisionCuota, DateDiff("d", ltCalendario(I - 1).FechaPago, ltCalendario(I).FechaPago), nSaldoCapital + lnSaldoC, pnPeriodoInt)
                    End If
                    If Not pbCuotaFijaFechaFija Then
                        ltCalendario(I).Total = CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, DateDiff("d", ltCalendario(I - 1).FechaPago, ltCalendario(I).FechaPago), pnPeriodoInt)
                    Else
                        ltCalendario(I).Total = nMontoCuotaTemp
                    End If
                End If
                
                If pbCuotaFijaFechaFija Then
                    ltCalendario(I).Interes = ltCalendario(I).Interes + nCDIntPend
                    ltCalendario(I).InteresConc = ltCalendario(I).InteresConc + nCDIntPendConc
                    nCDIntPend = 0#
                    nCDIntPendConc = 0#
                    If ltCalendario(I).Interes > (nMontoCuotaTemp - ltCalendario(I).InteresConc - ltCalendario(I).CapitalConc) Then
                        nCDIntPend = ltCalendario(I).Interes - (nMontoCuotaTemp - ltCalendario(I).InteresConc - ltCalendario(I).CapitalConc)
                        ltCalendario(I).Interes = (nMontoCuotaTemp - ltCalendario(I).InteresConc - ltCalendario(I).CapitalConc)
                    End If
                    If (ltCalendario(I).InteresConc > (nMontoCuotaTemp - ltCalendario(I).Interes - ltCalendario(I).CapitalConc)) And pnCapitalConcesional > 0 Then
                        nCDIntPendConc = ltCalendario(I).InteresConc - (nMontoCuotaTemp - ltCalendario(I).Interes - ltCalendario(I).CapitalConc)
                        ltCalendario(I).InteresConc = (nMontoCuotaTemp - ltCalendario(I).Interes - ltCalendario(I).CapitalConc)
                    End If
                End If

                If I = pnNroCuotas Then
                    ltCalendario(I).Capital = nSaldoCapital
                    ltCalendario(I).CapitalConc = lnSaldoC
                    If (ltCalendario(I).Interes + ltCalendario(I).Capital + ltCalendario(I).InteresConc + ltCalendario(I).CapitalConc) > ltCalendario(I).Total Then
                        ltCalendario(I).Total = ltCalendario(I).Capital + ltCalendario(I).Interes + ltCalendario(I).InteresConc + ltCalendario(I).CapitalConc
                    Else
                        ltCalendario(I).Interes = ltCalendario(I).Total - (ltCalendario(I).Capital + ltCalendario(I).InteresConc + ltCalendario(I).CapitalConc)
                        ltCalendario(I).InteresConc = ltCalendario(I).Total - (ltCalendario(I).Capital + ltCalendario(I).Interes + ltCalendario(I).CapitalConc)
                    End If
                Else
                    ltCalendario(I).Capital = ltCalendario(I).Total - (ltCalendario(I).Interes + ltCalendario(I).InteresConc + ltCalendario(I).CapitalConc) - ltCalendario(I).nComisionCuota
                    ltCalendario(I).CapitalConc = ltCalendario(I).Total - (ltCalendario(I).Interes + ltCalendario(I).InteresConc + ltCalendario(I).Capital) - ltCalendario(I).nComisionCuota
                End If
                
                ltCalendario(I).Saldo = nSaldoCapital + lnSaldoC - (ltCalendario(I).Capital + ltCalendario(I).CapitalConc)
                nSaldoCapital = nSaldoCapital - ltCalendario(I).Capital
                lnSaldoC = lnSaldoC - ltCalendario(I).CapitalConc
                ltCalendario(I).cTpoCuota = gCGTipoCuotCalIFCuota
                ltCalendario(I).cEstado = gTpoEstCuotaAdeudPend
                ltCalendario(I).nDiasPago = 0
                ltCalendario(I).nInteresPagado = 0
                ltCalendario(I).cMovNro = ""
                
                
                If pnTramoConcesional <> 0 Then
                    If I Mod 6 = 0 Then
                        lnContC = lnContC + 1
                        ReDim Preserve ltcalendarioC(lnContC)
                        ltcalendarioC(lnContC).Cuota = lnContC
                        ltcalendarioC(lnContC).FechaPago = ltCalendario(I).FechaPago
                        ltcalendarioC(lnContC).Total = lnMontoCuotaC
                        If lnContC = 1 Then
                            ltcalendarioC(lnContC).Interes = MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, ltcalendarioC(lnContC).FechaPago), lnSaldoC, pnPeriodoInt)
                            If pnComisionCuota <> 0 Then
                                ltcalendarioC(lnContC).nComisionCuota = MontoIntPerDias(pnComisionCuota, DateDiff("d", dDesembolso, ltcalendarioC(lnContC).FechaPago), lnSaldoC, pnPeriodoInt)
                            End If
                        Else
                            ltcalendarioC(lnContC).Interes = MontoIntPerDias(pnTasaInt, DateDiff("d", ltcalendarioC(lnContC - 1).FechaPago, ltcalendarioC(lnContC).FechaPago), lnSaldoC, pnPeriodoInt)
                            If pnComisionCuota <> 0 Then
                                ltcalendarioC(lnContC).nComisionCuota = MontoIntPerDias(pnComisionCuota, DateDiff("d", ltcalendarioC(lnContC - 1).FechaPago, ltcalendarioC(lnContC).FechaPago), lnSaldoC, pnPeriodoInt)
                            End If
                        End If
                        ltcalendarioC(lnContC).Capital = ltcalendarioC(lnContC).Total - ltcalendarioC(lnContC).Interes - ltcalendarioC(lnContC).nComisionCuota
                        ltcalendarioC(lnContC).Saldo = lnSaldoC - ltcalendarioC(lnContC).Capital
                        lnSaldoC = lnSaldoC - ltcalendarioC(lnContC).Capital
                        ltcalendarioC(lnContC).cTpoCuota = gCGTipoCuotCalIFNoConcesional
                        ltcalendarioC(lnContC).cEstado = gTpoEstCuotaAdeudPend
                        ltcalendarioC(lnContC).nDiasPago = 0
                        ltcalendarioC(lnContC).nInteresPagado = 0
                        ltcalendarioC(lnContC).cMovNro = ""
                    End If
                End If
                
            Next I
        End If
    End If
    'Actualizar si existe Periodo de Gracia
    If pnDiasGracia > 0 Then
        If pnTipoGracia = PrimeraCuota Then
            ltCalendario(1).Interes = ltCalendario(1).Interes + MontoIntPerDias(pnTasaInt, pnDiasGracia, pnMonto, pnPeriodoInt)
            ltCalendario(1).InteresConc = ltCalendario(1).InteresConc + MontoIntPerDias(pnTasaInt, pnDiasGracia, pnCapitalConcesional, pnPeriodoInt)
        End If
        If pnTipoGracia = Prorateada Or pnTipoGracia = Configurable Then
            For I = 1 To pnNroCuotas
                ltCalendario(I).Interes = MatGracia(I)
                ltCalendario(I).Total = ltCalendario(I).Total + MatGracia(I)
            Next I
        End If
        If pnTipoGracia = UltimaCuota Then
            ReDim Preserve ltCalendario(pnNroCuotas + 1)
            ltCalendario(pnNroCuotas).FechaPago = ltCalendario(pnNroCuotas - 1).FechaPago + pnDiasGracia
            ltCalendario(pnNroCuotas).Cuota = pnNroCuotas + 1
            ltCalendario(pnNroCuotas).Interes = MatGracia(pnNroCuotas)
            ltCalendario(pnNroCuotas).Total = ltCalendario(pnNroCuotas).Interes
            ltCalendario(pnNroCuotas).Capital = 0#
            ltCalendario(pnNroCuotas).Saldo = 0#
            ltCalendario(pnNroCuotas).cTpoCuota = gCGTipoCuotCalIFCuota
            ltCalendario(pnNroCuotas).cEstado = gTpoEstCuotaAdeudPend
            ltCalendario(pnNroCuotas).nDiasPago = 0
            ltCalendario(pnNroCuotas).nInteresPagado = 0
            ltCalendario(pnNroCuotas).cMovNro = ""
            
            ltCalendario(pnNroCuotas).InteresConc = MatGracia(pnNroCuotas)
            ltCalendario(pnNroCuotas).CapitalConc = 0#
            ltCalendario(pnNroCuotas).nComisionCuotaConc = 0# 'ALPA20130905
            
        End If
    End If
    
    MuestraDatos fg
End Sub

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
    dFecTemp = DateAdd("m", 1, CDate("01/" & Format(nMes, "00") & "/" & Format(nAnio, "0000"))) - 1
    
'    dFecTemp = CDate(Format(Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Trim(Str(nAnio)), "dd/mm/yyyy"))
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
 
Private Function MuestraDatos(fgCronograma As Object)
    Dim I As Integer
    Dim K As Integer
    Dim Inicio As Integer

    vnTotalCapital = 0
    vnTotalInteres = 0
    'ALPA 20130614***********
    vnTotalCapitalC = 0
    vnTotalInteresC = 0
    '************************
    'NAGL 20170519
    vnTotalComision = 0
    vnTotalComisionC = 0
    '************************
    vnTotalGeneral = 0
    If UBound(ltCalendario) > 0 Then
        fgCronograma.Rows = UBound(ltCalendario) + 1
        Inicio = 1
        For I = Inicio To UBound(ltCalendario)
'            If ltCalendario(i).Interes > 0 Then
                fgCronograma.TextMatrix(I, 0) = "-"
                fgCronograma.TextMatrix(I, 1) = ltCalendario(I).FechaPago
                fgCronograma.TextMatrix(I, 2) = Format(ltCalendario(I).Cuota, "#0")
                fgCronograma.TextMatrix(I, 3) = Format(ltCalendario(I).Capital, "#,#0.00")
'                If I > 0 Then
'                    vnTotalCapital = vnTotalCapital + ltCalendario(I).Capital
'                End If
                fgCronograma.TextMatrix(I, 4) = Format(ltCalendario(I).Interes, "#,#0.00")
                'vnTotalInteres = CCur(vnTotalInteres) + ltCalendario(I).Interes
                fgCronograma.TextMatrix(I, 5) = Format(ltCalendario(I).nComisionCuota, "#,#0.00")
                fgCronograma.TextMatrix(I, 6) = Format(ltCalendario(I).Total, "#,#0.00")
                fgCronograma.TextMatrix(I, 7) = Format(ltCalendario(I).Saldo, "#,#0.00")
                
                fgCronograma.TextMatrix(I, 8) = ltCalendario(I).cTpoCuota
                fgCronograma.TextMatrix(I, 9) = ltCalendario(I).cEstado
                
                fgCronograma.TextMatrix(I, 10) = ltCalendario(I).nDiasPago
                fgCronograma.TextMatrix(I, 11) = ltCalendario(I).nInteresPagado
                fgCronograma.TextMatrix(I, 12) = ltCalendario(I).cMovNro
                fgCronograma.TextMatrix(I, 13) = Format(ltCalendario(I).CapitalConc, "#,#0.00")
                fgCronograma.TextMatrix(I, 14) = Format(ltCalendario(I).InteresConc, "#,#0.00")
                fgCronograma.TextMatrix(I, 15) = Format(ltCalendario(I).nComisionCuotaConc, "#,#0.00") 'ALPA20130904
                'ALPA 20130614******
                If I > 0 Then
                    vnTotalCapital = vnTotalCapital + ltCalendario(I).Capital
                    vnTotalCapitalC = vnTotalCapitalC + ltCalendario(I).CapitalConc
                End If
                vnTotalInteres = CCur(vnTotalInteres) + ltCalendario(I).Interes
                vnTotalInteresC = CCur(vnTotalInteresC) + ltCalendario(I).InteresConc
                'ALPA 20130614********************************
                ''********************************************
                vnTotalComision = CCur(vnTotalComision) + ltCalendario(I).nComisionCuota
                vnTotalComisionC = CCur(vnTotalComisioC) + ltCalendario(I).nComisionCuotaConc
                'NAGL 20170519********************************
                
'                fgCronograma.TextMatrix(I, 15) = Format(ltCalendario(I).Total, "#,#0.00")
'                fgCronograma.TextMatrix(I, 16) = Format(ltCalendario(I).Saldo, "#,#0.00")
                fgCronograma.TextMatrix(I, 16) = Format(ltCalendario(I).Total, "#,#0.00")
                fgCronograma.TextMatrix(I, 17) = Format(ltCalendario(I).Saldo, "#,#0.00")

                '*********************************************
                '*******************
'            End If
        Next I
        vnTotalGeneral = vnTotalCapital + vnTotalInteres + vnTotalCapitalC + vnTotalInteresC
        fgCronograma.Row = 1
        fgCronograma.Col = 1
    End If

    
'    If UBound(ltcalendarioC) > 0 Then
'        Inicio = fgCronograma.Rows - 1
'        fgCronograma.Rows = fgCronograma.Rows + UBound(ltcalendarioC)
'        For I = 1 To UBound(ltcalendarioC)
'            If ltcalendarioC(I).Interes > 0 Then
'                fgCronograma.Row = I + Inicio
'                For K = 1 To fgCronograma.Cols - 1
'                    fgCronograma.Col = K
'                    fgCronograma.CellBackColor = "&H00E0E0E0"
'                Next
'                fgCronograma.TextMatrix(I + Inicio, 0) = "-"
'                fgCronograma.TextMatrix(I + Inicio, 1) = ltcalendarioC(I).FechaPago
'                fgCronograma.TextMatrix(I + Inicio, 2) = Format(ltcalendarioC(I).Cuota, "#0")
'                fgCronograma.TextMatrix(I + Inicio, 3) = Format(ltcalendarioC(I).Capital, "#,#0.00")
'                fgCronograma.TextMatrix(I + Inicio, 4) = Format(ltcalendarioC(I).Interes, "#,#0.00")
'                fgCronograma.TextMatrix(I + Inicio, 5) = Format(ltcalendarioC(I).nComisionCuota, "#,#0.00")
'                fgCronograma.TextMatrix(I + Inicio, 6) = Format(ltcalendarioC(I).Total, "#,#0.00")
'                fgCronograma.TextMatrix(I + Inicio, 7) = Format(ltcalendarioC(I).Saldo, "#,#0.00")
'                fgCronograma.TextMatrix(I + Inicio, 8) = ltcalendarioC(I).cTpoCuota
'                fgCronograma.TextMatrix(I + Inicio, 9) = ltcalendarioC(I).cEstado
'                fgCronograma.TextMatrix(I + Inicio, 10) = ltcalendarioC(I).nDiasPago
'                fgCronograma.TextMatrix(I + Inicio, 11) = ltcalendarioC(I).nInteresPagado
'                fgCronograma.TextMatrix(I + Inicio, 12) = ltcalendarioC(I).cMovNro
'
'                fgCronograma.TextMatrix(I + Inicio, 13) = Format(ltCalendario(I).CapitalConc, "#,#0.00")
'                fgCronograma.TextMatrix(I + Inicio, 14) = Format(ltCalendario(I).InteresConc, "#,#0.00")
'                fgCronograma.TextMatrix(I + Inicio, 15) = Format(ltCalendario(I).nComisionCuotaConc, "#,#0.00") 'ALPA20130904
'
'                'ALPA 20130614********************************
''                fgCronograma.TextMatrix(I + Inicio, 15) = Format(ltCalendario(I).Total, "#,#0.00")
''                fgCronograma.TextMatrix(I + Inicio, 16) = Format(ltCalendario(I).Saldo, "#,#0.00")
'                fgCronograma.TextMatrix(I + Inicio, 16) = Format(ltCalendario(I).Total, "#,#0.00")
'                fgCronograma.TextMatrix(I + Inicio, 17) = Format(ltCalendario(I).Saldo, "#,#0.00")
'                '*********************************************
'            End If
'        Next I
'        fgCronograma.Row = 1
'        fgCronograma.Col = 1
'    End If
End Function

Public Function GetCalendarioDatos(psPersCod As String, psIFTpo As String, psCtaIFCod As String, Optional pnEstadoCuota As CGTipoEstCuotaAdeud = -1, Optional pbConMov As Boolean = False, Optional bOrdencTpoCuota As String = "ASC") As ADODB.Recordset
    Dim oConect As DConecta
    Dim sql As String
    
    sql = " Select c.cPersCod, c.cCtaIFCod, c.cTpoCuota, c.nNroCuota, C.dVencimiento, "
    sql = sql & "    C.nInteresPagado, nCapital, nInteres, cEstado, nComision, nDiasPago, cMovNro "
    sql = sql & "From    CtaIfCalendario C "
    sql = sql & "Where  cIFTpo = '" & Left(psIFTpo, 2) & "' and cPersCod = '" & psPersCod & "' and cCtaIfcod = '" & psCtaIFCod & "' "
    If pnEstadoCuota <> -1 Then
        sql = sql & " and cEstado = '" & pnEstadoCuota & "' "
    End If
    If pbConMov Then
        sql = sql & " and cMovNro <> '' "
    End If
    sql = sql & "and bvigente = 1 "
    sql = sql & "Order BY C.cPersCod,C.cCtaIFCod, C.cTpoCuota " & bOrdencTpoCuota & ",C.nNroCuota"
    Set GetCalendarioDatos = oConect
    Set oConect = New DConecta
    If Not oConect.AbreConexion() Then
        Exit Function
    End If
    Set GetCalendarioDatos = oConect.CargaRecordSet(sql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Sub GeneraCalendarioBase(fg As Object, psCtaIFCod As String, pnCapital As Currency, pnTramo As Currency)
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim lnTotalPagar As Currency
    Dim I As Integer
    Dim lnCapitalCuota As Currency
    Dim lnCapital As Currency
    Dim lnDiasTrans As Long
    Dim lnSaldo As Currency
    Dim lnNumCuotas As Long
    Dim lnCapitalIni As Currency
    Dim lnComisionIni As Currency
    Dim oConect As DConecta
    
    Dim nTemp As Integer
    Dim nMontoOrigen As Double
    
    On Error GoTo ErrorGeneraCalBase
    Set oConect = New DConecta
    If Not oConect.AbreConexion() Then
        Exit Sub
    End If
    'ALPA 20110517*************************************
    sql = " Select c.cPersCod, c.cCtaIFCod, c.cTpoCuota, c.nNroCuota, C.dVencimiento, "
    sql = sql & "    C.nInteresPagado, nCapital, nInteres, cEstado, nComision, nDiasPago, cMovNro, "
    sql = sql & "    nCapitalConce=isnull(C.nCapitalConce,0), "
    sql = sql & "    nInteresConce=isnull(C.nInteresConce,0), "
    sql = sql & "    nComisionConce=isnull(C.nComisionConce,0) "
    sql = sql & "From    CtaIfCalendario C "
    sql = sql & "Where C.bVigente=1  and cIFTpo = '" & Left(psCtaIFCod, 2) & "' and cPersCod = '" & Mid(psCtaIFCod, 4, 13) & "' and cCtaIfcod = '" & Mid(psCtaIFCod, 18, 7) & "' "
    sql = sql & "Order BY C.cPersCod,C.cCtaIFCod, C.cTpoCuota DESC, C.nNroCuota"
    '**************************************************
    lnCapital = pnCapital
    If pnTramo <> 0 Then
        lnCapital = Round(pnCapital * pnTramo / 100, 2)
    End If
    Set rs = oConect.CargaRecordSet(sql)
    lnNumCuotas = rs.RecordCount - 1
    ReDim ltCalendario(0)
    ReDim ltcalendarioC(0)
    I = 0
    If Not RSVacio(rs) Then
        nTemp = rs!cTpoCuota
                
        Do While Not rs.EOF
            I = I + 1
            ReDim Preserve ltCalendario(I)
            'ALPA 20130614************************************
            lnCapitalCuota = rs!nCapital + rs!nCapitalConce
            'lnCapitalCuota = rs!nCapital
            '*************************************************
            ltCalendario(I).Cuota = rs!nNroCuota
            ltCalendario(I).FechaPago = Format(rs!dVencimiento, gsFormatoFechaView)
            ltCalendario(I).Interes = Format(rs!nInteres, "#0.00")
            ltCalendario(I).Capital = rs!nCapital
            'ALPA 20130614****************************
            ltCalendario(I).InteresConc = Format(rs!nInteresConce, "#0.00")
            ltCalendario(I).CapitalConc = rs!nCapitalConce
            '*****************************************
            ltCalendario(I).nComisionCuota = Format(rs!nComision, "#0.00")
            ltCalendario(I).Total = ltCalendario(I).Capital + ltCalendario(I).Interes + ltCalendario(I).nComisionCuota
            
            ltCalendario(I).cTpoCuota = rs!cTpoCuota
            ltCalendario(I).cEstado = rs!cEstado
            
            ltCalendario(I).nDiasPago = rs!nDiasPago
            ltCalendario(I).nInteresPagado = rs!nInteresPagado
            ltCalendario(I).cMovNro = rs!cMovNro
            If pnTramo <> 0 And rs!cTpoCuota = gCGTipoCuotCalIFNoConcesional Then
                If rs!nNroCuota = 1 Then
                    lnCapital = pnCapital - Round(pnCapital * pnTramo / 100, 2)
                End If
            End If
             
            If rs!nCapital > 0 Then
                lnCapital = lnCapital - lnCapitalCuota
            End If
                            
            '
            If nMontoOrigen = 0 Then
                nMontoOrigen = lnCapital + lnCapitalCuota
            End If
            
            If nTemp <> rs!cTpoCuota And lnCapital < 0 Then
                nTemp = rs!cTpoCuota
                lnCapital = pnCapital - nMontoOrigen - lnCapitalCuota
            End If
            '
            ltCalendario(I).nComisionCuotaConc = rs!nComisionConce 'ALPA20130905
            ltCalendario(I).Saldo = lnCapital
            rs.MoveNext
        Loop
    End If
    RSClose rs
    oConect.CierraConexion
    Set oConect = Nothing
    MuestraDatos fg

Exit Sub
ErrorGeneraCalBase:
    Call RaiseError(MyUnhandledError, "DCaja_Adeudados:CargaIndiceVAC Method")
End Sub
'*************************************ANGC202012*************************************
Public Sub GenerarCronogramaIterativo(fg As Object, nMonto As Currency, nCuotas As Integer, nTasa As Double, nGracia As Integer, dFecha As String, nTasaComision As Double)
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim nTemp As Integer
    Dim oConect As DConecta
    Dim I As Integer
             
    Set oConect = New DConecta
     If Not oConect.AbreConexion() Then
        Exit Sub
    End If
    
    nTasaComision = nTasaComision / 100#
    dFecha = Format(CDate(dFecha), "yyyymmdd")
    
    sql = "stp_sel_GenerarCalend_Adeudados '109', " & nMonto & "," & nCuotas & "," & nTasa & ", " & nGracia & ",'" & dFecha & "', " & nTasaComision & ",0,0"
    Set rs = oConect.CargaRecordSet(sql)
    
    ReDim ltCalendario(0)
    
    I = 0
    If Not RSVacio(rs) Then
        Do While Not rs.EOF
            I = I + 1
            ReDim Preserve ltCalendario(I)
            
            'lnCapitalCuota = rs!nCapital '+ rs!nCapitalConce
            ltCalendario(I).Cuota = rs!nCuota
            ltCalendario(I).FechaPago = Format(rs!dFecha, gsFormatoFechaView)
            ltCalendario(I).Interes = Format(rs!nInteresCompe, "#0.00")
            ltCalendario(I).Capital = Format(rs!nCapital, "#0.00")
            ltCalendario(I).InteresConc = 0#    'Format(rs!nInteresConce, "#0.00")
            ltCalendario(I).CapitalConc = 0#    'rs!nCapitalConce
            ltCalendario(I).nComisionCuota = Format(rs!Comision, "#0.00")
            ltCalendario(I).Total = ltCalendario(I).Capital + ltCalendario(I).Interes + ltCalendario(I).nComisionCuota
            ltCalendario(I).cTpoCuota = 2   'Tipo Cuota
            ltCalendario(I).cEstado = 0     'rs!cEstado
            ltCalendario(I).nDiasPago = rs!Dias
            ltCalendario(I).nInteresPagado = 0  'rs!nInteresPagado
            ltCalendario(I).cMovNro = ""     'rs!cMovNro

            ltCalendario(I).nComisionCuotaConc = 0  'rs!nComisionConce
            ltCalendario(I).Saldo = rs!nSaldo
            rs.MoveNext
        Loop
    End If
    
    RSClose rs
    oConect.CierraConexion
    Set oConect = Nothing
    MuestraDatos fg
End Sub

Public Property Get TotalCapital() As Currency
    TotalCapital = vnTotalCapital
End Property
'ALPA 20130614
Public Property Get TotalCapitalC() As Currency
    TotalCapitalC = vnTotalCapitalC
End Property
Public Property Let TotalCapital(ByVal vNewValue As Currency)
    vnTotalCapital = vNewValue
End Property
'ALPA 20130614
Public Property Let TotalCapitalC(ByVal vNewValue As Currency)
    vnTotalCapitalC = vNewValue
End Property
Public Property Get TotalInteres() As Currency
    TotalInteres = vnTotalInteres
End Property
'ALPA 20130614
Public Property Get TotalInteresC() As Currency
    TotalInteresC = vnTotalInteresC
End Property
Public Property Let TotalInteres(ByVal vNewValue As Currency)
    vnTotalInteres = vNewValue
End Property
'ALPA 20130614
Public Property Let TotalInteresC(ByVal vNewValue As Currency)
    vnTotalInteresC = vNewValue
End Property
'*****NAGL 20170519
Public Property Get TotalComision() As Currency
    TotalComision = vnTotalComision
End Property
Public Property Get TotalComisionC() As Currency
    TotalComisionC = vnTotalComisionC
End Property
Public Property Let TotalComision(ByVal vNewValue As Currency)
    vnTotalComision = vNewValue
End Property
Public Property Let TotalComisionC(ByVal vNewValue As Currency)
    vnTotalComisionC = vNewValue
End Property '****NAGL 20170519
Public Property Get TotalGeneral() As Currency
    TotalGeneral = vnTotalGeneral
End Property
Public Property Let TotalGeneral(ByVal vNewValue As Currency)
     vnTotalGeneral = vNewValue
End Property
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni

vsCentralPers = oIni.BasePersonas
vsCentralCom = oIni.BaseComunes
vsCentralImg = oIni.BaseImagenes
End Sub
