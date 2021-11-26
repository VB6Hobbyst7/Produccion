VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
Public Function ValidaGrabarDesembolso(ByVal psCtaCod As String, Optional ByVal pnPrestamo As Double, Optional ByVal pbDesembCC As Boolean, Optional ByVal psCtaAbo As String) As String
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
    
    If pbDesembCC Then
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

Public Function DesembolsarCredito(ByRef pnMovNro As Long, ByVal psCtaCod As String, ByVal pnPrestamo As Double, ByVal pnMontoDesembolso As Double, _
        ByVal pMatGastos As Variant, ByVal MatCargoAutom As Variant, ByVal pdHoy As Date, ByVal psCodAge As String, psCodUser As String, _
        ByVal pbDesembCC As Boolean, Optional ByVal pbCtaAhoNueva As Boolean = False, _
        Optional ByVal MatCredCanc As Variant = Nothing, Optional ByVal pnContMatCredCanc As Integer = -1, Optional ByVal psCtaAbo As String = "") As String

Dim R As ADODB.Recordset
Dim oDCredito As DCredito
Dim oFunciones As NContFunciones
Dim oBase As DCredActualizaBD
Dim oLinea As DLineaCredito
Dim oCalend As NCalendario
Dim oDCalend As Dcalendario
Dim oGastos As nGasto
Dim sLineaCod As String
Dim nMontoReservado As Double
Dim nMontoColocado As Double
Dim MatCalendPagos As Variant
Dim MatGracia As Variant
Dim nPeriodoGracia As Integer
Dim nTasaInteres As Double
Dim nNroCalen As Integer
Dim nNroProxDesemb As Integer
Dim I As Integer
Dim MatGastos As Variant
Dim MatDesemb() As String
Dim ContDesemb As Integer
'Dim MatGastosDesemb(1000, 3) As String
'Dim ContGastosDesemb As Integer
Dim nNumGastos As Integer
Dim nNumTransac As Long
Dim sMovNro As String
Dim nMovNro As Long
Dim nConsCred As String
Dim sMetLiquid As String
Dim bTransac As Boolean
'Para Cancelacion de Creditos
Dim MatCalendTemp As Variant
Dim MatCalendDistrTemp As Variant
Dim sMetLiqTmp As String
Dim RTmp As ADODB.Recordset

        On Error GoTo ErrorDesembolsarCredito
        bTransac = False
        'Valida datos antes de Desembolsar
        DesembolsarCredito = ValidaGrabarDesembolso(psCtaCod, pnPrestamo, pbDesembCC, psCtaAbo)
        If DesembolsarCredito <> "" Then
            Exit Function
        End If
        
        If pbDesembCC Then
            If pbCtaAhoNueva Then
                If Mid(psCtaAbo, 4, 2) <> psCodAge Then 'Cuenta de Otra Age
                    nConsCred = gCredDesembCtaNuevaDOA
                Else
                    nConsCred = gCredDesembCtaNueva
                End If
            Else
                If Mid(psCtaAbo, 4, 2) <> psCodAge Then 'Cuenta de Otra Age
                    nConsCred = gCredDesembCtaExistDOA
                Else
                    nConsCred = gCredDesembCtaExist
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
        Set R = oDCredito.RecuperaProducto(psCtaCod)
        nTasaInteres = Format(R!nTasaInteres, "#0.00")
        nNumTransac = IIf(IsNull(R!nTransacc), 0, R!nTransacc)
        R.Close
        Set R = Nothing
        Set R = oDCredito.RecuperaColocacCred(psCtaCod)
        nNroCalen = R!nNroCalen
        nNroProxDesemb = R!nNroProxDesemb
        sMetLiquid = R!cMetLiquidacion
        R.Close
        Set R = Nothing
        Set oDCredito = Nothing
        
        Set oFunciones = New NContFunciones
        sMovNro = oFunciones.GeneraMovNro(pdHoy, psCodAge, psCodUser)
        Set oFunciones = Nothing
                
        'Recupera Calendario de Desembolsos
        ContDesemb = 0
        Set oDCalend = New Dcalendario
        Set R = oDCalend.RecuperaCalendarioDesemb(psCtaCod)
        Do While Not R.EOF
            ContDesemb = ContDesemb + 1
            R.MoveNext
        Loop
        If ContDesemb > 0 Then
            ReDim MatDesemb(ContDesemb, 2)
            ContDesemb = 0
            R.MoveFirst
            Do While Not R.EOF
                ContDesemb = ContDesemb + 1
                MatDesemb(ContDesemb - 1, 0) = Format(R!dVenc, "dd/mm/yyyy")
                MatDesemb(ContDesemb - 1, 1) = Format(R!nCapital, "#0.00")
                R.MoveNext
            Loop
        End If
        R.Close
        Set R = Nothing
        Set oDCalend = Nothing
        
        Set oBase = Nothing
        Set oBase = New DCredActualizaBD
        Call oBase.dBeginTrans
        bTransac = True
        
        'Inserta Movimiento
        Call oBase.dInsertMov(sMovNro, nConsCred, "", gMovEstContabMovContable, gMovFlagVigente, False)
        nMovNro = oBase.dGetnMovNro(sMovNro)
        pnMovNro = nMovNro
        '***********************************************************
        'Proceso de Cancelacion de Creditos
        '**********************************************************
        Set oDCredito = New DCredito
        Set RTmp = New ADODB.Recordset
        For I = 0 To pnContMatCredCanc - 1
            Set RTmp = oDCredito.RecuperaDatosCreditoVigente(MatCredCanc(I, 0))
            sMetLiqTmp = RTmp!cMetLiquidacion
            RTmp.Close
            MatCalendTemp = RecuperaMatrizCalendarioPendiente(MatCredCanc(I, 0))
            MatCalendDistrTemp = MatrizDistribuirCancelacion(MatCredCanc(I, 0), MatCalendTemp, CDbl(MatCredCanc(I, 1)), sMetLiqTmp, pdHoy)
            Call AmortizarCredito(MatCredCanc(I, 0), MatCalendTemp, MatCalendDistrTemp, CDbl(MatCredCanc(I, 1)), pdHoy, sMetLiqTmp, gColocTipoPagoEfectivo, psCodAge, psCodUser, , oBase, nMovNro)
        Next I
        Set oDCredito = Nothing
        '**********************************************************
        'Final de Proceso de Cancelacion
        '**********************************************************
        
        'Actualiza la Saldos de Linea
        Set oLinea = New DLineaCredito
        Set R = oBase.RecuperaLineasCreditoSaldo(sLineaCod, True)
        Set oLinea = Nothing
        nMontoReservado = CDbl(Format(R!nMontoReservado, "#0.00"))
        nMontoColocado = CDbl(Format(R!nMontoColocado, "#0.00"))
        R.Close
        Set R = Nothing
        nMontoReservado = nMontoReservado - pnPrestamo
        nMontoColocado = nMontoColocado + pnPrestamo
        Call oBase.dUpdateLineaCreditoSaldo(sLineaCod, , , nMontoColocado, , nMontoReservado, False)
        
        'Genera Nuevo Calendario
        Set oDCredito = New DCredito
        Set R = oDCredito.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
        Set oDCredito = Nothing
        If R!nColocCalendCod <> gColocCalendCodCL And nNroProxDesemb = 1 Then
            Set oCalend = New NCalendario
            nPeriodoGracia = IIf(IsNull(R!nPeriodoGracia), 0, R!nPeriodoGracia)
            If nPeriodoGracia > 0 Then
                MatGracia = oCalend.GeneraGracia(R!nTipoGracia, CDbl(Format(TasaIntPerDias(nTasaInteres, nPeriodoGracia) * pnPrestamo, "#0.00")), R!nCuotas)
            End If
            
            MatCalendPagos = oCalend.GeneraCalendario(pnPrestamo, nTasaInteres, R!nCuotas, R!nPlazo, _
                pdHoy, DameTipoCuota(R!nColocCalendCod), DameTipoPeriodo(R!nColocCalendCod), _
                IIf(IsNull(R!nTipoGracia), 0, R!nTipoGracia), nPeriodoGracia, IIf(IsNull(R!nPeriodoFechaFija), 0, R!nPeriodoFechaFija), _
                IIf(R!nProxMes = 0, False, True), MatGracia)
            
            Set oCalend = Nothing
            
            Set oGastos = New nGasto
            MatGastos = oGastos.GeneraCalendarioGastos(MatCalendPagos, MatDesemb, nNumGastos, psCtaCod)
            Set oGastos = Nothing
            
            'Nuevo Calendario de Desembolsos
            For I = 0 To ContDesemb - 1
                Call oBase.dInsertColocCalendario(psCtaCod, nNroCalen + 1, gColocCalendAplDesembolso, I + 1, CDate(MatDesemb(I, 0)), gColocCalendEstadoPendiente, "Calendario de Desembolsos -  Desembolso", gColocCalendConceptoProcDesembolsado, False)
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplDesembolso, I + 1, gColocConceptoCodCapital, CDbl(MatDesemb(I, 1)), 0, "", False)
            Next I
            
            'Nuevos Gastos de Desembolsos
            For I = 0 To nNumGastos - 1
                If CInt(Trim(Right(MatGastos(I, 0), 10))) = gColocCalendAplDesembolso Then
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplDesembolso, CInt(Trim(MatGastos(I, 1))), CInt(Trim(Right(MatGastos(I, 2), 20))), CDbl(MatGastos(I, 3)), 0, "", False)
                End If
            Next I
            
            'Nuevo Calendario de Pagos
            For I = 0 To UBound(MatCalendPagos) - 1
                Call oBase.dInsertColocCalendario(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, I + 1, CDate(MatCalendPagos(I, 0)), gColocCalendEstadoPendiente, "Calendario de Pagos - Desembolso", gColocCalendConceptoProcDesembolsado, False)
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, I + 1, gColocConceptoCodCapital, CDbl(MatCalendPagos(I, 3)), 0, "", False)
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalendPagos(I, 4)), 0, "", False)
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", False)
                If CDbl(MatCalendPagos(I, 5)) > 0 Then
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresGracia, CDbl(MatCalendPagos(I, 5)), 0, "", False)
                End If
            Next I
            'Nuevos Gastos de Cuota
            For I = 0 To nNumGastos - 1
                If CInt(Trim(Right(MatGastos(I, 0), 10))) = gColocCalendAplCuota Then
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen + 1, CInt(Trim(Right(MatGastos(I, 0), 3))), CInt(Trim(MatGastos(I, 1))), CInt(Trim(Right(MatGastos(I, 2), 20))), CDbl(MatGastos(I, 3)), 0, "", False)
                End If
            Next I
        End If
        
        'Actualiza Estado
        Call oBase.dInsertColocacEstado(psCtaCod, pdHoy, gColocEstVigNorm, R!nCuotas, pnPrestamo, "Desembolso de Credito", R!nColocCalendCod, IIf(IsNull(R!nPeriodoFechaFija), 0, R!nPeriodoFechaFija), nPeriodoGracia, R!nPlazo, R!nTipoGracia, R!nTipoDesembolso, R!nProxMes, R!nCalendDinamico, False)
        'Actualiza Colocaciones
        Call oBase.dUpdateColocaciones(psCtaCod, CDate(MatCalendPagos(UBound(MatCalendPagos) - 1, 0)) - CDate(MatCalendPagos(0, 0)) _
            , CDate(MatCalendPagos(UBound(MatCalendPagos) - 1, 0)), pnPrestamo, , , pdHoy, False)
        'Actualiza ColocacCred
        Call oBase.dUpdateColocacCred(psCtaCod, 0, , , , IIf(UBound(MatCargoAutom) > 0, 1, 0), , 0, 1, , , , , , nNroCalen + 1, nNroProxDesemb + 1, , , False)
        'Actualiza Producto
        Call oBase.dUpdateProducto(psCtaCod, , pnPrestamo, gColocEstVigNorm, pdHoy, nNumTransac + 1, False)
        R.Close
        Set R = Nothing
        
        'Inserta Movimientos en MovCol
        Call oBase.dInsertMovCol(nMovNro, nConsCred, psCtaCod, nNroCalen + 1, pnPrestamo, 0, sMetLiquid, CDate(MatCalendPagos(UBound(MatCalendPagos) - 1, 0)) - CDate(MatCalendPagos(0, 0)), False, gColocEstAprob)
        
        Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, nNroCalen + 1, gColocConceptoCodCapital, nNroProxDesemb, pnPrestamo, False)
        
        'Ingresa el Monto de Todos los Gastos del desembolso
        For I = 0 To UBound(pMatGastos) - 1
            Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, nNroCalen + 1, CInt(pMatGastos(I, 2)), CInt(pMatGastos(I, 0)), pMatGastos(I, 1), False)
        Next I
                        
        'Ingreso de Movmientos de Cancelaciones de Creditos
        
        'Actualiza Calendario de Desembolsos
        Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplDesembolso, nNroProxDesemb, gColocConceptoCodCapital, , pnPrestamo, , False)
        
        'Actualiza Gastos de Desembolso
        For I = 0 To nNumGastos - 1
            Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen + 1, gColocCalendAplDesembolso, nNroProxDesemb, CInt(Trim(Right(MatGastos(I, 2), 20))), , CDbl(MatGastos(I, 3)), , False)
        Next I
        
        'Cambia a estado Cancelado el Desembolso
        Call oBase.dUpdateColocCalendario(psCtaCod, nNroCalen + 1, nNroProxDesemb, gColocCalendAplDesembolso, , gColocCalendEstadoPagado, , , False)
        
        'Actualiza Cuentas de Ahorro
        '****************** depende de las clases de Raul ****************
        If pbDesembCC Then
            oBase.CapAbonoCuentaAho psCtaAbo, pnMontoDesembolso, gAhoApeEfec, sMovNro, "", , , , , , , pdHoy, ""
        End If
        
        'Actualiza CargoAutomatico
        For I = 0 To UBound(MatCargoAutom) - 1
            Call oBase.dInsertColocCargoAutoma(psCtaCod, I + 1, MatCargoAutom(I), False)
        Next I
        Call oBase.dCommitTrans
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
Dim dTemp As Date
Dim nCuota As Integer
Dim nNroCalen As Integer

    On Error GoTo ErrorValidaCargaDatosDesembolso
    ValidaCargaDatosDesembolso = ""
    
    Set oCredDatos = New DCredito
    Set R = oCredDatos.RecuperaProducto(psCtaCod)
    Set oCredDatos = Nothing
    If R.RecordCount = 0 Then
        ValidaCargaDatosDesembolso = "Credito No ha sido Aprobado"
        R.Close
        Set R = Nothing
        Exit Function
    Else
        If R!nPrdEstado <> gColocEstAprob Then
            ValidaCargaDatosDesembolso = "Credito No ha sido Aprobado"
            R.Close
            Set R = Nothing
            Exit Function
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
    dTemp = CDate(Format(R!dVenc, "dd/mm/yyyy"))
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
    
    Exit Function

ErrorValidaCargaDatosDesembolso:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
    
End Function

Public Function SugerenciaCredito(ByVal psCtaCod As String, ByVal pnEstadoActual As ColocEstado, ByVal pnEstado As ColocEstado, ByVal pdEstado As Date, ByVal pnNroTransac As Long, ByVal psLineaCred As String, ByVal pnTasa As Double, ByVal pnMonto As Double, _
        ByVal pnNrocuotas As Integer, ByVal pnPlazo As Integer, ByVal pnTipoCuota As ColocTipoCalend, ByVal pnDiaFijo As Integer, _
        ByVal pnProxMes As Integer, ByVal pnTipoDesemb As ColocTiposDesembolso, ByVal pnCalendDinamico As Integer, _
        ByVal pnGraciaSug As Integer, ByVal pnTasaGracia As Double, ByVal pnTipoGracia As ColocTiposGracia, _
        ByVal nNroCalend As Integer, ByVal MatCalendDes As Variant, ByVal MatCalendPagos As Variant) As String
        
Dim oDCredito As DCredActualizaBD
Dim oLinea As DLineaCredito
Dim R As ADODB.Recordset

Dim I As Integer
    On Error GoTo ErrorSugerenciaCredito
    
    SugerenciaCredito = ValidaGarantia(psCtaCod, pdEstado, pnMonto)
    If SugerenciaCredito <> "" Then
        Exit Function
    End If
    
    Set oDCredito = New DCredActualizaBD
    Call oDCredito.dUpdateProducto(psCtaCod, pnTasa, 0#, pnEstado, pdEstado, pnNroTransac, True)
    
    'Inserta Tasa Compensatoria
    If pnEstadoActual <> pnEstado Then
        Call oDCredito.dInsertProductoTasaInteres(psCtaCod, gColocTasaCompNormal, pnTasa, True)
    Else
        Call oDCredito.dUpdateProductoTasaInteres(psCtaCod, gColocTasaCompNormal, pnTasa, True)
    End If
    'Inserta TasaGracia
    If pnGraciaSug > 0 Then
        Call oDCredito.dDeleteProductoTasaInteres(psCtaCod, gColocTasaCompGraCia)
        Call oDCredito.dInsertProductoTasaInteres(psCtaCod, gColocTasaCompGraCia, pnTasaGracia, True)
    Else
        Call oDCredito.dDeleteProductoTasaInteres(psCtaCod, gColocTasaCompGraCia)
    End If
        
    Call oDCredito.dUpdateColocaciones(psCtaCod, pnPlazo, CDate(MatCalendPagos(UBound(MatCalendPagos) - 1, 0)), 0, pnTipoCuota, psLineaCred, pdEstado, True)
    If pnEstadoActual <> pnEstado Then
        Call oDCredito.dInsertColocacEstado(psCtaCod, pdEstado, pnEstado, pnNrocuotas, pnMonto, "Sugerencia Analista", pnTipoCuota, pnDiaFijo, pnGraciaSug, pnPlazo, pnTipoGracia, pnTipoDesemb, pnProxMes, pnCalendDinamico, True)
    Else
        Call oDCredito.dUpdateColocacEstado(psCtaCod, pdEstado, pnEstado, pnNrocuotas, pnMonto, "Sugerencia Analista", pnTipoCuota, pnDiaFijo, pnGraciaSug, pnPlazo, pnTipoGracia, pnTipoDesemb, pnProxMes, pnCalendDinamico, True)
    End If
    
    Set oLinea = New DLineaCredito
    Set R = oLinea.RecuperaLineadeCredito(psLineaCred)
    Set oLinea = Nothing
    Call oDCredito.dUpdateColocacCred(psCtaCod, , , , , , , , , , , pnTipoCuota, pnCalendDinamico, pnTipoDesemb, nNroCalend, , R!cPersCod, IIf(IsNull(R!nMoneda), CInt(Mid(psLineaCred, 5, 1)), R!nMoneda), True)
    R.Close
    Set R = Nothing
    'Borrado de Calendario
    Call oDCredito.dDeleteColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplDesembolso, , , True)
    Call oDCredito.dDeleteColocCalendario(psCtaCod, nNroCalend, gColocCalendAplDesembolso, , True)

    Call oDCredito.dDeleteColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, , , True)
    Call oDCredito.dDeleteColocCalendario(psCtaCod, nNroCalend, gColocCalendAplCuota, , True)
    
    'Nuevo Calendario
    For I = 0 To UBound(MatCalendDes) - 1
        Call oDCredito.dInsertColocCalendario(psCtaCod, nNroCalend, gColocCalendAplDesembolso, I + 1, CDate(MatCalendDes(I, 0)), gColocCalendEstadoPendiente, "Calendario de Sugerencia Desembolsos", gColocCalendConceptoProcSugerido, True)
        Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplDesembolso, I + 1, gColocConceptoCodCapital, CDbl(MatCalendDes(I, 1)), 0, "", True)
    Next I
    For I = 0 To UBound(MatCalendPagos) - 1
        Call oDCredito.dInsertColocCalendario(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, CDate(MatCalendPagos(I, 0)), gColocCalendEstadoPendiente, "Calendario de Sugerencia Pagos", gColocCalendConceptoProcSugerido, True)
        If pnTipoCuota <> gColocCalendCodCL Then
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, gColocConceptoCodCapital, CDbl(MatCalendPagos(I, 3)), 0, "", True)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalendPagos(I, 4)), 0, "", True)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", True)
            If CDbl(MatCalendPagos(I, 5)) > 0 Then
                Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresGracia, CDbl(MatCalendPagos(I, 5)), 0, "", True)
            End If
        Else
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, gColocConceptoCodCapital, CDbl(MatCalendPagos(I, 3)), 0, "", True)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalendPagos(I, 4)), 0, "", True)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", True)
        End If
    Next I
    oDCredito.dEjecutaBatch
    Set oDCredito = Nothing
    
    Exit Function

ErrorSugerenciaCredito:
    Err.Raise Err.Number, "Error En Proceso SugerenciaCredito", Err.Description
    
End Function
Public Function EsRefinanciado(ByVal psCtaCod As String) As Boolean
Dim oCred As DCredito
Dim R As ADODB.Recordset

    On Error GoTo ErrorEsRefinanciado
    Set oCred = New DCredito
    Set R = oCred.RecuperaColocacRefinanc(psCtaCod)
    Set oCred = Nothing
    If R.RecordCount > 0 Then
        EsRefinanciado = True
    Else
        EsRefinanciado = False
    End If
    Exit Function

ErrorEsRefinanciado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function MontosARefinanciar(ByVal psCtaCod As String, ByVal pdHoy As Date) As Variant
Dim R As ADODB.Recordset
Dim MatCalend As Variant
Dim nMontoInt As Double
Dim nMontoGastos As Double
Dim nMontoCap As Double

    On Error GoTo ErrorMontosARefinanciar
    MatCalend = RecuperaMatrizCalendarioPendiente(psCtaCod)
    nMontoInt = MatrizInteresGastosAFecha(psCtaCod, MatCalend, pdHoy) - MatrizGastosFecha(psCtaCod, MatCalend)
    nMontoGastos = MatrizGastosFecha(psCtaCod, MatCalend)
    nMontoCap = MatrizCapitalAFecha(psCtaCod, MatCalend)
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

Public Function AprobacionCredito(ByVal psCtaCod As String, ByVal pnEstado As ColocEstado, ByVal pdEstado As Date, ByVal pnNroTransac As Long, ByVal psLineaCred As String, ByVal pnTasa As Double, ByVal pnMonto As Double, _
        ByVal pnNrocuotas As Integer, ByVal pnPlazo As Integer, ByVal pnTipoCuota As ColocTipoCalend, ByVal pnDiaFijo As Integer, _
        ByVal pnProxMes As Integer, ByVal pnTipoDesemb As ColocTiposDesembolso, ByVal pnCalendDinamico As Integer, _
        ByVal pnGraciaSug As Integer, ByVal pnTasaGracia As Double, ByVal pnTipoGracia As ColocTiposGracia, _
        ByVal nNroCalend As Integer, ByVal MatCalendDes As Variant, ByVal MatCalendPagos As Variant, ByVal MatGastos As Variant, ByVal pNumGastos As Integer, ByVal psApoderado As String, _
        Optional pbRefinanc As Boolean = False, Optional pbCapitalInter As Boolean = False, Optional ByVal psCodAge As String = "", Optional ByVal psCodUsu As String = "") As String
        
Dim oDCredito As DCredActualizaBD
Dim R As ADODB.Recordset
Dim RRef As ADODB.Recordset
Dim I As Integer
Dim nMontoReservado As Double
Dim bTransac As Boolean
Dim MatCredRef As Variant
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

Dim oGeneral As DGeneral
Dim nEstadCredRef As Integer
Dim sConsCred As String
Dim sConsCred2 As String
Dim nTipoCambioFijo As Double

    On Error GoTo ErrorSugerenciaCredito
    
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
        MatCredRef = RecuperaMatrizRefinanciados(psCtaCod)
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
    Call oDCredito.dInsertColocacEstado(psCtaCod, pdEstado, pnEstado, pnNrocuotas, pnMonto, "Aprobacion de Credito", pnTipoCuota, pnDiaFijo, pnGraciaSug, pnPlazo, pnTipoGracia, pnTipoDesemb, pnProxMes, pnCalendDinamico, False)
    
    Call oDCredito.dUpdateColocacCred(psCtaCod, , , , , , , , , , , pnTipoCuota, pnCalendDinamico, pnTipoDesemb, nNroCalend, 1, , , False)
                    
    'Borrado de Calendario
    'Call oDCredito.dDeleteColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplDesembolso, , , False)
    'Call oDCredito.dDeleteColocCalendario(psCtaCod, nNroCalend, gColocCalendAplDesembolso, , False)

    'Call oDCredito.dDeleteColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, , , False)
    'Call oDCredito.dDeleteColocCalendario(psCtaCod, nNroCalend, gColocCalendAplCuota, , False)
    
    
    'Nuevo Calendario
    'Desembolsos
    For I = 0 To UBound(MatCalendDes) - 1
        Call oDCredito.dInsertColocCalendario(psCtaCod, nNroCalend, gColocCalendAplDesembolso, I + 1, CDate(MatCalendDes(I, 0)), gColocCalendEstadoPendiente, "Calendario de Aprobacion Desembolsos", gColocCalendConceptoProcAprobado, False)
        Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplDesembolso, I + 1, gColocConceptoCodCapital, CDbl(MatCalendDes(I, 1)), 0, "", False)
    Next I
    'Pagos
    For I = 0 To UBound(MatCalendPagos) - 1
        Call oDCredito.dInsertColocCalendario(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, CDate(MatCalendPagos(I, 0)), gColocCalendEstadoPendiente, "Calendario de Aprobacion Pagos", gColocCalendConceptoProcAprobado, False)
        If pnTipoCuota <> gColocCalendCodCL Then
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, gColocConceptoCodCapital, CDbl(MatCalendPagos(I, 3)), 0, "", False)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalendPagos(I, 4)), 0, "", False)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", False)
            If CDbl(MatCalendPagos(I, 5)) > 0 Then
                Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresGracia, CDbl(MatCalendPagos(I, 5)), 0, "", False)
            End If
        Else
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, gColocConceptoCodCapital, CDbl(MatCalendPagos(I, 3)), 0, "", False)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalendPagos(I, 4)), 0, "", False)
            Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", False)
        End If
    Next I
    
    'Gastos
    For I = 0 To pNumGastos - 1
        Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, CInt(Trim(Right(MatGastos(I, 0), 3))), CInt(Trim(MatGastos(I, 1))), CInt(Trim(Right(MatGastos(I, 2), 20))), CDbl(MatGastos(I, 3)), 0, "", False)
    Next I
    
    '*********************************************************
    '************* REFINANCIACION DE CREDITO *****************
    '*********************************************************
    If pbRefinanc Then
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
            
            MatCredRef = RecuperaMatrizCalendarioPendiente(R!cCtaCodRef)
            Call oDCredito.dInsertColocacRefinanc(psCtaCod, R!cCtaCodRef, MatrizInteresGastosAFecha(R!cCtaCodRef, MatCredRef, pdEstado) + MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef), gColocEstadoRefinancAprobado, pdEstado, False)
                                    
            'Inserta en ColocacCredRefinancDet
                Call oDCredito.dInsertColocacRefinancDet(psCtaCod, R!cCtaCodRef, MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef), 0, gColocEstadoRefinancAprobado, gColocConceptoCodCapital, pdEstado, False)
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
                nMontoRefTmp = MatrizInteresGastosAFecha(R!cCtaCodRef, MatCredRef, pdEstado) + MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef) - MatrizGastosFecha(R!cCtaCodRef, MatCredRef)
                nMontoRefTmpCapital = MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef)
                nMontoRefTmpInteres = MatrizInteresGastosAFecha(R!cCtaCodRef, MatCredRef, pdEstado) - MatrizGastosFecha(R!cCtaCodRef, MatCredRef)
                nMontoRefTmpGastosSusp = MatrizGastosFecha(R!cCtaCodRef, MatCredRef)
                                                
                'Insertar Movimientos
                Call oDCredito.dInsertMovCol(nMovNro, sConsCred, R!cCtaCodRef, 0, nMontoRefTmp, 0, 0, 0, False, gColocEstAprob)
                Call oDCredito.dInsertMovColDet(nMovNro, sConsCred, R!cCtaCodRef, 0, gColocConceptoCodCapital, 0, nMontoRefTmpCapital, False)
                If nMontoRefTmpInteres > 0 Then
                    Call oDCredito.dInsertMovColDet(nMovNro, sConsCred, R!cCtaCodRef, 0, gColocConceptoCodInteresSuspenso, 0, nMontoRefTmpInteres, False)
                End If
                
                'Inserta Movimiento para Gastos en suspenso
                If nMontoRefTmpGastosSusp > 0 Then
                    Call oDCredito.dInsertMovCol(nMovNro, gCredRefGastosSuspenso, R!cCtaCodRef, 0, nMontoRefTmpGastosSusp, 0, 0, 0, False, gColocEstAprob)
                    Call oDCredito.dInsertMovColDet(nMovNro, gCredRefGastosSuspenso, R!cCtaCodRef, 0, gColocConceptoCodGastoSuspenso, 0, nMontoRefTmpGastosSusp, False)
                End If
                'Actualiza Calendario
                If nMontoRefTmpGastosSusp > 0 Then
                    Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, UBound(MatCalendPagos), gColocConceptoCodGastoVarios, nMontoRefTmpGastosSusp, 0, "", False)
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
                    Call oDCredito.dInsertMovCol(nMovNro, gCredRefIntSuspenso, R!cCtaCodRef, 0, nMontoRefTmpInterSusp, 0, 0, 0, False, gColocEstAprob)
                    Call oDCredito.dInsertMovColDet(nMovNro, gCredRefIntSuspenso, R!cCtaCodRef, 0, gColocConceptoCodInteresSuspenso, 0, nMontoRefTmpInterSusp, False)
                End If
                'Insertar Movimientos Gastos en suspenso
                If nMontoRefTmpGastosSusp > 0 Then
                    Call oDCredito.dInsertMovCol(nMovNro, gCredRefGastosSuspenso, R!cCtaCodRef, 0, nMontoRefTmpGastosSusp, 0, 0, 0, False, gColocEstAprob)
                    Call oDCredito.dInsertMovColDet(nMovNro, gCredRefGastosSuspenso, R!cCtaCodRef, 0, gColocConceptoCodGastoSuspenso, 0, nMontoRefTmpGastosSusp, False)
                End If
                If nMontoRefTmpInterSusp > 0 Then
                    'Actualiza Calendario
                    Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, UBound(MatCalendPagos), gColocConceptoCodInteresSuspenso, nMontoRefTmpInterSusp, 0, "", False)
                End If
                If nMontoRefTmpGastosSusp > 0 Then
                    Call oDCredito.dInsertColocCalendDet(psCtaCod, nNroCalend, gColocCalendAplCuota, UBound(MatCalendPagos), gColocConceptoCodGastoVarios, nMontoRefTmpGastosSusp, 0, "", False)
                End If
            End If
            
            'Actualiza Estado
            Call oDCredito.dUpdateProducto_Estado(R!cCtaCodRef, gColocEstRefinanc, pdEstado, , False, True)
                
            'Movimiento para Credito Nuevo
            If pbCapitalInter Then
                nMontoRefTmp = MatrizInteresGastosAFecha(R!cCtaCodRef, MatCredRef, pdEstado) + MatrizCapitalAFecha(R!cCtaCodRef, MatCredRef) - MatrizGastosFecha(R!cCtaCodRef, MatCredRef)
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
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Set oCredito = Nothing
        
        'Ingresar Movimientos para Nuevo Credito
        If pbCapitalInter Then
            'Inserta Movimientos
            Call oDCredito.dInsertMovCol(nMovNro, sConsCred2, psCtaCod, 0, nMontoRefTmpNew, 0, 0, 0, False, gColocEstAprob)
            Call oDCredito.dInsertMovColDet(nMovNro, sConsCred2, psCtaCod, 0, gColocConceptoCodCapital, 0, nMontoRefTmpCapitalNew, False)
            If nMontoRefTmpInteresNew > 0 Then
                Call oDCredito.dInsertMovColDet(nMovNro, sConsCred2, psCtaCod, 0, gColocConceptoCodInteresSuspenso, 0, nMontoRefTmpInteresNew, False)
            End If
        Else
            'Insertar Movimientos
            Call oDCredito.dInsertMovCol(nMovNro, sConsCred2, psCtaCod, 0, nMontoRefTmpNew, 0, 0, 0, False, gColocEstAprob)
            Call oDCredito.dInsertMovColDet(nMovNro, sConsCred2, psCtaCod, 0, gColocConceptoCodCapital, 0, nMontoRefTmpCapitalNew, False)
        End If
        
        'Actualiza Estado de Producto
        Call oDCredito.dUpdateProducto(psCtaCod, , pnMonto, gColocEstRefNorm, pdEstado, 0, False)
        'Actualiza Colocaciones
        Call oDCredito.dUpdateColocaciones(psCtaCod, , , pnMonto, , , pnEstado, False)
        
    Else
        Call oDCredito.dInsertMov(sMovNro, gCredAprobacion, "Aprobacion de Credito", gMovEstContabNoContable, gMovFlagVigente, False)
        nMovNro = oDCredito.dGetnMovNro(sMovNro)
        Call oDCredito.dInsertMovCol(nMovNro, gCredAprobacion, psCtaCod, 0, pnMonto, 0, "", 0, False, gColocEstAprob)
    End If
    
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
Dim I As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For I = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(I, 9)) 'Gastos
    Next I
    
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
Dim I As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For I = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(I, 7)) 'Interes Reprogramado
    Next I
    
    'Total Calculado es
    MatrizInteresReprogramadoFecha = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function

Public Function MatrizInteresSuspensoFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim I As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For I = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(I, 8)) 'Interes Suspenso
    Next I
    
    'Total Calculado es
    MatrizInteresSuspensoFecha = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function


Public Function MatrizInteresGraciaFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date, _
    Optional pnMontoCol As Double = -1) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim I As Integer
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
    Set R = oCredito.RecuperaProductoTasaInteres(psCtaCod, gColocTasaCompGraCia)
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
    If CDate(Format(pdHoy, "dd/mm/yyyy")) - CDate(Format(dFecIni, "dd/mm/yyyy")) >= nDiasGracia Then
        'Calculo de Montos Ya Ganados
        For I = 0 To UBound(MatCalend) - 1
            nMontoGanado = nMontoGanado + CDbl(MatCalend(I, 6)) 'Interes Moratorio
        Next I
    Else
        nMontoGanado = TasaIntPerDias(pnTasa, CDate(Format(pdHoy, "dd/mm/yyyy")) - CDate(Format(dFecIni, "dd/mm/yyyy"))) * nSaldoCal
    End If
    'Total Calculado es
    MatrizInteresGraciaFecha = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function



Public Function MatrizCapitalAFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant) As Double
Dim nMontoGanado As Double
Dim I As Integer
    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    For I = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(I, 3)) 'Capital
    Next I
    
    'Total Calculado es
    MatrizCapitalAFecha = CDbl(Format(nMontoGanado, "#0.00"))

End Function
Private Sub MatrizActualizarEstadoCalendCancelado(ByRef MatCalendDistrib As Variant)
Dim I As Integer

    For I = 0 To UBound(MatCalendDistrib) - 1
        MatCalendDistrib(I, 2) = Trim(Str(gColocCalendEstadoPagado))
    Next I
End Sub
Public Function MatrizInteresMorFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim I As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For I = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(I, 6)) 'Interes Moratorio
    Next I
    
    'Total Calculado es
    MatrizInteresMorFecha = CDbl(Format(nMontoGanado, "#0.00"))
    
End Function

Public Function MatrizCapitalCalendario(ByVal MatCalend As Variant) As Double
Dim I As Integer
    MatrizCapitalCalendario = 0
    For I = 0 To UBound(MatCalend) - 1
        MatrizCapitalCalendario = MatrizCapitalCalendario + CDbl(MatCalend(I, 3))
    Next I
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
Dim I As Integer
    MatrizIntCompCalendario = 0
    For I = 0 To UBound(MatCalend) - 1
        MatrizIntCompCalendario = MatrizIntCompCalendario + CDbl(MatCalend(I, 4))
    Next I
End Function

Public Function MatrizIntGraciaCalendario(ByVal MatCalend As Variant) As Double
Dim I As Integer
    MatrizIntGraciaCalendario = 0
    For I = 0 To UBound(MatCalend) - 1
        MatrizIntGraciaCalendario = MatrizIntGraciaCalendario + CDbl(MatCalend(I, 5))
    Next I
End Function

Public Function MatrizIntMoratorioCalendario(ByVal MatCalend As Variant) As Double
Dim I As Integer
    MatrizIntMoratorioCalendario = 0
    For I = 0 To UBound(MatCalend) - 1
        MatrizIntMoratorioCalendario = MatrizIntMoratorioCalendario + CDbl(MatCalend(I, 6))
    Next I
End Function

Public Function MatrizIntReprogCalendario(ByVal MatCalend As Variant) As Double
Dim I As Integer
    MatrizIntReprogCalendario = 0
    For I = 0 To UBound(MatCalend) - 1
        MatrizIntReprogCalendario = MatrizIntReprogCalendario + CDbl(MatCalend(I, 7))
    Next I
End Function

Public Function MatrizIntSuspensoCalendario(ByVal MatCalend As Variant) As Double
Dim I As Integer
    MatrizIntSuspensoCalendario = 0
    For I = 0 To UBound(MatCalend) - 1
        MatrizIntSuspensoCalendario = MatrizIntSuspensoCalendario + CDbl(MatCalend(I, 8))
    Next I
End Function

Public Function MatrizIntGastosCalendario(ByVal MatCalend As Variant) As Double
Dim I As Integer
    MatrizIntGastosCalendario = 0
    For I = 0 To UBound(MatCalend) - 1
        MatrizIntGastosCalendario = MatrizIntGastosCalendario + CDbl(MatCalend(I, 9))
    Next I
End Function

Public Function MatrizInteresCompAFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date, _
    Optional pnMontoCol As Double = -1, Optional pnTasaInteres As Double = -1) As Double
Dim R As ADODB.Recordset
Dim oCredito As DCredito
Dim oCalend As Dcalendario
Dim nMontoGanado As Double
Dim nMontoAFecha As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim I As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For I = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(I, 0)) Then
            nMontoGanado = nMontoGanado + CDbl(MatCalend(I, 4))
        End If
    Next I
    
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
    For I = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(I, 0)) Then
            dFecIni = CDate(MatCalend(I, 0))
        Else
            If (pdHoy - dFecIni) > 0 Then
                If CInt(MatCalend(I, 1)) = 1 Then
                    nMontoAFecha = TasaIntPerDias(pnTasa, pdHoy - dFecIni) * nSaldoCal
                Else
                    nMontoAFecha = TasaIntPerDias(pnTasa, pdHoy - dFecIni) * CDbl(MatCalend(I, 10))
                End If
            Else
                nMontoAFecha = 0
            End If
        End If
    Next I
    
    'Total Calculado es
    MatrizInteresCompAFecha = CDbl(Format(nMontoGanado + nMontoAFecha, "#0.00"))
    
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
                MatrizInteresCompAFecha(psCtaCod, MatCalend, pdHoy, nSaldoCal, pnTasa), "#0.00"))
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
                MatrizInteresCompAFecha(psCtaCod, MatCalend, pdHoy, nSaldoCal, pnTasa), "#0.00"))
End Function


Public Function MatrizInteresGastosAFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
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
                MatrizInteresCompAFecha(psCtaCod, MatCalend, pdHoy, nSaldoCal, pnTasa), "#0.00"))
    
End Function


Public Function ReprogramarCreditoenMemoriaTotal(ByVal psCtaCod As String, ByVal pdHoy As Date) As Variant
Dim oCred As DCredito
Dim oCalend As NCalendario
Dim oDCalend As Dcalendario
Dim RAprob As ADODB.Recordset
Dim MatCalend As Variant
Dim R As ADODB.Recordset
Dim nInteresReprog As Double
Dim nMontoPrestamo As Double
Dim nCuotasPendientes As Integer
Dim nTipocalend As Integer

    ReDim MatCalend(0, 0)
    Set oCred = New DCredito
    Set RAprob = oCred.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
    
    Set R = oCred.RecuperaProducto(psCtaCod)
    nMontoPrestamo = CDbl(Format(R!nSaldo, "#0.00"))
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
    MatCalend = oCalend.GeneraCalendario(nMontoPrestamo, R!nTasaInteres, nCuotasPendientes, IIf(IsNull(RAprob!nPlazo), 0, RAprob!nPlazo), _
        pdHoy, nTipocalend, IIf(RAprob!nPeriodoFechaFija > 0, 2, 1), Exonerada, 0, RAprob!nPeriodoFechaFija, True)
    Set oCalend = Nothing
    R.Close
    Set R = Nothing
    RAprob.Close
    Set RAprob = Nothing
    
    ReprogramarCreditoenMemoriaTotal = MatCalend
    
End Function

Public Function ReprogramarCreditoenMemoria(ByVal psCtaCod As String, ByVal pnTasa As Double, _
    ByVal pdFecOrig As Date, ByVal pdFecReprog As Date, ByVal nFilaReprog As Integer, ByVal pnTipoReprog As Integer) As Variant

Dim MontoAReprogramar As Double
Dim MontoAProratear As Double
Dim nDias As Integer
Dim I As Integer
Dim j As Integer
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

    'Carga Calendario A Reprogramar
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
    Set oCredito = Nothing
    nMontoCredito = IIf(IsNull(R!nMonto), 0, R!nMonto)
    nMontoSaldo = IIf(IsNull(R!nMonto), 0, R!nMonto)
    R.Close
    Set R = Nothing
    
    Set oCal = New Dcalendario
    Set R = oCal.RecuperaCalendarioPagos(psCtaCod)
    ReDim MatCalend(R.RecordCount, 18)
    Do While Not R.EOF
        MatCalend(R.Bookmark - 1, 0) = Format(R!dVenc, "dd/mm/yyyy") 'Fecha de Vencimiento
        MatCalend(R.Bookmark - 1, 1) = Trim(Str(R!nCuota)) 'NroCuota
        MatCalend(R.Bookmark - 1, 2) = Format(R!nCapital, "#0.00") 'Capital
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
        MatCalend(R.Bookmark - 1, 14) = Format(IIf(IsNull(R!nGasto), 0, R!nGasto), "#0.00") 'Gastos
        MatCalend(R.Bookmark - 1, 15) = Format(IIf(IsNull(R!nGastoPag), 0, R!nGastoPag), "#0.00") 'Gastos Pagados
        nMontoCredito = nMontoCredito - R!nCapital
        nMontoCredito = CDbl(Format(nMontoCredito, "#0.00"))
        MatCalend(R.Bookmark - 1, 16) = Format(nMontoCredito, "#0.00") 'Saldo Capital
        MatCalend(R.Bookmark - 1, 17) = Trim(Str(R!nColocCalendEstado))
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    MatCalend(nFilaReprog, 0) = Format(pdFecReprog, "dd/mm/yyyy")
    'Mover la Fechas
    nCapitalFusionado = 0
    For I = 0 To 6
        nMontoMovido(I, 0) = 0#
        nMontoMovido(I, 1) = 0#
    Next I
    For I = nFilaReprog + 1 To UBound(MatCalend) - 1
        If CDate(MatCalend(nFilaReprog, 0)) >= CDate(MatCalend(I, 0)) Then
            nCapitalFusionado = nCapitalFusionado + CDbl(MatCalend(I, 2))
            nCapitalFusionado = Format(nCapitalFusionado, "#0.00")
            
            'Mueve Capital a la Fila Reprogramada
            nMontoMovido(0, 0) = CDbl(Format((nMontoMovido(0, 0) + CDbl(MatCalend(I, 2))), "#0.00")) 'Capital Pendiente
            nMontoMovido(0, 1) = CDbl(Format((nMontoMovido(0, 1) + CDbl(MatCalend(I, 3))), "#0.00"))  'Capital Pagado
            'MatCalend(nFilaReprog, 2) = Format(CDbl(MatCalend(nFilaReprog, 2)) + CDbl(MatCalend(I, 2)), "#0.00")
            'MatCalend(nFilaReprog, 3) = Format(CDbl(MatCalend(nFilaReprog, 3)) + CDbl(MatCalend(I, 3)), "#0.00")
            
            'Mueve Intereses Compensatorio
            'nMontoMovido(1, 0) = nMontoMovido(1, 0) + CDbl(Format(CDbl(MatCalend(nFilaReprog, 4)) + CDbl(MatCalend(I, 4)), "#0.00")) 'Interes Compensatorio pendiente
            'nMontoMovido(1, 1) = nMontoMovido(1, 1) + CDbl(Format(CDbl(MatCalend(nFilaReprog, 5)) + CDbl(MatCalend(I, 5)), "#0.00")) 'Interes Compensatorio pendiente
            'MatCalend(nFilaReprog, 4) = Format(CDbl(MatCalend(nFilaReprog, 4)) + CDbl(MatCalend(I, 4)), "#0.00")
            'MatCalend(nFilaReprog, 5) = Format(CDbl(MatCalend(nFilaReprog, 5)) + CDbl(MatCalend(I, 5)), "#0.00")
            
            'Mueve Intereses Moratorio
            nMontoMovido(2, 0) = CDbl(Format((nMontoMovido(2, 0) + CDbl(MatCalend(I, 6))), "#0.00")) 'Interes Moratorio pendiente
            nMontoMovido(2, 1) = CDbl(Format((nMontoMovido(2, 1) + CDbl(MatCalend(I, 7))), "#0.00"))  'Interes Moratorio pendiente
            'MatCalend(nFilaReprog, 6) = Format(CDbl(MatCalend(nFilaReprog, 6)) + CDbl(MatCalend(I, 6)), "#0.00")
            'MatCalend(nFilaReprog, 7) = Format(CDbl(MatCalend(nFilaReprog, 7)) + CDbl(MatCalend(I, 7)), "#0.00")
            
            'Mueve Intereses Reprogramado
            nMontoMovido(3, 0) = CDbl(Format((nMontoMovido(3, 0) + CDbl(MatCalend(I, 8))), "#0.00")) 'Interes Reporgramado pendiente
            nMontoMovido(3, 1) = CDbl(Format((nMontoMovido(3, 1) + CDbl(MatCalend(I, 9))), "#0.00")) 'Interes Reprogramado pendiente
            'MatCalend(nFilaReprog, 8) = Format(CDbl(MatCalend(nFilaReprog, 8)) + CDbl(MatCalend(I, 8)), "#0.00")
            'MatCalend(nFilaReprog, 9) = Format(CDbl(MatCalend(nFilaReprog, 9)) + CDbl(MatCalend(I, 9)), "#0.00")
            
            'Mueve Intereses Gracia
            nMontoMovido(4, 0) = CDbl(Format((nMontoMovido(4, 0) + CDbl(MatCalend(I, 10))), "#0.00")) 'Interes Gracia pendiente
            nMontoMovido(4, 1) = CDbl(Format((nMontoMovido(4, 1) + CDbl(MatCalend(I, 11))), "#0.00"))  'Interes Gracia pendiente
            'MatCalend(nFilaReprog, 10) = Format(CDbl(MatCalend(nFilaReprog, 10)) + CDbl(MatCalend(I, 10)), "#0.00")
            'MatCalend(nFilaReprog, 11) = Format(CDbl(MatCalend(nFilaReprog, 11)) + CDbl(MatCalend(I, 11)), "#0.00")
            
            'Mueve Intereses Suspenso
            nMontoMovido(5, 0) = CDbl(Format((nMontoMovido(5, 0) + CDbl(MatCalend(I, 12))), "#0.00"))  'Interes Suspenso pendiente
            nMontoMovido(5, 1) = CDbl(Format((nMontoMovido(5, 1) + CDbl(MatCalend(I, 13))), "#0.00")) 'Interes Suspenso pendiente
            'MatCalend(nFilaReprog, 12) = Format(CDbl(MatCalend(nFilaReprog, 12)) + CDbl(MatCalend(I, 12)), "#0.00")
            'MatCalend(nFilaReprog, 13) = Format(CDbl(MatCalend(nFilaReprog, 13)) + CDbl(MatCalend(I, 13)), "#0.00")
            
            'Mueve Gastos Como Intereses Reprogramados
            nMontoMovido(6, 0) = CDbl(Format((nMontoMovido(6, 0) + CDbl(MatCalend(I, 14))), "#0.00"))  'Gastos pendiente
            nMontoMovido(6, 1) = CDbl(Format((nMontoMovido(6, 1) + CDbl(MatCalend(I, 15))), "#0.00"))  'Gastos pendiente
            'MatCalend(nFilaReprog, 14) = Format(CDbl(MatCalend(nFilaReprog, 14)) + CDbl(MatCalend(I, 14)), "#0.00")
            'MatCalend(nFilaReprog, 15) = Format(CDbl(MatCalend(nFilaReprog, 15)) + CDbl(MatCalend(I, 15)), "#0.00")
            MatCalend(I, 0) = ""
        Else
            'nMontoMovido(3, 0) = nMontoMovido(3, 0) + (TasaIntPerDias(pnTasa, (pdFecReprog - CDate(MatCalend(I - 1, 0)))) * CDbl(MatCalend(I - 1, 16))) 'Interes Reporgramado pendiente
            'nMontoMovido(3, 1) = nMontoMovido(3, 1) + CDbl(Format(CDbl(MatCalend(nFilaReprog, 9)) + CDbl(MatCalend(I, 9)), "#0.00")) 'Interes Reprogramado pendiente
            MatCalend(I - 1, 16) = Format(CDbl(MatCalend(I - 1, 16)) + nCapitalFusionado, "#0.00")
            Exit For
        End If
    Next I
    'Hallar Filas Reales
    nNumFilasReal = 0
    For I = 0 To UBound(MatCalend) - 1
        If MatCalend(I, 0) <> "" Then
            nNumFilasReal = nNumFilasReal + 1
        End If
    Next I
    
    ReDim MatCalTmp(nNumFilasReal, 18)
    nPunt = 0
    j = 0
    For I = 0 To UBound(MatCalend) - 1
        If MatCalend(I, 0) <> "" Then
            MatCalTmp(nPunt, 0) = MatCalend(I, 0)
            MatCalTmp(nPunt, 1) = nPunt + 1
            MatCalTmp(nPunt, 2) = MatCalend(I, 2)
            MatCalTmp(nPunt, 3) = MatCalend(I, 3)
            MatCalTmp(nPunt, 4) = MatCalend(I, 4)
            MatCalTmp(nPunt, 5) = MatCalend(I, 5)
            MatCalTmp(nPunt, 6) = MatCalend(I, 6)
            MatCalTmp(nPunt, 7) = MatCalend(I, 7)
            MatCalTmp(nPunt, 8) = MatCalend(I, 8)
            MatCalTmp(nPunt, 9) = MatCalend(I, 9)
            MatCalTmp(nPunt, 10) = MatCalend(I, 10)
            MatCalTmp(nPunt, 11) = MatCalend(I, 11)
            MatCalTmp(nPunt, 12) = MatCalend(I, 12)
            MatCalTmp(nPunt, 13) = MatCalend(I, 13)
            MatCalTmp(nPunt, 14) = MatCalend(I, 14)
            MatCalTmp(nPunt, 15) = MatCalend(I, 15)
            MatCalTmp(nPunt, 17) = MatCalend(I, 17)
            nPunt = nPunt + 1
            'Solo Pendientes
            If CInt(Trim(MatCalTmp(nPunt - 1, 17))) = gColocCalendEstadoPendiente Then
                j = j + 1 'Total Pendientes
            End If
        End If
    Next I
    
    'Solo para Capital
    If pnTipoReprog = 1 Then 'Proratear
        nMontoMovidoProrateado(0, 0) = Format((nMontoMovido(0, 0) / j), "#0.00")
        nMontoMovidoProrateado(0, 1) = Format((nMontoMovido(0, 1) / j), "#0.00")
        For I = 0 To UBound(MatCalTmp) - 1
            If CInt(Trim(MatCalTmp(I, 17))) = gColocCalendEstadoPendiente Then
                If I <> UBound(MatCalTmp) - 1 Then
                    MatCalTmp(I, 2) = Format(CDbl(MatCalTmp(I, 2)) + nMontoMovidoProrateado(0, 0), "#0.00")
                    nMontoMovido(0, 0) = nMontoMovido(0, 0) - nMontoMovidoProrateado(0, 0)
                    nMontoMovido(0, 0) = CDbl(Format(nMontoMovido(0, 0), "#0.00"))
                    
                    MatCalTmp(I, 3) = Format(CDbl(MatCalTmp(I, 3)) + nMontoMovidoProrateado(0, 1), "#0.00")
                    nMontoMovido(0, 1) = nMontoMovido(0, 1) - nMontoMovidoProrateado(0, 1)
                    nMontoMovido(0, 1) = CDbl(Format(nMontoMovido(0, 1), "#0.00"))
                Else
                    MatCalTmp(I, 2) = Format(CDbl(MatCalTmp(I, 2)) + nMontoMovido(0, 0), "#0.00")
                    MatCalTmp(I, 3) = Format(CDbl(MatCalTmp(I, 3)) + nMontoMovido(0, 1), "#0.00")
                End If
            End If
        Next I
    Else
        MatCalTmp(UBound(MatCalTmp) - 1, 2) = Format(CDbl(MatCalTmp(UBound(MatCalTmp) - 1, 2)) + nMontoMovido(0, 0), "#0.00")
        MatCalTmp(UBound(MatCalTmp) - 1, 3) = Format(CDbl(MatCalTmp(UBound(MatCalTmp) - 1, 3)) + nMontoMovido(0, 1), "#0.00")
    End If
    
    For I = 0 To UBound(MatCalTmp) - 1
        nMontoSaldo = nMontoSaldo - CDbl(MatCalTmp(I, 2))
        nMontoSaldo = CDbl(Format(nMontoSaldo, "#0.00"))
        MatCalTmp(I, 16) = Format(nMontoSaldo, "#0.00")
    Next I
    'Calculo del Nuevo Interes para la Fila despues del Reprogramado
    If (nPunt - 1) <> nFilaReprog Then
        nDias = CDate(MatCalTmp(nFilaReprog + 1, 0)) - CDate(MatCalTmp(nFilaReprog, 0))
    Else
        nDias = CDate(MatCalTmp(nFilaReprog, 0)) - CDate(MatCalTmp(nFilaReprog - 1, 0))
    End If
    MontoAReprogramar = TasaIntPerDias(pnTasa, nDias) * CDbl(MatCalTmp(nFilaReprog, 16))
    MontoAReprogramar = CDbl(Format(MontoAReprogramar, "#0.00"))
    MatCalTmp(nFilaReprog + 1, 4) = Format(MontoAReprogramar, "#0.00")
    
    'Calulo por Dias Reprogramados
    nDias = CDate(MatCalTmp(nFilaReprog, 0)) - pdFecOrig
    MontoAReprogramar = TasaIntPerDias(pnTasa, nDias) * CDbl(MatCalTmp(nFilaReprog - 1, 16))
    MontoAReprogramar = CDbl(Format(MontoAReprogramar, "#0.00"))

    If pnTipoReprog = 1 Then 'Proratear
        MontoAProratear = Format((MontoAReprogramar / UBound(MatCalTmp)), "#0.00")
        nMontoMovidoProrateado(2, 0) = Format((nMontoMovido(2, 0) / UBound(MatCalTmp)), "#0.00")
        nMontoMovidoProrateado(2, 1) = Format((nMontoMovido(2, 1) / UBound(MatCalTmp)), "#0.00")
        nMontoMovidoProrateado(3, 0) = Format((nMontoMovido(3, 0) / UBound(MatCalTmp)), "#0.00")
        nMontoMovidoProrateado(3, 1) = Format((nMontoMovido(3, 1) / UBound(MatCalTmp)), "#0.00")
        nMontoMovidoProrateado(4, 0) = Format((nMontoMovido(4, 0) / UBound(MatCalTmp)), "#0.00")
        nMontoMovidoProrateado(4, 1) = Format((nMontoMovido(4, 1) / UBound(MatCalTmp)), "#0.00")
        nMontoMovidoProrateado(5, 0) = Format((nMontoMovido(5, 0) / UBound(MatCalTmp)), "#0.00")
        nMontoMovidoProrateado(5, 1) = Format((nMontoMovido(5, 1) / UBound(MatCalTmp)), "#0.00")
        nMontoMovidoProrateado(6, 0) = Format((nMontoMovido(6, 0) / UBound(MatCalTmp)), "#0.00")
        nMontoMovidoProrateado(6, 1) = Format((nMontoMovido(6, 1) / UBound(MatCalTmp)), "#0.00")
        For I = 0 To UBound(MatCalTmp) - 1
            If CInt(Trim(MatCalTmp(I, 17))) = gColocCalendEstadoPendiente Then
                If I <> UBound(MatCalTmp) - 1 Then
                    MatCalTmp(I, 8) = Format(CDbl(MatCalTmp(I, 8)) + MontoAProratear, "#0.00")
                    MontoAReprogramar = MontoAReprogramar - MontoAProratear
                    MontoAReprogramar = CDbl(Format(MontoAReprogramar, "#0.00"))
                    
                    MatCalTmp(I, 6) = Format(CDbl(MatCalTmp(I, 6)) + nMontoMovidoProrateado(2, 0), "#0.00")
                    nMontoMovido(2, 0) = nMontoMovido(2, 0) - nMontoMovidoProrateado(2, 0)
                    nMontoMovido(2, 0) = CDbl(Format(nMontoMovido(2, 0), "#0.00"))
                    
                    MatCalTmp(I, 7) = Format(CDbl(MatCalTmp(I, 6)) + nMontoMovidoProrateado(2, 1), "#0.00")
                    nMontoMovido(2, 1) = nMontoMovido(2, 1) - nMontoMovidoProrateado(2, 1)
                    nMontoMovido(2, 1) = CDbl(Format(nMontoMovido(2, 1), "#0.00"))
                    
                    MatCalTmp(I, 8) = Format(CDbl(MatCalTmp(I, 8)) + nMontoMovidoProrateado(3, 0), "#0.00")
                    nMontoMovido(3, 0) = nMontoMovido(3, 0) - nMontoMovidoProrateado(3, 0)
                    nMontoMovido(3, 0) = CDbl(Format(nMontoMovido(3, 0), "#0.00"))
                    
                    MatCalTmp(I, 9) = Format(CDbl(MatCalTmp(I, 9)) + nMontoMovidoProrateado(3, 1), "#0.00")
                    nMontoMovido(3, 1) = nMontoMovido(3, 1) - nMontoMovidoProrateado(3, 1)
                    nMontoMovido(3, 1) = CDbl(Format(nMontoMovido(3, 1), "#0.00"))
                    
                    MatCalTmp(I, 10) = Format(CDbl(MatCalTmp(I, 10)) + nMontoMovidoProrateado(4, 0), "#0.00")
                    nMontoMovido(4, 0) = nMontoMovido(4, 0) - nMontoMovidoProrateado(4, 0)
                    nMontoMovido(4, 0) = CDbl(Format(nMontoMovido(4, 0), "#0.00"))
                    
                    MatCalTmp(I, 11) = Format(CDbl(MatCalTmp(I, 11)) + nMontoMovidoProrateado(4, 1), "#0.00")
                    nMontoMovido(4, 1) = nMontoMovido(4, 1) - nMontoMovidoProrateado(4, 1)
                    nMontoMovido(4, 1) = CDbl(Format(nMontoMovido(4, 1), "#0.00"))
                    
                    MatCalTmp(I, 12) = Format(CDbl(MatCalTmp(I, 12)) + nMontoMovidoProrateado(5, 0), "#0.00")
                    nMontoMovido(5, 0) = nMontoMovido(5, 0) - nMontoMovidoProrateado(5, 0)
                    nMontoMovido(5, 0) = CDbl(Format(nMontoMovido(5, 0), "#0.00"))
                    
                    MatCalTmp(I, 13) = Format(CDbl(MatCalTmp(I, 13)) + nMontoMovidoProrateado(5, 1), "#0.00")
                    nMontoMovido(5, 1) = nMontoMovido(5, 1) - nMontoMovidoProrateado(5, 1)
                    nMontoMovido(5, 1) = CDbl(Format(nMontoMovido(5, 1), "#0.00"))
                    
                    MatCalTmp(I, 14) = Format(CDbl(MatCalTmp(I, 14)) + nMontoMovidoProrateado(6, 0), "#0.00")
                    nMontoMovido(6, 0) = nMontoMovido(6, 0) - nMontoMovidoProrateado(6, 0)
                    nMontoMovido(6, 0) = CDbl(Format(nMontoMovido(6, 0), "#0.00"))
                    
                    MatCalTmp(I, 15) = Format(CDbl(MatCalTmp(I, 15)) + nMontoMovidoProrateado(6, 1), "#0.00")
                    nMontoMovido(6, 1) = nMontoMovido(6, 1) - nMontoMovidoProrateado(6, 1)
                    nMontoMovido(6, 1) = CDbl(Format(nMontoMovido(6, 1), "#0.00"))
                Else
                    MatCalTmp(I, 8) = Format(CDbl(MatCalTmp(I, 8)) + MontoAReprogramar, "#0.00")
                    MatCalTmp(I, 6) = Format(CDbl(MatCalTmp(I, 6)) + nMontoMovido(2, 0), "#0.00")
                    MatCalTmp(I, 7) = Format(CDbl(MatCalTmp(I, 7)) + nMontoMovido(2, 1), "#0.00")
                    MatCalTmp(I, 8) = Format(CDbl(MatCalTmp(I, 8)) + nMontoMovido(3, 0), "#0.00")
                    MatCalTmp(I, 9) = Format(CDbl(MatCalTmp(I, 9)) + nMontoMovido(3, 1), "#0.00")
                    MatCalTmp(I, 10) = Format(CDbl(MatCalTmp(I, 10)) + nMontoMovido(4, 0), "#0.00")
                    MatCalTmp(I, 11) = Format(CDbl(MatCalTmp(I, 11)) + nMontoMovido(4, 1), "#0.00")
                    MatCalTmp(I, 12) = Format(CDbl(MatCalTmp(I, 12)) + nMontoMovido(5, 0), "#0.00")
                    MatCalTmp(I, 13) = Format(CDbl(MatCalTmp(I, 13)) + nMontoMovido(5, 1), "#0.00")
                    MatCalTmp(I, 14) = Format(CDbl(MatCalTmp(I, 14)) + nMontoMovido(6, 0), "#0.00")
                    MatCalTmp(I, 15) = Format(CDbl(MatCalTmp(I, 15)) + nMontoMovido(6, 1), "#0.00")
                End If
            End If
        Next I
    Else    'A la Ultima Cuota
        MatCalTmp(UBound(MatCalTmp) - 1, 8) = Format(CDbl(MatCalTmp(UBound(MatCalTmp) - 1, 8)) + MontoAReprogramar, "#0.00")
        MatCalTmp(UBound(MatCalTmp) - 1, 6) = Format(CDbl(MatCalTmp(UBound(MatCalTmp) - 1, 6)) + nMontoMovido(2, 0), "#0.00")
        MatCalTmp(UBound(MatCalTmp) - 1, 7) = Format(CDbl(MatCalTmp(UBound(MatCalTmp) - 1, 7)) + nMontoMovido(2, 1), "#0.00")
        MatCalTmp(UBound(MatCalTmp) - 1, 8) = Format(CDbl(MatCalTmp(UBound(MatCalTmp) - 1, 8)) + nMontoMovido(3, 0), "#0.00")
        MatCalTmp(UBound(MatCalTmp) - 1, 9) = Format(CDbl(MatCalTmp(UBound(MatCalTmp) - 1, 9)) + nMontoMovido(3, 1), "#0.00")
        MatCalTmp(UBound(MatCalTmp) - 1, 10) = Format(CDbl(MatCalTmp(UBound(MatCalTmp) - 1, 10)) + nMontoMovido(4, 0), "#0.00")
        MatCalTmp(UBound(MatCalTmp) - 1, 11) = Format(CDbl(MatCalTmp(UBound(MatCalTmp) - 1, 11)) + nMontoMovido(4, 1), "#0.00")
        MatCalTmp(UBound(MatCalTmp) - 1, 12) = Format(CDbl(MatCalTmp(UBound(MatCalTmp) - 1, 12)) + nMontoMovido(5, 0), "#0.00")
        MatCalTmp(UBound(MatCalTmp) - 1, 13) = Format(CDbl(MatCalTmp(UBound(MatCalTmp) - 1, 13)) + nMontoMovido(5, 1), "#0.00")
        MatCalTmp(UBound(MatCalTmp) - 1, 14) = Format(CDbl(MatCalTmp(UBound(MatCalTmp) - 1, 14)) + nMontoMovido(6, 0), "#0.00")
        MatCalTmp(UBound(MatCalTmp) - 1, 15) = Format(CDbl(MatCalTmp(UBound(MatCalTmp) - 1, 15)) + nMontoMovido(6, 1), "#0.00")
    End If
    Set oCal = Nothing
    ReprogramarCreditoenMemoria = MatCalTmp
End Function

Public Function ReprogramarCredito(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pnTipoReprog As Integer) As String
'pnTipoReprog = 1 Reprogramacion con Edicion de Fecha
'pnTipoReprog = 2 Reprogramacion Total Otorgamiento de un Nuevo
'                 Credito bajo el Saldo de capital Existente y las mismas condiciones que el anterior
Dim oCredito As DCredito
Dim oBase As DCredActualizaBD
Dim oCalend As Dcalendario
Dim nNroCalen As Integer
Dim R As ADODB.Recordset
Dim I As Integer

    On Error GoTo ErrorReprogramarCredito
    ReprogramarCredito = ""
    Set oCredito = New DCredito
    Set R = oCredito.RecuperaColocacCred(psCtaCod)
    nNroCalen = R!nNroCalen + 1
    R.Close
    Set R = Nothing
    
    Set oBase = New DCredActualizaBD
    If pnTipoReprog = 1 Then
        For I = 0 To UBound(MatCalend) - 1
            Call oBase.dInsertColocCalendario(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(I, 1)), CDate(MatCalend(I, 0)), CInt(MatCalend(I, 17)), "Calendario Reprogramado", gColocCalendConceptoProcReprogramado, True)
            If CDbl(MatCalend(I, 2)) > 0 Then
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(I, 1)), gColocConceptoCodCapital, CDbl(MatCalend(I, 2)), CDbl(MatCalend(I, 3)), "", True)
            End If
            If CDbl(MatCalend(I, 4)) > 0 Then
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(I, 1)), gColocConceptoCodInteresCompensatorio, CDbl(MatCalend(I, 4)), CDbl(MatCalend(I, 5)), "", True)
            End If
            If CDbl(MatCalend(I, 6)) > 0 Then
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(I, 1)), gColocConceptoCodInteresMoratorio, CDbl(MatCalend(I, 6)), CDbl(MatCalend(I, 7)), "", True)
            End If
            If CDbl(MatCalend(I, 8)) > 0 Then
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(I, 1)), gColocConceptoCodInteresReprogramado, CDbl(MatCalend(I, 8)), CDbl(MatCalend(I, 9)), "", True)
            End If
            If CDbl(MatCalend(I, 10)) > 0 Then
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(I, 1)), gColocConceptoCodInteresGracia, CDbl(MatCalend(I, 10)), CDbl(MatCalend(I, 11)), "", True)
            End If
            If CDbl(MatCalend(I, 12)) > 0 Then
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(I, 1)), gColocConceptoCodInteresSuspenso, CDbl(MatCalend(I, 12)), CDbl(MatCalend(I, 13)), "", True)
            End If
            
            If CDbl(MatCalend(I, 14)) > 0 Then
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalend(I, 1)), gColocConceptoCodGastoVarios, CDbl(MatCalend(I, 14)), CDbl(MatCalend(I, 15)), "", True)
            End If
        Next I
    Else
        For I = 0 To UBound(MatCalend) - 1
            Call oBase.dInsertColocCalendario(psCtaCod, nNroCalen, gColocCalendAplCuota, I + 1, CDate(MatCalend(I, 0)), gColocCalendEstadoPendiente, "Calendario Reporgramado de Pagos", gColocCalendConceptoProcReprogramado, True)
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, I + 1, gColocConceptoCodCapital, CDbl(MatCalend(I, 3)), 0, "", True)
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresCompensatorio, CDbl(MatCalend(I, 4)), 0, "", True)
                Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresMoratorio, 0, 0, "", True)
                If CDbl(MatCalend(I, 5)) > 0 Then
                    Call oBase.dInsertColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, I + 1, gColocConceptoCodInteresGracia, CDbl(MatCalend(I, 5)), 0, "", True)
                End If
        Next I
    End If
    Call oBase.dUpdateColocacCred(psCtaCod, , , , , , , , , , , , , , nNroCalen, , , , True)
    oBase.dEjecutaBatch
    Set oBase = Nothing
    Exit Function

ErrorReprogramarCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Private Function MatrizDistribuirGastos(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, ByRef pnMonto As Double, _
        Optional ByVal DistVert As Boolean = False)
Dim I As Integer

    For I = 0 To UBound(MatCalend) - 1
        If CInt(MatCalend(I, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 And (CDbl(MatCalend(I, 9)) - CDbl(MatCalendDistrib(I, 9))) > 0 Then
            If pnMonto > (CDbl(MatCalend(I, 9)) - CDbl(MatCalendDistrib(I, 9))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(I, 9)) - CDbl(MatCalendDistrib(I, 9)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(I, 9) = MatCalend(I, 9)
            Else
                MatCalendDistrib(I, 9) = Format(CDbl(MatCalendDistrib(I, 9)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
            If Not DistVert Then
                Exit For
            End If
        End If
    Next I
    
End Function
'*********************************************************
'** Cubrir Verticalmente Todas las Moras de las Cuotas
'*********************************************************
Private Function MatrizDistribuirMora(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, ByRef pnMonto As Double)
Dim I As Integer

    For I = 0 To UBound(MatCalend) - 1
        If CInt(MatCalend(I, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(I, 6)) - CDbl(MatCalendDistrib(I, 6))) > 0 Then
            If pnMonto > (CDbl(MatCalend(I, 6)) - CDbl(MatCalendDistrib(I, 6))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(I, 6)) - CDbl(MatCalendDistrib(I, 6)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(I, 6) = MatCalend(I, 6)
            Else
                MatCalendDistrib(I, 6) = Format(CDbl(MatCalendDistrib(I, 6)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
    Next I
End Function

Private Function MatrizDistribuirInteresI(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, _
    ByRef pnMonto As Double, Optional ByVal DistVert As Boolean = False)
Dim I As Integer
Dim bSiCubrio As Boolean
    bSiCubrio = False
    For I = 0 To UBound(MatCalend) - 1
        'Cubre Interes en Suspenso
        If CInt(MatCalend(I, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(I, 8)) - CDbl(MatCalendDistrib(I, 8))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(I, 8)) - CDbl(MatCalendDistrib(I, 8))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(I, 8)) - CDbl(MatCalendDistrib(I, 8)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(I, 8) = MatCalend(I, 8)
            Else
                MatCalendDistrib(I, 8) = Format(CDbl(MatCalendDistrib(I, 8)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        'Cubre Intereses Reprogramados
        If CInt(MatCalend(I, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(I, 7)) - CDbl(MatCalendDistrib(I, 7))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(I, 7)) - CDbl(MatCalendDistrib(I, 7))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(I, 7)) - CDbl(MatCalendDistrib(I, 7)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(I, 7) = MatCalend(I, 7)
            Else
                MatCalendDistrib(I, 7) = Format(CDbl(MatCalendDistrib(I, 7)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        'Intereses de Gracia
        If CInt(MatCalend(I, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(I, 5)) - CDbl(MatCalendDistrib(I, 5))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(I, 5)) - CDbl(MatCalendDistrib(I, 5))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(I, 5)) - CDbl(MatCalendDistrib(I, 5)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(I, 5) = MatCalend(I, 5)
            Else
                MatCalendDistrib(I, 5) = Format(CDbl(MatCalendDistrib(I, 5)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        'Intereses Compensatorios
        If CInt(MatCalend(I, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(I, 4)) - CDbl(MatCalendDistrib(I, 4))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(I, 4)) - CDbl(MatCalendDistrib(I, 4))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(I, 4)) - CDbl(MatCalendDistrib(I, 4)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(I, 4) = MatCalend(I, 4)
            Else
                MatCalendDistrib(I, 4) = Format(CDbl(MatCalendDistrib(I, 4)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        If bSiCubrio Then
            If Not DistVert Then
                Exit For
            End If
        End If
    Next I
End Function

Private Function MatrizDistribuirCapital(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, _
    ByRef pnMonto As Double, Optional ByVal DistVert As Boolean = False)
Dim I As Integer

    For I = 0 To UBound(MatCalend) - 1
        If CInt(MatCalend(I, 2)) = gColocCalendEstadoPendiente And pnMonto > 0 _
            And (CDbl(MatCalend(I, 3)) - CDbl(MatCalendDistrib(I, 3))) > 0 Then
            If pnMonto > (CDbl(MatCalend(I, 3)) - CDbl(MatCalendDistrib(I, 3))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(I, 3)) - CDbl(MatCalendDistrib(I, 3)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(I, 3) = MatCalend(I, 3)
            Else
                MatCalendDistrib(I, 3) = Format(CDbl(MatCalendDistrib(I, 3)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
            If Not DistVert Then
                Exit For
            End If
        End If
    Next I
End Function

Private Sub MatrizAdicionaInteresSuspenso(ByVal MatCalend As Variant, ByVal nMontoSusp As Double)
    
End Sub

Private Sub MatrizAdicionaGastosSuspenso(ByVal MatCalend As Variant, ByVal nMontoSusp As Double)

End Sub

Private Sub MatrizActualizarEstadoCuota(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant)
Dim I As Integer
    For I = 0 To UBound(MatCalend) - 1
        If CInt(MatCalendDistrib(I, 2)) = gColocCalendEstadoPendiente Then
            If CDbl(MatCalend(I, 8)) = CDbl(MatCalendDistrib(I, 8)) And _
                CDbl(MatCalend(I, 7)) = CDbl(MatCalendDistrib(I, 7)) And _
                CDbl(MatCalend(I, 6)) = CDbl(MatCalendDistrib(I, 6)) And _
                CDbl(MatCalend(I, 5)) = CDbl(MatCalendDistrib(I, 5)) And _
                CDbl(MatCalend(I, 3)) = CDbl(MatCalendDistrib(I, 3)) Then
                MatCalendDistrib(I, 2) = Trim(Str(gColocCalendEstadoPagado))
            End If
        End If
    Next I
End Sub

Public Function MatrizDistribuirCancelacion(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pnMontoPago As Double, _
                 ByVal psMetLiquidacion As String, ByVal pdHoy As Date) As Variant

Dim MatCalendDistrib As Variant
Dim nMontoGastos As Double
Dim nMontoMora As Double
Dim nMontoInteres As Double
Dim nMontoCapital As Double
Dim nMontoTotalTmp As Double
Dim j As Integer

        MatCalendDistrib = CrearMatrizparaAmortizacion(MatCalend)
        nMontoTotalTmp = 0
        For j = 1 To 4
                Select Case Mid(psMetLiquidacion, j, 1)
                    Case "G"
                        nMontoGastos = MatrizGastosFecha(psCtaCod, MatCalend)
                        nMontoTotalTmp = nMontoTotalTmp + nMontoGastos
                        Call MatrizDistribuirGastos(MatCalend, MatCalendDistrib, nMontoGastos, True)
                    Case "M"
                        nMontoMora = MatrizInteresMorFecha(psCtaCod, MatCalend)
                        nMontoTotalTmp = nMontoTotalTmp + nMontoMora
                        Call MatrizDistribuirMora(MatCalend, MatCalendDistrib, nMontoMora)
                    Case "I"
                        nMontoInteres = CDbl(Format(MatrizInteresTotalesAFechaSinMora(psCtaCod, MatCalend, pdHoy), "#0.00"))
                        nMontoTotalTmp = nMontoTotalTmp + nMontoInteres
                        Call MatrizDistribuirInteresI(MatCalend, MatCalendDistrib, nMontoInteres, True)
                    Case "C"
                        nMontoCapital = MatrizCapitalAFecha(psCtaCod, MatCalend)
                        nMontoTotalTmp = nMontoTotalTmp + nMontoCapital
                        Call MatrizDistribuirCapital(MatCalend, MatCalendDistrib, nMontoCapital, True)
                End Select
        Next j
        Call MatrizActualizarEstadoCalendCancelado(MatCalendDistrib)
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
Dim j As Integer

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
Dim j As Integer

        MatCalendDistrib = CrearMatrizparaAmortizacion(MatCalend)
        nMontoTotalTmp = 0
        nMontoPago = pnMontoPago
        For j = 1 To 4
            Select Case Mid(psMetLiquidacion, j, 1)
                Case "G"
                    nMontoGastos = MatrizGastosFecha(psCtaCod, MatCalend)
                    nMontoPago = nMontoPago - nMontoGastos
                    Call MatrizDistribuirGastos(MatCalend, MatCalendDistrib, nMontoGastos, True)
                Case "M"
                    nMontoMora = MatrizInteresMorFecha(psCtaCod, MatCalend)
                    nMontoPago = nMontoPago - nMontoMora
                    Call MatrizDistribuirMora(MatCalend, MatCalendDistrib, nMontoMora)
                Case "I"
                    nMontoInteres = CDbl(Format(MatrizInteresTotalesAFechaSinMora(psCtaCod, MatCalend, pdHoy), "#0.00"))
                    nMontoPago = nMontoPago - nMontoInteres
                    Call MatrizDistribuirInteresI(MatCalend, MatCalendDistrib, nMontoInteres, True)
                Case "C"
                    nMontoCapital = MatrizCapitalVencido(MatCalend, pdHoy)
                    nMontoPago = nMontoPago - nMontoCapital
                    Call MatrizDistribuirCapital(MatCalend, MatCalendDistrib, nMontoCapital, True)
            End Select
            nMontoPago = CDbl(Format(nMontoPago, "#0.00"))
            Call MatrizActualizarEstadoCuota(MatCalend, MatCalendDistrib)
        Next j
        If nMontoPago > 0 Then
            Call MatrizDistribuirCapital(MatCalend, MatCalendDistrib, nMontoPago, True)
            Call MatrizActualizarEstadoCuota(MatCalend, MatCalendDistrib)
        End If
        MatrizDistribuirCalendDinamico = MatCalendDistrib
End Function

Public Function MatrizDistribuirMonto(ByVal MatCalend As Variant, ByVal pnMontoPago As Double, _
                    ByVal psMetLiquidacion As String) As Variant
Dim I As Integer
Dim j As Integer
Dim MatCalendDistrib As Variant
Dim nMonto As Double
        nMonto = pnMontoPago
        MatCalendDistrib = CrearMatrizparaAmortizacion(MatCalend)
        Do While nMonto > 0
            For j = 1 To 4
                Select Case Mid(psMetLiquidacion, j, 1)
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
            Next j
        Loop
        MatrizDistribuirMonto = MatCalendDistrib
End Function

Public Function MatrizSaldoCapital(ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant) As Double
Dim I As Integer
    MatrizSaldoCapital = 0
    For I = 0 To UBound(MatCalend) - 1
        MatrizSaldoCapital = MatrizSaldoCapital + (CDbl(MatCalend(I, 3)) - CDbl(MatCalendDistrib(I, 3)))
        MatrizSaldoCapital = CDbl(Format(MatrizSaldoCapital, "#0.00"))
    Next I
End Function

Public Function MatrizMontoCuota(ByVal MatCalend As Variant, ByVal pnNroCuota As Integer) As Double
Dim I As Integer
    MatrizMontoCuota = 0
    For I = 0 To UBound(MatCalend) - 1
        If CInt(MatCalend(I, 1)) = pnNroCuota Then
            MatrizMontoCuota = MatrizMontoCuota + CDbl(MatCalend(I, 3)) + CDbl(MatCalend(I, 4)) + CDbl(MatCalend(I, 5))
            MatrizMontoCuota = CDbl(Format(MatrizMontoCuota, "#0.00"))
            Exit Function
        End If
    Next I
End Function


Public Function MatrizCuotaPendiente(ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant) As Integer
Dim I As Integer
    MatrizCuotaPendiente = 0
    For I = 0 To UBound(MatCalend) - 1
        If CInt(MatCalendDistrib(I, 2)) = gColocCalendEstadoPendiente Then
            MatrizCuotaPendiente = CInt(MatCalendDistrib(I, 1))
            Exit For
        End If
    Next I
End Function
Public Function MatrizCapitalPagado(ByVal MatCalendDistrib As Variant) As Double
Dim I As Integer
    MatrizCapitalPagado = 0
    For I = 0 To UBound(MatCalendDistrib) - 1
        MatrizCapitalPagado = MatrizCapitalPagado + CDbl(MatCalendDistrib(I, 3))
    Next I
End Function

Public Function MatrizIntCompPagado(ByVal MatCalendDistrib As Variant) As Double
Dim I As Integer
    MatrizIntCompPagado = 0
    For I = 0 To UBound(MatCalendDistrib) - 1
        MatrizIntCompPagado = MatrizIntCompPagado + CDbl(MatCalendDistrib(I, 4))
    Next I
End Function

Public Function MatrizIntGraciaPagado(ByVal MatCalendDistrib As Variant) As Double
Dim I As Integer
    MatrizIntGraciaPagado = 0
    For I = 0 To UBound(MatCalendDistrib) - 1
        MatrizIntGraciaPagado = MatrizIntGraciaPagado + CDbl(MatCalendDistrib(I, 5))
    Next I
End Function

Public Function MatrizIntMorPagado(ByVal MatCalendDistrib As Variant) As Double
Dim I As Integer
    MatrizIntMorPagado = 0
    For I = 0 To UBound(MatCalendDistrib) - 1
        MatrizIntMorPagado = MatrizIntMorPagado + CDbl(MatCalendDistrib(I, 6))
    Next I
End Function

Public Function MatrizIntReprogPag(ByVal MatCalendDistrib As Variant) As Double
Dim I As Integer
    MatrizIntReprogPag = 0
    For I = 0 To UBound(MatCalendDistrib) - 1
        MatrizIntReprogPag = MatrizIntReprogPag + CDbl(MatCalendDistrib(I, 7))
    Next I
End Function

Public Function MatrizIntSuspensoPag(ByVal MatCalendDistrib As Variant) As Double
Dim I As Integer
    MatrizIntSuspensoPag = 0
    For I = 0 To UBound(MatCalendDistrib) - 1
        MatrizIntSuspensoPag = MatrizIntSuspensoPag + CDbl(MatCalendDistrib(I, 8))
    Next I
End Function

Public Function MatrizGastoPag(ByVal MatCalendDistrib As Variant) As Double
Dim I As Integer
    MatrizGastoPag = 0
    For I = 0 To UBound(MatCalendDistrib) - 1
        MatrizGastoPag = MatrizGastoPag + CDbl(MatCalendDistrib(I, 9))
    Next I
End Function

Public Function MatrizCuotasPagadas(ByVal MatCalendDistrib As Variant) As String
Dim I As Integer
    MatrizCuotasPagadas = ""
    For I = 0 To UBound(MatCalendDistrib) - 1
        If CInt(MatCalendDistrib(I, 2)) = gColocCalendEstadoPagado Then
            MatrizCuotasPagadas = MatrizCuotasPagadas & MatCalendDistrib(I, 1) & ","
        End If
    Next I
    If Len(MatrizCuotasPagadas) > 0 Then
        MatrizCuotasPagadas = Mid(MatrizCuotasPagadas, 1, Len(MatrizCuotasPagadas) - 1)
    End If
    
End Function
Public Function MatrizFechaCuotaPendiente(ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant) As Date
Dim I As Integer
    For I = 0 To UBound(MatCalend) - 1
        If CInt(MatCalendDistrib(I, 2)) = gColocCalendEstadoPendiente Then
            MatrizFechaCuotaPendiente = CDate(MatCalendDistrib(I, 0))
            Exit For
        End If
    Next I
End Function

Public Function MatrizEstadoCalendario(ByVal MatCalendDistrib As Variant) As Integer
Dim I As Integer
    MatrizEstadoCalendario = gColocCalendEstadoPagado
    For I = 0 To UBound(MatCalendDistrib) - 1
        If CInt(MatCalendDistrib(I, 2)) = gColocCalendEstadoPendiente Then
            MatrizEstadoCalendario = gColocCalendEstadoPendiente
            Exit For
        End If
    Next I
End Function

Public Function CrearMatrizparaAmortizacion(ByVal MatCalend As Variant) As Variant
Dim MatCalendAmortiz() As String
Dim I As Integer
    ReDim MatCalendAmortiz(UBound(MatCalend), 11)
    
    For I = 0 To UBound(MatCalend) - 1
        MatCalendAmortiz(I, 0) = MatCalend(I, 0)
        MatCalendAmortiz(I, 1) = MatCalend(I, 1)
        MatCalendAmortiz(I, 2) = MatCalend(I, 2)
        MatCalendAmortiz(I, 3) = "0.00"
        MatCalendAmortiz(I, 4) = "0.00"
        MatCalendAmortiz(I, 5) = "0.00"
        MatCalendAmortiz(I, 6) = "0.00"
        MatCalendAmortiz(I, 7) = "0.00"
        MatCalendAmortiz(I, 8) = "0.00"
        MatCalendAmortiz(I, 9) = "0.00"
        MatCalendAmortiz(I, 10) = MatCalend(I, 10)
    Next I
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

Function RecuperaMatrizCalendarioInicial(ByVal psCtaCod As String, Optional nMontoInicial As Double = -1) As Variant
Dim R As ADODB.Recordset
Dim RCred As ADODB.Recordset
Dim oCalend As Dcalendario
Dim oCredito As DCredito
Dim MatCalend() As String
Dim nMontoSaldo As Double

    On Error GoTo ErrorRecuperaMatrizCalendarioInicial
    Set oCalend = New Dcalendario
    Set R = oCalend.RecuperaCalendarioPagos(psCtaCod)
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
        MatCalend(R.Bookmark - 1, 0) = Format(R!dVenc, "dd/mm/yyyy")
        MatCalend(R.Bookmark - 1, 1) = Trim(Str(R!nCuota))
        MatCalend(R.Bookmark - 1, 2) = Trim(Str(R!nColocCalendEstado))
        MatCalend(R.Bookmark - 1, 3) = Format(R!nCapital, "#0.00")
        MatCalend(R.Bookmark - 1, 4) = Format(IIf(IsNull(R!nIntComp), 0, R!nIntComp), "#0.00")
        MatCalend(R.Bookmark - 1, 5) = Format(IIf(IsNull(R!nIntGracia), 0, R!nIntGracia), "#0.00")
        MatCalend(R.Bookmark - 1, 6) = Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00")
        MatCalend(R.Bookmark - 1, 7) = Format(IIf(IsNull(R!nIntReprog), 0, R!nIntReprog), "#0.00")
        MatCalend(R.Bookmark - 1, 8) = Format(IIf(IsNull(R!nIntSuspenso), 0, R!nIntSuspenso), "#0.00")
        MatCalend(R.Bookmark - 1, 9) = Format(IIf(IsNull(R!nGasto), 0, R!nGasto), "#0.00")
        nMontoSaldo = nMontoSaldo - R!nCapital
        nMontoSaldo = CDbl(Format(nMontoSaldo, "#0.00"))
        MatCalend(R.Bookmark - 1, 10) = Format(nMontoSaldo, "#0.00")
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    RecuperaMatrizCalendarioInicial = MatCalend
    Exit Function

ErrorRecuperaMatrizCalendarioInicial:
    ReDim MatCalend(0, 0)
    RecuperaMatrizCalendarioInicial = MatCalend
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function


Function RecuperaMatrizCalendarioPendiente(ByVal psCtaCod As String) As Variant
Dim R As ADODB.Recordset
Dim RCred As ADODB.Recordset
Dim oCalend As Dcalendario
Dim oCredito As DCredito
Dim MatCalend() As String
Dim nMontoSaldo As Double

    On Error GoTo ErrorRecuperaMatrizCalendarioPendiente
    Set oCalend = New Dcalendario
    Set R = oCalend.RecuperaCalendarioPagosPendiente(psCtaCod)
    Set oCalend = Nothing
    ReDim MatCalend(R.RecordCount, 11)
    Set oCredito = New DCredito
    Set RCred = oCredito.RecuperaColocaciones(psCtaCod)
    nMontoSaldo = CDbl(Format(RCred!nMontoCol, "#0.00"))
    Set oCredito = Nothing
    Do While Not R.EOF
        MatCalend(R.Bookmark - 1, 0) = Format(R!dVenc, "dd/mm/yyyy")
        MatCalend(R.Bookmark - 1, 1) = Trim(Str(R!nCuota))
        MatCalend(R.Bookmark - 1, 2) = Trim(Str(R!nColocCalendEstado))
        MatCalend(R.Bookmark - 1, 3) = Format(R!nCapital, "#0.00")
        MatCalend(R.Bookmark - 1, 4) = Format(IIf(IsNull(R!nIntComp), 0, R!nIntComp), "#0.00")
        MatCalend(R.Bookmark - 1, 5) = Format(IIf(IsNull(R!nIntGracia), 0, R!nIntGracia), "#0.00")
        MatCalend(R.Bookmark - 1, 6) = Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00")
        MatCalend(R.Bookmark - 1, 7) = Format(IIf(IsNull(R!nIntReprog), 0, R!nIntReprog), "#0.00")
        MatCalend(R.Bookmark - 1, 8) = Format(IIf(IsNull(R!nIntSuspenso), 0, R!nIntSuspenso), "#0.00")
        MatCalend(R.Bookmark - 1, 9) = Format(IIf(IsNull(R!nGasto), 0, R!nGasto), "#0.00")
        nMontoSaldo = nMontoSaldo - R!nSaldoCap
        nMontoSaldo = CDbl(Format(nMontoSaldo, "#0.00"))
        MatCalend(R.Bookmark - 1, 10) = Format(nMontoSaldo, "#0.00")
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    RecuperaMatrizCalendarioPendiente = MatCalend
    Exit Function

ErrorRecuperaMatrizCalendarioPendiente:
    ReDim MatCalend(0, 0)
    RecuperaMatrizCalendarioPendiente = MatCalend
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function MatrizGastosVencidos(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim I As Integer
    MatrizGastosVencidos = 0
    For I = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(0, 0)) Then
            MatrizGastosVencidos = MatrizGastosVencidos + CDbl(MatCalend(I, 3))
        End If
    Next I
End Function

Public Function MatrizIntCompVencido(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim I As Integer
    MatrizIntCompVencido = 0
    For I = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(0, 0)) Then
            MatrizIntCompVencido = MatrizIntCompVencido + CDbl(MatCalend(I, 4))
        End If
    Next I
End Function

Public Function MatrizCapitalVencido(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim I As Integer
    MatrizCapitalVencido = 0
    For I = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(0, 0)) Then
            MatrizCapitalVencido = MatrizCapitalVencido + CDbl(MatCalend(I, 9))
        End If
    Next I
End Function

Public Function MatrizDeudaAlaFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
    MatrizDeudaAlaFecha = MatrizCapitalVencido(MatCalend, pdHoy) + MatrizInteresGastosAFecha(psCtaCod, MatCalend, pdHoy)
End Function

Public Function MatrizMoraTotal(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim I As Integer
    MatrizMoraTotal = 0
    For I = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(0, 0)) Then
            MatrizMoraTotal = MatrizMoraTotal + CDbl(MatCalend(I, 6))
        End If
    Next I
End Function

Public Function MatrizCuotasEnMora(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Integer
Dim I As Integer
    MatrizCuotasEnMora = 0
    For I = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(0, 0)) Then
            MatrizCuotasEnMora = MatrizCuotasEnMora + 1
        End If
    Next I
End Function

Public Function MatrizMontoCalendDinamico(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date)
    
    MatrizMontoCalendDinamico = 0
    'MatrizMontoCalendDinamico = MatrizMontoCalendDinamico + MatrizCapitalVencido(MatCalend, pdHoy) _
    '    + MatrizGastosVencidos(MatCalend, pshoy) + MatrizInteresTotalesAFecha(psCtaCod, MatCalend, pdHoy)
    MatrizMontoCalendDinamico = MatrizMontoCalendDinamico + MatrizCapitalVencido(MatCalend, pdHoy) _
        + MatrizGastosFecha(psCtaCod, MatCalend) + MatrizInteresTotalesAFecha(psCtaCod, MatCalend, pdHoy)
End Function

'**********************************************************************
'* Funcion que trabaja conla Matriz de Calendario ya cargado
'*  y calcula lo que tiene que pagar el Cliente para Ponerse al Dia
'***********************************************************************
Public Function MatrizMontoAPagar(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim I As Integer
Dim j As Integer
    MatrizMontoAPagar = 0
    For I = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(I, 0)) Then
            For j = 3 To 9
                MatrizMontoAPagar = MatrizMontoAPagar + CDbl(MatCalend(I, j))
            Next j
        End If
    Next I
    If MatrizMontoAPagar = 0 Then
        For j = 3 To 9
            MatrizMontoAPagar = MatrizMontoAPagar + CDbl(MatCalend(0, j))
        Next j
    End If
End Function

Private Function ValidaAmortizacionCredito(ByVal pnMonto As Double) As String
    ValidaAmortizacionCredito = ""
    If pnMonto = 0 Then
        ValidaAmortizacionCredito = "Monto de Pago debe ser Mayor que Cero"
    End If
End Function

Public Function AmortizarCredito(ByVal psCtaCod, ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant, _
            ByVal pnMonto As Double, ByVal pdHoy As Date, ByVal psMetLiquid As String, _
            ByVal pnTipoPago As ColocTipoPago, ByVal psCodAge As String, ByVal psCodUser As String, _
            Optional psNroDoc As String = "", Optional ByVal pConnBase As DCredActualizaBD = Nothing, _
            Optional ByVal pnMovNro As Long = -1, Optional ByVal pnNroDacion As Integer = -1, Optional pbEnOtraCmac As Boolean = False, _
            Optional psPersCmac As String = "", Optional ByVal pnIntPend As Double = 0, Optional ByVal pnIntPendPagado As Double = 0) As String
            
Dim oBase As DCredActualizaBD
Dim nEstadoCred As Integer
Dim R As ADODB.Recordset
Dim oCred As DCredito
Dim oCalend As Dcalendario
Dim nTransacc As Long
Dim sLineaCred As String
Dim nMontoColocado As Double
Dim dFecPend As Date
Dim nDiasAtraso As Integer
Dim I As Integer
Dim nNroCalen As Integer
Dim nMontoGasto As Double
Dim sMovNro As String
Dim nMovNro As Long
Dim nConsCred As String
Dim pnPlazo As Integer
Dim bTran As Boolean
Dim oFunciones As New NContFunciones
Dim dFechaTran As Date
Dim nIntPend As Double

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
    nTransacc = R!nTransacc
    R.Close
    Set R = Nothing
    
    If pnTipoPago <> gColocTipoPagoDacionPago Then
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
                    nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorNorEfec, gCredPagNorNorCh)
                Else
                    nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorNorEOCEfec, gCredPagNorNorEOCCh)
                End If
            Case gColocEstVigNorm
                If Not pbEnOtraCmac Then
                    nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorMorEfec, gCredPagNorMorCh)
                Else
                    nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorMorEOCEfec, gCredPagNorMorEOCCh)
                End If
            Case gColocEstVigVenc
                If Not pbEnOtraCmac Then
                    nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorVenEfec, gCredPagNorVenCh)
                Else
                    nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorVenEOCEfec, gCredPagNorVenEOCCh)
                End If
        End Select
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
    R.Close
    Set R = Nothing
    
    Set oFunciones = New NContFunciones
    sMovNro = oFunciones.GeneraMovNro(pdHoy, psCodAge, psCodUser)
    Set oFunciones = Nothing
    
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
    dFechaTran = CDate(Format(Format(pdHoy, "dd/mm/yyyy") & " " & Format(oBase.dFechaHora, "hh:mm:ss"), "mm/dd/yyyy hh:mm:ss"))
    
    'Actualiza Producto
    If MatrizEstadoCalendario(MatCalendDistrib) = gColocCalendEstadoPagado Then
        nEstadoCred = gColocEstCancelado
    End If
    Call oBase.dUpdateProducto(psCtaCod, , MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCred, pdHoy, nTransacc + 1, False)
    
    'Actualiza ColocLineaCreditoSaldo
    Set oCred = New DCredito
    Set R = oCred.RecuperaColocaciones(psCtaCod)
    Set oCred = Nothing
    sLineaCred = R!cLineaCred
    R.Close
    Set R = Nothing
    Set R = oBase.RecuperaLineasCreditoSaldo(sLineaCred, True)
    nMontoColocado = R!nMontoColocado
    R.Close
    Set R = Nothing
    nMontoColocado = nMontoColocado - pnMonto
    Call oBase.dUpdateLineaCreditoSaldo(sLineaCred, , , nMontoColocado, , , False)
    
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
    Call oBase.dInsertMov(sMovNro, nConsCred, "", gMovEstContabMovContable, gMovFlagVigente, False)
    nMovNro = oBase.dGetnMovNro(sMovNro)
    '************************************************************************************
    'Si es una Pago en Otra Cmac
    '************************************************************************************
    If pbEnOtraCmac Then
        Call oBase.dInsertMovCMAC(nMovNro, psPersCmac, Format$(gTpoIFCmac, "00"), CInt(Mid(psCtaCod, 9, 1)), "", psNroDoc, nConsCred, pnMonto, False)
    End If
    '************************************************************************************
    'Si es Cancelado con un Desembolso Insertar su Referencia
    '************************************************************************************
    If Not pConnBase Is Nothing Then
        Call oBase.dInsertMovRef(pnMovNro, nMovNro, False)
    End If
    
    If pnTipoPago = gColocTipoPagoCheque Then
        Call oBase.dInsertaMovDoc(nMovNro, TpoDocCheque, psNroDoc, dFechaTran, False)
    End If
    
    If pnTipoPago <> gColocTipoPagoDacionPago Then
        Call oBase.dInsertMovCol(nMovNro, nConsCred, psCtaCod, nNroCalen, pnMonto, 0, psMetLiquid, pnPlazo, MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCred, False)
    Else
        Call oBase.dInsertMovCol(nMovNro, nConsCred, psCtaCod, nNroCalen, pnMonto, 0, psMetLiquid, pnPlazo, MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCred, False, pnNroDacion)
    End If
    
    'Actualiza calendario (ColocCalendario y ColocCalendDet)
    For I = 0 To UBound(MatCalendDistrib) - 1
        Call oBase.dUpdateColocCalendario(psCtaCod, nNroCalen, CInt(MatCalendDistrib(I, 1)), gColocCalendAplCuota, , CInt(MatCalendDistrib(I, 2)), "Pago de Cuota", gColocCalendConceptoProcAprobado, False)
        
        'Amortizando Capital
        If CDbl(MatCalendDistrib(I, 3)) > 0 Then
            Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(I, 1)), gColocConceptoCodCapital, , CDbl(MatCalendDistrib(I, 3)), , False, True)
            'Inserta Detalle Movimiento Capital
            Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, nNroCalen, gColocConceptoCodCapital, CInt(MatCalendDistrib(I, 1)), CDbl(MatCalendDistrib(I, 3)), False)
        End If
        'Amortizando Interes Compensatorio
        If CDbl(MatCalendDistrib(I, 4)) > 0 Then
            Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(I, 1)), gColocConceptoCodInteresCompensatorio, , CDbl(MatCalendDistrib(I, 4)), , False, True)
            'Inserta Detalle Movimiento Interes Compensatorio
            Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, nNroCalen, gColocConceptoCodInteresCompensatorio, CInt(MatCalendDistrib(I, 1)), CDbl(MatCalendDistrib(I, 4)), False)
        End If
        'Amortizando Interes Gracia
        If CDbl(MatCalendDistrib(I, 5)) > 0 Then
            Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(I, 1)), gColocConceptoCodInteresCompensatorio, , CDbl(MatCalendDistrib(I, 5)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, nNroCalen, gColocConceptoCodInteresGracia, CInt(MatCalendDistrib(I, 1)), CDbl(MatCalendDistrib(I, 5)), False)
        End If
        'Amortizando Interes Moratorio
        If CDbl(MatCalendDistrib(I, 6)) > 0 Then
            Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(I, 1)), gColocConceptoCodInteresCompensatorio, , CDbl(MatCalendDistrib(I, 6)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, nNroCalen, gColocConceptoCodInteresMoratorio, CInt(MatCalendDistrib(I, 1)), CDbl(MatCalendDistrib(I, 6)), False)
        End If
        'Amortizando Interes Reprog
        If CDbl(MatCalendDistrib(I, 7)) > 0 Then
            Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(I, 1)), gColocConceptoCodInteresCompensatorio, , CDbl(MatCalendDistrib(I, 7)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, nNroCalen, gColocConceptoCodInteresReprogramado, CInt(MatCalendDistrib(I, 1)), CDbl(MatCalendDistrib(I, 7)), False)
        End If
        'Amortizando Interes Suspenso
        If CDbl(MatCalendDistrib(I, 8)) > 0 Then
            Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(I, 1)), gColocConceptoCodInteresCompensatorio, , CDbl(MatCalendDistrib(I, 8)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, nNroCalen, gColocConceptoCodInteresSuspenso, CInt(MatCalendDistrib(I, 1)), CDbl(MatCalendDistrib(I, 7)), False)
        End If
        
        'Amortizando Gastos
        If CDbl(MatCalendDistrib(I, 9)) > 0 Then
            nMontoGasto = CDbl(MatCalendDistrib(I, 9))
            Set oCalend = New Dcalendario
            Set R = oCalend.RecuperaCalendarioGastos(psCtaCod, nNroCalen, CInt(MatCalendDistrib(I, 1)), gColocCalendAplCuota)
            Set oCalend = Nothing
            Do While Not R.EOF
                If nMontoGasto >= R!nMonto Then
                    Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(I, 1)), R!nColocConceptoCod, , R!nMonto, , False, True)
                    'Inserta Detalle Movimiento Gastos
                    Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, nNroCalen, R!nColocConceptoCod, CInt(MatCalendDistrib(I, 1)), R!nMonto, False)
                    nMontoGasto = nMontoGasto - R!nMonto
                Else
                    Call oBase.dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(I, 1)), R!nColocConceptoCod, , nMontoGasto, , False, True)
                    'Inserta Detalle Movimiento Gastos
                    Call oBase.dInsertMovColDet(nMovNro, nConsCred, psCtaCod, nNroCalen, R!nColocConceptoCod, CInt(MatCalendDistrib(I, 1)), CDbl(Format(nMontoGasto, "#0.00")), False)
                    nMontoGasto = 0
                End If
                
                nMontoGasto = CDbl(Format(nMontoGasto, "#0.00"))
                If nMontoGasto = 0 Then
                    Exit Do
                End If
                R.MoveNext
             Loop
             R.Close
             Set R = Nothing
        End If
    Next I
        
    If pnTipoPago <> gColocTipoPagoDacionPago Then
    '    Call oBase.dAnularColocGarantRec(pnNroDacion, gColocGarantRecEstadoCancelado, False)
    End If
    
    Call oBase.dCommitTrans
    Set oBase = Nothing
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
    ByVal pnDemanda As Boolean, ByVal pnDiasAtraso As Integer, ByVal pdHoy As Date, _
    ByVal psCodAge As String, ByVal psCodUser As String)
Dim MatCalend As Variant
Dim oBase As DCredActualizaBD
Dim nInteres As Double
Dim nMora As Double
Dim nGasto As Double
Dim dFechaEst As Date
Dim sMovNro As String
Dim oFunciones As NContFunciones
Dim pbTran As Boolean
Dim nMovNro As Long

    On Error GoTo ErrorTransferirARecuperaciones
    pbTran = False
    MatCalend = RecuperaMatrizCalendarioPendiente(psCtaCod)
    nInteres = MatrizInteresGastosAFecha(psCtaCod, MatCalend, pdHoy) - MatrizGastosFecha(psCtaCod, MatCalend) - MatrizInteresMorFecha(psCtaCod, MatCalend)
    nMora = MatrizInteresMorFecha(psCtaCod, MatCalend)
    nGasto = MatrizGastosFecha(psCtaCod, MatCalend)
    
    Set oFunciones = New NContFunciones
    sMovNro = oFunciones.GeneraMovNro(pdHoy, psCodAge, psCodUser)
    Set oFunciones = Nothing
    
    Set oBase = New DCredActualizaBD
    dFechaEst = oBase.dFechaHora
    Call oBase.dBeginTrans
    pbTran = True
    Call oBase.dUpdateProducto(psCtaCod, , , gColocEstJudicial, Format(dFechaEst, "mm/dd/yyyy hh:mm:ss"), , False)
    Call oBase.dInsertExpedJud(psCtaCod, "", 0, "", "", "", "", "", "", "", 0, False)
    Call oBase.dInsertColocCredRecup(psCtaCod, 0, nInteres, nMora, nGasto, 0, "GCIM", 0, 0, pnDemanda, "", False)
    Call oBase.dInsertColocCredRecupEstado(psCtaCod, Format(dFechaEst, "mm/dd/yyyy hh:mm:ss"), gColocEstRecVigJud, False)
    Call oBase.dInsertMov(sMovNro, gCredPasoARecup, "Paso a Judicial", gMovEstContabMovContable, gMovFlagVigente, False)
    nMovNro = oBase.dGetnMovNro(sMovNro)
    Call oBase.dInsertMovCol(nMovNro, gCredPasoARecup, psCtaCod, 0, pnSaldoCap, pnDiasAtraso, "", 0, 0, False)
    Call oBase.dInsertMovColDet(nMovNro, gCredPasoARecup, psCtaCod, 0, gColRecConceptoCodCapital, 0, pnSaldoCap, False)
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
            Call oBase.dInsertColocCalendDet(psCtaCod, R!nNroCalen, pnAplicado, R!nCuota, CInt(psCodGasto), nMonto, 0, "", False)
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
    Call oBase.dInsertColocCalificacionAnalista(psCtaCod, dFecha, pnNota, psComent, False)
    Set oBase = Nothing
    
    Exit Sub

ErrorActualizarNotaCredito:
        Err.Raise Err.Number, "Error En Proceso ActualizarNotaCredito", Err.Description
    
End Sub

Public Sub ActualizaMetasAnalista(ByVal psCodPers As String, ByVal pnTipoMeta As ColocTipoMetas, _
     ByVal pdInicial As Date, ByVal pdFinal As Date, ByVal pnTipoAct As Integer, Optional pMatMontos As Variant)

Dim I As Integer
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
            For I = 0 To UBound(pMatMontos) - 1
                Call oBase.dInsertColocMetasAnalista(psCodPers, pnTipoMeta, CInt(pMatMontos(I, 1)), pdInicial, pdFinal, CDbl(pMatMontos(I, 0)), False)
            Next I
            oBase.dCommitTrans
        Case 2
            oBase.dBeginTrans
            pbTran = True
            For I = 0 To UBound(pMatMontos) - 1
                Call oBase.dUpdateColocMetasAnalista(psCodPers, pnTipoMeta, CInt(pMatMontos(I, 1)), pdInicial, pdFinal, CDbl(pMatMontos(I, 0)), False)
            Next I
            oBase.dCommitTrans
        Case 3
            Call oBase.dDeleteColocMetasAnalista(psCodPers, pnTipoMeta, pdInicial, pdFinal, False)
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
    Set R = oDCred.RecuperaColocGarantia(psCtaCod)
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
    Set oDCred = Nothing
    
    Set oFunciones = New NContFunciones
    sMovNro = oFunciones.GeneraMovNro(pdHoy, psCodAge, psCodUser)
    Set oFunciones = Nothing
    
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
        Call oBase.dDeleteColocGarantia(R!cTpoDoc, R!cNroDoc, psCtaCod, False)
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
        Call oBase.dInsertMovRef(nMovNro, nMov, False)
    End If
    If pnTipoAcc = 2 And pbRefinanc Then 'Si es Retiro de Refinanciacion
        Do While Not RRef.EOF
            Call oBase.dUpdateProducto(R!cCtaCodRef, , , gColocEstVigVenc, , , False)
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
    ByVal sLpt As String, ByVal psNomInstitucion As String, Optional psNroDoc As String = "", Optional ByVal pbImpBoletas As Boolean = False) As String
    
Dim MatCalend() As Variant
Dim MatCalendDistribuido() As Variant
Dim I As Integer
Dim oBase As DCredActualizaBD
Dim oConstante As DConstante
Dim oDoc As NCredDoc
Dim sMovNro As String
Dim nMovNro As Long
Dim oFunciones As New NContFunciones
Dim nConsCred As String
Dim sError As String
Dim sCadImp As String

    On Error GoTo ErrorAmortizarPagoLote
    
    Set oConstante = New DConstante
    Set oDoc = New NCredDoc
    Set oFunciones = New NContFunciones
    sMovNro = oFunciones.GeneraMovNro(pdHoy, psCodAge, psCodUsu)
    Set oFunciones = Nothing
    
    Set oBase = New DCredActualizaBD
    
    ReDim MatCalend(UBound(MatPagos))
    ReDim MatCalendDistribuido(UBound(MatPagos))
    For I = 0 To UBound(MatPagos) - 1
        MatCalend(I) = RecuperaMatrizCalendarioPendiente(MatPagos(I, 0))
        MatCalendDistribuido(I) = MatrizDistribuirMonto(MatCalend(I), CDbl(MatPagos(I, 1)), MatPagos(I, 4))
    Next I
    
    oBase.dBeginTrans
    'Insert Movimiento
    Call oBase.dInsertMov(sMovNro, gCredPagLote, "PAGO EN LOTE", gMovEstContabMovContable, gMovFlagVigente, False)
    nMovNro = oBase.dGetnMovNro(sMovNro)
    For I = 0 To UBound(MatPagos) - 1
        Select Case CInt(MatPagos(I, 5))
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
        
        sError = oBase.AmortizarCreditoLote(MatPagos(I, 0), MatCalend(I), MatCalendDistribuido(I), CDbl(MatPagos(I, 1)), pdHoy, MatPagos(I, 4), pnTipoPago, psCodAge, psCodUsu, nMovNro, CInt(MatPagos(I, 5)), CLng(MatPagos(I, 2)), IIf(pnTipoPago = gColocTipoPagoCheque, psNroDoc, ""))
        If pbImpBoletas Then
            Call oDoc.ImprimeBoleta(MatPagos(I, 0), MatPagos(I, 3), psNomAge, oConstante.DameDescripcionConstante(gMoneda, CInt(Mid(MatPagos(I, 0), 9, 1))), _
                MatrizCuotasPagadas(MatCalendDistribuido(I)), Format(pdHoy, "dd/mm/yyyy"), Format(pdHoy, "hh:mm:ss"), Trim(Str(CInt(MatPagos(I, 2)) + 1)), "Credito Consumo", _
                MatrizCapitalPagado(MatCalendDistribuido(I)), MatrizIntCompPagado(MatCalendDistribuido(I)), _
                MatrizIntMorPagado(MatCalendDistribuido(I)), MatrizGastoPag(MatCalendDistribuido(I)), _
                MatrizIntGraciaPagado(MatCalendDistribuido(I)), _
                MatrizIntSuspensoPag(MatCalendDistribuido(I)) + MatrizIntReprogPag(MatCalendDistribuido(I)), _
                MatrizSaldoCapital(MatCalend(I), MatCalendDistribuido(I)), MatrizFechaCuotaPendiente(MatCalend(I), MatCalendDistribuido(I)), _
                psCodUsu, sLpt)
        End If
    Next I
    oBase.dCommitTrans
    Set oBase = Nothing
    Set oConstante = Nothing
    sCadImp = oDoc.ImprimePagoLote(MatPagos, MatCalend, MatCalendDistribuido, psNomAge, pdHoy, psCodUsu, psNomInstitucion, IIf(Mid(MatPagos(0, 0), 1, 9) = "1", "SOLES", "DOLARES"))
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
    Optional ByVal pbExtornoEnLote As Boolean = False, Optional ByVal poBase As DCredActualizaBD = Nothing, Optional ByVal pnMovNro As Long = -1)
    
Dim oDCred As DCredito
Dim oBase As DCredActualizaBD
Dim oFun As NContFunciones
Dim nMovNro As Long
Dim sMovNro As String
Dim R As ADODB.Recordset
    
    Set oDCred = New DCredito
    Set R = oDCred.RecuperaDatosExtorno(pnMov, psCtaCod)
    Set oDCred = Nothing
        
    
    If pbExtornoEnLote Then
        Set oBase = poBase
    Else
        Set oFun = New NContFunciones
        sMovNro = oFun.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
        Set oFun = Nothing
        Set oBase = New DCredActualizaBD
        Call oBase.dBeginTrans
    End If
    
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
    
    'Extorna Saldos Calendario
    If psCodOpe = gCredDesembCtaExist Or psCodOpe = gCredDesembCtaNueva Or psCodOpe = gCredDesembCtaExistDOA Or psCodOpe = gCredDesembCtaNuevaDOA Then
        Call oBase.dExtornaSaldosCalendario(pnMov, gColocCalendAplDesembolso, False)
    Else
        Call oBase.dExtornaSaldosCalendario(pnMov, gColocCalendAplCuota, False)
    End If
    
    'Extorna Saldos de Maestros
    'Colocaccred
    Call oBase.dUpdateColocacCred(psCtaCod, R!nDiasMora, , , , , , , R!nMinCuota, , , , , , , R!nMinCuota, , , False)
    'Producto
    Call oBase.dUpdateProducto(psCtaCod, , R!nCapital, R!nCredEstado, , , False, 1)
    
    '*********************************************
    'FALTA EL EXTO>RNO DE ABONO A CUENTA
    '*********************************************
    
    
    '**************************************************************
    'En Caso que el pago halla sido con una Dacion en Pago
    '**************************************************************
    If psCodOpe = gCredPagNorNorDacion Or psCodOpe = gCredPagNorMorDacion Or psCodOpe = gCredPagNorVenDacion _
        Or psCodOpe = gCredPagRefNorDacion Or psCodOpe = gCredPagNorNorDacion Or psCodOpe = gCredPagRefMorDacion Or psCodOpe = gCredPagRefVenDacion Then
        'Call oBase.dAnularColocGarantRec(R!nFlag, gColocGarantRecEstadoRegistrado, False)
    End If
    If Not pbExtornoEnLote Then
        Call oBase.dCommitTrans
    End If
    Set oBase = Nothing
    
    R.Close
    Set R = Nothing
    
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
        Call ExtornarCredito(R!cCtaCod, pdFecSis, psCodUser, psCodAge, pnMovNro, R!COPECOD, R!nMonto, , True, oBase, nMovNro)
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

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing


End Sub
