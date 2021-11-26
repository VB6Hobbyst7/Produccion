VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMFormatosEval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************
' Descripcion: Clase creado para los Formatos de Evaluacion de Créditos
' Referencia: ERS004 -2016
' Creado por: LUCV004-2016 | PEAC
' Fecha Creacion: 20160525
' Ultima Modificacion: 20160525
'************************************************************************

Public Function CargaDatosCredEvaluacion2(ByVal psCtaCod As String, _
                                        ByVal pnFormato As Integer, _
                                        ByRef prsDatCredEval As ADODB.Recordset, _
                                        Optional ByRef prsDatGastoNeg As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatGastoFam As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatOtrosIng As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatRef As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatActivos As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatPasivos As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatCuotaIFIs As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatPropuesta As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsCapacPagoNeta As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatRatios As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatActivoPasivo As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatIfiGastoNego As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatIfiGastoFami As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatVentaCosto As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatFlujoCaja As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatPDT As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatPDTDet As ADODB.Recordset = Nothing, _
                                        Optional ByRef pnCodActivo As Integer, Optional ByRef pnCodPasivo As Integer, _
                                        Optional ByRef prsDatParamFlujoCaja As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatIfiBalActCorri As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatBalGerlTotalPasivo As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatIfiNoSupervisadaGastoNego As ADODB.Recordset = Nothing, Optional ByRef prsDatIfiNoSupervisadaGastoFami As ADODB.Recordset = Nothing) As Boolean
                                        
                                        'prsDatIfiNoSupervisadaGastoNego, prsDatIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agregó
                                        'JOEP20171015 Segun ERS051-2017 Flujo de Caja Optional ByRef prsDatParamFlujoCaja As ADODB.Recordset = Nothing
                                        'JOEP20171015 Segun ERS051-2017 Flujo de Caja Optional ByRef prsDatBalGerlTotalPasivo As ADODB.Recordset = Nothing
                                        
On Error GoTo ErrorCargaDatosCreEval
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim prsColocacCred As ADODB.Recordset
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    CargaDatosCredEvaluacion2 = False
    
    'Carga en la Opcion de Mantenimiento de Formato
    Set prsDatCredEval = oDCOMFormatosEval.RecuperaColocacCredEval(psCtaCod)
    Set prsDatGastoNeg = oDCOMFormatosEval.RecuperaDatosCredEvalGastosNeg(psCtaCod)
    Set prsDatGastoFam = oDCOMFormatosEval.RecuperaDatosCredEvalGastosFam(psCtaCod)
    Set prsDatOtrosIng = oDCOMFormatosEval.RecuperaDatosCredEvalOtrosIngr(psCtaCod)
    Set prsDatCuotaIFIs = oDCOMFormatosEval.RecuperaDatosCredEvalCuotaIfis(psCtaCod)
    Set prsDatPropuesta = oDCOMFormatosEval.RecuperaDatosCredEvalPropuesta(psCtaCod)
    'Set prsCapacPagoNeta = oDCOMFormatosEval.RecuperaDatosCapacidadPagoNeta(psCtaCod)
    Set prsDatVentaCosto = oDCOMFormatosEval.RecuperaDatosCredEvalVentaCosto(psCtaCod)
    
    Set prsDatRef = oDCOMFormatosEval.RecuperaDatosReferidos(psCtaCod)
    Set prsDatRatios = oDCOMFormatosEval.RecuperaDatosRatios(psCtaCod)
    
    Set prsDatActivos = oDCOMFormatosEval.RecuperaDatosActivos(psCtaCod, pnCodActivo)
    Set prsDatPasivos = oDCOMFormatosEval.RecuperaDatosPasivos(psCtaCod, pnCodPasivo)
    
    Set prsDatActivoPasivo = oDCOMFormatosEval.RecuperaDatosActivoPasivo(psCtaCod)
    Set prsDatFlujoCaja = oDCOMFormatosEval.RecuperaDatosCredEvalFlujoCaja(psCtaCod)
    Set prsDatPDT = oDCOMFormatosEval.RecuperaDatosCredEvalPDT(psCtaCod)
    Set prsDatPDTDet = oDCOMFormatosEval.RecuperaDatosCredEvalPDTDet(psCtaCod)
    
    'CTI3 ERS0032020**************************************************************
    Dim pnFormatoGastosNego As Integer
    Dim pnCodCuotaIfiGastoNego As Integer
    pnFormatoGastosNego = gFormatoGastosNego
    pnCodCuotaIfiGastoNego = gCodCuotaIfiGastoNego
    If pnFormato = 4 Or pnFormato = 5 Or pnFormato = 6 Then
        pnFormatoGastosNego = gFormatoGastosNegoFlujoCaja
        pnCodCuotaIfiGastoNego = gCodCuotaIfiFlujoCaja
    End If
    
    Set prsDatIfiGastoNego = oDCOMFormatosEval.RecuperaDatosIfiCuota(psCtaCod, pnFormato, pnFormatoGastosNego, pnCodCuotaIfiGastoNego) 'CTI320200110 ERS003-2020. Agregó: gCodCuotaIfiGastoNego
    'Set prsDatIfiGastoNego = oDCOMFormatosEval.RecuperaDatosIfiCuota(psCtaCod, pnFormato, gFormatoGastosNego, gCodCuotaIfiGastoNego) 'CTI320200110 ERS003-2020. Agregó: gCodCuotaIfiGastoNego
    '***********************************************************************************************************
    Set prsDatIfiNoSupervisadaGastoNego = oDCOMFormatosEval.RecuperaDatosIfiCuota(psCtaCod, pnFormato, IIf(pnFormato >= 4, gFormatoGastosNegoFlujoCaja, gFormatoGastosNego), IIf(pnFormato >= 4, gCodCuotaIfiNoSupervisadaFlujoCaja, gCodCuotaIfiNoSupervisadaGastoNego)) 'CTI320200110 ERS003-2020. Agregó
    
    Set prsDatIfiGastoFami = oDCOMFormatosEval.RecuperaDatosIfiCuota(psCtaCod, pnFormato, gFormatoGastosFami, gCodCuotaIfiGastoFami) 'CTI320200110 ERS003-2020. Agregó: gCodCuotaIfiGastoFami
    Set prsDatIfiNoSupervisadaGastoFami = oDCOMFormatosEval.RecuperaDatosIfiCuota(psCtaCod, pnFormato, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami) 'CTI320200110 ERS003-2020. Agregó
    
  'JOEP20171015 Segun ERS051-2017 Flujo de Caja
    If pnFormato = 3 Or pnFormato = 4 Or pnFormato = 5 Or pnFormato = 6 Then
        Set prsDatParamFlujoCaja = oDCOMFormatosEval.RecuperaDatosParamFlujoCaja(psCtaCod, pnFormato)
    End If
    
    Set prsDatIfiBalActCorri = oDCOMFormatosEval.RecuperaDatosIfiCuota(psCtaCod, pnFormato, 7026, 109)
    Set prsDatBalGerlTotalPasivo = oDCOMFormatosEval.RecuperaDatosIfiCuota(psCtaCod, pnFormato, 7026, 201)
  'JOEP20171015 Segun ERS051-2017 Flujo de Caja
    
    CargaDatosCredEvaluacion2 = True
Exit Function
ErrorCargaDatosCreEval:
    MsgBox Err.Description, vbCritical, "Error"
End Function

'LUCV20160623
Public Function CargaDatosFlexEdit(ByVal pnFormato As Integer, ByVal pcCtaCod As String, _
                                   Optional ByRef prsFeGastoNeg As ADODB.Recordset = Nothing, _
                                   Optional ByRef prsDatGastoFam As ADODB.Recordset = Nothing, _
                                   Optional ByRef prsFeDatOtrosIng As ADODB.Recordset = Nothing, _
                                   Optional ByRef prsFeDatBalanGen As ADODB.Recordset = Nothing, _
                                   Optional ByRef prsDatActivos As ADODB.Recordset = Nothing, _
                                   Optional ByRef prsDatPasivos As ADODB.Recordset = Nothing, _
                                   Optional ByRef prsDatPasivosNo As ADODB.Recordset = Nothing, _
                                   Optional ByRef prsDatPatrimonio As ADODB.Recordset = Nothing, _
                                   Optional ByRef prsFeDatRef As ADODB.Recordset = Nothing, _
                                   Optional ByRef prsFeFlujoCaja As ADODB.Recordset = Nothing, _
                                   Optional ByRef prsFeDatActivosForm6 As ADODB.Recordset = Nothing, _
                                   Optional ByRef rsFeDatPasivosForm6 As ADODB.Recordset = Nothing, _
                                   Optional ByRef rsFeDatEstadoGanPerdForm6 As ADODB.Recordset = Nothing, _
                                   Optional ByRef rsFeDatCoeficienteFinanForm6 As ADODB.Recordset = Nothing, _
                                   Optional ByRef prsFeDatPDT As ADODB.Recordset = Nothing, _
                                   Optional ByRef prsDatRemBrutaTotalMes1 As ADODB.Recordset = Nothing, _
                                   Optional ByRef prsDatRemBrutaTotalMes2 As ADODB.Recordset = Nothing, _
                                   Optional ByVal pnTpoInsConv As Integer, Optional ByVal bSecSalud As Boolean, _
                                   Optional ByRef prsFeFlujoCajaHistorico As ADODB.Recordset = Nothing, _
                                   Optional ByRef prsDatParamFlujoCaja As ADODB.Recordset = Nothing) As Boolean 'LUCV20171015, ERS05120171015


On Error GoTo ErrorCargaDatosCreEval
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim prsDatGastoNeg As ADODB.Recordset
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    
    'CargaDatosCredEvaluacion2 = False
    If pnFormato = 1 Or pnFormato = 2 Or pnFormato = 3 Or pnFormato = 4 Or pnFormato = 5 Or pnFormato = 7 Then
        Set prsFeGastoNeg = oDCOMFormatosEval.RecuperaDatosFlexGastosNegocio(pcCtaCod, pnFormato) 'LUCV20160713
        Set prsDatGastoFam = oDCOMFormatosEval.RecuperaDatosFlexGastosFamiliares(pcCtaCod, pnFormato) 'LUCV20160713 / JOEP
        Set prsFeDatOtrosIng = oDCOMFormatosEval.RecuperaDatosFlexOtrosIngresos(pnFormato) 'LUCV20160713 / JOEP
        Set prsFeDatBalanGen = oDCOMFormatosEval.RecuperaDatosFlexBalanceGeneral(pcCtaCod, pnFormato) 'LUCV20160713
        Set prsDatActivos = oDCOMFormatosEval.RecuperaDatosFlexActivos(pnFormato) 'LUCV20160713
        Set prsDatPasivos = oDCOMFormatosEval.RecuperaDatosFlexPasivos(pcCtaCod, pnFormato) 'LUCV20160713
        Set prsFeFlujoCaja = oDCOMFormatosEval.RecuperaDatosFlexFlujoCaja(pnFormato, pcCtaCod) 'LUCV20160713
        Set prsFeDatPDT = oDCOMFormatosEval.RecuperaDatosFlexFlexDeclaracionPdt(pnFormato) 'LUCV20160713
    End If
    
    If pnFormato = 6 Then
        Set rsFeDatPasivosForm6 = oDCOMFormatosEval.ObtenerFormatosEvalPasivos(pcCtaCod, pnormato) 'PEAC 20160704
        Set prsFeDatActivosForm6 = oDCOMFormatosEval.ObtenerFormatosEvalActivos(pcCtaCod, pnFormato) 'PEAC 20160704
        Set rsFeDatEstadoGanPerdForm6 = oDCOMFormatosEval.ObtenerFormatosEvalEstadoGanPerd(pcCtaCod, pnFormato) 'PEAC 20160704
        Set rsFeDatCoeficienteFinanForm6 = oDCOMFormatosEval.ObtenerFormatosEvalCoeficienteFinan(pcCtaCod, pnFormato) 'PEAC 20160706
        
        Set prsFeDatOtrosIng = oDCOMFormatosEval.RecuperaDatosFlexOtrosIngresos(pnFormato) 'LUCV20160713 / JOEP
        Set prsFeFlujoCaja = oDCOMFormatosEval.RecuperaDatosFlexFlujoCaja(pnFormato, pcCtaCod) 'LUCV20160713
        Set prsFeFlujoCajaHistorico = oDCOMFormatosEval.RecuperaDatosFlexFlujoCaja(pnFormato, pcCtaCod, 1) 'LUCV20171015
        Set prsDatGastoFam = oDCOMFormatosEval.RecuperaDatosFlexGastosFamiliares(pcCtaCod, pnFormato) 'LUCV20160713 / JOEP
        Set prsFeDatPDT = oDCOMFormatosEval.RecuperaDatosFlexFlexDeclaracionPdt(pnFormato) 'LUCV20160713
        Set prsDatParamFlujoCaja = oDCOMFormatosEval.RecuperaDatosParamFlujoCaja(pcCtaCod, pnFormato) 'LUCV20171015, ERS0512017
    End If
    
    If pnFormato = 8 Then
        Set prsDatRemBrutaTotalMes1 = oDCOMFormatosEval.RecuperaDatosFlexRemBrutaTotalMes1(pcCtaCod, pnFormato, pnTpoInsConv, bSecSalud) 'JOEP
        Set prsDatRemBrutaTotalMes2 = oDCOMFormatosEval.RecuperaDatosFlexRemBrutaTotalMes2(pcCtaCod, pnFormato, pnTpoInsConv, bSecSalud) 'JOEP
        'Set prsFeGastoNeg = oDCOMFormatosEval.RecuperaDatosFlexGastosNegocio(pcCtaCod, pnFormato) 'CTI320200110 ERS003-2020. Agregó
        Set prsDatGastoFam = oDCOMFormatosEval.RecuperaDatosFlexGastosFamiliares(pcCtaCod, pnFormato) 'CTI320200110 ERS003-2020. Agregó
        Set prsFeDatOtrosIng = oDCOMFormatosEval.RecuperaDatosFlexOtrosIngresos(pnFormato) 'LUCV20160713 / JOEP
    End If
    
    If pnFormato = 9 Then 'CTI320200110 ERS003-2020. Agregó
        Set prsFeGastoNeg = oDCOMFormatosEval.RecuperaDatosFlexGastosNegocio(pcCtaCod, pnFormato)
        Set prsDatGastoFam = oDCOMFormatosEval.RecuperaDatosFlexGastosFamiliares(pcCtaCod, pnFormato)
    End If
    
CargaDatosFlexEdit = True
Exit Function
ErrorCargaDatosCreEval:
    MsgBox Err.Description, vbCritical, "Error"
End Function

Public Function ObtieneTipoPermisoCredEval(ByVal pCargo As String) As Integer
On Error GoTo ErrorObtieneTipoPermisoCredEval
    Dim oGen As COMDConstSistema.DCOMGeneral
    'Dim sCargosAnalistas As String, sCargosCoord As String, sCargosJef As String 'LUCV, Comentó
    'Dim sCodA() As String, sCodC() As String, sCodJ() As String 'LUCV, Comentó
    Dim i As Integer
    Dim sCargosFormatosEval As String
    Dim sCodCargo() As String
    
    Set oGen = New COMDConstSistema.DCOMGeneral
        'LUCV20160805, Comentó. Según Coordinado Usuario LIPA ******
        '    sCargosAnalistas = oGen.LeeConstSistema(gConstSistRHCargoCodAnalistasCredEval)
        '    sCargosCoord = oGen.LeeConstSistema(gConstSistRHCargoCodCoordCredEval)
        '    sCargosJef = oGen.LeeConstSistema(gConstSistRHCargoCodJefCredEval)
        '
        '    sCodA = Split(sCargosAnalistas, ",")
        '    sCodC = Split(sCargosCoord, ",")
        '    sCodJ = Split(sCargosJef, ",")
        '    ObtieneTipoPermisoCredEval = 0
        '
        '    For i = 0 To UBound(sCodA)
        '        If pCargo = sCodA(i) Then
        '            ObtieneTipoPermisoCredEval = 3
        '            Exit Function
        '        End If
        '    Next i
        '    For i = 0 To UBound(sCodC)
        '        If pCargo = sCodC(i) Then
        '            ObtieneTipoPermisoCredEval = 2
        '            Exit Function
        '        End If
        '    Next i
        '    For i = 0 To UBound(sCodJ)
        '        If pCargo = sCodJ(i) Then
        '            ObtieneTipoPermisoCredEval = 1
        '            Exit Function
        '        End If
        '    Next i
        'Fin LUCV20160805 *****
    sCargosFormatosEval = oGen.LeeConstSistema(gConstSistRHCargoCodAnalistasCredEval)
    sCodCargo = Split(sCargosFormatosEval, ",")
    ObtieneTipoPermisoCredEval = 0
    
    For i = 0 To UBound(sCodCargo)
        If pCargo = sCodCargo(i) Then
            ObtieneTipoPermisoCredEval = 3
            Exit Function
        End If
    Next i

    Exit Function
ErrorObtieneTipoPermisoCredEval:
    MsgBox Err.Description, vbCritical, "Error"
End Function

'LUCV20162206
Public Function GrabarCredFormEvalFormato1_5(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnInsertaActualiza As Integer, _
                                             ByVal psGiroNeg As String, ByVal pnExpEmpAnio As Integer, ByVal pnExpEmpMes As Integer, _
                                             ByVal pnTiempoLocalAnio As Integer, ByVal pnTiempoLocalMes As Integer, ByVal pnUltEndeuda As Double, _
                                             ByVal psFecUltEndeuda As String, ByVal pnCondLocal As Integer, ByVal psCondLocalOtros As String, _
                                             ByVal pnExposicion As Double, ByVal psFecEvaluacion As String, ByVal psFechaVisita As String, _
                                             ByVal psEntorFami As String, ByVal psGiroUbi As String, ByVal psExperiencia As String, _
                                             ByVal psFormalidad As String, ByVal psColaterales As String, ByVal psDestino As String, _
                                             ByVal psComentario As String, ByVal pMatReferidos As Variant, ByVal pMatCuotaOtraIfi As Variant, _
                                             ByVal pMatCuotaIfiGastoFami As Variant, ByVal prsGastoFam As ADODB.Recordset, ByVal prsOtrosIng As ADODB.Recordset, _
                                             Optional ByVal prsGastoNego As ADODB.Recordset = Nothing, Optional ByVal pnIngreso As Double = -1, _
                                             Optional ByVal pnEgreso As Double = -1, Optional ByVal pnMargen As Double = -1, _
                                             Optional ByVal pMatActiPasivo As Variant = Nothing, Optional ByVal pMatActiPasivoDet As Variant = Nothing, _
                                             Optional ByVal prsFlujoCaja As Variant = Nothing, Optional ByVal prsPDT As ADODB.Recordset, _
                                             Optional ByVal pnCodCapaPago As RatioIndicador = -1, Optional ByVal pnCapPago As Double = -1, _
                                             Optional ByVal pnCodEndeu As RatioIndicador = -1, Optional ByVal pnEndeu As Double = -1, _
                                             Optional ByVal pnCodIngrNego As RatioIndicador = -1, Optional ByVal pnIngreNetoNego As Double, _
                                             Optional ByVal pnCodExecMens As RatioIndicador = -1, Optional ByVal pnExceMensual As Double = -1, _
                                             Optional ByVal pnMes1 As Integer, Optional ByVal pnMes2 As Integer, _
                                             Optional ByVal pnMes3 As Integer, Optional ByVal pnAnio1 As Integer, _
                                             Optional ByVal pnAnio2 As Integer, Optional ByVal pnAnio3 As Integer, Optional ByVal pnCondicion As Integer = -1, _
                                             Optional ByVal pMatFlujoCaja As Variant, _
                                             Optional ByVal pMatBalActCorr As Variant, _
                                             Optional ByVal pMatBalGerlTotalPasivo As Variant, _
                                             Optional ByVal pMatIfiNoSupervisadaGastoNego As Variant = Nothing, _
                                             Optional ByVal pMatIfiNoSupervisadaGastoFami As Variant = Nothing) As Boolean
                                             
                                             'pMatIfiNoSupervisadaGastoNego, pMatIfiNoSupervisadaGastoFami'CTI320200110 ERS003-2020. Agregó
                                             'Optional ByVal pnCodTotaPat As Integer = -1, Optional ByVal nMontoPatrimonial As Currency) As Boolean 'LUCV, Comentó
                                             'Optional ByVal pMatFlujoCaja As Variant 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                                             'Optional ByVal pMatBalGerlTotalPasivo As Variant 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja

'Valido para Formato 1, 2, 3, 4, 5
Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
Dim i As Integer
On Error GoTo ErrorGrabarCredFormEvalFormato1_5
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    
    'oCredito.dBeginTrans
    'Registra / Actualiza Datos de la Cabecera(Informacion del Negocio) del Formato
    Call oDCOMFormatosEval.GrabarCredFormEvalCabeceraFormato1_5(psCtaCod, _
                                                                pnInsertaActualiza, _
                                                                psGiroNeg, _
                                                                pnExpEmpAnio, _
                                                                pnExpEmpMes, _
                                                                pnTiempoLocalAnio, _
                                                                pnTiempoLocalMes, _
                                                                pnUltEndeuda, _
                                                                psFecUltEndeuda, _
                                                                pnCondLocal, _
                                                                psCondLocalOtros, _
                                                                pnExposicion, _
                                                                psFecEvaluacion, _
                                                                pnFormato, _
                                                                psComentario)
    'Registra la Propuesta del Credito a Evaluar
    If pnCondicion <> 4 Then 'Registra si los crédito no son Refinanciados
        Call oDCOMFormatosEval.GrabarCredFormEvalPropuestaFormato1_5(psCtaCod, _
                                                                    pnInsertaActualiza, _
                                                                    psFechaVisita, _
                                                                    psEntorFami, _
                                                                    psGiroUbi, _
                                                                    psExperiencia, _
                                                                    psFormalidad, _
                                                                    psColaterales, _
                                                                    psDestino, _
                                                                    pnFormato)
    End If
    'Registra los Ratios/Indicadores Formato1_6
        Call oDCOMFormatosEval.GrabarCredFormEvalRatioIndicador1_5(psCtaCod, _
                                                                pnInsertaActualiza, _
                                                                pnCodCapaPago, _
                                                                pnCapPago, _
                                                                pnCodEndeu, _
                                                                pnEndeu, _
                                                                pnCodIngrNego, _
                                                                pnIngreNetoNego, _
                                                                pnCodExecMens, _
                                                                pnExceMensual, _
                                                                pnFormato)
    'Registrar Ventas y Costos (Formato 1,2,3)
    If pnIngreso <> -1 And pnEgreso <> -1 And pnMargen <> -1 Then
        Call oDCOMFormatosEval.EliminaCredFormEvalVentasCostos(psCtaCod, pnFormato)
        Call oDCOMFormatosEval.GuardarFormatoSinConvenio_NegocioVentasCostos(psCtaCod, pnFormato, pnIngreso, pnEgreso, pnMargen)
    End If
                                                                                                       
    'Para Guardar las Grillas (GastoNegocio, GastoFamiliares, OtrosIngresos, BalanceGeneral, Referidos)
    If Not prsGastoNego Is Nothing Then
    Call oDCOMFormatosEval.EliminaCredFormEvalGrillaGastoNegocio(psCtaCod, pnFormato)
        For i = 0 To prsGastoNego.RecordCount - 1
            Call oDCOMFormatosEval.InsertaCredFormEvalGastoNego(psCtaCod, pnFormato, prsGastoNego!n, CCur(prsGastoNego!Monto))
            prsGastoNego.MoveNext
        Next i
    End If
    
    If Not prsGastoFam Is Nothing Then
    Call oDCOMFormatosEval.EliminaCredFormEvalGrillaGastoFami(psCtaCod, pnFormato)
        For i = 0 To prsGastoFam.RecordCount - 1
            Call oDCOMFormatosEval.InsertaCredFormEvalGastoFami(psCtaCod, pnFormato, prsGastoFam!n, CCur(prsGastoFam!Monto))
            prsGastoFam.MoveNext
        Next i
    End If
    
    If Not prsOtrosIng Is Nothing Then
    Call oDCOMFormatosEval.EliminaCredFormEvalGrillaOtrosIngre(psCtaCod, pnFormato)
        For i = 0 To prsOtrosIng.RecordCount - 1
            Call oDCOMFormatosEval.InsertaCredFormEvalOtrosIngr(psCtaCod, pnFormato, prsOtrosIng!n, CCur(prsOtrosIng!Monto))
            prsOtrosIng.MoveNext
        Next i
    End If

    'Activos | Pasivos (Formato: 1 -3) - TOTALES
    If IsArray(pMatActiPasivo) Then
        If UBound(pMatActiPasivo) - 1 Then
        Call oDCOMFormatosEval.EliminaCredFormEvalGrillaActiPasi(psCtaCod, pnFormato)
            For i = 1 To UBound(pMatActiPasivo) - 1
                Call oDCOMFormatosEval.InsertaCredFormEvalActiPasi(psCtaCod, pnFormato, psFecEvaluacion, IIf(pMatActiPasivo(i, 1) = "", 0, pMatActiPasivo(i, 1)), pMatActiPasivo(i, 2), pMatActiPasivo(i, 5))
            Next i
        End If
        'If pnFormato = 1 Then
        'Inserta Patrimonio
        'Call oDCOMFormatosEval.InsertaCredFormEvalActiPasi(psCtaCod, pnFormato, psFecEvaluacion, gFormatoPasivos, pnCodTotaPat, nMontoPatrimonial)
        'End If
    End If
    'Activos | Pasivos(Formato: 1- 3)- DETALLES
    If IsArray(pMatActiPasivoDet) Then 'Pregunto Matriz no Sea Vacia
         If UBound(pMatActiPasivoDet) - 1 Then
        Call oDCOMFormatosEval.EliminaCredFormEvalGrillaActiPasiDet(psCtaCod, pnFormato)
            For i = 1 To UBound(pMatActiPasivoDet) - 1
                Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiDet(psCtaCod, pnFormato, psFecEvaluacion, IIf(pMatActiPasivoDet(i, 1) = "", 0, pMatActiPasivoDet(i, 1)), pMatActiPasivoDet(i, 2), pMatActiPasivoDet(i, 5))
                'feBalanceGeneral2.MoveNext
            Next i
        End If
    End If
    'Cuotas Otras IFIs ->GastosNegocio / FlujoCajaMensual
    If IsArray(pMatCuotaOtraIfi) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatCuotaOtraIfi) Then
        Call oDCOMFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(psCtaCod, pnFormato, IIf(pnFormato >= 4, gFormatoGastosNegoFlujoCaja, gFormatoGastosNego), IIf(pnFormato >= 4, gCodCuotaIfiFlujoCaja, gCodCuotaIfiGastoNego)) 'CTI320200110 ERS003-2020. Agregó:condicional: pnFormato >= 4
            For i = 0 To UBound(pMatCuotaOtraIfi) - 1
                Call oDCOMFormatosEval.InsertaCredFormEvalCuotaIFI(psCtaCod, pnFormato, IIf(pnFormato >= 4, gFormatoGastosNegoFlujoCaja, gFormatoGastosNego), IIf(pnFormato >= 4, gCodCuotaIfiFlujoCaja, gCodCuotaIfiGastoNego), pMatCuotaOtraIfi(i, 0), pMatCuotaOtraIfi(i, 1), CCur(pMatCuotaOtraIfi(i, 2))) 'CTI320200110 ERS003-2020. Agregó:condicional: pnFormato >= 4
            Next i
        End If
    End If
    
    'Cuotas IFIs ->GastosFamiliar
    If IsArray(pMatCuotaIfiGastoFami) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatCuotaIfiGastoFami) Then
        Call oDCOMFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(psCtaCod, pnFormato, gFormatoGastosFami)
            For i = 0 To UBound(pMatCuotaIfiGastoFami) - 1
                Call oDCOMFormatosEval.InsertaCredFormEvalCuotaIFI(psCtaCod, pnFormato, gFormatoGastosFami, gCodCuotaIfiGastoFami, pMatCuotaIfiGastoFami(i, 0), pMatCuotaIfiGastoFami(i, 1), CDbl(pMatCuotaIfiGastoFami(i, 2)))
            Next i
        End If
    End If
    
    
    'CTI320200110 ERS003-2020. Agregó:
    'Cuotas Otras IFIs No supervisadas (Negocio)
    If IsArray(pMatIfiNoSupervisadaGastoNego) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatIfiNoSupervisadaGastoNego) Then
        Call oDCOMFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(psCtaCod, pnFormato, IIf(pnFormato >= 4, gFormatoGastosNegoFlujoCaja, gFormatoGastosNego), IIf(pnFormato >= 4, gCodCuotaIfiNoSupervisadaFlujoCaja, gCodCuotaIfiNoSupervisadaGastoNego))  'CTI320200110 ERS003-2020. Agregó:gCodCuotaIfiNoSupervisadaGastoNego
            For i = 0 To UBound(pMatIfiNoSupervisadaGastoNego) - 1
                Call oDCOMFormatosEval.InsertaCredFormEvalCuotaIFI(psCtaCod, pnFormato, IIf(pnFormato >= 4, gFormatoGastosNegoFlujoCaja, gFormatoGastosNego), IIf(pnFormato >= 4, gCodCuotaIfiNoSupervisadaFlujoCaja, gCodCuotaIfiNoSupervisadaGastoNego), pMatIfiNoSupervisadaGastoNego(i, 0), pMatIfiNoSupervisadaGastoNego(i, 1), CCur(pMatIfiNoSupervisadaGastoNego(i, 2)))
            Next i
        End If
    End If
    
    'Cuotas de Ifis No supervisadas (Familiar)
    If IsArray(pMatIfiNoSupervisadaGastoFami) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatIfiNoSupervisadaGastoFami) Then
        Call oDCOMFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(psCtaCod, pnFormato, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami)
            For i = 0 To UBound(pMatIfiNoSupervisadaGastoFami) - 1
                Call oDCOMFormatosEval.InsertaCredFormEvalCuotaIFI(psCtaCod, pnFormato, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami, pMatIfiNoSupervisadaGastoFami(i, 0), pMatIfiNoSupervisadaGastoFami(i, 1), CDbl(pMatIfiNoSupervisadaGastoFami(i, 2)))
            Next i
        End If
    End If
    'Fin CTI320200110
    
    
    'Referidos
    If IsArray(pMatReferidos) Then
    If UBound(pMatReferidos) Then
        Call oDCOMFormatosEval.EliminaCredFormEvalGrillaReferidos(psCtaCod, pnFormato)
        For i = 1 To UBound(pMatReferidos)
            Call oDCOMFormatosEval.InsertaCredFormEvalReferidos(psCtaCod, pnFormato, IIf(pMatReferidos(i, 1) = "", 0, pMatReferidos(i, 1)), pMatReferidos(i, 2), pMatReferidos(i, 3), pMatReferidos(i, 4), pMatReferidos(i, 5), pMatReferidos(i, 6))
        Next i
    End If
    End If
    'Flujo de Caja Mensual - agregar constantes
    If Not prsFlujoCaja Is Nothing Then
        Call oDCOMFormatosEval.EliminaCredFormEvalFlujoCaja(psCtaCod, pnFormato)
        For i = 0 To prsFlujoCaja.RecordCount - 1
            Call oDCOMFormatosEval.InsertaCredFormFlujoCajaMensual(psCtaCod, pnFormato, prsFlujoCaja!nConsCod, prsFlujoCaja!nConsValor, prsFlujoCaja!concepto, prsFlujoCaja!Monto)
            prsFlujoCaja.MoveNext
        Next i
    End If
        
    'Declaracion PDT (Formato 4, 5)
    If Not prsPDT Is Nothing Then
        Call oDCOMFormatosEval.EliminaCredFormEvalPDT(psCtaCod, pnFormato)
        Call oDCOMFormatosEval.EliminaCredFormEvalPDTDet(psCtaCod, pnFormato)
        Call oDCOMFormatosEval.InsertaCredFormEvalPDT(psCtaCod, pnFormato, prsPDT!nConsCod, prsPDT!nConsValor, pnMes1, pnMes2, pnMes3, pnAnio1, pnAnio2, pnAnio3)
        For i = 0 To prsPDT.RecordCount - 1
            Call oDCOMFormatosEval.InsertaCredFormEvalPDTDet(psCtaCod, pnFormato, prsPDT!nConsCod, prsPDT!nConsValor, CLng(IIf(Trim(prsPDT!Mes1) = "", 0, prsPDT!Mes1)), _
            CLng(IIf(Trim(prsPDT!Mes2) = "", 0, prsPDT!Mes2)), CLng(IIf(Trim(prsPDT!Mes3) = "", 0, prsPDT!Mes3)), CLng(IIf(Trim(prsPDT!Promedio) = "", 0, prsPDT!Promedio)), _
            CDbl(IIf(Trim(prsPDT!VentasDeclaradas) = "", 0, Replace(prsPDT!VentasDeclaradas, "%", ""))))
            prsPDT.MoveNext
        Next i
    End If
    
    'Agrego Inicio JOEP20171015 Segun ERS051-2017 Flujo de Caja
    If IsArray(pMatFlujoCaja) Then
        If UBound(pMatFlujoCaja) Then
                For i = 1 To UBound(pMatFlujoCaja)
                    Call oDCOMFormatosEval.InsertaCredFormEvalParamFlujoCaja(psCtaCod, pnFormato, pMatFlujoCaja(i, 1), pMatFlujoCaja(i, 2), pMatFlujoCaja(i, 3), pMatFlujoCaja(i, 4), pMatFlujoCaja(i, 5))
                Next i
        End If
    End If
    'Cuotas IFIs ->Balance Activo Corriente
    If IsArray(pMatBalActCorr) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatBalActCorr) Then
            Call oDCOMFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(psCtaCod, pnFormato, 7026, 109)
                For i = 0 To UBound(pMatBalActCorr) - 1
                    Call oDCOMFormatosEval.InsertaCredFormEvalCuotaIFI(psCtaCod, pnFormato, 7026, 109, pMatBalActCorr(i, 0), pMatBalActCorr(i, 1), CDbl(pMatBalActCorr(i, 2)))
                Next i
        Else
            Call oDCOMFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(psCtaCod, pnFormato, 7026, 109)
        End If
    Else
        Call oDCOMFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(psCtaCod, pnFormato, 7026, 109)
    End If
    'Cuotas IFIs ->Balance General Total Pasivo Form1, Form2 Activo no Corriente
    If IsArray(pMatBalGerlTotalPasivo) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatBalGerlTotalPasivo) Then
            Call oDCOMFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(psCtaCod, pnFormato, 7026, 201)
                For i = 0 To UBound(pMatBalGerlTotalPasivo) - 1
                    Call oDCOMFormatosEval.InsertaCredFormEvalCuotaIFI(psCtaCod, pnFormato, 7026, 201, pMatBalGerlTotalPasivo(i, 0), pMatBalGerlTotalPasivo(i, 1), CDbl(pMatBalGerlTotalPasivo(i, 2)))
                Next i
        Else
            Call oDCOMFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(psCtaCod, pnFormato, 7026, 201)
        End If
    Else
        Call oDCOMFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(psCtaCod, pnFormato, 7026, 201)
    End If
    'Agrego Fin JOEP20171015 Segun ERS051-2017 Flujo de Caja
    
    'oCredito.dCommitTrans
    Set oCredito = Nothing
    GrabarCredFormEvalFormato1_5 = True
    Exit Function
ErrorGrabarCredFormEvalFormato1_5:
    'oCredito.dRollbackTrans
    GrabarCredFormEvalFormato1_5 = False
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'Fin LUCV20160622

''JOEP20160530
''-----------------------------------------------------
'Public Function ActualziarfrmCredFormEvalFormatoLineaCredAutomatica(ByVal cCtaCod As String, ByVal nCodForm As Integer, ByVal dFechaExpeCaja As Date, _
'                                                        ByVal cLinea As String)
'
'    Dim oCredFormEval As COMDCredito.DCOMFormatosEval
'    Set oCredFormEval = New COMDCredito.DCOMFormatosEval
'
'    oCredito.dBeginTrans
'
'    Call oCredFormEval.ActualizarfrmCredFormEvalFormatoLineaCredAutomatica(cCtaCod, nCodForm, Format(dFechaExpeCaja, "yyyyMMdd"), cCtaCod)
'
'    oCredito.dCommitTrans
'    Set oCredito = Nothing
'
'    ActualziarfrmCredFormEvalFormatoLineaCredAutomatica = True
'    Exit Function
'    oCredito.dRollbackTrans
'    ActualziarfrmCredFormEvalFormatoLineaCredAutomatica = False
'    Err.Raise Err.Number, "Error En Proceso", Err.Description
'End Function
'
'
''JOEP20160530
'Public Function GrabarfrmCredFormEvalFormatoLineaCredAutomatica(ByVal cCtaCod As String, ByVal nCodForm As Integer, ByVal cActividad As String, _
'                                                        ByVal nExposiCred As Double, ByVal nUltEndeSBS As Double, ByVal dUltEndeuSBS As Date, ByVal dFechaExpeCaja As Date, _
'                                                        ByVal cLinea As String)
'
'    Dim oCredFormEval As COMDCredito.DCOMFormatosEval
'    Set oCredFormEval = New COMDCredito.DCOMFormatosEval
'
'    oCredito.dBeginTrans
'
'    Call oCredFormEval.GrabarfrmCredFormEvalFormatoLineaCredAutomatica(cCtaCod, nCodForm, cActividad, nExposiCred, nUltEndeSBS, Format(dUltEndeuSBS, "yyyyMMdd"), Format(dFechaExpeCaja, "yyyyMMdd"), cLinea)
'
'    oCredito.dCommitTrans
'    Set oCredito = Nothing
'
'    GrabarfrmCredFormEvalFormatoLineaCredAutomatica = True
'    Exit Function
'    oCredito.dRollbackTrans
'    GrabarfrmCredFormEvalFormatoLineaCredAutomatica = False
'    Err.Raise Err.Number, "Error En Proceso", Err.Description
'End Function
'
''JOEP20160530
'Public Function ActualziarCredFormEvalFormatoNuevaLineaAutomatica(ByVal cNumLineaAut As String, ByVal cCtaCod As String, ByVal nCodForm As Integer, ByVal cPersCodLinea As String, _
'                                                            ByVal cDescripcion As String, ByVal dFecha As Date, ByVal nMontoMax As Double, _
'                                                            ByVal nCuotaMax As Double, ByVal nPlazoMax As Integer, ByVal nMoneda As Integer, _
'                                                            ByVal nTasaInt As Double, ByVal cPersCodAna As String, ByVal nMontoRiesgoUnico As Double, ByVal cObservacion As String, ByVal dEstado As Date, ByVal nEstado As Integer, _
'                                                            Optional ByVal rsPersonaFirma As ADODB.Recordset, Optional ByVal rsGarantias As ADODB.Recordset)
'
'
'Dim oCredFormEval As COMDCredito.DCOMFormatosEval
'
'Set oCredFormEval = New COMDCredito.DCOMFormatosEval
'
'Dim i As Integer
'
'
'    oCredito.dBeginTrans
'
'     Call oCredFormEval.GrabarfrmCredFormEvalFormatoNuevaLineaAutomaticaEStado(cNumLineaAut, nMontoMax, dEstado, nEstado, cPersCodLinea)
'
'     Call oCredFormEval.ActualziarfrmCredFormEvalFormatoNuevaLineaAutomatica(cNumLineaAut, cCtaCod, nCodForm, cPersCodLinea, cDescripcion, Format(dFecha, "yyyyMMdd"), nMontoMax, nCuotaMax _
'                                                                    , nPlazoMax, nMoneda, nTasaInt, cPersCodAna, nMontoRiesgoUnico, cObservacion, dEstado, nEstado)
'
'    If Not rsPersonaFirma Is Nothing Then
'            rsPersonaFirma.MoveFirst
'        For i = 0 To rsPersonaFirma.RecordCount - 1
'            Call oCredFormEval.ActualziarfrmCredFormEvalFormatoNuevaLineaAutomatica_PersonaFirma(cNumLineaAut, cCtaCod, 10, rsPersonaFirma!cPersCod)
'            rsPersonaFirma.MoveNext
'        Next i
'    End If
'
'    If Not rsGarantias Is Nothing Then
'            rsGarantias.MoveFirst
'        For i = 0 To rsGarantias.RecordCount - 1
'            Call oCredFormEval.ActualziarfrmCredFormEvalFormatoNuevaLineaAutomatica_Garantias(cNumLineaAut, cCtaCod, 10, rsGarantias!cNumGarant)
'            rsGarantias.MoveNext
'        Next i
'    End If
'
'oCredito.dCommitTrans
'    Set oCredito = Nothing
'
'    ActualziarCredFormEvalFormatoNuevaLineaAutomatica = True
'    Exit Function
''ErrorGrabarCredEvaluacionFormatoIyII:
'    oCredito.dRollbackTrans
'    ActualziarCredFormEvalFormatoNuevaLineaAutomatica = False
'    Err.Raise Err.Number, "Error En Proceso", Err.Description
'
'End Function
'
''JOEP20160530
'Public Function GrabarfrmCredFormEvalFormatoNuevaLineaAutomatica(ByVal cNumLineaAut As String, ByVal cCtaCod As String, ByVal nCodForm As Integer, ByVal cPersCodLinea As String, _
'                                                                 ByVal cDescripcion As String, ByVal dFecha As Date, ByVal nMontoMax As Double, _
'                                                                 ByVal nCuotaMax As Double, ByVal nPlazoMax As Integer, ByVal nMoneda As Integer, _
'                                                                 ByVal nTasaInt As Double, ByVal cPersCodAna As String, ByVal nMontoRU, ByVal cObservacion As String, ByVal dEstado As Date, ByVal nEstado As Integer, _
'                                                                 Optional ByVal rsPersonaFirma As ADODB.Recordset, Optional ByVal rsGarantias As ADODB.Recordset) As Boolean
'Dim oCredFormEval As COMDCredito.DCOMFormatosEval
'
'Set oCredFormEval = New COMDCredito.DCOMFormatosEval
'
'Dim i As Integer
'
'
'    oCredito.dBeginTrans
'
'    Call oCredFormEval.GrabarfrmCredFormEvalFormatoNuevaLineaAutomaticaEStado(cNumLineaAut, nMontoMax, dEstado, nEstado, cPersCodLinea)
'
'    Call oCredFormEval.GrabarfrmCredFormEvalFormatoNuevaLineaAutomatica(cNumLineaAut, cCtaCod, nCodForm, cPersCodLinea, cDescripcion, Format(dFecha, "yyyyMMdd"), nMontoMax, nCuotaMax _
'                                                                    , nPlazoMax, nMoneda, nTasaInt, cPersCodAna, nMontoRU, cObservacion, dEstado, nEstado)
'
'    If Not rsPersonaFirma Is Nothing Then
'            rsPersonaFirma.MoveFirst
'        For i = 0 To rsPersonaFirma.RecordCount - 1
'            Call oCredFormEval.GrabarfrmCredFormEvalFormatoNuevaLineaAutomatica_PersonaFirma(cNumLineaAut, cCtaCod, 10, rsPersonaFirma!cPersCod)
'            rsPersonaFirma.MoveNext
'        Next i
'    End If
'
'    If Not rsGarantias Is Nothing Then
'            rsGarantias.MoveFirst
'        For i = 0 To rsGarantias.RecordCount - 1
'            Call oCredFormEval.GrabarfrmCredFormEvalFormatoNuevaLineaAutomatica_Garantias(cNumLineaAut, cCtaCod, 10, rsGarantias!cNumGarant)
'            rsGarantias.MoveNext
'        Next i
'    End If
'
'    oCredito.dCommitTrans
'    Set oCredito = Nothing
'
'    GrabarfrmCredFormEvalFormatoNuevaLineaAutomatica = True
'    Exit Function
''ErrorGrabarCredEvaluacionFormatoIyII:
'    oCredito.dRollbackTrans
'    GrabarfrmCredFormEvalFormatoNuevaLineaAutomatica = False
'    Err.Raise Err.Number, "Error En Proceso", Err.Description
'End Function

'JOEP20160530
Public Function ActualizarfrmCredFormEvalFormatoParalelo(ByVal cCtaCod As String, ByVal nCodForm As Integer, ByVal dFechaExpeCaja As Date, _
                                                         ByVal nMontoApro As Double, ByVal nSaldoActual As Double, ByVal nVentas As Double, ByVal nCapPago As Double, ByVal nIngresoNeto As Double, _
                                                         ByVal nIncreIngresoDatos As Double, ByVal nIngreEstMontos As Double, ByVal nIncreIngresoEstiMontos As Double, ByVal nMargenbruto As Double, ByVal nOtrsoGastos As Double, _
                                                         ByVal nConsuFamili As Double, ByVal nCuotaCredVig As Double, ByVal nOtrosIng As Double, ByVal nMargenBrutoCaja As Double, ByVal nIngreResumen As Double, ByVal nIncreIngreResumen As Double, _
                                                         ByVal nMontoParalelo As Double, ByVal nMontoPropuesto As Double, _
                                                         ByVal dFecVisita As Date, ByVal cDestino As String, ByVal cEntornoFami As String, ByVal cGiroUbicacion As String, ByVal cExpeCrediticia As String, ByVal cFormalidadNegocio As String, ByVal cGarantias As String, _
                                                         ByVal prsGastoNego As ADODB.Recordset, ByVal prsGastoFami As ADODB.Recordset, _
                                                         ByVal pMatCuotaOtraIfi As Variant, ByVal pMatCuotaIfiGastoFami As Variant, _
                                                         ByVal pMatIfiNoSupervisadaGastoNego As Variant, ByVal pMatIfiNoSupervisadaGastoFami As Variant) As Boolean
                                                         'prsGastoNego, prsGastoFami, pMatCuotaOtraIfi, pMatCuotaIfiGastoFami, pMatIfiNoSupervisadaGastoNego, pMatIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020, Agregó
    
    Dim oCredFormEval As COMDCredito.DCOMFormatosEval
    Set oCredFormEval = New COMDCredito.DCOMFormatosEval
    
'    oCredito.dBeginTrans
    
        Call oCredFormEval.ActualizarfrmCredFormFormatoParalelo(cCtaCod, nCodForm, dFechaExpeCaja)
    
    'If Not rsEvaluacion Is Nothing Then
        Call oCredFormEval.ActualizarfrmCredFormFormatoParaleloEvaluacion(cCtaCod, nCodForm, nMontoApro, nSaldoActual, nVentas, nCapPago, nIngresoNeto, nIncreIngresoDatos, _
                                                                    nIngreEstMontos, nIncreIngresoEstiMontos, nMargenbruto, nOtrsoGastos, nConsuFamili, nCuotaCredVig, nOtrosIng, nMargenBrutoCaja, nIngreResumen, nIncreIngreResumen, _
                                                                    nMontoParalelo, nMontoPropuesto)
        
        Call oCredFormEval.ActualizarfrmCredFormFormatoParaleloPropuestaCredito(cCtaCod, nCodForm, dFecVisita, cDestino, cEntornoFami, cGiroUbicacion, cExpeCrediticia, cFormalidadNegocio, cGarantias)
   ' End If

    'oCredito.dCommitTrans
    
    'CTI320200110 ERS003-2020. Agregó:
    'Gastos del Negocio
    If Not prsGastoNego Is Nothing Then
    Call oCredFormEval.EliminaCredFormEvalGrillaGastoNegocio(cCtaCod, nCodForm)
        For i = 0 To prsGastoNego.RecordCount - 1
            Call oCredFormEval.InsertaCredFormEvalGastoNego(cCtaCod, nCodForm, prsGastoNego!n, CCur(prsGastoNego!Monto))
            prsGastoNego.MoveNext
        Next i
    End If
    'Gastos familiares
    If Not prsGastoFami Is Nothing Then
    Call oCredFormEval.EliminaCredFormEvalGrillaGastoFami(cCtaCod, nCodForm)
        For i = 0 To prsGastoFami.RecordCount - 1
            Call oCredFormEval.InsertaCredFormEvalGastoFami(cCtaCod, nCodForm, prsGastoFami!n, CCur(prsGastoFami!Monto))
            prsGastoFami.MoveNext
        Next i
    End If
    
    'Cuotas Otras IFIs ->GastosNegocio
    If IsArray(pMatCuotaOtraIfi) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatCuotaOtraIfi) Then
        Call oCredFormEval.EliminaCredFormEvalGrillaOtrasIfis(cCtaCod, nCodForm, gFormatoGastosNego, gCodCuotaIfiGastoNego) 'CTI320200110 ERS003-2020. Agregó:condicional: pnFormato >= 4
            For i = 0 To UBound(pMatCuotaOtraIfi) - 1
                Call oCredFormEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosNego, gCodCuotaIfiGastoNego, pMatCuotaOtraIfi(i, 0), pMatCuotaOtraIfi(i, 1), CCur(pMatCuotaOtraIfi(i, 2)))
            Next i
        End If
    End If
    
    'Cuotas IFIs ->GastosFamiliar
    If IsArray(pMatCuotaIfiGastoFami) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatCuotaIfiGastoFami) Then
        Call oCredFormEval.EliminaCredFormEvalGrillaOtrasIfis(cCtaCod, nCodForm, gFormatoGastosFami)
            For i = 0 To UBound(pMatCuotaIfiGastoFami) - 1
                Call oCredFormEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiGastoFami, pMatCuotaIfiGastoFami(i, 0), pMatCuotaIfiGastoFami(i, 1), CDbl(pMatCuotaIfiGastoFami(i, 2)))
            Next i
        End If
    End If
    
    'Cuotas Otras IFIs No supervisadas (Negocio)
    If IsArray(pMatIfiNoSupervisadaGastoNego) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatIfiNoSupervisadaGastoNego) Then
        Call oCredFormEval.EliminaCredFormEvalGrillaOtrasIfis(cCtaCod, nCodForm, gFormatoGastosNego, gCodCuotaIfiNoSupervisadaGastoNego)
            For i = 0 To UBound(pMatIfiNoSupervisadaGastoNego) - 1
                Call oCredFormEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosNego, gCodCuotaIfiNoSupervisadaGastoNego, pMatIfiNoSupervisadaGastoNego(i, 0), pMatIfiNoSupervisadaGastoNego(i, 1), CCur(pMatIfiNoSupervisadaGastoNego(i, 2)))
            Next i
        End If
    End If
    
    'Cuotas de Ifis No supervisadas (Familiar)
    If IsArray(pMatIfiNoSupervisadaGastoFami) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatIfiNoSupervisadaGastoFami) Then
        Call oCredFormEval.EliminaCredFormEvalGrillaOtrasIfis(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami)
            For i = 0 To UBound(pMatIfiNoSupervisadaGastoFami) - 1
                Call oCredFormEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami, pMatIfiNoSupervisadaGastoFami(i, 0), pMatIfiNoSupervisadaGastoFami(i, 1), CDbl(pMatIfiNoSupervisadaGastoFami(i, 2)))
            Next i
        End If
    End If
    'Fin CTI320200110
    
    
    Set oCredito = Nothing
    
    ActualizarfrmCredFormEvalFormatoParalelo = True
    Exit Function
ErrorGrabarCredEvaluacionFormatoIyII:
    'oCredito.dRollbackTrans
    ActualizarfrmCredFormEvalFormatoParalelo = False
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'JOEP20160530
Public Function GrabarfrmCredFormEvalFormatoParalelo(ByVal cCtaCod As String, ByVal nCodForm As Integer, ByVal cActividad As String, _
                                                    ByVal dFechaExpeCaja As Date, ByVal nUltEndeSBS As Double, ByVal nNCreditos As Integer, _
                                                    ByVal dUltEndeuSBS As Date, ByVal cCampaña As String, ByVal nExposiCred As Double, _
                                                    ByVal rsEvaluacion As ADODB.Recordset, ByVal nVentas As Double, ByVal nCapPago As Double, ByVal nIngresoNetoas As Double, _
                                                    ByVal nIncreIngresoDatos As Double, ByVal nIngreEstMontos As Double, ByVal nIncreIngresoEstiMontos As Double, ByVal nMargenbruto As Double, ByVal nOtrsoGastos As Double, ByVal nConsuFamili As Double, ByVal nCuota As Double, _
                                                    ByVal nOtrosIng As Double, ByVal nMargenBrutoCaja As Double, ByVal nIngreResumen As Double, ByVal nIncreIngreResumen As Double, ByVal nMontoParalelo As Double, _
                                                    ByVal dFecVisita As Date, ByVal cDestino As String, ByVal cEntornoFami As String, ByVal cGiroUbicacion As String, ByVal cExpeCrediticia As String, ByVal cFormalidadNegocio As String, ByVal cGarantias As String, _
                                                    ByVal prsGastoNego As ADODB.Recordset, ByVal prsGastoFami As ADODB.Recordset, _
                                                    ByVal pMatCuotaOtraIfi As Variant, ByVal pMatCuotaIfiGastoFami As Variant, _
                                                    ByVal pMatIfiNoSupervisadaGastoNego As Variant, ByVal pMatIfiNoSupervisadaGastoFami As Variant) As Boolean
                                                   'prsGastoNego, prsGastoFami, pMatCuotaOtraIfi, pMatCuotaIfiGastoFami, pMatIfiNoSupervisadaGastoNego, pMatIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020, Agregó
                                                   
 Dim oCredFormEval As COMDCredito.DCOMFormatosEval
 Set oCredFormEval = New COMDCredito.DCOMFormatosEval
    
    'oCredito.dBeginTrans
    
        Call oCredFormEval.GrabarfrmCredFormEvalFormatoParalelo(cCtaCod, 9, cActividad, Format(dFechaExpeCaja, "yyyyMMdd"), nUltEndeSBS _
                                                       , nNCreditos, Format(dUltEndeuSBS, "yyyyMMdd"), cCampaña, nExposiCred)
    
    'If Not rsEvaluacion Is Nothing Then
        Call oCredFormEval.dGrabarfrmCredFormEvalFormatoParaleloEvaluacion(cCtaCod, 9, rsEvaluacion!nMontoCol, rsEvaluacion!nSaldo, nVentas, nCapPago, nIngresoNetoas, _
                                                        nIncreIngresoDatos, nIngreEstMontos, nIncreIngresoEstiMontos, nMargenbruto, nOtrsoGastos, nConsuFamili, nCuota, nOtrosIng, nMargenBrutoCaja, nIngreResumen, nIncreIngreResumen, nMontoParalelo, rsEvaluacion!nMontoPro)
                                                        
        
        Call oCredFormEval.GrabarfrmCredFormEvalFormatoParaleloPropuestaCredito(cCtaCod, 9, Format(dFecVisita, "yyyyMMdd"), cDestino, cEntornoFami, cGiroUbicacion, cExpeCrediticia, cFormalidadNegocio, cGarantias)
   ' End If
    
    'CTI320200110 ERS003-2020. Agregó:
    'Gastos del Negocio
    If Not prsGastoNego Is Nothing Then
    Call oCredFormEval.EliminaCredFormEvalGrillaGastoNegocio(cCtaCod, nCodForm)
        For i = 0 To prsGastoNego.RecordCount - 1
            Call oCredFormEval.InsertaCredFormEvalGastoNego(cCtaCod, nCodForm, prsGastoNego!n, CCur(prsGastoNego!Monto))
            prsGastoNego.MoveNext
        Next i
    End If
    'Gastos familiares
    If Not prsGastoFami Is Nothing Then
    Call oCredFormEval.EliminaCredFormEvalGrillaGastoFami(cCtaCod, nCodForm)
        For i = 0 To prsGastoFami.RecordCount - 1
            Call oCredFormEval.InsertaCredFormEvalGastoFami(cCtaCod, nCodForm, prsGastoFami!n, CCur(prsGastoFami!Monto))
            prsGastoFami.MoveNext
        Next i
    End If
    
    'Cuotas Otras IFIs ->GastosNegocio
    If IsArray(pMatCuotaOtraIfi) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatCuotaOtraIfi) Then
        Call oCredFormEval.EliminaCredFormEvalGrillaOtrasIfis(cCtaCod, nCodForm, gFormatoGastosNego, gCodCuotaIfiGastoNego) 'CTI320200110 ERS003-2020. Agregó:condicional: pnFormato >= 4
            For i = 0 To UBound(pMatCuotaOtraIfi) - 1
                Call oCredFormEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosNego, gCodCuotaIfiGastoNego, pMatCuotaOtraIfi(i, 0), pMatCuotaOtraIfi(i, 1), CCur(pMatCuotaOtraIfi(i, 2)))
            Next i
        End If
    End If
    
    'Cuotas IFIs ->GastosFamiliar
    If IsArray(pMatCuotaIfiGastoFami) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatCuotaIfiGastoFami) Then
        Call oCredFormEval.EliminaCredFormEvalGrillaOtrasIfis(cCtaCod, nCodForm, gFormatoGastosFami)
            For i = 0 To UBound(pMatCuotaIfiGastoFami) - 1
                Call oCredFormEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiGastoFami, pMatCuotaIfiGastoFami(i, 0), pMatCuotaIfiGastoFami(i, 1), CDbl(pMatCuotaIfiGastoFami(i, 2)))
            Next i
        End If
    End If
    
    'Cuotas Otras IFIs No supervisadas (Negocio)
    If IsArray(pMatIfiNoSupervisadaGastoNego) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatIfiNoSupervisadaGastoNego) Then
        Call oCredFormEval.EliminaCredFormEvalGrillaOtrasIfis(cCtaCod, nCodForm, gFormatoGastosNego, gCodCuotaIfiNoSupervisadaGastoNego)
            For i = 0 To UBound(pMatIfiNoSupervisadaGastoNego) - 1
                Call oCredFormEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosNego, gCodCuotaIfiNoSupervisadaGastoNego, pMatIfiNoSupervisadaGastoNego(i, 0), pMatIfiNoSupervisadaGastoNego(i, 1), CCur(pMatIfiNoSupervisadaGastoNego(i, 2)))
            Next i
        End If
    End If
    
    'Cuotas de Ifis No supervisadas (Familiar)
    If IsArray(pMatIfiNoSupervisadaGastoFami) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatIfiNoSupervisadaGastoFami) Then
        Call oCredFormEval.EliminaCredFormEvalGrillaOtrasIfis(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami)
            For i = 0 To UBound(pMatIfiNoSupervisadaGastoFami) - 1
                Call oCredFormEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami, pMatIfiNoSupervisadaGastoFami(i, 0), pMatIfiNoSupervisadaGastoFami(i, 1), CDbl(pMatIfiNoSupervisadaGastoFami(i, 2)))
            Next i
        End If
    End If
    'Fin CTI320200110
   
   ' oCredito.dCommitTrans
    Set oCredito = Nothing
    
    GrabarfrmCredFormEvalFormatoParalelo = True
    Exit Function
ErrorGrabarCredEvaluacionFormatoIyII:
    'oCredito.dRollbackTrans
    GrabarfrmCredFormEvalFormatoParalelo = False
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'-----------------------------------------------------

'JOEP20160530
Public Function CargaDatosCredEvaluacionFP(ByVal psCtaCod As String, _
                                        ByVal pnFormato As Integer, _
                                        ByRef prsDatCredEval As ADODB.Recordset, _
                                        Optional ByRef prsDatGastoNeg As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatGastoFam As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatOtrosIng As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatRef As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatActivos As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatPasivos As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatPasivosNo As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatPDT As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatPDTDet As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatPatrimonio As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatPasPat As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatEstadoGP As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatRatios As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsDatIngNeg As ADODB.Recordset = Nothing) As Boolean


On Error GoTo ErrorCargaDatosCredEvaluacionFP
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim prsColocacCred As ADODB.Recordset
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set prsDatCredEval = oDCOMFormatosEval.RecuperarDatosCredEvalFormatoParalelo(psCtaCod)
   

    CargaDatosCredEvaluacionFP = True
    
Exit Function
ErrorCargaDatosCredEvaluacionFP:
    MsgBox Err.Description, vbCritical, "Error"
End Function

'JOEP20160530
Public Function ActualizarConsumoConConvenio_InfCliente(ByVal cCtaCod As String, ByVal nCodForm As Integer, ByVal cActividad As String, ByVal nAntgAnios As Integer, ByVal nAntgMes As Integer, _
                                                        ByVal nUltEndeSBS As Double, ByVal dUltEndeuSBS As Date, ByVal nTipoAportacion As Integer, ByVal nTipoInstitucion As Integer, _
                                                        ByVal nSectorSalud As Integer, ByVal nTipoPlanilla As Integer, ByVal cInstiConv As String, _
                                                        ByVal nMontoSol As Double, ByVal nNumCuotas As Integer, ByVal nExposiCred As Double, ByVal dFecFteIng As Date, _
                                                        ByVal dFechaMes1 As Integer, ByVal nAnoMes1 As Integer, ByVal nRemBruTotalMes1 As Double, ByVal nDescMes1 As Double, _
                                                        ByVal nIngrNetoMes1 As Double, ByVal dFechaMes2 As Integer, ByVal nAnoMes2 As Integer, ByVal nRemBruTotalMes2 As Double, _
                                                        ByVal nDescMes2 As Double, ByVal nIngrNetoMes2 As Double, ByVal nRemBruTotalPromedio As Double, ByVal nDescuentoPromedio As Double, _
                                                        ByVal nIngNetoPromedio As Double, ByVal nMonMaxIngDescuento As Double, ByVal nCapPagoPorc As Integer, ByVal nCapPagoTotal As Double, _
                                                        ByVal dFechaVisita As Date, ByVal cEntornoFami As String, ByVal cGiroUbica As String, ByVal cExpeCrediticia As String, _
                                                        ByVal cFormaNegocio As String, ByVal cColateGarantia As String, ByVal cDestino As String, ByVal cComentario As String, _
                                                        ByVal pMatReferidos As Variant, ByVal MtrRemuneracionBrutaTotal1 As Variant, ByVal MtrDescuento1 As Variant, _
                                                        ByVal MtrRemuneracionBrutaTotal2 As Variant, ByVal MtrDescuento2 As Variant, ByVal lnColocCondi As Integer, _
                                                        ByVal prsGastoFam As Variant, ByVal pMatCuotaIfiGastoFami As Variant, ByVal pMatIfiNoSupervisadaGastoFami As Variant, Optional prsIngresos As ADODB.Recordset, Optional ByVal pMtrNegocio As Variant) As Boolean
                                                        'pMatCuotaIfiGastoFami, pMatIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agregó

   Dim EvalMes1 As Integer
   Dim EvalMes2 As Integer
   
       EvalMes1 = 1
       EvalMes2 = 2
       
   Dim oFormatosEval As COMDCredito.DCOMFormatosEval
    
    Set oFormatosEval = New COMDCredito.DCOMFormatosEval
    
    'oFormatosEval.dBeginTrans
    
    Call oFormatosEval.EliminaCredFormEvalNegocio(cCtaCod, nCodForm)
    If IsArray(pMtrNegocio) Then 'Pregunto Matriz no Sea Vacia
    If UBound(pMtrNegocio, 2) Then
        For i = 1 To UBound(pMtrNegocio, 2)
            Call oFormatosEval.GuardarFormatoConConvenio_GastosNegocio(cCtaCod, nCodForm, pMtrNegocio(0, i), pMtrNegocio(2, i))
        Next i
    End If
    End If
    
    Call oFormatosEval.ActualizarFormatoConConvenio_InfCliente(cCtaCod, nCodForm, cActividad, nAntgAnios, nAntgMes, nUltEndeSBS, dUltEndeuSBS, nTipoAportacion, nTipoInstitucion, nSectorSalud, nTipoPlanilla, cInstiConv, nMontoSol, nNumCuotas, nExposiCred, dFecFteIng, cComentario)
    
    ' Guardo mi EvalMes1 Cuando es 1
    Call oFormatosEval.ActualizarFormatoConConvenio_EvalMes1(cCtaCod, dFechaMes1, nAnoMes1, nRemBruTotalMes1, nDescMes1, nIngrNetoMes1, EvalMes1)
    'FIN Guardo mi EvalMes1 Cuando es 1
    
    ' Guardo mi EvalMes2 Cuando es 2
    Call oFormatosEval.ActualizarFormatoConConvenio_EvalMes2(cCtaCod, dFechaMes2, nAnoMes2, nRemBruTotalMes2, nDescMes2, nIngrNetoMes2, EvalMes2)
    'FIN Guardo mi EvalMes2 Cuando es 2
    
    
    Call oFormatosEval.EliminarRemBrutaTotal(cCtaCod, EvalMes1)
    ' Guardo mi EvalMes1 Tabla de Remuneracion Bruta Total Cuando es 1
    If IsArray(MtrRemuneracionBrutaTotal1) Then 'Pregunto Matriz no Sea Vacia
    If UBound(MtrRemuneracionBrutaTotal1, 2) Then
        For i = 1 To UBound(MtrRemuneracionBrutaTotal1, 2)
            Call oFormatosEval.GuardarFormatoConConvenio_RemBrutaTotal(cCtaCod, EvalMes1, MtrRemuneracionBrutaTotal1(1, i), MtrRemuneracionBrutaTotal1(2, i), MtrRemuneracionBrutaTotal1(3, i))
        Next i
    End If
    End If
    'FIN Guardo mi EvalMes1 Tabla de Remuneracion Bruta Total Cuando es 1
    
    Call oFormatosEval.EliminarRemBrutaTotal(cCtaCod, EvalMes2)
    ' Guardo mi EvalMes2 Tabla de Remuneracion Bruta Total Cuando es 2
    If IsArray(MtrRemuneracionBrutaTotal2) Then 'Pregunto Matriz no Sea Vacia
    If UBound(MtrRemuneracionBrutaTotal2, 2) Then
        For i = 1 To UBound(MtrRemuneracionBrutaTotal2, 2)
            Call oFormatosEval.GuardarFormatoConConvenio_RemBrutaTotal(cCtaCod, EvalMes2, MtrRemuneracionBrutaTotal2(1, i), MtrRemuneracionBrutaTotal2(2, i), MtrRemuneracionBrutaTotal2(3, i))
        Next i
    End If
    End If
    'FIN Guardo mi EvalMes2 Tabla de Remuneracion Bruta Total Cuando es 2
    
    Call oFormatosEval.EliminarDescuento(cCtaCod, EvalMes1)
    ' Guardo mi EvalMes1 Tabla Descuento Cuando es 1
    If IsArray(MtrDescuento1) Then 'Pregunto Matriz no Sea Vacia
    If UBound(MtrDescuento1, 2) Then
        For i = 1 To UBound(MtrDescuento1, 2)
            Call oFormatosEval.GuardarFormatoConConvenio_Descuento(cCtaCod, EvalMes1, IIf(MtrDescuento1(0, i) = "", 0, MtrDescuento1(0, i)), MtrDescuento1(1, i), IIf(MtrDescuento1(2, i) = "", 0, MtrDescuento1(2, i)))
        Next i
    End If
    End If
    'FIN Guardo mi EvalMes1 Tabla Descuento Cuando es 1
    
    Call oFormatosEval.EliminarDescuento(cCtaCod, EvalMes2)
    ' Guardo mi EvalMes2 Tabla Descuento Cuando es 2
    If IsArray(MtrDescuento2) Then 'Pregunto Matriz no Sea Vacia
    If UBound(MtrDescuento2, 2) Then
        For i = 1 To UBound(MtrDescuento2, 2)
            Call oFormatosEval.GuardarFormatoConConvenio_Descuento(cCtaCod, EvalMes2, MtrDescuento2(0, i), MtrDescuento2(1, i), MtrDescuento2(2, i))
        Next i
    End If
    'FIN Guardo mi EvalMes2 Tabla Descuento Cuando es 2
    End If
    
    Call oFormatosEval.ActualizarFormatoConConvenio_Promedio(cCtaCod, nRemBruTotalPromedio, nDescuentoPromedio, nIngNetoPromedio, nMonMaxIngDescuento, nCapPagoPorc, nCapPagoTotal)
    
    If lnColocCondi <> 4 Then
        Call oFormatosEval.EliminaCredFormEvalPropuesta(cCtaCod, nCodForm)
        Call oFormatosEval.GuardarFormatoConConvenio_PropuestaCredito(cCtaCod, nCodForm, dFechaVisita, cEntornoFami, cGiroUbica, cExpeCrediticia, cFormaNegocio, cColateGarantia, cDestino)
    End If
    
    Call oFormatosEval.EliminaCredFormEvalGrillaReferidos(cCtaCod, nCodForm)
    If IsArray(pMatReferidos) Then 'Pregunto Matriz no Sea Vacia
    If UBound(pMatReferidos) Then
        'Call oDCOMFormatosEval.EliminaCredFormEvalGrillaReferidos(psCtaCod, pnFormato)
        For i = 1 To UBound(pMatReferidos)
            Call oFormatosEval.InsertaCredFormEvalReferidos(cCtaCod, nCodForm, pMatReferidos(i, 0), pMatReferidos(i, 1), IIf(pMatReferidos(i, 2) = "", 0, pMatReferidos(i, 2)), pMatReferidos(i, 3), pMatReferidos(i, 4), IIf(pMatReferidos(i, 5) = "", 0, pMatReferidos(i, 5)))
        Next i
    End If
    End If
    
    'CTI320200110 ERS003-2020. Agregó:
    'Registro de Gastos Familiares
    If Not prsGastoFam Is Nothing Then
    Call oFormatosEval.EliminaCredFormEvalGrillaGastoFami(cCtaCod, nCodForm)
        For i = 0 To prsGastoFam.RecordCount - 1
            Call oFormatosEval.InsertaCredFormEvalGastoFami(cCtaCod, nCodForm, prsGastoFam!n, CCur(prsGastoFam!Monto))
            prsGastoFam.MoveNext
        Next i
    End If
    'Cuotas IFIs ->GastosFamiliar
    If IsArray(pMatCuotaIfiGastoFami) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatCuotaIfiGastoFami) Then
        Call oFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(cCtaCod, nCodForm, gFormatoGastosFami)
            For i = 0 To UBound(pMatCuotaIfiGastoFami) - 1
                Call oFormatosEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiGastoFami, pMatCuotaIfiGastoFami(i, 0), pMatCuotaIfiGastoFami(i, 1), CDbl(pMatCuotaIfiGastoFami(i, 2)))
            Next i
        End If
    End If

    'Cuotas de Ifis No supervisadas (Familiar)
    If IsArray(pMatIfiNoSupervisadaGastoFami) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatIfiNoSupervisadaGastoFami) Then
        Call oFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami)
            For i = 0 To UBound(pMatIfiNoSupervisadaGastoFami) - 1
                Call oFormatosEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami, pMatIfiNoSupervisadaGastoFami(i, 0), pMatIfiNoSupervisadaGastoFami(i, 1), CDbl(pMatIfiNoSupervisadaGastoFami(i, 2)))
            Next i
        End If
    End If
        
    If Not prsIngresos Is Nothing Then
        For i = 0 To prsIngresos.RecordCount - 1
            Call oFormatosEval.ActualizarCredFormEvalSinConvenioOtrosIngr(cCtaCod, nCodForm, prsIngresos!n, prsIngresos!Monto)
            prsIngresos.MoveNext
        Next i
    End If
    'Fin CTI320200110

    'oFormatosEval.dCommitTrans
    
    Set oFormatosEval = Nothing
    ActualizarConsumoConConvenio_InfCliente = True
    Exit Function
    'oFormatosEval.dRollbackTrans
    ActualizarConsumoConConvenio_InfCliente = False
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

'JOEP20160530
Public Function GuardarConsumoConConvenio_InfCliente(ByVal cCtaCod As String, ByVal nCodForm As Integer, ByVal cActividad As String, ByVal nAntgAnios As Integer, ByVal nAntgMes As Integer, ByVal nUltEndeSBS As Double, ByVal dUltEndeuSBS As Date, _
                                                    ByVal nTipoAportacion As Integer, ByVal nTipoInstitucion As Integer, ByVal nSectorSalud As Integer, ByVal nTipoPlanilla As Integer, ByVal cInstiConv As String, _
                                                    ByVal nMontoSol As Double, ByVal nNumCuotas As Integer, ByVal nExposiCred As Double, ByVal dFecFteIng As Date, _
                                                    ByVal dFechaMes1 As Integer, ByVal nAnoMes1 As Integer, ByVal nRemBruTotalMes1 As Double, ByVal nDescMes1 As Double, ByVal nIngrNetoMes1 As Double, _
                                                    ByVal dFechaMes2 As Integer, ByVal nAnoMes2 As Integer, ByVal nRemBruTotalMes2 As Double, ByVal nDescMes2 As Double, ByVal nIngrNetoMes2 As Double, _
                                                    ByVal nRemBruTotalPromedio As Double, ByVal nDescuentoPromedio As Double, ByVal nIngNetoPromedio As Double, ByVal nMonMaxIngDescuento As Double, _
                                                    ByVal nCapPagoPorc As Integer, ByVal nCapPagoTotal As Double, _
                                                    ByVal dFechaVisita As Date, ByVal cEntornoFami As String, ByVal cGiroUbica As String, ByVal cExpeCrediticia As String, _
                                                    ByVal cFormaNegocio As String, ByVal cColateGarantia As String, ByVal cDestino As String, _
                                                    ByVal cComentario As String, ByVal pMatReferidos As Variant, _
                                                    ByVal MtrRemuneracionBrutaTotal1 As Variant, ByVal MtrDescuento1 As Variant, ByVal MtrRemuneracionBrutaTotal2 As Variant, ByVal MtrDescuento2 As Variant, ByVal lnColocCondi As Integer, _
                                                    ByVal prsGastoFam As Variant, ByVal pMatCuotaIfiGastoFami As Variant, ByVal pMatIfiNoSupervisadaGastoFami As Variant, Optional prsIngresos As ADODB.Recordset, Optional ByVal pMtrNegocio As Variant) As Boolean
                                                    'pMatCuotaIfiGastoFami, pMatIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agregó
                                                           
   Dim EvalMes1 As Integer
   Dim EvalMes2 As Integer
   
       EvalMes1 = 1
       EvalMes2 = 2
       
   Dim oFormatosEval As COMDCredito.DCOMFormatosEval
    Set oFormatosEval = New COMDCredito.DCOMFormatosEval
    
    Call oFormatosEval.EliminaCredFormEvalNegocio(cCtaCod, nCodForm)
    If IsArray(pMtrNegocio) Then 'Pregunto Matriz no Sea Vacia
    If UBound(pMtrNegocio, 2) Then
        For i = 1 To UBound(pMtrNegocio, 2)
            Call oFormatosEval.GuardarFormatoConConvenio_GastosNegocio(cCtaCod, nCodForm, pMtrNegocio(0, i), pMtrNegocio(2, i))
        Next i
    End If
    End If
    
    'oFormatosEval.dBeginTrans
    Call oFormatosEval.GuardarFormatoConConvenio_InfCliente(cCtaCod, nCodForm, cActividad, nAntgAnios, nAntgMes, nUltEndeSBS, dUltEndeuSBS, nTipoAportacion, nTipoInstitucion, nSectorSalud, nTipoPlanilla, cInstiConv, nMontoSol, nNumCuotas, nExposiCred, dFecFteIng, cComentario)
    
    ' Guardo mi EvalMes1 Cuando es 1
    Call oFormatosEval.GuardarFormatoConConvenio_EvalMes1(cCtaCod, dFechaMes1, nAnoMes1, nRemBruTotalMes1, nDescMes1, nIngrNetoMes1, EvalMes1)
    'FIN Guardo mi EvalMes1 Cuando es 1
    
    ' Guardo mi EvalMes2 Cuando es 2
    Call oFormatosEval.GuardarFormatoConConvenio_EvalMes2(cCtaCod, dFechaMes2, nAnoMes2, nRemBruTotalMes2, nDescMes2, nIngrNetoMes2, EvalMes2)
    'FIN Guardo mi EvalMes2 Cuando es 2
    
    ' Guardo mi EvalMes1 Tabla de Remuneracion Bruta Total Cuando es  1
    If IsArray(MtrRemuneracionBrutaTotal1) Then 'Pregunto Matriz no Sea Vacia
    If UBound(MtrRemuneracionBrutaTotal1, 2) Then
        For i = 1 To UBound(MtrRemuneracionBrutaTotal1, 2)
            Call oFormatosEval.GuardarFormatoConConvenio_RemBrutaTotal(cCtaCod, EvalMes1, MtrRemuneracionBrutaTotal1(1, i), MtrRemuneracionBrutaTotal1(2, i), MtrRemuneracionBrutaTotal1(3, i))
        Next i
    End If
    End If
    'FIN Guardo mi EvalMes1 Tabla de Remuneracion Bruta Total Cuando es 1 -
    
    ' Guardo mi EvalMes2 Tabla de Remuneracion Bruta Total Cuando es 2
    If IsArray(MtrRemuneracionBrutaTotal2) Then 'Pregunto Matriz no Sea Vacia
    If UBound(MtrRemuneracionBrutaTotal2, 2) Then
        For i = 1 To UBound(MtrRemuneracionBrutaTotal2, 2)
            Call oFormatosEval.GuardarFormatoConConvenio_RemBrutaTotal(cCtaCod, EvalMes2, MtrRemuneracionBrutaTotal2(1, i), MtrRemuneracionBrutaTotal2(2, i), MtrRemuneracionBrutaTotal2(3, i))
        Next i
    End If
    End If
    'FIN Guardo mi EvalMes2 Tabla de Remuneracion Bruta Total Cuando es 2
    
    ' Guardo mi EvalMes1 Tabla Descuento Cuando es 1
    If IsArray(MtrDescuento1) Then 'Pregunto Matriz no Sea Vacia
    If UBound(MtrDescuento1, 2) Then
        For i = 1 To UBound(MtrDescuento1, 2)
            Call oFormatosEval.GuardarFormatoConConvenio_Descuento(cCtaCod, EvalMes1, IIf(MtrDescuento1(0, i) = "", 0, MtrDescuento1(0, i)), MtrDescuento1(1, i), IIf(MtrDescuento1(2, i) = "", 0, MtrDescuento1(2, i)))
        Next i
    End If
    End If
    'FIN Guardo mi EvalMes1 Tabla Descuento Cuando es 1
    
    ' Guardo mi EvalMes2 Tabla Descuento Cuando es 2
    If IsArray(MtrDescuento2) Then 'Pregunto Matriz no Sea Vacia
    If UBound(MtrDescuento2, 2) Then
        For i = 1 To UBound(MtrDescuento2, 2)
            Call oFormatosEval.GuardarFormatoConConvenio_Descuento(cCtaCod, EvalMes2, MtrDescuento2(0, i), MtrDescuento2(1, i), MtrDescuento2(2, i))
        Next i
    End If
    End If
    'FIN Guardo mi EvalMes2 Tabla Descuento Cuando es 2
    
    
    Call oFormatosEval.GuardarFormatoConConvenio_Promedio(cCtaCod, nRemBruTotalPromedio, nDescuentoPromedio, nIngNetoPromedio, nMonMaxIngDescuento, nCapPagoPorc, nCapPagoTotal)
    
    If lnColocCondi <> 4 Then
        Call oFormatosEval.GuardarFormatoConConvenio_PropuestaCredito(cCtaCod, nCodForm, dFechaVisita, cEntornoFami, cGiroUbica, cExpeCrediticia, cFormaNegocio, cColateGarantia, cDestino)
    End If
    
    If IsArray(pMatReferidos) Then 'Pregunto Matriz no Sea Vacia
    If UBound(pMatReferidos) Then
        'Call oDCOMFormatosEval.EliminaCredFormEvalGrillaReferidos(psCtaCod, pnFormato)
        For i = 1 To UBound(pMatReferidos)
            Call oFormatosEval.InsertaCredFormEvalReferidos(cCtaCod, nCodForm, pMatReferidos(i, 0), pMatReferidos(i, 1), IIf(pMatReferidos(i, 2) = "", 0, pMatReferidos(i, 2)), pMatReferidos(i, 3), pMatReferidos(i, 4), IIf(pMatReferidos(i, 5) = "", 0, pMatReferidos(i, 5)))
        Next i
    End If
    End If
    
    
    'CTI320200110 ERS003-2020. Agregó:
    'Registro de Gastos Familiares
    If Not prsGastoFam Is Nothing Then
    Call oFormatosEval.EliminaCredFormEvalGrillaGastoFami(cCtaCod, nCodForm)
        For i = 0 To prsGastoFam.RecordCount - 1
            Call oFormatosEval.InsertaCredFormEvalGastoFami(cCtaCod, nCodForm, prsGastoFam!n, CCur(prsGastoFam!Monto))
            prsGastoFam.MoveNext
        Next i
    End If
    'Cuotas IFIs ->GastosFamiliar
    If IsArray(pMatCuotaIfiGastoFami) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatCuotaIfiGastoFami) Then
        Call oFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(cCtaCod, nCodForm, gFormatoGastosFami)
            For i = 0 To UBound(pMatCuotaIfiGastoFami) - 1
                Call oFormatosEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiGastoFami, pMatCuotaIfiGastoFami(i, 0), pMatCuotaIfiGastoFami(i, 1), CDbl(pMatCuotaIfiGastoFami(i, 2)))
            Next i
        End If
    End If

    'Cuotas de Ifis No supervisadas (Familiar)
    If IsArray(pMatIfiNoSupervisadaGastoFami) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatIfiNoSupervisadaGastoFami) Then
        Call oFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami)
            For i = 0 To UBound(pMatIfiNoSupervisadaGastoFami) - 1
                Call oFormatosEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami, pMatIfiNoSupervisadaGastoFami(i, 0), pMatIfiNoSupervisadaGastoFami(i, 1), CDbl(pMatIfiNoSupervisadaGastoFami(i, 2)))
            Next i
        End If
    End If
    
    If Not prsIngresos Is Nothing Then
        For i = 0 To prsIngresos.RecordCount - 1
            Call oFormatosEval.InsertaCredFormEvalOtrosIngr(cCtaCod, nCodForm, prsIngresos!n, prsIngresos!Monto)
            prsIngresos.MoveNext
        Next i
    End If
    'Fin CTI320200110
    
    'oFormatosEval.dCommitTrans
    
    Set oFormatosEval = Nothing
    
    GuardarConsumoConConvenio_InfCliente = True
    Exit Function
    'oFormatosEval.dRollbackTrans
    GuardarConsumoConConvenio_InfCliente = False
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

'JOEP20160530
Public Function ActualizarConsumoSinConvenio_InfCliente(ByVal cCtaCod As String, ByVal nCodForm As Integer, ByVal cActividad As String, ByVal nAntgAnios As Integer, _
                                                     ByVal nAntgMes As Integer, ByVal nUltEndeSBS As Double, ByVal nTipoAportacion As Integer, ByVal dUltEndeuSBS As Date, _
                                                     ByVal nMontoSol As Double, ByVal nNumCuotas As Integer, ByVal nExposiCred As Double, ByVal nNunDepend As Integer, _
                                                     ByVal cPersCodEmpleado As String, ByVal dFecFteIng As Date, _
                                                     ByVal prsIngresos As ADODB.Recordset, ByVal prsEgresos As ADODB.Recordset, _
                                                     ByVal pMtrBoletaPago As Variant, ByVal pnIngNegocio As Double, ByVal pnEgrVenta As Double, ByVal pnMargBruto As Double, ByVal pnIngNeto As Double, ByVal pMtrNegocio As Variant, ByVal pMtrReciboHono As Variant, ByVal pMtrIfis As Variant, _
                                                     ByVal dFechaVisita As Date, ByVal cEntornoFami As String, ByVal cGiroUbica As String, ByVal cExpeCrediticia As String, _
                                                     ByVal cFormaNegocio As String, ByVal cColateGarantia As String, ByVal cDestino As String, _
                                                     ByVal cComentario As String, ByVal pMatReferidos As Variant, ByVal lnColocCondi As Integer, _
                                                     Optional ByVal pMatIfiNoSupervisadaGastoFami As Variant = Nothing) As Boolean

    Dim nEstadoCuota As Integer
    nEstadoCuota = 1
    
    Dim oFormatosEval As COMDCredito.DCOMFormatosEval
    
    Set oFormatosEval = New COMDCredito.DCOMFormatosEval
    
    'oFormatosEval.dBeginTrans
    
    Call oFormatosEval.ActualizarFormatoSinConvenio_InfCliente(cCtaCod, nCodForm, cActividad, nAntgAnios, nAntgMes, nUltEndeSBS, nTipoAportacion, dUltEndeuSBS, nMontoSol, nNumCuotas, nExposiCred, nNunDepend, cPersCodEmpleado, dFecFteIng, cComentario)
    
    
    If lnColocCondi <> 4 Then
        Call oFormatosEval.EliminaCredFormEvalPropuesta(cCtaCod, nCodForm)
        Call oFormatosEval.GuardarFormatoSinConvenio_PropuestaCredito(cCtaCod, nCodForm, dFechaVisita, cEntornoFami, cGiroUbica, cExpeCrediticia, cFormaNegocio, cColateGarantia, cDestino)
    End If
    
    If Not prsIngresos Is Nothing Then
        For i = 0 To prsIngresos.RecordCount - 1
            Call oFormatosEval.ActualizarCredFormEvalSinConvenioOtrosIngr(cCtaCod, nCodForm, prsIngresos!n, prsIngresos!Monto)
            prsIngresos.MoveNext
        Next i
    End If
    
    If Not prsEgresos Is Nothing Then
        For i = 0 To prsEgresos.RecordCount - 1
            Call oFormatosEval.ActualizarCredFormEvalSinConvenioGastoFami(cCtaCod, nCodForm, prsEgresos!n, prsEgresos!Monto)
            prsEgresos.MoveNext
        Next i
    End If
    
    '------------------------------------------------------------------------------------------------------------------------------------------------
    
    Call oFormatosEval.EliminaCredFormEvalBoletaPago(cCtaCod, nCodForm)
    If IsArray(pMtrBoletaPago) Then
        If UBound(pMtrBoletaPago, 2) Then
            For i = 1 To UBound(pMtrBoletaPago, 2)
                Call oFormatosEval.GuardarFormatoConConvenio_BoletaPago(cCtaCod, nCodForm, pMtrBoletaPago(1, i), pMtrBoletaPago(2, i), pMtrBoletaPago(3, i), pMtrBoletaPago(4, i)) 'Agrego pMtrBoletaPago(4, i) JOEP20180618 ERS048-2018
            Next i
        End If
    End If
    Call oFormatosEval.ActualizarFormatoSinConvenio_NegocioVentasCostos(cCtaCod, nCodForm, pnIngNegocio, pnEgrVenta, pnMargBruto, pnIngNeto)
    
    Call oFormatosEval.EliminaCredFormEvalNegocio(cCtaCod, nCodForm)
    If IsArray(pMtrNegocio) Then 'Pregunto Matriz no Sea Vacia
    If UBound(pMtrNegocio, 2) Then
        For i = 1 To UBound(pMtrNegocio, 2)
            Call oFormatosEval.GuardarFormatoConConvenio_GastosNegocio(cCtaCod, nCodForm, pMtrNegocio(0, i), pMtrNegocio(2, i))
        Next i
    End If
    End If
    
    'Cuotas Otras IFIs ->GastosNegocio / FlujoCajaMensual
    Dim oDEval  As New COMDCredito.DCOMFormatosEval
    Call oDEval.EliminaCredFormEvalGrillaOtrasIfis(cCtaCod, 7, gFormatoGastosFami)
    If IsArray(pMtrIfis) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMtrIfis, 1) Then
            For i = 0 To UBound(pMtrIfis, 1) - 1
                'Call oFormatosEval.ActualizarCredFormEvalSinConvenioCuotaIFI(cCtaCod, nCodForm, gFormatoGastosNego, gCodCuotaIfiGastoNego, pMtrIfis(i, 0), pMtrIfis(i, 1), CDbl(pMtrIfis(i, 2)), nEstadoCuota)
                Call oFormatosEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiGastoFami, pMtrIfis(i, 0), pMtrIfis(i, 1), CDbl(pMtrIfis(i, 2)))
            Next i
        End If
    End If
    
    Call oFormatosEval.EliminaCredFormEvalReciboHonorarios(cCtaCod, nCodForm)
    If IsArray(pMtrReciboHono) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMtrReciboHono, 2) Then
            For i = 1 To UBound(pMtrReciboHono, 2)
                Call oFormatosEval.GuardarFormatoConConvenio_ReciboHonorarios(cCtaCod, nCodForm, pMtrReciboHono(1, i), pMtrReciboHono(2, i), pMtrReciboHono(3, i))
            Next i
        End If
    End If
    '------------------------------------------------------------------------------------------------------------------------------------------------
    
    Call oFormatosEval.EliminaCredFormEvalGrillaReferidos(cCtaCod, nCodForm)
    If IsArray(pMatReferidos) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatReferidos, 1) Then
            For i = 1 To UBound(pMatReferidos)
                If pMatReferidos(i, 0) <> "" Then
                    Call oFormatosEval.InsertaCredFormEvalReferidos(cCtaCod, nCodForm, pMatReferidos(i, 0), pMatReferidos(i, 1), pMatReferidos(i, 2), pMatReferidos(i, 3), pMatReferidos(i, 4), pMatReferidos(i, 5))
                End If
            Next i
        End If
    End If
    
        'Cuotas de Ifis No supervisadas (Familiar)
    If IsArray(pMatIfiNoSupervisadaGastoFami) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatIfiNoSupervisadaGastoFami) Then
        Call oFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami)
            For i = 0 To UBound(pMatIfiNoSupervisadaGastoFami) - 1
                Call oFormatosEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami, pMatIfiNoSupervisadaGastoFami(i, 0), pMatIfiNoSupervisadaGastoFami(i, 1), CDbl(pMatIfiNoSupervisadaGastoFami(i, 2)))
            Next i
        End If
    End If
    'Fin CTI320200110
    
    'oFormatosEval.dCommitTrans
    
    Set oFormatosEval = Nothing
    
    ActualizarConsumoSinConvenio_InfCliente = True
    Exit Function
    'oFormatosEval.dRollbackTrans
    ActualizarConsumoSinConvenio_InfCliente = False
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

'JOEP20160530
Public Function GuardarConsumoSinConvenio_InfCliente(ByVal cCtaCod As String, ByVal nCodForm As Integer, ByVal cActividad As String, ByVal nAntgAnios As Integer, _
            ByVal nAntgMes As Integer, ByVal nUltEndeSBS As Double, ByVal nTipoAportacion As Integer, ByVal dUltEndeuSBS As Date, ByVal nMontoSol As Double, _
            ByVal nNumCuotas As Integer, ByVal nExposiCred As Double, ByVal nNunDepend As Integer, ByVal cPersCodEmpleado As String, ByVal dFecFteIng As Date, _
            ByVal prsIngresos As ADODB.Recordset, ByVal prsEgresos As ADODB.Recordset, ByVal pMtrBoletaPago As Variant, ByVal pnIngNegocio As Double, _
            ByVal pnEgrVenta As Double, ByVal pnMargBruto As Double, ByVal pnIngNeto As Double, ByVal pMtrNegocio As Variant, ByVal pMtrReciboHono As Variant, _
            ByVal pMtrIfis As Variant, ByVal dFechaVisita As Date, ByVal cEntornoFami As String, ByVal cGiroUbica As String, ByVal cExpeCrediticia As String, _
            ByVal cFormaNegocio As String, ByVal cColateGarantia As String, ByVal cDestino As String, ByVal cComentario As String, ByVal pMatReferidos As Variant, ByVal lnColocCondi As Integer, _
            Optional ByVal pMatIfiNoSupervisadaGastoFami As Variant = Nothing) As Boolean

     
    Dim oFormatosEval As New COMDCredito.DCOMFormatosEval
    'oFormatosEval.dBeginTrans
    Call oFormatosEval.GuardarFormatoSinConvenio_InfCliente(cCtaCod, nCodForm, cActividad, nAntgAnios, nAntgMes, nUltEndeSBS, nTipoAportacion, dUltEndeuSBS, nMontoSol, nNumCuotas, nExposiCred, nNunDepend, cPersCodEmpleado, dFecFteIng, cComentario)
    
    If lnColocCondi <> 4 Then
        Call oFormatosEval.GuardarFormatoSinConvenio_PropuestaCredito(cCtaCod, nCodForm, dFechaVisita, cEntornoFami, cGiroUbica, cExpeCrediticia, cFormaNegocio, cColateGarantia, cDestino)
    End If
    
    If Not prsIngresos Is Nothing Then
        For i = 0 To prsIngresos.RecordCount - 1
            Call oFormatosEval.InsertaCredFormEvalOtrosIngr(cCtaCod, nCodForm, prsIngresos!n, prsIngresos!Monto)
            prsIngresos.MoveNext
        Next i
    End If
    
    If Not prsEgresos Is Nothing Then
        For i = 0 To prsEgresos.RecordCount - 1
            Call oFormatosEval.InsertaCredFormEvalGastoFami(cCtaCod, nCodForm, prsEgresos!n, prsEgresos!Monto)
            prsEgresos.MoveNext
        Next i
    End If
    
    '------------------------------------------------------------------------------------------------------------------------------------------------
    If IsArray(pMtrBoletaPago) Then
        If UBound(pMtrBoletaPago, 2) Then
            For i = 1 To UBound(pMtrBoletaPago, 2)
                Call oFormatosEval.GuardarFormatoConConvenio_BoletaPago(cCtaCod, nCodForm, pMtrBoletaPago(1, i), pMtrBoletaPago(2, i), pMtrBoletaPago(3, i), pMtrBoletaPago(4, i)) 'Agrego pMtrBoletaPago(4, i) JOEP20180618 ERS048-2018 'ACTA Nº 112-2018
            Next i
        End If
    End If
      
    Call oFormatosEval.GuardarFormatoSinConvenio_NegocioVentasCostos(cCtaCod, nCodForm, pnIngNegocio, pnEgrVenta, pnMargBruto, pnIngNeto)
    
    If IsArray(pMtrNegocio) Then
        If UBound(pMtrNegocio, 2) Then
            For i = 1 To UBound(pMtrNegocio, 2)
                Call oFormatosEval.GuardarFormatoConConvenio_GastosNegocio(cCtaCod, nCodForm, pMtrNegocio(0, i), pMtrNegocio(2, i))
            Next i
        End If
    End If
    'Cuotas Otras IFIs ->GastosNegocio / FlujoCajaMensual
    If IsArray(pMtrIfis) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMtrIfis, 1) Then
            For i = 0 To UBound(pMtrIfis, 1) - 1
                Call oFormatosEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiGastoFami, pMtrIfis(i, 0), pMtrIfis(i, 1), CDbl(pMtrIfis(i, 2)))
            Next i
        End If
    End If
    If IsArray(pMtrReciboHono) Then
        If IsArray(pMtrReciboHono) Then
            If UBound(pMtrReciboHono, 2) Then
                For i = 1 To UBound(pMtrReciboHono, 2)
                    Call oFormatosEval.GuardarFormatoConConvenio_ReciboHonorarios(cCtaCod, nCodForm, pMtrReciboHono(1, i), pMtrReciboHono(2, i), pMtrReciboHono(3, i))
                Next i
            End If
        End If
    End If
    '------------------------------------------------------------------------------------------------------------------------------------------------
    If IsArray(pMatReferidos) Then
        If UBound(pMatReferidos) Then
            For i = 1 To UBound(pMatReferidos)
                Call oFormatosEval.InsertaCredFormEvalReferidos(cCtaCod, nCodForm, pMatReferidos(i, 0), pMatReferidos(i, 1), pMatReferidos(i, 2), pMatReferidos(i, 3), pMatReferidos(i, 4), pMatReferidos(i, 5))
            Next i
        End If
    End If
    'oFormatosEval.dCommitTrans
    
        'Cuotas de Ifis No supervisadas (Familiar)
    If IsArray(pMatIfiNoSupervisadaGastoFami) Then 'Pregunto Matriz no Sea Vacia
        If UBound(pMatIfiNoSupervisadaGastoFami) Then
        Call oFormatosEval.EliminaCredFormEvalGrillaOtrasIfis(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami)
            For i = 0 To UBound(pMatIfiNoSupervisadaGastoFami) - 1
                Call oFormatosEval.InsertaCredFormEvalCuotaIFI(cCtaCod, nCodForm, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami, pMatIfiNoSupervisadaGastoFami(i, 0), pMatIfiNoSupervisadaGastoFami(i, 1), CDbl(pMatIfiNoSupervisadaGastoFami(i, 2)))
            Next i
        End If
    End If
    'Fin CTI320200110
    Set oFormatosEval = Nothing
    GuardarConsumoSinConvenio_InfCliente = True
    Exit Function
    'oFormatosEval.dRollbackTrans
    GuardarConsumoSinConvenio_InfCliente = False
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function



'*** PEAC 20160703
Public Function RecuperaDatosInformeComercial(ByVal psCtaCod As String, ByRef prsCred As ADODB.Recordset)
    'Dim oDCred As COMDCredito.DCOMCredito
    Dim oDCred As COMDCredito.DCOMFormatosEval
    Set oDCred = New COMDCredito.DCOMFormatosEval
    Set prsCred = oDCred.RecuperaDatosInfCome(psCtaCod)
    Set oDCred = Nothing
End Function
'RECO20160719**************************************************
Public Function CargarOtrasIfis() As ADODB.Recordset
    Dim oDCred As New COMDCredito.DCOMFormatosEval
    Set CargarOtrasIfis = oDCred.CargarOtrasIfis
    Set oDCred = Nothing
End Function
Public Function ObtieneDetalleInventario(ByVal psCtaCod As String, ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnTipoPat As Integer) As ADODB.Recordset
    Dim oDCred As New COMDCredito.DCOMFormatosEval
    Set ObtieneDetalleInventario = oDCred.ObtieneDetalleInventario(psCtaCod, pnConsCod, pnConsValor, pnTipoPat)
    Set oDCred = Nothing
End Function
Public Function ObtieneDetalleActiFijo(ByVal psCtaCod As String, ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnTipoPat As Integer) As ADODB.Recordset
    Dim oDCred As New COMDCredito.DCOMFormatosEval
    Set ObtieneDetalleActiFijo = oDCred.ObtieneDetalleActiFijo(psCtaCod, pnConsCod, pnConsValor, pnTipoPat)
    Set oDCred = Nothing
End Function
Public Function RecuperaDatosCtaCobrar(ByVal psCtaCod As String, ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnTipoPat As Integer, _
                                        Optional ByVal pnTpoConsulta As Integer = 1) As ADODB.Recordset
    Dim oDCred As New COMDCredito.DCOMFormatosEval
    Set RecuperaDatosCtaCobrar = oDCred.RecuperaDatosCtaCobrar(psCtaCod, pnConsCod, pnConsValor, pnTipoPat, pnTpoConsulta)
    Set oDCred = Nothing
End Function
Public Sub ActualizaDatosLineaCredNoutilizada(ByVal psCtaCod As String, ByVal pnMontoDeuda As Currency, ByVal pnMontoCuota As Currency)
    Dim oCred As New COMDCredito.DCOMFormatosEval
    Call oCred.ActualizaDatosLineaCredNoutilizada(psCtaCod, pnMontoDeuda, pnMontoCuota)
End Sub
'RECO FIN ****************************************************
'JOEP 20160725 ***********************************************
Public Function CredFormEvalObtieneDescuentoConv(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oDCred As New COMDCredito.DCOMFormatosEval
    Set CredFormEvalObtieneDescuentoConv = oDCred.CredFormEvalObtieneDescuentoConv(psCtaCod)
    Set oDCred = Nothing
End Function
'JOEP FIN*****************************************************


'*** PEAC 20160809 -- para proximo pase
'Public Function ActualizaCredFormEvalFormato6( _
'            ByVal pvVectorDatosGrles As Variant, _
'            ByVal pnRespElimina As Integer, _
'            ByVal prsActivos As adodb.Recordset, _
'            ByVal prsPasivos As adodb.Recordset, _
'            ByVal prsEstGanPer As adodb.Recordset, _
'            ByVal prsCoeFinan As adodb.Recordset, _
'            ByVal prsFlujoCaja As adodb.Recordset, _
'            ByVal prsGastosFam As adodb.Recordset, _
'            ByVal prsOtrosIng As adodb.Recordset, _
'            ByVal prsPDT As adodb.Recordset _
'            ) As Integer
'
'Dim oCredito As New COMDCredito.DCOMFormatosEval
'Dim bTrans As Boolean
'
'On Error GoTo Error
'ActualizaCredFormEvalFormato6 = 0
''oCredito.BeginTrans
''bTrans = True
''***************************************************************************************
''Call oCredito.ActualizaCredFormEvalFormato6(psCtaCod, psMovAprueba, pnEstadoA, pnEstadoB)
'
''----------------------CABECERA
'If pnRespElimina = 1 Then
''Format(mskFecReg.Text, "yyyymmdd")
'    Call oCredito.EliminaFormatoEvaluacion(pvVectorDatosGrles(1), pvVectorDatosGrles(16))
'End If
'
''    lvVectorDatosGrles(1) = sCtaCod
''    lvVectorDatosGrles(2) = 1
''    lvVectorDatosGrles(3) = txtGiroNeg2.Text
''    lvVectorDatosGrles(4) = CInt(spnExpEmpAnio.valor)
''    lvVectorDatosGrles(5) = CInt(spnExpEmpMes.valor)
''    lvVectorDatosGrles(6) = CInt(spnTiempoLocalAnio.valor)
''    lvVectorDatosGrles(7) = CInt(spnTiempoLocalMes.valor)
''    lvVectorDatosGrles(8) = CDbl(txtUltEndeuda.Text)
''    lvVectorDatosGrles(9) = Format(txtFecUltEndeuda.Text, "yyyymmdd")
''    lvVectorDatosGrles(10) = lnCondLocal
''    lvVectorDatosGrles(11) = IIf(txtCondLocalOtros.Visible = False, "", txtCondLocalOtros.Text)
''    lvVectorDatosGrles(12) = CDbl(Me.txtExposicioCreditoito2.Text)
''    lvVectorDatosGrles(13) = Format(txtFechaEvaluacion.Text, "yyyymmdd")
''    lvVectorDatosGrles(14) = lnNumForm
''    lvVectorDatosGrles(15) = txtComentario.Text
''    lvVectorDatosGrles(16) = Format(mskFecReg, "yyyyMMdd")
'
'Call oCredito.GrabarCredFormEvalCabeceraFormato1_5( _
'    pvVectorDatosGrles(1), _
'    pvVectorDatosGrles(2), _
'    pvVectorDatosGrles(3), _
'    pvVectorDatosGrles(4), _
'    pvVectorDatosGrles(5), _
'    pvVectorDatosGrles(6), _
'    pvVectorDatosGrles(7), _
'    pvVectorDatosGrles(8), _
'    pvVectorDatosGrles(9), _
'    pvVectorDatosGrles(10), _
'    pvVectorDatosGrles(11), _
'    pvVectorDatosGrles(12), _
'    pvVectorDatosGrles(13), _
'    pvVectorDatosGrles(14), _
'    pvVectorDatosGrles(15))
'
'
'' lvPrinResumenEstFin, lvPrinResumenEstFin(i).vPP(J).nImporte
'
''--------------------- ACTIVOS
'        If UBound(lvPrinResumenEstFin) > 0 Then
'            For i = 1 To UBound(lvPrinResumenEstFin)
'                If CDbl(Me.feActivos.TextMatrix(i, 4)) <> 0 Then
'                    ''If i = 1 Or i = 10 Or i = 17 Then
'                    If i = 17 Then
'                        Call oCredito.AgregaCredFormEvalEstFinActivosPasivosFormato6(sCtaCod, 6, Format(mskFecReg, "yyyyMMdd"), chkAudit.Value, CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), CDbl(Me.feActivos.TextMatrix(i, 2)), CDbl(Me.feActivos.TextMatrix(i, 3)), CDbl(Me.feActivos.TextMatrix(i, 4)))
'                    Else
'                        Call oCredito.AgregaCredFormEvalEstFinActivosPasivosFormato6Det(sCtaCod, 6, Format(mskFecReg, "yyyyMMdd"), chkAudit.Value, CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), CDbl(Me.feActivos.TextMatrix(i, 2)), CDbl(Me.feActivos.TextMatrix(i, 3)), CDbl(Me.feActivos.TextMatrix(i, 4)))
'                    End If
'                End If
'            Next i
'        End If
'
'        '-- activos det DETALLE formato6
'        If UBound(lvPrinResumenEstFin) > 0 Then
'            For i = 1 To UBound(lvPrinResumenEstFin)
'                If lvPrinResumenEstFin(i).nDetPP > 0 Then
'                    For J = 1 To UBound(lvPrinResumenEstFin(i).vPP)
'                        If lvPrinResumenEstFin(i).vPP(J).nImporte <> 0 Then
'                            Call oCredito.AgregaCredFormEvalEstFinActivosPasivosFormato6DetDetalle( _
'                                sCtaCod, 6, Format(mskFecReg, "yyyyMMdd"), chkAudit.Value, _
'                                CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), 1, _
'                                Format(lvPrinResumenEstFin(i).vPP(J).dFecha, "yyyyMMdd"), _
'                                lvPrinResumenEstFin(i).vPP(J).cDescripcion, _
'                                lvPrinResumenEstFin(i).vPP(J).nImporte, 0, "")
'                        End If
'                    Next J
'                End If
'
'                If lvPrinResumenEstFin(i).nDetPE > 0 Then
'                    For J = 1 To UBound(lvPrinResumenEstFin(i).vPE)
'                        If lvPrinResumenEstFin(i).vPE(J).nImporte <> 0 Then
'                            Call oCredito.AgregaCredFormEvalEstFinActivosPasivosFormato6DetDetalle( _
'                                sCtaCod, 6, Format(mskFecReg, "yyyyMMdd"), chkAudit.Value, _
'                                CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), 2, _
'                                Format(lvPrinResumenEstFin(i).vPE(J).dFecha, "yyyyMMdd"), _
'                                lvPrinResumenEstFin(i).vPE(J).cDescripcion, 0, _
'                                lvPrinResumenEstFin(i).vPE(J).nImporte, "")
'                        End If
'                    Next J
'                End If
'            Next i
'        End If
'
''-------------------------- PASIVOS
'
'        If UBound(lvPrinResumenEstFinPasivos) > 0 Then
'            For i = 1 To UBound(lvPrinResumenEstFinPasivos)
'                If CDbl(Me.fePasivos.TextMatrix(i, 4)) <> 0 Then
'                    If i = 23 Or i = 24 Or i = 25 Then
'                        Call oCredito.AgregaCredFormEvalEstFinActivosPasivosFormato6(sCtaCod, lnNumForm, Format(mskFecReg, "yyyyMMdd"), _
'                            chkAudit.Value, CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), CDbl(Me.fePasivos.TextMatrix(i, 2)), _
'                            CDbl(Me.fePasivos.TextMatrix(i, 3)), CDbl(Me.fePasivos.TextMatrix(i, 4)))
'                    Else
'                        Call oCredito.AgregaCredFormEvalEstFinActivosPasivosFormato6Det(sCtaCod, lnNumForm, Format(mskFecReg, "yyyyMMdd"), _
'                            chkAudit.Value, CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), _
'                            CDbl(Me.fePasivos.TextMatrix(i, 2)), CDbl(Me.fePasivos.TextMatrix(i, 3)), CDbl(Me.fePasivos.TextMatrix(i, 4)))
'                    End If
'                End If
'            Next i
'        End If
'
'        '-- pasivos det DETALLE formato6
'        If UBound(lvPrinResumenEstFinPasivos) > 0 Then
'            For i = 1 To UBound(lvPrinResumenEstFinPasivos)
'
'                If lvPrinResumenEstFinPasivos(i).nDetPP > 0 Then
'                    For J = 1 To UBound(lvPrinResumenEstFinPasivos(i).vPP)
'                        If lvPrinResumenEstFinPasivos(i).vPP(J).nImporte <> 0 Then
'                            lcCodifi = IIf(CInt(Me.fePasivos.TextMatrix(i, 6)) = 109 Or CInt(Me.fePasivos.TextMatrix(i, 6)) = 201, Right(lvPrinResumenEstFinPasivos(i).vPP(J).cDescripcion, 8), "")
'
'                            Call oCredito.AgregaCredFormEvalEstFinActivosPasivosFormato6DetDetalle( _
'                                sCtaCod, lnNumForm, Format(mskFecReg, "yyyyMMdd"), chkAudit.Value, _
'                                CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), 1, _
'                                Format(lvPrinResumenEstFinPasivos(i).vPP(J).dFecha, "yyyyMMdd"), _
'                                lvPrinResumenEstFinPasivos(i).vPP(J).cDescripcion, _
'                                lvPrinResumenEstFinPasivos(i).vPP(J).nImporte, 0, lcCodifi)
'                        End If
'                    Next J
'                End If
'
'                If lvPrinResumenEstFinPasivos(i).nDetPE > 0 Then
'                    For J = 1 To UBound(lvPrinResumenEstFinPasivos(i).vPE)
'                        If lvPrinResumenEstFinPasivos(i).vPE(J).nImporte <> 0 Then
'                            lcCodifi = IIf(CInt(Me.fePasivos.TextMatrix(i, 6)) = 109 Or CInt(Me.fePasivos.TextMatrix(i, 6)) = 201, Right(lvPrinResumenEstFinPasivos(i).vPE(J).cDescripcion, 8), "")
'
'                            Call oCredito.AgregaCredFormEvalEstFinActivosPasivosFormato6DetDetalle( _
'                                sCtaCod, lnNumForm, Format(mskFecReg, "yyyyMMdd"), chkAudit.Value, _
'                                CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), 2, _
'                                Format(lvPrinResumenEstFinPasivos(i).vPE(J).dFecha, "yyyyMMdd"), _
'                                lvPrinResumenEstFinPasivos(i).vPE(J).cDescripcion, 0, _
'                                lvPrinResumenEstFinPasivos(i).vPE(J).nImporte, lcCodifi)
'                        End If
'                    Next J
'                End If
'            Next i
'        End If
'
''------------------- ESTADOS DE GANANCIAS Y PERDIDAS
'
'        If UBound(lvPrinResumenEstFinEstGanPer) > 0 Then
'            For i = 1 To UBound(lvPrinResumenEstFinEstGanPer)
'                If Abs(CDbl(Me.feEstaGananPerd.TextMatrix(i, 2))) <> 0 Then
'                    Call oCredito.AgregaCredFormEvalEstFinEstGanPerFormato6(sCtaCod, lnNumForm, Format(mskFecReg, "yyyyMMdd"), chkAudit.Value, CInt(Me.feEstaGananPerd.TextMatrix(i, 3)), CInt(Me.feEstaGananPerd.TextMatrix(i, 4)), CDbl(Me.feEstaGananPerd.TextMatrix(i, 2)))
'                End If
'            Next i
'        End If
'
''------------------ COEFICIENTE FINANCIERO
'    If Me.feCoeFinan.Rows - 1 > 0 Then
'        For i = 1 To Me.feCoeFinan.Rows - 1
'
'            'If Abs(CDbl(Me.feCoeFinan.TextMatrix(i, 2))) <> "0.00" Then
'            If Me.feCoeFinan.TextMatrix(i, 2) <> "0.00" Then
'                Call oCredito.AgregaCredFormEvalCoeficienteFinanFormato6(sCtaCod, lnNumForm, Format(mskFecReg, "yyyyMMdd"), CInt(Me.feCoeFinan.TextMatrix(i, 3)), _
'                                    CInt(Me.feCoeFinan.TextMatrix(i, 4)), CDbl(IIf(Right(Me.feCoeFinan.TextMatrix(i, 2), 1) = "%", Left(Me.feCoeFinan.TextMatrix(i, 2), Len(Me.feCoeFinan.TextMatrix(i, 2)) - 1), Me.feCoeFinan.TextMatrix(i, 2))))
'            End If
'        Next i
'    End If
'
''------------------ FLUJO DE CAJA
'    If Me.feFlujoCajaMensual.Rows - 1 > 0 Then
'        For i = 1 To Me.feFlujoCajaMensual.Rows - 1
'            If Abs(CCur(Me.feFlujoCajaMensual.TextMatrix(i, 2))) <> 0 Then
'                Call oCredito.InsertaCredFormFlujoCajaMensual(sCtaCod, lnNumForm, 7027, CInt(Me.feFlujoCajaMensual.TextMatrix(i, 4)), Trim(Me.feFlujoCajaMensual.TextMatrix(i, 1)), CCur(Me.feFlujoCajaMensual.TextMatrix(i, 2)))
'            End If
'        Next i
'    End If
'
''------------------ IFIS FLUJO DE CAJA
'
'    'Cuotas IFIs ->FLUJO DE CAJA
'    If IsArray(MatIfiFlujoCaja) Then 'Pregunto Matriz no Sea Vacia
'        If UBound(MatIfiFlujoCaja) > 0 Then
'        Call oCredito.EliminaCredFormEvalGrillaOtrasIfis(sCtaCod, lnNumForm, 7027)
'            For i = 0 To UBound(MatIfiFlujoCaja) - 1
'                Call oCredito.InsertaCredFormEvalCuotaIFI(sCtaCod, lnNumForm, 7027, gCodCuotaIfiFlujoCaja, MatIfiFlujoCaja(i, 0), MatIfiFlujoCaja(i, 1), CDbl(MatIfiFlujoCaja(i, 2)))
'            Next i
'        End If
'    End If
'
''------------------  GASTOS FAMILIARES
'    If Me.feGastosFamiliares.Rows - 1 > 0 Then
'        For i = 1 To Me.feGastosFamiliares.Rows - 1
'            If Abs(CDbl(Me.feGastosFamiliares.TextMatrix(i, 3))) <> 0 Then
'                Call oCredito.InsertaCredFormEvalGastoFami(sCtaCod, lnNumForm, CInt(Me.feGastosFamiliares.TextMatrix(i, 1)), CInt(Me.feGastosFamiliares.TextMatrix(i, 3)))
'            End If
'        Next i
'    End If
'
'
'
''----------------------------------IFIS GASTOS FAMILIARES
'    'Cuotas IFIs ->GastosFamiliar
'    If IsArray(MatIfiGastoFami) Then 'Pregunto Matriz no Sea Vacia
'        If UBound(MatIfiGastoFami) > 0 Then
'        Call oCredito.EliminaCredFormEvalGrillaOtrasIfis(sCtaCod, lnNumForm, gFormatoGastosFami)
'            For i = 0 To UBound(MatIfiGastoFami) - 1
'                Call oCredito.InsertaCredFormEvalCuotaIFI(sCtaCod, lnNumForm, gFormatoGastosFami, gCodCuotaIfiGastoFami, MatIfiGastoFami(i, 0), MatIfiGastoFami(i, 1), CDbl(MatIfiGastoFami(i, 2)))
'            Next i
'        End If
'    End If
'
''------------------  OTROS INGRESOS
'
'    If Me.feOtrosIngresos.Rows - 1 > 0 Then
'        For i = 1 To Me.feOtrosIngresos.Rows - 1
'            If Abs(CDbl(Me.feOtrosIngresos.TextMatrix(i, 3))) <> 0 Then
'                Call oCredito.InsertaCredFormEvalOtrosIngr(sCtaCod, lnNumForm, CInt(Me.feOtrosIngresos.TextMatrix(i, 1)), CInt(Me.feOtrosIngresos.TextMatrix(i, 3)))
'            End If
'        Next i
'    End If
'
''------------------  Declaracion PDT (Formato 4, 5, y 6)
'    If Me.feDeclaracionPDT.Rows - 1 > 0 Then
'
'        'Declaracion PDT
'        sMes1 = DevolverMes(1, nAnio3, nMes3)
'        sMes2 = DevolverMes(2, nAnio2, nMes2)
'        sMes3 = DevolverMes(3, nAnio1, nMes1)
'
'        Call oCredito.EliminaCredFormEvalPDT(sCtaCod, lnNumForm)
'        Call oCredito.EliminaCredFormEvalPDTDet(sCtaCod, lnNumForm)
'
''Call oDCOMFormatosEval.InsertaCredFormEvalPDT(psCtaCod, pnFormato, _
''            prsPdt!nConsCod, _
''            prsPdt!nConsvalor, _
''            pnMes1, pnMes2, pnMes3, pnAnio1, pnAnio2, pnAnio3)
'
'        Call oCredito.InsertaCredFormEvalPDT(sCtaCod, lnNumForm, _
'                    CInt(Me.feDeclaracionPDT.TextMatrix(1, 2)), _
'                    CInt(Me.feDeclaracionPDT.TextMatrix(1, 3)), _
'                    nMes1, _
'                    nMes2, _
'                    nMes3, _
'                    nAnio1, _
'                    nAnio2, _
'                    nAnio3)
'
'        For i = 1 To Me.feDeclaracionPDT.Rows - 1
'            If Abs(CCur(Me.feDeclaracionPDT.TextMatrix(i, 4))) > 0 Then
'                Call oCredito.InsertaCredFormEvalPDTDet(sCtaCod, lnNumForm, _
'                    CInt(Me.feDeclaracionPDT.TextMatrix(i, 2)), _
'                    CInt(Me.feDeclaracionPDT.TextMatrix(i, 3)), _
'                    CInt(Me.feDeclaracionPDT.TextMatrix(i, 4)), _
'                    CInt(Me.feDeclaracionPDT.TextMatrix(i, 5)), _
'                    CInt(Me.feDeclaracionPDT.TextMatrix(i, 6)), _
'                    CInt(Me.feDeclaracionPDT.TextMatrix(i, 7)), _
'                    CInt(Me.feDeclaracionPDT.TextMatrix(i, 8)))
'            End If
'        Next i
'    End If
'
'
''    If Not prsPdt Is Nothing Then
''        Call oDCOMFormatosEval.EliminaCredFormEvalPDT(psCtaCod, pnFormato)
''        Call oDCOMFormatosEval.EliminaCredFormEvalPDTDet(psCtaCod, pnFormato)
''        Call oDCOMFormatosEval.InsertaCredFormEvalPDT(psCtaCod, pnFormato, prsPdt!nConsCod, prsPdt!nConsvalor, pnMes1, pnMes2, pnMes3, pnAnio1, pnAnio2, pnAnio3)
''        For i = 0 To prsPdt.RecordCount - 1
''            Call oDCOMFormatosEval.InsertaCredFormEvalPDTDet(psCtaCod, pnFormato, _
''                    prsPdt!nConsCod, prsPdt!nConsvalor, _
''                    CLng(IIf(Trim(prsPdt!Mes1) = "", 0, prsPdt!Mes1)), _
''                    CLng(IIf(Trim(prsPdt!Mes2) = "", 0, prsPdt!Mes2)), _
''                    CLng(IIf(Trim(prsPdt!Mes3) = "", 0, prsPdt!Mes3)), _
''                    CLng(IIf(Trim(prsPdt!Promedio) = "", 0, prsPdt!Promedio)), _
''                    CLng(IIf(Trim(prsPdt!VentasDeclaradas) = "", 0, prsPdt!VentasDeclaradas)))
''            prsPdt.MoveNext
''        Next i
''    End If
'
''--------------------- PROPUESTA DEL CREDITO
'
'    Call oCredito.AgregaCredFormEvalPropuCredFormato6(sCtaCod, lnNumForm, Format(txtFechaVisita, "yyyyMMdd"), txtEntornoFamiliar2.Text, txtGiroUbicacion2.Text, txtExperiencia2.Text, txtFormalidadNegocio2.Text, txtColaterales2, txtDestino2.Text)
'
''--------------------- COMENTARIOS/REFERIDOS (solo referidos, comentario esta en la cabcera)
'
'    ReDim MatReferidos(feReferidos.Rows - 1, 6)
'    For i = 1 To feReferidos.Rows - 1
'        MatReferidos(i, 1) = feReferidos.TextMatrix(i, 0)
'        MatReferidos(i, 2) = feReferidos.TextMatrix(i, 1)
'        MatReferidos(i, 3) = feReferidos.TextMatrix(i, 2)
'        MatReferidos(i, 4) = feReferidos.TextMatrix(i, 3)
'        MatReferidos(i, 5) = feReferidos.TextMatrix(i, 4)
'        MatReferidos(i, 6) = feReferidos.TextMatrix(i, 5)
'     Next i
'
'    If Me.feReferidos.Rows - 1 > 1 Then
'        For i = 1 To Me.feReferidos.Rows - 1
'            Call oCredito.InsertaCredFormEvalReferidos(sCtaCod, lnNumForm, MatReferidos(i, 1), MatReferidos(i, 2), MatReferidos(i, 3), MatReferidos(i, 4), MatReferidos(i, 5), MatReferidos(i, 6))
'        Next i
'    End If
''--------------------------- recalcula ratios
'
'        Set oDCred = New COMDCredito.DCOMFormatosEval
'        Call oDCred.RecalculaIndicadoresyRatiosEvaluacion(sCtaCod, 0, Format(mskFecReg, "yyyyMMdd"))
'
''----------------------------------------------------------------------------------------------------------------
'
'
'
''***************************************************************************************
'oCredito.CommitTrans
'bTrans = False
'ActualizaCredFormEvalFormato6 = 1
'    Exit Function
'Error:
'    If bTrans Then
'        oCredito.RollbackTrans
'        Set oCredito = Nothing
'        ActualizaCredFormEvalFormato6 = 0
'    End If
'    Err.Raise Err.Number, "Error en Proceso(ActualizaCredFormEvalFormato6)", Err.Description
'End Function

'***** LUCV20170302, ANEXO 001-2017
Public Function GrabarAlertaTempranaConfig(ByVal psMovNro As String, ByVal prsDatosCapaPago As ADODB.Recordset, ByVal prsDatosRatiosOtros As ADODB.Recordset, ByVal sNombreRatioCapaPago As String, ByVal psFormulaRatioCapaPago As String) As Boolean
Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
Dim i As Integer

On Error GoTo ErrorGrabarAlertaTempranaConfig

    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    'oDCOMFormatosEval.dBeginTrans
    If Not prsDatosCapaPago Is Nothing Then
    Call oDCOMFormatosEval.EliminarAlertaTempranaConfig(prsDatosCapaPago!nCodAlerta, prsDatosCapaPago!nProducto, prsDatosCapaPago!nSubProducto, prsDatosCapaPago!nCredito, prsDatosCapaPago!nTipoCredito)
        For i = 0 To prsDatosCapaPago.RecordCount - 1
            Call oDCOMFormatosEval.GrabarAlertaTempranaConfig(prsDatosCapaPago!nCodAlerta, psMovNro, prsDatosCapaPago!Valor, prsDatosCapaPago!nUnidad, prsDatosCapaPago!nEstado, prsDatosCapaPago!nTipoCliente, prsDatosCapaPago!nProducto, prsDatosCapaPago!nSubProducto, prsDatosCapaPago!nCredito, prsDatosCapaPago!nTipoCredito, sNombreRatioCapaPago, psFormulaRatioCapaPago)
            prsDatosCapaPago.MoveNext
        Next i
    End If
    
    If Not prsDatosRatiosOtros Is Nothing Then
    Call oDCOMFormatosEval.EliminarAlertaTempranaConfig
        For i = 0 To prsDatosRatiosOtros.RecordCount - 1
            Call oDCOMFormatosEval.GrabarAlertaTempranaConfig(prsDatosRatiosOtros!nCodAlerta, psMovNro, prsDatosRatiosOtros!Valor, prsDatosRatiosOtros!nUnidad, prsDatosRatiosOtros!nEstado, , , , , , prsDatosRatiosOtros!Descripcion, prsDatosRatiosOtros!Formula)
            prsDatosRatiosOtros.MoveNext
        Next i
    End If
    'oDCOMFormatosEval.dCommitTrans
    Set oCredito = Nothing
    GrabarAlertaTempranaConfig = True
    Exit Function
ErrorGrabarAlertaTempranaConfig:
    'oDCOMFormatosEval.dRollbackTrans
    GrabarAlertaTempranaConfig = False
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function CargaDatosAlertaTempranaConfig(ByVal psProducto As String, _
                                               ByVal psSubProducto As String, _
                                               ByVal pstipocredito As String, _
                                               ByVal psSubTipoCredito As String, _
                                               ByRef prsCapaPago As ADODB.Recordset, _
                                               ByRef prsCapaPagoDet As ADODB.Recordset, _
                                               ByRef prsRatioOtro As ADODB.Recordset, _
                                               ByRef prsCodRatios As ADODB.Recordset) As Boolean

On Error GoTo ErrorCargaDatosAlertaTempranaConfig
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    CargaDatosAlertaTempranaConfig = False
    
    Set prsCapaPago = oDCOMFormatosEval.RecuperaAlertaTempranaConfigCapaPago()
    Set prsCapaPagoDet = oDCOMFormatosEval.RecuperaAlertaTempranaConfigCapaPagoDet(psProducto, psSubProducto, pstipocredito, psSubTipoCredito)
    Set prsRatioOtro = oDCOMFormatosEval.RecuperaAlertaTempranaConfigRatioOtro()
    Set prsCodRatios = oDCOMFormatosEval.ObtenerCodAlertaTempranaConfig()
    
    CargaDatosAlertaTempranaConfig = True
Exit Function
ErrorCargaDatosAlertaTempranaConfig:
    MsgBox Err.Description, vbCritical, "Error"
End Function
'***** Fin LUCV20170302 **************

'*****->LUCV20170915-Agregó, según: ERS051-2017
Public Function CargaDatosFlexEEFF(ByVal pcPersCod As String, _
                                   Optional ByRef prsFeActivosEEFF As ADODB.Recordset = "Nothing", _
                                   Optional ByRef prsFePasivosEEFF As ADODB.Recordset = "Nothing", _
                                   Optional ByRef prsFeEstadoResultadoEEFF As ADODB.Recordset = "Nothing", _
                                   Optional ByRef prsFeFlujoEfectivoEEFF As ADODB.Recordset = "Nothing", _
                                   Optional ByRef prsFeCambiosPatrimonioEEFF As ADODB.Recordset = "Nothing") As Boolean
                                                          
On Error GoTo ErrorCargaDatosFlexEEFF
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    
    Set prsFeActivosEEFF = oDCOMFormatosEval.RecuperaDatosFlexActivosEEFF()
    Set prsFePasivosEEFF = oDCOMFormatosEval.RecuperaDatosFlexPasivosEEFF()
    Set prsFeEstadoResultadoEEFF = oDCOMFormatosEval.RecuperaDatosFlexEstadoResultadoEEFF()
    Set prsFeFlujoEfectivoEEFF = oDCOMFormatosEval.RecuperaDatosFlexFlujoEfectivoEEFF()
    Set prsFeCambiosPatrimonioEEFF = oDCOMFormatosEval.RecuperaDatosFlexCambiosPatrimonioEEFF()
    
CargaDatosFlexEEFF = True
Exit Function
ErrorCargaDatosFlexEEFF:
    MsgBox Err.Description, vbCritical, "Error"
End Function

Public Function GrabarPersEstadosFinancierosDetNew(ByVal pnCodEF As Integer, _
                                                    ByVal prsEstadoResultado As ADODB.Recordset, _
                                                    ByVal prsFlujoEfectivo As ADODB.Recordset, _
                                                    ByVal prsCambiosPatrimonio As ADODB.Recordset) As Boolean
Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
Dim i As Integer
On Error GoTo ErrorGrabarPersEstadosFinancierosDetNew
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
                                                                                                   
    'Estados de resultados
    If Not prsEstadoResultado Is Nothing Then
    Call oDCOMFormatosEval.EliminaPersEstadoFinancieroDetNew(pnCodEF, prsEstadoResultado!nConsCod)
        For i = 0 To prsEstadoResultado.RecordCount - 1
            Call oDCOMFormatosEval.InsertaPersEstadoFinancieroDetNew(pnCodEF, (i + 1), prsEstadoResultado!nEstFinanTipo, prsEstadoResultado!nConsCod, prsEstadoResultado!nConsValor, 0, prsEstadoResultado!Totales, prsEstadoResultado!Porcentaje)
            prsEstadoResultado.MoveNext
        Next i
    End If
    Set prsEstadoResultado = Nothing
    'Flujo de Efectivo
    If Not prsFlujoEfectivo Is Nothing Then
    Call oDCOMFormatosEval.EliminaPersEstadoFinancieroDetNew(pnCodEF, prsFlujoEfectivo!nConsCod)
        For i = 0 To prsFlujoEfectivo.RecordCount - 1
            Call oDCOMFormatosEval.InsertaPersEstadoFinancieroDetNew(pnCodEF, (i + 1), prsFlujoEfectivo!nEstFinanTipo, prsFlujoEfectivo!nConsCod, prsFlujoEfectivo!nConsValor, 0, prsFlujoEfectivo!Totales)
            prsFlujoEfectivo.MoveNext
        Next i
    End If
    Set prsFlujoEfecitvo = Nothing
    'Cambios en el Patrimonio
    If Not prsCambiosPatrimonio Is Nothing Then
    Call oDCOMFormatosEval.EliminaPersEstadoFinancieroDetNew(pnCodEF, prsCambiosPatrimonio!nConsCod)
        For i = 0 To prsCambiosPatrimonio.RecordCount - 1
            Call oDCOMFormatosEval.InsertaPersEstadoFinancieroDetNew(pnCodEF, (i + 1), prsCambiosPatrimonio!nEstFinanTipo, prsCambiosPatrimonio!nConsCod, prsCambiosPatrimonio!nConsValor, 0, prsCambiosPatrimonio!Total, , prsCambiosPatrimonio!CapitalSocial, _
                                                                     prsCambiosPatrimonio!CapitalAdicional, prsCambiosPatrimonio!ReservaLegal, prsCambiosPatrimonio!ResulAcumulado, prsCambiosPatrimonio!UtilidadEjercicio)
            prsCambiosPatrimonio.MoveNext
        Next i
    End If
    Set prsCambiosPatrimonio = Nothing
    GrabarPersEstadosFinancierosDetNew = True
    Exit Function
ErrorGrabarPersEstadosFinancierosDetNew:
    'oCredito.dRollbackTrans
    GrabarPersEstadosFinancierosDetNew = False
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function CargaDatosPersEstadosFinancierosDetNew(ByVal pnCodEF As Integer, _
                                                        ByRef prsFeActivosEEFF As ADODB.Recordset, _
                                                        ByRef prsFePasivosEEFF As ADODB.Recordset, _
                                                        ByRef prsFeEstadoResultadoEEFF As ADODB.Recordset, _
                                                        ByRef prsFeFlujoEfectivoEEFF As ADODB.Recordset, _
                                                        ByRef prsFeCambiosPatrimonioEEFF As ADODB.Recordset, _
                                                        ByRef prsDatosPersEstadosFinancieros As ADODB.Recordset) As Boolean
                                                                                        
On Error GoTo ErrorCargaDatosPersEstadosFinancierosDetNew
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim prsColocacCred As ADODB.Recordset
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    CargaDatosPersEstadosFinancierosDetNew = False
    
    Set prsFeActivosEEFF = oDCOMFormatosEval.RecuperaDatosActivosEEFF(pnCodEF)
    Set prsFePasivosEEFF = oDCOMFormatosEval.RecuperaDatosPasivosEEFF(pnCodEF)
    Set prsFeEstadoResultadoEEFF = oDCOMFormatosEval.RecuperaDatosEstadoResultadoEEFF(pnCodEF)
    Set prsFeFlujoEfectivoEEFF = oDCOMFormatosEval.RecuperaDatosFlujoEfectivoEEFF(pnCodEF)
    Set prsFeCambiosPatrimonioEEFF = oDCOMFormatosEval.RecuperaDatosCambiosPatrimonioEEFF(pnCodEF)
    Set prsDatosPersEstadosFinancieros = oDCOMFormatosEval.RecuperaDatosPersEstadosFinancieros(pnCodEF)
    
    CargaDatosPersEstadosFinancierosDetNew = True
Exit Function
ErrorCargaDatosPersEstadosFinancierosDetNew:
    MsgBox Err.Description, vbCritical, "Error"
    CargaDatosPersEstadosFinancierosDetNew = False
End Function

Public Function RecuperaDatosSubItemDetalle(ByVal pnCodEF As Integer, ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, Optional ByVal pnConsValorGrupo As Integer = -1) As ADODB.Recordset
    Dim oDCred As New COMDCredito.DCOMFormatosEval
    Set RecuperaDatosSubItemDetalle = oDCred.RecuperaDatosDetalleSubItem(pnCodEF, pnConsCod, pnConsValor, pnConsValorGrupo)
    Set oDCred = Nothing
End Function

'->********** LUCV20171015 ->**********
Public Function CargaDatosComparacionAnual(ByVal psPersCod As String, _
                                        ByVal psFechaInicial As String, _
                                        ByVal psFechaFinal As String, _
                                        ByRef prsBalanceActivo As ADODB.Recordset, _
                                        Optional ByRef prsBalancePasivoPat As ADODB.Recordset = Nothing, _
                                        Optional ByRef prsEERR As ADODB.Recordset = Nothing) As Boolean
                                        
On Error GoTo ErrorCargaDatosComparacionAnual
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    CargaDatosComparacionAnual = False
    
    Set prsBalanceActivo = oDCOMFormatosEval.RecuperaDatosComparativoActivo(psPersCod, psFechaInicial, psFechaFinal)
    Set prsBalancePasivoPat = oDCOMFormatosEval.RecuperaDatosComparativoPasivoPat(psPersCod, psFechaInicial, psFechaFinal)
    Set prsEERR = oDCOMFormatosEval.RecuperaDatosComparativoEERR(psPersCod, psFechaInicial, psFechaFinal)
    
    CargaDatosComparacionAnual = True
Exit Function
ErrorCargaDatosComparacionAnual:
    CargaDatosComparacionAnual = False
    MsgBox Err.Description, vbCritical, "Error"
End Function
Public Function CargaDatosFlujoCajaRpt(ByVal psCtaCod As String, _
                                        ByRef prsFlujoCajaRptObtieneDatosCabecera As ADODB.Recordset, _
                                        ByRef prsFlujoCajaRptObtieneDatosConceptos As ADODB.Recordset, _
                                        ByRef prsFlujoCajaRptObtieneDatosCuotas As ADODB.Recordset, _
                                        ByRef prsFlujoCajaRptObtieneDatosParametros As ADODB.Recordset) As Boolean
                                        
On Error GoTo ErrorCargaDatosFlujoCajaRpt
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    CargaDatosFlujoCajaRpt = False
    'Carga para la carga del archivo a generar
    Set prsFlujoCajaRptObtieneDatosCabecera = oDCOMFormatosEval.RecuperaFlujoCajaRptObtieneDatosCabecera(psCtaCod)
    Set prsFlujoCajaRptObtieneDatosConceptos = oDCOMFormatosEval.RecuperaFlujoCajaRptObtieneDatosConceptos(psCtaCod)
    Set prsFlujoCajaRptObtieneDatosCuotas = oDCOMFormatosEval.RecuperaFlujoCajaRptObtieneDatosCuotas(psCtaCod)
    Set prsFlujoCajaRptObtieneDatosParametros = oDCOMFormatosEval.RecuperaFlujoCajaRptObtieneDatosParametros(psCtaCod)
    CargaDatosFlujoCajaRpt = True
Exit Function
ErrorCargaDatosFlujoCajaRpt:
    CargaDatosFlujoCajaRpt = False
    MsgBox Err.Description, vbCritical, "Error"
End Function
'<-********** Fin LUCV20171015 <-**********
'*****<- Fin LUCV20170915

