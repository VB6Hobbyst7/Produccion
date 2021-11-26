VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NPigRemate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub nPigRegistraSobrante(ByVal rsSob As Recordset, ByVal pnRemate As Integer, ByVal pdFecha As String, _
            ByVal psCodAge As String, ByVal psCodUser As String)
Dim oPigGraba As DPigRemate
Dim lsCtaAhorro As String
Dim lnSobrante As Double
Dim lnTasaInt As Double
Dim lsMovRS As String
Dim lnMovRS As Long
Dim lsMovBloq As String
Dim rs As Recordset
Dim MatSobranteCtaExist As Variant
Dim MatSobranteCtaNueva As Variant
Dim lnContCN As Integer
Dim lnContCE As Integer
Dim I As Integer

    ReDim MatSobranteCtaNueva(rsSob.RecordCount, 5)
    ReDim MatSobranteCtaExist(rsSob.RecordCount, 5)
    
    If rsSob.RecordCount > 0 Then
        
        lnContCN = 0
        lnContCE = 0
        Set oPigGraba = New DPigRemate
        
        Do While Not rsSob.EOF
        
            lsCtaAhorro = oPigGraba.ValidaSiCtaAhorro(rsSob!cPersCod)
            If lsCtaAhorro = "" Then
                lnTasaInt = oPigGraba.GetCapTasaInteres(gCapAhorros, gMonedaNacional, gCapTasaNormal, rsSob!nSobrante, psCodAge)
                lnTasaInt = Format$(ConvierteTNAaTEA(lnTasaInt), "#,##0.00")
                
            Else
                lnTasaInt = 0
            End If
            
            If lsCtaAhorro = "" Then
                MatSobranteCtaNueva(lnContCN, 0) = rsSob!cCtaCod
                MatSobranteCtaNueva(lnContCN, 1) = rsSob!cPersCod
                MatSobranteCtaNueva(lnContCN, 2) = rsSob!nSobrante
                MatSobranteCtaNueva(lnContCN, 3) = lsCtaAhorro
                MatSobranteCtaNueva(lnContCN, 4) = lnTasaInt
                MatSobranteCtaNueva(lnContCN, 5) = 1
                lnContCN = lnContCN + 1
            Else
                MatSobranteCtaExist(lnContCE, 0) = rsSob!cCtaCod
                MatSobranteCtaExist(lnContCE, 1) = rsSob!cPersCod
                MatSobranteCtaExist(lnContCE, 2) = rsSob!nSobrante
                MatSobranteCtaExist(lnContCE, 3) = lsCtaAhorro
                MatSobranteCtaExist(lnContCE, 4) = lnTasaInt
                MatSobranteCtaExist(lnContCE, 5) = 0
                lnContCE = lnContCE + 1
            End If
            
            rsSob.MoveNext
        Loop
        
        Set oPigGraba = Nothing
        
    End If
    
    Set oPigGraba = New DPigRemate
    
    '========= APERTURA POR SOBRANTE DE REMATES ===========
    lsMovRS = oPigGraba.GeneraMovNro(pdFecha, Right(psCodAge, 2), psCodUser)
    oPigGraba.dInsertMov lsMovRS, gAhoApeSobRemate, "Apertura por Sobrante de Remate", gMovEstContabMovContable, gMovFlagVigente, False
    lnMovRS = oPigGraba.dGetnMovNro(lsMovRS)
    
    For I = 0 To UBound(MatSobranteCtaNueva) - 1
        If MatSobranteCtaNueva(I, 1) <> "" Then
            lsCtaAhorro = oPigGraba.ValidaSiCtaAhorro(MatSobranteCtaNueva(I, 1))
            If lsCtaAhorro = "" Then
                oPigGraba.dBeginTrans
                lsCtaAhorro = oPigGraba.CapAperturaCuenta(gCapAhorros, gMonedaNacional, MatSobranteCtaNueva(I, 1), psCodAge, _
                            MatSobranteCtaNueva(I, 4), MatSobranteCtaNueva(I, 2), pdFecha, 1, gPersonaNat, _
                              "Apertura por Sobrante de Remate", gPrdCtaTpoIndiv, lsMovRS, lnMovRS, gCapTasaNormal, gAhoApeSobRemate)
                    
                MatSobranteCtaNueva(I, 3) = lsCtaAhorro
                
                Call oPigGraba.dInsertaSobrante(MatSobranteCtaNueva(I, 0), pnRemate, MatSobranteCtaNueva(I, 3), MatSobranteCtaNueva(I, 2), _
                        lsMovRS, 0)
                oPigGraba.dCommitTrans
            Else
                MatSobranteCtaNueva(I, 5) = 0
                MatSobranteCtaExist(lnContCE, 0) = MatSobranteCtaNueva(I, 0)
                MatSobranteCtaExist(lnContCE, 1) = MatSobranteCtaNueva(I, 1)
                MatSobranteCtaExist(lnContCE, 2) = MatSobranteCtaNueva(I, 2)
                MatSobranteCtaExist(lnContCE, 3) = lsCtaAhorro
                MatSobranteCtaExist(lnContCE, 4) = MatSobranteCtaNueva(I, 4)
                MatSobranteCtaExist(lnContCE, 5) = 0
                lnContCE = lnContCE + 1
            End If
        End If
    Next I
    
    '============ BLOQUEO DE LA CUENTA NUEVA ============
    oPigGraba.dBeginTrans
    For I = 0 To UBound(MatSobranteCtaNueva) - 1
        If MatSobranteCtaNueva(I, 5) = 1 Then
            oPigGraba.NuevoBloqueoTotal MatSobranteCtaNueva(I, 3), gCapMotBlqTotSobRemate, "Bloqueo de Cta por Sobrante de Remate", lsMovRS
            oPigGraba.AgregaCaptacEstado MatSobranteCtaNueva(I, 3), gCapEstBloqTotal, lsMovRS
            oPigGraba.ActualizaEstadoCuenta MatSobranteCtaNueva(I, 3), gCapEstBloqTotal
        End If
    Next I
    oPigGraba.dCommitTrans
    
    '============= ABONO EN LA CUENTA EXISTENTE ==============
    lsMovRS = oPigGraba.GeneraMovNro(pdFecha, Right(psCodAge, 2), psCodUser)
    oPigGraba.dInsertMov lsMovRS, gAhoDepSobRemate, "Deposito por Sobrante de Remate", gMovEstContabMovContable, gMovFlagVigente, False
    lnMovRS = oPigGraba.dGetnMovNro(lsMovRS)
    
    For I = 0 To UBound(MatSobranteCtaExist) - 1
        If MatSobranteCtaExist(I, 3) <> "" Then
            oPigGraba.dBeginTrans
            Call oPigGraba.dInsertaSobrante(MatSobranteCtaExist(I, 0), pnRemate, MatSobranteCtaExist(I, 3), MatSobranteCtaExist(I, 2), lsMovRS, 0)
            Call oPigGraba.CapAbonoCuentaAho(MatSobranteCtaExist(I, 3), MatSobranteCtaExist(I, 2), gAhoDepSobRemate, lsMovRS, lnMovRS, "Abono por Sobrante de Remate")
            oPigGraba.dCommitTrans
        End If
    Next I
        
    oPigGraba.dUpdateColocPigRemate pnRemate
    
    Set oPigGraba = Nothing
    
End Sub

Public Sub nPigPagoSobrante(ByVal rs As Recordset, ByVal psctaAho As String, ByVal psMovNro As String, ByVal pnMonto As Currency, _
                            ByVal pdFecha As String, ByVal psCodAge As String, ByVal psCodUser As String)
                            
Dim oPigGraba As DPigRemate
Dim oContFunc As NContFunciones
Dim lsMovNro As String
Dim lsMovBloqueo As String
Dim lnMovNro As Long
Dim lnMovNroAho As Long

    Set oPigGraba = New DPigRemate
    lsMovBloqueo = oPigGraba.GetMovBloqueo(psctaAho)  'NUMERO DE MOVIMIENTO DEL BLOQUEO
    Sleep 2
    lsMovNro = oPigGraba.GeneraMovNro(pdFecha, psCodAge, psCodUser)
    oPigGraba.dBeginTrans
    
    oPigGraba.dInsertMov psMovNro, gPigOpePagoSobrantes, "Pago de Sobrante por Remate", gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oPigGraba.dGetnMovNro(psMovNro)
    
    '=========== ACTUALIZA ESTADO DEL SOBRANTE ============
    Do While Not rs.EOF
        oPigGraba.dUpdateSobrante rs!Contrato, 1
        oPigGraba.dInsertMovCol lnMovNro, gPigOpePagoSobrantes, rs!Contrato, 0, rs!Sobrante, 0, "", 0, gPigEstRematFact, 0
    rs.MoveNext
    Loop
    
    '=========== DESBLOQUEO DE LA CUENTA ==================
    oPigGraba.ActualizaBloqueoTot psctaAho, "Desbloqueo de Cuenta", psMovNro, gCapMotBlqTotSobRemate, lsMovBloqueo
    oPigGraba.ActualizaEstadoCuenta psctaAho, gCapEstActiva
    
    '=========== CANCELACION DE LA CUENTA =================
    lnMovNroAho = oPigGraba.CapCancelaCuentaAho(psctaAho, lsMovNro, "Cancelación por Sobrante Remate", gCapEstCancelada)
    
    oPigGraba.dInsertMovRef lnMovNroAho, lnMovNro
    
    oPigGraba.dCommitTrans
    Set oPigGraba = Nothing

End Sub

Public Sub nPigVentaRemate(ByVal psCtaCod As String, ByVal pnValorVenta As Currency, ByVal pnTipoVenta As Integer, _
                           ByVal psFechaGraba As String, ByVal prJoyas As Recordset, ByVal pnTipoProceso As Integer, _
                           ByVal pnEstado As Integer, Optional ByVal psComprador As String)
                           
Dim oGraba As DPigActualizaBD
Dim lnValorVenta As Currency
Dim Cliente As String
Dim lnPiezas As Integer
Dim I As Integer
Dim lnValorVta As Currency

    Set oGraba = New DPigActualizaBD
    oGraba.dBeginTrans

    Call oGraba.dUpdateProducto(psCtaCod, , , gPigEstRematPFact, psFechaGraba)
    
    If pnEstado = gPigEstRemat Then
        Call oGraba.dInsertColocacEstado(psCtaCod, psFechaGraba, gPigEstRematPFact, 1, 0, "Rematado Pendiente de Facturacion", gColocCalendCodPFCF, 0, 0, 0, 0, 0, 0, 0, 0)
        Call oGraba.dUpdateColocPigno(psCtaCod, , gPigEstRematPFact)
    ElseIf pnEstado = gPigEstRematPRes Then
        Call oGraba.dInsertColocacEstado(psCtaCod, psFechaGraba, gPigEstPResRematPFact, 1, 0, "Pend Rescate Rematado Pendiente de Facturacion", gColocCalendCodPFCF, 0, 0, 0, 0, 0, 0, 0, 0)
        Call oGraba.dUpdateColocPigno(psCtaCod, , gPigEstPResRematPFact)
    End If
    
    I = 0
    If pnTipoVenta = 1 Then 'EN CASO DE QUE LA VENTA SEA POR LOTE
        
        prJoyas.MoveFirst
        Do While Not prJoyas.EOF
            I = I + 1
            lnPiezas = prJoyas.RecordCount
            If I = lnPiezas Then
                lnValorVta = pnValorVenta
            Else
                lnValorVta = prJoyas!ValBase
                pnValorVenta = pnValorVenta - prJoyas!ValBase
            End If
            If lnValorVta > 0 Then
                Call oGraba.dUpdateSituacionPieza(psCtaCod, prJoyas!Pieza, gPigSituacionPendFacturar, pnTipoProceso, psComprador, lnValorVta)
            End If
            prJoyas.MoveNext
        Loop
        
    ElseIf pnTipoVenta = 2 Then  'EN CASO DE QUE LA VENTA SEA POR PIEZA
        
        prJoyas.MoveFirst
        Do While Not prJoyas.EOF
            Cliente = prJoyas!Cliente
            lnValorVta = prJoyas!ValVenta
            If lnValorVta > 0 Then
                Call oGraba.dUpdateSituacionPieza(psCtaCod, prJoyas!Pieza, gPigSituacionPendFacturar, pnTipoProceso, Cliente, lnValorVta)
            End If
            prJoyas.MoveNext
        Loop

    End If

    Set prJoyas = Nothing
    oGraba.dCommitTrans
    Set oGraba = Nothing

End Sub

Public Sub nPigFacturaVentaRemate(ByVal psMovNro As String, ByVal psFechaHora As String, ByVal psOpeCod As String, _
                                  ByVal prPiezas As Recordset, ByVal psNumDoc As String, ByVal pnTVenta As Currency, _
                                  ByVal psCodPers As String, Optional ByVal pnRemate As Integer)

Dim oGraba As DPigActualizaBD
Dim lnMovNro As Long
Dim lnMontoTran As Currency
Dim lnMontoPieza As Currency
Dim rsDet As Recordset

Dim lnCapital As Currency, lnComision As Currency
Dim lnInteresComp As Currency, lnImpuesto As Currency
Dim lnMontoEntregar As Currency
Dim lnConcepto As Integer
Dim I As Integer
Dim lsCodCta As String

    I = 0: lsCodCta = ""
    Set oGraba = New DPigActualizaBD
    oGraba.dBeginTrans
    
    '************************ INSERTA EN EL MOV
    oGraba.dInsertMov psMovNro, psOpeCod, "Facturacion Venta en Remate", gMovEstContabMovContable, gMovFlagVigente
    lnMovNro = oGraba.dGetnMovNro(psMovNro)
    
    '*********************** INSERTAR EN MOVDOCPIG
    oGraba.dInsertMovDocPig gPigTipoPoliza, psNumDoc, lnMovNro, pnTVenta, psCodPers, gPigTipoVentaATerceros, 1, psMovNro
        
    prPiezas.MoveFirst
    Do While Not prPiezas.EOF
        I = I + 1
        '************************ ACTUALIZAR LA INFORMACION CORRESPONDIENTE A COLOCACIONES
        Set rsDet = oGraba.dGetColocPigProcesoDet(prPiezas!Contrato, prPiezas!Pieza, pnRemate)

        oGraba.dUpdateProducto prPiezas!Contrato, , , gPigEstRematFact, psFechaHora
        If lsCodCta <> prPiezas!Contrato Then
            lsCodCta = prPiezas!Contrato
            oGraba.dInsertColocacEstado prPiezas!Contrato, psFechaHora, gPigEstRematFact, 0, 0, "Rematado Facturado", gColocCalendCodFFCF, 0, 0, 0, 0, 0, 0, 0, 0
        End If
        oGraba.dUpdateColocPigno prPiezas!Contrato, , gPigEstRematFact
        oGraba.dUpdateColocPigProceso prPiezas!NumRemate, prPiezas!TipoProceso, prPiezas!Contrato, prPiezas!Pieza, psMovNro, , , gPigSituacionFacturado
        
        '***************** INSERTA EN MOVDOCPIGDET -  VALOR DE LA PIEZA
        oGraba.dInsertMovDocPigDet lnMovNro, gPigTipoPoliza, psNumDoc, I, prPiezas!Contrato, prPiezas!Pieza, gColPigConceptoCodPrecioVenta, prPiezas!Importe
        '***************** INSERTA EN MOVDOCPIGDET -  COMISION DEL MARTILLERO
        oGraba.dInsertMovDocPigDet lnMovNro, gPigTipoPoliza, psNumDoc, I, prPiezas!Contrato, prPiezas!Pieza, gColPigConceptoCodComisionMartillero, prPiezas!ComMart
        oGraba.dInsertMovDocPigDet lnMovNro, gPigTipoPoliza, psNumDoc, I, prPiezas!Contrato, prPiezas!Pieza, gColPigConceptoCodIgv, prPiezas!IGV
        
        '***************** INSERTA EN MOVCOL
        lnMontoPieza = CCur(prPiezas!Importe)
        lnMontoTran = CCur(prPiezas!Importe) + CCur(prPiezas!ComMart) + CCur(prPiezas!IGV)
        oGraba.dInsertMovCol lnMovNro, psOpeCod, prPiezas!Contrato, prPiezas!Pieza, lnMontoTran, 0, "", 0, gPigEstRematPFact, 0
      
        '***************** INSERTA EN MOVCOLDET (Un Reg por cada concepto)
        Do While Not rsDet.EOF
            oGraba.dInsertMovColDet lnMovNro, psOpeCod, prPiezas!Contrato, prPiezas!Pieza, rsDet!nCodConcepto, prPiezas!Pieza, rsDet!nMonto
            lnMontoPieza = lnMontoPieza - CCur(rsDet!nMonto)
            rsDet.MoveNext
        Loop
        Set rsDet = Nothing
        
        '****************** SOBRANTE DE DINERO EN MOVCOLDET
        If lnMontoPieza > 0 Then
            oGraba.dInsertMovColDet lnMovNro, psOpeCod, prPiezas!Contrato, prPiezas!Pieza, gColPigConceptoCodPendCompensacion, prPiezas!Pieza, lnMontoPieza
        End If
        
        '****************** COMISION DEL MARTILLERO EN MOVCOLDET
        oGraba.dInsertMovColDet lnMovNro, psOpeCod, prPiezas!Contrato, prPiezas!Pieza, gColPigConceptoCodComisionMartillero, prPiezas!Pieza, prPiezas!ComMart
        oGraba.dInsertMovColDet lnMovNro, psOpeCod, prPiezas!Contrato, prPiezas!Pieza, gColPigConceptoCodIgv, prPiezas!Pieza, prPiezas!IGV
        
        '********** LIBERACION DE LA GARANTIA (si el contrato fue un contrato activo, se libera la garantia de la pieza vendida)
        If prPiezas!Estado = gPigEstRematPFact Then
            oGraba.dInsertMovColDet lnMovNro, psOpeCod, prPiezas!Contrato, prPiezas!Pieza, gColPigConceptoCodGarantia, prPiezas!Pieza, prPiezas!Tasacion
        End If
        '********* DESCARGO DE LA PIEZA EN CUSTODIA
        If prPiezas!Estado = gPigEstRematPFact Then
            oGraba.dInsertMovColDet lnMovNro, psOpeCod, prPiezas!Contrato, prPiezas!Pieza, gColPigConceptoCodGarantPendRescAge, prPiezas!Pieza, prPiezas!Tasacion
        ElseIf prPiezas!Estado = gPigEstPResRematPFact Then
            oGraba.dInsertMovColDet lnMovNro, psOpeCod, prPiezas!Contrato, prPiezas!Pieza, gColPigConceptoCodGarantPendRescAge, prPiezas!Pieza, prPiezas!Tasacion
        End If
        
        prPiezas.MoveNext
        
    Loop
    
    Set prPiezas = Nothing
    oGraba.dCommitTrans
    Set oGraba = Nothing
    Exit Sub
    
ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Public Function ExtornoPagoSobrante(ByVal pnMovNro As Long, ByVal psMovNro As String, ByVal pdFecha As String, _
                               ByVal psCodAge As String, psCodUser As String) As Long
Dim oPigGraba As DPigRemate
Dim sCtaAho As String
Dim lnMonto As Currency
Dim lnMovNro As Long
Dim lnMovAho As Long
Dim lnMovAhoExt As Long
Dim rs As Recordset
Dim rsAho As Recordset
Dim lsMovNro As String
Dim lsMovBloqueo As String

    Set oPigGraba = New DPigRemate
    
    '=========== Obtener los datos del Extorno ==============
    lnMovAho = oPigGraba.GetMovRef(pnMovNro)
    Set rsAho = oPigGraba.GetMovCap(lnMovAho)
    
    If Not (rsAho.EOF And rsAho.BOF) Then
        sCtaAho = rsAho!cCtaCod
        lnMonto = rsAho!nMonto
    End If
    Set rsAho = Nothing
    
    Set rs = oPigGraba.GetMovCol(pnMovNro)
    
    lsMovBloqueo = oPigGraba.GetMovBloqueo(sCtaAho)  'NUMERO DE MOVIMIENTO DEL BLOQUEO
    Sleep 2
    lsMovNro = oPigGraba.GeneraMovNro(pdFecha, psCodAge, psCodUser)   ' Mov para el extorno de Cancelacion
    
    oPigGraba.dBeginTrans
    
    '=========== ACTUALIZA MOV PAGO DE SOBRANTE
    oPigGraba.dInsertMov psMovNro, gPigOpeExtPagoSobrante, "Extorno de Pago Sobrante por Remate", gMovEstContabNoContable, gMovFlagDeExtorno
    oPigGraba.dUpdateMov pnMovNro, gMovFlagExtornado
    
    lnMovNro = oPigGraba.dGetnMovNro(psMovNro)
    
    '=========== ACTUALIZA ESTADO DEL SOBRANTE ============
    Do While Not rs.EOF
        oPigGraba.dUpdateSobrante rs!cCtaCod, 0
        rs.MoveNext
    Loop

    '======= EXTORNO DE LA CANCELACION
    lnMovAhoExt = oPigGraba.CapExtornoCancelacion(lnMovAho, gAhoExtCancSobRemate, sCtaAho, lsMovNro, lnMonto)

    oPigGraba.dInsertMovRef lnMovAhoExt, lnMovNro
    
    ExtornoPagoSobrante = lnMovNro
    
    oPigGraba.dCommitTrans
    Set oPigGraba = Nothing

End Function
