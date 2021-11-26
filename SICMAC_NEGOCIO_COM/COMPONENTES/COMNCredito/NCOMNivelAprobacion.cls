VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMNivelAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim oNivApr As COMDCredito.DCOMNivelAprobacion

'MAVM 20110429 ***
Public Sub InsertarNivApr(ByVal cNivAprCod As String, ByVal nNivelPor As Integer, ByVal cNivelCod As String, ByVal nMontoMin As Currency, ByVal nMontoMax As Currency, ByVal cRiesgoCod As String, ByVal FlexCargo As ADODB.Recordset)
    Dim i, J As Integer
    Dim objNivAprob As COMDCredito.DCOMNivelAprobacion
    Set objNivAprob = New COMDCredito.DCOMNivelAprobacion
    
On Error GoTo ControlError
    objNivAprob.InsertarNivApr cNivAprCod, nNivelPor, cNivelCod, nMontoMin, nMontoMax, cRiesgoCod
    
    If Not FlexCargo Is Nothing Then
        Do While Not FlexCargo.EOF
            objNivAprob.InsertarNivAprCargo cNivAprCod, FlexCargo.Fields(0)
            FlexCargo.MoveNext
        Loop
    End If
    Exit Sub
ControlError:
    Set objNivAprob = Nothing
End Sub
'***

'JUEZ 20121128 ********************************************************************
Public Sub dInsertaGruposApr(ByVal psGrupoDesc As String, ByVal pnMatTpoProd As Variant, ByVal pnMatAgencias As Variant)
    Dim psGrupoCod As String
    
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    
    Call oNivApr.dInsertaGruposApr(psGrupoDesc)
    psGrupoCod = oNivApr.RecuperaGrupoAprCod(psGrupoDesc)
    
    Call dInsertaGruposAprDetalle(psGrupoCod, pnMatTpoProd, 1)
    Call dInsertaGruposAprDetalle(psGrupoCod, pnMatAgencias, 2)
    Set oNivApr = Nothing
End Sub

Public Sub dActualizaGruposApr(ByVal psGrupoCod As String, ByVal psGrupoDesc As String, ByVal pnMatTpoProd As Variant, ByVal pnMatAgencias As Variant)
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    
    Call oNivApr.dActualizaGruposApr(psGrupoCod, psGrupoDesc)
    Call oNivApr.dEliminaGruposAprDetalle(psGrupoCod)
    Call dInsertaGruposAprDetalle(psGrupoCod, pnMatTpoProd, 1)
    Call dInsertaGruposAprDetalle(psGrupoCod, pnMatAgencias, 2)
    Set oNivApr = Nothing
End Sub

Public Sub dInsertaGruposAprDetalle(ByVal psGrupoCod As String, ByVal pnMatLista As Variant, ByVal pnTipoMat As Integer)
    Dim i As Integer
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    
    For i = 1 To UBound(pnMatLista)
        Call oNivApr.dInsertaGruposAprDetalle(psGrupoCod, pnMatLista(i, 0), pnTipoMat)
    Next i
    Set oNivApr = Nothing
End Sub

Public Sub dEliminaGruposApr(ByVal psGrupoCod As String)
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    Call oNivApr.dEliminaGrupoApr(psGrupoCod)
    Set oNivApr = Nothing
End Sub

Public Function VerificaDatosGrupoApr(ByVal psGrupoCod As String, ByVal psGrupoDesc As String, ByVal pnMatTpoProd As Variant, ByVal pnMatAgencias As Variant) As Boolean
    Dim i  As Integer
    Dim psTpoProd As String, psAgencias As String
    
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    
    For i = 1 To UBound(pnMatTpoProd)
        psTpoProd = psTpoProd & pnMatTpoProd(i, 0) & ","
    Next i
    psTpoProd = Mid(psTpoProd, 1, Len(psTpoProd) - 1)
    For i = 1 To UBound(pnMatAgencias)
        psAgencias = psAgencias & pnMatAgencias(i, 0) & ","
    Next i
    psAgencias = Mid(psAgencias, 1, Len(psAgencias) - 1)
    VerificaDatosGrupoApr = oNivApr.VerificaDatosGrupoApr(psGrupoCod, psGrupoDesc, psTpoProd, psAgencias)
    Set oNivApr = Nothing
End Function

Public Function VerificaDatosNivApr(ByVal psNivAprCod As String, ByVal psNivAprDesc As String) As Boolean
    Dim i  As Integer
    Dim psTpoProd As String, psAgencias As String
    
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    VerificaDatosNivApr = oNivApr.VerificaDatosNivApr(psNivAprCod, psNivAprDesc)
    Set oNivApr = Nothing
End Function

Public Sub dInsertaNivApr(ByVal psNivAprDesc As String, ByVal pnTipoReg As Integer, ByVal pnNumCantFirmas As Integer, _
                          ByVal pnCorrigeSug As Integer, ByVal pnValidaAg As Integer, ByVal pnMatValores As Variant)
    Dim psNivAprCod As String
    
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    
    psNivAprCod = oNivApr.dInsertaNivApr(psNivAprDesc, pnTipoReg, pnNumCantFirmas, pnCorrigeSug, pnValidaAg)
     
    Call dInsertaNivAprValores(psNivAprCod, pnMatValores)
    Set oNivApr = Nothing
End Sub

Public Sub dActualizaNivApr(ByVal psNivAprCod As String, ByVal psNivAprDesc As String, ByVal pnTipoReg As Integer, _
                            ByVal pnNumCantFirmas As Integer, ByVal pnCorrigeSug As Integer, ByVal pnValidaAg As Integer, _
                            ByVal pnMatValores As Variant)
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    
    Call oNivApr.dActualizaNivApr(psNivAprCod, psNivAprDesc, pnTipoReg, pnNumCantFirmas, pnCorrigeSug, pnValidaAg)
    Call oNivApr.dEliminaNivAprValores(psNivAprCod)
    Call dInsertaNivAprValores(psNivAprCod, pnMatValores)
    Set oNivApr = Nothing
End Sub

Public Sub dInsertaNivAprValores(ByVal psNivAprCod As String, ByVal pnMatValores As Variant)
    Dim i As Integer
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    
    For i = 0 To UBound(pnMatValores) - 1
        Call oNivApr.dInsertaNivAprValores(psNivAprCod, pnMatValores(i, 0), pnMatValores(i, 1))
    Next i
    Set oNivApr = Nothing
End Sub

Public Sub dEliminaNivApr(ByVal psNivAprCod As String)
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    Call oNivApr.dEliminaNivApr(psNivAprCod)
    Set oNivApr = Nothing
End Sub

Public Sub dActualizaOrdenNivel(ByVal psNivAprCodAct As String, ByVal psNivAprCodReemp As String, ByVal pnOrdenAct As Integer, ByVal pnOrdenReemp As Integer)
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    Call oNivApr.dActualizaOrdenNivel(psNivAprCodAct, psNivAprCodReemp, pnOrdenAct, pnOrdenReemp)
    Set oNivApr = Nothing
End Sub

Public Sub dInsertaParamGruposNiveles(ByVal psGrupoCod As String, ByVal psNivelesDesc As String, ByVal pnDesdeR1 As Double, ByVal nHastaR1 As Double, _
                                      ByVal nDesdeR2 As Double, ByVal nHastaR2 As Double, _
                                      Optional ByVal pnTpoCliente As Integer)
                                      'RECO20150129 ERS173-2014 Se agrego tipo de cliente
    Dim i As Integer
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    
    'Call oNivApr.dEliminaParamGruposNiveles(psGrupoCod)
    Call oNivApr.dInsertaParamGruposNiveles(psGrupoCod, psNivelesDesc, pnDesdeR1, nHastaR1, nDesdeR2, nHastaR2, pnTpoCliente)
    Set oNivApr = Nothing
End Sub

Public Sub dEliminaParamApr(ByVal psParamCod As String)
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    Call oNivApr.dEliminaParamApr(psParamCod)
    Set oNivApr = Nothing
End Sub

Public Sub dActualizaParamGruposNiveles(ByVal psParamCod As String, ByVal psGrupoCod As String, ByVal psNivelesDesc As String, ByVal pnDesdeR1 As Double, _
                                        ByVal nHastaR1 As Double, ByVal nDesdeR2 As Double, ByVal nHastaR2 As Double, _
                                        Optional ByVal pnTpoCliente As Integer)
                                        'RECO20150129 ERS173-2014 Se agrego tipo de cliente
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    Call oNivApr.dActualizaParamGruposNiveles(psParamCod, psGrupoCod, psNivelesDesc, pnDesdeR1, nHastaR1, nDesdeR2, nHastaR2, pnTpoCliente)
    Set oNivApr = Nothing
End Sub

Public Sub dDelegarAprobacion(ByVal psUserDelega As String, ByVal psUserDelegado As String, ByVal pdFechaDesde As String, ByVal pdFechaHasta As String)
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    Call oNivApr.dDelegarAprobacion(psUserDelega, psUserDelegado, pdFechaDesde, pdFechaHasta)
    Set oNivApr = Nothing
End Sub

Public Sub dRetirarDelegacionAprobacion(ByVal psUserDelega As String)
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    Call oNivApr.dRetirarDelegacionAprobacion(psUserDelega)
    Set oNivApr = Nothing
End Sub

Public Sub dInsertaTiposExoneraciones(ByVal psExoneraDesc As String, ByVal pnTipoExonera As Integer)
    Dim i As Integer
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    'Call oNivApr.dEliminaTipoExoneracion
    Call oNivApr.dInsertaTiposExoneraciones(psExoneraDesc, pnTipoExonera)
    Set oNivApr = Nothing
End Sub

Public Sub dActualizaTiposExoneraciones(ByVal psExoneraCod As String, ByVal psExoneraDesc As String, ByVal pnTipoExonera As Integer)
    Dim i As Integer
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    Call oNivApr.dActualizaTiposExoneraciones(psExoneraCod, psExoneraDesc, pnTipoExonera)
    Set oNivApr = Nothing
End Sub

Public Sub dInsertaNivelesExoneraciones(ByVal psTipoExoneraCod As String, ByVal prsDatos As ADODB.Recordset)
    Dim i As Integer
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    If Not prsDatos Is Nothing Then
        Call oNivApr.dEliminaNivelesExoneracion(psTipoExoneraCod)
        For i = 0 To prsDatos.RecordCount - 1
            Call oNivApr.dInsertaNivelesExoneraciones(psTipoExoneraCod, Trim(Right(prsDatos![Nivel de Aprobacion], 10)), prsDatos!Desde, prsDatos!Hasta)
            prsDatos.MoveNext
        Next i
    End If
    Set oNivApr = Nothing
End Sub

Public Sub dSolicitaModifAprobacionResultado(ByVal psCtaCod As String, ByVal psComent As String, _
                            Optional ByVal pnTipoReg As TipoNivelReg = TipoNivelReg.gTipoNivelAprob, Optional ByVal psMovNro As String = "") 'RECO20160526 ERS0022016
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    'Call oNivApr.dSolicitaModifAprobacionResultado(psCtaCod, psComent)  'RECO20160526 ERS0022016
    Call oNivApr.dSolicitaModifAprobacionResultado(psCtaCod, psComent, pnTipoReg, psMovNro)  'RECO20160526 ERS0022016
    Call dDesactivaAprobacionResultado(psCtaCod)
    Set oNivApr = Nothing
End Sub

Public Sub dDesactivaAprobacionResultado(ByVal psCtaCod As String)
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    Call oNivApr.dDesactivaAprobacionResultado(psCtaCod)
    Set oNivApr = Nothing
End Sub

Public Function AprobacionCreditoPorNivel(ByVal psCtaCod As String, ByVal psNivAprCod As String, ByVal pnCuotas As Integer, ByVal pnTasa As Double, _
                                          ByVal pnMonto As Double, ByVal psComent As String, ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUsu As String) As String
Dim oDCredito As COMDCredito.DCOMCredActBD
Dim oAmpliado As COMDCredito.DCOMAmpliacion
Dim oFunciones As COMNContabilidad.NCOMContFunciones
Dim bAmpliacion As Boolean
Dim sMovNroApr As String
Dim bTransac As String
Set oDCredito = New COMDCredito.DCOMCredActBD
Set oAmpliado = New COMDCredito.DCOMAmpliacion
Set oFunciones = New COMNContabilidad.NCOMContFunciones

On Error GoTo ControlError
    
    bAmpliacion = oAmpliado.ValidaCreditoaAmpliar(psCtaCod)
    Set oAmpliado = Nothing
    AprobacionCreditoPorNivel = ValidaGarantia(psCtaCod, pdFecSis, pnMonto, bAmpliacion)
    If AprobacionCreditoPorNivel <> "" Then
        Exit Function
    End If
    
    Set oFunciones = New COMNContabilidad.NCOMContFunciones
    sMovNroApr = oFunciones.GeneraMovNro(pdFecSis, psCodAge, psCodUsu)
    Set oFunciones = Nothing
    
    bTransac = False
    Call oDCredito.dBeginTrans
    bTransac = True
    'Tasa
    Call oDCredito.dUpdateProducto(psCtaCod, pnTasa, , , , , False)
    Call oDCredito.dUpdateProductoTasaInteres(psCtaCod, gColocLineaCredTasasIntCompNormal, pnTasa, False)
    
    'Monto
    Call oDCredito.dUpdateColocaciones(psCtaCod, , , pnMonto, , , pdFecSis, False)
    
    If oDCredito.ExisteColocacEstadoNivApr(psCtaCod) Then
        Call oDCredito.dUpdateColocacEstadoUltimoRegNivApr(psCtaCod, pdFecSis, gColocEstSug, pnCuotas, pnMonto, "Nueva Sugerencia Credito por Niv. Aprob.")
    Else
        Call oDCredito.dInsertColocacEstadoNivApr(psCtaCod, pdFecSis, gColocEstSug, pnCuotas, pnMonto, "Nueva Sugerencia Credito por Niv. Aprob.", False)
    End If
    Call oDCredito.dInsertaCredNivAprResultado(psCtaCod, psNivAprCod, pnCuotas, pnTasa, pnMonto, psComent, sMovNroApr, False)
    
    Call oDCredito.dInsertProductoPersonaNivAprApoderado(psCtaCod, psCodUsu)
    
    oDCredito.dCommitTrans

Exit Function

ControlError:
    If bTransac Then
        oDCredito.dRollbackTrans
        AprobacionCreditoPorNivel = "Hubo un error en la aprobación"
        Set oDCredito = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso Aprobacion del Credito por Nivel", Err.Description
End Function

Private Function ValidaGarantia(ByVal psCtaCod As String, ByVal pdGarantia As Date, ByVal pnMonto As Double, Optional ByVal bAmpliacion As Boolean) As String
Dim oCred As COMDCredito.DCOMCredito
Dim nPorcGarant As Double
Dim nMontoGarantia As Double
Dim rsCred As ADODB.Recordset

    On Error GoTo ErrorValidaSugerenciaAprobacion
    Set oCred = New COMDCredito.DCOMCredito
    nMontoGarantia = oCred.RecuperaMontoGarantiaCredito(psCtaCod, pdGarantia, bAmpliacion)
    If Mid(psCtaCod, 6, 3) = "517" Then
        Set rsCred = oCred.RecuperaParametro(1029)
        If Not rsCred.EOF And Not rsCred.BOF Then
            nPorcGarant = rsCred!nParamValor
            pnMonto = pnMonto * nPorcGarant
        End If
    End If
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
'END JUEZ *************************************************************************
'ALPA 20140321*********************************************************************
Public Function VerificaDatosNivAprCV(ByVal psCargoCod As String, ByVal psCargoDesc As String, ByVal psNivAprCod As String) As Boolean
    Dim i  As Integer
    Dim psTpoProd As String, psAgencias As String
    
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    VerificaDatosNivAprCV = oNivApr.VerificaDatosNivAprCV(psCargoCod, psCargoDesc, psNivAprCod)
    Set oNivApr = Nothing
End Function

Public Sub dInsertaNivAprCV(ByVal psNivAprCod As String, ByVal psCargoCod As String, ByVal pnTipoReg As String, ByVal pnNumCantFirmas As Integer, _
                            ByVal pnValidaAg As Integer, ByVal pnMontoDesde As Currency, ByVal pnMontoHasta As Currency, _
                            ByVal pnTCCmas As Currency, ByVal pnTCVmas As Currency)
    
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    
    psNivAprCod = oNivApr.dInsertaNivAprCV(psNivAprCod, psCargoCod, pnTipoReg, pnNumCantFirmas, pnValidaAg, pnMontoDesde, pnMontoHasta, pnTCCmas, pnTCVmas)
    Set oNivApr = Nothing
End Sub

Public Function ObtenerNivelesAprobacionCompraVentaxMonto(ByVal pnMonto As Currency, ByVal pnDifTCambio As Currency, ByVal pnTipo As Integer, ByRef pnNivel As String, ByVal nTipoEspecial As Integer) As Integer
'pnTipoCambio           =       1--Compra
'pnTipoCambio           =       2--Venta
    Dim oRs As ADODB.Recordset
    Dim nValidado As Integer
    nValidado = 0
    Set oRs = New ADODB.Recordset
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    Set oRs = oNivApr.ObtenerNivelesAprobacionCompraVentaxMonto(pnMonto, nTipoEspecial)
    If Not (oRs.BOF Or oRs.EOF) Then
        Do While Not oRs.EOF
            If pnTipo = 1 Then
                If pnDifTCambio > 0 Then
                    If Abs(pnDifTCambio) <= oRs!nTCCmas Then
                        nValidado = 1
                        pnNivel = oRs!cNivelCod
                        Exit Do
                        End If
                End If
            Else
                If pnDifTCambio < 0 Then
                    If Abs(pnDifTCambio) >= oRs!nTCVMas Then
                        nValidado = 1
                        pnNivel = oRs!cNivelCod
                        Exit Do
                    End If
                End If
            End If
            pnNivel = oRs!cNivelCod
        oRs.MoveNext
        Loop
    End If
    ObtenerNivelesAprobacionCompraVentaxMonto = nValidado
End Function

Public Function ObtenerCantidadAprobacionMovCompraVenta(ByVal psMovNro As String, ByRef nEstado As Integer) As Integer
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    Dim oRs As ADODB.Recordset
    Set oRs = oNivApr.ObtenerCantidadAprobacionMovCompraVenta(psMovNro)
    If Not (oRs.BOF Or oRs.EOF) Then
        nEstado = oRs!nEstado
        ObtenerCantidadAprobacionMovCompraVenta = oRs!nNroFirmas - oRs!nContador
    End If
    Set oRs = Nothing
End Function
'**********************************************************************************
'EJVG20160205 ERS002-2016***
Public Function GuardarAutorizacionConfig(ByVal psMovNro As String, ByVal prsAutorizacion As ADODB.Recordset, ByRef pvEscalonaConfig As TEscalonamientoConf) As Boolean
    Dim bTransac As Boolean
    Set oNivApr = New COMDCredito.DCOMNivelAprobacion
    
    On Error GoTo ErrGuardar
    
    oNivApr.dBeginTrans
    bTransac = True

    Do While Not prsAutorizacion.EOF
        oNivApr.dActualizarTipoAutorizacion prsAutorizacion!Codigo, IIf(prsAutorizacion!Habilitado = "1", 1, 2), psMovNro
        prsAutorizacion.MoveNext
    Loop
    
    oNivApr.dActualizarEscalonaConfig pvEscalonaConfig.nMoraPorcentaje, pvEscalonaConfig.nMontoCuoMayorA, pvEscalonaConfig.nMontoCuoMenorIgual, pvEscalonaConfig.nMontoCreMayorA, pvEscalonaConfig.nMontoCreMenorIgual
    
    oNivApr.dCommitTrans
    bTransac = False
    Set oNivApr = Nothing
    GuardarAutorizacionConfig = True
    Exit Function
ErrGuardar:
    GuardarAutorizacionConfig = False
    If bTransac Then
        oNivApr.dRollbackTrans
        Set oNivApr = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso Guardar Autorizaciones", Err.Description
End Function
'END EJVG *******

'RECO20160526 ERS0022016*****************************************************
Public Function ListaCredNivExoneracionCta(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oCredNiv As New COMDCredito.DCOMNivelAprobacion
    Set ListaCredNivExoneracionCta = oCredNiv.ListaCredNivExoneracionCta(psCtaCod)
End Function
Public Function RegistrarCredNivExoneraCabecera(ByVal psCtaCod As String, ByVal pnEstado As Integer, ByVal psMovNroReg As String) As Long
    Dim oCredNiv As New COMDCredito.DCOMNivelAprobacion
    Dim oDR As New ADODB.Recordset
    Set oDR = oCredNiv.RegistrarCredNivExoneraCabecera(psCtaCod, pnEstado, psMovNroReg)
    
    If Not (oDR.EOF And oDR.BOF) Then
        RegistrarCredNivExoneraCabecera = oDR!nID
    Else
        RegistrarCredNivExoneraCabecera = 0
    End If
    
End Function
Public Sub RegistrarCredNivExoneraDetalle(ByVal pnId As Long, ByVal pnItem As Integer, ByVal psExoneracion As String, ByVal psDescripcion As String, ByVal psNivAprCod As String _
                                       , ByVal pbManual As Integer, ByVal pnEstado As Integer, ByVal psMovNroReg As String)
    Dim oCredNiv As New COMDCredito.DCOMNivelAprobacion
    Call oCredNiv.RegistrarCredNivExoneraDetalle(pnId, pnItem, psExoneracion, psDescripcion, psNivAprCod, pbManual, pnEstado, psMovNroReg)
End Sub
'RECO FIN********************************************************************

