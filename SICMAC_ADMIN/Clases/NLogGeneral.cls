VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NLogGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
'** Nombre : NLogGeneral
'** Descripción : Clase para manejo de Acta de Conformidad creado segun ERS062-2013
'** Creación : EJVG, 20131009 09:00:00 AM
'**********************************************************************************
Option Explicit
Public Type TActaConformidadOrden
    nMovItem As Long
    sCtaContCod As String
    sObjeto As String
    sDescripcion As String
    nCantidad As Integer
    nTotal As Currency
End Type
Public Type TActaConformidadContrato
    nNPago As Integer
    sCtaContCod As String
    sDescripcion As String
    nMonto As Currency
End Type
'PASI20140917 ERS0772014
Public Type TContratoBS
    sAgeCod As String
    sObjeto As String
    sDescripcion As String
    nCantidad As Integer
    nPrecUnit As Currency
    nTotal As Currency
    sCtaContCod As String
    nMovItem As Integer
End Type
Public Type TComprobanteOrden
    nMovItem As Long
    sCtaContCod As String
    sObjeto As String
    sDescripcion As String
    nCantidad As Double
    nTotal As Currency
End Type
Public Type TComprobanteContratoxCronograma
    nNPago As Integer
    sCtaContCod As String
    sDescripcion As String
    nMonto As Currency
End Type
Public Type TComprobanteContratoxBien
    nMovItem As Long
    sCtaContCod As String
    sObjeto As String
    sDescripcion As String
    nCantidad As Double
    nTotal As Currency
End Type
Public Type TComprobanteContratoxObra
    sCtaContCod As String
    sDescripcion As String
    nMonto As Currency
End Type
'end PASI
Public Function GrabarActaConformidad_Orden(ByVal pdFecha As Date, ByVal psAgeCod As String, _
                                        ByVal psUserCod As String, ByVal psOpeCod As String, _
                                        ByVal pnDocOrigenTpo As Integer, ByVal pnDocTpo As Integer, _
                                        ByVal psDocReferencia As String, ByVal psAreaAgeCod As String, _
                                        ByVal psSubAreaDesc As String, ByVal pnMoneda As Moneda, _
                                        ByRef psActaNro As String, ByVal psProveedorCod As String, _
                                        ByVal psDescripcion As String, ByVal psObservacion As String, _
                                        ByRef pMatDetOrden() As TActaConformidadOrden, ByVal psCtaContCod As String, _
                                        ByVal pnTpoCambio As Currency, ByRef psMovNro As String, _
                                        Optional ByVal pnMovNroRef As Long = 0) As Long
    Dim oMov As New DMov
    Dim bTrans As Boolean
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lnMovItem As Integer
    Dim iref As Integer
    Dim lnMonto As Currency
    On Error GoTo ErrGrabar
    
    oMov.BeginTrans
    bTrans = True
    
    psActaNro = oMov.GetCorrelativoActaConformidad(pnDocTpo, psAgeCod, CStr(Year(pdFecha))) 'Volvemos a generar para evitar duplicidad
    If Len(Trim(psActaNro)) = 0 Then
        oMov.RollBackTrans
        Set oMov = Nothing
        Exit Function
    End If
    
    lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUserCod)
    'oMov.InsertaMov lsMovNro, psOpeCod, Left(psDescripcion, 250), gMovEstContabMovContable, gMovFlagVigente
    oMov.InsertaMov lsMovNro, psOpeCod, Left(psDescripcion, 250), gMovEstContabNoContable, gMovFlagVigente 'EJVG20140416
    lnMovNro = oMov.GetnMovNro(lsMovNro)
    oMov.InsertaMovGasto lnMovNro, psProveedorCod, "1"
    oMov.InsertaMovDoc lnMovNro, pnDocTpo, psActaNro, Left(lsMovNro, 8)
    oMov.InsertaMovActaConformidad lnMovNro, psAreaAgeCod, psSubAreaDesc, pnMoneda, pnDocOrigenTpo, psDocReferencia, psDescripcion, psObservacion
    If pnMovNroRef <> 0 Then
        oMov.InsertaMovActaConformidadOrden lnMovNro, pnMovNroRef
    End If

    For lnMovItem = 1 To UBound(pMatDetOrden)
        lnMonto = lnMonto + pMatDetOrden(lnMovItem).nTotal
        oMov.InsertaMovCta lnMovNro, lnMovItem, pMatDetOrden(lnMovItem).sCtaContCod, pMatDetOrden(lnMovItem).nTotal
        oMov.InsertaMovActaConformidadItem lnMovNro, lnMovItem, pMatDetOrden(lnMovItem).sDescripcion
        If pnDocOrigenTpo = LogTipoDocOrigenActaConformidad.OrdenCompra Then
            oMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjBienesServicios
            oMov.InsertaMovBS lnMovNro, lnMovItem, 1, pMatDetOrden(lnMovItem).sObjeto
            oMov.InsertaMovCant lnMovNro, lnMovItem, pMatDetOrden(lnMovItem).nCantidad
        End If
        If pnMovNroRef <> 0 Then
            oMov.InsertaMovActaConformidadOrdenDet lnMovNro, lnMovItem, pMatDetOrden(lnMovItem).nMovItem
        End If
    Next lnMovItem
    oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContCod, lnMonto * -1
    
    If pnMoneda = gMonedaExtranjera Then
        oMov.GeneraMovME lnMovNro, lsMovNro, pnTpoCambio, True
    End If
    'oMov.ActualizaSaldoMovimiento lsMovNro, "+"
    
    oMov.CommitTrans
    GrabarActaConformidad_Orden = lnMovNro
    psMovNro = lsMovNro
    bTrans = False
    
    Set oMov = Nothing
    Exit Function
ErrGrabar:
    If bTrans Then
        oMov.RollBackTrans
        Set oMov = Nothing
    End If
End Function
Public Function GrabarActaConformidad_Contrato(ByVal pdFecha As Date, ByVal psAgeCod As String, _
                                        ByVal psUserCod As String, ByVal psOpeCod As String, _
                                        ByVal pnDocOrigenTpo As Integer, ByVal pnDocTpo As Integer, _
                                        ByVal psDocReferencia As String, ByVal psAreaAgeCod As String, _
                                        ByVal psSubAreaDesc As String, ByVal pnMoneda As Moneda, _
                                        ByRef psActaNro As String, ByVal psProveedorCod As String, _
                                        ByVal psDescripcion As String, ByVal psObservacion As String, _
                                        ByRef pMatDetContrato() As TActaConformidadContrato, ByVal psCtaContCod As String, _
                                        ByVal pnTpoCambio As Currency, ByRef psMovNro As String) As Long
    Dim oMov As New DMov
    Dim bTrans As Boolean
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lnMovItem As Integer
    Dim iref As Integer
    Dim lnMonto As Currency
    On Error GoTo ErrGrabar
    
    oMov.BeginTrans
    bTrans = True
    
    psActaNro = oMov.GetCorrelativoActaConformidad(pnDocTpo, psAgeCod, CStr(Year(pdFecha))) 'Volvemos a generar para evitar duplicidad
    If Len(Trim(psActaNro)) = 0 Then
        oMov.RollBackTrans
        Set oMov = Nothing
        Exit Function
    End If
    
    lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUserCod)
    'oMov.InsertaMov lsMovNro, psOpeCod, Left(psDescripcion, 250), gMovEstContabMovContable, gMovFlagVigente
    oMov.InsertaMov lsMovNro, psOpeCod, Left(psDescripcion, 250), gMovEstContabNoContable, gMovFlagVigente 'EJVG20140416
    lnMovNro = oMov.GetnMovNro(lsMovNro)
    oMov.InsertaMovGasto lnMovNro, psProveedorCod, "1"
    oMov.InsertaMovDoc lnMovNro, pnDocTpo, psActaNro, Left(lsMovNro, 8)
    oMov.InsertaMovActaConformidad lnMovNro, psAreaAgeCod, psSubAreaDesc, pnMoneda, pnDocOrigenTpo, psDocReferencia, psDescripcion, psObservacion
    oMov.InsertaMovActaConformidadContrato lnMovNro, psDocReferencia

    For lnMovItem = 1 To UBound(pMatDetContrato)
        lnMonto = lnMonto + pMatDetContrato(lnMovItem).nMonto
        oMov.InsertaMovCta lnMovNro, lnMovItem, pMatDetContrato(lnMovItem).sCtaContCod, pMatDetContrato(lnMovItem).nMonto
        oMov.InsertaMovActaConformidadItem lnMovNro, lnMovItem, pMatDetContrato(lnMovItem).sDescripcion
        oMov.InsertaMovActaConformidadContratoDet lnMovNro, lnMovItem, pMatDetContrato(lnMovItem).nNPago
        oMov.ActualizaEstadoCuotaCronograma psDocReferencia, 0, pMatDetContrato(lnMovItem).nNPago, 2 'Modificado PASI se agredo valor 0 para el parametro NContref
    Next lnMovItem
    oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContCod, lnMonto * -1
    
    If pnMoneda = gMonedaExtranjera Then
        oMov.GeneraMovME lnMovNro, lsMovNro, pnTpoCambio, True
    End If
    'oMov.ActualizaSaldoMovimiento lsMovNro, "+"
    
    oMov.CommitTrans
    GrabarActaConformidad_Contrato = lnMovNro
    psMovNro = lsMovNro
    bTrans = False
    
    Set oMov = Nothing
    Exit Function
ErrGrabar:
    If bTrans Then
        oMov.RollBackTrans
        Set oMov = Nothing
    End If
End Function
Public Function ExtornaActaConformidad(ByVal pdFecha As Date, ByVal psAgeCod As String, ByVal psUserCod As String, ByVal psOpeCod As String, ByVal psGlosa As String, ByVal pnMovNroRef As Long) As Boolean
    Dim oMov As New DMov
    Dim bTrans As Boolean
    Dim lsMovNro As String
    Dim lnMovNro As Long
    On Error GoTo ErrExtornaActaConformidad
    
    oMov.BeginTrans
    bTrans = True
    
    lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUserCod)
    oMov.InsertaMov lsMovNro, psOpeCod, Left(psGlosa, 300), gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(lsMovNro)
    oMov.EliminaMov pnMovNroRef
    oMov.InsertaMovRef lnMovNro, pnMovNroRef
    oMov.ActualizaEstadoCuotasCronogramaxExtActaConformidad pnMovNroRef
    
    oMov.CommitTrans
    bTrans = False
    ExtornaActaConformidad = True
    Set oMov = Nothing
    Exit Function
ErrExtornaActaConformidad:
    If bTrans Then
        oMov.RollBackTrans
        Set oMov = Nothing
        ExtornaActaConformidad = False
    End If
End Function
Public Function GrabaComprobante(ByVal pdFecha As Date, ByVal psUserCod As String, ByVal psAgeCod As String, ByVal psOpeCod As String, ByVal psGlosa As String, _
                    ByVal pnMovNroAC As Long, ByVal pnMoneda As Moneda, ByVal pnDocTpo As Integer, ByVal psDocNro As String, pdDocFecha As Date, _
                    ByVal pnTpoPago As Integer, ByVal psIFICod As String, ByVal psIFICtaCod As String, ByVal pnMovNroComprob As Long) As Boolean
    Dim oMov As New DMov
    Dim bTrans As Boolean
    Dim lsMovNro As String
    Dim lnMovNro As Long
    On Error GoTo ErrGrabaComprobante
    
    oMov.BeginTrans
    bTrans = True
    
    lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUserCod)
    If pnMovNroComprob > 0 Then
        ExtornaComprobante pnMovNroComprob, oMov, pnMovNroAC
    End If
    oMov.InsertaMov lsMovNro, psOpeCod, psGlosa, gMovEstContabMovContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(lsMovNro)
    oMov.InsertaMovComprobante lnMovNro, pnMoneda, pnDocTpo, psDocNro, pdDocFecha, pnTpoPago, psIFICod, psIFICtaCod
    oMov.InsertaMovRef lnMovNro, pnMovNroAC
    oMov.ActualizaEstadoCuotasCronogramaxRegComprobante pnMovNroAC, False
    
    oMov.CommitTrans
    bTrans = False
    GrabaComprobante = True
    Set oMov = Nothing
    Exit Function
ErrGrabaComprobante:
    If bTrans Then
        oMov.RollBackTrans
        Set oMov = Nothing
        GrabaComprobante = False
    End If
End Function
Public Function ExtornaComprobante(ByVal pnMovNroComprob As Long, Optional ByRef poMov As DMov = Nothing, Optional ByVal pnMovNroAC As Long = 0) As Boolean
    Dim oMov As DMov
    Dim bTrans As Boolean
    Dim lsMovNro As String
    Dim lnMovNro As Long
    On Error GoTo ErrGrabaComprobante
    
    If poMov Is Nothing Then
        Set oMov = New DMov
        oMov.BeginTrans
    Else
        Set oMov = poMov
    End If
    bTrans = True
    oMov.EliminaMov pnMovNroComprob
    If pnMovNroAC > 0 Then
        oMov.ActualizaEstadoCuotasCronogramaxRegComprobante pnMovNroAC, True
    End If
    If poMov Is Nothing Then
        oMov.CommitTrans
    End If
    bTrans = False
    ExtornaComprobante = True
    If poMov Is Nothing Then
        Set oMov = Nothing
    End If
    Exit Function
ErrGrabaComprobante:
    If bTrans Then
        If poMov Is Nothing Then
            oMov.RollBackTrans
            Set oMov = Nothing
        End If
        ExtornaComprobante = False
    End If
End Function
'PASI20140917 ERS0772014
Public Function GrabarComprobanteOrden(ByVal pdFecha As Date, _
                                        ByVal psAgeCod As String, _
                                        ByVal psUserCod As String, _
                                        ByVal psOpeCod As String, _
                                        ByVal psGlosa As String, _
                                        ByVal psProveedorCod As String, _
                                        ByVal psAreaAgeCod As String, _
                                        ByVal pnMoneda As Integer, _
                                        ByVal pnDocOrigenTpo As Integer, _
                                        ByVal psObservacion As String, _
                                        ByVal pnDocTpo As Integer, _
                                        ByVal psDocNro As String, _
                                        ByVal psFechaDoc As Date, _
                                        ByVal pnTpoPago As Integer, _
                                        ByVal psIFICod As String, _
                                        ByVal psIFICtaCod As String, _
                                        ByRef pMatDetOrden() As TComprobanteOrden, _
                                        ByVal psCtaContCod As String, _
                                        ByVal pnTpoCambio As Currency, _
                                        ByRef psMovNro As String, _
                                        Optional ByVal pnMovNroRef As Long = 0) As Long 'PASI20150107 agrego psFechaDoc
    
    Dim oMov As New DMov
    Dim bTrans As Boolean
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lnMovItem As Integer
    Dim lnMonto As Currency
    On Error GoTo ErrorGrabarComprobanteOrden
    
    oMov.BeginTrans
    bTrans = True
    
    lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUserCod)
    oMov.InsertaMov lsMovNro, psOpeCod, psGlosa, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(lsMovNro)
    oMov.InsertaMovGasto lnMovNro, psProveedorCod, "1"
    'oMov.InsertaMovDoc lnMovNro, pnDocTpo, psDocNro, Left(lsMovNro, 8)
    'oMov.InsertaMovComprobanteReg lnMovNro, psAreaAgeCod, pnMoneda, pnDocOrigenTpo, psObservacion, pnTpoPago, psIFICod, psIFICtaCod, pnDocTpo, psDocNro, Left(lsMovNro, 8)
    oMov.InsertaMovComprobanteReg lnMovNro, psAreaAgeCod, pnMoneda, pnDocOrigenTpo, psObservacion, pnTpoPago, psIFICod, psIFICtaCod, pnDocTpo, psDocNro, Format(psFechaDoc, "yyyyMMdd") 'PASI20150107
    If pnMovNroRef <> 0 Then
        oMov.InsertaMovComprobanteOrden lnMovNro, pnMovNroRef
    End If
    
    For lnMovItem = 1 To UBound(pMatDetOrden)
        lnMonto = lnMonto + pMatDetOrden(lnMovItem).nTotal
        oMov.InsertaMovCta lnMovNro, lnMovItem, pMatDetOrden(lnMovItem).sCtaContCod, pMatDetOrden(lnMovItem).nTotal
        oMov.InsertaMovComprobanteItem lnMovNro, lnMovItem, pMatDetOrden(lnMovItem).sDescripcion
        If pnDocOrigenTpo = LogTipoDocOrigenComprobante.OrdenCompra Or pnDocOrigenTpo = LogTipoDocOrigenComprobante.CompraLibre Then
            oMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjBienesServicios
            oMov.InsertaMovBS lnMovNro, lnMovItem, 1, pMatDetOrden(lnMovItem).sObjeto
            oMov.InsertaMovCant lnMovNro, lnMovItem, pMatDetOrden(lnMovItem).nCantidad
        End If
        If pnMovNroRef <> 0 Then
            oMov.InsertaMovComprobanteOrdenDet lnMovNro, lnMovItem, pMatDetOrden(lnMovItem).nMovItem
        End If
    Next lnMovItem
    oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContCod, lnMonto * -1
    If pnMoneda = gMonedaExtranjera Then
        oMov.GeneraMovME lnMovNro, lsMovNro, pnTpoCambio, True
    End If
    oMov.CommitTrans
    GrabarComprobanteOrden = lnMovNro
    psMovNro = lsMovNro
    bTrans = False
    
    Set oMov = Nothing
    Exit Function
ErrorGrabarComprobanteOrden:
    If bTrans Then
        oMov.RollBackTrans
        Set oMov = Nothing
    End If
End Function
'vapa20160811
Public Function ValidaComprobanteReg(ByVal psDocTipo As Integer, ByVal psDocNro As String, ByVal psCodProveedor As String) As String
 Dim rs As ADODB.Recordset
 Dim sSql As String
Dim oconect As DConecta
Set oconect = New DConecta
    oconect.AbreConexion
    sSql = "stp_sel_ValidaMovComprobanteReg '" & psDocTipo & "','" & psDocNro & "','" & psCodProveedor & "'"
     Set rs = oconect.CargaRecordSet(sSql)
    ValidaComprobanteReg = rs!cVal
   oconect.CierraConexion
    Set oconect = Nothing
End Function
'end vapa

Public Function GrabaComprobanteContratoxServicio(ByVal pdFecha As Date, _
                                         ByVal psAgeCod As String, _
                                         ByVal psUserCod As String, _
                                         ByVal psOpeCod As String, _
                                         ByVal psncontrato As String, _
                                         ByVal pncontref As Integer, _
                                         ByVal psGlosa As String, _
                                         ByVal psProveedorCod As String, _
                                         ByVal psAreaAgeCod As String, _
                                         ByVal pnMoneda As Integer, _
                                         ByVal pnDocOrigenTpo As Integer, _
                                         ByVal psObservacion As String, _
                                         ByVal pnDocTpo As Integer, _
                                         ByVal psDocNro As String, _
                                         ByVal psFechaDoc As Date, _
                                         ByVal pnTpoPago As Date, _
                                         ByVal psIFICod As String, _
                                         ByVal psIFICtaCod As String, _
                                         ByRef pMatDetContrato() As TComprobanteContratoxCronograma, _
                                         ByRef pMatDetOrden() As TComprobanteOrden, _
                                         ByVal psCtaContCod As String, _
                                         ByVal pnTpoCambio As Currency, _
                                         ByRef psMovNro As String) 'PASI20150107 agrego psFechaDoc
    
    
    Dim oMov As New DMov
    Dim bTrans As Boolean
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lnMovItem As Integer
    Dim lnMonto As Currency
    
    On Error GoTo ErrorGrabar
    
    oMov.BeginTrans
    bTrans = True
    
    lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUserCod)
    oMov.InsertaMov lsMovNro, psOpeCod, psGlosa, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(lsMovNro)
    oMov.InsertaMovGasto lnMovNro, psProveedorCod, "1"
    'oMov.InsertaMovDoc lnMovNro, pnDocTpo, psDocNro, Left(lsMovNro, 8)
    'oMov.InsertaMovComprobanteReg lnMovNro, psAreaAgeCod, pnMoneda, pnDocOrigenTpo, psObservacion, pnTpoPago, psIFICod, psIFICtaCod, pnDocTpo, psDocNro, Left(lsMovNro, 8)
    oMov.InsertaMovComprobanteReg lnMovNro, psAreaAgeCod, pnMoneda, pnDocOrigenTpo, psObservacion, pnTpoPago, psIFICod, psIFICtaCod, pnDocTpo, psDocNro, Format(psFechaDoc, "yyyyMMdd") 'PASI20150107
    oMov.InsertaMovComprobanteContrato lnMovNro, psncontrato, pncontref
    
    For lnMovItem = 1 To UBound(pMatDetContrato)
        oMov.InsertaMovComprobanteContratoCronDet lnMovNro, lnMovItem, pMatDetContrato(lnMovItem).nNPago
        oMov.ActualizaEstadoCuotaCronograma psncontrato, pncontref, pMatDetContrato(lnMovItem).nNPago, 2
    Next lnMovItem
    For lnMovItem = 1 To UBound(pMatDetOrden)
        lnMonto = lnMonto + pMatDetOrden(lnMovItem).nTotal
        oMov.InsertaMovCta lnMovNro, lnMovItem, pMatDetOrden(lnMovItem).sCtaContCod, pMatDetOrden(lnMovItem).nTotal
        oMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjBienesServicios
        oMov.InsertaMovBS lnMovNro, lnMovItem, 1, pMatDetOrden(lnMovItem).sObjeto
        oMov.InsertaMovCant lnMovNro, lnMovItem, pMatDetOrden(lnMovItem).nCantidad
        oMov.InsertaMovComprobanteItem lnMovNro, lnMovItem, pMatDetOrden(lnMovItem).sDescripcion
    Next lnMovItem
    oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContCod, lnMonto * -1
    If pnMoneda = gMonedaExtranjera Then
        oMov.GeneraMovME lnMovNro, lsMovNro, pnTpoCambio, True
    End If
    oMov.ActualizaSaldoContrato psncontrato, pncontref, lnMonto * -1
    oMov.CommitTrans
    GrabaComprobanteContratoxServicio = lnMovNro
    psMovNro = lsMovNro
    bTrans = False
    Set oMov = Nothing
    Exit Function
ErrorGrabar:
    If bTrans Then
        oMov.RollBackTrans
        Set oMov = Nothing
    End If
End Function
Public Function GrabaComprobanteContratoxCronograma(ByVal pdFecha As Date, _
                                         ByVal psAgeCod As String, _
                                         ByVal psUserCod As String, _
                                         ByVal psOpeCod As String, _
                                         ByVal psncontrato As String, _
                                         ByVal pncontref As Integer, _
                                         ByVal psGlosa As String, _
                                         ByVal psProveedorCod As String, _
                                         ByVal psAreaAgeCod As String, _
                                         ByVal pnMoneda As Integer, _
                                         ByVal pnDocOrigenTpo As Integer, _
                                         ByVal psObservacion As String, _
                                         ByVal pnDocTpo As Integer, _
                                         ByVal psDocNro As String, _
                                         ByVal psFechaDoc As Date, _
                                         ByVal pnTpoPago As Date, _
                                         ByVal psIFICod As String, _
                                         ByVal psIFICtaCod As String, _
                                         ByRef pMatDetContrato() As TComprobanteContratoxCronograma, _
                                         ByVal psCtaContCod As String, _
                                         ByVal pnTpoCambio As Currency, _
                                         ByRef psMovNro As String) 'PASI20150107 agrego psFechaDoc
    Dim oMov As New DMov
    Dim bTrans As Boolean
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lnMovItem As Integer
    Dim lnMonto As Currency
    
    On Error GoTo ErrorGrabar
    
    oMov.BeginTrans
    bTrans = True
    
    lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUserCod)
    oMov.InsertaMov lsMovNro, psOpeCod, psGlosa, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(lsMovNro)
    oMov.InsertaMovGasto lnMovNro, psProveedorCod, "1"
    'oMov.InsertaMovDoc lnMovNro, pnDocTpo, psDocNro, Left(lsMovNro, 8)
    'oMov.InsertaMovComprobanteReg lnMovNro, psAreaAgeCod, pnMoneda, pnDocOrigenTpo, psObservacion, pnTpoPago, psIFICod, psIFICtaCod, pnDocTpo, psDocNro, Left(lsMovNro, 8)
    oMov.InsertaMovComprobanteReg lnMovNro, psAreaAgeCod, pnMoneda, pnDocOrigenTpo, psObservacion, pnTpoPago, psIFICod, psIFICtaCod, pnDocTpo, psDocNro, Format(psFechaDoc, "yyyyMMdd") 'PASI20150107
    oMov.InsertaMovComprobanteContrato lnMovNro, psncontrato, pncontref
    
    For lnMovItem = 1 To UBound(pMatDetContrato)
        lnMonto = lnMonto + pMatDetContrato(lnMovItem).nMonto
        oMov.InsertaMovCta lnMovNro, lnMovItem, pMatDetContrato(lnMovItem).sCtaContCod, pMatDetContrato(lnMovItem).nMonto
        oMov.InsertaMovComprobanteItem lnMovNro, lnMovItem, pMatDetContrato(lnMovItem).sDescripcion
        oMov.InsertaMovComprobanteContratoCronDet lnMovNro, lnMovItem, pMatDetContrato(lnMovItem).nNPago
        oMov.ActualizaEstadoCuotaCronograma psncontrato, pncontref, pMatDetContrato(lnMovItem).nNPago, 2
    Next lnMovItem
    oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContCod, lnMonto * -1
    If pnMoneda = gMonedaExtranjera Then
        oMov.GeneraMovME lnMovNro, lsMovNro, pnTpoCambio, True
    End If
    oMov.ActualizaSaldoContrato psncontrato, pncontref, lnMonto * -1
    
    oMov.CommitTrans
    GrabaComprobanteContratoxCronograma = lnMovNro
    psMovNro = lsMovNro
    bTrans = False
    
    Set oMov = Nothing
    Exit Function
ErrorGrabar:
    If bTrans Then
        oMov.RollBackTrans
        Set oMov = Nothing
    End If
End Function
Public Function GrabaComprobanteContratoxBienes(ByVal pdFecha As Date, _
                                         ByVal psAgeCod As String, _
                                         ByVal psUserCod As String, _
                                         ByVal psOpeCod As String, _
                                         ByVal psncontrato As String, _
                                         ByVal pncontref As Integer, _
                                         ByVal psGlosa As String, _
                                         ByVal psProveedorCod As String, _
                                         ByVal psAreaAgeCod As String, _
                                         ByVal pnMoneda As Integer, _
                                         ByVal pnDocOrigenTpo As Integer, _
                                         ByVal psObservacion As String, _
                                         ByVal pnDocTpo As Integer, _
                                         ByVal psDocNro As String, _
                                         ByVal psFechaDoc As Date, _
                                         ByVal pnTpoPago As Date, _
                                         ByVal psIFICod As String, _
                                         ByVal psIFICtaCod As String, _
                                         ByRef pMatDetContrato() As TComprobanteContratoxBien, _
                                         ByVal psCtaContCod As String, _
                                         ByVal pnTpoCambio As Currency, _
                                         ByRef psMovNro As String) 'PASI20150107 agrego psFechaDoc
    
    Dim oMov As New DMov
    Dim bTrans As Boolean
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lnMovItem As Integer
    Dim lnMonto As Currency
    
    On Error GoTo ErrorGrabar
    
    oMov.BeginTrans
    bTrans = True
    
    lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUserCod)
    oMov.InsertaMov lsMovNro, psOpeCod, psGlosa, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(lsMovNro)
    oMov.InsertaMovGasto lnMovNro, psProveedorCod, "1"
    'oMov.InsertaMovDoc lnMovNro, pnDocTpo, psDocNro, Left(lsMovNro, 8)
    'oMov.InsertaMovComprobanteReg lnMovNro, psAreaAgeCod, pnMoneda, pnDocOrigenTpo, psObservacion, pnTpoPago, psIFICod, psIFICtaCod, pnDocTpo, psDocNro, Left(lsMovNro, 8)
    oMov.InsertaMovComprobanteReg lnMovNro, psAreaAgeCod, pnMoneda, pnDocOrigenTpo, psObservacion, pnTpoPago, psIFICod, psIFICtaCod, pnDocTpo, psDocNro, Format(psFechaDoc, "yyyyMMdd") 'PASI20150107
    oMov.InsertaMovComprobanteContrato lnMovNro, psncontrato, pncontref
    For lnMovItem = 1 To UBound(pMatDetContrato)
        lnMonto = lnMonto + pMatDetContrato(lnMovItem).nTotal
        oMov.InsertaMovCta lnMovNro, lnMovItem, pMatDetContrato(lnMovItem).sCtaContCod, pMatDetContrato(lnMovItem).nTotal
        oMov.InsertaMovComprobanteItem lnMovNro, lnMovItem, pMatDetContrato(lnMovItem).sDescripcion
        oMov.InsertaMovComprobanteContratoBienDet lnMovNro, lnMovItem, pMatDetContrato(lnMovItem).nMovItem
        'Ultimo Agregado
            oMov.InsertaMovObj lnMovNro, lnMovItem, 1, ObjBienesServicios
            oMov.InsertaMovBS lnMovNro, lnMovItem, 1, pMatDetContrato(lnMovItem).sObjeto
        'end
        oMov.InsertaMovCant lnMovNro, lnMovItem, pMatDetContrato(lnMovItem).nCantidad
        oMov.ActualizaEstadoBienContrato psncontrato, pncontref, pMatDetContrato(lnMovItem).nMovItem, 3
    Next lnMovItem
    oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContCod, lnMonto * -1
    If pnMoneda = gMonedaExtranjera Then
        oMov.GeneraMovME lnMovNro, lsMovNro, pnTpoCambio, True
    End If
    oMov.ActualizaSaldoContrato psncontrato, pncontref, lnMonto * -1
    oMov.CommitTrans
    GrabaComprobanteContratoxBienes = lnMovNro
    psMovNro = lsMovNro
    bTrans = False
    
    Set oMov = Nothing
    Exit Function
ErrorGrabar:
    If bTrans Then
        oMov.RollBackTrans
        Set oMov = Nothing
    End If
End Function
Public Function GrabaComprobanteContratoxObra(ByVal pdFecha As Date, _
                                         ByVal psAgeCod As String, _
                                         ByVal psUserCod As String, _
                                         ByVal psOpeCod As String, _
                                         ByVal psncontrato As String, _
                                         ByVal pncontref As Integer, _
                                         ByVal psGlosa As String, _
                                         ByVal psProveedorCod As String, _
                                         ByVal psAreaAgeCod As String, _
                                         ByVal pnMoneda As Integer, _
                                         ByVal pnDocOrigenTpo As Integer, _
                                         ByVal psObservacion As String, _
                                         ByVal pnDocTpo As Integer, _
                                         ByVal psDocNro As String, _
                                         ByVal psFechaDoc As Date, _
                                         ByVal pnTpoPago As Date, _
                                         ByVal psIFICod As String, _
                                         ByVal psIFICtaCod As String, _
                                         ByRef pMatDetContrato() As TComprobanteContratoxObra, _
                                         ByVal psCtaContCod As String, _
                                         ByVal pnTpoCambio As Currency, _
                                         ByRef psMovNro As String) 'PASI20150107 agrego psFechaDoc
        
    Dim oMov As New DMov
    Dim bTrans As Boolean
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lnMovItem As Integer
    Dim lnMonto As Currency
    
    On Error GoTo ErrorGrabar
    
    oMov.BeginTrans
    bTrans = True
    
    lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUserCod)
    oMov.InsertaMov lsMovNro, psOpeCod, psGlosa, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(lsMovNro)
    oMov.InsertaMovGasto lnMovNro, psProveedorCod, "1"
    'oMov.InsertaMovDoc lnMovNro, pnDocTpo, psDocNro, Left(lsMovNro, 8)
    'oMov.InsertaMovComprobanteReg lnMovNro, psAreaAgeCod, pnMoneda, pnDocOrigenTpo, psObservacion, pnTpoPago, psIFICod, psIFICtaCod, pnDocTpo, psDocNro, Left(lsMovNro, 8)
    oMov.InsertaMovComprobanteReg lnMovNro, psAreaAgeCod, pnMoneda, pnDocOrigenTpo, psObservacion, pnTpoPago, psIFICod, psIFICtaCod, pnDocTpo, psDocNro, Format(psFechaDoc, "yyyyMMdd") 'PASI20150107
    oMov.InsertaMovComprobanteContrato lnMovNro, psncontrato, pncontref
    For lnMovItem = 1 To UBound(pMatDetContrato)
        lnMonto = lnMonto + pMatDetContrato(lnMovItem).nMonto
        oMov.InsertaMovCta lnMovNro, lnMovItem, pMatDetContrato(lnMovItem).sCtaContCod, pMatDetContrato(lnMovItem).nMonto
        oMov.InsertaMovComprobanteItem lnMovNro, lnMovItem, pMatDetContrato(lnMovItem).sDescripcion
    Next lnMovItem
    oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaContCod, lnMonto * -1
    If pnMoneda = gMonedaExtranjera Then
        oMov.GeneraMovME lnMovNro, lsMovNro, pnTpoCambio, True
    End If
    oMov.ActualizaSaldoContrato psncontrato, pncontref, lnMonto * -1
    oMov.CommitTrans
    GrabaComprobanteContratoxObra = lnMovNro
    psMovNro = lsMovNro
    bTrans = True
    Set oMov = Nothing
    Exit Function
ErrorGrabar:
    If bTrans Then
        oMov.RollBackTrans
        Set oMov = Nothing
    End If
End Function
Public Function GrabarActaConformidadNew(ByVal pdFecha As Date, ByVal psAgeCod As String, _
                                        ByVal psUserCod As String, ByVal psOpeCod As String, _
                                        ByVal pnDocOrigenTpo As Integer, ByVal pnDocTpo As Integer, _
                                        ByVal psDocReferencia As String, ByVal psAreaAgeCod As String, _
                                        ByVal psSubAreaDesc As String, ByVal pnMoneda As Moneda, _
                                        ByRef psActaNro As String, _
                                        ByVal psDescripcion As String, ByVal psObservacion As String, _
                                        ByRef psMovNro As String, _
                                        Optional ByVal psNGuia As String = "N/A", _
                                        Optional ByVal pnMovNroComp As Long = 0) As Long
   Dim oMov As New DMov
   Dim bTrans As Boolean
   Dim lsMovNro As String
   Dim lnMovNro As Long
   Dim rsNActa As ADODB.Recordset
   On Error GoTo ErrGrabar
   
   oMov.BeginTrans
    bTrans = True
    
    psActaNro = oMov.GetExisteNActaConformidadxContrato(pnMovNroComp)
    If psActaNro = "" Then
        psActaNro = oMov.GetCorrelativoActaConformidad(pnDocTpo, psAgeCod, CStr(Year(pdFecha)))
    End If
    
    'psActaNro = oMov.GetCorrelativoActaConformidad(pnDocTpo, psAgeCod, CStr(Year(pdFecha))) 'Se genera un nuevo correlativo para evitar duplicidad 'Comentado PASIERS0772014
    
    If Len(Trim(psActaNro)) = 0 Then
       oMov.RollBackTrans
       Set oMov = Nothing
       Exit Function
    End If
    lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUserCod)
    oMov.InsertaMov lsMovNro, psOpeCod, Left(psDescripcion, 250), gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(lsMovNro)
    oMov.InsertaMovDoc lnMovNro, pnDocTpo, psActaNro, Left(lsMovNro, 8)
    oMov.InsertaMovActaConformidadReg lnMovNro, psAreaAgeCod, psSubAreaDesc, pnMoneda, pnDocOrigenTpo, psDocReferencia, psDescripcion, psObservacion, psNGuia
    oMov.InsertaMovRef lnMovNro, pnMovNroComp
    oMov.ActualizaActaPendComprobante pnMovNroComp, psAreaAgeCod, lsMovNro, psActaNro
    oMov.CommitTrans
    GrabarActaConformidadNew = lnMovNro
    psMovNro = lsMovNro
    bTrans = False
    Set oMov = Nothing
Exit Function
ErrGrabar:
    If bTrans Then
        oMov.RollBackTrans
        Set oMov = Nothing
    End If
End Function
Public Function ExtornaComprobanteNew(ByVal pdFecha As Date, ByVal psAgeCod As String, ByVal psUserCod As String, ByVal psOpeCod As String, ByVal psGlosa As String, ByVal pnMovNroRef As Long, ByVal pnMonto As Currency, Optional ByVal pncontref = 0, Optional ByVal psncontrato As String = "") As Boolean
    Dim oMov As New DMov
    Dim bTrans As Boolean
    Dim lsMovNro As String
    Dim lnMovNro As Long
    
    On Error GoTo ErrExtornaComprobante
     oMov.BeginTrans
    bTrans = True
    lsMovNro = oMov.GeneraMovNro(pdFecha, psAgeCod, psUserCod)
    oMov.InsertaMov lsMovNro, psOpeCod, Left(psGlosa, 300), gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(lsMovNro)
    oMov.EliminaMov pnMovNroRef
    oMov.InsertaMovRef lnMovNro, pnMovNroRef
    oMov.ActualizaContratoxExtActaConformidad pnMovNroRef
    oMov.ActualizaSaldoContrato psncontrato, pncontref, pnMonto
    
    oMov.CommitTrans
    bTrans = False
    ExtornaComprobanteNew = True
    Set oMov = Nothing
    Exit Function
ErrExtornaComprobante:
    If bTrans Then
        oMov.RollBackTrans
        Set oMov = Nothing
        ExtornaComprobanteNew = False
    End If
End Function
'end PASI
