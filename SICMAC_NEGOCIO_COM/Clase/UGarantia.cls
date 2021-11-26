VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Atributos
Private sGarantID As String
Private nmoneda As Moneda
Private nClasificacion As Integer
Private nBienGarantia As Integer
Private sClasificacionSBS As String
Private sClasificacionInterna As String
Private sEmisorCod As String
Private sEmisorNombre As String
Private nDocumentoTpo As Integer
Private sDocumentoNro As String
Private dFechaConstata As Date
Private bPermiteNuevaValorizacion As Boolean
Private bPermiteNuevoTramiteLegal As Boolean
Private bVinculadoUltVAL As Boolean
Private bBloqueoLegal As Boolean
Private nNroCredVinc As Integer
Private nNroCredVig As Integer
Private nNroCredEnProceso As Integer 'EJVG20160420
Private bDJBloqueada As Boolean
Private nEstado As Integer
'***JGPA TI-ERS026-2018
Private nNroCredCancelados As Integer
Private nNroCredEnProceso2 As Integer
Private nNroCredVig2 As Integer
'**End JGPA

Private Propietarios() As tPropietarioRelacion
Private Valorizaciones() As tValorizacion
Private TramitesLegales() As tTramiteLegal
Private nTpoBienContrato As Integer 'CTI5 ERS012020
Private sEtapa As String 'CTI5 ERS012020
Private sCtaCodDesembolsado As String 'CTI5 ERS012020
Private bCambiadoBienFuturo As Boolean 'CTI5 ERS012020
Private dFechaSistema As Date 'CTI5 ERS012020
Private sAgeCtaDesembolsado As String 'CTI5 ERS012021
Private nSalaValida As Integer 'JOEP 20210930

'Propiedades
Public Property Get GarantID() As String
    GarantID = sGarantID
End Property
Public Property Let GarantID(ByVal vNewValue As String)
    sGarantID = vNewValue
End Property
Public Property Get Moneda() As Moneda
    Moneda = nmoneda
End Property
Public Property Let Moneda(ByVal vNewValue As Moneda)
    nmoneda = vNewValue
End Property
Public Property Get Clasificacion() As Integer
    Clasificacion = nClasificacion
End Property
Public Property Let Clasificacion(ByVal vNewValue As Integer)
    nClasificacion = vNewValue
End Property
Public Property Get BienGarantia() As Integer
    BienGarantia = nBienGarantia
End Property
Public Property Let BienGarantia(ByVal vNewValue As Integer)
    nBienGarantia = vNewValue
End Property
Public Property Get ClasificacionSBS() As String
    ClasificacionSBS = sClasificacionSBS
End Property
Public Property Let ClasificacionSBS(ByVal vNewValue As String)
    sClasificacionSBS = vNewValue
End Property
Public Property Get ClasificacionInterna() As String
    ClasificacionInterna = sClasificacionInterna
End Property
Public Property Let ClasificacionInterna(ByVal vNewValue As String)
    sClasificacionInterna = vNewValue
End Property
Public Property Get EmisorCod() As String
    EmisorCod = sEmisorCod
End Property
Public Property Let EmisorCod(ByVal vNewValue As String)
    sEmisorCod = vNewValue
End Property
Public Property Get EmisorNombre() As String
    EmisorNombre = sEmisorNombre
End Property
Public Property Let EmisorNombre(ByVal vNewValue As String)
    sEmisorNombre = vNewValue
End Property
Public Property Get DocumentoTpo() As Integer
    DocumentoTpo = nDocumentoTpo
End Property
Public Property Let DocumentoTpo(ByVal vNewValue As Integer)
    nDocumentoTpo = vNewValue
End Property
Public Property Get DocumentoNro() As String
    DocumentoNro = sDocumentoNro
End Property
Public Property Let DocumentoNro(ByVal vNewValue As String)
    sDocumentoNro = vNewValue
End Property
Public Property Get FechaConstata() As Date
    FechaConstata = dFechaConstata
End Property
Public Property Let FechaConstata(ByVal vNewValue As Date)
    dFechaConstata = vNewValue
End Property
Public Property Get PermiteNuevaValorizacion() As Boolean
    PermiteNuevaValorizacion = bPermiteNuevaValorizacion
End Property
Public Property Let PermiteNuevaValorizacion(ByVal vNewValue As Boolean)
    bPermiteNuevaValorizacion = vNewValue
End Property
Public Property Get PermiteNuevoTramiteLegal() As Boolean
    PermiteNuevoTramiteLegal = bPermiteNuevoTramiteLegal
End Property
Public Property Let PermiteNuevoTramiteLegal(ByVal vNewValue As Boolean)
    bPermiteNuevoTramiteLegal = vNewValue
End Property
Public Property Get VinculadoUltVAL() As Boolean
    VinculadoUltVAL = bVinculadoUltVAL
End Property
Public Property Let VinculadoUltVAL(ByVal vNewValue As Boolean)
    bVinculadoUltVAL = vNewValue
End Property
Public Property Get BloqueoLegal() As Boolean
    BloqueoLegal = bBloqueoLegal
End Property
Public Property Let BloqueoLegal(ByVal vNewValue As Boolean)
    bBloqueoLegal = vNewValue
End Property
Public Property Get NroCredVinc() As Integer
    NroCredVinc = nNroCredVinc
End Property
Public Property Let NroCredVinc(ByVal vNewValue As Integer)
    nNroCredVinc = vNewValue
End Property
Public Property Get NroCredVig() As Integer
    NroCredVig = nNroCredVig
End Property
Public Property Let NroCredVig(ByVal vNewValue As Integer)
    nNroCredVig = vNewValue
End Property
Public Property Get NroCredEnProceso() As Integer 'EJVG20160420
    NroCredEnProceso = nNroCredEnProceso
End Property
Public Property Let NroCredEnProceso(ByVal vNewValue As Integer) 'EJVG20160420
    nNroCredEnProceso = vNewValue
End Property
Public Property Get DJBloqueada() As Boolean
    DJBloqueada = bDJBloqueada
End Property
Public Property Let DJBloqueada(ByVal vNewValue As Boolean)
    bDJBloqueada = vNewValue
End Property
Public Property Get Estado() As Integer
    Estado = nEstado
End Property
Public Property Let Estado(ByVal vNewValue As Integer)
    nEstado = vNewValue
End Property

Public Property Get NroPropietarios() As Integer
    NroPropietarios = UBound(Propietarios)
End Property
'***JGPA TI-ERS026-2018-------------------------------------------
Public Property Get NroCredCancelados() As Integer
    NroCredCancelados = nNroCredCancelados
End Property
Public Property Let NroCredCancelados(ByVal vNewValue As Integer)
    nNroCredCancelados = vNewValue
End Property
Public Property Get NroCredEnProceso2() As Integer
    NroCredEnProceso2 = nNroCredEnProceso2
End Property
Public Property Let NroCredEnProceso2(ByVal vNewValue As Integer)
    nNroCredEnProceso2 = vNewValue
End Property
Public Property Get NroCredVig2() As Integer
    NroCredVig2 = nNroCredVig2
End Property
Public Property Let NroCredVig2(ByVal vNewValue As Integer)
    nNroCredVig2 = vNewValue
End Property
'**End JGPA----------------------------------------------------

'JOEP20210820
Public Property Let SalaValida(ByVal vNewValue As Integer)
    nSalaValida = vNewValue
End Property

Public Property Get SalaValida() As Integer
    SalaValida = nSalaValida
End Property
'JOEP20210820
Public Property Get NroPropietariosActivos() As Integer
    Dim Index As Integer
    Dim Cuenta As Integer
    For Index = 1 To NroPropietarios
        If ObtenerPropietario(Index).nGarantiaCambiosFila <> GarantFilaEliminada And ObtenerPropietario(Index).nGarantiaCambiosFila <> GarantFilaOculta Then
            Cuenta = Cuenta + 1
        End If
    Next
    NroPropietariosActivos = Cuenta
End Property
Public Property Get NroValorizaciones() As Integer
    NroValorizaciones = UBound(Valorizaciones)
End Property
Public Property Get NroValorizacionesActivas() As Integer
    Dim Index As Integer
    Dim Cuenta As Integer
    For Index = 1 To NroValorizaciones
        If ObtenerValorizacion(Index).nGarantiaCambiosFila <> GarantFilaEliminada And ObtenerValorizacion(Index).nGarantiaCambiosFila <> GarantFilaOculta Then
            Cuenta = Cuenta + 1
        End If
    Next
    NroValorizacionesActivas = Cuenta
End Property
Public Property Get IndexUltimaValorizacionActiva() As Integer
    Dim Index As Integer
    Dim IndexVAL As Integer
    For Index = 1 To NroValorizaciones
        If ObtenerValorizacion(Index).nGarantiaCambiosFila <> GarantFilaEliminada And ObtenerValorizacion(Index).nGarantiaCambiosFila <> GarantFilaOculta Then
            IndexVAL = Index
        End If
    Next
    IndexUltimaValorizacionActiva = IndexVAL
End Property
Public Property Get NroTramitesLegal() As Integer
    NroTramitesLegal = UBound(TramitesLegales)
End Property
Public Property Get NroTramitesLegalActivas() As Integer
    Dim Index As Integer
    Dim Cuenta As Integer
    For Index = 1 To NroTramitesLegal
        If ObtenerTramiteLegal(Index).nGarantiaCambiosFila <> GarantFilaEliminada And ObtenerTramiteLegal(Index).nGarantiaCambiosFila <> GarantFilaOculta Then
            Cuenta = Cuenta + 1
        End If
    Next
    NroTramitesLegalActivas = Cuenta
End Property
Public Property Get IndexUltimoTramiteLegalActiva() As Integer
    Dim Index As Integer
    Dim IndexTRA As Integer
    For Index = 1 To NroTramitesLegal
        If ObtenerTramiteLegal(Index).nGarantiaCambiosFila <> GarantFilaEliminada And ObtenerTramiteLegal(Index).nGarantiaCambiosFila <> GarantFilaOculta Then
            IndexTRA = Index
        End If
    Next
    IndexUltimoTramiteLegalActiva = IndexTRA
End Property
'CTI5 ERS0032020****************************
Public Property Let TpoBienContrato(ByVal vNewValue As Integer)
    nTpoBienContrato = vNewValue
End Property

Public Property Get TpoBienContrato() As Integer
    TpoBienContrato = nTpoBienContrato
End Property

Public Property Let Etapa(ByVal vNewValue As String)
    sEtapa = vNewValue
End Property

Public Property Get Etapa() As String
    Etapa = sEtapa
End Property

Public Property Let CtaCodDesembolsado(ByVal vNewValue As String)
    sCtaCodDesembolsado = vNewValue
End Property

Public Property Get CtaCodDesembolsado() As String
    CtaCodDesembolsado = sCtaCodDesembolsado
End Property

Public Property Let CambiadoBienFuturo(ByVal vNewValue As Boolean)
    bCambiadoBienFuturo = vNewValue
End Property

Public Property Get CambiadoBienFuturo() As Boolean
    CambiadoBienFuturo = bCambiadoBienFuturo
End Property

Public Property Let FechaSistema(ByVal vNewValue As Date)
    dFechaSistema = vNewValue
End Property

Public Property Get FechaSistema() As Date
    FechaSistema = dFechaSistema
End Property

Public Property Let AgeCtaDesembolsado(ByVal vNewValue As String)
    sAgeCtaDesembolsado = vNewValue
End Property

Public Property Get AgeCtaDesembolsado() As String
    AgeCtaDesembolsado = sAgeCtaDesembolsado
End Property

'*******************************************
'Métodos
Private Sub Class_Initialize()
    NuevaGarantia
End Sub
Private Sub NuevaGarantia()
    GarantID = ""
    Moneda = 0
    Clasificacion = 0
    BienGarantia = 0
    ClasificacionSBS = ""
    ClasificacionInterna = ""
    EmisorCod = ""
    EmisorNombre = ""
    DocumentoTpo = 0
    DocumentoNro = ""
    FechaConstata = CDate("01/01/1900")
    PermiteNuevaValorizacion = True 'Cuando es nuevo podrá agregar valorizacion y trámite legal
    PermiteNuevoTramiteLegal = True
    VinculadoUltVAL = False
    BloqueoLegal = False
    NroCredVinc = 0
    NroCredVig = 0
    NroCredEnProceso = 0
    ReDim Propietarios(0)
    ReDim Valorizaciones(0)
    ReDim TramitesLegales(0)
    TpoBienContrato = 0 'CTI05 ERS012020
End Sub
Public Sub InsertarPropietario(ByVal psPersCod As String, ByVal psNombre As String, ByVal psRelacionTpo As String, ByVal pnCambiosFila As eGarantiaCambiosFila)
    Dim Index As Integer
    Index = UBound(Propietarios) + 1
    ReDim Preserve Propietarios(Index)
    Propietarios(Index).sPersCod = psPersCod
    Propietarios(Index).sPersNombre = psNombre
    Propietarios(Index).nRelacionTpo = val(Trim(Right(psRelacionTpo, 3)))
    Propietarios(Index).sRelacionTpo = psRelacionTpo
    Propietarios(Index).nGarantiaCambiosFila = pnCambiosFila
End Sub
Public Sub ActualizarPropietario(ByVal pnFila As Integer, ByVal psPersCod As String, ByVal psNombre As String, ByVal psRelacionTpo As String)
    Propietarios(pnFila).sPersCod = psPersCod
    Propietarios(pnFila).sPersNombre = psNombre
    Propietarios(pnFila).nRelacionTpo = val(Trim(Right(psRelacionTpo, 3)))
    Propietarios(pnFila).sRelacionTpo = psRelacionTpo
    Propietarios(pnFila).nGarantiaCambiosFila = IIf(Propietarios(pnFila).nGarantiaCambiosFila = GarantFilaNueva, GarantFilaNueva, GarantFilaModificada)
End Sub
Public Sub EliminarPropietario(ByVal pnFila As Integer)
    Dim PropietariosTmp() As tPropietarioRelacion
    Dim indexTmp As Integer, Index As Integer
    
    If Propietarios(pnFila).nGarantiaCambiosFila = GarantFilaNueva Then 'Si es nuevo lo elimina de memoria
        ReDim PropietariosTmp(0)
        For Index = 1 To UBound(Propietarios)
            ReDim Preserve PropietariosTmp(Index)
            PropietariosTmp(Index) = Propietarios(Index)
        Next
        ReDim Propietarios(0)
        Index = 0
        For indexTmp = 1 To UBound(PropietariosTmp)
            If indexTmp <> pnFila Then
                Index = Index + 1
                ReDim Preserve Propietarios(Index)
                Propietarios(Index) = PropietariosTmp(indexTmp)
            End If
        Next
        Erase PropietariosTmp
    Else 'Caso contrario lo cambia de estado para eliminarlo cuando se grabe
        CambiarEstadoPropietario pnFila, GarantFilaEliminada
    End If
End Sub
Private Sub CambiarEstadoPropietario(ByVal pnFila As Integer, ByVal pnCambio As eGarantiaCambiosFila)
    Propietarios(pnFila).nGarantiaCambiosFila = pnCambio
End Sub
Public Function ObtenerPropietario(ByVal pnFila As Integer) As tPropietarioRelacion
    ObtenerPropietario = Propietarios(pnFila)
End Function
Public Function ObtenerPropietarioEspecifico(ByVal pnInterviniente As eTipoRelacion) As tPropietarioRelacion
    Dim Index As Integer
    Dim i As Integer

    For i = 1 To UBound(Propietarios)
        If Propietarios(i).nRelacionTpo = pnInterviniente Then
            Index = i
            Exit For
        End If
    Next
    ObtenerPropietarioEspecifico = Propietarios(Index)
End Function
Public Sub InsertarValorizacion(ByVal pdFecha As Date, ByVal psUltimaActualizacion As String, _
                                ByVal pnValorizacionTpo As eGarantiaTipoValorizacion, ByVal psValorizacionTpo As String, _
                                ByVal psGlosa As String, ByVal pnCambiosFila As eGarantiaCambiosFila, ByVal pbVinculado As Boolean, _
                                Optional ByRef pvValDirTot As Variant = Nothing, _
                                Optional ByRef pvDetalle As Variant = Nothing, _
                                Optional ByVal pvValorAutoLiquidable As Variant = Nothing, _
                                Optional ByVal pvValorTasacionInmobiliaria As Variant = Nothing, _
                                Optional ByVal pvValorTasacionVehicular As Variant = Nothing, _
                                Optional ByVal pvValorTasacionMobiliariaOtras As Variant = Nothing, _
                                Optional ByVal pbEliminar As Boolean = False, _
                                Optional ByVal pvValorGravamenFavorOtraIFi As Variant = Nothing)
    Dim lvValDirTot As tValorDirectoTotal
    Dim lvValDirDet As tValorDirectoDetallado
    Dim lvValAutoliquidable As tValorAutoLiquidable
    Dim lvValorTasacionInmobiliaria As tValorTasacionInmobiliaria
    Dim lvValorTasacionVehicular As tValorTasacionVehicular
    Dim lvValorTasacionMobiliariaOtras As tValorTasacionMobiliariaOtras
    Dim lvValorGravamenFavorOtraIFi As tValorGravamenFavorOtraIFi 'EJVG20160225
    
    Dim Index As Integer, indexDet As Integer
    Dim lnMonto As Currency
    
    Index = UBound(Valorizaciones) + 1
    ReDim Preserve Valorizaciones(Index)
    
    Valorizaciones(Index).dfecha = pdFecha
    Valorizaciones(Index).nValorizacionTpo = pnValorizacionTpo
    Valorizaciones(Index).sValorizacionTpo = psValorizacionTpo
    Valorizaciones(Index).sGlosa = psGlosa
    Valorizaciones(Index).nGarantiaCambiosFila = pnCambiosFila
    Valorizaciones(Index).bVinculado = pbVinculado
    Valorizaciones(Index).sUltimaActualizacion = psUltimaActualizacion
    Valorizaciones(Index).bEliminar = pbEliminar
    
    ReDim Valorizaciones(Index).vValorDirDet(0)
    
    Select Case pnValorizacionTpo
        Case eGarantiaTipoValorizacion.ValorDirectoTotal
            lvValDirTot = pvValDirTot
            Valorizaciones(Index).vValorDirectoTotal = lvValDirTot
            Valorizaciones(Index).nVRM = lvValDirTot.nVRM
        Case eGarantiaTipoValorizacion.ValorDirectoDetallado
            For indexDet = 1 To UBound(pvDetalle)
                lvValDirDet = pvDetalle(indexDet)
                ReDim Preserve Valorizaciones(Index).vValorDirDet(indexDet)
                Valorizaciones(Index).vValorDirDet(indexDet) = lvValDirDet
                lnMonto = lnMonto + lvValDirDet.nValor * lvValDirDet.nCantidad
            Next
            Valorizaciones(Index).nVRM = lnMonto
        Case eGarantiaTipoValorizacion.GarantiaAutoliquidable
            lvValAutoliquidable = pvValorAutoLiquidable
            Valorizaciones(Index).vValorAutoLiquidable = lvValAutoliquidable
            Valorizaciones(Index).nVRM = lvValAutoliquidable.nMontoCuentaDisp
        Case eGarantiaTipoValorizacion.TasacionInmobiliaria
            lvValorTasacionInmobiliaria = pvValorTasacionInmobiliaria
            Valorizaciones(Index).vValorTasacionInmobiliaria = lvValorTasacionInmobiliaria
            Valorizaciones(Index).nVRM = lvValorTasacionInmobiliaria.nVRM
        Case eGarantiaTipoValorizacion.TasacionVehicular
            lvValorTasacionVehicular = pvValorTasacionVehicular
            Valorizaciones(Index).vValorTasacionVehicular = lvValorTasacionVehicular
            Valorizaciones(Index).nVRM = lvValorTasacionVehicular.nVRM
        Case eGarantiaTipoValorizacion.OtrasTasacionesMobiliarias
            lvValorTasacionMobiliariaOtras = pvValorTasacionMobiliariaOtras
            Valorizaciones(Index).vValorTasacionMobiliariaOtras = lvValorTasacionMobiliariaOtras
            Valorizaciones(Index).nVRM = lvValorTasacionMobiliariaOtras.nVRM
        Case eGarantiaTipoValorizacion.GravamenFavorOtraIFi 'EJVG20160225
            lvValorGravamenFavorOtraIFi = pvValorGravamenFavorOtraIFi
            Valorizaciones(Index).vValorGravamenFavorOtraIFi = lvValorGravamenFavorOtraIFi
            For indexDet = 1 To UBound(lvValorGravamenFavorOtraIFi.vValorGravamenFavorOtraIFiDet)
                lnMonto = lnMonto + lvValorGravamenFavorOtraIFi.vValorGravamenFavorOtraIFiDet(indexDet).nGravado
            Next
            Valorizaciones(Index).nVRM = lvValorGravamenFavorOtraIFi.nValorComercial - lnMonto
            Valorizaciones(Index).sGlosa = lvValorGravamenFavorOtraIFi.sDescripcion
            Valorizaciones(Index).nVTpRegPub = lvValorGravamenFavorOtraIFi.nOpTpRegPub 'JOEP20180524 ERS009-2018
    End Select
End Sub
Public Sub ActualizarValorizacion(ByVal pnFila As Integer, ByVal psUltimaActualizacion As String, _
                                    Optional ByVal psGlosa As String = "", _
                                    Optional ByRef pvValDirTot As Variant = Nothing, _
                                    Optional ByRef pvDetalle As Variant = Nothing, _
                                    Optional ByVal pvValorAutoLiquidable As Variant = Nothing, _
                                    Optional ByVal pvValorTasacionInmobiliaria As Variant = Nothing, _
                                    Optional ByVal pvValorTasacionVehicular As Variant = Nothing, _
                                    Optional ByVal pvValorTasacionMobiliariaOtras As Variant = Nothing, _
                                    Optional ByVal pvValorGravamenFavorOtraIFi As Variant = Nothing)
    Dim lvValDirTot As tValorDirectoTotal
    Dim lvValDirDet As tValorDirectoDetallado
    Dim lvValAutoliquidable As tValorAutoLiquidable
    Dim lvValorTasacionInmobiliaria As tValorTasacionInmobiliaria
    Dim lvValorTasacionVehicular As tValorTasacionVehicular
    Dim lvValorTasacionMobiliariaOtras As tValorTasacionMobiliariaOtras
    Dim lvValorGravamenFavorOtraIFi As tValorGravamenFavorOtraIFi 'EJVG20160225
    
    Dim lnValorizacionTpo As eGarantiaTipoValorizacion
    Dim Index As Integer, indexDet As Integer
    Dim lnMonto As Currency
    
    Index = pnFila
    lnValorizacionTpo = Valorizaciones(Index).nValorizacionTpo
    If Valorizaciones(Index).nGarantiaCambiosFila = GarantFilaNueva Then
        Valorizaciones(Index).dfecha = fgFechaHoraMovDate(psUltimaActualizacion)
    End If
    Valorizaciones(Index).sGlosa = psGlosa
    If Valorizaciones(Index).nGarantiaCambiosFila <> GarantFilaNueva Then
        Valorizaciones(Index).nGarantiaCambiosFila = GarantFilaModificada
    End If
    Valorizaciones(Index).sUltimaActualizacion = psUltimaActualizacion
    
    ReDim Valorizaciones(Index).vValorDirDet(0)
    
    Select Case lnValorizacionTpo
        Case eGarantiaTipoValorizacion.ValorDirectoTotal
            lvValDirTot = pvValDirTot
            Valorizaciones(Index).vValorDirectoTotal = lvValDirTot
            Valorizaciones(Index).nVRM = lvValDirTot.nVRM
        Case eGarantiaTipoValorizacion.ValorDirectoDetallado
            For indexDet = 1 To UBound(pvDetalle)
                lvValDirDet = pvDetalle(indexDet)
                ReDim Preserve Valorizaciones(Index).vValorDirDet(indexDet)
                Valorizaciones(Index).vValorDirDet(indexDet) = lvValDirDet
                lnMonto = lnMonto + lvValDirDet.nValor * lvValDirDet.nCantidad
            Next
            Valorizaciones(Index).nVRM = lnMonto
        Case eGarantiaTipoValorizacion.GarantiaAutoliquidable
            lvValAutoliquidable = pvValorAutoLiquidable
            Valorizaciones(Index).vValorAutoLiquidable = lvValAutoliquidable
            Valorizaciones(Index).nVRM = lvValAutoliquidable.nMontoCuentaDisp
        Case eGarantiaTipoValorizacion.TasacionInmobiliaria
            lvValorTasacionInmobiliaria = pvValorTasacionInmobiliaria
            Valorizaciones(Index).vValorTasacionInmobiliaria = lvValorTasacionInmobiliaria
            Valorizaciones(Index).nVRM = lvValorTasacionInmobiliaria.nVRM
        Case eGarantiaTipoValorizacion.TasacionVehicular
            lvValorTasacionVehicular = pvValorTasacionVehicular
            Valorizaciones(Index).vValorTasacionVehicular = lvValorTasacionVehicular
            Valorizaciones(Index).nVRM = lvValorTasacionVehicular.nVRM
        Case eGarantiaTipoValorizacion.OtrasTasacionesMobiliarias
            lvValorTasacionMobiliariaOtras = pvValorTasacionMobiliariaOtras
            Valorizaciones(Index).vValorTasacionMobiliariaOtras = lvValorTasacionMobiliariaOtras
            Valorizaciones(Index).nVRM = lvValorTasacionMobiliariaOtras.nVRM
        Case eGarantiaTipoValorizacion.GravamenFavorOtraIFi 'EJVG20160225
            lvValorGravamenFavorOtraIFi = pvValorGravamenFavorOtraIFi
            Valorizaciones(Index).vValorGravamenFavorOtraIFi = lvValorGravamenFavorOtraIFi
            For indexDet = 1 To UBound(lvValorGravamenFavorOtraIFi.vValorGravamenFavorOtraIFiDet)
                lnMonto = lnMonto + lvValorGravamenFavorOtraIFi.vValorGravamenFavorOtraIFiDet(indexDet).nGravado
            Next
            Valorizaciones(Index).nVRM = lvValorGravamenFavorOtraIFi.nValorComercial - lnMonto
            Valorizaciones(Index).sGlosa = lvValorTasacionMobiliariaOtras.sDescripcion
    End Select
End Sub
Public Sub EliminarValorizacion(ByVal pnFila As Integer)
    Dim ValorizacionesTmp() As tValorizacion
    Dim indexTmp As Integer, Index As Integer
    
    If Valorizaciones(pnFila).nGarantiaCambiosFila = GarantFilaNueva Then 'Si es nuevo lo elimina de memoria
        ReDim ValorizacionesTmp(0)
        For Index = 1 To UBound(Valorizaciones)
            ReDim Preserve ValorizacionesTmp(Index)
            ValorizacionesTmp(Index) = Valorizaciones(Index)
        Next
        ReDim Valorizaciones(0)
        Index = 0
        For indexTmp = 1 To UBound(ValorizacionesTmp)
            If indexTmp <> pnFila Then
                Index = Index + 1
                ReDim Preserve Valorizaciones(Index)
                Valorizaciones(Index) = ValorizacionesTmp(indexTmp)
            End If
        Next
        Erase ValorizacionesTmp
    Else 'Caso contrario lo cambia de estado para eliminarlo cuando se grabe
        CambiarEstadoValorizacion pnFila, GarantFilaEliminada
    End If
End Sub
Private Sub CambiarEstadoValorizacion(ByVal pnFila As Integer, ByVal pnCambio As eGarantiaCambiosFila)
    Valorizaciones(pnFila).nGarantiaCambiosFila = pnCambio
End Sub
'Public Sub OcultarValorizaciones()
'    Dim Index As Integer
'
'    For Index = NroValorizaciones To 1 Step -1
'        CambiarEstadoValorizacion Index, GarantFilaOculta
'    Next
'End Sub
Public Function ObtenerValorizacion(ByVal pnFila As Integer) As tValorizacion
    ObtenerValorizacion = Valorizaciones(pnFila)
End Function
Public Function ObtenerValorizacionUltima(ByVal pnTpoValorizacion As eGarantiaTipoValorizacion, Optional ByVal pdFechaMax As Date = CDate("1900-01-01")) As Variant
    Dim Index As Integer
    Dim IndexBusca As Integer
    
    Dim vResultado As Variant
    Dim vTotal As tValorDirectoTotal
    Dim vDet() As tValorDirectoDetallado
    Dim vInmobiliaria As tValorTasacionInmobiliaria
    Dim vVehicular As tValorTasacionVehicular
    Dim vMobiliaria As tValorTasacionMobiliariaOtras
    Dim vGravamenIFi As tValorGravamenFavorOtraIFi 'EJVG20160225
    
    IndexBusca = -1
    
    If pnTpoValorizacion = ValorDirectoTotal Then
        vResultado = vTotal
    ElseIf pnTpoValorizacion = ValorDirectoDetallado Then
        vResultado = vDet
    ElseIf pnTpoValorizacion = TasacionInmobiliaria Then
        vResultado = vInmobiliaria
    ElseIf pnTpoValorizacion = TasacionVehicular Then
        vResultado = vVehicular
    ElseIf pnTpoValorizacion = OtrasTasacionesMobiliarias Then
        vResultado = vMobiliaria
    ElseIf pnTpoValorizacion = GravamenFavorOtraIFi Then 'EJVG20160225
        vResultado = vGravamenIFi
    End If
    
    For Index = 1 To NroValorizaciones
        If ObtenerValorizacion(Index).nGarantiaCambiosFila <> GarantFilaEliminada And ObtenerValorizacion(Index).nGarantiaCambiosFila <> GarantFilaOculta Then
            If DateDiff("D", pdFechaMax, CDate("1900-01-01")) = 0 Or (DateDiff("D", pdFechaMax, CDate("1900-01-01")) <> 0 And ObtenerValorizacion(Index).dfecha < pdFechaMax) Then 'Para sacar según fecha puedan ingresar
                If pnTpoValorizacion = ObtenerValorizacion(Index).nValorizacionTpo Then
                    IndexBusca = Index
                End If
            End If
        End If
    Next
    
    If IndexBusca <> -1 Then
        If pnTpoValorizacion = ValorDirectoTotal Then
            vResultado = ObtenerValorizacion(IndexBusca).vValorDirectoTotal
        ElseIf pnTpoValorizacion = ValorDirectoDetallado Then
            vResultado = ObtenerValorizacion(IndexBusca).vValorDirDet
        ElseIf pnTpoValorizacion = TasacionInmobiliaria Then
            vResultado = ObtenerValorizacion(IndexBusca).vValorTasacionInmobiliaria
        ElseIf pnTpoValorizacion = TasacionVehicular Then
            vResultado = ObtenerValorizacion(IndexBusca).vValorTasacionVehicular
        ElseIf pnTpoValorizacion = OtrasTasacionesMobiliarias Then
            vResultado = ObtenerValorizacion(IndexBusca).vValorTasacionMobiliariaOtras
        ElseIf pnTpoValorizacion = GravamenFavorOtraIFi Then 'EJVG20160225
            vResultado = ObtenerValorizacion(IndexBusca).vValorGravamenFavorOtraIFi
        End If
    End If
    
    ObtenerValorizacionUltima = vResultado
End Function
Public Sub InsertarTramiteLegal(ByVal pvTramiteLegal As Variant, Optional ByVal pnCambiosFila As eGarantiaCambiosFila = GarantFilaNueva, Optional ByVal pbEliminar As Boolean = False)
    Dim Index As Integer
   
    Index = UBound(TramitesLegales) + 1
    ReDim Preserve TramitesLegales(Index)
    
    TramitesLegales(Index) = pvTramiteLegal
    TramitesLegales(Index).nGarantiaCambiosFila = pnCambiosFila
    TramitesLegales(Index).bEliminar = pbEliminar
End Sub
Public Sub ActualizarTramiteLegal(ByVal pnFila As Integer, ByVal pvTramiteLegal As Variant)
    Dim Index As Integer
    Dim ldFecha As Date
    Dim lbVinculado As Boolean
    
    Index = pnFila
    'Recuperamos las llaves principales
    ldFecha = TramitesLegales(Index).dfecha
    lbVinculado = TramitesLegales(Index).bVinculado
    
    TramitesLegales(Index) = pvTramiteLegal
    
    If TramitesLegales(Index).nGarantiaCambiosFila = GarantFilaNueva Then
        TramitesLegales(Index).dfecha = fgFechaHoraMovDate(TramitesLegales(Index).sUltimaActualizacion)
    Else
        TramitesLegales(Index).dfecha = ldFecha
        TramitesLegales(Index).bVinculado = lbVinculado
        TramitesLegales(Index).nGarantiaCambiosFila = GarantFilaModificada
    End If
End Sub
Public Sub EliminarTramiteLegal(ByVal pnFila As Integer)
    Dim TramitesLegalesTmp() As tTramiteLegal
    Dim indexTmp As Integer, Index As Integer
    
    If TramitesLegales(pnFila).nGarantiaCambiosFila = GarantFilaNueva Then 'Si es nuevo lo elimina de memoria
        ReDim TramitesLegalesTmp(0)
        For Index = 1 To UBound(TramitesLegales)
            ReDim Preserve TramitesLegalesTmp(Index)
            TramitesLegalesTmp(Index) = TramitesLegales(Index)
        Next
        ReDim TramitesLegales(0)
        Index = 0
        For indexTmp = 1 To UBound(TramitesLegalesTmp)
            If indexTmp <> pnFila Then
                Index = Index + 1
                ReDim Preserve TramitesLegales(Index)
                TramitesLegales(Index) = TramitesLegalesTmp(indexTmp)
            End If
        Next
        Erase TramitesLegalesTmp
    Else 'Caso contrario lo cambia de estado para eliminarlo cuando se grabe
        CambiarEstadoTramiteLegal pnFila, GarantFilaEliminada
    End If
End Sub
Private Sub CambiarEstadoTramiteLegal(ByVal pnFila As Integer, ByVal pnCambio As eGarantiaCambiosFila)
    TramitesLegales(pnFila).nGarantiaCambiosFila = pnCambio
End Sub
Public Function ObtenerTramiteLegal(ByVal pnFila As Integer) As tTramiteLegal
    ObtenerTramiteLegal = TramitesLegales(pnFila)
End Function
Public Function ObtenerTramiteLegalUltima(Optional ByVal pdFechaMax As Date = CDate("1900-01-01")) As tTramiteLegal
    Dim Index As Integer
    Dim IndexBusca As Integer
    Dim vResultado As tTramiteLegal
    
    IndexBusca = -1
    For Index = 1 To NroTramitesLegal
        If ObtenerTramiteLegal(Index).nGarantiaCambiosFila <> GarantFilaEliminada And ObtenerTramiteLegal(Index).nGarantiaCambiosFila <> GarantFilaOculta Then
            If DateDiff("D", pdFechaMax, CDate("1900-01-01")) = 0 Or (DateDiff("D", pdFechaMax, CDate("1900-01-01")) <> 0 And ObtenerTramiteLegal(Index).dfecha < pdFechaMax) Then 'Para sacar según fecha puedan ingresar
                IndexBusca = Index
            End If
        End If
    Next
    
    If IndexBusca <> -1 Then
        vResultado = ObtenerTramiteLegal(IndexBusca)
    End If
    
    ObtenerTramiteLegalUltima = vResultado
End Function
Public Function GrabarDatosGarantia(ByVal pbEliminarRegCobCredProcesoXDescob As Boolean, _
                                    ByVal pbPermisoDarBajaGarantia As Boolean) As Boolean 'JGPA added fbPermisoDarBajaGarantia según TI-ERS026-2018
    Dim oNGar As New NCOMGarantia
    Dim rsGarantia As ADODB.Recordset
    
    On Error GoTo ErrRegistrar
    
    LlenaRecordSet_Garantia rsGarantia, pbEliminarRegCobCredProcesoXDescob
    
    GrabarDatosGarantia = oNGar.GrabarDatosGarantia(sGarantID, rsGarantia, Propietarios, Valorizaciones, TramitesLegales, pbPermisoDarBajaGarantia)
    
    Exit Function
ErrRegistrar:
    Err.Raise Err.Number, "Error al Registrar Garantía", Err.Description
End Function
Private Sub LlenaRecordSet_Garantia(ByRef pRsGarantia As ADODB.Recordset, ByVal pbEliminarRegCobCredProceso As Boolean)
    Set pRsGarantia = New ADODB.Recordset
         With pRsGarantia
        .Fields.Append "sGarantID", adVarChar, 8
        .Fields.Append "nMoneda", adSmallInt
        .Fields.Append "nClasificacion", adSmallInt
        .Fields.Append "nBienGarantia", adSmallInt
        .Fields.Append "cEmisorCod", adVarChar, 13
        .Fields.Append "nDocTpo", adSmallInt
        .Fields.Append "cDocNro", adVarChar, 50
        .Fields.Append "dConstatacion", adDate
        .Fields.Append "sUltimacionActualizacion", adVarChar, 25
        .Fields.Append "bEliminarRegCobCredProceso", adBoolean 'EJVG20160420
        .Fields.Append "nTpoBienContrato", adSmallInt 'CTI5 ERS01-2020
        .Fields.Append "sCtaCodDesembolsado", adVarChar, 18 'CTI5 ERS01-2020
        .Fields.Append "bCambiadoBienFuturo", adBoolean 'CTI5 ERS01-2020
        .Fields.Append "dFechaSistema", adDate 'CTI5 ERS01-2020
        .Fields.Append "sAgeCtaDesembolsado", adVarChar, 2
        .Fields.Append "sMovNroCreditoPoliza", adVarChar, 25
        .Open '
        .AddNew
        .Fields("sGarantID") = GarantID
        .Fields("nMoneda") = Moneda
        .Fields("nClasificacion") = Clasificacion
        .Fields("nBienGarantia") = BienGarantia
        .Fields("cEmisorCod") = EmisorCod
        .Fields("nDocTpo") = DocumentoTpo
        .Fields("cDocNro") = DocumentoNro
        .Fields("dConstatacion") = FechaConstata
        .Fields("sUltimacionActualizacion") = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        .Fields("bEliminarRegCobCredProceso") = pbEliminarRegCobCredProceso 'EJVG20160420
        .Fields("nTpoBienContrato") = TpoBienContrato 'CTI5 ERS01-2020
        .Fields("sCtaCodDesembolsado") = CtaCodDesembolsado 'CTI5 ERS01-2020
        .Fields("bCambiadoBienFuturo") = CambiadoBienFuturo 'CTI5 ERS01-2020
        .Fields("dFechaSistema") = FechaSistema 'CTI5 ERS01-2020
        .Fields("sAgeCtaDesembolsado") = AgeCtaDesembolsado 'CTI5 ERS01-2020
        .Fields("sMovNroCreditoPoliza") = GeneraMovNro(gdFecSis, AgeCtaDesembolsado, "SIST") 'CTI5 ERS01-2020
    End With
End Sub
Public Function cargarDatos(ByVal psGarantiaID As String) As Boolean
    Dim obj As New COMDCredito.DCOMGarantia
    Dim R As New ADODB.Recordset
    Dim P As New ADODB.Recordset
    Dim V As New ADODB.Recordset
    Dim TL As New ADODB.Recordset
    Dim rsValoriza As New ADODB.Recordset
    
    Dim lvValDirTot As tValorDirectoTotal
    Dim lvValDirDet() As tValorDirectoDetallado
    Dim lvValAutoliquidable As tValorAutoLiquidable
    Dim lvValorTasacionInmobiliaria As tValorTasacionInmobiliaria
    Dim lvValorTasacionVehicular As tValorTasacionVehicular
    Dim lvValorTasacionMobiliariaOtras As tValorTasacionMobiliariaOtras
    Dim lvValorGravamenFavorOtraIFi As tValorGravamenFavorOtraIFi
    
    Dim lvTramiteLegal As tTramiteLegal
    
    Dim iDet As Integer
    Dim bExistePropietario As Boolean, bExisteValorizacion As Boolean

    On Error GoTo ErrCargarDatos
    
    NuevaGarantia
    Set R = obj.RecuperaGarantiaxID(psGarantiaID)
    If Not R.EOF Then
        GarantID = R!cNumGarant
        Moneda = R!nmoneda
        Clasificacion = R!nClasificacion
        BienGarantia = R!nBienGarantia
        ClasificacionSBS = R!cClasificacionSBS
        ClasificacionInterna = R!cClasificacionInterna
        EmisorCod = R!cEmisorCod
        EmisorNombre = R!cEmisorNombre
        DocumentoTpo = R!nDocTpo
        DocumentoNro = R!cDocNro
        If IIf(IsNull(R!dConstatacion), 0, 1) = 1 Then
            FechaConstata = R!dConstatacion
        End If
        BloqueoLegal = R!bBloqueoLegal
        NroCredVinc = R!nNroCredVinc
        NroCredVig = R!nNroCredVig
        NroCredEnProceso = R!nNroCredEnProceso
        DJBloqueada = R!bDJBloqueada
        Estado = R!nEstado 'EJVG20160618
        '***JGPA TI-ERS026-2018-----------------
        NroCredCancelados = R!nNroCredCancelados
        NroCredEnProceso2 = R!nNroCredEnProceso2
        NroCredVig2 = R!nNroCredVig2
        '***End JGPA----------------------------
        'CTI5 ERS001-2020***********************
        TpoBienContrato = R!nTpoBienContrato
        '***************************************
                'JOEP20210820    ***********************
        SalaValida = R!nSaltaValida
        '***************************************
                                                                           
        Set P = obj.RecuperaGarantiaPersonaxID(psGarantiaID)
        Do While Not P.EOF
            InsertarPropietario P!cPersCod, P!cPersNombre, P!cRelacion, GarantFilaSinCambios
            P.MoveNext
        Loop

        Set V = obj.RecuperaGarantiaValorizaxID(psGarantiaID)
        Do While Not V.EOF
            Select Case V!nTipoValorizacion
                Case eGarantiaTipoValorizacion.ValorDirectoTotal
                    Set rsValoriza = obj.RecuperaGarantiaValorizaxTipo(psGarantiaID, V!dValorizacion, ValorDirectoTotal)
                    If Not rsValoriza.EOF Then
                        lvValDirTot.nVRM = rsValoriza!nVRM
                    End If
                Case eGarantiaTipoValorizacion.ValorDirectoDetallado
                    ReDim lvValDirDet(0)
                    Set rsValoriza = obj.RecuperaGarantiaValorizaxTipo(psGarantiaID, V!dValorizacion, ValorDirectoDetallado)
                    Do While Not rsValoriza.EOF
                        iDet = UBound(lvValDirDet) + 1
                        ReDim Preserve lvValDirDet(iDet)
                        lvValDirDet(iDet).sDescripcion = rsValoriza!cDescripcion
                        lvValDirDet(iDet).nCantidad = rsValoriza!nCantidad
                        lvValDirDet(iDet).nValor = rsValoriza!nValor
                        lvValDirDet(iDet).nDocTpo = rsValoriza!nDocTpo
                        lvValDirDet(iDet).sDocTpo = rsValoriza!cDocTpo
                        lvValDirDet(iDet).sDocNro = rsValoriza!cDocNro
                        rsValoriza.MoveNext
                    Loop
                Case eGarantiaTipoValorizacion.GarantiaAutoliquidable
                    Set rsValoriza = obj.RecuperaGarantiaValorizaxTipo(psGarantiaID, V!dValorizacion, GarantiaAutoliquidable)
                    If Not rsValoriza.EOF Then
                        lvValAutoliquidable.sCtaCod = rsValoriza!cCtaCod
                        lvValAutoliquidable.nFormaRetiro = rsValoriza!nFormaRetiro
                        lvValAutoliquidable.sFormaRetiro = rsValoriza!cFormaRetiro
                        lvValAutoliquidable.nMontoCuenta = rsValoriza!nMontoCuenta
                        lvValAutoliquidable.nMontoCuentaDisp = rsValoriza!nMontoCuentaDisp
                        lvValAutoliquidable.nMontoDisponible = rsValoriza!nMontoDisponible
                    End If
                Case eGarantiaTipoValorizacion.TasacionInmobiliaria
                    Set rsValoriza = obj.RecuperaGarantiaValorizaxTipo(psGarantiaID, V!dValorizacion, TasacionInmobiliaria)
                    If Not rsValoriza.EOF Then
                        lvValorTasacionInmobiliaria.sUbicacionGeografica = rsValoriza!cUbiGeoCod
                        lvValorTasacionInmobiliaria.sDireccion = rsValoriza!cDireccion
                        lvValorTasacionInmobiliaria.nValorComercial = rsValoriza!nValorComercial
                        lvValorTasacionInmobiliaria.nValorEdificacion = rsValoriza!nValorEdificacion
                        lvValorTasacionInmobiliaria.nVRM = rsValoriza!nVRM
                        lvValorTasacionInmobiliaria.nClase = rsValoriza!nClase
                        lvValorTasacionInmobiliaria.nCategoria = rsValoriza!nCategoria
                        lvValorTasacionInmobiliaria.nNroLocales = rsValoriza!nNroLocales
                        lvValorTasacionInmobiliaria.nNroPisos = rsValoriza!nNroPisos
                        lvValorTasacionInmobiliaria.nNroSotanos = rsValoriza!nNroSotanos
                        lvValorTasacionInmobiliaria.nAnioConstruccion = rsValoriza!nAnioConstruccion
                        lvValorTasacionInmobiliaria.dTasacion = rsValoriza!dTasacion
                        lvValorTasacionInmobiliaria.nTasacionTC = rsValoriza!nTasacionTC
                        lvValorTasacionInmobiliaria.sTasadorCod = rsValoriza!cPersCodTasador
                        lvValorTasacionInmobiliaria.sTasadorNombre = rsValoriza!cPersNombreTasador
                        lvValorTasacionInmobiliaria.sTasadorDNI = rsValoriza!Dni
                        lvValorTasacionInmobiliaria.sTasadorREPEV = rsValoriza!REPEV
                        lvValorTasacionInmobiliaria.bMigrado = V!bMigrado
                        
                        'CTI5 ERS0012020
                        lvValorTasacionInmobiliaria.sMz = rsValoriza!cMz
                        lvValorTasacionInmobiliaria.sLt = rsValoriza!cLt
                        lvValorTasacionInmobiliaria.Set = rsValoriza!cEt
                        'END
                    End If
                Case eGarantiaTipoValorizacion.TasacionVehicular
                    Set rsValoriza = obj.RecuperaGarantiaValorizaxTipo(psGarantiaID, V!dValorizacion, TasacionVehicular)
                    If Not rsValoriza.EOF Then
                        lvValorTasacionVehicular.sDescripcion = rsValoriza!cDescripcion
                        lvValorTasacionVehicular.sCarroceria = rsValoriza!cCarroceria
                        lvValorTasacionVehicular.sMarca = rsValoriza!cMarca
                        lvValorTasacionVehicular.sModelo = rsValoriza!cModelo
                        lvValorTasacionVehicular.sNroPlaca = rsValoriza!cNroPlaca
                        lvValorTasacionVehicular.sNroMotor = rsValoriza!cNroMotor
                        lvValorTasacionVehicular.sNroChasis = rsValoriza!cNroChasis
                        lvValorTasacionVehicular.nAnioFabricacion = rsValoriza!nAnioFabricacion
                        lvValorTasacionVehicular.nAnioAdquisicionTienda = rsValoriza!nAnioAdqTienda
                        lvValorTasacionVehicular.nValorComercial = rsValoriza!nValorComercial
                        lvValorTasacionVehicular.nVRM = rsValoriza!nVRM
                        lvValorTasacionVehicular.dTasacion = rsValoriza!dTasacion
                        lvValorTasacionVehicular.nTasacionTC = rsValoriza!nTasacionTC
                        lvValorTasacionVehicular.sTasadorCod = rsValoriza!cPersCodTasador
                        lvValorTasacionVehicular.sTasadorNombre = rsValoriza!cPersNombreTasador
                        lvValorTasacionVehicular.sTasadorDNI = rsValoriza!Dni
                        lvValorTasacionVehicular.sTasadorREPEV = rsValoriza!REPEV
                        lvValorTasacionVehicular.bMigrado = V!bMigrado
                    End If
                Case eGarantiaTipoValorizacion.OtrasTasacionesMobiliarias
                    Set rsValoriza = obj.RecuperaGarantiaValorizaxTipo(psGarantiaID, V!dValorizacion, OtrasTasacionesMobiliarias)
                    If Not rsValoriza.EOF Then
                        lvValorTasacionMobiliariaOtras.sDescripcion = rsValoriza!cDescripcion
                        lvValorTasacionMobiliariaOtras.sNroSerie = rsValoriza!cNroSerie
                        lvValorTasacionMobiliariaOtras.sMarca = rsValoriza!cMarca
                        lvValorTasacionMobiliariaOtras.sModelo = rsValoriza!cModelo
                        lvValorTasacionMobiliariaOtras.nAnioFabricacion = rsValoriza!nAnioFabricacion
                        lvValorTasacionMobiliariaOtras.nValorComercial = rsValoriza!nValorComercial
                        lvValorTasacionMobiliariaOtras.nVRM = rsValoriza!nVRM
                        lvValorTasacionMobiliariaOtras.dTasacion = rsValoriza!dTasacion
                        lvValorTasacionMobiliariaOtras.nTasacionTC = rsValoriza!nTasacionTC
                        lvValorTasacionMobiliariaOtras.sTasadorCod = rsValoriza!cPersCodTasador
                        lvValorTasacionMobiliariaOtras.sTasadorNombre = rsValoriza!cPersNombreTasador
                        lvValorTasacionMobiliariaOtras.sTasadorDNI = rsValoriza!Dni
                        lvValorTasacionMobiliariaOtras.sTasadorREPEV = rsValoriza!REPEV
                        lvValorTasacionMobiliariaOtras.bMigrado = V!bMigrado
                    End If
                Case eGarantiaTipoValorizacion.GravamenFavorOtraIFi 'EJVG20160227
                    Set rsValoriza = obj.RecuperaGarantiaValorizaxTipo(psGarantiaID, V!dValorizacion, GravamenFavorOtraIFi)
                    lvValorGravamenFavorOtraIFi.sDescripcion = rsValoriza!cDescripcion
                    lvValorGravamenFavorOtraIFi.nValorComercial = rsValoriza!nValorComercial
                    lvValorGravamenFavorOtraIFi.nOpTpRegPub = rsValoriza!nTpRegPub  'JOEP20180524 ERS009-2018
                    ReDim lvValorGravamenFavorOtraIFi.vValorGravamenFavorOtraIFiDet(0)
                    Do While Not rsValoriza.EOF
                        ReDim Preserve lvValorGravamenFavorOtraIFi.vValorGravamenFavorOtraIFiDet(rsValoriza.Bookmark)
                        lvValorGravamenFavorOtraIFi.vValorGravamenFavorOtraIFiDet(rsValoriza.Bookmark).sIFICod = rsValoriza!cPersCod
                        lvValorGravamenFavorOtraIFi.vValorGravamenFavorOtraIFiDet(rsValoriza.Bookmark).sIFINombre = rsValoriza!cPersNombre
                        lvValorGravamenFavorOtraIFi.vValorGravamenFavorOtraIFiDet(rsValoriza.Bookmark).sIFTpo = rsValoriza!cIFTpo
                        lvValorGravamenFavorOtraIFi.vValorGravamenFavorOtraIFiDet(rsValoriza.Bookmark).nGravado = rsValoriza!nGravado
                        lvValorGravamenFavorOtraIFi.bMigrado = V!bMigrado
                        rsValoriza.MoveNext
                    Loop
            End Select
            
            InsertarValorizacion V!dValorizacion, V!cUltimaActualizacion, V!nTipoValorizacion, V!cTipoValorizacion, V!cGlosa, GarantFilaSinCambios, V!bVinculado, lvValDirTot, lvValDirDet, lvValAutoliquidable, lvValorTasacionInmobiliaria, lvValorTasacionVehicular, lvValorTasacionMobiliariaOtras, False, lvValorGravamenFavorOtraIFi
            '***
            VinculadoUltVAL = V!bVinculado
            PermiteNuevaValorizacion = V!bPermiteAgregarValorizacion
            '***
            V.MoveNext
        Loop

        Set TL = obj.RecuperaTramiteLegalxID(psGarantiaID, VinculadoUltVAL, PermiteNuevaValorizacion)
        If Not TL.EOF Then
            Do While Not TL.EOF
                lvTramiteLegal.dfecha = TL!dRegistro
                lvTramiteLegal.sUltimaActualizacion = TL!cUltimaActualizacion
                lvTramiteLegal.sNroPartidaRegistral = TL!cNroPartidaRegistral
                lvTramiteLegal.nTipoTramite = TL!nTipoTramite
                lvTramiteLegal.cTipoTramite = TL!cTipoTramite
                lvTramiteLegal.nGravamen = TL!nGravamen
                lvTramiteLegal.nZonaRegistralID = TL!nZonaRegistralID
                lvTramiteLegal.cZonaRegistralNombre = TL!cZonaRegistralNombre
                lvTramiteLegal.nOficinaRegistralID = TL!nOficinaRegistralID
                lvTramiteLegal.cOficinaRegistralNombre = TL!cOficinaRegistralNombre
                lvTramiteLegal.sNotariaCod = TL!cNotariaPersCod
                lvTramiteLegal.sNotariaNombre = TL!cNotariaPersNombre
                lvTramiteLegal.nEstado = TL!nEstado
                lvTramiteLegal.cEstado = TL!cEstado
                lvTramiteLegal.nTipoCobertura = TL!nTipoCobertura
                lvTramiteLegal.sNroAsiento = TL!cNroAsiento
                lvTramiteLegal.dInscripcion = TL!dInscripcion
                lvTramiteLegal.bVinculado = TL!bVinculado
                lvTramiteLegal.bMigrado = TL!bMigrado
                    
                InsertarTramiteLegal lvTramiteLegal, GarantFilaSinCambios, False
                PermiteNuevoTramiteLegal = TL!bPermiteAgregarTramiteLegal
                TL.MoveNext
            Loop
        Else
            If VinculadoUltVAL And Not PermiteNuevaValorizacion Then 'Valorización ya tiene Cobertura(Gravamen)
                PermiteNuevoTramiteLegal = False
            End If
        End If
        cargarDatos = True
    End If
    
    RSClose rsValoriza
    RSClose R
    RSClose P
    RSClose V
    RSClose TL
    Set obj = Nothing
    Exit Function
ErrCargarDatos:
    NuevaGarantia
    cargarDatos = False
End Function
Public Function ExisteGarantia(ByRef psNumGarantExiste As String) As Boolean
    Dim oDGar As New COMDCredito.DCOMGarantia
    Dim rsGarantia As ADODB.Recordset
    
    On Error GoTo ErrExisteGarantia

    ExisteGarantia = oDGar.ExisteGarantiaNew(EmisorCod, DocumentoTpo, DocumentoNro, GarantID, psNumGarantExiste)
    
    Exit Function
ErrExisteGarantia:
    Err.Raise Err.Number, "Error al verificar la existencia de la Garantía", Err.Description
End Function
Public Function CadenaCoberturaGarantia(ByVal pbTramiteLegal As Boolean, ByVal pnTipoTramite As Integer, ByVal pnVRM As Currency, ByVal pnGravamen As Currency) As String
    Dim oDGar As New DCOMGarantia
    CadenaCoberturaGarantia = oDGar.CadenaCoberturaGarantia(GarantID, pbTramiteLegal, pnTipoTramite, pnVRM, pnGravamen)
    Set oDGar = Nothing
End Function
Public Function ObviarDescobertura() As Boolean
    Dim oDGar As New DCOMGarantia
    ObviarDescobertura = oDGar.ObviarDescobertura(GarantID)
    Set oDGar = Nothing
End Function
'*************************
