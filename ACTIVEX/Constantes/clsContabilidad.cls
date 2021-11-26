VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum CGTipoIF
    gTpoIFBanco = 1
    gTpoIFFinanciera = 2
    gTpoIFCmac = 3
    gTpoIFCrac = 4
    gTpoIFFuenteFinanciamiento = 5
    gTpoIFCooperativa = 6
    gTpoIFEDPYME = 7
End Enum

Public Enum CGEstadoCtaIF
    gEstadoCtaIFRegistrada = 0
    gEstadoCtaIFActiva = 1    'A
    gEstadoCtaIFCancelada = 2  'B
    gEstadoCtaIFAnulada = 3
    gEstadoCtaIFRestringida = 4
End Enum

Public Enum CGTipoCtaIF
    gTpoCtaIFCtaCte = 1
    gTpoCtaIFCtaAho = 2
    gTpoCtaIFCtaPF = 3
    gTpoCtaIFCtaAdeud = 5
    gTpoCtaIFAdeudProyecta = 6
End Enum

Public Enum CGTipoEstCuotaAdeud
    gTpoEstCuotaAdeudPend = 0
    gTpoEstCuotaAdeudCanc = 1
    gTpoEstCuotaAdeudOtros = 2
End Enum

Public Enum CGTipoPagoAdeud
    gTpoPagoAdeudCapital = 1
    gTpoPagoAdeudInteres = 2
    gTpoPagoAdeudGastos = 3
End Enum

Public Enum CGTipoCuotCalIF
    gCGTipoCuotCalIFDesem = 1
    gCGTipoCuotCalIFCuota = 2
    gCGTipoCuotCalIFNoConcesional = 6
End Enum
Public Enum CGEstadosChq
    gCGEstadosChqRecibido = 0
    gCGEstadosChqDepositado = 1
    gCGEstadosChqRechazado = 2
End Enum

Public Enum MovEstado
    gMovEstContabMovContable = 10
    gMovEstContabPendiente = 11
    gMovEstContabRechazado = 12
    gMovEstContabNoContable = 13
    gMovEstContabSustPendRendir = 14
    gMovEstPresupRechazado = 14
    gMovEstPresupPendiente = 15
    gMovEstPresupAceptado = 16
    gMovEstPresupProyectado = 17
    gMovEstLogIngBienAceptado = 20
    gMovEstLogIngBienRechazado = 21
    gMovEstLogSaleBienAlmacen = 22
End Enum

Public Enum MovFlag   ' "gsMovFlag"
    gMovFlagVigente = 0
    gMovFlagEliminado = 1   ' "X"
    gMovFlagExtornado = 2   '  "E"
    gMovFlagDeExtorno = 3   ' "N"
    gMovFlagDeLeido = 4     '  "L"
    gMovFlagModificado = 5  ' "M"
End Enum

Public Enum MovParalelo   '= "CMovParalelo"
    gMovParaleloTransferido = 1   ' "T"
End Enum

Public Enum CHTipoProc   '    Tipo de Proceso de Caja Chica
    gCHTipoProcHabilitacion = 1     'Habilitacion
    gCHTipoProcRendicion = 2   'Rendición en Contabilidad
    gCHTipoProcArqueo = 3  'Arqueo
    gCHTipoProcDesembolso = 4  'Desembolso de Caja Chica
    gCHTipoProcEgresoDirecto = 5    'Egreso Directo de Caja Chica
    gCHTipoProcArendir = 6     'A rendir Cuenta
    gCHTipoProcNuevaPorRendicion = 7 'Nueva por rendicion
    gCHTipoProcCancelada = 8 'Cancelada
End Enum

Public Enum ArendirTipo  '     Tipo de Arendir Cuenta
    gArendirTipoCajaGeneral = 1   'A rendir Cuenta de Caja General
    gArendirTipoViaticos = 2     'A rendir Cuenta de Viáticos
    gArendirTipoCajaChica = 3    'A rendir Cuenta de Caja Chica
    gArendirTipoAgencias = 4     'A rendir Cuenta de Agencias
End Enum

Public Enum ViaticosCateg         'Viaticos Categoría del Trabajador CMACT
    gViaticosCategFuncionario = 1  ' A     'Funcionario
    gViaticosCategEmpleado = 2     ' B    'Empleado
End Enum

Public Enum ViaticosAfectoA
    gViaticoAfectoADias = 1     '  D   Nro de Dias
    gViaticoAfectoATransporte = 2   ' T   Transporte
    gViaticoAfectoANinguno = 3      'N   Ninguno
End Enum

Public Enum ViaticosTransporte         'Viaticos Transporte
    gViaticosTransporteAereo = 1    'Aereo
    gViaticosTransporteMaritimo = 2     'Maritimo
    gViaticosTransporteTerrestre = 3    'Terrestre
End Enum

Public Enum ViaticosDestino            'Viaticos Destino
    gViaticosDestinoLaLibertad = 1        'La Libertad
    gViaticosDestinoLimaIqAreqCuz = 2     'Lima, Iquitos, Arequipa o Cuzco
    gViaticosDestinoOtros = 3             'Otras Ciudades
End Enum

Public Enum OpeTpoVisible
   gOpeTpoVisibleSi = 1   '1 SI
   gOpeTpoVisibleNo = 0   '0 NO
End Enum

Public Enum OpeCtaCaracter
   gOpeCtaCaracOpcional = 2      'Opcional
   gOpeCtaCaracObligatorio = 1   'Obligatorio
End Enum

Public Enum FiltroObj
    CtaOBjFiltroObj = 0
    CtaOBjFiltroIF = 1
End Enum
Public Enum TpoObjetos
    ObjPersona = 0
    ObjEntidadesFinancieras = 1
    ObjAccionesSectorPrivado = 2
    ObjCMACAgencias = 11
    ObjCMACArea = 12
    ObjCMACAgenciaArea = 13
    ObjBienesServicios = 18
    ObjGastosOrgInstalacion = 19
    ObjConceptosARendir = 20
    ObjArendir = 50
    ObjCtasAhorrosCMACS = 51                   'CUENTAS DE AHORROS DE CMACT
    ObjProductosCMACT = 60
    ObjDescomEfectivo = 90
    ObjServRecTerceros = 4511
    ObjGastosDirectorio = 4415                 'GASTOS DE DIRECTORIO
    ObjOtrasCargasDivGestion = 4719            'OTRAS CARGAS DIVERSAS DE GESTION
End Enum
Public Enum TpoMargenes
    MargenSuperior = 0
    MargenInferior = 1
    MargenDerecho = 2
    MargenIzquierdo = 3
End Enum
Public Enum OpeDocEstado
    OpeDocEstObligatorioDebeExistir = 11
    OpeDocEstObligatorioNoDebeExistir = 12
    OpeDocEstOpcionalDebeExistir = 21
    OpeDocEstOpcionalNoDebeExistir = 22
End Enum
Public Enum OpeDocMetodo
    OpeDocMetNroMovimiento = 1
    OpeDocMetAutogenerado = 2
    OpeDocMetDigitado = 3
End Enum
Public Enum ARendirFases
    ArendirSolicitud = 0
    ArendirRechazo
    ArendirAtencion
    ArendirSustentacion
    ArendirRendicion
    ArendirExtornoAtencion
    ArendirExtornoRendicion
End Enum

Public Enum CGEstadosOp
    PorEmitir = 0
    Emitidas = 1
    Anuladas = 2
    Rechazadas = 3
    Extraviadas = 4
End Enum

Public Enum CGEstadoConfCheque
    ChqCGSinConfirmacion = 0
    ChqCGNoConfirmado = 1
    ChqCGConfirmado = 2
End Enum
Public Enum CGTipoTransporte
    CGTipoTransporteBlindado = 0
    CGTipoTransportePropio = 1
    CGTipoTransporteAlquilado = 2
End Enum

Public Enum CGCtaIFConceptos
    CGCtaIFConCapital = 1
    CGCtaIFConInteres = 2
    CGCtaIFConIntAdel = 3
    CGCtaIFConMora = 4
    CGCtaIFConGastos = 5
    CGCtaIFConGastosServ = 6
    CGCtaIFConComisiones = 7
End Enum
Public Enum CGBloqueos
   gBloqueoBalance = 1
   gBloqueoSaldos = 2
End Enum


Public Enum CGAnxTipoCalculo
   gAnxTCalcTasa = 1
   gAnxTCalcVencimiento = 2
   gAnxTCalcEncajeBCR = 3
   gAnxTCalcPlazaCheque = 4
   gAnxTCalcSeries = 5
   gAnxTCalcFecha = 6
   gAnxTCalcFormula = 7
   gAnxTCalcFormulasAcumula = 8
   gAnxTCalcTotales = 9
End Enum

Public Enum CGAdeudCalTpoCuota
   gAdeudTpoCuotaFija = 1
   gAdeudTpoCuotaVariable = 2
End Enum

Public Enum ContAnx10TipoOficina
    gAnx10TipoOfPrincipal = 1
    gAnx10TipoOfSucursal = 2
    gAnx10TipoOfAgencia = 3
    gAnx10TipoOfEspecial = 4
    gAnx10TipoOfLocalCompar = 5
End Enum

Public Enum ARendirDestinoGasto
    gARendirDestinoGastoJudiciales = 1
    gARendirDestinoGastoOtros = 2
    gARendirDestinoGastoViaticos = 3
    gARendirDestinoGastoNoDefinido = 4
End Enum

