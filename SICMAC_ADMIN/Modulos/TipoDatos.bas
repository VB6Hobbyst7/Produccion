Attribute VB_Name = "TipoDatos"
Option Explicit

Global Const gsRHPlanillaSueldos = "E01"
Global Const gsRHPlanillaGratificacion = "E02"
Global Const gsRHPlanillaTercio = "E03"
Global Const gsRHPlanillaUtilidades = "E04"
Global Const gsRHPlanillaCTS = "E05"
Global Const gsRHPlanillaCTSFractal = "27" 'MAVM 20120504 CTS Fractal
Global Const gsRHPlanillaVacaciones = "E06"
Global Const gsRHPlanillaSubsidio = "E07"
Global Const gsRHPlanillaLiquidacion = "E08"
Global Const gsRHPlanillaBonificacionVacacinal = "E09"
Global Const gsRHPlanillaBonoProductividad = "E10"
Global Const gsRHPlanillaBonoAguinaldo = "E11"
Global Const gsRHPlanillaSubsidioEnfermedad = "E12"
Global Const gsRHPlanillaReintegro = "E13"
Global Const gsRHPlanillaDev5ta = "E14"
Global Const gsRHPlanillaMovilidad = "E15"

Global Const gnRHTotalTpo = 9999
Global Const gsRHConceptoUMESTRAB = "U_POR_MES_TRAB"
Global Const gsRHConceptoITOTING = "I_TOT_ING"
Global Const gsRHConceptoITOTINGCOD = "130"
Global Const gsRHConceptoINETOPAGARCOD = "112"
Global Const gsRHConceptoINETOPAGAR = "I_NETO_PAGAR"
Global Const gsRHConceptoDTOTDES = "D_TOT_DESC"
Global Const gsRHConceptoDTOTDESCOD = "215"

'ALPA 20101026
Global Const gsRHConceptoCanasNaviCOD = "194"
Global Const gsRHConceptoBonoCreEcCOD = "195"
'***************
Global Const gsRHConceptoVTOTREM = "V_TOT_REM"
Global Const gsRHConceptoITOTCTS = "I_TOT_CTS"
Global Const gsRHConceptoITOTTERCIO = "I_TERCIO"
Global Const gsRHConceptoITOTGRAT = "I_TOTAL_GRATIF"
Global Const gsRHConceptoVNETOPAGAR = "V_NETO_PAGAR"
Global Const gnRHNumDiasVac = 30

Public Const gsRHPlanillaProvGratificacion = 901
Public Const gsRHPlanillaProvVacaciones = 902
Public Const gsRHPlanillaProvCTS = 903
Public Const gsRHPlanillaProvTercio = 904
Public Const gsRHPlanillaProvBonVacaciones = 905

Global Const gsAgenciaPrinsipal = "07"

Global Const gnMes1 = 7
Global Const gnMes2 = 12
Global Const gnNumRem = 14
Global Const gsRIMNA = "O_RIMA"
Global Const gsIMP5TA = "D_IMP_5TA"

Global Const gsRHMotivoAbono = "101"
Global Const gsRHMotivoCargo = "201"


'Codigos de operacion para el calculo de provisiones y remuneraciones
'Sueldos
Global Const gsRHPlanillaSueldosRemEst = "622001"
Global Const gsRHPlanillaSueldosRemCon = "622002"
'Gratificacion
Global Const gsRHPlanillaGratificacionProvEst = "622101"
Global Const gsRHPlanillaGratificacionProvCon = "622102"
Global Const gsRHPlanillaGratificacionRemEst = "622103"
Global Const gsRHPlanillaGratificacionRemCon = "622104"
'Tercio
Global Const gsRHPlanillaTercioProvEst = "622201"
Global Const gsRHPlanillaTercioProvCon = "622202"
Global Const gsRHPlanillaTercioRemEst = "622203"
Global Const gsRHPlanillaTercioRemCon = "622204"
'Utilidades
Global Const gsRHPlanillaUtilidadesProvEst = "622301"
Global Const gsRHPlanillaUtilidadesProvCon = "622302"
Global Const gsRHPlanillaUtilidadesRemEst = "622303"
Global Const gsRHPlanillaUtilidadesRemCon = "622304"
'CTS
Global Const gsRHPlanillaCTSProvEst = "622401"
Global Const gsRHPlanillaCTSProvCon = "622402"
Global Const gsRHPlanillaCTSRemEst = "622403"
Global Const gsRHPlanillaCTSRemCon = "622404"
'Vacaciones
Global Const gsRHPlanillaVacacionesProvEst = "622501"
Global Const gsRHPlanillaVacacionesProvCon = "622502"
Global Const gsRHPlanillaVacacionesRemEst = "622503"
Global Const gsRHPlanillaVacacionesRemCon = "622504"
'Subsidios
Global Const gsRHPlanillaSubsidioRem = "622601"
Global Const gsRHPlanillaSubsidioEnfermedadRem = "622602"
'Liquidacion
Global Const gsRHPlanillaLiquidacionRem = "622701"
'Bonificacion Vacacional
Global Const gsRHPlanillaBonificacionVacacinalProvEst = "622801"
Global Const gsRHPlanillaBonificacionVacacinalProvCon = "622802"
Global Const gsRHPlanillaBonificacionVacacinalRemEst = "622803"
Global Const gsRHPlanillaBonificacionVacacinalRemCon = "622804"
'Bono Productividad
Global Const gsRHPlanillaBonoProductividadProvEst = "622901"
Global Const gsRHPlanillaBonoProductividadProvCon = "622902"
Global Const gsRHPlanillaBonoProductividadRemEst = "622903"
Global Const gsRHPlanillaBonoProductividadRemCon = "622904"

'Bono Reintegro
Global Const gsRHPlanillaReintegroProvEst = "623001"
Global Const gsRHPlanillaReintegroProvCon = "623002"
Global Const gsRHPlanillaReintegroRemEst = "623003"
Global Const gsRHPlanillaReintegroRemCon = "623004"

'Bono Rev 5ta Categoria
Global Const gsRHPlanillaDev5taRemEst = "623101"
Global Const gsRHPlanillaDev5taRemCon = "623102"

'Bono Movilidad
Global Const gsRHPlanillaMovilidadProvEst = "623201"
Global Const gsRHPlanillaMovilidadProvCon = "623202"
Global Const gsRHPlanillaMovilidadRemEst = "623203"
Global Const gsRHPlanillaMovilidadRemCon = "623204"

Public Enum RHEstadosRRHH
    gRHEstadosRRHHINACTFINCONTRATO = 101
    gRHEstadosRRHHACTIVO = 201
    gRHEstadosRRHHVACACIONES = 301
    gRHEstadosRRHHOTRASVACACIONES = 302
    gRHEstadosRRHHLICSINGOCE = 401
    gRHEstadosRRHHLICCONGOCE = 402
    gRHEstadosRRHHPERPERSONAL = 403
    gRHEstadosRRHHPERMEDICO = 404
    gRHEstadosRRHHPERCOMISION = 405
    gRHEstadosRRHHSUBSIDIADO = 501
    gRHEstadosRRHHDESCANSO = 502
    gRHEstadosRRHHSUSPENDIDO = 601
    gRHEstadosRRHHRETIRADODESPIDO = 701
    gRHEstadosRRHHRETIRADORENUNCIA = 702
    gRHEstadosRRHHRETIRADONORENOVACIONCONTRATO = 703
    gRHEstadosRRHHPORLIQUIDARDESPIDO = 801
    gRHEstadosRRHHPORLIQUIDARRENUNCIA = 802
    gRHEstadosRRHHPORLIQUIDARNORENOVACIONCONTRATO = 803
End Enum

Public Enum RHBeneficiarioRela
    gPersRelaBenefDesaCtivado = 0
    gPersRelaBenefACtivo = 1
End Enum

Public Enum RHAsistenciaMedicaRela
    gPersRelaAMPDesactivado = 0
    gPersRelaAMPActivo = 1
End Enum

Public Enum RHProcesoSeleccionTipo
    gRHProcSelTpoInterno = 0
    gRHProcSelTpoExterno = 1
    gRHProcSelTpoParaEmp = 2
End Enum

Public Enum RHProcesoSeleccionEstado
    gRHProcSelEstIniCiado = 0
    gRHProcSelEstFinalizado = 1
End Enum

Public Enum RHProcesoSeleccionResultado
    gRHProcSelResIngreso = 0
    gRHProcSelResSuplente = 1
    gRHProcSelResNoIngreso = 2
End Enum

Public Enum RHEstado
    gRHEstInactivo = 0
    gRHEstActivo = 1
    gRHEstVacacFisicas = 2
    gRHEstVacacSuspend = 3
    gRHEstLicSGoceHaber = 4
    gRHEstLicCGoceHaber = 5
    gRHEstSubs = 6
    gRHEstSuspend = 7
    gRHEstDespedido = 8
    gRHEstRetirado = 9
End Enum

Public Enum RHCondicion
    gRHCondEmpContratado = 0
    gRHCondEmpEstable = 1
End Enum

Public Enum RHCategoria
    gRHCategFunCionario = 1
    gRHCategEmpleado = 2
End Enum

Public Enum RHCargoEstado
    gRHCargoEstACtivo = 1
    gRHCargoEstEliminado = 2
End Enum

Public Enum RHConceptoIngresos
    gRHConcIngBoncarFam = 101
    gRHConcIngBonConsol = 102
    gRHConcIngBonVacacional = 103
    gRHConcIngBonCargo = 104
    gRHConcIngBonResponsabilidad = 105
    gRHConcIngCTSBonVac = 106
    gRHConcIngCTSBonGratif = 107
    gRHConcIngCTSTercio = 108
    gRHConcIngAFP1023 = 109
    gRHConcIngAFP0300 = 110
    gRHConcIngIngDiaSub = 111
    gRHConcIngNetoPagar = 112
    gRHConcIngONP = 113
    gRHConcIngPrueba1 = 114
    gRHConcIngReintegro = 115
    gRHConcIngRemNoAfeCta = 116
    gRHConcIngRemVacacional = 117
    gRHConcIngRemMensual = 118
    gRHConcIngSueBasico = 119
    gRHConcIngTerBonCargo = 120
    gRHConcIngTerBonCarFam = 121
    gRHConcIngTerBonConsol = 122
    gRHConcIngTerCio = 123
    gRHConcIngTerAFP1023 = 124
    gRHConcIngTerAFP300 = 125
    gRHConcIngTerSueBasico = 126
    gRHConcIngTerBonResp = 127
    gRHConcIngTotBonProduct = 128
    gRHConcIngTotCTS = 129
    gRHConcIngTotIngresos = 130
    gRHConcIngTotalUtilidades = 131
    gRHConcIngTotalGratifiCaCiones = 132
    gRHConcIngBonProduct = 133
    gRHConcIngIngresosLocacion = 134
    gRHConcIngAguinaldoNavideño = 135
    gRHConcIngAguinaldoJuguetes = 136
    gRHConcIngVacSueldoBasico = 137
    gRHConcIngVacBonCargo = 138
    gRHConcIngVacAFP300 = 139
    gRHConcIngVacBonResponsabilidad = 140
    gRHConcIngVacBonFam = 141
    gRHConcIngVacAFP1023 = 142
    gRHConcIngVacCTSTotRem = 143
    gRHConcIngVacBonConsol = 144
    gRHConcIngTerBonProduCtividad = 145
    gRHConcIngVacSueBasico = 146
    gRHConcIngUtilSueDias = 147
    gRHConcIngUtilSueMonto = 148
    gRHConcIngIncentivoProd = 149
    gRHConcIngDev5taCategoria = 150
    gRHConcIngReintegroCTS = 151
    gRHConcIngIncentivoPorProd = 152
    gRHConcIngRefrigerio = 153
    gRHConcIngLiquidaCTS = 154
    gRHConcIngLiquidaVacTruncas = 155
    gRHConcIngLiquidaGratificacion = 156
    gRHConcIngLiquidaDiasLaborados = 157
    gRHConcIngReintegroDiaFeriado = 158
    gRHConcIngLiquidaCompVac = 159
    gRHConcIngIngresoXReconsideracion = 160
    
End Enum

Public Enum RHConceptoDescuentos
    gRHConcDcto5taCat = 201
    gRHConcDctoAFPComVar = 202
    gRHConcDctoAFPCuotaFija = 203
    gRHConcDctoAFPPrima = 204
    gRHConcDctoAMP = 205
    gRHConcDctoBcoLima = 206
    gRHConcDctoEsSaludVida = 207
    gRHConcDctoJudFijo = 208
    gRHConcDctoJudPorc = 209
    gRHConcDctoHipotec = 210
    gRHConcDctoTardanzas = 211
    gRHConcDctoONP1300 = 212
    gRHConcDctoPrueba = 213
    gRHConcDctoPrestAdmin = 214
    gRHConcDctoTotalDctos = 215
    
    gRHConcDctoPrestAdminOtros = 224
End Enum

Public Enum RHConceptoAportaciones
    gsRHCAportExtSolid = 301
    gsRHCAportSegSoc = 302
    gsRHCAportTotalAport = 303
    gsRHCAportLiquidaExtSolid = 304
    gsRHCAportLiquidaSegSoc = 305
    gsRHCAportExtSolidManual = 306
    gsRHCAportSegSocManual = 307
End Enum

Public Enum RHConceptoVariablesUsuario
    gsRHCVarUsuCTSMesTrab = 401
    gsRHCVarUsuGratMesTrab = 402
    gsRHCVarUsuSubNumDias = 403
    gsRHCVarUsuPorcMesTrab = 404
End Enum

Public Enum RHConceptoVariablesGlobales
    gsRHCVarGlobTotDiasLab = 501
    gsRHCVarGlobPorcPagUTIL = 502
    gsRHCVarGlobSumTotIng = 503
    gsRHCVarGlobValUtil = 504
End Enum

Public Enum RHConceptoFuncionesConst
    gsRHCFuncConstAFPValComVar = 601
    gsRHCFuncConstAFPValorPrima = 602
    gsRHCFuncConstDctoJudiFijo = 603
    gsRHCFuncConstDctoJudiPorc = 604
    gsRHCFuncConst5taCat = 605
    gsRHCFuncConstAFP1023 = 606
    gsRHCFuncConstAFP0300 = 607
    gsRHCFuncConstCTSMesTrab = 608
    gsRHCFuncConstMinSal = 609
    gsRHCFuncConstMinTard = 610
    gsRHCFuncConstDiasnoTrab = 611
    gsRHCFuncConstNumMesTrab = 612
    gsRHCFuncConstTERNumMesTrab = 613
    gsRHCFuncConstPagoExeso = 614
    gsRHCFuncConstReintegro = 615
    gsRHCFuncConstSdoBas = 616
    gsRHCFuncConstSdoBasNivel = 617
    gsRHCFuncConstSdoContrato = 618
    gsRHCFuncConstUltBonVacac = 619
    gsRHCFuncConstUltGratif = 620
    gsRHCFuncConstUltTercio = 621
End Enum

Public Enum RHConceptoVariablesLocales
    gsRHCVarTotIngAcum = 701
    gsRHCVarAFPCuotaFija = 702
    gsRHCVarIngAntesAFP = 703
    gsRHCVarMontoAMP = 704
    gsRHCVarNumDiasLab = 705
    gsRHCVarNumDiasMes = 706
    gsRHCVarNumMinTrab = 707
    gsRHCVarPrueba = 708
    gsRHCVarPorcAsistMedFam = 709
    gsRHCVarPorcONP = 710
    gsRHCVarPorcFonavi = 711
    gsRHCVarPorc5taCat = 712
    gsRHCVarPorcSegSoc = 713
    gsRHCVarPorcExtSolid = 714
    gsRHCVarPorcIncONP = 715
    gsRHCVarPorcRPS = 716
    gsRHCVarPorcTercio = 717
    gsRHCVarPromDia = 718
    gsRHCVarSueldoNeto = 719
    gsRHCVarSumMonInt = 720
    gsRHCVarAFPTopePrima = 721
    gsRHCVarAPFTope = 722
    gsRHCVarTotalRemun = 723
    gsRHCVarUITValor = 724
    gsRHCVarUIT54 = 725
    gsRHCVarUIT07 = 726
    gsRHCVarDiaProm = 727
End Enum

Public Enum RHConceptoTipo
    gsRHTpoConcConst = 1
    gsRHTpoConcValCal = 2
    gsRHTpoConcValPreDef = 3
    gsRHTpoConcFunInt = 4
End Enum

Public Enum RHConceptoEstado
    gsRHCEstAct = 1
    gsRHCEstDesh = 0
End Enum

Public Enum RHImpuesto5taCat
    gsRHImp5taCatCDcto = 1
    gsRHImp5taCatSDcto = 0
End Enum

Public Enum RHAfectoMesTrabajado
    gsRHAfeCtoMesTrabajado = 1
    gsRHNoAfeCtoMesTrabajado = 0
End Enum

Global Const gRHEvaluacionComite = 6022
Global Const gGenTipoPeriodos = 1012

Public Enum RHConceptoTipoCal
    RHConceptoTipo_CONSTANTES = 1
    RHConceptoTipo_VALORCALCULADO = 2
    RHConceptoTipo_VALOR_PRE_DEFINIDO = 3
    RHConceptoTipo_FUNCIONES_INTERNAS = 4
End Enum

Public Enum RHTipoOpeEvaluacion
    RHTipoOpeEvaEscrito = 0
    RHTipoOpeEvaPsicologico = 1
    RHTipoOpeEvaEntrevista = 2
    RHTipoOpeEvaCurricular = 3
    RHTipoOpeEvaConsolidado = 4
End Enum

Public Enum RHContratoTipo
    RHContratoTipoIndeterminado = 0
    RHContratoTipoFijo = 1
    RHContratoTipoLocacion = 2
    RHContratoTipoFLaboral = 3
    RHContratoTipoPractica = 4
    RHContratoTipoSesigrista = 5
    RHContratoTipoDirector = 6
End Enum

Public Enum RHEmpleadoTurno
    RHEmpleadoTurnoUno = 1
    RHEmpleadoTurnoDos = 2
End Enum

Public Enum RHEstadosTpo
    RHEstadosTpoInactivo = 1
    RHEstadosTpoActivo = 2
    RHEstadosTpoVacaciones = 3
    RHEstadosTpoPermisosLicencias = 4
    RHEstadosTpoSubsidiado = 5
    RHEstadosTpoSuspendido = 6
    RHEstadosTpoRetirado = 7
End Enum

Public Enum RHPeriodoNoLab
    RHPeriodoNoLabSolicitado = 0
    RHPeriodoNoLabAprovado = 1
    RHPeriodoNoLabRechazado = 2
End Enum

Public Enum RHEmpleadoCuentasTpo
    RHEmpleadoCuentasTpoAhorro = 232
    RHEmpleadoCuentasTpoCTS = 234
End Enum

Public Enum RHExtraPlanillaOpeTpo
    RHExtraPlanillaOpeTpoCargo = 0
    RHExtraPlanillaOpeTpoAbono = 1
End Enum

Public Enum RHConceptosTpoVisible
    RHConceptosTpoVIngreso = 1
    RHConceptosTpoVEgreso = 2
    RHConceptosTpoVAportacion = 3
    RHConceptosTpoVVarUsuario = 4
    RHConceptosTpoVTodos = 5
End Enum

 

