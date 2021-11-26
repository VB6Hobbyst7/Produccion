VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecursosHumanos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
End Enum

Public Enum RHConceptoAportaciones
    gRHConcAportExtSolid = 301
    gRHConcAportSegSoc = 302
    gRHConcAportTotalAport = 303
End Enum

Public Enum RHConceptoVariablesUsuario
    gRHConcVarUsuCTSMesTrab = 401
    gRHConcVarUsuGratMesTrab = 402
    gRHConcVarUsuSubNumDias = 403
    gRHConcVarUsuPorcMesTrab = 404
End Enum

Public Enum RHConceptoVariablesGlobales
    gRHConcVarGlobTotDiasLab = 501
    gRHConcVarGlobPorcPagUTIL = 502
    gRHConcVarGlobSumTotIng = 503
    gRHConcVarGlobValUtil = 504
End Enum

Public Enum RHConceptoFuncionesConstantes
    gRHConcFuncConstAFPValComVar = 601
    gRHConcFuncConstAFPValorPrima = 602
    gRHConcFuncConstDctoJudiFijo = 603
    gRHConcFuncConstDctoJudiPorc = 604
    gRHConcFuncConst5taCat = 605
    gRHConcFuncConstAFP1023 = 606
    gRHConcFuncConstAFP0300 = 607
    gRHConcFuncConstCTSMesTrab = 608
    gRHConcFuncConstMinSal = 609
    gRHConcFuncConstMinTard = 610
    gRHConcFuncConstDiasnoTrab = 611
    gRHConcFuncConstNumMesTrab = 612
    gRHConcFuncConstTERNumMesTrab = 613
    gRHConcFuncConstPagoExeso = 614
    gRHConcFuncConstReintegro = 615
    gRHConcFuncConstSdoBas = 616
    gRHConcFuncConstSdoBasNivel = 617
    gRHConcFuncConstSdoContrato = 618
    gRHConcFuncConstUltBonVacac = 619
    gRHConcFuncConstUltGratif = 620
    gRHConcFuncConstUltTercio = 621
End Enum

Public Enum RHConceptoVariablesLocales
    gRHConcVarTotIngAcum = 701
    gRHConcVarAFPCuotaFija = 702
    gRHConcVarIngAntesAFP = 703
    gRHConcVarMontoAMP = 704
    gRHConcVarNumDiasLab = 705
    gRHConcVarNumDiasMes = 706
    gRHConcVarNumMinTrab = 707
    gRHConcVarPrueba = 708
    gRHConcVarPorcAsistMedFam = 709
    gRHConcVarPorcONP = 710
    gRHConcVarPorcFonavi = 711
    gRHConcVarPorc5taCat = 712
    gRHConcVarPorcSegSoc = 713
    gRHConcVarPorcExtSolid = 714
    gRHConcVarPorcIncONP = 715
    gRHConcVarPorcRPS = 716
    gRHConcVarPorcTercio = 717
    gRHConcVarPromDia = 718
    gRHConcVarSueldoNeto = 719
    gRHConcVarSumMonInt = 720
    gRHConcVarAFPTopePrima = 721
    gRHConcVarAPFTope = 722
    gRHConcVarTotalRemun = 723
    gRHConcVarUITValor = 724
    gRHConcVarUIT54 = 725
    gRHConcVarUIT07 = 726
    gRHConcVarDiaProm = 727
End Enum

Public Enum RHConceptoTipo
    gRHTpoConcConst = 1
    gRHTpoConcValCal = 2
    gRHTpoConcValPreDef = 3
    gRHTpoConcFunInt = 4
End Enum

Public Enum RHConceptoEstado
    gRHConcEstAct = 1
    gRHConcEstDesh = 0
End Enum

Public Enum RHImpuesto5taCat
    gRHImp5taCatCDcto = 1
    gRHImp5taCatSDcto = 0
End Enum

Public Enum RHAfectoMesTrabajado
    gRHAfeCtoMesTrabajado = 1
    gRHNoAfeCtoMesTrabajado = 0
End Enum

Public Enum RHProcesoSeleccionModal
    gRHProcSelModInterno = 0
    gRHProcSelModNuevo = 1
End Enum

