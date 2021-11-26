VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NDefOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit
Public Enum OpeNegPendientes
    gCapIngresoRegulaCG = 360101
End Enum

Public Enum OpeOpeCGAbonoNeg
    gCGPagProvAbonoDist = 201717
    gCGPagProvAbonoCent = 200207
    
    gCGArendirAtencAbonoDist = 201717
    gCGArendirAtencAbonoCent = 200207
    gCGArendirAtencCargoDist = 201817
    gCGArendirAtencCargoCent = 200207
End Enum

'-FONCODES
Public Enum OpeOpeCGAbonoNegFoncodes
    gCGPagProvAbonoFoncodesDist = 201702
    gCGPagProvAbonoFoncodesCent = 200210
End Enum
Public Enum OpeOpeCGCargoNegFoncodes
    gCGPagProvCargoFoncodesDist = 201802
    gCGPagProvCargoFoncodesCent = 200313
End Enum
'*****************************

Public Enum OpeCGMonedaExtranjera
    gOpeME = 400010
    gOpeMECompraAInst = 400011
    gOpeMEVentaAInst = 400012
    gOpeMECompraEfect = 400013
    gOpeMEVentaEfec = 400014

    gOpeMEExtCompraAInst = 400016
    gOpeMEExtVentaAInst = 400017
    gOpeMEExtCompraEfect = 400018
    gOpeMEExtVentaEfec = 400019
End Enum

Public Enum OpeCGArendirCta
    gCGArendirCtaMN = 401100
    gCGArendirCtaME = 402100
    gCGArendirCtaSolMN = 401110
    gCGArendirCtaSolME = 402110
    gCGArendirCtaRechMN = 401120
    gCGArendirCtaRechME = 402120
    
    '***Agregado por ELRO el 20120321, según OYP-RFC016-2012
    gCGArendirCtaSolcEditMN = 401124
    gCGArendirCtaSolcEditME = 402124
    gCGArendirCtaAprobMN = 401125
    gCGArendirCtaAprobME = 402125
    gCGArendirCtaExtAprobMN = 401176
    gCGArendirCtaExtAprobME = 402176
    gCGArendirCtatSust2MN = 401141
    gCGArendirCtatSust2ME = 402141
    gCGArendirCtaRend2MN = 401160
    gCGArendirCtaRend2ME = 402160
    gCGArendirCtaRendDsctPlanMN = 401167
    gCGArendirCtaRendDsctPlanME = 402167
    '***Fin Agregado por ELRO*******************************
    
    gCGArendirCtaAtencMN = 401130
    gCGArendirCtaAtencME = 402130
    
    gCGArendirCtaSustMN = 401140
    gCGArendirCtaSustME = 402140
    
    gCGArendirCtaRendMN = 401150
    gCGArendirCtaRendME = 402150
    gCGArendirCtaRendDevMN = 401161
    gCGArendirCtaRendDevME = 402161
    
    gCGArendirCtaExtAtencMN = 401171
    gCGArendirCtaExtAtencME = 402171
    
    gCGArendirCtaExtRendMN = 401180
    gCGArendirCtaExtRendME = 402180
    gCGArendirCtaRecxHonorarioProvBusca = 401190 'EJVG20140721
End Enum

Public Enum OpeCGArendirCtaAge
    gCGArendirCtaMNAge = 431100
    gCGArendirCtaMEAge = 432100
    gCGArendirCtaSolMNAge = 431110
    gCGArendirCtaSolMEAge = 432110
    gCGArendirCtaRechMNAge = 431120
    gCGArendirCtaRechMEAge = 432120
    
   
    gCGArendirCtaAtencMNAge = 431130
    gCGArendirCtaAtencMEAge = 432130
    
    gCGArendirCtaSustMNAge = 431140
    gCGArendirCtaSustMEAge = 432140
    
    gCGArendirCtaRendMNAge = 431150
    gCGArendirCtaRendMEAge = 432150
    gCGArendirCtaRendDevMNAge = 431161
    gCGArendirCtaRendDevMEAge = 432161
    
    gCGArendirCtaExtAtencMNAge = 431171
    gCGArendirCtaExtAtencMEAge = 432171
    
    gCGArendirCtaExtRendMNAge = 431180
    gCGArendirCtaExtRendMEAge = 432180
End Enum


Public Enum OpeCGArendirCtaViativo
    gCGArendirViatMN = 401200
    gCGArendirViatME = 402200
    
    gCGArendirViatSolMN = 401210
    gCGArendirViatSolME = 402210
    
    gCGArendirViatRechMN = 401220
    gCGArendirViatRechME = 402220
    
    '***Agregado por ELRO el 20120321, según OYP-RFC005-2012
    gCGArendirViatSolcEditMN = 401224
    gCGArendirViatSolcEditME = 402224
    gCGArendirViatAprobMN = 401225
    gCGArendirViatAprobME = 402225
    gCGArendirViatExtAprobMN = 401276
    gCGArendirViatExtAprobME = 402276
    gCGArendirViatSust2MN = 401241
    gCGArendirViatSust2ME = 402241
    gCGArendirViatRend2MN = 401260
    gCGArendirViatRend2ME = 402260
    gCGArendirViatRendDsctPlanMN = 401267
    gCGArendirViatRendDsctPlanME = 402267
    '***Fin Agregado por ELRO*******************************
    
    gCGArendirViatAtencMN = 401230
    gCGArendirViatAtencME = 402230
    
    gCGArendirViatSustMN = 401240
    gCGArendirViatSustME = 402240
    
    gCGArendirViatRendMN = 401250
    gCGArendirViatRendME = 402250
    
    gCGArendirViatRendDevMN = 401261
    gCGArendirViatRendDevME = 402261
    
    gCGArendirViatExtAtencMN = 401271
    gCGArendirViatExtAtencME = 402271
    
    gCGArendirViatExtRendMN = 401280
    gCGArendirViatExtRendME = 402280
    
    gCGArendirViatAmpMN = 401290
    gCGArendirViatAmpME = 402290
    gCGArendirViatRecxHonorarioProvBusca = 401291 'EJVG20140721
    
    '**Agregado por VAPI SEGUN ERS1792014
    gCGArendirViatReimpMN = 401292
    gCGArendirViatReimpME = 402292
    '**fin VAPI
    
End Enum
Public Enum OpeCajaChica
    gCHHabilitaNuevaMN = 401301
    gCHHabilitaNuevaME = 402301
    
    gCHMantenimientoMN = 401302
    gCHMantenimientoME = 402302
    
    gCHAutorizaDesembMN = 401310
    gCHAutorizaDesembME = 402310
    
    gCHDesembEfectivoMN = 401321
    gCHDesembEfectivoME = 402321
    
    gCHDesembOrdenPagoMN = 401322
    gCHDesembOrdenPagoME = 402322
    gCHDesembAbonoCtaMN = 401323
    gCHDesembAbonoCtaME = 402323
    
    gCHDesembCaChiCtaPendienteMN = 401324
    gCHDesembCaChiCtaPendienteME = 402324
    
    '***Agregado por ELRO según OYP-RFC047-2012
    gCHApropacionApeMN = 401325
    gCHApropacionApeME = 402325
    '***Fin Agregado por ELRO******************
    
    gCHExtDesembEfectivoMN = 401331
    gCHExtDesembEfectivoME = 402331
    
    gCHExtDesembOrdenPagoMN = 401332
    gCHExtDesembOrdenPagoME = 402332
    
    gCHExtDesembCtaPendienteMN = 401333
    gCHExtDesembCtaPendienteME = 402333
    
    '***Agregado por ELRO según OYP-RFC047-2012
    gCHExtApropacionApeMN = 401397
    gCHExtApropacionApeME = 402397
    '***Fin Agregado por ELRO******************
    
    gCHArendirCtaSolMN = 401341
    gCHArendirCtaSolME = 402341
    
    gCHArendirCtaRechMN = 401342
    gCHArendirCtaRechME = 402342
    
    gCHArendirCtaAtencMN = 401343
    gCHArendirCtaAtencME = 402343
    
    gCHArendirCtaSustMN = 401344
    gCHArendirCtaSustME = 402344
    
    gCHArendirCtaRendMN = 401350
    gCHArendirCtaRendME = 402350
    
    gCHArendirCtaRendExacMN = 401351
    gCHArendirCtaRendExacME = 402351
    
    gCHArendirCtaRendIngMN = 401352
    gCHArendirCtaRendIngME = 402352
    
    gCHArendirCtaRendEgreMN = 401353
    gCHArendirCtaRendEgreME = 402353
    
  
    gCHArendirCtaRendExtMN = 401360
    gCHArendirCtaRendExtME = 402360
    
    '***Modificado por ELRO según OYP-RFC047-2012
    gCHArendirCtaRendExtMN2 = 401365
    gCHArendirCtaRendExtME2 = 402365
    '***Fin Modificado por ELRO******************
    
    gCHArendirCtaRendExtAtencMN = 401361
    gCHArendirCtaRendExtAtencME = 402361
    
    gCHArendirCtaRendExtExactMN = 401362
    gCHArendirCtaRendExtExactME = 402362
    
    gCHArendirCtaRendExtIngMN = 401363
    gCHArendirCtaRendExtIngME = 402363
    
    gCHArendirCtaRendExtEgrMN = 401364
    gCHArendirCtaRendExtEgrME = 402364
    
    gCHEgreDirectoSolMN = 401371
    gCHEgreDirectoSolME = 402371
    
    gCHEgreDirectoRechMN = 401372
    gCHEgreDirectoRechME = 402372
    
    gCHEgreDirectoAtencMN = 401373
    gCHEgreDirectoAtencME = 402373
    
    gCHEgreDirectoExtAtencMN = 401374
    gCHEgreDirectoExtAtencME = 402374
    
    gCHRendContabMN = 401380
    gCHRendContabME = 402380
    
    gCHArqueoContabMN = 401390
    gCHArqueoContabME = 402390
    
    gCHRendContabExtMN = 401396
    gCHRendContabExtME = 402396
    gCHRendContabRecxHonorarioProvBusca = 401398 'EJVG20140721
End Enum

Public Enum OpeBovedaCajaGeneral
    gOpeBoveCGMN = 401400
    gOpeBoveCGME = 402400
    
    gOpeBoveCGHabAgeMN = 401401
    gOpeBoveCGHabAgeME = 402401
    
    gOpeBoveCGConfHabAgeBoveMN = 401402
    gOpeBoveCGConfHabAgeBoveME = 402402

    gOpeBoveCGExtHabAgeMN = 401404
    gOpeBoveCGExtHabAgeME = 402404
    
    gOpeBoveCGExtConfHabAgeBovMN = 401405
    gOpeBoveCGExtConfHabAgeBovME = 402405
End Enum

Public Enum OpeBovedaAgencia
    gOpeBoveAgeMM = 901001
    gOpeBoveAgeME = 902001
    
    gOpeBoveAgeConfHabCGMN = 901002
    gOpeBoveAgeConfHabCGME = 902002
    
    gOpeBoveAgeHabAgeACGMN = 901003
    gOpeBoveAgeHabAgeACGME = 902003
    
    gOpeBoveAgeHabEntreAgeMN = 901004
    gOpeBoveAgeHabEntreAgeME = 902004
    
    gOpeBoveAgeHabCajeroMN = 901005
    gOpeBoveAgeHabCajeroME = 902005
    
    gOpeBoveAgeExtConfHabCGME = 902007
    
    gOpeBoveAgeExtHabAgeACGMN = 901008
    gOpeBoveAgeExtHabAgeACGME = 902008
    
    gOpeBoveAgeExtHabEntreAgeMN = 901009
    gOpeBoveAgeExtHabEntreAgeMe = 902009
    
    gOpeBoveAgeExtHabCajeroMN = 901010
    gOpeBoveAgeExtHabCajeroME = 902010
    
End Enum
Public Enum OpeCajeroMonedaExtranjera
    gOpeCajeroME = 900020
    gOpeCajeroMETipoCambio = 900021
    gOpeCajeroMECompra = 900022
    gOpeCajeroMEVenta = 900023
    gOpeCajeroMEExtCompra = 900025
    gOpeCajeroMEExtVenta = 900026
End Enum

Public Enum OpeHabCajero
    gOpeHabCajeroMN = 901011
    gOpeHabCajeroME = 902011

    gOpeHabCajDevABoveMN = 901012
    gOpeHabCajDevABoveME = 902012
    
    gOpeHabCajTransfEfectCajerosMN = 901013
    gOpeHabCajTransfEfectCajerosME = 902013
    
    gOpeHabCajRegSobFaltMN = 901014
    gOpeHabCajRegSobFaltME = 902014

    gOpeHabCajIngEfectRegulaFaltMN = 901015
    gOpeHabCajIngEfectRegulaFaltME = 902015
    
    'Operaciones de Declaracion de Efectivo en una Agencia al cierre de dia
    gOpeHabCajRegEfect = 901007
    gOpeHabBoveRegEfect = 901016
    '----------
    
    gOpeHabCajConfHabBovAgeMN = 901017
    gOpeHabCajConfHabBovAgeME = 902017
    
    gOpeHabCajExtDevABoveMN = 901022
    gOpeHabCajExtDevABoveME = 902022

    gOpeHabCajExtTransfEfectCajerosMN = 901023
    gOpeHabCajExtTransfEfectCajerosME = 902023
    
    gOpeHabCajExtRegSobFaltMN = 901024
    gOpeHabCajExtRegSobFaltME = 902024

    gOpeHabCajExtIngEfectRegulaFaltMN = 901025
    gOpeHabCajExtIngEfectRegulaFaltME = 902025
    
    gOpeHabCajExtRegEfectMN = 901026
    gOpeHabCajExtRegEfectME = 902026
    
    gOpeHabCajExtConfHabBovAgeMN = 901027
    gOpeHabCajExtConfHabBovAgeME = 902027
    
End Enum

Public Enum OpeCajeroVarias
    gOpeCajeroVarSEDALIBMN = 901032
    gOpeCajeroVarHIDRANDINAMN = 901033
    gOpeCajeroVarTELEFONICAMN = 901034
    gOpeCajeroVarExtSEDALIBMN = 901036
    gOpeCajeroVarExtHIDRANDINAMN = 901037
    gOpeCajeroVarExtTELEFONICAMN = 901038
    
    gOpeCajeroVarSEDALIBME = 902032
    gOpeCajeroVarHIDRANDINAME = 902033
    gOpeCajeroVarTELEFONICAME = 902034
    gOpeCajeroVarExtSEDALIBME = 902036
    gOpeCajeroVarExtHIDRANDINAME = 902037
    gOpeCajeroVarExtTELEFONICAME = 902038
    
End Enum
Public Enum OpeCheques
    gOpeChequesAnulacion = 900031
    gOpeChequesRechazo = 900032
    gOpeChequesValorización = 900033
    gOpeChequesModFecVal = 900034
    gOpeChequesConsEstados = 900035
    gOpeChequesExtAnulación = 900037
    gOpeChequesExtRechazo = 900038
    gOpeChequesExtValorización = 900039
End Enum

Public Enum OpeContabilidad
    gContRegistroAsiento = 701101
    gContAjusteTipoCambio = 701201
    gContAjusteInflaIngre = 701202
    gContAjusteInflaHisto = 701203
    gContAjusteDespreciac = 701204
    gContCierreAnual = 701850
End Enum

Public Enum OpeCGRegulaVentanilla
    gOpeCGRVentanaIngresoMN = 401411
    gOpeCGRVentanaEgresoMN = 401412
    gOpeCGRVentanaIngresoME = 402411
    gOpeCGRVentanaEgresoME = 402412
    gOpeCGRVentRegTransferMN = 401413
    gOpeCGRVentRegTransferME = 402413
    'Extornos
    gOpeCGRVentanaIngresoMNExt = 401416
    gOpeCGRVentanaEgresoMNExt = 401417
    gOpeCGRVentanaIngresoMEExt = 402416
    gOpeCGRVentanaEgresoMEExt = 402417
    gOpeCGRVentRegTransferMNExt = 401418
    gOpeCGRVentRegTransferMEExt = 402418
End Enum

Public Enum OpeCGTransferencias
    gOpeCGTransfMN = 401420
    gOpeCGTransfBancosMN = 401421
    gOpeCGTransfBancosCMACSMN = 401422
    gOpeCGTransfCMACSBancosMN = 401423
    gOpeCGTransfMismoBancoMN = 401424
    
    gOpeCGTransfME = 402420
    gOpeCGTransfBancosME = 402421
    gOpeCGTransfBancosCMACSME = 402422
    gOpeCGTransfCMACSBancosME = 402423
    gOpeCGTransfMismoBancoME = 402424
    
    gOpeCGTransfExtBancosMN = 401431
    gOpeCGTransfExtCMACSBancosMN = 401432
    gOpeCGTransfExtBancosCMACSMN = 401433
    gOpeCGTransfExtMismoBancoMN = 401434

    gOpeCGTransfExtBancosME = 402431
    gOpeCGTransfExtCMACSBancosME = 402432
    gOpeCGTransfExtBancosCMACSME = 402433
    gOpeCGTransfExtMismoBancoME = 402434
End Enum

Public Enum OpeCGOpeBancos
    gOpeCGOpeBancosMN = 401500
    gOpeCGOpeBancosDepEfecMN = 401502
    gOpeCGOpeBancosRetEfecMN = 401503
    gOpeCGOpeBancosConfRetEfecMN = 401504
    gOpeCGOpeBancosRegChequesMN = 401505
    gOpeCGOpeBancosDepChequesMN = 401506
    gOpeCGOpeBancosDepDivBancosMN = 401507
    gOpeCGOpeBancosRetDivBancosMN = 401508
    gOpeCGOpeBancosRecibeChqAgMn = 401509
    gOpeCGOpeBancosRemesaAgenciaMN = 401571 'EJVG20140626
    gOpeCGOpeBancosRemesaVBHabilitacion = 401572 'EJVG20140711
    gOpeCGOpeBancosExtRemesaAgencia = 401574 'EJVG20140811
    'JEOM  MN
    gOpeCGOpeBancosRetCtasBancosDetracMN = 401581
    gOpeCGOpeBancosRetCtasBancosEmbargoMN = 401582
    gOpeCGOpeBancosOtrosDepositosMN = 401583
    gOpeCGOpeBancosOtrosRetirosMN = 401584
    gOpeCGOpeBancosRetFielCumplimientoMN = 401585
    'gOpeCGOpeBancosRemesaAgenciaMN = 401587 'EJVG20140509 'COMENTADO POR FRHU SOLICITADO POR PASI 20140610
    gOpeCGOpeBancosDepositoXActivacionSegTarjetaMN = 401587 'FRHU 20140522 ERS068-2014 RQ14283
    gOpeCGOpeBancosRetPagoSegTarjMN = 401588 'JUEZ 20140711
    gOpeCGOpeBancosDepComSegTarjMN = 401589 'JUEZ 20140711
'    gOpeCGOpeBancosRetCtasBancosEmbargoMNExt = 401588
'    gOpeCGOpeBancosRetCtasBancosDetracMNExt = 401589
    'Extornos MN
    gOpeCGOpeBancosRetCtasBancosEmbargoMNExt = 401591
    gOpeCGOpeBancosRetCtasBancosDetracMNExt = 401592
    gOpeCGOpeBancosOtrosDepositosMNExt = 401593
    gOpeCGOpeBancosOtrosRetirosMNExt = 401594
    gOpeCGOpeBancosRetFielCumplimientoMNExt = 401595
    gOpeCGOpeBancosDepositoXActivacionSegTarjetaMNExt = 401597 'FRHU 20140526 ERS068-2014 RQ14284
    gOpeCGOpeBancosRetPagoSegTarjMNExt = 401598 'JUEZ 20140711
    gOpeCGOpeBancosDepComSegTarjMNExt = 401599 'JUEZ 20140711
    
    
    gOpeCGOpeBancosME = 401500
    gOpeCGOpeBancosDepEfecME = 402502
    gOpeCGOpeBancosRetEfecME = 402503
    gOpeCGOpeBancosConfRetEfecME = 402504
    gOpeCGOpeBancosRegChequesME = 402505
    gOpeCGOpeBancosDepChequesME = 402506
    gOpeCGOpeBancosDepDivBancosME = 402507
    gOpeCGOpeBancosRetDivBancosME = 402508
    gOpeCGOpeBancosRecibeChqAgMe = 402509
    gOpeCGOpeBancosRemesaAgenciaMe = 402571 'EJVG20140626
    gOpeCGOpeBancosRemesaVBHabilitacionme = 402572 'EJVG20140711
    gOpeCGOpeBancosExtRemesaAgenciaME = 402574 'EJVG20140811
    'JEOM ME
    gOpeCGOpeBancosRetCtasBancosDetracME = 402581
    gOpeCGOpeBancosRetCtasBancosEmbargoME = 402582
    gOpeCGOpeBancosOtrosDepositosME = 402583
    gOpeCGOpeBancosOtrosRetirosME = 402584
    gOpeCGOpeBancosRetFielCumplimientoME = 402585
    
    
'    gOpeCGOpeBancosRetCtasBancosDetracMEExt = 402589
'    gOpeCGOpeBancosRetCtasBancosEmbargoMEExt = 402588
    'Extornos ME
    gOpeCGOpeBancosRetCtasBancosEmbargoMEExt = 402591
    gOpeCGOpeBancosRetCtasBancosDetracMEExt = 402592
    gOpeCGOpeBancosOtrosDepositosMEExt = 402593
    gOpeCGOpeBancosOtrosRetirosMEExt = 402594
    gOpeCGOpeBancosRetFielCumplimientoMEExt = 402595
    

    gOpeCGOpeAperCorrienteMN = 401512
    gOpeCGOpeAperAhorroMN = 401513
    gOpeCGOpeAperPlazoMN = 401514
    
    gOpeCGOpeAperCorrienteME = 402512
    gOpeCGOpeAperAhorroME = 402513
    gOpeCGOpeAperPlazoME = 402514
    
    gOpeCGOpeConfApertMN = 401515
    gOpeCGOpeIntDevPFMN = 401516
    gOpeCGOpeGastComBancosMN = 401517
    gOpeCGOpeCapIntPFMN = 401518
    gOpeCGOpeCancCtaCteMN = 401520
    gOpeCGOpeCancCtaAhoMN = 401521
    gOpeCGOpeCancCtaPFMN = 401522
    gOpeCGOpeMantCtaBancosMN = 401523
    gOpeCGOpeInteresAhoCtaCteMN = 401526
    gOpeCGOpeMantCtasBancosMN = 401540
    gOpeBancoyCajasMN = 401541 'ALPA 20110429
    gOpeBancoyCajasME = 402541 'ALPA 20110429
    
    gOpeCGOpeMantCtasBancoshabMN = 401578
    gOpeCGOpeMantCtasBancoshabME = 402578
    
    gOpeCGOpeMantCtasBancosReMN = 401579
    gOpeCGOpeMantCtasBancosReME = 402579
    
    gOpeCGOpeConfApertME = 402515
    gOpeCGOpeIntDevPFME = 402516
    gOpeCGOpeGastComBancosME = 402517
    gOpeCGOpeCapIntPFME = 402518
    gOpeCGOpeCancCtaCteME = 402520
    gOpeCGOpeCancCtaAhoME = 402521
    gOpeCGOpeCancCtaPFME = 402522
    gOpeCGOpeMantCtaBancosME = 402523
    gOpeCGOpeInteresAhoCtaCteME = 402526
    gOpeCGOpeMantCtasBancosME = 402540
    
    gOpeCGOpeMantCtasBancosHaMN = 401577
    gOpeCGOpeMantCtasBancosHaME = 402577
    

    'Extornos de Bancos
    gOpeCGExtBancos = 401550
    gOpeCGExtBcoDepRet = 401551
    gOpeCGExtBcoDepEfectivo = 401552
    gOpeCGExtBcoRetEfectivo = 401553
    gOpeCGExtBcoConfRetEfectivo = 401554
    gOpeCGExtBcoRegCheques = 401555
    gOpeCGExtBcoDepCheques = 401556
    gOpeCGExtBcoDepDiv = 401557
    gOpeCGExtBcoRetDiv = 401558
    gOpeCGExtBcoRecepChqRegAgencias = 401559
    gOpeCGExtBcoOpeCtas = 401560
    gOpeCGExtBcoApertCta = 401561
    gOpeCGExtBcoConfApert = 401565
    gOpeCGExtBcoIntDevengPF = 401566
    gOpeCGExtBcoGastComision = 401567
    gOpeCGExtBcoCapitalizaIntDPF = 401568
    gOpeCGExtBcoCancelaCtas = 401569
    gOpeCGExtBcoIntCtasAho = 401576

    'Extornos de Bancos ME
    gOpeCGExtBancosME = 402550
    gOpeCGExtBcoDepRetME = 402551
    gOpeCGExtBcoDepEfectivoME = 402552
    gOpeCGExtBcoRetEfectivoME = 402553
    gOpeCGExtBcoConfRetEfectivoME = 402554
    gOpeCGExtBcoRegChequesME = 402555
    gOpeCGExtBcoDepChequesME = 402556
    gOpeCGExtBcoDepDivME = 402557
    gOpeCGExtBcoRetDivME = 402558
    gOpeCGExtBcoRecepChqRegAgenciasME = 402559
    gOpeCGExtBcoOpeCtasME = 402560
    gOpeCGExtBcoApertCtaME = 402561
    gOpeCGExtBcoConfApertME = 402565
    gOpeCGExtBcoIntDevengPFME = 402566
    gOpeCGExtBcoGastComisionME = 402567
    gOpeCGExtBcoCapitalizaIntDPFME = 402568
    gOpeCGExtBcoCancelaCtasME = 402569
    gOpeCGExtBcoIntCtasAhoME = 402576

End Enum

Public Enum OpeCGOpeCMAC
    gOpeCGOpeCMACDepDivMN = 401602
    gOpeCGOpeCMACRetDivMN = 401603
    gOpeCGOpeCMACRegularizMN = 401604
    gOpeCGOpeCMACAperAhorrosMN = 401612
    gOpeCGOpeCMACAperPFMN = 401613
    gOpeCGOpeCMACConfAperMN = 401614
    gOpeCGOpeCMACIntDevPFMN = 401615
    gOpeCGOpeCMACGastosComMN = 401616
    gOpeCGOpeCMACCapIntDevPFMN = 401617
    gOpeCGOpeCMACCancAhorrosMN = 401619
    gOpeCGOpeCMACCancPFMN = 401620
    gOpeCGOpeCMACMantCtas1MN = 401640
    gOpeCGOpeCMACInteresAhoMN = 401626
    
    gOpeCGOpeCMACDepDivME = 402602
    gOpeCGOpeCMACRetDivME = 402603
    gOpeCGOpeCMACRegularizME = 402604
    gOpeCGOpeCMACAperAhorrosME = 402612
    gOpeCGOpeCMACAperPFME = 402613
    gOpeCGOpeCMACConfAperME = 402614
    gOpeCGOpeCMACIntDevPFME = 402615
    gOpeCGOpeCMACGastosComME = 402616
    gOpeCGOpeCMACCapIntDevPFME = 402617
    gOpeCGOpeCMACCancAhorrosME = 402619
    gOpeCGOpeCMACCancPFME = 402620
    gOpeCGOpeCMACMantCtas1ME = 402640
    gOpeCGOpeCMACInteresAhoME = 402626
    
    gOpeCGOpeCMACDepDivMNExt = 401652
    gOpeCGOpeCMACRetDivMNExt = 401653
    gOpeCGOpeCMACRegularizMNExt = 401654
    gOpeCGOpeCMACAperCtasMNExt = 401661
    gOpeCGOpeCMACConfAperMNExt = 401664
    gOpeCGOpeCMACIntDevPFMNExt = 401665
    gOpeCGOpeCMACGastosComMNExt = 401666
    gOpeCGOpeCMACCapIntDevPFMNExt = 401667
    gOpeCGOpeCMACCancelaMNExt = 401668
    gOpeCGOpeCMACInteresAhoMNExt = 401676
    
    gOpeCGOpeCMACDepDivMEExt = 402652
    gOpeCGOpeCMACRetDivMEExt = 402653
    gOpeCGOpeCMACRegularizMEExt = 402654
    gOpeCGOpeCMACAperCtasMEExt = 402661
    gOpeCGOpeCMACConfAperMEExt = 402664
    gOpeCGOpeCMACIntDevPFMEExt = 402665
    gOpeCGOpeCMACGastosComMEExt = 402666
    gOpeCGOpeCMACCapIntDevPFMEExt = 402667
    gOpeCGOpeCMACCancelaMEExt = 402668
    gOpeCGOpeCMACInteresAhoMEExt = 402676
End Enum

Public Enum OpeCGAdeudados
    gOpeCGAdeudadosMN = 401800             'ADEDUADOS MN
    gOpeCGAdeudaCalendarioMN = 401801      'Simulacion de Cronograma de Pagos
    gOpeCGAdeudaRegPagareMN = 401803       'Registro Pagarés MN
    gOpeCGAdeudaRegPagareConfMN = 401804   'Confirmacion de Pagaré
    gOpeCGAdeudaProvisionMN = 401805         'Provisión Adeudados MN
    gOpeCGAdeudaPagoCuotaMN = 401806       'Pago Adeudados MN
    gOpeCGAdeudaReprogramaMN = 401807      'Reprogramación Adeudados MN
    gOpeCGAdeudaExtornosMN = 401808        'Extornos Adeudados MN
    gOpeCGAdeudaExtRegistroMN = 401809     'Extornos Registro Pagarés Bancos MN
    gOpeCGAdeudaExtConfRegiMN = 401810     'Extornos Confirmación Pagare MN
    gOpeCGAdeudaExtProvisiónMN = 401811    'Extornos Provisión Adeudados Bancos MN
    gOpeCGAdeudaExtPagoCuotaMN = 401812    'Extornos Pago Adeudados Bancos MN
    gOpeCGAdeudaExtReprogramaMN = 401813   'Extornos Reprogramación Adeudados Bancos MN
    gOpeCGAdeudaMntPagaresMN = 401832      'Mantenimiento de Pagarés MN
    gOpeCGAdeudaMntAjusteEurosMN = 401834      'AJUSTE DE ADEUDOS EN EUROS
    gOpeCGAdeudaMntPagaresVinculadosMN = 401833      'Mantenimiento de Pagarés Vinculados MN
    
    gOpeCGAdeudaPresuProyecMN = 401901     'Proyecciones de Adeudados MN
    gOpeCGAdeudaPagoCuotaLoteMN = 401840   'Pago Adeudados en Lote MN
    
    gOpeCGAdeudadosME = 402800             'ADEDUADOS ME
    gOpeCGAdeudaCalendarioME = 402801      'Simulacion de Cronograma de Pagos
    gOpeCGAdeudaRegPagareME = 402803       'Registro Pagarés ME
    gOpeCGAdeudaRegPagareConfMe = 402804   'Confirmacion de Pagaré
    gOpeCGAdeudaProvisionME = 402805         'Provisión Adeudados ME
    gOpeCGAdeudaPagoCuotaME = 402806       'Pago Adeudados ME
    gOpeCGAdeudaReprogramaME = 402807      'Reprogramación Adeudados ME
    gOpeCGAdeudaExtornosME = 402808        'Extornos Adeudados ME
    gOpeCGAdeudaExtRegistroME = 402809     'Extornos Registro Pagarés Bancos ME
    gOpeCGAdeudaExtConfRegiME = 402810     'Extornos Confirmación Pagare ME
    gOpeCGAdeudaExtProvisiónME = 402811    'Extornos Provisión Adeudados Bancos ME
    gOpeCGAdeudaExtPagoCuotaME = 402812    'Extornos Pago Adeudados Bancos ME
    gOpeCGAdeudaExtReprogramaME = 402813   'Extornos Reprogramación Adeudados Bancos ME
    gOpeCGAdeudaMntPagaresME = 402832      'Mantenimiento de Pagarés ME
    gOpeCGAdeudaMntPagaresVinculadosME = 402833      'Mantenimiento de Pagarés Vinculados ME
    gOpeCGAdeudaMntAjusteEurosME = 402834      'AJUSTE DE ADEUDOS EN EUROS
    gOpeCGAdeudaPresuProyecME = 402901     'Proyecciones de Adeudados ME
    gOpeCGAdeudaPagoCuotaLoteME = 402840   'Pago Adeudados en Lote ME
End Enum


Public Enum OpeCGReportes
    OpeCGRepFlujoDiarioResMN = 461001
    OpeCGRepFlujoDiarioDetMN = 461002
    OpeCGRepRepBancosMN = 461010
    OpeCGRepRepBancosFlujoMensMN = 461011
    OpeCGRepRepBancosFlujoPFMN = 461012
    OpeCGRepRepBancosResumenPFMN = 461013
    OpeCGRepRepBancosConcentFdos = 461014
    
    OpeCGRepRepCMACSMN = 461020
    OpeCGRepRepCMACSFlujoMensMN = 461021
    OpeCGRepRepCMACSFlujoPFMN = 461022
    OpeCGRepRepCMACSResumenPFMN = 461023
    OpeCGRepProyecNroCliente = 461024 'ALPA 20090316
    'Ordenes de Pago
    OpeCGRepRepOPMN = 461030
    OpeCGRepRepOPGirMN = 461031
    'A Rendir Cuenta
    OpeCGRepArendirLibroAuxMN = 461041
    OpeCGRepArendirPendienteMN = 461042
    OpeCGRepArendirViaticosMN = 461043
    OpeCGRepArendirViaticoLibroAuxMN = 461044
    
    'Cheques
    OpeCGRepRepChqMN = 461100
    OpeCGRepRepChqRecMN = 461110
    OpeCGRepRepChqRecDetMN = 461111
    OpeCGRepRepChqRecResMN = 461112
    OpeCGRepRepChqValMN = 461120
    OpeCGRepRepChqValDetMN = 461121
    OpeCGRepRepChqValResMN = 461122
    OpeCGRepRepChqValorizadosMN = 461130
    OpeCGRepRepChqValorizadosDetMN = 461131
    OpeCGRepRepChqValorizadosResMN = 461132
    OpeCGRepRepChqAnulMN = 461140
    OpeCGRepRepChqAnulDetMN = 461141
    OpeCGRepRepChqAnulResMN = 461142
    OpeCGRepRepChqObsMN = 461150
    OpeCGRepRepChqObsDetMN = 461151
    OpeCGRepRepChqObsResMN = 461152
    OpeCGRepChqRecibidoCajaMN = 461160
    OpeCGRepChqDepositadoMN = 461170
    OpeCGRepChqPorCobrarConvMN = 461240 '*** PEAC 20100510
    
    'Cheques Remesados
    OpeCGRepChequesEnvMN = 461210
    OpeCGRepChequesAnulMN = 461220
    OpeCGRepChequesCobMN = 461230
    
    
    OpeCGAdeudRepGeneralMN = 461091
    OpeCGAdeudRepDetalleMN = 461092
    OpeCGAdeudRepSaldLinFinancDescalceMN = 461093
    OpeCGAdeudRepCortoPlazoMN = 461094
    OpeCGAdeudRepxFecVenc = 461095
    OpeCGAdeudRepVinculadosMN = 461096

    
    OpeCGAdeudRepGeneralME = 462091
    OpeCGADeudRepDetalleME = 462092
    OpeCGADeudRepDetalleCofideME = 462093
    OpeCGAdeudRepCortoPlazoME = 462094
    OpeCGAdeudRepxFecVencME = 462095
    OpeCGAdeudRepVinculadosME = 462096

'Flujo diario de Caja
    OpeCGRepSaldoBcos = 461701
    OpeCGRepSaldoCajAge = 461702
    
'Reportes al Exterior
    gOpeRptTripleJump = 461801 'EJVG20111014
    gOpeRptGtoFinanProd = 461802 'EJVG20111021
    gOpeRptConcentracFondos = 461803 'EJVG20121002
    gOpeRptELF = 461804 'ALPA20120117
    gOpRepCOFIDE = 461805 'MIOL20130627 RQ13329
    
'MONEDA EXTRANJERA
    OpeCGRepFlujoDiarioResME = 462001
    OpeCGRepFlujoDiarioDetME = 462002
    OpeCGRepRepBancosME = 462010
    OpeCGRepRepBancosFlujoMensME = 462011
    OpeCGRepRepBancosFlujoPFME = 462012
    OpeCGRepRepBancosResumenPFME = 462013
    OpeCGRepRepCMACSME = 462020
    OpeCGRepRepCMACSFlujoMensME = 462021
    OpeCGRepRepCMACSFlujoPFME = 462022
    OpeCGRepRepCMACSResumenPFME = 462023
    'Ordenes de Pago
    OpeCGRepRepOPME = 462030
    OpeCGRepRepOPGirME = 462031
    'A Rendir Cuenta
    OpeCGRepArendirLibroAuxME = 462041
    OpeCGRepArendirPendienteME = 462042
    OpeCGRepArendirViaticosME = 462043
    OpeCGRepArendirViaticoLibroAuxME = 462044
    OpeCGRepAdeudadoCalendarioMN = 461097
    OpeCGRepAdeudadoCalendarioME = 462097
    
    OpeCGRepAdeudadoCalendarioVigenteMN = 461098
    OpeCGRepAdeudadoCalendarioVigenteME = 462098
    
    OpeCGRepAnalisisDeCtaMN = 461099
    OpeCGRepAnalisisDeCtaME = 462099
    
        
    'Cheques
    OpeCGRepRepChqME = 462100
    OpeCGRepRepChqRecME = 462110
    OpeCGRepRepChqRecDetME = 462111
    OpeCGRepRepChqRecResME = 462112
    OpeCGRepRepChqValME = 462120
    OpeCGRepRepChqValDetME = 462121
    OpeCGRepRepChqValResME = 462122
    OpeCGRepRepChqValorizadosME = 462130
    OpeCGRepRepChqValorizadosDetME = 462131
    OpeCGRepRepChqValorizadosResME = 462132
    OpeCGRepRepChqAnulME = 462140
    OpeCGRepRepChqAnulDetME = 462141
    OpeCGRepRepChqAnulResME = 462142
    OpeCGRepRepChqObsME = 462150
    OpeCGRepRepChqObsDetME = 462151
    OpeCGRepRepChqObsResME = 462152
    OpeCGRepChqRecibidoCajaME = 462160
    OpeCGRepChqDepositadoME = 462170
    OpeCGRepChqPorCobrarConvME = 462240 '*** PEAC 20100510
    'Cheques Remesados
    OpeCGRepChequesEnvME = 462210
    OpeCGRepChequesAnulME = 462220
    OpeCGRepChequesCobME = 462230
    
End Enum

Public Enum OpeCGRepPresupuesto
    OpeCGRepPresuFlujoCaja = 461501
    OpeCGRepPresuServDeuda = 461502
    OpeCGRepPresuFinancia = 461503
    
    OpeCGRepPresuFlujoCajaME = 462501
    OpeCGRepPresuServDeudaME = 462502
    OpeCGRepPresuFinanciaME = 462503
End Enum

Public Enum OpeCGPagoSunat
   OpeCGPagoSunat = 411100
   OpeCGPagoSunatSegSocial = 411101
   OpeCGPagoSunatBoletaPag = 411102
   OpeCGPagoSunatIGVRenta = 411103
   
   OpeCGPagoSunatExt = 411105
   OpeCGPagoSunatSegSocialExt = 411106
   OpeCGPagoSunatBoletaPagExt = 411107
   OpeCGPagoSunatIGVRentaExt = 411108

End Enum

Public Enum OpeCGCartaFianza
   OpeCGCartaFianza = 411110
   OpeCGCartaFianzaIng = 411111
   OpeCGCartaFianzaSal = 411112
   OpeCGCartaFianzaME = 412110
   OpeCGCartaFianzaIngME = 412111
   OpeCGCartaFianzaSalME = 412112
   
   OpeCGCartaFianzaRepIngreso = 461401
   OpeCGCartaFianzaRepSalida = 461402
   OpeCGCartaFianzaRepIngresoME = 462401
   OpeCGCartaFianzaRepSalidaME = 462402

   OpeCGCartaFianzaExt = 411115
   OpeCGCartaFianzaIngExt = 411116
   OpeCGCartaFianzaSalExt = 411117
   OpeCGCartaFianzaMEExt = 412115
   OpeCGCartaFianzaIngMEExt = 412116
   OpeCGCartaFianzaSalMEExt = 412117

End Enum

Public Enum OpeCGOtrosOpeEfectivo
   OpeCGOtrosOpeEfectivo = 421000
   OpeCGOtrosOpeEfecIngr = 421001   'Ingresos de Efectivo
   OpeCGOtrosOpeEfecEgre = 421002   'Egresos de Efectivo
   OpeCGOtrosOpeEfecCamb = 421003   'Cambio de Billetaje
   OpeCGOtrosOpeChequeValorizacion = 421004   'EJVG20140125
   OpeCGOtrosOpeChequeRechazo = 421005   'EJVG20140129
   OpeCGOtrosOpeEfecOtro = 421009   'Otrs Operaciones
   OpeCGOtrosOpeRetPagSeguroDesgravamenMN = 421006 'PASIERS1362014
   OpeCGOtrosOpeRetPagSeguroIncendioMN = 421007 'PASIERS1362014
   
   
   OpeCGOtrosOpeEfectivoME = 422000
   OpeCGOtrosOpeEfecIngrME = 422001   'Ingresos de Efectivo
   OpeCGOtrosOpeEfecEgreME = 422002   'Egresos de Efectivo
   OpeCGOtrosOpeEfecCambme = 422003   'Cambio de Billetaje
   OpeCGOtrosOpeChequeValorizacionME = 422004   'EJVG20140125
   OpeCGOtrosOpeChequeRechazoME = 422005   'EJVG20140129 'PASI20140611 Cambio del valor 421005 por 422005
   OpeCGOtrosOpeEfecOtroME = 422009   'Otrs Operaciones
   OpeCGOtrosOpeRetPagSeguroDesgravamenME = 422006 'PASIERS1362014
   OpeCGOtrosOpeRetPagSeguroIncendioME = 422007 'PASIERS1362014
   
   OpeCGOtrosOpeEfectivoExt = 421010
   OpeCGOtrosOpeEfecIngrExt = 421011   'Ext Ingresos de Efectivo
   OpeCGOtrosOpeEfecEgreExt = 421012   'Ext Egresos de Efectivo
   OpeCGOtrosOpeEfecCambExt = 421013   'Ext Cambio de Billetaje
   OpeCGOtrosOpeChequeExtValorizacion = 421014   'EJVG20140314
   OpeCGOtrosOpeChequeExtRechazo = 421015   'EJVG20140314
   OpeCGOtrosOpeRetPagSeguroDesgravamenMNExt = 421016 'PASIERS1362014
   OpeCGOtrosOpeRetPagSeguroIncendioMNExt = 421017 'PASIERS1362014
   
   
   OpeCGOtrosOpeEfecOtroExt = 421019   'Ext Otrs Operaciones
    
   OpeCGOtrosOpeEfectivoMEExt = 422010
   OpeCGOtrosOpeEfecIngrMEExt = 422011   'Ext Ingresos de Efectivo
   OpeCGOtrosOpeEfecEgreMEExt = 422012   'Ext Egresos de Efectivo
   OpeCGOtrosOpeEfecCambMEExt = 422013   'Ext Cambio de Billetaje
   OpeCGOtrosOpeChequeExtValorizacionME = 422014   'EJVG20140314
   OpeCGOtrosOpeChequeExtRechazoME = 422015   'EJVG20140314
   OpeCGOtrosOpeEfecOtroMEExt = 422019   'Ext Otrs Operaciones
   OpeCGOtrosOpeRetPagSeguroDesgravamenMEExt = 422016  'PASIERS1362014
   OpeCGOtrosOpeRetPagSeguroIncendioMEExt = 422017  'PASIERS1362014

End Enum

Public Enum OpeCGOpeProveedores
   OpeCGOpeProveedor = 421100
   OpeCGOpeProvPago = 421110
   OpeCGOpeProvPagoListSUNAT = 421118
   OpeCGOpeProvPagoRegSUNAT = 421119
   OpeCGOpeProvPagoPagUNAT = 421120
   OpeCGOpeProvDevSUNAT = 421124 'PASIERS1242014
   
   OpeCGOpeProvRechazo = 421121
   OpeCGOpeProvEntrOP = 421122
   OpeCGOpeProvEntrCh = 421123
   
   OpeCGOpeProveedorME = 422100
   OpeCGOpeProvPagoME = 422110
   OpeCGOpeProvRechazoME = 422121
   OpeCGOpeProvEntrOPME = 422122
   OpeCGOpeProvEntrChME = 422123
   
    gOpeCGExtProveedores = 421130
    gOpeCGExtProvPago = 421140
    gOpeCGExtProvPagoEfectivo = 421141
    gOpeCGExtProvPagoTransfer = 421142
    gOpeCGExtProvPagoOPago = 421143
    gOpeCGExtProvPagoCheque = 421144
    gOpeCGExtProvPagoAbono = 421145
    gOpeCGExtProvPagoRechazo = 421151
    gOpeCGExtProvEntrOPago = 421152
    gOpeCGExtProvEntrCheques = 421153
    gOpeCGExtProvPagoSUNAT = 421154
    gOpeCGExtProvDevSUNAT = 421155 'PASIERS1242014

    gOpeCGExtProveedoresME = 421130
    gOpeCGExtProvPagoME = 421140
    gOpeCGExtProvPagoEfectivoME = 422141
    gOpeCGExtProvPagoTransferME = 422142
    gOpeCGExtProvPagoOPagoME = 422143
    gOpeCGExtProvPagoChequeME = 422144
    gOpeCGExtProvPagoAbonoME = 422145
    gOpeCGExtProvPagoRechazoME = 422151
    gOpeCGExtProvEntrOPagoME = 422152
    gOpeCGExtProvEntrChequesME = 422153
    
End Enum

'EJVG20121123 ***
Public Enum OpeCheque
    OpeChqRegistroTalonarioMN = 431191
    OpeChqRegistroTalonarioME = 432191
End Enum
'END EJVG *******

Public Enum OpeCGRepEncaje
   OpeCGRepEncajeEst = 461410       'ESTADISTICAS DE ENCAJE
   OpeCGRepEncajeConsolSdoEnc = 461411 'Consolidado de Saldos para Efecto de Encaje
   OpeCGRepEncajeAgencia = 461412   'Saldos para Efecto de Encaje por Agencia
   OpeCGRepEncajeConsolPosLiq = 461413 'Consolidado de Saldos para Posición de Liquidez
   OpeCGRepEncajeEstME = 462410       'ESTADISTICAS DE ENCAJE
   OpeCGRepEncajeConsolSdoEncME = 462411 'Consolidado de Saldos para Efecto de Encaje
   OpeCGRepEncajeAgenciaME = 462412   'Saldos para Efecto de Encaje por Agencia
   OpeCGRepEncajeConsolPosLiqME = 462413 'Consolidado de Saldos para Posición de Liquidez
   
   OpeEstEncajeSimulacionPlaEncajeMN = 461416 ' Simulacion de Planilla de Encaje MN
   OpeEstEncajeSimulacionPlaEncajeME = 462416 ' Simulacion de Planilla de Encaje ME
End Enum

Public Enum OpeCGRepEstadOperacion
   OpeCGRepEstadOpeUsuario = 461601
   OpeCGRepEstadOpeUsuarioME = 462601
End Enum

Public Enum OpeCGRepOtros
   OpeCGRepOtrBilletesFalsosMN = 461901
   OpeCGRepOtrBilletesFalsosME = 462901
End Enum

Public Enum OpeCGRemCheques
    gOpeCGRemChequesMN = 421210
    gOpeCGRemChequesME = 422210
    gOpeCGRemChequesMantMN = 421220
    gOpeCGRemChequesMantME = 422220
End Enum
    
'PROVISIONES
Public Enum ContProvisionProvee
    gContProvisionProveeMN = 701400
    gContProvOrdenCompraMN = 701401
    gContProvOrdenServicMN = 701402
    gContProvDirectaMN = 701403
    'Orden Compra
    gContProvDirectaOCMN = 701404
    'RRHH
    gContProvDirectaRRHHMN = 701405
    gContProvLogComprobanteMN = 701406 'EJVG20131113
    gContProvCarteraCredMN = 701410
    
    gContProvisionProveeME = 702400
    gContProvOrdenCompraME = 702401
    gContProvOrdenServicME = 702402
    gContProvDirectaME = 702403
    'Orden Compra
    gContProvDirectaOCME = 702404
    'RRHH
    gContProvDirectaRRHHME = 702405
    gContProvLogComprobanteME = 702406 'EJVG20131113
    gContProvCarteraCredME = 702410
End Enum

'OPERACIONES DE ANALISIS DE PENDIENTES
Public Enum ContAnalisPendientes
    gAnaALQUILERES = 741000
    gAnaRegGarantía = 741010
    gAnaDevGarantía = 741020
    gAnaDevGarantíaEfectivo = 741021
    gAnaDevGarantíaCheque = 741022
    gAnaDevGarantíaIngrVentanilla = 741026
    gAnaDevGarantíaGasto = 741028
    
    gAnaDEMANDAS = 741030
    gAnaDemandaIngr = 741031
    gAnaDemandaDescargo = 741040
    gAnaDemandaDescargoEfectivo = 741041
    gAnaDemandaDescargoCheque = 741042
    gAnaDemandaDescargoCargoCtaAho = 741043
    gAnaDemandaDescargoIngrVentanilla = 741046
    gAnaDemandaDescargoGasto = 741048
    
    gAnaINDEMNIZASINIESTROS = 741050
    gAnaIndemnizaSiniestroRegistro = 741051
    gAnaIndemnizaPago = 741060
    gAnaIndemnizaPagoEfectivo = 741061
    gAnaIndemnizaPagoCheque = 741062
    gAnaIndemnizaPagoIngrVentanilla = 741066
    gAnaIndemnizaPagoGasto = 741068
    
    gAnaOTRASCUENTASPORCOBRAR = 741080
    gAnaOtrasCtaCobrarRegistroAsiento = 741081
    gAnaOtrasCtaCobrarProvisPago = 741082
    gAnaOtrasCtaCobrarCancelaAsiento = 741089
    gAnaOtrasCtaCobrarCancelaDeuda = 741090
    gAnaOtrasCtaCobrarCancelaEfectivo = 741091
    gAnaOtrasCtaCobrarCancelaCheque = 741092
    gAnaOtrasCtaCobrarCancelaCargoCta = 741093
    gAnaOtrasCtaCobrarCancelaIngrVentanilla = 741096
    gAnaSubsidioPreyPostNatal = 741100
    gAnaSubsidioPrePostAbonoSeguro = 741101
    gAnaSubsidioPrePostAbonoCheque = 741102
    gAnaSubsidioPrePostAbonoEfectivo = 741103
    gAnaSubsidioEnfermedad = 741110
    gAnaSubsidioEnfermAbonoSeguro = 741111
    gAnaSubsidioEnfermAbonoSegCheque = 741112
    gAnaSubsidioEnfermAbonoSegEfectivo = 741113
    gAnaOTRASOperacCAJAGENERAL = 741120
    gAnaOtraOpeCGRegistroAsiento = 741121
    gAnaOtraOpeCGRegulaDocSustentatorio = 741122
    gAnaOtraOpeCGRegulaAsiento = 741123
    gAnaOtraOpeCGRegulaCajaGeneral = 741124
    gAnaOtraOpeCGRegulaCheque = 741125
    gAnaOtraOpeCGRegulaEfectivo = 741127
    gAnaOtraOpeCGRegulaIngrVentanilla = 741128
    gAnaOTRAPROVISION = 741300
    gAnaOtraProvisRegistProvision = 741301
    gAnaOtraProvisRegulaProvision = 741302
    gAnaOtraProvisRegulaProvAsiento = 741303
    gAnaCANJEORDENPAGOxCHEQUE = 741310
    gAnaCanjeOPRegulariza = 741311
    gAnaTRANSFERENCIAS = 741315
    gAnaTranferRegulariza = 741316
    gAnaOTRASOPERACLIQUID = 741330
    gAnaOtrasOpeLiqRegistAsiento = 741331
    gAnaOtrasOpeLiqRegulaAsiento = 741332
    gAnaOtrasOpeLiqRegulaCajaGen = 741333
    gAnaRECURSOSHUMANOS = 741340
    gAnaRecursosHumRegulaAsiento = 741342

    gAnaALQUILERESME = 742000
    gAnaRegGarantíaME = 742001
    gAnaDevGarantíaME = 742020
    gAnaDevGarantíaEfectivoME = 742021
    gAnaDevGarantíaChequeME = 742022
    gAnaDevGarantíaIngrVentanillaME = 742026
    gAnaDevGarantíaGastoME = 742028
    
    gAnaDEMANDASME = 742030
    gAnaDemandaIngrME = 742031
    gAnaDemandaDescargoME = 742040
    gAnaDemandaDescargoEfectivoME = 742041
    gAnaDemandaDescargoChequeME = 742042
    gAnaDemandaDescargoCargoCtaAhoME = 742043
    gAnaDemandaDescargoIngrVentanillaME = 742046
    gAnaDemandaDescargoGastoME = 742048
    
    gAnaINDEMNIZASINIESTROSME = 742050
    gAnaIndemnizaSiniestroRegistroME = 742051
    gAnaIndemnizaPagoME = 742060
    gAnaIndemnizaPagoEfectivoME = 742061
    gAnaIndemnizaPagoChequeME = 742062
    gAnaIndemnizaPagoIngrVentanillaME = 742066
    gAnaIndemnizaPagoGastoME = 742068
    
    gAnaOTRASCUENTASPORCOBRARME = 742080
    gAnaOtrasCtaCobrarRegistroAsientoME = 742081
    gAnaOtrasCtaCobrarProvisPagoME = 742082
    gAnaOtrasCtaCobrarCancelaAsientoME = 742089
    gAnaOtrasCtaCobrarCancelaDeudaME = 742090
    gAnaOtrasCtaCobrarCancelaEfectivoME = 742091
    gAnaOtrasCtaCobrarCancelaChequeME = 742092
    gAnaOtrasCtaCobrarCancelaCargoCtaME = 742093
    gAnaOtrasCtaCobrarCancelaIngrVentanillaME = 742096
    gAnaSubsidioPreyPostNatalME = 742100
    gAnaSubsidioPrePostAbonoSeguroME = 742101
    gAnaSubsidioPrePostAbonoChequeME = 742102
    gAnaSubsidioPrePostAbonoEfectivoME = 742103
    gAnaSubsidioEnfermedadME = 742110
    gAnaSubsidioEnfermAbonoSeguroME = 742111
    gAnaSubsidioEnfermAbonoSegChequeME = 742112
    gAnaSubsidioEnfermAbonoSegEfectivoME = 742113
    gAnaOTRASOperacCAJAGENERALME = 742120
    gAnaOtraOpeCGRegistroAsientoME = 742121
    gAnaOtraOpeCGRegulaDocSustentatorioME = 742122
    gAnaOtraOpeCGRegulaAsientoME = 742123
    gAnaOtraOpeCGRegulaCajaGeneralME = 742124
    gAnaOtraOpeCGRegulaChequeME = 742125
    gAnaOtraOpeCGRegulaEfectivoME = 742127
    gAnaOtraOpeCGRegulaIngrVentanillaME = 742128
    gAnaOTRAPROVISIONME = 742300
    gAnaOtraProvisRegistProvisionME = 742301
    gAnaOtraProvisRegulaProvisionME = 742302
    gAnaOtraProvisRegulaProvAsientoME = 741303
    gAnaCANJEORDENPAGOxCHEQUEME = 742310
    gAnaCanjeOPRegularizaME = 742311
    gAnaTRANSFERENCIASME = 742315
    gAnaTranferRegularizaME = 742316
    gAnaOTRASOPERACLIQUIDME = 742330
    gAnaOtrasOpeLiqRegistAsientoME = 742331
    gAnaOtrasOpeLiqRegulaAsientoME = 742332
    gAnaOtrasOpeLiqRegulaCajaGenME = 742333
    gAnaRECURSOSHUMANOSME = 742340
    gAnaRecursosHumRegulaAsientoME = 742342

End Enum
Public Enum RepCGEncajeBCR
   RepCGEncajeBCR = 761200   'INFORME DE ENCAJE AL BCR
   RepCGEncBCRObligacion = 761201  'Obligaciones Sujetas a Encaje
   RepCGEncBCRCredDeposi = 761202  'Créditos, Depósitos e Interbancarios
   RepCGEncBCRCredRecibi = 761203  'Créditos Recibidos no Sujetas a Encaje
   RepCGEncBCRObligaExon = 761204  'Obligaciones Exoneradas de Guardar Encaje
   RepCGEncBCRLinCredExt = 761205  'Líneas de Créditos y Créditos Provenientes del Exterior
   RepCGEncajeBCRME = 762200   'INFORME DE ENCAJE AL BCR
   RepCGEncBCRObligacionME = 762201  'Obligaciones Sujetas a Encaje
   RepCGEncBCRCredDeposiME = 762202  'Créditos, Depósitos e Interbancarios
   RepCGEncBCRCredRecibiME = 762203  'Créditos Recibidos no Sujetas a Encaje
   RepCGEncBCRObligaExonME = 762204  'Obligaciones Exoneradas de Guardar Encaje
   RepCGEncBCRLinCredExtME = 762205  'Líneas de Créditos y Créditos Provenientes del Exterior
End Enum

Public Enum RepContabilidad
    gContRegCompraGastos = 760200
    gContRegVentas = 760210
    gContRepBaseFormula = 760100
    gContRepBaseFormCarteraCont = 760104
    'EJVG20130502 ***
    gContRepBaseNotasEstadoSitFinan = 760105
    gContRepBaseNotasEstadoResultado = 760106
    gContRepBaseFormaB2 = 760107
    'END EJVG *******
    'ALPA 20130517******************
    gContRepEstadoSitFinanEEFF1 = 760108
    gContRepEstadoSitFinanEEFF2 = 760109
    gContRepEstadoSitFinanEEFF3 = 760110
    '*******************************
    gContRepEstadoFlujoEfectivo = 760116 ' FRHU 20140104 RQ13658
    gContRepHojaTrabajoFlujoEfectivo = 760118 ' FRHU 20140104 RQ13659
    gContRepBaseNotasComplementariaInfoAnual = 760117 'EJVG20140102
    gContRepPlanillaPagoProv = 760301
    gContRepControlGastoProv = 760302
    gContRepPagoProvDetalleDoc = 760303
    
    gContRepEstadIngGastos = 760251
    gContRepEstadProv = 760252
    
End Enum

'Captaciones antes estaban en frmrepahoconsol
Public Enum RepCaptaciones
    gContRepCaptacCVMonExtr = "763201"
    gContRepCaptacSituacCaptac = "763202"
    gContRepCaptacMovCV = "763203"
    gContRepCaptacIngPagos = "763204"
    gContRepCaptacCredDesem = "763205"
    
End Enum

Public Enum RepVarios
    gContRepInusuales = "763300"
    gContRepAsientoxCuenta = "764000" 'YIHU RS015-2015 20150406****************************
End Enum

'ALPA20140227****************************
Public Enum RepOperaciones
    gContRepGPSpot = "760611"
End Enum
'***************************************

Public Enum OtrosReportes
    gProvContxCtasCont = "763401"
    gProvContxOpe = "763402"
    gIntDevPFxFxAG = "763403"
    gIntDevCTSxFxAG = "763404"
    gPlazoFijoRango = "763405"
    gRepInstPubliResgos = "763406" 'Reporte de Instituciones Publicas
    gITFQuincena = "763407" 'Reporte Quincenal para pago de ITF GITU
    gPlazoFijoIntCash = "763408" 'By Capi 04022008
    gITFQuincenaTipoCambioDiario = "763409" 'ALPA 20081001
    gGenArchivoDaot = "763410" 'GITU 23-02-2009
    gGenPosCamDiaria = "763411" 'MIOL 20120703,SEGUN RQ12122
    'FRHU20140109 ***
    gPropuestaDistribuUtilidad = "763412"
    gPropuestaCapitaliUtilidad = "763413"
    gEstadoDeCambioPatrimonio = "763414"
    'END FRHU ******
End Enum

Public Enum ReportesBalanceContabilidad
    gPignoraticiosVencidos = "763501" 'Pignoraticios Vencidos
    'gCredRefinanciados = "763502" 'Creditos Refinanciados
    gCarteraCreditos = "763503" 'Cartera de Creditos
    gCarteraInteres = "763504" 'Cartera de Creditos
    gCreditosCastigados = "763505" 'Cartera de Creditos
    gInteresesDiferidos = "763506" 'Intereses Diferidos
    gPignoraticiosVigentes = "763507" 'Pignoraticios Vigentes
    gCreditosCondonados = "763508" 'Condonados
    gProvisionCreditos = "763509" 'Provision de Cartera
    gIntCreditosRefinanciados = "763510" 'Intereses Creditos Refinanciados
    gProvisionCartasFianzas = "763511" 'Provision de Cartas Fianzas
    gDetalleGarantias = "763512" 'Detalle de Garantias
    'By Capi 08022007
     gCarteraDetxIntDSP = "763513"
     gCarteraResxLineas = "763514"
    'ALPA 20090225*******************************************
     gInteresesCastigados = "763515"
    '********************************************************
    'ALPA 20101228*******************************************
     gVerificacionDevenSusp = "763520"
    '********************************************************
    'ALPA 20110727*******************************************
     gVerificacionDiferidos = "763521"
    '********************************************************
    gGastosCastigados = "763522" 'JUEZ 20130116
    gReporteComCartFi = "763523" 'ALPA 20130708
    gGenRepProyEjec = "763524" 'MIOL 20130725,SEGUN RQ13288
    gGenRepEEFFCont = "763526" 'ALPA 20150512
    gGenRepProvInd = "763527" 'PASI20160404
End Enum

Public Enum ReportesBalancePlaneamiento
    gCarteraCredPla = "763601" 'Cartera
    gCredDesembolsosPla = "763602" 'Desembolsos
    gCarteraVencidaPla = "763603" 'Cartera Vencida
    gCarteraRefinanciadaPla = "763604" 'Cartera Refinanciada
    gCarteraJudicialesPla = "763605" 'Cartera Refinanciada
    gCarteraRecupCapital = "763606" 'By Capi 18122007 para planeamiento Recuperacion de Capital
    gNumCliCtasEnAhorrosCreditos = "763607" '*** PEAC 20080915
    gOpeRptInfoEstadColocBCRP = "763608" 'EJVG20121112
    gCarteraDeAhorros = "763609" 'FRHU20140121 RQ13826
End Enum

Public Enum ReportesBalanceRiesgos
    gCredVigentesRiesgos = "763701" 'Creditos Vigentes
    gCredRefinanciadosRiesgos = "763702" 'Creditos Refinanciados
    gPlazoFijoRiesgos = "763703" 'Creditos Refinanciados
    gPPosCambRiesgos = "763704" 'MIOL 20130208 RFC038-2012
    gVarCambBackTestRiesgos = "763705" 'NAGL ERS002-2017 20170717
    gPosicionAfectaRiesgoCamb = "763706" 'NAGL ERS002-2017 20170717
    gRepSegTOSE = "763707" 'NAGL ERS002-2017 20170912
End Enum

'ANEXOS
Public Enum AnexosContabilidad
    gContAnx02CredTpoGarantia = 770020
    gContAnx03FujoCrediticio = 770030
    gContAnx09 = "770090"

    gContAnx07 = 770070
    gContAnx07N = 770040
    gContAnx10DepColocaPer = 770100
    
    gContAnx11MovDepsMonto = 770110
    gContAnx13DepsSEscMonto = 770130
    gContAnx13DepsSEscMonto_Nuevo = 770131
    
    gContAnx15A_Estad = 770152
    gContAnx15A_Efect = 770153
    gContAnx15A_Banco = 770154
    gContAnx15A_Repor = 770155
    gContAnx15B = 770156    'ALPA20130905
    gContAnx15C = 770157    'ALPA20130905
    gContAnex15BMens = 770158 'NAGL ERS079-2017 20171226
    gContAnx16LiqVenc = 770160
    gContAnx16A = 770161
    gContAnx16B = 770162
    gContAnx17A_FSD = 770171
    gContAnx17B_FSD = 770172
    
    gContAnx17ListadoGenCtas = 770174
    gContAnx17CtasFuncionarios = 770175
    gContAnx17ListadoFSD = 770176
    
    gContAnx18 = 770180
    gContAnx24 = 770240
End Enum

'Reportes SBS
Public Enum ReporteSBSContabilidad
    gContRep06Crediticio = 780060
    gContRepSpreadFinanciero = 780070 'ERS006-2017 JUCS
    gRiesgoSBSA02A = "780021"
    gRiesgoSBSA02B = "780022"
    gRiesgoSBSA050 = "780050"
    gRiesgoSBSA090 = "780090"
    gRiesgoSBSA190 = "780190"
    gRiesgoSBSA191 = "780191"
    gRiesgoSBSA200 = "780200"
    gRiesgoSBSA201 = "780201"
    gRiesgoSBSA210 = "780210"
    gPatrEfecAjxInfla = "780030"
    gRepCreditosIncumplidos = "780140"
    'ALPA 20090727********************************
    gRiesgoSBS2B1 = "780023"
    gRiesgoSBS2D = "780024"
    '*********************************************

    'MAVM 20090810 *********
    gFinancFinanciamientoRecibido = "780120"
    '***********************

End Enum
'FRHU20140109 ***
Public Enum ReportesActivosFijos
   gMovDepreAcumuladaIntangible = "763901"
   gMovActivoFijoIntangible = "763902"
   gCtoIntangibleOtrosAcivosAmo = "763903"
   gCtoInmuebleMaquinariaEquipo = "763904"
   gCtoDepreciacionDeActivoFijo = "763905"
End Enum
'END FRHU *******

'Balances
Public Enum BalancesCont
    gContBalanceFormaAB_1 = 760101
    gContBalanceFormaAB = 760102
    gContBalanceSectorial = 760103
End Enum

'Variables de operaciones para manejo de Pistas
'Add By GITU 21-01-2009
Public Enum VarOperacionesManejoPistas
    gOpcMenuAdmin = "790000"
    gMantTipoCambio = "790100"
    gOpcMenuConta = "790200"
    gMantParametrosViaticos = "790300"
    gMantDocumentos = "790500"
    gMantTasaImpuestos = "790700"
    gMantCalsifOperaciones = "790900"
    gMantOperaciones = "791200"
    gAnaliCtasDatosHistoMN = "791400"
    gAnaliCtasDatosHistoME = "791500"
    gMantCuentasContables = "791600"
    gMantCtasContablesBaseSBS = "791800"
    gMantObjetos = "791900"
    gMantAsientosContables = "792300"
    gExtAsientosCntConGenAsiento = "792400"
    gExtAsientosCntConElimAsiento = "792500"
    gMantCuentasContablesPend = "792700"
    gActuaReferenciasCtasPend = "792800"
    gProcesoValidaBalance = "793000"
    gGeneraBalanceHistorico = "793100"
    gGenraBalanceAjustado = "793200"
    gGeneraBalanceFormaAyB = "793300"
    gGeneraBalanceSectorial = "793400"
    gActualizaSaldos = "793500"
    gProcesaCierrContableAnual = "793400"
    gOpcMenuFinan = "795000"
    gConfirRecepcionCheques = "795100"
    gConsolEstadDiarias = "795200"
    gConsolBilletajeAgencias = "795300"
    gCierreDiarioContable = "795400"
    gControlPaqueteAdeudados = "795500"
End Enum
'End Gitu

Public Enum ExtornoCierre
    gContExtCierre = 701900
End Enum

Public Enum VarOtrosAjustes
    gReclasiGarantME = "702224"
End Enum
'EJVG20140721 ***
Public Enum TipoSistemaPensionProveeedor
    AFP = 1
    ONP = 2
End Enum
Public Enum TipoComisionAFPProveeedor
    Flujo = 1
    Saldo = 2
End Enum
Public Enum ConceptoRetencionSistemaPensionProveedor
    Aporte = 1
    SeguroAFP = 2
    ComsionAFP = 3
End Enum
'END EJVG ******

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

End Sub

