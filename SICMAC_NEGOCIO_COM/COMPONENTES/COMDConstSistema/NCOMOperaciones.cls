VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
End Enum

Public Enum OpeCGArendirCtaViativo
    gCGArendirViatMN = 401200
    gCGArendirViatME = 402200
    
    gCGArendirViatSolMN = 401210
    gCGArendirViatSolME = 402210
    
    gCGArendirViatRechMN = 401220
    gCGArendirViatRechME = 402220
    
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
    
    gCHExtDesembEfectivoMN = 401331
    gCHExtDesembEfectivoME = 402331
    
    gCHExtDesembOrdenPagoMN = 401332
    gCHExtDesembOrdenPagoME = 402332
    
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

End Enum

Public Enum OpeBovedaCajaGeneral
    gOpeBoveCG = 401400
    gOpeBoveCGHabAge = 402401
    gOpeBoveCGConfHabAgeBove = 401402
    gOpeBoveCGExtHabAge = 401404
    gOpeBoveCGExtConfHabAgeBov = 401405
End Enum

Public Enum OpeBovedaAgencia
    gOpeBoveAge = 901001
    gOpeBoveAgeConfHabCG = 901002
    gOpeBoveAgeHabAgeACG = 901003
    gOpeBoveAgeHabEntreAge = 901004
    gOpeBoveAgeHabCajero = 901005
    gOpeBoveAgeConfDevCaj = 901006
    gOpeBoveAgeRegEfect = 901007
    gOpeBoveAgeRegSobFalt = 901008
    gOpeBoveAgeRegSobrante = 901009
    gOpeBoveAgeRegFaltante = 901010

    gOpeBoveAgeExtConfHabCG = 901031
    gOpeBoveAgeExtHabAgeACG = 901032
    gOpeBoveAgeExtHabEntreAge = 901033
    gOpeBoveAgeExtHabCajero = 901034
    gOpeBoveAgeExtConfDevCaj = 901035
    gOpeBoveAgeExtRegEfect = 901036
    gOpeBoveAgeExtRegSobFalt = 901037
    gOpeBoveAgeExtRegSobrante = 901038
    gOpeBoveAgeExtRegFaltante = 901039
    'madm 20110201
    gOpeBoveAgeRegPreCuadre = 901040
    gOpeBoveAgeRegSobrantePreCuadre = 901041
    gOpeBoveAgeRegFaltantePreCuadre = 901042
    gOpeBoveAgeExtPreCuadre = 901029
    'madm 20110201
End Enum

Public Enum OpeCajeroMonedaExtranjera
    gOpeCajeroME = 900020
    gOpeCajeroMETipoCambio = 900021
    gOpeCajeroMECompra = 900022
    gOpeCajeroMEVenta = 900023
    'APRI RFC1705220001
    gOpeCajeroMECompraCheque = 900041
    gOpeCajeroMEVentaCheque = 900042
    'END APRI RFC1705220001
    gOpeCajeroMECompraCargoCta = 900027 'CTI4 ERS0112020
    gOpeCajeroMEVentaCargoCta = 900028 'CTI4 ERS0112020
    gOpeCajeroMEExtCompra = 909002
    gOpeCajeroMEExtVenta = 909003
    gOpeCajeroMETipoCambioEsp = 900024
    gOpeCajeroMECompraEsp = 900025
    gOpeCajeroMEVentaEsp = 900026
    gOpeCajeroMEExtCompraEsp = 909005
    gOpeCajeroMEExtVentaEsp = 909006
End Enum

Public Enum OpeHabCajero
    gOpeHabCajero = 901011
    gOpeHabCajDevABove = 901012
    gOpeHabCajTransfEfectCajeros = 901013
    gOpeHabCajRegSobFalt = 901014
    gOpeHabCajIngEfectRegulaFalt = 901015
    gOpeHabCajRegEfect = 901016
    gOpeHabCajConfHabBovAge = 901017
    gOpeHabCajDevBilletaje = 901018
    gOpeHabCajRegSobrante = 901019
    gOpeHabCajRegFaltante = 901020
    gOpeHabCajDevClienteRegulaSob = 901061
    gOpeHabCajIngRegulaSob = 901062
    
    gOpeHabCajExtDevABove = 901022
    gOpeHabCajExtTransfEfectCajeros = 901023
    gOpeHabCajExtRegSobFalt = 901024
    gOpeHabCajExtIngEfectRegulaFalt = 901025
    gOpeHabCajExtRegEfect = 901026
    gOpeHabCajExtConfHabBovAge = 901027
    gOpeHabCajExtDevBilletaje = 901028
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

Public Enum OpeCuadreCaja
    gOpeCajaCierreAgencia = 902001
    gOpeCajaExtCierreAgenica = 902031
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
    
    gOpeCGOpeBancosME = 401500
    gOpeCGOpeBancosDepEfecME = 402502
    gOpeCGOpeBancosRetEfecME = 402503
    gOpeCGOpeBancosConfRetEfecME = 402504
    gOpeCGOpeBancosRegChequesME = 402505
    gOpeCGOpeBancosDepChequesME = 402506
    gOpeCGOpeBancosDepDivBancosME = 402507
    gOpeCGOpeBancosRetDivBancosME = 402508
    gOpeCGOpeBancosRecibeChqAgMe = 402509

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
        
    gOpeCGOpeConfApertME = 402515
    gOpeCGOpeIntDevPFME = 402516
    gOpeCGOpeGastComBancosME = 402517
    gOpeCGOpeCapIntPFME = 402518
    gOpeCGOpeCancCtaCteME = 402520
    gOpeCGOpeCancCtaAhoME = 402521
    gOpeCGOpeCancCtaPFME = 402522
    gOpeCGOpeMantCtaBancosME = 402523
End Enum

Public Enum OpeCGOpeCMAC
    gOpeCGOpeCMACDepDivMN = 401602
    gOpeCGOpeCMACRetDivMN = 401603
    gOpeCGOpeCMACRegularizMN = 401604
    '***Modificado por ELRO el 20110923, según Acta 263-2011/TI-D
    gOpeCGOpeCMACDepEfeMN = 401605
    gOpeCGOpeCMACRetEfeMN = 401606
    gOpeCGOpeCMACConRetEfeMN = 401607
    '***Fin Modificado por ELRO**********************************
    '***Modificado por ELRO el 20110923, según Acta 269-2011/TI-D
    gOpeCGOpeCMACExtDepEfeMN = 401655
    gOpeCGOpeCMACExtRetEfeMN = 401656
    gOpeCGOpeCMACExtConRetEfeMN = 401657
    '***Fin Modificado por ELRO**********************************
    gOpeCGOpeCMACAperAhorrosMN = 401612
    gOpeCGOpeCMACAperPFMN = 401613
    gOpeCGOpeCMACConfAperMN = 401614
    gOpeCGOpeCMACIntDevPFMN = 401615
    gOpeCGOpeCMACGastosComMN = 401616
    gOpeCGOpeCMACCapIntDevPFMN = 401617
    gOpeCGOpeCMACCancAhorrosMN = 401619
    gOpeCGOpeCMACCancPFMN = 401620
    gOpeCGOpeCMACMantCtasMN = 401621
    
    gOpeCGOpeCMACDepDivME = 402602
    gOpeCGOpeCMACRetDivME = 402603
    gOpeCGOpeCMACRegularizME = 402604
    '***Modificado por ELRO el 20110923, según Acta 263-2011/TI-D
    gOpeCGOpeCMACDepEfeME = 402605
    gOpeCGOpeCMACRetEfeME = 402606
    gOpeCGOpeCMACConRetEfeME = 402607
    '***Fin Modificado por ELRO**********************************
    '***Modificado por ELRO el 20110923, según Acta 269-2011/TI-D
    gOpeCGOpeCMACExtDepEfeME = 402655
    gOpeCGOpeCMACExtRetEfeME = 402656
    gOpeCGOpeCMACExtConRetEfeME = 402657
    '***Fin Modificado por ELRO**********************************
    gOpeCGOpeCMACAperAhorrosME = 402612
    gOpeCGOpeCMACAperPFME = 402613
    gOpeCGOpeCMACConfAperME = 402614
    gOpeCGOpeCMACIntDevPFME = 402615
    gOpeCGOpeCMACGastosComME = 402616
    gOpeCGOpeCMACCapIntDevPFME = 402617
    gOpeCGOpeCMACCancAhorrosME = 402619
    gOpeCGOpeCMACCancPFME = 402620
    gOpeCGOpeCMACMantCtasME = 402621
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
    gOpeCGAdeudaExtProvisiónMN = 401810    'Extornos Provisión Adeudados Bancos MN
    gOpeCGAdeudaExtPagoCuotaMN = 401811    'Extornos Pago Adeudados Bancos MN
    gOpeCGAdeudaExtReprogramaMN = 401812   'Extornos Reprogramación Adeudados Bancos MN
    gOpeCGAdeudaMntPagaresMN = 401832      'Mantenimiento de Pagarés MN
    gOpeCGAdeudaProyeccionMN = 401833      'Proyecciones de Adeudados MN
    gOpeCGAdeudaReportesMN = 401834        'Reporte General de Adeudados

    gOpeCGAdeudadosME = 402800             'ADEDUADOS ME
    gOpeCGAdeudaCalendarioME = 402801      'Simulacion de Cronograma de Pagos
    gOpeCGAdeudaRegPagareME = 402803       'Registro Pagarés ME
    gOpeCGAdeudaRegPagareConfMe = 402804   'Confirmacion de Pagaré
    gOpeCGAdeudaProvisionME = 402805         'Provisión Adeudados ME
    gOpeCGAdeudaPagoCuotaME = 402806       'Pago Adeudados ME
    gOpeCGAdeudaReprogramaME = 402807      'Reprogramación Adeudados ME
    gOpeCGAdeudaExtornosME = 402808        'Extornos Adeudados ME
    gOpeCGAdeudaExtRegistroME = 402809     'Extornos Registro Pagarés Bancos ME
    gOpeCGAdeudaExtProvisiónME = 402810    'Extornos Provisión Adeudados Bancos ME
    gOpeCGAdeudaExtPagoCuotaME = 402811    'Extornos Pago Adeudados Bancos ME
    gOpeCGAdeudaExtReprogramaME = 402812   'Extornos Reprogramación Adeudados Bancos ME
    gOpeCGAdeudaMntPagaresME = 402832      'Mantenimiento de Pagarés ME
    gOpeCGAdeudaProyeccionME = 402833      'Proyecciones de Adeudados ME
    gOpeCGAdeudaReportesME = 402834        'Reporte General de AdeudadoE
End Enum

Public Enum OpeCGReportes
    OpeCGRepFlujoDiarioResMN = 461001
    OpeCGRepFlujoDiarioDetMN = 461002
    OpeCGRepRepBancosMN = 461010
    OpeCGRepRepBancosFlujoMensMN = 461011
    OpeCGRepRepBancosFlujoPFMN = 461012
    OpeCGRepRepBancosSaldosCtasMN = 461013
    OpeCGRepRepCMACSMN = 461020
    OpeCGRepRepCMACSFlujoMensMN = 461021
    OpeCGRepRepCMACSFlujoPFMN = 461022
    OpeCGRepRepCMACSSaldosCtasMN = 461023
    OpeCGRepRepOPMN = 461030
    OpeCGRepRepOPGirMN = 461031
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

'MONEDA EXTRANJERA
    OpeCGRepFlujoDiarioResME = 462001
    OpeCGRepFlujoDiarioDetME = 462002
    OpeCGRepRepBancosME = 462010
    OpeCGRepRepBancosFlujoMensME = 462011
    OpeCGRepRepBancosFlujoPFME = 462012
    OpeCGRepRepBancosSaldosCtasME = 462013
    OpeCGRepRepCMACSME = 462020
    OpeCGRepRepCMACSFlujoMensME = 462021
    OpeCGRepRepCMACSFlujoPFME = 462022
    OpeCGRepRepCMACSSaldosCtasME = 462023
    OpeCGRepRepOPME = 462030
    OpeCGRepRepOPGirME = 462031
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
End Enum


Public Enum OpeCGPagoSunat
   OpeCGPagoSunat = 411100
   OpeCGPagoSunatSegSocial = 411101
   OpeCGPagoSunatBoletaPag = 411102
   OpeCGPagoSunatIGVRenta = 411103
End Enum

Public Enum OpeCGCartaFianza
   OpeCGCartaFianza = 411110
   OpeCGCartaFianzaIng = 411111
   OpeCGCartaFianzaSal = 411112
   OpeCGCartaFianzaRepIngreso = 461401
   OpeCGCartaFianzaRepSalida = 461402
   OpeCGCartaFianzaME = 412110
   OpeCGCartaFianzaIngME = 412111
   OpeCGCartaFianzaSalME = 412112
   OpeCGCartaFianzaRepIngresoME = 462401
   OpeCGCartaFianzaRepSalidaME = 462402
End Enum

Public Enum OpeCGOtrosOpeEfectivo
   OpeCGOtrosOpeEfectivo = 421000
   OpeCGOtrosOpeEfecIngr = 421001   'Ingresos de Efectivo
   OpeCGOtrosOpeEfecEgre = 421002   'Egresos de Efectivo
   OpeCGOtrosOpeEfecCamb = 421003   'Cambio de Billetaje
   OpeCGOtrosOpeEfecOtro = 421009   'Otrs Operaciones
   
   OpeCGOtrosOpeEfectivoME = 422000
   OpeCGOtrosOpeEfecIngrME = 422001   'Ingresos de Efectivo
   OpeCGOtrosOpeEfecEgreME = 422002   'Egresos de Efectivo
   OpeCGOtrosOpeEfecCambme = 422003   'Cambio de Billetaje
   OpeCGOtrosOpeEfecOtroME = 422009   'Otrs Operaciones
End Enum

Public Enum OpeCGOpeProveedores
   OpeCGOpeProveedor = 421100
   OpeCGOpeProvPago = 421110
   OpeCGOpeProvRechazo = 421121
   OpeCGOpeProvEntrOP = 421122
   OpeCGOpeProvEntrCh = 421123
   OpeCGOpeProveedorME = 422100
   OpeCGOpeProvPagoME = 422110
   OpeCGOpeProvRechazoME = 422121
   OpeCGOpeProvEntrOPME = 422122
   OpeCGOpeProvEntrChME = 422123
End Enum

Public Enum OpeCGRepEncaje
   OpeCGRepEncajeEst = 461410       'ESTADISTICAS DE ENCAJE
   OpeCGRepEncajeConsolSdoEnc = 461411 'Consolidado de Saldos para Efecto de Encaje
   OpeCGRepEncajeAgencia = 461412   'Saldos para Efecto de Encaje por Agencia
   OpeCGRepEncajeConsolPosLiq = 461413 'Consolidado de Saldos para Posición de Liquidez
   OpeCGRepEncajeEstME = 462410       'ESTADISTICAS DE ENCAJE
   OpeCGRepEncajeConsolSdoEncME = 462411 'Consolidado de Saldos para Efecto de Encaje
   OpeCGRepEncajeAgenciaME = 462412   'Saldos para Efecto de Encaje por Agencia
   OpeCGRepEncajeConsolPosLiqME = 462413 'Consolidado de Saldos para Posición de Liquidez
End Enum

'PROVISIONES
Public Enum ContProvisionProvee
    gContProvisionProveeMN = 701400
    gContProvOrdenCompraMN = 701401
    gContProvOrdenServicMN = 701402
    gContProvDirectaMN = 701403
    gContProvCarteraCredMN = 701410
    
    gContProvisionProveeME = 702400
    gContProvOrdenCompraME = 702401
    gContProvOrdenServicME = 702402
    gContProvDirectaME = 702403
    gContProvCarteraCredME = 702410
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
    gContRepPlanillaPagoProv = 760301
    gContRepControlGastoProv = 760302
    
    gContRepEstadIngGastos = 760251
End Enum

'ANEXOS
Public Enum AnexosContabilidad
    gContAnx02CredTpoGarantia = 770020
    gContAnx03FujoCrediticio = 770030
    gContAnx07 = 770070
    gContAnx15A_Estad = 770152
    gContAnx15A_Efect = 770153
    gContAnx15A_Banco = 770154
    gContAnx15A_Repor = 770155
    gContAnx15B = 770156
    gContAnx16LiqVenc = 770160
    gContAnx16A = 770161
    gContAnx16B = 770162
    gContAnx17A_FSD = 770171
    gContAnx17B_FSD = 770172
End Enum

'Balances
Public Enum BalancesCont
    gContBalanceFormaAB_1 = 760101
    gContBalanceFormaAB = 760102
    gContBalanceSectorial = 760103
End Enum

'ANDE 20170512 DINERO ELECTRÓNICO
Public Enum OpeDineroElectronico
    gOpeDEConversion = 980001
    gOpeDEReConversion = 980002
End Enum

Public Enum ConceptoDineroElectronico
    gConceptoConversion = 301
    gConceptoReConversion = 302
End Enum
'FIN ANDE

