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
    
    gCGArendirCtaExtAtencMN = 401161
    gCGArendirCtaExtAtencME = 402161
    
    gCGArendirCtaExtRendMN = 401170
    gCGArendirCtaExtRendME = 402170
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
    
    gCGArendirViatExtAtencMN = 401261
    gCGArendirViatExtAtencME = 402261
    
    gCGArendirViatExtRendMN = 401270
    gCGArendirViatExtRendME = 402270
    
    gCGArendirViatAmpMN = 401280
    gCGArendirViatAmpME = 402280
    
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
    
    gOpeBoveAgeExtConfHabCGMN = 901007
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
    
    gOpeHabCajRegEfectMN = 901016
    gOpeHabCajRegEfectME = 902016
    
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
    gContBalanceFormaAB = 760102
    gContCierreAnual = 701850
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
    
    gOpeCGOpeBancosME = 401500
    gOpeCGOpeBancosDepEfecME = 401502
    gOpeCGOpeBancosRetEfecME = 401503
    gOpeCGOpeBancosConfRetEfecME = 401504
    gOpeCGOpeBancosRegChequesME = 401505
    gOpeCGOpeBancosDepChequesME = 401506
    gOpeCGOpeBancosDepDivBancosME = 401507
    gOpeCGOpeBancosRetDivBancosME = 401508
    
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
    OpeCGOpeCMACDepDivMN = 401602
    OpeCGOpeCMACRetDivMN = 401603
    OpeCGOpeCMACRegularizMN = 401604
    OpeCGOpeCMACAperAhorrosMN = 401612
    OpeCGOpeCMACAperPFMN = 401613
    OpeCGOpeCMACConfAperMN = 401614
    OpeCGOpeCMACIntDevPFMN = 401615
    OpeCGOpeCMACGastosComMN = 401616
    OpeCGOpeCMACCapIntDevPFMN = 401617
    OpeCGOpeCMACCancAhorrosMN = 401619
    OpeCGOpeCMACCancPFMN = 401620
    OpeCGOpeCMACMantCtasMN = 401621
    
    OpeCGOpeCMACDepDivME = 402602
    OpeCGOpeCMACRetDivME = 402603
    OpeCGOpeCMACRegularizME = 402604
    OpeCGOpeCMACAperAhorrosME = 402612
    OpeCGOpeCMACAperPFME = 402613
    OpeCGOpeCMACConfAperME = 402614
    OpeCGOpeCMACIntDevPFME = 402615
    OpeCGOpeCMACGastosComME = 402616
    OpeCGOpeCMACCapIntDevPFME = 402617
    OpeCGOpeCMACCancAhorrosME = 402619
    OpeCGOpeCMACCancPFME = 402620
    OpeCGOpeCMACMantCtasME = 402621
    
End Enum

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing


End Sub
