Attribute VB_Name = "Constantes"
Option Explicit

Global Const gsPersPersoneria = "gsPersPersoneria"
Global Const gsPersPersoneriaNat = "0"
Global Const gsPersPersoneriaJurSFL = "1"
Global Const gsPersPersoneriaJurCFL = "2"
Global Const gsPersPersoneriaJurCFLCMAC = "3"

Global Const gsPersIdTpo = "gsPersIdTpo"
Global Const gsPersIdDNI = "DNI"
Global Const gsPersIdRUC = "RUC"
Global Const gsPersIdFFPPAA = "FPA"
Global Const gsPersIdExtranjeria = "EXT"
Global Const gsPersIdSBS = "SBS"
Global Const gsPersIdIPSS = "IPS"
Global Const gsPersIdBrevete = "BRE"
Global Const gsPersIdRegPub = "RPU"
Global Const gsPersIdAFP = "AFP"
Global Const gsPersIdPartNaC = "PAR"
Global Const gsPersIdPasaporte = "PAS"
Global Const gsPersIdBolMil = "BOL"
Global Const gsPersIdLibMil = "LIB"
Global Const gsPersIdLibTribut = "LTB"
Global Const gsPersIdRUS = "RUS"

Global Const gsPersJurMagnitud = "gsPersJurMagnitud"
Global Const gsPersJurMagnitudGrande = "0"
Global Const gsPersJurMagnitudMediana = "1"
Global Const gsPersJurMagnitudPeque�a = "2"
Global Const gsPersJurMagnitudMiCro = "3"

Global Const gsAdeudTpoConC = "gsAdeudTpoConC"
Global Const gsAdeudTpoConCgsAp = "01"
Global Const gsAdeudTpoConCInt = "02"
Global Const gsAdeudTpoConCGastos = "03"

Global Const gsPersEstado = "gsPersEstado"
Global Const gsPersEstadoPNACtivo = "00"
Global Const gsPersEstadoPNFalleCido = "01"
Global Const gsPersEstadoPJACtivo = "10"
Global Const gsPersEstadoPJLiquidaCion = "11"
Global Const gsPersEstadoPJDisuelta = "12"
Global Const gsPersEstadoPJFusionada = "13"

Global Const gsPersRelaC = "gsPersRelaC"
Global Const gsPersRelaCConyugue = "00"
Global Const gsPersRelaCHijos = "01"
Global Const gsPersRelaCPadres = "02"
Global Const gsPersRelaCHermanos = "03"
Global Const gsPersRelaCOtros = "09"
Global Const gsPersRelaCRepresentante = "10"

Global Const gbPersRelaBenef = "gbPersRelaBenef"
Global Const gbPersRelaBenefDesaCtivado = "0"
Global Const gbPersRelaBenefACtivo = "1"

Global Const gbPersRelaAMP = "gbPersRelaAMP"
Global Const gbPersRelaAMPDesactivado = "0"
Global Const gbPersRelaAMPActivo = "1"

Global Const gbPersRelaEstado = "gbPersRelaEstado"
Global Const gbPersRelaEstadoDesativado = "0"
Global Const gbPersRelaEstadoACtivo = "1"
Global Const gsOPEstado = "gsOPEstado"
Global Const gsOPEstadoEmitido = "0"
Global Const gsOPEstadoCertifiCada = "1"
Global Const gsOPEstadoCobrado = "2"
Global Const gsOPEstadoReChazado = "3"
Global Const gsOPEstadoPerdido = "7"
Global Const gsOPEstadoExtornado = "8"
Global Const gsOPEstadoAnulado = "9"

Global Const gsChqEstado = "gsChqEstado"
Global Const gsChqEstEnValorizacion = "E"
Global Const gsChqEstValorizado = "V"
Global Const gsChqEstAnulado = "A"
Global Const gsChqEstRechazado = "R"
Global Const gsChqEstExtornado = "X"

Global Const gbColocPenalidad = "gbColocPenalidad"
Global Const gbColocPenalidadExonera = "1"
Global Const gbColocPenalidadNOExonera = "0"

Global Const gsColocRefinanciado = "gsColocRefinanciado"
Global Const gsColocRefCapInt = "1"
Global Const gsColocRefNoCapInt = "0"

Global Const gbColocProtesto = "gsColocProtesto"
Global Const gbColocNoProtesto = "0"

Global Const gsColocTipoGracia = "gsColocTipoGracia"
Global Const gsColocTipoGraCiaIniCio = "0"
Global Const gsColocTipoGraCiaUlt = "1"
Global Const gsColocTipoGraCiaProrat = "2"
Global Const gsColocTipoGraCiaConfig = "3"
Global Const gsColocTipoGraCiaExo = "4"

Global Const gsColocDestino = "gsColocDestino"
Global Const gsColocDestinoNoIndiCado = "0"
Global Const gsColocDestinoCapTrab = "1"
Global Const gsColocDestinoACtFijo = "2"
Global Const gsColocDestinoMixto = "3"
Global Const gsColocDestinoConsumo = "4"
Global Const gsColocDestinoOtros = "9"

Global Const gsColocCond = "gsColocCond"
Global Const gsColocCondNormal = "1"
Global Const gsColocCondReCurrente = "2"
Global Const gsColocCondParalelo = "3"

Global Const gsColocCalendCod = "gsColocCalendCod"
Global Const gsColocCalendCodPFCF = "000"
Global Const gsColocCalendCodPFCFPG = "001"
Global Const gsColocCalendCodPFCC = "010"
Global Const gsColocCalendCodPFCCPG = "011"
Global Const gsColocCalendCodPFCD = "020"
Global Const gsColocCalendCodPFCDPG = "021"
Global Const gsColocCalendCodFFCF = "100"
Global Const gsColocCalendCodFFCFPG = "101"
Global Const gsColocCalendCodFFCC = "110"
Global Const gsColocCalendCodFFCCPG = "111"
Global Const gsColocCalendCodFFCD = "120"
Global Const gsColocCalendCodFFCDPG = "121"
Global Const gsColocCalendCodCL = "290"
Global Const gsColocCalendCodCLPG = "291"

Global Const gsColocCalendApl = "gsColocCalendApl"
Global Const gsColocCalendAplDesembolso = "0"
Global Const gsColocCalendAplCuota = "1"
Global Const gsColocCalendAplJudiCial = "2"
Global Const gsColocCalendAplOtros = "9"

Global Const gsColocCalendFlag = "gsColocCalendFlag"
Global Const gsColocCalendFlagPendiente = "0"
Global Const gsColocCalendFlagPagado = "1"

Global Const gsColocConCeptoCod = "gsColocConCeptoCod"
Global Const gsColocConCeptoCodCapital = "1000"
Global Const gsColocConCeptoCodInteresCompensatorio = "1100"
Global Const gsColocConCeptoCodInteresMoratorio = "1101"
Global Const gsColocConCeptoCodInteresGraCia = "1102"
Global Const gsColocConCeptoCodInteresReprogramado = "1103"
Global Const gsColocConCeptoCodGasto01 = "1201"
Global Const gsColocConCeptoCodGastoVarios = "1299"
Global Const gsColocConCeptoCodPiCapital = "2000"
Global Const gsColocConCeptoCodPigInteresCompensatorio = "2100"
Global Const gsColocConCeptoCodPigInteresMoratorio = "2101"
Global Const gsColocConCeptoCodPigTasaCion = "2200"
Global Const gsColocConCeptoCodPiCustodia = "2201"
Global Const gsColocConCeptoCodPiCustodiaDiferida = "2202"
Global Const gsColocConCeptoCodPigImpuesto = "2203"
Global Const gsColocConCeptoCodJudGasto01 = "3001"
Global Const gsColocConCeptoCodJudGastoVarios = "3099"
Global Const gsColocConCeptoCodJudComision01 = "3100"
Global Const gsColocConCeptoCodJudComision99 = "3199"

Global Const gsColocConConcApl = "gsColocConConcApl"
Global Const gsColocConCeptoAplTodosDC = "00"
Global Const gsColocConCeptoAplTodosD = "10"
Global Const gsColocConCeptoAplDesembolso = "11"
Global Const gsColocConCeptoAplTodosC = "20"
Global Const gsColocConCeptoAplCuota = "21"

Global Const gsPrdEstado = "gsPrdEstado"
Global Const gsPrdEstCapVig = "1000"
Global Const gsPrdEstCapVigInac = "1001"
Global Const gsPrdEstCapBloqTot = "1100"
Global Const gsPrdEstCapBloqTotInac = "1101"
Global Const gsPrdEstCapBloqRet = "1200"
Global Const gsPrdEstCapBloqRetInac = "1201"
Global Const gsPrdEstCapAnu = "1300"
Global Const gsPrdEstCapAnuInac = "1301"
Global Const gsPrdEstCapCanc = "1400"
Global Const gsPrdEstCapCancInac = "1401"
Global Const gsPrdEstColocSolic = "2000"
Global Const gsPrdEstColocSug = "2001"
Global Const gsPrdEstColocAprob = "2002"
Global Const gsPrdEstColocRech = "2003"
Global Const gsPrdEstColocDesemb = "2004"
Global Const gsPrdEstColocVigNorm = "2020"
Global Const gsPrdEstColocVigVenc = "2021"
Global Const gsPrdEstColocVigMor = "2022"
Global Const gsPrdEstColocRefNorm = "2030"
Global Const gsPrdEstColocRefVenc = "2031"
Global Const gsPrdEstColocRefMor = "2032"
Global Const gsPrdEstColocEmbarg = "2040"
Global Const gsPrdEstPigReg = "2100"
Global Const gsPrdEstPigDesemb = "2101"
Global Const gsPrdEstPigCanc = "2102"
Global Const gsPrdEstPigEntJoya = "2103"
Global Const gsPrdEstPigVenc = "2104"
Global Const gsPrdEstPigRenov = "2105"
Global Const gsPrdEstPigRemate = "2106"
Global Const gsPrdEstPigAdjud = "2107"
Global Const gsPrdEstPigSubast = "2108"
Global Const gsPrdEstPigBarras = "2109"
Global Const gsPrdEstPigAnulNoDesemb = "2199"
Global Const gsPrdEstJudTras = "2200"
Global Const gsPrdEstJudNeg = "2201"
Global Const gsPrdEstJudProcIni = "2202"
Global Const gsPrdEstJudProcFin = "2203"
Global Const gsPrdEstJudCastig = "2900"

Global Const gsPrdPersRelac = "gsPrdPersRelac"
Global Const gsPrdPersRelacCaptaCTitular = "10"
Global Const gsPrdPersRelacCaptaCApoderado = "11"
Global Const gsPrdPersRelacCaptaCRepresentante = "12"
Global Const gsPrdPersRelacColoCTitular = "20"
Global Const gsPrdPersRelacColoCConyugue = "21"
Global Const gsPrdPersRelacColoCCodeudor = "22"
Global Const gsPrdPersRelacColoCRepresentante = "23"
Global Const gsPrdPersRelacColoCRepresCodeudor = "24"
Global Const gsPrdPersRelacColoCAnalista = "28"
Global Const gsPrdPersRelacColoCApoderado = "29"
Global Const gsPrdPersRelacJudEstudioJuridiCo = "30"
Global Const gsPrdPersRelacJudJuez = "31"
Global Const gsPrdPersRelacJudSeCretario = "32"
Global Const gsPrdPersRelacJudAbogadoReponsable = "33"
Global Const gsPrdPersRelacJudAbogadoContrario = "34"

Global Const gsPrdTasaInt = "gsPrdTasaInt"
Global Const gsPrdTasaIntCaptacNormal = "100"
Global Const gsPrdTasaIntCaptacEspeCial = "101"
Global Const gsPrdTasaIntColocCompNormal = "200"
Global Const gsPrdTasaIntColocCompGraCia = "201"
Global Const gsPrdTasaIntColocMoraNormal = "210"
Global Const gsPrdTasaIntColocMoraGraCia = "211"

Global Const gsPrdPFFormaRetiro = "gsPrdPFFormaRetiro"
Global Const gsPrdPFFormaRetiroAlFinal = "1"
Global Const gsPrdPFFormaRetiroMensual = "2"
Global Const gsPrdPFFormaRetiroLibre = "3"

Global Const gsColocCalifica = "gsColocCalifica"
Global Const gsColocCalificaNormal = "0"
Global Const gsColocCalificaPotenCial = "1"
Global Const gsColocCalificaDefiCiente = "2"
Global Const gsColocCalificaDudoso = "3"
Global Const gsColocCalificaPerdida = "4"

Global Const gsColocacNota = "gsColocacNota"
Global Const gsColocacNotaNormal = "0"
Global Const gsPrdCtaTpo = "gsPrdCtaTpo"
Global Const gsPrdCtaTpoIndividual = "I"
Global Const gsPrdCtaTpoManComunadaY = "Y"
Global Const gsPrdCtaTpoManComunadaO = "O"

Global Const gsMovEst = "gsMovEst"
Global Const gsMovEstContabMovContable = "10"
Global Const gsMovEstContabPendiente = "11"
Global Const gsMovEstContabRechazado = "12"
Global Const gsMovEstLogInicio = "20"
Global Const gsMovEstLogTramite = "21"
Global Const gsMovEstLogParaAtencion = "22"
Global Const gsMovEstLogAtencion = "23"
Global Const gsMovEstLogRechazado = "24"
Global Const gsMovEstLogAnulado = "25"

Global Const gsMovFlag = "gsMovFlag"
Global Const gsMovFlagVigente = ""
Global Const gsMovFlagEliminado = "X"
Global Const gsMovFlagExtornado = "E"
Global Const gsMovFlagDeExtorno = "N"
Global Const gsMovFlagDeLeido = "L"

Global Const gsMovParalelo = "CMovParalelo"
Global Const gsMovParaleloTransferido = "T"

Global Const gsRHProcSelTpo = "gsRHProcSeleccTpo"
Global Const gsRHProcSelTpoInterno = "0"
Global Const gsRHProcSelTpoExterno = "1"
Global Const gsRHProcSelTpoParaEmp = "2"

Global Const gsRHProcSelEst = "gsRHProcSelEst"
Global Const gsRHProcSelEstIniCiado = "0"
Global Const gsRHProcSelEstFinalizado = "1"

Global Const gsRHProcSelRes = "gsRHProcSelRes"
Global Const gsRHProcSelResIngreso = "0"
Global Const gsRHProcSelResSuplente = "1"
Global Const gsRHProcSelResNoIngreso = "2"

Global Const gsRHEst = "gsRHEst"
Global Const gsRHEstInactivo = "0"
Global Const gsRHEstActivo = "1"
Global Const gsRHEstVacacFisicas = "2"
Global Const gsRHEstVacacSuspend = "3"
Global Const gsRHEstLicSGoceHaber = "4"
Global Const gsRHEstLicCGoceHaber = "5"
Global Const gsRHEstSubs = "6"
Global Const gsRHEstSuspend = "7"
Global Const gsRHEstDespedido = "8"
Global Const gsRHEstRetirado = "9"

Global Const gsLogObtTpo = "gsLogObtTpo"
Global Const gsLogObtTpoNormal = "1"
Global Const gsLogObtTpoExtemporaneo = "2"

Global Const gsDocChqPlaza = "gsDocChqPlaza"
Global Const gsDocChqPlazaLoCal = "0"
Global Const gsDocChqPlazaRemota = "1"

Global Const gsTarjEst = "gsTarjEstado"
Global Const gsTarjEstActiva = "A"
Global Const gsTarjEstBloqueada = "B"
Global Const gsTarjEstCancelada = "C"
Global Const gsTarjEstVencida = "V"
Global Const gsTarjEstCambioClave = "N"
Global Const gsTarjEstNuevaClave = "K"

Global Const gsRHCondEmp = "gsRHCondEmp"
Global Const gsRHCondEmpContratado = "0"
Global Const gsRHCondEmpEstable = "1"

Global Const gsRHCateg = "gsRHCateg"
Global Const gsRHCategFunCionario = "A"
Global Const gsRHCategEmpleado = "B"

Global Const gsRHCargoEst = "gsRHCargoEst"
Global Const gsRHCargoEstACtivo = "A"
Global Const gsRHCargoEstEliminado = "E"

Global Const gsRHConc = "gsRHConc"
Global Const gsRHConcIng = "100"
Global Const gsRHConcIngBoncarFam = "101"
Global Const gsRHConcIngBonConsol = "102"
Global Const gsRHConcIngBonVacacional = "103"
Global Const gsRHConcIngBonCargo = "104"
Global Const gsRHConcIngBonResponsabilidad = "105"
Global Const gsRHConcIngCTSBonVac = "106"
Global Const gsRHConcIngCTSBonGratif = "107"
Global Const gsRHConcIngCTSTercio = "108"
Global Const gsRHConcIngAFP1023 = "109"
Global Const gsRHConcIngAFP0300 = "110"
Global Const gsRHConcIngIngDiaSub = "111"
Global Const gsRHConcIngNetoPagar = "112"
Global Const gsRHConcIngONP = "113"
Global Const gsRHConcIngPrueba1 = "114"
Global Const gsRHConcIngReintegro = "115"
Global Const gsRHConcIngRemNoAfeCta = "116"
Global Const gsRHConcIngRemVacacional = "117"
Global Const gsRHConcIngRemMensual = "118"
Global Const gsRHConcIngSueBasico = "119"
Global Const gsRHConcIngTerBonCargo = "120"
Global Const gsRHConcIngTerBonCarFam = "121"
Global Const gsRHConcIngTerBonConsol = "122"
Global Const gsRHConcIngTerCio = "123"
Global Const gsRHConcIngTerAFP1023 = "124"
Global Const gsRHConcIngTerAFP300 = "125"
Global Const gsRHConcIngTerSueBasico = "126"
Global Const gsRHConcIngTerBonResp = "127"
Global Const gsRHConcIngTotBonProduct = "128"
Global Const gsRHConcIngTotCTS = "129"
Global Const gsRHConcIngTotIngresos = "130"
Global Const gsRHConcIngTotalUtilidades = "131"
Global Const gsRHConcIngTotalGratifiCaCiones = "132"
Global Const gsRHConcIngBonProduct = "133"
Global Const gsRHConcIngIngresosLocacion = "134"
Global Const gsRHConcIngAguinaldoNavide�o = "135"
Global Const gsRHConcIngAguinaldoJuguetes = "136"
Global Const gsRHConcIngVacSueldoBasico = "137"
Global Const gsRHConcIngVacBonCargo = "138"
Global Const gsRHConcIngVacAFP300 = "139"
Global Const gsRHConcIngVacBonResponsabilidad = "140"
Global Const gsRHConcIngVacBonFam = "141"
Global Const gsRHConcIngVacAFP1023 = "142"
Global Const gsRHConcIngVacCTSTotRem = "143"
Global Const gsRHConcIngVacBonConsol = "144"
Global Const gsRHConcIngTerBonProduCtividad = "145"
Global Const gsRHConcIngVacSueBasico = "146"

Global Const gsRHConcDcto = "gsRHConcDcto"
Global Const gsRHConcDcto5taCat = "201"
Global Const gsRHConcDctoAFPComVar = "202"
Global Const gsRHConcDctoAFPCuotaFija = "203"
Global Const gsRHConcDctoAFPPrima = "204"
Global Const gsRHConcDctoAMP = "205"
Global Const gsRHConcDctoBcoLima = "206"
Global Const gsRHConcDctoEsSaludVida = "207"
Global Const gsRHConcDctoJudFijo = "208"
Global Const gsRHConcDctoJudPorc = "209"
Global Const gsRHConcDctoHipotec = "210"
Global Const gsRHConcDctoTardanzas = "211"
Global Const gsRHConcDctoONP1300 = "212"
Global Const gsRHConcDctoPrueba = "213"
Global Const gsRHConcDctoPrestAdmin = "214"
Global Const gsRHConcDctoTotalDctos = "215"

Global Const gsRHConcAport = "gsRHConcAport"
Global Const gsRHConcAportExtSolid = "301"
Global Const gsRHConcAportSegSoc = "302"
Global Const gsRHConcAportTotalAport = "303"
Global Const gsRHConcVarUsu = "gsRHConcVarUsu"
Global Const gsRHConcVarUsuCTSMesTrab = "401"
Global Const gsRHConcVarUsuGratMesTrab = "402"
Global Const gsRHConcVarUsuSubNumDias = "403"
Global Const gsRHConcVarUsuPorcMesTrab = "404"
Global Const gsRHConcVarGlob = "gsRHConcVarGlob"
Global Const gsRHConcVarGlobTotDiasLab = "501"
Global Const gsRHConcVarGlobPorcPagUTIL = "502"
Global Const gsRHConcVarGlobSumTotIng = "503"
Global Const gsRHConcVarGlobValUtil = "504"
Global Const gsRHConcFuncConst = "gsRHConcFuncConst"
Global Const gsRHConcFuncConstAFPValComVar = "601"
Global Const gsRHConcFuncConstAFPValorPrima = "602"
Global Const gsRHConcFuncConstDctoJudiFijo = "603"
Global Const gsRHConcFuncConstDctoJudiPorc = "604"
Global Const gsRHConcFuncConst5taCat = "605"
Global Const gsRHConcFuncConstAFP1023 = "606"
Global Const gsRHConcFuncConstAFP0300 = "607"
Global Const gsRHConcFuncConstCTSMesTrab = "608"
Global Const gsRHConcFuncConstMinSal = "609"
Global Const gsRHConcFuncConstMinTard = "610"
Global Const gsRHConcFuncConstDiasnoTrab = "611"
Global Const gsRHConcFuncConstNumMesTrab = "612"
Global Const gsRHConcFuncConstTERNumMesTrab = "613"
Global Const gsRHConcFuncConstPagoExeso = "614"
Global Const gsRHConcFuncConstReintegro = "615"
Global Const gsRHConcFuncConstSdoBas = "616"
Global Const gsRHConcFuncConstSdoBasNivel = "617"
Global Const gsRHConcFuncConstSdoContrato = "618"
Global Const gsRHConcFuncConstUltBonVacac = "619"
Global Const gsRHConcFuncConstUltGratif = "620"
Global Const gsRHConcFuncConstUltTercio = "621"
Global Const gsRHConcVar = "gsRHConcVar"
Global Const gsRHConcVarTotIngAcum = "701"
Global Const gsRHConcVarAFPCuotaFija = "702"
Global Const gsRHConcVarIngAntesAFP = "703"
Global Const gsRHConcVarMontoAMP = "704"
Global Const gsRHConcVarNumDiasLab = "705"
Global Const gsRHConcVarNumDiasMes = "706"
Global Const gsRHConcVarNumMinTrab = "707"
Global Const gsRHConcVarPrueba = "708"
Global Const gsRHConcVarPorcAsistMedFam = "709"
Global Const gsRHConcVarPorcONP = "710"
Global Const gsRHConcVarPorcFonavi = "711"
Global Const gsRHConcVarPorc5taCat = "712"
Global Const gsRHConcVarPorcSegSoc = "713"
Global Const gsRHConcVarPorcExtSolid = "714"
Global Const gsRHConcVarPorcIncONP = "715"
Global Const gsRHConcVarPorcRPS = "716"
Global Const gsRHConcVarPorcTercio = "717"
Global Const gsRHConcVarPromDia = "718"
Global Const gsRHConcVarSueldoNeto = "719"
Global Const gsRHConcVarSumMonInt = "720"
Global Const gsRHConcVarAFPTopePrima = "721"
Global Const gsRHConcVarAPFTope = "722"
Global Const gsRHConcVarTotalRemun = "723"
Global Const gsRHConcVarUITValor = "724"
Global Const gsRHConcVarUIT54 = "725"
Global Const gsRHConcVarUIT07 = "726"
Global Const gsRHConcVarDiaProm = "727"
Global Const gsRHTpoConc = "gsRHTpoConc"
Global Const gsRHTpoConcConst = "1"
Global Const gsRHTpoConcValCal = "2"
Global Const gsRHTpoConcValPreDef = "3"
Global Const gsRHTpoConcFunInt = "4"
Global Const gsRHConcEst = "gsRHConcEst"
Global Const gsRHConcEstAct = "1"
Global Const gsRHConcEstDesh = "0"
Global Const gsRHImp5taCat = "gsRHImp5taCat"
Global Const gsRHImp5taCatCDcto = "1"
Global Const gsRHImp5taCatSDcto = "0"
Global Const gsRHAfeCtoMesTrabajado = "gsRHAfeCtoMesTrabajado"
Global Const gsTpoCtaIf = "gsTpoCtaIf                           "
Global Const gsTpoCtaIfCtaCte = "01"
Global Const gsTpoCtaIfCtaAho = "02"
Global Const gsTpoCtaIfCtaPF = "03"
Global Const gsTpoCtaIfCtaAdeud = "05"
Global Const gsTpoIf = "gsTpoIf"
Global Const gsTpoIfBanco = "01"
Global Const gsTpoIfFinanciera = "02"
Global Const gsTpoIfCmac = "03"
Global Const gsTpoPagoAdeud = "gsTpoPagoAdeud"
Global Const gsTpoPagoAdeudCapital = "01"
Global Const gsTpoPagoAdeudInteres = "02"
Global Const gsTpoPagoAdeudGastos = "03"

