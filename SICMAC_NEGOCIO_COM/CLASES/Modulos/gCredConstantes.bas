Attribute VB_Name = "gCredConstantes"
Option Explicit

'********************************************************
'Filtro de Constantes
Global Const gCredFiltroProd = 1 'Para Filtrar Todos los Productos de Creditos
Global Const gCredFiltroProdCab = 2 'Para Filtrar Todos los Cabeceras de Productos de Creditos
Global Const gCredFiltroCalendApl = 3 'Para Filtrar Aplicacion de Calendario A Cuota o Desembolso
Global Const gCredFiltroInstRepres = 4 'Para Filtrar Representantes de Institucion
Global Const gCredFiltroGastosOperad = 5 'Para filtrar operaciones de Comisiones
'********************************************************

'Constantes de Impresion a WORD

Global Const cPlantillaCartaAMoroso1 = "\FormatoCarta\CartaMorosoAvisoPago.doc"
Global Const cPlantillaCartaAMoroso2 = "\FormatoCarta\CartaMorosoRequerimientoPago.doc"
Global Const cPlantillaCartaAMoroso3 = "\FormatoCarta\CartaMorosoNotificacionPrejudicial.doc"
Global Const cPlantillaCartaAMoroso4 = "\FormatoCarta\CartaMorosoUltimaNotificacionPrejudicial.doc"
Global Const cPlantillaCartaAMoroso5 = "\FormatoCarta\CartaMorosoCartaNotarial.doc"
Global Const cPlantillaCartaInvCredParalelo = "\FormatoCarta\CartaInvCredParalelo.doc"
Global Const cPlantillaCartaRecup = "\FormatoCarta\CartaRecup.doc"

'Garantias
Global Const gsGarantLevanta = "160101"


'Aprobacion de Credito
Global Const gCredAprobacion = "101501"

'Desembolso en Efectivo
Global Const gCredDesembEfec = "100101"
'Desembolso Abono a Cuenta Nueva
Global Const gCredDesembCtaNueva = "100102"
'Desembolso Abono a Cuenta Existente
Global Const gCredDesembCtaExist = "100103"
'Desembolso Abono a Cuenta Existente de Otra Agencia
Global Const gCredDesembCtaExistDOA = "100104"
'Desembolso Abono a Cuenta Nueva de Otra Agencia
Global Const gCredDesembCtaNuevaDOA = "100105"

'Desembolso Retiro Gastos De Agencia Local
Global Const gCredDesembCtaRetiroGastos = "100106"
'Desembolso Retiro Gastos De Agencia Remota
Global Const gCredDesembCtaRetiroGastosDOA = "100107"
'Desembolso Retiro Cancelacion Credito Agencia Local
Global Const gCredDesembCtaRetiroCancelacion = "100108"
'Desembolso Retiro Cancelacion Credito Agencia Remota
Global Const gCredDesembCtaRetiroCancelacionDOA = "100109"


'Pagos Lote
Global Const gCredPagLote = "102100"

'Registro de Dacion
Global Const gCredRegisDacion = "102200"

'Pagos Normales
Global Const gCredPagNorNor = "100200"
Global Const gCredPagNorNorEfec = "100201"
Global Const gCredPagNorNorCC = "100202"
Global Const gCredPagNorNorCh = "100203"
Global Const gCredPagNorNorDctoPla = "100204"
'Llamada Reccepcion en Otra Cmact
Global Const gCredPagNorNorEOCEfec = "100205"
Global Const gCredPagNorNorEOCCh = "100206"
'Dacion en Pago
Global Const gCredPagNorNorDacion = "100207"

'Pagos RFA
Global Const gCredPagNorRfaEfec = "100209"

Global Const gCredPagNorMor = "100300"
Global Const gCredPagNorMorEfec = "100301"
Global Const gCredPagNorMorCC = "100302"
Global Const gCredPagNorMorCh = "100303"
Global Const gCredPagNorMorDctoPla = "100304"
'Llamada Reccepcion en Otra Cmact
Global Const gCredPagNorMorEOCEfec = "100305"
Global Const gCredPagNorMorEOCCh = "100306"
'Dacion en Pago
Global Const gCredPagNorMorDacion = "100307"

Global Const gCredPagNorVen = "100400"
Global Const gCredPagNorVenEfec = "100401"
Global Const gCredPagNorVenCC = "100402"
Global Const gCredPagNorVenCh = "100403"
Global Const gCredPagNorVenDctoPla = "100404"
'Llamada Reccepcion en Otra Cmact
Global Const gCredPagNorVenEOCEfec = "100405"
Global Const gCredPagNorVenEOCCh = "100406"
'Dacion en Pago
Global Const gCredPagNorVenDacion = "100407"

Global Const gCredPagRefNor = "100500"
Global Const gCredPagRefNorEfec = "100501"
Global Const gCredPagRefNorCC = "100502"
Global Const gCredPagRefNorCh = "100503"
Global Const gCredPagRefNorDctoPla = "100504"
'Llamada Reccepcion en Otra Cmact
Global Const gCredPagRefNorEOCEfec = "100505"
Global Const gCredPagRefNorEOCCh = "100506"
'Dacion en Pago
Global Const gCredPagRefNorDacion = "100507"

Global Const gCredPagRefMor = "100600"
Global Const gCredPagRefMorEfec = "100601"
Global Const gCredPagRefMorCC = "100602"
Global Const gCredPagRefMorCh = "100603"
Global Const gCredPagRefMorDctoPla = "100604"
'Llamada Reccepcion en Otra Cmact
Global Const gCredPagRefMorEOCEfec = "100605"
Global Const gCredPagRefMorEOCCh = "100606"
'Dacion en Pago
Global Const gCredPagRefMorDacion = "100607"

Global Const gCredPagRefVen = "100700"
Global Const gCredPagRefVenEfec = "100701"
Global Const gCredPagRefVenCC = "100702"
Global Const gCredPagRefVenCh = "100703"
Global Const gCredPagRefVenDctoPla = "100704"
'Llamada Reccepcion en Otra Cmact
Global Const gCredPagRefVenEOCEfec = "100705"
Global Const gCredPagRefVenEOCCh = "100706"
'Dacion en Pago
Global Const gCredPagRefVenDacion = "100707"

'Pago en Suspenso
Global Const gsPagIntSus = "100800"
Global Const gsPagIntSusEfec = "100801"


'*************************************
'******   Para Refinanciacion ********
'*************************************
Global Const gCredRefinanciacion = "101200"
Global Const gCredRefNormSoles = "101201"
Global Const gCredRefMorSoles = "101202"
Global Const gCredRefVencSoles = "101203"
Global Const gCredRefIngSoles = "101261"

Global Const gCredRefRefNormSoles = "101204"
Global Const gCredRefRefMorSoles = "101205"
Global Const gCredRefRefVencSoles = "101206"
Global Const gCredRefRefIngSoles = "101207"


Global Const gCredRefNormDolares = "101247"
Global Const gCredRefMorDolares = "101248"
Global Const gCredRefVencDolares = "101249"
Global Const gCredRefIngDolares = "101262"

Global Const gCredRefRefNormDolares = "101250"
Global Const gCredRefRefMorDolares = "101251"
Global Const gCredRefRefVencDolares = "101252"
Global Const gCredRefRefIngDolares = "101253"

Global Const gCredRefDMNormSoles = "101208"
Global Const gCredRefDMMorSoles = "101209"
Global Const gCredRefDMVencSoles = "101210"
Global Const gCredRefIngDMSoles = "101263"

Global Const gCredRefRefDMNormSoles = "101211"
Global Const gCredRefRefDMMorSoles = "101212"
Global Const gCredRefRefDMVencSoles = "101213"
Global Const gCredRefRefIngDMSoles = "101214"

Global Const gCredRefDMNormDolares = "101215"
Global Const gCredRefDMMorDolares = "101216"
Global Const gCredRefDMVencDolares = "101217"
Global Const gCredRefIngDMDolares = "101264"

Global Const gCredRefRefDMNormDolares = "101218"
Global Const gCredRefRefDMMorDolares = "101219"
Global Const gCredRefRefDMVencDolares = "101220"
Global Const gCredRefRefIngDMDolares = "101221"

Global Const gCredRefCapIntNormSoles = "101222"
Global Const gCredRefCapIntMorSoles = "101223"
Global Const gCredRefCapIntVencSoles = "101224"
Global Const gCredRefCapIntIngSoles = "101265"

Global Const gCredRefRefCapIntNormSoles = "101225"
Global Const gCredRefRefCapIntMorSoles = "101226"
Global Const gCredRefRefCapIntVencSoles = "101227"
Global Const gCredRefRefCapIntIngSoles = "101228"

Global Const gCredRefCapIntDMNormSoles = "101229"
Global Const gCredRefCapIntDMMorSoles = "101230"
Global Const gCredRefCapIntDMVencSoles = "101231"
Global Const gCredRefCapIntIngDMSoles = "101266"

Global Const gCredRefRefCapIntDMNormSoles = "101232"
Global Const gCredRefRefCapIntDMMorSoles = "101233"
Global Const gCredRefRefCapIntDMVencSoles = "101234"
Global Const gCredRefRefCapIntIngDMSoles = "101235"

Global Const gCredRefCapIntNormDolares = "101254"
Global Const gCredRefCapIntMorDolares = "101255"
Global Const gCredRefCapIntVencDolares = "101256"
Global Const gCredRefCapIntIngDolares = "101267"

Global Const gCredRefRefCapIntNormDolares = "101257"
Global Const gCredRefRefCapIntMorDolares = "101258"
Global Const gCredRefRefCapIntVencDolares = "101259"
Global Const gCredRefRefCapIntIngDolares = "101260"

Global Const gCredRefCapIntDMNormDolares = "101236"
Global Const gCredRefCapIntDMMorDolares = "101237"
Global Const gCredRefCapIntDMVencDolares = "101238"
Global Const gCredRefCapIntIngDMDolares = "101268"

Global Const gCredRefRefCapIntDMNormDolares = "101239"
Global Const gCredRefRefCapIntDMMorDolares = "101240"
Global Const gCredRefRefCapIntDMVencDolares = "101241"
Global Const gCredRefRefCapIntIngDMDolares = "101242"

Global Const gCredRefIntSuspenso = "101243"
Global Const gCredRefGastosSuspenso = "101244"
Global Const gCredRefRevIntSuspenso = "101245"
Global Const gCredRefRevGastosSuspenso = "101246"


Global Const gCredMiVivCondonacion = "101301"

Global Const gCredPasoARecup = "130100"

'EXTORNOS
Global Const gCredExtPago = "109001"
Global Const gCredExtDesemb = "109002"
Global Const gCredExtPagoLote = "109003"
Global Const gCredExtAprobacion = "109004" 'por desaparecer

'Llamada de Otra CMAC 1070
Global Const geCredPagodeOtCj = "107100"
'Global Const geCredCanceldeOtCj = "127400"

'Recepcion en Otra Cmac
Global Const geCredPagoEnOtCj = "126400"

'Garantias
Global Const gGarantLevanta = "160101"
Global Const gGarantExtorno = "160901"

'ITF
Global Const gCredITF = "990103"
Global Const gCredITFDesemb = "990106"

'ITF en Otra Cmac
Global Const gCredITFEOC = "990303"

'RFA
Global Const gColocOpeRFA = "100209"
Global Const gColoOpeExRFA = "109005"

Global Const gColocRFACapital = "1010"
Global Const gColocRFAInteresMoratorio = "1108"
Global Const gColocRFAInteresCompesatorio = "1107"
Global Const gColocRFAInteresDiferido = "1109"
Global Const gColocRFAInteresDiferidoRFA = "1110"

'Constante de Creditos Automaticos
Global Const gCredAutomatico = 3
