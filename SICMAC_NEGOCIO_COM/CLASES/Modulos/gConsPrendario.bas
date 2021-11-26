Attribute VB_Name = "gConsPrendario"
Option Explicit



'***************************************************************
'Cadenas de Plantillas txt
Global Const cPlantillaActaAdju = "\FormatoCarta\CartaActaAdjudica.txt"
Global Const cPlantillaActaSuba = "\FormatoCarta\CartaActaSubasta.txt"
Global Const cPlantillaActaRema = "\FormatoCarta\CartaActaRemate.txt"
Global Const cPlantillaAvisoRemate = "\FormatoCarta\CartaAvisoRemate.txt"
Global Const cPlantillaAvisoSubasta = "\FormatoCarta\CartaAvisoSubasta.txt"
Global Const cPlantillaAvisoSobrante = "\FormatoCarta\CartaAvisoSobrante.txt"
Global Const cPlantillaAvisoVencimiento = "\FormatoCarta\CartaAvisoVencimiento.txt"
'************************************************************************************

''FreeFile de impresión
'Global ArcSal As Integer
'Hora para las grabaciones
'Global gdHoraGrab As Date
'************************************************************************************
'Ultima Ventas de Adjudicados a Barras
'Global Const cFecUltVentaBarras = "08/08/1998"
'************************************************************************************
'Public Const cMinPesoOro As Single = 1#
'Public Const cDiasParaRemate As Integer = 30
'************************************************************************************
'Colores para montos
'Fondo.
' Soles -> Blanco
'Global Const pColFonSoles = &H80000005
' Dolares -> Verde
'Global Const pColFonDolares = &HFF00&
'Primer Plano
' Ingreso -> Negro
'Global Const pColPriIngreso = &HFF0000   '&H0&
' Egreso -> Rojo
'Global Const pColPriEgreso = &HFF&
'************************************************************************************
'** CONSTANTES de los Codigo de Operacion **
' Registro Contrato Prendario
Global Const gsRegContrato = "030100" ' **
Global Const gsRegMonPres = "030101"
'Global Const gsRegIntAde = "030102"
'Global Const gsRegImp = "030103"
'Global Const gsRegTas = "030104"
'Global Const gsRegCus = "030105"
Global Const gsRegContGar = "030106"
Global Const gsRegOro14 = "030107"
Global Const gsRegOro16 = "030108"
Global Const gsRegOro18 = "030109"
Global Const gsRegOro21 = "030110"

' Desembolso de Credito Prendario
Global Const gsDesPrestamo = "030200" ' **
Global Const gsDesMonPag = "030201"
Global Const gsDesIntAde = "030202"
Global Const gsDesImp = "030203"
Global Const gsDesTas = "030204"
Global Const gsDesCus = "030205"

' Cancelacion/Rescate Normal de Credito Prendario
Global Const gsCanNorPrestamo = "030800" ' **
Global Const gsCanNorCap = "030801"
Global Const gsCanNorValTas = "030802"


' Cancelacion/Rescate Moroso de Credito Prendario
Global Const gsCanMorPrestamo = "031700"  ' **
Global Const gsCanMorCapPag = "031701"
Global Const gsCanMorIntMor = "031702"
Global Const gsCanMorCusMor = "031703"
Global Const gsCanMorImp = "031704"
Global Const gsCanMorPreRem = "031705"
Global Const gsCanMorValTas = "031706"

' Renovacion de Credito Prendario
Global Const gsRenPrestamo = "030500"  ' **
Global Const gsRenCapPag = "030501"
Global Const gsRenIntMor = "030502"
Global Const gsRenCusMor = "030503"
Global Const gsRenPreRem = "030504"
Global Const gsRenIntAde = "030505"
Global Const gsRenImp = "030506"
Global Const gsRenCus = "030507"

' Renovacion de Credito Prendario - Morosa
Global Const gsRenMorPrestamo = "031800"  ' **
Global Const gsRenMorCapPag = "031801"
Global Const gsRenMorIntMor = "031802"
Global Const gsRenMorCusMor = "031803"
Global Const gsRenMorPreRem = "031804"
Global Const gsRenMorIntAde = "031805"
Global Const gsRenMorImp = "031806"
Global Const gsRenMorCus = "031807"
Global Const gsRenMorCamCar = "031808"
Global Const gsRenMorValTasac = "031809"

' Anulacion de Prestamo Prendario
Global Const gsAnuContrato = "034900"
Global Const gsAnuCon = "034901"

' Anulacion de No Desembolsados
Global Const gsAnuNoDesemb = "030300"
Global Const gsAnuNoDes = "030301"

' Imprimir Duplicado
Global Const gsImpDuplicado = "031200"
Global Const gsImpDup = "031201"

' Imprimir Duplicado de Otra Agencia
Global Const gsImpDuplicadoDOA = "035400"
Global Const gsImpDupDOA = "035401"

' Imprimir Duplicado En Otra Agencia
Global Const gsImpDuplicadoEOA = "035500"
Global Const gsImpDupEOA = "035501"

' Devolucion de Prendas
Global Const gsDevJoyas = "031100"
Global Const gsDevJoyGar = "031101"
Global Const gsDevOro14 = "031102"
Global Const gsDevOro16 = "031103"
Global Const gsDevOro18 = "031104"
Global Const gsDevOro21 = "031105"

' Devolucion de Prendas De Otra Agencia
Global Const gsDevJoyasDOA = "035700"
Global Const gsDevJoyGarDOA = "035701"

' Devolucion de Prendas En Otra Agencia
Global Const gsDevJoyasEOA = "035800"
Global Const gsDevJoyGarEOA = "035801"
Global Const gsDevOro14EOA = "035802"
Global Const gsDevOro16EOA = "035803"
Global Const gsDevOro18EOA = "035804"
Global Const gsDevOro21EOA = "035805"

' Modificación de descripción de lote de Contrato
Global Const gsModContrato = "036400"

' Pago de Sobrantes
Global Const gsPagSobrante = "037400"
Global Const gsPagSob = "037401"
Global Const gsPagSobEnOtAg = "037402"
Global Const gsPagSobDeOtAg = "037403"

'Abono a cuenta de Ahorros
Global Const gsAboSobCta = "037200"

' Remate de Contratos
Global Const gsRemContrato = "037000"

' Venta Lotes en Remate
Global Const gsVtaRemate = "031300"
Global Const gsVtaRemSalCap = "031301"
Global Const gsVtaRemIntMor = "031302"
Global Const gsVtaRemCusMor = "031303"
Global Const gsVtaRemPreRem = "031304"
Global Const gsVtaRemCosRem = "031305"
Global Const gsVtaRemImp = "031306"
Global Const gsVtaRemSob = "031307"
Global Const gsVtaRemValTas = "031308"
Global Const gsVtaRemOro14 = "031309"
Global Const gsVtaRemOro16 = "031310"
Global Const gsVtaRemOro18 = "031311"
Global Const gsVtaRemOro21 = "031312"

' Adjudicaciones de Creditos
Global Const gsAdjudica = "031600"
Global Const gsAdjCap = "031601"
Global Const gsAdjIntMor = "031602"
Global Const gsAdjCusMor = "031603"
Global Const gsAdjPreRem = "031604"
Global Const gsAdjSobrant = "031605"
Global Const gsAdjValTas = "031606"
Global Const gsAdjOro14 = "031607"
Global Const gsAdjOro16 = "031608"
Global Const gsAdjOro18 = "031609"
Global Const gsAdjOro21 = "031610"

' Venta Adjudicados en Subasta
Global Const gsVtaSubasta = "031400"
Global Const gsVtaSubVtaNeta = "031401"
Global Const gsVtaSubImp = "031402"
Global Const gsVtaSubCostVtaAdj = "031403"
Global Const gsVtaSubJoyAdj = "031404"
Global Const gsVtaSubOro14 = "031405"
Global Const gsVtaSubOro16 = "031406"
Global Const gsVtaSubOro18 = "031407"
Global Const gsVtaSubOro21 = "031408"

' Cobrar Custodia Diferida
Global Const gsCobCusDiferida = "036000"
Global Const gsCobCusDif = "036001"
Global Const gsCobCusImp = "036002"

' Cobrar Custodia Diferida En Otra Agencia
Global Const gsCobCusDiferidaEOA = "036100"
Global Const gsCobCusDifEOA = "036101"
Global Const gsCobCusImpEOA = "036102"

' Cobrar Custodia Diferida De Otra Agencia
Global Const gsCobCusDiferidaDOA = "036200"
Global Const gsCobCusDifDOA = "036201"

'************************************************
'******* OPERACIONES CON OTRAS AGENCIAS  ********
'************************************************
'****** OPERACIONES EN OTRA AGENCIA
'**** Renovacion En Otra Agencia
Global Const gsRenEnOtAg = "032000"
Global Const gsRenEnOtAgCapPag = "032001"
Global Const gsRenEnOtAgIntMor = "032002"
Global Const gsRenEnOtAgCusMor = "032003"
Global Const gsRenEnOtAgPreRem = "032004"
Global Const gsRenEnOtAgIntAde = "032005"
Global Const gsRenEnOtAgImp = "032006"
Global Const gsRenEnOtAgCus = "032007"

'**** Renovacion En Otra Agencia - Morosa
Global Const gsRenMorEnOtAg = "032050"
Global Const gsRenMorEnOtAgCapPag = "032051"
Global Const gsRenMorEnOtAgIntMor = "032052"
Global Const gsRenMorEnOtAgCusMor = "032053"
Global Const gsRenMorEnOtAgPreRem = "032054"
Global Const gsRenMorEnOtAgIntAde = "032055"
Global Const gsRenMorEnOtAgImp = "032056"
Global Const gsRenMorEnOtAgCus = "032057"
Global Const gsRenMorEnOtAgCamCar = "032058"
Global Const gsRenMorEnOtAgValTasac = "032059"

'**** Cancelacion En Otra Agencia
Global Const gsCanNorEnOtAgPrestamo = "033200"
Global Const gsCanNorEnOtAgCapPag = "033201"
Global Const gsCanNorEnOtAgValTas = "033202"

'**** Cancelacion Morosa En Otra Agencia
Global Const gsCanMorEnOtAgPrestamo = "033500"
Global Const gsCanMorEnOtAgCapPag = "033501"
Global Const gsCanMorEnOtAgIntMor = "033502"
Global Const gsCanMorEnOtAgCusMor = "033503"
Global Const gsCanMorEnOtAgImp = "033504"
Global Const gsCanMorEnOtAgPreRem = "033505"
Global Const gsCanMorEnOtAgValTas = "033506"

'****** OPERACIONES DE OTRA AGENCIA
'****  Renovacion De Otra Agencia
Global Const gsRenDeOtAg = "034100"
Global Const gsRenDeOtAgMonTra = "034101"

'****  Cancelacion De Otra Agencia
Global Const gsCanNorDeOtAg = "034400"
Global Const gsCanNorDeOtAgMonTra = "034401"

'****  Cancelacion Morosa De Otra Agencia
Global Const gsCanMorDeOtAg = "034700"
Global Const gsCanMorDeOtAgMonTra = "034701"

'****** CONSTANTES DE EXTORNOS
'****  Cancelación/Rescate de contrato
Global Const gsExtCanPrestamo = "039000"
Global Const gsExtCanCapPag = "039001"
Global Const gsExtCanIntMor = "039002"
Global Const gsExtCanCusMor = "039003"
Global Const gsExtCanImp = "039004"
Global Const gsExtCanPreRem = "039005"
Global Const gsExtCanValTas = "039006"

'****  Renovación de contrato
Global Const gsExtRenPrestamo = "039100"
Global Const gsExtRenCapPag = "039101"
Global Const gsExtRenIntMor = "039102"
Global Const gsExtRenCusMor = "039103"
Global Const gsExtRenPreRem = "039104"
Global Const gsExtRenIntAde = "039105"
Global Const gsExtRenImp = "039106"
Global Const gsExtRenCus = "039107"
Global Const gsExtRenMorCamCar = "039108"
Global Const gsExtRenMorValTasac = "039109"

'****  Devolución de Prendas
Global Const gsExtDevJoyas = "039200"
Global Const gsExtDevJoyGar = "039201"
Global Const gsExtDevOro14 = "039202"
Global Const gsExtDevOro16 = "039203"
Global Const gsExtDevOro18 = "039204"
Global Const gsExtDevOro21 = "039205"

'****  Venta de Remate De Otra Agencia
Global Const gsVtaRemDeOtAg = "033700"
Global Const gsVtaRemDeOtAgEfe = "033701"

'****  Venta de Remate En Otra Agencia
Global Const gsVtaRemEnOtAg = "035000"
Global Const gsVtaRemEnOtAgSalCap = "035001"
Global Const gsVtaRemEnOtAgIntMor = "035002"
Global Const gsVtaRemEnOtAgCusmor = "035003"
Global Const gsVtaRemEnOtAgPreRem = "035004"
Global Const gsVtaRemEnOtAgCosRem = "035005"
Global Const gsVtaRemEnOtAgImp = "035006"
Global Const gsVtaRemEnOtAgSob = "035007"
Global Const gsVtaRemEnOtAgValTas = "035008"
Global Const gsVtaRemEnOtAgOro14 = "035009"
Global Const gsVtaRemEnOtAgOro16 = "035010"
Global Const gsVtaRemEnOtAgOro18 = "035011"
Global Const gsVtaRemEnOtAgOro21 = "035012"

'****  Venta de Subasta De Otra Agencia
Global Const gsVtaSubDeOtAg = "033800"
Global Const gsVtaSubDeOtAgEfe = "033801"

'****  Venta de Subasta En Otra Agencia
Global Const gsVtaSubEnOtAg = "035200"
Global Const gsVtaSubEnOtAgVtaNeta = "035201"
Global Const gsVtaSubEnOtAgImp = "035202"
Global Const gsVtaSubEnOtAgCostVtaAdj = "035203"
Global Const gsVtaSubEnOtAgJoyAdj = "035204"
Global Const gsVtaSubEnOtAgOro14 = "035205"
Global Const gsVtaSubEnOtAgOro16 = "035206"
Global Const gsVtaSubEnOtAgOro18 = "035207"
Global Const gsVtaSubEnOtAgOro21 = "035208"

'*********  OPERACIONES CON OTRAS CMACT
'**** Renovacion En Otra CMAC
Global Const gsRenEnOtCj = "032100"
Global Const gsRenEnOtCjCapPag = "032101"
Global Const gsRenEnOtCjIntMor = "032102"
Global Const gsRenEnOtCjCusMor = "032103"
Global Const gsRenEnOtCjPreRem = "032104"
Global Const gsRenEnOtCjIntAde = "032105"
Global Const gsRenEnOtCjImp = "032106"
Global Const gsRenEnOtCjCus = "032107"

'**** Renovacion En Otra CMAC - Morosa
Global Const gsRenMorEnOtCj = "032200"
Global Const gsRenMorEnOtCjCapPag = "032201"
Global Const gsRenMorEnOtCjIntMor = "032202"
Global Const gsRenMorEnOtCjCusMor = "032203"
Global Const gsRenMorEnOtCjPreRem = "032204"
Global Const gsRenMorEnOtCjIntAde = "032205"
Global Const gsRenMorEnOtCjImp = "032206"
Global Const gsRenMorEnOtCjCus = "032207"
Global Const gsRenMorEnOtCjCamCar = "032208"
Global Const gsRenMorEnOtCjValTasac = "032209"

'****  Cancelacion Contrato En Otra CMAC
Global Const gsCanNorEnOtCjPrestamo = "033300"
Global Const gsCanNorEnOtCjCapPag = "033301"
Global Const gsCanNorEnOtCjValTas = "033302"

'**** Cancelacion Morosa En Otra Agencia
Global Const gsCanMorEnOtCjPrestamo = "033600"
Global Const gsCanMorEnOtCjCapPag = "033601"
Global Const gsCanMorEnOtCjIntMor = "033602"
Global Const gsCanMorEnOtCjCusMor = "033603"
Global Const gsCanMorEnOtCjImp = "033604"
Global Const gsCanMorEnOtCjPreRem = "033605"
Global Const gsCanMorEnOtCjValTas = "033606"

' Cambio de Cartera en el Cierre Diario
Global Const gsPigCamCarNorMor = "036500"
Global Const gsPigCanValTasVigVen = "036502"

