Attribute VB_Name = "MVarPublicas"
'*******************************************************************
'************************ Constantes Públicas *******************
'*******************************************************************




Global Const gnNroDigitosDNI = 8
Global Const gnNroDigitosRUC = 11
Global Const gsFormatoFecha = "mm/dd/yyyy"
Global Const gsFormatoFechaHora = "mm/dd/yyyy hh:mm:ss"
Global Const gsFormatoFechaHoraView = "dd/mm/yyyy hh:mm:ss"
Global Const gsFormatoFechaView = "dd/mm/yyyy"

Global Const gcFormatoFecha = "mm/dd/yyyy"
Global Const gcFormatoFechaHora = "mm/dd/yyyy hh:mm:ss"
Global Const gcFormatoFechaHoraView = "dd/mm/yyyy hh:mm:ss"
Global Const gcFormatoFechaView = "dd/mm/yyyy"

Global Const gsFormatoMovFecha = "yyyymmdd"
Global Const gsFormatoMovFechaHora = "yyyymmddhhmmss"

Global Const gsFormatoNumeroView = "##,###,##0.00##"
Global Const gcFormView = "##,###,##0.00##"
Global Const gsFormatoNumeroDato = "#######0.00##"
Global Const gcFormDato = "#######0.00##"

Global Const gcFormatoTC = "#0.00##"
Global Const gnColPage = 79   'Columnas por página de Impresión
Global Const gnLinVert = 66   'Orientación Vertical
Global Const gnLinHori = 46   'Orientación Horizontal
Global Const gcFormatoMov = "yyyymmdd"
Global Const gbComunRemoto = False
Global Const IDPlantillaOP = "OPBatch"
Global Const IDPlantillaVOP = "OPVEBatch"
Global Const gsUsuarioBOVEDA = "BOVE"
'*******************************************************************
'************************ Variables Globales *******************
'*******************************************************************
Global gsDominio As String
Global gsPDC As String
Global gsCtaApe As String
Global gsDirBackup As String
Global gsServerName As String
Global gsUser As String
Global gsCodUser As String
Global gsCodPersUser As String
Global gsCodAge As String
Global gsCodArea As String
Global gsNomAge As String
Global gsNomCmac As String
Global gsCodCMAC As String
Global gbOpeOk As Boolean
Global gsRutaFirmas As String
Global gsDBName As String
Global gsUID As String
Global gsPWD As String
Global gsCodPersCMACT As String
Global gbRetiroSinFirma As Boolean
Global gbAgeEsp As Boolean
Global gbVerificaRegistroEfectivo As Boolean
Global gnTipCambio As Currency
Global gnTipCambioV As Currency
Global gnTipCambioC As Currency
Global gsRutaIcono As String
Global gnValidaSolCredito As Integer
Global gnValidaGarantia As Integer

'Variable de la Fecha del sitema
Global gdFecSis As Date
'Variable de la Fecha del fin de Mes
Global gdFecData As Date
'Variable de la Fecha Data Consolidada de Fin de Mes
Global gdFecDataFM As Date
'Variables de Conection
Global gsCodAgeN As String
Global sLpt  As String
Global gsConnection As String
'* Archivo de Impresion para Previo
Global nFicSal As Integer
Global gsInstCmac As String


'*******************************************************************
'************** VARIABLES PUBLICAS DE BASES CENTRALIZADAS ******************************
'*******************************************************************

Global gsCentralPers As String
Global gsCentralImg As String
Global gsCentralCom As String


'*******************************************************************
'************** VARIABLES PUBLICAS DE CONTABILIDAD ******************************
'*******************************************************************
Public glAceptar As Boolean
Public BON   As String 'BOLD ON
Public BOFF  As String 'Bold off
Public Con   As String 'Condensado ON
Public COFF  As String 'Condensado OFF
Global gcEmpresa As String   'Entidad Financiera
Global gcEmpresaLogo As String  'Logo de la Entidad Financiera
Global gcEmpresaRUC As String   'RUC de la Entidad Financiera
Global gnDocTpoOPago  As String 'Codigo Tipo de Documento Orden de Pago
Global gnDocTpoCheque As String 'Código Tipo de Documento Cheque
Global gnDocTpoFac    As String 'Código Tipo de Documento Factura
Global gnDocTpoCarta  As String 'Código Tipo de Documento Carta
Global gnDocTpoAbono  As String 'Código Tipo de Documento Nota de Abono
Global gnTasaCajaCh As Currency
Global gnDocTpoCargo  As String
Global gcModuloLogo  As String

Global Const gsContDebe = "D"
Global Const gsContHaber = "H"
Global Const gsContDebeDesc = "Debe"
Global Const gsContHaberDesc = "Haber"

Global Const gsOpeCtaCaracterObligaDesc = "OBLIGATORIO"
Global Const gsOpeCtaCaracterOpcionDesc = "OPCIONAL"

Global Const gsSI = "SI"
Global Const gsNO = "NO"

Global Const gsMenuAplicac = "1"

'Variables de Contabilidad
Global gsMovNro As String
Global gnMovNro As Long
Global gsGlosa  As String
Global gnImporte As Currency
Global gdFecha  As Date

Global gsOpeCod As String
Global gsOpeDesc As String
Global gsOpeDescPadre As String
Global gsOpeDescHijo As String

Global gsSimbolo As String

Global gcTitModulo    As String
Global glDiaCerrado As Boolean
Global gcCtaIGV   As String
Global gcDocTpoFac As String

Global gcDocTpoOPago As String
Global gcDocTpoCargo As String
Global gcDocTpoCarta As String
Global gcDocTpoAbono As String
Global gcDocTpoCheque As String

Global gcMN  As String
Global gcME  As String

Global gcMNDig  As String
Global gcMEDig  As String

Global gnMgSup   As Integer
Global gnMgIzq  As Integer
Global gnMgDer  As Integer
Global gnLinPage   As Integer
Global gnArendirImporte  As Currency
Global gnLinPageOP   As Integer
Global gcConvMED  As String
Global gcConvMES  As String
Global gcConvTipo   As String
Global gcCtaCaja   As String
Global gcCCHCta As String
Global gnEncajeExig  As Currency
Global gnTotalOblig  As Currency
Global gsCtaBancoMN  As String
Global gsCtaBancoME  As String
Global gsCtaBCRMN  As String
Global gsCtaBCRME  As String
Global gsCodAdeudado As String
Global gaObj() As String
Global gsDirPlantillas As String


Global GmyPSerial As Object         'ppoa   variable Puerto
Global GnTipoPinPad  As TipoPinPad     'Marca y Modelo  de PinPad

'____________________________________________________________
'
'*********VARIABLES PARA PERSONA DE LAVADO DE DINERO*********
'____________________________________________________________

Global gReaPersLavDinero As String
Global gBenPersLavDinero As String

'***************************





