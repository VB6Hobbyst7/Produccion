Attribute VB_Name = "gVarPublicas"
'*******************************************************************
'************************ Constantes Públicas *******************
'*******************************************************************
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

Global Const gsRHCredAdmCod = "214"
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
Global gsNomAge As String * 30
Global gsNomCmac As String
Global gsCodCMAC As String
Global gbOpeOk As Boolean
Global gsRutaFirmas As String
Global gsDBName As String
Global gsUID As String
Global gsPWD As String

Global gnTipCambio As Currency
Global gnTipCambioV As Currency
Global gnTipCambioC As Currency

'Variable de la Fecha del sitema
Global gdFecSis As Date
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
Public CON   As String 'Condensado ON
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
Global gsReciboEgreso As String

'Documentos
Public gnDocTpo  As Long
Public gsDocDesc As String
Public gsDocNro  As String

Public Sub CargaVarSistema(ByVal pbContabilidad As Boolean)
    Dim lsQrySis As String
    Dim rsQrySis As New ADODB.Recordset
    Dim oconect As DConecta
    Dim VSQL As String
    Dim lnStrConn As String
    Dim lnPosIni As Integer
    Dim lnPosFin As Integer
    Dim lnStr As String
    Set oconect = New DConecta
    
    If oconect.AbreConexion(gsConnection) = False Then
        Exit Sub
    End If
    
    lsQrySis = " Select * From ConstSistema " _
            & " Where nConsSisCod in (" & gConstSistFechaSistema & "," & gConstSistNombreAbrevCMAC & "," _
            & gConstSistRutaBackup & "," & gConstSistCodCMAC & "," & gConstSistMargenSupCartas & "," & gConstSistMagenIzqCartas & "," _
            & gConstSistMargenDerCartas & "," & gConstSistNroLineasPagina & "," & gConstSistNroLineasOrdenPago & "," _
            & gConstSistCtaConversionMEDol & "," & gConstSistCtaConversiónMESoles & "," & gConstSistTipoConverión & "," _
            & gConstSistNombreModulo & "," & gConstSistFechaInicioDia & "," & gConstSistDominio & "," & gConstSistPDC & ",40) ORDER BY nConsSisCod"
    
    Set rsQrySis = oconect.CargaRecordSet(lsQrySis)
    If rsQrySis.BOF Or rsQrySis.EOF Then
       rsQrySis.Close
       Set rsQrySis = Nothing
       MsgBox "Tabla VarSistema está vacia", vbInformation, "Aviso"
       gsNomAge = ""
       gdFecSis = ""
       gsInstCmac = ""
       gsNomCmac = ""
       gsCodCMAC = ""
       Exit Sub
    End If
    Do While Not rsQrySis.EOF
        Select Case Trim(rsQrySis!nConsSisCod)
                Case gConstSistFechaSistema
                        gdFecSis = CDate(Trim(rsQrySis!nConsSisValor))
                Case gConstSistNombreAbrevCMAC
                        gsInstCmac = Trim(rsQrySis!nConsSisValor)
                        gsNomCmac = Trim(rsQrySis!nConsSisDesc)
                Case gConstSistCodCMAC
                        gsCodCMAC = Trim(rsQrySis!nConsSisValor)
                Case gConstSistRutaBackup
                        gsDirBackup = Trim(rsQrySis!nConsSisValor)
                Case gConstSistNombreModulo '   "cEmpresa":
                        gcTitModulo = Trim(rsQrySis!nConsSisDesc)
                        gcModuloLogo = Trim(rsQrySis!nConsSisValor)
                Case gConstSistMargenSupCartas  ' "nMargSup":
                        gnMgSup = Val(rsQrySis!nConsSisValor)
                Case gConstSistMagenIzqCartas  ' "nMargIzq":
                        gnMgIzq = Val(rsQrySis!nConsSisValor)
                Case gConstSistMargenDerCartas  '  "nMargDer":
                        gnMgDer = Val(rsQrySis!nConsSisValor)
                Case gConstSistNroLineasPagina  '      "nLinPage":
                        gnLinPage = Val(rsQrySis!nConsSisValor)
                Case gConstSistNroLineasOrdenPago  '   "nLinPageOP":
                        gnLinPageOP = Val(rsQrySis!nConsSisValor)
                Case gConstSistCtaConversionMEDol '   "cConvMED":
                        gcConvMED = Trim(rsQrySis!nConsSisValor)
                Case gConstSistCtaConversiónMESoles '    "cConvMES":
                        gcConvMES = Trim(rsQrySis!nConsSisValor)
                Case gConstSistTipoConverión  '  "cConvTipo":
                        gcConvTipo = Trim(rsQrySis!nConsSisValor)
                Case gConstSistDominio
                        gsDominio = Trim(rsQrySis!nConsSisValor)
                Case gConstSistPDC
                        gsPDC = Trim(rsQrySis!nConsSisValor)
                Case 40
                    gcCtaIGV = Trim(rsQrySis!nConsSisValor)
        End Select
        rsQrySis.MoveNext
    Loop
    rsQrySis.Close
    Set rsQrySis = Nothing
    
    'Deduce el nombre del Servidor
    
    gsServerName = oconect.servername
    'Deduce el nombre de la Base de Datos
    gsDBName = oconect.DatabaseName
    lnStrConn = oconect.CadenaConexion
    'Deduce el nombre de usuario
    lnPosIni = InStr(1, lnStrConn, "UID=", vbTextCompare)
    If lnPosIni > 0 Then
        lnPosFin = InStr(lnPosIni, lnStrConn, ";", vbTextCompare)
        lnStr = Mid(lnStrConn, lnPosIni, lnPosFin - lnPosIni)
        lnPosIni = InStr(1, lnStr, "=", vbTextCompare)
        gsUID = Mid(lnStr, lnPosIni + 1, Len(lnStr) - lnPosIni)
    Else
        gsUID = ""
    End If
    'Deduce el password
    lnPosIni = InStr(1, lnStrConn, "PWD=", vbTextCompare)
    If lnPosIni > 0 Then
        lnPosFin = InStr(lnPosIni, lnStrConn, ";", vbTextCompare)
        lnStr = Mid(lnStrConn, lnPosIni, lnPosFin - lnPosIni)
        lnPosIni = InStr(1, lnStr, "=", vbTextCompare)
        gsPWD = Mid(lnStr, lnPosIni + 1, Len(lnStr) - lnPosIni)
    Else
        gsPWD = ""
    End If
    oconect.CierraConexion
    Set oconect = Nothing
    
    gcMN = "S/."
    gcME = "$"
    
    gcMNDig = "1"
    gcMEDig = "2"
End Sub


