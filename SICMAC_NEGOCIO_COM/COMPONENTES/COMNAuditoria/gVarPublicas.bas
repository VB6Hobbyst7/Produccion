Attribute VB_Name = "gVarPublicas"
'Public oImpresora As New ContsImp.clsConstImp
Public oImpresora As New COMFunciones.FCOMVarImpresion
'Public gImpresora As Impresoras
'Public Const READ_BUFF As Long = 255
'Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Global gsRUC As String

'*******************************************************************
'************************ Constantes Públicas *******************
'*******************************************************************
Global Const gnNroDigitosDNI = 8
Global Const gnNroDigitosRUC = 11

Global gsRutaIcono As String

Global Const gsMenuAplicac = "2" 'para identificar los menus que le corresponden
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

'*******************************************************************
'************************ Variables Globales *******************
'*******************************************************************
Global estadoAccesoLogistica As Integer
Global gsDominio As String
Global gsPDC As String
Global gsCtaApe As String
Global gsDirBackup As String
Global gsServerName As String
Global gsUser As String
Global gsCodUser As String
Global gsNomUser As String
Global gsCodPersUser As String
Global gsCodPersCargo As String
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
Global gsMaquina As String
Global gnTipCambio As Currency
Global gnTipCambioV As Currency
Global gnTipCambioC As Currency
Global gnTipCambioVE As Currency
Global gnTipCambioCE As Currency
Global gnTipCambioPonderado As Currency

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

Global gnIGV As Currency

Global gcDocTpoOPago As String
Global gcDocTpoCargo As String
Global gcDocTpoCarta As String
Global gcDocTpoAbono As String
Global gcDocTpoCheque As String

Global Const gcMN = "S/."
Global Const gcME = "US$"

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
' Para Contabilidad
Global gbBitCentral As Boolean


'Informacion a Recordar
Public gbInfRecOCOSVencimiento As Boolean
Public gbPermisoLogProveedorAG As Boolean


'* Para el mantenimiento de Permisos
'Public Type TMatmenu
'    nId As Integer
'    sCodigo As String
'    sName As String
'    sCaption As String
'    sIndex As String
'    nNumHijos As Integer
'    bCheck As Boolean
'    nPuntDer As Integer
'    nPuntAbajo As Integer
'    nNivel As Integer
'End Type
'
'Public MatMenuItems() As TMatmenu
'Public MatOperac(2000, 5) As String
'Public NroRegOpe As Integer

'Private Function DamePosicionNivel(ByVal psName As String) As Integer
'Dim I As Integer
'Dim Y As Integer
'    Y = 1
'    For I = 4 To Len(psName) Step 2
'        If Mid(psName, I, 2) = "00" Then
'            DamePosicionNivel = Y
'            Exit For
'        End If
'        Y = Y + 1
'    Next I
'    If Y = 6 Then
'        DamePosicionNivel = 5
'    End If
'
'End Function
'
'Public Sub CargaMenuArbol(ByRef PR As ADODB.Recordset, ByRef nPunt As Integer, ByRef pnId As Integer)
'Dim nPos As Integer
'Dim nPos2 As Integer
'Dim nPuntTemp As Integer
'
'        If PR.EOF Then
'            Exit Sub
'        End If
'
'        'Obtengo el Nivel
'        nPos = DamePosicionNivel(MatMenuItems(nPunt - 1).sName)
'        nPos2 = DamePosicionNivel(PR!cname)
'
'        If nPos2 > nPos Then 'Es Hijo
'            pnId = pnId + 1
'            ReDim Preserve MatMenuItems(nPunt + 1)
'            MatMenuItems(nPunt).nId = pnId - 1
'            MatMenuItems(nPunt).sCodigo = Trim(PR!cCodigo)
'            MatMenuItems(nPunt).sCaption = Trim(PR!cDescrip)
'            MatMenuItems(nPunt).sName = Trim(PR!cname)
'            MatMenuItems(nPunt).sIndex = Right(PR!cname, 2)
'            MatMenuItems(nPunt).bCheck = False
'            MatMenuItems(nPunt).nPuntDer = -1
'            MatMenuItems(nPunt).nPuntAbajo = -1
'            MatMenuItems(nPunt - 1).nPuntDer = MatMenuItems(nPunt).nId
'            MatMenuItems(nPunt).nNivel = nPos2
'            PR.MoveNext
'
'            nPunt = nPunt + 1
'            Call CargaMenuArbol(PR, nPunt, pnId)
'
'        End If
'
'        If nPos2 = nPos Then 'Son del Mismo Menu
'            pnId = pnId + 1
'            nPuntTemp = nPunt
'            If pnId <> nPunt + 1 Then
'                nPunt = pnId - 1
'            End If
'
'            ReDim Preserve MatMenuItems(nPunt + 1)
'            MatMenuItems(nPunt).nId = pnId - 1
'            MatMenuItems(nPunt).sCodigo = Trim(PR!cCodigo)
'            MatMenuItems(nPunt).sCaption = Trim(PR!cDescrip)
'            MatMenuItems(nPunt).sName = Trim(PR!cname)
'            MatMenuItems(nPunt).sIndex = Right(PR!cname, 2)
'            MatMenuItems(nPunt).bCheck = False
'            MatMenuItems(nPunt).nPuntDer = -1
'            MatMenuItems(nPunt).nPuntAbajo = -1
'            MatMenuItems(nPuntTemp - 1).nPuntAbajo = MatMenuItems(nPunt).nId
'            MatMenuItems(nPunt).nNivel = nPos2
'            PR.MoveNext
'
'            nPunt = nPunt + 1
'            Call CargaMenuArbol(PR, nPunt, pnId)
'        End If
'
'        If nPos2 < nPos Then 'Es un Menu de nivel Anterior
'            nPunt = nPunt - 1
'            Call CargaMenuArbol(PR, nPunt, pnId)
'        End If
'
'End Sub


