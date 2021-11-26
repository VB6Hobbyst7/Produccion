Attribute VB_Name = "gVarPublicas"
Option Explicit

Public oImpresora As New ContsImp.clsConstImp
Public gImpresora As Impresoras


'*******************************************************************
'************************ Constantes Públicas *******************
'*******************************************************************


Global Const gnNroDigitosDNI = 8
Global Const gnNroDigitosRUC = 11
Global Const gsFormatoFecha = "mm/dd/yyyy"
Global Const gsFormatoFechaHora = "mm/dd/yyyy hh:mm:ss"
Global Const gsFormatoFechaHoraView = "dd/mm/yyyy hh:mm:ss"
Global Const gsFormatoFechaHoraViewAMPM = "dd/mm/yyyy hh:mm:ss AMPM"
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
Global gsGruposUser As String 'EJVG20120419
Global gsCodUser As String
Global gsCodPersUser As String
Global gsNomPersUser As String 'JUEZ 20160405
Global gsCodAge As String
Global gsCodArea As String
Global gsNomArea As String 'JUEZ 20160405
Global gsCodCargo As String
Global gsNomCargo As String 'JUEZ 20160405
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
Global gsServPindVerify As String
Global gsPVKi As String

Global gsTitutloOP_Soles As String
Global gsTitutloOP_Dolares As String

Global gnDocumentoGarantia As Integer 'ARCV 31-10-2006

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
'Variable Global de Impresora TMU
Global gbImpTMU As Boolean
'Variable GLobal de Revision de Clasificacion del MODULO DE AUDITORIA
Global gRevisionId As Integer
Global gbBitCentral As Boolean

Global gSesionDirId As Integer 'MAVM 20100106
Global gNroAcuerdo As Integer 'MAVM 20100106
Global gNroInformeId As Integer 'MAVM 20100106
Global gMedidasCorrectivasId As Integer 'MAVM 20100106

Global gnAgenciaCredEval As Integer 'JUEZ 20121219
Global gnAgenciaHojaRutaNew As Boolean 'WIOR 20151109

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
Global gnTipoPinPad  As TipoPinPad     'Marca y Modelo  de PinPad

Global gnIGVValor As Currency ' RECO20131114

'** DAOR 20081125 *******************************
Global gnPinPadPuerto As Integer
Global gsNomMaquinaUsu As String
Global Const gIpPuertoPinVerifyPOS = "192.168.0.9:81" '"192.168.15.35:81"
Global Const gWKPOS = "81AE036D7855A288" '"2222222222222222"
Global Const gNMKPOS = 0
Global Const gCanalIdPOS = "_02"
Global Const gCanalIdATM = "_01"
'************************************************

'____________________________________________________________
'
'*********VARIABLES PARA PERSONA DE LAVADO DE DINERO*********
'____________________________________________________________

Global gReaPersLavDinero As String
Global gBenPersLavDinero As String
Global gBenPersLavDinero2 As String 'JACA 20110223
Global gBenPersLavDinero3 As String 'JACA 20110223
Global gBenPersLavDinero4 As String 'JACA 20110223
Global gOrdPersLavDinero As String 'By Capi 20012008
Global gVisPersLavDinero As String 'DAOR 20070511
Global gnTipoREU As Integer 'ALPA 20081003
Global gnMontoAcumulado As Double 'ALPA 20081003
Global gsOrigen As String 'ALPA 20081003

'***************************

'*********************************************'
'***** Variables para manejo de Tarjetas *****'
'************* GITU 2011-04-19 ***************'
'*********************************************'

Global gnTimeOutAg As Integer
Global gnCodOpeTarj As Integer

'Documentos
Public gnDocTpo  As Long
Public gsDocDesc As String
Public gsDocNro  As String

'******* VARIABLES DE CORREO *******
Global gsCorreoHost As String 'JUEZ 20160405
Global gsCorreoEnvia As String 'JUEZ 20160405
'***********************************

'*******VARIABLES DE MONEDA****************
'MARG ERS044-2016
Global Const gcPEN_SINGULAR = "SOL"
Global Const gcPEN_PLURAL = "SOLES"
Global Const gcPEN_SIMBOLO = "S/"

'******************************************

'***VARIABLE CONTADOR MENSAJE DE VISTO CONTINUIDAD ADMITIDO***
'MARG ERS046-2016
Global gnCountVCAdmitido As Integer
'*************************************************************

'**** ANDE 20170602 variable que almacena el usuario que dio visto en Arqueo de expedientes de ahorro
Global gcUsuarioVistoArqExpAho As String

'***** ande 20170914 para verificar si un usuario tiene firma al retirar, se hace uso en el control ImageDB metodo CargarFirma
Global gbTieneFirma As Boolean
'***** ande 20171011 variable que permite verificar si la firma fue actualizada.
Global gbFirmaActualizada As Boolean


'* Para el mantenimiento de Permisos
Public Type TMatmenu

    nId As Integer
    sCodigo As String
    sName As String
    sCaption As String
    sIndex As String
    nNumHijos As Integer
    bCheck As Boolean
    nPuntDer As Integer
    nPuntAbajo As Integer
    nNivel As Integer
End Type
Public MatMenuItems() As TMatmenu
'Public MatOperac(2000, 5) As String
Public MatOperac() As String
Public NroRegOpe As Integer

'Para el manejo de las Operaciones F2 y Extornos
Public gRsOpeF2 As ADODB.Recordset
Public gRsExtornos As ADODB.Recordset
Public gRsOpeCMACRecep As ADODB.Recordset
Public gRsOpeCMACLlam As ADODB.Recordset

'EJVG20161116 ***
Public Type TAtajoTeclado
    bOperaciones As Boolean
    bRenovaPigno As Boolean
    bRetiroIntDPF As Boolean
    bHabEfectivo As Boolean
    bDebitoSP As Boolean
    bCancelaDPF As Boolean
    bCancelaPigno As Boolean
    bCompraME As Boolean
    bRetiroEfectAho As Boolean
    bConfirHab As Boolean
    bDesembEfect As Boolean
    bDesembCta As Boolean
    bDepoEfectAHO As Boolean
    bVentaME As Boolean
    bAperGiro As Boolean
    bRetiroCTS As Boolean
    bPagoSERV As Boolean
    bPagoCRED As Boolean
    bCuadre As Boolean
    bAumentoCapDPF As Boolean
    bCancelGiro As Boolean
    bPosicion As Boolean
End Type
Public gAtajoTeclado As TAtajoTeclado
Public gbAtajoActivo As Boolean
'END EJVG *******
'ARCV 20-07-2006
Public gRsOpeRepo As ADODB.Recordset
'---------------

'Para el manejo de los Proyectos en las diferentes Empresas
Public gsProyectoActual As String
Public Type TCastigoSobFal
    nOpeCod As Long
    sUser As String
    nMonto As Currency
    nMovNro As Long
    sAge As String
    bEstado As Integer
   nMovNro2 As Long
End Type
Public Function LimpiaVarLavDinero()
    gReaPersLavDinero = ""
    gBenPersLavDinero = ""
    gnTipoREU = 0
End Function



'' PinPad
''Public myPSerial                As HCOMPINPADLib.Pinpad
Private Function DamePosicionNivel(ByVal psName As String) As Integer
Dim i As Integer
Dim Y As Integer
    Y = 1
    For i = 4 To Len(psName) Step 2
        If Mid(psName, i, 2) = "00" Then
            DamePosicionNivel = Y
            Exit For
        End If
        Y = Y + 1
    Next i
    If Y = 6 Then
        DamePosicionNivel = 5
    End If
    
End Function

Public Sub CargaMenuArbol(ByRef pR As ADODB.Recordset, ByRef nPunt As Integer, ByRef pnId As Integer)
Dim nPos As Integer
Dim nPos2 As Integer
Dim nPuntTemp As Integer

        If pR.EOF Then
            Exit Sub
        End If
        
        'Obtengo el Nivel
        nPos = DamePosicionNivel(MatMenuItems(nPunt - 1).sName)
        nPos2 = DamePosicionNivel(pR!cname)
        
        If nPos2 > nPos Then 'Es Hijo
            pnId = pnId + 1
            ReDim Preserve MatMenuItems(nPunt + 1)
            MatMenuItems(nPunt).nId = pnId - 1
            MatMenuItems(nPunt).sCodigo = Trim(pR!cCodigo)
            MatMenuItems(nPunt).sCaption = Trim(pR!cDescrip)
            MatMenuItems(nPunt).sName = Trim(pR!cname)
            MatMenuItems(nPunt).sIndex = Right(pR!cname, 2)
            MatMenuItems(nPunt).bCheck = False
            MatMenuItems(nPunt).nPuntDer = -1
            MatMenuItems(nPunt).nPuntAbajo = -1
            MatMenuItems(nPunt - 1).nPuntDer = MatMenuItems(nPunt).nId
            MatMenuItems(nPunt).nNivel = nPos2
            pR.MoveNext
            
            nPunt = nPunt + 1
            Call CargaMenuArbol(pR, nPunt, pnId)
            
        End If
        
        If nPos2 = nPos Then 'Son del Mismo Menu
            pnId = pnId + 1
            nPuntTemp = nPunt
            If pnId <> nPunt + 1 Then
                nPunt = pnId - 1
            End If
            
            ReDim Preserve MatMenuItems(nPunt + 1)
            MatMenuItems(nPunt).nId = pnId - 1
            MatMenuItems(nPunt).sCodigo = Trim(pR!cCodigo)
            MatMenuItems(nPunt).sCaption = Trim(pR!cDescrip)
            MatMenuItems(nPunt).sName = Trim(pR!cname)
            MatMenuItems(nPunt).sIndex = Right(pR!cname, 2)
            MatMenuItems(nPunt).bCheck = False
            MatMenuItems(nPunt).nPuntDer = -1
            MatMenuItems(nPunt).nPuntAbajo = -1
            MatMenuItems(nPuntTemp - 1).nPuntAbajo = MatMenuItems(nPunt).nId
            MatMenuItems(nPunt).nNivel = nPos2
            pR.MoveNext
            
            nPunt = nPunt + 1
            Call CargaMenuArbol(pR, nPunt, pnId)
        End If
        
        If nPos2 < nPos Then 'Es un Menu de nivel Anterior
            nPunt = nPunt - 1
            Call CargaMenuArbol(pR, nPunt, pnId)
        End If
        
End Sub

Public Sub CargaVarSistema(ByVal pbContabilidad As Boolean, _
                            ByVal prsQrySis As ADODB.Recordset, _
                            ByVal psServerName As String, _
                            ByVal psDatabaseName As String, _
                            ByVal psCadenaConexion As String)
    
'    Dim lsQrySis As String
'    Dim rsQrySis As New ADODB.Recordset
'    Dim oConect As COMConecta.DCOMConecta
    Dim VSQL As String
    Dim lnStrConn As String
    Dim lnPosIni As Integer
    Dim lnPosFin As Integer
    Dim lnStr As String
    
'    Set oConect = New COMConecta.DCOMConecta
    
'    If oConect.AbreConexion(gsConnection) = False Then
'        Exit Sub
'    End If
'
'    lsQrySis = " Select * From ConstSistema " _
'            & " Where nConsSisCod in (" & gConstSistFechaSistema & "," & gConstSistNombreAbrevCMAC & "," _
'            & gConstSistRutaBackup & "," & gConstSistCodCMAC & "," & gConstSistMargenSupCartas & "," & gConstSistMagenIzqCartas & "," _
'            & gConstSistMargenDerCartas & "," & gConstSistNroLineasPagina & "," & gConstSistNroLineasOrdenPago & "," _
'            & gConstSistCtaConversionMEDol & "," & gConstSistCtaConversiónMESoles & "," & gConstSistTipoConverión & "," _
'            & gConstSistNombreModulo & "," & gConstSistFechaInicioDia & "," & gConstSistDominio & "," & gConstSistPDC _
'            & ",40,50," & gConstPersCodCMACT & "," & gConstSistAgenciaEspecial & "," & gConstSistCierreMesNegocio & ", " _
'            & gConstSistRutaIcono & ",84,85," & gConstSistVerificaRegistroEfectivo & ") " _
'            & "ORDER BY nConsSisCod"
'
'    Set rsQrySis = oConect.CargaRecordSet(lsQrySis)
    If prsQrySis.BOF Or prsQrySis.EOF Then
       prsQrySis.Close
       Set prsQrySis = Nothing
       MsgBox "Tabla VarSistema está vacia", vbInformation, "Aviso"
       gdFecSis = ""
       gsInstCmac = ""
       gsNomCmac = ""
       gsCodCMAC = ""
       Exit Sub
    End If
    Do While Not prsQrySis.EOF
        Select Case Trim(prsQrySis!nConsSisCod)
                Case gConstSistFechaSistema
                        gdFecSis = CDate(Trim(prsQrySis!nConsSisValor))
                Case gConstSistCierreMesNegocio
                        gdFecDataFM = CDate(Trim(prsQrySis!nConsSisValor))
                        gdFecData = CDate(Trim(prsQrySis!nConsSisValor))
                Case gConstSistNombreAbrevCMAC
                        gsInstCmac = Trim(prsQrySis!nConsSisValor)
                        gsNomCmac = Trim(prsQrySis!nConsSisDesc)
                Case gConstSistCodCMAC
                        gsCodCMAC = Trim(prsQrySis!nConsSisValor)
                Case gConstSistRutaBackup
                        gsDirBackup = Trim(prsQrySis!nConsSisValor)
                Case gConstSistNombreModulo '   "cEmpresa":
                        gcTitModulo = Trim(prsQrySis!nConsSisDesc)
                        gcModuloLogo = Trim(prsQrySis!nConsSisValor)
                Case gConstSistMargenSupCartas  ' "nMargSup":
                        gnMgSup = val(prsQrySis!nConsSisValor)
                Case gConstSistMagenIzqCartas  ' "nMargIzq":
                        gnMgIzq = val(prsQrySis!nConsSisValor)
                Case gConstSistMargenDerCartas  '  "nMargDer":
                        gnMgDer = val(prsQrySis!nConsSisValor)
                Case gConstSistNroLineasPagina  '      "nLinPage":
                        gnLinPage = val(prsQrySis!nConsSisValor)
                Case gConstSistNroLineasOrdenPago  '   "nLinPageOP":
                        gnLinPageOP = val(prsQrySis!nConsSisValor)
                Case gConstSistCtaConversionMEDol '   "cConvMED":
                        gcConvMED = Trim(prsQrySis!nConsSisValor)
                Case gConstSistCtaConversiónMESoles '    "cConvMES":
                        gcConvMES = Trim(prsQrySis!nConsSisValor)
                Case gConstSistTipoConverión  '  "cConvTipo":
                        gcConvTipo = Trim(prsQrySis!nConsSisValor)
                Case gConstSistDominio
                        gsDominio = Trim(prsQrySis!nConsSisValor)
                Case gConstSistPDC
                        gsPDC = Trim(prsQrySis!nConsSisValor)
                Case gConstPersCodCMACT
                    gsCodPersCMACT = Trim(prsQrySis!nConsSisValor)
                Case 40
                    gcCtaIGV = Trim(prsQrySis!nConsSisValor)
                Case 50
                    gcEmpresaRUC = Trim(prsQrySis!nConsSisValor)
                Case gConstSistAgenciaEspecial
                    gbAgeEsp = IIf(Trim(prsQrySis!nConsSisValor) = "1", True, False)
                Case gConstSistRutaIcono
                    gsRutaIcono = Trim(prsQrySis!nConsSisValor)
                Case gConstSistVerificaRegistroEfectivo
                    gbVerificaRegistroEfectivo = IIf(prsQrySis!nConsSisValor = "1", True, False)
                Case 84
                  gnValidaSolCredito = prsQrySis!nConsSisValor
                
                Case 85
                    gnValidaGarantia = prsQrySis!nConsSisValor
                
                Case 301
                    gnDocumentoGarantia = prsQrySis!nConsSisValor
                Case 302
                    gsServPindVerify = prsQrySis!nConsSisValor
                Case 303
                    gsPVKi = prsQrySis!nConsSisValor
                Case 304
                    gsTitutloOP_Soles = prsQrySis!nConsSisValor
                Case 305
                    gsTitutloOP_Dolares = prsQrySis!nConsSisValor
        End Select
        prsQrySis.MoveNext
    Loop
    prsQrySis.Close
    Set prsQrySis = Nothing
    
    'Deduce el nombre del Servidor
    gsServerName = psServerName 'oConect.ServerName
    'Deduce el nombre de la Base de Datos
    gsDBName = psDatabaseName 'oConect.DatabaseName
    lnStrConn = psCadenaConexion 'oConect.CadenaConexion
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
    'oConect.CierraConexion
    'Set oConect = Nothing
    
    gcMN = "S/."
    gcME = "$"
    
    gcMNDig = "1"
    gcMEDig = "2"
End Sub



Public Sub OpeDiaCreaTemporal(psFecha As String, psCodUsu As String)
Dim sql As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

oCon.AbreConexion

sql = " if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[##Mov_" & psCodUsu & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" _
& " drop table [dbo].[##Mov_" & psCodUsu & "]" _
& " "
oCon.Ejecutar sql

sql = " CREATE TABLE [dbo].[##Mov_" & psCodUsu & "] (" _
& " [nMovNro] [int] NOT NULL ," _
& " [cMovNro] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," _
& " [cOpeCod] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," _
& " [cMovDesc] [varchar] (302) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," _
& " [nMovEstado] [int] NOT NULL ," _
& " [nMovFlag] [int] NULL" _
& " ) ON [PRIMARY]" _
& " "
oCon.Ejecutar sql
sql = " ALTER TABLE [dbo].[##Mov_" & psCodUsu & "] WITH NOCHECK ADD" _
& " CONSTRAINT [PK_Mov_" & psCodUsu & "] PRIMARY KEY CLUSTERED" _
& " (" _
& " [nMovNro]" _
& " ) WITH FILLFACTOR = 90 ON [PRIMARY]" _
& " "
oCon.Ejecutar sql

sql = " Insert Into [dbo].[##Mov_" & psCodUsu & "] (nMovNro, cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag)" _
& " Select nMovNro, cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag From Mov Where cMovNro Like '" & psFecha & "%'"
oCon.Ejecutar sql

oCon.CierraConexion

End Sub
Public Sub OpeDiaEliminaTemporal(psCodUsu As String)
Dim sql As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

oCon.AbreConexion

sql = " drop table [dbo].[##Mov_" & psCodUsu & "]"
oCon.Ejecutar sql

oCon.CierraConexion
End Sub

'***Agregado por ELRO el 20120314, según Acta N° 045-2012/TI-D
Public Function validarFechaSistema() As String
Dim sql As String
Dim oDCOMConecta As COMConecta.DCOMConecta
Dim rsFechaSistema As ADODB.Recordset
Set oDCOMConecta = New COMConecta.DCOMConecta
Set rsFechaSistema = New ADODB.Recordset

oDCOMConecta.AbreConexion

sql = "stp_sel_ValidarFechaSistema"
Set rsFechaSistema = oDCOMConecta.CargaRecordSet(sql)

If Not rsFechaSistema.BOF And Not rsFechaSistema.EOF Then
    validarFechaSistema = rsFechaSistema!nConsSisValor
Else
    validarFechaSistema = ""
End If

oDCOMConecta.CierraConexion
Set oDCOMConecta = Nothing
Set rsFechaSistema = Nothing

End Function
'***Fin Agregado por ELRO*************************************

Public Function ValidarFechaSistServer() As String '' ******** Agregado por ANGC 20200306
    Dim fechaSis As Date
    Dim FechaServer As String
    FechaServer = validarFechaSistema()
    fechaSis = gdFecSis
    FechaServer = CDate(FechaServer)
    
    If FechaServer <> fechaSis Then     '' ***** SI LAS FECHAS SON DIFERENTES EL SICMAC DEBE CERRARSE
        ValidarFechaSistServer = "La fecha del sistema y del servidor no coinciden."
    Else
        ValidarFechaSistServer = ""
    End If
End Function '' ******** FIN Agregado por ANGC 20200306

' *** RIRO SEGUN TI-ERS108-2013 ***

Public Function ValidarRFIII() As ADODB.Recordset

Dim oDCOMPersona As COMDPersona.DCOMPersonas
Dim oAcceso As UCOMAcceso
Dim clsGen As COMDConstSistema.DCOMGeneral
 
Dim rsValidar As ADODB.Recordset
Dim rsDatos As ADODB.Recordset

Set rsDatos = New ADODB.Recordset
Set oDCOMPersona = New COMDPersona.DCOMPersonas
Set rsDatos = oDCOMPersona.RecuperarDatosRF3(gsCodAge)

Set rsValidar = New ADODB.Recordset
rsValidar.Fields.Append "cGrupos", adVarChar, 10000
rsValidar.Fields.Append "cGrupoRF3", adVarChar, 100
rsValidar.Fields.Append "bOpcionesSimultaneas", adBoolean, 1
rsValidar.Fields.Append "bModoSupervisor", adBoolean, 1
rsValidar.Fields.Append "bPerfilCambiado", adBoolean, 1
rsValidar.Fields.Append "cUser", adVarChar, adVarChar, 5

If Not rsDatos.BOF And Not rsDatos.EOF Then
    
    Dim sGrupos, sTemporal As String
    Set oAcceso = New UCOMAcceso
    Call oAcceso.CargaGruposUsuario(rsDatos!cUser, gsDominio)
    sTemporal = oAcceso.DameGrupoUsuario
    Do While Len(sTemporal) > 0
        sGrupos = sGrupos & sTemporal & ","
        sTemporal = oAcceso.DameGrupoUsuario
    Loop
    sGrupos = Mid(sGrupos, 1, Len(sGrupos) - 1)

    rsValidar.Open
    rsValidar.AddNew
    
    Set clsGen = New COMDConstSistema.DCOMGeneral
    
    rsValidar.Fields("cGrupos") = sGrupos
    rsValidar.Fields("cGrupoRF3") = clsGen.GetConstante(10027, , "100", "1")!cDescripcion
    rsValidar.Fields("bOpcionesSimultaneas") = rsDatos!nAccionesSimultaneas
    rsValidar.Fields("cUser") = rsDatos!cUser
    
    Dim oPersona As New COMDPersona.DCOMPersonas
    Dim rsRF3 As New ADODB.Recordset
    
    ' *** COMENTADO RIRO20131102
    'If InStr(1, rsValidar!cGrupos, rsValidar!cGrupoRF3) > 0 Then
    '    rsValidar.Fields("bModoSupervisor") = True
    'Else
    '    rsValidar.Fields("bModoSupervisor") = False
    'End If
    
    ' Validando Modo supervisor
    If InStr(1, rsValidar!cGrupos, rsValidar!cGrupoRF3) > 0 Then
        Set rsRF3 = oPersona.RecuperarGruposRF3(rsDatos!cUser)
        If Not rsRF3 Is Nothing Then
            If Not rsRF3.BOF And Not rsRF3.EOF Then
                If rsRF3!nEstado = 1 Then
                    rsValidar.Fields("bModoSupervisor") = True
                Else
                    rsValidar.Fields("bModoSupervisor") = False
                End If
            Else
                rsValidar.Fields("bModoSupervisor") = False
            End If
        Else
            rsValidar.Fields("bModoSupervisor") = False
        End If
        
    Else
        rsValidar.Fields("bModoSupervisor") = False
    End If
    
    
    'Dim bPerfilEnSistema As Boolean
    'If InStr(1, gsGruposUser, rsValidar!cGrupoRF3) > 0 Then
    '    bPerfilEnSistema = True
    'Else
    '    bPerfilEnSistema = False
    'End If
    '
    'If bPerfilEnSistema = rsValidar!bModoSupervisor Then
    '    rsValidar.Fields("bPerfilCambiado") = False
    'Else
    '    rsValidar.Fields("bPerfilCambiado") = True
    'End If
    
    ' Validadndo perfil del sistema
    Dim bPerfilEnSistema As Boolean
    If InStr(1, gsGruposUser, rsValidar!cGrupoRF3) > 0 Then
    
        Set rsRF3 = oPersona.RecuperarGruposRF3(gsCodUser)
        If Not rsRF3 Is Nothing Then
            If Not rsRF3.BOF And Not rsRF3.EOF Then
                If rsRF3!nEstado = 1 Then
                    bPerfilEnSistema = True
                Else
                    bPerfilEnSistema = False
                End If
            Else
                bPerfilEnSistema = False
            End If
        Else
            bPerfilEnSistema = False
        End If
        
    Else
        bPerfilEnSistema = False
    End If
    
    If rsDatos!cUser = gsCodUser Then
        If bPerfilEnSistema = rsValidar!bModoSupervisor Then
            rsValidar.Fields("bPerfilCambiado") = False
        Else
            rsValidar.Fields("bPerfilCambiado") = True
        End If
    Else
        rsValidar.Fields("bPerfilCambiado") = False
    End If
    
Else
    rsValidar.Open
    
End If

Set ValidarRFIII = rsValidar

End Function

' *** FIN RIRO ***
'****************RECO20121115*************************************
'TORE : Comentar para compilar los formularios excluidos
Public Function CargaImpuestoFechaValor(Optional psCtaContCod As String = "", Optional pdFecha As Date = 0) As Currency
   On Error GoTo CargaImpuestoErr
   Dim sql As String
   Dim oDCOMConecta As COMConecta.DCOMConecta
   Dim pRs As ADODB.Recordset
   Set oDCOMConecta = New COMConecta.DCOMConecta
   Set pRs = New ADODB.Recordset

   oDCOMConecta.AbreConexion

   sql = "SELECT i.cCtaContCod, nImpTasa FROM ImpuestoFecha i WHERE i.cCtaContCod = '" & psCtaContCod & "'" _
          & IIf(pdFecha = 0, "", " and dFechaIniVig = (SELECT Max(dFechaIniVig) FROM ImpuestoFecha f WHERE f.cCtaContCod = i.cCtaContCod and f.dFechaIniVig <= '" & Format(pdFecha, gsFormatoFecha) & "' ) ")

    Set pRs = oDCOMConecta.CargaRecordSet(sql)


      If Not pRs.EOF Then
         CargaImpuestoFechaValor = pRs!nImpTasa
      End If
      RSClose pRs
      oDCOMConecta.CierraConexion
      Set oDCOMConecta = Nothing
      Set pRs = Nothing


   Exit Function
CargaImpuestoErr:
   Call RaiseError(MyUnhandledError, "gVarPublicas:CargaImpuestoFecha Method")
End Function
