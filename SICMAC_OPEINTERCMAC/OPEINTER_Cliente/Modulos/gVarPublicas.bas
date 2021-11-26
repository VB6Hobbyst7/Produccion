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
Global gsCodCargo As String
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

'** DAOR 20081125 *******************************
Global gnPinPadPuerto As Integer
Global gsNomMaquinaUsu As String
Global Const gIpPuertoPinVerifyPOS = "192.168.0.9:81"
Global Const gWKPOS = "81AE036D7855A288"
Global Const gNMKPOS = 0
Global Const gCanalIdPOS = "_02"
'************************************************

'____________________________________________________________
'
'*********VARIABLES PARA PERSONA DE LAVADO DE DINERO*********
'____________________________________________________________

Global gReaPersLavDinero As String
Global gBenPersLavDinero As String
Global gOrdPersLavDinero As String 'By Capi 20012008
Global gVisPersLavDinero As String 'DAOR 20070511
Global gnTipoREU As Integer 'ALPA 20081003
Global gnMontoAcumulado As Double 'ALPA 20081003
Global gsOrigen As String 'ALPA 20081003

'***************************




'Documentos
Public gnDocTpo  As Long
Public gsDocDesc As String
Public gsDocNro  As String

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

'ARCV 20-07-2006
Public gRsOpeRepo As ADODB.Recordset
'---------------

'Para el manejo de los Proyectos en las diferentes Empresas
Public gsProyectoActual As String


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

Public Sub CargaVarSistema()
    
    Dim sSQL As String
    Dim prsQrySis As New ADODB.Recordset
    'Dim oConect As DConecta
    Dim VSQL As String
    Dim lnStrConn As String
    Dim lnPosIni As Integer
    Dim lnPosFin As Integer
    Dim lnStr As String
    
    Dim clsFun As DFunciones.dFuncionesNeg
    Set clsFun = New DFunciones.dFuncionesNeg
    
'    sSQL = " Select nConsSisCod,nConsSisValor,nConsSisDesc From DBCMAC..ConstSistema " _
'            & " Where nConsSisCod in (" & gConstSistFechaSistema & "," & gConstSistNombreAbrevCMAC & "," _
'            & gConstSistRutaBackup & "," & gConstSistCodCMAC & "," & gConstSistMargenSupCartas & "," & gConstSistMagenIzqCartas & "," _
'            & gConstSistMargenDerCartas & "," & gConstSistNroLineasPagina & "," & gConstSistNroLineasOrdenPago & "," _
'            & gConstSistCtaConversionMEDol & "," & gConstSistCtaConversiónMESoles & "," & gConstSistTipoConverión & "," _
'            & gConstSistNombreModulo & "," & gConstSistFechaInicioDia & "," & gConstSistDominio & "," & gConstSistPDC _
'            & ",40,50," & gConstPersCodCMACT & "," & gConstSistAgenciaEspecial & "," & gConstSistCierreMesNegocio & ", " _
'            & gConstSistRutaIcono & ",84,85," & gConstSistVerificaRegistroEfectivo & ",301,302,303,304,305) " _
'            & "ORDER BY nConsSisCod"


    Set prsQrySis = clsFun.GetVarSistema
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
                        gdFecSis = CDate(Trim(prsQrySis!nConssisValor))
                Case gConstSistCierreMesNegocio
                        gdFecDataFM = CDate(Trim(prsQrySis!nConssisValor))
                        gdFecData = CDate(Trim(prsQrySis!nConssisValor))
                Case gConstSistNombreAbrevCMAC
                        gsInstCmac = Trim(prsQrySis!nConssisValor)
                        gsNomCmac = Trim(prsQrySis!nConsSisDesc)
                Case gConstSistCodCMAC
                        gsCodCMAC = Trim(prsQrySis!nConssisValor)
                Case gConstSistRutaBackup
                        gsDirBackup = Trim(prsQrySis!nConssisValor)
                Case gConstSistNombreModulo '   "cEmpresa":
                        gcTitModulo = Trim(prsQrySis!nConsSisDesc)
                        gcModuloLogo = Trim(prsQrySis!nConssisValor)
                Case gConstSistMargenSupCartas  ' "nMargSup":
                        gnMgSup = Val(prsQrySis!nConssisValor)
                Case gConstSistMagenIzqCartas  ' "nMargIzq":
                        gnMgIzq = Val(prsQrySis!nConssisValor)
                Case gConstSistMargenDerCartas  '  "nMargDer":
                        gnMgDer = Val(prsQrySis!nConssisValor)
                Case gConstSistNroLineasPagina  '      "nLinPage":
                        gnLinPage = Val(prsQrySis!nConssisValor)
                Case gConstSistNroLineasOrdenPago  '   "nLinPageOP":
                        gnLinPageOP = Val(prsQrySis!nConssisValor)
                Case gConstSistCtaConversionMEDol '   "cConvMED":
                        gcConvMED = Trim(prsQrySis!nConssisValor)
                Case gConstSistCtaConversiónMESoles '    "cConvMES":
                        gcConvMES = Trim(prsQrySis!nConssisValor)
                Case gConstSistTipoConverión  '  "cConvTipo":
                        gcConvTipo = Trim(prsQrySis!nConssisValor)
                Case gConstSistDominio
                        gsDominio = Trim(prsQrySis!nConssisValor)
                Case gConstSistPDC
                        gsPDC = Trim(prsQrySis!nConssisValor)
                Case gConstPersCodCMACT
                    gsCodPersCMACT = Trim(prsQrySis!nConssisValor)
                Case 40
                    gcCtaIGV = Trim(prsQrySis!nConssisValor)
                Case 50
                    gcEmpresaRUC = Trim(prsQrySis!nConssisValor)
                Case gConstSistAgenciaEspecial
                    gbAgeEsp = IIf(Trim(prsQrySis!nConssisValor) = "1", True, False)
                Case gConstSistRutaIcono
                    gsRutaIcono = Trim(prsQrySis!nConssisValor)
                Case gConstSistVerificaRegistroEfectivo
                    gbVerificaRegistroEfectivo = IIf(prsQrySis!nConssisValor = "1", True, False)
                Case 84
                  gnValidaSolCredito = prsQrySis!nConssisValor
                
                Case 85
                    gnValidaGarantia = prsQrySis!nConssisValor
                
                Case 301
                    gnDocumentoGarantia = prsQrySis!nConssisValor
                Case 302
                    gsServPindVerify = prsQrySis!nConssisValor
                Case 303
                    gsPVKi = prsQrySis!nConssisValor
                Case 304
                    gsTitutloOP_Soles = prsQrySis!nConssisValor
                Case 305
                    gsTitutloOP_Dolares = prsQrySis!nConssisValor
                    
        End Select
        prsQrySis.MoveNext
    Loop
    prsQrySis.Close
    Set prsQrySis = Nothing
    

    
    gcMN = "S/."
    gcME = "$"
    
    gcMNDig = "1"
    gcMEDig = "2"
End Sub

'Descomentar GITU

'Public Sub OpeDiaCreaTemporal(psFecha As String, psCodUsu As String)
'Dim Sql As String
'Dim oCon As DConecta
'Set oCon = New DConecta
'
'oCon.AbreConexion
'
'Sql = " if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[##Mov_" & psCodUsu & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" _
'& " drop table [dbo].[##Mov_" & psCodUsu & "]" _
'& " "
'oCon.Ejecutar Sql
'
'Sql = " CREATE TABLE [dbo].[##Mov_" & psCodUsu & "] (" _
'& " [nMovNro] [int] NOT NULL ," _
'& " [cMovNro] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," _
'& " [cOpeCod] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," _
'& " [cMovDesc] [varchar] (302) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," _
'& " [nMovEstado] [int] NOT NULL ," _
'& " [nMovFlag] [int] NULL" _
'& " ) ON [PRIMARY]" _
'& " "
'oCon.Ejecutar Sql
'Sql = " ALTER TABLE [dbo].[##Mov_" & psCodUsu & "] WITH NOCHECK ADD" _
'& " CONSTRAINT [PK_Mov_" & psCodUsu & "] PRIMARY KEY CLUSTERED" _
'& " (" _
'& " [nMovNro]" _
'& " ) WITH FILLFACTOR = 90 ON [PRIMARY]" _
'& " "
'oCon.Ejecutar Sql
'
'Sql = " Insert Into [dbo].[##Mov_" & psCodUsu & "] (nMovNro, cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag)" _
'& " Select nMovNro, cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag From Mov Where cMovNro Like '" & psFecha & "%'"
'oCon.Ejecutar Sql
'
'oCon.CierraConexion
'
'End Sub
'Public Sub OpeDiaEliminaTemporal(psCodUsu As String)
'Dim Sql As String
'Dim oCon As DConecta
'Set oCon = New DConecta
'
'oCon.AbreConexion
'
'Sql = " drop table [dbo].[##Mov_" & psCodUsu & "]"
'oCon.Ejecutar Sql
'
'oCon.CierraConexion
'End Sub
