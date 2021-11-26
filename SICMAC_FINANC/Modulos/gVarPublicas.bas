Attribute VB_Name = "gVarPublicas"
Public oImpresora As New clsConstImp
Public gImpresora As Impresoras
Global Const gnNroDigitosDNI = 8
Global Const gnNroDigitosRUC = 11
Global Const gMonedaVAC = 3

'*******************************************************************
'************************ Constantes Públicas *******************
'*******************************************************************
Global Const gsMenuAplicac = "3" 'para identificar los menus que le corresponden
Global Const gsFormatoFecha = "mm/dd/yyyy"
Global Const gsFormatoFechaHora = "mm/dd/yyyy hh:mm:ss"
Global Const gsFormatoFechaHoraView = "dd/mm/yyyy hh:mm:ss"
Global Const gsFormatoFechaView = "dd/mm/yyyy"

Global Const gsFormatoMovFecha = "yyyymmdd"
Global Const gsFormatoMovFechaHora = "yyyymmddhhmmss"

'Global Const gsFormatoNumeroView = "##,###,##0.00##"
Global Const gsFormatoNumeroView = "##,###,##0.00"
Global Const gsFormatoNumeroView3Dec = "##,###,##0.000"

Global Const gsFormatoNumeroDato = "#######0.00##"
Global Const gsFormatoNumeroDato2D = "#######0.00"

Global Const gcFormatoTC = "#0.00##"
Global Const gnColPage = 79   'Columnas por página de Impresión
Global Const gnLinVert = 66   'Orientación Vertical
Global Const gnLinHori = 46   'Orientación Horizontal
Global Const gbComunRemoto = False
Global Const IDPlantillaOP = "OPBatch"
Global Const IDPlantillaVOP = "OPVEBatch"

'''Global Const gcMN = "S/." 'MARG ERS044-2016
Global Const gcMN = "S/" 'MARG ERS044-2016
Global Const gcME = "US$"

Global Const gsColorME = "&H00008000"
Global Const gsColorMN = "&H00FF0000"
Global Const gsBackColorME = "&H00C0FFC0"
Global Const gsBackColorMN = "&H80000009"

Global Const gsACInaAct = "201001"
Global Const gsACDepNARH = "201701"
Global Const gsACRetNCRH = "201801"

'*******************************************************************
'************************ Variables Globales *******************
'*******************************************************************
Global gsDominio As String
Global gsPDC As String
Global gsCtaApe As String
Global gsDirBackup As String
Global gsServerName As String
Global gsUser       As String
Global gsCodUser    As String
Global gsCodAge     As String
Global gsCodAgeAsig As String
Global gsCodArea As String
Global gsCodCargo As String 'EJVG20111217
Global gsNomAge As String * 30
Global gsNomCmac As String
Global gsCodCMAC As String
Global gbOpeOk As Boolean
Global gsRutaFirmas As String
Global gsDBName As String
Global gsUID As String
Global gsPWD As String
Global gsCodPersUser As String 'Add By GITU 26-01-2009
Global gsGrupoUsu As String 'Juez 20120715

Global gnTipCambio As Currency
Global gnTipCambioV As Currency
Global gnTipCambioC As Currency
Global gnTipCambioVE As Currency
Global gnTipCambioCE As Currency
Global gnTipCambioPonderado As Currency
Global gnTipCambioPonderadoVenta As Currency
Global gsRutaIcono As String

Global gsProyectoActual As String

Global gsMesCerrado As String

'Variable de la Fecha del sitema
Global gdFecSis As Date
'Variables de Conection
Global gsCodAgeN As String

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
Global gsNomCmacRUC As String   'RUC de la Entidad Financiera
Global gcModuloLogo  As String

'Variables de Contabilidad
Global gsMovNro As String
Global gnMovNro As Long
Global gsGlosa  As String
Global gnImporte As Currency
Global gnSaldo   As Currency
Global gdFecha  As Date

Global gsOpeCod As String
Global gsOpeDesc As String
Global gsOpeDescPadre As String
Global gsOpeDescHijo As String

Global gsSimbolo As String

Global gcTitModulo    As String
Global glDiaCerrado As Boolean
Global gcCtaIGV   As String
Global gnMgSup   As Integer
Global gnMgIzq  As Integer
Global gnMgDer  As Integer
Global gnLinPage   As Integer
Global gnArendirImporte  As Currency
Global gnLinPageOP   As Integer
Global gcConvMED  As String
Global gcConvMES  As String

Global gnIGVValor As Currency

Global gcConvMEDAjTC  As String
Global gcConvMESAjTC  As String

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
Global gsPersNombre    As String
Global cCtaDetraccionProvision As String

'************BANCO PAGADOR*************' AMDO20140112

Global Const gAhoITFCargoCta = "990101" '990121
Global Const gAhoBancPagIni = "200265"
'**************************************

'***********Modulo de Contingencias ***********'** Juez 20120614
Global Const gActivoContingente = "1"
Global Const gPasivoContingente = "2"

Global Const gOrigenEventoPerdidas = "2" 'Eventos de Pérdidas

'CodOpe
Global Const gRegistroActivoContingente = "500501"
Global Const gRegistroPasivoContingente = "500502"
Global Const gRegistroInfTec = "500510"
Global Const gLiberarContigencia = "500520"
Global Const gDesestimarConting = "500530"
'*************************************************

'**********Modulo de Intangibles *************'PASI 20140318
'Codigos de Operacion
Global Const gRegistroActivacionIntangible = "500401" 'Activacion de Intangibles
Global Const gAmortizaIntangibleLicencia = "501402"
Global Const gAmortizaIntangibleSoftware = "501403"
Global Const gAmortizaIntangibleOtros = "501404"

Global Const gExtornoAmortizacion = "501410"
Global Const gBajaIntangible = "501499"
'*********************************************

'Documentos
Public gnDocTpo  As Long
Public gsDocDesc As String
Public gsDocNro  As String

Global gsRUC As String
Global gsEmpresa As String
Global gsEmpresaCompleto As String
Global gsEmpresaDireccion As String
Global gcPDC As String
Global gcDominio As String
Global gcWINNT As String
Global gbBitCentral As Boolean
Global gnAplicaITF As Integer
'Global gnImpITF As Currency
Global gnImpITF As Double '*** PEAC 20110331
Global gbBitRetencSistPensProv As Boolean 'EJVG20140724

Global gbBitTCPonderado As Boolean
'Global gnDocCuentaPendiente As Integer
Global Const gnDocCuentaPendiente = 80
Global gsFechaVersion As String 'ARLO20170626

Global gnTipoCambioEuro As Currency

Public sLPT As String

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
'Public MatOperac(1000, 5) As String
Public MatOperac(6000, 5) As String 'Antes 2000 NAGL 20191230
Public NroRegOpe As Integer

'Pendientes
Global Const gOpePendOpeAgencias = "760401"
Global Const gOpePendFaltantCaja = "760402"
Global Const gOpePendRendirCuent = "760403"
Global Const gOpePendDisponibRes = "760404"
Global Const gOpePendCtaCobrarDi = "760405"
Global Const gOpePendPagoSubsidi = "760406"
Global Const gOpePendCtaCobraDiv = "760407"

Global Const gOpePendOrdendePago = "760420"
Global Const gOpePendCobraLiquid = "760421"
Global Const gOpePendSobraRemate = "760422"
Global Const gOpePendOtrasProvis = "760423"
Global Const gOpePendSobrantCaja = "760424"
Global Const gOpePendOtrasOpeLiqPas = "760425"
Global Const gOpePendRecursHuman = "760426"
Global Const gOpePendOPCertifica = "760428"
Global Const gOpePendProvPagoProv = "760429"
Global Const gOpePendMntHistoric = "701230"

'RRHH
Global Const gsRHPlanillaSueldos = "E01"
Global Const gsRHPlanillaGratificacion = "E02"
Global Const gsRHPlanillaTercio = "E03"
Global Const gsRHPlanillaUtilidades = "E04"
Global Const gsRHPlanillaCTS = "E05"
Global Const gsRHPlanillaVacaciones = "E06"
Global Const gsRHPlanillaSubsidio = "E07"
Global Const gsRHPlanillaLiquidacion = "E08"
Global Const gsRHPlanillaBonificacionVacacinal = "E09"
Global Const gsRHPlanillaBonoProductividad = "E10"
Global Const gsRHPlanillaBonoAguinaldo = "E11"
Global Const gsRHPlanillaSubsidioEnfermedad = "E12"
Global Const gsRHPlanillaReintegro = "E13"
Global Const gsRHPlanillaDev5ta = "E14"
Global Const gsRHPlanillaMovilidad = "E15"

Global Const gnRHTotalTpo = 9999
Global Const gsRHConceptoUMESTRAB = "U_POR_MES_TRAB"
Global Const gsRHConceptoITOTING = "I_TOT_ING"
Global Const gsRHConceptoITOTINGCOD = "130"
Global Const gsRHConceptoINETOPAGARCOD = "112"
Global Const gsRHConceptoINETOPAGAR = "I_NETO_PAGAR"
Global Const gsRHConceptoDTOTDES = "D_TOT_DESC"
Global Const gsRHConceptoDTOTDESCOD = "215"

Global Const gsRHConceptoVTOTREM = "V_TOT_REM"
Global Const gsRHConceptoITOTCTS = "I_TOT_CTS"
Global Const gsRHConceptoITOTTERCIO = "I_TERCIO"
Global Const gsRHConceptoITOTGRAT = "I_TOTAL_GRATIF"
Global Const gsRHConceptoVNETOPAGAR = "V_NETO_PAGAR"


'Codigos de operacion para el calculo de provisiones y remuneraciones
'Sueldos
Global Const gsRHPlanillaSueldosRemEst = "622001"
Global Const gsRHPlanillaSueldosRemCon = "622002"
'Gratificacion
Global Const gsRHPlanillaGratificacionProvEst = "622101"
Global Const gsRHPlanillaGratificacionProvCon = "622102"
Global Const gsRHPlanillaGratificacionRemEst = "622103"
Global Const gsRHPlanillaGratificacionRemCon = "622104"
'Tercio
Global Const gsRHPlanillaTercioProvEst = "622201"
Global Const gsRHPlanillaTercioProvCon = "622202"
Global Const gsRHPlanillaTercioRemEst = "622203"
Global Const gsRHPlanillaTercioRemCon = "622204"
'Utilidades
Global Const gsRHPlanillaUtilidadesProvEst = "622301"
Global Const gsRHPlanillaUtilidadesProvCon = "622302"
Global Const gsRHPlanillaUtilidadesRemEst = "622303"
Global Const gsRHPlanillaUtilidadesRemCon = "622304"
'CTS
Global Const gsRHPlanillaCTSProvEst = "622401"
Global Const gsRHPlanillaCTSProvCon = "622402"
Global Const gsRHPlanillaCTSRemEst = "622403"
Global Const gsRHPlanillaCTSRemCon = "622404"
'Vacaciones
Global Const gsRHPlanillaVacacionesProvEst = "622501"
Global Const gsRHPlanillaVacacionesProvCon = "622502"
Global Const gsRHPlanillaVacacionesRemEst = "622503"
Global Const gsRHPlanillaVacacionesRemCon = "622504"
'Subsidios
Global Const gsRHPlanillaSubsidioRem = "622601"
Global Const gsRHPlanillaSubsidioEnfermedadRem = "622602"
'Liquidacion
Global Const gsRHPlanillaLiquidacionRem = "622701"
'Bonificacion Vacacional
Global Const gsRHPlanillaBonificacionVacacinalProvEst = "622801"
Global Const gsRHPlanillaBonificacionVacacinalProvCon = "622802"
Global Const gsRHPlanillaBonificacionVacacinalRemEst = "622803"
Global Const gsRHPlanillaBonificacionVacacinalRemCon = "622804"
'Bono Productividad
Global Const gsRHPlanillaBonoProductividadProvEst = "622901"
Global Const gsRHPlanillaBonoProductividadProvCon = "622902"
Global Const gsRHPlanillaBonoProductividadRemEst = "622903"
Global Const gsRHPlanillaBonoProductividadRemCon = "622904"

'Bono Reintegro
Global Const gsRHPlanillaReintegroProvEst = "623001"
Global Const gsRHPlanillaReintegroProvCon = "623002"
Global Const gsRHPlanillaReintegroRemEst = "623003"
Global Const gsRHPlanillaReintegroRemCon = "623004"

'Bono Rev 5ta Categoria
Global Const gsRHPlanillaDev5taRemEst = "623101"
Global Const gsRHPlanillaDev5taRemCon = "623102"

'Bono Movilidad
Global Const gsRHPlanillaMovilidadProvEst = "623201"
Global Const gsRHPlanillaMovilidadProvCon = "623202"
Global Const gsRHPlanillaMovilidadRemEst = "623203"

'*******VARIABLES DE MONEDA****************
'MARG ERS044-2016
'''Global gcPEN_SINGULAR As String
'''Global gcPEN_PLURAL As String
'''Global gcPEN_SIMBOLO As String
Global Const gcPEN_SINGULAR = "SOL"
Global Const gcPEN_PLURAL = "SOLES"
Global Const gcPEN_SIMBOLO = "S/"
'******************************************

Private Function DamePosicionNivel(ByVal psName As String) As Integer
Dim I As Integer
Dim y As Integer
    y = 1
    For I = 4 To Len(psName) Step 2
        If Mid(psName, I, 2) = "00" Then
            DamePosicionNivel = y
            Exit For
        End If
        y = y + 1
    Next I
    If y = 6 Then
        DamePosicionNivel = 5
    End If
    
End Function

Public Sub CargaMenuArbol(ByRef pR As ADODB.Recordset, ByRef nPunt As Integer, ByRef pnId As Integer)
Dim nPos As Integer
Dim nPos2 As Integer
Dim nPuntTemp As Integer
        'If Left(pR!cCodigo, 10) = "1601060000" Then
        '    nPunt = nPunt
        'End If

        If nPunt = 175 Then
            nPunt = nPunt
        End If

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
            
            If MatMenuItems(nPunt - 1).nPuntDer = 0 Then
                nPos = nPos
            End If
            
            
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
            
            If MatMenuItems(nPuntTemp - 1).nPuntDer = 0 Then
                nPos = nPos
            End If
            
            nPunt = nPunt + 1
            Call CargaMenuArbol(pR, nPunt, pnId)
        End If
        
        If nPos2 < nPos Then 'Es un Menu de nivel Anterior
            nPunt = nPunt - 1
            Call CargaMenuArbol(pR, nPunt, pnId)
        End If
        
End Sub

'***AGREGADO POR ANGC 20200306 COPIADO DE  SICMAN NEG  ELRO el 20120314, según Acta N° 045-2012/TI-D
Public Function validarFechaSistema() As String
    Dim sql As String
    Dim oConect As DConecta
    Dim rsFechaSistema As ADODB.Recordset
    Set oConect = New DConecta
    Set rsFechaSistema = New ADODB.Recordset
    
    If oConect.AbreConexion(gsConnection) = False Then
        Exit Function
    End If
    
    sql = "stp_sel_ValidarFechaSistema"
    Set rsFechaSistema = oConect.CargaRecordSet(sql)
    
    If Not rsFechaSistema.BOF And Not rsFechaSistema.EOF Then
        validarFechaSistema = rsFechaSistema!nConsSisValor
    Else
        validarFechaSistema = ""
    End If
    
    oConect.CierraConexion
    Set oConect = Nothing
    Set rsFechaSistema = Nothing
End Function
'***Fin POR ANGC 20200306*************************************

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


Public Sub CargaVarSistema(ByVal pbContabilidad As Boolean)
    Dim lsQrySis As String
    Dim rsQrySis As New ADODB.Recordset
    Dim oConect As DConecta
    Dim VSQL As String
    Dim lnStrConn As String
    Dim lnPosIni As Integer
    Dim lnPosFin As Integer
    Dim lnStr As String
    Dim pdFecIni As Date
    
    Set oConect = New DConecta
    If oConect.AbreConexion(gsConnection) = False Then
        Exit Sub
    End If
    
    lsQrySis = " Select * From ConstSistema " _
            & " Where nConsSisCod in (" & gConstSistFechaSistema & "," & gConstSistNombreAbrevCMAC & "," _
            & gConstSistRutaBackup & "," & gConstSistCodCMAC & "," & gConstSistMargenSupCartas & "," & gConstSistMagenIzqCartas & "," _
            & gConstSistMargenDerCartas & "," & gConstSistNroLineasPagina & "," & gConstSistNroLineasOrdenPago & "," _
            & gConstSistCtaConversionMEDol & "," & gConstSistCtaConversiónMESoles & "," & gConstSistTipoConverión & "," _
            & gConstSistNombreModulo & "," & gConstSistFechaInicioDia & "," & gConstSistDominio & "," & gConstSistPDC & "," & gConstSistCMACRuc & ",40," & gConstSistBitCentral & "," & gConstSistRutaIcono & ", 111, 112, 155, 156,481 ) ORDER BY nConsSisCod"
    
    Set rsQrySis = oConect.CargaRecordSet(lsQrySis)
    If rsQrySis.BOF Or rsQrySis.EOF Then
       rsQrySis.Close
       Set rsQrySis = Nothing
       MsgBox "Tabla VarSistema está vacia", vbInformation, "Aviso"
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
                Case gConstSistFechaInicioDia
                        pdFecIni = CDate(Trim(rsQrySis!nConsSisValor))
                Case gConstSistNombreAbrevCMAC
                        gsInstCmac = Trim(rsQrySis!nConsSisValor)
                        gsNomCmac = Trim(rsQrySis!nConsSisDesc)
                Case gConstSistCMACRuc
                        gsNomCmacRUC = Trim(rsQrySis!nConsSisValor)
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
                Case gConstSistBitCentral
                    gbBitCentral = (rsQrySis!nConsSisValor = "1")
                Case gConstSistRutaIcono
                    gsRutaIcono = Trim(rsQrySis!nConsSisValor)
                Case 111   'gConstSistCtaConversionMEDolAjTC
                    gcConvMEDAjTC = Trim(rsQrySis!nConsSisValor)
                Case 112   'gConstSistCtaConversionMESolAjTC
                    gcConvMESAjTC = Trim(rsQrySis!nConsSisValor)
                Case 155
                    gnAplicaITF = rsQrySis!nConsSisValor
                Case 156
                    gnImpITF = rsQrySis!nConsSisValor
                Case 481 'EJVG20140724
                    gbBitRetencSistPensProv = IIf(rsQrySis!nConsSisValor = "1", True, False)
            End Select
        rsQrySis.MoveNext
    Loop
    rsQrySis.Close
    Set rsQrySis = Nothing
    
    Dim oConst As New NConstSistemas
    
    If pdFecIni > gdFecSis Then
        If MsgBox(" ¿ Desea realizar Inicio de Dia " & pdFecIni & " ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbYes Then
            gsMovNro = GeneraMovNroActualiza(CDate(Format(Now, gsFormatoFechaView)), gsCodUser, gsCodCMAC, gsCodAge)
            oConst.ActualizaConstSistemas gConstSistFechaSistema, gsMovNro, pdFecIni
            gdFecSis = pdFecIni
            glDiaCerrado = False
        Else
            glDiaCerrado = True
        End If
    End If

    If Not gbBitCentral Then
        ' comentado por cambio de funcion
        'If oConect.AbreConexionRemota(gsCodAge, True, False, "01") Then
        If oConect.AbreConexion() Then
            lsQrySis = " SELECT rtrim(cValorVar) as nConsSisValor FROM VarSistema WHERE cNomVar IN ('dFecSis')"
            Set rsQrySis = oConect.CargaRecordSet(lsQrySis)
            If Not rsQrySis.EOF Then
                If CDate(Trim(rsQrySis!nConsSisValor)) > gdFecSis Then
                    gdFecSis = ldFecSisNeg
                    
                    gsMovNro = GeneraMovNroActualiza(CDate(Format(Now, gsFormatoFechaView)), gsCodUser, gsCodCMAC, gsCodAge)
                    oConst.ActualizaConstSistemas gConstSistFechaSistema, gsMovNro, CDate(Trim(rsQrySis!nConsSisValor))
                    oConst.ActualizaConstSistemas gConstSistFechaInicioDia, gsMovNro, CDate(Trim(rsQrySis!nConsSisValor))
                End If
            End If
        End If
    End If
    oConect.CierraConexion
    oConect.AbreConexion
    'Deduce el nombre del Servidor
    gsServerName = oConect.ServerName
    'Deduce el nombre de la Base de Datos
    gsDBName = oConect.DatabaseName
    lnStrConn = oConect.CadenaConexion
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
        '----------MARG ERS044-2016---------
    '''gcPEN_SINGULAR = "SOL"
    '''gcPEN_PLURAL = "SOLES"
    '''gcPEN_SIMBOLO = "S/"
    '-----------------------------------
    oConect.CierraConexion
    Set oConect = Nothing
End Sub




