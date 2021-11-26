VERSION 5.00
Begin VB.Form frmCFHojaAprob 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hoja de Aprobación de Carta Fianza"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   Icon            =   "frmCFHojaAprob.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   310
      Left            =   2400
      TabIndex        =   11
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   310
      Left            =   1320
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   310
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del Crédito "
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6735
      Begin VB.Label lblFecDesembolso 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4800
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblEstado 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Desemb.:"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   750
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Estado:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   750
         Width           =   735
      End
      Begin VB.Label lblTitular 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   465
         Width           =   6255
      End
      Begin VB.Label Label1 
         Caption         =   "Titular:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin SICMACT.ActXCodCta_New ActXCodCta 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Texto           =   "Crédito"
      EnabledCta      =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "frmCFHojaAprob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************************************************************
'* NOMBRE         : "frmCFHojaAprob"
'* DESCRIPCION    : Permite generar la hoja de aprobacion de una carta fianza.
'* CREACION       : RECO, 09/03/2016 12:00 PM
'************************************************************************************************************************************************

Option Explicit

Dim sDNI As String
Dim sRUC As String
Dim sPersTDoc As String

Public Sub Inicio(ByVal psCtaCod As String)
    Call CargaDatos(psCtaCod)
    Call ImprimeAprobacionCreditos(psCtaCod, sDNI, sRUC)
End Sub

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call CargaDatos(ActXCodCta.NroCuenta)
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiarFormulario
End Sub

Private Sub cmdImprimir_Click()
    If Len(ActXCodCta.NroCuenta) = 18 Then
        Call ImprimeAprobacionCreditos(ActXCodCta.NroCuenta, sDNI, sRUC)
    Else
        MsgBox "Ingrese un número de Carta Fianza válido", vbInformation, "Alerta"
        Call LimpiarFormulario
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
    Dim oCF As New COMDCartaFianza.DCOMCartaFianza
    Dim oDatos As New ADODB.Recordset
    
    Set oDatos = oCF.RecuperaCartaFianzaDetalle(psCtaCod)
    
    If Not (oDatos.EOF And oDatos.BOF) Then
        lblTitular = oDatos!cPersNombre
        lblEstado = oDatos!cEstado
        lblMonto = Format(oDatos!nMontoApr, gsFormatoNumeroView)
        lblFecDesembolso = oDatos!dFecDesemb
        sDNI = oDatos!cDNI
        sRUC = oDatos!cRUC
        
        If oDatos!nPersPersoneria = 1 Then
            sPersTDoc = 1
        Else
            sPersTDoc = 3
        End If
    Else
        MsgBox "No se encontraron datos del crédito", vbInformation, "Alerta SICMAC"
        Call LimpiarFormulario
    End If
End Function

Public Sub ImprimeAprobacionCreditos(ByVal psCtaCod As String, ByVal psCodNat As String, ByVal psCodJur As String)
    Dim oGeneral As New COMDConstSistema.DCOMGeneral
    Dim oDCredExoAut As New COMDCredito.DCOMNivelAprobacion
    Dim oDLeasing As New COMDCredito.DCOMleasing
    Dim oDPersGeneral As New comdpersona.DCOMPersGeneral
    Dim oDCred As New COMDCredito.DCOMCredito
    Dim oDBCred As New COMDCredito.DCOMCredDoc
    Dim oConst As New COMDConstSistema.DCOMGeneral
    Dim oDGarantia As New COMDCredito.DCOMGarantia
    Dim oCF As New COMDCartaFianza.DCOMCartaFianza
    Dim R As ADODB.Recordset
    Dim RCaliSbs As ADODB.Recordset, RRelaCred As ADODB.Recordset, RGarantCred As ADODB.Recordset
    Dim rBancos As ADODB.Recordset, RDatFin As ADODB.Recordset, rResGarTitAva As ADODB.Recordset
    Dim REstCivConvenio As ADODB.Recordset, RCredEval As ADODB.Recordset, RLeasing As ADODB.Recordset
    Dim RCredAmp As ADODB.Recordset, RCalfSbsRel As ADODB.Recordset, RExoAutCred As ADODB.Recordset
    Dim RRiesgoUnico As ADODB.Recordset, RCredGarant As ADODB.Recordset, RCredResulNivApr As ADODB.Recordset
    Dim RNivApr As ADODB.Recordset, ROpRiesgo As ADODB.Recordset, RComentAnalis As ADODB.Recordset
    Dim sConstante As String
    Dim sRiesgo As String
    Dim nTipoCambioFijo
    
    Set R = oCF.ObtieneDatosAprobacionCF(psCtaCod)
    Set RCaliSbs = oDCred.RecuperaCaliSbs(psCodNat, psCodJur, sPersTDoc)
    Set RRelaCred = oDCred.RecuperaRelacPers(psCtaCod)
    Set RDatFin = oDCred.RecuperaDatosFinan(psCtaCod)
    Set RCredEval = oDCred.RecuperaColocacCredEvalAprobacion(psCtaCod) ' ESTO
    Set ROpRiesgo = oDCred.RecuperaOpinionRiesgo(psCtaCod) ' Esto
    Set RCredAmp = oDCred.VerificarAmpliados(psCtaCod)
    Set RCalfSbsRel = oDCred.ObtieneCalifSBSRelacionCred(psCtaCod)
    Set RExoAutCred = oDCredExoAut.ObtieneExoneraAutoriCred(psCtaCod)
    Set RRiesgoUnico = oDCredExoAut.ObtineRiesgoUnicoCred(psCtaCod)
    Set RCredGarant = oDCred.ObtieneDatasGarantiaCred(psCtaCod)
    Set RCredResulNivApr = oDCredExoAut.ObtieneNivelAprResultado(psCtaCod)
    Set RNivApr = oDCredExoAut.RecuperaHistorialCredAprobados(psCtaCod)
    Set RComentAnalis = oDCred.RecuperaComentarioAnalistaSugerencia(psCtaCod)
    Set RLeasing = oDLeasing.Obtener_MontoFinanciarLeasing(psCtaCod)
    
    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Exit Sub
    End If
    
    Set RGarantCred = oDGarantia.RecuperaGarantiaCredito(psCtaCod)
    Set rBancos = oDBCred.RecuperaRelaBancosPersonaHojaApr(psCodNat, psCodJur, psCtaCod)
    Set rResGarTitAva = oDBCred.RecuperaResumenGarTitularAval(psCtaCod)
    Set REstCivConvenio = oDPersGeneral.RecuperaEstCivConvenio(psCtaCod, Format$(gdFecSis, "yyyymmdd"))
    
    sConstante = oConst.LeeConstSistema(473)
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(gdFecSis, TCFijoMes)
    Call ImprimeHojaAprobacionCred(R, RCaliSbs, RRelaCred, rBancos, RDatFin, rResGarTitAva, RGarantCred, psCtaCod, RCredEval, RCredAmp, RCalfSbsRel, RExoAutCred, RRiesgoUnico, RCredGarant, RCredResulNivApr, RNivApr, sRiesgo, sConstante, ROpRiesgo, RComentAnalis, nTipoCambioFijo)
    
    Set oGeneral = Nothing
    Set oConst = Nothing
    Set oDPersGeneral = Nothing
    Set oDBCred = Nothing
    Set oDGarantia = Nothing
    Set oDCred = Nothing
    Set oDLeasing = Nothing
End Sub


Private Sub ImprimeHojaAprobacionCred(ByRef pR As ADODB.Recordset, ByRef pRB As ADODB.Recordset, ByRef pRCliRela As ADODB.Recordset, ByRef pRRelaBcos As ADODB.Recordset, _
            ByRef pRDatFinan As ADODB.Recordset, ByRef pRResGarTitAva As ADODB.Recordset, ByRef pRGarantCred As ADODB.Recordset, _
            ByVal psCodCta, ByVal prsCredEval As ADODB.Recordset, ByVal prsCredAmp As ADODB.Recordset, ByVal prsCalfSBSRela As ADODB.Recordset, _
            ByVal prsExoAutCred As ADODB.Recordset, ByVal prsRiesgoUnico As ADODB.Recordset, ByVal prsCredGarant As ADODB.Recordset, _
            ByVal prsCredResulNivApr As ADODB.Recordset, ByVal pRNivApr As ADODB.Recordset, ByVal psRiesgo As String, ByVal psHabilitaNiveles As String, _
            ByVal prsOpRiesgo As ADODB.Recordset, ByVal prsComentAnalis As ADODB.Recordset, ByVal pnTpoCambio As Double)
    Dim oFun As New COMFunciones.FCOMImpresion
    'Dim oCliPre As New COMNCredito.NCOMCredito 'COMENTADO POR ARLO 20170722
    Dim oDCredExoAut As New COMDCredito.DCOMNivelAprobacion
    Dim rsConDPer As New comdpersona.DCOMPersonas
    Dim oCrDoc As New COMDCredito.DCOMCredDoc
    Dim oDoc  As New cPDF
    Dim lrDatosRatiosFI As New ADODB.Recordset
    Dim lrsRespNvlExo As ADODB.Recordset
    Dim lrsIngGas As ADODB.Recordset
    Dim oGarant As New COMNCredito.NCOMGarantia
    Dim bAplicaRatio As Boolean
    Dim lnliquidez As Double, lnCapacidadPago As Double, lnExcedente As Double, lnInventario As Double, lnCuota As Double, lnRemNeta As Double, lnEgresos As Double
    Dim lnPatriEmpre As Double, lnPatrimonio As Double, lnIngresoNeto As Double, lnRentabPatrimonial As Double, lnEndeudamiento As Double
    Dim lnMontoRiesgoUnico As Double, lnMontoExpEstCred As Double
    Dim bValidarCliPre As Boolean
    Dim nTipo As Integer, i As Integer, a As Integer, nPosicion As Integer
    Dim dFecFteIng As Date
    Dim nIngNeto As Double, nGasFamiliar As Double, nExpoCredAct As Double
    Dim nFTabla As Integer, nFTablaCabecera As Integer, lnFontSizeBody As Integer, nCanGarant As Integer
    Dim nValComer As Double, nValReali As Double, nValUtiliz As Double, nValDisp As Double, nValGrava As Double
    Dim nPosAmp As Integer, nCanIFs As Integer, nAdicionaFila As Integer, nIndx As Integer, nTmpPosic As Integer
    
    On Error GoTo ErrorImprimirPDF
    
    If oCrDoc.ObtenerFechaProNumFuente(pR!cNumeroFuente, gdFecSis).RecordCount > 0 Then
        dFecFteIng = Format(oCrDoc.ObtenerFechaProNumFuente(pR!cNumeroFuente, gdFecSis)!valor, "dd/mm/yyyy")
    End If
    Set lrDatosRatiosFI = oCrDoc.ObtieneRatiosFI(pR!cNumeroFuente, psCodCta)
    
    lnliquidez = 0: lnEndeudamiento = 0: lnRentabPatrimonial = 0: lnCapacidadPago = 0: lnPatrimonio = 0: lnInventario = 0: lnExcedente = 0: lnCuota = 0: lnRemNeta = 0: lnEgresos = 0
    
    If Not (lrDatosRatiosFI.EOF And lrDatosRatiosFI.BOF) Then
        Dim nIndexRatios As Integer
        For nIndexRatios = 0 To lrDatosRatiosFI.RecordCount - 1
            If lrDatosRatiosFI!nValor = 1 Then
                lnInventario = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 2 Then
                lnExcedente = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 3 Then
                lnCuota = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 4 Then
                lnPatrimonio = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 5 Then
                lnEndeudamiento = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 6 Then
                lnliquidez = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 9 Then
                lnRemNeta = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 10 Then
                lnEgresos = lrDatosRatiosFI!nMonto
            End If
            
            lrDatosRatiosFI.MoveNext
        Next
        lnCapacidadPago = lnCuota / IIf(lnExcedente = 0, 1, lnExcedente)
        lnRentabPatrimonial = IIf(lnExcedente = 0, 1, lnExcedente) / IIf(lnPatrimonio = 0, 1, lnPatrimonio)
        lnEndeudamiento = lnEndeudamiento / IIf(lnPatrimonio = 0, 1, lnPatrimonio)
    End If
    
    Set lrsIngGas = rsConDPer.ObtenerDatosHojEvaluaci(pR!cNumeroFuente, dFecFteIng)

    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Hoja de Aprobación de Créditos Nº " & gsCodUser
    oDoc.Title = "Hoja de Aprobación de Créditos Nº " & gsCodUser
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\" & IIf(nTipo = 1, "Previo", "") & psCodCta & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    oDoc.Fonts.Add "F1", "Times New Roman", TrueType, Normal, WinAnsiEncoding: oDoc.Fonts.Add "F2", "Times New Roman", TrueType, Bold, WinAnsiEncoding
    oDoc.Fonts.Add "F3", "Arial Narrow", TrueType, Normal, WinAnsiEncoding: oDoc.Fonts.Add "F4", "Arial Narrow", TrueType, Bold, WinAnsiEncoding
    
    nFTablaCabecera = 7: nFTabla = 7: lnFontSizeBody = 7

    oDoc.NewPage A4_Vertical
    lnMontoRiesgoUnico = 0#
    
    'bValidarCliPre = oCliPre.ValidarClientePreferencial(pR!cPersCod) 'COMENTADO POR ARLO 20170722
    bValidarCliPre = False 'ARLO 20170722

    If Not (prsRiesgoUnico.EOF And prsRiesgoUnico.BOF) Then
        lnMontoRiesgoUnico = prsRiesgoUnico!nMonto
    End If

    If Not (lrsIngGas.BOF And lrsIngGas.EOF) Then
        Dim j As Integer
        For j = 1 To lrsIngGas.RecordCount
            If (lrsIngGas!cCodHojEval = 300102) Then
                nIngNeto = Format(lrsIngGas!nUnico, gcFormView)
            ElseIf (lrsIngGas!cCodHojEval = 300201) Then
                nGasFamiliar = Format(lrsIngGas!nUnico, gcFormView)
            ElseIf (lrsIngGas!cCodHojEval = 400205) Then
                nExpoCredAct = Format(lrsIngGas!nUnico, gcFormView)
            End If
            lrsIngGas.MoveNext
        Next
    End If
    
    If bValidarCliPre = True Then
        oDoc.WTextBox 40, 70, 10, 450, "HOJA DE APROBACIÓN - CLIENTE PREFERENCIAL", "F3", 11, hCenter, , vbBlack
    Else
        oDoc.WTextBox 40, 70, 10, 450, "HOJA DE APROBACIÓN", "F3", 11, hCenter, , vbBlack
    End If
    
    oDoc.WTextBox 60, 55, 10, 200, "APROBACION DE CREDITOS", "F4", 9, hLeft
    oDoc.WTextBox 60, 354, 10, 200, "FECHA APROBACION:....................", "F4", 9, hRight
    oDoc.WTextBox 80, 55, 77, 510, "", "F3", 7, hLeft, , , 1, vbBlack
    oDoc.WTextBox 80, 57, 12, 450, "Cliente:", "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 88, 57, 12, 450, "DNI/RUC:", "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 96, 57, 12, 450, "Tipo de Credito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 104, 57, 12, 450, "Producto de Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 112, 57, 12, 450, "Monto :", "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 120, 57, 12, 450, "Exposición Riesgo Único:", "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 128, 57, 12, 450, "Fecha de Solicitud:", "F3", lnFontSizeBody, hLeft, , , , , , 2
    
    oDoc.WTextBox 80, 140, 12, 450, UCase(pR!Prestatario), "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 88, 140, 12, 450, pR!DniRuc & "/" & pR!Ruc, "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 96, 140, 12, 450, pR!Tipo_Cred, "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 104, 140, 12, 450, pR!Tipo_Prod, "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 112, 140, 12, 450, Format(pR!Ptmo_Propto, gcFormView), "F4", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 120, 140, 12, 450, Format(lnMontoRiesgoUnico, gcFormView), "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 128, 140, 12, 450, pR!Fec_Soli, "F3", lnFontSizeBody, hLeft, , , , , , 2
    
    lnMontoExpEstCred = nExpoCredAct
    
    oDoc.WTextBox 80, 350, 12, 450, "Nº Cta. Fianza:", "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 88, 350, 12, 450, "Condición", "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 96, 350, 12, 450, "Cod. Analista:", "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 104, 350, 12, 450, "Agencia:", "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 112, 350, 12, 450, "Giro/Act Negocio:", "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 120, 350, 12, 450, "Direc. Principal:", "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 128, 350, 12, 450, "Direc. Negocio:", "F3", lnFontSizeBody, hLeft, , , , , , 2
    
    oDoc.WTextBox 80, 400, 12, 450, pR!Nro_Credito, "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 88, 400, 12, 450, UCase(pR!cCondicion), "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 96, 400, 12, 450, UCase(pR!Analista), "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 104, 400, 12, 450, UCase(pR!Oficina), "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 112, 400, 12, 450, pR!ActiGiro, "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 120, 400, 12, 450, pR!dire_domicilio, "F3", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 128, 400, 12, 450, pR!dire_trabajo, "F3", lnFontSizeBody, hLeft, , , , , , 2
    
    If Not (prsOpRiesgo.EOF And prsOpRiesgo.BOF) Then
        oDoc.WTextBox 144, 400, 12, 450, prsOpRiesgo!cRiesgoValor, "F3", lnFontSizeBody, hLeft, , , , , , 2
    End If
    
    oDoc.WTextBox 163, 55, 60, 490, "DATOS DE LA CARTA FIANZA", "F4", lnFontSizeBody, hLeft
    oDoc.WTextBox 172, 55, 50, 510, "", "F3", lnFontSizeBody, hLeft, , , 1, vbBlack
    
    oDoc.WTextBox 173, 57, 12, 450, "Acreedor:", "F4", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 173, 140, 12, 450, pR!cAcreedor, "F3", lnFontSizeBody, hLeft, , , , , , 2
    
    oDoc.WTextBox 173, 350, 12, 450, "Avalado:", "F4", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 173, 400, 12, 450, pR!cAvalado, "F3", lnFontSizeBody, hLeft, , , , , , 2
    
    oDoc.WTextBox 183, 57, 12, 450, "Modalidad:", "F4", lnFontSizeBody, hLeft, , , , , , 2
    'oDoc.WTextBox 183, 140, 12, 450, UCase(pR!cModalidad), "F3", lnFontSizeBody, hLeft, , , , , , 2'Comento JOEP20181221 CP
    oDoc.WTextBox 183, 140, 12, 450, IIf(pR!nModalidad = 13, UCase(pR!OtrsModalidades), UCase(pR!cModalidad)), "F3", lnFontSizeBody, hLeft, , , , , , 2 'JOEP20181221 CP
    
    oDoc.WTextBox 183, 350, 12, 450, "Fec. Vencimiento:", "F4", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 183, 400, 12, 450, pR!dVencimiento, "F3", lnFontSizeBody, hLeft, , , , , , 2
    
    
    oDoc.WTextBox 192, 57, 12, 510, "Finalidad:", "F4", lnFontSizeBody, hLeft, , , , , , 2
    oDoc.WTextBox 200, 57, 12, 510, pR!cFinalidad, "F3", lnFontSizeBody, hLeft, , , , , , 2
    
    If CInt(pR!bExononeracionTasa) = 1 Then
        oDoc.WTextBox 210, 360, 12, 450, "EXONERADO DE TASA", "F3", lnFontSizeBody, hLeft, , , , , , 2
    End If
    
    nPosicion = 40
    
    oDoc.WTextBox 206 + nPosicion, 55 + 2, 60, 490, "GARANTIAS", "F4", lnFontSizeBody, hLeft
    
    oDoc.WTextBox 215 + nPosicion, 57, 12, 36, "COD", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox 215 + nPosicion, 93, 12, 60, "Tipo de Garantia", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox 215 + nPosicion, 153, 12, 15, "RG", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox 215 + nPosicion, 188 - 20, 12, 76, "Documento", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox 215 + nPosicion, 264 - 20, 12, 76, "Dirección", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox 215 + nPosicion, 340 - 20, 12, 20, "Mon.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox 215 + nPosicion, 367 - 27, 12, 40, "Val. Come.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox 215 + nPosicion, 407 - 27, 12, 40, "Val. Real.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox 215 + nPosicion, 447 - 27, 12, 40, "Val. Utilz.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox 215 + nPosicion, 487 - 27, 12, 40, "Val. Disp.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox 215 + nPosicion, 527 - 27, 12, 40, "Val. Grav.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox 215 + nPosicion, 567 - 27, 12, 25, "PROP.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    nPosicion = 227 + 40

    If Not (prsCredGarant.BOF And prsCredGarant.EOF) Then
        If prsCredGarant.RecordCount > 0 Then
            bAplicaRatio = oGarant.ObtieneGarantLiq(prsCredGarant!cNumGarant)
            psRiesgo = "Riesgo 2"
            For i = 1 To prsCredGarant.RecordCount
                Dim nAltoAdic As Integer, nValorMayor As Integer
                Dim nValCom As Double, nValRea As Double, nValUti As Double, nValDis As Double, nValGra As Double
                        
                    nValCom = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nTasacion * pnTpoCambio, prsCredGarant!nTasacion)
                    nValRea = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nRealizacion * pnTpoCambio, prsCredGarant!nRealizacion)
                    nValUti = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nGravado * pnTpoCambio, prsCredGarant!nGravado)
                    nValDis = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nDisponible * pnTpoCambio, prsCredGarant!nDisponible)
                    nValGra = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nValorGravado * pnTpoCambio, prsCredGarant!nValorGravado)
                    nValorMayor = ValorMayor(Len(prsCredGarant!cTpoGarant), Len(prsCredGarant!cClasGarant), Len(prsCredGarant!cDocDesc), Len(prsCredGarant!cDireccion))
                    nAltoAdic = (nValorMayor / 13) * 6
                    
                    If Trim(prsCredGarant!cClasGarant) = "GARANTIAS NO PREFERIDAS" Then
                        psRiesgo = "Riesgo 1"
                    End If
                    
                    oDoc.WTextBox nPosicion + a, 57, 12 + nAltoAdic, 36, prsCredGarant!cNumGarant, "F1", nFTabla, hCenter, , , 1, , , 2
                    oDoc.WTextBox nPosicion + a, 93, 12 + nAltoAdic, 60, prsCredGarant!cTpoGarant, "F1", nFTabla, hjustify, , , 1, , , 2
                    oDoc.WTextBox nPosicion + a, 153, 12 + nAltoAdic, 15, IIf(Trim(prsCredGarant!cClasGarant) = "GARANTIAS NO PREFERIDAS", 1, 2), "F1", nFTabla, hCenter, , , 1, , , 2
                    oDoc.WTextBox nPosicion + a, 188 - 20, 12 + nAltoAdic, 76, prsCredGarant!cDocDesc & " - Nº " & prsCredGarant!cNroDoc, "F1", nFTabla, hjustify, , , 1, , , 2
                    oDoc.WTextBox nPosicion + a, 264 - 20, 12 + nAltoAdic, 76, prsCredGarant!cDireccion, "F1", nFTabla, hLeft, , , 1, , , 2
                    oDoc.WTextBox nPosicion + a, 340 - 20, 12 + nAltoAdic, 20, prsCredGarant!cMoneda, "F1", nFTabla, hCenter, , , 1, , , 2
                    oDoc.WTextBox nPosicion + a, 367 - 27, 12 + nAltoAdic, 40, Format(nValCom, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                    oDoc.WTextBox nPosicion + a, 407 - 27, 12 + nAltoAdic, 40, Format(nValRea, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                    oDoc.WTextBox nPosicion + a, 447 - 27, 12 + nAltoAdic, 40, Format(nValUti, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                    oDoc.WTextBox nPosicion + a, 487 - 27, 12 + nAltoAdic, 40, Format(nValDis, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                    oDoc.WTextBox nPosicion + a, 527 - 27, 12 + nAltoAdic, 40, Format(nValGra, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                    oDoc.WTextBox nPosicion + a, 567 - 27, 12 + nAltoAdic, 25, Mid(prsCredGarant!cRelGarant, 1, 3) & ".", "F1", nFTabla, hCenter, , , 1, , , 2
                    nValComer = nValComer + nValCom
                    nValReali = nValReali + nValRea
                    nValUtiliz = nValUtiliz + nValUti
                    nValDisp = nValDisp + nValDis
                    nValGrava = nValGrava + nValGra
                    prsCredGarant.MoveNext
                    a = a + 12 + nAltoAdic
            Next
        End If
    End If
            
    oDoc.WTextBox nPosicion + a, 57, 12, 283, "TOTALES", "F1", nFTabla, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion + a, 340, 12, 40, Format(nValComer, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
    oDoc.WTextBox nPosicion + a, 380, 12, 40, Format(nValReali, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
    oDoc.WTextBox nPosicion + a, 420, 12, 40, Format(nValUtiliz, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
    oDoc.WTextBox nPosicion + a, 460, 12, 40, Format(nValDisp, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
    oDoc.WTextBox nPosicion + a, 500, 12, 40, Format(nValGrava, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
    oDoc.WTextBox nPosicion + a, 540, 12, 25, "", "F1", nFTabla, hLeft, , , 1, , , 2
    nPosicion = nPosicion + a + 20: i = 0: a = 0
            
    oDoc.WTextBox nPosicion + 2, 55, 60, 490, "COBERTURA DE GARANTIA", "F4", lnFontSizeBody, hLeft
    nPosAmp = nPosicion: nPosicion = nPosicion + 12
            
    oDoc.WTextBox nPosicion, 55, 12, 95, "Cobertura Exp. Este Crédito", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion, 150, 12, 95, "Cobertura Exp. Riesgo Único", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion, 245, 12, 95, "Tipo de Riesgo", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    nPosicion = nPosicion + 12
            
    If lnMontoExpEstCred = 0 Then
        oDoc.WTextBox nPosicion + a, 55, 12, 95, Format(0, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    Else
        oDoc.WTextBox nPosicion + a, 55, 12, 95, Format(nValReali / lnMontoExpEstCred, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    End If
            
    If lnMontoRiesgoUnico = 0 Then
        oDoc.WTextBox nPosicion + a, 150, 12, 95, Format(0, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    Else
        oDoc.WTextBox nPosicion + a, 150, 12, 95, Format(nValReali / lnMontoRiesgoUnico, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    End If
    
    oDoc.WTextBox nPosicion + a, 245, 12, 95, UCase(psRiesgo), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    nPosicion = nPosicion + 25
            
    If Not (prsCredAmp.BOF And prsCredAmp.EOF) Then
        oDoc.WTextBox nPosAmp + 2, 360, 60, 490, "AMPLIACIÓN DE CRÉDITO", "F4", lnFontSizeBody, hLeft
        nPosAmp = nPosAmp + 12
        oDoc.WTextBox nPosAmp, 360, 12, 77, "Crédito Nº", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosAmp, 437, 12, 60, "Saldo Capital", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        nPosAmp = nPosAmp + 12
        For i = 1 To prsCredAmp.RecordCount
            oDoc.WTextBox nPosAmp + a, 360, 12, 77, prsCredAmp!cCtaCodAmp, "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosAmp + a, 437, 12, 60, Format(prsCredAmp!nMonto, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
            prsCredAmp.MoveNext
            a = a + 12
        Next
        nPosicion = nPosAmp + a + 10
    End If
    
    If Not (pRRelaBcos.EOF And pRRelaBcos.BOF) Then
        nCanIFs = pRRelaBcos.RecordCount
    End If
    
    If nCanIFs > 5 Then
        nAdicionaFila = nAdicionaFila + 15
    End If
    
    If nCanIFs > 7 Then
        nAdicionaFila = nAdicionaFila + 15
    End If
        
    If nCanIFs > 9 Then
        nAdicionaFila = nAdicionaFila + 15
    End If
    
    nPosicion = nPosicion
    oDoc.WTextBox nPosicion + 6, 55, 60, 490, "RATIOS FINANCIEROS", "F4", lnFontSizeBody, hLeft
    oDoc.WTextBox nPosicion + 16, 55, 56 + nAdicionaFila, 240, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion + 16, 295, 56 + nAdicionaFila, 240, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion + 16, 295, 12, 139, "Institución", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion + 16, 434, 12, 28, "Moneda", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion + 16, 462, 12, 40, "Saldo", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion + 16, 502, 12, 33, "Relacion", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    nPosicion = nPosicion + 9
    oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Liquidez", "F1", 5, hLeft, , , , , , 2
    
    nTmpPosic = nPosicion + 8
    
    For nIndx = 1 To pRRelaBcos.RecordCount
        If Len(pRRelaBcos!Nombre) > 30 Then
            oDoc.WTextBox nTmpPosic + 9, 295, 9, 139, pRRelaBcos!Nombre, "F1", 5, hLeft, , , , , , 2
        Else
            oDoc.WTextBox nTmpPosic + 9, 295, 9, 139, pRRelaBcos!Nombre, "F1", nFTablaCabecera, hLeft, , , , , , 2
        End If
        
        oDoc.WTextBox nTmpPosic + 9, 434, 9, 28, pRRelaBcos!Moneda, "F1", nFTablaCabecera, hCenter, , , , , , 2
        oDoc.WTextBox nTmpPosic + 9, 462, 9, 40, Format(pRRelaBcos!Saldo, gcFormView), "F1", nFTablaCabecera, hRight, , , , , , 2
        oDoc.WTextBox nTmpPosic + 9, 502, 9, 33, Mid(pRRelaBcos!Relacion, 1, 3) & ".", "F1", nFTablaCabecera, hLeft, , , , , , 2
        pRRelaBcos.MoveNext
        nTmpPosic = nTmpPosic + 8
    Next
        
    If bAplicaRatio = True Then
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(0#, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Patrimonio", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(0#, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Endeudamiento Patrimonial", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(0# * 100, "0.00") & "%", "F1", nFTablaCabecera, hRight, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Inventario", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(0#, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Rentabilidad Patrimonial", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(0# * 100, "0.00") & "%", "F1", nFTablaCabecera, hRight, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Excedente", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(0#, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Capacid. De Pago", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(0# * 100, "0.00") & "%", "F1", nFTablaCabecera, hRight, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Cuota", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(0#, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12 + 20 + nAdicionaFila
    Else
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(lnliquidez, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Patrimonio", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(lnPatrimonio, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Endeudamiento Patrimonial", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(lnEndeudamiento * 100, "0.00") & "%", "F1", nFTablaCabecera, hRight, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Inventario", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(lnInventario, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Rentabilidad Patrimonial", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(lnRentabPatrimonial * 100, "0.00") & "%", "F1", nFTablaCabecera, hRight, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Excedente", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(lnExcedente, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Capacid. De Pago", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(lnCapacidadPago * 100, "0.00") & "%", "F1", nFTablaCabecera, hRight, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Cuota", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(lnCuota, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12 + 20 + nAdicionaFila
    End If
    oDoc.WTextBox nPosicion, 55, 60, 490, "CALIFICACIÓN Y RELACIÓN DE TITULARES / CÓNYUGE / AVALES", "F4", 7, hLeft
    oDoc.WTextBox nPosicion + 9, 55, 56, 480, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion + 9, 55, 12, 240, "Nombre", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion + 9, 295, 12, 115, "Relación", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion + 9, 410, 12, 25, "Normal", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion + 9, 435, 12, 25, "Poten.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion + 9, 460, 12, 25, "Defic.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion + 9, 485, 12, 25, "Dudos", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion + 9, 510, 12, 25, "Pérdida", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    
    If Not (prsCalfSBSRela.EOF And prsCalfSBSRela.BOF) Then
        Dim nIndx2 As Integer, nTmpPosic2 As Integer
        nTmpPosic2 = nPosicion + 14
        
        For nIndx2 = 1 To prsCalfSBSRela.RecordCount
            oDoc.WTextBox nTmpPosic2 + 9, 55, 11, 240, prsCalfSBSRela!cPersNombre, "F1", nFTablaCabecera, hLeft, , , , , , 1
            oDoc.WTextBox nTmpPosic2 + 9, 295, 11, 115, prsCalfSBSRela!cConsDescripcion, "F1", nFTablaCabecera, hLeft, , , , , , 1
            oDoc.WTextBox nTmpPosic2 + 9, 410, 11, 25, prsCalfSBSRela!Normal & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
            oDoc.WTextBox nTmpPosic2 + 9, 435, 11, 25, prsCalfSBSRela!Potencial & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
            oDoc.WTextBox nTmpPosic2 + 9, 460, 11, 25, prsCalfSBSRela!DEFICIENTE & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
            oDoc.WTextBox nTmpPosic2 + 9, 485, 11, 25, prsCalfSBSRela!DUDOSO & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
            oDoc.WTextBox nTmpPosic2 + 9, 510, 11, 25, prsCalfSBSRela!PERDIDA & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
            prsCalfSBSRela.MoveNext
            nTmpPosic2 = nTmpPosic2 + 8
        Next
    End If
    
    nPosicion = nPosicion + 70
    If Not (prsComentAnalis.BOF And prsComentAnalis.EOF) Then
        oDoc.WTextBox nPosicion, 55, 56, 330, "COMENTARIO ANALISTA", "F1", nFTablaCabecera, hLeft
        nPosicion = nPosicion + 9
        oDoc.WTextBox nPosicion, 55, 40, 480, prsComentAnalis!cComentAnalista, "F1", nFTablaCabecera, hjustify, , , 1, , , 2
        nPosicion = nPosicion + 48
    End If

    oDoc.WTextBox nPosicion, 55, 56, 330, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion, 55, 56, 165, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion, 55, 12, 165, "EXONERACIONES", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion, 220, 12, 165, "AUTORIZACIONES", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion, 385, 12, 150, "NIVELES DE APROBACION POR EXPOSICION", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
    oDoc.WTextBox nPosicion, 385, 56, 150, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        
    If psHabilitaNiveles = 1 Then
        If Not (prsExoAutCred.EOF And prsExoAutCred.BOF) Then
            Dim nIndx3 As Integer, nTmpPosic3 As Integer, nTmpPosicExo As Integer, nTmpPosicAut As Integer
            nTmpPosic3 = nPosicion + 14
            nTmpPosicExo = nPosicion + 2
            nTmpPosicAut = nPosicion + 2
                
            For nIndx3 = 1 To prsExoAutCred.RecordCount
                Dim texto As String
                
                If prsExoAutCred!nTipoExoneraCod = 1 Then
                    Set lrsRespNvlExo = oDCredExoAut.RecuperaRespExoAut(prsExoAutCred!cExoneraCod, IIf(lnMontoRiesgoUnico = 0, lnMontoExpEstCred, lnMontoRiesgoUnico))
                    oDoc.WTextBox nTmpPosicExo + 9, 55, 9, 150, prsExoAutCred!cExoneraDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                
                    If Not (lrsRespNvlExo.BOF And lrsRespNvlExo.EOF) Then
                        oDoc.WTextBox nTmpPosicExo + 9, 145, 9, 150, lrsRespNvlExo!cNivAprDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                    End If
                    nTmpPosicExo = nTmpPosicExo + 5
                Else
                    Set lrsRespNvlExo = oDCredExoAut.RecuperaRespExoAut(prsExoAutCred!cExoneraCod, IIf(lnMontoRiesgoUnico = 0, lnMontoExpEstCred, lnMontoRiesgoUnico))
                    oDoc.WTextBox nTmpPosicAut + 9, 225, 9, 240, prsExoAutCred!cExoneraDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                
                    If Not (lrsRespNvlExo.BOF And lrsRespNvlExo.EOF) Then
                        oDoc.WTextBox nTmpPosicAut + 9, 295, 9, 150, lrsRespNvlExo!cNivAprDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                    End If
                    nTmpPosicAut = nTmpPosicAut + 5
                End If
                prsExoAutCred.MoveNext
                Set lrsRespNvlExo = Nothing
            Next
        End If
            
        If Not (prsCredResulNivApr.EOF And prsCredResulNivApr.BOF) Then
            Dim nIndx4 As Integer, nTmpPosic4 As Integer
            nTmpPosic4 = nPosicion + 14
            For nIndx = 1 To prsCredResulNivApr.RecordCount
                oDoc.WTextBox nTmpPosic4 + 12, 390, 12, 305, prsCredResulNivApr!cNivAprDesc, "F1", nFTablaCabecera, hLeft
                nTmpPosic4 = nTmpPosic4 + 5
                prsCredResulNivApr.MoveNext
            Next
        End If
    End If
    nPosicion = nPosicion + 50
    oDoc.WTextBox nPosicion + 15, 150, 56, 150, "RESOLUCION DE COMITÉ, EN CONCLUSION: ", "F1", nFTablaCabecera, Left
    nPosicion = nPosicion + 12
    oDoc.WTextBox nPosicion + 15, 150, 56, 70, "MONTO", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 250, 56, 70, "CUOTAS", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 350, 56, 70, "TI", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 450, 56, 70, "VCTO", "F1", nFTablaCabecera, Left
    nPosicion = nPosicion + 20
    oDoc.WTextBox nPosicion + 15, 70, 56, 150, "APROBADO POR: ", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 150, 56, 70, "...................", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 250, 56, 70, "...................", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 350, 56, 70, "...................", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 450, 56, 70, "...................", "F1", nFTablaCabecera, Left
    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ErrorImprimirPDF:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Public Function ValorMayor(ByVal nV1 As Integer, ByVal nV2 As Integer, ByVal nV3 As Integer, ByVal nV4 As Integer) As Integer
    Dim nValM As Integer
    Dim nArreglo(4) As Integer
    Dim i As Integer, j As Integer
    
    nArreglo(0) = nV1: nArreglo(1) = nV2: nArreglo(2) = nV3: nArreglo(3) = nV4: ValorMayor = 0
    For i = 0 To 3
        Dim nContador As Integer
        For j = 0 To 3
            If nArreglo(i) >= nArreglo(j) Then
                nContador = nContador + 1
            End If
        Next
        If nContador = 4 Then
            ValorMayor = nArreglo(i)
        End If
        nContador = 0
    Next
End Function

Private Sub LimpiarFormulario()
    ActXCodCta.CMAC = "109"
    ActXCodCta.Age = gsCodAge
    ActXCodCta.Prod = "514"
    ActXCodCta.Cuenta = ""
    lblTitular.Caption = ""
    lblMonto.Caption = ""
    lblFecDesembolso.Caption = ""
    lblEstado.Caption = ""
End Sub

Private Sub Form_Load()
    Call LimpiarFormulario
End Sub
