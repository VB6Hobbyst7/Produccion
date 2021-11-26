VERSION 5.00
Begin VB.Form frmColPDesembCredAmpliado 
   Caption         =   "Credito Pignoraticio - Ampliacíon de Crédito"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7785
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   7125
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   4080
      TabIndex        =   18
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmbCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5280
      TabIndex        =   17
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Montos"
      Height          =   2415
      Left            =   4800
      TabIndex        =   10
      Top             =   4200
      Width           =   2775
      Begin VB.Label lblInteres 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Interes:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblSaldoNeto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblITF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   24
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblSaldoBruto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblMontoCredant 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblMontoCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Saldo Neto:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Cancel Cred Ant.:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "ITF:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo Bruto:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Monto Crédito:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdBuscaCtaAbono 
      Height          =   345
      Left            =   3840
      Picture         =   "frmColPDesembCredAmpliado.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Buscar ..."
      Top             =   5055
      Width           =   420
   End
   Begin SICMACT.ActXCodCta ActXCodCtaAbono 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Texto           =   "Cta Abono"
      EnabledCta      =   -1  'True
      Age             =   "01"
      Prod            =   "705"
      CMAC            =   "109"
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   4
      Top             =   4800
      Width           =   4575
      Begin VB.CheckBox chkDesembAbonoCta 
         Caption         =   "Desembolso con abono en Cuenta"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   0
         Width           =   2775
      End
      Begin VB.Label lblCapTpoProducto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label lblCapTitularCta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   19
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Titular:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   880
         Width           =   735
      End
   End
   Begin SICMACT.ActXCodCta ActXCodCtaCancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Texto           =   "Cta a Canc."
      Age             =   "01"
      Prod            =   "705"
      CMAC            =   "109"
   End
   Begin SICMACT.ActXColPDesCon AXDesCon 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6165
   End
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      Texto           =   "Contrato"
      EnabledCta      =   -1  'True
      Age             =   "01"
      Prod            =   "705"
      CMAC            =   "109"
   End
   Begin VB.CommandButton cmdBuscar 
      Height          =   345
      Left            =   3840
      Picture         =   "frmColPDesembCredAmpliado.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Buscar ..."
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "frmColPDesembCredAmpliado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'DESEMBOLSO DE CREDITO PIGNORATICIO AMPLIADO
'Archivo:  frmColPDesembCredAmpliado
'RECO   :  28/01/2014.
'Resumen:  Formulario realizar el desembolso de un crédito pignoraticio ampliado

Option Explicit

Dim lsPersCodTitular As String

Dim loItf1 As Double, loItf2 As Double
Dim vFecVenc As Date
Dim lnOpeCod As Long
Dim lsPersCod As String

Dim fsVarOpeDesc As String
Dim fnVarOpeCod As Long
Dim nRedondeoITF As Double
'****Cancelacion Credito Anterior*******
Dim fnVarOpeCodCanc As Long
Dim fsVarOpeDescCanc As String
Dim fsVarPersCodCMACCanc As String
Dim fsVarNombreCMACCanc As String

Dim fnTasaPreparacionRemateCanc As Double
Dim fnTasaCustodiaVencidaCanc As Double
Dim fnTasaImpuestoCanc As Double
Dim fnVarDiasCambCartCanc As Integer

Dim fnVarTasaInteresCanc As Double

Dim vTasaInteresCanc As Double ' peac 20070820

Dim vNroContratoCanc As String
Dim vFecContratoCanc As Date
Dim vFecVencimientoCanc As Date
Dim vdiasAtrasoCanc As Double
Dim vDiasAtrasoRealCanc As Double
Dim vDeudaCanc As Double
Dim vInteresVencidoCanc As Double
Dim vInteresMoratorioCanc As Double
Dim vImpuestoCanc As Double
Dim fnVarGastoCorrespondenciaCanc As Double
Dim vCostoCustodiaMoratorioCanc As Double
Dim vOroNetoCanc As Double
Dim vValorTasacionCanc As Double
Dim vSaldoCapitalCanc As Double
Dim vNewSaldoCapitalCanc As Double
Dim vPlazoCanc As Integer
Dim vEstadoCanc As String
Dim vCostoPreparacionRemateCanc As Double
Dim vPrestamoCanc As Double
Dim vTasaInteresVencidoCanc As Double
Dim vTasaInteresMoratorioCanc As Double
Dim vCostoCustMoratorioCanc As Double
Dim v14kCanc As Double
Dim v16kCanc As Double
Dim v18kCanc As Double
Dim v21kCanc As Double

Dim fnVarCostoNotificacionCanc As Currency 'PEAC 20070926

Dim gnNotifiAdjuCanc As Integer  ' peac 20080515
Dim gnNotifiCobCanc As Integer  ' PEAC 20080715

Dim gcCredAntiguoCanc As String  ' peac 20070923
Dim fnVarEstUltProcRemCanc As Integer 'DAOR 20070714
Dim fsColocLineaCredPigCanc As String, vFecEstadoCanc As Date ' PEAC 20070813
Dim vDiasAdelCanc As Integer, vInteresAdelCanc As Double, vMontoColCanc As Double ' PEAC 20070813
Dim nRedondeoITFCanc As Double ' BRGO 20110914
'**********
'DESEMBOLSO ABONO CTA*****************************
Dim lbDesembCC As Boolean
Dim lbCuentaNueva As Boolean
Private pRSRela As ADODB.Recordset
'FIN DESEMBOLSO ABONO CTA***********************************
'************RECATE CREDITO ANTERIOR**************
Dim fnMaxDiasCustodiaDiferidaResc As Double
Dim fnTasaIGVResc As Double
Dim fnPorcentajeCustodiaDiferidaResc As Double

Dim vCostoCustodiaExtraResc As Double
Dim vSaldoCustodiaExtraResc As Double

Dim fnVarOpeCodResc As Long
Dim fsVarOpeDescResc As String
Dim objPistaResc As COMManejador.Pista
Dim nDiasTranscurridosResc As Integer


Dim nProgAhorros As Integer
Dim nMontoAbonar As Double
Dim nPlazoAbonar As Integer
Dim sPromotorAho As String
'************FIN RESCATE CREDITO ANTERIOR*********
'*****CUENTA AHORRO*********************************
Public nProducto As Producto
'FIN CUENTA AHORRO**************************


Private Sub chkDesembAbonoCta_Click()
    If chkDesembAbonoCta.value = 1 Then
        ActXCodCtaAbono.Enabled = True
        cmdBuscaCtaAbono.Enabled = True
    Else
        ActXCodCtaAbono.Enabled = False
        cmdBuscaCtaAbono.Enabled = False
    End If
End Sub

Private Sub cmbCancelar_Click()
    Call Limpiar
End Sub

Private Sub cmdBuscaCtaAbono_Click()
    Dim clsPers As COMDPersona.UCOMPersona 'UPersona
    Set clsPers = New COMDPersona.UCOMPersona
    Set clsPers = frmBuscaPersona.Inicio
    
    If Not clsPers Is Nothing Then
        Dim sPers As String
        Dim rsPers As New ADODB.Recordset
        Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
        Dim sCta As String
        Dim sRelac As String * 15
        Dim sEstado As String
        Dim sProd As String
        Dim clsCuenta As UCapCuenta
        sPers = clsPers.sPersCod
        Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsPers = clsCap.GetCuentasPersona(sPers, nProducto, , , , , gsCodAge)
        Set clsCap = Nothing
        If sPers <> lsPersCodTitular Then
            MsgBox "El titular del crédito debe ser el mismo titular de la cuanto de ahorros", vbCritical, "Aviso"
            Exit Sub
        End If
        
        If Not (rsPers.EOF And rsPers.EOF) Then
            Do While Not rsPers.EOF
                sCta = rsPers("cCtaCod")
                sRelac = rsPers("cRelacion")
                sEstado = Trim(rsPers("cEstado"))
                frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
                rsPers.MoveNext
            Loop
            Set clsCuenta = New UCapCuenta
            Set clsCuenta = frmCapMantenimientoCtas.Inicia
            If clsCuenta Is Nothing Then
            Else
                If clsCuenta.sCtaCod <> "" Then
                    ActXCodCtaAbono.NroCuenta = clsCuenta.sCtaCod
                    Me.lblCapTitularCta.Caption = clsPers.sPersNombre
                    Me.lblCapTpoProducto.Caption = IIf(Mid(clsCuenta.sCtaCod, 6, 3) = "232", "AHORROS", IIf(Mid(clsCuenta.sCtaCod, 6, 3) = "233", "PLAZO FIJO", IIf(Mid(clsCuenta.sCtaCod, 6, 3) = "234", "CTS", "")))
                    ActXCodCtaAbono.SetFocusCuenta
                    SendKeys "{Enter}"
                End If
            End If
            Set clsCuenta = Nothing
        Else
            MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
        End If
        rsPers.Close
        Set rsPers = Nothing
    End If
    Set clsPers = Nothing
    ActXCodCtaAbono.SetFocusCuenta
End Sub

Private Sub cmdBuscar_Click()
    
    Me.AXCodCta.NroCuenta = frmColPListaCredEstado.Inicio("Créditos para Desembolsar", "2100", gsCodAge, 1)
    If Me.AXCodCta.NroCuenta = "" Then
        Call cmdCancelar_Click
    End If
End Sub

Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.Limpiar
    Me.lblInteres.Caption = Format(0, "#0.00")
    Me.lblMontoCredito.Caption = Format(0, "#0.00")
    Me.lblMontoCredant.Caption = Format(0, "#0.00")
    Me.lblSaldoBruto.Caption = Format(0, "#0.00")
    Me.lblITF.Caption = Format(0, "#0.00")
    Me.lblSaldoNeto.Caption = Format(0, "#0.00")
    Me.lblInteres.Caption = Format(0, "#0.00")
    Me.chkDesembAbonoCta.value = 0
    Me.cmdBuscaCtaAbono.Enabled = False
    ActXCodCtaAbono.NroCuenta = ""
    lblCapTitularCta.Caption = ""
    lblCapTpoProducto.Caption = ""
    ActXCodCtaCancelar.NroCuenta = ""
    lsPersCodTitular = ""
    Me.AXCodCta.Enabled = True
End Sub

Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim lbok As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim loValMontoPrestamo As Double
Dim loValOtrosCostos As Double
Dim lsmensaje As String
Dim oCredAnt As ADODB.Recordset
Dim oNCOMColPContrato As COMNColoCPig.NCOMColPContrato

Set oNCOMColPContrato = New COMNColoCPig.NCOMColPContrato

Set oCredAnt = oNCOMColPContrato.ObtieneCuentaCancPignoAmpliado(Me.AXCodCta.NroCuenta)

If Not (oCredAnt.BOF And oCredAnt.EOF) Then
    Me.ActXCodCtaCancelar.NroCuenta = oCredAnt!cCtaCodAmp
    Me.lblMontoCredant.Caption = Format(oCredAnt!nMontoAmp, "#0.00")
    
        gITF.fgITFParamAsume (Mid(psNroContrato, 4, 2)), Mid(psNroContrato, 6, 3)
        'Valida Contrato
        Set lrValida = New ADODB.Recordset
        Set loValContrato = New COMNColoCPig.NCOMColPValida
            Set lrValida = loValContrato.nValidaDesembolsoCredPignoraticio(psNroContrato, gdFecSis, 0, lsmensaje)
            If Trim(lsmensaje) <> "" Then
                 MsgBox lsmensaje, vbInformation, "Aviso"
                 Call Limpiar
                 Exit Sub
            End If
    
        Set loValContrato = Nothing
        
        If lrValida Is Nothing Then ' Hubo un Error
            Limpiar
            Set lrValida = Nothing
            Exit Sub
        End If
        'Muestra Datos
        lbok = fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, False)
        Me.lblInteres.Caption = Format(lrValida!nInteres, "#0.00")
        Me.lblMontoCredito.Caption = Format(CCur(AXDesCon.SaldoCapital), "#0.00")
        Me.lblSaldoBruto.Caption = Format(Val(Me.lblMontoCredito.Caption) - Val(Me.lblMontoCredant.Caption), "#0.00")
        'vFecVencimiento = Format(lrValida!dVenc, "dd/mm/yyyy")
        'vFecVenc = Format(lrValida!dVenc, "dd/mm/yyyy")
        'vFecEstado = lrValida!dPrdEstado '
        'vValorTasacion = Format(lrValida!nTasacion, "#0.00")
        
        Set lrValida = Nothing
        
        loValMontoPrestamo = CCur(AXDesCon.SaldoCapital)
        'loValOtrosCostos = CCur(Me.txtCostoTasacion) + CCur(txtCostoCustodia.Text) + CCur(Me.lblInteres.Caption) + CCur(txtImpuesto.Text)
        
        'Me.lblNetoRecibir.Caption = Format(CCur(AXDesCon.SaldoCapital), "#0.00")
        nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
        ' **************  ITF ***************
        If gITF.gbITFAplica Then
            If Not gITF.gbITFAsumidocreditos Then
                loItf1 = Format(gITF.fgITFCalculaImpuesto(Val(Me.lblSaldoBruto.Caption)), "#0.00")
                If Val(Me.lblSaldoBruto.Caption) >= 1000 Then
                    Me.lblITF = Format(loItf1, "#0.00")
                Else
                    Me.lblITF = Format(0, "#0.00")
                End If
                
                Me.lblSaldoNeto.Caption = Format(CDbl(Me.lblSaldoBruto.Caption) - CDbl(Me.lblITF), "#0.00")
            Else
                loItf1 = Format(gITF.fgITFCalculaImpuesto(loValMontoPrestamo), "#0.00")
                loItf2 = Format(gITF.fgITFCalculaImpuesto(loValOtrosCostos), "#0.00")
                Me.lblITF = Format(loItf1 + loItf2, "#0.00")
                Me.lblSaldoNeto.Caption = Format(Val(Me.lblSaldoBruto.Caption), "#0.00")
            End If
        Else
            Me.lblITF = Format(0, "#0.00")
            Me.lblSaldoNeto.Caption = Format(Val(Me.lblSaldoBruto.Caption), "#0.00")
        End If
        ' **************  ITF ***************
        'fgCalculaDeuda
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
            
        AXCodCta.Enabled = False
Else
    MsgBox "Crédito no encontrado", vbCritical, "Aviso"
End If

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BuscaContrato (AXCodCta.NroCuenta)
        If ActXCodCtaCancelar.NroCuenta <> "" Then
            BuscaContratoCancelar (ActXCodCtaCancelar.NroCuenta)
        End If
        If ActXCodCtaCancelar.NroCuenta <> "" Then
            BuscaContratoRescate (ActXCodCtaCancelar.NroCuenta)
        End If
        If ActXCodCtaCancelar.NroCuenta <> "" Then
            lsPersCodTitular = AXDesCon.listaClientes.ListItems(1).Text
            If Me.lblSaldoNeto.Caption >= 3500 Then
                MsgBox "Debe seleccionar una cuenta de ahorro para depositar el desembolso. ya que el monto a desembolsar es mayor o igual a 3500", vbInformation, "Aviso"
                chkDesembAbonoCta.value = 1
                Me.ActXCodCtaAbono.Enabled = True
                Me.cmdBuscaCtaAbono.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
    cmdGrabar.Enabled = False
    AXCodCta.Enabled = True
    AXCodCta.SetFocusCuenta
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarDesem As COMNColoCPig.NCOMColPContrato
Dim loColImp As COMNColoCPig.NCOMColPImpre
Dim clsprevio As New previo.clsprevio
Dim lsCadImp As String
Dim opt As Integer
Dim OptBt2 As Integer
Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsFechaHoraPrend As String
Dim lsCuenta As String
Dim nFicSal As Integer
Dim psCtaAhoNew As String
Dim oITF  As COMDConstSistema.FCOMITF
Dim oMov As COMDMov.DCOMMov

Dim lnValorItfAbono As Double
'Dim lnValorItfGasto As Double
'Dim lnValorItfCancelacion As Double

Dim lnMovNro As Long
Dim MatDatosAho(14) As String
Dim lcTextImp As String
Dim lnSaldoCap As Currency, lnInteresComp As Currency, lnImpuesto As Currency
Dim lnCostoTasacion As Currency, lnCostoCustodia As Currency
Dim lnMontoEntregar As Currency
Set loColImp = New COMNColoCPig.NCOMColPImpre

Dim lnMontoTransaccion As Currency

psCtaAhoNew = ""
lnValorItfAbono = 0: 'lnValorItfGasto = 0: lnValorItfCancelacion = 0

lsCuenta = Me.ActXCodCtaAbono.NroCuenta
lnSaldoCap = Me.AXDesCon.SaldoCapital
lnInteresComp = CCur(Me.lblInteres.Caption)
'lnImpuesto = CCur(Me.txtImpuesto.Text)
'lnCostoCustodia = CCur(Me.txtCostoCustodia.Text)
'lnCostoTasacion = CCur(Me.txtCostoTasacion.Text)
'lnMontoEntregar = CCur(Me.lblSaldoNeto)
lnMontoEntregar = CCur(Me.lblMontoCredito.Caption)

lbDesembCC = IIf(Me.chkDesembAbonoCta.value = 1, True, False)
lbCuentaNueva = False
'COMENTADO RENZO
'If lbDesembCC And LblTipoAbono.Caption = "" Then
'    MsgBox "Por favor seleccione una Cuenta de Ahorro afecta a ITF...", vbInformation, "Atención"
'    Exit Sub
'End If
'FIN COMENTARIO
If Me.lblSaldoNeto.Caption >= 3500 And Me.ActXCodCtaAbono.NroCuenta = "" Then
    MsgBox "El monto es mayor o igual a 3500 debe selecionar una cuenta de ahorros", vbCritical, "Aviso"
    Exit Sub
End If
'EJVG20120322 Verifica actualización Persona
Dim oPersona As New COMNPersona.NCOMPersona
If oPersona.NecesitaActualizarDatos(lsPersCodTitular, gdFecSis) Then
     MsgBox "Para continuar con la Operación Ud. debe actualizar los datos del" & Chr(13) & "Titular", vbInformation, "Aviso"
     Dim foPersona As New frmPersona
     If Not foPersona.realizarMantenimiento(lsPersCodTitular) Then
         MsgBox "No se ha realizado la actualización de los datos del Titular la Operación no puede continuar!", vbInformation, "Aviso"
         Exit Sub
     End If
End If

If MsgBox(" Grabar Desembolso de Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing

        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        lsFechaHoraPrend = fgFechaHoraPrend(lsMovNro)
        lnMontoTransaccion = CCur(Me.lblMontoCredant.Caption)
        Set loGrabarDesem = New COMNColoCPig.NCOMColPContrato

        'Grabar Desembolso Pignoraticio
        Dim clsExo As New COMNCaptaServicios.NCOMCaptaServicios
        If Not clsExo.EsCuentaExoneradaLavadoDinero(Me.AXCodCta.NroCuenta) Then
          Dim sPersLavDinero As String
          Dim nMontoLavDinero As Double, nTC As Double
          Dim clsLav As New COMNCaptaGenerales.NCOMCaptaDefinicion, nmoneda As Integer, nMonto As Double

            nMonto = CDbl(Me.lblSaldoNeto.Caption)
            Set clsExo = Nothing
            sPersLavDinero = ""
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            nmoneda = gMonedaNacional
            If nmoneda = gMonedaNacional Then
                'Modificar
                Dim clsTC As COMDConstSistema.NCOMTipoCambio
                Set clsTC = New COMDConstSistema.NCOMTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
            Else
                nTC = 1
            End If
            If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                'By Capi 18022008 no aplica a desembolsos segun manifestado por riesgos
                'sPersLavDinero = IniciaLavDinero()
                 sPersLavDinero = ""
                'If sPersLavDinero = "" Then Exit Sub
            End If
         Else
            Set clsExo = Nothing
         End If
                        
        '*** PEAC 20071206 - no se graba el interes solo se muestra ***********************************
            'Call loGrabarDesem.nDesembolsoCredPignoraticio(lsCuenta, lnSaldoCap, lsFechaHoraGrab, _
                 lsMovNro, lnMontoEntregar, lnInteresComp, lnImpuesto, lnCostoTasacion, lnCostoCustodia, gITF.gbITFAplica, gITF.gbITFAsumidocreditos, loItf1, loItf2, CCur(lblNetoRecibir.Caption), False)
            
        '*** PEAC 20090505
            'Call loGrabarDesem.nDesembolsoCredPignoraticio(lsCuenta, lnSaldoCap, lsFechaHoraGrab, _
                 lsMovNro, lnMontoEntregar, 0, lnImpuesto, lnCostoTasacion, lnCostoCustodia, gITF.gbITFAplica, gITF.gbITFAsumidocreditos, loItf1, loItf2, CCur(lblNetoRecibir.Caption), False)
            Dim loRegPig As COMDColocPig.DCOMColPActualizaBD
            Set loRegPig = New COMDColocPig.DCOMColPActualizaBD
            'RECO20140311 ERS160-2013*********************************************************************
            'CONTROL DE SISTEMA LAVADO DE ACTIVOS
            Dim lbResultadoVisto As Boolean
            Dim loVistoElectronico As frmVistoElectronico
            Dim bVisto As Boolean
            Set loVistoElectronico = New frmVistoElectronico
            If lnSaldoCap > 5000 Then
                bVisto = True
            End If
            If loGrabarDesem.ObtieneMontoDesembCreDPigMes(lsPersCod, Format(gdFecSis, "yyyyMM")) + lnSaldoCap > 5000 Then
                bVisto = True
            End If
            If bVisto = True Then
                loVistoElectronico.VistoMovNro = lsMovNro
                lbResultadoVisto = loVistoElectronico.Inicio(10, 120205)
                If Not lbResultadoVisto Then
                    MsgBox "Operación cancelada por el usuario", vbInformation, "Aviso"
                    Exit Sub
                End If
            End If
            'RECO FIN*************************************************************************************
            loRegPig.dBeginTrans
            Call loGrabarDesem.nDesembolsoCredPignoEfectivoAbono(Me.AXCodCta.NroCuenta, lbDesembCC, lbCuentaNueva, lnSaldoCap, lsFechaHoraGrab, _
                 lsMovNro, lnMontoEntregar, 0, lnImpuesto, lnCostoTasacion, lnCostoCustodia, IIf(lbDesembCC, False, gITF.gbITFAplica), gITF.gbITFAsumidocreditos, loItf1, loItf2, CCur(Me.lblSaldoNeto.Caption), lnMovNro, False, loRegPig, loRegPig.cnConecOpenValor)
            '*** BRGO 20110914 *********************
            If gITF.gbITFAplica Then
                Call loGrabarDesem.InsertaMovRedondeoITF(lsMovNro, 1, CCur(Me.lblITF.Caption) + nRedondeoITF, CCur(Me.lblITF.Caption), lnMovNro, loRegPig, loRegPig.cnConecOpenValor)
                'Set oMov = New COMDMov.DCOMMov
                'Call oMov.InsertaMovRedondeoITF(lsMovNro, 1, CCur(Me.lblITF.Caption) + nRedondeoITF, CCur(Me.lblITF.Caption))
                'Set oMov = Nothing
            End If
            '*** END BRGO *******************************
'*******REALIZA LA CANCELACION ************************************************************
            Call loGrabarDesem.nCancelacionCredPignoraticio(Me.ActXCodCtaCancelar.NroCuenta, vNewSaldoCapitalCanc, lsFechaHoraGrab, _
                     lsMovNro, lnMontoTransaccion, vSaldoCapitalCanc, vInteresVencidoCanc, _
                      vCostoCustodiaMoratorioCanc, vCostoPreparacionRemateCanc, vImpuestoCanc, _
                      vDiasAtrasoRealCanc, fnVarDiasCambCartCanc, vValorTasacionCanc, fnVarOpeCodCanc, fsVarOpeDescCanc, fsVarPersCodCMACCanc, _
                      gITF.gbITFAplica, gITF.gbITFAsumidocreditos, CCur(Me.lblITF), False, Val(fnVarGastoCorrespondenciaCanc), vInteresMoratorioCanc, fnVarCostoNotificacionCanc, _
                      vInteresAdelCanc, lnMovNro, loRegPig, loRegPig.cnConecOpenValor)
'*******FIN DE LA CANCELACION **************************************************************
'*******REALIZA RESCATE JOYA ***************************************************************
            Call loGrabarDesem.nRescataJoyaCredPignoraticio(Me.ActXCodCtaCancelar.NroCuenta, lsFechaHoraGrab, _
                     lsMovNro, Val(Me.AXDesCon.Oro14), Val(Me.AXDesCon.Oro16), _
                     Val(Me.AXDesCon.Oro18), Val(Me.AXDesCon.Oro21), Val(nDiasTranscurridosResc), vValorTasacionCanc, gColPOpeDevJoyas, False, lnMovNro, loRegPig, loRegPig.cnConecOpenValor)
'*******FIN RESCATE*************************************************************************
''********** aqui graba el abono en cta
    If Me.chkDesembAbonoCta.value = 1 Then
        If lbDesembCC Then
    
            ReDim MatDatosAhoNew(4)
    
            MatDatosAhoNew(0) = nProgAhorros
            MatDatosAhoNew(1) = nMontoAbonar
            MatDatosAhoNew(2) = nPlazoAbonar
            MatDatosAhoNew(3) = sPromotorAho
                
            Set oITF = New COMDConstSistema.FCOMITF
            oITF.fgITFParametros
            oITF.fgITFParamAsume Mid(lsCuenta, 4, 2)
        
            lnValorItfAbono = CDbl(lblITF.Caption)
            'lnValorItfGasto = oITF.fgITFDesembolso(0) 'pnMontoGastos
            'lnValorItfCancelacion = oITF.fgITFDesembolso(0) 'pnMontoCancel
            
            Dim pMatGastos() As String
            Dim pMatCargoAutom() As String
            
            If lbCuentaNueva Then
                'Call loGrabarDesem.DesembolsoPignoConAbonoCta(lnMovNro, MatDatosAho, lsCuenta, lnSaldoCap, _
                    lnMontoEntregar, pMatGastos, pMatCargoAutom, gdFecSis, Mid(lsCuenta, 4, 2), gsCodUser, True, True, , , Me.ActXCodCtaAbono.NroCuenta, _
                     pRSRela, pnTasa, pnPersoneria, pnTipoCuenta, pnTipoTasa, pbDocumento, psNroDoc, psCodIF, , , , , _
                    oITF.gbITFAplica, lnValorItfAbono, , , False, gITFCobroCargoPigno, , , , , , psCtaAhoNew, MatDatosAhoNew, , , lsMovNro)
            Else
                
                'Call loGrabarDesem.DesembolsoPignoConAbonoCta(lnMovNro, MatDatosAho, lsCuenta, lnSaldoCap, _
                    lnMontoEntregar, pMatGastos, pMatCargoAutom, gdFecSis, Mid(lsCuenta, 4, 2), gsCodUser, True, False, , , sCtaAho, , , , , , , , , _
                    , , , , _
                    oITF.gbITFAplica, lnValorItfAbono, , , False, gITFCobroCargoPigno)
    
                Call loGrabarDesem.DesembolsoPignoConAbonoCta(lnMovNro, MatDatosAho, lsCuenta, lnMontoEntregar, _
                    lnMontoEntregar, pMatGastos, pMatCargoAutom, gdFecSis, Mid(lsCuenta, 4, 2), gsCodUser, True, False, , , Me.ActXCodCtaAbono.NroCuenta, , , , , , , , , _
                    , , , , _
                    oITF.gbITFAplica, lnValorItfAbono, , , False, gITFCobroCargoPigno, , , , , , , , , , , loRegPig, loRegPig.cnConecOpenValor)
            End If
            
        End If
    End If
    loRegPig.dCommitTrans
    Set loRegPig = Nothing
    Set oITF = Nothing
            
'********************************************************
'RECO COMENTADO
'        If Trim(Me.LblTipoAbono.Caption) = "NUEVA" Then
'           'IMPRIME REGISTRO DE FISMAS
'           Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
'           Dim lsCadImpFirmas As String
'           Dim lsCadImpCartilla As String
'           Dim sTipoCuenta As String
'           If pnTipoCuenta = 0 Then
'                sTipoCuenta = "INDIVIDUAL"
'           ElseIf pnTipoCuenta = 1 Then
'                sTipoCuenta = "MALCOMUNADA"
'           ElseIf pnTipoCuenta = 2 Then
'                sTipoCuenta = "INDISTINTA"
'           End If
'           Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
'                clsMant.IniciaImpresora gImpresora
'                lsCadImpFirmas = clsMant.GeneraRegistroFirmas(psCtaAhoNew, sTipoCuenta, gdFecSis, False, pRSRela, gsNomAge, gdFecSis, gsCodUser)
'           Set clsMant = Nothing
'           Set rsRel = Nothing
'
'           'IMPRIME CARTILLA EN WORD
'           Dim lnTasaE As Double
'           lnTasaE = Round(((1 + (pnTasa / 100 / 12) / 30) ^ 360 - 1) * 100, 2)
'            ImpreCartillaAhoCorriente MatTitulares, psCtaAhoNew, lnTasaE, lnSaldoCap, nProgAhorros
'
'        End If
'FIN RECO
'********************************************************
        '*** IMPRIME REGISTRO DE FIRMAS - AHORRO
        'If Trim(LblTipoAbono.Caption) = "NUEVA" Then
        '    MsgBox "Coloque Papel Continuo Tamaño Carta, Para la Impresion del Registros de Firmas", vbInformation, "Aviso"
        '    clsprevio.PrintSpool sLpt, oImpresora.gPrnCondensadaON & lsCadImpFirmas & oImpresora.gPrnCondensadaOFF, False, gnLinPage   'ARCV 01-11-2006
        'End If

'********************************************************

        
        lsCadImp = loColImp.nPrintReciboDesembolso(vFecVenc, Me.AXCodCta.NroCuenta, lnSaldoCap, lsFechaHoraPrend, _
                       lnMontoEntregar, lnInteresComp, gsNomAge, gsCodUser, CDbl(lblITF.Caption), gImpresora, lbDesembCC, Me.ActXCodCtaAbono.NroCuenta, psCtaAhoNew)
                       
        

'*********************************************************
        lcTextImp = "Desea Imprimir las boletas de desembolsos"
        Do
        OptBt2 = MsgBox(lcTextImp, vbInformation + vbYesNo, "Aviso")
        If vbYes = OptBt2 Then
        lcTextImp = "Desea Reimprimir las boletas de desembolsos"
        MsgBox "Cambie de Papel para imprimir las boletas de desembolsos", vbExclamation, "Aviso"
        nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
                Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
                Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
                Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
                Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
                Print #nFicSal, lsCadImp & Chr$(12)
                Print #nFicSal, ""
                Close #nFicSal
        End If
        Loop Until OptBt2 = vbNo
            
'*********************************************************
         lsCadImp = loColImp.nPrintReciboCancelacion(gsNomAge, lsFechaHoraGrab, Me.ActXCodCtaCancelar.NroCuenta, AXDesCon.listaClientes.ListItems(1).SubItems(1), _
                    Format(AXDesCon.FechaPrestamo, "mm/dd/yyyy"), vdiasAtrasoCanc, CCur(AXDesCon.prestamo), vSaldoCapitalCanc, _
                    vInteresAdelCanc, vInteresVencidoCanc, vImpuestoCanc, vCostoCustodiaMoratorioCanc, vCostoPreparacionRemateCanc, _
                    lnMontoTransaccion, vNewSaldoCapitalCanc, vTasaInteresCanc, gsCodUser, fnVarCostoNotificacionCanc, fsVarNombreCMACCanc, "", Val(fnVarGastoCorrespondenciaCanc), _
                    CDbl(lblITF.Caption), gImpresora, vInteresMoratorioCanc, gbImpTMU)
            
            'Set loPrevio = New previo.clsprevio
                clsprevio.PrintSpool sLpt, lsCadImp, False, 22
                Do While True
                    If MsgBox("Reimprimir Recibo de Cancelación ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                        clsprevio.PrintSpool sLpt, lsCadImp, False, 22
                    Else
                        'Set loPrevio = Nothing
                        Exit Do
                    End If
                Loop
'*********************************************************
'*********************************************************
'    If MsgBox("Desea realizar impresiones ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'        'Set loPrevio = New previo.clsprevio
'            lsCadImp = loColImp.ImprimirRescate(Me.ActXCodCtaCancelar.NroCuenta, AXDesCon.listaClientes.ListItems(1).SubItems(1), AXDesCon.prestamo, AXDesCon.SaldoCapital, gdFecSis, gsCodUser, gImpresora)
'                clsprevio.PrintSpool sLpt, lsCadImp, False, 22
'                Do While True
'                    If MsgBox("Desea reimprimir ?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'                         clsprevio.PrintSpool sLpt, lsCadImp, False, 22
'                    Else
'                    Exit Do
'                    'Set loPrevio = Nothing
'                End If
'            Loop
'        'Set loPrevio = Nothing
'    End If
'*********************************************************
        Set loGrabarDesem = Nothing
        Set loColImp = Nothing
        Limpiar
        'Me.lblNetoRecibir = "0.00"
        'Me.lblITF = "0.00"
        'Me.LblTotalPagar = "0.00"

        AXCodCta.Enabled = True
        AXCodCta.SetFocus
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
    loRegPig.dRollbackTrans
End Sub
Private Function IniciaLavDinero() As String
    Dim i As Long
    Dim nRelacion As CaptacRelacPersona
    Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
    Dim nMonto As Double, nPersoneria As Integer
    Dim sCuenta As String
        nPersoneria = gPersonaNat
        If nPersoneria = gPersonaNat Then
                sPersCod = AXDesCon.listaClientes.ListItems(1).Text
                sNombre = AXDesCon.listaClientes.ListItems(1).SubItems(1)
                sDireccion = AXDesCon.listaClientes.ListItems(1).SubItems(2)
                sDocId = AXDesCon.listaClientes.ListItems(1).SubItems(7)
        Else
                sPersCod = AXDesCon.listaClientes.ListItems(1).Text
                sNombre = AXDesCon.listaClientes.ListItems(1).SubItems(1)
                sDireccion = AXDesCon.listaClientes.ListItems(1).SubItems(2)
                sDocId = AXDesCon.listaClientes.ListItems(1).SubItems(9)
        End If
    
    nMonto = CDbl(Me.lblSaldoNeto.Caption)
    sCuenta = AXCodCta.NroCuenta
    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, sCuenta, gColPOpeDesembolsoEFE, , gMonedaNacional)
End Function

'Finaliza el formulario actual
Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
            SendKeys "{Enter}"
        End If
    ElseIf KeyCode = 13 And Trim(AXCodCta.EnabledCta) And AXCodCta.Age <> "" And Trim(AXCodCta.Cuenta) = "" Then
                AXCodCta.SetFocusCuenta
                 Exit Sub
        
    End If
End Sub

'Inicializa el formulario actual
Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    fsVarOpeDesc = "DESEMBOLSO AMPLIACION"
    fnVarOpeCod = "120205"
    fsVarOpeDescCanc = "Cancelacion Pignoraticio"
    fnVarOpeCodCanc = "121200"
    fnVarOpeCodResc = "121800"
    fsVarOpeDescResc = "DEVOLUCION DE JOYAS"
    Limpiar
End Sub


Private Sub IniciaLavDineroCancelacion(poLavDinero As frmMovLavDinero)
Dim i As Long
Dim nRelacion As CaptacRelacPersona
Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
Dim nMonto As Double, nPersoneria As Integer
Dim sCuenta As String
'For i = 1 To grdCliente.Rows - 1
    'nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4)))
    nPersoneria = gPersonaNat
    If nPersoneria = gPersonaNat Then
        'If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
            poLavDinero.TitPersLavDinero = AXDesCon.listaClientes.ListItems(1).Text
            poLavDinero.TitPersLavDineroNom = AXDesCon.listaClientes.ListItems(1).SubItems(1)
            poLavDinero.TitPersLavDineroDir = AXDesCon.listaClientes.ListItems(1).SubItems(2)
            poLavDinero.TitPersLavDineroDoc = AXDesCon.listaClientes.ListItems(1).SubItems(7)
         '   Exit For
       ' End If
    Else
        'If nRelacion = gCapRelPersTitular Then
            poLavDinero.TitPersLavDinero = AXDesCon.listaClientes.ListItems(1).Text
            poLavDinero.TitPersLavDineroNom = AXDesCon.listaClientes.ListItems(1).SubItems(1)
            poLavDinero.TitPersLavDineroDir = AXDesCon.listaClientes.ListItems(1).SubItems(2)
            poLavDinero.TitPersLavDineroDoc = AXDesCon.listaClientes.ListItems(1).SubItems(9)
          '  Exit For
        'End If
    End If
'Next i
nMonto = CDbl((Me.lblMontoCredant.Caption))
sCuenta = AXCodCta.NroCuenta
'If sPersCodCMAC <> "" Then
'    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nmonto, sCuenta, sOperacion, , sTipoCuenta)
'Else
    'IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, sCuenta, CStr(fnVarOpeCod), , gMonedaNacional)
'End If
End Sub
Private Sub fgCalculaDeuda()
Dim loCalculos As COMNColoCPig.NCOMColPCalculos 'NColPCalculos
Dim lsmensaje As String
vdiasAtrasoCanc = DateDiff("d", Format(vFecVencimientoCanc, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
vDiasAtrasoRealCanc = vdiasAtrasoCanc
If vdiasAtrasoCanc <= 0 Then
        
        If gcCredAntiguoCanc = "A" Then
            vInteresAdelCanc = Round(0, 2)
        Else
        Set loCalculos = New COMNColoCPig.NCOMColPCalculos
            vDiasAdelCanc = DateDiff("d", Format(vFecEstadoCanc, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
            vInteresAdelCanc = loCalculos.nCalculaInteresAlVencimiento(vSaldoCapitalCanc, vTasaInteresCanc, vDiasAdelCanc)
            vInteresAdelCanc = Round(vInteresAdelCanc, 2)
        Set loCalculos = Nothing
        End If
    
    'end peac
    vdiasAtrasoCanc = 0
    vInteresVencidoCanc = 0
    vInteresMoratorioCanc = 0
    vCostoCustodiaMoratorioCanc = 0
    vImpuestoCanc = 0
    fnVarGastoCorrespondenciaCanc = 0
Else
    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
        
        If gcCredAntiguoCanc = "A" Then
            vInteresAdelCanc = Round(vInteresAdelCanc, 2)
        Else
            vDiasAdelCanc = DateDiff("d", Format(vFecEstadoCanc, "dd/mm/yyyy"), Format(vFecVencimientoCanc, "dd/mm/yyyy"))
             vInteresAdelCanc = loCalculos.nCalculaInteresAlVencimiento(vSaldoCapitalCanc, vTasaInteresCanc, vDiasAdelCanc)
            vInteresAdelCanc = Round(vInteresAdelCanc, 2)
        End If
    
        vInteresVencidoCanc = loCalculos.nCalculaInteresMoratorio(vSaldoCapitalCanc, vTasaInteresVencidoCanc, vdiasAtrasoCanc)
        vInteresVencidoCanc = Round(vInteresVencidoCanc, 2)
        
        vInteresMoratorioCanc = loCalculos.nCalculaInteresMoratorio(vSaldoCapitalCanc, vTasaInteresMoratorioCanc, vdiasAtrasoCanc)
        vInteresMoratorioCanc = Round(vInteresMoratorioCanc, 2)
        
        vCostoCustodiaMoratorioCanc = loCalculos.nCalculaCostoCustodiaMoratorio(vValorTasacionCanc, fnTasaCustodiaVencidaCanc, vdiasAtrasoCanc)
        vCostoCustodiaMoratorioCanc = Round(vCostoCustodiaMoratorioCanc, 2)
        
        vImpuestoCanc = (vInteresVencidoCanc + vCostoCustodiaMoratorioCanc + vInteresMoratorioCanc) * fnTasaImpuestoCanc
        vImpuestoCanc = Round(vImpuestoCanc, 2)
        fnVarGastoCorrespondenciaCanc = loCalculos.nCalculaGastosCorrespondencia(AXCodCta.NroCuenta, lsmensaje)
        
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
    Set loCalculos = Nothing
End If
vCostoPreparacionRemateCanc = 0

If vEstadoCanc = gColPEstPRema And fnVarEstUltProcRemCanc = 2 Then  ' Si esta en via de Remate
    vCostoPreparacionRemateCanc = fnTasaPreparacionRemateCanc * vValorTasacionCanc
    vCostoPreparacionRemateCanc = Round(vCostoPreparacionRemateCanc, 2)
End If

If gnNotifiAdjuCanc = 1 Then
    If gnNotifiCobCanc = 1 Then
        fnVarCostoNotificacionCanc = 0
    End If
Else
    fnVarCostoNotificacionCanc = 0
End If

 vDeudaCanc = vSaldoCapitalCanc + vInteresAdelCanc + vInteresVencidoCanc + vCostoCustodiaMoratorioCanc + vImpuestoCanc + vCostoPreparacionRemateCanc + fnVarGastoCorrespondenciaCanc + vInteresMoratorioCanc + fnVarCostoNotificacionCanc

End Sub

Private Sub BuscaContratoCancelar(ByVal psNroContrato As String)
Dim lbok As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida 'nColPValida
Dim loCalculos As COMNColoCPig.NCOMColPCalculos 'NColPCalculos
Dim loPigFunc As COMDColocPig.DCOMColPFunciones 'dColPFunciones
Dim lnDeuda As Currency, lnMinimoPagar As Currency
Dim lnDiasAtraso  As Integer
Dim lsFecVenTemp As String
Dim lsmensaje As String

Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim Rf As ADODB.Recordset
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
    
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        Set lrValida = loValContrato.nValidaCancelacionCredPignoraticio(psNroContrato, gdFecSis, 0, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Call Limpiar
             Exit Sub
        End If
        
    Set loValContrato = Nothing
    
    If lrValida Is Nothing Then ' Hubo un Error
        Limpiar
        Set lrValida = Nothing
        Exit Sub
    End If
    
    vValorTasacionCanc = Format(lrValida!nTasacion, "#0.00")
    vTasaInteresVencidoCanc = lrValida!nTasaIntVenc
    vTasaInteresMoratorioCanc = lrValida!nTasaIntMora
    vEstadoCanc = lrValida!nPrdEstado
    
    gcCredAntiguoCanc = lrValida!cCredB
    
    gnNotifiAdjuCanc = lrValida!nCodNotifiAdj
    gnNotifiCobCanc = lrValida!nCodNotifiCob
    
    vFecEstadoCanc = lrValida!dPrdEstado
    vSaldoCapitalCanc = lrValida!nMontoCol
    
    vSaldoCapitalCanc = Format(lrValida!nSaldo, "#0.00")
    vTasaInteresCanc = lrValida!nTasaInteres
    
    vFecVencimientoCanc = Format(lrValida!dVenc, "dd/mm/yyyy")
    
    fnVarEstUltProcRemCanc = lrValida!nEstUltProcRem
    
    If fgMuestraCredPig_AXDesCon(Me.AXCodCta.NroCuenta, Me.AXDesCon, False) Then

    End If
    
    
    lsFecVenTemp = vFecVencimientoCanc
    Set loPigFunc = New COMDColocPig.DCOMColPFunciones
    
    If loPigFunc.dVerSiFeriado(lsFecVenTemp, lsmensaje) = True Then
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
        Do While True
            lsFecVenTemp = DateAdd("d", 1, lsFecVenTemp)
            If Not loPigFunc.dVerSiFeriado(lsFecVenTemp, lsmensaje) = True Then
                If Trim(lsmensaje) <> "" Then
                    MsgBox lsmensaje, vbInformation, "Aviso"
                    Exit Sub
                End If
                Exit Do
            End If
        Loop
        If lsFecVenTemp = gdFecSis Then
            vFecVencimientoCanc = lsFecVenTemp
        End If
    End If
    Set loPigFunc = Nothing
    
    lnDiasAtraso = DateDiff("d", Format(lrValida!dVenc, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
    
    'Me.txtDiasAtraso = val(lnDiasAtraso) -RECO-B

    If Me.AXCodCta.Age <> "" Then
        Select Case CInt(Me.AXCodCta.Age)
            Case 1
               fnVarCostoNotificacionCanc = loParam.dObtieneColocParametro(3103)
            Case 2
               fnVarCostoNotificacionCanc = loParam.dObtieneColocParametro(3104)
            Case 3
               fnVarCostoNotificacionCanc = loParam.dObtieneColocParametro(3105)
            Case 4
               fnVarCostoNotificacionCanc = loParam.dObtieneColocParametro(3106)
            Case 5
               fnVarCostoNotificacionCanc = loParam.dObtieneColocParametro(3107)
            Case 6
               fnVarCostoNotificacionCanc = loParam.dObtieneColocParametro(3108)
            Case 7
               fnVarCostoNotificacionCanc = loParam.dObtieneColocParametro(3109)
            Case 9
               fnVarCostoNotificacionCanc = loParam.dObtieneColocParametro(3111)
            Case 10
               fnVarCostoNotificacionCanc = loParam.dObtieneColocParametro(3112)
            Case 12
               fnVarCostoNotificacionCanc = loParam.dObtieneColocParametro(3113)
            Case 13
               fnVarCostoNotificacionCanc = loParam.dObtieneColocParametro(3114)
            Case 24
               fnVarCostoNotificacionCanc = loParam.dObtieneColocParametro(3115)
            Case 25
               fnVarCostoNotificacionCanc = loParam.dObtieneColocParametro(3116)
            Case 31
               fnVarCostoNotificacionCanc = loParam.dObtieneColocParametro(3117)
        End Select
   End If
                Set loParam = Nothing
    fgCalculaDeuda
    
    'lblTotalDeuda.Caption = Format(CDbl(vDeudaCanc), "#0.00") -RECO-B
    If gITF.gbITFAplica Then
        If Not gITF.gbITFAsumidocreditos Then
            'lblITF.Caption = Format(gITF.fgITFCalculaImpuesto(lblTotalDeuda.Caption), "#0.00") -RECO-B
            'nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption)) -RECO-B
            'If nRedondeoITF > 0 Then -RECO-B
               'Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00") -RECO-B
            'End If -RECO-B
            'LblMontoPagar = Format(CDbl(Me.lblTotalDeuda.Caption) + CDbl(lblITF.Caption), "#0.00") -RECO-B
        Else
            'lblITF = Format(gITF.fgITFCalculaImpuesto(LblMontoPagar.Caption), "#0.00") -RECO-B
            'nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
            'If nRedondeoITF > 0 Then
               'Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
            'End If
            'LblMontoPagar = Format(CDbl(Me.lblTotalDeuda.Caption), "#0.00")
        End If
    Else
            'lblITF = Format(0, "#0.00")
            'LblMontoPagar = Format(CDbl(Me.lblTotalDeuda.Caption), "#0.00")
    End If
'txtCapital.Text = Format(vSaldoCapital, "#0.00")
'txtMora.Text = Format(vInteresMoratorio, "#0.00")
'txtInteres.Text = Format(vInteresAdel, "#0.00")
'txtIntVen.Text = Format(vInteresVencido, "#0.00")
'txtCostoCus.Text = Format(vCostoCustodiaMoratorio, "#0.00")
'txtCostoNoti.Text = Format(fnVarCostoNotificacion, "#0.00")
'txtCostoRemate.Text = Format(vCostoPreparacionRemate, "#0.00")
    
    Set lrValida = Nothing


        
         Set lafirma = New frmPersonaFirma
         Set ClsPersona = New COMDPersona.DCOMPersonas
        
         Set Rf = ClsPersona.BuscaCliente(gColPigFunciones.vcodper, BusquedaCodigo)
         
         If Not Rf.BOF And Not Rf.EOF Then
            lsPersCod = Rf!cPersCod
            If Rf!nPersPersoneria = 1 Then
            Call frmPersonaFirma.Inicio(Trim(gColPigFunciones.vcodper), Mid(gColPigFunciones.vcodper, 4, 2), False, False) 'MOD BY JATO 20210324 True --> False
        End If
         End If
         Set Rf = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Public Sub BuscaContratoRescate(ByVal psNroContrato As String)

Dim lbok As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim loCalculos As COMNColoCPig.NCOMColPCalculos
Dim lnCustodiaDiferida  As Currency
Dim lsmensaje As String
    
    Set lrValida = New ADODB.Recordset
        Set loValContrato = New COMNColoCPig.NCOMColPValida
            Set lrValida = loValContrato.nValidaRescateCredPignoraticio(psNroContrato, gsCodAge, gdFecSis, fnVarOpeCod, gsCodUser, lsmensaje)
            If Trim(lsmensaje) <> "" Then
                 MsgBox lsmensaje, vbInformation, "Aviso"
                 Call Limpiar
                 Exit Sub
            End If
            
        Set loValContrato = Nothing
    
        If lrValida Is Nothing Then ' Hubo un Error
            Limpiar
            Set lrValida = Nothing
            Exit Sub
        End If
        'reco quitado
        'lbok = fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, True)
        
        
        Set loCalculos = New COMNColoCPig.NCOMColPCalculos
            lnCustodiaDiferida = loCalculos.nCalculaCostoCustodiaDiferida(lrValida!nTasacion, IIf(IsNull(lrValida!nDiasTranscurridos), 0, lrValida!nDiasTranscurridos), lrValida!nPorcentajeCustodia, lrValida!nTasaIGV)
        Set loCalculos = Nothing
        'reco quitado
        'Me.lblCostoCustodia = Format(lnCustodiaDiferida - lrValida!nCustodiaDiferida, "#0.00")
        'Me.lblFecPago = Format(lrValida!dCancelado, "dd/mm/yyyy")
        nDiasTranscurridosResc = IIf(IsNull(lrValida!nDiasTranscurridos), 0, lrValida!nDiasTranscurridos)
        'Me.lblNroDuplic = lrValida!nNroDuplic
    Set lrValida = Nothing
        
    'AXCodCta.Enabled = False
    
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
