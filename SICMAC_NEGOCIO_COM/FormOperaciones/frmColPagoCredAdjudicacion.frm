VERSION 5.00
Begin VB.Form frmColPagoCredAdjudicacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pago de Credito x Adjudicacion"
   ClientHeight    =   3690
   ClientLeft      =   5490
   ClientTop       =   4935
   ClientWidth     =   6585
   Icon            =   "frmColPagoCredAdjudicacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   6375
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   5040
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Nro Credito"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
         CMAC            =   "109"
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Pagar:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblITF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ITF:"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblTpoCambio 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Cambio:"
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label txtTpoCambio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4320
         TabIndex        =   14
         Top             =   2520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblMontoSoles 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Soles:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label txtMontoSoles 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   2520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label txtMoneda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4320
         TabIndex        =   10
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Monto:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblPersCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Titular Cta:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lbltitular 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   1080
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmColPagoCredAdjudicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnMoneda As Integer

Dim lnNroCalen As Integer
Dim lnNroCuota As Integer
Dim lnDiasAtraso As Integer
Dim lsMetLiqui As String
Dim lnPlazo As Integer
Dim lnPrdEstado As Integer
Dim lnCapital As Currency
Dim lnIntComp As Currency
Dim lnIntMora As Currency
Dim lnGastos As Currency

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObtieneDatosCuenta AXCodCta.NroCuenta
        'Me.cmdAceptar.SetFocus
    End If
End Sub

Private Sub cmdAceptar_Click()
    'ALPA20131001
    Dim oTrans As Boolean
    On Error GoTo ErrorAdjudicacion
        oTrans = False
        If MsgBox("Se van ha Guardar los Datos", vbYesNo, "AVISO") = vbYes Then
            Dim oBase As COMDCredito.DCOMCredActBD
            Dim clsMant As COMDCaptaGenerales.DCOMCaptaMovimiento
            Dim ClsMov As COMNContabilidad.NCOMContFunciones
            Dim oColRec As COMNColocRec.NCOMColRecCredito
            
            Dim sMovNro As String
            Dim nMovNro As Long
            
            Set oBase = New COMDCredito.DCOMCredActBD
            Set clsMant = New COMDCaptaGenerales.DCOMCaptaMovimiento
            Set ClsMov = New COMNContabilidad.NCOMContFunciones
            Set oColRec = New COMNColocRec.NCOMColRecCredito
            
            sMovNro = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            oTrans = True
            oBase.dBeginTrans
                clsMant.AgregaMov sMovNro, 130206, "PAGO CREDITO X ADJUDICACION", gMovEstContabPendiente, gMovFlagVigente
                nMovNro = clsMant.GetnMovNro(sMovNro)
                'lnCapital
                Call oBase.dInsertMovCol(nMovNro, "130206", Me.AXCodCta.NroCuenta, lnNroCalen, CCur(Me.txtMonto), lnDiasAtraso, lsMetLiqui, lnPlazo, lnCapital, lnPrdEstado, False)
                oColRec.guardarMovPendientesRend nMovNro, IIf(Mid(Me.AXCodCta.NroCuenta, 9, 1) = "1", "19180711", "19280711"), CCur(Me.txtMonto)
                'Call oBase.dInsertMovColDet(nMovNro, "130206", Me.AXCodCta.NroCuenta, CLng(lnNroCalen), 3000, lnNroCuota, CCur(Me.txtMonto), False)
                Call oBase.dInsertMovColDet(nMovNro, "130206", Me.AXCodCta.NroCuenta, CLng(lnNroCalen), 3000, lnNroCuota, lnCapital, False)
                Call oBase.dInsertMovColDet(nMovNro, "130206", Me.AXCodCta.NroCuenta, CLng(lnNroCalen), 3100, lnNroCuota, lnIntComp, False)
                Call oBase.dInsertMovColDet(nMovNro, "130206", Me.AXCodCta.NroCuenta, CLng(lnNroCalen), 3101, lnNroCuota, lnIntMora, False)
                Call oBase.dInsertMovColDet(nMovNro, "130206", Me.AXCodCta.NroCuenta, CLng(lnNroCalen), 3201, lnNroCuota, lnGastos, False)
                oColRec.guardarPagoCredAdjudicacion "130206", Me.AXCodCta.NroCuenta, CCur(Me.txtMonto), gdFecSis, lnMoneda, sMovNro
            oBase.dCommitTrans
            'MsgBox "Se han Guardado los Datos", vbInformation, "AVISO"
            MsgBox "Coloque papel para la Boleta", vbInformation, "Aviso"
            Dim loImprime As COMNColocRec.NCOMColRecImpre
            Dim loPrevio As previo.clsprevio
            Dim sCadImprimir As String
            Set loImprime = New COMNColocRec.NCOMColRecImpre
                
            sCadImprimir = loImprime.nPrintReciboPagoCredRecup(gsNomAge, gdFecSis, AXCodCta.NroCuenta, _
            Me.lblTitular, CCur(Me.txtMonto), gsCodUser, "OPE: 130206-Pago de Credito x Adjudicacion", Me.lblITF, gImpresora, gbImpTMU)
            'Me.lbltitular, CCur(Me.txtMonto), gsCodUser, "OPE: 130206-Pago de Credito x Adjudicacion", CDbl(LblITF.Caption), gImpresora, gbImpTMU)
            Set loImprime = Nothing
            Set loPrevio = New previo.clsprevio
                loPrevio.PrintSpool sLpt, sCadImprimir, True, 22
                
                Do While True
                    If MsgBox("Reimprimir Recibo de Pago de Credito x Adjudicacion ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                        loPrevio.PrintSpool sLpt, sCadImprimir, True, 22
                    Else
                        Set loPrevio = Nothing
                        Exit Do
                    End If
                Loop
            LimpiarPantalla
            'INICIO JHCU ENCUESTA 16-10-2019
            Encuestas gsCodUser, gsCodAge, "ERS0292019", "130206"
            'FIN
        End If
        Exit Sub
'ALPA20131001
ErrorAdjudicacion:
    If oTrans = True Then
        oBase.dRollbackTrans
        oTrans = False
        err.Raise err.Number, "Error En Proceso AmortizarPagoLote", err.Description
    End If
End Sub

Private Sub LimpiarPantalla()
    Me.AXCodCta.NroCuenta = "109"
    Me.LblPersCod = ""
    Me.lblTitular = ""
    Me.txtMonto = ""
    Me.txtMoneda = ""
    
    Me.txtMontoSoles = ""
    Me.txtTpoCambio = ""
    
    Me.lblMontoSoles.Visible = True
    Me.txtMontoSoles.Visible = True
    Me.lblTpoCambio.Visible = True
    Me.txtTpoCambio.Visible = True
    Me.cmdAceptar.Enabled = False
    
End Sub

Private Sub cmdBuscar_Click()
    Dim loPers As COMDPersona.UCOMPersona 'UPersona
    Dim lsPersCod As String, lsPersNombre As String
    Dim lsEstados As String
    Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
    Dim lrCreditos As New ADODB.Recordset
    Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

    Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.inicio
    
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
    Set loPers = Nothing

    ' Selecciona Estados
    lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast

    If Trim(lsPersCod) <> "" Then
        Set loPersCredito = New COMDColocRec.DCOMColRecCredito
            Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)
        Set loPersCredito = Nothing
    End If

    Set loCuentas = New COMDPersona.UCOMProdPersona
        Set loCuentas = frmProdPersona.inicio(lsPersNombre, lrCreditos)
        If loCuentas.sCtaCod <> "" Then
            AXCodCta.Enabled = True
            AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
            AXCodCta.SetFocusCuenta
        End If
    Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdsalir_Click()
 Unload Me
End Sub
Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
    Dim oColRec As COMNColocRec.NCOMColRecCredito
    Dim rsCta As ADODB.Recordset
    Dim nCta As String
    Dim nITF As Double
    
    Set oColRec = New COMNColocRec.NCOMColRecCredito
    Set rsCta = New ADODB.Recordset
    
    Set rsCta = oColRec.ObtenerPagoCredAdjudicacion(sCuenta, "136303")

    If Not (rsCta.EOF And rsCta.BOF) Then
        
        LblPersCod.Caption = rsCta("cPersCod")
        lblTitular.Caption = UCase(PstaNombre(rsCta("Nombre")))
        
        Me.txtMonto = Format(rsCta("Monto"), "##,##0.00")
'        lnCapital = rsCta("Capital")
'        Me.txtMonto = Format(lnCapital, "##,##0.00")
        
        nITF = gITF.fgITFCalculaImpuesto(rsCta("Monto"))
        'nITF = gITF.fgITFCalculaImpuesto(lnCapital)
        lblITF = Format(nITF, "##,##0.00")
        
        Me.lblTotal = Format(rsCta("Monto") + nITF, "##,##0.00")
         'Me.lblTotal = Format(lnCapital + nITF, "##,##0.00")
        
        lnCapital = rsCta("Capital")
        lnIntComp = rsCta("IntComp")
        lnIntMora = rsCta("IntMora")
        lnGastos = rsCta("Gastos")
        lnNroCalen = rsCta("nNroCalen")
        lnNroCuota = rsCta("nNroProxCuota")
        lnDiasAtraso = rsCta("nDiasAtraso")
        lsMetLiqui = rsCta("cMetLiquidacion")
        lnPlazo = rsCta("nPlazo")
        lnPrdEstado = rsCta("nPrdEstado")
        lnMoneda = rsCta("Moneda")
        
        If lnMoneda = 1 Then
            Me.txtMoneda = "NACIONAL"
            
            Me.lblMontoSoles.Visible = False
            Me.txtMontoSoles.Visible = False
            Me.lblTpoCambio.Visible = False
            Me.txtTpoCambio.Visible = False
            
            Me.txtMontoSoles = ""
            Me.txtTpoCambio = ""
        Else
            Dim clsTC As COMDConstSistema.NCOMTipoCambio
            Dim nTC As Double
            Set clsTC = New COMDConstSistema.NCOMTipoCambio
            nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
            Me.txtMoneda = "EXTRANJERA"
            Me.lblMontoSoles.Visible = True
            Me.txtMontoSoles.Visible = True
            Me.lblTpoCambio.Visible = True
            Me.txtTpoCambio.Visible = True
            
            Me.txtMontoSoles = Format(CCur(lblTotal) * nTC, "##,##0.00")
            Me.txtTpoCambio = nTC
        End If
        Me.cmdAceptar.Visible = True
         Me.cmdAceptar.Enabled = True
    Else
        MsgBox "No se ha encontrado información de la cuenta ingresada"
        AXCodCta.SetFocus
    End If
End Sub
