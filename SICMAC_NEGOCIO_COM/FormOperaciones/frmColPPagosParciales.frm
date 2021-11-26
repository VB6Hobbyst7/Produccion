VERSION 5.00
Begin VB.Form frmColPPagosParciales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Amortización de Crédito"
   ClientHeight    =   8760
   ClientLeft      =   1935
   ClientTop       =   2385
   ClientWidth     =   7965
   Icon            =   "frmColPPagosParciales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   8300
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4440
      TabIndex        =   3
      Top             =   8300
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6795
      TabIndex        =   5
      Top             =   8300
      Width           =   975
   End
   Begin VB.Frame fraContenedor 
      Height          =   7590
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   7785
      Begin VB.TextBox txtCampReteAmor 
         Height          =   285
         Left            =   5520
         TabIndex        =   54
         Text            =   "0.00"
         Top             =   300
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   7110
         Picture         =   "frmColPPagosParciales.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Buscar ..."
         Top             =   270
         Width           =   420
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Plazo Nuevo"
         Height          =   1785
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   5760
         Width           =   7545
         Begin VB.TextBox txtMora 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   43
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtIntVen 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3960
            TabIndex        =   42
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtIntPend 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3960
            TabIndex        =   39
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtCostoNoti 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3960
            TabIndex        =   33
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtCostoCus 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   32
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtInteres 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   31
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtCapital 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   30
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TxtITF 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6240
            TabIndex        =   27
            Top             =   585
            Width           =   1215
         End
         Begin VB.TextBox TxtMontoTotal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   930
            Width           =   1215
         End
         Begin VB.ComboBox cboPlazoNuevo 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmColPPagosParciales.frx":040C
            Left            =   1440
            List            =   "frmColPPagosParciales.frx":0413
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   -120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtMontoPagar 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6240
            MaxLength       =   9
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblMora 
            AutoSize        =   -1  'True
            Caption         =   "Mora"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   45
            Top             =   1320
            Width           =   360
         End
         Begin VB.Label lblIntVen 
            AutoSize        =   -1  'True
            Caption         =   "Interés Vencido"
            Height          =   195
            Index           =   3
            Left            =   2640
            TabIndex        =   44
            Top             =   960
            Width           =   1110
         End
         Begin VB.Label lblInteres 
            AutoSize        =   -1  'True
            Caption         =   "Interés Pend.:"
            Height          =   195
            Index           =   2
            Left            =   2640
            TabIndex        =   48
            Top             =   600
            Width           =   990
         End
         Begin VB.Label lblCostoNoti 
            AutoSize        =   -1  'True
            Caption         =   "Costo Notif."
            Height          =   195
            Index           =   3
            Left            =   2640
            TabIndex        =   37
            Top             =   240
            Width           =   825
         End
         Begin VB.Label lblCostoCus 
            AutoSize        =   -1  'True
            Caption         =   "Costo Custodia"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   36
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label lblInteres 
            AutoSize        =   -1  'True
            Caption         =   "Interés"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   49
            Top             =   600
            Width           =   480
         End
         Begin VB.Label lblCapital 
            AutoSize        =   -1  'True
            Caption         =   "Capital"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ITF"
            Height          =   195
            Left            =   5250
            TabIndex        =   29
            Top             =   675
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Monto Pagar"
            Height          =   195
            Left            =   5250
            TabIndex        =   28
            Top             =   990
            Width           =   915
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Monto "
            Height          =   255
            Index           =   12
            Left            =   5250
            TabIndex        =   13
            Top             =   240
            Width           =   600
         End
      End
      Begin VB.Frame fraContenedor 
         Height          =   1365
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   4320
         Width           =   7485
         Begin VB.TextBox txtSaldoInt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            TabIndex        =   46
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtPgoAdelInt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   40
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtSaldoCapitalNuevo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6360
            TabIndex        =   23
            Top             =   600
            Width           =   960
         End
         Begin VB.TextBox txtPlazoActual 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6720
            MaxLength       =   2
            TabIndex        =   21
            Top             =   240
            Width           =   555
         End
         Begin VB.TextBox txtDiasAtraso 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   18
            Top             =   240
            Width           =   690
         End
         Begin VB.TextBox txtNroRenovacion 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   17
            Top             =   600
            Width           =   690
         End
         Begin VB.TextBox txtMontoMinimoPagar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            TabIndex        =   8
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtTotalDeuda 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            TabIndex        =   7
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label lblInteres 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Int.:"
            Height          =   192
            Index           =   3
            Left            =   2760
            TabIndex        =   47
            Top             =   960
            Width           =   696
         End
         Begin VB.Label lblInteres 
            AutoSize        =   -1  'True
            Caption         =   "Amortización Int.:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   41
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nuevo Saldo"
            Height          =   255
            Index           =   14
            Left            =   5280
            TabIndex        =   24
            Top             =   615
            Width           =   960
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Plazo Actual"
            Height          =   255
            Index           =   11
            Left            =   5280
            TabIndex        =   22
            Top             =   285
            Width           =   960
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Dias Atraso"
            Height          =   210
            Index           =   7
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   960
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nro Renovación"
            Height          =   330
            Index           =   8
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label lblMoneda 
            Height          =   255
            Left            =   2070
            TabIndex        =   14
            Top             =   210
            Width           =   255
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Mínimo a Pagar"
            Height          =   255
            Index           =   10
            Left            =   2520
            TabIndex        =   11
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Total Deuda"
            Height          =   195
            Index           =   9
            Left            =   2520
            TabIndex        =   10
            Top             =   240
            Width           =   1185
         End
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin SICMACT.ActXColPDesCon AXDesCon 
         Height          =   3735
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6588
      End
      Begin VB.Label lblCampRetenPrendAmor 
         Caption         =   "Campaña y Retencion"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   3840
         TabIndex        =   55
         Top             =   180
         Visible         =   0   'False
         Width           =   3135
      End
   End
   Begin VB.Frame fraFormaPago 
      Height          =   600
      Left            =   135
      TabIndex        =   50
      Top             =   7650
      Width           =   7785
      Begin VB.ComboBox CmbForPag 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   200
         Width           =   1785
      End
      Begin SICMACT.ActXCodCta txtCuentaCargo 
         Height          =   375
         Left            =   3120
         TabIndex        =   38
         Top             =   200
         Visible         =   0   'False
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCta      =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblNroDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Left            =   3105
         TabIndex        =   51
         Top             =   250
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label LblNumDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4300
         TabIndex        =   52
         Top             =   200
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblFormaPago 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
         Height          =   195
         Left            =   180
         TabIndex        =   53
         Top             =   250
         Width           =   855
      End
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      Caption         =   "<F12> Número Contrato Antiguo"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   8400
      Width           =   2280
   End
End
Attribute VB_Name = "frmColPPagosParciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************
'PAGOS PARCIALES DE CREDITOS PIGNORATICIOS - sera amortizacion
'Reutilizado a partir del formulario de amortizaciones
'Archivo:  frmColPPagosParciales.frm
'PEAC   :  01/11/2016.
'Resumen:  Nos permite pagar menos del interes calculado como minimo ó mas
'          del interes que viene a ser el pago del interes mas una parte
'          del capital.
'******************************************

Option Explicit

Dim fnVarOpeCod As Long
Dim fsVarOpeDesc As String
Dim fsVarPersCodCMAC As String
Dim fsVarNombreCMAC As String

Dim fnVarTasaPreparacionRemate As Double
Dim fnVarTasaImpuesto As Double
Dim fnVarTasaCustodia As Double
Dim fnVarTasaCustodiaVencida As Double

Dim fnVarTasaInteres As Double
Dim fnVarTasaInteresVencido As Double
Dim fnVarTasaInteresMoratorio As Double

Dim fnVarDiasCambCart As Double

Dim fnVarSaldoCap As Currency
Dim fnVarValorTasacion As Currency
Dim fnVarPlazo As Integer
Dim fdVarFecVencimiento As Date
Dim fnVarEstado As ColocEstado
Dim fnVarNroRenovacion As Integer
Dim fnVarFechaRenovacion As Date

Dim fnVarNewSaldoCap As Currency
Dim fnVarNewPlazo As Integer
Dim fsVarNewFecVencimiento As String
Dim fnVarCapitalPagado As Currency   ' Capital a Pagar
Dim fnVarFactor As Double
Dim fnVarInteres As Currency
Dim fnVarIntPagado As Currency '*** PEAC 20161028
Dim fnVarCostoCustodia As Currency
Dim fnVarCostoCustodiaVencida As Currency
Dim fnVarImpuesto As Currency
Dim fnVarCostoPreparacionRemate As Double

Dim fnVarDiasAtraso As Double
Dim vDiasAtrasoReal As Double
Dim vSumaCostoCustodia As Double
Dim fnVarDeuda As Currency

Dim fnVarMontoMinimo As Currency
Dim fnVarMontoAPagar As Currency

Dim fnVarCostoNotificacion As Currency '*** PEAC 20080515

Dim fsColocLineaCredPig As String ' PEAC 20070813
Dim vFecEstado As Date ' PEAC 20070813
Dim vDiasAdel As Integer, vInteresAdel As Double, vMontoCol As Double ' PEAC 20070813
Dim gcCredAntiguo As String  ' peac 20070923
Dim gnNotifiAdju As Integer  ' peac 20080515
Dim nRedondeoITF As Double  'BRGO 20110906
Dim gnPigPorcenPgoCap As Double  '*** PEAC 20160920
Dim gnPigVigMeses As Integer '*** PEAC 20160920
Dim vCapitalAdel As Currency
Dim fnCredRevolAntNue As Boolean '*** PEAC 20161105 - 1= antiguo, 0 ó null = nuevo
Dim fnIntPend As Double
Dim fnIntPendPagados As Double
Dim fnIntPendSaldo As Double '*** PEAC 20170522
Dim fnVarInteresVencido As Currency '*** PEAC 20170320
Dim fnVarInteresMoratorio As Currency '*** PEAC 20170320
Dim nPagoIntVenMora As Double '*** PEAC 20170331
Dim nUltIntAPagar As Double
Dim lnDiasAtraso  As Integer
Private nMontoVoucher As Currency 'CTI4 ERS0112020
Dim nMovNroRVD As Long, nMovNroRVDPend As Long 'CTI4 ERS0112020
Dim sNumTarj As String 'CTI4 ERS0112020
Dim loVistoElectronico As frmVistoElectronico 'CTI4 ERS0112020
Dim nRespuesta As Integer 'CTI4 ERS0112020

Public Sub Inicio(ByVal pnOpeCod As Long, ByVal psOpeDesc As String, _
        ByVal psPersCodCMAC As String, ByVal psNomCmac As String)

    fnVarOpeCod = pnOpeCod
    fsVarOpeDesc = psOpeDesc
    fsVarPersCodCMAC = psPersCodCMAC
    fsVarNombreCMAC = psNomCmac
    
    Select Case fnVarOpeCod
        Case gColPOpeAmortizEFE
            'txtDocumento.Visible = false
        Case gColPOpeAmortizCHQ
            'txtDocumento.Visible = True
    '    Case Else
    '        txtDocumento.Visible = False
    End Select
    CargaParametros
    Limpiar
    Me.Show 1

End Sub

Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.Limpiar
    txtDiasAtraso.Text = ""
    txtNroRenovacion.Text = ""
    txtTotalDeuda.Text = ""
    txtMontoMinimoPagar.Text = Format(0, "#0.00")
    TxtITF.Text = Format(0, "#0.00")
    TxtMontoTotal = Format(0, "#0.00")
    txtPlazoActual.Text = ""
    txtSaldoCapitalNuevo.Text = Format(0, "#0.00")
    cboPlazoNuevo.ListIndex = 0
    txtMontoPagar.Text = Format(0, "#0.00")
    fnVarCapitalPagado = 0
    vCapitalAdel = 0
    fnVarNewSaldoCap = 0
    txtSaldoCapitalNuevo.Text = Format(0, "#0.00")
    nRedondeoITF = 0
    Me.txtInteres.Text = Format(0, "#0.00")
    Me.txtIntPend.Text = Format(0, "#0.00")
    
    txtIntVen.Text = Format(0, "#0.00")
    txtMora.Text = Format(0, "#0.00")
    
    Me.txtCapital.Text = "": Me.txtCostoNoti.Text = ""
    Me.txtMora.Text = "": Me.txtPgoAdelInt.Text = ""
    Me.txtCostoCus.Text = ""
    CmbForPag.ListIndex = -1 'CTI4 ERS0112020
    txtCuentaCargo.NroCuenta = "" 'CTI4 ERS0112020
    LblNumDoc.Caption = "" 'CTI4 ERS0112020
    cmdGrabar.Enabled = False 'CTI4 ERS0112020
    sNumTarj = "" 'CTI4 ERS0112020
    'Me.txtProxFecha.Text = ""
     'JOEP20210922 campana prendario
    lblCampRetenPrendAmor.Visible = False
    lblCampRetenPrendAmor.Caption = ""
    txtCampReteAmor.Text = 0#
    'JOEP20210922 campana prendario
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call BuscaContrato(AXCodCta.NroCuenta)
End Sub

'Busca el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim lbok As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim lnDeuda As Currency, lnMinimoPagar As Currency
Dim lsmensaje As String
'----- MADM 20091120 ---------------------
Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim Rf As ADODB.Recordset

    Dim loParam As COMDColocPig.DCOMColPCalculos
    Set loParam = New COMDColocPig.DCOMColPCalculos

    fnCredRevolAntNue = loParam.dObtieneParamCredRevolNueAnt(Me.AXCodCta.NroCuenta)

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
    Set lrValida = loValContrato.nValidaPagoParcialCredPignoraticio(psNroContrato, gdFecSis, 0, gsCodUser, lsmensaje, fnCredRevolAntNue)
    If Trim(lsmensaje) <> "" Then
        '*** PEAC 20170131
'             MsgBox lsmensaje, vbInformation, "Aviso"
'             Exit Sub

        If Not lrValida Is Nothing Then
            If lrValida.RecordCount > 0 Then
                MsgBox lsmensaje, vbInformation + vbOKOnly, "Aviso"
                lsmensaje = ""
            End If
        Else
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
        '*** FIN PEAC
    End If
        
    Set loValContrato = Nothing

    If (lrValida Is Nothing) Then       ' Hubo un Error
        Limpiar
        Set lrValida = Nothing
        Exit Sub
    End If

    If Me.AXCodCta.Age <> "" Then

        If AXCodCta.Age = "33" Then
            Dim oColPNotif As New COMDColocPig.DCOMColPActualizaBD
            Dim drNotif As ADODB.Recordset
            
            Set oColPNotif = New COMDColocPig.DCOMColPActualizaBD
            Set drNotif = New ADODB.Recordset
            
            Set drNotif = oColPNotif.DevuelveValorNotificacionCarNotMinka(AXCodCta.NroCuenta)
            If Not (drNotif.EOF And drNotif.BOF) Then
                fnVarCostoNotificacion = drNotif!nValor
            Else
                fnVarCostoNotificacion = loParam.dObtieneParamPignoCostoNotif("COSTO NOTIFIC", Me.AXCodCta.Age)
            End If
        Else
            fnVarCostoNotificacion = loParam.dObtieneParamPignoCostoNotif("COSTO NOTIFIC", Me.AXCodCta.Age)
        End If
   End If

Set loParam = Nothing

    fnVarPlazo = lrValida!nPlazo
    fnVarSaldoCap = Format(lrValida!nSaldo, "#0.00")
    fnVarValorTasacion = lrValida!nTasacion
    nPagoIntVenMora = lrValida!nPagosIntVenMora

    fnVarTasaInteresVencido = lrValida!nTasaIntVenc
    fnVarTasaInteresMoratorio = lrValida!nTasaIntMora '*** PEAC 20170320

    nUltIntAPagar = lrValida!nUltIntAPagar
    fnVarEstado = lrValida!nPrdEstado
    fnVarTasaInteres = lrValida!nTasaInteres
    fdVarFecVencimiento = Format(lrValida!dVenc, "dd/mm/yyyy")
    fnVarFechaRenovacion = lrValida!dRenovacion
    gcCredAntiguo = lrValida!cCredB 'PEAC 20071106
    
    fnIntPendPagados = lrValida!nPagosPendIntPagados
    fnIntPendSaldo = lrValida!nPagosPendIntSaldo
    
    vFecEstado = Format(lrValida!dPrdEstado, "dd/mm/yyyy") ' PEAC 20070813
    fnVarSaldoCap = lrValida!nMontoCol ' PEAC 20070813
    gnNotifiAdju = lrValida!nCodNotifiAdj 'PEAC 20080515
    
    fnVarNroRenovacion = lrValida!nNroRenov
    fnVarNewPlazo = lrValida!nPlazo
    
    If fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, False) Then
    
    End If
     
    'lnDiasAtraso = DateDiff("d", Format(lrValida!dVenc, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy")) RIRO 20200401 Comentado
    lnDiasAtraso = lrValida!nDiasAtraso 'RIRO 20200401 Add
    Me.txtDiasAtraso = Val(lnDiasAtraso)
    txtNroRenovacion.Text = Val(lrValida!nNroRenov) + 1
    txtPlazoActual.Text = Val(lrValida!nPlazo)
    
    'JOEP20210921 campana prendario
    Call CampPrendVerificaCampanas(psNroContrato, gdFecSis, 2, txtNroRenovacion.Text)
    AXDesCon.TasaEfectivaMensual = fnVarTasaInteres
    'JOEP20210921 campana prendario
    
    If lnDiasAtraso > 0 Then
        '*** PEAC 20170329
        If fnCredRevolAntNue = True Then ' true =SI ES ANTIGUO
            If Trim(LeeConstanteSist(606)) = 0 Then '' 1=si permite 0=no permite
                MsgBox "Contrato se encuentra Vencido, por favor realice la renovación.", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    
    End If

    fgCalculaDeuda lnDiasAtraso 'RIRO 20200406 Se añadió lnDiasAtraso

    fgCalculaMinimoPagar
    
    txtMontoMinimoPagar.Text = Format(fnVarMontoMinimo, "#0.00")
    TxtMontoTotal.Text = Format(fnVarMontoMinimo + fnVarInteres, "#0.00")
    
    If gITF.gbITFAplica Then
        If Not gITF.gbITFAsumidocreditos Then
             
           TxtITF = gITF.fgITFCalculaImpuesto(txtMontoPagar)
           txtMontoPagar = CCur(TxtMontoTotal.Text) + CCur(TxtITF.Text)
           
           Dim Aux As String
           If InStr(1, CStr(TxtITF), ".", vbTextCompare) > 0 Then
                Aux = CDbl(CStr(Int(TxtITF)) & "." & Mid(CStr(TxtITF), InStr(1, CStr(TxtITF), ".", vbTextCompare) + 1, 2))
           Else
                Aux = CDbl(CStr(Int(TxtITF)))
           End If
            
            TxtITF.Text = Format(Aux, "#0.00")
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtITF.Text))
            If nRedondeoITF > 0 Then
               Me.TxtITF.Text = Format(CCur(Me.TxtITF.Text) - nRedondeoITF, "#,##0.00")
            End If

        Else
            Me.TxtITF = Format(gITF.fgITFCalculaImpuesto(txtMontoPagar.Text), "#0.00")
            '*** BRGO 20110908 ************************************************
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtITF.Text))
            If nRedondeoITF > 0 Then
               Me.TxtITF.Text = Format(CCur(Me.TxtITF.Text) - nRedondeoITF, "#,##0.00")
            End If
            '*** END BRGO
            TxtMontoTotal = Format(CDbl(Me.txtMontoPagar), "#0.00")
        End If
    Else
            Me.TxtITF = Format(0, "#0.00")
            TxtMontoTotal = Format(Me.txtMontoPagar, "#0.00")
    End If
    
    txtSaldoCapitalNuevo.Text = "0.00"
    
'*** PEAC 20080528
If CCur(AXDesCon.SaldoCapital) > 0 Then
    fnVarDeuda = CCur(AXDesCon.SaldoCapital) + TxtITF.Text + _
                fnVarInteres + _
                fnVarCostoCustodiaVencida + fnVarImpuesto + _
                fnVarCostoPreparacionRemate + fnVarCostoNotificacion + _
                fnVarInteresVencido + fnVarInteresMoratorio '*** PEAC 20170320

    txtTotalDeuda.Text = Format(fnVarDeuda, "#0.00")

    txtCapital.Text = 0 'Format(txtMontoPagar.Text - fnVarCostoNotificacion - fnVarInteres - TxtITF.Text, "#0.00")
    'txtCapital.Text = vCapitalAdel '*** PEAC 20161020
    txtInteres.Text = Format(fnVarInteres, "#0.00")
    txtCostoCus.Text = Format(fnVarCostoCustodiaVencida, "#0.00")
    txtCostoNoti.Text = Format(fnVarCostoNotificacion, "#0.00")
    txtSaldoCapitalNuevo.Text = Format(CCur(AXDesCon.SaldoCapital) - fnVarCapitalPagado, "#0.00")

    '*** PEAC 20161116  fnIntPendPagados
    Me.txtSaldoInt.Text = Format(fnIntPendSaldo, "#0.00")
    Me.txtPgoAdelInt.Text = Format(IIf(fnIntPendSaldo = 0, 0, fnIntPendPagados), "#0.00")
    
    '*** PEAC 20170320
    txtMora.Text = Format(fnVarInteresMoratorio, "#0.00")
    txtIntVen.Text = Format(fnVarInteresVencido, "#0.00")
    
End If

    Set lrValida = Nothing
    
    AXCodCta.Enabled = False
    cmdGrabar.Enabled = True
    txtMontoPagar.Enabled = True
    Me.txtMontoPagar.SetFocus
        
     Set lafirma = New frmPersonaFirma
     Set ClsPersona = New COMDPersona.DCOMPersonas
    
     Set Rf = ClsPersona.BuscaCliente(gColPigFunciones.vcodper, BusquedaCodigo)
     If Not Rf.BOF And Not Rf.EOF Then
        If Rf!nPersPersoneria = 1 Then
            Call frmPersonaFirma.Inicio(Trim(gColPigFunciones.vcodper), Mid(gColPigFunciones.vcodper, 4, 2), False, False) 'mod by jato 20210324
        End If
     End If
     Set Rf = Nothing

    CmbForPag.Enabled = True 'CTI4 ERS0112020
    CmbForPag.ListIndex = IndiceListaCombo(CmbForPag, 1) 'CTI4 ERS0112020

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

''*** PEAC 20161021
'Private Sub Check1_Click()
'    If Check1.value = 1 Then
'        Me.txtProxFecha = AXDesCon.FechaVencimiento
'    Else
'        Me.txtProxFecha = Format(DateAdd("d", 30, gdFecSis), "dd/MM/yyyy")
'    End If
'End Sub


Private Sub cmdBuscar_Click()

Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersContrato As COMDColocPig.DCOMColPContrato
Dim lrContratos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

lsEstados = gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov

If Trim(lsPersCod) <> "" Then
    Set loPersContrato = New COMDColocPig.DCOMColPContrato
        Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstados, Mid(gsCodAge, 4, 2))
    Set loPersContrato = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmCuentasPersona.Inicio(lsPersNombre, lrContratos) ' RIRO 20130724 SEGUN ERS101-2013
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Cancela el proceso actual y permite inicializar ls variables para otro proceso
Private Sub cmdCancelar_Click()
    Limpiar
    cmdGrabar.Enabled = False
    cboPlazoNuevo.Enabled = False
    txtMontoPagar.Enabled = False
    AXCodCta.Enabled = True
    CmbForPag.Enabled = False 'CTI4 ERS0112020
    AXCodCta.SetFocusCuenta
End Sub

'Actualiza los cambios en la basede datos
Private Sub cmdGrabar_Click()
Dim PersonaPago() As Variant
Dim nI As Integer
Dim fnCondicion As Integer
Dim regPersonaRealizaPago As Boolean
nI = 0
gnMovNro = 0

Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarRen As COMNColoCPig.NCOMColPContrato 'COMNColoCPig.NCOMColPContrato
Dim loImprime As COMNColoCPig.NCOMColPImpre
Dim loPrevio As previo.clsprevio

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsFechaVenc As String
Dim lnMontoTransaccion As Currency
Dim lsCadImprimir As String
Dim lsNombreCliente As String

Dim loLavDinero As frmMovLavDinero
Set loLavDinero = New frmMovLavDinero
Dim loMov As COMDMov.DCOMMov
Dim lsFecVenImp  As String
Dim lsBoletaCargo  As String 'CTI4 ERS0112020
Dim MatDatosAho(14) As String 'CTI4 ERS0112020
Dim lsNombreClienteCargoCta As String 'CTI4 ERS0112020
Dim objPig As COMDColocPig.DCOMColPContrato 'JOEP20210917 campana prendario
If Not ValidaFormaPago Then Exit Sub 'CTI4 ERS0112020

If Not ValidaAlGrabar Then Exit Sub

lsFechaVenc = Format$(fdVarFecVencimiento, "mm/dd/yyyy")
lsFecVenImp = Format$(fdVarFecVencimiento, "dd/MM/yyyy")

lsNombreCliente = AXDesCon.listaClientes.ListItems(1).ListSubItems.Item(1)

If CDbl(txtMontoPagar.Text) <= 0 And gcCredAntiguo <> "A" Then ' peac 20071106
   'MsgBox "El monto del Pago Parcial debe ser mayor a cero.", vbOKOnly + vbInformation, "AVISO"
   MsgBox "El monto de la Amortización debe ser mayor a cero.", vbOKOnly + vbInformation, "AVISO"
   Exit Sub
End If

Dim oDPersona As COMDPersona.DCOMPersona
Dim rsPersonaCred As ADODB.Recordset
Dim rsPersona As ADODB.Recordset
Dim Cont As Integer
Set oDPersona = New COMDPersona.DCOMPersona

Set rsPersonaCred = oDPersona.ObtenerPersCuentaRelac(Trim(AXCodCta.NroCuenta), gColRelPersTitular)

If rsPersonaCred.RecordCount > 0 Then
    If Not (rsPersonaCred.EOF And rsPersonaCred.BOF) Then
        For Cont = 0 To rsPersonaCred.RecordCount - 1
            ReDim Preserve PersonaPago(Cont, 1)
            PersonaPago(Cont, 0) = Trim(rsPersonaCred!cperscod)
            PersonaPago(Cont, 1) = Trim(rsPersonaCred!cPersNombre)
            Set rsPersona = oDPersona.ObtenerUltimaVisita(Trim(rsPersonaCred!cperscod))
            If rsPersona.RecordCount > 0 Then
                If Not (rsPersona.EOF And rsPersona.BOF) Then
                    If Trim(rsPersona!sUsual) = "3" Then
                    MsgBox PstaNombre(Trim(rsPersonaCred!cPersNombre), True) & "." & Chr(10) & "CLIENTE OBSERVADO: " & Trim(rsPersona!cVisObserva), vbInformation, "Aviso"
                        Call frmPersona.Inicio(Trim(rsPersonaCred!cperscod), PersonaActualiza)
                    End If
                   
                End If
            End If
            Call VerSiClienteActualizoAutorizoSusDatos(Trim(rsPersonaCred!cperscod), fnVarOpeCod) 'FRHU ERS077-2015 20151204
            Set rsPersona = Nothing
            rsPersonaCred.MoveNext
        Next Cont
    End If
End If

Dim oDPersonaAct As COMDPersona.DCOMPersona
Set oDPersonaAct = New COMDPersona.DCOMPersona
                If oDPersonaAct.VerificaExisteSolicitudDatos(gColPigFunciones.vcodper) Then
                    MsgBox Trim("SE SOLICITA DATOS DEL CLIENTE: " & lsNombreCliente) & "." & Chr(10), vbInformation, "Aviso"
                    Call frmActInfContacto.Inicio(gColPigFunciones.vcodper)
                End If

'If MsgBox("¿Desea realizar el Pago Parcial del Contrato Pignoraticio? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
If MsgBox("¿Desea realizar la Amortización del Contrato Pignoraticio? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then

    cmdGrabar.Enabled = False
        
        If fnVarCapitalPagado <= 0 Then
            fnVarCapitalPagado = 0
            lnMontoTransaccion = CCur(Me.txtMontoPagar.Text) + CCur(TxtITF.Text)
        Else
            lnMontoTransaccion = CCur(Me.txtMontoPagar.Text)
        End If
        
        fnVarNewSaldoCap = Format(CCur(AXDesCon.SaldoCapital) - vCapitalAdel - fnVarCapitalPagado, "#0.00")
      
        Dim clsExo As New COMNCaptaServicios.NCOMCaptaServicios
        If Not clsExo.EsCuentaExoneradaLavadoDinero(AXCodCta.NroCuenta) Then
          Dim sPersLavDinero As String
          Dim nMontoLavDinero As Double, nTC As Double
          Dim clsLav As New COMNCaptaGenerales.NCOMCaptaDefinicion, nmoneda As Integer, nMonto As Double
    
            nMonto = CCur(Me.txtMontoPagar.Text)
            
            Set clsExo = Nothing
            sPersLavDinero = ""
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            nmoneda = gMonedaNacional
            If nmoneda = gMonedaNacional Then
                Dim clsTC As COMDConstSistema.NCOMTipoCambio
                Set clsTC = New COMDConstSistema.NCOMTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
            Else
                nTC = 1
            End If
            If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                Call IniciaLavDinero(loLavDinero)
                sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, AXCodCta.NroCuenta, Mid(Me.Caption, 15), True, "", , , , , nmoneda)
                If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
            End If
         Else
            Set clsExo = Nothing
         End If
         Set clsExo = Nothing
        If loLavDinero.OrdPersLavDinero = "Exit" Then
            Dim oPersonaSPR As UPersona_Cli
            Dim oPersonaU As COMDPersona.UCOMPersona
            Dim nTipoConBN As Integer
            Dim sConPersona As String
            Dim pbClienteReforzado As Boolean
            Dim rsAgeParam As Recordset
            Dim objCred As COMNCredito.NCOMCredito
            Dim lnMonto As Double, lnTC As Double
            Dim ObjTc As COMDConstSistema.NCOMTipoCambio
            
            Set oPersonaU = New COMDPersona.UCOMPersona
            Set oPersonaSPR = New UPersona_Cli
            
            regPersonaRealizaPago = False
            pbClienteReforzado = False
            fnCondicion = 0
            
            For nI = 0 To UBound(PersonaPago)
                oPersonaSPR.RecuperaPersona Trim(PersonaPago(nI, 0))
                                    
                If oPersonaSPR.Personeria = 1 Then
                    If oPersonaSPR.Nacionalidad <> "04028" Then
                        sConPersona = "Extranjera"
                        fnCondicion = 1
                        pbClienteReforzado = True
                        Exit For
                    ElseIf oPersonaSPR.Residencia <> 1 Then
                        sConPersona = "No Residente"
                        fnCondicion = 2
                        pbClienteReforzado = True
                        Exit For
                    ElseIf oPersonaSPR.RPeps = 1 Then
                        sConPersona = "PEPS"
                        fnCondicion = 4
                        pbClienteReforzado = True
                        Exit For
                    ElseIf oPersonaU.ValidaEnListaNegativaCondicion(IIf(Trim(oPersonaSPR.ObtenerDNI) = "", oPersonaSPR.ObtenerNumeroDoc(0), oPersonaSPR.ObtenerDNI), oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                        If nTipoConBN = 1 Or nTipoConBN = 3 Then
                            sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                            fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                            pbClienteReforzado = True
                            Exit For
                        End If
                    End If
                Else
                    If oPersonaU.ValidaEnListaNegativaCondicion(oPersonaSPR.ObtenerDNI, oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                        If nTipoConBN = 1 Or nTipoConBN = 3 Then
                            sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                            fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                            pbClienteReforzado = True
                            Exit For
                        End If
                    End If
                End If
            Next nI
            
            If pbClienteReforzado Then
                MsgBox "El Cliente: " & Trim(PersonaPago(nI, 1)) & " es un Cliente de Procedimiento Reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
                frmPersRealizaOpeGeneral.Inicia fsVarOpeDesc & " (Persona " & sConPersona & ")", fnVarOpeCod
                regPersonaRealizaPago = frmPersRealizaOpeGeneral.PersRegistrar
                
                If Not regPersonaRealizaPago Then
                    MsgBox "Se va a proceder a Anular la Operación", vbInformation, "Aviso"
                    cmdGrabar.Enabled = True
                    Exit Sub
                End If
            Else
                fnCondicion = 0
                lnMonto = nMonto
                pbClienteReforzado = False
                
                Set ObjTc = New COMDConstSistema.NCOMTipoCambio
                lnTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set ObjTc = Nothing
            
            
                Set objCred = New COMNCredito.NCOMCredito
                Set rsAgeParam = objCred.obtieneCredPagoCuotasAgeParam(gsCodAge)
                Set objCred = Nothing
                
                If Mid(AXCodCta.NroCuenta, 9, 1) = 2 Then
                    lnMonto = Round(lnMonto * lnTC, 2)
                End If
            
                If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                    If lnMonto >= rsAgeParam!nMontoMin And lnMonto <= rsAgeParam!nMontoMax Then
                        frmPersRealizaOpeGeneral.Inicia fsVarOpeDesc, fnVarOpeCod
                        regPersonaRealizaPago = frmPersRealizaOpeGeneral.PersRegistrar
                        If Not regPersonaRealizaPago Then
                            MsgBox "Se va a proceder a Anular la Operación", vbInformation, "Aviso"
                            cmdGrabar.Enabled = True
                            Exit Sub
                        End If
                    End If
                End If
                
            End If
        End If
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarRen = New COMNColoCPig.NCOMColPContrato
            'Grabar Pago Parcial Pignoraticio - puede pagar menos del interes calculado - la dif debera pagar al siguiente operacion

        'CTI4 ERS0112020
        Select Case CInt(Trim(Right(CmbForPag.Text, 10)))
            Case gColocTipoPagoEfectivo
                fnVarOpeCod = gColPOpePagoParcialEfectivo
            Case gColocTipoPagoVoucher
                fnVarOpeCod = gColPOpePagoParcialNorVoucher
            Case gColocTipoPagoCargoCta
                fnVarOpeCod = gColPOpePagoParcialNorCargoCta
        End Select
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then lsNombreClienteCargoCta = PstaNombre(loGrabarRen.ObtieneNombreTitularCargoCta(txtCuentaCargo.NroCuenta))
        'end CTI4

            Call loGrabarRen.nPagoParcialCredPignoraticio(AXCodCta.NroCuenta, fnVarNewSaldoCap, lsFechaHoraGrab, _
                 lsMovNro, lsFechaVenc, fnVarNewPlazo, lnMontoTransaccion - CCur(Val(Me.TxtITF.Text)), fnVarCapitalPagado, fnVarInteresVencido, _
                 fnVarCostoCustodiaVencida, fnVarCostoPreparacionRemate, CCur(Me.txtInteres.Text), fnVarImpuesto, fnVarCostoCustodia, _
                 fnVarDiasAtraso, fnVarDiasCambCart, fnVarValorTasacion, fnVarOpeCod, _
                 fsVarOpeDesc, fsVarPersCodCMAC, gITF.gbITFAplica, gITF.gbITFAsumidocreditos, CCur(Val(Me.TxtITF.Text)), False, fnVarCostoNotificacion, gnMovNro, CDbl(Me.txtIntPend.Text), fnVarInteresMoratorio, _
                CInt(Trim(Right(CmbForPag.Text, 10))), nMovNroRVD, nMovNroRVDPend, txtCuentaCargo.NroCuenta, MatDatosAho)
                
'            Call loGrabarRen.nPagoParcialCredPignoraticio(AXCodCta.NroCuenta, fnVarNewSaldoCap, lsFechaHoraGrab, _
'                 lsMovNro, lsFechaVenc, fnVarNewPlazo, lnMontoTransaccion - CCur(Val(Me.TxtITF.Text)), fnVarCapitalPagado, fnVarInteresVencido, _
'                 fnVarCostoCustodiaVencida, fnVarCostoPreparacionRemate, CCur(Me.Txtinteres.Text), fnVarImpuesto, fnVarCostoCustodia, _
'                 fnVarDiasAtraso, fnVarDiasCambCart, fnVarValorTasacion, fnVarOpeCod, _
'                 fsVarOpeDesc, fsVarPersCodCMAC, gITF.gbITFAplica, gITF.gbITFAsumidocreditos, CCur(Val(Me.TxtITF.Text)), CInt(Trim(Right(CmbForPag.Text, 10))), False, fnVarCostoNotificacion, gnMovNro, CDbl(Me.txtIntPend.Text), fnVarInteresMoratorio, _
'                nMovNroRVD, nMovNroRVDPend, txtCuentaCargo.NroCuenta, MatDatosAho)

        Set loGrabarRen = Nothing
        
        If gITF.gbITFAplica Then
           Set loMov = New COMDMov.DCOMMov
           Call loMov.InsertaMovRedondeoITF(lsMovNro, 1, CCur(Me.TxtITF) + nRedondeoITF, CCur(Me.TxtITF))
           Set loMov = Nothing
        End If
        
        'JOEP20210917 campana prendario
        If lblCampRetenPrendAmor.Visible = True Then
            Set objPig = New COMDColocPig.DCOMColPContrato
            Call objPig.CampPrenRegCampCred(AXCodCta.NroCuenta, 0, "Amortizacion", AXDesCon.TasaEfectivaMensual, txtNroRenovacion.Text, 0, 0, lsMovNro, 5, 3)
            Set objPig = Nothing
        End If
        'JOEP20210917 campana prendario
        
        'ADD JHCU 09-07-2020 REVERSIÓN PIGNORATICIO
         Set loGrabarRen = New COMNColoCPig.NCOMColPContrato
         Call loGrabarRen.ReversionReprogramacion(lsMovNro, AXCodCta.NroCuenta)
         Set loGrabarRen = Nothing
        'FIN JHCU 09-07-2020

        Set loImprime = New COMNColoCPig.NCOMColPImpre

            lsCadImprimir = loImprime.nPrintReciboAmortizacion(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
                Format(AXDesCon.FechaPrestamo, "dd/MM/yyyy"), fnVarDiasAtraso, CCur(AXDesCon.SaldoCapital), fnVarCapitalPagado, _
                CDbl(Me.txtInteres.Text), fnVarImpuesto, fnVarCostoCustodiaVencida + fnVarCostoCustodia, _
                fnVarCostoPreparacionRemate, lnMontoTransaccion, fnVarNewSaldoCap, fnVarTasaInteres, _
                Val(Me.txtNroRenovacion.Text), Format(lsFecVenImp, "dd/MM/yyyy"), gsCodUser, fnVarNewPlazo, _
                fsVarNombreCMAC, " ", CDbl(Val(TxtITF.Text)), gImpresora, gbImpTMU, fnVarCostoNotificacion, fnVarInteresVencido, fnVarInteresMoratorio, CCur(Me.txtIntPend.Text))
        '*** PEAC - SE AGREGÓ : fnVarInteresVencido, fnVarInteresMoratorio, CCur(Me.txtIntPend.Text)

       'CTI4 ERS0112020
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            lsBoletaCargo = loImprime.ImprimeBoletaAhorro("RETIRO AHORROS", "CARGO A CUENTA POR AMORT. PIGNO.", "", CStr(lnMontoTransaccion + Me.TxtITF.Text), lsNombreClienteCargoCta, txtCuentaCargo.NroCuenta, "", CDbl(MatDatosAho(10)), CDbl(MatDatosAho(3)), "", 1, CDbl(MatDatosAho(11)), , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, gsCodCMAC, 0, , , gbImpTMU)
        End If
        'END CTI4

        Set loImprime = Nothing
        Set loPrevio = New previo.clsprevio
        
            loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
            Do While True
                'If MsgBox("¿Desea Reimprimir el Recibo del Pago Parcial? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                If MsgBox("¿Desea Reimprimir el Recibo de la Amortización? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
                    
                Else
                    Set loPrevio = Nothing
                    Exit Do
                End If
            Loop
            Set loPrevio = Nothing
            Set loLavDinero = Nothing
            
            'CTI4 ERS0112020
            If Trim(lsBoletaCargo) <> "" Then
                Set loPrevio = New previo.clsprevio
                loPrevio.PrintSpool sLpt, lsBoletaCargo, False, 22
                
                Do While True
                    If MsgBox("Desea reimprimir boleta del cargo a cuenta?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                        Set loPrevio = New previo.clsprevio
                            loPrevio.PrintSpool sLpt, lsBoletaCargo, False
                        Set loPrevio = Nothing
                    Else
                        Exit Do
                    End If
                Loop
            End If
            'END CTI4 ERS0112020

        If regPersonaRealizaPago And gnMovNro > 0 Then
            frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(AXCodCta.NroCuenta), fnCondicion
            regPersonaRealizaPago = False
        End If
        'CTI4 ERS0112020
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            Dim oMovOperacion As COMDMov.DCOMMov
            Dim nMovNroOperacion As Long
            Dim rsCli As New ADODB.Recordset
            Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
            Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
            Set oMovOperacion = New COMDMov.DCOMMov
            nMovNroOperacion = oMovOperacion.GetnMovNro(lsMovNro)

            loVistoElectronico.RegistraVistoElectronico nMovNroOperacion, , gsCodUser, nMovNroOperacion

            If nRespuesta = 2 Then
                Set rsCli = clsCli.GetPersonaCuenta(txtCuentaCargo.NroCuenta, gCapRelPersTitular)
                oSolicitud.ActualizarCapAutSinTarjetaVisto_nMovNro gsCodUser, gsCodAge, txtCuentaCargo.NroCuenta, rsCli!cperscod, nMovNroOperacion, CStr(gAhoCargoCtaAmortizaPignoNor)
            End If
            Set oMovOperacion = Nothing
            nRespuesta = 0
        End If
        'CTI4 end
        Limpiar
        
        AXCodCta.Enabled = True
        AXCodCta.SetFocus
        'INICIO JHCU ENCUESTA 16-10-2019
        Encuestas gsCodUser, gsCodAge, "ERS0292019", fnVarOpeCod
        'FIN
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
Private Sub IniciaLavDinero(poLavDinero As frmMovLavDinero)
Dim i As Long
Dim nRelacion As CaptacRelacPersona
Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
Dim nMonto As Double, nPersoneria As Integer
Dim sCuenta As String
    nPersoneria = gPersonaNat
    If nPersoneria = gPersonaNat Then
            poLavDinero.TitPersLavDinero = AXDesCon.listaClientes.ListItems(1).Text
            poLavDinero.TitPersLavDineroNom = AXDesCon.listaClientes.ListItems(1).SubItems(1)
            poLavDinero.TitPersLavDineroDir = AXDesCon.listaClientes.ListItems(1).SubItems(2)
            poLavDinero.TitPersLavDineroDoc = AXDesCon.listaClientes.ListItems(1).SubItems(7)
    Else
             poLavDinero.TitPersLavDinero = AXDesCon.listaClientes.ListItems(1).Text
             poLavDinero.TitPersLavDineroNom = AXDesCon.listaClientes.ListItems(1).SubItems(1)
             poLavDinero.TitPersLavDineroDir = AXDesCon.listaClientes.ListItems(1).SubItems(2)
             poLavDinero.TitPersLavDineroDoc = AXDesCon.listaClientes.ListItems(1).SubItems(9)
    End If
nMonto = CDbl(TxtMontoTotal.Text)
sCuenta = AXCodCta.NroCuenta
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub fgCalculaDeuda(Optional ByVal nDiasAtr As Integer = -1)
'nDiasAtr RIRO 20200406 ADD
Dim loCalculos As COMNColoCPig.NCOMColPCalculos
'fnVarDiasAtraso = DateDiff("d", fdVarFecVencimiento, gdFecSis) RIRO 20200406 COMENTADO
fnVarDiasAtraso = IIf(nDiasAtr < 0, DateDiff("d", fdVarFecVencimiento, gdFecSis), nDiasAtr) 'RIRO 20200401

If fnVarDiasAtraso <= 0 Then
    fnVarDiasAtraso = 0
    '*** PEAC 20170320
    fnVarInteresVencido = 0
    fnVarInteresMoratorio = 0
    '*** FIN PEAC
    
    If gcCredAntiguo = "A" Then
        fnVarInteres = Round(0, 2)
    Else
        Set loCalculos = New COMNColoCPig.NCOMColPCalculos
        vDiasAdel = DateDiff("d", vFecEstado, Format(gdFecSis, "dd/mm/yyyy"))
        
        fnVarInteres = loCalculos.nCalculaInteresAlVencimiento(CCur(AXDesCon.SaldoCapital), fnVarTasaInteres, vDiasAdel)
        '*** PEAC 20161221
        
        fnVarInteres = fnVarInteres + fnIntPendSaldo 'nUltIntAPagar
        fnVarInteres = Round(IIf((fnVarInteres) <= 0, 0, fnVarInteres), 2)
        
        '*** FIN PEAC
        'fnIntPend = Round(fnVarInteres, 2)
        fnIntPend = Round(fnIntPendSaldo, 2)
        
        Set loCalculos = Nothing
    End If

    fnVarInteresVencido = 0
    fnVarCostoCustodia = 0
    fnVarImpuesto = 0
Else
    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
        
        '*** PEAC 20170906 - MEJORA EL CALCULO DE LOS DIAS DE MORA
        If Format(vFecEstado, "yyyymmdd") >= Format(fdVarFecVencimiento, "yyyymmdd") Then
            vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
        Else
            'vDiasAdel = 30
            vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(fdVarFecVencimiento, "dd/mm/yyyy"))
        End If
        
'        If vFecEstado > Format(AXDesCon.FechaVencimiento) Then
'            vDiasAdel = 30
'        Else
'            vDiasAdel = DateDiff("d", vFecEstado, Format(AXDesCon.FechaVencimiento, "dd/mm/yyyy"))
'        End If
        '*** FIN PEAC
                
        fnVarInteres = loCalculos.nCalculaInteresAlVencimiento(CCur(AXDesCon.SaldoCapital), fnVarTasaInteres, vDiasAdel)
                
        fnVarInteres = fnVarInteres + fnIntPendSaldo
                
        '*** PEAC 20170320
        'If nPagoIntVenMora > 0 then
        If DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy")) = 0 Then '*** PEAC 20170331
            fnVarInteresVencido = Round(0, 2)
            fnVarInteresMoratorio = Round(0, 2)
        Else
        
            If Format(vFecEstado, "yyyymmdd") >= Format(fdVarFecVencimiento, "yyyymmdd") Then
                fnVarDiasAtraso = vDiasAdel
            End If
        
            fnVarInteresVencido = loCalculos.nCalculaInteresMoratorio(CCur(AXDesCon.SaldoCapital), fnVarTasaInteresVencido, fnVarDiasAtraso, fnVarInteres)
            fnVarInteresVencido = Round(fnVarInteresVencido, 2)
            
            fnVarInteresMoratorio = loCalculos.nCalculaInteresMoratorio(CCur(AXDesCon.SaldoCapital), fnVarTasaInteresMoratorio, fnVarDiasAtraso)
            fnVarInteresMoratorio = Round(fnVarInteresMoratorio, 2)
        End If

        '*** FIN PEAC
        
        fnVarCostoCustodiaVencida = loCalculos.nCalculaCostoCustodiaMoratorio(fnVarValorTasacion, fnVarTasaCustodiaVencida, fnVarDiasAtraso)
        fnVarCostoCustodiaVencida = Round(fnVarCostoCustodiaVencida, 2)
        
        fnVarImpuesto = (fnVarInteresVencido + fnVarInteresMoratorio + fnVarCostoCustodiaVencida) * fnVarTasaImpuesto '*** PEAC 20170320 - SE AGREGO fnVarInteresMoratorio
        fnVarImpuesto = Round(fnVarImpuesto, 2)
        
    Set loCalculos = Nothing
End If
fnVarCostoPreparacionRemate = 0

If fnVarEstado = gColPEstPRema Then    ' Si esta en via de Remate
    fnVarCostoPreparacionRemate = fnVarTasaPreparacionRemate * fnVarValorTasacion
    fnVarCostoPreparacionRemate = Round(fnVarCostoPreparacionRemate, 2)
End If

fnVarCostoNotificacion = 0

End Sub

Private Sub fgCalculaMinimoPagar()
Dim loCalculos As COMNColoCPig.NCOMColPCalculos

' **************************************
' ** Calculo del Monto Minimo a Pagar **
' **************************************
Set loCalculos = New COMNColoCPig.NCOMColPCalculos
    fnVarFactor = loCalculos.nCalculaFactorRenovacion(fnVarTasaInteres, fnVarNewPlazo)
    fnVarCostoCustodia = loCalculos.nCalculaCostoCustodia(fnVarValorTasacion, fnVarTasaCustodia, fnVarNewPlazo)
    fnVarCostoCustodia = Round(fnVarCostoCustodia, 2)
    
    fnVarImpuesto = (fnVarInteresVencido + fnVarInteresMoratorio + fnVarInteres + fnVarCostoCustodia + fnVarCostoCustodiaVencida) * fnVarTasaImpuesto
    fnVarImpuesto = Round(fnVarImpuesto, 2)
    
    If fnVarFechaRenovacion = gdFecSis Then
        fnVarMontoMinimo = Round(0.01, 2)
    Else
        If gcCredAntiguo = "A" Then
            fnVarMontoMinimo = Round(0.01, 2)
        Else
            'fnVarMontoMinimo = fnVarInteresVencido + fnVarInteresMoratorio + fnVarCostoCustodiaVencida + fnVarInteres + fnVarCostoCustodia + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarCostoNotificacion
            'fnVarMontoMinimo = fnVarInteresVencido + fnVarInteresMoratorio + fnVarCostoCustodiaVencida + fnVarCostoCustodia + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarCostoNotificacion
            fnVarMontoMinimo = fnVarInteresVencido + fnVarInteresMoratorio + fnVarCostoCustodiaVencida + fnVarCostoCustodia + fnVarImpuesto + fnVarCostoNotificacion
            fnVarMontoMinimo = Round(fnVarMontoMinimo, 2)

        End If
    End If
    
Set loCalculos = Nothing
End Sub
Private Sub cboPlazoNuevo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

End If
End Sub
Private Sub cboPlazoNuevo_Click()
    fnVarNewPlazo = Val(cboPlazoNuevo.Text)
    fgCalculaMinimoPagar
    txtMontoMinimoPagar.Text = Format(fnVarMontoMinimo, "#0.00")
    txtMontoPagar.Text = Format(fnVarMontoMinimo, "#0.00")
    fnVarCapitalPagado = 0
    txtSaldoCapitalNuevo.Text = Format(fnVarSaldoCap - fnVarCapitalPagado, "#0.00")
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
Call CargaControles 'CTI4 ERS0112020
'JOEP20210922 campana prendario
    lblCampRetenPrendAmor.Visible = False
    lblCampRetenPrendAmor.Caption = ""
    txtCampReteAmor.Text = 0#
'JOEP20210922 campana prendario
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColProConsumoPrendario, False)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
        End If
    ElseIf KeyCode = 13 And Trim(AXCodCta.EnabledCta) And AXCodCta.Age <> "" And Trim(AXCodCta.Cuenta) = "" Then
                AXCodCta.SetFocusCuenta
                 Exit Sub
    End If
End Sub

Private Sub txtMontoPagar_Change()
Dim loValida As COMNColoCPig.NCOMColPValida 'CTI4ERS0112020
Dim bEsMismoTitular As Boolean 'CTI4 ERS0112020
    Set loValida = New COMNColoCPig.NCOMColPValida
    
fnVarInteres = fnVarInteres
bEsMismoTitular = loValida.EsMismoTitulardeCuentaPignoYAhorro(txtCuentaCargo.NroCuenta, AXCodCta.NroCuenta)

If IsNumeric(txtMontoPagar.Text) Then
    fnVarCapitalPagado = IIf((txtMontoPagar.Text - fnVarInteres - fnVarInteresMoratorio - fnVarInteresVencido - TxtITF.Text) > 0, (txtMontoPagar.Text - fnVarInteres - fnVarInteresMoratorio - fnVarInteresVencido - TxtITF.Text), 0) ' peac 20070820

     If gITF.gbITFAplica And Not bEsMismoTitular Then
        If Not gITF.gbITFAsumidocreditos Then
            TxtMontoTotal.Text = "0.00"

            If txtMontoPagar.Text <> txtMontoMinimoPagar.Text Then
              TxtITF.Text = Format(gITF.fgITFCalculaImpuesto(txtMontoPagar.Text))
              nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtITF.Text))
              If nRedondeoITF > 0 Then
                  Me.TxtITF.Text = Format(CCur(Me.TxtITF.Text) - nRedondeoITF, "#,##0.00")
              End If
                  TxtMontoTotal.Text = CCur(txtMontoPagar.Text) + CCur(TxtITF.Text)
              End If
           
           Dim Aux As String
           If InStr(1, CStr(TxtITF), ".", vbTextCompare) > 0 Then
            Aux = CDbl(CStr(Int(TxtITF)) & "." & Mid(CStr(TxtITF), InStr(1, CStr(TxtITF), ".", vbTextCompare) + 1, 2))
           Else
            Aux = CDbl(CStr(Int(TxtITF)))
           End If
            TxtITF.Text = Format(TxtITF.Text, "#0.00")
        Else
            Me.TxtITF = gITF.fgITFCalculaImpuesto(txtMontoPagar.Text)
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtITF.Text))
            If nRedondeoITF > 0 Then
               Me.TxtITF.Text = Format(CCur(Me.TxtITF.Text) - nRedondeoITF, "#,##0.00")
            End If
            TxtMontoTotal = Format(CDbl(Me.txtMontoPagar), "#0.00")
        End If
    Else
            Me.TxtITF = Format(0, "#0.00")
            TxtMontoTotal = Format(Me.txtMontoPagar, "#0.00")
    End If

    Me.TxtITF = Me.TxtITF
    
    fnVarNewSaldoCap = Format(fnVarSaldoCap - fnVarCapitalPagado, "#0.00")
    
    txtSaldoCapitalNuevo.Text = fnVarNewSaldoCap

If CCur(AXDesCon.SaldoCapital) > 0 Then
    fnVarDeuda = CCur(AXDesCon.SaldoCapital) + TxtITF.Text + fnVarInteres + fnVarInteresVencido + fnVarInteresMoratorio + fnVarCostoCustodiaVencida + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarCostoNotificacion
    txtTotalDeuda.Text = Format(fnVarDeuda, "#0.00")
    
    txtCapital.Text = Format(IIf((txtMontoPagar.Text - fnVarCostoNotificacion - fnVarInteres - fnVarInteresVencido - fnVarInteresMoratorio - TxtITF.Text) > 0, (txtMontoPagar.Text - fnVarCostoNotificacion - fnVarInteres - fnVarInteresVencido - fnVarInteresMoratorio - TxtITF.Text), 0), "#0.00") '*** PEAC 20161020

    'txtInteres.Text = Format(fnVarInteres, "#0.00") '*** PEAC 20170321
    txtInteres.Text = Format(fnVarInteres, "#0.00")
    
    txtCostoCus.Text = Format(fnVarCostoCustodiaVencida, "#0.00")
    txtCostoNoti.Text = Format(fnVarCostoNotificacion, "#0.00")

    txtSaldoCapitalNuevo.Text = Format(CCur(AXDesCon.SaldoCapital) - fnVarCapitalPagado, "#0.00")

    Me.txtIntPend.Text = Format(IIf((CDbl(Me.txtMontoMinimoPagar.Text) - CDbl(Me.txtMontoPagar.Text)) < 0, 0, CDbl(Me.txtMontoMinimoPagar.Text) - CDbl(Me.txtMontoPagar.Text)), "#0.00")

End If

If Not IsNumeric(Me.txtInteres.Text) Then
    Me.txtInteres.Text = "0.00"
End If
fnVarIntPagado = IIf(CCur(Me.txtMontoPagar.Text) > CCur(Me.txtInteres.Text), CCur(Me.txtInteres.Text), CCur(Me.txtMontoPagar.Text) - fnVarInteresVencido - fnVarInteresMoratorio)
'*** PEAC 20170321
If fnVarIntPagado < 0 Then
    fnVarIntPagado = 0
'    Me.txtMontoPagar.Text = CCur(Me.txtMontoPagar.Text) + Abs(fnVarIntPagado)
End If
Me.txtInteres.Text = Format(fnVarIntPagado, "#0.00")
'Me.txtMontoPagar.Text = CCur(Me.txtMontoPagar.Text) + Abs(fnVarIntPagado)
End If
End Sub

Private Sub txtMontoPagar_GotFocus()
    fEnfoque txtMontoPagar
End Sub
Private Function ValidaAlGrabar() As Boolean
    ValidaAlGrabar = False
    
    If Not IsNumeric(txtMontoPagar) Then
        txtMontoPagar.Text = "0.00"
    End If
    
    '*** PEAC 20170525
    If Not IsNumeric(txtTotalDeuda) Then
        txtTotalDeuda.Text = "0.00"
    End If
    If Not IsNumeric(TxtMontoTotal) Then
        TxtMontoTotal.Text = "0.00"
    End If
    '*** END PEAC

    'If CDbl(txtMontoPagar) >= CDbl(txtTotalDeuda) Then  'Monto Pagar >= Total Deuda
    If CDbl(TxtMontoTotal) >= CDbl(txtTotalDeuda) Then  'Monto Pagar >= Total Deuda
       MsgBox " Cliente debe Cancelar su contrato ", , " Aviso "
       cmdGrabar.Enabled = False
       Exit Function
    End If
        
    If Val(txtMontoPagar) <= 0 Then
       cmdGrabar.Enabled = False
       Exit Function
    End If
    
        
    If CDbl(txtMontoPagar) < CDbl(Me.txtMora.Text) + CDbl(Me.txtIntVen.Text) Then
        MsgBox "Se debe pagar por lo menos el Interes Vencido y la Mora."
        Exit Function
    End If
    
    
'    If (CDbl(Me.txtCapital.Text)) >= (CCur(AXDesCon.SaldoCapital) * gnPigPorcenPgoCap) Then
'        MsgBox "El monto del Capital es mayor o igual al " & gnPigPorcenPgoCap * 100 & "% del Saldo Capital, por favor realice la operación de Pago Anticipado o Renovación"
'        Exit Function
'    End If
    
    
    
    ValidaAlGrabar = True

End Function

Private Sub txtMontoPagar_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoPagar, KeyAscii)

If KeyAscii = 13 Then
    fnVarInteres = fnVarInteres ' peac 20070820

    If Not ValidaAlGrabar Then Exit Sub

    vSumaCostoCustodia = fnVarCostoCustodia + fnVarCostoCustodiaVencida

    Dim loCalculos As COMNColoCPig.NCOMColPCalculos
    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
        If gdFecSis > AXDesCon.FechaVencimiento Then
        
                If Format(vFecEstado, "yyyymmdd") >= Format(fdVarFecVencimiento, "yyyymmdd") Then
                    vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
                Else
                    'vDiasAdel = 30
                    vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(fdVarFecVencimiento, "dd/mm/yyyy"))
                End If
        
'                If vFecEstado > Format(AXDesCon.FechaVencimiento, "dd/mm/yyyy") Then
'                    vDiasAdel = 30
'                Else
'                    vDiasAdel = DateDiff("d", vFecEstado, Format(AXDesCon.FechaVencimiento, "dd/mm/yyyy"))
'                End If
            Else
                vDiasAdel = DateDiff("d", vFecEstado, Format(gdFecSis, "dd/mm/yyyy"))
        End If

        If gcCredAntiguo = "A" Then
            fnVarInteres = Round(0, 2)
        Else
            
            fnVarInteres = loCalculos.nCalculaInteresAlVencimiento(CCur(AXDesCon.SaldoCapital), fnVarTasaInteres, vDiasAdel)
            fnVarInteres = Round(fnVarInteres, 2)
                        
            fnVarInteres = fnVarInteres + fnIntPendSaldo

        End If

    If CCur(AXDesCon.SaldoCapital) > 0 Then
        'fnVarDeuda = CCur(AXDesCon.SaldoCapital) + TxtITF.Text + fnVarInteres + fnVarInteresVencido + fnVarInteresMoratorio + fnVarCostoCustodiaVencida + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarCostoNotificacion
        fnVarDeuda = CCur(AXDesCon.SaldoCapital) + TxtITF.Text + fnVarInteres + fnVarInteresVencido + fnVarInteresMoratorio + fnVarCostoCustodiaVencida + fnVarImpuesto + fnVarCostoNotificacion
        txtTotalDeuda.Text = Format(fnVarDeuda, "#0.00")
        
        'txtCapital.Text = Format(IIf((txtMontoPagar.Text - (fnVarInteres - fnIntPendPagados) - fnVarInteresVencido - fnVarInteresMoratorio - txtITF.Text) > 0, (txtMontoPagar.Text - (fnVarInteres - fnIntPendPagados) - fnVarInteresVencido - fnVarInteresMoratorio - txtITF.Text), 0), "#0.00")
        'txtCapital.Text = Format(fnVarCapitalPagado, "#0.00")
        txtCapital.Text = Format(IIf((txtMontoPagar.Text - fnVarInteres - fnVarInteresVencido - fnVarInteresMoratorio - TxtITF.Text) > 0, (txtMontoPagar.Text - fnVarInteres - fnVarInteresVencido - fnVarInteresMoratorio - TxtITF.Text), 0), "#0.00")
        
        txtCostoCus.Text = Format(fnVarCostoCustodiaVencida, "#0.00")
        txtCostoNoti.Text = Format(fnVarCostoNotificacion, "#0.00")
        txtSaldoCapitalNuevo.Text = Format(CCur(AXDesCon.SaldoCapital) - fnVarCapitalPagado, "#0.00")
        
        '*** PEAC 20170320
        txtMora.Text = Format(fnVarInteresMoratorio, "#0.00")
        txtIntVen.Text = Format(fnVarInteresVencido, "#0.00")

        'If CDbl(txtMontoPagar.Text) >= CDbl(txtMontoMinimoPagar) And CDbl(txtMontoPagar.Text) <= fnVarInteres Then '*** PEAC 20170630
        If CDbl(txtMontoPagar.Text) >= CDbl(txtMontoMinimoPagar) Then '*** PEAC 20170630
            'fnVarCapitalPagado = 0 '*** PEAC 20170630
            If lnDiasAtraso > 0 Then
                txtInteres.Text = Format(CDbl(txtMontoPagar.Text) - CDbl(txtMontoMinimoPagar), "#0.00")
                'txtIntPend.Text = Format((fnVarInteres - fnIntPendSaldo) - (CDbl(txtMontoPagar.Text) - CDbl(txtMontoMinimoPagar)), "#0.00")
                txtIntPend.Text = Format(fnVarInteres - (CDbl(txtMontoPagar.Text) - CDbl(txtMontoMinimoPagar)), "#0.00")
            Else
                'txtInteres.Text = Format(IIf(CCur(Me.txtMontoPagar.Text) > (fnVarInteres - fnIntPendPagados), (fnVarInteres - fnIntPendPagados), CCur(Me.txtMontoPagar.Text)), "#0.00")
                txtInteres.Text = Format(IIf(CCur(Me.txtMontoPagar.Text) > fnVarInteres, fnVarInteres, CCur(Me.txtMontoPagar.Text)), "#0.00")
                'txtIntPend.Text = Format(fnVarInteres - fnIntPendPagados - CDbl(txtInteres.Text), "#0.00")
                txtIntPend.Text = Format((fnVarInteres) - CDbl(txtInteres.Text), "#0.00")
            End If
            
        End If
        
        If CDbl(txtMontoPagar.Text) >= fnVarInteres + fnVarInteresVencido + fnVarInteresMoratorio Then
            'fnVarCapitalPagado = Format(CDbl(txtMontoPagar.Text) - CDbl(txtMontoMinimoPagar) - fnVarInteres, "#0.00")
            'fnVarCapitalPagado = Format(CDbl(txtMontoPagar.Text) - fnVarInteres, "#0.00")
            txtInteres.Text = Format(fnVarInteres, "#0.00")
            txtIntPend.Text = Format(0, "#0.00")
        End If
               
    End If

    cmdGrabar.Enabled = True
    EnfocaControl cmdGrabar
End If
End Sub

Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
    fnVarTasaPreparacionRemate = loParam.dObtieneColocParametro(gConsColPTasaPreparaRemate)
    fnVarTasaImpuesto = loParam.dObtieneColocParametro(gConsColPTasaImpuesto)
    fnVarTasaCustodiaVencida = loParam.dObtieneColocParametro(gConsColPTasaCustodiaVencida)
    fnVarTasaCustodia = loParam.dObtieneColocParametro(gConsColPTasaCustodia)
    fnVarDiasCambCart = loParam.dObtieneColocParametro(gConsColPDiasCambioCartera)
    gnPigPorcenPgoCap = loParam.dObtieneColocParametro(9050) '*** PEAC 20160920
    gnPigVigMeses = loParam.dObtieneColocParametro(9051) '*** PEAC 20160920
    
    If Me.AXCodCta.Age <> "" Then
        Select Case CInt(Me.AXCodCta.Age)
            Case 1
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3103)
            Case 2
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3104)
            Case 3
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3105)
            Case 4
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3106)
            Case 5
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3107)
            Case 6
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3108)
            Case 7
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3109)
            Case 9
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3111)
            Case 10
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3112)
            Case 12
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3113)
            Case 13
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3114)
            Case 24
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3115)
            Case 25
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3116)
            Case 31
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3117)
        End Select
   End If
    
Set loParam = Nothing
End Sub

Private Sub txtMontoPagar_LostFocus()
    
     If Trim(txtMontoPagar.Text) = "" Then
        txtMontoPagar.Text = "0.00"
    End If
    txtMontoPagar.Text = Format(txtMontoPagar.Text, "#0.00")
    
    Call txtMontoPagar_KeyPress(13)
    
End Sub

Private Sub TxtMontoTotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtMontoTotal.Text = Format(TxtMontoTotal.Text, "#0.00")
    txtMontoPagar_Change
 End If
End Sub

'CTI4 ERS0112020 *****************
Private Sub CmbForPag_Click()
    EstadoFormaPago IIf(CmbForPag.ListIndex = -1, -1, CInt(Trim(Right(IIf(CmbForPag.Text = "", "-1", CmbForPag.Text), 10))))
    If CmbForPag.ListIndex <> -1 Then
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher Then
     
            Dim oformVou As New frmCapRegVouDepBus
            Dim lnTipMot As Integer
            Dim sGlosa As String
            Dim sIF As String
            Dim sVaucher As String
            Dim sPersCod As String
            Dim sNombre As String
            Dim sDireccion As String
            Dim sDocumento As String
                        
            lnTipMot = 20 ' Amortizacion Credito Pignoraticio
            oformVou.iniciarFormularioDeposito CInt(Mid(AXCodCta.NroCuenta, 9, 1)), lnTipMot, sGlosa, sIF, sVaucher, nMontoVoucher, sPersCod, nMovNroRVD, nMovNroRVDPend, sNombre, sDireccion, sDocumento, AXCodCta.NroCuenta
            If Len(sVaucher) = 0 Then Exit Sub
            LblNumDoc.Caption = Trim(Mid(sVaucher, 1, Len(sVaucher) - 10))
            cmdGrabar.Enabled = True
        ElseIf CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            Dim sCuenta As String
            
            sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(gAhoCargoCtaAmortizaPignoNor), sNumTarj, 232, False)
            If Val(Mid(sCuenta, 6, 3)) <> "232" And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            If Len(sCuenta) = 18 Then
                If CInt(Mid(sCuenta, 9, 1)) <> CInt(Mid(AXCodCta.NroCuenta, 9, 1)) Then
                    MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la cuenta a amortizar.", vbOKOnly + vbInformation, App.Title
                End If
            End If
            If Len(sCuenta) = 0 Then txtCuentaCargo.SetFocusAge: Exit Sub
            txtCuentaCargo.NroCuenta = sCuenta
            txtCuentaCargo.Enabled = False
            txtMontoPagar_Change
            cmdGrabar.Enabled = True
            cmdGrabar.SetFocus
        End If
    End If
End Sub
Private Sub EstadoFormaPago(ByVal nFormaPago As Integer)
    LblNumDoc.Caption = ""
    txtCuentaCargo.NroCuenta = ""
    Select Case nFormaPago
        Case -1
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargo.Visible = False
            cmdGrabar.Enabled = True
        Case gColocTipoPagoEfectivo
            txtCuentaCargo.Visible = False
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            lblNroDocumento.Visible = False
            cmdGrabar.Enabled = True
        Case gColocTipoPagoCargoCta
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargo.Visible = True
            txtCuentaCargo.Enabled = True
            txtCuentaCargo.CMAC = gsCodCMAC
            txtCuentaCargo.Prod = Trim(Str(gCapAhorros))
            cmdGrabar.Enabled = False
        Case gColocTipoPagoVoucher
            LblNumDoc.Visible = True
            lblNroDocumento.Visible = True
            txtCuentaCargo.Visible = False
            cmdGrabar.Enabled = False
    End Select
End Sub
Private Function ValidaFormaPago() As Boolean
Dim clsCap As New COMNCaptaGenerales.NCOMCaptaMovimiento
    ValidaFormaPago = False
    If CmbForPag.ListIndex = -1 Then
        MsgBox "No se ha seleccionado la forma de pago. Verifique.", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher And Len(Trim(LblNumDoc.Caption)) = 0 Then
        MsgBox "No se ha seleccionado el voucher correctamente. Verifique.", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher And Len(Trim(LblNumDoc.Caption)) > 0 _
        And CCur(TxtMontoTotal.Text) <> CCur(nMontoVoucher) Then
        MsgBox "El Monto de Transacción debe ser igual al Monto Total. Verifique.", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
    
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta And Len(txtCuentaCargo.NroCuenta) <> 18 Then
        MsgBox "Favor de ingresar la cuenta de ahorros a debitar", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
        
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
        If Not clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, CDbl(TxtMontoTotal.Text)) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
            Exit Function
        End If
    End If
    
    ValidaFormaPago = True
End Function
Private Sub txtCuentaCargo_KeyPress(KeyAscii As Integer)
    Dim sMsg As String
    If KeyAscii = 13 Then sMsg = ValidaCuentaACargo(txtCuentaCargo.NroCuenta)
    If Len(sMsg) > 0 Then
        MsgBox sMsg, vbInformation, "¡Aviso!"
        txtCuentaCargo.SetFocus
        Exit Sub
    End If
    If Len(txtCuentaCargo.NroCuenta) = 18 Then
        If CInt(Mid(txtCuentaCargo.NroCuenta, 9, 1)) <> CInt(Mid(AXCodCta.NroCuenta, 9, 1)) Then
            MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la cuenta a amortizar.", vbOKOnly + vbInformation, App.Title
        End If
    End If
    ObtieneDatosCuenta txtCuentaCargo.NroCuenta
End Sub
Private Function ValidaCuentaACargo(ByVal psCuenta As String) As String
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim sMsg As String
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    sMsg = clsCap.ValidaCuentaOperacion(psCuenta)
    ValidaCuentaACargo = sMsg
End Function
Private Sub ObtieneDatosCuenta(ByVal psCuenta As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsCta As ADODB.Recordset
    Dim rsV As ADODB.Recordset
    Dim lnTpoPrograma As Integer
    Dim lsTieneTarj As String
    Dim lbVistoVal As Boolean
    
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsV = New ADODB.Recordset
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(psCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
        If sNumTarj = "" Then
            If Mid(psCuenta, 6, 3) = "232" And lnTpoPrograma <> 1 Then
                Set rsV = clsMant.ValidaTarjetizacion(psCuenta, lsTieneTarj)
                If rsV.RecordCount > 0 Then
                    Dim tipoCta As Integer
                    tipoCta = rsCta("nPrdCtaTpo")
                    If tipoCta = 0 Or tipoCta = 2 Then
                        Dim rsCli As New ADODB.Recordset
                        Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
                        Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
                        Dim bExitoSol As Integer
                        Set rsCli = clsCli.GetPersonaCuenta(psCuenta, gCapRelPersTitular)
                        nRespuesta = oSolicitud.SolicitarVistoAtencionSinTarjeta(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gAhoCargoCtaAmortizaPignoNor))
                    
                        If nRespuesta = 1 Then '1:Tiene Visto de atencion sin tarjeta pendiente de autorizar
                             MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                             Exit Sub
                        End If
                        If nRespuesta = 2 Then '2:Tiene visto aceptado
                            MsgBox "La solicitud de atención sin tarjeta fue Aprobada, proceda con la atención", vbInformation, "Aviso"
                        End If
                        If nRespuesta = 3 Then '3:Tiene visto rechazado
                           If MsgBox("La solicitud de atención sin tarjeta fue RECHAZADA. ¿Desea realizar una nueva solicitud?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                Call oSolicitud.RegistrarVistoDeUsuario(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gAhoCargoCtaAmortizaPignoNor))
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        If nRespuesta = 4 Then '4:Se permite registrar la solicitud
                            Dim mensaje As String
                            If lsTieneTarj = "SI" Then
                                mensaje = "El Cliente posee tarjeta. Para continuar deberá registrar el Motivo de Autorización y comunicar al Coordinador o Jefe de Operaciones para su Aprobación. ¿Desea Continuar?"
                            Else
                                mensaje = "El Cliente NO posee tarjeta activa. Para continuar deberá registrar el Motivo de Autorización y comunicar al Coordinador o Jefe de Operaciones para su Aprobación. ¿Desea Continuar?"
                            End If
                        
                            If MsgBox(mensaje, vbInformation + vbYesNo, "Aviso") = vbYes Then
                        
                                bExitoSol = frmCapMotivoAutorizacion.Inicio(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gAhoCargoCtaAmortizaPignoNor))
                                If bExitoSol > 0 Then
                                    MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                                End If
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        lbVistoVal = loVistoElectronico.Inicio(5, gAhoCargoCtaAmortizaPignoNor)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema"
                            Exit Sub
                        End If
                    End If
                ElseIf lsTieneTarj = "NO" And rsV.RecordCount > 0 Then
                    If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? ", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'add by marg ers 065-2017
                        lbVistoVal = loVistoElectronico.Inicio(5, gAhoCargoCtaAmortizaPignoNor)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema" 'ADD BY MARG ERS065-2017
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            End If
        End If
        txtCuentaCargo.Enabled = False
        txtMontoPagar_Change
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
    End If
End Sub
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    Set R = oCons.RecuperaConstantes(gColocTipoPago, , , 4)
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(R, CmbForPag)
    Set loVistoElectronico = New frmVistoElectronico
    Exit Sub
ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub
'END CTI4

'JOEP20210916 campana prendario
Private Sub CampPrendVerificaCampanas(ByVal pcCtaCod As String, ByVal pdFechaSis As String, ByVal pnModulo As Integer, Optional ByVal pnRenva As Integer)
    Dim oCampPrend As COMDColocPig.DCOMColPContrato
    Dim rsCampPrend As ADODB.Recordset
    Set oCampPrend = New COMDColocPig.DCOMColPContrato
    
    Set rsCampPrend = oCampPrend.CampPrendarioDesbCampa(pcCtaCod, pdFechaSis, pnModulo, pnRenva)
    If Not (rsCampPrend.BOF And rsCampPrend.EOF) Then
        lblCampRetenPrendAmor.Caption = rsCampPrend!cResultado
        lblCampRetenPrendAmor.Visible = True
        txtCampReteAmor.Text = rsCampPrend!nCampana
    Else
        lblCampRetenPrendAmor.Visible = False
        lblCampRetenPrendAmor.Caption = ""
        txtCampReteAmor.Text = 0#
    End If
    Set oCampPrend = Nothing
    RSClose oCampPrend
End Sub
'JOEP20210916 campana prendario
