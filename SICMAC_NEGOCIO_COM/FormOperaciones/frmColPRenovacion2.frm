VERSION 5.00
Begin VB.Form frmColPRenovacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Renovación de Crédito"
   ClientHeight    =   8325
   ClientLeft      =   720
   ClientTop       =   2115
   ClientWidth     =   7995
   Icon            =   "frmColPRenovacion2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   5640
      TabIndex        =   4
      Top             =   7870
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4440
      TabIndex        =   3
      Top             =   7870
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   6795
      TabIndex        =   5
      Top             =   7870
      Width           =   975
   End
   Begin VB.Frame fraContenedor 
      Height          =   7275
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   7785
      Begin VB.TextBox txtCampRete 
         Height          =   285
         Left            =   5640
         TabIndex        =   52
         Text            =   "0.00"
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   7110
         Picture         =   "frmColPRenovacion2.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Buscar ..."
         Top             =   270
         Width           =   420
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Plazo Nuevo"
         Height          =   1290
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   5880
         Width           =   7425
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
            Left            =   3840
            TabIndex        =   49
            Top             =   840
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
            Left            =   3840
            TabIndex        =   37
            Top             =   480
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
            Left            =   3840
            TabIndex        =   36
            Top             =   120
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
            Left            =   1080
            TabIndex        =   45
            Top             =   120
            Width           =   1215
         End
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
            Left            =   1080
            TabIndex        =   34
            Top             =   480
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
            Left            =   1080
            TabIndex        =   33
            Top             =   840
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
            Left            =   6120
            TabIndex        =   29
            Top             =   900
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
            Left            =   6120
            TabIndex        =   28
            Top             =   555
            Width           =   1215
         End
         Begin VB.ComboBox cboPlazoNuevo 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmColPRenovacion2.frx":040C
            Left            =   1440
            List            =   "frmColPRenovacion2.frx":0413
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
            Left            =   6120
            MaxLength       =   9
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblCostoNoti 
            AutoSize        =   -1  'True
            Caption         =   "Costo Notificación"
            Height          =   195
            Index           =   5
            Left            =   2400
            TabIndex        =   41
            Top             =   960
            Width           =   1290
         End
         Begin VB.Label lblCostoCus 
            AutoSize        =   -1  'True
            Caption         =   "Costo Custodia"
            Height          =   195
            Index           =   4
            Left            =   2400
            TabIndex        =   40
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label lblIntVen 
            AutoSize        =   -1  'True
            Caption         =   "Interés Vencido"
            Height          =   195
            Index           =   3
            Left            =   2400
            TabIndex        =   39
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label lblCapital 
            AutoSize        =   -1  'True
            Caption         =   "Capital"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   32
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lblMora 
            AutoSize        =   -1  'True
            Caption         =   "Mora"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   31
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lblInteres 
            AutoSize        =   -1  'True
            Caption         =   "Interés"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   30
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Monto Pagar"
            Height          =   195
            Left            =   5115
            TabIndex        =   27
            Top             =   960
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ITF"
            Height          =   195
            Left            =   5145
            TabIndex        =   26
            Top             =   645
            Width           =   240
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Monto "
            Height          =   195
            Index           =   12
            Left            =   5130
            TabIndex        =   13
            Top             =   270
            Width           =   495
         End
      End
      Begin VB.Frame fraContenedor 
         Height          =   1365
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   4440
         Width           =   7485
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
            TabIndex        =   42
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
         Begin VB.Label lblProxFechaPago 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2880
            TabIndex        =   44
            Top             =   1080
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.Label lblInteres 
            AutoSize        =   -1  'True
            Caption         =   "Amortización Int.:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   43
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
         Top             =   720
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6588
      End
      Begin VB.Label lblCampRetenPrend 
         Caption         =   "Campaña y Retencion"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   3960
         TabIndex        =   53
         Top             =   120
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label lblCampRetenPrendAmor 
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Visible         =   0   'False
         Width           =   3135
      End
   End
   Begin VB.Frame fraFormaPago 
      Height          =   600
      Left            =   135
      TabIndex        =   46
      Top             =   7250
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
         TabIndex        =   47
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
         TabIndex        =   48
         Top             =   200
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblFormaPago 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
         Height          =   195
         Left            =   180
         TabIndex        =   50
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
      Top             =   8000
      Width           =   2280
   End
End
Attribute VB_Name = "frmColPRenovacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* RENOVACION  DE CONTRATO PIGNORATICIO
'Archivo:  frmColPRenovacion.frm
'LAYG   :  10/07/2001.
'Resumen:  Nos permite renovar un contrato cambiando asi su fecha
'          de Vencimiento hacia otra
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
Dim fnVarTopRenovaciones As Integer
Dim fnVarTopRenovacionesNuevo As Integer

Dim fnVarSaldoCap As Currency
Dim fnVarValorTasacion As Currency
Dim fnVarPlazo As Integer
Dim fdVarFecVencimiento As Date
Dim fnVarEstado As ColocEstado
Dim fnVarNroRenovacion As Integer

Dim fnVarNewSaldoCap As Currency
Dim fnVarNewPlazo As Integer
Dim fsVarNewFecVencimiento As String
Dim fnVarCapitalPagado As Currency   ' Capital a Pagar
Dim fnVarFactor As Double
Dim fnVarInteresVencido As Currency
Dim fnVarInteresMoratorio As Currency
Dim fnVarInteres As Currency
Dim fnVarCostoCustodia As Currency
Dim fnVarCostoCustodiaVencida As Currency
Dim fnVarImpuesto As Currency
Dim fnVarCostoPreparacionRemate As Double

Dim fnVarInteCompVencido As Currency

Dim fnVarDiasAtraso As Double, fnVarDiasAtrasoReal As Double
Dim vDiasAtrasoReal As Double
Dim vSumaCostoCustodia As Double
Dim fnVarDeuda As Currency
Dim fnVarGastoCorrespondencia As Currency

'*********
Dim fnVarMontoMinimo As Currency
Dim fnVarMontoAPagar As Currency

Dim fnVarCostoNotificacion As Currency 'ARCV 14-03-2007

Dim fnVarEstUltProcRem As Integer 'DAOR 20070714
Dim fsColocLineaCredPig As String, vFecEstado As Date ' PEAC 20070813
Dim vDiasAdel As Integer, vInteresAdel As Double, vMontoCol As Double ' PEAC 20070813
Dim gcCredAntiguo As String  ' peac 20070923
Dim gnNotifiAdju As Integer  ' peac 20080515
Dim gnNotifiCob As Integer  ' *** PEAC 20080715
Dim lsTpoCredCod As String
Dim nRedondeoITF As Double 'BRGO 20110914
Dim fnDiasFer As Integer '20141226 ERS170-2014
Dim gnPigVigMeses As Integer '*** PEAC 20160920
Dim gnPigPorcenPgoCap As Double '*** PEAC 20160920
Dim vCapitalAdel As Double '*** PEAC 20160921
Dim fnCredRevolAntNue As Boolean '*** PEAC 20161105 - 1= antiguo, 0 ó null = nuevo
Dim fnPgoAdelInt As Double '*** Peac 20161116
Dim lsFechaVenc As String
Dim lsFechaVencImp As String
Dim nPagoIntVenMor As Double '*** PEAC 20170330
Dim fnIntPendSaldo As Double '*** PEAC 20170522
Dim fnIntComVenPgdo As Double '*** PEAC 201
Dim fnIntMoraPgdo As Double '*** PEAC 20170522
Dim ldfecVigDomFer As Date '*** PEAC 20190520
Dim ldfecDesembolso As Date '*** PEAC 20190520
Dim fnDiasFerImp As Integer '*** PEAC 20190711
Dim fdVarFecVectoUnico As Date '*** PEAC 20190723
Private nMontoVoucher As Currency 'CTI4 ERS0112020
Dim nMovNroRVD As Long, nMovNroRVDPend As Long 'CTI4 ERS0112020
Dim sNumTarj As String 'CTI4 ERS0112020
Dim loVistoElectronico As frmVistoElectronico 'CTI4 ERS0112020
Dim nRespuesta As Integer 'CTI4 ERS0112020
'JOEP20210916 campana prendario
Dim nCampPrenRetencion As Integer
Dim nCampPrenCampana As Integer
Dim pgnCumpleCampna As Integer
Dim pgTasaOriginal As Double
'JOEP20210916 campana prendario
'Dim ventana As Integer 'MADM 20090928
Public Sub Inicio(ByVal pnOpeCod As Long, ByVal psOpeDesc As String, _
        ByVal psPersCodCMAC As String, ByVal psNomCmac As String)

    fnVarOpeCod = pnOpeCod
    fsVarOpeDesc = psOpeDesc
    fsVarPersCodCMAC = psPersCodCMAC
    fsVarNombreCMAC = psNomCmac
    
    Select Case fnVarOpeCod
        Case gColPOpeRenovacEFE
            'txtDocumento.Visible = false
        Case gColPOpeRenovacCHQ
            'txtDocumento.Visible = True
    '    Case Else
    '        txtDocumento.Visible = False
    'Add by GITU 10-07-2013
        Case "122700"
            cmdgrabar.Visible = False
            TxtMontoTotal.Enabled = False
    'End Gitu
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
    txtPlazoActual.Text = ""
    txtSaldoCapitalNuevo.Text = Format(0, "#0.00")
    If cboPlazoNuevo.ListCount > 0 Then
     cboPlazoNuevo.ListIndex = 0
    End If
    txtMontoPagar.Text = Format(0, "#0.00")
    TxtMontoTotal.Text = Format(0, "#0.00")
    TxtITF.Text = Format(0, "#0.00")
    'WIOR 20130917 ************************

    txtCapital.Text = Format(0, "#0.00")
    TxtMora.Text = Format(0, "#0.00")
    Txtinteres.Text = Format(0, "#0.00")
    txtIntVen.Text = Format(0, "#0.00")
    Me.txtCostoCus.Text = ""
    Me.txtCostoNoti.Text = ""
    
    'WIOR FIN *****************************
    fnDiasFer = 0 'RECO20141226 ERS170-2014
       
    Me.lblProxFechaPago.Caption = ""
    Me.lblProxFechaPago.Visible = False
    
    lsFechaVenc = "__/__/____"
    CmbForPag.ListIndex = -1 'CTI4 ERS0112020
    txtCuentaCargo.NroCuenta = "" 'CTI4 ERS0112020
    LblNumDoc.Caption = "" 'CTI4 ERS0112020
    cmdgrabar.Enabled = False 'CTI4 ERS0112020
    sNumTarj = "" 'CTI4 ERS0112020
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call BuscaContrato(AXCodCta.NroCuenta)
End Sub

'Busca el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim loCalculos As COMNColoCPig.NCOMColPCalculos
Dim loPigFunc As COMDColocPig.DCOMColPFunciones
Dim lnDeuda As Currency, lnMinimoPagar As Currency
Dim lnDiasAtraso  As Integer
Dim lsFecVenTemp As String
Dim lsmensaje As String
'----- MADM
Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim Rf As ADODB.Recordset
'----- MADM

'On Error GoTo ControlError

    '*** PEAC 20161104
    '*** obtiene si credito es nuevo o antiguo para aplicar cambios
    '*** de cred revolventes
    Dim loParam As COMDColocPig.DCOMColPCalculos
    Set loParam = New COMDColocPig.DCOMColPCalculos
    
    fnCredRevolAntNue = loParam.dObtieneParamCredRevolNueAnt(Me.AXCodCta.NroCuenta)

    'Valida Contrato
    gITF.fgITFParamAsume (Mid(psNroContrato, 4, 2)), Mid(psNroContrato, 6, 3)
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida

        Set lrValida = loValContrato.nValidaRenovacionCredPignoraticio(psNroContrato, gdFecSis, 0, gsCodUser, lsmensaje, fnCredRevolAntNue)
        If Trim(lsmensaje) <> "" Then
        '*** PEAC 20161214
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
    ' Asigna Valores a Variables del Form
          
'Dim loParam As COMDColocPig.DCOMColPCalculos
'Set loParam = New COMDColocPig.DCOMColPCalculos
    '******** madm 20091204 **********************************

If Me.AXCodCta.Age <> "" Then
        'RECO20140623 ERS081-2014*******************************************
'        Select Case CInt(Me.AXCodCta.Age)
'            Case 1
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3103)
'            Case 2
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3104)
'            Case 3
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3105)
'            Case 4
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3106)
'            Case 5
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3107)
'            Case 6
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3108)
'            Case 7
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3109)
'            Case 9
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3111)
'            Case 10
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3112)
'            Case 12
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3113)
'            Case 13
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3114)
'            Case 24
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3115)
'            Case 25
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3116)
'            Case 31
'             fnVarCostoNotificacion = loParam.dObtieneColocParametro(3117)
'        End Select
        'RECO20140722 ERS114-2014******************************************

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
        'RECO FIN**********************************************************
        'fnVarCostoNotificacion = loParam.dObtieneParamPignoCostoNotif("COSTO NOTIFIC", Me.AXCodCta.Age)
        'RECO FIN***********************************************************
   End If
   ' fnVarCostoNotificacion = loParam.dObtieneColocParametro(3103) 'ARCV 14-03-2007
   '*********** madm *****************************************
Set loParam = Nothing

    fnVarPlazo = lrValida!nPlazo
    fnVarSaldoCap = Format(lrValida!nSaldo, "#0.00")
    fnVarValorTasacion = lrValida!nTasacion
    nPagoIntVenMor = lrValida!nPagosIntVenMora
    
    fnVarTasaInteresVencido = lrValida!nTasaIntVenc
    fnVarTasaInteresMoratorio = lrValida!nTasaIntMora
    
    fnVarEstado = lrValida!nPrdEstado
    vFecEstado = lrValida!dPrdEstado ' PEAC 20070813
    fnVarSaldoCap = lrValida!nMontoCol ' PEAC 20070813
    fnVarTasaInteres = lrValida!nTasaInteres
    fdVarFecVencimiento = Format(lrValida!dVenc, "dd/mm/yyyy")
    fdVarFecVectoUnico = Format(lrValida!dVenc, "dd/mm/yyyy")
    fnVarNroRenovacion = lrValida!nNroRenov
    
    '*** PEAC 20071205 - no importa si antes era de 14 o 7 dias, para adelante será de 30
    'fnVarNewPlazo = lrValida!nPlazo
    fnVarNewPlazo = Val(cboPlazoNuevo.Text)
    '***************************************
    
    gcCredAntiguo = lrValida!cCredB 'PEAC 20070925
    gnNotifiAdju = lrValida!nCodNotifiAdj 'PEAC 20080515
    gnNotifiCob = lrValida!nCodNotifiCob  ' PEAC 20080715
    
    fnIntPendSaldo = lrValida!nPagosPendIntSaldo
    fnPgoAdelInt = lrValida!nPagosPendIntPagados '*** PEAC 20161117
    
    fnVarEstUltProcRem = lrValida!nEstUltProcRem 'DAOR 20070714
    
    ldfecVigDomFer = lrValida!dfecvigdomfer '*** PEAC 20190520
    ldfecDesembolso = lrValida!dDesembolso '*** PEAC 20190520
    
    
    'Muestra Datos
    If fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, False) Then
    
    End If

    'If lrValida!codigoant <> "" Then
    
    '*** PEAC 20080412 - el numero de renvaciones no tiene limite segun Acta 070-2008/TI-D
'        If lrValida!nNroRenov >= fnVarTopRenovaciones Then ' controla el Nro Maximo Renovaciones
'            MsgBox "Contrato llego al Nro Maximo de Renovaciones  (" & Trim(Str(fnVarTopRenovaciones)) & ") ", vbInformation, "Aviso"
'            Exit Sub
'        End If
    '*** FIN PEAC
        
'    'Else
'        If lrValida!nNroRenov >= fnVarTopRenovacionesNuevo Then ' controla el Nro Maximo Renovaciones
'            MsgBox "Contrato llego al Nro Maximo de Renovaciones  (" & Trim(Str(fnVarTopRenovaciones)) & ") ", vbInformation, "Aviso"
'            Exit Sub
'        End If
'    'End If
    

    ' Fecha de Vencimiento es feriado - OJO
    lsFecVenTemp = fdVarFecVencimiento
    Set loPigFunc = New COMDColocPig.DCOMColPFunciones
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Sub
    End If
    
    '*** PEAC 20190508
    Dim rsFecVenvFeri As ADODB.Recordset
    Set rsFecVenvFeri = New ADODB.Recordset
    Set rsFecVenvFeri = loPigFunc.dObtieneFechaVencimientoFeriado(Format(fdVarFecVencimiento, "yyyyMMdd"), Me.AXCodCta.Age, Format(gdFecSis, "yyyyMMdd"))

'   COMENTADO POR PEAC - este proceso se envio a la funcion dObtieneFechaVencimientoFeriado
'    If loPigFunc.dVerSiFeriado(lsFecVenTemp, lsmensaje) = True Then
'        Dim nDiasFerTmp As Integer
'        Do While True
'            lsFecVenTemp = DateAdd("d", 1, lsFecVenTemp)
'
'            fnDiasFer = fnDiasFer + 1 'RECO20141226 ERS170-2014
'            nDiasFerTmp = fnDiasFer
'            If Not loPigFunc.dVerSiFeriado(lsFecVenTemp, lsmensaje) = True Then
'                If Trim(lsmensaje) <> "" Then
'                     MsgBox lsmensaje, vbInformation, "Aviso"
'                     Exit Sub
'                End If
'                Exit Do
'            End If
'        Loop
'        If lsFecVenTemp = gdFecSis Then
'            fdVarFecVencimiento = lsFecVenTemp
'        Else
'            fnDiasFer = 0 'RECO20141226 ERS170-2014
'        End If
''        If gdFecSis <= fdVarFecVencimiento Then 'RECO20141226 ERS170-2014
''            fnDiasFer = 0 'RECO20141226 ERS170-2014
''        End If 'RECO20141226 ERS170-2014
'    End If
'    Set loPigFunc = Nothing
    
    'lnDiasAtraso = DateDiff("d", Format(lrValida!dVenc, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
    lnDiasAtraso = rsFecVenvFeri!nDiasAtraso
    
    fdVarFecVencimiento = rsFecVenvFeri!dNuevaFecVenc
    
    If Format(ldfecDesembolso, "yyyyMMdd") >= Format(ldfecVigDomFer, "yyyyMMdd") Then
        fnDiasFer = 0
    Else
        fnDiasFer = rsFecVenvFeri!nCuentaDomFer
    End If
    
    'fnDiasFer = rsFecVenvFeri!nCuentaDomFer
    
'**** fin PEAC
    
    'vDiasAtrasoReal = vDiasAtraso
    Me.txtDiasAtraso = Val(lnDiasAtraso)
    txtNroRenovacion.Text = Val(lrValida!nNroRenov) + 1
    txtPlazoActual.Text = Val(lrValida!nPlazo)
    
    Me.txtPgoAdelInt.Text = Format(fnPgoAdelInt, "#0.00") '*** PEAC 20161117
    
     'JOEP20210914 Campana Prendario
    Call CampPrendVerificaRetencion(psNroContrato, gdFecSis)
    If lblCampRetenPrend.Visible = False Then
        Call CampPrendVerificaCampanas(psNroContrato, gdFecSis, 1, txtNroRenovacion.Text)
    End If
    'JOEP20210914 Campana Prendario
    
    'vDiasAtrasoReal = vDiasAtraso
'    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
'        lnDeuda = loCalculos.nCalculaDeudaPignoraticio(fnVarSaldoCap, fdVarFecVencimiento, fnVarValorTasacion, _
'                fnVarTasaInteresVencido, fnVarTasaCustodiaVencida, fnVarTasaImpuesto, fnVarEstado, _
'                fnVarTasaPreparacionRemate, gdFecSis)
'
'        lnMinimoPagar = loCalculos.nCalculaMinimoPagar(fnVarSaldoCap, fnVarTasaInteres, fnVarPlazo, fnVarTasaCustodia, _
'                fdVarFecVencimiento, fnVarValorTasacion, fnVarTasaInteresVencido, fnVarTasaCustodiaVencida, _
'                fnVarTasaImpuesto, fnVarEstado, fnVarTasaPreparacionRemate, gdFecSis)
'    Set loCalculos = Nothing
        
    ' Calcula el Monto Total de la Deuda
    fgCalculaDeuda lnDiasAtraso 'Se añadió parámetro RIRO 20200406
    
    fgCalculaMinimoPagar
        
    ' Muestra datos
    
    txtTotalDeuda.Text = Format(fnVarDeuda, "#0.00")
    fnVarGastoCorrespondencia = 0
    txtMontoMinimoPagar.Text = Format(fnVarMontoMinimo + fnVarGastoCorrespondencia, "#0.00")
    'txtMontoPagar.Text = Format(fnVarMontoMinimo + fnVarGastoCorrespondencia, "#0.00")
   
     Dim v1 As Double, v2 As Double, v3 As Double
     
     
    
    'TxtMontoTotal.Text = Format((CDbl(fnVarMontoMinimo) + CDbl(fnVarGastoCorrespondencia)) / (1 - gITF.gnITFPorcent), "#0.00")
    TxtMontoTotal.Text = Format((CDbl(fnVarMontoMinimo) + CDbl(fnVarGastoCorrespondencia)) / (1 - gITF.gnITFPorcent), "#0.00")
    txtSaldoCapitalNuevo.Text = "0.00"
    cboPlazoNuevo.Enabled = False 'True
    Dim i As Integer
    For i = 1 To cboPlazoNuevo.ListCount
        If cboPlazoNuevo.List(i - 1) = (lrValida!nPlazo) Then
            cboPlazoNuevo.ListIndex = i - 1
            cboPlazoNuevo_Click
            Exit For
        End If
    Next i
    
    'cboPlazoNuevo.Clear
    'cboPlazoNuevo.AddItem (lrValida!nPlazo)
    'cboPlazoNuevo.ListIndex = 0
    'cboPlazoNuevo.Enabled = False
    'txtMontoPagar.Enabled = True
    AXCodCta.Enabled = False
    'cboPlazoNuevo.SetFocus
    
  
    
    Set lrValida = Nothing
        
    '----- ITF -----
     If gITF.gbITFAplica Then
        If Not gITF.gbITFAsumidocreditos Then
            'Me.TxtITF = Format(gITF.fgITFCalculaImpuestoIncluido(TxtMontoTotal), "#0.00")
            'txtMontoPagar = Format(CDbl(Me.TxtMontoTotal) - CDbl(Me.TxtITF), "#0.00")
            'txtMontoPagar.Text = Format(gITF.fgITFCalculaImpuestoIncluido(TxtMontoTotal), "#0.00")
            
            'Me.TxtITF = Format(CDbl(Me.TxtMontoTotal) - CDbl(Me.txtMontoPagar), "#0.00")
           txtMontoPagar = Format(txtMontoMinimoPagar.Text, "#0.00")
           Me.TxtITF = gITF.fgITFCalculaImpuesto(txtMontoPagar)
           '*** BRGO 20110908 ************************************************
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtITF.Text))
            If nRedondeoITF > 0 Then
               Me.TxtITF.Text = Format(CCur(Me.TxtITF.Text) - nRedondeoITF, "#,##0.00")
            End If
            '*** END BRGO
           Me.TxtMontoTotal = Format(Val(txtMontoPagar.Text) + Val(TxtITF.Text))
        Else
            Me.TxtITF = gITF.fgITFCalculaImpuesto(TxtMontoTotal)
           '*** BRGO 20110908 ************************************************
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtITF.Text))
            If nRedondeoITF > 0 Then
               Me.TxtITF.Text = Format(CCur(Me.TxtITF.Text) - nRedondeoITF, "#,##0.00")
            End If
            '*** END BRGO
            txtMontoPagar = Format(CDbl(Me.TxtMontoTotal), "#0.00")
        End If
        
    Else
        Me.TxtITF = Format(0, "#0.00")
        txtMontoPagar = Format(Me.TxtMontoTotal, "#0.00")
    End If

'*** PEAC 20190509
'*** verifica que prox fecha de venc no sea domingo ni feriado
    Dim rsFecVenvFeri1 As ADODB.Recordset
    Set rsFecVenvFeri1 = New ADODB.Recordset
    Set rsFecVenvFeri1 = loPigFunc.dObtieneFechaVencimientoFeriado(Format(gdFecSis + fnVarNewPlazo, "yyyyMMdd"), Me.AXCodCta.Age, Format(gdFecSis, "yyyyMMdd"))

'*** PEAC 20170314
'validar si es l ultimo pago

'*** PEAC 20190521
fnDiasFerImp = 0
If Format(ldfecDesembolso, "yyyyMMdd") >= Format(ldfecVigDomFer, "yyyyMMdd") Then
    lsFechaVenc = Format$(rsFecVenvFeri1!dNuevaFecVenc, "dd/mm/yyyy") '*** PEAC 20190510
Else
    lsFechaVenc = Format$(gdFecSis + fnVarNewPlazo, "dd/mm/yyyy") '*** PEAC 20190510
End If
fnDiasFerImp = rsFecVenvFeri1!nCuentaDomFer
lsFechaVencImp = lsFechaVenc

'lsFechaVenc = Format$(gdFecSis + fnVarNewPlazo, "dd/mm/yyyy") '*** PEAC 20190510
'lsFechaVenc = Format$(rsFecVenvFeri1!dNuevaFecVenc, "dd/mm/yyyy") '*** PEAC 20190510

'*** FIN PEAC

Set loPigFunc = Nothing '*** PEAC 20190510

Me.lblProxFechaPago.Caption = "PROXIMA FECHA DE PAGO: " & lsFechaVenc
Me.lblProxFechaPago.Visible = True

'*** PEAC 20080528
fnVarDeuda = CCur(AXDesCon.SaldoCapital) + vInteresAdel + fnVarInteresVencido + fnVarCostoCustodiaVencida _
        + fnVarImpuesto + fnVarGastoCorrespondencia + fnVarInteresMoratorio + fnVarCostoNotificacion + TxtITF.Text
'*** PEAC 20160921

'fnVarDeuda = CCur(AXDesCon.SaldoCapital) + vCapitalAdel + vInteresAdel + fnVarInteresVencido + fnVarCostoCustodiaVencida _
'        + fnVarImpuesto + fnVarGastoCorrespondencia + fnVarInteresMoratorio + fnVarCostoNotificacion + TxtITF.Text


'fnVarDeuda = CCur(AXDesCon.SaldoCapital) + TxtITF.Text + fnVarInteres + fnVarInteresVencido + fnVarCostoCustodiaVencida + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarCostoNotificacion
txtTotalDeuda.Text = Format(fnVarDeuda - TxtITF.Text, "#0.00")

'txtCapital.Text = Format(TxtMontoTotal.Text - fnVarCostoNotificacion - vInteresAdel - fnVarInteresVencido - TxtITF.Text, "#0.00")
'txtCapital.Text = Format(TxtMontoTotal.Text - fnVarCostoNotificacion - vInteresAdel - fnVarInteresMoratorio - fnVarInteresVencido - TxtITF.Text, "#0.00")

txtCapital.Text = Format(TxtMontoTotal.Text - fnVarCostoNotificacion - vInteresAdel - fnVarInteresMoratorio - fnVarInteresVencido - TxtITF.Text, "#0.00") '*** PEAC 20161019

TxtMora.Text = Format(fnVarInteresMoratorio, "#0.00")
txtIntVen.Text = Format(fnVarInteresVencido, "#0.00")
Txtinteres.Text = Format(vInteresAdel, "#0.00")
txtCostoCus.Text = Format(fnVarCostoCustodiaVencida, "#0.00")
txtCostoNoti.Text = Format(fnVarCostoNotificacion, "#0.00")
    
txtSaldoCapitalNuevo.Text = Format(CCur(AXDesCon.SaldoCapital) - vCapitalAdel - fnVarCapitalPagado, "0#.00")
'*** FIN PEAC 20080528
    
'    If gITF.gbITFAplica Then
'        If Not gITF.gbITFAsumidoCreditos Then
'            Me.TxtITF = Format(gITF.fgITFCalculaImpuestoIncluido(Me.txtMontoPagar), "#0.00")
'            Me.TxtMontoTotal = CCur(Me.txtMontoPagar) + CCur(Me.TxtITF)
'        Else
'            Me.TxtITF = Format(gITF.fgITFCalculaImpuesto(TxtMontoTotal), "#0.00")
'            Me.TxtMontoTotal = Format(txtMontoPagar, "#0.00")
'        End If
'    Else
'        Me.TxtITF = Format(0, "#0.00")
'        txtMontoPagar = Format(Me.TxtMontoTotal, "#0.00")
'    End If
    '---------------

    cmdgrabar.Enabled = True
   ' Me.cboPlazoNuevo.SetFocus
    'cmdGrabar.SetFocus
    
    '************ firma madm
        'If ventana = 0 Then
         Set lafirma = New frmPersonaFirma
         Set ClsPersona = New COMDPersona.DCOMPersonas
        
         Set Rf = ClsPersona.BuscaCliente(gColPigFunciones.vcodper, BusquedaCodigo)
         If Not Rf.BOF And Not Rf.EOF Then
            If Rf!nPersPersoneria = 1 Then
            'Call frmPersonaFirma.Inicio(Trim(gColPigFunciones.vcodper), Mid(gColPigFunciones.vcodper, 4, 2), False, True) 'jato 20210322
            Call frmPersonaFirma.Inicio(Trim(gColPigFunciones.vcodper), Mid(gColPigFunciones.vcodper, 4, 2), False, False) 'jato 20210322
        End If
    End If
         Set Rf = Nothing
       '  End If
       '  ventana = 0
        '************ firma madm

    AXCodCta.Enabled = False
    CmbForPag.Enabled = True 'CTI4 ERS0112020
    CmbForPag.ListIndex = IndiceListaCombo(CmbForPag, 1) 'CTI4 ERS0112020
    
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdBuscar_Click()

Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersContrato As COMDColocPig.DCOMColPContrato
Dim lrContratos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

'On Error GoTo ControlError
Call cmdCancelar_Click 'WIOR 20130917
Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov

If Trim(lsPersCod) <> "" Then
    Set loPersContrato = New COMDColocPig.DCOMColPContrato
        Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstados, Mid(gsCodAge, 4, 2))
    Set loPersContrato = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    'Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrContratos) Comentado Por RIRO 20130724 SEGUN ERS101-2013
    Set loCuentas = frmCuentasPersona.Inicio(lsPersNombre, lrContratos) 'RIRO 20130724 SEGUN ERS101-2013
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing
'ventana = 1
'Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
'        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Cancela el proceso actual y permite inicializar ls variables para otro proceso
Private Sub cmdCancelar_Click()
    Limpiar
    cmdgrabar.Enabled = False
    cboPlazoNuevo.Enabled = False
   ' txtMontoPagar.Enabled = False
   
    AXCodCta.Enabled = True
    CmbForPag.Enabled = False 'CTI4 ERS0112020
    AXCodCta.SetFocusCuenta
End Sub

'Actualiza los cambios en la basede datos
Private Sub cmdGrabar_Click()
'WIOR 20130301 ******************
Dim PersonaPago() As Variant
Dim nI As Integer
Dim fnCondicion As Integer
Dim regPersonaRealizaPago As Boolean
nI = 0
gnMovNro = 0
'WIOR ***************************
'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarRen As COMNColoCPig.NCOMColPContrato
Dim loImprime As COMNColoCPig.NCOMColPImpre
Dim loPrevio As previo.clsprevio

Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim oMov As COMDMov.DCOMMov

Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lsFechaVencAnt As String '*** PEAC 20161213
Dim lsFecVnctoUnico As String
Dim lnMontoTransaccion As Currency
Dim lsCadImprimir As String
Dim lsNombreCliente As String
Dim lsBoletaCargo  As String 'CTI4 ERS0112020
Dim MatDatosAho(14) As String 'CTI4 ERS0112020
Dim lsNombreClienteCargoCta As String 'CTI4 ERS0112020

Dim lsFecVenImp  As String 'RECO20160412
'peac 20070818
'lsFechaVenc = Format$(fdVarFecVencimiento + fnVarNewPlazo, "mm/dd/yyyy")
lsFechaVenc = Format$(gdFecSis + fnVarNewPlazo, "mm/dd/yyyy")
lsFecVenImp = Format$(gdFecSis + fnVarNewPlazo, "dd/MM/yyyy") 'RECO20160412
lsFechaVencAnt = Format$(fdVarFecVencimiento, "yyyymmdd") '*** PEAC 20161213
lsFecVnctoUnico = Format$(fdVarFecVectoUnico, "yyyymmdd") '*** PEAC 20190723
'end peac

lnMontoTransaccion = CCur(Me.txtMontoPagar.Text)
lsNombreCliente = AXDesCon.listaClientes.ListItems(1).ListSubItems.Item(1)
'WIOR 20121009 Clientes Observados *************************************
Dim oDPersona As COMDPersona.DCOMPersona
Dim rsPersonaCred As ADODB.Recordset
Dim rsPersona As ADODB.Recordset
Dim Cont As Integer
Set oDPersona = New COMDPersona.DCOMPersona

If Not ValidaFormaPago Then Exit Sub 'CTI4 ERS0112020

Set rsPersonaCred = oDPersona.ObtenerPersCuentaRelac(Trim(AXCodCta.NroCuenta), gColRelPersTitular)

If rsPersonaCred.RecordCount > 0 Then
    If Not (rsPersonaCred.EOF And rsPersonaCred.BOF) Then
        For Cont = 0 To rsPersonaCred.RecordCount - 1
            Call VerSiClienteActualizoAutorizoSusDatos(Trim(rsPersonaCred!cperscod), fnVarOpeCod) 'FRHU ERS077-2015 20151204 OBSERVACION
            'WIOR 20130301 **************
            ReDim Preserve PersonaPago(Cont, 1)
            PersonaPago(Cont, 0) = Trim(rsPersonaCred!cperscod)
            PersonaPago(Cont, 1) = Trim(rsPersonaCred!cPersNombre)
            'WIOR **********************
            Set rsPersona = oDPersona.ObtenerUltimaVisita(Trim(rsPersonaCred!cperscod))
            If rsPersona.RecordCount > 0 Then
                If Not (rsPersona.EOF And rsPersona.BOF) Then
                    If Trim(rsPersona!sUsual) = "3" Then
                    MsgBox PstaNombre(Trim(rsPersonaCred!cPersNombre), True) & "." & Chr(10) & "CLIENTE OBSERVADO: " & Trim(rsPersona!cVisObserva), vbInformation, "Aviso"
                        Call frmPersona.Inicio(Trim(rsPersonaCred!cperscod), PersonaActualiza)
                    End If
                End If
            End If
            'Call VerSiClienteActualizoAutorizoSusDatos(Trim(rsPersonaCred!cPersCod), fnVarOpeCod) 'FRHU ERS077-2015 20151204
            Set rsPersona = Nothing
            rsPersonaCred.MoveNext
        Next Cont
    End If
End If
'WIOR FIN ***************************************************************

'*** AMDO20130705 TI-ERS063-2013
            Dim oDPersonaAct As COMDPersona.DCOMPersona
            Set oDPersonaAct = New COMDPersona.DCOMPersona
                            If oDPersonaAct.VerificaExisteSolicitudDatos(gColPigFunciones.vcodper) Then
                                MsgBox Trim("SE SOLICITA DATOS DEL CLIENTE: " & lsNombreCliente) & "." & Chr(10), vbInformation, "Aviso"
                                Call frmActInfContacto.Inicio(gColPigFunciones.vcodper)
                            End If
'***END AMDO

If MsgBox(" Grabar Renovación de Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdgrabar.Enabled = False
        'WIOR 20130301 ***SEGUN TI-ERS005-2013 ************************************************************
        If frmMovLavDinero.OrdPersLavDinero = "Exit" Or frmMovLavDinero.OrdPersLavDinero = "" Then
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
                    cmdgrabar.Enabled = True
                    Exit Sub
                End If
            Else
                fnCondicion = 0
                lnMonto = lnMontoTransaccion
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
                            cmdgrabar.Enabled = True
                            Exit Sub
                        End If
                    End If
                End If
                
            End If
        End If
        'WIOR FIN ***************************************************************
        fnVarNewSaldoCap = Format(CCur(AXDesCon.SaldoCapital) - vCapitalAdel - fnVarCapitalPagado, "0#.00")
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarRen = New COMNColoCPig.NCOMColPContrato
            'Grabar Renovacion Pignoraticio
'            Call loGrabarRen.nRenovacionCredPignoraticio(AXCodCta.NroCuenta, fnVarNewSaldoCap, lsFechaHoraGrab, _
'                 lsMovNro, lsFechaVenc, fnVarNewPlazo, lnMontoTransaccion, fnVarCapitalPagado, fnVarInteresVencido, _
'                 fnVarCostoCustodiaVencida, fnVarCostoPreparacionRemate, fnVarInteres, fnVarImpuesto, fnVarCostoCustodia, _
'                 fnVarDiasAtrasoReal, Val(txtNroRenovacion.Text), fnVarDiasCambCart, fnVarValorTasacion, fnVarOpeCod, _
'                 fsVarOpeDesc, fsVarPersCodCMAC, fnVarGastoCorrespondencia, gITF.gbITFAplica, gITF.gbITFAsumidocreditos, CCur(Val(Me.TxtITF.Text)), False, fnVarInteresMoratorio)

'*** PEAC 20161221 - SE CAMBIARON :
'fnVarCapitalPagado por CDbl(Me.txtCapital.Text), fnVarInteres por CDbl(Me.txtInteres.Text)

'CTI4 ERS0112020
    Select Case CInt(Trim(Right(CmbForPag.Text, 10)))
        Case gColocTipoPagoEfectivo
            fnVarOpeCod = gColPOpeRenovacEFE
        Case gColocTipoPagoVoucher
            fnVarOpeCod = gColPOpeRenovacVoucher
        Case gColocTipoPagoCargoCta
            fnVarOpeCod = gColPOpeRenovacCargoCta
    End Select
If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then lsNombreClienteCargoCta = PstaNombre(loGrabarRen.ObtieneNombreTitularCargoCta(txtCuentaCargo.NroCuenta))
'end CTI4

'JOEP20210914 Campana Prendario
If lblCampRetenPrend.Visible = True Then
    Dim oRegCampPrend As COMDColocPig.DCOMColPContrato
    Set oRegCampPrend = New COMDColocPig.DCOMColPContrato
    
    If nCampPrenRetencion = 1 Then
        Call oRegCampPrend.CampPrenRegRenovacionRetencion(AXCodCta.NroCuenta, txtCampRete.Text, lsMovNro, 1)
    End If
    
    If nCampPrenCampana = 1 Then
        Set oRegCampPrend = New COMDColocPig.DCOMColPContrato
        Call oRegCampPrend.CampPrenRegCampCred(AXCodCta.NroCuenta, txtCampRete.Text, "Renovacion", AXDesCon.TasaEfectivaMensual, txtNroRenovacion.Text, 0, 0, lsMovNro, 4, 2)
    End If
    
    Set oRegCampPrend = Nothing
End If
'JOEP20210914 Campana Prendario

Call loGrabarRen.nRenovacionCredPignoraticio(AXCodCta.NroCuenta, fnVarNewSaldoCap, lsFechaHoraGrab, _
lsMovNro, lsFechaVenc, fnVarNewPlazo, lnMontoTransaccion, CDbl(Me.txtCapital.Text), fnVarInteresVencido, _
fnVarCostoCustodiaVencida, fnVarCostoPreparacionRemate, CDbl(Me.Txtinteres.Text), fnVarImpuesto, fnVarCostoCustodia, _
fnVarDiasAtrasoReal, Val(txtNroRenovacion.Text), fnVarDiasCambCart, fnVarValorTasacion, fnVarOpeCod, _
fsVarOpeDesc, fsVarPersCodCMAC, fnVarGastoCorrespondencia, gITF.gbITFAplica, gITF.gbITFAsumidocreditos, _
CCur(Val(Me.TxtITF.Text)), False, fnVarInteresMoratorio, fnVarCostoNotificacion, gnMovNro, CDbl(Me.txtPgoAdelInt.Text), lsFecVnctoUnico, _
CInt(Trim(Right(CmbForPag.Text, 10))), nMovNroRVD, nMovNroRVDPend, txtCuentaCargo.NroCuenta, MatDatosAho)

'WIOR 20130301 agrego gnMovNro
'PEAC 20161213 se agregó lsFechaVencAnt
'PEAC 20190723 se reeemplazo lsFechaVencAnt por  lsFecVnctoUnico
'fnVarPlazo
            
        Set loGrabarRen = Nothing
        '*** BRGO 20110915 **************************
        If gITF.gbITFAplica Then
            Set oMov = New COMDMov.DCOMMov
            Call oMov.InsertaMovRedondeoITF(lsMovNro, 1, CCur(Val(Me.TxtITF.Text)) + nRedondeoITF, CCur(Val(Me.TxtITF.Text)))
            Set oMov = Nothing
        End If
        Set loContFunct = Nothing
        Set oMov = Nothing
        '********************************************
        '************
          'ADD JHCU 09-07-2020 REVERSIÓN PIGNORATICIO
         Set loGrabarRen = New COMNColoCPig.NCOMColPContrato
         Call loGrabarRen.ReversionReprogramacion(lsMovNro, AXCodCta.NroCuenta)
         Set loGrabarRen = Nothing
        'FIN JHCU 09-07-2020
        '************
        Set loImprime = New COMNColoCPig.NCOMColPImpre
'            lsCadImprimir = loImprime.nPrintReciboRenovacion(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
'                Format(AXDesCon.FechaPrestamo, "mm/dd/yyyy"), fnVarDiasAtraso, CCur(AXDesCon.SaldoCapital), fnVarCapitalPagado, _
'                fnVarInteres, fnVarInteresMoratorio, fnVarImpuesto, fnVarCostoCustodiaVencida + fnVarCostoCustodia, _
'                fnVarCostoPreparacionRemate, lnMontoTransaccion, fnVarNewSaldoCap, fnVarTasaInteres, _
'                Val(Me.txtNroRenovacion.Text), Format(lsFechaVenc, "mm/dd/yyyy"), gsCodUser, fnVarNewPlazo, fsVarNombreCMAC, fnVarGastoCorrespondencia, " ", CDbl(Val(TxtITF.Text)), gImpresora, fnVarInteresVencido)
'            lsCadImprimir = loImprime.nPrintReciboRenovacion(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
'                Format(AXDesCon.FechaPrestamo, "mm/dd/yyyy"), fnVarDiasAtraso, CCur(AXDesCon.SaldoCapital), fnVarCapitalPagado, _
'                fnVarInteres, fnVarInteresMoratorio, fnVarImpuesto, fnVarCostoCustodiaVencida + fnVarCostoCustodia, _
'                fnVarCostoPreparacionRemate, lnMontoTransaccion, fnVarNewSaldoCap, fnVarTasaInteres, _
'                Val(Me.txtNroRenovacion.Text), Format(lsFechaVenc, "mm/dd/yyyy"), gsCodUser, fnVarNewPlazo, fsVarNombreCMAC, fnVarCostoNotificacion, " ", CDbl(Val(TxtITF.Text)), gImpresora, fnVarInteresVencido, gbImpTMU) 'ARCV fnvarCostoNotificacion x fnVarGastoCorrespondencia

                '*** PEAC 20161221 - SE CAMBIARON :
                'fnVarCapitalPagado por CDbl(Me.txtCapital.Text), fnVarInteres por CDbl(Me.txtInteres.Text)
                'joep20210922 campana prendario
                If lblCampRetenPrend.Visible = True And pgTasaOriginal <> 0 And pgnCumpleCampna = 1 Then
                    fnVarTasaInteres = pgTasaOriginal
                End If
                'joep20210922 campana prendario
                
                lsCadImprimir = loImprime.nPrintReciboRenovacion(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
                Format(AXDesCon.FechaPrestamo, "dd/MM/yyyy"), fnVarDiasAtraso, CCur(AXDesCon.SaldoCapital), CDbl(Me.txtCapital.Text), _
                CDbl(Me.Txtinteres.Text), fnVarInteresMoratorio, fnVarImpuesto, fnVarCostoCustodiaVencida + fnVarCostoCustodia, _
                fnVarCostoPreparacionRemate, lnMontoTransaccion, fnVarNewSaldoCap, fnVarTasaInteres, _
                Val(Me.txtNroRenovacion.Text), Format(lsFechaVencImp, "dd/MM/yyyy"), gsCodUser, fnVarNewPlazo, fsVarNombreCMAC, fnVarCostoNotificacion, " ", CDbl(Val(TxtITF.Text)), gImpresora, fnVarInteresVencido, gbImpTMU, fnCredRevolAntNue, gnPigPorcenPgoCap, _
                fnDiasFerImp, txtCampRete.Text, pgnCumpleCampna) 'ARCV fnvarCostoNotificacion x fnVarGastoCorrespondencia
                '*** 20171109 - JHONY M. - Se agregó fnCredRevolAntNue, gnPigPorcenPgoCap
                '*** PEAC 20190711 - SE CAMBIO "lsFecVenImp" por "lsFechaVencImp" y se agregó "fnDiasFerImp"
            
        'CTI4 ERS0112020
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            lsBoletaCargo = loImprime.ImprimeBoletaAhorro("RETIRO AHORROS", "CARGO A CUENTA POR RENOV. PIGNO.", "", CStr(lnMontoTransaccion + Me.TxtITF.Text), lsNombreClienteCargoCta, txtCuentaCargo.NroCuenta, "", CDbl(MatDatosAho(10)), CDbl(MatDatosAho(3)), "", 1, CDbl(MatDatosAho(11)), , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, gsCodCMAC, 0, , , gbImpTMU)
        End If
        'END CTI4
                
            Set loImprime = Nothing
            
            Set loPrevio = New previo.clsprevio
            loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
            
            Do While True
                If MsgBox("Reimprimir Recibo de Renovación ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
                Else
                    Set loPrevio = Nothing
                    Exit Do
                End If
            Loop
            
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
        Set loPrevio = Nothing
        'WIOR 20130301 ************************************************************
        If regPersonaRealizaPago And gnMovNro > 0 Then
            frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(AXCodCta.NroCuenta), fnCondicion
            regPersonaRealizaPago = False
        End If
        'WIOR FIN *****************************************************************
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
                oSolicitud.ActualizarCapAutSinTarjetaVisto_nMovNro gsCodUser, gsCodAge, txtCuentaCargo.NroCuenta, rsCli!cperscod, nMovNroOperacion, CStr(gAhoCargoCtaRenovaPigno)
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

'Termina el formulario actual
Private Sub cmdSalir_Click()
    Unload Me
End Sub

'Procedimiento que calcula la deuda del cliente
Private Sub fgCalculaDeuda(Optional ByVal nDiasAtr As Integer = -1)
Dim loCalculos As COMNColoCPig.NCOMColPCalculos
Dim lsmensaje As String

'fnVarDiasAtraso = DateDiff("d", fdVarFecVencimiento, gdFecSis) RIRO 20200401 COMENTADO
fnVarDiasAtraso = IIf(nDiasAtr < 0, DateDiff("d", fdVarFecVencimiento, gdFecSis), nDiasAtr) 'RIRO 20200401

fnVarDiasAtrasoReal = fnVarDiasAtraso ' Dias Atraso Real

If fnVarDiasAtraso <= 0 Then
    fnVarDiasAtraso = 0
    fnVarInteresVencido = 0
    fnVarInteresMoratorio = 0
    fnVarCostoCustodia = 0
    fnVarImpuesto = 0
    Set loCalculos = New COMNColoCPig.NCOMColPCalculos

        'PEAC 20070813
        If gcCredAntiguo = "A" Then
            vInteresAdel = Round(0, 2)
        Else
            vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
            vDiasAdel = vDiasAdel - fnDiasFer 'RECO20141226 ERS170-2014
            '*** PEAC 20080806 ******************************
            'vInteresAdel = loCalculos.nCalculaInteresAdelantado(CCur(AXDesCon.SaldoCapital), fnVarTasaInteres, vDiasAdel)
             vInteresAdel = loCalculos.nCalculaInteresAlVencimiento(CCur(AXDesCon.SaldoCapital), fnVarTasaInteres, vDiasAdel)
            '*** FIN PEAC ***********************************
            'vInteresAdel = Round(vInteresAdel, 2)
            
            vInteresAdel = vInteresAdel + fnIntPendSaldo 'nUltIntAPagar
            'vInteresAdel = Round(IIf((vInteresAdel - fnPgoAdelInt) < 0, vInteresAdel, vInteresAdel - fnPgoAdelInt), 2) '*** PEAC 20161117
            vInteresAdel = Round(vInteresAdel, 2) '*** PEAC 20171109
            
        End If
        
        If fnCredRevolAntNue = True Then ' true =SI ES ANTIGUO
            vCapitalAdel = 0
        Else
            vCapitalAdel = Round(CCur(AXDesCon.SaldoCapital) * gnPigPorcenPgoCap, 2) '*** PEAC 20161020
        End If
        
        fnVarGastoCorrespondencia = loCalculos.nCalculaGastosCorrespondencia(AXCodCta.NroCuenta, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
    
    Set loCalculos = Nothing
Else

    Set loCalculos = New COMNColoCPig.NCOMColPCalculos

        'PEAC 20070813
        If gcCredAntiguo = "A" Then
            vInteresAdel = Round(0, 2)
        Else
        
        '*** PEAC 20170906 - MEJORA EL CALCULO DE LOS DIAS DE MORA
        If Format(vFecEstado, "yyyymmdd") >= Format(fdVarFecVencimiento, "yyyymmdd") Then
            vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
        Else
            'vDiasAdel = 30
            vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(fdVarFecVencimiento, "dd/mm/yyyy"))
        End If
        vDiasAdel = vDiasAdel - fnDiasFer '*** PEAC 20190521
        
'        If Format(vFecEstado, "dd/mm/yyyy") > Format(fdVarFecVencimiento, "dd/mm/yyyy") Then
'            vDiasAdel = 30
'        Else
'            vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(fdVarFecVencimiento, "dd/mm/yyyy"))
'            If vDiasAdel <= 0 Then
'                vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
'            End If
'        End If
        '*** FIN PEAC

         '*** PEAC 20080806 **********************************
          vInteresAdel = loCalculos.nCalculaInteresAlVencimiento(CCur(AXDesCon.SaldoCapital), fnVarTasaInteres, vDiasAdel)
         '*** FIN PEAC ***************************************
         
         '*** PEAC 20161221
         vInteresAdel = vInteresAdel + fnIntPendSaldo
         vInteresAdel = Round(vInteresAdel, 2) '*** PEAC 20171109
         
         'vInteresAdel = Round(IIf((vInteresAdel - fnPgoAdelInt) < 0, vInteresAdel, vInteresAdel - fnPgoAdelInt), 2) '*** PEAC 20161117
         '*** FIN PEAC
        End If
        
        If fnCredRevolAntNue = True Then ' true =SI ES ANTIGUO
            vCapitalAdel = 0
        Else
            vCapitalAdel = Round(CCur(AXDesCon.SaldoCapital) * gnPigPorcenPgoCap, 2) '*** PEAC 20161020
        End If
        
        'Agregar Calculo de Interes Compensatorio Vencido
        
        'fnVarInteresVencido = loCalculos.nCalculaInteresMoratorio(fnVarSaldoCap, fnVarTasaInteresVencido, fnVarDiasAtraso)
        'fnVarInteresMoratorio = loCalculos.nCalculaInteresMoratorio(fnVarSaldoCap, fnVarTasaInteresMoratorio, fnVarDiasAtraso)
        
        'If nPagoIntVenMor > 0 then
        If DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy")) = 0 Then
            fnVarInteresVencido = Round(0, 2)
            fnVarInteresMoratorio = Round(0, 2)
        Else
        
            If Format(vFecEstado, "yyyymmdd") >= Format(fdVarFecVencimiento, "yyyymmdd") Then
                fnVarDiasAtraso = vDiasAdel
            End If
        
            fnVarInteresVencido = loCalculos.nCalculaInteresMoratorio(CCur(AXDesCon.SaldoCapital), fnVarTasaInteresVencido, fnVarDiasAtraso, vInteresAdel)
            'fnVarInteresVencido = Round(fnVarInteresVencido, 2)
            'fnVarInteresVencido = Round(fnVarInteresVencido - fnIntComVenPgdo, 2)
            fnVarInteresVencido = Round(fnVarInteresVencido, 2)
        
            fnVarInteresMoratorio = loCalculos.nCalculaInteresMoratorio(CCur(AXDesCon.SaldoCapital), fnVarTasaInteresMoratorio, fnVarDiasAtraso)
            'fnVarInteresMoratorio = Round(fnVarInteresMoratorio, 2)
            'fnVarInteresMoratorio = Round(fnVarInteresMoratorio - fnIntMoraPgdo, 2)
            fnVarInteresMoratorio = Round(fnVarInteresMoratorio, 2)
            
        End If

        fnVarCostoCustodiaVencida = loCalculos.nCalculaCostoCustodiaMoratorio(fnVarValorTasacion, fnVarTasaCustodiaVencida, fnVarDiasAtraso)
        fnVarCostoCustodiaVencida = Round(fnVarCostoCustodiaVencida, 2)
        
        fnVarImpuesto = (fnVarInteresVencido + fnVarInteresMoratorio + fnVarCostoCustodiaVencida) * fnVarTasaImpuesto
        fnVarImpuesto = Round(fnVarImpuesto, 2)
        
        fnVarGastoCorrespondencia = loCalculos.nCalculaGastosCorrespondencia(AXCodCta.NroCuenta, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
    Set loCalculos = Nothing
End If
fnVarCostoPreparacionRemate = 0
'Modificado por DAOR 20070714, Si esta para remate y además estado en el ultimo remate fue no vendido=2

'*** PEAC 20080515
If fnVarEstado = gColPEstPRema And fnVarEstUltProcRem = 2 Then  ' Si esta en via de Remate
    fnVarCostoPreparacionRemate = fnVarTasaPreparacionRemate * CDbl(fnVarValorTasacion)
    fnVarCostoPreparacionRemate = Round(fnVarCostoPreparacionRemate, 2)
End If

'If fnVarEstado <> gColPEstPRema Then  ' Si no esta en via de Remate
'    fnVarCostoNotificacion = 0
'End If


'*** PEAC 20080515
'If gnNotifiAdju <> 1 And gnNotifiCob <> 0 Then
'    fnVarCostoNotificacion = 0
'End If

If gnNotifiAdju = 1 Then
    If gnNotifiCob = 1 Then
        fnVarCostoNotificacion = 0
    End If
Else
    fnVarCostoNotificacion = 0
End If



'Agrega Gastos de Correspondencia
'ARCV 14-03-2007
'fnVarDeuda = fnVarSaldoCap + fnVarInteresVencido + fnVarCostoCustodiaVencida _
        + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarGastoCorrespondencia + fnVarInteresMoratorio

'fnVarDeuda = fnVarSaldoCap + vInteresAdel + fnVarInteresVencido + fnVarCostoCustodiaVencida _
        + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarGastoCorrespondencia + fnVarInteresMoratorio + fnVarCostoNotificacion
'-------
End Sub

Private Sub fgCalculaMinimoPagar()
Dim loCalculos As COMNColoCPig.NCOMColPCalculos
' **************************************
' ** Calculo del Monto Minimo a Pagar **
' **************************************
'    Dim NumeroDias As Single
Set loCalculos = New COMNColoCPig.NCOMColPCalculos
    'fnVarInteres = 0
    fnVarFactor = loCalculos.nCalculaFactorRenovacion(fnVarTasaInteres, fnVarNewPlazo)
    'Ubicacion corte
    fnVarCostoCustodia = Round(loCalculos.nCalculaCostoCustodia(fnVarValorTasacion, fnVarTasaCustodia, fnVarNewPlazo), 2)
    fnVarCostoCustodia = Round(fnVarCostoCustodia, 2)
    
    'PEAC 20070813
'    If gcCredAntiguo = "A" Then
'        fnVarInteres = 0
'    Else
'        '*** PEAC 20080806 *****************************************
'        'fnVarInteres = loCalculos.nCalculaInteresAdelantado(CCur(AXDesCon.SaldoCapital), fnVarTasaInteres, vDiasAdel)
'         fnVarInteres = loCalculos.nCalculaInteresAlVencimiento(CCur(AXDesCon.SaldoCapital), fnVarTasaInteres, vDiasAdel)
'        '*** FIN PEAC **********************************************
'        fnVarInteres = Round(fnVarInteres, 2)
'    End If
    
    
    'fnVarInteres = Round(fnVarSaldoCap * fnVarFactor, 2)
    'fnVarInteres = Round(fnVarInteres, 2)
    
'*** PEAC 20170325 - esta variable ya está calculado
'    fnVarImpuesto = Round((fnVarInteresVencido + fnVarInteres + fnVarCostoCustodia + fnVarCostoCustodiaVencida + fnVarInteresMoratorio) * fnVarTasaImpuesto, 2)
'    fnVarImpuesto = Round(fnVarImpuesto, 2)
    
    'ARCV 14-03-2007
    'fnVarMontoMinimo = fnVarInteresVencido + fnVarCostoCustodiaVencida + fnVarInteres + fnVarCostoCustodia + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarInteresMoratorio
    'fnVarMontoMinimo = fnVarInteresVencido + fnVarCostoCustodiaVencida + fnVarInteres + fnVarCostoCustodia + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarInteresMoratorio + fnVarCostoNotificacion
    
    'fnVarMontoMinimo = fnVarInteresVencido + fnVarCostoCustodiaVencida + fnVarInteres + fnVarCostoCustodia + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarInteresMoratorio + fnVarCostoNotificacion + vCapitalAdel '*** PEAC 20161019
    fnVarMontoMinimo = fnVarInteresVencido + fnVarCostoCustodiaVencida + vInteresAdel + fnVarCostoCustodia + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarInteresMoratorio + fnVarCostoNotificacion + vCapitalAdel '*** PEAC 20161019
    'fnVarMontoMinimo = Round(fnVarMontoMinimo - fnPgoAdelInt, 2)
    fnVarMontoMinimo = Round(fnVarMontoMinimo, 2)
    
Set loCalculos = Nothing
End Sub
'Valida el campo cboplazonuevo
Private Sub cboPlazoNuevo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtMontoTotal.SetFocus
    'TXTMONTOPAGAR.SETFOCUS
End If
End Sub
Private Sub cboPlazoNuevo_Click()
    fnVarNewPlazo = Val(cboPlazoNuevo.Text)
    fgCalculaMinimoPagar
    txtMontoMinimoPagar.Text = Format(fnVarMontoMinimo + fnVarGastoCorrespondencia, "#0.00")
    'txtMontoPagar.Text = Format(fnVarMontoMinimo + fnVarGastoCorrespondencia, "#0.00")
    TxtMontoTotal.Text = Format((CDbl(fnVarMontoMinimo) + CDbl(fnVarGastoCorrespondencia)) / (1 - gITF.gnITFPorcent), "#0.00")
    fnVarCapitalPagado = 0
    fnVarNewSaldoCap = fnVarSaldoCap
    txtSaldoCapitalNuevo.Text = Format(fnVarSaldoCap - vCapitalAdel - fnVarCapitalPagado, "#0.00")
    
    '------ ITF ------
    If gITF.gbITFAplica Then
        If Not gITF.gbITFAsumidocreditos Then
'       Me.TxtITF = Format(gITF.fgITFCalculaImpuestoIncluido(Me.txtMontoPagar), "#0.00")
'       Me.TxtMontoTotal = CCur(Me.txtMontoPagar) + CCur(Me.TxtITF)
        'txtMontoPagar.Text = Format(gITF.fgITFCalculaImpuestoIncluido(Val(Me.TxtMontoTotal)), "#0.00")
        'Me.TxtITF = Format(CDbl(Me.TxtMontoTotal) - CDbl(Me.txtMontoPagar), "#0.00")
           If Val(TxtMontoTotal.Text) > Val(txtMontoMinimoPagar.Text) Then
             Me.TxtITF = gITF.fgITFCalculaImpuesto(TxtMontoTotal)
             txtMontoPagar = Format(CDbl(Me.TxtMontoTotal) - CDbl(Me.TxtITF), "#0.00")
           Else
             txtMontoPagar = Format(txtMontoMinimoPagar.Text, "#0.00")
             Me.TxtITF = gITF.fgITFCalculaImpuesto(txtMontoPagar)
             Me.TxtMontoTotal = Format(Val(txtMontoPagar.Text) + Val(TxtITF.Text))
           End If
           
        Else
            Me.TxtITF = Format(gITF.fgITFCalculaImpuesto(TxtMontoTotal), "#0.00")
            Me.TxtMontoTotal = Format(txtMontoPagar, "#0.00")
        End If
    Else
        Me.TxtITF = Format(0, "#0.00")
        txtMontoPagar = Format(Me.TxtMontoTotal, "#0.00")
    End If
    '-----------------
    cmdgrabar.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
        End If
    ElseIf KeyCode = 13 And AXCodCta.EnabledCta And AXCodCta.Age <> "" And Trim(AXCodCta.Cuenta) = "" Then
                AXCodCta.SetFocusCuenta
                 Exit Sub
    End If
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
TxtMontoTotal.Text = "0.00"

'JOEP20210914 Campana Prendario
lblCampRetenPrend.Visible = False
txtCampRete.Text = "0.00"
nCampPrenRetencion = 0
nCampPrenCampana = 0
pgnCumpleCampna = 0
'JOEP20210914 Campana Prendario

'ventana = 0
    Call CargaControles 'CTI4 ERS0112020
    
If fnVarOpeCod = 122700 Then
    fraFormaPago.Visible = False
Else
    fraFormaPago.Visible = True
End If
    
End Sub

Private Sub txtMontoPagar_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoPagar, KeyAscii)
Dim lnMontoPagar As Currency
Dim lnMontoPagarOriginal As Double
If KeyAscii = 13 Then
    If Trim(txtTotalDeuda) <> "" Then
        If CDbl(txtMontoPagar) >= CDbl(txtTotalDeuda) Then  'Monto Pagar >= Total Deuda
           MsgBox " Cliente debe RESCATAR su contrato ", , " Aviso "
           txtMontoPagar.SetFocus
           cmdgrabar.Enabled = False
           Exit Sub
        End If
    End If
    If CDbl(txtMontoPagar) < CDbl(txtMontoMinimoPagar) Then 'Monto Pagar < MinimoPagar
       MsgBox " Monto a Pagar debe ser Mayor que el Mínimo", , " Aviso "
       txtMontoPagar.SetFocus
       cmdgrabar.Enabled = False
       Exit Sub
    End If
    
    lnMontoPagarOriginal = CDbl(Me.txtMontoPagar.Text)
    
    'Distribuye los importes a las diferentes rubros
    fnVarCapitalPagado = CDbl(txtMontoPagar.Text) - CDbl(txtMontoMinimoPagar.Text) '+ CDbl(TxtITF.Text) 'peac 20070820
    'fnVarCapitalPagado = 0 'peac 20070820
    'lnMontoPagar = CDbl(txtMontoPagar.Text) - fnVarGastoCorrespondencia
    lnMontoPagar = CDbl(txtMontoPagar.Text) - fnVarInteres - fnVarGastoCorrespondencia 'peac 20070820
    vSumaCostoCustodia = fnVarCostoCustodia + fnVarCostoCustodiaVencida
    fnVarGastoCorrespondencia = 0
    If lnMontoPagar > CDbl(txtMontoMinimoPagar) - fnVarGastoCorrespondencia Then ' Monto Pagar = Minimo Pagar
'*****para diferencianorma
'        fnVarCapitalPagado = (lnMontoPagar - Round(fnVarFactor * fnVarSaldoCap, 2) - vSumaCostoCustodia - fnVarInteresVencido - _
'            fnVarCostoPreparacionRemate - Round(fnVarTasaImpuesto * fnVarFactor * fnVarSaldoCap, 2) - fnVarTasaImpuesto * vSumaCostoCustodia - _
'            fnVarTasaImpuesto * fnVarInteresVencido) / (1 - fnVarFactor - fnVarTasaImpuesto * fnVarFactor)
'
         ' AY *****
         
         'ARCV 14-03-2007
        'fnVarCapitalPagado = (Val(txtMontoPagar) - fnVarFactor * fnVarSaldoCap - fnVarInteresVencido - fnVarInteresMoratorio - _
                                fnVarCostoPreparacionRemate - fnVarTasaImpuesto * fnVarFactor * fnVarSaldoCap - _
                                (fnVarTasaImpuesto * (fnVarInteresVencido + fnVarInteresMoratorio))) / (1 - fnVarFactor - fnVarTasaImpuesto * fnVarFactor)
            'peac 20070820
         'fnVarCapitalPagado = (Val(txtMontoPagar) - fnVarFactor * fnVarSaldoCap - fnVarInteresVencido - fnVarInteresMoratorio - _
                                fnVarCostoPreparacionRemate - fnVarCostoNotificacion - fnVarTasaImpuesto * fnVarFactor * fnVarSaldoCap - _
                                (fnVarTasaImpuesto * (fnVarInteresVencido + fnVarInteresMoratorio))) / (1 - fnVarFactor - fnVarTasaImpuesto * fnVarFactor)
        '---------
    fnVarCapitalPagado = Round(fnVarCapitalPagado, 2)
    Dim loCalculos As COMNColoCPig.NCOMColPCalculos
    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
    'PEAC 20070813
    fnVarDiasAtraso = DateDiff("d", fdVarFecVencimiento, gdFecSis)
    fnVarDiasAtrasoReal = fnVarDiasAtraso ' Dias Atraso Real
    If fnVarDiasAtraso <= 0 Then
       fnVarDiasAtraso = 0
        vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
    Else
        vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(fdVarFecVencimiento, "dd/mm/yyyy"))
    End If
    If gcCredAntiguo = "A" Then
        fnVarInteres = 0
    Else
        '*** PEAC 20080806 ************************************
        'fnVarInteres = loCalculos.nCalculaInteresAdelantado(CCur(AXDesCon.SaldoCapital), fnVarTasaInteres, vDiasAdel)
         fnVarInteres = loCalculos.nCalculaInteresAlVencimiento(CCur(AXDesCon.SaldoCapital), fnVarTasaInteres, vDiasAdel - fnDiasFer)
        '*** FIN PEAC *****************************************
        
        'fnVarInteres = Round(fnVarInteres, 2)
        fnVarInteres = Round(fnVarInteres - fnPgoAdelInt, 2) '*** PEAC 20161117
        
    End If

    Set loCalculos = Nothing
'        fnVarInteres = fnVarFactor * (fnVarSaldoCap - fnVarCapitalPagado)
'        fnVarInteres = Round(fnVarInteres, 2)
        
        fnVarImpuesto = fnVarTasaImpuesto * (fnVarInteres + fnVarInteresVencido + fnVarInteresMoratorio + vSumaCostoCustodia)
        fnVarImpuesto = Round(fnVarImpuesto, 2)
'       'PARA REDONDEO
        'ARCV 14-03-2007
        'fnVarInteres = lnMontoPagar - fnVarCapitalPagado - fnVarInteresVencido - vSumaCostoCustodia - fnVarImpuesto - fnVarCostoPreparacionRemate - fnVarInteresMoratorio
        
        
        'comentado 14/08/07 peac
'        fnVarInteres = lnMontoPagar - fnVarCapitalPagado - fnVarInteresVencido - vSumaCostoCustodia - fnVarImpuesto - fnVarCostoPreparacionRemate - fnVarInteresMoratorio - fnVarCostoNotificacion
'        fnVarInteres = Round(fnVarInteres, 2)
    End If
    txtMontoPagar.Text = Format(txtMontoPagar.Text, "#0.00")
    'fnVarNewSaldoCap = Format(fnVarSaldoCap - fnVarCapitalPagado, "#0.00")
    'fnVarNewSaldoCap = Format(fnVarSaldoCap - vCapitalAdel - fnVarCapitalPagado, "#0.00") '*** PEAC 20160921
    'fnVarNewSaldoCap = Format(fnVarSaldoCap - fnVarCapitalPagado, "#0.00") '*** PEAC 20161019
    fnVarNewSaldoCap = Round(fnVarSaldoCap - fnVarCapitalPagado, 2) '*** PEAC 20170928
    txtSaldoCapitalNuevo.Text = Format(fnVarNewSaldoCap, "#0.00")
    
    '*** PEAC 20080528
    fnVarDeuda = CCur(AXDesCon.SaldoCapital) + vInteresAdel + fnVarInteresVencido + fnVarCostoCustodiaVencida _
            + fnVarImpuesto + fnVarGastoCorrespondencia + fnVarInteresMoratorio + fnVarCostoNotificacion + TxtITF.Text
    '*** PEAC 20160921
    'fnVarDeuda = CCur(AXDesCon.SaldoCapital) + vCapitalAdel + vInteresAdel + fnVarInteresVencido + fnVarCostoCustodiaVencida _
     '       + fnVarImpuesto + fnVarGastoCorrespondencia + fnVarInteresMoratorio + fnVarCostoNotificacion + TxtITF.Text
    
    txtTotalDeuda.Text = Format(fnVarDeuda, "#0.00")
    
    'txtCapital.Text = Format(TxtMontoTotal.Text - fnVarCostoNotificacion - vInteresAdel - fnVarInteresMoratorio - fnVarInteresVencido - TxtITF.Text, "#0.00")
    'txtCapital.Text = Format(TxtMontoTotal.Text - fnVarCostoNotificacion - vCapitalAdel - vInteresAdel - fnVarInteresMoratorio - fnVarInteresVencido - TxtITF.Text, "#0.00") '*** PEAC 20160921
    txtCapital.Text = Format(TxtMontoTotal.Text - fnVarCostoNotificacion - vInteresAdel - fnVarInteresMoratorio - fnVarInteresVencido - TxtITF.Text, "#0.00") '*** PEAC 20161019
    
    TxtMora.Text = Format(fnVarInteresMoratorio, "#0.00")
    txtIntVen.Text = Format(fnVarInteresVencido, "#0.00")
    Txtinteres.Text = Format(vInteresAdel, "#0.00")
    txtCostoCus.Text = Format(fnVarCostoCustodiaVencida, "#0.00")
    txtCostoNoti.Text = Format(fnVarCostoNotificacion, "#0.00")
        
    txtSaldoCapitalNuevo.Text = Format(CCur(AXDesCon.SaldoCapital) - vCapitalAdel - fnVarCapitalPagado, "0#.00")
    '*** FIN PEAC 20080528
    
    If lnMontoPagarOriginal < CDbl(Me.txtMontoPagar.Text) Then
        MsgBox "No se puede pagar menos del mínimo a pagar. Tiene que realizar una amortización.", vbExclamation + vbOKOnly, "Atención"
    End If
    
    cmdgrabar.Enabled = True
    cmdgrabar.SetFocus
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
    fnVarTopRenovaciones = loParam.dObtieneColocParametro(gConsColPMaxNroRenovac)
    fnVarTopRenovacionesNuevo = loParam.dObtieneColocParametro(3057)
    gnPigPorcenPgoCap = loParam.dObtieneColocParametro(9050) '*** PEAC 20160920
    gnPigVigMeses = loParam.dObtieneColocParametro(9051) '*** PEAC 20160920
    
   'madm 20091204 ******************************
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

   ' fnVarCostoNotificacion = loParam.dObtieneColocParametro(3103) 'ARCV 14-03-2007
   '********** end madm **************************
Set loParam = Nothing
End Sub

Private Sub TxtMontoTotal_GotFocus()
    fEnfoque TxtMontoTotal
End Sub

Private Sub TxtMontoTotal_KeyPress(KeyAscii As Integer)
Dim nMontoPagar As Double
Dim loValida As COMNColoCPig.NCOMColPValida 'CTI4ERS0112020
Dim bEsMismoTitular As Boolean 'CTI4 ERS0112020
    Set loValida = New COMNColoCPig.NCOMColPValida
    
bEsMismoTitular = loValida.EsMismoTitulardeCuentaPignoYAhorro(txtCuentaCargo.NroCuenta, AXCodCta.NroCuenta)
KeyAscii = NumerosDecimales(TxtMontoTotal, KeyAscii)
If KeyAscii = 13 Then
    
    cmdgrabar.Enabled = True
    
    If gITF.gbITFAplica And Not bEsMismoTitular Then
        If Not gITF.gbITFAsumidocreditos Then
            'Me.TxtITF = Format(gITF.fgITFCalculaImpuestoIncluido(TxtMontoTotal), "#0.00")
            'txtMontoPagar = Format(CDbl(Me.TxtMontoTotal) - CDbl(Me.TxtITF), "#0.00")
            'txtMontoPagar.Text = Format(gITF.fgITFCalculaImpuestoIncluido(TxtMontoTotal), "#0.00")
           ' Me.TxtITF = Format(CDbl(Me.TxtMontoTotal) - CDbl(Me.txtMontoPagar), "#0.00")
           If Val(TxtMontoTotal.Text) > Val(txtMontoMinimoPagar.Text) Then
              Me.TxtITF = gITF.fgITFCalculaImpuesto(TxtMontoTotal)
              '*** BRGO 20110908 ************************************************
              nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtITF.Text))
              If nRedondeoITF > 0 Then
                 Me.TxtITF.Text = Format(CCur(Me.TxtITF.Text) - nRedondeoITF, "#,##0.00")
              End If
              '*** END BRGO
              txtMontoPagar = Format(CDbl(Me.TxtMontoTotal) - CDbl(Me.TxtITF), "#0.00")
           Else
             txtMontoPagar = Format(txtMontoMinimoPagar.Text, "#0.00")
             Me.TxtITF = gITF.fgITFCalculaImpuesto(txtMontoPagar)
              '*** BRGO 20110908 ************************************************
              nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtITF.Text))
              If nRedondeoITF > 0 Then
                 Me.TxtITF.Text = Format(CCur(Me.TxtITF.Text) - nRedondeoITF, "#,##0.00")
              End If
              '*** END BRGO
             Me.TxtMontoTotal = Format(Val(txtMontoPagar.Text) + Val(TxtITF.Text))
           End If
           
        Else
            Me.TxtITF = Format(gITF.fgITFCalculaImpuesto(TxtMontoTotal), "#0.00")
            '*** BRGO 20110908 ************************************************
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtITF.Text))
            If nRedondeoITF > 0 Then
               Me.TxtITF.Text = Format(CCur(Me.TxtITF.Text) - nRedondeoITF, "#,##0.00")
            End If
            '*** END BRGO
            txtMontoPagar = Format(CDbl(Me.TxtMontoTotal) - CDbl(Me.TxtITF), "#0.00")
        End If
    Else
        Me.TxtITF = Format(0, "#0.00")
        txtMontoPagar = Format(Me.TxtMontoTotal, "#0.00")
    End If
    
    If Trim(txtTotalDeuda) <> "" Then
    
        'peac 20071219
        'If CDbl(txtMontoPagar) >= CDbl(txtTotalDeuda) Then  'Monto Pagar >= Total Deuda
        If CDbl(TxtMontoTotal) >= CCur(AXDesCon.SaldoCapital) Then  'Monto Pagar >= saldo capital
        
           MsgBox " Cliente debe RESCATAR su contrato ", , " Aviso "
           TxtMontoTotal.SetFocus
           cmdgrabar.Enabled = False
           Exit Sub
        End If
    Else
        cmdgrabar.Enabled = False
        Exit Sub
    End If
    
    If CDbl(txtMontoPagar) < CDbl(txtMontoMinimoPagar) Then 'Monto Pagar < MinimoPagar
       MsgBox " Monto a Pagar debe ser Mayor que el Mìnimo", , " Aviso "
     '  txtMontoPagar.SetFocus
       cmdgrabar.Enabled = False
       Exit Sub
    End If
    
    Call txtMontoPagar_KeyPress(13)
  
  If cmdgrabar.Enabled Then
    cmdgrabar.SetFocus
  End If
End If
End Sub


Private Sub TxtMontoTotal_LostFocus()
    cmdgrabar.Enabled = False
    fEnfoque TxtMontoTotal
    cmdgrabar.Enabled = True
    
    If Trim(TxtMontoTotal.Text) = "" Then
        TxtMontoTotal.Text = "0.00"
    End If
    TxtMontoTotal.Text = Format(TxtMontoTotal.Text, "#0.00")
    
    Call TxtMontoTotal_KeyPress(13)
    
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
                        
            lnTipMot = 18 ' Renovacion Credito Pignoraticio
            oformVou.iniciarFormularioDeposito CInt(Mid(AXCodCta.NroCuenta, 9, 1)), lnTipMot, sGlosa, sIF, sVaucher, nMontoVoucher, sPersCod, nMovNroRVD, nMovNroRVDPend, sNombre, sDireccion, sDocumento, AXCodCta.NroCuenta
            If Len(sVaucher) = 0 Then Exit Sub
            LblNumDoc.Caption = Trim(Mid(sVaucher, 1, Len(sVaucher) - 10))
            cmdgrabar.Enabled = True
        ElseIf CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            Dim sCuenta As String
            
            sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(gAhoCargoCtaRenovaPigno), sNumTarj, 232, False)
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
            AsignaValorITF
            cmdgrabar.Enabled = True
            cmdgrabar.SetFocus
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
            cmdgrabar.Enabled = True
        Case gColocTipoPagoEfectivo
            txtCuentaCargo.Visible = False
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            lblNroDocumento.Visible = False
            cmdgrabar.Enabled = True
        Case gColocTipoPagoCargoCta
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargo.Visible = True
            txtCuentaCargo.Enabled = True
            txtCuentaCargo.CMAC = gsCodCMAC
            txtCuentaCargo.Prod = Trim(Str(gCapAhorros))
            cmdgrabar.Enabled = False
        Case gColocTipoPagoVoucher
            LblNumDoc.Visible = True
            lblNroDocumento.Visible = True
            txtCuentaCargo.Visible = False
            cmdgrabar.Enabled = False
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
                        nRespuesta = oSolicitud.SolicitarVistoAtencionSinTarjeta(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gAhoCargoCtaRenovaPigno))
                    
                        If nRespuesta = 1 Then '1:Tiene Visto de atencion sin tarjeta pendiente de autorizar
                             MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                             Exit Sub
                        End If
                        If nRespuesta = 2 Then '2:Tiene visto aceptado
                            MsgBox "La solicitud de atención sin tarjeta fue Aprobada, proceda con la atención", vbInformation, "Aviso"
                        End If
                        If nRespuesta = 3 Then '3:Tiene visto rechazado
                           If MsgBox("La solicitud de atención sin tarjeta fue RECHAZADA. ¿Desea realizar una nueva solicitud?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                Call oSolicitud.RegistrarVistoDeUsuario(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gAhoCargoCtaRenovaPigno))
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
                        
                                bExitoSol = frmCapMotivoAutorizacion.Inicio(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gAhoCargoCtaRenovaPigno))
                                If bExitoSol > 0 Then
                                    MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                                End If
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        lbVistoVal = loVistoElectronico.Inicio(5, gAhoCargoCtaRenovaPigno)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema"
                            Exit Sub
                        End If
                    End If
                ElseIf lsTieneTarj = "NO" And rsV.RecordCount > 0 Then
                    If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? ", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'add by marg ers 065-2017
                        lbVistoVal = loVistoElectronico.Inicio(5, gAhoCargoCtaRenovaPigno)
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
        AsignaValorITF
        cmdgrabar.Enabled = True
        cmdgrabar.SetFocus
    End If
End Sub
Private Sub AsignaValorITF()
Dim loValida As COMNColoCPig.NCOMColPValida
Dim bEsMismoTitular As Boolean
    Set loValida = New COMNColoCPig.NCOMColPValida
    
    bEsMismoTitular = loValida.EsMismoTitulardeCuentaPignoYAhorro(txtCuentaCargo.NroCuenta, AXCodCta.NroCuenta)
    
    If gITF.gbITFAplica And Not bEsMismoTitular Then
        If Not gITF.gbITFAsumidocreditos Then
            txtMontoPagar = Format(txtMontoMinimoPagar.Text, "#0.00")
            Me.TxtITF = gITF.fgITFCalculaImpuesto(txtMontoPagar)
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtITF.Text))
            If nRedondeoITF > 0 Then
               Me.TxtITF.Text = Format(CCur(Me.TxtITF.Text) - nRedondeoITF, "#,##0.00")
            End If
            Me.TxtMontoTotal = Format(Val(txtMontoPagar.Text) + Val(TxtITF.Text))
        Else
            Me.TxtITF = gITF.fgITFCalculaImpuesto(TxtMontoTotal)
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtITF.Text))
            If nRedondeoITF > 0 Then
               Me.TxtITF.Text = Format(CCur(Me.TxtITF.Text) - nRedondeoITF, "#,##0.00")
            End If
            txtMontoPagar = Format(CDbl(Me.TxtMontoTotal), "#0.00")
        End If
    
    Else
        Me.TxtITF = Format(0, "#0.00")
        txtMontoPagar = Format(Me.TxtMontoTotal, "#0.00")
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
Private Sub CampPrendVerificaRetencion(ByVal pcCtaCod As String, ByVal pdFechaSist As String)
    Dim objCampRete As COMDColocPig.DCOMColPContrato
    Dim rsCampRete As ADODB.Recordset
    Set objCampRete = New COMDColocPig.DCOMColPContrato
    
    
    Set rsCampRete = objCampRete.CampPrendarioDesVerificaCampReten(pcCtaCod, pdFechaSist)
    
    If Not (rsCampRete.BOF And rsCampRete.EOF) Then
        lblCampRetenPrend.Visible = True
        txtCampRete.Text = rsCampRete!TENNew
        lblCampRetenPrend.Caption = rsCampRete!cResultado
        nCampPrenRetencion = 1
        nCampPrenCampana = 0
    Else
        lblCampRetenPrend.Visible = False
        lblCampRetenPrend.Caption = ""
        txtCampRete.Text = 0#
        nCampPrenRetencion = 0
        nCampPrenCampana = 0
    End If
    
    Set objCampRete = Nothing
    RSClose rsCampRete
End Sub

Private Sub CampPrendVerificaCampanas(ByVal pcCtaCod As String, ByVal pdFechaSis As String, ByVal pnModulo As Integer, Optional ByVal pnRenva As Integer)
    Dim oCampPrend As COMDColocPig.DCOMColPContrato
    Dim rsCampPrend As ADODB.Recordset
    Set oCampPrend = New COMDColocPig.DCOMColPContrato
    
    Set rsCampPrend = oCampPrend.CampPrendarioDesbCampa(pcCtaCod, pdFechaSis, pnModulo, pnRenva)
    If Not (rsCampPrend.BOF And rsCampPrend.EOF) Then
        lblCampRetenPrend.Caption = rsCampPrend!cResultado
        lblCampRetenPrend.Visible = True
        txtCampRete.Text = rsCampPrend!nCampana
        nCampPrenRetencion = 0
        nCampPrenCampana = 1
        pgnCumpleCampna = rsCampPrend!Cumple
        pgTasaOriginal = rsCampPrend!TEMOriginal
    Else
        lblCampRetenPrend.Visible = False
        lblCampRetenPrend.Caption = ""
        txtCampRete.Text = 0#
        nCampPrenRetencion = 0
        nCampPrenCampana = 0
        pgnCumpleCampna = 0
        pgTasaOriginal = 0
    End If
    Set oCampPrend = Nothing
    RSClose oCampPrend
End Sub
'JOEP20210916 campana prendario

