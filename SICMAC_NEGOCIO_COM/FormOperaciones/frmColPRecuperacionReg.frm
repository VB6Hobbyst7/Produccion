VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmColPRecuperacionReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio : Recuperación Contrato Adjudicado"
   ClientHeight    =   8805
   ClientLeft      =   915
   ClientTop       =   2100
   ClientWidth     =   8190
   Icon            =   "frmColPRecuperacionReg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraContenedor 
      Height          =   7770
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   7935
      Begin VB.Frame fraContenedor 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1005
         Index           =   1
         Left            =   75
         TabIndex        =   11
         Top             =   4320
         Width           =   7665
         Begin VB.TextBox txtNomAdj 
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
            Left            =   825
            Locked          =   -1  'True
            TabIndex        =   15
            Tag             =   "txtnombre"
            Top             =   540
            Width           =   3780
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "&Buscar..."
            Enabled         =   0   'False
            Height          =   300
            Left            =   2520
            TabIndex        =   0
            Top             =   195
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.TextBox txtCodAdj 
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
            Height          =   330
            Left            =   825
            Locked          =   -1  'True
            TabIndex        =   14
            Tag             =   "txtcodigo"
            Top             =   195
            Width           =   1455
         End
         Begin VB.TextBox txtTriAdj 
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
            Left            =   6420
            TabIndex        =   13
            Tag             =   "txtTributario"
            Top             =   540
            Width           =   1080
         End
         Begin VB.TextBox txtNatAdj 
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
            Left            =   6435
            TabIndex        =   12
            Tag             =   "txtDocumento"
            Top             =   210
            Width           =   1080
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nombre :"
            Height          =   225
            Index           =   7
            Left            =   135
            TabIndex        =   19
            Top             =   585
            Width           =   735
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Código :"
            Height          =   225
            Index           =   8
            Left            =   120
            TabIndex        =   18
            Top             =   255
            Width           =   615
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Doc.Natural : "
            Height          =   255
            Index           =   2
            Left            =   5265
            TabIndex        =   17
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Doc.Tributario : "
            Height          =   255
            Index           =   3
            Left            =   5280
            TabIndex        =   16
            Top             =   585
            Width           =   1110
         End
      End
      Begin VB.Frame txtPrecioVta 
         Height          =   2430
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   5280
         Width           =   7695
         Begin VB.TextBox txtDsctoInt 
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
            Left            =   3525
            TabIndex        =   52
            Text            =   "0"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtCustodiaDif 
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
            Left            =   3525
            TabIndex        =   50
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtPreVntaFinal 
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
            Left            =   5400
            TabIndex        =   47
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox txtDevol 
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
            Left            =   4080
            TabIndex        =   45
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox txtPreVnta 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   2760
            TabIndex        =   43
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox txtDiasTrans 
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
            TabIndex        =   41
            Top             =   2040
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
            Left            =   3525
            TabIndex        =   33
            Top             =   600
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
            Left            =   885
            TabIndex        =   32
            Top             =   960
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
            Left            =   885
            TabIndex        =   31
            Top             =   240
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
            Left            =   3525
            TabIndex        =   30
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtIntComp 
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
            Left            =   885
            TabIndex        =   29
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox TxtMontoTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   27
            Text            =   "0"
            Top             =   885
            Width           =   1260
         End
         Begin VB.TextBox TxtItf 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   25
            Text            =   "0"
            Top             =   540
            Width           =   1260
         End
         Begin VB.TextBox TxtInteres 
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
            Left            =   9720
            TabIndex        =   23
            Text            =   "0"
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtPreVentaBruta 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   1
            Text            =   "0"
            Top             =   195
            Width           =   1260
         End
         Begin VB.TextBox txtDeuda 
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
            Height          =   315
            Left            =   840
            TabIndex        =   8
            Text            =   "0"
            Top             =   1340
            Width           =   1215
         End
         Begin MSMask.MaskEdBox txtFecAdju 
            Height          =   300
            Left            =   120
            TabIndex        =   39
            Top             =   2040
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblDscto 
            Caption         =   "Dscto. Int.:"
            Height          =   255
            Left            =   2160
            TabIndex        =   53
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblCostoNoti 
            AutoSize        =   -1  'True
            Caption         =   "Custodia Diferida:"
            Height          =   195
            Index           =   0
            Left            =   2160
            TabIndex        =   51
            Top             =   1080
            Width           =   1245
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "(Deuda + Importe Adicional)"
            Height          =   195
            Index           =   15
            Left            =   5040
            TabIndex        =   49
            Top             =   1200
            Width           =   1965
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Importe Adicional:"
            Height          =   195
            Index           =   12
            Left            =   5400
            TabIndex        =   48
            Top             =   1800
            Width           =   1305
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Devolución:"
            Height          =   195
            Index           =   11
            Left            =   4080
            TabIndex        =   46
            Top             =   1800
            Width           =   930
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Interés Adicional:"
            Height          =   195
            Index           =   9
            Left            =   2760
            TabIndex        =   44
            Top             =   1800
            Width           =   1290
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Dias Transc.:"
            Height          =   270
            Index           =   6
            Left            =   1440
            TabIndex        =   42
            Top             =   1800
            Width           =   1050
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Adjud.:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   40
            Top             =   1800
            Width           =   1050
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblCostoNoti 
            AutoSize        =   -1  'True
            Caption         =   "Costo Notificación:"
            Height          =   195
            Index           =   5
            Left            =   2160
            TabIndex        =   38
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblIntVen 
            AutoSize        =   -1  'True
            Caption         =   "Int. Vcdo."
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   690
         End
         Begin VB.Label lblCapital 
            AutoSize        =   -1  'True
            Caption         =   "Capital"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   480
         End
         Begin VB.Label lblMora 
            AutoSize        =   -1  'True
            Caption         =   "Mora:"
            Height          =   195
            Index           =   1
            Left            =   3000
            TabIndex        =   35
            Top             =   360
            Width           =   405
         End
         Begin VB.Label lblInteres 
            AutoSize        =   -1  'True
            Caption         =   "Interés"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   480
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Total Venta Final:"
            Height          =   195
            Index           =   4
            Left            =   5010
            TabIndex        =   28
            Top             =   975
            Width           =   1245
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "ITF Efectivo:"
            Height          =   195
            Index           =   1
            Left            =   5010
            TabIndex        =   26
            Top             =   630
            Width           =   915
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Interes Moratorio:"
            Height          =   195
            Index           =   0
            Left            =   9720
            TabIndex        =   24
            Top             =   360
            Visible         =   0   'False
            Width           =   1290
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Deuda :"
            Height          =   195
            Index           =   13
            Left            =   150
            TabIndex        =   10
            Top             =   1410
            Width           =   570
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Monto Recup.:"
            Height          =   195
            Index           =   14
            Left            =   5010
            TabIndex        =   9
            Top             =   285
            Width           =   1065
         End
      End
      Begin MSMask.MaskEdBox txtNroDocumento 
         Height          =   330
         Left            =   5445
         TabIndex        =   2
         Top             =   5790
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###-#####"
         PromptChar      =   "_"
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   45
         TabIndex        =   21
         Top             =   255
         Width           =   3615
         _extentx        =   6376
         _extenty        =   661
         texto           =   "Crédito"
         enabledcta      =   -1
         enabledprod     =   -1
         enabledage      =   -1
      End
      Begin SICMACT.ActXColPDesCon AXDesCon 
         Height          =   3855
         Left            =   90
         TabIndex        =   22
         Top             =   630
         Width           =   7635
         _extentx        =   13467
         _extenty        =   6800
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Nro.Documento :"
         Height          =   225
         Index           =   10
         Left            =   4095
         TabIndex        =   20
         Top             =   5835
         Visible         =   0   'False
         Width           =   1305
      End
   End
   Begin VB.Frame fraFormaPago 
      Height          =   600
      Left            =   135
      TabIndex        =   58
      Top             =   7750
      Width           =   7935
      Begin VB.ComboBox CmbForPag 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   200
         Width           =   1785
      End
      Begin SICMACT.ActXCodCta txtCuentaCargo 
         Height          =   375
         Left            =   3120
         TabIndex        =   54
         Top             =   200
         Visible         =   0   'False
         Width           =   3630
         _extentx        =   6403
         _extenty        =   661
         texto           =   "Cuenta N°:"
         enabledcta      =   -1
         enabledage      =   -1
      End
      Begin VB.Label lblNroDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Left            =   3105
         TabIndex        =   55
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
         TabIndex        =   56
         Top             =   200
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblFormaPago 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
         Height          =   195
         Left            =   180
         TabIndex        =   57
         Top             =   250
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   315
      Left            =   7050
      TabIndex        =   5
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5100
      TabIndex        =   3
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   6075
      TabIndex        =   4
      Top             =   8400
      Width           =   975
   End
End
Attribute VB_Name = "frmColPRecuperacionReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* REGISTRO DE VENTA DE SUBASTA.
'Archivo:  frmColPSubastaRegVenta.frm
'LAYG   :  18/07/2001.
'Resumen:  Nos permite registrar una venta de un contrato que se está Subastando
Option Explicit

Dim fnVarNroSubasta As String
Dim fsVarNroSubasta As String
Dim pDifeDiasRema As Integer
Dim fnVarTasaImpuesto As Double
Dim fnVarTasaPreparacionRemate As Double
Dim fnVarTasaInteresVencido As Double
Dim fnVarTasaIGV As Double
Dim pAgeRemSub As String * 2
Dim fnVarDiasAtraso As Integer
Dim fnVarInteresVencido As Double

Dim fnVtaJoyaSINSubasta As Integer

Dim fnVarPrecioAdjudica As Currency
Dim nRedondeoITF As Double

Dim lnPreVta As Double '*** PEAC 20180320
Dim lnDevol As Double '*** PEAC 20180320
Dim lnDeudatotal As Currency '*** PEAC 20180712
Dim lnPreVtaInicial As Double '*** PEAC 20180813
Dim lnDsctoIntTotal As Double '*** PEAC 20200729
Dim lbCampanaDscoInt As Integer '*** PEAC 20190726
Dim lbDscoInt As Integer '*** PEAC 20190726
Dim lnPorcenDsctoInt As Integer '*** PEAC 20200811
Dim lnITF As Double 'PEAC 20200903
Private nMontoVoucher As Currency 'CTI4 ERS0112020
Dim nMovNroRVD As Long, nMovNroRVDPend As Long 'CTI4 ERS0112020
Dim sNumTarj As String 'CTI4 ERS0112020
Dim loVistoElectronico As frmVistoElectronico 'CTI4 ERS0112020
Dim nRespuesta As Integer 'CTI4 ERS0112020
'JAOR20210225
Dim nTipoCamp As Integer 'Tipo de acción campaña:  1=Matener calif. 2=Afectar calif.
'Inicializa el formulario
Public Sub Inicio(ByVal psNroProceso As String)
    
'/*Verificar cantidad de operaciones disponibles ANDE 20171218*/
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim bProsigue As Boolean
    Dim cMsgValid As String
    bProsigue = oCaptaLN.OperacionPermitida(gsCodUser, gdFecSis, gsOpeCod, cMsgValid)
    If bProsigue = False Then
        MsgBox cMsgValid, vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
'/*end ande*/

    fsVarNroSubasta = psNroProceso
    CargaParametros
    If fnVtaJoyaSINSubasta = 1 Then ' Venta sin subasta
    
    End If
    
    Limpiar
    Me.Show 1
End Sub

'Inicializa las variables
Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.Limpiar
    txtCodAdj = ""
    txtNomAdj = ""
    txtNatAdj = ""
    txtTriAdj = ""
    txtDeuda.Text = Format(0, "#0.00")
    'txtPreBaseVenta.Text = Format(0, "#0.00")
    Me.TxtInteres.Text = Format(0, "#0.00")
    txtPreVentaBruta.Text = Format(0, "#0.00")
    
    '*** PEAC 20180712
    Me.txtCapital.Text = Format(0, "#0.00")
    Me.txtIntComp.Text = Format(0, "#0.00")
    Me.txtIntVen.Text = Format(0, "#0.00")
    Me.txtDiasTrans.Text = Format(0, "#0.00")
    Me.txtPreVnta.Text = Format(0, "#0.00")
    Me.txtDevol.Text = Format(0, "#0.00")
    Me.txtPreVntaFinal.Text = Format(0, "#0.00")
    Me.txtMora.Text = Format(0, "#0.00")
    Me.txtCostoNoti.Text = Format(0, "#0.00")
    Me.txtCustodiaDif.Text = Format(0, "#0.00") '*** PEAC 20190726
    Me.txtPreVnta.Enabled = False
    Me.txtFecAdju.Text = "__/__/____"
    Me.txtDsctoInt.Text = Format(0, "#0.00") 'PEAC 20200903
    CmbForPag.ListIndex = -1 'CTI4 ERS0112020
    txtCuentaCargo.NroCuenta = "" 'CTI4 ERS0112020
    LblNumDoc.Caption = "" 'CTI4 ERS0112020
    cmdGrabar.Enabled = False 'CTI4 ERS0112020
    sNumTarj = ""  'CTI4 ERS0112020
End Sub

'Busca el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)
    Dim loCalculos As COMNColoCPig.NCOMColPCalculos
    Dim lbok As Boolean
    Dim lrValida As ADODB.Recordset
    Dim loValContrato As COMNColoCPig.NCOMColPValida
    Dim lnPrecioMercado As Currency
    'Dim lnPrecioBase As Currency, lnDeudatotal As Currency, fdVarFecVencimiento As Date
    Dim lnPrecioBase As Currency, fdVarFecVencimiento As Date
    Dim lrValorVenta As ADODB.Recordset '*** PEAC 20180323

    ''*** PEAC 20090303 ***********************
    Dim loCalcula As NCOMColPCalculos
    Dim vdiasAtraso As Integer, vDiasCustodia As Integer
    Dim vdiasAtrasoMora As Integer
    Dim lnIntVencido As Double
    Dim lnIntMoratorio As Double, lnSaldo As Double, lnCostoNotif As Double
    Dim lnIntAdelantado As Double, vCostoCustodiaMoratorio As Currency
    Dim lnDeuda As Currency
    Dim lsmensaje As String
    Dim lnAdjuCustDif As Double '*** PEAC 20190726
    
    'JAOR20210223 Porcentaje Descuento campaña
    Dim rsBusca As ADODB.Recordset
    Dim rsUpdate As ADODB.Recordset
    Dim valBuscaCamp As COMNColoCPig.NCOMColPValida
    Dim updPorcDesc As COMDColocPig.DCOMColPActualizaBD
    Dim nPorcDesc As Integer
    Dim nEstadoCamp As Integer
    Dim sResultModal As String
    Dim bEsDecimal As Boolean 'Evitar desbordamiento
    'Dim nTipoCamp As Integer
        
    nPorcDesc = 0
    nEstadoCamp = 0
    nTipoCamp = 0
    bEsDecimal = False
    
    '*** PEAC 20200729
    lnDsctoIntTotal = 0
    lbCampanaDscoInt = 0
    lbDscoInt = 0
    
    'On Error GoTo ControlError

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        Set lrValida = loValContrato.nValidaRegVentaSubastaCredPignoraticioSINSubasta(psNroContrato, "A", fsVarNroSubasta, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
    If lrValida Is Nothing Then ' Hubo un Error
        Limpiar
        Set lrValida = Nothing
        Exit Sub
    End If
    
    'JAOR2021 - Buscar si crédito aplica para campaña
    Set rsBusca = New ADODB.Recordset
    Set valBuscaCamp = New COMNColoCPig.NCOMColPValida
    Set rsBusca = valBuscaCamp.BuscaCredPignoCampana(psNroContrato, nEstadoCamp)
    
    If (nEstadoCamp = 1) Then
         Do
         bEsDecimal = False
            If (rsBusca!cTipBeneficio = "A") Then
                sResultModal = InputBox("El crédito califica para un descuento máximo de " & rsBusca!nPorcMax & " %" & vbNewLine & vbNewLine _
                & " " & "+ MANTENER CALIFICACIÓN:   [  1% - 40%]  " & vbNewLine _
                & " " & "+ AFECTAR" & space(5) & "CALIFICACIÓN:   [41% - 80%] " & vbNewLine & vbNewLine _
                & "Ingrese porcentaje de descuento: ", "Crédito apto para descuento", " ")
            ElseIf (rsBusca!cTipBeneficio = "B") Then
                sResultModal = InputBox("El crédito califica para un descuento máximo de " & rsBusca!nPorcMax & " %" & vbNewLine & vbNewLine _
                & " " & "+ MANTENER CALIFICACIÓN:   [  1% - 30%]  " & vbNewLine _
                & " " & "+ AFECTAR" & space(5) & "CALIFICACIÓN:   [31% - 50%] " & vbNewLine & vbNewLine _
                & "Ingrese porcentaje de descuento: ", "Crédito apto para descuento", " ")
            ElseIf (rsBusca!cTipBeneficio = "C") Then
                sResultModal = InputBox("El crédito califica para un descuento máximo de " & rsBusca!nPorcMax & " %" & vbNewLine & vbNewLine _
                & " " & "+ MANTENER CALIFICACIÓN:   [  1% - 20%]  " & vbNewLine _
                & " " & "+ AFECTAR" & space(5) & "CALIFICACIÓN:   [21% - 40%] " & vbNewLine & vbNewLine _
                & "Ingrese porcentaje de descuento: ", "Crédito apto para descuento", " ")
            End If
            
           
           
            If Trim(sResultModal) = "" Then
                MsgBox "Inserte valor porcentual", vbInformation, "Action Cancelled"
                Exit Sub
            ElseIf IsNumeric(sResultModal) Then
                'If CInt(sResultModal) > CDbl(sResultModal) Then
                If sResultModal > 100 Then
                     MsgBox "No se permiten valores mayores a 100" & vbCrLf & "Intente de nuevo", vbInformation, " "
                ElseIf sResultModal <> CInt(sResultModal) Then
                    bEsDecimal = True
                    MsgBox "No se aceptan decimales " & vbCrLf & "Intente de nuevo", vbInformation, " "
                ElseIf (sResultModal > rsBusca!nPorcMax) Then
                   MsgBox "El monto ingresado es mayor al máximo permitido " & vbCrLf & "Intente de nuevo", vbInformation, " "
                ElseIf (sResultModal < 0) Then
                    MsgBox "No se aceptan valores negativos " & vbCrLf & "Intente de nuevo", vbInformation, " "
                Else
                    nPorcDesc = sResultModal
                End If
            Else
                MsgBox "Solo se permiten números enteros positivos" & vbCrLf & "Try Again", vbInformation, "Intergers Only"
            End If
            
         Loop While (Trim(sResultModal) = "") Or (sResultModal < 0) Or (sResultModal > 100) Or (bEsDecimal = True) Or (sResultModal > rsBusca!nPorcMax)
        
        'Determinar si mantiene o afecta calificación
        If (rsBusca!cTipBeneficio = "A") Then
            If (nPorcDesc <= 40) Then
                 nTipoCamp = 1
            ElseIf (nPorcDesc >= 41 And nPorcDesc <= 80) Then
                  nTipoCamp = 2
            End If
        ElseIf (rsBusca!cTipBeneficio = "B") Then
            If (nPorcDesc <= 30) Then
                 nTipoCamp = 1
            ElseIf (nPorcDesc >= 31 And nPorcDesc <= 50) Then
                 nTipoCamp = 2
            End If
            
        ElseIf (rsBusca!cTipBeneficio = "C") Then
            If (nPorcDesc <= 20) Then
                 nTipoCamp = 1
            ElseIf (nPorcDesc >= 21 And nPorcDesc <= 40) Then
                 nTipoCamp = 2
            End If
        End If
        
        Set updPorcDesc = New COMDColocPig.DCOMColPActualizaBD
        Call updPorcDesc.dUpdateColocPigDescInteres(psNroContrato, nPorcDesc)
    
    End If 'END nEstadoCamp=1
 
    'PEAC 20200729
    lbCampanaDscoInt = lrValida!nAplicaCampDsctoInt
    lbDscoInt = lrValida!nAplicaCredDscto
    'FIN PEAC
    
    
    '**Comentado por DAOR 20070915, Según Memo 1492-2007-CMAC-M
    'If lrValida!bExcepVta = 0 Then
    '    If DateDiff("d", lrValida!dPrdEstado, gdFecSis) > 60 Then
    '            MsgBox "Contrato con mas de 60 dias de adjudicado para ser recuperado.", vbOKOnly + vbInformation, "AVISO"
    '            Limpiar
    '            Set lrValida = Nothing
    '            Exit Sub
    '    End If
    'End If
    
    
    ''*** PEAC 20090303
    'lnDeudatotal = loValContrato.CalcDeudaRemate(lrValida!dprocesorem, lrValida!agerem, lrValida!cnroprocesorem, lrValida!cCtaCod)
    'fdVarFecVencimiento = lrValida!dprocesorem 'JAOR20210216 COMENTÓ
    fdVarFecVencimiento = lrValida!dPrdEstado 'JAOR20210216 AGREGÓ - ACTA014-2021
    
    Set loValContrato = Nothing

    '*** PEAC 20201028
    vdiasAtraso = DateDiff("d", lrValida!dVenc, fdVarFecVencimiento)
    vdiasAtrasoMora = lrValida!nDiasMora
    
    Set loCalcula = New NCOMColPCalculos
        If vdiasAtraso <= 0 Then
            vdiasAtraso = 0
            lnIntVencido = 0
            lnIntAdelantado = 0
            lnIntMoratorio = 0
            vCostoCustodiaMoratorio = 0
        Else
            lnIntAdelantado = loCalcula.nCalculaInteresAlVencimiento(lrValida!nSaldo, lrValida!nTasaIntVenc, 30)
            lnIntAdelantado = Round(lnIntAdelantado, 2)
            
            lnIntVencido = loCalcula.nCalculaInteresMoratorio(lrValida!nSaldo, lrValida!nTasaIntVenc, vdiasAtraso, lnIntAdelantado)
            lnIntVencido = Round(lnIntVencido, 2)
            
'            lnIntVencido = loCalcula.nCalculaInteresMoratorio(lrValida!nSaldo, lrValida!nTasaIntVenc, vdiasAtraso)
'            lnIntVencido = Round(lnIntVencido, 2)
          
            lnIntMoratorio = loCalcula.nCalculaInteresMoratorio(lrValida!nSaldo, lrValida!nTasaIntMora, vdiasAtrasoMora)
            lnIntMoratorio = Round(lnIntMoratorio, 2)
                    
'            lnIntAdelantado = loCalcula.nCalculaInteresAlVencimiento(lrValida!nSaldo, lrValida!nTasaIntVenc, 30)
'            lnIntAdelantado = Round(lnIntAdelantado, 2)
        
        End If
    Set loCalcula = Nothing
    
    ''*** PEAC 20090227 *******************************
    'lnDeuda = loCalcula.nCalculaDeudaPignoraticio(!nsaldo, !dVenc, !nTasacion, !ntasaintvenc, pnTasaCustodiaVencida, pnTasaImpuesto, !nPrdEstado, pnTasaPreparacionRemate, gdFecSis)
    
    lnSaldo = lrValida!nSaldo
    lnCostoNotif = lrValida!nNotificacion
    lnAdjuCustDif = lrValida!nAdjudicaCustDif '*** PEAC 20190726
    
    'lnDeudatotal = lnSaldo + Round(lnIntMoratorio, 2) + Round(lnIntVencido, 2) + Round(lnIntAdelantado, 2) + lnCostoNotif
    lnDeudatotal = lnSaldo + Round(lnIntMoratorio, 2) + Round(lnIntVencido, 2) + Round(lnIntAdelantado, 2) + lnCostoNotif + lnAdjuCustDif '***PEAC 20190726
    lnDeudatotal = Round(lnDeudatotal, 2)
    '*********************************************************
    
    'Muestra Datos
    lbok = fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, False)
    
   'txtDeuda = Format(lrValida!nDeuda, "#0.00")
    'Obtiene el precio Base de Subasta
    ' Si es vta sin subasta Calcula el precio de mercado
    ' en base al precio de la ultima adjudicacion
    
''*** PEAC 20090303 - COMENTADO PORQUE NO HACE NADA
'    If fnVtaJoyaSINSubasta = 1 Then
'        Dim loCalcula As COMNColoCPig.NCOMColPCalculos
'         Set loCalcula = New COMNColoCPig.NCOMColPCalculos
'            'lnPrecioMercado = loCalcula.nCalculaPrecioMercadoVtaAdjudicadoSINSubasta(psNroContrato)
'            'lnPrecioBase = loCalcula.nCalculaValorVentaAdjudicadoSINSubasta(lrValida!nSaldo, lnPrecioMercado, fnVarTasaIGV)
'         Set loCalcula = Nothing
'    End If


    'txtPreBaseVenta = Format(lnPrecioBase, "#0.00")
    'txtPreVentaBruta = Format(lnPrecioBase, "#0.00")
     
    '*** PEAC 20090303 - ******************************
    Me.txtCapital.Text = lrValida!nSaldo
    Me.txtIntComp.Text = lnIntAdelantado
    Me.txtIntVen.Text = lnIntVencido
    Me.txtMora.Text = lnIntMoratorio
    Me.txtCostoNoti.Text = lrValida!nNotificacion
    Me.txtCustodiaDif.Text = lrValida!nAdjudicaCustDif
    
    '***************************************************
    
    txtDeuda.Text = lnDeudatotal

    '*** PEAC 20180320
    Set lrValorVenta = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        Set lrValorVenta = loValContrato.nCobroPrecioVentaLoteAdjudicado(psNroContrato, lnDeudatotal, Format(gdFecSis, "yyyymmdd"), lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
    If lrValorVenta Is Nothing Then ' Hubo un Error
        Limpiar
        Set lrValorVenta = Nothing
        Exit Sub
    End If

    Me.txtPreVnta.Enabled = True
    
    lnPreVta = lrValorVenta!IntComp + lrValorVenta!IntMora
    lnDevol = lrValorVenta!nMontoDevuelto
    
    Me.txtFecAdju = Format(lrValorVenta!dPrdEstado, "dd/MM/yyyy")
    Me.txtDiasTrans = lrValorVenta!nDiastrans
    Me.txtPreVnta = Format(lnPreVta, "#0.00")
    Me.txtDevol = Format(lnDevol, "#0.00")
    
    lnPreVtaInicial = lnPreVta
    
    'Me.txtPreVntaFinal = Format(lnPreVta + lnDevol, "#0.00")
    
    '*** FIN PEAC
    
    '*** PEAC 20180320
    'txtPreVentaBruta = lnDeudatotal
    'txtPreVentaBruta = lnDeudatotal + (lnPreVta + lnDevol)
   
    
    '*** FIN PEAC
    
    fnVarDiasAtraso = DateDiff("d", fdVarFecVencimiento, gdFecSis)
       
'*** PEAC 20090303 - se comento porque estos calculos serán reeemplazados por otros arriba mencionados
'    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
'            fnVarInteresVencido = loCalculos.nCalculaInteresMoratorio(lnDeudatotal, fnVarTasaInteresVencido, fnVarDiasAtraso)
'            fnVarInteresVencido = Round(fnVarInteresVencido, 2)
'            TxtInteres.Text = Format(fnVarInteresVencido, "#0.00")
'            txtPreVentaBruta.Text = Format(fnVarInteresVencido + lnDeudatotal)
'            fnVarPrecioAdjudica = Format(lrValida!nValRegistroAdj, "#0.00")
'    Set lrValida = Nothing

'*** INICIO PEAC 20180905
    If lnDevol = 0 Then
        Set lrValida = New ADODB.Recordset
        Set loValContrato = New COMNColoCPig.NCOMColPValida
            Set lrValida = loValContrato.BuscaDevCredPignoPorAbonar(psNroContrato, lsmensaje)
            If Trim(lsmensaje) <> "" Then
                 MsgBox lsmensaje, vbInformation, "Aviso"
                 'Exit Sub
            End If
    '    If lrValida Is Nothing Then ' Hubo un Error
    '        Limpiar
    '        Set lrValida = Nothing
    '        Exit Sub
    '    End If
    End If
'*** FIN PEAC 20180905
        
    Me.txtCodAdj.Text = AXDesCon.listaClientes.ListItems(1).Text
    Me.txtNomAdj.Text = AXDesCon.listaClientes.ListItems(1).SubItems(1)
    Me.txtNatAdj.Text = AXDesCon.listaClientes.ListItems(1).SubItems(7)
    Me.txtTriAdj.Text = AXDesCon.listaClientes.ListItems(1).SubItems(9)
    txtPreVentaBruta.Enabled = True
    'txtPreVentaBruta.SetFocus
     
    cmdGrabar.Enabled = True
    
    AXCodCta.Enabled = False
    cmdBuscar.Enabled = True
    'cmdBuscar.SetFocus
    CmbForPag.Enabled = True 'CTI4 ERS0112020
    CmbForPag.ListIndex = IndiceListaCombo(CmbForPag, 1) 'CTI4 ERS0112020

   Call CalculaDsctoInt(Me.txtIntComp, Me.txtIntVen, Me.txtMora, Me.txtPreVnta, psNroContrato) 'PEAC 20200729 'JAOR20210223 AGREGO psNroContrato
    Call CalculaTotalesVta(lnDeudatotal, lnPreVta, lnDevol, lnDsctoIntTotal) 'PEAC 20200729

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
'PEAC 20200727
Private Sub CalculaDsctoInt(ByVal pnIntComp As Double, ByVal pnIntVenc As Double, ByVal pnIntMora As Double, ByVal pnIntAdic As Double, Optional ByVal psNroContrato As String)
    Dim lnIntComp As Double
    Dim lnIntVenc As Double
    Dim lnIntMora As Double
    Dim lnIntAdic As Double
    Dim lnTotInt As Double
    Dim lnDsctoInt As Double
    
    Dim lrValida As ADODB.Recordset
    
    lnPorcenDsctoInt = 0
    If lbCampanaDscoInt = 1 And lbDscoInt = 1 Then
    
        lnIntComp = pnIntComp: lnIntVenc = pnIntVenc: lnIntAdic = pnIntAdic: lnIntMora = pnIntMora
        lnTotInt = lnIntComp + lnIntVenc + lnIntAdic + lnIntMora
    
        Set lrValida = New ADODB.Recordset
        Dim loParam As COMDColocPig.DCOMColPCalculos
        Set loParam = New COMDColocPig.DCOMColPCalculos
            Set lrValida = loParam.nObtieneDsctoInt(lnTotInt, psNroContrato)
        Set loParam = Nothing
        
        lnPorcenDsctoInt = lrValida!nPorcen
        
        Me.txtDsctoInt = lrValida!nDsctoInt
        Me.lblDscto = "Dscto.Int." + Str(lnPorcenDsctoInt) + "%:"
        
        lnDsctoIntTotal = lrValida!nDsctoInt
    End If
    
    Set lrValida = Nothing
End Sub

Private Sub CalculaTotalesVta(ByVal pnDeudatotal As Double, ByVal pnPreVta As Double, ByVal pnDevol As Double, Optional ByVal pnDsctoInt As Double = 0, Optional ByVal pnITF As Double = 0)
    Me.txtPreVnta.Text = Format(pnPreVta, "#0.00")
    Me.txtPreVntaFinal = Format(pnPreVta + pnDevol, "#0.00")
    'txtPreVentaBruta = pnDeudatotal + (pnPreVta + pnDevol)
    txtPreVentaBruta = pnDeudatotal + (pnPreVta + pnDevol) - pnDsctoInt 'PEAC 20200903
    Me.TxtMontoTotal = pnDeudatotal + (pnPreVta + pnDevol) - pnDsctoInt + lnITF 'PEAC 20200729
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call BuscaContrato(AXCodCta.NroCuenta)
End Sub

'Busca el Adjudicatario por nombre y/o documento
Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String
Dim liFil As Integer
Dim ls As String
On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
Set loPers = frmBuscaPersona.Inicio

If Not loPers Is Nothing Then
    lsPersCod = loPers.sPersCod
    Me.txtCodAdj = loPers.sPersCod
    txtNomAdj = PstaNombre(loPers.sPersNombre, False)
    
    txtPreVentaBruta.Enabled = True
    'txtPreVentaBruta.SetFocus

End If

Set loPers = Nothing
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Cancela el proceso actual e inicializa uno nuevo
Private Sub cmdCancelar_Click()
    Limpiar
    DesactivaBoton (False)
    AXCodCta.Enabled = True
    CmbForPag.Enabled = False 'CTI4 ERS0112020
    AXCodCta.SetFocus
End Sub

'Graba los cambios en la base de datos
Private Sub cmdGrabar_Click()

'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarVta As COMNColoCPig.NCOMColPContrato 'NColPContrato
Dim oMov As COMDMov.DCOMMov

Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lnVtaNeta As Currency
Dim lnImpuestoVtaSub As Currency
Dim lnCostVtaAdj As Currency

Dim lnMovNro As Long
Dim lnMovNroResult As Long 'RECO20160226 ERS040-2015

Dim lsBoletaCargo  As String 'CTI4 ERS0112020
Dim MatDatosAho(14) As String 'CTI4 ERS0112020
Dim lsNombreClienteCargoCta As String 'CTI4 ERS0112020

'*** PEAC 20180713
'Dim lbResultadoVisto As Boolean
'Dim sPersVistoCod  As String
'Dim sPersVistoCom As String
'Dim loVistoElectronico As frmVistoElectronico
'Set loVistoElectronico = New frmVistoElectronico
'************************************************

If Len(txtCodAdj) <> 13 Then
    MsgBox " Falta ingresar el Adjudicatario ", vbInformation, " Aviso "
    cmdBuscar.Enabled = True
    'cmdBuscar.SetFocus
ElseIf CCur(txtPreVentaBruta) <= 0 Then
    MsgBox " Falta ingresar el precio de venta bruta ", vbInformation, " Aviso "
    'txtPreVentaBruta.SetFocus
'ElseIf Len(Trim(txtNroDocumento)) <> 8 Then
'    MsgBox " Falta ingresar el número de documento ", vbInformation, " Aviso "
'    txtNroDocumento.SetFocus
End If

'asigna valores a variables
lnVtaNeta = Round(Val(txtPreVentaBruta.Text) / (1 + fnVarTasaIGV), 2)
lnImpuestoVtaSub = Val(txtPreVentaBruta.Text) - lnVtaNeta

'If lnPreVtaInicial <> lnPreVta Then
'    MsgBox "El valor ingresado "
'    lbResultadoVisto = loVistoElectronico.Inicio(3, "122900")
'    If Not lbResultadoVisto Then Exit Sub
'End If

If Not ValidaFormaPago Then Exit Sub 'CTI4 ERS0112020

If MsgBox(" Grabar Venta de Joyas Adjudicadas ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        Dim lsPersDirecAct As String
        cmdGrabar.Enabled = False
        cmdBuscar.Enabled = False
        txtPreVentaBruta.Enabled = False
        txtNroDocumento.Enabled = False
        lsPersDirecAct = frmPersActualizaDireccion.Iniciaformulario(Me.AXDesCon.listaClientes.ListItems(1).Text, AXCodCta.NroCuenta)   'RECO20160329 ERS040-2015
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarVta = New COMNColoCPig.NCOMColPContrato
            'Grabar Venta de Remate
            
        'CTI4 ERS0112020
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then lsNombreClienteCargoCta = PstaNombre(loGrabarVta.ObtieneNombreTitularCargoCta(txtCuentaCargo.NroCuenta))
        'end CTI4
            
            'Grabar Venta de CREDITO ADJUDICADO
            Call loGrabarVta.nSubastaVentaCredPignoraticioSINSubasta(AXCodCta.NroCuenta, fsVarNroSubasta, lsFechaHoraGrab, _
                 lsMovNro, CCur(Val(Me.txtPreVentaBruta.Text)), CCur(Val(Me.txtDeuda.Text)), 0, fnVarPrecioAdjudica, _
                 Val(Me.AXDesCon.Oro14), Val(Me.AXDesCon.Oro16), Val(Me.AXDesCon.Oro18), Val(Me.AXDesCon.Oro21), False, txtNroDocumento.Text, Me.AXDesCon.listaClientes.ListItems(1).Text, _
                  gITF.gbITFAplica, gITF.gbITFAsumidocreditos, CCur(Val(Me.TxtItf.Text)), CCur(Val(Me.TxtInteres.Text)) _
                 , True, , lnMovNroResult, lnPreVta, lnDevol, CCur(Val(Me.txtDsctoInt.Text)), CInt(Trim(Right(CmbForPag.Text, 10))), nMovNroRVD, nMovNroRVDPend, txtCuentaCargo.NroCuenta, MatDatosAho)
            '*** PEAC 20200811 - se agregó CCur(Val(Me.txtDsctoInt.Text))
            '*** PEAC 20180712 - se agregó lnPreVta, lnDevol
            '*** PEAC 20190717 - SE CAMBIO "CCur(Val(Me.TxtMontoTotal.Text))" POR "CCur(Val(Me.txtPreVentaBruta.Text))"

            'PEAC 20200904
            If CCur(Val(Me.txtDsctoInt.Text)) > 0 Then
                
                'JAOR20210225
                If (nTipoCamp = 0 Or nTipoCamp = 1) Then
                    loGrabarVta.RegistroPignoSegmentacion Me.AXDesCon.listaClientes.ListItems(1).Text, lsMovNro
                ElseIf (nTipoCamp = 2) Then
                    loGrabarVta.AfectarPignoSegmentacion AXCodCta.NroCuenta
                End If
            
            End If
            
           '*** BRGO 20110915 ***************************
           If gITF.gbITFAplica Then
                Set oMov = New COMDMov.DCOMMov
                Call oMov.InsertaMovRedondeoITF(lsMovNro, 1, CCur(Val(Me.TxtItf.Text)) + nRedondeoITF, CCur(Val(Me.TxtItf.Text)))
           End If
           Set oMov = Nothing
           '*********************************************

        Set loGrabarVta = Nothing

        'Impresión
        If MsgBox(" Desea realizar impresión de Recibo ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            Dim oImpColP As COMNColoCPig.NCOMColPImpre
            Dim lsCadImp As String
            Set oImpColP = New COMNColoCPig.NCOMColPImpre
                lsCadImp = oImpColP.ImpRecupSub(AXCodCta.NroCuenta, txtNomAdj.Text, CCur(Val(txtPreVentaBruta.Text)), CCur(Val(Me.TxtMontoTotal.Text)), CCur(Val(Me.TxtItf.Text)), gsNomAge, gdFecSis, gsCodUser, gImpresora, gbImpTMU, CCur(Val(Me.txtDsctoInt.Text)))
           'CTI4 ERS0112020
            If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
                lsBoletaCargo = oImpColP.ImprimeBoletaAhorro("RETIRO AHORROS", "CARGO A CUENTA POR REC. CO. ADJ", "", CStr(CCur(Val(Me.txtPreVentaBruta.Text)) + CCur(Val(Me.TxtItf.Text))), lsNombreClienteCargoCta, txtCuentaCargo.NroCuenta, "", CDbl(MatDatosAho(10)), CDbl(MatDatosAho(3)), "", 1, CDbl(MatDatosAho(11)), , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, gsCodCMAC, 0, , , gbImpTMU)
            End If
            'END CTI4
            Set oImpColP = Nothing
            
            Dim loPrevio As previo.clsprevio
            Set loPrevio = New previo.clsprevio
                loPrevio.PrintSpool sLpt, lsCadImp, False 'CTI4 ERS0112020
            Set loPrevio = Nothing

            Do While True
                If MsgBox("Desea reimprimir boleta de la operación?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    Set loPrevio = New previo.clsprevio
                        loPrevio.PrintSpool sLpt, lsCadImp, False
                    Set loPrevio = Nothing
                Else
                    Exit Do
                End If
            Loop
            

            'CTI4 ERS0112020
            If Trim(lsBoletaCargo) <> "" Then
            Set loPrevio = New previo.clsprevio
            loPrevio.PrintSpool sLpt, lsBoletaCargo, False
            Set loPrevio = Nothing
            
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
        End If
        
        'Call ImprimeComprobanteAdj(False, AXCodCta.NroCuenta, txtNomAdj.Text, AXDesCon, Val(txtDeuda.Text), lnMovNroResult, , , , , lsPersDirecAct)
        'Call ImprimeComprobanteAdj(False, AXCodCta.NroCuenta, txtNomAdj.Text, txtCodAdj.Text, AXDesCon, Val(txtDeuda.Text), lnMovNroResult, , , , , lsPersDirecAct) 'NAGL ERS 012-2017 Agregó:txtCodAdj.Text
        Call ImprimeComprobanteAdj(False, AXCodCta.NroCuenta, txtNomAdj.Text, txtCodAdj.Text, AXDesCon, CCur(Val(Me.txtPreVentaBruta.Text)), lnMovNroResult, , , , , lsPersDirecAct, CCur(Val(Me.txtDsctoInt.Text))) '*** PEAC 20180905
        nRedondeoITF = 0
        
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
                oSolicitud.ActualizarCapAutSinTarjetaVisto_nMovNro gsCodUser, gsCodAge, txtCuentaCargo.NroCuenta, rsCli!cperscod, nMovNroOperacion, CStr(gAhoCargoCtaRecupCoAdj)
            End If
            Set oMovOperacion = Nothing
            nRespuesta = 0
        End If
        'CTI4 end

        '*** PEAC 20180713
        Limpiar
        DesactivaBoton (False)
        AXCodCta.Enabled = True
        AXCodCta.SetFocus
        
'************ Registrar actividad de opertaciones especiales - ANDE 2017-12-18
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim RVerOpe As ADODB.Recordset
    Dim nEstadoActividad As Integer
    nEstadoActividad = oCaptaLN.RegistrarActividad(gsOpeCod, gsCodUser, gdFecSis)
  
    If nEstadoActividad = 1 Then
        MsgBox "He detectado un problema; su operación no fue afectada, pero por favor comunciar a TI-Desarrollo.", vbError, "Error"
    ElseIf nEstadoActividad = 2 Then
        MsgBox "Ha usado el total de operaciones permitidas para el día de hoy. Si desea realizar más operaciones, comuníquese con el área de Operaciones.", vbInformation + vbOKOnly, "Aviso"
        Unload Me
    End If
    
    ' END ANDE ******************************************************************

        
    
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub DesactivaBoton(ByVal pbEstado As Boolean)
    cmdGrabar.Enabled = pbEstado
    cmdBuscar.Enabled = pbEstado
    txtPreVentaBruta.Enabled = pbEstado
    txtNroDocumento.Enabled = pbEstado
End Sub


'Finaliza el formulario
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
        End If
    End If
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
Limpiar
Call CargaControles 'CTI4 ERS0112020
End Sub
Private Sub txtNomAdj_KeyPress(KeyAscii As Integer)
 KeyAscii = SoloLetras(KeyAscii)
End Sub

'Valida el campo txtnrodocumento
Private Sub txtNroDocumento_KeyPress(KeyAscii As Integer)
Dim loValid As COMNColoCPig.NCOMColPValida
Dim lbExiste As Boolean
If KeyAscii = 13 And Len(Trim(txtNroDocumento)) = 8 Then
    Set loValid = New COMNColoCPig.NCOMColPValida
        lbExiste = loValid.nDocumentoEmitido(3, txtNroDocumento.Text, "'" & geColPVtaSubasta & "'")
    Set loValid = Nothing
    If lbExiste = True Then
        MsgBox "Número de Boleta duplicada" & vbCr & "Ingrese un número diferente", vbInformation, " Aviso "
    Else
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
    End If
End If
End Sub

Private Sub txtPreVentaBruta_Change()
Dim loValida As COMNColoCPig.NCOMColPValida 'CTI4ERS0112020
Dim bEsMismoTitular As Boolean 'CTI4 ERS0112020
    Set loValida = New COMNColoCPig.NCOMColPValida

bEsMismoTitular = loValida.EsMismoTitulardeCuentaPignoYAhorro(txtCuentaCargo.NroCuenta, AXCodCta.NroCuenta)
If IsNumeric(txtPreVentaBruta.Text) Then

     If gITF.gbITFAplica And Not bEsMismoTitular Then
        
        If Not gITF.gbITFAsumidocreditos Then
            'Me.TxtMontoTotal.Text = "0.00"
            
            
              'Dim aux As String
              '  If InStr(1, CStr(TxtItf), ".", vbTextCompare) > 0 Then
              '       aux = CDbl(CStr(Int(TxtItf)) & "." & Mid(CStr(TxtItf), InStr(1, CStr(TxtItf), ".", vbTextCompare) + 1, 2))
              '  Else
              '       aux = CDbl(CStr(Int(TxtItf)))
              '  End If
                 TxtItf.Text = Format(TxtItf.Text, "#0.00")
                 Me.txtDsctoInt.Text = Format(txtDsctoInt.Text, "#0.00")
                 
                 'TxtItf.Text = Format(gITF.fgITFCalculaImpuesto(Me.txtPreVentaBruta.Text))
                 TxtItf.Text = Format(gITF.fgITFCalculaImpuesto(Me.txtPreVentaBruta.Text - Me.txtDsctoInt.Text))  'PEAC 20200828
                 '*** BRGO 20110908 ************************************************
                 nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtItf.Text))
                 If nRedondeoITF > 0 Then
                    Me.TxtItf.Text = Format(CCur(Me.TxtItf.Text) - nRedondeoITF, "#,##0.00")
                 End If
                '*** END BRGO
                lnITF = CCur(Val(Me.TxtItf.Text))
                 TxtMontoTotal.Text = CCur(Val(txtPreVentaBruta.Text)) + CCur(TxtItf.Text)
        Else
                 'Me.TxtItf = gITF.fgITFCalculaImpuesto(Me.txtPreVentaBruta.Text)
                 Me.TxtItf = gITF.fgITFCalculaImpuesto(Me.txtPreVentaBruta.Text - Me.txtDsctoInt.Text)  'PEAC 2020828 se agregó Me.txtDsctoInt.Text
                 
                 '*** BRGO 20110908 ************************************************
                 nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtItf.Text))
                 If nRedondeoITF > 0 Then
                    Me.TxtItf.Text = Format(CCur(Me.TxtItf.Text) - nRedondeoITF, "#,##0.00")
                 End If
                '*** END BRGO
                lnITF = CCur(Val(Me.TxtItf.Text))
                 Me.TxtMontoTotal = Format(CDbl(Me.txtPreVentaBruta.Text), "#0.00")
        End If
    Else
                 Me.TxtItf = Format(0, "#0.00")
                 TxtMontoTotal = Format(CDbl(Me.txtPreVentaBruta.Text), "#0.00")
    End If
End If

End Sub

'Valida el campo txtpreventabruta
Private Sub txtPreVentaBruta_GotFocus()
    fEnfoque txtPreVentaBruta
End Sub
Private Sub txtPreVentaBruta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreVentaBruta, KeyAscii)
If KeyAscii = 13 Then
'    If Val(txtPreVentaBruta.Text) < Val(txtPreBaseVenta.Text) Then
'        MsgBox " Precio Venta debe ser mayor a Precio Base ", vbInformation, " Aviso "
'        txtPreVentaBruta.SetFocus
'    Else
        txtNroDocumento.Enabled = True
        'txtNroDocumento.SetFocus
'    End If
End If
End Sub
Private Sub txtPreVentaBruta_LostFocus()
txtPreVentaBruta = Format(Val(txtPreVentaBruta), "#0.00")
VeriVenBru
End Sub

'Procedimiento de verificación de la venta bruta
Private Sub VeriVenBru()
'    If Val(txtPreVentaBruta.Text) < Val(txtPreBaseVenta.Text) Then
'        MsgBox " Precio Venta debe ser mayor a Precio Base ", vbInformation, " Aviso "
'        txtPreVentaBruta.SetFocus
'    End If
End Sub

Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Set loParam = New COMDColocPig.DCOMColPCalculos
    fnVarTasaPreparacionRemate = loParam.dObtieneColocParametro(gConsColPTasaPreparaRemate)
    fnVarTasaImpuesto = loParam.dObtieneColocParametro(gConsColPTasaImpuesto)
    fnVarTasaIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)
    fnVarTasaInteresVencido = loParam.dObtieneTasaInteres("01011130501", gColocLineaCredTasasIntMoratNormal)
    'pAgeRemSub = Right(ReadVarSis("CPR", "cAgeRemSub"), 2)
    'pDifeDiasRema = Val(ReadVarSis("CPR", "nDifeDiasRema"))
    
Set loParam = Nothing
Set loConstSis = New COMDConstSistema.NCOMConstSistema
    fnVtaJoyaSINSubasta = loConstSis.LeeConstSistema(123)  '0=Con Subasta // 1=Sin Subasta //
Set loConstSis = Nothing

End Sub

Private Sub txtPreVnta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreVnta, KeyAscii)
    
    
    If KeyAscii = 13 Then
        
        If Not IsNumeric(txtPreVnta.Text) Then
            txtPreVnta.Text = "0.00"
        ElseIf Trim(txtPreVnta.Text) = "" Then
            txtPreVnta.Text = "0.00"
        End If
    
        txtPreVnta.Text = Format(txtPreVnta.Text, "#0.00")

        If bValidaValorVta(CCur(Me.txtDeuda.Text), CCur(Me.txtPreVnta.Text)) Then
            lnPreVta = CCur(Me.txtPreVnta.Text)
            Call CalculaTotalesVentaEnter
            Me.txtPreVentaBruta.SetFocus
        Else
            MsgBox "El monto ingresado no puede ser menor al monto de la deuda...", vbOKOnly + vbExclamation, "Atencion"
            Call CalculaDsctoInt(Me.txtIntComp, Me.txtIntVen, Me.txtMora, Me.txtPreVnta, Me.AXCodCta.NroCuenta)   'JAOR20210302
            Call CalculaTotalesVta(lnDeudatotal, lnPreVtaInicial, lnDevol, lnDsctoIntTotal)
        End If
    End If
End Sub
'*** PEAC 20180713
Private Function bValidaValorVta(ByVal pnValorDeuda As Double, ByVal pnValorVta As Double) As Boolean
    bValidaValorVta = True
    If pnValorVta < pnValorDeuda Then
        bValidaValorVta = False
    End If
End Function

Private Sub CalculaTotalesVentaEnter()
Dim loValida As COMNColoCPig.NCOMColPValida 'CTI4ERS0112020
Dim bEsMismoTitular As Boolean 'CTI4 ERS0112020
    Set loValida = New COMNColoCPig.NCOMColPValida

bEsMismoTitular = loValida.EsMismoTitulardeCuentaPignoYAhorro(txtCuentaCargo.NroCuenta, AXCodCta.NroCuenta)
    If IsNumeric(txtPreVnta.Text) Then
        
        Call CalculaDsctoInt(Me.txtIntComp, Me.txtIntVen, Me.txtMora, Me.txtPreVnta) 'PEAC 20200729
        Call CalculaTotalesVta(lnDeudatotal, lnPreVta, lnDevol, lnDsctoIntTotal)
    
         If gITF.gbITFAplica And Not bEsMismoTitular Then
            If Not gITF.gbITFAsumidocreditos Then
                     TxtItf.Text = Format(TxtItf.Text, "#0.00")
                     'TxtItf.Text = Format(gITF.fgITFCalculaImpuesto(Me.txtPreVentaBruta.Text))
                     TxtItf.Text = Format(gITF.fgITFCalculaImpuesto(Me.txtPreVentaBruta.Text - Me.txtDsctoInt.Text)) 'PEAC 20200828
                     nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtItf.Text))
                     If nRedondeoITF > 0 Then
                        Me.TxtItf.Text = Format(CCur(Me.TxtItf.Text) - nRedondeoITF, "#,##0.00")
                     End If
                     TxtMontoTotal.Text = CCur(Val(txtPreVentaBruta.Text)) + CCur(TxtItf.Text)
            Else
                     'Me.TxtItf = gITF.fgITFCalculaImpuesto(Me.txtPreVentaBruta.Text)
                     Me.TxtItf = gITF.fgITFCalculaImpuesto(Me.txtPreVentaBruta.Text - Me.txtDsctoInt.Text)
                     nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtItf.Text))
                     If nRedondeoITF > 0 Then
                        Me.TxtItf.Text = Format(CCur(Me.TxtItf.Text) - nRedondeoITF, "#,##0.00")
                     End If
                     Me.TxtMontoTotal = Format(CDbl(Me.txtPreVentaBruta.Text), "#0.00")
            End If
        Else
            Me.TxtItf = Format(0, "#0.00")
            TxtMontoTotal = Format(CDbl(Me.txtPreVentaBruta.Text), "#0.00")
        End If
    End If

End Sub

Private Sub txtPreVnta_LostFocus()
        
        If Not IsNumeric(txtPreVnta.Text) Then
            txtPreVnta.Text = "0.00"
        ElseIf Trim(txtPreVnta.Text) = "" Then
            txtPreVnta.Text = "0.00"
        End If
    
        txtPreVnta.Text = Format(txtPreVnta.Text, "#0.00")
        
        If bValidaValorVta(CCur(Me.txtDeuda.Text), CCur(Me.txtPreVnta.Text)) Then
            lnPreVta = CCur(Me.txtPreVnta.Text)
            Call CalculaTotalesVentaEnter
        Else
            MsgBox "El monto ingresado no puede ser menor al monto de la deuda...", vbOKOnly + vbExclamation, "Atencion"
           Call CalculaDsctoInt(Me.txtIntComp, Me.txtIntVen, Me.txtMora, Me.txtPreVnta, Me.AXCodCta.NroCuenta) 'PEAC 20200729 - JAOR20210302
            Call CalculaTotalesVta(lnDeudatotal, lnPreVtaInicial, lnDevol, lnDsctoIntTotal)
        End If
End Sub
'**
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
                        
            lnTipMot = 21 ' Recuperacion de Contrato Ajudicado Credito Pignoraticio
            oformVou.iniciarFormularioDeposito CInt(Mid(AXCodCta.NroCuenta, 9, 1)), lnTipMot, sGlosa, sIF, sVaucher, nMontoVoucher, sPersCod, nMovNroRVD, nMovNroRVDPend, sNombre, sDireccion, sDocumento, AXCodCta.NroCuenta
            If Len(sVaucher) = 0 Then Exit Sub
            LblNumDoc.Caption = Trim(Mid(sVaucher, 1, Len(sVaucher) - 10))
            cmdGrabar.Enabled = True
        ElseIf CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            Dim sCuenta As String
            
            sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(gAhoCargoCtaRecupCoAdj), sNumTarj, 232, False)
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
            txtPreVentaBruta_Change
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
                        nRespuesta = oSolicitud.SolicitarVistoAtencionSinTarjeta(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gAhoCargoCtaRecupCoAdj))
                    
                        If nRespuesta = 1 Then '1:Tiene Visto de atencion sin tarjeta pendiente de autorizar
                             MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                             Exit Sub
                        End If
                        If nRespuesta = 2 Then '2:Tiene visto aceptado
                            MsgBox "La solicitud de atención sin tarjeta fue Aprobada, proceda con la atención", vbInformation, "Aviso"
                        End If
                        If nRespuesta = 3 Then '3:Tiene visto rechazado
                           If MsgBox("La solicitud de atención sin tarjeta fue RECHAZADA. ¿Desea realizar una nueva solicitud?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                Call oSolicitud.RegistrarVistoDeUsuario(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gAhoCargoCtaRecupCoAdj))
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
                        
                                bExitoSol = frmCapMotivoAutorizacion.Inicio(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gAhoCargoCtaRecupCoAdj))
                                If bExitoSol > 0 Then
                                    MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                                End If
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        lbVistoVal = loVistoElectronico.Inicio(5, gAhoCargoCtaRecupCoAdj)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema"
                            Exit Sub
                        End If
                    End If
                ElseIf lsTieneTarj = "NO" And rsV.RecordCount > 0 Then
                    If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? ", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'add by marg ers 065-2017
                        lbVistoVal = loVistoElectronico.Inicio(5, gAhoCargoCtaRecupCoAdj)
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
        txtPreVentaBruta_Change
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

