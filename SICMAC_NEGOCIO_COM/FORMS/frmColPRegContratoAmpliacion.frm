VERSION 5.00
Begin VB.Form frmColPRegContratoAmpliacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Registro Contrato Ampliacíon"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   Icon            =   "frmColPRegContratoAmpliacion.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   58
      Top             =   5520
      Width           =   9495
      Begin SICMACT.TxtBuscar txtBuscarLinea 
         Height          =   345
         Left            =   1200
         TabIndex        =   59
         Top             =   240
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblLineaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3000
         TabIndex        =   61
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label18 
         Caption         =   "Linea Cred."
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
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdVerRetasacion 
      Caption         =   "Ver"
      Height          =   375
      Left            =   4920
      TabIndex        =   57
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cboTipcta 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmColPRegContratoAmpliacion.frx":030A
      Left            =   5640
      List            =   "frmColPRegContratoAmpliacion.frx":0317
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   120
      Width           =   1830
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   6480
      TabIndex        =   48
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7560
      TabIndex        =   47
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8640
      TabIndex        =   46
      Top             =   6480
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   7080
      TabIndex        =   37
      Top             =   3360
      Width           =   2535
      Begin VB.Label lblInteres 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   45
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Interes"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1560
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   2400
         X2              =   120
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblMontoDesemb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   43
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lbltextMontoDesemb 
         Caption         =   "Desembolso:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblMontoDeudaAct 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   41
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Deuda a la Fecha:"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblMontoBruto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Monto Bruto:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdImpVolTas 
      Caption         =   "Volante de Tasación"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nuevos Valores de Tasación"
      Height          =   975
      Left            =   120
      TabIndex        =   29
      Top             =   4560
      Width           =   6855
      Begin VB.ComboBox cboPlazoNuevo 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmColPRegContratoAmpliacion.frx":0352
         Left            =   5160
         List            =   "frmColPRegContratoAmpliacion.frx":0354
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblOroNetoNew 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   55
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Oro Neto"
         Height          =   255
         Left            =   2160
         TabIndex        =   54
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblFecVen 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5160
         TabIndex        =   53
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Fec. Venc."
         Height          =   255
         Left            =   4200
         TabIndex        =   52
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Plazo (días)"
         Height          =   255
         Left            =   4200
         TabIndex        =   34
         Top             =   285
         Width           =   975
      End
      Begin VB.Label lblMontoPrestamo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   33
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Monto"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   620
         Width           =   615
      End
      Begin VB.Label lblTasaNueva 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Tasación"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraContenedor 
      Caption         =   "Datos de Créditos a Cancelar"
      Height          =   1065
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   6795
      Begin VB.TextBox txtPiezas 
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
         Left            =   3360
         MaxLength       =   5
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cboPlazo 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmColPRegContratoAmpliacion.frx":0356
         Left            =   5580
         List            =   "frmColPRegContratoAmpliacion.frx":035D
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblMontoPresAnt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5600
         TabIndex        =   51
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Prestamo"
         Height          =   255
         Index           =   9
         Left            =   4530
         TabIndex        =   28
         Top             =   615
         Width           =   1335
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Tasación "
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   27
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Plazo  (dias)"
         Height          =   255
         Index           =   8
         Left            =   4560
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Piezas"
         Height          =   210
         Index           =   2
         Left            =   2520
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Oro Neto  (gr)"
         Height          =   210
         Index           =   10
         Left            =   120
         TabIndex        =   24
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Oro Bruto  (gr)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lblOroNeto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblValorTasacion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         TabIndex        =   21
         Top             =   615
         Width           =   1095
      End
      Begin VB.Label lblOroBruto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraPiezasDet 
      Caption         =   "Detalle de Piezas"
      Height          =   2055
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   7335
      Begin SICMACT.FlexEdit FEJoyas 
         Height          =   1695
         Left            =   120
         TabIndex        =   16
         Top             =   225
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   2990
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   2
         EncabezadosNombres=   "Num-Pzas-Material-PBruto-PNetoAnt-Tasac A-Tasac-Descripcion-Item-PNeto"
         EncabezadosAnchos=   "400-450-1030-650-650-700-700-2500-0-650"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-4-X-X-7-X-9"
         ListaControles  =   "0-0-3-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-R-L-R-R-R-R-L-R-C"
         FormatosEdit    =   "0-3-1-2-2-2-2-0-3-2"
         TextArray0      =   "Num"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cliente"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   7335
      Begin VB.Label lblClienteDOI 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5640
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "DOI:"
         Height          =   255
         Left            =   5280
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblClienteNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Height          =   345
      Left            =   3840
      Picture         =   "frmColPRegContratoAmpliacion.frx":0365
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar ..."
      Top             =   120
      Width           =   420
   End
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      Texto           =   "Contrato"
      EnabledCta      =   -1  'True
      Prod            =   "705"
      CMAC            =   "109"
   End
   Begin VB.Label lblCredRetasado 
      Caption         =   "CRÉDITO RETASADO"
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
      TabIndex        =   56
      Top             =   6585
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Tipo de contrato"
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   50
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label lblTituloCalificacion 
      Caption         =   "Última Calificación Según SBS - RCC "
      Height          =   375
      Left            =   7560
      TabIndex        =   14
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblCalificacionNormal 
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
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7560
      TabIndex        =   13
      Top             =   600
      Width           =   1875
   End
   Begin VB.Label Label5 
      Caption         =   "T.E.M."
      Height          =   255
      Left            =   7560
      TabIndex        =   12
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblPorcentajeTasa 
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
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7560
      TabIndex        =   11
      Top             =   2640
      Width           =   1875
   End
   Begin VB.Label lblCalificacionPotencial 
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
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7560
      TabIndex        =   10
      Top             =   960
      Width           =   1875
   End
   Begin VB.Label lblCalificacionDeficiente 
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
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7560
      TabIndex        =   9
      Top             =   1320
      Width           =   1875
   End
   Begin VB.Label lblCalificacionDudoso 
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
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7560
      TabIndex        =   8
      Top             =   1680
      Width           =   1875
   End
   Begin VB.Label lblCalificacionPerdida 
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
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7560
      TabIndex        =   7
      Top             =   2040
      Width           =   1875
   End
End
Attribute VB_Name = "frmColPRegContratoAmpliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'REGISTRO DE CONTRATO PIGNORATICIO AMPLIADO
'Archivo:  frmColPRegContratoAmpliacion
'RECO   :  28/01/2014.
'Resumen:  Formulario Permite registrar ampliaciones de créditos pignoraticios

Option Explicit

'******** Parametros de Colocaciones
Dim fnPorcentajePrestamo As Double
Dim fnImpresionesContrato As Double
Dim fnMinPesoOro As Double
Dim fnMaxMontoPrestamo1 As Double
Dim fnRangoPreferencial As Double

Dim fnTasaInteresAdelantado As Double ' Cambia si es tasa preferencial


Dim fnTasaCustodia As Double
Dim fnTasaTasacion As Double
Dim fnTasaImpuesto As Double
Dim fnTasaPreparacionRemate As Double
Dim fnPrecioOro14 As Double
Dim fnPrecioOro16 As Double
Dim fnPrecioOro18 As Double
Dim fnPrecioOro21 As Double
'****** Variables del formulario
Dim fsCodZona As String * 12
Dim fsContrato As String
Dim fsNroContrato As String * 12
Dim fnJoyasDet As Integer  ' 0 = Sin Detalle (Trujillo) / 1 = Con Detalle (Santa)
Dim lnJoyas As Integer

Dim vFecVenc As Date
Dim vOroBruto As Double
Dim vOroNeto As Double
Dim vPiezas As Integer
Dim vPrestamo As Double
Dim vCostoTasacion As Double
Dim vCostoCustodia As Double
Dim vInteres As Double
Dim vImpuesto As Double
Dim vNetoPagar As Double
'** lstCliente - ListView
Dim lstTmpCliente As ListItem
Dim lsIteSel As String
Dim vContAnte As Boolean
Dim gColocLineaCredPig As String
Dim lsSQL As String
Dim objPista As COMManejador.Pista
Dim fsColocLineaCredPig  As String
Dim fbMalCalificacion As Boolean
Dim nTpoCliente As Integer
Dim lnPesoNetoDesc As Double
Dim fbClienteCPP As Boolean
Dim lsPersCod As String
Dim lsPersApellido As String
Dim lsPersNombre As String
Dim lsPersDireccDomicilio As String
Dim lnSaldoActual As Double
Dim ln14k As Double
Dim ln16k As Double
Dim ln18k As Double
Dim ln21k As Double

Dim lrPersonas As ADODB.Recordset
'RECO COMPLEMENTO
Dim lnVolanteTasacion As Integer
Dim vTasaInteresVencido As Double
Dim vTasaInteresMoratorio As Double
Dim vInteresMoratorio As Double
Dim vEstado  As Integer
Dim vSaldoCapital As Double
Dim vFecVencimiento As Date
Dim vdiasAtraso As Double
Dim vDeuda As Double
Dim vInteresVencido As Double
Dim vCostoCustodiaMoratorio As Double
Dim vCostoPreparacionRemate As Double
Dim fnVarGastoCorrespondencia As Double
Dim fnVarCostoNotificacion As Currency 'PEAC 20070926
Dim vTasaInteres As Double ' peac 20070820
Dim vDiasAdel As Integer
Dim vFecEstado As Date
Dim vInteresAdel As Double
Dim fnTasaCustodiaVencida As Double
Dim fnVarEstUltProcRem As Integer

Dim gcCredAntiguo As String
Dim gnNotifiAdju As Integer
Dim gnNotifiCob As Integer
'RECO FIN*********

Dim vValorTasacion As Double

Dim fsPlazoSist As String 'RECO20150421
Dim sLineaTmp As String 'ALPA20150617**
Dim RLinea As ADODB.Recordset 'ALPA20150617**
Dim lnTasaInicial As Currency 'ALPA20150617**
Dim lnTasaFinal As Currency 'ALPA20150617**
Dim lnTasaGracia As Currency 'ALPA20150617**
Dim lnTasaCompes As Currency 'ALPA20150617**
Dim lnTasaMorato As Currency 'ALPA20150617**

Private Sub cmbVolanTasac_Click()
    Dim lrJoyas  As New ADODB.Recordset
    Dim lnPlazo As Integer
    Dim lnOroBruto As Double
    Dim lnOroNeto As Double
    Dim lnPiezas As Integer
    Dim lnValTasacion As Currency
    Dim lsCadImprimir As String
    Dim loRegImp As COMNColoCPig.NCOMColPImpre
    Dim loPrevio As previo.clsprevio
    
        If ValidarMsh = True Then Exit Sub
        Set lrJoyas = FEJoyas.GetRsNew
        lnPlazo = val(cboPlazo.Text)
        lnOroBruto = val(lblOroBruto.Caption)
        lnOroNeto = val(lblOroNetoNew.Caption)
        lnPiezas = val(txtPiezas.Text)
        lnValTasacion = CCur(Me.lblTasaNueva.Caption)
        If MsgBox("Imprimir Volante de Tasación ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            Set loRegImp = New COMNColoCPig.NCOMColPImpre
                lsCadImprimir = ""
                lsCadImprimir = loRegImp.nPrintVolanteTasacion(lnPlazo, lnOroBruto, lnOroNeto, lnValTasacion, lnPiezas, lrJoyas, gImpresora)
            Set loRegImp = Nothing
            Set loPrevio = New previo.clsprevio
            
            Dim oImp As New ContsImp.clsConstImp
                oImp.Inicia gImpresora
                gPrnSaltoLinea = oImp.gPrnSaltoLinea
                gPrnSaltoPagina = oImp.gPrnSaltoPagina
            Set oImp = Nothing
                loPrevio.PrintSpool sLpt, lsCadImprimir & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea, False
        End If
        Set loPrevio = Nothing
End Sub

Private Sub cmdBuscar_Click()

Dim loPers As comdpersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersContrato As COMDColocPig.DCOMColPContrato
Dim lrContratos As New ADODB.Recordset
Dim loCuentas As comdpersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New comdpersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

lsEstados = "2101,2104,2106,2107" 'Estados Créditos Pigno Vigentes

If Trim(lsPersCod) <> "" Then
    Set loPersContrato = New COMDColocPig.DCOMColPContrato
        Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstados, Mid(gsCodAge, 4, 2))
    Set loPersContrato = Nothing
End If
    
Set loCuentas = New comdpersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrContratos)
    If loCuentas.sCtaCod <> "" And gsCodAge = Mid(loCuentas.sCtaCod, 4, 2) Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        Me.AXCodCta.Enabled = True
        AXCodCta.SetFocusCuenta
    Else
        Me.AXCodCta.CMAC = "109"
        Me.AXCodCta.Prod = "705"
        Me.AXCodCta.Age = "01"
        Me.AXCodCta.Cuenta = ""
        Me.AXCodCta.Enabled = False
        MsgBox "El crédito no pertenece a la agencia", vbCritical, "Aviso"
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdGrabar_Click()
    Dim pbTran As Boolean
    Dim lsCtaReprestamo As String
    Dim lsMovNro As String
    Dim lsFechaHoraGrab As String
    Dim lnMontoPrestamo As Currency
    Dim lnNetoRecibir As Currency
    Dim lnPlazo As Integer
    Dim lsFechaVenc As String
    Dim lnOroBruto As Double
    Dim lnOroNeto As Double
    Dim lnPiezas As Integer
    Dim lnValTasacion As Currency
    Dim lsTipoContrato As String
    Dim lsLote As String
    Dim lnIntAdelantado As Currency, lnCostoTasac As Currency, lnCostoCustodia As Currency, lnImpuesto As Currency
    Dim lrJoyas  As New ADODB.Recordset
    Dim loRegPig As COMNColoCPig.NCOMColPContrato
    Dim loRegImp As COMNColoCPig.NCOMColPImpre
    Dim loContFunct As COMNContabilidad.NCOMContFunciones
    Dim lsContrato As String
    Dim loPrevio As previo.clsprevio
    Dim lnNumImp As Integer
    Dim pnITF As Double
    Dim lsCadImprimir As String
    Dim lsmensaje As String
    Dim HojaResumen1 As String
    Dim HojaResumen2 As String
    Dim lbResultadoVisto As Boolean
    Dim sPersVistoCod  As String
    Dim sPersVistoCom As String
    Dim pnMovNro As Long
    Dim loVistoElectronico As SICMACT.frmVistoElectronico
    Set loVistoElectronico = New SICMACT.frmVistoElectronico
       
    pbTran = False

    If lnVolanteTasacion = 0 Then
        MsgBox "Debe imprimir el volante de tasación", vbCritical, "Aviso"
        Exit Sub
    End If
    
    If ValidaDatosGrabar = False Then Exit Sub
    
        'Asigno los valores a los parametros
        If lsPersCod = "" Then
            MsgBox "Crédito no tiene Titular verifíque la información", vbCritical, "Aviso"
            Exit Sub
        End If
            
        lnMontoPrestamo = CCur(lblMontoPrestamo.Caption)
        lnNetoRecibir = CCur(Me.lblMontoDesemb.Caption)
        lnPlazo = val(cboPlazo.Text)
        lsFechaVenc = Format$(Me.lblFecVen, "mm/dd/yyyy")
        lnOroBruto = val(lblOroBruto.Caption)
        lnOroNeto = val(lblOroNetoNew.Caption)
        lnPiezas = val(txtPiezas.Text)
        lnValTasacion = CCur(Me.lblTasaNueva.Caption)
        lsTipoContrato = Switch(cboTipcta.ListIndex = 0, "0", cboTipcta.ListIndex = 1, "1", cboTipcta.ListIndex = 2, "2")
        lsLote = ""
        lnIntAdelantado = CCur(Me.lblInteres.Caption)
        
        Set lrJoyas = FEJoyas.GetRsNew
    
        'Validar ingreso de Joyas
    
    If ValidarMsh = True Then Exit Sub
    
    
            '*** el codigo de operacio falta definir para reg de contrato pig por miestras se puso 150100
            lbResultadoVisto = loVistoElectronico.Inicio(1, gColRegistraContratoPig, lsPersCod)
            If Not lbResultadoVisto Then
                Exit Sub
            End If
'ALPA 20150617************************************************************
    Dim oCredPers As COMDCredito.DCOMCredito
    Set oCredPers = New COMDCredito.DCOMCredito
    Dim oRsVal As ADODB.Recordset
    Set oRsVal = New ADODB.Recordset
    Set oRsVal = oCredPers.RecValidaProcentajeCredito(gdFecSis, gsCodAge, CDbl(lblMontoDesemb.Caption), sLineaTmp)
    If Not (oRsVal.BOF Or oRsVal.EOF) Then
        If oRsVal!nSaldoAdeudado > 0 Then
        If Round((oRsVal!nSaldoCredito / oRsVal!nSaldoAdeudado) * 100, 2) > Round(oRsVal!nPorMax, 2) Then
            MsgBox "No existe saldo para esta aprobación, consultar con el Area de Creditos", vbInformation, "Aviso"
            Exit Sub
        End If
        Else
            MsgBox "No existe saldo para esta aprobación, consultar con el Area de Creditos", vbInformation, "Aviso"
            Exit Sub
        End If
    Else
            MsgBox "No existe saldo para esta aprobación, consultar con el Area de Creditos", vbInformation, "Aviso"
            Exit Sub
    End If
    Set oRsVal = Nothing
    Set oCredPers = Nothing
'**************************************************************************
    If MsgBox("¿Grabar Contrato Prestamo Pignoraticio ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        
        'Genera Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        
        Set loRegPig = New COMNColoCPig.NCOMColPContrato
            If loRegPig.ObtieneEstadoCredPignoAmp(Me.AXCodCta.NroCuenta) = 2100 Then
                MsgBox "El crédito se encuentra pendiente de desembolso", vbInformation, "Aviso"
                Call cmdCancelar_Click
                Exit Sub
            End If
            If fbMalCalificacion = False Then
                lsContrato = loRegPig.nRegistraContratoPignoraticioDetalle(gsCodCMAC & gsCodAge, gMonedaNacional, _
                lrPersonas, fnTasaInteresAdelantado, lnMontoPrestamo, lsFechaHoraGrab, lnPlazo, _
                lsFechaVenc, lnOroBruto, lnOroNeto, lnValTasacion, lnPiezas, lsTipoContrato, _
                lsLote, ln14k, ln16k, ln18k, ln21k, lsMovNro, lnIntAdelantado, lnCostoTasac, _
                lnCostoCustodia, lnImpuesto, lrJoyas, pnMovNro, , , , Me.AXCodCta.NroCuenta, val(Me.lblMontoDeudaAct.Caption), gsCodUser, , , sLineaTmp, lnTasaGracia, lnTasaMorato)
                
            End If
    
            pbTran = False
        Set loRegPig = Nothing
        
        objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, , lsContrato, gCodigoCuenta
        
        MsgBox "Se ha generado Contrato Nro " & lsContrato, vbInformation, "Aviso"
    
    
            If MsgBox("Imprimir Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                Set loRegImp = New COMNColoCPig.NCOMColPImpre
                    lsCadImprimir = ""
                    pnITF = gITF.fgITFCalculaImpuesto(val(lblMontoDesemb.Caption))
                    lsCadImprimir = loRegImp.nPrintContratoPignoraticioDet(lsContrato, True, lrPersonas, fnTasaInteresAdelantado, _
                        lnMontoPrestamo, lsFechaHoraGrab, Format(lsFechaVenc, "dd/MM/yyyy"), lnPlazo, lnOroBruto, lnOroNeto, lnValTasacion, _
                        lnPiezas, lsLote, ln14k, ln16k, ln18k, ln21k, lnIntAdelantado, lnCostoTasac, lnCostoCustodia, lnImpuesto, gsCodUser, , lsmensaje, gImpresora, pnITF)
                    If Trim(lsmensaje) <> "" Then
                        MsgBox lsmensaje, vbInformation, "Aviso"
                        Exit Sub
                    End If
                Set loPrevio = New previo.clsprevio
                
                Dim oImp As New ContsImp.clsConstImp
                    oImp.Inicia gImpresora
                    gPrnSaltoLinea = oImp.gPrnSaltoLinea
                    gPrnSaltoPagina = oImp.gPrnSaltoPagina
                Set oImp = Nothing
                    loPrevio.PrintSpool sLpt, lsCadImprimir, False
                    Do While True
                    Dim cad As String
                        If MsgBox("Reimprimir Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                                    loPrevio.PrintSpool sLpt, gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & lsCadImprimir, False
                                    loPrevio.PrintSpool sLpt, gPrnSaltoLinea '& gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoPagina & lsCadImprimir, False
                        Else
                            Exit Do
                        End If
                    Loop
        
                lrPersonas.MoveFirst
                MsgBox "Coloque Papel para Imprimir Hoja de Resumen...", vbInformation, "Aviso"
                lsCadImprimir = ""
            
                Call loRegImp.nPrintHojaResumen(lsContrato, lsPersNombre, lsPersApellido, lblClienteDOI.Caption, lsPersDireccDomicilio, _
                                lnValTasacion, lnMontoPrestamo, lnNetoRecibir, pnITF, gdFecSis, Format(lsFechaVenc, "dd/MM/yyyy"), lnIntAdelantado, gImpresora, lsmensaje, HojaResumen1, HojaResumen2)
            
                loPrevio.PrintSpool sLpt, HojaResumen1, False
                    MsgBox "Presione Enter para imprimir la segunda hoja", vbInformation, "Aviso"
                loPrevio.PrintSpool sLpt, HojaResumen2, False
                
                If MsgBox("Reimprimir Hoja de Resumen ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    loPrevio.PrintSpool sLpt, HojaResumen1, False
                    MsgBox "Presione Enter para imprimir la segunda hoja", vbInformation, "Aviso"
                    loPrevio.PrintSpool sLpt, HojaResumen2, False
                End If
                
                HojaResumen1 = ""
                HojaResumen2 = ""
                
                Set loRegImp = Nothing
            End If
        
        loVistoElectronico.RegistraVistoElectronico (pnMovNro)
        
        fbMalCalificacion = False
        lblCalificacionNormal = ""
        lblCalificacionPotencial = ""
        lblCalificacionDeficiente = ""
        lblCalificacionDudoso = ""
        lblCalificacionPerdida = ""
        lblPorcentajeTasa = ""
        fbClienteCPP = False
    End If

    
 
 Set loPrevio = Nothing
 Set loRegPig = Nothing
    
Call cmdCancelar_Click


Exit Sub

ControlError:   ' Rutina de control de errores.
    'Verificar que se halla iniciado transaccion y la cierra
    'If pbTran Then dbCmact.RollbackTrans
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
    Limpiar
End Sub

Private Sub cmdImpVolTas_Click()
    Dim lrJoyas  As New ADODB.Recordset
    Dim lnPlazo As Integer
    Dim lnOroBruto As Double
    Dim lnOroNeto As Double
    Dim lnPiezas As Integer
    Dim lnValTasacion As Currency
    Dim lsCadImprimir As String
    Dim loRegImp As COMNColoCPig.NCOMColPImpre
    Dim loPrevio As previo.clsprevio
    
        If ValidarMsh = True Then Exit Sub
        Set lrJoyas = FEJoyas.GetRsNew
        lnPlazo = val(cboPlazo.Text)
        lnOroBruto = val(lblOroBruto.Caption)
        lnOroNeto = val(lblOroNetoNew.Caption)
        lnPiezas = val(txtPiezas.Text)
        lnValTasacion = CCur(Me.lblTasaNueva.Caption)
        If MsgBox("Imprimir Volante de Tasación ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            Set loRegImp = New COMNColoCPig.NCOMColPImpre
                lsCadImprimir = ""
                lsCadImprimir = loRegImp.nPrintVolanteTasacion(lnPlazo, lnOroBruto, lnOroNeto, lnValTasacion, lnPiezas, lrJoyas, gImpresora)
            Set loRegImp = Nothing
            Set loPrevio = New previo.clsprevio
            
            Dim oImp As New ContsImp.clsConstImp
                oImp.Inicia gImpresora
                gPrnSaltoLinea = oImp.gPrnSaltoLinea
                gPrnSaltoPagina = oImp.gPrnSaltoPagina
            Set oImp = Nothing
                loPrevio.PrintSpool sLpt, lsCadImprimir & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea, False
            lnVolanteTasacion = 1
        End If
        Set loPrevio = Nothing
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
    If val(Me.lblMontoDesemb) < 0 Then
        MsgBox "El monto de prestamo es menor a la deuda anterior no se puede ampliar crédito.", vbCritical, "Aviso"
        Call cmdCancelar_Click
    End If
End Sub

Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim loCalculos As COMNColoCPig.NCOMColPCalculos 'NColPCalculos
Dim lrValida As ADODB.Recordset
Dim lbOk As Boolean
Dim lbCan As Boolean
Dim lsmensaje As String
Dim loParam As COMDColocPig.DCOMColPCalculos

Dim loCredPContrato As New COMNColoCPig.NCOMColPContrato 'RECO20120823 ERS074-2014

Set loParam = New COMDColocPig.DCOMColPCalculos

'On Error GoTo ControlError

    'Valida Contrato si esta Cancelado
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        Set lrValida = loValContrato.nValidaCancelacionCredPignoraticio(psNroContrato, gdFecSis, 0, lsmensaje)
    Set loValContrato = Nothing
    
    
    vValorTasacion = Format(lrValida!nTasacion, "#0.00")
    vTasaInteresVencido = lrValida!nTasaIntVenc
    vTasaInteresMoratorio = lrValida!nTasaIntMora
    vEstado = lrValida!nPrdEstado
    vSaldoCapital = Format(lrValida!nSaldo, "#0.00")
    vFecVencimiento = Format(lrValida!dVenc, "dd/mm/yyyy")
    vdiasAtraso = DateDiff("d", Format(vFecVencimiento, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
    vTasaInteres = lrValida!nTasaInteres
    vFecEstado = lrValida!dPrdEstado
    fnVarEstUltProcRem = lrValida!nEstUltProcRem
    gcCredAntiguo = lrValida!cCredB
    gnNotifiAdju = lrValida!nCodNotifiAdj 'PEAC 20080515
    gnNotifiCob = lrValida!nCodNotifiCob 'PEAC 20080715
    
    vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
    
    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
    vInteresMoratorio = loCalculos.nCalculaInteresMoratorio(vSaldoCapital, vTasaInteresMoratorio, vdiasAtraso)
    vInteresAdel = loCalculos.nCalculaInteresAlVencimiento(vSaldoCapital, vTasaInteres, vDiasAdel)
    vInteresAdel = Round(vInteresAdel, 2)
    vInteresVencido = loCalculos.nCalculaInteresMoratorio(vSaldoCapital, vTasaInteresVencido, vdiasAtraso)
    vInteresVencido = Round(vInteresVencido, 2)
    vCostoCustodiaMoratorio = loCalculos.nCalculaCostoCustodiaMoratorio(vValorTasacion, fnTasaCustodiaVencida, vdiasAtraso)
    vCostoCustodiaMoratorio = Round(vCostoCustodiaMoratorio, 2)
    vImpuesto = (vInteresVencido + vCostoCustodiaMoratorio + vInteresMoratorio) * fnTasaImpuesto
    vImpuesto = Round(vImpuesto, 2)
    fnVarGastoCorrespondencia = loCalculos.nCalculaGastosCorrespondencia(AXCodCta.NroCuenta, lsmensaje)
    
    '**************
    If vdiasAtraso <= 0 Then
        'PEAC 20070813
        If gcCredAntiguo = "A" Then
            vInteresAdel = Round(0, 2)
        Else
        Set loCalculos = New COMNColoCPig.NCOMColPCalculos
            vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
            '*** PEAC 20080806 ***************************
            'vInteresAdel = loCalculos.nCalculaInteresAdelantado(vSaldoCapital, vTasaInteres, vDiasAdel)
             vInteresAdel = loCalculos.nCalculaInteresAlVencimiento(vSaldoCapital, vTasaInteres, vDiasAdel)
            '*** FIN PEAC ********************************
            vInteresAdel = Round(vInteresAdel, 2)
        Set loCalculos = Nothing
        End If
    
    
    
    'end peac
    vdiasAtraso = 0
    vInteresVencido = 0
    vInteresMoratorio = 0
    vCostoCustodiaMoratorio = 0
    vImpuesto = 0
    fnVarGastoCorrespondencia = 0
Else
    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
        'PEAC 20070813
        
        If gcCredAntiguo = "A" Then
            vInteresAdel = Round(vInteresAdel, 2)
        Else
            vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(vFecVencimiento, "dd/mm/yyyy"))
            '*** PEAC 20080806 *********************************
            'vInteresAdel = loCalculos.nCalculaInteresAdelantado(vSaldoCapital, vTasaInteres, vDiasAdel)
             vInteresAdel = loCalculos.nCalculaInteresAlVencimiento(vSaldoCapital, vTasaInteres, vDiasAdel)
            '*** FIN PEAC **************************************
            vInteresAdel = Round(vInteresAdel, 2)
        End If
    
        vInteresVencido = loCalculos.nCalculaInteresMoratorio(vSaldoCapital, vTasaInteresVencido, vdiasAtraso)
        vInteresVencido = Round(vInteresVencido, 2)
        
        vInteresMoratorio = loCalculos.nCalculaInteresMoratorio(vSaldoCapital, vTasaInteresMoratorio, vdiasAtraso)
        vInteresMoratorio = Round(vInteresMoratorio, 2)
        
        vCostoCustodiaMoratorio = loCalculos.nCalculaCostoCustodiaMoratorio(vValorTasacion, fnTasaCustodiaVencida, vdiasAtraso)
        vCostoCustodiaMoratorio = Round(vCostoCustodiaMoratorio, 2)
        
        vImpuesto = (vInteresVencido + vCostoCustodiaMoratorio + vInteresMoratorio) * fnTasaImpuesto
        vImpuesto = Round(vImpuesto, 2)
        fnVarGastoCorrespondencia = loCalculos.nCalculaGastosCorrespondencia(AXCodCta.NroCuenta, lsmensaje)
        
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
    Set loCalculos = Nothing
End If
    '**************
    If vEstado = gColPEstPRema And fnVarEstUltProcRem = 2 Then  ' Si esta en via de Remate
        vCostoPreparacionRemate = fnTasaPreparacionRemate * vValorTasacion
        vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
    End If
    
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
    
If gnNotifiAdju = 1 Then
    If gnNotifiCob = 1 Then
        fnVarCostoNotificacion = 0
    End If
Else
    fnVarCostoNotificacion = 0
End If
    vDeuda = vSaldoCapital + vInteresAdel + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate + fnVarGastoCorrespondencia + vInteresMoratorio + fnVarCostoNotificacion
    
    'Muestra Datos
    lbOk = MuestraCredPig(psNroContrato)
    If lbOk = False Then
        AXCodCta.SetFocusCuenta
        Exit Sub
    End If
    Me.cmdGrabar.Enabled = True
    Set lrValida = Nothing
'RECO20120823 ERS074-2014************************************* ESTO NO SALE FALTA TERMINAR
If loCredPContrato.ObtieneHistorialCredRetasacion(AXCodCta.NroCuenta) = True Then
    lblCredRetasado.Visible = True
    cmdVerRetasacion.Visible = True
End If
'RECO FIN ****************************************************
Call CargarDatosProductoCrediticio
Call MostrarLineas
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Function MuestraCredPig(ByVal psNroContrato As String) As Boolean
Dim lrCredPig As ADODB.Recordset
Dim lrCredPigCostos As ADODB.Recordset
Dim lrCredPigPersonas As ADODB.Recordset
Dim lrCredPigJoyasDet As ADODB.Recordset
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnJoyasDet As Integer
Dim contador As Long
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim lnValTasNueva As Double
Dim lrCredPigKilotaje As ADODB.Recordset



    MuestraCredPig = True
    Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
    
        Set lrCredPig = loMuestraContrato.dObtieneDatosCreditoPignoraticio(psNroContrato)
        Set lrCredPigCostos = loMuestraContrato.dObtieneDatosCreditoPignoraticioCostos(psNroContrato)
        Set lrCredPigPersonas = loMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psNroContrato)
        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(psNroContrato, True)
        Set lrCredPigKilotaje = loMuestraContrato.ObtieneValorKilotajePesoCredPigno(psNroContrato)
    Set loMuestraContrato = Nothing
        
    If lrCredPig.BOF And lrCredPig.EOF Then
        lrCredPig.Close
        Set lrCredPig = Nothing
        Set lrCredPigPersonas = Nothing
        MsgBox " No se encuentra el Credito Pignoraticio " & psNroContrato, vbInformation, " Aviso "
        MuestraCredPig = False
        Exit Function
    Else
        If lrCredPig!nPrdEstado = 2101 Or lrCredPig!nPrdEstado = 2104 Or lrCredPig!nPrdEstado = 2106 Or lrCredPig!nPrdEstado = 2107 Then
          
            Me.lblOroBruto.Caption = lrCredPig!nOroBruto
            Me.lblOroNeto.Caption = lrCredPig!nOroNeto
            
            CargaDatosComboTipo Me.cboTipcta, Trim(lrCredPig!cTipCta)
            
            Me.txtPiezas.Text = lrCredPig!nPiezas
            Me.lblValorTasacion.Caption = lrCredPig!nTasacion
            Me.lblMontoPresAnt.Caption = lrCredPig!nMontoCol + vInteresMoratorio
            
            CargaDatosComboPlazo Me.cboPlazo, Trim(lrCredPig!nPlazo)
            CargaDatosComboPlazo Me.cboPlazoNuevo, Trim(lrCredPig!nPlazo)
            
            'Me.lblInteres.Caption = lrCredPig!nTasaInteres
            lnSaldoActual = lrCredPig!nSaldo
            lrCredPig.Close
            Set lrCredPig = Nothing
             
            ' Mostrar Clientes
            'Set lrPersonas = lrCredPigPersonas
        
            If Not (lrCredPigKilotaje.EOF And lrCredPigKilotaje.BOF) Then
                Dim j As Integer
                 For j = 1 To lrCredPigKilotaje.RecordCount
                    If lrCredPigKilotaje!cKilataje = 14 Then
                        ln14k = lrCredPigKilotaje!nPesoOro
                    End If
                    If lrCredPigKilotaje!cKilataje = 16 Then
                        ln16k = lrCredPigKilotaje!nPesoOro
                    End If
                    If lrCredPigKilotaje!cKilataje = 18 Then
                        ln18k = lrCredPigKilotaje!nPesoOro
                    End If
                    If lrCredPigKilotaje!cKilataje = 21 Then
                        ln21k = lrCredPigKilotaje!nPesoOro
                    End If
                    lrCredPigKilotaje.MoveNext
                Next
            End If
            Set lrPersonas = lrCredPigPersonas
            If Not (lrCredPigPersonas.EOF And lrCredPigPersonas.BOF) Then
                
                lblClienteNombre.Caption = lrCredPigPersonas!cPersNombre & " " & lrCredPigPersonas!cpersapellido
                lblClienteDOI.Caption = lrCredPigPersonas!NroDNI
                lsPersCod = lrCredPigPersonas!cPersCod
                lsPersApellido = lrCredPigPersonas!cpersapellido
                lsPersNombre = lrCredPigPersonas!cPersNombre
                lsPersDireccDomicilio = lrCredPigPersonas!cPersDireccDomicilio
            Else
                MuestraCredPig = False
                Exit Function
            End If
        
            ObtenerTipoCiente (lsPersCod)
            If (ObtenerCalificacionRCC(AXCodCta.NroCuenta) = False) Then
                Exit Function
            End If
        
            'lrCredPigPersonas.Close
            'Set lrCredPigPersonas = Nothing
            
            'Para mostrar el Detalle de las joyas
            Set loConstSis = New COMDConstSistema.NCOMConstSistema
                lnJoyasDet = loConstSis.LeeConstSistema(109)
            Set loConstSis = Nothing
            If lnJoyasDet = 1 Then
                If Not (lrCredPigJoyasDet.EOF And lrCredPigJoyasDet.BOF) Then
                    Dim loColPCalculos As COMDColocPig.DCOMColPCalculos
                    Dim lnPOro As Double
                    Dim loColContrato As COMDColocPig.DCOMColPContrato
                    Dim lnValorPOro As Double
                    Dim lnMatOro As Integer
                    Dim loDR As ADODB.Recordset
                    Dim i As Integer
                    
                    Set loColContrato = New COMDColocPig.DCOMColPContrato
                    Set loDR = New ADODB.Recordset
                    
                    FEJoyas.Clear
                    FEJoyas.FormaCabecera
                    For i = 1 To lrCredPigJoyasDet.RecordCount
                        FEJoyas.AdicionaFila
                        FEJoyas.TextMatrix(i, 1) = lrCredPigJoyasDet!nPiezas
                        FEJoyas.TextMatrix(i, 2) = lrCredPigJoyasDet!cKilataje
                        FEJoyas.TextMatrix(i, 3) = lrCredPigJoyasDet!nPesoBruto
                        FEJoyas.TextMatrix(i, 4) = lrCredPigJoyasDet!nPesoNeto
                        FEJoyas.TextMatrix(i, 5) = lrCredPigJoyasDet!nValTasac
                        
                        lnPesoNetoDesc = CDbl(lrCredPigJoyasDet!nPesoBruto) * 0.1
                        lnPesoNetoDesc = CDbl(lrCredPigJoyasDet!nPesoBruto) - lnPesoNetoDesc
            
                        Set loColPCalculos = New COMDColocPig.DCOMColPCalculos
                                             
                        lnPOro = loColPCalculos.dObtienePrecioMaterial(1, val(Left(FEJoyas.TextMatrix(i, 2), 2)), 1) 'APRI 20170408  CAMBIO de Right (X,3) -> Left (X,2)
                                           
                        lnMatOro = Left(FEJoyas.TextMatrix(i, 2), 2) 'APRI 20170408  CAMBIO de Right (X,3) -> Left (X,2)
         
                            Set loDR = loColContrato.PigObtenerValorTasacionxTpoClienteKt(nTpoCliente)
                            If Not (loDR.BOF And loDR.EOF) Then
                                If lnMatOro = 14 Then
                                    lnValorPOro = loDR!n14kt
                                ElseIf lnMatOro = 16 Then
                                    lnValorPOro = loDR!n16kt
                                ElseIf lnMatOro = 18 Then
                                    lnValorPOro = loDR!n18kt
                                ElseIf lnMatOro = 21 Then
                                    lnValorPOro = loDR!n21kt
                                End If
                            End If
                            
                            If lnPOro <= 0 Then
                                MsgBox "Precio del Material No ha sido ingresado en el Tarifario, actualice el Tarifario", vbInformation, "Aviso"
                                Exit Function
                            End If
                            Set loColPCalculos = Nothing
                            'FEJoyas.TextMatrix(i, 9) = lnPesoNetoDesc
                        If (FEJoyas.TextMatrix(FEJoyas.row, 4) <= lnPesoNetoDesc) Then
                            FEJoyas.TextMatrix(i, 6) = Format$(val(FEJoyas.TextMatrix(FEJoyas.row, 4) * lnValorPOro), "#####.00")
                            lnPesoNetoDesc = val(FEJoyas.TextMatrix(FEJoyas.row, 4))
                        Else
                            FEJoyas.TextMatrix(i, 6) = Format$(val(lnPesoNetoDesc * lnValorPOro), "#####.00")
                        End If
                        FEJoyas.TextMatrix(i, 9) = lnPesoNetoDesc
                        FEJoyas.TextMatrix(i, 7) = lrCredPigJoyasDet!cDescrip
                        
                        lnValTasNueva = lnValTasNueva + CDbl(FEJoyas.TextMatrix(i, 6))
                        SumaColumnas
                        lrCredPigJoyasDet.MoveNext
                    Next
                    
                    Me.lblTasaNueva.Caption = lnValTasNueva
                    DesbloquearControles
                    CalculaPrestamo
                    
                End If
            End If
        Else
         MsgBox "El número de crédito no está sujeto a este tipo de operación", vbCritical, "Aviso"
         Call cmdCancelar_Click
        End If
    End If
Exit Function

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Function


Private Sub CargaDatosComboTipo(ByVal pcCombo As ComboBox, ByVal psCod As String)
    Dim i As Integer
    Dim sCad As String
    For i = 0 To pcCombo.ListCount
        sCad = Left(Trim(pcCombo.List(i)), (Len(Trim(pcCombo.List(i))) - 5))
        If UCase(Trim(sCad)) = Trim(psCod) Then
            pcCombo.ListIndex = i
            Exit Sub
        End If
        i = i + 1
    Next
End Sub

Private Sub CargaDatosComboPlazo(ByVal pcCombo As ComboBox, ByVal psCod As String)
    Dim i As Integer
    Dim sCad As String
    For i = 0 To pcCombo.ListCount
        If Trim(pcCombo.List(i)) = Trim(psCod) Then
            pcCombo.ListIndex = i
            Exit Sub
        End If
        i = i + 1
    Next
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    AXCodCta.Age = gsCodAge
    CargaParametros
    Limpiar
    fsColocLineaCredPig = "0101117550101"
    Set objPista = New COMManejador.Pista
    gsOpeCod = gPigRegistrarContrato
    Call CargaPlazo 'RECO20140421
End Sub

Public Sub ObtenerTipoCiente(ByVal psPersCod As String)
    lnPesoNetoDesc = 0
    Dim loPigContrato As COMDColocPig.DCOMColPContrato
    Set loPigContrato = New COMDColocPig.DCOMColPContrato
    Dim poDR As ADODB.Recordset
    Set poDR = New ADODB.Recordset
    Set poDR = loPigContrato.dVerificarCredPignoAdjudicado(psPersCod)
    If Not (poDR.BOF And poDR.EOF) Then
        nTpoCliente = 1
    Else
        Set poDR = Nothing
        Set poDR = loPigContrato.dVerificarCredPignoDesembolso(psPersCod)
        If Not (poDR.BOF And poDR.EOF) Then
            nTpoCliente = 2
        Else
            nTpoCliente = 1
        End If
    End If
End Sub

Public Function ObtenerCalificacionRCC(ByVal psCtaCod As String) As Boolean
    Dim rsCalificacionSBS As ADODB.Recordset
    Dim loPersContrato As COMDColocPig.DCOMColPContrato
    Set rsCalificacionSBS = New ADODB.Recordset
    Set loPersContrato = New COMDColocPig.DCOMColPContrato
    Dim lrPersContrato As ADODB.Recordset
    Dim oDCOMCreditos As DCOMCreditos
    Set oDCOMCreditos = New DCOMCreditos
    
    Set lrPersContrato = loPersContrato.ObtieneDatosPersonaXCredito(psCtaCod)
    ObtenerCalificacionRCC = True
    If Not (lrPersContrato.BOF And lrPersContrato.EOF) Then
        Set rsCalificacionSBS = oDCOMCreditos.DatosPosicionClienteCalificacionSBS(IIf(lrPersContrato!nPersPersoneria = 1, True, False), _
                                                                                  IIf(lrPersContrato!nPersPersoneria = 1, Trim(lrPersContrato!NroDNI), Trim(lrPersContrato!NroRuc)), "")
        If Not rsCalificacionSBS.BOF And Not rsCalificacionSBS.EOF Then
            lblTituloCalificacion = "Última Calificación Según SBS - RCC " & Format(rsCalificacionSBS!Fec_Rep, "dd/mm/yyyy")
            fbMalCalificacion = False
            fbClienteCPP = False
            
            lblCalificacionNormal = "Normal " & rsCalificacionSBS!nNormal & " %"
            lblCalificacionPotencial = "Potencial " & rsCalificacionSBS!nPotencial & " %"
            lblCalificacionDeficiente = "Deficiente " & rsCalificacionSBS!nDeficiente & " %"
            lblCalificacionDudoso = "Dudoso " & rsCalificacionSBS!nDudoso & " %"
            lblCalificacionPerdida = "Perdida " & rsCalificacionSBS!nPerdido & " %"
            
            'RECO20140527 RFC1405270001***********************************************************
            If CDbl(rsCalificacionSBS!nDudoso) <> 0 Or CDbl(rsCalificacionSBS!nPerdido) <> 0 Or CDbl(rsCalificacionSBS!nDeficiente) Then
            'If CDbl(rsCalificacionSBS!nDudoso) <> 0 Or CDbl(rsCalificacionSBS!nPerdido) <> 0 Then
            'RECO20140527 FIN*********************************************************************
                    MsgBox "Clientes con calificación Deficiente, Dudoso y Perdida no pueden ser atendidos ", vbCritical, "Aviso"
                    ObtenerCalificacionRCC = False
                    Call cmdCancelar_Click
                    Exit Function
            End If
            
            If CDbl(rsCalificacionSBS!nPotencial) <> 0 Or CDbl(rsCalificacionSBS!nDeficiente) <> 0 Then
                fbClienteCPP = True = True
            End If
            
        Else
            lblTituloCalificacion = "Última Calificación Según SBS - RCC "
            fbMalCalificacion = False
            lblCalificacionNormal = "No Registrado"
            lblCalificacionPotencial = "No Registrado"
            lblCalificacionDeficiente = "No Registrado"
            lblCalificacionDudoso = "No Registrado"
            lblCalificacionPerdida = "No Registrado"
            fbClienteCPP = False
    
        End If
    End If

End Function

Private Sub cmdCancelar_Click()
    Limpiar
    txtPiezas.Enabled = False
    cboPlazo.Enabled = False
    lblValorTasacion.Enabled = False
    cboPlazo.ListIndex = 0
    cboTipcta.Enabled = False
    Me.AXCodCta.Cuenta = ""
    Me.cmdBuscar.Enabled = True
    fbMalCalificacion = False
    lblCalificacionNormal = ""
    lblCalificacionPotencial = ""
    lblCalificacionDeficiente = ""
    lblCalificacionDudoso = ""
    lblCalificacionPerdida = ""
    lblPorcentajeTasa = ""
    Me.lblClienteNombre.Caption = ""
    Me.lblClienteDOI.Caption = ""
    Me.lblMontoBruto.Caption = Format(0, "#0.00")
    Me.lblMontoDesemb.Caption = Format(0, "#0.00")
    Me.lblTasaNueva.Caption = Format(0, "#0.00")
    Me.lblMontoDeudaAct.Caption = Format(0, "#0.00")
    Me.lblFecVen.Caption = ""
    fbClienteCPP = False
    Me.lblClienteNombre.Caption = ""
    Me.lblClienteDOI.Caption = ""
    Me.lblFecVen.Caption = ""
    Me.lblTasaNueva.Caption = Format(0, "#0.00")
    Me.lblMontoBruto.Caption = Format(0, "#0.00")
    Me.lblMontoDeudaAct.Caption = Format(0, "#0.00")
    Me.lblMontoDesemb.Caption = Format(0, "#0.00")
    Me.lblInteres.Caption = Format(0, "#0.00")
    lnVolanteTasacion = 0
    lblCredRetasado.Visible = False  'RECO20120823 ERS074-2014
    cmdVerRetasacion.Visible = False 'RECO20120823 ERS074-2014
    txtBuscarLinea.Text = ""
    lblLineaDesc.Caption = ""
End Sub
Private Sub SumaColumnas()
    Dim i As Integer
    Dim lnPiezasT As Integer, lnPBrutoT As Double, lnPNetoT As Double, lnTasacT As Double, lnPNetoTNew As Double
    Dim lbNroPigAdlCli As Boolean
    Dim lsPersCod As String
    lsPersCod = lsPersCod

    lnPiezasT = 0: lnPBrutoT = 0:       lnPNetoT = 0:       lnTasacT = 0 ':         lnPrestamoT = 0
    'Total Piezas
    lnPiezasT = FEJoyas.SumaRow(1)
    txtPiezas.Text = Format$(lnPiezasT, "##")

    'PESO BRUTO
    lnPBrutoT = FEJoyas.SumaRow(3)
    lblOroBruto.Caption = Format$(lnPBrutoT, "######.00")

    'PESO NETO
    lnPNetoT = FEJoyas.SumaRow(4)
    
    'Tasacion
    lnTasacT = FEJoyas.SumaRow(5)
    'PESO NETO NUEVO
    lnPNetoTNew = FEJoyas.SumaRow(9)
    
    lbNroPigAdlCli = devolverPignoraticiosAdjudicadosCliente(lsPersCod, gdFecSis)
    If lbNroPigAdlCli = False Then
        lblOroNeto.Caption = Format$(lnPNetoT, "######.00")
        lblOroNeto.ForeColor = &H80000008
        lblOroNetoNew.Caption = Format$(lnPNetoTNew, "######.00")
        lblOroNetoNew.ForeColor = &H80000008
    Else
        lblOroNeto.Caption = Format$(lnPNetoT - (lnPNetoT * 0.15), "######.00")
        lblOroNeto.ForeColor = &HFF&
        lnTasacT = lnTasacT - lnTasacT * 0.15
        lblOroNetoNew.Caption = Format$(lnPNetoTNew - (lnPNetoTNew * 0.15), "######.00")
        lblOroNetoNew.ForeColor = &HFF&
        
    End If
  
        lblValorTasacion.Caption = Format$(lnTasacT, "######.00")

    cboPlazo.ListIndex = 0
    vPrestamo = val(lblValorTasacion.Caption) * fnPorcentajePrestamo
    Me.lblMontoPrestamo.Caption = Format(vPrestamo, "#0.00")
    Me.lblMontoBruto.Caption = Format(CDbl(Me.lblMontoPrestamo.Caption), "#0.00")
    'Me.lblMontoDeudaAct.Caption = Format(lnSaldoActual, "#0.00")
    Me.lblMontoDeudaAct.Caption = Format(vDeuda, "#0.00")
    Me.lblMontoDesemb.Caption = Format(CDbl(Me.lblMontoBruto.Caption) - CDbl(Me.lblMontoDeudaAct.Caption), "#0.00")
    If CDbl(Me.lblMontoPrestamo.Caption) > 0 Then
        cmdGrabar.Enabled = True
        cmdImpVolTas.Enabled = True
    End If
    CalculaCostosAsociados
End Sub
Private Function devolverPignoraticiosAdjudicadosCliente(ByVal pcCodcli As String, ByVal pdFecSis As Date) As Boolean
    Dim oDCOMCredDoc As COMDCredito.DCOMCredDoc
    Dim rsPigAdjCli As ADODB.Recordset
    Set oDCOMCredDoc = New COMDCredito.DCOMCredDoc
    
    Set rsPigAdjCli = oDCOMCredDoc.devolverPignoraticiosAdjudicadosCliente(pcCodcli, pdFecSis)
    
    If Not rsPigAdjCli.BOF And Not rsPigAdjCli.EOF Then
        If rsPigAdjCli!nNroCtaAdjudicados > 3 Then
            devolverPignoraticiosAdjudicadosCliente = True
        Else
            devolverPignoraticiosAdjudicadosCliente = False
        End If
    Else
        devolverPignoraticiosAdjudicadosCliente = False
    End If
    
    Set rsPigAdjCli = Nothing
    Set oDCOMCredDoc = Nothing
End Function


Private Sub CalculaCostosAsociados()
    Dim loCostos As COMNColoCPig.NCOMColPCalculos
    Dim loTasaInt As COMDColocPig.DCOMColPCalculos
    
    Set loCostos = New COMNColoCPig.NCOMColPCalculos
    
        vPrestamo = CDbl(Me.lblMontoPrestamo.Caption)
        Set loTasaInt = New COMDColocPig.DCOMColPCalculos
            If fbClienteCPP = False Then
                'ALPA 20150625************************************************************
                'fnTasaInteresAdelantado = loTasaInt.dObtieneTasaInteres(fsColocLineaCredPig, "1")
                fnTasaInteresAdelantado = lnTasaInicial
                lblPorcentajeTasa = CStr(fnTasaInteresAdelantado) & " %"
            Else
                'ALPA 20150625************************************************************
                Dim oClases As New clases.NConstSistemas
                'fnTasaInteresAdelantado = oClases.LeeConstSistema(453)
                fnTasaInteresAdelantado = lnTasaFinal
                lblPorcentajeTasa = CStr(fnTasaInteresAdelantado) & " %"
            End If
            
        Set loTasaInt = Nothing
        
        vCostoTasacion = loCostos.nCalculaCostoTasacion(val(lblValorTasacion.Caption), fnTasaTasacion)
        vCostoCustodia = loCostos.nCalculaCostoCustodia(val(lblValorTasacion.Caption), fnTasaCustodia, val(cboPlazo.Text))
        vInteres = loCostos.nCalculaInteresAlVencimiento(val(Me.lblMontoPrestamo.Caption), fnTasaInteresAdelantado, val(cboPlazo.Text))
        vImpuesto = loCostos.nCalculaImpuestoDesembolso(vCostoTasacion, 0, vCostoCustodia, fnTasaImpuesto)
        vNetoPagar = val(Me.lblMontoPrestamo.Caption) - vCostoTasacion - vCostoCustodia - vImpuesto
    
    Set loCostos = Nothing
    'Me.lblInteres = Format(vInteres, "#0.00")
    Me.lblFecVen = Format(DateAdd("d", val(Me.cboPlazoNuevo.Text), gdFecSis), "dd/mm/yyyy")
End Sub
Private Sub CargaParametros()
    Dim loParam As COMDColocPig.DCOMColPCalculos
    Dim loConstSis As COMDConstSistema.NCOMConstSistema
    Set loParam = New COMDColocPig.DCOMColPCalculos
    
        fnTasaInteresAdelantado = loParam.dObtieneTasaInteres(fsColocLineaCredPig, "1")
        fnTasaCustodia = loParam.dObtieneColocParametro(gConsColPTasaCustodia)
        fnTasaTasacion = loParam.dObtieneColocParametro(gConsColPTasaTasacion)
        fnTasaImpuesto = loParam.dObtieneColocParametro(gConsColPTasaImpuesto)
        fnTasaPreparacionRemate = loParam.dObtieneColocParametro(gConsColPTasaPreparaRemate)
        fnTasaCustodiaVencida = loParam.dObtieneColocParametro(gConsColPTasaCustodiaVencida)
        fnRangoPreferencial = loParam.dObtieneColocParametro(3019)
        fnPorcentajePrestamo = loParam.dObtieneColocParametro(gConsColPPorcentajePrestamo)
        fnImpresionesContrato = loParam.dObtieneColocParametro(gConsColPNroImpresionesContrato)
        fnMaxMontoPrestamo1 = loParam.dObtieneColocParametro(gConsColPLim1MontoPrestamo)
        
    Set loParam = Nothing
    Set loConstSis = New COMDConstSistema.NCOMConstSistema
        fnJoyasDet = loConstSis.LeeConstSistema(109)
        fsPlazoSist = loConstSis.LeeConstSistema(503) 'RECO20150421
    Set loConstSis = Nothing
End Sub

Private Sub Limpiar()
    vContAnte = False
    lblOroBruto.Caption = Format(0, "#0.00")
    lblOroNeto.Caption = Format(0, "#0.00")
    lblOroNetoNew.Caption = Format(0, "#0.00")
    txtPiezas.Text = Format(0, "#0")
    cboPlazo.ListIndex = 0
    lblValorTasacion.Caption = Format(0, "#0.00")
    Me.lblMontoPresAnt.Caption = Format(0, "#0.00")
    Me.lblMontoPrestamo.Caption = Format(0, "#0.00")
    Me.lblInteres = Format(0, "#0.00")
    lsPersCod = ""
    cboTipcta.ListIndex = 0
    FEJoyas.Clear
    FEJoyas.Rows = 2
    FEJoyas.FormaCabecera
    lnJoyas = 0
    Me.AXCodCta.Cuenta = ""
    Me.cmdGrabar.Enabled = False
    'cboPlazoNuevo.ListIndex = 0 'RECO20150421
    Me.cboPlazoNuevo.Enabled = False
    Me.cmdCancelar.Enabled = False
    Me.cmdImpVolTas.Enabled = False
End Sub

Public Sub DesbloquearControles()
    Me.cmdGrabar.Enabled = True
    Me.cboPlazoNuevo.Enabled = True
    Me.cmdCancelar.Enabled = True
    Me.cmdImpVolTas.Enabled = True
End Sub

Private Sub cboPlazoNuevo_Click()
    If val(lblMontoPrestamo.Caption) = 0 And lsPersCod > "" Then
        CalculaPrestamo
    End If
    If val(lblMontoPrestamo.Caption) <> 0 Then
        CalculaCostosAsociados
    End If
End Sub

Private Sub CalculaPrestamo()
    Dim loCostos As COMNColoCPig.NCOMColPCalculos
    Set loCostos = New COMNColoCPig.NCOMColPCalculos
    
    vPrestamo = val(lblTasaNueva.Caption) * fnPorcentajePrestamo
    Me.lblMontoPrestamo.Caption = Format(vPrestamo, "#0.00")
    Me.lblMontoBruto.Caption = Format(vPrestamo, "#0.00")
    vInteres = loCostos.nCalculaInteresAlVencimiento(val(Me.lblMontoPrestamo.Caption), fnTasaInteresAdelantado, val(cboPlazo.Text))
    Me.lblInteres = Format(vInteres, "#0.00")
    Me.lblMontoDesemb.Caption = Format(CDbl(Me.lblMontoBruto.Caption) - CDbl(Me.lblMontoDeudaAct.Caption), "#0.00")
End Sub

Private Sub cboplazo_KeyPress(KeyAscii As Integer)
    vPrestamo = val(lblTasaNueva.Caption) * fnPorcentajePrestamo
    lblMontoPrestamo.Caption = Format(vPrestamo, "#0.00")
End Sub

Private Function ValidaDatosGrabar() As Boolean
Dim lbOk As Boolean
lbOk = True
If lsPersCod = "" Then
    MsgBox "Falta ingresar el cliente" & vbCr & _
    " Cancele operación ", , " Aviso "
    lbOk = False
    Exit Function
End If

If val(lblOroBruto) < val(lblOroNetoNew) Then
    MsgBox " Oro Neto debe ser menor o igual a Oro Bruto ", vbInformation, " Aviso "
    lbOk = False
    Exit Function
End If
' Monto de Prestamo < 60% de Valor de Tasacion
If val(lblMontoPrestamo.Caption) > val(Format(fnPorcentajePrestamo * val(Me.lblTasaNueva.Caption), "#0.00")) Then
    MsgBox " Monto de Prestamo debe ser menor al " & fnPorcentajePrestamo & " % del Valor de Tasacion ", vbInformation, " Aviso "
    'txtMontoPrestamo.SetFocus
    lbOk = False
    Exit Function
End If
If Trim(lblMontoPrestamo.Caption) = "" Then
    MsgBox " Falta ingresar Monto de Prestamo " & vbCr & " No se puede grabar con datos inconclusos ", vbInformation, " Aviso "
    'txtMontoPrestamo.SetFocus
    lbOk = False
    Exit Function
End If
' llena las tipos de Kilatajes

    If FEJoyas.Rows < 1 Then
        MsgBox " No ha ingresado el Detalle de las Joyas" & vbCr & " No se puede grabar con datos inconclusos ", vbInformation, " Aviso "
        'cmdAgregar.SetFocus
        lbOk = False
        Exit Function
    Else
        FEJoyas.row = 0
'        txt14k.Text = 0: txt16k.Text = 0: txt18k.Text = 0: txt21k.Text = 0
'        Do While FEJoyas.row < FEJoyas.Rows - 1
'            Select Case val(Right(FEJoyas.TextMatrix(FEJoyas.row + 1, 2), 3))
'                Case 14
'                    ln14k = val(txt14k.Text) + val(FEJoyas.TextMatrix(FEJoyas.row + 1, 4))
'                Case 16
'                    ln16k = val(txt16k.Text) + val(FEJoyas.TextMatrix(FEJoyas.row + 1, 4))
'                Case 18
'                    ln18k = val(txt18k.Text) + val(FEJoyas.TextMatrix(FEJoyas.row + 1, 4))
'                Case 21
'                    ln21k = val(txt21k.Text) + val(FEJoyas.TextMatrix(FEJoyas.row + 1, 4))
'            End Select
'            FEJoyas.row = FEJoyas.row + 1
'        Loop
    End If

ValidaDatosGrabar = lbOk
End Function

Public Function ValidarMsh() As Boolean
    Dim nFilas As Integer
    Dim i As Integer
    nFilas = FEJoyas.Rows
    For i = 0 To nFilas - 1
        If FEJoyas.TextMatrix(i, 1) = "" Then
            ValidarMsh = True
            MsgBox "Ingrese el detalle de Joyas", vbInformation, "Aviso"
            Exit Function
        End If
    Next
End Function

'RECO20150421 *****************************************
Private Sub CargaPlazo()
    Dim i As Integer
    Dim sPlazo As String
    For i = 1 To Len(fsPlazoSist)
        If Mid(fsPlazoSist, i, 1) <> "," Then
            sPlazo = sPlazo & Mid(fsPlazoSist, i, 1)
        Else
            cboPlazoNuevo.AddItem (sPlazo)
            sPlazo = ""
        End If
    Next
End Sub
'END RECO**********************************************
'ALPA 20150616*****************************************
Private Sub CargarDatosProductoCrediticio()
Dim sCodigo As String
Dim sCtaCodOrigen As String
Dim oLineas As COMDCredito.DCOMLineaCredito
Set RLinea = New ADODB.Recordset
sCodigo = Mid(txtBuscarLinea.Text, 5, Len(txtBuscarLinea.Text))
sLineaTmp = sCodigo
Set oLineas = New COMDCredito.DCOMLineaCredito
txtBuscarLinea.Text = ""
lblLineaDesc.Caption = ""
Set RLinea = oLineas.RecuperaLineadeCreditoProductoCrediticio("705", "0", Trim(Right((txtBuscarLinea.psDescripcion), 15)), sLineaTmp, lblLineaDesc, "1", CCur(IIf((lblMontoDesemb.Caption) = "", 0, lblMontoDesemb.Caption)), 0)
Set oLineas = Nothing
       If RLinea.RecordCount > 0 Then
          If txtBuscarLinea.Text = "" Then
            txtBuscarLinea.Text = "XXX"
          End If
          Call CargaDatosLinea
          If txtBuscarLinea.Text = "XXX" Then
            txtBuscarLinea.Text = ""
          End If
       Else
            lnTasaInicial = 0
            lnTasaFinal = 0
       End If
Call MostrarLineas
End Sub
'Private Sub cmdLineas_Click()
'Dim oLineas As COMDCredito.DCOMLineaCredito
'Dim sCtaCod As String
'
'bBuscarLineas = True
'sCtaCod = ActxCta.NroCuenta
'Set oLineas = New COMDCredito.DCOMLineaCredito
'txtBuscarLinea.Text = ""
'lblLineaDesc.Caption = ""
'If ActxCta.Cuenta = "" Then
'    MsgBox "Ingrese Nº de Crédito", vbCritical, "Aviso"
'    Exit Sub
'End If
'txtBuscarLinea.rs = oLineas.RecuperaLineasProductoArbol("755", "1", , gsCodAge, 30, CDbl(lblMontoDesemb.Caption), 1, , 0, gdFecSis)
'Set oLineas = Nothing
'End Sub
Private Sub txtBuscarLinea_EmiteDatos()
Dim sCodigo As String
Dim oCred As COMNCredito.NCOMCredito
Dim RLinea As ADODB.Recordset
Dim bExisteLineaCred As Boolean
Dim sCtaCodOrigen As String
sCodigo = Mid(txtBuscarLinea.Text, 5, Len(txtBuscarLinea.Text))
If sCodigo <> "" Then
    sLineaTmp = sCodigo
    If txtBuscarLinea.psDescripcion <> "" Then lblLineaDesc = txtBuscarLinea.psDescripcion Else lblLineaDesc = ""
        Set oCred = New COMNCredito.NCOMCredito
        Set oCred = Nothing
        bExisteLineaCred = True
        If bExisteLineaCred Then
        Else
            MsgBox "No existen Líneas de Crédito con el Plazo seleccionado", vbInformation, "Aviso"
            txtBuscarLinea.Text = ""
            lblLineaDesc = ""
        End If
Else
    lblLineaDesc = ""
End If
End Sub
Private Sub MostrarLineas()
    txtBuscarLinea.Text = ""
    lblLineaDesc.Caption = ""
    If val(lblMontoDesemb.Caption) > 0 Then
    Dim oLineas As COMDCredito.DCOMLineaCredito
    Dim lrsLineas As ADODB.Recordset
    Set lrsLineas = New ADODB.Recordset
    
    Set oLineas = New COMDCredito.DCOMLineaCredito
    Set lrsLineas = oLineas.RecuperaLineasProductoArbol("755", "1", , gsCodAge, 30, CDbl(lblMontoDesemb.Caption), 1, , 0, gdFecSis)
    Set oLineas = Nothing
    txtBuscarLinea.rs = lrsLineas
    End If
End Sub
Private Sub CargaDatosLinea()
ReDim MatCalend(0, 0)
ReDim MatrizCal(0, 0)
    
    If Trim(txtBuscarLinea.Text) = "" Then
        Exit Sub
    End If
    If RLinea.BOF Or RLinea.EOF Then
        lnTasaInicial = 0#
        lnTasaFinal = 0#
        Exit Sub
    End If
    lnTasaInicial = RLinea!nTasaIni
    lnTasaFinal = RLinea!nTasafin

    If fbClienteCPP = False Then
        fnTasaInteresAdelantado = lnTasaInicial
        lnTasaCompes = fnTasaInteresAdelantado
        lnTasaGracia = IIf(IsNull(RLinea!nTasaGraciaIni), 0#, RLinea!nTasaGraciaIni)
        lnTasaMorato = IIf(IsNull(RLinea!nTasaMoraIni), 0#, RLinea!nTasaMoraIni)
        lblPorcentajeTasa = CStr(fnTasaInteresAdelantado) & " %"
    Else
        Dim oClases As New clases.NConstSistemas
        fnTasaInteresAdelantado = lnTasaFinal
        lnTasaCompes = fnTasaInteresAdelantado
        lnTasaGracia = IIf(IsNull(RLinea!nTasaGraciaFin), 0#, RLinea!nTasaGraciaFin)
        lnTasaMorato = IIf(IsNull(RLinea!nTasaMoraFin), 0#, RLinea!nTasaMoraFin)
        lblPorcentajeTasa = CStr(fnTasaInteresAdelantado) & " %"
    End If
    If RLinea!nTasaIni <> RLinea!nTasafin Then
        If fnTasaInteresAdelantado >= RLinea!nTasaIni And fnTasaInteresAdelantado <= RLinea!nTasafin Then
            lnTasaCompes = fnPorcentajePrestamo
        Else
            lnTasaCompes = 0#
        End If
    End If
    CalculaCostosAsociados
End Sub
'******************************************************
