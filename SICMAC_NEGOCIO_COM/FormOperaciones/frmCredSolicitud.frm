VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredSolicitud 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitud de Credito"
   ClientHeight    =   8490
   ClientLeft      =   1965
   ClientTop       =   2340
   ClientWidth     =   9315
   Icon            =   "frmCredSolicitud.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdpresolicitud 
      Caption         =   "Pre Solicitud"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   80
      Top             =   105
      Width           =   1335
   End
   Begin VB.CheckBox chkAutAmpliacion 
      Caption         =   "Solicitar Autorización de Ampliación Excepcional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   78
      Top             =   7350
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox txtVendedor 
      Height          =   285
      Left            =   5880
      TabIndex        =   71
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdEnvioEstCta 
      Caption         =   "Envio Estado Cta."
      Enabled         =   0   'False
      Height          =   360
      Left            =   7500
      TabIndex        =   70
      Top             =   2640
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ComboBox cmbCondicionOtra2 
      Height          =   315
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   67
      Top             =   560
      Width           =   1860
   End
   Begin VB.ComboBox cmbCondicionOtra 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   58
      Top             =   560
      Width           =   1860
   End
   Begin VB.CheckBox ChkCap 
      Caption         =   "Capit. de Interes"
      Height          =   150
      Left            =   7560
      TabIndex        =   52
      Top             =   2520
      Width           =   1500
   End
   Begin SICMACT.ActXCodCta ActXCtaCred 
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3585
      _ExtentX        =   6535
      _ExtentY        =   688
      Texto           =   "Credito :"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
      CMAC            =   "112"
   End
   Begin VB.Frame fraCreditos 
      Height          =   2805
      Left            =   30
      TabIndex        =   25
      Top             =   4920
      Width           =   9225
      Begin VB.CommandButton cmdCreditoVerde 
         Caption         =   "E.A"
         Height          =   255
         Left            =   5640
         TabIndex        =   87
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cmbSubDestCred 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   1845
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDestinoDetalleAguaS 
         Caption         =   "A.S"
         Height          =   255
         Left            =   5160
         TabIndex        =   81
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdDestinoDetalle 
         Caption         =   "&Detalle.."
         Height          =   285
         Left            =   5160
         TabIndex        =   79
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame fraPromotor 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   75
         Top             =   2280
         Width           =   8895
         Begin VB.ComboBox cmbPromotor 
            Height          =   315
            Left            =   960
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   120
            Width           =   3180
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Promotor  :"
            Height          =   195
            Left            =   120
            TabIndex        =   77
            Top             =   165
            Width           =   765
         End
      End
      Begin Spinner.uSpinner spnPlazo 
         Height          =   300
         Left            =   7800
         TabIndex        =   56
         Top             =   1080
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   529
         Max             =   750
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin Spinner.uSpinner spnCuotas 
         Height          =   315
         Left            =   4470
         TabIndex        =   55
         Top             =   1050
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   556
         Max             =   360
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin VB.Frame fraLineaCred 
         Height          =   810
         Left            =   105
         TabIndex        =   26
         Top             =   120
         Width           =   8970
         Begin VB.ComboBox cmbTpDoc 
            Height          =   315
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Top             =   420
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cmbMoneda 
            Height          =   315
            Left            =   5640
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   420
            Width           =   1800
         End
         Begin VB.ComboBox cmbSubProducto 
            Height          =   315
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   420
            Width           =   2460
         End
         Begin VB.ComboBox cmbProductoCMACM 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   420
            Width           =   2895
         End
         Begin VB.Label lblTpDoc 
            Caption         =   "Label7"
            Height          =   195
            Left            =   7560
            TabIndex        =   86
            Top             =   210
            Width           =   1335
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   5640
            TabIndex        =   30
            Top             =   210
            Width           =   585
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Producto"
            Height          =   195
            Left            =   3120
            TabIndex        =   29
            Top             =   210
            Width           =   645
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Categoria"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   210
            Width           =   675
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Datos del Crédito"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   105
            TabIndex        =   27
            Top             =   15
            Width           =   1215
         End
      End
      Begin VB.ComboBox cmbAnalista 
         Height          =   315
         Left            =   1950
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1845
         Width           =   3180
      End
      Begin VB.TextBox txtMontoSol 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1950
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1020
         Width           =   1260
      End
      Begin VB.ComboBox cmbDestCred 
         Height          =   315
         Left            =   1920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1425
         Width           =   3180
      End
      Begin VB.ComboBox cmbInstitucion 
         Height          =   315
         Left            =   1005
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2355
         Visible         =   0   'False
         Width           =   4140
      End
      Begin VB.ComboBox cmbModular 
         Height          =   315
         Left            =   6375
         TabIndex        =   14
         Top             =   2355
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSMask.MaskEdBox txtfechaAsig 
         Height          =   315
         Left            =   7860
         TabIndex        =   11
         Top             =   1455
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label LblCondProd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   7860
         TabIndex        =   54
         Top             =   1875
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Condicion Producto :"
         Height          =   240
         Left            =   6225
         TabIndex        =   53
         Top             =   1890
         Width           =   1560
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Analista Responsable :"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   1890
         Width           =   1620
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Monto Solicitado :"
         Height          =   195
         Left            =   255
         TabIndex        =   37
         Top             =   1095
         Width           =   1275
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Plazo (Dias) :"
         Height          =   195
         Left            =   6225
         TabIndex        =   36
         Top             =   1095
         Width           =   930
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Cuotas :"
         Height          =   195
         Left            =   3660
         TabIndex        =   35
         Top             =   1095
         Width           =   585
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Asignación :"
         Height          =   195
         Left            =   6225
         TabIndex        =   34
         Top             =   1500
         Width           =   1365
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Destino de Crédito :"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   1485
         Width           =   1395
      End
      Begin VB.Line Line1 
         X1              =   165
         X2              =   9000
         Y1              =   2265
         Y2              =   2265
      End
      Begin VB.Label lblinstitucion 
         AutoSize        =   -1  'True
         Caption         =   "Institución :"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   2415
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblModular 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Modular :"
         Height          =   195
         Left            =   5235
         TabIndex        =   31
         Top             =   2415
         Visible         =   0   'False
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdexaminar 
      Caption         =   "E&xaminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   105
      Width           =   1215
   End
   Begin VB.Frame fracontrol 
      Height          =   585
      Left            =   30
      TabIndex        =   21
      Top             =   7680
      Width           =   9255
      Begin VB.CommandButton cmdEvaluar 
         Caption         =   "E&valuar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3380
         TabIndex        =   69
         Top             =   165
         Width           =   900
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7240
         TabIndex        =   50
         ToolTipText     =   "Cancelar"
         Top             =   165
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eli&minar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1920
         TabIndex        =   17
         Top             =   165
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdImprimir 
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
         Height          =   375
         Left            =   2925
         Picture         =   "frmCredSolicitud.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Imprimir Solicitud"
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8230
         TabIndex        =   20
         ToolTipText     =   "Salir"
         Top             =   165
         Width           =   975
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "Registro &Cobertura>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5330
         TabIndex        =   19
         ToolTipText     =   "Registro de Cobertura"
         Top             =   165
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.CommandButton cmdGarantias 
         Caption         =   "Garan&tías"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4300
         TabIndex        =   18
         ToolTipText     =   "Crear y/o actualizar Garantías"
         Top             =   165
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1005
         TabIndex        =   16
         Top             =   165
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CommandButton cmdCancela 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1005
         TabIndex        =   24
         Top             =   165
         Width           =   900
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   75
         TabIndex        =   23
         ToolTipText     =   "Grabar Datos"
         Top             =   165
         Width           =   930
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   15
         Top             =   165
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdRelaciones 
      Caption         =   "&Relaciones"
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
      Height          =   360
      Left            =   7500
      TabIndex        =   3
      Top             =   2130
      Width           =   1485
   End
   Begin VB.ComboBox cmbCondicion 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   560
      Width           =   1860
   End
   Begin VB.Frame fracliente 
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
      Height          =   3495
      Left            =   60
      TabIndex        =   39
      Top             =   1440
      Width           =   9180
      Begin VB.ComboBox CboAutoriazaUsoDatos 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCredSolicitud.frx":088C
         Left            =   8040
         List            =   "frmCredSolicitud.frx":0896
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   3000
         Width           =   930
      End
      Begin VB.TextBox txtDetalleMotivoRef 
         Height          =   285
         Left            =   2400
         MaxLength       =   150
         TabIndex        =   73
         Top             =   3120
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.ComboBox cmbMotivoRef 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   2760
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.CommandButton cmdSeleccionarFuentes 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   63
         Top             =   2385
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton CmdAmpliacion 
         Caption         =   "&Ampliacion"
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
         Height          =   360
         Left            =   7440
         TabIndex        =   60
         Top             =   280
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CommandButton CmdSustitucionDeudor 
         Caption         =   "&Sustit. Deudor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7440
         TabIndex        =   59
         Top             =   1680
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton CmdRefinanc 
         Caption         =   "&Refinanciar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7440
         TabIndex        =   51
         Top             =   1950
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton cmdFuentes 
         Caption         =   "&Fuentes Ingreso"
         Height          =   330
         Left            =   7440
         TabIndex        =   5
         Top             =   2400
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton CmdVerFteIngreso 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         TabIndex        =   61
         Top             =   2400
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComctlLib.ListView ListaRelacion 
         Height          =   1305
         Left            =   435
         TabIndex        =   49
         Top             =   1020
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   2302
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre de Cliente"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Relación"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Valor. Rel."
            Object.Width           =   0
         EndProperty
      End
      Begin VB.ComboBox cmbFuentes 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1845
         Visible         =   0   'False
         Width           =   4680
      End
      Begin VB.Label lblAutorizarUsoDatos 
         Caption         =   "Autorizar Uso de Datos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7440
         TabIndex        =   82
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label lblDetalleMotivoRef 
         Caption         =   "Detalle del motivo :"
         Height          =   255
         Left            =   360
         TabIndex        =   74
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblCodigo 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   720
         TabIndex        =   46
         Tag             =   "txtcodigo"
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label lblMotivoRef 
         Caption         =   "Motivo Refinanciacion:"
         Height          =   375
         Left            =   360
         TabIndex        =   66
         Top             =   2760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblSelFuentes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Seleccionar Fuentes de Ingreso"
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
         Left            =   2115
         TabIndex        =   64
         Top             =   2385
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   570
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1920
         TabIndex        =   45
         Tag             =   "txtnombre"
         Top             =   315
         Width           =   5460
      End
      Begin VB.Label lblDocnat 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2070
         TabIndex        =   44
         Tag             =   "txtdocumento"
         Top             =   675
         Width           =   1770
      End
      Begin VB.Label lblNat 
         AutoSize        =   -1  'True
         Caption         =   "Doc. de Identidad:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   435
         TabIndex        =   43
         Top             =   728
         Width           =   1320
      End
      Begin VB.Label lblJur 
         AutoSize        =   -1  'True
         Caption         =   "RUC:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4800
         TabIndex        =   42
         Top             =   735
         Width           =   390
      End
      Begin VB.Label lblDocTrib 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5400
         TabIndex        =   41
         Tag             =   "txttributario"
         Top             =   675
         Width           =   1770
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Fuentes de Ingreso :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   405
         TabIndex        =   40
         Top             =   2385
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Label lblVendedor 
      Caption         =   "Vendedor:"
      Height          =   255
      Left            =   4920
      TabIndex        =   72
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Casa Comercial :"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   5400
      TabIndex        =   68
      Top             =   240
      Width           =   1185
   End
   Begin VB.Label LblCampana 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   4680
      TabIndex        =   62
      Top             =   900
      Width           =   3580
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Campaña :"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   3840
      TabIndex        =   57
      Top             =   240
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Condición:"
      ForeColor       =   &H80000006&
      Height          =   315
      Left            =   120
      TabIndex        =   48
      Top             =   550
      Width           =   750
   End
End
Attribute VB_Name = "frmCredSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'***** add pti1 ers070-2018 18/12/2019
Private Enum TiposBusquedaNombre
    BusqApellidoPaterno = 1
    BusqApellidoMaterno = 2
    BusqApellidoCasada = 3
    BusqNombres = 4
End Enum
Dim MatPersona(1) As TActAutDatos 'add pti1 ers070-2018
Dim ultEstado As Integer
Dim sAgeReg As String
Dim dfreg As String
'fin add pti1 ers070-2018
'******** fin pti1
Dim bfCredSolicitud As Boolean 'PTI1 ADD 24082018 ERS027-2017
'MARG ERS003-2018----------
Private WithEvents Req As WinHttp.WinHttpRequest
Attribute Req.VB_VarHelpID = -1
'END MARG------------------

'agregado por vapi SEGÙN ERS TI-ERS001-2017
Dim nPresolicitudId As Integer
Dim cPersCodPreSol As String
Dim bPresol As Boolean
Dim rsPresol As ADODB.Recordset
Dim bPreSolOperacion As Boolean
Dim bPresolAmpliAuto As Boolean
'fin agreagdo por vapi
Dim nDestino As Integer '**ARLO2017113 ERS070-2017

Enum TModiffrmCredSolicitud
    Registrar = 1
    Consulta = 2
End Enum
Private oRelPersCred As UCredRelac_Cli
Private oPersona As UPersona_Cli 'COMDPersona.DCOMPersona
Private cmdEjecutar As Integer
Private nPermiso As TModiffrmCredSolicitud
Private MatCredRef As Variant
Private bRefinanciar As Boolean
Private MatTipoFte() As String
Dim fvListaCompraDeuda() As TCompraDeuda 'EJVG20160201 ERS002-2016
Dim fvListaAguaSaneamiento() As TAguaSaneamiento 'EAAS20180801 ERS054-2018
Dim fvListaCreditoVerde() As TCreditoVerde 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Dim nMontoCreditoVariable As Double 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Dim nCentinela As Integer 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Dim nSumaAguaSaneamiento As Double 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Dim nSumaCreditoVerde As Double 'EAAS20191401 SEGUN 018-GM-DI_CMACM
' CMACICA_CSTS - 10/11/2003 ---------------------------------------
Private MatCredSust As Variant
Private bSustituirDeudor As Boolean
Private bRefinanciarSustituir As Boolean
'------------------------------------------------------------------

Dim nIndtempo As Integer

'Creditos Ampliados
Private bAmpliacion  As Boolean
Private bLeasing As Boolean
Private nMontoAmpliadoAnt As Double
Public rsAmpliado As ADODB.Recordset
Public nCampanaCod As Integer

Public cCampanaDesc As String
Dim fsCtaCod As String
Dim i As Integer '*** PEAC 20080813
'ARCV 27-10-2006
Private sCodInstitucion As String
Private sCodModular As String
Private sCargo As String
Private sCARBEN As String
Private sT_Plani As String
Private rsInstituc As ADODB.Recordset

Private MatFuentes As Variant 'ARCV 29-12-2006
Private MatFteFecEval As Variant
Dim objPista As COMManejador.Pista

Dim nCasaCod As Integer 'madm 20100721

'***Modificado por ELRO 20110831, según Acta 222-2011/TI-D
Dim fnPersoneriaTitular As Integer
'WIOR 20120914 *******************************
Dim fnTipo As Integer
Dim nExisteAgeEvalCred As Integer
'WIOR FIN *************************************
'JUEZ 20130527 *********************
Dim fbRegistraEnvio As Boolean
Dim fbActivaEnvio As Boolean
Dim frsEnvEstCta As ADODB.Recordset
Dim fnModoEnvioEstCta As Integer
Dim fnDebitoMismaCta As Integer
Dim fnModoEnvioEstCtaSiNo As Boolean 'APRI20180215 ERS036-2017
'END JUEZ **************************
Private oCredAgrico As frmCredAgricoSelec 'WIOR 20130723
Private fbActivo As Boolean   'WIOR 20130723
Dim nMontoTotal As Double 'FRHU 20140424 TI-ERS015-2014
Private fbRegPromotores As Boolean  'WIOR 20140509
Dim fnMontoExpEsteCred_NEW As Double 'EJVG20160712
Dim fbEliminarEvaluacion As Boolean 'EJVG20160713
Dim sCtaCod As String 'EAAS 20180811
Dim nInicioActDa As Integer 'add pti1 ers070-2018

'JOEP20190919 ERS042 CP-2018
Dim cValorIni As String
Dim cmdBtnCancelar As Integer
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 CP-2018
Dim MinMonto As Double, MaxMonto As Double
Dim MinCuota As Integer, MaxCuota As Integer
Dim MinPlazo As Integer, MaxPlazo As Integer
Dim nMatMontoPre As Variant
Dim nTpCmbTpDoc As Integer
Dim nTpDoc As Long, nTpIngr As Long, nTpInt As Long, nSubDestino As Long, nSubDestAnt As Long
Dim bEntrotxtMontoSol As Boolean
'JOEP20190919 ERS042 CP-2018
Dim sCodTitular As String 'ARLO20181126 ERS68-2018

Property Let PersoneriaTitular(ByVal nNewValue As Integer)
fnPersoneriaTitular = nNewValue
End Property

Property Get PersoneriaTitular() As Integer
PersoneriaTitular = fnPersoneriaTitular
End Property
'*********************************************************

Private Sub DefineCondicionCredito()
Dim oCred As COMDCredito.DCOMCredito
Dim nValor As Integer
'Dim sTipoProducto As String
'sTipoProducto = Right(Trim(cmbProductoCMACM.Text), 5)
'sTipoProducto = IIf(Trim(sTipoProducto) = "", "000", sTipoProducto)
    Set oCred = New COMDCredito.DCOMCredito
    'nValor = oCred.DefineCondicionCredito(oRelPersCred.TitularPersCod, sTipoProducto, gdFecSis, bRefinanciar)
    'nValor = oCred.DefineCondicionCredito(oRelPersCred.TitularPersCod, , gdFecSis, bRefinanciar)
    If Not (oRelPersCred Is Nothing) Then 'WIOR 20140820
        'nValor = oCred.DefineCondicionCredito(oRelPersCred.TitularPersCod, , gdFecSis, bRefinanciar, val(spnCuotas.valor)) 'EJVG20130503
        nValor = oCred.DefineCondicionCredito(oRelPersCred.TitularPersCod, , gdFecSis, bRefinanciar, 0) 'WIOR 20141210
    End If 'WIOR 20140820
    
    Set oCred = Nothing
    'ARCV 20-02-2007
    If bAmpliacion Then
        cmbCondicion.ListIndex = IndiceListaCombo(cmbCondicion, 5)
    Else
        cmbCondicion.ListIndex = IndiceListaCombo(cmbCondicion, Trim(str(nValor)))
    End If
    '-----
    'STS - 04112003 - PARA ESTE CASO, NO ES NECESARIO
    'cmbCondicionOtra.ListIndex = IndiceListaCombo(cmbCondicionOtra, Trim(Str(nValor)))
    '--------------------------------------------------------------------------
End Sub

Public Sub RefinanciaCredito(ByVal pnPermiso As TModiffrmCredSolicitud)
    bRefinanciar = True
    bLeasing = False
    bSustituirDeudor = False
    bAmpliacion = False 'JUEZ 20130719
    nPermiso = pnPermiso
    CmdRefinanc.Visible = True
    'By Capi 28102008
    lblMotivoRef.Visible = True
    cmbMotivoRef.Visible = True
    '
    'JAME20140509 ***
    lblDetalleMotivoRef.Visible = True
    txtDetalleMotivoRef.Visible = True
    'END JAME *******
    txtMontoSol.Enabled = False
    Me.Caption = "Solicitud de Refinanciacion"
    Me.Show 1
End Sub

Public Sub SustitucionCredito(ByVal pnPermiso As TModiffrmCredSolicitud)

    bRefinanciar = False
    bSustituirDeudor = True
    nPermiso = pnPermiso
    CmdRefinanc.Visible = False
    CmdSustitucionDeudor.Visible = True
    txtMontoSol.Enabled = False
    bLeasing = False
    Me.Caption = "Solicitud de Sustitución de Deudor"
    Me.Show 1
End Sub
Public Sub AmpliacionCredito(ByVal pnPermiso As TModiffrmCredSolicitud)

    bAmpliacion = True
    bRefinanciar = False 'JUEZ 20130719
    'bSustituirDeudor = True
    nPermiso = pnPermiso
    'CmdRefinanc.Visible = False
    'CmdSustitucionDeudor.Visible = True
    CmdAmpliacion.Visible = True
    txtMontoSol.Enabled = True
    bLeasing = False
    Me.Caption = "Ampliacion de un Credito"
    chkAutAmpliacion.Visible = True 'JUEZ 20160509
    Me.Show 1
End Sub


Public Sub Inicio(ByVal pnPermiso As TModiffrmCredSolicitud)
    bRefinanciar = False
    bSustituirDeudor = False
    bAmpliacion = False 'ARCV 13-03-2007
    nPermiso = pnPermiso
    bLeasing = False
    ChkCap.Visible = False
    'WIOR 20120914*****************************
    'Dim NCredito As COMNCredito.NCOMCredito
    fnTipo = CInt(pnPermiso)
    'Set NCredito = New COMNCredito.NCOMCredito
    'nExisteAgeEvalCred = NCredito.ObtieneAgenciaCredEval(gsCodAge)
    
    'If nExisteAgeEvalCred = 0 Then
    '    Me.cmdEvaluar.Visible = False
    'Else
    '    Me.cmdEvaluar.Visible = True
    '    Me.cmdEvaluar.Enabled = False
    'End If
    'WIOR FIN ************************************
    cmdEnvioEstCta.Visible = True 'JUEZ 20130527
    
    
    'agregado por vapi SEGÙN ERS TI-ERS001-2017
    If nPermiso = Registrar Then
        cmdpresolicitud.Visible = True
    Else
        cmdpresolicitud.Visible = False
    End If
    'fin agregado por vapi
    
    
    Me.Show 1
End Sub
Public Sub Inicioleasing(ByVal pnPermiso As TModiffrmCredSolicitud)
    bRefinanciar = False
    bSustituirDeudor = False
    bAmpliacion = False 'ARCV 13-03-2007
    nPermiso = pnPermiso
    bLeasing = True
    ChkCap.Visible = False
    Me.Caption = "Solicitud de Arrendamiento Financiero"
    ActXCtaCred.texto = "Operación"
    Label18.Caption = "Datos de la Operación"
    Label10.Caption = "Destino OP"
    cmdNuevo.Enabled = False 'EJVJ20120720
    Me.Show 1
End Sub

Private Sub ControlesPermiso()
    If nPermiso = Consulta Then
        cmdNuevo.Enabled = False
        cmdEditar.Enabled = False
        cmdEliminar.Enabled = False
        cmdGarantias.Enabled = False
        cmdGravar.Enabled = False
        cmdEvaluar.Enabled = False
        ''WIOR 20120914 ******************************************
        'If nExisteAgeEvalCred = 0 Then
        '    Me.cmdEvaluar.Visible = False
        'Else
        '    Me.cmdEvaluar.Visible = True
        '    Me.cmdEvaluar.Enabled = False
        'End If
        ''WIOR FIN ***********************************************
    End If
    If nPermiso = Registrar Then
        cmdNuevo.Enabled = True
        cmdEditar.Enabled = True
        cmdEliminar.Enabled = True
        cmdGarantias.Enabled = True
        cmdGravar.Enabled = True
        cmdEvaluar.Enabled = True
        ''WIOR 20120914 ******************************************
        'If nExisteAgeEvalCred = 0 Then
        '    Me.cmdEvaluar.Visible = False
        'Else
        '    Me.cmdEvaluar.Visible = True
        '    Me.cmdEvaluar.Enabled = True
        'End If
        ''WIOR FIN ***********************************************
    End If
End Sub

Private Function ExisteTitular() As Boolean
    ExisteTitular = False
    oRelPersCred.IniciarMatriz
    Do While Not oRelPersCred.EOF
        If oRelPersCred.ObtenerValorRelac = gColRelPersTitular Then
            ExisteTitular = True
            Exit Do
        End If
        oRelPersCred.siguiente
    Loop
End Function

Private Function ValidaDatos() As Boolean
Dim oCred As COMNCredito.NCOMCredito
Dim nMontoFte As Double
Dim sMonedaFteCod As String
Dim sValor As String
Dim i As Integer
    
    ValidaDatos = True
    
    If Not ExisteTitular Then
        MsgBox "Debe Ingresar el Titular del Credito", vbInformation, "Aviso"
        ValidaDatos = False
        cmdRelaciones.SetFocus
        Exit Function
    End If
    'Valida Seleccion de Condicion del Credito
    If cmbCondicion.ListIndex = -1 Then
        MsgBox "Debe Seleccionar la Condicion del Credito", vbInformation, "Aviso"
        ValidaDatos = False
        cmbCondicion.SetFocus
        Exit Function
    End If
    
    ' CMACICA_CSTS - 04112003 --------------------------------------------------------
    'Valida Seleccion de Condicion2 del
    
    If cmbCondicionOtra.ListIndex = -1 Then
        MsgBox "Debe Seleccionar la Condicion 2 del Credito", vbInformation, "Aviso"
        ValidaDatos = False
        'MAVM 200907
        'CmbCondicionOtra.SetFocus
        Exit Function
    End If
    '---------------------------------------------------------------------------------
    
    'Valida Ingreso de Relaciones de Cliente
    If oRelPersCred.NroRelaciones = 0 Then
        MsgBox "Debe Ingresar por lo Menos el Titular de la Cuenta", vbInformation, "Aviso"
        cmdRelaciones.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    
    '**** COMENTADO POR PEAC 20080412
    
'    'Valida Selecion de Fuentes de Ingreso
'    'If cmbFuentes.ListIndex = -1 Then
'    If Not IsArray(MatFuentes) Then
'        MsgBox "Debe Selecionar un Fuente de Ingreso para el Credito", vbInformation, "Aviso"
'        ValidaDatos = False
'        Exit Function
'    Else
'        If UBound(MatFuentes) = 0 Then
'            MsgBox "Debe Selecionar un Fuente de Ingreso para el Credito", vbInformation, "Aviso"
'            ValidaDatos = False
'            Exit Function
'        End If
'    End If
    '**** FIN PEAC
    
    
    'Valida Selecion de Tipo de Credito
    If cmbProductoCMACM.ListIndex = -1 Then
        MsgBox "Debe Selecionar el Tipo del Credito", vbInformation, "Aviso"
        cmbProductoCMACM.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    'Valida Selecion de Sub Tipo de Credito
    If cmbSubProducto.ListIndex = -1 Then
        MsgBox "Debe Selecionar el Sub Tipo del Credito", vbInformation, "Aviso"
        cmbSubProducto.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    'Valida Selecion de Moneda
    If cmbMoneda.ListIndex = -1 Then
        MsgBox "Debe Selecionar la Moneda del Credito", vbInformation, "Aviso"
        cmbMoneda.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    'Valida Monto Solicitado
    If Trim(txtMontoSol.Text) = "" Or txtMontoSol.Text = "0.00" Then
        MsgBox "Debe Ingresar el Monto del Prestamo ", vbInformation, "Aviso"
        If txtMontoSol.Enabled Then
            txtMontoSol.SetFocus
        End If
        ValidaDatos = False
        Exit Function
    End If
    
    'Valida Nro de Cuotas
    If spnCuotas.valor = "" Or spnCuotas.valor = "0" Then
        MsgBox "Debe Ingresar el Numero de Cuotas del Credito ", vbInformation, "Aviso"
        spnCuotas.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    'Valida Plazo de las Cuotas
    If spnPlazo.valor = "" Or spnPlazo.valor = "0" Then
        MsgBox "Debe Ingresar el plazo de las Cuotas del Credito ", vbInformation, "Aviso"
        spnPlazo.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    'Valida Seleccion del Destino del Credito
    If cmbDestCred.ListIndex = -1 Then
        MsgBox "Debe Seleccionar el Destino del Credito ", vbInformation, "Aviso"
        If cmbDestCred.Visible And cmbDestCred.Enabled Then cmbDestCred.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    'JOEP20190919 ERS042 CP-2018 Valida Seleccion del Sub Destino del Credito si es que tiene
    If cmbTpDoc.Visible = True And Trim(Right(cmbSubProducto.Text, 10)) = "520" And spnCuotas.valor = 1 And cmbTpDoc.ListIndex = -1 Then 'ARLO20190304
        MsgBox "Ingrese el Tipo de Documento del Crédito", vbInformation, "Aviso"
        ValidaDatos = False
        cmbTpDoc.SetFocus
        Exit Function
    End If
    
    If cmbTpDoc.Visible = True And cmbTpDoc.Text = "" And Trim(Right(cmbSubProducto.Text, 10)) <> "520" Then
        MsgBox "Ingrese el Tipo de Documento del Crédito", vbInformation, "Aviso"
        ValidaDatos = False
        cmbTpDoc.SetFocus
        Exit Function
    End If
    
    If cmbSubDestCred.Visible = True And cmbSubDestCred.Text = "" Then
        MsgBox "Ingrese el Sub Destino de Crédito", vbInformation, "Aviso"
        ValidaDatos = False
        cmbSubDestCred.SetFocus
        Exit Function
    End If
    
    If cmbTpDoc.Visible = True And Trim(Right(cmbSubProducto.Text, 10)) = "703" Then
        If IsArray(nMatMontoPre) = False Then
            MsgBox "Ingrese el Monto para el Tipo de Interés", vbInformation, "Aviso"
            ValidaDatos = False
            txtMontoSol.Text = "0.00"
            cmbTpDoc.SetFocus
            Exit Function
        End If
    End If
    
    If Trim(Right(cmbSubProducto.Text, 10)) = "525" Then
        If cmbTpDoc.Visible = True Then
            If cmbTpDoc.Text = "" Then
                MsgBox "Ingrese el Tipo de Documento del Crédito", vbInformation, "Aviso"
                ValidaDatos = False
                cmbTpDoc.SetFocus
                Exit Function
            ElseIf cmbTpDoc.Text <> "" And Trim(Right(cmbTpDoc.Text, 8)) = 22004 And CInt(spnPlazo.valor) > 90 Then
                MsgBox "Plazo maxímo es 90", vbInformation, "Aviso"
                ValidaDatos = False
                cmbTpDoc.SetFocus
                Exit Function
            ElseIf cmbTpDoc.Text <> "" And CInt(spnPlazo.valor) > 120 Then
                MsgBox "Plazo maxímo es 120", vbInformation, "Aviso"
                ValidaDatos = False
                cmbTpDoc.SetFocus
                Exit Function
            End If
        End If
    End If
'JOEP20190919 ERS042 CP-2018
        
    'Valida Seleccion del Analista del Credito
    If cmbAnalista.ListIndex = -1 Then
        MsgBox "Debe Seleccionar el Analista del Credito ", vbInformation, "Aviso"
        cmbAnalista.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    'Valida Seleccion de la Institucion
    'ARCV 27-10-2006
    'If CInt(Trim(Right(cmbSubTipo.Text, 5))) = gColConsuDctoPlan And cmbInstitucion.ListIndex = -1 Then
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If ((objProducto.GetResultadoCondicionCatalogo("N0000080", Trim(Right(cmbSubProducto.Text, 10)))) = gColProConsumoPerDesPla) And sCodInstitucion = "" Then
'**ARLO20180712 ERS042 - 2018
    'If (CInt(Trim(Right(cmbSubProducto.Text, 5))) = 512 Or CInt(Trim(Right(cmbSubProducto.Text, 5))) = gColProConsumoPerDesPla) And sCodInstitucion = "" Then
        MsgBox "Debe Seleccionar la Institucion para el Descuento por Planilla ", vbInformation, "Aviso"
        'cmbInstitucion.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    'Valida Ingreso del Codigo Modular
    'If CInt(Trim(Right(cmbSubProducto.Text, 5))) = gColConsuDctoPlan And sCodModular = "" Then
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If ((objProducto.GetResultadoCondicionCatalogo("N0000080", Trim(Right(cmbSubProducto.Text, 10)))) = gColProConsumoPerDesPla) And sCodInstitucion = "" Then
    '**ARLO20180712 ERS042 - 2018
    'If (CInt(Trim(Right(cmbSubProducto.Text, 5))) = 512 Or CInt(Trim(Right(cmbSubProducto.Text, 5))) = gColProConsumoPerDesPla) And sCodModular = "" Then
        MsgBox "Debe Ingresar el Codigo Modular del Cliente", vbInformation, "Aviso"
        'cmbModular.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    'Valida fecha de Solicitud
    If Len(ValidaFecha(txtfechaAsig.Text)) > 0 Then
        MsgBox ValidaFecha(txtfechaAsig.Text), vbInformation, "Aviso"
        txtfechaAsig.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    'WIOR 20140509 ***************************
    'If fbRegPromotores And fraPromotor.Visible Then 'LUCV20171013, Comentó, según INC1710120009
    If fbRegPromotores And fraPromotor.Enabled Then  'LUCV20171013, Agregó, según INC1710120009
        If cmbPromotor.ListIndex = -1 Then
'***** LUCV20170425, Comentó
'            MsgBox "Debe Seleccionar el Promotor del Credito ", vbInformation, "Aviso"
'            cmbPromotor.SetFocus
'            ValidaDatos = False
'            Exit Function

            If MsgBox("¿Desea registrar al promotor del crédito?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                cmbPromotor.SetFocus
                ValidaDatos = False
                Exit Function
            Else
                'fraPromotor.Visible = True 'LUCV20171013, Comentó, según INC1710120009
                 fraPromotor.Enabled = True 'LUCV20171013, Agregó, según INC1710120009
            End If
'***** LUCV20170425, Fin
        End If
    End If
    'WIOR FIN ********************************
    
    'Valida Caducidad de Fuente de Ingreso
    Dim nPos As Integer
    Dim rsFteIng As ADODB.Recordset
    Dim rsFIDep As ADODB.Recordset
    Dim rsFIInd As ADODB.Recordset
    
    Set oCred = New COMNCredito.NCOMCredito
    
    ReDim MatFteFecEval(0)
    'JAME20142803************************* Valida Detalle Motivo Refinanciación
    If bRefinanciar Then
        If Len(Trim(txtDetalleMotivoRef.Text)) = 0 Then
             MsgBox "Debe ingresar el Motivo del Detalle de la Refinanciación ", vbInformation, "Aviso"
             ValidaDatos = False
             If txtDetalleMotivoRef.Visible And txtDetalleMotivoRef.Enabled Then txtDetalleMotivoRef.SetFocus
             Exit Function
         End If
     End If
    'Fin Jame******************************
    'EJVG20160203 ERS002-2016*** Compra Deuda
    If val(Trim(Right(cmbDestCred, 3))) = ColocDestino.gColocDestinoCambEstructPasivo Then
        If Not IsArray(fvListaCompraDeuda) Then
            ReDim fvListaCompraDeuda(0)
        End If
        If Not bRefinanciar Then 'ARLO20180319
            If UBound(fvListaCompraDeuda) <= 0 Then
                MsgBox "Ud. debe ingresar el detalle de la compra de deuda.", vbInformation, "Aviso"
                EnfocaControl cmdDestinoDetalle
                ValidaDatos = False
                Exit Function
            End If
        End If 'ARLO20180319
    End If
    'END EJVG *******
    'Valida Nombre del Vendedor
    'Jame 20142606 *************************
    If txtVendedor.Visible = True And txtVendedor.Text = "" Then
        MsgBox "Debe Ingresar el Nombre del Vendedor", vbInformation, "Aviso"
        txtVendedor.SetFocus
        ValidaDatos = False
    End If
    'Fin Jame ******************************
    
    '**ARLO2017113 --INICIO ERS070-217
    
    nDestino = Trim(Right(cmbDestCred.Text, 5))
    
    Dim nCantidad As Integer
    Dim maxValue As Double
    Dim lvListaCompraDeudaNew(1) As TCompraDeuda
    '**ARLO20180531 INICIO
    Dim oTC  As New COMDConstSistema.NCOMTipoCambio
    Dim nTpoC As Double
    Dim nMontoSol, nSaldoComp, nDesem As Double
    nTpoC = CDbl(oTC.EmiteTipoCambio(gdFecSis, TCFijoDia))
    nMontoSol = IIf(Trim(Right(cmbMoneda.Text, 1)) = 1, val(txtMontoSol.Text), val(txtMontoSol.Text) * nTpoC)
    '**ARLO20180531 FIN
    
    '**ARLO20180531 MODIFICO INICIO
    For i = 1 To UBound(fvListaCompraDeuda)
        If (fvListaCompraDeuda(i).nmoneda) = 1 Then
            nSaldoComp = fvListaCompraDeuda(i).nSaldoComprar
        Else
            nSaldoComp = fvListaCompraDeuda(i).nSaldoComprar * nTpoC
        End If
        maxValue = IIf((fvListaCompraDeuda(1).nmoneda) = 1, fvListaCompraDeuda(1).nSaldoComprar, fvListaCompraDeuda(1).nSaldoComprar * nTpoC)
        If maxValue < nSaldoComp Then
        maxValue = nSaldoComp
        End If
    Next
    
    For i = 1 To UBound(fvListaCompraDeuda)
        If (fvListaCompraDeuda(i).nmoneda) = 1 Then
            nSaldoComp = fvListaCompraDeuda(i).nSaldoComprar
        Else
            nSaldoComp = fvListaCompraDeuda(i).nSaldoComprar * nTpoC
        End If
        If maxValue = nSaldoComp Then
        lvListaCompraDeudaNew(1) = fvListaCompraDeuda(i)
        End If
    Next
    
    For i = 1 To UBound(lvListaCompraDeudaNew)
        If (lvListaCompraDeudaNew(i).nmoneda) = 1 Then
            nSaldoComp = lvListaCompraDeudaNew(i).nSaldoComprar
            nDesem = lvListaCompraDeudaNew(i).nMontoDesembolso
            
        Else
            nSaldoComp = lvListaCompraDeudaNew(i).nSaldoComprar * nTpoC
            nDesem = lvListaCompraDeudaNew(i).nMontoDesembolso * nTpoC
        End If
    Next
    
    For i = 1 To UBound(lvListaCompraDeudaNew)
        If (lvListaCompraDeudaNew(i).nDestino = 3) Then
            If (nSaldoComp <= 1000) Then
            nCantidad = 6
            ElseIf (nSaldoComp > 1000 And nSaldoComp <= 5000) Then
            nCantidad = 12
            ElseIf (nSaldoComp > 5000 And nSaldoComp <= 15000) Then
            nCantidad = 18
            ElseIf (nSaldoComp > 15000) Then
            nCantidad = 24
            End If
        ElseIf (nMontoSol >= nDesem) Then
            nCantidad = lvListaCompraDeudaNew(i).nNroCuotasPactadas
            nCantidad = nCantidad + Math.Round(nCantidad * 0.4)
        ElseIf (nMontoSol > nSaldoComp And nMontoSol < nDesem) Then
            nCantidad = lvListaCompraDeudaNew(i).nNroCuotasPactadas
        Else
            nCantidad = lvListaCompraDeudaNew(i).nNroCuotasPactadas - lvListaCompraDeudaNew(i).nNroCuotasPagadas
            nCantidad = nCantidad + Math.Round(nCantidad * 0.4)
        End If
    Next
    '**ARLO20180531 MODIFICO FIN
    If Not bRefinanciar Then '**ARLO20180317  ERS070-217
    'ARLO20180712 ERS042 - 2018
        Set objProducto = New COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018
        If objProducto.GetResultadoCondicionCatalogo("N0000063", Trim(Right(cmbSubProducto.Text, 10))) Then
    'ARLO20180712 ERS042 - 2018
        'If (CInt(Trim(Right(cmbSubProducto.Text, 5)))) <> 704 Then
            If (nDestino = 14) Then
                If CInt(spnCuotas.valor) > nCantidad Then
                    MsgBox "El número de cuotas debe ser menor o igual a " & nCantidad, vbInformation, "Aviso" 'ARLO20180321
                    spnCuotas.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
            End If
        End If
    End If '**ARLO20180317  ERS070-217
    Set oTC = Nothing
    '**ARLO2017113 --FIN ERS070-217
    
    '**ARLO20180315 INICIO ERS070 - 2017 ANEXO 01
    Dim Y As Integer
    Dim nTotalCompra As Double

    If Not bRefinanciar Then
        If (nDestino = 14) Then
            For Y = 1 To UBound(fvListaCompraDeuda)
                
                If (fvListaCompraDeuda(Y).nmoneda) = 1 Then
                    nTotalCompra = nTotalCompra + fvListaCompraDeuda(Y).nSaldoComprar
                Else
                    nTotalCompra = nTotalCompra + fvListaCompraDeuda(Y).nSaldoComprar * nTpoC
                End If
                
            Next Y
            
            If (CDbl(nMontoSol) < nTotalCompra) Then
                    MsgBox "Monto Solicitado debe ser mayor o igual al Saldo a Comprar", vbInformation, "Alerta"
                    ValidaDatos = False
                    Exit Function
            End If
        End If
    End If
    Set oTC = Nothing
    '**ARLO20180315 FIN ERS070 - 2017 ANEXO 01
    
    'Call oCred.CargarFtesIngreso(rsFteIng, rsFIDep, rsFIInd, oRelPersCred.TitularPersCod, , cmbFuentes.ListIndex)
    
'    For i = 0 To UBound(MatFuentes) - 1
'
'        Call oCred.CargarFtesIngreso(rsFteIng, rsFIDep, rsFIInd, oRelPersCred.TitularPersCod, , MatFuentes(i))
'
'        Call oPersona.RecuperaFtesdeIngreso(oRelPersCred.TitularPersCod, rsFteIng)
'        Call oPersona.RecuperaFtesIngresoDependiente(MatFuentes(i), rsFIDep)
'        Call oPersona.RecuperaFtesIngresoIndependiente(MatFuentes(i), rsFIInd)
'
'        ReDim Preserve MatFteFecEval(UBound(MatFteFecEval) + 1)
'        MatFteFecEval(UBound(MatFteFecEval) - 1) = oPersona.ObtenerFteIngFecEval(MatFuentes(i), IIf(oPersona.ObtenerFteIngIngresoTipo(MatFuentes(i)) = gPersFteIngresoTipoDependiente, oPersona.ObtenerFteIngIngresoNumeroFteDep(MatFuentes(i)) - 1, oPersona.ObtenerFteIngIngresoNumeroFteIndep(MatFuentes(i)) - 1))
'
'        'nPos = oPersona.UbicaPosicionFteIngreso(Trim(Mid(cmbFuentes.Text, 100, 20)), CDate(Trim(Right(cmbFuentes.Text, 20))), MatFuentes(i))
'        If gdFecSis >= oPersona.ObtenerFteIngFecCaducac(MatFuentes(i), IIf(oPersona.ObtenerFteIngIngresoTipo(MatFuentes(i)) = gPersFteIngresoTipoDependiente, oPersona.ObtenerFteIngIngresoNumeroFteDep(MatFuentes(i)) - 1, oPersona.ObtenerFteIngIngresoNumeroFteIndep(MatFuentes(i)) - 1)) Then
'            MsgBox "Fuente de Ingreso a Caducado Ingrese otra Fuente de Ingreso Actual", vbInformation, "Aviso"
'            'cmbFuentes.SetFocus
'            ValidaDatos = False
'            Exit Function
'        End If
'
'        'Valida Fuente de Ingreso de Credito Pyme y Comercial Sea una Fuente de Ingreso Independiente
'        If CInt(Right(cmbSubTipo.Text, 3)) = gColPYMEAgro Or _
'            CInt(Right(cmbSubTipo.Text, 3)) = gColPYMEEmp Or _
'            CInt(Right(cmbSubTipo.Text, 3)) = gColPYMEPesq Or _
'            CInt(Right(cmbSubTipo.Text, 3)) = gColComercEmp Or _
'            CInt(Right(cmbSubTipo.Text, 3)) = gColComercAgro Or _
'            CInt(Right(cmbSubTipo.Text, 3)) = gColComercPesq Then
'            'nPos = oPersona.UbicaPosicionFteIngreso(Trim(Mid(cmbFuentes.Text, 100, 20)), CDate(Trim(Right(cmbFuentes.Text, 20))))
'            If CInt(oPersona.ObtenerFteIngTipo(MatFuentes(i))) <> gPersFteIngresoTipoIndependiente Then
'                MsgBox "Debe Seleccionar una Fuente de Ingreso Independiente para Este tipo de Credito", vbInformation, "Aviso"
'                'cmbFuentes.SetFocus
'                ValidaDatos = False
'                Exit Function
'            End If
'        End If
'
'        'Valida Fuente de Ingreso de Credito Consumo Sea una Fuente de Ingreso Dependiente
'        If CInt(Right(cmbSubTipo.Text, 3)) = gColConsuDctoPlan Or _
'            CInt(Right(cmbSubTipo.Text, 3)) = gColConsuPlazoFijo Or _
'            CInt(Right(cmbSubTipo.Text, 3)) = gColConsCTS Or _
'            CInt(Right(cmbSubTipo.Text, 3)) = gColConsuUsosDiv Or _
'            CInt(Right(cmbSubTipo.Text, 3)) = gColConsuPrendario Or _
'            CInt(Right(cmbSubTipo.Text, 3)) = gColConsuPrestAdm Then
'            'nPos = oPersona.UbicaPosicionFteIngreso(Trim(Mid(cmbFuentes.Text, 100, 20)), CDate(Trim(Right(cmbFuentes.Text, 20))))
'            If CInt(oPersona.ObtenerFteIngTipo(MatFuentes(i))) <> gPersFteIngresoTipoDependiente Then
'                MsgBox "Debe Seleccionar una Fuente de Ingreso Dependiente para Este tipo de Credito", vbInformation, "Aviso"
'                'cmbFuentes.SetFocus
'                ValidaDatos = False
'                Exit Function
'            End If
'        End If
'
'        '23/092004:LMMD Desabilitado por recomendaciones de Javier Cabrera
'        'Valida que la Institucion y la Fuente de Igreso sean las mismas
'
'    '    If CInt(Trim(Right(cmbSubTipo.Text, 10))) = gColConsuDctoPlan Then
'    '        If Trim(Mid(cmbFuentes.Text, 100, 20)) <> Trim(Right(cmbInstitucion.Text, 20)) Then
'    '            MsgBox "La Fuente de Ingreso no Pertenece a la Institucion", vbInformation, "Aviso"
'    '            cmbFuentes.SetFocus
'    '            ValidaDatos = False
'    '            Exit Function
'    '        End If
'    '    End If
'
'        'CMACICA_CSTS - 25/11/2003 ------------------------------------------------------------------------------------------
'        'Valida el Monto Total a la Fecha (Otros Prestamos Sistema Financiero + Prestamos CMAC + Monto del Prestamo) para distingur un Credito Mes y un Comercial
'        If CInt(Right(cmbSubTipo.Text, 3)) = gColPYMEAgro Or _
'            CInt(Right(cmbSubTipo.Text, 3)) = gColPYMEEmp Or _
'            CInt(Right(cmbSubTipo.Text, 3)) = gColPYMEPesq Or _
'            CInt(Right(cmbSubTipo.Text, 3)) = gColComercEmp Or _
'            CInt(Right(cmbSubTipo.Text, 3)) = gColComercAgro Or _
'            CInt(Right(cmbSubTipo.Text, 3)) = gColComercPesq Then
'
'            nMontoFte = oPersona.ObtenerFteIngCreditosCmact(MatFuentes(i), IIf(oPersona.ObtenerFteIngIngresoTipo(MatFuentes(i)) = gPersFteIngresoTipoDependiente, oPersona.ObtenerFteIngIngresoNumeroFteDep(MatFuentes(i)) - 1, oPersona.ObtenerFteIngIngresoNumeroFteIndep(MatFuentes(i)) - 1))
'            nMontoFte = nMontoFte + oPersona.ObtenerFteIngOtrosCreditos(MatFuentes(i), IIf(oPersona.ObtenerFteIngIngresoTipo(MatFuentes(i)) = gPersFteIngresoTipoDependiente, oPersona.ObtenerFteIngIngresoNumeroFteDep(MatFuentes(i)) - 1, oPersona.ObtenerFteIngIngresoNumeroFteIndep(MatFuentes(i)) - 1))
'            sMonedaFteCod = oPersona.ObtenerFteIngMoneda(MatFuentes(i))
'
'            Set oCred = New COMNCredito.NCOMCredito
'            sValor = oCred.ValidaMontoParaTipoCredito(Mid(Right(cmbTipoCred.Text, 3), 1, 2), Trim(Right(cmbMoneda.Text, 2)), CDbl(txtMontoSol.Text), sMonedaFteCod, nMontoFte, gdFecSis)
'            If sValor <> "" Then
'                If MsgBox(sValor & vbCrLf & "Desea continuar", vbInformation + vbQuestion, "Aviso") = vbNo Then
'                    'cmbTipoCred.SetFocus
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            End If
'
'            Set oCred = Nothing
'        End If
'        '------------------------------------------------------------------------------------------------------------------------
'    Next i

'JOEP20171107 3097-2017-GM Acta193.
If Trim(Right(cmbCondicion.Text, 5)) <> "4" Then 'JOEP20190301 CP
    Dim rsValDes As ADODB.Recordset
    Dim obDCredValDes As COMDCredito.DCOMCredito
    Set obDCredValDes = New COMDCredito.DCOMCredito
    
    Set rsValDes = obDCredValDes.ValidadDestinoConsEmp(CInt(Trim(Right(cmbProductoCMACM.Text, 5))), CInt(Trim(Right(cmbSubProducto.Text, 5))), CInt(Trim(Right(cmbDestCred.Text, 5))))
    
    If Not (rsValDes.EOF And rsValDes.BOF) Then
        If rsValDes!cMensaje <> "" Then
            MsgBox rsValDes!cMensaje, vbInformation, "Aviso" 'EAAS201809 CAMBIO DE TITULO DE MENSAJE
            rsValDes.Close
            Set obDCredValDes = Nothing
            ValidaDatos = False
            Exit Function
        End If
    rsValDes.Close
    Set obDCredValDes = Nothing
    End If
End If 'JOEP20190301 CP
'JOEP20171107 3097-2017-GM Acta193.

'INICIO EAAS20180815 /CAMBIO DE CODIGO EN LA LOGICA SEGUN ACTA 147 EAAS20180904
Dim obDCredValDesAguaSaneamiento As COMDCredito.DCOMCredito
Set obDCredValDesAguaSaneamiento = New COMDCredito.DCOMCredito

Dim nProducto As Integer, npdestino As Integer, sCodigo As String, nCuotas As Integer, sMensaje As String

nProducto = CInt(Trim(Right(cmbProductoCMACM.Text, 5)))
npdestino = CInt(Trim(Right(cmbDestCred.Text, 5)))
sCodigo = lblCodigo
nCuotas = spnCuotas.valor

If UBound(fvListaAguaSaneamiento) > 0 Then
    Call obDCredValDesAguaSaneamiento.getMensajeValidacion(nProducto, npdestino, sCodigo, sMensaje, nCuotas)
        If sMensaje <> "" Then
            MsgBox sMensaje, vbInformation, "Aviso"
            Set obDCredValDesAguaSaneamiento = Nothing
            ValidaDatos = False
            Exit Function
        End If

    Set obDCredValDesAguaSaneamiento = Nothing
End If
'FIN EAAS20180815 /CAMBIO DE CODIGO EN LA LOGICA SEGUN ACTA 147 EAAS20180904

'JOEP ERS004-2016 Si Cumple Condicion para Credito Paralelo

'ARLO20180712 ERS042 - 2018
Set objProducto = New COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018
If objProducto.GetResultadoCondicionCatalogo("N0000064", Trim(Right(cmbSubProducto.Text, 10))) Then     '**END ARLO
'ARLO20180712 ERS042 - 2018

'If (CInt(Trim(Right(cmbSubProducto.Text, 5)))) = 513 Then

    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Dim rs As ADODB.Recordset
    Set rs = oDCOMFormatosEval.ValidarParalelo(lblCodigo)

    If Not (rs.EOF And rs.BOF) Then
        If Left(rs!cTpoCredCod, 1) = 7 Or Left(rs!cTpoCredCod, 1) = 8 Then
        MsgBox "El Cliente no cumple condición para Créditos Paralelo: " & Chr(13) & "" & Chr(13) & _
               "No Aplica a los Tipos de Créditos: " & Chr(13) & _
               "CREDITO DE CONSUMO NO REVOLVENTE" & Chr(13) & _
               "CREDITOS HIPOTECARIOS PARA VIVIENDA", vbInformation, "Alerta"
        ValidaDatos = False
        Exit Function
        End If
    End If

End If
'FIN JOEP ERS004-2016 Si Cumple Condicion para Credito Paralelo

'INICIO EAAS20180815
If Trim(Right(cmbCondicion.Text, 5)) <> "4" Then 'JOEP20190301 CP
If (UBound(fvListaAguaSaneamiento) = 0 And CInt(Trim(Right(cmbDestCred.Text, 5))) = 26) Then
                    MsgBox "Ingrese el detalle del destino Agua y saneamiento", vbInformation, "Alerta"
                    ValidaDatos = False
                    Exit Function
End If
    Dim nSumaTotalAguaSaneamiento As Double
    nSumaTotalAguaSaneamiento = 0
    Dim ixCD As Integer
    For ixCD = 1 To UBound(fvListaAguaSaneamiento)
        nSumaTotalAguaSaneamiento = nSumaTotalAguaSaneamiento + fvListaAguaSaneamiento(ixCD).nMontoS
    Next
    If (nSumaTotalAguaSaneamiento > CDbl(txtMontoSol.Text)) Then
                    MsgBox "La suma de los subdestinos de agua y saneamiento es mayor al monto solicitado", vbInformation, "Alerta"
                    ValidaDatos = False
                    Exit Function
    End If
    
    If (nSumaTotalAguaSaneamiento <> CDbl(txtMontoSol.Text) And nDestino = 26) Then
                    MsgBox "La suma de los subdestinos de agua y saneamiento debe ser igual al monto solicitado", vbInformation, "Alerta"
                    ValidaDatos = False
                    Exit Function
    End If
End If 'JOEP20190301 CP
'End Function
'END EAAS20180815

'add pti1 ers070-2018 18/12/2018***********************
    If nPermiso = Registrar Then
        If CboAutoriazaUsoDatos.ListIndex = -1 And CboAutoriazaUsoDatos.Visible Then
            MsgBox "Falta ingresar la autorización de datos", vbInformation, "Alerta"
            ValidaDatos = False
            Exit Function
        End If
    End If

'JOEP20190919 ERS042 CP-2018
'If Trim(Right(cmbDestCred.Text, 5)) <> 14 Then 'Comento JOEP2090307 Mejora
    nSubDestino = IIf(cmbSubDestCred.Visible = True And cmbSubDestCred.Text <> "", Trim(Right(cmbSubDestCred.Text, 8)), 0)
    nTpDoc = IIf(cmbTpDoc.Visible = True And cmbTpDoc.Text <> "" And nTpCmbTpDoc = 2, Trim(Right(cmbTpDoc, 10)), 0)
    nTpIngr = IIf(cmbTpDoc.Visible = True And cmbTpDoc.Text <> "" And nTpCmbTpDoc = 1, Trim(Right(cmbTpDoc, 10)), 0)
    nTpInt = IIf(cmbTpDoc.Visible = True And cmbTpDoc.Text <> "" And nTpCmbTpDoc = 3, Trim(Right(cmbTpDoc, 10)), 0)

If Trim(Right(cmbDestCred.Text, 5)) <> 14 Then 'JOEP2090307 Mejora
    If Not CatalogoValidador(4000, txtMontoSol.Text, spnCuotas.valor, spnPlazo.valor, CInt(Trim(Right(cmbDestCred.Text, 3))), nSubDestino, nTpDoc, nTpIngr, sT_Plani) Then
        spnCuotas.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Not CatalogoValidador(46000, txtMontoSol.Text, spnCuotas.valor, spnPlazo.valor, CInt(Trim(Right(cmbDestCred.Text, 3))), nSubDestino, nTpDoc, nTpIngr, sT_Plani) Then
        spnCuotas.SetFocus
        ValidaDatos = False
        Exit Function
    End If
End If

    If Not CP_Mensajes(5, Trim(Right(cmbSubProducto.Text, 5))) Then
        cmbDestCred.SetFocus
        ValidaDatos = False
        Exit Function
    End If
'JOEP20190919 ERS042 CP-2018
End Function

Private Sub CargaDatosTitular()
    lblCodigo.Caption = oPersona.PersCodigo
    lblNombre.Caption = oPersona.NombreCompleto
    lblDocnat.Caption = oPersona.ObtenerDNI
    lblDocTrib.Caption = oPersona.ObtenerRUC
End Sub

Private Sub LimpiaPantalla()
    cmdImprimir.Enabled = False
    Call LimpiaControles(Me, True)
    Call InicializaCombos(Me)
    
    If nIndtempo = -99 Then
    Else
        cmbAnalista.ListIndex = nIndtempo
        cmbAnalista.Enabled = False
    End If
    
    spnCuotas.valor = "0"
    spnPlazo.valor = "0"
    cmbFuentes.Clear
    cmbModular.Clear
    txtfechaAsig.Text = "__/__/____"
    ListaRelacion.ListItems.Clear
    ActXCtaCred.NroCuenta = ""
    ActXCtaCred.CMAC = gsCodCMAC
    ActXCtaCred.Age = gsCodAge
    ActXCtaCred.Enabled = True
    cmdGarantias.Enabled = False
    cmdGravar.Enabled = False
    cmdEvaluar.Enabled = False
    cmdEliminar.Enabled = False
    cmdEditar.Enabled = False
    If bRefinanciar Then
        ChkCap.Enabled = True
    End If
    
    ' CMACICA_CSTS - 10/11/2003 -------------------------------------------------------------
    
    If bSustituirDeudor Then
        ChkCap.Enabled = False
    End If
    
    ' ---------------------------------------------------------------------------------------
    ''WIOR 20120914 ******************************************
    'If nExisteAgeEvalCred = 0 Then
    '    Me.cmdEvaluar.Visible = False
    'Else
    '    Me.cmdEvaluar.Visible = True
    '    cmdEvaluar.Enabled = False
    'End If
    ''WIOR FIN ***********************************************
    'JUEZ 20130527 *************
    fbRegistraEnvio = False
    fbActivaEnvio = False
    Set frsEnvEstCta = Nothing
    fnModoEnvioEstCta = 0
    fnModoEnvioEstCtaSiNo = 0 'APRI20180215 ERS036-2017
    fnDebitoMismaCta = 0
    'END JUEZ ******************
    Set oCredAgrico = New frmCredAgricoSelec 'WIOR 20130723
    fbActivo = False 'WIOR 20130723
    chkAutAmpliacion.value = 0 'JUEZ 20160509
End Sub

Private Sub CargaFuentesIngreso(ByVal psPersCod As String)
Dim i As Integer
Dim MatFte As Variant

    On Error GoTo ErrorCargaFuentesIngreso
    Set oPersona = Nothing
    Set oPersona = New UPersona_Cli 'COMDPersona.DCOMPersona
    Call oPersona.RecuperaPersona_Solicitud(psPersCod)  'oPersona.RecuperaPersona(psPersCod)
    
    oPersona.PersCodigo = TitularCredito
'ARCV 29-12-2006
'    cmbFuentes.Clear
'    'For i = 0 To oPersona.NumeroFtesIngreso - 1
'    '    cmbFuentes.AddItem oPersona.ObtenerFteIngRazonSocial(i) & Space(100 - Len(oPersona.ObtenerFteIngRazonSocial(i))) & oPersona.ObtenerFteIngFuente(i) & Space(50 - Len(oPersona.ObtenerFteIngFuente(i))) & Format(oPersona.ObtenerFteIngFecEval(i), "dd/mm/yyyy hh:mm:ss")
'    'Next i
'
'    MatFte = oPersona.FiltraFuentesIngresoPorRazonSocial
'    If IsArray(MatFte) Then
'        ReDim MatTipoFte(UBound(MatFte))
'        For i = 0 To UBound(MatFte) - 1
'            'cmbFuentes.AddItem MatFte(i, 2) & Space(100 - Len(MatFte(i, 2))) & MatFte(i, 6) & Space(50 - Len(MatFte(i, 6))) & Format(MatFte(i, 4), "dd/mm/yyyy hh:mm:ss")
'            cmbFuentes.AddItem MatFte(i, 2) & Space(100 - Len(MatFte(i, 2))) & MatFte(i, 6) & Space(50 - Len(MatFte(i, 6))) & MatFte(i, 8)
'            MatTipoFte(i) = MatFte(i, 1)
'        Next i
'    End If
'
'    If cmbFuentes.ListCount > 0 Then
'        cmbFuentes.ListIndex = UBound(MatFte) - 1
'    End If
'------------
    Exit Sub

ErrorCargaFuentesIngreso:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Function TitularCredito() As String
    TitularCredito = ""
    oRelPersCred.IniciarMatriz
    Do While Not oRelPersCred.EOF
        If CInt(oRelPersCred.ObtenerValorRelac) = gColRelPersTitular Then
            TitularCredito = oRelPersCred.ObtenerCodigo
            Exit Do
        End If
        oRelPersCred.siguiente
    Loop
    
End Function
Private Sub HabilitaIngresoSolicitud(ByVal pbHabilita As Boolean)
    Call HabilitaControles(Me, pbHabilita)
    
    If nIndtempo = -99 Then
    Else
        cmbAnalista.Enabled = False
    End If
    
    fracontrol.Enabled = True
    cmdGrabar.Visible = pbHabilita
    cmdCancela.Visible = pbHabilita
    cmdNuevo.Visible = Not pbHabilita
    cmdEditar.Visible = Not pbHabilita
    cmdEliminar.Visible = False
    cmdGarantias.Visible = Not pbHabilita
    cmdGravar.Visible = Not pbHabilita
    cmdEvaluar.Visible = Not pbHabilita
    cmdsalir.Visible = Not pbHabilita
    CmdLimpiar.Visible = Not pbHabilita
    
    cmdGrabar.Enabled = pbHabilita
    cmdCancela.Enabled = pbHabilita
    cmdNuevo.Enabled = Not pbHabilita
    cmdEditar.Enabled = Not pbHabilita
    cmdEliminar.Enabled = Not pbHabilita
    cmdGarantias.Enabled = Not pbHabilita
    cmdGravar.Enabled = Not pbHabilita
    cmdEvaluar.Enabled = Not pbHabilita
    cmdsalir.Enabled = Not pbHabilita
    cmdFuentes.Enabled = pbHabilita
    cmdexaminar.Enabled = Not pbHabilita
    ActXCtaCred.Enabled = Not pbHabilita
    If pbHabilita Then
       txtfechaAsig = Format(gdFecSis, "dd/mm/yyyy")
    End If
    If bRefinanciar Then
       txtMontoSol.Enabled = False
    End If
    
    ' CMACICA_STS - 10112003 ------------------------------------------------------------
    If bSustituirDeudor Then
       txtMontoSol.Enabled = False
    End If
    '------------------------------------------------------------------------------------
    'AVMM -- 19122006 ------------
    'cmbCondicion.Enabled = True'Comento JOEP20180919 ERS042 CP-2018
    cmbCondicion.Enabled = False 'Agrego JOEP20180919 ERS042 CP-2018
    '-----------------------------
    'cmbCondicion.Enabled = False
    'MAVM 200907
    'CmbCondicionOtra.Enabled = False
    ' CMACICA_STS - 04112003 ------------------------------------------------------------
    'cmbCondicionOtra.Enabled = True
    '------------------------------------------------------------------------------------
    
    If bRefinanciar = True Then
        ChkCap.Enabled = pbHabilita
        ChkCap.Visible = True
    End If
    
    ' CMACICA_STS - 10112003 ------------------------------------------------------------
    If bSustituirDeudor = True Then
        ChkCap.Enabled = pbHabilita
        ChkCap.Visible = False
    End If
    '------------------------------------------------------------------------------------
    
    '->***** LUCV20171013, Agregó, según INC1710120009
    If fbRegPromotores Then
        fraPromotor.Enabled = pbHabilita
        cmbPromotor.Enabled = pbHabilita
    Else
        fraPromotor.Enabled = Not pbHabilita
        cmbPromotor.Enabled = Not pbHabilita
    End If
    '<-***** Fin LUCV20171013
    
'     If bAmpliacion = True Then
'        CmdAmpliacion.Enabled = True
'     End If
    ''WIOR 20120914 ******************************************
    'If nExisteAgeEvalCred = 0 Then
    '    Me.cmdEvaluar.Visible = False
    'Else
    '    Me.cmdEvaluar.Visible = Not pbHabilita
    '    cmdEvaluar.Enabled = Not pbHabilita
    'End If
    ''WIOR FIN ***********************************************
    cmdEnvioEstCta.Enabled = pbHabilita  'WIOR 20130713
    'JOEP20190204 CP
    If bRefinanciar = True Then
        cmbDestCred.Enabled = False
    End If
    'JOEP20190204 CP
End Sub
Private Sub HabilitaCreditoDsctoPlanilla(ByVal pbHabilita As Boolean)
    cmbInstitucion.Enabled = pbHabilita
    cmbModular.Enabled = pbHabilita
    lblinstitucion.Enabled = pbHabilita
    lblModular.Enabled = pbHabilita
    cmbInstitucion.Visible = pbHabilita
    cmbInstitucion.ListIndex = -1
    cmbModular.Visible = pbHabilita
    cmbModular.Text = ""
    lblinstitucion.Visible = pbHabilita
    lblModular.Visible = pbHabilita
    
    '**DAOR 20070807*********************
    If pbHabilita = False Then
        sCodInstitucion = ""
        sCodModular = ""
    End If
    '************************************
End Sub

Private Sub ActualizarListaPersRelacCred()
Dim s As ListItem
Dim oPREDA As COMDPersona.DCOMPersonas 'JUEZ 20130717
    ListaRelacion.ListItems.Clear
    oRelPersCred.IniciarMatriz
    Do While Not oRelPersCred.EOF
        'JUEZ 20130717 ********************************************************
        Set oPREDA = New COMDPersona.DCOMPersonas
        If oPREDA.VerificarPersonaPREDA(oRelPersCred.ObtenerCodigo, 1) Then
            MsgBox "El " & IIf(oRelPersCred.ObtenerValorRelac = CStr(gColRelPersTitular), "Titular", "cliente") & " " & oRelPersCred.ObtenerNombre & " es un cliente PREDA no sujeto de Crédito, consultar a Coordinador de Producto Agropecuario", vbInformation, "Aviso"
            If oRelPersCred.ObtenerValorRelac = CStr(gColRelPersTitular) Then
                Do While Not oRelPersCred.EOF
                    Call oRelPersCred.EliminarRelacion(oRelPersCred.ObtenerCodigo, oRelPersCred.ObtenerValorRelac)
                Loop
                Exit Sub
            End If
            Call oRelPersCred.EliminarRelacion(oRelPersCred.ObtenerCodigo, oRelPersCred.ObtenerValorRelac)
        Else
        'END JUEZ *************************************************************
            Set s = ListaRelacion.ListItems.Add(, , oRelPersCred.ObtenerNombre)
            s.SubItems(1) = oRelPersCred.ObtenerRelac
            s.SubItems(2) = oRelPersCred.ObtenerCodigo
            s.SubItems(3) = oRelPersCred.ObtenerValorRelac
            '***Modificado por ELRO 20110831, según Acta 222-2011/TI-D
            If oRelPersCred.ObtenerValorRelac = CStr(gColRelPersTitular) Then
                PersoneriaTitular = CInt(oRelPersCred.ObtenerValorPersoneria)
            End If
            '*********************************************************
        End If
        oRelPersCred.siguiente
    Loop
End Sub

Private Sub CargaPersonasRelacCred(ByVal psCtaCod As String, _
                                    ByVal prsRelac As ADODB.Recordset)
Dim s As ListItem
    On Error GoTo ErrorCargaPersonasRelacCred
    ListaRelacion.ListItems.Clear
    Set oRelPersCred = New UCredRelac_Cli
    Call oRelPersCred.CargaRelacPersCred(psCtaCod, prsRelac)
    oRelPersCred.IniciarMatriz
    Do While Not oRelPersCred.EOF
        Set s = ListaRelacion.ListItems.Add(, , oRelPersCred.ObtenerNombre)
        s.SubItems(1) = oRelPersCred.ObtenerRelac
        s.SubItems(2) = oRelPersCred.ObtenerCodigo
        s.SubItems(3) = oRelPersCred.ObtenerValorRelac
        '***Modificado por ELRO 20111017, según Acta 222-2011/TI-D
        If oRelPersCred.ObtenerValorRelac = CStr(gColRelPersTitular) Then
            PersoneriaTitular = CInt(oRelPersCred.ObtenerValorPersoneria)
        End If
        '*********************************************************
        oRelPersCred.siguiente
    Loop
    
    Exit Sub

ErrorCargaPersonasRelacCred:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

'Private Sub CargaRefinanciados(ByRef MatCalend As Variant)
'Dim oNegCredito As COMNCredito.NCOMCredito
'Dim MatCalendTemp As Variant
'Dim i As Integer
'
'    On Error GoTo ErrorCargaRefinanciados
'    If bRefinanciar Or bSustituirDeudor Then
'        Set oNegCredito = New COMNCredito.NCOMCredito
'        MatCalendTemp = oNegCredito.RecuperaMatrizRefinanciados(ActXCtaCred.NroCuenta)
'        Set oNegCredito = Nothing
'        ReDim MatCalend(UBound(MatCalendTemp), 9)
'        For i = 0 To UBound(MatCalendTemp) - 1
'            MatCalend(i, 0) = MatCalendTemp(i, 1)
'            MatCalend(i, 1) = MatCalendTemp(i, 2)
'            MatCalend(i, 2) = MatCalendTemp(i, 3) 'Capital
'            MatCalend(i, 3) = MatCalendTemp(i, 5) 'Int Comp
'            MatCalend(i, 4) = MatCalendTemp(i, 9) 'Int Moratorio
'            MatCalend(i, 5) = MatCalendTemp(i, 7) 'Int Gracia
'            MatCalend(i, 6) = MatCalendTemp(i, 13) 'Int Suspenso
'            MatCalend(i, 7) = MatCalendTemp(i, 11) 'Int Reprog
'            MatCalend(i, 8) = MatCalendTemp(i, 15) 'gastos
'            MatCalend(i, 9) = MatCalendTemp(i, 2)
'        Next i
'    End If
'    Exit Sub
'
'ErrorCargaRefinanciados:
'        MsgBox Err.Description, vbCritical, "Aviso"
'
'
'End Sub

Private Sub CargaDatos(ByVal psCtaCod As String)

Dim oCredito As COMNCredito.NCOMCredito   ' COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim R2 As ADODB.Recordset
Dim i As Integer
Dim nIndice As Integer
Dim rsRelac As ADODB.Recordset
Dim rsAporte As ADODB.Recordset 'JOEP20181201 CP

    Set oRelPersCred = Nothing
    Set oPersona = Nothing
    Set oCredito = New COMNCredito.NCOMCredito ' COMDCredito.DCOMCredito
    
    'Call oCredito.CargaObjetosControles(psCtaCod, bRefinanciar, bSustituirDeudor, R, R2, MatCredRef, MatCredSust, rsRelac)
    'Call oCredito.CargaObjetosControles(psCtaCod, bRefinanciar, bSustituirDeudor, R, R2, MatCredRef, MatCredSust, rsRelac, fvListaCompraDeuda, fvListaAguaSaneamiento) 'EJVG20160203 ERS002-2016 'EAAS 20180807 SEGUN ERS054-2018 fvListaAguaSaneamiento
    Call oCredito.CargaObjetosControles(psCtaCod, bRefinanciar, bSustituirDeudor, R, R2, MatCredRef, MatCredSust, rsRelac, fvListaCompraDeuda, fvListaAguaSaneamiento, rsAporte, fvListaCreditoVerde) 'JOEP20181201 CP rsAporte 'EAAS20191401 fvListaCreditoVerde SEGUN 018-GM-DI_CMACM
    'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM
     If UBound(fvListaAguaSaneamiento) >= 1 Then
        nSumaAguaSaneamiento = fvListaAguaSaneamiento(1).nSumaAguaSaneamiento
     End If
     If UBound(fvListaCreditoVerde) >= 1 Then
        nSumaCreditoVerde = fvListaCreditoVerde(1).nSumaCreditoVerde
     End If
    'FIN EAAS20191401 SEGUN 018-GM-DI_CMACM
    'Carga Personas Relacionadas con el Credito
    Call CargaPersonasRelacCred(psCtaCod, rsRelac)
    
    fnMontoExpEsteCred_NEW = 0# 'EJVG20160712
    If Not R.BOF And Not R.EOF Then
        'Carga Condicion del Credito
        nIndice = IndiceListaCombo(cmbCondicion, IIf(IsNull(R!nColocCondicionProd), 0, R!nColocCondicionProd))
        If nIndice <> -1 Then
            LblCondProd.Caption = cmbCondicion.List(nIndice)
        End If
        cmbCondicion.ListIndex = IndiceListaCombo(cmbCondicion, R!nColocCondicion)
        
        ' STS - 04112003 --------------------------------------------------------------------------
        'ALPA-MAVM 20090801********************************************
        'Carga Condicion 2 del Credito
        'cmbCondicionOtra.ListIndex = IndiceListaCombo(cmbCondicionOtra, R!nColocCondicion2)
        cmbCondicionOtra.ListIndex = IndiceListaCombo(cmbCondicionOtra, R!idCampana)
        '**************************************************************
        cmbCondicionOtra2.ListIndex = IndiceListaCombo(cmbCondicionOtra2, IIf(IsNull(R!id_CasaCom), 0, R!id_CasaCom))
        ' -----------------------------------------------------------------------------------------
        
        lblCodigo.Caption = R!cPersCod
        lblNombre.Caption = R!cPersNombre
        lblDocnat.Caption = IIf(IsNull(R!Dni), "", R!Dni)
        lblDocTrib.Caption = IIf(IsNull(R!Ruc), "", R!Ruc)
        Me.txtVendedor = R!cVendedor
'        cmbTipoCred.ListIndex = IndiceListaCombo(cmbTipoCred, Mid(psCtaCod, 6, 1) & "00")
 '       cmbSubTipo.ListIndex = IndiceListaCombo(cmbSubTipo, Mid(psCtaCod, 6, 1) & Mid(psCtaCod, 7, 2))
        'cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, Mid(psCtaCod, 9, 1))'Comento JOEP20190919 ERS042 CP-2018
        'Comento JOEP20190919 ERS042 CP-2018
'        If Mid(psCtaCod, 9, 1) = "1" Then
'            txtMontoSol.ForeColor = vbBlue
'        Else
'            txtMontoSol.ForeColor = &H289556
'        End If
        'Comento JOEP20190919 ERS042 CP-2018
        txtMontoSol.Text = Format(R!nMonto, "#0.00")
        spnCuotas.valor = Format(R!nCuotas, "#0")
        spnPlazo.valor = Format(R!nPlazo, "#0")
        'cmbDestCred.ListIndex = IndiceListaCombo(cmbDestCred, Trim(str(CInt(R!nColocDestino))))'Comento JOEP20190919 ERS042 CP-2018
        cmbAnalista.ListIndex = IndiceListaCombo(cmbAnalista, IIf(IsNull(R!cAnalista), "", R!cAnalista))
        'WIOR 20140510 *********************
        If fbRegPromotores Then
           'fraPromotor.Visible = True 'LUCV20171013, Comentó, según INC1710120009
            fraPromotor.Enabled = True 'LUCV20171013, Agregó, según INC1710120009
            cmbPromotor.ListIndex = IndiceListaCombo(cmbPromotor, Trim(R!cPromotor))
        Else
            cmbPromotor.ListIndex = IndiceListaCombo(cmbPromotor, Trim(R!cPromotor))
            fraPromotor.Enabled = False 'LUCV20171013, Agregó, según INC1710120009
            'fraPromotor.Visible = False 'LUCV20171013, Comentó, según INC1710120009
        End If
        'WIOR FIN **************************
        txtfechaAsig.Text = Format(R!dPrdEstado, "dd/mm/yyyy")
        ChkCap.value = IIf(R!bRefCapInt, 1, 0)
        cmdImprimir.Enabled = True
        
        
        'Carga Fuentes de Ingreso
        Call CargaFuentesIngreso(TitularCredito)
        For i = 0 To cmbFuentes.ListCount - 1
            If R!cNumFuente = Trim(Mid(cmbFuentes.List(i), 150, 9)) Then
                cmbFuentes.ListIndex = i
            End If
        Next i
        
        'Carga Refinanciados
        'If bRefinanciar Then
        '   Call CargaRefinanciados(MatCredRef)
        'Else
        '    If bSustituirDeudor Then
        '       Call CargaRefinanciados(MatCredSust)
        '    End If
        'End If
        
        ChkCap.Enabled = False
        
        'cmbProductoCMACM.ListIndex = IndiceListaCombo(cmbProductoCMACM, Mid(psCtaCod, 6, 1) & "00")
        cmbProductoCMACM.ListIndex = IndiceListaCombo(cmbProductoCMACM, Mid(R!cTpoProdCod, 1, 1) & "00") 'EJVG20160714
        'cmbSubProducto.ListIndex = IndiceListaCombo(cmbSubProducto, Mid(psCtaCod, 6, 1) & Mid(psCtaCod, 7, 2))
        cmbSubProducto.ListIndex = IndiceListaCombo(cmbSubProducto, R!cTpoProdCod) 'EJVG20160714
         'WIOR 20130723 ******************************
        
        'Agrego JOEP20190919 ERS042 CP-2018
        If cmbSubProducto.Text <> "" Then
            If bRefinanciar = True Then
                cmbDestCred.ListIndex = IndiceListaCombo(cmbDestCred, Trim(str(CInt(R!nColocDestino))))
            Else
                Call CatalogoLlenaCombox(Trim(Right(cmbSubProducto.Text, 3)), 2000, Trim(str(CInt(R!nColocDestino))))
            End If
        End If
        
        Call CP_DatosDefaut(46000)
        cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, Mid(psCtaCod, 9, 1)) 'Cambio de Posicion JOEP20190919 ERS042 CP-2018
        Call PintaMoneda(Mid(psCtaCod, 9, 1))
        If cmbTpDoc.Visible = True Then
            cmbTpDoc.ListIndex = IndiceListaCombo(cmbTpDoc, IIf(R!nTpDoc = -1, IIf(R!nTpIngr = -1, R!nTpInt, R!nTpIngr), R!nTpDoc))
        End If
        
        If cmbSubDestCred.Visible = True Then
            cmbSubDestCred.ListIndex = IndiceListaCombo(cmbSubDestCred, R!nSubDestino)
        End If
        
        If Not (rsAporte.BOF And rsAporte.EOF) Then
            If Trim(Right(cmbSubProducto.Text, 3)) = "703" Then
                ReDim nMatMontoPre(1, 4)
                nMatMontoPre(1, 1) = rsAporte!nMonto
                nMatMontoPre(1, 2) = rsAporte!nAporte
                'nMatMontoPre(1, 3) = rsAporte!nMontoDisponible'Joep20190313 Comento
                nMatMontoPre(1, 3) = rsAporte!nMontoSoli 'Joep20190313
                'nMatMontoPre(1, 4) = rsAporte!nMontoSoli 'Joep20190313 Comento
                nMatMontoPre(1, 4) = rsAporte!nMontoDisponible 'Joep20190313
            Else
                ReDim nMatMontoPre(1, 3)
                nMatMontoPre(1, 1) = rsAporte!nMonto
                nMatMontoPre(1, 2) = rsAporte!nAporte
                nMatMontoPre(1, 3) = rsAporte!nMontoSoli
            End If
        End If
        'Agrego JOEP20190919 ERS042 CP-2018
         
        Set oCredAgrico = New frmCredAgricoSelec
        If Trim(Right(cmbProductoCMACM.Text, 5)) = "600" Then
            oCredAgrico.inicia psCtaCod, Trim(Right(cmbSubProducto.Text, 5))
        End If
        'WIOR FIN ***********************************
        'Carga Codigos Modulares de Titular
        'Set R2 = oCredito.CodigosModulares(oRelPersCred.TitularPersCod)
        cmbModular.Clear
        Do While Not R2.EOF
            cmbModular.AddItem Trim(R2!cCodModular)
            R2.MoveNext
        Loop
        R2.Close
        cmbInstitucion.ListIndex = IndiceListaCombo(cmbInstitucion, IIf(IsNull(R!cPersConvenio), "", R!cPersConvenio))
        cmbModular.ListIndex = IndiceListaCombo(cmbModular, IIf(IsNull(R!cCodModular), "", R!cCodModular), 0)
        If bAmpliacion Then chkAutAmpliacion.value = R!bAutSolicAmp 'JUEZ 20160509
        fnMontoExpEsteCred_NEW = R!nMontoExpCredito 'EJVG20160713
    Else
        cmdImprimir.Enabled = False
    End If
    R.Close
    Set R = Nothing
    Set R2 = Nothing
    Set oCredito = Nothing
    fbActivo = True 'WIOR 20130723
    'add pti1 ers070-2018 18/12/2018 ****************
    Call ActAutDatos
  
    
End Sub

'Private Sub CargaAnalistas()
'Dim R As ADODB.Recordset
'Dim ssql As String
'Dim oconecta As COMConecta.DCOMConecta
'Dim sAnalistas As String
'Dim oGen As COMDConstSistema.DCOMGeneral
'Dim bMuestraSoloAnalistaActual As Integer
'
'
'    On Error GoTo ERRORCargaAnalistas
'
'    Set oGen = New COMDConstSistema.DCOMGeneral
'    sAnalistas = oGen.LeeConstSistema(gConstSistRHCargoCodAnalistas)
'    bMuestraSoloAnalistaActual = oGen.LeeConstSistema(58)
'    Set oGen = Nothing
'
'    ssql = "Select R.cPersCod, P.cPersNombre from RRHH R inner join Persona P ON R.cPersCod = P.cpersCod "
'    ssql = ssql & " AND nRHEstado in (201,301) "
'    ssql = ssql & " inner join RHCargos RC ON R.cPersCod = RC.cPersCod "
'    ssql = ssql & " where  RC.cRHCargoCod in (" & sAnalistas & ") AND RC.dRHCargoFecha = (select MAX(dRHCargoFecha) from RHCargos RHC2 where RHC2.cPersCod = RC.cPersCod) "
'    ssql = ssql & " and R.cAgenciaActual='" & gsCodAge & "'"
'    ssql = ssql & " order by P.cPersNombre "
'
'
'    Set oconecta = New COMConecta.DCOMConecta
'    oconecta.AbreConexion
'    Set R = oconecta.CargaRecordSet(ssql)
'    oconecta.CierraConexion
'    Set oconecta = Nothing
'
'    nIndtempo = -99
'    cmbAnalista.Clear
'    Do While Not R.EOF
'
'        If bMuestraSoloAnalistaActual = 1 Then
'            If R!cPersCod = gsCodPersUser Then
'                nIndtempo = R.AbsolutePosition - 1
'            End If
'        End If
'
'        cmbAnalista.AddItem PstaNombre(R!cPersNombre) & Space(100) & R!cPersCod
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'
'    If bMuestraSoloAnalistaActual = 1 Then
'        cmbAnalista.Enabled = False
'        cmbAnalista.ListIndex = nIndtempo
'    End If
'
'
'    Exit Sub
'ERRORCargaAnalistas:
'    MsgBox Err.Description, vbCritical, "Aviso"
'End Sub

'Private Sub CargaSubProducto(ByVal psTipo As String)'Comento JOEP20190118 CP
Private Sub CargaSubProducto(ByVal psTipo As String, Optional ByVal nOpeRefinanciado As Boolean = False, Optional ByVal nOpeAmpliado As Boolean = False) 'JOEP20190118 CP
Dim oCred As COMDCredito.DCOMCredito
Dim ssql As String
Dim RTemp As ADODB.Recordset
    On Error GoTo ERRORCargaSubProducto
    Set oCred = New COMDCredito.DCOMCredito
    'Set RTemp = oCred.RecuperaSubProductosCrediticios(psTipo, gsCodCargo) 'NAGL 20171121'Cometno JOEP20190118 CP
    Set RTemp = oCred.RecuperaSubProductosCrediticios(psTipo, gsCodCargo, nOpeRefinanciado, nOpeAmpliado) 'JOEP20190118 CP
    Set oCred = Nothing
    cmbSubProducto.Clear
    
    If cmdEjecutar = 1 Then cmbSubProducto.Enabled = True 'Agrego JOEP20190919 ERS042 CP-2018
    
    Do While Not RTemp.EOF
        cmbSubProducto.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    Call CambiaTamañoCombo(cmbSubProducto, 250)
    'cmbMoneda.Enabled = True 'WIOR 20151222
    If Not bAmpliacion Or rsAmpliado Is Nothing Then cmbMoneda.Enabled = True    'JUEZ 20160509
    Exit Sub
    
ERRORCargaSubProducto:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

'Private Sub CargaTiposCredito(ByVal psTipoFuente As String)
'Dim oCred As COMDCredito.DCOMCredito
'Dim RTemp As ADODB.Recordset
'
'    Set oCred = New COMDCredito.DCOMCredito
'    Set RTemp = oCred.RecuperaProductosDeSolicitudDeCredito
'    Set oCred = Nothing
'    cmbTipoCred.Clear
'
'    Do While Not RTemp.EOF
'        If psTipoFuente = "D" And (Mid(Trim(Str(RTemp!nConsValor)), 1, 1) = "3" Or Mid(Trim(Str(RTemp!nConsValor)), 1, 1) = "4") Then 'Fte Dependiente
'            cmbTipoCred.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
'        End If
'        If psTipoFuente = "I" And (Mid(Trim(Str(RTemp!nConsValor)), 1, 1) = "1" Or Mid(Trim(Str(RTemp!nConsValor)), 1, 1) = "2") Then 'Fte Independiente
'            cmbTipoCred.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
'        End If
'        RTemp.MoveNext
'    Loop
'    RTemp.Close
'
'    Set RTemp = Nothing
'    If cmbTipoCred.ListCount > 0 Then
'        cmbTipoCred.ListIndex = 0
'    End If
'    Call CambiaTamañoCombo(cmbTipoCred, 300)
'End Sub

Private Sub CargaControles()

    Call Cargar_Objetos_Controles
'Dim oconecta As COMConecta.DCOMConecta
'Dim ssql As String
'Dim RTemp As ADODB.Recordset
'Dim oCred As COMDCredito.DCOMCredito
'
'    On Error GoTo ERRORCargaControles
'
'    'Carga Condiciones de un Credito
'    Call CargaComboConstante(gColocCredCondicion, cmbCondicion)
'    Call CambiaTamañoCombo(cmbCondicion)
'
'    'STS - 04112003 ------------------------------------------------------------------
'    'Carga Condiciones 2 de un Credito
'    Call CargaComboConstante(gColocCredCondicionOtra, cmbCondicionOtra)
'    Call CambiaTamañoCombo(cmbCondicionOtra)
'    '---------------------------------------------------------------------------------
'
'    'Carga Tipos de Credito
'    Set oCred = New COMDCredito.DCOMCredito
'    Set RTemp = oCred.RecuperaProductosDeSolicitudDeCredito
'    Set oCred = Nothing
'    cmbTipoCred.Clear
'    Do While Not RTemp.EOF
'        cmbTipoCred.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
'        RTemp.MoveNext
'    Loop
'    RTemp.Close
'    Set RTemp = Nothing
'    If cmbTipoCred.ListCount > 0 Then
'        cmbTipoCred.ListIndex = 0
'    End If
'    Call CambiaTamañoCombo(cmbTipoCred, 300)
'    'Carga Monedas
'    Call CargaComboConstante(gMoneda, cmbMoneda)
'    'Carga Destino de Credito
'    Call CargaComboConstante(gColocDestino, cmbDestCred)
'
'    'Carga Analistas
'    Call CargaAnalistas
'    'Carga Instituciones
'    Call CargaComboPersonasTipo(gPersTipoConvenio, cmbInstitucion)
'    Call CambiaTamañoCombo(cmbInstitucion, 400)
'    spnCuotas.Valor = 0
'    spnPlazo.Valor = 0
'    Exit Sub
'
'
'ERRORCargaControles:
'    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub Cargar_Objetos_Controles()

Dim oCred As COMDCredito.DCOMCredito
Dim rsCondCred As ADODB.Recordset
Dim rsCondCred2 As ADODB.Recordset
Dim rsTipoCred As ADODB.Recordset
Dim rsDestCred As ADODB.Recordset
Dim rsAnalista As ADODB.Recordset
Dim rsMoneda As ADODB.Recordset
'Dim rsInstituc As ADODB.Recordset
Dim bMuestraSoloAnalistaActual As Integer
'By Capi 28102008
Dim rsMotivoRef As ADODB.Recordset
'
Dim rsPromotores As ADODB.Recordset 'WIOR 20140509

   On Error GoTo ErrorCargar_Objetos_Controles
    
   Set oCred = New COMDCredito.DCOMCredito
   
   'By Capi 28102008
   'Call oCred.Cargar_Objetos_Controles(rsCondCred, rsCondCred2, rsTipoCred, rsDestCred, rsAnalista, bMuestraSoloAnalistaActual, rsMoneda, rsInstituc, gsCodAge)
   Call oCred.Cargar_Objetos_Controles(rsCondCred, rsCondCred2, rsTipoCred, rsDestCred, rsAnalista, bMuestraSoloAnalistaActual, rsMoneda, rsInstituc, gsCodAge, rsMotivoRef, rsPromotores) 'WIOR 2014509 AGREGO rsPromotores
   '
   Set oCred = Nothing
   'Carga Condiciones de un Credito
   Call Llenar_Combo_con_Recordset(rsCondCred, cmbCondicion)
   Call CambiaTamañoCombo(cmbCondicion)
    
   'Carga Condiciones 2 de un Credito/ Campañas
   Call Llenar_Combo_con_Recordset(rsCondCred2, cmbCondicionOtra)
   Call CambiaTamañoCombo(cmbCondicionOtra)
    
   'Carga Tipos de Credito
   Call Llenar_Combo_con_Recordset(rsTipoCred, cmbProductoCMACM)
   If cmbProductoCMACM.ListCount > 0 Then
        'cmbProductoCMACM.ListIndex = 0 'Comento JOEP20190919 ERS042 CP-2018
        cmbProductoCMACM.ListIndex = -1 'Agrego JOEP20190919 ERS042 CP-2018
   End If
   Call CambiaTamañoCombo(cmbProductoCMACM, 300)
    
   'Carga Monedas
   Call Llenar_Combo_con_Recordset(rsMoneda, cmbMoneda)
   
   'Carga Destino de Credito
   Call Llenar_Combo_con_Recordset(rsDestCred, cmbDestCred)
   'ALPA 20150414*****************************************
   Call CambiaTamañoCombo(cmbDestCred, 320)
   '******************************************************
   'Carga Analistas
   
   'By Capi 28102008 carga motivos de refinanciacion
   Call Llenar_Combo_con_Recordset(rsMotivoRef, cmbMotivoRef)
   '
    
    'Carga Analistas
    nIndtempo = -99
    cmbAnalista.Clear
    
    Do While Not rsAnalista.EOF
        If bMuestraSoloAnalistaActual = 1 Then
            If rsAnalista!cPersCod = gsCodPersUser Then
                nIndtempo = rsAnalista.AbsolutePosition - 1
            End If
        End If
        cmbAnalista.AddItem PstaNombre(rsAnalista!cPersNombre) & Space(100) & rsAnalista!cPersCod
        rsAnalista.MoveNext
    Loop
    
    If bMuestraSoloAnalistaActual = 1 Then
        cmbAnalista.Enabled = False
        cmbAnalista.ListIndex = nIndtempo
    End If

    'WIOR 20140509 ****************************
    If fbRegPromotores Then
        'fraPromotor.Visible = True 'LUCV20171013, Comentó, según INC1710120009
        fraPromotor.Enabled = True 'LUCV20171013, Agregó, según INC1710120009
        cmbPromotor.Clear
        Do While Not rsPromotores.EOF
            cmbPromotor.AddItem PstaNombre(rsPromotores!cPersNombre) & Space(100) & rsPromotores!cPersCod
            rsPromotores.MoveNext
        Loop
    Else
        'fraPromotor.Visible = False 'LUCV20171013, Comentó, según INC1710120009
        fraPromotor.Enabled = False 'LUCV20171013, Agregó, según INC1710120009
        cmbPromotor.Clear
        Do While Not rsPromotores.EOF
            cmbPromotor.AddItem PstaNombre(rsPromotores!cPersNombre) & Space(100) & rsPromotores!cPersCod
            rsPromotores.MoveNext
        Loop
        
        
    End If
    'WIOR FIN *********************************
    
    'Carga Instituciones
    'Call Llenar_Combo_con_Recordset(rsInstituc, cmbInstitucion)
'ARCV 28-10-2006
'    cmbInstitucion.Clear
'    Do While Not rsInstituc.EOF
'        cmbInstitucion.AddItem PstaNombre(rsInstituc!cPersNombre) & Space(250) & rsInstituc!cPersCod
'        rsInstituc.MoveNext
'    Loop
'    Call CambiaTamañoCombo(cmbInstitucion, 400)
    
    spnCuotas.valor = 0
    spnPlazo.valor = 0
    
    Set rsCondCred = Nothing
    Set rsCondCred2 = Nothing
    Set rsTipoCred = Nothing
    Set rsDestCred = Nothing
    Set rsAnalista = Nothing
    Set rsMoneda = Nothing
    'Set rsInstituc = Nothing
    'By Capi 28102008
    Set rsMotivoRef = Nothing
    '
    Set rsPromotores = Nothing 'WIOR 20140509
    
    Exit Sub
    
ErrorCargar_Objetos_Controles:
    MsgBox Err.Description, vbInformation, "Aviso"

End Sub

Private Sub ActXCtaCred_KeyPress(KeyAscii As Integer)
Dim sCadTmp As Integer
Dim oNeg As COMNCredito.NCOMCredito

     If KeyAscii = 13 Then
        Set oNeg = New COMNCredito.NCOMCredito
        ActXCtaCred.Enabled = False
        'If bRefinanciar Then
         If bRefinanciar Or bSustituirDeudor Then
            If Not oNeg.EsRefinanciado(ActXCtaCred.NroCuenta) Then
                MsgBox "Credito No Es una Refinanciacion", vbInformation, "Aviso"
                Exit Sub
            End If
        ElseIf bAmpliacion Then 'LUCV20180417, Agregó
            If oNeg.EsRefinanciado(ActXCtaCred.NroCuenta) Then
                MsgBox "El crédito no se solicitó como Ampliado. Es Refinanciado.", vbInformation, "Aviso"
                Exit Sub
            End If
            '->***** LUCV20180417, Agregó
            If Not oNeg.EsAmpliado(ActXCtaCred.NroCuenta) Then
                MsgBox "El crédito no se solicitó como Ampliado.", vbInformation, "Aviso"
                Exit Sub
            End If
            '<-***** Fin LUCV20180417
        Else
            If oNeg.EsRefinanciado(ActXCtaCred.NroCuenta) Then
                MsgBox "Credito No Es una Solicitud Normal", vbInformation, "Aviso"
                Exit Sub
            End If
            '->***** LUCV20180417, Agregó
            If oNeg.EsAmpliado(ActXCtaCred.NroCuenta) Then
                MsgBox "El crédito no es una Solicitud Normal. Es Ampliado.", vbInformation, "Aviso"
                Exit Sub
            End If
            '<-***** Fin LUCV20180417
        End If
        Set oNeg = Nothing
        Call CargaDatos(ActXCtaCred.NroCuenta)
        cmdEditar.Enabled = True
        cmdEliminar.Enabled = True
        cmdGarantias.Enabled = True
        cmdGravar.Enabled = True
        cmdEvaluar.Enabled = True
        If lblCodigo.Caption = "" Then
            MsgBox "Cuenta No Existe ", vbInformation, "Aviso"
            cmdEjecutar = 1
            cmdCancela_Click
        End If
        Call ControlesPermiso
     End If
End Sub


Private Sub cmbAnalista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbInstitucion.Visible Then
            cmbInstitucion.SetFocus
        Else
            cmdGrabar.SetFocus
        End If
    End If
End Sub

Private Sub CmbCondicion_Change()
If Not bAmpliacion Then
    If InStr(1, cmbCondicion.Text, "AMPLIADO") > 0 Then
        MsgBox "No se puede escoger esta condicion de crédito", vbInformation, "Mensaje"
        cmbCondicion.ListIndex = 0
    End If
End If

If Not bRefinanciar Then
    If InStr(1, cmbCondicion.Text, "REFINANCIADO") > 0 Then
        MsgBox "No se puede escoger esta condicion de crédito", vbInformation, "Mensaje"
        cmbCondicion.ListIndex = 0
    End If
End If

Call DefineCondicionCredito 'MAVM 20110613

End Sub

Private Sub CmbCondicion_Click()
    Call CmbCondicion_Change
End Sub

Private Sub cmbCondicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'cmbCondicionOtra.SetFocus
    End If
End Sub

Private Sub cmbCondicionOtra_Click()
'JOEP20190114 CP
Set nMatMontoPre = Nothing
txtMontoSol.Text = "0.00"
'JOEP20190114 CP

    If Right(cmbCondicionOtra.Text, 1) = "2" Then 'Cuando la condicion es campaña
        LblCampana.Visible = True
        Call FrmCredListaCampanas.Inicio(gsCodAge, ActXCtaCred.NroCuenta)
        'Call ObtenerDescripcionCampana(nCampanaCod)
        LblCampana.Caption = cCampanaDesc
    Else
        LblCampana.Visible = False
        'JOEP20190114 CP
        If bRefinanciar = True Then
            cmbCondicionOtra.ListIndex = IndiceListaCombo(cmbCondicionOtra, 0)
            cmbCondicionOtra.Enabled = False
        End If
        'JOEP20190114 CP
    End If
    
     'MADM 20100719 ***************
    Call CargaSubCampanaConv(Trim(Right(cmbCondicionOtra.Text, 3)))
    '***********************************
    txtMontoSol.Enabled = True 'JOEP20190114 CP
End Sub

Private Sub CargaSubCampanaConv(ByVal psTipo As String)
Dim oCred As COMDCredito.DCOMCreditos
Dim ssql As String
Dim RTemp As ADODB.Recordset
    On Error GoTo ERRORCargaSubCampanaConv
    Set oCred = New COMDCredito.DCOMCreditos
    Set RTemp = oCred.DevolverDatos_CasaComercialID(IIf(psTipo = "", 0, psTipo))
    Set oCred = Nothing
    cmbCondicionOtra2.Clear
    Do While Not RTemp.EOF
        cmbCondicionOtra2.AddItem RTemp!cNombre & Space(200) & RTemp!Id
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    Call CambiaTamañoCombo(cmbCondicionOtra2, 200)
   
    If cmbCondicionOtra2.ListCount > 0 Then
        cmbCondicionOtra2.ListIndex = 0
    End If
    Exit Sub
ERRORCargaSubCampanaConv:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

'Sub ObtenerDescripcionCampana(ByVal pnIdCampana As Integer)
'    Dim odCamp As COMDCredito.DCOMCampanas
'
'    Set odCamp = New COMDCredito.DCOMCampanas
'    LblCampana.Caption = odCamp.DesCampanaXIdCampana(pnIdCampana)
'    Set odCamp = Nothing
'End Sub
Private Sub cmbCondicionOtra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdRelaciones.SetFocus
    End If
End Sub

'MIOL 20130625, SEGUN RQ13335 ********************************
Private Sub cmbCondicionOtra2_Click()
    If Trim(Right(cmbCondicionOtra2.Text, 3)) = "0" Then
        Me.lblVendedor.Visible = False
        Me.txtVendedor.Visible = False
    End If
    Set objProducto = New COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018
    If cmbSubProducto.Text <> "" Then
    If objProducto.GetResultadoCondicionCatalogo("N0000065", Trim(Right(cmbSubProducto.Text, 10))) Then '**ARLO20180712 ERS042 - 2018
        'If CInt(Trim(Right(cmbSubProducto.Text, 10))) = 510 Or CInt(Trim(Right(cmbSubProducto.Text, 10))) = 706 Or CInt(Trim(Right(cmbSubProducto.Text, 10))) = 511 Then
            If Trim(Right(cmbCondicionOtra2.Text, 3)) <> "0" Then
                Me.lblVendedor.Visible = True
                Me.txtVendedor.Visible = True
            End If
        End If
    End If
End Sub
'END MIOL ****************************************************

'EJVG20160201 ERS002-2016***
Private Sub cmbDestCred_Click()
    Dim rsSubDestinos As ADODB.Recordset 'JOEP20190919 ERS042 CP-2018
    Dim P As COMDCredito.DCOMCredito 'JOEP20190919 ERS042 CP-2018
    Dim n As Double 'JOEP20190919 ERS042 CP-2018
    
    'EAAS 20180727 ERS-054-2018
    Dim nClick As Integer
    cmdDestinoDetalle.Visible = False
    cmdDestinoDetalleAguaS.Visible = False
    cmdCreditoVerde.Visible = False 'EAAS20191401 SEGUN 018-GM-DI_CMACM
If bRefinanciar = False Then 'JOEP20190919 ERS042 CP-2018
    Select Case val(Trim(Right(cmbDestCred.Text, 3)))
        Case ColocDestino.gColocDestinoCambEstructPasivo:
            cmdDestinoDetalle.Visible = True
            'INICIO EAAS 20180727 ERS-054-2018
            cmdDestinoDetalleAguaS.Visible = False
            If cmbDestCred.ListCount > 1 Then
                Set nMatMontoPre = Nothing 'JOEP20190919 ERS042 CP-2018
            End If
        Case ColocDestino.gColocDestinoAguaSaneamiento:
            cmdDestinoDetalleAguaS.Visible = True
            If (cmdEjecutar = 2) Then
                ReDim fvListaAguaSaneamiento(0)
            End If
            If (cmdEjecutar <> -1) Then
                cmdDestinoDetalleAguaS_Click
                nClick = 1
            End If
            If cmbDestCred.ListCount > 1 Then
                Set nMatMontoPre = Nothing 'JOEP20190919 ERS042 CP-2018
            End If
        Case ColocDestino.gColocDestinoActFijo:
            cmdDestinoDetalleAguaS.Visible = True
            cmdCreditoVerde.Visible = True 'EAAS20191401 SEGUN 018-GM-DI_CMACM
            If cmbDestCred.ListCount > 1 Then
                Set nMatMontoPre = Nothing 'JOEP20190919 ERS042 CP-2018
            End If
        Case ColocDestino.gColocDestinoConsumo:
            cmdDestinoDetalleAguaS.Visible = True
            cmdCreditoVerde.Visible = True 'EAAS20191401 SEGUN 018-GM-DI_CMACM
            If cmbDestCred.ListCount > 1 Then
                Set nMatMontoPre = Nothing 'JOEP20190919 ERS042 CP-2018
            End If
        Case ColocDestino.gColocDestinoRefracionVivienda:
            cmdDestinoDetalleAguaS.Visible = True
            If cmbDestCred.ListCount > 1 Then
                Set nMatMontoPre = Nothing 'JOEP20190919 ERS042 CP-2018
            End If
        Case ColocDestino.gColocDestinoRemodelacionVivienda:
            cmdDestinoDetalleAguaS.Visible = True
            If cmbDestCred.ListCount > 1 Then
                Set nMatMontoPre = Nothing 'JOEP20190919 ERS042 CP-2018
            End If
        Case ColocDestino.gColocDestinoAmpliacionVivienda:
            cmdDestinoDetalleAguaS.Visible = True
            If cmbDestCred.ListCount > 1 Then
                Set nMatMontoPre = Nothing 'JOEP20190919 ERS042 CP-2018
            End If
        Case ColocDestino.gColocDestinoMejoramientoVivienda:
            cmdDestinoDetalleAguaS.Visible = True
        'FIN EAAS 20180727 ERS-054-2018
        'JOEP20190919 ERS042 CP-2018
        Case 19
            cmdDestinoDetalleAguaS.Visible = True
                If cmbDestCred.ListCount > 1 Then
                    Set nMatMontoPre = Nothing
                End If
         'JOEP20190919 ERS042 CP-2018
    End Select
    'INICIO EAAS 20180727 ERS-054-2018
    If ((cmdEjecutar = 2 Or cmdEjecutar = 1) And cmbDestCred.Text <> "") Then
        If (UBound(fvListaAguaSaneamiento) > 0 And nClick <> 1) Then
            MsgBox "Se está cambiando el destino, el detalle agua y saneamiento se limpiará", vbInformation, "Alerta"
            nSumaAguaSaneamiento = 0 'EAAS20191401 SEGUN 018-GM-DI_CMACM
            nSumaCreditoVerde = 0 'EAAS20191401 SEGUN 018-GM-DI_CMACM
            ReDim fvListaAguaSaneamiento(0)
        End If
'        ReDim fvListaAguaSaneamiento(0)
    End If
    'FIN EAAS 20180727 ERS-054-2018
    'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM
    If ((cmdEjecutar = 2 Or cmdEjecutar = 1) And cmbDestCred.Text <> "") Then
        If (UBound(fvListaCreditoVerde) > 0 And nClick <> 1) Then
            MsgBox "Se está cambiando el destino, el detalle crédito verde se limpiará", vbInformation, "Alerta"
            nSumaAguaSaneamiento = 0
            nSumaCreditoVerde = 0
            ReDim fvListaCreditoVerde(0)
        End If
    End If
    'FIN EAAS20191401 SEGUN 018-GM-DI_CMACM
End If

'Agrego JOEP20190919 ERS042 CP-2018
If cmdBtnCancelar = 0 Then
    Set P = New COMDCredito.DCOMCredito
    If cmbSubProducto.Text <> "" And cmbDestCred.Text <> "" Then
        If Not CP_Mensajes(3, Trim(Right(cmbSubProducto.Text, 5))) Then Exit Sub
        Select Case Trim(Right(cmbSubProducto.Text, 5))
            Case 521, 718
                Set rsSubDestinos = P.CatalogoProCargaSubDestinos(Trim(Right(cmbSubProducto.Text, 5)), CInt(Trim(Right(cmbDestCred.Text, 3))), lblCodigo, Right(cmbCondicion.Text, 10))
                If Not (rsSubDestinos.BOF And rsSubDestinos.EOF) Then
                    If rsSubDestinos!nConsValor <> 0 Then
                            Call Llenar_Combo_con_Recordset(rsSubDestinos, cmbSubDestCred)
                            Call CambiaTamañoCombo(cmbSubDestCred, 300)
                            cmbSubDestCred.Visible = True
                            If cmbSubDestCred.Enabled = True Then
                                cmbSubDestCred.SetFocus
                            End If
                    Else
                        cmbSubDestCred.Visible = False
                        Call CP_CargaAporte(lblCodigo, Trim(Right(cmbSubProducto.Text, 5)), 0, Trim(Right(cmbDestCred.Text, 9)), , , IIf(cmbMoneda.Text = "", "0", Right(cmbMoneda.Text, 3)))
                    End If
                Else
                    cmbSubDestCred.Visible = False
                    Set nMatMontoPre = Nothing
                    txtMontoSol.Enabled = True
                End If
            Case 803
                Call CP_CargaAporte(lblCodigo, Trim(Right(cmbSubProducto.Text, 5)), 0, Trim(Right(cmbDestCred.Text, 9)), , , IIf(cmbMoneda.Text = "", "0", Right(cmbMoneda.Text, 3)))
        End Select
    End If
    RSClose rsSubDestinos
End If

If Not CP_Mensajes(5, Trim(Right(cmbSubProducto.Text, 5))) Then
    cmbDestCred.ListIndex = IndiceListaCombo(cmbDestCred, 1)
    Exit Sub
End If
'Agrego JOEP20190919 ERS042 CP-2018
End Sub

'Agrego JOEP20190919 ERS042 CP-2018
Private Sub cmbMoneda_Click()
    If cmbMoneda.Text <> "" Then Call PintaMoneda(Right(cmbMoneda.Text, 1))
    Call CP_DatosDefaut(4000)
End Sub

Private Sub cmbSubDestCred_Click()

If cmbSubDestCred.Text = "" Then Exit Sub
   
    If Trim(Right(cmbSubProducto.Text, 5)) = "521" Then
        Call CP_CargaAporte(lblCodigo, Trim(Right(cmbSubProducto.Text, 5)), Trim(Right(cmbSubDestCred.Text, 9)), Trim(Right(cmbDestCred.Text, 9)), Trim(Right(cmbCondicion.Text, 10)))
    ElseIf Trim(Right(cmbSubProducto.Text, 5)) = "718" Then
        Call CP_CargaAporte(lblCodigo, Trim(Right(cmbSubProducto.Text, 5)), Trim(Right(cmbSubDestCred.Text, 9)), Trim(Right(cmbDestCred.Text, 9)), Trim(Right(cmbCondicion.Text, 10)))
    End If
End Sub

Private Sub cmbTpDoc_Click()
If Not CP_Mensajes(2, Trim(Right(cmbSubProducto.Text, 5))) Then Exit Sub
    If Trim(Right(cmbSubProducto.Text, 10)) = "703" And cmbTpDoc.Text <> "" Then
        Call CP_CargaAporte(lblCodigo, Trim(Right(cmbSubProducto.Text, 5)), 0, IIf(cmbDestCred.Text = "", 0, Trim(Right(cmbDestCred.Text, 9))), Trim(Right(cmbCondicion.Text, 10)), IIf(cmbTpDoc.Text = "", 0, Trim(Right(cmbTpDoc.Text, 9))), Trim(Right(cmbMoneda.Text, 3)), Trim(Right(cmbCondicionOtra, 3)))
    Else
        Call CP_DatosDefaut(7000)
    End If
End Sub
'Agrego JOEP20190919 ERS042 CP-2018
'EAAS20191401 SEGUN 018-GM-DI_CMACM
Private Sub cmdCreditoVerde_Click()
Dim ofrmCreditoVerde As frmCredVerde
On Error GoTo ErrcmdCreditoVerde
Set ofrmCreditoVerde = New frmCredVerde
    If (Trim(Right(cmbSubProducto.Text, 5))) = "" Then
        MsgBox "Por Favor,Debe Ingrese el tipo de Producto", vbInformation, "Aviso"
        Exit Sub
    End If
    

    If (txtMontoSol.Text) = "" Then
        MsgBox "Por Favor,Debe Ingrese el monto a solicitar", vbInformation, "Aviso"
        Exit Sub
    End If
If Not IsArray(fvListaCreditoVerde) Then ReDim fvListaCreditoVerde(0)
nMontoCreditoVariable = CDbl(txtMontoSol.Text) - nSumaAguaSaneamiento - nSumaCreditoVerde
nCentinela = 0
If (nMontoCreditoVariable <> CDbl(txtMontoSol.Text) And nMontoCreditoVariable <> 0) Then
nCentinela = 1
End If
ofrmCreditoVerde.Inicio fvListaCreditoVerde, CInt(Trim(Right(cmbSubProducto.Text, 5))), cmbDestCred.Text, nMontoCreditoVariable, IIf(cmdEjecutar = 1, "", ActXCtaCred.NroCuenta), nSumaCreditoVerde
Exit Sub
ErrcmdCreditoVerde:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
'INICIO EAAS 20180727 SEGUN ERS-052-2018***
Private Sub cmdDestinoDetalleAguaS_Click()
Dim ofrmAguaSaneamiento As frmCredAguaSaneamiento
On Error GoTo ErrDestinoDetalleAguaS
Set ofrmAguaSaneamiento = New frmCredAguaSaneamiento
    If (Trim(Right(cmbSubProducto.Text, 5))) = "" Then
        MsgBox "Por Favor,Debe Ingrese el tipo de Producto", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If (txtMontoSol.Text) = "" Or CCur(txtMontoSol.Text) = 0 Then
    'MsgBox "Por Favor, ingrese el monto a solicitar", vbInformation, "Aviso"
'JOEP20190919 ERS042 CP-2018
        If cmbSubDestCred.Visible = True Then
            If cmbSubDestCred.Text = "" Then
                MsgBox "Seleccione el Sub Destino", vbInformation, "Aviso"
                cmbSubDestCred.SetFocus
            Else
                MsgBox "Por Favor, ingrese el monto a solicitar", vbInformation, "Aviso"
            End If
        Else
            MsgBox "Por Favor, ingrese el monto a solicitar", vbInformation, "Aviso"
        End If
        Exit Sub
    End If
'JOEP20190919 ERS042 CP-2018

'JOEP20190919 ERS042 CP-2018
If cmbTpDoc.Visible = True Then
    If Not CP_Mensajes(1, Trim(Right(cmbSubProducto.Text, 5))) Then Exit Sub
'Dim nTpDoc As Long
'Dim nTpInt As Long
    nSubDestino = IIf(cmbSubDestCred.Visible = True And cmbSubDestCred.Text <> "", Trim(Right(cmbSubDestCred.Text, 8)), 0)
    nTpDoc = IIf(cmbTpDoc.Visible = True And cmbTpDoc.Text <> "" And nTpCmbTpDoc = 2, Trim(Right(cmbTpDoc, 10)), 0)
    nTpIngr = IIf(cmbTpDoc.Visible = True And cmbTpDoc.Text <> "" And nTpCmbTpDoc = 1, Trim(Right(cmbTpDoc, 10)), 0)
    nTpInt = IIf(cmbTpDoc.Visible = True And cmbTpDoc.Text <> "" And nTpCmbTpDoc = 3, Trim(Right(cmbTpDoc, 10)), 0)
    
    If Not CatalogoValidador(4000, txtMontoSol.Text, spnCuotas.valor, spnPlazo.valor, CInt(Trim(Right(cmbDestCred.Text, 3))), nSubDestino, nTpDoc, nTpIngr) Then
        If txtMontoSol.Enabled = True Then
            txtMontoSol.SetFocus
        End If
        Exit Sub
    End If
End If
'JOEP20190919 ERS042 CP-2018

If Not IsArray(fvListaAguaSaneamiento) Then ReDim fvListaAguaSaneamiento(0)
'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM
nMontoCreditoVariable = CDbl(txtMontoSol.Text) - nSumaAguaSaneamiento - nSumaCreditoVerde
nCentinela = 0
If (nMontoCreditoVariable <> CDbl(txtMontoSol.Text) Or nMontoCreditoVariable = 0) Then ' EAAS20191004 SEGUN 018-GM-DI_CMACM
nCentinela = 1
End If
'FIN EAAS20191401 SEGUN 018-GM-DI_CMACM
ofrmAguaSaneamiento.Inicio fvListaAguaSaneamiento, CInt(Trim(Right(cmbSubProducto.Text, 5))), cmbDestCred.Text, nMontoCreditoVariable, IIf(cmdEjecutar = 1, "", ActXCtaCred.NroCuenta), nCentinela, nSumaAguaSaneamiento 'EAAS20191401 SEGUN 018-GM-DI_CMACM nMontoCreditoVariable
Exit Sub
ErrDestinoDetalleAguaS:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

'FIN EAAS 20180727 SEGUN ERS-052-2018***
'END EJVG *******
Private Sub cmbDestCred_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtfechaAsig.SetFocus
    End If
End Sub

Private Sub cmbFuentes_Click()
    
    If cmbFuentes.ListIndex >= 0 Then
       ' Call CargaTiposCredito(MatTipoFte(cmbFuentes.ListIndex))
    End If
End Sub

Private Sub cmbFuentes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdFuentes.SetFocus
    End If
End Sub

Private Sub cmbInstitucion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbModular.SetFocus
    End If
End Sub


Private Sub cmbModular_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub

Private Sub CmbMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtMontoSol.Enabled Then
            txtMontoSol.SetFocus
        Else
            spnCuotas.SetFocus
        End If
    End If
End Sub

'Private Sub cmbSubProducto_Click()
'    Dim sMensaje  As String
'    Dim oDCred As COMDCredito.DCOMCredito
'    Dim nValor As Integer
'    Dim nIndice As Integer
'    Set oDCred = New COMDCredito.DCOMCredito
'
'    cmbCondicionOtra.Enabled = True
'    If cmbSubProducto.ListIndex <> -1 Then
'        'If CInt(Trim(Right(cmbSubTipo.Text, 10))) = gColConsuDctoPlan Then
'        'MAVM 20091028 Nuevo Producto
'        'If CInt(Trim(Right(cmbSubProducto.Text, 10))) = gColConsuDctoPlan Or CInt(Trim(Right(cmbSubProducto.Text, 10))) = gColPYMEConv Then
'        If CInt(Trim(Right(cmbSubProducto.Text, 10))) = 512 Or CInt(Trim(Right(cmbSubProducto.Text, 10))) = gColConsuDctoPlan Or CInt(Trim(Right(cmbSubProducto.Text, 10))) = gColProConsumoPerDesPla Or CInt(Trim(Right(cmbSubProducto.Text, 10))) = gColProHIpoConveTrayDir Then
'        'ALPA 20100609 ************************************************
'        '    Call HabilitaCreditoDsctoPlanilla(True)
'            Call HabilitaCreditoDsctoPlanilla(False)
'        '**************************************************************
'            If rsInstituc.RecordCount = 0 Then
'                MsgBox "No existen instituciones por convenio", vbInformation, "Mensaje"
'                Exit Sub
'            End If
'            rsInstituc.MoveFirst
'            Call frmCredSolicitudConvenio.Inicio(rsInstituc)
'            sCARBEN = frmCredSolicitudConvenio.fsCARBEN
'            sCargo = frmCredSolicitudConvenio.fsCargo
'            sCodInstitucion = frmCredSolicitudConvenio.fsCodInstitucion
'            sCodModular = frmCredSolicitudConvenio.fsCodModular
'            sT_Plani = frmCredSolicitudConvenio.fsT_Plani
'        Else
'            Call HabilitaCreditoDsctoPlanilla(False)
'        End If
'        'EJVG20130222 ***
'        Set oDCred = New COMDCredito.DCOMCredito
'        If bAmpliacion Then
'            LblCondProd.Caption = Trim(" " & cmbCondicion.List(4))
'        Else
'            nValor = oDCred.DefineCondicionCredito(oRelPersCred.TitularPersCod, CInt(Right(cmbSubProducto.Text, 3)), gdFecSis, bRefinanciar)
'            nIndice = IndiceListaCombo(cmbCondicion, Trim(str(nValor)))
'            If nIndice <> -1 Then
'                LblCondProd.Caption = Trim(" " & cmbCondicion.List(nIndice))
'            End If
'        End If
'        'END EJVG *******
'    End If
'    Set oDCred = Nothing
'End Sub
'EJVG20130503 ***
Private Sub cmbSubProducto_Click()
    Dim oDCred As COMDCredito.DCOMCredito
    oCredAgrico.Registrar = False 'WIOR 20130723
    cmbCondicionOtra.Enabled = True
    
'Agrego JOEP20190919 ERS042 CP-2018
    If bRefinanciar = False Then
        Select Case Trim(Right(cmbSubProducto.Text, 10))
            Case 520, 525
                CP_CargaTpDoc Trim(Right(cmbSubProducto.Text, 10)), 22000
            Case 707, 718
                CP_CargaTpDoc Trim(Right(cmbSubProducto.Text, 10)), 34000
            Case 703
                CP_CargaTpDoc Trim(Right(cmbSubProducto.Text, 10)), 48000
            Case Else
                nTpCmbTpDoc = 0 '0:Nada - Para identificar el tipo de Combo cmbTpDoc
                cmbTpDoc.ListIndex = -1
                cmbTpDoc.Visible = False
                lblTpDoc.Visible = False
                sT_Plani = ""
        End Select
    End If
'Agrego JOEP20190919 ERS042 CP-2018
        
    If cmbSubProducto.ListIndex <> -1 Then
'**ARLO20180712 ERS042 - 2018
        Set objProducto = New COMDCredito.DCOMCredito
        If objProducto.GetResultadoCondicionCatalogo("N0000066", Trim(Right(cmbSubProducto.Text, 10))) Then
'**ARLO20180712 ERS042 - 2018
        'If CInt(Trim(Right(cmbSubProducto.Text, 10))) = 512 Or CInt(Trim(Right(cmbSubProducto.Text, 10))) = gColConsuDctoPlan Or CInt(Trim(Right(cmbSubProducto.Text, 10))) = gColProConsumoPerDesPla Or CInt(Trim(Right(cmbSubProducto.Text, 10))) = gColProHIpoConveTrayDir Then
            Call HabilitaCreditoDsctoPlanilla(False)
            If rsInstituc.RecordCount = 0 Then
                MsgBox "No existen instituciones por convenio", vbInformation, "Mensaje"
                Exit Sub
            End If
            rsInstituc.MoveFirst
            Call frmCredSolicitudConvenio.Inicio(rsInstituc)
            sCARBEN = frmCredSolicitudConvenio.fsCARBEN
            sCargo = frmCredSolicitudConvenio.fsCargo
            sCodInstitucion = frmCredSolicitudConvenio.fsCodInstitucion
            sCodModular = frmCredSolicitudConvenio.fsCodModular
            sT_Plani = frmCredSolicitudConvenio.fsT_Plani
        Else
            Call HabilitaCreditoDsctoPlanilla(False)
            'WIOR 20130723 ******************************
            If Trim(Right(cmbProductoCMACM.Text, 5)) = "600" Then
                If fbActivo Then
                    Call oCredAgrico.inicia(, Trim(Right(cmbSubProducto.Text, 5)))
                    If Not oCredAgrico.Registrar Then
                        cmbSubProducto.ListIndex = -1
                        Exit Sub
                    End If
                End If
            Else
                 oCredAgrico.Registrar = False
            End If
            'WIOR FIN ***********************************
        End If
        'MIOL 20130726, SEGUN ERS074-SOBRE RUEDAS ULTIMO CORREGIDO
        Me.lblVendedor.Visible = False
        Me.txtVendedor.Visible = False
'**ARLO20180712 ERS042 - 2018
        Set objProducto = New COMDCredito.DCOMCredito
        If objProducto.GetResultadoCondicionCatalogo("N0000067", Trim(Right(cmbSubProducto.Text, 10))) Then
'**ARLO20180712 ERS042 - 2018
        'If CInt(Trim(Right(cmbSubProducto.Text, 10))) = 510 Or CInt(Trim(Right(cmbSubProducto.Text, 10))) = 706 Or CInt(Trim(Right(cmbSubProducto.Text, 10))) = 511 Then
            If Trim(Right(cmbCondicionOtra2.Text, 3)) <> "0" Then
                Me.lblVendedor.Visible = True
                Me.txtVendedor.Visible = True
            End If
        End If
        'END MIOL *******************************
        Call CargaConfigMonedaTpoProd(Trim(Right(cmbSubProducto.Text, 3))) 'WIOR 20151222
        Call EstableceCondicionSubProducto
        
'Agrego JOEP20190919 ERS042 CP-2018
        If cmdEjecutar = 1 Then
            cmbSubDestCred.Visible = False
            cmdDestinoDetalleAguaS.Visible = False
        End If
                        
        If cmdEjecutar <> -1 Then
            If Not CP_Mensajes(4, Right(cmbSubProducto.Text, 3)) Then Exit Sub
            Call CP_HabilitaControles(True)
            Call CatalogoLlenaCombox(Right(cmbSubProducto.Text, 3), 5000)
            Call CatalogoLlenaCombox(Right(cmbSubProducto.Text, 3), 2000, , Trim(Right(IIf(cmbCondicion.Text = "", 0, cmbCondicion.Text), 9)))
            Call limpiaCatalogo
            Call CP_DatosDefaut(46000)
            Call CP_DatosDefaut(7000)
            If bRefinanciar = True Then
                cmbTpDoc.Visible = False
                lblTpDoc.Visible = False
                cmbDestCred.Enabled = False
                cmbCondicionOtra.Enabled = False
            End If
        End If
        Set nMatMontoPre = Nothing
'Agrego JOEP20190919 ERS042 CP-2018
        
    End If
End Sub
Private Sub EstableceCondicionSubProducto()
    Dim oDCred As New COMDCredito.DCOMCredito
    Dim nValor As Integer
    Dim nIndice As Integer
    
    If bAmpliacion Then
        LblCondProd.Caption = Trim(" " & cmbCondicion.List(4))
    Else
        'nValor = oDCred.DefineCondicionCredito(oRelPersCred.TitularPersCod, val(Right(cmbSubProducto.Text, 3)), gdFecSis, bRefinanciar, val(spnCuotas.valor))
        nValor = oDCred.DefineCondicionCredito(oRelPersCred.TitularPersCod, val(Right(cmbSubProducto.Text, 3)), gdFecSis, bRefinanciar, 0) 'WIOR 20141210
        nIndice = IndiceListaCombo(cmbCondicion, Trim(str(nValor)))
        If nIndice <> -1 Then
            LblCondProd.Caption = Trim(" " & cmbCondicion.List(nIndice))
        End If
    End If
    Set oDCred = Nothing
End Sub
'END EJVG *******

Private Sub cmbSubProducto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbMoneda.Enabled Then
            cmbMoneda.SetFocus
        Else
            If spnCuotas.Enabled = True Then
                spnCuotas.SetFocus
            End If
        End If
    End If
End Sub

Private Sub cmbProductoCMACM_Click()
Dim oCred As COMDCredito.DCOMCredito
Dim nValor As Integer
Dim sValor As String
Dim nIndice As Integer
    
    'JOEP20190208 CP
    lblTpDoc.Visible = False
    cmbTpDoc.Visible = False
    cmbSubDestCred.Visible = False
    'JOEP20190208 CP
    
    '***Modificado por ELRO 20110831, según Acta 222-2011/TI-D
    If Trim(Right(cmbProductoCMACM.Text, 3)) = "700" And PersoneriaTitular <> gPersonaNat Then
    MsgBox "Persona Jurídica no corresponde tipo de credito", vbInformation, "Aviso"
    cmbProductoCMACM.ListIndex = -1
    If cmbProductoCMACM.Enabled = True Then
        cmbProductoCMACM.SetFocus
    End If
    Exit Sub
    End If
    If Trim(Right(cmbProductoCMACM.Text, 3)) = "800" And PersoneriaTitular <> gPersonaNat Then
    MsgBox "Persona Jurídica no corresponde tipo de credito", vbInformation, "Aviso"
    cmbProductoCMACM.ListIndex = -1
    If cmbProductoCMACM.Enabled = True Then
        cmbProductoCMACM.SetFocus
    End If
    Exit Sub
    End If
    '*********************************************************
    'ALPA 20100603 BASII****************
    'Call CargaSubProducto(Trim(Right(cmbProductoCMACM.Text, 3)))'Comento JOEP20190118 CP
    Call CargaSubProducto(Trim(Right(cmbProductoCMACM.Text, 3)), bRefinanciar, bAmpliacion)
    'Call CargaSubTiposCredito(Trim(Right(cmbProductoCMACM.Text, 3)))
    '***********************************
    If Not oRelPersCred Is Nothing And Trim(cmbProductoCMACM.Text) <> "" Then
        ' CMACICA_CSTS - 18/11/2003 - ------------------------------------------------------------------
        Set oCred = New COMDCredito.DCOMCredito
        sValor = oCred.VerificaCombinacionProd(oRelPersCred.TitularPersCod, CInt(Mid(Right(cmbProductoCMACM.Text, 3), 1, 2)), IIf(Right(cmbCondicionOtra, 1) = "1", True, False))
        
        If sValor <> "" And bAmpliacion = False Then
           MsgBox sValor, vbInformation, "Aviso"
           cmbProductoCMACM.ListIndex = -1
           cmbProductoCMACM.SetFocus
        Else
'            Set oCred = New COMDCredito.DCOMCredito
'            'ARCV 20-02-2007
'            If bAmpliacion Then
'                LblCondProd.Caption = " " & cmbCondicion.List(4)
'                LblCondProd.Caption = Trim(LblCondProd.Caption)
'            Else
'                nValor = oCred.DefineCondicionCredito(oRelPersCred.TitularPersCod, CInt(Mid(Right(cmbProductoCMACM.Text, 3), 1, 3)), gdFecSis, bRefinanciar)
'                nIndice = IndiceListaCombo(cmbCondicion, Trim(str(nValor)))
'                'cmbCondicion.ListIndex = IndiceListaCombo(cmbCondicion, Trim(str(nValor)))
'                If nIndice <> -1 Then
'                    LblCondProd.Caption = " " & cmbCondicion.List(nIndice)
'                    LblCondProd.Caption = Trim(LblCondProd.Caption)
'                End If
'            End If
        End If
        Set oCred = Nothing
        ' ----------------------------------------------------------------------------------------------
    End If
'Agrego JOEP20190919 ERS042 CP-2018
    If cmdEjecutar = 1 Then Call CP_LimpiaCombo
'Agrego JOEP20190919 ERS042 CP-2018
End Sub

Private Sub cmbProductoCMACM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbSubProducto.Enabled = True Then
            cmbSubProducto.SetFocus
        End If
    End If
End Sub

Private Sub CmdAmpliacion_Click()
Dim rsAmpliadoNew, RsValidaAmpNew As ADODB.Recordset 'ARLO20181126
Dim oCredito As COMDCredito.DCOMCredito 'ARLO20181126

'JOEP20190225 CP
If Not CP_Mensajes(7, Trim(Right(cmbSubProducto.Text, 5))) Then
    Exit Sub
End If
'JOEP20190225 CP
    'ARCV 09-03-2007
    If cmbMoneda.ListIndex = -1 Then
        MsgBox "Seleccione la Moneda del Credito", vbInformation, "Aviso"
        'cmbMoneda.SetFocus 'Comento JOEP20190213 CP
        'JOEP20190213 CP
        If cmbMoneda.Enabled = True Then
            cmbMoneda.SetFocus
        End If
        'JOEP20190213 CP
        Exit Sub
    Else
        cmbMoneda.Enabled = False
    End If
    
    'FRHU 20140424 TI-ERS015-2014
    
    'FrmCredAmpliado.cPersCod = lblCodigo.Caption
    'FrmCredAmpliado.cMoneda = Trim(Right(cmbMoneda.Text, 20))
    'FrmCredAmpliado.Show vbModal
    nMontoTotal = 0
    'Call frmCredAmpliadoNew.Inicio(lblCodigo.Caption, Trim(Right(cmbMoneda.Text, 20)), lblNombre.Caption, nMontoTotal, rsAmpliado) 'COMENTADO POR VAPI SEGUN ERS 001-2017
    
    'agregado por vapi SEGÙN ERS TI-ERS001-2017
    If bPresolAmpliAuto Then
        'Call frmCredAmpliadoNew.Inicio(lblCodigo.Caption, Trim(Right(cmbMoneda.Text, 20)), lblNombre.Caption, nMontoTotal, rsAmpliado, rsPresol!cCtaCodAmpliado) 'Comento JOEP20190225 CP
        Call frmCredAmpliadoNew.Inicio(lblCodigo.Caption, Trim(Right(cmbMoneda.Text, 20)), lblNombre.Caption, nMontoTotal, rsAmpliado, rsPresol!cCtaCodAmpliado, Trim(Right(cmbSubProducto.Text, 5))) 'JOEP20190225 CP
        bPresolAmpliAuto = False
    Else
        'Call frmCredAmpliadoNew.Inicio(lblCodigo.Caption, Trim(Right(cmbMoneda.Text, 20)), lblNombre.Caption, nMontoTotal, rsAmpliado)'Comento JOEP20190225 CP
        Call frmCredAmpliadoNew.Inicio(lblCodigo.Caption, Trim(Right(cmbMoneda.Text, 20)), lblNombre.Caption, nMontoTotal, rsAmpliado, , Trim(Right(cmbSubProducto.Text, 5))) 'JOEP20190225 CP
    End If
    'fin agregado por vapi
    
    If frmCredAmpliadoNew.nIdCampana > 0 Then
        cmbCondicionOtra.ListIndex = IndiceListaCombo(cmbCondicionOtra, Trim(str(FrmCredAmpliado.nIdCampana)))
    End If
    'FIN FRHU 20140424
End Sub

Private Sub cmdCancela_Click()
'JOEP20190919 ERS042 CP-2018
cmdBtnCancelar = 1
nSubDestAnt = 0
ReDim nMatMontoPre(0)
bEntrotxtMontoSol = False
lblTpDoc.Visible = False
cmbTpDoc.Visible = False
'JOEP20190919 ERS042 CP-2018
    Call HabilitaIngresoSolicitud(False)
    ActXCtaCred.Enabled = False
    cmbProductoCMACM.Enabled = False
    cmbSubProducto.Enabled = False
    cmbMoneda.Enabled = False
    If cmdEjecutar = 1 Then
        Set MatCredRef = Nothing
        Set MatCredSust = Nothing
        Unload frmCredRefinanc
    End If
    
    Call LimpiaPantalla
    
    cmdEjecutar = -1
    cmdRelaciones.Enabled = False
    CmdAmpliacion.Enabled = False
    cmdEnvioEstCta.Enabled = False 'JUEZ 20130527
    'bAmpliacion = False
    ReDim MatFuentes(0)
    Set rsAmpliado = Nothing
    '***Modificado por ELRO 20111017, según Acta 222-2011/TI-D
    PersoneriaTitular = 0
    '*********************************************************
    If bLeasing Then 'EJVJ20120720
        cmdNuevo.Enabled = False
    End If
    ReDim fvListaCompraDeuda(0) 'EJVG20160201 ERS002-2016
    ReDim fvListaAguaSaneamiento(0) 'EAAS20181002 ERS054-2018
    ReDim fvListaCreditoVerde(0) 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    nSumaAguaSaneamiento = 0 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    nSumaCreditoVerde = 0 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    'agregado por vapi SEGÙN ERS TI-ERS001-2017
    If nPermiso = Registrar And Not bPresol Then
        cmdpresolicitud.Enabled = True
        bPresolAmpliAuto = False
        bPresol = False
        bPreSolOperacion = False
    End If
    'fin agregado por vapi
'Agrego JOEP20190919 ERS042 CP-2018
    cmdDestinoDetalleAguaS.Visible = False
    cmdCreditoVerde.Visible = False
    cmbSubDestCred.Visible = False
    Call Cargar_Objetos_Controles
    cmdBtnCancelar = 0
'Agrego JOEP20190919 ERS042 CP-2018
End Sub
'EJVG20160201 ERS002-2016***
Private Sub cmdDestinoDetalle_Click()
    Dim ofrmCompraDeuda As frmCredCompraDeuda
    
    On Error GoTo ErrDestinoDetalle
    'ARLO20180317
    If (Trim(Right(cmbSubProducto.Text, 5))) = "" Then
        MsgBox "Por Favor,Debe Ingrese el tipo de Producto", vbInformation, "Aviso"
        Exit Sub
    End If
    'ARLO20180317
    
    If Not bRefinanciar Then 'ARLO20180317
    Select Case val(Trim(Right(cmbDestCred.Text, 3)))
        Case ColocDestino.gColocDestinoCambEstructPasivo:
            Set ofrmCompraDeuda = New frmCredCompraDeuda
            If Not IsArray(fvListaCompraDeuda) Then ReDim fvListaCompraDeuda(0)
            ofrmCompraDeuda.Inicio fvListaCompraDeuda, (CInt(Trim(Right(cmbSubProducto.Text, 5)))) 'ARLO20180317
    End Select
    End If 'ARLO20180317
    
    Set ofrmCompraDeuda = Nothing
    Exit Sub
ErrDestinoDetalle:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
'END EJVG *******
Private Sub cmdEditar_Click()
    Call HabilitaIngresoSolicitud(True)
    ChkCap.Enabled = False
    ActXCtaCred.Enabled = False
    cmbProductoCMACM.Enabled = False
    cmbSubProducto.Enabled = False
    cmbMoneda.Enabled = False
    cmdRelaciones.Enabled = True
    cmdEnvioEstCta.Enabled = True 'JUEZ 20130527
    If bAmpliacion Then chkAutAmpliacion.Enabled = False 'JUEZ 20160509
    'ALPA 20100604 B2************************************
'Comento JOEP20190919 ERS042 CP-2018
    'cmbProductoCMACM.Enabled = True
    'cmbSubProducto.Enabled = True
'Comento JOEP20190919 ERS042 CP-2018
    '****************************************************
    cmdEjecutar = 2
'Agrego JOEP20190919 ERS042 CP-2018
    If cmbSubProducto.Text = "" Then
        MsgBox "No se migro el producto correctamente" & " Comuníquese con el Área de TI-Desarrollo", vbInformation, "Aviso"
        Call cmdCancela_Click
        Exit Sub
    End If
'Agrego JOEP20190919 ERS042 CP-2018
End Sub

Private Sub CmdEliminar_Click()
Dim oCredito As COMDCredito.DCOMCredito
    If MsgBox("Se va a Eliminar la Solicitud de Credito y Todos los Datos se perderan definitivamente." & Chr(10) & " Desea Continuar ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set oCredito = New COMDCredito.DCOMCredito
        Call oCredito.EliminaSolicitud(ActXCtaCred.NroCuenta)
        Set oCredito = Nothing
        Call LimpiaPantalla
    End If
End Sub

'JUEZ 20130527 ****************************************************************
Private Sub cmdEnvioEstCta_Click()
    fbActivaEnvio = True
    'frmEnvioEstadoCta.InicioCol ActXCtaCred.NroCuenta, LlenaRecordSet_Cliente
    frmEnvioEstadoCta.InicioCol ActXCtaCred.NroCuenta, LlenaRecordSet_Cliente, False 'APRI20180404 ERS036-2017
    fbRegistraEnvio = frmEnvioEstadoCta.RegistraEnvio
    Set frsEnvEstCta = frmEnvioEstadoCta.RecordSetDatos
    fnModoEnvioEstCta = frmEnvioEstadoCta.ModoEnvioEstCta
    fnModoEnvioEstCtaSiNo = frmEnvioEstadoCta.RegistraEnvioSiNo 'APRI20180404 ERS036-2017
    fnDebitoMismaCta = frmEnvioEstadoCta.DebitoMismaCta
End Sub
'END JUEZ *********************************************************************

'WIOR 20120914 *********************************************************************************************************
Private Sub cmdEvaluar_Click()
    Dim oEval As New COMDCredito.DCOMFormatosEval
    Dim oRs As New ADODB.Recordset
    Dim nFormEmpr As Boolean
    
    Set oRs = oEval.RecuperaFormatoEvaluacion(ActXCtaCred.NroCuenta)
    If (oRs.EOF And oRs.BOF) Then
        'If ValidaMultiForm(Right(cmbProductoCMACM.Text, 3)) Then'Comento JOEP20190919 ERS042 CP-2018
        If ValidaMultiForm(Right(cmbSubProducto.Text, 3)) Then 'JOEP20190919 ERS042 CP-2018
            If MsgBox("¿Desea utilizar un formato empresarial?", vbYesNo + vbInformation, "Alerta") = vbYes Then
                nFormEmpr = True
            Else
                nFormEmpr = False
            End If
        End If
    End If
    Call EvaluarCredito(ActXCtaCred.NroCuenta, False, 2000, CInt(Right(cmbProductoCMACM.Text, 3)), CInt(Right(cmbSubProducto, 3)), fnMontoExpEsteCred_NEW, , , nFormEmpr)
End Sub
'Private Sub EvaluarCredito(ByVal pcCtaCod As String)
'Dim DCredito As COMDCredito.DCOMCredito
'Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
'Dim rsDCredito As ADODB.Recordset
'Dim nEstado As Integer
'Dim nFomato As Integer
'Dim nmonto As Double
'Dim cPrd As String
'Dim cSPrd As String
'Set DCredito = New COMDCredito.DCOMCredito
'Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
'
'
'nEstado = DCredito.RecuperaEstadoCredito(pcCtaCod)
'
'If nEstado = 0 Then
'    MsgBox "Nº de Crédito no Existe.", vbInformation, "Aviso"
'    Exit Sub
'Else
'    If nEstado = 2000 Then
'
'
'        Set rsDCredito = DCredito.RecuperaSolicitudDatoBasicos(pcCtaCod)
'        If rsDCredito.RecordCount > 0 Then
'            nmonto = CDbl(Trim(rsDCredito!nmonto))
'            cSPrd = Trim(rsDCredito!cTpoProdCod)
'            cPrd = Mid(cSPrd, 1, 1) & "00"
'            If Mid(pcCtaCod, 9, 1) = "2" Then
'                nmonto = nmonto * CDbl(oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoDia))
'            End If
'        End If
'
'        nFomato = DCredito.AsignarFormato(cPrd, cSPrd, nmonto)
'
'        Select Case nFomato
'            Case 0, 4, 5: MsgBox "Crédito no se adecua para este Proceso.", vbInformation, "Aviso"
'            Case 1: Call frmCredEvalFormato1.Inicio(pcCtaCod, fnTipo)
'            Case 2: Call frmCredEvalFormato2.Inicio(pcCtaCod, fnTipo)
'            Case 3: Call frmCredEvalFormato3.Inicio(pcCtaCod, fnTipo)
'            'Case 4: Call frmCredEvalFormato4.Inicio(pcCtaCod, fnTipo)
'            'Case 5: Call frmCredEvalFormato5.Inicio(pcCtaCod, fnTipo)
'        End Select
'    Else
'        MsgBox "Nº de Crédito no se encuentra en estado Solicitado.", vbInformation, "Aviso"
'        Exit Sub
'    End If
'End If
'End Sub
'WIOR FIN *********************************************************************************************************
Private Sub cmdExaminar_Click()
Dim sCta As String
Dim oNeg As COMNCredito.NCOMCredito
    Screen.MousePointer = 11
ReDim nMatMontoPre(0) 'JOEP20190919 ERS042 CP-2018
    
    ' CMACICA_CSTS - 10/11/2003 --------------------------------------------------------------------
    If bRefinanciar Then
       bRefinanciarSustituir = True
    Else
        If bSustituirDeudor Then
           bRefinanciarSustituir = True
        Else
           bRefinanciarSustituir = False
        End If
    End If
    '------------------------------------------------------------------------------------------------
    
    'sCta = frmCredPersEstado.Inicio(Array(gColocEstSolic), "Solicitudes de Credito", , bRefinanciarSustituir, , gsCodAge)
    'sCta = frmCredPersEstado.Inicio(Array(gColocEstSolic), "Solicitudes de Credito", , bRefinanciarSustituir, , gsCodAge, bLeasing, , , bAmpliacion)
    sCta = frmCredPersEstado.Inicio(Array(gColocEstSolic), "Solicitudes de Credito", , bRefinanciarSustituir, , gsCodAge, bLeasing, , , bAmpliacion, gsCodCargo) 'JOEP20190205 CP gsCodCargo
    'LUCV20180417, Agregó bAmpliacion, según incidente en la edición de ampliados
    sCtaCod = sCta 'EAAS20180811 SEGUN ERS-054-2018
    If Len(Trim(sCta)) > 0 Then
        ActXCtaCred.NroCuenta = sCta
        Set oNeg = New COMNCredito.NCOMCredito
        ActXCtaCred.Enabled = False
        If bRefinanciar Or bSustituirDeudor Then
            If Not oNeg.EsRefinanciado(ActXCtaCred.NroCuenta) Then
                MsgBox "Credito No Es una Refinanciacion", vbInformation, "Aviso"
                Exit Sub
            End If
        ElseIf bAmpliacion Then 'LUCV20180417, Agregó
            If oNeg.EsRefinanciado(ActXCtaCred.NroCuenta) Then
                MsgBox "El Crédito, No se solicitó como Ampliado. Es Refinanciado.", vbInformation, "Aviso"
                Exit Sub
            End If
            '->***** LUCV20180417, Agregó
            If Not oNeg.EsAmpliado(ActXCtaCred.NroCuenta) Then
                MsgBox "El Crédito, No se solicitó como Ampliado.", vbInformation, "Aviso"
                Exit Sub
            End If
            '<-***** Fin LUCV20180417
        Else
            If oNeg.EsRefinanciado(ActXCtaCred.NroCuenta) Then
                MsgBox "Credito No Es una Solicitud Normal", vbInformation, "Aviso"
                Exit Sub
            End If
            '->***** LUCV20180417, Agregó
            If oNeg.EsAmpliado(ActXCtaCred.NroCuenta) Then
                MsgBox "El Crédito, No es una Solicitud Normal. Es Ampliado.", vbInformation, "Aviso"
                Exit Sub
            End If
            '<-***** Fin LUCV20180417
        End If
        Set oNeg = Nothing
        Call CargaDatos(sCta)
        cmdEditar.Enabled = True
        cmdEliminar.Enabled = True
        cmdGarantias.Enabled = True
        cmdGravar.Enabled = True
        cmdEvaluar.Enabled = True
        cmbMoneda.Enabled = False 'Agrego JOEP20190919 ERS042 CP-2018
        Call ControlesPermiso
        If bLeasing Then 'EJVJ20120720
            cmdNuevo.Enabled = False
        End If
    End If
    
    'Agrego JOEP20190919 ERS042 CP-2018
If Len(Trim(sCta)) > 0 Then
    If cmbSubProducto.Text = "" And cmdEjecutar = -1 Then
        MsgBox "No se migro el producto correctamente" & " Comuníquese con el Área de TI-Desarrollo", vbInformation, "Aviso"
        Call cmdCancela_Click
        Exit Sub
    End If
End If
'Agrego JOEP20190919 ERS042 CP-2018
    
End Sub

Private Sub cmdFuentes_Click()
    If Not ExisteTitular Then
        MsgBox "Debe Ingresar el Titular del Credito", vbInformation, "Aviso"
        cmdRelaciones.SetFocus
        Exit Sub
    End If
    
  
    Call frmPersona.Inicio(TitularCredito, PersonaActualiza)
    Call CargaFuentesIngreso(TitularCredito)
    oPersona.PersCodigo = TitularCredito
   
End Sub

Private Sub cmdGarantias_Click()
    'EJVG20150707 ***
    Dim frm As New frmGarantia
    If MsgBox("Seleccione [SI] para Registrar Nuevas Garantías." & Chr(13) & "Seleccione [NO] para Editar Garantías.", vbInformation + vbYesNo, "Aviso") = vbYes Then
        frm.Registrar
    Else
        frm.Editar
    End If
    Set frm = Nothing
    'If gsProyectoActual = "H" Then
    '    frmPersGarantiasHC.Show 1
    'Else
    '    'frmPersGarantias.Show 1'WIOR 20140912 COMENTO
    '    'WIOR 20140912 ************************
    '    Dim nCamp As Long
    '    nCamp = 0
    '    If Trim(CmbCondicionOtra.Text) <> "" Then
    '        nCamp = CLng(Right(CmbCondicionOtra.Text, 3))
    '   End If
    '
    '    frmPersGarantias.Inicio RegistroGarantia, , , nCamp
    '    'WIOR FIN ***************************
    'End If
    'END EJVG *******
End Sub

Private Sub cmdGrabar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim psNuevaCta As String
Dim oGen As COMDConstSistema.DCOMGeneral
Dim sMovAct As String
Dim bRefSusti As Boolean
Dim MatCredRefSust As Variant
Dim bSustiDeudor As Boolean
Dim oDPersGen As COMDPersona.DCOMPersGeneral 'JUEZ 20140603
Dim bSolicitaAutSectorEcon As Boolean 'JUEZ 20140603
Dim RsValidaAmp As ADODB.Recordset 'JUEZ 20160509
Dim lsComentario As String 'JUEZ 20160509

Dim bSolicitaAutZonaGeog As Boolean 'JOEP ERS047 20170901
Dim bSolicitaAutTpCredito As Boolean 'JOEP ERS047 20170901

'Se agrego para manejar el Tema de los Componentes
Dim MatCredRelaciones As Variant
Dim lbEliminaCobertura As Boolean 'EJVG20151015

'*** PEAC 20080811
Dim lbResultadoVisto As Boolean
Dim sPersVistoCod  As String
Dim sPersVistoCom As String
Dim loVistoElectronico As frmVistoElectronico
Set loVistoElectronico = New frmVistoElectronico
'RECO 20141226 ERS168-2014*********************
Dim oTC  As New COMDConstSistema.NCOMTipoCambio
Dim nTpoC As Double
Dim nMontoSol As Double
nTpoC = CDbl(oTC.EmiteTipoCambio(gdFecSis, TCFijoDia))
nMontoSol = IIf(Trim(Right(cmbMoneda.Text, 1)) = 1, val(txtMontoSol.Text), val(txtMontoSol.Text) * nTpoC)

'**ARLO20180712 ERS042 - 2018
Set objProducto = New COMDCredito.DCOMCredito
If objProducto.GetResultadoCondicionCatalogo("N0000068", Trim(Right(cmbSubProducto.Text, 10))) And nMontoSol < 350 Then
'**ARLO20180712 ERS042 - 2018
'If Trim(Right(cmbSubProducto.Text, 3)) = "703" And nMontoSol < 350 Then
    MsgBox "No se puede solicitar un monto menor a  350 o su equivalente en moneda extranjera si el subproducto es: Rapiflash con tu plazo fijo", vbInformation, "Alerta"
    Exit Sub
End If
'RECO FIN *************************************
'FRHU 20160615 ERS002-2016
Dim lnFila As Integer
Dim lsListaPrdPersRelac As String
Dim lsPrdPersRelac As String
Dim lsListaCondCred As String
Dim lsCondCred As String
Dim lsPersCod As String
'FIN FRHU 20160615
'FRHU 20140514 Observacion
If Not ValidaDatos Then
    Exit Sub
End If
'FIN FRHU 20140514
'***Modificado por ELRO 20110831, según Acta 222-2011/TI-D
If Trim(Right(cmbProductoCMACM.Text, 3)) = "700" And PersoneriaTitular <> gPersonaNat Then
    MsgBox "Persona Jurídica no corresponde tipo de credito", vbInformation, "Aviso"
    cmbProductoCMACM.ListIndex = -1
    If cmbProductoCMACM.Enabled = True Then
        cmbProductoCMACM.SetFocus
    End If
    Exit Sub
End If
If Trim(Right(cmbProductoCMACM.Text, 3)) = "800" And PersoneriaTitular <> gPersonaNat Then
    MsgBox "Persona Jurídica no corresponde tipo de credito", vbInformation, "Aviso"
    cmbProductoCMACM.ListIndex = -1
    If cmbProductoCMACM.Enabled = True Then
        cmbProductoCMACM.SetFocus
    End If
    Exit Sub
End If
'ALPA 20160419********************************************
If spnPlazo.valor < 30 Then
    MsgBox "El crédito no cumple con las especificaciones de plazo, no debe ser menor a 30 días", vbInformation, "Aviso"
    Exit Sub
End If
'*********************************************************
'*********************************************************
'*********************************************************
'ALPA 20140708********************************************
Dim objCreditoDA As COMDCredito.DCOMCredito
Set objCreditoDA = New COMDCredito.DCOMCredito
Dim oRsA As ADODB.Recordset
Set oRsA = New ADODB.Recordset
Dim oRsRel As New ADODB.Recordset 'FRHU 20160615 ERS002-2016
Dim nEdad As Integer
Dim sPersonaCodCony As String
Dim nIdCampana As Integer 'ARLO20170818
Dim nCantCompraIFIS As Integer '**ARLO2017113 ERS070 - 2017
Dim rsCrediCancelado As ADODB.Recordset '**ARLO2017113 ERS070 - 2017
Dim rsRCC As ADODB.Recordset '**ARLO2017113 ERS070 - 2017
Dim lcMovNro As String 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
Dim nFormatoEliminado As Integer 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
Dim nFormato_NEW As Integer 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018

nIdCampana = Trim(Right(cmbCondicionOtra.Text, 10)) 'ARLO20170818

    'ARLO 20170911
    Dim oDCreditos As COMDCredito.DCOMCreditos
    Set oDCreditos = New COMDCredito.DCOMCreditos
    Dim oRsC As ADODB.Recordset
    Set oRsC = New ADODB.Recordset
    
    If (oDCreditos.VerificaCampania(nIdCampana)) Then  'BY ARLO20171127 'ARLO20170103 COMENTO Not
    'If (nIdCampana <> 116) Then 'ARLO20170818 'BY ARLO20171127 COMENT
    'If Not oDCreditos.VerificaClienteCampaniaSolicitud(lblCodigo) Then
        Set oRsC = oDCreditos.VerificaClienteCampaniaSolicitud(lblCodigo.Caption)
        If Not (oRsC.EOF And oRsC.BOF) Then
            If (oRsC!bImpreso <> 1) Then
                MsgBox "El Cliente pertenece a la data de la [Campaña Automático], pero falta imprimir la carta y concretar la visita para poder continuar.  ", vbInformation, "Aviso"
                Exit Sub
            End If
            
            If (oRsC!bImpreso = 1 And oRsC!nIdEstadoVisita <> 3) Then
                MsgBox "El Cliente pertenece a la data de la [Campaña Automático], pero falta concretar la visita para poder continuar ", vbInformation, "Aviso"
                Exit Sub
                End If
            Else
                MsgBox "El Cliente no pertenece a la data de la [Campaña Automático], no podrá registrarlo con esta campaña.  ", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    Set oDCreditos = Nothing
    Set oRsC = Nothing
    'End If 'COMENTADO POR ARLO20171226
    'ARLO 20170911


'**ARLO2017113 INICIO ERS070 - 2017
If Not bRefinanciar Then '**ARLO20180317 ERS070 - 2017 --ANEXO 02
    Set oDCreditos = New COMDCredito.DCOMCreditos
    Set rsRCC = oDCreditos.ObtenerCalificacionRCC(Me.lblCodigo.Caption, "")

'**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000069", Trim(Right(cmbSubProducto.Text, 10))) Then
'**ARLO20180712 ERS042 - 2018
   ' If ((CInt(Trim(Right(cmbSubProducto.Text, 5)))) <> 704) Then
    '**ARLO20180528 COMENTADO - INICIO
    '    For i = 1 To UBound(fvListaCompraDeuda)
    '    If fvListaCompraDeuda(i).bRecompra Then
    '        Set rsCrediCancelado = oDCreditos.obtenerCreditosCancelados(Me.lblCodigo.Caption)
    '        If Not (rsCrediCancelado.EOF And rsCrediCancelado.BOF) Then 'ARLO20182203
    '            If DateDiff("M", rsCrediCancelado!dFechaCancelado, gdFecSis) >= 3 Then 'ARLO20180411
    '              MsgBox "El Cliente no cuenta con crédito cancelado hace 3 meses.", vbInformation, "Aviso"
    '              Exit Sub
    '            End If
    '        End If
    '    End If
    '    Next
    '**ARLO20180528 COMENTADO - INICIO
        
        If (nDestino = 14) Then
            If Not (rsRCC.EOF And rsRCC.BOF) Then
                    If (rsRCC!Calif_0 <> 100) Then
                        MsgBox "El Cliente no tiene calificación 100% normal. ", vbInformation, "Aviso"
                        Exit Sub
                    End If
            End If
        End If
    End If
    If (nDestino = 14) Then 'ARLO20180323
        If Not (rsRCC.EOF And rsRCC.BOF) Then
                nCantCompraIFIS = rsRCC!Can_Ents - UBound(fvListaCompraDeuda)
                If (nCantCompraIFIS + 1) > 3 Then '**ARLO20180317 ERS070 - 2017 --ANEXO 02
                    MsgBox "El Cliente no cumple con los requisitos de compra de deuda, máximo debe contar con 3 IFIS" & Chr(13) & _
                    "(incluyendo Caja Maynas) después de la compra.", vbInformation, "Alerta"
                    Exit Sub
                End If
        End If
    End If 'ARLO20180323
 End If '**ARLO20180317 ERS070 - 2017 --ANEXO 02
'**ARLO2017113 FIN ERS070 - 2017

'Agrego JOEP20190919 ERS042 CP-2018
If Not CP_validadConfiguracion Then
    Exit Sub
End If
'Agrego JOEP20190919 ERS042 CP-2018

'Comento JOEP20190919 ERS042 CP-2018
'Set oRsA = objCreditoDA.RecuperaProductoCredicioActivo(Trim(Right(cmbSubProducto.Text, 10)), Trim(Right(cmbCondicionOtra.Text, 10)))
'If Not (oRsA.BOF Or oRsA.EOF) Then
'        'Titular************************************************************
'        oRelPersCred.IniciarMatriz
'        nEdad = EdadPersona(oRelPersCred.ObtenerMatrizRelacionesRelacion(20, 2), gdFecSis)
'        Set oRsA = objCreditoDA.RecuperaAutorizacionCreditos(lblCodigo.Caption, Trim(Right(cmbSubProducto.Text, 10)), Trim(Right(cmbCondicionOtra.Text, 10)), nEdad, 1, CInt(Round((spnCuotas.valor * spnPlazo.valor) / 30, 0)), ActXCtaCred.Age, CInt(Right(cmbDestCred.Text, 3)), Trim(Right(cmbMoneda.Text, 3)), val(txtMontoSol.Text), Trim(Right(cmbCondicion.Text, 5)), 20) 'JOEP20181015 Acta 158-2018 Trim(Right(cmbCondicion.Text, 5)),20
'        If Not (oRsA.BOF Or oRsA.EOF) Then
'            If oRsA!nPase = 0 Then
'                MsgBox "El Titular - crédito no cumple con los requisitos para una solicitud, ver " & IIf(IsNull(oRsA!cTpoDesc), "", oRsA!cTpoDesc), vbInformation, "Aviso"
'                Exit Sub
'            End If
'        Else
'            MsgBox "El Titular - crédito no cumple con los requisitos para una solicitud", vbInformation, "Aviso"
'            Exit Sub
'        End If
'        'Conyugue************************************************************
'        oRelPersCred.IniciarMatriz
'        sPersonaCodCony = oRelPersCred.ObtenerMatrizRelacionesRelacion(21, 1)
'        If Trim(sPersonaCodCony) <> "" Then
'            nEdad = EdadPersona(oRelPersCred.ObtenerMatrizRelacionesRelacion(21, 2), gdFecSis)
'            'Set oRsA = objCreditoDA.RecuperaAutorizacionCreditos(sPersonaCodCony, Trim(Right(cmbSubProducto.Text, 10)), Trim(Right(cmbCondicionOtra.Text, 10)), nEdad, 1, CInt(Round((spnCuotas.valor * spnPlazo.valor) / 30, 0)), ActXCtaCred.Age, CInt(Right(cmbDestCred.Text, 3)), Trim(Right(cmbMoneda.Text, 3)), val(txtMontoSol.Text)) 'Comentado JOEP20180202
'             Set oRsA = objCreditoDA.RecuperaAutorizacionCreditos(sPersonaCodCony, Trim(Right(cmbSubProducto.Text, 10)), Trim(Right(cmbCondicionOtra.Text, 10)), nEdad, 1, CInt(Round((spnCuotas.valor * spnPlazo.valor) / 30, 0)), ActXCtaCred.Age, CInt(Right(cmbDestCred.Text, 3)), Trim(Right(cmbMoneda.Text, 3)), val(txtMontoSol.Text), Trim(Right(cmbCondicion.Text, 5)), 21) 'JOEP20180202 Acta 021-2018 - 'JOEP20181015 Acta 158-2018 21
'            If Not (oRsA.BOF Or oRsA.EOF) Then
'                If oRsA!nPase = 0 Then
'                    MsgBox "El Cónyuge  del crédito no cumple con los requisitos para una solicitud," & IIf(IsNull(oRsA!cTpoDesc), "", oRsA!cTpoDesc), vbInformation, "Aviso"
'                    Exit Sub
'                End If
'            Else
'                MsgBox "El Cónyuge  del crédito no cumple con los requisitos para una solicitud", vbInformation, "Aviso"
'                Exit Sub
'            End If
'        End If
'        'Codeudor************************************************************
'        oRelPersCred.IniciarMatriz
'        sPersonaCodCony = oRelPersCred.ObtenerMatrizRelacionesRelacion(22, 1)
'        If Trim(sPersonaCodCony) <> "" Then
'            nEdad = EdadPersona(oRelPersCred.ObtenerMatrizRelacionesRelacion(22, 2), gdFecSis)
'            'Set oRsA = objCreditoDA.RecuperaAutorizacionCreditos(sPersonaCodCony, Trim(Right(cmbSubProducto.Text, 10)), Trim(Right(cmbCondicionOtra.Text, 10)), nEdad, 1, CInt(Round((spnCuotas.valor * spnPlazo.valor) / 30, 0)), ActXCtaCred.Age, CInt(Right(cmbDestCred.Text, 3)), Trim(Right(cmbMoneda.Text, 3)), val(txtMontoSol.Text))'Comentado JOEP20180202
'             Set oRsA = objCreditoDA.RecuperaAutorizacionCreditos(sPersonaCodCony, Trim(Right(cmbSubProducto.Text, 10)), Trim(Right(cmbCondicionOtra.Text, 10)), nEdad, 1, CInt(Round((spnCuotas.valor * spnPlazo.valor) / 30, 0)), ActXCtaCred.Age, CInt(Right(cmbDestCred.Text, 3)), Trim(Right(cmbMoneda.Text, 3)), val(txtMontoSol.Text), Trim(Right(cmbCondicion.Text, 5)), 22) 'JOEP20180202 Acta 021-2018 - 'JOEP20181015 Acta 158-2018 22
'            If Not (oRsA.BOF Or oRsA.EOF) Then
'                If oRsA!nPase = 0 Then
'                    MsgBox "El Codeudor del crédito no cumple con los requisitos para una solicitud, ver " & IIf(IsNull(oRsA!cTpoDesc), "", oRsA!cTpoDesc), vbInformation, "Aviso"
'                    Exit Sub
'                End If
'            Else
'                MsgBox "El Cónyuge  del crédito no cumple con los requisitos para una solicitud", vbInformation, "Aviso"
'                Exit Sub
'            End If
'        End If
'        'FRHU 20160615 ERS002-2016
'        lsListaPrdPersRelac = Trim(LeeConstanteSist(gConstSistIntervinientesDelCredito))
'        lsListaCondCred = Trim(LeeConstanteSist(gConstSistCondicionesDelCreditoAvalidar))
'        lsCondCred = Trim(Right(cmbCondicion.Text, 5))
'        If InStr(1, lsListaCondCred, lsCondCred) > 0 Then
'            For lnFila = 1 To ListaRelacion.ListItems.count
'                lsPrdPersRelac = Trim(ListaRelacion.ListItems(lnFila).ListSubItems(3))
'                If InStr(1, lsListaPrdPersRelac, lsPrdPersRelac) > 0 Then
'                    lsPersCod = Trim(ListaRelacion.ListItems(lnFila).ListSubItems(2))
'                    'Set oRsRel = objCreditoDA.ValidarCondicionesDeIntervinientesEnElCredito(psNuevaCta, lsPersCod, Trim(Right(cmbCondicion.Text, 5)))
'                    'Set oRsRel = objCreditoDA.ValidarCondicionesDeIntervinientesEnElCredito(psNuevaCta, lsPersCod, Trim(Right(cmbCondicion.Text, 5)), gdFecSis, gsCodUser, gsCodAge) 'FRHU 20160815 ANEXO002 ERS002-2016
'                    Set oRsRel = objCreditoDA.ValidarCondicionesDeIntervinientesEnElCredito(psNuevaCta, lsPersCod, Trim(Right(cmbCondicion.Text, 5)), gdFecSis, gsCodUser, gsCodAge, Trim(Right(cmbSubProducto.Text, 3))) 'FRHU 20160815 Observación
'                    If Not (oRsRel.BOF And oRsRel.EOF) Then
'                        If oRsRel!nPase = 0 Then
'                            MsgBox IIf(IsNull(oRsRel!cTpoDesc), "", oRsRel!cTpoDesc), vbInformation, "Aviso"
'                            Exit Sub
'                        End If
'                    End If
'                End If
'            Next lnFila
'        End If
'        'FIN FRHU
'Else
'    MsgBox "No existe configuración de este sub producto y la campaña", vbInformation, "Aviso"
'    Exit Sub
'End If
''ENDALPA 20140708********************************************
'Comento JOEP20190919 ERS042 CP-2018

'Agrego JOEP20190919 ERS042 CP-2018
    'FRHU 20160615 ERS002-2016
        lsListaPrdPersRelac = Trim(LeeConstanteSist(gConstSistIntervinientesDelCredito))
        lsListaCondCred = Trim(LeeConstanteSist(gConstSistCondicionesDelCreditoAvalidar))
        lsCondCred = Trim(Right(cmbCondicion.Text, 5))
        If InStr(1, lsListaCondCred, lsCondCred) > 0 Then
            For lnFila = 1 To ListaRelacion.ListItems.count
                lsPrdPersRelac = Trim(ListaRelacion.ListItems(lnFila).ListSubItems(3))
                If InStr(1, lsListaPrdPersRelac, lsPrdPersRelac) > 0 Then
                    lsPersCod = Trim(ListaRelacion.ListItems(lnFila).ListSubItems(2))
                    Set oRsRel = objCreditoDA.ValidarCondicionesDeIntervinientesEnElCredito(psNuevaCta, lsPersCod, Trim(Right(cmbCondicion.Text, 5)), gdFecSis, gsCodUser, gsCodAge, Trim(Right(cmbSubProducto.Text, 3))) 'FRHU 20160815 Observación
                    If Not (oRsRel.BOF And oRsRel.EOF) Then
                        If oRsRel!nPase = 0 Then
                            MsgBox IIf(IsNull(oRsRel!cTpoDesc), "", oRsRel!cTpoDesc), vbInformation, "Aviso"
                            Exit Sub
                        End If
                    End If
                End If
            Next lnFila
        End If
    'FIN FRHU
'Agrego JOEP20190919 ERS042 CP-2018

'***  Validar Credito Vinculado --- 28-09-2006 ***
Dim oPers  As COMDPersona.UCOMPersona
Set oPers = New COMDPersona.UCOMPersona
    If oPers.fgVerificaEmpleado(lblCodigo.Caption) Then
        MsgBox "Este es un Crédito Vinculado...Empleado de la Caja", vbInformation, "Aviso"
    End If
Set oPers = Nothing
'*****************APRI20170630 TI-ERS025-2017*****************
Dim rs As ADODB.Recordset
Set oPers = New COMDPersona.UCOMPersona
Set rs = oPers.ObtenerVinculadoRiesgoUnico(lblCodigo.Caption, "", 0)

    If Not (rs.BOF And rs.EOF) Then
        If rs.RecordCount = 1 Then
            If rs!nTotal = 1 Then
                If MsgBox("El vinculado " & rs!cPersNombre & " tiene un crédito que se encuentra en " & rs!cEstado & ". ¿Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                Exit Sub
                End If
            Else
                If MsgBox("El vinculado " & rs!cPersNombre & " tiene " & rs!nTotal & " créditos que se encuentran en mora. ¿Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                Exit Sub
                End If
            End If
        ElseIf rs.RecordCount > 1 Then
            If MsgBox("El cliente tiene vinculados en persona que se encuentra en mora. ¿Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
Set oPers = Nothing
'*****************END APRI20170630 ***************************
Set oPers = New COMDPersona.UCOMPersona
    If oPers.fgVerificaEmpleadoVincualdo(lblCodigo.Caption) Then
        MsgBox "Este es un Crédito Vinculado...Pariente de Empleado", vbInformation, "Aviso"
    End If
Set oPers = Nothing


    '*** PEAC 20100512
    If cmdEjecutar = 1 Then  'valida solo cuando es nueva solicitud
        Set oPers = New COMDPersona.UCOMPersona
            If oPers.fgVerificaCredAnalistaCliente(lblCodigo.Caption, Trim(Right(cmbAnalista.Text, 15))) Then
                MsgBox "Analista no puede digitar crédito porque existe un crédito pendiente.", vbInformation, "Aviso"
                Exit Sub
            End If
        Set oPers = Nothing
    End If
    '*** FIN PEAC
    
    '*** PEAC 20080811 ******************************************************
    For i = 0 To UBound(oRelPersCred.ObtenerMatrizRelaciones) - 1
        '*** el codigo de operacio falta definir para la solicitud por miestras se puso 401110
        lbResultadoVisto = loVistoElectronico.Inicio(1, gColAperturaSolicitudCred, oRelPersCred.ObtenerMatrizRelaciones(i, 0))
        If Not lbResultadoVisto Then
            Exit Sub
        End If
        
        If CInt(Trim(Right(cmbSubProducto.Text, 10))) = 515 Then
            
        End If
        
        
    Next i
    '*** FIN PEAC ************************************************************
    '*** BRGO 20111102 *************************************
'**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000071", Trim(Right(cmbSubProducto.Text, 10))) Then
'**ARLO20180712 ERS042 - 2018
    'If CInt(Trim(Right(cmbSubProducto.Text, 10))) = "517" Then
        If spnCuotas.valor > 60 Then
            MsgBox "El Sub Producto seleccionado no permite el registro de cuotas mayores a 60"
            spnCuotas.SetFocus
            Exit Sub
        End If
    End If
    '*** End BRGO
    Dim oDCredi As COMDCredito.DCOMCreditos  'BY ARLO20171226
    Set oDCredi = New COMDCredito.DCOMCreditos 'BY ARLO20171226
    If Not (oDCredi.VerificaCampania(nIdCampana)) Then 'BY ARLO20171127
    'If (nIdCampana <> 116) Then 'ARLO20170818 'BY ARLO20171127 COMENT
        'JUEZ 20140603 *******************************************************
        Set oDPersGen = New COMDPersona.DCOMPersGeneral
        If oDPersGen.VerificaSuperaUmbralSectorEcon(lblCodigo.Caption, CDbl(txtMontoSol.Text), CInt(Trim(Right(cmbMoneda.Text, 2)))) Then
            If MsgBox("El crédito supera el porcentaje máximo por sector; se podrá continuar, pero no se podrá sugerir si no se tiene la autorización de Riesgos, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                bSolicitaAutSectorEcon = True
            Else
                Exit Sub
            End If
        End If
        Set oDPersGen = Nothing
        'END JUEZ ************************************************************
    End If 'ARLO20170818
    
    'JOEP ERS047 20170901 *******************************************************
    Set oCredito = New COMDCredito.DCOMCredito
    If oCredito.VerificaSuperaUmbralZonaGeog(gsCodAge, CDbl(txtMontoSol.Text), CInt(Trim(Right(cmbMoneda.Text, 2)))) Then
        If MsgBox("El Crédito supera el porcentaje máximo por Zona Geográfica; se podrá continuar, pero no se podrá sugerir si no se tiene la autorización de Riesgos, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            bSolicitaAutZonaGeog = True
        Else
            Exit Sub
        End If
    End If
    Set oCredito = Nothing
    
     Set oCredito = New COMDCredito.DCOMCredito
    If oCredito.VerificaSuperaUmbralTpCredito(Trim(Right(cmbSubProducto.Text, 5)), CDbl(txtMontoSol.Text), CInt(Trim(Right(cmbMoneda.Text, 2)))) Then
        If MsgBox("El crédito supera el porcentaje máximo por Tipo de Credito; se podrá continuar, pero no se podrá sugerir si no se tiene la autorización de Riesgos, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            bSolicitaAutTpCredito = True
        Else
            Exit Sub
        End If
    End If
    Set oCredito = Nothing
    'END JOEP ERS047 ************************************************************
    
        'AMDO 20130702 TI-ERS063-2013 ****************************************************
            Dim oDPersonaAct As COMDPersona.DCOMPersona
            Dim conta As Integer
            Dim sPersCod As String
            Set oDPersonaAct = New COMDPersona.DCOMPersona
            For conta = 1 To ListaRelacion.ListItems.count
            sPersCod = Trim(ListaRelacion.ListItems(conta).ListSubItems(2))
                            If oDPersonaAct.VerificaExisteSolicitudDatos(sPersCod) Then
                                MsgBox Trim("SE SOLICITA DATOS DEL CLIENTE: " & ListaRelacion.ListItems(conta).Text) & "." & Chr(10), vbInformation, "Aviso"
                                Call frmActInfContacto.Inicio(sPersCod)
                            End If
            Next conta
    'AMDO FIN ********************************************************************************
    'FRHU 20140402 ERS026-2014 RQ14129
    Dim oCred As New COMDCredito.DCOMCredito
    Dim oRs As ADODB.Recordset
    Dim nConta As Integer
    Dim sCodPers As String
    Dim sNomPers As String
    sCodPers = ""
    For nConta = 1 To ListaRelacion.ListItems.count
        sCodPers = sCodPers & Trim(ListaRelacion.ListItems(nConta).ListSubItems(2)) & ","
    Next nConta
    Set oRs = oCred.ObtenerUltimaActualizacion(Mid(sCodPers, 1, Len(sCodPers) - 1))
    sNomPers = ""
    If Not oRs.EOF And Not oRs.BOF Then
        Do While Not oRs.EOF
            sNomPers = sNomPers & oRs!cPersNombre & vbNewLine
            oRs.MoveNext
        Loop
        MsgBox "Nombre de los clientes que requieren actualizarse: " & vbNewLine & _
               sNomPers & vbNewLine & _
               "No podrá aprobarse el crédito hasta que se hayan realizado las actualizaciones.", vbInformation
    End If
    'FIN FRHU 20140402
    'EJVG20160712 *** Exposición Este Crédito
    fbEliminarEvaluacion = False
    nFormatoEliminado = -1 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
    nFormato_NEW = -1 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
    If cmdEjecutar = 2 Then
        GenerarDataExposicionEsteCredito ActXCtaCred.NroCuenta, CDbl(txtMontoSol.Text), fnMontoExpEsteCred_NEW 'Seteamos el valor de la nueva exposición
        If NecesitaFormatoEvaluacion(ActXCtaCred.NroCuenta, 2000, CInt(Right(cmbProductoCMACM.Text, 3)), CInt(Right(cmbSubProducto, 3)), fnMontoExpEsteCred_NEW, fbEliminarEvaluacion, nFormatoEliminado, nFormato_NEW) Then
        'LUCV20181220 Agregó: nFormatoEliminado, nFormato_NEW, Anexo01 de Acta 199-2018
            Exit Sub
        End If
    End If
    'END EJVG *******
    '**ARLO20200716 Incidente
    Set objCreditoDA = New COMDCredito.DCOMCredito
    If bRefinanciar Then
        If IsArray(MatCredRef) Then
            If UBound(MatCredRef) > 0 Then
                For i = 0 To UBound(MatCredRef) - 1
                    Set oRsA = objCreditoDA.ValidadCreditoRefinanciadoTitular(MatCredRef(i, 0), Trim(Me.lblCodigo.Caption))
                    If Not (oRsA.EOF And oRsA.BOF) Then
                        If oRsA!cMensaje <> "" Then
                            MsgBox oRsA!cMensaje, vbInformation, "Aviso"
                            Exit Sub
                        End If
                    End If
                    Set oRsA = Nothing
                Next i
            End If
        End If
    End If
    Set objCreditoDA = Nothing
    '**END ARLO
    'On Error GoTo ErrorCmdGrabar_Click
    If MsgBox("Se va a Grabar los Datos, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub 'vbYes Then
            'FRHU 20140514 Observacion
            'If Not ValidaDatos Then
               'Exit Sub
            'End If
            'FIN FRHU 20140514
        ' CMACICA_CSTS - 10/11/2003 --------------------------------------------------------------------
        If bRefinanciar Then
           bRefinanciarSustituir = True
           bSustiDeudor = False
           
           MatCredRefSust = MatCredRef
        Else
            If bSustituirDeudor Then
               bRefinanciarSustituir = True
               MatCredRefSust = MatCredSust
               bSustiDeudor = True
            Else
               bRefinanciarSustituir = False
               bSustiDeudor = False
            End If
        End If
        '------------------------------------------------------------------------------------------------
        
        'Se Agrego para manejar el tema de las relaciones
        MatCredRelaciones = oRelPersCred.ObtenerMatrizRelaciones
        
        'ARCV 30-12-2006
        Dim MatFtesSel As Variant
        'Dim i As Integer

        'ReDim MatFtesSel(UBound(MatFuentes), 2)
            
'        For i = 0 To UBound(MatFuentes) - 1
'            MatFtesSel(i, 0) = oPersona.ObtenerFteIngcNumFuente(MatFuentes(i))
'            MatFtesSel(i, 1) = MatFteFecEval(i)
'        Next i
        If Trim(cmbSubProducto.Text) = "" Then
            MsgBox "Se debe asignar el sub producto", vbInformation, "AVISO"
        End If
        
        lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
        If cmdEjecutar = 1 Then
        
            If rsAmpliado Is Nothing And bAmpliacion = True Then
                MsgBox "Se debe asignar el credito que se debe ampliar", vbInformation, "AVISO"
                Exit Sub
            End If
            
            'FRHU 20140424 TI-ERS015-2014
            If bAmpliacion = True Then
                Dim pnMontoITf As Double
                Dim oITF As COMDConstSistema.FCOMITF
                Set oITF = New COMDConstSistema.FCOMITF
                oITF.fgITFParametros
                pnMontoITf = oITF.fgITFDesembolso(nMontoTotal)
                Set oITF = Nothing
                nMontoTotal = nMontoTotal + pnMontoITf
                If Me.txtMontoSol.Text < nMontoTotal Then
                    MsgBox "El Monto Solicitado en la ampliacion de credito, debe ser mayor a el monto total de los creditos a cancelar", vbInformation
                    Exit Sub
                End If
            End If
            'FIN FRHU 20140424
            
            'ARLO 20181126 ERS068 - 2018*****************************************
            Dim sMensaje As String
            If bAmpliacion Then
                If chkAutAmpliacion.value = 0 Then  'ARLO20190403
                    Dim rsAmpliadoRefinan, rsSMS As ADODB.Recordset
                    Dim cCtaCodRefinan As String
                    Set rsAmpliadoRefinan = rsAmpliado.Clone
                    rsAmpliadoRefinan.MoveFirst
                    Do While Not rsAmpliadoRefinan.EOF
                        cCtaCodRefinan = cCtaCodRefinan + "," + rsAmpliadoRefinan("cCtaCod")
                        rsAmpliadoRefinan.MoveNext
                    Loop
                    Set oCredito = New COMDCredito.DCOMCredito
                    Set rsSMS = oCredito.ValidaCreditoRefinanciarAmpliado(cCtaCodRefinan, val(Me.txtMontoSol.Text))
                    If Not (rsSMS.EOF And rsSMS.BOF) Then
                        If rsSMS!sMensaje <> "" Then
                            MsgBox rsSMS!sMensaje, vbInformation, "Aviso"
                            Exit Sub
                        End If
                    End If
                    Set oCredito = Nothing
                End If
            End If
            'ARLO 20181126 ERS068 - 2018*****************************************
            
    '**ARLO20190308 ERS068-2018
    Set objCreditoDA = New COMDCredito.DCOMCredito
    If bRefinanciar Then
        If IsArray(MatCredRef) Then
            If UBound(MatCredRef) > 0 Then
                For i = 0 To UBound(MatCredRef) - 1
                    Set oRsA = objCreditoDA.getObtieneCuotasPagadasCreditoRefinan(MatCredRef(i, 0), Trim(Right(cmbSubProducto.Text, 10)))
                    If Not (oRsA.EOF And oRsA.BOF) Then
                        If oRsA!cMensaje <> "" Then
                            Call objCreditoDA.InsertarVBCuotas6MinimasRiesgos(lblCodigo, gsCodAge, CInt(Trim(Right(cmbMoneda.Text, 2))), CDbl(txtMontoSol.Text), oRsA!nCuotasPagadas, oRsA!nTotalCuotas, lcMovNro)
                            MsgBox oRsA!cMensaje, vbInformation, "Aviso"
                            Exit Sub
                        End If
                    End If
                    Set oRsA = Nothing
                Next i
            End If
                Set oRsA = objCreditoDA.getObtieneMensajeVB6Cuotas(lblCodigo)
                If Not (oRsA.EOF And oRsA.BOF) Then
                    If oRsA!cMensaje <> "" Then
                        MsgBox "Ciente cuenta con V.B de Riesgo : " + oRsA!cMensaje, vbInformation, "Aviso" 'arlo20191215
                    End If
                End If
                Set oRsA = Nothing
        End If
    End If
    Set objCreditoDA = Nothing
    '**END ARLO

            
            'JUEZ 20160509 ******************************************************
            If bAmpliacion Then
                If chkAutAmpliacion.value = 0 Then
                    Dim nParamPorcAmp As Double, nMontoTotalValida As Double, nMontoPorcMontoTotal As Double
                    Dim oDParam As COMDCredito.DCOMParametro
                    Set oDParam = New COMDCredito.DCOMParametro
                        nParamPorcAmp = oDParam.RecuperaValorParametro(3503)
                    Set oDParam = Nothing
                    
                    rsAmpliado.MoveFirst
                    Do While Not rsAmpliado.EOF
                        Set oCredito = New COMDCredito.DCOMCredito
                            Set RsValidaAmp = oCredito.ValidaCreditoAmpliar(rsAmpliado("cCtaCod"))
                        Set oCredito = Nothing
                        
                        If Not RsValidaAmp.BOF And Not RsValidaAmp.EOF Then
                            If RsValidaAmp!sMensaje <> "" Then
                                MsgBox "El crédito " & rsAmpliado("cCtaCod") & " no cumple con lo requerido para ampliación: " & CStr(RsValidaAmp!sMensaje), vbInformation, "Aviso"
                                
                                nMontoTotalValida = nMontoTotal - pnMontoITf
                                nMontoPorcMontoTotal = nMontoTotal * (nParamPorcAmp / 100)
                                If txtMontoSol.Text <= nMontoTotalValida + nMontoPorcMontoTotal Then
                                    MsgBox "El monto solicitado debe superar al monto total de los créditos a ampliar en un " & CStr(nParamPorcAmp) & "% adicional, para este caso el monto solicitado debe superar el monto de " & Format(nMontoTotalValida + nMontoPorcMontoTotal, "#,###.00"), vbInformation, "Aviso"
                                    rsAmpliado.MoveFirst
                                    Exit Sub
                                End If
                                
                                rsAmpliado.MoveFirst
                                Exit Sub
                            End If
                        End If
                        rsAmpliado.MoveNext
                    Loop
                Else
                    lsComentario = frmCredSolicAutAmpComent.ObtenerComentario
                    If Trim(lsComentario) = "" Then
                        MsgBox "Deber ingresar el motivo de la solicitud de exoneración", vbInformation, "Aviso"
                        Exit Sub
                    End If
                End If
            End If
            'END JUEZ ***********************************************************
            
        '    Set oGen = New COMDConstSistema.DCOMGeneral
        '    psNuevaCta = gsCodCMAC & oGen.GeneraNuevaCuenta(gsCodAge, CInt(Mid(Trim(Right(cmbTipoCred.Text, 5)), 1, 1) & Right(cmbSubTipo.Text, 2)), CInt(Trim(Right(cmbMoneda.Text, 2))))
        '    sMovAct = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        '    Set oGen = Nothing
            
'            If bAmpliacion = True Then
'                If CDbl(txtMontoSol.Text) < CDbl(rsAmpliado(3)) Then
'                     MsgBox "El monto solicitado debe ser mayor que el monto de credito ampliado", vbInformation, "AVISO"
'                     Exit Sub
'                End If
'            End If
            
'            If bAmpliacion = True Then
'                If ValidaMontoAmpliacion(CDbl(rsAmpliado(3)), CInt(Mid(rsAmpliado(0), 9, 1)), Val(Me.txtMontoSol), Mid(psNuevaCta, 9, 1)) = False Then
'                    MsgBox "El monto del nuevo credito no cubre el monto del credito ampliado", vbInformation, "AVISO"
'                    Exit Sub
'                End If
'
'                If ListaValidaMontoAmpliacion(rsAmpliado, txtMontoSol.Text, Mid(psNuevaCta, 9, 1)) = False Then
'                    MsgBox "El monto del nuevo credito no cubre el monto del credito ampliado", vbInformation, "AVISO"
'                    Exit Sub
'                End If
'            End If
            
            Set oCredito = New COMDCredito.DCOMCredito
                        
            nCampanaCod = Trim(Right(cmbCondicionOtra.Text, 5))
            nCasaCod = Trim(Right(cmbCondicionOtra2.Text, 5)) 'MADM 20100719
            
            'Call oCredito.NuevaSolicitud(MatCredRelaciones, psNuevaCta, Trim(Right(cmbCondicion.Text, 5)), _
                    Trim(Right(cmbCondicionOtra.Text, 5)), oPersona.ObtenerFteIngcNumFuente(cmbFuentes.ListIndex), CDbl(txtMontoSol.Text), CInt(spnCuotas.Valor), CInt(spnPlazo.Valor), Trim(Right(cmbDestCred.Text, 5)), Trim(Right(cmbAnalista.Text, 15)), _
                    sCodInstitucion, sCodModular, gdFecSis, sMovAct, bRefinanciarSustituir, CInt(Trim(Right(LblCondProd.Caption, 5))), oPersona.ObtenerFteIngFecEval(cmbFuentes.ListIndex, IIf(oPersona.ObtenerFteIngIngresoTipo(cmbFuentes.ListIndex) = gPersFteIngresoTipoDependiente, oPersona.ObtenerFteIngIngresoNumeroFteDep(cmbFuentes.ListIndex) - 1, oPersona.ObtenerFteIngIngresoNumeroFteIndep(cmbFuentes.ListIndex) - 1)), MatCredRefSust, IIf(ChkCap.value = 1, True, False), bSustiDeudor, bAmpliacion, rsAmpliado, nCampanaCod, _
                    gdFecSis, gsCodUser, gsCodCMAC, gsCodAge, CInt(Mid(Trim(Right(cmbTipoCred.Text, 5)), 1, 1) & Right(cmbSubTipo.Text, 2)), CInt(Trim(Right(cmbMoneda.Text, 2))), fsCtaCod, sCargo, sCARBEN, sT_Plani)
                    'Trim(Right(CmbInstitucion.Text, 15)), cmbModular.Text, gdFecSis, sMovAct, bRefinanciar, CInt(Trim(Right(LblCondProd.Caption, 5))), oPersona.ObtenerFteIngFecEval(cmbFuentes.ListIndex, IIf(oPersona.ObtenerFteIngIngresoTipo(cmbFuentes.ListIndex) = gPersFteIngresoTipoDependiente, oPersona.ObtenerFteIngIngresoNumeroFteDep(cmbFuentes.ListIndex) - 1, oPersona.ObtenerFteIngIngresoNumeroFteIndep(cmbFuentes.ListIndex) - 1)), MatCredRef, IIf(ChkCap.value = 1, True, False))
            
'            Call oCredito.NuevaSolicitud(MatCredRelaciones, psNuevaCta, Trim(Right(cmbCondicion.Text, 5)), _
'                    Trim(Right(cmbCondicionOtra.Text, 5)), MatFtesSel, CDbl(txtMontoSol.Text), CInt(spnCuotas.Valor), CInt(spnPlazo.Valor), Trim(Right(cmbDestCred.Text, 5)), Trim(Right(cmbAnalista.Text, 15)), _
'                    sCodInstitucion, sCodModular, gdFecSis, sMovAct, bRefinanciarSustituir, CInt(Trim(Right(LblCondProd.Caption, 5))), MatCredRefSust, IIf(ChkCap.value = 1, True, False), bSustiDeudor, bAmpliacion, rsAmpliado, nCampanaCod, _
'                    gdFecSis, gsCodUser, gsCodCMAC, gsCodAge, CInt(Mid(Trim(Right(cmbTipoCred.Text, 5)), 1, 1) & Right(cmbSubTipo.Text, 2)), CInt(Trim(Right(cmbMoneda.Text, 2))), fsCtaCod, sCargo, sCARBEN, sT_Plani)
            
            'By capi 28102008
'            Call oCredito.NuevaSolicitud(MatCredRelaciones, psNuevaCta, Trim(Right(cmbCondicion.Text, 5)), _
'                    Trim(Right(cmbCondicionOtra.Text, 5)), CDbl(txtMontoSol.Text), CInt(spnCuotas.Valor), CInt(spnPlazo.Valor), Trim(Right(cmbDestCred.Text, 5)), Trim(Right(cmbAnalista.Text, 15)), _
'                    sCodInstitucion, sCodModular, gdFecSis, sMovAct, bRefinanciarSustituir, CInt(Trim(Right(LblCondProd.Caption, 5))), MatCredRefSust, IIf(ChkCap.value = 1, True, False), bSustiDeudor, bAmpliacion, rsAmpliado, nCampanaCod, _
'                    gdFecSis, gsCodUser, gsCodCMAC, gsCodAge, CInt(Mid(Trim(Right(cmbTipoCred.Text, 5)), 1, 1) & Right(cmbSubTipo.Text, 2)), CInt(Trim(Right(cmbMoneda.Text, 2))), fsCtaCod, sCargo, sCARBEN, sT_Plani)
        
        'PEAC Modificado por ALPA 20090104******************************************************
'             Call oCredito.NuevaSolicitud(MatCredRelaciones, psNuevaCta, Trim(Right(cmbCondicion.Text, 5)), _
'                    Trim(Right(cmbCondicionOtra.Text, 5)), CDbl(txtMontoSol.Text), CInt(spnCuotas.Valor), CInt(spnPlazo.Valor), Trim(Right(cmbDestCred.Text, 5)), Trim(Right(cmbAnalista.Text, 15)), _
'                    sCodInstitucion, sCodModular, gdFecSis, sMovAct, bRefinanciarSustituir, CInt(Trim(Right(LblCondProd.Caption, 5))), MatCredRefSust, IIf(ChkCap.value = 1, True, False), bSustiDeudor, bAmpliacion, rsAmpliado, nCampanaCod, _
'                    gdFecSis, gsCodUser, gsCodCMAC, gsCodAge, CInt(Mid(Trim(Right(cmbTipoCred.Text, 5)), 1, 1) & Right(cmbSubTipo.Text, 2)), CInt(Trim(Right(cmbMoneda.Text, 2))), fsCtaCod, sCargo, sCARBEN, sT_Plani, CInt(Trim(Right(cmbMotivoRef.Text, 5))))
                    
            Call oCredito.NuevaSolicitud(MatCredRelaciones, psNuevaCta, Trim(Right(cmbCondicion.Text, 5)), _
                    Trim(Right(cmbCondicionOtra.Text, 5)), CDbl(txtMontoSol.Text), CInt(spnCuotas.valor), CInt(spnPlazo.valor), Trim(Right(cmbDestCred.Text, 5)), Trim(Right(cmbAnalista.Text, 15)), _
                    sCodInstitucion, sCodModular, gdFecSis, sMovAct, bRefinanciarSustituir, CInt(Trim(Right(LblCondProd.Caption, 5))), MatCredRefSust, IIf(ChkCap.value = 1, True, False), bSustiDeudor, bAmpliacion, rsAmpliado, nCampanaCod, _
                    gdFecSis, gsCodUser, gsCodCMAC, gsCodAge, CInt(Mid(Trim(Right(cmbProductoCMACM.Text, 5)), 1, 1) & Right(cmbSubProducto.Text, 2)), CInt(Trim(Right(cmbMoneda.Text, 2))), fsCtaCod, sCargo, sCARBEN, sT_Plani, CInt(IIf(Trim(Right(cmbMotivoRef.Text, 5)) = "", 0, Trim(Right(cmbMotivoRef.Text, 5)))), nCasaCod, Me.txtVendedor.Text, Trim(txtDetalleMotivoRef.Text), IIf(fraPromotor.Visible And cmbPromotor.ListIndex <> -1, fbRegPromotores, False), Trim(Right(cmbPromotor.Text, 15)), fvListaCompraDeuda, fvListaAguaSaneamiento, _
                    nSubDestino, nTpDoc, nTpIngr, nTpInt, fvListaCreditoVerde) 'EAAS 20180807 SEGUN ERS-054-2018 fvListaAguaSaneamiento
'JOEP20190919 ERS042 CP-2018 nSubDestino, nTpDoc, nTpIngr,nTpInt
'EAAS20191401 SEGUN 018-GM-DI_CMACM fvListaCreditoVerde
             'agregado por vapi SEGÙN ERS TI-ERS001-2017
            If bPreSolOperacion Then
                Dim oHojaRuta As COMDCredito.DCOMhojaRuta
                Set oHojaRuta = New COMDCredito.DCOMhojaRuta
                Call oHojaRuta.ActualizarcCtaPreSol(rsPresol!nPresolicitudId, psNuevaCta)
                bPresolAmpliAuto = False
                bPresol = False
                bPreSolOperacion = False
            End If
            'fin agregado por vapi
                    
                    
                    'WIOR 20140509 AGREGO IIf(fraPromotor.Visible, fbRegPromotores, False),Trim(Right(cmbPromotor.Text, 15))
                    'LUCV20170425 Modificó: IIf(fraPromotor.Visible, fbRegPromotores, False)
        '****************************************************************************************
            'MIOL 20130626, RQ13335 Agrego Parametro Me.txtVendedor.Text
            'EJVG20160203 ERS002-2016 agregó fvListaCompraDeuda
            ''** PEAC 20090126
            
            'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, , psNuevaCta, gCodigoCuenta 'RECO20161020 ERS060-2016
            
             'RECO20161020 ERS060-2016*********************************
            Dim oNCOMColocEval As New NCOMColocEval
            'Dim lcMovNro As String 'LUCV20181220 Comentó, Anexo01 de Acta 199-2018
            
            'lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Comentó, Anexo01 de Acta 199-2018
            Set objPista = New COMManejador.Pista 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
            'objPista.InsertarPista 190260, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, , psNuevaCta, gCodigoCuenta 'JOEP22052017 'LUCV20181220 Comentó, Anexo01 de Acta 199-2018
            objPista.InsertarPista gCredRegistrarActualizaSoliCred, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Solicitud del Crédito (Sicmac Negocio)", psNuevaCta, gCodigoCuenta  'JOEP22052017 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
            Set objPista = Nothing 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
            
            If Not ValidaExisteRegProceso(psNuevaCta, gTpoRegCtrlSolicitud) Then
                'objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, , psNuevaCta, gCodigoCuenta ' Comentado JOEP22052017
                Call oNCOMColocEval.insEstadosExpediente(psNuevaCta, "Solicitud de Crédito", lcMovNro, "", "", "", 1, 2000, gTpoRegCtrlSolicitud)
                Set oNCOMColocEval = Nothing
            End If
            'RECO FIN ************************************************
            
            '----------------
            Set oCredito = Nothing
            MsgBox " Nro de Credito : " & psNuevaCta, vbInformation, "Aviso"
            Set MatCredRef = Nothing
            Set MatCredSust = Nothing
            Set MatCredRefSust = Nothing
            
            Unload frmCredRefinanc
        Else
            'MAVM 20090801*************************************************************
            nCampanaCod = Trim(Right(cmbCondicionOtra.Text, 5))
            nCasaCod = Trim(Right(cmbCondicionOtra2.Text, 5)) 'MADM 20100719
            '**************************************************************************
            'sMovAct = GeneraMovNroActualiza(gdFecSis, gsCodUser, Mid(gsCodAge, 1, 3), gsCodAge)
            psNuevaCta = ActXCtaCred.NroCuenta
            Set oCredito = New COMDCredito.DCOMCredito
            
'            Call oCredito.ActualizaSolicitud(MatCredRelaciones, psNuevaCta, Trim(Right(cmbCondicion.Text, 5)), _
'                    Trim(Right(cmbCondicionOtra.Text, 5)), MatFtesSel, CDbl(txtMontoSol.Text), CInt(spnCuotas.Valor), CInt(spnPlazo.Valor), Trim(Right(cmbDestCred.Text, 5)), Trim(Right(cmbAnalista.Text, 15)), _
'                    sCodInstitucion, sCodModular, gdFecSis, sMovAct, bRefinanciarSustituir, CInt(Trim(Right(IIf(Trim(LblCondProd.Caption) = "", "0", LblCondProd.Caption), 5))), MatCredRefSust, IIf(ChkCap.value = 1, True, False), , nCampanaCod, _
'                    gdFecSis, gsCodUser, gsCodCMAC, gsCodAge, fsCtaCod, sCargo, sCARBEN, sT_Plani)
            'WIOR 20120914 ******************************************
            'If nExisteAgeEvalCred = 1 Then
            '    Dim oEvalCred As COMDCredito.DCOMCredito
            '    Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
            '    Dim rsEvalCred As ADODB.Recordset
            '    Dim nFormato As Integer
            '    Dim nFormatoAct As Integer
            '    Dim nMontoEval As Double
            '    Dim nMontoAct As Double
            '    Dim cPrd As String
            '    Dim cSPrd As String
            '    Dim nMin As Double
            '    Dim nMax As Double
            '    Dim nTC As Double
            '
            '    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
            '    Set oEvalCred = New COMDCredito.DCOMCredito
            '
            '    nTC = CDbl(oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoDia))
            '    nMontoAct = CDbl(txtMontoSol.Text)
            '
            '    Set rsEvalCred = oEvalCred.RecuperaSolicitudDatoBasicos(psNuevaCta)
            '    If rsEvalCred.RecordCount > 0 Then
            '        nMontoEval = CDbl(Trim(rsEvalCred!nMonto))
            '        cSPrd = Trim(rsEvalCred!cTpoProdCod)
            '        cPrd = Mid(cSPrd, 1, 1) & "00"
            '        If Mid(psNuevaCta, 9, 1) = "2" Then
            '            nMontoEval = nMontoEval * nTC
            '        End If
            '    End If
            '
            '    If Trim(Right(cmbMoneda, 2)) = "2" Then
            '        nMontoAct = nMontoAct * nTC
            '    End If
            '
            '    nFormato = oEvalCred.AsignarFormato(cPrd, cSPrd, nMontoEval, nMin, nMax)
            '    nFormatoAct = oEvalCred.AsignarFormato(Trim(Right(cmbProductoCMACM.Text, 5)), _
            '                Trim(Right(cmbSubProducto.Text, 5)), nMontoAct)
            '
            '    If nFormato <> nFormatoAct Then
            '        If MsgBox("Este credito ya esta asignado al Formato " & nFormato & ", si modifica la solicitud a este monto se asignará al Formato " & nFormatoAct & _
            '           Chr(10) & "y los datos del Anterior Formato se eliminaran por completo." & _
            '           Chr(10) & _
            '           Chr(10) & "Estas Seguro de continuar con la Actualizacón?", _
            '           vbInformation + vbYesNo, "Aviso") = vbYes Then
            '
            '            Call oEvalCred.EliminaDatosFormato(psNuevaCta)
            '            MsgBox "Datos del Formato " & nFormato & " Eliminados Satisfactoriamente.", vbInformation, "Aviso"
            '        Else
            '            If Mid(psNuevaCta, 9, 1) = "2" Then
            '                nMin = nMin / nTC
            '                nMax = nMax / nTC
            '            End If
            '            MsgBox "Monto Establecidos Para el Formato de Evaluación Actual" & _
            '            Chr(10) & "Monto Minimo: " & Format(nMin, "#0.00") & _
            '            Chr(10) & "Monto Máximo: " & Format(nMax, "#0.00"), _
            '            vbInformation, "Aviso"
            '            Exit Sub
            '        End If
            '    End If
            'End If
            'Set oTipoCam = Nothing
            'Set oEvalCred = Nothing
            ''WIOR FIN ***********************************************
            
            'By capi 28102008
'            Call oCredito.ActualizaSolicitud(MatCredRelaciones, psNuevaCta, Trim(Right(cmbCondicion.Text, 5)), _
'                    Trim(Right(cmbCondicionOtra.Text, 5)), CDbl(txtMontoSol.Text), CInt(spnCuotas.Valor), CInt(spnPlazo.Valor), Trim(Right(cmbDestCred.Text, 5)), Trim(Right(cmbAnalista.Text, 15)), _
'                    sCodInstitucion, sCodModular, gdFecSis, sMovAct, bRefinanciarSustituir, CInt(Trim(Right(IIf(Trim(LblCondProd.Caption) = "", "0", LblCondProd.Caption), 5))), MatCredRefSust, IIf(ChkCap.value = 1, True, False), , nCampanaCod, _
'                    gdFecSis, gsCodUser, gsCodCMAC, gsCodAge, fsCtaCod, sCargo, sCARBEN, sT_Plani)
'           ALPA 20100604 B2************************************************************************
            Call oCredito.ActualizaSolicitud(MatCredRelaciones, psNuevaCta, Trim(Right(cmbCondicion.Text, 5)), _
                     Trim(Right(cmbCondicionOtra.Text, 5)), CDbl(txtMontoSol.Text), CInt(spnCuotas.valor), CInt(spnPlazo.valor), Trim(Right(cmbDestCred.Text, 5)), Trim(Right(cmbAnalista.Text, 15)), _
                    sCodInstitucion, sCodModular, gdFecSis, sMovAct, bRefinanciarSustituir, CInt(Trim(Right(IIf(Trim(LblCondProd.Caption) = "", "0", LblCondProd.Caption), 5))), MatCredRefSust, IIf(ChkCap.value = 1, True, False), , nCampanaCod, _
                    gdFecSis, gsCodUser, gsCodCMAC, gsCodAge, fsCtaCod, sCargo, sCARBEN, sT_Plani, CInt(Trim(Right(IIf(cmbMotivoRef.Text = "", 0, cmbMotivoRef.Text), 5))), CInt(Mid(Trim(Right(cmbProductoCMACM.Text, 5)), 1, 1) & Right(cmbSubProducto.Text, 2)), nCasaCod, Me.txtVendedor.Text, Trim(txtDetalleMotivoRef.Text), IIf(fraPromotor.Visible And cmbPromotor.ListIndex <> -1, fbRegPromotores, False), Trim(Right(cmbPromotor.Text, 15)), lbEliminaCobertura, fvListaCompraDeuda, fnMontoExpEsteCred_NEW, fbEliminarEvaluacion, fvListaAguaSaneamiento, nSubDestino, nTpDoc, nTpIngr, nTpInt, fvListaCreditoVerde) 'CInt(Trim(Right(cmbMotivoRef.Text, 5))))
            'WIOR 20140509 AGREGO IIf(fraPromotor.Visible, fbRegPromotores, False),Trim(Right(cmbPromotor.Text, 15))
            'EJVG20160203 ERS002-2016 agregó fvListaCompraDeuda
            'LUCV20170425 Modificó: IIf(fraPromotor.Visible, fbRegPromotores, False)
            'JOEP20190919 ERS042 CP-2018 Agrego nSubDestino, nTpDoc, nTpIngr
            'EAAS20191401 SEGUN 018-GM-DI_CMACM fvListaCreditoVerde
'           ***************************************************************************************************
            'MIOL 20130626, RQ13335 Agrego Parametro Me.txtVendedor.Text
            ''*** PEAC 20090126
            'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, , psNuevaCta, gCodigoCuenta 'Comentado JOEP22052017
             Set objPista = New COMManejador.Pista 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
            'objPista.InsertarPista 190260, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, , psNuevaCta, gCodigoCuenta 'JOEP22052017 'LUCV20181220 Comentó, Anexo01 de Acta 199-2018
            objPista.InsertarPista gCredRegistrarActualizaSoliCred, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Solicitud del Crédito (Sicmac Negocio)", psNuevaCta, gCodigoCuenta 'JOEP22052017 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
            Set objPista = Nothing 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                    
            Set oCredito = Nothing
        End If
        cmdImprimir.Enabled = True
    'End If
    
    
    'JOEP20190919 ERS042 CP-2018
    Dim objAporte_CP As COMDCredito.DCOMCredito
    If IsArray(nMatMontoPre) Then
        Set objAporte_CP = New COMDCredito.DCOMCredito
        If UBound(nMatMontoPre) > 0 Then
            If Trim(Right(cmbSubProducto.Text, 5)) = "703" Then
                Call objAporte_CP.CP_GrabaAporte(psNuevaCta, CCur(nMatMontoPre(1, 1)), CCur(nMatMontoPre(1, 2)), CCur(nMatMontoPre(1, 3)), CCur(nMatMontoPre(1, 4)))
            Else
                Call objAporte_CP.CP_GrabaAporte(psNuevaCta, CCur(nMatMontoPre(1, 1)), CCur(nMatMontoPre(1, 2)), CCur(nMatMontoPre(1, 3)), -1)
            End If
        Else
            Set objAporte_CP = New COMDCredito.DCOMCredito
            Call objAporte_CP.CP_DeleteAporte(psNuevaCta)
            Set objAporte_CP = Nothing
        End If
        Set objAporte_CP = Nothing
    Else
        Set objAporte_CP = New COMDCredito.DCOMCredito
        Call objAporte_CP.CP_DeleteAporte(psNuevaCta)
        Set objAporte_CP = Nothing
    End If
    'JOEP20190919 ERS042 CP-2018
    
    '*** PEAC 20080807
    '*** aqui no tenemos el nMovNro para poder relacional con el credito
    '*** que hacemos ?
    loVistoElectronico.RegistraVistoElectronico (0)
    
    '*** FIN PEAC
    
    'JUEZ 20130527 *******************************************************************************************
    Dim oEnvEstCta As COMDCaptaGenerales.DCOMCaptaGenerales
    If fbRegistraEnvio Then
        If fnModoEnvioEstCtaSiNo Then
            Call frmEnvioEstadoCta.GuardarRegistroEnvioEstadoCta(1, psNuevaCta, frsEnvEstCta, fnModoEnvioEstCta, 0, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
        Else
            'APRI20180310 ERS036-2107
             Set oEnvEstCta = New COMDCaptaGenerales.DCOMCaptaGenerales
             Call oEnvEstCta.EliminarRegistroEnvioEstadoCta(psNuevaCta)
             Set oEnvEstCta = Nothing
            'END APRI
        End If
    Else
        Dim rsEnvio As ADODB.Recordset
        Set oEnvEstCta = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsEnvio = oEnvEstCta.RecuperaDatosEnvioEstadoCta(psNuevaCta)
        Set oEnvEstCta = Nothing
        If IsNull(rsEnvio!nModoEnvio) Then
            Call frmEnvioEstadoCta.GuardarRegistroEnvioEstadoCta(1, psNuevaCta, LlenaRecordSet_Cliente, 1, 0, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
        End If
    End If
    'END JUEZ ************************************************************************************************
    'WIOR 20130723 ************************************************
    If oCredAgrico.Registrar Then
        Call oCredAgrico.InsertaActDatos(psNuevaCta, oCredAgrico.TipoAct, oCredAgrico.SubTipoAct, oCredAgrico.TotalHect, oCredAgrico.HectProd, oCredAgrico.Animales, oCredAgrico.CodCooperativa)
    Else
        Call oCredAgrico.EliminaAgricolasCred(psNuevaCta)
        Set oCredAgrico = New frmCredAgricoSelec
    End If
    'WIOR FIN ******************************************************
    
    'JUEZ 20140603 ***************************************************
    Set oDPersGen = New COMDPersona.DCOMPersGeneral
    If bSolicitaAutSectorEcon Then
        Call oDPersGen.InsertarSolicitudAutorizacionRiesgos(psNuevaCta, CDbl(txtMontoSol.Text))
    Else
        Call oDPersGen.EliminarSolicitudAutorizacionRiesgos(psNuevaCta)
    End If
    Set oDPersGen = Nothing
    'END JUEZ ********************************************************
    
    'JOEP ERS047 20170901 ***************************************************
    Set oCredito = New COMDCredito.DCOMCredito
    If bSolicitaAutZonaGeog Then
        Call oCredito.InsertarSolicitudAutorizacionZonaGeog(psNuevaCta, CDbl(txtMontoSol.Text))
    Else
        Call oCredito.EliminarSolicitudAutorizacionZonaxProduxGarant(psNuevaCta, Trim(Right(cmbSubProducto.Text, 5)), 1)
    End If
    Set oCredito = Nothing
    
    Set oCredito = New COMDCredito.DCOMCredito
    If bSolicitaAutTpCredito Then
        Call oCredito.InsertarSolicitudAutorizacionTpCredito(psNuevaCta, Trim(Right(cmbSubProducto.Text, 5)), CDbl(txtMontoSol.Text))
    Else
        Call oCredito.EliminarSolicitudAutorizacionZonaxProduxGarant(psNuevaCta, Trim(Right(cmbSubProducto.Text, 5)), 2)
    End If
    Set oCredito = Nothing
    'END JOEP ERS47 20170901 ********************************************************
    
    'EJVG20151015 ***
    If lbEliminaCobertura Then
        MsgBox "Ud. tendrá que realizar nuevamente el [Registro de Cobertura] de las Garantías con el Próducto", vbInformation, "Aviso"
        EnfocaControl cmdGravar
    End If
    'END EJVG *******
    'JUEZ 20160509 ************************************************************
    If bAmpliacion And chkAutAmpliacion.value = 1 Then
        Set oCredito = New COMDCredito.DCOMCredito
            oCredito.RegistraSolicitudAutorizacionAmpliacion psNuevaCta, CDbl(txtMontoSol.Text), nMontoTotal - pnMontoITf, lsComentario, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set oCredito = Nothing
    End If
    'END JUEZ *****************************************************************
    'FRHU 20160712 ERS002-2016 CAMBIO
    Dim lbSoloTitular As Boolean
    Dim lsPersCodTitular As String
    lbSoloTitular = True
    oRelPersCred.IniciarMatriz
    Do While Not oRelPersCred.EOF
        If lsPersCodTitular = "" Then
            lsPersCodTitular = oRelPersCred.ObtenerMatrizRelacionesRelacion(20, 1)
        End If
        If oRelPersCred.ObtenerValorRelac = gColRelPersConyugue Or oRelPersCred.ObtenerValorRelac = gColRelPersCodeudor Then
            lbSoloTitular = False
            Call verificarAlertas(lsPersCodTitular, oRelPersCred.ObtenerCodigo)
        End If
        oRelPersCred.siguiente
    Loop
    If lbSoloTitular Then Call verificarAlertas(lsPersCodTitular, "")

    'MARG ERS003-2018--------------------
    'Call ValidarScoreExperian(psNuevaCta, MatCredRelaciones)'comment by marg 201906
    'END MARG----------------------------
    Call ValidarScore(psNuevaCta, MatCredRelaciones, 1) 'add by marg 201906

    'FIN FRHU 20160712
    If Not (oDCredi.VerificaCampania(nIdCampana)) Then 'BY ARLO20171127
    'If (nIdCampana <> 116) Then 'ARLO20170818 'BY ARLO20171127 COMENT
        'FRHU 20160802 ERS002-2016
        If ValidarExisteNivelAprobacionParaAutorizacion(psNuevaCta) Then
            'FRHU 20160615 ERS002-2016
            Set oCredito = New COMDCredito.DCOMCredito
            Call oCredito.RegistraAutorizacionesRequeridas(Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss"), gsCodUser, gsCodAge, psNuevaCta)
            If oCredito.verificarExisteAutorizaciones(psNuevaCta) Then
                Call frmCredNewNivAutorizaVer.Consultar(psNuevaCta)
            End If
            Set oCredito = Nothing
            'FIN FRHU
        End If
        'FRHU 20160802
    End If  'ARLO20170818

    'EJVG20160712 ***
    If fbEliminarEvaluacion Then
        MsgBox "Ud. deberá registrar nuevamente la Evaluación del Crédito", vbInformation, "Aviso"
        Set objPista = New COMManejador.Pista 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
        objPista.InsertarPista gCredEliminacionEvaluacionCred, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gEliminar, "Evaluacion Credito: Se eliminó Formato Nro. " & nFormatoEliminado & " - Formato Nuevo Nro. " & nFormato_NEW & " ", psNuevaCta, gCodigoCuenta 'JOEP22052017 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
        Set objPista = Nothing 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
        EnfocaControl cmdEvaluar
    End If
    'END EJVG *******
    Call HabilitaIngresoSolicitud(False)
    ActXCtaCred.NroCuenta = psNuevaCta
    ActXCtaCred.Enabled = False
    cmbProductoCMACM.Enabled = False
    cmbSubProducto.Enabled = False
    cmbMoneda.Enabled = False
    cmdRelaciones.Enabled = True
    cmdEditar.Enabled = True
    cmdEliminar.Enabled = True
    'cmdEnvioEstCta.Enabled = True 'JUEZ 20130527
    cmdEjecutar = -1
    nCampanaCod = 0
    '***Modificado por ELRO 20111017, según Acta 222-2011/TI-D
    'PersoneriaTitular = 0'WIOR 20130719 COMENTO ESTA VARIABLE
    '*********************************************************
    If bLeasing Then 'EJVJ20120720
        cmdNuevo.Enabled = False
    End If
    
    'add pti1 ers0702018 18/12/2018
    If nPermiso = Registrar And CboAutoriazaUsoDatos.Visible Then
        Dim oNPersona As New COMNPersona.NCOMPersona
        Dim oCont As New COMNContabilidad.NCOMContFunciones
        Dim rsPersona As New ADODB.Recordset
        Dim sMovNro As String
        
        Dim MatPersona(1 To 2) As TActAutDatos
        Dim sUbicGeografica As String
        Dim sNombres As String, sApePat As String, sApeMat As String, sApeCas As String
        Dim sSexo As String, sEstadoCivil As String
        Dim sDomicilio As String, cNacionalidad As String, sRefDomicilio As String
        Dim sTelefonos As String, sCelular As String, sEmail As String
        Dim sPersIDTpo As String, sPersIDnro As String
        Dim lbClienteEsPersonaNatural As Boolean
        Dim sPerCodcli As String
        Dim nAutorizaUsoDatos As Integer
        
        sPerCodcli = Trim(lblCodigo.Caption)
        lbClienteEsPersonaNatural = oNPersona.ClienteEsPersonaNatural(sPerCodcli)
        nAutorizaUsoDatos = CboAutoriazaUsoDatos.ListIndex
        If lbClienteEsPersonaNatural Then
            Set rsPersona = oNPersona.ObtenerDatosParaActAutDeCliente(sPerCodcli)
            sApePat = BuscaNombre(rsPersona!cPersNombre, BusqApellidoPaterno)
            sApeMat = BuscaNombre(rsPersona!cPersNombre, BusqApellidoMaterno)
            sNombres = BuscaNombre(rsPersona!cPersNombre, BusqNombres)
            sApeCas = BuscaNombre(rsPersona!cPersNombre, BusqApellidoCasada)
            sSexo = Trim(IIf(IsNull(rsPersona!cPersnatSexo), "", rsPersona!cPersnatSexo))
            sEstadoCivil = Trim(IIf(IsNull(rsPersona!nPersNatEstCiv), "", rsPersona!nPersNatEstCiv))
            sDomicilio = rsPersona!cPersDireccDomicilio
            sUbicGeografica = rsPersona!cPersDireccUbiGeo
            sTelefonos = IIf(IsNull(rsPersona!cPersTelefono), "", rsPersona!cPersTelefono)
            sCelular = IIf(IsNull(rsPersona!cPersCelular), "", rsPersona!cPersCelular)
            sEmail = IIf(IsNull(rsPersona!cEmail), "", rsPersona!cEmail)
            cNacionalidad = Trim(IIf(IsNull(rsPersona!cNacionalidad), "", rsPersona!cNacionalidad))
            sRefDomicilio = Trim(IIf(IsNull(rsPersona!cPersRefDomicilio), "", rsPersona!cPersRefDomicilio))
            sPersIDTpo = Trim(IIf(IsNull(rsPersona!cPersIDTpo), "", rsPersona!cPersIDTpo))
            sPersIDnro = Trim(IIf(IsNull(rsPersona!cPersIDnro), "", rsPersona!cPersIDnro))
            
            MatPersona(1).sNombres = sNombres
            MatPersona(1).sApePat = sApePat
            MatPersona(1).sApeMat = sApeMat
            MatPersona(1).sApeCas = sApeCas
            MatPersona(1).sPersIDTpo = sPersIDTpo
            MatPersona(1).sPersIDnro = sPersIDnro
            MatPersona(1).sSexo = sSexo
            MatPersona(1).sEstadoCivil = sEstadoCivil
            MatPersona(1).cNacionalidad = cNacionalidad
            MatPersona(1).sDomicilio = sDomicilio
            MatPersona(1).sRefDomicilio = sRefDomicilio
            MatPersona(1).sUbicGeografica = sUbicGeografica
            MatPersona(1).sCelular = sCelular
            MatPersona(1).sTelefonos = sTelefonos
            MatPersona(1).sEmail = sEmail
            
            MatPersona(2).sNombres = sNombres
            MatPersona(2).sApePat = sApePat
            MatPersona(2).sApeMat = sApeMat
            MatPersona(2).sApeCas = sApeCas
            MatPersona(2).sPersIDTpo = sPersIDTpo
            MatPersona(2).sPersIDnro = sPersIDnro
            MatPersona(2).sSexo = sSexo
            MatPersona(2).sEstadoCivil = sEstadoCivil
            MatPersona(2).cNacionalidad = cNacionalidad
            MatPersona(2).sDomicilio = sDomicilio
            MatPersona(2).sRefDomicilio = sRefDomicilio
            MatPersona(2).sUbicGeografica = sUbicGeografica
            MatPersona(2).sCelular = sCelular
            MatPersona(2).sTelefonos = sTelefonos
            MatPersona(2).sEmail = sEmail
            
            If CboAutoriazaUsoDatos.ListIndex = 0 Then
            nAutorizaUsoDatos = 1
            End If
            
            If CboAutoriazaUsoDatos.ListIndex = 1 Then
            nAutorizaUsoDatos = 0
            End If
            
            
            If CboAutoriazaUsoDatos.Visible And (nInicioActDa = -1 Or nInicioActDa = 0) Then
            sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            Call oNPersona.InsertarPersActAutDatos(sMovNro, 0, sPerCodcli, MatPersona(), nAutorizaUsoDatos, 2, 1, 0, 1, 3)
            MsgBox "Recordar que debe imprimir y solicitar la firma de la cartilla de Autorización de Uso de Datos", vbInformation, "Aviso"
            End If
            
        End If
      
    
    End If
    
    
    'RECO20160718***************************************
    Call CargaDatos(psNuevaCta)
    cmdEditar.Enabled = True
    cmdEliminar.Enabled = True
    cmdGarantias.Enabled = True
    cmdGravar.Enabled = True
    cmdEvaluar.Enabled = True
    Call ControlesPermiso
    If bLeasing Then 'EJVJ20120720
        cmdNuevo.Enabled = False
    End If
    'RECO FIN********************************************
    'Exit Function
    Exit Sub

'ErrorCmdGrabar_Click:
'        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Function CargarMatrizRelaCredito(oRelPersCred As UCredRelac_Cli) As Variant
Dim i As Integer
Dim MatTemp As Variant

oRelPersCred.IniciarMatriz
Do While Not oRelPersCred.EOF
    ReDim Preserve MatTemp(i)
    oRelPersCred.ObtenerCodigo
    oRelPersCred.ObtenerValorRelac
    oRelPersCred.siguiente
    i = i + 1
Loop

CargarMatrizRelaCredito = MatTemp
End Function

Private Sub cmdGravar_Click()
    'EJVG20150707 ***
    frmGarantiaCobertura.Inicio InicioGravamenxSolicitud, Credito, ActXCtaCred.NroCuenta, bLeasing
'If bLeasing = True Then
'    frmCredGarantCred.Inicioleasing PorSolicitud, ActXCtaCred.NroCuenta
'Else
'    frmCredGarantCred.Inicio PorSolicitud, ActXCtaCred.NroCuenta
'End If
    'END EJVG *******
End Sub

Private Sub cmdImprimir_Click()
Dim oNCredDoc As COMNCredito.NCOMCredDoc
Dim oPrev As previo.clsprevio
    Set oNCredDoc = New COMNCredito.NCOMCredDoc
    Set oPrev = New previo.clsprevio
    oPrev.Show oNCredDoc.ImprimeSolicitud(ActXCtaCred.NroCuenta, gsNomAge, gdFecSis, gsCodUser, gsNomCmac), "Registro de Solicitud"

    'Add pti1 ers070-2018 ******************************
            Dim spercod As String
            Dim snumcuenta As String
            Dim oNPersona As New COMNPersona.NCOMPersona
            Dim rsImpresion As ADODB.Recordset
            Dim rsPersona As ADODB.Recordset
            Dim CondicionCred  As String
            Dim cPersCod  As String
            Dim CantImpr As Integer
            Dim CanR As Integer
            Dim Tsi As Integer
            Dim Tno As Integer
            
            
            Dim sUbicGeografica As String
            Dim sNombres As String, sApePat As String, sApeMat As String, sApeCas As String
            Dim sSexo As String, sEstadoCivil As String
            Dim sDomicilio As String, cNacionalidad As String, sRefDomicilio As String
            Dim sTelefonos As String, sCelular As String, sEmail As String
            Dim sPersIDTpo As String, sPersIDnro As String
            
            snumcuenta = Trim(ActXCtaCred.NroCuenta)
            Set rsImpresion = oNPersona.DupDocCred(snumcuenta)
            If Not (rsImpresion.EOF And rsImpresion.BOF) Then
                 CondicionCred = rsImpresion!CondicionCred
                 cPersCod = rsImpresion!cPersCod
                 CantImpr = rsImpresion!CantImpr
                 CanR = rsImpresion!CanR
                 Tsi = rsImpresion!Tsi
                 Tno = rsImpresion!Tno
                
                 
                 Set rsPersona = oNPersona.ObtenerDatosParaActAutDeCliente(cPersCod)
                If Not (rsPersona.EOF And rsPersona.BOF) And CanR > 0 Then
                    ultEstado = rsImpresion!ultEstado
                    sAgeReg = rsImpresion!Agencia
                    dfreg = rsImpresion!freg
                    sApePat = BuscaNombre(rsPersona!cPersNombre, BusqApellidoPaterno)
                    sApeMat = BuscaNombre(rsPersona!cPersNombre, BusqApellidoMaterno)
                    sNombres = BuscaNombre(rsPersona!cPersNombre, BusqNombres)
                    sApeCas = BuscaNombre(rsPersona!cPersNombre, BusqApellidoCasada)
                    sSexo = Trim(IIf(IsNull(rsPersona!cPersnatSexo), "", rsPersona!cPersnatSexo))
                    sEstadoCivil = Trim(IIf(IsNull(rsPersona!nPersNatEstCiv), "", rsPersona!nPersNatEstCiv))
                    sDomicilio = rsPersona!cPersDireccDomicilio
                    sUbicGeografica = rsPersona!cPersDireccUbiGeo
                    sTelefonos = IIf(IsNull(rsPersona!cPersTelefono), "", rsPersona!cPersTelefono)
                    sCelular = IIf(IsNull(rsPersona!cPersCelular), "", rsPersona!cPersCelular)
                    sEmail = IIf(IsNull(rsPersona!cEmail), "", rsPersona!cEmail)
                    cNacionalidad = Trim(IIf(IsNull(rsPersona!cNacionalidad), "", rsPersona!cNacionalidad))
                    sRefDomicilio = Trim(IIf(IsNull(rsPersona!cPersRefDomicilio), "", rsPersona!cPersRefDomicilio))
                    sPersIDTpo = Trim(IIf(IsNull(rsPersona!cPersIDTpo), "", rsPersona!cPersIDTpo))
                    sPersIDnro = Trim(IIf(IsNull(rsPersona!cPersIDnro), "", rsPersona!cPersIDnro))
                    
                       MatPersona(1).sNombres = sNombres
                       MatPersona(1).sApePat = sApePat
                       MatPersona(1).sApeMat = sApeMat
                       MatPersona(1).sApeCas = sApeCas
                       MatPersona(1).sPersIDTpo = sPersIDTpo
                       MatPersona(1).sPersIDnro = sPersIDnro
                       MatPersona(1).sSexo = sSexo
                       MatPersona(1).sEstadoCivil = sEstadoCivil
                       MatPersona(1).cNacionalidad = cNacionalidad
                       MatPersona(1).sDomicilio = sDomicilio
                       MatPersona(1).sRefDomicilio = sRefDomicilio
                       MatPersona(1).sUbicGeografica = sUbicGeografica
                       MatPersona(1).sCelular = sCelular
                       MatPersona(1).sTelefonos = sTelefonos
                       MatPersona(1).sEmail = sEmail
                     If CantImpr = 0 Then
                             If CondicionCred = 1 Then
                                'Cliente Nuevo
                                Call ImprimirPdfCartillaAutorizacion
                                Call oNPersona.imprimeCartAut(gsCodUser, snumcuenta)
                             Else
                                If CanR = 1 Then
                                    'si es un cliente recurrente y por primera vez autoriza sus datos
                                    Call ImprimirPdfCartillaAutorizacion
                                    Call oNPersona.imprimeCartAut(gsCodUser, snumcuenta)
                                    
                                Else
                                  'si es un cliente recurrente y CAMBIA DE NO A SI
                                  If CanR > 1 And Tsi = 1 And ultEstado = 1 Then
                                    Call ImprimirPdfCartillaAutorizacion
                                    Call oNPersona.imprimeCartAut(gsCodUser, snumcuenta)
                                  End If
                                
                                End If
                             End If
                     Else
                        
                         If vbYes = MsgBox("¿Desea Re-Imprimir la cartilla de Autorización de Uso de Datos?", vbInformation + vbYesNo) Then
                                    Call ImprimirPdfCartillaAutorizacion
                                    Call oNPersona.imprimeCartAut(gsCodUser, snumcuenta)
                         End If
                        
                    End If
                End If
                
           End If
            
            
    Set oPrev = Nothing
    Set oNCredDoc = Nothing
    
End Sub

Private Sub cmdLimpiar_Click()
    Call cmdCancela_Click
    Call LimpiaPantalla
    Call ControlesPermiso
    cmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    cmdGarantias.Enabled = False
    cmdGravar.Enabled = False
    cmdEvaluar.Enabled = False
    'JOEP20190919 ERS042 CP-2018
    cmbSubDestCred.Visible = False
    nSubDestAnt = 0
    ReDim nMatMontoPre(0)
    bEntrotxtMontoSol = False
    'JOEP20190919 ERS042 CP-2018
    If bLeasing Then 'EJVG20120720
        cmdNuevo.Enabled = False
    End If
    ''WIOR 20120914 ******************************************
    'If nExisteAgeEvalCred = 0 Then
    '    cmdEvaluar.Visible = False
    'Else
    '    cmdEvaluar.Visible = True
    '    cmdEvaluar.Enabled = False
    'End If
    ''WIOR FIN ***********************************************
End Sub

Private Sub cmdNuevo_Click()
'JOEP20190919 ERS042 CP-2018
    nSubDestAnt = 0
    ReDim nMatMontoPre(0)
'JOEP20190919 ERS042 CP-2018
    Call cmdLimpiar_Click
    Set oRelPersCred = Nothing
    Set oPersona = Nothing
    Set oRelPersCred = New UCredRelac_Cli
    cmdEjecutar = 1
    'WIOR 20140509 ****************************
    If fbRegPromotores Then
        fraPromotor.Visible = True
    End If
    'WIOR FIN *********************************
    Call HabilitaIngresoSolicitud(True)
    
    If bRefinanciar Then
        ChkCap.Enabled = True
    End If
    
    'CMACICA_CSTS - 10112003 ---------------------------------------------------------------------
    If bSustituirDeudor Then
       ChkCap.Enabled = False
    End If
    '---------------------------------------------------------------------------------------------
    bfCredSolicitud = True 'PTI1 ADD 24082018 ERS027-2017
    cmdRelaciones_Click
    bfCredSolicitud = False 'PTI1 ADD 24082018 ERS027-2017
    
    'AGREGADO POR PTI1 22/08/2018 ERS027-2017
      If frmCredRelaCta.getnEstadoVerifica = 1 Then
         Exit Sub
      End If
    'FIN AGREGADO POR PTI1
    
    'Carga Fuentes de Ingreso
    If oRelPersCred.NroRelaciones = 0 Then
        Call cmdCancela_Click
        Call LimpiaPantalla
        cmdEditar.Enabled = False
        cmdEliminar.Enabled = False
        cmdGarantias.Enabled = False
        cmdGravar.Enabled = False
        cmdEvaluar.Enabled = False
        Exit Sub
    End If
    Call CargaFuentesIngreso(TitularCredito) 'ARCV 29-12-2006
    
    Call CargaDatosTitular
    cmdRelaciones.Enabled = True
    CmdAmpliacion.Enabled = True
    cmbCondicionOtra.Enabled = True
    cmbCondicionOtra.ListIndex = 0
    cmdEnvioEstCta.Enabled = True 'JUEZ 20130527
    
    'MADM 20100719
    cmbCondicionOtra2.Enabled = True
    cmbCondicionOtra2.ListIndex = 0
    'END MADM
    
    'bAmpliacion = False '12-03-2007
    fbActivo = True 'WIOR 20130723
    'add pti1 ers070-2018 18/12/2018
    Call ActAutDatos
    Call CP_HabilitaControles(False) 'Agrego JOEP20190919 ERS042 CP-2018
    
        '**ARLO20181126 ERS068 - 2018
    If bRefinanciar Then
            Dim rs As ADODB.Recordset
            Dim oDCred As COMDCredito.DCOMCredito
            Set oDCred = New COMDCredito.DCOMCredito
            Set rs = oDCred.ValidaPropuestaRefinanciado(sCodTitular)
            If Not (rs.EOF And rs.BOF) Then
                If rs!sMensaje <> "" Then
                    MsgBox rs!sMensaje, vbInformation, "Aviso"
                    cmdCancela_Click
                    sCodTitular = ""
                    Exit Sub
                End If
            End If
            Set oDCred = Nothing
            Set rs = Nothing
    End If
    '**ARLO END
    
End Sub

'Private Sub cmdpresolicitud_Click() 'COMENTADO POR PTI1 22/08/2018 ERS027-2017
Public Sub cmdpresolicitud_Click() 'ADD POR PTI1 22/08/2018
Dim sPersona As String 'ADD PTI1 22/08/2018
sPersona = frmCredRelaCta.getsPersona 'ADD PTI1 22/08/2018
nPresolicitudId = frmPreSolicitud.Inicio(gsCodUser, bAmpliacion, sPersona) 'ADD PTI1 22/08/2018
'nPresolicitudId = frmPreSolicitud.Inicio(gsCodUser, bAmpliacion) 'COMENTADO POR PTI1 ERS027-2017
 If nPresolicitudId <> -1 Then
    Dim oHojaRuta As COMDCredito.DCOMhojaRuta
    Set oHojaRuta = New COMDCredito.DCOMhojaRuta
    Set rsPresol = oHojaRuta.ObtenerPreSolicitudesXid(nPresolicitudId)
    Set oHojaRuta = Nothing
    'add pti1 ERS027-2017
    If Not (rsPresol.BOF And rsPresol.EOF) Then
        cPersCodPreSol = rsPresol!cPersCod
        bPresol = True
        bPreSolOperacion = True
        bPresolAmpliAuto = True
        MsgBox "Ha elegido una Pre-solicitud, ahora debe elegir el boton Nuevo para continuar", vbInformation, "Aviso"
        cmdpresolicitud.Enabled = False
    Else
       MsgBox "No existen Pre-solicitudes pendientes", vbInformation, "Aviso"
    End If
    'fin pti1 ERS027-2017
     'comentado por pti1 ers027-2017
'    cPersCodPreSol = rsPresol!cPersCod
'    bPresol = True
'    bPreSolOperacion = True
'    bPresolAmpliAuto = True
'    MsgBox "Ha elegido una Pre-solicitud, ahora debe elegir el boton Nuevo para continuar", vbInformation, "Aviso"
'    cmdpresolicitud.Enabled = False
 Else
    MsgBox "No existen Pre-solicitudes pendientes", vbInformation, "Aviso"
 End If
End Sub

Private Sub CmdRefinanc_Click()
Dim i As Integer
Dim nMontoRef As Double
Dim Coma As String
Dim fnDestino As Integer 'JAME20140509
'JOEP20190115 CP
Dim obj As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Dim cCtaCodRef As String
'JOEP20190115 CP

Coma = "'"

'Agrego JOEP20171222 Acta 226-2017
If bRefinanciar Then
    If ChkCap.Enabled = True And ChkCap.value = 0 Then
        'If MsgBox("Seguro que no desea capitalizar el interés. ¿Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        If MsgBox("No desea capitalizar el interés, mora y gastos", vbQuestion + vbYesNo, "Aviso") = vbYes Then Exit Sub
    End If
End If
'Agrego JOEP20171222 Acta 226-2017

    If cmbMoneda.ListIndex = -1 Then
        MsgBox "Seleccione la Moneda del Credito", vbInformation, "Aviso"
        'cmbMoneda.SetFocus 'comento JOEP20190115 CP
        'JOEP20190115 CP
        If cmbMoneda.Enabled = True Then
            cmbMoneda.SetFocus
        End If
        'JOEP20190115 CP
        Exit Sub
    Else
        cmbMoneda.Enabled = False
    End If
    MatCredRef = frmCredRefinanc.Inicio(CInt(Trim(Right(cmbMoneda.Text, 20))), MatCredRef, IIf(ChkCap.value = 1, True, False), False, fnDestino)
    'By Capi 18092008 para el control de interes a capitalizar
    'ChkCap.Enabled = False 'Comento JOEP20190204 CP
    
    '**ARLO20180319 ERS070 - 2017 ANEXO 02
    If bRefinanciar = True And fnDestino = 14 Then
        Me.cmdDestinoDetalle.Enabled = False
    End If
    '**ARLO20180319 ERS070 - 2017 ANEXO 02
    
    If IsArray(MatCredRef) Then
        If UBound(MatCredRef) > 0 Then
            cCtaCodRef = "" 'JOEP20190115 CP
            fsCtaCod = "" 'JOEP20190115 CP
            nMontoRef = 0
            For i = 0 To UBound(MatCredRef) - 1
                nMontoRef = nMontoRef + CDbl(MatCredRef(i, 9))
                nMontoRef = CDbl(Format(nMontoRef, "#0.00"))
                If i = 0 Then
                    fsCtaCod = fsCtaCod & "(" & Coma & MatCredRef(i, 0) & Coma & ","
                    cCtaCodRef = MatCredRef(i, 0) 'JOEP20190115 CP
                Else
                    fsCtaCod = fsCtaCod & Coma & MatCredRef(i, 0) & Coma & ","
                    cCtaCodRef = cCtaCodRef & "," & MatCredRef(i, 0) 'JOEP20190115 CP
                End If
            Next i
            txtMontoSol.Text = Format(nMontoRef, "#0.00")
            'JAME20140509 ***
            'Comento JOEP CP ERS042-2018
            'cmbDestCred.ListIndex = IndiceListaCombo(cmbDestCred, fnDestino)
            'cmbDestCred.Enabled = False
            'Comento JOEP CP ERS042-2018
            'END JAME *******
            ChkCap.Enabled = False 'JOEP20190204 CP
        Else
            ChkCap.Enabled = True 'JOEP20190204 CP
            txtMontoSol.Text = "0.00"
        End If
    Else
        ChkCap.Enabled = True 'JOEP20190204 CP
        txtMontoSol.Text = "0.00"
    End If
    If fsCtaCod <> "" Then
'JOEP CP ERS042-2018
        Set obj = New COMDCredito.DCOMCredito
        Set rs = obj.CP_getDestinoRef(cCtaCodRef)
        If Not (rs.BOF And rs.EOF) Then
            fnDestino = rs!nConsValor
            Call Llenar_Combo_con_Recordset(rs, cmbDestCred)
            cmbDestCred.ListIndex = IndiceListaCombo(cmbDestCred, fnDestino)
        End If
        
        If cmbDestCred.Text <> "" Then
            cmbDestCred.Enabled = False
            txtMontoSol.Enabled = False
        End If
'JOEP CP ERS042-2018
        fsCtaCod = Left(fsCtaCod, Len(fsCtaCod) - 1) & ")"
    End If
'JOEP CP ERS042-2018
Set obj = Nothing
RSClose rs
'JOEP CP ERS042-2018
End Sub


Private Sub cmdSeleccionarFuentes_Click()
    Call frmCredSolicitud_SelecFtes.Inicio(oPersona.PersCodigo)
    MatFuentes = frmCredSolicitud_SelecFtes.MatFuentes

End Sub

Private Sub CmdSustitucionDeudor_Click()
Dim i As Integer
Dim nMontoSust As Double
        
    If cmbMoneda.ListIndex = -1 Then
        MsgBox "Seleccione la Moneda del Credito", vbInformation, "Aviso"
        cmbMoneda.SetFocus
        Exit Sub
    Else
        cmbMoneda.Enabled = False
    End If
    
    MatCredSust = frmCredRefinanc.Inicio(CInt(Trim(Right(cmbMoneda.Text, 20))), MatCredSust, IIf(ChkCap.value = 1, True, False), True)
    If IsArray(MatCredSust) Then
        If UBound(MatCredSust) > 0 Then
            nMontoSust = 0
            For i = 0 To UBound(MatCredSust) - 1
                nMontoSust = nMontoSust + CDbl(MatCredSust(i, 9))
                nMontoSust = CDbl(Format(nMontoSust, "#0.00"))
            Next i
            txtMontoSol.Text = Format(nMontoSust, "#0.00")
        Else
            txtMontoSol.Text = "0.00"
        End If
    Else
        txtMontoSol.Text = "0.00"
    End If

End Sub

Private Sub cmdRelaciones_Click()
Dim oDCred As COMDCredito.DCOMCredito
Dim sMensaje As String
'Dim lsPersCodTitular As String 'FRHU 20160702 ERS002-2016 'SE QUITO 20160712

'agregado por vapi SEGÙN ERS TI-ERS001-2017
    If bPresol Then
        Call frmCredRelaCta.Inicio(oRelPersCred, InicioSolicitud, , cmdEjecutar, True, cPersCodPreSol)
    Else
        cPersCodPreSol = "" 'ADD PTI1 ERS027-2017 25082018
        frmCredRelaCta.ActualizaEstadoVerifica (0) 'ADD PTI1 ERS027-2017 25082018
        Call frmCredRelaCta.Inicio(oRelPersCred, InicioSolicitud, , cmdEjecutar)
      
       'AGREGADO POR PTI1 22/08/2018
        If frmCredRelaCta.getnEstadoVerifica = 1 Then
         Call cmdCancela_Click
         Exit Sub
        End If
        'FIN AGREGADO POR PTI1
    End If
    'fin agregado por vapi

    'Call frmCredRelaCta.Inicio(oRelPersCred, InicioSolicitud, , cmdEjecutar) 'COMENTADO POR VAPI SEGUN ERS 001-2017
    
'    Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
'        'sCadImp = sCadImp & oCOMNCredDoc.ImprimeRepClientesHisNegativo(MatAgencias, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
'        Call oCOMNCredDoc.ImprimeRepClientesHisNegativo(MatAgencias, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
'    Set oCOMNCredDoc = Nothing
    
    oRelPersCred.IniciarMatriz
    'AQUI VERIFICAR SI TAMBIEN SE COLOCARA LA VALIDACION DE MAS DE 65 AÑOS
    Do While Not oRelPersCred.EOF
        If oRelPersCred.ObtenerValorEdad(gdFecSis) < 18 And oRelPersCred.ObtenerValorPersoneria = 1 Then
             MsgBox "La Persona : " & oRelPersCred.ObtenerNombre & " es Menor de Edad, No puede estar realacionada con un credito", vbInformation, "Aviso"
            Call oRelPersCred.EliminarRelacion(oRelPersCred.ObtenerCodigo, oRelPersCred.ObtenerValorRelac)
        End If
        ' CMACICA_CSTS - 14/11/2003 ---------------------------------------------------------------------
        If oRelPersCred.ObtenerValorEdad(gdFecSis) > 65 And oRelPersCred.ObtenerValorPersoneria = 1 Then
           MsgBox "La Persona : " & oRelPersCred.ObtenerNombre & " supera los 65 años de Edad.", vbInformation, "Aviso"
        End If
        '------------------------------------------------------------------------------------------------
        If oRelPersCred.ObtenerValorRelac = gColRelPersTitular Then
            'lsPersCodTitular = oRelPersCred.ObtenerCodigo 'FRHU 20160702 ERS002-2016 'SE QUITO 20160712
            Set oDCred = New COMDCredito.DCOMCredito
            'CUSCO
            'If oDCred.NumerosCredEnJudicial(oRelPersCred.ObtenerCodigo) > 0 Then
            '    MsgBox "Titular tiene creditos en judicial", vbInformation, "Aviso"
            'End If
            sCodTitular = oRelPersCred.ObtenerCodigo 'ARLO20181126 ERS068-2018
            sMensaje = oDCred.ValidaTitularCredito(oRelPersCred.ObtenerCodigo)
            If sMensaje <> "" Then
                MsgBox sMensaje, vbInformation, "Aviso"
            End If
            Set oDCred = Nothing
        End If
        'FRHU 20160702 ERS002-2016 ' SE QUITO 20160712
        'If oRelPersCred.ObtenerValorRelac = gColRelPersConyugue Then
            'Call verificarAlertas(lsPersCodTitular, oRelPersCred.ObtenerCodigo)
        'End If
        'If oRelPersCred.ObtenerValorRelac = gColRelPersCodeudor Then
            'Call verificarAlertas(lsPersCodTitular, oRelPersCred.ObtenerCodigo)
        'End If
        'FIN FRHU 20160702
        oRelPersCred.siguiente
    Loop
    
    Call ActualizarListaPersRelacCred
    
    
    'agregado por vapi SEGÙN ERS TI-ERS001-2017
    If bPresol Then
        cmbProductoCMACM.ListIndex = IndiceListaCombo(cmbProductoCMACM, Mid(rsPresol!nConsValorProducto, 1, 1) & "00")
        cmbSubProducto.ListIndex = IndiceListaCombo(cmbSubProducto, rsPresol!nConsValorSubProducto)
        cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, rsPresol!nmoneda)
        cmbDestCred.ListIndex = IndiceListaCombo(cmbDestCred, rsPresol!nConsValorDestino)
        cmbAnalista.ListIndex = IndiceListaCombo(cmbAnalista, IIf(IsNull(rsPresol!cPersCodAnalista), "", rsPresol!cPersCodAnalista))
        txtMontoSol.Text = Format(rsPresol!nMonto, "#0.00")
        spnCuotas.valor = rsPresol!nCuotas
        spnPlazo.valor = rsPresol!nPlazo
        bPresol = False
    End If
    'fin vapi
    
    Call DefineCondicionCredito
    
End Sub

Private Sub cmdsalir_Click()
    Set MatCredRef = Nothing
    Set MatCredSust = Nothing
    bAmpliacion = False
    bRefinanciar = False
    bSustituirDeudor = False
    bRefinanciarSustituir = False
    '***Modificado por ELRO 20111017, según Acta 222-2011/TI-D
    PersoneriaTitular = 0
    '*********************************************************
    'JOEP20190131 CP
    ReDim nMatMontoPre(0)
    bEntrotxtMontoSol = False
    cmbSubDestCred.Visible = False
    nSubDestAnt = 0
    lblTpDoc.Visible = False
    cmbTpDoc.Visible = False
    'JOEP20190131 CP
    Unload frmCredRefinanc
    Unload Me
End Sub


Private Sub CmdVerFteIngreso_Click()
    If cmbFuentes.ListCount > 0 Then
        If cmbFuentes.ListIndex <> -1 Then
             FrmCredverFteIngreso.inicia (Right(cmbFuentes.Text, 8))
        Else
            MsgBox "No ha escogido ninguna fuente de Ingreso", vbInformation, "Mensaje"
        End If
    Else
        MsgBox "No existe fuentes de ingreso", vbInformation, "Mensaje"
    End If
End Sub

Private Sub Form_Load()
'JOEP20190919 ERS042 CP-2018
    MinMonto = 0
    MaxMonto = 0
    MaxCuota = 0
    MinCuota = 0
    MaxPlazo = 0
    MinPlazo = 0
    nTpCmbTpDoc = 0
    nSubDestAnt = 0
    ReDim nMatMontoPre(0)
    bEntrotxtMontoSol = False
    lblTpDoc.Visible = False
    cmbTpDoc.Visible = False
    cmbSubDestCred.Visible = False
'JOEP20190919 ERS042 CP-2018

    CentraForm Me
    'WIOR 20140509 ******************************
    fbRegPromotores = VerificaGruposRegPromotores(gsGruposUser)
    'WIOR FIN ***********************************
    'JUEZ 20160509 *************************
    '->***** LUCV20171013, Comentó según INC1710120009
    chkAutAmpliacion.Left = 4560
    'If fbRegPromotores Then
    '  chkAutAmpliacion.Left = 4560
    'Else
    '    chkAutAmpliacion.Left = 240
    'End If
    '<-***** LUCV20171013
    'END JUEZ ******************************
    Call CargaControles
    Call HabilitaIngresoSolicitud(False)
    Call ControlesPermiso
    If bRefinanciar Then
        ChkCap.Enabled = True
    End If
    
    'CMACICA_CSTS - 10112003 -----------------------------------------------------------
    If bSustituirDeudor Then
       ChkCap.Enabled = False
    End If
    '------------------------------------------------------------------------------------
   
    cmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    cmdGarantias.Enabled = False
    cmdGravar.Enabled = False
    cmdEvaluar.Enabled = False
    ActXCtaCred.CMAC = gsCodCMAC
    ActXCtaCred.Age = gsCodAge
    cmdEjecutar = -1
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCredRegistrarActualizaSoliCred
    Set oCredAgrico = New frmCredAgricoSelec 'WIOR 20130723
    fbActivo = False 'WIOR 20130723
    ReDim fvListaCompraDeuda(0) 'EJVG20160201 ERS002-2016
    ReDim fvListaAguaSaneamiento(0) 'EAAS20180912
    ReDim fvListaCreditoVerde(0) 'EAAS20191401 SEGUN 018-GM-DI_CMACM
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oRelPersCred = Nothing
    Set objPista = Nothing
    PersoneriaTitular = 0 'WIOR 20130719
    '***MARG ERS046-2016***AGREGADO 20161109***
    gsOpeCod = ""
    '***MARG ERS046-2016***
End Sub

'EJVG20130503 ***
Private Sub SpnCuotas_Change()
    Call DefineCondicionCredito
    Call EstableceCondicionSubProducto
''JOEP20190919 ERS042 CP-2018
    Call CP_ValidaProdUniCuota
    If cmbTpDoc.Visible = True Then
        spnCuotas.valor = IIf(spnCuotas.valor = "", 0, spnCuotas.valor)
        If bRefinanciar = False And spnCuotas.valor > MaxCuota Then
            MsgBox "La cuota máxima es " & MaxCuota, vbInformation, "Aviso"
            spnCuotas.valor = MaxCuota
            Exit Sub
        End If
    End If
''JOEP20190919 ERS042 CP-2018
End Sub
'END EJVG *******
Private Sub spnCuotas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'JOEP20190919 ERS042 CP-2018
    spnCuotas.valor = IIf(spnCuotas.valor = "", 0, spnCuotas.valor)
        If bRefinanciar = False And spnCuotas.valor < MinCuota Then
            MsgBox "La cuota mínima es " & MinCuota, vbInformation, "Aviso"
            spnCuotas.valor = MinCuota
            Exit Sub
        End If
        If bRefinanciar = False And spnCuotas.valor > MaxCuota Then
            MsgBox "La cuota máxima es " & MaxCuota, vbInformation, "Aviso"
            spnCuotas.valor = MaxCuota
            Exit Sub
        End If
'JOEP20190919 ERS042 CP-2018
        spnPlazo.SetFocus
    End If
End Sub

Private Sub spnPlazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    'JOEP20190919 ERS042 CP-2018
            If Right(cmbSubProducto.Text, 3) = 525 Then
                If spnPlazo.valor > MaxPlazo Then
                    MsgBox "El plazo máximo  es " & MaxPlazo & " días", vbInformation, "Aviso"
                    spnPlazo.valor = MaxPlazo
                    spnPlazo.SetFocus
                    Exit Sub
                End If
                If spnPlazo.valor < MinPlazo Then
                    MsgBox "El plazo mínimo es " & MinPlazo & " días", vbInformation, "Aviso"
                    spnPlazo.valor = MinPlazo
                    spnPlazo.SetFocus
                    Exit Sub
                End If
            End If
'JOEP20190919 ERS042 CP-2018
        If cmbDestCred.Visible And cmbDestCred.Enabled Then cmbDestCred.SetFocus
    End If
End Sub

Private Sub txtfechaAsig_GotFocus()
    fEnfoque txtfechaAsig
End Sub

Private Sub txtfechaAsig_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbAnalista.Enabled = True Then
            cmbAnalista.SetFocus
        End If
    End If
End Sub

Private Sub txtMontoSol_Change()
If cmbMoneda.Text <> "" Then Call PintaMoneda(Right(cmbMoneda.Text, 1))    'Agrego JOEP20190919 ERS042 CP-2018
'Comento JOEP20190919 ERS042 CP-2018
'    If cmbMoneda.ListIndex = 0 Then
'        txtMontoSol.ForeColor = &H289556
'    Else
'        txtMontoSol.ForeColor = vbBlue
'    End If
'Comento JOEP20190919 ERS042 CP-2018
End Sub

Private Sub txtMontoSol_GotFocus()
'Agrego JOEP20190919 ERS042 CP-2018
    If Not CP_Mensajes(2, "") Then Exit Sub
    If Not (CP_Mensajes(6, Trim(Right(cmbSubProducto.Text, 10)), False, False)) Then Exit Sub
    
    If cmdEjecutar = 2 Then
        If bEntrotxtMontoSol = False Then
            Call CP_CargaAporte(lblCodigo, Trim(Right(cmbSubProducto.Text, 5)), IIf(cmbSubDestCred.Text = "", "0", Trim(Right(cmbSubDestCred.Text, 9))), Trim(Right(cmbDestCred.Text, 9)), Trim(Right(cmbCondicion.Text, 10)), IIf(cmbTpDoc.Text = "", 0, Trim(Right(cmbTpDoc.Text, 9))), Trim(Right(cmbMoneda.Text, 3)), Trim(Right(cmbCondicionOtra, 3)))
        End If
    End If
    
'Agrego JOEP20190919 ERS042 CP-2018
    fEnfoque txtMontoSol
    bEntrotxtMontoSol = True 'Agrego JOEP20190919 ERS042 CP-2018
End Sub

Private Sub txtMontoSol_KeyPress(KeyAscii As Integer)
'JOEP20190919 ERS042 CP-2018
    txtMontoSol = IIf(txtMontoSol = "", 0, txtMontoSol)
    KeyAscii = NumerosDecimales(txtMontoSol, KeyAscii)
    
    If cmdEjecutar = 2 Then
        Call CP_CargaAporte(lblCodigo, Trim(Right(cmbSubProducto.Text, 5)), IIf(cmbSubDestCred.Text = "", 0, Trim(Right(cmbSubDestCred.Text, 9))), Trim(Right(cmbDestCred.Text, 9)), Trim(Right(cmbCondicion.Text, 10)))
        bEntrotxtMontoSol = True
    End If
 'JOEP20190919 ERS042 CP-2018
 
    If KeyAscii = 13 Then
'JOEP20190919 ERS042 CP-2018
    If CDbl(txtMontoSol.Text) < MinMonto Then
        MsgBox "Monto mínimo para el tipo de Producto " & IIf(Trim(Right(cmbMoneda.Text, 1)) = 1, "S/ ", "$ ") & Format(MinMonto, "#,#0.00"), vbInformation, "Aviso"
        txtMontoSol.Text = Format(MinMonto, "#0.00")
    End If
'JOEP20190919 ERS042 CP-2018
    'arlo20200429 begin
    If CDbl(txtMontoSol.Text) < MaxMonto Then
        MsgBox "Monto mínimo para el tipo de Producto " & IIf(Trim(Right(cmbMoneda.Text, 1)) = 1, "S/ ", "$ ") & Format(MaxMonto, "#,#0.00"), vbInformation, "Aviso"
        txtMontoSol.Text = Format(MaxMonto, "#0.00")
    End If
    'arlo20200429 end
        spnCuotas.SetFocus
        txtMontoSol.Text = Format(txtMontoSol.Text, "#0.00") 'JOEP20190919 ERS042 CP-2018
    End If
End Sub

Private Sub txtMontoSol_LostFocus()
bEntrotxtMontoSol = False 'JOEP20190919 ERS042 CP-2018
    If Len(Trim(txtMontoSol.Text)) = 0 Then
         txtMontoSol.Text = "0.00"
    End If
'JOEP20190919 ERS042 CP-2018
    If cmbDestCred.Text = "" And Trim(Right(cmbSubProducto.Text, 10)) = "521" And bRefinanciar = False And bAmpliacion = False Then Exit Sub
    
    If CDbl(txtMontoSol.Text) < MinMonto And cmbMoneda.Text <> "" Then
        MsgBox "Monto mínimo para el tipo de Producto " & IIf(Trim(Right(cmbMoneda.Text, 1)) = 1, "s/ ", "$ ") & Format(MinMonto, "#,#0.00"), vbInformation, "Aviso"
        txtMontoSol.Text = Format(MinMonto, "#0.00")
        txtMontoSol.SetFocus
    End If
'JOEP20190919 ERS042 CP-2018
    txtMontoSol.Text = Format(txtMontoSol.Text, "#0.00")
End Sub

'Function ValidaMontoAmpliacion(ByVal nMontoAmpliado As Double, ByVal nMonedaAmpliado As Integer, _
'                               ByVal nMonto As Double, ByVal nMoneda As Integer) As Boolean
'
'    Dim oAmpliado As COMDCredito.DCOMAmpliacion
'
'    Set oAmpliado = New COMDCredito.DCOMAmpliacion
'    ValidaMontoAmpliacion = oAmpliado.ValidaMontoAmpliado(nMontoAmpliado, nMonedaAmpliado, nMonto, nMoneda, gdFecSis)
'    Set oAmpliado = Nothing
'End Function
'
'Function ListaValidaMontoAmpliacion(ByVal rs As ADODB.Recordset, ByVal nMontoSol As Double, ByVal nMoneda As Integer) As Boolean
'    Dim oAmpliado As COMDCredito.DCOMAmpliacion
'
'    Set oAmpliado = New COMDCredito.DCOMAmpliacion
'    ListaValidaMontoAmpliacion = oAmpliado.ValidaMontoAmpliadoLista(rs, nMontoSol, nMoneda)
'    Set oAmpliado = Nothing
'End Function

'JUEZ 20130527 *******************************************************************
Private Function LlenaRecordSet_Cliente() As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim i As Integer
Set rs = New ADODB.Recordset

With rs
    .Fields.Append "codigo", adVarChar, 13
    .Fields.Append "relacion", adInteger
    .Open
    
    For i = 1 To ListaRelacion.ListItems.count
        .AddNew
        .Fields("codigo") = ListaRelacion.ListItems.Item(i).SubItems(2)
        .Fields("relacion") = ListaRelacion.ListItems.Item(i).SubItems(3) 'APRI2018 ERS036-2017
    Next i
End With

Set LlenaRecordSet_Cliente = rs
End Function
'END JUEZ ************************************************************************
'MIOL 20130625, SEGUN RQ13335 **
Private Sub txtVendedor_KeyPress(KeyAscii As Integer)
KeyAscii = Letras(KeyAscii)
End Sub
'END MIOL **********************
'WIOR 20140509 ************************
Private Function VerificaGruposRegPromotores(ByVal psGrupos As String) As Boolean
Dim oCredito As COMNCredito.NCOMCredito

Set oCredito = New COMNCredito.NCOMCredito
VerificaGruposRegPromotores = oCredito.VerificaGruposRegPromotores(psGrupos)
End Function
'WIOR FIN *****************************

'WIOR 20151222 ***
Private Sub CargaConfigMonedaTpoProd(ByVal psTipo As String)
Dim oCred As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Set oCred = New COMDCredito.DCOMCredito
Set rs = oCred.ObtieneConfigMonedaTpoProd(psTipo)
Set oCred = Nothing

'cmbMoneda.Enabled = True
If Not bAmpliacion Or rsAmpliado Is Nothing Then cmbMoneda.Enabled = True 'JUEZ 20160509
If Not (rs.EOF And rs.BOF) Then
    If rs.RecordCount = 1 Then
        cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, CInt(rs!nmoneda))
        cmbMoneda.Enabled = False
    End If
End If

End Sub
'WIOR FIN ********
'FRHU 20160702 ERS002-2016
Private Sub verificarAlertas(ByVal psPersCodTitular As String, ByVal psPersCodConyuCodeu As String)
    Dim oCred As New COMDCredito.DCOMCredito
    Dim rs As New ADODB.Recordset
        
    If psPersCodTitular <> "" Then
        Set rs = oCred.verificarAlertasRelaCred(psPersCodTitular, psPersCodConyuCodeu)
        If Not (rs.BOF And rs.EOF) Then
            If rs!nAlerta <> 0 Then
                MsgBox IIf(IsNull(rs!cTpoDesc), "", rs!cTpoDesc), vbInformation, "AVISO"
            End If
        End If
    End If
End Sub
'FIN FRHU 20160702
'RECO 2016072016 ***************************************************************
Private Function ValidaMultiForm(ByVal psProdCab As String) As Boolean
    Dim oGen As New COMDConstSistema.DCOMGeneral
    Dim nIndice As Integer
    Dim sProdCad As String
    Dim sProdCod As String
    sProdCad = oGen.LeeConstSistema(530)
    ValidaMultiForm = False
    
    For nIndice = 1 To Len(sProdCad)
        If Mid(sProdCad, nIndice, 1) <> "," Then
            sProdCod = sProdCod & Mid(sProdCad, nIndice, 1)
        Else
            If psProdCab = sProdCod Then
                ValidaMultiForm = True
                Exit Function
            End If
            sProdCod = ""
        End If
    Next
End Function
'RECO FIN **********************************************************************
'MARG ERS003-2018----------------------------------------
'Public Function ValidarScoreExperian(ByVal psNuevaCta As String, Optional ByVal MatCredRelaciones As Variant) As Boolean 'comment by marg201906
Public Function ValidarScoreExperian(ByVal psNuevaCta As String, Optional ByVal MatCredRelaciones As Variant, Optional ByRef pbExitoScore As Boolean = False, Optional ByVal pbEsConsultaExperian = True) As Boolean 'add by marg201906
    Dim oScore As COMDCredito.DCOMCredito
    Dim esAplicableValidacionScore As Boolean
    Dim bPermiteSugerencia As Boolean
    Set oScore = New COMDCredito.DCOMCredito
    bPermiteSugerencia = False
    esAplicableValidacionScore = oScore.esAplicableValidacionScore(psNuevaCta)
    If esAplicableValidacionScore Then
    
        Dim existeDataExperianEnBD As Boolean
        Dim esDataExperianActualizado As Boolean
        Dim cPersCodTitular As String
        Dim bExitoConsulta As Boolean
        Dim bExitoScore As Boolean
        Dim rsInforme As ADODB.Recordset
        Dim rsScore As ADODB.Recordset
        Dim rsCliente As ADODB.Recordset
        Dim MatScoreCliente As Variant
        Dim iCR As Integer
        Dim mensajeScore As String
        
        'CTI1 20180719 ***
        Dim rsRelFinales As ADODB.Recordset
        Dim cParamExtraXML As String
        Dim cParamExtraJSON As String
        Dim bPasarParamExperian As Boolean
        Set rsRelFinales = oScore.ObtenerRelacionesFinalesExperian(psNuevaCta)
        
        If Not rsRelFinales.BOF And Not rsRelFinales.EOF Then
        'CTI1 FIN ********
            bExitoConsulta = False
            'For iCR = 0 To UBound(MatCredRelaciones) - 1'COMENTADO POR CTI1 20180720
                'Set rsCliente = oScore.obtenerCliente(MatCredRelaciones(iCR, 0), "")'COMENTADO POR WIOR 20180720
            'CTI1 20180720 ***
            For iCR = 0 To rsRelFinales.RecordCount - 1
                'If Not rsCliente.BOF And Not rsCliente.EOF Then
                    cParamExtraJSON = ""
                    cParamExtraXML = ""
                    bPasarParamExperian = False
                
                    If Trim(rsRelFinales!cParamExtraJSON) <> "" Then
                        cParamExtraJSON = Trim(rsRelFinales!cParamExtraJSON)
                    End If
                    
                    If Trim(rsRelFinales!cParamExtraXML) <> "" Then
                        cParamExtraXML = Trim(rsRelFinales!cParamExtraXML)
                    End If
                    
                    If cParamExtraXML <> "" And cParamExtraJSON <> "" Then
                        bPasarParamExperian = True
                    End If
            'CTI1 FIN ********
                    
                    'existeDataExperianEnBD = oScore.existeDataExperianEnBD(CStr(rsCliente!cPersIDTpo), rsCliente!cPersIDnro, rsCliente!ApellidoPaterno, rsCliente!cPersCod) 'CTI1 20180802 COMENTÓ
                    existeDataExperianEnBD = oScore.existeDataExperianEnBD(CStr(rsRelFinales!cPersIDTpo), rsRelFinales!cPersIDnro, rsRelFinales!ApellidoPaterno, rsRelFinales!cPersCod, bPasarParamExperian, cParamExtraXML) 'CTI1 20180802
                    If existeDataExperianEnBD Then
                        'esDataExperianActualizado = oScore.esDataExperianActualizado(CStr(rsCliente!cPersIDTpo), rsCliente!cPersIDnro, rsCliente!ApellidoPaterno, rsCliente!cPersCod) 'CTI1 20180802 COMENTÓ
                        esDataExperianActualizado = oScore.esDataExperianActualizado(CStr(rsRelFinales!cPersIDTpo), rsRelFinales!cPersIDnro, rsRelFinales!ApellidoPaterno, rsRelFinales!cPersCod, bPasarParamExperian, cParamExtraXML) 'CTI1 20180802
                        If esDataExperianActualizado Then
                            bExitoConsulta = True
                        Else
                            'bExitoConsulta = ConsultarDataExperianOnline(CStr(rsCliente!cPersIDTpo), rsCliente!cPersIDnro, rsCliente!ApellidoPaterno, 1, 5, 1, gsCodUser, rsCliente!cPersCod, 1) 'CTI1 20180802 COMENTÓ
                            bExitoConsulta = ConsultarDataExperianOnline(CStr(rsRelFinales!cPersIDTpo), rsRelFinales!cPersIDnro, rsRelFinales!ApellidoPaterno, 1, 5, 1, gsCodUser, rsRelFinales!cPersCod, 1, bPasarParamExperian, cParamExtraJSON) 'CTI1 20180802
                        End If
                    Else
                        'bExitoConsulta = ConsultarDataExperianOnline(CStr(rsCliente!cPersIDTpo), rsCliente!cPersIDnro, rsCliente!ApellidoPaterno, 1, 5, 1, gsCodUser, rsCliente!cPersCod, 1)'CTI1 20180802 COMENTÓ
                        bExitoConsulta = ConsultarDataExperianOnline(CStr(rsRelFinales!cPersIDTpo), rsRelFinales!cPersIDnro, rsRelFinales!ApellidoPaterno, 1, 5, 1, gsCodUser, rsRelFinales!cPersCod, 1, bPasarParamExperian, cParamExtraJSON) 'CTI1 20180802
                    End If
                    If bExitoConsulta = False Then
                        'MsgBox "El Scoring No se pudo calcular por un problema en la conectividad. Consulte a través del aplicativo móvil al Titular y los relacionados al crédito para el cálculo de su Scoring", vbInformation, "SICMACT"
                        'Call showMensajeProblemaDataExperian 'comment by marg201906
                        Call showMensajeProblemaDataExperian(pbEsConsultaExperian) 'add by marg201906
                        oScore.EliminarInformeCredito (psNuevaCta)
                        Exit For
                    End If
                'Else
                '    'MsgBox "El Scoring No se pudo calcular por un problema en la conectividad. Consulte a través del aplicativo móvil al Titular y los relacionados al crédito para el cálculo de su Scoring", vbInformation, "SICMACT"
                '    Call showMensajeProblemaDataExperian
                '    Exit For
                'End If
                'rsCliente.Close
                'Set rsCliente = Nothing
                rsRelFinales.MoveNext
            Next iCR
    
            bExitoScore = False
            If bExitoConsulta = True Then
                ReDim MatScoreCliente(UBound(MatCredRelaciones), 4)
                oScore.EliminarInformeCredito (psNuevaCta)
                rsRelFinales.MoveFirst 'CTI1 20180802
                'For iCR = 0 To UBound(MatCredRelaciones) - 1'COMENTADO POR CTI120180720
                For iCR = 0 To rsRelFinales.RecordCount - 1 'CTI120180720
                    'CTI120180720 ***
                    cParamExtraJSON = ""
                    cParamExtraXML = ""
                    bPasarParamExperian = False
                
                    If Trim(rsRelFinales!cParamExtraJSON) <> "" Then
                        cParamExtraJSON = Trim(rsRelFinales!cParamExtraJSON)
                    End If
                    
                    If Trim(rsRelFinales!cParamExtraXML) <> "" Then
                        cParamExtraXML = Trim(rsRelFinales!cParamExtraXML)
                    End If
                    
                    If cParamExtraXML <> "" And cParamExtraJSON <> "" Then
                        bPasarParamExperian = True
                    End If
                    'CTI1 FIN *******
                
                    'consultar cliente
                    'Set rsCliente = oScore.obtenerCliente(MatCredRelaciones(iCR, 0), "")'COMENTADO POR CTI120180720
                    'If Not rsCliente.BOF And Not rsCliente.EOF Then
                        'consultar informe
                        'Set rsInforme = oScore.obtenerInforme(CStr(rsCliente!cPersIDTpo), rsCliente!cPersIDnro, rsCliente!ApellidoPaterno, rsCliente!cPersCod) 'CTI120180720 COMENTÓ
                        Set rsInforme = oScore.obtenerInforme(CStr(rsRelFinales!cPersIDTpo), rsRelFinales!cPersIDnro, rsRelFinales!ApellidoPaterno, rsRelFinales!cPersCod, bPasarParamExperian, cParamExtraXML)
                        'CTI120180720 - SE AGREGO LOS PARAMETROS bPasarParamExperian, cParamExtra
    
                        If Not rsInforme.BOF And Not rsInforme.EOF Then
                            'consultar score
                            Set rsScore = oScore.obtenerScore(rsInforme!nIdInforme)
                            If Not rsScore.BOF And Not rsScore.EOF Then
                                'asignar score al cliente
                                'MatScoreCliente(iCR, 0) = MatCredRelaciones(iCR, 0) 'codigo persona'COMENTADO POR CTI120180720
                                'MatScoreCliente(iCR, 1) = MatCredRelaciones(iCR, 1) 'relacion persona'COMENTADO POR CTI120180720
                                MatScoreCliente(iCR, 0) = rsRelFinales!cPersCod 'codigo persona'CTI120180802
                                MatScoreCliente(iCR, 1) = rsRelFinales!nPrdPersRelac 'relacion persona'CTI120180802
                                MatScoreCliente(iCR, 2) = rsScore!cPuntaje 'puntaje score
                                MatScoreCliente(iCR, 3) = rsScore!cClasificacion 'clasificacion score
                                bExitoScore = True
                                'oScore.RegistrarInformeCredito rsInforme!nIdInforme, psNuevaCta, CInt(MatCredRelaciones(iCR, 1))'CTI1 20180731 COMENTÓ
                                oScore.RegistrarInformeCredito rsInforme!nIdInforme, psNuevaCta, CInt(rsRelFinales!nPrdPersRelac) 'CTI120180802
                            Else
                                'MsgBox "El Scoring No se pudo calcular por un problema en la conectividad. Consulte a través del aplicativo móvil al Titular y los relacionados al crédito para el cálculo de su Scoring", vbInformation, "SICMACT"
                                'Call showMensajeProblemaDataExperian 'add by marg201906
                                Call showMensajeProblemaDataExperian(pbEsConsultaExperian) 'add by marg201906
                                oScore.EliminarInformeCredito (psNuevaCta)
                                bExitoScore = False
                                Exit For
                            End If
                            rsScore.Close
                            Set rsScore = Nothing
                        Else
                            'MsgBox "El Scoring No se pudo calcular por un problema en la conectividad. Consulte a través del aplicativo móvil al Titular y los relacionados al crédito para el cálculo de su Scoring", vbInformation, "SICMACT"
                            'Call showMensajeProblemaDataExperian 'comment by marg201906
                            Call showMensajeProblemaDataExperian(pbEsConsultaExperian) 'add by marg201906
                            oScore.EliminarInformeCredito (psNuevaCta)
                            bExitoScore = False
                            Exit For
                        End If
                        Call rsInforme.Close
                        Set rsInforme = Nothing
                    'Else
                    '    'MsgBox "El Scoring No se pudo calcular por un problema en la conectividad. Consulte a través del aplicativo móvil al Titular y los relacionados al crédito para el cálculo de su Scoring", vbInformation, "SICMACT"
                    '    Call showMensajeProblemaDataExperian
                    '    oScore.EliminarInformeCredito (psNuevaCta)
                    '    bExitoScore = False
                    '    Exit For
                    'End If
                    'rsCliente.Close
                    'Set rsCliente = Nothing
                    rsRelFinales.MoveNext 'CTI1 20180802
                Next iCR
            End If
    
            If bExitoScore = True Then
                'mensajeScore = getMensajeScore(MatScoreCliente)
                'MsgBox mensajeScore, vbInformation, "SICMACT"
                Set rsScore = oScore.getDecisionScore(psNuevaCta)
                If Not rsScore.BOF And Not rsScore.EOF Then
                    mensajeScore = rsScore!mensaje
                    bPermiteSugerencia = CBool(rsScore!bPermiteSugerencia)
                    pbExitoScore = bExitoScore 'add marg 201906
                    MsgBox mensajeScore, vbInformation, "AVISO"
                End If
                rsScore.Close
                Set rsScore = Nothing
            End If
        'CTI1 20180802 ***
        Else
            'Call showMensajeProblemaDataExperian 'comment by marg201906
            Call showMensajeProblemaDataExperian(pbEsConsultaExperian) 'add by marg201906
        End If
        'CTI1 FIN ********
    Else
        bPermiteSugerencia = True
    End If
    
    ValidarScoreExperian = bPermiteSugerencia
End Function
Private Function ConsultarScoreExperian(ByVal pcUserConsulta As String, ByVal pcPersCodCliente As String) As Boolean
    Set Req = New WinHttp.WinHttpRequest
    Dim postData As String
    Dim urlSimaynas As String
    urlSimaynas = Trim(LeeConstanteSist(708))
    postData = "cUserConsulta=" & pcUserConsulta & "&cPersCodCliente=" & pcPersCodCliente
    With Req
        '.Open "POST", "http://192.168.15.215:65229/ScoreCliente/ConsultarHistorialFromSicmacm?" & postData, Async:=False 'ojo cambiar a simaynas cuando pase a produccion
        .Open "POST", urlSimaynas & "/ScoreCliente/ConsultarHistorialFromSicmacm?" & postData, Async:=False 'ojo cambiar a simaynas cuando pase a produccion
        .Send
    End With
    If Req.Status = 200 Then
        ConsultarScoreExperian = True
    Else
        ConsultarScoreExperian = False
    End If
End Function
Private Function ConsultarDataExperianOnline(pcTipoid As String, pcId As String, pcApellido As String, _
                                             pnModalidad As Integer, pnTipoCred As Integer, _
                                             pnConsulta As Integer, pcUserConsulta As String, _
                                             pcPersCodCliente As String, pnSistema As Integer, _
                                             Optional ByVal pbDatosParam As Boolean = False, _
                                             Optional ByVal pcParamExtraJSON As String = "") As Boolean
    Dim postData As String
    Dim urlSimaynas As String
    urlSimaynas = Trim(LeeConstanteSist(708))
    'CTI1 20180720B AGREGO LOS SIGUIENTES PARAMETROS
    'Optional ByVal pbDatosParam As Boolean = False, Optional ByVal pMatParam As Variant = Null
                                             
    On Error GoTo ErrorConsultarDataExperianOnline
    Set Req = New WinHttp.WinHttpRequest

    postData = "cTipoid=" & pcTipoid & "&cId=" & pcId & "&cApellido=" & pcApellido & "&nModalidad=" & pnModalidad & _
               "&nTipoCred=" & pnTipoCred & "&nConsulta=" & pnConsulta & "&cUserConsulta=" & pcUserConsulta & _
               "&cPersCodCliente=" & pcPersCodCliente & "&nSistema=" & pnSistema
    
    'CTI1 20180720 ***
    If pbDatosParam Then
        If Trim(pcParamExtraJSON) <> "" Then
            postData = postData & "&bDatosParam=true"
            postData = postData & "&paramExtras=" & pcParamExtraJSON
        End If
    End If
    'CTI1 ************
    
    With Req
        '.Open "POST", "http://192.168.15.215:65229/ScoreCliente/ConsultarDataExperianOnline?" & postData, Async:=False 'ojo cambiar a simaynas cuando pase a produccion
        .Open "POST", urlSimaynas & "/ScoreCliente/ConsultarDataExperianOnline?" & postData, Async:=False 'ojo cambiar a simaynas cuando pase a produccion
        .Send
    End With
    If Req.Status = 200 Then
        ConsultarDataExperianOnline = True
    Else
        ConsultarDataExperianOnline = False
    End If
    Exit Function
ErrorConsultarDataExperianOnline:
    ConsultarDataExperianOnline = False
End Function
Private Sub showMensajeProblemaDataExperian(Optional ByVal pbEsConsultaExperian As Boolean = True) 'change marg201906
    'MsgBox "El Scoring No se pudo calcular por un problema en la conectividad. Consulte a través del Simaynas o Aplicativo Movil al Titular y los relacionados al crédito para el cálculo de su Scoring", vbInformation, "AVISO" 'comment by marg 201906
    
    'add by marg
    If pbEsConsultaExperian Then
        MsgBox "El Scoring No se pudo calcular por un problema en la conectividad. Consulte a través del Simaynas o Aplicativo Movil al Titular y los relacionados al crédito para el cálculo de su Scoring", vbInformation, "AVISO"
    End If
    '''
End Sub
'END MARG------------------------------------------------
'add pti1 ers0702018 18/12/2018
Private Sub ActAutDatos()
  Dim oNPersona As New COMNPersona.NCOMPersona
    Dim lbClienteActualizoAutorizoSusDatos As ADODB.Recordset
    Dim lbClienteEsPersonaNatural As Boolean
    nInicioActDa = 0
    If lblCodigo.Caption = "" Then
        Exit Sub
    End If
    
    lbClienteEsPersonaNatural = oNPersona.ClienteEsPersonaNatural(Trim(lblCodigo.Caption))
    
      If lbClienteEsPersonaNatural Then
            Set lbClienteActualizoAutorizoSusDatos = oNPersona.ClienteActualizoAutorizoSusDatos(Trim(lblCodigo.Caption))
               If Not (lbClienteActualizoAutorizoSusDatos.EOF And lbClienteActualizoAutorizoSusDatos.BOF) Then
                'ya se registro el usuario
                Dim autorizos As Integer
                autorizos = Trim(lbClienteActualizoAutorizoSusDatos!nAutorizaUsoDatos)
        
                
                If autorizos = 1 Then
                CboAutoriazaUsoDatos.Visible = False
                lblAutorizarUsoDatos.Visible = False
                CboAutoriazaUsoDatos.ListIndex = 0
                nInicioActDa = 1
                Else
                CboAutoriazaUsoDatos.Visible = True
                lblAutorizarUsoDatos.Visible = True
                CboAutoriazaUsoDatos.ListIndex = 1
                nInicioActDa = 0
                End If
                
            Else
             CboAutoriazaUsoDatos.Visible = True
             lblAutorizarUsoDatos.Visible = True
             CboAutoriazaUsoDatos.ListIndex = -1
             nInicioActDa = -1
            End If
            

        End If
End Sub

Private Function BuscaNombre(ByVal psNombre As String, ByVal nTipoBusqueda As TiposBusquedaNombre) As String 'add pti1 ers070-2018 18/12/2018
Dim sCadTmp As String
Dim PosIni As Integer
Dim PosFin As Integer
Dim PosIni2 As Integer
    sCadTmp = ""
    Select Case nTipoBusqueda
        Case 1 'Busqueda de Apellido Paterno
            If Mid(psNombre, 1, 1) <> "/" And Mid(psNombre, 1, 1) <> "\" And Mid(psNombre, 1, 1) <> "," Then
                PosIni = 1
                PosFin = InStr(1, psNombre, "/")
                If PosFin = 0 Then
                    PosFin = InStr(1, psNombre, "\")
                    If PosFin = 0 Then
                        PosFin = InStr(1, psNombre, ",")
                        If PosFin = 0 Then
                            PosFin = Len(psNombre)
                        End If
                    End If
                End If
                sCadTmp = Mid(psNombre, PosIni, PosFin - PosIni)
            Else
                sCadTmp = ""
            End If
        Case 2 'Apellido materno
           PosIni = InStr(1, psNombre, "/")
           If PosIni <> 0 Then
                PosIni = PosIni + 1
                PosFin = InStr(1, psNombre, "\")
                If PosFin = 0 Then
                    PosFin = InStr(1, psNombre, ",")
                    If PosFin = 0 Then
                        PosFin = Len(psNombre)
                    End If
                End If
                sCadTmp = Mid(psNombre, PosIni, PosFin - PosIni)
            Else
                sCadTmp = ""
            End If
        Case 3 'Apellido de casada
           PosIni = InStr(1, psNombre, "\")
           If PosIni <> 0 Then
                PosIni2 = InStr(1, psNombre, "VDA")
                If PosIni2 <> 0 Then
                    PosIni = PosIni2 + 3
                    PosFin = InStr(1, psNombre, ",")
                    If PosFin = 0 Then
                        PosFin = Len(psNombre)
                    End If
                Else
                    PosIni = PosIni + 1
                    PosFin = InStr(1, psNombre, ",")
                    If PosFin = 0 Then
                        PosFin = Len(psNombre)
                    End If
                End If
                sCadTmp = Trim(Mid(psNombre, PosIni, PosFin - PosIni))
            Else
                sCadTmp = ""
            End If
        Case 4 'Nombres
            PosIni = InStr(1, psNombre, ",")
            If PosIni <> 0 Then
                PosIni = PosIni + 1
                PosFin = Len(psNombre)
                sCadTmp = Mid(psNombre, PosIni, (PosFin + 1) - PosIni)
            Else
                sCadTmp = ""
            End If
            
    End Select
    BuscaNombre = sCadTmp
End Function
'ADD PTI1 23/08/2018 ERS027-2017
Property Get getcPersCodPreSol() As String
getcPersCodPreSol = cPersCodPreSol
End Property

Property Get getbfCredSolicitud() As String
getbfCredSolicitud = bfCredSolicitud
End Property

Property Get getbAmpliacion() As Boolean
getbAmpliacion = bAmpliacion
End Property
'END PTI1

'add PTI1 ERS070-2018 11/12/2018
Private Sub ImprimirPdfCartillaAutorizacion()
    Dim sParrafoUno As String
    Dim sParrafoDos As String
    Dim sParrafoTres As String
    Dim sParrafoCuatro As String
    Dim sParrafoCinco As String
    Dim sParrafoSeis As String
    Dim sParrafoSiete As String
    Dim sParrafoOcho As String
    Dim oDoc As cPDF
    Dim nAltura As Integer
    
    Set oDoc = New cPDF
    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Cartilla Autorización y Actualización de datos personales"
    oDoc.Title = "Cartilla Autorización y Actualización de datos personales"
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\CartillaAutorizacionActualizacionDeDatos" & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    
    'oDoc.LoadImageFromFile App.path & "\logo_cmacmaynas.bmp", "Logo"
    oDoc.LoadImageFromFile App.Path & "\Logo_2015.jpg", "Logo" 'O
    
    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical
    '<body>
    nAltura = 20
    oDoc.WTextBox 10, 10, 780, 575, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
    'oDoc.WImage 60, 480, 35, 100, "Logo"
    oDoc.WImage 70, 460, 50, 100, "Logo" 'O
    oDoc.WTextBox 90, 50, 15, 500, "AUTORIZACIÓN PARA EL TRATAMIENTO DE DATOS PERSONALES", "F2", 11, hCenter 'agregado por pti1 ers070-2018 05/12/2018
     
    
    oDoc.WTextBox 125, 56, 360, 520, (MatPersona(1).sNombres & " " & MatPersona(1).sApePat & " " & MatPersona(1).sApeMat & IIf(Len(MatPersona(1).sApeCas) = 0, "", " " & IIf(MatPersona(1).sEstadoCivil = "2", "DE", "VDA") & " " & MatPersona(1).sApeCas)), "F1", 11, hjustify
    oDoc.WTextBox 125, 484, 360, 520, (Trim(MatPersona(1).sPersIDnro)), "F1", 11, hjustify
    oDoc.WTextBox 125, 56, 360, 520, ("___________________________________________________________"), "F1", 11, hjustify
    oDoc.WTextBox 125, 481, 360, 520, ("____________"), "F1", 11, hjustify
    oDoc.WTextBox 125, 35, 10, 520, ("Yo, " & String(120, vbTab) & "  con DOI N° " & String(22, vbTab) & ""), "F1", 11, hjustify
    
    'sParrafoUno = "Yo " & String(78, vbTab) & "  con DOI N° " & String(20, vbTab) & "autorizo y otorgo por tiempo  "
    sParrafoUno = "autorizo y otorgo por tiempo indefinido, " & String(0.52, vbTab) & "mi consentimiento libre, previo, expreso, inequívoco e informado a" & Chr$(13) & _
                   "la " & String(0.52, vbTab) & "CAJA MUNICIPAL DE AHORRO Y CRÉDITO DE MAYNAS " & String(0.52, vbTab) & "S.A. " & String(0.52, vbTab) & "(en " & String(0.52, vbTab) & "adelante," & String(0.52, vbTab) & " ""LA CAJA""), " & String(0.51, vbTab) & " para " & String(0.51, vbTab) & " el" & Chr$(13) & _
                   "tratamiento de mis datos personales proporcionados " & String(0.7, vbTab) & " en contexto de la contratación de cualquier producto " & Chr$(13) & _
                   "(activo y/o pasivo)" & String(0.52, vbTab) & " o" & String(0.51, vbTab) & " servicio, " & String(0.52, vbTab) & " así " & String(0.52, vbTab) & "como " & String(0.52, vbTab) & "resultado" & String(0.52, vbTab) & "de " & String(0.52, vbTab) & " la suscripción de contratos, " & String(0.52, vbTab) & " formularios, " & String(0.52, vbTab) & " y a los " & Chr$(13) & _
                   "recopilados anteriormente, actualmente y/o por recopilar por " & String(0.52, vbTab) & "LA CAJA. " & String(0.53, vbTab) & "Asimismo, " & String(0.53, vbTab) & "otorgo " & String(0.53, vbTab) & "mi autorización" & Chr$(13) & _
                   "para el envío de información  promocional y/o publicitaria de los servicios y productos que" & String(0.53, vbTab) & " LA CAJA ofrece, " & Chr$(13) & _
                   "a tráves de cualquier medio de comunicación que se considere apropiado para su difusión, " & String(0.53, vbTab) & "y " & String(0.52, vbTab) & "para" & String(0.53, vbTab) & " su uso " & Chr$(13) & _
                   "en la gestión administrativa " & String(0.53, vbTab) & " y " & String(0.5, vbTab) & " comercial de  " & String(0.53, vbTab) & "LA  " & String(0.53, vbTab) & "CAJA " & String(0.53, vbTab) & " que guarde relación con su objeto social.  " & String(0.53, vbTab) & "En " & String(0.52, vbTab) & "ese " & Chr$(13) & _
                   "sentido, autorizo a LA CAJA al uso de mis datos personales para tratamientos que supongan el " & String(0.52, vbTab) & "desarrollo" & Chr$(13) & _
                   "de acciones y actividades comerciales, incluyendo la realización de estudios  de  mercado, " & String(0.53, vbTab) & " elaboración " & String(0.52, vbTab) & "de" & Chr$(13) & _
                   "perfiles de compra " & String(0.53, vbTab) & " y evaluaciones financieras. " & String(0.54, vbTab) & " El uso y tratamiento de mis datos personales, " & String(0.54, vbTab) & "se sujetan" & Chr$(13) & _
                   "a lo establecido por el artículo 13° de la Ley N° 29733 - Ley de Protección de Datos Personales."
    
  
    sParrafoDos = "Declaro conocer el compromiso de " & String(0.52, vbTab) & "LA CAJA " & String(0.52, vbTab) & " por garantizar el mantenimiento de la confidencialidad" & String(0.52, vbTab) & " y " & String(0.52, vbTab) & "el " & Chr$(13) & _
                  "tratamiento seguro de mis datos personales, incluyendo el resguardo en las transferencias de " & String(0.52, vbTab) & "los mismos, " & Chr$(13) & _
                  "que se realicen " & String(0.53, vbTab) & "en cumplimiento de la " & String(0.55, vbTab) & " Ley N° 29733 - Ley de Protección " & String(0.53, vbTab) & " de Datos Personales. De" & String(0.53, vbTab) & "igual " & Chr$(13) & _
                  "manera, declaro " & String(0.52, vbTab) & "conocer que los datos personales " & String(0.55, vbTab) & "proporcionados por mi persona serán incorporados " & String(0.52, vbTab) & "al " & Chr$(13) & _
                  "Banco de Datos de Clientes de  " & String(0.6, vbTab) & " LA CAJA, el cual  " & String(0.55, vbTab) & "se encuentra debidamente registrado ante la" & String(0.52, vbTab) & " Dirección " & Chr$(13) & _
                  "Nacional  " & String(0.55, vbTab) & " de  " & String(0.55, vbTab) & " Protección de Datos " & String(0.55, vbTab) & "Personales, para lo cual " & String(0.55, vbTab) & " autorizo a LA CAJA " & String(0.52, vbTab) & "que " & String(0.55, vbTab) & " recopile, registre, " & Chr$(13) & _
                  "organice, " & String(0.55, vbTab) & "almacene, " & String(0.55, vbTab) & "conserve, bloquee, suprima, extraiga, consulte, utilice, transfiera, exporte, importe" & String(0.52, vbTab) & " o " & Chr$(13) & _
                  "procese de cualquier otra forma mis datos personales, con las limitaciones que prevé la Ley."
                 
                 
    sParrafoTres = "Del mismo modo, y siempre que así lo estime necesario, declaro conocer que podré ejercitar mis derechos " & Chr$(13) & _
                   "de " & String(0.55, vbTab) & " acceso, " & String(0.56, vbTab) & " rectificación, " & String(0.58, vbTab) & " cancelación " & String(0.55, vbTab) & " y " & String(0.55, vbTab) & " oposición relativos a este tratamiento, de conformidad " & String(0.52, vbTab) & "con lo " & Chr$(13) & _
                   "establecido" & String(0.51, vbTab) & " en " & String(0.5, vbTab) & "el " & String(0.6, vbTab) & " Titulo" & String(0.54, vbTab) & " III " & String(0.54, vbTab) & " de la Ley N° 29733 - Ley de Protección de Datos " & String(0.52, vbTab) & " Personales" & String(0.52, vbTab) & " acercándome " & Chr$(13) & _
                   "a cualquiera de las Agencias de LA CAJA a nivel nacional."

   sParrafoCuatro = "Asimismo, " & String(1.4, vbTab) & " declaro " & String(1.4, vbTab) & " conocer " & String(1.4, vbTab) & " el " & String(1.4, vbTab) & "compromiso " & String(1.4, vbTab) & " de " & String(1.4, vbTab) & " LA " & String(1.4, vbTab) & "CAJA " & String(1.4, vbTab) & " por " & String(1.4, vbTab) & "respetar " & String(1.4, vbTab) & "los " & String(1.4, vbTab) & "principios " & String(1.4, vbTab) & "de " & String(1.4, vbTab) & " legalidad, " & Chr$(13) & _
                    "consentimiento, finalidad, proporcionalidad, calidad, disposición de recurso, y nivel de protección adecuado," & Chr$(13) & _
                    "conforme lo dispone la Ley N° 29733 - Ley de Protección de Datos Personales," & String(1.4, vbTab) & " para " & String(1.4, vbTab) & "el " & String(1.4, vbTab) & "tratamiento de los" & Chr$(13) & _
                    "datos personales otorgados por mi persona."
                  
    sParrafoCinco = "Esta autorización es" & String(1.5, vbTab) & " indefinida y se mantendrá inclusive" & String(0.5, vbTab) & " después de terminada(s) la(s) operación(es)" & String(0.52, vbTab) & " y/o " & Chr$(13) & _
                    "el(los) Contrato(s) que tenga" & String(1.5, vbTab) & " o pueda tener con LA CAJA" & String(1.3, vbTab) & " sin perjuicio de " & String(0.5, vbTab) & "poder ejercer mis derechos " & String(0.52, vbTab) & "de " & Chr$(13) & _
                    "acceso, rectificación, cancelación y oposición mencionados en el presente documento."
     
     Dim cfecha  As String 'pti1 add
      
     cfecha = Choose(Month(dfreg), "Enero", "Febrero", "Marzo", "Abril", _
                                        "Mayo", "Junio", "Julio", "Agosto", _
                                        "Setiembre", "Octubre", "Noviembre", "Diciembre")
            Dim nTamanio As Integer
            Dim Spac As Integer
            Dim Index As Integer
            Dim Princ As Integer
            Dim CantCarac As Integer
            Dim txtcDescrip As String
            Dim contador As Integer
            Dim nCentrar As Integer
            Dim nTamLet As Integer
            Dim spacvar As Integer
             
           
           
            nTamanio = Len(sParrafoUno)
            spacvar = 23
            Spac = 138
            Index = 1
            Princ = 1
            CantCarac = 0
            
            nTamLet = 6: contador = 0: nCentrar = 80
            
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoUno, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoUno, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoUno, Index, CantCarac)
                        oDoc.WTextBox Spac, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoUno, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoUno, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoUno, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            nTamanio = Len(sParrafoDos)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoDos, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoDos, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoDos, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoDos, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoDos, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoDos, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            nTamanio = Len(sParrafoTres)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoTres, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoTres, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoTres, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoTres, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoTres, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoTres, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            nTamanio = Len(sParrafoCuatro)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoCuatro, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoCuatro, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoCuatro, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoCuatro, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoCuatro, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoCuatro, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            
            nTamanio = Len(sParrafoCinco)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoCinco, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoCinco, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoCinco, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoCinco, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoCinco, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoCinco, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
             '******************PTI1  2018/12/20 ERS070-2018
                    
    
                  'oDoc.WTextBox 136, 35, 120, 520, sParrafoUno, "F1", 11, hjustify, , , 4, vbBlue
                  'oDoc.WTextBox 310, 35, 88, 520, sParrafoDos, "F1", 11, hjustify, , , 1, vbBlue
                  'oDoc.WTextBox 430, 35, 44, 520, sParrafoTres, "F1", 11, hjustify, , , 1, vbBlue
                  'oDoc.WTextBox 485, 35, 44, 520, sParrafoCuatro, "F1", 11, hjustify, , , 1, vbBlue
                  'oDoc.WTextBox 560, 35, 33, 520, sParrafoCinco, "F1", 11, hjustify
    


    oDoc.WTextBox 610, 35, 60, 300, ("En " & sAgeReg & " a los " & Day(dfreg) & " días del mes de " & cfecha & " de " & Year(dfreg)) & ".", "F1", 11, hLeft 'O  agregado  por pti1
    oDoc.WTextBox 670, 35, 90, 200, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    oDoc.WTextBox 730, 35, 60, 180, "________________________________________", "F1", 8, hCenter
    oDoc.WTextBox 745, 90, 60, 80, "Firma", "F1", 10, hCenter
    
    sParrafoSeis = "¿Autorizas a Caja Maynas para el tratamiento de sus datos personales?"
    
    oDoc.WTextBox 670, 280, 60, 250, sParrafoSeis, "F1", 11, hLeft 'O  agregado  por pti1
   
   
    oDoc.WTextBox 712, 300, 15, 20, "SI", "F1", 8, hCenter
    oDoc.WTextBox 742, 300, 15, 20, "NO", "F1", 8, hCenter
    
    oDoc.WTextBox 690, 420, 70, 80, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    oDoc.WTextBox 740, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    oDoc.WTextBox 710, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    If ultEstado = 1 Then
        oDoc.WTextBox 710, 280, 15, 20, "X", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
        oDoc.WTextBox 710, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    Else
        oDoc.WTextBox 740, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
        oDoc.WTextBox 740, 280, 15, 20, "X", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
    End If
    

            
    oDoc.PDFClose
    oDoc.Show
    '</body>
End Sub


'Agrego JOEP20190919 ERS042 CP-2018
Private Sub CatalogoLlenaCombox(ByVal nParProd As Integer, ByVal nParCod As Integer, Optional ByVal nDestino As Long = 0, Optional ByVal nCondicion As Long = 0)
Dim objCatalogoLlenaCombox As COMDCredito.DCOMCredito
Dim rsCatalogoCombox As ADODB.Recordset
Dim nDestinoAct As Integer
Dim nIdCampana As Integer 'arlo20200429
nIdCampana = IIf(Trim(Right(cmbCondicionOtra.Text, 10)) = "", 0, Trim(Right(cmbCondicionOtra.Text, 10))) 'arlo20200429
Set objCatalogoLlenaCombox = New COMDCredito.DCOMCredito
Set rsCatalogoCombox = objCatalogoLlenaCombox.getCatalogoCombo(nParProd, nParCod, nDestino, nCondicion, nIdCampana) 'arlo20200429

    If Not (rsCatalogoCombox.BOF And rsCatalogoCombox.EOF) Then
       If cmdEjecutar = 1 Then
         Select Case nParCod
            Case 5000
                Call Llenar_Combo_con_Recordset(rsCatalogoCombox, cmbMoneda)
            Case 2000
                Call Llenar_Combo_con_Recordset(rsCatalogoCombox, cmbDestCred)
            End Select
        ElseIf cmdEjecutar = -1 Then
            Call Llenar_Combo_con_Recordset(rsCatalogoCombox, cmbDestCred)
            cmbDestCred.ListIndex = IndiceListaCombo(cmbDestCred, nDestino)
        End If
    Else
    
    Set rsCatalogoCombox = objCatalogoLlenaCombox.GetCatalogoParametro(nParCod)
    
        Select Case nParCod
            Case 5000
                cmbMoneda.Clear
            Case 2000
                cmbDestCred.Clear
        End Select
        MsgBox "No existe configuración de Producto " & Trim(Left(cmbSubProducto.Text, 150)) & " - " & rsCatalogoCombox!cConsDescripcion & " Comuníquese con el Área de Productos Crediticios", vbInformation, "Aviso"
    End If
Set objCatalogoLlenaCombox = Nothing
RSClose rsCatalogoCombox
End Sub

Private Function CatalogoValidador(ByVal nParCod As Currency, Optional ByVal nMonto As Double = 0, Optional ByVal nCuota As Integer = 0, Optional ByVal nPlazo As Integer = 0, Optional ByVal nDestino As Long = 0, Optional ByVal nSubDest As Long = 0, Optional ByVal nTpDoc As Long = 0, Optional ByVal nTpIngr As Long = 0, Optional ByVal cTpPla As String = "") As Boolean
Dim objCatalogoDestino As COMDCredito.DCOMCredito
Dim rsCatalogoDestino As ADODB.Recordset
On Error GoTo ErrorCatalogoValidador
cValorIni = "0"
CatalogoValidador = True

    If cmbSubProducto.ListIndex = -1 Then
        If cmdEjecutar = -1 Then
            MsgBox "Verificar la migracion del Prodcuto", vbInformation, "Aviso"
        Else
            MsgBox "Debe Selecionar el Producto", vbInformation, "Aviso"
        End If
        
        If cmbSubProducto.Enabled = True Then
            cmbSubProducto.SetFocus
        End If
        Exit Function
    End If
  
    Set objCatalogoDestino = New COMDCredito.DCOMCredito
    Set rsCatalogoDestino = objCatalogoDestino.getCatalogoValidador(CInt(Trim(Right(cmbSubProducto.Text, 10))), nParCod, Trim(Right(cmbMoneda.Text, 2)), nMonto, nCuota, nPlazo, nDestino, nSubDest, nTpDoc, nTpIngr, cTpPla, gModSolicitud, IIf(bRefinanciar = False, 0, 1), IIf(bAmpliacion = False, 0, 1))
        
    If Not (rsCatalogoDestino.BOF And rsCatalogoDestino.EOF) Then
        If rsCatalogoDestino!mensaje <> "" Then
            If rsCatalogoDestino!ValorMax = 0 Then
                MsgBox rsCatalogoDestino!mensaje, vbInformation, "AVISO"
            Else
               MsgBox rsCatalogoDestino!mensaje & " " & IIf(nParCod = 4000, Format(rsCatalogoDestino!ValorMax, "#,#0.00"), rsCatalogoDestino!ValorMax), vbInformation, "AVISO"
            End If
            CatalogoValidador = False
            cValorIni = rsCatalogoDestino!ValorMax
        End If
    End If
Set objCatalogoDestino = Nothing
RSClose rsCatalogoDestino
    Exit Function
ErrorCatalogoValidador:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Public Function CP_validadConfiguracion() As Boolean
Dim objValConfProdCred As COMDCredito.DCOMCredito
Dim rsConfProCred As ADODB.Recordset
Dim sPersonaRelacNomb As String
Dim sPersonaCod As String
Dim sPersRelacCod As Integer
Dim i As Integer
Dim CTA As String
Set objValConfProdCred = New COMDCredito.DCOMCredito
Dim rsAmpliadoCatalogo As ADODB.Recordset 'ARLO20190510
Dim nIdCampana As Integer 'arlo20200429
nIdCampana = IIf(Trim(Right(cmbCondicionOtra.Text, 10)) = "", 0, Trim(Right(cmbCondicionOtra.Text, 10))) 'arlo20200429

CP_validadConfiguracion = True
Set rsConfProCred = objValConfProdCred.ValidaIfExistReqCond(Trim(Right(cmbSubProducto.Text, 10)))

If Not (rsConfProCred.BOF Or rsConfProCred.EOF) Then
'Validaciones Solicitados Por Usuario
CTA = ""
If Not rsAmpliado Is Nothing Then
    
    Set rsAmpliadoCatalogo = rsAmpliado.Clone 'ARLO20190510
    'ARLO CAMBIO rsAmpliado A rsAmpliadoCatalogo
    If Trim(Right(cmbSubProducto.Text, 10)) = "507" And bAmpliacion = True Then
        rsAmpliadoCatalogo.MoveFirst
        For i = 1 To rsAmpliadoCatalogo.RecordCount
            CTA = IIf(i = 1, "", CTA) & IIf(i = 1, "", ",") & rsAmpliadoCatalogo("cCtaCod")
            rsAmpliadoCatalogo.MoveNext
        Next i
    End If
End If

    oRelPersCred.IniciarMatriz
    Do While Not oRelPersCred.EOF
        sPersRelacCod = oRelPersCred.ObtenerValorRelac
        sPersonaRelacNomb = oRelPersCred.ObtenerRelac
        sPersonaCod = oRelPersCred.ObtenerMatrizRelacionesRelacion(sPersRelacCod, 1)
        Set rsConfProCred = objValConfProdCred.ValidaCondExtras(Trim(Right(cmbSubProducto.Text, 10)), sPersonaCod, sPersRelacCod, Trim(Right(cmbCondicion.Text, 5)), Trim(Right(cmbCondicionOtra.Text, 3)), Trim(Right(cmbDestCred, 3)), CTA, IIf(cmbTpDoc.Visible = True, Right(cmbTpDoc.Text, 9), 0), Right(cmbMoneda.Text, 3), gsCodAge, IIf(bRefinanciar = False, 0, 1), IIf(bAmpliacion = False, 0, 1), nIdCampana)
        If Not (rsConfProCred.BOF Or rsConfProCred.EOF) Then
            If rsConfProCred!cTpoDesc <> "" Then
                MsgBox "" & rsConfProCred!cTpoDesc & "", vbInformation, "Aviso"
                CP_validadConfiguracion = False
                Exit Function
            End If
        End If
        oRelPersCred.siguiente
    Loop
    
    i = 1
    'Titular,Conyuge,Codeudor,Garante,Representante
    oRelPersCred.IniciarMatriz
    For i = 1 To ListaRelacion.ListItems.count
        sPersRelacCod = oRelPersCred.ObtenerValorRelac
        sPersonaRelacNomb = oRelPersCred.ObtenerRelac
        If sPersRelacCod = 20 Or sPersRelacCod = 21 Or sPersRelacCod = 22 Or sPersRelacCod = 23 Or sPersRelacCod = 24 Or sPersRelacCod = 25 Then
            sPersonaCod = oRelPersCred.ObtenerMatrizRelacionesRelacion(sPersRelacCod, 1)
            If Trim(sPersonaCod) <> "" Then
                Set rsConfProCred = objValConfProdCred.ValidaRequisitos(Trim(Right(cmbSubProducto.Text, 10)), sPersonaCod, sPersRelacCod, Trim(Right(cmbCondicion.Text, 5)), chkAutAmpliacion.value, nTpIngr, IIf(bRefinanciar = False, 0, 1), IIf(bAmpliacion = False, 0, 1))
                If Not (rsConfProCred.BOF Or rsConfProCred.EOF) Then
                    If rsConfProCred!cParDescripcion <> "" Then
                        MsgBox "El " & sPersonaRelacNomb & " [" & oRelPersCred.ObtenerNombre & "] no cumple con los requisitos para una solicitud, ver " & rsConfProCred!cParDescripcion & "", vbInformation, "Aviso"
                        CP_validadConfiguracion = False
                        Exit Function
                    End If
                End If
            End If
        End If
    oRelPersCred.siguiente
    Next i
Else
    MsgBox "No existe configuración de este producto " & Trim(Left(cmbSubProducto.Text, 150)) & " Comuníquese con el Área de Productos Crediticios", vbInformation, "Aviso"
    CP_validadConfiguracion = False
    Exit Function
End If

Set objValConfProdCred = Nothing
RSClose rsConfProCred
End Function

Private Sub limpiaCatalogo()
    txtMontoSol.Text = "0.00"
    spnCuotas.valor = "0"
    spnPlazo.valor = "0"
    cmdDestinoDetalle.Visible = False
End Sub

Private Sub CP_HabilitaControles(ByVal Habilitar As Boolean)
    cmbSubProducto.Enabled = Habilitar
    cmbMoneda.Enabled = Habilitar
    txtMontoSol.Enabled = Habilitar
    spnCuotas.Enabled = Habilitar
    spnPlazo.Enabled = Habilitar
    cmbDestCred.Enabled = Habilitar
    cmdDestinoDetalle.Enabled = Habilitar
    cmbAnalista.Enabled = Habilitar
End Sub
Private Sub PintaMoneda(ByVal nTpMoneda As Integer)
    If nTpMoneda = 2 Then
        txtMontoSol.ForeColor = &H289556
    Else
        txtMontoSol.ForeColor = vbBlue
    End If
End Sub

Private Sub CP_LimpiaCombo()
    cmbMoneda.Clear
    cmbMoneda.Enabled = False
    txtMontoSol.Text = "0.00"
    txtMontoSol.Enabled = False
    spnCuotas.valor = 0
    spnCuotas.Enabled = False
    spnPlazo.valor = 0
    spnPlazo.Enabled = False
    cmbDestCred.Clear
    cmbDestCred.Enabled = False
    cmdDestinoDetalleAguaS.Visible = False
    cmbSubDestCred.Clear
    cmbSubDestCred.Visible = False
End Sub
Private Sub CP_DatosDefaut(ByVal cParCod As Currency)
    Dim oDCred As COMDCredito.DCOMCredito
    Dim rsDefaut As ADODB.Recordset
    Set oDCred = New COMDCredito.DCOMCredito
    Dim nTpDocDatDef As Long
    Dim nIdCampana As Integer 'arlo20200429
    nIdCampana = IIf(Trim(Right(cmbCondicionOtra.Text, 10)) = "", 0, Trim(Right(cmbCondicionOtra.Text, 10))) 'arlo20200429
    nTpDocDatDef = IIf(cmbTpDoc.Visible = False Or cmbTpDoc.Text = "", 0, Trim(Right(cmbTpDoc.Text, 10)))
    Set rsDefaut = oDCred.CatalogoProDefaut(Trim(Right(cmbSubProducto.Text, 5)), cParCod, IIf(cmbMoneda.Text = "", 0, Trim(Right(cmbMoneda.Text, 1))), nTpDocDatDef, 1, IIf(bRefinanciar = False, 0, 1), IIf(bAmpliacion = False, 0, 1), nIdCampana)
        
    If Not (rsDefaut.BOF And rsDefaut.EOF) Then
        If cParCod = 4000 Then 'Monto
            If rsDefaut!MinMonto <> -1 And rsDefaut!MaxMonto <> -1 Then
                MinMonto = Format(rsDefaut!MinMonto, "#,#0.00")
                MaxMonto = Format(rsDefaut!MaxMonto, "#,#0.00")
            Else
                MinMonto = 0
                MaxMonto = 9999999
            End If
        ElseIf cParCod = 46000 Then 'Cuota
            If rsDefaut!MinCuota <> -1 And rsDefaut!MaxCuota <> -1 Then
                MinCuota = rsDefaut!MinCuota
                MaxCuota = rsDefaut!MaxCuota
            Else
                MinCuota = 0
                MaxCuota = 999
            End If
        ElseIf cParCod = 7000 Then 'Plazo
            If rsDefaut!MinPlazo <> -1 And rsDefaut!MaxPlazo <> -1 Then
                MinPlazo = rsDefaut!MinPlazo
                MaxPlazo = rsDefaut!MaxPlazo
            Else
                MinPlazo = 0
                MaxPlazo = 999
            End If
        'End If
        'If cmdEjecutar <> -1 Then
        'End If
        End If
        
        'arlo20200429 begin
        If nIdCampana = 136 Then
                MinPlazo = rsDefaut!MinPlazo
                MaxPlazo = rsDefaut!MaxPlazo
                MinCuota = rsDefaut!MinCuota
                MaxCuota = rsDefaut!MaxCuota
                MinMonto = Format(rsDefaut!MinMonto, "#,#0.00")
                MaxMonto = Format(rsDefaut!MaxMonto, "#,#0.00")
        End If
        'arlo20200429 end
    End If
Set oDCred = Nothing
RSClose rsDefaut
End Sub

Private Sub CP_ValidaProdUniCuota()
Dim oDCred As COMDCredito.DCOMCredito
Dim rsTpDoc As ADODB.Recordset
Set oDCred = New COMDCredito.DCOMCredito
    If Trim(Right(cmbSubProducto.Text, 10)) = "520" Then
        If IIf(spnCuotas.valor = "", 0, spnCuotas.valor) = 1 Then
            If bRefinanciar = False Then
                Set rsTpDoc = oDCred.ObtieneTpDoc(Trim(Right(cmbSubProducto.Text, 10)), 22000)
                Call Llenar_Combo_con_Recordset(rsTpDoc, cmbTpDoc)
                Call CambiaTamañoCombo(cmbTpDoc, 220)
                    lblTpDoc.Caption = "Tipo Documentos"
                    nTpCmbTpDoc = 2 '2:Tipo Documentos - Para identificar el tipo de Combo cmbTpDoc
                    cmbTpDoc.ListIndex = -1
                    cmbTpDoc.Visible = True
                    lblTpDoc.Visible = True
            End If
        Else
            nTpCmbTpDoc = 0 '0:Nada - Para identificar el tipo de Combo cmbTpDoc
            cmbTpDoc.ListIndex = -1
            cmbTpDoc.Visible = False
            lblTpDoc.Visible = False
            sT_Plani = ""
        End If
    End If
Set oDCred = Nothing
RSClose rsTpDoc
End Sub

Private Sub CP_CargaTpDoc(cTpProd As String, nParCod As Long)
Dim oDCred As COMDCredito.DCOMCredito
Dim rsTpDoc As ADODB.Recordset
Set oDCred = New COMDCredito.DCOMCredito
Set rsTpDoc = oDCred.ObtieneTpDoc(cTpProd, nParCod)
    If Not (rsTpDoc.BOF And rsTpDoc.EOF) Then
        Select Case cTpProd
            Case 703
                Call Llenar_Combo_con_Recordset(rsTpDoc, cmbTpDoc)
                Call CambiaTamañoCombo(cmbTpDoc, 110)
                lblTpDoc.Caption = "Tipo de Interés"
                nTpCmbTpDoc = 3 '1:Tipo de Interes - Para identificar el tipo de Combo cmbTpDoc
                cmbTpDoc.ListIndex = -1
                cmbTpDoc.Visible = True
                lblTpDoc.Visible = True
            Case 707, 718
                Call Llenar_Combo_con_Recordset(rsTpDoc, cmbTpDoc)
                If cTpProd = 718 Then
                    Call CambiaTamañoCombo(cmbTpDoc, 180)
                Else
                    Call CambiaTamañoCombo(cmbTpDoc, 100)
                End If
                lblTpDoc.Caption = "Tipo de Ingresos"
                nTpCmbTpDoc = 1 '1:Tipo de Ingresos - Para identificar el tipo de Combo cmbTpDoc
                cmbTpDoc.ListIndex = -1
                cmbTpDoc.Visible = True
                lblTpDoc.Visible = True
            Case 525
                Call Llenar_Combo_con_Recordset(rsTpDoc, cmbTpDoc)
                Call CambiaTamañoCombo(cmbTpDoc, 220)
                lblTpDoc.Caption = "Tipo Documentos"
                nTpCmbTpDoc = 2 '2:Tipo Documentos - Para identificar el tipo de Combo cmbTpDoc
                cmbTpDoc.ListIndex = -1
                cmbTpDoc.Visible = True
                lblTpDoc.Visible = True
            Case 520
                If spnCuotas.valor = 1 Then
                    Call Llenar_Combo_con_Recordset(rsTpDoc, cmbTpDoc)
                    Call CambiaTamañoCombo(cmbTpDoc, 220)
                    lblTpDoc.Caption = "Tipo Documentos"
                    nTpCmbTpDoc = 2 '2:Tipo Documentos - Para identificar el tipo de Combo cmbTpDoc
                    cmbTpDoc.ListIndex = -1
                    cmbTpDoc.Visible = True
                    lblTpDoc.Visible = True
                Else
                    nTpCmbTpDoc = 0 '0:Nada - Para identificar el tipo de Combo cmbTpDoc
                    cmbTpDoc.ListIndex = -1
                    cmbTpDoc.Visible = False
                    lblTpDoc.Visible = False
                    sT_Plani = ""
                End If
        End Select
    Else
        cmbTpDoc.ListIndex = -1
        cmbTpDoc.Visible = False
        lblTpDoc.Visible = False
        sT_Plani = ""
        nTpCmbTpDoc = 0 '0:Nada - Para identificar el tipo de Combo cmbTpDoc
    End If
Set oDCred = Nothing
RSClose rsTpDoc
End Sub
Private Sub CP_CargaAporte(ByVal cPersCod As String, ByVal cCodProd As String, Optional ByVal nSubDestino As Long = 0, Optional ByVal nDestino As Long = 0, Optional ByVal nCondicion As Long = 0, Optional ByVal nTpInteres As Long = 0, Optional ByVal nTpMonedad As Long = 0, Optional ByVal nCampana As Integer = -1)
Dim obMontoPre As COMDCredito.DCOMCredito
Dim rsMontoPre As ADODB.Recordset
    Set obMontoPre = New COMDCredito.DCOMCredito
If cmdEjecutar = -1 Then Exit Sub
    Set rsMontoPre = obMontoPre.ObtieneAporteCalifInter(cPersCod, cCodProd, nSubDestino, nDestino, nCondicion, nTpInteres, , nCampana)
        If Not (rsMontoPre.BOF And rsMontoPre.EOF) Then
            If cmdEjecutar <> 2 And cmbSubDestCred.Visible = True Then
                If cmbSubDestCred.Visible = True Then
                    If nSubDestAnt <> Right(cmbSubDestCred.Text, 8) Then
                        Set nMatMontoPre = Nothing
                        nSubDestAnt = Right(cmbSubDestCred.Text, 8)
                    End If
                Else
                    If nSubDestAnt <> Right(cmbDestCred.Text, 8) Then
                        Set nMatMontoPre = Nothing
                        If cmbDestCred.Text <> "" Then
                            nSubDestAnt = Right(cmbDestCred.Text, 8)
                        End If
                    End If
                End If
            End If
                        
            If rsMontoPre!m = 0 Then
                txtMontoSol.Enabled = True
            Else
                If IsArray(nMatMontoPre) Then
                    If UBound(nMatMontoPre) > 0 Then
                        frmCredMontoPresupuestado.Inicio rsMontoPre!m, nMatMontoPre, nDestino, cCodProd, MinMonto, nSubDestino, nTpMonedad, nTpInteres
                            If IsArray(nMatMontoPre) Then
                                If Trim(Right(cmbSubProducto.Text, 5)) = "703" Then
                                    'txtMontoSol.Text = Format(nMatMontoPre(1, 4), "0.00")'Comento JOEP20190313
                                    txtMontoSol.Text = Format(nMatMontoPre(1, 3), "0.00") ' JOEP20190313
                                Else
                                    txtMontoSol.Text = Format(nMatMontoPre(1, 3), "0.00")
                                End If
                            Else
                                txtMontoSol.Text = "0.00"
                            End If
                    End If
                Else
                    Set nMatMontoPre = Nothing
                    frmCredMontoPresupuestado.Inicio rsMontoPre!m, nMatMontoPre, nDestino, cCodProd, MinMonto, nSubDestino, nTpMonedad, nTpInteres
                    If UBound(nMatMontoPre) > 0 Then
                        If Trim(Right(cmbSubProducto.Text, 5)) = "703" Then
                            'txtMontoSol.Text = Format(nMatMontoPre(1, 4), "0.00")'Comento JOEP20190313
                            txtMontoSol.Text = Format(nMatMontoPre(1, 3), "0.00") ' JOEP20190313
                        Else
                            txtMontoSol.Text = Format(nMatMontoPre(1, 3), "0.00")
                        End If
                    Else
                        txtMontoSol.Text = "0.00"
                    End If
                End If
                txtMontoSol.Enabled = False
            End If
        Else
            Set nMatMontoPre = Nothing
            If cmbSubDestCred.Visible = True Then
                If Not (Trim(Right(cmbSubDestCred.Text, 10)) = 35001 Or Trim(Right(cmbSubDestCred.Text, 10)) = 35002) Then
                    txtMontoSol.Text = "0.00"
                End If
            End If
            txtMontoSol.Enabled = True
        End If
Set obMontoPre = Nothing
RSClose rsMontoPre
End Sub
Private Function CP_Mensajes(ByVal nTpMsn As Integer, ByVal cTpProd As String, Optional ByVal bRefi As Boolean = False, Optional ByVal bAmpli As Boolean = False) As Boolean
CP_Mensajes = True
If nTpMsn = 1 And cTpProd = "718" Then
    If cmbTpDoc.Visible = True And cmbTpDoc.Text = "" And cmbTpDoc.Enabled = True Then
        MsgBox "Seleccione el Tipo de Ingresos", vbInformation, "Aviso"
        cmbDestCred.ListIndex = -1
        cmbTpDoc.SetFocus
        CP_Mensajes = False
    End If
End If
If nTpMsn = 2 Then
    If cmbMoneda.Text = "" And cmbMoneda.Enabled = True Then
        MsgBox "Seleccione el tipo de moneda", vbInformation, "Aviso"
        cmbMoneda.SetFocus
        CP_Mensajes = False
    End If
End If
If nTpMsn = 3 Then
    If (cmbTpDoc.Visible = True And cmbTpDoc.Text = "" And cmbTpDoc.Enabled = True) Then
        If Trim(Right(cmbSubProducto.Text, 5)) = 525 Then
            MsgBox "Selecione el tipo de Documentos", vbInformation, "Aviso"
        Else
            MsgBox "Selecione el tipo de Ingreso", vbInformation, "Aviso"
        End If
        cmbDestCred.ListIndex = -1
        If cmbTpDoc.Enabled = True Then cmbTpDoc.SetFocus
        CP_Mensajes = False
    End If
End If
If nTpMsn = 4 Then
    If (cmbCondicion.Text = "") Then
        MsgBox "Selecione el tipo de Condición", vbInformation, "Aviso"
        CP_Mensajes = False
    End If
End If
If nTpMsn = 5 Then
    If cmbTpDoc.Visible = True Then
        If cTpProd = 520 And Trim(Right(cmbDestCred.Text, 5)) <> 1 And (Trim(Right(cmbTpDoc.Text, 10)) = 22001 Or Trim(Right(cmbTpDoc.Text, 10)) = 22002 Or Trim(Right(cmbTpDoc.Text, 10)) = 22003 Or Trim(Right(cmbTpDoc.Text, 10)) = 22004) Then
            MsgBox "El destino seleccionado no coresponde al Tipo de Documento", vbInformation, "Aviso"
            CP_Mensajes = False
        End If
    End If
End If
If nTpMsn = 6 Then
    If cTpProd = 521 And cmbDestCred.Text = "" And bRefinanciar = bRefi And bAmpliacion = bAmpli Then
        MsgBox "Seleccione el destino.", vbInformation, "Aviso"
        cmbDestCred.SetFocus
        CP_Mensajes = False
    End If
End If
If nTpMsn = 7 Then
    If cmbSubProducto.Text = "" Then
        MsgBox "Seleccione el Tipo de Producto.", vbInformation, "Aviso"
        If cmbSubProducto.Enabled = True Then
            cmbSubProducto.SetFocus
        End If
        CP_Mensajes = False
    End If
End If
End Function
'Agrego JOEP20190919 ERS042 CP-2018
'MARG201906 ERS011-2019*************
Public Function ValidarScore(ByVal pcCtaCod As String, Optional ByVal pMatRelaciones As Variant, Optional ByVal pnFormulario As Integer) As Boolean
    Dim oScore As COMDCredito.DCOMCredito
    Dim isConsultaExperian As Boolean
    Dim isAprobacionSinScore As Boolean
    Dim bPermiteSugerencia As Boolean
    
    Set oScore = New COMDCredito.DCOMCredito
    isConsultaExperian = oScore.isConsultaExperian
    If isConsultaExperian Then
        bPermiteSugerencia = ValidarScoreExperian(pcCtaCod, pMatRelaciones)
    Else
        'validacion de aprobacion sin score
        isAprobacionSinScore = oScore.isAprobacionSinScore
        If isAprobacionSinScore Then
            bPermiteSugerencia = True
            'insertar credito para posterior consulta a experian
            oScore.InsertarCreditoSinScore pcCtaCod, gsCodUser
        Else
            Dim pbExitoScoreExperian As Boolean
            bPermiteSugerencia = ValidarScoreExperian(pcCtaCod, pMatRelaciones, pbExitoScoreExperian, False)
            If Not pbExitoScoreExperian Then
                bPermiteSugerencia = ValidarScoreManual(pcCtaCod, pMatRelaciones, pnFormulario)
            End If
        End If
    End If
    ValidarScore = bPermiteSugerencia
End Function
Private Function ValidarScoreManual(ByVal psNuevaCta As String, Optional ByVal MatCredRelaciones As Variant, Optional ByVal pnFormulario As Integer) As Boolean
    Dim oScore As COMDCredito.DCOMCredito
    Dim esAplicableValidacionScore As Boolean
    Dim bPermiteSugerencia As Boolean
    
    Set oScore = New COMDCredito.DCOMCredito
    bPermiteSugerencia = False
    esAplicableValidacionScore = oScore.esAplicableValidacionScoreManual(psNuevaCta)
    
    If esAplicableValidacionScore Then
        Dim mensajeScore As String
        Dim rsScore As ADODB.Recordset
        Set rsScore = oScore.getDecisionScoreManual(psNuevaCta, pnFormulario) '1:solicitud, 2: sugerencia
        If Not rsScore.BOF And Not rsScore.EOF Then
            mensajeScore = rsScore!mensaje
            bPermiteSugerencia = CBool(rsScore!bPermiteSugerencia)
            MsgBox mensajeScore, vbInformation, "AVISO"
        End If
        rsScore.Close
        Set rsScore = Nothing
    Else
        bPermiteSugerencia = True
    End If
    
    ValidarScoreManual = bPermiteSugerencia
End Function
'END MARG***************************
