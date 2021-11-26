VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredFormEvalFormato6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Créditos - Evaluación - Formato 6"
   ClientHeight    =   11175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16515
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredFormEvalFormato6.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11175
   ScaleWidth      =   16515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Ratios - Indicadores:"
      ForeColor       =   &H8000000D&
      Height          =   2895
      Left            =   14700
      TabIndex        =   122
      ToolTipText     =   "Datos del flujo de Caja Proyectado"
      Top             =   8200
      Width           =   1755
      Begin SICMACT.EditMoney txtCapacidadNeta2 
         Height          =   300
         Left            =   90
         TabIndex        =   123
         Top             =   435
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtIngresoNeto2 
         Height          =   300
         Left            =   120
         TabIndex        =   124
         Top             =   930
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtExcedenteMensual2 
         Height          =   300
         Left            =   120
         TabIndex        =   125
         Top             =   1425
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtRentabilidad 
         Height          =   300
         Left            =   120
         TabIndex        =   126
         Top             =   1935
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin SICMACT.EditMoney txtLiquidezCte 
         Height          =   300
         Left            =   120
         TabIndex        =   127
         Top             =   2415
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Liquidez Cte.:"
         Height          =   195
         Left            =   120
         TabIndex        =   133
         Top             =   2235
         Width           =   990
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rentabilidad:"
         Height          =   195
         Left            =   120
         TabIndex        =   132
         Top             =   1755
         Width           =   945
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excedente:"
         Height          =   195
         Left            =   120
         TabIndex        =   131
         Top             =   1230
         Width           =   825
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso Neto:"
         Height          =   195
         Left            =   120
         TabIndex        =   130
         Top             =   735
         Width           =   1005
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacidad de Pago:"
         Height          =   195
         Left            =   120
         TabIndex        =   129
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   0
         Left            =   1440
         TabIndex        =   128
         Top             =   480
         Width           =   165
      End
   End
   Begin VB.Frame FrameDatos 
      Caption         =   "Datos Flujo Proyect."
      ForeColor       =   &H8000000D&
      Height          =   3255
      Left            =   14700
      TabIndex        =   108
      Top             =   4920
      Width           =   1755
      Begin SICMACT.EditMoney txtIncrVentasContado 
         Height          =   300
         Left            =   120
         TabIndex        =   109
         ToolTipText     =   "Incremento de ventas al contado - Anual"
         Top             =   450
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtIncrCompraMercaderia 
         Height          =   300
         Left            =   120
         TabIndex        =   110
         ToolTipText     =   "Incremento de Compras de Mercaderias - Anual"
         Top             =   1050
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtIncrPagoPersonal 
         Height          =   300
         Left            =   120
         TabIndex        =   111
         ToolTipText     =   "Incremento de Consumo - Anual"
         Top             =   1650
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtIncrGastoVentas 
         Height          =   300
         Left            =   120
         TabIndex        =   112
         ToolTipText     =   "Incremento de Pago Personal -Anual"
         Top             =   2250
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtIncrConsumo 
         Height          =   300
         Left            =   120
         TabIndex        =   113
         ToolTipText     =   "Incremento de Gastos de Ventas - Anual"
         Top             =   2850
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
         Text            =   "0"
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Anual"
         Height          =   195
         Index           =   11
         Left            =   960
         TabIndex        =   137
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Anual"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   136
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. ventas contado:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   121
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. Compra Mercad:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   120
         Top             =   810
         Width           =   1575
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. de Consumo:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   119
         Top             =   2610
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. Pago Personal:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   118
         Top             =   1410
         Width           =   1470
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. Gasto Ventas:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   117
         Top             =   2010
         Width           =   1410
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Anual"
         Height          =   195
         Index           =   8
         Left            =   960
         TabIndex        =   116
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Anual"
         Height          =   195
         Index           =   9
         Left            =   960
         TabIndex        =   115
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Anual"
         Height          =   195
         Index           =   10
         Left            =   960
         TabIndex        =   114
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Frame fraBotones 
      Height          =   3500
      Left            =   14740
      TabIndex        =   81
      Top             =   1400
      Width           =   1725
      Begin VB.CommandButton cmdMNME 
         Caption         =   "MN - ME"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   75
         TabIndex        =   142
         Top             =   1080
         Visible         =   0   'False
         Width           =   1600
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   80
         TabIndex        =   94
         Top             =   555
         Width           =   1600
      End
      Begin VB.CommandButton cmdVerCar2 
         Caption         =   "Ver CAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   75
         TabIndex        =   87
         Top             =   1800
         Width           =   1600
      End
      Begin VB.CommandButton cmdInformeVista2 
         Caption         =   "Informe Visita"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   75
         TabIndex        =   86
         Top             =   1440
         Width           =   1600
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   80
         TabIndex        =   85
         Top             =   180
         Width           =   1600
      End
      Begin VB.CommandButton cmdFlujoCaja 
         Caption         =   "Ver Flujo Caja Mensual Proyec. Histórico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   75
         TabIndex        =   84
         Top             =   2860
         Width           =   1600
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Hoja Evaluación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   75
         TabIndex        =   83
         Top             =   2160
         Width           =   1600
      End
      Begin VB.CommandButton cmdImpEEFF 
         Caption         =   "EE.FF."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   75
         TabIndex        =   82
         Top             =   2520
         Width           =   1600
      End
      Begin VB.Line Line1 
         X1              =   180
         X2              =   1600
         Y1              =   1005
         Y2              =   1005
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Negocio"
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   15
      TabIndex        =   68
      Top             =   0
      Width           =   16320
      Begin VB.Frame Frame17 
         Height          =   760
         Left            =   100
         TabIndex        =   69
         Top             =   510
         Width           =   15975
         Begin VB.TextBox txtNombreCliente2 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1400
            TabIndex        =   3
            Top             =   120
            Width           =   5895
         End
         Begin VB.OptionButton OptCondLocal2 
            Caption         =   "Propia"
            Height          =   240
            Index           =   1
            Left            =   5140
            TabIndex        =   8
            Top             =   450
            Width           =   855
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Alquilada"
            Height          =   240
            Index           =   2
            Left            =   5970
            TabIndex        =   9
            Top             =   450
            Width           =   975
         End
         Begin VB.OptionButton OptCondLocal2 
            Caption         =   "Ambulante"
            Height          =   240
            Index           =   3
            Left            =   7020
            TabIndex        =   10
            Top             =   450
            Width           =   1095
         End
         Begin VB.OptionButton OptCondLocal2 
            Caption         =   "Otros"
            Height          =   240
            Index           =   4
            Left            =   8180
            TabIndex        =   11
            Top             =   450
            Width           =   810
         End
         Begin VB.TextBox txtCondLocalOtros 
            Height          =   285
            Left            =   8985
            TabIndex        =   12
            Top             =   440
            Visible         =   0   'False
            Width           =   2715
         End
         Begin MSMask.MaskEdBox txtFecUltEndeuda 
            Height          =   300
            Left            =   14520
            TabIndex        =   14
            Top             =   435
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin Spinner.uSpinner spnTiempoLocalAnio 
            Height          =   280
            Left            =   1400
            TabIndex        =   6
            Top             =   420
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            Max             =   99
            MaxLength       =   2
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
         Begin Spinner.uSpinner spnTiempoLocalMes 
            Height          =   280
            Left            =   2580
            TabIndex        =   7
            Top             =   420
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            Max             =   12
            MaxLength       =   2
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
         Begin Spinner.uSpinner spnExpEmpAnio 
            Height          =   280
            Left            =   9255
            TabIndex        =   4
            Top             =   120
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            Max             =   99
            MaxLength       =   2
            Enabled         =   0   'False
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
         Begin Spinner.uSpinner spnExpEmpMes 
            Height          =   280
            Left            =   10440
            TabIndex        =   5
            Top             =   120
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            Max             =   12
            MaxLength       =   2
            Enabled         =   0   'False
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
         Begin SICMACT.EditMoney txtUltEndeuda 
            Height          =   300
            Left            =   14520
            TabIndex        =   13
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   840
            TabIndex        =   79
            Top             =   160
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exp. como Empresario:"
            Height          =   200
            Left            =   7635
            TabIndex        =   78
            Top             =   160
            Width           =   1650
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tiempo en el local:"
            Height          =   190
            Left            =   60
            TabIndex        =   77
            Top             =   450
            Width           =   1320
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Condición local:"
            Height          =   195
            Left            =   4080
            TabIndex        =   76
            Top             =   450
            Width           =   1110
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "años"
            Height          =   255
            Left            =   10040
            TabIndex        =   75
            Top             =   160
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "años"
            Height          =   255
            Left            =   2160
            TabIndex        =   74
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "meses"
            Height          =   255
            Left            =   11200
            TabIndex        =   73
            Top             =   160
            Width           =   615
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "meses"
            Height          =   200
            Left            =   3350
            TabIndex        =   72
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Último endeudamiento RCC:"
            Height          =   195
            Left            =   12480
            TabIndex        =   71
            Top             =   165
            Width           =   2010
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fec. último endeud. SBS RCC:"
            Height          =   195
            Left            =   12300
            TabIndex        =   70
            Top             =   465
            Width           =   2160
         End
      End
      Begin VB.TextBox txtGiroNeg2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4800
         TabIndex        =   2
         Top             =   180
         Width           =   5300
      End
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   495
         Left            =   105
         TabIndex        =   1
         Top             =   165
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
         Texto           =   "Crédito"
      End
      Begin MSMask.MaskEdBox txtFechaEvaluacion 
         Height          =   300
         Left            =   11355
         TabIndex        =   88
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin SICMACT.EditMoney txtExposicionCredito2 
         Height          =   300
         Left            =   14760
         TabIndex        =   92
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp. con este crédito:"
         Height          =   195
         Left            =   13200
         TabIndex        =   93
         Top             =   225
         Width           =   1590
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Eval.:"
         Height          =   195
         Left            =   10245
         TabIndex        =   89
         Top             =   225
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Giro del Neg.:"
         Height          =   195
         Left            =   3840
         TabIndex        =   80
         Top             =   225
         Width           =   990
      End
   End
   Begin TabDlg.SSTab SSTabIngresos2 
      Height          =   9885
      Left            =   0
      TabIndex        =   0
      Top             =   1340
      Width           =   14680
      _ExtentX        =   25903
      _ExtentY        =   17436
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Balance General"
      TabPicture(0)   =   "frmCredFormEvalFormato6.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblBalance"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblSolesBalance"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fePasivos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "feActivos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frameAgregar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Estado de Resultados"
      TabPicture(1)   =   "frmCredFormEvalFormato6.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(1)=   "Label33"
      Tab(1).Control(2)=   "Label34"
      Tab(1).Control(3)=   "Frame6"
      Tab(1).Control(4)=   "Command3"
      Tab(1).Control(5)=   "txtCantEGP"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Coeficiente Financiero"
      TabPicture(2)   =   "frmCredFormEvalFormato6.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame12"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Evaluación"
      TabPicture(3)   =   "frmCredFormEvalFormato6.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame16"
      Tab(3).Control(1)=   "Frame15"
      Tab(3).Control(2)=   "Frame14"
      Tab(3).Control(3)=   "frameFlujoEval"
      Tab(3).Control(4)=   "Line3"
      Tab(3).Control(5)=   "Line2"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Propuesta del Crédito"
      TabPicture(4)   =   "frmCredFormEvalFormato6.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).Control(1)=   "Frame9"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Comentarios y Referidos"
      TabPicture(5)   =   "frmCredFormEvalFormato6.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdAgregar2"
      Tab(5).Control(1)=   "cmdQuitar2"
      Tab(5).Control(2)=   "Frame11"
      Tab(5).Control(3)=   "Frame10"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Flujo Caja Histórico"
      TabPicture(6)   =   "frmCredFormEvalFormato6.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "frameFlujoHisto"
      Tab(6).ControlCount=   1
      Begin VB.Frame frameFlujoHisto 
         Caption         =   "Flujo de Caja Histórico [Expresado en soles]:"
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
         Height          =   8835
         Left            =   -74520
         TabIndex        =   134
         Top             =   720
         Width           =   6480
         Begin VB.TextBox txtFinanciamiento 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   4560
            TabIndex        =   141
            Text            =   "0"
            Top             =   7920
            Width           =   1695
         End
         Begin VB.TextBox txtInversion 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   4560
            TabIndex        =   139
            Text            =   "0"
            Top             =   8235
            Width           =   1695
         End
         Begin SICMACT.FlexEdit feFlujoCajaHistorico 
            Height          =   7215
            Left            =   120
            TabIndex        =   135
            Top             =   360
            Width           =   6240
            _ExtentX        =   11007
            _ExtentY        =   12515
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Concepto-Monto-nConsCod-nConsValor"
            EncabezadosAnchos=   "300-4100-1700-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-C-C"
            FormatosEdit    =   "0-0-2-0-0"
            CantEntero      =   12
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
         End
         Begin VB.Label Label35 
            Caption         =   "FINANCIAMIENTO :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   140
            Top             =   7920
            Width           =   1575
         End
         Begin VB.Label Label32 
            Caption         =   "INVERSION :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3540
            TabIndex        =   138
            Top             =   8280
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   -74760
         TabIndex        =   104
         Top             =   360
         Width           =   3495
         Begin MSMask.MaskEdBox txtFechaVisita 
            Height          =   300
            Left            =   1680
            TabIndex        =   105
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Visita:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   240
            TabIndex        =   106
            Top             =   300
            Width           =   1305
         End
      End
      Begin VB.Frame frameAgregar 
         Height          =   520
         Left            =   165
         TabIndex        =   95
         Top             =   320
         Width           =   14415
         Begin VB.CheckBox chkAudit 
            Caption         =   "Auditado"
            Height          =   195
            Left            =   2760
            TabIndex        =   107
            Top             =   180
            Width           =   975
         End
         Begin VB.CommandButton cmdAgregaEEFF 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   12480
            TabIndex        =   97
            Top             =   120
            Width           =   1800
         End
         Begin VB.ComboBox CboFecRegEEFF 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   96
            Top             =   120
            Width           =   1170
         End
         Begin MSMask.MaskEdBox mskFecReg 
            Height          =   300
            Left            =   120
            TabIndex        =   98
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label31 
            Caption         =   "Seleccione Fecha:"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   180
            Width           =   1335
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Declaración PDT:"
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
         Height          =   1420
         Left            =   -74760
         TabIndex        =   67
         Top             =   8325
         Width           =   12075
         Begin SICMACT.FlexEdit feDeclaracionPDT 
            Height          =   1095
            Left            =   45
            TabIndex        =   22
            Top             =   240
            Width           =   11115
            _ExtentX        =   19606
            _ExtentY        =   1931
            Rows            =   3
            Cols0           =   9
            FixedCols       =   2
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Mes/Detalle-nConsCod-nConsValor----Promedio-%Vent. Decl."
            EncabezadosAnchos=   "0-3000-0-0-1600-1600-1600-1500-1500"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-4-5-6-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-L-R-C-C-C-C-R"
            FormatosEdit    =   "0-0-0-0-4-4-4-0-2"
            CantEntero      =   12
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
            CellBackColor   =   -2147483633
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Otros Ingresos :"
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
         Height          =   2415
         Left            =   -68400
         TabIndex        =   66
         Top             =   4560
         Width           =   6375
         Begin SICMACT.FlexEdit feOtrosIngresos 
            Height          =   1815
            Left            =   75
            TabIndex        =   21
            Top             =   320
            Width           =   6240
            _ExtentX        =   11007
            _ExtentY        =   3201
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-N-Concepto-Monto-Aux"
            EncabezadosAnchos=   "0-300-4000-1800-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-3-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-R-C"
            FormatosEdit    =   "0-0-0-2-0"
            CantEntero      =   12
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Gastos Familiares : "
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
         Height          =   3015
         Left            =   -68355
         TabIndex        =   65
         Top             =   1440
         Width           =   6375
         Begin SICMACT.FlexEdit feGastosFamiliares 
            Height          =   2415
            Left            =   75
            TabIndex        =   20
            Top             =   320
            Width           =   6240
            _ExtentX        =   11218
            _ExtentY        =   4260
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-N-Concepto-Monto-Aux"
            EncabezadosAnchos=   "0-300-4000-1800-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-3-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-R-C"
            FormatosEdit    =   "0-0-0-2-0"
            CantEntero      =   12
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
      Begin VB.Frame frameFlujoEval 
         Caption         =   "Flujo de Caja Mensual :"
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
         Height          =   7515
         Left            =   -74920
         TabIndex        =   64
         Top             =   720
         Width           =   6480
         Begin SICMACT.FlexEdit feFlujoCajaMensual 
            Height          =   7215
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   6240
            _ExtentX        =   11007
            _ExtentY        =   12515
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Concepto-Monto-nConsCod-nConsValor"
            EncabezadosAnchos=   "300-4100-1700-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-C-C"
            FormatosEdit    =   "0-0-2-0-0"
            CantEntero      =   12
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Propuesta del Credito:"
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
         Height          =   8055
         Left            =   -74760
         TabIndex        =   57
         Top             =   1320
         Width           =   12135
         Begin VB.TextBox txtEntornoFamiliar2 
            Height          =   900
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   600
            Width           =   11655
         End
         Begin VB.TextBox txtGiroUbicacion2 
            Height          =   900
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   1800
            Width           =   11655
         End
         Begin VB.TextBox txtExperiencia2 
            Height          =   900
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   3120
            Width           =   11655
         End
         Begin VB.TextBox txtFormalidadNegocio2 
            Height          =   900
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   4320
            Width           =   11655
         End
         Begin VB.TextBox txtColaterales2 
            Height          =   900
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   5520
            Width           =   11655
         End
         Begin VB.TextBox txtDestino2 
            Height          =   900
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   6840
            Width           =   11655
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el entorno familiar del cliente o representante:"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   360
            Width           =   3795
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el giro y la ubicación del negocio:"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   1560
            Width           =   2820
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sólo la experiencia crediticia:"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   2880
            Width           =   2070
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la consistencia de la información y la formalidad del negocio:"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   4080
            Width           =   4770
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   5280
            Width           =   2400
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el destino y el impacto del mismo:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   58
            Top             =   6600
            Width           =   2850
         End
      End
      Begin VB.CommandButton cmdAgregar2 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74280
         TabIndex        =   31
         Top             =   6240
         Width           =   1170
      End
      Begin VB.CommandButton cmdQuitar2 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -72840
         TabIndex        =   32
         Top             =   6240
         Width           =   1170
      End
      Begin VB.Frame Frame11 
         Caption         =   "Referidos :"
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
         Height          =   3135
         Left            =   -74880
         TabIndex        =   56
         Top             =   3000
         Width           =   13455
         Begin SICMACT.FlexEdit feReferidos 
            Height          =   2655
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   13155
            _ExtentX        =   23204
            _ExtentY        =   4683
            Cols0           =   7
            HighLight       =   1
            EncabezadosNombres=   "N°-Nombre-DNI-Telef.-Referido-DNI-Aux"
            EncabezadosAnchos=   "500-5000-1250-1250-5000-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2-3-4-5-X"
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-L-L-L-C-L-L"
            FormatosEdit    =   "0-0-0-0-0-0-0"
            TextArray0      =   "N°"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   495
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Comentarios :"
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
         Height          =   2415
         Left            =   -74880
         TabIndex        =   55
         Top             =   480
         Width           =   13455
         Begin VB.TextBox txtComentario 
            Height          =   2010
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   29
            Top             =   240
            Width           =   13095
         End
      End
      Begin VB.TextBox txtCantEGP 
         Height          =   405
         Left            =   -65760
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   1020
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Mostrar"
         Height          =   375
         Left            =   -65040
         TabIndex        =   17
         Top             =   1020
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame12 
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
         Height          =   7935
         Left            =   -74880
         TabIndex        =   53
         Top             =   720
         Width           =   7455
         Begin SICMACT.FlexEdit feCoeFinan 
            Height          =   7575
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   6960
            _ExtentX        =   12277
            _ExtentY        =   13361
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Concepto-Monto-nConsCod-nConsValor"
            EncabezadosAnchos=   "300-4500-1700-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-L-C"
            FormatosEdit    =   "0-0-2-0-0"
            CantEntero      =   12
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame6 
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
         Height          =   7950
         Left            =   -74640
         TabIndex        =   52
         Top             =   960
         Width           =   7335
         Begin SICMACT.FlexEdit feEstaGananPerd 
            Height          =   7635
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   6840
            _ExtentX        =   12065
            _ExtentY        =   13467
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Concepto-Monto-nConsCod-nConsValor"
            EncabezadosAnchos=   "300-4500-1700-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-C-C"
            FormatosEdit    =   "0-0-2-0-0"
            CantEntero      =   12
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
         End
      End
      Begin VB.Frame frmCredEvalFormato1 
         Caption         =   " Gastos del Negocio :"
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
         Height          =   6015
         Left            =   -74880
         TabIndex        =   39
         Top             =   360
         Width           =   9975
         Begin VB.TextBox txtEntornoFamiliar 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   45
            Top             =   480
            Width           =   9735
         End
         Begin VB.TextBox txtGiroUbicacion 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   44
            Top             =   1440
            Width           =   9735
         End
         Begin VB.TextBox txtExperiencia 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   43
            Top             =   2400
            Width           =   9735
         End
         Begin VB.TextBox txtFormalidadNegocio 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   42
            Top             =   3360
            Width           =   9735
         End
         Begin VB.TextBox txtColaterales 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   41
            Top             =   4320
            Width           =   9735
         End
         Begin VB.TextBox txtDestino 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   40
            Top             =   5280
            Width           =   9735
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el entorno familiar del cliente o representante:"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   3795
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el giro y la ubicación del negocio:"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   1200
            Width           =   2820
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sólo la experiencia crediticia:"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   2160
            Width           =   2070
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la consistencia de la información y la formalidad del negocio:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   3120
            Width           =   4770
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   4080
            Width           =   2400
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   5040
            Width           =   2400
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Comentarios :"
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
         Height          =   2655
         Left            =   -74760
         TabIndex        =   37
         Top             =   360
         Width           =   9975
         Begin VB.TextBox Text1 
            Height          =   2010
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   9735
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Comentarios :"
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
         Height          =   2655
         Left            =   -74760
         TabIndex        =   35
         Top             =   3360
         Width           =   9975
         Begin SICMACT.FlexEdit FlexEdit1 
            Height          =   1935
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   9720
            _ExtentX        =   17145
            _ExtentY        =   3413
            Cols0           =   6
            HighLight       =   1
            EncabezadosNombres=   "N°-Nombre-DNI-Telef.-Referido-DNI"
            EncabezadosAnchos=   "1000-2800-1000-1500-2300-1000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-R-L-C-C-C"
            FormatosEdit    =   "0-2-0-0-0-0"
            TextArray0      =   "N°"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   1005
            RowHeight0      =   300
         End
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74640
         TabIndex        =   34
         Top             =   6120
         Width           =   1170
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73200
         TabIndex        =   33
         Top             =   6120
         Width           =   1170
      End
      Begin SICMACT.FlexEdit feActivos 
         Height          =   8655
         Left            =   60
         TabIndex        =   90
         Top             =   1185
         Width           =   7400
         _ExtentX        =   13044
         _ExtentY        =   15266
         Cols0           =   7
         ScrollBars      =   1
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-ACTIVOS-P.P.-P.E.-Total-nConsCod-nConsValor"
         EncabezadosAnchos=   "0-3200-1300-1300-1500-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-4-X-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-R-C-C"
         FormatosEdit    =   "0-0-2-0-0-0-0"
         CantEntero      =   12
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
      End
      Begin SICMACT.FlexEdit fePasivos 
         Height          =   8655
         Left            =   7460
         TabIndex        =   91
         Top             =   1200
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   15266
         Cols0           =   7
         ScrollBars      =   1
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-PASIVOS-P.P.-P.E.-Total-nConsCod-nConsValor"
         EncabezadosAnchos=   "0-3200-1300-1300-1300-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-R-C-C"
         FormatosEdit    =   "0-0-2-0-0-0-0"
         CantEntero      =   12
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
      End
      Begin VB.Label Label34 
         Caption         =   "ESTADO DE RESULTADOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -72360
         TabIndex        =   103
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label33 
         Caption         =   "( Expresado en Soles )"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -72120
         TabIndex        =   102
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblSolesBalance 
         Caption         =   "(Expresado en soles)"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   7440
         TabIndex        =   101
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblBalance 
         Caption         =   "BALANCE GENERAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   5760
         TabIndex        =   100
         Top             =   840
         Width           =   1575
      End
      Begin VB.Line Line3 
         X1              =   -68400
         X2              =   -62040
         Y1              =   8160
         Y2              =   8160
      End
      Begin VB.Line Line2 
         X1              =   -75000
         X2              =   -75000
         Y1              =   0
         Y2              =   4800
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad de G y P:"
         Height          =   195
         Left            =   -67200
         TabIndex        =   54
         Top             =   1140
         Visible         =   0   'False
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmCredFormEvalFormato6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmCredFormEvalFormato6
'** Descripción : Formulario para evaluación de Creditos que tienen el tipo de evaluación 6
'** Referencia  : ERS004-2016
'** Creación    : PEAC, 20160610 09:00:00 AM
'**********************************************************************************************
Option Explicit

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

Dim fnTipoCliente As Integer
Dim fsCtaCod As String
Dim gsOpeCod As String
Dim fnTipoRegMant As Integer
Dim fnTipoPermiso As Integer
Dim fbPermiteGrabar As Boolean
Dim fbBloqueaTodo As Boolean
Dim lnIndMaximaCapPago As Double
Dim lnIndCuotaUNM As Double
Dim lnIndCuotaExcFam As Double
Dim lnCondLocal As Integer
Dim rsCredEval As ADODB.Recordset
Dim rsInd As ADODB.Recordset
Dim fnTotalRef As Currency

Dim rsFeGastoNeg As ADODB.Recordset
Dim rsFeDatGastoFam As ADODB.Recordset
Dim rsFeDatOtrosIng As ADODB.Recordset
Dim rsFeDatRef As ADODB.Recordset
Dim rsFeDatActivos As ADODB.Recordset
Dim rsFeDatPasivos As ADODB.Recordset
Dim rsFeDatPasivosNo As ADODB.Recordset
Dim rsFeDatPDT As ADODB.Recordset
Dim rsFeDatPDTDet As ADODB.Recordset
Dim rsFeDatPatrimonio As ADODB.Recordset
Dim rsFeDatPasPat As ADODB.Recordset
Dim rsFeDatRatios As ADODB.Recordset
Dim rsFeDatIngNeg As ADODB.Recordset
Dim rsFeDatActivosForm6 As ADODB.Recordset
Dim rsFeDatPasivosForm6 As ADODB.Recordset
Dim rsFeDatEstadoGanPerdForm6 As ADODB.Recordset
Dim rsFeDatCoeficienteFinanForm6 As ADODB.Recordset
Dim rsDatRatios As ADODB.Recordset

Dim rsInfVisita As ADODB.Recordset
Dim rsDatActivos As ADODB.Recordset
Dim rsDatPasivos As ADODB.Recordset
Dim rsDatActivosForm6Det As ADODB.Recordset
Dim rsDatPasivosForm6det As ADODB.Recordset
Dim rsDatEstadoGananPerdForm6 As ADODB.Recordset
Dim rsDatCoeFinanForm6 As ADODB.Recordset
Dim rsDatFlujoCaja As ADODB.Recordset
Dim rsDatIfiflujocaja As ADODB.Recordset
Dim rsDatFlujoCajaHistorico As ADODB.Recordset 'LUCV20171015, Agregó según ERS0512017
Dim rsDatIfiflujocajaHistorico As ADODB.Recordset 'LUCV20171015, Agregó según ERS0512017
Dim rsDatGastoFam As ADODB.Recordset
Dim rsDatOtrosIng As ADODB.Recordset
Dim rsDatPDT As ADODB.Recordset
Dim rsDatPDTDet As ADODB.Recordset
Dim rsDatIfiGastoFami As ADODB.Recordset
Dim rsDatRatiosIndi As ADODB.Recordset

Dim rsDatEstadoGP As ADODB.Recordset
Dim rsDatIngNeg As ADODB.Recordset
Dim rsFeFlujoCaja As ADODB.Recordset
Dim rsFeFlujoCajaHistorico As ADODB.Recordset 'LUCV20171015, Agregó según ERS0512017
Dim rsDatParamFlujoCaja As ADODB.Recordset 'LUCV20171015, Agregó según ERS0512017
Dim rsDLineaCNU As ADODB.Recordset

Dim rsDatGastoNeg As ADODB.Recordset

Dim cuotaifi As Integer
Dim fnTotalRefGastoNego As Currency
Dim fnTotalRefGastoFami As Currency
Dim fnTotalRefFlujoCaja As Currency

Dim fsCliente As String
Dim fsGiroNego As String
Dim fsAnioExp As Integer
Dim fsMesExp As Integer
Dim fsUserAnalista  As String

Dim fnMontoDeudaSbs As Currency

Dim nTasaIngNeg As Double
Dim nTasaGastoNeg As Double
Dim nTasaGastoFam As Double
Dim nTasaOtrosIng As Double

Dim fnPasivoPE As Double
Dim fnPasivoPP As Double
Dim fnPasivoTOTAL As Double

Dim fnActivoPE As Double
Dim fnActivoPP As Double
Dim fnActivoTOTAL As Double

Dim cSPrd As String, cPrd As String
Dim objPista As COMManejador.Pista
Dim nFormato, nPersoneria As Integer
Dim fnMontoIni As Double
Dim lnMin As Double, lnMax As Double
Dim lnMinDol As Double, lnMaxDol As Double
Dim nTC As Double
Dim i, j As Integer
Dim sMes1 As String, sMes2 As String, sMes3 As String
Dim nMes1 As Integer, nMes2 As Integer, nMes3 As Integer
Dim nAnio1 As Integer, nAnio2 As Integer, nAnio3 As Integer
Dim nMontoPDT, nMontoAct, nMontoPas, nMontoPasN As Double

Dim oFrm6 As frmCredFormEvalDetalleFormato6

Dim fbGrabar As Boolean
Dim lnColocCondi As Integer
Dim fbTieneReferido6Meses As Boolean 'LUCV20171115, Agregó según correo: RUSI
Dim lnNumForm As Integer

Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval

Dim prsDatActivosForm6 As ADODB.Recordset
Dim prsDatPasivosForm6 As ADODB.Recordset
Dim prsDatEstadoGananPerdForm6 As ADODB.Recordset
Dim prsDatActivosForm6Det As ADODB.Recordset
Dim prsDatPasivosForm6det As ADODB.Recordset
Dim prsDatCoeFinanForm6 As ADODB.Recordset
Dim rsDatPropuesta As ADODB.Recordset
Dim rsDatRef As ADODB.Recordset

Dim lcValorIngre As String
Dim lnPrdEstado As Integer
Dim fnTotalPasivo9 As Currency
Dim lcCodifi As String
Dim lcDescripcionDet As String
Dim lcComentario As String
Dim nTotal As Double

'LUCV20170915 *****-> Comentó y agregó, según ERS051-2017
'Dim lvPrincipalActivos() As tForEvalResumenEstFinFormato6
'Dim lvPrincipalPasivos() As tForEvalResumenEstFinFormato6
'Dim lvPrincipalEstGanPer() As tForEvalResumenEstFinFormato6
'Dim lvPrincipalCoefiFinan() As tForEvalResumenEstFinFormato6
'Dim lvDetalleActivos() As tForEvalEstFinFormato6 'matriz para activos
'Dim lvDetallePasivos() As tForEvalEstFinFormato6 'matriz para pasivos
Dim lvPrincipalActivos() As tFormEvalPrincipalEstFinFormato6    'Matriz Principal-> Activos
Dim lvPrincipalPasivos() As tFormEvalPrincipalEstFinFormato6    'Matriz Principal-> Pasivos
Dim lvPrincipalEstGanPer() As tFormEvalPrincipalEstFinFormato6  'Matriz Principal-> Ganancias y pérdidas
Dim lvPrincipalCoefiFinan() As tFormEvalPrincipalEstFinFormato6 'Matriz Principal-> Coeficiente Financiero
'Detalle Activos y Pasivos
Dim lvDetalleActivos() As tFormEvalDetalleEstFinFormato6 'Matriz Detalle->Activos
Dim lvDetallePasivos() As tFormEvalDetalleEstFinFormato6 'Matriz Detalle->Pasivos
'<***** Fin LUCV20170915

Dim MatIfiGastoNego As Variant 'matriz de ifis de fluj de caja
Dim MatIfiFlujoCajaHistorico As Variant 'LUCV20171015, Agregó según ERS0512017
Dim MatIfiGastoFami As Variant
Dim MatReferidos As Variant
Dim MatIfiPasivo9() As Variant

Dim lnTotActivo1 As Double
Dim lnTotActivo2 As Double
Dim lnTotPasivo1 As Double
Dim lnTotPasivo2 As Double
Dim lnResulEjer1 As Double
Dim lnResulEjer2 As Double
Dim lnResulAcum1 As Double
Dim lnResulAcum2 As Double
Dim lnCapiAdici1 As Double
Dim lnExceReval1 As Double
Dim lnReservaLe1 As Double
Dim lnCapiAdici2 As Double
Dim lnExceReval2 As Double
Dim lnReservaLe2 As Double

Dim lnVtasContado As Currency
Dim lnCobrosCtaCre As Currency
Dim lnCobrosActFijo As Currency
Dim lnEgrePorCom As Currency
Dim lcFecRegEF As String
Dim ldFecRegEF As Date
Dim lnMontoSol As Currency
Dim lnCompraDeuda As Currency
Dim lnMontoAmpliado As Currency

Dim lcMovNro As String 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018

Public Function DisableCloseButton(frm As Form) As Boolean
'PURPOSE: Removes X button from a form
'EXAMPLE: DisableCloseButton Me
'RETURNS: True if successful, false otherwise
'NOTES:   Also removes Exit Item from
'         Control Box Menu
    Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long
    
    lHndSysMenu = GetSystemMenu(frm.hwnd, 0)

    'remove close button
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)
   'Remove seperator bar
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
    'Return True if both calls were successful
    DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)
End Function
Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtGiroNeg2.SetFocus
    End If
End Sub

Private Sub CboFecRegEEFF_Click()
'    If KeyAscii = 13 Then
        If MsgBox("¿Está seguro de obtener datos de esta fecha?", vbQuestion + vbYesNo, "Pregunta") = vbNo Then Exit Sub
        
        Dim oNCred As COMDCredito.DCOMFormatosEval
        Set oNCred = New COMDCredito.DCOMFormatosEval

        lcFecRegEF = Format(Me.CboFecRegEEFF.Text, "yyyyMMdd")
        ldFecRegEF = CDate(Me.CboFecRegEEFF.Text)
'        If IsDate(mskFecReg.Text) Then
            'busca fecha
            If oNCred.ValidaFechaCredFormEval(fsCtaCod, 6, lcFecRegEF).RecordCount = 0 Then
                MsgBox "No existe datos en la fecha ingresada.", vbOKOnly + vbInformation, "Atención"
    '            Exit Sub
            Else
                'carga datos
                Me.feActivos.Enabled = True
                Me.fePasivos.Enabled = True
                Me.feEstaGananPerd.Enabled = True
                Me.feFlujoCajaMensual.Enabled = True
                Me.feGastosFamiliares.Enabled = True
                Me.feOtrosIngresos.Enabled = True
                Me.feDeclaracionPDT.Enabled = True
                Me.feFlujoCajaHistorico.Enabled = True 'LUCV20171015, Agregó según ERS0512017
                
                Call CargaDatosBusqueda(fsCtaCod, 6, lcFecRegEF)
            End If
'        End If
'    End If
End Sub
Private Sub cmdAgregaEEFF_Click()
    Call frmCredFormEvalFormato6_EstFinan.Inicio(fsCtaCod, 6)
    Call CargaDatosCboFecEEFF
    Call CargarFlexEdit
    Call HabilitaControles(False, False, False)
    Call LimpiaControles
    
    If lnColocCondi = 4 Then
        Me.cmdInformeVista2.Enabled = False
        Me.cmdVerCar2.Enabled = False
    End If
End Sub
Private Sub cmdGuardar_Click()
    Dim oNCred As COMDCredito.DCOMFormatosEval
    Dim i, j As Integer
    Dim nId As String
    Dim rsBuscaFe As ADODB.Recordset
    Dim nSuma As Double
    Set oNCred = New COMDCredito.DCOMFormatosEval
    Dim lnNumForm As Integer
    Dim GrabarDatos As Boolean
    Dim MatReferidos As Variant
    Dim oDCred As COMDCredito.DCOMFormatosEval
    Set oDCred = New COMDCredito.DCOMFormatosEval
    lnNumForm = 6
    Dim nResp As Integer
    Dim lsMensajeIfi As String 'LUCV20161115
    
    ' valida si se ingresó la fecha de eval
    'If Not IsDate(Me.mskFecReg) Then
    If lcFecRegEF = "" Then
        MsgBox "Seleccione una fecha en los Estados Financieros ...", vbOKOnly + vbInformation, "Atención"
        Exit Sub
    End If
    
    ' valida que la feha no exista
'    nResp = 0 'No reemplaza datos
'    If oNCred.ValidaFechaCredFormEval(fsCtaCod, 6, lcFecRegEF).RecordCount > 0 Then
'        If MsgBox("La fecha que ingresó ya fue registrada, " & Chr(10) & "Desea Reemplazar con estos nuevos datos? ", vbYesNo, "Atención") = vbNo Then Exit Sub
        nResp = 1 'Si reeemplaza datos
'    End If
    
    If spnTiempoLocalAnio.valor = 0 And Me.spnTiempoLocalMes.valor = 0 Then
        MsgBox "El tiempo del Local en años y meses no debe ser cero", vbInformation + vbOKOnly, "Atención"
        spnTiempoLocalAnio.SetFocus
        Exit Sub
    End If

    If OptCondLocal2.iTem(1).value = False And OptCondLocal.iTem(2).value = False And _
        OptCondLocal2.iTem(3).value = False And OptCondLocal2.iTem(4).value = False Then
            MsgBox "Seleccione una condición del local.", vbOKOnly + vbInformation, "Atención"
            OptCondLocal2.iTem(1).SetFocus
            Exit Sub
    End If

     If txtFecUltEndeuda.Text = "__/__/____" Then
            MsgBox "Ingrese la fecha Ultima de Endeudamiento...", vbOKOnly + vbInformation, "Atención"
            txtFecUltEndeuda.SetFocus
            Exit Sub
     End If
     
     If txtFechaEvaluacion.Text = "__/__/____" Then
            MsgBox "Ingrese la fecha de evaluación...", vbOKOnly + vbInformation, "Atención"
            Exit Sub
     End If
   
    '-- verifica si activos tiene datos
        nSuma = 0
        If UBound(lvPrincipalActivos) > 0 Then
            For i = 1 To UBound(lvPrincipalActivos)
                If CDbl(Me.feActivos.TextMatrix(i, 4)) > 0 Then
                    nSuma = nSuma + CDbl(Me.feActivos.TextMatrix(i, 4))
                End If
            Next i
        End If
        If nSuma = 0 Then
            MsgBox "Ingrese datos en los Activos...", vbInformation + vbOKOnly, "Atención"
            Me.SSTabIngresos2.Tab = 0
            Me.feActivos.SetFocus
            Exit Sub
        End If
    
        If CCur(Me.feActivos.TextMatrix(1, 4)) = 0 Then
            MsgBox "El monto de Activo Corriente no debe ser cero.", vbOKOnly, "Atención"
            Me.SSTabIngresos2.Tab = 0
            Me.feActivos.SetFocus
            Exit Sub
        End If
        If CCur(Me.feActivos.TextMatrix(17, 4)) = 0 Then
            MsgBox "El monto de Total Activo no debe ser cero.", vbOKOnly, "Atención"
            Me.feActivos.SetFocus
            Exit Sub
        End If
        If CCur(Me.feActivos.TextMatrix(17, 4)) <> CCur(Me.fePasivos.TextMatrix(25, 4)) Then
            Me.SSTabIngresos2.Tab = 0
            MsgBox "El Activo y Pasivo no cuadran, por favor verificar.", vbOKOnly, "Atención"
            Exit Sub
        End If
    
    '-- verifica si pasivos tiene datos
        nSuma = 0
        If UBound(lvPrincipalPasivos) > 0 Then
            For i = 1 To UBound(lvPrincipalPasivos)
                If CDbl(Me.fePasivos.TextMatrix(i, 4)) > 0 Then
                    nSuma = nSuma + CDbl(Me.fePasivos.TextMatrix(i, 4))
                End If
            Next i
        End If
        If nSuma = 0 Then
            MsgBox "Ingrese datos en los Pasivos...", vbOKOnly + vbInformation, "Atención"
            Me.SSTabIngresos2.Tab = 0
            Me.fePasivos.SetFocus
            Exit Sub
        End If
    
        If CCur(Me.fePasivos.TextMatrix(1, 4)) = 0 Then
            MsgBox "El monto de Pasivo Corriente no debe ser cero.", vbOKOnly, "Atención"
            Me.SSTabIngresos2.Tab = 0
            Me.fePasivos.SetFocus
            Exit Sub
        End If
        If CCur(Me.fePasivos.TextMatrix(23, 4)) = 0 Then
            MsgBox "El monto de Total Pasivo no debe ser cero.", vbOKOnly, "Atención"
            Me.fePasivos.SetFocus
            Exit Sub
        End If
    
        If CCur(Me.fePasivos.TextMatrix(24, 4)) = 0 Then
            MsgBox "El monto del Patrimonio no debe ser cero.", vbOKOnly, "Atención"
            Me.SSTabIngresos2.Tab = 0
            Me.fePasivos.SetFocus
            Exit Sub
        End If
    
    '-- verifica si est ganan y perd tiene datos
        nSuma = 0
        If UBound(lvPrincipalEstGanPer) > 0 Then
            For i = 1 To UBound(lvPrincipalEstGanPer)
                If CDbl(Me.feEstaGananPerd.TextMatrix(i, 2)) > 0 Then
                    nSuma = nSuma + CDbl(Me.feEstaGananPerd.TextMatrix(i, 2))
                End If
            Next i
        End If
        If nSuma = 0 Then
            MsgBox "Ingrese datos en los Estados de Ganancias y Perdidas...", vbOKOnly + vbInformation, "Atención"
            Me.SSTabIngresos2.Tab = 1
            Me.feEstaGananPerd.SetFocus
            Exit Sub
        End If
    
    'valida flujo de caja
    If CCur(Me.feFlujoCajaMensual.TextMatrix(22, 2)) = 0 Then 'saldo disponible de flujo de caja
        MsgBox "El Saldo disponible del Flujo de Caja no puede ser cero, por favor verifique...", vbOKOnly + vbInformation, "Atención"
        Me.SSTabIngresos2.Tab = 3
        Me.feFlujoCajaMensual.SetFocus
        SendKeys "{TAB}"
        Exit Sub
    End If

    'Si es refinanciado no valida ingreso de datos de propuesta de crédito.
    If lnColocCondi <> 4 Then
        If txtFechaVisita.Text = "__/__/____" Then
            MsgBox "Ingrese la fecha de visita en la Pestaña de Propuesta de Crédito...", vbOKOnly + vbInformation, "Atención"
            Me.SSTabIngresos2.Tab = 4
            txtFechaVisita.SetFocus
            Exit Sub
        End If
    End If
    
    '--- valida coment y referidos - si condicion de credito es NUEVO obliga a llenar
    'If lnColocCondi = 1 Then 'LUCV20171115, Comentó según correo RUSI
    If Not fbTieneReferido6Meses Then
        If txtComentario.Text = "" Then
                MsgBox "Ingrese el Comentario en la Pestaña Comentarios y Referidos...", vbOKOnly + vbInformation, "Atención"
                Me.SSTabIngresos2.Tab = 5
                txtComentario.SetFocus
                Exit Sub
        End If
        If ValidaDatos = False Then 'propuesta del credito
            SSTabIngresos2.Tab = 4
            Exit Sub
        End If
        
        If ValidaDatosReferencia = False Then 'Contenido de feReferidos2: Referidos
            SSTabIngresos2.Tab = 5
            Exit Sub
        End If
    End If

    'LUCV20161115, Agregó->Según ERS068-2016
    If Not ValidaIfiExisteCompraDeuda(fsCtaCod, MatIfiGastoFami, MatIfiGastoNego, lsMensajeIfi) Or Len(Trim(lsMensajeIfi)) > 0 Then
        MsgBox "Ifi y Cuota registrada en detalle de cambio de estructura de pasivos no coincide:  " & Chr(10) & Chr(10) & " " & lsMensajeIfi & " ", vbInformation, "Aviso"
        Me.SSTabIngresos2.Tab = 3
        Exit Sub
    End If

    If MsgBox("Los Datos ingresados se guardarán, ¿ Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub

'----------------------CABECERA
'oDCred.BeginTrans
If nResp = 1 Then
'Format(mskFecReg.Text, "yyyymmdd")
    Call oDCred.EliminaFormatoEvaluacion(fsCtaCod, lcFecRegEF)
End If

'    lvVectorDatosGrles(1) = fsCtaCod
'    lvVectorDatosGrles(2) = 1
'    lvVectorDatosGrles(3) = txtGiroNeg2.Text
'    lvVectorDatosGrles(4) = CInt(spnExpEmpAnio.valor)
'    lvVectorDatosGrles(5) = CInt(spnExpEmpMes.valor)
'    lvVectorDatosGrles(6) = CInt(spnTiempoLocalAnio.valor)
'    lvVectorDatosGrles(7) = CInt(spnTiempoLocalMes.valor)
'    lvVectorDatosGrles(8) = CDbl(txtUltEndeuda.Text)
'    lvVectorDatosGrles(9) = Format(txtFecUltEndeuda.Text, "yyyymmdd")
'    lvVectorDatosGrles(10) = lnCondLocal
'    lvVectorDatosGrles(11) = IIf(txtCondLocalOtros.Visible = False, "", txtCondLocalOtros.Text)
'    lvVectorDatosGrles(12) = CDbl(Me.txtExposicioDCredito2.Text)
'    lvVectorDatosGrles(13) = Format(txtFechaEvaluacion.Text, "yyyymmdd")
'    lvVectorDatosGrles(14) = lnNumForm
'    lvVectorDatosGrles(15) = txtComentario.Text
'    lvVectorDatosGrles(16) = Format(mskFecReg, "yyyyMMdd")

    '->***** LUCV20171015, Agregó: ERS0512017 ->*****
    Dim MatFlujoCaja As Variant
    Set MatFlujoCaja = Nothing
    ReDim MatFlujoCaja(1, 5)
        For i = 1 To 1
            MatFlujoCaja(i, 1) = txtIncrVentasContado
            MatFlujoCaja(i, 2) = txtIncrCompraMercaderia
            MatFlujoCaja(i, 3) = txtIncrPagoPersonal
            MatFlujoCaja(i, 4) = txtIncrGastoVentas
            MatFlujoCaja(i, 5) = txtIncrConsumo
        Next i
        
    If IsArray(MatFlujoCaja) Then
        If UBound(MatFlujoCaja) Then
                For i = 1 To UBound(MatFlujoCaja)
                    Call oDCOMFormatosEval.InsertaCredFormEvalParamFlujoCaja(fsCtaCod, nFormato, MatFlujoCaja(i, 1), MatFlujoCaja(i, 2), MatFlujoCaja(i, 3), MatFlujoCaja(i, 4), MatFlujoCaja(i, 5))
                Next i
        End If
    End If
    '<-***** Fin LUCV20171015 <-*****

Call oDCred.GrabarCredFormEvalCabeceraFormato1_5(fsCtaCod, _
                                                1, _
                                                txtGiroNeg2.Text, _
                                                CInt(spnExpEmpAnio.valor), _
                                                CInt(spnExpEmpMes.valor), _
                                                CInt(spnTiempoLocalAnio.valor), _
                                                CInt(spnTiempoLocalMes.valor), _
                                                CDbl(txtUltEndeuda.Text), _
                                                Format(txtFecUltEndeuda.Text, "yyyymmdd"), _
                                                lnCondLocal, _
                                                IIf(txtCondLocalOtros.Visible = False, "", txtCondLocalOtros.Text), _
                                                CDbl(Me.txtExposicionCredito2.Text), _
                                                Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                                                lnNumForm, _
                                                txtComentario.Text, _
                                                txtInversion.Text, _
                                                txtFinanciamiento.Text)

' lvPrincipalActivos, lvPrincipalActivos(i).vPP(J).nImporte

'--------------------- ACTIVOS
        If UBound(lvPrincipalActivos) > 0 Then
            For i = 1 To UBound(lvPrincipalActivos)
                If CDbl(Me.feActivos.TextMatrix(i, 4)) <> 0 Then
                    ''If i = 1 Or i = 10 Or i = 17 Then
                    If i = 17 Then
                        Call oDCred.AgregaCredFormEvalEstFinActivosPasivosFormato6(fsCtaCod, 6, lcFecRegEF, chkAudit.value, CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), CDbl(Me.feActivos.TextMatrix(i, 2)), CDbl(Me.feActivos.TextMatrix(i, 3)), CDbl(Me.feActivos.TextMatrix(i, 4)))
                    Else
                        Call oDCred.AgregaCredFormEvalEstFinActivosPasivosFormato6Det(fsCtaCod, 6, lcFecRegEF, chkAudit.value, CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), CDbl(Me.feActivos.TextMatrix(i, 2)), CDbl(Me.feActivos.TextMatrix(i, 3)), CDbl(Me.feActivos.TextMatrix(i, 4)))
                    End If
                End If
            Next i
        End If

        '-- activos det DETALLE formato6
        If UBound(lvPrincipalActivos) > 0 Then
            For i = 1 To UBound(lvPrincipalActivos)
                If lvPrincipalActivos(i).nDetPP > 0 Then
                    For j = 1 To UBound(lvPrincipalActivos(i).vPP)
                        If lvPrincipalActivos(i).vPP(j).nImporte <> 0 Then
                            Call oDCred.AgregaCredFormEvalEstFinActivosPasivosFormato6DetDetalle( _
                                fsCtaCod, 6, lcFecRegEF, chkAudit.value, _
                                CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), 1, _
                                Format(lvPrincipalActivos(i).vPP(j).dFecha, "yyyyMMdd"), _
                                lvPrincipalActivos(i).vPP(j).cDescripcion, _
                                lvPrincipalActivos(i).vPP(j).nImporte, 0, "")
                        End If
                    Next j
                End If

                If lvPrincipalActivos(i).nDetPE > 0 Then
                    For j = 1 To UBound(lvPrincipalActivos(i).vPE)
                        If lvPrincipalActivos(i).vPE(j).nImporte <> 0 Then
                            Call oDCred.AgregaCredFormEvalEstFinActivosPasivosFormato6DetDetalle( _
                                fsCtaCod, 6, lcFecRegEF, chkAudit.value, _
                                CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), 2, _
                                Format(lvPrincipalActivos(i).vPE(j).dFecha, "yyyyMMdd"), _
                                lvPrincipalActivos(i).vPE(j).cDescripcion, 0, _
                                lvPrincipalActivos(i).vPE(j).nImporte, "")
                        End If
                    Next j
                End If
            Next i
        End If

'-------------------------- PASIVOS
        
        If UBound(lvPrincipalPasivos) > 0 Then
            For i = 1 To UBound(lvPrincipalPasivos)
                If CDbl(Me.fePasivos.TextMatrix(i, 4)) <> 0 Then
                    If i = 23 Or i = 24 Or i = 25 Then
                        Call oDCred.AgregaCredFormEvalEstFinActivosPasivosFormato6(fsCtaCod, lnNumForm, lcFecRegEF, _
                            chkAudit.value, CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), CDbl(Me.fePasivos.TextMatrix(i, 2)), _
                            CDbl(Me.fePasivos.TextMatrix(i, 3)), CDbl(Me.fePasivos.TextMatrix(i, 4)))
                    Else
                        Call oDCred.AgregaCredFormEvalEstFinActivosPasivosFormato6Det(fsCtaCod, lnNumForm, lcFecRegEF, _
                            chkAudit.value, CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), _
                            CDbl(Me.fePasivos.TextMatrix(i, 2)), CDbl(Me.fePasivos.TextMatrix(i, 3)), CDbl(Me.fePasivos.TextMatrix(i, 4)))
                    End If
                End If
            Next i
        End If

        '-- pasivos det DETALLE formato6
        If UBound(lvPrincipalPasivos) > 0 Then
            For i = 1 To UBound(lvPrincipalPasivos)
                
                If lvPrincipalPasivos(i).nDetPP > 0 Then
                    For j = 1 To UBound(lvPrincipalPasivos(i).vPP)
                        If lvPrincipalPasivos(i).vPP(j).nImporte <> 0 Then
                            lcCodifi = IIf(CInt(Me.fePasivos.TextMatrix(i, 6)) = 109 Or CInt(Me.fePasivos.TextMatrix(i, 6)) = 201, Right(lvPrincipalPasivos(i).vPP(j).cDescripcion, 13), "") 'LUCV20161115, Modificó->Según ERS068-2016 [Modificó 8 por 13]
                        
                            Call oDCred.AgregaCredFormEvalEstFinActivosPasivosFormato6DetDetalle( _
                                fsCtaCod, lnNumForm, lcFecRegEF, chkAudit.value, _
                                CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), 1, _
                                Format(lvPrincipalPasivos(i).vPP(j).dFecha, "yyyyMMdd"), _
                                lvPrincipalPasivos(i).vPP(j).cDescripcion, _
                                lvPrincipalPasivos(i).vPP(j).nImporte, 0, lcCodifi)
                        End If
                    Next j
                End If

                If lvPrincipalPasivos(i).nDetPE > 0 Then
                    For j = 1 To UBound(lvPrincipalPasivos(i).vPE)
                        If lvPrincipalPasivos(i).vPE(j).nImporte <> 0 Then
                            lcCodifi = IIf(CInt(Me.fePasivos.TextMatrix(i, 6)) = 109 Or CInt(Me.fePasivos.TextMatrix(i, 6)) = 201, Right(lvPrincipalPasivos(i).vPE(j).cDescripcion, 13), "") 'LUCV20161115, Modificó->Según ERS068-2016 [Modificó 8 por 13]
                        
                            Call oDCred.AgregaCredFormEvalEstFinActivosPasivosFormato6DetDetalle( _
                                fsCtaCod, lnNumForm, lcFecRegEF, chkAudit.value, _
                                CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), 2, _
                                Format(lvPrincipalPasivos(i).vPE(j).dFecha, "yyyyMMdd"), _
                                lvPrincipalPasivos(i).vPE(j).cDescripcion, 0, _
                                lvPrincipalPasivos(i).vPE(j).nImporte, lcCodifi)
                        End If
                    Next j
                End If
            Next i
        End If

'------------------- ESTADOS DE GANANCIAS Y PERDIDAS

        If UBound(lvPrincipalEstGanPer) > 0 Then
            For i = 1 To UBound(lvPrincipalEstGanPer)
                If Abs(CDbl(Me.feEstaGananPerd.TextMatrix(i, 2))) <> 0 Then
                    Call oDCred.AgregaCredFormEvalEstFinEstGanPerFormato6(fsCtaCod, lnNumForm, lcFecRegEF, chkAudit.value, CInt(Me.feEstaGananPerd.TextMatrix(i, 3)), CInt(Me.feEstaGananPerd.TextMatrix(i, 4)), CDbl(Me.feEstaGananPerd.TextMatrix(i, 2)))
                End If
            Next i
        End If

'------------------ COEFICIENTE FINANCIERO
    If Me.feCoeFinan.rows - 1 > 0 Then
        For i = 1 To Me.feCoeFinan.rows - 1
        
            'If Abs(CDbl(Me.feCoeFinan.TextMatrix(i, 2))) <> "0.00" Then
            If Me.feCoeFinan.TextMatrix(i, 2) <> "0.00" Then
                Call oDCred.AgregaCredFormEvalCoeficienteFinanFormato6(fsCtaCod, lnNumForm, lcFecRegEF, CInt(Me.feCoeFinan.TextMatrix(i, 3)), _
                                    CInt(Me.feCoeFinan.TextMatrix(i, 4)), CDbl(IIf(Right(Me.feCoeFinan.TextMatrix(i, 2), 1) = "%", Left(Me.feCoeFinan.TextMatrix(i, 2), Len(Me.feCoeFinan.TextMatrix(i, 2)) - 1), Me.feCoeFinan.TextMatrix(i, 2))))
            End If
        Next i
    End If

'------------------ FLUJO DE CAJA
    If Me.feFlujoCajaMensual.rows - 1 > 0 Then
        For i = 1 To Me.feFlujoCajaMensual.rows - 1
            If Abs(CCur(Me.feFlujoCajaMensual.TextMatrix(i, 2))) <> 0 Then
                Call oDCred.InsertaCredFormFlujoCajaMensual(fsCtaCod, lnNumForm, 7027, CInt(Me.feFlujoCajaMensual.TextMatrix(i, 4)), Trim(Me.feFlujoCajaMensual.TextMatrix(i, 1)), CCur(Me.feFlujoCajaMensual.TextMatrix(i, 2)))
            End If
        Next i
    End If
'------------------ IFIS FLUJO DE CAJA
    'Cuotas IFIs ->FLUJO DE CAJA
    If IsArray(MatIfiGastoNego) Then 'Pregunto Matriz no Sea Vacia
        If UBound(MatIfiGastoNego) > 0 Then
        Call oDCred.EliminaCredFormEvalGrillaOtrasIfis(fsCtaCod, lnNumForm, 7027)
            For i = 0 To UBound(MatIfiGastoNego) - 1
                Call oDCred.InsertaCredFormEvalCuotaIFI(fsCtaCod, lnNumForm, 7027, gCodCuotaIfiFlujoCaja, MatIfiGastoNego(i, 0), MatIfiGastoNego(i, 1), CDbl(MatIfiGastoNego(i, 2)))
            Next i
        End If
    End If

'-*****>LUCV20171015, Agregó según ERS0512017:
'------------------FLUJO DE CAJA HISTORICO
    If Me.feFlujoCajaHistorico.rows - 1 > 0 Then
        For i = 1 To Me.feFlujoCajaHistorico.rows - 1
            If Abs(CCur(Me.feFlujoCajaHistorico.TextMatrix(i, 2))) <> 0 Then
                Call oDCred.InsertaCredFormFlujoCajaMensual(fsCtaCod, lnNumForm, 7027, CInt(Me.feFlujoCajaHistorico.TextMatrix(i, 4)), Trim(Me.feFlujoCajaHistorico.TextMatrix(i, 1)), CCur(Me.feFlujoCajaHistorico.TextMatrix(i, 2)), 1)
            End If
        Next i
    End If
'------------------ IFIS FLUJO DE CAJA HISTORICO
'Cuotas IFIs ->FLUJO DE CAJA HISTORICO
    If IsArray(MatIfiFlujoCajaHistorico) Then
        If UBound(MatIfiFlujoCajaHistorico) > 0 Then
        Call oDCred.EliminaCredFormEvalGrillaOtrasIfis(fsCtaCod, lnNumForm, 7027, , 1)
            For i = 0 To UBound(MatIfiFlujoCajaHistorico) - 1
                Call oDCred.InsertaCredFormEvalCuotaIFI(fsCtaCod, lnNumForm, 7027, gCodCuotaIfiFlujoCaja, MatIfiFlujoCajaHistorico(i, 0), MatIfiFlujoCajaHistorico(i, 1), CDbl(MatIfiFlujoCajaHistorico(i, 2)), 1)
            Next i
        End If
    End If
'<-***** Fin LUCV20171015

'------------------  GASTOS FAMILIARES
    If Me.feGastosFamiliares.rows - 1 > 0 Then
        For i = 1 To Me.feGastosFamiliares.rows - 1
            If Abs(CDbl(Me.feGastosFamiliares.TextMatrix(i, 3))) <> 0 Then
                Call oDCred.InsertaCredFormEvalGastoFami(fsCtaCod, lnNumForm, CInt(Me.feGastosFamiliares.TextMatrix(i, 1)), CCur(Me.feGastosFamiliares.TextMatrix(i, 3)))
            End If
        Next i
    End If
'----------------------------------IFIS GASTOS FAMILIARES
    'Cuotas IFIs ->GastosFamiliar
    If IsArray(MatIfiGastoFami) Then 'Pregunto Matriz no Sea Vacia
        If UBound(MatIfiGastoFami) > 0 Then
        Call oDCred.EliminaCredFormEvalGrillaOtrasIfis(fsCtaCod, lnNumForm, gFormatoGastosFami)
            For i = 0 To UBound(MatIfiGastoFami) - 1
                Call oDCred.InsertaCredFormEvalCuotaIFI(fsCtaCod, lnNumForm, gFormatoGastosFami, gCodCuotaIfiGastoFami, MatIfiGastoFami(i, 0), MatIfiGastoFami(i, 1), CDbl(MatIfiGastoFami(i, 2)))
            Next i
        End If
    End If

'------------------  OTROS INGRESOS

    If Me.feOtrosIngresos.rows - 1 > 0 Then
        For i = 1 To Me.feOtrosIngresos.rows - 1
            If Abs(CDbl(Me.feOtrosIngresos.TextMatrix(i, 3))) <> 0 Then
                Call oDCred.InsertaCredFormEvalOtrosIngr(fsCtaCod, lnNumForm, CDbl(Me.feOtrosIngresos.TextMatrix(i, 1)), CCur(Me.feOtrosIngresos.TextMatrix(i, 3)))
            End If
        Next i
    End If

'------------------  Declaracion PDT (Formato 4, 5, y 6)
    If Me.feDeclaracionPDT.rows - 1 > 0 Then

        'Declaracion PDT
        sMes1 = DevolverMes(1, nAnio3, nMes3)
        sMes2 = DevolverMes(2, nAnio2, nMes2)
        sMes3 = DevolverMes(3, nAnio1, nMes1)

        Call oDCred.EliminaCredFormEvalPDT(fsCtaCod, lnNumForm)
        Call oDCred.EliminaCredFormEvalPDTDet(fsCtaCod, lnNumForm)
        
        Call oDCred.InsertaCredFormEvalPDT(fsCtaCod, lnNumForm, _
                    CInt(Me.feDeclaracionPDT.TextMatrix(1, 2)), _
                    CInt(Me.feDeclaracionPDT.TextMatrix(1, 3)), _
                    nMes1, _
                    nMes2, _
                    nMes3, _
                    nAnio1, _
                    nAnio2, _
                    nAnio3)

        For i = 1 To Me.feDeclaracionPDT.rows - 1
            If Abs(CCur(Me.feDeclaracionPDT.TextMatrix(i, 4))) > 0 Then
                Call oDCred.InsertaCredFormEvalPDTDet(fsCtaCod, lnNumForm, _
                    CInt(Me.feDeclaracionPDT.TextMatrix(i, 2)), _
                    CInt(Me.feDeclaracionPDT.TextMatrix(i, 3)), _
                    CCur(Me.feDeclaracionPDT.TextMatrix(i, 4)), _
                    CCur(Me.feDeclaracionPDT.TextMatrix(i, 5)), _
                    CCur(Me.feDeclaracionPDT.TextMatrix(i, 6)), _
                    CCur(Me.feDeclaracionPDT.TextMatrix(i, 7)), _
                    CCur(Replace(Me.feDeclaracionPDT.TextMatrix(i, 8), "%", "")))
            End If
        Next i
    End If

'--------------------- PROPUESTA DEL CREDITO
    'Si es refinanciado (4) no graba la propuesta del crédito(inf.visita)
    If lnColocCondi <> 4 Then
        Call oDCred.AgregaCredFormEvalPropuCredFormato6(fsCtaCod, lnNumForm, Format(txtFechaVisita, "yyyyMMdd"), txtEntornoFamiliar2.Text, txtGiroUbicacion2.Text, txtExperiencia2.Text, txtFormalidadNegocio2.Text, txtColaterales2, txtDestino2.Text)
    End If
'--------------------- COMENTARIOS/REFERIDOS (solo referidos, comentario esta en la cabcera)

    ReDim MatReferidos(feReferidos.rows - 1, 6)
    For i = 1 To feReferidos.rows - 1
        MatReferidos(i, 1) = feReferidos.TextMatrix(i, 0)
        MatReferidos(i, 2) = feReferidos.TextMatrix(i, 1)
        MatReferidos(i, 3) = feReferidos.TextMatrix(i, 2)
        MatReferidos(i, 4) = feReferidos.TextMatrix(i, 3)
        MatReferidos(i, 5) = feReferidos.TextMatrix(i, 4)
        MatReferidos(i, 6) = feReferidos.TextMatrix(i, 5)
     Next i
    
    If Me.feReferidos.rows - 1 > 1 Then
        For i = 1 To Me.feReferidos.rows - 1
            Call oDCred.InsertaCredFormEvalReferidos(fsCtaCod, lnNumForm, MatReferidos(i, 1), MatReferidos(i, 2), MatReferidos(i, 3), MatReferidos(i, 4), MatReferidos(i, 5), MatReferidos(i, 6))
        Next i
    End If
'--------------------------- recalcula ratios

        Set oDCred = New COMDCredito.DCOMFormatosEval
        Call oDCred.RecalculaIndicadoresyRatiosEvaluacion(fsCtaCod, 0, lcFecRegEF)

'----------------------------------------------------------------------------------------------------------------
'JOEP20180725 ERS034-2018
        Call EmiteFormRiesgoCamCred(sCtaCod)
'JOEP20180725 ERS034-2018
    
    'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 6", fsCtaCod, gCodigoCuenta
    Set objPista = New COMManejador.Pista
    'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 6", fsCtaCod, gCodigoCuenta 'RECO20161020 ERS0-2016
    'RECO20161020 ERS060-2016 **********************************************************
     Dim oNCOMColocEval As New NCOMColocEval
     'Dim lcMovNro As String 'LUCV20181220 Comentó, Anexo01 de Acta 199-2018
     lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
     
     If Not ValidaExisteRegProceso(fsCtaCod, gTpoRegCtrlEvaluacion) Then
        'lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Comentó, Anexo01 de Acta 199-2018
        'objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 6", fsCtaCod, gCodigoCuenta 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
        Call oNCOMColocEval.insEstadosExpediente(fsCtaCod, "Evaluacion de Credito", lcMovNro, "", "", "", 1, 2001, gTpoRegCtrlEvaluacion)
        Set oNCOMColocEval = Nothing
     End If
     'RECO FIN **************************************************************************
    
    'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
     If fnTipoRegMant = 1 Then
        objPista.InsertarPista gCredRegistrarEvaluacionCred, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 6", fsCtaCod, gCodigoCuenta 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
        Set objPista = Nothing 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
        MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
     Else
        objPista.InsertarPista gCredMantenimientoEvaluacionCred, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Evaluacion Credito Formato 6", fsCtaCod, gCodigoCuenta 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
        Set objPista = Nothing 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
        MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
     End If
    'Fin LUCV20181220
    
    
'/////////////////////////////////////////////////////// inicio prox pase
'    Dim oDCred As COMNCredito.NCOMFormatosEval
'    Dim nOK As Integer
'    Dim lvVectorDatosGrles(1 To 16) As Variant
'    Dim sMgsError As String
'    Set oDCred = New COMNCredito.NCOMFormatosEval
'
'    lvVectorDatosGrles(1) = fsCtaCod
'    lvVectorDatosGrles(2) = 1
'    lvVectorDatosGrles(3) = txtGiroNeg2.Text
'    lvVectorDatosGrles(4) = CInt(spnExpEmpAnio.valor)
'    lvVectorDatosGrles(5) = CInt(spnExpEmpMes.valor)
'    lvVectorDatosGrles(6) = CInt(spnTiempoLocalAnio.valor)
'    lvVectorDatosGrles(7) = CInt(spnTiempoLocalMes.valor)
'    lvVectorDatosGrles(8) = CDbl(txtUltEndeuda.Text)
'    lvVectorDatosGrles(9) = Format(txtFecUltEndeuda.Text, "yyyymmdd")
'    lvVectorDatosGrles(10) = lnCondLocal
'    lvVectorDatosGrles(11) = IIf(txtCondLocalOtros.Visible = False, "", txtCondLocalOtros.Text)
'    lvVectorDatosGrles(12) = CDbl(Me.txtExposicionCredito2.Text)
'    lvVectorDatosGrles(13) = Format(txtFechaEvaluacion.Text, "yyyymmdd")
'    lvVectorDatosGrles(14) = lnNumForm
'    lvVectorDatosGrles(15) = txtComentario.Text
'    lvVectorDatosGrles(16) = Format(mskFecReg, "yyyyMMdd")
'
'    ' lvPrincipalActivos, lvPrincipalActivos(i).vPP(J).nImporte
'
'    nOK = oDCred.ActualizaCredFormEvalFormato6( _
'             lvVectorDatosGrles, nResp, Me.feActivos.Recordset, Me.fePasivos.Recordset, _
'             Me.feEstaGananPerd.Recordset, Me.feCoeFinan.Recordset, Me.feFlujoCajaMensual.Recordset, _
'             Me.feGastosFamiliares.Recordset, Me.feOtrosIngresos.Recordset, _
'             Me.feDeclaracionPDT.Recordset)
'
'    If nOK = 0 Then
'        sMgsError = "Ha ocurrido un Error al tratar de grabar la evaluación..." & Chr(10) & "¿Desea Continuar?"
'        If MsgBox(sMgsError, vbCritical + vbYesNo, "Error") = vbNo Then
'            MsgBox "El formulario se cerrará", vbInformation, "Aviso"
'            Unload Me
'        End If
'    End If
'/////////////////////////////////////////////////////// fin prox pase

    MsgBox "Se guardaron los datos ingresados satisfactoriamente.", vbInformation, "Atención"
    fbGrabar = True
    
    Call CargaDatosCboFecEEFF
    Call CargarFlexEdit
    Call HabilitaControles(True, False, False)
    Call LimpiaControles
    If lnColocCondi = 4 Then
        Me.cmdInformeVista2.Enabled = False
        Me.cmdVerCar2.Enabled = False
    End If

End Sub

Private Sub LimpiaControles()
    lcFecRegEF = ""
    Me.txtFechaVisita.Text = "__/__/____"
    txtEntornoFamiliar2.Text = ""
    txtGiroUbicacion2.Text = ""
    txtExperiencia2.Text = ""
    txtFormalidadNegocio2.Text = ""
    txtColaterales2.Text = ""
    txtDestino2.Text = ""
    
    chkAudit.value = 0
    Me.feActivos.Enabled = False
    Me.fePasivos.Enabled = False
    Me.feEstaGananPerd.Enabled = False
    Me.feFlujoCajaMensual.Enabled = False
    Me.feGastosFamiliares.Enabled = False
    Me.feOtrosIngresos.Enabled = False
    Me.feDeclaracionPDT.Enabled = False
    Me.feFlujoCajaHistorico.Enabled = False 'LUCV20171015, Agregó según ERS0512017
    
    txtCapacidadNeta2.Text = "0.00"
    txtIngresoNeto2.Text = "0.00"
    txtExcedenteMensual2.Text = "0.00"
    txtRentabilidad.Text = "0.00"
    txtLiquidezCte.Text = "0.00"
    
    txtIncrVentasContado.Text = "0.00"
    txtIncrCompraMercaderia.Text = "0.00"
    txtIncrPagoPersonal.Text = "0.00"
    txtIncrGastoVentas.Text = "0.00"
    txtIncrConsumo.Text = "0.00"
End Sub

Private Sub CabeceraImpCuadros(ByVal prsInfVisita As ADODB.Recordset)

 Dim A As Integer
    Dim B As Integer
    Dim nFila As Integer
    Dim oDoc  As cPDF
    A = 50
    B = 29
    Set oDoc = New cPDF

    oDoc.WImage 40, 60, 35, 105, "Logo"
    oDoc.WTextBox 40, 60, 35, 390, prsInfVisita!Agencia, "F2", 10, hLeft
    
    oDoc.WTextBox 40, 60, 35, 390, "FECHA", "F2", 10, hRight
    oDoc.WTextBox 40, 60, 35, 450, Format(gdFecSis, "dd/mm/yyyy"), "F2", 10, hRight
    oDoc.WTextBox 40, 60, 35, 490, Format(Time, "hh:mm:ss"), "F2", 10, hRight
    
    oDoc.WTextBox 90 - B, 60, 15, 160, "Cliente", "F2", 10, hLeft
    oDoc.WTextBox 90 - B, 60, 15, 80, ":", "F2", 10, hRight
    oDoc.WTextBox 90 - B, 150, 15, 500, prsInfVisita!cPersNombre, "F1", 10, hjustify
    
    oDoc.WTextBox 90 - B, 350 + 20, 15, 160, "Analista", "F2", 10, hjustify
    oDoc.WTextBox 90 - B, 390 + 20, 15, 80, ":", "F2", 10, hjustify
    oDoc.WTextBox 90 - B, 402 + 20, 15, 500, prsInfVisita!UserAnalista, "F1", 10, hjustify
    
    oDoc.WTextBox 100 - B, 60, 15, 160, "Usuario", "F2", 10, hLeft
    oDoc.WTextBox 100 - B, 60, 15, 80, ":", "F2", 10, hRight
    oDoc.WTextBox 100 - B, 150, 15, 118, gsCodUser, "F1", 10, hjustify
    
    oDoc.WTextBox 100 - B, 350, 15, 160, "Producto", "F2", 10, hjustify
    oDoc.WTextBox 100 - B, 390, 15, 80, ":", "F2", 10, hjustify
    oDoc.WTextBox 100 - B, 402, 15, 118, prsInfVisita!cConsDescripcion, "F1", 10, hjustify
    
    oDoc.WTextBox 110 - B, 60, 15, 160, "Credito", "F2", 10, hLeft
    oDoc.WTextBox 110 - B, 60, 15, 80, ":", "F2", 10, hRight
    oDoc.WTextBox 110 - B, 150, 15, 500, prsInfVisita!cCtaCod, "F1", 10, hjustify
    
    oDoc.WTextBox 120 - B, 60, 15, 160, "Cod. Cliente", "F2", 10, hLeft
    oDoc.WTextBox 120 - B, 60, 15, 80, ":", "F2", 10, hRight
    oDoc.WTextBox 120 - B, 150, 15, 500, prsInfVisita!cPersCod, "F1", 10, hjustify
    
    oDoc.WTextBox 120 - B, 270, 15, 160, "Doc. Natural", "F2", 10, hjustify
    oDoc.WTextBox 120 - B, 328, 15, 80, ":", "F2", 10, hjustify
    oDoc.WTextBox 120 - B, 335, 15, 500, prsInfVisita!DNI, "F1", 10, hjustify
    
    oDoc.WTextBox 120 - B, 400, 15, 160, "Doc. Juridico", "F2", 10, hjustify
    oDoc.WTextBox 120 - B, 460, 15, 80, ":", "F2", 10, hjustify
    oDoc.WTextBox 120 - B, 470, 15, 500, IIf(prsInfVisita!Ruc = "NULL", "-", prsInfVisita!Ruc), "F1", 10, hjustify

End Sub

Private Sub cmdImpEEFF_Click()

    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    
    Dim oDCredUltFec As COMDCredito.DCOMFormatosEval
    Set oDCredUltFec = New COMDCredito.DCOMFormatosEval

    Dim RSUltDef As ADODB.Recordset

    If lcFecRegEF = "" Then
        Set RSUltDef = oDCredUltFec.ObtieneUltFecEEFFForm6(fsCtaCod)
        lcFecRegEF = RSUltDef!cFecUltEEFF
    End If

    If lnPrdEstado = 2000 Then
        MsgBox "El crédito debe estar por lo menos en estado Sugerido para mostrar el reporte.", vbOKOnly, "Mensaje"
        Exit Sub
    End If

    Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(fsCtaCod)
    Set rsDatActivos = oDCOMFormatosEval.ObtenerFormatosEvalActivos(fsCtaCod, lnNumForm, lcFecRegEF)
    Set rsDatPasivos = oDCOMFormatosEval.ObtenerFormatosEvalPasivos(fsCtaCod, lnNumForm, lcFecRegEF)
    Set rsDatActivosForm6Det = oDCOMFormatosEval.ObtenerFormatosEvalActivosDet(fsCtaCod, lnNumForm, lcFecRegEF)
    Set rsDatPasivosForm6det = oDCOMFormatosEval.ObtenerFormatosEvalPasivosDet(fsCtaCod, lnNumForm, lcFecRegEF)
    Set rsDatEstadoGananPerdForm6 = oDCOMFormatosEval.ObtenerFormatosEvalEstadoGanPerd(fsCtaCod, lnNumForm, lcFecRegEF)
    Set rsDatCoeFinanForm6 = oDCOMFormatosEval.ObtenerFormatosEvalCoeficienteFinan(fsCtaCod, lnNumForm, lcFecRegEF)
    
    Set rsDatFlujoCaja = oDCOMFormatosEval.RecuperaDatosCredEvalFlujoCaja(fsCtaCod)
    Set rsDatGastoFam = oDCOMFormatosEval.RecuperaDatosCredEvalGastosFam(fsCtaCod)
    Set rsDatOtrosIng = oDCOMFormatosEval.RecuperaDatosCredEvalOtrosIngr(fsCtaCod)
    Set rsDatPDT = oDCOMFormatosEval.RecuperaDatosCredEvalPDT(fsCtaCod)
    Set rsDatPDTDet = oDCOMFormatosEval.RecuperaDatosCredEvalPDTDet(fsCtaCod)
    Set rsDatIfiGastoFami = oDCOMFormatosEval.RecuperaDatosIfiCuota(fsCtaCod, lnNumForm, gFormatoGastosFami)
    Set rsDatRatios = oDCOMFormatosEval.RecuperaDatosRatios(fsCtaCod)
    
    If Not GeneraImpresionReporte6(rsInfVisita, rsDatActivos, rsDatPasivos, rsDatActivosForm6Det, _
                                    rsDatPasivosForm6det, rsDatEstadoGananPerdForm6, rsDatCoeFinanForm6, _
                                    rsDatFlujoCaja, rsDatGastoFam, rsDatOtrosIng, rsDatPDT, rsDatPDTDet, rsDatIfiGastoFami, _
                                    rsDatIfiGastoFami, rsDatRatios) Then
        MsgBox "Los Datos de la propuesta del Credito no han sido Registrados Correctamente", vbInformation, "Aviso"
    End If

    Set rsInfVisita = Nothing
    Set rsDatActivos = Nothing
    Set rsDatPasivos = Nothing
    Set rsDatActivosForm6Det = Nothing
    Set rsDatPasivosForm6det = Nothing
    Set rsDatEstadoGananPerdForm6 = Nothing
    Set rsDatCoeFinanForm6 = Nothing
    Set rsDatFlujoCaja = Nothing
    Set rsDatGastoFam = Nothing
    Set rsDatOtrosIng = Nothing
    Set rsDatPDT = Nothing
    Set rsDatPDTDet = Nothing
    Set rsDatIfiGastoFami = Nothing
    Set rsDatRatios = Nothing

End Sub

Private Sub cmdImprimir_Click()
       
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    
    Dim oDCredUltFec As COMDCredito.DCOMFormatosEval
    Set oDCredUltFec = New COMDCredito.DCOMFormatosEval

    Dim RSUltDef As ADODB.Recordset

    If lcFecRegEF = "" Then
        Set RSUltDef = oDCredUltFec.ObtieneUltFecEEFFForm6(fsCtaCod)
        If RSUltDef!cFecUltEEFF = "" Then
            MsgBox "Este crédito no tiene Evaluación", vbInformation + vbOKOnly, "Atención"
            Exit Sub
        Else
            lcFecRegEF = RSUltDef!cFecUltEEFF
        End If
        RSUltDef.Close
        Set RSUltDef = Nothing
    End If

    If lnPrdEstado = 2000 Then
        MsgBox "El crédito debe estar por lo menos en estado Sugerido para mostrar el reporte.", vbOKOnly, "Mensaje"
        Exit Sub
    End If

    Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(fsCtaCod)
    Set rsDatCoeFinanForm6 = oDCOMFormatosEval.ObtenerFormatosEvalCoeficienteFinan(fsCtaCod, lnNumForm, lcFecRegEF)
    Set rsDatFlujoCaja = oDCOMFormatosEval.RecuperaDatosCredEvalFlujoCaja(fsCtaCod)
    Set rsDatRatios = oDCOMFormatosEval.RecuperaDatosRatios(fsCtaCod)
    Set rsDatGastoFam = oDCOMFormatosEval.RecuperaDatosCredEvalGastosFam(fsCtaCod)
    Set rsDatOtrosIng = oDCOMFormatosEval.RecuperaDatosCredEvalOtrosIngr(fsCtaCod)
    Set rsDatActivos = oDCOMFormatosEval.ObtenerFormatosEvalActivos(fsCtaCod, lnNumForm, lcFecRegEF)
    Set rsDatPasivos = oDCOMFormatosEval.ObtenerFormatosEvalPasivos(fsCtaCod, lnNumForm, lcFecRegEF)
    
    If Not GeneraHojaEvalReporte6(rsInfVisita, rsDatActivos, rsDatPasivos, rsDatCoeFinanForm6, _
                                    rsDatFlujoCaja, rsDatRatios, rsDatGastoFam, rsDatOtrosIng) Then
        MsgBox "Los Datos de la propuesta del Credito no han sido Registrados Correctamente", vbInformation, "Aviso"
    End If

    Set rsInfVisita = Nothing
    Set rsDatCoeFinanForm6 = Nothing
    Set rsDatFlujoCaja = Nothing

End Sub

Private Sub cmdInformeVista2_Click()
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim rsInfVisita As ADODB.Recordset

    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsInfVisita = New ADODB.Recordset
    Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(fsCtaCod)
    Me.cmdInformeVista2.Enabled = False
    If (rsInfVisita.EOF And rsInfVisita.BOF) Then
        Set oDCOMFormatosEval = Nothing
        
        MsgBox "No existe datos para este reporte.", vbOKOnly, "Atención"
        Exit Sub
        
    End If

    Call CargaInformeVisitaPDF(rsInfVisita)
    'Call CargaInformeVisitaPDF_form6(rsInfVisita)
    cmdInformeVista2.Enabled = True
End Sub
'JOEP20180725 ERS034-2018
Private Sub cmdMNME_Click()
    Call frmCredFormEvalCredCel.Inicio(ActXCodCta.NroCuenta, 11)
End Sub
'JOEP20180725 ERS034-2018

Private Sub cmdVerCar2_Click()
    
    Dim oCred As COMNCredito.NCOMFormatosEval
    Dim oDCredSbs As COMDCredito.DCOMFormatosEval
    Dim R As ADODB.Recordset
    Dim lcDNI, lcRUC As String

    Dim RSbs, RDatFin1, RCap As ADODB.Recordset

    Set oCred = New COMNCredito.NCOMFormatosEval
    Call oCred.RecuperaDatosInformeComercial(ActXCodCta.NroCuenta, R)
    Set oCred = Nothing

    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para el Reporte...", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If lnPrdEstado = 2000 Then
        MsgBox "El crédito debe estar por lo menos en estado Sugerido para mostrar el reporte.", vbOKOnly, "Mensaje"
        Exit Sub
    End If
    
    lcDNI = Trim(R!dni_deudor)
    lcRUC = Trim(R!ruc_deudor)
    
    Set oDCredSbs = New COMDCredito.DCOMFormatosEval
        Set RSbs = oDCredSbs.RecuperaCaliSbs(lcDNI, lcRUC)
        Set RDatFin1 = oDCredSbs.RecuperaDatosFinan(ActXCodCta.NroCuenta, 6)
    Set oDCredSbs = Nothing
    
    Call ImprimeInformeCriteriosAceptacionRiesgoFormatoEval(ActXCodCta.NroCuenta, gsNomAge, gsCodUser, R, RSbs, RDatFin1)
End Sub

Private Sub feActivos_EnterCell()
    Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0)) 'celda que se activa el textbuscar
        Case 2, 4, 5, 6, 7, 8, 9, 11, 14, 16
            Me.feActivos.ListaControles = "0-0-1-1-0-0-0"
        Case Else
            Me.feActivos.ListaControles = "0-0-0-0-0-0-0"
        End Select

    Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0)) 'celda que  no se puede editar
        Case 1, 10, 17
            Me.feActivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case Else
            Me.feActivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        End Select
End Sub

Private Sub feActivos_GotFocus()
    If Me.feActivos.row = 16 And Me.feActivos.Col = 4 Then
        Me.fePasivos.SetFocus
    End If
End Sub

Private Sub feActivos_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feActivos.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        
        Select Case CInt(feActivos.TextMatrix(pnRow, 0))
            Case 15 'negativos
                If CCur(feActivos.TextMatrix(pnRow, pnCol)) > 0 Then
                    feActivos.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(feActivos.TextMatrix(pnRow, pnCol))) * -1, "#,#0.00")   '"0.00"
                End If
            Case 5 ' posi o negativo
            
            Case Else
                If CCur(feActivos.TextMatrix(pnRow, pnCol)) < 0 Then
                    feActivos.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(feActivos.TextMatrix(pnRow, pnCol))), "#,#0.00")  '"0.00"
                End If
        End Select
'        If feActivos.TextMatrix(pnRow, pnCol) < 0 Then
'            feActivos.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(feActivos.TextMatrix(pnRow, pnCol))), "#,#0.00")   '"0.00"
'        End If
        
    Else
        feActivos.TextMatrix(pnRow, pnCol) = "0.00"
    End If

    Call CalculaCeldas(1)
    Call CalculaCeldas(2)
End Sub

Private Sub feActivos_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)

Dim pnMonto As Double
Dim index As Integer
Dim nTotal As Double
      
    'If Me.mskFecReg = "__/__/____" Then
    If lcFecRegEF = "" Then
        MsgBox "Debe registrar una fecha", vbInformation + vbOKOnly, "Atención"
        Exit Sub
    End If
      
    If feActivos.TextMatrix(1, 0) = "" Then Exit Sub

    index = CInt(feActivos.TextMatrix(feActivos.row, 0))
    nTotal = 0
    Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0))
        Case 2, 4, 5, 6, 7, 8, 9, 11, 14, 16
            Set oFrm6 = New frmCredFormEvalDetalleFormato6
            
            If feActivos.Col = 2 Then 'column P.P.
                If IsArray(lvPrincipalActivos(index).vPP) Then
                    lvDetalleActivos = lvPrincipalActivos(index).vPP
                    nTotal = lvPrincipalActivos(index).nImportePP
                Else
                    ReDim lvDetalleActivos(0)
                End If
            End If
            
            If feActivos.Col = 3 Then 'column P.P.
                If IsArray(lvPrincipalActivos(index).vPE) Then
                    lvDetalleActivos = lvPrincipalActivos(index).vPE
                    nTotal = lvPrincipalActivos(index).nImportePE
                Else
                    ReDim lvDetalleActivos(0)
                End If
            End If
            
            If oFrm6.Registrar(True, 1, lvPrincipalActivos(index).cConcepto, lvDetalleActivos, lvDetalleActivos, nTotal, CInt(feActivos.TextMatrix(Me.feActivos.row, 5)), CInt(feActivos.TextMatrix(Me.feActivos.row, 6)), Trim(str(ldFecRegEF))) Then
                If feActivos.Col = 2 Then 'column P.P.
                    lvPrincipalActivos(index).vPP = lvDetalleActivos
                End If
                If feActivos.Col = 3 Then ' columna P.E.
                    lvPrincipalActivos(index).vPE = lvDetalleActivos
                End If
            End If

            If feActivos.Col = 2 Then
                Me.feActivos.TextMatrix(Me.feActivos.row, Me.feActivos.Col) = Format(nTotal, "#,#0.00")
                If nTotal <> 0 Then
                    lvPrincipalActivos(index).nImportePP = nTotal
                    lvPrincipalActivos(index).nDetPP = 1
                Else
                    lvPrincipalActivos(index).nImportePP = nTotal
                    lvPrincipalActivos(index).nDetPP = 0
                End If
            End If

            If feActivos.Col = 3 Then
                Me.feActivos.TextMatrix(Me.feActivos.row, Me.feActivos.Col) = Format(nTotal, "#,#0.00")
                If nTotal <> 0 Then
                    lvPrincipalActivos(index).nImportePE = nTotal
                    lvPrincipalActivos(index).nDetPE = 1
                Else
                    lvPrincipalActivos(index).nImportePE = nTotal
                    lvPrincipalActivos(index).nDetPE = 0
                End If
            End If
            
            Call CalculaCeldas(1)
            Call CalculaCeldas(2)
        End Select
End Sub

Private Sub feActivos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim Editar() As String
    Editar = Split(Me.feActivos.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    Select Case pnRow
        Case 2, 4, 5, 6, 7, 8, 9, 11, 14, 16
            Cancel = False
            SendKeys "{TAB}"
            Exit Sub
    End Select
End Sub


Private Sub CalculaCeldas(pnActPas As Integer)
    Dim m1, m2 As Double
    Dim s1, s2, s3, s4 As Double 'para pasivos y activos
    Dim s5, s6, s7, s8, s9 As Double 'para est gana y perdi

    Dim nPorcentajeVentas As Double
    Dim nPorcentajeCompras As Double
    Dim nMontoDeclarado As Double


    nPorcentajeVentas = 0: nPorcentajeCompras = 0: nMontoDeclarado = 0

    s1 = 0: s2 = 0: s3 = 0: s4 = 0
    s5 = 0: s6 = 0: s7 = 0: s8 = 0: s9 = 0

    If pnActPas = 1 Then '-- activos
            For i = 2 To 9
                s1 = s1 + CDbl(Me.feActivos.TextMatrix(i, 2))
                s2 = s2 + CDbl(Me.feActivos.TextMatrix(i, 3))
            Next i

            Me.feActivos.TextMatrix(1, 2) = Format(s1, "#,#0.00")
            Me.feActivos.TextMatrix(1, 3) = Format(s2, "#,#0.00")

            For i = 11 To 16
                s3 = s3 + CDbl(Me.feActivos.TextMatrix(i, 2))
                s4 = s4 + CDbl(Me.feActivos.TextMatrix(i, 3))
            Next i

            Me.feActivos.TextMatrix(10, 2) = Format(s3, "#,#0.00")
            Me.feActivos.TextMatrix(10, 3) = Format(s4, "#,#0.00")

            Me.feActivos.TextMatrix(17, 2) = Format(s1 + s3, "#,#0.00")
            Me.feActivos.TextMatrix(17, 3) = Format(s2 + s4, "#,#0.00")

            '-- TOTALIZA TOTAL PATRIMONIO EN PP y PE (TOT ACTIVO - TOT PASIVO)
'            Me.fePasivos.TextMatrix(17, 2) = Format(CDbl(Me.feActivos.TextMatrix(17, 2)) - CDbl(Me.fePasivos.TextMatrix(16, 2)), "#,#0.00")
'            Me.fePasivos.TextMatrix(17, 3) = Format(CDbl(Me.feActivos.TextMatrix(17, 3)) - CDbl(Me.fePasivos.TextMatrix(16, 3)), "#,#0.00")
            Me.fePasivos.TextMatrix(24, 2) = Format(CDbl(Me.feActivos.TextMatrix(17, 2)) - CDbl(Me.fePasivos.TextMatrix(23, 2)), "#,#0.00")
            Me.fePasivos.TextMatrix(24, 3) = Format(CDbl(Me.feActivos.TextMatrix(17, 3)) - CDbl(Me.fePasivos.TextMatrix(23, 3)), "#,#0.00")

            '-- TOTALIZA TOTAL PASIVO Y PATRIMONIO EN PP y PE (TOT PASIVO + TOT PATRIMONIO)
'            Me.fePasivos.TextMatrix(18, 2) = Format(CDbl(Me.fePasivos.TextMatrix(16, 2)) + CDbl(Me.fePasivos.TextMatrix(17, 2)), "#,#0.00")
'            Me.fePasivos.TextMatrix(18, 3) = Format(CDbl(Me.fePasivos.TextMatrix(16, 3)) + CDbl(Me.fePasivos.TextMatrix(17, 3)), "#,#0.00")
            Me.fePasivos.TextMatrix(25, 2) = Format(CDbl(Me.fePasivos.TextMatrix(23, 2)) + CDbl(Me.fePasivos.TextMatrix(24, 2)), "#,#0.00")
            Me.fePasivos.TextMatrix(25, 3) = Format(CDbl(Me.fePasivos.TextMatrix(23, 3)) + CDbl(Me.fePasivos.TextMatrix(24, 3)), "#,#0.00")

            For i = 1 To Me.feActivos.rows - 1

                m1 = CDbl(Me.feActivos.TextMatrix(i, 2))
                m2 = CDbl(Me.feActivos.TextMatrix(i, 3))

                Me.feActivos.TextMatrix(i, 4) = Format(m1 + m2, "#,#0.00")
                Me.feActivos.TextMatrix(i, 2) = Format(m1, "#,#0.00")
                Me.feActivos.TextMatrix(i, 3) = Format(m2, "#,#0.00")
            Next i

            Call CalculaCoeFinan
    End If

    s1 = 0: s2 = 0: s3 = 0: s4 = 0
    s5 = 0: s6 = 0: s7 = 0: s8 = 0: s9 = 0

    If pnActPas = 2 Then '-- pasivos
        
        lnTotActivo1 = CCur(Me.feActivos.TextMatrix(17, 2))
        lnTotActivo2 = CCur(Me.feActivos.TextMatrix(17, 3))
        lnTotPasivo1 = CCur(Me.fePasivos.TextMatrix(23, 2))
        lnTotPasivo2 = CCur(Me.fePasivos.TextMatrix(23, 3))
        lnResulEjer1 = CCur(Me.fePasivos.TextMatrix(21, 2))
        lnResulEjer2 = CCur(Me.fePasivos.TextMatrix(21, 3))
        lnResulAcum1 = CCur(Me.fePasivos.TextMatrix(22, 2))
        lnResulAcum2 = CCur(Me.fePasivos.TextMatrix(22, 3))

        lnCapiAdici1 = CCur(Me.fePasivos.TextMatrix(18, 2))
        lnCapiAdici2 = CCur(Me.fePasivos.TextMatrix(18, 3))
        lnExceReval1 = CCur(Me.fePasivos.TextMatrix(19, 2))
        lnExceReval2 = CCur(Me.fePasivos.TextMatrix(19, 3))
        lnReservaLe1 = CCur(Me.fePasivos.TextMatrix(20, 2))
        lnReservaLe2 = CCur(Me.fePasivos.TextMatrix(20, 3))
        
        'calcula capital
        Me.fePasivos.TextMatrix(17, 2) = Format(lnTotActivo1 - lnTotPasivo1 - lnResulEjer1 - lnResulAcum1 - lnCapiAdici1 - lnExceReval1 - lnReservaLe1, "#,#0.00")
        Me.fePasivos.TextMatrix(17, 3) = Format(lnTotActivo2 - lnTotPasivo2 - lnResulEjer2 - lnResulAcum2 - lnCapiAdici2 - lnExceReval2 - lnReservaLe2, "#,#0.00")

        For i = 2 To 9
            s1 = s1 + CDbl(Me.fePasivos.TextMatrix(i, 2))
            s2 = s2 + CDbl(Me.fePasivos.TextMatrix(i, 3))
        Next i

        Me.fePasivos.TextMatrix(1, 2) = Format(s1, "#,#0.00")
        Me.fePasivos.TextMatrix(1, 3) = Format(s2, "#,#0.00")

        For i = 11 To 15
            s3 = s3 + CDbl(Me.fePasivos.TextMatrix(i, 2))
            s4 = s4 + CDbl(Me.fePasivos.TextMatrix(i, 3))
        Next i

        Me.fePasivos.TextMatrix(10, 2) = Format(s3, "#,#0.00")
        Me.fePasivos.TextMatrix(10, 3) = Format(s4, "#,#0.00")

        For i = 17 To 22
            s5 = s5 + CDbl(Me.fePasivos.TextMatrix(i, 2))
            s6 = s6 + CDbl(Me.fePasivos.TextMatrix(i, 3))
        Next i

        Me.fePasivos.TextMatrix(16, 2) = Format(s5, "#,#0.00")
        Me.fePasivos.TextMatrix(16, 3) = Format(s6, "#,#0.00")


        '-- TOTALIZA TOTAL PASIVO EN PP y PE (PAS CORR + PAS NO CORR)
        Me.fePasivos.TextMatrix(23, 2) = Format(s1 + s3, "#,#0.00")
        Me.fePasivos.TextMatrix(23, 3) = Format(s2 + s4, "#,#0.00")

        '-- TOTALIZA TOTAL PATRIMONIO EN PP y PE (TOT ACTIVO - TOT PASIVO)
    '            Me.fePasivos.TextMatrix(17, 2) = Format(CDbl(Me.feActivos.TextMatrix(17, 2)) - CDbl(Me.fePasivos.TextMatrix(16, 2)), "#,#0.00")
    '            Me.fePasivos.TextMatrix(17, 3) = Format(CDbl(Me.feActivos.TextMatrix(17, 3)) - CDbl(Me.fePasivos.TextMatrix(16, 3)), "#,#0.00")
        Me.fePasivos.TextMatrix(24, 2) = Format(s5, "#,#0.00")
        Me.fePasivos.TextMatrix(24, 3) = Format(s6, "#,#0.00")

        '-- TOTALIZA TOTAL PASIVO Y PATRIMONIO EN PP y PE (TOT PASIVO + TOT PATRIMONIO)
        Me.fePasivos.TextMatrix(25, 2) = Format(CDbl(Me.fePasivos.TextMatrix(23, 2)) + CDbl(Me.fePasivos.TextMatrix(24, 2)), "#,#0.00")
        Me.fePasivos.TextMatrix(25, 3) = Format(CDbl(Me.fePasivos.TextMatrix(23, 3)) + CDbl(Me.fePasivos.TextMatrix(24, 3)), "#,#0.00")

        For i = 1 To Me.fePasivos.rows - 1

            m1 = CDbl(Me.fePasivos.TextMatrix(i, 2))
            m2 = CDbl(Me.fePasivos.TextMatrix(i, 3))

            Me.fePasivos.TextMatrix(i, 4) = Format(m1 + m2, "#,#0.00")
            Me.fePasivos.TextMatrix(i, 2) = Format(m1, "#,#0.00")
            Me.fePasivos.TextMatrix(i, 3) = Format(m2, "#,#0.00")
        Next i
        
        Call CalculaCoeFinan
    End If

    s1 = 0: s2 = 0: s3 = 0: s4 = 0
    s5 = 0: s6 = 0: s7 = 0: s8 = 0: s9 = 0

    If pnActPas = 3 Then '-- EST DE GANACIAS Y PERDIDAS

            For i = 1 To 2
                s5 = s5 + CDbl(Me.feEstaGananPerd.TextMatrix(i, 2))
            Next i

            For i = 4 To 5
                s6 = s6 + CDbl(Me.feEstaGananPerd.TextMatrix(i, 2))
            Next i

            For i = 7 To 8
                s7 = s7 + CDbl(Me.feEstaGananPerd.TextMatrix(i, 2))
            Next i

            For i = 10 To 15
                If (i = 12 Or i = 14) Then
                    s8 = s8 - CCur(Me.feEstaGananPerd.TextMatrix(i, 2))
                Else
                    s8 = s8 + CCur(Me.feEstaGananPerd.TextMatrix(i, 2))
                End If
            Next i

            For i = 17 To 18
                s9 = s9 + CDbl(Me.feEstaGananPerd.TextMatrix(i, 2))
            Next i

            Me.feEstaGananPerd.TextMatrix(3, 2) = Format(s5, "#,#0.00")
            Me.feEstaGananPerd.TextMatrix(6, 2) = Format(s5 - s6, "#,#0.00")
            Me.feEstaGananPerd.TextMatrix(9, 2) = Format(s5 - s6 - s7, "#,#0.00")
            Me.feEstaGananPerd.TextMatrix(16, 2) = Format((s5 - s6 - s7) + s8, "#,#0.00")
            Me.feEstaGananPerd.TextMatrix(19, 2) = Format(((s5 - s6 - s7) + s8) - s9, "#,#0.00")

            Call CalculaCoeFinan
    End If

    If pnActPas = 4 Then  '--Promedio Declaracion PDT
        
        For i = 1 To feDeclaracionPDT.rows - 1
            nMontoDeclarado = CDbl(Me.feDeclaracionPDT.TextMatrix(i, 4)) + CDbl(Me.feDeclaracionPDT.TextMatrix(i, 5)) + CDbl(Me.feDeclaracionPDT.TextMatrix(i, 6))
            nMontoDeclarado = nMontoDeclarado / 3
            Me.feDeclaracionPDT.TextMatrix(i, 7) = Format(nMontoDeclarado, "#,#0.00")
        Next
            'Para el %Declarado
            If CDbl(feFlujoCajaMensual.TextMatrix(1, 2)) = 0 Then
                MsgBox "Ingrese las Ventas en el Flujo de Caja", vbInformation + vbOKOnly, "Atención"
                Me.feDeclaracionPDT.TextMatrix(feDeclaracionPDT.row, feDeclaracionPDT.Col) = "0.00"
                Exit Sub
            End If
            If CDbl(feFlujoCajaMensual.TextMatrix(3, 2)) = 0 Then
                MsgBox "Ingrese las Compras en el Flujo de Caja", vbInformation + vbOKOnly, "Atención"
                Me.feDeclaracionPDT.TextMatrix(feDeclaracionPDT.row, feDeclaracionPDT.Col) = "0.00"
                Exit Sub
            End If
            
            nPorcentajeVentas = Round(CDbl(feDeclaracionPDT.TextMatrix(1, 7)) / CDbl(feFlujoCajaMensual.TextMatrix(1, 2)), 4)
            nPorcentajeCompras = Round(CDbl(feDeclaracionPDT.TextMatrix(2, 7)) / CDbl(feFlujoCajaMensual.TextMatrix(3, 2)), 4)
    
            Me.feDeclaracionPDT.TextMatrix(1, 8) = CStr(nPorcentajeVentas * 100) & "%"
            Me.feDeclaracionPDT.TextMatrix(2, 8) = CStr(nPorcentajeCompras * 100) & "%"
    End If

    If pnActPas = 5 Then '--flujo de caja
        s1 = 0
        lnVtasContado = 0
        lnCobrosCtaCre = 0
        lnCobrosActFijo = 0
        lnEgrePorCom = 0
        
        lnVtasContado = CCur(Me.feFlujoCajaMensual.TextMatrix(1, 2))
        lnCobrosCtaCre = CCur(Me.feFlujoCajaMensual.TextMatrix(2, 2))
        lnCobrosActFijo = CCur(Me.feFlujoCajaMensual.TextMatrix(3, 2))
        lnEgrePorCom = CCur(Me.feFlujoCajaMensual.TextMatrix(4, 2))
        'margen bruto
        Me.feFlujoCajaMensual.TextMatrix(5, 2) = Format(lnVtasContado + lnCobrosCtaCre + lnCobrosActFijo - lnEgrePorCom, "#,#0.00")
        For i = 7 To 21
            s1 = s1 + CCur(Me.feFlujoCajaMensual.TextMatrix(i, 2))
        Next i
        'otros egresos
        Me.feFlujoCajaMensual.TextMatrix(6, 2) = Format(s1, "#,#0.00")
        'saldo disponible
        Me.feFlujoCajaMensual.TextMatrix(22, 2) = Format(CCur(Me.feFlujoCajaMensual.TextMatrix(5, 2)) - (CCur(Me.feFlujoCajaMensual.TextMatrix(6, 2)) - CCur(Me.feFlujoCajaMensual.TextMatrix(18, 2)) - CCur(Me.feFlujoCajaMensual.TextMatrix(19, 2))), "#,#0.00")
    End If
    If pnActPas = 6 Then '-- flujo de caja Historico 'LUCV20171015, Agregó según ERS0512017
        s1 = 0
        lnVtasContado = 0
        lnCobrosCtaCre = 0
        lnCobrosActFijo = 0
        lnEgrePorCom = 0
        
        lnVtasContado = CCur(Me.feFlujoCajaHistorico.TextMatrix(1, 2))
        lnCobrosCtaCre = CCur(Me.feFlujoCajaHistorico.TextMatrix(2, 2))
        lnCobrosActFijo = CCur(Me.feFlujoCajaHistorico.TextMatrix(3, 2))
        lnEgrePorCom = CCur(Me.feFlujoCajaHistorico.TextMatrix(4, 2))
        'margen bruto
        Me.feFlujoCajaHistorico.TextMatrix(5, 2) = Format(lnVtasContado + lnCobrosCtaCre + lnCobrosActFijo - lnEgrePorCom, "#,#0.00")
        For i = 7 To 21
            s1 = s1 + CCur(Me.feFlujoCajaHistorico.TextMatrix(i, 2))
        Next i
        'otros egresos
        Me.feFlujoCajaHistorico.TextMatrix(6, 2) = Format(s1, "#,#0.00")
        'saldo disponible
        Me.feFlujoCajaHistorico.TextMatrix(22, 2) = Format(CCur(Me.feFlujoCajaHistorico.TextMatrix(5, 2)) - (CCur(Me.feFlujoCajaHistorico.TextMatrix(6, 2)) - CCur(Me.feFlujoCajaHistorico.TextMatrix(18, 2)) - CCur(Me.feFlujoCajaHistorico.TextMatrix(19, 2))), "#,#0.00")
    End If
End Sub

Private Sub CalculaCoeFinan()
    
'------------ calcula coeficiente financiero
'   patri=tot activo17-tot pasivo16
'    2=activo1-pasivo1
'   3=activo1-pasvivo1-activo9
'   4=activo1/pasivo1
'   5=(activo1-activo9-activo8)/pasivo1
'   7=si(activo17<>0,patri/activo17,0)
'   8=si(patri<>0,pasivo16/patri
'   9=si(patri<>0,(pasivo2+pasivo9+pasivo11)/patri)
'   10=si(egp3<>0,(pasivo16/(egp3/(mes*30)*360))))
'   11=si(activo17<>0,(pasivo16/activo17))
'   13=si(activo17<>0,(sgp3/(mes*30)*360/activo17))
'   14=si(egp3<>0,(egp4+egp5)/egp3)
'   15=si(egp9<>0,(egp12/egp9))
'   16=si(egp9<>0,(egp11-egp12)/egp9)
'   17=si(egp1<>0,(activos4*mes*30)/egp1)
'   18=si(egp1<>0,(pasivos5*mes*30)/(egp4+egp5))
'   19=si(egp1<>0,(activos8*mes*30)/(egp4+egp5))
        
'    If Me.mskFecReg = "__/__/____" Then
'        MsgBox "Debe registrar una fecha"
'        Exit Sub
'    End If

    Dim nMes As Integer
    'nMES = Month(Me.mskFecReg)
    nMes = Month(ldFecRegEF)
    
    Dim Activo14, Pasivo14, Activo44, Activo84, Activo94, Activo174 As Double
    Dim Pasivo24, Pasivo54, Pasivo94, Pasivo114, Pasivo164, Pasivo174 As Double
    Dim EstGanPer12, EstGanPer32, EstGanPer42, EstGanPer52, EstGanPer92, EstGanPer112, EstGanPer122 As Double
    Dim EstGanPer192 As Currency
    Dim nMOntoCred As Currency

    nMOntoCred = CCur(Me.txtExposicionCredito2.Text)

    Activo14 = CDbl(Me.feActivos.TextMatrix(1, 4))
    Activo44 = CDbl(Me.feActivos.TextMatrix(4, 4))
    Activo84 = CDbl(Me.feActivos.TextMatrix(8, 4))
    Activo94 = CDbl(Me.feActivos.TextMatrix(9, 4))
    Activo174 = IIf(CDbl(Me.feActivos.TextMatrix(17, 4)) = 0, 1, CDbl(Me.feActivos.TextMatrix(17, 4)))
    
    Pasivo14 = IIf(CDbl(Me.fePasivos.TextMatrix(1, 4)) = 0, 1, CDbl(Me.fePasivos.TextMatrix(1, 4)))
    Pasivo24 = CDbl(Me.fePasivos.TextMatrix(2, 4))
    Pasivo54 = CDbl(Me.fePasivos.TextMatrix(5, 4))
    Pasivo94 = CDbl(Me.fePasivos.TextMatrix(9, 4))
    Pasivo114 = CDbl(Me.fePasivos.TextMatrix(11, 4))
    
    'total pasivo
    'Pasivo164 = CDbl(Me.fePasivos.TextMatrix(16, 4))
    Pasivo164 = CDbl(Me.fePasivos.TextMatrix(23, 4))
    'Pasivo164 = CDbl(Me.fePasivos.TextMatrix(23, 4)) '- lnCompraDeuda - lnMontoAmpliado
    
    'patrimonio
    'Pasivo174 = IIf(CDbl(Me.fePasivos.TextMatrix(17, 4)) = 0, 1, CDbl(Me.fePasivos.TextMatrix(17, 4)))
    Pasivo174 = IIf(CDbl(Me.fePasivos.TextMatrix(24, 4)) = 0, 1, CDbl(Me.fePasivos.TextMatrix(24, 4)))
    
    EstGanPer12 = IIf(CDbl(Me.feEstaGananPerd.TextMatrix(1, 2)) = 0, 1, CDbl(Me.feEstaGananPerd.TextMatrix(1, 2)))
    EstGanPer32 = IIf(CDbl(Me.feEstaGananPerd.TextMatrix(3, 2)) = 0, 1, CDbl(Me.feEstaGananPerd.TextMatrix(3, 2)))
    EstGanPer42 = IIf(CDbl(Me.feEstaGananPerd.TextMatrix(4, 2)) = 0, 1, CDbl(Me.feEstaGananPerd.TextMatrix(4, 2)))
    EstGanPer52 = IIf(CDbl(Me.feEstaGananPerd.TextMatrix(5, 2)) = 0, 1, CDbl(Me.feEstaGananPerd.TextMatrix(5, 2)))
    EstGanPer92 = IIf(CDbl(Me.feEstaGananPerd.TextMatrix(9, 2)) = 0, 1, CDbl(Me.feEstaGananPerd.TextMatrix(9, 2)))
    
    EstGanPer112 = CDbl(Me.feEstaGananPerd.TextMatrix(11, 2))
    EstGanPer122 = CDbl(Me.feEstaGananPerd.TextMatrix(12, 2))
    EstGanPer192 = CDbl(Me.feEstaGananPerd.TextMatrix(19, 2))
    
    
    Me.feCoeFinan.TextMatrix(2, 2) = Format(Activo14 - Pasivo14, "#,#0.00")
    Me.feCoeFinan.TextMatrix(3, 2) = Format(Activo14 - Pasivo14 - Activo94, "#,#0.00")
    Me.feCoeFinan.TextMatrix(4, 2) = Format(Activo14 / Pasivo14, "#,#0.00")
    Me.feCoeFinan.TextMatrix(5, 2) = Format((Activo14 - Activo94 - Activo84) / Pasivo14, "#,#0.00")
    'ENDEUDAMIENTO
    Me.feCoeFinan.TextMatrix(7, 2) = Format(IIf(Activo174 <> 0, Pasivo174 / Activo174, 0) * 100, "#,#0.00") & "%"
    
    'Me.feCoeFinan.TextMatrix(8, 2) = Format(IIf(Activo174 <> 0, Pasivo164 / Pasivo174, 0), "#,#0.00")
    'Me.feCoeFinan.TextMatrix(8, 2) = Format(IIf(Activo174 <> 0, (Pasivo164 + lnMontoSol) / Pasivo174, 0) * 100, "#,#0.00") & "%"
    Me.feCoeFinan.TextMatrix(8, 2) = Format(IIf(Activo174 <> 0, ((Pasivo164 - lnCompraDeuda - lnMontoAmpliado) + lnMontoSol) / Pasivo174, 0) * 100, "#,#0.00") & "%"
    
    Me.feCoeFinan.TextMatrix(9, 2) = Format(IIf(Activo174 <> 0, (Pasivo24 + Pasivo94 + Pasivo114) / Pasivo174, 0) * 100, "#,#0.00") & "%"
    Me.feCoeFinan.TextMatrix(10, 2) = Format(IIf(EstGanPer32 <> 0, Pasivo164 / ((EstGanPer32 / (nMes * 30)) * 360), 0) * 100, "#,#0.00") & "%"
    Me.feCoeFinan.TextMatrix(11, 2) = Format(IIf(Activo174 <> 0, Pasivo164 / Activo174, 0) * 100, "#,#0.00") & "%"
    
    Me.feCoeFinan.TextMatrix(13, 2) = Format(IIf(Activo174 <> 0, ((EstGanPer32 / (nMes * 30)) * 360) / Activo174, 0), "#,#0.00")
    Me.feCoeFinan.TextMatrix(14, 2) = Format(IIf(EstGanPer32 <> 0, (EstGanPer42 + EstGanPer52) / EstGanPer32, 0), "#,#0.00")
    Me.feCoeFinan.TextMatrix(15, 2) = Format(IIf(EstGanPer92 <> 0, Abs(Round((EstGanPer122 / EstGanPer92) * 100, 1)), 0), "#,#0.00") & "%"
    Me.feCoeFinan.TextMatrix(16, 2) = Format(IIf(EstGanPer92 <> 0, Round((EstGanPer112 - (EstGanPer122 * -1)) / EstGanPer92 * 100, 1), 0), "#,#0.00") & "%"
    Me.feCoeFinan.TextMatrix(17, 2) = Format(IIf(EstGanPer12 <> 0, Round((Activo44 * nMes * 30) / EstGanPer12, 0), 0), "#,#0.00")
    Me.feCoeFinan.TextMatrix(18, 2) = Format(IIf(EstGanPer12 <> 0, Round((Pasivo54 * nMes * 30) / (EstGanPer42 + EstGanPer52), 0), 0), "#,#0.00")
    Me.feCoeFinan.TextMatrix(19, 2) = Format(IIf(EstGanPer12 <> 0, Round((Activo84 * nMes * 30) / (EstGanPer42 + EstGanPer52), 0), 0), "#,#0.00")

    Me.feCoeFinan.TextMatrix(21, 2) = Format(IIf(EstGanPer32 <> 0, Round(EstGanPer192 / EstGanPer32 * 100, 2), 0), "#,#0.00") & "%"
    Me.feCoeFinan.TextMatrix(22, 2) = Format(IIf(Pasivo174 <> 0, Round(EstGanPer192 / Pasivo174 * 100, 2), 0), "#,#0.00") & "%"
    Me.feCoeFinan.TextMatrix(23, 2) = Format(IIf(Activo174 <> 0, Round(EstGanPer192 / Activo174 * 100, 2), 0), "#,#0.00") & "%"
    Me.feCoeFinan.TextMatrix(24, 2) = Format(IIf(Activo174 <> 0, Round(EstGanPer92 / Activo174 * 100, 2), 0), "#,#0.00") & "%"

End Sub

Private Sub feCoeFinan_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim Editar() As String
    
    Editar = Split(Me.feCoeFinan.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    Select Case pnRow
        Case 2, 5, 6, 7, 9, 11
            Cancel = False
            SendKeys "{TAB}"
            Exit Sub
    End Select

End Sub

Private Sub feCoeFinan_RowColChange()
    If feCoeFinan.Col = 2 Then
        feCoeFinan.AvanceCeldas = Vertical
    Else
        feCoeFinan.AvanceCeldas = Horizontal
    End If

End Sub

Private Sub feDeclaracionPDT_OnCellChange(pnRow As Long, pnCol As Long)
    Call CalculaCeldas(4)
End Sub

Private Sub feDeclaracionPDT_OnChangeCombo()
    'Call CalculaCeldas(4)
End Sub

Private Sub feEstaGananPerd_EnterCell()
    Select Case CInt(feEstaGananPerd.TextMatrix(Me.feEstaGananPerd.row, 0)) 'celda que  o se puede editar
        Case 3, 6, 9, 16, 19
            Me.feEstaGananPerd.ColumnasAEditar = "X-X-X-X-X"
        Case Else
            Me.feEstaGananPerd.ColumnasAEditar = "X-X-2-X-X"
        End Select
End Sub

Private Sub feEstaGananPerd_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(Me.feEstaGananPerd.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
'        Select Case CInt(feEstaGananPerd.TextMatrix(pnRow, 0))
'            Case 1, 2, 11, 13 'positivos
                If feEstaGananPerd.TextMatrix(pnRow, pnCol) < 0 Then
                    feEstaGananPerd.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(feEstaGananPerd.TextMatrix(pnRow, pnCol))), "#,#0.00")   '"0.00"
                End If
'            Case 4, 5, 7, 8, 12, 14 'negativos
'                If feEstaGananPerd.TextMatrix(pnRow, pnCol) > 0 Then
'                    feEstaGananPerd.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(feEstaGananPerd.TextMatrix(pnRow, pnCol))) * -1, "#,#0.00") '"0.00"
'                End If
'        End Select
    Else
        feEstaGananPerd.TextMatrix(pnRow, pnCol) = "0.00"
    End If
    Call CalculaCeldas(3)
End Sub

Private Sub feEstaGananPerd_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
'    Call CalculaCeldas(3)
End Sub

Private Sub feEstaGananPerd_RowColChange()
    If feEstaGananPerd.Col = 2 Then
        feEstaGananPerd.AvanceCeldas = Vertical
    Else
        feEstaGananPerd.AvanceCeldas = Horizontal
    End If
End Sub
Private Sub feFlujoCajaMensual_EnterCell()
    Select Case CInt(feFlujoCajaMensual.TextMatrix(Me.feFlujoCajaMensual.row, 0)) 'celda que se activa el textbuscar
        Case 19
            Me.feFlujoCajaMensual.ListaControles = "0-0-1-0-0"
        Case Else
            Me.feFlujoCajaMensual.ListaControles = "0-0-0-0-0"
    End Select

    Select Case CInt(feFlujoCajaMensual.TextMatrix(Me.feFlujoCajaMensual.row, 0)) 'celda que  o se puede editar
        Case 5, 6, 18, 22
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X"
        Case Else
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-2-X-X"
    End Select
End Sub

Private Sub feFlujoCajaMensual_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(Me.feFlujoCajaMensual.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feFlujoCajaMensual.TextMatrix(pnRow, pnCol) < 0 Then
            feFlujoCajaMensual.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(feFlujoCajaMensual.TextMatrix(pnRow, pnCol))), "#,#0.00")   '"0.00"
        End If
    Else
        feFlujoCajaMensual.TextMatrix(pnRow, pnCol) = "0.00"
    End If
    Call CalculaCeldas(5)

    If Me.feFlujoCajaMensual.row = 19 And Me.feFlujoCajaMensual.Col = 2 Then
        Me.feGastosFamiliares.SetFocus
        feGastosFamiliares.row = 1
        feGastosFamiliares.Col = 3
        SendKeys "{TAB}"
    End If
End Sub
Private Sub feFlujoCajaMensual_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
    psCodigo = 0
    psDescripcion = ""
    psDescripcion = feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 1) 'Cuotas Otras IFIs
    psCodigo = feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 2) 'Monto
    If psCodigo = 0 Then
        fnTotalRefFlujoCaja = 0
        Set MatIfiGastoNego = Nothing
        frmCredFormEvalCuotasIfis.Inicio (CCur(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 2))), fnTotalRefFlujoCaja, MatIfiGastoNego
        psCodigo = Format(fnTotalRefFlujoCaja, "#,##0.00")
    Else
        frmCredFormEvalCuotasIfis.Inicio (CCur(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 2))), fnTotalRefFlujoCaja, MatIfiGastoNego
        psCodigo = Format(fnTotalRefFlujoCaja, "#,##0.00")
    End If
End Sub
Private Sub feFlujoCajaMensual_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim Editar() As String
    
    Editar = Split(Me.feFlujoCajaMensual.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{TAB}"
        Exit Sub
    End If
            
    Select Case pnRow
        Case 19
            Cancel = False
            SendKeys "{TAB}"
            Exit Sub
    End Select
End Sub
Private Sub feFlujoCajaMensual_RowColChange()
    If feFlujoCajaMensual.Col = 2 Then
        feFlujoCajaMensual.AvanceCeldas = Vertical
    Else
        feFlujoCajaMensual.AvanceCeldas = Horizontal
    End If
End Sub

Private Sub feGastosFamiliares_Click()
    If feGastosFamiliares.Col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami Then
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If

    Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) 'celda que  o se puede editar
        Case gCodCuotaIfiGastoFami
            Me.feGastosFamiliares.CellBackColor = &HC0FFFF
            Me.feGastosFamiliares.ForeColorRow (&H80000007), True
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
        Case gCodDeudaLCNUGastoFami
            Me.feGastosFamiliares.BackColorRow vbWhite
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
        Case Else
            Me.feGastosFamiliares.BackColorRow (&HFFFFFF)
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
    End Select

End Sub

Private Sub feGastosFamiliares_EnterCell()
    If feGastosFamiliares.Col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami Then
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If

    Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) 'celda que  o se puede editar
        Case gCodCuotaIfiGastoFami
            Me.feGastosFamiliares.CellBackColor = &HC0FFFF
            Me.feGastosFamiliares.ForeColorRow (&H80000007), True
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
        Case gCodDeudaLCNUGastoFami
            Me.feGastosFamiliares.BackColorRow vbWhite
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
        Case Else
            Me.feGastosFamiliares.BackColorRow (&HFFFFFF)
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
    End Select
End Sub

Private Sub feGastosFamiliares_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feGastosFamiliares.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feGastosFamiliares.TextMatrix(pnRow, pnCol) < 0 Then
            feGastosFamiliares.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feGastosFamiliares.TextMatrix(pnRow, pnCol) = 0
    End If

    If Me.feGastosFamiliares.row = 6 And Me.feGastosFamiliares.Col = 3 Then
        Me.feOtrosIngresos.SetFocus
        feOtrosIngresos.row = 1
        feOtrosIngresos.Col = 3
        SendKeys "{TAB}"
    End If


End Sub

Private Sub feGastosFamiliares_OnClickTxtBuscar(psMontoIfiGastoFami As String, psDescripcion As String)

    psMontoIfiGastoFami = 0
    psDescripcion = ""
    psDescripcion = feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2) 'Cuotas Otras IFIs
    psMontoIfiGastoFami = feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3) 'Monto
    
    If psMontoIfiGastoFami = 0 Then
        fnTotalRefGastoFami = 0
        Set MatIfiGastoFami = Nothing
        frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiGastoFami
        psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
    Else
        frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiGastoFami
        psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
    End If

End Sub

Private Sub feGastosFamiliares_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    
    Editar = Split(Me.feGastosFamiliares.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    Select Case pnRow
        Case 5
            Cancel = False
            SendKeys "{TAB}"
            Exit Sub
    End Select

End Sub

Private Sub feGastosFamiliares_RowColChange()
    If feGastosFamiliares.Col = 3 Then
        feGastosFamiliares.AvanceCeldas = Vertical
    Else
        feGastosFamiliares.AvanceCeldas = Horizontal
    End If
    
    If feGastosFamiliares.Col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 0)) = gCodCuotaIfiGastoFami Then
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If

    Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) 'celda que  o se puede editar
        Case gCodCuotaIfiGastoFami
            Me.feGastosFamiliares.CellBackColor = &HC0FFFF
            Me.feGastosFamiliares.ForeColorRow (&H80000007), True
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
        Case gCodDeudaLCNUGastoFami
            Me.feGastosFamiliares.BackColorRow vbWhite
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
        Case Else
            Me.feGastosFamiliares.BackColorRow (&HFFFFFF)
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
    End Select

End Sub

Private Sub feOtrosIngresos_RowColChange()
    If feOtrosIngresos.Col = 3 Then
        feOtrosIngresos.AvanceCeldas = Vertical
    Else
        feOtrosIngresos.AvanceCeldas = Horizontal
    End If

End Sub

Private Sub fePasivos_EnterCell()
   
    Select Case CInt(Me.fePasivos.TextMatrix(Me.fePasivos.row, 0)) 'celda que se activa el textbuscar
        Case 2, 5, 6, 7, 9, 11
            Me.fePasivos.ListaControles = "0-0-1-1-0-0-0"
        Case Else
            Me.fePasivos.ListaControles = "0-0-0-0-0-0-0"
        End Select
        
    Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0)) 'celda que  no se puede editar
        Case 1, 10, 16, 17, 23, 24, 25
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case Else
            Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        End Select
        
End Sub

Private Sub fePasivos_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(fePasivos.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        
        Select Case CInt(fePasivos.TextMatrix(pnRow, 0))
            Case 21, 22 ' ingresa positivos o negativos
                'fePasivos.TextMatrix(pnRow, pnCol) = Format(CCur(fePasivos.TextMatrix(pnRow, pnCol)), "#,#0.00")   '"0.00"
            'Case 4, 5, 7, 8, 12, 14 'negativos
            Case Else
                If fePasivos.TextMatrix(pnRow, pnCol) < 0 Then
                    fePasivos.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(fePasivos.TextMatrix(pnRow, pnCol))), "#,#0.00")   '"0.00"
                End If
        End Select
        
'        If fePasivos.TextMatrix(pnRow, pnCol) < 0 Then
'            fePasivos.TextMatrix(pnRow, pnCol) = Format(Abs(val(fePasivos.TextMatrix(pnRow, pnCol))), "#,#0.00")   '"0.00"
'        End If
        
    Else
        fePasivos.TextMatrix(pnRow, pnCol) = "0.00"
    End If
    nTotal = CDbl(fePasivos.TextMatrix(fePasivos.row, fePasivos.Col))
    Call CalculaCeldas(2)
End Sub

Private Sub fePasivos_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)

Dim pnMonto As Double
'Dim lvDetallePasivos() As tForEvalEstFinFormato6 'matriz para activos, Comentó
Dim lvDetallePasivos() As tFormEvalDetalleEstFinFormato6 'LUCV20171015, Agregó

Dim index As Integer
       
    'If Me.mskFecReg = "__/__/____" Then
    If lcFecRegEF = "" Then
        MsgBox "Debe registrar una fecha", vbInformation + vbOKOnly, "Atención"
        Exit Sub
    End If

    If Me.fePasivos.TextMatrix(1, 0) = "" Then Exit Sub
    
    index = CInt(fePasivos.TextMatrix(fePasivos.row, 0))
    nTotal = 0
    Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0))
        Case 2, 5, 6, 7
            Set oFrm6 = New frmCredFormEvalDetalleFormato6
            
            If fePasivos.Col = 2 Then 'column P.P.
                If IsArray(lvPrincipalPasivos(index).vPP) Then
                    lvDetallePasivos = lvPrincipalPasivos(index).vPP
                    nTotal = lvPrincipalPasivos(index).nImportePP
                Else
                    ReDim lvDetallePasivos(0)
                End If
            End If
            
            If fePasivos.Col = 3 Then 'column P.E.
                If IsArray(lvPrincipalPasivos(index).vPE) Then
                    lvDetallePasivos = lvPrincipalPasivos(index).vPE
                    nTotal = lvPrincipalPasivos(index).nImportePE
                Else
                    ReDim lvDetallePasivos(0)
                End If
            End If
                        
            If oFrm6.Registrar(True, 1, lvPrincipalPasivos(index).cConcepto, lvDetallePasivos, lvDetallePasivos, nTotal, CInt(fePasivos.TextMatrix(Me.fePasivos.row, 5)), CInt(fePasivos.TextMatrix(Me.fePasivos.row, 6)), Trim(str(ldFecRegEF))) Then
                If fePasivos.Col = 2 Then 'column P.P.
                    lvPrincipalPasivos(index).vPP = lvDetallePasivos
                End If
                If fePasivos.Col = 3 Then ' columna P.E.
                    lvPrincipalPasivos(index).vPE = lvDetallePasivos
                End If
                
            End If
            
            If fePasivos.Col = 2 Then
                Me.fePasivos.TextMatrix(Me.fePasivos.row, Me.fePasivos.Col) = Format(nTotal, "#,#0.00")
                If nTotal <> 0 Then
                    lvPrincipalPasivos(index).nImportePP = nTotal
                    lvPrincipalPasivos(index).nDetPP = 1
                Else
                    lvPrincipalPasivos(index).nImportePP = nTotal
                    lvPrincipalPasivos(index).nDetPP = 0
                End If
                Call CalculaCeldas(2)
            End If
            
            If fePasivos.Col = 3 Then
                Me.fePasivos.TextMatrix(Me.fePasivos.row, Me.fePasivos.Col) = Format(nTotal, "#,#0.00")
                If nTotal <> 0 Then
                    lvPrincipalPasivos(index).nImportePE = nTotal
                    lvPrincipalPasivos(index).nDetPE = 1
                Else
                    lvPrincipalPasivos(index).nImportePE = nTotal
                    lvPrincipalPasivos(index).nDetPE = 0
                End If
                Call CalculaCeldas(2)
            End If

            Call CalculaCeldas(2)
            Call CalculaCeldas(2)
            
        Case 9, 11 'detalle de Ifis
            
            'Set oFrm6 = New frmCredFormEvalDetalleFormato6
            
            If fePasivos.Col = 2 Then 'column P.P.
                If IsArray(lvPrincipalPasivos(index).vPP) Then
                    lvDetallePasivos = lvPrincipalPasivos(index).vPP
                    nTotal = lvPrincipalPasivos(index).nImportePP
                Else
                    ReDim lvDetallePasivos(0)
                End If
            End If
            
            If fePasivos.Col = 3 Then 'column P.E.
                If IsArray(lvPrincipalPasivos(index).vPE) Then
                    lvDetallePasivos = lvPrincipalPasivos(index).vPE
                    nTotal = lvPrincipalPasivos(index).nImportePE
                Else
                    ReDim lvDetallePasivos(0)
                End If
            End If
                        
            If frmCredFormEvalIfisDetalleFormato6.Registrar(True, 1, lvPrincipalPasivos(index).cConcepto, lvDetallePasivos, lvDetallePasivos, nTotal, CInt(fePasivos.TextMatrix(Me.fePasivos.row, 5)), CInt(fePasivos.TextMatrix(Me.fePasivos.row, 6)), Trim(str(ldFecRegEF))) Then
                If fePasivos.Col = 2 Then 'column P.P.
                    lvPrincipalPasivos(index).vPP = lvDetallePasivos
                End If
                If fePasivos.Col = 3 Then ' columna P.E.
                    lvPrincipalPasivos(index).vPE = lvDetallePasivos
                End If
                
            End If
            
            If fePasivos.Col = 2 Then
                Me.fePasivos.TextMatrix(Me.fePasivos.row, Me.fePasivos.Col) = Format(nTotal, "#,#0.00")
                If nTotal <> 0 Then
                    lvPrincipalPasivos(index).nImportePP = nTotal
                    lvPrincipalPasivos(index).nDetPP = 1
                Else
                    lvPrincipalPasivos(index).nImportePP = nTotal
                    lvPrincipalPasivos(index).nDetPP = 0
                End If
                Call CalculaCeldas(2)
            End If
            
            If fePasivos.Col = 3 Then
                Me.fePasivos.TextMatrix(Me.fePasivos.row, Me.fePasivos.Col) = Format(nTotal, "#,#0.00")
                If nTotal <> 0 Then
                    lvPrincipalPasivos(index).nImportePE = nTotal
                    lvPrincipalPasivos(index).nDetPE = 1
                Else
                    lvPrincipalPasivos(index).nImportePE = nTotal
                    lvPrincipalPasivos(index).nDetPE = 0
                End If
                Call CalculaCeldas(2)
            End If

            Call CalculaCeldas(2)
            Call CalculaCeldas(2)
            
        End Select
End Sub

Private Sub fePasivos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    
    Editar = Split(Me.fePasivos.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    Select Case pnRow
        Case 2, 5, 6, 7, 9, 11
            Cancel = False
            SendKeys "{TAB}"
            Exit Sub
    End Select
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        KeyCode = 0
    End If
End Sub
Private Sub Form_Load()
    DisableCloseButton Me
'JOEP20180725 ERS034-2018
    If fnTipoRegMant = 3 Then
        If Not ConsultaRiesgoCamCred(fsCtaCod) Then
            cmdMNME.Visible = True
        End If
    End If
'JOEP20180725 ERS034-2018
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set MatIfiGastoNego = Nothing 'LUCV20161115
    Set MatIfiGastoFami = Nothing 'LUCV20161115
End Sub
Private Sub mskFecReg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim oNCred As COMDCredito.DCOMFormatosEval
        Set oNCred = New COMDCredito.DCOMFormatosEval
            
        If IsDate(mskFecReg.Text) Then
            'busca fecha
            If oNCred.ValidaFechaCredFormEval(fsCtaCod, 6, Format(mskFecReg, "yyyyMMdd")).RecordCount = 0 Then
                MsgBox "No existe datos en la fecha ingresada.", vbOKOnly + vbInformation, "Atención"
    '            Exit Sub
            Else
                'carga datos
                Me.feActivos.Enabled = True
                Me.fePasivos.Enabled = True
                Me.feEstaGananPerd.Enabled = True
                Me.feFlujoCajaMensual.Enabled = True
                Me.feGastosFamiliares.Enabled = True
                Me.feOtrosIngresos.Enabled = True
                Me.feDeclaracionPDT.Enabled = True
                Me.feFlujoCajaHistorico.Enabled = True 'LUCV20171015, Agregó según ERS0512017
                
                Call CargaDatosBusqueda(fsCtaCod, 6, Format(mskFecReg, "yyyyMMdd"))
            End If
        End If
    End If
End Sub

Private Sub CargaDatosBusqueda(ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pcFecReg As String)

    Dim nFila As Integer
    Dim NumRegRS As Integer
    Dim lnFila As Integer
    Dim nFilaDet As Integer
    Dim lcTitOriginal As String
    Dim i As Integer
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    '----------------------------------------- LLENAR ACTIVOS
    Set prsDatActivosForm6 = oDCOMFormatosEval.ObtenerFormatosEvalActivos(psCtaCod, pnFormato, pcFecReg)
    Set prsDatActivosForm6Det = oDCOMFormatosEval.ObtenerFormatosEvalActivosDet(psCtaCod, pnFormato, pcFecReg)
    
    feActivos.Clear
    feActivos.FormaCabecera
    feActivos.rows = 2
        Call LimpiaFlex(feActivos)
        nFila = 0
        NumRegRS = prsDatActivosForm6.RecordCount
        ReDim lvPrincipalActivos(NumRegRS)

    If Not (prsDatActivosForm6.EOF And prsDatActivosForm6.BOF) Then
        prsDatActivosForm6.MoveFirst
        If prsDatActivosForm6!nAuditado = True Then
            Me.chkAudit.value = 1
        Else
            Me.chkAudit.value = 0
        End If
    End If
    
    Do While Not (prsDatActivosForm6.EOF)
        feActivos.AdicionaFila
        lnFila = feActivos.row
        feActivos.TextMatrix(lnFila, 1) = prsDatActivosForm6!Concepto
        feActivos.TextMatrix(lnFila, 2) = Format(prsDatActivosForm6!PP, "#,#0.00")
        feActivos.TextMatrix(lnFila, 3) = Format(prsDatActivosForm6!PE, "#,#0.00")
        feActivos.TextMatrix(lnFila, 4) = Format(prsDatActivosForm6!Total, "#,#0.00")
        feActivos.TextMatrix(lnFila, 5) = prsDatActivosForm6!nConsCod
        feActivos.TextMatrix(lnFila, 6) = prsDatActivosForm6!nConsValor
                
        '----------------- llena matriz activos
        nFila = nFila + 1
        lvPrincipalActivos(nFila).cConcepto = prsDatActivosForm6!Concepto
        lvPrincipalActivos(nFila).nImportePP = prsDatActivosForm6!PP
        lvPrincipalActivos(nFila).nImportePE = prsDatActivosForm6!PE
        lvPrincipalActivos(nFila).nConsCod = prsDatActivosForm6!nConsCod
        lvPrincipalActivos(nFila).nConsValor = prsDatActivosForm6!nConsValor
        lvPrincipalActivos(nFila).nDetPP = prsDatActivosForm6!nDetPP
        lvPrincipalActivos(nFila).nDetPE = prsDatActivosForm6!nDetPE

        If prsDatActivosForm6!nDetPP <> 0 Then
            nFilaDet = 0

            'lvPrincipalActivos(nFila).nDetPP = 1
            ReDim lvPrincipalActivos(nFila).vPP(prsDatActivosForm6!nDetPP)
            prsDatActivosForm6Det.MoveFirst
            Do While Not (prsDatActivosForm6Det.EOF)
            
                If prsDatActivosForm6Det!nConsCod = prsDatActivosForm6!nConsCod And prsDatActivosForm6Det!nConsValor = prsDatActivosForm6!nConsValor _
                        And prsDatActivosForm6Det!nTipoPatri = 1 Then
                    nFilaDet = nFilaDet + 1
                    lvPrincipalActivos(nFila).vPP(nFilaDet).cDescripcion = prsDatActivosForm6Det!cDescripcionDet
                    lvPrincipalActivos(nFila).vPP(nFilaDet).dFecha = prsDatActivosForm6Det!dFechaDet
                    lvPrincipalActivos(nFila).vPP(nFilaDet).nImporte = prsDatActivosForm6Det!PP
                End If
                prsDatActivosForm6Det.MoveNext
            Loop
        End If
        
        If prsDatActivosForm6!nDetPE <> 0 Then
            nFilaDet = 0
            ReDim lvPrincipalActivos(nFila).vPE(prsDatActivosForm6!nDetPE)

            prsDatActivosForm6Det.MoveFirst
            Do While Not (prsDatActivosForm6Det.EOF)
            
                If prsDatActivosForm6Det!nConsCod = prsDatActivosForm6!nConsCod And prsDatActivosForm6Det!nConsValor = prsDatActivosForm6!nConsValor _
                    And prsDatActivosForm6Det!nTipoPatri = 2 Then
                    
                    nFilaDet = nFilaDet + 1
                    
                    lvPrincipalActivos(nFila).vPE(nFilaDet).cDescripcion = prsDatActivosForm6Det!cDescripcionDet
                    lvPrincipalActivos(nFila).vPE(nFilaDet).dFecha = prsDatActivosForm6Det!dFechaDet
                    lvPrincipalActivos(nFila).vPE(nFilaDet).nImporte = prsDatActivosForm6Det!PE
                End If
                    
                prsDatActivosForm6Det.MoveNext
            Loop
            
        End If
        
        Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0))
            Case 2, 4, 5, 6, 7, 8, 9, 11, 14, 16
                 Me.feActivos.BackColorRow &HC0FFFF, True 'color amarillo claro
        End Select
        
        Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0))
            Case 1, 10, 17
                 Me.feActivos.BackColorRow QBColor(8), True 'gris
        End Select
        
        prsDatActivosForm6.MoveNext
    Loop
    prsDatActivosForm6.Close
    prsDatActivosForm6Det.Close
    Set prsDatActivosForm6 = Nothing
    Set prsDatActivosForm6Det = Nothing

'------------------------------------------------- LLENAR PASIVOS
    Set prsDatPasivosForm6 = oDCOMFormatosEval.ObtenerFormatosEvalPasivos(psCtaCod, pnFormato, pcFecReg)
    Set prsDatPasivosForm6det = oDCOMFormatosEval.ObtenerFormatosEvalPasivosDet(psCtaCod, pnFormato, pcFecReg)

    Me.fePasivos.Clear
    fePasivos.FormaCabecera
    fePasivos.rows = 2
        Call LimpiaFlex(fePasivos)
        nFila = 0
        NumRegRS = prsDatPasivosForm6.RecordCount
        ReDim lvPrincipalPasivos(NumRegRS)
        
    Do While Not (prsDatPasivosForm6.EOF)
        fePasivos.AdicionaFila
        lnFila = fePasivos.row
        fePasivos.TextMatrix(lnFila, 1) = prsDatPasivosForm6!Concepto
        fePasivos.TextMatrix(lnFila, 2) = Format(prsDatPasivosForm6!PP, "#,#0.00")
        fePasivos.TextMatrix(lnFila, 3) = Format(prsDatPasivosForm6!PE, "#,#0.00")
        fePasivos.TextMatrix(lnFila, 4) = Format(prsDatPasivosForm6!Total, "#,#0.00")
        fePasivos.TextMatrix(lnFila, 5) = prsDatPasivosForm6!nConsCod
        fePasivos.TextMatrix(lnFila, 6) = prsDatPasivosForm6!nConsValor

        '----------------- llena matriz pasivos
        nFila = nFila + 1
        lvPrincipalPasivos(nFila).cConcepto = prsDatPasivosForm6!Concepto
        lvPrincipalPasivos(nFila).nImportePP = prsDatPasivosForm6!PP
        lvPrincipalPasivos(nFila).nImportePE = prsDatPasivosForm6!PE
        lvPrincipalPasivos(nFila).nConsCod = prsDatPasivosForm6!nConsCod
        lvPrincipalPasivos(nFila).nConsValor = prsDatPasivosForm6!nConsValor
        lvPrincipalPasivos(nFila).nDetPP = prsDatPasivosForm6!nDetPP
        lvPrincipalPasivos(nFila).nDetPE = prsDatPasivosForm6!nDetPE

        If prsDatPasivosForm6!nDetPP <> 0 Then
            nFilaDet = 0
            
            ReDim lvPrincipalPasivos(nFila).vPP(prsDatPasivosForm6!nDetPP)
            prsDatPasivosForm6det.MoveFirst
            Do While Not (prsDatPasivosForm6det.EOF)
            
                If prsDatPasivosForm6det!nConsCod = prsDatPasivosForm6!nConsCod And prsDatPasivosForm6det!nConsValor = prsDatPasivosForm6!nConsValor _
                        And prsDatPasivosForm6det!nTipoPatri = 1 Then
                        
                    nFilaDet = nFilaDet + 1
                    
                    lcDescripcionDet = IIf(prsDatPasivosForm6det!nConsValor = 109 Or prsDatPasivosForm6det!nConsValor = 201, Trim(prsDatPasivosForm6det!cDescripcionDet) & Space(150) & prsDatPasivosForm6det!cCodIfi, Trim(prsDatPasivosForm6det!cDescripcionDet))
                    
                    lvPrincipalPasivos(nFila).vPP(nFilaDet).cDescripcion = lcDescripcionDet
                    lvPrincipalPasivos(nFila).vPP(nFilaDet).dFecha = prsDatPasivosForm6det!dFechaDet
                    lvPrincipalPasivos(nFila).vPP(nFilaDet).nImporte = prsDatPasivosForm6det!PP
    
                End If
                    
                prsDatPasivosForm6det.MoveNext
            Loop
            
        End If
        
        If prsDatPasivosForm6!nDetPE <> 0 Then
            nFilaDet = 0
            ReDim lvPrincipalPasivos(nFila).vPE(prsDatPasivosForm6!nDetPE)
            
            prsDatPasivosForm6det.MoveFirst
            Do While Not (prsDatPasivosForm6det.EOF)
            
                If prsDatPasivosForm6det!nConsCod = prsDatPasivosForm6!nConsCod And prsDatPasivosForm6det!nConsValor = prsDatPasivosForm6!nConsValor _
                    And prsDatPasivosForm6det!nTipoPatri = 2 Then
                    
                    nFilaDet = nFilaDet + 1
                    
                    lcDescripcionDet = IIf(prsDatPasivosForm6det!nConsValor = 109 Or prsDatPasivosForm6det!nConsValor = 201, Trim(prsDatPasivosForm6det!cDescripcionDet) & Space(150) & prsDatPasivosForm6det!cCodIfi, Trim(prsDatPasivosForm6det!cDescripcionDet))
                    
                    lvPrincipalPasivos(nFila).vPE(nFilaDet).cDescripcion = lcDescripcionDet
                    lvPrincipalPasivos(nFila).vPE(nFilaDet).dFecha = prsDatPasivosForm6det!dFechaDet
                    lvPrincipalPasivos(nFila).vPE(nFilaDet).nImporte = prsDatPasivosForm6det!PE
    
                End If
                    
                prsDatPasivosForm6det.MoveNext
            Loop
        End If
        
        Select Case CInt(Me.fePasivos.TextMatrix(Me.fePasivos.row, 0))
            Case 2, 5, 6, 7, 9, 11
                 Me.fePasivos.BackColorRow &HC0FFFF, True 'color amarillo claro
        End Select
        
        Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0))
            Case 1, 10, 16, 23, 24, 25
                 Me.fePasivos.BackColorRow QBColor(8), True 'gris
        End Select
        prsDatPasivosForm6.MoveNext
    Loop
    prsDatPasivosForm6.Close
    prsDatPasivosForm6det.Close
    Set prsDatPasivosForm6 = Nothing
    Set prsDatPasivosForm6det = Nothing


'------------------------------------------------- LLENAR ESTADOS DE GANANCIAS Y PERDIDAS
    Set prsDatEstadoGananPerdForm6 = oDCOMFormatosEval.ObtenerFormatosEvalEstadoGanPerd(psCtaCod, pnFormato, pcFecReg)

    Me.feEstaGananPerd.Clear
    feEstaGananPerd.FormaCabecera
    feEstaGananPerd.rows = 2
        Call LimpiaFlex(feEstaGananPerd)
        nFila = 0
        NumRegRS = prsDatEstadoGananPerdForm6.RecordCount
        ReDim lvPrincipalEstGanPer(NumRegRS)

    Do While Not (prsDatEstadoGananPerdForm6.EOF)
        feEstaGananPerd.AdicionaFila
        lnFila = feEstaGananPerd.row
        feEstaGananPerd.TextMatrix(lnFila, 1) = prsDatEstadoGananPerdForm6!Concepto
        feEstaGananPerd.TextMatrix(lnFila, 2) = Format(prsDatEstadoGananPerdForm6!nMonto, "#,#0.00")
        feEstaGananPerd.TextMatrix(lnFila, 3) = prsDatEstadoGananPerdForm6!nConsCod
        feEstaGananPerd.TextMatrix(lnFila, 4) = prsDatEstadoGananPerdForm6!nConsValor
        
        Select Case CInt(feEstaGananPerd.TextMatrix(Me.feEstaGananPerd.row, 0))
            Case 3, 6, 9, 16, 19
                'Me.feActivos.CellForeColor() = QBColor(1) 'color azul
                'Me.feEstaGananPerd.CellBackColor() = QBColor(8) 'gris
                Me.feEstaGananPerd.BackColorRow QBColor(8), True 'gris
        End Select
        
        prsDatEstadoGananPerdForm6.MoveNext
    Loop
    prsDatEstadoGananPerdForm6.Close
    Set prsDatEstadoGananPerdForm6 = Nothing
    
'------------------------------------------------- LLENAR COEFICIENTE FINANCIERO
    Set prsDatCoeFinanForm6 = oDCOMFormatosEval.ObtenerFormatosEvalCoeficienteFinan(psCtaCod, pnFormato, pcFecReg)

    Me.feCoeFinan.Clear
    feCoeFinan.FormaCabecera
    feCoeFinan.rows = 2
        Call LimpiaFlex(feCoeFinan)
        nFila = 0

    Do While Not (prsDatCoeFinanForm6.EOF)
        feCoeFinan.AdicionaFila
        lnFila = feCoeFinan.row
        feCoeFinan.TextMatrix(lnFila, 1) = prsDatCoeFinanForm6!Concepto
        feCoeFinan.TextMatrix(lnFila, 2) = Format(prsDatCoeFinanForm6!nMonto, "#,#0.00")
        feCoeFinan.TextMatrix(lnFila, 3) = prsDatCoeFinanForm6!nConsCod
        feCoeFinan.TextMatrix(lnFila, 4) = prsDatCoeFinanForm6!nConsValor
        
        Select Case CInt(feCoeFinan.TextMatrix(Me.feCoeFinan.row, 0))
            Case 1, 6, 12, 20
                Me.feCoeFinan.BackColorRow QBColor(8), True 'gris
        End Select
        
        prsDatCoeFinanForm6.MoveNext
    Loop
    prsDatCoeFinanForm6.Close
    Set prsDatCoeFinanForm6 = Nothing

'    Call CalculaCoeFinan
    
'-------------------------------- LLENAR flujo de caja
    'Set rsDatFlujoCaja = oDCOMFormatosEval.RecuperaDatosCredEvalFlujoCajaForm6(psCtaCod)
    Set rsDatFlujoCaja = oDCOMFormatosEval.RecuperaDatosFlexFlujoCaja(pnFormato, psCtaCod)
    Set rsDatIfiflujocaja = oDCOMFormatosEval.RecuperaDatosIfiCuota(psCtaCod, pnFormato, 7027)
    
    If rsDatFlujoCaja.RecordCount <> 0 Then
        Me.feFlujoCajaMensual.Clear
        feFlujoCajaMensual.FormaCabecera
        feFlujoCajaMensual.rows = 2
        Call LimpiaFlex(feFlujoCajaMensual)
        nFila = 0
            Do While Not (rsDatFlujoCaja.EOF)
                feFlujoCajaMensual.AdicionaFila
                lnFila = feFlujoCajaMensual.row
                feFlujoCajaMensual.TextMatrix(lnFila, 1) = rsDatFlujoCaja!cConsDescripcion
                feFlujoCajaMensual.TextMatrix(lnFila, 2) = Format(rsDatFlujoCaja!nMonto, "#,##0.00")
                feFlujoCajaMensual.TextMatrix(lnFila, 3) = rsDatFlujoCaja!nConsCod
                feFlujoCajaMensual.TextMatrix(lnFila, 4) = rsDatFlujoCaja!nConsValor
                
                Select Case CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 0))
                    Case 5, 6, 18, 22
                         Me.feFlujoCajaMensual.BackColorRow QBColor(8), True 'gris
                    Case 19
                         Me.feFlujoCajaMensual.BackColorRow &HC0FFFF, True 'color amarillo claro
                End Select
                rsDatFlujoCaja.MoveNext
            Loop
            
            Call LimpiaFlex(frmCredFormEvalCuotasIfis.feCuotaIfis)
            Do While Not (rsDatIfiflujocaja.EOF)
                frmCredFormEvalCuotasIfis.feCuotaIfis.AdicionaFila
                lnFila = frmCredFormEvalCuotasIfis.feCuotaIfis.row
                frmCredFormEvalCuotasIfis.feCuotaIfis.TextMatrix(lnFila, 1) = rsDatIfiflujocaja!cDescripcion
                frmCredFormEvalCuotasIfis.feCuotaIfis.TextMatrix(lnFila, 2) = Format(rsDatIfiflujocaja!nMonto, "#,##0.00")
                rsDatIfiflujocaja.MoveNext
            Loop
    
            If Not (rsDatIfiflujocaja.BOF And rsDatIfiflujocaja.EOF) Then
                '--- IFIS DE FLUJO DE CAJA
                ReDim MatIfiGastoNego(rsDatIfiflujocaja.RecordCount, 4)
                i = 0
                rsDatIfiflujocaja.MoveFirst
                Do While Not rsDatIfiflujocaja.EOF
                    MatIfiGastoNego(i, 0) = rsDatIfiflujocaja!nNroCuota
                    MatIfiGastoNego(i, 1) = rsDatIfiflujocaja!cDescripcion
                    MatIfiGastoNego(i, 2) = Format(IIf(IsNull(rsDatIfiflujocaja!nMonto), 0, rsDatIfiflujocaja!nMonto), "#,##0.00")
                    rsDatIfiflujocaja.MoveNext
                      i = i + 1
                Loop
            End If
        End If
        rsDatFlujoCaja.Close
        Set rsDatFlujoCaja = Nothing
        rsDatIfiflujocaja.Close
        Set rsDatIfiflujocaja = Nothing
        
    '->***** LUCV20171015, Acregó segun ERS0512017 (LLENAR FLUJO DE CAJA HISTORICO )->*****
      'Set rsDatFlujoCaja = oDCOMFormatosEval.RecuperaDatosCredEvalFlujoCajaForm6(psCtaCod)
    Set rsDatFlujoCajaHistorico = oDCOMFormatosEval.RecuperaDatosFlexFlujoCaja(pnFormato, psCtaCod, 1)
    Set rsDatIfiflujocajaHistorico = oDCOMFormatosEval.RecuperaDatosIfiCuota(psCtaCod, pnFormato, 7027, , 1)
    
    If rsDatFlujoCajaHistorico.RecordCount <> 0 Then
        Me.feFlujoCajaHistorico.Clear
        feFlujoCajaHistorico.FormaCabecera
        feFlujoCajaHistorico.rows = 2
        Call LimpiaFlex(feFlujoCajaHistorico)
        nFila = 0
            Do While Not (rsDatFlujoCajaHistorico.EOF)
                feFlujoCajaHistorico.AdicionaFila
                lnFila = feFlujoCajaHistorico.row
                feFlujoCajaHistorico.TextMatrix(lnFila, 1) = rsDatFlujoCajaHistorico!cConsDescripcion
                feFlujoCajaHistorico.TextMatrix(lnFila, 2) = Format(rsDatFlujoCajaHistorico!nMonto, "#,##0.00")
                feFlujoCajaHistorico.TextMatrix(lnFila, 3) = rsDatFlujoCajaHistorico!nConsCod
                feFlujoCajaHistorico.TextMatrix(lnFila, 4) = rsDatFlujoCajaHistorico!nConsValor
                
                Select Case CInt(feFlujoCajaHistorico.TextMatrix(feFlujoCajaHistorico.row, 0))
                    Case 1, 4
                        Me.feFlujoCajaHistorico.ForeColorRow vbGrayText
                    Case 5, 6, 22
                         Me.feFlujoCajaHistorico.BackColorRow QBColor(8), True 'gris
                    Case 19
                         Me.feFlujoCajaHistorico.BackColorRow &HC0FFFF, True 'color amarillo claro
                End Select
                rsDatFlujoCajaHistorico.MoveNext
            Loop
            
'            Call LimpiaFlex(frmCredFormEvalCuotasIfis.feCuotaIfis)
'            Do While Not (rsDatIfiflujocaja.EOF)
'                frmCredFormEvalCuotasIfis.feCuotaIfis.AdicionaFila
'                lnFila = frmCredFormEvalCuotasIfis.feCuotaIfis.row
'                frmCredFormEvalCuotasIfis.feCuotaIfis.TextMatrix(lnFila, 1) = rsDatIfiflujocaja!cDescripcion
'                frmCredFormEvalCuotasIfis.feCuotaIfis.TextMatrix(lnFila, 2) = Format(rsDatIfiflujocaja!nMonto, "#,##0.00")
'                rsDatIfiflujocaja.MoveNext
'            Loop
    
            If Not (rsDatIfiflujocajaHistorico.BOF And rsDatIfiflujocajaHistorico.EOF) Then
                '--- IFIS DE FLUJO DE CAJA
                ReDim MatIfiFlujoCajaHistorico(rsDatIfiflujocajaHistorico.RecordCount, 4)
                i = 0
                rsDatIfiflujocajaHistorico.MoveFirst
                Do While Not rsDatIfiflujocajaHistorico.EOF
                    MatIfiFlujoCajaHistorico(i, 0) = rsDatIfiflujocajaHistorico!nNroCuota
                    MatIfiFlujoCajaHistorico(i, 1) = rsDatIfiflujocajaHistorico!cDescripcion
                    MatIfiFlujoCajaHistorico(i, 2) = Format(IIf(IsNull(rsDatIfiflujocajaHistorico!nMonto), 0, rsDatIfiflujocajaHistorico!nMonto), "#,##0.00")
                    rsDatIfiflujocajaHistorico.MoveNext
                      i = i + 1
                Loop
            End If
        End If
        rsDatFlujoCajaHistorico.Close
        Set rsDatFlujoCajaHistorico = Nothing
        rsDatIfiflujocajaHistorico.Close
        Set rsDatIfiflujocajaHistorico = Nothing
        '<-***** LUCV20171015 <-*****
        
    '------------------------------------------------- LLENAR gastos familiares
        Set rsDatGastoFam = oDCOMFormatosEval.RecuperaDatosCredEvalGastosFamForm6(psCtaCod)
        Set rsDatIfiGastoFami = oDCOMFormatosEval.RecuperaDatosIfiCuota(psCtaCod, pnFormato, gFormatoGastosFami)
        
        If rsDatGastoFam.RecordCount > 0 Then
            Call LimpiaFlex(feGastosFamiliares)
            rsDatGastoFam.MoveFirst
            Do While Not (rsDatGastoFam.EOF)
                feGastosFamiliares.AdicionaFila
                lnFila = feGastosFamiliares.row
                feGastosFamiliares.TextMatrix(lnFila, 1) = rsDatGastoFam!nConsValor
                feGastosFamiliares.TextMatrix(lnFila, 2) = rsDatGastoFam!cConsDescripcion
                feGastosFamiliares.TextMatrix(lnFila, 3) = Format(rsDatGastoFam!nMonto, "#,##0.00")
    
            Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) 'celda que  o se puede editar
                Case gCodCuotaIfiGastoFami
                    Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
                    Me.feGastosFamiliares.ForeColorRow (&H80000007), True
                    Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                Case gCodDeudaLCNUGastoFami
                    Me.feGastosFamiliares.BackColorRow vbWhite
                    Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
                Case Else
                    Me.feGastosFamiliares.BackColorRow &HFFFFFF
                    Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
            End Select
            rsDatGastoFam.MoveNext
            Loop
            
            Call LimpiaFlex(frmCredFormEvalCuotasIfis.feCuotaIfis)
                Do While Not (rsDatIfiGastoFami.EOF)
                    frmCredFormEvalCuotasIfis.feCuotaIfis.AdicionaFila
                    lnFila = frmCredFormEvalCuotasIfis.feCuotaIfis.row
                    frmCredFormEvalCuotasIfis.feCuotaIfis.TextMatrix(lnFila, 1) = rsDatIfiGastoFami!cDescripcion
                    frmCredFormEvalCuotasIfis.feCuotaIfis.TextMatrix(lnFila, 2) = Format(rsDatIfiGastoFami!nMonto, "#,##0.00")
                    rsDatIfiGastoFami.MoveNext
                Loop
                        
            If Not (rsDatIfiGastoFami.EOF And rsDatIfiGastoFami.BOF) Then
            '--- IFIS DE GASTOS FAMILIARES
            ReDim MatIfiGastoFami(rsDatIfiGastoFami.RecordCount, 4)
            j = 0
            rsDatIfiGastoFami.MoveFirst
            Do While Not rsDatIfiGastoFami.EOF
                MatIfiGastoFami(j, 0) = rsDatIfiGastoFami!nNroCuota
                MatIfiGastoFami(j, 1) = rsDatIfiGastoFami!cDescripcion
                MatIfiGastoFami(j, 2) = Format(IIf(IsNull(rsDatIfiGastoFami!nMonto), 0, rsDatIfiGastoFami!nMonto), "#,##0.00")
                rsDatIfiGastoFami.MoveNext
                j = j + 1
            Loop
            End If

    End If
    rsDatGastoFam.Close
    Set rsDatGastoFam = Nothing
    rsDatIfiGastoFami.Close
    Set rsDatIfiGastoFami = Nothing
        
'------------------------------------------------- LLENAR otros ingresos
    Set rsDatOtrosIng = oDCOMFormatosEval.RecuperaDatosCredEvalOtrosIngresosForm6(psCtaCod)
    If rsDatOtrosIng.RecordCount > 0 Then
        Call LimpiaFlex(feOtrosIngresos)
            Do While Not (rsDatOtrosIng.EOF)
                feOtrosIngresos.AdicionaFila
                lnFila = feOtrosIngresos.row
                feOtrosIngresos.TextMatrix(lnFila, 1) = rsDatOtrosIng!nConsValor
                feOtrosIngresos.TextMatrix(lnFila, 2) = rsDatOtrosIng!cConsDescripcion
                feOtrosIngresos.TextMatrix(lnFila, 3) = Format(rsDatOtrosIng!nMonto, "#,##0.00")
                rsDatOtrosIng.MoveNext
            Loop
    End If
    rsDatOtrosIng.Close
    Set rsDatOtrosIng = Nothing
    
'------------------------------------------------- LLENAR declara PDT
    Set rsDatPDT = oDCOMFormatosEval.RecuperaDatosCredEvalPDT(psCtaCod)
    Set rsDatPDTDet = oDCOMFormatosEval.RecuperaDatosCredEvalPDTDet(psCtaCod)

    If rsDatPDTDet.RecordCount > 0 Then
        lnFila = 1
        Do While Not (rsDatPDTDet.EOF)
            'feDeclaracionPDT.AdicionaFila
            feDeclaracionPDT.TextMatrix(lnFila, 2) = Format(rsDatPDTDet!nConsCod, "#,##0.00")
            feDeclaracionPDT.TextMatrix(lnFila, 3) = Format(rsDatPDTDet!nConsValor, "#,##0.00")
            feDeclaracionPDT.TextMatrix(lnFila, 4) = Format(rsDatPDTDet!nMontoMes1, "#,##0.00")
            feDeclaracionPDT.TextMatrix(lnFila, 5) = Format(rsDatPDTDet!nMontoMes2, "#,##0.00")
            feDeclaracionPDT.TextMatrix(lnFila, 6) = Format(rsDatPDTDet!nMontoMes3, "#,##0.00")
            feDeclaracionPDT.TextMatrix(lnFila, 7) = Format(rsDatPDTDet!nPromedio, "#,##0.00")
            feDeclaracionPDT.TextMatrix(lnFila, 8) = Format(rsDatPDTDet!nPorcentajeVent, "#,##0.00") & "%"
            rsDatPDTDet.MoveNext
            lnFila = lnFila + 1
        Loop
        
        Do While Not (rsDatPDT.EOF)
            feDeclaracionPDT.TextMatrix(0, 4) = DevolverNombreMes(CInt(rsDatPDT!nMes1))
            feDeclaracionPDT.TextMatrix(0, 5) = DevolverNombreMes(CInt(rsDatPDT!nMes2))
            feDeclaracionPDT.TextMatrix(0, 6) = DevolverNombreMes(CInt(rsDatPDT!nMes3))
            rsDatPDT.MoveNext
        Loop
    End If
    rsDatPDT.Close
    Set rsDatPDT = Nothing
    
    rsDatPDTDet.Close
    Set rsDatPDTDet = Nothing

    '--------------------------------------------LLENAR PROPUESTA DEL CREDITO
    Set rsDatPropuesta = oDCOMFormatosEval.RecuperaDatosCredEvalPropuesta(psCtaCod)

    Do While Not (rsDatPropuesta.EOF)
        txtFechaVisita.Text = Format(rsDatPropuesta!dFecVisita, "dd/mm/yyyy")
        txtEntornoFamiliar2.Text = Trim(rsDatPropuesta!cEntornoFami)
        txtGiroUbicacion2.Text = Trim(rsDatPropuesta!cGiroUbica)
        txtExperiencia2.Text = Trim(rsDatPropuesta!cExpeCrediticia)
        txtFormalidadNegocio2.Text = Trim(rsDatPropuesta!cFormalNegocio)
        txtColaterales2.Text = Trim(rsDatPropuesta!cColateGarantia)
        txtDestino2.Text = Trim(rsDatPropuesta!cDestino)
        txtComentario.Text = lcComentario
        rsDatPropuesta.MoveNext
    Loop
    '--------------------------------------------------- LLENAR REFERIDOS
    
    Set rsDatRef = oDCOMFormatosEval.RecuperaDatosReferidos(psCtaCod)
        Call LimpiaFlex(feReferidos)
            Do While Not (rsDatRef.EOF)
                feReferidos.AdicionaFila
                lnFila = feReferidos.row
                feReferidos.TextMatrix(lnFila, 0) = rsDatRef!nCodRef
                feReferidos.TextMatrix(lnFila, 1) = rsDatRef!cNombre
                feReferidos.TextMatrix(lnFila, 2) = rsDatRef!cDniNom
                feReferidos.TextMatrix(lnFila, 3) = rsDatRef!cTelf
                feReferidos.TextMatrix(lnFila, 4) = rsDatRef!cReferido
                feReferidos.TextMatrix(lnFila, 5) = rsDatRef!cDNIRef
                rsDatRef.MoveNext
            Loop
        rsDatRef.Close
        Set rsDatRef = Nothing

        lnTotActivo1 = CDbl(Me.feActivos.TextMatrix(17, 2))
        lnTotActivo2 = CDbl(Me.feActivos.TextMatrix(17, 3))
        lnTotPasivo1 = CDbl(Me.fePasivos.TextMatrix(23, 2))
        lnTotPasivo2 = CDbl(Me.fePasivos.TextMatrix(23, 3))
        lnResulEjer1 = CDbl(Me.fePasivos.TextMatrix(21, 2))
        lnResulEjer2 = CDbl(Me.fePasivos.TextMatrix(21, 3))
        lnResulAcum1 = CDbl(Me.fePasivos.TextMatrix(22, 2))
        lnResulAcum2 = CDbl(Me.fePasivos.TextMatrix(22, 3))

        Me.fePasivos.TextMatrix(17, 2) = lnTotActivo1 - lnTotPasivo1 - lnResulEjer1 - lnResulAcum1
        Me.fePasivos.TextMatrix(17, 3) = lnTotActivo2 - lnTotPasivo2 - lnResulEjer2 - lnResulAcum2
        
        Call CargaRatiosIndicadores
        Call CargaParametrosEvaluacion
        
        CalculaCeldas (1)
        CalculaCeldas (2)
        CalculaCeldas (2)
        CalculaCeldas (3)
        Call CalculaCoeFinan
        
        MsgBox "Se cargaron los datos Satisfactoriamente.", vbOKOnly, "Atención"

End Sub
Private Function DevolverMes(ByVal pnMes As Integer, ByRef pnAnio As Integer, ByRef pnMesN As Integer) As String 'Cargar Ultimo 3 Meses -> Registrar
    Dim nIndMes As Integer
    nIndMes = CInt(Mid(gdFecSis, 4, 2)) - pnMes
    pnAnio = CInt(Mid(gdFecSis, 7, 4))
        If nIndMes < 1 Then
            nIndMes = nIndMes + 12
            pnAnio = pnAnio - 1
        End If
    pnMesN = nIndMes
    DevolverMes = Choose(nIndMes, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
End Function
Private Sub OptCondLocal_Click(index As Integer)
    Select Case index
    Case 2
        Me.txtCondLocalOtros.Visible = False
        Me.txtCondLocalOtros.Text = ""
    End Select
    lnCondLocal = index
End Sub
Private Sub OptCondLocal2_Click(index As Integer)
    Select Case index
    Case 1, 3
    
        Me.txtCondLocalOtros.Visible = False
        Me.txtCondLocalOtros.Text = ""
    Case 4
        Me.txtCondLocalOtros.Visible = True
        Me.txtCondLocalOtros.Text = ""
    End Select
    lnCondLocal = index
End Sub

'***** LUCV20160528 - FeReferidos2
Private Sub cmdQuitar2_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feReferidos.EliminaFila (feReferidos.row)
        'txtTotalIfis.Text = Format(SumarCampo(feReferidos2, 2), "#,##0.00")
    End If
End Sub

'***** LUCV20160528
Private Sub cmdAgregar2_Click()
    
    If feReferidos.rows - 1 < 25 Then
        feReferidos.lbEditarFlex = True
        feReferidos.AdicionaFila
        feReferidos.SetFocus
        feReferidos.AvanceCeldas = Horizontal
        SendKeys "{Enter}"
    Else
        MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub feReferidos_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 1 Or pnCol = 4 Then
        feReferidos.TextMatrix(pnRow, pnCol) = UCase(feReferidos.TextMatrix(pnRow, pnCol))
    End If
    
    Select Case pnCol
    Case 2
        If IsNumeric(feReferidos.TextMatrix(pnRow, pnCol)) Then
            If Len(feReferidos.TextMatrix(pnRow, pnCol)) = 8 Then
                Select Case CCur(feReferidos.TextMatrix(pnRow, pnCol))
                    Case Is > 0
                    Case Else
                        MsgBox "Por favor, verifique el DNI", vbInformation, "Alerta"
                        feReferidos.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            Else
                MsgBox "El DNI, tiene que ser 8 dígitos.", vbInformation, "Alerta"
                feReferidos.TextMatrix(pnRow, pnCol) = 0
            End If
            
        Else
            MsgBox "El DNI, tiene que ser numérico.", vbInformation, "Alerta"
            feReferidos.TextMatrix(pnRow, pnCol) = 0
        End If
    Case 3
        If IsNumeric(feReferidos.TextMatrix(pnRow, pnCol)) Then
            If Len(feReferidos.TextMatrix(pnRow, pnCol)) = 9 Then
                Select Case CCur(feReferidos.TextMatrix(pnRow, pnCol))
                Case Is > 0
                Case Else
                    MsgBox "Teléfono Mal Ingresado", vbInformation, "Alerta"
                    feReferidos.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            ElseIf Len(feReferidos.TextMatrix(pnRow, pnCol)) < 9 Then
                MsgBox "Faltan caracteres en el teléfono / celular.", vbInformation, "Alerta"
                feReferidos.TextMatrix(pnRow, pnCol) = 0
            Else
                MsgBox "Solo se acepta nueve(9) dígitos en el teléfono / celular.", vbInformation, "Alerta"
                feReferidos.TextMatrix(pnRow, pnCol) = 0
                
            End If
        Else
            MsgBox "El telefono, solo permite ingreso de datos tipo numérico." & Chr(10) & "Ejemplo: 065404040, 984047523 ", vbInformation, "Alerta"
            feReferidos.TextMatrix(pnRow, pnCol) = 0
        End If
    End Select
End Sub

Private Sub cmdsalir_Click()
    Unload Me
    rsDatParamFlujoCaja.Close 'LUCV20171015, ERS051-2017
    Set rsDatParamFlujoCaja = Nothing 'LUCV20171015, ERS051-2017
End Sub

Public Function Inicio(ByVal psTipoRegMant As Integer, ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnProducto As Integer, _
                    ByVal pnSubProducto As Integer, ByVal pnMontoExpCred As Double, ByVal pbImprimir As Boolean) As Boolean


    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim rsDCredito As ADODB.Recordset
    Dim rsDCredEval As ADODB.Recordset
    Dim rsDColCred As ADODB.Recordset
    Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio

    lnNumForm = 6

    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio

    Me.feActivos.Enabled = False
    Me.fePasivos.Enabled = False
    Me.feEstaGananPerd.Enabled = False
    Me.feFlujoCajaMensual.Enabled = False
    Me.feGastosFamiliares.Enabled = False
    Me.feOtrosIngresos.Enabled = False
    Me.feDeclaracionPDT.Enabled = False
    Me.feFlujoCajaHistorico.Enabled = False 'LUCV20171015, Agregó según ERS0512017

    fsCtaCod = psCtaCod
    nTC = oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoDia)

    fnTipoRegMant = psTipoRegMant
    ActXCodCta.NroCuenta = fsCtaCod
    
    lnCompraDeuda = 0: lnMontoAmpliado = 0 'PEAC 20160926
    
    fnTipoPermiso = oNCOMFormatosEval.ObtieneTipoPermisoCredEval(gsCodCargo) ' Obtener el tipo de Permiso, Segun Cargo
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval

    'Obtener. Deuda de Linea de Credito No Autorizada ****
    Set rsDLineaCNU = oDCOMFormatosEval.RecuperaDeudaLineaCreditoNU(fsCtaCod)
    Set rsDCredEval = oDCOMFormatosEval.RecuperaColocacCredEval(fsCtaCod) 'Ojo: Recuperar Credito Si ha sido Registrado el Form. Eval.

    'monto solicitado
    Dim oDatComu As COMDCredito.DCOMFormatosEval
    Set oDatComu = New COMDCredito.DCOMFormatosEval
    Dim rsDatComuFormEval As ADODB.Recordset
'    Dim lnMontoSol As Currency
    Set rsDatComuFormEval = oDatComu.ObtieneDatosComunes(fsCtaCod)
    lnMontoSol = rsDatComuFormEval!nMontoSol
    
    lnCompraDeuda = rsDatComuFormEval!nCompraDeuda 'PEAC 20160926
    lnMontoAmpliado = rsDatComuFormEval!nAmpliado 'PEAC 20160926
    
    rsDatComuFormEval.Close
    Set rsDatComuFormEval = Nothing
    Set oDatComu = Nothing
    
    gsOpeCod = "" 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
    lcMovNro = "" 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
    
    Call CargaControlesInicio

'    If fnTipoRegMant = 1 Then 'Para el Evento: "Registrar"
'        If Not rsCredEval.EOF Then
'            'Call Mantenimiento
'            fnTipoRegMant = 2
'        Else
'            Call Registro
'            fnTipoRegMant = 1
'        End If
'    ElseIf fnTipoRegMant = 2 Then  'Para el Evento. "Mantenimiento"
'        If rsDCredEval.EOF Then
'            Call Registro
'            fnTipoRegMant = 1
'        Else
'            'Call Mantenimiento
'            fnTipoRegMant = 2
'        End If
'    ElseIf fnTipoRegMant = 3 Then  'Para el Evento. "consulta"
'            Call Consulta
'    End If

    Set rsDCredito = oDCOMFormatosEval.RecuperaSolicitudDatoBasicosEval(fsCtaCod) ' Datos Basicos del Credito Solicitado

    If (rsDCredito!cActiGiro) = "" Then
        MsgBox "Por favor, actualizar los datos del cliente. " & Chr(13) & " (Actividad o Giro del negocio)", vbInformation, "Alerta"
        Exit Function
    End If
    
    lnPrdEstado = rsDCredito!nPrdEstado
    lnColocCondi = rsDCredito!nColocCondicion
    fbTieneReferido6Meses = rsDCredito!bTieneReferido6Meses 'LUCV20171115, Agregó segun correo RUSI
    fnMontoIni = Trim(rsDCredito!nMonto)
    fsCliente = Trim(rsDCredito!cPersNombre)
    fsGiroNego = IIf((rsDCredito!cActiGiro) = "", "", (rsDCredito!cActiGiro))
    fsAnioExp = CInt(rsDCredito!nAnio)
    fsMesExp = CInt(rsDCredito!nMes)

    Me.txtFechaEvaluacion = Format(rsDCredito!dFecEval, "dd/MM/yyyy")
    
    Me.spnTiempoLocalAnio.valor = CInt(rsDCredito!nTmpoLocalAnio)
    Me.spnTiempoLocalMes.valor = CInt(rsDCredito!nTmpoLocalMes)
    Me.txtInversion.Text = Format(rsDCredito!nInversionFlujo, "#,##0.00")
    Me.txtFinanciamiento.Text = Format(rsDCredito!nFinanciamientoFlujo, "#,##0.00")
    
    Select Case rsDCredito!nCondiLocal
        Case 1
            OptCondLocal2(1).value = True
        Case 2
            OptCondLocal(2).value = True
        Case 3
            OptCondLocal2(3).value = True
        Case 4
            OptCondLocal2(4).value = True
            txtCondLocalOtros.Text = rsDCredito!cCondiLocalOtro
    End Select

    'SI CONDICION DE CREDITO ES NUEVO
    '->***** LUCV20171115, Comentó, según correo: RUSI
    'If lnColocCondi = 1 Then
    If Not fbTieneReferido6Meses Then
        Me.txtComentario.Enabled = True
        Me.feReferidos.Enabled = True
        Me.cmdAgregar2.Enabled = True
        Me.cmdQuitar2.Enabled = True
    Else
        Me.txtComentario.Enabled = False
        Me.feReferidos.Enabled = False
        Me.cmdAgregar2.Enabled = False
        Me.cmdQuitar2.Enabled = False
    End If
    '<-***** Fin LUCV20171115
    
    'si credito es refinanciado no se ingresa propuesta del credito(inf. visita)
    If lnColocCondi = 4 Then
        txtFechaVisita.Enabled = False
        txtEntornoFamiliar2.Enabled = False
        txtGiroUbicacion2.Enabled = False
        txtExperiencia2.Enabled = False
        txtFormalidadNegocio2.Enabled = False
        txtColaterales2.Enabled = False
        txtDestino2.Enabled = False
    Else
        txtFechaVisita.Enabled = True
        txtEntornoFamiliar2.Enabled = True
        txtGiroUbicacion2.Enabled = True
        txtExperiencia2.Enabled = True
        txtFormalidadNegocio2.Enabled = True
        txtColaterales2.Enabled = True
        txtDestino2.Enabled = True
    End If

    fnMontoDeudaSbs = Format(CCur(rsDCredito!nMontoUltimaDeudaSBS), "#,##0.00")
    Me.txtExposicionCredito2.Text = Format(pnMontoExpCred, "#,#0.00")
    
    Me.txtNombreCliente2.Text = fsCliente
    Me.txtGiroNeg2.Text = fsGiroNego
    
    txtUltEndeuda.Text = Format(fnMontoDeudaSbs, "#,##0.00")
    txtFecUltEndeuda.Text = Format(IIf(rsDCredito!dFechaUltimaDeudaSBS = "", "__/__/____", rsDCredito!dFechaUltimaDeudaSBS), "dd/mm/yyyy")

    txtUltEndeuda.Enabled = False
    spnExpEmpAnio.valor = fsAnioExp
    spnExpEmpMes.valor = fsMesExp
     
    cSPrd = Trim(rsDCredito!cTpoProdCod)
    cPrd = Mid(cSPrd, 1, 1) & "00"
    fbPermiteGrabar = False
    fbBloqueaTodo = False
   
    If fnTipoPermiso = 2 Then
       If rsDCredEval.RecordCount = 0 Then ' Si no hay credito registrado
            MsgBox "El analista no ha registrado la Evaluacion respectiva", vbExclamation, "Aviso"
            fbPermiteGrabar = False
        Else
            fbPermiteGrabar = True
         End If
    End If
    
    Set rsDCredito = Nothing
    Set rsDCredEval = Nothing
    
    Set rsDColCred = oDCOMFormatosEval.RecuperaColocacCred(fsCtaCod) ' PARA VERFICAR SI FUE VERIFICADO
    If rsDColCred!nVerifCredEval = 1 Then
        MsgBox "Ud. no puede editar la evaluación, ya se realizó la verificacion del credito", vbExclamation, "Aviso"
        fbBloqueaTodo = True
    End If
    
    nFormato = pnFormato
    
    Set oDCOMFormatosEval = Nothing
    Set oTipoCam = Nothing
    
    If fnTipoRegMant = 3 Then  'LUCV20171015, Agregó: Para Consultas
        Call Consulta
        'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
        gsOpeCod = gCredConsultarEvaluacionCred
        lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gConsultar, "Evaluacion Credito Formato 6", fsCtaCod, gCodigoCuenta
        Set objPista = Nothing
        'Fin LUCV20181220
    End If
    
    If CargaDatos Then
        If CargaControles(fnTipoPermiso, fbPermiteGrabar, fbBloqueaTodo) Then
            If fnTipoRegMant = 1 Then 'Para el Evento: "Registrar"
                If Not rsCredEval.EOF Then
                    'Call Mantenimiento
                    fnTipoRegMant = 2
                Else
                    Call Registro
                    fnTipoRegMant = 1
                End If
            ElseIf fnTipoRegMant = 2 Then  'Para el Evento. "Mantenimiento"
                If rsCredEval.EOF Then
                    Call Registro
                    fnTipoRegMant = 1
                Else
                    'Call Mantenimiento
                    fnTipoRegMant = 2
                End If
            ElseIf fnTipoRegMant = 3 Then  'Para el Evento. "consulta"
                    Call Consulta
            End If
        Else
            Unload Me
            Exit Function
        End If
    Else
        If CargaControles(1, False) Then
        End If
    End If
    
    
    'LUCV20171015, Comentó según ERS051-2017
'    If lnPrdEstado = 2000 Then
'        'Me.SSTabRatios.Enabled = False
'        Me.SSTabRatios.Visible = False 'LUCV20170424
'        Me.Height = 9900 '9800 'LUCV20170424
'    Else
'        'Me.SSTabRatios.Enabled = True
'        Me.SSTabRatios.Visible = True
'        Me.Height = 11000 ' 10900 'LUCV20170424
'        Call HabilitaControles(True, True, True)
'    End If
    If lnPrdEstado <> 2000 Then
        Call HabilitaControles(True, True, True)
    End If
    'Fin LUCV20171015
    
'    '----------------------------------LLENAR PROPUESTA DEL CREDITO
'    Set rsDatPropuesta = oDCOMFormatosEval.RecuperaDatosCredEvalPropuesta(psCtaCod)
'
'    If Not (rsDatPropuesta.EOF And rsDatPropuesta.BOF) Then
'        Do While Not rsDatPropuesta.EOF
'            txtFechaVisita.Text = Format(rsDatPropuesta!dFecVisita, "dd/mm/yyyy")
'            txtEntornoFamiliar2.Text = Trim(rsDatPropuesta!cEntornoFami)
'            txtGiroUbicacion2.Text = Trim(rsDatPropuesta!cGiroUbica)
'            txtExperiencia2.Text = Trim(rsDatPropuesta!cExpeCrediticia)
'            txtFormalidadNegocio2.Text = Trim(rsDatPropuesta!cFormalNegocio)
'            txtColaterales2.Text = Trim(rsDatPropuesta!cColateGarantia)
'            txtDestino2.Text = Trim(rsDatPropuesta!cDestino)
'            txtComentario.Text = lcComentario
'            rsDatPropuesta.MoveNext
'        Loop
'    End If
'    '------------------------------------------------------------------
    Me.SSTabIngresos2.Tab = 0
    If Not pbImprimir Then
        Me.Show 1
    Else
        cmdImprimir_Click
    End If
    Inicio = fbGrabar
End Function
Public Sub Consulta()
    Me.cmdGuardar.Enabled = False
    Me.cmdVerCar2.Enabled = False
    Me.cmdInformeVista2.Enabled = False
    Me.cmdAgregaEEFF.Enabled = False
    Me.cmdImprimir.Enabled = False
    Me.cmdImpEEFF.Enabled = False
End Sub

'***** LUCV20160529 / feReferidos2
Public Function ValidaDatosReferencia() As Boolean
    Dim i As Integer, j As Integer
    ValidaDatosReferencia = False

    If feReferidos.rows - 1 < 2 Then
        MsgBox "Debe registrar por lo menos 2 referencias para continuar", vbInformation, "Aviso"
        cmdAgregar2.SetFocus
        ValidaDatosReferencia = False
        Exit Function
    End If
    
    For i = 1 To feReferidos.rows - 1  'Verfica Tipo de Valores del DNI
        If Trim(feReferidos.TextMatrix(i, 1)) <> "" Then
            For j = 1 To Len(Trim(feReferidos.TextMatrix(i, 2)))
                If (Mid(feReferidos.TextMatrix(i, 2), j, 1) < "0" Or Mid(feReferidos.TextMatrix(i, 2), j, 1) > "9") Then
                   MsgBox "Uno de los Digitos del primer DNI de la fila " & i & " no es un Numero", vbInformation, "Aviso"
                   feReferidos.SetFocus
                   ValidaDatosReferencia = False
                   Exit Function
                End If
            Next j
        End If
    Next i
    
    For i = 1 To feReferidos.rows - 1  'Verfica Longitud del DNI
        If Trim(feReferidos.TextMatrix(i, 1)) <> "" Then
            If Len(Trim(feReferidos.TextMatrix(i, 2))) <> gnNroDigitosDNI Then
                MsgBox "Primer DNI de la fila " & i & " no es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
                feReferidos.SetFocus
                ValidaDatosReferencia = False
                Exit Function
            End If
        End If
    Next i
    
    For i = 1 To feReferidos.rows - 1  'Verfica Tipo de Valores del Telefono
        If Trim(feReferidos.TextMatrix(i, 1)) <> "" Then
            For j = 1 To Len(Trim(feReferidos.TextMatrix(i, 3)))
                If (Mid(feReferidos.TextMatrix(i, 3), j, 1) < "0" Or Mid(feReferidos.TextMatrix(i, 3), j, 1) > "9") Then
                   MsgBox "Uno de los Digitos del teléfono de la fila " & i & " no es un Numero", vbInformation, "Aviso"
                   feReferidos.SetFocus
                   ValidaDatosReferencia = False
                   Exit Function
                End If
            Next j
        End If
    Next i
    
    For i = 1 To feReferidos.rows - 1 'Verfica Tipo de Valores del DNI 2
        If Trim(feReferidos.TextMatrix(i, 1)) <> "" Then
            For j = 1 To Len(Trim(feReferidos.TextMatrix(i, 5)))
                If (Mid(feReferidos.TextMatrix(i, 5), j, 1) < "0" Or Mid(feReferidos.TextMatrix(i, 5), j, 1) > "9") Then
                   MsgBox "Uno de los Digitos del segundo DNI de la fila " & i & " no es un Numero", vbInformation, "Aviso"
                   feReferidos.SetFocus
                   ValidaDatosReferencia = False
                   Exit Function
                End If
            Next j
        End If
    Next i
    ValidaDatosReferencia = True
End Function


Public Function ValidaDatos() As Boolean
ValidaDatos = False

    If txtFechaVisita.Text = "__/__/____" Then
        MsgBox "Ingrese la fecha de visita en la propuesta del Crédito", vbOKOnly, "Atención"
        Exit Function
    End If
    If txtEntornoFamiliar2.Text = "" Then
        MsgBox "Ingrese el entorno familiar en la propuesta del Crédito", vbOKOnly, "Atención"
        Exit Function
    End If
    If txtGiroUbicacion2.Text = "" Then
        MsgBox "Ingrese la ubicación del Giro en la propuesta del Crédito", vbOKOnly, "Atención"
        Exit Function
    End If
    If txtExperiencia2.Text = "" Then
        MsgBox "Ingrese la experiencia en la propuesta del Crédito", vbOKOnly, "Atención"
        Exit Function
    End If
    If txtFormalidadNegocio2.Text = "" Then
        MsgBox "Ingrese la formalidad del negocio en la propuesta del Crédito", vbOKOnly, "Atención"
        Exit Function
    End If
    If txtColaterales2.Text = "" Then
        MsgBox "Ingrese los colaterales y garantías en la propuesta del Crédito", vbOKOnly, "Atención"
        Exit Function
    End If
    If txtDestino2.Text = "" Then
        MsgBox "Ingrese sobre el destino e impacto en la propuesta del Crédito", vbOKOnly, "Atención"
        Exit Function
    End If

ValidaDatos = True
End Function
Private Function CargaControles(ByVal TipoPermiso As Integer, ByVal pPermiteGrabar As Boolean, Optional ByVal pBloqueaTodo As Boolean = False) As Boolean
    If TipoPermiso = 1 Then
        Call HabilitaControles(False, False, False)
        CargaControles = True
    ElseIf TipoPermiso = 2 Then
        Call HabilitaControles(False, True, pPermiteGrabar)
        CargaControles = True
    ElseIf TipoPermiso = 3 Then
        Call HabilitaControles(False, False, True)
        CargaControles = True
    Else
        MsgBox "No tiene Permisos para este módulo", vbInformation, "Aviso"
        CargaControles = False
    End If
    If pBloqueaTodo Then
        Call HabilitaControles(False, False, False)
        CargaControles = True
    End If
End Function

Private Function HabilitaControles(ByVal pbHabilitaA As Boolean, ByVal pbHabilitaB As Boolean, ByVal pbHabilitaGuardar As Boolean)
    cmdInformeVista2.Enabled = pbHabilitaA
    cmdVerCar2.Enabled = pbHabilitaA
    cmdImprimir.Enabled = pbHabilitaA
    cmdImpEEFF.Enabled = pbHabilitaA
    cmdFlujoCaja.Enabled = pbHabilitaA
End Function

Private Function Registro()
    gsOpeCod = gCredRegistrarEvaluacionCred
    txtNombreCliente2.Text = fsCliente
    txtNombreCliente2.Enabled = False
    txtGiroNeg2.Text = fsGiroNego
    txtGiroNeg2.Enabled = False
End Function

Private Sub CargaControlesInicio()
    Call CargaDatosCboFecEEFF
    Call CargarFlexEdit
End Sub

Private Sub CargaDatosCboFecEEFF()
    Dim oDCred As COMDCredito.DCOMFormatosEval
    Dim rs As ADODB.Recordset
    
    Set oDCred = New COMDCredito.DCOMFormatosEval
    Set rs = oDCred.RecuperaDatosCredEvalFechasEEFFEvalForm6(fsCtaCod)
    Set oDCred = Nothing
    
'    lcFecRegEF = ""
    CboFecRegEEFF.Clear
    Do While Not rs.EOF
        CboFecRegEEFF.AddItem rs!dfecReg  '& Space(250) & rs!cPersCod
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
Private Sub CargarFlexEdit()
    Dim lnFila As Integer
    Dim CargarFlexEdit As Boolean
    Dim nMonto As Double
    
    Dim i As Integer
    Dim nFila As Integer
    Dim NumRegRS As Integer
    Dim NumRegRSPasivos As Integer
    Dim NumRegRSEstGanPer As Integer
    Dim NumRegRSCoefiFinan As Integer
    Dim nMontoIni As Double

    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval

    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval

    cuotaifi = 9
    nMonto = Format(0, "00.00")
    feActivos.Clear
    feActivos.FormaCabecera

CargarFlexEdit = oNCOMFormatosEval.CargaDatosFlexEdit(6, fsCtaCod _
                                                     , _
                                                     , rsFeDatGastoFam _
                                                     , rsFeDatOtrosIng _
                                                     , _
                                                     , _
                                                     , _
                                                     , _
                                                     , _
                                                     , _
                                                     , rsFeFlujoCaja _
                                                     , rsFeDatActivosForm6 _
                                                     , rsFeDatPasivosForm6 _
                                                     , rsFeDatEstadoGanPerdForm6 _
                                                     , rsFeDatCoeficienteFinanForm6 _
                                                     , rsFeDatPDT, , , , , _
                                                     rsFeFlujoCajaHistorico, _
                                                     rsDatParamFlujoCaja)
                                                     'LUCV20171015, ERS051-2017 -Agregó: rsFeFlujoCajaHistorico, rsDatParamFlujoCaja

'---------------------------- Activos
    feActivos.FormaCabecera
        
    feActivos.rows = 2
        Call LimpiaFlex(feActivos)
        nFila = 0
        NumRegRS = rsFeDatActivosForm6.RecordCount
        ReDim lvPrincipalActivos(NumRegRS)
        
    Do While Not rsFeDatActivosForm6.EOF
        feActivos.AdicionaFila
        lnFila = feActivos.row
        feActivos.TextMatrix(lnFila, 1) = rsFeDatActivosForm6!Concepto
        feActivos.TextMatrix(lnFila, 2) = Format(rsFeDatActivosForm6!PP, "#,#0.00")
        feActivos.TextMatrix(lnFila, 3) = Format(rsFeDatActivosForm6!PE, "#,#0.00")
        feActivos.TextMatrix(lnFila, 4) = Format(rsFeDatActivosForm6!Total, "#,#0.00")
        feActivos.TextMatrix(lnFila, 5) = rsFeDatActivosForm6!nConsCod
        feActivos.TextMatrix(lnFila, 6) = rsFeDatActivosForm6!nConsValor
                                
        Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0))
            Case 2, 4, 5, 6, 7, 8, 9, 11, 14, 16
                 Me.feActivos.BackColorRow &HC0FFFF, True 'color amarillo claro
        End Select
        
        Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0))
            Case 1, 10, 17
                Me.feActivos.BackColorRow QBColor(8), True 'gris
        End Select
        
        rsFeDatActivosForm6.MoveNext
    Loop
    rsFeDatActivosForm6.Close
    Set rsFeDatActivosForm6 = Nothing

'----------------- Pasivos
    fePasivos.FormaCabecera
    fePasivos.rows = 2
        Call LimpiaFlex(fePasivos)
        
        nFila = 0
        NumRegRSPasivos = rsFeDatPasivosForm6.RecordCount
        ReDim lvPrincipalPasivos(NumRegRSPasivos)
        
    Do While Not rsFeDatPasivosForm6.EOF
        fePasivos.AdicionaFila
        lnFila = fePasivos.row
        fePasivos.TextMatrix(lnFila, 1) = rsFeDatPasivosForm6!Concepto
        fePasivos.TextMatrix(lnFila, 2) = Format(rsFeDatPasivosForm6!PP, "#,#0.00")
        fePasivos.TextMatrix(lnFila, 3) = Format(rsFeDatPasivosForm6!PE, "#,#0.00")
        fePasivos.TextMatrix(lnFila, 4) = Format(rsFeDatPasivosForm6!Total, "#,#0.00")
        fePasivos.TextMatrix(lnFila, 5) = rsFeDatPasivosForm6!nConsCod
        fePasivos.TextMatrix(lnFila, 6) = rsFeDatPasivosForm6!nConsValor
        
        '-----------------pinta items que se ingresaran detalles en pasivos
        Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0))
            Case 2, 5, 6, 7, 9
                Me.fePasivos.BackColorRow &HC0FFFF, True
        End Select

        Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0))
            Case 1, 10, 16
                Me.fePasivos.BackColorRow QBColor(8), True 'color gris
        End Select
        
        rsFeDatPasivosForm6.MoveNext
    Loop
    rsFeDatPasivosForm6.Close
    Set rsFeDatPasivosForm6 = Nothing

'----------------- ESTADO DE GANANCIA Y PERDIDAS
     
    feEstaGananPerd.FormaCabecera
    feEstaGananPerd.rows = 2
        Call LimpiaFlex(feEstaGananPerd)
        
        nFila = 0
        NumRegRSEstGanPer = rsFeDatEstadoGanPerdForm6.RecordCount
        ReDim lvPrincipalEstGanPer(NumRegRSEstGanPer)

    Do While Not rsFeDatEstadoGanPerdForm6.EOF
        feEstaGananPerd.AdicionaFila
        lnFila = feEstaGananPerd.row
        feEstaGananPerd.TextMatrix(lnFila, 1) = rsFeDatEstadoGanPerdForm6!Concepto
        feEstaGananPerd.TextMatrix(lnFila, 2) = Format(rsFeDatEstadoGanPerdForm6!nMonto, "#,#0.00")
        feEstaGananPerd.TextMatrix(lnFila, 3) = rsFeDatEstadoGanPerdForm6!nConsCod
        feEstaGananPerd.TextMatrix(lnFila, 4) = rsFeDatEstadoGanPerdForm6!nConsValor

'        '-----------------pinta items DE TOTALES
        Select Case CInt(feEstaGananPerd.TextMatrix(Me.feEstaGananPerd.row, 0))
            Case 3, 6, 9, 16, 19
                'Me.feEstaGananPerd.CellBackColor() = QBColor(8)
                Me.feEstaGananPerd.BackColorRow QBColor(8), True ' color gris
        End Select
        
        rsFeDatEstadoGanPerdForm6.MoveNext
    Loop
    rsFeDatEstadoGanPerdForm6.Close
    Set rsFeDatEstadoGanPerdForm6 = Nothing

'----------------- COEFICIENTE FINANCIERO
    Me.feCoeFinan.FormaCabecera
    feCoeFinan.rows = 2
        Call LimpiaFlex(feCoeFinan)
        
        nFila = 0
        NumRegRSCoefiFinan = rsFeDatCoeficienteFinanForm6.RecordCount
        ReDim lvPrincipalCoefiFinan(NumRegRSCoefiFinan)

    Do While Not rsFeDatCoeficienteFinanForm6.EOF
        feCoeFinan.AdicionaFila
        lnFila = feCoeFinan.row
        feCoeFinan.TextMatrix(lnFila, 1) = rsFeDatCoeficienteFinanForm6!Concepto
        feCoeFinan.TextMatrix(lnFila, 2) = Format(rsFeDatCoeficienteFinanForm6!nMonto, "#,#0.00")
        feCoeFinan.TextMatrix(lnFila, 3) = rsFeDatCoeficienteFinanForm6!nConsCod
        feCoeFinan.TextMatrix(lnFila, 4) = rsFeDatCoeficienteFinanForm6!nConsValor
        
        '-----------------pinta items TOTALES
        Select Case CInt(feCoeFinan.TextMatrix(Me.feCoeFinan.row, 0))
            Case 1, 6, 12, 20
                Me.feCoeFinan.CellBackColor() = QBColor(8)
        End Select
        
        rsFeDatCoeficienteFinanForm6.MoveNext
    Loop
    rsFeDatCoeficienteFinanForm6.Close
    Set rsFeDatCoeficienteFinanForm6 = Nothing

'------------------- FLUJO DE CAJA
    feFlujoCajaMensual.FormaCabecera
    feFlujoCajaMensual.rows = 2
    Call LimpiaFlex(feFlujoCajaMensual)
        Do While Not rsFeFlujoCaja.EOF
            feFlujoCajaMensual.AdicionaFila
            lnFila = feFlujoCajaMensual.row
            feFlujoCajaMensual.TextMatrix(lnFila, 1) = rsFeFlujoCaja!cConsDescripcion
            feFlujoCajaMensual.TextMatrix(lnFila, 2) = Format(rsFeFlujoCaja!nMonto, "#,#0.00")
            feFlujoCajaMensual.TextMatrix(lnFila, 3) = rsFeFlujoCaja!nConsCod
            feFlujoCajaMensual.TextMatrix(lnFila, 4) = rsFeFlujoCaja!nConsValor
            
            Select Case CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 0))
                Case 5, 6, 18, 22
                     Me.feFlujoCajaMensual.BackColorRow QBColor(8), True 'gris
                Case 19
                     Me.feFlujoCajaMensual.BackColorRow &HC0FFFF, True 'color amarillo claro
            End Select
            rsFeFlujoCaja.MoveNext
        Loop
    rsFeFlujoCaja.Close
    Set rsFeFlujoCaja = Nothing

'------------------------------- GASTOS FAMILIARES
    
    feGastosFamiliares.FormaCabecera
    feGastosFamiliares.rows = 2
    Call LimpiaFlex(feGastosFamiliares)
        Do While Not rsFeDatGastoFam.EOF
            feGastosFamiliares.AdicionaFila
            lnFila = feGastosFamiliares.row
            feGastosFamiliares.TextMatrix(lnFila, 1) = rsFeDatGastoFam!nConsValor
            feGastosFamiliares.TextMatrix(lnFila, 2) = rsFeDatGastoFam!cConsDescripcion
            feGastosFamiliares.TextMatrix(lnFila, 3) = Format(rsFeDatGastoFam!nMonto, "#,#0.00")
        Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) 'celda que  o se puede editar
            Case gCodCuotaIfiGastoFami
                Me.feGastosFamiliares.BackColorRow &HC0FFFF, True 'color amarillo
                Me.feGastosFamiliares.ForeColorRow (&H80000007), True
                Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
            Case gCodDeudaLCNUGastoFami
                Me.feGastosFamiliares.BackColorRow vbWhite
                Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
            Case Else
            Me.feGastosFamiliares.BackColorRow &HFFFFFF
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
        End Select
            rsFeDatGastoFam.MoveNext
        Loop
    rsFeDatGastoFam.Close
    Set rsFeDatGastoFam = Nothing
'----------------------- OTROS INGRESOS

    feOtrosIngresos.FormaCabecera
    feOtrosIngresos.rows = 2
    Call LimpiaFlex(feOtrosIngresos)
        Do While Not rsFeDatOtrosIng.EOF
            feOtrosIngresos.AdicionaFila
            lnFila = feOtrosIngresos.row
            feOtrosIngresos.TextMatrix(lnFila, 1) = rsFeDatOtrosIng!nConsValor
            feOtrosIngresos.TextMatrix(lnFila, 2) = rsFeDatOtrosIng!cConsDescripcion
            feOtrosIngresos.TextMatrix(lnFila, 3) = Format(nMontoIni, "#,#0.00")
            rsFeDatOtrosIng.MoveNext
        Loop
    rsFeDatOtrosIng.Close
    Set rsFeDatOtrosIng = Nothing
'-------------------------------------------------------------------------------
    'Declaracion PDT
    sMes1 = DevolverMes(1, nAnio3, nMes3)
    sMes2 = DevolverMes(2, nAnio2, nMes2)
    sMes3 = DevolverMes(3, nAnio1, nMes1)
    
    'feDeclaracionPDT.Clear
    feDeclaracionPDT.FormaCabecera
    feDeclaracionPDT.rows = 2
        
    feDeclaracionPDT.TextMatrix(0, 4) = sMes3
    feDeclaracionPDT.TextMatrix(0, 5) = sMes2
    feDeclaracionPDT.TextMatrix(0, 6) = sMes1
    
    feDeclaracionPDT.TextMatrix(0, 1) = "Mes/Detalle" '& Space(8)
    For i = 1 To 2
        feDeclaracionPDT.AdicionaFila
        'feDeclaracionPDT.TextMatrix(i, 1) = Choose(i, "Compras" & Space(8), "Ventas" & Space(8))
        feDeclaracionPDT.TextMatrix(i, 1) = rsFeDatPDT!cConsDescripcion
        feDeclaracionPDT.TextMatrix(i, 2) = rsFeDatPDT!nConsCod
        feDeclaracionPDT.TextMatrix(i, 3) = rsFeDatPDT!nConsValor
        feDeclaracionPDT.TextMatrix(i, 4) = Choose(i, "0.00", "0.00") 'Mes3
        feDeclaracionPDT.TextMatrix(i, 5) = Choose(i, "0.00", "0.00") 'Mes2
        feDeclaracionPDT.TextMatrix(i, 6) = Choose(i, "0.00", "0.00") 'Mes1
        feDeclaracionPDT.TextMatrix(i, 7) = Choose(i, "0.00", "0.00") 'Promedio
        feDeclaracionPDT.TextMatrix(i, 8) = Choose(i, "0.00", "0.00") '%Ventas
        rsFeDatPDT.MoveNext
    Next i

    '->***** LUCV20171015, Agregó según ERS0512017: FLUJO DE CAJA PROYECTADO
    feFlujoCajaHistorico.FormaCabecera
    feFlujoCajaHistorico.rows = 2
    Call LimpiaFlex(feFlujoCajaHistorico)
        Do While Not rsFeFlujoCajaHistorico.EOF
            feFlujoCajaHistorico.AdicionaFila
            lnFila = feFlujoCajaHistorico.row
            feFlujoCajaHistorico.TextMatrix(lnFila, 1) = rsFeFlujoCajaHistorico!cConsDescripcion
            feFlujoCajaHistorico.TextMatrix(lnFila, 2) = Format(rsFeFlujoCajaHistorico!nMonto, "#,#0.00")
            feFlujoCajaHistorico.TextMatrix(lnFila, 3) = rsFeFlujoCajaHistorico!nConsCod
            feFlujoCajaHistorico.TextMatrix(lnFila, 4) = rsFeFlujoCajaHistorico!nConsValor
            
            Select Case CInt(feFlujoCajaHistorico.TextMatrix(feFlujoCajaHistorico.row, 0))
                Case 1, 4
                    Me.feFlujoCajaHistorico.ForeColorRow vbGrayText
                Case 5, 6, 22
                     Me.feFlujoCajaHistorico.BackColorRow QBColor(8), True 'gris
                Case 19
                     Me.feFlujoCajaHistorico.BackColorRow &HC0FFFF, True 'color amarillo claro
            End Select
            rsFeFlujoCajaHistorico.MoveNext
        Loop
    rsFeFlujoCajaHistorico.Close
    Set rsFeFlujoCajaHistorico = Nothing
    '<-***** Fin LUCV20171015
    Call CargaRatiosIndicadores
    Call CargaParametrosEvaluacion
End Sub
'LUCV20171015, ERS0512017
Private Sub CargaParametrosEvaluacion()
    If Not (rsDatParamFlujoCaja.BOF And rsDatParamFlujoCaja.EOF) Then
        txtIncrVentasContado.Text = Format(rsDatParamFlujoCaja!nIncVentCont, "#0.00")
        txtIncrCompraMercaderia.Text = Format(rsDatParamFlujoCaja!nIncCompMerc, "#0.00")
        txtIncrPagoPersonal.Text = Format(rsDatParamFlujoCaja!nIncPagPers, "#0.00")
        txtIncrGastoVentas.Text = Format(rsDatParamFlujoCaja!nIncGastvent, "#0.00")
        txtIncrConsumo.Text = Format(rsDatParamFlujoCaja!nIncConsu, "#0.00")
    End If
End Sub
'Fin LUCV20171015

Private Sub CargaRatiosIndicadores()

    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    
    '----------- CARGA RATIOS E INDICADORES
    If lnPrdEstado > 2000 Then
        Set rsDatRatiosIndi = oDCOMFormatosEval.RecuperaDatosRatios(fsCtaCod)
    If (rsDatRatiosIndi.EOF And rsDatRatiosIndi.BOF) Then Exit Sub
    'If R.EOF And R.BOF Then
    
        txtCapacidadNeta2.Text = CStr(rsDatRatiosIndi!nCapPagNeta * 100) & "%"  '-- capacidad de pago
        'txtCapacidadRDS2.Text = CStr(0) & "%"    '-- liquidez cte *
        'txtEndeudamiento2.Text = CStr(0)  '-- endeudamiento *
        'EditMoney1.Text = CStr(0) & "%"  '-- rentabilidad *
        txtIngresoNeto2.Text = Format(rsDatRatiosIndi!nIngreNeto, "#,##0.00")  '-- ingreso neto
        txtExcedenteMensual2.Text = Format(rsDatRatiosIndi!nExceMensual, "#,##0.00")  '-- excedente mensual
        txtRentabilidad.Text = Format(rsDatRatiosIndi!nRentaPatri * 100, "#,##0.00") & "%" '-- rentabilidad
        txtLiquidezCte.Text = Format(rsDatRatiosIndi!nLiquidezCte, "#,##0.00")  '-- Liquidez Cte.
        
        rsDatRatiosIndi.Close
        Set rsDatRatiosIndi = Nothing
    End If
End Sub
Private Function CargaDatos() As Boolean
On Error GoTo ErrorCargaDatos
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim i As Integer
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
  
    Set rsCredEval = oDCOMFormatosEval.RecuperaColocacCredEval(fsCtaCod)
    
    If Not rsCredEval.EOF() Then
        lcComentario = Trim(rsCredEval!cComentario)
    Else
        lcComentario = ""
    End If
    Exit Function
ErrorCargaDatos:
    CargaDatos = False
    MsgBox Err.Description + ": Error al carga datos", vbCritical, "Error"
End Function
Private Sub spnTiempoLocalAnio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spnTiempoLocalMes.SetFocus
    End If
End Sub

Private Sub spnTiempoLocalMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.OptCondLocal2(1).SetFocus
    End If
End Sub
Private Sub txtColaterales2_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtDestino2.SetFocus
    End If
End Sub
Private Sub txtDestino2_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras3(KeyAscii, True)
End Sub
Private Sub txtEntornoFamiliar2_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtGiroUbicacion2.SetFocus
    End If
End Sub
Private Sub txtExperiencia2_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtFormalidadNegocio2.SetFocus
    End If
End Sub
Private Sub txtFechaVisita_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEntornoFamiliar2.SetFocus
    End If
End Sub
Private Sub txtFormalidadNegocio2_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtColaterales2.SetFocus
    End If
End Sub
Private Sub txtGiroUbicacion2_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtExperiencia2.SetFocus
    End If
End Sub
'->***** LUCV20171015, Agregó según ERS0512017
Private Sub txtFinanciamiento_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtFinanciamiento, KeyAscii, 10, , True)
    If KeyAscii = 45 Then KeyAscii = 0
    If KeyAscii = 13 Then
        txtInversion.SetFocus
    End If
End Sub
Private Sub txtFinanciamiento_LostFocus()
 If Trim(txtFinanciamiento.Text) = "" Then
        txtFinanciamiento.Text = "0.00"
    Else
        txtFinanciamiento.Text = Format(txtFinanciamiento.Text, "###," & String(15, "#") & "#0.00")
    End If
End Sub
Private Sub txtInversion_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtInversion, KeyAscii, 10, , True)
    If KeyAscii = 45 Then KeyAscii = 0
    If KeyAscii = 13 Then
        txtIncrVentasContado.SetFocus
    End If
End Sub
Private Sub txtInversion_LostFocus()
 If Trim(txtInversion.Text) = "" Then
        txtInversion.Text = "0.00"
    Else
        txtInversion.Text = Format(txtInversion.Text, "###," & String(15, "#") & "#0.00")
    End If
End Sub
Private Sub txtIncrVentasContado_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtIncrCompraMercaderia.SetFocus
    End If
End Sub
Private Sub txtIncrCompraMercaderia_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtIncrPagoPersonal.SetFocus
    End If
End Sub
Private Sub txtIncrPagoPersonal_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtIncrGastoVentas.SetFocus
    End If
End Sub
Private Sub txtIncrGastoVentas_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        cmdGuardar.SetFocus
    End If
End Sub
Private Sub feFlujoCajaHistorico_EnterCell()
    Select Case CInt(feFlujoCajaHistorico.TextMatrix(Me.feFlujoCajaHistorico.row, 0)) 'celda que se activa el textbuscar
        Case 19
            Me.feFlujoCajaHistorico.ListaControles = "0-0-1-0-0"
        Case Else
            Me.feFlujoCajaHistorico.ListaControles = "0-0-0-0-0"
    End Select

    Select Case CInt(feFlujoCajaHistorico.TextMatrix(Me.feFlujoCajaHistorico.row, 0)) 'celda que  o se puede editar
        Case 1, 4, 5, 6, 22
            Me.feFlujoCajaHistorico.ColumnasAEditar = "X-X-X-X-X"
        Case Else
            Me.feFlujoCajaHistorico.ColumnasAEditar = "X-X-2-X-X"
    End Select
End Sub
Private Sub feFlujoCajaHistorico_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(Me.feFlujoCajaHistorico.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feFlujoCajaHistorico.TextMatrix(pnRow, pnCol) < 0 Then
            feFlujoCajaHistorico.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(feFlujoCajaHistorico.TextMatrix(pnRow, pnCol))), "#,#0.00")   '"0.00"
        End If
    Else
        feFlujoCajaHistorico.TextMatrix(pnRow, pnCol) = "0.00"
    End If
    Call CalculaCeldas(6)
End Sub
Private Sub feFlujoCajaHistorico_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
    psCodigo = 0
    psDescripcion = ""
    psDescripcion = feFlujoCajaHistorico.TextMatrix(feFlujoCajaHistorico.row, 1) 'Cuotas Otras IFIs
    psCodigo = feFlujoCajaHistorico.TextMatrix(feFlujoCajaHistorico.row, 2) 'Monto
    If psCodigo = 0 Then
        fnTotalRefFlujoCaja = 0
        Set MatIfiFlujoCajaHistorico = Nothing
        frmCredFormEvalCuotasIfis.Inicio (CCur(feFlujoCajaHistorico.TextMatrix(feFlujoCajaHistorico.row, 2))), fnTotalRefFlujoCaja, MatIfiFlujoCajaHistorico
        psCodigo = Format(fnTotalRefFlujoCaja, "#,##0.00")
    Else
        frmCredFormEvalCuotasIfis.Inicio (CCur(feFlujoCajaHistorico.TextMatrix(feFlujoCajaHistorico.row, 2))), fnTotalRefFlujoCaja, MatIfiFlujoCajaHistorico
        psCodigo = Format(fnTotalRefFlujoCaja, "#,##0.00")
    End If
End Sub
Private Sub feFlujoCajaHistorico_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Editar = Split(Me.feFlujoCajaHistorico.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{TAB}"
        Exit Sub
    End If
    Select Case pnRow
        Case 19
            Cancel = False
            SendKeys "{TAB}"
            Exit Sub
    End Select
End Sub
Private Sub feFlujoCajaHistorico_RowColChange()
    If feFlujoCajaHistorico.Col = 2 Then
        feFlujoCajaHistorico.AvanceCeldas = Vertical
    Else
        feFlujoCajaHistorico.AvanceCeldas = Horizontal
    End If
End Sub
Private Sub cmdFlujoCaja_Click()
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As New Excel.Workbook
    Dim xlsHoja As New Excel.Worksheet
    Dim lsArchivo As String
    Dim CargaValores As Boolean
    
    Dim rsFlujoCajaRptObtieneDatosCabecera As ADODB.Recordset
    Dim rsFlujoCajaRptObtieneDatosCuotas As ADODB.Recordset
    Dim rsFlujoCajaRptObtieneDatosConceptos As ADODB.Recordset
    Dim rsFlujoCajaRptObtieneDatosParametros As ADODB.Recordset
    
    Dim oNFormatosEval As COMNCredito.NCOMFormatosEval
    Set oNFormatosEval = New COMNCredito.NCOMFormatosEval
    
    CargaValores = oNFormatosEval.CargaDatosFlujoCajaRpt(fsCtaCod, _
                                                         rsFlujoCajaRptObtieneDatosCabecera, _
                                                         rsFlujoCajaRptObtieneDatosConceptos, _
                                                         rsFlujoCajaRptObtieneDatosCuotas, _
                                                         rsFlujoCajaRptObtieneDatosParametros)
   If lnPrdEstado <> 2000 Then
        'Ruta del archivo generado
        lsArchivo = "\spooler\RptFlujoCajaMensualHistorico_" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
        
        'Configuración de Hoja
        Set xlsLibro = xlsAplicacion.Workbooks.Add
        Set xlsHoja = xlsLibro.Worksheets.Add
        xlsHoja.Name = "FlujoCajaHistoricoOperativo"
        xlsHoja.PageSetup.Orientation = xlLandscape
        xlsHoja.PageSetup.CenterHorizontally = True
        xlsHoja.PageSetup.Zoom = 60
        Call GeneraHojaFlujoCajaHistoricoRpt(rsFlujoCajaRptObtieneDatosCabecera, rsFlujoCajaRptObtieneDatosConceptos, rsFlujoCajaRptObtieneDatosCuotas, rsFlujoCajaRptObtieneDatosParametros, xlsHoja)
        
        'proteger Libro
        'xlsAplicacion.ActiveWorkbook.Protect ("" & UCase(gsCodUser) & "" & Format(gdFecSis, "YYYYMMDD") & "") 'TEMPORAL
        'xlsAplicacion.Worksheets("FlujoCajaHistoricoOperativo").Protect ("" & UCase(gsCodUser) & "" & Format(gdFecSis, "YYYYMMDD") & "") 'TEMPORAL
        
        MsgBox "Se ha generado satisfactoriamente el de flujo de caja mensual / histórico", vbInformation, "Aviso"
        xlsHoja.SaveAs App.Path & lsArchivo
        xlsAplicacion.Visible = True
        xlsAplicacion.Windows(1).Visible = True
    
        Set xlsAplicacion = Nothing
        Set xlsLibro = Nothing
        Set xlsHoja = Nothing
        Exit Sub
    Else
        MsgBox "El crédito debe estar por lo menos en estado Sugerido para mostrar el flujo de caja mensual proyectado - histórico.", vbOKOnly, "Mensaje"
        Exit Sub
    End If
End Sub
Public Sub GeneraHojaFlujoCajaHistoricoRpt(ByVal prsFlujoCajaRptObtieneDatosCabecera As ADODB.Recordset, _
                                           ByVal prsFlujoCajaRptObtieneDatosConceptos As ADODB.Recordset, _
                                           ByVal prsFlujoCajaRptObtieneDatosCuotas As ADODB.Recordset, _
                                           ByVal prsFlujoCajaRptObtieneDatosParametros As ADODB.Recordset, _
                                           ByRef xlsHoja As Worksheet)
   Dim i As Integer
   Dim lnFila As Integer
   Dim dFechaEval As Date
   Dim dfechaEvalHist As String
   Dim A As Integer
   Dim Z As Integer
   Dim nCol As Integer
   Dim nColInicio As Integer
   Dim nColFin As Integer
   Dim lnIncVentCont As Double
   Dim lnIncCompMerc As Double
   Dim lnIncConsu As Double
   Dim lnIncPagPers As Double
   Dim lnIncGastvent As Double
   lnIncVentCont = 0: lnIncCompMerc = 0: lnIncConsu = 0: lnIncPagPers = 0: lnIncGastvent = 0
   
    'Datos de la cabecera
    xlsHoja.Cells(2, 1) = "FLUJO DE CAJA MENSUAL PRESUPUESTADO"
    xlsHoja.Range(xlsHoja.Cells(2, 1), xlsHoja.Cells(2, 12)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(2, 1), xlsHoja.Cells(2, 12)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(2, 1), xlsHoja.Cells(2, 12)).Font.Bold = True
    
    xlsHoja.Cells(4, 1) = "CLIENTE: "
    xlsHoja.Range(xlsHoja.Cells(4, 1), xlsHoja.Cells(4, 1)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(4, 1), xlsHoja.Cells(4, 1)).HorizontalAlignment = xlRight
    xlsHoja.Range(xlsHoja.Cells(4, 1), xlsHoja.Cells(4, 1)).Font.Bold = True
    
    xlsHoja.Cells(5, 1) = "ANALISTA: "
    xlsHoja.Range(xlsHoja.Cells(5, 1), xlsHoja.Cells(5, 1)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(5, 1), xlsHoja.Cells(5, 1)).HorizontalAlignment = xlRight
    xlsHoja.Range(xlsHoja.Cells(5, 1), xlsHoja.Cells(5, 1)).Font.Bold = True
    
    xlsHoja.Cells(6, 1) = "DNI: "
    xlsHoja.Range(xlsHoja.Cells(6, 1), xlsHoja.Cells(6, 1)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(6, 1), xlsHoja.Cells(6, 1)).HorizontalAlignment = xlRight
    xlsHoja.Range(xlsHoja.Cells(6, 1), xlsHoja.Cells(6, 1)).Font.Bold = True
    
    xlsHoja.Cells(7, 1) = "RUC: "
    xlsHoja.Range(xlsHoja.Cells(7, 1), xlsHoja.Cells(7, 1)).MergeCells = True
    xlsHoja.Range(xlsHoja.Cells(7, 1), xlsHoja.Cells(7, 1)).HorizontalAlignment = xlRight
    xlsHoja.Range(xlsHoja.Cells(7, 1), xlsHoja.Cells(7, 1)).Font.Bold = True
    
    If Not (prsFlujoCajaRptObtieneDatosCabecera.EOF And prsFlujoCajaRptObtieneDatosCabecera.BOF) Then
        dFechaEval = prsFlujoCajaRptObtieneDatosCabecera!fechaEval
        dfechaEvalHist = prsFlujoCajaRptObtieneDatosCabecera!fechaEvalHist
        
        xlsHoja.Cells(4, 2) = prsFlujoCajaRptObtieneDatosCabecera!NombreClie
        xlsHoja.Range(xlsHoja.Cells(4, 2), xlsHoja.Cells(4, 6)).MergeCells = True
        
        xlsHoja.Cells(5, 2) = prsFlujoCajaRptObtieneDatosCabecera!NombreAnal
        xlsHoja.Range(xlsHoja.Cells(5, 2), xlsHoja.Cells(5, 6)).MergeCells = True
    
        If prsFlujoCajaRptObtieneDatosCabecera!nPersId = 1 Then
            xlsHoja.Cells(6, 2) = prsFlujoCajaRptObtieneDatosCabecera!nDoc
            xlsHoja.Range(xlsHoja.Cells(6, 2), xlsHoja.Cells(6, 6)).MergeCells = True
        Else
            xlsHoja.Cells(7, 2) = prsFlujoCajaRptObtieneDatosCabecera!nDocTrib
            xlsHoja.Range(xlsHoja.Cells(7, 2), xlsHoja.Cells(7, 6)).MergeCells = True
        End If
    Else
        MsgBox "Hubo un Error Comuníquese con el Área de TI", vbInformation, "Aviso"
        Exit Sub
    End If
    
    xlsHoja.Cells(9, 1) = "Conceptos / Meses"
    xlsHoja.Range(xlsHoja.Cells(9, 1), xlsHoja.Cells(10, 1)).MergeCells = True
        
    xlsHoja.Cells(9, 2) = "Flujo Histórico"
    xlsHoja.Cells(10, 2) = dfechaEvalHist
    
    xlsHoja.Cells(9, 3) = "Flujo Mensual"
    xlsHoja.Cells(10, 3) = Format(dFechaEval, "mmm-yyyy")
    
    xlsHoja.Range(xlsHoja.Cells(9, 1), xlsHoja.Cells(10, 200)).Font.Bold = True
    xlsHoja.Range(xlsHoja.Cells(9, 1), xlsHoja.Cells(10, 200)).HorizontalAlignment = xlCenter

    lnFila = 11
    'Conceptos
    If Not (prsFlujoCajaRptObtieneDatosConceptos.EOF And prsFlujoCajaRptObtieneDatosConceptos.BOF) Then
        For i = 1 To prsFlujoCajaRptObtieneDatosConceptos.RecordCount
            CuadroExcel xlsHoja, 1, lnFila, 3, lnFila
            xlsHoja.Cells(lnFila, 1) = prsFlujoCajaRptObtieneDatosConceptos!Descripcion
            xlsHoja.Range(xlsHoja.Cells(lnFila, 2), xlsHoja.Cells(lnFila, 200)).NumberFormat = "#,##0.00"
            xlsHoja.Cells(lnFila, 2) = prsFlujoCajaRptObtieneDatosConceptos!MontoHistorico
            xlsHoja.Cells(lnFila, 3) = prsFlujoCajaRptObtieneDatosConceptos!MontoMensual
            If prsFlujoCajaRptObtieneDatosConceptos!Descripcion = "INVERSION" Then
                lnFila = lnFila + 2
            Else
                lnFila = lnFila + 1
            End If
            CuadroExcel xlsHoja, 2, lnFila, 3, lnFila - 1
            prsFlujoCajaRptObtieneDatosConceptos.MoveNext
        Next i
    Else
        MsgBox "Hubo un Error Comuníquese con el Área de TI", vbInformation, "Aviso"
        Exit Sub
    End If
    
    'Calculos Finales: Totales->
    xlsHoja.Cells(11, 2) = "=SUM(B12:B16)" 'Ingresos Operativos - Hist.
    xlsHoja.Cells(11, 3) = "=SUM(C12:C16)" 'Ingresos Operativos - Mes
    
    xlsHoja.Cells(17, 2) = "=SUM(B18:B32)" 'Egresos Operativos - Hist.
    xlsHoja.Cells(17, 3) = "=SUM(C18:C32)" 'Egresos Operativos - Mes
    
    xlsHoja.Cells(33, 2) = "=(B11-B17)" 'Flujo Operativo - Hist.
    xlsHoja.Cells(33, 3) = "=(C11-C17)" 'Flujo Operativo - Mes
    
    xlsHoja.Cells(37, 2) = "=(B33+B34-B35-B36)" 'Flujo Financiero - Hist.
    xlsHoja.Cells(37, 3) = "=(C33+C34-C35-C36)" 'Flujo Financiero - Mes
    
    xlsHoja.Cells(40, 2) = "=(B37-B38)" 'Saldo - Hist.
    xlsHoja.Cells(40, 3) = "=(C37-C38)" 'Saldo - Mes
    
    xlsHoja.Cells(42, 2) = "=(B40+B41)" 'Saldo Acumulado - Hist.
    xlsHoja.Cells(42, 3) = "=(C40+C41)" 'Saldo Acumulado - Mes
        
    'Parametros
    If Not (prsFlujoCajaRptObtieneDatosParametros.EOF And prsFlujoCajaRptObtieneDatosParametros.BOF) Then
        lnIncVentCont = prsFlujoCajaRptObtieneDatosParametros!nIncVentCont
        lnIncCompMerc = prsFlujoCajaRptObtieneDatosParametros!nIncCompMerc
        lnIncPagPers = prsFlujoCajaRptObtieneDatosParametros!nIncPagPers
        lnIncConsu = prsFlujoCajaRptObtieneDatosParametros!nIncConsu
        lnIncGastvent = prsFlujoCajaRptObtieneDatosParametros!nIncGastvent
        
        xlsHoja.Cells(lnFila + 1, 1) = "DATOS ADICIONALES"
        xlsHoja.Range(xlsHoja.Cells(lnFila + 1, 1), xlsHoja.Cells(lnFila + 1, 1)).Font.Bold = True
        xlsHoja.Range(xlsHoja.Cells(lnFila + 1, 1), xlsHoja.Cells(lnFila + 1, 1)).HorizontalAlignment = xlCenter
        CuadroExcel xlsHoja, 1, lnFila + 1, 2, lnFila + 1
        xlsHoja.Cells(lnFila + 2, 1) = "Fecha de Pago"
        xlsHoja.Cells(lnFila + 2, 2) = Format(prsFlujoCajaRptObtieneDatosParametros!dFechaPago, "DD/MM/YYYY")
        CuadroExcel xlsHoja, 1, lnFila + 2, 2, lnFila + 2

        xlsHoja.Cells(lnFila + 4, 2) = "Mes"
        xlsHoja.Cells(lnFila + 4, 3) = "Anual"
        CuadroExcel xlsHoja, 2, lnFila + 4, 3, lnFila + 4
        xlsHoja.Range(xlsHoja.Cells(lnFila + 4, 2), xlsHoja.Cells(lnFila + 4, 3)).Font.Bold = True
        xlsHoja.Range(xlsHoja.Cells(lnFila + 4, 2), xlsHoja.Cells(lnFila + 4, 3)).HorizontalAlignment = xlCenter

        xlsHoja.Cells(lnFila + 5, 1) = "Incremento de ventas al contado "
        xlsHoja.Cells(lnFila + 6, 1) = "Incremento de Compra de Mercaderias"
        xlsHoja.Cells(lnFila + 7, 1) = "Incremento de Consumo"
        xlsHoja.Cells(lnFila + 8, 1) = "Incremento de Pago Personal"
        xlsHoja.Cells(lnFila + 9, 1) = "Ingremento de Gastos de Ventas"

        xlsHoja.Cells(lnFila + 5, 2) = Format(((1 + lnIncVentCont / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlsHoja.Cells(lnFila + 6, 2) = Format(((1 + lnIncCompMerc / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlsHoja.Cells(lnFila + 7, 2) = Format(((1 + lnIncConsu / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlsHoja.Cells(lnFila + 8, 2) = Format(((1 + lnIncPagPers / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlsHoja.Cells(lnFila + 9, 2) = Format(((1 + lnIncGastvent / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"

        xlsHoja.Cells(lnFila + 5, 3) = Format(lnIncVentCont, "#0.0") & "%"
        xlsHoja.Cells(lnFila + 6, 3) = Format(lnIncCompMerc, "#0.0") & "%"
        xlsHoja.Cells(lnFila + 7, 3) = Format(lnIncConsu, "#0.0") & "%"
        xlsHoja.Cells(lnFila + 8, 3) = Format(lnIncPagPers, "#0.0") & "%"
        xlsHoja.Cells(lnFila + 9, 3) = Format(lnIncGastvent, "#0.0") & "%"

        CuadroExcel xlsHoja, 1, lnFila + 5, 3, lnFila + 9, True
        CuadroExcel xlsHoja, 1, lnFila + 5, 3, lnFila + 9, False
        xlsHoja.Range(xlsHoja.Cells(lnFila + 5, 2), xlsHoja.Cells(lnFila + 9, 3)).HorizontalAlignment = xlCenter
'    Else
'        MsgBox "Hubo un Error Comuníquese con el Área de TI", vbInformation, "Aviso"
'        Exit Sub
    End If
  
  'Obtener las Letras del Abecedario A-Z
    Dim MatAZ As Variant
    Dim P As Integer
    P = 1
    Set MatAZ = Nothing
    ReDim MatAZ(1, 140)
    For i = 65 To 90
        MatAZ(1, P) = ChrW(i)
        P = P + 1
    Next i
           
    Dim MatLetrasRep As Variant
    Dim Y As Integer
    Set MatLetrasRep = Nothing
    Y = 1
    ReDim MatLetrasRep(1, 131)
    For A = 1 To 130
        If A <= 26 Then
                MatLetrasRep(1, Y) = ChrW(65) & MatAZ(1, Y) 'AA,AB,AC......AZ
            Y = Y + 1
        ElseIf (A >= 27 And A <= 52) Then
            If A = 27 Then
                P = 1
            End If
                MatLetrasRep(1, Y) = ChrW(66) & MatAZ(1, P) 'BA,BB,BC......BZ
            Y = Y + 1
            P = P + 1
        ElseIf (A >= 53 And A <= 78) Then
            If A = 53 Then
                P = 1
            End If
                MatLetrasRep(1, Y) = ChrW(67) & MatAZ(1, P) 'CA,CB,CC......CZ
            Y = Y + 1
            P = P + 1
        ElseIf (A >= 79 And A <= 104) Then
            If A = 79 Then
                P = 1
            End If
                MatLetrasRep(1, Y) = ChrW(68) & MatAZ(1, P) 'DA,DB,DC......DZ
            Y = Y + 1
            P = P + 1
        End If
    Next A
    
'Cuotas
i = 0
Y = 0
Z = 0
nCol = 4
nColInicio = 4
nColFin = 0
   If Not (prsFlujoCajaRptObtieneDatosCuotas.EOF And prsFlujoCajaRptObtieneDatosCuotas.BOF) Then
        For i = 1 To prsFlujoCajaRptObtieneDatosCuotas.RecordCount

            If i >= 24 Then
                Y = Y + 1
            End If

            xlsHoja.Cells(9, nCol) = prsFlujoCajaRptObtieneDatosCuotas!nCuota
            xlsHoja.Cells(10, nCol) = Format(prsFlujoCajaRptObtieneDatosCuotas!dFechaCuotas, "mmm-yyyy")

            '11: INGRESOS OPERATIVOS
            xlsHoja.Cells(11, nCol) = "=SUM(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "12" & ":" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "16)"
            '12:Ventas al contado
            xlsHoja.Cells(12, nCol) = Round((xlsHoja.Cells(12, nCol - 1) * ((1 + lnIncVentCont / 100) ^ (1 / 12) - 1) + xlsHoja.Cells(12, nCol - 1)), 1)
            '13:Cobros (por ventas al crédito)
            xlsHoja.Cells(13, nCol) = "=C13"
            '14:Cobros por ventas de activos fijos
            'xlsHoja.Cells(14, nCol) = "=C14"
            '15:Financiamiento
            'xlsHoja.Cells(15, nCol) = "=C15"
            '16: Otros Ingresos
            xlsHoja.Cells(16, nCol) = "=C16"
            
            '17:EGRESOS OPERATIVOS
            xlsHoja.Cells(17, nCol) = "=SUM(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "18" & ":" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "32)"
            '18:Egresos por compras
            xlsHoja.Cells(18, nCol) = Round((xlsHoja.Cells(17, nCol - 1) * ((1 + lnIncCompMerc / 100) ^ (1 / 12) - 1) + xlsHoja.Cells(17, nCol - 1)), 1)
            '19:Personal
            xlsHoja.Cells(19, nCol) = Round((xlsHoja.Cells(19, nCol - 1) * ((1 + lnIncPagPers / 100) ^ (1 / 12) - 1) + xlsHoja.Cells(19, nCol - 1)), 1)
            
            '20:Alquiler de locales
            xlsHoja.Cells(20, nCol) = "=C20"
            '21:Alquiler de equipos
            xlsHoja.Cells(21, nCol) = "=C21"
            '22:Servicios (Luz, Agua, Telefono, Cel.)
            xlsHoja.Cells(22, nCol) = "=C22"
            '23:Utiles de oficina
            xlsHoja.Cells(23, nCol) = "=C23"
            '24:Rep. y Matto. de equipos
            xlsHoja.Cells(24, nCol) = "=C24"
            '25:Rep. y Matto. de vehículos
            xlsHoja.Cells(25, nCol) = "=C25"
            '26:Seguros
            xlsHoja.Cells(26, nCol) = "=C26"
            '27:Transporte/Conbustible/Gas
            xlsHoja.Cells(27, nCol) = "=C27"
            '28: Contador
            xlsHoja.Cells(28, nCol) = "=C28"
            '29: Suntat + Impuestos
            xlsHoja.Cells(29, nCol) = "=C29"
            
            '30: Publicidad y otros gastos de ventas (**Nuevo)
            xlsHoja.Cells(30, nCol) = Round((xlsHoja.Cells(30, nCol - 1) * ((1 + lnIncGastvent / 100) ^ (1 / 12) - 1) + xlsHoja.Cells(30, nCol - 1)), 1)
            '31: Otros
            xlsHoja.Cells(31, nCol) = "=C31"
            '32:Consumo Per.Nat.
            xlsHoja.Cells(32, nCol) = Round((xlsHoja.Cells(32, nCol - 1) * ((1 + lnIncConsu / 100) ^ (1 / 12) - 1) + xlsHoja.Cells(32, nCol - 1)), 1)
            
            '33: FLUJO OPERATIVO
            xlsHoja.Cells(33, nCol) = "=(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "11" & "-" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "17)"
            '34:Cobro de Prestamo y dividendos
            xlsHoja.Cells(34, nCol) = "0.00"
            '35:Pago de cuotas préstamo vigentes
            xlsHoja.Cells(35, nCol) = "=C35"
            '36: Pago de cuotas prestamos solicitado
            xlsHoja.Cells(36, nCol) = "=C36"
            '37:FLUJO FINANCIERO
            xlsHoja.Cells(37, nCol) = "=(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "33" & "+" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "34" & "-" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "35" & "-" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "36)"
            '38:Inversion
            xlsHoja.Cells(38, nCol) = "0.00"
            
            '40:Saldo
            xlsHoja.Cells(40, nCol) = "=(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "37" & ")"
            If xlsHoja.Cells(40, nCol) < 0 Then  'Color de celda
                xlsHoja.Range(xlsHoja.Cells(40, nCol), xlsHoja.Cells(40, nCol)).Cells.Interior.Color = RGB(255, 145, 145)
            End If
            
            '41:Saldo Disponible
            If i >= 25 Then
                Z = Z + 1
            End If
            xlsHoja.Cells(41, nCol) = "=(" & IIf(i >= 25, MatLetrasRep(1, Z), MatAZ(1, i + 2)) & "42)"
            If xlsHoja.Cells(41, nCol) < 0 Then  'Color de celda
                xlsHoja.Range(xlsHoja.Cells(41, nCol), xlsHoja.Cells(41, nCol)).Cells.Interior.Color = RGB(255, 145, 145)
            End If
            
            '42:Saldo Acumulado
            xlsHoja.Cells(42, nCol) = "=(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "40" & "+" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "41)"
            If xlsHoja.Cells(42, nCol) < 0 Then  'Color de celda
                xlsHoja.Range(xlsHoja.Cells(42, nCol), xlsHoja.Cells(42, nCol)).Cells.Interior.Color = RGB(255, 145, 145)
            End If
            nCol = nCol + 1
    
            If (i Mod 12) = 0 Then
                    nColFin = nCol - 1
                        xlsHoja.Cells(8, nColInicio) = "Año" & (i / 12)
                        xlsHoja.Range(xlsHoja.Cells(8, nColInicio), xlsHoja.Cells(8, nColFin)).HorizontalAlignment = xlCenter
                        xlsHoja.Range(xlsHoja.Cells(8, nColInicio), xlsHoja.Cells(8, nColFin)).MergeCells = True
                        xlsHoja.Range(xlsHoja.Cells(8, nColInicio), xlsHoja.Cells(8, nColFin)).Font.Bold = True
                    nColInicio = nColFin + 1
           ' Else
           '     nColInicio = nColFin + 1
            End If
            prsFlujoCajaRptObtieneDatosCuotas.MoveNext
        Next i
    'Para la celda si no cumple un año
    xlsHoja.Range(xlsHoja.Cells(8, nColInicio), xlsHoja.Cells(8, nCol - 1)).MergeCells = True
    Else
        MsgBox "Hubo un Error Comuníquese con el Área de TI", vbInformation, "Aviso"
        Exit Sub
    End If

CuadroExcel xlsHoja, 4, 8, nCol - 1, 8, False 'Año
CuadroExcel xlsHoja, 1, 9, nCol - 1, 38, False
CuadroExcel xlsHoja, 1, 10, nCol - 1, 38, False
CuadroExcel xlsHoja, 1, 11, nCol - 1, 38, True
CuadroExcel xlsHoja, 1, 40, nCol - 1, 42, True

CuadroExcel xlsHoja, 1, 11, nCol - 1, 11, True 'Gris
CuadroExcel xlsHoja, 1, 17, nCol - 1, 17, True 'Gris
CuadroExcel xlsHoja, 1, 33, nCol - 1, 33, True 'Gris
CuadroExcel xlsHoja, 1, 37, nCol - 1, 37, True 'Gris

xlsHoja.Range(xlsHoja.Cells(8, 4), xlsHoja.Cells(9, nCol - 1)).Cells.Interior.Color = RGB(238, 248, 255) 'Celeste
xlsHoja.Range(xlsHoja.Cells(9, 1), xlsHoja.Cells(10, nCol - 1)).Cells.Interior.Color = RGB(238, 248, 255) 'Celeste
xlsHoja.Range(xlsHoja.Cells(44, 1), xlsHoja.Cells(44, 2)).Cells.Interior.Color = RGB(238, 248, 255) 'Celeste
xlsHoja.Range(xlsHoja.Cells(47, 2), xlsHoja.Cells(47, 3)).Cells.Interior.Color = RGB(238, 248, 255) 'Celeste

xlsHoja.Range(xlsHoja.Cells(11, 1), xlsHoja.Cells(11, nCol - 1)).Cells.Interior.Color = RGB(220, 220, 220) ' Gris
xlsHoja.Range(xlsHoja.Cells(17, 1), xlsHoja.Cells(17, nCol - 1)).Cells.Interior.Color = RGB(220, 220, 220) ' Gris
xlsHoja.Range(xlsHoja.Cells(33, 1), xlsHoja.Cells(33, nCol - 1)).Cells.Interior.Color = RGB(220, 220, 220) ' Gris
xlsHoja.Range(xlsHoja.Cells(37, 1), xlsHoja.Cells(37, nCol - 1)).Cells.Interior.Color = RGB(220, 220, 220) ' Gris

xlsHoja.Range(xlsHoja.Cells(11, 1), xlsHoja.Cells(11, nCol - 1)).Font.Bold = True 'Fondo Gris
xlsHoja.Range(xlsHoja.Cells(17, 1), xlsHoja.Cells(17, nCol - 1)).Font.Bold = True 'Fondo Gris
xlsHoja.Range(xlsHoja.Cells(33, 1), xlsHoja.Cells(33, nCol - 1)).Font.Bold = True 'Fondo Gris
xlsHoja.Range(xlsHoja.Cells(37, 1), xlsHoja.Cells(37, nCol - 1)).Font.Bold = True 'Fondo Gris

xlsHoja.Range(xlsHoja.Cells(12, 1), xlsHoja.Cells(12, nCol - 1)).Font.Bold = True 'Calculo Ventas Contado
xlsHoja.Range(xlsHoja.Cells(18, 1), xlsHoja.Cells(18, nCol - 1)).Font.Bold = True 'Calculo Egreso Compras
xlsHoja.Range(xlsHoja.Cells(19, 1), xlsHoja.Cells(19, nCol - 1)).Font.Bold = True 'Calculo Personal
xlsHoja.Range(xlsHoja.Cells(32, 1), xlsHoja.Cells(32, nCol - 1)).Font.Bold = True 'Calculo Consumo
xlsHoja.Range(xlsHoja.Cells(30, 1), xlsHoja.Cells(30, nCol - 1)).Font.Bold = True 'Calculo Inc. Gasto Ventas

xlsHoja.Cells.Select
xlsHoja.Cells.Font.Name = "Arial"
xlsHoja.Cells.Font.Size = 9
xlsHoja.Cells.EntireColumn.AutoFit
End Sub
'<-***** Fin LUCV20171015
