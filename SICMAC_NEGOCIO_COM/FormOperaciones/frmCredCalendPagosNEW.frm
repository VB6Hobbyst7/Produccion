VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredCalendPagosNEW 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendario de Pagos"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11055
   Icon            =   "frmCredCalendPagosNEW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTasaAnuales 
      Caption         =   " &Tasas Anuales "
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
      Height          =   1550
      Left            =   9120
      TabIndex        =   73
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Label lblTasaCostoEfectivoAnual 
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
         Height          =   255
         Left            =   360
         TabIndex        =   77
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblTCEA 
         Caption         =   "Tasa Costo Efectivo Anual"
         Height          =   495
         Left            =   120
         TabIndex        =   76
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblTasaEfectivaAnual 
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
         Height          =   255
         Left            =   360
         TabIndex        =   75
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblTEA 
         Caption         =   "Tasa Efectiva Anual"
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame FraPigno 
      Caption         =   " &Pignoraticios"
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
      Height          =   1910
      Left            =   9120
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      Begin VB.TextBox txtInteresPigno 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         MaxLength       =   9
         TabIndex        =   72
         Top             =   1520
         Width           =   1050
      End
      Begin VB.TextBox txtTotalPigno 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         MaxLength       =   9
         TabIndex        =   70
         Top             =   1010
         Width           =   1050
      End
      Begin MSMask.MaskEdBox txtFechaVencPigno 
         Height          =   315
         Left            =   240
         TabIndex        =   68
         ToolTipText     =   "Presione Enter"
         Top             =   480
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label16 
         Caption         =   "Interes"
         Height          =   200
         Left            =   480
         TabIndex        =   71
         Top             =   1310
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Total"
         Height          =   200
         Left            =   480
         TabIndex        =   69
         Top             =   800
         Width           =   495
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha Vencimiento"
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblFechaVenc 
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
         Height          =   255
         Left            =   360
         TabIndex        =   66
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmbSubProducto 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   525
      Width           =   2300
   End
   Begin VB.Frame Frame5 
      Height          =   3585
      Left            =   0
      TabIndex        =   34
      Top             =   3720
      Width           =   10995
      Begin SICMACT.FlexEdit FECalend 
         Height          =   3165
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   5689
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-Fecha Venc.-Cuota-Cuota + ITF-Cuotas-Capital-Interes-Int. Gracia-Gatos/Comis-Seg. Desg-Seg. Mult.-Saldo Capital"
         EncabezadosAnchos=   "0-1000-600-1000-1100-1000-1000-1000-1000-1000-1000-1200"
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
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-R-R-R-R-R-R-R-R-C"
         FormatosEdit    =   "0-0-0-2-2-2-2-2-2-2-2-2"
         SelectionMode   =   1
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
      End
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   33
      ToolTipText     =   "Generar el Calendario de Pagos"
      Top             =   7785
      Width           =   1455
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4155
      TabIndex        =   32
      ToolTipText     =   "Imprimir el Calendario de Pagos"
      Top             =   7785
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   31
      ToolTipText     =   "Salir del Calendario de Pagos"
      Top             =   7785
      Width           =   1455
   End
   Begin VB.CommandButton cmdResumen 
      Caption         =   "&Resumen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7650
      TabIndex        =   0
      ToolTipText     =   "Resumen del Calendario de Pagos"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin TabDlg.SSTab SSCalend 
      Height          =   3585
      Left            =   240
      TabIndex        =   36
      Top             =   3720
      Visible         =   0   'False
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   6324
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Buen Pagador"
      TabPicture(0)   =   "frmCredCalendPagosNEW.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FECalBPag"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mal Pagador"
      TabPicture(1)   =   "frmCredCalendPagosNEW.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FECalMPag"
      Tab(1).ControlCount=   1
      Begin SICMACT.FlexEdit FECalBPag 
         Height          =   2880
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   5080
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-Fecha Venc.-Cuota-Cuotas-Capital-Interes-Int. Gracia-Gastos-Seg Desg-Seg Inmueb-Saldo Capital-Cuota + ITF"
         EncabezadosAnchos=   "400-1000-600-1200-1000-1000-1000-1000-1000-1000-1200-1000"
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
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-R-R-R-R-R-R-R-R-C"
         FormatosEdit    =   "0-0-0-2-3-2-2-2-2-2-2-2"
         SelectionMode   =   1
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin SICMACT.FlexEdit FECalMPag 
         Height          =   2880
         Left            =   -74760
         TabIndex        =   38
         Top             =   360
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   5080
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-Fecha Venc.-Cuota-Cuotas-Capital-Interes-Int. Gracia-Gastos-Seg Desg-Seg Inmueb-Saldo Capital-Cuota + ITF"
         EncabezadosAnchos=   "400-1000-600-1200-1000-1000-1000-1000-1000-1000-1200-1200"
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
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-R-R-R-R-R-R-R-R-R"
         FormatosEdit    =   "0-0-0-2-3-2-2-2-2-2-2-2"
         SelectionMode   =   1
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
   End
   Begin VB.Frame FraDatos 
      Caption         =   "Condiciones"
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
      Height          =   3675
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9015
      Begin VB.TextBox txtTEM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3050
         MaxLength       =   7
         TabIndex        =   81
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtInteresAnual 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   78
         Top             =   1200
         Width           =   615
      End
      Begin VB.Frame FraHipoteca 
         Height          =   1335
         Left            =   6480
         TabIndex        =   58
         Top             =   1080
         Width           =   2390
         Begin VB.TextBox txtEdificacion 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            MaxLength       =   7
            TabIndex        =   80
            Top             =   840
            Width           =   2150
         End
         Begin VB.CheckBox chkHipoteca 
            Caption         =   "¿Tiene Hipoteca?"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   0
            Width           =   1935
         End
         Begin VB.ComboBox cmbHipoteca 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   240
            Width           =   2150
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Valor Edificación:"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   600
            Width           =   1230
         End
      End
      Begin VB.Frame FraTpoCliente 
         Caption         =   "Calificación Cliente"
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
         Height          =   650
         Left            =   6480
         TabIndex        =   56
         Top             =   240
         Width           =   2390
         Begin VB.ComboBox cmbTpoCliente 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   240
            Width           =   2150
         End
      End
      Begin VB.ComboBox cmbMoneda 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   840
         Width           =   1325
      End
      Begin VB.ComboBox cmbProductoCMACM 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   240
         Width           =   2300
      End
      Begin VB.Frame FraFechaPago 
         Height          =   650
         Left            =   3960
         TabIndex        =   48
         Top             =   1110
         Width           =   2390
         Begin MSMask.MaskEdBox txtFechaPago 
            Height          =   315
            Left            =   1250
            TabIndex        =   49
            ToolTipText     =   "Presione Enter"
            Top             =   280
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha de Pago"
            Height          =   255
            Left            =   50
            TabIndex        =   50
            Top             =   280
            Width           =   1095
         End
      End
      Begin VB.Frame FraTipoPeriodo 
         Caption         =   "Tipo Periodo"
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
         Height          =   890
         Left            =   3960
         TabIndex        =   17
         Top             =   240
         Width           =   2390
         Begin VB.OptionButton OptTipoPeriodo 
            Caption         =   "Fecha Fija"
            Height          =   315
            Index           =   1
            Left            =   150
            TabIndex        =   47
            Top             =   540
            Width           =   1035
         End
         Begin VB.OptionButton OptTipoPeriodo 
            Caption         =   "Periodo Fijo"
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   21
            Top             =   285
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.Frame Frame6 
            Height          =   550
            Left            =   1320
            TabIndex        =   18
            Top             =   240
            Width           =   1000
            Begin VB.TextBox TxtDiaFijo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   615
               MaxLength       =   2
               TabIndex        =   19
               Top             =   150
               Width           =   330
            End
            Begin VB.Label LblDia 
               AutoSize        =   -1  'True
               Caption         =   "&Dia 1:"
               Enabled         =   0   'False
               Height          =   195
               Left            =   90
               TabIndex        =   20
               Top             =   180
               Width           =   420
            End
         End
      End
      Begin VB.Frame FraGracia 
         Height          =   945
         Left            =   3960
         TabIndex        =   10
         Top             =   1725
         Width           =   2390
         Begin VB.TextBox TxtTasaGraciaNEW 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1380
            MaxLength       =   7
            TabIndex        =   79
            Top             =   500
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CheckBox ChkPerGra 
            Caption         =   "Periodo &Gracia"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   210
            Width           =   1350
         End
         Begin VB.TextBox TxtPerGra 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1590
            MaxLength       =   4
            TabIndex        =   13
            Text            =   "0"
            Top             =   210
            Width           =   390
         End
         Begin VB.TextBox TxtTasaGracia 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1380
            MaxLength       =   7
            TabIndex        =   12
            Top             =   500
            Width           =   615
         End
         Begin VB.OptionButton optTipoGracia 
            Caption         =   "Capitalizar"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   660
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label LblTasaGracia 
            Caption         =   "Tasa :"
            Height          =   165
            Left            =   870
            TabIndex        =   16
            Top             =   500
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label LblPorcGracia 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2040
            TabIndex        =   15
            Top             =   500
            Visible         =   0   'False
            Width           =   150
         End
      End
      Begin VB.TextBox TxtInteres 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   9
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox TxtMonto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.Frame fraGastoCom 
         Caption         =   "Gastos/Comisión"
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
         Height          =   900
         Left            =   120
         TabIndex        =   2
         Top             =   2700
         Visible         =   0   'False
         Width           =   6225
         Begin VB.Frame Frame2 
            Height          =   615
            Left            =   0
            TabIndex        =   62
            Top             =   240
            Width           =   2415
            Begin VB.CheckBox chkSegDesgra 
               Caption         =   "Seg. Desgravamen:"
               Height          =   255
               Left            =   120
               TabIndex        =   64
               Top             =   0
               Width           =   1935
            End
            Begin VB.ComboBox cmbSeguroDes 
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   63
               Top             =   240
               Width           =   2220
            End
         End
         Begin VB.Frame fraEnvioEst 
            Height          =   615
            Left            =   2520
            TabIndex        =   3
            Top             =   240
            Width           =   2415
            Begin VB.ComboBox cmbEnvioEst 
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   240
               Width           =   2220
            End
            Begin VB.CheckBox chkEnvioEst 
               Caption         =   "Envío Estado Cuenta"
               Height          =   255
               Left            =   120
               TabIndex        =   4
               Top             =   0
               Width           =   1935
            End
         End
      End
      Begin Spinner.uSpinner SpnCuotas 
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   1560
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   450
         Max             =   300
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
      Begin MSComCtl2.DTPicker DTFecDesemb 
         Height          =   300
         Left            =   1440
         TabIndex        =   7
         Top             =   2265
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   529
         _Version        =   393216
         Format          =   269484033
         CurrentDate     =   37054
      End
      Begin Spinner.uSpinner SpnPlazo 
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   1905
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   503
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
      Begin VB.Frame FraTipoCuota 
         Caption         =   "Tipo Cuota"
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
         Height          =   550
         Left            =   3960
         TabIndex        =   22
         Top             =   240
         Width           =   2385
         Begin VB.OptionButton OptTipoCuota 
            Caption         =   "Fijo"
            Height          =   255
            Index           =   0
            Left            =   19120
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.Label lblTEM 
         AutoSize        =   -1  'True
         Caption         =   "T.E.M"
         Height          =   195
         Left            =   2400
         TabIndex        =   82
         Top             =   1245
         Width           =   435
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   525
         Width           =   645
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Categoría"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblmonto 
         AutoSize        =   -1  'True
         Caption         =   "&Monto"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   885
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "T.E.A"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   1245
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro Cuotas"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   1605
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Periodo (Dias)"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   1935
         Width           =   990
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha Desembolso"
         Height          =   435
         Left            =   120
         TabIndex        =   26
         Top             =   2235
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2100
         TabIndex        =   25
         Top             =   1230
         Width           =   150
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Interes + Int. Gracia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   2220
      TabIndex        =   46
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6120
      TabIndex        =   45
      Top             =   7320
      Width           =   1410
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5610
      TabIndex        =   44
      Top             =   7320
      Width           =   435
   End
   Begin VB.Label lblInteres 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4020
      TabIndex        =   43
      Top             =   7320
      Width           =   1410
   End
   Begin VB.Label lblCapital 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   660
      TabIndex        =   42
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Capital"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   0
      TabIndex        =   41
      Top             =   7320
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total+ITF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7845
      TabIndex        =   40
      Top             =   7320
      Width           =   840
   End
   Begin VB.Label lblTotalCONITF 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   8730
      TabIndex        =   39
      Top             =   7320
      Width           =   1410
   End
End
Attribute VB_Name = "frmCredCalendPagosNEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCredCalendPagosNEW
'***     Descripcion:       Simulador de Calendario de Pagos a diferentes condiciones de pago
'***     Creado por:        ARLO
'***     Maquina:           01A-DETI-05
'***     Fecha-Tiempo:         23/06/2018 12:03:12 PM
'***     Ultima Modificacion: Creacion de la Opcion
'*****************************************************************************************

Option Explicit
Dim nTipoGracia As Integer
Dim MatGracia As Variant
Dim psCtaCod As String
Dim lsCtaCodLeasing As String
Dim bGraciaGenerada As Boolean
Dim MatCalend As Variant
Dim MatCalend_2 As Variant
Dim MatResul As Variant
Dim MatResulDiff As Variant
Dim MatCalendDesParcial As Variant
Dim bTipoDesembolso As TCalendTipoDesemb
Dim bDesembParcialGenerado As Boolean
Dim nSugerAprob As Integer '1 suegerencia 2 aprobacion
Dim bMuestraGastos As Boolean
Dim bDesemParcial As Boolean
Dim MatDesPar As Variant
Dim cCtaCodG As String
Dim lnLeasing As Integer
'Enum TCalendTipoDesemb
'    DesembolsoTotal = 0
'    DesembolsoParcial = 1
'End Enum

' Para evitar la Caida en el caso que surjan Errores de Validacion en
' la Generacion del Calendario
Dim bErrorValidacion As Boolean
'Para trabajar los Gastos con Componentes
'Dim MatGastos As Variant
'Dim nNumgastos As Integer

'Para almacenar el valor del Capital antes de Capitalizar la Gracia
Dim nMontoCapInicial As Double
Dim bRenovarCredito As Boolean
Dim nInteresAFecha As Double
Dim nTasaCostoEfectivoAnual As Double
Dim nTasaEfectivaAnual As Double
Dim nCuotMensBono As Double
Dim nCuotMens As Double
Private MatGastos As Variant
Private MatDesemb As Variant
Private sCtaCodRep As String
Dim nTotalcuotasLeasing As Currency
Dim nIntGraInicial As Double
Dim lbLogicoBF As Integer
Dim ldFechaBF As Date
Dim lnMontoMivivienda As Currency
Dim lnCuotaMivienda As Integer
Private fbInicioSim As Boolean
Dim nTpoSubPro, nTpoPro, nTpoCliente, nTpoHipo, nTotalPigno, nInteresPigno As Integer
Dim cTpoPro, cTpoSubPro, cTpoMoneda, cTpoCliente, cTpoHipo As String
Dim dFechaPagoPigno As Date
Dim nTasaNormal, nTasaCPP, nTEA, nLocalHipo, nCasaHabiHipo, nSegMult As Double
Dim nSegMutlTotal, nSegMutlUnidad, nTem As Double
Dim bProxMesN, bSimulador As Boolean
Dim nMultiplicaSegMutl, pnDiaFijo2 As Integer
Dim bProxMes, bExisteModal As Boolean
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018

Public Function Renovar(ByVal nMonto As Double, ByVal dFecDes As Date, ByVal nTasa As Double, _
                        ByVal pnInteresAFecha As Double, Optional psCtaCod As String) As Variant

    TxtMonto.Text = Format(nMonto, "#0.00")
    TxtInteres.Text = Format(nTasa, "#0.00")
    DTFecDesemb.value = Format(dFecDes, "dd/mm/yyyy")
    
    bRenovarCredito = True
    nInteresAFecha = pnInteresAFecha
    sCtaCodRep = psCtaCod
    TxtMonto.Enabled = False
    TxtInteres.Enabled = False
    DTFecDesemb.Enabled = False
    Me.Show 1
    
    bRenovarCredito = False
    nInteresAFecha = 0
    Renovar = MatCalend
    
End Function

Public Sub Simulacion(ByVal pTipoSimulacion As TCalendTipoDesemb)
    DTFecDesemb.value = Format(gdFecSis, "dd/mm/yyyy")
    Set MatCalendDesParcial = Nothing
    If pTipoSimulacion = DesembolsoTotal Then
        TxtMonto.Text = "0.00"
        TxtMonto.Enabled = True
        SpnCuotas.valor = 1
        SpnCuotas.Enabled = True
        bDesemParcial = False
    Else
        bDesemParcial = True
        TxtMonto.Text = "0.00"
        TxtMonto.Enabled = False
        SpnCuotas.valor = 1
        SpnCuotas.Enabled = False
    End If
    bTipoDesembolso = pTipoSimulacion
    Me.Show 1
End Sub

Public Sub SoloMuestraMatrices(ByVal pMatCalend As Variant, ByVal pMatResul As Variant, ByVal MatGastos As Variant, ByVal nNumGasto As Integer, _
                ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Integer, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnTasaGracia As Double, ByVal pnDiaFijo As Integer, _
                ByVal bProxMes As Boolean, ByVal pMatGracia As Variant, ByVal pnMiViv As Integer, _
                ByVal pnCuotaCom As Integer, ByRef MatMiViv_2 As Variant, Optional ByVal pnSugerAprob As Integer = 0, _
                Optional ByVal pbNoMostrarCalendario As Boolean = False, _
                Optional ByVal pnDiaFijo2 As Integer = 0, Optional ByVal pbIncrementaCapital As Boolean = False, _
                Optional ByRef pnTasCosEfeAnu As Double, Optional ByRef psCtaCodLeasing As String = "", Optional pbLogicoBF As Integer = 0, Optional pdFechaBF As Date = CDate("1900-01-01")) ' DAOR 20070419, pnTasCosEfeAnu:Tasa Costo Efectivo Anual

Dim i, j As Integer
Dim nTotalInteres As Double
Dim nTotalCapital As Double
Dim nTotalGasto As Double
Dim nTotalGastoSeg As Double
Dim lnSalCap As Double
Dim oCredito As COMNCredito.NCOMCredito
Dim nRedondeoITF As Double
Dim nTotalcuotasCONItF As Double

        nSugerAprob = pnSugerAprob
        TxtMonto.Text = Format(pnMonto, "#0.00")
        TxtInteres.Text = Format(pnTasaInt, "#0.0000")
        SpnCuotas.valor = Trim(str(pnNroCuotas))
        SpnPlazo.valor = Trim(str(pnPeriodo))
        lbLogicoBF = pbLogicoBF
        ldFechaBF = pdFechaBF
        
        ChkPerGra.value = IIf(pnDiasGracia <> "0", 1, 0)
        
        DTFecDesemb.value = Format(pdFecDesemb, "dd/mm/yyyy")
        OptTipoCuota(pnTipoCuota - 1).value = True
        OptTipoPeriodo(pnTipoPeriodo - 1).value = True
        If pnTipoPeriodo = FechaFija Then
            TxtDiaFijo.Text = Trim(str(pnDiaFijo))
             '   ChkProxMes.value = IIf(bProxMes, 1, 0)
            'Se agrego para manejar la opcion de 2 dias fijos
            'TxtDiaFijo2.Text = Trim(str(pnDiaFijo2))
        End If
        nTipoGracia = pnTipoGracia
        TxtPerGra.Text = Trim(str(pnDiasGracia))
        TxtTasaGracia.Text = Format(pnTasaGracia, "#0.0000")
        Set MatGracia = Nothing
        MatGracia = pMatGracia
        cmdAplicar.Enabled = False
        FraDatos.Enabled = False
        FraFechaPago.Enabled = False
        'ChkCuotaCom.value = pnCuotaCom
        
        'Cambios para las opciones de gracia
        If pnTipoGracia = EnCuotas - 1 Then
            optTipoGracia(1).value = True
        End If
        If pnTipoGracia = Capitalizada - 1 Then
            optTipoGracia(0).value = True
        End If

        If Len(Trim(psCtaCodLeasing)) = 0 Then
            If IsArray(MatGastos) Then
                For j = 0 To UBound(pMatCalend) - 1
                    nTotalGasto = 0
                    nTotalGastoSeg = 0
                    For i = 0 To UBound(MatGastos) - 1
                        
                        If Trim(Right(MatGastos(i, 0), 2)) = "1" And _
                           (Trim(MatGastos(i, 1)) = Trim(pMatCalend(j, 1)) And Trim(Right(MatGastos(i, 2), 4)) = "1217") Then
                            nTotalGastoSeg = nTotalGastoSeg + CDbl(MatGastos(i, 3))
                            pMatCalend(j, 6) = Format(nTotalGastoSeg, "#0.00")
                        ElseIf (Trim(Right(MatGastos(i, 0), 2)) = "1" And Trim(MatGastos(i, 1)) = Trim(pMatCalend(j, 1)) And Trim(Right(MatGastos(i, 2), 4)) = "1272") Then
                            pMatCalend(j, 14) = Format(CDbl(MatGastos(i, 3)), "#0.00")
                        Else
                            If Trim(MatGastos(i, 1)) = "*" Or (Trim(Right(MatGastos(i, 0), 2)) = "1" And Trim(MatGastos(i, 1)) = Trim(pMatCalend(j, 1)) And Trim(Right(MatGastos(i, 2), 4)) <> "1217") Then
                                nTotalGasto = nTotalGasto + CDbl(MatGastos(i, 3))
                                pMatCalend(j, 8) = Format(nTotalGasto, "#0.00")
                            End If
                        End If
                    Next i
                Next j
            End If
        End If
        
            nTotalInteres = 0
            nTotalCapital = 0
            LimpiaFlex FECalend
            For i = 0 To UBound(pMatCalend) - 1
                FECalend.AdicionaFila
                FECalend.TextMatrix(i + 1, 1) = Trim(pMatCalend(i, 0))
                FECalend.TextMatrix(i + 1, 2) = Trim(pMatCalend(i, 1))
                
                '**ARLO20180712 ERS042 - 2018
                Set objProducto = New COMDCredito.DCOMCredito
                If objProducto.GetResultadoCondicionCatalogo("N0000089", Mid(psCtaCodLeasing, 6, 3)) Then
                'If Mid(psCtaCodLeasing, 6, 3) = "515" Or Mid(psCtaCodLeasing, 6, 3) = "516" Then
                '**ARLO20180712 ERS042 - 2018
                    pMatCalend(i, 2) = Format(CDbl(pMatCalend(i, 2)), "#0.00")
                Else
                    pMatCalend(i, 2) = Format(CDbl(pMatCalend(i, 2)) + CDbl(pMatCalend(i, 6)) + CDbl(pMatCalend(i, 8)), "#0.00")
                End If

                If nTipoGracia = 6 Then
                    pMatCalend(i, 2) = Trim(CDbl(pMatCalend(i, 2)) + CDbl(pMatCalend(i, 11)))
                    FECalend.TextMatrix(i + 1, 3) = Format(Trim(CDbl(pMatCalend(i, 2)) + CDbl(pMatCalend(i, 14))), "#0.00")
                Else
                    FECalend.TextMatrix(i + 1, 3) = Trim(CDbl(pMatCalend(i, 2)) + CDbl(pMatCalend(i, 14)))
                End If
                
                FECalend.row = i + 1
                FECalend.col = 3
                FECalend.CellForeColor = vbBlue
                FECalend.TextMatrix(i + 1, 4) = Trim(pMatCalend(i, 3))
                
                FECalend.TextMatrix(i + 1, 5) = Trim(pMatCalend(i, 4))
                
                FECalend.TextMatrix(i + 1, 6) = Trim(pMatCalend(i, 5))

                If Len(Trim(psCtaCodLeasing)) = 18 Then
                        FECalend.TextMatrix(i + 1, 7) = Trim(pMatCalend(i, 6))
                        FECalend.TextMatrix(i + 1, 8) = Format(Trim(pMatCalend(i, 8)), "#0.00")
                Else
                        FECalend.TextMatrix(i + 1, 7) = Trim(pMatCalend(i, 8))
                        FECalend.TextMatrix(i + 1, 8) = Trim(pMatCalend(i, 6))
                End If
                FECalend.TextMatrix(i + 1, 9) = Trim(pMatCalend(i, 14))
                FECalend.TextMatrix(i + 1, 10) = Trim(pMatCalend(i, 7))

                If Not (i = 0 And nTipoGracia = 6) Then
                    nTotalCapital = nTotalCapital + CDbl(Trim(pMatCalend(i, 3)))
                    nTotalInteres = nTotalInteres + CDbl(Trim(pMatCalend(i, 4))) + CDbl(Trim(pMatCalend(i, 5)))
                End If

                nRedondeoITF = fgDiferenciaRedondeoITF(fgITFCalculaImpuesto(CDbl(pMatCalend(i, 2))))
                If nRedondeoITF > 0 Then
                    FECalend.TextMatrix(i + 1, 11) = Format(CDbl(pMatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(pMatCalend(i, 2))) - nRedondeoITF, "0.00")
                Else
                    FECalend.TextMatrix(i + 1, 11) = Format(CDbl(pMatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(pMatCalend(i, 2))), "0.00")
                End If
            
            If Not (pnTipoGracia = gColocTiposGraciaCapitalizada And i = 0) Then
                nTotalcuotasCONItF = nTotalcuotasCONItF + CDbl(FECalend.TextMatrix(i + 1, 11))
            End If
            
            Next i
            
            If pnTipoGracia = 6 Then
                FECalend.TextMatrix(1, 3) = ""
                FECalend.TextMatrix(1, 5) = ""
                FECalend.TextMatrix(1, 7) = ""
                FECalend.TextMatrix(1, 8) = ""
                FECalend.TextMatrix(1, 11) = ""
            End If
            
            lblCapital.Caption = Format(nTotalCapital, "#0.00")
            
            lblInteres.Caption = Format(nTotalInteres, "#0.00")
            lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#0.00")
            FECalend.row = 1
            FECalend.TopRow = 1

        fraTasaAnuales.Visible = True
        Set oCredito = New COMNCredito.NCOMCredito
            nTasaEfectivaAnual = Round(oCredito.TasaIntPerDias(pnTasaInt, 360) * 100, 2)
        If pnTipoGracia = 6 Then
            Dim Y As Integer
            Dim MatCalendTemp() As String
            ReDim MatCalendTemp(UBound(pMatCalend) - 1, 13)
            For i = 0 To UBound(pMatCalend) - 2
                For Y = 0 To 13
                    MatCalendTemp(i, Y) = pMatCalend(i + 1, Y)
                Next Y
            Next i
            Erase pMatCalend
            ReDim pMatCalend(UBound(MatCalendTemp), 13)
            
            For i = 0 To UBound(MatCalendTemp)
                For Y = 0 To 13
                    pMatCalend(i, Y) = MatCalendTemp(i, Y)
                Next Y
            Next i
            Erase MatCalendTemp
        End If
            
            nTasaCostoEfectivoAnual = oCredito.GeneraTasaCostoEfectivoAnual(pdFecDesemb, IIf(nTpoSubPro = 301, pnMonto, pnMonto - 12500), pMatCalend, pnTasaInt, lsCtaCodLeasing, pnTipoPeriodo)
            lblTasaCostoEfectivoAnual.Caption = nTasaCostoEfectivoAnual & " %"
            lblTasaEfectivaAnual.Caption = nTasaEfectivaAnual & " %"
            
            pnTasCosEfeAnu = nTasaCostoEfectivoAnual
        Set oCredito = Nothing
        If UBound(pMatCalend) = 0 Then
            cmdImprimir.Enabled = False
        Else
            cmdImprimir.Enabled = True
        End If
        MatCalend = pMatCalend
        MatResul = pMatResul
        
        '**ARLO20180712 ERS042 - 2018
        Set objProducto = New COMDCredito.DCOMCredito
        If objProducto.GetResultadoCondicionCatalogo("N0000090", Mid(psCtaCodLeasing, 6, 3)) Then
        'If Mid(psCtaCodLeasing, 6, 3) = "515" Or Mid(psCtaCodLeasing, 6, 3) = "516" Then
         '**ARLO20180712 ERS042 - 2018
            If nTotalcuotasLeasing > 0 Then
                nTotalcuotasLeasing = Format(nTotalcuotasLeasing + fgITFCalculaImpuesto(CDbl(nTotalcuotasLeasing)), "0.00")
            Else
                lblTotalCONITF.Caption = Format(CDbl(lblTotal.Caption) + fgITFCalculaImpuesto(CDbl(lblTotal.Caption)), "0.00")
            End If
        Else
            lblTotalCONITF.Caption = Format(nTotalcuotasCONItF, "#0.00")
        End If
        
        Me.Show 1
End Sub

'Modificado CACV para trabajar los Gastos con los Componentes
Public Function Inicio(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Integer, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnTasaGracia As Double, ByVal pnDiaFijo As Integer, _
                ByVal bProxMes As Boolean, ByVal pMatGracia As Variant, ByVal pnMiViv As Integer, _
                ByVal pnCuotaCom As Integer, ByRef MatMiViv_2 As Variant, Optional ByVal pnSugerAprob As Integer = 0, _
                Optional ByVal pbNoMostrarCalendario As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
                Optional ByVal pMatDesPar As Variant = "", Optional ByVal bQuincenal As Boolean, Optional ByVal psCtaCod As String, _
                Optional ByRef pbErrorValidacion As Boolean = False, Optional ByVal pnDiaFijo2 As Integer = 0, _
                Optional ByVal pbIncrementaCapital As Boolean = False, _
                Optional ByVal bGracia As Boolean = False, Optional ByVal dFechaPago As Date, Optional pnLeasing As Integer = 0, Optional psCredCodLeasing As String = "", _
                Optional ByVal pnValorInmueb As Double, Optional ByRef pnIntGraInicial As Double = 0, Optional ByVal pnCuotaBalon As Integer = 0, _
                Optional pbLogicoBF As Boolean = False, Optional pdFechaBF As Date = CDate("1900-01-01"), Optional pnMontoMivivienda As Currency = 0#, _
                Optional pnCuotaMivienda As Integer, Optional ByVal pArrMIVIVIENDA As Variant) As Variant
        Dim oParam As COMDCredito.DCOMParametro
        Set oParam = New COMDCredito.DCOMParametro
        Dim nTramoNoConsMonto As Double
        Dim nTramoConsMonto As Double
        Dim nTramoNoConsPorcen As Double
        lnCuotaMivienda = pnCuotaMivienda
        'ChkCalMiViv.value = pnMiViv
        lnMontoMivivienda = pnMontoMivivienda
        lnLeasing = pnLeasing
        lsCtaCodLeasing = psCredCodLeasing
        lbLogicoBF = pbLogicoBF
        ldFechaBF = pdFechaBF
        lnMontoMivivienda = pnMontoMivivienda
        'txtCuotaInicial.Text = Format(lnMontoMivivienda - pnMonto, "#0.00")
        'Me.txtValorInmueble.Text = Format(lnMontoMivivienda, "#0.00")
        'txtBonoBuenPagador.Text = Format(nTramoNoConsPorcen, "#0.00")
        bDesemParcial = pbDesemParcial
        MatDesPar = pMatDesPar
        cCtaCodG = psCtaCod
        nSugerAprob = pnSugerAprob
        TxtMonto.Text = Format(pnMonto, "#0.00")
        TxtInteres.Text = Format(pnTasaInt, "#0.0000")
        SpnCuotas.valor = Trim(str(pnNroCuotas))
        SpnPlazo.valor = Trim(str(pnPeriodo))
        TxtPerGra.Text = Trim(str(pnDiasGracia))
        DTFecDesemb.value = Format(pdFecDesemb, "dd/mm/yyyy")
        
        OptTipoCuota(pnTipoCuota - 1).value = True
        OptTipoPeriodo(pnTipoPeriodo - 1).value = True
        If pnTipoPeriodo = FechaFija Then
            TxtDiaFijo.Text = Trim(str(pnDiaFijo))
            '    ChkProxMes.value = IIf(bProxMes, 1, 0)
            'TxtDiaFijo2.Text = Trim(str(pnDiaFijo2))
            
        End If
        
        txtFechaPago.Text = CDate(dFechaPago)
        
        nTipoGracia = pnTipoGracia
        
        ChkPerGra.value = IIf(bGracia, 1, 0)
        
        TxtTasaGracia.Text = Format(pnTasaGracia, "#0.0000")
        bGraciaGenerada = True
        Set MatGracia = Nothing
        cmdAplicar.Enabled = False
        FraDatos.Enabled = False
        FraFechaPago.Enabled = False
        'ChkCuotaCom.value = pnCuotaCom
        MatGracia = pMatGracia
        
        'Cambios para las opciones de gracia
        If pnTipoGracia = EnCuotas - 1 Then
            optTipoGracia(1).value = True
        End If
        If pnTipoGracia = Capitalizada - 1 Then
            optTipoGracia(0).value = True
        End If
        
        If bQuincenal = True Then
            'ChkQuincenal.value = 1
        End If
        
        If pnCuotaBalon > 0 Then
            'chkCuotaBalon.value = 1
            'uspCuotaBalon.valor = pnCuotaBalon
        End If
        
        'txtValorInmueble.Text = ""
        'txtCuotaInicial.Text = ""
        'txtBonoBuenPagador.Text = ""
        If pnMiViv = 1 Then
            If IsArray(pArrMIVIVIENDA) Then
                If Trim(pArrMIVIVIENDA(0)) <> "" Then
                    'txtValorInmueble.Text = Format(CDbl(pArrMIVIVIENDA(0)), "###," & String(15, "#") & "#0.00")
                    'txtCuotaInicial.Text = Format(CDbl(pArrMIVIVIENDA(1)), "###," & String(15, "#") & "#0.00")
                    'txtBonoBuenPagador.Text = Format(CDbl(pArrMIVIVIENDA(2)), "###," & String(15, "#") & "#0.00")
                End If
            End If
        End If
        
        Call cmdAplicar_Click
        
        If pnTipoGracia = 6 Then
            pnIntGraInicial = nIntGraInicial
        End If
        
        cmdResumen.Enabled = False
        'FraMivivienda.Enabled = False
        
        cmbSeguroDes.Visible = False
        
        If bErrorValidacion = True Then
            pbNoMostrarCalendario = True
        End If
        
        If Not pbNoMostrarCalendario Then
            Me.Show 1
        End If
        
        
        Inicio = MatCalend
        MatMiViv_2 = MatResul
        
        pbErrorValidacion = bErrorValidacion
        cCtaCodG = ""
        nSugerAprob = 0
End Function


Private Function ValidaDatos() As Boolean
    ValidaDatos = True
    
    '203 Producto Comsumo Prendario
    If (nTpoSubPro = 203) Then
        'Valida Selecion de Tipo de Cliente
        If cmbTpoCliente.ListIndex = -1 Then
            MsgBox "Debe Selecionar el Tipo de Cliente", vbInformation, "Aviso"
            cmbTpoCliente.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    
    Else
        'Interes ARLO20181029
        If (Me.fraGracia.Enabled = True) Then
            If CDbl(Me.txtInteresAnual) <> CDbl(Me.TxtTasaGraciaNEW) Then
                MsgBox "Falto Presionar Enter en el Campo Fecha de Pago", vbInformation, "Aviso"
                ValidaDatos = False
                If txtInteresAnual.Enabled Then txtInteresAnual.SetFocus
                Exit Function
            End If
        End If
            
        'Interes
        If Trim(txtInteresAnual.Text) = "" Then
                MsgBox "Ingrese la Tasa de Interes", vbInformation, "Aviso"
                ValidaDatos = False
                If txtInteresAnual.Enabled Then txtInteresAnual.SetFocus
                Exit Function
        Else
            If CDbl(Me.txtInteresAnual.Text) = 0 Then
                MsgBox "Ingrese la Tasa de Interes", vbInformation, "Aviso"
                ValidaDatos = False
                If TxtInteres.Enabled Then TxtInteres.SetFocus
                Exit Function
            End If
        End If
        
        If CDbl(Me.txtInteresAnual.Text) = 0 Then
                MsgBox "Falto Presionar Enter en el Campo T.E.A", vbInformation, "Aviso"
                ValidaDatos = False
                If txtInteresAnual.Enabled Then txtInteresAnual.SetFocus
                Exit Function
        End If
        
        
        'Numero de Cuotas
        If Trim(SpnCuotas.valor) = "" Or CInt(SpnCuotas.valor) <= 0 Then
            MsgBox "Ingrese el Numero de Cuotas del Prestamo", vbInformation, "Aviso"
            ValidaDatos = False
            If SpnCuotas.Enabled Then SpnCuotas.SetFocus
            Exit Function
        End If
        
        'Plazo de Cuotas
        If (Trim(SpnPlazo.valor) = "" Or SpnPlazo.valor = "0" Or SpnPlazo.valor = "00") And OptTipoPeriodo(0).value = True Then
            MsgBox "Ingrese el Plazo de las Cuotas del Prestamo", vbInformation, "Aviso"
            ValidaDatos = False
            If SpnPlazo.Enabled Then SpnPlazo.SetFocus
            Exit Function
        End If
        
        'Plazo de Cuotas 'ARLO20180809
        If (OptTipoPeriodo(0).value = True) Then
            If CInt(SpnPlazo.valor) < 30 Then
                MsgBox "El Plazo del crédito tiene que ser mayor o igual a 30 días", vbInformation, "Aviso"
                ValidaDatos = False
                If SpnPlazo.Enabled Then SpnPlazo.SetFocus
                Exit Function
            End If
        Else
            If CDate(CDate(DTFecDesemb.value) + CDate(30)) > CDate(txtFechaPago.Text) Then
                MsgBox "El Plazo del crédito tiene que ser mayor o igual a 30 días", vbInformation, "Aviso"
                ValidaDatos = False
                If SpnPlazo.Enabled Then SpnPlazo.SetFocus
                Exit Function
            End If
        End If
        
        'Fecha de Desembolso
        If ValidaFecha(DTFecDesemb.value) <> "" Then
            MsgBox ValidaFecha(DTFecDesemb.value), vbInformation, "Aviso"
            ValidaDatos = False
            If DTFecDesemb.Enabled Then DTFecDesemb.SetFocus
            Exit Function
        End If
        
        'Valida dia de Fecha Fija
        If OptTipoPeriodo(1).value And (Trim(TxtDiaFijo.Text) = "" Or Trim(TxtDiaFijo.Text) = "0" Or Trim(TxtDiaFijo.Text) = "00") Then
            MsgBox "Ingrese el día del mes que vencerán todas las cuotas", vbInformation, "Aviso"
            ValidaDatos = False
            If TxtDiaFijo.Enabled Then TxtDiaFijo.SetFocus
            Exit Function
        End If
        'Valida Generacion de Tipos de Periodo de Gracia
        If ChkPerGra.value = 1 Then
            If (TxtPerGra.Text = "00" Or TxtPerGra.Text = "0") Then
                MsgBox "Ingrese los Días de Gracia", vbInformation, "Aviso"
                ValidaDatos = False
                If TxtPerGra.Enabled Then TxtPerGra.SetFocus
                Exit Function
            Else
                If (TxtTasaGracia.Text = "0.00" Or TxtTasaGracia.Text = "") Then
                    MsgBox "Ingrese la Tasa de Gracia ", vbInformation, "Aviso"
                    ValidaDatos = False
                    If TxtTasaGracia.Enabled Then TxtTasaGracia.SetFocus
                        Exit Function
                End If
            End If
        End If
            
        If CInt(TxtPerGra.Text) > 0 Then
            Dim dFechaGracia As Date
            Dim nDiasGraciaPermitido As Integer
            If OptTipoPeriodo(1).value Then 'Fecha Fija
            End If
        End If
    
        If OptTipoPeriodo(0).value = True Then
            If CDate(CDate(DTFecDesemb.value) + CDate(SpnPlazo.valor) + CDate(TxtPerGra.Text)) <> CDate(txtFechaPago.Text) Then
                MsgBox "Falto Presionar Enter en el Campo Fecha de Desembolso", vbInformation, "Aviso"
                DTFecDesemb.SetFocus
                ValidaDatos = False
                Exit Function
            End If
        Else
            If Len(Trim(lsCtaCodLeasing)) = 0 Then
                If CDate(TxtPerGra.Text) <> "0" Then
                    If CDate(CDate(DTFecDesemb.value) + CDate(30) + CDate(TxtPerGra.Text)) <> CDate(txtFechaPago.Text) Then
                        MsgBox "Falto Presionar Enter en el Campo Fecha de Desembolso", vbInformation, "Aviso"
                        DTFecDesemb.SetFocus
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
            End If
        End If
        If fbInicioSim Then
            If fraGastoCom.Visible Then
                If (Me.chkSegDesgra.value = 1) Then
                    If Trim(cmbSeguroDes.Text) = "" Then
                        MsgBox "Ingrese el tipo de Seguro Desgravamen", vbInformation, "Aviso"
                        cmbSeguroDes.SetFocus
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
                If chkEnvioEst.value = 1 Then
                    If Trim(cmbEnvioEst.Text) = "" Then
                        MsgBox "Ingrese el Tipo de Envío del estado de cuenta", vbInformation, "Aviso"
                        cmbEnvioEst.SetFocus
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
        'Valida Selecion de Tipo de Moneda
        If cmbMoneda.ListIndex = -1 Then
            MsgBox "Debe Selecionar el Tipo de Moneda", vbInformation, "Aviso"
            cmbMoneda.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        'Monto de Prestamo
        If Trim(TxtMonto.Text) = "" Then
            MsgBox "Ingrese el Monto del Prestamo", vbInformation, "Aviso"
            ValidaDatos = False
            If TxtMonto.Enabled Then TxtMonto.SetFocus
            Exit Function
        ElseIf CDbl(TxtMonto.Text) <= 0 Then
            MsgBox "El Monto del Prestamo debe ser Mayor a Cero[0]", vbInformation, "Aviso"
            ValidaDatos = False
            If TxtMonto.Enabled Then TxtMonto.SetFocus
            Exit Function
        End If
        If nTpoSubPro <> 203 Then
        If (Me.chkHipoteca.value = 1) Then
            If (Me.txtEdificacion.Text = "" Or Me.txtEdificacion.Text = "0.00") Then
            MsgBox "Ingrese el Monto de la Edificación", vbInformation, "Aviso"
            ValidaDatos = False
            If txtEdificacion.Enabled Then txtEdificacion.SetFocus
            Exit Function
            End If
        End If
        End If

End Function

Private Sub HabilitaFechaFija(ByVal pbHabilita As Boolean)
    SpnPlazo.Enabled = Not pbHabilita
    SpnPlazo.valor = IIf(pbHabilita, "0", SpnPlazo.valor)
    LblDia.Enabled = pbHabilita
    TxtDiaFijo.Enabled = pbHabilita
    Me.cmbProductoCMACM.Enabled = IIf(pbHabilita, True, True)
    Me.fraGracia.Enabled = IIf(pbHabilita, False, False)
    Me.cmbSubProducto.Enabled = Not pbHabilita
    Me.cmbMoneda.Enabled = Not pbHabilita
    Me.cmbTpoCliente.Enabled = Not pbHabilita
    Me.cmbHipoteca.Enabled = Not pbHabilita
    Me.cmdImprimir.Enabled = pbHabilita
    Me.txtTEM.Enabled = IIf(pbHabilita, False, False) 'ARLO20181029
    If (Me.cmbProductoCMACM.ListIndex = -1) Then
        Me.cmbSubProducto.Enabled = pbHabilita
    End If
    If pbHabilita Then
        TxtDiaFijo.Text = "00"
        If TxtDiaFijo.Enabled And TxtDiaFijo.Visible And FraDatos.Enabled Then
            TxtDiaFijo.SetFocus
        End If
    End If
End Sub

Private Sub CargaControles()
    FECalend.RowHeight(0) = 250
    FECalend.RowHeight(1) = 250
    TxtDiaFijo.Text = "00"
    TxtTasaGracia.Text = "00"
End Sub
Private Sub chkEnvioEst_Click()
    If chkEnvioEst.value = 1 Then
        cmbEnvioEst.Enabled = True
    Else
        cmbEnvioEst.Enabled = False
    End If
End Sub
Private Sub chkHipoteca_Click()
    If chkHipoteca.value = 1 Then
        cmbHipoteca.Enabled = True
        txtEdificacion.Enabled = True
    Else
        cmbHipoteca.Enabled = False
        txtEdificacion.Enabled = False
        cmbHipoteca.ListIndex = -1
    End If
End Sub

Private Sub ChkPerGra_Click()
Dim i As Integer

    ReDim MatGracia(CInt(SpnCuotas.valor))

    For i = 0 To CInt(SpnCuotas.valor) - 1
        MatGracia(i) = "0.00"
    Next i
    Call LimpiaFlex(FECalend)
    If ChkPerGra.value = 1 Then
        LblTasaGracia.Enabled = True
        TxtTasaGracia.Enabled = True
        LblPorcGracia.Enabled = True

        If txtInteresAnual.Text <> "" Then
        TxtTasaGraciaNEW.Text = Format(txtInteresAnual.Text, "#0.00") 'ARLO20181029
        TxtTasaGraciaNEW.Enabled = False
        End If
        TxtTasaGracia.Text = nTem
        optTipoGracia(0).value = False
        
        'Para Fecha Fija no Aplica
        If OptTipoPeriodo(1).value = True Then
            optTipoGracia(0).Enabled = True
        Else
            optTipoGracia(0).Enabled = True
        End If
    Else
        LblTasaGracia.Enabled = False
        TxtTasaGracia.Enabled = False
        LblPorcGracia.Enabled = False
        TxtPerGra.Enabled = False
        TxtPerGra.Text = "0"
        TxtTasaGracia.Text = "0.00"
        optTipoGracia(0).Enabled = False
        optTipoGracia(0).value = False
        
        GenerarFechaPago
        If OptTipoPeriodo(1).value = True Then
            ChkPerGra.Enabled = False
        End If
        Call txtFechaPago_KeyPress(13)
    End If
End Sub

Private Sub chkSegDesgra_Click()
        
        If (chkSegDesgra.value = 1) Then
            Me.cmbSeguroDes.Enabled = True
        Else
            Me.cmbSeguroDes.Enabled = False
        End If
End Sub

Private Sub cmbHipoteca_Click()
    Dim nEdificacion As Double
    If (cmbHipoteca.Text <> "") Then
        nTpoHipo = CInt(Trim(Right(cmbHipoteca.Text, 3)))
        cTpoHipo = Trim(Left(cmbHipoteca.Text, 30))
        txtEdificacion.SetFocus
        nEdificacion = val(txtEdificacion.Text)
        txtEdificacion.Text = Format(nEdificacion, "#0.00")
    End If
End Sub
Private Sub cmbMoneda_Click()
 If (cmbMoneda.Text <> "") Then
    cTpoMoneda = Trim(Left(cmbMoneda.Text, 15))
    If txtInteresAnual.Enabled Then
    txtInteresAnual.SetFocus
    End If
    If txtInteresAnual.Text = "" Then
    txtInteresAnual.Text = "0.00"
    Else
    txtInteresAnual.Text = Format(txtInteresAnual.Text, "#0.00")
    End If
 End If
End Sub

Private Sub cmbTpoCliente_Click()
 If (cmbTpoCliente.Text <> "") Then
 nTpoCliente = CInt(Trim(Right(cmbTpoCliente.Text, 3)))
 cTpoCliente = Trim(Left(cmbTpoCliente.Text, 30))
 End If
End Sub
Private Sub cmdAplicar_Click()
Dim i As Integer
Dim nTipoCuota As Integer
Dim nTipoPeriodo As Integer
Dim nTotalInteres As Double
Dim nTotalCapital As Double
Dim oParam As COMDCredito.DCOMParametro
Dim nTramoNoConsMonto As Double
Dim nTramoConsMonto As Double
Dim nTramoNoConsPorcen As Double
Dim nPlazoMiViv As Integer
Dim nPlazoMiVivMax As Integer
Dim nRedondeoITF As Double
Dim nTotalcuotasCONItF As Double
Dim nTotalcuotasLeasing As Double
Dim lnSalCapital As Double
Dim oCredito As COMNCredito.NCOMCredito
Dim nPlazoMinHipo As Integer
Dim nPlazoMaxHipo As Integer
Dim nValorMaxParam As Long
Dim nValorMinParam As Long
Dim dFechaFinMes As Date

nTotalcuotasLeasing = 0

    nIntGraInicial = 0
    nMontoCapInicial = 0
    
    Call LimpiaFlex(FECalend)
    Call LimpiaFlex(FECalBPag)
    Call LimpiaFlex(FECalMPag)
    MatResul = Array(0)
    MatResulDiff = Array(0)
    MatCalend = Array(0)
       
    If Not ValidaDatos Then
        bErrorValidacion = True
        Exit Sub
    Else
        bErrorValidacion = False
    End If
    
    Call txtInteresAnual_KeyPress(13)

    
    'If nTpoPro = 3 Then
        If (nTpoSubPro = 301) Then
            nValorMinParam = 1027396
            nValorMaxParam = 1027397
        ElseIf (nTpoSubPro = 302) Then
            nValorMinParam = 1027398
            nValorMaxParam = 1027399
        ElseIf (nTpoSubPro = 303) Then
            nValorMinParam = 3065
            nValorMaxParam = 3066
        End If

        If nTpoSubPro = 201 Then 'PLAZO FIJO
                nValorMinParam = 1027401
                nValorMaxParam = 1027400
        End If
        
'ARLO20190229 ERS042-2018 - COMENTADO
        'Porcentaje Real sin dividir enter 100
'        If (nTpoPro = 1 Or nTpoPro = 2) And nTpoSubPro <> 201 Then
'            nPlazoMinHipo = 0
'            nPlazoMaxHipo = 5
'        Else
'            Set oParam = New COMDCredito.DCOMParametro
'            nPlazoMinHipo = oParam.RecuperaValorParametro(nValorMinParam)
'            nPlazoMaxHipo = oParam.RecuperaValorParametro(nValorMaxParam)
'            Set oParam = Nothing
'        End If
'
'        If (CInt(spnCuotas.valor) * 30) / 360 < nPlazoMinHipo Then
'            MsgBox "El Plazo del Credito debe ser Minimo " & nPlazoMinHipo & " Años", vbInformation, "Aviso"
'            bErrorValidacion = True
'            Exit Sub
'        End If
'
'        If (CInt(spnCuotas.valor) * 30) / 360 > nPlazoMaxHipo Then
'            MsgBox "El Plazo del Credito debe ser Maximo " & nPlazoMaxHipo & " Años", vbInformation, "Aviso"
'            bErrorValidacion = True
'            Exit Sub
'        End If
'ARLO20190229 ERS042-2018

        If (nTpoPro = 8) Then 'ARLO20190229 ERS042-2018 CAMBIO DE 3 A 8
            If (Me.chkHipoteca.value = 1) Then
                If (cmbHipoteca.Text = "") Then
                        MsgBox "Debe Selecionar un tipo de incendio para la Hipoteca", vbInformation, "Aviso"
                        bErrorValidacion = True
                        Exit Sub
                End If
                If (txtEdificacion.Text = "" Or txtEdificacion.Text = "0.00") Then
                        MsgBox "Debe Selecionar un tipo de incendio para la Hipoetca", vbInformation, "Aviso"
                        bErrorValidacion = True
                        Exit Sub
                End If
            Else
                MsgBox "Ud. Seleciono el producto hipotecario, se debera activar el check tiene Hipoteca ", vbInformation, "Aviso"
                Me.chkHipoteca.value = 1
                Me.chkHipoteca.SetFocus
                bErrorValidacion = True
                Exit Sub
            End If
        End If

    Call LimpiaFlex(FECalend)
    Call LimpiaFlex(FECalBPag)
    Call LimpiaFlex(FECalMPag)
    For i = 0 To 2
        If OptTipoCuota(i).value Then
            nTipoCuota = i + 1
            Exit For
        End If
    Next i
    For i = 0 To 1
        If OptTipoPeriodo(i).value Then
            nTipoPeriodo = i + 1
            Exit For
        End If
    Next i
    
    If (optTipoGracia(0).value = False And ChkPerGra.value = 1) Then
        Call GeneraGracia
    End If
     
    nTasaNormal = 151.82
    nTasaCPP = 213.84
    
    Dim oPol As COMDCredito.DCOMPoliza
    Dim rs As ADODB.Recordset
    
    Set oPol = New COMDCredito.DCOMPoliza
    
    Set rs = oPol.RecuperaTasasPolizas()
    Set oPol = Nothing
    
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            If (gsCodAge = rs!cAgeCod) Then
                If (rs!Inmueble = "CASA HABITACION") Then
                    nCasaHabiHipo = rs(2)
                ElseIf (rs!Inmueble = "LOCAL COMERCIAL") Then
                    nLocalHipo = rs(2)
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    
    If (Me.ChkPerGra.value = 1) Then
         dFechaFinMes = fin_del_Mes(Me.txtFechaPago.Text)
         nMultiplicaSegMutl = Redondeo((CDate(dFechaFinMes) - CDate(DTFecDesemb.value)) / 30, 0)
    Else
         nMultiplicaSegMutl = 1
    End If
    
    If (nTpoHipo = 1) Then
        nSegMult = (nLocalHipo * 1.2154 / 1000) / 12
    Else
        nSegMult = (nCasaHabiHipo * 1.2154 / 1000) / 12
    End If
       
    If nTpoSubPro <> 203 Then
        If (Me.chkHipoteca.value = 1) Then
            nSegMutlTotal = (nSegMult * CDbl(txtEdificacion.Text) * val(Me.SpnCuotas.valor))
            nSegMutlUnidad = Round(nSegMult * CDbl(txtEdificacion.Text), 2)
            If (nSegMutlUnidad < 22.79 And cTpoMoneda = "SOLES") Then
                nSegMutlUnidad = 22.79
            ElseIf (nSegMutlUnidad < 7.6 And cTpoMoneda = "DOLARES") Then
                nSegMutlUnidad = 7.6
            End If
        Else
            nSegMutlUnidad = 0
        End If
    End If
    If (nTpoSubPro = 203) Then
            
            If (nTpoCliente = 1) Then
                nTEA = nTasaNormal
            Else
                nTEA = nTasaCPP
            End If
            
            nInteresPigno = Round(((nTEA / 100 + 1) ^ (1 / 12) - 1) * CDbl(Me.TxtMonto.Text), 2)
            txtTotalPigno.Text = (nInteresPigno) + CDbl(Me.TxtMonto.Text)
            txtTotalPigno.Text = Format(txtTotalPigno.Text, "#0.00")
            txtInteresPigno.Text = Format(nInteresPigno, "#0.00")
            txtFechaVencPigno.Text = CDate(Me.DTFecDesemb.value + 30)
            
    Else
        'Se Agrego para manejar la Capitalizacion de la Gracia
        If optTipoGracia(0).value Then
            Set oCredito = New COMNCredito.NCOMCredito
            'Para realizar los cálculos
            nMontoCapInicial = CDbl(TxtMonto.Text)
            nIntGraInicial = oCredito.MontoIntPerDias(CDbl(TxtTasaGracia.Text), CInt(TxtPerGra.Text), CDbl(TxtMonto.Text))
            nTipoGracia = gColocTiposGraciaCapitalizada
            Set oCredito = Nothing
        End If
        '*********************************************lsCtaCodLeasing
        If Len(Trim(lsCtaCodLeasing)) = 0 Then
                
            'MatCalend = GeneraCalendario(CDbl(TxtMonto.Text), CDbl(TxtInteres.Text), CInt(SpnCuotas.valor), _
            '            CInt(SpnPlazo.valor), CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), nTipoCuota, nTipoPeriodo, _
            '            nTipoGracia, CInt(TxtPerGra.Text), CInt(TxtDiaFijo.Text), bProxMesN, MatGracia, True, False, bDesemParcial, MatDesPar, , , , _
            '            False, CDbl(TxtTasaGracia.Text), 0, nMontoCapInicial, False, bRenovarCredito, nInteresAFecha, nIntGraInicial, _
            '            0)

            Dim oNGasto As New COMNCredito.NCOMGasto
            Dim lnTasaSegDes As Double
            Dim RGastosSegDes As ADODB.Recordset
            Dim MatCalendSegDes As Variant

            Dim oGastosCab As New COMDCredito.DCOMGasto
            Set RGastosSegDes = oGastosCab.RecuperaGastosCabecera(1)
            RGastosSegDes.Filter = " nPrdConceptoCod = 1217"
            lnTasaSegDes = 0
            
            If Me.cmbSeguroDes.Enabled Or Me.cmbEnvioEst.Enabled Then
                lnTasaSegDes = IIf(Trim(Right(cmbSeguroDes.Text, 4)) = "1", RGastosSegDes!nValor, RGastosSegDes!nValorDosTit)
            End If

            MatCalend = GeneraCalendarioNuevo(CDbl(TxtMonto.Text), _
                        CDbl(TxtInteres.Text), _
                        CInt(SpnCuotas.valor), _
                        CInt(SpnPlazo.valor), _
                        CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), _
                        nTipoCuota, _
                        nTipoPeriodo, _
                        nTipoGracia, _
                        CInt(TxtPerGra.Text), _
                        CInt(TxtDiaFijo.Text), _
                        bProxMesN, _
                        MatGracia, _
                        True, False, bDesemParcial, _
                        MatDesPar, , , , _
                        False, _
                        CDbl(TxtTasaGracia.Text), _
                        0, nMontoCapInicial, _
                        False, bRenovarCredito, _
                        nInteresAFecha, nIntGraInicial, _
                        0, _
                        cCtaCodG, lnTasaSegDes, MatCalendSegDes, , _
                        nSegMutlUnidad, nSegMult)

        Else
        MatCalend = GeneraCalendarioLeasing(CDbl(TxtMonto.Text), CDbl(TxtInteres.Text), CInt(SpnCuotas.valor), _
                    CInt(SpnPlazo.valor), CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), nTipoCuota, nTipoPeriodo, _
                    nTipoGracia, CInt(TxtPerGra.Text), CInt(TxtDiaFijo.Text), 0, MatGracia, True, False, bDesemParcial, MatDesPar, , , , _
                    optTipoGracia(1).value, CDbl(TxtTasaGracia.Text), "", nMontoCapInicial, False, bRenovarCredito, nInteresAFecha, lsCtaCodLeasing)
        End If

        If bRenovarCredito Then
            Call ObtenerGastosEnReprogramacion
        End If
        
        If Me.cmbSeguroDes.Enabled Or Me.cmbEnvioEst.Enabled Then
            Call ObtenerDesgravamen
        Else
            For i = 0 To UBound(MatCalend) - 1
                If (i = 0) And nTipoGracia <> gColocTiposGraciaCapitalizada Then
                    'MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 8)) + (nSegMutlUnidad * nMultiplicaSegMutl))
                    'Monto de Cuota = Capital + Interes Compensatorio + Interes Gracia + Gasto +(Seguro Poliza Incendio + Gracia Poliza Incendio) + Seguro Desgramen
                    MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 16)) + CDbl(MatCalend(i, 8)), "#0.00")
                    
                ElseIf (i = 1) And nTipoGracia = gColocTiposGraciaCapitalizada Then
                '    MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 8)) + (nSegMutlUnidad * nMultiplicaSegMutl))
                    MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 16)) + CDbl(MatCalend(i, 8)), "#0.00")
                Else
                '    MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 8)) + nSegMutlUnidad)
                    MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 16)) + CDbl(MatCalend(i, 8)), "#0.00")
                End If
            Next i
        End If

        nTotalInteres = 0
        nTotalCapital = 0
        nTotalcuotasCONItF = 0
        lnSalCapital = val(TxtMonto.Text)

        For i = 0 To UBound(MatCalend) - 1
            FECalend.AdicionaFila
            FECalend.TextMatrix(i + 1, 1) = Trim(MatCalend(i, 0))
            If Len(Trim(lsCtaCodLeasing)) = 18 Then
                txtFechaPago.Text = CDate(Trim(MatCalend(0, 0)))
            End If
            If i = 0 Then
                If CDate(Trim(MatCalend(i, 0))) <> CDate(txtFechaPago.Text) And nTipoGracia <> 1 Then
                    If nTipoGracia <> 6 Then
                        Call LimpiaFlex(FECalend)
                        MsgBox "Falto Presionar Enter en el Campo Fecha de Pago", vbInformation, "Aviso"
                        bErrorValidacion = True
                        Exit Sub
                    End If
                End If
            End If

            FECalend.TextMatrix(i + 1, 2) = Trim(MatCalend(i, 1))
            
            If nTipoGracia = 6 Then
                MatCalend(i, 2) = Trim(CDbl(MatCalend(i, 2)) + CDbl(MatCalend(i, 11)))
                FECalend.TextMatrix(i + 1, 4) = Format(Trim(MatCalend(i, 2)), "#0.00") '3
            Else
                FECalend.TextMatrix(i + 1, 4) = Trim(MatCalend(i, 2)) '3
            End If

            FECalend.row = i + 1
            FECalend.col = 3
            FECalend.CellForeColor = vbBlue
            FECalend.TextMatrix(i + 1, 5) = Trim(MatCalend(i, 3)) 'Amort Cap '4

            If nTipoGracia = 6 Then
                MatCalend(i, 4) = Trim(CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 11)))
                FECalend.TextMatrix(i + 1, 6) = Format(Trim(MatCalend(i, 4)), "#0.00") '5
            Else
                FECalend.TextMatrix(i + 1, 6) = Trim(MatCalend(i, 4)) '5
            End If
            
            'Interes Gracia
            FECalend.TextMatrix(i + 1, 7) = Trim(MatCalend(i, 5)) '6

            If Len(Trim(lsCtaCodLeasing)) = 0 Then
                FECalend.TextMatrix(i + 1, 8) = Trim(MatCalend(i, 6)) '7
                FECalend.TextMatrix(i + 1, 9) = Trim(MatCalend(i, 8)) '8
            Else
                FECalend.TextMatrix(i + 1, 8) = Trim(MatCalend(i, 6)) '7
                FECalend.TextMatrix(i + 1, 9) = Trim(MatCalend(i, 8)) '8
            End If

            FECalend.TextMatrix(i + 1, 11) = Trim(MatCalend(i, 7)) '10
            
            If (chkHipoteca.value = 1) Then
                If (i = 0 And nTipoGracia <> gColocTiposGraciaCapitalizada) Then
                    FECalend.TextMatrix(i + 1, 10) = CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 16)) 'nSegMutlUnidad * nMultiplicaSegMutl
                    MatCalend(i, 13) = CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 16)) 'nSegMutlUnidad * nMultiplicaSegMutl
                ElseIf (i = 1 And nTipoGracia = gColocTiposGraciaCapitalizada) Then
                    FECalend.TextMatrix(i + 1, 10) = CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 16)) 'nSegMutlUnidad * nMultiplicaSegMutl
                    MatCalend(i, 13) = CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 16)) 'nSegMutlUnidad * nMultiplicaSegMutl
                Else
                    FECalend.TextMatrix(i + 1, 10) = CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 16)) 'nSegMutlUnidad
                    MatCalend(i, 13) = CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 16)) 'nSegMutlUnidad '9
                End If
            Else
                FECalend.TextMatrix(i + 1, 10) = Trim(MatCalend(i, 13)) '9
            End If
            
            If Not (i = 0 And nTipoGracia = gColocTiposGraciaCapitalizada) Then
                nTotalCapital = nTotalCapital + CDbl(Trim(MatCalend(i, 3)))
                nTotalInteres = nTotalInteres + CDbl(Trim(MatCalend(i, 4))) + CDbl(Trim(MatCalend(i, 5)))
            End If

            FECalend.TextMatrix(i + 1, 3) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))), "0.00") '11
            nRedondeoITF = fgDiferenciaRedondeoITF(fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))))
            If nRedondeoITF > 0 Then
                FECalend.TextMatrix(i + 1, 3) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))) - nRedondeoITF, "0.00") '11
            Else
                FECalend.TextMatrix(i + 1, 3) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))), "0.00") '11
            End If

            If Not (nTipoGracia = gColocTiposGraciaCapitalizada And i = 0) Then
                nTotalcuotasCONItF = nTotalcuotasCONItF + CDbl(FECalend.TextMatrix(i + 1, 3))
                nTotalcuotasLeasing = nTotalcuotasLeasing + CDbl(FECalend.TextMatrix(i + 1, 3))
            End If
        Next i

        Set oCredito = Nothing
        If nTipoGracia = gColocTiposGraciaCapitalizada Then
            FECalend.TextMatrix(1, 3) = ""
            FECalend.TextMatrix(1, 4) = ""
            FECalend.TextMatrix(1, 6) = ""
            FECalend.TextMatrix(1, 8) = ""
            FECalend.TextMatrix(1, 9) = ""
            FECalend.TextMatrix(1, 10) = ""
        End If

        lblCapital.Caption = Format(nTotalCapital, "#0.00")
        lblInteres.Caption = Format(nTotalInteres, "#0.00")
        lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#0.00")
        lblTotalCONITF.Caption = Format(nTotalcuotasCONItF, "#0.00")
        FECalend.row = 1
        FECalend.TopRow = 1
    End If

    If (nTpoSubPro = 203) Then
        FraPigno.Visible = True
        FraPigno.Enabled = False
        fraTasaAnuales.Visible = True
        lblTasaEfectivaAnual = nTEA
        lblTasaCostoEfectivoAnual = nTEA
        nTasaEfectivaAnual = nTEA
    Else
        fraTasaAnuales.Visible = True
        Set oCredito = New COMNCredito.NCOMCredito
            nTasaEfectivaAnual = Round(oCredito.TasaIntPerDias(CDbl(TxtInteres), 360) * 100, 2)
            If nTipoGracia = 6 Then
                Dim Y As Integer
                Dim MatCalendTemp() As String
                ReDim MatCalendTemp(UBound(MatCalend) - 1, 14)
                For i = 0 To UBound(MatCalend) - 1
                    For Y = 0 To 14
                        MatCalendTemp(i, Y) = MatCalend(i + 1, Y)
                    Next Y
                Next i
                Erase MatCalend
                ReDim MatCalend(UBound(MatCalendTemp), 14)
                
                For i = 0 To UBound(MatCalendTemp)
                    For Y = 0 To 14
                        MatCalend(i, Y) = MatCalendTemp(i, Y)
                    Next Y
                Next i
                Erase MatCalendTemp
            End If
            nTasaCostoEfectivoAnual = oCredito.GeneraTasaCostoEfectivoAnual(CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), CDbl(TxtMonto.Text), MatCalend, CDbl(TxtInteres), lsCtaCodLeasing, nTipoPeriodo)
            'lblTasaEfectivaAnual.Caption = nTasaEfectivaAnual & " %"
            lblTasaEfectivaAnual.Caption = Format(Round(CDbl(txtInteresAnual.Text), 3), "#0.00") & " %" 'ARLO20180809
            If (nTasaCostoEfectivoAnual < Round(CDbl(txtInteresAnual.Text), 3)) Then    'ARLO20180809
                lblTasaCostoEfectivoAnual.Caption = Format(Round(CDbl(txtInteresAnual.Text), 3), "#0.00") & " %"
            Else
                lblTasaCostoEfectivoAnual.Caption = Format(nTasaCostoEfectivoAnual, "#0.00") & " %"
            End If
        Set oCredito = Nothing
    End If
    If UBound(MatCalend) = 0 And nTpoSubPro <> 203 Then
        cmdImprimir.Enabled = False
    Else
        cmdImprimir.Enabled = True
        If (nTpoPro = 3) Then
            cmdResumen.Visible = True
            cmdResumen.Enabled = True
        End If
    End If
    If (Me.cmbProductoCMACM.ListIndex = -1) Then
        Me.cmbSubProducto.Enabled = False
    Else
        Me.cmbSubProducto.Enabled = True
    End If
    
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
    bExisteModal = False 'ARLO20180809
End Sub
Private Sub GeneraGracia()
Dim oCredito As COMNCredito.NCOMCredito

Set oCredito = New COMNCredito.NCOMCredito

If CDbl(Me.TxtTasaGraciaNEW.Text) <= 0# Then
    MsgBox "Ingrese la Tasa de Interes de Gracia", vbInformation, "Aviso"
    TxtTasaGraciaNEW.SetFocus
    Exit Sub
End If

    MatGracia = frmCredGracia.Inicio(CInt(TxtPerGra.Text), oCredito.MontoIntPerDias(CDbl(TxtTasaGracia.Text), CInt(TxtPerGra.Text), CDbl(TxtMonto.Text)), CInt(SpnCuotas.valor), nTipoGracia, psCtaCod, , bSimulador)
    
    Set oCredito = Nothing
    bGraciaGenerada = True
    Call LimpiaFlex(FECalend)
End Sub

Private Sub EjecutaReporte()
Dim loRep As COMNCredito.NCOMCalendario
Dim lsDestino As String  ' P= Previo // I = Impresora // A = Archivo // E = Excel
Dim lsCadImp As String   'cadena q forma
Dim x As Integer
Dim loPrevio As previo.clsprevio
Dim Agencia As String
Dim lnAge As Integer

Dim TCuota As String
Dim Periodo As String
Dim tIndex As Integer
Dim Interes As Double
Dim Monto As Double
Dim Cuota As Double
Dim Plazo As Integer
Dim Vigencia As Date
Dim pnCuotas As Integer
Dim ArrSimulador() As Variant
Dim pnTpoReporte As Integer

If OptTipoCuota(0).value Then
    tIndex = 0
End If
Select Case tIndex
    Case 0: TCuota = "Cuota Fija"
    Case 1: TCuota = "Cuota Creciente"
    Case 2: TCuota = "Cuota Decreciente"
End Select

ReDim ArrSimulador(3)
ArrSimulador(0) = 0 'Para decir que es del simulador
ArrSimulador(1) = 0 'Para decir el envio de estado de cuenta
ArrSimulador(2) = 0 'Seguro Desgravamen

    If (nTpoPro = 3) Then
        pnTpoReporte = 2
    ElseIf (nTpoSubPro = 203) Then
        pnTpoReporte = 3
        nTotalPigno = CDbl(val(txtTotalPigno))
    Else
        pnTpoReporte = 1
    End If
If (txtFechaVencPigno.Text = "__/__/____") Then
    dFechaPagoPigno = Format(1 / 5 / 2018, "dd/mm/yyyy")
    Else
    dFechaPagoPigno = Format(txtFechaVencPigno.Text, "dd/mm/yyyy")
End If

If (txtTotalPigno.Text = "") Then
    TxtInteres.Text = 0
End If
If (txtInteresPigno.Text = "") Then
    TxtInteres.Text = 0
End If

If (TxtInteres.Text = "") Then
    TxtInteres.Text = 0
End If

If (OptTipoPeriodo(1).value) Then
    SpnPlazo.valor = 30
End If

If fbInicioSim Then
    'If fraGastoCom.Visible Then
    If Me.chkSegDesgra.value = 1 Then
        ArrSimulador(0) = 1
        ArrSimulador(2) = CInt(Trim(Right(Me.cmbSeguroDes.Text, 4)))
        If chkEnvioEst = 1 Then
            If Trim(Right(Me.cmbEnvioEst.Text, 4)) = "2" Then
                ArrSimulador(1) = 1
            End If
        End If
    Else
        ArrSimulador(0) = 0
    End If
End If

    Periodo = IIf(OptTipoPeriodo(0).value = True, "Periodo Fijo - ", "Fecha Fija - ")
    TCuota = Periodo & TCuota
    pnCuotas = SpnCuotas.valor
    Set loRep = New COMNCredito.NCOMCalendario
    loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
    lsCadImp = lsCadImp & Chr(10) & loRep.ReporteCalendarioNEW(pnTpoReporte, MatCalend, MatResul, _
    TCuota, CDbl(TxtInteres.Text), TxtMonto.Text, SpnCuotas.valor, SpnPlazo.valor, DTFecDesemb.value, nSugerAprob, IIf(bDesemParcial, MatDesPar, ""), gbITFAplica, gnITFPorcent, gnITFMontoMin, cCtaCodG, pnCuotas, _
    nTasaEfectivaAnual, nTasaCostoEfectivoAnual, lsCtaCodLeasing, nTipoGracia, nIntGraInicial, CInt(TxtPerGra.Text), , ArrSimulador, _
    cTpoSubPro, cTpoPro, cTpoMoneda, dFechaPagoPigno, nTotalPigno, nInteresPigno, nTasaNormal, nTasaCPP, cTpoCliente)

lsDestino = "P"
Set loRep = Nothing
If lsDestino = "P" Then
    If Len(Trim(lsCadImp)) > 0 Then
        Set loPrevio = New previo.clsprevio
            loPrevio.Show lsCadImp, "Calendario de Pagos - Simulacion", True, , gImpresora
        Set loPrevio = Nothing
    Else
        MsgBox "No Existen Datos para el reporte ", vbInformation, "Aviso"
    End If
ElseIf lsDestino = "A" Then
End If

End Sub
Private Sub cmdImprimir_Click()
        If Len(Trim(FECalend.TextMatrix(1, 1))) = 0 And nTpoSubPro <> 203 Then
            MsgBox "No existen datos para imprimir", vbExclamation, "Aviso"
            Exit Sub
        Else
            EjecutaReporte
        End If
End Sub

Private Sub cmdResumen_Click()
Dim oGastos As COMDCredito.DCOMGasto
Dim nGastoAdministracion As Double
Dim loRep As COMNCredito.NCOMCalendario
Dim lsDestino As String  ' P= Previo // I = Impresora // A = Archivo // E = Excel
Dim lsCadImp As String   'cadena q forma
Dim x As Integer
Dim loPrevio As previo.clsprevio
Dim Agencia As String
Dim lnAge As Integer

Dim TCuota As String
Dim Periodo As String
Dim tIndex As Integer
Dim Interes As Double
Dim Monto As Double
Dim Cuota As Double
Dim Plazo As Integer
Dim Vigencia As Date
Dim pnCuotas As Integer
Dim pnTpoReporte As Integer
 
If OptTipoCuota(0).value Then
    tIndex = 0
'Else
'    If OptTipoCuota(1).value Then
'        tIndex = 1
'    Else
'        tIndex = 2
'    End If
End If
Select Case tIndex
    Case 0: TCuota = "Cuota Fija"
    Case 1: TCuota = "Cuota Creciente"
    Case 2: TCuota = "Cuota Decreciente"
End Select
 
    Periodo = IIf(OptTipoPeriodo(0).value = True, "Periodo Fijo - ", "Fecha Fija - ")
    TCuota = Periodo & TCuota
    pnCuotas = SpnCuotas.valor
    Set loRep = New COMNCredito.NCOMCalendario
    
    Set oGastos = New COMDCredito.DCOMGasto
    nGastoAdministracion = oGastos.GetGastoAdmMiViv(1224)
    
    If (nTpoPro = 3) Then
    pnTpoReporte = 2
    Else
    pnTpoReporte = 1
    End If
    
    If (Me.txtEdificacion = "") Then
    Me.txtEdificacion = 0
    End If
    
    
    loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
    lsCadImp = lsCadImp & Chr(10) & loRep.ReporteResumenMiViviendaNEW(pnTpoReporte, MatCalend, MatResul, _
    TCuota, CDbl(TxtInteres.Text), TxtMonto.Text, SpnCuotas.valor, SpnPlazo.valor, DTFecDesemb.value, nSugerAprob, IIf(bDesemParcial, MatDesPar, ""), gbITFAplica, gnITFPorcent, gnITFMontoMin, cCtaCodG, pnCuotas, _
    nTasaEfectivaAnual, nTasaCostoEfectivoAnual, lsCtaCodLeasing, Me.txtEdificacion, 0, 0, TxtPerGra.Text, Me.cmbSeguroDes.Text, nGastoAdministracion, nCuotMens, nCuotMensBono, _
    cTpoSubPro, cTpoPro, cTpoMoneda, cTpoHipo)

lsDestino = "P"
Set loRep = Nothing
If lsDestino = "P" Then
    If Len(Trim(lsCadImp)) > 0 Then
        Set loPrevio = New previo.clsprevio
            loPrevio.Show oImpresora.gPrnTpoLetraSansSerif1PDef & oImpresora.gPrnTamLetra10CPIDef & lsCadImp, "Calendario de Pagos - Simulacion"
        Set loPrevio = Nothing
    Else
        MsgBox "No Existen Datos para el reporte ", vbInformation, "Aviso"
    End If
ElseIf lsDestino = "A" Then
End If
End Sub
Private Sub DTFecDesemb_Change()
    bGraciaGenerada = False
    Call LimpiaFlex(FECalend)
    GenerarFechaPago
End Sub

Private Sub Form_Load()
    CentraForm Me
    nTasaEfectivaAnual = 0: nTasaCostoEfectivoAnual = 0
    Call CargaControles
    Set MatCalendDesParcial = Nothing
    bGraciaGenerada = False
    DTFecDesemb.value = gdFecSis
    Call HabilitaFechaFija(False)
    bSimulador = True
End Sub
Private Sub OptTipoCuota_Click(Index As Integer)
    Call LimpiaFlex(FECalend)
End Sub

Private Sub optTipoGracia_Click(Index As Integer)
    If Index = 0 Then
    Else
        If OptTipoPeriodo(1).value Then
            MsgBox "Gracia en Cuotas no es aplicable para este Periodo", vbInformation, "Mensaje"
            Exit Sub
        End If
    End If
End Sub

Private Sub OptTipoPeriodo_Click(Index As Integer)
    Call LimpiaFlex(FECalend)
    If Index = 1 Then
        Call HabilitaFechaFija(True)
        optTipoGracia(0).Enabled = False
        'Frame6.Enabled = False
        ChkPerGra.Enabled = False
        txtFechaPago.Text = DTFecDesemb.value
        GenerarFechaPago
    Else
        Call HabilitaFechaFija(False)
        optTipoGracia(0).Enabled = True
        GenerarFechaPago
        ChkPerGra.value = 0
        TxtDiaFijo.Text = "00"
    End If
End Sub
Private Sub spnCuotas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If SpnPlazo.Enabled Then
            SpnPlazo.SetFocus
        Else
            If DTFecDesemb.Enabled Then DTFecDesemb.SetFocus
        End If
    End If
End Sub

Private Sub SpnPlazo_Change()
    bGraciaGenerada = False
    GenerarFechaPago
    ChkPerGra.value = 0
End Sub

Private Sub spnPlazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And OptTipoCuota(0).Enabled Then
        OptTipoCuota(0).SetFocus
    End If
End Sub

Private Sub SSCalend_Click(PreviousTab As Integer)
Dim nTotalInteres As Double
Dim nTotalCapital As Double
Dim nTotalCONITF As Double
Dim i As Integer
'Para mostrar los totales de MiVivienda cuando no se muestre Cuota+ITF
Dim bTieneTotalITF As Boolean
'***************************

    If FECalBPag.rows <= 2 Then
        Exit Sub
        lblCapital.Caption = "0.00"
        lblInteres.Caption = "0.00"
        lblTotal.Caption = "0.00"
        lblTotalCONITF.Caption = "0.00"
    End If
    If SSCalend.Tab = 0 Then
            If FECalBPag.TextMatrix(1, 11) = "" Then bTieneTotalITF = True
            If FECalBPag.TextMatrix(1, 9) = "" Then bTieneTotalITF = True
            For i = 0 To UBound(MatCalend) - 1
                nTotalCapital = nTotalCapital + CDbl(Trim(FECalBPag.TextMatrix(i + 1, 4)))
                If Not bTieneTotalITF Then nTotalCONITF = nTotalCONITF + CDbl(Trim(FECalBPag.TextMatrix(i + 1, 11)))
                nTotalInteres = nTotalInteres + CDbl(Trim(FECalBPag.TextMatrix(i + 1, 5)))
            Next i
            lblCapital.Caption = Format(nTotalCapital, "#0.00")
            lblInteres.Caption = Format(nTotalInteres, "#0.00")
            lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#0.00")
            If Not bTieneTotalITF Then
                lblTotalCONITF.Caption = Format(nTotalCONITF, "#0.00")
            Else
                lblTotalCONITF.Caption = Format(CDbl(lblTotal.Caption) + fgITFCalculaImpuesto(CDbl(lblTotal.Caption)), "0.00")
            End If
            
    Else
            If FECalMPag.TextMatrix(1, 11) = "" Then bTieneTotalITF = True
            For i = 0 To UBound(MatCalend) - 1
                nTotalCapital = nTotalCapital + CDbl(Trim(FECalMPag.TextMatrix(i + 1, 4)))
                If Not bTieneTotalITF Then nTotalCONITF = nTotalCONITF + CDbl(Trim(FECalMPag.TextMatrix(i + 1, 11)))
                nTotalInteres = nTotalInteres + CDbl(Trim(FECalMPag.TextMatrix(i + 1, 5)))
            Next i
            lblCapital.Caption = Format(nTotalCapital, "#0.00")
            lblInteres.Caption = Format(nTotalInteres, "#0.00")
            lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#0.00")
            If Not bTieneTotalITF Then
                lblTotalCONITF.Caption = Format(nTotalCONITF, "#0.00")
            Else
                lblTotalCONITF.Caption = Format(CDbl(lblTotal.Caption) + fgITFCalculaImpuesto(CDbl(lblTotal.Caption)), "0.00")
            End If
            
    End If
End Sub
Private Sub TxtDiaFijo_Change()
    If TxtDiaFijo.Text = "" Then TxtDiaFijo.Text = "00"
    If CInt(TxtDiaFijo.Text) > 31 Then
        TxtDiaFijo.Text = "00"
    End If
    Call LimpiaFlex(FECalend)
End Sub
Private Sub TxtDiaFijo_GotFocus()
    fEnfoque TxtDiaFijo
End Sub

Private Sub TxtDiaFijo_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        txtFechaPago.SetFocus
    End If
End Sub
Private Sub TxtDiaFijo_LostFocus()
    TxtDiaFijo.Text = Right("00" + Trim(TxtDiaFijo.Text), 2)
End Sub
Private Sub txtFechaPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        nTipoGracia = 0
        If Not Trim(ValidaFecha(txtFechaPago.Text)) = "" Then
            MsgBox Trim(ValidaFecha(txtFechaPago.Text)), vbInformation, "Aviso"
            Exit Sub
        End If
        If OptTipoPeriodo(0).value = True Then

            If (Trim(SpnPlazo.valor) = "" Or SpnPlazo.valor = "0" Or SpnPlazo.valor = "00") Then
                MsgBox "Ingrese el Plazo de las Cuotas del Prestamo", vbInformation, "Aviso"
                If SpnPlazo.Enabled Then SpnPlazo.SetFocus
                Exit Sub
            End If

            If CDate(DTFecDesemb.value + SpnPlazo.valor) > CDate(txtFechaPago.Text) Then
                MsgBox "La Fecha de Pago No puede ser Menor que el Plazo", vbInformation, "Aviso"
                txtFechaPago.Text = CDate(DTFecDesemb.value + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
                txtFechaPago.SetFocus
                ChkPerGra.value = 0
                Exit Sub
            End If
            If txtFechaPago.Text > CDate(DTFecDesemb.value + SpnPlazo.valor) Then
                ChkPerGra.Enabled = True
                ChkPerGra.value = 1
                TxtPerGra.Text = CInt(CDate(txtFechaPago.Text) - CDate(DTFecDesemb.value + SpnPlazo.valor))
            Else
                TxtPerGra.Text = 0
                ChkPerGra.value = 0
            End If
        End If
        If OptTipoPeriodo(1).value = True Then
            If CDate(DTFecDesemb.value) > CDate(txtFechaPago.Text) Then
                MsgBox "La Fecha de Pago No puede ser Menor que la F. Desembolso", vbInformation, "Aviso"
                txtFechaPago.Text = CDate(DTFecDesemb.value + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
                txtFechaPago.SetFocus
                ChkPerGra.value = 0
                Exit Sub
            End If
            If Month(DTFecDesemb.value) = Month(txtFechaPago.Text) And Year(DTFecDesemb.value) = Year(txtFechaPago.Text) Then
                bProxMesN = 0
            Else
                bProxMesN = 1
            End If
            If CDate(DTFecDesemb.value + 30) < txtFechaPago.Text Then
                ChkPerGra.Enabled = True
                ChkPerGra.value = 1
                TxtPerGra.Text = CInt(CDate(txtFechaPago.Text) - CDate(DTFecDesemb.value + 30))
            Else
                ChkPerGra.value = 0
            End If
            TxtDiaFijo.Text = Right("00" + Trim(Day(txtFechaPago.Text)), 2)
        End If
        If (Me.ChkPerGra.value = 0) Then
            TxtTasaGraciaNEW.Text = "0.00"
        Else
            TxtTasaGraciaNEW.Text = Format(Me.txtInteresAnual.Text, "#0.00") 'ARLO20181029
        End If
        TxtTasaGracia.Text = nTem
        GenerarFechaPago
    End If
End Sub
Private Sub txtInteresAnual_GotFocus()
    fEnfoque txtInteresAnual
End Sub
Private Sub txtInteresAnual_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtInteresAnual, KeyAscii, , 4)
    If KeyAscii = 13 Then
        If nTpoSubPro <> 203 Then
        If (txtInteresAnual.Text = "0" Or txtInteresAnual.Text = "0.00") Then
            MsgBox "La tasa de interés tiene que ser mayor que cero", vbInformation, "Aviso"
            txtInteresAnual.SetFocus
            Exit Sub
        End If
        End If
        Me.TxtTasaGraciaNEW = Me.txtInteresAnual 'ARLO2019015
        nTem = Redondeo(((CDbl(txtInteresAnual.Text) / 100 + 1) ^ (1 / 12) * 100 - 100), 3)
        TxtInteres.Text = nTem
        txtTEM.Text = Format(nTem, "#0.00") 'ARLO20181029
        If SpnCuotas.Enabled Then
            SpnCuotas.SetFocus
        Else
            If SpnPlazo.Enabled Then SpnPlazo.SetFocus
        End If
    End If
End Sub
Private Sub txtinteres_LostFocus()
    If Trim(TxtInteres.Text) = "" Then
        TxtInteres.Text = "0.00"
    Else
        TxtInteres.Text = Format(TxtInteres.Text, "#0.0000")
    End If
End Sub
Private Sub txtMonto_GotFocus()
    fEnfoque TxtMonto
End Sub
Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtMonto, KeyAscii)
    If KeyAscii = 13 Then
        cmbMoneda.Enabled = True
        cmbMoneda.SetFocus
    End If
End Sub
Private Sub txtMonto_LostFocus()
    If Trim(TxtMonto.Text) = "" Then
        TxtMonto.Text = "0.00"
    Else
        TxtMonto.Text = Format(TxtMonto.Text, "#0.00")
    End If
End Sub
Private Sub TxtPerGra_GotFocus()
    fEnfoque TxtPerGra
End Sub
Private Sub TxtPerGra_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        If TxtTasaGracia.Enabled Then
            TxtTasaGracia.SetFocus
        Else
            cmdAplicar.SetFocus
        End If
    End If
End Sub
Private Sub TxtTasaGracia_GotFocus()
    fEnfoque TxtTasaGracia
End Sub
Private Sub TxtTasaGracia_LostFocus()
    If Trim(TxtTasaGracia.Text) = "" Then
        TxtTasaGracia.Text = "0.00"
    Else
        TxtTasaGracia.Text = Format(TxtTasaGracia.Text, "#0.0000")
    End If
End Sub
'*Función que devuelve el tipo de cuota
Private Function getTipoCuota() As Integer
Dim i As Integer
    For i = 0 To 2
        If OptTipoCuota(i).value Then
            getTipoCuota = i + 1
            Exit For
        End If
    Next i
End Function
'**Función que obtiene los gastos en la reprogramación
Private Sub ObtenerGastosEnReprogramacion()
Dim oNGasto As COMNCredito.NCOMGasto
Dim oDCredito As COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Dim nNumGastos As Integer
Dim nTotalGasto As Double
Dim i, j As Integer

        ReDim MatDesemb(1, 2)
        MatDesemb(0, 0) = Format(DTFecDesemb.value, "dd/mm/yyyy")
        MatDesemb(0, 1) = Format(TxtMonto.Text, "#0.00")
    
        Set oDCredito = New COMDCredito.DCOMCredito
        Set rsCredito = oDCredito.RecuperaDatosParaGenerarGastosEnReprog(sCtaCodRep)
        Set oDCredito = Nothing
                    
        Set oNGasto = New COMNCredito.NCOMGasto
            MatGastos = oNGasto.GeneraCalendarioGastos_NEW(CDbl(TxtMonto.Text), CDbl(TxtInteres.Text), CInt(SpnCuotas.valor), _
                                CInt(SpnPlazo.valor), CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), getTipoCuota, _
                                IIf(OptTipoPeriodo(0).value, 1, 2), nTipoGracia, CInt(TxtPerGra.Text), _
                                CDbl(TxtTasaGracia.Text), CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)), _
                                bProxMes, MatGracia, False, 0, MatCalend_2, _
                                MatDesemb, nNumGastos, gdFecSis, _
                                sCtaCodRep, 1, "RP", rsCredito("cTipoGasto"), _
                                CDbl(MatCalend(0, 2)), CDbl(TxtMonto.Text), , , , , , , , True, _
                                2, True, False, MatDesemb, False, , _
                                pnDiaFijo2, True, , _
                                gnITFMontoMin, gnITFPorcent, gbITFAplica, rsCredito("nExoSeguroDes"))
        Set oNGasto = Nothing
        Set rsCredito = Nothing
        Call frmCredReprogCred.EstablecerGastos(MatGastos, True, nNumGastos, IIf(OptTipoPeriodo(0).value, 1, 2), CInt(SpnPlazo.valor))
        '***************************************************************************
        'Adicionamos los Gastos
        '***************************************************************************
        If IsArray(MatGastos) Then
            For j = 0 To UBound(MatCalend) - 1
                nTotalGasto = 0
                For i = 0 To UBound(MatGastos) - 1
                    If Trim(Right(MatGastos(i, 0), 2)) = "1" And _
                       (Trim(MatGastos(i, 1)) = Trim(MatCalend(j, 1)) _
                         Or Trim(MatGastos(i, 1)) = "*") Then
                        nTotalGasto = nTotalGasto + CDbl(MatGastos(i, 3))
                    End If
                Next i
                MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
            Next j
        End If
    
End Sub
'**Función que obtiene los gastos de desgravamen
Private Sub ObtenerDesgravamen()
Dim oNGasto As COMNCredito.NCOMGasto
Dim oDCredito As COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Dim nNumGastos As Integer
Dim nTotalGasto As Double
Dim nTotalGastoSeg As Double
Dim i, j As Integer
Dim oGastos As COMDCredito.DCOMGasto
Dim RGastosSegDes As ADODB.Recordset
Dim RGastosEnvCue As ADODB.Recordset
Dim nMunMesesPorDia, nNumMesesPorMes As Integer
                    
Set oGastos = New COMDCredito.DCOMGasto

Dim nSegDes As Double

nSegDes = 0

If fbInicioSim And fraGastoCom.Visible Then
    Set RGastosSegDes = oGastos.RecuperaGastosCabecera(1)
    RGastosSegDes.Filter = " nPrdConceptoCod = 1217"
    nSegDes = IIf(Trim(Right(cmbSeguroDes.Text, 4)) = "1", RGastosSegDes!nValor, RGastosSegDes!nValorDosTit)
    nSegDes = 0
    
    Set RGastosEnvCue = oGastos.RecuperaGastosCabecera(1)
    RGastosEnvCue.Filter = " nPrdConceptoCod = 1249"
End If

Set oGastos = Nothing

        ReDim MatDesemb(1, 2)
        MatDesemb(0, 0) = Format(DTFecDesemb.value, "dd/mm/yyyy")
        MatDesemb(0, 1) = Format(TxtMonto.Text, "#0.00")
    

        Set oNGasto = New COMNCredito.NCOMGasto
            MatGastos = oNGasto.GeneraCalendarioGastos_NEW(CDbl(TxtMonto.Text), CDbl(TxtInteres.Text), CInt(SpnCuotas.valor), _
                                CInt(SpnPlazo.valor), CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), getTipoCuota, _
                                IIf(OptTipoPeriodo(0).value, 1, 2), nTipoGracia, CInt(TxtPerGra.Text), _
                                CDbl(TxtTasaGracia.Text), CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)), _
                                0, MatGracia, 0, 0, MatCalend_2, _
                                MatDesemb, nNumGastos, gdFecSis, _
                                sCtaCodRep, 1, "SI", "F", _
                                CDbl(MatCalend(0, 2)), CDbl(TxtMonto.Text), , , , , , , , True, _
                                2, True, False, MatDesemb, False, , _
                                0, , , _
                                gnITFMontoMin, gnITFPorcent, gbITFAplica, 0, , , , , , nIntGraInicial, , , 0)
        Set oNGasto = Nothing
        Set rsCredito = Nothing
        Call frmCredReprogCred.EstablecerGastos(MatGastos, True, nNumGastos, IIf(OptTipoPeriodo(0).value, 1, 2), CInt(SpnPlazo.valor))
        '***************************************************************************
        'Adicionamos los Gastos
        '***************************************************************************
               
        If IsArray(MatGastos) Then
            For j = 0 To UBound(MatCalend) - 1
                nTotalGasto = 0
                nTotalGastoSeg = 0
                If fbInicioSim And fraGastoCom.Visible Then
                    
                    'RIRO 20200313 El seguro desgravamen, ya es calculado en la genración de calendario.
                    'If j = 0 Or CInt(MatCalend(j, 1)) = 1 Then
                    '    nMunMesesPorDia = Round(DateDiff("d", DTFecDesemb.value, MatCalend(j, 0)) / 30, 0)
                    '    nNumMesesPorMes = DateDiff("m", DTFecDesemb.value, MatCalend(j, 0))
                    '    If Not Me.cmbSeguroDes.Enabled Then
                    '        MatCalend(j, 6) = "0.00"
                    '    Else
                    '        MatCalend(j, 6) = Format(TxtMonto.Text * (nSegDes / 100) * IIf(nMunMesesPorDia >= nNumMesesPorMes, nMunMesesPorDia, nNumMesesPorMes), "#0.00")
                    '    End If
                    'Else
                    '    If Not Me.cmbSeguroDes.Enabled Then
                    '         MatCalend(j, 6) = "0.00"
                    '    Else
                    '        MatCalend(j, 6) = Format(MatCalend(j - 1, 7) * (nSegDes / 100), "#0.00")
                    '    End If
                    'End If
                    'FECalend.EncabezadosNombres = "-Fecha Venc.-Cuota-Cuota + ITF-Cuotas-Capital-Interes-Int. Gracia-Gast/Comis-Seg.Desg(1)-Seg. Mult.-Saldo Capital"
                    
                    If chkEnvioEst = 1 Then
                        If Trim(Right(Me.cmbEnvioEst.Text, 4)) = "2" Then
                            'FECalend.EncabezadosNombres = "-Fecha Venc.-Cuota-Cuota + ITF-Cuotas-Capital-Interes-Int. Gracia-Gast/Com(2)-Seg.Desg(1)-Seg. Mult.-Saldo Capital"
                            'MatCalend(j, 8) = Format(RGastosEnvCue!nValor, "#0.00")
                            MatCalend(j, 6) = Format(CDbl(MatCalend(j, 6)) + CDbl(RGastosEnvCue!nValor), "#0.00")
                        End If
                    End If
                Else
                    For i = 0 To UBound(MatGastos) - 1
                        If Trim(Right(MatGastos(i, 0), 2)) = "1" And _
                           (Trim(MatGastos(i, 1)) = Trim(MatCalend(j, 1))) Then
                            nTotalGasto = CDbl(MatGastos(i, 3))
                            MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
                        Else
                            If Trim(MatGastos(i, 1)) = "*" Then
                                nTotalGasto = CDbl(MatGastos(i, 3))
                                MatCalend(j, 8) = Format(nTotalGasto, "#0.00")
                            End If
    
                        End If
                        
                    Next i
                End If
            Next j
        End If
               
        For i = 0 To UBound(MatCalend) - 1
            If (i = 0) And nTipoGracia <> gColocTiposGraciaCapitalizada Then
                'MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 8)) + (nSegMutlUnidad * nMultiplicaSegMutl))
                MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 16)) + CDbl(MatCalend(i, 8)), "#0.00")
                
            ElseIf (i = 1) And nTipoGracia = gColocTiposGraciaCapitalizada Then
                'MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 8)) + (nSegMutlUnidad * nMultiplicaSegMutl))
                MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 16)) + CDbl(MatCalend(i, 8)), "#0.00")
                
            Else
                'MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 8)) + nSegMutlUnidad)
                MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 16)) + CDbl(MatCalend(i, 8)), "#0.00")
            End If
        Next i

End Sub
Private Sub ObtenerDesgravamenHipot(Optional ByVal pnValorInmueble As Double, Optional ByVal pnTramoConsMonto As Double, Optional pnMontoMivivienda As Currency = 0#, Optional ByVal pnCuotaMivienda As Double = -1)
Dim oNGasto As COMNCredito.NCOMGasto
Dim oDCredito As COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Dim nNumGastos As Integer
Dim nTotalGasto As Double
Dim nTotalGastoSeg As Double
Dim i, j As Integer

        ReDim MatDesemb(1, 2)
        MatDesemb(0, 0) = Format(DTFecDesemb.value, "dd/mm/yyyy")
        MatDesemb(0, 1) = Format(TxtMonto.Text, "#0.00")
    
        Set oNGasto = New COMNCredito.NCOMGasto
        MatGastos = oNGasto.GeneraCalendarioGastos_NEW(CDbl(TxtMonto.Text), CDbl(TxtInteres.Text), CInt(SpnCuotas.valor), _
                                CInt(SpnPlazo.valor), CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), getTipoCuota, _
                                IIf(OptTipoPeriodo(0).value, 1, 2), nTipoGracia, CInt(TxtPerGra.Text), _
                                CDbl(TxtTasaGracia.Text), CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)), _
                                bProxMes, MatGracia, True, 0, MatCalend_2, _
                                MatDesemb, nNumGastos, gdFecSis, _
                                sCtaCodRep, 1, "SI", "F", _
                                CDbl(MatCalend(0, 2)), IIf(True, pnTramoConsMonto, CDbl(TxtMonto.Text)), , , , , , , , True, _
                                2, True, False, MatDesemb, False, , _
                                pnDiaFijo2, True, , _
                                pnCuotaMivienda, gnITFPorcent, gbITFAplica, 0, , , , IIf(True, Left(cmbSeguroDes.Text, 1), ""), pnValorInmueble, , , , , , , , pnMontoMivivienda)
        Set oNGasto = Nothing
        Set rsCredito = Nothing
        Call frmCredReprogCred.EstablecerGastos(MatGastos, True, nNumGastos, IIf(OptTipoPeriodo(0).value, 1, 2), CInt(SpnPlazo.valor))
                
        If IsArray(MatGastos) Then
            For j = 0 To UBound(MatCalend) - 1
                nTotalGasto = 0
                nTotalGastoSeg = 0
                For i = 0 To UBound(MatGastos) - 1
                
                'If ChkCalMiViv.value = 0 Then
                If nTpoPro <> 3 Then
                    If Trim(Right(MatGastos(i, 0), 2)) = "1" And _
                       (Trim(MatGastos(i, 1)) = Trim(MatCalend(j, 1))) Then
                        nTotalGasto = CDbl(MatGastos(i, 3))
                        MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
                    Else
                        If Trim(MatGastos(i, 1)) = "*" Then
                            nTotalGasto = CDbl(MatGastos(i, 3))
                            MatCalend(j, 8) = Format(nTotalGasto, "#0.00")
                        End If
                    End If
                Else
                    If Trim(Right(MatGastos(i, 0), 2)) = "1" And _
                       (Trim(MatGastos(i, 1)) = Trim(MatCalend(j, 1))) Then
                        If Left(Trim(MatGastos(i, 2)), 2) <> "PO" Then
                            nTotalGasto = CDbl(MatGastos(i, 3))
                            MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
                        Else
                            nTotalGasto = CDbl(MatGastos(i, 3))
                            MatCalend(j, 9) = Format(nTotalGasto, "#0.00")
                        End If
                    Else
                        If Trim(MatGastos(i, 1)) = "*" Then
                            If Left(Trim(MatGastos(i, 2)), 2) = "CO" And i <> UBound(MatGastos) - 1 Then
                                nTotalGasto = CDbl(MatGastos(i, 3))
                                MatCalend(j, 8) = Format(nTotalGasto, "#0.00")
                            Else
                                If i <> UBound(MatGastos) - 1 Then
                                    nTotalGasto = CDbl(MatGastos(i, 3))
                                    MatCalend(j, 9) = Format(nTotalGasto, "#0.00")
                                End If
                            End If
                        Else
                            If Left(Trim(MatGastos(i, 2)), 2) <> "PO" And _
                               (Trim(MatGastos(i, 1)) = Trim(MatCalend(j, 1))) Then
                                nTotalGasto = CDbl(MatGastos(i, 3))
                                MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
                            Else
                                If (Trim(MatGastos(i, 1)) = Trim(MatCalend(j, 1))) Then
                                    nTotalGasto = CDbl(MatGastos(i, 3))
                                    MatCalend(j, 9) = Format(nTotalGasto, "#0.00")
                                End If
                            End If
                        End If
                    End If
                End If
                '***
                Next i
            Next j
        End If
        
        For i = 0 To UBound(MatCalend) - 1
            MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 9)), "#0.00")
        Next i
End Sub
Private Sub GenerarFechaPago()
    If OptTipoPeriodo(0).value = True Then
        txtFechaPago.Text = CDate(DTFecDesemb.value + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor) + TxtPerGra.Text)
    End If
    If OptTipoPeriodo(1).value = True Then
        If SpnPlazo.Enabled = True Then
            txtFechaPago.Text = CDate(DTFecDesemb.value)
            TxtDiaFijo.Text = Right("00" + Trim(Day(txtFechaPago.Text)), 2)
                                        
            If Month(DTFecDesemb.value) = Month(CDate(txtFechaPago.Text)) And Year(gdFecSis) = Year(CDate(txtFechaPago.Text)) Then
                bProxMes = 0
            Else
                bProxMes = 1
            End If
        End If
    End If
    If (CDate(DateAdd("D", 30, DTFecDesemb.value)) < CDate(txtFechaPago.Text)) Then  'ARLO20180809
        fraGracia.Enabled = True
    Else
        fraGracia.Enabled = False
    End If
End Sub
Public Sub InicioSim()
Dim oCons As COMDConstantes.DCOMConstantes
Dim oCred As COMDCredito.DCOMCredito
Dim nId As Integer
        
fraGastoCom.Visible = True
fbInicioSim = True
Set oCons = New COMDConstantes.DCOMConstantes
Set oCred = New COMDCredito.DCOMCredito

Call Llenar_Combo_con_Recordset(oCons.RecuperaConstantes(7096), cmbSeguroDes)
Call Llenar_Combo_con_Recordset(oCons.RecuperaConstantes(9110), cmbEnvioEst)
'Call Llenar_Combo_con_Recordset(oCons.RecuperaConstantes(10091), cmbProductoCMACM)
Call Llenar_Combo_con_Recordset(oCons.RecuperaConstantes(1011), cmbMoneda)
Call Llenar_Combo_con_Recordset(oCons.RecuperaConstantes(10093), cmbTpoCliente)
Call Llenar_Combo_con_Recordset(oCons.RecuperaConstantes(10094), cmbHipoteca)

If Not (bExisteModal) Then 'ARLO20180809
    'Call Llenar_Combo_con_Recordset(oCons.RecuperaConstantes(10091), cmbProductoCMACM) 'ARLO20190229 ERS042-2018 COMENTADO
    Call Llenar_Combo_con_Recordset(oCred.RecuperaProductosCrediticios, cmbProductoCMACM) 'ARLO20190229 ERS042-2018
    Me.Show 1
End If

End Sub
Private Sub cmbProductoCMACM_Click()
    Call CargaSubProducto(Trim(Right(cmbProductoCMACM.Text, 3)))
    nTpoPro = CInt(Trim(Right(cmbProductoCMACM.Text, 3)))
    cTpoPro = Trim(Left(cmbProductoCMACM.Text, 15))
    If (Me.cmbProductoCMACM.ListIndex = -1) Then
        Me.cmbSubProducto.Enabled = False
    Else
        Me.cmbSubProducto.Enabled = True
    End If
End Sub
Private Sub cmbProductoCMACM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbSubProducto.SetFocus
    End If
End Sub
Private Sub CargaSubProducto(ByVal psTipo As String)
Dim oCred As COMDCredito.DCOMCredito
Dim ssql As String
Dim RTemp As ADODB.Recordset
    On Error GoTo ERRORCargaSubProducto
    Set oCred = New COMDCredito.DCOMCredito
    'Set RTemp = oCred.RecuperaSubProductosCrediticiosNEW(psTipo) 'ARLO20190229 ERS042-2018 COMENTADO
    Set RTemp = oCred.RecuperaSubProductosCrediticios(psTipo, "", 0, 0) 'ARLO20190229 ERS042-2018
    Set oCred = Nothing
    cmbSubProducto.Clear
    Do While Not RTemp.EOF
        cmbSubProducto.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    Call CambiaTamañoCombo(cmbSubProducto, 250)
    Exit Sub
    
ERRORCargaSubProducto:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub
Private Sub cmbSubProducto_Click()
    nTpoSubPro = CInt(Trim(Right(cmbSubProducto.Text, 3)))
    cTpoSubPro = Trim(Left(cmbSubProducto.Text, 30))
    
    bExisteModal = True
    
    'Call CargaControles
    LimpiaFlex FECalend
    TxtMonto.Text = "0.00"
    TxtInteres.Text = "0.00"
    SpnCuotas.valor = 0
    SpnPlazo.valor = 0
    cmbMoneda.ListIndex = -1
    TxtMonto.SetFocus
    cmbTpoCliente.ListIndex = -1
    cmdImprimir.Enabled = False
    cmdResumen.Visible = False
    fraTasaAnuales.Left = 9120  'ARLO20180809
    fraTasaAnuales.top = 2160  'ARLO20180809
    FraFechaPago.Left = 3960  'ARLO20180809
    FraFechaPago.top = 1110  'ARLO20180809
    FraTpoCliente.Left = 6480  'ARLO20180809
    FraTpoCliente.top = 240  'ARLO20180809
    FraHipoteca.Visible = True 'ARLO20180809
    Me.fraGracia.Visible = True 'ARLO20180809
    Me.fraEnvioEst.Visible = True  'ARLO20181029
    Me.fraGastoCom.Visible = True  'ARLO20181029
    Me.FECalend.Visible = True     'ARLO20181029
    Call InicioSim
    
    If (nTpoSubPro = 203) Then 'ARLO20180809 '203 PIGNORATICIOS
    'If (nTpoSubPro = 203 Or nTpoSubPro = 201) Then
        Me.cmbTpoCliente.Enabled = True
        Me.FraHipoteca.Enabled = False
        Me.chkHipoteca.Enabled = False
        Me.TxtInteres.Enabled = False
        Me.SpnCuotas.Enabled = False
        Me.SpnPlazo.Enabled = False
        FraTipoCuota.Enabled = False
        FraTipoPeriodo.Enabled = False
        fraGracia.Enabled = False
        FraFechaPago.Enabled = False
        fraGastoCom.Enabled = False
        FraPigno.Visible = False
        fraTasaAnuales.Enabled = True
        fraTasaAnuales.Visible = False
        FraFechaPago.Visible = False
        FraTpoCliente.Left = 3960   'ARLO20180809
        FraTpoCliente.top = 1110    'ARLO20180809
        fraGracia.Visible = False   'ARLO20180809
        FraHipoteca.Visible = False 'ARLO20180809
        FraPigno.top = 240          'ARLO20180809
        FraPigno.Left = 6900        'ARLO20180809
        fraTasaAnuales.Left = 6900  'ARLO20180809
        fraTasaAnuales.top = 2100  'ARLO20180809
        cmbHipoteca.ListIndex = -1
        txtEdificacion.Text = ""
        txtInteresAnual.Enabled = False
        txtInteresAnual.Text = ""
        FECalend.Enabled = False 'ARLO20180809
        cmbMoneda.RemoveItem (1) 'ARLO20180809
        txtTEM.Enabled = True 'ARLO20181029
        Me.fraEnvioEst.Visible = False  'ARLO20181029
        Me.fraGastoCom.Visible = False  'ARLO20181029
        'Me.FECalend.Visible = False     'ARLO20181029
    ElseIf (nTpoSubPro = 201) Then
            Me.cmbTpoCliente.Enabled = False
            Me.txtInteresAnual.Enabled = True
            Me.SpnCuotas.Enabled = True
            Me.SpnPlazo.Enabled = True
            Me.FraFechaPago.Visible = True
            Me.FraHipoteca.Enabled = True       'ARLO20180809
            Me.chkHipoteca.Enabled = True       'ARLO20180809
            Me.FraTipoPeriodo.Enabled = True    'ARLO20180809
            Me.FraFechaPago.Enabled = True      'ARLO20180809
            Me.FraPigno.Visible = False         'ARLO20180809
            Me.fraGastoCom.Enabled = True       'ARLO20180809
            Me.fraEnvioEst.Enabled = True       'ARLO20180809
            Me.fraTasaAnuales.Visible = False   'ARLO20180809

    Else
        Me.cmbTpoCliente.Enabled = False
        Me.FraHipoteca.Enabled = True
        Me.chkHipoteca.Enabled = True
        'Me.TxtInteres.Enabled = True
        Me.SpnCuotas.Enabled = True
        FraTipoCuota.Enabled = True
        FraTipoPeriodo.Enabled = True
        fraGracia.Enabled = True
        FraTipoCuota.Visible = True
        FraFechaPago.Enabled = True
        fraTasaAnuales.Enabled = True
        fraTasaAnuales.Visible = False
        FraPigno.Visible = False
        Me.FraFechaPago.Visible = True
        Me.SpnPlazo.Enabled = True
        Me.fraGastoCom.Enabled = True       'ARLO20180809
        Me.fraEnvioEst.Enabled = True       'ARLO20180809
        Me.txtInteresAnual.Enabled = True
        If (Me.chkHipoteca.value = 1) Then
            Me.FraHipoteca.Enabled = True
        End If
    End If
End Sub
Private Sub txtEdificacion_GotFocus()
    fEnfoque txtEdificacion
End Sub
Private Sub txtEdificacion_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtEdificacion, KeyAscii, , 4)
txtEdificacion.MaxLength = "10"
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtEdificacion.MaxLength = "10"
        txtEdificacion.Text = Format(txtEdificacion.Text, "#,#00.00")
        Me.cmdAplicar.SetFocus
    End If
End Sub
Private Sub txtEdificacion_LostFocus()
    If Trim(txtEdificacion.Text) = "" Then
        txtEdificacion.Text = "0.00"
    Else
        txtEdificacion.Text = Format(txtEdificacion.Text, "#0.00")
    End If
End Sub
Public Function fin_del_Mes(Fecha As Variant) As Date
 
    If IsDate(Fecha) Then
        fin_del_Mes = DateAdd("m", 1, Fecha)
        fin_del_Mes = DateSerial(Year(fin_del_Mes), Month(fin_del_Mes), 1)
        fin_del_Mes = DateAdd("d", -1, fin_del_Mes)
    End If
 
End Function
Private Function Redondeo(ByVal Numero, ByVal Decimales)
    Redondeo = Int(Numero * 10 ^ Decimales + 1 / 2) / 10 ^ Decimales
End Function
Private Sub Form_Unload(Cancel As Integer)
    bExisteModal = False
End Sub
