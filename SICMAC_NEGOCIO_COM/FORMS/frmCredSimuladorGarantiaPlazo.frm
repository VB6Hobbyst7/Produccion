VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredSimuladorGarantiaPlazo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simulador de Crédito con Garantía Plazo Fijo"
   ClientHeight    =   11205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11310
   Icon            =   "frmCredSimuladorGarantiaPlazo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11205
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "Aplicar"
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
      Left            =   3000
      TabIndex        =   71
      Top             =   10680
      Width           =   1575
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
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
      Left            =   4800
      TabIndex        =   57
      ToolTipText     =   "Imprimir el Calendario de Pagos"
      Top             =   10665
      Width           =   1455
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
      Height          =   375
      Left            =   6480
      TabIndex        =   56
      ToolTipText     =   "Salir del Calendario de Pagos"
      Top             =   10665
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   " Calendario de Pagos "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      TabIndex        =   22
      Top             =   4680
      Width           =   11055
      Begin VB.Frame FraFechaPago 
         Height          =   735
         Left            =   7560
         TabIndex        =   68
         Top             =   1560
         Width           =   1335
         Begin MSMask.MaskEdBox txtFechaPago 
            Height          =   315
            Left            =   120
            TabIndex        =   69
            ToolTipText     =   "Presione Enter"
            Top             =   360
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
            Left            =   120
            TabIndex        =   70
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame FraTipoCuota 
         Caption         =   " Tipo Cuota "
         ForeColor       =   &H80000007&
         Height          =   2025
         Left            =   3360
         TabIndex        =   52
         Top             =   240
         Width           =   1545
         Begin VB.OptionButton OptTipoCuota 
            Caption         =   "Fijo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton OptTipoCuota 
            Caption         =   "Creciente"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   54
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton OptTipoCuota 
            Caption         =   "Decreciente"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   53
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.Frame FraTipoPeriodo 
         Caption         =   " Tipo Periodo "
         ForeColor       =   &H80000007&
         Height          =   2040
         Left            =   5040
         TabIndex        =   43
         Top             =   240
         Width           =   2385
         Begin VB.OptionButton OptTipoPeriodo 
            Caption         =   "Periodo Fijo"
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   51
            Top             =   285
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.Frame Frame6 
            Height          =   820
            Left            =   150
            TabIndex        =   45
            Top             =   960
            Width           =   2025
            Begin VB.CheckBox ChkProxMes 
               Caption         =   "Prox Mes"
               Enabled         =   0   'False
               Height          =   210
               Left            =   1020
               TabIndex        =   48
               Top             =   180
               Width           =   960
            End
            Begin VB.TextBox TxtDiaFijo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   615
               MaxLength       =   2
               TabIndex        =   47
               Top             =   150
               Width           =   330
            End
            Begin VB.TextBox TxtDiaFijo2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   615
               MaxLength       =   2
               TabIndex        =   46
               Text            =   "00"
               Top             =   480
               Width           =   330
            End
            Begin VB.Label LblDia 
               AutoSize        =   -1  'True
               Caption         =   "&Dia 1:"
               Enabled         =   0   'False
               Height          =   195
               Left            =   90
               TabIndex        =   50
               Top             =   180
               Width           =   420
            End
            Begin VB.Label lblDia2 
               AutoSize        =   -1  'True
               Caption         =   "&Dia 2:"
               Height          =   195
               Left            =   90
               TabIndex        =   49
               Top             =   510
               Width           =   420
            End
         End
         Begin VB.OptionButton OptTipoPeriodo 
            Caption         =   "Fecha Fija"
            Height          =   315
            Index           =   1
            Left            =   150
            TabIndex        =   44
            Top             =   600
            Width           =   1035
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Gracia "
         Height          =   1305
         Left            =   7560
         TabIndex        =   35
         Top             =   240
         Width           =   2385
         Begin VB.CheckBox ChkPerGra 
            Caption         =   "Periodo &Gracia"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox TxtPerGra 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1590
            MaxLength       =   4
            TabIndex        =   39
            Text            =   "0"
            Top             =   225
            Width           =   510
         End
         Begin VB.CommandButton CmdGracia 
            Caption         =   "-->"
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
            Left            =   1680
            TabIndex        =   38
            Top             =   570
            Width           =   405
         End
         Begin VB.TextBox TxtTasaGracia 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   660
            MaxLength       =   7
            TabIndex        =   37
            Top             =   600
            Width           =   615
         End
         Begin VB.OptionButton optTipoGracia 
            Caption         =   "Capitalizar"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   36
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label LblTasaGracia 
            Caption         =   "Tasa :"
            Height          =   165
            Left            =   150
            TabIndex        =   42
            Top             =   615
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
            Left            =   1320
            TabIndex        =   41
            Top             =   630
            Width           =   150
         End
      End
      Begin VB.Frame FraDatos 
         Caption         =   " Condiciones "
         ForeColor       =   &H80000007&
         Height          =   1995
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   3135
         Begin VB.TextBox TxtMonto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            MaxLength       =   9
            TabIndex        =   25
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox TxtInteres 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            MaxLength       =   7
            TabIndex        =   24
            Top             =   550
            Width           =   615
         End
         Begin Spinner.uSpinner SpnCuotas 
            Height          =   255
            Left            =   1680
            TabIndex        =   26
            Top             =   890
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
            Left            =   1680
            TabIndex        =   27
            Top             =   1530
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   393216
            Format          =   91422721
            CurrentDate     =   37054
         End
         Begin Spinner.uSpinner SpnPlazo 
            Height          =   285
            Left            =   1680
            TabIndex        =   28
            Top             =   1200
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
         Begin VB.Label Label18 
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
            Left            =   2340
            TabIndex        =   34
            Top             =   580
            Width           =   150
         End
         Begin VB.Label Label19 
            Caption         =   "Fecha Desembolso"
            Height          =   435
            Left            =   240
            TabIndex        =   33
            Top             =   1470
            Width           =   885
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Periodo (Dias)"
            Height          =   195
            Left            =   240
            TabIndex        =   32
            Top             =   1230
            Width           =   990
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Nro Cuotas"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   910
            Width           =   795
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Interes (Mensual)"
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   580
            Width           =   1215
         End
         Begin VB.Label lblmonto 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
            Height          =   195
            Left            =   240
            TabIndex        =   29
            Top             =   285
            Width           =   450
         End
      End
      Begin SICMACT.FlexEdit FECalend 
         Height          =   2865
         Left            =   120
         TabIndex        =   66
         Top             =   2400
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   5054
         Cols0           =   11
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-Fecha Venc.-Cuota-Cuotas-Capital-Interes-Int. Gracia-Gast/Comis-Seg. Desg-Saldo Capital-Cuota + ITF"
         EncabezadosAnchos=   "400-1000-600-1200-1000-1000-1000-1000-1000-1200-1000"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-R-R-R-R-R-R-R-C"
         FormatosEdit    =   "0-0-0-2-3-2-2-2-2-2-2"
         SelectionMode   =   1
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Garantías Disponibles "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   335
         Left            =   120
         TabIndex        =   67
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   " Totales "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   4080
         TabIndex        =   7
         Top             =   3000
         Width           =   6855
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Dólares :"
            Height          =   195
            Left            =   75
            TabIndex        =   21
            Top             =   870
            Width           =   630
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Soles :"
            Height          =   195
            Left            =   75
            TabIndex        =   20
            Top             =   480
            Width           =   480
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Linea Disponible"
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
            Left            =   5160
            TabIndex        =   19
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Cred."
            Height          =   195
            Left            =   3720
            TabIndex        =   18
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Lim. Cobertura"
            Height          =   195
            Left            =   2280
            TabIndex        =   17
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Saldo"
            Height          =   195
            Left            =   840
            TabIndex        =   16
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lblDolLineaDisp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5160
            TabIndex        =   15
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label lblDolSaldoCred 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   3720
            TabIndex        =   14
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblDolLimCobert 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   2280
            TabIndex        =   13
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblDolSaldo 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   840
            TabIndex        =   12
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblSolLineaDisp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5160
            TabIndex        =   11
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblSolSaldoCred 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   3720
            TabIndex        =   10
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblSolLimCobert 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   2280
            TabIndex        =   9
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblSolSaldo 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   840
            TabIndex        =   8
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   335
         Left            =   9720
         TabIndex        =   5
         Top             =   280
         Width           =   1095
      End
      Begin SICMACT.TxtBuscar TxtBCodPers 
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Top             =   300
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
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
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin SICMACT.FlexEdit fePlazoFijo 
         Height          =   2055
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   3625
         Cols0           =   11
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nro-Nro Cuenta-Sub Producto-Moneda-Saldo MN-% Cobert-Limite Cob.-Linea Disponible-Saldo Cred-Otros Bloq-Detalle"
         EncabezadosAnchos=   "400-1800-2200-800-1200-800-1200-1300-1200-1200-800"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "R-L-L-C-R-R-R-R-R-R-C"
         FormatosEdit    =   "3-0-0-0-2-0-2-2-2-2-0"
         TextArray0      =   "Nro"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblTipoCamb 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   3165
         TabIndex        =   73
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "TC:"
         Height          =   195
         Left            =   2880
         TabIndex        =   72
         Top             =   3885
         Width           =   255
      End
      Begin VB.Label lblNomPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4320
         TabIndex        =   4
         Top             =   300
         Width           =   5055
      End
      Begin VB.Label lblDocPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2760
         TabIndex        =   3
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   335
         Width           =   525
      End
   End
   Begin VB.Label Label26 
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
      Left            =   2700
      TabIndex        =   65
      Top             =   10200
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
      Left            =   6600
      TabIndex        =   64
      Top             =   10200
      Width           =   1410
   End
   Begin VB.Label Label25 
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
      Left            =   6090
      TabIndex        =   63
      Top             =   10200
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
      Left            =   4500
      TabIndex        =   62
      Top             =   10200
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
      Left            =   1140
      TabIndex        =   61
      Top             =   10200
      Width           =   1335
   End
   Begin VB.Label Label24 
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
      Left            =   480
      TabIndex        =   60
      Top             =   10200
      Width           =   585
   End
   Begin VB.Label Label23 
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
      Left            =   8325
      TabIndex        =   59
      Top             =   10200
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
      Left            =   9210
      TabIndex        =   58
      Top             =   10200
      Width           =   1410
   End
End
Attribute VB_Name = "frmCredSimuladorGarantiaPlazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmCredSimuladorGarantiaPlazo
'** Descripción : Formulario para simular generación de creditos Rapiflash según TI-ERS138-2013
'** Creación : JUEZ, 20140103 03:30:00 PM
'*****************************************************************************************************

Option Explicit

Dim oDPers As COMDPersona.DCOMPersonas
Dim rs As ADODB.Recordset

'Variables Calendario
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

Dim bErrorValidacion As Boolean

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

Private Sub CmdBuscar_Click()
    Dim oNCred As COMNCredito.NCOMCredito
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim lnFila As Integer
    Dim nTpoCambCompra As Double, nSumSaldoMN As Double, nSumSaldoCredMN As Double, nSumLimiteCob As Double, nSumLineaDisp As Double
    Dim nSaldoCredMN As Double, nLimCobert As Double, nLineaDisp As Double
    Dim nMontoCancMN As Double, nIntGanadoMN As Double
    
    Set oNCred = New COMNCredito.NCOMCredito
    Set rs = oNCred.RecuperaDPFParaGarantia(TxtBCodPers.Text)
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            fePlazoFijo.AdicionaFila
            lnFila = fePlazoFijo.row
            fePlazoFijo.TextMatrix(lnFila, 1) = rs!cCtaCod
            fePlazoFijo.TextMatrix(lnFila, 2) = rs!cSubProducto
            fePlazoFijo.TextMatrix(lnFila, 3) = rs!cMoneda
            fePlazoFijo.TextMatrix(lnFila, 4) = Format(rs!nSaldoMN, "#,##0.00")
            fePlazoFijo.TextMatrix(lnFila, 5) = rs!nCobertura
            fePlazoFijo.TextMatrix(lnFila, 9) = Format(rs!nBloqueoParcial, "#,##0.00")
            fePlazoFijo.TextMatrix(lnFila, 10) = "Ver"
            
            Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                nMontoCancMN = clsCap.GetSaldoCancelacion(rs!cCtaCod, gdFecSis, gsCodAge) * IIf(Mid(rs!cCtaCod, 9, 1) = "1", 1, CDbl(rs!nTpoCambVenta))
                nIntGanadoMN = nMontoCancMN - rs!nSaldoMN
            Set clsCap = Nothing
            nLimCobert = (CDbl(rs!nSaldoMN) - (IIf(CInt(rs!nformaretiro) = 4, -nIntGanadoMN, 0))) * (rs!nPorcCobert)
            nSaldoCredMN = CalculaSaldoCredMN(rs!cCtaCod, CDbl(rs!nTpoCambVenta))
            nLineaDisp = nLimCobert - nSaldoCredMN - rs!nBloqueoParcial
            
            fePlazoFijo.TextMatrix(lnFila, 6) = Format(nLimCobert, "#,##0.00")
            fePlazoFijo.TextMatrix(lnFila, 7) = Format(nLineaDisp, "#,##0.00")
            fePlazoFijo.TextMatrix(lnFila, 8) = Format(nSaldoCredMN, "#,##0.00")
            
            nSumSaldoMN = nSumSaldoMN + CDbl(rs!nSaldoMN)
            nSumLimiteCob = nSumLimiteCob + nLimCobert
            nSumSaldoCredMN = nSumSaldoCredMN + nSaldoCredMN
            nSumLineaDisp = nSumLineaDisp + nLineaDisp
            nTpoCambCompra = rs!nTpoCambCompra
            rs.MoveNext
        Loop
        
        lblSolSaldo.Caption = Format(nSumSaldoMN, "#,##0.00")
        lblDolSaldo.Caption = Format(nSumSaldoMN / nTpoCambCompra, "#,##0.00")
        lblSolLimCobert.Caption = Format(nSumLimiteCob, "#,##0.00")
        lblDolLimCobert.Caption = Format(nSumLimiteCob / nTpoCambCompra, "#,##0.00")
        lblSolSaldoCred.Caption = Format(nSumSaldoCredMN, "#,##0.00")
        lblDolSaldoCred.Caption = Format(nSumSaldoCredMN / nTpoCambCompra, "#,##0.00")
        lblSolLineaDisp.Caption = Format(nSumLineaDisp, "#,##0.00")
        lblDolLineaDisp.Caption = Format(nSumLineaDisp / nTpoCambCompra, "#,##0.00")
        
        lblTipoCamb.Caption = Format(nTpoCambCompra, "#,##0.000")
        
        cmdBuscar.Enabled = False
        cmdCancelar.Enabled = True
    Else
        MsgBox "El cliente no posee Plazos Fijos Activos", vbInformation, "Aviso"
        LimpiarControles
        LimpiarControlesSimulador
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub fePlazoFijo_Click()
    If fePlazoFijo.TextMatrix(fePlazoFijo.row, fePlazoFijo.Col) <> "" Then
        If fePlazoFijo.Col = 10 Then
            Dim MatSaldos() As String
            
            ReDim MatSaldos(3, 1)
            MatSaldos(0, 1) = fePlazoFijo.TextMatrix(fePlazoFijo.row, 4)
            MatSaldos(1, 1) = fePlazoFijo.TextMatrix(fePlazoFijo.row, 7)
            MatSaldos(2, 1) = fePlazoFijo.TextMatrix(fePlazoFijo.row, 8)
            MatSaldos(3, 1) = fePlazoFijo.TextMatrix(fePlazoFijo.row, 9)
            frmCredConsultaGarantDPF.Inicio CStr(fePlazoFijo.TextMatrix(fePlazoFijo.row, 1)), MatSaldos, fePlazoFijo.TextMatrix(fePlazoFijo.row, 5), fePlazoFijo.TextMatrix(fePlazoFijo.row, 2)
        End If
    End If
End Sub

Private Sub Form_Load()
    Call CargaControles
    bGraciaGenerada = False
    DTFecDesemb.value = gdFecSis
    Call HabilitaFechaFija(False)
End Sub

Private Sub TxtBCodPers_EmiteDatos()
    If Trim(TxtBCodPers.Text) = "" Then
        Exit Sub
    End If
    
    Set oDPers = New COMDPersona.DCOMPersonas
    Set rs = oDPers.RecuperaDatosPersona_Basic(TxtBCodPers.Text)
    lblNomPers.Caption = rs("cPersNombre")
    lblDocPers.Caption = rs("nDOINro")
    Set rs = Nothing
    Set oDPers = Nothing
    TxtBCodPers.Enabled = False
    cmdBuscar.Enabled = True
    cmdBuscar.SetFocus
End Sub

Private Sub LimpiarControles()
    Call LimpiaFlex(fePlazoFijo)
    TxtBCodPers.Text = ""
    lblDocPers.Caption = ""
    lblNomPers.Caption = ""
    Call LimpiaFlex(fePlazoFijo)
    lblSolSaldo.Caption = ""
    lblSolLimCobert.Caption = ""
    lblSolSaldoCred.Caption = ""
    lblSolLineaDisp.Caption = ""
    lblDolSaldo.Caption = ""
    lblDolLimCobert.Caption = ""
    lblDolSaldoCred.Caption = ""
    lblDolLineaDisp.Caption = ""
    lblTipoCamb.Caption = ""
    TxtBCodPers.Enabled = True
    cmdBuscar.Enabled = False
End Sub

Private Sub cmdCancelar_Click()
    LimpiarControles
    LimpiarControlesSimulador
    TxtBCodPers.SetFocus
End Sub

Private Function CalculaSaldoCredMN(ByVal psCtaCod As String, ByVal pnTpoCamb As Double) As Double
    Dim oNCred As COMNCredito.NCOMCredito
    Dim oDCred As COMDCredito.DCOMCredito
    Dim oDCredAct As COMDCredito.DCOMCredActBD
    Dim oNCF As COMNCartaFianza.NCOMCartaFianzaValida
    Dim rsCred As ADODB.Recordset, RGas As ADODB.Recordset
    Dim MatCalend As Variant
    Dim nSaldoKFecha As Double, nIntCompFecha As Double, nGastoFecha As Double, nIntMorFecha As Double, nDeudaFechaMN As Double
    Dim nSaldoCredCob As Double
    
    CalculaSaldoCredMN = 0
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set rsCred = oDCred.RecuperaCreditosGarantiaDPF(psCtaCod)
    
    If rsCred.RecordCount > 0 Then
        Set oNCred = New COMNCredito.NCOMCredito
        
        Do While Not rsCred.EOF
            If Mid(rsCred!cCtaCod, 6, 3) = "121" Or Mid(rsCred!cCtaCod, 6, 3) = "221" Or Mid(rsCred!cCtaCod, 6, 3) = "514" Then
                Set oNCF = New COMNCartaFianza.NCOMCartaFianzaValida
                Set RGas = oNCF.RecuperaDatosT(rsCred!cCtaCod)
                Set oNCF = Nothing
                nDeudaFechaMN = RGas!nSaldo * IIf(Mid(rsCred!cCtaCod, 9, 1) = "1", 1, pnTpoCamb)
            Else
                MatCalend = oNCred.RecuperaMatrizCalendarioPendienteHistorial(rsCred!cCtaCod)
                nSaldoKFecha = Format(oNCred.MatrizCapitalAFecha(rsCred!cCtaCod, MatCalend), "#0.00")
                nIntCompFecha = Format(oNCred.MatrizInteresTotalesAFechaSinMora(rsCred!cCtaCod, MatCalend, gdFecSis) + oNCred.MatrizInteresGraAFecha(rsCred!cCtaCod, MatCalend, gdFecSis), "#0.00")
                Set oDCredAct = New COMDCredito.DCOMCredActBD
                Set RGas = oDCredAct.CargaRecordSet(" SELECT nGasto=dbo.ColocCred_ObtieneGastoFechaCredito('" & rsCred!cCtaCod & "','" & Format(gdFecSis, "mm/dd/yyyy") & "')")
                nGastoFecha = RGas!nGasto
                Set oDCredAct = Nothing
                nIntMorFecha = Format(oNCred.ObtenerMoraVencida(gdFecSis, MatCalend), "#0.00")
                
                nDeudaFechaMN = (nSaldoKFecha + nIntCompFecha + nGastoFecha + nIntMorFecha) * IIf(Mid(rsCred!cCtaCod, 9, 1) = "1", 1, pnTpoCamb)
            End If
            nSaldoCredCob = rsCred!nPorcMontoCobert * nDeudaFechaMN
            'CalculaSaldoCredMN = CalculaSaldoCredMN + nDeudaFechaMN
            CalculaSaldoCredMN = CalculaSaldoCredMN + nSaldoCredCob
            rsCred.MoveNext
        Loop
        Set oNCred = Nothing
    End If
End Function

Private Sub txtMonto_Change()
    If ChkPerGra.value = 1 Then
        bGraciaGenerada = True
    Else
        bGraciaGenerada = False
    End If
    Call LimpiaFlex(FECalend)
End Sub

Private Sub txtMonto_GotFocus()
    fEnfoque TxtMonto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtMonto, KeyAscii)
    If KeyAscii = 13 Then
        TxtInteres.SetFocus
    End If
End Sub

Private Sub txtMonto_LostFocus()
    If Trim(TxtMonto.Text) = "" Then
        TxtMonto.Text = "0.00"
    Else
        If lblSolLineaDisp.Caption = "" Then
            MsgBox "Es necesario buscar el cliente para verificar su linea disponible", vbInformation, "Aviso"
            cmdCancelar_Click
            Exit Sub
        End If
        
        If CDbl(TxtMonto.Text) > CDbl(lblSolLineaDisp.Caption) Then
            MsgBox "El monto excede de la linea disponible", vbInformation, "Aviso"
            TxtMonto.Text = "0.00"
            TxtMonto.SetFocus
        Else
            TxtMonto.Text = Format(TxtMonto.Text, "#0.00")
        End If
    End If
End Sub

Private Sub txtInteres_Change()
     If ChkPerGra.value = 1 Then
        bGraciaGenerada = True
    Else
        bGraciaGenerada = False
    End If
    Call LimpiaFlex(FECalend)
End Sub

Private Sub txtinteres_GotFocus()
    fEnfoque TxtInteres
End Sub

Private Sub txtInteres_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtInteres, KeyAscii, , 4)
    If KeyAscii = 13 Then
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

Private Sub SpnCuotas_Change()
     If ChkPerGra.value = 1 Then
        bGraciaGenerada = True
    Else
        bGraciaGenerada = False
    End If
    Call LimpiaFlex(FECalend)
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
     If ChkPerGra.value = 1 Then
        bGraciaGenerada = True
    Else
        bGraciaGenerada = False
    End If
    Call LimpiaFlex(FECalend)
    GenerarFechaPago
    ChkPerGra.value = 0
End Sub

Private Sub spnPlazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And OptTipoCuota(0).Enabled Then
        OptTipoCuota(0).SetFocus
    End If
End Sub

Private Sub OptTipoCuota_Click(Index As Integer)
    Call LimpiaFlex(FECalend)
End Sub

Private Sub OptTipoPeriodo_Click(Index As Integer)
    Call LimpiaFlex(FECalend)
    If Index = 1 Then
        Call HabilitaFechaFija(True)
        optTipoGracia(0).Enabled = False
        Frame6.Enabled = False
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

Private Sub DTFecDesemb_Change()
     If ChkPerGra.value = 1 Then
        bGraciaGenerada = True
    Else
        bGraciaGenerada = False
    End If
    Call LimpiaFlex(FECalend)
    GenerarFechaPago
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
                ChkProxMes.value = 0
            Else
                ChkProxMes.value = 1
            End If
        End If
    End If
End Sub
'***

Private Sub ChkProxMes_Click()
     If ChkPerGra.value = 1 Then
        bGraciaGenerada = True
    Else
        bGraciaGenerada = False
    End If
    Call LimpiaFlex(FECalend)
End Sub

Private Sub TxtDiaFijo2_Change()
If TxtDiaFijo2.Text = "" Then TxtDiaFijo2.Text = "00"

If CInt(TxtDiaFijo2.Text) = 0 Then
    ChkProxMes.Enabled = True
Else
    ChkProxMes.Enabled = False
End If

If CInt(TxtDiaFijo2.Text) > 31 Then
    TxtDiaFijo2.Text = "00"
End If
Call LimpiaFlex(FECalend)
End Sub

Private Sub TxtDiaFijo2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
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
                ChkProxMes.value = 0
            Else
                ChkProxMes.value = 1
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
    End If
End Sub

Private Sub cmdImprimir_Click()
    If Len(Trim(FECalend.TextMatrix(1, 1))) = 0 Then
        MsgBox "No existen datos para imprimir", vbExclamation, "Aviso"
        Exit Sub
    Else
        EjecutaReporte
    End If
End Sub

Private Sub HabilitaFechaFija(ByVal pbHabilita As Boolean)
    SpnPlazo.Enabled = Not pbHabilita
    SpnPlazo.valor = IIf(pbHabilita, "0", SpnPlazo.valor)
    LblDia.Enabled = pbHabilita
    TxtDiaFijo.Enabled = pbHabilita
    ChkProxMes.Enabled = pbHabilita
    ChkProxMes.value = 0
    TxtDiaFijo2.Enabled = pbHabilita
    lblDia2.Enabled = pbHabilita
    If pbHabilita Then
        TxtDiaFijo.Text = "00"
        If TxtDiaFijo.Enabled And TxtDiaFijo.Visible And FraDatos.Enabled Then
            TxtDiaFijo.SetFocus
        End If
    End If
End Sub

Private Sub EjecutaReporte()
Dim loRep As COMNCredito.NCOMCalendario
Dim lsDestino As String  ' P= Previo // I = Impresora // A = Archivo // E = Excel
Dim lsCadImp As String
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

If OptTipoCuota(0).value Then
    tIndex = 0
Else
    If OptTipoCuota(1).value Then
        tIndex = 1
    Else
        tIndex = 2
    End If
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
    loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
    lsCadImp = lsCadImp & Chr(10) & loRep.ReporteCalendario(1, MatCalend, MatResul, _
    TCuota, CDbl(TxtInteres.Text), TxtMonto.Text, SpnCuotas.valor, SpnPlazo.valor, DTFecDesemb.value, nSugerAprob, IIf(bDesemParcial, MatDesPar, ""), gbITFAplica, gnITFPorcent, gnITFMontoMin, cCtaCodG, pnCuotas, _
    nTasaEfectivaAnual, nTasaCostoEfectivoAnual, lsCtaCodLeasing, nTipoGracia, nIntGraInicial, CInt(TxtPerGra.Text), lblNomPers.Caption) 'DAOR 20070403

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
Dim nRedondeoITF As Double
Dim nTotalcuotasCONItF As Double
Dim nTotalcuotasLeasing As Double
Dim lnSalCapital As Double

nTotalcuotasLeasing = 0

    nIntGraInicial = 0
    nMontoCapInicial = 0
    
    Call LimpiaFlex(FECalend)
    MatResul = Array(0)
    MatResulDiff = Array(0)
    MatCalend = Array(0)
       
    If Not ValidaDatos Then
        bErrorValidacion = True
        Exit Sub
    Else
        bErrorValidacion = False
    End If

    Call LimpiaFlex(FECalend)

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
            
    If optTipoGracia(0).value Then
        Dim oCredito As COMNCredito.NCOMCredito
        Set oCredito = New COMNCredito.NCOMCredito
        
        nMontoCapInicial = CDbl(TxtMonto.Text)
        
        nIntGraInicial = oCredito.MontoIntPerDias(CDbl(TxtTasaGracia.Text), CInt(TxtPerGra.Text), CDbl(TxtMonto.Text))

        nTipoGracia = gColocTiposGraciaCapitalizada

        Set oCredito = Nothing
    End If
    
    If Len(Trim(lsCtaCodLeasing)) = 0 Then
    MatCalend = GeneraCalendario(CDbl(TxtMonto.Text), CDbl(TxtInteres.Text), CInt(SpnCuotas.valor), _
                CInt(SpnPlazo.valor), CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), nTipoCuota, nTipoPeriodo, _
                nTipoGracia, CInt(TxtPerGra.Text), CInt(TxtDiaFijo.Text), ChkProxMes.value, MatGracia, True, False, bDesemParcial, MatDesPar, , , , _
                0, CDbl(TxtTasaGracia.Text), CInt(TxtDiaFijo2.Text), nMontoCapInicial, False, bRenovarCredito, nInteresAFecha, nIntGraInicial)
    Else
    MatCalend = GeneraCalendarioLeasing(CDbl(TxtMonto.Text), CDbl(TxtInteres.Text), CInt(SpnCuotas.valor), _
                CInt(SpnPlazo.valor), CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), nTipoCuota, nTipoPeriodo, _
                nTipoGracia, CInt(TxtPerGra.Text), CInt(TxtDiaFijo.Text), ChkProxMes.value, MatGracia, True, False, bDesemParcial, MatDesPar, , , , _
                0, CDbl(TxtTasaGracia.Text), CInt(TxtDiaFijo2.Text), nMontoCapInicial, False, bRenovarCredito, nInteresAFecha, lsCtaCodLeasing)
    End If

    If cmdAplicar.Enabled Then
        Call ObtenerDesgravamen
    Else
        For i = 0 To UBound(MatCalend) - 1
            MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 8)), "#0.00")
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
        '***
    
        FECalend.TextMatrix(i + 1, 2) = Trim(MatCalend(i, 1))
        If nTipoGracia = 6 Then
            MatCalend(i, 2) = Trim(CDbl(MatCalend(i, 2)) + CDbl(MatCalend(i, 11)))
            FECalend.TextMatrix(i + 1, 3) = Format(Trim(MatCalend(i, 2)), "#0.00")
        Else
            FECalend.TextMatrix(i + 1, 3) = Trim(MatCalend(i, 2))
        End If
        '***
        
        FECalend.row = i + 1
        FECalend.Col = 3
        FECalend.CellForeColor = vbBlue
        FECalend.TextMatrix(i + 1, 4) = Trim(MatCalend(i, 3)) 'Amort Cap
        
        If nTipoGracia = 6 Then
            MatCalend(i, 4) = Trim(CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 11)))
            FECalend.TextMatrix(i + 1, 5) = Format(Trim(MatCalend(i, 4)), "#0.00")
        Else
            FECalend.TextMatrix(i + 1, 5) = Trim(MatCalend(i, 4))
        End If
        
        FECalend.TextMatrix(i + 1, 6) = Trim(MatCalend(i, 5))
    
        If Len(Trim(lsCtaCodLeasing)) = 0 Then
            FECalend.TextMatrix(i + 1, 7) = Trim(MatCalend(i, 8))
            FECalend.TextMatrix(i + 1, 8) = Trim(MatCalend(i, 6))
        Else
            FECalend.TextMatrix(i + 1, 7) = Trim(MatCalend(i, 6))
            FECalend.TextMatrix(i + 1, 8) = Trim(MatCalend(i, 8))
        End If
        
        FECalend.TextMatrix(i + 1, 9) = Trim(MatCalend(i, 7))

        If Not (i = 0 And nTipoGracia = gColocTiposGraciaCapitalizada) Then
            nTotalCapital = nTotalCapital + CDbl(Trim(MatCalend(i, 3)))
            nTotalInteres = nTotalInteres + CDbl(Trim(MatCalend(i, 4))) + CDbl(Trim(MatCalend(i, 5)))
        End If

        FECalend.TextMatrix(i + 1, 10) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))), "0.00")
        nRedondeoITF = fgDiferenciaRedondeoITF(fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))))
        If nRedondeoITF > 0 Then
            FECalend.TextMatrix(i + 1, 10) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))) - nRedondeoITF, "0.00")
        Else
            FECalend.TextMatrix(i + 1, 10) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))), "0.00")
        End If

        If Not (nTipoGracia = gColocTiposGraciaCapitalizada And i = 0) Then
            nTotalcuotasCONItF = nTotalcuotasCONItF + CDbl(FECalend.TextMatrix(i + 1, 10))
            nTotalcuotasLeasing = nTotalcuotasLeasing + CDbl(FECalend.TextMatrix(i + 1, 10))
        End If
    Next i
    
    Set oCredito = Nothing
    If nTipoGracia = gColocTiposGraciaCapitalizada Then
        FECalend.TextMatrix(1, 3) = ""
        FECalend.TextMatrix(1, 5) = ""
        FECalend.TextMatrix(1, 7) = ""
        FECalend.TextMatrix(1, 8) = ""
        FECalend.TextMatrix(1, 10) = ""
    End If
    
    lblCapital.Caption = Format(nTotalCapital, "#,##0.00")
    lblInteres.Caption = Format(nTotalInteres, "#,##0.00")
    lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#,##0.00")
    lblTotalCONITF.Caption = Format(nTotalcuotasCONItF, "#,##0.00")
    FECalend.row = 1
    FECalend.TopRow = 1
    
    'fraTasaAnuales.Visible = True
    Set oCredito = New COMNCredito.NCOMCredito
        nTasaEfectivaAnual = Round(oCredito.TasaIntPerDias(CDbl(TxtInteres), 360) * 100, 2)
        
        If nTipoGracia = 6 Then
            Dim y As Integer
            Dim MatCalendTemp() As String
            ReDim MatCalendTemp(UBound(MatCalend) - 1, 13)
            For i = 0 To UBound(MatCalend) - 1
                For y = 0 To 13
                    MatCalendTemp(i, y) = MatCalend(i + 1, y)
                Next y
            Next i
            Erase MatCalend
            ReDim MatCalend(UBound(MatCalendTemp), 13)
            
            For i = 0 To UBound(MatCalendTemp)
                For y = 0 To 13
                    MatCalend(i, y) = MatCalendTemp(i, y)
                Next y
            Next i
            Erase MatCalendTemp
        End If
        
        nTasaCostoEfectivoAnual = oCredito.GeneraTasaCostoEfectivoAnual(CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), CDbl(TxtMonto.Text), MatCalend, CDbl(TxtInteres), lsCtaCodLeasing)

        'lblTasaEfectivaAnual.Caption = nTasaEfectivaAnual & " %"
        'lblTasaCostoEfectivoAnual.Caption = nTasaCostoEfectivoAnual & " %"
        
    If UBound(MatCalend) = 0 Then
        cmdImprimir.Enabled = False
    Else
        cmdImprimir.Enabled = True
    End If
    
End Sub

Private Function ValidaDatos() As Boolean
    ValidaDatos = True
    
    'Monto de Prestamo
    If Trim(TxtMonto.Text) = "" Then
        MsgBox "Ingrese el Monto del Prestamo", vbInformation, "Aviso"
        ValidaDatos = False
        If TxtMonto.Enabled Then TxtMonto.SetFocus
        Exit Function
    End If
    
    'Interes
    If Trim(TxtInteres.Text) = "" Then
        MsgBox "Ingrese la Tasa de Interes", vbInformation, "Aviso"
        ValidaDatos = False
        If TxtInteres.Enabled Then TxtInteres.SetFocus
        Exit Function
    Else
        If CDbl(TxtInteres.Text) = 0 Then
            MsgBox "Ingrese la Tasa de Interes", vbInformation, "Aviso"
            ValidaDatos = False
            If TxtInteres.Enabled Then TxtInteres.SetFocus
            Exit Function
        End If
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
    
    'Fecha de Desembolso
    If ValidaFecha(DTFecDesemb.value) <> "" Then
        MsgBox ValidaFecha(DTFecDesemb.value), vbInformation, "Aviso"
        ValidaDatos = False
        If DTFecDesemb.Enabled Then DTFecDesemb.SetFocus
        Exit Function
    End If
    
    'Valida dia de Fecha Fija
    If OptTipoPeriodo(1).value And (Trim(TxtDiaFijo.Text) = "" Or Trim(TxtDiaFijo.Text) = "0" Or Trim(TxtDiaFijo.Text) = "00") Then
        MsgBox "Ingrese el Dia del Mes que Venceran todas las cuotas", vbInformation, "Aviso"
        ValidaDatos = False
        If TxtDiaFijo.Enabled Then TxtDiaFijo.SetFocus
        Exit Function
    End If
    If CInt(TxtDiaFijo2.Text) > 0 And (CInt(TxtDiaFijo2.Text) <= CInt(TxtDiaFijo.Text)) Then
        MsgBox "El 2do dia tiene que ser mayor al 1ro", vbInformation, "Mensaje"
        ValidaDatos = False
        TxtDiaFijo2.Text = "00"
        TxtDiaFijo2.SetFocus
        Exit Function
    End If
    
    'Valida Generacion de Tipos de Periodo de Gracia
    If ChkPerGra.value = 1 Then
        If (TxtPerGra.Text = "00" Or TxtPerGra.Text = "0") Then
            MsgBox "Ingrese los Dias de Gracia", vbInformation, "Aviso"
            ValidaDatos = False
            If TxtPerGra.Enabled Then TxtPerGra.SetFocus
            Exit Function
        Else
            If (TxtTasaGracia.Text = "0.00" Or TxtTasaGracia.Text = "") Then
                MsgBox "Ingrese la Tasa de Gracia ", vbInformation, "Aviso"
                ValidaDatos = False
                If TxtTasaGracia.Enabled Then TxtTasaGracia.SetFocus
                    Exit Function
                Else
                    If Not bGraciaGenerada And (optTipoGracia(0).value = False) Then
                        ValidaDatos = False
                        MsgBox "Seleccione un Tipo de Gracia", vbInformation, "Aviso"
                    If CmdGracia.Enabled Then
                        CmdGracia.SetFocus
                    End If
                    Exit Function
                End If
            End If
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
    
End Function

Private Sub ObtenerDesgravamen()
Dim oNGasto As COMNCredito.NCOMGasto
Dim oDCredito As COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Dim nNumGastos As Integer
Dim nTotalGasto As Double
Dim nTotalGastoSeg As Double
Dim i, J As Integer
        
        ReDim MatDesemb(1, 2)
        MatDesemb(0, 0) = Format(DTFecDesemb.value, "dd/mm/yyyy")
        MatDesemb(0, 1) = Format(TxtMonto.Text, "#0.00")
    
        Set oNGasto = New COMNCredito.NCOMGasto
            MatGastos = oNGasto.GeneraCalendarioGastos_NEW(CDbl(TxtMonto.Text), CDbl(TxtInteres.Text), CInt(SpnCuotas.valor), _
                                CInt(SpnPlazo.valor), CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), getTipoCuota, _
                                IIf(OptTipoPeriodo(0).value, 1, 2), nTipoGracia, CInt(TxtPerGra.Text), _
                                CDbl(TxtTasaGracia.Text), CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)), _
                                ChkProxMes.value, MatGracia, 0, 0, MatCalend_2, _
                                MatDesemb, nNumGastos, gdFecSis, _
                                sCtaCodRep, 1, "SI", "F", _
                                CDbl(MatCalend(0, 2)), CDbl(TxtMonto.Text), , , , , , , , True, _
                                2, True, False, MatDesemb, False, , _
                                CInt(TxtDiaFijo2.Text), , , _
                                gnITFMontoMin, gnITFPorcent, gbITFAplica, 0, , , , , , nIntGraInicial, , , 0)

        Set oNGasto = Nothing
        Set rsCredito = Nothing
        Call frmCredReprogCred.EstablecerGastos(MatGastos, True, nNumGastos, IIf(OptTipoPeriodo(0).value, 1, 2), CInt(SpnPlazo.valor))
        '***************************************************************************
        'Adicionamos los Gastos
        '***************************************************************************
               
        If IsArray(MatGastos) Then
            For J = 0 To UBound(MatCalend) - 1
                nTotalGasto = 0
                nTotalGastoSeg = 0
                For i = 0 To UBound(MatGastos) - 1
                'Comentado por MAVM para separar los gastos 20100320
'                    If Trim(Right(MatGastos(I, 0), 2)) = "1" And _
'                       (Trim(MatGastos(I, 1)) = Trim(MatCalend(J, 1)) _
'                         Or Trim(MatGastos(I, 1)) = "*") Then
'                        nTotalGasto = nTotalGasto + CDbl(MatGastos(I, 3))
'                    End If
                    
                    If Trim(Right(MatGastos(i, 0), 2)) = "1" And _
                       (Trim(MatGastos(i, 1)) = Trim(MatCalend(J, 1))) Then
                        nTotalGasto = CDbl(MatGastos(i, 3))
                        MatCalend(J, 6) = Format(nTotalGasto, "#0.00")
                    Else
                        If Trim(MatGastos(i, 1)) = "*" Then
                            nTotalGasto = CDbl(MatGastos(i, 3))
                            MatCalend(J, 8) = Format(nTotalGasto, "#0.00")
                        End If

                    End If

'                    If Trim(Right(MatGastos(i, 0), 2)) = "1" And _
'                       (Trim(MatGastos(i, 1)) = Trim(MatCalend(J, 1)) And Trim(Right(MatGastos(i, 2), 4)) = "1217") Then
'                        nTotalGastoSeg = nTotalGastoSeg + CDbl(MatGastos(i, 3))
'                        MatCalend(J, 6) = Format(nTotalGastoSeg, "#0.00")
'                    Else
'                        If Trim(MatGastos(i, 1)) = "*" Or (Trim(Right(MatGastos(i, 0), 2)) = "1" And Trim(MatGastos(i, 1)) = Trim(MatCalend(J, 1)) And Trim(Right(MatGastos(i, 2), 4)) <> "1217") Then
'                            nTotalGasto = nTotalGasto + CDbl(MatGastos(i, 3))
'                            MatCalend(J, 8) = Format(nTotalGasto, "#0.00")
'                        End If
'
'                    End If
                    
                Next i
                'Add By GITU 06-08-2008
                'Descomentar cuando esten seguros de los cambios GITU
'                If j > 0 Then
'                    If InStr(Trim(Str(nTotalGasto)), ".") > 0 Then
'                        MatCalend(0, 6) = Format(MatCalend(0, 6) + (nTotalGasto - Val(Left(Trim(Str(nTotalGasto)), InStr(Trim(nTotalGasto), ".") - 1))), "#0.00")
'                        MatCalend(j, 6) = Format(Val(Left(Trim(nTotalGasto), InStr(Trim(nTotalGasto), ".") - 1)), "#0.00")
'                    Else
'                        MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
'                    End If
'                Else
'                    MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
'                End If
                'End GITU
                
                'Comentado por MAVM 20100320
                'MatCalend(J, 6) = Format(nTotalGasto, "#0.00")
                
                'MatCalend(j, 2) = Format(CDbl(MatCalend(j, 2)) + CDbl(MatCalend(j, 6)), "#0.00")
            Next J
        End If
               
        For i = 0 To UBound(MatCalend) - 1
            'MAVM 20100320
            'MatCalend(I, 2) = Format(CDbl(MatCalend(I, 3)) + CDbl(MatCalend(I, 4)) + CDbl(MatCalend(I, 5)) + CDbl(MatCalend(I, 6)), "#0.00")
            MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 8)), "#0.00")
        Next i

End Sub

Private Function getTipoCuota() As Integer
Dim i As Integer
    For i = 0 To 2
        If OptTipoCuota(i).value Then
            getTipoCuota = i + 1
            Exit For
        End If
    Next i
End Function

Private Sub CargaControles()
    FECalend.RowHeight(0) = 250
    FECalend.RowHeight(1) = 250
    TxtDiaFijo.Text = "00"
    TxtTasaGracia.Text = "00"
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
        
        TxtTasaGracia.Text = "0.00"
        CmdGracia.Enabled = True
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
        CmdGracia.Enabled = False
        optTipoGracia(0).Enabled = False
        optTipoGracia(0).value = False
        
        GenerarFechaPago
        If OptTipoPeriodo(1).value = True Then
            ChkPerGra.Enabled = False
        End If
    End If
End Sub

Private Sub CmdGracia_Click()
Dim oCredito As COMNCredito.NCOMCredito

Set oCredito = New COMNCredito.NCOMCredito

If CDbl(TxtTasaGracia.Text) <= 0# Then
    MsgBox "Ingrese la Tasa de Interes de Gracia", vbInformation, "Aviso"
    TxtTasaGracia.SetFocus
    Exit Sub
End If

    MatGracia = frmCredGracia.Inicio(CInt(TxtPerGra.Text), oCredito.MontoIntPerDias(CDbl(TxtTasaGracia.Text), CInt(TxtPerGra.Text), CDbl(TxtMonto.Text)), CInt(SpnCuotas.valor), nTipoGracia, psCtaCod)
    
    Set oCredito = Nothing
    bGraciaGenerada = True
    Call LimpiaFlex(FECalend)
End Sub

Private Sub LimpiarControlesSimulador()
    TxtMonto.Text = 0
    TxtInteres.Text = 0
    SpnCuotas.valor = 0
    SpnPlazo.valor = 0
    DTFecDesemb.value = gdFecSis
    OptTipoCuota(0).value = 1
    OptTipoPeriodo(0).value = 1
    TxtDiaFijo.Text = 0
    TxtDiaFijo2.Text = 0
    ChkPerGra.value = 0
    TxtPerGra.Text = 0
    TxtTasaGracia.Text = 0
    Call LimpiaFlex(FECalend)
End Sub
